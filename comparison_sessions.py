"""Durable comparison snapshots; closing a view never deletes its saved results."""
import json
import os
import re
import sqlite3
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from comparison import Comparison

FIELDS = ('mode','query','filters','names','quantity_mode','display_columns','summary',
          'columns','shared','qty_pairs','has_qty','field_pairs','key_columns')


def path_for(root, token):
    if not isinstance(token,str) or not re.fullmatch(r'[0-9a-f]{32}',token):
        raise ValueError('معرّف جلسة غير صالح.')
    return Path(root)/f'{token}.sqlite'


def update_metadata(conn, **changes):
    raw=conn.execute('SELECT data FROM session_metadata WHERE id=1').fetchone()
    state=json.loads(raw[0]);state['session'].update(changes)
    conn.execute('UPDATE session_metadata SET data=? WHERE id=1',(json.dumps(state,ensure_ascii=False),))


def save(engine,root,token,title):
    root=Path(root);root.mkdir(parents=True,exist_ok=True)
    path=path_for(root,token)
    now=datetime.now(timezone.utc).isoformat()
    engine.summary['snapshot_created_at']=now
    state={'version':1,'engine':{name:getattr(engine,name) for name in FIELDS},
           'session':{'id':token,'title':title,'created_at':now,'updated_at':now,'view':{}}}
    fd,stage=tempfile.mkstemp(prefix='.session-',suffix='.sqlite',dir=root);os.close(fd)
    source=sqlite3.connect(engine.path);target=sqlite3.connect(stage)
    try:
        source.backup(target)
        target.execute('CREATE TABLE session_metadata(id INTEGER PRIMARY KEY CHECK(id=1),data TEXT NOT NULL)')
        target.execute('INSERT INTO session_metadata VALUES(1,?)',(json.dumps(state,ensure_ascii=False),));target.commit()
        source.close();target.close()
        os.replace(stage,path)
    finally:
        source.close();target.close()
        if os.path.exists(stage):os.remove(stage)
    engine.close();engine.temp=None;engine.path=str(path)


def read(root,token):
    path=path_for(root,token)
    if not path.is_file():raise ValueError('المقارنة المحفوظة غير موجودة.')
    conn=sqlite3.connect(path.resolve().as_uri()+'?mode=ro',uri=True)
    try:
        row=conn.execute('SELECT data FROM session_metadata WHERE id=1').fetchone()
        state=json.loads(row[0])
        if state.get('version')!=1:raise ValueError('صيغة جلسة المقارنة غير مدعومة.')
        return state
    finally:conn.close()


def restore(root,token):
    state=read(root,token);engine=Comparison.__new__(Comparison)
    for name in FIELDS:setattr(engine,name,state['engine'][name])
    engine.qty_pairs=[tuple(pair) for pair in engine.qty_pairs]
    engine.field_pairs=[tuple(pair) for pair in engine.field_pairs]
    engine.temp=None;engine.path=str(path_for(root,token));engine.report=lambda _:None
    return engine,state['session']


def remember(root,token,**changes):
    path=path_for(root,token)
    if not path.is_file():raise ValueError('المقارنة المحفوظة غير موجودة.')
    conn=sqlite3.connect(path)
    try:
        update_metadata(conn,**changes);conn.commit()
    finally:conn.close()


def listing(root):
    sessions=[]
    for path in Path(root).glob('*.sqlite'):
        if not re.fullmatch(r'[0-9a-f]{32}',path.stem):continue
        try:
            state=read(root,path.stem)
            conn=sqlite3.connect(path.resolve().as_uri()+'?mode=ro',uri=True)
            try:counts=dict(conn.execute('SELECT section,COUNT(*) FROM results GROUP BY section'))
            finally:conn.close()
            sessions.append({**state['session'],'names':state['engine']['names'],'mode':state['engine']['mode'],'counts':counts})
        except (ValueError,sqlite3.Error,KeyError,TypeError):
            sessions.append({'id':path.stem,'title':path.stem,'error':'تعذر قراءة الجلسة؛ لم يتم حذفها.'})
    return sorted(sessions,key=lambda s:s.get('updated_at',''),reverse=True)
