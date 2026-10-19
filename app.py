#!/usr/bin/env python3
"""
app.py — واجهة الويب المحلية لأداة البحث في ملفات الإكسل
=============================================================
تشغيل: python3 app.py
تفتح تلقائياً على المتصفح على http://127.0.0.1:5000

بيدعم أكتر من ملف إكسل/قاعدة في نفس الوقت — تقدر تضيف ملف جديد من غير ما يمسح
الملفات اللي فهرستها قبل كده، وتبدّل بينهم من قائمة أعلى الصفحة.

الإعدادات (قائمة الملفات، مسار كل قاعدة، آخر ملف مفتوح) بتتحفظ في config.json
بجانب هذا الملف (أو بجانب ملف الـ exe لو اتحول).
"""

import io
import json
import os
import sqlite3
import sys
import threading
import time
import uuid
import webbrowser
import tempfile
import atexit
from concurrent.futures import ThreadPoolExecutor
from functools import wraps

from flask import Flask, jsonify, request, send_file, render_template

import excel_db_core as core
from comparison import Comparison
import comparison_sessions as saved_comparisons

# ----------------------------------------------------------------------
# تحديد مسار العمل (يشتغل صح سواء .py عادي أو .exe مبني بـ PyInstaller)
# ----------------------------------------------------------------------
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(APP_DIR, "config.json")
DATABASES_DIR = os.path.join(APP_DIR, "databases")
os.makedirs(DATABASES_DIR, exist_ok=True)

RESOURCE_DIR = getattr(sys, '_MEIPASS', APP_DIR)
app = Flask(__name__, template_folder=os.path.join(RESOURCE_DIR,'templates'), static_folder=os.path.join(RESOURCE_DIR,'static'))
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
app.json.sort_keys = False
APP_VERSION = '2026.10.19-saved-review-v8'
APP_PORT = int(os.environ.get('EXCEL_SEARCH_PORT', '5000'))
operation_lock = threading.RLock()
jobs_lock = threading.RLock()
executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="excel-work")
jobs = {}
comparisons = {}


def start_job(work):
    with jobs_lock:
        # Bound retained terminal jobs, but never discard active jobs.
        for key in list(jobs):
            if len(jobs) < 32:
                break
            if jobs[key]['state'] in ('done', 'error', 'cancelled'):
                del jobs[key]
        if len(jobs) >= 32:
            raise ValueError('توجد مهام كثيرة. انتظر اكتمال بعضها.')
        token = uuid.uuid4().hex
        job = {'state': 'queued', 'message': 'في انتظار التنفيذ', 'cancel': threading.Event()}
        jobs[token] = job

    def progress(message):
        if job['cancel'].is_set():
            raise InterruptedError('تم إلغاء العملية.')
        with jobs_lock:
            job['message'] = message

    def run():
        try:
            with operation_lock:
                progress('بدء التنفيذ...')
                job['state'] = 'running'
                result = work(progress)
                with jobs_lock:
                    job.update(state='done', result=result, message='اكتملت العملية')
        except InterruptedError as exc:
            with jobs_lock:
                job.update(state='cancelled', error=str(exc))
        except Exception as exc:
            app.logger.exception('Background operation failed')
            with jobs_lock:
                job.update(state='error', error=str(exc))
    executor.submit(run)
    return jsonify({'job_id': token}), 202


@app.route('/api/jobs/<token>')
def job_status(token):
    with jobs_lock:
        job = jobs.get(token)
        if not job:
            return jsonify(error='المهمة غير موجودة. أعد المحاولة.'), 404
        return jsonify({k: v for k, v in job.items() if k != 'cancel'})


@app.route('/api/jobs/<token>/cancel', methods=['POST'])
def cancel_job(token):
    with jobs_lock:
        job = jobs.get(token)
        if not job:
            return jsonify(error='المهمة غير موجودة.'), 404
        if job['state'] in ('queued', 'running'):
            job['cancel'].set()
    return jsonify(ok=True)


def serialized(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        with operation_lock:
            return fn(*args, **kwargs)
    return wrapped


def cleanup_comparisons():
    for item in comparisons.values():
        item['engine'].close()


atexit.register(cleanup_comparisons)


# ----------------------------------------------------------------------
# إدارة إعدادات التطبيق: قائمة الملفات المفهرسة + الملف النشط حالياً
# ----------------------------------------------------------------------
def load_app_config():
    if not os.path.exists(CONFIG_PATH):
        return {"active_db_id": None, "databases": {}}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # توافق مع النسخة القديمة (ملف واحد فقط) — بترحّله تلقائياً للصيغة الجديدة
    if "databases" not in cfg and cfg.get("db_path"):
        old_id = "db_1"
        cfg = {
            "active_db_id": old_id,
            "databases": {
                old_id: {
                    "name": os.path.basename(cfg.get("excel_path", "ملف قديم")),
                    "db_path": cfg["db_path"],
                    "excel_path": cfg.get("excel_path"),
                    "unique_col": cfg.get("unique_col"),
                }
            },
        }
        save_app_config(cfg)
    cfg.setdefault("active_db_id", None)
    cfg.setdefault("databases", {})
    return cfg


def save_app_config(cfg):
    fd, path = tempfile.mkstemp(prefix='.config-', suffix='.json', dir=os.path.dirname(CONFIG_PATH))
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(path, CONFIG_PATH)
    finally:
        if os.path.exists(path):
            os.remove(path)


def get_active_entry(cfg):
    data = request.get_json(silent=True) or {}
    explicit = request.args.get('db_id') or data.get('db_id')
    if request.method == 'POST' and not explicit:
        return None, None
    active_id = explicit or cfg.get("active_db_id")
    if not active_id:
        return None, None
    entry = cfg.get("databases", {}).get(active_id)
    return active_id, entry


# ----------------------------------------------------------------------
# API — إدارة قائمة الملفات (databases)
# ----------------------------------------------------------------------
@app.route("/api/databases")
def api_list_databases():
    cfg = load_app_config()
    out = []
    for db_id, entry in cfg.get("databases", {}).items():
        row_count = None
        columns = []
        if os.path.exists(entry["db_path"]):
            try:
                stats = core.get_stats(entry["db_path"])
                row_count = stats.get("row_count")
                columns = stats.get("search_columns") or []
            except Exception:
                row_count = None
        out.append({
            "id": db_id,
            "name": entry.get("name"),
            "row_count": row_count,
            "columns": columns,
            "active": db_id == cfg.get("active_db_id"),
        })
    return jsonify({"databases": out})


@app.route("/api/switch", methods=["POST"])
@serialized
def api_switch():
    data = request.get_json(force=True)
    db_id = data.get("db_id")
    cfg = load_app_config()
    if db_id not in cfg.get("databases", {}):
        return jsonify({"error": "الملف ده مش موجود."}), 400
    cfg["active_db_id"] = db_id
    save_app_config(cfg)
    return jsonify({"ok": True})


@app.route("/api/delete_db", methods=["POST"])
@serialized
def api_delete_db():
    data = request.get_json(force=True)
    db_id = data.get("db_id")
    cfg = load_app_config()
    entry = cfg.get("databases", {}).pop(db_id, None)
    if not entry:
        return jsonify({"error": "الملف ده مش موجود."}), 400
    if os.path.exists(entry["db_path"]):
        os.remove(entry["db_path"])
    if cfg.get("active_db_id") == db_id:
        remaining = list(cfg["databases"].keys())
        cfg["active_db_id"] = remaining[0] if remaining else None
    save_app_config(cfg)
    return jsonify({"ok": True})


# ----------------------------------------------------------------------
# API — الإعداد والفهرسة
# ----------------------------------------------------------------------
@app.route("/api/stats")
def api_stats():
    cfg = load_app_config()
    db_id, entry = get_active_entry(cfg)
    if not entry or not os.path.exists(entry["db_path"]):
        return jsonify({"configured": False, "has_any_db": bool(cfg.get("databases"))})
    stats = core.get_stats(entry["db_path"])
    stats["configured"] = True
    stats["db_id"] = db_id
    stats["display_name"] = entry.get("name")
    conn = sqlite3.connect(entry["db_path"])
    stats["local_columns"] = core.get_config(conn, "local_columns", []) or []
    conn.close()
    return jsonify(stats)


@app.route("/api/setup", methods=["POST"])
@serialized
def api_setup():
    data = request.get_json(force=True)
    excel_path = (data.get("excel_path") or "").strip().strip('"')
    unique_col = (data.get("unique_col") or "").strip() or None
    name = (data.get("name") or "").strip()
    header_row_raw = str(data.get("header_row") or "").strip()
    header_row = int(header_row_raw) if header_row_raw.isdigit() else None

    if not excel_path:
        return jsonify({"error": "من فضلك اكتب مسار ملف الإكسل."}), 400
    if not os.path.exists(excel_path):
        return jsonify({"error": f"مفيش ملف موجود في المسار ده: {excel_path}"}), 400
    if not name:
        name = os.path.splitext(os.path.basename(excel_path))[0]

    db_id = "db_" + uuid.uuid4().hex[:8]
    db_path = os.path.join(DATABASES_DIR, f"{db_id}.sqlite")

    def work(progress):
        core.build_db(excel_path, db_path, unique_col=unique_col, header_row=header_row,
                      last_column=data.get("last_column"), progress=progress)
        cfg = load_app_config()
        cfg["databases"][db_id] = {"name": name, "db_path": db_path, "excel_path": excel_path, "unique_col": unique_col}
        cfg["active_db_id"] = db_id
        save_app_config(cfg)
        return {"ok": True, "db_id": db_id}
    if data.get('async'):
        return start_job(work)
    try:
        return jsonify(work(lambda _: None))
    except Exception as exc:
        return jsonify(error=str(exc)), 400



@app.route("/api/sync", methods=["POST"])
@serialized
def api_sync():
    cfg = load_app_config()
    _, entry = get_active_entry(cfg)
    if not entry:
        return jsonify({"error": "لسه ماتمش إعداد أي قاعدة."}), 400
    force = bool((request.get_json(silent=True) or {}).get("force"))
    def work(progress):
        if not os.path.exists(entry['db_path']):
            raise ValueError('الملف لم يعد موجودًا في الأرشيف.')
        return core.sync_db(entry['db_path'], force=force, progress=progress)
    if (request.get_json(silent=True) or {}).get('async'):
        return start_job(work)
    try:
        return jsonify(work(lambda _: None))
    except Exception as exc:
        return jsonify(error=str(exc)), 400



# ----------------------------------------------------------------------
# API — البحث والتصدير (على الملف النشط حالياً)
# ----------------------------------------------------------------------
def _strip_internal(rows):
    out = []
    for r in rows:
        d = dict(r)
        d.pop("_hash", None)
        d.pop("_key", None)
        overrides_raw = d.pop("_overrides", None)
        try:
            d["_protected"] = json.loads(overrides_raw) if overrides_raw else []
        except (TypeError, json.JSONDecodeError):
            d["_protected"] = []
        out.append(d)
    return out


@app.route("/api/search")
def api_search():
    cfg = load_app_config()
    _, entry = get_active_entry(cfg)
    if not entry:
        return jsonify({"error": "لسه ماتمش إعداد أي قاعدة."}), 400
    db_path = entry["db_path"]
    q = request.args.get("q", "").strip()
    try:
        limit = max(1, min(int(request.args.get("limit", 20)), 200))
        offset = max(0, int(request.args.get("offset", 0)))
        filters = json.loads(request.args.get("filters", "[]"))
        if request.args.get("col") and request.args.get("val") and not filters:
            filters = [{"col": request.args["col"], "op": "equals", "value": request.args["val"]}]
        rows, total = core.query_rows(db_path, q, filters, limit, offset)
    except (ValueError, TypeError) as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"rows": rows, "total": total})



@app.route("/api/export")
@serialized
def api_export():
    from exports import write_search_export
    cfg=load_app_config()
    if request.args.get('all_files')=='1':
        entries=list(cfg.get('databases',{}).values())
    else:
        _,entry=get_active_entry(cfg)
        if not entry:
            return jsonify(error='الملف غير موجود.'),400
        entries=[entry]
    output=tempfile.TemporaryFile(suffix='.xlsx')
    try:
        filters=json.loads(request.args.get('filters','[]'))
        write_search_export(entries,request.args.get('q','').strip(),filters,output)
        output.seek(0)
        response=send_file(output,as_attachment=True,download_name='نتائج_البحث.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response.call_on_close(output.close)
        return response
    except (ValueError,TypeError) as exc:
        output.close()
        return jsonify(error=str(exc)),400
    except Exception:
        output.close()
        raise


# ----------------------------------------------------------------------
# API — بحث موحّد في كل الملفات + مقارنة ملفين
# ----------------------------------------------------------------------
@app.route("/api/search_all")
def api_search_all():
    cfg = load_app_config()
    q = request.args.get("q", "").strip()
    try:
        filters=json.loads(request.args.get('filters','[]'))
        if not isinstance(filters,list) or len(filters)>12 or any(not isinstance(f,dict) for f in filters):
            raise ValueError('قائمة فلاتر غير صالحة.')
        if not q and not filters:
            return jsonify(groups=[],skipped=[])
        groups,skipped=[],[]
        for db_id,entry in cfg.get('databases',{}).items():
            from row_filters import missing_column
            if any(missing_column(f, core.get_columns(entry['db_path'])) for f in filters):
                skipped.append(entry.get('name'))
                continue
            rows,total=core.query_rows(entry['db_path'],q,filters,limit=10)
            if total:
                groups.append({'db_id':db_id,'name':entry.get('name'),'rows':rows,'total':total})
        return jsonify(groups=groups,skipped=skipped)
    except (ValueError,TypeError) as exc:
        return jsonify(error=str(exc)),400



@app.route("/api/columns")
def api_columns():
    db_id = request.args.get("db_id")
    cfg = load_app_config()
    entry = cfg.get("databases", {}).get(db_id)
    if not entry or not os.path.exists(entry["db_path"]):
        return jsonify({"error": "الملف مش موجود."}), 400
    return jsonify({"columns": core.get_columns(entry["db_path"])})


@app.route("/api/compare", methods=["POST"])
def api_compare():
    data = request.get_json(force=True)
    cfg = load_app_config()
    entry_a = cfg.get("databases", {}).get(data.get("db_id_a"))
    entry_b = cfg.get("databases", {}).get(data.get("db_id_b"))
    if not entry_a or not entry_b or data.get("db_id_a") == data.get("db_id_b"):
        return jsonify(error="اختار ملفين مختلفين."), 400
    def work(progress):
        engine = Comparison(entry_a['db_path'], entry_b['db_path'], data.get('key_col_a'),
                            data.get('key_col_b'), query=str(data.get('query') or ''),
                            mode=data.get('mode', 'key'), threshold=float(data.get('threshold',65)),
                            progress=progress, filters_a=data.get('filters_a'), filters_b=data.get('filters_b'),
                            field_a=data.get('field_a') or None, field_b=data.get('field_b') or None,
                            name_a=entry_a.get('name'), name_b=entry_b.get('name'),
                            quantity_mode=data.get('quantity_mode') or None,
                            quantity_pairs=data.get('quantity_pairs'),
                            display_columns_a=data.get('display_columns_a'),
                            display_columns_b=data.get('display_columns_b'),
                            key_columns_a=data.get('key_columns_a'),key_columns_b=data.get('key_columns_b'))
        token = uuid.uuid4().hex
        title=str(data.get('title') or f"{entry_a.get('name')} ⇔ {entry_b.get('name')}").strip()[:160]
        try:
            progress('حفظ لقطة المقارنة على القرص...')
            saved_comparisons.save(engine,comparisons_root(),token,title)
        except BaseException:
            engine.close();raise
        trim_comparison_cache()
        comparisons[token] = {'engine': engine, 'name_a': entry_a.get('name'), 'name_b': entry_b.get('name')}
        return comparison_page(token)
    if data.get('async'):
        return start_job(work)
    try:
        with operation_lock:
            return jsonify(work(lambda _: None))
    except (ValueError, TypeError) as exc:
        return jsonify(error=str(exc)), 400


def comparisons_root():
    return os.path.join(os.path.dirname(CONFIG_PATH),'comparisons')


def trim_comparison_cache():
    while len(comparisons)>=12:
        comparisons.pop(next(iter(comparisons)))['engine'].close()


def get_comparison(token):
    if token not in comparisons:
        try:
            engine,_=saved_comparisons.restore(comparisons_root(),token)
        except (ValueError,sqlite3.Error,KeyError,TypeError):
            return None
        trim_comparison_cache()
        comparisons[token]={'engine':engine,'name_a':engine.names['a'],'name_b':engine.names['b']}
    return comparisons[token]


def comparison_page(token, section=None, offset=None, qty_only=None):
    item = get_comparison(token)
    if not item:
        raise ValueError('المقارنة المحفوظة غير موجودة أو تعذر فتحها.')
    session=saved_comparisons.read(comparisons_root(),token)['session']
    view=session.get('view',{})
    result=item['engine'].page(section or view.get('section'),offset if offset is not None else view.get('offset',0),
                              qty_only=qty_only if qty_only is not None else view.get('qty_only',False))
    if result['offset'] and result['offset']>=result['total']:
        result=item['engine'].page(result['section'],0,qty_only=result['qty_only'])
    saved_comparisons.remember(comparisons_root(),token,view={k:result[k] for k in ('section','offset','qty_only')})
    return {**result,'comparison_id':token,'session':session,'name_a':item['name_a'],'name_b':item['name_b']}


@app.route('/api/comparisons')
@serialized
def list_comparisons():
    return jsonify(sessions=saved_comparisons.listing(comparisons_root()))


@app.route('/api/compare_rename',methods=['POST'])
@serialized
def rename_comparison():
    data=request.get_json(force=True);title=str(data.get('title') or '').strip()
    try:
        if not title or len(title)>160:raise ValueError('اكتب اسمًا من 1 إلى 160 حرفًا.')
        saved_comparisons.remember(comparisons_root(),data.get('comparison_id'),title=title)
        return jsonify(ok=True)
    except (ValueError,sqlite3.Error):return jsonify(error='تعذر تغيير اسم المقارنة.'),400


@app.route('/api/compare_delete',methods=['POST'])
@serialized
def delete_comparison():
    token=(request.get_json(force=True) or {}).get('comparison_id')
    try:
        path=saved_comparisons.path_for(comparisons_root(),token)
        item=comparisons.pop(token,None)
        if item:item['engine'].close()
        path.unlink(missing_ok=True)
        return jsonify(ok=True)
    except (ValueError,OSError):return jsonify(error='تعذر حذف المقارنة المحفوظة.'),400


@app.route('/api/compare_page', methods=['POST'])
@serialized
def api_compare_page():
    data = request.get_json(force=True)
    try:
        return jsonify(comparison_page(data.get('comparison_id'), data.get('section'),
                                       int(data['offset']) if data.get('offset') is not None else None,
                                       bool(data['qty_only']) if 'qty_only' in data else None))
    except (ValueError, TypeError) as exc:
        return jsonify(error=str(exc)), 400


@app.route('/api/compare_decision', methods=['POST'])
@serialized
def api_compare_decision():
    data = request.get_json(force=True)
    item = get_comparison(data.get('comparison_id'))
    if not item:
        return jsonify(error='جلسة المقارنة انتهت.'), 400
    try:
        item['engine'].decide(int(data.get('id')), data.get('decision'),data.get('note'))
        return jsonify(ok=True)
    except (ValueError, TypeError) as exc:
        return jsonify(error=str(exc)), 400


@app.route('/api/compare_close', methods=['POST'])
@serialized
def api_compare_close():
    item = comparisons.pop((request.get_json(force=True) or {}).get('comparison_id'), None)
    if item:
        item['engine'].close()
    return jsonify(ok=True)


@app.route('/api/compare_export', methods=['POST'])
@serialized
def api_compare_export():
    item = get_comparison((request.get_json(force=True) or {}).get('comparison_id'))
    if not item:
        return jsonify(error='جلسة المقارنة انتهت. أعد المقارنة أولًا.'), 400
    output = tempfile.TemporaryFile(suffix='.xlsx')
    try:
        item['engine'].export(output)
        output.seek(0)
        response = send_file(output, as_attachment=True, download_name='تقرير_المقارنة.xlsx',
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response.call_on_close(output.close)
        return response
    except Exception:
        output.close()
        raise


# ----------------------------------------------------------------------
# API — تعديل الخلايا وإضافة أعمدة محلية
# ----------------------------------------------------------------------
@app.route("/api/update_cell", methods=["POST"])
@serialized
def api_update_cell():
    data = request.get_json(force=True)
    cfg = load_app_config()
    _, entry = get_active_entry(cfg)
    if not entry:
        return jsonify({"error": "لسه ماتمش إعداد أي قاعدة."}), 400
    row_id = data.get("id")
    col = data.get("col")
    value = data.get("value", "")
    if row_id is None or not col:
        return jsonify({"error": "بيانات ناقصة."}), 400
    try:
        core.update_cell(entry["db_path"], row_id, col, value)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"ok": True})


@app.route("/api/clear_override", methods=["POST"])
@serialized
def api_clear_override():
    data = request.get_json(force=True)
    cfg = load_app_config()
    _, entry = get_active_entry(cfg)
    if not entry:
        return jsonify({"error": "لسه ماتمش إعداد أي قاعدة."}), 400
    row_id = data.get("id")
    col = data.get("col")
    if row_id is None or not col:
        return jsonify({"error": "بيانات ناقصة."}), 400
    try:
        core.clear_override(entry["db_path"], row_id, col)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"ok": True})


@app.route("/api/add_column", methods=["POST"])
@serialized
def api_add_column():
    data = request.get_json(force=True)
    cfg = load_app_config()
    _, entry = get_active_entry(cfg)
    if not entry:
        return jsonify({"error": "لسه ماتمش إعداد أي قاعدة."}), 400
    name = (data.get("name") or "").strip()
    default = data.get("default", "")
    if not name:
        return jsonify({"error": "اكتب اسم العمود."}), 400
    try:
        clean_name = core.add_local_column(entry["db_path"], name, default)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"ok": True, "column": clean_name})


@app.route("/")
def index():
    response = app.make_response(render_template('index.html', app_version=APP_VERSION))
    response.headers['Cache-Control'] = 'no-store'
    return response


@app.route('/api/health')
def health():
    return jsonify(ok=True, version=APP_VERSION)


# ----------------------------------------------------------------------
# الواجهة في templates و static؛ تُضمّن عند بناء EXE
# ----------------------------------------------------------------------



def _open_browser():
    time.sleep(1.2)
    webbrowser.open(f"http://127.0.0.1:{APP_PORT}")


if __name__ == "__main__":
    if os.environ.get('EXCEL_SEARCH_NO_BROWSER') != '1':
        threading.Thread(target=_open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=APP_PORT, debug=False)
