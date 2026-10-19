"""Bounded-memory search exports, including multiple files and filter scope."""
import sqlite3
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
import excel_db_core as core
from row_filters import validate, missing_column


def write_search_export(entries, query, filters, output):
    if not isinstance(filters, list) or len(filters) > 12:
        raise ValueError('قائمة فلاتر غير صالحة.')
    for item in filters:
        validate(item)
    wb=Workbook(write_only=True)
    summary=wb.create_sheet('الملفات')
    def append(sheet,values):
        cells=[]
        for value in values:
            cell=WriteOnlyCell(sheet,value=value)
            if isinstance(value,str):
                cell.data_type='s'
            cells.append(cell)
        sheet.append(cells)
    append(summary,['الملف','ورقة النتائج','عدد النتائج','الحالة'])
    for number,entry in enumerate(entries,1):
        path=entry['db_path']
        allowed=set(core.get_columns(path))
        if any(missing_column(item, allowed) for item in filters):
            append(summary,[entry.get('name',''),'',0,'مستبعد: أحد أعمدة الفلترة غير موجود'])
            continue
        rows,total=core.query_rows(path,query,filters,limit=500)
        conn=sqlite3.connect(path)
        try:
            columns=[row[1] for row in conn.execute('PRAGMA table_info(data)') if row[1] not in ('_hash','_key','_overrides')]
        finally:
            conn.close()
        title=f'النتائج {number}'
        sheet=wb.create_sheet(title);append(sheet,columns)
        append(summary,[entry.get('name',''),title,total,'تم'])
        offset,line,part=0,1,1
        while rows:
            for row in rows:
                if line>=1048576:
                    part+=1;sheet=wb.create_sheet(f'{title}-{part}');append(sheet,columns);line=1
                append(sheet,[row.get(col,'') for col in columns]);line+=1
            offset+=len(rows)
            if offset>=total:
                break
            rows,_=core.query_rows(path,query,filters,limit=500,offset=offset)
    wb.save(output)
