#!/usr/bin/env python3
"""
excel_search_db.py — أداة سطر الأوامر
=======================================
تحويل ملف إكسل إلى قاعدة SQLite قابلة للبحث السريع، مع مزامنة تزايدية.

أمثلة:
  python3 excel_search_db.py build --excel data.xlsx --db data.sqlite --unique-col "رقم_الهاتف"
  python3 excel_search_db.py sync  --db data.sqlite
  python3 excel_search_db.py search --db data.sqlite --query "كلمة"
  python3 excel_search_db.py shell --db data.sqlite
"""

import argparse
import sqlite3
import sys

import excel_db_core as core


def cmd_build(args):
    core.build_db(
        args.excel, args.db, sheet=args.sheet, text_cols=args.text_cols,
        unique_col=args.unique_col, header_row=args.header_row, progress=print,
        last_column=args.last_column,
    )


def cmd_sync(args):
    core.sync_db(args.db, excel_path=args.excel, progress=print, force=args.force)


def _print_rows(rows):
    if not rows:
        print("🚫 مفيش نتائج مطابقة.")
        return
    for i, row in enumerate(rows, 1):
        print(f"\n--- نتيجة {i} (id={row.get('id')}) ---")
        for c, v in row.items():
            if c == "id" or not v:
                continue
            print(f"  {c}: {v}")


def cmd_search(args):
    rows, total = core.search(args.db, args.query, limit=args.limit)
    print(f"🔎 عدد النتائج الكلي: {total}\n")
    _print_rows(rows)


def cmd_filter(args):
    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(f"SELECT * FROM data WHERE {args.where} LIMIT ?", (args.limit,)).fetchall()
    except sqlite3.OperationalError as e:
        sys.exit(f"❌ خطأ في الشرط: {e}")
    finally:
        conn.close()
    cleaned = []
    for r in rows:
        d = dict(r)
        d.pop("_hash", None)
        d.pop("_key", None)
        cleaned.append(d)
    _print_rows(cleaned)


def cmd_shell(args):
    print("🔎 وضع البحث التفاعلي — اكتب كلمة البحث (exit للخروج)\n")
    while True:
        try:
            q = input("بحث> ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if q.lower() in ("exit", "quit", "خروج"):
            break
        if not q:
            continue
        rows, total = core.search(args.db, q, limit=15)
        print(f"({total} نتيجة)")
        _print_rows(rows)


def cmd_stats(args):
    s = core.get_stats(args.db)
    for k, v in s.items():
        print(f"{k}: {v}")


def main():
    parser = argparse.ArgumentParser(description="تحويل إكسل إلى قاعدة SQLite مع بحث نصي سريع ومزامنة تزايدية")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("build", help="بناء القاعدة من ملف إكسل (أول مرة)")
    p.add_argument("--excel", required=True)
    p.add_argument("--db", required=True)
    p.add_argument("--sheet", default=None)
    p.add_argument("--text-cols", default=None)
    p.add_argument("--unique-col", default=None)
    p.add_argument("--header-row", type=int, default=None, help="رقم صف العناوين (1-indexed)، اختياري - افتراضياً بيتكتشف تلقائياً")
    p.add_argument("--last-column", default=None, help="آخر عمود للفهرسة، مثل N أو 14")
    p.set_defaults(func=cmd_build)

    p = sub.add_parser("sync", help="مزامنة تزايدية (بعد أول build)")
    p.add_argument("--db", required=True)
    p.add_argument("--excel", default=None, help="اختياري، لو الملف اتنقل مكان")
    p.add_argument("--force", action="store_true", help="فرض المزامنة حتى لو الملف نفس القديم")
    p.set_defaults(func=cmd_sync)

    p = sub.add_parser("search", help="بحث نصي كامل")
    p.add_argument("--db", required=True)
    p.add_argument("--query", required=True)
    p.add_argument("--limit", type=int, default=20)
    p.set_defaults(func=cmd_search)

    p = sub.add_parser("filter", help="فلترة بشرط SQL مباشر")
    p.add_argument("--db", required=True)
    p.add_argument("--where", required=True)
    p.add_argument("--limit", type=int, default=50)
    p.set_defaults(func=cmd_filter)

    p = sub.add_parser("shell", help="وضع بحث تفاعلي")
    p.add_argument("--db", required=True)
    p.set_defaults(func=cmd_shell)

    p = sub.add_parser("stats", help="عرض معلومات القاعدة")
    p.add_argument("--db", required=True)
    p.set_defaults(func=cmd_stats)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
