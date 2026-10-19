import sqlite3
import unittest
from row_filters import predicate
from excel_db_core import normalize_arabic


class RowFilterTests(unittest.TestCase):
    def test_wide_rows_nulls_and_literal_characters(self):
        conn=sqlite3.connect(':memory:')
        try:
            conn.create_function('norm',1,lambda v:normalize_arabic(v).casefold())
            columns=[f'c{i}' for i in range(1100)]
            conn.execute('CREATE TABLE data('+','.join('"'+c+'" TEXT' for c in columns)+')')
            conn.execute('INSERT INTO data(c1099) VALUES(?)',('المسجد %_quoted',))
            conn.execute('INSERT INTO data(c1099) VALUES(NULL)')
            clause,params=predicate({'scope':'all','op':'not_contains','value':'%_'},columns,lambda v:normalize_arabic(v).casefold())
            self.assertEqual(conn.execute('SELECT COUNT(*) FROM data WHERE '+clause,params).fetchone()[0],1)
            self.assertEqual(conn.execute('SELECT c1099 FROM data WHERE '+clause,params).fetchone()[0],None)
        finally:conn.close()
