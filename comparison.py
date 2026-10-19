"""Disk-backed comparison snapshots and explicitly reviewed text suggestions."""
import json
import re
import sqlite3
import tempfile
import math
from datetime import datetime, timezone
from pathlib import Path
from difflib import SequenceMatcher

from excel_db_core import normalize_arabic


def norm(value):
    return normalize_arabic(value).casefold()


_ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")


def parse_number(value):
    """يحوّل النص إلى رقم إن أمكن: أرقام عربية-هندية، فواصل آلاف، ومسافات."""
    if value is None:
        return None
    text = str(value).translate(_ARABIC_DIGITS)
    text = re.sub(r"\s+", "", text)
    text = text.replace("٬", "").replace("،", "").replace(",", "").replace("٫", ".")
    if not text:
        return None
    try:
        number = float(text)
        return number if math.isfinite(number) else None
    except ValueError:
        return None


class Comparison:
    def __init__(self, path_a, path_b, key_a, key_b, query="", mode="key",
                 threshold=65, progress=None, filters_a=None, filters_b=None,
                 field_a=None, field_b=None, name_a='الملف الأول', name_b='الملف الثاني',
                 quantity_mode=None, quantity_pairs=None, display_columns_a=None, display_columns_b=None,
                 key_columns_a=None, key_columns_b=None):
        if mode not in ("key", "similarity"):
            raise ValueError("نوع المقارنة غير معروف.")
        if not 0 <= float(threshold) <= 100:
            raise ValueError("نسبة التشابه يجب أن تكون بين صفر و100.")
        self.key_columns = {'a': key_columns_a if key_columns_a is not None else [key_a],
                            'b': key_columns_b if key_columns_b is not None else [key_b]}
        for columns in self.key_columns.values():
            if not isinstance(columns,list) or not 1 <= len(columns) <= 6 or any(not isinstance(c,str) or not c for c in columns) or len(set(columns)) != len(columns):
                raise ValueError('اختر من عمود واحد إلى 6 أعمدة ربط مختلفة لكل ملف.')
        if len(self.key_columns['a']) != len(self.key_columns['b']):
            raise ValueError('عدد أعمدة الربط يجب أن يكون متساويًا في الملفين وبنفس الترتيب.')
        if mode=='similarity' and any(len(cols)!=1 for cols in self.key_columns.values()):
            raise ValueError('تشابه الوصف يستخدم عمود وصف واحد؛ الربط المركب متاح في وضع المفتاح.')
        key_a,key_b=self.key_columns['a'][0],self.key_columns['b'][0]
        self.temp = tempfile.TemporaryDirectory(prefix="excel-compare-")
        self.path = str(Path(self.temp.name) / "snapshot.sqlite")
        self.mode, self.query = mode, query
        self.filters = {'a': [] if filters_a is None else filters_a, 'b': [] if filters_b is None else filters_b}
        self.names = {'a': name_a, 'b': name_b}
        self.quantity_mode = quantity_mode
        self.display_columns = {'a': display_columns_a if display_columns_a is not None else [],
                                'b': display_columns_b if display_columns_b is not None else []}
        self.report = progress or (lambda message: None)
        self.summary = {"mode": mode, "query": query, "blank_a": 0, "blank_b": 0,
                        'filters_a':self.filters['a'], 'filters_b':self.filters['b'],
                        'key_a':key_a,'key_b':key_b,'key_columns_a':self.key_columns['a'],
                        'key_columns_b':self.key_columns['b'],'field_a':field_a,'field_b':field_b}
        self.columns = {}
        conn = sqlite3.connect(self.path)
        try:
            conn.execute("PRAGMA temp_store=FILE")
            conn.execute("CREATE TABLE results(id INTEGER PRIMARY KEY, section TEXT, aid INTEGER, bid INTEGER, score REAL, qty_changed INTEGER, body TEXT, note TEXT NOT NULL DEFAULT '')")
            for tag, path, key in (("a", path_a, key_a), ("b", path_b, key_b)):
                self._snapshot(conn, tag, path, key)
            self.shared = [c for c in self.columns['a'] if c in self.columns['b'] and c not in self.key_columns['a']+self.key_columns['b']]
            self.qty_pairs = self._qty_pairs()
            if quantity_mode:
                from quantity_comparison import configure
                self.qty_pairs = configure(self.columns, quantity_mode, quantity_pairs, self.display_columns)
                self.summary.update(quantity_mode=quantity_mode, quantity_pairs=quantity_pairs,
                                    display_columns=self.display_columns)
            self.has_qty = bool(self.qty_pairs)
            if bool(field_a) != bool(field_b):
                raise ValueError('اختر عمود المقارنة في الملفين معًا أو اتركهما لمقارنة كل الأعمدة المشتركة.')
            if field_a and (field_a not in self.columns['a'] or field_b not in self.columns['b']):
                raise ValueError('عمود المقارنة المختار غير موجود.')
            self.field_pairs = [(field_a,field_b)] if field_a else [(c,c) for c in self.shared]
            if mode=='similarity':
                self.field_pairs = [(key_a,key_b)] + [(c,c) for c in self.shared]
            if quantity_mode:
                self.field_pairs = self.qty_pairs
            self.summary['shared_columns'] = self.shared
            if mode == 'key':
                if not self.field_pairs:
                    raise ValueError("لا توجد أعمدة مشتركة للمقارنة غير المفتاح. اختار أعمدة مناسبة أو استخدم وضع تشابه الوصف.")
                self._exact(conn)
            else:
                self._similar(conn, float(threshold))
                self.summary['threshold'] = threshold
            conn.execute("CREATE INDEX results_section ON results(section,id)")
            conn.execute("CREATE UNIQUE INDEX accepted_a ON results(aid) WHERE section='accepted'")
            conn.execute("CREATE UNIQUE INDEX accepted_b ON results(bid) WHERE section='accepted'")
            conn.commit()
        except BaseException:
            conn.close()
            self.close()
            raise
        finally:
            conn.close()

    def _snapshot(self, target, tag, path, key):
        self.report(f"قراءة لقطة الملف {'الأول' if tag == 'a' else 'الثاني'}...")
        source = sqlite3.connect(Path(path).resolve().as_uri() + '?mode=ro', uri=True)
        source.row_factory = sqlite3.Row
        terms = norm(self.query).split()
        try:
            source.execute('BEGIN')
            cols = [r[1] for r in source.execute('PRAGMA table_info(data)')
                    if r[1] not in ('id', '_hash', '_key', '_overrides')]
            if any(c not in cols for c in self.key_columns[tag]):
                raise ValueError('أحد أعمدة الربط غير موجود في الملف. راجع الأعمدة المختارة.')
            filters=self.filters[tag]
            if not isinstance(filters,list) or len(filters)>12:
                raise ValueError('قائمة فلاتر المقارنة غير صالحة.')
            conditions,params=[],[]
            for item in filters:
                from row_filters import predicate
                clause, values = predicate(item, cols, norm)
                conditions.append(clause)
                params.extend(values)
            source.create_function('norm',1,norm)
            where=' AND '.join(conditions) or '1=1'
            self.summary['source_count_'+tag]=source.execute('SELECT COUNT(*) FROM data').fetchone()[0]
            self.summary['filtered_count_'+tag]=source.execute('SELECT COUNT(*) FROM data WHERE '+where,params).fetchone()[0]
            self.columns[tag] = cols
            target.execute(f'CREATE TABLE {tag}(id INTEGER PRIMARY KEY, k TEXT, description TEXT, body TEXT, matched INTEGER)')
            insert = target.cursor()
            for number, row in enumerate(source.execute('SELECT * FROM data WHERE '+where+' ORDER BY id',params), 1):
                if number % 250 == 0:
                    self.report(f"قراءة {number} صف من الملف {tag}...")
                body = {c: row[c] for c in cols}
                parts = [str(body[c]).strip() if body[c] is not None else '' for c in self.key_columns[tag]]
                if not all(parts):
                    self.summary['blank_' + tag] += 1
                    continue
                raw = parts[0] if len(parts)==1 else json.dumps(parts,ensure_ascii=False,separators=(',',':'))
                # Codes are literal: Arabic spelling variants must not merge distinct codes.
                description = norm(raw)
                text = norm(' '.join(str(v or '') for v in body.values())) if terms else ''
                matched = all(term in text for term in terms)
                insert.execute(f'INSERT INTO {tag} VALUES (?,?,?,?,?)',
                               (row['id'], raw, description, json.dumps(body, ensure_ascii=False), int(matched)))
            target.execute(f'CREATE INDEX {tag}_key ON {tag}(k)')
            if self.mode == 'key':
                duplicates = target.execute(f'SELECT k,COUNT(*) FROM {tag} GROUP BY k HAVING COUNT(*)>1 LIMIT 4').fetchall()
                if duplicates:
                    sample = '، '.join(f'{k} ({n} صفوف)' for k, n in duplicates)
                    raise ValueError(f"مفتاح الملف {'الأول' if tag == 'a' else 'الثاني'} مكرر: {sample}. اختار مفتاحًا فريدًا أو استخدم اقتراحات تشابه الوصف.")
        finally:
            source.close()

    def _qty_columns(self, tag):
        """أعمدة الكميات: كل عمود اسمه بعد التطبيع يحتوي «سنه» أو «كميه»."""
        columns = self.columns.get(tag) or []
        ordinal = ('الاول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس')

        def order(name):
            n = norm(name)
            for index, word in enumerate(ordinal, 1):
                if word in n:
                    return index
            return len(ordinal) + 1

        picked = [c for c in columns if 'سنه' in norm(c) or 'كميه' in norm(c)]
        return sorted(picked, key=lambda c: (order(c), columns.index(c)))

    def _qty_pairs(self):
        cols_a, cols_b = self._qty_columns('a'), self._qty_columns('b')
        by_norm = {}
        for c in cols_b:
            by_norm.setdefault(norm(c), c)
        pairs, used_b = [], set()
        for c in cols_a:
            match = by_norm.get(norm(c))
            pairs.append((c, match))
            if match:
                used_b.add(match)
        for c in cols_b:
            if c not in used_b:
                pairs.append((None, c))
        return pairs

    def _body(self, a, b, key):
        row_a, row_b = json.loads(a) if a else {}, json.loads(b) if b else {}
        diffs = {}
        for ca, cb in self.field_pairs:
            text_a = str(row_a.get(ca) or '').strip()
            text_b = str(row_b.get(cb) or '').strip()
            if text_a == text_b:
                continue
            num_a, num_b = parse_number(text_a), parse_number(text_b)
            if num_a is not None and num_b is not None and abs(num_a - num_b) <= 1e-6:
                continue
            entry = {'a': text_a, 'b': text_b}
            if num_a is not None and num_b is not None:
                entry['delta'] = round(num_b - num_a, 6)
            diffs[ca if ca == cb else ca + ' ↔ ' + cb] = entry
        if self.mode=='key' and len(self.key_columns['a'])>1:
            key=' | '.join(json.loads(key))
        body = {'key': key, 'row_a': row_a, 'row_b': row_b, 'diffs': diffs}
        if self.has_qty:
            if self.quantity_mode:
                from quantity_comparison import compare_quantities
                body['qty'] = compare_quantities(row_a, row_b, self.qty_pairs, self.quantity_mode,
                                                parse_number, both=bool(a) and bool(b))
            else:
                body['qty'] = self._qty_block(row_a, row_b, both=bool(a) and bool(b))
        return body

    def _qty_block(self, row_a, row_b, both=True):
        """ملخص الكميات: عمود لكل سنة/كمية + الإجماليات + الفرق (الثاني - الأول)."""
        cols, changed = [], False
        total_a = total_b = 0.0
        filled_a = filled_b = 0
        for ca, cb in self.qty_pairs:
            text_a = str((row_a.get(ca) if ca else '') or '').strip()
            text_b = str((row_b.get(cb) if cb else '') or '').strip()
            num_a, num_b = parse_number(text_a), parse_number(text_b)
            if num_a is not None:
                total_a += num_a
                filled_a += 1
            if num_b is not None:
                total_b += num_b
                filled_b += 1
            delta = None
            if num_a is not None and num_b is not None:
                delta = round(num_b - num_a, 6)
            if both and ca and cb and (text_a or text_b):
                if text_a == text_b:
                    pass
                elif num_a is not None and num_b is not None:
                    if abs(num_a - num_b) > 1e-6:
                        changed = True
                elif (num_a is None) != (num_b is None):
                    changed = True
                else:
                    changed = True
            cols.append({'name': ca or cb, 'a': text_a, 'b': text_b,
                         'a_num': num_a, 'b_num': num_b, 'delta': delta})
        return {'cols': cols, 'a_total': round(total_a, 6), 'b_total': round(total_b, 6),
                'delta_total': round(total_b - total_a, 6),
                'a_filled': filled_a, 'b_filled': filled_b, 'changed': changed}

    def _exact(self, conn):
        self.report("مقارنة الصفوف وحفظ النتائج على القرص...")
        reader = conn.execute('SELECT a.id,b.id,a.k,a.body,b.body FROM a JOIN b ON a.k=b.k WHERE a.matched OR b.matched ORDER BY a.id')
        for number, (aid, bid, key, a, b) in enumerate(reader, 1):
            if number % 250 == 0:
                self.report(f"مقارنة {number} صف مشترك...")
            body = self._body(a, b, key)
            changed = body['qty']['changed'] if self.quantity_mode else bool(body['diffs'])
            section = 'changed' if changed else 'identical'
            conn.execute('INSERT INTO results(section,aid,bid,qty_changed,body) VALUES (?,?,?,?,?)',
                         (section, aid, bid, 1 if body.get('qty', {}).get('changed') else 0,
                          json.dumps(body, ensure_ascii=False)))
        for tag, other in (('a', 'b'), ('b', 'a')):
            for rid, key, body in conn.execute(f'SELECT id,k,body FROM {tag} WHERE matched AND NOT EXISTS(SELECT 1 FROM {other} WHERE {other}.k={tag}.k) ORDER BY id'):
                item = self._body(body if tag == 'a' else None, body if tag == 'b' else None, key)
                conn.execute('INSERT INTO results(section,aid,bid,qty_changed,body) VALUES (?,?,?,?,?)',
                             ('only_' + tag, rid if tag == 'a' else None, rid if tag == 'b' else None, 0,
                              json.dumps(item, ensure_ascii=False)))

    def _similar(self, conn, threshold):
        self.report("اقتراح تطابقات الوصف؛ كل اقتراح يحتاج مراجعتك...")
        conn.execute('CREATE VIRTUAL TABLE descriptions USING fts5(text)')
        conn.execute('INSERT INTO descriptions(rowid,text) SELECT id,description FROM b')
        examined = 0
        for aid, description, a, matched in conn.execute('SELECT id,description,body,matched FROM a ORDER BY id'):
            # Rare/long terms improve candidate retrieval; cap candidates, not source rows.
            words = sorted(set(re.findall(r'\w+', description)), key=lambda word: (-len(word), word))[:32]
            if not words:
                continue
            match = ' OR '.join('"' + word.replace('"', '""') + '"' for word in words)
            candidates = conn.execute('SELECT b.id,b.description,b.body,b.matched FROM descriptions JOIN b ON b.id=descriptions.rowid WHERE descriptions MATCH ? ORDER BY bm25(descriptions),b.id LIMIT 25', (match,)).fetchall()
            ranked = []
            for bid, other, b, bmatched in candidates:
                if not (matched or bmatched):
                    continue
                score = round(100 * SequenceMatcher(None, description, other).ratio(), 1)
                if score >= threshold:
                    ranked.append((score, bid, b))
            for score, bid, b in sorted(ranked, key=lambda item: (-item[0], item[1]))[:3]:
                item = self._body(a, b, f'{aid} ⇔ {bid}')
                conn.execute('INSERT INTO results(section,aid,bid,score,qty_changed,body) VALUES (?,?,?,?,?,?)',
                             ('suggested', aid, bid, score, 1 if item.get('qty', {}).get('changed') else 0,
                              json.dumps(item, ensure_ascii=False)))
            examined += 1
            if examined % 250 == 0:
                self.report(f"تم فحص وصف {examined} صف...")
        self.summary['candidate_limit'] = 25
        self.summary['suggestions_per_row'] = 3
        self.summary['without_suggestion_a'] = conn.execute('SELECT COUNT(*) FROM a WHERE matched AND NOT EXISTS(SELECT 1 FROM results WHERE aid=a.id)').fetchone()[0]

    def counts(self, conn):
        return dict(conn.execute('SELECT section,COUNT(*) FROM results GROUP BY section'))

    def qty_counts(self, conn):
        return dict(conn.execute('SELECT section,COUNT(*) FROM results WHERE qty_changed=1 GROUP BY section'))

    def page(self, section=None, offset=0, limit=20, qty_only=False):
        section = section or ('suggested' if self.mode == 'similarity' else 'changed')
        allowed = ('suggested', 'accepted', 'rejected') if self.mode == 'similarity' else ('changed', 'identical', 'only_a', 'only_b')
        if section not in allowed or offset < 0 or not 1 <= limit <= 100:
            raise ValueError('صفحة مقارنة غير صالحة.')
        conn = sqlite3.connect(self.path)
        try:
            where = 'section=?' + (' AND qty_changed=1' if qty_only else '')
            items = []
            for rid, aid, bid, score, body, note in conn.execute(f'SELECT id,aid,bid,score,body,note FROM results WHERE {where} ORDER BY id LIMIT ? OFFSET ?', (section, limit, offset)):
                item = json.loads(body)
                item.update(id=rid, row_id_a=aid, row_id_b=bid, score=score, status=section, note=note)
                from review_text import segments
                review_pairs=list(zip(self.key_columns['a'],self.key_columns['b']))
                if self.quantity_mode:
                    review_pairs += [(c,c if c in self.columns['b'] else None) for c in self.display_columns['a']]
                    review_pairs += [(c if c in self.columns['a'] else None,c) for c in self.display_columns['b']]
                elif self.mode=='key':review_pairs += self.field_pairs
                if self.mode=='similarity':
                    review_pairs += [(c,c) for c in self.shared if 'وحد' in norm(c) or norm(c) in ('unit','code','الكود','كود')]
                item['review_fields']=[]
                for ca,cb in dict.fromkeys(review_pairs):
                    va,vb=item['row_a'].get(ca,''),item['row_b'].get(cb,'')
                    sa,sb=segments(va,vb)
                    item['review_fields'].append({'a_column':ca,'b_column':cb,'a':va,'b':vb,'a_segments':sa,'b_segments':sb})
                items.append(item)
            counts = self.counts(conn)
            total = conn.execute(f'SELECT COUNT(*) FROM results WHERE {where}', (section,)).fetchone()[0]
            return {**self.summary, 'counts': counts, 'qty_counts': self.qty_counts(conn),
                    'has_qty': self.has_qty, 'qty_only': bool(qty_only), 'items': items,
                    'section': section, 'offset': offset, 'limit': limit, 'total': total}
        finally:
            conn.close()

    def decide(self, rid, decision, note=None):
        if self.mode != 'similarity' or decision not in ('accepted', 'rejected', 'suggested'):
            raise ValueError('قرار مراجعة غير صالح.')
        if note is not None and (not isinstance(note,str) or len(note)>2000):
            raise ValueError('ملاحظة المراجعة يجب ألا تتجاوز 2000 حرف.')
        conn = sqlite3.connect(self.path)
        try:
            if not conn.execute('SELECT 1 FROM results WHERE id=?', (rid,)).fetchone():
                raise ValueError('الاقتراح غير موجود.')
            conn.execute('UPDATE results SET section=? WHERE id=?', (decision, rid))
            if note is not None:
                conn.execute('UPDATE results SET note=? WHERE id=?',(note,rid))
            if self.temp is None:
                from comparison_sessions import update_metadata
                update_metadata(conn, updated_at=datetime.now(timezone.utc).isoformat())
            conn.commit()
        except sqlite3.IntegrityError:
            raise ValueError('أحد الصفين مرتبط بالفعل بتطابق مقبول. ألغِ قبول التطابق السابق أولًا.') from None
        finally:
            conn.close()

    def export(self, destination):
        from comparison_report import export_comparison
        export_comparison(self, destination)

    def close(self):
        if self.temp is not None:
            self.temp.cleanup()
