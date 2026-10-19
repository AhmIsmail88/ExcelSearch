"""Explicit quantity mappings, independent of matching and display columns."""
import math

MODES = {'yearly': 'سنة بسنة مع الإجمالي', 'total': 'إجمالي السنوات المختارة',
         'custom': 'سنوات مخصصة'}


def configure(columns, mode, pairs, display):
    if mode not in MODES:
        raise ValueError('طريقة مقارنة الكميات غير صالحة.')
    if not isinstance(pairs, list) or not 1 <= len(pairs) <= 12:
        raise ValueError('أضف من سنة واحدة إلى 12 سنة للمقارنة.')
    result = []
    for pair in pairs:
        if not isinstance(pair, dict) or any(pair.get(tag) not in columns[tag] for tag in ('a', 'b')):
            raise ValueError('اختر عمود الكمية المقابل في كل ملف لكل سنة.')
        result.append((pair['a'], pair['b']))
    if any(len({p[i] for p in result}) != len(result) for i in (0, 1)):
        raise ValueError('لا تكرر عمود كمية في أكثر من سنة؛ حتى لا يتكرر احتسابه في الإجمالي.')
    for tag in ('a', 'b'):
        selected = display[tag]
        if not isinstance(selected, list) or len(selected) > 20 or any(not isinstance(c, str) or c not in columns[tag] for c in selected):
            raise ValueError('أعمدة العرض في التقرير غير صالحة أو غير موجودة في الملف.')
        display[tag] = list(dict.fromkeys(selected))
    return result


def compare_quantities(row_a, row_b, pairs, mode, parse_number, both=True):
    cols, totals, filled, invalid = [], {'a': 0.0, 'b': 0.0}, {'a': 0, 'b': 0}, False
    different, distribution = False, False
    for ca, cb in pairs:
        texts = {tag: '' if row.get(col) is None else str(row.get(col)).strip()
                 for tag, row, col in [('a', row_a, ca), ('b', row_b, cb)]}
        nums = {tag: parse_number(texts[tag]) for tag in ('a', 'b')}
        bad = any(texts[tag] and (nums[tag] is None or not math.isfinite(nums[tag])) for tag in ('a', 'b'))
        invalid |= bad
        for tag in ('a', 'b'):
            if nums[tag] is not None and math.isfinite(nums[tag]):
                totals[tag] += nums[tag]; filled[tag] += 1
            else:
                nums[tag] = None
        na, nb = nums['a'], nums['b']
        equal = (na is not None and nb is not None and abs(nb-na) <= 1e-6) or (not texts['a'] and not texts['b'])
        different |= not equal
        distribution |= abs((nb or 0)-(na or 0)) > 1e-6
        cols.append({'name': ca if ca == cb else ca+' ↔ '+cb, 'a': texts['a'], 'b': texts['b'],
                     'a_num': na, 'b_num': nb, 'delta': round(nb-na, 6) if na is not None and nb is not None and both else None,
                     'changed': both and not equal, 'invalid': bad})
    usable = both and not invalid and all(filled.values())
    delta = round(totals['b']-totals['a'], 6) if usable else None
    same_total = usable and abs(delta) <= 1e-6
    if not both:
        status = 'بند موجود في ملف واحد ضمن النطاق'
    elif invalid:
        status = 'قيمة كمية غير رقمية — تحتاج مراجعة'
    elif not all(filled.values()):
        status = 'بيانات الكميات غير مكتملة — لا يمكن تأكيد الإجمالي'
    elif same_total and distribution:
        status = 'الإجمالي متساوٍ — توزيع السنوات مختلف' if mode != 'custom' else 'الإجمالي متساوٍ — توزيع الأعمدة المختارة مختلف'
    elif same_total and different:
        status = 'الإجمالي متساوٍ — توجد فروق بين الفراغ والصفر'
    elif same_total:
        status = 'الكميات متساوية'
    else:
        status = 'الإجمالي مختلف'
    changed = both and (not usable or (not same_total if mode == 'total' else different))
    return {'cols': cols, 'a_total': round(totals['a'], 6) if filled['a'] and not invalid else None,
            'b_total': round(totals['b'], 6) if filled['b'] and not invalid else None,
            'delta_total': delta, 'a_filled': filled['a'], 'b_filled': filled['b'],
            'changed': changed, 'distribution_changed': both and distribution,
            'status': status, 'mode': mode, 'incomplete': not usable}
