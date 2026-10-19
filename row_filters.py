"""Shared predicates for search, comparison and exports (norm is supplied by caller)."""
OPERATORS = ('contains', 'not_contains', 'equals', 'empty', 'not_empty')


def validate(item):
    if not isinstance(item, dict) or item.get('op', 'contains') not in OPERATORS:
        raise ValueError('شرط فلترة غير صالح.')
    scope = item.get('scope', 'column')
    if scope not in ('column', 'all'):
        raise ValueError('نطاق الفلترة غير صالح.')
    if scope == 'all' and item.get('op', 'contains') not in ('contains', 'not_contains'):
        raise ValueError('لكل الأعمدة اختر يحتوي أو لا يحتوي.')
    if item.get('op', 'contains') in ('contains', 'not_contains', 'equals') and not str(item.get('value') or '').strip():
        raise ValueError('اكتب الكلمة المطلوبة في الفلتر.')


def missing_column(item, columns):
    validate(item)
    return item.get('scope') != 'all' and item.get('col') not in columns


def predicate(item, columns, normalize):
    if missing_column(item, columns):
        raise ValueError('عمود الفلترة غير موجود. راجع الشروط المحفوظة.')
    op = item.get('op', 'contains')
    value = normalize(item.get('value', ''))
    names = list(columns) if item.get('scope') == 'all' else [item['col']]
    quoted = ['"'+name.replace('"', '""')+'"' for name in names]
    if op in ('empty', 'not_empty'):
        return f"norm({quoted[0]}) {'=' if op == 'empty' else '<>'} ''", []
    if not value:
        raise ValueError('اكتب الكلمة المطلوبة في الفلتر.')
    if op == 'equals':
        return f'norm({quoted[0]}) = ?', [value]
    # Balance the OR tree so wide Excel sheets do not hit SQLite's expression depth.
    clauses = [f'instr(norm({col}), ?) > 0' for col in quoted]
    while len(clauses) > 1:
        clauses = ['('+ ' OR '.join(clauses[i:i+2]) +')' for i in range(0, len(clauses), 2)]
    expression = clauses[0] if clauses else '0'
    if op == 'not_contains':
        expression = f'NOT ({expression})'
    return expression, [value]*len(quoted)
