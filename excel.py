"""Dependency-free XLSX writer with cached formula values and editable periods."""
import io
import re
import zipfile
from collections import Counter, defaultdict
from datetime import date
from decimal import Decimal
from pathlib import Path
from xml.sax.saxutils import escape

from engine import (BOUTIQUES, UNKNOWN, PROFILE_FIELDS, PROFILE_NOTE, RETURN_NOTE,
                    SALES_HEAD, REG_HEAD, STAFF_HEAD, REG_STAFF_HEAD, sales_values,
                    reg_values, period, tr)

STYLES = {'body': 3, 'title': 5, 'section': 7, 'header': 12, 'note': 15,
          'integer': 18, 'money': 20, 'decimal': 22, 'percent': 24, 'date': 26,
          'inputDate': 30, 'total': 33, 'totalMoney': 35, 'totalPercent': 37}
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


def xml(value):
    return escape(re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(value)), {'"': '&quot;', "'": '&apos;'})


def col(index):
    result = ''
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def serial(day):
    return (day - date(1899, 12, 30)).days


class Sheet:
    def __init__(self, name, widths):
        self.name, self.widths = name, widths
        self.cells, self.heights, self.tables, self.merges = {}, {}, [], []
        self.validation = None
        self.freeze = 0

    def put(self, row, column, value, style='body', formula=None):
        self.cells[row, column] = ('' if value is None else value, style, formula)

    def row(self, row, values, start=1, style='body'):
        for i, value in enumerate(values):
            self.put(row, start + i, value, style)

    def note(self, row, text, end=7):
        self.put(row, 1, text, 'note')
        self.merges.append(f'B{row}:{col(end)}{row}')
        self.heights[row] = 44

    def table(self, row, column, headers, count, name):
        self.row(row, headers, column, 'header')
        self.heights[row] = 64
        self.tables.append((name, f'{col(column)}{row}:{col(column+len(headers)-1)}{row+max(1,count)}', headers))


def make_sheets(prepared, from_date, to_date, lang='ru'):
    a, z = period(from_date, to_date)
    T = lambda v: tr(v, lang)
    L = lambda ru, en: ru if lang == 'ru' else en
    na = T('Нет данных')
    facts, events = prepared['facts'], prepared['events']
    selected = [x for x in facts if a <= x['day'] <= z]
    chosen = [x for x in events if a <= x['day'] <= z]
    sales = Sheet('Sales — Adapted', [3, 29, 34, 24, 26, 25, 27, 24])
    regs = Sheet('Registrations — Adapted', [3, 29, 34, 30, 32, 25, 24, 24, 24] + [22] * 6)
    lines = Sheet('Sales data', [22, 22, 25, 32, 24, 24, 20] + [24] * 11)
    registrations = Sheet('Registration data', [24, 24, 29, 34, 28, 24, 24])
    profiles = Sheet('Customer profile data', [24] + [22] * 7)
    sheets = [sales, regs, lines, registrations, profiles]
    N, R = max(5, len(facts) + 4), max(5, len(events) + 4)
    sr = lambda c: f"'Sales data'!${c}$5:${c}${N}"
    er = lambda c: f"'Registration data'!${c}$5:${c}${R}"
    pr = lambda c: f"'Customer profile data'!${c}$5:${c}${R}"
    base = f'{sr("B")},">="&$C$4,{sr("B")},"<="&$E$4'
    ebase = f'{er("B")},">="&$C$4,{er("B")},"<="&$E$4'
    linked = f'{sr("B")},">="&\'Sales — Adapted\'!$C$4,{sr("B")},"<="&\'Sales — Adapted\'!$E$4'
    all_dates = [d for limits in prepared['ranges'].values() for d in limits]
    lo, hi = serial(min(all_dates)), serial(max(all_dates))
    guard = f'OR(NOT(ISNUMBER(C4)),NOT(ISNUMBER(E4)),C4>E4,C4<{lo},E4>{hi})'
    valid = lambda formula: f'IF({guard},NA(),{formula})'

    for sheet, title in [(sales, 'Продажи по бутикам'), (regs, 'Регистрации и покупки')]:
        sheet.put(2, 1, T(title), 'title')
        sheet.merges.append('B2:H2')
        sheet.heights[2] = 32
        sheet.put(4, 1, T('Начало периода'))
        sheet.put(4, 2, serial(a), 'inputDate')
        sheet.put(4, 3, T('Конец периода'))
        sheet.put(4, 4, serial(z), 'inputDate')
        sheet.validation = (lo, hi)
        sheet.note(5, L('Измените жёлтые даты. Таблицы пересчитаются в Excel. Даты на двух отчётных листах независимы.',
                        'Edit the yellow dates to recalculate in Excel. The two report sheets have independent dates.'))
        sheet.note(6, L('Полнота периода зависит от выгрузок. Для новых дат и статусов загрузите CSV заново.',
                        'Period completeness depends on the exports. Reload CSVs for new dates and statuses.'))
    # Data sheets keep all uploaded dates, including those outside the initial period.
    for sheet, title, note in [
        (lines, L('Учтённые позиции заказов', 'Eligible order lines'), L('Источник: CSV заказов. Москва; шесть бутиков; done; только клиенты с ID.', 'Source: order CSV. Moscow dates; six boutiques; done; identified customers only.')),
        (registrations, L('Первые доступные регистрации', 'Earliest available registrations'), L('Источники: CSV регистраций и клиентов. Сотрудник — текущее поле карточки.', 'Sources: registration and customer CSVs. Staff uses the current customer field.')),
        (profiles, T('Поля карточки клиента'), T(PROFILE_NOTE)),
    ]:
        sheet.put(1, 0, title, 'title')
        sheet.merges.append('A1:G1')
        sheet.put(2, 0, note, 'note')
        sheet.merges.append('A2:G2')
        sheet.heights[2] = 48
        sheet.freeze = 4
    lh = L(['ID заказа', 'Дата заказа (Москва)', 'Бутик', 'Продавец', 'ID клиента', 'Продажи, ₽', 'Товаров',
            'Вес заказа бутика', 'Вес клиента', 'Вес клиента бутика', 'Вес заказа продавца', 'Вес клиента продавца',
            'Дата регистрации', 'Класс регистрации', 'Вес прежнего заказа', 'Вес прежнего клиента', 'Вес прежнего клиента бутика', 'Вес заказа сети'],
           ['Order ID', 'Order date (Moscow)', 'Boutique', 'Salesperson', 'Customer ID', 'Net sales, RUB', 'Units',
            'Boutique order weight', 'Customer weight', 'Boutique customer weight', 'Staff order weight', 'Staff customer weight',
            'Registration date', 'Registration class', 'Prior registration order weight', 'Prior registration customer weight',
            'Prior registration boutique customer weight', 'Network order weight'])
    lines.table(4, 0, lh, len(facts), 'EligibleLines')
    # Distinct counts are sums of reciprocal frequencies inside the selected period.
    definitions = [(7, ('order', 'b')), (8, ('customer',)), (9, ('customer', 'b')),
                   (10, ('order', 'b', 'seller')), (11, ('customer', 'b', 'seller')),
                   (14, ('order',)), (15, ('customer',)), (16, ('customer', 'b')), (17, ('order',))]
    field_col = {'order': 'A', 'customer': 'E', 'b': 'C', 'seller': 'D'}
    counts = {column: Counter(tuple(x[k] for k in keys) for x in selected if column not in (14, 15, 16) or x['category'] == 1)
              for column, keys in definitions}
    for i, x in enumerate(facts, 5):
        lines.row(i, [x['order'], serial(x['day']), T(x['b']), T(x['seller']), x['customer'], x['amount'], x['qty']], 0)
        lines.put(i, 1, serial(x['day']), 'date')
        lines.put(i, 5, x['amount'], 'money')
        lines.put(i, 12, serial(x['reg_day']) if x['reg_day'] else '', 'date')
        lines.put(i, 13, x['category'], 'integer')
        for column, keys in definitions:
            returning = column in (14, 15, 16)
            active = a <= x['day'] <= z and (not returning or x['category'] == 1)
            count = counts[column][tuple(x[k] for k in keys)]
            criteria = ','.join(f'{sr(field_col[k])},${field_col[k]}{i}' for k in keys)
            if returning:
                criteria += f',{sr("N")},1'
            condition = f"$B{i}>='Sales — Adapted'!$C$4,$B{i}<='Sales — Adapted'!$E$4" + (f',N{i}=1' if returning else '')
            lines.put(i, column, 1 / count if active else 0, 'decimal', f'IF(AND({condition}),1/COUNTIFS({criteria},{linked}),0)')
    registrations.table(4, 0, [T('ID клиента'), L('Дата регистрации (Москва)', 'Registration date (Moscow)'),
                              T('Бутик регистрации'), T('Сотрудник'), L('Нет покупки в тот же день', 'No same-day purchase'),
                              T('Заполнено полей'), L('Неполный источник', 'Incomplete source')], len(events), 'RegistrationEvents')
    profiles.table(4, 0, [T('ID клиента')] + [T(f[0]) for f in PROFILE_FIELDS] + [T('Заполнено полей')], len(events), 'CustomerProfileData')
    for i, x in enumerate(events, 5):
        registrations.row(i, [x['customer'], serial(x['day']), T(x['b']), T(x['seller'])], 0)
        registrations.put(i, 1, serial(x['day']), 'date')
        registrations.put(i, 4, x['no'], 'integer', f'IF(COUNTIFS({sr("E")},A{i},{sr("B")},B{i})=0,1,0)')
        registrations.put(i, 5, na if x['filled'] is None else x['filled'], 'integer', f'INDEX({pr("H")},MATCH(A{i},{pr("A")},0))')
        registrations.put(i, 6, int(x['filled'] is None), 'integer', f'IF(ISNUMBER(F{i}),0,1)')
        profiles.row(i, [x['customer']] + [na if f is None else f for f in x['flags']], 0)
        profiles.put(i, 7, na if x['filled'] is None else x['filled'], 'integer', f'IF(COUNT(B{i}:G{i})=6,SUM(B{i}:G{i}),"{na}")')

    def sale_row(r, boutique, seller=None, returning=False, total=False):
        sub = [x for x in selected if (boutique is None or x['b'] == boutique) and
               (seller is None or x['seller'] == seller) and (not returning or x['category'] == 1)]
        values = sales_values(sub)
        sales.row(r, [T('Все бутики' if boutique is None else boutique)] + ([] if seller is None else [T(seller)]))
        start = 2 if seller is None else 3
        criteria = base
        if boutique is not None:
            criteria += f',{sr("C")},B{r}'
        if seller is not None:
            criteria += f',{sr("D")},C{r}'
        if returning:
            criteria += f',{sr("N")},1'
        count_cols = ('O', 'P' if boutique is None else 'Q') if returning else ('R' if boutique is None else 'H', 'I' if boutique is None else 'J')
        if seller is not None:
            count_cols = ('K', 'L')
        for j, c in enumerate(('F', 'G') + count_cols):
            f = f'SUMIFS({sr(c)},{criteria})'
            if j >= 2:
                f = f'ROUND({f},0)'
            sales.put(r, start+j, values[j], ('totalMoney' if j == 0 else 'total') if total else ('money' if j == 0 else 'integer'), valid(f))
        if seller is None:
            for c, v, numerator in [(6, values[4], 'C'), (7, values[5], 'D')]:
                sales.put(r, c, v, 'money' if c == 6 else 'decimal', f'IF(E{r}=0,"",{numerator}{r}/E{r})')

    sales.table(8, 1, list(map(T, SALES_HEAD)), 6, 'SalesSummary')
    for i, b in enumerate(BOUTIQUES, 9):
        sale_row(i, b)
    sale_row(15, None, total=True)
    sales.note(17, L('Клиенты и заказы в общем итоге уникальны по сети. Фильтры строк не изменяют общий итог.',
                     'Network totals deduplicate customers and orders. Row filters do not change the overall total.'))
    sales.put(20, 1, T('Детализация по продавцам'), 'section')
    pairs = sorted({(x['b'], x['seller']) for x in facts})
    sales.table(21, 1, list(map(T, STAFF_HEAD)), len(pairs), 'SalesByStaff')
    for i, (b, s) in enumerate(pairs, 22):
        sale_row(i, b, s)
    sales.note(24+len(pairs), T('Количество заказов и клиентов по продавцам не суммируется.'))
    sales.note(25+len(pairs), L('Сумма позиции уже включает количество. Возвраты и отменённые позиции исключены из исходных заказов.',
                              'Line amounts already include quantity. Returned and cancelled items are excluded from original orders.'))
    ret_title = 29+len(pairs)
    sales.put(ret_title, 1, T('Продажи ранее зарегистрированным клиентам'), 'section')
    sales.note(ret_title+1, T(RETURN_NOTE))
    sales.table(ret_title+3, 1, list(map(T, SALES_HEAD)), 6, 'PreviouslyRegisteredSales')
    for i, b in enumerate(BOUTIQUES + [None], ret_title+4):
        sale_row(i, b, returning=True, total=b is None)
    u = ret_title+12
    sales.note(u, T('Продажи клиентам без известной даты регистрации'), 5)
    sales.put(u, 6, sum((x['amount'] for x in selected if x['category'] == -1), Decimal(0)), 'money', valid(f'SUMIFS({sr("F")},{base},{sr("N")},-1)'))

    def reg_row(r, boutique, seller=None, total=False):
        sub = [x for x in chosen if (boutique is None or x['b'] == boutique) and (seller is None or x['seller'] == seller)]
        vals = reg_values(sub)
        regs.row(r, [T('Всего, включая неизвестный бутик' if boutique is None else boutique)] + ([] if seller is None else [T(seller)]))
        start = 2 if seller is None else 3
        c = [col(start+i) for i in range(6)]
        criteria = ebase + (f',{er("C")},B{r}' if boutique is not None else '') + (f',{er("D")},C{r}' if seller is not None else '')
        formulas = [valid(f'COUNTIFS({criteria})'), valid(f'SUMIFS({er("E")},{criteria})'),
                    f'IF({c[0]}{r}=0,"",{c[1]}{r}/{c[0]}{r})', f'{c[0]}{r}*6',
                    valid(f'IF(COUNTIFS({criteria},{er("G")},1)>0,"{na}",SUMIFS({er("F")},{criteria}))'),
                    f'IF({c[3]}{r}=0,"",IF(ISNUMBER({c[4]}{r}),{c[4]}{r}/{c[3]}{r},"{na}"))']
        for j, v in enumerate(vals):
            if j in (4, 5) and sub and any(x['filled'] is None for x in sub):
                v = na
            style = ('totalPercent' if total else 'percent') if j in (2, 5) else ('total' if total else 'integer')
            regs.put(r, start+j, v, style, formulas[j])

    regs.table(8, 1, list(map(T, REG_HEAD)), 7, 'RegistrationSummary')
    for i, b in enumerate(BOUTIQUES + [UNKNOWN], 9):
        reg_row(i, b)
    reg_row(16, None, total=True)
    regs.heights[16] = 48
    regs.note(18, T(PROFILE_NOTE), 8)
    regs.put(20, 1, T('Зарегистрировавшие сотрудники'), 'section')
    pairs = sorted({(x['b'], x['seller']) for x in events})
    regs.table(21, 1, list(map(T, REG_STAFF_HEAD)), len(pairs), 'RegistrationsByStaff')
    for i, (b, s) in enumerate(pairs, 22):
        reg_row(i, b, s)
    regs.note(24+len(pairs), L('Клиент учитывается по первой доступной регистрации. Покупка в тот же день в любом из шести бутиков исключает его из «Не купили».',
                             'Each customer uses the earliest available registration. Same-day purchases in any of the six boutiques exclude them from non-buyers.'))
    ptitle = 29+len(pairs)
    regs.put(ptitle, 1, T('Заполненность по клиентам'), 'section')
    regs.note(ptitle+1, L('Фильтр «В периоде» = 1. После изменения дат повторно примените фильтр. Поля: 1 — заполнено, 0 — пусто.',
                         'Filter In period = 1. Reapply filters after date changes. Fields: 1 = present, 0 = empty.'), 14)
    headers = [T('Бутик регистрации'), T('Сотрудник'), T('ID клиента'), L('Дата регистрации', 'Registration date'),
               L('В периоде', 'In period'), T('Возможных полей'), T('Заполнено полей'), T('Заполненность')] + [T(f[0]) for f in PROFILE_FIELDS]
    regs.table(ptitle+3, 1, headers, len(events), 'RegisteredCustomerProfiles')
    for i, x in enumerate(events, ptitle+4):
        active = a <= x['day'] <= z
        regs.row(i, [T(x['b']), T(x['seller']), x['customer'], serial(x['day'])])
        regs.put(i, 4, serial(x['day']), 'date')
        regs.put(i, 5, int(active), 'integer', valid(f'IF(AND(E{i}>=$C$4,E{i}<=$E$4),1,0)'))
        regs.put(i, 6, 6 if active else 0, 'integer', f'F{i}*6')
        regs.put(i, 7, (na if x['filled'] is None else x['filled']) if active else 0, 'integer', f'IF(F{i}=0,0,IF(COUNT(J{i}:O{i})=6,SUM(J{i}:O{i}),"{na}"))')
        regs.put(i, 8, (na if x['filled'] is None else x['filled']/6) if active else '', 'percent', f'IF(G{i}=0,"",IF(ISNUMBER(H{i}),H{i}/G{i},"{na}"))')
        for j, flag in enumerate(x['flags']):
            regs.put(i, 9+j, na if flag is None else flag, 'integer', f'INDEX({pr(col(j+1))},MATCH(D{i},{pr("A")},0))')
    return sheets


def pack(sheets):
    """Write literal strings as inlineStr: CSV values cannot become Excel formulas."""
    files, table_paths, table_id = {}, [], 0
    for si, sheet in enumerate(sheets, 1):
        rows = defaultdict(list)
        for (r, c), (value, style, formula) in sorted(sheet.cells.items()):
            attrs = f'r="{col(c)}{r}" s="{STYLES[style]}"'
            if formula:
                kind = ' t="str"' if isinstance(value, str) else ''
                cell = f'<c {attrs}{kind}><f>{xml(formula)}</f><v>{xml(value)}</v></c>'
            elif isinstance(value, (int, float, Decimal)):
                cell = f'<c {attrs}><v>{value}</v></c>'
            else:
                cell = f'<c {attrs} t="inlineStr"><is><t xml:space="preserve">{xml(value)}</t></is></c>'
            rows[r].append(cell)
        body = ''.join(f'<row r="{r}" ht="{sheet.heights.get(r,32)}" customHeight="1">{"".join(cells)}</row>' for r, cells in sorted(rows.items()))
        ids = []
        for name, ref, headers in sheet.tables:
            table_id += 1
            ids.append(table_id)
            path = f'xl/tables/table{table_id}.xml'
            table_paths.append(path)
            columns = ''.join(f'<tableColumn id="{j}" name="{xml(h)}"/>' for j, h in enumerate(headers, 1))
            files[path] = DECL + f'<table xmlns="{NS}" id="{table_id}" name="{name}" displayName="{name}" ref="{ref}" totalsRowShown="0"><autoFilter ref="{ref}"/><tableColumns count="{len(headers)}">{columns}</tableColumns><tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/></table>'
        pane = f'<pane ySplit="{sheet.freeze}" topLeftCell="A{sheet.freeze+1}" activePane="bottomLeft" state="frozen"/>' if sheet.freeze else ''
        widths = ''.join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(sheet.widths, 1))
        merges = f'<mergeCells count="{len(sheet.merges)}">' + ''.join(f'<mergeCell ref="{ref}"/>' for ref in sheet.merges) + '</mergeCells>' if sheet.merges else ''
        validations = ''
        if sheet.validation:
            lo, hi = sheet.validation
            validations = f'<dataValidations count="1"><dataValidation type="date" operator="between" showErrorMessage="1" sqref="C4 E4"><formula1>{lo}</formula1><formula2>{hi}</formula2></dataValidation></dataValidations>'
        parts = f'<tableParts count="{len(ids)}">' + ''.join(f'<tablePart r:id="rId{i}"/>' for i in range(1, len(ids)+1)) + '</tableParts>'
        files[f'xl/worksheets/sheet{si}.xml'] = DECL + f'<worksheet xmlns="{NS}" xmlns:r="{REL}"><sheetViews><sheetView showGridLines="0" workbookViewId="0">{pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="32"/><cols>{widths}</cols><sheetData>{body}</sheetData>{merges}{validations}<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>{parts}</worksheet>'
        files[f'xl/worksheets/_rels/sheet{si}.xml.rels'] = DECL + f'<Relationships xmlns="{PKG}">' + ''.join(f'<Relationship Id="rId{i}" Type="{REL}/table" Target="../tables/table{tid}.xml"/>' for i, tid in enumerate(ids, 1)) + '</Relationships>'
    files['xl/styles.xml'] = Path(__file__).with_name('styles.xml').read_text()
    files['xl/workbook.xml'] = DECL + f'<workbook xmlns="{NS}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>' + ''.join(f'<sheet name="{xml(s.name)}" sheetId="{i}" r:id="rId{i}"/>' for i, s in enumerate(sheets, 1)) + '</sheets><calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>'
    files['xl/_rels/workbook.xml.rels'] = DECL + f'<Relationships xmlns="{PKG}">' + ''.join(f'<Relationship Id="rId{i}" Type="{REL}/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets)+1)) + f'<Relationship Id="rId{len(sheets)+1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'
    files['_rels/.rels'] = DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    overrides = [('xl/workbook.xml', 'sheet.main'), ('xl/styles.xml', 'styles')] + [(f'xl/worksheets/sheet{i}.xml', 'worksheet') for i in range(1, len(sheets)+1)] + [(p, 'table') for p in table_paths]
    files['[Content_Types].xml'] = DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + ''.join(f'<Override PartName="/{path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.{kind}+xml"/>' for path, kind in overrides) + '</Types>'
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, 'w', zipfile.ZIP_DEFLATED) as archive:
        for path, content in files.items():
            archive.writestr(path, content)
    return stream.getvalue()


def export_xlsx(prepared, from_date, to_date, lang='ru'):
    return pack(make_sheets(prepared, from_date, to_date, lang))
