"""Mindbox CSV contracts and report calculations, independent of the interface."""
import csv
import io
import json
import re
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path

BOUTIQUES = ['Авиапарк', 'Галерея', 'Белая Дача', 'Пулково', 'Внуково', 'ГУМ']
UNKNOWN, NONE = 'Бутик не определён', 'Не указан'
MOSCOW = timezone(timedelta(hours=3))
REQUIRED = {
    'orders': ['OrderIdsMindboxId', 'OrderFirstActionDateTimeUtc', 'OrderFirstActionChannelName',
               'OrderCustomerIdsMindboxId', 'OrderLineQuantity', 'OrderLinePriceOfLine',
               'OrderLineStatusIdsExternalId', 'OrderLineCustomFieldsConsultantPerItem'],
    'customers': ['CustomerIdsMindboxId', 'CustomerCustomFieldsConsultantWhoRegistered'],
    'registrations': ['CustomerActionCustomerIdsMindboxId', 'CustomerActionDateTimeUtc', 'CustomerActionChannelName'],
}
PROFILE_FIELDS = [
    ('Имя', ['CustomerFirstName', 'FirstName', 'Имя']),
    ('Фамилия', ['CustomerLastName', 'LastName', 'Фамилия']),
    ('Отчество', ['CustomerMiddleName', 'MiddleName', 'Отчество']),
    ('Дата рождения', ['CustomerBirthDate', 'CustomerBirthday', 'BirthDate', 'Birthday', 'Дата рождения', 'ДР']),
    ('Телефон', ['CustomerMobilePhone', 'CustomerMobilePhoneNumber', 'MobilePhone', 'Телефон']),
    ('Email', ['CustomerEmail', 'Email', 'Электронная почта']),
]
TRANSLATIONS = json.loads(Path(__file__).with_name('translations.json').read_text())
TRANSLATIONS.update(dict(zip(
    ['Имя', 'Фамилия', 'Отчество', 'Дата рождения', 'Телефон', 'Возможных полей', 'Заполнено полей',
     'Заполненность', 'Нет данных', 'ID клиента', 'Поля карточки клиента', 'Заполненность по клиентам',
     'Продажи ранее зарегистрированным клиентам', 'Продажи клиентам без известной даты регистрации'],
    ['First name', 'Last name', 'Middle name', 'Date of birth', 'Phone', 'Possible fields', 'Completed fields',
     'Field completion rate', 'Not available', 'Customer ID', 'Customer profile fields', 'Profile completeness by customer',
     'Sales to previously registered customers', 'Sales to customers with unknown registration date'])))
SALES_HEAD = ['Бутик', 'Продажи, ₽', 'Товаров', 'Заказов', 'Покупателей', 'Средний чек, ₽', 'Товаров на заказ']
REG_HEAD = ['Бутик регистрации', 'Новых клиентов', 'Не купили в день регистрации',
            'Доля без покупки в день регистрации', 'Возможных полей', 'Заполнено полей', 'Заполненность']
STAFF_HEAD = ['Бутик', 'Продавец', 'Продажи позиций, ₽', 'Товаров', 'Заказов с позициями', 'Покупателей']
REG_STAFF_HEAD = [REG_HEAD[0], 'Сотрудник'] + REG_HEAD[1:]
PROFILE_NOTE = ('Заполненность карточки на момент выгрузки. Непустое значение считается заполненным; '
                'валидность контактов не проверяется. Отсутствие колонки — «Нет данных», а не пустое поле.')
RETURN_NOTE = ('Дата покупки позже даты первой доступной регистрации, по Москве. '
               'Это ранее зарегистрированные клиенты, не обязательно повторная покупка.')
TRANSLATIONS[PROFILE_NOTE] = ('Profile completeness at export time. Non-empty values count as completed; '
                            'contact validity is not checked. An absent column means unavailable data, not an empty field.')
TRANSLATIONS[RETURN_NOTE] = ('Purchase date is later than the earliest available registration date in Moscow time. '
                           'These customers are not necessarily repeat buyers.')


def tr(value, lang='ru'):
    if not isinstance(value, str) or lang == 'ru':
        return value
    if value in TRANSLATIONS:
        return TRANSLATIONS[value]
    for prefix in ['Не найдены колонки: ', 'Неверное число колонок, строка ', 'Неверная дата UTC: ',
                   'Неверная дата: ', 'Неверная сумма или количество: ']:
        if value.startswith(prefix):
            return TRANSLATIONS[prefix] + value[len(prefix):]
    return value


def parse_csv(content, kind):
    try:
        text = content.decode('utf-8-sig')
    except UnicodeDecodeError:
        text = content.decode('cp1251')
    # Count fields with the CSV parser, so separators inside quoted headers don't matter.
    candidates = []
    for delimiter in [';', ',', '\t']:
        try:
            header = next(csv.reader(io.StringIO(text), delimiter=delimiter, strict=True))
            candidates.append((len(header), delimiter))
        except (StopIteration, csv.Error):
            pass
    if not candidates:
        raise ValueError('CSV пуст или не содержит данных')
    delimiter = max(candidates, key=lambda x: x[0])[1]
    try:
        reader = csv.reader(io.StringIO(text, newline=''), delimiter=delimiter, strict=True)
        headers = [h.strip().lstrip('\ufeff') for h in next(reader)]
        if len(set(headers)) != len(headers) or '' in headers:
            raise ValueError('Повторяющиеся названия колонок')
        missing = set(REQUIRED[kind]) - set(headers)
        if missing:
            raise ValueError('Не найдены колонки: ' + ', '.join(sorted(missing)))
        rows = []
        for row in reader:
            if not any(row):
                continue
            if len(row) != len(headers):
                raise ValueError('Неверное число колонок, строка ' + str(reader.line_num))
            rows.append(dict(zip(headers, (cell.strip() for cell in row))))
    except csv.Error as exc:
        raise ValueError('Незакрытая кавычка в CSV') from exc
    if not rows:
        raise ValueError('CSV пуст или не содержит данных')
    return rows


def utc(value):
    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(\.\d+)?Z?', value):
        raise ValueError('Неверная дата UTC: ' + value)
    try:
        return datetime.fromisoformat(value.rstrip('Z')).replace(tzinfo=timezone.utc)
    except ValueError as exc:
        raise ValueError('Неверная дата UTC: ' + value) from exc


def moscow_day(value):
    return value.astimezone(MOSCOW).date()


def number(value):
    if not value:
        raise ValueError('Пустая сумма или количество')
    try:
        n = Decimal(re.sub(r'\s', '', value).replace(',', '.'))
        if not n.is_finite() or n < 0:
            raise InvalidOperation
        return n
    except InvalidOperation as exc:
        raise ValueError('Неверная сумма или количество: ' + value) from exc


def detect_mapping(rows):
    headers = {h.lower(): h for h in rows[0]}
    return [next((headers[a.lower()] for a in aliases if a.lower() in headers), '') for _, aliases in PROFILE_FIELDS]


def date_ranges(data):
    ranges = {}
    for kind, field in [('orders', 'OrderFirstActionDateTimeUtc'), ('registrations', 'CustomerActionDateTimeUtc')]:
        days = [moscow_day(utc(r[field])) for r in data[kind]]
        ranges[kind] = (min(days), max(days))
    return ranges


def prepare(data, mapping=None):
    """Validate the full exports, retaining all dates for editable Excel periods."""
    mapping = detect_mapping(data['customers']) if mapping is None else mapping
    if len(mapping) != 6 or len(set(filter(None, mapping))) != len(list(filter(None, mapping))):
        raise ValueError('Choose a different CSV column for each of the six profile fields / Выберите разные колонки для шести полей')
    if any(k and k not in data['customers'][0] for k in mapping):
        raise ValueError('Не найдены колонки: ' + ', '.join(k for k in mapping if k and k not in data['customers'][0]))
    customers, profiles = {}, {}
    for row in data['customers']:
        cid = row['CustomerIdsMindboxId']
        if not cid:
            raise ValueError('Клиент без ID')
        if cid in customers:
            raise ValueError('Повторяющийся ID в выгрузке клиентов')
        customers[cid] = row['CustomerCustomFieldsConsultantWhoRegistered'] or NONE
        profiles[cid] = [int(bool(row[k].strip())) if k and k in row else None for k in mapping]
    facts, seen, purchases = [], set(), set()
    missing, excluded = 0, 0
    for row in data['orders']:
        timestamp = utc(row['OrderFirstActionDateTimeUtc'])
        b = row['OrderFirstActionChannelName']
        if row['OrderLineStatusIdsExternalId'] != 'done' or b not in BOUTIQUES:
            excluded += 1
            continue
        cid = row['OrderCustomerIdsMindboxId']
        if not cid:
            missing += 1
            continue
        if not row['OrderIdsMindboxId']:
            raise ValueError('Заказ без ID')
        signature = tuple(sorted(row.items()))
        if signature in seen:
            raise ValueError('Обнаружены одинаковые строки заказов. Удалите дубли в исходной выгрузке.')
        seen.add(signature)
        day = moscow_day(timestamp)
        fact = dict(order=row['OrderIdsMindboxId'], day=day, b=b, customer=cid,
                    amount=number(row['OrderLinePriceOfLine']), qty=number(row['OrderLineQuantity']),
                    seller=row['OrderLineCustomFieldsConsultantPerItem'] or NONE)
        facts.append(fact)
        purchases.add((cid, day))
    first = {}
    for row in data['registrations']:
        cid = row['CustomerActionCustomerIdsMindboxId']
        if not cid:
            raise ValueError('Регистрация без ID клиента')
        timestamp = utc(row['CustomerActionDateTimeUtc'])
        b = row['CustomerActionChannelName']
        b = b if b in BOUTIQUES else UNKNOWN
        old = first.get(cid)
        if old and old['timestamp'] == timestamp and old['b'] != b:
            raise ValueError('Одинаковое время регистрации клиента в разных бутиках')
        if not old or timestamp < old['timestamp']:
            first[cid] = dict(customer=cid, timestamp=timestamp, day=moscow_day(timestamp), b=b,
                              seller=customers.get(cid, NONE), flags=profiles.get(cid, [None] * 6))
    for event in first.values():
        event['no'] = int((event['customer'], event['day']) not in purchases)
        event['filled'] = None if None in event['flags'] else sum(event['flags'])
    for fact in facts:
        reg = first.get(fact['customer'])
        fact['reg_day'] = reg['day'] if reg else None
        fact['category'] = -1 if not reg else (1 if fact['day'] > reg['day'] else 0 if fact['day'] == reg['day'] else -2)
    return dict(facts=facts, events=list(first.values()), missing=missing, excluded=excluded,
                ranges=date_ranges(data), customers=customers, mapping=mapping)


def sales_values(facts):
    amount = sum((x['amount'] for x in facts), Decimal(0))
    qty = sum((x['qty'] for x in facts), Decimal(0))
    orders = len({x['order'] for x in facts})
    return [amount, qty, orders, len({x['customer'] for x in facts}),
            amount / orders if orders else None, qty / orders if orders else None]


def reg_values(events):
    n, no = len(events), sum(x['no'] for x in events)
    filled = None if any(x['filled'] is None for x in events) else sum(x['filled'] for x in events)
    return [n, no, no / n if n else None, n * 6, filled, filled / (n * 6) if n and filled is not None else None]


def period(from_date, to_date):
    try:
        a, z = date.fromisoformat(from_date), date.fromisoformat(to_date)
        if a > z:
            raise ValueError
        return a, z
    except (ValueError, TypeError) as exc:
        raise ValueError('Выберите корректный период') from exc


def report(prepared, from_date, to_date, lang='ru'):
    a, z = period(from_date, to_date)
    facts = [x for x in prepared['facts'] if a <= x['day'] <= z]
    events = [x for x in prepared['events'] if a <= x['day'] <= z]
    returning = [x for x in facts if x['category'] == 1]
    tables = []
    def table(key, title, headers, rows, percent=(), total=False):
        tables.append(dict(key=key, title=tr(title, lang), headers=[tr(h, lang) for h in headers],
                           rows=[[tr(v, lang) for v in row] for row in rows], percent=list(percent), total=total))
    table('sales', 'Продажи по бутикам', SALES_HEAD,
          [[b] + sales_values([x for x in facts if x['b'] == b]) for b in BOUTIQUES]
          + [['Все бутики'] + sales_values(facts)], total=True)
    staff = defaultdict(list)
    for x in facts:
        staff[x['b'], x['seller']].append(x)
    table('staff', 'Детализация по продавцам', STAFF_HEAD,
          [[b, s] + sales_values(v)[:4] for (b, s), v in sorted(staff.items())])
    table('returning', 'Продажи ранее зарегистрированным клиентам', SALES_HEAD,
          [[b] + sales_values([x for x in returning if x['b'] == b]) for b in BOUTIQUES]
          + [['Все бутики'] + sales_values(returning)], total=True)
    table('unknown', 'Продажи клиентам без известной даты регистрации', SALES_HEAD,
          [['Все бутики'] + sales_values([x for x in facts if x['category'] == -1])])
    table('registrations', 'Регистрации и покупки', REG_HEAD,
          [[b] + reg_values([x for x in events if x['b'] == b]) for b in BOUTIQUES + [UNKNOWN]]
          + [['Всего, включая неизвестный бутик'] + reg_values(events)], percent=(3, 6), total=True)
    staff = defaultdict(list)
    for x in events:
        staff[x['b'], x['seller']].append(x)
    table('reg_staff', 'Зарегистрировавшие сотрудники', REG_STAFF_HEAD,
          [[b, s] + reg_values(v) for (b, s), v in sorted(staff.items())], percent=(4, 7))
    table('profiles', 'Заполненность по клиентам',
          ['Бутик', 'ID клиента'] + [f[0] for f in PROFILE_FIELDS] + ['Возможных полей', 'Заполнено полей', 'Заполненность'],
          [[x['b'], x['customer']] + x['flags'] + [6, x['filled'], x['filled'] / 6 if x['filled'] is not None else None]
          for x in sorted(events, key=lambda x: (x['b'], x['customer']))], percent=(10,))
    warnings = []
    def warn(count, ru, en):
        if count:
            warnings.append(f'{count}: ' + (ru if lang == 'ru' else en))
    warn(sum(x['b'] == UNKNOWN for x in events), 'регистраций с неизвестным бутиком; включены в общий итог.', 'registrations with unknown boutique; included in total.')
    warn(sum(x['seller'] == NONE for x in facts), 'позиций без продавца.', 'items with no salesperson.')
    warn(sum(x['seller'] == NONE for x in events), 'регистраций без сотрудника.', 'registrations with no staff member.')
    warn(sum(x['customer'] not in prepared['customers'] for x in events), 'клиентов нет в выгрузке клиентов.', 'registrants absent from customer export.')
    warn(prepared['missing'], 'офлайн-позиций без ID клиента исключено.', 'offline items without customer ID excluded.')
    warn(sum(x['filled'] is None for x in events), 'карточек с неполными исходными данными о полях.', 'profiles with incomplete source columns.')
    if a < date(2026, 10, 9):
        warnings.append(tr('Период включает даты до 09.10.2026: детализация по продавцам неполна.', lang))
    for kind, (lo, hi) in prepared['ranges'].items():
        label = tr('Заказы' if kind == 'orders' else 'Регистрации', lang)
        warnings.append(f'{label}: {lo} — {hi}. ' + ('Полнота зависит от исходной выгрузки.' if lang == 'ru' else 'Completeness depends on the source export.'))
    warnings.append(tr('Для регистраций используется самое раннее событие в загруженном файле. Если более ранние события не выгружены, число новых клиентов может быть завышено.', lang))
    return dict(tables=tables, warnings=warnings, notes=[tr(PROFILE_NOTE, lang), tr(RETURN_NOTE, lang),
                tr('Количество заказов и клиентов по продавцам не суммируется.', lang)],
                period=[from_date, to_date], lang=lang, excluded=prepared['excluded'])


def json_default(value):
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    raise TypeError(type(value).__name__)
