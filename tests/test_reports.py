"""Business-rule and HTTP tests. All fixtures are synthetic."""
import base64
import copy
import csv
import io
import json
import sys
import threading
import unittest
import urllib.error
import urllib.request
import zipfile
from datetime import date
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from engine import REQUIRED, PROFILE_FIELDS, parse_csv, prepare, report, number, utc, moscow_day
from excel import export_xlsx, make_sheets, NS
from app import ReportServer, csv_export


def fixture():
    def customer(cid):
        row = dict(CustomerIdsMindboxId=cid, CustomerCustomFieldsConsultantWhoRegistered='Demo advisor')
        row.update({f[1][0]: 'Demo' for f in PROFILE_FIELDS})
        row['CustomerEmail'] = ''
        return row
    def order(oid, cid, timestamp, amount='100', qty='1', boutique='Авиапарк', seller='Demo advisor', status='done'):
        return dict(zip(REQUIRED['orders'], [oid, timestamp, boutique, cid, qty, amount, status, seller]))
    def reg(cid, timestamp, boutique='Авиапарк'):
        return dict(zip(REQUIRED['registrations'], [cid, timestamp, boutique]))
    return dict(customers=[customer('001'), customer('002'), customer('003')],
                orders=[order('001', '001', '2026-10-08 21:00:00', '100,50', '2'),
                        order('001', '001', '2026-10-08 21:00:00', '50', '1', seller='Demo assistant'),
                        order('002', '001', '2026-10-10 10:00:00', '200', boutique='ГУМ'),
                        order('003', '002', '2026-10-09 08:00:00', '300'),
                        order('004', '003', '2026-10-09 08:00:00', '400', status='returned'),
                        order('005', 'missing', '2026-10-09 08:00:00', '25'),
                        order('006', '', '2026-10-09 08:00:00', '999'),
                        order('007', '001', '2026-10-10 21:00:00', '10')],
                registrations=[reg('001', '2026-10-09 12:00:00'), reg('002', '2026-10-09 12:00:00'),
                               reg('003', '2026-10-10 08:00:00', 'Интеграция касс'),
                               reg('001', '2026-10-10 12:00:00', 'ГУМ')])


def csv_bytes(rows, delimiter=';', encoding='utf-8-sig'):
    stream = io.StringIO(newline='')
    writer = csv.DictWriter(stream, fieldnames=list(rows[0]), delimiter=delimiter)
    writer.writeheader()
    writer.writerows(rows)
    return stream.getvalue().encode(encoding)


class BusinessTests(unittest.TestCase):
    def setUp(self):
        self.data = fixture()

    def result(self, data=None):
        return report(prepare(data or self.data), '2026-10-09', '2026-10-10')

    def test_sales_and_distinct_network_totals(self):
        result = self.result()
        tables = {t['key']: t for t in result['tables']}
        self.assertEqual(tables['sales']['rows'][-1][1:5], [Decimal('675.50'), Decimal(6), 4, 3])
        self.assertEqual(tables['returning']['rows'][-1][1:5], [Decimal(200), Decimal(1), 1, 1])
        self.assertEqual(tables['unknown']['rows'][0][1], 25)

    def test_first_registration_and_purchase_before_registration(self):
        tables = {t['key']: t for t in self.result()['tables']}
        self.assertEqual(tables['registrations']['rows'][-1][1:3], [3, 1])
        self.assertEqual(tables['registrations']['rows'][0][1:3], [2, 0])
        self.assertEqual(tables['registrations']['rows'][-1][4:7], [18, 15, 15/18])

    def test_moscow_boundaries(self):
        self.assertEqual(moscow_day(utc('2026-10-08 21:00:00')), date(2026, 10, 9))
        self.assertEqual(moscow_day(utc('2026-10-10 21:00:00')), date(2026, 10, 11))

    def test_missing_column_is_not_zero(self):
        for row in self.data['customers']:
            del row['CustomerEmail']
        tables = {t['key']: t for t in self.result()['tables']}
        self.assertEqual(tables['registrations']['rows'][-1][4:7], [18, None, None])

    def test_missing_customer_profile(self):
        self.data['customers'].pop()
        events = prepare(self.data)['events']
        self.assertIsNone(next(x for x in events if x['customer'] == '003')['filled'])

    def test_mapping_override(self):
        self.data['customers'][0]['Other email'] = 'test@example.invalid'
        for row in self.data['customers'][1:]:
            row['Other email'] = ''
        mapping = [f[1][0] for f in PROFILE_FIELDS]
        mapping[-1] = 'Other email'
        self.assertEqual(prepare(self.data, mapping)['events'][0]['filled'], 6)
        with self.assertRaises(ValueError):
            prepare(self.data, [mapping[0]] * 6)

    def test_duplicates_and_ambiguous_registration(self):
        data = copy.deepcopy(self.data)
        data['orders'].append(data['orders'][0].copy())
        with self.assertRaises(ValueError):
            prepare(data)
        data = copy.deepcopy(self.data)
        data['customers'].append(data['customers'][0].copy())
        with self.assertRaises(ValueError):
            prepare(data)
        data = copy.deepcopy(self.data)
        data['registrations'].append({**data['registrations'][0], 'CustomerActionChannelName': 'ГУМ'})
        with self.assertRaises(ValueError):
            prepare(data)

    def test_csv_encodings_and_quoting(self):
        self.data['customers'][0]['CustomerFirstName'] = 'Анна; "тест"\nновая строка'
        for encoding in ['utf-8-sig', 'cp1251']:
            for delimiter in [';', ',', '\t']:
                self.assertEqual(parse_csv(csv_bytes(self.data['customers'], delimiter, encoding), 'customers'), self.data['customers'])
        for raw in [b'a,a\n1,2', b'abc\n1', csv_bytes(self.data['customers']) + b'wrong;count\n']:
            with self.assertRaises(ValueError):
                parse_csv(raw, 'customers')

    def test_invalid_numbers_dates_and_periods(self):
        for value in ['', '-1', 'NaN', 'Infinity', 'text']:
            with self.assertRaises(ValueError):
                number(value)
        self.assertEqual(number('1\u00a0234,50'), Decimal('1234.50'))
        with self.assertRaises(ValueError):
            utc('2026-02-30 10:00:00')
        with self.assertRaises(ValueError):
            report(prepare(self.data), '2026-10-11', '2026-10-09')

    def test_empty_period(self):
        tables = report(prepare(self.data), '2026-12-01', '2026-12-02')['tables']
        self.assertEqual(tables[0]['rows'][-1][1:5], [0, 0, 0, 0])

    def test_xlsx_structure_cached_values_and_literal_text(self):
        self.data['orders'][0]['OrderLineCustomFieldsConsultantPerItem'] = '=1+1'
        p = prepare(self.data)
        for lang in ['ru', 'en']:
            raw = export_xlsx(p, '2026-10-09', '2026-10-10', lang)
            with zipfile.ZipFile(io.BytesIO(raw)) as archive:
                self.assertEqual(len([x for x in archive.namelist() if x.startswith('xl/tables/')]), 9)
                for name in archive.namelist():
                    ET.fromstring(archive.read(name))
                sheet = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
                ns = {'s': NS}
                self.assertEqual(sheet.find('.//s:c[@r="C15"]/s:v', ns).text, '675.50')
                self.assertIsNotNone(sheet.find('.//s:c[@r="C15"]/s:f', ns))
                source = ET.fromstring(archive.read('xl/worksheets/sheet3.xml'))
                cell = source.find('.//s:c[@r="D5"]', ns)
                self.assertEqual(cell.attrib['t'], 'inlineStr')
                self.assertEqual(cell.find('s:is/s:t', ns).text, '=1+1')
                self.assertIsNone(cell.find('s:f', ns))

    def test_csv_formula_injection(self):
        self.data['orders'][0]['OrderLineCustomFieldsConsultantPerItem'] = '=1+1'
        self.assertIn("'=1+1", csv_export(self.result()).decode('utf-8-sig'))


class HttpTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = ReportServer(('127.0.0.1', 0))
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.url = f'http://127.0.0.1:{cls.server.server_port}'

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join()

    def request(self, path, body, token=True):
        headers = {'Content-Type': 'application/json'}
        if token:
            headers['X-Report-Token'] = self.server.token
        return urllib.request.urlopen(urllib.request.Request(self.url + path, json.dumps(body).encode(), headers))

    def test_upload_build_export_flow(self):
        body = {'files': {k: base64.b64encode(csv_bytes(v)).decode() for k, v in fixture().items()}}
        loaded = json.load(self.request('/api/load', body))
        self.assertEqual(loaded['counts']['customers'], 3)
        payload = dict(id=loaded['id'], from_date='2026-10-09', to_date='2026-10-10', mapping=loaded['mapping'], lang='en')
        result = json.load(self.request('/api/report', payload))
        self.assertEqual(result['tables'][0]['rows'][-1][1], 675.5)
        self.assertEqual(result['tables'][0]['headers'][0], 'Boutique')
        self.assertTrue(self.request('/api/excel', payload).read().startswith(b'PK'))
        self.assertIn(b'675.50', self.request('/api/csv', payload).read())

    def test_no_token_and_foreign_origin_rejected(self):
        with self.assertRaises(urllib.error.HTTPError) as ctx:
            self.request('/api/load', {}, token=False)
        self.assertEqual(ctx.exception.code, 403)
        req = urllib.request.Request(self.url+'/api/load', b'{}', {'X-Report-Token': self.server.token, 'Origin': 'https://example.invalid'})
        with self.assertRaises(urllib.error.HTTPError) as ctx:
            urllib.request.urlopen(req)
        self.assertEqual(ctx.exception.code, 403)

    def test_invalid_file_error(self):
        with self.assertRaises(urllib.error.HTTPError) as ctx:
            self.request('/api/load', {'files': {'orders': base64.b64encode(b'bad').decode()}})
        self.assertEqual(ctx.exception.code, 400)


if __name__ == '__main__':
    unittest.main()
