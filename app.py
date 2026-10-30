"""Local Python report application. Run: python3 app.py"""
import argparse
import base64
import csv
import io
import json
import secrets
import threading
import webbrowser
from collections import OrderedDict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit

from engine import REQUIRED, PROFILE_FIELDS, parse_csv, prepare, report, json_default, tr
from excel import export_xlsx

ROOT = Path(__file__).resolve().parent
MAX_REQUEST = 128 * 1024 * 1024


def csv_export(result):
    stream = io.StringIO(newline='')
    writer = csv.writer(stream, delimiter=';')
    def row(values):
        # Keep literal CSV text from being interpreted as a spreadsheet formula.
        writer.writerow(["'" + v if isinstance(v, str) and v.startswith(('=', '+', '-', '@', '\t', '\r')) else v for v in values])
    row(['Emiliano Zapata', *result['period'], 'Moscow'])
    for table in result['tables']:
        row([])
        row([table['title']])
        row(table['headers'])
        for values in table['rows']:
            row([tr('Нет данных', result['lang']) if v is None else v for v in values])
    for note in result['warnings'] + result['notes']:
        row([note])
    return stream.getvalue().encode('utf-8-sig')


class ReportServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, address):
        super().__init__(address, Handler)
        self.token = secrets.token_urlsafe(32)
        self.datasets = OrderedDict()
        self.lock = threading.Lock()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *_):
        pass  # Never log uploaded customer data.

    def send(self, body, content_type='application/json; charset=utf-8', status=200, filename=None):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'")
        if filename:
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, value, status=200):
        self.send(json.dumps(value, ensure_ascii=False, default=json_default).encode('utf-8'), status=status)

    def valid_host(self):
        return self.headers.get('Host') == f'127.0.0.1:{self.server.server_port}'

    def do_GET(self):
        if not self.valid_host():
            self.send_json({'error': 'Invalid host'}, 403)
            return
        path = urlsplit(self.path).path
        if path == '/':
            content = (ROOT / 'static/index.html').read_text().replace('__TOKEN__', self.server.token)
            self.send(content.encode(), 'text/html; charset=utf-8')
        elif path in ['/app.js', '/style.css']:
            self.send((ROOT / 'static' / path[1:]).read_bytes(), 'text/javascript; charset=utf-8' if path.endswith('.js') else 'text/css; charset=utf-8')
        else:
            self.send_json({'error': 'Not found'}, 404)

    def do_POST(self):
        lang = 'ru'
        if not self.valid_host() or not secrets.compare_digest(self.headers.get('X-Report-Token', ''), self.server.token):
            self.send_json({'error': 'Reload this application / Обновите страницу приложения'}, 403)
            return
        origin = self.headers.get('Origin')
        if origin and origin != f'http://127.0.0.1:{self.server.server_port}':
            self.send_json({'error': 'Invalid origin'}, 403)
            return
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= MAX_REQUEST:
                self.send_json({'error': 'Upload limit: 128 MB / Лимит запроса: 128 МБ'}, 413)
                return
            body = json.loads(self.rfile.read(length))
            lang = body.get('lang', 'ru')
            if lang not in ('ru', 'en'):
                raise ValueError('Invalid language')
            if self.path in ('/api/load', '/api/demo'):
                data = {}
                for kind in REQUIRED:
                    content = (ROOT / 'demo' / f'{kind}.csv').read_bytes() if self.path == '/api/demo' else base64.b64decode(body['files'][kind], validate=True)
                    try:
                        data[kind] = parse_csv(content, kind)
                    except ValueError as exc:
                        raise ValueError(kind + ': ' + tr(str(exc), lang)) from exc
                prepared = prepare(data)
                key = secrets.token_urlsafe(18)
                with self.server.lock:
                    self.server.datasets[key] = data
                    while len(self.server.datasets) > 8:
                        self.server.datasets.popitem(last=False)
                ranges = prepared['ranges']
                lo = max(v[0] for v in ranges.values())
                hi = min(v[1] for v in ranges.values())
                self.send_json(dict(id=key, counts={k: len(v) for k, v in data.items()},
                                    headers=list(data['customers'][0]), mapping=prepared['mapping'],
                                    from_date=lo.isoformat() if lo <= hi else '', to_date=hi.isoformat() if lo <= hi else '',
                                    ranges=ranges))
                return
            with self.server.lock:
                data = self.server.datasets.get(body.get('id'))
            if data is None:
                raise ValueError('Загрузите три CSV заново / Load the three CSV files again')
            prepared = prepare(data, body.get('mapping'))
            start, end = body['from_date'], body['to_date']
            result = report(prepared, start, end, lang)
            filename = f'Emiliano-Zapata-{start}-{end}'
            if self.path == '/api/report':
                self.send_json(result)
            elif self.path == '/api/excel':
                self.send(export_xlsx(prepared, start, end, lang), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=filename+'.xlsx')
            elif self.path == '/api/csv':
                self.send(csv_export(result), 'text/csv; charset=utf-8', filename=filename+'.csv')
            else:
                self.send_json({'error': 'Not found'}, 404)
        except (ValueError, KeyError, TypeError, UnicodeError) as exc:
            self.send_json({'error': tr(str(exc), lang)}, 400)
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception:
            self.send_json({'error': 'Не удалось сформировать отчёт / Unable to build report'}, 500)


def main():
    parser = argparse.ArgumentParser(description='Emiliano Zapata — Python reports')
    parser.add_argument('--port', type=int, default=0, help='Local port; 0 selects an available one')
    parser.add_argument('--no-browser', action='store_true')
    args = parser.parse_args()
    server = ReportServer(('127.0.0.1', args.port))
    url = f'http://127.0.0.1:{server.server_port}'
    print(f'Emiliano Zapata: {url}\nДля остановки / To stop: Ctrl+C', flush=True)
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nПриложение остановлено / Application stopped')
    finally:
        server.server_close()
        server.datasets.clear()


if __name__ == '__main__':
    main()
