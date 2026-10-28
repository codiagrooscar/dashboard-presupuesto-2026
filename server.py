import http.server
import socketserver
import os
import json
import urllib.parse
import secrets
import time
from pathlib import Path
from http import cookies

PORT = 3025
BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent if BASE_DIR.name == 'web_dashboard' else BASE_DIR

VALID_USERS = {
    'oscar.ocampo': 'Codiagro2026!',
    'oscar': 'Codiagro2026!',
    'direccion': 'Codiagro2026!',
    'comercial': 'Codiagro2026!',
    'admin': 'Codiagro2026!'
}

ACTIVE_SESSIONS = {}

class DashboardHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def get_session_user(self):
        cookie_header = self.headers.get('Cookie')
        if not cookie_header:
            return None
        c = cookies.SimpleCookie()
        try:
            c.load(cookie_header)
            if 'codiagro_auth_token' in c:
                token = c['codiagro_auth_token'].value
                session = ACTIVE_SESSIONS.get(token)
                if session:
                    if time.time() - session['created'] < 7 * 86400:
                        return session['user']
                    else:
                        ACTIVE_SESSIONS.pop(token, None)
        except Exception:
            pass
        return None

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/api/login':
            content_length = int(self.headers.get('Content-Length', 0))
            post_body = self.rfile.read(content_length)
            try:
                data = json.loads(post_body.decode('utf-8'))
                username = str(data.get('username', '')).strip().lower()
                password = str(data.get('password', ''))

                if username in VALID_USERS and VALID_USERS[username] == password:
                    token = secrets.token_hex(24)
                    ACTIVE_SESSIONS[token] = {
                        'user': username,
                        'created': time.time()
                    }

                    resp_body = json.dumps({
                        'success': True,
                        'user': username,
                        'token': token
                    }).encode('utf-8')

                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Set-Cookie', f'codiagro_auth_token={token}; Path=/; Max-Age=604800; SameSite=Lax')
                    self.send_header('Content-Length', str(len(resp_body)))
                    self.end_headers()
                    self.wfile.write(resp_body)
                    return
                else:
                    err_body = json.dumps({
                        'success': False,
                        'error': 'Usuario o contraseña incorrectos'
                    }).encode('utf-8')
                    self.send_response(401)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Content-Length', str(len(err_body)))
                    self.end_headers()
                    self.wfile.write(err_body)
                    return
            except Exception as e:
                err_body = json.dumps({'success': False, 'error': str(e)}).encode('utf-8')
                self.send_response(400)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(err_body)))
                self.end_headers()
                self.wfile.write(err_body)
                return

        elif path == '/api/logout':
            cookie_header = self.headers.get('Cookie')
            if cookie_header:
                c = cookies.SimpleCookie()
                try:
                    c.load(cookie_header)
                    if 'codiagro_auth_token' in c:
                        token = c['codiagro_auth_token'].value
                        ACTIVE_SESSIONS.pop(token, None)
                except Exception:
                    pass
            resp_body = json.dumps({'success': True}).encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Set-Cookie', 'codiagro_auth_token=; Path=/; Max-Age=0; SameSite=Lax')
            self.send_header('Content-Length', str(len(resp_body)))
            self.end_headers()
            self.wfile.write(resp_body)
            return

        self.send_error(404, "Endpoint no encontrado")

    def do_HEAD(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path in ['/login.html', '/codiagro_logo.png']:
            return super().do_HEAD()

        user = self.get_session_user()
        if not user:
            self.send_response(302)
            self.send_header('Location', '/login.html')
            self.end_headers()
            return

        if path in ['/download/mrp', '/download/fabricacion', '/Orden_Fabricacion_MRP_Codiagro.xlsx']:
            target_file = BASE_DIR / 'Orden_Fabricacion_MRP_Codiagro.xlsx'
            if not target_file.exists():
                target_file = PROJECT_DIR / 'Orden_Fabricacion_MRP_Codiagro.xlsx'
            if target_file.exists():
                size = target_file.stat().st_size
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename="Orden_Fabricacion_MRP_Codiagro.xlsx"')
                self.send_header('Content-Length', str(size))
                self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                return

        return super().do_HEAD()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/api/auth/status':
            user = self.get_session_user()
            resp = json.dumps({
                'authenticated': bool(user),
                'user': user
            }).encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Length', str(len(resp)))
            self.end_headers()
            self.wfile.write(resp)
            return

        elif path == '/api/logout':
            cookie_header = self.headers.get('Cookie')
            if cookie_header:
                c = cookies.SimpleCookie()
                try:
                    c.load(cookie_header)
                    if 'codiagro_auth_token' in c:
                        token = c['codiagro_auth_token'].value
                        ACTIVE_SESSIONS.pop(token, None)
                except Exception:
                    pass
            self.send_response(302)
            self.send_header('Set-Cookie', 'codiagro_auth_token=; Path=/; Max-Age=0; SameSite=Lax')
            self.send_header('Location', '/login.html')
            self.end_headers()
            return

        if path in ['/login.html', '/codiagro_logo.png']:
            return super().do_GET()

        user = self.get_session_user()
        if not user:
            if path in ['/', '/index.html', '']:
                self.send_response(302)
                self.send_header('Location', '/login.html')
                self.end_headers()
                return
            elif path.endswith('.json') or path.endswith('.js') or path.startswith('/api/') or path.startswith('/download/'):
                self.send_response(401)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(b'{"error": "Autenticacion requerida", "login_url": "/login.html"}')
                return
            else:
                self.send_response(302)
                self.send_header('Location', '/login.html')
                self.end_headers()
                return

        if path == '/api/data':
            json_file = BASE_DIR / 'dashboard_data.json'
            if not json_file.exists():
                json_file = PROJECT_DIR / 'dashboard_data.json'
            if not json_file.exists():
                import sys
                scripts_dir = str(PROJECT_DIR / 'scripts')
                if scripts_dir not in sys.path:
                    sys.path.insert(0, scripts_dir)
                from build_dataset import build_dataset
                data = build_dataset()
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False)
            
            with open(json_file, 'r', encoding='utf-8') as f:
                content = f.read().encode('utf-8')

            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Length', str(len(content)))
            self.send_header('Cache-Control', 'no-cache')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(content)
            return

        elif path in ['/download/mrp', '/download/fabricacion', '/Orden_Fabricacion_MRP_Codiagro.xlsx']:
            target_file = BASE_DIR / 'Orden_Fabricacion_MRP_Codiagro.xlsx'
            if not target_file.exists():
                target_file = PROJECT_DIR / 'Orden_Fabricacion_MRP_Codiagro.xlsx'
            if not target_file.exists():
                self.send_error(404, "Fichero Orden de Fabricacion no encontrado")
                return

            with open(target_file, 'rb') as f:
                data = f.read()

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="Orden_Fabricacion_MRP_Codiagro.xlsx"')
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(data)
            return

        elif path == '/download/excel':
            excel_file_unidades = PROJECT_DIR / 'Seguimiento_Presupuesto_Sep_2026_Unidades.xlsx'
            excel_file_std = PROJECT_DIR / 'Seguimiento_Presupuesto_Sep_2026.xlsx'
            target_file = excel_file_unidades if excel_file_unidades.exists() else excel_file_std
            if not target_file.exists():
                self.send_error(404, "Fichero Excel no encontrado")
                return

            with open(target_file, 'rb') as f:
                data = f.read()

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="Seguimiento_Presupuesto_Sep_2026.xlsx"')
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(data)
            return

        elif path == '/' or path == '':
            self.path = '/index.html'

        return super().do_GET()

def run_server():
    os.chdir(PROJECT_DIR)
    socketserver.TCPServer.allow_reuse_address = True
    HOST = "127.0.0.1"
    with socketserver.TCPServer((HOST, PORT), DashboardHandler) as httpd:
        print(f"=== Servidor Dashboard Comercial con Autenticación en http://{HOST}:{PORT} ===")
        print(f"  Usuario: oscar.ocampo (o oscar / comercial)")
        print(f"  Contraseña: Codiagro2026!")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor detenido.")

if __name__ == '__main__':
    run_server()
