import http.server
import socketserver
import os
import json
import urllib.parse
import secrets
import time
from pathlib import Path
from http import cookies

PORT = 3026
BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent if BASE_DIR.name in ['web_dashboard', 'web_dashboard_direccion'] else BASE_DIR

VALID_USERS = {
    'oscar.ocampo': 'Codiagro2026!',
    'oscar': 'Codiagro2026!',
    'codiagrooscar@gmail.com': 'Codiagro2026!',
    'oscar.ocampo@codiagro.com': 'Codiagro2026!',
    'direccion': 'Codiagro2026!',
    'admin': 'Codiagro2026!'
}

# In-memory session store: { token: { 'user': username, 'created': timestamp } }
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
                    # 7 days expiration
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

        if path == '/api/guardar_acuerdo':
            user = self.get_session_user()
            if not user:
                err_body = json.dumps({'success': False, 'error': 'No autorizado: sesión requerida'}).encode('utf-8')
                self.send_response(401)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(err_body)))
                self.end_headers()
                self.wfile.write(err_body)
                return

            content_length = int(self.headers.get('Content-Length', 0))
            post_body = self.rfile.read(content_length)
            try:
                data = json.loads(post_body.decode('utf-8'))
                from datetime import datetime
                def v_date(s):
                    if not s: return False
                    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d"):
                        try:
                            dt = datetime.strptime(str(s).strip(), fmt)
                            if 1900 <= dt.year <= 2100: return True
                        except ValueError: pass
                    return False

                cli = str(data.get("cliente") or data.get("cliente_id") or "").strip()
                doc = str(data.get("documento_origen") or "").strip()
                fec = str(data.get("fecha") or data.get("fecha_acuerdo") or "").strip()
                vig = str(data.get("vigencia") or data.get("fecha_limite") or "").strip()
                resp = str(data.get("responsable_validacion") or data.get("responsable") or "").strip()

                if not cli or not doc or not fec or not vig or not resp or not v_date(fec) or not v_date(vig):
                    self.send_response(400)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.end_headers()
                    self.wfile.write(json.dumps({'success': False, 'error': 'Datos documentales incompletos o fechas de calendario no válidas.'}).encode('utf-8'))
                    return

                gastos_file = PROJECT_DIR / "Gastos" / "acuerdos_comerciales_registrados.json"
                acuerdos = []
                if gastos_file.exists():
                    try:
                        with open(gastos_file, "r", encoding="utf-8") as gf:
                            acuerdos = json.load(gf)
                    except Exception:
                        acuerdos = []

                # Eliminar metadatos internos del cliente antes de consolidar centralmente
                for k in [k for k in list(data.keys()) if k.startswith("_")]:
                    del data[k]

                import re
                existing_ids = {a.get("id") for a in acuerdos if a.get("id")}
                seq_numbers = []
                for existing_id in existing_ids:
                    m = re.match(r"^ACU-(\d{4})-(\d+)$", str(existing_id))
                    if m:
                        seq_numbers.append(int(m.group(2)))
                next_seq = (max(seq_numbers) + 1) if seq_numbers else 1
                current_year = datetime.now().year

                accion = str(data.get("accion") or "crear").lower()
                req_op_key = str(data.get("clave_operacion") or data.get("idempotency_key") or "").strip()
                req_tema = str(data.get("tema") or data.get("contenido") or data.get("detalle") or "").strip()

                if accion == "editar":
                    aid = str(data.get("id") or "").strip()
                    if not aid or aid not in existing_ids:
                        err_body = json.dumps({'success': False, 'error': f'Acuerdo no encontrado para editar (ID: {aid}).'}).encode('utf-8')
                        self.send_response(404)
                        self.send_header('Content-Type', 'application/json; charset=utf-8')
                        self.send_header('Content-Length', str(len(err_body)))
                        self.end_headers()
                        self.wfile.write(err_body)
                        return
                    data["id"] = aid
                    acuerdos = [data if a.get("id") == aid else a for a in acuerdos]
                else:
                    # Reconocimiento de reintentos:
                    # 1. Por clave única de operación (idempotency key) si está presente
                    # 2. O por coincidencia exacta de cliente, documento, fecha Y contenido/tema
                    existing_match = None
                    for a in acuerdos:
                        a_op_key = str(a.get("clave_operacion") or a.get("idempotency_key") or "").strip()
                        if req_op_key and a_op_key and req_op_key == a_op_key:
                            existing_match = a
                            break

                        a_cli = str(a.get("cliente") or "").strip().upper()
                        a_doc = str(a.get("documento_origen") or "").strip().upper()
                        a_fec = str(a.get("fecha") or "").strip()
                        a_tema = str(a.get("tema") or a.get("contenido") or a.get("detalle") or "").strip()

                        # Si coincide cliente, documento y fecha:
                        if a_cli == cli.upper() and a_doc == doc.upper() and a_fec == fec:
                            # Solo es el mismo acuerdo si el contenido solicitado coincide exactamente
                            if req_tema and a_tema and req_tema == a_tema:
                                existing_match = a
                                break

                    if existing_match:
                        # Reintento del mismo envío: devolver el registro original existente sin duplicar
                        aid = existing_match["id"]
                        data = existing_match
                    else:
                        # Alta nueva (acuerdo nuevo para la misma acta o nueva operación)
                        aid = f"ACU-{current_year}-{next_seq:03d}"
                        while aid in existing_ids:
                            next_seq += 1
                            aid = f"ACU-{current_year}-{next_seq:03d}"
                        data["id"] = aid
                        if not req_op_key:
                            data["clave_operacion"] = f"OP-{current_year}-{aid}"
                        acuerdos.insert(0, data)

                gastos_file.parent.mkdir(parents=True, exist_ok=True)
                with open(gastos_file, "w", encoding="utf-8") as gf:
                    json.dump(acuerdos, gf, ensure_ascii=False, indent=2)

                bi_json = BASE_DIR / "bi_intelligence_data.json"
                if bi_json.exists():
                    try:
                        with open(bi_json, "r", encoding="utf-8") as bf:
                            bi_data = json.load(bf)
                        mr = bi_data.setdefault("mejora_10_modo_reunion", {})
                        mr_list = [a for a in mr.get("acuerdos_registrados", []) if a.get("id") != aid]
                        mr_list.insert(0, data)
                        mr["acuerdos_registrados"] = mr_list
                        mr["total_acuerdos"] = len(mr_list)
                        mr["estado"] = "Operativo"
                        with open(bi_json, "w", encoding="utf-8") as bf:
                            json.dump(bi_data, bf, ensure_ascii=False, indent=2)
                        
                        bi_js = BASE_DIR / "bi_intelligence_data.js"
                        with open(bi_js, "w", encoding="utf-8") as jf:
                            jf.write("window.BI_INTELLIGENCE_DATA = " + json.dumps(bi_data, ensure_ascii=False, indent=2) + ";\n")
                    except Exception as e:
                        print("Aviso al sincronizar bi_intelligence_data:", e)

                resp_body = json.dumps({'success': True, 'acuerdo': data}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(resp_body)))
                self.end_headers()
                self.wfile.write(resp_body)
                return
            except Exception as e:
                err_body = json.dumps({'success': False, 'error': str(e)}).encode('utf-8')
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(err_body)))
                self.end_headers()
                self.wfile.write(err_body)
                return

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

        # Static assets needed for login
        if path in ['/login.html', '/codiagro_logo.png']:
            return super().do_HEAD()

        # Auth check for protected resources
        user = self.get_session_user()
        if not user:
            self.send_response(302)
            self.send_header('Location', '/login.html')
            self.end_headers()
            return

        if path in ['/download/treasury', '/download/tesoreria', '/Prevision_Tesoreria_CashFlow_2026.xlsx']:
            target_file = BASE_DIR / 'Prevision_Tesoreria_CashFlow_2026.xlsx'
            v2_file = BASE_DIR / 'Prevision_Tesoreria_CashFlow_2026_v2.xlsx'
            if v2_file.exists() and (not target_file.exists() or v2_file.stat().st_mtime > target_file.stat().st_mtime):
                target_file = v2_file
            elif not target_file.exists():
                target_file = PROJECT_DIR / 'Prevision_Tesoreria_CashFlow_2026.xlsx'
            if target_file.exists():
                size = target_file.stat().st_size
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename="Prevision_Tesoreria_CashFlow_2026.xlsx"')
                self.send_header('Content-Length', str(size))
                self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                return

        elif path in ['/download/mrp', '/download/fabricacion', '/Orden_Fabricacion_MRP_Codiagro.xlsx']:
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

        # Status check
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

        # Public assets needed for login
        if path in ['/login.html', '/codiagro_logo.png']:
            return super().do_GET()

        # Check authentication for everything else
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

        # Authenticated requests
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

        elif path in ['/download/treasury', '/download/tesoreria', '/Prevision_Tesoreria_CashFlow_2026.xlsx']:
            target_file = BASE_DIR / 'Prevision_Tesoreria_CashFlow_2026.xlsx'
            v2_file = BASE_DIR / 'Prevision_Tesoreria_CashFlow_2026_v2.xlsx'
            if v2_file.exists() and (not target_file.exists() or v2_file.stat().st_mtime > target_file.stat().st_mtime):
                target_file = v2_file
            elif not target_file.exists():
                target_file = PROJECT_DIR / 'Prevision_Tesoreria_CashFlow_2026.xlsx'
            
            with open(target_file, 'rb') as f:
                data = f.read()

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="Prevision_Tesoreria_CashFlow_2026.xlsx"')
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(data)
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
        print(f"=== Servidor Dashboard Dirección General con Autenticación en http://{HOST}:{PORT} ===")
        print("  Autenticación habilitada (Acceso confidencial Dirección General)")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor detenido.")

if __name__ == '__main__':
    run_server()
