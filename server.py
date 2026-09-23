import http.server
import socketserver
import os
import json
import urllib.parse
from pathlib import Path

PORT = 3025
BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent if BASE_DIR.name == 'web_dashboard' else BASE_DIR

class DashboardHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/api/data':
            json_file = PROJECT_DIR / 'dashboard_data.json'
            if not json_file.exists():
                # Generate if not exists
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
            self.end_headers()
            self.wfile.write(content)
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
            self.end_headers()
            self.wfile.write(data)
            return

        elif path == '/' or path == '':
            self.path = '/index.html'

        return super().do_GET()

def run_server():
    os.chdir(PROJECT_DIR)
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), DashboardHandler) as httpd:
        print(f"=== Servidor Dashboard Presupuesto Septiembre 2026 iniciado en http://localhost:{PORT} ===")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor detenido.")

if __name__ == '__main__':
    run_server()
