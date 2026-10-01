import http.server
import socketserver
import os
import sys
from pathlib import Path

PORT = 3030
BASE_DIR = Path(__file__).resolve().parent

class PresentationHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        super().end_headers()

def run_server():
    os.chdir(BASE_DIR)
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), PresentationHandler) as httpd:
        print(f"================================================================")
        print(f"  CODIAGRO - APP PRESENTACION EJECUTIVA PRESUPUESTO 2027")
        print(f"  Servidor iniciado en: http://localhost:{PORT}")
        print(f"================================================================")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor cerrado.")

if __name__ == '__main__':
    run_server()
