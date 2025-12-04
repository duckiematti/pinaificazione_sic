#!/usr/bin/env python3
"""
Server HTTP minimo per eseguire comandi Python dall'HTML
CROSS-PLATFORM: Funziona su Windows, Mac e Linux
"""
from http.server import HTTPServer, SimpleHTTPRequestHandler
import json
import subprocess
import os
import sys
import platform
from urllib.parse import parse_qs

# Determina il comando Python corretto per la piattaforma
# Su Windows: C:\Python311\python.exe
# Su Mac/Linux: /usr/bin/python3
PYTHON_CMD = sys.executable

class PianificazioneHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/api/list-pdfs':
            try:
                pdf_dir = 'stampe_pdf'
                if os.path.exists(pdf_dir):
                    pdfs = []
                    for file in os.listdir(pdf_dir):
                        if file.endswith('.pdf'):
                            filepath = os.path.join(pdf_dir, file)
                            size = os.path.getsize(filepath)
                            # Converti bytes in KB/MB
                            if size < 1024 * 1024:
                                size_str = f"{size // 1024} KB"
                            else:
                                size_str = f"{size // (1024 * 1024)} MB"
                            pdfs.append({'name': file, 'size': size_str})
                    response = {'pdfs': pdfs}
                else:
                    response = {'pdfs': []}
            except Exception as e:
                response = {'pdfs': [], 'error': str(e)}
            
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(response).encode())
        else:
            # Default GET handler per file statici
            super().do_GET()
    
    def do_POST(self):
        if self.path == '/api/execute':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))
            
            command = data.get('command', '')
            
            # CROSS-PLATFORM: Sostituisci python3 con il comando Python corretto
            # Su Windows: usa python, Su Mac/Linux: usa python3
            # Soluzione: usa sys.executable (il Python che sta eseguendo il server)
            command = command.replace('python3', f'"{PYTHON_CMD}"')
            
            # Su Windows, converte echo Unix in echo Windows
            if platform.system() == 'Windows':
                # echo "1" | python -> echo 1 | python (rimuove virgolette)
                import re
                command = re.sub(r'echo "(\d+)"', r'echo \1', command)
            
            # Esegui comando
            try:
                result = subprocess.run(
                    command,
                    shell=True,
                    capture_output=True,
                    text=True,
                    timeout=30
                )
                
                response = {
                    'success': True,
                    'stdout': result.stdout,
                    'stderr': result.stderr,
                    'returncode': result.returncode
                }
            except subprocess.TimeoutExpired:
                response = {
                    'success': False,
                    'error': 'Comando timeout (>30 secondi)'
                }
            except Exception as e:
                response = {
                    'success': False,
                    'error': str(e)
                }
            
            # Invia risposta
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(response).encode())
        
        elif self.path == '/api/check-file':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))
            
            filepath = data.get('filepath', '')
            
            try:
                exists = os.path.exists(filepath)
                response = {'exists': exists}
            except Exception as e:
                response = {'exists': False, 'error': str(e)}
            
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(response).encode())
        
        elif self.path == '/api/download-pdf':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))
            
            filename = data.get('filename', '')
            filepath = os.path.join('stampe_pdf', filename)
            
            try:
                if os.path.exists(filepath) and filename.endswith('.pdf'):
                    with open(filepath, 'rb') as f:
                        pdf_data = f.read()
                    
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/pdf')
                    self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(pdf_data)
                else:
                    self.send_response(404)
                    self.send_header('Content-Type', 'application/json')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(json.dumps({'error': 'File non trovato'}).encode())
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode())
        
        elif self.path == '/api/open-file':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))
            
            filepath = data.get('filepath', '')
            
            try:
                subprocess.run(['open', filepath], check=True)
                response = {'success': True}
            except Exception as e:
                response = {'success': False, 'error': str(e)}
            
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(response).encode())
    
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

if __name__ == '__main__':
    port = 8765
    server = HTTPServer(('localhost', port), PianificazioneHandler)
    print(f"\n{'='*70}")
    print("🌐 SERVER PIANIFICAZIONE AVVIATO")
    print(f"{'='*70}")
    print(f"\n📋 Apri nel browser: http://localhost:{port}/pianificazione.html")
    print(f"\n⏹️  Premi CTRL+C per fermare")
    print(f"{'='*70}\n")
    
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n\n✅ Server fermato")
