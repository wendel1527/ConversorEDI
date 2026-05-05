"""
Servidor local da aplicação Cielo EDI Converter.
Recebe requisições do app.html e executa o conversor Python.
"""

import http.server
import json
import os
import sys
import io
import threading
import webbrowser
import traceback
from pathlib import Path
from urllib.parse import urlparse, parse_qs

# Garantir que o diretório do script está no PATH
BASE_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(BASE_DIR))

try:
    from cielo_csv_para_edi import converter, ler_csv, _parse_valor
    CONVERSOR_OK = True
except ImportError as e:
    CONVERSOR_OK = False
    IMPORT_ERROR = str(e)


PORT = 8765


class CieloHandler(http.server.BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        pass  # silenciar logs do servidor no terminal

    def _send_json(self, data: dict, status: int = 200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path: Path):
        ext  = path.suffix.lower()
        mime = {
            ".html": "text/html; charset=utf-8",
            ".css" : "text/css",
            ".js"  : "application/javascript",
            ".ico" : "image/x-icon",
        }.get(ext, "application/octet-stream")
        content = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path   = parsed.path

        if path in ("/", "/index.html"):
            html_file = BASE_DIR / "app.html"
            if html_file.exists():
                self._send_file(html_file)
            else:
                self._send_json({"erro": "app.html não encontrado"}, 500)
            return

        if path == "/api/status":
            self._send_json({
                "ok": CONVERSOR_OK,
                "erro": "" if CONVERSOR_OK else IMPORT_ERROR,
                "versao": "1.0.0",
            })
            return

        if path == "/api/listar_csvs":
            qs     = parse_qs(parsed.query)
            pasta  = qs.get("pasta", [""])[0]
            if not pasta:
                pasta = str(BASE_DIR / "input")
            p = Path(pasta)
            if not p.exists():
                self._send_json({"arquivos": [], "pasta": str(p)})
                return
            arquivos = sorted(
                str(f) for f in p.glob("*.csv")
            ) + sorted(
                str(f) for f in p.glob("*.CSV")
            )
            self._send_json({"arquivos": list(dict.fromkeys(arquivos)), "pasta": str(p)})
            return

        if path == "/api/preview_csv":
            qs    = parse_qs(parsed.query)
            arq   = qs.get("arquivo", [""])[0]
            if not arq or not Path(arq).exists():
                self._send_json({"erro": "Arquivo não encontrado"}, 400)
                return
            try:
                linhas, meta = ler_csv(arq)
                total = len(linhas)
                creditos  = sum(1 for r in linhas if _parse_valor(r.get("Valor bruto","0")) > 0)
                debitos   = total - creditos
                estabs    = sorted(set(r.get("Estabelecimento","").strip() for r in linhas if r.get("Estabelecimento")))
                bandeiras = sorted(set(r.get("Bandeira","").strip() for r in linhas if r.get("Bandeira")))
                datas_pag = sorted(set(r.get("Data de pagamento","").strip() for r in linhas if r.get("Data de pagamento")))
                self._send_json({
                    "ok": True,
                    "total_linhas": total,
                    "creditos": creditos,
                    "debitos": debitos,
                    "estabelecimentos": estabs,
                    "bandeiras": bandeiras,
                    "data_ini": datas_pag[0]  if datas_pag else "",
                    "data_fim": datas_pag[-1] if datas_pag else "",
                    "meta": meta,
                })
            except Exception as e:
                self._send_json({"ok": False, "erro": str(e)}, 400)
            return

        # Servir arquivos estáticos da pasta do app
        file_path = BASE_DIR / path.lstrip("/")
        if file_path.exists() and file_path.is_file():
            self._send_file(file_path)
            return

        self._send_json({"erro": "Rota não encontrada"}, 404)

    def do_POST(self):
        parsed = urlparse(self.path)
        path   = parsed.path

        length  = int(self.headers.get("Content-Length", 0))
        body    = self.rfile.read(length)
        try:
            data = json.loads(body) if body else {}
        except json.JSONDecodeError:
            self._send_json({"erro": "JSON inválido"}, 400)
            return

        if path == "/api/converter":
            if not CONVERSOR_OK:
                self._send_json({"ok": False, "erro": f"Módulo conversor não carregado: {IMPORT_ERROR}"}, 500)
                return

            arquivo_csv = data.get("arquivo_csv", "")
            pasta_saida = data.get("pasta_saida", "")

            if not arquivo_csv:
                self._send_json({"ok": False, "erro": "Caminho do CSV não informado."}, 400)
                return
            if not Path(arquivo_csv).exists():
                self._send_json({"ok": False, "erro": f"Arquivo não encontrado: {arquivo_csv}"}, 400)
                return
            if not pasta_saida:
                pasta_saida = str(BASE_DIR / "output")

            # Capturar o stdout do conversor para enviar ao frontend
            old_stdout = sys.stdout
            sys.stdout = captura = io.StringIO()
            resultado_path = None
            erro_fatal = None

            try:
                resultado_path = converter(arquivo_csv, pasta_saida)
            except Exception as e:
                erro_fatal = traceback.format_exc()
            finally:
                sys.stdout = old_stdout

            log = captura.getvalue()

            if erro_fatal:
                self._send_json({
                    "ok": False,
                    "erro": erro_fatal,
                    "log": log,
                })
            else:
                self._send_json({
                    "ok": True,
                    "arquivo_gerado": str(resultado_path),
                    "log": log,
                })
            return

        if path == "/api/abrir_pasta":
            pasta = data.get("pasta", "")
            if pasta and Path(pasta).exists():
                import subprocess, platform
                sistema = platform.system()
                try:
                    if sistema == "Windows":
                        os.startfile(pasta)
                    elif sistema == "Darwin":
                        subprocess.Popen(["open", pasta])
                    else:
                        subprocess.Popen(["xdg-open", pasta])
                    self._send_json({"ok": True})
                except Exception as e:
                    self._send_json({"ok": False, "erro": str(e)})
            else:
                self._send_json({"ok": False, "erro": "Pasta não encontrada"})
            return

        self._send_json({"erro": "Rota não encontrada"}, 404)


def iniciar_servidor(abrir_browser: bool = True):
    (BASE_DIR / "input").mkdir(exist_ok=True)
    (BASE_DIR / "output").mkdir(exist_ok=True)

    servidor = http.server.ThreadingHTTPServer(("127.0.0.1", PORT), CieloHandler)
    url = f"http://127.0.0.1:{PORT}"

    print(f"\n{'='*55}")
    print(f"  Cielo EDI Converter — Servidor iniciado")
    print(f"  Acesse: {url}")
    print(f"  Para encerrar: pressione Ctrl+C")
    print(f"{'='*55}\n")

    if abrir_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()

    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\n  Servidor encerrado.")
        servidor.shutdown()


if __name__ == "__main__":
    iniciar_servidor(abrir_browser=True)
