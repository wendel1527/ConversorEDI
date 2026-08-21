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
    from cielo_csv_para_edi import converter as _converter_cielo
    from cielo_csv_para_edi import ler_csv as _ler_csv_cielo
    from cielo_csv_para_edi import _parse_valor
    CIELO_OK = True
except ImportError as e:
    CIELO_OK = False
    IMPORT_ERROR = str(e)

try:
    from rede_csv_para_edi import converter as _converter_rede
    REDE_OK = True
except ImportError:
    REDE_OK = False

try:
    from getnet_csv_para_edi import converter as _converter_getnet
    GETNET_OK = True
except ImportError:
    GETNET_OK = False

CONVERSOR_OK = CIELO_OK  # compatibilidade

def _converter_por_adquirente(adquirente: str, csv_path: str, pasta_saida: str) -> str:
    """Despacha para o conversor correto conforme adquirente."""
    adq = str(adquirente or "cielo").strip().lower()
    if adq == "rede":
        if not REDE_OK:
            raise ImportError("Módulo rede_csv_para_edi não carregado.")
        return _converter_rede(csv_path, pasta_saida)
    elif adq == "getnet":
        if not GETNET_OK:
            raise ImportError("Módulo getnet_csv_para_edi não carregado.")
        # forcar_lq=True: o conciliador só reconhece o indicador 'LQ' (liquidado);
        # 'PF' faz o arquivo inteiro ser rejeitado.
        return _converter_getnet(csv_path, pasta_saida, forcar_lq=True)
    else:
        if not CIELO_OK:
            raise ImportError("Módulo cielo_csv_para_edi não carregado.")
        return _converter_cielo(csv_path, pasta_saida)


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
                "versao": "1.2.0",
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
            ) + sorted(
                str(f) for f in p.glob("*.xlsx")
            ) + sorted(
                str(f) for f in p.glob("*.XLSX")
            )
            self._send_json({"arquivos": list(dict.fromkeys(arquivos)), "pasta": str(p)})
            return

        if path == "/api/preview_csv":
            qs    = parse_qs(parsed.query)
            arq   = qs.get("arquivo", [""])[0]
            adq   = qs.get("adquirente", ["cielo"])[0].lower()
            if not arq or not Path(arq).exists():
                self._send_json({"erro": "Arquivo não encontrado"}, 400)
                return
            try:
                # GetNet: planilha XLSX (aba Detalhado) — tratamento próprio
                if adq == "getnet":
                    if not GETNET_OK:
                        self._send_json({"ok": False, "erro": "Módulo getnet_csv_para_edi não carregado."}, 500)
                        return
                    from getnet_csv_para_edi import ler_xlsx as _ler_getnet, _classificar
                    linhas, meta = _ler_getnet(arq)
                    total    = len(linhas)
                    creditos = sum(1 for r in linhas if _classificar(r) in ("venda", "liquidacao"))
                    debitos  = total - creditos
                    estabs   = sorted(set(r.get("estabelecimento", "").strip() for r in linhas if r.get("estabelecimento")))
                    bandeiras = sorted(set(r.get("bandeira", "").strip() for r in linhas if r.get("bandeira")))
                    di = meta.get("venc_min")
                    df = meta.get("venc_max")
                    self._send_json({
                        "ok": True,
                        "total_linhas": total,
                        "creditos": creditos,
                        "debitos": debitos,
                        "estabelecimentos": estabs,
                        "bandeiras": bandeiras,
                        "data_ini": di.strftime("%d/%m/%Y") if di else "",
                        "data_fim": df.strftime("%d/%m/%Y") if df else "",
                        "meta": {"ignoradas_saldo": meta.get("ignoradas_saldo", 0),
                                 "cpf_cnpj": meta.get("cpf_cnpj", "")},
                    })
                    return

                # Selecionar ler_csv e nomes de colunas conforme adquirente
                if adq == "rede":
                    from rede_csv_para_edi import ler_csv as _ler
                    col_vb   = "valor bruto da parcela original"
                    col_estab = "estabelecimento"
                    col_band  = "bandeira"
                    col_data  = "data do recebimento"
                else:
                    from cielo_csv_para_edi import ler_csv as _ler
                    col_vb   = "Valor bruto"
                    col_estab = "Estabelecimento"
                    col_band  = "Bandeira"
                    col_data  = "Data de pagamento"

                linhas, meta = _ler(arq)
                total     = len(linhas)
                creditos  = sum(1 for r in linhas if _parse_valor(r.get(col_vb, "0")) > 0)
                debitos   = total - creditos
                estabs    = sorted(set(r.get(col_estab, "").strip() for r in linhas if r.get(col_estab)))
                bandeiras = sorted(set(r.get(col_band, "").strip() for r in linhas if r.get(col_band)))
                datas_pag = sorted(set(r.get(col_data, "").strip() for r in linhas if r.get(col_data)))
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

        if parsed.path == "/api/download":
            qs       = parse_qs(parsed.query)
            arquivo  = qs.get("arquivo", [""])[0]
            p        = Path(arquivo)
            output_dir = (BASE_DIR / "output").resolve()
            try:
                p.resolve().relative_to(output_dir)
            except ValueError:
                self._send_json({"erro": "Acesso negado"}, 403)
                return
            if not p.exists() or not p.is_file():
                self._send_json({"erro": "Arquivo não encontrado"}, 404)
                return
            content = p.read_bytes()
            nome    = p.name
            self.send_response(200)
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Disposition", f'attachment; filename="{nome}"')
            self.send_header("Content-Length", str(len(content)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content)
            return

        if parsed.path == "/api/listar_edis":
            p = BASE_DIR / "output"
            if not p.exists():
                self._send_json({"arquivos": []})
                return
            arquivos = sorted(str(f) for f in p.glob("*.txt"))
            self._send_json({"arquivos": arquivos})
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

        length = int(self.headers.get("Content-Length", 0))

        # /api/upload usa multipart/form-data — tratar antes de tentar json.loads
        if path == "/api/upload":
            content_type = self.headers.get("Content-Type", "")
            if "multipart/form-data" not in content_type:
                self._send_json({"erro": "Content-Type deve ser multipart/form-data"}, 400)
                return
            # Extrair boundary
            boundary = None
            for part in content_type.split(";"):
                part = part.strip()
                if part.startswith("boundary="):
                    boundary = part[len("boundary="):].strip().encode()
                    break
            if not boundary:
                self._send_json({"erro": "Boundary não encontrado"}, 400)
                return
            # Ler corpo
            raw = self.rfile.read(length)
            # Parser manual de multipart — extrai o primeiro arquivo .csv
            delimiter = b"--" + boundary
            parts     = raw.split(delimiter)
            destino   = None
            for part in parts:
                if b'filename=' not in part:
                    continue
                # Separar cabeçalhos do conteúdo
                if b"\r\n\r\n" in part:
                    headers_raw, file_content = part.split(b"\r\n\r\n", 1)
                else:
                    continue
                # Remover o \r\n final adicionado pelo multipart
                if file_content.endswith(b"\r\n"):
                    file_content = file_content[:-2]
                # Extrair nome do arquivo
                nome_arquivo = None
                for h in headers_raw.decode("latin-1", errors="replace").split("\r\n"):
                    if "filename=" in h:
                        import re as _re
                        m = _re.search(r'filename="([^"]+)"', h)
                        if m:
                            nome_arquivo = Path(m.group(1)).name
                            break
                if not nome_arquivo:
                    continue
                if not (nome_arquivo.lower().endswith(".csv") or nome_arquivo.lower().endswith(".xlsx")):
                    self._send_json({"erro": "Apenas arquivos .csv ou .xlsx são aceitos"}, 400)
                    return
                input_dir = BASE_DIR / "input"
                input_dir.mkdir(exist_ok=True)
                destino = input_dir / nome_arquivo
                destino.write_bytes(file_content)
                break
            if destino and destino.exists():
                self._send_json({"ok": True, "arquivo": str(destino), "nome": destino.name})
            else:
                self._send_json({"erro": "Nenhum arquivo CSV ou XLSX encontrado no upload"}, 400)
            return

        # Demais rotas usam JSON
        body = self.rfile.read(length)
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
            adquirente  = data.get("adquirente", "cielo")

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
                resultado_path = _converter_por_adquirente(adquirente, arquivo_csv, pasta_saida)
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
