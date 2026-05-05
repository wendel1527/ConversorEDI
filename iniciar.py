#!/usr/bin/env python3
"""
Cielo EDI Converter — Iniciador
================================
Duplo clique para iniciar a aplicação desktop.

Requisitos:
  • Python 3.9+  (sem instalações adicionais necessárias)
  • Os arquivos abaixo devem estar na MESMA pasta que este script:
      - servidor.py
      - cielo_csv_para_edi.py
      - app.html

Uso via terminal (opcional):
  python iniciar.py
  python iniciar.py --sem-browser    # não abre o browser automaticamente
"""

import sys
import os

# Garantir que estamos no diretório do script
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Verificar versão do Python
if sys.version_info < (3, 9):
    print("=" * 55)
    print("  ERRO: Python 3.9 ou superior é necessário.")
    print(f"  Versão atual: {sys.version}")
    print("  Baixe em: https://www.python.org/downloads/")
    print("=" * 55)
    input("Pressione Enter para fechar...")
    sys.exit(1)

# Verificar arquivos necessários
import pathlib
faltando = []
for arq in ["servidor.py", "cielo_csv_para_edi.py", "app.html"]:
    if not pathlib.Path(arq).exists():
        faltando.append(arq)

if faltando:
    print("=" * 55)
    print("  ERRO: Arquivo(s) necessário(s) não encontrado(s):")
    for f in faltando:
        print(f"    ✗  {f}")
    print()
    print("  Certifique-se de que todos os arquivos estão")
    print("  na mesma pasta que iniciar.py")
    print("=" * 55)
    input("Pressione Enter para fechar...")
    sys.exit(1)

# Iniciar servidor
abrir_browser = "--sem-browser" not in sys.argv

from servidor import iniciar_servidor
iniciar_servidor(abrir_browser=abrir_browser)
