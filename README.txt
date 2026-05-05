============================================================
  Cielo EDI Converter — Guia de Uso
  Versão 1.0.0 · Manual Cielo v15.15 (fev/2026)
============================================================

O QUE É
-------
Converte o CSV de Recebíveis exportado do portal Cielo para
o formato EDI de posição fixa CIELO04D (Liquidação/Pagamento).


REQUISITOS
----------
• Python 3.9 ou superior
  Download: https://www.python.org/downloads/
  ⚠ Durante a instalação, marque "Add Python to PATH"

• Nenhuma biblioteca extra é necessária.
  Usa apenas módulos padrão do Python.


COMO INICIAR
------------
Windows:
  Duplo clique em "iniciar.py"
  (se não funcionar, clique com botão direito → "Abrir com Python")

Mac / Linux:
  Terminal: python3 iniciar.py
  ou: chmod +x iniciar.py && ./iniciar.py

A aplicação abrirá automaticamente no seu navegador padrão.
URL de acesso: http://127.0.0.1:8765


COMO USAR
---------
1. Exporte o CSV do portal Cielo:
   Relatórios → Recebíveis Detalhado → Exportar CSV

2. Abra a aplicação (iniciar.py)

3. Na aba "Converter":
   a) Clique na área de upload ou arraste o arquivo CSV
   b) Confira o preview (total de linhas, bandeiras, período)
   c) Escolha a pasta de saída (ou deixe em branco para "output/")
   d) Clique em "Converter"

4. O arquivo EDI será gerado na pasta de saída:
   Formato: CIELO04D_<estab>_<data_ini>_<data_fim>.txt
   Exemplo: CIELO04D_1028105247_19032026_31032026.txt


FILTROS AUTOMÁTICOS
-------------------
✅ INCLUÍDOS:  Lançamentos com Valor bruto > 0 (créditos)
❌ EXCLUÍDOS:  Cancelamentos, chargebacks, estornos (valor negativo)

Os lançamentos excluídos ficam a cargo do analista financeiro.


ERROS DE CONVERSÃO
------------------
Se alguma linha do CSV não puder ser convertida:
• A conversão continua para as demais linhas
• Uma tabela de erros é exibida com:
    - Número da linha no CSV
    - Estabelecimento
    - Código de autorização
    - NSU/DOC
    - Descrição do erro
• Corrija as linhas indicadas no CSV e reconverta


ESTRUTURA DE PASTAS
--------------------
Após a primeira execução, serão criadas automaticamente:

  pasta_da_aplicacao/
  ├── iniciar.py              ← Iniciar aqui (duplo clique)
  ├── servidor.py             ← Servidor local
  ├── cielo_csv_para_edi.py   ← Motor de conversão
  ├── app.html                ← Interface visual
  ├── input/                  ← Coloque seus CSV aqui
  └── output/                 ← Arquivos EDI gerados aqui


PORTAS E SEGURANÇA
------------------
• A aplicação roda APENAS localmente (127.0.0.1:8765)
• Nenhum dado é enviado para a internet
• Para encerrar: feche o terminal ou pressione Ctrl+C


PROBLEMAS COMUNS
----------------
"Python não é reconhecido":
→ Reinstale o Python marcando "Add Python to PATH"

"Porta 8765 já está em uso":
→ Feche outra instância da aplicação ou reinicie o computador

"Arquivo CSV não reconhecido":
→ Verifique se é o CSV exportado do portal Cielo (separador ;)
→ O arquivo deve estar em encoding latin-1 (padrão do portal)

============================================================
