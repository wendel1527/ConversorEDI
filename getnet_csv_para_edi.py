# -*- coding: utf-8 -*-
"""
getnet_csv_para_edi.py — Motor GetNet -> Extrato Eletronico V10 (CEADM100 / 400 bytes)
=====================================================================================

Converte a planilha de "Recebiveis - Extrato Detalhado" exportada do portal GetNet
para o arquivo EDI Extrato Eletronico, layout V10 (posicao FIXA, 400 bytes por
registro, terminador CRLF, encoding latin-1). Referencia: "EXTRATO ELETRONICO
LAYOUT V10" e "Manual V10" (Getnet/Santander), validado contra arquivo real v10.1.

Modulo do projeto conversor-EDI (adquirente: GetNet).
Helpers PROPRIOS — NAO reaproveita helpers de Cielo (_mon) nem Rede (_mon_eevc).

Zero dependencias externas: apenas biblioteca padrao (zipfile + xml).

Registros gerados
-----------------
    0  Header do arquivo
    1  Detalhe do RV (Resumo de Vendas)         — 1 por grupo de RV
    2  Detalhe do CV (Comprovante de Venda)     — 1 por transacao de venda
    3  Ajustes (cancelamento / chargeback)      — 1 por ajuste
    9  Trailer (quantidade total de registros)

Regras de negocio (confirmadas com o usuario)
---------------------------------------------
* Agrupamento de RV (planilha nao expoe o numero do RV -> gerado sequencial):
    chave = Estabelecimento + Produto + Data de Vencimento + Data da Venda
* Indicador de tipo de pagamento (pos 169-170):
    "Vendas"              -> 'PF'  (pagamento futuro / previsao)
    "Pagamento Realizado" -> 'LQ'  (liquidado; conforme arquivo real GetNet)
* "Saldo Anterior"           -> ignorado
* "Cancelamento/Chargeback"  -> registro tipo 3 (ajuste), sinal '-'

Funcoes publicas: ler_xlsx(caminho) e converter(xlsx, saida).
"""

import io
import os
import re
import sys
import zipfile
import datetime
import xml.etree.ElementTree as ET

VERSAO_MODULO = "1.0.0"

# CNPJ/Nome do adquirente (header) — conforme arquivo real GetNet
CNPJ_ADQUIRENTE = "10440482000154"
NOME_ADQUIRENTE = "GETNET S.A."
COD_ADQUIRENTE = "GS"
VERSAO_LAYOUT = "SANT. V.10.1 400 BYTES"
ARQUIVO_VERSAO = "CEADM100"

TAM_REGISTRO = 400

# ---------------------------------------------------------------------------
# TABELAS DE LOOKUP (GetNet)
# ---------------------------------------------------------------------------

# Bandeira/Modalidade (coluna da planilha) -> Codigo do Produto (Tabela I, pos 17-18)
PRODUTO = {
    "visa credito": "SV",
    "visa debito": "SE",
    "mastercard credito": "SM",
    "mastercard debito": "SR",
    "elo credito": "EC",
    "elo debito": "ED",
    "amex credito": "AC",
    "amex debito": "AC",
    "hipercard credito": "HC",
    "hiper credito": "HC",
}

# Tipo de Lancamento (coluna da planilha) -> indicador tipo pagamento (pos 169-170)
INDICADOR = {
    "vendas": "PF",
    "pagamento realizado": "LQ",
}

# Motivo do ajuste (Tabela II, pos 76-77 do registro 3)
MOTIVO_AJUSTE = {
    "cancelamento": "03",
    "chargeback": "04",
    "reversao de chargeback": "16",
}

# Status da transacao (Tabela III, pos 144 do registro 2)
STATUS_APROVADA = "C"
STATUS_CANCELADA = "X"


# ===========================================================================
# LEITOR XLSX MINIMO (stdlib) — le uma aba pelo nome, resolve shared strings,
# numeros e datas (sistema serial 1900). Suficiente para a planilha GetNet.
# ===========================================================================

_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
       "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

# numFmtId builtin considerados data/hora no OOXML
_BUILTIN_DATE_FMT = {14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47}


def _col_to_idx(ref):
    """'B8' -> 1 (indice 0-based da coluna)."""
    letters = re.match(r"([A-Z]+)", ref).group(1)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _serial_para_data(serial):
    """Converte serial de data do Excel (sistema 1900) para datetime.date/datetime."""
    try:
        serial = float(serial)
    except (TypeError, ValueError):
        return None
    # Excel trata 1900 como bissexto (bug historico): dias >= 60 sao deslocados em 1.
    dias = int(serial)
    frac = serial - dias
    if dias >= 60:
        dias -= 1
    base = datetime.datetime(1899, 12, 31)
    dt = base + datetime.timedelta(days=dias, seconds=round(frac * 86400))
    return dt


def _ler_shared_strings(z):
    try:
        data = z.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    out = []
    for si in root.findall("m:si", _NS):
        # concatena todos os <t> (texto simples ou rich text)
        txt = "".join(t.text or "" for t in si.iter("{%s}t" % _NS["m"]))
        out.append(txt)
    return out


def _ler_estilos_datas(z):
    """Retorna set de indices de cellXfs (s=...) cujo numFmt e' data/hora."""
    try:
        data = z.read("xl/styles.xml")
    except KeyError:
        return set()
    root = ET.fromstring(data)
    # numFmts customizados
    custom_date = set()
    numfmts = root.find("m:numFmts", _NS)
    if numfmts is not None:
        for nf in numfmts.findall("m:numFmt", _NS):
            fid = int(nf.get("numFmtId"))
            code = (nf.get("formatCode") or "").lower()
            if any(tok in code for tok in ("yy", "mm", "dd", "hh", "ss")) and "[" not in code:
                custom_date.add(fid)
    date_style_idx = set()
    cellxfs = root.find("m:cellXfs", _NS)
    if cellxfs is not None:
        for i, xf in enumerate(cellxfs.findall("m:xf", _NS)):
            fid = int(xf.get("numFmtId", "0"))
            if fid in _BUILTIN_DATE_FMT or fid in custom_date:
                date_style_idx.add(i)
    return date_style_idx


def _nome_para_arquivo_sheet(z, nome):
    """Mapeia nome da aba -> caminho xl/worksheets/sheetN.xml."""
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid_para_alvo = {}
    for rel in rels:
        rid_para_alvo[rel.get("Id")] = rel.get("Target")
    for sh in wb.find("m:sheets", _NS).findall("m:sheet", _NS):
        if sh.get("name") == nome:
            rid = sh.get("{%s}id" % _NS["r"])
            alvo = rid_para_alvo.get(rid, "")
            if not alvo.startswith("xl/"):
                alvo = "xl/" + alvo.lstrip("/")
            return alvo
    return None


def ler_aba_xlsx(caminho, nome_aba):
    """Le uma aba do XLSX e retorna lista de linhas (cada linha = lista de celulas).

    Cada celula ja vem no tipo apropriado: str, float/int ou datetime.
    Celulas vazias viram None. Preenche o comprimento das linhas de forma esparsa.
    """
    with zipfile.ZipFile(caminho) as z:
        shared = _ler_shared_strings(z)
        estilos_data = _ler_estilos_datas(z)
        alvo = _nome_para_arquivo_sheet(z, nome_aba)
        if alvo is None:
            raise ValueError("Aba '%s' nao encontrada no XLSX." % nome_aba)
        root = ET.fromstring(z.read(alvo))

    sheetdata = root.find("m:sheetData", _NS)
    linhas = []
    for row in sheetdata.findall("m:row", _NS):
        celulas = {}
        maxc = -1
        for c in row.findall("m:c", _NS):
            ref = c.get("r")
            idx = _col_to_idx(ref)
            maxc = max(maxc, idx)
            t = c.get("t")            # tipo: s=shared, str, inlineStr, b, e...
            s = c.get("s")            # indice de estilo
            v_el = c.find("m:v", _NS)
            valor = None
            if t == "s":
                if v_el is not None and v_el.text is not None:
                    valor = shared[int(v_el.text)]
            elif t == "inlineStr":
                is_el = c.find("m:is", _NS)
                if is_el is not None:
                    valor = "".join(x.text or "" for x in is_el.iter("{%s}t" % _NS["m"]))
            elif t == "str":
                valor = v_el.text if v_el is not None else None
            elif t == "b":
                valor = (v_el.text == "1") if v_el is not None else None
            else:
                # numero (ou data, conforme estilo)
                if v_el is not None and v_el.text not in (None, ""):
                    if s is not None and int(s) in estilos_data:
                        valor = _serial_para_data(v_el.text)
                    else:
                        num = float(v_el.text)
                        valor = int(num) if num.is_integer() else num
            celulas[idx] = valor
        linha = [celulas.get(i) for i in range(maxc + 1)]
        linhas.append(linha)
    return linhas


# ===========================================================================
# HELPERS DE FORMATACAO (PROPRIOS DO MODULO GETNET)
# ===========================================================================

def _txt(v):
    """Normaliza celula para string 'crua' (trata None e o marcador '-')."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if s == "-":
        return ""
    return s


def _n(val, tam):
    """Numerico: extrai apenas digitos, zero-fill a esquerda, corta a 'tam'."""
    s = re.sub(r"\D", "", str(val or ""))
    return s.rjust(tam, "0")[-tam:] if s else "0" * tam


def _a(val, tam):
    """Alfanumerico: left-justify com espacos, corta a 'tam'."""
    s = _txt(val)
    return s.ljust(tam)[:tam]


def _val12(valor):
    """Valor monetario -> 12 digitos, 2 casas implicitas, SEM sinal.
    Ex.: 1057.00 -> '000000105700' ; -30.82 -> '000000003082'."""
    try:
        centavos = int(round(abs(float(valor)) * 100))
    except (TypeError, ValueError):
        centavos = 0
    return str(centavos).rjust(12, "0")[-12:]


def _sinal(valor):
    """'+' para credito (>=0), '-' para debito (<0)."""
    try:
        return "-" if float(valor) < 0 else "+"
    except (TypeError, ValueError):
        return "+"


def _para_data(v):
    """Normaliza celula (datetime, serial Excel ou 'dd/mm/aaaa') -> datetime.date ou None."""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    # Serial do Excel (int/float) — colunas de data podem vir SEM estilo de data,
    # chegando como numero cru (ex.: 46253 = 18/08/2026). Faixa ~1954..2119.
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 20000 <= float(v) <= 80000:
            dt = _serial_para_data(v)
            return dt.date() if dt else None
    s = _txt(v)
    if not s:
        return None
    # numero em texto tambem pode ser serial
    if re.fullmatch(r"\d{4,6}", s):
        n = int(s)
        if 20000 <= n <= 80000:
            dt = _serial_para_data(n)
            if dt:
                return dt.date()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _data_ddmmaaaa(v):
    """-> 'DDMMAAAA' (8). Vazio/invalido -> '00000000'."""
    d = _para_data(v)
    return d.strftime("%d%m%Y") if d else "00000000"


def _hora_hhmmss(v):
    """'16:10:29' / '16:10' -> 'HHMMSS' (6). Vazio -> '000000'."""
    s = _txt(v)
    if not s:
        return "000000"
    partes = re.split(r"[:h]", s)
    partes = [p for p in partes if p != ""]
    while len(partes) < 3:
        partes.append("0")
    try:
        h, m, seg = (int(partes[0]), int(partes[1]), int(partes[2]))
        return "%02d%02d%02d" % (h, m, seg)
    except ValueError:
        return "000000"


def _estab15(v):
    """Codigo do estabelecimento -> 15 pos, alfanumerico left-justify (so digitos uteis)."""
    s = re.sub(r"\D", "", _txt(v))
    return s.ljust(15)[:15]


def _cnpj14(v):
    return _n(v, 14)


def _cartao19(v):
    """Numero do cartao mascarado -> 19 pos, mantendo a mascara do CSV."""
    s = _txt(v)
    return s.ljust(19)[:19]


def _parcelas(v):
    """'2 de 6' -> (2, 6) ; '1 de 1' -> (1,1) ; '-'/vazio -> (1,1)."""
    s = _txt(v)
    m = re.search(r"(\d+)\s*de\s*(\d+)", s, re.IGNORECASE)
    if m:
        return int(m.group(1)), int(m.group(2))
    return 1, 1


def _autoriz(v, tam=10):
    """Codigo de autorizacao -> 'tam' pos com ZERO-FILL a esquerda (convencao GetNet).
    Preserva caracteres alfanumericos. Ex.: '603019' -> '0000603019' ; 'YTHT8A' -> '0000YTHT8A'."""
    s = _txt(v)
    if not s:
        return "0" * tam
    return s.rjust(tam, "0")[-tam:]


def _produto(bandeira):
    """Bandeira/Modalidade -> codigo do produto (Tabela I). Fallback: espacos."""
    chave = _txt(bandeira).lower()
    chave = (chave.replace("é", "e").replace("ê", "e").replace("í", "i")
                  .replace("ó", "o").replace("á", "a").replace("â", "a")
                  .replace("ã", "a").replace("ç", "c"))
    chave = re.sub(r"\s+", " ", chave).strip()
    return PRODUTO.get(chave, "  ")


def _captura(terminal):
    """Heuristica de forma de captura a partir do terminal logico.
    Terminais 'TF...' costumam ser TEF; numericos, POS. Fallback: POS."""
    s = _txt(terminal).upper()
    if s.startswith("TF"):
        return "TEF"
    if s:
        return "POS"
    return "POS"


# ===========================================================================
# LEITURA / NORMALIZACAO DA PLANILHA
# ===========================================================================

# indices (0-based) das colunas na aba "Detalhado" (cabecalho na linha 8 do Excel)
COL = {
    "ec_centralizador": 0,
    "estabelecimento": 1,
    "cpf_cnpj": 2,
    "data_vencimento": 3,
    "bandeira": 4,
    "tipo_lancamento": 5,
    "lancamento": 6,
    "valor_liquido": 7,
    "valor_liquidado": 8,
    "cartao": 9,
    "autorizacao": 10,
    "nsu": 11,
    "terminal": 12,
    "data_venda": 13,
    "hora_venda": 14,
    "valor_venda": 15,
    "parcelas": 16,
    "valor_parcela": 17,
    "descontos": 18,
    "valor_liq_parcela": 19,
    "data_contrato": 20,
    "instituicao_neg": 21,
    "contrato_reg": 22,
    "valor_atual_contrato": 23,
    "valor_liq_contrato": 24,
    "valor_a_liquidar": 25,
}

ABA_DADOS = "Detalhado"
TITULO_ESPERADO = "Recebimentos - Extrato Detalhado"


def ler_xlsx(caminho):
    """Le a aba 'Detalhado' e retorna (linhas, meta).

    linhas : lista de dicts, uma por transacao valida (exclui 'Saldo Anterior').
    meta   : dict com estabelecimento, cpf_cnpj, data do movimento, contadores.
    """
    todas = ler_aba_xlsx(caminho, ABA_DADOS)

    # localiza a linha de cabecalho (contem 'ESTABELECIMENTO COMERCIAL')
    hdr_idx = None
    for i, linha in enumerate(todas[:20]):
        if any(_txt(c).upper() == "ESTABELECIMENTO COMERCIAL" for c in linha):
            hdr_idx = i
            break
    if hdr_idx is None:
        raise ValueError("Cabecalho da aba 'Detalhado' nao localizado.")

    linhas = []
    ignoradas_saldo = 0
    estab_arquivo = None
    cpf_arquivo = None
    venc_min = venc_max = None

    for linha in todas[hdr_idx + 1:]:
        def get(nome):
            idx = COL[nome]
            return linha[idx] if idx < len(linha) else None

        estab = _txt(get("estabelecimento"))
        if not estab:
            continue

        tipo = _txt(get("tipo_lancamento"))
        if tipo.lower().startswith("saldo anterior"):
            ignoradas_saldo += 1
            continue

        reg = {
            "ec_centralizador": _txt(get("ec_centralizador")),
            "estabelecimento": estab,
            "cpf_cnpj": _txt(get("cpf_cnpj")),
            "data_vencimento": get("data_vencimento"),
            "bandeira": _txt(get("bandeira")),
            "tipo_lancamento": tipo,
            "lancamento": _txt(get("lancamento")),
            "valor_liquido": get("valor_liquido"),
            "valor_liquidado": get("valor_liquidado"),
            "cartao": _txt(get("cartao")),
            "autorizacao": _txt(get("autorizacao")),
            "nsu": _txt(get("nsu")),
            "terminal": _txt(get("terminal")),
            "data_venda": get("data_venda"),
            "hora_venda": _txt(get("hora_venda")),
            "valor_venda": get("valor_venda"),
            "parcelas": _txt(get("parcelas")),
            "valor_parcela": get("valor_parcela"),
            "descontos": get("descontos"),
            "valor_liq_parcela": get("valor_liq_parcela"),
        }
        linhas.append(reg)

        if estab_arquivo is None:
            estab_arquivo = estab
            cpf_arquivo = reg["cpf_cnpj"]
        d = _para_data(reg["data_vencimento"])
        if d:
            venc_min = d if venc_min is None or d < venc_min else venc_min
            venc_max = d if venc_max is None or d > venc_max else venc_max

    meta = {
        "estabelecimento": estab_arquivo or "",
        "cpf_cnpj": cpf_arquivo or "",
        "data_movimento": venc_max or datetime.date.today(),
        "venc_min": venc_min,
        "venc_max": venc_max,
        "ignoradas_saldo": ignoradas_saldo,
        "total_linhas": len(linhas),
    }
    return linhas, meta


# ===========================================================================
# GERACAO DOS REGISTROS (posicao fixa, 400 pos)
# ===========================================================================

def _campo(buf, ini, fim, valor):
    """Escreve 'valor' no buffer (lista de chars) nas posicoes 1-based [ini, fim]."""
    valor = str(valor)
    tam = fim - ini + 1
    valor = valor[:tam].ljust(tam)
    buf[ini - 1:fim] = list(valor)


def _linha_branca():
    return [" "] * TAM_REGISTRO


def gerar_registro_0(meta, sequencia=1):
    """Header do arquivo (tipo 0)."""
    buf = _linha_branca()
    agora = datetime.datetime.now()
    _campo(buf, 1, 1, "0")
    _campo(buf, 2, 9, agora.strftime("%d%m%Y"))
    _campo(buf, 10, 15, agora.strftime("%H%M%S"))
    _campo(buf, 16, 23, meta["data_movimento"].strftime("%d%m%Y"))
    _campo(buf, 24, 31, ARQUIVO_VERSAO)
    _campo(buf, 32, 46, _estab15(meta["estabelecimento"]))
    _campo(buf, 47, 60, _cnpj14(CNPJ_ADQUIRENTE))
    _campo(buf, 61, 80, _a(NOME_ADQUIRENTE, 20))
    _campo(buf, 81, 89, _n(sequencia, 9))
    _campo(buf, 90, 91, _a(COD_ADQUIRENTE, 2))
    _campo(buf, 92, 116, _a(VERSAO_LAYOUT, 25))
    # 117-400 reservado (espacos)
    return "".join(buf)


def gerar_registro_1(rv, meta):
    """Detalhe do RV (tipo 1). 'rv' e' o dict do grupo agregado."""
    buf = _linha_branca()
    produto = rv["produto"]
    _campo(buf, 1, 1, "1")
    _campo(buf, 2, 16, _estab15(rv["estabelecimento"]))
    _campo(buf, 17, 18, produto)                              # produto/bandeira
    _campo(buf, 19, 21, _a(rv["captura"], 3))                 # forma de captura
    _campo(buf, 22, 30, _n(rv["numero_rv"], 9))               # numero do RV (sintetico)
    _campo(buf, 31, 38, _data_ddmmaaaa(rv["data_venda"]))     # data do RV
    _campo(buf, 39, 46, _data_ddmmaaaa(rv["data_vencimento"]))# data do pagamento
    _campo(buf, 47, 49, _n(0, 3))                             # banco (nao exportado)
    _campo(buf, 50, 55, _n(0, 6))                             # agencia (nao exportado)
    _campo(buf, 56, 66, _n(0, 11))                            # conta (nao exportado)
    _campo(buf, 67, 75, _n(rv["qtd_cv"], 9))                  # CVs aceitos
    _campo(buf, 76, 84, _n(0, 9))                             # CVs rejeitados
    _campo(buf, 85, 96, _val12(rv["valor_bruto"]))           # valor bruto
    _campo(buf, 97, 108, _val12(rv["valor_liquido"]))        # valor liquido
    _campo(buf, 109, 120, _val12(0))                         # tarifa
    _campo(buf, 121, 132, _val12(rv["valor_desconto"]))      # taxa de desconto (MDR)
    _campo(buf, 133, 144, _val12(0))                         # valor rejeitado
    _campo(buf, 145, 156, _val12(rv["valor_liquido"]))       # valor credito
    _campo(buf, 157, 168, _val12(0))                         # valor encargos
    _campo(buf, 169, 170, rv["indicador"])                   # indicador tipo pagamento
    _campo(buf, 171, 172, _n(rv["parcela"], 2))              # numero da parcela
    _campo(buf, 173, 174, _n(rv["qtd_parcelas"], 2))         # qtd de parcelas
    _campo(buf, 175, 189, _estab15(rv["ec_centralizador"] or rv["estabelecimento"]))
    _campo(buf, 282, 284, "986")                             # moeda (real)
    _campo(buf, 286, 286, rv["sinal"])                       # sinal
    _campo(buf, 287, 288, "CC")                              # metadado 1 (conta corrente)
    return "".join(buf)


def gerar_registro_2(cv, rv):
    """Detalhe do CV (tipo 2)."""
    buf = _linha_branca()
    parc, total = cv["parcela"], cv["qtd_parcelas"]
    _campo(buf, 1, 1, "2")
    _campo(buf, 2, 16, _estab15(cv["estabelecimento"]))
    _campo(buf, 17, 25, _n(rv["numero_rv"], 9))              # numero do RV (pai)
    _campo(buf, 26, 37, _n(cv["nsu"], 12))                   # NSU
    _campo(buf, 38, 45, _data_ddmmaaaa(cv["data_venda"]))    # data da transacao
    _campo(buf, 46, 51, _hora_hhmmss(cv["hora_venda"]))      # hora da transacao
    _campo(buf, 52, 70, _cartao19(cv["cartao"]))             # numero do cartao
    _campo(buf, 71, 82, _val12(cv["valor_venda"]))           # valor da transacao
    _campo(buf, 83, 94, _val12(0))                           # valor do saque
    _campo(buf, 95, 106, _val12(0))                          # taxa de embarque
    _campo(buf, 107, 108, _n(total, 2))                      # total de parcelas
    _campo(buf, 109, 110, _n(parc, 2))                       # numero da parcela
    _campo(buf, 111, 122, _val12(cv["valor_parcela"]))       # valor da parcela
    _campo(buf, 123, 130, _data_ddmmaaaa(rv["data_vencimento"]))  # data pagamento
    _campo(buf, 131, 140, _autoriz(cv["autorizacao"], 10))   # codigo de autorizacao (zero-fill esq.)
    _campo(buf, 141, 143, _a(cv["captura"], 3))              # forma de captura
    _campo(buf, 144, 144, cv["status"])                      # status da transacao
    _campo(buf, 145, 159, _estab15(cv["ec_centralizador"] or cv["estabelecimento"]))  # centralizador
    _campo(buf, 160, 167, _a(cv["terminal"], 8))            # codigo do terminal
    _campo(buf, 168, 170, "986")                             # moeda
    _campo(buf, 171, 171, "N")                               # origem emissor (Brasil)
    _campo(buf, 172, 172, cv["sinal"])                       # sinal
    _campo(buf, 176, 187, _val12(cv["valor_comissao"]))      # comissao (MDR)
    return "".join(buf)


def gerar_registro_3(aj, numero_rv):
    """Ajuste (tipo 3) — cancelamento / chargeback."""
    buf = _linha_branca()
    _campo(buf, 1, 1, "3")
    _campo(buf, 2, 16, _estab15(aj["estabelecimento"]))
    _campo(buf, 17, 25, _n(numero_rv, 9))                    # RV ajustado
    _campo(buf, 26, 33, _data_ddmmaaaa(aj["data_venda"]))    # data do RV
    _campo(buf, 34, 41, _data_ddmmaaaa(aj["data_vencimento"]))  # data do pagamento
    _campo(buf, 63, 63, "-")                                 # sinal do ajuste (debito)
    _campo(buf, 64, 75, _val12(aj["valor_liq_parcela"]))     # valor do ajuste
    _campo(buf, 76, 77, aj["motivo"])                        # motivo (Tabela II)
    _campo(buf, 86, 104, _cartao19(aj["cartao"]))            # numero do cartao
    _campo(buf, 105, 113, _n(numero_rv, 9))                  # RV original
    _campo(buf, 114, 125, _n(aj["nsu"], 12))                 # NSU original
    _campo(buf, 126, 133, _data_ddmmaaaa(aj["data_venda"]))  # data transacao original
    _campo(buf, 134, 135, "LQ")                             # indicador tipo pagamento
    _campo(buf, 136, 143, _a(aj.get("terminal", ""), 8))    # terminal (transacao original)
    _campo(buf, 144, 151, _data_ddmmaaaa(aj["data_vencimento"]))  # data pagamento original
    _campo(buf, 152, 154, "986")                             # moeda
    return "".join(buf)


def gerar_registro_9(qtd_total):
    """Trailer (tipo 9)."""
    buf = _linha_branca()
    _campo(buf, 1, 1, "9")
    _campo(buf, 2, 10, _n(qtd_total, 9))
    return "".join(buf)


# ---------------------------------------------------------------------------
# CONVENCAO DE PREENCHIMENTO GETNET (zero-fill de campos numericos)
# ---------------------------------------------------------------------------
# A GetNet zero-preenche TODOS os campos numericos ao longo dos 400 bytes; deixar
# espacos em posicoes numericas faz o parser do conciliador rejeitar o arquivo.
# Ranges (1-based, inclusivos) derivados do arquivo real por tipo de registro:
# posicao e' zero-fill quando e' digito em TODOS os registros reais daquele tipo.
_ZERO_FILL_RANGES = {
    "1": [(22, 168), (171, 180), (190, 284), (289, 333)],
    "2": [(17, 57), (64, 66), (71, 134), (145, 150), (162, 170), (176, 187)],
    "3": [(17, 56), (64, 86), (105, 133), (152, 166)],
}


def _aplicar_zero_fill(linha):
    """Substitui ESPACO por '0' nas posicoes numericas (zero-fill) do tipo do registro.
    Preserva os valores ja escritos (so afeta espacos remanescentes)."""
    tipo = linha[0]
    ranges = _ZERO_FILL_RANGES.get(tipo)
    if not ranges:
        return linha
    buf = list(linha)
    for ini, fim in ranges:
        for p in range(ini - 1, fim):
            if buf[p] == " ":
                buf[p] = "0"
    return "".join(buf)


# ===========================================================================
# PIPELINE DE CONVERSAO
# ===========================================================================

def _classificar(reg):
    """Retorna 'venda', 'liquidacao', 'ajuste' ou None."""
    tipo = reg["tipo_lancamento"].lower()
    lanc = reg["lancamento"].lower()
    if "cancelamento" in tipo or "chargeback" in tipo or "cancelamento" in lanc:
        return "ajuste"
    if "pagamento realizado" in tipo or "liquidado" in lanc:
        return "liquidacao"
    if "venda" in tipo or "venda" in lanc:
        return "venda"
    return None


def _motivo_ajuste(reg):
    """Motivo do ajuste (Tabela II) a partir da coluna 'Lancamento'.
    Obs.: a coluna 'Tipo de Lancamento' vale sempre 'Cancelamento/Chargeback',
    entao a distincao e' feita pelo 'Lancamento' (ex.: 'Cancelamento De Venda')."""
    lanc = reg["lancamento"].lower()
    if "reversao" in lanc and "chargeback" in lanc:
        return MOTIVO_AJUSTE["reversao de chargeback"]
    if "chargeback" in lanc:
        return MOTIVO_AJUSTE["chargeback"]
    return MOTIVO_AJUSTE["cancelamento"]


def converter(xlsx, saida=None, log=None, forcar_lq=False):
    """Pipeline completo: le a planilha GetNet e grava o EDI Extrato Eletronico V10.

    forcar_lq: se True, todos os RVs recebem indicador 'LQ' (liquidado) — util quando
    o conciliador so reconhece o codigo que a GetNet realmente emite. Se False (padrao),
    usa 'PF' para vendas futuras e 'LQ' para pagamentos realizados.

    Retorna o caminho do arquivo gerado.
    """
    if log is None:
        log = sys.stdout

    linhas, meta = ler_xlsx(xlsx)
    log.write("GetNet -> Extrato Eletronico V10 (modulo %s)\n" % VERSAO_MODULO)
    log.write("Estabelecimento: %s | CPF/CNPJ: %s\n" % (meta["estabelecimento"], meta["cpf_cnpj"]))
    log.write("Linhas uteis: %d | 'Saldo Anterior' ignoradas: %d\n"
              % (meta["total_linhas"], meta["ignoradas_saldo"]))

    registros = []            # linhas finais do arquivo
    erros = []
    seq_rv = 0
    grupos = {}               # chave -> numero_rv atribuido
    rv_dados = {}             # numero_rv -> dict agregado do RV
    rv_cvs = {}               # numero_rv -> lista de CVs
    ajustes = []              # lista de (aj, numero_rv)

    n_venda = n_liq = n_aj = 0

    for i, reg in enumerate(linhas, start=1):
        try:
            classe = _classificar(reg)
            produto = _produto(reg["bandeira"])
            captura = _captura(reg["terminal"])
            parc, total = _parcelas(reg["parcelas"])

            if classe == "ajuste":
                ajustes.append({
                    "estabelecimento": reg["estabelecimento"],
                    "data_venda": reg["data_venda"],
                    "data_vencimento": reg["data_vencimento"],
                    "valor_liq_parcela": reg["valor_liq_parcela"] if reg["valor_liq_parcela"] not in (None, "") else reg["valor_liquido"],
                    "motivo": _motivo_ajuste(reg),
                    "cartao": reg["cartao"],
                    "nsu": reg["nsu"],
                    "terminal": reg["terminal"],
                })
                n_aj += 1
                continue

            # chave de agrupamento de RV: inclui parcela e total para que todos os
            # CVs de um RV tenham o MESMO plano (evita RV '03/06' com CV '03/04',
            # que o conciliador acusa como erro no crédito parcelado).
            d_venc = _para_data(reg["data_vencimento"])
            d_venda = _para_data(reg["data_venda"])
            chave = (reg["estabelecimento"], produto,
                     d_venc.isoformat() if d_venc else "",
                     d_venda.isoformat() if d_venda else "",
                     parc, total)

            if chave not in grupos:
                seq_rv += 1
                grupos[chave] = seq_rv
                indicador = INDICADOR.get(reg["tipo_lancamento"].lower(),
                                          "LQ" if classe == "liquidacao" else "PF")
                if forcar_lq:
                    indicador = "LQ"
                rv_dados[seq_rv] = {
                    "numero_rv": seq_rv,
                    "estabelecimento": reg["estabelecimento"],
                    "ec_centralizador": reg["ec_centralizador"],
                    "produto": produto,
                    "captura": captura,
                    # data do RV: usa a data da venda; se ausente (ex.: liquidação),
                    # cai para a data de vencimento (nunca deixa 00000000, que quebra o parser)
                    "data_venda": reg["data_venda"] if _para_data(reg["data_venda"]) else reg["data_vencimento"],
                    "data_vencimento": reg["data_vencimento"],
                    "indicador": indicador,
                    "parcela": parc,
                    "qtd_parcelas": total,
                    "valor_bruto": 0.0,
                    "valor_liquido": 0.0,
                    "valor_desconto": 0.0,
                    "qtd_cv": 0,
                    "sinal": "+",
                }
                rv_cvs[seq_rv] = []

            numero_rv = grupos[chave]
            rv = rv_dados[numero_rv]

            def _f(x):
                try:
                    return float(x)
                except (TypeError, ValueError):
                    return 0.0

            if classe == "venda":
                cv = {
                    "estabelecimento": reg["estabelecimento"],
                    "ec_centralizador": reg["ec_centralizador"],
                    "nsu": reg["nsu"],
                    "data_venda": reg["data_venda"],
                    "hora_venda": reg["hora_venda"],
                    "cartao": reg["cartao"],
                    "valor_venda": _f(reg["valor_venda"]),
                    "valor_parcela": _f(reg["valor_parcela"]),
                    "parcela": parc,
                    "qtd_parcelas": total,
                    "autorizacao": reg["autorizacao"],
                    "captura": captura,
                    "terminal": reg["terminal"],
                    "status": STATUS_APROVADA,
                    "valor_comissao": abs(_f(reg["descontos"])),
                    "sinal": _sinal(reg["valor_liq_parcela"] if reg["valor_liq_parcela"] not in (None, "") else 1),
                }
                rv_cvs[numero_rv].append(cv)
                rv["valor_bruto"] += abs(_f(reg["valor_parcela"]))
                rv["valor_liquido"] += abs(_f(reg["valor_liq_parcela"]))
                rv["valor_desconto"] += abs(_f(reg["descontos"]))
                rv["qtd_cv"] += 1
                n_venda += 1
            else:  # liquidacao (sem CV — valor agregado)
                v = abs(_f(reg["valor_liq_parcela"] if reg["valor_liq_parcela"] not in (None, "") else reg["valor_liquido"]))
                rv["valor_bruto"] += v
                rv["valor_liquido"] += v
                n_liq += 1

        except Exception as e:  # noqa: BLE001 — reporta e continua
            erros.append((i, reg.get("estabelecimento", ""), reg.get("nsu", ""), str(e)))

    # monta o arquivo na ordem: 0, (1 + N*2)..., 3..., 9
    registros.append(gerar_registro_0(meta, sequencia=1))
    for numero_rv in sorted(rv_dados):
        rv = rv_dados[numero_rv]
        registros.append(_aplicar_zero_fill(gerar_registro_1(rv, meta)))
        for cv in rv_cvs[numero_rv]:
            registros.append(_aplicar_zero_fill(gerar_registro_2(cv, rv)))
    for aj in ajustes:
        seq_rv += 1
        registros.append(_aplicar_zero_fill(gerar_registro_3(aj, seq_rv)))
    total_registros = len(registros) + 1  # inclui o proprio trailer
    registros.append(gerar_registro_9(total_registros))

    # nome do arquivo de saida
    estab = re.sub(r"\D", "", meta["estabelecimento"]) or "GETNET"
    di = (meta["venc_min"] or meta["data_movimento"]).strftime("%d%m%Y")
    df = (meta["venc_max"] or meta["data_movimento"]).strftime("%d%m%Y")
    nome = "GETNET_EE_V10_%s_%s_%s.txt" % (estab, di, df)

    if saida is None:
        destino = nome
    elif os.path.isdir(saida):
        destino = os.path.join(saida, nome)
    else:
        destino = saida

    with open(destino, "w", encoding="latin-1", newline="") as fh:
        for r in registros:
            fh.write(r + "\r\n")

    log.write("RVs gerados: %d | CVs: %d | Liquidacoes: %d | Ajustes: %d\n"
              % (len(rv_dados), n_venda, n_liq, n_aj))
    log.write("Total de registros no arquivo: %d\n" % total_registros)
    if erros:
        log.write("\n%d linha(s) com erro:\n" % len(erros))
        for lin, est, nsu, msg in erros[:20]:
            log.write("  linha %d | estab %s | NSU %s | %s\n" % (lin, est, nsu, msg))
    log.write("Arquivo gerado: %s\n" % destino)
    return destino


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("uso: python getnet_csv_para_edi.py <planilha.xlsx> [saida|pasta]")
        sys.exit(1)
    entrada = sys.argv[1]
    destino = sys.argv[2] if len(sys.argv) > 2 else None
    converter(entrada, destino)
