"""
=============================================================================
  Conversor: CSV Recebíveis Cielo (portal) → EDI Posição Fixa (CIELO04D)
  Manual de Especificação Técnica v15.15 – fev/2026
  Versão 1.0.1 – correção _estab() para CNPJ formatado
=============================================================================
  Entrada  : CSV exportado do portal Cielo (separador ';', encoding latin-1)
             Pasta: input/
  Saída    : Arquivo TXT EDI posição fixa no formato CIELO04D
             Pasta: output/
             Nome : CIELO04D_<estab>_<data_proc>_<per_ini>_<per_fim>.txt

  Estrutura do arquivo gerado:
    Registro 0  – Header
    Registro D  – UR Agenda (1 por grupo de Chave UR)
    Registro E  – Detalhe do Lançamento (1 por linha do CSV)
    Registro 9  – Trailer

  Uso:
    python cielo_csv_para_edi.py                  → modo interativo
    python cielo_csv_para_edi.py arquivo.csv       → converte direto
    python cielo_csv_para_edi.py arquivo.csv saida/ → pasta de saída custom
=============================================================================
"""

import csv
import json
import re
import shutil
import sys
from collections import defaultdict
from datetime import datetime, date
from pathlib import Path


# ============================================================================
# TABELAS DE LOOKUP  (CSV texto → código EDI)
# ============================================================================

BANDEIRA = {
    "Visa"            : "001",
    "Mastercard"      : "002",
    "American Express": "003",
    "Amex"            : "003",
    "Elo"             : "007",
    "Hipercard"       : "009",
    "Diners"          : "005",
    "JCB"             : "006",
    "Cabal"           : "008",
    "Maestro"         : "002",
    "Discover"        : "004",
}

TIPO_LANCAMENTO = {
    "Venda à vista"             : "01",
    "Venda débito"              : "02",
    "Venda parcelada"           : "03",
    "Cancelamento de venda"     : "06",
    "Chargeback débito"         : "07",
    "Chargeback crédito"        : "08",
    "Chargeback parcelado"      : "09",
    "Ajuste crédito"            : "04",
    "Ajuste débito"             : "05",
    "Negociação de recebíveis"  : "11",
    "Cessão de recebíveis"      : "11",
}

# Forma de pagamento: chave = (bandeira_cod, modalidade_base)
# Modalidade base: remover "XX/YY" do final
FORMA_PAGAMENTO = {
    # Visa
    ("001", "crédito à vista")          : "040",
    ("001", "visa crédito à vista")     : "040",
    ("001", "débito à vista")           : "041",
    ("001", "visa electron débito")     : "041",
    ("001", "parcelado loja")           : "043",
    ("001", "visa parcelado loja")      : "043",
    ("001", "parcelado cliente")        : "030",
    ("001", "visa parcelado cliente")   : "030",
    ("001", "crediário")                : "064",
    # Mastercard
    ("002", "crédito à vista")          : "010",
    ("002", "mastercard crédito à vista"): "010",
    ("002", "débito à vista")           : "011",
    ("002", "maestro")                  : "011",
    ("002", "parcelado loja")           : "012",
    ("002", "mastercard parcelado loja"): "012",
    ("002", "parcelado cliente")        : "012",
    # American Express
    ("003", "crédito à vista")          : "082",
    ("003", "amex crédito à vista")     : "082",
    ("003", "parcelado loja")           : "083",
    ("003", "amex parcelado loja")      : "083",
    ("003", "parcelado banco")          : "084",
    # Elo
    ("007", "crédito à vista")          : "070",
    ("007", "elo crédito a vista")      : "070",
    ("007", "débito à vista")           : "071",
    ("007", "parcelado loja")           : "072",
    ("007", "elo parcelado loja")       : "072",
    # Hipercard
    ("009", "crédito à vista")          : "164",
    ("009", "parcelado loja")           : "165",
    # Diners
    ("005", "crédito à vista")          : "020",
    ("005", "parcelado loja")           : "021",
}

STATUS_PAGAMENTO = {
    "Pago"               : "10",
    "Previsto"           : "00",
    "Enviado para banco" : "03",
    "Agendado"           : "00",
    "Rejeitado"          : "06",
    "Pendente"           : "00",
    "Antecipado"         : "04",
    "Debitado em conta"  : "46",
}

CANAL_VENDA = {
    "Máquina"    : "001",
    "Internet"   : "004",
    "E-Commerce" : "007",
    "TEF"        : "008",
    "Mobile"     : "010",
    "Pix"        : "019",
}

TIPO_CAPTURA = {
    "Leitura de chip"  : "07",
    "Contactless"      : "07",
    "Tarja magnética"  : "05",
    "Venda digitada"   : "81",
    "E-Commerce"       : "10",
    "Não se aplica"    : "05",
    ""                 : "05",
}

# Tabela IX do manual: Descrição CSV → código de ajuste EDI
DESCRICAO_AJUSTE = {
    "Cobrança de venda cancelada pelo estabelecimento comercial" : "0151",
    "Cancelamento de transação"                                  : "0151",
    "Chargeback"                                                 : "0171",
    "Contestação de venda"                                       : "0171",
    "Aluguel de equipamento"                                     : "0052",
    "Plano cielo"                                                : "0054",
    "Cobrança indevida de aluguel de máquina"                    : "0065",
}

ORIGEM_CARTAO = {
    "Emitido no Brasil"   : "N",
    "Emitido no exterior" : "S",
    ""                    : "N",
}

TIPO_LIQUIDACAO = {
    "crédito"  : "002",
    "débito"   : "001",
    "voucher"  : "004",
    ""         : "000",
}


# ============================================================================
# HELPERS DE FORMATAÇÃO
# ============================================================================

def _n(val, tam: int) -> str:
    """Formata inteiro com zeros à esquerda, truncando se necessário."""
    try:
        v = int(round(float(str(val).replace(",", ".").replace(" ", "") or 0)))
    except (ValueError, TypeError):
        v = 0
    return str(abs(v)).zfill(tam)[-tam:]


def _estab(val, tam: int = 10) -> str:
    """
    Formata número de estabelecimento/CNPJ para o EDI.
    Remove pontuação (pontos, barras, traços) antes de aplicar zfill.
    Ex: '28.084.180/30'   → '2808418030'
        '028.084.180/0001-90' → '0280841800'  (trunca para tam)
        '1028105247'       → '1028105247'
    Use SEMPRE este helper para campos de estabelecimento — nunca _n().
    """
    s = re.sub(r"[.\-/]", "", str(val or "").strip())
    # Manter apenas dígitos
    s = re.sub(r"\D", "", s)
    return s.zfill(tam)[-tam:]


def _a(val, tam: int) -> str:
    """Formata alfanumérico com espaços à direita, truncando se necessário."""
    s = str(val or "").strip()
    return s[:tam].ljust(tam)


def _mon(valor_float: float, tam: int = 13) -> tuple[str, str]:
    """
    Converte float para (sinal, valor_inteiro_sem_decimal).
    Ex: -10.87 → ('-', '0000000001087')
    """
    sinal = "-" if valor_float < 0 else "+"
    centavos = abs(round(valor_float * 100))
    return sinal, str(centavos).zfill(tam)[-tam:]


def _pct(pct_str: str, tam: int = 5) -> str:
    """
    Converte '2,45%' para inteiro com 2 decimais implícitas.
    Ex: '2,45%' → '00245'  (conforme exemplo real CIELO04D)
    """
    s = str(pct_str).strip().rstrip("%").strip()
    s = s.replace(",", ".").replace(" ", "")
    if not s:
        return "0" * tam
    try:
        v = float(s) * 100          # 2 casas decimais implícitas
        return str(int(round(v))).zfill(tam)[-tam:]
    except ValueError:
        return "0" * tam


def _dt_para_aaaammdd(dt_str: str) -> str:
    """
    Converte data DD/MM/AAAA → AAAAMMDD.
    Retorna '00000000' se inválida.
    """
    s = str(dt_str).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y%m%d")
        except ValueError:
            pass
    return "00000000"


def _dt_para_ddmmaaaa(dt_str: str) -> str:
    """
    Converte data DD/MM/AAAA → DDMMAAAA.
    Retorna '00000000' se inválida.
    """
    s = str(dt_str).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d%m%Y")
        except ValueError:
            pass
    return "00000000"


def _parse_valor(s: str) -> float:
    """Converte '2.500,00', '-61,25', '-1.290,00' para float preservando o sinal."""
    s = str(s).strip().strip('"').strip("'").replace(" ", "")
    # Preservar sinal antes de remover separadores
    negativo = s.startswith("-")
    s = s.lstrip("+-")
    s = s.replace(".", "").replace(",", ".")
    try:
        v = float(s)
        return -v if negativo else v
    except ValueError:
        return 0.0


def _lookup(tabela: dict, chave: str, padrao: str = "") -> str:
    """Busca insensível a maiúsculas/minúsculas em tabelas de lookup."""
    chave_norm = str(chave).strip()
    # Tentativa direta
    if chave_norm in tabela:
        return tabela[chave_norm]
    # Case insensitive
    for k, v in tabela.items():
        if isinstance(k, str) and k.lower() == chave_norm.lower():
            return v
    return padrao


def _forma_pagamento(bandeira_cod: str, forma_csv: str) -> str:
    """
    Determina o código da forma de pagamento a partir da bandeira e
    da descrição textual do CSV (ex: 'Crédito parcelado loja 10/10').
    """
    # Remover "XX/YY" do final
    forma_base = re.sub(r"\s+\d{1,2}/\d{1,2}$", "", forma_csv.strip()).strip().lower()

    # Tentativa com chave (bandeira, base)
    chave = (bandeira_cod, forma_base)
    if chave in FORMA_PAGAMENTO:
        return FORMA_PAGAMENTO[chave]

    # Tentativa apenas pela base (sem bandeira)
    for (_, base), cod in FORMA_PAGAMENTO.items():
        if base == forma_base:
            return cod

    # Fallback por palavras-chave
    base_l = forma_base.lower()
    if "parcelado loja" in base_l or "parcelado" in base_l:
        return {"001": "043", "002": "012", "003": "083",
                "007": "072", "009": "165"}.get(bandeira_cod, "043")
    if "crédito à vista" in base_l or "crédito a vista" in base_l:
        return {"001": "040", "002": "010", "003": "082",
                "007": "070", "009": "164"}.get(bandeira_cod, "040")
    if "débito" in base_l:
        return {"001": "041", "002": "011", "007": "071"}.get(bandeira_cod, "041")
    return "000"


def _tipo_liquidacao_from_forma(forma_csv: str) -> str:
    """Infere tipo de liquidação a partir da forma de pagamento."""
    f = forma_csv.lower()
    if "débito" in f or "electron" in f:
        return "001"
    if "voucher" in f or "vale" in f:
        return "004"
    return "002"  # crédito (padrão)


def _cpf_cnpj_da_chave(chave_ur: str, fallback: str = "") -> str:
    """
    Extrai o CPF/CNPJ do recebedor da chave UR.
    Estrutura da chave (89 chars úteis):
      [0:14]  cpf_cnpj_UR
      [14:28] cpf_cnpj_titular
      [28:38] data_vencimento (AAAA-MM-DD)
      [38:40] tipo_lancamento
      [40:42] tipo_liquidacao
      [42:46] bandeira+tipo
      [46:56] estabelecimento_submissor
      [56:70] cpf_cnpj_recebedor  ← este campo
      [70:89] zeros
    """
    chave = str(chave_ur).strip()
    if len(chave) >= 70:
        cpf = chave[56:70].strip()
        if cpf and cpf != "00000000000000":
            return cpf
    return fallback


def _conta_formatada(conta_str: str) -> tuple[str, str]:
    """
    Separa conta '57559-4' → conta='00000000000000057559', digito='4'.
    Se não há hífen, usa o último char como dígito.
    """
    s = str(conta_str).strip()
    if "-" in s:
        partes = s.rsplit("-", 1)
        conta  = partes[0].replace(".", "").replace("-", "").zfill(20)[-20:]
        digito = partes[1][:1]
    else:
        conta  = s.replace(".", "").zfill(21)[:20]
        digito = s[-1] if s else "0"
    return conta, digito


def _agencia_formatada(ag_str: str) -> str:
    """
    Formata agência para 5 chars no padrão EDI Cielo:
    4 dígitos da agência (zerofill) + '0' como dígito verificador padrão.
    Ex: '656' → '0656' + '0' = '06560'
    """
    s = str(ag_str).strip().replace("-", "").replace(" ", "")
    # Se já tem 5 chars, usar direto
    if len(s) == 5 and s.isdigit():
        return s
    # Se tem 4 chars, adicionar dígito '0'
    if len(s) == 4 and s.isdigit():
        return s + "0"
    # Caso geral: zerofill para 4 + dígito '0'
    return s.zfill(4)[:4] + "0"


# ============================================================================
# LEITURA DO CSV
# ============================================================================

def ler_csv(caminho: str) -> tuple[list[dict], dict]:
    """
    Lê o CSV de Recebíveis do portal Cielo.

    Retorna:
        linhas    : lista de dicts com os dados
        metadados : dict com estabelecimento, data_proc, etc.
    """
    caminho = str(caminho)
    metadados: dict = {}
    header_idx: int = -1

    with open(caminho, encoding="latin-1", errors="replace") as f:
        raw = f.readlines()

    # Extrair metadados das primeiras linhas (antes do header de colunas)
    for i, line in enumerate(raw):
        s = line.strip()
        if s.startswith("Estabelecimento:"):
            metadados["estabelecimento"] = s.split(":", 1)[1].strip().split(";")[0].strip()
        if s.startswith("Data de pagamento:"):
            datas = s.split(":", 1)[1].strip().split(";")[0].strip()
            partes = [p.strip() for p in datas.split("à")]
            if len(partes) == 2:
                metadados["data_ini"] = _dt_para_aaaammdd(partes[0])
                metadados["data_fim"] = _dt_para_aaaammdd(partes[1])
        if s.startswith("Usuário:"):
            metadados["usuario"] = s.split(":", 1)[1].strip().split(";")[0].strip()
        if s.startswith("CPF/CNPJ:"):
            metadados["cpf_cnpj"] = (
                s.split(":", 1)[1].strip().split(";")[0]
                .replace(".", "").replace("/", "").replace("-", "").strip()
            )
        if s.startswith("Data de pagamento;"):
            header_idx = i
            break

    if header_idx == -1:
        raise ValueError("Header 'Data de pagamento;...' não encontrado no CSV.")

    # Leitura com csv.DictReader
    linhas = []
    reader = csv.DictReader(
        (line for line in raw[header_idx:]),
        delimiter=";",
    )
    for row in reader:
        # Limpar aspas extras e espaços em branco
        limpo = {k: (v.strip().strip('"').strip() if v else "") for k, v in row.items()}
        # Ignorar linhas sem dados úteis
        if not limpo.get("Estabelecimento"):
            continue
        linhas.append(limpo)

    # Se metadados não extraídos das linhas de cabeçalho, inferir das linhas
    if not metadados.get("estabelecimento") and linhas:
        metadados["estabelecimento"] = linhas[0].get("Estabelecimento", "")
    if not metadados.get("data_ini") and linhas:
        datas = sorted(
            _dt_para_aaaammdd(l["Data de pagamento"])
            for l in linhas if l.get("Data de pagamento")
        )
        if datas:
            metadados["data_ini"] = datas[0]
            metadados["data_fim"] = datas[-1]

    metadados.setdefault("data_proc", datetime.now().strftime("%Y%m%d"))
    metadados.setdefault("data_ini",  metadados["data_proc"])
    metadados.setdefault("data_fim",  metadados["data_proc"])
    metadados.setdefault("estabelecimento", "0000000000")
    metadados.setdefault("cpf_cnpj", "00000000000000")

    return linhas, metadados


# ============================================================================
# GERAÇÃO DOS REGISTROS
# ============================================================================

def gerar_registro_0(meta: dict, sequencia: int = 1) -> str:
    """
    Registro 0 – Header (250 posições).
    """
    estab    = _estab(meta["estabelecimento"], 10)
    dt_proc  = meta["data_proc"]          # AAAAMMDD
    dt_ini   = meta["data_ini"]           # AAAAMMDD
    dt_fim   = meta["data_fim"]           # AAAAMMDD
    seq      = _n(sequencia, 7)
    empresa  = _a("CIELO", 5)
    opcao    = "04"                       # CIELO04 = Liquidação/Pagamento
    trans    = "I"
    caixa    = _a("", 20)
    versao   = "015"
    hierarq  = "03"
    ind_cad  = "S"
    uso      = _a("", 174)               # Uso Cielo (brancos)

    linha = (
        "0"       +  # pos  1
        estab     +  # pos  2-11
        dt_proc   +  # pos 12-19
        dt_ini    +  # pos 20-27
        dt_fim    +  # pos 28-35
        seq       +  # pos 36-42
        empresa   +  # pos 43-47
        opcao     +  # pos 48-49
        trans     +  # pos 50
        caixa     +  # pos 51-70
        versao    +  # pos 71-73
        hierarq   +  # pos 74-75
        ind_cad   +  # pos 76
        uso          # pos 77-250
    )
    return linha[:250].ljust(250)


def gerar_registro_d(grupo: list[dict], meta: dict) -> str:
    """
    Registro D – UR Agenda (254+ posições).
    Um registro D por grupo de linhas com o mesmo Código da Unidade de Recebível.
    Os valores são somados de todas as linhas do grupo.
    Suporta múltiplos estabelecimentos: cpf_cnpj é extraído da chave_ur da linha.
    """
    ref = grupo[0]  # linha de referência para campos fixos do grupo

    estab_linha  = str(ref.get("Estabelecimento", meta["estabelecimento"])).strip()
    estab        = _estab(estab_linha, 10)

    # CPF/CNPJ: extrair da chave_ur (contém o cpf do recebedor deste estabelecimento)
    chave_ref    = ref.get("Código da Unidade de recebível", "")
    cpf_raw      = _cpf_cnpj_da_chave(chave_ref, meta["cpf_cnpj"])
    cpf_cnpj     = _a(cpf_raw.zfill(14), 14)

    bandeira_txt = ref.get("Bandeira", "")
    bandeira_cod = _lookup(BANDEIRA, bandeira_txt, "000")

    forma_csv    = ref.get("Forma de pagamento", "")
    tipo_liq     = _tipo_liquidacao_from_forma(forma_csv)
    matriz_pag   = _estab(estab_linha, 10)   # matriz = o próprio estab da linha

    status_txt   = ref.get("Status de pagamento", "Pago")
    status_cod   = _lookup(STATUS_PAGAMENTO, status_txt, "03")

    # Somar valores monetários de todas as linhas do grupo
    soma_bruto = sum(_parse_valor(l.get("Valor bruto", "0")) for l in grupo)
    soma_taxa  = sum(_parse_valor(l.get("Valor Taxa/Tarifa", "0")) for l in grupo)
    soma_liq   = sum(_parse_valor(l.get("Valor líquido", "0")) for l in grupo)

    sinal_vb, vb = _mon(soma_bruto)
    sinal_tx, tx = _mon(soma_taxa)
    sinal_vl, vl = _mon(soma_liq)

    banco_raw, agencia_raw, conta_raw = (
        ref.get("Banco", "0000"),
        ref.get("Agência", "00000"),
        ref.get("Conta", ""),
    )
    banco    = _n(banco_raw, 4)
    agencia  = _agencia_formatada(agencia_raw)
    conta, digito = _conta_formatada(conta_raw)

    qtd_lanc = _n(len(grupo), 6)

    tipo_lanc_txt = ref.get("Tipo de lançamento", "Venda parcelada")
    tipo_lanc_cod = _lookup(TIPO_LANCAMENTO, tipo_lanc_txt, "03")
    if tipo_lanc_cod == "03" and "parcelado" not in forma_csv.lower():
        tipo_lanc_cod = "01"

    chave_ur = _a(ref.get("Código da Unidade de recebível", ""), 100)
    tipo_lanc_orig = tipo_lanc_cod

    # Campos adicionais pos 255+
    dt_pag   = _dt_para_ddmmaaaa(ref.get("Data de pagamento", ""))
    dt_pag2  = dt_pag
    dt_pag3  = dt_pag
    recebedor_estab = _estab(estab_linha, 10)
    flags    = "NNN"
    cpf_neg  = _a("0" * 14, 14)

    linha = (
        "D"             +  # pos  1
        estab           +  # pos  2-11
        cpf_cnpj        +  # pos 12-25
        cpf_cnpj        +  # pos 26-39
        cpf_cnpj        +  # pos 40-53
        bandeira_cod    +  # pos 54-56
        tipo_liq        +  # pos 57-59
        matriz_pag      +  # pos 60-69
        status_cod      +  # pos 70-71
        sinal_vb        +  # pos 72
        vb              +  # pos 73-85
        sinal_tx        +  # pos 86
        tx              +  # pos 87-99
        sinal_vl        +  # pos 100
        vl              +  # pos 101-113
        banco           +  # pos 114-117
        agencia         +  # pos 118-122
        conta           +  # pos 123-142
        digito          +  # pos 143
        qtd_lanc        +  # pos 144-149
        tipo_lanc_cod   +  # pos 150-151
        chave_ur        +  # pos 152-251
        tipo_lanc_orig  +  # pos 252-253
        "0"             +  # pos 254  (indicativo_reenvio = Normal)
        # pos 255+ (campos adicionais)
        "0" * 14        +  # pos 255-268 (saldo negociação zeros)
        dt_pag          +  # pos 269-276 (data vencimento DDMMAAAA)
        "01011001"      +  # pos 277-284 (data envio banco = sem envio)
        dt_pag3         +  # pos 285-292 (data vencimento original)
        recebedor_estab +  # pos 293-302
        flags           +  # pos 303-305
        cpf_neg            # pos 306-319
    )
    return linha


def gerar_registro_e(row: dict, meta: dict) -> str:
    """
    Registro E – Detalhe do Lançamento (760 posições).
    Uma linha do CSV → um Registro E.
    Suporta múltiplos estabelecimentos: cpf_cnpj extraído da chave_ur da linha.
    """
    estab_linha  = str(row.get("Estabelecimento", meta["estabelecimento"])).strip()
    estab        = _estab(estab_linha, 10)

    # CPF/CNPJ do recebedor: extrair da chave_ur desta linha
    chave_linha  = row.get("Código da Unidade de recebível", "")
    cpf_linha    = _cpf_cnpj_da_chave(chave_linha, meta["cpf_cnpj"])

    bandeira_txt = row.get("Bandeira", "")
    bandeira_cod = _lookup(BANDEIRA, bandeira_txt, "000")

    forma_csv    = row.get("Forma de pagamento", "")
    tipo_liq     = _tipo_liquidacao_from_forma(forma_csv)

    # Parcela e total
    try:
        parcela = int(row.get("Número da parcela") or 0)
    except (ValueError, TypeError):
        parcela = 0
    try:
        total_parc = int(row.get("Quantidade total de parcelas") or 0)
    except (ValueError, TypeError):
        total_parc = 0

    # Lançamento parcelado → parcela/total, senão zeros
    tipo_lanc_txt  = row.get("Tipo de lançamento", "Venda parcelada")
    tipo_lanc_cod  = _lookup(TIPO_LANCAMENTO, tipo_lanc_txt, "03")
    if tipo_lanc_cod == "03" and "parcelado" not in forma_csv.lower():
        tipo_lanc_cod = "01"
    is_parcelado = tipo_lanc_cod in ("03",)
    parc_str     = _n(parcela if is_parcelado else 0, 2)
    total_str    = _n(total_parc if is_parcelado else 0, 2)

    cod_auth  = _a(row.get("Código de autorização", "").strip(), 6)
    chave_ur  = _a(row.get("Código da Unidade de recebível", ""), 100)

    # Código da transação recebida = Código da venda
    cod_transac = _a(row.get("Código da venda", ""), 22)

    # cod_ajuste: inferido da Descrição do CSV (quando disponível)
    descricao   = row.get("Descrição", "").strip()
    cod_ajuste_val = _lookup(DESCRICAO_AJUSTE, descricao, "")
    # Ajustes e cancelamentos (tipo_lanc 04-10 exceto 03) sempre têm cod_ajuste
    if not cod_ajuste_val and tipo_lanc_cod not in ("01", "02", "03"):
        cod_ajuste_val = ""  # sem info suficiente → branco
    cod_ajuste  = _a(cod_ajuste_val, 4)

    forma_cod = _forma_pagamento(bandeira_cod, forma_csv)

    # Indicativos (N por padrão)
    ind_promo  = "N"
    ind_dcc    = "N"
    ind_commin = "N"
    ind_ra_tc  = "3"  # Sem produtos de prazo
    ind_tzero  = "N"
    ind_rejeit = "N"
    ind_tardia = "N"

    # Cartão: '498453 **** 6761' → bin=498453, ultimos=6761
    cartao = row.get("Número do cartão", "")
    bin_cartao = ""
    ult_digitos = ""
    m = re.match(r"(\d{6})\s*\*+\s*(\d{4})", cartao)
    if m:
        bin_cartao  = m.group(1)
        ult_digitos = m.group(2)

    nsu       = _a(row.get("NSU/DOC", "").zfill(6)[-6:], 6)
    nota_fisc = _n(row.get("Nota fiscal", "0") or "0", 10)
    tid       = _a(row.get("TID", ""), 20)
    pedido    = _a(row.get("Número do pedido", ""), 20)

    # Taxas
    taxa_mdr  = _pct(row.get("Taxa administrativa (MDR)", "0"))
    taxa_ra   = _pct(row.get("Taxa do prazo de recebimento", "0"))
    taxa_vnd  = taxa_mdr  # taxa_venda = MDR

    # Valores
    v_total   = _parse_valor(row.get("Valor total da transação", "0") or "0")
    v_bruto   = _parse_valor(row.get("Valor bruto", "0"))
    v_liq     = _parse_valor(row.get("Valor líquido", "0"))
    v_taxa    = _parse_valor(row.get("Valor Taxa/Tarifa", "0"))
    v_mdr     = _parse_valor(row.get("Valor da taxa administrativa (MDR)", "0") or "0")
    v_ra      = _parse_valor(row.get("Valor da taxa do prazo de recebimento", "0") or "0")
    v_saque   = _parse_valor(row.get("Valor do saque", "0") or "0")

    # Comissão = taxa administrativa (MDR)
    # Se v_mdr não disponível, usa v_taxa
    if v_mdr == 0 and v_taxa != 0:
        v_mdr = v_taxa

    sinal_vt, val_vt = _mon(v_total)
    sinal_vb, val_vb = _mon(v_bruto)
    sinal_vl, val_vl = _mon(v_liq)
    sinal_cm, val_cm = _mon(v_mdr)
    sinal_ra, val_ra = _mon(v_ra)
    sinal_sq, val_sq = _mon(v_saque)

    # Zeros para campos não usados neste fluxo
    zero13 = "+" + "0" * 13
    zero_sinal_val = lambda: "+" + "0" * 13

    hora_raw = row.get("Hora da venda", "").strip()
    # CSV tem 'HH:MM', EDI precisa de 'HHMMSS' - completar segundos com '00'
    hora = re.sub(r"[:\s]", "", hora_raw).ljust(6, "0")[:6]

    # Grupo de cartões: Amex sempre '00'; outros '01' se emitido no Brasil
    if bandeira_cod == "003":
        grupo = "00"
    elif row.get("Origem do cartão", "").strip():
        grupo = "01"
    else:
        grupo = "00"

    # CPF/CNPJ recebedor: usa o extraído da chave_ur (correto por estabelecimento)
    cpf_receb = _a(cpf_linha.zfill(14), 14)

    # cod_unico e cod_original são chaves internas Cielo, não disponíveis no CSV
    cod_unico = _a("", 15)
    cod_orig  = _a("", 15)
    id_efeito = _a("", 15)

    canal_txt  = row.get("Canal da venda", "Máquina")
    canal_cod  = _lookup(CANAL_VENDA, canal_txt, "001")

    terminal   = _a(row.get("Número da máquina", "").strip(), 8)
    tipo_lanc_orig = tipo_lanc_cod

    # tipo_transacao: só preenchido para vendas (01/02/03); zeros para ajustes/outros
    is_venda = tipo_lanc_cod in ("01", "02", "03")
    if is_venda:
        if "débito" in forma_csv.lower():
            tipo_trans = "001"
        elif "parcelado" in forma_csv.lower():
            tipo_trans = "003"
        else:
            tipo_trans = "002"
    else:
        tipo_trans = "000"  # cancelamentos, negociações, ajustes

    # Matriz de pagamento: usar o estabelecimento da linha (pode ser filial)
    matriz_pag  = _estab(estab_linha, 10)

    # Datas: para cancelamentos, dt_capt e dt_lanc = Data do lançamento (quando ocorreu)
    # Para vendas normais, dt_capt = dt_lanc = Data do lançamento (ou Data da venda)
    dt_auth  = _dt_para_ddmmaaaa(row.get("Data da venda", ""))
    dt_capt  = _dt_para_ddmmaaaa(row.get("Data do lançamento", ""))
    dt_lanc  = _dt_para_ddmmaaaa(row.get("Data do lançamento", ""))
    dt_lanc0 = _dt_para_ddmmaaaa(row.get("Data da venda", ""))

    lote_raw = str(row.get("Número do lote", "0") or "0").strip()
    # CSV tem '004250516' (9 chars); EDI usa 7 chars sem zeros extras à esq
    lote_limpo = lote_raw.lstrip("0") or "0"
    lote     = lote_limpo.zfill(7)[-7:]

    num_trans_proc = _a("", 22)
    motivo_rej     = _a("", 3)

    dt_venc_orig = _dt_para_ddmmaaaa(row.get("Data de pagamento", ""))
    tipo_cartao  = "00"

    orig_cartao = _lookup(ORIGEM_CARTAO, row.get("Origem do cartão", ""), "N")
    ind_mdr_tc  = "N"
    ind_parc_cl = "N"

    uso_cielo_557 = "0" * 4   # pos 557-560: uso futuro Cielo
    cod_precif    = _a("", 5)  # pos 561-565: modelo precificação (não disponível no CSV)

    banco_raw   = row.get("Banco", "0000")
    agencia_raw = row.get("Agência", "00000")
    conta_raw   = row.get("Conta", "")
    banco    = _n(banco_raw, 4)
    agencia  = _agencia_formatada(agencia_raw)
    conta, digito = _conta_formatada(conta_raw)

    arn      = _a("", 23)
    ind_neg  = "N"

    captura_txt = row.get("Tipo de captura", "")
    captura_cod = _lookup(TIPO_CAPTURA, captura_txt, "05")

    cpf_neg = _a("0" * 14, 14)
    uso_fin = _a("", 38)   # uso Cielo pos 723-760

    linha = (
        "E"            +  # pos   1
        estab          +  # pos   2-11
        bandeira_cod   +  # pos  12-14
        tipo_liq       +  # pos  15-17
        parc_str       +  # pos  18-19
        total_str      +  # pos  20-21
        cod_auth       +  # pos  22-27
        tipo_lanc_cod  +  # pos  28-29
        chave_ur       +  # pos  30-129
        cod_transac    +  # pos 130-151
        cod_ajuste     +  # pos 152-155
        forma_cod      +  # pos 156-158
        ind_promo      +  # pos 159
        ind_dcc        +  # pos 160
        ind_commin     +  # pos 161
        ind_ra_tc      +  # pos 162
        ind_tzero      +  # pos 163
        ind_rejeit     +  # pos 164
        ind_tardia     +  # pos 165
        _a(bin_cartao, 6)  +  # pos 166-171
        _a(ult_digitos, 4) +  # pos 172-175
        nsu            +  # pos 176-181
        nota_fisc      +  # pos 182-191
        tid            +  # pos 192-211
        pedido         +  # pos 212-231
        taxa_mdr       +  # pos 232-236
        taxa_ra        +  # pos 237-241
        taxa_vnd       +  # pos 242-246
        sinal_vt + val_vt +  # pos 247-260
        sinal_vb + val_vb +  # pos 261-274
        sinal_vl + val_vl +  # pos 275-288
        sinal_cm + val_cm +  # pos 289-302  (comissão = MDR)
        zero_sinal_val() +   # pos 303-316  (comissão mínima)
        zero_sinal_val() +   # pos 317-330  (entrada Cias. Aéreas)
        zero_sinal_val() +   # pos 331-344  (tarifa MDR detalhe)
        sinal_ra + val_ra +  # pos 345-358  (recebimento automático)
        sinal_sq + val_sq +  # pos 359-372  (saque)
        zero_sinal_val() +   # pos 373-386  (tarifa embarque)
        zero_sinal_val() +   # pos 387-400  (valor pendente)
        zero_sinal_val() +   # pos 401-414  (valor dívida)
        zero_sinal_val() +   # pos 415-428  (valor cobrado)
        zero_sinal_val() +   # pos 429-442  (tarifa administrativa = ZERO no E; consta no D)
        zero_sinal_val() +   # pos 443-456  (Cielo Promo)
        zero_sinal_val() +   # pos 457-470  (DCC conversor)
        hora           +  # pos 471-476
        grupo          +  # pos 477-478
        cpf_receb      +  # pos 479-492
        bandeira_cod   +  # pos 493-495  (bandeira autorização = liquidação)
        cod_unico      +  # pos 496-510
        cod_orig       +  # pos 511-525
        id_efeito      +  # pos 526-540
        canal_cod      +  # pos 541-543
        terminal       +  # pos 544-551
        tipo_lanc_orig +  # pos 552-553
        tipo_trans     +  # pos 554-556
        uso_cielo_557  +  # pos 557-560
        cod_precif     +  # pos 561-565
        dt_auth        +  # pos 566-573
        dt_capt        +  # pos 574-581
        dt_lanc        +  # pos 582-589
        dt_lanc0       +  # pos 590-597
        lote           +  # pos 598-604
        num_trans_proc +  # pos 605-626
        motivo_rej     +  # pos 627-629
        dt_venc_orig   +  # pos 630-637
        matriz_pag     +  # pos 638-647
        tipo_cartao    +  # pos 648-649
        orig_cartao    +  # pos 650
        ind_mdr_tc     +  # pos 651
        ind_parc_cl    +  # pos 652
        banco          +  # pos 653-656
        agencia        +  # pos 657-661
        conta          +  # pos 662-681
        digito         +  # pos 682
        arn            +  # pos 683-705
        ind_neg        +  # pos 706
        captura_cod    +  # pos 707-708
        cpf_neg        +  # pos 709-722
        uso_fin           # pos 723-760
    )

    # Garantir exatamente 760 posições
    linha = linha[:760].ljust(760)
    return linha


def gerar_registro_9(
    total_regs: int,
    val_liq: float,
    val_bruto: float,
    qtd_e: int,
    val_cedido: float = 0.0,
    val_gravame: float = 0.0,
) -> str:
    """
    Registro 9 – Trailer (95 posições).

    total_regs : total de registros D + E (sem header e trailer)
    val_liq    : soma dos valores líquidos dos registros D
    val_bruto  : soma dos valores brutos dos registros D
    qtd_e      : quantidade total de registros E
    """
    sinal_vl, vl = _mon(val_liq, 17)
    sinal_vb, vb = _mon(val_bruto, 17)
    sinal_ced, ced = _mon(val_cedido, 17)
    sinal_grav, grav = _mon(val_gravame, 17)

    linha = (
        "9"                    +  # pos  1
        _n(total_regs, 11)     +  # pos  2-12
        sinal_vl + vl          +  # pos 13-30
        _n(qtd_e, 11)          +  # pos 31-41
        sinal_vb + vb          +  # pos 42-59
        sinal_ced + ced        +  # pos 60-77
        sinal_grav + grav         # pos 78-95
    )
    return linha[:95].ljust(95)


# ============================================================================
# PIPELINE PRINCIPAL
# ============================================================================

def converter(caminho_csv: str, pasta_saida: str) -> Path:
    """
    Pipeline completo: CSV → arquivo EDI CIELO04D.

    Filtragem: somente linhas com Valor bruto > 0 (créditos/recebíveis).
    Lançamentos negativos (cancelamentos, chargebacks) são ignorados —
    tratativa financeira feita pelo analista.

    Retorna o caminho do arquivo gerado.
    """
    pasta_saida = Path(pasta_saida)
    pasta_saida.mkdir(parents=True, exist_ok=True)

    print(f"\n  » Lendo CSV: {caminho_csv}")
    linhas, meta = ler_csv(caminho_csv)
    total_lido = len(linhas)
    print(f"    {total_lido} linhas lidas")
    print(f"    Estabelecimento : {meta['estabelecimento']}")
    print(f"    Período         : {meta['data_ini']} → {meta['data_fim']}")

    if not linhas:
        raise ValueError("CSV sem linhas de dados.")

    # ------------------------------------------------------------------
    # FILTRAGEM: somente lançamentos de crédito (Valor bruto > 0)
    # Cancelamentos, chargebacks e estornos ficam a cargo do analista.
    # ------------------------------------------------------------------
    linhas_credito = [
        r for r in linhas
        if _parse_valor(r.get("Valor bruto", "0")) > 0
    ]
    ignoradas = total_lido - len(linhas_credito)
    if ignoradas:
        tipos_ignorados: dict[str, int] = {}
        for r in linhas:
            if _parse_valor(r.get("Valor bruto", "0")) <= 0:
                t = r.get("Tipo de lançamento", "Desconhecido").strip()
                tipos_ignorados[t] = tipos_ignorados.get(t, 0) + 1
        print(f"    Filtradas (débito): {ignoradas} linha(s) ignorada(s):")
        for tipo, cnt in sorted(tipos_ignorados.items()):
            print(f"      • {tipo}: {cnt}")

    print(f"    Linhas a converter : {len(linhas_credito)}")

    if not linhas_credito:
        raise ValueError("Nenhuma linha de crédito encontrada após filtragem.")

    # ------------------------------------------------------------------
    # Atualizar datas do meta com base nas linhas filtradas
    # (a data_ini/fim pode mudar se as linhas removidas eram as extremas)
    # ------------------------------------------------------------------
    datas_pag = sorted(
        _dt_para_aaaammdd(r["Data de pagamento"])
        for r in linhas_credito
        if r.get("Data de pagamento")
    )
    if datas_pag:
        meta["data_ini"] = datas_pag[0]
        meta["data_fim"] = datas_pag[-1]

    # ------------------------------------------------------------------
    # Agrupar por Código da Unidade de Recebível (chave D)
    # ------------------------------------------------------------------
    grupos: dict[str, list[dict]] = defaultdict(list)
    for row in linhas_credito:
        chave = row.get("Código da Unidade de recebível", "").strip()
        grupos[chave].append(row)
    print(f"    Grupos (Reg D)  : {len(grupos)}")
    print(f"    Registros E     : {sum(len(g) for g in grupos.values())}")

    # ------------------------------------------------------------------
    # Gerar linhas do arquivo com tratativa de exceção por linha
    # ------------------------------------------------------------------
    linhas_edi:   list[str] = []
    erros:        list[str] = []   # log de erros por linha
    soma_vl       = 0.0
    soma_vb       = 0.0
    qtd_e_total   = 0
    qtd_e_erro    = 0

    # Registro 0
    linhas_edi.append(gerar_registro_0(meta))

    # Número de linha global no CSV original para rastreio
    # Montar índice: (chave_ur, auth) → número da linha no CSV original
    num_linha_csv: dict[tuple, int] = {}
    for idx, row in enumerate(linhas, start=1):
        chave = row.get("Código da Unidade de recebível", "")
        auth  = row.get("Código de autorização", "")
        num_linha_csv[(chave, auth)] = idx

    for chave, grupo in grupos.items():
        # --- Registro D ---
        try:
            linhas_edi.append(gerar_registro_d(grupo, meta))
        except Exception as exc:
            ref  = grupo[0]
            auth = ref.get("Código de autorização", "—")
            estab = ref.get("Estabelecimento", "—")
            msg = (
                f"[Reg D] Estab {estab} | Auth {auth} | "
                f"Chave UR {chave[:30]}... | Erro: {exc}"
            )
            erros.append(msg)
            continue   # pula o grupo inteiro se o D falhar

        # Somar valores do D apenas se gerou com sucesso
        soma_vl += sum(_parse_valor(r.get("Valor líquido", "0")) for r in grupo)
        soma_vb += sum(_parse_valor(r.get("Valor bruto",   "0")) for r in grupo)

        # --- Registros E ---
        for row in grupo:
            auth      = row.get("Código de autorização", "—").strip()
            nsu       = row.get("NSU/DOC", "—").strip()
            estab_row = row.get("Estabelecimento", "—").strip()
            chave_row = row.get("Código da Unidade de recebível", "")
            num_linha = num_linha_csv.get((chave_row, auth), "?")
            try:
                linhas_edi.append(gerar_registro_e(row, meta))
                qtd_e_total += 1
            except Exception as exc:
                qtd_e_erro += 1
                msg = (
                    f"[Reg E] Linha CSV {num_linha} | "
                    f"Estab {estab_row} | "
                    f"Auth {auth} | NSU {nsu} | "
                    f"Erro: {exc}"
                )
                erros.append(msg)

    # total_regs = D + E bem-sucedidos (sem header e trailer)
    qtd_d_ok   = sum(1 for l in linhas_edi if l.startswith("D"))
    total_regs = qtd_d_ok + qtd_e_total

    # Registro 9
    linhas_edi.append(gerar_registro_9(total_regs, soma_vl, soma_vb, qtd_e_total))

    # ------------------------------------------------------------------
    # Nome do arquivo: CIELO04D_<estab>_<data_ini_DDMMAAAA>_<data_fim_DDMMAAAA>.txt
    # Ex: CIELO04D_1028105247_20032026_31032026.txt
    # ------------------------------------------------------------------
    estab   = meta["estabelecimento"]
    # Converter datas de AAAAMMDD para DDMMAAAA para o nome do arquivo
    def _aaaammdd_para_ddmmaaaa(s: str) -> str:
        if len(s) == 8 and s.isdigit():
            return s[6:8] + s[4:6] + s[0:4]
        return s
    dt_ini_fmt = _aaaammdd_para_ddmmaaaa(meta["data_ini"])
    dt_fim_fmt = _aaaammdd_para_ddmmaaaa(meta["data_fim"])
    nome_arq      = f"CIELO04D_{estab}_{dt_ini_fmt}_{dt_fim_fmt}.txt"
    caminho_saida = pasta_saida / nome_arq

    # --- Gravar arquivo (encoding latin-1) ---
    with open(caminho_saida, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(linhas_edi) + "\n")

    # ------------------------------------------------------------------
    # Relatório de conversão
    # ------------------------------------------------------------------
    SEP_R = "-" * 50
    print(f"\n  ✅  Arquivo gerado: {caminho_saida}")
    print(f"  {SEP_R}")
    print(f"     Total lidas           : {total_lido}")
    print(f"     Filtradas (débito)    : {ignoradas}")
    print(f"     Convertidas (crédito) : {len(linhas_credito)}")
    print(f"  {SEP_R}")
    print(f"     Registros 0  : 1")
    print(f"     Registros D  : {qtd_d_ok}")
    print(f"     Registros E  : {qtd_e_total}")
    if qtd_e_erro:
        print(f"     Reg E c/ erro: {qtd_e_erro}  ⚠️")
    print(f"     Registros 9  : 1")
    print(f"     Total linhas : {len(linhas_edi)}")
    print(f"  {SEP_R}")
    print(f"     Período      : {dt_ini_fmt} → {dt_fim_fmt}")
    print(f"     Valor bruto  : R$ {soma_vb:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    print(f"     Valor líquido: R$ {soma_vl:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    if erros:
        print(f"\n  ⚠️   {len(erros)} erro(s) durante a conversão:")
        print(f"  {SEP_R}")
        for i, e in enumerate(erros, 1):
            print(f"     [{i:02d}] {e}")
        print(f"  {SEP_R}")
        print("  Corrija as linhas acima no CSV e reconverta.")

    return caminho_saida


# ============================================================================
# GERENCIAMENTO DE PASTAS
# ============================================================================

def preparar_pastas(base: Path) -> tuple[Path, Path]:
    pasta_input  = base / "input"
    pasta_output = base / "output"
    pasta_input.mkdir(parents=True, exist_ok=True)
    pasta_output.mkdir(parents=True, exist_ok=True)
    return pasta_input, pasta_output


def listar_csvs(pasta: Path) -> list[Path]:
    arquivos = []
    for ext in ("*.csv", "*.CSV"):
        arquivos.extend(pasta.glob(ext))
    return sorted(set(arquivos))


# ============================================================================
# MAIN
# ============================================================================

SEP  = "=" * 62
SEP2 = "-" * 62


def main():
    print(f"\n{SEP}")
    print("  Conversor CSV Recebíveis Cielo → EDI CIELO04D")
    print("  Entrada → input/   |   Saída → output/")
    print(SEP)

    base = Path.cwd()
    pasta_input, pasta_output = preparar_pastas(base)

    # --- Modo direto via argumentos ---
    if len(sys.argv) >= 2:
        arquivo_csv = Path(sys.argv[1])
        if not arquivo_csv.exists():
            print(f"\n  ❌  Arquivo não encontrado: {arquivo_csv}")
            sys.exit(1)
        pasta_sai = Path(sys.argv[2]) if len(sys.argv) >= 3 else pasta_output
        # Copiar para input/ se não estiver lá
        destino = pasta_input / arquivo_csv.name
        if arquivo_csv.resolve() != destino.resolve():
            shutil.copy2(arquivo_csv, destino)
            print(f"\n  📥  Copiado para input/: {arquivo_csv.name}")
        try:
            converter(str(destino), str(pasta_sai))
        except Exception as e:
            print(f"\n  ❌  Erro: {e}")
            sys.exit(1)
        return

    # --- Modo interativo ---
    print(f"\n  Pasta de entrada : {pasta_input}")
    print(f"  Pasta de saída   : {pasta_output}\n")

    while True:
        print(SEP2)
        print("  [1] Informar caminho do arquivo CSV")
        print("  [2] Escolher arquivo já em input/")
        print("  [3] Converter TODOS os CSV em input/")
        print("  [0] Sair")
        print(SEP2)
        op = input("  Escolha: ").strip()

        if op == "0":
            print("\n  Até mais!\n")
            break

        elif op == "1":
            caminho = input("\n  Caminho do CSV: ").strip().strip('"')
            arq = Path(caminho)
            if not arq.exists():
                print(f"\n  ❌  Não encontrado: {arq}\n")
                continue
            destino = pasta_input / arq.name
            if arq.resolve() != destino.resolve():
                shutil.copy2(arq, destino)
                print(f"  📥  Copiado para input/")
            try:
                converter(str(destino), str(pasta_output))
            except Exception as e:
                print(f"\n  ❌  Erro: {e}\n")

        elif op == "2":
            csvs = listar_csvs(pasta_input)
            if not csvs:
                print(f"\n  ℹ️  Nenhum CSV em {pasta_input}\n")
                continue
            print()
            for i, c in enumerate(csvs, 1):
                print(f"  [{i}] {c.name}")
            print("  [0] Voltar")
            esc = input("\n  Número: ").strip()
            if esc == "0":
                continue
            try:
                arq = csvs[int(esc) - 1]
            except (ValueError, IndexError):
                print("  ❌  Opção inválida.\n")
                continue
            try:
                converter(str(arq), str(pasta_output))
            except Exception as e:
                print(f"\n  ❌  Erro: {e}\n")

        elif op == "3":
            csvs = listar_csvs(pasta_input)
            if not csvs:
                print(f"\n  ℹ️  Nenhum CSV em {pasta_input}\n")
                continue
            for i, arq in enumerate(csvs, 1):
                print(f"\n  [{i}/{len(csvs)}] {arq.name}")
                print(SEP2)
                try:
                    converter(str(arq), str(pasta_output))
                except Exception as e:
                    print(f"  ❌  Erro: {e}")
        else:
            print("  ❌  Opção inválida.\n")


if __name__ == "__main__":
    main()
