#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
 GERADOR DE PLANILHA SILOMS - Resultado de Licitacao
 Fonte: API Dados Abertos Compras.gov.br (Modulo Contratacoes / Lei 14.133)
=============================================================================
 Reproduz o layout de 15 colunas do modelo SILOMS:

   LOTE | ITEM | REQUISICAO | CNPJ | EMPRESA | QTDE | UND |
   VALOR UNIT LIC | VALOR TOTAL LICI | PRAZO | DESCRICAO |
   SITUACAO | FORNECEDOR | MODELO/VERSAO | MARCA

 Os filtros sao os mesmos do site de referencia: UASG, tipo de licitacao e
 numero do processo. O script PERGUNTA os tres a cada execucao.

 INSTALACAO (uma vez so):
     pip install requests pandas openpyxl

 USO:
     Duplo-clique em RODAR_PLANILHA.bat  (ou: py gerar_planilha_siloms.py)
     e responda as perguntas na tela.
=============================================================================
"""

import json
import re
import sys
import time
from datetime import date, datetime, timedelta

import requests
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# =============================================================================
# PADROES  -- o script pergunta UASG, tipo e numero a cada execucao.
#             Os valores abaixo sao apenas as sugestoes que aparecem entre
#             colchetes; basta apertar ENTER para aceitar.
# =============================================================================

UASG_PADRAO = "120641"          # BAPV. Aparece como sugestao na pergunta.
TIPO_PADRAO = "PREGAO"          # Tipo sugerido.

PRAZO_PADRAO = 30               # Valor fixo da coluna PRAZO

# --- Comportamentos do layout (raramente mudam) -----------------------------

# False -> emite a linha "Item deserto/fracassado..." apenas nos itens que
#          realmente ficaram sem vencedor.  [PADRAO - comportamento correto]
# True  -> emite essa linha para TODO item, reproduzindo o defeito do arquivo
#          modelo (2 linhas por item homologado).
LINHA_PLACEHOLDER_SEMPRE = False

# True  -> mantem somente o 1o classificado de cada item.
# False -> traz todos os classificados (util em SRP com cadastro de reserva).
SOMENTE_PRIMEIRO_CLASSIFICADO = True

# True  -> preenche UND, DESCRICAO e SITUACAO com os dados da API.
# False -> deixa em branco, igual ao modelo.
PREENCHER_EXTRAS = False

DEBUG = False                   # True imprime as URLs chamadas (para diagnostico)

# =============================================================================
# CONSTANTES
# =============================================================================

BASE = "https://dadosabertos.compras.gov.br"
TAMANHO_PAGINA = 200
JANELA_DIAS = 365
PAUSA = 0.15                # pausa entre paginas
TIMEOUT_REQ = 30            # intervalo maximo sem receber bytes
LIMITE_RESPOSTA = 40_000_000   # teto de uma resposta, em bytes
TEMPO_LIMITE = 180          # limite da consulta inteira (o servidor corta em 300)
JANELA_ITENS = 45           # dias antes/depois da publicacao para buscar itens

MODALIDADES = {
    "PREGAO": 5,
    "DISPENSA": 6,
    "CONCORRENCIA": 3,
    "INEXIGIBILIDADE": 7,
    "CREDENCIAMENTO": 12,
}

COLUNAS = [
    "LOTE", "ITEM", "REQUISIÇÃO", "CNPJ", "EMPRESA", "QTDE", "UND",
    "VALOR UNIT LIC", "VALOR TOTAL LICI", "PRAZO", "DESCRIÇÃO",
    "SITUAÇÃO", "FORNECEDOR", "MODELO/VERSAO", "MARCA",
]

TEXTO_SEM_RESULTADO = "Item deserto/fracassado ou sem retorno na base consultada"

SESSAO = requests.Session()
SESSAO.headers.update({
    "accept": "*/*",
    "User-Agent": "gerador-siloms/1.0 (uso administrativo interno)",
})


def log(msg):
    print(msg, flush=True)


def dbg(msg):
    if DEBUG:
        print(f"    [debug] {msg}", flush=True)


# =============================================================================
# ACESSO A API
# =============================================================================

def _data_de(valor):
    """Le uma data da API ('2026-03-10T00:00:00Z' ou '2026-03-10') como date."""
    if not valor:
        return None
    txt = str(valor).strip().replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(txt).date()
    except ValueError:
        try:
            return date.fromisoformat(txt[:10])
        except ValueError:
            return None


def janelas_de_data(inicio, fim, dias=JANELA_DIAS):
    d0, d1 = date.fromisoformat(inicio), date.fromisoformat(fim)
    atual = d0
    while atual <= d1:
        prox = min(atual + timedelta(days=dias - 1), d1)
        yield atual.isoformat(), prox.isoformat()
        atual = prox + timedelta(days=1)


class TempoEsgotado(Exception):
    """A consulta inteira passou do prazo; abortamos para responder a tempo."""


_inicio_consulta = None
_limite_consulta = TEMPO_LIMITE


def iniciar_prazo(segundos=None):
    """Marca o inicio da consulta. Sem isso nao ha limite de tempo."""
    global _inicio_consulta, _limite_consulta
    _inicio_consulta = time.monotonic()
    _limite_consulta = segundos or TEMPO_LIMITE


def limpar_prazo():
    global _inicio_consulta
    _inicio_consulta = None


def _checar_prazo():
    if _inicio_consulta is None:
        return
    gasto = time.monotonic() - _inicio_consulta
    if gasto > _limite_consulta:
        raise TempoEsgotado(
            f"a consulta passou de {int(_limite_consulta)} segundos"
        )


def consultar(endpoint, params, obrigatorio=True, filtro=None):
    """
    Consulta paginada.

    'filtro' e aplicado pagina a pagina, antes de acumular. Isso importa:
    os endpoints de itens e de resultado nao aceitam filtrar por licitacao,
    so por UASG e data. Sem filtrar na hora, guardariamos na memoria milhares
    de registros de outras compras para descartar no fim.
    """
    registros, pagina = [], 1
    while True:
        _checar_prazo()
        p = dict(params, pagina=pagina, tamanhoPagina=TAMANHO_PAGINA)
        try:
            # stream=True e leitura em pedacos: o 'timeout' do requests nao
            # limita a duracao TOTAL da requisicao, so o intervalo sem receber
            # bytes. Uma resposta grande entregue devagar passava da trava de
            # tempo e ia morrer no limite do servidor. Lendo em pedacos, o
            # prazo e verificado durante o download, nao so entre paginas.
            r = SESSAO.get(f"{BASE}/{endpoint}", params=p,
                           timeout=TIMEOUT_REQ, stream=True)
            if pagina == 1:
                dbg(r.url)

            if r.status_code != 200:
                msg = f"HTTP {r.status_code} em {endpoint}"
                r.close()
                if obrigatorio:
                    raise RuntimeError(msg)
                log(f"    !! {msg}")
                return registros

            corpo = bytearray()
            for pedaco in r.iter_content(64 * 1024):
                _checar_prazo()
                corpo.extend(pedaco)
                if len(corpo) > LIMITE_RESPOSTA:
                    r.close()
                    raise TempoEsgotado(
                        "a resposta do Compras.gov.br veio grande demais"
                    )
            r.close()
        except requests.RequestException as e:
            if obrigatorio:
                raise
            log(f"    !! falha de rede em {endpoint}: {e}")
            return registros

        try:
            dados = json.loads(corpo.decode("utf-8", "replace")) if corpo else {}
        except ValueError:
            if obrigatorio:
                raise RuntimeError(f"resposta ilegível de {endpoint}")
            log(f"    !! resposta ilegível de {endpoint}")
            return registros
        lote = dados.get("resultado") or []
        registros.extend(lote if filtro is None else [x for x in lote if filtro(x)])

        restantes = dados.get("paginasRestantes", 0) or 0
        total_pag = dados.get("totalPaginas", 1) or 1
        if restantes <= 0 or pagina >= total_pag or not lote:
            break
        pagina += 1
        time.sleep(PAUSA)

    time.sleep(PAUSA)
    return registros


# =============================================================================
# HELPERS
# =============================================================================

def parse_numero_compra(valor, ano_config):
    """
    Normaliza o numero da compra e devolve (numero_9_digitos, ano).
      "90009/2026" -> ("900092026", 2026)
      "900092026"  -> ("900092026", 2026)
      "90009"      -> usa ANO_COMPRA
    """
    digitos = re.sub(r"\D", "", str(valor))
    if not digitos:
        raise ValueError("NUMERO_COMPRA vazio ou invalido.")

    if "/" in str(valor):
        num, ano = str(valor).split("/", 1)
        num, ano = re.sub(r"\D", "", num), int(re.sub(r"\D", "", ano))
    elif len(digitos) == 9:
        num, ano = digitos[:5], int(digitos[5:])
    elif ano_config:
        num, ano = digitos, int(ano_config)
    else:
        raise ValueError(
            "Informe o ano junto. Use o formato '90009/2026' ou '900092026'."
        )

    if not (2000 <= ano <= 2100):
        raise ValueError(f"Ano invalido: {ano}")
    return f"{int(num):05d}{ano}", ano


def so_digitos(v):
    return re.sub(r"\D", "", str(v or ""))


def numero_item_do_id(id_compra_item, fallback=None):
    """Extrai o numero do item a partir do idCompraItem, com fallback."""
    d = re.findall(r"\d+", str(id_compra_item or ""))
    if d:
        try:
            return int(d[-1])
        except ValueError:
            pass
    return fallback


def fmt_cnpj(ni):
    """CNPJ/CPF como texto, preservando zeros a esquerda."""
    d = so_digitos(ni)
    if not d:
        return ""
    if len(d) <= 11:
        return d.zfill(11)
    return d.zfill(14)


def num(valor, padrao=0.0):
    """
    Converte para float sem quebrar.
    A API ora devolve numero, ora texto ("1.234,56" ou "1234.56"), ora None.
    Qualquer coisa que nao der para converter vira o padrao.
    """
    if valor is None or valor == "":
        return padrao
    if isinstance(valor, (int, float)):
        return float(valor)

    txt = str(valor).strip()
    # formato brasileiro: 1.234,56 -> 1234.56
    if "," in txt:
        txt = txt.replace(".", "").replace(",", ".")
    txt = re.sub(r"[^\d.\-]", "", txt)
    try:
        return float(txt)
    except ValueError:
        return padrao


# =============================================================================
# EXTRACAO
# =============================================================================

def buscar_dados(uasg, numero, ano, cod_modalidade):
    # ---- 1. Localizar a contratacao -----------------------------------------
    log("\n[1/3] Localizando a contratacao...")
    contratacoes = []
    for d0, d1 in janelas_de_data(f"{ano}-01-01", f"{ano}-12-31"):
        contratacoes += consultar(
            "modulo-contratacoes/1_consultarContratacoes_PNCP_14133",
            {
                "unidadeOrgaoCodigoUnidade": uasg,
                "dataPublicacaoPncpInicial": d0,
                "dataPublicacaoPncpFinal": d1,
                "codigoModalidade": cod_modalidade,
            },
            obrigatorio=False,
        )

    alvo = [
        c for c in contratacoes
        if so_digitos(c.get("numeroCompra")).lstrip("0")
        == numero.lstrip("0")
    ]
    # fallback: comparar apenas o numero sequencial, sem o ano
    if not alvo:
        seq = numero[:5].lstrip("0")
        alvo = [
            c for c in contratacoes
            if so_digitos(c.get("numeroCompra")).lstrip("0").startswith(seq)
            and str(c.get("anoCompraPncp", ano)) == str(ano)
        ]

    if not alvo:
        log(f"    !! Compra {numero} nao encontrada na UASG {uasg} em {ano}.")
        log(f"       Foram varridas {len(contratacoes)} contratacao(oes) da modalidade.")
        if contratacoes:
            amostra = sorted({str(c.get("numeroCompra")) for c in contratacoes})[:15]
            log(f"       Numeros disponiveis (amostra): {', '.join(amostra)}")
        return None, [], []

    compra = alvo[0]
    id_compra = compra.get("idCompra")
    log(f"    OK  {compra.get('numeroCompra')} | {compra.get('modalidadeNome')}")
    log(f"        processo: {compra.get('processo')}")
    log(f"        objeto..: {str(compra.get('objetoCompra') or '')[:80]}")
    log(f"        idCompra: {id_compra}")

    # ---- janela de datas ----------------------------------------------------
    # Os endpoints de itens e resultado nao filtram por licitacao, so por UASG
    # e data. Varrer o ano inteiro sao dezenas de requisicoes e estoura o
    # tempo do servidor. Como os itens entram no PNCP junto com a compra,
    # buscamos so em volta da data dela.
    ref = _data_de(compra.get("dataPublicacaoPncp")
                   or compra.get("dataInclusaoPncp")) or date(ano, 1, 1)
    hoje = date.today()

    ini_itens = max(date(ano, 1, 1), ref - timedelta(days=JANELA_ITENS))
    fim_itens = min(hoje, ref + timedelta(days=JANELA_ITENS))
    if fim_itens < ini_itens:
        fim_itens = ini_itens

    so_desta_compra = lambda x: x.get("idCompra") == id_compra

    # ---- 2. Itens da contratacao --------------------------------------------
    log("\n[2/3] Itens da contratacao...")
    log(f"    janela: {ini_itens} a {fim_itens}")

    def _buscar_itens(d0, d1):
        base = {
            "unidadeOrgaoCodigoUnidade": uasg,
            "dataInclusaoPncpInicial": d0.isoformat(),
            "dataInclusaoPncpFinal": d1.isoformat(),
        }
        lote = consultar(
            "modulo-contratacoes/2_consultarItensContratacoes_PNCP_14133",
            base, obrigatorio=False, filtro=so_desta_compra,
        )
        if lote:
            return lote
        # a doc marca materialOuServico como obrigatorio; tentamos os dois
        for tipo in ("M", "S"):
            lote += consultar(
                "modulo-contratacoes/2_consultarItensContratacoes_PNCP_14133",
                {**base, "materialOuServico": tipo},
                obrigatorio=False, filtro=so_desta_compra,
            )
        return lote

    itens = _buscar_itens(ini_itens, fim_itens)

    # se a janela estreita nao achou nada, alarga uma vez para o ano inteiro
    if not itens:
        log("    nada na janela; ampliando para o ano todo")
        itens = _buscar_itens(date(ano, 1, 1), min(hoje, date(ano, 12, 31)))

    log(f"    {len(itens)} item(ns) da compra")

    # ---- 3. Resultados (vencedores) -----------------------------------------
    # O resultado sai depois da publicacao, entao a janela vai da compra ate hoje.
    log("\n[3/3] Resultados dos itens...")
    ini_res = max(date(ano, 1, 1), ref - timedelta(days=15))
    fim_res = max(hoje, ini_res)
    log(f"    janela: {ini_res} a {fim_res}")

    resultados = []
    for d0, d1 in janelas_de_data(ini_res.isoformat(), fim_res.isoformat()):
        resultados += consultar(
            "modulo-contratacoes/3_consultarResultadoItensContratacoes_PNCP_14133",
            {
                "unidadeOrgaoCodigoUnidade": uasg,
                "dataResultadoPncpInicial": d0,
                "dataResultadoPncpFinal": d1,
            },
            obrigatorio=False, filtro=so_desta_compra,
        )

    log(f"    {len(resultados)} resultado(s) da compra")

    return compra, itens, resultados


# =============================================================================
# MONTAGEM DO LAYOUT SILOMS
# =============================================================================

def montar_planilha(itens, resultados):
    # indice de resultados por idCompraItem
    por_item = {}
    for r in resultados:
        por_item.setdefault(r.get("idCompraItem"), []).append(r)

    for lista in por_item.values():
        lista.sort(key=lambda x: (
            x.get("ordemClassificacaoSrp") or 9999,
            x.get("sequencialResultado") or 9999,
        ))

    # se o endpoint de itens falhar, monta a partir dos proprios resultados
    if not itens and resultados:
        log("    (!) Sem retorno no endpoint de itens; montando pelos resultados.")
        vistos, itens = set(), []
        for r in resultados:
            k = r.get("idCompraItem")
            if k in vistos:
                continue
            vistos.add(k)
            itens.append({
                "idCompraItem": k,
                "numeroItemCompra": numero_item_do_id(k, r.get("numeroItemPncp")),
                "numeroGrupo": 0,
                "descricaoResumida": "",
                "unidadeMedida": "",
            })

    def chave_item(i):
        n = i.get("numeroItemCompra") or numero_item_do_id(i.get("idCompraItem"))
        try:
            return int(n)
        except (TypeError, ValueError):
            return 10**9

    itens = sorted(itens, key=chave_item)

    linhas = []
    for it in itens:
        num_item = chave_item(it)
        num_item = "" if num_item == 10**9 else float(num_item)
        grupo = num(it.get("numeroGrupo"))
        lote = grupo if grupo else ""

        vencedores = por_item.get(it.get("idCompraItem"), [])
        if SOMENTE_PRIMEIRO_CLASSIFICADO and vencedores:
            vencedores = vencedores[:1]

        base_extras = {
            "UND": it.get("unidadeMedida") or "" if PREENCHER_EXTRAS else "",
            "DESCRIÇÃO": it.get("descricaoResumida") or "" if PREENCHER_EXTRAS else "",
        }

        # linha de item sem resultado
        if LINHA_PLACEHOLDER_SEMPRE or not vencedores:
            linhas.append({
                "LOTE": lote, "ITEM": num_item, "REQUISIÇÃO": "",
                "CNPJ": TEXTO_SEM_RESULTADO, "EMPRESA": "",
                "QTDE": 0.0, "UND": base_extras["UND"],
                "VALOR UNIT LIC": 0.0, "VALOR TOTAL LICI": 0.0,
                "PRAZO": None, "DESCRIÇÃO": base_extras["DESCRIÇÃO"],
                "SITUAÇÃO": "", "FORNECEDOR": "",
                "MODELO/VERSAO": "", "MARCA": "",
            })

        for v in vencedores:
            qtd = num(v.get("quantidadeHomologada"))
            vu = num(v.get("valorUnitarioHomologado"))
            vt = num(v.get("valorTotalHomologado"))
            if vt == 0.0 and qtd and vu:
                vt = round(qtd * vu, 4)

            linhas.append({
                "LOTE": lote,
                "ITEM": num_item,
                "REQUISIÇÃO": "",
                "CNPJ": fmt_cnpj(v.get("niFornecedor")),
                "EMPRESA": v.get("nomeRazaoSocialFornecedor") or "",
                "QTDE": qtd,
                "UND": base_extras["UND"],
                "VALOR UNIT LIC": vu,
                "VALOR TOTAL LICI": vt,
                "PRAZO": float(PRAZO_PADRAO),
                "DESCRIÇÃO": base_extras["DESCRIÇÃO"],
                "SITUAÇÃO": (v.get("situacaoCompraItemResultadoNome") or ""
                             if PREENCHER_EXTRAS else ""),
                "FORNECEDOR": "",
                "MODELO/VERSAO": "",
                "MARCA": "",
            })

    return pd.DataFrame(linhas, columns=COLUNAS)


def gravar_xlsx(df, caminho):
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet0", index=False)
        ws = writer.sheets["Sheet0"]

        idx = {c: i + 1 for i, c in enumerate(COLUNAS)}

        # cabecalho
        for i in range(1, len(COLUNAS) + 1):
            cel = ws.cell(row=1, column=i)
            cel.font = Font(bold=True)
            cel.fill = PatternFill("solid", fgColor="DDDDDD")
            cel.alignment = Alignment(horizontal="center", vertical="center")

        # formatos por coluna
        for linha in range(2, len(df) + 2):
            for col in ("VALOR UNIT LIC", "VALOR TOTAL LICI"):
                ws.cell(row=linha, column=idx[col]).number_format = "0.0000"
            c = ws.cell(row=linha, column=idx["CNPJ"])
            c.number_format = "@"          # texto: preserva zero a esquerda
            c.alignment = Alignment(horizontal="left")

        larguras = {
            "LOTE": 8, "ITEM": 8, "REQUISIÇÃO": 12, "CNPJ": 52, "EMPRESA": 45,
            "QTDE": 10, "UND": 8, "VALOR UNIT LIC": 15, "VALOR TOTAL LICI": 17,
            "PRAZO": 8, "DESCRIÇÃO": 30, "SITUAÇÃO": 16, "FORNECEDOR": 18,
            "MODELO/VERSAO": 16, "MARCA": 14,
        }
        for col, larg in larguras.items():
            ws.column_dimensions[get_column_letter(idx[col])].width = larg

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions


# =============================================================================
# FORMULARIO NO CONSOLE
# =============================================================================

def perguntar(rotulo, padrao=None, obrigatorio=True, validador=None):
    """
    Pergunta um valor ao usuario. ENTER aceita o padrao entre colchetes.
    'validador' recebe o texto digitado e devolve (ok, valor_ou_mensagem).
    """
    sufixo = f" [{padrao}]" if padrao else ""
    while True:
        try:
            resp = input(f"  {rotulo}{sufixo}: ").strip()
        except EOFError:
            raise KeyboardInterrupt

        if not resp and padrao is not None:
            resp = str(padrao)

        if not resp:
            if not obrigatorio:
                return ""
            print("     >> Campo obrigatorio. Tente novamente.")
            continue

        if validador is None:
            return resp

        ok, valor = validador(resp)
        if ok:
            return valor
        print(f"     >> {valor}")


def valida_uasg(txt):
    d = re.sub(r"\D", "", txt)
    if not d:
        return False, "Informe apenas numeros. Ex.: 120641"
    if not (5 <= len(d) <= 6):
        return False, "A UASG tem 5 ou 6 digitos. Ex.: 120641"
    return True, d


def valida_tipo(txt):
    t = txt.strip().upper()
    atalhos = {
        "1": "PREGAO", "2": "DISPENSA", "3": "CONCORRENCIA",
        "4": "INEXIGIBILIDADE", "5": "CREDENCIAMENTO",
        "PREGÃO": "PREGAO", "CONCORRÊNCIA": "CONCORRENCIA",
        "DISPENSA DE LICITACAO": "DISPENSA",
        "DISPENSA DE LICITAÇÃO": "DISPENSA",
    }
    t = atalhos.get(t, t)
    if t not in MODALIDADES:
        return False, f"Opcoes: {', '.join(MODALIDADES)} (ou os numeros 1 a 5)"
    return True, t


def valida_compra(txt):
    try:
        numero, ano = parse_numero_compra(txt, None)
    except ValueError as e:
        return False, str(e)
    return True, (numero, ano)


def coletar_filtros():
    """Formulario equivalente ao do site de referencia."""
    print()
    print("-" * 70)
    print(" FILTROS  (ENTER aceita o valor sugerido entre colchetes)")
    print("-" * 70)

    uasg = perguntar("UASG.................", UASG_PADRAO, validador=valida_uasg)

    print()
    print("     Tipo de licitacao:")
    print("       1 - PREGAO            2 - DISPENSA DE LICITACAO")
    print("       3 - CONCORRENCIA      4 - INEXIGIBILIDADE")
    print("       5 - CREDENCIAMENTO")
    tipo = perguntar("Tipo.................", TIPO_PADRAO, validador=valida_tipo)

    print()
    numero, ano = perguntar(
        "N. do Processo.......", None, validador=valida_compra
    )

    return uasg, tipo, numero, ano


# =============================================================================
# EXECUCAO DE UMA CONSULTA
# =============================================================================

def executar(uasg, tipo, numero, ano):
    """Gera a planilha de uma compra. Devolve True se gerou o arquivo."""
    cod_mod = MODALIDADES[tipo]

    print()
    log("-" * 70)
    log(f" UASG......: {uasg}")
    log(f" Tipo......: {tipo} (codigo {cod_mod})")
    log(f" Processo..: {numero[:5]}/{ano}")
    log("-" * 70)

    compra, itens, resultados = buscar_dados(uasg, numero, ano, cod_mod)
    if compra is None:
        return False

    df = montar_planilha(itens, resultados)
    if df.empty:
        log("\n!! Nenhuma linha gerada. Planilha nao criada.")
        return False

    saida = f"compra_{numero}.xlsx"
    try:
        gravar_xlsx(df, saida)
    except PermissionError:
        log(f"\n!! Nao consegui gravar {saida}.")
        log("   O arquivo provavelmente esta aberto no Excel. Feche e tente de novo.")
        return False

    com_venc = int((df["CNPJ"] != TEXTO_SEM_RESULTADO).sum())
    sem_venc = int((df["CNPJ"] == TEXTO_SEM_RESULTADO).sum())
    total = df.loc[df["CNPJ"] != TEXTO_SEM_RESULTADO, "VALOR TOTAL LICI"].sum()
    valor_fmt = (f"{total:,.2f}".replace(",", "@")
                 .replace(".", ",").replace("@", "."))

    log("\n" + "-" * 70)
    log(f" Itens distintos....: {df['ITEM'].nunique()}")
    log(f" Linhas com vencedor: {com_venc}")
    log(f" Linhas placeholder.: {sem_venc}")
    log(f" Valor homologado...: R$ {valor_fmt}")
    log("-" * 70)
    log(f"\n>> Planilha gerada: {saida}")
    return True


# =============================================================================
# MAIN
# =============================================================================

def main():
    log("=" * 70)
    log(" GERADOR DE PLANILHA SILOMS - Resultado de Licitacao")
    log(" Fonte: API Dados Abertos do Compras.gov.br")
    log("=" * 70)

    ultimo = None   # guarda (uasg, tipo) para sugerir na proxima rodada

    while True:
        if ultimo:
            global UASG_PADRAO, TIPO_PADRAO
            UASG_PADRAO, TIPO_PADRAO = ultimo

        try:
            uasg, tipo, numero, ano = coletar_filtros()
        except KeyboardInterrupt:
            log("\n\nEncerrado.")
            return

        try:
            executar(uasg, tipo, numero, ano)
        except KeyboardInterrupt:
            log("\n\nConsulta interrompida.")
        except Exception as e:
            log(f"\n!! Falha nesta consulta: {type(e).__name__}: {e}")
            if DEBUG:
                import traceback
                traceback.print_exc()

        ultimo = (uasg, tipo)

        print()
        try:
            de_novo = input("Gerar outra planilha? (S/N) [N]: ").strip().upper()
        except (EOFError, KeyboardInterrupt):
            return
        if de_novo not in ("S", "SIM", "Y", "YES"):
            return
        print()


if __name__ == "__main__":
    codigo = 0
    try:
        main()
    except KeyboardInterrupt:
        log("\nInterrompido pelo usuario.")
        codigo = 1
    except SystemExit as e:
        codigo = e.code or 0
    except Exception as e:
        log(f"\n!! ERRO: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        codigo = 1

    # Mantem a janela aberta quando o script e executado por duplo-clique.
    try:
        input("\nPressione ENTER para fechar...")
    except EOFError:
        pass

    sys.exit(codigo)
