#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
 Cliente da API de Dados Abertos do Compras.gov.br
=============================================================================
 Tres cuidados moldam este arquivo, cada um aprendido na pratica:

 1. PRAZO PROPRIO. O servidor derruba a requisicao depois de um tempo. Aqui
    a consulta desiste antes disso e responde com uma mensagem, em vez de
    morrer em silencio.

 2. LEITURA QUE NAO TRAVA. O 'timeout' do requests nao limita a duracao
    total da chamada, so o intervalo sem receber bytes. Uma resposta que
    nunca comeca, ou que chega devagar, passa por ele. Por isso a chamada
    roda numa thread com prazo, e o corpo e lido em pedacos.

 3. PAGINAS EM PARALELO. Os endpoints de itens e de resultado nao aceitam
    filtrar por licitacao, so por unidade e data. Buscar pagina por pagina,
    em fila, e o que estourava o tempo. Aqui a primeira pagina revela quantas
    existem e as demais vem juntas.
=============================================================================
"""

import json
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta

import requests

VERSAO = "2.1.0"

BASE = "https://dadosabertos.compras.gov.br"

TAMANHO_PAGINA = 500        # registros por pagina (maximo aceito pela API)
PARALELO = 6                # paginas buscadas ao mesmo tempo
TIMEOUT_REQ = 25            # intervalo maximo sem receber bytes
TEMPO_LIMITE = 120          # prazo da consulta inteira
LIMITE_RESPOSTA = 60_000_000

MODALIDADES = {
    5: "Pregão Eletrônico",
    6: "Dispensa Eletrônica",
    7: "Inexigibilidade",
    3: "Concorrência Eletrônica",
    12: "Credenciamento",
    1: "Convite",
    2: "Tomada de Preços",
}

# as fases que um certame percorre, em ordem
FASES = [
    "Aguardando abertura",
    "Em disputa",
    "Em habilitação",
    "Em homologação",
    "Homologado",
]

SESSAO = requests.Session()
SESSAO.headers.update({
    "accept": "application/json",
    "User-Agent": "painel-certames/2.0",
})


# =============================================================================
# PRAZO
# =============================================================================

class TempoEsgotado(Exception):
    """A consulta passou do prazo e foi interrompida de proposito."""


_local = threading.local()


def iniciar_prazo(segundos=None):
    _local.inicio = time.monotonic()
    _local.limite = segundos or TEMPO_LIMITE
    _local.requisicoes = 0
    _local.etapa = ""


def limpar_prazo():
    _local.inicio = None


def etapa(nome):
    _local.etapa = nome


def diagnostico():
    """Onde a consulta estava e quanto ja tinha gasto. Vai na mensagem de erro."""
    inicio = getattr(_local, "inicio", None)
    return {
        "etapa": getattr(_local, "etapa", ""),
        "requisicoes": getattr(_local, "requisicoes", 0),
        "segundos": round(time.monotonic() - inicio, 1) if inicio else 0,
    }


def _restante():
    inicio = getattr(_local, "inicio", None)
    if inicio is None:
        return None
    return getattr(_local, "limite", TEMPO_LIMITE) - (time.monotonic() - inicio)


def _checar_prazo():
    r = _restante()
    if r is not None and r <= 0:
        d = diagnostico()
        raise TempoEsgotado(
            f"{d['segundos']}s em “{d['etapa']}”, {d['requisicoes']} requisições"
        )


# =============================================================================
# REQUISICAO
# =============================================================================

def _get(url, params):
    """
    Faz a requisicao numa thread e desiste se ela nao voltar no prazo.
    Uma leitura de socket parada nao pode ser interrompida de fora; o que
    da para fazer e parar de esperar por ela.
    """
    caixa = {}

    def tentar():
        try:
            caixa["r"] = SESSAO.get(url, params=params,
                                    timeout=TIMEOUT_REQ, stream=True)
        except BaseException as e:                       # noqa: BLE001
            caixa["erro"] = e

    limite = TIMEOUT_REQ + 8
    r = _restante()
    if r is not None:
        if r <= 0:
            _checar_prazo()
        limite = min(limite, r)

    t = threading.Thread(target=tentar, daemon=True)
    t.start()
    t.join(limite)

    if t.is_alive():
        raise TempoEsgotado(
            f"o Compras.gov.br não respondeu em {int(limite)} segundos"
        )
    if "erro" in caixa:
        raise caixa["erro"]
    return caixa["r"]


def _pagina(endpoint, params, numero):
    """Baixa uma pagina e devolve o JSON decodificado."""
    _local.requisicoes = getattr(_local, "requisicoes", 0) + 1
    p = dict(params, pagina=numero, tamanhoPagina=TAMANHO_PAGINA)
    r = _get(f"{BASE}/{endpoint}", p)
    try:
        if r.status_code != 200:
            raise RuntimeError(f"HTTP {r.status_code} em {endpoint}")
        corpo = bytearray()
        for pedaco in r.iter_content(64 * 1024):
            _checar_prazo()
            corpo.extend(pedaco)
            if len(corpo) > LIMITE_RESPOSTA:
                raise TempoEsgotado("a resposta veio grande demais")
    finally:
        r.close()
    if not corpo:
        return {}
    return json.loads(corpo.decode("utf-8", "replace"))


def consultar(endpoint, params, filtro=None, tolerante=True):
    """
    Busca todas as paginas de um endpoint.

    A primeira pagina diz quantas existem; as demais sao buscadas em
    paralelo. 'filtro' e aplicado pagina a pagina, antes de acumular, para
    nao guardar na memoria registros que serao descartados.
    """
    registros = []

    def juntar(dados):
        lote = dados.get("resultado") or []
        if filtro is None:
            registros.extend(lote)
        else:
            registros.extend(x for x in lote if filtro(x))
        return lote

    try:
        primeira = _pagina(endpoint, params, 1)
    except TempoEsgotado:
        raise
    except Exception:
        if tolerante:
            return registros
        raise

    lote = juntar(primeira)
    total = primeira.get("totalPaginas") or 1
    if not lote or total <= 1:
        return registros

    restantes = list(range(2, int(total) + 1))
    executor = ThreadPoolExecutor(max_workers=PARALELO)
    try:
        for i in range(0, len(restantes), PARALELO):
            _checar_prazo()
            bloco = restantes[i:i + PARALELO]
            futuros = [executor.submit(_pagina, endpoint, params, n) for n in bloco]
            for f in as_completed(futuros):
                try:
                    juntar(f.result())
                except TempoEsgotado:
                    raise
                except Exception:
                    if not tolerante:
                        raise
    finally:
        executor.shutdown(wait=False, cancel_futures=True)

    return registros


# =============================================================================
# AUXILIARES
# =============================================================================

def so_digitos(v):
    return re.sub(r"\D", "", str(v or ""))


def para_data(valor):
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


def numero_legivel(numero_compra, ano):
    """'000612026' vira '61/2026'."""
    d = so_digitos(numero_compra)
    if len(d) >= 9:
        seq, a = d[:-4], d[-4:]
        return f"{int(seq)}/{a}"
    return f"{d or '?'}/{ano}"


def _fase(compra, hoje=None):
    """
    Deduz em que fase o certame esta.

    A API de dados abertos nao expoe o andamento operacional do
    Compras.gov.br — ela traz a situacao do PNCP (divulgada, revogada,
    anulada, suspensa), as datas de proposta e se ja existe resultado.
    A fase abaixo e deduzida desses campos, nao lida do sistema.
    """
    hoje = hoje or date.today()

    situacao = str(compra.get("situacaoCompraNomePncp") or "").strip()
    if situacao and "divulgad" not in situacao.lower():
        return {"nome": situacao, "ordem": None, "excecao": True}

    abertura = para_data(compra.get("dataAberturaPropostaPncp"))
    encerramento = para_data(compra.get("dataEncerramentoPropostaPncp"))
    tem_resultado = bool(compra.get("existeResultado"))
    homologado = float(compra.get("valorTotalHomologado") or 0) > 0

    if abertura and hoje < abertura:
        nome = "Aguardando abertura"
    elif abertura and encerramento and abertura <= hoje <= encerramento:
        nome = "Em disputa"
    elif encerramento and hoje > encerramento and not tem_resultado:
        nome = "Em habilitação"
    elif tem_resultado and not homologado:
        nome = "Em homologação"
    elif homologado:
        nome = "Homologado"
    elif tem_resultado:
        nome = "Em homologação"
    else:
        nome = "Em habilitação"

    return {"nome": nome, "ordem": FASES.index(nome) + 1, "excecao": False}


# =============================================================================
# CERTAMES
# =============================================================================

def listar_certames(uasg, ano, modalidades=(5, 6, 7, 3, 12)):
    """Todos os certames da unidade no ano, com a fase deduzida."""
    etapa("lista de certames")
    hoje = date.today()
    fim = min(date(ano, 12, 31), hoje) if ano == hoje.year else date(ano, 12, 31)

    # As modalidades vao juntas. Em fila, cinco chamadas lentas somavam mais
    # que o prazo inteiro da consulta -- era o gargalo real da listagem.
    def uma(mod):
        return consultar(
            "modulo-contratacoes/1_consultarContratacoes_PNCP_14133",
            {
                "unidadeOrgaoCodigoUnidade": str(uasg),
                "dataPublicacaoPncpInicial": f"{ano}-01-01",
                "dataPublicacaoPncpFinal": fim.isoformat(),
                "codigoModalidade": mod,
            },
        )

    achados, faltaram = {}, []
    executor = ThreadPoolExecutor(max_workers=min(len(modalidades), PARALELO))
    try:
        futuros = {executor.submit(uma, m): m for m in modalidades}
        for f in as_completed(futuros):
            mod = futuros[f]
            try:
                lote = f.result()
            except TempoEsgotado:
                # uma modalidade lenta nao derruba as outras: seguimos com o
                # que chegou e avisamos o que faltou
                faltaram.append(MODALIDADES.get(mod, str(mod)))
                continue
            except Exception:
                faltaram.append(MODALIDADES.get(mod, str(mod)))
                continue
            for c in lote:
                if c.get("contratacaoExcluida"):
                    continue
                achados[c.get("idCompra")] = c
    finally:
        executor.shutdown(wait=False, cancel_futures=True)

    if faltaram and not achados:
        raise TempoEsgotado(
            "nenhuma modalidade respondeu a tempo"
        )

    saida = []
    for c in achados.values():
        mod = int(c.get("codigoModalidade") or 0)
        ano_c = int(c.get("anoCompraPncp") or ano)
        fase = _fase(c, hoje)
        saida.append({
            "idCompra": c.get("idCompra"),
            "titulo": (f"{MODALIDADES.get(mod, c.get('modalidadeNome') or 'Contratação')} "
                       f"{uasg} - {numero_legivel(c.get('numeroCompra'), ano_c)}"),
            "modalidade": MODALIDADES.get(mod, c.get("modalidadeNome") or ""),
            "numero": numero_legivel(c.get("numeroCompra"), ano_c),
            "uasg": str(uasg),
            "ano": ano_c,
            "objeto": (c.get("objetoCompra") or "").strip(),
            "processo": c.get("processo") or "",
            "fase": fase["nome"],
            "fase_ordem": fase["ordem"],
            "fase_excecao": fase["excecao"],
            "srp": bool(c.get("srp")),
            "abertura": (para_data(c.get("dataAberturaPropostaPncp")) or ""),
            "encerramento": (para_data(c.get("dataEncerramentoPropostaPncp")) or ""),
            "publicacao": (para_data(c.get("dataPublicacaoPncp")) or ""),
            "valor_estimado": float(c.get("valorTotalEstimado") or 0),
            "valor_homologado": float(c.get("valorTotalHomologado") or 0),
            "tem_resultado": bool(c.get("existeResultado")),
        })

    saida.sort(key=lambda x: (str(x["publicacao"]), x["numero"]), reverse=True)
    for c in saida:
        for campo in ("abertura", "encerramento", "publicacao"):
            c[campo] = c[campo].isoformat() if c[campo] else ""
    return saida, faltaram


def resultado_certame(uasg, id_compra, publicacao=None, janela=45):
    """
    Itens e resultados de um certame.

    Os dois endpoints so filtram por unidade e data, entao a busca e feita
    numa janela em torno da publicacao e os registros de outras compras sao
    descartados ja na chegada.
    """
    hoje = date.today()
    ref = para_data(publicacao) or hoje
    ano = ref.year
    desta = lambda x: x.get("idCompra") == id_compra

    # ---- itens -------------------------------------------------------------
    etapa("itens do certame")
    ini = max(date(ano, 1, 1), ref - timedelta(days=janela))
    fim = min(hoje, ref + timedelta(days=janela))
    if fim < ini:
        fim = ini

    def buscar_itens(d0, d1):
        base = {
            "unidadeOrgaoCodigoUnidade": str(uasg),
            "dataInclusaoPncpInicial": d0.isoformat(),
            "dataInclusaoPncpFinal": d1.isoformat(),
        }
        achado = consultar(
            "modulo-contratacoes/2_consultarItensContratacoes_PNCP_14133",
            base, filtro=desta)
        if achado:
            return achado
        for tipo in ("M", "S"):
            achado += consultar(
                "modulo-contratacoes/2_consultarItensContratacoes_PNCP_14133",
                {**base, "materialOuServico": tipo}, filtro=desta)
        return achado

    itens = buscar_itens(ini, fim)
    if not itens:                       # janela estreita nao achou: abre o ano
        itens = buscar_itens(date(ano, 1, 1), min(hoje, date(ano, 12, 31)))

    # ---- resultados --------------------------------------------------------
    etapa("resultado dos itens")
    ini_r = max(date(ano, 1, 1), ref - timedelta(days=15))
    fim_r = max(hoje, ini_r)
    resultados = consultar(
        "modulo-contratacoes/3_consultarResultadoItensContratacoes_PNCP_14133",
        {
            "unidadeOrgaoCodigoUnidade": str(uasg),
            "dataResultadoPncpInicial": ini_r.isoformat(),
            "dataResultadoPncpFinal": fim_r.isoformat(),
        },
        filtro=desta)

    return itens, resultados
