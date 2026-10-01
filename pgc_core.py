#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
 PGC - Plano de Contratacoes Anual
=============================================================================
 Consulta o modulo PGC da API Dados Abertos do Compras.gov.br e consolida
 o resultado em UMA LINHA POR DFD (Documento de Formalizacao de Demanda).

 A API devolve os dados no nivel de item; cada item carrega a identidade do
 DFD e do projeto de compra. Aqui esses itens sao agrupados por DFD.

 Depende de gerar_planilha_siloms.py (reaproveita a paginacao e a sessao).
=============================================================================
"""

import re
from datetime import datetime

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import gerar_planilha_siloms as core


# =============================================================================
# COLUNAS DA PLANILHA
# =============================================================================

COLUNAS = [
    "DFD", "ORDEM", "OBJETO DO DFD", "AREA", "PRIORIDADE",
    "DATA PREVISTA FORMALIZACAO", "QTD ITENS", "VALOR TOTAL",
    "CONTRATACAO", "INICIO DO PROCESSO", "FIM DO PROCESSO",
    "DURACAO (DIAS)", "ITEM PNCP", "EM EXECUCAO",
    "PUBLICACAO PNCP", "UASG",
]

COLUNAS_DATA = {
    "DATA PREVISTA FORMALIZACAO", "INICIO DO PROCESSO",
    "FIM DO PROCESSO", "PUBLICACAO PNCP",
}


# =============================================================================
# CONSULTAS
# =============================================================================

def obter_cnpj_orgao(uasg):
    """
    Descobre o CNPJ do orgao a partir do codigo da UASG.
    O endpoint de PGC exige o CNPJ, que o usuario nao precisa saber de cor.
    Devolve (cnpj, nome_uasg) ou (None, None).
    """
    registros = core.consultar(
        "modulo-uasg/1_consultarUasg",
        {"codigoUasg": str(uasg), "statusUasg": "true"},
        obrigatorio=False,
    )
    if not registros:
        # tenta sem o filtro de status, caso a unidade esteja inativa
        registros = core.consultar(
            "modulo-uasg/1_consultarUasg",
            {"codigoUasg": str(uasg)},
            obrigatorio=False,
        )
    if not registros:
        return None, None

    reg = registros[0]
    cnpj = re.sub(r"\D", "", str(reg.get("cnpjCpfOrgao") or ""))
    return (cnpj or None), reg.get("nomeUasg")


def buscar_pgc(uasg, ano, cnpj_orgao):
    """Consulta os itens do plano de contratacoes da UASG no ano informado."""
    registros = core.consultar(
        "modulo-pgc/1_consultarPgcDetalhe",
        {
            "orgao": cnpj_orgao,
            "anoPcaProjetoCompra": int(ano),
            "codigoUasg": str(uasg),
        },
        obrigatorio=False,
    )
    # a API as vezes ignora o filtro de UASG; garante o recorte
    alvo = str(uasg).lstrip("0")
    return [
        r for r in registros
        if str(r.get("codigoUasg") or "").lstrip("0") == alvo
    ] or registros


# =============================================================================
# CONSOLIDACAO POR DFD
# =============================================================================

def _data_curta(valor):
    """Converte a data ISO da API para date, descartando hora e fuso."""
    if not valor:
        return None
    try:
        return datetime.fromisoformat(str(valor).replace("Z", "+00:00")).date()
    except (ValueError, TypeError):
        return None


def _primeiro(serie):
    """Primeiro valor nao vazio de uma coluna dentro do grupo."""
    for v in serie:
        if v not in (None, "", 0):
            return v
    return ""


def _unico_ou_varios(serie):
    """Devolve o valor unico do grupo ou 'varios' quando houver divergencia."""
    vistos = [v for v in dict.fromkeys(serie) if v not in (None, "")]
    if not vistos:
        return ""
    if len(vistos) == 1:
        return vistos[0]
    return f"{len(vistos)} diferentes"


def montar_por_dfd(registros):
    """Agrupa os itens do PGC em uma linha por DFD."""
    if not registros:
        return pd.DataFrame(columns=COLUNAS)

    linhas = []
    for r in registros:
        num = r.get("numeroArtefato")
        ano = r.get("anoArtefato")
        dfd = f"{num}/{ano}" if num and ano else (str(num or "") or "sem número")

        linhas.append({
            "_chave": (num, ano, r.get("ordemDfd")),
            "DFD": dfd,
            "ORDEM": r.get("ordemDfd"),
            "OBJETO DO DFD": r.get("descricaoObjetoDfd") or "",
            "AREA": r.get("codigoAreaDfd") or "",
            "PRIORIDADE": r.get("nivelPrioridadeDfd"),
            "DATA PREVISTA FORMALIZACAO":
                _data_curta(r.get("dataPrevistaFormalizacaoDemanda")),
            "_valor": float(r.get("valorTotalItem") or 0),
            "CONTRATACAO": r.get("tituloProjetoCompra") or "",
            "INICIO DO PROCESSO": _data_curta(r.get("dataInicioProcessoCompra")),
            "FIM DO PROCESSO": _data_curta(r.get("dataFimProcessoCompra")),
            "DURACAO (DIAS)": r.get("duracaoProcessoCompra"),
            "ITEM PNCP": r.get("numeroItemPncp"),
            "EM EXECUCAO": "Sim" if r.get("statusContratacaoExecucao") else "Não",
            "PUBLICACAO PNCP": _data_curta(r.get("dataHoraPublicacaoPncp")),
            "UASG": r.get("codigoUasg") or "",
        })

    bruto = pd.DataFrame(linhas)

    agrupado = bruto.groupby("_chave", dropna=False, sort=False).agg(**{
        "DFD": ("DFD", "first"),
        "ORDEM": ("ORDEM", "first"),
        "OBJETO DO DFD": ("OBJETO DO DFD", _primeiro),
        "AREA": ("AREA", _primeiro),
        "PRIORIDADE": ("PRIORIDADE", "first"),
        "DATA PREVISTA FORMALIZACAO": ("DATA PREVISTA FORMALIZACAO", _primeiro),
        "QTD ITENS": ("_valor", "size"),
        "VALOR TOTAL": ("_valor", "sum"),
        "CONTRATACAO": ("CONTRATACAO", _unico_ou_varios),
        "INICIO DO PROCESSO": ("INICIO DO PROCESSO", _primeiro),
        "FIM DO PROCESSO": ("FIM DO PROCESSO", _primeiro),
        "DURACAO (DIAS)": ("DURACAO (DIAS)", "first"),
        "ITEM PNCP": ("ITEM PNCP", _unico_ou_varios),
        "EM EXECUCAO": ("EM EXECUCAO", _unico_ou_varios),
        "PUBLICACAO PNCP": ("PUBLICACAO PNCP", _primeiro),
        "UASG": ("UASG", "first"),
    }).reset_index(drop=True)

    # ordena por numero e ordem do DFD
    def chave_ordem(v):
        try:
            return int(v)
        except (TypeError, ValueError):
            return 10 ** 9

    agrupado["_o1"] = agrupado["DFD"].map(
        lambda d: chave_ordem(str(d).split("/")[0])
    )
    agrupado["_o2"] = agrupado["ORDEM"].map(chave_ordem)
    agrupado = (agrupado.sort_values(["_o1", "_o2"])
                        .drop(columns=["_o1", "_o2"]))

    for col in COLUNAS:
        if col not in agrupado.columns:
            agrupado[col] = ""

    return agrupado[COLUNAS]


# =============================================================================
# PLANILHA
# =============================================================================

def gravar_xlsx(df, caminho):
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="DFD", index=False)
        ws = writer.sheets["DFD"]

        idx = {c: i + 1 for i, c in enumerate(COLUNAS)}

        for i in range(1, len(COLUNAS) + 1):
            cel = ws.cell(row=1, column=i)
            cel.font = Font(bold=True)
            cel.fill = PatternFill("solid", fgColor="DDDDDD")
            cel.alignment = Alignment(horizontal="center", vertical="center")

        for linha in range(2, len(df) + 2):
            ws.cell(row=linha, column=idx["VALOR TOTAL"]).number_format = \
                'R$ #,##0.00'
            for col in COLUNAS_DATA:
                ws.cell(row=linha, column=idx[col]).number_format = "DD/MM/YYYY"
            c = ws.cell(row=linha, column=idx["OBJETO DO DFD"])
            c.alignment = Alignment(vertical="top", wrap_text=True)

        larguras = {
            "DFD": 12, "ORDEM": 8, "OBJETO DO DFD": 55, "AREA": 12,
            "PRIORIDADE": 11, "DATA PREVISTA FORMALIZACAO": 16,
            "QTD ITENS": 10, "VALOR TOTAL": 16, "CONTRATACAO": 40,
            "INICIO DO PROCESSO": 15, "FIM DO PROCESSO": 15,
            "DURACAO (DIAS)": 12, "ITEM PNCP": 12, "EM EXECUCAO": 12,
            "PUBLICACAO PNCP": 15, "UASG": 10,
        }
        for col, larg in larguras.items():
            ws.column_dimensions[get_column_letter(idx[col])].width = larg

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
