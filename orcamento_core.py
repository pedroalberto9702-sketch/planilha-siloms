#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
 ORCAMENTO - leitura das planilhas do Tesouro Gerencial
=============================================================================
 Dois extratos:

   CREDITO DISPONIVEL  uma aba por UG responsavel. Cabecalho na linha 6,
                       subtitulos na 7, dados a partir da 8. As abas NAO tem
                       as mesmas colunas (algumas nao trazem CREDITO
                       DISPONIVEL), e os campos hierarquicos (acao, fonte,
                       PTRES, PI) vem em branco quando repetem a linha acima.

   EMPENHOS            uma aba. Cabecalho na linha 3, subtitulos na 4, dados
                       a partir da 5. O identificador do favorecido precisa
                       ser lido como TEXTO: 6 digitos = UG, 11 = CPF,
                       14 = CNPJ. Lido como numero, perde o zero a esquerda.

 Por isso as colunas sao localizadas pelo NOME no cabecalho, nunca pela
 posicao: a extracao muda de um mes para o outro.
=============================================================================
"""

import re
import unicodedata

import pandas as pd


# =============================================================================
# AUXILIARES
# =============================================================================

def _chave(texto):
    """Normaliza um rotulo: sem acento, sem espaco duplo, maiusculo."""
    t = unicodedata.normalize("NFKD", str(texto or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().upper()


def _para_numero(serie):
    """Converte para float aceitando texto, formato brasileiro e vazio."""
    if serie.dtype.kind in "if":
        return serie.fillna(0.0).astype(float)

    txt = serie.astype(str).str.strip()
    tem_virgula = txt.str.contains(",", na=False)
    txt = txt.where(~tem_virgula,
                    txt.str.replace(".", "", regex=False)
                       .str.replace(",", ".", regex=False))
    txt = txt.str.replace(r"[^\d.\-]", "", regex=True)
    return pd.to_numeric(txt, errors="coerce").fillna(0.0)


def _achar_cabecalho(bruto, marcador, limite=15):
    """Devolve o indice (base 0) da linha de cabecalho que contem o marcador."""
    alvo = _chave(marcador)
    for i in range(min(limite, len(bruto))):
        for v in bruto.iloc[i].tolist():
            if alvo in _chave(v):
                return i
    raise ValueError(f"não encontrei o cabeçalho (procurando por “{marcador}”)")


def _mapa_colunas(linha):
    """{rotulo normalizado: indice} a partir de uma linha de cabecalho."""
    mapa = {}
    for i, v in enumerate(linha):
        k = _chave(v)
        if k and k != "NAN" and k not in mapa:
            mapa[k] = i
    return mapa


def _indice(mapa, *nomes, obrigatorio=True):
    """Procura o primeiro rotulo que comece com algum dos nomes dados."""
    for nome in nomes:
        alvo = _chave(nome)
        if alvo in mapa:
            return mapa[alvo]
    for nome in nomes:
        alvo = _chave(nome)
        for k, i in mapa.items():
            if k.startswith(alvo):
                return i
    if obrigatorio:
        raise ValueError(f"coluna “{nomes[0]}” não encontrada na planilha")
    return None


# =============================================================================
# CREDITO DISPONIVEL
# =============================================================================

def ler_credito(arquivo):
    """
    Le o extrato de Credito Disponivel (todas as abas) e devolve um DataFrame
    com uma linha por registro e a UG responsavel identificada.
    """
    abas = pd.read_excel(arquivo, sheet_name=None, header=None)
    if not abas:
        raise ValueError("a planilha não tem nenhuma aba")

    partes, problemas = [], []
    for nome_aba, bruto in abas.items():
        try:
            partes.append(_ler_aba_credito(nome_aba, bruto))
        except Exception as e:
            problemas.append(f"{nome_aba}: {e}")

    if not partes:
        raise ValueError(
            "nenhuma aba pôde ser lida. " + " | ".join(problemas[:3])
        )

    df = pd.concat(partes, ignore_index=True)
    df = df[~((df["recebido"] == 0) & (df["disponivel"] == 0)
              & (df["empenhadas"] == 0))]
    return df.reset_index(drop=True)


def _ler_aba_credito(nome_aba, bruto):
    i_cab = _achar_cabecalho(bruto, "Ação Governo")
    mapa = _mapa_colunas(bruto.iloc[i_cab].tolist())

    i_acao = _indice(mapa, "Ação Governo")
    i_fonte = _indice(mapa, "Fonte Recursos Detalhada", "Fonte")
    i_ptres = _indice(mapa, "PTRES")
    i_pi = _indice(mapa, "PI")
    i_nd = _indice(mapa, "Item Informação")
    i_receb = _indice(mapa, "Crédito Recebido")
    i_disp = _indice(mapa, "CREDITO DISPONIVEL", "Crédito Disponível",
                     obrigatorio=False)
    i_emp = _indice(mapa, "DESPESAS EMPENHADAS", obrigatorio=False)

    # a coluna logo depois do PI traz o nome do PI e nao tem rotulo
    i_pi_nome = i_pi + 1 if (i_pi + 1) not in mapa.values() else None

    # UG responsavel: aparece como "UG Responsável: NOME:CODIGO" acima do cabecalho
    ug_nome, ug_cod = nome_aba, ""
    for i in range(i_cab):
        for v in bruto.iloc[i].tolist():
            txt = str(v or "")
            if "UG RESPONS" in _chave(txt):
                bruto_ug = txt.split(":", 1)[-1].strip()
                m = re.search(r"(.*?):(\d+)\s*$", bruto_ug)
                if m:
                    ug_nome, ug_cod = m.group(1).strip(), m.group(2)
                else:
                    ug_nome = bruto_ug
                break

    dados = bruto.iloc[i_cab + 2:].reset_index(drop=True)

    out = pd.DataFrame({
        "ug_codigo": ug_cod,
        "ug_nome": ug_nome,
        "acao": dados.iloc[:, i_acao],
        "fonte": dados.iloc[:, i_fonte],
        "ptres": dados.iloc[:, i_ptres],
        "pi": dados.iloc[:, i_pi],
        "pi_nome": dados.iloc[:, i_pi_nome] if i_pi_nome is not None else "",
        "nd": dados.iloc[:, i_nd],
        "recebido": _para_numero(dados.iloc[:, i_receb]),
        "disponivel": (_para_numero(dados.iloc[:, i_disp])
                       if i_disp is not None else 0.0),
        "empenhadas": (_para_numero(dados.iloc[:, i_emp])
                       if i_emp is not None else 0.0),
    })

    # os campos hierarquicos so aparecem na primeira linha do grupo
    for col in ("acao", "fonte", "ptres", "pi", "pi_nome"):
        out[col] = (out[col].astype("object").where(out[col].notna())
                    .ffill().fillna(""))
        out[col] = out[col].astype(str).str.strip().replace("nan", "")

    out["nd"] = (out["nd"].astype(str).str.replace(r"\D", "", regex=True)
                 .str[:6])
    return out[out["nd"].str.len() > 0]


def resumo_credito(df):
    """Numeros de topo e quebras por UG e por natureza de despesa."""
    por_ug = (df.groupby(["ug_codigo", "ug_nome"], as_index=False)
                [["recebido", "disponivel", "empenhadas"]].sum()
                .sort_values("recebido", ascending=False))
    por_nd = (df.groupby("nd", as_index=False)
                [["recebido", "disponivel", "empenhadas"]].sum()
                .sort_values("recebido", ascending=False))
    por_acao = (df[df["acao"] != ""].groupby("acao", as_index=False)
                  [["recebido", "disponivel", "empenhadas"]].sum()
                  .sort_values("recebido", ascending=False))
    return {
        "recebido": float(df["recebido"].sum()),
        "disponivel": float(df["disponivel"].sum()),
        "empenhadas": float(df["empenhadas"].sum()),
        "linhas": int(len(df)),
        "ugs": int(df["ug_nome"].nunique()),
        "por_ug": por_ug.to_dict("records"),
        "por_nd": por_nd.to_dict("records"),
        "por_acao": por_acao.head(15).to_dict("records"),
    }


# =============================================================================
# EMPENHOS
# =============================================================================

def ler_empenhos(arquivo):
    """Le o extrato de Empenhos e devolve um DataFrame por nota de empenho."""
    bruto = pd.read_excel(arquivo, sheet_name=0, header=None, dtype=object)

    i_cab = _achar_cabecalho(bruto, "NE CCor")
    mapa = _mapa_colunas(bruto.iloc[i_cab].tolist())

    i_ne = _indice(mapa, "NE CCor")
    i_dia = _indice(mapa, "NE CCor - Dia Emissão", "Dia Emissão")
    i_fav = _indice(mapa, "NE CCor - Favorecido", "Favorecido")
    i_desc = _indice(mapa, "NE CCor - Descrição", "Descrição")
    i_nd = _indice(mapa, "Natureza Despesa Detalhada", "Natureza Despesa")
    # "DESPESAS EMPENHADAS" e prefixo de "DESPESAS EMPENHADAS A LIQUIDAR":
    # procuramos o rotulo mais longo primeiro.
    i_aliq = _indice(mapa, "DESPESAS EMPENHADAS A LIQUIDAR")
    i_liq = _indice(mapa, "DESPESAS LIQUIDADAS")
    i_emp = next((i for k, i in mapa.items()
                  if k.startswith("DESPESAS EMPENHADAS") and i not in (i_aliq,)),
                 None)
    if i_emp is None:
        raise ValueError("coluna “DESPESAS EMPENHADAS” não encontrada")

    # a modalidade fica na linha de subtitulo, logo abaixo de "Item Informação"
    i_modal = _indice(mapa, "Item Informação", obrigatorio=False)

    dados = bruto.iloc[i_cab + 2:].reset_index(drop=True)

    def texto(col, repetir=False):
        """
        Coluna como texto. Com repetir=True, herda o valor da linha de cima
        quando vem vazia: uma mesma NE com duas naturezas de despesa gera
        linhas de continuacao em que so a ND e os valores sao preenchidos.
        """
        s = dados.iloc[:, col]
        s = s.where(s.notna())
        if repetir:
            s = s.ffill()
        return s.astype("object").fillna("").astype(str).str.strip()

    out = pd.DataFrame({
        "ne": texto(i_ne, repetir=True),
        "emissao": texto(i_dia, repetir=True),
        "favorecido_ni": texto(i_fav, repetir=True),
        "favorecido_nome": texto(i_fav + 1, repetir=True),
        "descricao": texto(i_desc, repetir=True),
        "nd_detalhada": texto(i_nd),
        "nd_nome": texto(i_nd + 1),
        "modalidade": texto(i_modal) if i_modal is not None else "",
        "empenhado": _para_numero(dados.iloc[:, i_emp]),
        "a_liquidar": _para_numero(dados.iloc[:, i_aliq]),
        "liquidado": _para_numero(dados.iloc[:, i_liq]),
    })

    out = out[out["ne"].str.len() > 0].reset_index(drop=True)
    out["favorecido_ni"] = out["favorecido_ni"].str.replace(r"\D", "", regex=True)
    out["nd"] = out["nd_detalhada"].str.replace(r"\D", "", regex=True).str[:6]

    # 14 digitos = CNPJ (fornecedor, pode ter contrato);
    # 11 = CPF; menos = a propria UG (folha, auxilios, diarias)
    tam = out["favorecido_ni"].str.len()
    out["tipo_favorecido"] = tam.map(
        lambda n: "CNPJ" if n == 14 else ("CPF" if n == 11 else "UG")
    )
    return out


def resumo_empenhos(df):
    """Numeros de topo e quebras por ND, modalidade e fornecedor."""
    medidas = ["empenhado", "a_liquidar", "liquidado"]

    por_nd = (df.groupby(["nd"], as_index=False)
                .agg(empenhado=("empenhado", "sum"),
                     a_liquidar=("a_liquidar", "sum"),
                     liquidado=("liquidado", "sum"),
                     nes=("ne", "nunique"),
                     nome=("nd_nome", "first"))
                .sort_values("empenhado", ascending=False))

    por_mod = (df.groupby("modalidade", as_index=False)
                 .agg(empenhado=("empenhado", "sum"),
                      a_liquidar=("a_liquidar", "sum"),
                      liquidado=("liquidado", "sum"),
                      nes=("ne", "count"))
                 .sort_values("empenhado", ascending=False))

    pj = df[df["tipo_favorecido"] == "CNPJ"]
    por_forn = (pj.groupby("favorecido_ni", as_index=False)
                  .agg(nome=("favorecido_nome", "first"),
                       empenhado=("empenhado", "sum"),
                       a_liquidar=("a_liquidar", "sum"),
                       liquidado=("liquidado", "sum"),
                       nes=("ne", "count"))
                  .sort_values("empenhado", ascending=False))

    # fornecedor com CNPJ pode estar ligado a contrato; a propria UG, nunca
    forn = float(pj["empenhado"].sum())
    outros = float(df[df["tipo_favorecido"] != "CNPJ"]["empenhado"].sum())

    return {
        "empenhado": float(df["empenhado"].sum()),
        "a_liquidar": float(df["a_liquidar"].sum()),
        "liquidado": float(df["liquidado"].sum()),
        "nes": int(df["ne"].nunique()),
        "fornecedores": int(pj["favorecido_ni"].nunique()),
        "valor_fornecedores": forn,
        "valor_proprios": outros,
        "nes_fornecedores": int(pj["ne"].nunique()),
        "nes_proprios": int(df["ne"].nunique() - pj["ne"].nunique()),
        "por_nd": por_nd.to_dict("records"),
        "por_modalidade": por_mod.to_dict("records"),
        "por_fornecedor": por_forn.head(25).to_dict("records"),
    }
