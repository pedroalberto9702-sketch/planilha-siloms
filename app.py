#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
 Painel de Certames — acompanhamento de contratações públicas
=============================================================================
 Carrega os certames de uma unidade e gera a planilha de resultado de
 qualquer um deles.

 A lista fica guardada em memória depois da primeira busca, então gerar a
 planilha de um certame não refaz a consulta inteira. O botão Atualizar
 descarta o que está guardado e busca de novo.

 Arquivos:
   comprasnet.py   cliente da API
   planilha.py     montagem do .xlsx
   gunicorn.conf.py  limites do servidor (o render.yaml não vale para
                     serviços criados pelo painel do Render)
=============================================================================
"""

import os
import re
import threading
import time
from datetime import date, datetime

from flask import Flask, jsonify, render_template_string, request, send_file

import comprasnet as api
import planilha as pl

VERSAO = "2.1.0"

app = Flask(__name__)
app.config["JSON_SORT_KEYS"] = False

# ---- memória da sessão do servidor -----------------------------------------
# {(uasg, ano): {"certames": [...], "em": timestamp}}
_CACHE = {}
_TRAVA = threading.Lock()
VALIDADE = 30 * 60          # meia hora; o botão Atualizar ignora isso


# =============================================================================
# PÁGINA
# =============================================================================

PAGINA = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Painel de Certames</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{
  --papel:#EDF1F4;
  --carta:#FFFFFF;
  --tinta:#13232E;
  --tinta2:#5C7183;
  --tinta3:#8698A6;
  --linha:#D3DDE4;
  --linha2:#E7EDF1;
  --acao:#0E6A6F;
  --acao-forte:#0A5155;
  --acao-fraco:#DCEBEB;
  --alerta:#A44A22;
  --alerta-fraco:#F7EBE4;
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{
  background:var(--papel);
  color:var(--tinta);
  font-family:Archivo,system-ui,sans-serif;
  font-size:15px;line-height:1.5;
  font-variant-numeric:tabular-nums;
  -webkit-font-smoothing:antialiased;
}
:focus-visible{outline:2px solid var(--acao);outline-offset:2px}

/* ---------- barra de controle ---------- */
.controle{
  position:sticky;top:0;z-index:30;
  background:var(--tinta);color:#fff;
  padding:16px 22px;
}
.controle .dentro{
  max-width:1120px;margin:0 auto;
  display:flex;align-items:flex-end;gap:18px;flex-wrap:wrap;
}
.marca{margin-right:auto}
.marca h1{
  margin:0;font-size:21px;font-weight:700;letter-spacing:-.015em;
}
.marca p{
  margin:3px 0 0;font-size:13px;color:#9FB4C0;font-weight:400;
}
.campo{display:flex;flex-direction:column;gap:5px}
.campo label{font-size:12px;font-weight:500;color:#9FB4C0}
.campo input{
  font-family:Archivo,sans-serif;font-size:15px;font-weight:500;
  font-variant-numeric:tabular-nums;
  color:#fff;background:rgba(255,255,255,.08);
  border:1px solid rgba(255,255,255,.22);border-radius:4px;
  padding:8px 11px;width:112px;
}
.campo input:focus{outline:none;border-color:#6FC3C6;background:rgba(255,255,255,.14)}
.campo input.erro{border-color:#E4A08A;background:rgba(228,160,138,.14)}
#ano{width:84px}

.atualizar{
  font-family:Archivo,sans-serif;font-size:14px;font-weight:600;color:#fff;
  background:var(--acao);border:1px solid #17888E;border-radius:4px;
  padding:9px 18px;cursor:pointer;
}
.atualizar:hover:not(:disabled){background:#12807F}
.atualizar:disabled{opacity:.55;cursor:progress}

/* ---------- conteúdo ---------- */
.folha{max-width:1120px;margin:0 auto;padding:22px}

.resumo{
  display:flex;align-items:baseline;gap:14px;flex-wrap:wrap;
  margin-bottom:14px;font-size:13.5px;color:var(--tinta2);
}
.resumo b{color:var(--tinta);font-weight:600;font-size:15px}

.filtros{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:16px}
.filtro{
  font-family:Archivo,sans-serif;font-size:13px;font-weight:500;
  color:var(--tinta2);background:transparent;
  border:1px solid var(--linha);border-radius:999px;
  padding:5px 13px;cursor:pointer;
}
.filtro:hover{border-color:var(--tinta3);color:var(--tinta)}
.filtro.ativo{background:var(--tinta);border-color:var(--tinta);color:#fff}
.filtro .conta{opacity:.6;margin-left:5px;font-weight:400}

/* ---------- lista ---------- */
.lista{
  background:var(--carta);border:1px solid var(--linha);border-radius:6px;
  overflow:hidden;
}
.certame{
  display:grid;
  grid-template-columns:minmax(0,1fr) 190px 150px;
  gap:18px;align-items:center;
  padding:15px 20px;border-bottom:1px solid var(--linha2);
}
.certame:last-child{border-bottom:none}
.certame:hover{background:#F8FBFC}

.identidade{min-width:0}
.identidade .nome{
  font-size:15.5px;font-weight:600;letter-spacing:-.01em;margin-bottom:3px;
}
.identidade .objeto{
  font-size:13px;color:var(--tinta2);
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
}
.identidade .datas{
  font-size:12.5px;color:var(--tinta3);margin-top:4px;
}
.selo{
  display:inline-block;font-size:11px;font-weight:600;
  background:var(--acao-fraco);color:var(--acao-forte);
  border-radius:3px;padding:1px 6px;margin-left:7px;vertical-align:1px;
}

/* trilha de fase: a posição carrega a informação */
.fase .nome{font-size:13.5px;font-weight:600;margin-bottom:6px}
.fase .trilha{display:flex;gap:2px}
.fase .trilha i{
  display:block;height:5px;flex:1;border-radius:1px;background:var(--linha);
}
.fase .trilha i.feito{background:var(--acao)}
.fase.excecao .nome{color:var(--alerta)}
.fase.excecao .trilha i{background:var(--alerta-fraco)}
.fase.excecao .trilha i:first-child{background:var(--alerta)}

.gerar{
  font-family:Archivo,sans-serif;font-size:13.5px;font-weight:600;
  color:var(--acao-forte);background:#fff;
  border:1px solid var(--acao);border-radius:4px;
  padding:8px 12px;cursor:pointer;width:100%;
}
.gerar:hover:not(:disabled){background:var(--acao-fraco)}
.gerar:disabled{opacity:.5;cursor:progress;border-color:var(--linha);color:var(--tinta3)}
.gerar.pronto{background:var(--acao);border-color:var(--acao);color:#fff}

.recado{
  grid-column:1 / -1;font-size:13px;color:var(--tinta2);
  background:#F6F9FA;border-left:3px solid var(--acao);
  padding:9px 12px;border-radius:0 3px 3px 0;margin-top:4px;
}
.recado.ruim{border-left-color:var(--alerta);background:var(--alerta-fraco);color:#7A3717}

/* ---------- estados ---------- */
.vazio{
  background:var(--carta);border:1px solid var(--linha);border-radius:6px;
  padding:44px 26px;text-align:center;
}
.vazio h2{margin:0 0 7px;font-size:17px;font-weight:600}
.vazio p{margin:0;font-size:14px;color:var(--tinta2);max-width:460px;
         margin-left:auto;margin-right:auto}
.vazio.falhou h2{color:var(--alerta)}

.carregando{
  background:var(--carta);border:1px solid var(--linha);border-radius:6px;
  padding:30px 26px;
}
.barra{height:3px;background:var(--linha2);border-radius:2px;overflow:hidden}
.barra i{display:block;height:100%;width:34%;background:var(--acao);
         border-radius:2px;animation:desliza 1.3s ease-in-out infinite}
@keyframes desliza{0%{margin-left:-34%}100%{margin-left:100%}}
.carregando p{margin:12px 0 0;font-size:13.5px;color:var(--tinta2);text-align:center}
@media (prefers-reduced-motion:reduce){
  .barra i{animation:none;width:100%}
}

.nota{
  max-width:1120px;margin:16px auto 34px;padding:0 22px;
  font-size:12.5px;color:var(--tinta3);line-height:1.6;
}
.nota a{color:var(--acao)}

@media (max-width:860px){
  .certame{grid-template-columns:1fr;gap:12px}
  .fase{max-width:260px}
  .gerar{width:auto;padding:8px 16px}
  .identidade .objeto{white-space:normal}
}
@media (max-width:560px){
  .controle{padding:14px 16px}
  .controle .dentro{gap:12px}
  .marca{flex:1 0 100%}
  .folha{padding:16px}
}
</style>
</head>
<body>

<header class="controle">
  <div class="dentro">
    <div class="marca">
      <h1>Painel de Certames</h1>
      <p>Contratações publicadas no PNCP</p>
    </div>
    <div class="campo">
      <label for="uasg">Unidade (UASG)</label>
      <input type="text" id="uasg" inputmode="numeric" maxlength="6" value="">
    </div>
    <div class="campo">
      <label for="ano">Ano</label>
      <input type="text" id="ano" inputmode="numeric" maxlength="4" value="">
    </div>
    <button class="atualizar" id="btn">Atualizar</button>
  </div>
</header>

<main class="folha">
  <div id="resumo" class="resumo" hidden></div>
  <div id="filtros" class="filtros" hidden></div>
  <div id="palco"></div>
</main>

<p class="nota">
  A fase de cada certame é deduzida das datas de proposta e da existência de
  resultado publicados no PNCP. A API de dados abertos não expõe o andamento
  interno do Compras.gov.br, então use esta coluna como indicação, não como
  situação oficial. Dados do
  <a href="https://dadosabertos.compras.gov.br" target="_blank" rel="noopener">Compras.gov.br</a>,
  com atraso de alguns dias em relação ao sistema.
  <span id="versao"></span>
</p>

<script>
const $ = s => document.querySelector(s);
const palco = $('#palco'), elResumo = $('#resumo'), elFiltros = $('#filtros');
const campoUasg = $('#uasg'), campoAno = $('#ano'), btn = $('#btn');

let certames = [], faltaram = [], filtro = 'Todos';

const esc = t => String(t ?? '').replace(/[&<>"]/g, c =>
  ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const brl = n => (n||0).toLocaleString('pt-BR',
  {style:'currency',currency:'BRL',minimumFractionDigits:2});
const dia = s => s ? s.split('-').reverse().slice(0,2).join('/') : '';

const FASES = ['Aguardando abertura','Em disputa','Em habilitação',
               'Em homologação','Homologado'];

// ---- memória do navegador: evita redigitar a unidade todo dia ----
try{
  campoUasg.value = localStorage.getItem('uasg') || '';
}catch(e){}
campoAno.value = String(new Date().getFullYear());

function trilha(c){
  if(c.fase_excecao){
    return `<div class="fase excecao"><div class="nome">${esc(c.fase)}</div>
      <div class="trilha">${'<i></i>'.repeat(5)}</div></div>`;
  }
  const n = c.fase_ordem || 0;
  let seg = '';
  for(let i=1;i<=5;i++) seg += `<i class="${i<=n?'feito':''}"></i>`;
  return `<div class="fase"><div class="nome">${esc(c.fase)}</div>
    <div class="trilha">${seg}</div></div>`;
}

function datas(c){
  const p = [];
  if(c.abertura)     p.push('abertura ' + dia(c.abertura));
  if(c.encerramento) p.push('propostas até ' + dia(c.encerramento));
  if(!p.length && c.publicacao) p.push('publicado em ' + dia(c.publicacao));
  if(c.valor_homologado > 0) p.push('homologado ' + brl(c.valor_homologado));
  return p.join(', ');
}

function desenhar(){
  const vis = filtro === 'Todos' ? certames : certames.filter(c => c.fase === filtro);

  if(!certames.length){
    elResumo.hidden = elFiltros.hidden = true;
    return;
  }

  const atual = new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'});
  elResumo.hidden = false;
  elResumo.innerHTML = `<span><b>${certames.length}</b> certames na unidade
      ${esc(campoUasg.value)} em ${esc(campoAno.value)}</span>
    <span>atualizado às ${atual}</span>`
    + (faltaram.length ? `<span style="color:var(--alerta)">
        ${esc(faltaram.join(', '))} não respondeu a tempo</span>` : '');

  const contas = {};
  certames.forEach(c => contas[c.fase] = (contas[c.fase]||0)+1);
  const ordem = FASES.filter(f => contas[f]);
  Object.keys(contas).forEach(f => { if(!ordem.includes(f)) ordem.push(f); });

  elFiltros.hidden = false;
  elFiltros.innerHTML = [['Todos', certames.length]]
    .concat(ordem.map(f => [f, contas[f]]))
    .map(([f,n]) => `<button class="filtro ${f===filtro?'ativo':''}"
        onclick="trocarFiltro('${f.replace(/'/g,"\\'")}')">${esc(f)}
        <span class="conta">${n}</span></button>`).join('');

  if(!vis.length){
    palco.innerHTML = `<div class="vazio"><h2>Nenhum certame nesta fase</h2>
      <p>Escolha outra fase ou volte para Todos.</p></div>`;
    return;
  }

  palco.innerHTML = '<div class="lista">' + vis.map((c,i) => `
    <div class="certame" id="cert-${i}">
      <div class="identidade">
        <div class="nome">${esc(c.titulo)}${c.srp?'<span class="selo">SRP</span>':''}</div>
        <div class="objeto" title="${esc(c.objeto)}">${esc(c.objeto || 'Sem objeto informado')}</div>
        <div class="datas">${esc(datas(c))}</div>
      </div>
      ${trilha(c)}
      <button class="gerar" onclick="gerar(${i})">Gerar planilha</button>
    </div>`).join('') + '</div>';
}

function trocarFiltro(f){ filtro = f; desenhar(); }

function mostrarCarregando(texto){
  elResumo.hidden = elFiltros.hidden = true;
  palco.innerHTML = `<div class="carregando">
    <div class="barra"><i></i></div><p>${esc(texto)}</p></div>`;
}

// ---- lista de certames ----
async function atualizar(forcar){
  const uasg = campoUasg.value.replace(/\D/g,'');
  const ano  = campoAno.value.replace(/\D/g,'');
  campoUasg.classList.toggle('erro', !(uasg.length>=5 && uasg.length<=6));
  campoAno.classList.toggle('erro', ano.length!==4);
  if(uasg.length<5 || uasg.length>6){ campoUasg.focus(); return; }
  if(ano.length!==4){ campoAno.focus(); return; }
  try{ localStorage.setItem('uasg', uasg); }catch(e){}

  btn.disabled = true; btn.textContent = 'Buscando...';
  mostrarCarregando('Consultando o Compras.gov.br. Na primeira busca do dia o '
    + 'servidor precisa acordar, o que leva até um minuto.');
  certames = [];

  try{
    const r = await fetch(`/api/certames?uasg=${uasg}&ano=${ano}`
                          + (forcar ? '&forcar=1' : ''));
    const d = await r.json().catch(() => ({}));
    if(!r.ok || !d.ok){
      palco.innerHTML = `<div class="vazio falhou">
        <h2>${esc(d.titulo || 'Não consegui buscar os certames')}</h2>
        <p>${esc(d.mensagem || 'Tente de novo em alguns instantes.')}</p></div>`;
      return;
    }
    certames = d.certames || [];
    faltaram = d.faltaram || [];
    if(!certames.length){
      palco.innerHTML = `<div class="vazio">
        <h2>Nenhum certame encontrado</h2>
        <p>A unidade ${esc(uasg)} não tem contratações publicadas no PNCP em
           ${esc(ano)}. Confira o número da unidade e o ano.</p></div>`;
      return;
    }
    filtro = 'Todos';
    desenhar();
  }catch(err){
    palco.innerHTML = `<div class="vazio falhou"><h2>A conexão caiu</h2>
      <p>Verifique sua internet e clique em Atualizar.</p></div>`;
  }finally{
    btn.disabled = false; btn.textContent = 'Atualizar';
  }
}

// ---- planilha de um certame ----
async function gerar(i){
  const c = certames[i];
  const linha = document.getElementById('cert-'+i);
  const b = linha.querySelector('.gerar');
  linha.querySelectorAll('.recado').forEach(e => e.remove());

  b.disabled = true; b.textContent = 'Gerando...';
  try{
    const r = await fetch('/api/planilha', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({uasg:c.uasg, idCompra:c.idCompra,
                            publicacao:c.publicacao, titulo:c.titulo})
    });

    if(r.ok && (r.headers.get('content-type')||'').includes('spreadsheet')){
      const blob = await r.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = (r.headers.get('x-arquivo') || 'resultado.xlsx');
      document.body.appendChild(a); a.click(); a.remove();
      URL.revokeObjectURL(url);
      b.textContent = 'Baixada';
      b.classList.add('pronto');
      const info = r.headers.get('x-resumo');
      if(info) recado(linha, info, false);
      setTimeout(() => { b.textContent='Gerar planilha'; b.classList.remove('pronto');
                         b.disabled=false; }, 4000);
      return;
    }

    const d = await r.json().catch(() => ({}));
    recado(linha, (d.titulo ? d.titulo + '. ' : '') +
                  (d.mensagem || 'Não consegui gerar a planilha.'), true);
  }catch(err){
    recado(linha, 'A conexão caiu durante a geração. Tente de novo.', true);
  }finally{
    if(b.textContent === 'Gerando...'){ b.disabled=false; b.textContent='Gerar planilha'; }
  }
}

function recado(linha, texto, ruim){
  const d = document.createElement('div');
  d.className = 'recado' + (ruim ? ' ruim' : '');
  d.textContent = texto;
  linha.appendChild(d);
}

btn.addEventListener('click', () => atualizar(true));
[campoUasg, campoAno].forEach(el =>
  el.addEventListener('keydown', e => { if(e.key==='Enter') atualizar(true); }));

palco.innerHTML = `<div class="vazio"><h2>Escolha a unidade</h2>
  <p>Informe o código da UASG e o ano, depois clique em Atualizar para carregar
     os certames publicados no PNCP.</p></div>`;
if(campoUasg.value) atualizar(false);

fetch('/versao').then(r=>r.json()).then(d=>{
  $('#versao').textContent = ' Versão ' + d.versao + '.';
}).catch(()=>{});
</script>
</body>
</html>
"""


# =============================================================================
# ROTAS
# =============================================================================

@app.errorhandler(Exception)
def erro_inesperado(e):
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return e
    import traceback
    traceback.print_exc()
    tb = traceback.extract_tb(e.__traceback__)
    onde = ""
    if tb:
        u = tb[-1]
        onde = f" ({os.path.basename(u.filename)}, linha {u.lineno})"
    return jsonify(
        ok=False,
        titulo="Erro inesperado no servidor",
        mensagem=f"{type(e).__name__}: {e}{onde}",
    ), 200


@app.route("/")
def inicio():
    return render_template_string(PAGINA)


@app.route("/versao")
def versao():
    return jsonify(
        versao=VERSAO,
        cliente=api.VERSAO,
        tempo_limite=api.TEMPO_LIMITE,
        paralelo=api.PARALELO,
        tamanho_pagina=api.TAMANHO_PAGINA,
        em_cache=sorted(f"{u}/{a}" for u, a in _CACHE),
    )


def _validar(uasg, ano):
    u = re.sub(r"\D", "", str(uasg or ""))
    a = re.sub(r"\D", "", str(ano or ""))
    if not (5 <= len(u) <= 6):
        return None, None, "Informe a UASG com 5 ou 6 dígitos."
    if len(a) != 4 or not (2000 <= int(a) <= 2100):
        return None, None, "Informe o ano com 4 dígitos."
    return u, int(a), None


@app.route("/api/certames")
def rota_certames():
    uasg, ano, erro = _validar(request.args.get("uasg"), request.args.get("ano"))
    if erro:
        return jsonify(ok=False, titulo="Dados incompletos", mensagem=erro)

    forcar = request.args.get("forcar") == "1"
    chave = (uasg, ano)

    if not forcar:
        with _TRAVA:
            guardado = _CACHE.get(chave)
        if guardado and (time.time() - guardado["em"]) < VALIDADE:
            return jsonify(ok=True, certames=guardado["certames"],
                           faltaram=guardado.get("faltaram") or [], do_cache=True)

    api.iniciar_prazo()
    try:
        certames, faltaram = api.listar_certames(uasg, ano)
    except api.TempoEsgotado as e:
        return jsonify(
            ok=False, titulo="A consulta demorou demais",
            mensagem=f"O Compras.gov.br não respondeu a tempo ({e}). "
                     "Tente de novo em alguns minutos.",
        )
    finally:
        api.limpar_prazo()

    # resultado parcial vale mais que erro: guardamos e avisamos o que faltou
    with _TRAVA:
        _CACHE[chave] = {"certames": certames, "faltaram": faltaram,
                         "em": time.time()}
        for k in [k for k in _CACHE if k != chave][:-4]:
            _CACHE.pop(k, None)

    return jsonify(ok=True, certames=certames, faltaram=faltaram, do_cache=False)


@app.route("/api/planilha", methods=["POST"])
def rota_planilha():
    dados = request.get_json(silent=True) or {}
    uasg = re.sub(r"\D", "", str(dados.get("uasg") or ""))
    id_compra = str(dados.get("idCompra") or "").strip()
    if not uasg or not id_compra:
        return jsonify(ok=False, titulo="Certame não identificado",
                       mensagem="Atualize a lista e tente de novo.")

    api.iniciar_prazo()
    try:
        itens, resultados = api.resultado_certame(
            uasg, id_compra, dados.get("publicacao"))
    except api.TempoEsgotado as e:
        return jsonify(
            ok=False, titulo="A consulta demorou demais",
            mensagem=f"O Compras.gov.br não respondeu a tempo ({e}). "
                     "Tente de novo em alguns minutos.",
        )
    finally:
        api.limpar_prazo()

    if not itens and not resultados:
        return jsonify(
            ok=False, titulo="Sem itens publicados",
            mensagem="O PNCP ainda não publicou os itens deste certame. "
                     "Isso costuma levar alguns dias após a divulgação.",
        )

    df = pl.montar(itens, resultados)
    if df.empty:
        return jsonify(ok=False, titulo="Nada para montar",
                       mensagem="Os itens vieram sem identificação.")

    r = pl.resumo(df)
    nome = f"resultado_{id_compra}.xlsx"
    # a troca de separadores vale só para o número, não para a frase
    valor = f"{r['valor']:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    aviso = (f"{r['itens']} itens, {r['com_vencedor']} com vencedor, "
             f"{r['sem_vencedor']} desertos ou fracassados. "
             f"Total homologado R$ {valor}.")

    resposta = send_file(
        pl.gerar_xlsx(df), as_attachment=True, download_name=nome,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resposta.headers["x-arquivo"] = nome
    resposta.headers["x-resumo"] = aviso
    return resposta


# =============================================================================
# EXECUÇÃO LOCAL
# =============================================================================

if __name__ == "__main__":
    porta = int(os.environ.get("PORT", 5000))
    print(f"Painel de Certames {VERSAO} — http://127.0.0.1:{porta}")
    app.run(host="0.0.0.0", port=porta, debug=False)
