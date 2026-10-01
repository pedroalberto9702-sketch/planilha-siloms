#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
 GERADOR TOP COMPRASNET - versao WEB
=============================================================================
 Sobe um site local com o formulario (UASG, tipo, processo) e devolve o
 arquivo .xlsx pronto para download.

 REQUISITOS:
     - gerar_planilha_siloms.py na MESMA PASTA (contem toda a logica)
     - pip install flask requests pandas openpyxl

 USO:
     Duplo-clique em ABRIR_SITE.bat
     ou:  py app.py
     e acesse http://127.0.0.1:5000 no navegador.
=============================================================================
"""

import io
import os
import time
import re
import sys
import uuid

import pandas as pd
import tempfile
import threading
import webbrowser
from datetime import datetime

try:
    from flask import Flask, request, jsonify, send_file, render_template_string
except ImportError:
    print("!! Flask nao instalado. Rode:  py -m pip install flask")
    sys.exit(1)

# --- importa a logica do script de linha de comando --------------------------
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import gerar_planilha_siloms as core
except ImportError:
    print("!! Arquivo gerar_planilha_siloms.py nao encontrado nesta pasta.")
    print("   Os dois arquivos precisam ficar juntos.")
    sys.exit(1)

core.DEBUG = False

try:
    import pgc_core
except ImportError:
    pgc_core = None

try:
    import orcamento_core
except ImportError:
    orcamento_core = None

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 24 * 1024 * 1024

PASTA_SAIDA = os.path.join(tempfile.gettempdir(), "siloms_planilhas")
os.makedirs(PASTA_SAIDA, exist_ok=True)

# guarda os arquivos gerados nesta sessao: token -> (caminho, nome_exibicao)
ARQUIVOS = {}

# nomes com acento para exibir na tela
NOMES_TIPO = {
    "PREGAO": "Pregão",
    "DISPENSA": "Dispensa de licitação",
    "CONCORRENCIA": "Concorrência",
    "INEXIGIBILIDADE": "Inexigibilidade",
    "CREDENCIAMENTO": "Credenciamento",
}


# =============================================================================
# TEMPLATE
# =============================================================================

PAGINA = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Gerador TOP Comprasnet</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{
  --tinta:#0B1F33;
  --tinta-fraca:#5A6B7C;
  --papel:#EDF0F4;
  --campo:#FFFFFF;
  --azul:#1E5AA8;
  --azul-escuro:#153F75;
  --verde:#1B6E52;
  --verde-dado:#108A5E;
  --vermelho:#A8321F;
  --linha:#C8D2DD;
  --sombra:0 1px 2px rgba(11,31,51,.06),0 8px 24px rgba(11,31,51,.08);
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{
  background:var(--papel);
  color:var(--tinta);
  font-family:"IBM Plex Sans",system-ui,sans-serif;
  font-size:15px;line-height:1.5;
  padding:32px 16px 64px;
  -webkit-font-smoothing:antialiased;
}
.folha{
  max-width:720px;margin:0 auto;transition:max-width .18s ease;
  background:var(--campo);
  border:1px solid var(--linha);
  border-radius:3px;
  box-shadow:var(--sombra);
  overflow:hidden;
}

/* -------- cabecalho institucional -------- */
.tarja{
  background:var(--tinta);color:#fff;
  padding:20px 28px 18px;
  border-bottom:3px solid var(--azul);
}
.tarja .orgao{
  font-family:"IBM Plex Mono",monospace;
  font-size:11px;letter-spacing:.14em;text-transform:uppercase;
  color:#8FA6BC;margin-bottom:10px;
}
.tarja h1{
  margin:0;font-size:23px;font-weight:600;letter-spacing:-.01em;
}
.tarja .sub{
  margin-top:6px;font-size:13px;color:#A9BDD0;
}

.folha.larga{max-width:1020px}

/* -------- abas -------- */
.abas{
  display:flex;border-bottom:1px solid var(--linha);background:#F7F9FB;
  overflow-x:auto;-webkit-overflow-scrolling:touch;
}
.aba{
  flex:1 0 auto;padding:14px 16px;white-space:nowrap;border:none;background:none;cursor:pointer;
  font-family:"IBM Plex Sans",sans-serif;font-size:13.5px;font-weight:600;
  color:var(--tinta-fraca);border-bottom:3px solid transparent;
  transition:color .15s,border-color .15s;width:auto;border-radius:0;
}
.aba:hover{color:var(--tinta);background:none}
.aba.ativa{color:var(--azul);border-bottom-color:var(--azul);background:#fff}
.aba:disabled{background:none;cursor:not-allowed;opacity:.5}
.painel{display:none}
.painel.ativo{display:block}

/* -------- formulario -------- */
form{padding:28px}
.campo{margin-bottom:22px}
.campo:last-of-type{margin-bottom:26px}
label{
  display:block;
  font-family:"IBM Plex Mono",monospace;
  font-size:11px;font-weight:500;
  letter-spacing:.12em;text-transform:uppercase;
  color:var(--tinta-fraca);
  margin-bottom:7px;
}
.dica{
  font-family:"IBM Plex Sans",sans-serif;
  text-transform:none;letter-spacing:0;font-size:12px;
  color:var(--tinta-fraca);margin-top:6px;
}
input[type=text],select{
  width:100%;
  font-family:"IBM Plex Mono",monospace;
  font-size:16px;
  color:var(--tinta);
  background:var(--campo);
  border:1px solid var(--linha);
  border-radius:2px;
  padding:11px 13px;
}
select{font-family:"IBM Plex Sans",sans-serif;font-size:15px;cursor:pointer}
input:focus,select:focus{
  outline:none;border-color:var(--azul);
  box-shadow:0 0 0 3px rgba(30,90,168,.14);
}
input.erro{border-color:var(--vermelho);box-shadow:0 0 0 3px rgba(168,50,31,.12)}
.msg-erro{
  color:var(--vermelho);font-size:13px;margin-top:6px;display:none;
}

button{
  width:100%;
  font-family:"IBM Plex Sans",sans-serif;
  font-size:15px;font-weight:600;
  color:#fff;background:var(--azul);
  border:none;border-radius:2px;
  padding:14px;cursor:pointer;
  transition:background .15s;
}
button:hover:not(:disabled){background:var(--azul-escuro)}
button:disabled{background:#9BAABA;cursor:progress}
button:focus-visible{outline:3px solid var(--tinta);outline-offset:2px}

/* -------- comprovante -------- */
.saida{display:none;border-top:1px solid var(--linha);padding:26px 28px 30px}
.saida.ativa{display:block}
.carimbo{
  font-family:"IBM Plex Mono",monospace;
  font-size:11px;letter-spacing:.14em;text-transform:uppercase;
  color:var(--tinta-fraca);margin-bottom:14px;
}
.linha-dado{
  display:flex;justify-content:space-between;align-items:baseline;
  gap:12px;padding:9px 0;border-bottom:1px dotted var(--linha);
}
.linha-dado .rot{font-size:13px;color:var(--tinta-fraca)}
.linha-dado .val{
  font-family:"IBM Plex Mono",monospace;font-size:14px;font-weight:500;
  text-align:right;word-break:break-word;
}
.linha-dado .val.destaque{color:var(--verde);font-weight:600;font-size:15px}
.baixar{
  display:block;margin-top:22px;text-align:center;text-decoration:none;
  font-family:"IBM Plex Sans",sans-serif;font-weight:600;font-size:15px;
  color:#fff;background:var(--verde);border-radius:2px;padding:14px;
}
.baixar:hover{filter:brightness(1.1)}
.aviso{
  margin-top:16px;padding:12px 14px;border-radius:2px;font-size:13px;
  background:#FDF3E7;border-left:3px solid #C8862B;color:#6B4A12;
}
.falha{
  display:none;border-top:1px solid var(--linha);
  padding:22px 28px 26px;
}
.falha.ativa{display:block}
.falha h2{
  margin:0 0 8px;font-size:16px;font-weight:600;color:var(--vermelho);
}
.falha p{margin:0 0 10px;font-size:14px}
.falha ul{margin:8px 0 0;padding-left:20px;font-size:13.5px;color:var(--tinta-fraca)}
.falha code{
  font-family:"IBM Plex Mono",monospace;font-size:12.5px;
  background:var(--papel);padding:1px 5px;border-radius:2px;
}

/* -------- upload -------- */
.solta{
  border:2px dashed var(--linha);border-radius:3px;
  padding:26px 18px;text-align:center;cursor:pointer;
  transition:border-color .15s,background .15s;background:#FAFBFC;
}
.solta:hover,.solta.sobre{border-color:var(--azul);background:#F2F6FB}
.solta .titulo{
  font-family:"IBM Plex Sans",sans-serif;font-weight:600;font-size:15px;
  letter-spacing:0;text-transform:none;color:var(--tinta);margin-bottom:5px;
}
.solta .ajuda{
  font-family:"IBM Plex Sans",sans-serif;font-size:13px;font-weight:400;
  letter-spacing:0;text-transform:none;color:var(--tinta-fraca);
}
.solta input{display:none}
.arquivo{
  margin-top:12px;font-family:"IBM Plex Mono",monospace;font-size:12.5px;
  letter-spacing:0;text-transform:none;color:var(--verde);
}

/* -------- numeros de destaque -------- */
.placas{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:22px 0 6px}
.placa{border:1px solid var(--linha);border-radius:3px;padding:13px 14px;background:#FAFBFC}
.placa .rot{
  font-family:"IBM Plex Mono",monospace;font-size:10.5px;letter-spacing:.1em;
  text-transform:uppercase;color:var(--tinta-fraca);margin-bottom:6px;
}
.placa .num{font-size:19px;font-weight:600;letter-spacing:-.02em;line-height:1.2}
.placa.destaque{background:#F0F7F4;border-color:#BFE0D2}
.placa.destaque .num{color:var(--verde-dado)}
.placa .pe{font-size:11.5px;color:var(--tinta-fraca);margin-top:3px}

/* -------- tabelas -------- */
.bloco{margin-top:26px}
.bloco h3{
  font-family:"IBM Plex Mono",monospace;font-size:11px;letter-spacing:.12em;
  text-transform:uppercase;color:var(--tinta-fraca);font-weight:500;
  margin:0 0 10px;
}
.rolagem{overflow-x:auto;-webkit-overflow-scrolling:touch}
table.dados{border-collapse:collapse;width:100%;font-size:13px}
table.dados th{
  text-align:left;font-weight:600;font-size:11px;letter-spacing:.06em;
  text-transform:uppercase;color:var(--tinta-fraca);
  border-bottom:1px solid var(--linha);padding:7px 9px;white-space:nowrap;
}
table.dados td{border-bottom:1px solid #EEF1F5;padding:7px 9px;vertical-align:middle}
table.dados td.n{
  text-align:right;font-family:"IBM Plex Mono",monospace;
  font-size:12.5px;white-space:nowrap;
}
table.dados tr:hover td{background:#F7F9FB}
table.dados .fraco{color:var(--tinta-fraca)}
table.dados td.corta{
  max-width:200px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
}
table.dados td.corta.estreito{max-width:210px}
table.dados.densa th,table.dados.densa td{padding:7px 7px}

/* barra de proporcao: liquidado x a liquidar */
.prop{display:flex;gap:2px;height:8px;min-width:90px;border-radius:2px;overflow:hidden}
.prop i{display:block;height:100%}
.prop .liq{background:var(--verde-dado)}
.prop .alq{background:var(--azul)}
.legenda{display:flex;gap:16px;font-size:12px;color:var(--tinta-fraca);margin-bottom:10px}
.legenda span{display:flex;align-items:center;gap:6px}
.legenda i{width:10px;height:10px;border-radius:2px;display:block}

/* -------- barra de progresso -------- */
.progresso{display:none;padding:0 28px 26px}
.progresso.ativa{display:block}
.etapa{
  display:flex;align-items:center;gap:10px;
  font-family:"IBM Plex Mono",monospace;font-size:12.5px;
  color:var(--tinta-fraca);padding:5px 0;
}
.etapa .bolinha{
  width:7px;height:7px;border-radius:50%;background:var(--linha);flex:none;
}
.etapa.rodando .bolinha{background:var(--azul);animation:pulsa 1.1s infinite}
.etapa.rodando{color:var(--tinta)}
.etapa.feita .bolinha{background:var(--verde)}
@keyframes pulsa{0%,100%{opacity:1}50%{opacity:.25}}
@media (prefers-reduced-motion:reduce){
  .etapa.rodando .bolinha{animation:none}
}
@media (max-width:600px){
  .placas{grid-template-columns:1fr}
  body{padding:16px 12px 40px}
  form,.saida,.falha,.progresso,.orc{padding-left:18px;padding-right:18px}
  .tarja{padding:18px}
}
</style>
</head>
<body>

<div class="folha">
  <div class="tarja">
    <div class="orgao">Base Aérea de Porto Velho &middot; Seção de Licitações e Contratos 2026</div>
    <h1>Gerador TOP Comprasnet</h1>
    <div class="sub">Licitações, plano de contratações e execução orçamentária</div>
  </div>

  <div class="abas">
    <button class="aba ativa" id="aba-lic" onclick="trocaAba('lic')">Licitação</button>
    <button class="aba" id="aba-pgc" onclick="trocaAba('pgc')">Plano de contratações</button>
    <button class="aba" id="aba-cre" onclick="trocaAba('cre')">Crédito</button>
    <button class="aba" id="aba-emp" onclick="trocaAba('emp')">Empenhos</button>
  </div>

  <div class="painel ativo" id="painel-lic">
  <form id="form" autocomplete="off">
    <div class="campo">
      <label for="uasg">UASG</label>
      <input type="text" id="uasg" name="uasg" value="120641" inputmode="numeric" maxlength="6">
      <div class="msg-erro" id="erro-uasg"></div>
      <div class="dica">Código da unidade. BAPV = 120641.</div>
    </div>

    <div class="campo">
      <label for="tipo">Tipo de licitação</label>
      <select id="tipo" name="tipo">
        <option value="PREGAO">Pregão</option>
        <option value="DISPENSA">Dispensa de licitação</option>
        <option value="CONCORRENCIA">Concorrência</option>
        <option value="INEXIGIBILIDADE">Inexigibilidade</option>
        <option value="CREDENCIAMENTO">Credenciamento</option>
      </select>
    </div>

    <div class="campo">
      <label for="processo">Nº do processo</label>
      <input type="text" id="processo" name="processo" placeholder="90009/2026">
      <div class="msg-erro" id="erro-processo"></div>
      <div class="dica">Informe com o ano, no formato 90009/2026.</div>
    </div>

    <button type="submit" id="btn">Gerar planilha</button>
  </form>
  </div>

  <div class="painel" id="painel-pgc">
    <form id="form-pgc" autocomplete="off">
      <div class="campo">
        <label for="uasg-pgc">UASG</label>
        <input type="text" id="uasg-pgc" name="uasg" value="120641" inputmode="numeric" maxlength="6">
        <div class="msg-erro" id="erro-uasg-pgc"></div>
        <div class="dica">Código da unidade. BAPV = 120641.</div>
      </div>

      <div class="campo">
        <label for="ano-pgc">Ano do plano</label>
        <input type="text" id="ano-pgc" name="ano" value="2026" inputmode="numeric" maxlength="4">
        <div class="msg-erro" id="erro-ano-pgc"></div>
        <div class="dica">Ano do PCA. Uma linha por DFD, com os itens somados.</div>
      </div>

      <button type="submit" id="btn-pgc">Gerar planilha</button>
    </form>
  </div>

  <div class="painel orc" id="painel-cre" style="padding:28px">
    <label class="solta" id="solta-cre">
      <div class="titulo">Enviar o extrato de Crédito Disponível</div>
      <div class="ajuda">Arquivo .xlsx do Tesouro Gerencial. Clique ou arraste aqui.</div>
      <div class="arquivo" id="nome-cre"></div>
      <input type="file" id="arq-cre" accept=".xlsx,.xls">
    </label>
    <div id="res-cre"></div>
  </div>

  <div class="painel orc" id="painel-emp" style="padding:28px">
    <label class="solta" id="solta-emp">
      <div class="titulo">Enviar o extrato de Empenhos</div>
      <div class="ajuda">Arquivo .xlsx do Tesouro Gerencial. Clique ou arraste aqui.</div>
      <div class="arquivo" id="nome-emp"></div>
      <input type="file" id="arq-emp" accept=".xlsx,.xls">
    </label>
    <div id="res-emp"></div>
  </div>

  <div class="progresso" id="progresso"></div>

  <div class="saida" id="saida">
    <div class="carimbo">Planilha gerada</div>
    <div id="dados"></div>
    <div id="aviso-container"></div>
    <a class="baixar" id="baixar" href="#">Baixar arquivo .xlsx</a>
  </div>

  <div class="falha" id="falha">
    <h2 id="falha-titulo">Não encontrei essa licitação</h2>
    <p id="falha-texto"></p>
    <div id="falha-extra"></div>
  </div>
</div>

<p style="max-width:1020px;margin:14px auto 0;text-align:center;
          font-family:'IBM Plex Mono',monospace;font-size:11px;
          letter-spacing:.08em;color:#9AA8B6">versão __VERSAO__</p>

<script>
const form = document.getElementById('form');
const btn = document.getElementById('btn');
const progresso = document.getElementById('progresso');
const form_pgc = document.getElementById('form-pgc');
const btnPgc = document.getElementById('btn-pgc');
const saida = document.getElementById('saida');
const falha = document.getElementById('falha');

function limpaErros(){
  ['uasg','processo','uasg-pgc','ano-pgc'].forEach(c=>{
    document.getElementById(c).classList.remove('erro');
    document.getElementById('erro-'+c).style.display='none';
  });
}
function marcaErro(campo, texto){
  const i = document.getElementById(campo);
  const e = document.getElementById('erro-'+campo);
  i.classList.add('erro'); e.textContent = texto; e.style.display='block';
  i.focus();
}

const ETAPAS_LIC = ['Localizando a contratação','Lendo os itens',
                    'Buscando os vencedores','Montando a planilha'];
const ETAPAS_PGC = ['Identificando o órgão','Lendo o plano de contratações',
                    'Agrupando por DFD'];

let timers = [];
function animaEtapas(rotulos){
  progresso.innerHTML = rotulos.map((t,i)=>
    `<div class="etapa" data-etapa="${i+1}"><span class="bolinha"></span>${t}</div>`
  ).join('');
  progresso.classList.add('ativa');
  const marcos = [0, 1500, 5000, 9000];
  timers = rotulos.map((_,i)=>setTimeout(()=>{
    const atual = document.querySelector(`[data-etapa="${i+1}"]`);
    const ant = document.querySelector(`[data-etapa="${i}"]`);
    if(ant) ant.className='etapa feita';
    if(atual) atual.className='etapa rodando';
  }, marcos[i] !== undefined ? marcos[i] : 9000));
}
function paraEtapas(){
  timers.forEach(clearTimeout); timers=[];
  progresso.classList.remove('ativa');
}

const ABAS = ['lic','pgc','cre','emp'];
function trocaAba(qual){
  document.querySelector('.folha').classList.toggle('larga',
      qual==='cre' || qual==='emp');
  ABAS.forEach(a=>{
    document.getElementById('aba-'+a).classList.toggle('ativa', a===qual);
    document.getElementById('painel-'+a).classList.toggle('ativo', a===qual);
  });
  saida.classList.remove('ativa');
  falha.classList.remove('ativa');
  paraEtapas();
}

// ---- utilidades de formatacao ----
const brl = n => (n||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
const esc = t => String(t==null?'':t).replace(/[&<>"]/g, c =>
  ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

function placa(rot, num, pe, destaque){
  return `<div class="placa${destaque?' destaque':''}">
    <div class="rot">${esc(rot)}</div>
    <div class="num">R$ ${brl(num)}</div>
    ${pe?`<div class="pe">${esc(pe)}</div>`:''}</div>`;
}

function tabela(titulo, colunas, linhas, legenda, classe){
  return `<div class="bloco"><h3>${esc(titulo)}</h3>
    ${legenda||''}
    <div class="rolagem"><table class="dados">
      <thead><tr>${colunas.map(c=>`<th${c.n?' style="text-align:right"':''}>${esc(c.t)}</th>`).join('')}</tr></thead>
      <tbody>${linhas.join('')}</tbody>
    </table></div></div>`;
}

const LEGENDA_LIQ = `<div class="legenda">
  <span><i style="background:var(--verde-dado)"></i>Liquidado</span>
  <span><i style="background:var(--azul)"></i>A liquidar</span></div>`;

function barraProp(liq, alq){
  const t = (liq||0)+(alq||0);
  if(t<=0) return '<span class="fraco">—</span>';
  const p = Math.round(liq/t*100);
  return `<div class="prop" title="${p}% liquidado">
    <i class="liq" style="width:${p}%"></i><i class="alq" style="width:${100-p}%"></i></div>`;
}

// ---- envio de planilha ----
function ligarUpload(tipo, rota, render){
  const campo = document.getElementById('arq-'+tipo);
  const area  = document.getElementById('solta-'+tipo);
  const nome  = document.getElementById('nome-'+tipo);
  const res   = document.getElementById('res-'+tipo);

  ['dragenter','dragover'].forEach(ev=>area.addEventListener(ev, e=>{
    e.preventDefault(); area.classList.add('sobre');}));
  ['dragleave','drop'].forEach(ev=>area.addEventListener(ev, e=>{
    e.preventDefault(); area.classList.remove('sobre');}));
  area.addEventListener('drop', e=>{
    if(e.dataTransfer.files.length){ campo.files = e.dataTransfer.files; enviar(); }});
  campo.addEventListener('change', enviar);

  async function enviar(){
    const f = campo.files[0];
    if(!f) return;
    nome.textContent = 'Lendo ' + f.name + '...';
    res.innerHTML = '';
    const fd = new FormData(); fd.append('arquivo', f);
    try{
      const r = await fetch(rota, {method:'POST', body:fd});
      if(!r.ok){
        nome.textContent = '';
        res.innerHTML = `<div class="aviso">O servidor recusou o arquivo (erro ${r.status}). Se ele for muito grande, exporte um período menor.</div>`;
        return;
      }
      const d = await r.json();
      if(!d.ok){
        nome.textContent = '';
        res.innerHTML = `<div class="aviso"><strong>${esc(d.titulo||'Não consegui ler')}</strong><br>${esc(d.mensagem||'')}</div>`;
        return;
      }
      nome.textContent = f.name + ' — lido com sucesso';
      res.innerHTML = render(d);
    }catch(err){
      nome.textContent = '';
      res.innerHTML = '<div class="aviso">A conexão caiu durante o envio. Tente de novo.</div>';
    }
  }
}

function mostraFalha(titulo, texto, extra){
  document.getElementById('falha-titulo').textContent = titulo;
  document.getElementById('falha-texto').textContent = texto;
  document.getElementById('falha-extra').innerHTML = extra || '';
  falha.classList.add('ativa');
}

function trataHttp(status){
  if(status === 502 || status === 503){
    mostraFalha('O servidor está acordando',
      'Este site hiberna quando fica parado. Espere cerca de um minuto e '
      + 'clique em Gerar planilha de novo.');
  }else if(status === 504){
    mostraFalha('A consulta demorou demais',
      'O Compras.gov.br está lento agora. Tente novamente em alguns minutos.');
  }else{
    mostraFalha('O servidor não respondeu', 'Tente de novo em alguns instantes.');
  }
}

form.addEventListener('submit', async (ev)=>{
  ev.preventDefault();
  limpaErros();
  saida.classList.remove('ativa');
  falha.classList.remove('ativa');

  const uasg = document.getElementById('uasg').value.trim();
  const processo = document.getElementById('processo').value.trim();
  if(!uasg){ marcaErro('uasg','Informe a UASG.'); return; }
  if(!processo){ marcaErro('processo','Informe o número do processo.'); return; }

  btn.disabled = true;
  btn.textContent = 'Consultando o Compras.gov.br...';
  animaEtapas(ETAPAS_LIC);

  try{
    const r = await fetch('/gerar', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({uasg, processo, tipo: document.getElementById('tipo').value})
    });

    if(!r.ok){ paraEtapas(); trataHttp(r.status); return; }

    const d = await r.json();
    paraEtapas();

    if(!d.ok){
      if(d.campo){ marcaErro(d.campo, d.mensagem); }
      else{ mostraFalha(d.titulo || 'Não deu para gerar', d.mensagem, d.extra); }
      return;
    }

    const f = (n)=> n.toLocaleString('pt-BR',{minimumFractionDigits:2, maximumFractionDigits:2});
    document.getElementById('dados').innerHTML = `
      <div class="linha-dado"><span class="rot">Licitação</span><span class="val">${d.numero_exib}</span></div>
      <div class="linha-dado"><span class="rot">Processo</span><span class="val">${d.processo||'—'}</span></div>
      <div class="linha-dado"><span class="rot">Objeto</span><span class="val">${d.objeto||'—'}</span></div>
      <div class="linha-dado"><span class="rot">Itens</span><span class="val">${d.itens}</span></div>
      <div class="linha-dado"><span class="rot">Com vencedor</span><span class="val">${d.com_vencedor}</span></div>
      <div class="linha-dado"><span class="rot">Desertos / fracassados</span><span class="val">${d.sem_vencedor}</span></div>
      <div class="linha-dado"><span class="rot">Valor homologado</span><span class="val destaque">R$ ${f(d.valor)}</span></div>`;
    document.getElementById('aviso-container').innerHTML =
      d.aviso ? `<div class="aviso">${d.aviso}</div>` : '';
    document.getElementById('baixar').href = '/baixar/' + d.token;
    saida.classList.add('ativa');
    saida.scrollIntoView({behavior:'smooth', block:'nearest'});

  }catch(err){
    paraEtapas();
    mostraFalha('Não consegui completar a consulta',
      'A conexão caiu no meio do caminho. Verifique sua internet e clique '
      + 'em Gerar planilha de novo.');
  }finally{
    btn.disabled = false;
    btn.textContent = 'Gerar planilha';
  }
});
form_pgc.addEventListener('submit', async (ev)=>{
  ev.preventDefault();
  limpaErros();
  saida.classList.remove('ativa');
  falha.classList.remove('ativa');

  const uasg = document.getElementById('uasg-pgc').value.trim();
  const ano  = document.getElementById('ano-pgc').value.trim();
  if(!uasg){ marcaErro('uasg-pgc','Informe a UASG.'); return; }
  if(!/^\d{4}$/.test(ano)){ marcaErro('ano-pgc','Informe o ano com 4 dígitos.'); return; }

  btnPgc.disabled = true;
  btnPgc.textContent = 'Consultando o Compras.gov.br...';
  animaEtapas(ETAPAS_PGC);

  try{
    const r = await fetch('/gerar-pgc', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({uasg, ano})
    });
    if(!r.ok){ paraEtapas(); trataHttp(r.status); return; }

    const d = await r.json();
    paraEtapas();

    if(!d.ok){
      if(d.campo){ marcaErro(d.campo, d.mensagem); }
      else{ mostraFalha(d.titulo || 'Não deu para gerar', d.mensagem, d.extra); }
      return;
    }

    const f = (n)=> n.toLocaleString('pt-BR',{minimumFractionDigits:2, maximumFractionDigits:2});
    document.getElementById('dados').innerHTML = `
      <div class="linha-dado"><span class="rot">Unidade</span><span class="val">${d.nome_uasg||d.uasg}</span></div>
      <div class="linha-dado"><span class="rot">Plano</span><span class="val">PCA ${d.ano}</span></div>
      <div class="linha-dado"><span class="rot">DFDs</span><span class="val">${d.dfds}</span></div>
      <div class="linha-dado"><span class="rot">Itens somados</span><span class="val">${d.itens}</span></div>
      <div class="linha-dado"><span class="rot">Valor planejado</span><span class="val destaque">R$ ${f(d.valor)}</span></div>`;
    document.getElementById('aviso-container').innerHTML =
      d.aviso ? `<div class="aviso">${d.aviso}</div>` : '';
    document.getElementById('baixar').href = '/baixar/' + d.token;
    saida.classList.add('ativa');
    saida.scrollIntoView({behavior:'smooth', block:'nearest'});

  }catch(err){
    paraEtapas();
    mostraFalha('Não consegui completar a consulta',
      'A conexão caiu no meio do caminho. Verifique sua internet e clique '
      + 'em Gerar planilha de novo.');
  }finally{
    btnPgc.disabled = false;
    btnPgc.textContent = 'Gerar planilha';
  }
});

// ================= CREDITO =================
ligarUpload('cre', '/upload-credito', d => {
  const r = d.resumo;
  let h = `<div class="placas">
    ${placa('Crédito recebido', r.recebido)}
    ${placa('Empenhado', r.empenhadas)}
    ${placa('Disponível', r.disponivel, null, true)}
  </div>
  <div class="pe" style="font-size:12px;color:var(--tinta-fraca);margin-top:4px">
    ${r.linhas} registros em ${r.ugs} unidade(s) gestora(s)</div>`;

  h += tabela('Por unidade gestora',
    [{t:'UG'},{t:'Código'},{t:'Recebido',n:1},{t:'Empenhado',n:1},{t:'Disponível',n:1}],
    r.por_ug.map(u=>`<tr>
      <td>${esc(u.ug_nome)}</td><td class="fraco">${esc(u.ug_codigo)}</td>
      <td class="n">${brl(u.recebido)}</td><td class="n">${brl(u.empenhadas)}</td>
      <td class="n"><strong>${brl(u.disponivel)}</strong></td></tr>`));

  h += tabela('Por natureza de despesa',
    [{t:'ND'},{t:'Recebido',n:1},{t:'Empenhado',n:1},{t:'Disponível',n:1}],
    r.por_nd.map(n=>`<tr>
      <td>${esc(n.nd)}</td>
      <td class="n">${brl(n.recebido)}</td><td class="n">${brl(n.empenhadas)}</td>
      <td class="n"><strong>${brl(n.disponivel)}</strong></td></tr>`));

  if(r.por_acao.length) h += tabela('Por ação de governo',
    [{t:'Ação'},{t:'Recebido',n:1},{t:'Empenhado',n:1},{t:'Disponível',n:1}],
    r.por_acao.map(a=>`<tr>
      <td>${esc(a.acao)}</td>
      <td class="n">${brl(a.recebido)}</td><td class="n">${brl(a.empenhadas)}</td>
      <td class="n">${brl(a.disponivel)}</td></tr>`));
  return h;
});

// ================= EMPENHOS =================
ligarUpload('emp', '/upload-empenhos', d => {
  const r = d.resumo;
  let h = `<div class="placas">
    ${placa('Empenhado', r.empenhado, r.nes + ' notas')}
    ${placa('A liquidar', r.a_liquidar)}
    ${placa('Liquidado', r.liquidado)}
  </div>`;

  h += tabela('Vínculo com contrato',
    [{t:'Grupo'},{t:'Notas',n:1},{t:'Empenhado',n:1}],
    [`<tr><td>Fornecedor com CNPJ <span class="fraco">— pode ter contrato</span></td>
        <td class="n">${r.nes_fornecedores}</td><td class="n">${brl(r.valor_fornecedores)}</td></tr>`,
     `<tr><td>Própria UG ou CPF <span class="fraco">— folha, auxílios, diárias</span></td>
        <td class="n">${r.nes_proprios}</td><td class="n">${brl(r.valor_proprios)}</td></tr>`]);

  h += tabela('Por natureza de despesa',
    [{t:'ND'},{t:'Descrição'},{t:'Notas',n:1},{t:'Empenhado',n:1},
     {t:'A liquidar',n:1},{t:'Liquidado',n:1},{t:'Execução'}],
    r.por_nd.map(n=>`<tr>
      <td>${esc(n.nd)}</td><td class="fraco corta estreito" title="${esc(n.nome||'')}">${esc(n.nome||'')}</td>
      <td class="n">${n.nes}</td><td class="n">${brl(n.empenhado)}</td>
      <td class="n">${brl(n.a_liquidar)}</td><td class="n">${brl(n.liquidado)}</td>
      <td style="width:96px">${barraProp(n.liquidado, n.a_liquidar)}</td></tr>`),
    LEGENDA_LIQ, ' densa');

  h += tabela('Por modalidade',
    [{t:'Modalidade'},{t:'Notas',n:1},{t:'Empenhado',n:1},{t:'A liquidar',n:1}],
    r.por_modalidade.map(m=>`<tr>
      <td>${esc(m.modalidade||'—')}</td><td class="n">${m.nes}</td>
      <td class="n">${brl(m.empenhado)}</td><td class="n">${brl(m.a_liquidar)}</td></tr>`));

  h += tabela(`Maiores fornecedores (${r.fornecedores} no total)`,
    [{t:'CNPJ'},{t:'Fornecedor'},{t:'Notas',n:1},{t:'Empenhado',n:1},{t:'A liquidar',n:1}],
    r.por_fornecedor.map(f=>`<tr>
      <td class="fraco" style="font-family:'IBM Plex Mono',monospace;font-size:12px">${esc(f.favorecido_ni)}</td>
      <td class="corta" title="${esc(f.nome||'')}">${esc(f.nome||'')}</td><td class="n">${f.nes}</td>
      <td class="n">${brl(f.empenhado)}</td><td class="n">${brl(f.a_liquidar)}</td></tr>`));
  return h;
});
</script>
</body>
</html>
"""


# =============================================================================
# ROTAS
# =============================================================================

def limpar_antigos(horas=2):
    """Remove planilhas geradas ha mais de N horas (o link ja expirou)."""
    limite = time.time() - horas * 3600
    for nome in os.listdir(PASTA_SAIDA):
        caminho = os.path.join(PASTA_SAIDA, nome)
        try:
            if os.path.isfile(caminho) and os.path.getmtime(caminho) < limite:
                os.remove(caminho)
        except OSError:
            pass
    for token in [t for t, (c, _) in ARQUIVOS.items() if not os.path.exists(c)]:
        ARQUIVOS.pop(token, None)


@app.errorhandler(Exception)
def erro_inesperado(e):
    """
    Converte qualquer falha nao prevista em uma resposta legivel na tela,
    em vez do erro 500 generico, que nao diz nada ao usuario.
    """
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return e

    import traceback
    traceback.print_exc()          # continua indo para o log do servidor

    linha = ""
    tb = traceback.extract_tb(e.__traceback__)
    if tb:
        ult = tb[-1]
        linha = (f"{os.path.basename(ult.filename)}, linha {ult.lineno}, "
                 f"em {ult.name}()")

    return jsonify(
        ok=False,
        titulo="Erro inesperado ao gerar a planilha",
        mensagem=f"{type(e).__name__}: {e}",
        extra=(f"<p style='font-size:13px;margin-top:10px'>Origem: "
               f"<code>{linha}</code></p>" if linha else ""),
    ), 200


@app.route("/versao")
def versao():
    """Diz qual código está no ar. Evita diagnosticar a versão errada."""
    return jsonify(
        versao=getattr(core, "VERSAO", "?"),
        tempo_limite=getattr(core, "TEMPO_LIMITE", None),
        timeout_requisicao=getattr(core, "TIMEOUT_REQ", None),
        tamanho_pagina=getattr(core, "TAMANHO_PAGINA", None),
        modulos={"pgc": pgc_core is not None,
                 "orcamento": orcamento_core is not None},
    )


@app.route("/")
def inicio():
    return render_template_string(
        PAGINA.replace("__VERSAO__", getattr(core, "VERSAO", "?"))
    )


@app.route("/gerar", methods=["POST"])
def gerar():
    limpar_antigos()
    dados = request.get_json(silent=True) or {}

    # ---- validacao ----------------------------------------------------------
    uasg = re.sub(r"\D", "", str(dados.get("uasg", "")))
    if not (5 <= len(uasg) <= 6):
        return jsonify(ok=False, campo="uasg",
                       mensagem="A UASG tem 5 ou 6 dígitos. Exemplo: 120641.")

    tipo = str(dados.get("tipo", "")).strip().upper()
    if tipo not in core.MODALIDADES:
        return jsonify(ok=False, titulo="Tipo inválido",
                       mensagem="Escolha um tipo de licitação na lista.")

    try:
        numero, ano = core.parse_numero_compra(dados.get("processo", ""), None)
    except ValueError:
        return jsonify(ok=False, campo="processo",
                       mensagem="Informe o processo com o ano. Exemplo: 90009/2026.")

    # ---- consulta -----------------------------------------------------------
    saida_log = io.StringIO()
    original = core.log
    core.log = lambda m: saida_log.write(str(m) + "\n")
    core.iniciar_prazo()
    try:
        compra, itens, resultados = core.buscar_dados(
            uasg, numero, ano, core.MODALIDADES[tipo]
        )
    except core.TempoEsgotado:
        return jsonify(
            ok=False, titulo="A consulta demorou demais",
            mensagem="O Compras.gov.br não respondeu a tempo. Isso costuma ser "
                     "lentidão momentânea do portal.",
            extra="<ul><li>Tente de novo em alguns minutos.</li>"
                  "<li>Se insistir, confira se o tipo de licitação está "
                  "correto — buscar na modalidade errada varre muito mais "
                  "dados sem necessidade.</li></ul>",
        )
    except Exception as e:
        return jsonify(
            ok=False, titulo="Não consegui consultar o Compras.gov.br",
            mensagem=f"A consulta falhou: {type(e).__name__}: {e}",
        )
    finally:
        core.log = original
        core.limpar_prazo()

    if compra is None:
        texto = saida_log.getvalue()
        amostra = ""
        m = re.search(r"Numeros disponiveis \(amostra\): (.+)", texto)
        if m:
            nums = [n.strip() for n in m.group(1).split(",")][:12]
            itens_html = "".join(
                f"<code>{n[:5]}/{n[5:]}</code> " for n in nums if len(n) >= 9
            )
            amostra = (
                "<p style='font-size:13.5px;margin-top:12px'>"
                f"Nessa UASG e modalidade eu encontrei:</p><p>{itens_html}</p>"
            )
        return jsonify(
            ok=False,
            titulo="Não encontrei essa licitação",
            mensagem=f"A compra {numero[:5]}/{ano} não apareceu na UASG {uasg} "
                     f"como {NOMES_TIPO.get(tipo, tipo.title())}.",
            extra=amostra + (
                "<ul>"
                "<li>Confira se o tipo está certo — dispensa cadastrada como "
                "dispensa não aparece na busca por pregão.</li>"
                "<li>Licitações muito recentes levam alguns dias para sair no PNCP.</li>"
                "</ul>"
            ),
        )

    df = core.montar_planilha(itens, resultados)
    if df.empty:
        return jsonify(
            ok=False, titulo="Licitação encontrada, mas sem itens",
            mensagem="A contratação existe, porém o PNCP ainda não publicou "
                     "os itens dela. Tente novamente em alguns dias.",
        )

    # ---- grava --------------------------------------------------------------
    token = uuid.uuid4().hex
    nome = f"compra_{numero}.xlsx"
    caminho = os.path.join(PASTA_SAIDA, f"{token}_{nome}")
    core.gravar_xlsx(df, caminho)
    ARQUIVOS[token] = (caminho, nome)

    com_venc = int((df["CNPJ"] != core.TEXTO_SEM_RESULTADO).sum())
    sem_venc = int((df["CNPJ"] == core.TEXTO_SEM_RESULTADO).sum())
    total = float(df.loc[df["CNPJ"] != core.TEXTO_SEM_RESULTADO,
                         "VALOR TOTAL LICI"].sum())
    n_itens = int(df["ITEM"].nunique())

    aviso = ""
    if com_venc == 0:
        aviso = ("Nenhum item veio com vencedor. Se a licitação já foi "
                 "homologada, o resultado ainda não subiu para o PNCP.")
    elif sem_venc > com_venc:
        aviso = ("A maior parte dos itens está sem resultado. Vale conferir "
                 "no portal se a homologação já foi publicada por completo.")

    objeto = str(compra.get("objetoCompra") or "")
    if len(objeto) > 90:
        objeto = objeto[:90].rstrip() + "…"

    return jsonify(
        ok=True, token=token,
        numero_exib=f"{numero[:5]}/{ano} — {NOMES_TIPO.get(tipo, tipo.title())}",
        processo=compra.get("processo") or "",
        objeto=objeto,
        itens=n_itens, com_vencedor=com_venc, sem_vencedor=sem_venc,
        valor=total, aviso=aviso,
    )


@app.route("/gerar-pgc", methods=["POST"])
def gerar_pgc():
    limpar_antigos()
    if pgc_core is None:
        return jsonify(
            ok=False, titulo="Módulo indisponível",
            mensagem="O arquivo pgc_core.py não está na pasta do servidor.",
        )

    dados = request.get_json(silent=True) or {}

    uasg = re.sub(r"\D", "", str(dados.get("uasg", "")))
    if not (5 <= len(uasg) <= 6):
        return jsonify(ok=False, campo="uasg-pgc",
                       mensagem="A UASG tem 5 ou 6 dígitos. Exemplo: 120641.")

    ano_txt = re.sub(r"\D", "", str(dados.get("ano", "")))
    if len(ano_txt) != 4:
        return jsonify(ok=False, campo="ano-pgc",
                       mensagem="Informe o ano com 4 dígitos. Exemplo: 2026.")
    ano = int(ano_txt)

    try:
        cnpj, nome_uasg = pgc_core.obter_cnpj_orgao(uasg)
    except Exception as e:
        return jsonify(
            ok=False, titulo="Não consegui identificar o órgão",
            mensagem=f"A consulta da UASG falhou: {type(e).__name__}. "
                     "Tente novamente em alguns instantes.",
        )

    if not cnpj:
        return jsonify(
            ok=False, campo="uasg-pgc",
            mensagem=f"Não encontrei a UASG {uasg} no Compras.gov.br.",
        )

    core.iniciar_prazo()
    try:
        registros = pgc_core.buscar_pgc(uasg, ano, cnpj)
    except core.TempoEsgotado:
        return jsonify(
            ok=False, titulo="A consulta demorou demais",
            mensagem="O Compras.gov.br não respondeu a tempo. Tente de novo "
                     "em alguns minutos.",
        )
    except Exception as e:
        return jsonify(
            ok=False, titulo="Não consegui consultar o plano",
            mensagem=f"A consulta falhou: {type(e).__name__}: {e}",
        )
    finally:
        core.limpar_prazo()

    if not registros:
        return jsonify(
            ok=False, titulo="Nenhum item no plano",
            mensagem=f"A UASG {uasg} não tem itens de PCA {ano} publicados.",
            extra=(
                "<ul>"
                "<li>Confira se o ano está certo.</li>"
                "<li>O PGC depende da divulgação do PCA no PNCP e tem "
                "defasagem de alguns dias — planos recém-publicados podem "
                "ainda não aparecer.</li>"
                "</ul>"
            ),
        )

    df = pgc_core.montar_por_dfd(registros)
    if df.empty:
        return jsonify(
            ok=False, titulo="Não consegui agrupar por DFD",
            mensagem="Os itens vieram sem identificação de DFD.",
        )

    token = uuid.uuid4().hex
    nome = f"pgc_{uasg}_{ano}.xlsx"
    caminho = os.path.join(PASTA_SAIDA, f"{token}_{nome}")
    pgc_core.gravar_xlsx(df, caminho)
    ARQUIVOS[token] = (caminho, nome)

    total = float(df["VALOR TOTAL"].sum())
    n_itens = int(df["QTD ITENS"].sum())

    aviso = ""
    if total == 0:
        aviso = ("Todos os DFDs vieram com valor zerado. Confira no PGC se "
                 "os itens já têm estimativa de preço lançada.")

    return jsonify(
        ok=True, token=token, uasg=uasg, nome_uasg=nome_uasg or "",
        ano=ano, dfds=int(len(df)), itens=n_itens, valor=total, aviso=aviso,
    )


def _ler_enviado(leitor, rotulo):
    """Recebe o arquivo do formulário, passa pelo leitor e devolve o resumo."""
    if orcamento_core is None:
        return jsonify(ok=False, titulo="Módulo indisponível",
                       mensagem="O arquivo orcamento_core.py não está no servidor.")

    arq = request.files.get("arquivo")
    if arq is None or not arq.filename:
        return jsonify(ok=False, titulo="Nenhum arquivo",
                       mensagem="Escolha a planilha antes de enviar.")

    if not arq.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify(ok=False, titulo="Formato não aceito",
                       mensagem="Envie o arquivo .xlsx exportado do Tesouro Gerencial.")

    try:
        dados = io.BytesIO(arq.read())
        df = leitor(dados)
    except ValueError as e:
        return jsonify(
            ok=False, titulo=f"Não reconheci o extrato de {rotulo}",
            mensagem=f"{e}. Confira se é mesmo a planilha de {rotulo} e se ela "
                     "veio inteira do Tesouro Gerencial.",
        )
    except Exception as e:
        return jsonify(ok=False, titulo="Não consegui ler a planilha",
                       mensagem=f"{type(e).__name__}: {e}")

    if df.empty:
        return jsonify(ok=False, titulo="Planilha sem dados",
                       mensagem="O arquivo foi lido, mas não tem nenhuma linha.")
    return df


@app.route("/upload-credito", methods=["POST"])
def upload_credito():
    r = _ler_enviado(orcamento_core.ler_credito, "Crédito Disponível")
    if not isinstance(r, pd.DataFrame):
        return r
    return jsonify(ok=True, resumo=orcamento_core.resumo_credito(r))


@app.route("/upload-empenhos", methods=["POST"])
def upload_empenhos():
    r = _ler_enviado(orcamento_core.ler_empenhos, "Empenhos")
    if not isinstance(r, pd.DataFrame):
        return r
    return jsonify(ok=True, resumo=orcamento_core.resumo_empenhos(r))


@app.route("/baixar/<token>")
def baixar(token):
    reg = ARQUIVOS.get(token)
    if not reg or not os.path.exists(reg[0]):
        return ("Arquivo expirado. Gere a planilha novamente.", 404)
    caminho, nome = reg
    return send_file(caminho, as_attachment=True, download_name=nome)


# =============================================================================
# INICIALIZACAO
# =============================================================================

def main():
    porta = int(os.environ.get("PORT", 5000))
    local = os.environ.get("RENDER") is None

    if local:
        print("=" * 66)
        print(" GERADOR TOP COMPRASNET - servidor local")
        print("=" * 66)
        print(f"\n  Abra no navegador:  http://127.0.0.1:{porta}")
        print("\n  Para outros computadores da rede usarem, veja o IP com")
        print("  o comando ipconfig e passe  http://SEU_IP:%d" % porta)
        print("\n  MANTENHA ESTA JANELA ABERTA enquanto estiver usando.")
        print("  Para encerrar, feche a janela ou aperte Ctrl+C.\n")
        print("=" * 66)
        threading.Timer(
            1.2, lambda: webbrowser.open(f"http://127.0.0.1:{porta}")
        ).start()

    app.run(host="0.0.0.0", port=porta, debug=False)


if __name__ == "__main__":
    main()
