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

app = Flask(__name__)

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
  max-width:720px;margin:0 auto;
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

/* -------- abas -------- */
.abas{
  display:flex;border-bottom:1px solid var(--linha);background:#F7F9FB;
}
.aba{
  flex:1;padding:14px 10px;border:none;background:none;cursor:pointer;
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
  body{padding:16px 12px 40px}
  form,.saida,.falha,.progresso{padding-left:18px;padding-right:18px}
  .tarja{padding:18px}
}
</style>
</head>
<body>

<div class="folha">
  <div class="tarja">
    <div class="orgao">Base Aérea de Porto Velho &middot; Seção de Licitações e Contratos 2026</div>
    <h1>Gerador TOP Comprasnet</h1>
    <div class="sub">Resultado de licitação extraído do Compras.gov.br</div>
  </div>

  <div class="abas">
    <button class="aba ativa" id="aba-lic" onclick="trocaAba('lic')">Resultado de licitação</button>
    <button class="aba" id="aba-pgc" onclick="trocaAba('pgc')">Plano de contratações</button>
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

function trocaAba(qual){
  const outro = qual === 'lic' ? 'pgc' : 'lic';
  document.getElementById('aba-'+qual).classList.add('ativa');
  document.getElementById('aba-'+outro).classList.remove('ativa');
  document.getElementById('painel-'+qual).classList.add('ativo');
  document.getElementById('painel-'+outro).classList.remove('ativo');
  saida.classList.remove('ativa');
  falha.classList.remove('ativa');
  paraEtapas();
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


@app.route("/")
def inicio():
    return render_template_string(PAGINA)


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
    try:
        compra, itens, resultados = core.buscar_dados(
            uasg, numero, ano, core.MODALIDADES[tipo]
        )
    except Exception as e:
        core.log = original
        return jsonify(
            ok=False, titulo="Não consegui consultar o Compras.gov.br",
            mensagem=f"A consulta falhou: {type(e).__name__}. "
                     "Verifique sua conexão e tente de novo em alguns instantes.",
        )
    finally:
        core.log = original

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

    try:
        registros = pgc_core.buscar_pgc(uasg, ano, cnpj)
    except Exception as e:
        return jsonify(
            ok=False, titulo="Não consegui consultar o plano",
            mensagem=f"A consulta falhou: {type(e).__name__}. "
                     "Verifique sua conexão e tente de novo.",
        )

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
