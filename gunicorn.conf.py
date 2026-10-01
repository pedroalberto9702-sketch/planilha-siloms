# -*- coding: utf-8 -*-
"""
=============================================================================
 Configuracao do gunicorn
=============================================================================
 O gunicorn le este arquivo sozinho quando encontra "gunicorn.conf.py" na
 pasta de trabalho -- mesmo que o comando de inicializacao seja apenas
 "gunicorn app:app", sem parametro nenhum.

 Isso importa aqui: o render.yaml so vale para servicos criados por
 Blueprint. Um servico criado pelo painel usa o comando guardado la, e o
 nosso rodava com os padroes do gunicorn -- entre eles um limite de 30
 segundos por requisicao, curto demais para consultar a API do
 Compras.gov.br. Com este arquivo a configuracao passa a viver junto do
 codigo, versionada, sem depender de ninguem lembrar de ajustar o painel.
=============================================================================
"""

import os

# ---- rede -------------------------------------------------------------------
bind = f"0.0.0.0:{os.environ.get('PORT', '10000')}"

# ---- processos --------------------------------------------------------------
# A instancia gratuita tem pouca memoria, e pandas sozinho ja ocupa boa parte
# dela. Um processo com algumas threads atende melhor do que varios processos.
workers = 1
threads = 4
worker_class = "gthread"

# ---- tempo ------------------------------------------------------------------
# 300s e o teto do servidor. O site tem trava propria bem menor (TEMPO_LIMITE,
# em gerar_planilha_siloms.py), para responder com uma mensagem clara antes de
# chegar perto daqui. Este numero existe so como rede de seguranca.
timeout = 300
graceful_timeout = 30
keepalive = 5

# ---- registro ---------------------------------------------------------------
accesslog = "-"
errorlog = "-"
loglevel = "info"

# a instancia gratuita hiberna; reciclar o processo de vez em quando evita
# que vazamentos lentos se acumulem entre um despertar e outro
max_requests = 200
max_requests_jitter = 40
