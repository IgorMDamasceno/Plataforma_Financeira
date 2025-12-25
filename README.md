# Plataforma_Financeira

Prompt para IA (gerar plataforma no Google Apps Script + Dashboard)

Crie uma plataforma de Dashboard de Controle Financeiro (Patrimônio) para rodar dentro do Google Apps Script de uma Google Planilha, usando:

1 arquivo index.html (frontend)

1 arquivo Code.gs (backend Apps Script)

A planilha como banco de dados (armazenar histórico e posições atuais)

Objetivo do produto

A plataforma NÃO é para controle de gastos. Ela é exclusivamente para:

Registrar o montante total do usuário (patrimônio) e

Mostrar a evolução do patrimônio ao longo do tempo

Mostrar o crescimento percentual mês a mês (MoM) e também crescimento no período (YTD e total)

Mostrar como o dinheiro está distribuído por:

Banco/Corretora (ex: Nubank, Inter, XP, etc.)

Tipo de investimento (Tesouro Direto, CDB, Criptomoedas, FIIs, Ações, Fundos, Reserva, etc.)

(Opcional) Ativo (ex: “Tesouro Selic 2029”, “BTC”, “HGLG11”, etc.)

Requisitos técnicos (obrigatórios)

Implementar um Web App no Apps Script (HTML Service) que abra uma dashboard.

Estrutura mínima:

Code.gs deve conter rotas/funções chamadas via google.script.run

index.html deve conter toda a UI (HTML + CSS + JS)

A UI deve ser moderna, intuitiva, responsiva (desktop e mobile).

Use Tailwind via CDN (pode usar <script src="https://cdn.tailwindcss.com"></script>) para acelerar visual.

Use Chart.js via CDN para os gráficos.

A planilha deve ser usada como armazenamento, com abas sugeridas e criadas automaticamente se não existirem.

Modelo de dados (planilha)

Crie e utilize estas abas (com cabeçalhos):

PATRIMONIO_HIST (histórico do patrimônio total)

date (YYYY-MM-DD)

total_value (número)

notes (texto opcional)

POSICOES (posição detalhada por banco e investimento; pode ser a “foto” atual ou histórico por data)

date (YYYY-MM-DD)

bank (texto)

investment_type (texto)

asset (texto opcional)

value (número)

BANCOS (cadastro simples)

bank (texto)

TIPOS (cadastro simples)

investment_type (texto)

Fluxo da plataforma

A dashboard precisa ter:

A) Tela principal (Dashboard)

Cards no topo:

Patrimônio atual

Crescimento no mês (%) (comparando mês atual vs mês anterior)

Crescimento no mês (R$)

Crescimento 12 meses (%) (se houver dados)

Gráfico 1: Linha do patrimônio ao longo do tempo (últimos 12-24 meses se existir)

Gráfico 2: Barras do crescimento mensal (%) (mês a mês)

Gráfico 3: Pizza/Donut de alocação por tipo de investimento

Gráfico 4: Pizza/Donut de alocação por banco

Tabela “Resumo” com:

banco | valor | % do total

tipo | valor | % do total

B) Tela/Modal de “Atualizar Patrimônio”

Form para inserir:

Data (default: hoje)

Valor total do patrimônio

Observação (opcional)

Ao salvar:

inserir linha em PATRIMONIO_HIST

recalcular e refletir na dashboard

C) Tela/Modal de “Atualizar Posições”

Form para adicionar linhas com:

Data (default: hoje)

Banco (dropdown + opção “novo banco”)

Tipo (dropdown + opção “novo tipo”)

Ativo (opcional)

Valor

Permitir adicionar múltiplas linhas antes de salvar (estilo “+ Adicionar linha”)

Ao salvar:

inserir linhas em POSICOES

dashboard atualiza alocações usando a data mais recente disponível (ou filtro selecionado)

D) Filtros

Filtro por período:

“Últimos 3 meses”, “6 meses”, “12 meses”, “Tudo”

Filtro por data de posição (para gráficos de alocação):

“Mais recente” (padrão)

ou selecionar uma data específica existente em POSICOES

Regras de cálculo (importante)

Patrimônio “atual” = último registro em PATRIMONIO_HIST por data.

Crescimento MoM (%) = (total_mes_atual - total_mes_anterior) / total_mes_anterior * 100

Se não houver mês anterior, mostrar “—”.

Crescimento MoM (R$) = total_mes_atual - total_mes_anterior

Para crescimento mensal, usar fechamento do mês (último registro de cada mês).

Alocação:

pegar POSICOES do dia mais recente (ou do dia selecionado)

somar por bank e por investment_type

calcular percentuais sobre o total do dia

Qualidade e UX

Visual clean e moderno:

cards com sombras suaves

tipografia forte

espaçamentos bons

skeleton/loading enquanto busca dados

Botões claros:

“Adicionar patrimônio”

“Adicionar posições”

Mensagens de erro amigáveis:

valores inválidos

campos obrigatórios

Performance:

buscar dados em lote no backend

evitar múltiplas leituras na planilha

Segurança:

validar inputs no backend

tratar NaN / strings

Entregáveis obrigatórios

Código completo do Code.gs

Código completo do index.html

Função onOpen() criando menu “Dashboard Financeiro”

Função para abrir sidebar ou dialog (preferência: sidebar)

Função para doGet() (se usar Web App) opcional, mas a solução deve funcionar pelo menu da planilha

Função para criar abas e cabeçalhos automaticamente (setup inicial)

Observação

Não use serviços externos pagos. Pode usar CDN de Tailwind e Chart.js. Não precisa integrar APIs bancárias; a entrada é manual.

Gere o código pronto para eu colar no Apps Script e rodar.
