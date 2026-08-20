# RPA Pagamento Perícia - Robo_ajg

## Objetivo
Automação Flask-based (Robo_ajg) que:
1. Extrai processos de arquivo `.docx`
2. Salva em `resultado.xlsx` (12 colunas)
3. Abre Chrome para PJe (login manual)
4. Verifica "Laudo Médico" juntada via timeline text search
5. Marca `juntada` OK/NÃO
6. Aciona robô AJG que processa apenas `juntada=OK` para nomeação/honorários

## Estrutura do Projeto

### Arquivos Principais
- `app.py` — Servidor Flask com todas as rotas (upload, start, stop, logs, status, AJG)
- `scraper.py` — Etapa 1 (extracao docx) + Etapa 2 (verificar_juntadas + buscar_pje)
- `ajg_scraper.py` — Robô AJG com classe `ScraperAJG`, `ler_resultado()` (filtra juntada=OK)
- `templates/index.html` — Interface completa com seções PJe + AJG, SSE streams
- `util_chrome.py` — Função `criar_driver(user_data_dir, prefs, debugger_address)`
- `util_log.py` — Utilitário `escrever_log()`
- `paths.py` — Helper `base_dir()`

### Dados
- `resultado.xlsx` — Saída com processos, 12 colunas (inclui `juntada`)
- `entrada.docx` — Documento de entrada (gerado automaticamente no upload)

## Fluxo Completo

### Etapa 1: Extração (.docx → resultado.xlsx)
1. Upload `.docx` na interface `/upload-arquivo`
2. Clicar "Iniciar extração"
3. Dados extraídos e salvos em `resultado.xlsx`
4. Chrome abre com porta 9222 para PJe
5. Retorno: `chrome_aberto`

### Etapa 2: Verificação PJe (juntada Laudo Médico)
1. Aguardar login manual — clicar "Já loguei no PJe"
2. Sistema verifica cada processo buscando "JUNTADA DE LAUDO MÉDICO" na timeline
3. Cada processo recebe `juntada: OK` ou `juntada: NÃO`
4. Pode clicar "Parar" a qualquer momento (now com `stop_event` confiável)
5. Ao terminar, seção AJG aparece automaticamente

### Etapa 3: Robô AJG (nomeação/honorários)
1. Clicar "Iniciar Robô AJG"
2. Se Chrome PJe fechado, abre novo Chrome para AJG
3. Clicar "Já loguei no AJG"
4. Robot processa somente processos com `juntada=OK`
5. Flow por processo: Novo → PERITO → profissão → data → polo passivo → avancar
   - Tratamento de alertas (nomeacao paga, numero invalido)
   - Colar numero do processo + TAB
   - Selecionar parte ativa
   - Preencher honorarios + perito + solicitacao
   - Voltar ao inicio para próximo processo

## Endpoints API

### PJe
- `GET /` — Interface principal
- `POST /upload-arquivo` — Upload `.docx`
- `POST /start` — Inicia extração + abre Chrome PJe
- `POST /stop` — Para scraper
- `GET /logs` — SSE stream de logs
- `GET /status` — Status (running/done)
- `GET /resultado-existe` — Verifica se resultado.xlsx existe
- `GET /download` — Download do resultado.xlsx

### AJG
- `POST /start-ajg` — Inicia robô AJG
- `POST /login-ajg` — Confirma login no AJG (define login_event)
- `POST /stop-ajg` — Para robô AJG + fecha navegador
- `GET /logs-ajg` — SSE stream logs AJG
- `GET /status-ajg` — Status AJG

## Funcionalidades Implementadas

✅ Extracao de 25 processos do .docx com todos os 12 campos  
✅ Valor automatico: R$350 se "Piripiri" em `local`, R$300 senão  
✅ Verificacao PJe timeline text search (`.text-upper.texto-movimento`)  
✅ Botao "Parar" confiavel (check `stop_event` em multiplos pontos + driver.quit)  
✅ Fallback Chrome: se porta 9222 indisponivel, abre novo Chrome  
✅ Filtro AJG: processa somente `juntada=OK`  
✅ Logs detalhados por ETAPA (antes travava silenciosamente — agora mostra ETAPA 1/2/3)  
✅ Traceback em erros gerais (com `traceback.format_exc()`)  
✅ Interface com seções PJe + AJG (AJG hidden ate conclusao da Etapa 2)  
✅ SSE streams `/logs` e `/logs-ajg` para atualizacao em tempo real  
✅ `docx_fix.py` removido (template ja tem `accept=".docx"`)  
✅ 11 scripts de debug removidos (add_login_btn.py, check_accept.py, etc.)  

## Observacoes

- Chrome path verificado via `_caminho_chrome()` no Windows
- Login manual necessario tanto no PJe quanto no AJG
- O robô AJG so processa processos com `juntada=OK` (filtrado em `ler_resultado()`)
- O botao "Parar" fecha o driver e seta `stop_event`
- Arquivo `entrada.docx` e criado automaticamente no upload (sobranome do usuario original)
- O log agora mostra exatamente onde o robô esta (ETAPA 1, 2, 3 + passos 1/2/3/4/5/6 dentro da Etapa 5)

## Resumo das Correções Recentes

| Problema | Correção |
|---|---|
| "Parar" nao funcionava | `stop_event` checado em `buscar_pje()` + `driver.quit()` em `parar()` |
| AJG parava apos login | `navegar_para_nomeacao()` envolvido em try/except + logs passo a passo |
| Erro Chrome 9222 silencioso | Fallback para novo Chrome + log limpo "Chrome PJe indisponivel..." |
| Leitura resultado.xlsx sem filtro | `ler_resultado()` ja filtra `juntada == 'OK'` corretamente |
| Sem erro log ao navegar | Agora mostra `ERRO ao navegar para nomeacao: ...` ou sucesso "Tela carregada" |

## Próximos Passos

- Testar fluxo completo: upload .docx → extração → login PJe → verificar juntada → "Parar" → seção AJG aparece → iniciar AJG → login AJG → processar somente juntada=OK
- Adicionar modo headless Chrome (opcional)
- Persistencia de configuracoes (URLs, portas)