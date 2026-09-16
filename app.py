import json, os, subprocess, threading, time, sys, traceback
import webbrowser
from flask import Flask, Response, render_template, request, jsonify, send_file
from paths import base_dir, pasta_logs
from scraper import Scraper, CAMINHO_DOCX, CAMINHO_SAIDA
from ajg_scraper import ScraperAJG
from paths import base_dir
from util_chrome import criar_driver

app = Flask(__name__)
__version__ = '1.1.0'

# ── Captura global de erros (exe windowed fecha em silêncio) ─────
def _pasta_logs():
    import paths
    pasta = os.path.join(paths.base_dir(), 'logs')
    try:
        os.makedirs(pasta, exist_ok=True)
    except Exception:
        pass
    return pasta


def _salvar_crash(exc, tb=None):
    try:
        pasta = _pasta_logs()
        with open(os.path.join(pasta, 'erro_crash.txt'), 'a', encoding='utf-8') as f:
            f.write(f'\n===== {time.strftime("%d/%m/%Y %H:%M:%S")} =====\n')
            f.write(f'Excecao: {exc}\n')
            if tb:
                f.write('Traceback:\n')
                f.write(''.join(tb))
            else:
                f.write(traceback.format_exc())
    except Exception:
        pass


def _hook_excecao_global(tp, val, tb):
    try:
        _salvar_crash(val, traceback.format_exception(tp, val, tb))
    finally:
        sys.__excepthook__(tp, val, tb)


sys.excepthook = _hook_excecao_global
threading.excepthook = lambda args: _salvar_crash(args.exc_value, args.exc_traceback)


def _criar_estrutura():
    """Garante a existencia das pastas de saida na inicializacao."""
    _pasta_logs()
    try:
        os.makedirs(os.path.join(paths.base_dir(), 'perfil_do_chrome'), exist_ok=True)
    except Exception:
        pass


_criar_estrutura()

scraper = None
scraper_thread = None

ajg_scraper = None
ajg_thread = None


def _caminho_chrome():
    candidatos = [
        os.path.join(os.environ.get('PROGRAMFILES', r'C:\Program Files'),
                     r'Google\Chrome\Application\chrome.exe'),
        os.path.join(os.environ.get('PROGRAMFILES(X86)', r'C:\Program Files (x86)'),
                     r'Google\Chrome\Application\chrome.exe'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''),
                     r'Google\Chrome\Application\chrome.exe'),
    ]
    for c in candidatos:
        if c and os.path.exists(c):
            return c
    return None


def abrir_interface(url):
    chrome = _caminho_chrome()
    if chrome:
        subprocess.Popen([chrome, '--new-window', url])
    else:
        webbrowser.open(url)


def inicializar_driver():
    """Inicializa o driver do Selenium com o perfil persistente do Chrome."""
    global scraper
    perfil = os.path.join(base_dir(), 'perfil_do_chrome')
    os.makedirs(perfil, exist_ok=True)
    driver = criar_driver(
        user_data_dir=perfil,
        prefs={'download.default_directory': os.getcwd(), 
               'download.prompt_for_download': False, 
               'download.directory_upgrade': True, 
               'profile.default_content_settings.popups': 0}
    )
    scraper.driver = driver
    scraper.log('Driver inicializado com perfil persistente do Chrome')


def abrir_chrome_pje():
    """Abre Chrome com perfil persistente para o PJe."""
    chrome_path = _caminho_chrome()
    perfil = os.path.join(base_dir(), 'perfil_do_chrome')
    if chrome_path:
        subprocess.Popen([chrome_path, f'--user-data-dir={perfil}', '--remote-debugging-port=9222', '--new-window', 'https://pje1g.trf1.jus.br/pje/'])
    else:
        webbrowser.open('https://pje1g.trf1.jus.br/pje/')
    return chrome_path


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload-arquivo', methods=['POST'])
def upload_arquivo():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Nome de arquivo vazio'}), 400
    if not file.filename.endswith('.docx'):
        return jsonify({'error': 'Apenas arquivos .docx'}), 400
    file.save(CAMINHO_DOCX)
    return jsonify({'status': 'ok', 'msg': 'Arquivo .docx carregado com sucesso'})


CAMINHO_AJG_XLSX = os.path.join(base_dir(), 'resultado_ajg.xlsx')

@app.route('/upload-arquivo-ajg', methods=['POST'])
def upload_arquivo_ajg():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Nome de arquivo vazio'}), 400
    if not file.filename.endswith('.xlsx'):
        return jsonify({'error': 'Apenas arquivos .xlsx'}), 400
    file.save(CAMINHO_AJG_XLSX)
    return jsonify({'status': 'ok', 'msg': 'Arquivo .xlsx carregado com sucesso para o AJG'})


@app.route('/start', methods=['POST'])
def iniciar():
    global scraper, scraper_thread
    if scraper and scraper.running:
        return jsonify({'error': 'Ja esta em execucao'}), 400
    if not scraper:
        scraper = Scraper()
    # Etapa 1: Extrair dados do .docx (sem Chrome ainda)
    try:
        scraper_thread = threading.Thread(target=scraper.extrair_dados_sem_chrome, daemon=True)
        scraper_thread.start()
        # Aguardar conclusão da extração (com timeout)
        if scraper.extracao_concluida.wait(timeout=30):
            # Etapa 2: Abrir Chrome e iniciar thread de verificacao
            try:
                abrir_chrome_pje()
            except Exception as e:
                return jsonify({'error': f'Erro ao abrir Chrome: {e}'}), 500
            # Iniciar thread de verificacao de juntadas (espera login_event)
            v_thread = threading.Thread(target=scraper.verificar_juntadas, daemon=True)
            v_thread.start()
            return jsonify({'status': 'chrome_aberto', 'msg': 'Chrome aberto para PJe. Faca login e clique em "Ja loguei no PJe".'})
        else:
            return jsonify({'error': 'Timeout na extração de dados'}), 500
    except Exception as e:
        return jsonify({'error': f'Erro na extração: {e}'}), 500


@app.route('/stop', methods=['POST'])
def parar():
    if scraper:
        scraper.parar()
        return jsonify({'status': 'parando'})
    return jsonify({'error': 'sem scraper ativo'}), 400


@app.route('/close-browser', methods=['POST'])
def fechar_navegador_pje():
    if scraper and scraper.driver:
        try:
            scraper.driver.quit()
        except Exception:
            pass
    return jsonify({'status': 'fechado'})


@app.route('/shutdown', methods=['POST'])
def shutdown():
    """Fecha tudo e encerra o exe de forma garantida."""
    def _encerrar():
        time.sleep(0.8)
        try:
            if scraper and getattr(scraper, 'driver', None):
                scraper.driver.quit()
        except Exception:
            pass
        try:
            if ajg_scraper and getattr(ajg_scraper, 'driver', None):
                ajg_scraper.driver.quit()
        except Exception:
            pass
        try:
            from util_processo import matar_processo_chrome as _m
            _m()
        except Exception:
            pass
        try:
            from scraper import matar_processo_chrome as _m1
            _m1()
        except Exception:
            pass
        try:
            from ajg_scraper import matar_processo_chrome as _m2
            _m2()
        except Exception:
            pass
        try:
            os.kill(os.getpid(), 9)
        except Exception:
            os._exit(0)

    threading.Thread(target=_encerrar, daemon=True).start()
    return jsonify({'status': 'encerrando', 'msg': 'Encerrando o aplicativo...'})


@app.route('/logs')
def stream_logs():
    def generate():
        last_idx = 0
        while True:
            if scraper:
                novos = scraper.get_new_logs(last_idx)
                for linha in novos:
                    last_idx += 1
                    yield f'data: {json.dumps({"msg": linha})}\n\n'
                if scraper.done and last_idx >= len(scraper.logs):
                    yield f'data: {json.dumps({"done": True})}\n\n'
                    break
                elif scraper.done:
                    continue
            time.sleep(0.5)
            yield f'data: {json.dumps({"ping": True})}\n\n'

    return Response(generate(), mimetype='text/event-stream', headers={
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
    })


@app.route('/status')
def status():
    if scraper:
        return jsonify({
            'running': scraper.running,
            'done': scraper.done,
            'log_count': len(scraper.logs),
        })
    return jsonify({'running': False, 'done': False, 'log_count': 0})


@app.route('/conectar-pje', methods=['GET', 'POST'])
def conectar_pje():
    """Abre Chrome para o PJe caso ainda nao esteja aberto.
    O usuario faz login e clica em "Ja loguei"."""
    perfil = os.path.join(base_dir(), "perfil_do_chrome")
    chrome_path = _caminho_chrome()
    if chrome_path:
        subprocess.Popen([chrome_path, f"--user-data-dir={perfil}", "--new-window", "https://pje1g.trf1.jus.br/pje/"])
    else:
        webbrowser.open("https://pje1g.trf1.jus.br/pje/")
    return jsonify({"status": "chrome_aberto", "msg": "Chrome aberto para PJe. Faça login e clique em \"Ja loguei\"."})


@app.route('/login', methods=['POST'])
def login_pje():
    """Usuario clicou que ja logou no PJe. O robô já pode assumir o controle."""
    if scraper:
        scraper.login_event.set()
        return jsonify({"status": "ok", "msg": "Sessao PJe herdada. Robo vai assumir."})
    return jsonify({"error": "Sem scraper ativo"}), 400


@app.route('/ja-loguei', methods=['POST'])
def ja_loguei():
    """Usuario clicou que ja logou no PJe. O robô já pode assumir o controle."""
    if scraper:
        scraper.login_event.set()
        return jsonify({"status": "ok", "msg": "Sessao PJe herdada. Robo vai assumir."})
    return jsonify({"error": "Sem scraper ativo"}), 400


@app.route('/resultado-existe')
def resultado_existe():
    return jsonify({'existe': os.path.exists(CAMINHO_SAIDA)})


@app.route('/varas')
def listar_varas():
    """Retorna as varas (unidades) encontradas no xlsx.
    ?source=ajg -> lê de resultado_ajg.xlsx (upload do usuario)
    senão -> lê de resultado.xlsx (saida do PJe)
    """
    import openpyxl
    source = request.args.get('source', '').strip()
    caminho = CAMINHO_AJG_XLSX if source == 'ajg' and os.path.exists(CAMINHO_AJG_XLSX) else CAMINHO_SAIDA
    if not os.path.exists(caminho):
        return jsonify({'varas': []})
    try:
        wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
        ws = wb.active
        cabecalhos = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        vara_idx = cabecalhos.index('vara') if 'vara' in cabecalhos else None
        if vara_idx is None:
            wb.close()
            return jsonify({'varas': []})
        varas = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            v = str(row[vara_idx]).strip() if row[vara_idx] else ''
            if v and v.lower() != 'none':
                varas.add(v)
        wb.close()
        return jsonify({'varas': sorted(varas)})
    except Exception:
        return jsonify({'varas': []})


@app.route('/download')
def download():
    if os.path.exists(CAMINHO_SAIDA):
        return send_file(CAMINHO_SAIDA, as_attachment=True)
    return jsonify({'error': 'Arquivo nao encontrado'}), 404


# ── AJG Scraper ──────────────────────────────────────────────────

@app.route('/start-ajg', methods=['POST'])
def iniciar_ajg():
    global ajg_scraper, ajg_thread
    if ajg_scraper and ajg_scraper.running:
        return jsonify({'error': 'Robo AJG ja esta em execucao'}), 400
    dados = request.get_json(silent=True) or {}
    vara = dados.get('vara', '').strip()
    if not vara:
        return jsonify({'error': 'Selecione a unidade (Vara) antes de iniciar.'}), 400
    # Usa arquivo uploadado (resultado_ajg.xlsx) ou fallback para resultado.xlsx do PJe
    caminho_ajg = CAMINHO_AJG_XLSX if os.path.exists(CAMINHO_AJG_XLSX) else None
    ajg_scraper = ScraperAJG(debugger_address=None, vara=vara, caminho_arquivo=caminho_ajg)
    ajg_thread = threading.Thread(target=ajg_scraper.run, daemon=True)
    ajg_thread.start()
    return jsonify({'status': 'iniciado', 'vara': vara})


@app.route('/login-ajg', methods=['POST'])
def confirmar_login_ajg():
    if ajg_scraper:
        ajg_scraper.confirmar_login()
        return jsonify({'status': 'ok'})
    return jsonify({'error': 'sem robo AJG ativo'}), 400


@app.route('/stop-ajg', methods=['POST'])
def parar_ajg():
    if ajg_scraper:
        ajg_scraper.parar()
        return jsonify({'status': 'parando'})
    return jsonify({'error': 'sem robo AJG ativo'}), 400


@app.route('/close-browser-ajg', methods=['POST'])
def fechar_navegador_ajg():
    if ajg_scraper:
        ajg_scraper.fechar_navegador()
        return jsonify({'status': 'fechado'})
    return jsonify({'error': 'sem robo AJG ativo'}), 400


@app.route('/logs-ajg')
def stream_logs_ajg():
    def generate():
        last_idx = 0
        while True:
            if ajg_scraper:
                novos = ajg_scraper.get_new_logs(last_idx)
                for linha in novos:
                    last_idx += 1
                    yield f'data: {json.dumps({"msg": linha})}\n\n'
                if ajg_scraper.done and last_idx >= len(ajg_scraper.logs):
                    yield f'data: {json.dumps({"done": True})}\n\n'
                    break
                elif ajg_scraper.done:
                    continue
            time.sleep(0.5)
            yield f'data: {json.dumps({"ping": True})}\n\n'

    return Response(generate(), mimetype='text/event-stream', headers={
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
    })


@app.route('/status-ajg')
def status_ajg():
    if ajg_scraper:
        return jsonify({
            'running': ajg_scraper.running,
            'done': ajg_scraper.done,
            'log_count': len(ajg_scraper.logs),
        })
    return jsonify({'running': False, 'done': False, 'log_count': 0})


@app.route('/processos')
def lista_processos():
    """Lista processos do resultado.xlsx, com filtro opcional por vara."""
    vara_filtro = request.args.get('vara', '').strip()
    
    if not os.path.exists(CAMINHO_SAIDA):
        return jsonify({'processos': [], 'msg': 'resultado.xlsx nao encontrado.'})
    
    import openpyxl
    wb = openpyxl.load_workbook(CAMINHO_SAIDA, read_only=True, data_only=True)
    ws = wb.active
    cabecalhos = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    processos = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        dados_row = {}
        for col_idx, coluna in enumerate(cabecalhos):
            dados_row[coluna] = str(row[col_idx]) if row[col_idx] else ''
        if dados_row.get('nr_processo'):
            # Aplicar filtro por vara se informado
            if vara_filtro and vara_filtro.upper() not in str(dados_row.get('vara', '')).upper():
                continue
            processos.append(dados_row)
    wb.close()
    
    return jsonify({'processos': processos, 'total': len(processos)})


if __name__ == '__main__':
    PORT = int(os.environ.get('PORT', 8080))
    print(f'Servidor rodando em http://localhost:{PORT}')
    print(f'Para acessar de outro computador, use http://<SEU_IP>:{PORT}')
    try:
        from util_processo import limpar_na_inicializacao
        limpar_na_inicializacao()
    except Exception:
        pass
    if not os.environ.get('WERKZEUG_RUN_MAIN') and not app.debug:
        threading.Timer(1.5, lambda: abrir_interface(f'http://localhost:{PORT}')).start()
    app.run(debug=False, host='0.0.0.0', port=PORT)