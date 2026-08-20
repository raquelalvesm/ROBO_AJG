import json, os, subprocess, threading, time
import webbrowser
from flask import Flask, Response, render_template, request, jsonify, send_file
from scraper import Scraper, CAMINHO_DOCX, CAMINHO_SAIDA
from ajg_scraper import ScraperAJG
from paths import base_dir
from util_chrome import criar_driver

app = Flask(__name__)

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
    # Conecta no Chrome já aberto pelo PJe (porta 9222)
    # Se ja tiver Chrome com debug rodando, reusa; se nao, abre novo Chrome
    ajg_scraper = ScraperAJG(debugger_address=None)
    ajg_thread = threading.Thread(target=ajg_scraper.run, daemon=True)
    ajg_thread.start()
    return jsonify({'status': 'iniciado'})


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


if __name__ == '__main__':
    PORT = int(os.environ.get('PORT', 8080))
    print(f'Servidor rodando em http://localhost:{PORT}')
    print(f'Para acessar de outro computador, use http://<SEU_IP>:{PORT}')
    if not os.environ.get('WERKZEUG_RUN_MAIN') and not app.debug:
        threading.Timer(1.5, lambda: abrir_interface(f'http://localhost:{PORT}')).start()
    app.run(debug=False, host='0.0.0.0', port=PORT)