import os, re, threading, time
from datetime import datetime
import docx
import openpyxl
import pyperclip
from openpyxl.styles import Font, Alignment
from paths import base_dir

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException
    from webdriver_manager.chrome import ChromeDriverManager
    selenium_importado = True
except Exception:
    selenium_importado = False
    webdriver = None
    Options = None
    Service = None
    By = None
    Keys = None
    ActionChains = None
    WebDriverWait = None
    EC = None
    TimeoutException = None
    ChromeDriverManager = None

CAMINHO_DOCX = os.path.join(base_dir(), 'entrada.docx')
CAMINHO_SAIDA = os.path.join(base_dir(), 'resultado.xlsx')
MESES = {'janeiro': '01', 'fevereiro': '02', 'marco': '03', 'abril': '04', 'maio': '05', 'junho': '06',
         'julho': '07', 'agosto': '08', 'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12'}
COLUNAS_SAIDA = ['nr_processo', 'perito', 'profissao', 'data_nomeacao', 'local', 'objeto_acao', 'vara', 'autor', 'polo_passivo', 'assinatura_autor', 'valor', 'juntada']
REGEX_NR_PROCESSO = re.compile(r'(\d{7}-\d{2}\.\d{4}\.\d{1,2}\.\d{1,2}\.\d{4})')
VALOR_PIRIPIRI = 'R$ 350,00'
VALOR_OUTROS = 'R$ 300,00'


def matar_processo_chrome():
    """Mata todos os processos chrome.exe e chromedriver.exe em execução."""
    try:
        from util_processo import matar_processo_chrome as _m
        _m()
    except Exception:
        pass


class Scraper:
    def __init__(self):
        self.logs = []
        self.log_event = threading.Event()
        self.login_event = threading.Event()
        self.stop_event = threading.Event()
        self.running = False
        self.done = False
        self.resultados = []
        self.driver = None
        self.extracao_concluida = threading.Event()  # Novo: sinaliza quando extração termina

    def log(self, msg):
        ts = datetime.now().strftime('%H:%M:%S')
        entry = f'[{ts}] {msg}'
        self.logs.append(entry)
        self.log_event.set()
        print(entry)

    def get_new_logs(self, last_idx):
        if last_idx < len(self.logs):
            return self.logs[last_idx:]
        return []

    def parar(self):
        self.stop_event.set()
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
        # Mata processos Chrome/ChromeDriver em segundo plano
        try:
            matar_processo_chrome()
        except Exception:
            pass
        self.driver = None

    def extrair_numero_processo(self, texto):
        m = REGEX_NR_PROCESSO.search(texto)
        return m.group(1) if m else None

    def extrair_dados_docx(self, doc):
        dados = []
        for tabela in doc.tables:
            for linha in tabela.rows:
                celulas = [c.text.strip() for c in linha.cells]
                if len(celulas) < 2:
                    continue
                col1 = celulas[1]
                nr = self.extrair_numero_processo(col1)
                if not nr:
                    continue
                if col1.startswith('N') and 'DO PROCESSO' in col1.upper():
                    continue
                assinatura_autor = ''
                for c in celulas:
                    if c.strip().upper() in ('OK', 'OK.'):
                        assinatura_autor = c.strip().upper()
                        break
                dados.append({
                    'nr_processo': nr,
                    'objeto_acao': col1,
                    'vara': '',
                    'autor': '',
                    'polo_passivo': '',
                    'assinatura_autor': assinatura_autor,
                })
                # Log do conteúdo da célula para debug
                self.log(f'  [DEBUG] Celula col1: {repr(col1[:200])}')
                # Tenta extrair vara e partes de várias formas
                # Formato novo: "Nº – Objeto/ VARA / AUTOR X POLO" (separador " / ")
                partes_slash = [p.strip() for p in col1.split('/')]
                if len(partes_slash) >= 3:
                    # Formato com " / " como separador
                    dados[-1]['vara'] = partes_slash[1].strip()
                    partes_texto = partes_slash[2].strip()
                    if ' X ' in partes_texto:
                        px = partes_texto.split(' X ', 1)
                        dados[-1]['autor'] = px[0].strip()
                        dados[-1]['polo_passivo'] = px[1].strip()
                        dados[-1]['partes'] = partes_texto
                elif len(partes_slash) == 2:
                    # Pode ter vara + partes juntas
                    dados[-1]['vara'] = partes_slash[1].strip()
                else:
                    # Fallback: procura por " X " em qualquer linha
                    linhas_col1 = col1.replace('\r\n', '\n').replace('\r', '\n').split('\n')
                    for l in linhas_col1:
                        l_strip = l.strip()
                        if not l_strip:
                            continue
                        if ' X ' in l_strip and len(l_strip) > 10:
                            partes = l_strip.split(' X ', 1)
                            dados[-1]['autor'] = partes[0].strip()
                            dados[-1]['polo_passivo'] = partes[1].strip()
                            dados[-1]['partes'] = l_strip
        if not dados:
            self.log('Nenhuma tabela valida encontrada no documento.')
        return dados

    def extrair_metadados(self, doc):
        metadados = {}
        for p in doc.paragraphs:
            t = p.text.strip()
            if not t:
                continue
            t_upper = t.upper()
            if t_upper.startswith('PERITO'):
                metadados['perito'] = t.split(':', 1)[-1].strip()
            elif t_upper.startswith('DATA DA PERÍCIA') or 'DATA DA PERICIA' in t_upper:
                metadados['data_nomeacao'] = t.split(':', 1)[-1].strip()
            elif t_upper.startswith('LOCAL') or 'CIDADE DE REALIZAÇÃO' in t_upper or 'CIDADE DE REALIZACAO' in t_upper:
                metadados['local'] = t.split(':', 1)[-1].strip()
            elif t_upper.startswith('PROFISS'):
                metadados['profissao'] = t.split(':', 1)[-1].strip()
        return metadados

    def conectar_pje(self):
        if self.driver:
            return
        if not selenium_importado:
            raise Exception('Selenium nao instalado. Instale: pip install selenium webdriver-manager')
        options = Options()
        options.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.implicitly_wait(10)
        self.log('Conectado ao Chrome via debugger 127.0.0.1:9222')
    
    def _check_stop(self):
        """Verifica se deve parar e lança exceção se sim."""
        if self.stop_event.is_set():
            raise Exception('Interrompido pelo usuario')

    def entrar_ngframe(self):
        self.driver.switch_to.default_content()
        iframe = WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#ngFrame')))
        self.driver.switch_to.frame(iframe)
        time.sleep(2)

    def sessao_valida(self):
        if not self.driver:
            return False
        try:
            self.driver.current_window_handle
            return True
        except Exception:
            return False

    def _caminho_chrome(self):
        import os
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

    def recuperar_sessao(self):
        """Relança Chrome com perfil persistente (login mantido). Evita matar o Chrome que vamos subir."""
        try:
            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None
        except Exception:
            self.driver = None
        # NÃO mata todos os chrome.exe aqui - pode matar o que vamos subir
        # Apenas mata chromedriver órfãos
        try:
            from util_processo import matar_chromedriver
            matar_chromedriver()
        except Exception:
            pass
        time.sleep(3)
        try:
            import subprocess
            chrome_path = self._caminho_chrome()
            perfil = os.path.join(base_dir(), 'perfil_do_chrome')
            if chrome_path:
                subprocess.Popen([chrome_path, f'--user-data-dir={perfil}',
                                  '--remote-debugging-port=9222', '--new-window',
                                  'https://pje1g.trf1.jus.br/pje/'])
            # Aguarda a porta 9222 responder antes de reconectar
            import urllib.request
            for _ in range(60):
                try:
                    urllib.request.urlopen('http://127.0.0.1:9222/json', timeout=2)
                    break
                except Exception:
                    time.sleep(1)
            self.conectar_pje()
            time.sleep(5)  # espera a página carregar completamente
            return self.sessao_valida()
        except Exception as e:
            self.log(f'Falha ao recuperar sessao do Chrome: {e}')
            self.driver = None
            return False

    def buscar_pje(self, nr_processo, tentativas=2):
        while tentativas > 0:
            if not self.driver:
                self.log('Sessao do Chrome nao disponivel. Tentando reconectar...')
                if not self.recuperar_sessao():
                    return 'ERRO: Chrome desconectado'
            if not self.sessao_valida():
                self.log('Sessao do Chrome indisponivel. Tentando reconectar...')
                if not self.recuperar_sessao():
                    return 'ERRO: Chrome desconectado'
            try:
                self._check_stop()
                self.log('Login PJe confirmado. Iniciando busca...')
                self.driver.switch_to.window(self.driver.window_handles[0])
                self.driver.switch_to.default_content()
                self._check_stop()
                self.entrar_ngframe()
                self._check_stop()
                # Clica no ícone Home (fa fa-home) para resetar o estado do ngFrame
                try:
                    home = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.fa.fa-home')))
                    home.click()
                    time.sleep(2)
                except Exception:
                    pass
                menu = WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#liConsultaProcessual a')))
                menu.click()
                time.sleep(2)
                self._check_stop()
                # Cola o NÚMERO COMPLETO do processo ANTES de entrar no frameConsultaProcessual
                # (campo de busca fica no nível do ngFrame, não dentro do iframe)
                pyperclip.copy(nr_processo)
                ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
                time.sleep(0.3)
                ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
                time.sleep(3)
                self._check_stop()
                frame_consulta = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'frameConsultaProcessual')))
                self.driver.switch_to.frame(frame_consulta)
                botao = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#fPP\\:searchProcessos')))
                botao.click()
                time.sleep(3)
                self._check_stop()
                link_processo = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.btn-link.btn-condensed')))
                link_processo.click()
                time.sleep(3)
                self._check_stop()
                self.driver.switch_to.window(self.driver.window_handles[-1])
                time.sleep(3)
                self._check_stop()
                self.driver.switch_to.default_content()
                campo = self.driver.find_element(By.ID, 'divTimeLine:txtPesquisa')
                campo.clear()
                campo.send_keys('Juntada de Laudo Médico')
                time.sleep(2)
                self._check_stop()
                botoes = self.driver.find_elements(By.TAG_NAME, 'button')
                clicked = False
                for b in botoes:
                    bid = b.get_attribute('id') or ''
                    if 'btnPesquisa' in bid or 'pesquisa' in bid.lower():
                        b.click()
                        time.sleep(3)
                        clicked = True
                        break
                if not clicked:
                    campo.send_keys(Keys.RETURN)
                    time.sleep(3)
                self._check_stop()
                
                # Busca direta na timeline - procura por "Juntada de Laudo Médico" nos movimentos
                try:
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".text-upper.texto-movimento"))
                    )
                    self._check_stop()
                    movimentos = self.driver.find_elements(By.CSS_SELECTOR, ".text-upper.texto-movimento")
                    for mov in movimentos:
                        texto = mov.text.strip().upper()
                        if "JUNTADA DE LAUDO MÉDICO" in texto or "JUNTADA DE LAUDO MEDICO" in texto:
                            self.log(f'Laudo médico encontrado na timeline: {mov.text[:150]}')
                            # ── VALIDAÇÃO ANTI-FALSO-POSITIVO ─────────────
                            # Confirma que a timeline aberta é realmente do processo pesquisado.
                            # Se o PJe não recarregou a timeline (ficou cache do processo anterior),
                            # o número exibido não bate com o alvo → NÃO marcar OK.
                            try:
                                digitos_alvo = re.sub(r'\D', '', nr_processo)
                                nucleo_alvo = digitos_alvo[:12] if len(digitos_alvo) >= 12 else digitos_alvo

                                def _timeline_exibe_alvo():
                                    corpo = self.driver.find_element(By.TAG_NAME, 'body').text
                                    digitos_corpo = re.sub(r'\D', '', corpo)
                                    return bool(nucleo_alvo) and nucleo_alvo in digitos_corpo

                                if _timeline_exibe_alvo():
                                    self.log(f'Validado: timeline exibe o processo {nr_processo}. Confirmando OK.')
                                    return 'OK'

                                # Cache do processo anterior. Forca um F5/reload da timeline para
                                # limpar o cache do processo anterior e re-executar a consulta no
                                # servidor, recarregando o movimento correto do processo alvo.
                                for tent in range(1, 4):
                                    self.log(f'Cache detectado (tentativa {tent}/3). Aplicando F5 na timeline...')
                                    try:
                                        self.driver.refresh()
                                        time.sleep(5)
                                        self._check_stop()
                                    except Exception as re_:
                                        self.log(f'Nao foi possivel dar F5 na timeline ({re_}).')
                                        break
                                    try:
                                        WebDriverWait(self.driver, 15).until(
                                            EC.presence_of_element_located((By.CSS_SELECTOR, ".text-upper.texto-movimento"))
                                        )
                                    except Exception:
                                        pass
                                    time.sleep(2)
                                    self._check_stop()
                                    if _timeline_exibe_alvo():
                                        self.log(f'Validado (apos F5 na timeline): timeline exibe o '
                                                 f'processo {nr_processo}. Confirmando OK.')
                                        return 'OK'
                                self.log(f'ALERTA FALSO POSITIVO: timeline NAO exibe o processo {nr_processo} '
                                         f'(mesmo apos F5 3x - provavelmente cache permanente do processo '
                                         f'anterior). Marcando NÃO para evitar erro.')
                                return 'NÃO'
                            except Exception as ve:
                                self.log(f'Falha ao validar processo na timeline ({ve}). Tratando como NÃO.')
                                return 'NÃO'
                    self.log('Laudo médico NÃO encontrado na timeline')
                    return 'NÃO'
                except Exception as e:
                    msg_erro = str(e).split('\n')[0].strip()
                    if 'timeout' in msg_erro.lower() or 'no such element' in msg_erro.lower():
                        self.log(f'Laudo medico nao encontrado no processo {nr_processo}.')
                    else:
                        self.log(f'Erro ao buscar laudo no processo: {msg_erro}')
                    return 'NÃO'
            except Exception as e:
                if 'Interrompido pelo usuario' in str(e):
                    self.log('Busca interrompida pelo usuario')
                    return 'PARADO'
                msg = str(e)
                eh_sessao = (not msg.strip() or 'no such window' in msg.lower()
                             or 'target frame detached' in msg.lower()
                             or 'chrome not reachable' in msg.lower()
                             or 'GetHandleVerifier' in msg
                             or 'NoneType' in msg)
                tentativas -= 1
                if eh_sessao and tentativas > 0:
                    self.log(f'Erro de sessao ao verificar {nr_processo}. Relancando Chrome e repetindo...')
                    if self.recuperar_sessao():
                        time.sleep(3)
                        if self.sessao_valida():
                            continue
                    tentativas = 0
                    self.log(f'Erro ao verificar processo {nr_processo}: {str(e).split(chr(10))[0]}')
                    return f'ERRO: {str(e).split(chr(10))[0]}'
            finally:
                if self.driver:
                    try:
                        if len(self.driver.window_handles) >= 2:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        # Reseta o estado do ngFrame para a próxima iteração
                        self.driver.switch_to.default_content()
                        self.entrar_ngframe()
                        try:
                            home = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.fa.fa-home')))
                            home.click()
                            time.sleep(2)
                        except Exception:
                            pass
                    except Exception as fe:
                        self.log(f'Aviso: falha ao fechar aba extra ou resetar frame: {fe}')
        return 'ERRO: excedeu tentativas'

    def salvar_resultado(self, dados_lista, caminho_saida):
        self.log(f'Salvando {len(dados_lista)} resultados...')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Resultado'
        fonte = Font(bold=True)
        for col_idx, coluna in enumerate(COLUNAS_SAIDA, 1):
            cell = ws.cell(row=1, column=col_idx, value=coluna)
            cell.font = fonte
            cell.alignment = Alignment(horizontal='center')
        for row_idx, dados in enumerate(dados_lista, 2):
            for col_idx, coluna in enumerate(COLUNAS_SAIDA, 1):
                ws.cell(row=row_idx, column=col_idx, value=dados.get(coluna, ''))
        wb.save(caminho_saida)
        self.log(f'Resultado salvo em {caminho_saida}')

    def verificar_juntadas(self):
        """Etapa 2: Espera login do usuario e verifica juntada de Laudo Medico no PJe."""
        # IMPORTANTE: threading.Event e "colante" - se ficou setado de uma execucao
        # anterior, wait() retorna na hora e o botao de login nunca aparece.
        # Limpa no inicio de CADA execucao para sempre aguardar o login manual.
        self.login_event.clear()
        self.log('Etapa 2: Aguardando login no PJe...')
        self.log('ATENCAO: Navegue manualmente ate o menu Painel > Painel do Usuario no navegador Chrome ABERTO.')
        self.log('Somente depois de chegar na tela do Painel de Usuario, clique no botao "Ja loguei no PJe".')
        self.login_event.wait()
        self.log('Login PJe confirmado. Conectando ao Chrome...')
        self.running = True
        self.done = False
        try:
            self.conectar_pje()
            time.sleep(2)
            total = len(self.resultados)
            for i, proc in enumerate(self.resultados, 1):
                if self.stop_event.is_set():
                    self.log('Verificacao interrompida pelo usuario.')
                    break
                nr = proc.get('nr_processo', '')
                self.log(f'[{i}/{total}] Verificando processo {nr}...')
                resultado = self.buscar_pje(nr)
                proc['juntada'] = resultado
                self.log(f'[{i}/{total}] Processo {nr}: juntada = {resultado}')
                self.salvar_resultado(self.resultados, CAMINHO_SAIDA)
            self.log('Verificacao de juntadas concluida.')
        except Exception as e:
            self.log(f'ERRO NA ETAPA 2: {e}')
        finally:
            self.running = False
            self.done = True

    def extrair_dados_sem_chrome(self):
        """Etapa 1: Extrair dados do .docx e salvar em resultado.xlsx.
        NÃO abre Chrome ainda. Apenas processa o documento."""
        self.log('Etapa 1: Iniciando extração de dados do .docx...')
        try:
            caminho_docx = CAMINHO_DOCX
            if not os.path.exists(caminho_docx):
                docxs = [f for f in os.listdir(base_dir()) if f.lower().endswith('.docx') and f != os.path.basename(caminho_docx)]
                if docxs:
                    caminho_docx = os.path.join(base_dir(), docxs[0])
                    self.log(f'Usando documento alternativo: {docxs[0]}')
                else:
                    self.log('ERRO: Nenhum arquivo .docx encontrado.')
                    return
            if not os.path.exists(caminho_docx):
                self.log(f'ERRO: Arquivo {caminho_docx} nao encontrado.')
                return
            self.log(f'Abrindo documento: {caminho_docx}')
            doc = docx.Document(caminho_docx)
            self.log(f'Tabelas encontradas: {len(doc.tables)}')
            metadados = self.extrair_metadados(doc)
            self.log(f'Metadados: {metadados}')
            dados = self.extrair_dados_docx(doc)
            # Deduplicar por nr_processo (mantém a primeira ocorrência)
            vistos = set()
            dados_unicos = []
            for d in dados:
                nr = d.get('nr_processo')
                if nr and nr not in vistos:
                    vistos.add(nr)
                    dados_unicos.append(d)
            if len(dados_unicos) != len(dados):
                self.log(f'Removidos {len(dados) - len(dados_unicos)} processos duplicados.')
            dados = dados_unicos
            for d in dados:
                d.update(metadados)
                local = (d.get('local') or '').upper()
                d['valor'] = VALOR_PIRIPIRI if 'PIRIPIRI' in local else VALOR_OUTROS
            for i, d in enumerate(dados, 1):
                self.log(f'  [{i}] nr_processo: {d.get("nr_processo")}')
                self.log(f'      perito: {d.get("perito", "N/A")}')
                self.log(f'      profissao: {d.get("profissao", "N/A")}')
                self.log(f'      data_nomeacao: {d.get("data_nomeacao", "N/A")}')
                self.log(f'      local: {d.get("local", "N/A")}')
                self.log(f'      vara: {d.get("vara", "N/A")}')
                self.log(f'      autor: {d.get("autor", "N/A")}')
                self.log(f'      polo_passivo: {d.get("polo_passivo", "N/A")}')
                self.log(f'      valor: {d.get("valor", "N/A")}')
            self.resultados = dados
            if dados:
                self.salvar_resultado(dados, CAMINHO_SAIDA)
                self.log(f'Resultado salvo em {CAMINHO_SAIDA}. Total: {len(dados)} processos.')
                self.log('Etapa 1 concluida. Dados extraidos e salvos.')
                self.log('Etapa 2: Abrira Chrome automaticamente para login no PJe.')
            else:
                self.log('Nenhum processo encontrado no .docx.')
        except Exception as e:
            self.log(f'ERRO NA ETAPA 1: {e}')
        finally:
            self.running = False
            self.extracao_concluida.set()  # Sinaliza que extração terminou
