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
    """Matiza todos os processos chrome.exe e chromedriver.exe em execução."""
    try:
        import subprocess
        # Tenta matar processos chrome e chromedriver
        subprocess.run(['taskkill', '/IM', 'chrome.exe', '/F'], capture_output=True, shell=True)
        subprocess.run(['taskkill', '/IM', 'chromedriver.exe', '/F'], capture_output=True, shell=True)
        time.sleep(1)
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
                if col1.count('\n') > 0:
                    linhas_col1 = col1.split('\n')
                    if len(linhas_col1) > 1:
                        for l in linhas_col1[1:]:
                            l_strip = l.strip()
                            if 'Vara' in l_strip:
                                dados[-1]['vara'] = l_strip.replace('/', '').strip()
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
            elif t_upper.startswith('LOCAL'):
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

    def buscar_pje(self, nr_processo):
        if not self.driver:
            return 'SEM_CONEXAO'
        try:
            # Wait for login event (user clicked "Já loguei no PJe")
            self.login_event.wait()
            self._check_stop()
            self.log('Login PJe confirmado. Iniciando busca...')
            self.driver.switch_to.window(self.driver.window_handles[0])
            self.entrar_ngframe()
            self._check_stop()
            menu = WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#liConsultaProcessual a')))
            menu.click()
            time.sleep(2)
            self._check_stop()
            pyperclip.copy(nr_processo)
            ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
            time.sleep(3)
            self._check_stop()
            frame_consulta = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'frameConsultaProcessual')))
            self.driver.switch_to.frame(frame_consulta)
            time.sleep(1)
            self._check_stop()
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
            
            # NOVO: Busca direta na timeline - procura por "Juntada de Laudo Médico" nos movimentos
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
                        return 'OK'
                self.log('Laudo médico NÃO encontrado na timeline')
                return 'NÃO'
            except Exception as e:
                self.log(f'Erro ao buscar na timeline: {e}')
                return 'NÃO'
        except Exception as e:
            if 'Interrompido pelo usuario' in str(e):
                self.log('Busca interrompida pelo usuario')
                return 'PARADO'
            self.log(f'ERRO buscar_pje {nr_processo}: {e}')
            return f'ERRO: {e}'
        finally:
            if len(self.driver.window_handles) >= 2:
                self.driver.close()
                self.driver.switch_to.window(self.driver.window_handles[0])

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
        self.log('Etapa 2: Aguardando login no PJe...')
        self.login_event.wait()
        self.log('Login PJe confirmado. Conectando ao Chrome...')
        self.running = True
        self.done = False
        try:
            self.conectar_pje()
            self.log('Navegando para Painel > Painel do Usuario...')
            try:
                from selenium.webdriver.common.by import By
                from selenium.webdriver.support.ui import WebDriverWait
                from selenium.webdriver.support import expected_conditions as EC
                menu_painel = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.LINK_TEXT, 'Painel'))
                )
                menu_painel.click()
                time.sleep(2)
                submenu = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.LINK_TEXT, 'Painel do Usuario'))
                )
                submenu.click()
                time.sleep(3)
                self.log('Painel do Usuario carregado.')
            except Exception as e:
                self.log(f'Navegacao para Painel do Usuario falhou: {e}. Continuando...')
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
