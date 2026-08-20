import os, re, time, unicodedata, threading
from datetime import datetime
import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.common.exceptions import (
    NoSuchElementException, TimeoutException, StaleElementReferenceException,
    UnexpectedAlertPresentException, NoAlertPresentException,
)
from selenium.webdriver.common.action_chains import ActionChains
import pyperclip
from util_log import escrever_log
from util_chrome import criar_driver
from paths import base_dir


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

CAMINHO_RESULTADO = os.path.join(base_dir(), 'resultado.xlsx')
URL_LOGIN_AJG = ('https://ajg1.cjf.jus.br/aj/seguranca/efetuarloginintranet/'
                 'efetuarLoginIntranet_efetuarLogin.jsf')


class ScraperAJG:
    def __init__(self, debugger_address=None):
        self.logs = []
        self.log_event = threading.Event()
        self.login_event = threading.Event()
        self.stop_event = threading.Event()
        self.running = False
        self.done = False
        self.driver = None
        self.debugger_address = debugger_address

    def log(self, msg):
        ts = datetime.now().strftime('%H:%M:%S')
        entry = f'[{ts}] {msg}'
        self.logs.append(entry)
        self.log_event.set()
        print(entry)
        escrever_log(f'log_ajg_{datetime.now().strftime("%Y%m%d")}.txt', entry)

    def get_new_logs(self, last_idx):
        if last_idx < len(self.logs):
            return self.logs[last_idx:]
        return []

    @staticmethod
    def remover_acentos(texto):
        return ''.join(c for c in unicodedata.normalize('NFD', texto)
                       if unicodedata.category(c) != 'Mn')

    @staticmethod
    def digito(s):
        return re.sub(r'\D', '', s or '')

    def ler_resultado(self, caminho):
        wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
        ws = wb.active
        cabecalhos = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        resultados = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            dados_row = {}
            for col_idx, coluna in enumerate(cabecalhos):
                dados_row[coluna] = str(row[col_idx]) if row[col_idx] else ''
            if dados_row.get('nr_processo'):
                # Filtrar apenas processos com juntada = OK
                if dados_row.get('juntada', '').strip().upper() == 'OK':
                    # Mapear campos do PJe para nomes esperados pelo AJG
                    dados_row['nome'] = dados_row.get('perito', '')  # perito -> nome
                    dados_row['data_servico'] = dados_row.get('data_nomeacao', '')  # data_nomeacao -> data_servico
                    # valor já existe
                    resultados.append(dados_row)
        wb.close()
        return resultados

    def esperar_ajg(self, driver):
        try:
            el = driver.find_element(By.ID, 'loadingDiv')
            return el.value_of_css_property('display') == 'none'
        except (NoSuchElementException, StaleElementReferenceException):
            return True

    def esperar_ajg_com_timeout(self, driver, timeout=120):
        WebDriverWait(driver, timeout).until(self.esperar_ajg)

    def handle_alert(self):
        """Aceita o alerta atual (se houver) e retorna o texto; senao None."""
        try:
            alerta = WebDriverWait(self.driver, 2).until(EC.alert_is_present())
            txt = alerta.text
            self.log(f'  ALERTA: {txt}')
            alerta.accept()
            return txt
        except TimeoutException:
            return None
        except NoAlertPresentException:
            return None

    def cancelar_e_voltar_inicio(self):
        """Clica cancelar (com confirmacao) e bt_inicio_sistema para ir ao proximo processo."""
        try:
            btn_cancelar = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'formAJIntranet:cancelar'))
            )
            btn_cancelar.click()
            time.sleep(1)
            self.handle_alert()  # confirmacao "Deseja cancelar a operacao..."
        except Exception as e:
            self.log(f'  Erro ao clicar cancelar: {e}')

        try:
            btn_inicio = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'formAJIntranet:bt_inicio_sistema'))
            )
            btn_inicio.click()
            self.esperar_ajg_com_timeout(self.driver)
            time.sleep(1)
        except Exception as e:
            self.log(f'  Erro ao voltar ao inicio: {e}')

    def navegar_para_nomeacao(self):
        """Abre a tela de nomeacao de profissionais (tambem ao voltar ao inicio
        entre processos)."""
        self.log('  PASSO: Procurando link "Nomeacao de Profissionais"...')
        try:
            el = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, 'link_tela_com_tooltip_1nomeacaodeprofissionais'))
            )
            el.click()
            self.log('  PASSO: Link "Nomeacao de Profissionais" clicado (por ID).')
        except TimeoutException:
            self.log('  PASSO: Link por ID nao encontrado, buscando por texto...')
            links = self.driver.find_elements(By.TAG_NAME, 'a')
            for l in links:
                lid = l.get_attribute('id') or ''
                ltext = l.text.strip()[:50]
                if 'nomeacao' in lid.lower() or 'nomeacao' in ltext.lower():
                    l.click()
                    self.log(f'  PASSO: Link clicado: id={lid} texto={ltext}')
                    break
            else:
                self.log('  ERRO: Nenhum link de nomeacao encontrado na pagina.')
                raise Exception('Link "Nomeacao de Profissionais" nao encontrado.')
        self.log('  PASSO: Aguardando carregamento da tela de nomeacao...')
        self.esperar_ajg_com_timeout(self.driver)
        time.sleep(1)
        self.log('  PASSO: Tela de nomeacao pronta.')

    def preencher_nomeacao(self, dados):
        """Novo -> PERITO -> profissao -> data -> polo passivo -> concBenef -> avancar."""
        btn_novo = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'formAJIntranet:novo'))
        )
        btn_novo.click()
        self.esperar_ajg_com_timeout(self.driver)
        time.sleep(1)

        Select(WebDriverWait(self.driver, 60).until(
            EC.visibility_of_element_located((By.ID, 'formAJIntranet:id_categoriasProfissionais'))
        )).select_by_visible_text('PERITO')

        profissoes = Select(WebDriverWait(self.driver, 60).until(
            EC.visibility_of_element_located((By.ID, 'formAJIntranet:id_profissoes'))
        ))
        for option in profissoes.options:
            if self.remover_acentos(option.text).lower() == self.remover_acentos(
                    dados.get('profissao', '')).lower():
                profissoes.select_by_visible_text(option.text)
                break
        else:
            raise Exception(f"Profissao '{dados.get('profissao')}' nao encontrada no AJG.")

        if dados.get('data_nomeacao'):
            campo_data = self.driver.find_element(By.ID, 'formAJIntranet:id_dataNomeacao')
            campo_data.clear()
            campo_data.send_keys(dados['data_nomeacao'])

        if 'INSS' in dados.get('polo_passivo', '').upper():
            self.driver.find_element(By.XPATH,
                '/html/body/form/div[3]/span/table/tbody/tr[7]/td[2]/fieldset/span/span/input[1]').click()
        else:
            self.driver.find_element(By.XPATH,
                '/html/body/form/div[3]/span/table/tbody/tr[7]/td[2]/fieldset/span/span/input[2]').click()

        self.driver.find_element(By.ID, 'formAJIntranet:concBenefAssistDeficBenefPrevIncLab').click()
        self.log('  Dados da nomeacao preenchidos.')
        self.driver.find_element(By.ID, 'formAJIntranet:avancar').click()

    def colar_numero_e_tab(self, dados):
        """Cola nr_processo + TAB. Retorna True se for preciso pular o processo."""
        campo = WebDriverWait(self.driver, 60).until(
            EC.visibility_of_element_located((By.ID, 'formAJIntranet:id_NumeroProcessoJudicial'))
        )
        campo.click()
        time.sleep(0.5)

        match_cnj = re.search(r'\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}', dados.get('nr_processo', ''))
        nr_processo_ajg = match_cnj.group(0) if match_cnj else dados.get('nr_processo', '')

        pyperclip.copy(nr_processo_ajg)
        self.log(f'  Clipboard: {nr_processo_ajg}')
        ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        time.sleep(1)

        # alerta apos colar
        if self.handle_alert():
            return True

        try:
            valor_campo = campo.get_attribute('value')
        except StaleElementReferenceException:
            campo = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, 'formAJIntranet:id_NumeroProcessoJudicial'))
            )
            valor_campo = campo.get_attribute('value')
        self.log(f'  Valor no campo apos colar: {valor_campo}')

        if self.digito(valor_campo) != self.digito(nr_processo_ajg):
            self.log('  Colagem nao conferiu, pulando processo...')
            return True

        campo.send_keys(Keys.TAB)
        time.sleep(2)

        # alerta apos TAB
        if self.handle_alert():
            return True

        return False

    def selecionar_parte_ativa(self, dados):
        """Aguarda lista de partes, seleciona autor/requerente/exequente/impetrante
        (ou a unica parte), associa, salva parte_associada no Excel e avanca."""

        def _tem_parte_util(_driver):
            try:
                select_el = _driver.find_element(By.ID, 'formAJIntranet:selectOneListboxAssistido')
                for op in select_el.find_elements(By.TAG_NAME, 'option'):
                    txt = (op.text or '').strip().lower()
                    if txt and txt not in ('selecione', 'selecione...'):
                        return True
                return False
            except (NoSuchElementException, StaleElementReferenceException):
                return False

        try:
            WebDriverWait(self.driver, 20,
                          ignored_exceptions=(StaleElementReferenceException,)).until(_tem_parte_util)
        except TimeoutException:
            self.log('  Lista de partes NAO populou apos TAB. Tentando refresh + TAB...')
            self.driver.refresh()
            time.sleep(2)
            try:
                WebDriverWait(self.driver, 20,
                              ignored_exceptions=(StaleElementReferenceException,)).until(_tem_parte_util)
            except TimeoutException:
                self.log('  ERRO: lista de partes seguiu vazia. Pulando processo.')
                return False

        padrao_parte_ativa = re.compile(r'\b(autora?|requerente|exequente|impetrante)\b', re.IGNORECASE)
        autor_encontrado = False
        nome_parte_selecionada = None

        for _ in range(3):
            try:
                select_el = self.driver.find_element(By.ID, 'formAJIntranet:selectOneListboxAssistido')
                opcoes = select_el.find_elements(By.TAG_NAME, 'option')

                for op in opcoes:
                    texto = (op.text or '').strip()
                    if not texto or texto.lower() in ('selecione', 'selecione...'):
                        continue
                    if padrao_parte_ativa.search(texto):
                        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(op))
                        op.click()
                        autor_encontrado = True
                        nome_parte_selecionada = texto
                        self.log(f'  Parte ativa selecionada: {texto}')
                        break

                if autor_encontrado:
                    break

                partes_reais = [op for op in opcoes
                                if (op.text or '').strip().lower() not in ('selecione', 'selecione...', '')]
                if len(partes_reais) == 1:
                    WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(partes_reais[0]))
                    partes_reais[0].click()
                    autor_encontrado = True
                    nome_parte_selecionada = partes_reais[0].text.strip()
                    self.log(f'  Unica parte real selecionada: {nome_parte_selecionada}')
                    break

            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        if not autor_encontrado:
            self.log('  ERRO: nao foi possivel selecionar parte ativa. Pulando processo.')
            return False

        # Associar parte
        self.driver.find_element(By.ID, 'formAJIntranet:id_0000001009').click()
        self.esperar_ajg_com_timeout(self.driver)
        time.sleep(1)
        self.log('  Parte associada com sucesso.')

        # Salvar parte_associada no resultado.xlsx
        try:
            wb = openpyxl.load_workbook(CAMINHO_RESULTADO)
            ws = wb.active
            cabecalhos = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            col_parte = None
            for idx, col in enumerate(cabecalhos, 1):
                if col == 'parte_associada':
                    col_parte = idx
                    break
            if col_parte is None:
                col_parte = len(cabecalhos) + 1
                ws.cell(row=1, column=col_parte, value='parte_associada')

            nr_alvo = str(dados.get('nr_processo', '')).strip()
            for row in range(2, ws.max_row + 1):
                if ws.cell(row=row, column=1).value and str(ws.cell(row=row, column=1).value).strip() == nr_alvo:
                    ws.cell(row=row, column=col_parte, value=nome_parte_selecionada)
                    self.log(f'  Nome da parte salvo no Excel: {nome_parte_selecionada}')
                    break
            wb.save(CAMINHO_RESULTADO)
            wb.close()
        except Exception as e:
            self.log(f'  ERRO ao salvar parte no Excel (feche o arquivo se estiver aberto): {e}')

        # Avancar para honorarios
        self.driver.find_element(By.ID, 'formAJIntranet:avancar').click()
        self.esperar_ajg_com_timeout(self.driver)
        time.sleep(1)
        self.log('  Avancou para etapa de honorarios.')
        return True

    def preencher_honorarios_e_solicitacao(self, dados):
        """Honorarios -> busca perito -> valor -> concluir -> criar solicitacao -> final."""
        # Honorarios
        Select(self.driver.find_element(By.ID, 'formAJIntranet:id_honorarios')).select_by_visible_text('PERITAS(OS)')
        if dados.get('data_servico'):
            self.driver.find_element(By.ID, 'formAJIntranet:id_dataPrestServ').send_keys(dados['data_servico'])
        self.driver.find_element(By.ID, 'formAJIntranet:localPrestServ').click()
        time.sleep(1)

        # Abrir busca profissional (carrega tabela gaveta_03 na mesma pagina)
        btn_pesq = WebDriverWait(self.driver, 30).until(
            EC.element_to_be_clickable((By.ID, 'formAJIntranet:botaoPesquisarProfissinal'))
        )
        self.driver.execute_script("arguments[0].click();", btn_pesq)

        WebDriverWait(self.driver, 30).until(
            lambda d: len(d.find_element(By.ID, 'gaveta_03').find_elements(By.TAG_NAME, 'tr')) > 1
        )
        time.sleep(1)

        nome_perito = dados.get('nome', '')
        nome_perito_norm = self.remover_acentos(nome_perito).lower()
        tabela_peritos = self.driver.find_element(By.ID, 'gaveta_03')
        linhas = tabela_peritos.find_elements(By.TAG_NAME, 'tr')

        encontrado = False
        for linha in linhas[1:]:  # pula header "Nome do profissional"
            texto = self.remover_acentos(linha.text).lower()
            if nome_perito_norm in texto:
                self.driver.execute_script("arguments[0].click();", linha.find_element(By.TAG_NAME, 'input'))
                encontrado = True
                self.log(f'  Perito selecionado: {nome_perito}')
                break

        if not encontrado:
            for l in linhas:
                self.log(f'  Linha tabela: {l.text}')
            raise Exception(f'Perito "{nome_perito}" nao encontrado na gaveta_03')

        # Confirmar selecao do perito
        try:
            btn_confirmar = self.driver.find_element(By.ID, 'formAJIntranet:pesquisar')
            btn_confirmar.click()
        except NoSuchElementException:
            candidatos = self.driver.find_elements(By.XPATH,
                "//input[@type='button' or @type='submit']"
                "[contains(translate(@value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'confirm') or "
                "contains(translate(@value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'selecion') or "
                "contains(translate(@value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'ok')]")
            if candidatos:
                self.driver.execute_script("arguments[0].click();", candidatos[0])
            else:
                self.log('  AVISO: botao confirmar do perito nao encontrado; tentando avancar...')
        self.esperar_ajg_com_timeout(self.driver)
        time.sleep(1)
        self.log('  Perito confirmado.')

        # Valor estimado honorarios
        self.driver.find_element(By.ID, 'formAJIntranet:avancar').click()
        self.esperar_ajg_com_timeout(self.driver)
        time.sleep(1)

        campo_valor = WebDriverWait(self.driver, 60).until(
            EC.visibility_of_element_located((By.ID, 'formAJIntranet:id_valorEstimadoHonorarios'))
        )
        if dados.get('valor'):
            campo_valor.send_keys(dados['valor'])
        self.log('  Preenchi os dados de honorarios (valor estimado).')

        WebDriverWait(self.driver, 60).until(
            EC.visibility_of_element_located((By.ID, 'formAJIntranet:concluir'))
        ).click()
        self.esperar_ajg_com_timeout(self.driver)
        time.sleep(1)

        WebDriverWait(self.driver, 60).until(
            EC.visibility_of_element_located((By.ID, 'formAJIntranet:criarSolicitacao'))
        ).click()
        self.esperar_ajg_com_timeout(self.driver)
        time.sleep(1)
        self.log('  Criei solicitacao de pagamento.')

        # Dados finais da solicitacao
        WebDriverWait(self.driver, 60).until(
            EC.visibility_of_element_located((By.ID, 'formAJIntranet:dataPrestacao'))
        ).send_keys(dados.get('data_servico', ''))

        self.driver.find_element(By.ID, 'copiar').click()
        self.driver.find_element(By.ID, 'formAJIntranet:motivos2').click()
        self.driver.find_element(By.ID, 'formAJIntranet:decisaoFundamentada').send_keys('.')

        WebDriverWait(self.driver, 60).until(
            EC.visibility_of_element_located((By.ID, 'formAJIntranet:concluir'))
        ).click()
        self.esperar_ajg_com_timeout(self.driver)
        time.sleep(1)
        self.log('  Cliquei em concluir (solicitacao criada).')

        # Voltar ao inicio para o proximo processo
        WebDriverWait(self.driver, 120).until(self.esperar_ajg)
        self.driver.find_element(By.ID, 'formAJIntranet:bt_inicio_sistema').click()
        self.esperar_ajg_com_timeout(self.driver)
        time.sleep(1)

    def run(self):
        self.running = True
        self.done = False

        try:
            self.log('ETAPA 1: Lendo resultado.xlsx...')
            if not os.path.exists(CAMINHO_RESULTADO):
                self.log('ERRO: resultado.xlsx nao encontrado.')
                return

            resultados = self.ler_resultado(CAMINHO_RESULTADO)
            self.log(f'Processos com juntada=OK: {len(resultados)}')

            if not resultados:
                self.log('Nenhum processo com juntada=OK.')
                return

            for idx, r in enumerate(resultados, 1):
                self.log(f'  [{idx}] {r.get("nr_processo")} | {r.get("nome")} | {r.get("profissao")} | data={r.get("data_servico")} | valor={r.get("valor")}')

            self.log('ETAPA 2: Abrindo Chrome AJG...')
            try:
                # Com debugger_address=None, abre Chrome novo (mais rapido)
                self.log('Abrindo Chrome AJG (nova instância)...')
                self.driver = criar_driver()
                self.driver.get(URL_LOGIN_AJG)
                self.log('Chrome AJG aberto com sucesso.')
            except Exception as e:
                self.log(f'Erro ao abrir Chrome AJG: {e}')
                return

            self.log('ETAPA 3: Aguardando login manual...')
            self.log('Faca login e clique em "Ja loguei no AJG".')
            self.login_event.wait()
            if self.stop_event.is_set():
                self.log('Parado pelo usuario.')
                return
            self.log('Login AJG confirmado!')

            self.log('ETAPA 4: Navegando para Nomeacao de Profissionais...')
            self.driver.switch_to.window(self.driver.window_handles[0])
            time.sleep(2)
            self.navegar_para_nomeacao()
            self.log('Tela de Nomeacao de Profissionais carregada!')

            self.log(f'ETAPA 5: Processando {len(resultados)} processo(s)...')
            for i, dados in enumerate(resultados, 1):
                if self.stop_event.is_set():
                    self.log('Execucao interrompida pelo usuario.')
                    break

                nr_processo = dados.get('nr_processo', '').strip()
                self.log(f'--- Processo {i}/{len(resultados)}: {nr_processo} ---')

                try:
                    if i > 1:
                        self.log('  Voltando para Nomeacao de Profissionais...')
                        self.navegar_para_nomeacao()

                    self.log('  [1/6] Novo -> PERITO -> profissao -> data -> avancar...')
                    try:
                        self.preencher_nomeacao(dados)
                    except UnexpectedAlertPresentException:
                        pass

                    self.log('  [2/6] Verificando alerta apos avancar...')
                    alerta = self.handle_alert()
                    if alerta and ('paga' in alerta.lower() or 'nomea' in alerta.lower()):
                        self.log('  ALERTA: nomeacao paga. Cancelando.')
                        self.cancelar_e_voltar_inicio()
                        continue

                    self.log('  [3/6] Aguardando loading apos avancar...')
                    try:
                        WebDriverWait(self.driver, 120).until(self.esperar_ajg)
                    except TimeoutException:
                        self.log('  Timeout no loading. Recarregando...')
                        self.driver.refresh()
                        time.sleep(2)

                    self.log('  [4/6] Colando numero do processo + TAB...')
                    if self.colar_numero_e_tab(dados):
                        self.log('  ALERTA: problema ao colar numero. Cancelando.')
                        self.cancelar_e_voltar_inicio()
                        continue

                    self.log('  [5/6] Selecionando parte ativa...')
                    if not self.selecionar_parte_ativa(dados):
                        self.cancelar_e_voltar_inicio()
                        continue

                    self.log('  [6/6] Preenchendo honorarios + solicitacao...')
                    self.preencher_honorarios_e_solicitacao(dados)
                    self.log(f'  PROCESSO {nr_processo} CONCLUIDO!')

                except Exception as e:
                    import traceback
                    self.log(f'  ERRO no processo {nr_processo}: {e}')
                    self.log(traceback.format_exc())
                    try:
                        self.cancelar_e_voltar_inicio()
                    except Exception:
                        pass
                    continue

            self.log('=== EXECUCAO FINALIZADA ===')

        except Exception as e:
            import traceback
            self.log(f'ERRO GERAL: {type(e).__name__}: {e}')
            self.log(traceback.format_exc())
        finally:
            self.running = False
            self.done = True

    def confirmar_login(self):
        self.login_event.set()

    def fechar_navegador(self):
        if self.driver:
            try:
                self.driver.quit()
                self.log('Navegador fechado.')
            except:
                pass

    def parar(self):
        self.stop_event.set()
        self.login_event.set()
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
        self.driver = None
        # Mata processos Chrome/ChromeDriver em segundo plano
        matar_processo_chrome()
