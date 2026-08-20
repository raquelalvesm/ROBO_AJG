#!/usr/bin/env python
# coding: utf-8
"""
Robo AJG - Nomeacao de profissionais (fluxo completo standalone)

Fluxo por processo (lido de resultado.xlsx):
  1. Novo -> PERITO -> profissao -> data nomeacao -> polo passivo -> concBenef -> avancar
  2. Se alerta apos avancar ("nomeacao paga"): log, cancelar (+ confirmacao) e
     bt_inicio_sistema -> proximo processo
  3. Colar nr_processo + TAB
  4. Se alerta apos colar/TAB: log, cancelar (+ confirmacao) e bt_inicio_sistema -> proximo
  5. Parte ativa: aguardar selectOneListboxAssistido, selecionar autor/requerente/exequente/
     impetrante (ou a unica parte), clicar associar, salvar parte_associada no Excel, avancar
  6. Honorarios: PERITAS(OS), data servico, local, buscar profissional (tabela gaveta_03),
     selecionar perito, confirmar, avancar
  7. Valor estimado -> concluir -> criar solicitacao -> dados finais -> concluir
  8. bt_inicio_sistema -> proximo processo
"""

import os, re, time, unicodedata
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

CAMINHO_SAIDA = os.path.join(base_dir(), 'resultado.xlsx')
URL_LOGIN_AJG = ('https://ajg1.cjf.jus.br/aj/seguranca/efetuarloginintranet/'
                 'efetuarLoginIntranet_efetuarLogin.jsf')


def log(msg):
    linha = f'[{datetime.now().strftime("%H:%M:%S")}] {msg}'
    print(linha)
    escrever_log(f'log_ajg_{datetime.now().strftime("%Y%m%d")}.txt', linha)


def remover_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto)
                   if unicodedata.category(c) != 'Mn')


def esperar_ajg(driver):
    try:
        el = driver.find_element(By.ID, 'loadingDiv')
        return el.value_of_css_property('display') == 'none'
    except (NoSuchElementException, StaleElementReferenceException):
        return True


def esperar_ajg_com_timeout(driver, timeout=120):
    WebDriverWait(driver, timeout).until(esperar_ajg)


def handle_alert(driver):
    """Se houver alerta, registra no log, aceita e retorna o texto; senao None."""
    try:
        alerta = WebDriverWait(driver, 2).until(EC.alert_is_present())
        txt = alerta.text
        log(f'ALERTA: {txt}')
        alerta.accept()
        return txt
    except TimeoutException:
        return None
    except NoAlertPresentException:
        return None


def cancelar_e_voltar_inicio(driver):
    """Clica cancelar (com confirmacao) e bt_inicio_sistema para ir ao proximo processo."""
    try:
        btn_cancelar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'formAJIntranet:cancelar'))
        )
        btn_cancelar.click()
        time.sleep(1)
        handle_alert(driver)  # confirmacao "Deseja cancelar a operacao..."
    except Exception as e:
        log(f'Erro ao clicar cancelar: {e}')

    try:
        btn_inicio = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'formAJIntranet:bt_inicio_sistema'))
        )
        btn_inicio.click()
        esperar_ajg_com_timeout(driver)
        time.sleep(1)
    except Exception as e:
        log(f'Erro ao voltar ao inicio: {e}')


def navegar_para_nomeacao(driver):
    """Abre a tela de nomeacao de profissionais (tambem ao voltar para o inicio
    entre processos)."""
    try:
        el = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'link_tela_com_tooltip_1nomeacaodeprofissionais'))
        )
        el.click()
        log('Clicado em Nomeacao de Profissionais.')
    except TimeoutException:
        links = driver.find_elements(By.TAG_NAME, 'a')
        for l in links:
            lid = l.get_attribute('id') or ''
            ltext = l.text.strip()[:50]
            if 'nomeacao' in lid.lower() or 'nomeacao' in ltext.lower():
                l.click()
                log(f'Clicado em link: id={lid} texto={ltext}')
                break
        else:
            raise Exception('Link de nomeacao de profissionais nao encontrado.')
    esperar_ajg_com_timeout(driver)
    time.sleep(1)


def preencher_nomeacao(driver, dados):
    """Celula 3: Novo -> PERITO -> profissao -> data -> polo passivo -> concBenef -> avancar."""
    btn_novo = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'formAJIntranet:novo'))
    )
    btn_novo.click()
    esperar_ajg_com_timeout(driver)
    time.sleep(1)

    Select(WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.ID, 'formAJIntranet:id_categoriasProfissionais'))
    )).select_by_visible_text('PERITO')

    profissoes = Select(WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.ID, 'formAJIntranet:id_profissoes'))
    ))
    for option in profissoes.options:
        if remover_acentos(option.text).lower() == remover_acentos(dados['profissao']).lower():
            profissoes.select_by_visible_text(option.text)
            break
    else:
        raise Exception(f"Profissao '{dados['profissao']}' nao encontrada no seletor do AJG.")

    driver.find_element(By.ID, 'formAJIntranet:id_dataNomeacao').send_keys(dados['data_nomeacao'])

    if 'INSS' in dados['polo_passivo']:
        driver.find_element(By.XPATH,
            '/html/body/form/div[3]/span/table/tbody/tr[7]/td[2]/fieldset/span/span/input[1]').click()
    else:
        driver.find_element(By.XPATH,
            '/html/body/form/div[3]/span/table/tbody/tr[7]/td[2]/fieldset/span/span/input[2]').click()

    driver.find_element(By.ID, 'formAJIntranet:concBenefAssistDeficBenefPrevIncLab').click()
    log('Dados da nomeacao preenchidos.')

    driver.find_element(By.ID, 'formAJIntranet:avancar').click()


def colar_numero_e_tab(driver, dados):
    """Celula 4: cola nr_processo + TAB, tratando alertas. Retorna True se houve alerta."""
    campo = WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.ID, 'formAJIntranet:id_NumeroProcessoJudicial'))
    )
    campo.click()
    time.sleep(0.5)

    pyperclip.copy(dados['nr_processo'])
    log(f'Clipboard: {dados["nr_processo"]}')
    ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
    time.sleep(1)

    # alerta apos colar
    if handle_alert(driver):
        return True

    try:
        valor_campo = campo.get_attribute('value')
    except StaleElementReferenceException:
        campo = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'formAJIntranet:id_NumeroProcessoJudicial'))
        )
        valor_campo = campo.get_attribute('value')
    log(f'Valor no campo apos colar: {valor_campo}')

    campo.send_keys(Keys.TAB)
    time.sleep(2)

    # alerta apos TAB
    if handle_alert(driver):
        return True

    return False


def selecionar_parte_ativa(driver, dados):
    """Celula 5: aguarda lista de partes, seleciona autor/requerente/exequente/impetrante
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
        WebDriverWait(driver, 20, ignored_exceptions=(StaleElementReferenceException,)).until(_tem_parte_util)
    except TimeoutException:
        log('Lista de partes NAO populou apos TAB. Tentando refresh + TAB...')
        driver.refresh()
        time.sleep(2)
        try:
            WebDriverWait(driver, 20, ignored_exceptions=(StaleElementReferenceException,)).until(_tem_parte_util)
        except TimeoutException:
            log('ERRO: lista de partes seguiu vazia. Pulando processo.')
            return False

    padrao_parte_ativa = re.compile(r'\b(autora?|requerente|exequente|impetrante)\b', re.IGNORECASE)
    autor_encontrado = False
    nome_parte_selecionada = None

    for _ in range(3):
        try:
            select_el = driver.find_element(By.ID, 'formAJIntranet:selectOneListboxAssistido')
            opcoes = select_el.find_elements(By.TAG_NAME, 'option')

            for op in opcoes:
                texto = (op.text or '').strip()
                if not texto or texto.lower() in ('selecione', 'selecione...'):
                    continue
                if padrao_parte_ativa.search(texto):
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(op))
                    op.click()
                    autor_encontrado = True
                    nome_parte_selecionada = texto
                    log(f'Parte ativa selecionada: {texto}')
                    break

            if autor_encontrado:
                break

            partes_reais = [op for op in opcoes
                            if (op.text or '').strip().lower() not in ('selecione', 'selecione...', '')]
            if len(partes_reais) == 1:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable(partes_reais[0]))
                partes_reais[0].click()
                autor_encontrado = True
                nome_parte_selecionada = partes_reais[0].text.strip()
                log(f'Unica parte real selecionada: {nome_parte_selecionada}')
                break

        except StaleElementReferenceException:
            time.sleep(0.5)
            continue

    if not autor_encontrado:
        log('ERRO: nao foi possivel selecionar parte ativa. Pulando processo.')
        return False

    # Associar parte
    driver.find_element(By.ID, 'formAJIntranet:id_0000001009').click()
    esperar_ajg_com_timeout(driver)
    time.sleep(1)
    log('Parte associada com sucesso.')

    # Salvar parte_associada no resultado.xlsx
    try:
        wb = openpyxl.load_workbook(CAMINHO_SAIDA)
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

        nr_alvo = str(dados['nr_processo']).strip()
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value and str(ws.cell(row=row, column=1).value).strip() == nr_alvo:
                ws.cell(row=row, column=col_parte, value=nome_parte_selecionada)
                log(f'Nome da parte salvo no Excel: {nome_parte_selecionada}')
                break
        wb.save(CAMINHO_SAIDA)
        wb.close()
    except Exception as e:
        log(f'ERRO ao salvar parte no Excel (feche o arquivo se estiver aberto): {e}')

    # Avancar para honorarios
    driver.find_element(By.ID, 'formAJIntranet:avancar').click()
    esperar_ajg_com_timeout(driver)
    time.sleep(1)
    log('Avançou para etapa de honorários.')
    return True


def preencher_honorarios_e_solicitacao(driver, dados):
    """Celula 6+: honorarios -> busca perito -> valor -> concluir -> criar solicitacao -> final."""
    # Honorarios
    Select(driver.find_element(By.ID, 'formAJIntranet:id_honorarios')).select_by_visible_text('PERITAS(OS)')
    driver.find_element(By.ID, 'formAJIntranet:id_dataPrestServ').send_keys(dados['data_servico'])
    driver.find_element(By.ID, 'formAJIntranet:localPrestServ').click()
    time.sleep(1)

    # Abrir busca profissional (carrega tabela gaveta_03 na mesma pagina)
    btn_pesq = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.ID, 'formAJIntranet:botaoPesquisarProfissinal'))
    )
    driver.execute_script("arguments[0].click();", btn_pesq)

    WebDriverWait(driver, 30).until(
        lambda d: len(d.find_element(By.ID, 'gaveta_03').find_elements(By.TAG_NAME, 'tr')) > 1
    )
    time.sleep(1)

    nome_perito = dados['nome']
    nome_perito_norm = remover_acentos(nome_perito).lower()
    tabela_peritos = driver.find_element(By.ID, 'gaveta_03')
    linhas = tabela_peritos.find_elements(By.TAG_NAME, 'tr')

    encontrado = False
    for linha in linhas[1:]:  # pula header "Nome do profissional"
        texto = remover_acentos(linha.text).lower()
        if nome_perito_norm in texto:
            driver.execute_script("arguments[0].click();", linha.find_element(By.TAG_NAME, 'input'))
            encontrado = True
            log(f'Perito selecionado: {nome_perito}')
            break

    if not encontrado:
        for l in linhas:
            log(f'Linha tabela: {l.text}')
        raise Exception(f'Perito "{nome_perito}" nao encontrado na gaveta_03')

    # Confirmar selecao do perito
    try:
        btn_confirmar = driver.find_element(By.ID, 'formAJIntranet:pesquisar')
        btn_confirmar.click()
    except NoSuchElementException:
        candidatos = driver.find_elements(By.XPATH,
            "//input[@type='button' or @type='submit']"
            "[contains(translate(@value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'confirm') or "
            "contains(translate(@value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'selecion') or "
            "contains(translate(@value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'ok')]")
        if candidatos:
            driver.execute_script("arguments[0].click();", candidatos[0])
        else:
            log('AVISO: botao confirmar do perito nao encontrado; tentando avancar...')
    esperar_ajg_com_timeout(driver)
    time.sleep(1)
    log('Perito confirmado.')

    # Valor estimado honorarios
    driver.find_element(By.ID, 'formAJIntranet:avancar').click()
    esperar_ajg_com_timeout(driver)
    time.sleep(1)

    campo_valor = WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.ID, 'formAJIntranet:id_valorEstimadoHonorarios'))
    )
    campo_valor.send_keys(dados['valor'])
    log('Preenchi os dados de honorarios (valor estimado).')

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.ID, 'formAJIntranet:concluir'))
    ).click()
    esperar_ajg_com_timeout(driver)
    time.sleep(1)

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.ID, 'formAJIntranet:criarSolicitacao'))
    ).click()
    esperar_ajg_com_timeout(driver)
    time.sleep(1)
    log('Criei solicitacao de pagamento.')

    # Dados finais da solicitacao
    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.ID, 'formAJIntranet:dataPrestacao'))
    ).send_keys(dados['data_servico'])

    driver.find_element(By.ID, 'copiar').click()
    driver.find_element(By.ID, 'formAJIntranet:motivos2').click()
    driver.find_element(By.ID, 'formAJIntranet:decisaoFundamentada').send_keys('.')

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.ID, 'formAJIntranet:concluir'))
    ).click()
    esperar_ajg_com_timeout(driver)
    time.sleep(1)
    log('Cliquei em concluir (solicitacao criada).')

    # Voltar ao inicio para o proximo processo
    WebDriverWait(driver, 120).until(esperar_ajg)
    driver.find_element(By.ID, 'formAJIntranet:bt_inicio_sistema').click()
    esperar_ajg_com_timeout(driver)
    time.sleep(1)


def main():
    driver = criar_driver()
    driver.get(URL_LOGIN_AJG)
    time.sleep(3)
    log('Pagina do AJG aberta. Faca login manualmente.')
    input('Pressione ENTER apos estar logado no AJG...')

    # Navegar para tela de nomeacao de profissionais
    navegar_para_nomeacao(driver)

    # Ler processos do resultado.xlsx
    wb = openpyxl.load_workbook(CAMINHO_SAIDA, read_only=True, data_only=True)
    ws = wb.active
    cabecalhos = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    resultados_excel = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        dados_row = {}
        for col_idx, coluna in enumerate(cabecalhos):
            dados_row[coluna] = str(row[col_idx]) if row[col_idx] else ''
        if dados_row.get('nr_processo'):
            resultados_excel.append(dados_row)
    wb.close()
    log(f'Processos lidos do resultado.xlsx: {len(resultados_excel)}')

    # LOOP PRINCIPAL
    for i, dados in enumerate(resultados_excel, 1):
        log(f'=== Processo {i}/{len(resultados_excel)}: {dados.get("nr_processo")} ===')

        try:
            # Ao voltar para a tela de inicio (processos seguintes), clicar
            # novamente em Nomeacao de Profissionais antes de rodar o fluxo.
            if i > 1:
                navegar_para_nomeacao(driver)

            # 1. Novo -> formulario -> avancar
            preencher_nomeacao(driver, dados)

            # 2. Alerta apos avancar ("nomeacao paga")
            try:
                esperar_ajg_com_timeout(driver)
            except UnexpectedAlertPresentException:
                pass
            if handle_alert(driver):
                log('Alerta apos avancar (nomeacao paga). Cancelando e indo para o proximo.')
                cancelar_e_voltar_inicio(driver)
                continue

            # 3. Colar numero + TAB (com checagem de alerta)
            if colar_numero_e_tab(driver, dados):
                log('Alerta apos colar/TAB. Cancelando e indo para o proximo.')
                cancelar_e_voltar_inicio(driver)
                continue

            # 4. Aguarda loading apos TAB
            try:
                WebDriverWait(driver, 120).until(esperar_ajg)
            except TimeoutException:
                log('Timeout aguardando loading apos TAB. Recarregando...')
                driver.refresh()
                time.sleep(2)

            # 5. Parte ativa
            if not selecionar_parte_ativa(driver, dados):
                cancelar_e_voltar_inicio(driver)
                continue

            # 6. Honorarios + perito + solicitacao
            preencher_honorarios_e_solicitacao(driver, dados)
            log(f'Processo {i} concluido.')

        except Exception as e:
            log(f'ERRO no processo {i}: {e}')
            try:
                cancelar_e_voltar_inicio(driver)
            except Exception:
                pass
            continue

    log('Execucao finalizada.')

    if len(resultados_excel) == 1:
        log('Apenas 1 processo processado. O navegador permanece aberto para conferencia manual.')
    else:
        log(f'{len(resultados_excel)} processos processados.')
    input('Pressione ENTER para fechar o navegador...')
    driver.quit()


if __name__ == '__main__':
    main()
