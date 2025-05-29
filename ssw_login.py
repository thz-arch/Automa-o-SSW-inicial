from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
import logging
import datetime
import pyautogui

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Salva o log de execução em arquivo para rastreio
log_file = f"ssw_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logging.getLogger().addHandler(file_handler)
logging.info(f'Log salvo em {log_file}')

# Configuração do Edge
options = Options()
options.add_argument("--start-maximized")

service = Service(EdgeChromiumDriverManager().install())
driver = webdriver.Edge(service=service, options=options)

logging.info('Abrindo a página de login...')
# Abre a página
driver.get("https://sistema.ssw.inf.br/")

wait = WebDriverWait(driver, 30)
wait.until(EC.presence_of_element_located((By.ID, "1")))  # Domínio
wait.until(EC.presence_of_element_located((By.ID, "2")))  # CPF
wait.until(EC.presence_of_element_located((By.ID, "3")))  # Usuário
wait.until(EC.presence_of_element_located((By.ID, "4")))  # Senha

logging.info('Preenchendo campos de login...')
for campo, valor in [
    ("1", "BIN"),           # Domínio
    ("2", "46833604894"),   # CPF
    ("3", "thiago"),        # Usuário
    ("4", "178523"),        # Senha
]:
    logging.info(f'Preenchendo campo {campo}...')
    driver.execute_script(
        "var campo = document.getElementById(arguments[0]);"
        "campo.focus();"
        "campo.value = arguments[1];"
        "campo.dispatchEvent(new Event('input', {bubbles:true}));"
        "campo.dispatchEvent(new Event('change', {bubbles:true}));"
        "campo.blur();",
        campo, valor
    )

logging.info('Clicando no botão de login...')
wait.until(EC.element_to_be_clickable((By.ID, "5"))).click()  # Botão de login

# Aguarda o redirecionamento e tenta acessar o menu
try:
    WebDriverWait(driver, 30).until(lambda d: "menu" in d.page_source.lower() or "menu" in d.current_url.lower())
    logging.info('Login realizado! Redirecionando para a aba menu...')
    # Aguarda o campo de tela na aba menu
    campo_tela = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "3")))
    logging.info('Preenchendo campo de tela com "101"...')
    driver.execute_script(
        "var campo = document.getElementById(arguments[0]);"
        "campo.focus();"
        "campo.value = arguments[1];"
        "campo.dispatchEvent(new Event('input', {bubbles:true}));"
        "campo.dispatchEvent(new Event('change', {bubbles:true}));"
        "campo.blur();",
        "3", "101"
    )
    logging.info('Campo de tela preenchido com sucesso!')

    # Aguarda o redirecionamento para a tela 101 (pode abrir em nova janela/aba)
    logging.info('Aguardando redirecionamento para a tela 101...')
    WebDriverWait(driver, 30).until(lambda d: len(d.window_handles) > 1 or "nota fiscal" in d.page_source.lower())
    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])
        logging.info('Redirecionado para nova janela/aba da tela 101!')
    else:
        logging.info('Redirecionamento para tela 101 na mesma janela!')

    # --- NOVO: Preenchimento dos campos de busca ---
    # Sempre usar o ano '23' para o campo de data inicial
    data_ini = (datetime.datetime.now() - datetime.timedelta(days=730)).strftime("%d%m23")
    logging.info('Aguardando campo de data inicial (t_data_ini)...')
    campo_data = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "t_data_ini")))
    logging.info('Campo de data inicial encontrado!')
    logging.info(f'Preenchendo campo de data inicial com "{data_ini}"...')
    driver.execute_script(
        "var campo = document.getElementById(arguments[0]);"
        "campo.focus();"
        "campo.value = arguments[1];"
        "campo.dispatchEvent(new Event('input', {bubbles:true}));"
        "campo.dispatchEvent(new Event('change', {bubbles:true}));"
        "campo.blur();",
        "t_data_ini", data_ini
    )
    logging.info('Campo de data inicial preenchido!')

    logging.info('Aguardando campo de nota fiscal (t_nro_nf)...')
    campo_nota = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "t_nro_nf")))
    logging.info('Campo de nota fiscal encontrado!')
    logging.info('Preenchendo campo de nota fiscal com "315550"...')
    driver.execute_script(
        "var campo = arguments[0];"
        "campo.focus();"
        "campo.value = arguments[1];"
        "campo.dispatchEvent(new Event('input', {bubbles:true}));"
        "campo.dispatchEvent(new Event('change', {bubbles:true}));"
        "campo.blur();",
        campo_nota, "315550"
    )
    logging.info('Campo de nota fiscal preenchido!')

    # Clica no botão correto após preencher os campos
    logging.info('Aguardando botão de busca (id=4)...')
    btn_buscar = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.ID, "4"))
    )
    logging.info('Botão de busca encontrado!')
    logging.info('Clicando no botão de busca...')
    btn_buscar.click()
    logging.info('Botão de busca clicado!')

    # Aguarda nova janela/aba após clicar no botão de busca
    handles_antes = set(driver.window_handles)
    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(handles_antes))
    handles_depois = set(driver.window_handles)
    novas = list(handles_depois - handles_antes)
    if novas:
        driver.switch_to.window(novas[0])
        logging.info('Foco alterado para a janela de detalhe da nota!')
    else:
        logging.info('Nenhuma nova janela/aba foi aberta após clicar no botão de busca!')

    # Salva o HTML da tela de detalhe da nota (nova janela)
    with open('ssw_tela_detalhe_nota.html', 'w', encoding='utf-8') as f:
        f.write(driver.page_source)
    print('HTML da tela de detalhe da nota salvo em ssw_tela_detalhe_nota.html')

    # --- Busca botão Ocorrências apenas pelo texto exato na tela de detalhes da nota ---
    btn_ocorrencias = None
    links = driver.find_elements(By.TAG_NAME, "a")
    for link in links:
        if link.text.strip().lower() == "ocorrências":
            btn_ocorrencias = link
            break
    if not btn_ocorrencias:
        print('Botão "Ocorrências" não encontrado na tela de detalhes da nota!')
        input('Pressione Enter para encerrar e fechar o navegador...')
        driver.quit()
        exit()

    handles_antes = set(driver.window_handles)
    try:
        btn_ocorrencias.click()
        time.sleep(2)
    except Exception as e:
        driver.execute_script("arguments[0].scrollIntoView();", btn_ocorrencias)
        time.sleep(1)
        loc = btn_ocorrencias.location_once_scrolled_into_view
        size = btn_ocorrencias.size
        window_pos = driver.get_window_position()
        window_x, window_y = window_pos['x'], window_pos['y']
        click_x = window_x + loc['x'] + size['width'] // 2
        click_y = window_y + loc['y'] + size['height'] // 2
        pyautogui.moveTo(click_x, click_y, duration=0.5)
        pyautogui.click()
        time.sleep(2)
    # Aguarda nova janela/aba
    try:
        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(handles_antes))
        handles_depois = set(driver.window_handles)
        novas = list(handles_depois - handles_antes)
        if not novas:
            raise Exception('Nenhuma nova janela/aba foi aberta após clicar em Ocorrências!')
        driver.switch_to.window(novas[0])
        print('Foco alterado para a janela de Ocorrências!')
        print(f'Título da janela de Ocorrências: {driver.title}')
        with open('ssw_tela_ocorrencias.html', 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        print('HTML da tela de ocorrências salvo em ssw_tela_ocorrencias.html')
    except Exception as e:
        print('Nenhuma nova janela/aba foi aberta após clicar em Ocorrências!')

    input('Pressione Enter para encerrar e fechar o navegador...')
    driver.quit()
except Exception as e:
    import traceback
    print('Erro durante a automação!')
    print(traceback.format_exc())
    logging.error(f'Erro durante a automação: {e}')
    driver.quit()
