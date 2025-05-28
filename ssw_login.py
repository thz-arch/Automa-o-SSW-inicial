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

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Salva o log de execução em arquivo para rastreio
import datetime
log_file = f"ssw_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logging.getLogger().addHandler(file_handler)
logging.info(f'Log de execução salvo em {log_file}')

# Configuração do Edge
options = Options()
options.add_argument("--start-maximized")

service = Service(EdgeChromiumDriverManager().install())
driver = webdriver.Edge(service=service, options=options)

logging.info('Abrindo a página de login...')
# Abre a página
driver.get("https://sistema.ssw.inf.br/")

logging.info('Aguardando campos de login ficarem disponíveis...')
# Não é necessário clicar no checkbox para exibir os campos
# Preenche os campos diretamente
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

    # Não precisa mais aguardar a tela 101, pois será redirecionado para ocorrências
    # Remove qualquer espera ou ação extra após clicar no botão de busca

    # Tenta clicar no botão de ocorrências primeiro por ID, se não encontrar, tenta pelo XPath /html/body/form/a[17], garantindo que a janela de ocorrências seja aberta.
    try:
        logging.info('Aguardando botão "Ocorrências" (id=link_ocor)...')
        btn_ocorrencias = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "link_ocor"))
        )
        logging.info('Botão "Ocorrências" encontrado por ID!')
    except Exception:
        logging.info('Botão "Ocorrências" por ID não encontrado, tentando por XPath /html/body/form/a[17]...')
        btn_ocorrencias = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/form/a[17]"))
        )
        logging.info('Botão "Ocorrências" encontrado por XPath!')
    logging.info('Clicando no botão "Ocorrências"...')
    handles_antes = set(driver.window_handles)
    print(f'Janelas antes do clique: {len(handles_antes)}')
    # Tenta executar o JavaScript do onclick diretamente, se existir
    onclick = btn_ocorrencias.get_attribute("onclick")
    if onclick:
        driver.execute_script(onclick)
        print('JavaScript do link "Ocorrências" executado diretamente!')
    else:
        btn_ocorrencias.click()
        print('Clique padrão executado (sem JavaScript customizado).')

    # Aguarda até 10s para ver se uma nova janela foi aberta
    try:
        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(handles_antes))
        handles_depois = set(driver.window_handles)
        print(f'Janelas depois do clique: {len(handles_depois)}')
        nova_handle = list(handles_depois - handles_antes)[0]
        driver.switch_to.window(nova_handle)
        print('Foco alterado para a janela de Ocorrências!')
        print(f'Título da janela de Ocorrências: {driver.title}')
    except Exception:
        print('Nenhuma nova janela detectada. Continuando na aba atual.')
        print(f'Título da aba atual: {driver.title}')

    # NOVO: Listar todos os iframes após o clique em Ocorrências
    print('\n--- Lista de iframes após clicar em Ocorrências ---')
    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    print(f'Total de iframes encontrados: {len(iframes)}')
    for idx, iframe in enumerate(iframes):
        try:
            print(f'Iframe {idx}: name={iframe.get_attribute("name")}, id={iframe.get_attribute("id")}, src={iframe.get_attribute("src")}')
        except Exception as e:
            print(f'Iframe {idx}: erro ao acessar atributos ({e})')
    print('-----------------------------------------------\n')

    # Mostra todas as janelas/abas abertas e seus títulos
    for idx, handle in enumerate(driver.window_handles):
        try:
            driver.switch_to.window(handle)
            print(f'Aba {idx}: Título = {driver.title}')
        except Exception as e:
            print(f'Aba {idx}: Não foi possível acessar o título (janela pode ter sido fechada).')
    # Garante o foco na última janela existente
    try:
        driver.switch_to.window(driver.window_handles[-1])
        print('Foco garantido na última janela/aba!')
        print(f'Título da janela/aba: {driver.title}')
    except Exception:
        print('Não foi possível garantir o foco na última janela/aba (todas podem ter sido fechadas).')

    # Aguarda ação manual ou futura automação
    input('Pressione Enter para encerrar e fechar o navegador...')
    driver.quit()
except Exception as e:
    import traceback
    print('Erro durante a automação!')
    print(traceback.format_exc())
    logging.error(f'Erro durante a automação: {e}')
    driver.quit()
