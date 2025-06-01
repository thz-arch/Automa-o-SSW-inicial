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
import pyperclip
import os

# Função para obter o número da nota fiscal de forma dinâmica

def obter_numero_nota():
    """
    Obtém o número da nota fiscal de forma dinâmica.
    Futuramente, pode ser adaptada para buscar via API/POST.
    """
    try:
        with open('nota_atual.txt', 'r') as f:
            nota = f.read().strip()
            if nota:
                return nota
    except Exception:
        pass
    # Fallback para valor fixo (pode ser removido no futuro)
    return "315550"

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

try:
    WebDriverWait(driver, 30).until(lambda d: "menu" in d.page_source.lower() or "menu" in d.current_url.lower())
    logging.info('Login realizado! Redirecionando para a aba menu...')
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

    logging.info('Aguardando redirecionamento para a tela 101...')
    WebDriverWait(driver, 30).until(lambda d: len(d.window_handles) > 1 or "nota fiscal" in d.page_source.lower())
    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])
        logging.info('Redirecionado para nova janela/aba da tela 101!')
    else:
        logging.info('Redirecionamento para tela 101 na mesma janela!')

    # --- NOVO: Preenchimento dos campos de busca ---
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
    nota = obter_numero_nota()
    logging.info(f'Preenchendo campo de nota fiscal com "{nota}"...')
    driver.execute_script(
        "var campo = arguments[0];"
        "campo.focus();"
        "campo.value = arguments[1];"
        "campo.dispatchEvent(new Event('input', {bubbles:true}));"
        "campo.dispatchEvent(new Event('change', {bubbles:true}));"
        "campo.blur();",
        campo_nota, nota
    )
    logging.info('Campo de nota fiscal preenchido!')

    logging.info('Aguardando botão de busca (id=4)...')
    btn_buscar = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.ID, "4"))
    )
    logging.info('Botão de busca encontrado!')
    logging.info('Clicando no botão de busca...')
    btn_buscar.click()
    logging.info('Botão de busca clicado!')

    handles_antes = set(driver.window_handles)
    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(handles_antes))
    handles_depois = set(driver.window_handles)
    novas = list(handles_depois - handles_antes)
    if novas:
        driver.switch_to.window(novas[0])
        logging.info('Foco alterado para a janela de detalhe da nota!')
    else:
        logging.info('Nenhuma nova janela/aba foi aberta após clicar no botão de busca!')

    with open('ssw_tela_detalhe_nota.html', 'w', encoding='utf-8') as f:
        f.write(driver.page_source)
    print('HTML da tela de detalhe da nota salvo em ssw_tela_detalhe_nota.html')

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

    # --- Interação na tela de ocorrências: digitar 96 no campo código, dar Enter e clicar em "buscar no meu micro" com pyautogui ---
    try:
        campo_codigo = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "3")))
        campo_codigo.clear()
        campo_codigo.send_keys("96")
        campo_codigo.send_keys("\ue007")  # Tecla Enter
        logging.info('Ocorrência "96" enviada no campo código!')
        time.sleep(2)
        try:
            btn_buscar_micro = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "buscarBtn")))
            btn_buscar_micro.click()
            print('Clique realizado no botão "Buscar no meu micro" com Selenium!')
            time.sleep(2)
        except Exception as e:
            print('Clique com Selenium falhou, tentando com pyautogui...')
            btn_buscar_micro = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "buscarBtn")))
            driver.execute_script("arguments[0].scrollIntoView();", btn_buscar_micro)
            time.sleep(1)
            loc = btn_buscar_micro.location_once_scrolled_into_view
            size = btn_buscar_micro.size
            window_pos = driver.get_window_position()
            window_x, window_y = window_pos['x'], window_pos['y']
            click_x = window_x + loc['x'] + size['width'] // 2
            click_y = window_y + loc['y'] + size['height'] // 2
            pyautogui.moveTo(click_x, click_y, duration=0.5)
            pyautogui.click()
            print('Clique realizado no botão "Buscar no meu micro" com pyautogui!')
        time.sleep(2)
        pasta = r'G:\Custos Extras\COMPROVANTE DE DESCARGA'
        arquivos = [f for f in os.listdir(pasta) if nota in f and f.lower().endswith((".pdf", ".jpg", ".jpeg", ".png"))]
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(0.5)
        pyperclip.copy(nota)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(1.5)
        pyautogui.press('down')
        pyautogui.press('enter')
        print(f'Nota {nota} pesquisada e selecionada no pop-up!')
        if not arquivos:
            print(f'Arquivo da nota {nota} não encontrado na pasta!')
            arquivo = input('Se não anexou, digite ou cole o caminho COMPLETO do arquivo para anexar (ou pressione Enter para cancelar): ').strip()
            if not arquivo:
                print('Nenhum arquivo informado. Pulando esta nota.')
                logging.warning(f'Arquivo da nota {nota} não informado manualmente.')
            else:
                pyperclip.copy(arquivo)
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(1)
                pyautogui.press('enter')
                print(f'Arquivo {arquivo} selecionado no pop-up!')
                try:
                    input_upload = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "uploadBtn")))
                    input_upload.send_keys(arquivo)
                    print(f'Arquivo {arquivo} enviado pelo input de upload na tela de ocorrências!')
                except Exception as e:
                    print('Não foi possível enviar o arquivo pelo input de upload!')
                    logging.error(f'Erro ao enviar pelo input de upload: {e}')
        else:
            arquivo = os.path.abspath(os.path.join(pasta, arquivos[0]))
            try:
                input_upload = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "uploadBtn")))
                input_upload.send_keys(arquivo)
                print(f'Arquivo {arquivo} enviado pelo input de upload na tela de ocorrências!')
            except Exception as e:
                print('Não foi possível enviar o arquivo pelo input de upload!')
                logging.error(f'Erro ao enviar pelo input de upload: {e}')
            try:
                btn_finalizar = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "4")))
                input('Valide os dados na tela. Pressione Enter para finalizar a ocorrência e prosseguir...')
                btn_finalizar.click()
                print('Ocorrência finalizada e pronto para próxima nota!')
            except Exception as e:
                print('Não foi possível finalizar a ocorrência automaticamente!')
                logging.error(f'Erro ao finalizar ocorrência: {e}')
    except Exception as e:
        print('Não foi possível interagir com o campo código ou clicar no botão "Buscar no meu micro"!')
        logging.error(f'Erro ao interagir na tela de ocorrências: {e}')
        print(f"Ocorreu um erro: {e}")
finally:
    input('Pressione Enter para fechar o navegador e encerrar o script...')
    logging.info('Fechando o navegador...')
    driver.quit()
    logging.info('Navegador fechado.')
