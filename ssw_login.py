print('==== INICIANDO SSW_LOGIN.PY ====', flush=True)
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
import sys

# Configurações de codificação para saída em UTF-8
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# Função para obter o número da nota fiscal de forma dinâmica

def obter_numero_nota():
    """
    Obtém o número da nota fiscal do arquivo nota_atual.txt.
    """
    try:
        with open('nota_atual.txt', 'r') as f:
            nota = f.read().strip()
            if nota:
                return nota
    except Exception:
        pass
    return None  # Não retorna valor fixo, só retorna se houver nota no arquivo

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
    handles_antes = set(driver.window_handles)
    btn_buscar.click()
    logging.info('Botão de busca clicado!')
    # Aguarda até 15s por uma nova janela/aba
    WebDriverWait(driver, 15).until(lambda d: len(d.window_handles) > len(handles_antes))
    handles_depois = set(driver.window_handles)
    novas = list(handles_depois - handles_antes)
    if novas:
        driver.switch_to.window(list(novas)[-1])
        logging.info(f'Foco alterado para a nova janela/aba da tela intermediária! Título: {driver.title}')
    else:
        logging.error('Nenhuma nova janela/aba detectada após busca!')
        raise Exception('Nenhuma nova janela/aba detectada após busca!')
    # Agora busca as linhas da tela intermediária normalmente
    linhas = driver.find_elements(By.CSS_SELECTOR, 'tr[class*="srtr"]')
    linhas_dados = [l for l in linhas if l.text.strip() != ""]
    logging.info(f"Linhas CTRC encontradas (qtd={len(linhas_dados)}):")
    for idx, l in enumerate(linhas_dados):
        logging.info(f"Linha {idx}: {l.text}")
    if len(linhas_dados) == 0:
        logging.error("Nenhuma linha CTRC encontrada na tela intermediária! Salvando HTML para diagnóstico.")
        with open('ssw_tela_intermediaria_vazia.html', 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        print('Nenhuma linha CTRC encontrada na tela intermediária! Veja ssw_tela_intermediaria_vazia.html')
        raise Exception('Nenhuma linha CTRC encontrada na tela intermediária!')

    if len(linhas_dados) == 0:
        logging.info("Não há tela intermediária de CTRCs, indo direto para a tela da nota.")
        with open('ssw_tela_detalhe_nota_sem_intermediaria.html', 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        # Verifica se realmente está na tela de detalhe da nota (exemplo: procura por algum texto típico)
        page_text = driver.page_source.lower()
        if 'ocorrências' not in page_text and 'destinatário' not in page_text:
            logging.error('Nenhum CTRC encontrado para a nota fiscal pesquisada. Abortando.')
            print('Nenhum CTRC encontrado para a nota fiscal pesquisada. Abortando.')
            raise Exception('Nenhum CTRC encontrado para a nota fiscal pesquisada.')
        else:
            logging.info('Tela de detalhe da nota detectada, seguindo normalmente.')
            clicou_nota = True
    if len(linhas_dados) > 1:
        logging.info(f"Tela intermediária detectada: {len(linhas_dados)} CTRCs. Selecionando o de maior frete.")
        maior_valor = 0
        melhor_link = None
        fretes_log = []
        for idx, tr in enumerate(linhas_dados):
            try:
                tds = tr.find_elements(By.TAG_NAME, 'td')
                link_bin = None
                for a in tr.find_elements(By.TAG_NAME, 'a'):
                    if 'sra2' in a.get_attribute('class'):
                        link_bin = a
                        break
                if not link_bin or len(tds) < 11:
                    logging.info(f"Linha {idx} ignorada (sem link BIN ou colunas insuficientes).")
                    continue
                valor_txt = tds[10].text.strip().replace('.', '').replace(',', '.')
                valor_num = 0
                try:
                    valor_num = float(''.join([c for c in valor_txt if c.isdigit() or c == '.']))
                except Exception as e:
                    logging.warning(f"Erro ao converter valor de frete '{valor_txt}': {e}")
                fretes_log.append(valor_num)
                logging.info(f"Linha {idx}: Frete={valor_num}, Link={link_bin.text.strip()}")
                if valor_num > maior_valor:
                    maior_valor = valor_num
                    melhor_link = link_bin
            except Exception as e:
                logging.warning(f"Erro ao processar linha {idx}: {e}")
        logging.info(f"Valores de frete encontrados: {fretes_log}")
        if melhor_link:
            # Salva handles antes do clique
            handles_antes = set(driver.window_handles)
            melhor_link.click()
            logging.info(f"Linha de maior valor de frete ({maior_valor}) selecionada (click no link).")
            # Aguarda e troca o foco para a nova janela/aba da tela da nota
            WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(handles_antes))
            handles_depois = set(driver.window_handles)
            novas = list(handles_depois - handles_antes)
            if novas:
                driver.switch_to.window(list(novas)[-1])
                logging.info(f'Foco alterado para a janela de detalhe da nota! Título: {driver.title}')
            else:
                logging.warning('Nenhuma nova janela/aba foi aberta após clicar no CTRC!')
            time.sleep(2)
            clicou_nota = True
    elif len(linhas_dados) == 1:
        logging.info("Apenas um CTRC encontrado, indo direto para a tela da nota.")
        clicou_nota = True
    else:
        logging.error("Nenhuma linha encontrada para clicar! (pode ser tela da nota já aberta)")
        clicou_nota = True  # <-- Corrigido: segue normalmente
except Exception as e:
    logging.error(f"Erro ao tentar selecionar a nota de maior valor de frete: {e}")
    clicou_nota = False

if not clicou_nota:
    print('Não foi possível clicar na nota de maior valor. Abortando antes de abrir ocorrências.')
    logging.error('Não foi possível clicar na nota de maior valor. Abortando antes de abrir ocorrências.')
    with open('ssw_tela_erro_clicando_nota.html', 'w', encoding='utf-8') as f:
        f.write(driver.page_source)
    raise Exception('Não foi possível clicar na nota de maior valor.')

# Só espera por nova janela/aba se clicou em uma linha (caso de múltiplas notas)
# if len(linhas) > 1:
#     handles_antes = set(driver.window_handles)
#     WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(handles_antes))
#     handles_depois = set(driver.window_handles)
#     novas = list(handles_depois - handles_antes)
#     if novas:
#         driver.switch_to.window(novas[0])
#         logging.info('Foco alterado para a janela de detalhe da nota!')
#     else:
#         logging.info('Nenhuma nova janela/aba foi aberta após clicar no botão de busca!')
# else:
#     logging.info('Apenas uma nota, permanecendo na janela atual.')

with open('ssw_tela_detalhe_nota.html', 'w', encoding='utf-8') as f:
    f.write(driver.page_source)
print('HTML da tela de detalhe da nota salvo em ssw_tela_detalhe_nota.html')

# Aguarda até que algum link esteja presente na tela (garante carregamento)
try:
    WebDriverWait(driver, 10).until(lambda d: len(d.find_elements(By.TAG_NAME, "a")) > 0)
except Exception:
    logging.warning('Nenhum link <a> encontrado após aguardar carregamento da tela da nota.')

btn_ocorrencias = None
links = driver.find_elements(By.TAG_NAME, "a")
logging.info(f"Links encontrados na tela da nota: {[l.text.strip() for l in links]}")
for link in links:
    if link.text.strip().lower() == "ocorrências":
        btn_ocorrencias = link
        break
if not btn_ocorrencias:
    print('Botão "Ocorrências" não encontrado na tela de detalhes da nota!')
    logging.error('Botão "Ocorrências" não encontrado na tela de detalhes da nota!')
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
    logging.error('Nenhuma nova janela/aba foi aberta após clicar em Ocorrências!')
    with open('ssw_tela_erro_ocorrencias.html', 'w', encoding='utf-8') as f:
        f.write(driver.page_source)
    print('HTML da tela atual salvo em ssw_tela_erro_ocorrencias.html para diagnóstico.')
    raise

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
    # --- A PARTIR DAQUI, NÃO FAZ MAIS FILTRAGEM ANTES DO POP-UP ---
    # Automação do pop-up
    pasta = r'C:\Users\057-PC\Documents\COMPROVANTES'
    # Busca arquivo que contenha a nota em qualquer parte do nome
    # Diagnóstico: loga todos os arquivos da pasta antes do filtro
    todos_arquivos = os.listdir(pasta)
    logging.info(f"Arquivos na pasta antes do filtro: {todos_arquivos}")
    print(f"Arquivos na pasta antes do filtro: {todos_arquivos}")
    # Função para normalizar strings (minúsculo, sem espaços, hífens, zeros à esquerda)
    def normaliza(s):
        return s.lower().replace(' ', '').replace('-', '').lstrip('0')
    nota_normalizada = normaliza(nota)
    # Loga arquivos que contêm a nota, mesmo sem extensão
    arquivos_com_nota = [
        os.path.join(pasta, f)
        for f in todos_arquivos
        if nota_normalizada in normaliza(f)
    ]
    logging.info(f"Arquivos que contêm a nota (qualquer extensão): {arquivos_com_nota}")
    print(f"Arquivos que contêm a nota (qualquer extensão): {arquivos_com_nota}")
    # Filtro padrão: só arquivos com extensão válida
    arquivos = [
        os.path.join(pasta, f)
        for f in todos_arquivos
        if f.lower().endswith((".pdf", ".jpg", ".jpeg", ".png")) and nota_normalizada in normaliza(f)
    ]
    logging.info(f"Nota buscada (normalizada): {nota_normalizada}")
    print(f"Nota buscada (normalizada): {nota_normalizada}")
    logging.info(f"Arquivos encontrados: {arquivos}")
    print(f"Arquivos encontrados: {arquivos}")
    # Preenche o campo do pop-up com o caminho completo do arquivo encontrado
    if arquivos:
        caminho_arquivo = arquivos[0]
        logging.info(f"Arquivo selecionado para upload: {caminho_arquivo}")
        # Garante que o pop-up está em foco e o campo de nome está selecionado
        try:
            import pygetwindow as gw
        except ImportError:
            import subprocess
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pygetwindow'])
            import pygetwindow as gw
        time.sleep(1.2)  # Aguarda o pop-up abrir completamente
        # Tenta ativar a janela do pop-up "Abrir" (ou "Open")
        popup = None
        for w in gw.getAllTitles():
            if "abrir" in w.lower() or "open" in w.lower():
                popup = gw.getWindowsWithTitle(w)[0]
                break
        if popup:
            popup.activate()
            time.sleep(0.5)
        else:
            logging.warning('Pop-up "Abrir" não encontrado para ativar!')
        # Apenas cola o caminho do arquivo e dá enter, assumindo que o campo já está em foco
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyperclip.copy(caminho_arquivo)
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.7)
        campo_nome = ''
        max_wait = 10  # segundos
        waited = 0
        while waited < max_wait:
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
            campo_nome = pyperclip.paste().strip()
            logging.info(f"Campo do pop-up após colar: {campo_nome}")
            if os.path.basename(caminho_arquivo).lower() in campo_nome.lower() or caminho_arquivo.lower() in campo_nome.lower():
                break
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)
            pyperclip.copy(caminho_arquivo)
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.7)
            waited += 0.7
        print('Pressionando Enter para anexar o arquivo...')
        logging.info('Pressionando Enter para anexar o arquivo...')
        pyautogui.press('enter')
        print('Arquivo selecionado e "Abrir" acionado!')
        logging.info('Arquivo selecionado e "Abrir" acionado!')
        print(f'Arquivo referente à nota {nota} anexado!')
        logging.info(f'Arquivo referente à nota {nota} anexado!')
        nome_arquivo = os.path.basename(arquivos[0])
        # Aguarda 10 segundos para garantir que o arquivo foi carregado
        time.sleep(10)
        # Após aguardar, clica no botão de finalizar
        try:
            btn_finalizar = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "9")))
            btn_finalizar.click()
            print('Ocorrência finalizada e pronto para próxima nota!')
            logging.info('Ocorrência finalizada e pronto para próxima nota!')
            # Envia confirmação para o n8n
            try:
                import requests
                n8n_url = 'https://n8n-cloudrun-953483141430.southamerica-east1.run.app/webhook-test/f68e35af-e7a5-4608-990b-254674e1b45a'
                payload = {'nota': nota, 'status': 'anexado'}
                resp = requests.post(n8n_url, json=payload, timeout=10)
                print(f'Confirmação enviada para o n8n! Status: {resp.status_code} - Resposta: {resp.text}')
                logging.info(f'Webhook n8n: status={resp.status_code}, resposta={resp.text}')
            except Exception as e:
                print(f'Falha ao enviar confirmação para o n8n: {e}')
                logging.error(f'Falha ao enviar confirmação para o n8n: {e}')
        except Exception as e:
            print('Não foi possível finalizar a ocorrência automaticamente!')
            logging.error(f'Erro ao finalizar ocorrência: {e}')
    else:
        print(f'Nenhum arquivo encontrado para a nota {nota}! Feche o pop-up manualmente ou selecione o arquivo correto.')
        logging.warning(f'Nenhum arquivo encontrado para a nota {nota}!')
        pyautogui.press('esc')  # Fecha o pop-up do Windows
        print(f'Nenhum arquivo encontrado para a nota {nota}! Pulando para próxima nota.')
        logging.info(f'Nenhum arquivo encontrado para a nota {nota}! Pulando para próxima nota.')
except Exception as e:
    print('Não foi possível enviar o arquivo pelo input de upload!')
    logging.error(f'Erro ao enviar pelo input de upload: {e}')
try:
    # Verifica se o nome do arquivo aparece na tela antes de finalizar
    # logging.info(f'Verificando se o arquivo "{nome_arquivo}" foi anexado corretamente...')
    try:
        WebDriverWait(driver, 10).until(
            lambda d: nome_arquivo.lower() in d.page_source.lower()
        )
        # print(f'Arquivo "{nome_arquivo}" anexado com sucesso!')
        pode_finalizar = True
    except Exception:
        # print(f'ATENÇÃO: O arquivo "{nome_arquivo}" não foi encontrado na tela após o upload!')
        # logging.error(f'Arquivo "{nome_arquivo}" não encontrado na tela após upload.')
        # print('Pulando finalização da ocorrência e seguindo para próxima nota.')
        pode_finalizar = False
    if not pode_finalizar:
        # Pula a finalização e volta ao início do while True
        pass
    else:
        # Clica no botão <9> para finalizar ocorrência automaticamente
        btn_finalizar = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "9")))
        btn_finalizar.click()
        print('Ocorrência finalizada e pronto para próxima nota!')
        # Envia confirmação para o n8n
        try:
            import requests
            n8n_url = 'https://n8n-cloudrun-953483141430.southamerica-east1.run.app/webhook-test/f68e35af-e7a5-4608-990b-254674e1b45a'  # Substitua pelo endpoint real do n8n
            payload = {'nota': nota, 'status': 'anexado'}
            requests.post(n8n_url, json=payload, timeout=5)
            print('Confirmação enviada para o n8n!')
        except Exception as e:
            print(f'Falha ao enviar confirmação para o n8n: {e}')
except Exception as e:
    print('Não foi possível finalizar a ocorrência automaticamente!')
    logging.error(f'Erro ao finalizar ocorrência: {e}')
except Exception as e:
    print('Não foi possível interagir com o campo código ou clicar no botão "Buscar no meu micro"!')
    logging.error(f'Erro ao interagir na tela de ocorrências: {e}')
    with open('ssw_tela_erro_ocorrencias.html', 'w', encoding='utf-8') as f:
        f.write(driver.page_source)
    print('HTML da tela atual salvo em ssw_tela_erro_ocorrencias.html para diagnóstico.')
    print(f"Ocorreu um erro: {e}")
    raise
