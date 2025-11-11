import os
import time
import threading
import sqlite3
import re
import json
from datetime import datetime
import traceback
from selenium import webdriver
from path.paths import paths
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    WebDriverException,
    ElementNotInteractableException,
    JavascriptException,
)
from processos.process_web import (
    confirmando_wa_tgrid, confirmando_wa_tmsselbr, expand_shadow_element, shadow_button, shadow_input, 
    wait_for_element, click_element, confirmar_element, selecionar_elemento, acessa_container, compara_data,
    clicar_elemento_shadow_dom, verificar_situacao, clicar_repetidamente, definir_nfe, wait_for_click, normal_input,
    button, acessar_valor, tentar_alterar_valor, usar_gatilho, gatilho_erro, confirma_valor, altera_nota, shadow_input_quant,
    confirma_valor_quant,
)
from utils.services import NotaServico

# Variável global para rastrear o número de tentativas
tentativas = 0
limite_tentativas = 3
# Variáveis globais de controle
monitoring = True
connection_successful = False
filial_selector = paths["filial_container"]
unidade_selector = paths["enter_unidade"]
data_selector = paths["data_container"]
amb_selector = paths["ambiente_container"]
cnpj_selector = paths["cnpj_container"]
input_pesquisa = paths["pesquisa_cnpj"]
filial_unidade = paths["confirma_unidade"]
btn_filial_unidade = paths["btn_unidade"]
btn_ok_cnpj = paths["btn_ok_cnpj"]
menu_pagto = paths["pesquisa_pagto"]
btn_ok_pagto_nat = paths["btn_ok_pagto_nat"]
unidades = ['0102', '0103', '0104']

# Função para carregar os dados do JSON
def carregar_dados(json_path):
    # Carregando os dados
    try:
        
        if os.path.exists(json_path):
            with open(json_path, "r", encoding="utf-8") as f:
                cnpj_dict = json.load(f)
            print("Dados carregados com sucesso!")
            return cnpj_dict
        
        return {}
    except FileNotFoundError:
        print("Arquivo não encontrado.")
    except json.JSONDecodeError:
        print("Erro ao decodificar o JSON.")

# Função para salvar os dados no JSON
def salvar_dados(dados, json_path):
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)

# Função para ajustar o número da nota com base na unidade
def ajustar_numero_nota(numero_nota, unidade):
    if unidade == 1:  # São Luís
        return numero_nota.lstrip('0')  # Remove zeros à esquerda
    elif unidade == 2:  # Parauapebas
        return numero_nota[-4:]  # Extrai os últimos 4 dígitos
    else:
        raise ValueError("Unidade desconhecida")

def atualizar_status(numero_nota, pasta_nfe):
    numero_nota_sem_zeros = numero_nota.lstrip('0')  # Remover os zeros à esquerda para a busca

    for arquivo in os.listdir(pasta_nfe):
        if arquivo.startswith(f"NFE {numero_nota_sem_zeros}"):
            if arquivo.endswith('X.pdf'):
                return "Inserido" # Retorna o número formatado
            return "Encontrado"
    return "Não encontrado" # Valor padrão se nenhuma condição for atendida

def criar_servico(cnpj, cond_pagto, natureza, osKairos, preco, numero_nota, data):
    servico = NotaServico(cnpj, cond_pagto, natureza, osKairos, preco, numero_nota, data)
    return servico

def configurar_driver():
    """
    Configura e retorna o WebDriver para o Chrome.
    """
    # Configurações do navegador
    chrome_options = Options()
    #chrome_options.add_argument("--headless=new")  # modo invisível
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--ignore-ssl-errors")
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--disable-site-isolation-trials")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--start-maximized")  # Tela cheia
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    return driver

def abrir_site(driver, url, log_queue):
    """
    Inicializa o navegador, acessa o site especificado e realiza interações iniciais necessárias.
    """
    try:
        driver.get(url)
        log_queue.put(f"\nSite acessado: {url}")
        print(f"Site acessado: {url}")
        # Lógica para interagir com elementos na página
        return True
    except Exception as e:
        print(f"Erro ao abrir o site: {e}")
        log_queue.put(f"\nErro ao abrir o site: {e}")
        return False

def fechar_site(driver):
    """
    Fecha o navegador, encerra o site especificado e realiza interações finais necessárias.
    """
    global monitoring
    monitoring = False
    driver.quit()

def iniciar_driver(unidade, db_nome, mes_ano, log_queue, mes_selecionado):
    """
    Inicia o WebDriver, acessa o site e executa o processo principal.
    Tenta novamente até 10 vezes em caso de erro.
    """
    max_tentativas = 10
    tentativa = 1

    while tentativa <= max_tentativas:
        print(f"\n🔄 Tentativa {tentativa}/{max_tentativas} de iniciar o processo...")
        log_queue.put(f"\n🔄 Tentativa {tentativa}/{max_tentativas} de iniciar o processo...")
        driver = None

        try:
            # Buscar notas com status "Encontrado"
            notas_db = carregar_notas(db_nome, mes_ano)

            if not notas_db:
                log_queue.put("\nNenhuma nota com status 'Encontrado' encontrada.")
                print("Nenhuma nota com status 'Encontrado' encontrada.")
                return
            
            driver = configurar_driver()
            url = "link"

            site_aberto = abrir_site(driver, url, log_queue)
            if not site_aberto:
                raise Exception("Falha ao abrir o site.")

            print("✅ Site acessado com sucesso, prosseguindo com a lógica...")
            log_queue.put("\n✅ Site acessado com sucesso, prosseguindo com a lógica...")
            sucesso = main_process(driver, url, db_nome, unidade, mes_ano, log_queue, mes_selecionado)

            if sucesso:
                print("✅ Processamento concluído com sucesso.")
                log_queue.put("✅ Processamento concluído com sucesso.")
                if driver:
                    try:
                        driver.quit()
                        print("🛑 Driver finalizado.")
                        log_queue.put("🛑 Driver finalizado.")
                    except Exception as e:
                        print(f"⚠️ Erro ao finalizar driver: {e}")
                        log_queue.put(f"⚠️ Erro ao finalizar driver: {e}")

                return # Sai do loop com sucesso
            else:
                print(f"❌ Erro na tentativa {tentativa}: {e}")
                log_queue.put(f"❌ Erro na tentativa {tentativa}: {e}")
                tentativa += 1
                if driver:
                    try:
                        driver.quit()
                        print("🛑 Driver finalizado.")
                        log_queue.put("🛑 Driver finalizado.")
                    except Exception as e:
                        print(f"⚠️ Erro ao finalizar driver: {e}")
                        log_queue.put(f"⚠️ Erro ao finalizar driver: {e}")

        except Exception as e:
            print(f"❌ Erro na tentativa {tentativa}: {e}")
            log_queue.put(f"❌ Erro na tentativa {tentativa}: {e}")
            tentativa += 1
            if driver:
                try:
                    driver.quit()
                    print("🛑 Driver finalizado.")
                    log_queue.put("🛑 Driver finalizado.")
                except Exception as e:
                    print(f"⚠️ Erro ao finalizar driver: {e}")
                    log_queue.put(f"⚠️ Erro ao finalizar driver: {e}")

            time.sleep(3)  # Pequena pausa antes de tentar novamente

def monitor_connection_thread(driver, url, log_queue, stop_monitoring):
    """
    Inicia a thread de monitoramento da conexão.
    """
    monitor_thread = threading.Thread(target=monitor_connection, args=(driver, url, log_queue, stop_monitoring))
    monitor_thread.start()
    return monitor_thread

def monitor_connection(driver, url, log_queue, stop_monitoring, max_attempts=5, check_interval=5):
    """
    Monitora a conexão em segundo plano e retenta se houver erro.
    Para a thread se `stop_monitoring` for acionado.
    """
    global connection_successful
    attempt = 0

    while not stop_monitoring.is_set() and attempt < max_attempts and not connection_successful:
        try:
            log_queue.put(f"[Monitor] Tentativa {attempt + 1} de {max_attempts} para acessar {url}...")
            print(f"[Monitor] Tentativa {attempt + 1} de {max_attempts} para acessar {url}...")

            # Aguarda a página carregar um elemento essencial
            wait_for_element(driver, By.CSS_SELECTOR, "wa-dialog.startParameters")
            log_queue.put("[Monitor] Conexão bem-sucedida!")
            print("[Monitor] Conexão bem-sucedida!")
            connection_successful = True
            return  # Sai da função ao conectar com sucesso

        except Exception as e:
            log_queue.put(f"[Monitor] Erro ao tentar conectar: {e}")
            print(f"[Monitor] Erro ao tentar conectar: {e}")
            attempt += 1
            time.sleep(check_interval)

    if not connection_successful:
        log_queue.put("[Monitor] Falha ao conectar após todas as tentativas.")
        print("[Monitor] Falha ao conectar após todas as tentativas.")

def fechar_iframe(driver, log_queue):
    """
    Função para fechar o iframe acessado voltando para o documento principal do contexto.
    """
    try:
        driver.switch_to.default_content()
        print("Contexto retornado para o documento principal.")
        log_queue.put("Contexto retornado para o documento principal.")
    except Exception as e:
        print(f"Erro ao fechar o iframe: {e}")
        log_queue.put(f"Erro ao fechar o iframe: {e}")

def process_shadow_dom(driver, log_queue):
    """
    Processa interações no Shadow DOM para clicar no botão OK e localizar outros elementos.
    """
    print("Selecionando tipo de ambiente no servidor...")
    log_queue.put("\nSelecionando tipo de ambiente no servidor...")

    # Localiza o combobox dentro do Shadow DOM
    wa_combo_box = wait_for_element(
        driver, 
        By.CSS_SELECTOR, 
        'wa-dialog.startParameters > fieldset[id="fieldsetEnv"] > wa-combobox[id="selectEnv"]'
    )
    shadow_combo_box = expand_shadow_element(driver, wa_combo_box)
    select_element = shadow_combo_box.find_element(By.CSS_SELECTOR, "select")

    # Opção desejada
    desired_value = "czls4f_prod"
    desired_option = select_element.find_element(By.CSS_SELECTOR, f"option[value='{desired_value}']")

    # Opção atual selecionada
    current_option = select_element.find_element(By.CSS_SELECTOR, "option:checked")
    current_value = current_option.get_attribute("value")

    if current_value == desired_value:
        print(f"\nAmbiente '{desired_value}' já está selecionado, não será alterado.")
        log_queue.put(f"\nAmbiente '{desired_value}' já está selecionado, não será alterado.")
    else:
        desired_option.click()
        print(f"\nAmbiente '{desired_value}' selecionado.")
        log_queue.put(f"\nAmbiente '{desired_value}' selecionado.")

    time.sleep(1)

    print("\nAguardando wa-dialog...")
    log_queue.put("\nAguardando wa-dialog...")
    shadow_button(driver, "wa-dialog.startParameters", "wa-button[title='Botão confirmar']", log_queue)

    time.sleep(3)

def locate_and_access_iframe(driver, log_queue):
    """
    Localiza o iframe dentro do Shadow DOM e alterna para ele.
    """
    print("Aguardando próximo wa-dialog...")
    log_queue.put("\nAguardando próximo wa-dialog do iFrame...")
    
    wa_dialog_2 = wait_for_element(driver, By.ID, 'COMP3000')
    print("Acessando o wa-image...")
    log_queue.put("Acessando o wa-image...")
    
    wa_image_1 = wait_for_element(wa_dialog_2, By.ID, 'COMP3008')
    print("Acessando o wa-webview...")
    log_queue.put("Acessando o wa-webview...")
    
    wa_webview_1 = wait_for_element(wa_image_1, By.ID, 'COMP3010')
    print("Acessando shadow root do webview...")
    log_queue.put("Acessando shadow root do webview...")
    
    shadow_root_2 = expand_shadow_element(driver, wa_webview_1)
    print("Acessando o iframe dentro do shadowRoot...")
    log_queue.put("Acessando o iframe dentro do shadowRoot...")
    iframe = wait_for_element(shadow_root_2, By.CSS_SELECTOR, 'iframe[src*="kairoscomercio136240.protheus.cloudtotvs.com.br"]')

    if iframe:
        print("Iframe localizado com sucesso.")
        log_queue.put("Iframe localizado com sucesso.")
        driver.switch_to.frame(iframe)
        print("Dentro do iframe.")
        log_queue.put("Dentro do iframe.")
    else:
        raise Exception("Iframe não encontrado.")

def perform_login(driver, login, password, log_queue):
    """
    Preenche os campos de login e senha e realiza a autenticação.
    """
    try:
        normal_input(driver, '.po-field-container-content', '[name="login"]', login, "User",log_queue)
        
        normal_input(driver, '[name="password"]', 'input[name="password"]', password, "Password", log_queue)

        time.sleep(2)
        button_enter = wait_for_element(driver, By.CSS_SELECTOR, 'po-button')
        click_element(button_enter, (By.CSS_SELECTOR, "button.po-button[p-kind=primary]"))
        print("Botão Entrar clicado com sucesso!")
        log_queue.put("Botão Entrar clicado com sucesso!")
        time.sleep(2)
    except Exception as e:
        print(f"Erro durante o login: {e}")
        log_queue.put(f"Erro durante o login: {e}")

def abrir_menu_unidade(driver, unidade, data, log_queue):
    """
    Função inicial para inserir a data e filial correta que deseja(Tela inicial).
    """
    print("\nAcessando container da data...")
    log_queue.put("Acessando container da data...")
    
    normal_input(driver, data_selector, 'input', data, "Data", log_queue)
    
    print("Data retroagida ou inserida.")
    log_queue.put("Data retroagida ou inserida.")

    print("Acessando container da filial...")
    log_queue.put("\nAcessando container da filial...")
    
    normal_input(driver, filial_selector, 'input', unidades[unidade-1], "Filial", log_queue)

    print("Acessando ambiente 05...")
    log_queue.put("\nAcessando ambiente 05...")
    
    container_amb = wait_for_element(driver, By.CSS_SELECTOR, amb_selector)
    WebDriverWait(driver, 20).until(EC.visibility_of(container_amb))
    amb_field = wait_for_click(container_amb, By.CSS_SELECTOR, 'input')

    # Garantir que o elemento esteja visível
    driver.execute_script("arguments[0].scrollIntoView(true);", amb_field)
    normal_input(driver, amb_selector, 'input', '5', "Ambiente", log_queue)

    amb_field.send_keys(Keys.TAB)

    print("Acesso Concluído.")
    log_queue.put("Acesso Concluído.")
    # Procurando e clicando no botão
    container_but = wait_for_element(driver, By.CSS_SELECTOR, unidade_selector)

    ActionChains(driver).move_to_element(container_but).perform()
    print("Busca do container do botão Enter completa.")
    log_queue.put("Busca do container do botão Enter completa.")
    click_element(container_but, (By.CSS_SELECTOR, "button"))
    print("Botão de entrar na unidade clicado com sucesso!")
    log_queue.put("Botão de entrar na unidade clicado com sucesso!")

    fechar_iframe(driver, log_queue)

    time.sleep(5)

def rotina_venda(driver, log_queue):
    """
    Função que após apertar o botão de Favoritos acessa a rotina Pedidos de Venda.
    """
    print("Buscando pesquisa de rotina.")
    log_queue.put("\nBuscando pesquisa de rotina.")
    campo_rotina = wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP3053"] > wa-text-input[id="COMP3056"]')
    shadow_input(driver, 'wa-panel[id="COMP3053"] > wa-text-input[id="COMP3056"]', "Pedidos de Venda", log_queue)

    valor_atual = acessar_valor(campo_rotina).strip()
    print(f"Valor atual do campo: {valor_atual}")
    log_queue.put(f"Valor atual do campo: {valor_atual}")

    if valor_atual != "Pedidos de Venda":
        if tentar_alterar_valor(driver, campo_rotina, "Pedidos de Venda", log_queue, 'wa-panel[id="COMP3053"] > wa-text-input[id="COMP3056"]'):
            print("Valor alterado com sucesso.")
            log_queue.put("Valor alterado com sucesso.")
        else:
            print("Falha ao alterar valor")
            log_queue.put("Falha ao alterar valor")
    else:
        print("O valor já está correto, nenhuma alteração necessária.")
        log_queue.put("O valor já está correto, nenhuma alteração necessária.")

    print("Rotina inserida com sucesso.")
    log_queue.put("Rotina inserida com sucesso.")

    print("Buscando botão...")
    input_rotina = wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP3053"] > wa-text-input[id="COMP3056"]')
    btn_pesq = wait_for_element(driver, By.CSS_SELECTOR, 'button.button-image')
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_pesq)
    time.sleep(1)
    click_element(input_rotina, (By.CSS_SELECTOR, 'button.button-image'))

    shadow_button(driver, 'wa-menu-item[id="COMP4523"]', '.caption[title="Faturamento (1)"]', log_queue)
    shadow_button(driver, 'wa-menu-item[id="COMP4519"]', '.caption[title="Faturamento (1)"]', log_queue)
    shadow_button(driver, 'wa-menu-item[id= "COMP4520"]', '.caption[title="Pedidos de Venda"]', log_queue)

    print("Buscando segunda tela de validação...")
    log_queue.put("Buscando segunda tela de validação...")
    print("Abrindo wa-dialog do menu...")
    log_queue.put("Abrindo wa-dialog do menu...")

def apertar_incluir(driver, log_queue):
    """
    Função para apertar o botão de incluir em Pedidos de Venda
    """
    print("Buscando wa-panel da rotina de Pedidos...")
    log_queue.put("\nBuscando wa-panel da rotina de Pedidos...")
    wait_for_element(driver, By.ID, 'COMP4584')
    print("Tela carregada com sucesso.")
    log_queue.put("Tela carregada com sucesso.")
    time.sleep(5)

    print("Buscando botão de incluir...")
    log_queue.put("Buscando botão de incluir...")
    for i in range(0,5):    
        try:
            print(f"Tentativa: {i+1}")
            log_queue.put(f"Tentativa: {i+1}")
            btn_incluir = wait_for_element(driver, By.ID, 'COMP4586')

            print("Botão encontrado e expandindo shadow DOM...")
            log_queue.put("Botão encontrado e expandindo shadow DOM...")
            shadow_button(driver, 'wa-button[id="COMP4586"]', 'button', log_queue)
            time.sleep(2)
            
            if wait_for_element(driver, By.ID, 'COMP6000', timeout=10):
                print("Botão clicado com sucesso.")
                log_queue.put("Botão clicado com sucesso.")
                print("Abrindo tela de filiais...")
                log_queue.put("Abrindo tela de filiais...")
                break
            else:
                print(f"Erro na tentativa: {i+1}, tentando novamente...")
                log_queue.put(f"Erro na tentativa: {i+1}, tentando novamente...")
        except Exception as e:
            print(f"Erro: {e}")
            log_queue.put(f"Erro: {e}")

def abrir_pedido(driver, unidade, log_queue):
    """
    Função que adiciona a unidade novamente e confirma a abertura do pedido.
    """
    wa_dialog_filial = wait_for_element(driver, By.ID, 'COMP6000')
    print("Menu de filiais aberto.")
    log_queue.put("\nMenu de filiais aberto.")
    
    unidade_desejada = unidades[unidade-1]
    print("Aguardando input de filial...")
    log_queue.put("Aguardando input de filial...")
    wait_for_element(wa_dialog_filial, By.CSS_SELECTOR, filial_unidade)

    shadow_input(driver, filial_unidade, unidade_desejada, log_queue)
    print("Valor digitado.")

    wait_for_element(wa_dialog_filial, By.CSS_SELECTOR, btn_filial_unidade)
    print("Botão encontrado.")
    log_queue.put("Botão encontrado.")
    shadow_button(driver, btn_filial_unidade, 'button', log_queue)
    
    confirmando_wa_tgrid(driver, "COMP6012", 41, unidade, abrir_pedido, None, log_queue)

    print("Acessando painel...")
    log_queue.put("Acessando painel...")
    time.sleep(2)
    print("Procurando botão OK...")
    log_queue.put("Procurando botão OK...")
    wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP6001"] > wa-button[id="COMP6057"]')
    print("Botão encontrado.")
    log_queue.put("Botão encontrado.")

    shadow_button(driver, 'wa-panel[id="COMP6001"] > wa-button[id="COMP6057"]', 'button', log_queue)
    time.sleep(5)

def alterar_data(driver, data, log_queue):
    """
    Função para alterar a data do sistema.
    """
    close_button = 'wa-dialog[id="COMP4500"] > wa-panel[id="COMP4503"] > wa-button[id="COMP4514"]'
    seletor_data = 'wa-dialog[id="COMP4500"] > wa-panel[id="COMP4502"] > wa-text-input[id="COMP4507"]'
    confirma = 'wa-panel[id="COMP4504"] > wa-panel[id="COMP4520"] > wa-button[id="COMP4522"]'

    print("Fechando rotina atual...")
    log_queue.put("\nFechando rotina atual...")
    wait_for_element(driver, By.CSS_SELECTOR, close_button)
    print("Botão para fechar rotina encontrado.")
    log_queue.put("\nBotão para fechar rotina encontrado.")
    shadow_button(driver, close_button, 'button', log_queue)

    rotina_venda(driver, log_queue)

    print("Aberto.")
    log_queue.put("Aberto.")

    print("Acessando data do sistema...")
    log_queue.put("\nAcessando data do sistema...")
    campo_data = wait_for_element(driver, By.CSS_SELECTOR, seletor_data)
    print("Alterando data...")
    log_queue.put("\nAlterando data...")
    shadow_input(driver, seletor_data, data, log_queue)
    print("Data alterada.")
    log_queue.put("\nData alterada.")

    valor_atual = acessar_valor(campo_data)
    print(f"Valor atual do campo: {valor_atual}")
    log_queue.put(f"Valor atual do campo: {valor_atual}")

    if valor_atual != data:
        if tentar_alterar_valor(driver, campo_data, data, log_queue, seletor_data):
            print("Valor alterado com sucesso.")
            log_queue.put("Valor alterado com sucesso.")
        else:
            print("Falha ao alterar valor")
            log_queue.put("Falha ao alterar valor")
    else:
        print("O valor já está correto, nenhuma alteração necessária.")
        log_queue.put("O valor já está correto, nenhuma alteração necessária.")

    print("Aguardando botão de confirmar...")
    log_queue.put("\nAguardando botão de confirmar...")
    wait_for_element(driver, By.CSS_SELECTOR, confirma)
    shadow_button(driver, confirma, 'button', log_queue)

    time.sleep(7)

def busca_cnpj(driver, nota, log_queue):
    """
    Busca apenas o container do campo de CNPJ no sistema
    """
    # Caminho do arquivo JSON
    json_path = r"C:\Users\Pedro\Documents\BOT-SERVICES\path\cnpj.json"
    
    cnpj_input = nota.getCNPJ()
    cnpj_dict = carregar_dados(json_path)

    print("Buscando menu para pesquisar CNPJ...")
    log_queue.put("\nBuscando menu para pesquisar CNPJ...")
    wait_for_element(driver, By.CSS_SELECTOR, cnpj_selector)
    print("Encontrado.")
    log_queue.put("Encontrado.")

    codigo = cnpj_dict.get(cnpj_input, "NOT FOUND")
    print(f"Codigo Protheus(CNPJ): {codigo}")
    log_queue.put(f"Codigo Protheus(CNPJ): {codigo}")

    return codigo

def inserir_cnpj(driver, codigo, nota, log_queue):
    campo_codigo = wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP6009"] > wa-text-input[id="COMP6011"]')
    shadow_input(driver, 'wa-panel[id="COMP6009"] > wa-text-input[id="COMP6011"]', codigo, log_queue)

    valor_atual = acessar_valor(campo_codigo)
    print(f"Valor atual do campo: {valor_atual}")
    log_queue.put(f"Valor atual do campo: {valor_atual}")

    if valor_atual != codigo:  
        if tentar_alterar_valor(driver, campo_codigo, codigo, log_queue, 'wa-panel[id="COMP6009"] > wa-text-input[id="COMP6011"]'):
            print("Valor alterado com sucesso.")
            log_queue.put("Valor alterado com sucesso.")
            inserir_services(driver, nota, log_queue)
        else:
            print("Falha ao alterar valor")
            log_queue.put("Falha ao alterar valor")
    else:
        print("O valor já está correto, nenhuma alteração necessária.")
        log_queue.put("O valor já está correto, nenhuma alteração necessária.")
        inserir_services(driver, nota, log_queue)

def inserir_cnpj_pesquisa(driver, nota, log_queue):
    """
    Após buscado e acessado o botão de pesquisa, inserir cnpj.
    """
    # Caminho do arquivo JSON
    json_path = r"C:\Users\Pedro\Documents\BOT-SERVICES\path\cnpj.json"

    cnpj_dict = carregar_dados(json_path)

    container_cnpj = wait_for_element(driver, By.CSS_SELECTOR, cnpj_selector)
    btn_pesq = wait_for_element(driver, By.CSS_SELECTOR, 'button.button-image')
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_pesq)
    time.sleep(1) 
    click_element(container_cnpj, (By.CSS_SELECTOR, 'button.button-image'))
    print("Botão de pesquisa clicado.")
    log_queue.put("Botão de pesquisa clicado.")
    time.sleep(3)

    field_cnpj = wait_for_element(driver, By.CSS_SELECTOR, input_pesquisa)
    print("Field CNPJ encontrado.")
    log_queue.put("\nField CNPJ encontrado.")
    cnpj = nota.getCNPJ()
    print(f"Cliente: {cnpj}.")

    shadow_input(driver, input_pesquisa, cnpj, log_queue)
    print("CNPJ inserido com sucesso.")
    log_queue.put("CNPJ inserido com sucesso.")
    time.sleep(3)

    print("Pesquisando CNPJ...")
    log_queue.put("Pesquisando CNPJ...")
    pesquisar = wait_for_element(driver, By.CSS_SELECTOR, 'wa-button[id="COMP7534"]')
    shadow_pesquisar = expand_shadow_element(driver, pesquisar)
    button(driver, shadow_pesquisar, log_queue)
    time.sleep(3)
    
    confirmando_wa_tgrid(driver, "COMP7523", 15, nota, inserir_cnpj_pesquisa, "CNPJ", log_queue)

    confirmar_ok = wait_for_element(driver, By.CSS_SELECTOR, btn_ok_cnpj)
    shadow_ok = expand_shadow_element(driver, confirmar_ok)
    print("Acessando shadow-button...")
    log_queue.put("Acessando shadow-button...")
    button(driver, shadow_ok, log_queue)
    time.sleep(1.5)

    valor_atual = acessar_valor(container_cnpj)
    print(f"Valor atual do campo: {valor_atual}")
    log_queue.put(f"Valor atual do campo: {valor_atual}")

    novo_codigo = valor_atual
    
    adicionados = 0

    if cnpj not in cnpj_dict:
        cnpj_dict[cnpj] = novo_codigo
        adicionados += 1

    # Salvar apenas se houver novos CNPJs
    if adicionados > 0:
        salvar_dados(cnpj_dict, json_path)
        print(f"{adicionados} novos CNPJs foram adicionados ao arquivo.")
    else:
        print("Nenhum novo CNPJ foi adicionado.")

    inserir_services(driver, nota, log_queue)

def inserir_services(driver, nota, log_queue):
    """
    Acessa selection box do tipo de nota (S - SERVIÇO, M - MATERIAL, R - RETORNO)
    """
    print("Selecionando tipo de nota...")
    log_queue.put("\nSelecionando tipo de nota...")
    wa_combo_box = wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP6009"] > wa-combobox[id="COMP6013"]')
    shadow_combo_box = expand_shadow_element(driver, wa_combo_box)
    select_element = shadow_combo_box.find_element(By.CSS_SELECTOR, "select")
    # Escolha a segunda opção dentro do 'select'
    option = select_element.find_element(By.CSS_SELECTOR, "option:nth-child(2)")
    option.click()
    print("Serviço selecionado.")
    log_queue.put("Serviço selecionado.")
    time.sleep(1)

    busca_forma_pagto(driver, nota, log_queue)

def busca_forma_pagto(driver, nota, log_queue):
    """
    Busca apenas o container do campo de CNPJ no sistema
    """
    # Caminho do arquivo JSON
    json_path = r"C:\Users\Pedro\Documents\BOT-SERVICES\path\forma_pag.json"
    
    pagto_input = nota.getPAGTO()
    pagto_dict = carregar_dados(json_path)

    print("Buscando menu para pesquisar CNPJ...")
    log_queue.put("\nBuscando menu para pesquisar CNPJ...")
    wait_for_element(driver, By.CSS_SELECTOR, cnpj_selector)
    print("Encontrado.")
    log_queue.put("Encontrado.")

    codigo = pagto_dict.get(pagto_input, "NOT FOUND")
    print(f"Codigo Protheus(PAGTO): {codigo}")
    log_queue.put(f"Codigo Protheus(PAGTO): {codigo}")

    if codigo == "NOT FOUND":
        inserir_forma_pagto_pesquisa(driver, nota, log_queue)
    else:
        inserir_forma_pagto(driver, codigo, nota, log_queue)

def inserir_forma_pagto(driver, codigo, nota, log_queue):
    print("Informando forma de pagamento...")
    log_queue.put("\nInformando forma de pagamento...")
    usar_gatilho(driver, codigo, 'wa-panel[id="COMP6009"] > wa-text-input[id="COMP6014"]', inserir_iss, log_queue, nota)
    inserir_iss(driver, nota, log_queue)

def inserir_forma_pagto_pesquisa(driver, nota, log_queue):
    """
    Confirmar forma de pagamento acessando a lupa.
    """
    # Caminho do arquivo JSON
    json_path = r"C:\Users\Pedro\Documents\BOT-SERVICES\path\forma_pag.json"

    print("Informando forma de pagamento...")
    log_queue.put("\nInformando forma de pagamento...")
    container_pagto = wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP6009"] > wa-text-input[id="COMP6014"]')
    btn_pesq = wait_for_element(driver, By.CSS_SELECTOR, 'button.button-image')
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_pesq)
    time.sleep(1) 
    click_element(container_pagto, (By.CSS_SELECTOR, 'button.button-image'))
    print("Botão de pesquisa clicado.")
    log_queue.put("Botão de pesquisa clicado.")
    time.sleep(1)

    print("Procurando menu de pagamento...")
    log_queue.put("Procurando menu de pagamento...")
    pesquisa_pagto = wait_for_element(driver, By.CSS_SELECTOR, menu_pagto)
    print("Menu encontrado.")
    log_queue.put("Menu encontrado.")
    cond_pagto = nota.getPAGTO()
    shadow_input(driver, menu_pagto, cond_pagto, log_queue)
    time.sleep(1)

    button_pesq_pagto = wait_for_element(driver, By.CSS_SELECTOR, 'wa-button[id="COMP7534"]')
    shadow_pesq_pagto = expand_shadow_element(driver, button_pesq_pagto)
    button(driver, shadow_pesq_pagto, log_queue)
    time.sleep(1)

    confirmando_wa_tgrid(driver, "COMP7523", 29, nota, inserir_forma_pagto_pesquisa, "PAGTO", log_queue)
    
    confirmar_ok = wait_for_element(driver, By.CSS_SELECTOR, btn_ok_pagto_nat)
    shadow_ok = expand_shadow_element(driver, confirmar_ok)
    button(driver, shadow_ok, log_queue)

    print("Forma de pagamento adicionada.")
    log_queue.put("Forma de pagamento adicionada.")  
    time.sleep(1)

    valor_atual = acessar_valor(container_pagto)
    print(f"Valor atual do campo: {valor_atual}")
    log_queue.put(f"Valor atual do campo: {valor_atual}")

    novo_codigo = valor_atual
    
    adicionados = 0

    pagto_dict = carregar_dados(json_path)

    codigo = pagto_dict.get(valor_atual, "NOT FOUND")

    if codigo not in pagto_dict:
        pagto_dict[codigo] = novo_codigo
        adicionados += 1

    # Salvar apenas se houver novos CNPJs
    if adicionados > 0:
        salvar_dados(pagto_dict, json_path)
        print(f"{adicionados} novos CNPJs foram adicionados ao arquivo.")
    else:
        print("Nenhum novo CNPJ foi adicionado.")

    inserir_iss(driver, nota, log_queue)

def inserir_iss(driver, nota, log_queue):
    """
    Manipula a selection box para definir se tem tributação ou não
    """
    print("Selecionando tributação...")
    log_queue.put("\nSelecionando tributação...")
    wa_combo_box = wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP6009"] > wa-combobox[id="COMP6016"]')
    shadow_combo_box = expand_shadow_element(driver, wa_combo_box)
    select_element = shadow_combo_box.find_element(By.CSS_SELECTOR, "select")
    natureza = nota.getNAT()
    if natureza == "30102011" or natureza == "30102012":
        # Escolha a segunda opção dentro do 'select'
        option = select_element.find_element(By.CSS_SELECTOR, "option:nth-child(2)")
        option.click()
    elif natureza == "30102002" or natureza == "30102003":
        # Escolha a segunda opção dentro do 'select'
        option = select_element.find_element(By.CSS_SELECTOR, "option:nth-child(3)")
        option.click()

    print("Tributo selecionado.")
    log_queue.put("Tributo selecionado.")
    time.sleep(1)
    buscar_natureza(driver, nota, log_queue)

def buscar_natureza(driver, nota, log_queue):
    """
    Apenas busca o container do menu da natureza.
    """
    print("Abrindo menu de natureza do serviço...")
    log_queue.put("\nAbrindo menu de natureza do serviço...")
    
    codigo = nota.getNAT()

    print(f"Codigo Protheus(NATUREZA): {codigo}")
    log_queue.put(f"Codigo Protheus(NATUREZA): {codigo}")

    campo_codigo = wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP6009"] > wa-text-input[id="COMP6017"]')
    shadow_input(driver, 'wa-panel[id="COMP6009"] > wa-text-input[id="COMP6017"]', codigo, log_queue)

    campo_codigo.send_keys(Keys.RETURN)
    time.sleep(1)

    valor_atual = acessar_valor(campo_codigo)
    print(f"Valor atual do campo: {valor_atual}")
    log_queue.put(f"Valor atual do campo: {valor_atual}")

    if int(valor_atual) != int(codigo):
        print(f"\n{int(valor_atual)}")
        print(f"\n{int(codigo)}")
        if tentar_alterar_valor(driver, campo_codigo, codigo, log_queue, 'wa-panel[id="COMP6009"] > wa-text-input[id="COMP6017"]'):
            print("Valor alterado com sucesso.")
            log_queue.put("Valor alterado com sucesso.")
            abrir_vinculo_os(driver, nota, log_queue)
        else:
            print("Falha ao alterar valor")
            log_queue.put("Falha ao alterar valor")
    else:
        print("O valor já está correto, nenhuma alteração necessária.")
        log_queue.put("O valor já está correto, nenhuma alteração necessária.")

    time.sleep(1)

    abrir_vinculo_os(driver, nota, log_queue)

def abrir_vinculo_os(driver, nota, log_queue):
    """
    Abre menu do popup em outras ações para vincular OSs
    """
    time.sleep(1)
    print("Abrindo popup...")
    log_queue.put("\nAbrindo popup...")
    outras_button = wait_for_element(driver, By.CSS_SELECTOR, 'wa-button[id="COMP6171"]')
    shadow_outras = expand_shadow_element(driver, outras_button)
    button(driver, shadow_outras, log_queue)

    opcao_menu = wait_for_element(driver, By.CSS_SELECTOR, 'wa-menu-popup[id="COMP6170"] > wa-menu-popup-item[id="COMP6186"]')
    print("Popup aberto.")
    log_queue.put("Popup aberto.")
    print("Expandindo shadow do popup...")
    log_queue.put("Expandindo shadow do popup...")
    shadow_opcao = expand_shadow_element(driver, opcao_menu)
    print("Aberto e acessando...")
    log_queue.put("Aberto e acessando...")
    WebDriverWait(shadow_opcao, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'span.caption'))
    )
    rotina_element = shadow_opcao.find_element(By.CSS_SELECTOR, 'span.caption')
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", rotina_element)
    driver.execute_script("arguments[0].click();", rotina_element)
    print("Abrindo Vincular OSs...")
    time.sleep(1)
    vincular_os(driver, nota, log_queue)

def vincular_os(driver, nota, log_queue):
    """
    Abre o menu após acessar o botão de vínculo no popup e insere a OS correspondente ao serviço.
    """
    osKairos = nota.getOS()  # Isso agora pode ser uma lista de OS formatadas
    for os in osKairos:
        os_input = wait_for_element(driver, By.CSS_SELECTOR, 'wa-dialog[id="COMP7500"] > wa-text-input[id="COMP7502"]')
        shadow = expand_shadow_element(driver, os_input)
        inserir = wait_for_click(shadow, By.CSS_SELECTOR, 'input')
        time.sleep(1)
    
        print("Ativando foco no input.")
        log_queue.put("\nAtivando foco no input.")
        driver.execute_script("arguments[0].focus();", inserir)
        time.sleep(1)
    
        print("Usando Actions.")
        log_queue.put("Usando Actions.")
        ActionChains(driver).move_to_element(inserir).perform()
        print("Tecla BACKSPACE apertada.")
        log_queue.put("Tecla BACKSPACE apertada.")
        os_input.send_keys(Keys.BACKSPACE)  # Clear any existing text in the input field
        time.sleep(1)
        
        for char in os:
            os_input.send_keys(char)
            time.sleep(0.2)
    
        time.sleep(1)
    
        inserir.send_keys(Keys.RETURN)  # Press Enter to submit
        print("Primeiro enter pressionado.")
        log_queue.put("Primeiro enter pressionado.")
        time.sleep(1)
    
        confirmando_wa_tmsselbr(driver, "COMP7504", 35, nota, vincular_os, log_queue, os)

        enter_sec = wait_for_click(driver, By.CSS_SELECTOR, 'wa-tmsselbr[id="COMP7504"]')
        
        enter_sec.send_keys(Keys.RETURN)  # Press Enter to submit
        print("Segundo enter pressionado.")
        log_queue.put("Segundo enter pressionado.")
        time.sleep(1)
    
    print("Buscando botão de confirmar...")
    log_queue.put("Buscando botão de confirmar...")
    confirmar_os = wait_for_element(driver, By.CSS_SELECTOR, 'wa-dialog[id="COMP7500"] > wa-button[id="COMP7508"]')
    print("Encontrado.\nExpandindo shadow element...")
    log_queue.put("Encontrado.\nExpandindo shadow element...")
    shadow_os = expand_shadow_element(driver, confirmar_os)
    button(driver, shadow_os, log_queue)
    time.sleep(1)
    
    corpo_nota(driver, nota, log_queue)

def corpo_nota(driver, nota, log_queue):
    """
    Mesma lógica da função sobre o cabeçalho da nota mas agora define as informações do corpo.
    """
    selecionar_produto(driver, nota, log_queue)

def selecionar_produto(driver, nota, log_queue):
    """
    Busca acessar o campo do produto na tabela.
    """
    try:
        print("\nBuscando seleção de produto...")
        log_queue.put("\nBuscando seleção de produto...")
        # Esperar até que o elemento wa-dialog seja carregado
        dialog_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'wa-dialog[id="COMP6000"] > wa-tgetdados[id="COMP6160"]'))
        )

        # Obter o shadowRoot do elemento
        shadow_root = driver.execute_script("return arguments[0].shadowRoot", dialog_element)

        if shadow_root:
            target_element = shadow_root.find_element(By.CSS_SELECTOR, 'div.horizontal-scroll > table > tbody > tr#\\30 > td#\\31 > div')
            print(f"Elemento localizado: {target_element}")
            log_queue.put(f"Elemento localizado: {target_element}")
            target_element.click()
            time.sleep(1)
            print("Seleção concluída.")
            log_queue.put("Seleção concluída.")
            selecionar_elemento(driver, shadow_root, 'div.horizontal-scroll > table > tbody > tr#\\30 > td#\\31 > div', log_queue)
            inserir_produto(driver, log_queue, nota)
        else:
            print('ShadowRoot não encontrado no elemento.')
            log_queue.put("ShadowRoot não encontrado no elemento.")
    except Exception as e:
        print(f'Ocorreu um erro: {e}')
        log_queue.put(f'Ocorreu um erro: {e}')

def inserir_produto(driver, log_queue, nota, codigo="3500.0980"):
    """
    Após a busca ser concluída digita o código e a tecla enter é apertada para confirmar a escolha.
    """
    print('\nElemento encontrado e tecla "Enter" enviada com sucesso!')
    log_queue.put('Elemento encontrado e tecla "Enter" enviada com sucesso!')
    
    wa_dialog = wait_for_element(driver, By.CSS_SELECTOR, 'wa-dialog[id="COMP7500"] > wa-text-input[id="COMP7501"]')
    print("Inserindo código do produto...")
    log_queue.put("Inserindo código do produto...")
    time.sleep(1)

    inserir = shadow_input(driver,'wa-dialog[id="COMP7500"] > wa-text-input[id="COMP7501"]', codigo, log_queue)

    valor_atual = acessar_valor(wa_dialog).strip()  # Remove espaços extras
    print(f"Valor atual do campo: {valor_atual}")
    log_queue.put(f"Valor atual do campo: {valor_atual}")
    codigo = codigo.strip()  # Garante que também esteja sem espaços extras

    confirma_valor(driver, valor_atual, codigo, wa_dialog, log_queue, selecionar_quantidade, nota)
        
    inserir.send_keys(Keys.RETURN)  # Press Enter to submit
    print("Primeiro enter pressionado.")
    log_queue.put("Primeiro enter pressionado.")
    time.sleep(1)

    selecionar_quantidade(driver, log_queue, nota)

def selecionar_quantidade(driver, log_queue, nota):
    """
    Localiza o container da quantidade e insere.
    """
    seletor = 'div.horizontal-scroll > table > tbody > tr#\\30 > td#\\35 > div'
    element = 'wa-dialog[id="COMP6000"] > wa-tgetdados[id="COMP6160"]'
    print("\nBuscando campo da quantidade...")
    log_queue.put("\nBuscando campo da quantidade...")
    acessa_container(driver, element, seletor, inserir_quantidade, log_queue, nota)

def inserir_quantidade(driver, log_queue, nota, quant="1"):
    """
    Inserindo quantidade de serviços que por padrão é sempre 1
    """
    print('Elemento encontrado e tecla "Enter" enviada com sucesso!')
    log_queue.put('\nElemento encontrado e tecla "Enter" enviada com sucesso!')
    wa_dialog = wait_for_element(driver, By.CSS_SELECTOR, 'wa-dialog[id="COMP7500"] > wa-text-input[id="COMP7501"]')
    time.sleep(1)

    inserir = shadow_input_quant(driver, 'wa-dialog[id="COMP7500"] > wa-text-input[id="COMP7501"]', quant, log_queue)

    valor_atual = acessar_valor(wa_dialog)
    print(f"Valor atual do campo: {valor_atual}")
    log_queue.put(f"Valor atual do campo: {valor_atual}")

    valor_atual = str(float(valor_atual))  # Converte para float e depois de volta para string
    quant = str(float(quant))  # Faz o mesmo para 'quant'

    confirma_valor_quant(driver, valor_atual, quant, wa_dialog, log_queue, selecionar_preco, nota)

    inserir.send_keys(Keys.RETURN)  # Press Enter to submit
    print("Primeiro enter pressionado.")
    log_queue.put("Primeiro enter pressionado.")
    time.sleep(1)
        
    print("Serviço quantificado.")
    log_queue.put("Serviço quantificado.")
    time.sleep(1)
    selecionar_preco(driver, log_queue, nota)

def selecionar_preco(driver, log_queue, nota):
    """
    Localiza o container dos preços e acessa para poder inserir.
    """
    element = 'wa-dialog[id="COMP6000"] > wa-tgetdados[id="COMP6160"]'
    seletor = 'div.horizontal-scroll > table > tbody > tr#\\30 > td#\\36 > div'
    print("\nBuscando campo de preço...")
    acessa_container(driver, element, seletor, inserir_preco, log_queue, nota)

def inserir_preco(driver, log_queue, nota):
    """
    Após acessado o container, a função inserir o valor no formato do sistema.
    """
    preco = nota.getPRECO()
    log_queue.put(f"Preço: R$ {preco}")
    print(f"Preço: R$ {preco}")
    preco = f"{preco:.2f}"
    preco_format = str(preco).replace(".", ",") 
    log_queue.put(f"Preço formatado: R$ {preco_format}")
    print(f"Preço formatado: R$ {preco_format}")    
    print('Elemento encontrado e tecla "Enter" enviada com sucesso!')
    log_queue.put('Elemento encontrado e tecla "Enter" enviada com sucesso!')
    print("Inserindo preço...")
    log_queue.put("Inserindo preço...")
    wa_dialog = wait_for_element(driver, By.CSS_SELECTOR, 'wa-dialog[id="COMP7500"] > wa-text-input[id="COMP7501"]')
    time.sleep(1)

    inserir = shadow_input_quant(driver,'wa-dialog[id="COMP7500"] > wa-text-input[id="COMP7501"]', preco_format, log_queue)

    valor_atual = acessar_valor(wa_dialog)
    print(f"Valor atual do campo: {valor_atual}")
    log_queue.put(f"Valor atual do campo: {valor_atual}")

    # Converter 'valor_atual' para float, depois formatá-lo com duas casas decimais e vírgula
    valor_atual_formatado = f"{float(valor_atual):.2f}".replace(".", ",")

    print(f"Valor atual formatado: {valor_atual}")
    log_queue.put(f"Valor atual formatado: {valor_atual}")

    confirma_valor_quant(driver, valor_atual_formatado, preco_format, wa_dialog, log_queue, selecionar_tes)

    inserir.send_keys(Keys.RETURN)  # Press Enter to submit
    print("Preço digitado.")
    log_queue.put("Preço digitado.")
        
    time.sleep(1)
    selecionar_tes(driver, log_queue)

def selecionar_tes(driver, log_queue):
    """
    Mesma lógica, acessa o container para ficar disponível a inserção.
    """
    element = 'wa-dialog[id="COMP6000"] > wa-tgetdados[id="COMP6160"]'
    seletor = 'div.horizontal-scroll > table > tbody > tr#\\30 > td#\\31 0 > div'
    print("\nBuscando campo de TES...")
    log_queue.put("\nBuscando campo de TES...")
    acessa_container(driver, element, seletor, inserir_tes, log_queue)

def inserir_tes(driver, log_queue):
    try:
        print("Inserindo TES...")
        log_queue.put("\nInserindo TES...")

        wa_dialog = wait_for_element(driver, By.CSS_SELECTOR, 'wa-dialog[id="COMP7500"] > wa-text-input[id="COMP7501"]')
        time.sleep(1)

        shadow_input(driver, 'wa-dialog[id="COMP7500"] > wa-text-input[id="COMP7501"]', "568", log_queue)

        print("TES inserida com sucesso.")
        log_queue.put("TES inserida com sucesso.")
        time.sleep(1)

    except Exception as e:
        print(f"Erro ao inserir TES: {e}")
        log_queue.put(f"Erro ao inserir TES: {e}")
        raise  # Levanta a exceção para análise

def encerrar_pedido(driver, log_queue):
    """
    Função para encerrar o pedido de venda depois de completo.
    """
    try:
        print("Salvando a nota...")
        log_queue.put("\nSalvando a nota...")
        wa_button_save = wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP6164"] > wa-button[id="COMP6166"]')
        shadow_save = expand_shadow_element(driver, wa_button_save)
        button(driver, shadow_save, log_queue)
        print("Nota salva com sucesso.")
        log_queue.put("Nota salva com sucesso.")
        time.sleep(2.5)

        print("Acessando painel de cancelar...")
        # Verificar painel de retorno
        return_panel = wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP6164"]')
        if return_panel:
            print("Painel de retorno encontrado. Fechando...")
            log_queue.put("\nPainel de retorno encontrado. Fechando...")
            
            button_cancel = wait_for_element(return_panel, By.CSS_SELECTOR, 'wa-button[id="COMP6167"]')
            shadow_cancel = expand_shadow_element(driver, button_cancel)
            button(driver, shadow_cancel, log_queue)
            print("Nota encerrada.")
            log_queue.put("Nota encerrada.")
            time.sleep(1)
            return True
        else:
            print("Painel de retorno não encontrado.")
            log_queue.put("Painel de retorno não encontrado.")
            return False

    except Exception as e:
        print(f"Erro ao encerrar pedido: {e}")
        log_queue.put(f"Erro ao encerrar pedido: {e}")
        return False

def preparar_doc(driver, log_queue, num_nota):
    """
    Lógica para preparar o documento de saída e associar o número da nota.
    """
    print("Buscando botão para preparar Documento.")
    log_queue.put("\nBuscando botão para preparar Documento.")
    prep_doc = wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP4585"] > wa-button[id="COMP4587"]')
    shadow_doc = expand_shadow_element(driver, prep_doc)
    button(driver, shadow_doc, log_queue)

    print("Esperando menu de vínculo da nota...")
    log_queue.put("\nEsperando menu de vínculo da nota...")
    wait_for_element(driver, By.CSS_SELECTOR, 'wa-dialog[id="COMP6000"]')
    print("Menu aberto com sucesso.")
    log_queue.put("Menu aberto com sucesso.")
    
    time.sleep(1)
    clicar_repetidamente(driver, log_queue, 'wa-panel[id="COMP6004"] > wa-button[id="COMP6014"]', 'wa-panel[id="COMP6004"] > wa-button[id="COMP6015"]')

    time.sleep(1)
    
    print("Vinculando nota...")
    log_queue.put("\nVinculando nota...")
    
    clicar_elemento_shadow_dom(
        driver, "COMP6000", "COMP6004", 
        'div.horizontal-scroll > table > tbody > tr#\\33 > td#\\31 > div', 
        log_queue, num_nota)

    time.sleep(1)
    shadow_button(driver, 'wa-dialog[id="COMP6000"] > wa-button[id="COMP6005"]', 'button', log_queue)
    time.sleep(1)
    shadow_button(driver, 'wa-panel[id="COMP6066"] > wa-button[id="COMP6068"]', 'button', log_queue)

def renomeia_pdf(numero_nota, pasta_nfe, log_queue, db_nome, mes_ano, unidade):
    """
    Renomeia o arquivo após inserido no sistema e altera seu status no Banco de Dados.
    """
    log_queue.put("\nAjustando nome do arquivo...")
    print("Ajustando nome do arquivo...")

    # Verifica se a pasta existe
    if not os.path.exists(pasta_nfe):
        print(f"Erro: Pasta {pasta_nfe} não encontrada!")
        log_queue.put(f"Erro: Pasta {pasta_nfe} não encontrada!")
        return

    numero_nota_formatado = ajustar_numero_nota(numero_nota, unidade)

    print("Arquivos na pasta antes da renomeação:", os.listdir(pasta_nfe))

    for arquivo in os.listdir(pasta_nfe):
        print(f"Verificando: {arquivo} vs esperado: NFE {numero_nota_formatado}.pdf")
        log_queue.put(f"Verificando: {arquivo} vs esperado: NFE {numero_nota_formatado}.pdf")
        if arquivo.startswith(f"NFE {numero_nota_formatado}") and arquivo.endswith('.pdf'):
            caminho_antigo = os.path.join(pasta_nfe, arquivo)
            caminho_novo = os.path.join(pasta_nfe, arquivo.replace('.pdf', ' X.pdf'))
            
            print(f"Renomeando {caminho_antigo} para {caminho_novo}")
            try:
                os.rename(caminho_antigo, caminho_novo)
                log_queue.put(f"Arquivo renomeado: {caminho_antigo} -> {caminho_novo}")
                print(f"Arquivo renomeado com sucesso: {caminho_antigo} -> {caminho_novo}")
            except Exception as e:
                print(f"Erro ao renomear: {e}")
                log_queue.put(f"Erro ao renomear: {e}")
            break
    else:
        print(f"Nenhum arquivo correspondente encontrado para NFE {numero_nota_formatado}")

    # Atualizar status no banco de dados
    conn = sqlite3.connect(f'notas_{db_nome}.db')
    cursor = conn.cursor()
    tabela = f"{mes_ano.lower()}"

    cursor.execute(f'''
        UPDATE "{tabela}" 
        SET status_nfe = "Inserido" 
        WHERE numero_nota = ?
        ''', (numero_nota,))
    
    conn.commit()
    conn.close()

    print(f"Status da nota {numero_nota} atualizado para 'Inserido'.")
    log_queue.put(f"\nStatus da nota {numero_nota} atualizado para 'Inserido'.")

def formatar_os_kairos(os_kairos, log_queue):
    """
    Formata a osKairos para garantir que a primeira parte antes do hífen tenha 4 caracteres e esteja no formato correto.
    Exemplo: 0288-S11D, 17-TGM -> 0017-TGM, 0045-ALU.
    """
    if os_kairos:
        partes = os_kairos.split()
        
        os_encontradas = []
        
        for parte in partes:
            # Tentativa de capturar o padrão de número seguido de hífen e letras (ex: 0288-S11D, 17-TGM)
            match = re.match(r'(\d+[-_]\w+)(?:[\s,;]*)?', parte)
            if match:
                os_formatada = match.group(1)
                # Garantir que a primeira parte tenha 4 caracteres
                prefixo, sufixo = os_formatada.split('-')
                prefixo = prefixo.zfill(4)  # Preenche com zeros à esquerda, se necessário
                os_encontradas.append(f"{prefixo}-{sufixo}")
            else:
                # Caso não consiga capturar uma OS formatada, verifica se há uma sequência de números e um sufixo comum
                match_multiple = re.match(r'(\d+)(?:[-_]?)(\w+)?', parte)
                if match_multiple:
                    numero = match_multiple.group(1).zfill(4)  # Preenche com zeros à esquerda
                    sufixo = match_multiple.group(2) if match_multiple.group(2) else ""
                    if sufixo:
                        os_encontradas.append(f"{numero}-{sufixo}")
                    else:
                        os_encontradas.append(numero)
        
# Coloca os resultados na fila de log para auditoria
        if os_encontradas:
            for os_item in os_encontradas:
                log_queue.put(os_item)

            return os_encontradas
    return os_kairos    

def carregar_notas(db_nome, mes_ano):
    """ Carrega as notas do banco de dados SQLite. """
    with sqlite3.connect(f'notas_{db_nome}.db') as conn:
        cursor = conn.cursor()
        cursor.execute(f"""
            SELECT tomador_cnpj, tipo_pagamento, natureza, ordem_servico, 
                   valor_total, numero_nota, data_emissao 
            FROM {mes_ano.lower()} 
            WHERE status_nfe = 'Encontrado'
        """)
        return cursor.fetchall()

def inicializar_sistema(driver, unidade, data_inicial, log_queue):
    """ Realiza login e inicializações no sistema. """
    process_shadow_dom(driver, log_queue)
    locate_and_access_iframe(driver, log_queue)
    perform_login(driver, "user", "password", log_queue)
    abrir_menu_unidade(driver, unidade, data_inicial, log_queue)

def processar_notas(driver, notas, data_anterior, unidade, log_queue, mes_ano, db_nome, mes_selecionado):
    """ Processa todas as notas recuperadas do banco de dados. """

    for nota in notas:
        print(f"Iniciando processamento da nota {nota.getNumNOTA()}.")
        log_queue.put(f"\nIniciando processamento da nota {nota.getNumNOTA()}.")
        
        if nota.getDATA() != data_anterior:
            print("Mudança na data encontrada.")
            log_queue.put("\nMudança na data encontrada.")
            time.sleep(1)
            alterar_data(driver, nota.getDATA(), log_queue)
        
        apertar_incluir(driver, log_queue)
        abrir_pedido(driver, unidade, log_queue)

        time.sleep(1)

        codigo = busca_cnpj(driver, nota, log_queue)
        
        if codigo == "NOT FOUND":
            inserir_cnpj_pesquisa(driver, nota, log_queue)
        else:
            inserir_cnpj(driver, codigo, nota, log_queue)
        
        sucesso = encerrar_pedido(driver, log_queue)

        if sucesso:
            status = verificar_situacao(driver, log_queue)
            if status in ["Em Aberto", "Liberado"]:
                preparar_doc(driver, log_queue, nota.getNumNOTA())
                caminho_nfe = definir_nfe(unidade, mes_ano.split("_")[1].strip(), mes_selecionado)
                renomeia_pdf(nota.getNumNOTA(), caminho_nfe, log_queue, db_nome, mes_ano, unidade)
                data_anterior = nota.getDATA()
                time.sleep(3)
            else:
                log_queue.put("\nNota não foi salva corretamente. Tente de novo...")

def main_process(driver, url, db_nome, unidade, mes_ano, log_queue, mes_selecionado):
    """
    Gerencia o fluxo principal do processo para múltiplas notas, agora com integração ao banco de dados SQLite.
    O mês e ano da tabela são passados como parâmetro, e a data em abrir_menu_unidade é baseada na primeira nota com status "Encontrado".
    """
    global connection_successful, monitoring

    stop_monitoring = threading.Event()
    monitor_thread = monitor_connection_thread(driver, url, log_queue, stop_monitoring)

    try:
        log_queue.put("Iniciando o código principal...")
        print("Iniciando o código principal...")

        # Aguardar conexão
        while not connection_successful:
            log_queue.put("Aguardando conexão...")
            print("Aguardando conexão...")
            time.sleep(1)

        if connection_successful:
            log_queue.put("\nConexão estabelecida. Iniciando processamento!")
            print("Conexão estabelecida. Iniciando processamento!")

            # Buscar notas com status "Encontrado"
            notas_db = carregar_notas(db_nome, mes_ano)

            # Lista de notas a processar
            notas = []
            for nota_db in notas_db:
                cnpj, cond_pagto, natureza, os_kairos, preco, numero_nota, data_emissao = nota_db

                # Ignorar condições de pagamento que não existem no dicionário
                if cond_pagto is None:
                    log_queue.put(f"Ignorando nota {numero_nota} devido à condição de pagamento inválida: {cond_pagto}")
                    print(f"Ignorando nota {numero_nota} devido à condição de pagamento inválida: {cond_pagto}")
                    continue

                os_kairos_formatada = formatar_os_kairos(os_kairos, log_queue)
                servico = criar_servico(cnpj, cond_pagto, natureza, os_kairos_formatada, preco, numero_nota, data_emissao)

                notas.append(servico)

            # Definir a data inicial baseada na primeira nota
            data_inicial = notas[0].getDATA()
            log_queue.put(f"Data inicial baseada na primeira nota: {data_inicial}")
            print(f"Data inicial baseada na primeira nota: {data_inicial}")

            # Inicializar sistema
            inicializar_sistema(driver, unidade, data_inicial, log_queue)

            # Executar fluxo principal da nota
            rotina_venda(driver, log_queue)
            
            shadow_button(
            driver, 
            'wa-dialog[id="COMP4500"] > wa-panel[id="COMP4503"] > wa-panel[id="COMP4504"] > wa-panel[id="COMP4520"] > wa-button[id="COMP4522"]', 
            'button', 
            log_queue)

            print("Aberto.")
            log_queue.put("Aberto.")

            time.sleep(3)

            # Processar notas
            processar_notas(driver, notas, data_inicial, unidade, log_queue, mes_ano, db_nome, mes_selecionado)

            print("Processamento de todas as notas concluído.")
            log_queue.put("\nProcessamento de todas as notas concluído.")

            return True
        else:
            print("Conexão não estabelecida. Verifique a lógica de monitoramento.")
            log_queue.put("\nConexão não estabelecida. Verifique a lógica de monitoramento.")

    except (NoSuchElementException, ElementNotInteractableException, TimeoutException, JavascriptException, WebDriverException) as e:
        msg = f"Erro Selenium: {e}"
        log_queue.put(msg)
        print(msg)
        print(traceback.format_exc())

        return False

    except Exception as e:
        msg = f"Erro no processo principal: {e}"
        log_queue.put(msg)
        print(msg)
        print(traceback.format_exc())

        return False

    finally:
        stop_monitoring.set()
        monitor_thread.join()
        print("Finalizando driver e monitoramento.")
        log_queue.put("Finalizando driver e monitoramento.")