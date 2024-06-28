from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
from openpyxl import Workbook
from time import sleep
import sys
from pathlib import Path

# Verificar se estamos rodando como um executável PyInstaller
if getattr(sys, 'frozen', False):
    # Definir o caminho para o arquivo Excel dentro do diretório temporário gerado pelo PyInstaller
    arquivo_excel = Path(sys._MEIPASS) / 'Inserir_Push.xlsx'
else:
    # Caso contrário, assumir que estamos rodando o script diretamente
    arquivo_excel = Path(__file__).parent / 'Inserir_Push.xlsx'

# Carregar os dados da planilha Excel
try:
    workbook = openpyxl.load_workbook(arquivo_excel)
    ler_planilha = workbook['Plan1']
except FileNotFoundError:
    print(f"Arquivo '{arquivo_excel}' nao encontrado. Verifique o caminho do arquivo.")
    sys.exit(1)  # Sair do script se o arquivo não for encontrado

# Criar uma nova planilha para armazenar os resultados
nova_planilha = Workbook()
planilha_sheet = nova_planilha.active
planilha_sheet.title = "Resultados"
planilha_sheet.append(["Processo", "Status"])

# Configurações do navegador
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')  # Iniciar maximizado
options.add_argument('--disable-infobars')  # Desabilitar infobars
options.add_argument('--disable-extensions')  # Desabilitar extensões
options.page_load_strategy = 'normal'  # Vai carregar a página normalmente
options.add_experimental_option("detach", True)  # Após executar o script Selenium, o navegador continuará aberto

# Inicializa o serviço do ChromeDriver automaticamente usando o webdriver_manager
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(options=options, service=service)

# Abrindo PJE e fazendo o login
driver.get("https://pje1g.trf3.jus.br/pje/login.seam")  # URL do PJE

# A página de login tem um iframe que barra a automação, esse trecho burla esse iframe
frame_correta = driver.find_element(By.XPATH, '//*[@id="ssoFrame"]')
driver.switch_to.frame(frame_correta)

# Identificando os IDs e inserindo o login
login = driver.find_element(By.ID, 'username').send_keys('52312819805')  # Alterar CPF aqui
senha = driver.find_element(By.ID, 'password').send_keys('Eg@5044326')  # Alterar senha aqui
sleep(2)
clic = driver.find_element(By.ID, 'kc-login').click()

# Esperar até que o redirecionamento ocorra e a nova página seja carregada
try:
    # Esperar até que um elemento específico da nova página esteja presente
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, 'elemento_xpath_da_nova_pagina')))
    print("Login bem-sucedido e redirecionado para a nova pagina.")
except Exception as e:
    print(f"Ocorreu um erro durante o login ou o redirecionamento: {e}")

# Redirecionar para uma URL específica após o login
nova_url = "https://pje1g.trf3.jus.br/pje/Push/listView.seam"
driver.get(nova_url)
print(f"Redirecionado para {nova_url}")

# Lendo os dados da planilha
for linha in ler_planilha.iter_rows(min_row=1):
    processo = linha[0].value
    if processo:
        seletor_processo = driver.find_element(By.XPATH, '//*[@id="j_id148:inputNumeroProcesso-inputNumeroProcessoDecoration:inputNumeroProcesso-inputNumeroProcesso"]')
        seletor_processo.send_keys(processo)
        seletor_processo.send_keys(Keys.RETURN)

        wait = WebDriverWait(driver, 5)

        try:
            # Verificar se o processo já está no site
            resultado = wait.until(EC.presence_of_element_located((By.XPATH, f'//[contains(text(), "{processo}")]')))
            sleep(10)
            print(f'Processo {processo} já esta no site.')
            planilha_sheet.append([processo, "Já está no site"])
        except:
            print(f'Processo {processo} nao encontrado no site.')
            planilha_sheet.append([processo, "Nao encontrado no site"])

            # Clique no botão para incluir o processo
            try:
                botao_incluir = driver.find_element(By.XPATH, '//*[@id="j_id148:btnIncluirAcompanhamento"]')  # Altere para o ID do botão de incluir
                botao_incluir.click()

                # Preencher o campo do processo e submeter
                input_processo = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="j_id148:inputNumeroProcesso-inputNumeroProcessoDecoration:inputNumeroProcesso-inputNumeroProcesso"]')))  # Altere para o ID do campo de input do processo
                input_processo.send_keys(processo)

                print(f'Processo {processo} foi incluido no site.')
                planilha_sheet.append([processo, "Incluído no site"])
            except Exception as e:
                print(f'Erro ao incluir o processo {processo}: {e}')
                planilha_sheet.append([processo, f'Erro ao incluir: {e}'])

# Fechar o navegador e salvar a nova planilha
try:
    nova_planilha.save('Resultados_Inclusao_Processos.xlsx')
    print("Planilha de resultados salva com sucesso.")
except Exception as e:
    print(f"Erro ao salvar a planilha de resultados: {e}")

driver.quit()