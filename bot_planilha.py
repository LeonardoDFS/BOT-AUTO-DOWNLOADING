import os
import time
import chromedriver_autoinstaller  # Biblioteca para instalar o ChromeDriver automaticamente
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# Instala o ChromeDriver automaticamente
chromedriver_autoinstaller.install()

# Caminho do diretório para download
download_dir = r"C:\Users\Leonardo\Desktop\planilhas teste"

# Função para definir as opções do Chrome
def get_chrome_options(download_dir):
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument("--headless")  # Descomente esta linha para rodar o Chrome em modo headless (sem interface gráfica)
    chrome_options.add_argument("--disable-gpu")  # Desabilita a GPU
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,  # Muda o diretório de download
        "download.prompt_for_download": False,        # Baixa automaticamente sem pedir confirmação
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True                  # Habilita o download de arquivos potencialmente "perigosos"
    })
    return chrome_options

# Função para apagar todos os arquivos em um diretório
def clear_download_directory(download_dir):
    try:
        # Lista todos os arquivos no diretório
        files = os.listdir(download_dir)
        for file in files:
            file_path = os.path.join(download_dir, file)
            # Verifica se é um arquivo (não deletar diretórios por engano)
            if os.path.isfile(file_path):
                os.remove(file_path)  # Remove o arquivo
        print(f"Todos os arquivos em {download_dir} foram apagados.")
    except Exception as e:
        print(f"Erro ao apagar arquivos: {e}")

# Função para verificar se o usuário está logado
def is_logged_in(driver):
    try:
        driver.find_element(By.ID, "logout")  # Exemplo de um botão de logout
        return True
    except NoSuchElementException:
        return False

# Função para realizar o login
def login(driver, username, password):
    try:
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "email")))
        email_field = driver.find_element(By.ID, "email")
        email_field.clear()
        email_field.send_keys(username)
        time.sleep(1)

        password_field = driver.find_element(By.ID, "password")
        password_field.clear()
        password_field.send_keys(password)
        
        login_button = driver.find_element(By.ID, "sign-in")
        login_button.click()
        
        time.sleep(1)
        WebDriverWait(driver, 20).until(is_logged_in(driver))

    except TimeoutException:
        print("O tempo de espera para o login expirou.")
    except Exception as e:
        print(f"Ocorreu um erro durante o login: {e}")

# Função para renomear o arquivo baixado
def rename_downloaded_file(download_dir, new_name):
    time.sleep(5)  # Espera o download ser concluído
    files = os.listdir(download_dir)
    # Ordena os arquivos por data de criação, do mais recente para o mais antigo
    files = sorted(files, key=lambda x: os.path.getctime(os.path.join(download_dir, x)), reverse=True)
    for file in files:
        if file.endswith(".xls") or file.endswith(".xlsx"):  # Verifica se é um arquivo Excel
            old_file = os.path.join(download_dir, file)
            new_file = os.path.join(download_dir, new_name)
            try:
                os.rename(old_file, new_file)
                print(f"Arquivo renomeado para: {new_name}")
            except Exception as e:
                print(f"Erro ao renomear o arquivo: {e}")
            break

#Função para tirar a interface gráfica
def get_chrome_options(download_dir):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")  # Rodar o Chrome em modo headless (sem interface gráfica)
    chrome_options.add_argument("--disable-gpu")  # Desabilita a GPU
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,  # Muda o diretório de download
        "download.prompt_for_download": False,        # Baixa automaticamente sem pedir confirmação
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True                  # Habilita o download de arquivos potencialmente "perigosos"
    })
    return chrome_options

# Função para realizar o download das planilhas
def download_planilha_TELB5(driver):
    try:
        #ABA DOCUMENTOS
        #<<--------------------------------1° Planilha------------------------------------------>>
        # Clicando no menu de Documentos
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "documentsMenu"))).click()
        time.sleep(1)
        
        # Clicando na lista de documentos
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "documentsMenuList"))).click()
        time.sleep(1)
        
        # Clicando em Todos os Projetos
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "selectProject"))).click()
        time.sleep(1)
                   
        # Clicando no LOTE 5
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "1392"))).click()
        time.sleep(1)
        
        # Dentro do LOTE 5, clicando em exportar
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "exportButton"))).click()
        time.sleep(1)
        
        # Clicando em Baixar todos
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "exportAll"))).click()
        time.sleep(3)
        
        # Renomeia o arquivo baixado
        rename_downloaded_file(download_dir, "Planilha_LT05.xlsx")

        #<<--------------------------------2° Planilha------------------------------------------>>
        # 2° Etapa: Download em Documentos do LT05, Logo após irá clicar em todos os projetos, 
        # escolhendo a opção LT05 e clicando em exportar, e por último em todos
        
        # Clicando em Todos os Projetos
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "selectProject"))).click()
        time.sleep(2)
                   
        # Clicando no TELB5-LT05
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "1528"))).click()
        time.sleep(3)
        
        # Dentro do TELB5-LT05, clicando em exportar
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "exportButton"))).click()
        time.sleep(2)
        
        # Clicando em Baixar todos
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "exportAll"))).click()
        time.sleep(4)

        # Renomeia o arquivo baixado
        rename_downloaded_file(download_dir, "CONSAG.xlsx")

    #ABA HISTÓRICO
    #<<--------------------------------3° Planilha------------------------------------------>>
        # Clicando no menu de Documentos
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "historyMenuLink"))).click()
        time.sleep(1)
        
        # Clicando na lista de documentos
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "transitionhistoryMenuList"))).click()
        time.sleep(1)
        
        # Clicando em Todos os Projetos
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "selectProject"))).click()
        time.sleep(1)
        
        # Clicando no LOTE 05
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "1392"))).click()
        time.sleep(1)
        
        #Clicando em Todas as Versões
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "versionFilter"))).click()
        time.sleep(1)
        
        #Clicando em Última Versão
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "lastVersion"))).click()
        time.sleep(1)
        
        #Clicando em Todas as Versões
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "versionFilter"))).click()
        time.sleep(1)
                           
        # Dentro do LOTE 5, clicando em exportar
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "exportButton"))).click()
        time.sleep(1)
        
        # Clicando em Relatório Horizontal
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "handleToggleHorizontal"))).click()
        time.sleep(1)
        
        #Clicando em todos no relatório horizontal
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "horizontalExportAll"))).click()
        time.sleep(3)
        
        # Renomeia o arquivo baixado
        rename_downloaded_file(download_dir, "Planilha_LT05_Tramitacoes.xlsx")
    
    #<<--------------------------------4° Planilha------------------------------------------>>
        # 4° Etapa: Download em Histórico do TELB5-LT05, Logo após irá clicar em todos os projetos, 
        # escolhendo a opção LT05 e clicando em exportar, e por último em todos
        
        # Clicando em Todos os Projetos
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "selectProject"))).click()
        time.sleep(2)
                   
        # Clicando no TELB5-LT05
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "1528"))).click()
        time.sleep(3)
        
        #Clicando em Todas as Versões
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "versionFilter"))).click()
        time.sleep(1)
        
        #Clicando em Última Versão
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "lastVersion"))).click()
        time.sleep(1)
        
        #Clicando em Todas as Versões
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "versionFilter"))).click()
        time.sleep(1)
                                 
        # Dentro do TELB5-LT05, clicando em exportar
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "exportButton"))).click()
        time.sleep(2)
        
        #Clicando em todos no relatório horizontal
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "horizontalExportAll"))).click()
        time.sleep(3)
        # Renomeia o arquivo baixado
        rename_downloaded_file(download_dir, "CONSAG_Tramitacoes.xlsx")
        
    except TimeoutException:
        print("O tempo de espera para o download expirou.")
    except NoSuchElementException:
        print("Não foi possível encontrar um dos elementos necessários para o download.")
    except Exception as e:
        print(f"Ocorreu um erro durante o processo de download: {e}")

# Configuração do WebDriver (navegador Chrome neste exemplo)
chrome_options = get_chrome_options(download_dir)
driver = webdriver.Chrome(options=chrome_options)  # Usa o ChromeDriver instalado automaticamente

try:
    clear_download_directory(download_dir)
    
    chrome_options = get_chrome_options(download_dir)
    driver = webdriver.Chrome(options=chrome_options)

    driver.get("https://app.keepcontrol.com.br/#/login")

    # Faça o login
    login(driver, "email@teste.com", "senha123456789")

    # Continue com outras ações após o login
    download_planilha_TELB5(driver)

finally:
    # Encerre o WebDriver
    driver.quit()
