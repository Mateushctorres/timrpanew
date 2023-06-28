import time
import csv
import configparser
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException


def gravar():
    tabela.to_csv(saida)  

def ler_arquivo(arquivo):
    cpf=[]
    
    with open(arquivo, 'r', encoding='utf-8') as f: 

        csv_reader = csv.DictReader(f, delimiter=',')
    
        for row in csv_reader:
            cpf.append(row)
        
    return cpf

def gravar_arquivo(arquivo, tabela):
    keys = tabela[0].keys()
    keys = list(keys)# + ['A/R']  # Adicione o cabeçalho da coluna 'A/R'
    with open(arquivo, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        writer.writerows(tabela)

#Carregando o arquivo config.ini
config = configparser.ConfigParser()
config.read('config.ini')

#Lendo o config.ini e salvando variaveis
caminho = '' 
GRAVAR_A_CADA = int(config["bot"]["gravar_a_cada"])
caminho = config["bot"]["caminho_webdriver"]
cep = config["bot"]["cep"]
numero = config["bot"]["numero"]
arquivo = config["bot"]["arquivo"]
saida = config["bot"]["arquivo_saida"]

print('# Carregando tabela ' + arquivo)
tabela = ler_arquivo(arquivo)
tabela_saida = []

#Carregando site
service = Service(caminho)

driver = webdriver.Chrome(service=service)
driver.get('https://apptimvendas.timbrasil.com.br/#/login')
driver.maximize_window()


#Setando timeout 
wait = WebDriverWait(driver, 30)
wait_inic = WebDriverWait(driver, 50) 

#Fazendo Login Pg Inicial
print('# Aguardando login.... 30 segs')

wait_inic.until(EC.visibility_of_element_located((By.XPATH, '/html/body/ion-app/ng-component/ion-nav/page-installation-address/ion-content/div[2]/ion-card/ion-card-content/ion-list/ion-item[1]/div[1]/div/ion-input/input')))

#Preenchendo campo CEP
driver.find_element("xpath", '/html/body/ion-app/ng-component/ion-nav/page-installation-address/ion-content/div[2]/ion-card/ion-card-content/ion-list/ion-item[1]/div[1]/div/ion-input/input').click()     
driver.find_element("xpath", '/html/body/ion-app/ng-component/ion-nav/page-installation-address/ion-content/div[2]/ion-card/ion-card-content/ion-list/ion-item[1]/div[1]/div/ion-input/input').send_keys(cep)    

#Preenchendo campo Número
driver.find_element("xpath", '/html/body/ion-app/ng-component/ion-nav/page-installation-address/ion-content/div[2]/ion-card/ion-card-content/ion-list/ion-item[2]/div[1]/div[1]/ion-input/input').click()     
driver.find_element("xpath", '/html/body/ion-app/ng-component/ion-nav/page-installation-address/ion-content/div[2]/ion-card/ion-card-content/ion-list/ion-item[2]/div[1]/div[1]/ion-input/input').send_keys(numero)  

#Click em Pesquisar
driver.find_element("xpath", '/html/body/ion-app/ng-component/ion-nav/page-installation-address/ion-content/div[2]/ion-card/ion-card-content/ion-list/ion-item[1]/div[1]/button/span').click()    
wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/ion-app/ng-component/ion-nav/page-installation-address/ion-content/div[2]/ion-card/ion-card-header/ion-col[2]')))

#time.sleep(3)

#Marca campo Permitir edição manual
driver.find_element("xpath", '//*[@id="checkbox02"]').click()     

#Click em Próximo
time.sleep(3)
driver.find_element("xpath", '//*[@id="button-back"]').click()     

#Clica em Sim
wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/ion-app/ion-alert/div')))
driver.find_element("xpath", '/html/body/ion-app/ion-alert/div/div[3]/button[2]').click()     

# Make a loop to iterate through all the elements in the list
INTERVALO = 48  # Intervalo desejado

for i in range(len(tabela)):
    print('# Iniciando processo ' + str(i+1) + ' de ' + str(len(tabela)))

    if (i+1) % INTERVALO == 0:
        print('CPF igual a 48. Pulando para o próximo intervalo.')
        continue
    #Pegar CPF no arquivo .csv e consultar
    campocpf = '/html/body/ion-app/ng-component/ion-nav/page-client-identification/ion-content/div[2]/ion-card/ion-card-content/ion-item/div[1]/div/ion-input/input'
    wait.until(EC.element_to_be_clickable((By.XPATH, campocpf)))
    time.sleep(3)
    driver.find_element("xpath", campocpf).click()
    driver.find_element("xpath", campocpf).send_keys(Keys.CONTROL + "a")
    driver.find_element("xpath", campocpf).send_keys(Keys.DELETE)
    driver.find_element("xpath", campocpf).send_keys(tabela[i]['CPF'])

    #Próxima tela
    wait.until(EC.element_to_be_clickable((By.XPATH,  '/html/body/ion-app/ng-component/ion-nav/page-client-identification/ion-content/div[2]/button[2]/span')))
    driver.find_element("xpath", '/html/body/ion-app/ng-component/ion-nav/page-client-identification/ion-content/div[2]/button[2]/span').click()
    #wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/ion-app/ng-component/ion-nav/page-base-customer-data-pf/ion-content/div[2]/ion-card[1]/ion-card-header')))

    #Preenchendo campo Nome
    camponome = '/html/body/ion-app/ng-component/ion-nav/page-base-customer-data-pf/ion-content/div[2]/ion-card[1]/ion-card-content/ion-list/ion-item[2]/div[1]/div/ion-input/input'
    wait.until(EC.element_to_be_clickable((By.XPATH, camponome)))
    driver.find_element("xpath", camponome).click()
    driver.find_element("xpath", camponome).send_keys(Keys.CONTROL + "a")
    driver.find_element("xpath", camponome).send_keys(Keys.DELETE)
    driver.find_element("xpath", camponome).send_keys(tabela[i]['NOME'])

    #Preenchendo campo Data de Nascimento
    campodn = '/html/body/ion-app/ng-component/ion-nav/page-base-customer-data-pf/ion-content/div[2]/ion-card[1]/ion-card-content/ion-list/ion-item[3]/div[1]/div/ion-input/input'
    wait.until(EC.element_to_be_clickable((By.XPATH, campodn)))
    driver.find_element("xpath", campodn).click()
    driver.find_element("xpath", campodn).send_keys(Keys.CONTROL + "a")
    driver.find_element("xpath", campodn).send_keys(Keys.DELETE)
    driver.find_element("xpath", campodn).send_keys(tabela[i]['DN'])

    #Preenchendo campo Nome da Mãe
    campomae = '/html/body/ion-app/ng-component/ion-nav/page-base-customer-data-pf/ion-content/div[2]/ion-card[1]/ion-card-content/ion-list/ion-item[4]/div[1]/div/ion-input/input'
    wait.until(EC.element_to_be_clickable((By.XPATH, campomae)))
    driver.find_element("xpath", campomae).click()
    driver.find_element("xpath", campomae).send_keys(Keys.CONTROL + "a")
    driver.find_element("xpath", campomae).send_keys(Keys.DELETE)
    driver.find_element("xpath", campomae).send_keys(tabela[i]['MAE'])

    #Click em Próximo
    wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="button-next"]/span')))
    driver.find_element("xpath", '//*[@id="button-next"]/span').click()

    # Aprovado ou Reprovado
    resultado = ""
    try:
        time.sleep(15)
        driver.find_element(By.XPATH, '/html/body/ion-app/ion-alert/div/div[3]/button[1]/span').click()
        print('** Reprovado')
        resultado = 'Reprovado'
    except NoSuchElementException:
        wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/ion-app/ng-component/ion-nav/page-offers-available/ion-content/div[2]/ion-card[2]/ion-card-header')))
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="button-back-plans"]/span').click()
        time.sleep(1)
        print('** Aprovado')
        resultado = 'Aprovado'
    
    tabela[i]['A/R'] = resultado

    print('** novo loop')

    driver.get('https://apptimvendas.timbrasil.com.br/#/client-identification')

    gravar_arquivo(saida, tabela)
