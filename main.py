import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import pandas as pd
from selenium.webdriver.common.keys import Keys
import re
import unicodedata

options = uc.ChromeOptions()

driver = uc.Chrome(options = options, version_main=138)

locator = {
    # Login
    'usuario': (By.CSS_SELECTOR, '#name'),
    'senha': (By.CSS_SELECTOR, '#password'),
    'entrar': (By.CSS_SELECTOR, '#enter'),

    # Filtro
    'grupos_elementos_ul': (By.XPATH, '/html/body/div/main/div[1]/form/div/div[1]/div/div[1]/ul/li[1]/div[2]/div/div[1]/div[2]/ul'),
    'grupos_host_ul': (By.XPATH, '/html/body/div/main/form/div[2]/div[1]/ul/li[3]/div[2]/div/div[1]/div[2]/ul'),
    'host': (By.XPATH, '/html/body/div/main/form/table/tbody/tr/td[2]/a'),
    'tela_inicial': (By.CSS_SELECTOR, '#page-title-general'),
    'X': (By.XPATH, '/html/body/div/main/div[1]/form/div/div[1]/div/div[1]/ul/li[1]/div[2]/div/div[1]/div[2]/ul/li[1]/span/span[2]'),
    'grupo_host_input': (By.CSS_SELECTOR, '#groups__ms'),
    'nome_host_input': (By.CSS_SELECTOR, '#filter_host'),
    'IP_input': (By.CSS_SELECTOR, '#filter_ip'),
    'lista_hosts': (By.CSS_SELECTOR, '#cancel'),
    'erro': (By.XPATH, '/html/body/div/output[1]/a'),
    'IP': (By.XPATH, '/html/body/div/main/form/div[2]/div[1]/ul/li[4]/div[2]/div[1]/div[4]/div/div[3]/input'),
    'proxy': (By.CSS_SELECTOR, '#filter_monitored_by > li:nth-child(1) > label'),
    'host_cadastrados': (By.XPATH, '/html/body/div/main/form/table/tbody'),

    # Ação
    'aplicar': (By.CSS_SELECTOR, '#tab_0 > div:nth-child(2) > button:nth-child(1)'),
    'clonar': (By.CSS_SELECTOR, '#clone'),
    'nome_host': (By.CSS_SELECTOR, '#host'),
    'adicionar': (By.CSS_SELECTOR, '#add'),
    'atualizar': (By.CSS_SELECTOR, '#update'),
    'cancelar': (By.CSS_SELECTOR, '#cancel'),
    
}

driver.get('https://zabbix.tely.com.br/index.php')

WebDriverWait(driver, 10).until(EC.element_to_be_clickable(locator['usuario'])).send_keys('Admin')
driver.find_element(*locator['senha']).send_keys('ZAmw390711@')
driver.find_element(*locator['entrar']).click()

WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator['tela_inicial']))

driver.get('https://zabbix.tely.com.br/hosts.php')

planilha = pd.read_excel('MW CORP e GOV.xlsx', sheet_name='Inclusão 2025 (PRTG+Zabbix) - N')


############### Para filtrar alguma coluna ###############
# planilha = planilha[
#    (planilha['Fabricante'].str.upper() != 'SWITCHES MIKROTIK') &
#    (planilha['Fabricante'].str.upper() != '')
# ]

# planilha = planilha[planilha['Fabricante'].str.upper() == 'SWITCHES MIKROTIK']

def Limpar_filtros():
    grupos_elementos_ul = WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator['grupos_elementos_ul'])).find_elements(By.XPATH, 'li')
    if len(grupos_elementos_ul) > 0:
        for item in grupos_elementos_ul:
            item.find_element(By.XPATH, 'span/span[2]').click()

    driver.find_element(*locator['nome_host_input']).clear()
    sleep(2)

def Limpar_filtros():
    grupos_elementos_ul = WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator['grupos_elementos_ul'])).find_elements(By.XPATH, 'li')
    if len(grupos_elementos_ul) > 0:
        for item in grupos_elementos_ul:
            item.find_element(By.XPATH, 'span/span[2]').click()

    driver.find_element(*locator['nome_host_input']).clear()
    sleep(2)

    driver.find_element(*locator['IP_input']).clear()


def Concatenar_nome_ip(linha):
        nome = linha['NOME']
        nome = unicodedata.normalize('NFD', nome)
        nome = nome.encode('ascii', 'ignore').decode('utf-8')
        nome = nome.replace('ç', 'c').replace('Ç', 'C')
        nome_do_host_regex = re.sub(r"[^a-zA-Z0-9._ \-]", "", nome)
        nome_ip = f"{nome_do_host_regex} - {linha['IP']}"
        return nome_ip

def Clonar_host_():
    for _, linha in planilha.iterrows():

        Limpar_filtros()

        driver.find_element(*locator['grupo_host_input']).send_keys(linha['Fabricante'])
        sleep(1)
        driver.find_element(*locator['grupo_host_input']).send_keys(Keys.ENTER)
        sleep(2)
        driver.find_element(*locator['aplicar']).click()

        WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator['host'])).click()
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator['clonar'])).click()

        nome_ip = Concatenar_nome_ip(linha)

        sleep(3)
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator['nome_host']))
        campo_nome_host = driver.find_element(*locator['nome_host'])
        campo_nome_host.clear()

        sleep(3)
        driver.find_element(*locator['nome_host']).send_keys(nome_ip)

        driver.find_element(*locator['IP']).clear()
        sleep(2)
        driver.find_element(*locator['IP']).send_keys(linha['IP'])
        sleep(2)
        driver.find_element(*locator['adicionar']).click()

        sleep(3)
        if driver.find_elements(*locator['erro']):
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator['cancelar'])).click()


        sleep(10)

def Adicao_de_grupo():
    for i, linha in planilha.iterrows():


        nome_ip = Concatenar_nome_ip(linha)

        Limpar_filtros()

        driver.find_element(*locator['nome_host_input']).send_keys(nome_ip)
        sleep(4)
        driver.find_element(*locator['IP_input']).send_keys(linha['IP'])
        sleep(4)
        driver.find_element(*locator['aplicar']).click()
        
        if driver.find_element(By.XPATH,'/html/body/div/main/form/table/tbody/tr/td').text == 'Sem dados encontrados.':    
            planilha.loc[i,'Verificação do bot'] = 'Host não encontrado'

        else: 

            WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator['host'])).click()
            grupos_host_ul = WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator['grupos_host_ul'])).find_elements(By.XPATH, 'li')
            adicionar = []

            if len(grupos_host_ul) > 0:
                for item in grupos_host_ul:
            
                    if item.text == linha['Categoria'].upper():
                        adicionar.append(1)

                    if item.text == linha['Grupo'].upper():
                        adicionar.append(2)

                    if item.text == linha['Grupos (Host)'].upper():
                        adicionar.append(3)
            
            if not 1 in adicionar:
                # Categoria
                driver.find_element(*locator['grupo_host_input']).send_keys(linha['Categoria'].upper())
                sleep(5)
                driver.find_element(*locator['grupo_host_input']).send_keys(Keys.ENTER)

            if not 2 in adicionar:                                                           
                # Categoria
                driver.find_element(*locator['grupo_host_input']).send_keys(linha['Grupo'].upper())
                sleep(5)
                driver.find_element(*locator['grupo_host_input']).send_keys(Keys.ENTER)

            if not 3 in adicionar:
                # Grupos (Host)
                driver.find_element(*locator['grupo_host_input']).send_keys(linha['Grupos (Host)'].upper())
                sleep(5)
                driver.find_element(*locator['grupo_host_input']).send_keys(Keys.ENTER)     
            
            planilha.loc[i,'Verificação do bot'] = 'Incluído nos grupos'

            sleep(5)
            driver.find_element(*locator['atualizar']).click()


        sleep(10)

    planilha.to_excel('Planilha dos grupos atualziada.xlsx', index=False)

def Verificacao_de_duplicidade():
    for i, linha in planilha.iterrows():

        
        Limpar_filtros()


        driver.find_element(*locator['IP_input']).send_keys(linha['IP'])
        # driver.find_element(*locator['IP_input']).send_keys('10.8.34.138')
        driver.find_element(*locator['proxy']).click()
        sleep(2)
        driver.find_element(*locator['aplicar']).click()
        
        repedidos = 0


        grupos_elementos_ul = WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator['host_cadastrados'])).find_elements(By.XPATH, 'tr')
        
        if driver.find_element(By.XPATH,'/html/body/div/main/form/table/tbody/tr/td').text == 'Sem dados encontrados.':
            planilha.loc[i,'Status'] = 'Não adicionado'
            
        elif len(grupos_elementos_ul) > 0:
            quantidade = 0
            for item in grupos_elementos_ul:
                grupos_elementos_ul = WebDriverWait(driver, 30).until(EC.element_to_be_clickable(locator['host_cadastrados'])).find_element(By.XPATH, f'/html/body/div/main/form/table/tbody/tr[{1}]/td[9]')
                if grupos_elementos_ul.text == f"{linha['IP']}: 161":
                    repedidos += 1
                quantidade += 1
                
            if repedidos > 1:
                planilha.loc[i,'Status'] = 'Duplicado'
            else:
                planilha.loc[i,'Status'] = 'Único'
        
        sleep(10)

    planilha.to_excel('Atualizado.xlsx', index=False)
    



Adicao_de_grupo()
print('fim do codigo')
            