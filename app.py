import pyautogui
from time import sleep
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC


wb_debitos_maringa = openpyxl.load_workbook('codigos_maringa.xlsx')
sheet_debitos_maringa = wb_debitos_maringa['Planilha1']

wb_debitos_tapejara = openpyxl.load_workbook('codigos_tapejara.xlsx')
sheet_debitos_tapejara = wb_debitos_tapejara['Planilha1']

wb_resultado = openpyxl.Workbook()
sheet_resultado = wb_resultado.active
sheet_resultado.title = "Empresas Sem Débitos"
sheet_resultado.append(['Nome da Empresa', 'Código Municipal', 'Mensagem'])

def pegar_debitos_maringa():
    driver = webdriver.Chrome()
    driver.get('https://tributos.maringa.pr.gov.br/portal-contribuinte/consulta-debitos')
    
        
    select_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//select[@id='select-filter']"))
        )
    
    select = Select(select_element)
    select.select_by_value("1")  
        
    sleep(1)

    for linha in sheet_debitos_maringa.iter_rows(min_row=2,max_row=2):
            nome_empresa_maringa = linha[0].value
            codigo_municipal_maringa = linha[1].value

            pyautogui.press('TAB')
            pyautogui.write(str(codigo_municipal_maringa))
            sleep(1)

            pyautogui.press('TAB')
            sleep(1)
            pyautogui.press('ENTER')
            sleep(2)
        
            empresa_sem_debitos = EC.visibility_of_element_located((By.XPATH,"//article[@class='info mt-xs']"))
            if empresa_sem_debitos:
                 sheet_resultado.append([nome_empresa_maringa, codigo_municipal_maringa, 'Empresa sem débitos'])
            else:
                  continue
            

                 

    wb_resultado.save('empresas_sem_debitos_maringa.xlsx')
    driver.quit()  

def pegar_debitos_tapejara():
    driver = webdriver.Chrome()
    driver.get('https://tapejara.eloweb.net/portal-contribuinte/consulta-debitos')
      
    selecionar_elemento = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//select[@id='select-filter']"))
    )
    
    selecionar = Select(selecionar_elemento)
    selecionar.select_by_value("1")  
    sleep(1) 

    for linha in sheet_debitos_tapejara.iter_rows(min_row=2,max_row=2):
            nome_empresa_tapejara = linha[0].value
            codigo_municipal_tapejara = linha[1].value

            pyautogui.press('TAB')
            pyautogui.write(str(codigo_municipal_tapejara))
            sleep(1)

            pyautogui.press('TAB')
            sleep(1)
            pyautogui.press('ENTER')
            sleep(2)
        
            empresa_sem_debitos = EC.visibility_of_element_located((By.XPATH,"//article[@class='info mt-xs']"))
            if empresa_sem_debitos:
                 sheet_resultado.append([nome_empresa_tapejara, codigo_municipal_tapejara, 'Empresa sem débitos'])
            else:
                  print('empresa com debitos')
            

                 

    wb_resultado.save('empresas_sem_debitos.xlsx')
    driver.quit()  
    

pegar_debitos_tapejara()
