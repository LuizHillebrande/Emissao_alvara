import pyautogui
from time import sleep
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import customtkinter as ctk
import os
import sys
import threading

if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

def atualizar_progresso_maringa(progresso_atual, total_linhas):
    progresso = progresso_atual / total_linhas
    progress_maringa.set(progresso)
    app.update_idletasks()  # Atualiza a interface gráfica

def atualizar_progresso_tapejara(progresso_atual, total_linhas):
    progresso = progresso_atual / total_linhas
    progress_tapejara.set(progresso)
    app.update_idletasks()

maringa_file = os.path.join(application_path, 'codigos_maringa.xlsx')
tapejara_file = os.path.join(application_path, 'codigos_tapejara.xlsx')

wb_debitos_maringa = openpyxl.load_workbook('codigos_maringa.xlsx')
sheet_debitos_maringa = wb_debitos_maringa['Planilha1']

wb_debitos_tapejara = openpyxl.load_workbook('codigos_tapejara.xlsx')
sheet_debitos_tapejara = wb_debitos_tapejara['Planilha1']

wb_resultado = openpyxl.Workbook()
sheet_resultado = wb_resultado.active
sheet_resultado.title = "Empresas Sem Débitos"
sheet_resultado.append(['Nome da Empresa', 'Código Municipal', 'Mensagem'])

def ler_progresso_maringa():
    if os.path.exists("progresso_maringa.txt"):
        with open("progresso_maringa.txt", "r") as file:
            return int(file.read().strip())
    return 2  # Se não houver progresso registrado, começa da linha 2

# Função para salvar o progresso
def salvar_progresso_maringa(linha):
    with open("progresso_maringa.txt", "w") as file:
        file.write(str(linha))


def ler_progresso_tapejara():
    arquivo_progresso = "progresso_tapejara.txt"
    if os.path.exists(arquivo_progresso):
        with open(arquivo_progresso, "r") as file:
            return int(file.read().strip())
    else:
        # Cria o arquivo de progresso se não existir
        with open(arquivo_progresso, "w") as file:
            file.write("2")  # Inicia com a linha 2
        return 2

def salvar_progresso_tapejara(linha):
    with open("progresso_tapejara.txt", "w") as file:
        file.write(str(linha))

def pegar_debitos_maringa_thread():
    # Função que roda o Selenium em uma thread separada
    thread = threading.Thread(target=pegar_debitos_maringa)
    thread.start()

def pegar_debitos_tapejara_thread():
    thread = threading.Thread(target=pegar_debitos_tapejara)
    thread.start()


def pegar_debitos_maringa():
    driver = webdriver.Chrome()
    driver.get('https://tributos.maringa.pr.gov.br/portal-contribuinte/consulta-debitos')
    
    ultima_linha_processada_maringa = ler_progresso_maringa()
    
    total_linhas = sheet_debitos_maringa.max_row
    for linha in sheet_debitos_maringa.iter_rows(min_row=ultima_linha_processada_maringa, max_row=5):
            nome_empresa_maringa = linha[0].value
            codigo_municipal_maringa = linha[1].value

            select_element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//select[@id='select-filter']"))
                )
            
            select = Select(select_element)
            select.select_by_value("1")  
                
            sleep(1)

            if not codigo_municipal_maringa:
                continue

            campo_cod_municipal = WebDriverWait(driver, 20).until(
                EC.visibility_of_element_located((By.XPATH, "(//input[@placeholder='Digite o cadastro...'])[2]"))
            )
            campo_cod_municipal.clear()
            sleep(1)
            campo_cod_municipal.send_keys(str(codigo_municipal_maringa))

            pyautogui.press('TAB')
            sleep(1)
            pyautogui.press('ENTER')
            
            sleep(2)
        
            empresa_sem_debitos = EC.visibility_of_element_located((By.XPATH,"//article[@class='info mt-xs']"))
            if empresa_sem_debitos:
                 sheet_resultado.append([nome_empresa_maringa, codigo_municipal_maringa, 'Empresa sem débitos'])
                 sleep(1)
            else:
                  continue
            
            salvar_progresso_maringa(linha[0].row + 1)
            app.after(0, atualizar_progresso_maringa, linha[0].row, total_linhas)

                 

    wb_resultado.save('empresas_sem_debitos_maringa.xlsx')
    driver.quit()  

def pegar_debitos_tapejara():
    
    driver = webdriver.Chrome()
    driver.get('https://tapejara.eloweb.net/portal-contribuinte/consulta-debitos')

    ultima_linha_processada_tapejara = ler_progresso_tapejara()
    total_linhas = sheet_debitos_tapejara.max_row

    for linha in sheet_debitos_tapejara.iter_rows(min_row=ultima_linha_processada_tapejara, max_row=5):
            nome_empresa_tapejara = linha[0].value
            codigo_municipal_tapejara = linha[1].value
            
            selecionar_elemento = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//select[@id='select-filter']"))
            )
            
            selecionar = Select(selecionar_elemento)
            selecionar.select_by_value("1")  
            sleep(1) 
            
            if not codigo_municipal_tapejara:
                continue

            campo_cod_municipal = WebDriverWait(driver, 20).until(
                EC.visibility_of_element_located((By.XPATH, "(//input[@placeholder='Digite o cadastro...'])[2]"))
            )
            campo_cod_municipal.clear()
            sleep(1)
            campo_cod_municipal.send_keys(str(codigo_municipal_tapejara))

            pyautogui.press('TAB')
            sleep(1)
            pyautogui.press('ENTER')
            sleep(2)
        
            empresa_sem_debitos = EC.visibility_of_element_located((By.XPATH,"//article[@class='info mt-xs']"))
            if empresa_sem_debitos:
                 sheet_resultado.append([nome_empresa_tapejara, codigo_municipal_tapejara, 'Empresa sem débitos'])
            else:
                  print('empresa com debitos')
            
            salvar_progresso_tapejara(linha[0].row + 1)
            app.after(0, atualizar_progresso_tapejara, linha[0].row, total_linhas)
                 

    wb_resultado.save('empresas_sem_debitos.xlsx')
    driver.quit()  

#INTERFACE GRÁFICA

ctk.set_appearance_mode("Dark")  
ctk.set_default_color_theme("dark-blue")  

app = ctk.CTk()
app.title("Office automation")
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
app.geometry(f"{screen_width}x{screen_height-40}+0+0")

title_label = ctk.CTkLabel(
    app, 
    text="Controle de Débitos Municipais", 
    font=("Helvetica", 30, "bold"), 
    text_color="#ffffff"
)
title_label.pack(pady=20)

description_label = ctk.CTkLabel(
    app, 
    text="Selecione a prefeitura para consultar os débitos das empresas.", 
    font=("Helvetica", 16), 
    text_color="#c9c9c9", 
    wraplength=450,  
    justify="center"
)
description_label.pack(pady=10)

progress_maringa = ctk.CTkProgressBar(app, width=300)
progress_maringa.set(0.0)  # Inicializa com progresso 0%
progress_maringa.pack(pady=10)


button_maringa = ctk.CTkButton(
    app, 
    text="Prefeitura de Maringá/PR", 
    command=pegar_debitos_maringa_thread,
    font=("Helvetica", 14), 
    width=300, 
    height=40, 
    fg_color="#007acc",  
    hover_color="#005b99",  
)
button_maringa.pack(pady=15)

progress_tapejara = ctk.CTkProgressBar(app, width=300)
progress_tapejara.set(0.0)  # Inicializa com progresso 0%
progress_tapejara.pack(pady=10)

button_tapejara = ctk.CTkButton(
    app, 
    text="Prefeitura de Tapejara/PR", 
    command=pegar_debitos_tapejara_thread,
    font=("Helvetica", 14), 
    width=300, 
    height=40, 
    fg_color="#007acc",  
    hover_color="#005b99",  
)
button_tapejara.pack(pady=15)


footer_label = ctk.CTkLabel(
    app, 
    text="Desenvolvido por Luiz Fernando Hillebrande", 
    font=("Helvetica", 10), 
    text_color="#c9c9c9"
)
footer_label.pack(side="bottom", pady=25)

app.grid_rowconfigure(0, weight=1)
app.grid_columnconfigure(0, weight=1)

def sair_tela_cheia( event = None):
    app.attributes('-fullscreen', False)

app.bind("<Escape>", sair_tela_cheia)

app.mainloop()
