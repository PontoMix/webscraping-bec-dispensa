import time

#Importando o debugger do python para criar breakpoints e verificar se está funcionando o código criado
import pdb

#Definir um determinado período de tempo para que um elemento apareça antes de prosseguir é com Web Driver Wait. 
from selenium.webdriver.support.wait import WebDriverWait

#EC significa condições esperadas, que são condições que devem ser cumpridas para que uma determinada ação possa ser tomada
from selenium.webdriver.support import expected_conditions as EC

import os

##Importando o webdriver da biblioteca Selenium para acessar a Internet e carregar páginas
from selenium import webdriver

##Importando o Options para poder manipular o webdriver e suas propriedades, 
#como executar várias operações, desativar extensões, desativar pop-ups, etc
from selenium.webdriver.chrome.options import Options

#Importando as chaves especiais para utilizar teclas no teclado ao usar Selenium e 
#apagar textos pré-preenchido em campos de entradas
from selenium.webdriver.common.keys import Keys

#Importando o By para localizar elementos em uma página web
from selenium.webdriver.common.by import By

#Importando ActionChains para utilizar o mouse e teclas do teclado
from selenium.webdriver.common.action_chains import ActionChains

#Pandas para criar, manipular e visualizar tabelas
import pandas as pd

#Para formatar data e horário
from datetime import datetime

#OpenXL para escrever arquivos .xls (Excel) com o Pandas por meio do módulo "Xmlt" para arquivos .xls
import openpyxl 

#Time para utilizar datas quando salvar os arquivos .xls
import time

import re


#Pegando a data e horário de hoje, no momento que criou o arquivo .xls 
TodayDate = time.strftime("%d-%m-%Y %H-%M-%S")
DateSheet = time.strftime("%d-%m-%Y")

#Criando um nome padronizado para os arquivos .xls que terá os detalhes de cada OC
excelfilename = "Dispensa Detalhada Completa - " + TodayDate + ".xls"

home_dir = os.path.expanduser("~")
path_with_filename = os.path.join("C:\\", "WebScraping Licitações - Dispensa", "Detalhes Produtos - Dispensa", excelfilename)

#Criando um nome padronizado para os arquivos .xls que terá as informações da tabela das OCs
excelfilenameallocs = "Tabela Dispensa Completa - " + TodayDate + ".xls" 

path_with_filenameallocs = os.path.join("C:\\", "WebScraping Licitações - Dispensa", "Tabela OCs - Dispensa", excelfilenameallocs)

#Criando um nome padronizado para as folhas do .xls
excelsheet = "Dispensa Completa - " + DateSheet

#Importando as variáveis de ambiente para utilizar com segurança o login e senha do usuário
from config import database_infos

get_login = database_infos['login']
get_pass = database_infos['password']
get_username_pc = database_infos['username_pc']



def bec_allocsdispensa():
    
    browser_driver = webdriver.Chrome()

    #Fazendo solicitação para abrir e navegar na página da BEC
    browser_driver.get("https://www.bec.sp.gov.br/BECSP/Home/Home.aspx")

    #Inicializando o WebDriverWait
    waitWDW = WebDriverWait(browser_driver, 10)

    #Maximizando a Tela do Browser
    browser_driver.maximize_window()

    #Confirmando que é o site correto aquele que está aberto
    assert "BEC" in browser_driver.title

    #Procurando a tag certa do botão "Negociações Eletrônicas" 
    btn_ne = browser_driver.find_element(By.LINK_TEXT, "Negociações Eletrônicas")
 
    ##Fazendo com que clique no botão
    btn_ne.send_keys(Keys.RETURN)

    #Procurando as tags certas com XPATH e preenchendo os campos "CNJP/CPF" e "Senha"
    login = browser_driver.find_element(By.XPATH, "//input[@id='TextLogin']") #Se parar de funcionar, utilize a class="TextLogin" ou o id="TextLogin"
    login.send_keys(get_login)

    password = browser_driver.find_element(By.XPATH, "//*[@id='TextSenha']") #Se parar de funcionar, utilize a class="TextSenha" ou o id="TextSenha"
    password.send_keys(get_pass)

    #Marcando a caixa de declaração
    statement_box = browser_driver.find_element(By.XPATH, "//*[@id='chkAceite']") #Se parar de funcionar, utilize a class="chkAceite" ou o id="chkAceite"
    statement_box.click()

    #Clicando no botão de entrar
    btn_enter = browser_driver.find_element(By.ID, "Btn_Confirmar") #Se parar de funcionar, utilize o id="Btn_Confirmar"
    btn_enter.click()

    current_url = browser_driver.current_url
    
    if current_url == "https://www.bec.sp.gov.br/fornecedor_ui/TermoResponsabilidade.aspx?Dzqeio6gALuoR%2flQf2tFB6zBkp9ETq5P44%2bgrURdFf66JmFgqUpWHFjTKO2RLNZR":
        waitWDW = WebDriverWait(browser_driver, 10)
        reconfirm_checkbox = browser_driver.find_element(By.ID, "//*[@id='ctl00_c_area_conteudo_chkDeclaracao']")
        reconfirm_checkbox.click()
        ok_button = browser_driver.find_element(By.ID, "//*[@id='ctl00_c_area_conteudo_Button1']")
        ok_button.click()
        #Passando o mouse por cima da lista "Participar"
        join_menu_list = waitWDW.until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Participar']"))) 
        actions = ActionChains(browser_driver)
        actions.move_to_element(join_menu_list).pause(2).perform()
        
        #Escolhendo o item da lista certa, que é o Dispensa e clicando nele
        convite_item_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Dispensa de Licitação - Cotações']"))) 
        convite_item_list.click()
        
        situation_select_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_Wuc_filtroPesquisaOc1_c_ddlListaSituacao']")))
        situation_select_list.click()
        
        all_option_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_Wuc_filtroPesquisaOc1_c_ddlListaSituacao']/option[1]")))
        all_option_list.click()

        pe_btn_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_bt33022_Pesquisa']")))  #Se parar de funcionar, utilizar o id="ctl00_conteudo_Pesquisa", css_selector="#pesquisa" ou text_link="Pesquisar"
        pe_btn_search.click()
        time.sleep(2)  
    
    else:
        #Passando o mouse por cima da lista "Participar"
        join_menu_list = waitWDW.until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Participar']")))
        actions = ActionChains(browser_driver)
        actions.move_to_element(join_menu_list).pause(2).perform()
        
        
        #Escolhendo o item da lista certa, que é o Dispensa e clicando nele
        convite_item_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Dispensa de Licitação - Cotações']")))
        convite_item_list.click()
        
        situation_select_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_Wuc_filtroPesquisaOc1_c_ddlListaSituacao']")))
        situation_select_list.click()
        
        all_option_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_Wuc_filtroPesquisaOc1_c_ddlListaSituacao']/option[1]")))
        all_option_list.click()
        
        pe_btn_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_bt33022_Pesquisa']")))  #Se parar de funcionar, utilizar o id="ctl00_conteudo_Pesquisa", css_selector="#pesquisa" ou text_link="Pesquisar"
        pe_btn_search.click()
        time.sleep(2)
        
    ###Lista para armazenar os resultados da coleta de dados do Dispensa###
    result_all_ocs = []
    
    #Lista para armazenar os resultados da coleta de dados da descrição, quantidade, uf, telefone e e-mails do Dispensa
    details_products_dispensa = []    
    
    global iterator, m, n, i, j, k
    iterator = 1
    m = 0
    n = 0
    i = 1
    j = 1
    k = 1
    
    #Procurando todos os elementos da tabela da Dispensa    
    convite_oc_len = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[3]/a")[0:] 
    convite_uc = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[7]/table/tbody/tr[2]/td")[0:]
    convite_town = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[7]/table/tbody/tr[3]/td")[0:] 
    convite_situation = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[7]/table/tbody/tr[4]/td")[0:]
    convite_oc_number = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[3]/a")[0:] 
    
        
    #rows = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[3]")[0]

    for i in range(len(convite_oc_len)):
    #for i, row in enumerate(rows[iterator:], start=i+1):
        
        convite_initialdate = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[4]")[i]
       
        #Separando por meio do espaço a data do horário
        date_time_list = convite_initialdate.text.split(" ")
        date_value = date_time_list[0]
        hour_value = date_time_list[1]
        
        #Convertendo nos tipos de dados corretos
        date_obj = datetime.strptime(date_value, '%d/%m/%Y').date()
        time_obj = datetime.strptime(hour_value, '%H:%M:%S').time()
 
        result_all_ocs.append({"UC": convite_uc[i].text,
                               "Cidade": convite_town[i].text,
                               "OC": convite_oc_number[i].text,
                               "Data": date_obj,
                               "Hora": time_obj,
                               "Situação": convite_situation[i].text})  
        
        
        
        
    table_oc = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']")
    rows_ocs = table_oc.find_elements(By.XPATH, "//tbody/tr/td[2]")
    for m in range(k, len(rows_ocs)):
                          
                        #Entrando em cada OC para pegar as informações (descrições) dos produtos que desejam      
                        link = waitWDW.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/form/div[3]/div/div/div/div/div[2]/div[4]/div[2]/div/table/tbody/tr[{k+1}]/td[3]/a"))) 
                        #Pressionado CRTL e clicando no link
                        ActionChains(browser_driver).key_down(Keys.CONTROL).click(link).key_up(Keys.CONTROL).perform()

                        # Mudando para a nova aba aberta
                        browser_driver.switch_to.window(browser_driver.window_handles[-1])
                        time.sleep(2)
        
                        #Pegando os valores essenciais dos produtos
                        #Pegando o número da OC para colocar junto com os detalhes dos itens
                        oc_number_dispensa = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_DetalhesOfertaCompra1_txtOC']").text 
                        #Tabela com os detalhes dos itens que estão sendo solicitidas pela OC na Dispensa
                        details_table_oc_dispensa = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grd_fornecedor_lance']") 
                        #Linhas da tabela
                        rows_details_oc_dispensa = details_table_oc_dispensa.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grd_fornecedor_lance']/tbody/tr")[1:]
        
        
                        #Se a tabela possuir mais do que 1 item (1 linha com detalhes da descrição), o seguinte código será executado:
                        if len(rows_details_oc_dispensa) > 1:
                        
                                        item_values = []
                                        code_values = []
                                        description_values = []
                                        quantity_values = []
                                        uf_values = []
                                        minreductionporcent = []
                                        
                                        
                                        for row in rows_details_oc_dispensa:
                                            
                                            item_value = row.find_element(By.XPATH, "./td[3]").text
                                            code_value = row.find_element(By.XPATH, "./td[4]").text
                                            description_value = row.find_element(By.XPATH, "./td[5]").text
                                            quantity_value = row.find_element(By.XPATH, "./td[6]").text
                                            uf_value = row.find_element(By.XPATH, "./td[7]").text
                                            minreductionporcent = row.find_element(By.XPATH, "./td[11]").text
                        
                                            item_values = int(item_value)
                                            code_values = int(code_value)
                                            description_values = description_value
                                            quantity_values = int(float(quantity_value))
                                            uf_values = uf_value
                                            minreductionporcents = minreductionporcent
                                            minreductionporcents = minreductionporcents.replace('%', '')
                                            minreductionporcents = float(minreductionporcents)
                                            
                                            
                                            details_products_dispensa.append({
                                                "OC": oc_number_dispensa,
                                                "Item": item_values,
                                                "SIAF.": code_values,
                                                "Desc.": description_values,
                                                "Qtd.": quantity_values,
                                                "Und.": uf_values,
                                                "Red. Mínima": minreductionporcents})
                                            
                                            
                                        #Fechando a aba e voltando para a aba principal
                                        browser_driver.back()
                                        k+=1
                                        time.sleep(5)
                                        
                        
                        else:
                                        for row in rows_details_oc_dispensa:
                                            
                                            item_value = row.find_element(By.XPATH, "./td[3]").text
                                            code_value = row.find_element(By.XPATH, "./td[4]").text
                                            description_value = row.find_element(By.XPATH, "./td[5]").text
                                            quantity_value = row.find_element(By.XPATH, "./td[6]").text
                                            uf_value = row.find_element(By.XPATH, "./td[7]").text
                                            minreductionporcent = row.find_element(By.XPATH, "./td[11]").text
                        
                                            item_values = int(item_value)
                                            code_values = int(code_value)
                                            description_values = description_value
                                            quantity_values = int(float(quantity_value))
                                            uf_values = uf_value
                                            minreductionporcents = minreductionporcent
                                            minreductionporcents = minreductionporcents.replace('%', '')
                                            minreductionporcents = float(minreductionporcents)
                                            
                                            
                                            details_products_dispensa.append({
                                                "OC": oc_number_dispensa,
                                                "Item": item_values,
                                                "SIAF.": code_values,
                                                "Desc.": description_values,
                                                "Qtd.": quantity_values,
                                                "Und.": uf_values,
                                                "Red. Mínima": minreductionporcents})
                                            
                                            
                                        #Fechando a aba e voltando para a aba principal
                                        browser_driver.back()
                                        k+=1
                                        time.sleep(5)
                                        
    #Pause de 20 segundos depois de fazer Scraping de 15 páginas, para depois continuar e diminuir as chances de dar algum erro             
    #if (m+1) % 15 == 0:
    #       time.sleep(20)
                                        
    #Fechando o Browser depois de terminar
    browser_driver.close()
    browser_driver.quit()
        
        
        
        
         

                                                ######################################################################################
                                                ##Criando uma table (Excel) para visualizar os valores coletados da lista do Dispensa## 
                                                ######################################################################################
            
    #Utilizando pandas para criar e visualizar uma tabela formatada com os valores coletados da lista da Dispensa
    
    #Tabela de OCs
    df_final_table_ocs = pd.DataFrame(result_all_ocs)
    
    #Valores coletados dentro de cada OC (detalhes)
    df_final_products_ocs = pd.DataFrame(details_products_dispensa) 
    
    #Criando um Pandas Excel Writer para usar o Openpyxl como engine e salvar os detalhes das OCs selecionadas.
    writer = pd.ExcelWriter(path_with_filename, engine='openpyxl')
    
    #Criando um Pandas Excel Writer para salvar os dados atuais da tabela de OCs
    writer2 = pd.ExcelWriter(path_with_filenameallocs, engine='openpyxl')

    #Criando um arquivo .xls para utilizar os dados dos detalhes de OCs no Excel
    df_final_products_ocs.to_excel(writer, sheet_name=DateSheet, header=True, index=False)
    
    #Criando arquivo .xls para ver os dados gerais da tabela de OCs
    df_final_table_ocs.to_excel(writer2, sheet_name=DateSheet, header=True, index=False)
    
    #print(df_oc_details_data.dtypes)
    
    #Fechando o Pandas Excel Writer e fazendo o output do arquivo .xls
    writer.close()
    writer2.close() 

    #print(df_final_data)
    print(df_final_table_ocs)
    print(df_final_products_ocs)
    print('DataFrame is written to Excel File successfully!!!')
    
#bec_allocsdispensa()