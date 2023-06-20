import time
import pdb
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from datetime import datetime
import openpyxl 
import time
import re
 
TodayDate = time.strftime("%d-%m-%Y %H-%M-%S")
DateSheet = time.strftime("%d-%m-%Y")

excelfilename = "Dispensa Detalhada Reduzida - " + TodayDate + ".xls"

home_dir = os.path.expanduser("~")
path_with_filename = os.path.join("C:\\", "WebScraping Licitações - Dispensa", "Detalhes Produtos - Dispensa", excelfilename)

excelfilenameallocs = "Tabela Dispensa Reduzida - " + TodayDate + ".xls" 

path_with_filenameallocs = os.path.join("C:\\", "WebScraping Licitações - Dispensa", "Tabela OCs - Dispensa", excelfilenameallocs)

excelsheet = "Dispensa Reduzida - " + DateSheet

from config import database_infos

get_login = database_infos['login']
get_pass = database_infos['password']
get_username_pc = database_infos['username_pc']


def bec_filterdispensa(field_value):
    
    name_category = field_value
    
    browser_driver = webdriver.Chrome()

    browser_driver.get("https://www.bec.sp.gov.br/BECSP/Home/Home.aspx")

    waitWDW = WebDriverWait(browser_driver, 10)

    browser_driver.maximize_window()

    assert "BEC" in browser_driver.title
 
    btn_ne = browser_driver.find_element(By.LINK_TEXT, "Negociações Eletrônicas")
 
    btn_ne.send_keys(Keys.RETURN)

    login = browser_driver.find_element(By.XPATH, "//input[@id='TextLogin']")
    login.send_keys(get_login)

    password = browser_driver.find_element(By.XPATH, "//*[@id='TextSenha']")
    password.send_keys(get_pass)

    statement_box = browser_driver.find_element(By.XPATH, "//*[@id='chkAceite']") 
    statement_box.click()

    btn_enter = browser_driver.find_element(By.ID, "Btn_Confirmar") 
    btn_enter.click()

    current_url = browser_driver.current_url
    
    if current_url == "https://www.bec.sp.gov.br/fornecedor_ui/TermoResponsabilidade.aspx?Dzqeio6gALuoR%2flQf2tFB6zBkp9ETq5P44%2bgrURdFf66JmFgqUpWHFjTKO2RLNZR":
        waitWDW = WebDriverWait(browser_driver, 10)
        reconfirm_checkbox = browser_driver.find_element(By.ID, "//*[@id='ctl00_c_area_conteudo_chkDeclaracao']")
        reconfirm_checkbox.click()
        ok_button = browser_driver.find_element(By.ID, "//*[@id='ctl00_c_area_conteudo_Button1']")
        ok_button.click()
        
        join_menu_list = waitWDW.until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Participar']"))) 
        actions = ActionChains(browser_driver)
        actions.move_to_element(join_menu_list).pause(2).perform()
        
        convite_item_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Dispensa de Licitação - Cotações']"))) 
        convite_item_list.click()
        
        situation_select_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_Wuc_filtroPesquisaOc1_c_ddlListaSituacao']")))
        situation_select_list.click()
        
        all_option_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_Wuc_filtroPesquisaOc1_c_ddlListaSituacao']/option[1]")))
        all_option_list.click()

        pe_btn_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_bt33022_Pesquisa']")))  
        pe_btn_search.click()
        time.sleep(2)  
        
    
    else:
       
        join_menu_list = waitWDW.until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Participar']")))
        actions = ActionChains(browser_driver)
        actions.move_to_element(join_menu_list).pause(2).perform()
        
        convite_item_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Dispensa de Licitação - Cotações']")))
        convite_item_list.click()
        
        situation_select_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_Wuc_filtroPesquisaOc1_c_ddlListaSituacao']")))
        situation_select_list.click()
        
        all_option_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_Wuc_filtroPesquisaOc1_c_ddlListaSituacao']/option[1]")))
        all_option_list.click()
        
        pe_btn_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_bt33022_Pesquisa']"))) 
        pe_btn_search.click()
        time.sleep(2)
        
    
    
    input_category = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_c_area_conteudo_Wuc_filtroPesquisaOc1_c_txt_ItemMaterial_desc']")  
    input_category.send_keys(name_category)    
        
    button_advanced_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_bt33022_Pesquisa']")))  
    button_advanced_search.click()

    result_all_ocs = []
    details_products_dispensa = []    
    
    global iterator, m, n, i, j, k
    iterator = 1
    m = 0
    n = 0
    i = 0
    j = 1
    k = 1
    
    convite_oc_len = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[3]/a")[0:] 
    convite_uc = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[7]/table/tbody/tr[2]/td")[0:]
    convite_town = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[7]/table/tbody/tr[3]/td")[0:] 
    convite_situation = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[7]/table/tbody/tr[4]/td")[0:]
    convite_oc_number = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[3]/a")[0:] 
    
    for i in range(len(convite_oc_len)):
        
        convite_initialdate = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr/td[4]")[i]
        date_time_list = convite_initialdate.text.split(" ")

        date_value = date_time_list[0]
        hour_value = date_time_list[1]
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
                                
                        link = waitWDW.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/form/div[3]/div/div/div/div/div[2]/div[4]/div[2]/div/table/tbody/tr[{k+1}]/td[3]/a"))) 
                        ActionChains(browser_driver).key_down(Keys.CONTROL).click(link).key_up(Keys.CONTROL).perform()
                        browser_driver.switch_to.window(browser_driver.window_handles[-1])
                        time.sleep(2)
        

                        oc_number_dispensa = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_DetalhesOfertaCompra1_txtOC']").text 
                        details_table_oc_dispensa = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grd_fornecedor_lance']") 
                        rows_details_oc_dispensa = details_table_oc_dispensa.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grd_fornecedor_lance']/tbody/tr")[1:]
        
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

                                        browser_driver.back()
                                        k+=1
                                        time.sleep(5)
                                                    
    #if (m+1) % 15 == 0:
    #        time.sleep(20)

    browser_driver.close()
    browser_driver.quit()
    
    
    
    
                                                ######################################################################################
                                                ##Criando uma table (Excel) para visualizar os valores coletados da lista do Dispensa## 
                                                ######################################################################################
    
    df_final_table_ocs = pd.DataFrame(result_all_ocs)
    df_final_products_ocs = pd.DataFrame(details_products_dispensa) 
    
    
    #Criando as duas novas colunas para colocar o Valor Atual e o Valor Mínimo do Produto sendo disputado na Dispensa
    #df_final_products_ocs['Valor Inicial'] = 100.0
    #df_final_products_ocs['Valor Mínimo'] = 0.0

    #=ABS(H2*(G2/100)-H2)
    
    #Definindo a função lambda para calcular o valor mínimo de acordo com a Porcentagem Mínima de Redução e o Valor Atual
    #def min_value_func(row):
    #    return abs(row.iloc[7] * (row.iloc[6] / 100) - row.iloc[7])

    #df_final_products_ocs['Valor Mínimo'] = df_final_products_ocs.apply(min_value_func, axis=1)

    #Adicionando as colunas de valores iniciais e mínimos
    #df_final_products_ocs['Valor Inicial'] = df_final_products_ocs['Valor Inicial']
    #df_final_products_ocs['Valor Mínimo'] = df_final_products_ocs.apply(min_value_func, axis=1)

    
    writer = pd.ExcelWriter(path_with_filename, engine='openpyxl')
    writer2 = pd.ExcelWriter(path_with_filenameallocs, engine='openpyxl')
    
    df_final_products_ocs.to_excel(writer, sheet_name=DateSheet, header=True, index=False)
    df_final_table_ocs.to_excel(writer2, sheet_name=DateSheet, header=True, index=False)
    

    writer.close()
    writer2.close() 


    print(df_final_table_ocs)
    print(df_final_products_ocs)
    
    print('DataFrame is written to Excel File successfully!!!')
    
#bec_filterdispensa()