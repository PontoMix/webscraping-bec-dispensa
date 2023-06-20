#Pacote para fazer uma interface gráfica com Python
import tkinter as tk

#Importando o pacote para modificar as fontes dos widgets
from tkinter import font

import win32com.client as win32

#Importando as funções para rodar a Dispensa de acordo com o checkbox
from dispensaallocs import bec_allocsdispensa
from dispensaocsfilter import bec_filterdispensa


def category_productOCs(entry_number):
    field_value = entry_number.get()
    print(field_value)
    
    if field_value != "":
            run_filterdispensaocs(field_value) #Passando o valor do parâmetro para a próxima função usar    
    else:
            secondary_screen = tk.Tk()
            screen_width2 = root.winfo_screenwidth()
            screen_height2 = root.winfo_screenheight()

            width2 = 500
            height2 = 400
            x = (screen_width2/2) - (width2/2)
            y = (screen_height2/2) - (height2/2)
            secondary_screen.geometry("{}x{}+{}+{}".format(width2, height2, int(x), int(y)))

            secondary_screen.title("ERROR! - Campo não preenchido")

            content2 = tk.Frame(secondary_screen, background="#6818FF")
            content2.pack(fill=tk.BOTH, expand=True)

            font_backroot = font.Font(family="Arial", size=18, weight="bold")
            label = tk.Label(content2, text="POR FAVOR,\n COLOQUE NO PRIMEIRO CAMPO\nO NOME DO ITEM/CATEGORIA.\n\n E QUE SEJA UM ITEM POSSÍVEL DE PESQUISAR!\n\n EXEMPLO: COZINHA.", font=font_backroot, background="#6818FF", fg="snow")
            label.pack(side=tk.TOP, anchor=tk.CENTER, padx=5, pady=100)

            back_button = tk.Button(content2, bg="#DEA228", font=font_backroot, fg="snow", text="VOLTAR", relief="raised", borderwidth=5, command=secondary_screen.destroy)
            back_button.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)

            secondary_screen.mainloop()
            
            
            
#Função para pegar todos os valores da Tabela da Dispensa de acordo com a categoria/item digitado e os dados de cada OC presentes nas tabela de detalhes dos produtos            
def run_filterdispensaocs(field_value):
        
        bec_filterdispensa(field_value)
        
        finished_screen2 = tk.Tk()
        screen_width_finished2 = finished_screen2.winfo_screenwidth()
        screen_height_finished2 = finished_screen2.winfo_screenheight()

        width_finished2 = 500
        height_finished2 = 400

        x = (screen_width_finished2/2) - (width_finished2/2)
        y = (screen_height_finished2/2) - (height_finished2/2)

        finished_screen2.geometry("{}x{}+{}+{}".format(width_finished2, height_finished2, int(x), int(y)))
        finished_screen2.title("SUCESSO! - Scraping finalizado com sucesso")

        finished_content2 = tk.Frame(finished_screen2, background="#6818FF")
        finished_content2.pack(fill=tk.BOTH, expand=True)

        font_backroot2 = font.Font(family="Arial", size=20, weight="bold")
        label_finished = tk.Label(master=finished_content2, text="FINALIZADO COM SUCESSO! \n O SCRAPING DE TODOS OS CONVITES \n FORA EXECUTADO COM ÊXITO!", font=font_backroot2, background="#6818FF", fg="snow")
        label_finished.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=70)

        back_button = tk.Button(master=finished_content2, text="VOLTAR", font=font_backroot2, bg="#DEA228", fg="snow", relief="raised", borderwidth=5, command=finished_screen2.destroy)
        back_button.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)

        finished_screen2.mainloop()
        


#Função para pegar todos os valores da Tabela da Dispensa e os dados de cada OC presente na tabela de detalhes dos produtos          
def run_alldispensaocs():
        
        bec_allocsdispensa() 
        
        finished_screen2 = tk.Tk()
        screen_width_finished2 = finished_screen2.winfo_screenwidth()
        screen_height_finished2 = finished_screen2.winfo_screenheight()

        width_finished2 = 500
        height_finished2 = 400

        x = (screen_width_finished2/2) - (width_finished2/2)
        y = (screen_height_finished2/2) - (height_finished2/2)

        finished_screen2.geometry("{}x{}+{}+{}".format(width_finished2, height_finished2, int(x), int(y)))
        finished_screen2.title("SUCESSO! - Scraping finalizado com sucesso")

        finished_content2 = tk.Frame(finished_screen2, background="#6818FF")
        finished_content2.pack(fill=tk.BOTH, expand=True)

        font_backroot2 = font.Font(family="Arial", size=20, weight="bold")
        label_finished = tk.Label(master=finished_content2, text="FINALIZADO COM SUCESSO! \n O SCRAPING DE TODOS OS CONVITES \n FORA EXECUTADO COM ÊXITO!", font=font_backroot2, background="#6818FF", fg="snow")
        label_finished.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=70)

        back_button = tk.Button(master=finished_content2, text="VOLTAR", font=font_backroot2, bg="#DEA228", fg="snow", relief="raised", borderwidth=5, command=finished_screen2.destroy)
        back_button.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)

        finished_screen2.mainloop()           
                                                                    ######################################
                                                                    ###CRIAÇÃO DE UMA INTERFACE GRÁFICA###
                                                                    ######################################
                                                                    

#Criação da Janela onde estará todos os Widgets (controles, elementos, janelas,...) - É o elemento Pai da hierarquia de widgets
root = tk.Tk()

#Pegando a Largura e Altura da tela do monitor
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

#Definindo os valores para a Largura e Altura da janela
width = 600
height = 450

#Calculando as coordenadas de X e Y para o centro da tela com base no tamanho do monitor e da janela do aplicativo
x = (screen_width/2) - (width/2)
y = (screen_height/2) - (height/2)

#Definindo/Formatando os valores geométricos (tamanhos) da janela 
root.geometry("{}x{}+{}+{}".format(width, height, int(x), int(y)))
root.title("WebScrapping - Dispensa")

#Criação do elemento filho do root
content = tk.Frame(root, background="#6818FF")
content.pack(fill=tk.BOTH, expand=True)

#Criação de uma padronização das fontes dos widgets
font_all = font.Font(family="Arial", size=18, weight="bold")

#Criação dos próximos elementos da hierarquia

#Criação do campo de entrada (widget)
label = tk.Label(master = content, text="Digite o nome do item ou categoria:", font=font_all, background="#6818FF", fg="snow")
label.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)

entry_number = tk.Entry(master=content)
entry_number.pack(side=tk.TOP, anchor=tk.CENTER, padx=20, pady=5)
entry_number.focus()

#Criação do botão que irá invocar a função para rodar a função bec_convite()
button_onlyxtoy = tk.Button(master = content, text="Scraping", font=font_all, command=lambda:category_productOCs(entry_number), bg="#DEA228", fg="snow", relief="raised", borderwidth=5)
button_onlyxtoy.pack(side=tk.TOP, anchor=tk.CENTER, padx=20, pady=20)

#Botão para fazer o Scraping de todas as OCs
button_all = tk.Button(master = content, text="Scraping de todas OCs", font=font_all, command=run_alldispensaocs, bg="#DEA228", fg="snow", relief="raised", borderwidth=5)
button_all.pack(side=tk.TOP, anchor=tk.CENTER, padx=20, pady=20)


root.mainloop()