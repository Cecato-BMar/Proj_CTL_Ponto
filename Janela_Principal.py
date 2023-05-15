#Importando bibliotecas
import customtkinter as ctk    # customctkinter
from datetime import datetime, timedelta # datetime
import openpyxl                #openpyxl  manipulação de excel 
from tkinter import messagebox #tkinter

#criando um arquivo excell 
workbook = openpyxl.Workbook()
workbook = openpyxl.load_workbook('Controle de Ponto.xlsx')

# Selecionando uma página
pag = workbook.active

# Criando as colunas
pag['A1'] = 'Dia Entrada'
pag['B1'] = 'Hora entrada'
pag['C1'] = 'Dia saíada'
pag['D1'] = 'Hora Saída'
pag['E1'] = 'Horas Trabalhadas' 

#1- Pegar a data e hora Criando variáveis
d_ent = []
h_ent = []
d_saida = []
h_saida = []
f_data = '%d /%m/ %Y'
f_hora = '%H:%M:%S'
exib_h_ent = datetime
exib_h_saida = datetime
horas_trab = []

#Criar a janela principal
janela = ctk.CTk()
janela.geometry("500x400")
janela.title("**BMAR**")

#Customizar janela 
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

# Design da janela
texto = ctk.CTkLabel(janela, text= "Controle de Ponto**BMAR**", font=("Arial", 20))
texto.pack(padx=10, pady=10)
# Criar frames para entrada e saída
frame_entrada = ctk.CTkFrame(master=janela)
frame_entrada.pack(pady=20)
frame_saida = ctk.CTkFrame(master=janela)
frame_saida.pack(pady=20)
frame_gravar = ctk.CTkFrame(master=janela)
frame_gravar.pack(pady=20)

# Criar as variáveis para as últimas datas
ultima_d_ent = ctk.StringVar()
ultima_d_saida = ctk.StringVar()
ultima_h_ent = ctk.StringVar()
ultima_h_saida = ctk.StringVar()
ht = ctk.StringVar()
saldo_h = ctk.StringVar()
#Função para pegar a hora atual
def pegar_hora():
    dt = datetime.now()
    data = (dt.strftime(f_data))
    hora = (dt.strftime(f_hora)) 
    return [data, hora]

# Função para pegar a entrada
def entrada():    
    exib_d_ent = pegar_hora()[0]
    exib_h_ent = pegar_hora()[1]
    print("Dia entrada:" + exib_d_ent + "--- Hora entrada:"+ exib_h_ent)
    ultima_d_ent.set(exib_d_ent) # atualizar a variável com a última data
    label_ent.configure(textvariable=ultima_d_ent) # atualizar o rótulo com a última data
    ultima_h_ent.set(exib_h_ent) # atualizar a variável com a última data
    label_ent.configure(textvariable=ultima_h_ent) # atualizar o rótulo com a última data
    d_ent.append(exib_d_ent)    
    h_ent.append(exib_h_ent)
    
    
    # Função para pegar a saída
def bsaida():    
    exib_d_saida = pegar_hora()[0]
    exib_h_saida = pegar_hora()[1]
    print("Dia saída:" + exib_d_saida + "--- Hora saída:"+ exib_h_saida)
    ultima_d_saida.set(exib_d_saida) # atualizar a variável com a última data
    label_saida.configure(textvariable=ultima_d_saida) # atualizar o rótulo com a última data
    ultima_h_saida.set(exib_h_saida) # atualizar a variável com a última data
    label_saida.configure(textvariable=ultima_h_saida) # atualizar o rótulo com a última data 
    d_saida.append(exib_d_saida)
    h_saida.append(exib_h_saida)  
    
def horas_trabalhadas():    
    for i in range(len(h_ent)):        
        horas =  datetime.strptime(h_saida[i], f_hora)-datetime.strptime(h_ent[i], f_hora)        
        horas_trab.append(horas)   
    total_horas = sum(horas_trab, timedelta())               
    print(horas_trab)
    print(total_horas)
    ht.set(total_horas) # atualizar a variável com a última data
    label_htrabalhadas.configure(textvariable=ht) # atualizar o rótulo com a última data
    ht.set(total_horas) # atualizar a variável com a última data 
    
    
def gravar(): 
    horas_trabalhadas()   
    for i in range(len(d_ent)):        
        dados = (d_ent[i],h_ent[i],d_saida[i], h_saida[i], horas_trab[i])       
        pag.append(dados)
        #print(dados)
    workbook.save('Controle de Ponto.xlsx')             
    messagebox.showinfo("Dados Salvos", "Os dados foram salvos com sucesso!")            
    print(total_horas())
    
    
def total_horas():
    total_timedelta = timedelta()

    for row in pag.iter_rows(min_row=2, values_only=True):
        valor_celula = row[5 - 1]  # Subtrai 1 porque as colunas no Excel são indexadas a partir de 1
        if isinstance(valor_celula, timedelta):
            total_timedelta += valor_celula
    saldo_h.set(total_timedelta) # atualizar a variável com a última data
    label_totais.configure(textvariable=saldo_h) # atualizar o rótulo com a última data
    saldo_h.set(total_timedelta) # atualizar a variável com a última data 
    return total_timedelta 
     


    
# Botão de entrada e label
btn_entrada = ctk.CTkButton(master=frame_entrada, text="Entrada", command=entrada)
btn_entrada.pack(side="left", padx=5)
label_ent = ctk.CTkLabel(master=frame_entrada, text="***ENTRADA***", fg_color="green", corner_radius=20)
label_ent.pack(side="left", padx=15)
# Botão de saída e label
btn_saida = ctk.CTkButton(master=frame_saida, text="Saída", command=bsaida)
btn_saida.pack(side="left", padx=5)
label_saida = ctk.CTkLabel(master=frame_saida, text="****SAÍDA****", fg_color="red", corner_radius=20)
label_saida.pack(side="left", padx=15)
#Criar um botão para salvar
btn_gravar = ctk.CTkButton(master=frame_gravar, text="Gravar", command=gravar)
btn_gravar.pack(pady=30 )        
# Label horas trabalhadas no dia
label_htrabalhadas = ctk.CTkLabel(master=frame_gravar, text="****Horas Trabalhadas****", fg_color="red", corner_radius=20)
label_htrabalhadas.pack(side="left", padx=15, pady=10)
# Label horas total
label_totais = ctk.CTkLabel(master=frame_gravar, text="******Horas Totais******", fg_color="teal", corner_radius=20)
label_totais.pack(side="right", padx=15, pady=10)


janela.mainloop()