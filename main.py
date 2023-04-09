# Projeto edstinado a controlar o ponto de um funcionário no trabaljo.
from datetime import datetime
import openpyxl


#criando um arquivo excell 
workbook = openpyxl.Workbook()
workbook = openpyxl.load_workbook('Controle de Ponto.xlsx')
# Selecionando uma página
pag = workbook.active
# Criando as colunas
pag['A1'] = 'Entrada'
pag['B1'] = 'Saída'
 #1- Pegar a data e hora
2

ent = []
saida = []
dt = datetime.now()
dt = (dt.strftime('%d /%m/ %Y %H:%M:%S '))
print(dt)


def painel():
    print("=" * 30)
    print("PROGRAMA CONTROLE DE PONTO")
    print("=" * 30)
    print("1- Entrada.")
    print("2- Saída.")
    print("3- Sair do programa.")
    opcao = int(input("Opção : "))
    ponto(opcao)
    print("=" * 30)
    print("-"*13 + "FIM" + "-"*13 )
    print("=" * 30)
    

#2- Atribuir a entrada ou saída
   
def ponto(opcao):
     while opcao != 3 :
            
        if opcao == 1 :
            dt1 = datetime.now()
            ent.append(dt1.strftime('%d /%m/ %Y %H:%M:%S ')) 
            
            opcao = int(input("Opção : "))
            
     
        elif opcao == 2 :
            dt2 = datetime.now()
            saida.append(dt2.strftime('%d /%m/ %Y %H:%M:%S '))
            
            opcao = int(input("Opção : "))
           
            
       
     print(f"Horários de entrada: {ent} "   )
     print(f"Horários de saida:   {saida} "   )
     print("Saindo...")
     for i in range(len(ent)):
        dados = (ent[i],saida[i])
        pag.append(dados)
        
    
     workbook.save('Controle de Ponto.xlsx')
     
     
    
painel()



   



    





 
    





 #3- Salvar em um arquivo excell

