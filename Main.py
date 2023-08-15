from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkcalendar import DateEntry
import pyautogui
import pyodbc
import pandas as pd
from tkinter.filedialog import asksaveasfilename
from openpyxl import Workbook
from babel.dates import format_date, parse_date, get_day_names, get_month_names
from babel.numbers import *
from openpyxl import workbook
from openpyxl.styles import Font, Alignment
import threading 
from datetime import date, datetime, timedelta
import os
import tkinter.font as tkFont
from time import sleep
from cx_Freeze import setup, Executable

senha_admin="123456"

def pedir_senha1():
   

    
    senha_top = Toplevel(janela_principal)
    senha_top.geometry("340x230")
    texto= Label(senha_top,text="informe a senha de administrador")
    receber_senha = Entry(senha_top,show="*")
    receber_senha.place(x=100,y=100)   
    
    texto.place(x=80,y=50)
    def entrar_senha2():
        pegar_senha = receber_senha.get()
        if pegar_senha != senha_admin:
            pyautogui.alert("Senha errada!")
            senha_top.destroy()
        else:
            senha_top.destroy()
            def config_banco():
                
                #função de configuração do banco de dados
                def salvar_config_banco():
                    #função para salvar a configuração
                    def salvar_variavel_banco():
                    
                        instancia = instancia1.get("1.0", "end")
                        with open("instancia.txt","w") as arquivo:
                            arquivo.write(instancia)
                        banco = instanciabanco.get("1.0", "end")
                        with open("banco.txt","w")as arquivo:
                            arquivo.write(banco)

                    salvar_variavel_banco()
                    def carregar_variavel_banco():
                        with open("instancia.txt", "r")as arquivo:
                            instancia = arquivo.read()
                        instancia1.delete("1.0", "end")
                        instancia1.insert("1.0", instancia)
                        with open("banco.txt", "r")as arquivo:
                            banco = arquivo.read()
                        instanciabanco.delete("1.0", "end")
                        instanciabanco.insert("1.0", instancia)
                    carregar_variavel_banco
                    Janela_config_instanc.destroy()

                


                Janela_config_instanc = Toplevel(janela_principal)
                Janela_config_instanc.geometry("330x230")
                Label(Janela_config_instanc, text="Informe a instância").place(x=115, y=15)
                instancia1 = Text(Janela_config_instanc)
                instancia1.place(relx = 0.1, rely = 0.17,relwidth = 0.8,relheight=0.2)
                Label(Janela_config_instanc, text="Informe o banco").place(x=115, y=90)
                instanciabanco = Text(Janela_config_instanc)
                instanciabanco.place(relx = 0.1, rely = 0.57,relwidth = 0.8,relheight=0.2)
                botao_salvar = Button(Janela_config_instanc, text="Salvar", command=salvar_config_banco)
                botao_salvar.place(x=125, y=180)
                botao_salvar.config(padx=15, pady=5 )
            config_banco()
    botao_senha = Button(senha_top,text="Entrar", command=entrar_senha2).place(x=140,y=180)


def pedir_senha2():
    senha_top2 = Toplevel(janela_principal)
    senha_top2.geometry("340x230")
    texto= Label(senha_top2,text="informe a senha de administrador")
    receber_senha2 = Entry(senha_top2, show="*")
    receber_senha2.place(x=100,y=100)   
    texto.place(x=80,y=50)
    def entrar_senha2():
        pegar_senha = receber_senha2.get()
        if pegar_senha != senha_admin:
            pyautogui.alert("Senha errada!")
            senha_top2.destroy()
        else:
            senha_top2.destroy()
            def config_script():

                #janela de configuração do script


                Janela_config_script = Toplevel(janela_principal)
                Janela_config_script.geometry("630x740")
                
                
                Label(Janela_config_script, text="Informe a primeira parte do Script").place(x=240, y=10)
                instancia = Text(Janela_config_script)
                instancia.place(relx = 0.1, rely = 0.040,relwidth = 0.8,relheight=0.90)
                
                
                
                
                rodar_script1 = Text(Janela_config_script)
            
               

                
                
                def salvar_script():
                    def salvar_variavel():
                        
                        coluna = instancia.get("1.0", "end")
                        texto = rodar_script1.get("1.0", "end")
                        
                        
                        with open("Script1.sql", "w", encoding="utf-8") as arquivo:
                            arquivo.write(coluna)
                        
                        
                        
                    salvar_variavel()


                   
                    Janela_config_script.destroy()

                
                
                
                botao_salvar = Button(Janela_config_script, text="Salvar", command=salvar_script)
                botao_salvar.place(x=280, y=700)
                botao_salvar.config(padx=15, pady=5 )
            config_script()
    botao_senha = Button(senha_top2,text="Entrar", command=entrar_senha2).place(x=140,y=180)


    











janela_principal = Tk()
janela_principal.title("Gerador de relatório")
fontsize=tkFont.Font(family="Popins", size=20)
janela_principal.geometry("430x240")
janela_principal.maxsize(430,240)
janela_principal.iconbitmap(r"relatorio-de-negocios.ico")
descricao = Label(text="Data Inicial", background="#c80028", foreground='white')
gerador_Relatorio = Label(text="Gerador de relatório", background="#c80028", foreground='white', font=fontsize)
gerador_Relatorio.place(x=110, y=7)





janela_principal.configure(bg="#c80028")
descricao.place(x=70,y=120)
cal_entrada = DateEntry(janela_principal, selectmode='day', year=2023, month=1,day=1)
cal_entrada.place(x=70,y=150)

def gerar_relatorio():
    with open("instancia.txt", "r") as arquivo:
        instancia_conectar = arquivo.read()
        
    with open("banco.txt",'r') as arquivo:
        banco_conectar = arquivo.read()
    def controledata():
        data_entrada_selecionada = cal_entrada.get_date()
        data_saida_selecionada = cal_final.get_date()
        
        d1 = datetime.strptime(str(data_entrada_selecionada), '%Y-%m-%d').date()
        d2 = datetime.strptime(str(data_saida_selecionada), '%Y-%m-%d').date()
        hoje = datetime.now().date()
      

        if d2>hoje:
            pyautogui.alert("A data final não pode ser superior a hoje!")
        else:

            if data_entrada_selecionada > data_saida_selecionada:
                pyautogui.alert("A data inicial não pode ser superior à data final")
                        
        
            else:
                soma_dias = abs((d1 - d2).days)
                print(soma_dias)
        
                if soma_dias > 365:
                    pyautogui.alert("Intervalo não pode ser superior a 12 meses (365 dias)")
            
                
                else:
                    with open("datainicial.txt", "w") as arquivo:
                        arquivo.write(f"{data_entrada_selecionada}T00:00:00")
                    with open("datainicial.txt", "r") as arquivo:
                        texto_com_quebra = arquivo.read()
                        texto_sem_quebra = texto_com_quebra.replace('\n', '')

                    with open("datafinal.txt", "w") as arquivo:
                        arquivo.write(f"{data_saida_selecionada}T23:59:00")
                    
                    

                    
                    
                    
                    
                    

                    def identificar_datas():
                        with open("datainicial.txt","r") as arquivo:
                            nova_data_inicio = arquivo.read()
                        with open("datafinal.txt","r") as arquivo:
                            nova_data_fim = arquivo.read()
                        
                        nome_arquivo_sql = 'Script1.sql'
                        # Abra o arquivo SQL e leia o conteúdo
                        with open(nome_arquivo_sql, "r", encoding="utf-8-sig") as arquivo:
                            conteudo = arquivo.read()

                        # Use regex para encontrar todas as ocorrências de datas no formato "YYYY-MM-DD" ou "YYYY-MM-DDTHH:mm:ss"
                        padrao_data = r'\b\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}:\d{2})?\b'
                        datas_encontradas = re.findall(padrao_data, conteudo)

                        # Substituir as datas encontradas pelo novo intervalo em pares
                        novo_conteudo = conteudo
                        for i in range(0, len(datas_encontradas), 2):
                            if i + 1 < len(datas_encontradas):
                                data_inicio = datas_encontradas[i]
                                data_fim = datas_encontradas[i + 1]
                                novo_conteudo = novo_conteudo.replace(data_inicio, nova_data_inicio).replace(data_fim, nova_data_fim)

                        # Sobrescrever o arquivo SQL com o novo conteúdo
                        with open(nome_arquivo_sql, "w", encoding="utf-8-sig") as arquivo:
                            arquivo.write(novo_conteudo)

                        print(f"As datas foram alteradas no arquivo '{nome_arquivo_sql}'.")
                    identificar_datas()

                
                    root = Tk()
                    root.withdraw()
                    file_path = asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])
                    root.destroy()

                    if not file_path:
                        pyautogui.alert("Nenhum caminho selecionado")
                    else:
                       
                        # Criar uma nova planilha Excel
                        workbook = Workbook()

                        # Obter a planilha ativa (por padrão, é criada uma chamada "Sheet")
                        sheet = workbook.active


                        with open("Script1.sql", "r", encoding="utf-8-sig") as arquivo:
                            comandosql = arquivo.read().replace('GO','')

                        with open("instancia.txt","r") as arquivo:
                                arm_serv = arquivo.read().replace('\n','')
                        with open("banco.txt", "r")as arquivo:
                            banco = arquivo.read().replace('\n','')

                        SERVER=arm_serv
                        database =banco
                        id ='LOGIN'
                        password='SENHA'
                        try:
                            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+SERVER+';DATABASE='+database+';UID='+id+';PWD='+password+';Encrypt=no')
                        except:
                            pyautogui.alert("Não foi possível conectar ao banco de dados")
                        cursor = cnxn.cursor()
                        

                        cursor.execute(comandosql)
                        def alerta_relatorio():
                            pyautogui.alert("Gerando relatório aguarde até 2 minutos")
                        thread = threading.Thread(target=alerta_relatorio)
                        thread.start()
                        # Obtém o resultado da consulta
                        try:
                            resultado = cursor.fetchall()
                            nomes_colunas = [column[0] for column in cursor.description]
                            dados = {coluna: [] for coluna in nomes_colunas}

                            # # Preenche o dicionário com os dados da consulta
                            for linha in resultado:
                                for coluna, valor in zip(nomes_colunas, linha):
                                    dados[coluna].append(valor)

                    



                            df = pd.DataFrame.from_dict(dados, orient='index').transpose()
                                
                                
                    

                        
                        
                            
                            caminho_planilha = file_path
                            df.to_excel(caminho_planilha, index=False)

                            cursor.close()
                            
                            cnxn.close()
                        
                            pyautogui.alert(f"Relatório do dia {data_entrada_selecionada} até {data_saida_selecionada} gerado!") 
                        except:
                            pyautogui.alert("informação não localizada no banco de dados")
                        
                        
                        
    controledata()


                

    

   
    




    
    

     





descricao2 = Label(text="Data Final", background="#c80028", foreground='white')
descricao2.place(x=270,y=120)

cal_final = DateEntry(janela_principal, selectmode='day', year=2023, month=1,day=1)
cal_final.place(x=270,y=150)





def janela_superior():
    nova_janela = Toplevel(janela_principal)
    nova_janela.iconbitmap(r"relatorio-de-negocios.ico")
    nova_janela.geometry("390x230")
    nova_janela.maxsize(390,230)
    descricao_auto = Label(nova_janela, text="agendamento para gerar relatório de maneira automatica,")
    descricao_auto2= Label(nova_janela, text=" informe em horas de quanto em quanto tempo gerar o relatorio")
    var1 = tk.IntVar()
   
    def checkbox_mark():
        if (var1.get() ==0):
            print("Não enviar E-mail")
        else:
            print("enviar E-mail")
    descricao_auto.place(x=25,y=30)
    checkbox = tk.Checkbutton(nova_janela, text='Enviar E-mail', variable=var1,onvalue=1, offvalue=0,command=checkbox_mark)
    checkbox.pack()
    descricao_auto2.place(x=5,y=50)
    hora_entrada = Entry(nova_janela)
    hora_entrada.place(x=15, y=90)
    def agendar():
        def timer():
            
            hora = int(campo_input2)
            contador = 3600 * hora
            while True:
                nova_janela.destroy()
                now = datetime.now()
                data_atual = now.strftime("%H:%M:%S")
                data_atual += ':' + '00'
                print(data_atual)
                sleep(contador)
                gerar_relatorio()
        
        hora = hora_entrada.get()
        if not hora.isnumeric():
            pyautogui.alert("Digite um número válido")
        else:

            if hora == "0":
                pyautogui.alert("A hora deve ser superior a 0")
            else:
                campo_input2 = hora
                print(campo_input2)
                timer()
    def nova_janela_email():
        janela_secundaria = Toplevel(nova_janela)
        janela_secundaria.geometry("550x430")
        janela_secundaria.maxsize(550,430)
        la_descricao = Label(janela_secundaria, text="Informe o Host")
        la_descricao.place(x=15,y=5)
        
        la_porta = Label(janela_secundaria, text="Informe a Porta")
        la_porta.place(x=200,y=5)
        portavar = Entry(janela_secundaria)
        portavar.place(x=200, y=25)

        smtpvar = Entry(janela_secundaria)
        smtpvar.place(x=15, y=25)



                
    thread_nova  = threading.Thread(target=agendar)
   
    bota_salvar = Button(nova_janela,text="Agendar", command=thread_nova.start)
  
    configurar_Email = Button(nova_janela, text="Configurar E-mail", command=nova_janela_email)
    configurar_Email.place(x=115, y=120)
    bota_salvar.place(x=15, y=120)





botao_banco_de_dados = Button(janela_principal,text="Instância", command=pedir_senha1)
botao_banco_de_dados.place(x=70,y=59)
botao_banco_de_dados.config(padx=30, pady=5)
Script = Button(janela_principal,text="Script", command=pedir_senha2)
botao_gerar = Button(janela_principal, text="Gerar Relatorio", command=gerar_relatorio)

botao_gerar.place(x=170,y=190)
agendamento_relatorio = Button(janela_principal, text="Relatorio automático", command=janela_superior)
agendamento_relatorio.place(x=270, y=190)
agendamento_relatorio.configure(padx=5, pady=5)
Script.configure(padx=40, pady=5)
Script.place(x=270,y=59)


janela_principal.mainloop()
