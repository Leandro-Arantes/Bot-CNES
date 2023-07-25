import tkinter as tk
from tkinter import filedialog
import subprocess
import os
import datetime

# Função para executar o script de Automação de Marcações
def executar_script():
    login = login_var.get()
    senha = senha_var.get()
    cpf = cpf_var.get()

    # Chamar o script 'Automação-Marcações.py' com os argumentos de login, senha e cpf
    script_path = r"C:\CNESBot\Automação-Marcações.py"
    subprocess.run(['python', script_path, login, senha, cpf])

# Função para abrir o arquivo Marcacoes.xlsx
def abrir_arquivo():
    arquivo_path = r"C:\CNESBot\Marcacoes.xlsx"
    os.startfile(arquivo_path)
    
def conferir_resultado():
    caminho_pasta = r"C:\CNESBot\RESULTADO"
    os.startfile(caminho_pasta)

def atualizar_hora():
    data_hora_atual = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    texto_label3.config(text=data_hora_atual)
    # Atualização a cada 1000ms(1 segundo)
    texto_label3.after(1000, atualizar_hora)

# Função para validar a entrada do login e limitar a 6 caracteres
def validar_login(char):
    return len(entry_login.get()) < 6 or char == ''

# Criar a janela principal
window = tk.Tk()
window.title("Executar Script de Automação de Marcações")
window.geometry("400x380")

# Variáveis para armazenar as informações inseridas pelo usuário
login_var = tk.StringVar()
senha_var = tk.StringVar()
cpf_var = tk.StringVar()

# Criar os rótulos e campos de entrada para login, senha e CPF
label_login = tk.Label(window, text="Login:")
label_login.pack(pady=1)
entry_login = tk.Entry(window, width=15, textvariable=login_var, validate="key", validatecommand=(window.register(validar_login), '%S'))
entry_login.pack(pady=5)

label_senha = tk.Label(window, text="Senha:")
label_senha.pack(pady=1)
entry_senha = tk.Entry(window, show='*', width=15, textvariable=senha_var)
entry_senha.pack(pady=5)

label_cpf = tk.Label(window, text="CPF:")
label_cpf.pack(pady=1)
entry_cpf = tk.Entry(window, show='*', width=15, textvariable=cpf_var)
entry_cpf.pack(pady=5)

# Criar o botão para abrir o arquivo Marcacoes.xlsx
btn_abrir_arquivo = tk.Button(window, text="Abrir arquivo Marcações para realizar alteração", command=abrir_arquivo)
btn_abrir_arquivo.pack(pady=10)

# Criar o botão para executar o script
btn_executar = tk.Button(window, text="Iniciar inserção das marcações", command=executar_script)
btn_executar.pack(pady=10)

# Criar botão para verificar os arquivos de HABILITAÇÕES e INCENTIVOS inseridos
btn_resultado = tk.Button(window, text="Conferir RESULTADO", command=conferir_resultado)
btn_resultado.pack(padx=5, pady=10)

texto_label = tk.Label(window, text="Verificar arquivo TXT mais recente para conferir a Habilitação e Incentivo.")
texto_label.pack(pady=2)

texto_label2 = tk.Label(window, text="Adesão a Programas e Projetos ainda em desenvolvimento")
texto_label2.pack(pady=2)

# Criar o Label para exibir a data e hora atual
texto_label3 = tk.Label(window, text="")
texto_label3.pack(pady=2)

# Chamar a função para iniciar a atualização da hora
atualizar_hora()

# Iniciar o loop principal da aplicação
window.mainloop()

