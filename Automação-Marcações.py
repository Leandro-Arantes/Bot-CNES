# Execução de funcionalidades que envolvem Sistema Operacional
import os as os
import subprocess
import sys
import logging

# Receber dados de login
def automatizar_marcações(login, senha, cpf):
    # Aqui você pode utilizar os valores de login, senha e cpf recebidos como argumentos para realizar a automação das marcações
    
    logging.info(f'login: {login}')
    logging.info(f'senha: {senha}')
    logging.info(f'cpf: {cpf}')

if __name__ == "__main__":
    if len(sys.argv) == 4:
        # O primeiro argumento é o próprio nome do script (ignoramos)
        login = sys.argv[1]
        senha = sys.argv[2]
        cpf = sys.argv[3]
        automatizar_marcações(login, senha, cpf)
    else:
        print("Uso: python Automação-Marcações.py <login> <senha> <cpf>")

# Definição do script de incentivo
def iniciar_script_inc():
    script_path = os.path.join(os.getcwd(), r'C:\CNESBot\CNES-Automatização-Incentivos.py')
    subprocess.run(['python', script_path, login, senha, cpf])

# Definição do script de habilitação
def iniciar_script_hab():
    script_path = os.path.join(os.getcwd(), r'C:\CNESBot\CNES-Automatização-Habilitações.py')
    subprocess.run(['python', script_path, login, senha, cpf])

# Executa o script de incentivos e aguarda a conclusão
print('Executando script de Incentivos...')
iniciar_script_inc()

# Executa o script de habilitações somente após a conclusão do script de incentivos
print('Executando script de Habilitações...')
iniciar_script_hab()

print('Scripts executados. Conferir pasta RESULTADO.')