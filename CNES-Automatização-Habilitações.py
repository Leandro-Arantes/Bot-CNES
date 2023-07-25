# Importações para leitura do arquivo Excel
import pandas as pd

# Importações para automatizar a web
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert as Alert
import time
import requests as requests
from selenium.webdriver.chrome.options import Options
import argparse as argparse
import os as os
import numpy as np
from IPython.display import display
import datetime
from selenium.common.exceptions import NoSuchElementException
import sys
import logging

def automatizar_habilitacoes(login, senha, cpf):
    # Aqui você pode utilizar os valores de login, senha e cpf recebidos como argumentos para automatizar as habilitações
    logging.info(f'login: {login}')
    logging.info(f'senha: {senha}')
    logging.info(f'cpf: {cpf}')

if __name__ == "__main__":
    if len(sys.argv) == 4:
        # O primeiro argumento é o próprio nome do script (ignoramos)
        login = sys.argv[1]
        senha = sys.argv[2]
        cpf = sys.argv[3]
        automatizar_habilitacoes(login, senha, cpf)
    else:
        print("Uso: python CNES-Automatização-Habilitações.py <login> <senha> <cpf>")

# Carregar o arquivo de Marcações
marcacoes = r'C:\CNESBot\Marcacoes.xlsx'
df = pd.read_excel(marcacoes)
df['CNES'] = df['CNES'].astype(str).str.zfill(7)
display(df)

# Carregar o arquivo de Cod/Descricao das Marcações
cod_marcacoes = r'C:\CNESBot\Cod_Inc_Hab.xlsx'
df1 = pd.read_excel(cod_marcacoes)
display(df1)

# Converter a coluna "codigo" em ambos os DataFrames para o tipo objeto (string)
df['codigo'] = df['codigo'].astype(str).str.zfill(4)
df1['codigo'] = df1['codigo'].astype(str).str.zfill(4)


# Selecionar valores iguais da tabela e realizar o merge
resultado = pd.merge(df, df1, on='codigo', how='left')
display(resultado)

# Considerar somente linhas com CNES preenchido
count = resultado[resultado['CNES'].notnull()]['CNES'].count()
print(count)

# Configura as opções do ChromeDriver
chrome_options = Options()
chrome_options.add_argument('--headless')  # Executar em modo headless, sem abrir a janela do navegador

# Define o caminho para o ChromeDriver
driver_caminho = "chromedriver.exe"

# Configura e inicializa o ChromeDriver
navegador = Chrome(options=chrome_options)
#navegador = webdriver.Chrome('chromedriver.exe')

# Arquivo Marcações Inseridas
pasta_base = r'C:\CNESBot\RESULTADO'
# Obter a data e hora atual
data_hora_atual = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
data_hora_formatado = datetime.datetime.now().strftime("%d/%m/%Y às %H:%M:%S")
# Nome do arquivo com a data e hora atual
nome_arquivo = f"insercao_habilitacoes_{data_hora_atual}.txt"
# Caminho completo do arquivo
caminho = os.path.join(pasta_base, nome_arquivo)
largura = 150
largura_restante = 110

# Define o conteúdo de cada coluna
coluna1 = 'CNES'.center(8)
coluna2 = 'Código'.center(8)
coluna3 = 'Descrição'.center(largura_restante-1)
coluna4 = 'Status'.ljust(5)

# Verifica se a pasta base existe, caso contrário, cria-a
if not os.path.exists(pasta_base):
    os.makedirs(pasta_base)

# INICIO HABILITAÇÃO
#Página HABILITAÇÕES
habilitacao = ('http://cnes2.datasus.gov.br/Mod_Excluir_Habilitacao_Login.asp?Prim=1&Tipo_Habilitacao=H')
try: navegador.get(habilitacao)
except: navegador.get('http://cnes2.datasus.gov.br/Manutencao.htm')
finally: navegador.get(habilitacao)

#Identificador dos campos login HABILITAÇÕES
usuario = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr/td/form/table/tbody/tr[2]/td[2]/input')
senha = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr/td/form/table/tbody/tr[3]/td[2]/input')
cpf = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr/td/form/table/tbody/tr[4]/td[2]/input')


#Preenchimento de login
usuario.send_keys(f'{login}')
senha.send_keys(f'{senha}')
cpf.send_keys(f'{cpf}')
time.sleep(2)
navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr/td/form/table/tbody/tr[5]/td/input').click()

with open(caminho, 'w') as arquivo:
    arquivo.write('CONFERÊNCIA DAS MARCAÇÕES - HABILITAÇÕES'.center(largura, '-') + '\n')
    arquivo.write(f'Arquivo iniciado em: {data_hora_formatado}'.center(largura, '-') + '\n')
    arquivo.write('Identificação dos estabelecimentos e habilitações para conferência'.center(largura, '-') + '\n')
    arquivo.write('-' * largura + '\n')
    arquivo.write('{}|{}|{}| {}\n'.format(coluna1, coluna2, coluna3, coluna4))
    arquivo.write('-' * largura + '\n')

    # Busca CNES
    for index, row in resultado.iterrows():
        valor = str(row['CNES'])
        if valor == 'nan':
            break
        print('Iniciando a linha ' + str(index + 2) + ' CNES ' + str(row['CNES']) + ' Codigo ' + str(row['codigo']) + ' ' + str(row['acao']))
        if str(row['codigo'])[:1] == '8':
            print(f"Marcação {str(row['codigo'])} não é habilitação. Marcação para o CNES {str(row['CNES'])}, Código {str(row['codigo'])}, não inserida. Passando para próxima linha.")
            continue
        cnes = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/center[1]/form/table/tbody/tr[2]/td[4]/input')
        cnes.send_keys(row['CNES'])
        time.sleep(2)
        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/center/form/table/tbody/tr[2]/td[5]/input').click()
        time.sleep(2)

        # Acessar as Habilitações ativas da Unidade Pesquisada
        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr[1]/td[2]/font/a').click()
        cnes = str(row['CNES'])[:7]
        if cnes[6] == ".":
            cnes = '0' + cnes[:6]
        if cnes[5] == '.':
            cnes = '00' + cnes[:5]
        if cnes[4] == '.':
            cnes = '000' + cnes[:4]
        if cnes[3] == '.':
            cnes = '0000' + cnes[:3]
        if cnes[2] == '.':
            cnes = '00000' + cnes[:2]
        if cnes[1] == '.':
            cnes = '000000' + cnes[:1]
        codigo = str(row['codigo'])[:4]
        if codigo[3] == ".":
            codigo = "0" + codigo[:3]
        if codigo[2] == '.':
            codigo = '00' + codigo[:2]
        # Definir valor para a coluna descricao
        descricao = str(row['descricao'])
        acao = str(row['acao'])
        dt_portaria = str(row['data_portaria'])
        portaria = str(row['portaria'])
        comp_ini = str(row['data_inicial'])
        comp_fim = str(row['data_final'])
        # Verificar se o valor é NaN antes de converter para inteiro
        if pd.isnull(row['leitos']):
            leitos = ''
        else:
            leitos = str(int(row['leitos']))
        #Desabilitar
        if str(row['acao']) == 'desabilitar':
            try:
                navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form[2]/table/tbody/tr/td[2]/input').click()
                # Marcação
                elemento = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form[1]/table[2]/tbody')
                texto_elemento = elemento.text
                if codigo in texto_elemento:
                    elemento = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form[1]/table[2]/tbody')
                    texto_elemento = elemento.text
                    pesquisa_elemento = codigo in texto_elemento
                    print(texto_elemento)
                    linha_resultado = None
                    for indice, linha in enumerate(texto_elemento.split('\n'), 1):
                        if codigo in linha:
                            linha_resultado = indice
                            break

                    if linha_resultado is not None:
                        print("O termo", codigo, "foi encontrado na linha", linha_resultado)
                    else:
                        print("O termo", codigo, "não foi encontrado na tabela. Revisar manualmente.")
                        continue
                                
                    # Selecionar código de Habilitação
                    if linha_resultado == '0':
                        selecao = navegador.find_element(By.XPATH, "/html/body/table/tbody/tr/td/form/table[2]/tbody/tr/td[3]/font/input").click()
                        time.sleep(2)
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[3]/tbody/tr/td/input').click()
                    else:
                        selecao = navegador.find_element(By.XPATH, "/html/body/table/tbody/tr/td/form/table[2]/tbody/tr[" + str(linha_resultado) + "]/td[7]/font/input").click()
                        time.sleep(2)
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[3]/tbody/tr/td/input').click()
                            
                    # Preenchimento Competências
                    # Competência Final
                    if str(row['data_final']) is None:
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[1]').send_keys('00')
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[2]').send_keys('0000')
                    else:
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[1]').send_keys(str(row['data_final'])[:2])
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[2]').send_keys(str(row['data_final'])[-4:])
                            
                    # Portaria
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(Keys.CONTROL + "a")
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(Keys.BACKSPACE)
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(str(row['portaria'])) 
                            
                    # Data da Portaria
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(Keys.CONTROL + "a")
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(Keys.BACKSPACE)
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(str(row['data_portaria']))

                    # Confirmar habilitação
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[5]/td/input').click()

                    # Salvar no arquivo
                    arquivo.write('{:<8}|{:<8}| {:<{}}|{:27}\n'.format(str(row['CNES']).center(8), str(row['codigo']).center(8), descricao, largura_restante, 'Desabilitado'))
                    print(f"A ação de {str(row['acao'])} do CNES {str(row['CNES'])} foi realizada")
                                
                    try:
                        alerta = navegador.switch_to.alert
                        alerta.accept()
                                
                        # Nova Marcação
                        navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                        time.sleep(4)
                    except:
                        navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                        time.sleep(4)
            
            except Exception as e:
                # Caso ocorra uma exceção, imprima o erro e continue para a próxima linha
                arquivo.write('{:<8}|{:<8}| {:<{}}|{:27}\n'.format(str(row['CNES']).center(8), str(row['codigo']).center(8), descricao, largura_restante, 'Erro'))
                print(f"Erro na linha {index + 2}: {e}")
                continue
            
        if str(row['acao']) == 'habilitar':
            # Marcação
            time.sleep(1)
            try:
                elemento = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form[1]/table[2]/tbody')
                texto_elemento = elemento.text
                pesquisa_elemento = str(row['codigo']) in texto_elemento
                linha_resultado = None
                for indice, linha in enumerate(texto_elemento.split('\n'), 1):
                    if str(row['codigo']) in linha:
                        linha_resultado = indice
                        break
                # Alterar
                if linha_resultado is not None:
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form[2]/table/tbody/tr/td[3]/input').click()
                    elemento = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[2]/tbody')
                    texto_elemento = elemento.text
                    pesquisa_elemento = str(row['codigo']) in texto_elemento
                    linha_resultado = None
                    for indice, linha in enumerate(texto_elemento.split('\n'),1):
                        if str(row['codigo']) in linha:
                            linha_resultado = indice
                            break
                    if linha_resultado is not None:
                        print("O termo", str(row['codigo']), "foi encontrado na linha", linha_resultado)
                    else:
                        print("O termo", str(row['codigo']), "não foi encontrado na tabela.")
                    
                    # Seleção da marcação
                    if enumerate(texto_elemento.split('\n'),1) == '1':
                        navegador.find_element(By.XPATH, "/html/body/table/tbody/tr/td/form/table[2]/tbody/tr/td[3]/font/input").click()
                    else:
                        navegador.find_element(By.XPATH, "/html/body/table/tbody/tr/td/form/table[2]/tbody/tr["+ str(linha_resultado) +"]/td[3]/font/input").click()
                    
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[3]/tbody/tr/td/input').click() #Botão "Alterar habilitação" depois do código selecionado

                    # Competências
                    # Competência Final
                    if str(row['data_final']) is None:
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[1]').send_keys('00')
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[2]').send_keys('0000')
                    else:    
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[1]').send_keys(str(row['data_final'])[:2])
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[2]').send_keys(str(row['data_final'])[-4:])
                    
                    # Portaria
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(Keys.CONTROL + "a")
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(Keys.BACKSPACE)
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(str(row['portaria']))

                    # Data Portaria
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(Keys.CONTROL + "a")
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(Keys.BACKSPACE)
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(str(row['data_portaria']))
                    time.sleep(5)

                    # Leitos
                    if str(row['leitos']) == 'nan':
                        # Confirmar habilitação
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[5]/td/input').click()

                        try:
                            alerta = navegador.switch_to.alert
                            alerta.accept()
                            # Nova Marcação
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(4)
                        except:
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(4)
                    else:
                        #Inserir Leitos
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[7]/input').send_keys(Keys.CONTROL + "a")
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[7]/input').send_keys(Keys.BACKSPACE)
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[7]/input').send_keys(leitos)
                        time.sleep(2)

                        # Confirmar alteração
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[5]/td/input').click()

                        try:
                            alerta = navegador.switch_to.alert
                            alerta.accept()
                            # Nova Marcação
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(2)
                        except:
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(2) 
                    
                    # Salvar no arquivo
                    arquivo.write('{:<8}|{:<8}| {:<{}}|{:27}\n'.format(str(row['CNES']).center(8), str(row['codigo']).center(8), descricao, largura_restante, 'Alterado'))
                    print(f"A ação de {str(row['acao'])} do CNES {str(row['CNES'])} foi realizada")           
                
                else:
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form[2]/table/tbody/tr/td[1]/input').click() #Botão Habilitar
                    elemento = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[2]/tbody')
                    texto_elemento = elemento.text
                    pesquisa_elemento = str(row['codigo']) in texto_elemento
                    linha_resultado = None
                    for indice, linha in enumerate(texto_elemento.split('\n'),1):
                        if str(row['codigo']) in linha:
                            linha_resultado = indice
                            break
                    if linha_resultado is not None:
                        print("O termo", str(row['codigo']), "foi encontrado na linha", linha_resultado)
                    else:
                        print("O termo", str(row['codigo']), "não foi encontrado na tabela.")
                    
                    navegador.find_element(By.XPATH, "/html/body/table/tbody/tr/td/form/table[2]/tbody/tr["+ str(linha_resultado) +"]/td[3]/font/input").click() #selecionar a habilitação a ser inserida
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[3]/tbody/tr/td/input').click() #Botão "Habilitar" depois de selecionar a habilitação

                    # Competências
                    # Competência Inicial
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[3]/select[1]').send_keys(str(row['data_inicial'])[:2])
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[3]/select[2]').send_keys(str(row['data_inicial'])[-4:])

                    # Competência Final
                    if str(row['data_final']) is None:
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[1]').send_keys('00')
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[2]').send_keys('0000')
                    else:
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[1]').send_keys(str(row['data_final'])[:2])
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[2]').send_keys(str(row['data_final'])[-4:])
                    
                    # Portaria
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(Keys.CONTROL + "a")
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(Keys.BACKSPACE)
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(str(row['portaria']))

                    #Data Portaria
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(Keys.CONTROL + "a")
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(Keys.BACKSPACE)
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(str(row['data_portaria']))
                    time.sleep(2)

                    # Leitos
                    if str(row['leitos']) == 'nan':
                        # Desmarcar "Pendência" 
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[7]/input').click()
                        time.sleep(2)
                        
                        # Confirmar habilitação
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[5]/td/input').click()

                        try:
                            alerta = navegador.switch_to.alert
                            alerta.accept()
                            # Nova Marcação
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(2)
                        except:
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(2)
                    else:
                        #Inserir Leitos
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[7]/input').send_keys(Keys.CONTROL + "a")
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[7]/input').send_keys(Keys.BACKSPACE)
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[7]/input').send_keys(leitos)
                        
                        # Desmarcar "Pendência" 
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[8]/input').click()
                        time.sleep(2)
                        
                        try:
                            alerta = navegador.switch_to.alert
                            alerta.accept()
                            # Nova Marcação
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(4)
                        except:
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(4)
                    
                    # Salvar no arquivo
                    arquivo.write('{:<8}|{:<8}| {:<{}}|{:27}\n'.format(str(row['CNES']).center(8), str(row['codigo']).center(8), descricao, largura_restante, 'Habilitado'))
                    print(f"A ação de {str(row['acao'])} do CNES {str(row['CNES'])} foi realizada")
                    
            except NoSuchElementException:
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form[2]/table/tbody/tr/td[1]/input').click() #Botão Habilitar
                    elemento = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[2]/tbody')
                    texto_elemento = elemento.text
                    pesquisa_elemento = str(row['codigo']) in texto_elemento
                    linha_resultado = None
                    for indice, linha in enumerate(texto_elemento.split('\n'),1):
                        if str(row['codigo']) in linha:
                            linha_resultado = indice
                            break
                    if linha_resultado is not None:
                        print("O termo", str(row['codigo']), "foi encontrado na linha", linha_resultado)
                    else:
                        print("O termo", str(row['codigo']), "não foi encontrado na tabela.")
                    
                    navegador.find_element(By.XPATH, "/html/body/table/tbody/tr/td/form/table[2]/tbody/tr["+ str(linha_resultado) +"]/td[3]/font/input").click() #selecionar a habilitação a ser inserida
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[3]/tbody/tr/td/input').click() #Botão "Habilitar" depois de selecionar a habilitação

                    # Competências
                    # Competência Inicial
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[3]/select[1]').send_keys(str(row['data_inicial'])[:2])
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[3]/select[2]').send_keys(str(row['data_inicial'])[-4:])

                    # Competência Final
                    if str(row['data_final']) is None:
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[1]').send_keys('00')
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[2]').send_keys('0000')
                    else:
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[1]').send_keys(str(row['data_final'])[:2])
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[4]/select[2]').send_keys(str(row['data_final'])[-4:])
                    
                    # Portaria
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(Keys.CONTROL + "a")
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(Keys.BACKSPACE)
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[5]/input').send_keys(str(row['portaria']))

                    #Data Portaria
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(Keys.CONTROL + "a")
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(Keys.BACKSPACE)
                    navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[6]/input').send_keys(str(row['data_portaria']))
                    time.sleep(3)

                    # Leitos
                    if str(row['leitos']) == 'nan':
                        # Desmarcar "Pendência" 
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[7]/input').click()
                        time.sleep(3)
                        
                        # Confirmar habilitação
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[5]/td/input').click()

                        try:
                            alerta = navegador.switch_to.alert
                            alerta.accept()
                            # Nova Marcação
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(3)
                        except:
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(3)
                    else:
                        #Inserir Leitos
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[7]/input').send_keys(Keys.CONTROL + "a")
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[7]/input').send_keys(Keys.BACKSPACE)
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[7]/input').send_keys(leitos)
                        
                        # Desmarcar "Pendência" 
                        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table[2]/tbody/tr[4]/td[8]/input').click()
                        time.sleep(2)
                        
                        try:
                            alerta = navegador.switch_to.alert
                            alerta.accept()
                            # Nova Marcação
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(3)
                        except:
                            navegador.get('http://cnes2.datasus.gov.br/Lista_Es_Nome_habilitacao.asp')
                            time.sleep(3)
                    
                    # Salvar no arquivo
                    arquivo.write('{:<8}|{:<8}| {:<{}}|{:27}\n'.format(str(row['CNES']).center(8), str(row['codigo']).center(8), descricao, largura_restante, 'Habilitado'))
                    print(f"A ação de {str(row['acao'])} do CNES {str(row['CNES'])} foi realizada")
            
            except Exception as erro:
                # Caso ocorra uma exceção, imprima o erro e continue para a próxima linha
                arquivo.write('{:<8}|{:<8}| {:<{}}|{:27}\n'.format(str(row['CNES']).center(8), str(row['codigo']).center(8), descricao, largura_restante, 'Erro'))
                print(f"Erro na linha {index + 2}: {erro}")
                continue
            
    arquivo.write('-' * largura + '\n')
    arquivo.write(' ' * largura + '\n')
    arquivo.write('-' * largura + '\n')
    arquivo.write('Todas as habilitações contidas do arquivo Excel foram verificadas' + '\n')
#FIM HABILITAÇÃO
print('Todas as habilitações contidas do arquivo Excel foram inseridas')