# Bot-CNES
Automação de inserção de Marcações no sistema do Cadastro Nacional de Estabelecimentos de Saúde (CNES)

Essa automação permite a inserção de habilitações e incentivos na funcionalidade online específica do CNES a partir do preenchimento da planilha Marcacoes.xlsx.

Após conclusão da inserção das marcações contidas na planilha, é criada uma pasta chamada RESULTADO, com 2 arquivos TXT para conferência da inserção, um para habilitação e outro para incentivo, 
com informação linha a linha da inserção, se foi bem sucedida ou teve um erro, para que seja verificado manualmente em caso de erro.

Deve-se executar o script CNES-Bot.py, conferir se a planilha Marcacoes.xlsx está com os dados corretos, depois informar as credenciais de acesso e clicar em Iniciar inserção das marcações.
Executará o arquivo Automação-Marcações.py, que por sua vez, executa, sequencialmente, os arquivos CNES-Automatização-Incentivos.py e CNES-Automatização-Habilitações.py.

Após execução, será exibida em tela a informação para verificar o RESULTADO.
Para isso, basta clicar no botão Conferir RESULTADO, onde será exibida a pasta com os 2 arquivos TXT.
