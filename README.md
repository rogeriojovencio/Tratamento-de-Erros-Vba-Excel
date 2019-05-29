# Tratamento de Erros Vba Excel

Objetivo:
Tendo em vista produtividade e  na construção de sistemas vba, pelo programador, foi criado um modulo e uma classe, com o fim de agilizar a localização e interceptação de erros, otimizando a mensgem para o usuario do sistema, e auxiliando o programador na captura dos possiveis erros que poderão ocorrer no  sistema.
 -  Modulos:  mdlError.bas
 - Classes: clsFormatError.cls

#Escopo:
1:  Primeiros Passos:
•	Crie uma aba chamada menu para atribuir os menus de acesso ao seu projeto.
•	Crie um modulo padrão ou utilize o modulo de sua preferencia.
•	Crie o procedimento  que sera analisado pelo tratamento de erro dentro do modulo. 

#2: Inserir o modulo e a classe  para tratamento de erros	

•	Click na tecla Ctrl + F11, para abrir o projeto na parte dos modulos.
•	Click na raiz do projeto ( obs> verá o nome "VBAProject(nome do arquivo.xlsm)”
•	Click com o botão direiro e escolha "import file"
•	Selecione o modulo e classe que serão importados, porem tem que ser um a um.
#3: Chamada do modulo e classe:
•	 Dentro de seu procedimentos ou função , crie uma variavel erro.
#	exemplo :
            'Public sub Nome_Da_Funcao() As String
                           ‘On Error GoTo erro
                               '=====INICIO DO SEU PROCESSO========                                
                                      'processamento1
                                      'processamento1
                                      'processamento1
                                '=====FIM DO SEU PROCESSO===========
                                '    Exit Function
                                'erro:                             
            ' Obs:Invoque o modulo de tratamento de erros
            ' Call mdlError.MsnError(Err.Number, Err.Description, "Nome_Da_Funcao", Now())
 
             End Function
        
    
#4: Detalhes do Processo:
	mdlError.bas
o	->  sub MsnError   ex: [MsnError (scodError As Long, MsnError As String, sProcedimento As String, sdata As String)]

•	Há somente a necessidade de chamar este procedimento em seu tratamento de erros.
•	Passanddo como parametros o codigo de erro interceptado.
•	Passando a mensagem de erro interceptada.
•	Nome do procedimento aonde ocorreu o erro.
•	Data e hora que ocorreu o erro.
•	O Modulo irá  chamar a classe de erros.    
Obs:-> (os procedimento abaixo listado é executado pela classe, sem a necessidade de qualquer programação)
    
	 Verificar se a tabela de erros ja esta criada.
		•Caso não estiver criada, cria a tabela
			•Atribui o formato ja  programado na classe.
					o Verifica se o codigo de erro ja exite.
							 caso existir:
							o	verifica se a mensagem programada esta descrita no sistema.
								o	caso não estiver
										informa a mensagem do sistema.
								o	caso existir
										informa a mensagem programada
						fim da verificação.
						Adiciona no contador a quantidade de erros ocorridos deste codigo
						Emite a mensagem de erro formatada na tela, com formato programado.
              
•	#Ganhos:
1)	Rapida localização de erros quando houver.
2)	Rapida Identificação quando erro sistemico e ou erro de entrada de dados do usuario, identificação caminho, etc...
3)	Rapidez no desenvolvimento de tratamento de erros.
4)	Padronização de mensagens de erros, apresentada para o usuario.
5)	Padronização no estilo da mensagem de erros.
6)	e outros...

