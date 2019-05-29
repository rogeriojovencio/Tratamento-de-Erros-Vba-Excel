Attribute VB_Name = "mdlError"
Dim clsformatErrorClass As New clsFormatError
Public Const shtName = "shtErr"
Public Const sSistema = "Sistema de Gestao Relat�rio CGA"
Public intErro As Integer    'Intercepta o procedimento chamador caso for encontrado erro nos procedimentos chamados.
Sub MsnError(scodError As Long, MsnError As String, sProcedimento As String, sdata As String)

    With clsformatErrorClass
        .code = scodError
        .Msn1 = MsnError
        .Procedimento = sProcedimento
        .data = sdata
        .Sistema = sSistema
        .fcnMessageError
    End With
    
    '=============================XX====================================================================================
    '=======================INSTRU��ES DE USO===========================================================================
    '=============================XX====================================================================================
    'Primeiros passos:
        '* Crie uma aba chamda menu para atribuir os menus de acesso ao seu projeto.
        ' Inserir o Modulo e a classe disponibilizada em seu projeto vba.
            '* Click na tecla Ctrl + F11, para abrir o projeto na parte dos modulos.
            '* click na raiz do projeto ( obs> ver� o nome "VBAProject(nome do arquivo.xlsm)
            '* click com o bot�o direiro e escolha "import file"
            '* selecione os modulos e classes que ser�o importados, porem tem que ser um a um.
        
    ' Segundos passos:
        'quanto a chamada:-
            '* nos seus procedimentos, crie uma variavel erro.
               'exemplo :
                            'Public Function Nome_Da_Funcao() As String
                            'On Error GoTo erro
                            ' intErro = 0    'OPCIONAL->  toda variavel integer no VBA,  tem como valor padr�o 0, porem como ela � publica a necessidade de atribuir o valor no momento que ela entra no procedimento.
                                '=====INICIO DO SEU PROCESSO========
                                
                                      'processamento1
                                      'processamento1
                                      'processamento1
                                '=====FIM DO SEU PROCESSO===========
                                '    Exit Function
                                'erro:
                                '
                                '    intErro = 1 OPCIONAL->     'Observa��o: ou voce atual no retorno da fun��o ou utiliza a variavel intErro para tomar alguma decis�o, logo apos o erro ou acerto da fun��o.
                                '    Call mdlError.MsnError(Err.Number, Err.Description, "Nome_Da_Funcao", Now())
                                '    Application.Cursor = xlNormal
                                
                            'End Function
        
    
'obs:
'    - h� somente a necessidade de chamar este procedimento em seu tratamento de erros.
'    - passanddo como parametros o codigo de erro interceptado
'    - passando a mensagem de erro interceptada
'    ' nome do procedimento aonde ocorreu o erro
'    - Data e hora que ocorreu o erro.

    'O Modulo ir�  chamar a clase de erros:
    
    'obs:-> (os procedimentos abaixo listados � executado pela classe, sem a necessidade de qualquer programa��o)
    
       'Verificar se a tabela de erros ja esta criada.
           'caso n�o estiver criada cria a tabela
               ' atribui o formato ja  programado na classe.
           
        'Verifica se o codigo de erro ja exite.
            ' caso existir
                'verifica se a mensagem programada esta descrita no sistema
                        'caso n�o estiver
                            'informa a mensagem do sistema
                        'caso existir
                            'informa a mensagem programada
              'fim da verifica��o.
              'adiciona no contador a quantidade de erros ocorridos deste codigo
              'Emite a mensagem de erro formatada na tela, com formato programado.
              
    'Ganhos:
    ' 1)'Rapida localiza��o de erros quando houver
    ' 2)'Rapida Identifica��o quando erro sistemico e ou erro de entrada de dados do usuario, identifica��o caminho, etc...
    ' 3)'Rapidez no desenvolvimento de tratamento de erros.
    ' 4) Padroniza��o de mensagens de erros, apresentada para o usuario
    ' 5) padroniza��o no estilo da mentagem de erros
    ' 6) e outros...
    '=============================XX====================================================================================
    '=======================INSTRU��ES DE USO===========================================================================
    '=============================XX====================================================================================
End sub



