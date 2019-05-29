Attribute VB_Name = "mdlError"
Dim clsformatErrorClass As New clsFormatError
Public Const shtName = "shtErr"
Public Const sSistema = "Sistema de Gestao Relatório CGA"
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
    '=======================INSTRUÇÕES DE USO===========================================================================
    '=============================XX====================================================================================
    'Primeiros passos:
        '* Crie uma aba chamda menu para atribuir os menus de acesso ao seu projeto.
        ' Inserir o Modulo e a classe disponibilizada em seu projeto vba.
            '* Click na tecla Ctrl + F11, para abrir o projeto na parte dos modulos.
            '* click na raiz do projeto ( obs> verá o nome "VBAProject(nome do arquivo.xlsm)
            '* click com o botão direiro e escolha "import file"
            '* selecione os modulos e classes que serão importados, porem tem que ser um a um.
        
    ' Segundos passos:
        'quanto a chamada:-
            '* nos seus procedimentos, crie uma variavel erro.
               'exemplo :
                            'Public Function Nome_Da_Funcao() As String
                            'On Error GoTo erro
                            ' intErro = 0    'OPCIONAL->  toda variavel integer no VBA,  tem como valor padrão 0, porem como ela é publica a necessidade de atribuir o valor no momento que ela entra no procedimento.
                                '=====INICIO DO SEU PROCESSO========
                                
                                      'processamento1
                                      'processamento1
                                      'processamento1
                                '=====FIM DO SEU PROCESSO===========
                                '    Exit Function
                                'erro:
                                '
                                '    intErro = 1 OPCIONAL->     'Observação: ou voce atual no retorno da função ou utiliza a variavel intErro para tomar alguma decisão, logo apos o erro ou acerto da função.
                                '    Call mdlError.MsnError(Err.Number, Err.Description, "Nome_Da_Funcao", Now())
                                '    Application.Cursor = xlNormal
                                
                            'End Function
        
    
'obs:
'    - há somente a necessidade de chamar este procedimento em seu tratamento de erros.
'    - passanddo como parametros o codigo de erro interceptado
'    - passando a mensagem de erro interceptada
'    ' nome do procedimento aonde ocorreu o erro
'    - Data e hora que ocorreu o erro.

    'O Modulo irá  chamar a clase de erros:
    
    'obs:-> (os procedimentos abaixo listados é executado pela classe, sem a necessidade de qualquer programação)
    
       'Verificar se a tabela de erros ja esta criada.
           'caso não estiver criada cria a tabela
               ' atribui o formato ja  programado na classe.
           
        'Verifica se o codigo de erro ja exite.
            ' caso existir
                'verifica se a mensagem programada esta descrita no sistema
                        'caso não estiver
                            'informa a mensagem do sistema
                        'caso existir
                            'informa a mensagem programada
              'fim da verificação.
              'adiciona no contador a quantidade de erros ocorridos deste codigo
              'Emite a mensagem de erro formatada na tela, com formato programado.
              
    'Ganhos:
    ' 1)'Rapida localização de erros quando houver
    ' 2)'Rapida Identificação quando erro sistemico e ou erro de entrada de dados do usuario, identificação caminho, etc...
    ' 3)'Rapidez no desenvolvimento de tratamento de erros.
    ' 4) Padronização de mensagens de erros, apresentada para o usuario
    ' 5) padronização no estilo da mentagem de erros
    ' 6) e outros...
    '=============================XX====================================================================================
    '=======================INSTRUÇÕES DE USO===========================================================================
    '=============================XX====================================================================================
End sub



