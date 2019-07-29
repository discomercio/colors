Attribute VB_Name = "mod_BDPARAM"
Option Explicit
' ____________________________________________________________________________________________________
'|
'|  BANCO DE DADOS
'|
'|  IMPORTANTE:
'|    1. SEMPRE QUE ABRIR UMA CONEXÃO, CERTIFIQUE-SE DE EXECUTAR OS COMANDOS RETORNADOS POR
'|       BD_COMANDOS_INICIALIZACAO().
'|    3. ANTES DE IMPLEMENTAR ALTERAÇÕES, LEIA AS OBSERVAÇÕES AO FINAL DESTA SEÇÃO !!
'|

 
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' ******* ANTES DE COMPILAR, CONFIGURE AQUI O AMBIENTE !!         *******
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Global Const DESENVOLVIMENTO = True
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' ******* ANTES DE COMPILAR, CONFIGURE AQUI O BD A SER USADO !!   *******
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    #Const BD_DIRETIVA_TIPO_SERVIDOR = 2
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



  ' DEFINE SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~
    Global Const BD_SERVIDOR_ACCESS = 1
    Global Const BD_SERVIDOR_SQLSERVER = 2
    Global Const BD_SERVIDOR_ORACLE = 4
    Global Const BD_SERVIDOR_MSDE = 8
    
  ' CONSTANTES PARA COMPILAÇÃO CONDICIONAL: DEVEM SEGUIR OS
  ' MESMOS CÓDIGOS USADOS PARA DEFINIR O TIPO DE SERVIDOR DE BD
    #Const BD_DIRETIVA_SERVIDOR_ACCESS = 1
    #Const BD_DIRETIVA_SERVIDOR_SQLSERVER = 2
    #Const BD_DIRETIVA_SERVIDOR_ORACLE = 4
    #Const BD_DIRETIVA_SERVIDOR_MSDE = 8
    
    
    Global BD_STRING_CONEXAO_SERVIDOR As String
    Global BD_STRING_CONEXAO_SERVIDOR_AT As String
    Global BD_STRING_CONEXAO_SERVIDOR_CEP As String
        
    Type TIPO_PARAMETROS_CONEXAO_BD
        NOME_SERVIDOR As String
        NOME_BD As String
        ID_USUARIO As String
        SENHA_USUARIO As String
        descricao As String
        End Type
    
    Global bd_selecionado As TIPO_PARAMETROS_CONEXAO_BD
    Global bd_selecionado_at As TIPO_PARAMETROS_CONEXAO_BD
    Global bd_selecionado_cep As TIPO_PARAMETROS_CONEXAO_BD
      
      
  ' DEFINE OS PARÂMETROS PARA ABRIR BD
    #If BD_DIRETIVA_TIPO_SERVIDOR = BD_DIRETIVA_SERVIDOR_SQLSERVER Then
      ' SQL SERVER
      ' ~~~~~~~~~~
        Global Const BD_TIPO_SERVIDOR = BD_SERVIDOR_SQLSERVER
        Global Const BD_ID_SGBD = BD_SERVIDOR_SQLSERVER
        Global Const BD_DRIVER = "{SQL Server}"
        Global Const BD_OLEDB_PROVIDER = "SQLOLEDB"
    
    #ElseIf BD_DIRETIVA_TIPO_SERVIDOR = BD_DIRETIVA_SERVIDOR_MSDE Then
      ' MSDE (SQL SERVER)
      ' ~~~~~~~~~~~~~~~~~
        Global Const BD_TIPO_SERVIDOR = BD_SERVIDOR_SQLSERVER
        Global Const BD_ID_SGBD = BD_SERVIDOR_MSDE
        Global Const BD_DRIVER = "{SQL Server}"
        Global Const BD_OLEDB_PROVIDER = "SQLOLEDB"
              
    #ElseIf BD_DIRETIVA_TIPO_SERVIDOR = BD_DIRETIVA_SERVIDOR_ORACLE Then
      ' ORACLE
      ' ~~~~~~
        Global Const BD_TIPO_SERVIDOR = BD_SERVIDOR_ORACLE
        Global Const BD_ID_SGBD = BD_SERVIDOR_ORACLE
        Global Const BD_DRIVER = "{Microsoft ODBC for Oracle}"
        Global Const BD_OLEDB_PROVIDER = "OraOLEDB.Oracle"  ' IMPORTANTE: "MSDAORA" NÃO SUPORTA CAMPO MEMO
        #End If
        
    
    
  ' PARÂMETROS GERAIS DE CONFIGURAÇÃO
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' DEFINE OS PARÂMETROS QUE PODEM SER DEPENDENTES DO SERVIDOR DE BANCO DE DADOS
  ' MESMO QUE SEJAM IGUAIS AGORA, PODE SER INCLUÍDO UM NOVO SGBD COM PARÂMETROS DIFERENTES.
  
    #If (BD_DIRETIVA_TIPO_SERVIDOR = BD_DIRETIVA_SERVIDOR_SQLSERVER) Or _
        (BD_DIRETIVA_TIPO_SERVIDOR = BD_DIRETIVA_SERVIDOR_MSDE) Then
      ' CARACTER CURINGA USADO NO LIKE
        Global Const BD_CURINGA_TODOS = "%"
        Global Const BD_CURINGA_SINGLE_CHAR = "_"
        Global Const BD_DATE_PART_DAY = "day"
        Global Const BD_DATE_PART_MONTH = "month"
            
    #ElseIf BD_DIRETIVA_TIPO_SERVIDOR = BD_DIRETIVA_SERVIDOR_ORACLE Then
      ' CARACTER CURINGA USADO NO LIKE
        Global Const BD_CURINGA_TODOS = "%"
        Global Const BD_CURINGA_SINGLE_CHAR = "_"
        Global Const BD_DATE_PART_DAY = "day"
        Global Const BD_DATE_PART_MONTH = "month"
        
        #End If

    
    Global Const SQL_COLLATE_CASE_ACCENT = " COLLATE Latin1_General_CI_AI"

  ' LockType: Indicates the type of locks placed on records during editing.
  ' Opções: adLockOptimistic, adLockPessimistic, adLockReadOnly (default), adLockBatchOptimistic
    Global Const BD_POLITICA_LOCKING = adLockOptimistic

  ' CursorLocation: Sets or returns the location of the cursor service.
  ' Opções: adUseClient, adUseServer e adUseNone (obsoleto)
    Global Const BD_POLITICA_CURSOR = adUseClient
  
  ' ConnectionTimeout: Indicates how long to wait while establishing a connection before terminating the attempt and generating an error.
  ' Tempo em segundos.
    Global Const BD_CONNECTION_TIMEOUT = 45
  
  ' CommandTimeout: Indicates how long to wait while executing a command before terminating the attempt and generating an error.
  ' Obs: the Command object’s CommandTimeout property does not inherit the value of the Connection object’s CommandTimeout value.
  ' Tempo em segundos.
    Global Const BD_COMMAND_TIMEOUT = 900
    
  ' CacheSize: Indicates the number of records from a Recordset object that are cached locally in memory.
    Global Const BD_CACHE_CONSULTA = 30
    
  ' CursorType: Indicates the type of cursor used in a Recordset object.
  ' Opções: adOpenForwardOnly (default), adOpenKeyset, adOpenDynamic, adOpenStatic
    Global Const BD_CURSOR_SOMENTE_LEITURA = adOpenStatic
    Global Const BD_CURSOR_EDICAO = adOpenKeyset
    
  ' Especifica a quantidade de bytes a ser retornado pelo método GetChunk(BD_MAX_CHUNKSIZE) do objeto Field.
    Global Const BD_MAX_CHUNKSIZE = 65000
    
    
    
    


  ' FUNÇÃO QUE MONTA SUBSTRING DO JOIN PARA CONSULTAS SQL
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Global Const BD_NENHUM_JOIN = 0
    Global Const BD_INNER_JOIN = 1
    Global Const BD_LEFT_JOIN = 2
    Global Const BD_RIGHT_JOIN = 4
    
    
    Type TIPO_CAMPOS_RESTRICAO_JOIN
       campo_tabela_left As String
       campo_tabela_right As String
       End Type
       
       
    Type TIPO_PARAMETRO_JOIN
       tipo_join As Integer
       nome_tabela As String
       campos_join(1 To 10) As TIPO_CAMPOS_RESTRICAO_JOIN
       End Type



' ____________________________________________________________________________________________________
'|
'|   O B S E R V A Ç Õ E S
'|
'|   CursorLocation
'|   ==============
'|   Notou-se um consumo elevado de memória quando configurado como adUseClient,
'|   entretanto, adUseServer mostrou-se muito lento.
'|
'|   Recordsets
'|   ==========
'|   Ao abrir vários recordsets, pode ser que várias conexões sejam abertas de modo automático
'|   e implícito caso o driver perceba que não é possível compartilhar a conexão entre
'|   os recordsets. Isso ocorre principalmente se os recordsets forem criados usando um objeto
'|   "command" com a propriedade "prepared=True" (provavelmente porque nesse caso o servidor de
'|   BD cria uma stored procedure temporária). Configurar o CursorType=adOpenForwardOnly também
'|   faz o SQL Server abrir uma nova conexão para o recordset.
'|   Para executar uma transação, é necessário que exista apenas uma única conexão, portanto, é
'|   preciso garantir que os recordsets desnecessários estejam fechados nesse momento.
'|
'|   Order By
'|   ========
'|   No SQL Server 6.5, notou-se que em alguns casos a presença da cláusula ORDER BY impedia
'|   que um recordset fosse aberto no modo adOpenKeyset ou adOpenDynamic.
'|
'|   Cláusula Distinct
'|   =================
'|   Bug: esta cláusula funciona somente se o CursorLocation for igual a adUseServer.
'|   Se o CursorLocation for adUseClient, a cláusula DISTINCT é ignorada.
'|
'|   Case Sensitive no Operador Like
'|   ===============================
'|   No Oracle, ao fazer uma consulta usando o operador LIKE, lembre-se de que a comparação
'|   sempre é "case sensitive". Por exemplo: SELECT * FROM T_USUARIO WHERE NOME LIKE 'Mar%'
'|   irá encontrar "Maria", mas não irá encontrar "MARIANA" ou "marialva".
'|   Portanto, a consulta deveria estar escrita da seguinte forma:
'|   "SELECT * FROM T_USUARIO WHERE " & bd_monta_uppercase("NOME") & " LIKE 'MAR%'"
'|
'|   adCmdTable
'|   ==========
'|   Ao abrir um recordset usando a opção adCmdTable ao invés de adCmdText faz com que
'|   uma grande quantidade de dados seja transferida do servidor.  Isso somente é
'|   percebido quando as tabelas possuem muitos registros, pois nesta situação há uma
'|   longa demora para abrir o recordset.
'|   A solução adotada foi abrir o recordset usando "SELECT * FROM T_TABELA WHERE ..."
'|   e especificando restrições na cláusula where que não encontrarão nenhum registro.
'|   A partir daí, é executado o método AddNew do recordset normalmente.
'|   Exemplo de diferença de performance:
'|   Usando uma base de dados com 6400 regs em t_contrato, 12000 regs em t_versao e
'|   8200 regs em t_item foi realizado um teste cadastrando-se um contrato-base contendo
'|   somente 1 item contratual e nenhum lançamento.
'|   adCmdTable:
'|       Oracle demora 30 segundos p/ gravar, envia 3 MB e recebe 10 MB de dados do servidor.
'|       Sql Server demora 6 segundos p/ gravar, envia 3 KB (isso mesmo, KiloBytes) e recebe 5 MB de dados do servidor.
'|   adCmdText:
'|       Oracle demora 2 segundos p/ gravar, envia 15,2 KB e recebe 38 KB de dados do servidor.
'|       Sql Server demora 1 segundo p/ gravar, envia 3,7 KB e recebe 7 KB de dados do servidor.
'|
'|   Mode
'|   ====
'|   A propriedade "Mode" da conexão não funciona para o SQL Server.  Ao contrário
'|   do que diz a documentação da Microsoft (e seus exemplos), ela não tem efeito
'|   para o SQL Server.  No Access, ela funciona normalmente.
'|
'|
'|
'|



    Global Const msg_ERRO_ACESSO_CONCORRENTE = "Há outro usuário acessando o registro.  Será necessário aguardar que esse usuário termine a operação que está realizando para que o registro seja desbloqueado."


Function obtem_data_servidor(ByRef data_hora As Date, ByRef msg_erro As String) As Boolean
' ______________________________________________________________________________
'|
'|  CONSULTA A DATA NO SERVIDOR DE BANCO DE DADOS
'|
    
Dim s As String
Dim t As ADODB.Recordset
    
    On Error GoTo ODS_TRATA_ERRO
    
    obtem_data_servidor = False
    
    msg_erro = ""
        
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
    s = bd_monta_getdate("data_sistema")
    t.Open s, dbc, , , adCmdText
    If Not t.EOF Then
        If IsDate(t("data_sistema")) Then
            obtem_data_servidor = True
            data_hora = t("data_sistema")
            End If
        End If
        
    GoSub ODS_FECHA_TABELAS
    
Exit Function
    
    
    
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 
'   ODS_FECHA_TABELAS
' 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ODS_FECHA_TABELAS:
'==================
    bd_desaloca_recordset t, True
    Return
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 
'   TRATAMENTO DE ERRO
' 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ODS_TRATA_ERRO:
'==============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    GoSub ODS_FECHA_TABELAS
    Exit Function
    
End Function


Function bd_filtra_acentuacao(ByVal campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE SUBSTITUI LETRAS ACENTUADAS POR LETRAS SEM ACENTUAÇÃO.
'|

Dim i As Integer
Dim s As String
Dim s_resp As String
Dim s_char As String


    On Error GoTo BD_FILTRA_ACENTUACAO_ERRO
    
    
    bd_filtra_acentuacao = ""

    
    s_resp = ""

    For i = 1 To Len(campo)
        s_char = Mid$(campo, i, 1)
        
        Select Case s_char
            Case "ÿ", "ý"
                s_char = "y"
                
            Case "Ÿ", "Ý"
                s_char = "Y"
                
            Case "ñ"
                s_char = "n"
                
            Case "Ñ"
                s_char = "N"
                
            Case "ç"
                s_char = "c"
                
            Case "Ç"
                s_char = "C"
                
            Case "á", "à", "ã", "â", "ä", "å"
                s_char = "a"
                
            Case "Á", "À", "Ã", "Â", "Ä", "Å"
                s_char = "A"
                
            Case "é", "è", "ê", "ë"
                s_char = "e"
                
            Case "É", "È", "Ê", "Ë"
                s_char = "E"
                
            Case "í", "ì", "î", "ï"
                s_char = "i"
                
            Case "Í", "Ì", "Î", "Ï"
                s_char = "I"
                
            Case "ó", "ò", "õ", "ô", "ö"
                s_char = "o"
                
            Case "Ó", "Ò", "Õ", "Ô", "Ö"
                s_char = "O"
                
            Case "ú", "ù", "û", "ü"
                s_char = "u"
                
            Case "Ú", "Ù", "Û", "Ü"
                s_char = "U"
                
            End Select
        
        
        s_resp = s_resp & s_char
        
        Next
        
        
    bd_filtra_acentuacao = s_resp
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_FILTRA_ACENTUACAO_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function


Function bd_monta_as(campo As String, apelido As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A FUNÇÃO QUE RENOMEIA O CAMPO EM UMA CONSULTA, PARA CONSULTAS SQL
'|    NO SERVIDOR DE BD EM USO.
'|  - PARÂMETROS:
'|    - CAMPO: NOME DO CAMPO NA TABELA
'|    - APELIDO: APELIDO DO CAMPO NA CONSULTA
'|
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_AS_ERRO
    
    
    bd_monta_as = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_as = campo & " AS " & apelido
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_as = campo & " AS " & apelido


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_as = campo & " AS " & apelido
        
        End Select
        
    
    If bd_monta_as <> "" Then bd_monta_as = " " & Trim$(bd_monta_as)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_AS_ERRO:
'~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   
        

End Function


Function bd_monta_ascend(ByVal campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA O COMANDO DE ORDENAÇÃO ASCENDENTE PARA A
'|    CLÁUSULA "ORDER BY", PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO.
'|  - PARÂMETROS:
'|    - CAMPO: NOME DO CAMPO NA TABELA USADO NA ORDENAÇÃO.
'|
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|


Dim s As String


    On Error GoTo BD_MONTA_ASCEND_ERRO
    
    
    bd_monta_ascend = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_ascend = campo & " ASC"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_ascend = campo & " ASC"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_ascend = campo & " ASC"
        
        End Select
        
    
    If bd_monta_ascend <> "" Then bd_monta_ascend = " " & Trim$(bd_monta_ascend)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_ASCEND_ERRO:
'~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   
        

End Function



Function bd_monta_descend(ByVal campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA O COMANDO DE ORDENAÇÃO DESCENDENTE PARA A
'|    CLÁUSULA "ORDER BY", PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO.
'|  - PARÂMETROS:
'|    - CAMPO: NOME DO CAMPO NA TABELA USADO NA ORDENAÇÃO.
'|
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|


Dim s As String


    On Error GoTo BD_MONTA_DESCEND_ERRO
    
    
    bd_monta_descend = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_descend = campo & " DESC"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_descend = campo & " DESC"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_descend = campo & " DESC"
        
        End Select
        
    
    If bd_monta_descend <> "" Then bd_monta_descend = " " & Trim$(bd_monta_descend)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_DESCEND_ERRO:
'~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   
        

End Function




Function bd_comandos_inicializacao(v() As String) As Boolean
' _________________________________________________________________________________________________
'|
'|  _ *** IMPORTANTE ***: OS COMANDOS RETORNADOS POR ESTA FUNÇÃO DEVEM SER EXECUTADOS
'|    SEMPRE QUE QUE UMA CONEXÃO COM O BANCO DE DADOS FOR ABERTA.
'|
'|  _ A FUNÇÃO RETORNA 'TRUE' SE HOUVER COMANDOS A SEREM EXECUTADOS, CASO CONTRÁRIO,
'|    RETORNA 'FALSE'.
'|  _ OS COMANDOS EM QUESTÃO SÃO NECESSÁRIOS PARA CONFIGURAR A SESSÃO DO BANCO DE DADOS,
'|    PORTANTO, DEVEM SER EXECUTADOS UMA ÚNICA VEZ AO ABRIR A CONEXÃO.
'|  _ CADA COMANDO A SER EXECUTADO CORRESPONDE A UMA POSIÇÃO DO VETOR.
'|

Dim s As String
Dim i As Integer
Dim i_index As Integer


    On Error GoTo BD_COMANDOS_INICIALIZACAO_ERRO
        
    
    bd_comandos_inicializacao = False
    
    i_index = 0
    ReDim v(i_index)
    v(i_index) = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
        ' NOP
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' NOP


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
          ' PARA CLÁUSULA "ORDER BY" FUNCIONAR CORRETAMENTE
            v(i_index) = "ALTER SESSION SET NLS_SORT = BINARY"
            
        End Select
        
        
  ' HÁ ALGUM COMANDO A SER EXECUTADO ?
    For i = LBound(v) To UBound(v)
        If Trim$(v(i)) <> "" Then
            bd_comandos_inicializacao = True
            Exit For
            End If
        Next
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_COMANDOS_INICIALIZACAO_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function

Function bd_monta_escape(campo As String, texto_comparacao As String, converter_para_literais As String, caracter_escape As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA UMA RESTRIÇÃO DA CLÁUSULA WHERE CUJO CAMPO CONTENHA NA
'|    SEQUÊNCIA A SER PESQUISADA UM CARACTER ESPECIAL.  PARA QUE A CONDIÇÃO NO
'|    'WHERE' SEJA CONSTRUÍDA CORRETAMENTE, SÃO COLOCADOS 'ESCAPES' PARA QUE O
'|    CARACTER ESPECIAL SEJA TRATADO COMO UM LITERAL.
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s_bd_wildcards As String
Dim deve_converter As Boolean
Dim tem_wildcard As Boolean
Dim usou_caracter_escape As Boolean
Dim s As String
Dim i As Long


    On Error GoTo BD_MONTA_ESCAPE_ERRO
    
    
    bd_monta_escape = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            s_bd_wildcards = "*?#[]"
            
          ' VERIFICA SE O TEXTO TEM WILDCARDS
            tem_wildcard = False
            For i = 1 To Len(texto_comparacao)
                If InStr(s_bd_wildcards, Mid$(texto_comparacao, i, 1)) <> 0 Then
                    tem_wildcard = True
                    Exit For
                    End If
                Next
            
            
            s = ""
            For i = 1 To Len(texto_comparacao)
                deve_converter = False
              ' É UM DOS CARACTERES WILDCARDS A SEREM CONVERTIDOS ?
                If InStr(converter_para_literais, Mid$(texto_comparacao, i, 1)) <> 0 Then
                  ' CARACTER REALMENTE É UM WILDCARD DO BD EM USO ?
                    If InStr(s_bd_wildcards, Mid$(texto_comparacao, i, 1)) <> 0 Then deve_converter = True
                    End If
                
                If deve_converter Then
                  ' ADICIONA O ESCAPE P/ QUE O WILDCARD SEJA TRATADO COMO LITERAL
                    s = s & "[" & Mid$(texto_comparacao, i, 1) & "]"
                Else
                    s = s & Mid$(texto_comparacao, i, 1)
                    End If
                Next
                
                
            If tem_wildcard Then
                bd_monta_escape = campo & " LIKE '" & s & "'"
            Else
                bd_monta_escape = campo & " = '" & s & "'"
                End If
                
            
        
        
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            s_bd_wildcards = "%_"
            
          ' VERIFICA SE O TEXTO TEM WILDCARDS
            tem_wildcard = False
            For i = 1 To Len(texto_comparacao)
                If InStr(s_bd_wildcards, Mid$(texto_comparacao, i, 1)) <> 0 Then
                    tem_wildcard = True
                    Exit For
                    End If
                Next
            
            
            usou_caracter_escape = False
            s = ""
            
            For i = 1 To Len(texto_comparacao)
                deve_converter = False
              ' É UM DOS CARACTERES WILDCARDS A SEREM CONVERTIDOS ?
                If InStr(converter_para_literais, Mid$(texto_comparacao, i, 1)) <> 0 Then
                  ' CARACTER REALMENTE É UM WILDCARD DO BD EM USO ?
                    If InStr(s_bd_wildcards, Mid$(texto_comparacao, i, 1)) <> 0 Then deve_converter = True
                    End If
                
                If deve_converter Then
                  ' ADICIONA O ESCAPE P/ QUE O WILDCARD SEJA TRATADO COMO LITERAL
                    s = s & caracter_escape & Mid$(texto_comparacao, i, 1)
                    usou_caracter_escape = True
                Else
                  ' P/ STRINGS QUE CONTENHAM O CARACTER USADO NO ESCAPE É NECESSÁRIO
                  ' QUE ESSE CARACTER SEJA REPETIDO DUAS VEZES
                    If Mid$(texto_comparacao, i, 1) = caracter_escape Then
                        s = s & caracter_escape
                        usou_caracter_escape = True
                        End If
                        
                    s = s & Mid$(texto_comparacao, i, 1)
                    End If
                Next
                
                
            If tem_wildcard Then
                bd_monta_escape = campo & " LIKE '" & s & "'"
            Else
                bd_monta_escape = campo & " = '" & s & "'"
                End If
                
                
            If usou_caracter_escape Then
                bd_monta_escape = bd_monta_escape & " ESCAPE '" & caracter_escape & "'"
                End If





        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            s_bd_wildcards = "%_"
            
          ' VERIFICA SE O TEXTO TEM WILDCARDS
            tem_wildcard = False
            For i = 1 To Len(texto_comparacao)
                If InStr(s_bd_wildcards, Mid$(texto_comparacao, i, 1)) <> 0 Then
                    tem_wildcard = True
                    Exit For
                    End If
                Next
            
            
            usou_caracter_escape = False
            s = ""
            
            For i = 1 To Len(texto_comparacao)
                deve_converter = False
              ' É UM DOS CARACTERES WILDCARDS A SEREM CONVERTIDOS ?
                If InStr(converter_para_literais, Mid$(texto_comparacao, i, 1)) <> 0 Then
                  ' CARACTER REALMENTE É UM WILDCARD DO BD EM USO ?
                    If InStr(s_bd_wildcards, Mid$(texto_comparacao, i, 1)) <> 0 Then deve_converter = True
                    End If
                
                If deve_converter Then
                  ' ADICIONA O ESCAPE P/ QUE O WILDCARD SEJA TRATADO COMO LITERAL
                    s = s & caracter_escape & Mid$(texto_comparacao, i, 1)
                    usou_caracter_escape = True
                Else
                  ' P/ STRINGS QUE CONTENHAM O CARACTER USADO NO ESCAPE É NECESSÁRIO
                  ' QUE ESSE CARACTER SEJA REPETIDO DUAS VEZES
                    If Mid$(texto_comparacao, i, 1) = caracter_escape Then
                        s = s & caracter_escape
                        usou_caracter_escape = True
                        End If
                        
                    s = s & Mid$(texto_comparacao, i, 1)
                    End If
                Next
                
                
            If tem_wildcard Then
                bd_monta_escape = campo & " LIKE '" & s & "'"
            Else
                bd_monta_escape = campo & " = '" & s & "'"
                End If
                
                
            If usou_caracter_escape Then
                bd_monta_escape = bd_monta_escape & " ESCAPE '" & caracter_escape & "'"
                End If
        
        
        End Select
        
        
            
    If bd_monta_escape <> "" Then bd_monta_escape = " " & Trim$(bd_monta_escape)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_ESCAPE_ERRO:
'~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function


Function bd_monta_to_char(campo As String, n_char As Integer) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A FUNÇÃO QUE CONVERTE NÚMERO PARA CHAR, PARA CONSULTAS SQL
'|    NO SERVIDOR DE BD EM USO.
'|  _ É NECESSÁRIO ESPECIFICAR O TAMANHO DO CAMPO CHAR (EX: CHAR(2), CHAR(10), ETC.)
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_TO_CHAR_ERRO
    
    
    bd_monta_to_char = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_to_char = "CSTR(" & campo & ")"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_to_char = "CONVERT(CHAR(" & CStr(n_char) & "), " & campo & ")"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_to_char = "TO_CHAR(" & campo & ")"
        
        End Select
        
    
    If bd_monta_to_char <> "" Then bd_monta_to_char = " " & Trim$(bd_monta_to_char)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_TO_CHAR_ERRO:
'~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   
        

End Function


Function bd_filtra_aspas(ByVal campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE ANALISA O TEXTO PARA VERIFICAR SE EXISTEM ASPAS NO TEXTO.
'|    SE HOUVER, ALTERA O TEXTO PARA QUE ELE POSSA SER USADO COMO PARÂMETRO
'|    DA CLÁUSULA WHERE DE CONSULTAS SQL.
'|

Dim i As Integer
Dim s As String
Dim s_char As String


    On Error GoTo BD_FILTRA_ASPAS_ERRO
    
    
    bd_filtra_aspas = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            s = ""
            For i = 1 To Len(campo)
                s_char = Mid$(campo, i, 1)
                s = s & s_char
                If s_char = "'" Then s = s & s_char
                Next
                
            bd_filtra_aspas = s
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            s = ""
            For i = 1 To Len(campo)
                s_char = Mid$(campo, i, 1)
                s = s & s_char
                If s_char = "'" Then s = s & s_char
                Next
                
            bd_filtra_aspas = s


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            s = ""
            For i = 1 To Len(campo)
                s_char = Mid$(campo, i, 1)
                s = s & s_char
                If s_char = "'" Then s = s & s_char
                Next
                
            bd_filtra_aspas = s
        
        End Select
        
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_FILTRA_ASPAS_ERRO:
'~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function

Function bd_monta_parametro_accent_insensitive(ByVal campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE PREPARA UM PARÂMETRO QUE SERÁ USADO EM UMA CLÁUSULA WHERE
'|    DE MODO QUE A COMPARAÇÃO SEJA INDEPENDENTE DA ACENTUAÇÃO DO CAMPO.
'|    PARA ISSO, OS CARACTERES QUE POSSUEM "VERSÕES" COM E SEM ACENTO SERÃO
'|    SUBSTITUÍDOS POR UM WILDCARD.
'|    OBVIAMENTE, PROVAVELMENTE SERÁ NECESSÁRIO UM PROCESSAMENTO POSTERIOR
'|    P/ ELIMINAR REGISTROS INDESEJADOS, POIS, POR EXEMPLO:
'|    SE EXECUTARMOS ESTA FUNÇÃO COM A PALAVRA "MAÇÃ", O RETORNO SERÁ "MA__".
'|    AO EXECUTAR UM SQL C/ ESSE PARÂMETRO, SERÃO RETORNADAS TODAS AS PALAVRAS
'|    QUE SE INICIEM COM "MA" SEGUIDAS DE DUAS LETRAS QUALQUER, SENDO QUE, NA
'|    VERDADE, SOMENTE SÃO DESEJADOS OS REGISTROS C/ AS PALAVRAS MACA, MAÇA,
'|    MACÃ E MAÇÃ.
'|    PORTANTO, NESTE CASO, SERÁ NECESSÁRIO UM PROCESSAMENTO PARA DESPREZAR OS
'|    REGISTROS C/ CAMPOS DIFERENTES QUE O DESEJADO: MALA, MAPA, ETC.
'|

Dim i As Integer
Dim s As String
Dim s_resp As String
Dim s_char As String
Dim s_char_a As String


    On Error GoTo BD_MPAI_TRATA_ERRO
    
    
    bd_monta_parametro_accent_insensitive = ""

    
    s_resp = ""

    For i = 1 To Len(campo)
        s_char_a = Mid$(campo, i, 1)
        s_char = bd_filtra_acentuacao(s_char_a)
        
      ' É UM CARACTER ACENTUADO ?
        If (s_char <> s_char_a) Then s_char = BD_CURINGA_SINGLE_CHAR
                
        Select Case UCase$(s_char)
            Case "C": s_char = BD_CURINGA_SINGLE_CHAR
            Case "A": s_char = BD_CURINGA_SINGLE_CHAR
            Case "E": s_char = BD_CURINGA_SINGLE_CHAR
            Case "I": s_char = BD_CURINGA_SINGLE_CHAR
            Case "O": s_char = BD_CURINGA_SINGLE_CHAR
            Case "U": s_char = BD_CURINGA_SINGLE_CHAR
            End Select
            
        s_resp = s_resp & s_char
        Next
        
        
    bd_monta_parametro_accent_insensitive = s_resp
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MPAI_TRATA_ERRO:
'~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function


Function bd_monta_uppercase(ByVal campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A FUNÇÃO UPPERCASE (RETORNA O TEXTO EM LETRAS MAIÚSCULAS)
'|    PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO.
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_UPPERCASE_ERRO
    
    
    bd_monta_uppercase = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_uppercase = "UCASE(" & campo & ")"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_uppercase = "UPPER(" & campo & ")"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_uppercase = "UPPER(" & campo & ")"
        
        End Select
        
    
    If bd_monta_uppercase <> "" Then bd_monta_uppercase = " " & Trim$(bd_monta_uppercase)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_UPPERCASE_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function

Function bd_monta_lowercase(ByVal campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A FUNÇÃO LOWERCASE (RETORNA O TEXTO EM LETRAS MINÚSCULAS)
'|    PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO.
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_LOWERCASE_ERRO
    
    
    bd_monta_lowercase = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_lowercase = "LCASE(" & campo & ")"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_lowercase = "LOWER(" & campo & ")"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_lowercase = "LOWER(" & campo & ")"
        
        End Select
        
    
    If bd_monta_lowercase <> "" Then bd_monta_lowercase = " " & Trim$(bd_monta_lowercase)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_LOWERCASE_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function


Function bd_monta_string_conexao_sgbd(ByRef msg_erro As String) As Boolean
' _________________________________________________________________________________________________________
'|
'|  OBTÉM OS PARÂMETROS DE CONEXÃO AO SERVIDOR DE BANCO DE DADOS.
'|  OS PARÂMETROS PODEM SER FIXOS OU LIDOS A PARTIR DE ARQUIVO INI.
'|

    On Error GoTo BD_MSCSGBD_TRATA_ERRO
    
    
    bd_monta_string_conexao_sgbd = False
    
    msg_erro = ""
    
    BD_STRING_CONEXAO_SERVIDOR = ""
    
  ' SQL SERVER
    If BD_ID_SGBD = BD_SERVIDOR_SQLSERVER Then
      ' OLE DB APPROACH: SQLSERVER
        BD_STRING_CONEXAO_SERVIDOR = "Provider=" & BD_OLEDB_PROVIDER & _
                                     ";Data Source=" & bd_selecionado.NOME_SERVIDOR & _
                                     ";Initial Catalog=" & bd_selecionado.NOME_BD & _
                                     ";User Id=" & bd_selecionado.ID_USUARIO & _
                                     ";Password=" & bd_selecionado.SENHA_USUARIO
            
  ' MSDE
    ElseIf BD_ID_SGBD = BD_SERVIDOR_MSDE Then
      ' OLE DB APPROACH: SQLSERVER
        BD_STRING_CONEXAO_SERVIDOR = "Provider=" & BD_OLEDB_PROVIDER & _
                                     ";Data Source=" & bd_selecionado.NOME_SERVIDOR & _
                                     ";Initial Catalog=" & bd_selecionado.NOME_BD & _
                                     ";User Id=" & bd_selecionado.ID_USUARIO & _
                                     ";Password=" & bd_selecionado.SENHA_USUARIO
                
  ' ORACLE
    ElseIf BD_ID_SGBD = BD_SERVIDOR_ORACLE Then
      ' OLE DB APPROACH: ORACLE
        BD_STRING_CONEXAO_SERVIDOR = "Provider=" & BD_OLEDB_PROVIDER & _
                                     ";Data Source=" & bd_selecionado.NOME_BD & _
                                     ";User Id=" & bd_selecionado.ID_USUARIO & _
                                     ";Password=" & bd_selecionado.SENHA_USUARIO
        End If
        
    
  ' HÁ PARÂMETROS VÁLIDOS ?
    If Trim$(BD_STRING_CONEXAO_SERVIDOR) <> "" Then bd_monta_string_conexao_sgbd = True
        
        
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MSCSGBD_TRATA_ERRO:
'=====================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    Exit Function
    

End Function

Function bd_monta_string_conexao_at(ByRef msg_erro As String) As Boolean
' _________________________________________________________________________________________________________
'|
'|  OBTÉM OS PARÂMETROS DE CONEXÃO AO SERVIDOR DE BANCO DE DADOS.
'|  OS PARÂMETROS PODEM SER FIXOS OU LIDOS A PARTIR DE ARQUIVO INI.
'|

    On Error GoTo BD_MSCAT_TRATA_ERRO
    
    
    bd_monta_string_conexao_at = False
    
    msg_erro = ""
    
    BD_STRING_CONEXAO_SERVIDOR_AT = ""
    
  ' SQL SERVER
    If BD_ID_SGBD = BD_SERVIDOR_SQLSERVER Then
      ' OLE DB APPROACH: SQLSERVER
        BD_STRING_CONEXAO_SERVIDOR_AT = "Provider=" & BD_OLEDB_PROVIDER & _
                                     ";Data Source=" & bd_selecionado_at.NOME_SERVIDOR & _
                                     ";Initial Catalog=" & bd_selecionado_at.NOME_BD & _
                                     ";User Id=" & bd_selecionado_at.ID_USUARIO & _
                                     ";Password=" & bd_selecionado_at.SENHA_USUARIO
            
  ' MSDE
    ElseIf BD_ID_SGBD = BD_SERVIDOR_MSDE Then
      ' OLE DB APPROACH: SQLSERVER
        BD_STRING_CONEXAO_SERVIDOR_AT = "Provider=" & BD_OLEDB_PROVIDER & _
                                     ";Data Source=" & bd_selecionado_at.NOME_SERVIDOR & _
                                     ";Initial Catalog=" & bd_selecionado_at.NOME_BD & _
                                     ";User Id=" & bd_selecionado_at.ID_USUARIO & _
                                     ";Password=" & bd_selecionado_at.SENHA_USUARIO
                
  ' ORACLE
    ElseIf BD_ID_SGBD = BD_SERVIDOR_ORACLE Then
      ' OLE DB APPROACH: ORACLE
        BD_STRING_CONEXAO_SERVIDOR_AT = "Provider=" & BD_OLEDB_PROVIDER & _
                                     ";Data Source=" & bd_selecionado_at.NOME_BD & _
                                     ";User Id=" & bd_selecionado_at.ID_USUARIO & _
                                     ";Password=" & bd_selecionado_at.SENHA_USUARIO
        End If
        
    
  ' HÁ PARÂMETROS VÁLIDOS ?
    If Trim$(BD_STRING_CONEXAO_SERVIDOR_AT) <> "" Then bd_monta_string_conexao_at = True
        
        
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MSCAT_TRATA_ERRO:
'=====================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    Exit Function
    

End Function

Function bd_monta_string_conexao_cep(ByRef msg_erro As String) As Boolean
' _________________________________________________________________________________________________________
'|
'|  OBTÉM OS PARÂMETROS DE CONEXÃO AO SERVIDOR DE BANCO DE DADOS.
'|  OS PARÂMETROS PODEM SER FIXOS OU LIDOS A PARTIR DE ARQUIVO INI.
'|

    On Error GoTo BD_MSCCEP_TRATA_ERRO
    
    
    bd_monta_string_conexao_cep = False
    
    msg_erro = ""
    
    BD_STRING_CONEXAO_SERVIDOR_CEP = ""
    
  ' SQL SERVER
    If BD_ID_SGBD = BD_SERVIDOR_SQLSERVER Then
      ' OLE DB APPROACH: SQLSERVER
        BD_STRING_CONEXAO_SERVIDOR_CEP = "Provider=" & BD_OLEDB_PROVIDER & _
                                     ";Data Source=" & bd_selecionado_cep.NOME_SERVIDOR & _
                                     ";Initial Catalog=" & bd_selecionado_cep.NOME_BD & _
                                     ";User Id=" & bd_selecionado_cep.ID_USUARIO & _
                                     ";Password=" & bd_selecionado_cep.SENHA_USUARIO

  ' MSDE
    ElseIf BD_ID_SGBD = BD_SERVIDOR_MSDE Then
      ' OLE DB APPROACH: SQLSERVER
        BD_STRING_CONEXAO_SERVIDOR_CEP = "Provider=" & BD_OLEDB_PROVIDER & _
                                     ";Data Source=" & bd_selecionado_cep.NOME_SERVIDOR & _
                                     ";Initial Catalog=" & bd_selecionado_cep.NOME_BD & _
                                     ";User Id=" & bd_selecionado_cep.ID_USUARIO & _
                                     ";Password=" & bd_selecionado_cep.SENHA_USUARIO

  ' ORACLE
    ElseIf BD_ID_SGBD = BD_SERVIDOR_ORACLE Then
      ' OLE DB APPROACH: ORACLE
        BD_STRING_CONEXAO_SERVIDOR_CEP = "Provider=" & BD_OLEDB_PROVIDER & _
                                     ";Data Source=" & bd_selecionado_cep.NOME_BD & _
                                     ";User Id=" & bd_selecionado_cep.ID_USUARIO & _
                                     ";Password=" & bd_selecionado_cep.SENHA_USUARIO
        End If


  ' HÁ PARÂMETROS VÁLIDOS ?
    If Trim$(BD_STRING_CONEXAO_SERVIDOR_CEP) <> "" Then bd_monta_string_conexao_cep = True


Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MSCCEP_TRATA_ERRO:
'=====================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    Exit Function
    

End Function


Sub bd_desaloca_command(c As ADODB.Command)
' _________________________________________________________________________________________________
'|
'|  DESALOCA A MEMÓRIA ASSOCIADA AO COMMAND 'C'.
'|

Dim s As String

    
    On Error GoTo BD_DESALOCA_COMMAND_ERRO
    
    If Not (c Is Nothing) Then Set c = Nothing
        

Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_DESALOCA_COMMAND_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Sub
        

End Sub


Sub bd_desaloca_parameter(p As ADODB.Parameter)
' _________________________________________________________________________________________________
'|
'|  DESALOCA A MEMÓRIA ASSOCIADA AO PARAMETER 'P'.
'|

Dim s As String

    
    On Error GoTo BD_DESALOCA_PARAMETER_ERRO
    
    If Not (p Is Nothing) Then Set p = Nothing
        

Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_DESALOCA_PARAMETER_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Sub
        


End Sub

Function bd_monta_dateadd(intervalo As String, qtde As String, campo_data As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA FUNÇÃO DATEADD PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO.
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String
Dim sinal As String
Dim s_aux As String


    On Error GoTo BD_MONTA_DATEADD_ERRO
    
    
    bd_monta_dateadd = ""

    

  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            If Mid$(Trim$(qtde), 1, 1) = "-" Then
                sinal = "-"
              ' REMOVE SINAL
                s_aux = Mid$(Trim$(qtde), 2, Len(qtde))
            Else
                sinal = ""
                s_aux = qtde
                End If
            
            s_aux = sinal & "CInt('0' & " & s_aux & ")"
            
            s = "DATEADD('" & intervalo & "'," & s_aux & "," & campo_data & ")"
            bd_monta_dateadd = s
            
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            s = "DATEADD(" & intervalo & "," & qtde & "," & campo_data & ")"
            bd_monta_dateadd = s
            
            

        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            Select Case intervalo
                Case BD_DATE_PART_DAY
                    sinal = "+"
                  ' JÁ VEM C/ SINAL ?
                    If Mid$(Trim$(qtde), 1, 1) = "-" Then sinal = ""
                    If Mid$(Trim$(qtde), 1, 1) = "+" Then sinal = ""
                    s = "(" & campo_data & sinal & qtde & ")"
                
                Case BD_DATE_PART_MONTH: s = "ADD_MONTHS(" & campo_data & "," & qtde & ")"
                
                Case Else: s = ""
                
                End Select
            
            bd_monta_dateadd = s
            
        End Select
        
    
    If bd_monta_dateadd <> "" Then bd_monta_dateadd = " " & Trim$(bd_monta_dateadd)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_DATEADD_ERRO:
'~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   
        

End Function

Function bd_monta_getdate(nome_do_campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A FUNÇÃO 'GETDATE' PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO
'|    (A FUNÇÃO GETDATE RETORNA A DATA/HORA DO SERVIDOR).
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_GETDATE_ERRO
    
    
    bd_monta_getdate = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_getdate = "SELECT NOW AS " & nome_do_campo & " FROM T_NUMERACAO"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_getdate = "SELECT '" & nome_do_campo & "' = GETDATE()"



        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_getdate = "SELECT SYSDATE AS " & nome_do_campo & " FROM T_NUMERACAO"
        
        End Select
        
    
    If bd_monta_getdate <> "" Then bd_monta_getdate = " " & Trim$(bd_monta_getdate)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_GETDATE_ERRO:
'~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function

Function bd_monta_is_not_null(campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A COMPARAÇÃO 'IS NOT NULL' PARA CONSULTAS SQL NO SERVIDOR
'|    DE BD EM USO
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_IS_NOT_NULL_ERRO
    
    
    bd_monta_is_not_null = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_is_not_null = campo & " IS NOT NULL"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_is_not_null = campo & " IS NOT NULL"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_is_not_null = campo & " IS NOT NULL"
        
        End Select
        
    
    If bd_monta_is_not_null <> "" Then bd_monta_is_not_null = " " & Trim$(bd_monta_is_not_null)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_IS_NOT_NULL_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function

Function bd_monta_is_null(campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A COMPARAÇÃO 'IS NULL' PARA CONSULTAS SQL NO SERVIDOR
'|    DE BD EM USO
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_IS_NULL_ERRO
    
    
    bd_monta_is_null = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_is_null = campo & " IS NULL"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_is_null = campo & " IS NULL"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_is_null = campo & " IS NULL"
        
        End Select
        
    
    If bd_monta_is_null <> "" Then bd_monta_is_null = " " & Trim$(bd_monta_is_null)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_IS_NULL_ERRO:
'~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function

Function bd_monta_len(campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A FUNÇÃO 'LEN' PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO
'|    (A FUNÇÃO LEN RETORNA O COMPRIMENTO DO TEXTO PASSADO POR PARÂMETRO).
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_LEN_ERRO
    
    
    bd_monta_len = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_len = "LEN(" & campo & ")"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_len = "LEN(" & campo & ")"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_len = "LENGTH(" & campo & ")"
        
        End Select
        
    
    If bd_monta_len <> "" Then bd_monta_len = " " & Trim$(bd_monta_len)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_LEN_ERRO:
'~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function

Function bd_monta_moeda(valor As Variant) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA UM VALOR MONETÁRIO EM UM FORMATO PRÓPRIO PARA UTILIZAR EM STRINGS
'|    DE CONSULTA SQL NO SERVIDOR DE BD EM USO.
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String
Dim s_aux As String


    On Error GoTo BD_MONTA_MOEDA_ERRO
    
    
    bd_monta_moeda = ""

    
    If Not IsNumeric(valor) Then Exit Function
    
    

  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            s = Format$(valor, "###########0.00")
          ' SE SEPARADOR DECIMAL FOR VÍRGULA, SUBSTITUI POR PONTO
            s = substitui_caracteres(s, ",", ".")
            bd_monta_moeda = s
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            s = Format$(valor, "###########0.00")
          ' SE SEPARADOR DECIMAL FOR VÍRGULA, SUBSTITUI POR PONTO
            s = substitui_caracteres(s, ",", ".")
            bd_monta_moeda = s
            

        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            s = Format$(valor, "###########0.00")
          ' SE SEPARADOR DECIMAL FOR VÍRGULA, SUBSTITUI POR PONTO
            s = substitui_caracteres(s, ",", ".")
            bd_monta_moeda = s
            
        End Select
        
    
    If bd_monta_moeda <> "" Then bd_monta_moeda = " " & Trim$(bd_monta_moeda)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_MOEDA_ERRO:
'~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   
        

End Function

Function bd_monta_decimal(ByVal valor As Variant, ByVal casas_decimais As Integer) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA UM VALOR DECIMAL EM UM FORMATO PRÓPRIO PARA UTILIZAR EM STRINGS
'|    DE CONSULTA SQL NO SERVIDOR DE BD EM USO.
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim i As Integer
Dim s As String
Dim s_aux As String
Dim s_mascara As String


    On Error GoTo BD_MONTA_DECIMAL_TRATA_ERRO
    
    
    bd_monta_decimal = ""

    
    If Not IsNumeric(valor) Then Exit Function
    
    
  ' PREPARA MÁSCARA P/ FORMATAÇÃO
    s_mascara = ""
    For i = 1 To casas_decimais
        s_mascara = s_mascara & "0"
        Next
    
    If s_mascara <> "" Then s_mascara = "." & s_mascara
    s_mascara = "###########0" & s_mascara
    

  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            s = Format$(valor, s_mascara)
          ' SE SEPARADOR DECIMAL FOR VÍRGULA, SUBSTITUI POR PONTO
            s = substitui_caracteres(s, ",", ".")
            bd_monta_decimal = s
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            s = Format$(valor, s_mascara)
          ' SE SEPARADOR DECIMAL FOR VÍRGULA, SUBSTITUI POR PONTO
            s = substitui_caracteres(s, ",", ".")
            bd_monta_decimal = s
            

        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            s = Format$(valor, s_mascara)
          ' SE SEPARADOR DECIMAL FOR VÍRGULA, SUBSTITUI POR PONTO
            s = substitui_caracteres(s, ",", ".")
            bd_monta_decimal = s
            
        End Select
        
    
    If bd_monta_decimal <> "" Then bd_monta_decimal = " " & Trim$(bd_monta_decimal)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_DECIMAL_TRATA_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   
        

End Function


Function bd_monta_numero(ByVal valor As Variant) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA UM NÚMERO (DECIMAL OU INTEIRO) EM UM FORMATO PRÓPRIO
'|    PARA UTILIZAR EM STRINGS DE CONSULTA SQL NO SERVIDOR DE BD EM USO.
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Const ID_SEPARADOR_TEMP = "V"
Dim i As Integer
Dim s As String
Dim c As String
Dim s_valor As String
Dim s_separador_sistema As String

    On Error GoTo BD_MONTA_NUMERO_TRATA_ERRO
    
    
    bd_monta_numero = ""

    
    If Not IsNumeric(valor) Then Exit Function
        
'   FORÇA CONVERSÃO PARA NÚMERO
    valor = valor * 1
        
'   DETERMINA O SEPARADOR DECIMAL DO SISTEMA
    s = CStr(0.5)
    s_separador_sistema = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then
            s_separador_sistema = c
            Exit For
            End If
        Next

    If s_separador_sistema = "" Then Exit Function
    
    
  ' SUBSTITUI O SEPARADOR DECIMAL POR UM CARACTER ESPECIAL
    s = CStr(valor)
    s = substitui_caracteres(s, s_separador_sistema, ID_SEPARADOR_TEMP)
    s_valor = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If IsNumeric(c) Or (c = "-") Or (c = ID_SEPARADOR_TEMP) Then s_valor = s_valor & c
        Next

    
  
  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            s = substitui_caracteres(s_valor, ID_SEPARADOR_TEMP, ".")
            bd_monta_numero = s
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            s = substitui_caracteres(s_valor, ID_SEPARADOR_TEMP, ".")
            bd_monta_numero = s
            

        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            s = substitui_caracteres(s_valor, ID_SEPARADOR_TEMP, ".")
            bd_monta_numero = s
            
        End Select
        
    
    If bd_monta_numero <> "" Then bd_monta_numero = " " & Trim$(bd_monta_numero)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_NUMERO_TRATA_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   
        

End Function



Function bd_monta_month(campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A FUNÇÃO MONTH PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO
'|    (A FUNÇÃO MONTH RETORNA UM NÚMERO DE 1-12 REFERENTE AO MÊS DA DATA INDICADA).
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_MONTH_ERRO
    
    
    bd_monta_month = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_month = "MONTH(" & campo & ")"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_month = "DATEPART(month, " & campo & ")"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_month = "TO_NUMBER(TO_CHAR(" & campo & ", 'MM'))"
        
        End Select
        
    
    If bd_monta_month <> "" Then bd_monta_month = " " & Trim$(bd_monta_month)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_MONTH_ERRO:
'~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function

Function bd_monta_sum(campo As String, nome_alias As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A FUNÇÃO SUM (INCLUSIVE COM UM ALIAS, CASO SEJA ESPECIFICADO)
'|    PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO.
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_SUM_ERRO
    
    
    bd_monta_sum = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_sum = "SUM(" & campo & ")"
            If Trim$(nome_alias) <> "" Then bd_monta_sum = bd_monta_sum & " AS " & Trim$(nome_alias)
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_sum = "SUM(" & campo & ")"
            If Trim$(nome_alias) <> "" Then bd_monta_sum = bd_monta_sum & " AS " & Trim$(nome_alias)


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_sum = "SUM(" & campo & ")"
            If Trim$(nome_alias) <> "" Then bd_monta_sum = bd_monta_sum & " AS " & Trim$(nome_alias)
        
        End Select
        
    
    If bd_monta_sum <> "" Then bd_monta_sum = " " & Trim$(bd_monta_sum)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_SUM_ERRO:
'~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function

Function bd_monta_year(campo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A FUNÇÃO YEAR PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO
'|    (A FUNÇÃO YEAR RETORNA APENAS O ANO DA DATA ESPECIFICADA COMO UM NÚMERO: 1999,2000,ETC).
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_YEAR_ERRO
    
    
    bd_monta_year = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_year = "YEAR(" & campo & ")"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_year = "DATEPART(year, " & campo & ")"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_year = "TO_NUMBER(TO_CHAR(" & campo & ", 'YYYY'))"
        
        End Select
        
    
    If bd_monta_year <> "" Then bd_monta_year = " " & Trim$(bd_monta_year)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_YEAR_ERRO:
'~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   

End Function

Function bd_obtem_mes(ByVal i As Variant, por_extenso As Boolean) As String
' ________________________________________________________________________________
'|
'|  RETORNA O MÊS COM OS 3 PRIMEIROS CARACTERES E EM INGLÊS
'|

Dim s As String
Dim j As Integer

    On Error Resume Next

    If IsNumeric(i) Then j = CInt(i) Else j = 0

    Select Case j
        Case 1: s = "JANUARY"
        Case 2: s = "FEBRUARY"
        Case 3: s = "MARCH"
        Case 4: s = "APRIL"
        Case 5: s = "MAY"
        Case 6: s = "JUNE"
        Case 7: s = "JULY"
        Case 8: s = "AUGUST"
        Case 9: s = "SEPTEMBER"
        Case 10: s = "OCTOBER"
        Case 11: s = "NOVEMBER"
        Case 12: s = "DECEMBER"
        Case Else: s = ""
        End Select

    
    If Not por_extenso Then s = Mid$(s, 1, 3)
    
    bd_obtem_mes = s


End Function


Function bd_monta_data(data As Variant, inclui_hora As Boolean) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A DATA EM UM FORMATO PRÓPRIO PARA COMPARAÇÕES NA CLÁUSULA 'WHERE'
'|    PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO.
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String
Dim s_aux As String


    On Error GoTo BD_MONTA_DATA_ERRO
    
    
    bd_monta_data = ""

    
    If Not IsDate(data) Then
        bd_monta_data = " NULL"
        Exit Function
        End If
        
    

  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            s = Format$(data, "mm/dd/yyyy")
            If inclui_hora Then s = s & " " & Format$(data, "hh:mm:ss AM/PM")
            s = "#" & s & "#"
            bd_monta_data = s
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            s = Format$(data, "mm dd yyyy")
            s_aux = Mid$(s, 1, InStr(1, s, " ") - 1)
            s_aux = bd_obtem_mes(s_aux, False)
            s = s_aux & Mid$(s, InStr(1, s, " "), Len(s))
            If inclui_hora Then s = s & " " & Format$(data, "hh:mm:ss AM/PM")
            bd_monta_data = "'" & s & "'"
            

        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            s = Format$(data, "mm-dd-yyyy")
            If inclui_hora Then s = s & " " & Format$(data, "hh:mm:ss")
            s = "TO_DATE('" & s & "', 'MM-DD-YYYY"
            If inclui_hora Then s = s & " HH24:MI:SS"
            s = s & "')"
            bd_monta_data = s
            
        End Select
        
    
    If bd_monta_data <> "" Then bd_monta_data = " " & Trim$(bd_monta_data)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_DATA_ERRO:
'~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   
        

End Function


Function bd_monta_left(campo As String, n As Integer) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A FUNÇÃO LEFT (RETORNA OS 'N' PRIMEIROS CARACTERES À ESQUERDA)
'|    PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO.
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_LEFT_ERRO
    
    
    bd_monta_left = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_left = "LEFT$(" & campo & ", " & CStr(n) & ")"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_left = "SUBSTRING(" & campo & ", 1, " & CStr(n) & ")"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_left = "SUBSTR(" & campo & ", 1, " & CStr(n) & ")"
        
        End Select
        
    
    If bd_monta_left <> "" Then bd_monta_left = " " & Trim$(bd_monta_left)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_LEFT_ERRO:
'~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   
        

End Function


Sub bd_desaloca_recordset(t As ADODB.Recordset, libera_variavel As Boolean)
' _________________________________________________________________________________________________
'|
'|  DESALOCA A MEMÓRIA ASSOCIADA AO RECORDSET 'T'.
'|

Dim s As String

    On Error GoTo BD_DESALOCA_RECORDSET_ERRO
    
    If Not (t Is Nothing) Then
        If t.State <> adStateClosed Then t.Close
        If libera_variavel Then Set t = Nothing
        End If
        

Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_DESALOCA_RECORDSET_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    If s <> "" Then s = "Erro ao fechar recordset" & vbCrLf & s
    
    aguarde INFO_NORMAL, m_id
    aviso s

    Exit Sub
        

End Sub

Function sqlMontaGetdateSomenteData() As String

    sqlMontaGetdateSomenteData = "Convert(varchar(10), getdate(), 121)"
    
End Function

Function substitui_caracteres(Texto As String, antigo As String, novo As String) As String
' _________________________________________________________________________________________________
'|
'|  _ SUBSTITUI O CARACTER ESPECIFICADO PELO NOVO
'|

Dim i As Integer
Dim s As String

    
    On Error GoTo SUBSTITUI_CARACTERES_ERRO
    
    
    substitui_caracteres = ""
    
    s = ""
    
    For i = 1 To Len(Texto)
        If Mid$(Texto, i, 1) = antigo Then
           If novo <> "" Then If Asc(novo) <> 0 Then s = s & novo
        Else
           s = s & Mid$(Texto, i, 1)
           End If
        Next
    
    substitui_caracteres = s

Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
SUBSTITUI_CARACTERES_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
    

End Function
Function repete_caracteres(ByVal texto_base As String, ByVal n_vezes As Integer) As String
' _________________________________________________________________________________________________
'|
'|  _ REPETE 'N_VEZES' O CONTEÚDO DE 'TEXTO_BASE'.
'|

Dim s As String
Dim i As Integer


    On Error GoTo REPETE_CARACTERES_ERRO
   
   
    repete_caracteres = ""

    s = ""
    
    For i = 1 To n_vezes
        s = s & texto_base
        Next
    
    
    repete_caracteres = s
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REPETE_CARACTERES_ERRO:
'~~~~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
    

End Function

Sub bd_monta_join(v() As TIPO_PARAMETRO_JOIN, resp_from As String, resp_where As String)
' _________________________________________________________________________________________________
'|
'|  _ SUPONDO UMA CONSULTA COM AS TABELAS T_1, T_2, T_3, T_4 DO SEGUINTE MODO:
'|    T_1 INNER JOIN T_2 ON T_1.A = T_2.A LEFT OUTER JOIN T_3 ON T_2.B = T_3.B
'|    RIGHT OUTER JOIN T_4 ON T_3.C = T_4.C
'|  _ O NOME DA TABELA T_1 DEVE SER PASSADO NA PRIMEIRA POSIÇÃO DO VETOR V(), SEM PRECISAR
'|    PREENCHER OS CAMPOS QUE DEFINEM O TIPO DE JOIN, ETC.
'|  _ AS DEMAIS TABELAS DEVEM SER PASSADAS ATRAVÉS DAS POSIÇÕES SUBSEQUENTES DO VETOR V(),
'|    INDICANDO O  NOME DA TABELA, O TIPO DE JOIN A SER FEITO E OS CAMPOS USADOS PARA
'|    ESTABELECER O RELACIONAMENTO ENTRE AS TABELAS.
'|  _ OS CAMPOS QUE ESTABELECEM O RELACIONAMENTO ENTRE AS TABELAS DEVEM SER PASSADOS
'|    EM V().CAMPOS_JOIN().CAMPO_TABELA_LEFT E V().CAMPOS_JOIN().CAMPO_TABELA_RIGHT
'|    IMPORTANTE: É FUNDAMENTAL QUE ESTES CAMPOS ESTEJAM PREENCHIDOS NA ORDEM CORRETA,
'|    SENÃO O "LEFT JOIN" E O "RIGHT JOIN" PODERÃO FICAR ERRADOS.
'|  _ AS VARIÁVEIS DE RETORNO ESTARÃO VAZIAS OU COM STRINGS COM UM ESPAÇO EM BRANCO
'|    À ESQUERDA (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|  _ LEMBRE-SE:
'|    INNER JOIN: INCLUI SOMENTE OS REGISTROS QUE SATISFAÇAM AOS CRITÉRIOS DO JOIN
'|        NAS DUAS TABELAS.
'|    LEFT JOIN: INCLUI TODOS OS REGISTROS DA TABELA DA ESQUERDA, MESMO QUE NÃO
'|        EXISTAM REGISTROS CORRESPONDENTES NA TABELA DA DIREITA.
'|    RIGHT JOIN: INCLUI TODOS OS REGISTROS DA TABELA DA DIREITA, MESMO QUE NÃO
'|        EXISTAM REGISTROS CORRESPONDENTES NA TABELA DA ESQUERDA.
'|

Dim ic As Integer
Dim i As Integer
Dim s As String


    On Error GoTo BD_MONTA_JOIN_ERRO
    

    resp_from = ""
    resp_where = ""
       
       
       
  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            
            resp_from = Trim$(v(LBound(v)).nome_tabela)
            
            For ic = LBound(v) + 1 To UBound(v)
                
                If Trim$(v(ic).nome_tabela) <> "" Then
                
                    resp_from = "(" & resp_from
                    
                    Select Case v(ic).tipo_join
                        Case BD_INNER_JOIN
                            resp_from = resp_from & " INNER JOIN "
                        Case BD_LEFT_JOIN
                            resp_from = resp_from & " LEFT JOIN "
                        Case BD_RIGHT_JOIN
                            resp_from = resp_from & " RIGHT JOIN "
                            
                        End Select
                    
                    resp_from = resp_from & v(ic).nome_tabela & " ON "
                    
                    For i = LBound(v(ic).campos_join) To UBound(v(ic).campos_join)
                        If Trim$(v(ic).campos_join(i).campo_tabela_left) = "" Then Exit For
                        If i > LBound(v(ic).campos_join) Then resp_from = resp_from & " AND "
                        resp_from = resp_from & "(" & Trim$(v(ic).campos_join(i).campo_tabela_left) & _
                                                "=" & Trim$(v(ic).campos_join(i).campo_tabela_right) & ")"
                        Next
                    
                    resp_from = resp_from & ")"
                
                    End If
                    
                Next
            
            
            
            
            
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            
            resp_from = Trim$(v(LBound(v)).nome_tabela)
            
            For ic = LBound(v) + 1 To UBound(v)
                
                If Trim$(v(ic).nome_tabela) <> "" Then

                    resp_from = "(" & resp_from
                    
                    Select Case v(ic).tipo_join
                        Case BD_INNER_JOIN
                            resp_from = resp_from & " INNER JOIN "
                        Case BD_LEFT_JOIN
                            resp_from = resp_from & " LEFT OUTER JOIN "
                        Case BD_RIGHT_JOIN
                            resp_from = resp_from & " RIGHT OUTER JOIN "
                            
                        End Select
                    
                    resp_from = resp_from & v(ic).nome_tabela & " ON "
                    
                    For i = LBound(v(ic).campos_join) To UBound(v(ic).campos_join)
                        If Trim$(v(ic).campos_join(i).campo_tabela_left) = "" Then Exit For
                        If i > LBound(v(ic).campos_join) Then resp_from = resp_from & " AND "
                        resp_from = resp_from & "(" & Trim$(v(ic).campos_join(i).campo_tabela_left) & _
                                                "=" & Trim$(v(ic).campos_join(i).campo_tabela_right) & ")"
                        Next
                    
                    resp_from = resp_from & ")"
                    
                    End If
                    
                Next
            
            
            
            
            
        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            
            For ic = LBound(v) To UBound(v)
                If Trim$(v(ic).nome_tabela) <> "" Then
                    If Trim$(resp_from) <> "" Then resp_from = resp_from & ", "
                    resp_from = resp_from & v(ic).nome_tabela
                    End If
                Next
                
            resp_from = resp_from & " "
                
                
            For ic = LBound(v) + 1 To UBound(v)
                
                If Trim$(v(ic).nome_tabela) <> "" Then
                    
                    Select Case v(ic).tipo_join
                        Case BD_INNER_JOIN
                            For i = LBound(v(ic).campos_join) To UBound(v(ic).campos_join)
                                If Trim$(v(ic).campos_join(i).campo_tabela_left) = "" Then Exit For
                                If Trim$(resp_where) <> "" Then resp_where = resp_where & " AND "
                                resp_where = resp_where & Trim$(v(ic).campos_join(i).campo_tabela_left) & _
                                                          "=" & Trim$(v(ic).campos_join(i).campo_tabela_right)
                                Next
                        
                        Case BD_LEFT_JOIN
                            For i = LBound(v(ic).campos_join) To UBound(v(ic).campos_join)
                                If Trim$(v(ic).campos_join(i).campo_tabela_left) = "" Then Exit For
                                If Trim$(resp_where) <> "" Then resp_where = resp_where & " AND "
                                resp_where = resp_where & Trim$(v(ic).campos_join(i).campo_tabela_left) & _
                                                          "=" & Trim$(v(ic).campos_join(i).campo_tabela_right) & " (+)"
                                Next
                        
                        Case BD_RIGHT_JOIN
                            For i = LBound(v(ic).campos_join) To UBound(v(ic).campos_join)
                                If Trim$(v(ic).campos_join(i).campo_tabela_left) = "" Then Exit For
                                If Trim$(resp_where) <> "" Then resp_where = resp_where & " AND "
                                resp_where = resp_where & Trim$(v(ic).campos_join(i).campo_tabela_left) & " (+) " & _
                                                          "=" & Trim$(v(ic).campos_join(i).campo_tabela_right)
                                Next
                            
                        End Select
                    
                    End If
                
                Next
                  
        End Select
       
       
       
    If resp_from <> "" Then resp_from = " " & Trim$(resp_from)
    If resp_where <> "" Then resp_where = " " & Trim$(resp_where)
                
       
Exit Sub







'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_JOIN_ERRO:
'==================
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Sub
   
   
End Sub

Function bd_monta_right(campo As String, n As Integer) As String
' _________________________________________________________________________________________________
'|
'|  _ FUNÇÃO QUE RETORNA A FUNÇÃO RIGHT (RETORNA OS 'N' ÚLTIMOS CARACTERES À DIREITA)
'|    PARA CONSULTAS SQL NO SERVIDOR DE BD EM USO.
'|  _ A STRING DE RETORNO ESTARÁ EM BRANCO OU COM UM ESPAÇO EM BRANCO À ESQUERDA
'|    (LÓGICA: CADA SUBSTRING A SER CONCATENADA DEVE INSERIR SEU PRÓPRIO
'|    ESPAÇO EM BRANCO).
'|

Dim s As String


    On Error GoTo BD_MONTA_RIGHT_ERRO
    
    
    bd_monta_right = ""


  ' MONTA SQL ESPECÍFICO PARA CADA SERVIDOR DE BD
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Select Case BD_TIPO_SERVIDOR
        
        Case BD_SERVIDOR_ACCESS
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_right = "RIGHT$(" & campo & ", " & CStr(n) & ")"
            
        
        Case BD_SERVIDOR_SQLSERVER
       '~~~~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_right = "RIGHT(" & campo & ", " & CStr(n) & ")"


        Case BD_SERVIDOR_ORACLE
       '~~~~~~~~~~~~~~~~~~~~~~~
            bd_monta_right = "SUBSTR(" & campo & ", -" & CStr(n) & ")"
        
        End Select
        
    
    If bd_monta_right <> "" Then bd_monta_right = " " & Trim$(bd_monta_right)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BD_MONTA_RIGHT_ERRO:
'~~~~~~~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    
    aguarde ".", m_id
    aviso s

    Exit Function
   
        
End Function


