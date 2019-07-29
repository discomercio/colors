Attribute VB_Name = "mod_BD"
Option Explicit

  ' CONEXÃO AO BANCO DE DADOS
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~
    Global dbc As ADODB.Connection

  ' CONEXÃO AO BANCO DE DADOS DE ASSISTÊNCIA TÉCNICA
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Global dbcAssist As ADODB.Connection
  
  ' CONEXÃO AO BANCO DE DADOS DE CEP
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Global dbcCep As ADODB.Connection
  
  ' CONSTANTES PARA USAR COM O BANCO DE DADOS
    Const BD_DATA_NULA = "DEC 30 1899"


    Type TIPO_LOG_VIA_VETOR
        nome As String
        valor As String
        End Type


Function atualiza_NFe_imagem_com_retorno_NFe_T1(ByVal lngNsuNFeImagem As Long, _
                                                ByVal codigo_retorno_NFe_T1 As String, _
                                                ByVal msg_retorno_NFe_T1 As String, _
                                                ByRef strMsgErro As String) As Boolean
' DECLARAÇÕES
Const NomeDestaRotina = "atualiza_NFe_imagem_com_retorno_NFe_T1()"
Dim s As String
Dim t_NFe_IMAGEM As ADODB.Recordset
    
    On Error GoTo ATUALIZ_NFE_IMAGEM_TRATA_ERRO
    
    atualiza_NFe_imagem_com_retorno_NFe_T1 = False
    
    strMsgErro = ""
    
  ' T_NFE_IMAGEM
    Set t_NFe_IMAGEM = New ADODB.Recordset
    With t_NFe_IMAGEM
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
        
    s = "SELECT * FROM t_NFe_IMAGEM WHERE (id = " & CStr(lngNsuNFeImagem) & ")"
    If t_NFe_IMAGEM.State <> adStateClosed Then t_NFe_IMAGEM.Close
    t_NFe_IMAGEM.Open s, dbc, , , adCmdText
    If t_NFe_IMAGEM.EOF Then
        strMsgErro = "Não foi encontrado o registro em t_NFe_IMAGEM com id=" & CStr(lngNsuNFeImagem) & "!!"
        GoSub ATUALIZ_NFE_IMAGEM_FECHA_TABELAS
        aviso_erro strMsgErro
        Exit Function
        End If
        
    t_NFe_IMAGEM("codigo_retorno_NFe_T1") = codigo_retorno_NFe_T1
    t_NFe_IMAGEM("msg_retorno_NFe_T1") = msg_retorno_NFe_T1
    t_NFe_IMAGEM.Update
    
    GoSub ATUALIZ_NFE_IMAGEM_FECHA_TABELAS
    
    atualiza_NFe_imagem_com_retorno_NFe_T1 = True
    
Exit Function
    
    
    
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ATUALIZ_NFE_IMAGEM_TRATA_ERRO:
'=============================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    strMsgErro = s
    GoSub ATUALIZ_NFE_IMAGEM_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ATUALIZ_NFE_IMAGEM_FECHA_TABELAS:
'================================
  ' RECORDSETS
    bd_desaloca_recordset t_NFe_IMAGEM, True
    Return
    
End Function


Function consiste_municipio_IBGE_ok(ByRef dbcNFe As ADODB.Connection, ByVal municipio As String, ByVal uf As String, ByRef lista_sugerida_municipios As String, ByRef msg_erro As String) As Boolean
Const NomeDestaRotina = "consiste_municipio_IBGE_ok()"
Dim s As String
Dim strCodUF As String
Dim tUF As ADODB.Recordset
Dim tMunicipio As ADODB.Recordset

    On Error GoTo CMIBGEOK_TRATA_ERRO
    
    consiste_municipio_IBGE_ok = False
    lista_sugerida_municipios = ""
    msg_erro = ""
    
'   CONSISTE PARÂMETROS
    If Trim$("" & municipio) = "" Then
        msg_erro = "Não é possível consistir o município através da relação de municípios do IBGE: nenhum município foi informado!!"
        Exit Function
        End If
    
    If Trim$("" & uf) = "" Then
        msg_erro = "Não é possível consistir o município através da relação de municípios do IBGE: a UF não foi informada!!"
        Exit Function
        End If
        
    If Not UF_ok(uf) Then
        msg_erro = "Não é possível consistir o município através da relação de municípios do IBGE: a UF é inválida (" & uf & ")!!"
        Exit Function
        End If
    
'   CRIA OBJETOS DE ACESSO AO BD
    Set tUF = New ADODB.Recordset
    With tUF
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
    Set tMunicipio = New ADODB.Recordset
    With tMunicipio
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   OBTÉM O CÓDIGO DA UF
    s = "SELECT " & _
            "*" & _
        " FROM NFE_UF" & _
        " WHERE" & _
            " (SiglaUF = '" & UCase(uf) & "')"
    tUF.Open s, dbcNFe, , , adCmdText
    If tUF.EOF Then
        msg_erro = "Não é possível consistir o município através da relação de municípios do IBGE: a UF '" & uf & "' não foi localizada na relação do IBGE!!"
        GoSub CMIBGEOK_FECHA_TABELAS
        Exit Function
        End If
    
    strCodUF = Trim("" & tUF("CodUF"))
    
    s = "SELECT " & _
            "*" & _
        " FROM NFE_MUNICIPIO" & _
        " WHERE" & _
            " (CodMunic LIKE '" & strCodUF & BD_CURINGA_TODOS & "')" & _
            " AND (Descricao = '" & bd_filtra_aspas(municipio) & "' COLLATE Latin1_General_CI_AI)"
    tMunicipio.Open s, dbcNFe, , , adCmdText
    If Not tMunicipio.EOF Then
    '   ACHOU O MUNICÍPIO NA LISTA!!
        consiste_municipio_IBGE_ok = True
        GoSub CMIBGEOK_FECHA_TABELAS
        Exit Function
        End If
    
'   NÃO ENCONTROU O MUNICÍPIO, ENTÃO MONTA UMA LISTA DE SUGESTÕES C/ OS POSSÍVEIS MUNICÍPIOS
'   SERÃO DADOS COMO SUGESTÃO TODOS OS MUNICÍPIOS DA UF QUE SE INICIEM C/ A MESMA LETRA DO MUNICÍPIO INFORMADO
    s = "SELECT " & _
            "*" & _
        " FROM NFE_MUNICIPIO" & _
        " WHERE" & _
            " (CodMunic LIKE '" & strCodUF & BD_CURINGA_TODOS & "')" & _
            " AND (Descricao LIKE '" & left(municipio, 1) & BD_CURINGA_TODOS & "' COLLATE Latin1_General_CI_AI)" & _
        " ORDER BY" & _
            " Descricao"
    If tMunicipio.State <> adStateClosed Then tMunicipio.Close
    tMunicipio.Open s, dbcNFe, , , adCmdText
    Do While Not tMunicipio.EOF
        If lista_sugerida_municipios <> "" Then lista_sugerida_municipios = lista_sugerida_municipios & Chr(13)
        lista_sugerida_municipios = lista_sugerida_municipios & Trim("" & tMunicipio("Descricao"))
        tMunicipio.MoveNext
        Loop
    
    GoSub CMIBGEOK_FECHA_TABELAS
    
Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CMIBGEOK_TRATA_ERRO:
'===================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    msg_erro = s
    GoSub CMIBGEOK_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CMIBGEOK_FECHA_TABELAS:
'======================
  ' RECORDSETS
    bd_desaloca_recordset tUF, True
    bd_desaloca_recordset tMunicipio, True
    Return

End Function

Public Function geraNsu(ByVal idNsu As String, ByRef nsu As Long, ByRef strMsgErro As String) As Boolean
' ______________________________________________________________________________
'|
'|  Gera o NSU para a chave informada
'|
'|  Parâmetros:
'|      idNsu: identificação da chave para gerar o NSU, normalmente é o próprio nome da tabela para a qual se deseja gerar o NSU para se usar como ID
'|      nsu: retorna o NSU gerado
'|      strMsgErro: retorna a mensagem de erro em caso de exception
'|
'|  Retorno da função:
'|      true: sucesso ao gerar o NSU
'|      false: falha ao gerar o NSU
'|

Const MAX_TENTATIVAS = 10
Dim intQtdeTentativas As Integer
Dim blnSucesso As Boolean
Dim lngRetorno As Long
Dim lngNsuUltimo As Long
Dim lngNsuNovo As Long
Dim lngRecordsAffected As Long
Dim t As ADODB.Recordset
Dim strSql As String
    
    On Error GoTo GN_TRATA_ERRO
    
    geraNsu = False
    strMsgErro = ""
    
'   RECORDSET
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
    strSql = "SELECT" & _
                " Count(*) AS qtde" & _
            " FROM t_FIN_CONTROLE" & _
            " WHERE" & _
                " (id='" & idNsu & "')"
    If t.State <> adStateClosed Then t.Close
    t.Open strSql, dbc, , , adCmdText
    If Not t.EOF Then lngRetorno = t("qtde")
    
'   NÃO ESTÁ CADASTRADO, ENTÃO CADASTRA AGORA
    If lngRetorno = 0 Then
        strSql = "INSERT INTO t_FIN_CONTROLE (" & _
                    "id, " & _
                    "nsu, " & _
                    "dt_hr_ult_atualizacao" & _
                ") VALUES (" & _
                    "'" & idNsu & "'," & _
                    "0," & _
                    "getdate()" & _
                ")"
        Call dbc.Execute(strSql, lngRecordsAffected)
        If lngRecordsAffected <> 1 Then
            strMsgErro = "Falha ao criar o registro para geração de NSU (" & idNsu & ")!!"
            Exit Function
            End If
        End If
    
'   LAÇO DE TENTATIVAS PARA GERAR O NSU (DEVIDO A ACESSO CONCORRENTE)
    Do
        intQtdeTentativas = intQtdeTentativas + 1
        
    '   OBTÉM O ÚLTIMO NSU USADO
        strSql = "SELECT" & _
                    " nsu" & _
                " FROM t_FIN_CONTROLE" & _
                " WHERE" & _
                    " id = '" & idNsu & "'"
        If t.State <> adStateClosed Then t.Close
        t.Open strSql, dbc, , , adCmdText
        If t.EOF Then
            strMsgErro = "Falha ao localizar o registro para geração de NSU (" & idNsu & ")!!"
            Exit Function
        Else
            lngNsuUltimo = t("nsu")
            End If
            
    '   INCREMENTA 1
        lngNsuNovo = lngNsuUltimo + 1
        
    '   TENTA ATUALIZAR O BANCO DE DADOS
        strSql = "UPDATE t_FIN_CONTROLE SET" & _
                    " nsu = " & CStr(lngNsuNovo) & _
                " WHERE" & _
                    " (id = '" & idNsu & "')" & _
                    " AND (nsu = " & CStr(lngNsuUltimo) & ")"
        Call dbc.Execute(strSql, lngRecordsAffected)
        If lngRecordsAffected = 1 Then
            blnSucesso = True
            nsu = lngNsuNovo
        Else
            Sleep 100
            End If
            
        Loop While (Not blnSucesso) And (intQtdeTentativas < MAX_TENTATIVAS)
        
    If Not blnSucesso Then
        strMsgErro = "Falha ao tentar gerar o NSU!!"
        GoSub GN_FECHA_TABELAS
        Exit Function
        End If
    
    geraNsu = True
    
    GoSub GN_FECHA_TABELAS
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GN_TRATA_ERRO:
'~~~~~~~~~~~~~
    strMsgErro = CStr(Err) & ": " & Error$(Err)
    GoSub GN_FECHA_TABELAS
    Exit Function
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GN_FECHA_TABELAS:
'================
  ' RECORDSETS
    bd_desaloca_recordset t, True
    Return
    
End Function


Function grava_log(ByVal usuario As String, ByVal loja As String, ByVal pedido As String, ByVal id_cliente As String, ByVal operacao As String, ByVal complemento As String)
' ___________________________________________
' GRAVA LOG
'
Dim s  As String
Dim msg_erro  As String
Dim rs As ADODB.Recordset
    grava_log = False
    Set rs = New ADODB.Recordset
    rs.CursorType = BD_CURSOR_EDICAO
    rs.LockType = BD_POLITICA_LOCKING
    s = "select * from t_LOG where (data < '" & BD_DATA_NULA & "')"
    rs.Open s, dbc, , , adCmdText
    If Err = 0 Then
        rs.AddNew
        If Err = 0 Then
          ' LEMBRANDO QUE A DATA É INSERIDA, VIA DEFAULT DA COLUNA, COM O VALOR DE getdate()
            rs("usuario") = usuario
            rs("loja") = loja
            rs("pedido") = pedido
            rs("id_cliente") = id_cliente
            rs("operacao") = operacao
            rs("complemento") = complemento
            rs.Update
            If Err = 0 Then grava_log = True
            End If
        End If
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function

Function BD_inicia() As Boolean
' ______________________________________________________________________________
'|
'|  INICIA ACESSO AO BANCO DE DADOS
'|

Dim s As String
Dim t As ADODB.Recordset
Dim v() As String
Dim i As Integer
Dim n_tempo As Double

    On Error GoTo BDI_TRATA_ERRO
    
    BD_inicia = False
    
  ' RECORDSET
    Set t = New ADODB.Recordset
    t.CursorType = BD_CURSOR_SOMENTE_LEITURA
    t.LockType = BD_POLITICA_LOCKING
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ABRE O BANCO DE DADOS
      
    If Not bd_monta_string_conexao_sgbd(s) Then
        aviso_erro s
        Exit Function
        End If
        
    Set dbc = New ADODB.Connection
    dbc.CursorLocation = BD_POLITICA_CURSOR
    dbc.ConnectionTimeout = BD_CONNECTION_TIMEOUT
    dbc.CommandTimeout = BD_COMMAND_TIMEOUT
    dbc.Open BD_STRING_CONEXAO_SERVIDOR
   

  ' COMANDOS DE INICIALIZAÇÃO
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~
  ' ESTES COMANDOS DEVEM SER EXECUTADOS SEMPRE QUE SE ABRIR UMA CONEXÃO COM O BD !!
    If bd_comandos_inicializacao(v()) Then
        For i = LBound(v) To UBound(v)
            s = Trim$(v(i))
            If s <> "" Then
                If t.State <> adStateClosed Then t.Close
                t.Open s, dbc, , , adCmdText
                End If
            Next
        End If


  ' SINCRONIZA DATA/HORA C/ SERVIDOR
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' RECORDSET
    If t.State <> adStateClosed Then t.Close
    Set t = Nothing
    
    Set t = New ADODB.Recordset
    t.CursorType = adOpenDynamic
    t.LockType = adLockPessimistic

    s = bd_monta_getdate("data_sistema")
    
    If t.State <> adStateClosed Then t.Close
    t.Open s, dbc, , , adCmdText
    If Not t.EOF Then
      ' SE PROFILE IMPEDE ALTERAÇÃO DE DATA/HORA, PROSSEGUE MESMO ASSIM
        On Error Resume Next
        Time = t("data_sistema")
        Date = t("data_sistema")
      ' RESTAURA TRATAMENTO DE ERRO
        On Error GoTo BDI_TRATA_ERRO
        
      ' SE NÃO CONSEGUIU ACERTAR HORÁRIO, VERIFICA DIFERENÇA DE HORÁRIO C/ SERVIDOR
        n_tempo = Now - t("data_sistema")
        If Abs(n_tempo) > (MAX_ERRO_RELOGIO_EM_MINUTOS * (1 / (24 * 60))) Then
            s = "Não foi possível acertar automaticamente o relógio desta estação !!" & _
                vbCrLf & "Não será possível continuar porque o relógio está " & IIf(n_tempo > 0, "adiantado", "atrasado") & _
                " em " & CStr(Fix(Abs(n_tempo) * 24 * 60)) & " minutos !!"
            aviso_erro s
            GoSub BDI_FECHA_TABELAS
          ' ENCERRA MÓDULO
          ' ~~~~~~~~~
            BD_Fecha
            End
          ' ~~~~~~~~~
            End If
        End If


  ' FECHA RECORDSETS
    GoSub BDI_FECHA_TABELAS
   
    BD_inicia = True

Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 
'   BDI_FECHA_TABELAS
' 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BDI_FECHA_TABELAS:
'==================
    bd_desaloca_recordset t, True
    
    Return
    



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 
'   ERRO AO ABRIR OU ACESSAR BD
' 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BDI_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err)
    GoSub BDI_FECHA_TABELAS
    
    BD_Fecha
    
    aviso s
    
'   ENCERRA O PROGRAMA
    End

End Function

Sub BD_Fecha()
' _______________________________________________________________________________________
'|
'|  FECHA AS VARIÁVEIS DE ACESSO AO BANCO DE DADOS
'|

    On Error Resume Next

    If Not (dbc Is Nothing) Then
        If dbc.State <> adStateClosed Then dbc.Close
        Set dbc = Nothing
        End If

End Sub

' FORAM IMPLEMENTADAS FUNÇÕES PARA ABERTURA/FECHAMENTO DO BANCO DE DADOS DE
' ASSISTÊNCIA TÉCNICA NOS MOLDES DAS FUNÇÕES DO SGBD, PARA:
' 1 - EVITAR CONSTANTE ABERTURA/FECHAMENTO DA CONEXÃO COM O MESMO, VISTO QUE PODEM SER
'     SOLICITADAS VÁRIAS INFORMAÇÕES DE ASSISTÊNCIA DENTRO DA MESMA SESSÃO DO NF;
' 2 - PERMITIR QUE AS MESMAS APENAS SEJAM CHAMADAS NO CASO DE UTILIZAÇÃO DAS INFORMAÇÕES
'     DE ASSISTÊNCIA

Function BD_Assist_inicia() As Boolean
' ______________________________________________________________________________
'|
'|  INICIA ACESSO AO BANCO DE DADOS DE ASSISTÊNCIA TÉCNICA
'|

Dim s As String

    On Error GoTo BDAI_TRATA_ERRO
    
    BD_Assist_inicia = False
        
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ABRE O BANCO DE DADOS
      
    
    If Not bd_monta_string_conexao_at(s) Then
        aviso_erro s
        Exit Function
        End If
        
    Set dbcAssist = New ADODB.Connection
    dbcAssist.CursorLocation = BD_POLITICA_CURSOR
    dbcAssist.ConnectionTimeout = BD_CONNECTION_TIMEOUT
    dbcAssist.CommandTimeout = BD_COMMAND_TIMEOUT
    dbcAssist.Open BD_STRING_CONEXAO_SERVIDOR_AT
      
    BD_Assist_inicia = True

Exit Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 
'   ERRO AO ABRIR OU ACESSAR BD
' 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BDAI_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err)
    
    BD_Assist_Fecha
    
    aviso s
    
End Function

' FORAM IMPLEMENTADAS FUNÇÕES PARA ABERTURA/FECHAMENTO DO BANCO DE DADOS DE
' CEP NOS MOLDES DAS FUNÇÕES DO SGBD, PARA:
' 1 - EVITAR CONSTANTE ABERTURA/FECHAMENTO DA CONEXÃO COM O MESMO, VISTO QUE PODEM SER
'     SOLICITADAS VÁRIAS INFORMAÇÕES DE CEP DENTRO DA MESMA SESSÃO DO NF;

Function BD_CEP_inicia() As Boolean
' ______________________________________________________________________________
'|
'|  INICIA ACESSO AO BANCO DE DADOS DE CEP
'|

Dim s As String

    On Error GoTo BDCI_TRATA_ERRO
    
    BD_CEP_inicia = False
        
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ABRE O BANCO DE DADOS
      
    
    If Not bd_monta_string_conexao_cep(s) Then
        aviso_erro s
        Exit Function
        End If
        
    Set dbcCep = New ADODB.Connection
    dbcCep.CursorLocation = BD_POLITICA_CURSOR
    dbcCep.ConnectionTimeout = BD_CONNECTION_TIMEOUT
    dbcCep.CommandTimeout = BD_COMMAND_TIMEOUT
    dbcCep.Open BD_STRING_CONEXAO_SERVIDOR_CEP
      
    BD_CEP_inicia = True

Exit Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 
'   ERRO AO ABRIR OU ACESSAR BD
' 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BDCI_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err)
    
    BD_CEP_Fecha
    
    aviso s
    
End Function


Sub BD_Assist_Fecha()
' _______________________________________________________________________________________
'|
'|  FECHA AS VARIÁVEIS DE ACESSO AO BANCO DE DADOS DE ASSISTÊNCIA TÉCNICA
'|

    On Error Resume Next

    If Not (dbcAssist Is Nothing) Then
        If dbcAssist.State <> adStateClosed Then dbcAssist.Close
        Set dbcAssist = Nothing
        End If

End Sub


Sub BD_CEP_Fecha()
' _______________________________________________________________________________________
'|
'|  FECHA AS VARIÁVEIS DE ACESSO AO BANCO DE DADOS DE CEP
'|

    On Error Resume Next

    If Not (dbcCep Is Nothing) Then
        If dbcCep.State <> adStateClosed Then dbcCep.Close
        Set dbcCep = Nothing
        End If

End Sub


Function grava_NFe_imagem(ByVal usuario As String, _
                          ByVal lngSerieNFe As Long, _
                          ByVal lngNumeroNFe As Long, _
                          ByRef rNFeImg As TIPO_NFe_IMG, _
                          ByRef vNFeImgItem() As TIPO_NFe_IMG_ITEM, _
                          ByRef vNFeImgTagDup() As TIPO_NFe_IMG_TAG_DUP, _
                          ByRef vNFeImgNFeRef() As TIPO_NFe_IMG_NFe_REFERENCIADA, _
                          ByRef vNFeImgPag() As TIPO_NFe_IMG_PAG, _
                          ByRef lngNsuNFeImagem As Long, _
                          ByRef strMsgErro As String) As Boolean
' DECLARAÇÕES
Const NomeDestaRotina = "grava_NFe_imagem()"
Dim ic As Integer
Dim intOrdem As Integer
Dim lngNsuNFeImagemItem As Long
Dim lngNsuNFeImagemTagDup As Long
Dim lngNsuNFeImagemNFeReferenciada As Long
Dim lngNsuNFeImagemPag As Long
Dim s As String
Dim t_NFe_IMAGEM As ADODB.Recordset
Dim t_NFe_IMAGEM_ITEM As ADODB.Recordset
Dim t_NFe_IMAGEM_TAG_DUP As ADODB.Recordset
Dim t_NFe_IMAGEM_NFe_REFERENCIADA As ADODB.Recordset
Dim t_NFe_IMAGEM_PAG As ADODB.Recordset
    
    On Error GoTo GRAVA_NFE_IMAGEM_TRATA_ERRO
    
    grava_NFe_imagem = False
    
    lngNsuNFeImagem = 0
    strMsgErro = ""
    
  ' T_NFE_IMAGEM
    Set t_NFe_IMAGEM = New ADODB.Recordset
    With t_NFe_IMAGEM
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
    
  ' T_NFE_IMAGEM_ITEM
    Set t_NFe_IMAGEM_ITEM = New ADODB.Recordset
    With t_NFe_IMAGEM_ITEM
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
    
  ' T_NFE_IMAGEM_TAG_DUP
    Set t_NFe_IMAGEM_TAG_DUP = New ADODB.Recordset
    With t_NFe_IMAGEM_TAG_DUP
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
    
  ' T_NFE_IMAGEM_NFE_REFERENCIADA
    Set t_NFe_IMAGEM_NFe_REFERENCIADA = New ADODB.Recordset
    With t_NFe_IMAGEM_NFe_REFERENCIADA
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
    
  ' T_NFE_IMAGEM_PAG
    Set t_NFe_IMAGEM_PAG = New ADODB.Recordset
    With t_NFe_IMAGEM_PAG
        .CursorType = BD_CURSOR_EDICAO
        .LockType = BD_POLITICA_LOCKING
        End With
    
'   ~~~~~~~~~~~~~~~
    dbc.BeginTrans
'   ~~~~~~~~~~~~~~~
    On Error GoTo GRAVA_NFE_IMAGEM_TRATA_ERRO_TRANSACAO
    
    If Not geraNsu(NSU_T_NFe_IMAGEM, lngNsuNFeImagem, strMsgErro) Then
        If strMsgErro <> "" Then strMsgErro = Chr(13) & Chr(13) & strMsgErro
        strMsgErro = "Falha ao gravar os dados de pagamento!!" & strMsgErro
    '   ~~~~~~~~~~~~~~~~~
        dbc.RollbackTrans
    '   ~~~~~~~~~~~~~~~~~
        GoSub GRAVA_NFE_IMAGEM_FECHA_TABELAS
        aviso_erro strMsgErro
        Exit Function
        End If
    
'   LEMBRANDO QUE OS CAMPOS 'data' E 'data_hora' SÃO PREENCHIDOS AUTOMATICAMENTE POR UM "CONSTRAINT DEFAULT"
    s = "SELECT " & _
            "*" & _
        " FROM t_NFe_IMAGEM" & _
        " WHERE" & _
            " (id = -1)"
    If t_NFe_IMAGEM.State <> adStateClosed Then t_NFe_IMAGEM.Close
    t_NFe_IMAGEM.Open s, dbc, , , adCmdText
    t_NFe_IMAGEM.AddNew
    t_NFe_IMAGEM("id") = lngNsuNFeImagem
    t_NFe_IMAGEM("id_nfe_emitente") = rNFeImg.id_nfe_emitente
    t_NFe_IMAGEM("NFe_serie_NF") = lngSerieNFe
    t_NFe_IMAGEM("NFe_numero_NF") = lngNumeroNFe
    t_NFe_IMAGEM("versao_layout_NFe") = ID_VERSAO_LAYOUT_NFe
    t_NFe_IMAGEM("usuario") = usuario
    t_NFe_IMAGEM("pedido") = rNFeImg.pedido
    t_NFe_IMAGEM("operacional__email") = rNFeImg.operacional__email
    t_NFe_IMAGEM("ide__natOp") = rNFeImg.ide__natOp
    t_NFe_IMAGEM("ide__indPag") = rNFeImg.ide__indPag
    t_NFe_IMAGEM("ide__serie") = rNFeImg.ide__serie
    t_NFe_IMAGEM("ide__nNF") = rNFeImg.ide__nNF
    t_NFe_IMAGEM("ide__dEmi") = rNFeImg.ide__dEmi
    t_NFe_IMAGEM("ide__dEmiUTC") = rNFeImg.ide__dEmiUTC
    t_NFe_IMAGEM("ide__dSaiEnt") = rNFeImg.ide__dSaiEnt
    t_NFe_IMAGEM("ide__tpNF") = rNFeImg.ide__tpNF
    t_NFe_IMAGEM("ide__idDest") = rNFeImg.ide__idDest
    t_NFe_IMAGEM("ide__cMunFG") = rNFeImg.ide__cMunFG
    t_NFe_IMAGEM("ide__tpAmb") = rNFeImg.ide__tpAmb
    t_NFe_IMAGEM("ide__finNFe") = rNFeImg.ide__finNFe
    t_NFe_IMAGEM("ide__indFinal") = rNFeImg.ide__indFinal
    t_NFe_IMAGEM("ide__indPres") = rNFeImg.ide__indPres
    t_NFe_IMAGEM("ide__IEST") = rNFeImg.ide__IEST
    t_NFe_IMAGEM("dest__CNPJ") = rNFeImg.dest__CNPJ
    t_NFe_IMAGEM("dest__CPF") = rNFeImg.dest__CPF
    t_NFe_IMAGEM("dest__xNome") = rNFeImg.dest__xNome
    t_NFe_IMAGEM("dest__xLgr") = rNFeImg.dest__xLgr
    t_NFe_IMAGEM("dest__nro") = rNFeImg.dest__nro
    t_NFe_IMAGEM("dest__xCpl") = rNFeImg.dest__xCpl
    t_NFe_IMAGEM("dest__xBairro") = rNFeImg.dest__xBairro
    t_NFe_IMAGEM("dest__cMun") = rNFeImg.dest__cMun
    t_NFe_IMAGEM("dest__xMun") = rNFeImg.dest__xMun
    t_NFe_IMAGEM("dest__UF") = rNFeImg.dest__UF
    t_NFe_IMAGEM("dest__CEP") = rNFeImg.dest__CEP
    t_NFe_IMAGEM("dest__cPais") = rNFeImg.dest__cPais
    t_NFe_IMAGEM("dest__xPais") = rNFeImg.dest__xPais
    t_NFe_IMAGEM("dest__fone") = rNFeImg.dest__fone
    t_NFe_IMAGEM("dest__IE") = rNFeImg.dest__IE
    t_NFe_IMAGEM("dest__ISUF") = rNFeImg.dest__ISUF
    t_NFe_IMAGEM("dest__idEstrangeiro") = rNFeImg.dest__idEstrangeiro
    t_NFe_IMAGEM("dest__indIEDest") = rNFeImg.dest__indIEDest
    t_NFe_IMAGEM("dest__email") = rNFeImg.dest__email
    t_NFe_IMAGEM("entrega__CNPJ") = rNFeImg.entrega__CNPJ
    t_NFe_IMAGEM("entrega__CPF") = rNFeImg.entrega__CPF
    t_NFe_IMAGEM("entrega__xLgr") = rNFeImg.entrega__xLgr
    t_NFe_IMAGEM("entrega__nro") = rNFeImg.entrega__nro
    t_NFe_IMAGEM("entrega__xCpl") = rNFeImg.entrega__xCpl
    t_NFe_IMAGEM("entrega__xBairro") = rNFeImg.entrega__xBairro
    t_NFe_IMAGEM("entrega__cMun") = rNFeImg.entrega__cMun
    t_NFe_IMAGEM("entrega__xMun") = rNFeImg.entrega__xMun
    t_NFe_IMAGEM("entrega__UF") = rNFeImg.entrega__UF
    t_NFe_IMAGEM("total__vBC") = rNFeImg.total__vBC
    t_NFe_IMAGEM("total__vICMS") = rNFeImg.total__vICMS
    t_NFe_IMAGEM("total__vICMSDeson") = rNFeImg.total__vICMSDeson
    t_NFe_IMAGEM("total__vBCST") = rNFeImg.total__vBCST
    t_NFe_IMAGEM("total__vST") = rNFeImg.total__vST
    If PARTILHA_ICMS_ATIVA Then
        t_NFe_IMAGEM("total__vFCPUFDest") = rNFeImg.total__vFCPUFDest
        t_NFe_IMAGEM("total__vICMSUFDest") = rNFeImg.total__vICMSUFDest
        t_NFe_IMAGEM("total__vICMSUFRemet") = rNFeImg.total__vICMSUFRemet
        End If
    t_NFe_IMAGEM("total__vFCP") = rNFeImg.total__vFCP
    t_NFe_IMAGEM("total__vFCPST") = rNFeImg.total__vFCPST
    t_NFe_IMAGEM("total__vFCPSTRet") = rNFeImg.total__vFCPSTRet
    t_NFe_IMAGEM("total__vIPIDevol") = rNFeImg.total__vIPIDevol
    t_NFe_IMAGEM("total__vProd") = rNFeImg.total__vProd
    t_NFe_IMAGEM("total__vFrete") = rNFeImg.total__vFrete
    t_NFe_IMAGEM("total__vSeg") = rNFeImg.total__vSeg
    t_NFe_IMAGEM("total__vDesc") = rNFeImg.total__vDesc
    t_NFe_IMAGEM("total__vII") = rNFeImg.total__vII
    t_NFe_IMAGEM("total__vIPI") = rNFeImg.total__vIPI
    t_NFe_IMAGEM("total__vPIS") = rNFeImg.total__vPIS
    t_NFe_IMAGEM("total__vCOFINS") = rNFeImg.total__vCOFINS
    t_NFe_IMAGEM("total__vOutro") = rNFeImg.total__vOutro
    t_NFe_IMAGEM("total__vNF") = rNFeImg.total__vNF
    t_NFe_IMAGEM("total__vTotTrib") = rNFeImg.total__vTotTrib
    t_NFe_IMAGEM("transp__modFrete") = rNFeImg.transp__modFrete
    t_NFe_IMAGEM("transporta__CNPJ") = rNFeImg.transporta__CNPJ
    t_NFe_IMAGEM("transporta__CPF") = rNFeImg.transporta__CPF
    t_NFe_IMAGEM("transporta__xNome") = rNFeImg.transporta__xNome
    t_NFe_IMAGEM("transporta__IE") = rNFeImg.transporta__IE
    t_NFe_IMAGEM("transporta__xEnder") = rNFeImg.transporta__xEnder
    t_NFe_IMAGEM("transporta__xMun") = rNFeImg.transporta__xMun
    t_NFe_IMAGEM("transporta__UF") = rNFeImg.transporta__UF
    t_NFe_IMAGEM("vol__qVol") = rNFeImg.vol__qVol
    t_NFe_IMAGEM("vol__esp") = rNFeImg.vol__esp
    t_NFe_IMAGEM("vol__marca") = rNFeImg.vol__marca
    t_NFe_IMAGEM("vol__nVol") = rNFeImg.vol__nVol
    t_NFe_IMAGEM("vol__pesoL") = rNFeImg.vol__pesoL
    t_NFe_IMAGEM("vol__pesoB") = rNFeImg.vol__pesoB
    t_NFe_IMAGEM("vol_nLacre") = rNFeImg.vol_nLacre
    t_NFe_IMAGEM("infAdic__infAdFisco") = rNFeImg.infAdic__infAdFisco
    t_NFe_IMAGEM("infAdic__infCpl") = rNFeImg.infAdic__infCpl
    t_NFe_IMAGEM("codigo_retorno_NFe_T1") = rNFeImg.codigo_retorno_NFe_T1
    t_NFe_IMAGEM("msg_retorno_NFe_T1") = rNFeImg.msg_retorno_NFe_T1
    t_NFe_IMAGEM.Update

'   GRAVA OS ITENS
    intOrdem = 0
    For ic = LBound(vNFeImgItem) To UBound(vNFeImgItem)
        If Trim$(vNFeImgItem(ic).det__nItem) <> "" Then
            intOrdem = intOrdem + 1
            
            If Not geraNsu(NSU_T_NFe_IMAGEM_ITEM, lngNsuNFeImagemItem, strMsgErro) Then
                If strMsgErro <> "" Then strMsgErro = Chr(13) & Chr(13) & strMsgErro
                strMsgErro = "Falha ao gravar os dados da NFe em " & NSU_T_NFe_IMAGEM_ITEM & "!!" & strMsgErro
            '   ~~~~~~~~~~~~~~~~~
                dbc.RollbackTrans
            '   ~~~~~~~~~~~~~~~~~
                GoSub GRAVA_NFE_IMAGEM_FECHA_TABELAS
                aviso_erro strMsgErro
                Exit Function
                End If
            
            s = "SELECT " & _
                    "*" & _
                " FROM t_NFe_IMAGEM_ITEM" & _
                " WHERE" & _
                    " (id = -1)"
            If t_NFe_IMAGEM_ITEM.State <> adStateClosed Then t_NFe_IMAGEM_ITEM.Close
            t_NFe_IMAGEM_ITEM.Open s, dbc, , , adCmdText
            t_NFe_IMAGEM_ITEM.AddNew
            t_NFe_IMAGEM_ITEM("id") = lngNsuNFeImagemItem
            t_NFe_IMAGEM_ITEM("id_nfe_imagem") = lngNsuNFeImagem
            t_NFe_IMAGEM_ITEM("ordem") = intOrdem
            t_NFe_IMAGEM_ITEM("fabricante") = vNFeImgItem(ic).fabricante
            t_NFe_IMAGEM_ITEM("produto") = vNFeImgItem(ic).produto
            t_NFe_IMAGEM_ITEM("det__nItem") = vNFeImgItem(ic).det__nItem
            t_NFe_IMAGEM_ITEM("det__cProd") = vNFeImgItem(ic).det__cProd
            t_NFe_IMAGEM_ITEM("det__cEAN") = vNFeImgItem(ic).det__cEAN
            t_NFe_IMAGEM_ITEM("det__xProd") = vNFeImgItem(ic).det__xProd
            t_NFe_IMAGEM_ITEM("det__NCM") = vNFeImgItem(ic).det__NCM
            If PARTILHA_ICMS_ATIVA Then t_NFe_IMAGEM_ITEM("det__CEST") = vNFeImgItem(ic).det__CEST
            t_NFe_IMAGEM_ITEM("det__indEscala") = vNFeImgItem(ic).det__indEscala
            t_NFe_IMAGEM_ITEM("det__EXTIPI") = vNFeImgItem(ic).det__EXTIPI
            t_NFe_IMAGEM_ITEM("det__genero") = vNFeImgItem(ic).det__genero
            t_NFe_IMAGEM_ITEM("det__CFOP") = vNFeImgItem(ic).det__CFOP
            t_NFe_IMAGEM_ITEM("det__uCom") = vNFeImgItem(ic).det__uCom
            t_NFe_IMAGEM_ITEM("det__qCom") = vNFeImgItem(ic).det__qCom
            t_NFe_IMAGEM_ITEM("det__vUnCom") = vNFeImgItem(ic).det__vUnCom
            t_NFe_IMAGEM_ITEM("det__vProd") = vNFeImgItem(ic).det__vProd
            t_NFe_IMAGEM_ITEM("det__cEANTrib") = vNFeImgItem(ic).det__cEANTrib
            t_NFe_IMAGEM_ITEM("det__uTrib") = vNFeImgItem(ic).det__uTrib
            t_NFe_IMAGEM_ITEM("det__qTrib") = vNFeImgItem(ic).det__qTrib
            t_NFe_IMAGEM_ITEM("det__vUnTrib") = vNFeImgItem(ic).det__vUnTrib
            t_NFe_IMAGEM_ITEM("det__vFrete") = vNFeImgItem(ic).det__vFrete
            t_NFe_IMAGEM_ITEM("det__vSeg") = vNFeImgItem(ic).det__vSeg
            t_NFe_IMAGEM_ITEM("det__vDesc") = vNFeImgItem(ic).det__vDesc
            t_NFe_IMAGEM_ITEM("ICMS__orig") = vNFeImgItem(ic).ICMS__orig
            t_NFe_IMAGEM_ITEM("ICMS__CST") = vNFeImgItem(ic).ICMS__CST
            t_NFe_IMAGEM_ITEM("ICMS__modBC") = vNFeImgItem(ic).ICMS__modBC
            t_NFe_IMAGEM_ITEM("ICMS__pRedBC") = vNFeImgItem(ic).ICMS__pRedBC
            t_NFe_IMAGEM_ITEM("ICMS__vBC") = vNFeImgItem(ic).ICMS__vBC
            t_NFe_IMAGEM_ITEM("ICMS__pICMS") = vNFeImgItem(ic).ICMS__pICMS
            t_NFe_IMAGEM_ITEM("ICMS__vICMS") = vNFeImgItem(ic).ICMS__vICMS
            t_NFe_IMAGEM_ITEM("ICMS__vICMSDeson") = vNFeImgItem(ic).ICMS__vICMSDeson
            t_NFe_IMAGEM_ITEM("ICMS__modBCST") = vNFeImgItem(ic).ICMS__modBCST
            t_NFe_IMAGEM_ITEM("ICMS__pMVAST") = vNFeImgItem(ic).ICMS__pMVAST
            t_NFe_IMAGEM_ITEM("ICMS__pRedBCST") = vNFeImgItem(ic).ICMS__pRedBCST
            t_NFe_IMAGEM_ITEM("ICMS__vBCST") = vNFeImgItem(ic).ICMS__vBCST
            t_NFe_IMAGEM_ITEM("ICMS__pICMSST") = vNFeImgItem(ic).ICMS__pICMSST
            t_NFe_IMAGEM_ITEM("ICMS__vICMSST") = vNFeImgItem(ic).ICMS__vICMSST
            t_NFe_IMAGEM_ITEM("PIS__CST") = vNFeImgItem(ic).PIS__CST
            t_NFe_IMAGEM_ITEM("PIS__vBC") = vNFeImgItem(ic).PIS__vBC
            t_NFe_IMAGEM_ITEM("PIS__pPIS") = vNFeImgItem(ic).PIS__pPIS
            t_NFe_IMAGEM_ITEM("PIS__vPIS") = vNFeImgItem(ic).PIS__vPIS
            t_NFe_IMAGEM_ITEM("PIS__qBCProd") = vNFeImgItem(ic).PIS__qBCProd
            t_NFe_IMAGEM_ITEM("PIS__vAliqProd") = vNFeImgItem(ic).PIS__vAliqProd
            t_NFe_IMAGEM_ITEM("COFINS__CST") = vNFeImgItem(ic).COFINS__CST
            t_NFe_IMAGEM_ITEM("COFINS__vBC") = vNFeImgItem(ic).COFINS__vBC
            t_NFe_IMAGEM_ITEM("COFINS__pCOFINS") = vNFeImgItem(ic).COFINS__pCOFINS
            t_NFe_IMAGEM_ITEM("COFINS__vCOFINS") = vNFeImgItem(ic).COFINS__vCOFINS
            t_NFe_IMAGEM_ITEM("COFINS__qBCProd") = vNFeImgItem(ic).COFINS__qBCProd
            t_NFe_IMAGEM_ITEM("COFINS__vAliqProd") = vNFeImgItem(ic).COFINS__vAliqProd
            t_NFe_IMAGEM_ITEM("IPI__CST") = vNFeImgItem(ic).IPI__CST
            t_NFe_IMAGEM_ITEM("IPI__clEnq") = vNFeImgItem(ic).IPI__clEnq
            t_NFe_IMAGEM_ITEM("IPI__CNPJProd") = vNFeImgItem(ic).IPI__CNPJProd
            t_NFe_IMAGEM_ITEM("IPI__cSelo") = vNFeImgItem(ic).IPI__cSelo
            t_NFe_IMAGEM_ITEM("IPI__qSelo") = vNFeImgItem(ic).IPI__qSelo
            t_NFe_IMAGEM_ITEM("IPI__cEnq") = vNFeImgItem(ic).IPI__cEnq
            t_NFe_IMAGEM_ITEM("IPI__vBC") = vNFeImgItem(ic).IPI__vBC
            t_NFe_IMAGEM_ITEM("IPI__qUnid") = vNFeImgItem(ic).IPI__qUnid
            t_NFe_IMAGEM_ITEM("IPI__vUnid") = vNFeImgItem(ic).IPI__vUnid
            t_NFe_IMAGEM_ITEM("IPI__pIPI") = vNFeImgItem(ic).IPI__pIPI
            t_NFe_IMAGEM_ITEM("IPI__vIPI") = vNFeImgItem(ic).IPI__vIPI
            If PARTILHA_ICMS_ATIVA Then
                t_NFe_IMAGEM_ITEM("ICMSUFDest__vBCUFDest") = vNFeImgItem(ic).ICMSUFDest__vBCUFDest
                t_NFe_IMAGEM_ITEM("ICMSUFDest__pFCPUFDest") = vNFeImgItem(ic).ICMSUFDest__pFCPUFDest
                t_NFe_IMAGEM_ITEM("ICMSUFDest__pICMSUFDest") = vNFeImgItem(ic).ICMSUFDest__pICMSUFDest
                t_NFe_IMAGEM_ITEM("ICMSUFDest__pICMSInter") = vNFeImgItem(ic).ICMSUFDest__pICMSInter
                t_NFe_IMAGEM_ITEM("ICMSUFDest__pICMSInterPart") = vNFeImgItem(ic).ICMSUFDest__pICMSInterPart
                t_NFe_IMAGEM_ITEM("ICMSUFDest__vFCPUFDest") = vNFeImgItem(ic).ICMSUFDest__vFCPUFDest
                t_NFe_IMAGEM_ITEM("ICMSUFDest__vICMSUFDest") = vNFeImgItem(ic).ICMSUFDest__vICMSUFDest
                t_NFe_IMAGEM_ITEM("ICMSUFDest__vICMSUFRemet") = vNFeImgItem(ic).ICMSUFDest__vICMSUFRemet
                End If
            t_NFe_IMAGEM_ITEM("det__infAdProd") = vNFeImgItem(ic).det__infAdProd
            t_NFe_IMAGEM_ITEM("det__vOutro") = vNFeImgItem(ic).det__vOutro
            t_NFe_IMAGEM_ITEM("det__indTot") = vNFeImgItem(ic).det__indTot
            t_NFe_IMAGEM_ITEM("det__xPed") = vNFeImgItem(ic).det__xPed
            t_NFe_IMAGEM_ITEM("det__nItemPed") = vNFeImgItem(ic).det__nItemPed
            t_NFe_IMAGEM_ITEM("det__vTotTrib") = vNFeImgItem(ic).det__vTotTrib
            t_NFe_IMAGEM_ITEM("ICMS__vBCSTRet") = vNFeImgItem(ic).ICMS__vBCSTRet
            t_NFe_IMAGEM_ITEM("ICMS__vICMSSTRet") = vNFeImgItem(ic).ICMS__vICMSSTRet
            t_NFe_IMAGEM_ITEM.Update
            End If
        Next

'   GRAVA OS DADOS INFORMATIVOS SOBRE AS PARCELAS DE PAGAMENTO
    intOrdem = 0
    For ic = LBound(vNFeImgTagDup) To UBound(vNFeImgTagDup)
        If Trim$(vNFeImgTagDup(ic).nDup) <> "" Then
            intOrdem = intOrdem + 1
            
            If Not geraNsu(NSU_T_NFe_IMAGEM_TAG_DUP, lngNsuNFeImagemTagDup, strMsgErro) Then
                If strMsgErro <> "" Then strMsgErro = Chr(13) & Chr(13) & strMsgErro
                strMsgErro = "Falha ao gravar os dados da NFe em " & NSU_T_NFe_IMAGEM_TAG_DUP & "!!" & strMsgErro
            '   ~~~~~~~~~~~~~~~~~
                dbc.RollbackTrans
            '   ~~~~~~~~~~~~~~~~~
                GoSub GRAVA_NFE_IMAGEM_FECHA_TABELAS
                aviso_erro strMsgErro
                Exit Function
                End If
            
            s = "SELECT " & _
                    "*" & _
                " FROM t_NFe_IMAGEM_TAG_DUP" & _
                " WHERE" & _
                    " (id = -1)"
            If t_NFe_IMAGEM_TAG_DUP.State <> adStateClosed Then t_NFe_IMAGEM_TAG_DUP.Close
            t_NFe_IMAGEM_TAG_DUP.Open s, dbc, , , adCmdText
            t_NFe_IMAGEM_TAG_DUP.AddNew
            t_NFe_IMAGEM_TAG_DUP("id") = lngNsuNFeImagemTagDup
            t_NFe_IMAGEM_TAG_DUP("id_nfe_imagem") = lngNsuNFeImagem
            t_NFe_IMAGEM_TAG_DUP("ordem") = intOrdem
            t_NFe_IMAGEM_TAG_DUP("nDup") = vNFeImgTagDup(ic).nDup
            t_NFe_IMAGEM_TAG_DUP("dVenc") = vNFeImgTagDup(ic).dVenc
            t_NFe_IMAGEM_TAG_DUP("vDup") = vNFeImgTagDup(ic).vDup
            t_NFe_IMAGEM_TAG_DUP.Update
            End If
        Next

'   GRAVA OS DADOS DE NF REFERENCIADA
    intOrdem = 0
    For ic = LBound(vNFeImgNFeRef) To UBound(vNFeImgNFeRef)
        If Trim$(vNFeImgNFeRef(ic).refNFe) <> "" Then
            intOrdem = intOrdem + 1
            
            If Not geraNsu(NSU_T_NFe_IMAGEM_NFe_REFERENCIADA, lngNsuNFeImagemNFeReferenciada, strMsgErro) Then
                If strMsgErro <> "" Then strMsgErro = Chr(13) & Chr(13) & strMsgErro
                strMsgErro = "Falha ao gravar os dados da NFe em " & NSU_T_NFe_IMAGEM_NFe_REFERENCIADA & "!!" & strMsgErro
            '   ~~~~~~~~~~~~~~~~~
                dbc.RollbackTrans
            '   ~~~~~~~~~~~~~~~~~
                GoSub GRAVA_NFE_IMAGEM_FECHA_TABELAS
                aviso_erro strMsgErro
                Exit Function
                End If
            
            s = "SELECT " & _
                    "*" & _
                " FROM t_NFe_IMAGEM_NFe_REFERENCIADA" & _
                " WHERE" & _
                    " (id = -1)"
            If t_NFe_IMAGEM_NFe_REFERENCIADA.State <> adStateClosed Then t_NFe_IMAGEM_NFe_REFERENCIADA.Close
            t_NFe_IMAGEM_NFe_REFERENCIADA.Open s, dbc, , , adCmdText
            t_NFe_IMAGEM_NFe_REFERENCIADA.AddNew
            t_NFe_IMAGEM_NFe_REFERENCIADA("id") = lngNsuNFeImagemNFeReferenciada
            t_NFe_IMAGEM_NFe_REFERENCIADA("id_nfe_imagem") = lngNsuNFeImagem
            t_NFe_IMAGEM_NFe_REFERENCIADA("ordem") = intOrdem
            t_NFe_IMAGEM_NFe_REFERENCIADA("refNFe") = vNFeImgNFeRef(ic).refNFe
            t_NFe_IMAGEM_NFe_REFERENCIADA("NFe_serie_NF_referenciada") = vNFeImgNFeRef(ic).NFe_serie_NF_referenciada
            t_NFe_IMAGEM_NFe_REFERENCIADA("NFe_numero_NF_referenciada") = vNFeImgNFeRef(ic).NFe_numero_NF_referenciada
            t_NFe_IMAGEM_NFe_REFERENCIADA.Update
            End If
        Next





'   GRAVA OS DADOS DE PAGAMENTO DA NF
    intOrdem = 0
    For ic = LBound(vNFeImgPag) To UBound(vNFeImgPag)
        If Trim$(vNFeImgPag(ic).pag__indPag) <> "" Then
            intOrdem = intOrdem + 1
            
            If Not geraNsu(NSU_T_NFe_IMAGEM_PAG, lngNsuNFeImagemPag, strMsgErro) Then
                If strMsgErro <> "" Then strMsgErro = Chr(13) & Chr(13) & strMsgErro
                strMsgErro = "Falha ao gravar os dados da NFe em " & NSU_T_NFe_IMAGEM_PAG & "!!" & strMsgErro
            '   ~~~~~~~~~~~~~~~~~
                dbc.RollbackTrans
            '   ~~~~~~~~~~~~~~~~~
                GoSub GRAVA_NFE_IMAGEM_FECHA_TABELAS
                aviso_erro strMsgErro
                Exit Function
                End If
            
            s = "SELECT " & _
                    "*" & _
                " FROM t_NFe_IMAGEM_PAG" & _
                " WHERE" & _
                    " (id = -1)"
            If t_NFe_IMAGEM_PAG.State <> adStateClosed Then t_NFe_IMAGEM_PAG.Close
            t_NFe_IMAGEM_PAG.Open s, dbc, , , adCmdText
            t_NFe_IMAGEM_PAG.AddNew
            t_NFe_IMAGEM_PAG("id") = lngNsuNFeImagemPag
            t_NFe_IMAGEM_PAG("id_nfe_imagem") = lngNsuNFeImagem
            t_NFe_IMAGEM_PAG("ordem") = intOrdem
            t_NFe_IMAGEM_PAG("pag__indPag") = vNFeImgPag(ic).pag__indPag
            t_NFe_IMAGEM_PAG("pag__tPag") = vNFeImgPag(ic).pag__tPag
            t_NFe_IMAGEM_PAG("pag__vPag") = vNFeImgPag(ic).pag__vPag
            t_NFe_IMAGEM_PAG.Update
            End If
        Next





'   ~~~~~~~~~~~~~~~
    dbc.CommitTrans
'   ~~~~~~~~~~~~~~~
    On Error GoTo GRAVA_NFE_IMAGEM_TRATA_ERRO

    GoSub GRAVA_NFE_IMAGEM_FECHA_TABELAS
    
    grava_NFe_imagem = True
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GRAVA_NFE_IMAGEM_TRATA_ERRO_TRANSACAO:
'=====================================
    strMsgErro = CStr(Err) & ": " & Error$(Err)
    On Error Resume Next
    dbc.RollbackTrans
    GoSub GRAVA_NFE_IMAGEM_FECHA_TABELAS
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GRAVA_NFE_IMAGEM_TRATA_ERRO:
'===========================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    strMsgErro = s
    GoSub GRAVA_NFE_IMAGEM_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GRAVA_NFE_IMAGEM_FECHA_TABELAS:
'==============================
  ' RECORDSETS
    bd_desaloca_recordset t_NFe_IMAGEM, True
    bd_desaloca_recordset t_NFe_IMAGEM_ITEM, True
    bd_desaloca_recordset t_NFe_IMAGEM_TAG_DUP, True
    bd_desaloca_recordset t_NFe_IMAGEM_NFe_REFERENCIADA, True
    bd_desaloca_recordset t_NFe_IMAGEM_PAG, True
    Return

End Function

Function le_NFe_imagem(ByVal id_nfe_imagem As Long, _
                       ByRef rNFeImg As TIPO_NFe_IMG, _
                       ByRef vNFeImgItem() As TIPO_NFe_IMG_ITEM, _
                       ByRef vNFeImgTagDup() As TIPO_NFe_IMG_TAG_DUP, _
                       ByRef vNFeImgNFeRef() As TIPO_NFe_IMG_NFe_REFERENCIADA, _
                       ByRef vNFeImgPag() As TIPO_NFe_IMG_PAG, _
                       ByRef msg_erro As String) As Boolean
' DECLARAÇÕES
Const NomeDestaRotina = "le_NFe_imagem()"
Dim s As String
Dim t_NFe_IMAGEM As ADODB.Recordset
Dim t_NFe_IMAGEM_ITEM As ADODB.Recordset
Dim t_NFe_IMAGEM_TAG_DUP As ADODB.Recordset
Dim t_NFe_IMAGEM_NFe_REFERENCIADA As ADODB.Recordset
Dim t_NFe_IMAGEM_PAG As ADODB.Recordset

    On Error GoTo LE_NFE_IMAGEM_TRATA_ERRO
    
    le_NFe_imagem = False
    msg_erro = ""
    
    limpa_TIPO_NFe_IMG rNFeImg
    
    ReDim vNFeImgItem(0)
    limpa_TIPO_NFe_IMG_ITEM vNFeImgItem()
    
    ReDim vNFeImgTagDup(0)
    limpa_TIPO_NFe_IMG_TAG_DUP vNFeImgTagDup()
    
    ReDim vNFeImgNFeRef(0)
    limpa_TIPO_NFe_IMG_NFe_REFERENCIADA vNFeImgNFeRef()

    ReDim vNFeImgPag(0)
    limpa_TIPO_NFe_IMG_PAG vNFeImgPag()

'   T_NFE_IMAGEM
    Set t_NFe_IMAGEM = New ADODB.Recordset
    With t_NFe_IMAGEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   T_NFE_IMAGEM_ITEM
    Set t_NFe_IMAGEM_ITEM = New ADODB.Recordset
    With t_NFe_IMAGEM_ITEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   T_NFE_IMAGEM_TAG_DUP
    Set t_NFe_IMAGEM_TAG_DUP = New ADODB.Recordset
    With t_NFe_IMAGEM_TAG_DUP
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   T_NFE_IMAGEM_NFE_REFERENCIADA
    Set t_NFe_IMAGEM_NFe_REFERENCIADA = New ADODB.Recordset
    With t_NFe_IMAGEM_NFe_REFERENCIADA
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

'   T_NFE_IMAGEM_PAG
    Set t_NFe_IMAGEM_PAG = New ADODB.Recordset
    With t_NFe_IMAGEM_PAG
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

    s = "SELECT " & _
            "*" & _
        " FROM t_NFe_IMAGEM" & _
        " WHERE" & _
            " (id = " & CStr(id_nfe_imagem) & ")"
    t_NFe_IMAGEM.Open s, dbc, , , adCmdText
    If t_NFe_IMAGEM.EOF Then
        msg_erro = "Não foi encontrado o registro com os dados da NFe referenciada (t_NFe_IMAGEM.id=" & CStr(id_nfe_imagem) & ")!!"
        GoSub LE_NFE_IMAGEM_FECHA_TABELAS
        Exit Function
        End If
    
    With rNFeImg
        .id = t_NFe_IMAGEM("id")
        .id_nfe_emitente = t_NFe_IMAGEM("id_nfe_emitente")
        .NFe_serie_NF = t_NFe_IMAGEM("NFe_serie_NF")
        .NFe_numero_NF = t_NFe_IMAGEM("NFe_numero_NF")
        .versao_layout_NFe = Trim$("" & t_NFe_IMAGEM("versao_layout_NFe"))
        .pedido = t_NFe_IMAGEM("pedido")
        .operacional__email = Trim$("" & t_NFe_IMAGEM("operacional__email"))
        .ide__natOp = Trim$("" & t_NFe_IMAGEM("ide__natOp"))
        .ide__indPag = Trim$("" & t_NFe_IMAGEM("ide__indPag"))
        .ide__serie = Trim$("" & t_NFe_IMAGEM("ide__serie"))
        .ide__nNF = Trim$("" & t_NFe_IMAGEM("ide__nNF"))
        .ide__dEmi = Trim$("" & t_NFe_IMAGEM("ide__dEmi"))
        .ide__dEmiUTC = Trim$("" & t_NFe_IMAGEM("ide__dEmiUTC"))
        .ide__dSaiEnt = Trim$("" & t_NFe_IMAGEM("ide__dSaiEnt"))
        .ide__tpNF = Trim$("" & t_NFe_IMAGEM("ide__tpNF"))
        .ide__idDest = Trim$("" & t_NFe_IMAGEM("ide__idDest"))
        .ide__cMunFG = Trim$("" & t_NFe_IMAGEM("ide__cMunFG"))
        .ide__tpAmb = Trim$("" & t_NFe_IMAGEM("ide__tpAmb"))
        .ide__finNFe = Trim$("" & t_NFe_IMAGEM("ide__finNFe"))
        .ide__indFinal = Trim$("" & t_NFe_IMAGEM("ide__indFinal"))
        .ide__indPres = Trim$("" & t_NFe_IMAGEM("ide__indPres"))
        .ide__IEST = Trim$("" & t_NFe_IMAGEM("ide__IEST"))
        .dest__CNPJ = Trim$("" & t_NFe_IMAGEM("dest__CNPJ"))
        .dest__CPF = Trim$("" & t_NFe_IMAGEM("dest__CPF"))
        .dest__xNome = Trim$("" & t_NFe_IMAGEM("dest__xNome"))
        .dest__xLgr = Trim$("" & t_NFe_IMAGEM("dest__xLgr"))
        .dest__nro = Trim$("" & t_NFe_IMAGEM("dest__nro"))
        .dest__xCpl = Trim$("" & t_NFe_IMAGEM("dest__xCpl"))
        .dest__xBairro = Trim$("" & t_NFe_IMAGEM("dest__xBairro"))
        .dest__cMun = Trim$("" & t_NFe_IMAGEM("dest__cMun"))
        .dest__xMun = Trim$("" & t_NFe_IMAGEM("dest__xMun"))
        .dest__UF = Trim$("" & t_NFe_IMAGEM("dest__UF"))
        .dest__CEP = Trim$("" & t_NFe_IMAGEM("dest__CEP"))
        .dest__cPais = Trim$("" & t_NFe_IMAGEM("dest__cPais"))
        .dest__xPais = Trim$("" & t_NFe_IMAGEM("dest__xPais"))
        .dest__fone = Trim$("" & t_NFe_IMAGEM("dest__fone"))
        .dest__IE = Trim$("" & t_NFe_IMAGEM("dest__IE"))
        .dest__ISUF = Trim$("" & t_NFe_IMAGEM("dest__ISUF"))
        .dest__idEstrangeiro = Trim$("" & t_NFe_IMAGEM("dest__idEstrangeiro"))
        .dest__indIEDest = Trim$("" & t_NFe_IMAGEM("dest__indIEDest"))
        .dest__email = Trim$("" & t_NFe_IMAGEM("dest__email"))
        .entrega__CNPJ = Trim$("" & t_NFe_IMAGEM("entrega__CNPJ"))
        .entrega__CPF = Trim$("" & t_NFe_IMAGEM("entrega__CPF"))
        .entrega__xLgr = Trim$("" & t_NFe_IMAGEM("entrega__xLgr"))
        .entrega__nro = Trim$("" & t_NFe_IMAGEM("entrega__nro"))
        .entrega__xCpl = Trim$("" & t_NFe_IMAGEM("entrega__xCpl"))
        .entrega__xBairro = Trim$("" & t_NFe_IMAGEM("entrega__xBairro"))
        .entrega__cMun = Trim$("" & t_NFe_IMAGEM("entrega__cMun"))
        .entrega__xMun = Trim$("" & t_NFe_IMAGEM("entrega__xMun"))
        .entrega__UF = Trim$("" & t_NFe_IMAGEM("entrega__UF"))
        .total__vBC = Trim$("" & t_NFe_IMAGEM("total__vBC"))
        .total__vICMS = Trim$("" & t_NFe_IMAGEM("total__vICMS"))
        .total__vICMSDeson = Trim$("" & t_NFe_IMAGEM("total__vICMSDeson"))
        .total__vBCST = Trim$("" & t_NFe_IMAGEM("total__vBCST"))
        .total__vST = Trim$("" & t_NFe_IMAGEM("total__vST"))
        .total__vProd = Trim$("" & t_NFe_IMAGEM("total__vProd"))
        .total__vFrete = Trim$("" & t_NFe_IMAGEM("total__vFrete"))
        .total__vSeg = Trim$("" & t_NFe_IMAGEM("total__vSeg"))
        .total__vDesc = Trim$("" & t_NFe_IMAGEM("total__vDesc"))
        .total__vII = Trim$("" & t_NFe_IMAGEM("total__vII"))
        .total__vIPI = Trim$("" & t_NFe_IMAGEM("total__vIPI"))
        .total__vPIS = Trim$("" & t_NFe_IMAGEM("total__vPIS"))
        .total__vCOFINS = Trim$("" & t_NFe_IMAGEM("total__vCOFINS"))
        .total__vOutro = Trim$("" & t_NFe_IMAGEM("total__vOutro"))
        .total__vNF = Trim$("" & t_NFe_IMAGEM("total__vNF"))
        .total__vTotTrib = Trim$("" & t_NFe_IMAGEM("total__vTotTrib"))
        .transp__modFrete = Trim$("" & t_NFe_IMAGEM("transp__modFrete"))
        .transporta__CNPJ = Trim$("" & t_NFe_IMAGEM("transporta__CNPJ"))
        .transporta__CPF = Trim$("" & t_NFe_IMAGEM("transporta__CPF"))
        .transporta__xNome = Trim$("" & t_NFe_IMAGEM("transporta__xNome"))
        .transporta__IE = Trim$("" & t_NFe_IMAGEM("transporta__IE"))
        .transporta__xEnder = Trim$("" & t_NFe_IMAGEM("transporta__xEnder"))
        .transporta__xMun = Trim$("" & t_NFe_IMAGEM("transporta__xMun"))
        .transporta__UF = Trim$("" & t_NFe_IMAGEM("transporta__UF"))
        .vol__qVol = Trim$("" & t_NFe_IMAGEM("vol__qVol"))
        .vol__esp = Trim$("" & t_NFe_IMAGEM("vol__esp"))
        .vol__marca = Trim$("" & t_NFe_IMAGEM("vol__marca"))
        .vol__nVol = Trim$("" & t_NFe_IMAGEM("vol__nVol"))
        .vol__pesoL = Trim$("" & t_NFe_IMAGEM("vol__pesoL"))
        .vol__pesoB = Trim$("" & t_NFe_IMAGEM("vol__pesoB"))
        .vol_nLacre = Trim$("" & t_NFe_IMAGEM("vol_nLacre"))
        .infAdic__infAdFisco = Trim$("" & t_NFe_IMAGEM("infAdic__infAdFisco"))
        .infAdic__infCpl = Trim$("" & t_NFe_IMAGEM("infAdic__infCpl"))
        .codigo_retorno_NFe_T1 = Trim$("" & t_NFe_IMAGEM("codigo_retorno_NFe_T1"))
        .msg_retorno_NFe_T1 = Trim$("" & t_NFe_IMAGEM("msg_retorno_NFe_T1"))
        End With

    s = "SELECT " & _
            "*" & _
        " FROM t_NFe_IMAGEM_ITEM" & _
        " WHERE" & _
            " (id_nfe_imagem = " & CStr(id_nfe_imagem) & ")" & _
        " ORDER BY" & _
            " ordem"
    t_NFe_IMAGEM_ITEM.Open s, dbc, , , adCmdText
    If t_NFe_IMAGEM_ITEM.EOF Then
        msg_erro = "Não foi encontrado o registro com os dados da NFe referenciada (t_NFe_IMAGEM_ITEM.id=" & CStr(id_nfe_imagem) & ")!!"
        GoSub LE_NFE_IMAGEM_FECHA_TABELAS
        Exit Function
        End If
            
    Do While Not t_NFe_IMAGEM_ITEM.EOF
        If (vNFeImgItem(UBound(vNFeImgItem)).id > 0) Or _
           (Trim$(vNFeImgItem(UBound(vNFeImgItem)).produto) <> "") Then
            ReDim Preserve vNFeImgItem(UBound(vNFeImgItem) + 1)
            End If
        
        With vNFeImgItem(UBound(vNFeImgItem))
            .id = t_NFe_IMAGEM_ITEM("id")
            .id_nfe_imagem = t_NFe_IMAGEM_ITEM("id_nfe_imagem")
            .fabricante = Trim$("" & t_NFe_IMAGEM_ITEM("fabricante"))
            .produto = Trim$("" & t_NFe_IMAGEM_ITEM("produto"))
            .det__nItem = Trim$("" & t_NFe_IMAGEM_ITEM("det__nItem"))
            .det__cProd = Trim$("" & t_NFe_IMAGEM_ITEM("det__cProd"))
            .det__cEAN = Trim$("" & t_NFe_IMAGEM_ITEM("det__cEAN"))
            .det__xProd = Trim$("" & t_NFe_IMAGEM_ITEM("det__xProd"))
            .det__NCM = Trim$("" & t_NFe_IMAGEM_ITEM("det__NCM"))
            .det__CEST = Trim$("" & t_NFe_IMAGEM_ITEM("det__CEST"))
            .det__indEscala = Trim$("" & t_NFe_IMAGEM_ITEM("det__indEscala"))
            .det__EXTIPI = Trim$("" & t_NFe_IMAGEM_ITEM("det__EXTIPI"))
            .det__genero = Trim$("" & t_NFe_IMAGEM_ITEM("det__genero"))
            .det__CFOP = Trim$("" & t_NFe_IMAGEM_ITEM("det__CFOP"))
            .det__uCom = Trim$("" & t_NFe_IMAGEM_ITEM("det__uCom"))
            .det__qCom = Trim$("" & t_NFe_IMAGEM_ITEM("det__qCom"))
            .det__vUnCom = Trim$("" & t_NFe_IMAGEM_ITEM("det__vUnCom"))
            .det__vProd = Trim$("" & t_NFe_IMAGEM_ITEM("det__vProd"))
            .det__cEANTrib = Trim$("" & t_NFe_IMAGEM_ITEM("det__cEANTrib"))
            .det__uTrib = Trim$("" & t_NFe_IMAGEM_ITEM("det__uTrib"))
            .det__qTrib = Trim$("" & t_NFe_IMAGEM_ITEM("det__qTrib"))
            .det__vUnTrib = Trim$("" & t_NFe_IMAGEM_ITEM("det__vUnTrib"))
            .det__vFrete = Trim$("" & t_NFe_IMAGEM_ITEM("det__vFrete"))
            .det__vSeg = Trim$("" & t_NFe_IMAGEM_ITEM("det__vSeg"))
            .det__vDesc = Trim$("" & t_NFe_IMAGEM_ITEM("det__vDesc"))
            .ICMS__orig = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__orig"))
            .ICMS__CST = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__CST"))
            .ICMS__modBC = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__modBC"))
            .ICMS__pRedBC = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__pRedBC"))
            .ICMS__vBC = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__vBC"))
            .ICMS__pICMS = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__pICMS"))
            .ICMS__vICMS = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__vICMS"))
            .ICMS__vICMSDeson = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__vICMSDeson"))
            .ICMS__modBCST = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__modBCST"))
            .ICMS__pMVAST = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__pMVAST"))
            .ICMS__pRedBCST = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__pRedBCST"))
            .ICMS__vBCST = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__vBCST"))
            .ICMS__pICMSST = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__pICMSST"))
            .ICMS__vICMSST = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__vICMSST"))
            .PIS__CST = Trim$("" & t_NFe_IMAGEM_ITEM("PIS__CST"))
            .PIS__vBC = Trim$("" & t_NFe_IMAGEM_ITEM("PIS__vBC"))
            .PIS__pPIS = Trim$("" & t_NFe_IMAGEM_ITEM("PIS__pPIS"))
            .PIS__vPIS = Trim$("" & t_NFe_IMAGEM_ITEM("PIS__vPIS"))
            .PIS__qBCProd = Trim$("" & t_NFe_IMAGEM_ITEM("PIS__qBCProd"))
            .PIS__vAliqProd = Trim$("" & t_NFe_IMAGEM_ITEM("PIS__vAliqProd"))
            .COFINS__CST = Trim$("" & t_NFe_IMAGEM_ITEM("COFINS__CST"))
            .COFINS__vBC = Trim$("" & t_NFe_IMAGEM_ITEM("COFINS__vBC"))
            .COFINS__pCOFINS = Trim$("" & t_NFe_IMAGEM_ITEM("COFINS__pCOFINS"))
            .COFINS__vCOFINS = Trim$("" & t_NFe_IMAGEM_ITEM("COFINS__vCOFINS"))
            .COFINS__qBCProd = Trim$("" & t_NFe_IMAGEM_ITEM("COFINS__qBCProd"))
            .COFINS__vAliqProd = Trim$("" & t_NFe_IMAGEM_ITEM("COFINS__vAliqProd"))
            .IPI__CST = Trim$("" & t_NFe_IMAGEM_ITEM("IPI__CST"))
            .IPI__clEnq = Trim$("" & t_NFe_IMAGEM_ITEM("IPI__clEnq"))
            .IPI__CNPJProd = Trim$("" & t_NFe_IMAGEM_ITEM("IPI__CNPJProd"))
            .IPI__cSelo = Trim$("" & t_NFe_IMAGEM_ITEM("IPI__cSelo"))
            .IPI__qSelo = Trim$("" & t_NFe_IMAGEM_ITEM("IPI__qSelo"))
            .IPI__cEnq = Trim$("" & t_NFe_IMAGEM_ITEM("IPI__cEnq"))
            .IPI__vBC = Trim$("" & t_NFe_IMAGEM_ITEM("IPI__vBC"))
            .IPI__qUnid = Trim$("" & t_NFe_IMAGEM_ITEM("IPI__qUnid"))
            .IPI__vUnid = Trim$("" & t_NFe_IMAGEM_ITEM("IPI__vUnid"))
            .IPI__pIPI = Trim$("" & t_NFe_IMAGEM_ITEM("IPI__pIPI"))
            .IPI__vIPI = Trim$("" & t_NFe_IMAGEM_ITEM("IPI__vIPI"))
            If PARTILHA_ICMS_ATIVA Then
                .ICMSUFDest__vBCUFDest = t_NFe_IMAGEM_ITEM("ICMSUFDest__vBCUFDest")
                .ICMSUFDest__pFCPUFDest = t_NFe_IMAGEM_ITEM("ICMSUFDest__pFCPUFDest")
                .ICMSUFDest__pICMSUFDest = t_NFe_IMAGEM_ITEM("ICMSUFDest__pICMSUFDest")
                .ICMSUFDest__pICMSInter = t_NFe_IMAGEM_ITEM("ICMSUFDest__pICMSInter")
                .ICMSUFDest__pICMSInterPart = t_NFe_IMAGEM_ITEM("ICMSUFDest__pICMSInterPart")
                .ICMSUFDest__vFCPUFDest = t_NFe_IMAGEM_ITEM("ICMSUFDest__vFCPUFDest")
                .ICMSUFDest__vICMSUFDest = t_NFe_IMAGEM_ITEM("ICMSUFDest__vICMSUFDest")
                .ICMSUFDest__vICMSUFRemet = t_NFe_IMAGEM_ITEM("ICMSUFDest__vICMSUFRemet")
                End If
            .det__infAdProd = Trim$("" & t_NFe_IMAGEM_ITEM("det__infAdProd"))
            .det__vOutro = Trim$("" & t_NFe_IMAGEM_ITEM("det__vOutro"))
            .det__indTot = Trim$("" & t_NFe_IMAGEM_ITEM("det__indTot"))
            .det__xPed = Trim$("" & t_NFe_IMAGEM_ITEM("det__xPed"))
            .det__nItemPed = Trim$("" & t_NFe_IMAGEM_ITEM("det__nItemPed"))
            .det__vTotTrib = Trim$("" & t_NFe_IMAGEM_ITEM("det__vTotTrib"))
            .ICMS__vBCSTRet = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__vBCSTRet"))
            .ICMS__vICMSSTRet = Trim$("" & t_NFe_IMAGEM_ITEM("ICMS__vICMSSTRet"))
            End With
            
        t_NFe_IMAGEM_ITEM.MoveNext
        Loop
        
    s = "SELECT " & _
            "*" & _
        " FROM t_NFe_IMAGEM_TAG_DUP" & _
        " WHERE" & _
            " (id_nfe_imagem = " & CStr(id_nfe_imagem) & ")" & _
        " ORDER BY" & _
            " ordem"
    t_NFe_IMAGEM_TAG_DUP.Open s, dbc, , , adCmdText
    Do While Not t_NFe_IMAGEM_TAG_DUP.EOF
        If (vNFeImgTagDup(UBound(vNFeImgTagDup)).id > 0) Or _
           (Trim$(vNFeImgTagDup(UBound(vNFeImgTagDup)).nDup) <> "") Then
            ReDim Preserve vNFeImgTagDup(UBound(vNFeImgTagDup) + 1)
            End If
            
        With vNFeImgTagDup(UBound(vNFeImgTagDup))
            .id = t_NFe_IMAGEM_TAG_DUP("id")
            .id_nfe_imagem = t_NFe_IMAGEM_TAG_DUP("id_nfe_imagem")
            .nDup = Trim$("" & t_NFe_IMAGEM_TAG_DUP("nDup"))
            .dVenc = Trim$("" & t_NFe_IMAGEM_TAG_DUP("dVenc"))
            .vDup = Trim$("" & t_NFe_IMAGEM_TAG_DUP("vDup"))
            End With
        
        t_NFe_IMAGEM_TAG_DUP.MoveNext
        Loop
        
    s = "SELECT " & _
            "*" & _
        " FROM t_NFe_IMAGEM_NFe_REFERENCIADA" & _
        " WHERE" & _
            " (id_nfe_imagem = " & CStr(id_nfe_imagem) & ")" & _
        " ORDER BY" & _
            " ordem"
    t_NFe_IMAGEM_NFe_REFERENCIADA.Open s, dbc, , , adCmdText
    Do While Not t_NFe_IMAGEM_NFe_REFERENCIADA.EOF
        If (vNFeImgNFeRef(UBound(vNFeImgNFeRef)).id > 0) Or _
           (Trim$(vNFeImgNFeRef(UBound(vNFeImgNFeRef)).refNFe) <> "") Then
            ReDim Preserve vNFeImgNFeRef(UBound(vNFeImgNFeRef) + 1)
            End If
        With vNFeImgNFeRef(UBound(vNFeImgNFeRef))
            .id = t_NFe_IMAGEM_NFe_REFERENCIADA("id")
            .id_nfe_imagem = t_NFe_IMAGEM_NFe_REFERENCIADA("id_nfe_imagem")
            .refNFe = Trim$("" & t_NFe_IMAGEM_NFe_REFERENCIADA("refNFe"))
            .NFe_serie_NF_referenciada = t_NFe_IMAGEM_NFe_REFERENCIADA("NFe_serie_NF_referenciada")
            .NFe_numero_NF_referenciada = t_NFe_IMAGEM_NFe_REFERENCIADA("NFe_numero_NF_referenciada")
            End With
        
        t_NFe_IMAGEM_NFe_REFERENCIADA.MoveNext
        Loop
    
    s = "SELECT " & _
            "*" & _
        " FROM t_NFe_IMAGEM_PAG" & _
        " WHERE" & _
            " (id_nfe_imagem = " & CStr(id_nfe_imagem) & ")" & _
        " ORDER BY" & _
            " ordem"
    t_NFe_IMAGEM_PAG.Open s, dbc, , , adCmdText
    Do While Not t_NFe_IMAGEM_PAG.EOF
        If (vNFeImgPag(UBound(vNFeImgPag)).id > 0) Or _
           (Trim$(vNFeImgPag(UBound(vNFeImgPag)).pag__indPag) <> "") Then
            ReDim Preserve vNFeImgPag(UBound(vNFeImgPag) + 1)
            End If
        With vNFeImgPag(UBound(vNFeImgPag))
            .id = t_NFe_IMAGEM_PAG("id")
            .id_nfe_imagem = t_NFe_IMAGEM_PAG("id_nfe_imagem")
            .pag__indPag = Trim$("" & t_NFe_IMAGEM_PAG("pag__indPag"))
            .pag__tPag = Trim$("" & t_NFe_IMAGEM_PAG("pag__tPag"))
            .pag__vPag = Trim$("" & t_NFe_IMAGEM_PAG("pag__vPag"))
            End With
        
        t_NFe_IMAGEM_PAG.MoveNext
        Loop
    
    le_NFe_imagem = True
    
    GoSub LE_NFE_IMAGEM_FECHA_TABELAS
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LE_NFE_IMAGEM_TRATA_ERRO:
'========================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    msg_erro = s
    GoSub LE_NFE_IMAGEM_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LE_NFE_IMAGEM_FECHA_TABELAS:
'===========================
  ' RECORDSETS
    bd_desaloca_recordset t_NFe_IMAGEM, True
    bd_desaloca_recordset t_NFe_IMAGEM_ITEM, True
    bd_desaloca_recordset t_NFe_IMAGEM_TAG_DUP, True
    bd_desaloca_recordset t_NFe_IMAGEM_NFe_REFERENCIADA, True
    bd_desaloca_recordset t_NFe_IMAGEM_PAG, True
    Return
    
End Function

Function NFeObtemProximoNumero(ByVal id_nfe_emitente As Integer, ByRef strNFeSerie As String, ByRef strNFeNumero As String, ByRef msg_erro As String) As Boolean

Const MAX_TENTATIVAS = 10
Dim intQtdeTentativas As Integer
Dim blnSucesso As Boolean
Dim s As String
Dim lngSerieNF As Long
Dim lngUltimoNumeroNF As Long
Dim lngProximoNumeroNF As Long
Dim lngRecordsAffected As Long
Dim t As ADODB.Recordset

    On Error GoTo NFE_OPN_TRATA_ERRO
    
    NFeObtemProximoNumero = False
    
  ' INICIALIZAÇÃO
    strNFeSerie = ""
    strNFeNumero = ""
    msg_erro = ""
    
    lngSerieNF = 0
    lngUltimoNumeroNF = 0
    lngProximoNumeroNF = 0
    
  ' RECORDSET
    Set t = New ADODB.Recordset
    t.CursorType = BD_CURSOR_SOMENTE_LEITURA
    t.LockType = BD_POLITICA_LOCKING
    
'   LAÇO DE TENTATIVAS PARA GERAR O NSU (DEVIDO A ACESSO CONCORRENTE)
    Do
        intQtdeTentativas = intQtdeTentativas + 1
    
        s = "SELECT" & _
                " NFe_serie_NF," & _
                " NFe_numero_NF" & _
            " FROM t_NFE_EMITENTE" & _
            " WHERE" & _
                " (id = " & CStr(id_nfe_emitente) & ")"
        If t.State <> adStateClosed Then t.Close
        t.Open s, dbc, , , adCmdText
        If t.EOF Then
            msg_erro = "Falha ao localizar o registro do emitente durante a geração do próximo número de NFe!!"
            GoSub NFE_OPN_FECHA_TABELAS
            Exit Function
            End If
        
        lngSerieNF = t("NFe_serie_NF")
        lngUltimoNumeroNF = t("NFe_numero_NF")
        lngProximoNumeroNF = lngUltimoNumeroNF + 1
        
        If lngSerieNF = 0 Then
            msg_erro = "O nº de série da NFe não foi definido!!"
            GoSub NFE_OPN_FECHA_TABELAS
            Exit Function
            End If
            
        s = "UPDATE t_NFE_EMITENTE SET" & _
                " NFe_numero_NF = " & CStr(lngProximoNumeroNF) & _
            " WHERE" & _
                " (id = " & CStr(id_nfe_emitente) & ")" & _
                " AND (NFe_serie_NF = " & CStr(lngSerieNF) & ")" & _
                " AND (NFe_numero_NF = " & CStr(lngUltimoNumeroNF) & ")"
        Call dbc.Execute(s, lngRecordsAffected)
        If lngRecordsAffected = 1 Then
            blnSucesso = True
            strNFeSerie = CStr(lngSerieNF)
            strNFeNumero = CStr(lngProximoNumeroNF)
        Else
            Sleep 100
            End If
    
        Loop While (Not blnSucesso) And (intQtdeTentativas < MAX_TENTATIVAS)
        
    
'   NÃO CONSEGUIU GERAR O NÚMERO?
    If Not blnSucesso Then
        msg_erro = "Falha ao tentar gerar o próximo número de NFe após " & CStr(MAX_TENTATIVAS) & " tentativas consecutivas!!"
        GoSub NFE_OPN_FECHA_TABELAS
        Exit Function
        End If
    
    GoSub NFE_OPN_FECHA_TABELAS
    
    NFeObtemProximoNumero = True
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_OPN_FECHA_TABELAS:
'=====================
    bd_desaloca_recordset t, True
    Return


NFE_OPN_TRATA_ERRO:
'==================
    s = CStr(Err) & ": " & Error$(Err)
    GoSub NFE_OPN_FECHA_TABELAS
    aviso s
    Exit Function

End Function

Function NFeObtemUltimoNumeroEmitido(ByVal id_nfe_emitente As Integer, ByRef lngNFeSerie As Long, ByRef lngNFeNumeroNF As Long, ByRef msg_erro As String) As Boolean

Dim s As String
Dim t As ADODB.Recordset

    On Error GoTo NFE_OUNE_TRATA_ERRO
    
    NFeObtemUltimoNumeroEmitido = False
    
  ' INICIALIZAÇÃO
    lngNFeSerie = 0
    lngNFeNumeroNF = 0
    msg_erro = ""
    
  ' RECORDSET
    Set t = New ADODB.Recordset
    t.CursorType = BD_CURSOR_SOMENTE_LEITURA
    t.LockType = BD_POLITICA_LOCKING
    
    s = "SELECT" & _
            " NFe_serie_NF," & _
            " NFe_numero_NF" & _
        " FROM t_NFE_EMITENTE" & _
        " WHERE" & _
            " (id = " & CStr(id_nfe_emitente) & ")"
    If t.State <> adStateClosed Then t.Close
    t.Open s, dbc, , , adCmdText
    If t.EOF Then
        msg_erro = "Falha ao localizar o registro do emitente durante a leitura do último número de NFe emitido!!"
        GoSub NFE_OUNE_FECHA_TABELAS
        Exit Function
        End If
    
    lngNFeSerie = t("NFe_serie_NF")
    lngNFeNumeroNF = t("NFe_numero_NF")
    
    GoSub NFE_OUNE_FECHA_TABELAS
    
    NFeObtemUltimoNumeroEmitido = True
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_OUNE_FECHA_TABELAS:
'======================
    bd_desaloca_recordset t, True
    Return


NFE_OUNE_TRATA_ERRO:
'===================
    s = CStr(Err) & ": " & Error$(Err)
    GoSub NFE_OUNE_FECHA_TABELAS
    aviso s
    Exit Function

End Function


Function obtem_emitentes_usuario(iduser As String, ByRef v() As TIPO_CINCO_COLUNAS, ByRef qtd As Integer) As Boolean

'o vetor v armazenará as seguintes informações do Emitente:
'- na coluna c1, a concatenação de UF / apelido
'- na coluna c2, o id
'- na coluna c3, o texto fixo relacionado ao Emitente, caso haja

Dim s As String
Dim s_uf As String
Dim t As ADODB.Recordset

    On Error GoTo NFE_EMIT_TRATA_ERRO
    
    obtem_emitentes_usuario = False
    
  ' INICIALIZAÇÃO
    ReDim v(0)
    qtd = 0
    
  ' RECORDSETS
    Set t = New ADODB.Recordset
    t.CursorType = BD_CURSOR_SOMENTE_LEITURA
    t.LockType = BD_POLITICA_LOCKING
    
    s = "SELECT" & _
            " n.uf, n.apelido, n.razao_social, n.id, n.texto_fixo_especifico" & _
        " FROM t_NFE_EMITENTE n" & _
            " INNER JOIN t_USUARIO_X_NFe_EMITENTE u ON n.id = u.id_nfe_emitente" & _
        " WHERE" & _
            " (n.st_ativo = 1)" & _
            " AND (u.excluido_status = 0)" & _
            " AND (u.usuario = '" & usuario.id & "')" & _
        " ORDER BY n.apelido, n.uf"
    
    If t.State <> adStateClosed Then t.Close
    t.Open s, dbc, , , adCmdText
    If Not t.EOF Then
        Do While Not t.EOF
            If v(UBound(v)).c1 <> "" Then ReDim Preserve v(UBound(v) + 1)
            s_uf = Trim("" & t("uf"))
            'v(UBound(v)).c1 = Trim("" & t("apelido")) & " - " & Trim("" & t("razao_social")) & " (" & s_uf & ")"
            v(UBound(v)).c1 = Trim("" & t("apelido")) & " (" & s_uf & ")"
            v(UBound(v)).c2 = Trim(CStr(t("id")))
            v(UBound(v)).c3 = Trim("" & t("texto_fixo_especifico"))
            qtd = qtd + 1
            t.MoveNext
            Loop
        End If
    
    GoSub NFE_EMIT_FECHA_TABELAS
    
    If qtd <= 0 Then Exit Function
    
    obtem_emitentes_usuario = True
    
Exit Function




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_EMIT_FECHA_TABELAS:
'======================
    bd_desaloca_recordset t, True
    Return


NFE_EMIT_TRATA_ERRO:
'===================
    s = CStr(Err) & ": " & Error$(Err)
    GoSub NFE_EMIT_FECHA_TABELAS
    aviso s
    Exit Function
    
End Function

Sub get_registro_t_parametro(ByVal id_registro As String, ByRef rx As TIPO_t_PARAMETRO)
Dim r As ADODB.Recordset
Dim s As String
    
    On Error GoTo GRTP_TRATA_ERRO
    
    rx.campo_inteiro = 0
    rx.campo_monetario = 0
    rx.campo_real = 0
    rx.campo_data = 0
    rx.campo_texto = ""
    rx.dt_hr_ult_atualizacao = 0
    rx.usuario_ult_atualizacao = ""
    
  ' RECORDSETS
    Set r = New ADODB.Recordset
    r.CursorType = BD_CURSOR_SOMENTE_LEITURA
    r.LockType = BD_POLITICA_LOCKING
    
    id_registro = Trim("" & id_registro)
    s = "SELECT " & _
            "*" & _
        " FROM t_PARAMETRO" & _
        " WHERE" & _
            " (id = '" & id_registro & "')"
    If r.State <> adStateClosed Then r.Close
    r.Open s, dbc, , , adCmdText
    If Not r.EOF Then
        rx.id = Trim("" & r("id"))
        rx.campo_inteiro = r("campo_inteiro")
        rx.campo_monetario = r("campo_monetario")
        rx.campo_real = r("campo_real")
        If Not IsNull(r("campo_data")) Then rx.campo_data = r("campo_data")
        rx.campo_texto = "" & r("campo_texto")
        End If
    
    GoSub GRTP_FECHA_TABELAS
    
    Exit Sub
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GRTP_FECHA_TABELAS:
'======================
    bd_desaloca_recordset r, True
    Return


GRTP_TRATA_ERRO:
'===================
    s = CStr(Err) & ": " & Error$(Err)
    GoSub GRTP_FECHA_TABELAS
    aviso s
    Exit Sub
    
End Sub


Sub obtem_parametros_t_versao(ByRef cor As String, ByRef identificador As String)

'   OBTÉM A COR DE FUNDO PADRÃO E O IDENTIFICADOR DO AMBIENTE,
'   EXISTENTES NA TABELA T_VERSAO
Dim s As String
Dim rs

    s = "SELECT cor_fundo_padrao, identificador_ambiente FROM t_VERSAO" & _
        " WHERE (modulo = '" & Trim$(App.Title) & "')"
    
    Set rs = dbc.Execute(s)
    If Not rs.EOF Then
        cor = Trim$("" & rs("cor_fundo_padrao"))
        identificador = Trim$("" & rs("identificador_ambiente"))
        End If
    
End Sub


Function isCampoExistenteEmTabela(ByVal campo As String, ByVal tabela As String) As Boolean

Dim s As String
Dim rs

    isCampoExistenteEmTabela = False
    
    s = "IF EXISTS(select 1 from syscolumns " & _
                        "where id = object_id('" & tabela & _
                        "') and name = '" & campo & "') " & _
            "SELECT 'S' AS resposta ELSE SELECT 'N' AS resposta"

    Set rs = dbc.Execute(s)
    If Not rs.EOF Then
        isCampoExistenteEmTabela = Trim$("" & rs("resposta")) = "S"
        End If
    
    
End Function

Function NFeExisteNotaTriangularEmEmissao(ByRef id As Long, usuario As String, ByRef msg_erro As String) As Boolean
'Deve-se bloquear a emissão de notas em paralelo quando
'   - existe o parâmetro definindo o tempo de espera na tabela t_PARAMETRO
'   - a operação triangular está em andamento (emissao_status = 1)
'   - a nota de venda ainda não foi emitida (Nfe_venda_emissao_status < 2)

        Dim s As String
        Dim t As ADODB.Recordset

        On Error GoTo NFE_ENFTEE_TRATA_ERRO

        NFeExisteNotaTriangularEmEmissao = False

        msg_erro = ""

        Set t = New ADODB.Recordset
        t.CursorType = BD_CURSOR_SOMENTE_LEITURA
        t.LockType = BD_POLITICA_LOCKING

        s = "SELECT" & _
                " *" & _
            " FROM t_NFE_TRIANGULAR" & _
            " WHERE" & _
                " (emissao_status = 1)" & _
            " AND" & _
                " (Nfe_venda_emissao_status < 2)" & _
            " AND EXISTS" & _
                " (select p.campo_inteiro from t_PARAMETRO p where p.id = 'NF_MaxSegundos_EsperaNotaTriangular' and (datediff(s, dt_hr_status, getdate()) <= p.campo_inteiro))"
        If t.State <> adStateClosed Then t.Close
        t.Open s, dbc, , , adCmdText
        If t.EOF Then
            GoSub NFE_ENFTEE_FECHA_TABELAS
            Exit Function
            End If

        id = t("id")
        usuario = t("usuario_emissao_status")

        GoSub NFE_ENFTEE_FECHA_TABELAS

        NFeExisteNotaTriangularEmEmissao = True

        Exit Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_ENFTEE_FECHA_TABELAS:
'======================
        bd_desaloca_recordset t, True
        Return


NFE_ENFTEE_TRATA_ERRO:
'===================
        s = CStr(Err) & ": " & Error$(Err)
        GoSub NFE_ENFTEE_FECHA_TABELAS
        aviso (s)
        Exit Function

    End Function


Function pedido_vinculado_a_nf_triangular(ByRef pedido As String, _
                                            ByRef iStatus As Integer, _
                                            ByRef lngId As Long, _
                                            ByRef lngSerieRemessa As Long, _
                                            ByRef lngNumVenda As Long, _
                                            ByRef lngNumRemessa As Long, _
                                            ByRef rec_cnpj_cpf As String, _
                                            ByRef rec_nome As String, _
                                            ByRef rec_rg As String, _
                                            ByRef rec_endereco As String, _
                                            ByRef rec_numero As String, _
                                            ByRef rec_complemento As String, _
                                            ByRef rec_bairro As String, _
                                            ByRef rec_cidade As String, _
                                            ByRef rec_uf As String, _
                                            ByRef rec_cep As String) As Boolean

        Dim s As String
        Dim t As ADODB.Recordset

        On Error GoTo PVANFT_TRATA_ERRO

        pedido_vinculado_a_nf_triangular = False

        Set t = New ADODB.Recordset
        t.CursorType = BD_CURSOR_SOMENTE_LEITURA
        t.LockType = BD_POLITICA_LOCKING

        iStatus = 0
        lngId = 0
        lngSerieRemessa = 0
        lngNumVenda = 0
        lngNumRemessa = 0
        rec_cnpj_cpf = ""
        rec_nome = ""
        rec_rg = ""
        rec_endereco = ""
        rec_numero = ""
        rec_complemento = ""
        rec_bairro = ""
        rec_cidade = ""
        rec_uf = ""
        rec_cep = ""
        s = "SELECT" & _
                " *" & _
            " FROM t_NFE_TRIANGULAR" & _
            " WHERE" & _
                " (pedido = '" & pedido & "')" & _
            " AND " & _
                " (Nfe_venda_emissao_status = " & CStr(ST_NFT_EMITIDA) & ")" & _
            " AND " & _
                " (emissao_status NOT IN (" & CStr(ST_NFT_CANCELADA_USUARIO) & "," & CStr(ST_NFT_CANCELADA_TIMEOUT) & "," & CStr(ST_NFT_CANCELADA_SISTEMA) & "))" & _
            " AND" & _
                " (id_nfe_emitente = " & usuario.emit_id & ")"
        If t.State <> adStateClosed Then t.Close
        t.Open s, dbc, , , adCmdText
        If t.EOF Then
            GoSub PVANFT_FECHA_TABELAS
            Exit Function
            End If
        
        iStatus = t("emissao_status")
        lngId = t("id")
        lngSerieRemessa = t("Nfe_serie_remessa")
        lngNumVenda = t("Nfe_numero_venda")
        lngNumRemessa = t("Nfe_numero_remessa")
        rec_cnpj_cpf = Trim("" & t("recebedor_cnpj_cpf"))
        rec_nome = Trim("" & t("recebedor_nome"))
        rec_rg = Trim("" & t("recebedor_rg"))
        rec_endereco = Trim("" & t("recebedor_endereco"))
        rec_numero = Trim("" & t("recebedor_numero"))
        rec_complemento = Trim("" & t("recebedor_complemento"))
        rec_bairro = Trim("" & t("recebedor_bairro"))
        rec_cidade = Trim("" & t("recebedor_cidade"))
        rec_uf = Trim("" & t("recebedor_uf"))
        rec_cep = Trim("" & t("recebedor_cep"))

        GoSub PVANFT_FECHA_TABELAS

        pedido_vinculado_a_nf_triangular = True

        Exit Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PVANFT_FECHA_TABELAS:
'======================
        bd_desaloca_recordset t, True
        Return


PVANFT_TRATA_ERRO:
'===================
        s = CStr(Err) & ": " & Error$(Err)
        GoSub PVANFT_FECHA_TABELAS
        aviso (s)
        Exit Function

    End Function


Function RetornaOperacoesTriangularesPendentes() As String

Dim s As String
Dim sRetorno As String
Dim t As ADODB.Recordset

        On Error GoTo NFE_ROTP_TRATA_ERRO

        sRetorno = ""

        Set t = New ADODB.Recordset
        t.CursorType = BD_CURSOR_SOMENTE_LEITURA
        t.LockType = BD_POLITICA_LOCKING

        s = "SELECT" & _
                " *" & _
            " FROM t_NFE_TRIANGULAR" & _
            " WHERE" & _
                " (emissao_status = 1)" & _
            " AND" & _
                " (id_nfe_emitente = " & usuario.emit_id & ")"
        If t.State <> adStateClosed Then t.Close
        t.Open s, dbc, , , adCmdText
        If Not t.EOF Then
            sRetorno = "OPERAÇÕES TRIANGULARES PENDENTES" & vbCrLf & _
                       "================================" & vbCrLf & vbCrLf
            Do While Not t.EOF
                sRetorno = sRetorno & "Pedido: " & t("Pedido") & vbCrLf
                sRetorno = sRetorno & "NF Venda: " & NFeFormataNumeroNF(t("Nfe_numero_venda")) & _
                            IIf(t("Nfe_venda_emissao_status") = 2, " (Emitida)", " (Pendente)") & vbCrLf
                sRetorno = sRetorno & "NF Remessa: " & NFeFormataNumeroNF(t("Nfe_numero_remessa")) & _
                            IIf(t("Nfe_remessa_emissao_status") = 2, " (Emitida)", " (Pendente)") & vbCrLf & vbCrLf
                t.MoveNext
                Loop
            End If

        GoSub NFE_ROTP_FECHA_TABELAS

        RetornaOperacoesTriangularesPendentes = sRetorno

        Exit Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_ROTP_FECHA_TABELAS:
'======================
        bd_desaloca_recordset t, True
        Return


NFE_ROTP_TRATA_ERRO:
'===================
        s = CStr(Err) & ": " & Error$(Err)
        GoSub NFE_ROTP_FECHA_TABELAS
        aviso (s)
        Exit Function

End Function

Function RetornaNumeracaoRemessaPendente() As String

Dim s As String
Dim sRetorno As String
Dim t As ADODB.Recordset
Dim r As ADODB.Recordset

        On Error GoTo NFE_RNRP_TRATA_ERRO

        sRetorno = ""

        Set t = New ADODB.Recordset
        t.CursorType = BD_CURSOR_SOMENTE_LEITURA
        t.LockType = BD_POLITICA_LOCKING

        Set r = New ADODB.Recordset
        r.CursorType = BD_CURSOR_SOMENTE_LEITURA
        r.LockType = BD_POLITICA_LOCKING

        s = "SELECT" & _
                " *" & _
            " FROM t_NFE_TRIANGULAR" & _
            " WHERE" & _
                " (emissao_status = 3)" & _
            " AND" & _
                " (Nfe_remessa_emissao_status in (0, 1))" & _
            " AND" & _
                " (id_nfe_emitente = " & usuario.emit_id & ")"
        If t.State <> adStateClosed Then t.Close
        t.Open s, dbc, , , adCmdText
        Do While Not t.EOF

            s = "SELECT" & _
                    " NFe_serie_nf, NFe_numero_nf" & _
                " FROM t_NFE_EMISSAO" & _
                " WHERE" & _
                    " (NFe_serie_nf = " & CStr(t("Nfe_serie_remessa")) & ")" & _
                " AND" & _
                    " (NFe_numero_nf = " & CStr(t("Nfe_numero_remessa")) & ")" & _
            " AND" & _
                " (id_nfe_emitente = " & usuario.emit_id & ")"
            If r.State <> adStateClosed Then r.Close
            r.Open s, dbc, , , adCmdText
            
            If r.EOF Then
                If sRetorno = "" Then
                    sRetorno = "OPERAÇÕES TRIANGULARES - NUMERAÇÃO PENDENTE" & vbCrLf & _
                               "===========================================" & vbCrLf & vbCrLf
                    End If
                sRetorno = sRetorno & NFeFormataNumeroNF(t("Nfe_numero_remessa")) & vbCrLf
                End If
                
            t.MoveNext
            Loop

        GoSub NFE_RNRP_FECHA_TABELAS

        RetornaNumeracaoRemessaPendente = sRetorno

        Exit Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NFE_RNRP_FECHA_TABELAS:
'======================
        bd_desaloca_recordset t, True
        bd_desaloca_recordset r, True
        Return


NFE_RNRP_TRATA_ERRO:
'===================
        s = CStr(Err) & ": " & Error$(Err)
        GoSub NFE_RNRP_FECHA_TABELAS
        aviso (s)
        Exit Function

End Function




