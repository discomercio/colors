Attribute VB_Name = "mod_BD"
Option Explicit

  ' CONEXÃO AO BANCO DE DADOS
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~
    Global dbc As ADODB.Connection

  ' CONSTANTES PARA USAR COM O BANCO DE DADOS
    Const BD_DATA_NULA = "DEC 30 1899"


    Type TIPO_LOG_VIA_VETOR
        nome As String
        valor As String
        End Type


Function remove_tabelas_temporarias(ByVal prefixoTabelaTemporaria As String, ByRef msg_erro As String) As Boolean
Dim s  As String
Dim rs As ADODB.Recordset
        
    On Error GoTo RTT_TRATA_ERRO
    
    remove_tabelas_temporarias = False
    
    msg_erro = ""
    
    If Trim$(prefixoTabelaTemporaria) = "" Then
        msg_erro = "Não foi informado o prefixo usado no nome das tabelas temporárias!!"
        Exit Function
        End If
    
    Set rs = New ADODB.Recordset
    With rs
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
    s = "SELECT" & _
            " name" & _
        " FROM sysobjects" & _
        " WHERE" & _
            " (type = 'U')" & _
            " AND (name LIKE '" & prefixoTabelaTemporaria & "%')" & _
        " ORDER BY" & _
            " name"
    If rs.State <> adStateClosed Then rs.Close
    rs.Open s, dbc, , , adCmdText
    Do While Not rs.EOF
        s = "DROP TABLE " & Trim$("" & rs("name"))
        dbc.Execute s
        rs.MoveNext
        Loop
    
    remove_tabelas_temporarias = True
    
    GoSub RTT_FECHA_TABELAS
        
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
RTT_TRATA_ERRO:
'==============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    GoSub RTT_FECHA_TABELAS
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
RTT_FECHA_TABELAS:
'=================
    bd_desaloca_recordset rs, True
    Return
    
End Function

Function obtem_site_sistema()

'   OBTÉM A INFORMAÇÃO DE QUAL SITE ESTÁ SENDO CONECTADO
Dim s As String
Dim rs

    obtem_site_sistema = ""
    
    s = "SELECT mensagem FROM t_VERSAO" & _
        " WHERE (modulo = '" & Trim$(App.Title) & "')"
    
    Set rs = dbc.Execute(s)
    If Not rs.EOF Then obtem_site_sistema = Trim$("" & rs("mensagem"))
    
End Function

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

Function grava_log(ByVal usuario As String, ByVal loja As String, ByVal pedido As String, ByVal id_cliente As String, ByVal operacao As String, ByVal complemento As String, ByRef msg_erro As String) As Boolean
' ___________________________________________
' GRAVA LOG
'
Dim s  As String
Dim rs As ADODB.Recordset
        
    On Error GoTo GL_TRATA_ERRO
    
    grava_log = False
    
    msg_erro = ""
    
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
    
    GoSub GL_FECHA_TABELAS
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GL_TRATA_ERRO:
'=============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    GoSub GL_FECHA_TABELAS
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GL_FECHA_TABELAS:
'================
    bd_desaloca_recordset rs, True
    Return
    
    
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
        If Not DESENVOLVIMENTO Then
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
        rx.id = r("id")
        rx.campo_inteiro = r("campo_inteiro")
        rx.campo_monetario = r("campo_monetario")
        rx.campo_real = r("campo_real")
        If Not IsNull(r("campo_data")) Then rx.campo_data = r("campo_data")
        rx.campo_texto = Trim("" & r("campo_texto"))
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


