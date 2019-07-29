VERSION 5.00
Begin VB.Form f_PEDIDO_ASSISTENCIA 
   Caption         =   "Consulta de Pedido da Assistência Técnica"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9765
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5355
      Picture         =   "f_PEDIDO_ASSISTENCIA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1425
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "Confirma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2760
      Picture         =   "f_PEDIDO_ASSISTENCIA.frx":019D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1425
   End
   Begin VB.Frame pnDados 
      Caption         =   "Dados do Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   3105
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   9075
      Begin VB.Label lblInfoAssist 
         Caption         =   "lblInfoAssist"
         Height          =   2295
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   8535
      End
   End
   Begin VB.Frame pnPedido 
      Caption         =   "Nº Pedido Assistência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      Begin VB.TextBox c_pedido_assist 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   360
         MaxLength       =   15
         TabIndex        =   1
         Top             =   345
         Width           =   1545
      End
   End
   Begin VB.Label lblInfoAguarde 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "consultando banco de dados..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "f_PEDIDO_ASSISTENCIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Function obtem_info_pedido_assistencia(ByVal pedido As String, ByRef strResposta As String, ByRef strEndEntregaFormatado As String, ByRef strEndEntregaUf As String, ByRef strEndClienteUf As String, ByRef strCNPJCPF, ByRef strMsgErro As String) As Boolean
' CONSTANTES
Const NomeDestaRotina = "obtem_info_pedido_assistencia()"
' STRINGS
Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim s_endereco As String
Dim s_bairro As String
Dim s_cep As String
Dim s_cidade As String
Dim s_uf As String
Dim s_nome As String
Dim s_cnpj_cpf As String
Dim s_ie_rg As String
Dim s_obs_1 As String
Dim s_info As String
Dim s_end_linha_1 As String
Dim s_end_linha_2 As String
Dim s_end_entrega As String
Dim pedido_a As String
Dim s_id_cliente As String

' BANCO DE DADOS
Dim t_PEDIDO As ADODB.Recordset
Dim t_PEDIDO_ITEM As ADODB.Recordset
Dim t_DESTINATARIO As ADODB.Recordset

    On Error GoTo OIPA_TRATA_ERRO
    
    cmdConfirma.Enabled = False
    
    obtem_info_pedido_assistencia = False
    strMsgErro = ""
    strResposta = ""
    strEndEntregaFormatado = ""
    strEndEntregaUf = ""
    strEndClienteUf = ""
    
    pedido = Trim$("" & pedido)
    pedido = normaliza_num_pedido(pedido)
    
    If pedido = "" Then
        strMsgErro = "Não foi informado o número do pedido!"
        cmdConfirma.Enabled = True
        Exit Function
        End If
        
    lblInfoAguarde.Visible = True
    Screen.MousePointer = vbHourglass
    Screen.ActiveForm.MousePointer = vbHourglass
    Me.Refresh
    
  ' T_PEDIDO
    Set t_PEDIDO = New ADODB.Recordset
    With t_PEDIDO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_PEDIDO_ITEM
    Set t_PEDIDO_ITEM = New ADODB.Recordset
    With t_PEDIDO_ITEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  ' T_DESTINATARIO (PODE SER T_CLIENTE OU T_LOJA)
    Set t_DESTINATARIO = New ADODB.Recordset
    With t_DESTINATARIO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
    s_endereco = ""
    s_bairro = ""
    s_cep = ""
    s_cidade = ""
    s_uf = ""
    s_nome = ""
    s_cnpj_cpf = ""
    s_ie_rg = ""
    s_obs_1 = ""
    s_end_entrega = ""
        
'   VERIFICA O PEDIDO
    s_id_cliente = ""
    pedido_a = ""
    s_erro = ""
    s = "SELECT" & _
            " pedido, st_entrega, id_cliente, obs_1, st_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep" & _
        " FROM t_PEDIDO" & _
        " WHERE" & _
            " (pedido = '" & Trim$(pedido) & "')"
    If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
    t_PEDIDO.Open s, dbcAssist, , , adCmdText
    If t_PEDIDO.EOF Then
        If s_erro <> "" Then s_erro = s_erro & vbCrLf
        s_erro = s_erro & "Pedido " & Trim$(pedido) & " não está cadastrado !!"
    Else
    '   ENDEREÇO DE ENTREGA
        If (s_end_entrega = "") And (CLng(t_PEDIDO("st_end_entrega")) <> 0) Then
            s_end_entrega = formata_endereco(Trim("" & t_PEDIDO("EndEtg_endereco")), Trim("" & t_PEDIDO("EndEtg_endereco_numero")), Trim("" & t_PEDIDO("EndEtg_endereco_complemento")), Trim("" & t_PEDIDO("EndEtg_bairro")), Trim("" & t_PEDIDO("EndEtg_cidade")), Trim("" & t_PEDIDO("EndEtg_uf")), Trim("" & t_PEDIDO("EndEtg_cep")))
            s_end_entrega = UCase$(s_end_entrega)
            strEndEntregaFormatado = s_end_entrega
            strEndEntregaUf = UCase$(Trim("" & t_PEDIDO("EndEtg_uf")))
            If s_end_entrega <> "" Then s_end_entrega = vbCrLf & "ENTREGA: " & s_end_entrega
            End If
    
        If UCase$(Trim$("" & t_PEDIDO("st_entrega"))) = ST_ENTREGA_CANCELADO Then
            If s_erro <> "" Then s_erro = s_erro & vbCrLf
            s_erro = s_erro & "Pedido " & Trim$(pedido) & " está cancelado !!"
            End If
            
        s_id_cliente = Trim$("" & t_PEDIDO("id_cliente"))
        
        If Trim$("" & t_PEDIDO("obs_1")) <> "" Then
            If s_obs_1 <> "" Then s_obs_1 = s_obs_1 & vbCrLf
            s = Trim$("" & t_PEDIDO("obs_1"))
            s = substitui_caracteres(s, vbCr, " ")
            s = substitui_caracteres(s, vbLf, " ")
            s_obs_1 = s_obs_1 & s
            End If
        End If
    
    s = "SELECT pedido, fabricante, produto FROM t_PEDIDO_ITEM WHERE (pedido='" & Trim$(pedido) & "')"
    If t_PEDIDO_ITEM.State <> adStateClosed Then t_PEDIDO_ITEM.Close
    t_PEDIDO_ITEM.Open s, dbcAssist, , , adCmdText
    If t_PEDIDO_ITEM.EOF Then
        If s_erro <> "" Then s_erro = s_erro & vbCrLf
        s_erro = s_erro & "Não foi encontrado nenhum produto relacionado ao pedido " & Trim$(pedido) & "!!"
        End If
        
'   ENCONTROU ERRO ?
    If s_erro <> "" Then
        strMsgErro = s_erro
        GoSub OIPA_FECHA_TABELAS
        lblInfoAguarde.Visible = False
        Screen.MousePointer = vbDefault
        Screen.ActiveForm.MousePointer = vbDefault
        Me.Refresh
        cmdConfirma.Enabled = True
        Exit Function
        End If
        

'   OBTÉM DADOS DO DESTINATÁRIO DA NOTA
    s = "SELECT * FROM t_CLIENTE WHERE (id='" & s_id_cliente & "')"
    t_DESTINATARIO.Open s, dbcAssist, , , adCmdText
    If t_DESTINATARIO.EOF Then
        strMsgErro = "Cliente com nº registro " & s_id_cliente & " não foi encontrado!!"
        GoSub OIPA_FECHA_TABELAS
        lblInfoAguarde.Visible = False
        Screen.MousePointer = vbDefault
        Screen.ActiveForm.MousePointer = vbDefault
        Me.Refresh
        cmdConfirma.Enabled = True
        Exit Function
        End If


    s_endereco = UCase$(Trim$("" & t_DESTINATARIO("endereco")))
    s_aux = UCase$(Trim$("" & t_DESTINATARIO("endereco_numero")))
    If s_aux <> "" Then s_endereco = s_endereco & ", " & s_aux
    s_aux = UCase$(Trim$("" & t_DESTINATARIO("endereco_complemento")))
    If s_aux <> "" Then s_endereco = s_endereco & " " & s_aux

'   BAIRRO
    s_bairro = UCase$(Trim$("" & t_DESTINATARIO("bairro")))

'   CEP
    s_cep = Trim$("" & t_DESTINATARIO("cep"))

'   CIDADE
    s_cidade = UCase$(Trim$("" & t_DESTINATARIO("cidade")))

'   UF
    s_uf = UCase$(Trim$("" & t_DESTINATARIO("uf")))
    strEndClienteUf = UCase$(Trim$("" & t_DESTINATARIO("uf")))

'   NOME/RAZÃO SOCIAL DO CLIENTE
    s_nome = UCase$(Trim$("" & t_DESTINATARIO("nome")))

'   CNPJ/CPF
    s_cnpj_cpf = Trim$("" & t_DESTINATARIO("cnpj_cpf"))
    strCNPJCPF = s_cnpj_cpf

'   INSCRIÇÃO ESTADUAL
    If UCase$(Trim$("" & t_DESTINATARIO("tipo"))) = ID_PJ Then
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("ie")))
    Else
        s_ie_rg = UCase$(Trim$("" & t_DESTINATARIO("rg")))
        End If
    
    s_end_linha_1 = s_endereco
    If (s_end_linha_1 <> "") And (s_bairro <> "") Then s_end_linha_1 = s_end_linha_1 & "  -  "
    s_end_linha_1 = s_end_linha_1 & s_bairro
    
    s_end_linha_2 = s_cidade
    If (s_end_linha_2 <> "") And (s_uf <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & s_uf
    If (s_end_linha_2 <> "") And (s_cep <> "") Then s_end_linha_2 = s_end_linha_2 & "  -  "
    s_end_linha_2 = s_end_linha_2 & cep_formata(s_cep)
        
    If (s_end_linha_1 <> "") And (s_end_linha_2 <> "") Then s_end_linha_1 = s_end_linha_1 & vbCrLf
    
    s_info = s_nome & vbCrLf
    
    If s_cnpj_cpf <> "" Then s_info = s_info & "CNPJ/CPF: " & cnpj_cpf_formata(s_cnpj_cpf) & vbCrLf
    If s_ie_rg <> "" Then s_info = s_info & "IE/RG: " & s_ie_rg & vbCrLf
            
    s_info = s_info & _
             s_end_linha_1 & s_end_linha_2 & _
             s_end_entrega & vbCrLf & vbCrLf & _
             "OBSERVAÇÕES I" & vbCrLf & _
             s_obs_1
    
    GoSub OIPA_FECHA_TABELAS
    
    lblInfoAguarde.Visible = False
    Screen.MousePointer = vbDefault
    Screen.ActiveForm.MousePointer = vbDefault
    Me.Refresh
    cmdConfirma.Enabled = True

    strResposta = s_info
    obtem_info_pedido_assistencia = True
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIPA_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & Me.Name & " - " & NomeDestaRotina
    GoSub OIPA_FECHA_TABELAS
    lblInfoAguarde.Visible = False
    Screen.MousePointer = vbDefault
    Screen.ActiveForm.MousePointer = vbDefault
    Me.Refresh
    cmdConfirma.Enabled = True
    strMsgErro = s
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OIPA_FECHA_TABELAS:
'=================
  ' RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    bd_desaloca_recordset t_DESTINATARIO, True
    Return
    
End Function

Sub trata_pesquisa_produto()
    Dim s_Resposta As String
    Dim s_EndEntregaFormatado As String
    Dim s_EndEntregaUf As String
    Dim s_EndClienteUf As String
    Dim s_MsgErro As String

    s_Resposta = ""
    
End Sub

Private Sub c_pedido_assist_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdConfirma.SetFocus
        End If

End Sub

Private Sub c_pedido_assist_LostFocus()
    Dim s As String
    Dim s_MsgErro As String
    
    s_assist_Pedido = ""
    c_pedido_assist = Trim$(c_pedido_assist)
    s = normaliza_num_pedido(c_pedido_assist)
    
    If s <> "" Then
        c_pedido_assist = s
        If obtem_info_pedido_assistencia(c_pedido_assist, s_assist_Resposta, s_assist_EndEntregaFormatado, _
                                         s_assist_EndEntregaUf, s_assist_EndClienteUf, s_assist_Cliente_CNPJ_CPF, s_MsgErro) Then
            bln_assist_pedido_ok = True
            s_assist_Pedido = c_pedido_assist
            lblInfoAssist.Caption = s_assist_Resposta
            cmdConfirma.SetFocus
        Else
            If s_MsgErro <> "" Then aviso_erro s_MsgErro
            c_pedido_assist.SetFocus
            End If
        End If

End Sub

Private Sub cmdCancela_Click()
    
    bln_assist_pedido_ok = False
    Screen.MousePointer = vbDefault
    Screen.ActiveForm.MousePointer = vbDefault
    
    Unload Me
    
End Sub

Private Sub cmdConfirma_Click()
    If (Not bln_assist_pedido_ok) Or (Trim$(c_pedido_assist) = "") Then
        aviso_erro "Nº de pedido inválido!!!"
        c_pedido_assist.SetFocus
        Exit Sub
        End If
        
    Unload Me
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        cmdCancela_Click
        Exit Sub
        End If

End Sub

Private Sub Form_Load()
    
    bln_assist_pedido_ok = False
    
    lblInfoAssist.Caption = ""
    
End Sub
