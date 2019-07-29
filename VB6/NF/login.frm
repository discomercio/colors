VERSION 5.00
Begin VB.Form f_LOGIN 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONTROLE DE ACESSO"
   ClientHeight    =   2325
   ClientLeft      =   4245
   ClientTop       =   3585
   ClientWidth     =   2670
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "login.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   178
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox c_senha 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   150
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Tag             =   "senha"
      Text            =   "1234567890"
      Top             =   1050
      Width           =   2400
   End
   Begin VB.TextBox c_usuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   150
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "usuário"
      Text            =   "MMMMMMMMMM"
      Top             =   330
      Width           =   2400
   End
   Begin VB.PictureBox b_NAO 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "login.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   1035
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1740
      Width           =   1035
   End
   Begin VB.PictureBox b_OK 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1545
      Picture         =   "login.frx":06CF
      ScaleHeight     =   480
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   1740
      Width           =   1035
   End
   Begin VB.Label l_senha 
      AutoSize        =   -1  'True
      Caption         =   "senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   840
      Width           =   435
   End
   Begin VB.Label l_usuario 
      AutoSize        =   -1  'True
      Caption         =   "usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "f_LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub b_NAO_Click()

'   ~~~
    End
'   ~~~

End Sub

Private Sub b_ok_Click()

Dim rs
Dim sx
Dim s As String
Dim senha_real As String

    On Error GoTo BOK_TRATA_ERRO
    
    c_usuario = UCase$(Trim$(c_usuario))
    c_senha = UCase$(Trim$(c_senha))

    If c_usuario = "" Then
        c_usuario.SetFocus
        Exit Sub
        End If

    If c_senha = "" Then
        c_senha.SetFocus
        Exit Sub
        End If

'   OBTÉM DADOS DO LOGIN
    With usuario
        .id = c_usuario
        .senha = c_senha
        End With
    
    Screen.MousePointer = vbHourglass
    c_usuario.BackColor = COR_CINZA
    c_senha.BackColor = COR_CINZA
    b_OK.Visible = False
    b_NAO.Visible = False

    
' VERIFICA SENHA E CATEGORIA DO USUÁRIO
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "verificando usuário"
    s = "SELECT * FROM  t_USUARIO WHERE (usuario = '" & usuario.id & "')"
    Set rs = dbc.Execute(s)
    
    If rs.EOF Then
        GoSub BOK_ACESSO_NEGADO
        aviso_erro "USUÁRIO NÃO CADASTRADO!!"
      ' ENCERRA O PROGRAMA
        BD_Fecha
        BD_CEP_Fecha
       '~~~
        End
       '~~~
        End If
        
    If rs("bloqueado") <> 0 Then
        GoSub BOK_ACESSO_NEGADO
        aviso_erro "ACESSO NÃO PERMITIDO!!"
      ' ENCERRA O PROGRAMA
        BD_Fecha
        BD_CEP_Fecha
       '~~~
        End
       '~~~
        End If
        
    If IsNull(rs("dt_ult_alteracao_senha")) Then
        GoSub BOK_ACESSO_NEGADO
        s = "ACESSO NEGADO!!" & _
            vbCrLf & vbCrLf & "A senha está expirada e deve ser alterada antes de acessar este módulo!!"
        aviso_erro s
      ' ENCERRA O PROGRAMA
        BD_Fecha
        BD_CEP_Fecha
       '~~~
        End
       '~~~
        End If
        
    s = "SELECT" & _
            " tPU.usuario" & _
        " FROM t_PERFIL_X_USUARIO tPU" & _
            " INNER JOIN t_PERFIL tP ON (tPU.id_perfil=tP.id)" & _
            " INNER JOIN t_PERFIL_ITEM tPI ON (tP.id=tPI.id_perfil)" & _
        " WHERE" & _
            " (tPI.id_operacao = " & CStr(OP_CEN_MODULO_NF_ACESSO_AO_MODULO) & ")" & _
            " AND (tPU.usuario = '" & usuario.id & "')"
    Set sx = dbc.Execute(s)
    If Not sx.EOF Then
        usuario.perfil_acesso_ok = True
    Else
        usuario.perfil_acesso_ok = False
        End If
        
    senha_real = ""
    s = Trim("" & rs("datastamp"))
    decodifica_dado s, senha_real
    If senha_real <> usuario.senha Then
        GoSub BOK_ACESSO_NEGADO
        s = "ACESSO INVÁLIDO!!"
        aviso_erro s
      ' ENCERRA O PROGRAMA
        BD_Fecha
        BD_CEP_Fecha
       '~~~
        End
       '~~~
        End If
           

  ' VERSAO OK ?
    If Not versao_esta_atualizada() Then
        GoSub BOK_ACESSO_NEGADO
        aviso_erro "Versão do aplicativo está desatualizada!!"
      ' ENCERRA O PROGRAMA
        BD_Fecha
        BD_CEP_Fecha
       '~~~
        End
       '~~~
        End If


    c_usuario.BackColor = COR_VERDE_ESCURO
    c_senha.BackColor = COR_VERDE_ESCURO

    aguarde INFO_NORMAL, m_id
    Unload Me

    Screen.MousePointer = vbDefault

Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BOK_ACESSO_NEGADO:
'=================
    c_usuario.BackColor = COR_VERMELHO
    c_senha.BackColor = COR_VERMELHO
    Return
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BOK_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err)
    aviso_erro s
   'ENCERRA O PROGRAMA
   '~~~
    End
   '~~~
    Exit Sub
    
    
End Sub

Function versao_esta_atualizada()

'   VERIFICA SE A VERSAO DO APLICATIVO ESTA ATUALIZADA
Dim s As String
Dim v As String
Dim rs

    versao_esta_atualizada = False
    
    s = "SELECT versao FROM t_VERSAO" & _
        " WHERE (modulo = '" & Trim$(App.Title) & "')"
    
    v = ""
    Set rs = dbc.Execute(s)
    If Not rs.EOF Then v = Trim$("" & rs("versao"))
        
    versao_esta_atualizada = (v = m_id_versao)
    
End Function

Private Sub Form_Activate()

Dim s As String

    On Error GoTo FA_TRATA_ERRO
    
    If painel_ativo Is Me Then Exit Sub
    
    Set painel_ativo = Me

    Screen.MousePointer = Default

    c_usuario.SetFocus

Exit Sub



FA_TRATA_ERRO:
'~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    MsgBox s, vbOKOnly + vbCritical, App.Title
    Exit Sub

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        b_NAO_Click
        Exit Sub
        End If
        
End Sub

Private Sub Form_Load()

Dim s As String

    On Error GoTo FL_TRATA_ERRO
    
 '> CENTRALIZA O PAINEL NA TELA
    ScaleMode = vbPixels
    Line (1, 1)-(ScaleWidth - 3, b_OK.top - 8), COR_BRANCO, B
    Line (2, 2)-(ScaleWidth - 2, b_OK.top - 7), COR_CINZA_ESCURO, B

    c_usuario = ""
    c_senha = ""
    
    left = painel_principal.left + (painel_principal.Width - Width) \ 2
    top = painel_principal.top + (painel_principal.Height - Height) \ 2
    
Exit Sub



FL_TRATA_ERRO:
'~~~~~~~~~~~~~
    s = CStr(Err) & ": " & Error$(Err)
    MsgBox s, vbOKOnly + vbCritical, App.Title
    Exit Sub
    
End Sub

Private Sub c_senha_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        b_ok_Click
        End If

End Sub

Private Sub c_usuario_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        b_ok_Click
        End If

End Sub

