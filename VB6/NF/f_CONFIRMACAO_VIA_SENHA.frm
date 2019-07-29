VERSION 5.00
Begin VB.Form f_CONFIRMACAO_VIA_SENHA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmação da operação"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9510
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame pnSenha 
      Caption         =   "Digite a senha para confirmar a operação"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   4620
      Width           =   4710
      Begin VB.TextBox c_senha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   375
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   345
         Width           =   3945
      End
   End
   Begin VB.Frame pnMensagem 
      Caption         =   "Mensagem"
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
      Height          =   4185
      Left            =   218
      TabIndex        =   3
      Top             =   225
      Width           =   9075
      Begin VB.Label l_mensagem_informativa 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3420
         Left            =   240
         TabIndex        =   5
         Top             =   435
         Width           =   8535
         WordWrap        =   -1  'True
      End
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
      Left            =   2745
      Picture         =   "f_CONFIRMACAO_VIA_SENHA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5865
      Width           =   1425
   End
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
      Left            =   5340
      Picture         =   "f_CONFIRMACAO_VIA_SENHA.frx":0252
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5865
      Width           =   1425
   End
End
Attribute VB_Name = "f_CONFIRMACAO_VIA_SENHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strMensagemInformativa As String
Public strSenhaCorreta As String
Public blnResultadoFormOk As Boolean

Sub trata_botao_cancela()

    blnResultadoFormOk = False
    Unload Me

End Sub

Sub trata_botao_confirma()

Dim s As String

    On Error GoTo TBTCONF_TRATA_ERRO
    
    If Trim$(c_senha) = "" Then
        aviso_erro "Digite a senha para confirmar a operação!"
        c_senha.SetFocus
        Exit Sub
        End If
    
    If UCase$(Trim$(c_senha)) <> UCase$(Trim$(strSenhaCorreta)) Then
        aviso_erro "Senha não confere!"
        c_senha.SetFocus
        Exit Sub
        End If
        
    blnResultadoFormOk = True
    
    Unload Me
    
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TBTCONF_TRATA_ERRO:
'==================
    s = CStr(Err) & ": " & Error$(Err)
    aviso_erro s
    Exit Sub
    
End Sub

Private Sub c_senha_GotFocus()

    c_senha.SelStart = 0
    c_senha.SelLength = Len(c_senha)
    
End Sub

Private Sub c_senha_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdConfirma.SetFocus
        trata_botao_confirma
        End If

End Sub


Private Sub cmdCancela_Click()

    trata_botao_cancela
    
End Sub


Private Sub cmdConfirma_Click()

    trata_botao_confirma

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

'   ESC
    If KeyAscii = 27 Then
        KeyAscii = 0
        trata_botao_cancela
        Exit Sub
        End If

End Sub

Private Sub Form_Load()

    blnResultadoFormOk = False
    
    c_senha.Text = ""
    l_mensagem_informativa.Caption = strMensagemInformativa
    
End Sub


