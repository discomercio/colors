VERSION 5.00
Begin VB.Form f_DATA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe a data"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame pn_campos 
      Height          =   1065
      Left            =   225
      TabIndex        =   3
      Top             =   135
      Width           =   4020
      Begin VB.TextBox c_data 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1365
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "01/01/2010"
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label l_tit_data 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   480
         Width           =   420
      End
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
      Left            =   2820
      Picture         =   "f_DATA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
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
      Left            =   225
      Picture         =   "f_DATA.frx":019D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1425
   End
End
Attribute VB_Name = "f_DATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnResultadoFormOk As Boolean
Public dtDataSelecionada As Date

Sub trata_botao_cancela()

    blnResultadoFormOk = False
    Unload Me

End Sub

Sub trata_botao_confirma()

Dim s As String

    On Error GoTo TBTCONF_TRATA_ERRO
    
    If Trim$(c_data) = "" Then
        aviso_erro "Informe uma data!!"
        c_data.SetFocus
        Exit Sub
        End If
        
    If Not IsDate(c_data) Then
        aviso_erro "Informe uma data válida!!"
        c_data.SetFocus
        Exit Sub
        End If
    
    dtDataSelecionada = CDate(c_data)
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


Private Sub c_data_GotFocus()

    c_data.SelStart = 0
    c_data.SelLength = Len(c_data)

End Sub


Private Sub c_data_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        data_ok c_data
        If IsDate(c_data) Then trata_botao_confirma
        Exit Sub
        End If
        
    KeyAscii = filtra_data(KeyAscii)
    
End Sub


Private Sub c_data_LostFocus()

    c_data = Trim$(c_data)
    If c_data = "" Then Exit Sub
    
    data_ok c_data
    
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

    c_data = ""
    
End Sub


