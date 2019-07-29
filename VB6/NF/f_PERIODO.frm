VERSION 5.00
Begin VB.Form f_PERIODO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe o período"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame pn_UF 
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   3975
      Begin VB.TextBox c_UF 
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
         Height          =   360
         Left            =   720
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   600
      End
      Begin VB.Label l_tit_UF 
         AutoSize        =   -1  'True
         Caption         =   "UF"
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
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   255
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
      Left            =   225
      Picture         =   "f_PERIODO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
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
      Left            =   2820
      Picture         =   "f_PERIODO.frx":0252
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1425
   End
   Begin VB.Frame pn_campos 
      Height          =   945
      Left            =   240
      TabIndex        =   5
      Top             =   135
      Width           =   4020
      Begin VB.TextBox c_dt_termino 
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
         Height          =   360
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "01/01/2010"
         Top             =   420
         Width           =   1440
      End
      Begin VB.TextBox c_dt_inicio 
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
         Height          =   360
         Left            =   180
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "01/01/2010"
         Top             =   420
         Width           =   1440
      End
      Begin VB.Label l_tit_periodo 
         AutoSize        =   -1  'True
         Caption         =   "Período"
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
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   690
      End
      Begin VB.Label l_tit_ate 
         AutoSize        =   -1  'True
         Caption         =   "até"
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
         Left            =   1860
         TabIndex        =   6
         Top             =   495
         Width           =   285
      End
   End
End
Attribute VB_Name = "f_PERIODO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnResultadoFormOk As Boolean
Public dtInicioSelecionada As Date
Public dtTerminoSelecionada As Date
Public strUFSelecionada As String

Sub trata_botao_cancela()

    blnResultadoFormOk = False
    Unload Me

End Sub

Sub trata_botao_confirma()

Dim s As String

    On Error GoTo TBTCONF_TRATA_ERRO
    
'   DATA DE INÍCIO
    If Trim$(c_dt_inicio) = "" Then
        aviso_erro "Informe a data de início do período!!"
        c_dt_inicio.SetFocus
        Exit Sub
        End If
        
    If Not IsDate(c_dt_inicio) Then
        aviso_erro "Informe uma data de início válida!!"
        c_dt_inicio.SetFocus
        Exit Sub
        End If
    
'   DATA DE TÉRMINO
    If Trim$(c_dt_termino) = "" Then
        aviso_erro "Informe a data de término do período!!"
        c_dt_termino.SetFocus
        Exit Sub
        End If
        
    If Not IsDate(c_dt_termino) Then
        aviso_erro "Informe uma data de término válida!!"
        c_dt_termino.SetFocus
        Exit Sub
        End If
    
'   DATA DE TÉRMINO POSTERIOR À DATA DE INÍCIO?
    If CDate(c_dt_inicio) > CDate(c_dt_termino) Then
        aviso_erro "A data de término é anterior à data de início!!"
        c_dt_termino.SetFocus
        Exit Sub
        End If
    
    dtInicioSelecionada = CDate(c_dt_inicio)
    dtTerminoSelecionada = CDate(c_dt_termino)
    
'   UF
    If Not UF_ok(c_UF) Then
        aviso_erro "UF inválida!!"
        c_UF.SetFocus
        Exit Sub
        End If
        
    strUFSelecionada = c_UF
        
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

Private Sub c_dt_inicio_GotFocus()

    c_dt_inicio.SelStart = 0
    c_dt_inicio.SelLength = Len(c_dt_inicio)

End Sub


Private Sub c_dt_inicio_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        data_ok c_dt_inicio
        If IsDate(c_dt_inicio) Then c_dt_termino.SetFocus
        Exit Sub
        End If
        
    KeyAscii = filtra_data(KeyAscii)

End Sub


Private Sub c_dt_inicio_LostFocus()

    c_dt_inicio = Trim$(c_dt_inicio)
    If c_dt_inicio = "" Then Exit Sub
    
    data_ok c_dt_inicio

End Sub

Private Sub c_dt_termino_GotFocus()

    c_dt_termino.SelStart = 0
    c_dt_termino.SelLength = Len(c_dt_termino)

End Sub


Private Sub c_dt_termino_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        data_ok c_dt_termino
        If IsDate(c_dt_termino) Then c_UF.SetFocus
        Exit Sub
        End If
        
    KeyAscii = filtra_data(KeyAscii)

End Sub


Private Sub c_dt_termino_LostFocus()

    c_dt_termino = Trim$(c_dt_termino)
    If c_dt_termino = "" Then Exit Sub
    
    data_ok c_dt_termino

End Sub



Private Sub c_UF_GotFocus()

    c_UF.SelStart = 0
    c_UF.SelLength = Len(c_UF)

End Sub

Private Sub c_UF_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        trata_botao_confirma
        Exit Sub
        End If
    
    KeyAscii = filtra_letra(KeyAscii)
    If KeyAscii <> 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub c_UF_LostFocus()

    c_UF = UCase$(Trim$(c_UF))

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

    c_dt_inicio = ""
    c_dt_termino = ""

End Sub


