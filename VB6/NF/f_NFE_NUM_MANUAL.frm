VERSION 5.00
Begin VB.Form f_NFE_NUM_MANUAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preenchimento manual do nº da NFe"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   8520
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
      Left            =   3330
      Picture         =   "f_NFE_NUM_MANUAL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2025
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
      Left            =   1440
      Picture         =   "f_NFE_NUM_MANUAL.frx":019D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2025
      Width           =   1425
   End
   Begin VB.TextBox c_numero_NFe 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   0
      Top             =   1440
      Width           =   1290
   End
   Begin VB.TextBox c_serie_NFe 
      Alignment       =   2  'Center
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
      Height          =   330
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   900
      Width           =   1290
   End
   Begin VB.TextBox c_emitente 
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
      Height          =   330
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   6675
   End
   Begin VB.Label l_numero_NFe 
      AutoSize        =   -1  'True
      Caption         =   "Nº NFe"
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
      Left            =   690
      TabIndex        =   7
      Top             =   1500
      Width           =   630
   End
   Begin VB.Label l_emitente 
      AutoSize        =   -1  'True
      Caption         =   "Emitente"
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
      Left            =   570
      TabIndex        =   6
      Top             =   420
      Width           =   750
   End
   Begin VB.Label l_serie_NFe 
      AutoSize        =   -1  'True
      Caption         =   "Série"
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
      Left            =   870
      TabIndex        =   5
      Top             =   960
      Width           =   450
   End
End
Attribute VB_Name = "f_NFE_NUM_MANUAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngNFeUltSerieEmitida As Long
Public lngNFeUltNumeroNfEmitido As Long
Public lngNFeSerieManual As Long
Public lngNFeNumeroNfManual As Long
Public strDescricaoEmitente As String
Public blnResultadoFormOk As Boolean

Sub trata_botao_cancela()

    blnResultadoFormOk = False
    Unload Me

End Sub


Sub trata_botao_confirma()

Dim s As String
Dim lngSerie As Long
Dim lngNumero As Long

    On Error GoTo TBTCONF_TRATA_ERRO
    
    If Not IsNumeric(c_serie_NFe) Then
        aviso_erro "Série da NFe foi preenchida em formato inválido!"
        c_serie_NFe.SetFocus
        Exit Sub
        End If

    If Not IsNumeric(c_numero_NFe) Then
        aviso_erro "Número da NFe foi preenchida em formato inválido!"
        c_numero_NFe.SetFocus
        Exit Sub
        End If
        
    lngSerie = CLng(c_serie_NFe)
    lngNumero = CLng(c_numero_NFe)
    
    If (lngSerie <> lngNFeUltSerieEmitida) Then
        aviso_erro "Série da NFe não pode ser diferente da série atual!"
        c_serie_NFe.SetFocus
        Exit Sub
        End If
        
    If (lngNumero > lngNFeUltNumeroNfEmitido) Then
        aviso_erro "Não é possível usar um número de NFe maior que o último número emitido automaticamente!"
        c_numero_NFe.SetFocus
        Exit Sub
        End If
        
    lngNFeSerieManual = lngSerie
    lngNFeNumeroNfManual = lngNumero
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


Private Sub c_numero_NFe_GotFocus()

    c_numero_NFe.SelStart = 0
    c_numero_NFe.SelLength = Len(c_numero_NFe)
    
End Sub

Private Sub c_numero_NFe_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdConfirma.SetFocus
        trata_botao_confirma
        Exit Sub
        End If
        
    KeyAscii = filtra_numerico(KeyAscii)
    
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

    lngNFeSerieManual = 0
    lngNFeNumeroNfManual = 0
    blnResultadoFormOk = False
    
    c_emitente = strDescricaoEmitente
    c_serie_NFe = CStr(lngNFeUltSerieEmitida)
    
End Sub


