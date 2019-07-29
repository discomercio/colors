VERSION 5.00
Begin VB.Form f_MAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Criptografia"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton b_decripto 
      Caption         =   "&Descriptografa"
      Height          =   420
      Left            =   4410
      TabIndex        =   5
      Top             =   1020
      Width           =   1320
   End
   Begin VB.CommandButton b_cripto 
      Caption         =   "&Criptografa"
      Height          =   420
      Left            =   4410
      TabIndex        =   4
      Top             =   300
      Width           =   1320
   End
   Begin VB.TextBox c_cripto 
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Top             =   1065
      Width           =   4170
   End
   Begin VB.TextBox c_decripto 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   345
      Width           =   4170
   End
   Begin VB.Label l_cripto 
      AutoSize        =   -1  'True
      Caption         =   "Texto Criptografado"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   825
      Width           =   1395
   End
   Begin VB.Label l_decripto 
      AutoSize        =   -1  'True
      Caption         =   "Texto Descriptografado"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "f_MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub b_cripto_Click()
Dim s As String

    c_cripto = ""
    
    If c_decripto = "" Then Exit Sub
    
    If criptografa(c_decripto, s) Then c_cripto = s
    
End Sub


Private Sub b_decripto_Click()
Dim s As String

    c_decripto = ""
    
    If c_cripto = "" Then Exit Sub
    
    If decriptografa(c_cripto, s) Then c_decripto = s

End Sub


Private Sub c_cripto_GotFocus()

    With c_cripto
        .SelStart = 0
        .SelLength = Len(.Text)
        End With

End Sub

Private Sub c_cripto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then b_decripto_Click
    
End Sub


Private Sub c_decripto_GotFocus()

    With c_decripto
        .SelStart = 0
        .SelLength = Len(.Text)
        End With
        
End Sub

Private Sub c_decripto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then b_cripto_Click

End Sub


