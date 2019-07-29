VERSION 5.00
Begin VB.Form f_CHECKBOX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opções"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton b_desmarcar 
      Caption         =   "&Desmarcar todos"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2940
      Width           =   1485
   End
   Begin VB.CommandButton b_marcar 
      Caption         =   "&Marcar todos"
      Height          =   375
      Left            =   165
      TabIndex        =   1
      Top             =   2940
      Width           =   1485
   End
   Begin VB.CommandButton b_cancela 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6645
      TabIndex        =   4
      Top             =   2940
      Width           =   1215
   End
   Begin VB.CommandButton b_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   2940
      Width           =   1215
   End
   Begin VB.ListBox lista 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      ItemData        =   "f_CHKBOX.frx":0000
      Left            =   75
      List            =   "f_CHKBOX.frx":000D
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   90
      Width           =   7875
   End
End
Attribute VB_Name = "f_CHECKBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim confirmou As Boolean

Public Function executa_selecao(ByRef v_xls_a_selecionar() As String, ByRef v_xls_selecionada() As String) As Boolean
    
Dim i As Integer

    executa_selecao = False
    
    ReDim v_xls_selecionada(0)
    v_xls_selecionada(UBound(v_xls_selecionada)) = ""
    
    lista.Clear
    For i = LBound(v_xls_a_selecionar) To UBound(v_xls_a_selecionar)
        If v_xls_a_selecionar(i) <> "" Then
            lista.AddItem v_xls_a_selecionar(i)
            End If
        Next

    For i = (lista.ListCount - 1) To 0 Step -1
        lista.Selected(i) = True
        Next
        
    lista.Refresh
    
    confirmou = False
    Me.Show vbModal
    
    If Not confirmou Then Exit Function
    
'   RETORNA A LISTA DOS ITENS SELECIONADOS
    For i = 0 To (lista.ListCount - 1)
        If lista.Selected(i) Then
            If Trim$(v_xls_selecionada(UBound(v_xls_selecionada))) <> "" Then ReDim Preserve v_xls_selecionada(UBound(v_xls_selecionada) + 1)
            v_xls_selecionada(UBound(v_xls_selecionada)) = lista.List(i)
            End If
        Next
        
    executa_selecao = True
    
End Function


Private Sub b_cancela_Click()

    Hide
    
End Sub

Private Sub b_desmarcar_Click()

Dim i As Integer

    For i = (lista.ListCount - 1) To 0 Step -1
        lista.Selected(i) = False
        Next
        
    lista.Refresh

End Sub

Private Sub b_marcar_Click()

Dim i As Integer

    For i = (lista.ListCount - 1) To 0 Step -1
        lista.Selected(i) = True
        Next
        
    lista.Refresh
    
End Sub

Private Sub b_ok_Click()

    confirmou = True
    Hide
    
End Sub


Private Sub Form_Activate()

    If cor_fundo_padrao <> "" Then
        Me.BackColor = cor_fundo_padrao
        End If
        
End Sub

