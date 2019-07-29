VERSION 5.00
Begin VB.Form f_PROCESSANDO 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Processando ..."
   ClientHeight    =   1860
   ClientLeft      =   2490
   ClientTop       =   2280
   ClientWidth     =   5865
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
   HelpContextID   =   6
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   124
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   391
   Begin VB.PictureBox base 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   105
      ScaleHeight     =   1095
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   60
      Width           =   4485
      Begin VB.Image b_CANCELA 
         Height          =   480
         Left            =   1605
         Top             =   465
         Width           =   1035
      End
   End
   Begin VB.Image i_b_on 
      Height          =   480
      Left            =   4680
      Picture         =   "processo.frx":0000
      Top             =   75
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image i_b_off 
      Height          =   480
      Left            =   4680
      Picture         =   "processo.frx":03BB
      Top             =   615
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image i_b_down 
      Height          =   480
      Left            =   4680
      Picture         =   "processo.frx":0780
      Top             =   1185
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "f_PROCESSANDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ________________________________________________________________________________
'|
'|  A VARIÁVEL "PROCESSO_CANCELADO" (TIPO INTEGER) DEVE SER DECLARADA
'|  COMO GLOBAL EM ALGUM DOS MÓDULOS DO PROJETO
'|

Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

' MOVIMENTAÇÃO DO PAINEL
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type POINTAPI
    X As Long
    Y As Long
    End Type

' MOVIMENTAÇÃO DO PAINEL
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

' PolyFill() Modes
Const ALTERNATE = 1
Const WINDING = 2
Const POLYFILL_LAST = 2

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public processo_cancelado As Boolean

Public Sub exibe_painel()

Dim iLeft, iTop As Integer


    On Error Resume Next


    processo_cancelado = False


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   POSICIONA NO CENTRO DO PAINEL
' 
    iLeft = painel_ativo.Left + (painel_ativo.Width - Width) / 2
    iTop = painel_ativo.Top + (painel_ativo.Height - Height) / 2
    Move iLeft, iTop
    Show
    Refresh
    
End Sub


Public Sub fecha_painel()

    Unload Me
    
End Sub


Private Sub b_cancela_Click()
  
    processo_cancelado = True

End Sub

Private Sub b_Cancela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Aperta_Cursor

    If b_CANCELA.Picture <> i_b_down Then b_CANCELA.Picture = i_b_down

End Sub

Private Sub b_CANCELA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    b_CANCELA.MousePointer = vbDefault
    Screen.ActiveForm.MousePointer = vbDefault
    Screen.MousePointer = vbDefault

  ' BOTÃO PRESSIONADO ?
    If Button = 1 Then If b_CANCELA.Picture = i_b_down Then Exit Sub
        
  ' B_CANCELA LIGADO ?
    If b_CANCELA.Picture <> i_b_on Then b_CANCELA.Picture = i_b_on


End Sub

Private Sub b_Cancela_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Solta_cursor

    If b_CANCELA.Picture <> i_b_off Then b_CANCELA.Picture = i_b_off

End Sub



Private Sub Aperta_Cursor()

Dim lpPoint As POINTAPI
Dim l As Long

' 
'   SOBRE O CURSOR DO MOUSE DE ALGUNS PIXELS,
'   DANDO A IMPRESSÃO QUE APERTOU UM BOTÃO
' 
    
    l = GetCursorPos(lpPoint)
    l = SetCursorPos(lpPoint.X, lpPoint.Y - 2)
    

End Sub

Private Sub Solta_cursor()
    
Dim lpPoint As POINTAPI
' 
'   DESCE O CURSOR DO MOUSE DE ALGUNS PIXELS DANDO A IMPRESSÃO QUE SOLTOU UM BOTÃO
' 
   
    GetCursorPos lpPoint
    SetCursorPos lpPoint.X, lpPoint.Y + 2

End Sub


Private Sub movimenta_painel(p As Form, Button As Integer)
' ______________________________________________________________________________
'|
'|  MOVIMENTA O FORM INDICADO PELO PARÂMETRO.
'|  ESTA ROTINA DEVE SER CHAMADA A PARTIR DO EVENTO MOUSEMOVE, SENDO QUE O
'|  PARÂMETRO BUTTON É DO PRÓPRIO EVENTO.
'|  TIPICAMENTE, A CHAMADA SERÁ:  movimenta_painel Me, Button
'|

Dim lngReturnValue As Long

  ' É O BOTÃO ESQUERDO ?
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(p.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If


End Sub


Private Sub base_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If b_CANCELA.Picture <> i_b_off Then b_CANCELA.Picture = i_b_off

  ' MOVIMENTA O PAINEL
    movimenta_painel Me, Button

End Sub


Private Sub Form_Activate()

    If cor_fundo_padrao <> "" Then
        Me.BackColor = cor_fundo_padrao
        End If
        
End Sub

Private Sub Form_Deactivate()

    ZOrder

End Sub

Private Sub Form_Load()

Const LARGURA_SOMBRA = 6
Const ALTURA_SOMBRA = 6
Const OFF_X_SOMBRA = 10
Const OFF_Y_SOMBRA = 10
Const np = 9
Dim r(np - 1) As POINTAPI
Dim dw_a As Integer
Dim i As Integer
Dim X As Integer
Dim Y As Integer
Dim W As Integer
Dim h As Integer
Dim s As String

    b_CANCELA.Picture = i_b_off

    base.Appearance = 0
    base.AutoRedraw = True
    base.BorderStyle = 0
    base.ScaleMode = vbPixels
    base.BackColor = COR_FRENTE_PAINEL
    base.TabStop = False
    
    ScaleMode = vbPixels
    BackColor = COR_SOMBRA_PAINEL
    
    base.Top = 0
    base.Left = 0
    
    s = "P R O C E S S A N D O ..."
    With base
        .FontName = "Courier New"
        .FontSize = 16
        .FontBold = True
        .ForeColor = COR_AZUL_ESCURO
        
        .Width = .TextWidth(s) + 80
        
        Y = 5
        .CurrentX = (.ScaleWidth - .TextWidth(s)) \ 2
        .CurrentY = Y
        
        base.Print s
        End With
    
    b_CANCELA.Top = Y + base.TextHeight("X") + Y
    b_CANCELA.Left = (base.ScaleWidth - b_CANCELA.Width) \ 2
    
    base.Height = b_CANCELA.Top + b_CANCELA.Height + 10

   
    dw_a = base.DrawWidth
    base.DrawWidth = 2
    base.ForeColor = COR_BORDA_PAINEL
    base.Line (base.DrawWidth \ 2, base.DrawWidth \ 2)-(base.ScaleWidth - maior(1, base.DrawWidth \ 2), base.ScaleHeight - maior(1, base.DrawWidth \ 2)), COR_BORDA_PAINEL, B
    base.DrawWidth = dw_a
    
    Width = (base.Width + LARGURA_SOMBRA) * Screen.TwipsPerPixelX
    Height = (base.Height + ALTURA_SOMBRA) * Screen.TwipsPerPixelY
    
    
  ' RECORTA TELA
    i = 0: With r(i): .X = base.Left: .Y = base.Top: End With
    i = i + 1: With r(i): .X = base.Left + base.Width: .Y = base.Top: End With
    i = i + 1: With r(i): .X = base.Left + base.Width: .Y = base.Top + OFF_Y_SOMBRA: End With
    i = i + 1: With r(i): .X = base.Left + base.Width + LARGURA_SOMBRA: .Y = base.Top + OFF_Y_SOMBRA: End With
    i = i + 1: With r(i): .X = base.Left + base.Width + LARGURA_SOMBRA: .Y = base.Top + base.Height + ALTURA_SOMBRA: End With
    i = i + 1: With r(i): .X = base.Left + OFF_X_SOMBRA: .Y = base.Top + base.Height + ALTURA_SOMBRA: End With
    i = i + 1: With r(i): .X = base.Left + OFF_X_SOMBRA: .Y = base.Top + base.Height: End With
    i = i + 1: With r(i): .X = base.Left: .Y = base.Top + base.Height: End With
    i = i + 1: With r(i): .X = base.Left: .Y = base.Top: End With

    recorta_tela r(), Me.hwnd
        
    
End Sub

Private Sub recorta_tela(r() As POINTAPI, ByVal Handle As Long)
' ____________________________________________________________________________________________________________________________
'|
'|  EXECUTA AS FUNÇÕES DA API PARA RECORTAR A TELA.
'|

Dim rgnNew As Long
Dim resultado As Long
Dim np As Integer

    np = UBound(r) - LBound(r) + 1

    rgnNew = CreatePolygonRgn(r(LBound(r)), np, WINDING)
    resultado = SetWindowRgn(Handle, rgnNew, True)
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    b_CANCELA.MousePointer = vbDefault
    Screen.ActiveForm.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass

  ' MOVIMENTA O PAINEL
    movimenta_painel Me, Button

End Sub

