Attribute VB_Name = "mod_PAINEL"
Option Explicit

Type POINTAPI
    X As Long
    Y As Long
    End Type

' MOVIMENTAÇÃO DO PAINEL
Global Const WM_NCLBUTTONDOWN = &HA1
Global Const HTCAPTION = 2

Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' PolyFill() Modes
Const ALTERNATE = 1
Const WINDING = 2
Const POLYFILL_LAST = 2

Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


Sub movimenta_painel(p As Form, Button As Integer)
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


Sub recorta_tela(r() As POINTAPI, ByVal Handle As Long)
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

