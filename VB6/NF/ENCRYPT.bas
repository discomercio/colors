Attribute VB_Name = "mod_ENCRYPT"
Option Explicit

Function criptografa(ByVal origem As String, destino As String) As Boolean

Dim n As Integer
Dim i As Byte
Dim s_destino As String
Dim s_origem As String
Dim s As String
         
    On Error GoTo CRIPTO_ERRO
        
    criptografa = False
    destino = ""
    
    s_destino = ""
    i = CInt(Right$(CStr(Sqr(Len(origem))), 1))
    s_origem = texto_rotaciona_direita(origem, i)
    For n = 1 To Len(s_origem)
        s = Mid$(s_origem, n, 1)
        i = Asc(s)
        rotaciona_acima i, Len(s_origem) - (n - 1)
        s_destino = s_destino & Chr$(i)
        Next
    
    destino = s_destino

    criptografa = True


Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CRIPTO_ERRO:
'===========
    MsgBox CStr(Err) & Error$(Err), vbOKOnly + vbExclamation + vbSystemModal, "PARA SUA INFORMA플O"
    
    Exit Function
    
    
End Function






Function decriptografa(ByVal origem As String, destino As String) As Boolean

Dim n As Integer
Dim i As Byte
Dim s_destino As String
Dim s As String
         
    On Error GoTo DECRIPTO_ERRO
        
    decriptografa = False
    destino = ""
    
    s_destino = ""
    For n = 1 To Len(origem)
        s = Mid$(origem, n, 1)
        i = Asc(s)
        rotaciona_abaixo i, Len(origem) - (n - 1)
        s_destino = s_destino & Chr$(i)
        Next
    
    i = CInt(Right$(CStr(Sqr(Len(origem))), 1))
    destino = texto_rotaciona_esquerda(s_destino, i)
    
    decriptografa = True


Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DECRIPTO_ERRO:
'=============
    MsgBox CStr(Err) & Error$(Err), vbOKOnly + vbExclamation + vbSystemModal, "PARA SUA INFORMA플O"
    
    Exit Function
    
    
End Function


Sub rotaciona_acima(numero As Byte, ByVal incremento As Byte)
Dim n As Integer
        
    n = numero
    n = n + incremento
    If n <= 255 Then numero = n Else numero = n Mod 255
    
End Sub


Sub rotaciona_abaixo(numero As Byte, ByVal decremento As Byte)
Dim n As Integer

    n = numero
    n = n - decremento
    If n >= 0 Then numero = n Else numero = 256 - Abs(n)
    
End Sub



Function texto_rotaciona_direita(ByVal origem As String, ByVal casas As Integer) As String

Dim i As Byte
Dim s_destino As String
Dim s As String
         
    On Error GoTo TRD_TRATA_ERRO
        
    texto_rotaciona_direita = ""
    
    s_destino = origem
    For i = 1 To casas
        s = Right$(s_destino, 1)
        s_destino = Left$(s_destino, Len(s_destino) - 1)
        s_destino = s & s_destino
        Next
    
    texto_rotaciona_direita = s_destino

Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TRD_TRATA_ERRO:
'==============
    MsgBox CStr(Err) & Error$(Err), vbOKOnly + vbExclamation + vbSystemModal, "PARA SUA INFORMA플O"
    
    Exit Function
    

End Function



Function texto_rotaciona_esquerda(ByVal origem As String, ByVal casas As Integer) As String

Dim i As Byte
Dim s_destino As String
Dim s As String
         
    On Error GoTo TRE_TRATA_ERRO
        
    texto_rotaciona_esquerda = ""
    
    s_destino = origem
    For i = 1 To casas
        s = Left$(s_destino, 1)
        s_destino = Right$(s_destino, Len(s_destino) - 1)
        s_destino = s_destino & s
        Next
    
    texto_rotaciona_esquerda = s_destino

Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TRE_TRATA_ERRO:
'==============
    MsgBox CStr(Err) & Error$(Err), vbOKOnly + vbExclamation + vbSystemModal, "PARA SUA INFORMA플O"
    
    Exit Function
    

End Function



