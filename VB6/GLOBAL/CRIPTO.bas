Attribute VB_Name = "mod_CRIPTO"
Option Explicit

' ========================== IMPORTANTE ==================================================================
' Estas rotinas de criptografia, além de criptografar/descriptografar a senha,
' convertem os caracteres da senha criptograda para códigos em hexadecimal.
' Com isso evita-se problemas de acentuação e/ou conversão de idiomas no banco
' de dados e dificulta-se ainda mais a interpretação dos dados.
' Obviamente, as rotinas são 'case sensitive', ou seja, letras maiúsculas e
' minúsculas geram resultados diferentes.
' 
' A senha digitada pelo usuário nunca poderá ultrapassar o MENOR dos seguintes
' limites:
'    a) 255 caracteres
'    b) ((TAMANHO_SENHA_FORMATADA / 2) - 2) caracteres
' 
' ========================================================================================================


'   FATOR (CRIPTOGRAFIA): ATÉ 9999
Private Const FATOR_CRIPTO = 1209

Private Const TAMANHO_SENHA_FORMATADA = 32  ' PROCURAR USAR SEMPRE POTÊNCIA DE 2
Private Const PREFIXO_SENHA_FORMATADA = "0x"
Private Const TAMANHO_CAMPO_COMPRIMENTO_SENHA = 2

Function converte_bin_para_dec(ByVal numero As String) As Byte
' ____________________________________________________________________________________________________
'|
'|  CONVERTE UM NÚMERO BINÁRIO PARA SUA FORMA DECIMAL.
'|

Dim i As Integer
Dim n_byte As Byte

    
    converte_bin_para_dec = 0
    
  ' TRANSFORMA BINÁRIO -> DECIMAL
    n_byte = 0
    For i = 1 To 8
        If Mid$(numero, 8 - (i - 1), 1) = "1" Then
            n_byte = n_byte + (2 ^ (i - 1))
            End If
        Next

    converte_bin_para_dec = n_byte
    

End Function


Function converte_dec_para_bin(ByVal numero As Byte) As String
' ____________________________________________________________________________________________________
'|
'|  CONVERTE UM NÚMERO DECIMAL PARA SUA FORMA BINÁRIA.
'|  O NÚMERO É PREENCHIDO C/ ZEROS À ESQUERDA, SE NECESSÁRIO.
'|

Dim i As Integer
Dim s As String
Dim s_byte As String

    
    converte_dec_para_bin = ""
    
  ' TRANSFORMA DECIMAL -> BINÁRIO ('0101...')
    s_byte = ""
    For i = 1 To 8
        If (numero And (2 ^ (i - 1))) Then s = "1" Else s = "0"
        s_byte = s & s_byte
        Next

    converte_dec_para_bin = s_byte
    
End Function

Sub rotaciona_direita(numero As Byte, ByVal casas As Byte)
' ____________________________________________________________________________________________________
'|
'|  ROTACIONA O BYTE PARA DIREITA 'CASAS' POSIÇÕES.
'|  IMPORTANTE: O ÚLTIMO BIT DA DIREITA SERÁ COLOCADO NA 1ª CASA DA ESQUERDA.
'|

Dim i As Integer
Dim s As String
Dim s_byte As String


  ' TRANSFORMA DECIMAL -> BINÁRIO ('0101...')
    s_byte = converte_dec_para_bin(numero)
        
  ' ROTACIONA
    For i = 1 To casas
        s = Right$(s_byte, 1)
        s_byte = Left$(s_byte, Len(s_byte) - 1)
        s_byte = s & s_byte
        Next
        
  ' TRANSFORMA BINÁRIO -> DECIMAL
    numero = converte_bin_para_dec(s_byte)
        
        
End Sub

Sub rotaciona_esquerda(numero As Byte, ByVal casas As Byte)
' ____________________________________________________________________________________________________
'|
'|  ROTACIONA O BYTE PARA ESQUERDA 'CASAS' POSIÇÕES.
'|  IMPORTANTE: O 1º BIT DA ESQUERDA SERÁ COLOCADO NA ÚLTIMA CASA DA DIREITA.
'|

Dim i As Integer
Dim s As String
Dim s_byte As String


  ' TRANSFORMA DECIMAL -> BINÁRIO ('0101...')
    s_byte = converte_dec_para_bin(numero)
        
  ' ROTACIONA
    For i = 1 To casas
        s = Left$(s_byte, 1)
        s_byte = Right$(s_byte, Len(s_byte) - 1)
        s_byte = s_byte & s
        Next
        
  ' TRANSFORMA BINÁRIO -> DECIMAL
    numero = converte_bin_para_dec(s_byte)
    
    
End Sub

Sub shift_direita(numero As Byte, ByVal casas As Byte)
' ____________________________________________________________________________________________________
'|
'|  DESLOCA O BYTE PARA DIREITA 'CASAS' POSIÇÕES.
'|  IMPORTANTE: AS CASAS DA ESQUERDA SERÃO PREENCHIDAS COM ZEROS.
'|

Dim i As Integer
Dim s As String
Dim s_byte As String


  ' TRANSFORMA DECIMAL -> BINÁRIO ('0101...')
    s_byte = converte_dec_para_bin(numero)
        
  ' ROTACIONA
    For i = 1 To casas
        s_byte = Left$(s_byte, Len(s_byte) - 1)
        s_byte = "0" & s_byte
        Next
        
  ' TRANSFORMA BINÁRIO -> DECIMAL
    numero = converte_bin_para_dec(s_byte)
        
        

End Sub

Sub shift_esquerda(numero As Byte, ByVal casas As Byte)
' ____________________________________________________________________________________________________
'|
'|  DESLOCA O BYTE PARA ESQUERDA 'CASAS' POSIÇÕES.
'|  IMPORTANTE: AS CASAS DA DIREITA SERÃO PREENCHIDAS COM ZEROS.
'|

Dim i As Integer
Dim s As String
Dim s_byte As String


  ' TRANSFORMA DECIMAL -> BINÁRIO ('0101...')
    s_byte = converte_dec_para_bin(numero)
        
  ' ROTACIONA
    For i = 1 To casas
        s_byte = Right$(s_byte, Len(s_byte) - 1)
        s_byte = s_byte & "0"
        Next
        
  ' TRANSFORMA BINÁRIO -> DECIMAL
    numero = converte_bin_para_dec(s_byte)
    
    

End Sub


Function codifica_dado(ByVal origem As String, destino As String, Optional ByVal inclui_preenchimento As Boolean = False) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'    CODIFICA O VALOR DADO POR 'ORIGEM', UTILIZANDO A CHAVE PRÉ-DEFINIDA.
' 
'   OBSERVAÇÕES:
'   ============
'   Esta função gera a senha criptografada, depois converte cada um dos caracteres
'   criptografados para seu respectivo código hexadecimal e adiciona o prefixo '0xNN',
'   sendo que 'NN' é um número hexadecimal indicando o tamanho da senha.
'   O tamanho da senha indica os 'NN' caracteres da direita que devem ser utilizados
'   para descriptografar a senha. Os caracteres restantes da esquerda são apenas para
'   preenchimento e devem ser ignorados.
'   Lembre-se de que a senha ocupa no BD, no mínimo (sem os caracteres de preenchimento):
'   2 x (tamanho descriptografado) + 2 bytes do '0x' + 2 bytes do 'NN'
'   Por exemplo:
'   'AbCdEf' -> '0x0c34330210feccf497b2907e4d61ac7ad0be04ac09a3cd679061bb9d7fd923'
'            => '0x' -> prefixo a ser descartado.
'            =>   '0c' -> os 12 caracteres da direita contém a senha criptografada.
'            =>     '34330210feccf497b2907e4d61ac7ad0be04ac09a3cd6790' -> caracteres preenchimento que devem ser descartados.
'            =>                                                     '61bb9d7fd923' -> caracteres da senha criptografada.
' 
'   A senha criptografada, portanto, é gerada em hexadecimal, com tamanho
'   formatado para que seu comprimento total fique sempre com TAMANHO_SENHA_FORMATADA
'   bytes (incluindo o '0xNN').
'   Deve-se lembrar que a senha (descriptografada) em si poderá ter no máximo:
'   ((TAMANHO_SENHA_FORMATADA / 2) - 2) caracteres
'---------------------------------------------------------------------------------------------------------------------------------------------
    
Dim i As Byte
Dim i_tam_senha As Integer
Dim i_chave As Byte
Dim i_dado As Byte
Dim k As Byte
Dim s_origem As String
Dim s_destino As String
Dim s As String
Dim chave As String

    
    On Error GoTo COD_DADO_ERRO
    
    
    codifica_dado = False
    
  ' DEFAULT
    destino = ""
    
    
  ' SENHA ORIGEM ESTÁ VAZIA !
    If Trim$(origem) = "" Then Exit Function
    
  ' SENHA EXCEDE TAMANHO
    If Len(Trim$(origem)) > ((TAMANHO_SENHA_FORMATADA - Len(PREFIXO_SENHA_FORMATADA) - TAMANHO_CAMPO_COMPRIMENTO_SENHA) \ 2) Then Exit Function
    
  ' GERA CHAVE DE CRIPTOGRAFIA
    If Not gera_chave_codificacao(FATOR_CRIPTO, chave) Then Exit Function
    
    
    s_destino = ""
    s_origem = Trim$(origem)
     
  ' CRITOGRAFA PELA CHAVE
    For i = 1 To Len(s_origem)
        i_chave = Asc(Mid$(chave, i, 1))
        shift_esquerda i_chave, 1
        i_chave = i_chave + 1
        
        i_dado = Asc(Mid$(s_origem, i, 1))
        rotaciona_esquerda i_dado, 1
        
        k = i_chave Xor i_dado
        s_destino = s_destino & Chr$(k)
        Next i

  ' ASCII --> HEXADECIMAL
    s_origem = s_destino
    s_destino = ""
    For i = 1 To Len(s_origem)
        k = Asc(Mid$(s_origem, i, 1))
        s = Hex(k)
        While Len(s) < 2: s = "0" & s: Wend
        s_destino = s_destino & s
        Next i
        
        
  ' GUARDA O TAMANHO REAL DA SENHA
    i_tam_senha = Len(s_destino)
        
    If inclui_preenchimento Then
      ' COLOCA MÁSCARA (IMITA FORMATO TIMESTAMP)
        i = 0
        While Len(s_destino) < (TAMANHO_SENHA_FORMATADA - Len(PREFIXO_SENHA_FORMATADA) - TAMANHO_CAMPO_COMPRIMENTO_SENHA)
          ' AO INVÉS DE PREENCHER C/ ZEROS, GERA CÓDIGO P/ PREENCHIMENTO
            i = i + 1
            s = Hex(i Xor CInt("&H" & Mid$(s_destino, Len(s_destino) - (i - 1), 1)) Xor CInt("&H" & Mid$(s_destino, Len(s_destino) - i, 1)))
              
          ' ADICIONA UM CARACTER POR VEZ, P/ NÃO CORRER O RISCO DE ULTRAPASSAR TAMANHO MÁXIMO
            s_destino = Right$(s, 1) & s_destino
            Wend
    
      ' ADICIONA PREFIXO E TAMANHO REAL DA SENHA
        s = Hex(i_tam_senha)
        While Len(s) < 2: s = "0" & s: Wend
    Else
        While Len(s_destino) < (TAMANHO_SENHA_FORMATADA - Len(PREFIXO_SENHA_FORMATADA) - TAMANHO_CAMPO_COMPRIMENTO_SENHA)
            s_destino = "0" & s_destino
            Wend
            
        s = "00"
        End If
        
    s_destino = PREFIXO_SENHA_FORMATADA & s & s_destino
        
    destino = LCase(s_destino)

    codifica_dado = True
    
Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
COD_DADO_ERRO:
'=============
    MsgBox CStr(Err) & ": " & Error$(Err), vbOKOnly + vbExclamation + vbSystemModal, "PARA SUA INFORMAÇÃO"
    
    Exit Function
    

End Function




Function decodifica_dado(ByVal origem As String, destino As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'    DECODIFICA O VALOR DADO POR 'ORIGEM', UTILIZANDO A CHAVE PRÉ-DEFINIDA.
' 
'   OBSERVAÇÕES:
'   ============
'   Esta função descriptografa a senha, convertendo os códigos hexadecimais
'   de volta para os caracteres ASCII criptografados e, depois, descriptografando
'   a senha.
'   As 4 primeiras posições ('0xNN') formam o prefixo da senha, sendo que '0x'
'   deve ser descartado e 'NN' indica o tamanho da senha.
'   O tamanho da senha indica os 'NN' caracteres da direita que devem ser utilizados
'   para descriptografar a senha. Os caracteres restantes da esquerda são apenas para
'   preenchimento e devem ser ignorados.
'   Lembre-se de que a senha ocupa no BD, no mínimo (sem os caracteres de preenchimento):
'   2 x (tamanho descriptografado) + 2 bytes do '0x' + 2 bytes do 'NN'
'   Por exemplo:
'   'AbCdEf' -> '0x0c34330210feccf497b2907e4d61ac7ad0be04ac09a3cd679061bb9d7fd923'
'            => '0x' -> prefixo a ser descartado.
'            =>   '0c' -> os 12 caracteres da direita contém a senha criptografada.
'            =>     '34330210feccf497b2907e4d61ac7ad0be04ac09a3cd6790' -> caracteres preenchimento que devem ser descartados.
'            =>                                                     '61bb9d7fd923' -> caracteres da senha criptografada.
' 
'   A senha criptografada é gerada com formatação para que seu tamanho
'   total fique sempre com TAMANHO_SENHA_FORMATADA bytes (incluindo o '0xNN').
'   Deve-se lembrar que a senha (descriptografada) em si poderá ter no máximo:
'   ((TAMANHO_SENHA_FORMATADA / 2) - 2) caracteres
'---------------------------------------------------------------------------------------------------------------------------------------------

Dim i As Byte
Dim i_chave As Byte
Dim i_dado As Byte
Dim k As Byte
Dim s_origem As String
Dim s_destino As String
Dim s As String
Dim chave As String
     
     
    On Error GoTo DECOD_DADO_ERRO
    
    
    decodifica_dado = False
    
  ' DEFAULT
    destino = ""
    
    
  ' GERA CHAVE DE CRIPTOGRAFIA
    If Not gera_chave_codificacao(FATOR_CRIPTO, chave) Then Exit Function
    
    
    s_destino = ""
    s_origem = Trim$(origem)

  ' POSSUI PREFIXO '0x' ?
    If Left$(s_origem, Len(PREFIXO_SENHA_FORMATADA)) <> PREFIXO_SENHA_FORMATADA Then Exit Function
    
  ' RETIRA PREFIXO '0x' DA MÁSCARA (IMITA FORMATO TIMESTAMP)
    s_origem = Right$(s_origem, Len(s_origem) - Len(PREFIXO_SENHA_FORMATADA))
    s_origem = UCase(s_origem)
  
  ' RETIRA CARACTERES DE PREENCHIMENTO (IMITA FORMATO TIMESTAMP)
    s = Left$(s_origem, TAMANHO_CAMPO_COMPRIMENTO_SENHA)
    s = "&H" & s
    If IsNumeric(s) Then i = CInt(s) Else i = 0
    If i <> 0 Then
        s_origem = Right$(s_origem, i)
    Else
        Do While Mid$(s_origem, 1, 2) = "00"
            s_origem = Right$(s_origem, Len(s_origem) - 2)
            Loop
        End If
        
  ' HEXADECIMAL --> ASCII
    For i = 1 To Len(s_origem) Step 2
        s = Mid$(s_origem, i, 2)
        s = "&H" & s
        s_destino = s_destino & Chr$(s)
        Next i
    
    
  ' DECRIPTOGRAFA PELA CHAVE
    s_origem = s_destino
    s_destino = ""
    For i = 1 To Len(s_origem)
        i_chave = Asc(Mid$(chave, i, 1))
        shift_esquerda i_chave, 1
        i_chave = i_chave + 1
        
        i_dado = Asc(Mid$(s_origem, i, 1))
        k = i_chave Xor i_dado
        
        rotaciona_direita k, 1
        s_destino = s_destino & Chr$(k)
        Next i

    
    destino = s_destino

    decodifica_dado = True


Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DECOD_DADO_ERRO:
'===============
    MsgBox CStr(Err) & Error$(Err), vbOKOnly + vbExclamation + vbSystemModal, "PARA SUA INFORMAÇÃO"
    
    Exit Function
    
    
End Function




Function gera_chave_codificacao(ByVal fator As Long, chave_gerada As String) As Boolean

'-------------------------------------------------------------------------------
'    Gera a chave para criptografia.
'-------------------------------------------------------------------------------
    
Const COD_MINIMO = 35
Const COD_MAXIMO = 96
Const TAMANHO_CHAVE = 128

Dim i As Integer
Dim k As Currency
Dim s As String

    
    On Error GoTo GCCOD_ERRO
    
    
    gera_chave_codificacao = False
    
  ' DEFAULT
    chave_gerada = ""
    
    
    s = ""
    For i = 1 To TAMANHO_CHAVE
        k = (COD_MAXIMO - COD_MINIMO) + 1
        k = (k * fator)
        k = (k * i) + COD_MINIMO
        k = k Mod 128
        s = s & Chr$(k)
        Next i

    
    chave_gerada = s
    
    gera_chave_codificacao = True

Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GCCOD_ERRO:
'==========
    MsgBox CStr(Err) & Error$(Err), vbOKOnly + vbExclamation + vbSystemModal, "PARA SUA INFORMAÇÃO"
    
    Exit Function
    
    
End Function


