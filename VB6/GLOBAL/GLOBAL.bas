Attribute VB_Name = "mod_GLOBAL"
Option Explicit

  ' CONTROLA SE HÁ ROTINAS EM EXECUÇÃO
    Global em_execucao As Boolean
    
  ' VARIÁVEIS PARA CONTROLE DE HORÁRIO DE VERÃO (NECESSÁRIAS PARA O LAYOUT NFE 3.10)
    Global blnHorarioVerao As Boolean
    Global intHorarioVeraoAtivo As Integer
    Global intHorarioVeraoAtivoInicio As Integer
    Global sHorarioVeraoData As String
    Global blnInfoAdicParc As Boolean
    Global intInfoAdicParc As Integer
    Global intInfoAdicParcInicio As Integer

  ' VARIÁVEIS UTILIZADAS PARA EMISSÃO DE NOTAS REFERENTES A ASSISTÊNCIA TÉCNICA
    Global s_assist_Pedido As String
    Global s_assist_Resposta As String
    Global s_assist_EndEntregaFormatado As String
    Global s_assist_EndEntregaUf As String
    Global s_assist_EndClienteUf As String
    Global s_assist_Cliente_CNPJ_CPF As String
    Global bln_assist_pedido_ok As Boolean

  ' TOLERÂNCIA MÁXIMA ACEITA DE ERRO NO RELÓGIO DA ESTAÇÃO C/ RELAÇÃO AO HORÁRIO NO SERVIDOR DE BD.
  ' SUA UTILIZAÇÃO BASICAMENT É FEITA NOS CASOS EM QUE O USUÁRIO ESTÁ IMPEDIDO DE ATUALIZAR O RELÓGIO DA MÁQUINA.
    Global Const MAX_ERRO_RELOGIO_EM_MINUTOS = 30


  ' PAINEL INFO (MENSAGENS DE STATUS)
    Public Enum INFO_STATUS_OPTIONS
        INFO_NORMAL = 1
        INFO_EXECUTANDO = 2
        End Enum
        
  ' CHAVE DO REGISTRO PARA CONFIGURAÇÕES DO USUÁRIO
    Global Const REG_CHAVE_USUARIO_USUARIO = "SOFTWARE\PRAGMATICA\Sistema Contratos\AMB\Usuario"

  ' STATUS RELACIONADOS A NOTAS FISCAIS TRIANGULARES
    Global Const ST_NFT_NOP = -1
    Global Const ST_NFT_NAO_EMITIDA = 0
    Global Const ST_NFT_EM_PROCESSAMENTO = 1
    Global Const ST_NFT_EMITIDA = 2
    Global Const ST_NFT_CANCELADA_USUARIO = 3
    Global Const ST_NFT_CANCELADA_TIMEOUT = 4
    Global Const ST_NFT_CANCELADA_SISTEMA = 5


    Global painel_ativo As Form
    Global painel_principal As Form

    Type TIPO_CINCO_COLUNAS
        c1 As String
        c2 As String
        c3 As String
        c4 As String
        c5 As String
        End Type

  ' U S U Á R I O
    Type TIPO_USUARIO
        id As String
        perfil_acesso_ok As Boolean
        loja As String
        emit As String
        emit_uf As String
        emit_id As String
        senha As String
        nome As String
        bloqueado As Long
        dt_cadastro As Date
        dt_ult_atualizacao As Date
        dt_ult_alteracao_senha As Date
        End Type

    Global usuario As TIPO_USUARIO

    Global vEmitsUsuario() As TIPO_CINCO_COLUNAS
    
    Global txtFixoEspecifico As String

    'V A R I Á V E I S   G E R A I S
    Global cor_fundo_padrao As String
    
    Global identificador_ambiente_padrao As String
    
    Global sPedidoTriangular As String
    
    Global sPedidoDANFETelaAnterior As String
    
    Global sNFAnteriorSerie As String
    
    Global sNFAnteriorNumero As String
    
    Global sNFAnteriorEmitente As String
    
    Global sAvisosAExibir As String
    
    Global blnNotaTriangularAtiva As Boolean
    
    'P A R Â M E T R O
    '(registro da tabela t_PARAMETRO)
    Type TIPO_t_PARAMETRO
        id As String
        campo_inteiro As Integer
        campo_monetario As Double
        campo_real As Double
        campo_data As Date
        campo_texto As String
        dt_hr_ult_atualizacao As Date
        usuario_ult_atualizacao As String
        End Type
        
    Global param_geracaoboletos As TIPO_t_PARAMETRO
    
    Global param_notatriangular As TIPO_t_PARAMETRO
    
    Global param_atualizanfnopedido As TIPO_t_PARAMETRO
    
    Global param_pedidomemorizacaoenderecos As TIPO_t_PARAMETRO

    Global Const FORMATO_MOEDA = "###,###,###,##0.00"
    Global Const FORMATO_VALOR = "###,###,###,##0.00"
    Global Const FORMATO_DATA = "dd/mm/yyyy"
    Global Const FORMATO_DATA_HORA = "dd/mm/yyyy hh:mm:ss"
    Global Const FORMATO_NUMERO = "###,###,###,##0"
    Global Const FORMATO_INTEIRO = "###,###,###,##0"

    Global Const SIMBOLO_MONETARIO = "R$"





'-------------------------------------------------------------------------------
'   CORES - NOMES PADRONIZADOS
'-------------------------------------------------------------------------------
    Global Const COR_PRETO = &H0&
    
    Global Const COR_VERMELHO = &HFF&
    Global Const COR_VERMELHO_ESCURO = &H80&
    
    Global Const COR_VERDE = &HFF00&
    Global Const COR_VERDE_ESCURO = &H8000&
    
    Global Const COR_AMARELO = &HFFFF&
    Global Const COR_AMARELO_ESCURO = &H8080&
    
    Global Const COR_AZUL = &HFF0000
    Global Const COR_AZUL_ESCURO = &H800000
    
    Global Const COR_MAGENTA = &HFF00FF
    Global Const COR_MAGENTA_ESCURO = &H800080
    
    Global Const COR_CIAN = &HFFFF00
    Global Const COR_CIAN_ESCURO = &H808000
    
    Global Const COR_CINZA = &HC0C0C0
    Global Const COR_CINZA_ESCURO = &H808080
    Global Const COR_CINZA_CLARO = &HE0E0E0
    
    Global Const COR_BRANCO = &HFFFFFF


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Function filtra_data(tecla As Integer) As Integer
' ________________________________________________
'|
'|  PERMITE SOMENTE CÓDIGOS NUMÉRICOS  E A BARRA
'|

    filtra_data = tecla
 
  ' PERMITE A PASSAGEM DOS CÓDIGOS DE CONTROLE
    If tecla < Asc(" ") Then Exit Function
    
  ' FILTRA CÓDIGOS DIFERENTES DE ['0'..'9']
    If ((tecla < Asc("0")) Or (tecla > Asc("9"))) And (tecla <> Asc("/")) Then
        filtra_data = 0
        Beep
        End If

End Function

Sub data_ok(d As Control)

Dim s As String
Dim s_ano As String


    If Trim$(d) = "" Then Exit Sub


    s = d
    
    If InStr(s, "/") Then
        If Len(s) = 8 Then
            If IsNumeric(right$(s, 2)) Then
                If CInt(right$(s, 2)) > 50 Then
                    s = Format$(d, "dd/mm/19yy")
                Else
                    s = Format$(d, "dd/mm/20yy")
                    End If
                End If

        Else
            s = Format$(d, "dd/mm/yyyy")
            End If


    Else
        If Len(s) = 6 Then
            If IsNumeric(right$(s, 2)) Then
                If CInt(right$(s, 2)) > 50 Then
                    s = Format$(d, "@@/@@/19@@")
                Else
                    s = Format$(d, "@@/@@/20@@")
                    End If
                End If

        Else
            s = Format$(d, "@@/@@/@@@@")
            End If
        End If



  ' CORRIGE ANOS MENORES QUE 1900
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    s_ano = retorna_so_digitos(right$(s, 4))
    If (Len(s_ano) = 4) And (s_ano < "1900") Then
        If CInt(right$(s_ano, 2)) > 50 Then
            s_ano = "19" & right$(s_ano, 2)
        Else
            s_ano = "20" & right$(s_ano, 2)
            End If
        
        s = Format$(s, "dd") & "/" & Format$(s, "mm") & "/" & s_ano
        End If
    


  ' FINALIZAÇÃO
  ' ~~~~~~~~~~~
    If IsDate(s) Then
        d = Format$(s, "dd/mm/yyyy")
    Else
      ' DATA INVÁLIDA
        aviso_erro "Data especificada não é uma data válida!!"
        d.SelStart = 0
        d.SelLength = Len(d)
        CAMPO_setfocus d
        Exit Sub
        End If
        

End Sub

Sub CAMPO_setfocus(c As Control)
' _______________________________________________________________________________________________
'|
'|  PREVINE O RUN-TIME ERROR '5' QUE OCORRE APENAS NO EXECUTÁVEL (NÃO OCORRE DENTRO DO VB).
'|  QUANDO SE NAVEGA ENTRE PAINÉIS, ALGUNS CONTROLES FICAM AUTOMATICAMENTE COM A PROPRIEDADE
'|  'ENABLED' = FALSE, O QUE GERA ERRO AO TENTAR ATRIBUIR O FOCO.
'|
    
    On Error Resume Next
    
    
    If c.Visible And c.Enabled Then c.SetFocus
    
    
End Sub

Function UF_ok(ByVal uf As String) As Boolean
' _________________________________________________________
'|
'|  VERIFICAÇÃO DA SIGLA DA UNIDADE FEDERATIVA
'|

Const sigla = "AC AL AM AP BA CE DF ES GO MA MG MS MT PA PB PE PI PR RJ RN RO RR RS SC SE SP TO  "
    
    UF_ok = False
    
    uf = Trim$(uf)

    If uf = "" Then
        UF_ok = True
        Exit Function
        End If

    If Len(uf) <> 2 Then
        UF_ok = False
        Exit Function
        End If

    UF_ok = (InStr(sigla, uf) <> 0)

End Function

Function formata_perc(ByVal valor As Single) As String
Const FORMATO_PERCENTUAL_2DEC = "##0.00"
Dim strSeparadorDecimal As String
Dim strValorFormatado As String
Dim i As Integer
Dim c As String
Dim s As String

    strSeparadorDecimal = ""
    s = Format$(0.5, FORMATO_PERCENTUAL_2DEC)
    For i = Len(s) To 1 Step -1
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then
            strSeparadorDecimal = c
            Exit For
            End If
        Next

    If strSeparadorDecimal = "" Then strSeparadorDecimal = ","
    
    strValorFormatado = Format$(valor, FORMATO_PERCENTUAL_2DEC)
    strValorFormatado = substitui_caracteres(strValorFormatado, strSeparadorDecimal, "V")
    strValorFormatado = substitui_caracteres(strValorFormatado, ".", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, ",", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, "V", ",")
    
    formata_perc = strValorFormatado
End Function

Function formata_numero_2dec(ByVal valor As Double) As String
Const FORMATO_NUMERO_2DEC = "###,###,###,##0.00"
Dim strSeparadorDecimal As String
Dim strValorFormatado As String
Dim i As Integer
Dim c As String
Dim s As String

    strSeparadorDecimal = ""
    s = Format$(0.5, FORMATO_NUMERO_2DEC)
    For i = Len(s) To 1 Step -1
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then
            strSeparadorDecimal = c
            Exit For
            End If
        Next

    If strSeparadorDecimal = "" Then strSeparadorDecimal = ","
    
    strValorFormatado = Format$(valor, FORMATO_NUMERO_2DEC)
    strValorFormatado = substitui_caracteres(strValorFormatado, strSeparadorDecimal, "V")
    strValorFormatado = substitui_caracteres(strValorFormatado, ".", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, ",", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, "V", ",")
    
    formata_numero_2dec = strValorFormatado
End Function

Function formata_numero_3dec(ByVal valor As Double) As String
Const FORMATO_NUMERO_3DEC = "###,###,###,##0.000"
Dim strSeparadorDecimal As String
Dim strValorFormatado As String
Dim i As Integer
Dim c As String
Dim s As String

    strSeparadorDecimal = ""
    s = Format$(0.5, FORMATO_NUMERO_3DEC)
    For i = Len(s) To 1 Step -1
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then
            strSeparadorDecimal = c
            Exit For
            End If
        Next

    If strSeparadorDecimal = "" Then strSeparadorDecimal = ","
    
    strValorFormatado = Format$(valor, FORMATO_NUMERO_3DEC)
    strValorFormatado = substitui_caracteres(strValorFormatado, strSeparadorDecimal, "V")
    strValorFormatado = substitui_caracteres(strValorFormatado, ".", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, ",", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, "V", ",")
    
    formata_numero_3dec = strValorFormatado
End Function


Function formata_numero_1dec(ByVal valor As Double) As String
Const FORMATO_NUMERO_1DEC = "###,###,###,##0.0"
Dim strSeparadorDecimal As String
Dim strValorFormatado As String
Dim i As Integer
Dim c As String
Dim s As String

    strSeparadorDecimal = ""
    s = Format$(0.5, FORMATO_NUMERO_1DEC)
    For i = Len(s) To 1 Step -1
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then
            strSeparadorDecimal = c
            Exit For
            End If
        Next

    If strSeparadorDecimal = "" Then strSeparadorDecimal = ","
    
    strValorFormatado = Format$(valor, FORMATO_NUMERO_1DEC)
    strValorFormatado = substitui_caracteres(strValorFormatado, strSeparadorDecimal, "V")
    strValorFormatado = substitui_caracteres(strValorFormatado, ".", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, ",", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, "V", ",")
    
    formata_numero_1dec = strValorFormatado
End Function

Function filtra_cnpj_cpf(ByVal KeyAscii As Integer)
Dim c As String
    
    filtra_cnpj_cpf = 0
    c = Chr$(KeyAscii)
    If ((c < "0") Or (c > "9")) And (c <> ".") And (c <> "/") And (c <> "-") Then Exit Function
    filtra_cnpj_cpf = KeyAscii
    
End Function

Function filtra_moeda(ByVal KeyAscii As Integer)
Dim c As String
    
    filtra_moeda = 0
    
    c = Chr(KeyAscii)
    If ((c < "0") Or (c > "9")) And (c <> ".") And (c <> ",") Then Exit Function
    
    filtra_moeda = KeyAscii
    
End Function



Function filtra_peso(ByVal KeyAscii As Integer)
Dim c As String
    
    filtra_peso = 0
    
    c = Chr(KeyAscii)
    If ((c < "0") Or (c > "9")) And (c <> ".") And (c <> ",") Then Exit Function
    
    filtra_peso = KeyAscii
    
End Function

Public Function filtra_perc(ByRef c As Control, ByRef tecla As Integer) As Integer
' _________________________________
'|                                 |
'|  VALOR NUMÉRICO TIPO PERCENTUAL |
'|_________________________________|
           

    filtra_perc = tecla

  ' PERMITE A PASSAGEM DOS CÓDIGOS DE CONTROLE
    If tecla < Asc(" ") Then Exit Function
    
  ' PERMITE APENAS UMA VÍRGULA
    If tecla = Asc(",") Then
        If InStr(c, ",") <> 0 Then
            filtra_perc = 0
            Beep
            End If
    
  ' FILTRA CÓDIGOS DIFERENTES DE ['0'..'9',","]
    ElseIf (tecla < Asc("0") Or tecla > Asc("9")) Then
        filtra_perc = 0
        Beep
        End If
    
End Function


Function cnpj_cpf_ok(ByVal numero As String) As Boolean

Dim i As Integer
Dim tudo_igual As Boolean

    cnpj_cpf_ok = False
    
    numero = retorna_so_digitos(numero)
    
    If Len(numero) = 11 Then
        If Not cpf_ok(numero) Then Exit Function
    ElseIf Len(numero) = 14 Then
        If Not cnpj_ok(numero) Then Exit Function
    Else
        Exit Function
        End If
        
    tudo_igual = True
    For i = 1 To (Len(numero) - 1)
        If Mid$(numero, i, 1) <> Mid$(numero, i + 1, 1) Then
            tudo_igual = False
            Exit For
            End If
        Next
        
    If tudo_igual Then Exit Function
    
    cnpj_cpf_ok = True
    
End Function


Function cpf_ok(ByVal cpf As String) As Integer

Dim d As Integer
Dim i As Integer

'   _________________________________
'   VERIFICA OS 'CHECK DIGITS' DO CPF
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    If Trim$(cpf) = "" Then Exit Function

    If Len(cpf) <> 11 Then Exit Function


' 
'   VERIFICA O PRIMEIRO CHECK DIGIT
' 
    d = 0
    For i = 1 To 9
        d = d + (11 - i) * Val(Mid$(cpf, i, 1))
        Next i

    d = 11 - (d Mod 11)
    If d > 9 Then d = 0
    
    If d <> Val(Mid$(cpf, 10, 1)) Then
        cpf_ok = False
        Exit Function
        End If

' 
'   VERIFICA O SEGUNDO CHECK DIGIT
' 
    d = 0
    For i = 2 To 10
        d = d + (12 - i) * Val(Mid$(cpf, i, 1))
        Next i
    
    d = 11 - (d Mod 11)
    If d > 9 Then d = 0
    
    If d <> Val(Mid$(cpf, 11, 1)) Then
        cpf_ok = False
        Exit Function
        End If

    
    
    cpf_ok = True


End Function



Function cnpj_ok(ByVal cgc As String) As Integer

Dim d As Integer
Dim i As Integer


Const p1 = "543298765432"
Const p2 = "6543298765432"

    If Trim$(cgc) = "" Then Exit Function

    If Len(cgc) <> 14 Then Exit Function


' 
'  VERIFICA O PRIMEIRO CHECK DIGIT
' 
    d = 0
    For i = 1 To 12
        d = d + Val(Mid$(p1, i, 1)) * Val(Mid$(cgc, i, 1))
        Next i

    d = 11 - (d Mod 11)
    If d > 9 Then d = 0
    
    If d <> Val(Mid$(cgc, 13, 1)) Then
        cnpj_ok = False
        Exit Function
        End If

' 
'   VERIFICA O SEGUNDO CHECK DIGIT
' 
    d = 0
    For i = 1 To 13
        d = d + Val(Mid$(p2, i, 1)) * Val(Mid$(cgc, i, 1))
        Next i

    d = 11 - (d Mod 11)
    If d > 9 Then d = 0
    
    If d <> Val(Mid$(cgc, 14, 1)) Then
        cnpj_ok = False
        Exit Function
        End If
    
    cnpj_ok = True
     
End Function


Function converte_para_currency(ByVal numero As String) As Currency
Dim i As Integer
Dim s As String
Dim c As String
Dim s_sep As String
Dim s_sep_sys As String
Dim s_valor As String

    converte_para_currency = 0
    numero = Trim$("" & numero)
    If numero = "" Then Exit Function
    s_sep = retorna_separador_decimal(numero)
    s_sep_sys = retorna_separador_decimal_sistema()
    If s_sep_sys = "" Then Exit Function
    numero = substitui_caracteres(numero, s_sep, "V")
    
    s = ""
    For i = 1 To Len(numero)
        c = Mid$(numero, i, 1)
        If (Not IsNumeric(c)) And (c <> "-") And (c <> "V") Then c = ""
        s = s & c
        Next
    
    s_valor = substitui_caracteres(s, "V", s_sep_sys)
    If Not IsNumeric(s_valor) Then Exit Function

    converte_para_currency = CCur(s_valor)
End Function

Function converte_para_single(ByVal numero As String) As Single
Dim i As Integer
Dim s As String
Dim c As String
Dim s_sep As String
Dim s_sep_sys As String
Dim s_valor As String

    converte_para_single = 0
    numero = Trim$("" & numero)
    If numero = "" Then Exit Function
    s_sep = retorna_separador_decimal(numero)
    s_sep_sys = retorna_separador_decimal_sistema()
    If s_sep_sys = "" Then Exit Function
    numero = substitui_caracteres(numero, s_sep, "V")
    
    s = ""
    For i = 1 To Len(numero)
        c = Mid$(numero, i, 1)
        If (Not IsNumeric(c)) And (c <> "-") And (c <> "V") Then c = ""
        s = s & c
        Next
    
    s_valor = substitui_caracteres(s, "V", s_sep_sys)
    If Not IsNumeric(s_valor) Then Exit Function

    converte_para_single = CSng(s_valor)
End Function

Function converte_para_double(ByVal numero As String) As Double
Dim i As Integer
Dim s As String
Dim c As String
Dim s_sep As String
Dim s_sep_sys As String
Dim s_valor As String

    converte_para_double = 0
    numero = Trim$("" & numero)
    If numero = "" Then Exit Function
    s_sep = retorna_separador_decimal(numero)
    s_sep_sys = retorna_separador_decimal_sistema()
    If s_sep_sys = "" Then Exit Function
    numero = substitui_caracteres(numero, s_sep, "V")
    
    s = ""
    For i = 1 To Len(numero)
        c = Mid$(numero, i, 1)
        If (Not IsNumeric(c)) And (c <> "-") And (c <> "V") Then c = ""
        s = s & c
        Next
    
    s_valor = substitui_caracteres(s, "V", s_sep_sys)
    If Not IsNumeric(s_valor) Then Exit Function

    converte_para_double = CDbl(s_valor)
End Function


Public Function arredonda_para_monetario(ByVal numero As Double) As Currency
'   ARREDONDA PARA APENAS 2 CASAS DECIMAIS
    arredonda_para_monetario = converte_para_currency(Format$(numero, FORMATO_MOEDA))
End Function

Function formata_endereco(ByVal endereco As String, ByVal endereco_numero As String, ByVal endereco_complemento As String, ByVal bairro As String, ByVal cidade As String, ByVal uf As String, ByVal cep As String) As String
Dim s_aux As String
Dim strResposta As String
    
    strResposta = ""
    
    If Trim(endereco) <> "" Then
        strResposta = Trim(endereco)
        s_aux = Trim(endereco_numero)
        If s_aux <> "" Then strResposta = strResposta & ", " & s_aux
        s_aux = Trim(endereco_complemento)
        If s_aux <> "" Then strResposta = strResposta & " " & s_aux
        s_aux = Trim(bairro)
        If s_aux <> "" Then strResposta = strResposta & " - " & s_aux
        s_aux = Trim(cidade)
        If s_aux <> "" Then strResposta = strResposta & " - " & s_aux
        s_aux = Trim(uf)
        If s_aux <> "" Then strResposta = strResposta & " - " & s_aux
        s_aux = Trim(cep)
        If s_aux <> "" Then strResposta = strResposta & " - " & cep_formata(s_aux)
        End If
        
    formata_endereco = strResposta
    
End Function


Public Function formata_moeda(ByVal valor As Currency) As String
    formata_moeda = Format$(valor, FORMATO_MOEDA)
End Function

Function NFeFormataCampo(ByVal nomeCampo As String, ByVal valorCampo As String) As String
    NFeFormataCampo = nomeCampo & "=" & substitui_caracteres(valorCampo, ";", ",") & ";" & vbCrLf
End Function

Function NFeFormataData(ByVal data As Date) As String
    NFeFormataData = Format$(data, "yyyy-mm-dd")
End Function

Function NFeFormataDataHoraUTC(ByVal data As Date, ByVal HVerao As Boolean) As String
    If HVerao Then
        NFeFormataDataHoraUTC = Format$(data, "yyyy-mm-dd") + "T" + Format$(data, "hh:nn:ss") + "-02:00"
    Else
        NFeFormataDataHoraUTC = Format$(data, "yyyy-mm-dd") + "T" + Format$(data, "hh:nn:ss") + "-03:00"
        End If
End Function

Function NFeFormataMoeda2Dec(ByVal valor As Currency) As String
Const FORMATO_MOEDA_NFE_2DEC = "###########0.00"
Dim strSeparadorDecimal As String
Dim strValorFormatado As String
Dim i As Integer
Dim c As String
Dim s As String

    strSeparadorDecimal = ""
    s = Format$(0.5, FORMATO_MOEDA_NFE_2DEC)
    For i = Len(s) To 1 Step -1
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then
            strSeparadorDecimal = c
            Exit For
            End If
        Next

    If strSeparadorDecimal = "" Then strSeparadorDecimal = ","
    
    strValorFormatado = Format$(valor, FORMATO_MOEDA_NFE_2DEC)
    strValorFormatado = substitui_caracteres(strValorFormatado, strSeparadorDecimal, "V")
    strValorFormatado = substitui_caracteres(strValorFormatado, ".", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, ",", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, "V", ".")
    
    NFeFormataMoeda2Dec = strValorFormatado
End Function


Public Function NFeFormataSerieNF(ByVal numeroSerieNF As Variant) As String

Dim s_resp As String
    
    s_resp = Trim$(CStr(numeroSerieNF))
    Do While Len(s_resp) < 3: s_resp = "0" & s_resp: Loop
    NFeFormataSerieNF = s_resp
    
End Function

Public Function NFeFormataNumeroNF(ByVal numeroNF As Variant) As String

Dim s_resp As String
    
    s_resp = Trim$(CStr(numeroNF))
    Do While Len(s_resp) < 9: s_resp = "0" & s_resp: Loop
    NFeFormataNumeroNF = s_resp
    
End Function


Function NFeFormataPercentual2Dec(ByVal valor As Single) As String
Const FORMATO_PERCENTUAL_NFE_2DEC = "##0.00"
Dim strSeparadorDecimal As String
Dim strValorFormatado As String
Dim i As Integer
Dim c As String
Dim s As String

    strSeparadorDecimal = ""
    s = Format$(0.5, FORMATO_PERCENTUAL_NFE_2DEC)
    For i = Len(s) To 1 Step -1
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then
            strSeparadorDecimal = c
            Exit For
            End If
        Next

    If strSeparadorDecimal = "" Then strSeparadorDecimal = ","
    
    strValorFormatado = Format$(valor, FORMATO_PERCENTUAL_NFE_2DEC)
    strValorFormatado = substitui_caracteres(strValorFormatado, strSeparadorDecimal, "V")
    strValorFormatado = substitui_caracteres(strValorFormatado, ".", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, ",", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, "V", ".")
    
    NFeFormataPercentual2Dec = strValorFormatado
End Function



Function NFeFormataNumero4Dec(ByVal numero As Variant) As String
Const FORMATO_NUMERO_NFE_4DEC = "###########0.0000"
Dim strSeparadorDecimal As String
Dim strValorFormatado As String
Dim i As Integer
Dim c As String
Dim s As String

    strSeparadorDecimal = ""
    s = Format$(0.5, FORMATO_NUMERO_NFE_4DEC)
    For i = Len(s) To 1 Step -1
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then
            strSeparadorDecimal = c
            Exit For
            End If
        Next

    If strSeparadorDecimal = "" Then strSeparadorDecimal = ","
    
    strValorFormatado = Format$(numero, FORMATO_NUMERO_NFE_4DEC)
    strValorFormatado = substitui_caracteres(strValorFormatado, strSeparadorDecimal, "V")
    strValorFormatado = substitui_caracteres(strValorFormatado, ".", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, ",", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, "V", ".")
    
    NFeFormataNumero4Dec = strValorFormatado
End Function


Function NFeFormataNumero3Dec(ByVal numero As Variant) As String
Const FORMATO_NUMERO_NFE_3DEC = "###########0.000"
Dim strSeparadorDecimal As String
Dim strValorFormatado As String
Dim i As Integer
Dim c As String
Dim s As String

    strSeparadorDecimal = ""
    s = Format$(0.5, FORMATO_NUMERO_NFE_3DEC)
    For i = Len(s) To 1 Step -1
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then
            strSeparadorDecimal = c
            Exit For
            End If
        Next

    If strSeparadorDecimal = "" Then strSeparadorDecimal = ","
    
    strValorFormatado = Format$(numero, FORMATO_NUMERO_NFE_3DEC)
    strValorFormatado = substitui_caracteres(strValorFormatado, strSeparadorDecimal, "V")
    strValorFormatado = substitui_caracteres(strValorFormatado, ".", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, ",", "")
    strValorFormatado = substitui_caracteres(strValorFormatado, "V", ".")
    
    NFeFormataNumero3Dec = strValorFormatado
End Function

Public Function retorna_separador_decimal(ByVal numero As String) As String
Dim i As Integer
Dim c As String
Dim s_num As String
Dim s_resp As String
Dim n_ponto As Integer
Dim n_virg As Integer
Dim s_ult_sep As String
Dim n_digitos_finais As Integer

    n_digitos_finais = 0
    n_ponto = 0
    n_virg = 0
    s_ult_sep = ""
    s_num = Trim$("" & numero)
    For i = Len(s_num) To 1 Step -1
        c = Mid$(s_num, i, 1)
        If (c = ".") Then
            n_ponto = n_ponto + 1
            If (s_ult_sep = "") Then s_ult_sep = c
        ElseIf (c = ",") Then
            n_virg = n_virg + 1
            If (s_ult_sep = "") Then s_ult_sep = c
            End If
        If IsNumeric(c) And (n_ponto = 0) And (n_virg = 0) Then n_digitos_finais = n_digitos_finais + 1
        Next
        
'   DEFAULT
    s_resp = ","
    If (s_ult_sep = ".") Then
        If (n_ponto = 1) And (n_virg = 0) And (n_digitos_finais = 3) Then
        '   NOP: CONSIDERA 123.456 COMO CENTO E VINTE E TRÊS MIL E QUATROCENTOS E CINQUENTA E SEIS
        ElseIf (n_ponto = 1) Then
            s_resp = "."
            End If
    ElseIf (s_ult_sep = ",") Then
        If (n_virg > 1) And (n_ponto = 0) Then s_resp = "."
        End If
        
    retorna_separador_decimal = s_resp
    
End Function


Public Function retorna_separador_decimal_sistema() As String
Dim i As Integer
Dim c As String
Dim s As String
Dim s_sep_sys As String

    retorna_separador_decimal_sistema = ""
    s_sep_sys = ""
    s = CStr(0.5)
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If Not IsNumeric(c) Then
            s_sep_sys = c
            Exit For
            End If
        Next
    If s_sep_sys = "" Then Exit Function
    retorna_separador_decimal_sistema = s_sep_sys
End Function

Function RTrimCrLf(ByVal Texto As String) As String
' __________________________________________________________________________________
'|
'|  REMOVE CARACTERES BRANCO, CR E LF DO FINAL DO TEXTO
'|

Dim c As String
Dim s_resp As String
Dim i As Long

    s_resp = ""
    For i = Len(Texto) To 1 Step -1
        c = Mid$(Texto, i, 1)
        If (c <> " ") And (c <> vbCr) And (c <> vbLf) Then
            s_resp = Mid$(Texto, 1, i)
            Exit For
            End If
        Next
        
    RTrimCrLf = s_resp
    
End Function

Function LTrimCrLf(ByVal Texto As String) As String
' __________________________________________________________________________________
'|
'|  REMOVE CARACTERES BRANCO, CR E LF DO INÍCIO DO TEXTO
'|

Dim c As String
Dim s_resp As String
Dim i As Long

    s_resp = ""
    For i = 1 To Len(Texto)
        c = Mid$(Texto, i, 1)
        If (c <> " ") And (c <> vbCr) And (c <> vbLf) Then
            s_resp = Mid$(Texto, i)
            Exit For
            End If
        Next
        
    LTrimCrLf = s_resp
    
End Function

Public Function sqlFormataDecimal(ByVal valor As Currency) As String
Dim strValorFormatado As String
Dim strSeparadorDecimal As String
Dim vlNumeroAuxiliar As Currency
Dim strNumeroAuxiliar As String

    vlNumeroAuxiliar = 0.5
    strNumeroAuxiliar = CStr(vlNumeroAuxiliar)
    
    If InStr(strNumeroAuxiliar, ".") > 0 Then
        strSeparadorDecimal = "."
    ElseIf InStr(strNumeroAuxiliar, ",") > 0 Then
        strSeparadorDecimal = ","
        End If
        
    strValorFormatado = CStr(valor)
    If Len(strSeparadorDecimal) > 0 Then
        strValorFormatado = Replace(strValorFormatado, strSeparadorDecimal, "V")
        strValorFormatado = Replace(strValorFormatado, ".", "")
        strValorFormatado = Replace(strValorFormatado, ",", "")
        strValorFormatado = Replace(strValorFormatado, "V", ".")
        End If
        
    sqlFormataDecimal = strValorFormatado
End Function

Public Function sqlMontaDateParaSqlDateTime(ByVal dtReferencia As Date) As String
' __________________________________________________________________________________________________
'|
'|  GERA UMA EXPRESSÃO PARA O SQL SERVER PODER MANIPULAR A DATA REPRESENTADA PELO PARÂMETRO
'|

Dim strDataHora As String
Dim strSql As String

    strDataHora = Format$(dtReferencia, "yyyy-mm-dd hh:mm:ss")
    strSql = "Convert(datetime, '" & strDataHora & "', 120)"
    sqlMontaDateParaSqlDateTime = strSql
End Function

Function TrimCrLf(ByVal Texto As String) As String
' __________________________________________________________________________________
'|
'|  REMOVE CARACTERES BRANCO, CR E LF DO INÍCIO E DO FINAL DO TEXTO
'|

    TrimCrLf = LTrimCrLf(RTrimCrLf(Texto))
    
End Function


Function tem_digito(ByVal Texto As String) As Boolean
Dim i As Long
Dim achou As Boolean
    
    tem_digito = False
    
    Texto = Trim("" & Texto)
    achou = False
    For i = 1 To Len(Texto)
        If IsNumeric(Mid(Texto, i, 1)) Then
            achou = True
            Exit For
            End If
        Next
    
    If achou Then tem_digito = True
    
End Function


Function iniciais_em_maiusculas(ByVal Texto As String) As String
Const palavras_minusculas = "|A|AS|E|O|OS|UM|UNS|UMA|UMAS|DA|DAS|DE|DO|DOS|EM|NA|NAS|NO|NOS|COM|SEM|POR|PELO|PARA|PRA|P/|S/|C/|TEM|OU|E/OU|"
Const palavras_maiusculas = "|II|III|IV|VI|VII|VIII|IX|XI|XII|XIII|XIV|XV|XVI|XVII|XVIII|XIX|XX|XXI|XXII|XXIII|"
Dim letra As String
Dim palavra As String
Dim frase As String
Dim s As String
Dim i As Long
Dim i_max As Long
    
    iniciais_em_maiusculas = ""
    
    frase = ""
    palavra = ""
    Texto = "" & Texto
    i_max = Len(Texto)
    For i = 1 To i_max
        letra = Mid(Texto, i, 1)
        palavra = palavra & letra
        If (letra = " ") Or (i = i_max) Then
            s = "|" & UCase(Trim(palavra)) & "|"
            If (InStr(palavras_minusculas, s) <> 0) And (frase <> "") Then
                palavra = LCase(palavra)
            ElseIf (InStr(palavras_maiusculas, s) <> 0) Then
                palavra = UCase(palavra)
            Else
            '   SE POSSUI DÍGITOS, ENTÃO É ALGUM TIPO DE CÓDIGO
                If Not tem_digito(palavra) Then
                    palavra = UCase(left(palavra, 1)) & LCase(Mid(palavra, 2))
                    End If
                End If
            frase = frase & palavra
            palavra = ""
            End If
        Next
    
    iniciais_em_maiusculas = frase
    
End Function


Function cnpj_cpf_formata(ByVal cnpj_cpf As String) As String
Dim s  As String
Dim s_resp  As String

    cnpj_cpf = "" & cnpj_cpf
    s = retorna_so_digitos(cnpj_cpf)
    
  ' CPF
    If Len(s) = 11 Then
        s_resp = Mid(s, 1, 3) & "." & Mid(s, 4, 3) & "." & Mid(s, 7, 3) & "/" & Mid(s, 10, 2)
  ' CNPJ
    ElseIf Len(s) = 14 Then
        s_resp = Mid(s, 1, 2) & "." & Mid(s, 3, 3) & "." & Mid(s, 6, 3) & "/" & Mid(s, 9, 4) & "-" & Mid(s, 13, 2)
  ' DESCONHECIDO
    Else
        s_resp = cnpj_cpf
        End If
    
    cnpj_cpf_formata = s_resp
    
End Function

Function retorna_so_digitos(ByVal s_numero As String) As String
Dim s As String
Dim c As String
Dim i As Integer
    s = ""
    For i = 1 To Len(s_numero)
        c = Mid(s_numero, i, 1)
        If IsNumeric(c) Then s = s & c
        Next
    retorna_so_digitos = s
End Function

Function cep_formata(ByVal cep As String) As String
Dim s_cep As String
    s_cep = "" & cep
    s_cep = retorna_so_digitos(s_cep)
    
    cep_formata = s_cep
    
    If ((s_cep = "") Or (Len(s_cep) = 5) Or (Not cep_ok(s_cep))) Then Exit Function
    
    s_cep = Mid(s_cep, 1, 5) & "-" & Mid(s_cep, 6, 3)
    
    cep_formata = s_cep
    
End Function

Function cep_ok(ByVal cep As String) As Boolean
Dim s_cep As String
    cep_ok = False
    s_cep = "" & cep
    s_cep = retorna_so_digitos(s_cep)
    If ((Len(s_cep) = 0) Or (Len(s_cep) = 5) Or (Len(s_cep) = 8)) Then cep_ok = True
End Function

Function telefone_formata(ByVal telefone As String) As String
Dim i As Integer
Dim s_tel As String
    s_tel = "" & telefone
    s_tel = retorna_so_digitos(s_tel)
    
    telefone_formata = s_tel
    
    If ((s_tel = "") Or (Len(s_tel) > 8) Or (Not telefone_ok(s_tel))) Then Exit Function
         
    i = Len(s_tel) - 4
    s_tel = Mid(s_tel, 1, i) & "-" & Mid(s_tel, i + 1, Len(s_tel))
    
    telefone_formata = s_tel
    
End Function

Function telefone_formata_2(ByVal telefone As String) As String
Dim i As String
Dim s_tel
    
    s_tel = "" & telefone
    s_tel = retorna_so_digitos(s_tel)
    
    telefone_formata_2 = s_tel
    
    If ((s_tel = "") Or (Len(s_tel) > 9) Or (Not telefone_ok(s_tel))) Then Exit Function
         
    i = Len(s_tel) - 4
    s_tel = Mid(s_tel, 1, i) & "-" & Mid(s_tel, i + 1, Len(s_tel))
    
    telefone_formata_2 = s_tel
    
End Function


Function telefone_ok(ByVal telefone As String) As Boolean
Dim s_tel As String
    telefone_ok = False
    s_tel = "" & telefone
    s_tel = retorna_so_digitos(s_tel)
    If ((Len(s_tel) = 0) Or (Len(s_tel) >= 6)) Then telefone_ok = True
End Function

Function IsAlgarismo(ByVal c As String) As Boolean
    c = Trim("" & c)
    IsAlgarismo = ((c >= "0") And (c <= "9") And (Len(c) = 1))
End Function


Function IsLetra(ByVal c As String) As Boolean
    c = UCase(Trim("" & c))
    IsLetra = ((c >= "A") And (c <= "Z") And (Len(c) = 1))
End Function

Function maiuscula(ByVal KeyAscii As Integer) As Integer
' ______________________________________________________________________________
'|
'|  TRANSFORMA EM LETRA EM MAIÚSCULA
'|

    maiuscula = Asc(UCase$(Chr$(KeyAscii)))

End Function

Function cor_em_decimal_para_rgb(ByVal cor_em_decimal As Long) As String
Dim s As String
    s = Hex(cor_em_decimal)
    Do While Len(s) < 6: s = "0" & s: Loop
    s = Mid$(s, 5, 2) & Mid$(s, 3, 2) & Mid$(s, 1, 2)
    cor_em_decimal_para_rgb = s
End Function

Function cor_nova_tonalidade(ByVal cor As Long, ByVal percentual As Long) As Long
' ________________________________________________________________________________________________
'|
'|  CALCULA NOVAS TONALIDADES PARA A COR ESPECIFICADA.
'|  PARA MANTER A TONALIDADE, O DESLOCAMENTO É FEITO POR UM FATOR PERCENTUAL.
'|  FATORES POSITIVOS INDICAM CORES MAIS CLARAS E
'|  FATORES NEGATIVOS CORES MAIS ESCURAS.
'|

Dim s As String
Dim lr As Long
Dim lg As Long
Dim lb As Long


    s = Hex(cor)
    Do While Len(s) < 6: s = "0" & s: Loop
    
  ' RED
    lr = CLng("&H" & Mid$(s, 5, 2))
    lr = lr + (lr * percentual / 100)
    If lr < 0 Then lr = 0
    If lr > 255 Then lr = 255
    
  ' GREEN
    lg = CLng("&H" & Mid$(s, 3, 2))
    lg = lg + (lg * percentual / 100)
    If lg < 0 Then lg = 0
    If lg > 255 Then lg = 255
    lg = CLng("&H" & Hex(lg) & "00")
        
  ' BLUE
    lb = CLng("&H" & Mid$(s, 1, 2))
    lb = lb + (lb * percentual / 100)
    If lb < 0 Then lb = 0
    If lb > 255 Then lb = 255
    lb = CLng("&H" & Hex(lb) & "0000")
        
    
    cor_nova_tonalidade = lr + lg + lb
    
    
End Function



Function obtem_mes_em_portugues(ByVal numero_mes As Variant, Optional ByVal abreviado As Boolean = False) As String
' ________________________________________________________________________________
'|
'|  RETORNA O MÊS COM OS 3 PRIMEIROS CARACTERES E EM PORTUGUÊS.
'|

Dim s As String
Dim j As Integer

    
    On Error Resume Next

    
    obtem_mes_em_portugues = ""
    
    If IsNumeric(numero_mes) Then j = CInt(numero_mes) Else j = 0

    Select Case j
        Case 1: s = "JANEIRO"
        Case 2: s = "FEVEREIRO"
        Case 3: s = "MARÇO"
        Case 4: s = "ABRIL"
        Case 5: s = "MAIO"
        Case 6: s = "JUNHO"
        Case 7: s = "JULHO"
        Case 8: s = "AGOSTO"
        Case 9: s = "SETEMBRO"
        Case 10: s = "OUTUBRO"
        Case 11: s = "NOVEMBRO"
        Case 12: s = "DEZEMBRO"
        Case Else: s = ""
        End Select

    
    If abreviado Then s = Mid$(s, 1, 3)
    
    obtem_mes_em_portugues = s


End Function



Public Sub INFO_Inicia(ByRef f As Form)
' __________________________________________________________________________________________________
'|
'|  INICIALIZA ÁREA DE INFORMAÇÕES
'|

Const X_SEPARACAO = 5
Dim s As String


    With f
        .base_status.BackColor = vbMenuBar
        .agora.BackColor = .base_status.BackColor
        .hoje.BackColor = .base_status.BackColor
        .info.BackColor = .base_status.BackColor
    
        .base_status.ScaleMode = f.ScaleMode
        .base_status.Height = .info.Height + 6
        .base_status.BorderStyle = vbBSNone
        .base_status.Line (0, 0)-(f.ScaleWidth, 0), f.BackColor
        .base_status.Line (0, 0)-(f.ScaleWidth, 0), vbMenuBar
        .base_status.Line (0, 1)-(f.ScaleWidth, 1), vb3DHighlight
        .base_status.Align = vbAlignNone
        .base_status.top = .base_status.top + 1
        
        .agora.Move 2, 4
        desenha_borda_system .base_status, .agora, -1
        .agora = left$(Time$, 5)
    
        .hoje.Move .agora.left + .agora.Width + X_SEPARACAO, .agora.top
        desenha_borda_system .base_status, .hoje, -1
        s = Format$(Date, "dd/mm/yyyy")
        s = Mid$(s, 1, 2) & "." & obtem_mes_em_portugues(Mid(s, 4, 2), True) & "." & Mid$(s, 7, 4)
        .hoje = s
        
        .info.Move .hoje.left + .hoje.Width + X_SEPARACAO, .agora.top, .base_status.ScaleWidth - (.hoje.left + .hoje.Width + X_SEPARACAO) - .agora.left
        desenha_borda_system .base_status, .info, -1
        
        .base_status.top = f.ScaleHeight - .base_status.Height
        .base_status.ZOrder
        End With
        
        
End Sub


Sub desenha_borda_system(f As Control, c As Control, l As Integer)

Dim claro As Long
Dim escuro As Long
    
'  DESENHA AS SOMBRAS
'  ~~~~~~~~~~~~~~~~~~
    If l >= 0 Then
        claro = vb3DHighlight
        escuro = vb3DDKShadow
    Else
        claro = vb3DDKShadow
        escuro = vb3DHighlight
        End If
    
    f.ScaleMode = vbPixels
    f.Line (c.left - 2, c.top - 1)-(c.left + c.Width + 1, c.top - 1), claro
    f.Line (c.left - 2, c.top - 1)-(c.left - 2, c.top + c.Height + 1), claro
    f.Line (c.left + c.Width + 1, c.top - 1)-(c.left + c.Width + 1, c.top + c.Height + 2), escuro
    f.Line (c.left - 2, c.top + c.Height + 1)-(c.left + c.Width + 1, c.top + c.Height + 1), escuro

End Sub



Sub aguarde(ByVal tipo As INFO_STATUS_OPTIONS, Optional ByVal mensagem As String = m_id)

Const TEXTO_PONTINHOS = " . . ."
Dim s As String
Dim f As Form
Dim ha_status_bar As Boolean

    On Error Resume Next

    If painel_ativo Is Nothing Then Exit Sub
    
    ha_status_bar = True
    Set f = painel_ativo
    
    On Error GoTo AGUARDE_TRATA_AUSENCIA_STATUS_BAR
    s = painel_ativo.info.Name
    If Not ha_status_bar Then Set f = painel_principal
    On Error Resume Next

    Select Case tipo
        Case INFO_EXECUTANDO
            Screen.MousePointer = vbHourglass
            Screen.ActiveForm.MousePointer = vbHourglass
            f.info.BackColor = COR_AMARELO
            f.info.FontBold = True
            s = mensagem
            If right$(s, Len(TEXTO_PONTINHOS)) <> TEXTO_PONTINHOS Then s = s & TEXTO_PONTINHOS
            f.info = s
        Case INFO_NORMAL
            f.info.BackColor = f.hoje.BackColor
            f.info.FontBold = False
            f.info = mensagem
            Screen.ActiveForm.MousePointer = vbDefault
            Screen.MousePointer = vbDefault
            End Select
    
    f.Refresh
    
Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
AGUARDE_TRATA_AUSENCIA_STATUS_BAR:
'=================================
    ha_status_bar = False
    Resume Next
    
End Sub


Public Function maior(num_1, num_2) As Variant
' ______________________________________________________________________________
'|
'|  RETORNA O MAIOR DOS 2 NÚMEROS
'|

    If num_1 > num_2 Then
        maior = num_1
    Else
        maior = num_2
        End If
        
End Function
Public Function menor(num_1, num_2) As Variant
' ______________________________________________________________________________
'|
'|  RETORNA O MENOR DOS 2 NÚMEROS
'|

    If num_1 < num_2 Then
        menor = num_1
    Else
        menor = num_2
        End If
        
End Function


Function retorna_texto_sem_nulos(ByVal Texto As String) As String
' _______________________________________________________________________________
'|
'|  RETORNA O TEXTO ENCONTRADO ANTES DO PRIMEIRO CARACTER NULO.
'|

Dim s As String
Dim i As Variant

    s = ""

    i = InStr(1, Texto, Chr$(0))
    If Not IsNumeric(i) Then i = 0
    
    i = i - 1

  ' NÃO ACHOU CARACTER NULO
    If i = -1 Then
        s = Texto
  ' COPIA SOMENTE ATÉ O CARACTER NULO
    ElseIf i > 0 Then
        s = Mid$(Texto, 1, i)
        End If

   retorna_texto_sem_nulos = s

End Function

Sub alerta(ByVal mensagem As String)

Dim status_ok As Boolean
Dim status_a As Integer
Dim mensagem_a As String

    
    status_ok = obtem_status_painel_info(status_a, mensagem_a)
        
    aguarde INFO_NORMAL
    
    MsgBox mensagem, vbOKOnly + vbExclamation, App.Title

    If status_ok Then aguarde status_a, mensagem_a
    
    
End Sub

Sub aviso(ByVal mensagem As String)

Dim status_ok As Boolean
Dim status_a As Integer
Dim mensagem_a As String

    
    status_ok = obtem_status_painel_info(status_a, mensagem_a)
        
    aguarde INFO_NORMAL
    
    MsgBox mensagem, vbOKOnly + vbInformation, App.Title

    If status_ok Then aguarde status_a, mensagem_a
    
    
End Sub



Sub aviso_erro(ByVal mensagem As String)

Dim status_ok As Boolean
Dim status_a As Integer
Dim mensagem_a As String

    
    status_ok = obtem_status_painel_info(status_a, mensagem_a)
        
    aguarde INFO_NORMAL
    
    MsgBox mensagem, vbOKOnly + vbCritical, App.Title
    
    If status_ok Then aguarde status_a, mensagem_a
    
End Sub




Function confirma(ByVal mensagem As String) As Boolean

Dim status_ok As Boolean
Dim status_a As Integer
Dim mensagem_a As String

    
    status_ok = obtem_status_painel_info(status_a, mensagem_a)
        
    aguarde INFO_NORMAL
    
    confirma = (MsgBox(mensagem, vbYesNo + vbQuestion, App.Title) = vbYes)
    
    If status_ok Then aguarde status_a, mensagem_a
    
    
End Function



Public Function obtem_status_painel_info(ByRef status As Integer, ByRef mensagem As String) As Boolean
' ____________________________________________________________________________________
'|
'|  OBTÉM O STATUS E A MENSAGEM APRESENTADA NO PAINEL INFO (MENSAGENS DE STATUS).
'|


    On Error GoTo OSPI_TRATA_ERRO
    
    obtem_status_painel_info = False
    
  ' DEFAULT
    status = INFO_NORMAL
    mensagem = m_id
    
    If painel_ativo Is Nothing Then Exit Function
    
    If painel_ativo.info.BackColor = COR_AMARELO Then status = INFO_EXECUTANDO
    mensagem = painel_ativo.info.Caption
    
    obtem_status_painel_info = True
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OSPI_TRATA_ERRO:
'===============
    Err.Clear
    Exit Function
    
    
End Function

Function DomingoDePascoa(ByVal ano As Integer) As Date
' ____________________________________________________________________________________
'|
'|  OBTÉM A DATA DO DOMINGO DE PÁSCOA PARA O ANO FORNECIDO
'|  (objetivo: obter o último domingo do horário de verão, caso o terceiro domingo de fevereiro caia no Carnaval)
'|

    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    Dim e As Integer
    Dim f As Integer
    Dim g As Integer
    Dim h As Integer
    Dim i As Integer
    Dim k As Integer
    Dim l As Integer
    Dim m As Integer
    Dim mes As Integer
    Dim dia As Integer
    Dim domingo As Date

    a = ano Mod 19
    b = CInt(Int(ano / 100))
    c = ano Mod 100
    d = CInt(Int(b / 4))
    e = b Mod 4
    f = CInt(Int((b + 8) / 25))
    g = CInt(Int((b - f + 1) / 3))
    h = (19 * a + b - d - g + 15) Mod 30
    i = CInt(Int(c / 4))
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = CInt(Int((a + 11 * h + 22 * l) / 451))
    mes = CInt(Int((h + l - 7 * m + 114) / 31))
    dia = ((h + l - 7 * m + 114) Mod 31) + 1
    domingo = DateSerial(ano, mes, dia)

    DomingoDePascoa = domingo
End Function

Function DomingoDeCarnaval(ByVal ano As Integer) As Date
' ____________________________________________________________________________________
'|
'|  OBTÉM A DATA DO DOMINGO DE CARNAVAL PARA O ANO FORNECIDO
'|  (objetivo: obter o último domingo do horário de verão, caso o terceiro domingo de fevereiro caia no Carnaval)
'|
    Dim domingo

    domingo = DomingoDePascoa(ano) - 49

    DomingoDeCarnaval = domingo
End Function

Function InicioHorarioVerao(ByVal ano As Integer) As Date
' ____________________________________________________________________________________
'|
'|  OBTÉM A DATA DO DOMINGO DE INÍCIO DO HORÁRIO DE VERÃO
'|  (segundo DECRETO Nº 6.558, DE 8 DE SETEMBRO DE 2008)
'|
    'terceiro domingo de outubro
    Dim primeiroDeOutubro As Date
    Dim primeiroDomingoDeOutubro As Date
    Dim terceiroDomingoDeOutubro As Date
    
    primeiroDeOutubro = DateSerial(ano, 10, 1)
    'primeiroDomingoDeOutubro = primeiroDeOutubro + ((7 - CInt(Int(primeiroDeOutubro.DayOfWeek))) Mod 7)
    primeiroDomingoDeOutubro = primeiroDeOutubro
    Do While DatePart("w", primeiroDomingoDeOutubro) <> 1
        primeiroDomingoDeOutubro = primeiroDomingoDeOutubro + 1
        DoEvents
        Loop
    terceiroDomingoDeOutubro = primeiroDomingoDeOutubro + 14

    InicioHorarioVerao = terceiroDomingoDeOutubro
End Function

Function TerminoHorarioVerao(ByVal ano As Integer) As Date
' ____________________________________________________________________________________
'|
'|  OBTÉM A DATA DO DOMINGO DE TÉRMINO DO HORÁRIO DE VERÃO
'|  (segundo DECRETO Nº 6.558, DE 8 DE SETEMBRO DE 2008)
'|
    Dim primeiroDeFevereiro As Date
    Dim primeiroDomingoDeFevereiro As Date
    Dim terceiroDomingoDeFevereiro As Date

    primeiroDeFevereiro = DateSerial(ano + 1, 2, 1)
    'primeiroDomingoDeFevereiro = primeiroDeFevereiro.AddDays(CInt(Int(7 - primeiroDeFevereiro.DayOfWeek)) Mod 7)
    primeiroDomingoDeFevereiro = primeiroDeFevereiro
    Do While DatePart("w", primeiroDomingoDeFevereiro) <> 1
        primeiroDomingoDeFevereiro = primeiroDomingoDeFevereiro + 1
        DoEvents
        Loop
    terceiroDomingoDeFevereiro = primeiroDomingoDeFevereiro + 14
    If (terceiroDomingoDeFevereiro = DomingoDeCarnaval(ano)) Then
        terceiroDomingoDeFevereiro = terceiroDomingoDeFevereiro + 7
        End If

    TerminoHorarioVerao = terceiroDomingoDeFevereiro
End Function

Public Function converte_cor_VB2Web(colorcode As String) As String
Dim vcolor
    vcolor = Hex(Val(colorcode))
    If Len(vcolor) < 6 Then
       vcolor = String(6 - Len(vcolor), "0") & vcolor
        End If
    converte_cor_VB2Web = Mid(vcolor, 5, 2) & Mid(vcolor, 3, 2) & Mid(vcolor, 1, 2)
End Function

Public Function converte_cor_Web2VB(colorcode As String) As String
    If left(colorcode, 1) = "#" Then colorcode = Mid(colorcode, 2, Len(colorcode) - 1)
    If Len(colorcode) < 6 Then colorcode = colorcode & String(6 - Len(colorcode), "0")
    converte_cor_Web2VB = "&H" & Mid(colorcode, 5, 2) & Mid(colorcode, 3, 2) & Mid(colorcode, 1, 2)
End Function

'Function configura_registry_usuario_cd_padrao(cd_ini As String, cd_fim As String) As Boolean
'' ------------------------------------------------------------------------
''   CONFIGURA O REGISTRY PARA QUE O PROGRAMA MEMORIZE
''   O ID DO CENTRO DE DISTRIBUIÇÃO SELECIONADO PELO USUÁRIO,
''   DE FORMA A PRÉ-SELECIONAR O MESMO CD NA PRÓXIMA
''   ENTRADA NO PROGRAMA.
'
'Dim s_chave As String
'Dim s_campo As String
'Dim s_valor As String
'
'    configura_registry_usuario_cd_padrao = False
'
'    If (cd_ini <> cd_fim) Then
'        s_chave = REG_CHAVE_USUARIO_USUARIO
'
'        s_campo = "UsuCD"
'        s_valor = cd_fim
'        If Not registry_usuario_grava_string(s_chave, s_campo, s_valor) Then Exit Function
'
'        End If
'
'    configura_registry_usuario_cd_padrao = True
'
'End Function
'
'Function le_registry_usuario_cd_padrao(cd_valor As String) As Boolean
'' ------------------------------------------------------------------------
''   LE O REGISTRY PARA QUE O PROGRAMA VERIFIQUE QUAL O CENTRO DE
''   DISTRIBUIÇÃO QUE O USUÁRIO COSTUMA ACESSAR, SE HOUVER
'
'Dim s_chave As String
'Dim s_campo As String
'Dim s_valor As String
'
'    le_registry_usuario_cd_padrao = False
'
'    s_chave = REG_CHAVE_USUARIO_USUARIO
'
'    s_campo = "UsuCD"
'    s_valor = ""
'    If Not registry_usuario_le_string(s_chave, s_campo, s_valor) Then Exit Function
'    cd_valor = s_valor
'
'    le_registry_usuario_cd_padrao = True
'
'End Function

Function configura_registry_usuario_emit_padrao(emit_ini As String, emit_fim As String) As Boolean
' ------------------------------------------------------------------------
'   CONFIGURA O REGISTRY PARA QUE O PROGRAMA MEMORIZE
'   O ID DO EMITENTE SELECIONADO PELO USUÁRIO,
'   DE FORMA A PRÉ-SELECIONAR O MESMO EMITENTE NA PRÓXIMA
'   ENTRADA NO PROGRAMA.

Dim s_chave As String
Dim s_campo As String
Dim s_valor As String

    configura_registry_usuario_emit_padrao = False
    
    If (emit_ini <> emit_fim) Then
        s_chave = REG_CHAVE_USUARIO_USUARIO
        
        s_campo = "UsuEmit"
        s_valor = emit_fim
        If Not registry_usuario_grava_string(s_chave, s_campo, s_valor) Then Exit Function
        
        End If
        
    configura_registry_usuario_emit_padrao = True
    
End Function

Function le_registry_usuario_emit_padrao(emit_valor As String) As Boolean
' ------------------------------------------------------------------------
'   LE O REGISTRY PARA QUE O PROGRAMA VERIFIQUE QUAL O EMITENTE
'   QUE O USUÁRIO COSTUMA ACESSAR, SE HOUVER

Dim s_chave As String
Dim s_campo As String
Dim s_valor As String

    le_registry_usuario_emit_padrao = False
    
    s_chave = REG_CHAVE_USUARIO_USUARIO
    
    s_campo = "UsuEmit"
    s_valor = ""
    If Not registry_usuario_le_string(s_chave, s_campo, s_valor) Then Exit Function
    emit_valor = s_valor
    
    le_registry_usuario_emit_padrao = True
    
End Function

Function configura_registry_usuario_cor_fundo_padrao(cor As String) As Boolean
' ------------------------------------------------------------------------
'   CONFIGURA O REGISTRY PARA QUE O PROGRAMA MEMORIZE
'   A COR DE FUNDO DOS FORMs QUE DEVE SER EXIBIDA
'   NA ABERTURA DO PROGRAMA, ANTES DE CONSULTÁ-LA
'   NO BANCO DE DADOS.

Dim s_chave As String
Dim s_campo As String
Dim s_valor As String

    configura_registry_usuario_cor_fundo_padrao = False
    
    s_chave = REG_CHAVE_USUARIO_USUARIO
    
    s_campo = "UsuCorFundoPadrao"
    s_valor = cor
    If Not registry_usuario_grava_string(s_chave, s_campo, s_valor) Then Exit Function
        
    configura_registry_usuario_cor_fundo_padrao = True
    
End Function

Function le_registry_usuario_cor_fundo_padrao(cor As String) As Boolean
' ------------------------------------------------------------------------
'   LE O REGISTRY PARA QUE O PROGRAMA VERIFIQUE QUAL COR DE FUNDO DOS FORMs
'   DEVE SER EXIBIDA, ANTES DE CONSULTÁ-LA NO BANCO DE DADOS

Dim s_chave As String
Dim s_campo As String
Dim s_valor As String

    le_registry_usuario_cor_fundo_padrao = False
    
    s_chave = REG_CHAVE_USUARIO_USUARIO
    
    s_campo = "UsuCorFundoPadrao"
    s_valor = ""
    If Not registry_usuario_le_string(s_chave, s_campo, s_valor) Then Exit Function
    cor = s_valor
    
    le_registry_usuario_cor_fundo_padrao = True
    
End Function

