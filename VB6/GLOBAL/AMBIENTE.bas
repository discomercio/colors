Attribute VB_Name = "mod_AMBIENTE"
Option Explicit

'  ROTINAS PARA CONFIGURAÇÃO REGIONAL
'  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 1. ESTE MÓDULO CONTÉM FUNÇÕES QUE CONFIGURAM OU RESTAURAM OS PARÂMETROS DA
'    CONFIGURAÇÃO REGIONAL.
'    ISSO É MUITO IMPORTANTE PORQUE AFETA FUNÇÕES COMO CDATE/CVDATE, CCUR, ETC.
'    IMAGINE DUAS VARIÁVEIS STRING CONTENDO "10/01/2001" (10.JAN.2001) E
'    "123.456,78".  SE A CONFIGURAÇÃO REGIONAL FOR "MM/DD/YY" P/ DATAS E O
'    SEPARADOR DECIMAL FOR "." AO INVÉS DE ",", AS FUNÇÕES CDATE E CCUR IRÃO
'    RETORNAR "01/10/2001" (01.OUT.2001) E 123,46
' 2. OS PARÂMETROS DA CONFIGURAÇÃO REGIONAL QUE FOREM ALTERADOS POR ESTE MÓDULO
'    TERÃO SEUS VALORES ORIGINAIS GRAVADOS EM UM ARQUIVO .INI
' 3. IMPORTANTE: AS ALTERAÇÕES NA CONFIGURAÇÃO REGIONAL NÃO FAZEM EFEITO P/
'    OS PROGRAMAS QUE ESTÃO EM EXECUÇÃO, NEM MESMO P/ ESTE MÓDULO QUE ESTÁ
'    EFETUANDO OS COMANDOS DE ALTERAÇÃO.  É PRECISO QUE OS PROGRAMAS SEJAM
'    INICIADOS DEPOIS QUE AS ALTERAÇÕES TENHAM SIDO FEITAS.
' 4. A ALTERAÇÃO NA CONFIGURAÇÃO REGIONAL CONTINUA MESMO SE FOR FEITO UM
'    REBOOT. ENTRETANTO, SE UM OUTRO USUÁRIO FIZER O LOGON NA MÁQUINA, ENTÃO
'    A CONFIGURAÇÃO REGIONAL SERÁ A DEFAULT.
' 


Const NOME_ARQ_INI_AMBIENTE = "SC_AMB.INI"
Const SECAO_PARAMETROS_ORIGINAIS = "CONFIGURACAO_REGIONAL_ORIGINAL"
Const REG_CHAVE_PARAMETROS_ORIGINAIS = "SOFTWARE\PRAGMATICA\Sistema Contratos\AMB\Parametros Originais"



' MANIPULAÇÃO DA CONFIGURAÇÃO REGIONAL
Public Const WM_SETTINGCHANGE = &H1A
'same as the old WM_WININICHANGE
Public Const HWND_BROADCAST = &HFFFF&

Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal szlpLCData As Integer) As Boolean
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Declare Function GetUserDefaultLCID Lib "kernel32" () As Long


' Locale Types.
' These types are used for the GetLocaleInfoW NLS API routine.

' LOCALE_NOUSEROVERRIDE is also used in GetTimeFormatW and GetDateFormatW.
Public Const LOCALE_NOUSEROVERRIDE = &H80000000  ' do not use user overrides

Public Const LOCALE_ILANGUAGE = &H1         '  language id
Public Const LOCALE_SLANGUAGE = &H2         '  localized name of language
Public Const LOCALE_SENGLANGUAGE = &H1001      '  English name of language
Public Const LOCALE_SABBREVLANGNAME = &H3         '  abbreviated language name
Public Const LOCALE_SNATIVELANGNAME = &H4         '  native name of language
Public Const LOCALE_ICOUNTRY = &H5         '  country code
Public Const LOCALE_SCOUNTRY = &H6         '  localized name of country
Public Const LOCALE_SENGCOUNTRY = &H1002      '  English name of country
Public Const LOCALE_SABBREVCTRYNAME = &H7         '  abbreviated country name
Public Const LOCALE_SNATIVECTRYNAME = &H8         '  native name of country
Public Const LOCALE_IDEFAULTLANGUAGE = &H9         '  default language id
Public Const LOCALE_IDEFAULTCOUNTRY = &HA         '  default country code
Public Const LOCALE_IDEFAULTCODEPAGE = &HB         '  default code page

Public Const LOCALE_SLIST = &HC         '  list item separator
Public Const LOCALE_IMEASURE = &HD         '  0 = metric, 1 = US

Public Const LOCALE_SDECIMAL = &HE         '  decimal separator
Public Const LOCALE_STHOUSAND = &HF         '  thousand separator
Public Const LOCALE_SGROUPING = &H10        '  digit grouping
Public Const LOCALE_IDIGITS = &H11        '  number of fractional digits
Public Const LOCALE_ILZERO = &H12        '  leading zeros for decimal
Public Const LOCALE_SNATIVEDIGITS = &H13        '  native ascii 0-9

Public Const LOCALE_SCURRENCY = &H14        '  local monetary symbol
Public Const LOCALE_SINTLSYMBOL = &H15        '  intl monetary symbol
Public Const LOCALE_SMONDECIMALSEP = &H16        '  monetary decimal separator
Public Const LOCALE_SMONTHOUSANDSEP = &H17        '  monetary thousand separator
Public Const LOCALE_SMONGROUPING = &H18        '  monetary grouping
Public Const LOCALE_ICURRDIGITS = &H19        '  # local monetary digits
Public Const LOCALE_IINTLCURRDIGITS = &H1A        '  # intl monetary digits
Public Const LOCALE_ICURRENCY = &H1B        '  positive currency mode
Public Const LOCALE_INEGCURR = &H1C        '  negative currency mode

Public Const LOCALE_SDATE = &H1D        '  date separator
Public Const LOCALE_STIME = &H1E        '  time separator
Public Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Public Const LOCALE_SLONGDATE = &H20        '  long date format string
Public Const LOCALE_STIMEFORMAT = &H1003      '  time format string
Public Const LOCALE_IDATE = &H21        '  short date format ordering
Public Const LOCALE_ILDATE = &H22        '  long date format ordering
Public Const LOCALE_ITIME = &H23        '  time format specifier
Public Const LOCALE_ICENTURY = &H24        '  century format specifier
Public Const LOCALE_ITLZERO = &H25        '  leading zeros in time field
Public Const LOCALE_IDAYLZERO = &H26        '  leading zeros in day field
Public Const LOCALE_IMONLZERO = &H27        '  leading zeros in month field
Public Const LOCALE_S1159 = &H28        '  AM designator
Public Const LOCALE_S2359 = &H29        '  PM designator

Public Const LOCALE_SDAYNAME1 = &H2A        '  long name for Monday
Public Const LOCALE_SDAYNAME2 = &H2B        '  long name for Tuesday
Public Const LOCALE_SDAYNAME3 = &H2C        '  long name for Wednesday
Public Const LOCALE_SDAYNAME4 = &H2D        '  long name for Thursday
Public Const LOCALE_SDAYNAME5 = &H2E        '  long name for Friday
Public Const LOCALE_SDAYNAME6 = &H2F        '  long name for Saturday
Public Const LOCALE_SDAYNAME7 = &H30        '  long name for Sunday
Public Const LOCALE_SABBREVDAYNAME1 = &H31        '  abbreviated name for Monday
Public Const LOCALE_SABBREVDAYNAME2 = &H32        '  abbreviated name for Tuesday
Public Const LOCALE_SABBREVDAYNAME3 = &H33        '  abbreviated name for Wednesday
Public Const LOCALE_SABBREVDAYNAME4 = &H34        '  abbreviated name for Thursday
Public Const LOCALE_SABBREVDAYNAME5 = &H35        '  abbreviated name for Friday
Public Const LOCALE_SABBREVDAYNAME6 = &H36        '  abbreviated name for Saturday
Public Const LOCALE_SABBREVDAYNAME7 = &H37        '  abbreviated name for Sunday
Public Const LOCALE_SMONTHNAME1 = &H38        '  long name for January
Public Const LOCALE_SMONTHNAME2 = &H39        '  long name for February
Public Const LOCALE_SMONTHNAME3 = &H3A        '  long name for March
Public Const LOCALE_SMONTHNAME4 = &H3B        '  long name for April
Public Const LOCALE_SMONTHNAME5 = &H3C        '  long name for May
Public Const LOCALE_SMONTHNAME6 = &H3D        '  long name for June
Public Const LOCALE_SMONTHNAME7 = &H3E        '  long name for July
Public Const LOCALE_SMONTHNAME8 = &H3F        '  long name for August
Public Const LOCALE_SMONTHNAME9 = &H40        '  long name for September
Public Const LOCALE_SMONTHNAME10 = &H41        '  long name for October
Public Const LOCALE_SMONTHNAME11 = &H42        '  long name for November
Public Const LOCALE_SMONTHNAME12 = &H43        '  long name for December
Public Const LOCALE_SABBREVMONTHNAME1 = &H44        '  abbreviated name for January
Public Const LOCALE_SABBREVMONTHNAME2 = &H45        '  abbreviated name for February
Public Const LOCALE_SABBREVMONTHNAME3 = &H46        '  abbreviated name for March
Public Const LOCALE_SABBREVMONTHNAME4 = &H47        '  abbreviated name for April
Public Const LOCALE_SABBREVMONTHNAME5 = &H48        '  abbreviated name for May
Public Const LOCALE_SABBREVMONTHNAME6 = &H49        '  abbreviated name for June
Public Const LOCALE_SABBREVMONTHNAME7 = &H4A        '  abbreviated name for July
Public Const LOCALE_SABBREVMONTHNAME8 = &H4B        '  abbreviated name for August
Public Const LOCALE_SABBREVMONTHNAME9 = &H4C        '  abbreviated name for September
Public Const LOCALE_SABBREVMONTHNAME10 = &H4D        '  abbreviated name for October
Public Const LOCALE_SABBREVMONTHNAME11 = &H4E        '  abbreviated name for November
Public Const LOCALE_SABBREVMONTHNAME12 = &H4F        '  abbreviated name for December
Public Const LOCALE_SABBREVMONTHNAME13 = &H100F

Public Const LOCALE_SPOSITIVESIGN = &H50        '  positive sign
Public Const LOCALE_SNEGATIVESIGN = &H51        '  negative sign
Public Const LOCALE_IPOSSIGNPOSN = &H52        '  positive sign position
Public Const LOCALE_INEGSIGNPOSN = &H53        '  negative sign position
Public Const LOCALE_IPOSSYMPRECEDES = &H54        '  mon sym precedes pos amt
Public Const LOCALE_IPOSSEPBYSPACE = &H55        '  mon sym sep by space from pos amt
Public Const LOCALE_INEGSYMPRECEDES = &H56        '  mon sym precedes neg amt
Public Const LOCALE_INEGSEPBYSPACE = &H57        '  mon sym sep by space from neg amt

' Time Flags for GetTimeFormatW.
Public Const TIME_NOMINUTESORSECONDS = &H1         '  do not use minutes or seconds
Public Const TIME_NOSECONDS = &H2         '  do not use seconds
Public Const TIME_NOTIMEMARKER = &H4         '  do not use time marker
Public Const TIME_FORCE24HOURFORMAT = &H8         '  always use 24 hour format

' Date Flags for GetDateFormatW.
Public Const DATE_SHORTDATE = &H1         '  use short date picture
Public Const DATE_LONGDATE = &H2         '  use long date picture



Private Sub aviso_erro(ByVal mensagem As String)

    Beep
    MsgBox mensagem, vbOKOnly + vbCritical + vbApplicationModal, "ERRO"

End Sub



Public Function restaura_configuracao_regional() As Boolean
' __________________________________________________________________________________________
'|
'| _ RESTAURA OS PARÂMETROS DA CONFIGURAÇÃO REGIONAL A PARTIR DOS VALORES
'|   ANTIGOS QUE ESTAVAM SALVOS NO ARQUIVO .INI
'| _ IMPORTANTE: AS ALTERAÇÕES NA CONFIGURAÇÃO REGIONAL NÃO FAZEM EFEITO P/
'|   OS PROGRAMAS QUE ESTÃO EM EXECUÇÃO, NEM MESMO P/ ESTE PROGRAMA QUE ESTÁ
'|   EFETUANDO OS COMANDOS DE ALTERAÇÃO.  É PRECISO QUE OS PROGRAMAS SEJAM
'|   INICIADOS DEPOIS QUE AS ALTERAÇÕES TENHAM SIDO FEITAS.
'|

Dim s_erro As String
Dim id_feature As Long
Dim s_param As String
Dim dwLCID As Long
Dim i As Integer
Dim r_reg() As TIPO_REGISTRY

    
    restaura_configuracao_regional = False
    

  ' PREPARA MENSAGEM DE ERRO
    s_erro = "Erro ao restaurar a configuração regional !" & _
             Chr$(13) & "Não é possível continuar !"
             
    
  ' INICIA ALTERAÇÃO DOS PARÂMETROS DA CONFIGURAÇÃO REGIONAL
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    dwLCID = GetUserDefaultLCID()
    
    If registry_le_chave(REG_CHAVE_PARAMETROS_ORIGINAIS, r_reg()) Then
        For i = LBound(r_reg) To UBound(r_reg)
            If Trim$(r_reg(i).campo) <> "" Then
                If IsNumeric(r_reg(i).campo) Then
                    id_feature = CLng(r_reg(i).campo)
                    s_param = r_reg(i).valor
                    GoSub RCR_ALTERA_PARAMETRO
                    End If
                End If
            Next
        
        End If
        
    
  ' FINALIZAÇÃO
  ' ~~~~~~~~~~~
    PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
    
    
    restaura_configuracao_regional = True
    
    
Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
RCR_ALTERA_PARAMETRO:
'~~~~~~~~~~~~~~~~~~~~
  ' ALTERA CONFIGURAÇÃO
    If SetLocaleInfo(dwLCID, id_feature, s_param) = False Then
        s_erro = s_erro & Chr$(13) & "Parâmetro: " & CStr(id_feature)
        aviso_erro s_erro
        Exit Function
        End If
    
    Return


End Function




Public Function verifica_configuracao_regional(Optional ByVal corrigir_configuracao As Boolean) As Boolean
' __________________________________________________________________________________________
'|
'| _ APENAS VERIFICA OU CONFIGURA OS PARÂMETROS DA CONFIGURAÇÃO REGIONAL.
'| _ ISSO É MUITO IMPORTANTE PORQUE AFETA FUNÇÕES COMO CDATE/CVDATE, CCUR, ETC.
'|   IMAGINE DUAS VARIÁVEIS STRING CONTENDO "10/01/2001" (10.JAN.2001) E
'|   "123.456,78".  SE A CONFIGURAÇÃO REGIONAL FOR "MM/DD/YY" P/ DATAS E O
'|   SEPARADOR DECIMAL FOR "." AO INVÉS DE ",", AS FUNÇÕES CDATE E CCUR IRÃO
'|   RETORNAR "01/10/2001" (01.OUT.2001) E 123,46
'| _ IMPORTANTE: AS ALTERAÇÕES NA CONFIGURAÇÃO REGIONAL NÃO FAZEM EFEITO P/
'|   OS PROGRAMAS QUE ESTÃO EM EXECUÇÃO, NEM MESMO P/ ESTE PROGRAMA QUE ESTÁ
'|   EFETUANDO OS COMANDOS DE ALTERAÇÃO.  É PRECISO QUE OS PROGRAMAS SEJAM
'|   INICIADOS DEPOIS QUE AS ALTERAÇÕES TENHAM SIDO FEITAS.
'| _ A ALTERAÇÃO NA CONFIGURAÇÃO REGIONAL CONTINUA MESMO SE FOR FEITO UM
'|   REBOOT. ENTRETANTO, SE UM OUTRO USUÁRIO FIZER O LOGON NA MÁQUINA, ENTÃO
'|   A CONFIGURAÇÃO REGIONAL SERÁ A DEFAULT.
'|

Const MAX_TAM_PARAM_V = 1000

Dim s_erro As String
Dim id_feature As Long
Dim s_param_v As String
Dim s_param_n As String
Dim dwLCID As Long
Dim s_erro_grava_reg As String

    
    verifica_configuracao_regional = False
    

  ' MENSAGEM DE ERRO: VAZIO = OK, SENÃO = ERRO
    s_erro_grava_reg = ""
    
       
  ' INICIA ALTERAÇÃO DOS PARÂMETROS DA CONFIGURAÇÃO REGIONAL
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    dwLCID = GetUserDefaultLCID()
    
    
  ' ***  D A T A  ***
  ' ~~~~~~~~~~~~~~~~~
 
  ' SEPARADOR DE DATA
    id_feature = LOCALE_SDATE
    s_param_n = "/"
    GoSub VCR_ALTERA_PARAMETRO
        
  ' FORMATO CURTO P/ DATA
    id_feature = LOCALE_SSHORTDATE
    s_param_n = "dd/MM/yyyy"
    GoSub VCR_ALTERA_PARAMETRO
    
  ' FORMATO LONGO P/ DATA
    id_feature = LOCALE_SLONGDATE
    If (dwLCID = &H416) Or (dwLCID = &H816) Then
      ' PORTUGUÊS BRASIL (0x416) OU PORTUGUÊS STANDARD (0x816)
        s_param_n = "dddd, d' de 'MMMM' de 'yyyy"
    Else
        s_param_n = "dddd', 'MMMM d', 'yyyy"
        End If
        
    GoSub VCR_ALTERA_PARAMETRO
    
    
    
  ' ***  M O E D A  ***
  ' ~~~~~~~~~~~~~~~~~~~
  
  ' SEPARADOR DECIMAL P/ MOEDA
    id_feature = LOCALE_SMONDECIMALSEP
    s_param_n = ","
    GoSub VCR_ALTERA_PARAMETRO
      
  ' SEPARADOR DE MILHAR P/ MOEDA
    id_feature = LOCALE_SMONTHOUSANDSEP
    s_param_n = "."
    GoSub VCR_ALTERA_PARAMETRO
        
  ' SÍMBOLO P/ MOEDA
    id_feature = LOCALE_SCURRENCY
    s_param_n = "R$"
    GoSub VCR_ALTERA_PARAMETRO
        
  ' QUANTIDADE DE DECIMAIS P/ MOEDA
    id_feature = LOCALE_ICURRDIGITS
    s_param_n = "2"
    GoSub VCR_ALTERA_PARAMETRO
        
  ' POSIÇÃO DO SÍMBOLO DE MOEDA
    id_feature = LOCALE_ICURRENCY
    s_param_n = "0"
    GoSub VCR_ALTERA_PARAMETRO
        
  ' FORMATO P/ MOEDA NEGATIVA
    id_feature = LOCALE_INEGCURR
    s_param_n = "0"
    GoSub VCR_ALTERA_PARAMETRO
        
  ' FORMATAÇÃO DE MOEDA (AGRUPAMENTO DE MILHAR)
    id_feature = LOCALE_SMONGROUPING
    s_param_n = "3;0"
    GoSub VCR_ALTERA_PARAMETRO
    
    
    
  ' ***  N Ú M E R O  ***
  ' ~~~~~~~~~~~~~~~~~~~~~
    
  ' SEPARADOR DECIMAL P/ NÚMERO
    id_feature = LOCALE_SDECIMAL
    s_param_n = ","
    GoSub VCR_ALTERA_PARAMETRO
        
  ' SEPARADOR DE MILHAR P/ NÚMERO
    id_feature = LOCALE_STHOUSAND
    s_param_n = "."
    GoSub VCR_ALTERA_PARAMETRO
        
  ' SÍMBOLO P/ NÚMERO NEGATIVO
    id_feature = LOCALE_SNEGATIVESIGN
    s_param_n = "-"
    GoSub VCR_ALTERA_PARAMETRO
        
  ' QUANTIDADE DE DECIMAIS P/ NÚMERO
    id_feature = LOCALE_IDIGITS
    s_param_n = "2"
    GoSub VCR_ALTERA_PARAMETRO
        
  ' FORMATAÇÃO DE NÚMERO (AGRUPAMENTO DE MILHAR)
    id_feature = LOCALE_SGROUPING
    s_param_n = "3;0"
    GoSub VCR_ALTERA_PARAMETRO
        
  ' ZERO À ESQUERDA P/ NÚMERO DECIMAL
    id_feature = LOCALE_ILZERO
    s_param_n = "1"
    GoSub VCR_ALTERA_PARAMETRO
    
    
        
  ' ***  H O R A  ***
  ' ~~~~~~~~~~~~~~~~~
  
  ' SEPARADOR P/ HORA
    id_feature = LOCALE_STIME
    s_param_n = ":"
    GoSub VCR_ALTERA_PARAMETRO
       
  ' FORMATO P/ HORA
    id_feature = LOCALE_STIMEFORMAT
  ' HH em maiúsculo significa 24-hour clock
    s_param_n = "HH:mm:ss"
    GoSub VCR_ALTERA_PARAMETRO
    
  ' SÍMBOLO P/ AM
    id_feature = LOCALE_S1159
    s_param_n = ""
    GoSub VCR_ALTERA_PARAMETRO
       
  ' SÍMBOLO P/ PM
    id_feature = LOCALE_S2359
    s_param_n = ""
    GoSub VCR_ALTERA_PARAMETRO
       
    
    
  ' ***  O U T R O S  ***
  ' ~~~~~~~~~~~~~~~~~~~~~
  
  ' SISTEMA MÉTRICO
    id_feature = LOCALE_IMEASURE
    s_param_n = "0"
    GoSub VCR_ALTERA_PARAMETRO
    
  ' SEPARADOR DE LISTAS
    id_feature = LOCALE_SLIST
    s_param_n = ";"
    GoSub VCR_ALTERA_PARAMETRO
    
    
    
  ' FINALIZAÇÃO
  ' ~~~~~~~~~~~
    If corrigir_configuracao Then
        PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
    
      ' SE OCORREU AO SALVAR OS PARÂMETROS ORIGINAIS, EMITE UM AVISO, MAS PROSSEGUE COM A EXECUÇÃO
        If s_erro_grava_reg <> "" Then
            s_erro_grava_reg = "Houve erro ao tentar salvar os seguintes parâmetros originais no registry " & _
                                Chr$(13) & s_erro_grava_reg
            aviso s_erro_grava_reg
            End If
        End If
        
        
    verifica_configuracao_regional = True
    
    
Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
VCR_ALTERA_PARAMETRO:
'~~~~~~~~~~~~~~~~~~~~
  ' OBTÉM A CONFIGURAÇÃO ATUAL
    s_param_v = Space$(MAX_TAM_PARAM_V)
    If GetLocaleInfo(dwLCID, id_feature, s_param_v, Len(s_param_v)) <> False Then
      
      ' ESTÁ DIFERENTE !!
        If s_param_n <> retorna_texto_sem_nulos(s_param_v) Then
            
            If corrigir_configuracao Then
              ' GRAVA NO REGISTRY O PARÂMETRO ORIGINAL QUE SERÁ ALTERADO POR UM NOVO
                If Not registry_grava_string(REG_CHAVE_PARAMETROS_ORIGINAIS, CStr(id_feature), retorna_texto_sem_nulos(s_param_v)) Then
                    If s_erro_grava_reg <> "" Then s_erro_grava_reg = s_erro_grava_reg & Chr$(13)
                    s_erro_grava_reg = s_erro_grava_reg & "Parâmetro: " & CStr(id_feature)
                    End If
            
              ' ALTERA CONFIGURAÇÃO
                If SetLocaleInfo(dwLCID, id_feature, s_param_n) = False Then
                  ' MENSAGEM DE ERRO
                    s_erro = "Erro ao alterar a configuração regional!" & _
                             Chr$(13) & "Não é possível continuar !" & _
                             Chr$(13) & "Parâmetro: " & CStr(id_feature)
                    aviso_erro s_erro
                    Exit Function
                    End If
            
            Else
                aviso_erro "Parâmetro: " & CStr(id_feature) & _
                           " da configuração regional está incorreto !" & _
                           Chr$(13) & "Não é possível continuar !"
                Exit Function
                End If
            
            End If
        End If
    
    
    Return


End Function



