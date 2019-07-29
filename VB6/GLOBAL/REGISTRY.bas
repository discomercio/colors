Attribute VB_Name = "mod_REGISTRY"
Option Explicit

' ___________________________________________________________________________________________________
'|
'|  OBSERVAÇÃO IMPORTANTE: QUANDO O CÓDIGO DE ERRO RETORNADO FOR "ERROR_MORE_DATA" (234),
'|  PROVAVELMENTE A CAUSA SEJA O TAMANHO INSUFICIENTE DE ALGUMA VARIÁVEL QUE FOI FORNECIDA
'|  PARA ACOMODAR OS DADOS DE RETORNO.
'|  POR EXEMPLO, A FUNÇÃO RegEnumKeyEx(), QUE ENUMERA AS SUB-CHAVES, ACEITA QUE OS PARÂMETROS
'|  "lpClass" E "lpcClass" SEJAM PASSADOS COMO NULOS NO WINDOWS 98, MAS OCORRE O ERRO
'|  ERROR_MORE_DATA NO WINDOWS NT.  OU SEJA, PARA O NT, TORNA-SE OBRIGATÓRIO FORNECER VARIÁVEIS
'|  DEVIDAMENTE DIMENSIONADAS PARA QUE ESSA FUNÇÃO POSSA RETORNAR OS DADOS.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Type TIPO_REGISTRY
    campo As String
    valor As String
    tipo_dado As Long
    End Type
    
    
Const MAX_KEY_LENGTH = 255

'   No more data is available.
Const ERROR_NO_MORE_ITEMS = 259&
'   The configuration registry key is invalid.
Const ERROR_BADKEY = 1010&

    

'-------------------------------------------------------------------------------
'   REGISTRY DO WINDOWS
'-------------------------------------------------------------------------------
' Acesso ao Registry
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003

Const ERROR_SUCCESS = 0&

Const DELETE = &H10000
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000

Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_READ = &H20000
Const STANDARD_RIGHTS_WRITE = &H20000
Const STANDARD_RIGHTS_EXECUTE = &H20000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_ALL = &H1F0000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
                  KEY_QUERY_VALUE Or _
                  KEY_ENUMERATE_SUB_KEYS Or _
                  KEY_NOTIFY) And _
                  (Not SYNCHRONIZE))
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

' Reg Create Type Values...
Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore

' Reg Data Types...
Const REG_NONE = 0                       ' No value type
' Constant for a string variable type.
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Const REG_BINARY = 3                     ' Free form binary
Const REG_DWORD = 4                      ' 32-bit number
Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Const REG_LINK = 6                       ' Symbolic Link (unicode)
Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Const REG_RESOURCE_LIST = 8              ' Resource list in the resource map
Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Const REG_CREATED_NEW_KEY = &H1                      ' New Registry Key created
Const REG_OPENED_EXISTING_KEY = &H2                      ' Existing Key opened
Const REG_WHOLE_HIVE_VOLATILE = &H1                      ' Restore whole hive volatile
Const REG_REFRESH_HIVE = &H2                      ' Unwind changes to last flush
Const REG_NOTIFY_CHANGE_NAME = &H1                      ' Create or delete (child)
Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Const REG_NOTIFY_CHANGE_LAST_SET = &H4                      ' Time stamp
Const REG_NOTIFY_CHANGE_SECURITY = &H8
Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
    End Type

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
    End Type

Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
    End Type

Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    Sacl As ACL
    Dacl As ACL
    End Type



' Registry API prototypes

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long



Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Function registry_grava_binario(ByVal chave As String, ByVal campo As String, ByVal valor As String, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  GRAVA DADOS DO TIPO BINÁRIO NO REGISTRY DO WINDOWS
'|
'|  SE JÁ EXISTE, ALTERA O VALOR (E O TIPO DE DADO, SE FOR O CASO).
'|  SE NÃO EXISTE, GRAVA O NOVO DADO.
'|
'|  CHAVE: é similar a um diretório, sendo que no RegEdit são as pastas que aparecem
'|         no lado esquerdo.
'|  CAMPO: é o nome do campo propriamente dito.
'|  VALOR: é o valor atribuído ao campo, que deve ser informado em formato HEXADECIMAL.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim n_disp As Long
Dim v() As Byte
Dim i As Long
Dim i_idx As Long
Dim s As String
Dim w_valor As String

    
    On Error GoTo REGISTRY_GRAVA_BINARIO_TRATA_ERRO
    
    
    registry_grava_binario = False
    
    msg_erro = ""
    
    
  ' CONVERTE O VALOR HEXADECIMAL EM STRING P/ VETOR DO TIPO BYTE
    w_valor = UCase$(Trim$(valor))
    If left$(w_valor, 2) = "&H" Then w_valor = right$(w_valor, Len(w_valor) - 2)
    If Len(w_valor) Mod 2 <> 0 Then w_valor = "0" & w_valor
    
    i_idx = 0
    ReDim v(1 To 1)
    For i = 1 To (Len(w_valor) \ 2)
        s = Mid$(w_valor, (i - 1) * 2 + 1, 2)
        i_idx = i_idx + 1
        ReDim Preserve v(1 To i_idx)
        v(i_idx) = converte_hex_para_dec(s)
        Next
        
    
  ' ABRE (ou CRIA) CHAVE NO REGISTRY
    r = RegCreateKeyEx(HKEY_LOCAL_MACHINE, chave, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, n_disp)
    If r <> ERROR_SUCCESS Then Exit Function
    
  ' ACERTA CAMPO DA CHAVE NO REGISTRY
    r = RegSetValueEx(hKey, campo, 0&, REG_BINARY, v(1), UBound(v))
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_GRAVA_BINARIO_FIM_ERRO

  ' FECHA CHAVE
    r = RegCloseKey(hKey)
    If r <> ERROR_SUCCESS Then Exit Function
    
    
  ' SOMENTE CONSIDERA GRAVADO SE CONSEGUIR FECHAR O HANDLE COM SUCESSO
    registry_grava_binario = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_GRAVA_BINARIO_TRATA_ERRO:
'=================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_GRAVA_BINARIO_FIM_ERRO
    
    Exit Function
    
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_GRAVA_BINARIO_FIM_ERRO:
'===============================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function



Function converte_dec_para_hex(ByVal numero As Byte) As String
' ____________________________________________________________________________________________________
'|
'|  CONVERTE UM NÚMERO DECIMAL PARA SUA FORMA HEXADECIMAL.
'|  O NÚMERO É PREENCHIDO C/ ZEROS À ESQUERDA, SE NECESSÁRIO.
'|

Dim s As String

    s = Hex(numero)
    While Len(s) < 2: s = "0" & s: Wend
        
    converte_dec_para_hex = s
    
End Function


Function converte_hex_para_dec(ByVal numero As String) As Byte
' ____________________________________________________________________________________________________
'|
'|  CONVERTE UM NÚMERO HEXADECIMAL PARA SUA FORMA DECIMAL.
'|

Dim s As String

    s = UCase$(Trim$(numero))
    If left$(s, 2) <> "&H" Then s = "&H" & s
    If IsNumeric(s) Then converte_hex_para_dec = CByte(s)
    
End Function


Function registry_grava_string(ByVal chave As String, ByVal campo As String, ByVal valor As String, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  GRAVA OS DADOS DO TIPO TEXTO NO REGISTRY DO WINDOWS
'|
'|  SE JÁ EXISTE, ALTERA O VALOR (E O TIPO DE DADO, SE FOR O CASO).
'|  SE NÃO EXISTE, GRAVA O NOVO DADO.
'|
'|  CHAVE: é similar a um diretório, sendo que no RegEdit são as pastas que aparecem
'|         no lado esquerdo.
'|  CAMPO: é o nome do campo propriamente dito.
'|  VALOR: é o valor atribuído ao campo.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim n_disp As Long

    
    On Error GoTo REGISTRY_GRAVA_STRING_TRATA_ERRO
    
    
    registry_grava_string = False
    
    msg_erro = ""
    
    
  ' ABRE (ou CRIA) CHAVE NO REGISTRY
    r = RegCreateKeyEx(HKEY_LOCAL_MACHINE, chave, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, n_disp)
    If r <> ERROR_SUCCESS Then Exit Function
    
  ' ACERTA CAMPO DA CHAVE NO REGISTRY
  ' STRINGS DEVEM SER PASSADAS COM BYVAL
    If Len(valor) = 0 Then valor = String(1, 0)
    r = RegSetValueEx(hKey, campo, 0&, REG_SZ, ByVal valor, Len(valor))
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_GRAVA_STRING_FIM_ERRO

  ' FECHA CHAVE
    r = RegCloseKey(hKey)
    If r <> ERROR_SUCCESS Then Exit Function
    
    
  ' SOMENTE CONSIDERA GRAVADO SE CONSEGUIR FECHAR O HANDLE COM SUCESSO
    registry_grava_string = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_GRAVA_STRING_TRATA_ERRO:
'================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_GRAVA_STRING_FIM_ERRO
    
    Exit Function
    
    



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_GRAVA_STRING_FIM_ERRO:
'==============================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    

End Function


Function registry_grava_numero(ByVal chave As String, ByVal campo As String, ByVal valor As Long, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  GRAVA DADOS DO TIPO NUMÉRICO NO REGISTRY DO WINDOWS
'|
'|  SE JÁ EXISTE, ALTERA O VALOR (E O TIPO DE DADO, SE FOR O CASO).
'|  SE NÃO EXISTE, GRAVA O NOVO DADO.
'|
'|  CHAVE: é similar a um diretório, sendo que no RegEdit são as pastas que aparecem
'|         no lado esquerdo.
'|  CAMPO: é o nome do campo propriamente dito.
'|  VALOR: é o valor atribuído ao campo.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim n_disp As Long

    
    On Error GoTo REGISTRY_GRAVA_NUMERO_TRATA_ERRO
    
    
    registry_grava_numero = False
    
    msg_erro = ""
    
    
  ' ABRE (ou CRIA) CHAVE NO REGISTRY
    r = RegCreateKeyEx(HKEY_LOCAL_MACHINE, chave, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, n_disp)
    If r <> ERROR_SUCCESS Then Exit Function
    
  ' ACERTA CAMPO DA CHAVE NO REGISTRY
    r = RegSetValueEx(hKey, campo, 0&, REG_DWORD, valor, Len(valor))
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_GRAVA_NUMERO_FIM_ERRO

  ' FECHA CHAVE
    r = RegCloseKey(hKey)
    If r <> ERROR_SUCCESS Then Exit Function
    
    
  ' SOMENTE CONSIDERA GRAVADO SE CONSEGUIR FECHAR O HANDLE COM SUCESSO
    registry_grava_numero = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_GRAVA_NUMERO_TRATA_ERRO:
'================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_GRAVA_NUMERO_FIM_ERRO
    
    Exit Function
    
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_GRAVA_NUMERO_FIM_ERRO:
'==============================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function

Function registry_le_string(ByVal chave As String, ByVal campo As String, ByRef valor As String, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  LÊ UM CAMPO DO TIPO TEXTO NO REGISTRY DO WINDOWS
'|  O VALOR É RETORNADO ATRAVÉS DO PARÂMETRO 'VALOR'.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim s_valor As String
Dim tam_s_valor As Long


    On Error GoTo REGISTRY_LE_STRING_TRATA_ERRO
    
    
    registry_le_string = False
    
    valor = ""
    msg_erro = ""
        

  ' ABRE CHAVE NO REGISTRY
    r = RegOpenKeyEx(HKEY_LOCAL_MACHINE, chave, 0&, KEY_READ, hKey)
    If r <> ERROR_SUCCESS Then Exit Function


  ' TENTA DETERMINAR O TAMANHO TOTAL DO CAMPO
    r = RegQueryValueExNULL(hKey, campo, 0&, REG_SZ, 0&, tam_s_valor)
    If r <> ERROR_SUCCESS Then
        msg_erro = "Falha ao tentar determinar o tamanho total do campo !"
        GoTo REGISTRY_LE_STRING_FIM_ERRO
        End If

  ' LÊ O CAMPO
    s_valor = String(tam_s_valor, 0)
    r = RegQueryValueExString(hKey, campo, 0&, REG_SZ, s_valor, tam_s_valor)
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_LE_STRING_FIM_ERRO
    
    valor = left$(s_valor, tam_s_valor - 1)
    
  ' JÁ LEU OS DADOS, ENTÃO RETORNA TRUE MESMO QUE FALHE AO TENTAR FECHAR O HANDLE
    registry_le_string = True
    
    
  ' FECHA CHAVE
    Call RegCloseKey(hKey)
                
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_LE_STRING_TRATA_ERRO:
'=============================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_LE_STRING_FIM_ERRO
    
    Exit Function
    
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_LE_STRING_FIM_ERRO:
'===========================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function

Function registry_le_binario(ByVal chave As String, ByVal campo As String, ByRef valor As String, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  LÊ UM CAMPO DO TIPO BINÁRIO NO REGISTRY DO WINDOWS
'|  O VALOR É RETORNADO ATRAVÉS DO PARÂMETRO 'VALOR', QUE É UMA STRING
'|  CONTENDO A REPRESENTAÇÃO HEXADECIMAL DO CONTEÚDO DO CAMPO.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim s_valor As String
Dim tam_s_valor As Long
Dim s As String
Dim i As Long


    On Error GoTo REGISTRY_LE_BINARIO_TRATA_ERRO
    
    
    registry_le_binario = False
    
    valor = ""
    msg_erro = ""
        

  ' ABRE CHAVE NO REGISTRY
    r = RegOpenKeyEx(HKEY_LOCAL_MACHINE, chave, 0&, KEY_READ, hKey)
    If r <> ERROR_SUCCESS Then Exit Function


  ' TENTA DETERMINAR O TAMANHO TOTAL DO CAMPO
    r = RegQueryValueExNULL(hKey, campo, 0&, REG_SZ, 0&, tam_s_valor)
    If r <> ERROR_SUCCESS Then
        msg_erro = "Falha ao tentar determinar o tamanho total do campo !"
        GoTo REGISTRY_LE_BINARIO_FIM_ERRO
        End If

  ' LÊ O CAMPO
    s_valor = String(tam_s_valor, 0)
    r = RegQueryValueExString(hKey, campo, 0&, REG_SZ, s_valor, tam_s_valor)
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_LE_BINARIO_FIM_ERRO
    
    s = left$(s_valor, tam_s_valor)
    For i = 1 To Len(s)
        valor = valor & converte_dec_para_hex(Asc(Mid$(s, i, 1)))
        Next
        
        
  ' JÁ LEU OS DADOS, ENTÃO RETORNA TRUE MESMO QUE FALHE AO TENTAR FECHAR O HANDLE
    registry_le_binario = True
    
    
  ' FECHA CHAVE
    Call RegCloseKey(hKey)
                
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_LE_BINARIO_TRATA_ERRO:
'==============================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_LE_BINARIO_FIM_ERRO
    
    Exit Function
    
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_LE_BINARIO_FIM_ERRO:
'============================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function


Function registry_le_numero(ByVal chave As String, ByVal campo As String, ByRef valor As Long, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  LÊ UM CAMPO DO TIPO NÚMERO NO REGISTRY DO WINDOWS
'|  O VALOR É RETORNADO ATRAVÉS DO PARÂMETRO 'VALOR'.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim n_valor As Long
Dim tam_n_valor As Long


    On Error GoTo REGISTRY_LE_NUMERO_TRATA_ERRO
    
    
    registry_le_numero = False
    
    valor = 0
    msg_erro = ""
        

  ' ABRE CHAVE NO REGISTRY
    r = RegOpenKeyEx(HKEY_LOCAL_MACHINE, chave, 0&, KEY_READ, hKey)
    If r <> ERROR_SUCCESS Then Exit Function

  ' LÊ O CAMPO
    tam_n_valor = Len(n_valor)
    r = RegQueryValueExLong(hKey, campo, 0&, REG_DWORD, n_valor, tam_n_valor)
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_LE_NUMERO_FIM_ERRO
    
    valor = n_valor
    
  ' JÁ LEU OS DADOS, ENTÃO RETORNA TRUE MESMO QUE FALHE AO TENTAR FECHAR O HANDLE
    registry_le_numero = True
    
    
  ' FECHA CHAVE
    Call RegCloseKey(hKey)
                
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_LE_NUMERO_TRATA_ERRO:
'=============================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_LE_NUMERO_FIM_ERRO
    
    Exit Function
    
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_LE_NUMERO_FIM_ERRO:
'===========================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function


Function registry_le_chave(ByVal chave As String, ByRef r_reg() As TIPO_REGISTRY, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  LÊ TODOS OS DADOS DE UMA CHAVE NO REGISTRY DO WINDOWS
'|  OS VALORES SÃO RETORNADOS NO VETOR PASSADO COMO PARÂMETRO.
'|
'|  Esta rotina está preparada para manipular somente os seguintes tipos de dados:
'|    a) Número inteiro (longo)
'|    b) Texto
'|
'|  CHAVE: é similar a um diretório, sendo que no RegEdit são as pastas que aparecem
'|         no lado esquerdo.
'|  CAMPO: é o nome do campo propriamente dito.
'|  VALOR: é o valor associado ao campo.
'|

Dim r As Long
Dim r_aux As Long
Dim hKey As Long
Dim i As Long
Dim i_pos As Long
Dim tam_campo As Long
Dim s_campo As String
Dim s_valor As String
Dim tam_s_valor As Long
Dim n_valor As Long
Dim tipo_dado As Long
Dim s_resp As String
Dim r_ok As Boolean

    
    On Error GoTo REGISTRY_LE_CHAVE_TRATA_ERRO
    
    
    registry_le_chave = False
    
    msg_erro = ""
    
    
    ReDim r_reg(0)
    With r_reg(0)
        .campo = ""
        .valor = ""
        .tipo_dado = 0
        End With
        
    
  ' ABRE CHAVE NO REGISTRY
    r = RegOpenKeyEx(HKEY_LOCAL_MACHINE, chave, 0&, KEY_READ, hKey)
    If r <> ERROR_SUCCESS Then Exit Function
    
    
    i_pos = 0
    Do
        tam_campo = 2000
        s_campo = String(tam_campo, 0)
        
        tam_s_valor = 2000
        s_valor = String(tam_s_valor, 0)
        
        r = RegEnumValue(hKey, i_pos, ByVal s_campo, tam_campo, 0&, tipo_dado, ByVal s_valor, tam_s_valor)
        
        If r = ERROR_SUCCESS Then
            s_resp = ""
            r_ok = False
            
          ' NOME DO CAMPO
            s_campo = left$(s_campo, tam_campo)
          
          ' VALOR É TIPO NÚMERO
            If tipo_dado = REG_DWORD Then
                If registry_le_numero(chave, s_campo, n_valor) Then
                    r_ok = True
                    s_resp = CStr(n_valor)
                    End If
          
          ' VALOR É DO TIPO BINÁRIO
            ElseIf tipo_dado = REG_BINARY Then
                If registry_le_binario(chave, s_campo, s_valor) Then
                    r_ok = True
                    s_resp = s_valor
                    End If
                
          ' VALOR É TIPO TEXTO
            Else
                If registry_le_string(chave, s_campo, s_valor) Then
                    r_ok = True
                    s_resp = s_valor
                    End If
                End If
                
                
          ' ADICIONA ESTE CAMPO AO VETOR DE RESPOSTA
            If r_ok Then
              ' PRECISA CRIAR NOVA ENTRADA NO VETOR ? (LEMBRE-SE DE QUE O CAMPO "DEFAULT" OU "PADRÃO" NÃO POSSUI NOME, MAS PODE RECEBER VALOR
                If (Trim$(r_reg(UBound(r_reg)).campo) <> "") Or (Trim$(r_reg(UBound(r_reg)).valor) <> "") Then ReDim Preserve r_reg(UBound(r_reg) + 1)
                With r_reg(UBound(r_reg))
                    .campo = left$(s_campo, tam_campo)
                    .valor = s_resp
                    .tipo_dado = tipo_dado
                    End With
                End If
            End If
        
        i_pos = i_pos + 1
        
        Loop While r = ERROR_SUCCESS
         
         
    
  ' JÁ LEU OS DADOS, ENTÃO RETORNA TRUE MESMO QUE FALHE AO TENTAR FECHAR O HANDLE
    registry_le_chave = True
    
         
  ' FECHA CHAVE
    Call RegCloseKey(hKey)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_LE_CHAVE_TRATA_ERRO:
'============================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_LE_CHAVE_FIM_ERRO
    
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_LE_CHAVE_FIM_ERRO:
'==========================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function



Function registry_obtem_subchaves(ByVal chave As String, ByRef r_reg() As String, Optional ByRef msg_erro As String, Optional ByVal inclui_chave_raiz_na_resposta As Boolean) As Boolean
' ________________________________________________________________________________________________
'|
'|  OBTÉM A RELAÇÃO DE SUB-CHAVES CONTIDAS NA CHAVE ESPECIFICADA
'|
'|  A LISTA DE SUB-CHAVES SERÁ DEVOLVIDA NO VETOR PASSADO POR PARÂMETRO.
'|
'|  O parâmetro 'inclui_chave_raiz_na_resposta' define se o parâmetro 'chave'
'|  também deve ser incluído na lista de sub-chaves da resposta.
'|  O default é não incluir, sendo que para incluir, deve-se obrigatoriamente
'|  fornecer o valor TRUE.
'|

Dim r As Long
    
    
    On Error GoTo REGISTRY_OBTEM_SUBCHAVES_TRATA_ERRO
    
    
    registry_obtem_subchaves = False
    
    msg_erro = ""
    
    
    ReDim r_reg(1 To 1)
    r_reg(1) = ""



  ' OBTÉM RELAÇÃO DE SUBCHAVES
    If inclui_chave_raiz_na_resposta Then
        r = RegGetSubKeysRecursively(HKEY_LOCAL_MACHINE, chave, r_reg(), chave)
    Else
        r = RegGetSubKeysRecursively(HKEY_LOCAL_MACHINE, chave, r_reg())
        End If
        
    If r <> ERROR_SUCCESS Then
        msg_erro = "Falha ao tentar obter a relação de sub-chaves da chave " & chave & " !" & _
                    vbCrLf & "(Código de erro = " & CStr(r) & ")"
        Exit Function
        End If


  ' ORDENA DE MODO CRESCENTE
    ORDENA_lista r_reg(), 1, UBound(r_reg)
    
    
    registry_obtem_subchaves = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_OBTEM_SUBCHAVES_TRATA_ERRO:
'===================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    Exit Function
    

End Function

Private Sub QuickSort_Lista(vetor() As String, ByVal inf As Integer, ByVal sup As Integer)
' _________________________________________________________________________________________
'|
'|  ALGORITMO DE ORDENAÇÃO QUICKSORT
'|  OBS: ALGORITMO É RECURSIVO
'|

Dim i As Integer
Dim j As Integer
Dim ref As String
Dim temp As String


    On Error GoTo QLISTA_ERRO




  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' LAÇO DE ORDENAÇÃO
    
    Do
        i = inf
        j = sup
        ref = vetor((inf + sup) \ 2)
        
        Do
            
            Do
                If ref > vetor(i) Then i = i + 1 Else Exit Do
                Loop

            
            
            Do
                If ref < vetor(j) Then j = j - 1 Else Exit Do
                Loop



            If i <= j Then
                temp = vetor(i)
                vetor(i) = vetor(j)
                vetor(j) = temp
                i = i + 1
                j = j - 1
                End If

            Loop Until i > j

        
        
        If inf < j Then QuickSort_Lista vetor(), inf, j
        
        inf = i
        
        Loop Until i >= sup




Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
QLISTA_ERRO:
'===========
    MsgBox CStr(Err) & ": " & Error$(Err), vbCritical
    Exit Sub




End Sub



Public Sub ORDENA_lista(vetor() As String, ByVal inf As Integer, ByVal sup As Integer)

    If inf > sup Then Exit Sub
    
    QuickSort_Lista vetor(), inf, sup

End Sub



Function RegGetSubKeysRecursively(ByVal hKey_root As Long, ByVal nome_chave As String, ByRef r_reg() As String, Optional ByVal nome_chave_completo As String) As Long
' _____________________________________________________________________________________________
'|
'|  OBTÉM A RELAÇÃO DE SUB-CHAVES DA CHAVE ESPECIFICADA
'|
'|  IMPORTANTE: ESTA FUNÇÃO É RECURSIVA !!
'|
'|  O parâmetro 'nome_chave_completo' define se o parâmetro 'nome_chave' será incluído ou não
'|  na lista de sub-chaves da resposta.  Ou seja, se a rotina chamadora quer que o parâmetro
'|  'nome_chave' seja incluído na resposta, então ela deve repetir o mesmo valor de 'nome_chave'
'|  para 'nome_chave_completo'.
'|  Se o parâmetro 'nome_chave_completo' não for passado ou se for passado em branco, então
'|  o parâmetro 'nome_chave' não será incluído na lista de resposta.
'|

Dim r As Long
Dim tam_s_subchave As Long
Dim s_subchave As String
Dim tam_s_classe As Long
Dim s_classe As String
Dim hKey As Long
Dim ft As FILETIME
Dim i_chave As Long
Dim s As String

    
    On Error GoTo REGGETSUBKEYSRECURSIVELY_TRATA_ERRO
    
        
  ' Do not allow NULL or empty key name
    If Trim$(nome_chave) = "" Then
        RegGetSubKeysRecursively = ERROR_BADKEY
        Exit Function
        End If
    
    
  ' ABRE A CHAVE NO REGISTRY
    r = RegOpenKeyEx(hKey_root, nome_chave, 0&, KEY_ALL_ACCESS, hKey)
    
    
  ' LAÇO QUE PROCURA PELAS SUBCHAVES
    i_chave = 0
    Do While r = ERROR_SUCCESS
    
        tam_s_subchave = MAX_KEY_LENGTH
        s_subchave = String(tam_s_subchave, 0)
        
        tam_s_classe = MAX_KEY_LENGTH
        s_classe = String(tam_s_classe, 0)
        
        r = RegEnumKeyEx(hKey, i_chave, s_subchave, tam_s_subchave, 0&, s_classe, tam_s_classe, ft)
        
        s_subchave = left$(s_subchave, tam_s_subchave)
        If r = ERROR_NO_MORE_ITEMS Then
          ' DESLIGA CÓDIGO DE ERRO ERROR_NO_MORE_ITEMS
            r = ERROR_SUCCESS
            
            If Trim$(nome_chave_completo) <> "" Then
                If Trim$(r_reg(UBound(r_reg))) <> "" Then ReDim Preserve r_reg(LBound(r_reg) To UBound(r_reg) + 1)
                r_reg(UBound(r_reg)) = nome_chave_completo
                End If
                
            Exit Do
        
        ElseIf r = ERROR_SUCCESS Then
          ' NAS CHAMADAS RECURSIVAS, É OBRIGATÓRIO FORNECER O PARÂMETRO PARA 'NOME_CHAVE_COMPLETO'
            If Trim$(nome_chave_completo) <> "" Then
                s = nome_chave_completo
            Else
                s = nome_chave
                End If
                
            r = RegGetSubKeysRecursively(hKey, s_subchave, r_reg(), barra_invertida_add(s) & s_subchave)
            End If
        
        i_chave = i_chave + 1
        Loop
    
    
    
  ' VALOR DE RETORNO
    RegGetSubKeysRecursively = r
    
    Call RegCloseKey(hKey)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGGETSUBKEYSRECURSIVELY_TRATA_ERRO:
'===================================
    RegGetSubKeysRecursively = Err
    
    Err.Clear
    
    GoTo REGGETSUBKEYSRECURSIVELY_FIM_ERRO
    
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGGETSUBKEYSRECURSIVELY_FIM_ERRO:
'=================================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function
 

Private Function barra_invertida_add(ByVal Texto As String) As String
' _________________________________________________________________________________________
'|
'|  ADICIONA A BARRA INVERTIDA NO FINAL
'|
'|  OBS: NÃO REMOVER ESPAÇOS EM BRANCO, POIS OS NOMES DAS CHAVES PODEM CONTER ESPAÇOS !!
'|

Dim s As String

    s = Texto
    If right$(RTrim$(s), 1) <> "\" Then s = s & "\"
    
    barra_invertida_add = s
    
End Function



Private Function barra_invertida_del(ByVal Texto As String) As String
' _________________________________________________________________________________________
'|
'|  REMOVE A BARRA INVERTIDA NO FINAL
'|
'|  OBS: NÃO REMOVER ESPAÇOS EM BRANCO, POIS OS NOMES DAS CHAVES PODEM CONTER ESPAÇOS !!
'|

Dim s As String

    s = Texto
    If right$(RTrim$(s), 1) = "\" Then s = left$(RTrim$(s), Len(RTrim$(s)) - 1)
    
    barra_invertida_del = s
    
End Function


Function registry_obtem_inf_campo(ByVal chave As String, ByVal campo As String, ByRef tipo_dado As Long, ByRef tamanho_dado As Long, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  OBTÉM O TIPO DE DADO E O TAMANHO DO CAMPO ESPECIFICADO
'|

Dim r As Long
Dim hKey As Long


    On Error GoTo REGISTRY_OBTEM_INF_CAMPO_TRATA_ERRO
    
    
    registry_obtem_inf_campo = False
    
    msg_erro = ""
    
    tipo_dado = REG_NONE
    tamanho_dado = 0
        

  ' ABRE CHAVE NO REGISTRY
    r = RegOpenKeyEx(HKEY_LOCAL_MACHINE, chave, 0&, KEY_READ, hKey)
    If r <> ERROR_SUCCESS Then Exit Function


  ' TENTA DETERMINAR O TIPO DE DADO DO CAMPO
    r = RegQueryValueExNULL(hKey, campo, 0&, tipo_dado, 0&, tamanho_dado)
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_OBTEM_INF_CAMPO_FIM_ERRO

    If tipo_dado = REG_SZ Then tamanho_dado = tamanho_dado - 1
    If tamanho_dado < 0 Then tamanho_dado = 0
    
  
  ' JÁ OBTEVE AS INFORMAÇÕES, ENTÃO RETORNA TRUE MESMO QUE FALHE AO TENTAR FECHAR O HANDLE
    registry_obtem_inf_campo = True
    
    
  ' FECHA CHAVE
    Call RegCloseKey(hKey)
                
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_OBTEM_INF_CAMPO_TRATA_ERRO:
'===================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_OBTEM_INF_CAMPO_FIM_ERRO
    
    Exit Function
    
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_OBTEM_INF_CAMPO_FIM_ERRO:
'=================================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function

Function registry_remove_campo(ByVal chave As String, ByVal campo As String, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  REMOVE UM CAMPO DA CHAVE NO REGISTRY DO WINDOWS
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
    
    
    On Error GoTo REGISTRY_REMOVE_CAMPO_TRATA_ERRO
    
    
    registry_remove_campo = False
    
    msg_erro = ""


  ' ABRE (ou CRIA) CHAVE NO REGISTRY
    r = RegOpenKeyEx(HKEY_LOCAL_MACHINE, chave, 0&, KEY_WRITE, hKey)
    If r <> ERROR_SUCCESS Then Exit Function


  ' REMOVE CAMPO DA CHAVE NO REGISTRY
    r = RegDeleteValue(hKey, campo)
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_REMOVE_CAMPO_FIM_ERRO


  ' FECHA CHAVE
    r = RegCloseKey(hKey)
    If r <> ERROR_SUCCESS Then Exit Function
    

  ' SOMENTE CONSIDERA REMOVIDO SE CONSEGUIR FECHAR O HANDLE COM SUCESSO
    registry_remove_campo = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_REMOVE_CAMPO_TRATA_ERRO:
'================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_REMOVE_CAMPO_FIM_ERRO
    
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_REMOVE_CAMPO_FIM_ERRO:
'==============================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    

End Function


Function registry_remove_chave(ByVal chave As String, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  REMOVE A CHAVE E SUAS SUBCHAVES NO REGISTRY DO WINDOWS
'|
'|  NO WINDOWS 95, A FUNÇÃO REGDELETEKEY() REMOVE AUTOMATICAMENTE
'|  AS SUBCHAVES QUE POSSAM EXISTIR.
'|  NO WINDOWS NT, ENTRETANTO, UMA CHAVE SOMENTE PODE SER REMOVIDA
'|  SE NÃO POSSUIR SUBCHAVES.
'|  PORTANTO, A SOLUÇÃO É IMPLEMENTAR UMA FUNÇÃO RECURSIVA, CONFORME
'|  SUGESTÃO DO ARTIGO Q142491, SENDO QUE ESTA ROTINA NÃO PREVINE CONTRA
'|  REMOÇÃO PARCIAL DE CHAVES.
'|

Dim r As Long
    
    
    On Error GoTo REGISTRY_REMOVE_CHAVE_TRATA_ERRO
    
    
    registry_remove_chave = False
    
    msg_erro = ""


  ' REMOVE CHAVE NO REGISTRY
    r = RegDeleteKeyRecursively(HKEY_LOCAL_MACHINE, chave)
    If r <> ERROR_SUCCESS Then
        msg_erro = "Falha ao tentar remover a chave " & chave & " no registry !" & _
                    vbCrLf & "(Código de erro = " & CStr(r) & ")"
        Exit Function
        End If


    registry_remove_chave = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_REMOVE_CHAVE_TRATA_ERRO:
'================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    Exit Function
    
    
End Function



Function RegDeleteKeyRecursively(ByVal hKey_root As Long, ByVal nome_chave As String) As Long
' _____________________________________________________________________________________________
'|
'|  REMOVE TODA A CHAVE (E SUAS SUBCHAVES) NO REGISTRY DO WINDOWS
'|
'|  IMPORTANTE: ESTA FUNÇÃO É RECURSIVA !!
'|
'|  NO WINDOWS 95, A FUNÇÃO REGDELETEKEY() REMOVE AUTOMATICAMENTE
'|  AS SUBCHAVES QUE POSSAM EXISTIR.
'|  NO WINDOWS NT, ENTRETANTO, UMA CHAVE SOMENTE PODE SER REMOVIDA
'|  SE NÃO POSSUIR SUBCHAVES.
'|  PORTANTO, A SOLUÇÃO É IMPLEMENTAR UMA FUNÇÃO RECURSIVA, CONFORME
'|  SUGESTÃO DO ARTIGO Q142491, SENDO QUE ESTA ROTINA NÃO PREVINE CONTRA
'|  REMOÇÃO PARCIAL DE CHAVES.
'|

Dim r As Long
Dim tam_s_subchave As Long
Dim s_subchave As String
Dim tam_s_classe As Long
Dim s_classe As String
Dim hKey As Long
Dim ft As FILETIME

    
    On Error GoTo REGDELETEKEYRECURSIVELY_TRATA_ERRO
    
        
  ' Do not allow NULL or empty key name
    If Trim$(nome_chave) = "" Then
        RegDeleteKeyRecursively = ERROR_BADKEY
        Exit Function
        End If
    
    
  ' ABRE A CHAVE NO REGISTRY
    r = RegOpenKeyEx(hKey_root, nome_chave, 0&, KEY_ALL_ACCESS Or DELETE, hKey)
    
    
  ' LAÇO QUE PROCURA PELAS SUBCHAVES
    Do While r = ERROR_SUCCESS
    
        tam_s_subchave = MAX_KEY_LENGTH
        s_subchave = String(tam_s_subchave, 0)
        
        tam_s_classe = MAX_KEY_LENGTH
        s_classe = String(tam_s_classe, 0)
        
      ' DEVE-SE USAR SEMPRE O ÍNDICE ZERO (0), JÁ QUE APÓS A REMOÇÃO, AS CHAVES
      ' SÃO REORGANIZADAS.
        r = RegEnumKeyEx(hKey, 0, s_subchave, tam_s_subchave, 0&, s_classe, tam_s_classe, ft)
        
        s_subchave = left$(s_subchave, tam_s_subchave)
        If r = ERROR_NO_MORE_ITEMS Then
            r = RegDeleteKey(hKey_root, nome_chave)
            Exit Do
        ElseIf r = ERROR_SUCCESS Then
            r = RegDeleteKeyRecursively(hKey, s_subchave)
            End If
        
        Loop
    
    
    
  ' VALOR DE RETORNO
    RegDeleteKeyRecursively = r
    
    Call RegCloseKey(hKey)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGDELETEKEYRECURSIVELY_TRATA_ERRO:
'==================================
    RegDeleteKeyRecursively = Err
    
    Err.Clear
    
    GoTo REGDELETEKEYRECURSIVELY_FIM_ERRO
    
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGDELETEKEYRECURSIVELY_FIM_ERRO:
'================================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function
 
' ________________________________________________________________________________________________
'|
'|  AS FUNÇÕES ABAIXO REPRODUZEM AS FUNÇÕES DE MANIPULAÇÃO DO REGISTRY CRIADAS ANTERIORMENTE,
'|  COM A DIFERENÇA EM ALTERAR O CONTEÚDO DE HKEY_CURRENT_USER (AS FUNÇÕES ACIMA ALTERAM O
'|  CONTEÚDO DE HKEY_LOCAL_MACHINE).
'|  DESTA FORMA, NADA DO QUE ESTAVA FUNCIONANDO NO SISTEMA REFERENTE AO REGISTRY TEVE QUE
'|  SER ALTERADO.


Function registry_usuario_grava_binario(ByVal chave As String, ByVal campo As String, ByVal valor As String, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  GRAVA DADOS DO TIPO BINÁRIO NO REGISTRY DO WINDOWS
'|
'|  SE JÁ EXISTE, ALTERA O VALOR (E O TIPO DE DADO, SE FOR O CASO).
'|  SE NÃO EXISTE, GRAVA O NOVO DADO.
'|
'|  CHAVE: é similar a um diretório, sendo que no RegEdit são as pastas que aparecem
'|         no lado esquerdo.
'|  CAMPO: é o nome do campo propriamente dito.
'|  VALOR: é o valor atribuído ao campo, que deve ser informado em formato HEXADECIMAL.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim n_disp As Long
Dim v() As Byte
Dim i As Long
Dim i_idx As Long
Dim s As String
Dim w_valor As String

    
    On Error GoTo REGISTRY_USUARIO_GRAVA_BINARIO_TRATA_ERRO
    
    
    registry_usuario_grava_binario = False
    
    msg_erro = ""
    
    
  ' CONVERTE O VALOR HEXADECIMAL EM STRING P/ VETOR DO TIPO BYTE
    w_valor = UCase$(Trim$(valor))
    If left$(w_valor, 2) = "&H" Then w_valor = right$(w_valor, Len(w_valor) - 2)
    If Len(w_valor) Mod 2 <> 0 Then w_valor = "0" & w_valor
    
    i_idx = 0
    ReDim v(1 To 1)
    For i = 1 To (Len(w_valor) \ 2)
        s = Mid$(w_valor, (i - 1) * 2 + 1, 2)
        i_idx = i_idx + 1
        ReDim Preserve v(1 To i_idx)
        v(i_idx) = converte_hex_para_dec(s)
        Next
        
    
  ' ABRE (ou CRIA) CHAVE NO REGISTRY
    r = RegCreateKeyEx(HKEY_CURRENT_USER, chave, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, n_disp)
    If r <> ERROR_SUCCESS Then Exit Function
    
  ' ACERTA CAMPO DA CHAVE NO REGISTRY
    r = RegSetValueEx(hKey, campo, 0&, REG_BINARY, v(1), UBound(v))
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_USUARIO_GRAVA_BINARIO_FIM_ERRO

  ' FECHA CHAVE
    r = RegCloseKey(hKey)
    If r <> ERROR_SUCCESS Then Exit Function
    
    
  ' SOMENTE CONSIDERA GRAVADO SE CONSEGUIR FECHAR O HANDLE COM SUCESSO
    registry_usuario_grava_binario = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_GRAVA_BINARIO_TRATA_ERRO:
'=========================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_USUARIO_GRAVA_BINARIO_FIM_ERRO
    
    Exit Function
    
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_GRAVA_BINARIO_FIM_ERRO:
'=======================================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function


Function registry_usuario_grava_string(ByVal chave As String, ByVal campo As String, ByVal valor As String, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  GRAVA OS DADOS DO TIPO TEXTO NO REGISTRY DO WINDOWS
'|
'|  SE JÁ EXISTE, ALTERA O VALOR (E O TIPO DE DADO, SE FOR O CASO).
'|  SE NÃO EXISTE, GRAVA O NOVO DADO.
'|
'|  CHAVE: é similar a um diretório, sendo que no RegEdit são as pastas que aparecem
'|         no lado esquerdo.
'|  CAMPO: é o nome do campo propriamente dito.
'|  VALOR: é o valor atribuído ao campo.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim n_disp As Long

    
    On Error GoTo REGISTRY_USUARIO_GRAVA_STRING_TRATA_ERRO
    
    
    registry_usuario_grava_string = False
    
    msg_erro = ""
    
    
  ' ABRE (ou CRIA) CHAVE NO REGISTRY
    r = RegCreateKeyEx(HKEY_CURRENT_USER, chave, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, n_disp)
    If r <> ERROR_SUCCESS Then Exit Function
    
  ' ACERTA CAMPO DA CHAVE NO REGISTRY
  ' STRINGS DEVEM SER PASSADAS COM BYVAL
    If Len(valor) = 0 Then valor = String(1, 0)
    r = RegSetValueEx(hKey, campo, 0&, REG_SZ, ByVal valor, Len(valor))
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_USUARIO_GRAVA_STRING_FIM_ERRO

  ' FECHA CHAVE
    r = RegCloseKey(hKey)
    If r <> ERROR_SUCCESS Then Exit Function
    
    
  ' SOMENTE CONSIDERA GRAVADO SE CONSEGUIR FECHAR O HANDLE COM SUCESSO
    registry_usuario_grava_string = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_GRAVA_STRING_TRATA_ERRO:
'========================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_USUARIO_GRAVA_STRING_FIM_ERRO
    
    Exit Function
    
    



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_GRAVA_STRING_FIM_ERRO:
'======================================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    

End Function


Function registry_usuario_grava_numero(ByVal chave As String, ByVal campo As String, ByVal valor As Long, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  GRAVA DADOS DO TIPO NUMÉRICO NO REGISTRY DO WINDOWS
'|
'|  SE JÁ EXISTE, ALTERA O VALOR (E O TIPO DE DADO, SE FOR O CASO).
'|  SE NÃO EXISTE, GRAVA O NOVO DADO.
'|
'|  CHAVE: é similar a um diretório, sendo que no RegEdit são as pastas que aparecem
'|         no lado esquerdo.
'|  CAMPO: é o nome do campo propriamente dito.
'|  VALOR: é o valor atribuído ao campo.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim n_disp As Long

    
    On Error GoTo REGISTRY_USUARIO_GRAVA_NUMERO_TRATA_ERRO
    
    
    registry_usuario_grava_numero = False
    
    msg_erro = ""
    
    
  ' ABRE (ou CRIA) CHAVE NO REGISTRY
    r = RegCreateKeyEx(HKEY_CURRENT_USER, chave, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, n_disp)
    If r <> ERROR_SUCCESS Then Exit Function
    
  ' ACERTA CAMPO DA CHAVE NO REGISTRY
    r = RegSetValueEx(hKey, campo, 0&, REG_DWORD, valor, Len(valor))
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_USUARIO_GRAVA_NUMERO_FIM_ERRO

  ' FECHA CHAVE
    r = RegCloseKey(hKey)
    If r <> ERROR_SUCCESS Then Exit Function
    
    
  ' SOMENTE CONSIDERA GRAVADO SE CONSEGUIR FECHAR O HANDLE COM SUCESSO
    registry_usuario_grava_numero = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_GRAVA_NUMERO_TRATA_ERRO:
'========================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_USUARIO_GRAVA_NUMERO_FIM_ERRO
    
    Exit Function
    
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_GRAVA_NUMERO_FIM_ERRO:
'======================================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function

Function registry_usuario_le_string(ByVal chave As String, ByVal campo As String, ByRef valor As String, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  LÊ UM CAMPO DO TIPO TEXTO NO REGISTRY DO WINDOWS
'|  O VALOR É RETORNADO ATRAVÉS DO PARÂMETRO 'VALOR'.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim s_valor As String
Dim tam_s_valor As Long


    On Error GoTo REGISTRY_USUARIO_LE_STRING_TRATA_ERRO
    
    
    registry_usuario_le_string = False
    
    valor = ""
    msg_erro = ""
        

  ' ABRE CHAVE NO REGISTRY
    r = RegOpenKeyEx(HKEY_CURRENT_USER, chave, 0&, KEY_READ, hKey)
    If r <> ERROR_SUCCESS Then Exit Function


  ' TENTA DETERMINAR O TAMANHO TOTAL DO CAMPO
    r = RegQueryValueExNULL(hKey, campo, 0&, REG_SZ, 0&, tam_s_valor)
    If r <> ERROR_SUCCESS Then
        msg_erro = "Falha ao tentar determinar o tamanho total do campo !"
        GoTo REGISTRY_USUARIO_LE_STRING_FIM_ERRO
        End If

  ' LÊ O CAMPO
    s_valor = String(tam_s_valor, 0)
    r = RegQueryValueExString(hKey, campo, 0&, REG_SZ, s_valor, tam_s_valor)
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_USUARIO_LE_STRING_FIM_ERRO
    
    valor = left$(s_valor, tam_s_valor - 1)
    
  ' JÁ LEU OS DADOS, ENTÃO RETORNA TRUE MESMO QUE FALHE AO TENTAR FECHAR O HANDLE
    registry_usuario_le_string = True
    
    
  ' FECHA CHAVE
    Call RegCloseKey(hKey)
                
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_LE_STRING_TRATA_ERRO:
'=====================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_USUARIO_LE_STRING_FIM_ERRO
    
    Exit Function
    
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_LE_STRING_FIM_ERRO:
'===================================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function

Function registry_usuario_le_binario(ByVal chave As String, ByVal campo As String, ByRef valor As String, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  LÊ UM CAMPO DO TIPO BINÁRIO NO REGISTRY DO WINDOWS
'|  O VALOR É RETORNADO ATRAVÉS DO PARÂMETRO 'VALOR', QUE É UMA STRING
'|  CONTENDO A REPRESENTAÇÃO HEXADECIMAL DO CONTEÚDO DO CAMPO.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim s_valor As String
Dim tam_s_valor As Long
Dim s As String
Dim i As Long


    On Error GoTo REGISTRY_USUARIO_LE_BINARIO_TRATA_ERRO
    
    
    registry_usuario_le_binario = False
    
    valor = ""
    msg_erro = ""
        

  ' ABRE CHAVE NO REGISTRY
    r = RegOpenKeyEx(HKEY_CURRENT_USER, chave, 0&, KEY_READ, hKey)
    If r <> ERROR_SUCCESS Then Exit Function


  ' TENTA DETERMINAR O TAMANHO TOTAL DO CAMPO
    r = RegQueryValueExNULL(hKey, campo, 0&, REG_SZ, 0&, tam_s_valor)
    If r <> ERROR_SUCCESS Then
        msg_erro = "Falha ao tentar determinar o tamanho total do campo !"
        GoTo REGISTRY_USUARIO_LE_BINARIO_FIM_ERRO
        End If

  ' LÊ O CAMPO
    s_valor = String(tam_s_valor, 0)
    r = RegQueryValueExString(hKey, campo, 0&, REG_SZ, s_valor, tam_s_valor)
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_USUARIO_LE_BINARIO_FIM_ERRO
    
    s = left$(s_valor, tam_s_valor)
    For i = 1 To Len(s)
        valor = valor & converte_dec_para_hex(Asc(Mid$(s, i, 1)))
        Next
        
        
  ' JÁ LEU OS DADOS, ENTÃO RETORNA TRUE MESMO QUE FALHE AO TENTAR FECHAR O HANDLE
    registry_usuario_le_binario = True
    
    
  ' FECHA CHAVE
    Call RegCloseKey(hKey)
                
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_LE_BINARIO_TRATA_ERRO:
'======================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_USUARIO_LE_BINARIO_FIM_ERRO
    
    Exit Function
    
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_LE_BINARIO_FIM_ERRO:
'====================================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function


Function registry_usuario_le_numero(ByVal chave As String, ByVal campo As String, ByRef valor As Long, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  LÊ UM CAMPO DO TIPO NÚMERO NO REGISTRY DO WINDOWS
'|  O VALOR É RETORNADO ATRAVÉS DO PARÂMETRO 'VALOR'.
'|
'|  SE AO LER, GRAVAR, ALTERAR OU APAGAR UM CAMPO, FOR PASSADO UMA STRING VAZIA COMO
'|  NOME DO CAMPO, O WINDOWS IRÁ AUTOMATICAMENTE ASSUMIR O CAMPO (DEFAULT) OU (PADRÃO)
'|  PARA REALIZAR A OPERAÇÃO.
'|

Dim r As Long
Dim hKey As Long
Dim n_valor As Long
Dim tam_n_valor As Long


    On Error GoTo REGISTRY_USUARIO_LE_NUMERO_TRATA_ERRO
    
    
    registry_usuario_le_numero = False
    
    valor = 0
    msg_erro = ""
        

  ' ABRE CHAVE NO REGISTRY
    r = RegOpenKeyEx(HKEY_CURRENT_USER, chave, 0&, KEY_READ, hKey)
    If r <> ERROR_SUCCESS Then Exit Function

  ' LÊ O CAMPO
    tam_n_valor = Len(n_valor)
    r = RegQueryValueExLong(hKey, campo, 0&, REG_DWORD, n_valor, tam_n_valor)
    If r <> ERROR_SUCCESS Then GoTo REGISTRY_USUARIO_LE_NUMERO_FIM_ERRO
    
    valor = n_valor
    
  ' JÁ LEU OS DADOS, ENTÃO RETORNA TRUE MESMO QUE FALHE AO TENTAR FECHAR O HANDLE
    registry_usuario_le_numero = True
    
    
  ' FECHA CHAVE
    Call RegCloseKey(hKey)
                
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_LE_NUMERO_TRATA_ERRO:
'=====================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_USUARIO_LE_NUMERO_FIM_ERRO
    
    Exit Function
    
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_LE_NUMERO_FIM_ERRO:
'===================================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function


Function registry_usuario_le_chave(ByVal chave As String, ByRef r_reg() As TIPO_REGISTRY, Optional ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________
'|
'|  LÊ TODOS OS DADOS DE UMA CHAVE NO REGISTRY DO WINDOWS
'|  OS VALORES SÃO RETORNADOS NO VETOR PASSADO COMO PARÂMETRO.
'|
'|  Esta rotina está preparada para manipular somente os seguintes tipos de dados:
'|    a) Número inteiro (longo)
'|    b) Texto
'|
'|  CHAVE: é similar a um diretório, sendo que no RegEdit são as pastas que aparecem
'|         no lado esquerdo.
'|  CAMPO: é o nome do campo propriamente dito.
'|  VALOR: é o valor associado ao campo.
'|

Dim r As Long
Dim r_aux As Long
Dim hKey As Long
Dim i As Long
Dim i_pos As Long
Dim tam_campo As Long
Dim s_campo As String
Dim s_valor As String
Dim tam_s_valor As Long
Dim n_valor As Long
Dim tipo_dado As Long
Dim s_resp As String
Dim r_ok As Boolean

    
    On Error GoTo REGISTRY_USUARIO_LE_CHAVE_TRATA_ERRO
    
    
    registry_usuario_le_chave = False
    
    msg_erro = ""
    
    
    ReDim r_reg(0)
    With r_reg(0)
        .campo = ""
        .valor = ""
        .tipo_dado = 0
        End With
        
    
  ' ABRE CHAVE NO REGISTRY
    r = RegOpenKeyEx(HKEY_CURRENT_USER, chave, 0&, KEY_READ, hKey)
    If r <> ERROR_SUCCESS Then Exit Function
    
    
    i_pos = 0
    Do
        tam_campo = 2000
        s_campo = String(tam_campo, 0)
        
        tam_s_valor = 2000
        s_valor = String(tam_s_valor, 0)
        
        r = RegEnumValue(hKey, i_pos, ByVal s_campo, tam_campo, 0&, tipo_dado, ByVal s_valor, tam_s_valor)
        
        If r = ERROR_SUCCESS Then
            s_resp = ""
            r_ok = False
            
          ' NOME DO CAMPO
            s_campo = left$(s_campo, tam_campo)
          
          ' VALOR É TIPO NÚMERO
            If tipo_dado = REG_DWORD Then
                If registry_usuario_le_numero(chave, s_campo, n_valor) Then
                    r_ok = True
                    s_resp = CStr(n_valor)
                    End If
          
          ' VALOR É DO TIPO BINÁRIO
            ElseIf tipo_dado = REG_BINARY Then
                If registry_usuario_le_binario(chave, s_campo, s_valor) Then
                    r_ok = True
                    s_resp = s_valor
                    End If
                
          ' VALOR É TIPO TEXTO
            Else
                If registry_usuario_le_string(chave, s_campo, s_valor) Then
                    r_ok = True
                    s_resp = s_valor
                    End If
                End If
                
                
          ' ADICIONA ESTE CAMPO AO VETOR DE RESPOSTA
            If r_ok Then
              ' PRECISA CRIAR NOVA ENTRADA NO VETOR ? (LEMBRE-SE DE QUE O CAMPO "DEFAULT" OU "PADRÃO" NÃO POSSUI NOME, MAS PODE RECEBER VALOR
                If (Trim$(r_reg(UBound(r_reg)).campo) <> "") Or (Trim$(r_reg(UBound(r_reg)).valor) <> "") Then ReDim Preserve r_reg(UBound(r_reg) + 1)
                With r_reg(UBound(r_reg))
                    .campo = left$(s_campo, tam_campo)
                    .valor = s_resp
                    .tipo_dado = tipo_dado
                    End With
                End If
            End If
        
        i_pos = i_pos + 1
        
        Loop While r = ERROR_SUCCESS
         
         
    
  ' JÁ LEU OS DADOS, ENTÃO RETORNA TRUE MESMO QUE FALHE AO TENTAR FECHAR O HANDLE
    registry_usuario_le_chave = True
    
         
  ' FECHA CHAVE
    Call RegCloseKey(hKey)
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_LE_CHAVE_TRATA_ERRO:
'====================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoTo REGISTRY_USUARIO_LE_CHAVE_FIM_ERRO
    
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_LE_CHAVE_FIM_ERRO:
'==================================
  ' TENTA SE ASSEGURAR DE O HANDLE SERÁ FECHADO SEMPRE
    On Error Resume Next
    Call RegCloseKey(hKey)
    
    Exit Function
    
    
End Function



Function registry_usuario_obtem_subchaves(ByVal chave As String, ByRef r_reg() As String, Optional ByRef msg_erro As String, Optional ByVal inclui_chave_raiz_na_resposta As Boolean) As Boolean
' ________________________________________________________________________________________________
'|
'|  OBTÉM A RELAÇÃO DE SUB-CHAVES CONTIDAS NA CHAVE ESPECIFICADA
'|
'|  A LISTA DE SUB-CHAVES SERÁ DEVOLVIDA NO VETOR PASSADO POR PARÂMETRO.
'|
'|  O parâmetro 'inclui_chave_raiz_na_resposta' define se o parâmetro 'chave'
'|  também deve ser incluído na lista de sub-chaves da resposta.
'|  O default é não incluir, sendo que para incluir, deve-se obrigatoriamente
'|  fornecer o valor TRUE.
'|

Dim r As Long
    
    
    On Error GoTo REGISTRY_USUARIO_OBTEM_SUBCHAVES_TRATA_ERRO
    
    
    registry_usuario_obtem_subchaves = False
    
    msg_erro = ""
    
    
    ReDim r_reg(1 To 1)
    r_reg(1) = ""



  ' OBTÉM RELAÇÃO DE SUBCHAVES
    If inclui_chave_raiz_na_resposta Then
        r = RegGetSubKeysRecursively(HKEY_CURRENT_USER, chave, r_reg(), chave)
    Else
        r = RegGetSubKeysRecursively(HKEY_CURRENT_USER, chave, r_reg())
        End If
        
    If r <> ERROR_SUCCESS Then
        msg_erro = "Falha ao tentar obter a relação de sub-chaves da chave " & chave & " !" & _
                    vbCrLf & "(Código de erro = " & CStr(r) & ")"
        Exit Function
        End If


  ' ORDENA DE MODO CRESCENTE
    ORDENA_lista r_reg(), 1, UBound(r_reg)
    
    
    registry_usuario_obtem_subchaves = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REGISTRY_USUARIO_OBTEM_SUBCHAVES_TRATA_ERRO:
'===========================================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    Exit Function
    

End Function



