Attribute VB_Name = "mod_FILE"
Option Explicit

' BUG: A FUNÇÃO DO VB FileCopy() ÀS VEZES FALHA QUANDO É USADA NO WINDOWS NT
' P/ COPIAR UM ARQUIVO P/ UMA UNIDADE DE REDE MAPEADA OU UM UNC COM O ERRO:
'    75: Path/File access error
' PORTANTO, DÊ PREFERÊNCIA AO CopyFile() DA API DO WINDOWS !!

' CopyFile
'    bFailIfExists [in] Specifies how this operation is to proceed if a file of the same name as that specified by lpNewFileName already exists. If this parameter is TRUE and the new file already exists, the function fails. If this parameter is FALSE and the new file already exists, the function overwrites the existing file and succeeds.
' Return Values
'    If the function succeeds, the return value is nonzero.
'    If the function fails, the return value is zero. To get extended error information, call GetLastError.
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long



Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long



' DECLARAÇÕES PARA SHELLEXECUTE()
Private Const SW_SHOWNORMAL = 1

Const SE_ERR_FNF = 2&
Const SE_ERR_PNF = 3&
Const SE_ERR_ACCESSDENIED = 5&
Const SE_ERR_OOM = 8&
Const SE_ERR_DLLNOTFOUND = 32&
Const SE_ERR_SHARE = 26&
Const SE_ERR_ASSOCINCOMPLETE = 27&
Const SE_ERR_DDETIMEOUT = 28&
Const SE_ERR_DDEFAIL = 29&
Const SE_ERR_DDEBUSY = 30&
Const SE_ERR_NOASSOC = 31&
Const ERROR_BAD_FORMAT = 11&

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long


Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800

Function FileObtemAtributos(ByVal nome_arquivo As String, ByRef atributos As Long, ByRef msg_erro As String) As Boolean
' _______________________________________________________________________________________________
'|
'|  OBTÉM OS ATRIBUTOS DO ARQUIVO ESPECIFICADO.
'|
'|  VALORES DE RETORNO:
'|    'TRUE': OS ATRIBUTOS DO ARQUIVO FORAM LIDOS CORRETAMENTE E SERÃO DEVOLVIDOS
'|            NO PARÂMETRO 'atributos'.
'|    'FALSE': 1) O ARQUIVO NÃO EXISTE !!
'|             2) HOUVE ERRO E A DESCRIÇÃO DO ERRO SERÁ RETORNADA EM 'msg_erro'.
'|

Dim i_ret As Long
    
    
    On Error GoTo FILEOBTEMATTR_TRATA_ERRO
    
    
    FileObtemAtributos = False
    
    msg_erro = ""
    atributos = -1
    
    i_ret = GetFileAttributes(nome_arquivo)
    If i_ret <> -1 Then
        atributos = i_ret
        FileObtemAtributos = True
        End If


Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FILEOBTEMATTR_TRATA_ERRO:
'========================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    Exit Function
    
    
End Function

Function start_doc(ByVal nome_arquivo As String, ByRef msg_erro As String) As Boolean
' ________________________________________________________________________________________________________
'|
'|  EXIBE O CONTEÚDO DO ARQUIVO ATRAVÉS DA CHAMADA À FUNÇÃO DA API
'|  SHELLEXECUTE(), QUE SE ENCARREGARÁ DE ATIVAR O APLICATIVO ASSOCIADO
'|  AO TIPO DE ARQUIVO.
'|

Dim s As String
Dim s_arq As String
Dim r As Variant
Dim scr_hDC As Long


    On Error GoTo STARTDOC_TRATA_ERRO
    
    
    start_doc = False
    
    msg_erro = ""
    
    
    If Trim$(nome_arquivo) = "" Then
        msg_erro = "Nome do arquivo não foi fornecido"
        Exit Function
        End If
        
        
  ' TENTA OBTER O NOME CURTO, SE NÃO CONSEGUIR, USA O NOME LONGO
    s_arq = GetShortName(nome_arquivo, s)
    If Trim$(s_arq) = "" Then s_arq = Trim$(nome_arquivo)
    
    scr_hDC = GetDesktopWindow()
    r = ShellExecute(scr_hDC, "Open", s_arq, "", ExtractFilePath(s_arq), SW_SHOWNORMAL)
    
    
    
  ' O K  ! !
  ' ~~~~~~~~
    If r > 32 Then
        start_doc = True
        Exit Function
        End If
    
    
    
  ' HOUVE ERRO !!
  ' ~~~~~~~~~~~~~
    Select Case r
        Case SE_ERR_FNF
            msg_erro = "Arquivo NÃO encontrado: " & s_arq
        Case SE_ERR_PNF
            msg_erro = "Caminho NÃO encontrado"
        Case SE_ERR_ACCESSDENIED
            msg_erro = "Acesso negado"
        Case SE_ERR_OOM
            msg_erro = "Sem memória"
        Case SE_ERR_DLLNOTFOUND
            msg_erro = "DLL não encontrada"
        Case SE_ERR_SHARE
            msg_erro = "Ocorreu uma violação de compartilhamento"
        Case SE_ERR_ASSOCINCOMPLETE
            msg_erro = "Associação de tipo de arquivo inválida ou incompleta"
        Case SE_ERR_DDETIMEOUT
            msg_erro = "DDE: tempo limite excedido"
        Case SE_ERR_DDEFAIL
            msg_erro = "DDE: falha na transação"
        Case SE_ERR_DDEBUSY
            msg_erro = "DDE: ocupado"
        Case SE_ERR_NOASSOC
            msg_erro = "Não há aplicativo associado a este tipo de arquivo"
        Case ERROR_BAD_FORMAT
            msg_erro = "Arquivo EXE inválido ou erro no arquivo"
        Case Else
            msg_erro = "Erro desconhecido"
        End Select
   
              
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
STARTDOC_TRATA_ERRO:
'===================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    
    Exit Function


End Function

Function ExtractFileExt(ByVal nome_arquivo As String) As String
' ___________________________________________________________________________________________________________________
'|
'|  RETORNA SOMENTE A EXTENSÃO DO NOME DO ARQUIVO, SE HOUVER.
'|

Dim s_resp As String
Dim s As String
Dim i As Integer
Dim achou_ponto As Boolean

    
    ExtractFileExt = ""
    
    achou_ponto = False
    
    s_resp = ""
    For i = Len(nome_arquivo) To 1 Step -1
        s = Mid$(nome_arquivo, i, 1)
        If (s <> "\") And (s <> ".") Then
            s_resp = s & s_resp
        Else
            If s = "." Then achou_ponto = True
            Exit For
            End If
        Next
        
        
  ' NOME DE ARQUIVO NÃO POSSUI EXTENSÃO !!
    If Not achou_ponto Then s_resp = ""
    
    
    ExtractFileExt = s_resp
    
    
End Function

Function ExtractFileName(ByVal nome_arquivo As String, Optional ByVal remover_extensao As Boolean) As String
' ___________________________________________________________________________________________________________________
'|
'|  RETORNA SOMENTE O NOME DO ARQUIVO, SEM PATH
'|

Dim s_resp As String
Dim s As String
Dim i As Integer

    
    ExtractFileName = ""
    
    s_resp = ""
    For i = Len(nome_arquivo) To 1 Step -1
        s = Mid$(nome_arquivo, i, 1)
        If s <> "\" Then
            s_resp = s & s_resp
        Else
            Exit For
            End If
        Next
        
        
  ' REMOVE A EXTENSÃO TAMBÉM ?
    If remover_extensao And (InStr(s_resp, ".") <> 0) Then
        Do While Len(s_resp) > 0
            s = right$(s_resp, 1)
            s_resp = left$(s_resp, Len(s_resp) - 1)
            If s = "." Then Exit Do
            Loop
        End If
        
    
    ExtractFileName = s_resp
    
    
End Function


Function ExtractFileDrive(ByVal nome_arquivo As String) As String
' ___________________________________________________________________________________________________________________
'|
'|  RETORNA SOMENTE A LETRA DA UNIDADE
'|

    ExtractFileDrive = ""
    
    If Mid$(nome_arquivo, 2, 1) <> ":" Then Exit Function
        
    ExtractFileDrive = left$(nome_arquivo, 2)
        
End Function



Function ExtractFilePath(ByVal nome_arquivo As String) As String
' ___________________________________________________________________________________________________________________
'|
'|  RETORNA O PATH DO ARQUIVO, SEM O SEU NOME
'|

Dim s_resp As String
Dim s As String
Dim i As Integer

    ExtractFilePath = ""
    
    s_resp = ""
    For i = Len(nome_arquivo) To 1 Step -1
        s = Mid$(nome_arquivo, i, 1)
        If s = "\" Then
            s_resp = Mid$(nome_arquivo, 1, i - 1)
            Exit For
            End If
        Next
        
    
    ExtractFilePath = s_resp
    
    
End Function



Function barra_invertida_add(ByVal Texto As String) As String
' _________________________________________________________________________________________
'|
'|  ADICIONA A BARRA INVERTIDA NO FINAL
'|
'|  OBS: NÃO REMOVE ESPAÇOS EM BRANCO
'|

Dim s As String

    s = Texto
    If right$(RTrim$(s), 1) <> "\" Then s = s & "\"
    
    barra_invertida_add = s
    
End Function




Function barra_invertida_del(ByVal Texto As String) As String
' _________________________________________________________________________________________
'|
'|  REMOVE A BARRA INVERTIDA NO FINAL
'|
'|  OBS: NÃO REMOVE ESPAÇOS EM BRANCO
'|

Dim s As String

    s = Texto
    If right$(RTrim$(s), 1) = "\" Then s = left$(RTrim$(s), Len(RTrim$(s)) - 1)
    
    barra_invertida_del = s
    
End Function



Function ForceDirectories(ByVal nome_diretorio As String, ByRef msg_erro As String) As Boolean
' ___________________________________________________________________________________________________________________
'|
'|  CRIA O DIRETÓRIO.
'|  SE FOR NECESSÁRIO, O DIRETÓRIOS DE NÍVEL SUPERIOR TAMBÉM SERÃO CRIADOS.
'|  IMPORTANTE: ESTA ROTINA É RECURSIVA !!
'|

Dim s As String
Dim s_erro As String


    On Error GoTo FORCEDIRECTORIES_TRATA_ERRO

    
    ForceDirectories = False
    msg_erro = ""
    
    nome_diretorio = barra_invertida_del(nome_diretorio)
    s = ExtractFilePath(nome_diretorio)
    If (Trim$(s) <> "") And (right$(Trim$(s), 1) <> ":") Then
        If Not DirectoryExists(s, s_erro) Then If Not ForceDirectories(s, msg_erro) Then Exit Function
        End If
        
    If Not DirectoryExists(nome_diretorio, s_erro) Then MkDir nome_diretorio
    
    
    ForceDirectories = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FORCEDIRECTORIES_TRATA_ERRO:
'===========================
    msg_erro = "ForceDirectories (" & nome_diretorio & ") - " & CStr(Err) & ": " & Error$(Err)
    
    Exit Function
    
    
End Function


Private Function GetShortName(ByVal sLongFileName As String, ByRef msg_erro As String) As String
' ________________________________________________________________________________________________________________
'|
'|  LEMBRE-SE: NO WINDOWS NT 4, A FUNÇÃO GetShortPathName() NÃO CONSEGUE RETORNAR UM NOME
'|  NO FORMATO CURTO QUANDO O ARQUIVO ESTÁ LOCALIZADO EM UM SERVIDOR DE ARQUIVOS NA REDE.
'|  NESSES CASOS, O VALOR DE RETORNO É UMA STRING VAZIA.
'|  ISSO OCORRE TANTO PARA NOMES UNC (\\NOME_SERVIDOR\PASTA1\...) QUANTO PARA UNIDADES DE
'|  REDE MAPEADAS.
'|  NO CASO DE NOMES UNC, HÁ UM CASO ESPECÍFICO: SE A PASTA COMPARTILHADA CHAMA-SE
'|  "TEMPORARIO" (TEM MAIS QUE 8 CARACTERES), AS PASTAS SUBSEQUENTES É QUE NÃO PODEM
'|  EXCEDER O LIMITE DE 8 CARACTERES, OU SEJA, SOMENTE A PASTA QUE ESTÁ COMPARTILHADA
'|  PODE EXCEDER ESSE LIMITE.
'|

       
Dim lRetVal As Long, sShortPathName As String, iLen As Integer

    
    On Error GoTo GETSHORTNAME_TRATA_ERRO
    
    msg_erro = ""
    
       
  ' Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

  ' Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       
  ' Strip away unwanted characters.
    GetShortName = left(sShortPathName, lRetVal)
       
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GETSHORTNAME_TRATA_ERRO:
'=======================
    msg_erro = "GetShortName - " & CStr(Err) & ": " & Error$(Err)
    Err.Clear
    Exit Function
    
       
End Function



Function FileExists(ByVal nome_arquivo As String, ByRef msg_erro As String) As Boolean
' __________________________________________________________________________________________________________________________
'|
'|  INDICA SE O ARQUIVO EXISTE OU NÃO
'|

    On Error GoTo FILEEXISTS_TRATA_ERRO


    FileExists = False
    
    msg_erro = ""
    
    
    If Trim$(nome_arquivo) = "" Then Exit Function
    
    FileExists = (Dir$(nome_arquivo, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) <> "")
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FILEEXISTS_TRATA_ERRO:
'=====================
    msg_erro = "FileExists - " & CStr(Err) & ": " & Error$(Err)
    Err.Clear
    Exit Function
    
    
End Function


Function DirectoryExists(ByVal nome_diretorio As String, ByRef msg_erro As String) As Boolean
' __________________________________________________________________________________________________________________________
'|
'|  INDICA SE O DIRETÓRIO EXISTE OU NÃO
'|

Dim s As String

    
    On Error GoTo DIRECTORYEXISTS_TRATA_ERRO
    
    
    DirectoryExists = False
    
    msg_erro = ""
    
    
    If Trim$(nome_diretorio) = "" Then Exit Function
    
    
  ' ADICIONA-SE "." NO FINAL PELOS SEGUINTES MOTIVOS:
  '    1) WINDOWS NT: É SEMPRE OBRIGATÓRIO QUE TENHA O "." NO FINAL, TANTO P/ UNC QUANTO P/ DIRETÓRIOS LOCAIS
  '    2) WINDOWS 95/98: DIRETÓRIOS C/ NOMES UNC DEVEM TER O "." NO FINAL P/ EVITAR O ERRO "52: BAD FILE NAME OR NUMBER"
  ' LEMBRE-SE: "." REFERE-SE AO DIRETÓRIO CORRENTE E ".." REFERE-SE AO DIRETÓRIO ANTERIOR
    s = barra_invertida_add(nome_diretorio) & "."
    
    DirectoryExists = (Trim$(Dir$(s, vbDirectory)) <> "")
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DIRECTORYEXISTS_TRATA_ERRO:
'==========================
    msg_erro = "DirectoryExists - " & CStr(Err) & ": " & Error$(Err)
    Err.Clear
    Exit Function
    
    
End Function



Function FileAtivaReadOnly(ByVal nome_arquivo As String, ByRef msg_erro As String) As Boolean
' _______________________________________________________________________________________________
'|
'|  ATIVA O ATRIBUTO READ-ONLY DO ARQUIVO, MANTENDO OS DEMAIS ATRIBUTOS JÁ ATIVOS.
'|
'|  VALORES DE RETORNO:
'|    'TRUE': OS ATRIBUTOS DO ARQUIVO FORAM CONFIGURADOS CORRETAMENTE.
'|    'FALSE': 1) O ARQUIVO NÃO EXISTE !!
'|             2) HOUVE ERRO E A DESCRIÇÃO DO ERRO SERÁ RETORNADA EM 'msg_erro'.
'|

Dim atributos As Long
Dim i_ret As Long
    
    On Error GoTo FARO_TRATA_ERRO
    
    FileAtivaReadOnly = False
    
    msg_erro = ""
    
    i_ret = GetFileAttributes(nome_arquivo)
    If i_ret = -1 Then Exit Function
    
    atributos = i_ret Or FILE_ATTRIBUTE_READONLY
    
    i_ret = SetFileAttributes(nome_arquivo, atributos)
    FileAtivaReadOnly = (i_ret <> 0)

Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FARO_TRATA_ERRO:
'===============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    Exit Function
    
    
End Function



Function FileDesativaAtributos(ByVal nome_arquivo As String, ByRef msg_erro As String) As Boolean
' _______________________________________________________________________________________________
'|
'|  DESATIVA OS ATRIBUTOS DO ARQUIVO, MANTENDO O ATRIBUTO "ARQUIVO" ATIVADO.
'|
'|  VALORES DE RETORNO:
'|    'TRUE': OS ATRIBUTOS DO ARQUIVO FORAM CONFIGURADOS CORRETAMENTE.
'|    'FALSE': 1) O ARQUIVO NÃO EXISTE !!
'|             2) HOUVE ERRO E A DESCRIÇÃO DO ERRO SERÁ RETORNADA EM 'msg_erro'.
'|

Dim i_ret As Long
    
    On Error GoTo FDA_TRATA_ERRO
    
    FileDesativaAtributos = False
    
    msg_erro = ""
    
    i_ret = GetFileAttributes(nome_arquivo)
    If i_ret = -1 Then Exit Function
    
  ' É DIRETÓRIO ?
    If (i_ret And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then Exit Function
    
    i_ret = SetFileAttributes(nome_arquivo, 0)
    FileDesativaAtributos = (i_ret <> 0)

Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FDA_TRATA_ERRO:
'==============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    Exit Function
    
    
End Function




Function FileDelete(ByVal nome_arquivo As String, Optional ByRef msg_erro As String) As Boolean
' _______________________________________________________________________________________________
'|
'|  APAGA O(S) ARQUIVO(S) ESPECIFICADO(S) NO PARÂMETRO.
'|  O NOME DO ARQUIVO PODE CONTER CARACTERES CURINGA P/ APAGAR
'|  UM CONJUNTO DE ARQUIVOS.
'|
'|  VALORES DE RETORNO:
'|    'TRUE': OS ARQUIVOS FORAM APAGADOS.
'|    'FALSE': HOUVE ERRO E A DESCRIÇÃO DO ERRO SERÁ RETORNADA EM 'msg_erro'.
'|

Dim s As String
Dim s_arq As String

    On Error GoTo FD_TRATA_ERRO
    
    FileDelete = False
    
    msg_erro = ""
    
    s_arq = Dir$(nome_arquivo, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
    Do While s_arq <> ""
        If (s_arq <> ".") And (s_arq <> "..") Then
            s = barra_invertida_add(ExtractFilePath(nome_arquivo)) & s_arq
            FileDesativaAtributos s, msg_erro
            Kill s
            End If
      
      ' OBTÉM O NOME DO PRÓXIMO ARQUIVO (NO CASO DE HAVER CARACTERES CURINGA)
        s_arq = Dir$
        Loop
    
    FileDelete = True

Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FD_TRATA_ERRO:
'=============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    Exit Function
    
    
End Function

Function FileSetAttributes(ByVal nome_arquivo As String, ByVal atributos As Long, ByRef msg_erro As String) As Boolean
' _______________________________________________________________________________________________
'|
'|  CONFIGURA OS ATRIBUTOS DO ARQUIVO ESPECIFICADO.
'|
'|  VALORES DE RETORNO:
'|    'TRUE': OS ATRIBUTOS DO ARQUIVO FORAM CONFIGURADOS CORRETAMENTE.
'|    'FALSE': 1) O ARQUIVO NÃO EXISTE !!
'|             2) HOUVE ERRO E A DESCRIÇÃO DO ERRO SERÁ RETORNADA EM 'msg_erro'.
'|

Dim i_ret As Long
    
    On Error GoTo FSA_TRATA_ERRO
    
    FileSetAttributes = False
    
    msg_erro = ""
    
    i_ret = SetFileAttributes(nome_arquivo, atributos)
    FileSetAttributes = (i_ret <> 0)

Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FSA_TRATA_ERRO:
'==============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    Exit Function
    
    
End Function


Function RemoveDir(ByVal nome_diretorio As String, ByRef msg_erro As String) As Boolean
' ___________________________________________________________________________________________________________________
'|
'|  REMOVE O DIRETÓRIO E TODOS OS SEUS ARQUIVOS E SUBDIRETÓRIOS.
'|  IMPORTANTE: ESTA ROTINA É RECURSIVA !!
'|

Dim s As String
Dim s_dir As String
Dim s_arq As String
Dim s_erro As String
Dim i As Long
Dim v_dir() As String
Dim v_arq() As String


    On Error GoTo REMOVEDIR_TRATA_ERRO


    RemoveDir = False
    msg_erro = ""
    
    
  ' APAGA TODOS OS ARQUIVOS
    ReDim v_arq(0)
    s_arq = Dir$(barra_invertida_add(nome_diretorio) & "*.*", vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
    Do While (s_arq <> "")
        s = barra_invertida_add(nome_diretorio) & s_arq
        If Trim$(v_arq(UBound(v_arq))) <> "" Then ReDim Preserve v_arq(UBound(v_arq) + 1)
        v_arq(UBound(v_arq)) = s
        
        s_arq = Dir$
        Loop

    For i = LBound(v_arq) To UBound(v_arq)
        If FileExists(v_arq(i), s_erro) Then
            SetAttr v_arq(i), vbNormal
            Kill v_arq(i)
            End If
        Next


  ' REMOVE TODOS OS SUBDIRETÓRIOS
    ReDim v_dir(0)
    nome_diretorio = barra_invertida_del(nome_diretorio)
    s_dir = Dir$(barra_invertida_add(nome_diretorio) & "*.*", vbDirectory)
    
    Do While (s_dir <> "")
        If (s_dir <> ".") And (s_dir <> "..") Then
            s = barra_invertida_add(nome_diretorio) & s_dir
            If Trim$(v_dir(UBound(v_dir))) <> "" Then ReDim Preserve v_dir(UBound(v_dir) + 1)
            v_dir(UBound(v_dir)) = s
            End If
            
        s_dir = Dir$
        Loop
    
    For i = LBound(v_dir) To UBound(v_dir)
        If DirectoryExists(v_dir(i), s_erro) Then If Not RemoveDir(v_dir(i), msg_erro) Then Exit Function
        Next
    
    
    
  ' REMOVE O DIRETÓRIO
    RmDir nome_diretorio
        
    
    RemoveDir = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
REMOVEDIR_TRATA_ERRO:
'====================
    msg_erro = "RemoveDir - " & CStr(Err) & ": " & Error$(Err)
    
    Exit Function
    
    
End Function


