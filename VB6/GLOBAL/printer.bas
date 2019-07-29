Attribute VB_Name = "mod_PRINTER"
'�_____________________________________________________________________________________
'|
'|  ESTE M�DULO VISA FORNECER FUN��ES QUE RETORNEM INFORMA��ES SOBRE A
'|  CONFIGURA��O DA IMPRESSORA, PRINCIPALMENTE SOBRE AS MARGENS DO PAPEL.
'|
'|  BUG (Q156696): PlayEnhMetaFile() N�O LIBERA RECURSOS DO GDI NO WINDOWS
'|      95 E 98. PORTANTO, PARA RELAT�RIOS COM MUITAS P�GINAS, SER� PRECISO
'|      IMPLEMENTAR UMA ALTERNATIVA OU EXECUTAR O PROGRAMA EM WINDOWS NT/2000.
'|
'|  BUG (PRINTER.ENDDOC): H� ALGUM BUG NO COMANDO PRINTER.ENDDOC QUE FAZ PERDER
'|      RECURSOS DO GDI, MAIS OU MENOS 0.5% A CADA EXECU��O.
'|      OS RECURSOS PERDIDOS NUNCA MAIS S�O RECUPERADOS, MESMO FECHANDO O APLICATIVO.
'|      PORTANTO, SEMPRE QUE POSS�VEL, DEVE-SE USAR O PRINTER.NEWPAGE
'|
'|  BUG (Q176634): AO DESENHAR UM QUADRADO USANDO O M�TODO LINE OU UM C�RCULO
'|      USANDO CIRCLE, O FUNDO N�O FICA TRANSPARENTE E ENCOBRE O QUE ESTIVER POR
'|      BAIXO.  A SOLU��O USADA PARA ESTE BUG, ENTRETANTO, PROVOCA A PERDA DE
'|      RECURSOS DO GDI (MAIS OU MENOS 0,125% A CADA EXECU��O), SENDO QUE,
'|      EXCEPCIONALMENTE NESTE CASO, OS RECURSOS S�O RECUPERADOS QUANDO O
'|      APLICATIVO � FECHADO.
'|      A PERDA DE RECURSOS OCORRE POR P�GINA EM QUE SE EXECUTA OS COMANDOS DE
'|      CORRE��O DO BUG, OU SEJA, SE OS COMANDOS FOREM EXECUTADOS 1 VEZ OU 1000
'|      VEZES EM UMA MESMA P�GINA, A PERDA SER� A MESMA.
'|
'|

Option Explicit


Private Const PRINTER_MARGEM_TOPO_POL = 0.18
Private Const PRINTER_MARGEM_INF_POL = 0.25
Private Const PRINTER_MARGEM_ESQ_POL = 0.25
Private Const PRINTER_MARGEM_DIR_POL = 0.25


Type TIPO_DIMENSAO_LOGO
    largura As Single
    altura As Single
    End Type



Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

'�Constants for nIndex argument of GetDeviceCaps
Private Const HORZRES = 8
Private Const VERTRES = 10
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const PHYSICALWIDTH = 110
Private Const PHYSICALHEIGHT = 111
Private Const PHYSICALOFFSETX = 112
Private Const PHYSICALOFFSETY = 113


'�IMPRESS�O DE METAFILE
Type RECTL
    left As Long
    top As Long
    right As Long
    bottom As Long
    End Type

Private Type RECT_16B
    left As Integer
    top As Integer
    right As Integer
    bottom As Integer
    End Type

Private Type SIZEL
    cx As Long
    cy As Long
    End Type

Type ENHMETAHEADER
    iType As Long
    nSize As Long
    rclBounds As RECTL
    rclFrame As RECTL
    dSignature As Long
    nVersion As Long
    nBytes As Long
    nRecords As Long
    nHandles As Integer
    sReserved As Integer
    nDescription As Long
    offDescription As Long
    nPalEntries As Long
    szlDevice As SIZEL
    szlMillimeters As SIZEL
    End Type

Private Type APMFILEHEADER
    key As Long
    hMF As Integer
    bbox As RECT_16B
    inch As Integer
    Reserved As Long
    checksum As Integer
    End Type
    
Private Type METAHEADER
    mtType As Integer
    mtHeaderSize As Integer
    mtVersion As Integer
    mtSize As Long
    mtNoObjects As Integer
    mtMaxRecord As Long
    mtNoParameters As Integer
    End Type
        
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
    End Type
    
    
Private Declare Function GetObj Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long

Private Declare Function GetEnhMetaFile Lib "gdi32" Alias "GetEnhMetaFileA" (ByVal lpszMetaFile As String) As Long
Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hEMF As Long, lpRect As Any) As Long
Private Declare Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hEMF As Long, ByVal cbBuffer As Long, lpemh As ENHMETAHEADER) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEMF As Long) As Long

Private Declare Function GetMetaFile Lib "gdi32" Alias "GetMetaFileA" (ByVal lpFileName As String) As Long
Private Declare Function PlayMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMF As Long) As Long
Private Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hMF As Long, ByVal nSize As Long, lpvData As Any) As Long
Private Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Byte) As Long
Private Declare Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hDCRef As Long, lpmfp As Any) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long

Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal iBkMode As Long) As Long


Public Function logo_imprime(ByVal nome_arquivo As String, _
                             ByRef dimensao_impressa As TIPO_DIMENSAO_LOGO, _
                             dimensao_maxima As TIPO_DIMENSAO_LOGO, _
                             ByVal originX As Single, ByVal OriginY As Single, _
                             Optional ByVal largura As Single, _
                             Optional ByVal altura As Single) As Boolean
'�________________________________________________________________________________________________________________________
'|
'|  IMPRIME O LOGOTIPO CONTIDO NO ARQUIVO ESPECIFICADO EM 'NOME_ARQUIVO'.
'|
'|  O LOGOTIPO PODE SER DO TIPO:
'|     BITMAP (.BMP)
'|     ENHANCED METAFILE (.EMF), METAFILE (.WMF)
'|     GIF (.GIF)
'|     JPEG (.JPG)
'|
'|  1. ORIGINX E ORIGINY: POSI��O INICIAL DA FIGURA, EM MIL�METROS
'|  2. 'LARGURA' E 'ALTURA': DIMENS�ES, EM MIL�METROS, EM QUE A FIGURA DEVE SER IMPRESSA.
'|  3. SE 'LARGURA' E 'ALTURA' N�O FOREM FORNECIDOS, IMPRIME NO TAMANHO ORIGINAL.
'|  4. SE APENAS UMA DAS DIMENS�ES FOR FORNECIDA, CALCULA A PROPOR��O E APLICA P/
'|     A DIMENS�O N�O FORNECIDA (PARA EVITAR DISTOR��ES, RECOMENDA-SE FORNECER APENAS UMA
'|     DAS DIMENS�ES).
'|
'|  BUG (Q156696): A FUN��O DA API PlayEnhMetaFile() N�O LIBERA RECURSOS DO GDI
'|      NO WINDOWS 95 E 98. ISSO VALE TAMB�M PARA O M�TODO 'PAINTPICTURE' QUANDO
'|      SE IMPRIME METAFILES.  ESSE PROBLEMA N�O OCORRE EM WINDOWS NT/2000.
'|

Dim s As String
Dim s_erro As String
Dim h_max As Single
Dim w_max As Single
Dim fator As Single
Dim largura_ok As Boolean
Dim altura_ok As Boolean
Dim scalemode_a As Long
Dim p_logo As Picture
Dim xi As Single
Dim yi As Single
Dim dx As Single
Dim dy As Single


    On Error GoTo LOGO_IMPRIME_TRATA_ERRO
    
    
    logo_imprime = False
    
    
  '�SALVA CONTEXTO
    scalemode_a = Printer.ScaleMode
    
    
  '�INICIALIZA VALOR DE RETORNO
    With dimensao_impressa
        .largura = 0
        .altura = 0
        End With


  '�PARA ASSEGURAR TOTAL COMPATIBILIDADE, OBT�M O NOME NO FORMATO 8.3
    s = GetShortName(nome_arquivo)
    If Trim$(s) <> "" Then nome_arquivo = s
    
    
  '�ARQUIVO N�O EXISTE !
    If Not FileExists(nome_arquivo, s_erro) Then
      '�RESTAURA CONTEXTO
        Printer.ScaleMode = scalemode_a
        Exit Function
        End If
    
    
  '�CARREGA O LOGOTIPO DO ARQUIVO
    Set p_logo = LoadPicture(nome_arquivo)
      
      
  '�AJUSTA TAMANHO PARA PIXELS
    yi = convert_printer_mm_to_twipsY(OriginY) / Printer.TwipsPerPixelY
    xi = convert_printer_mm_to_twipsX(originX) / Printer.TwipsPerPixelX
            
            
  '�CALCULA DIMENS�ES
    largura_ok = False
    If IsNumeric(largura) Then If CSng(largura) > 0 Then largura_ok = True
    altura_ok = False
    If IsNumeric(altura) Then If CSng(altura) > 0 Then altura_ok = True
    
    If largura_ok And altura_ok Then
        dy = convert_printer_mm_to_twipsY(altura) / Printer.TwipsPerPixelY
        dx = convert_printer_mm_to_twipsX(largura) / Printer.TwipsPerPixelX
    ElseIf altura_ok Then
        fator = (convert_printer_mm_to_twipsY(altura) / Printer.TwipsPerPixelY) / p_logo.Height
        dy = convert_printer_mm_to_twipsY(altura) / Printer.TwipsPerPixelY
        dx = p_logo.Width * fator
    ElseIf largura_ok Then
        fator = (convert_printer_mm_to_twipsX(largura) / Printer.TwipsPerPixelX) / p_logo.Width
        dx = convert_printer_mm_to_twipsX(largura) / Printer.TwipsPerPixelX
        dy = p_logo.Height * fator
    Else
        dy = p_logo.Height
        dx = p_logo.Width
        End If
    
            
  '�EXCEDE ALGUM LIMITE M�XIMO ?
    h_max = convert_printer_mm_to_twipsY(dimensao_maxima.altura) / Printer.TwipsPerPixelY
    w_max = convert_printer_mm_to_twipsX(dimensao_maxima.largura) / Printer.TwipsPerPixelX
    
    If dy > h_max Then
        fator = h_max / dy
        dy = h_max
        dx = dx * fator
        End If
    
    If dx > w_max Then
        fator = w_max / dx
        dx = w_max
        dy = dy * fator
        End If
        
          
  '�RETORNA DIMENS�ES IMPRESSAS
    dimensao_impressa.largura = convert_printer_pixelsX_to_mm(dx)
    dimensao_impressa.altura = convert_printer_pixelsY_to_mm(dy)
    
       
  '�IMPRIME !!
    Printer.ScaleMode = vbPixels
    Printer.PaintPicture p_logo, xi, yi, dx, dy
            
    Set p_logo = Nothing
    
        
  '�RESTAURA CONTEXTO
    Printer.ScaleMode = scalemode_a
    
    
    logo_imprime = True
    
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LOGO_IMPRIME_TRATA_ERRO:
'=======================
    s = CStr(Err) & ": " & Error$(Err)
    s = "Impress�o do logotipo: " & nome_arquivo & vbCrLf & vbCrLf & s
    MsgBox s, vbCritical
    
  '�RESTAURA CONTEXTO
    Printer.ScaleMode = scalemode_a
    
    Exit Function
    
    
End Function

Private Function GetShortName(ByVal sLongFileName As String) As String
'�________________________________________________________________________________________________________________
'|
'|  LEMBRE-SE: NO WINDOWS NT 4, A FUN��O GetShortPathName() N�O CONSEGUE RETORNAR UM NOME
'|  NO FORMATO CURTO QUANDO O ARQUIVO EST� LOCALIZADO EM UM SERVIDOR DE ARQUIVOS NA REDE.
'|  NESSES CASOS, O VALOR DE RETORNO � UMA STRING VAZIA.
'|  ISSO OCORRE TANTO PARA NOMES UNC (\\NOME_SERVIDOR\PASTA1\...) QUANTO PARA UNIDADES DE
'|  REDE MAPEADAS.
'|  NO CASO DE NOMES UNC, H� UM CASO ESPEC�FICO: SE A PASTA COMPARTILHADA CHAMA-SE
'|  "TEMPORARIO" (TEM MAIS QUE 8 CARACTERES), AS PASTAS SUBSEQUENTES � QUE N�O PODEM
'|  EXCEDER O LIMITE DE 8 CARACTERES, OU SEJA, SOMENTE A PASTA QUE EST� COMPARTILHADA
'|  PODE EXCEDER ESSE LIMITE.
'|

Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       
   'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

   'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       
   'Strip away unwanted characters.
    GetShortName = left(sShortPathName, lRetVal)
       
       
End Function

Public Function convert_printer_mm_to_pixelsX(ByVal valor_em_mm) As Variant
'�______________________________________________________________________________________________________
'|
'|  CONVERTE DE MIL�METROS PARA PIXELS NO SENTIDO HORIZONTAL
'|

    convert_printer_mm_to_pixelsX = Printer.ScaleX(valor_em_mm, vbMillimeters, vbPixels)

End Function

Public Function convert_printer_mm_to_pixelsY(ByVal valor_em_mm) As Variant
'�______________________________________________________________________________________________________
'|
'|  CONVERTE DE MIL�METROS PARA PIXELS NO SENTIDO VERTICAL
'|

    convert_printer_mm_to_pixelsY = Printer.ScaleY(valor_em_mm, vbMillimeters, vbPixels)

End Function

Public Function convert_printer_pixelsX_to_mm(ByVal valor_em_pixels) As Variant
'�______________________________________________________________________________________________________
'|
'|  CONVERTE DE PIXELS PARA MIL�METROS NO SENTIDO HORIZONTAL
'|

    convert_printer_pixelsX_to_mm = Printer.ScaleX(valor_em_pixels, vbPixels, vbMillimeters)

End Function

Public Function convert_printer_pixelsY_to_mm(ByVal valor_em_pixels) As Variant
'�______________________________________________________________________________________________________
'|
'|  CONVERTE DE PIXELS PARA MIL�METROS NO SENTIDO VERTICAL
'|

    convert_printer_pixelsY_to_mm = Printer.ScaleY(valor_em_pixels, vbPixels, vbMillimeters)

End Function



Sub printer_circulo(ByVal xi As Single, ByVal yi As Single, ByVal raio As Single, Optional ByVal cor As Long = -1, Optional ByVal fill_style As Long = -1)
'�_________________________________________________________________________________________________________________________________________________________
'|
'|  ESTA FUN��O DEVE SER UTILIZADA SEMPRE QUE SE DESEJAR DESENHAR UM C�RCULO USANDO O
'|  M�TODO PRINTER.CIRCLE
'|
'|  O OBJETIVO DESTA FUN��O � REDUZIR A PERDA DE RECURSOS DO GDI QUE OCORREM AO
'|  EXECUTAR OS COMANDOS NECESS�RIOS P/ CORRIGIR O BUG Q176634.
'|  USANDO ESTA FUN��O, OS COMANDOS QUE CORRIGEM O BUG SOMENTE S�O EXECUTADOS QUANDO
'|  ESTRITAMENTE NECESS�RIOS.
'|
'|  BUG (Q176634): AO DESENHAR UM QUADRADO USANDO O M�TODO LINE OU UM C�RCULO
'|      USANDO CIRCLE, O FUNDO N�O FICA TRANSPARENTE E ENCOBRE O QUE ESTIVER POR
'|      BAIXO.  A SOLU��O USADA PARA ESTE BUG, ENTRETANTO, PROVOCA A PERDA DE
'|      RECURSOS DO GDI (MAIS OU MENOS 0,125% A CADA EXECU��O), SENDO QUE,
'|      EXCEPCIONALMENTE NESTE CASO, OS RECURSOS S�O RECUPERADOS QUANDO O
'|      APLICATIVO � FECHADO.
'|      A PERDA DE RECURSOS OCORRE POR P�GINA EM QUE SE EXECUTA OS COMANDOS DE
'|      CORRE��O DO BUG, OU SEJA, SE OS COMANDOS FOREM EXECUTADOS 1 VEZ OU 1000
'|      VEZES EM UMA MESMA P�GINA, A PERDA SER� A MESMA.
'|


Dim fillstyle_a As Long


  '�SALVA FILLSTYLE ORIGINAL
    fillstyle_a = Printer.FillStyle
    
  '�CORRIGE BUG Q176634
    Printer.FillStyle = vbHorizontalLine
    Printer.Print ""
    Printer.FillStyle = vbFSTransparent
        
    
  '�N�O FORNECEU FILLSTYLE NO PAR�METRO
  '�~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If fill_style = -1 Then
      '�RESTAURA FILLSTYLE ORIGINAL
        Printer.FillStyle = fillstyle_a
    
        If cor = -1 Then
          '�N�O ESPECIFICOU COR
            Printer.Circle (xi, yi), raio
        Else
          '�ESPECIFICOU COR
            Printer.Circle (xi, yi), raio, cor
            End If
            
            
  '�FORNECEU FILLSTYLE NO PAR�METRO
  '�~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Else
        Printer.FillStyle = fill_style
        
        If cor = -1 Then
          '�N�O ESPECIFICOU COR
            Printer.Circle (xi, yi), raio
        Else
          '�ESPECIFICOU COR
            Printer.Circle (xi, yi), raio, cor
            End If
        
      '�RESTAURA FILLSTYLE ORIGINAL
        Printer.FillStyle = fillstyle_a
        End If
        
    
End Sub

Public Sub printer_inicializa_pagina()
'�____________________________________________________________________________________
'|
'|  EXECUTA PROCEDIMENTOS DE INICIALIZA��O PARA NOVO DOCUMENTO E/OU P�GINA.
'|

Dim CurrentY_a As Single
Dim CurrentX_a As Single

    
    CurrentY_a = Printer.CurrentY
    CurrentX_a = Printer.CurrentX
    
    
  '�POSICIONA EM COORDENADA V�LIDA: SE CURRENTY ULTRAPASSAR SCALEHEIGHT, UMA NOVA
  '�P�GINA SER� CRIADA AUTOMATICAMENTE, COMO SE UM NEWPAGE TIVESSE SIDO EXECUTADO.
    Printer.CurrentY = 0
    Printer.CurrentX = 0
   

  '�BUG: PARA QUE OS LOGOTIPOS EM WMF/EMF N�O SAIAM EM BRANCO, GARANTE QUE A IMPRESSORA
  '�~~~  EST� INICIALIZADA. ISSO DEVE SER FEITO AP�S O NEWPAGE E O ENDDOC.
    Printer.Print " "


  '�RESTAURA POSI��O ORIGINAL
    Printer.CurrentY = CurrentY_a
    Printer.CurrentX = CurrentX_a
    
    
End Sub

Public Function convert_pol_to_mm(ByVal valor_em_pol) As Variant
'�______________________________________________________________________________________________________
'|
'|  CONVERTE DE POLEGADAS PARA MIL�METROS
'|

    convert_pol_to_mm = valor_em_pol * 25.4
    
End Function

Public Function convert_mm_to_pol(ByVal valor_em_mm) As Variant
'�______________________________________________________________________________________________________
'|
'|  CONVERTE DE MIL�METROS PARA POLEGADAS
'|

    convert_mm_to_pol = valor_em_mm / 25.4
    
End Function
Public Function convert_printer_mm_to_twipsX(ByVal valor_em_mm) As Variant
'�______________________________________________________________________________________________________
'|
'|  CONVERTE DE MIL�METROS PARA TWIPS NO SENTIDO HORIZONTAL
'|

    convert_printer_mm_to_twipsX = Printer.ScaleX(valor_em_mm, vbMillimeters, vbTwips)
    
End Function

Public Function convert_printer_pol_to_twipsX(ByVal valor_em_pol) As Variant
'�______________________________________________________________________________________________________
'|
'|  CONVERTE DE POLEGADAS PARA TWIPS NO SENTIDO HORIZONTAL
'|

    convert_printer_pol_to_twipsX = Printer.ScaleX(valor_em_pol, vbInches, vbTwips)
    
End Function


Public Function convert_printer_mm_to_twipsY(ByVal valor_em_mm) As Variant
'�______________________________________________________________________________________________________
'|
'|  CONVERTE DE MIL�METROS PARA TWIPS NO SENTIDO VERTICAL
'|

    convert_printer_mm_to_twipsY = Printer.ScaleY(valor_em_mm, vbMillimeters, vbTwips)
    
End Function

Public Function convert_printer_pol_to_twipsY(ByVal valor_em_pol) As Variant
'�______________________________________________________________________________________________________
'|
'|  CONVERTE DE POLEGADAS PARA TWIPS NO SENTIDO VERTICAL
'|

    convert_printer_pol_to_twipsY = Printer.ScaleY(valor_em_pol, vbInches, vbTwips)
    
End Function



Public Function get_printer_resolucao_dpi_x() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A RESOLU��O DA IMPRESSORA PARA O SENTIDO HORIZONTAL
'|

   get_printer_resolucao_dpi_x = GetDeviceCaps(Printer.hdc, LOGPIXELSX)
   
   
End Function
Public Function get_printer_resolucao_dpi_y() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A RESOLU��O DA IMPRESSORA PARA O SENTIDO VERTICAL
'|

   get_printer_resolucao_dpi_y = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
   
   
End Function


Public Function get_printer_margem_fisica_esq_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|   RETORNA A MARGEM ESQUERDA M�NIMA ACEITA PELA IMPRESSORA, EM POLEGADAS.
'|
'| _ O VALOR RETORNADO INDICA QUAL � MARGEM ESQUERDA M�NIMA ACEITA
'|   PELA IMPRESSORA.
'| _ ENTRETANTO, QUANDO A IMPRESS�O � FEITA EM LANDSCAPE (PAPEL DEITADO)
'|   OCORRE UM PROBLEMA NAS IMPRESSORAS JATO DE TINTA: A IMPRESS�O FICA
'|   MUITO DESLOCADA P/ A ESQUERDA.  ISSO OCORRE PORQUE ESSE TIPO DE
'|   IMPRESSORA TEM CARACTER�STICAS MEC�NICAS QUE PERMITEM TER UMA
'|   MARGEM SUPERIOR MUITO PEQUENA, MAS A MARGEM INFERIOR PRECISA SER
'|   RELATIVAMENTE GRANDE, POIS O PAPEL PRECISA ESTAR "PRESO" PELOS
'|   TRACIONADORES P/ QUE FIQUE FIRME E AINDA � NECESS�RIO UM ESPA�O
'|   SUFICIENTE P/ A CABE�A DE IMPRESS�O PODER SE DESLOCAR.
'| _ PORTANTO, USE SEMPRE A FUN��O GET_PRINTER_MARGEM_UTIL_ESQ_POL() P/
'|   O VALOR DA MARGEM ESQUERDA.
'| _ ESSE INCONVENIENTE N�O OCORRE P/ A MARGEM SUPERIOR, POIS MESMO QUE
'|   QUE A MARGEM SUPERIOR SEJA MUITO MENOR QUE A INFERIOR, A IMPRESS�O
'|   N�O FICA PARECENDO DESLOCADA E SEM SIMETRIA.
'|

Dim i As Single

    i = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
    i = i / get_printer_resolucao_dpi_x()
    
    get_printer_margem_fisica_esq_pol = i
   
   
End Function
Public Function get_printer_margem_util_esq_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A MARGEM ESQUERDA DA IMPRESSORA, EM POLEGADAS.
'|
'|  O VALOR RETORNADO � RESULTADO DE UM PROCESSAMENTO QUE VISA OBTER UMA
'|  PADRONIZA��O ENTRE IMPRESSORAS LASER E JATO DE TINTA, J� QUE AS IMPRESSORAS
'|  JATO DE TINTA TEM A MARGEM SUPERIOR PEQUENA E A MARGEM INFERIOR GRANDE.
'|
'|  IMPORTANTE: O OBJETIVO � OBTER VALORES IGUAIS P/ AS MARGENS DIREITA E
'|              ESQUERDA P/ QUE FIQUEM SIM�TRICAS.
'|

Dim i As Single

    i = maior(get_printer_margem_fisica_esq_pol(), get_printer_margem_fisica_dir_pol())
    i = maior(i, PRINTER_MARGEM_ESQ_POL)
    
    get_printer_margem_util_esq_pol = i
   
   
End Function

  
Public Function get_printer_margem_fisica_dir_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|   RETORNA A MARGEM DIREITA M�NIMA ACEITA PELA IMPRESSORA, EM POLEGADAS.
'|
'| _ O VALOR RETORNADO INDICA QUAL � MARGEM DIREITA M�NIMA ACEITA
'|   PELA IMPRESSORA.
'| _ ENTRETANTO, QUANDO A IMPRESS�O � FEITA EM LANDSCAPE (PAPEL DEITADO)
'|   OCORRE UM PROBLEMA NAS IMPRESSORAS JATO DE TINTA: A IMPRESS�O FICA
'|   MUITO DESLOCADA P/ A ESQUERDA.  ISSO OCORRE PORQUE ESSE TIPO DE
'|   IMPRESSORA TEM CARACTER�STICAS MEC�NICAS QUE PERMITEM TER UMA
'|   MARGEM SUPERIOR MUITO PEQUENA, MAS A MARGEM INFERIOR PRECISA SER
'|   RELATIVAMENTE GRANDE, POIS O PAPEL PRECISA ESTAR "PRESO" PELOS
'|   TRACIONADORES P/ QUE FIQUE FIRME E AINDA � NECESS�RIO UM ESPA�O
'|   SUFICIENTE P/ A CABE�A DE IMPRESS�O PODER SE DESLOCAR.
'| _ PORTANTO, USE SEMPRE A FUN��O GET_PRINTER_MARGEM_UTIL_DIR_POL() P/
'|   O VALOR DA MARGEM DIREITA.
'| _ ESSE INCONVENIENTE N�O OCORRE P/ A MARGEM SUPERIOR, POIS MESMO QUE
'|   QUE A MARGEM SUPERIOR SEJA MUITO MENOR QUE A INFERIOR, A IMPRESS�O
'|   N�O FICA PARECENDO DESLOCADA E SEM SIMETRIA.
'|

Dim i As Single

    i = get_printer_largura_fisica_pol()
    i = i - get_printer_largura_util_real_pol()
    i = i - get_printer_margem_fisica_esq_pol()
             
    get_printer_margem_fisica_dir_pol = i
   
   
End Function

Public Function get_printer_offset_dir_mm(ByVal abscissa_final_em_mm As Single) As Variant
'�______________________________________________________________________________________________________
'|
'|  CALCULA E RETORNA O VALOR NECESS�RIO PARA SER ADICIONADO � MARGEM DIREITA.
'|
'|  ESTA ROTINA TORNOU-SE NECESS�RIA PORQUE NOTOU-SE UM BUG NA "HP DeskJet 930C Series" QUE
'|  INFORMA UM VALOR ERRADO PARA A LARGURA �TIL DA P�GINA QUANDO A ORIENTA��O DO PAPEL EST�
'|  CONFIGURADA PARA PAISAGEM.  NESTE CASO, A LARGURA INFORMADA � ALGUNS MILIMETROS MAIOR
'|  DO QUE REALMENTE A IMPRESSORA � CAPAZ DE IMPRIMIR.
'|

Dim sw As Single
Dim i_min As Single
Dim s As String

    
    get_printer_offset_dir_mm = 0
  
  
  '�OBT�M A LARGURA �TIL DA IMPRESSORA
    sw = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbMillimeters)
    
    
  '�DEFINE O ESPA�AMENTO M�NIMO COM RELA��O � MARGEM DIREITA
    i_min = 1
    If Printer.Orientation = vbPRORLandscape Then
        s = UCase$(Trim$(Printer.DeviceName))
        If (InStr(s, "HP ") <> 0) And (InStr(s, " DESKJET") <> 0) Then i_min = 4
        End If
    
    
  '�VERIFICA QUAL A DIST�NCIA ENTRE A MAIOR ABSCISSA UTILIZADA E A M�XIMA ABSCISSA ACEITA PELA IMPRESSORA
    If (sw - abscissa_final_em_mm) < i_min Then
      '�RETORNA O VALOR INDICANDO O QUANTO A MARGEM DEVE AUMENTAR
        get_printer_offset_dir_mm = i_min - (sw - abscissa_final_em_mm)
        End If
    
    
End Function

Public Function get_printer_margem_fisica_inf_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A MARGEM INFERIOR M�NIMA ACEITA PELA IMPRESSORA, EM POLEGADAS.
'|

Dim i As Single

    i = get_printer_altura_fisica_pol()
    i = i - get_printer_altura_util_real_pol()
    i = i - get_printer_margem_fisica_topo_pol()
   
    get_printer_margem_fisica_inf_pol = i
   
   
End Function

Public Function get_printer_margem_util_inf_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A MARGEM INFERIOR DA IMPRESSORA, EM POLEGADAS.
'|
'|  O VALOR RETORNADO � RESULTADO DE UM PROCESSAMENTO QUE VISA OBTER UMA
'|  PADRONIZA��O ENTRE IMPRESSORAS LASER E JATO DE TINTA, J� QUE AS IMPRESSORAS
'|  JATO DE TINTA TEM A MARGEM SUPERIOR PEQUENA E A MARGEM INFERIOR GRANDE.
'|
'|  IMPORTANTE: AS MARGENS SUPERIOR E INFERIOR N�O PRECISAM SER IGUAIS.
'|

Dim i As Single

    i = maior(get_printer_margem_fisica_inf_pol(), PRINTER_MARGEM_INF_POL)
   
    get_printer_margem_util_inf_pol = i
   
   
End Function


Public Function get_printer_margem_util_topo_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A MARGEM SUPERIOR DA IMPRESSORA, EM POLEGADAS.
'|
'|  O VALOR RETORNADO � RESULTADO DE UM PROCESSAMENTO QUE VISA OBTER UMA
'|  PADRONIZA��O ENTRE IMPRESSORAS LASER E JATO DE TINTA, J� QUE AS IMPRESSORAS
'|  JATO DE TINTA TEM A MARGEM SUPERIOR PEQUENA E A MARGEM INFERIOR GRANDE.
'|
'|  IMPORTANTE: AS MARGENS SUPERIOR E INFERIOR N�O PRECISAM SER IGUAIS.
'|

Dim i As Single

    i = maior(get_printer_margem_fisica_topo_pol(), PRINTER_MARGEM_TOPO_POL)
    
    get_printer_margem_util_topo_pol = i
    
    
End Function




Public Function get_printer_largura_util_real_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A LARGURA M�XIMA DO PAPEL EM QUE A IMPRESSORA PODE REALMENTE IMPRIMIR.
'|  O VALOR EST� EM POLEGADAS.
'|

Dim i As Single

    i = GetDeviceCaps(Printer.hdc, HORZRES)
    i = i / get_printer_resolucao_dpi_x()
    
    get_printer_largura_util_real_pol = i
   
   
End Function
Public Function get_printer_altura_util_real_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A ALTURA M�XIMA DO PAPEL EM QUE A IMPRESSORA PODE REALMENTE IMPRIMIR.
'|  O VALOR EST� EM POLEGADAS.
'|

Dim i As Single

    i = GetDeviceCaps(Printer.hdc, VERTRES)
    i = i / get_printer_resolucao_dpi_y()
    
    get_printer_altura_util_real_pol = i
    
    
End Function

Public Function get_printer_altura_util_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A ALTURA DO PAPEL EM QUE A IMPRESSORA PODE IMPRIMIR, EM POLEGADAS.
'|
'|  O VALOR RETORNADO � RESULTADO DE UM PROCESSAMENTO QUE VISA OBTER UMA
'|  PADRONIZA��O ENTRE IMPRESSORAS LASER E JATO DE TINTA, J� QUE AS IMPRESSORAS
'|  JATO DE TINTA TEM A MARGEM SUPERIOR PEQUENA E A MARGEM INFERIOR GRANDE.
'|

Dim i As Single

    i = get_printer_altura_fisica_pol()
    i = i - get_printer_margem_util_topo_pol()
    i = i - get_printer_margem_util_inf_pol()
    
    get_printer_altura_util_pol = i
    
    
End Function

Public Function get_printer_largura_util_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A LARGURA DO PAPEL EM QUE A IMPRESSORA PODE IMPRIMIR, EM POLEGADAS.
'|
'|  O VALOR RETORNADO � RESULTADO DE UM PROCESSAMENTO QUE VISA OBTER UMA
'|  PADRONIZA��O ENTRE IMPRESSORAS LASER E JATO DE TINTA, J� QUE AS IMPRESSORAS
'|  JATO DE TINTA TEM A MARGEM SUPERIOR PEQUENA E A MARGEM INFERIOR GRANDE.
'|

Dim i As Single

    i = get_printer_largura_fisica_pol()
    i = i - get_printer_margem_util_esq_pol()
    i = i - get_printer_margem_util_dir_pol()
    
    get_printer_largura_util_pol = i
   
   
End Function

Public Function get_printer_margem_util_dir_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A MARGEM DIREITA DA IMPRESSORA, EM POLEGADAS.
'|
'|  O VALOR RETORNADO � RESULTADO DE UM PROCESSAMENTO QUE VISA OBTER UMA
'|  PADRONIZA��O ENTRE IMPRESSORAS LASER E JATO DE TINTA, J� QUE AS IMPRESSORAS
'|  JATO DE TINTA TEM A MARGEM SUPERIOR PEQUENA E A MARGEM INFERIOR GRANDE.
'|
'|  IMPORTANTE: O OBJETIVO � OBTER VALORES IGUAIS P/ AS MARGENS DIREITA E
'|              ESQUERDA P/ QUE FIQUEM SIM�TRICAS.
'|

Dim i As Single

    i = maior(get_printer_margem_fisica_esq_pol(), get_printer_margem_fisica_dir_pol())
    i = maior(i, PRINTER_MARGEM_DIR_POL)
             
    get_printer_margem_util_dir_pol = i
   
   
End Function


Public Function get_printer_largura_fisica_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A LARGURA ABSOLUTA (F�SICA) DO PAPEL QUE EST� SENDO UTILIZADO NA IMPRESSORA.
'|  O VALOR EST� EM POLEGADAS.
'|

Dim i As Single

    i = GetDeviceCaps(Printer.hdc, PHYSICALWIDTH)
    i = i / get_printer_resolucao_dpi_x()
    
    get_printer_largura_fisica_pol = i
    
    
End Function
Public Function get_printer_altura_fisica_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A ALTURA ABSOLUTA (F�SICA) DO PAPEL QUE EST� SENDO UTILIZADO NA IMPRESSORA.
'|  O VALOR EST� EM POLEGADAS.
'|

Dim i As Single

    i = GetDeviceCaps(Printer.hdc, PHYSICALHEIGHT)
    i = i / get_printer_resolucao_dpi_y()
   
    get_printer_altura_fisica_pol = i
    
    
End Function

Public Function get_printer_margem_fisica_topo_pol() As Variant
'�______________________________________________________________________________________________________
'|
'|  RETORNA A MARGEM SUPERIOR M�NIMA ACEITA PELA IMPRESSORA, EM POLEGADAS.
'|

Dim i As Single

    i = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
    i = i / get_printer_resolucao_dpi_y()
   
    get_printer_margem_fisica_topo_pol = i
    
    
End Function



Public Function maior(num_1, num_2) As Variant
'�______________________________________________________________________________
'|
'|  RETORNA O MAIOR DOS 2 N�MEROS
'|

    If num_1 > num_2 Then
        maior = num_1
    Else
        maior = num_2
        End If
        
End Function

Public Function menor(num_1, num_2) As Variant
'�______________________________________________________________________________
'|
'|  RETORNA O MENOR DOS 2 N�MEROS
'|

    If num_1 < num_2 Then
        menor = num_1
    Else
        menor = num_2
        End If
        
End Function

Sub printer_quadro(ByVal xi As Single, ByVal yi As Single, ByVal xf As Single, ByVal yf As Single, Optional ByVal cor As Long = -1, Optional ByVal preencher_fundo As Boolean)
'�_________________________________________________________________________________________________________________________________________________________
'|
'|  ESTA FUN��O DEVE SER UTILIZADA SEMPRE QUE SE DESEJAR DESENHAR UM QUADRADO USANDO OS
'|  PAR�METROS 'B' OU 'BF' DO M�TODO PRINTER.LINE
'|
'|  O OBJETIVO DESTA FUN��O � EVITAR A PERDA DE RECURSOS DO GDI QUE OCORREM AO
'|  EXECUTAR OS COMANDOS NECESS�RIOS P/ CORRIGIR O BUG Q176634.
'|
'|  BUG (Q176634): AO DESENHAR UM QUADRADO USANDO O M�TODO LINE OU UM C�RCULO
'|      USANDO CIRCLE, O FUNDO N�O FICA TRANSPARENTE E ENCOBRE O QUE ESTIVER POR
'|      BAIXO.  A SOLU��O USADA PARA ESTE BUG, ENTRETANTO, PROVOCA A PERDA DE
'|      RECURSOS DO GDI (MAIS OU MENOS 0,125% A CADA EXECU��O), SENDO QUE,
'|      EXCEPCIONALMENTE NESTE CASO, OS RECURSOS S�O RECUPERADOS QUANDO O
'|      APLICATIVO � FECHADO.
'|      A PERDA DE RECURSOS OCORRE POR P�GINA EM QUE SE EXECUTA OS COMANDOS DE
'|      CORRE��O DO BUG, OU SEJA, SE OS COMANDOS FOREM EXECUTADOS 1 VEZ OU 1000
'|      VEZES EM UMA MESMA P�GINA, A PERDA SER� A MESMA.
'|


  '�QUADRADO PREENCHIDO
  '�~~~~~~~~~~~~~~~~~~~
    If preencher_fundo Then
      '�NESTE CASO, O BUG Q176634 N�O IMPORTA, J� QUE A COR DE FUNDO N�O SER� TRANSPARENTE
        If cor = -1 Then
          '�N�O ESPECIFICOU COR
            Printer.Line (xi, yi)-(xf, yf), , BF
        Else
          '�ESPECIFICOU COR
            Printer.Line (xi, yi)-(xf, yf), cor, BF
            End If
  
      '�BUG: Ap�s a fun��o LINE com op��o B ou BF, ocorre perda da configura��o de
      '�~~~~ cor de fundo.  Necess�rio restaurar transpar�ncia (ver MSDN: Q183163)
        SetBkMode Printer.hdc, vbFSTransparent
  
  
  '�QUADRADO VAZADO
  '�~~~~~~~~~~~~~~~
    Else
      '�NESTE CASO, EVITA O BUG Q176634 DESENHANDO CADA UMA DAS LINHAS SEPARADAMENTE
        If cor = -1 Then
          '�N�O ESPECIFICOU COR
            Printer.Line (xf, yf)-(xi, yf)
            Printer.Line (xi, yf)-(xi, yi)
            Printer.Line (xi, yi)-(xf, yi)
            Printer.Line (xf, yi)-(xf, yf)
        Else
          '�ESPECIFICOU COR
            Printer.Line (xf, yf)-(xi, yf), cor
            Printer.Line (xi, yf)-(xi, yi), cor
            Printer.Line (xi, yi)-(xf, yi), cor
            Printer.Line (xf, yi)-(xf, yf), cor
            End If
        End If
        
        
End Sub


Sub printer_assinala_x(ByVal xi As Single, ByVal yi As Single, ByVal xf As Single, ByVal yf As Single, Optional ByVal espessura As Integer = -1, Optional ByVal cor As Long = -1)
'�_________________________________________________________________________________________________________________________________________________________
'|
'|  DESENHA UM "X" NO QUADRADO INDICADO PELAS COORDENADAS DO PAR�METRO.
'|

Dim drawwidth_a As Single
Dim drawstyle_a As Integer

    drawwidth_a = Printer.DrawWidth
    drawstyle_a = Printer.DrawStyle
    
    If espessura <> -1 Then Printer.DrawWidth = espessura
    Printer.DrawStyle = vbSolid
    
    If cor = -1 Then
      '�N�O ESPECIFICOU COR
        Printer.Line (xi, yi)-(xf, yf)
        Printer.Line (xi, yf)-(xf, yi)
    Else
      '�ESPECIFICOU COR
        Printer.Line (xi, yi)-(xf, yf), cor
        Printer.Line (xi, yf)-(xf, yi), cor
        End If
        
    Printer.DrawWidth = drawwidth_a
    Printer.DrawStyle = drawstyle_a
    
End Sub



