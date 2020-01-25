<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<%
	Const adTypeBinary = 1
	Const adTypeText = 2
	Const chunkSize = 2048

	dim s
	dim strFileName, strFilePath
	dim intIdNFeEmitente
	dim blnForceDownload ' Força o download (dialog box "Save As") ou permite que seja aberto diretamente dentro do browser?

	strFileName=Request("file")

	s=Request("emitente")
	if IsNumeric(s) then intIdNFeEmitente=CInt(s) else intIdNFeEmitente=0
	strFilePath=obtem_path_xml_nfe(intIdNFeEmitente)

	s = Ucase(Trim(Request("force")))
	if (s="FALSE") Or (s="F") Or (s="NAO") Or (s="N") then blnForceDownload=False else blnForceDownload=True

	if Len(strFileName) > 0 then
		Call DownloadFile(strFileName, strFilePath, blnForceDownload)
		Response.End
		end if
	

' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

Sub DownloadFile(byval strFileName, byval strFilePath, byval blnForceDownload)
dim fso, objFile, strFullFileName
dim fileSize, blnBinary, strContentType
dim objStream, strAllFile, iSz
dim i
dim strVersaoIIS
	
	strVersaoIIS = Request.ServerVariables("SERVER_SOFTWARE")
	
	strFullFileName = strFilePath
	if Right(strFullFileName,1)<>"\" then strFullFileName=strFullFileName & "\"
	strFullFileName = strFullFileName & strFileName
	
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	
	If Not(fso.FileExists(strFullFileName)) Then
		Set fso=Nothing
		Err.Raise 20000, "Download de Arquivo", "Erro: arquivo não encontrado: " & strFullFileName
		Response.END
		End If
	
	Set objFile=fso.GetFile(strFullFileName)
	fileSize=objFile.Size
	Set objFile=Nothing
	
	' Check whether file is binary or not and get content type of the file (according to its extension)
	blnBinary=GetContentType(strFileName, strContentType)
	strAllFile=""
	
	'Read the file contents.
	'Force download? If so, add proper header:
	If blnForceDownload Then
		Response.AddHeader "Content-Disposition", "attachment; filename=" & strFileName
		end if

	If blnBinary Then
		Set objStream=Server.CreateObject("ADODB.Stream")
		
		'Added to breakup chunk
		Response.Buffer = False 
		
		'this might be long...
		Server.ScriptTimeout = 30000
		
		objStream.Open
		objStream.Type = 1 'adTypeBinary
		objStream.LoadFromFile strFullFileName
		
		'Added to breakup chunk
		iSz = objStream.Size
		if InStr(strVersaoIIS, "IIS/6.") <> 0 then Response.AddHeader "Content-Length", iSz
		Response.Charset = "UTF-8"
		Response.ContentType = strContentType
		For i = 1 To iSz \ chunkSize
			If Not Response.IsClientConnected Then Exit For
			Response.BinaryWrite objStream.Read(chunkSize)
			Next 
		
		If iSz Mod chunkSize > 0 Then 
			If Response.IsClientConnected Then 
				Response.BinaryWrite objStream.Read(iSz Mod chunkSize)
				End If 
			End If
			
  		objStream.Close
		Set objStream = Nothing
	Else
		Set objFile=fso.OpenTextFile(strFullFileName,1) 'forReading
		If Not(objFile.AtEndOfStream) Then
			strAllFile=objFile.ReadAll
			End If
		
		objFile.Close
		Set objFile=Nothing
		Response.Write(strAllFile)
		End If
	
	Set fso=Nothing
	if Response.Buffer then Response.Flush
	Response.END
End Sub

Function GetContentType(ByVal strName, ByRef ContentType)
'Return whether binary or not, put type into second parameter
dim strExtension
	
	strExtension="."&GetExtension(strName)
	Select Case Lcase(strExtension)
		Case ".asf"
			ContentType = "video/x-ms-asf"
			GetContentType=True
		Case ".avi"
			ContentType = "video/avi"
			GetContentType=True
		Case ".doc"
			ContentType = "application/msword"
			GetContentType=True
		Case ".zip"
			ContentType = "application/zip"
			GetContentType=True
		Case ".xls"
			ContentType = "application/vnd.ms-excel"
			GetContentType=True
		Case ".gif"
			ContentType = "image/gif"
			GetContentType=True
		Case ".jpg", ".jpeg"
			ContentType = "image/jpeg"
			GetContentType=True
		Case ".wav"
			ContentType = "audio/wav"
			GetContentType=True
		Case ".mp3"
			ContentType = "audio/mpeg3"
			GetContentType=True
		Case ".wma" 
			ContentType = "audio/wma"
			GetContentType=True
		Case ".mpg", ".mpeg"
			ContentType = "video/mpeg"
			GetContentType=True
		Case ".pdf"
			ContentType = "application/pdf"
			GetContentType=True
		Case ".rtf"
			ContentType = "application/rtf"
			GetContentType=True
		Case ".htm", ".html"
			ContentType = "text/html"
			GetContentType=False
		Case ".asp"
			ContentType = "text/asp"
			GetContentType=False
		Case ".txt"
			ContentType = "text/plain"
			GetContentType=False
		Case ".xml"
			ContentType = "text/xml"
			GetContentType=False
		Case Else
			'Handle All Other Files
			ContentType = "application/octet-stream"
			GetContentType=True
	End Select
End Function

Function GetExtension(strName)
dim arrTmp
	arrTmp=Split(strName, ".")
	GetExtension=arrTmp(UBound(arrTmp))
End Function

%>