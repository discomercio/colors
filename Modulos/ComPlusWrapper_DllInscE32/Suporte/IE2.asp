<%@ Language=VBScript %>
<%OPTION EXPLICIT%>

<%
'     ===============================
'	  I P . A S P
'     ===============================
'
'
'	  S E R V E R   S I D E   S C R I P T I N G
'
'      SSSSSSS   EEEEEEEEE  RRRRRRRR   VVV   VVV  IIIII  DDDDDDDD    OOOOOOO   RRRRRRRR
'     SSS   SSS  EEE        RRR   RRR  VVV   VVV   III   DDD   DDD  OOO   OOO  RRR   RRR
'      SSS       EEE        RRR   RRR  VVV   VVV   III   DDD   DDD  OOO   OOO  RRR   RRR
'       SSSS     EEEEEE     RRRRRRRR   VVV   VVV   III   DDD   DDD  OOO   OOO  RRRRRRRR
'          SSS   EEE        RRR RRR     VVV VVV    III   DDD   DDD  OOO   OOO  RRR RRR
'     SSS   SSS  EEE        RRR  RRR     VVVVV     III   DDD   DDD  OOO   OOO  RRR  RRR
'      SSSSSSS   EEEEEEEEE  RRR   RRR     VVV     IIIII  DDDDDDDD    OOOOOOO   RRR   RRR


' _____________________________________________________________________________________________
'
'						I N I C I A L I Z A     P Á G I N A     A S P
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear

	dim resultado
	dim strIE, strUF

	strIE = Trim(Request("IE"))
	strUF = Trim(Request("UF"))

	if strIE = "" then
		Response.Write "Informe no parâmetro 'IE' o número da Inscrição Estadual a ser validado!"
		Response.End
		end if

	if strUF = "" then
		Response.Write "Informe no parâmetro 'UF' a sigla da UF do número da Inscrição Estadual a ser validado!"
		Response.End
		end if

	dim objIE
	set objIE = CreateObject("ComPlusWrapper_DllInscE32.ComPlusWrapper_DllInscE32")

	resultado = objIE.ConsisteInscricaoEstadual(strIE, strUF)
	Response.Write "IE: " & strIE & ", UF: " & strUF & ", ConsisteInscricaoEstadual() = " & resultado

	Response.Write "<br>"

	if objIE.isInscricaoEstadualOk(strIE, strUF) then
		Response.Write "IE: " & strIE & ", UF: " & strUF & ", isInscricaoEstadualOk() = OK"
	else
		Response.Write "IE: " & strIE & ", UF: " & strUF & ", isInscricaoEstadualOk() = INVÁLIDO"
		end if

	set objIE = Nothing
	Response.End
	
%>




<%
'	  C L I E N T   S I D E   S C R I P T I N G
'
'      CCCCCCC   LLL        IIIII  EEEEEEEEE  NNN   NNN  TTTTTTTTT EEEEEEEEE
'     CCC   CCC  LLL         III   EEE        NNNN  NNN     TTT    EEE
'     CCC        LLL         III   EEE        NNNNN NNN     TTT    EEE
'     CCC        LLL         III   EEEEEE     NNN NNNNN     TTT    EEEEEE
'     CCC        LLL         III   EEE        NNN  NNNN     TTT    EEE
'     CCC   CCC  LLL   LLL   III   EEE        NNN   NNN     TTT    EEE
'      CCCCCCC   LLLLLLLLL  IIIII  EEEEEEEEE  NNN   NNN     TTT    EEEEEEEEE 
%>

<html>

<head>
	<title>Consiste número de IE</title>
</head>


<!-- C A S C A D I N G   S T Y L E   S H E E T      CCCCCCC    SSSSSSS    SSSSSSS     CCC   CCC  SSS   SSS  SSS   SSS     CCC        SSS        SSS     CCC         SSSS       SSSS     CCC            SSSS       SSSS     CCC   CCC  SSS   SSS  SSS   SSS      CCCCCCC    SSSSSSS    SSSSSSS-->

<body>
<center>
<br>

<h1></h1>

</center>


</body>
</html>


