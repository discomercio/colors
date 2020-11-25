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

	dim i
	dim resultado
	dim strExecutaValidacao, strIE, strUF
	dim c, strIENormalizado
	dim objIE

	strExecutaValidacao = Ucase(Trim(Request("executa_validacao")))
	strIE = Ucase(Trim(Request("IE")))
	strUF = Ucase(Trim(Request("UF")))

	if strExecutaValidacao = "S" then
		Response.Write "<body onload='bVoltar.focus();'>"
		Response.Write "<center>"

		if strIE = "" then
			Response.Write "Informe no parâmetro 'IE' o número da Inscrição Estadual a ser validado!"
			Response.End
			end if

		if strUF = "" then
			Response.Write "Informe no parâmetro 'UF' a sigla da UF do número da Inscrição Estadual a ser validado!"
			Response.End
			end if

		strIENormalizado = strIE
		if strIENormalizado <> "ISENTO" then
			strIENormalizado = ""
			for i = 1 to len(strIE)
				c = Mid(strIE, i, 1)
				if IsNumeric(c) then strIENormalizado = strIENormalizado & c
				next
			end if

		set objIE = CreateObject("ComPlusWrapper_DllInscE32.ComPlusWrapper_DllInscE32")

		resultado = objIE.ConsisteInscricaoEstadual(strIENormalizado, strUF)
		Response.Write "IE: " & strIE & ", UF: " & strUF & ", ConsisteInscricaoEstadual() = " & resultado

		Response.Write "<br>"

		if objIE.isInscricaoEstadualOk(strIENormalizado, strUF) then
			Response.Write "IE: " & strIE & ", UF: " & strUF & ", isInscricaoEstadualOk() = OK"
		else
			Response.Write "IE: " & strIE & ", UF: " & strUF & ", isInscricaoEstadualOk() = INVÁLIDO"
			end if

		set objIE = Nothing

		Response.Write "<br><br>"
		Response.Write "<a href='IE-Consiste.asp' name='bVoltar'>VOLTAR</a>"
		Response.Write "</center>"
		Response.Write "</body>"

		Response.End
		end if

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

<body onload="f.ie.focus();">
<center>
<br>

<form action="IE-Consiste.asp" method="POST" name="f">
<input type="HIDDEN" name="executa_validacao" value="S">

<table>
<tr>
	<td align="right">IE</td>
	<td><input type="TEXT" name="ie"></td>
</tr>
<tr>
	<td align="right">UF</td>
	<td><input type="TEXT" name="uf"></td>
</tr>
<tr>
	<td colspan=2 align="center"><input type="SUBMIT"></td>
</tr>
</table>
</form>

</center>


</body>
</html>


