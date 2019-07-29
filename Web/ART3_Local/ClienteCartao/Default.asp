<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<%
	On Error GoTo 0
	Err.Clear

	dim strScript
	dim strUrlBase, strServerName, strLocalAddr, strURL
	strServerName = Ucase(Trim(Request.ServerVariables("server_name")))
	strLocalAddr = Trim(Request.ServerVariables("local_addr"))
	strURL = request.ServerVariables("URL")
	strUrlBase = SITE_CLIENTE_URL_BASE

'	LEMBRANDO QUE O ENDEREÇO 'PAGAMENTO.BONSHOP.COM.BR' É O SITE ALTERNATIVO LOCALIZADO NO SERVIDOR DO SISTEMA E QUE NÃO POSSUI CERTIFICADO SSL INSTALADO
	if (strServerName = "XXXDNN2581") Or (strServerName = "LOCALHOST") Or (strServerName = "PAGAMENTO.BONSHOP.COM.BR") then strUrlBase = ""
	if (strServerName = "HOMOLOGACAO.CENTRAL85.COM.BR") then strUrlBase = ""
	if (strServerName = "APPSERVER") Or (strServerName = "WIN2008R2") Or (strServerName = "WIN2008R2BS") then strUrlBase = ""

	if Instr(Ucase(strURL), "HOMOLOG") > 0 then strUrlBase = ""
	if Instr(Ucase(strURL), "TEST") > 0 then strUrlBase = ""

	if strLocalAddr <> "" then
		if Instr(strLocalAddr, "182.168.0.") > 0 then strUrlBase = ""
		if Instr(strLocalAddr, "182.168.2.") > 0 then strUrlBase = ""
		if Instr(strLocalAddr, "192.168.0.") > 0 then strUrlBase = ""
		if Instr(strLocalAddr, "192.168.2.") > 0 then strUrlBase = ""
		if Instr(strLocalAddr, "127.0.0.") > 0 then strUrlBase = ""
		if strLocalAddr = "::1" then strUrlBase = ""
		end if
	
	strScript = "<script language='JavaScript'>" & chr(13) & _
				"	var urlDestino = '" & strUrlBase & "Id.asp';" & chr(13) & _
				"	window.name = '" & SITE_CLIENTE_TITULO_JANELA & "';" & chr(13) & _
				"</script>" & chr(13)
%>


<%=DOCTYPE_LEGADO%>

<html>
<head>
<title><%=SITE_CLIENTE_TITULO_JANELA%></title>
</head>

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<% =strScript %>

<script language="JavaScript" type="text/javascript">
var blnVersaoNavegadorOk = true;
</script>

<script language="JavaScript" type="text/javascript">
	$(document).ready(function() {
		if (!blnVersaoNavegadorOk) {
			alert("Este navegador não é suportado para o acesso à área restrita!!\nPor favor, utilize o Internet Explorer versão 7 ou superior!!");
			$("#divMsg").html("Este navegador não é suportado para o acesso à área restrita!!<br />Por favor, utilize o Internet Explorer versão 7 ou superior!!");
			$("#divMsg").css({ 'color': 'red' });
			$("#divMsg").show();
		} else {
			window.location = urlDestino;
			$("#divMsg").html("O navegador será automaticamente redirecionado para um ambiente seguro em instantes...<br />Ou clique <a href='" + urlDestino + "'>aqui</a> para ser redirecionado agora.");
			$("#divMsg").show();
		}
	});
</script>

<style type="text/css">
.DivMsg
{
	border:2px double;
	padding:10px;
	font-family:Arial;
	font-weight:bold;
	font-size:10pt;
	text-align:center;
	color:Black;
	background-color:#F5F5F5;
	width:75%;
}
</style>

<body>
	<center>
	<div id="divMsg" class="C DivMsg" style="display:none;"></div>
	</center>
</body>

</html>
