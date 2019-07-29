<%@ language=VBScript%>
<%OPTION EXPLICIT%>
<%
	On Error GoTo 0
	Err.Clear

'	ENCERRA A SESSÃO
	Session("loja_atual")=" "
	Session("usuario_atual")=" "
	Session("senha_atual")=" "
	Session.Abandon

%>

<html>

<head>
	<title>LOJA</title>
	</head>

<script language="JavaScript" type="text/javascript">
	window.close();
</script>

</html>
