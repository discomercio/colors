<%@ language=VBScript%>
<%OPTION EXPLICIT%>
<%
	On Error GoTo 0
	Err.Clear

'	LEMBRANDO QUE NAS P�GINAS ACESSADAS DIRETAMENTE PELOS CLIENTES PARA FAZER O PAGAMENTO N�O SE DEVE USAR 'SESSION',
'	J� QUE OCORRERAM V�RIOS CASOS DE CLIENTES QUE N�O CONSEGUIRAM INICIAR A SESS�O (PROVAVELMENTE POR PROBLEMAS NA CONFIGURA��O DE COOKIES OU MESMO DEVIDO A ANTIVIRUS)

'	ENCERRA A SESS�O
	Session.Contents.RemoveAll
	Session.Abandon
	
	Response.Redirect("../ClienteCartao/Id.asp")
%>

<html>

<head>
	<title><%=SITE_CLIENTE_TITULO_JANELA%></title>
	</head>

<script language="JavaScript" type="text/javascript">
	window.close();
</script>

</html>
