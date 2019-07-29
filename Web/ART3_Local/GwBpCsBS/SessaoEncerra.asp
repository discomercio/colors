<%@ language=VBScript%>
<%OPTION EXPLICIT%>
<%
	On Error GoTo 0
	Err.Clear

'	LEMBRANDO QUE NAS PÁGINAS ACESSADAS DIRETAMENTE PELOS CLIENTES PARA FAZER O PAGAMENTO NÃO SE DEVE USAR 'SESSION',
'	JÁ QUE OCORRERAM VÁRIOS CASOS DE CLIENTES QUE NÃO CONSEGUIRAM INICIAR A SESSÃO (PROVAVELMENTE POR PROBLEMAS NA CONFIGURAÇÃO DE COOKIES OU MESMO DEVIDO A ANTIVIRUS)

'	ENCERRA A SESSÃO
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
