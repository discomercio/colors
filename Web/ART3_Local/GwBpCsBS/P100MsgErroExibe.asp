<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Braspag.asp"    -->

<%
'     ===========================================
'	  P100MsgErroExibe.asp
'     ===========================================
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

'	OBS: A P�GINA QUE EXIBE A MENSAGEM DE ERRO � ACIONADA ATRAV�S DOS SEGUINTES PASSOS:
'	~~~~
'	1) A P�GINA QUE EXECUTA A TRANSA��O C/ A BRASPAG VIA WEB SERVICE RESULTA EM ALGUM
'		ERRO DURANTE O PROCESSAMENTO.
'	2) A P�GINA INTERMEDI�RIA DE EXIBI��O DA MENSAGEM DE ERRO PREPARA OS DADOS EM
'		CAMPOS HIDDEN DE UM FORM, L� E APAGA OS DADOS ARMAZENADOS ATRAV�S DA SESSION
'		E, POR FIM, FAZ UM SUBMIT() P/ A P�GINA FINAL DE EXIBI��O.
'	3) COM ESTE MECANISMO, SE O USU�RIO ACIONAR O REFRESH NA P�GINA DE EXIBI��O, EVITAM-SE
'		OS SEGUINTES PROBLEMAS:
'		A) REEXECUTAR O PROCESSAMENTO DA TRANSA��O.
'		B) PARA OS DADOS ARMAZENADOS NA SESSION, A PARTIR DA 2� EXECU��O OS DADOS J� TERIAM
'			SIDO APAGADOS.

	On Error GoTo 0
	Err.Clear

	dim s, usuario, loja
	usuario = BRASPAG_USUARIO_CLIENTE

	dim pedido_selecionado, id_pedido_base
	pedido_selecionado = Trim(Request("pedido_selecionado"))
	id_pedido_base = retorna_num_pedido_base(pedido_selecionado)

	dim cnpj_cpf_selecionado
	cnpj_cpf_selecionado = retorna_so_digitos(Request("cnpj_cpf_selecionado"))

	dim strIdPagtoGwPag
	strIdPagtoGwPag = Trim(Request("idPagtoGwPag"))

	dim alerta
	alerta = Trim(Request("alerta"))

	dim strScriptWindowName
	strScriptWindowName = _
				"<script language='JavaScript'>" & chr(13) & _
				"	window.name = '" & SITE_CLIENTE_TITULO_JANELA & "';" & chr(13) & _
				"</script>" & chr(13)
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title><%=SITE_CLIENTE_TITULO_JANELA%></title>
	</head>


<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__SSL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<% =strScriptWindowName %>

<script language="JavaScript" type="text/javascript">
function Navega(url) {
	window.location.href = url;
}
</script>



<!-- C A S C A D I N G   S T Y L E   S H E E T

	 CCCCCCC    SSSSSSS    SSSSSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	CCC        SSS        SSS
	CCC         SSSS       SSSS
	CCC            SSSS       SSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	 CCCCCCC    SSSSSSS    SSSSSSS
-->

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__E_LOGO_TOP_BS_CSS%>" Rel="stylesheet" Type="text/css">

<style type="text/css">
body::before
{
	content: '';
	border: none;
	margin-top: 0px;
	margin-bottom: 0px;
	padding: 0px;
}
</style>



<!-- ************************************************************ -->
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();window.status='';" link=#000000 alink=#000000 vlink=#000000>
<center>

<table id="tbl_logotipo_bonshop" width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td align="center"><img alt="<%=SITE_CLIENTE_HEADER__ALT_IMG_TEXT%>" src="../imagem/<%=SITE_CLIENTE_HEADER__LOGOTIPO%>" /></td>
	</tr>
</table>
<table id="pagina_tbl_cabecalho" cellspacing="0px" cellpadding="0px">
	<tbody>
		<tr style="height:78px;">
			<td id="topo_verde" colspan="3">
				<div id="moldura_do_letreiro">
					<div id="letreiro_div" style="display:block;"></div>
				</div>
				<div id="telefone"></div>
			</td>
		</tr>
		<tr>
			<td id="topo_azul" colspan="3">&nbsp;</td>
		</tr>
	</tbody>
</table>

<br />

<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" href="javascript:Navega('../ClienteCartao/Id.asp')"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>

<% if SITE_CLIENTE_EXIBIR_LOGO_SSL then %>
<script language="JavaScript" type="text/javascript">
	logo_ssl_corner("../imagem/ssl/ssl_corner.gif");
</script>
<% end if %>

</body>

</html>
