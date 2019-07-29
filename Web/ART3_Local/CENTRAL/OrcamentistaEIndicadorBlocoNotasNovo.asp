<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  P E D I D O B L O C O N O T A S N O V O . A S P
'     ===============================================
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
'
'
'	REVISADO P/ IE10


	On Error GoTo 0
	Err.Clear

	dim usuario, apelido_selecionado, url_origem
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	apelido_selecionado = Trim(request("id_selecionado"))
	if (apelido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTISTA_INDICADOR_NAO_ESPECIFICADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))	

	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	url_origem = Request("url_origem")
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
	<title>LOJA - Novo Bloco de Notas Indicador</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var MAX_TAM_MENSAGEM_BLOCO_NOTAS_RELACIONAMENTO = <%=MAX_TAM_MENSAGEM_BLOCO_NOTAS_RELACIONAMENTO%>;
function calcula_tamanho_restante() {
	var f, s;
	f = fCAD;
	s = "" + fCAD.c_mensagem.value;
	f.c_tamanho_restante.value = MAX_TAM_MENSAGEM_BLOCO_NOTAS_RELACIONAMENTO - s.length;
}

function fCADBlocoNotasNovoConfirma(f) {
var s;

	s = "" + f.c_mensagem.value;
	if (s.length == 0) {
		alert('É necessário escrever o texto da mensagem!!');
		f.c_mensagem.focus();
		return;
		}
	if (s.length > MAX_TAM_MENSAGEM_BLOCO_NOTAS_RELACIONAMENTO) {
		alert('Conteúdo da mensagem excede em ' + (s.length - MAX_TAM_MENSAGEM_BLOCO_NOTAS_RELACIONAMENTO) + ' caracteres o tamanho máximo de ' + MAX_TAM_MENSAGEM_BLOCO_NOTAS_RELACIONAMENTO + '!!');
		f.c_mensagem.focus();
		return;
		}


		dCONFIRMA.style.visibility = "hidden";
		f.url_origem.value = "<%=url_origem%>";
	window.status = "Aguarde ...";
	f.submit();
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<body onload="fCAD.c_mensagem.focus();">
<center>

<form id="fCAD" name="fCAD" method="post" action="OrcamentistaEIndicadorBlocoNotasNovoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="apelido_selecionado" id="apelido_selecionado" value='<%=apelido_selecionado%>'>
<input type="hidden" name="url_origem" id="url_origem" value"<%=url_origem%>" />

<!--  I D E N T I F I C A Ç Ã O   D O   O R C A M E N T I S T A  -->  
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="left" valign="bottom"><p style="font-weight:bold;font-size:18pt;font-family:Times New Roman, serif;font-style:italic">Bloco de Notas</p></td>
	<td align="right" valign="bottom"><p style="font-weight:bold;font-size:14pt;font-family:Times New Roman, serif;font-style:italic">Indicador <%=apelido_selecionado%></p></td>
</tr>
</table>
<br>

<table>
<tr>
	<td align="right" valign="bottom">
		<span class="PLLd">Tamanho restante:</span><input name="c_tamanho_restante" id="c_tamanho_restante" tabindex=-1 readonly class="TA" style="width:35px;text-align:right;" value="<%=MAX_TAM_MENSAGEM_BLOCO_NOTAS_RELACIONAMENTO%>" />
	</td>
</tr>
<tr>
	<td>
	<table class="Q" style="width:649px;" cellSpacing="0">
		<tr>
			<td><p class="Rf">MENSAGEM</p>
				<textarea name="c_mensagem" id="c_mensagem" class="PLLe" rows="10" 
					style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_MENSAGEM_BLOCO_NOTAS_RELACIONAMENTO);" onblur="this.value=trim(this.value);calcula_tamanho_restante();"
					onkeyup="calcula_tamanho_restante();"
					></textarea>
			</td>
		</tr>
	</table>
	</td>
</tr>

</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a name="bCANCELA" id="bCANCELA" href="javascript:history.back()" title="cancela o cadastramento de nova mensagem no bloco de notas">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="RIGHT"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fCADBlocoNotasNovoConfirma(fCAD)" title="grava a mensagem no bloco de notas">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>