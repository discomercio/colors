<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===================================================================
'	  P E D I D O P R E D E V O L U Ç A O M E N S A G E M N O V A . A S P
'     ===================================================================
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

	dim usuario, loja, pedido_selecionado, id_devolucao
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
	id_devolucao = Trim(request("id_devolucao"))
	if (id_devolucao = "") then Response.Redirect("aviso.asp?id=" & ERR_PARAMETRO_OBRIGATORIO_NAO_ESPECIFICADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    dim url_origem
    url_origem = Trim(request("url_origem"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ESCREVER_MSG, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
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
	<title>LOJA<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function calcula_tamanho_restante() {
var f, s;
	f = fPED;
	s = "" + fPED.c_texto.value;
	f.c_tamanho_restante.value = MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS - s.length;
}

function fPEDPreDevolucaoMensagemNovaConfirma(f) {
var s;

//  Texto contendo uma mensagem referente à esta pré-devolução
	s = "" + f.c_texto.value;
	if (s.length == 0) {
		alert('É necessário escrever o texto da mensagem referente a esta pré-devolução!!');
		f.c_texto.focus();
		return;
		}
	if (s.length > MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS) {
		alert('Conteúdo do texto excede em ' + (s.length - MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS) + ' caracteres o tamanho máximo de ' + MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS + '!!');
		f.c_texto.focus();
		return;
		}

	dCONFIRMA.style.visibility="hidden";
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

<body onload="fPED.c_texto.focus();">
<center>

<form id="fPED" name="fPED" method="post" action="PedidoPreDevolucaoMensagemNovaConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="id_devolucao" id="id_devolucao" value='<%=id_devolucao%>'>
<input type="hidden" name="url_origem" id="url_origem" value="<%=url_origem%>" />

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="left" valign="bottom"><p class="PEDIDO">Mensagem para Pré-Devolução do Pedido <%=pedido_selecionado%></p></td>
</tr>
</table>
<br>

<table>
<tr>
	<td align="right" valign="bottom">
		<span class="PLLd">Tamanho restante:</span><input name="c_tamanho_restante" id="c_tamanho_restante" tabindex=-1 readonly class="TA" style="width:35px;text-align:right;" value='<%=Cstr(MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS)%>' />
	</td>
</tr>
<tr>
	<td>
	<table class="Q" style="width:649px;" cellSpacing="0">
		<tr>
			<td><p class="Rf">MENSAGEM</p>
				<textarea name="c_texto" id="c_texto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_MENSAGEM_OCORRENCIAS_EM_PEDIDOS)%>" 
					style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS);" onblur="this.value=trim(this.value);calcula_tamanho_restante();"
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
	<td><a name="bCANCELA" id="bCANCELA" href="javascript:history.back()" title="cancela o cadastramento da nova mensagem">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDPreDevolucaoMensagemNovaConfirma(fPED)" title="grava a nova mensagem">
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