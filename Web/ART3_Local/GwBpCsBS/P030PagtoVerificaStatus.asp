<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Braspag.asp"    -->
<!-- #include file = "../global/BraspagCS.asp"    -->

<%
'     ===============================================
'	  P030PagtoVerificaStatus.asp
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


	On Error GoTo 0
	Err.Clear

	dim s, usuario, loja, pedido_selecionado, id_pedido_base

	usuario = BRASPAG_USUARIO_CLIENTE

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	id_pedido_base = retorna_num_pedido_base(pedido_selecionado)
	
	dim cnpj_cpf_selecionado
	cnpj_cpf_selecionado = retorna_so_digitos(Request("cnpj_cpf_selecionado"))
	
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, t_PAG_PAYMENT, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	If Not cria_recordset_otimista(t_PAG_PAYMENT, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim alerta
	alerta = ""

	dim mensagem, erro_fatal, prosseguir_automaticamente
	erro_fatal = False
	prosseguir_automaticamente = True
	mensagem = ""
	
	dim r_pedido
	if Not le_pedido(id_pedido_base, r_pedido, msg_erro) then
		alerta = msg_erro
	else
		loja = r_pedido.loja
		end if
	
'	PESQUISA POR PAGAMENTOS JÁ EFETUADOS
'	LEMBRANDO QUE UM PEDIDO PODE SER PAGO USANDO VÁRIOS CARTÕES E O CLIENTE PODE TER REALIZADO VÁRIAS TENTATIVAS
	dim msg_pagto_anterior, qtde_pagto_anterior
	msg_pagto_anterior = ""
	qtde_pagto_anterior = 0
	s = "SELECT " & _
			" trx_TX_data_hora," & _
			"t_PAGTO_GW_PAG_PAYMENT.*" & _
		" FROM t_PAGTO_GW_PAG" & _
			" INNER JOIN t_PAGTO_GW_PAG_PAYMENT ON (t_PAGTO_GW_PAG.id = t_PAGTO_GW_PAG_PAYMENT.id_pagto_gw_pag)" & _
		" WHERE" & _
			" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')" & _
			" AND " & _
			"(" & _
				"(ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "')" & _
				" OR " & _
				"(ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA & "')" & _
			")" & _
		" ORDER BY" & _
			" id"
	if t_PAG_PAYMENT.State <> 0 then t_PAG_PAYMENT.Close
	t_PAG_PAYMENT.open s, cn
	do while Not t_PAG_PAYMENT.EOF
		prosseguir_automaticamente = False
		qtde_pagto_anterior = qtde_pagto_anterior + 1
			
		if msg_pagto_anterior <> "" then
			msg_pagto_anterior = msg_pagto_anterior & _
								"<tr>" & chr(13) & _
								"<td align='left'>&nbsp;</td>" & chr(13) & _
								"</tr>" & chr(13)
			end if
			
		msg_pagto_anterior = msg_pagto_anterior & _
					"<tr>" & chr(13) & _
					"<td align='left'>" & chr(13) & _
					"	<table border='0' cellpadding='0' class='N TblMsg'>" & chr(13) & _
					"	<tr><td nowrap class='Td1'>Bandeira:&nbsp;</td><td align='left'>" & BraspagDescricaoBandeira(Trim("" & t_PAG_PAYMENT("bandeira"))) & "</td></tr>" & chr(13) & _
					"	<tr><td nowrap class='Td1'>Valor:&nbsp;</td><td align='left'>" & SIMBOLO_MONETARIO & " " & formata_moeda(t_PAG_PAYMENT("valor_transacao")) & "</td></tr>" & chr(13) & _
					"	<tr><td nowrap class='Td1'>Opção de pagamento:&nbsp;</td><td align='left'>" & BraspagDescricaoParcelamento(Trim("" & t_PAG_PAYMENT("req_PaymentDataRequest_PaymentPlan")), Trim("" & t_PAG_PAYMENT("req_PaymentDataRequest_NumberOfPayments")), t_PAG_PAYMENT("valor_transacao")) & "</td></tr>" & chr(13) & _
					"	<tr><td nowrap class='Td1'>Data:&nbsp;</td><td align='left'>" & formata_data_hora(t_PAG_PAYMENT("trx_TX_data_hora")) & "</td></tr>" & chr(13) & _
					"	<tr><td nowrap class='Td1' style='vertical-align:middle;'>Consultar comprovante:&nbsp;</td><td align='left'><a href='javascript:ComprovanteConsulta(" & chr(34) & Trim("" & t_PAG_PAYMENT("id_pagto_gw_pag")) & chr(34) & ");'><img src='../imagem/doc_preview_22.png' /></a></td></tr>" & chr(13) & _
					"	</table>" & chr(13) & _
					"</td>" & chr(13) & _
					"</tr>" & chr(13)
		
		t_PAG_PAYMENT.MoveNext
		loop

'	ENCONTROU PAGAMENTOS EFETUADOS ANTERIORMENTE?
	if qtde_pagto_anterior > 0 then
		if mensagem <> "" then mensagem = mensagem & "<br>" & chr(13)
		if qtde_pagto_anterior > 1 then
			mensagem = mensagem & "O pedido " & id_pedido_base & " já originou as seguintes transações de pagamento <span style='color:Green;'>bem-sucedidas</span>:"
		else
			mensagem = mensagem & "O pedido " & id_pedido_base & " já originou a seguinte transação de pagamento <span style='color:Green;'>bem-sucedida</span>:"
			 end if
		mensagem = mensagem & _
					"<br>" & chr(13) & _
					"<table cellpadding='0' cellspacing='0' border='0'>" & chr(13) & _
					msg_pagto_anterior & _
					"</table>" & chr(13)
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

<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title><%=SITE_CLIENTE_TITULO_JANELA%></title>
	</head>


<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__SSL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function ComprovanteConsulta(_id_pagto_gw_pag) {
	fComprovante.id_pagto_gw_pag.value = _id_pagto_gw_pag;
	fComprovante.action = "P110ComprovanteConsulta.asp";
	window.status = "Aguarde ...";
	fComprovante.submit();
}
function fPEDConsulta() {
	fPED.action = "../ClienteCartao/PedidoConsulta.asp";
	window.status = "Aguarde ...";
	fPED.submit();
}
function fPAGTOConclui( f ) {
	f.action = "P040PagtoOpcoes.asp";
	window.status = "Aguarde ...";
	f.submit();
}
function fPAGTOConcluiExec( ) {
	fPAGTOConclui(fPAGTO);
}
function fPAGTOConcluiDivExec( ) {
	dCONFIRMA.style.visibility='hidden';
	fPAGTOConclui(fPAGTO);
}
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
<link href="<%=URL_FILE__E_LOGO_TOP_BS_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
body::before
{
	content: '';
	border: none;
	margin-top: 0px;
	margin-bottom: 0px;
	padding: 0px;
}
.TblMsg {
	margin-top: 6px;
	color: navy;
	}
.Td1 {
	vertical-align: top;
	text-align: right;
	}
</style>


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();window.status='';">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ****************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESUMO DO PEDIDO  ***************** -->
<!-- ****************************************************************** -->
<body 
<% if prosseguir_automaticamente then %>
onload="setTimeout('fPAGTOConcluiExec()', 100)"
<% else %>
onload="window.status='';"
<% end if %>
>
<center>

<table class="notPrint" id="tbl_logotipo_bonshop" width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td align="center"><img alt="<%=SITE_CLIENTE_HEADER__ALT_IMG_TEXT%>" src="../imagem/<%=SITE_CLIENTE_HEADER__LOGOTIPO%>" /></td>
	</tr>
</table>
<table class="notPrint" id="pagina_tbl_cabecalho" cellspacing="0px" cellpadding="0px">
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

<form id="fPED" name="fPED" method="post">
<input type="hidden" name='pedido_selecionado' value='<%=pedido_selecionado%>'>
<input type="hidden" name='cnpj_cpf_selecionado' value='<%=cnpj_cpf_selecionado%>'>
</form>

<form id="fComprovante" name="fComprovante" method="post">
<input type="hidden" name="id_pagto_gw_pag" value="" />
<input type="hidden" name='pedido_selecionado' value='<%=pedido_selecionado%>'>
<input type="hidden" name='cnpj_cpf_selecionado' value='<%=cnpj_cpf_selecionado%>'>
</form>

<form id="fPAGTO" name="fPAGTO" method="post">
<input type="hidden" name='pedido_selecionado' value='<%=pedido_selecionado%>'>
<input type="hidden" name='cnpj_cpf_selecionado' value='<%=cnpj_cpf_selecionado%>'>


<% if prosseguir_automaticamente then %>
	<br>
	<span class="PEDIDO">Aguarde, verificando o pedido...</span>
<% else %>
	<!-- **********  MENSAGEM DE ERRO  ********** -->
	<center>
	<br>
	<!--  T E L A  -->
	<p class="T">A T E N Ç Ã O</p>
	<div style="width:600px;font-weight:bold;" align="center"><span style='margin:5px 2px 5px 2px;'><%=mensagem%></span></div>
	</center>
<% end if %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>


<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr>
<% if (Not prosseguir_automaticamente) And (Not erro_fatal) then %>
	<td align="left"><a name="bVOLTAR" href="javascript:fPEDConsulta()" title="volta para a página do pedido">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right"><div name='dCONFIRMA' id='dCONFIRMA'><a name="bCONFIRMA" href="javascript:fPAGTOConcluiDivExec()" title="Continua com o processo de pagamento">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
<% else %>
	<td align="center"><a name="bVOLTAR" href="javascript:fPEDConsulta()" title="volta para a página do pedido">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
<% end if %>
</tr>
</table>

</form>

</center>

<% if SITE_CLIENTE_EXIBIR_LOGO_SSL then %>
<script language="JavaScript" type="text/javascript">
	logo_ssl_corner("../imagem/ssl/ssl_corner.gif");
</script>
<% end if %>

</body>

<% end if %>

</html>


<%
	if t_PAG_PAYMENT.State <> 0 then t_PAG_PAYMENT.Close
	set t_PAG_PAYMENT = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>