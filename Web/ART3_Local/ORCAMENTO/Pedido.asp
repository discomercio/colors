<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<%
'     ===========================================
'	  P E D I D O . A S P
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
'
'
'	REVISADO P/ IE10


	On Error GoTo 0
	Err.Clear

	dim s, usuario, loja, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim url_back
	url_back = Trim(request("url_back"))
	
	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then 
		if url_back <> "" then
			Response.Redirect("Resumo.asp")
		else
			Response.Redirect("Aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
			end if
		end if
	
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	if Len(pedido_selecionado) > TAM_MAX_ID_PEDIDO then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_INVALIDO)
	
	dim i, n, s_fabricante, s_produto, s_descricao, s_descricao_html, s_qtde, s_preco_lista, s_desc_dado
	dim s_vl_unitario, s_vl_TotalItem, m_TotalItem, m_TotalDestePedido, m_TotalItemComRA, m_TotalDestePedidoComRA
	dim s_preco_NF, m_TotalFamiliaParcelaRA
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("Aviso.asp?id=" & ERR_CONEXAO)

	dim r_pedido, v_item, alerta
	dim blnOrcamentistaOuIndicadorOK
	alerta=""
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
	else
		blnOrcamentistaOuIndicadorOK = False
		if Trim(r_pedido.orcamentista) = usuario then blnOrcamentistaOuIndicadorOK = True 
		if (Trim(r_pedido.indicador) = usuario) and (Trim(r_pedido.orcamentista) = usuario or Trim(r_pedido.orcamentista) = "") then blnOrcamentistaOuIndicadorOK = True 
		if Not blnOrcamentistaOuIndicadorOK then Response.Redirect("Aviso.asp?id=" & ERR_PEDIDO_INVALIDO)
		if Not le_pedido_item(pedido_selecionado, v_item, msg_erro) then alerta = msg_erro
		end if

	dim blnTemRA
	blnTemRA = False
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			if Trim("" & v_item(i).produto) <> "" then
				if v_item(i).preco_NF <> v_item(i).preco_venda then
					blnTemRA = True
					exit for
					end if
				end if
			next
		end if
	
	dim n_offset_tabela_ocorrencia, blnHaOcorrenciaEmAberto
	dim s_aux, s2, s3, s4, r_loja, r_cliente, s_cor, s_falta, v_pedido
	dim v_disp
	dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF
	dim vl_saldo_a_pagar, s_vl_saldo_a_pagar, st_pagto
	dim v_item_devolvido, s_devolucoes
	dim pedido_splitado
	dim v_pedido_perda, s_perdas, vl_total_perdas
	s_devolucoes = ""
	pedido_splitado = False
	s_perdas = ""
	vl_total_perdas = 0

	if alerta = "" then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then
			redim v_disp(Ubound(v_item))
			for i=Lbound(v_disp) to Ubound(v_disp)
				set v_disp(i) = New cl_ITEM_STATUS_ESTOQUE
				v_disp(i).pedido		= v_item(i).pedido
				v_disp(i).fabricante	= v_item(i).fabricante
				v_disp(i).produto		= v_item(i).produto
				v_disp(i).qtde			= v_item(i).qtde
				next
			
			if Not estoque_verifica_status_item(v_disp, msg_erro) then Response.Redirect("Aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			end if
			
	'	OBT�M OS N�MEROS DE PEDIDOS QUE COMP�EM ESTA FAM�LIA DE PEDIDOS
		if Not recupera_familia_pedido(pedido_selecionado, v_pedido, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		n=0
		for i=Lbound(v_pedido) to Ubound(v_pedido)
			if Trim(v_pedido(i))<>"" then n=n+1
			next
			
		if n > 1 then pedido_splitado = True
		
	'	OBT�M OS VALORES A PAGAR, J� PAGO E O STATUS DE PAGAMENTO (PARA TODA A FAM�LIA DE PEDIDOS)
		if Not calcula_pagamentos(pedido_selecionado, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		m_TotalFamiliaParcelaRA = vl_TotalFamiliaPrecoNF - vl_TotalFamiliaPrecoVenda
		vl_saldo_a_pagar = vl_TotalFamiliaPrecoNF - vl_TotalFamiliaPago - vl_TotalFamiliaDevolucaoPrecoNF
		s_vl_saldo_a_pagar = formata_moeda(vl_saldo_a_pagar)
	'	VALORES NEGATIVOS REPRESENTAM O 'CR�DITO' QUE O CLIENTE POSSUI EM CASO DE PEDIDOS CANCELADOS QUE HAVIAM SIDO PAGOS
		if (st_pagto = ST_PAGTO_PAGO) And (vl_saldo_a_pagar > 0) then s_vl_saldo_a_pagar = ""
		
	'	H� DEVOLU��ES?
		if Not le_pedido_item_devolvido(pedido_selecionado, v_item_devolvido, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		for i=Lbound(v_item_devolvido) to Ubound(v_item_devolvido)
			with v_item_devolvido(i)
				if .produto <> "" then
					if .qtde = 1 then s = "" else s = "s"
					if s_devolucoes <> "" then s_devolucoes = s_devolucoes & chr(13) & "<br>" & chr(13)
					s_devolucoes = s_devolucoes & formata_data(.devolucao_data) & " " & _
								   formata_hhnnss_para_hh_nn(.devolucao_hora) & " - " & _
								   formata_inteiro(.qtde) & " unidade" & s & " do " & .produto & " - " & produto_formata_descricao_em_html(.descricao_html)
					if Trim(.motivo) <> "" then	s_devolucoes = s_devolucoes & " (" & .motivo & ")"
					if .NFe_numero_NF > 0 then s_devolucoes = s_devolucoes & " [NF: " & .NFe_numero_NF & "]"
					end if
				end with
			next
		
	'	H� PERDAS?
		if Not le_pedido_perda(pedido_selecionado, v_pedido_perda, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		for i=Lbound(v_pedido_perda) to Ubound(v_pedido_perda)
			with v_pedido_perda(i)
				if .id <> "" then
					vl_total_perdas = vl_total_perdas + .valor
					if s_perdas <> "" then s_perdas = s_perdas & chr(13) & "<br>" & chr(13)
					s_perdas = s_perdas & formata_data(.data) & " " & _
							   formata_hhnnss_para_hh_nn_ss(.hora) & ": " & SIMBOLO_MONETARIO & " " & formata_moeda(.valor)
					if Trim(.obs) <> "" then s_perdas = s_perdas & " (" & .obs & ")"
					end if
				end with
			next
		end if




' _____________________________________________________________________________________________
'
'									F  U  N  �  �  E  S 
' _____________________________________________________________________________________________

' ___________________________________
' EXIBE_FAMILIA_PEDIDO
'
function exibe_familia_pedido(byval pedido_selecionado, byref v_pedido)
const PEDIDOS_POR_LINHA = 8
dim i
dim n
dim x
	exibe_familia_pedido = ""
	if Ubound(v_pedido) = Lbound(v_pedido) then exit function

	x = "<table width='649' class='Q' cellspacing='0'>" & chr(13) & _
		"<tr><td align='left'>" & chr(13) & _
		"<p class='Rf'>FAM�LIA DE PEDIDOS</p>" & chr(13) & _
		"<table width='100%' class='QT' cellspacing='0'>" & chr(13) & _
		"<tr>" & chr(13)
	
	n = 0
	for i = Lbound(v_pedido) to Ubound(v_pedido)
		if Trim(v_pedido(i))<>"" then
			n = n+1
			if n > PEDIDOS_POR_LINHA then 
				n = 1
				x = x & "</tr>" & chr(13) & "<tr>"
				end if
			x = x & "<td width='12.5%' class='L' style='text-align:left;color:black;' align='left'>"
			if v_pedido(i) <> pedido_selecionado then 
				x = x & "<a href='pedido.asp?pedido_selecionado=" & Trim(v_pedido(i)) & "&url_back=X" & _
						"' title='clique para consultar o pedido' class='L' style='color:black;'>"
				end if
			if v_pedido(i) = pedido_selecionado then
				x = x & "<span style='color:gray;'>" & Trim(v_pedido(i)) & "<span>"
			else
				x = x & Trim(v_pedido(i))
				end if
			if v_pedido(i) <> pedido_selecionado then x = x & "</a>"
			x = x & "</td>" & chr(13)
			end if
		next
			
	if (n Mod PEDIDOS_POR_LINHA)<> 0 then
		for i = ((n Mod PEDIDOS_POR_LINHA)+1) to PEDIDOS_POR_LINHA
			x = x & "<td align='left'>&nbsp;</td>" & chr(13)
			next
		end if
	
	x = x & "</tr></table>" & chr(13) & _
			"</td></tr></table>" & chr(13) & _
			"<br>"
	
	exibe_familia_pedido = x
end function

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
	<title><%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	$(document).ready(function() {
		var topo = $('#divConsultaPedido').offset().top - parseFloat($('#divConsultaPedido').css('margin-top').replace(/auto/, 0)) - parseFloat($('#divConsultaPedido').css('padding-top').replace(/auto/, 0));
		$('#divConsultaPedido').addClass('divFixo');
		$("#divClienteConsultaView").hide();
		$('#divInternoClienteConsultaView').addClass('divFixo');
		sizeDivClienteConsultaView();

		$(document).keyup(function(e) {
		    if (e.keyCode == 27) fechaDivClienteConsultaView();
		});

		$("#divClienteConsultaView").click(function() {
		    fechaDivClienteConsultaView();
		});

		$("#imgFechaDivClienteConsultaView").click(function() {
		    fechaDivClienteConsultaView();
		});

		$(".tdGarInd").hide();
		// Para a nova vers�o da forma de pagamento
		if ($(".tdGarInd").prev("td").hasClass("MD")) { $(".tdGarInd").prev("td").removeClass("MD") };
		// Para a vers�o antiga da forma de pagamento
		if ($(".tdGarInd").prev("td").hasClass("MDB")) { $(".tdGarInd").prev("td").removeClass("MDB").addClass("MB") }
	});

//Every resize of window
$(window).resize(function() {
    sizeDivClienteConsultaView();
});

function sizeDivClienteConsultaView() {
    var newHeight = $(document).height() + "px";
    $("#divClienteConsultaView").css("height", newHeight);
}

function fechaDivClienteConsultaView() {
    $("#divClienteConsultaView").fadeOut();
    $("#iframeClienteConsultaView").attr("src", "");
}

function fCLIConsultaView(id_cliente, usuario) {
    sizeDivClienteConsultaView();
    $("#iframeClienteConsultaView").attr("src", "ClienteConsultaView.asp?cliente_selecionado=" + id_cliente + "&usuario=" + usuario + "&ocultar_botoes=S");
    $("#divClienteConsultaView").fadeIn();
}
</script>

<script language="JavaScript" type="text/javascript">
<%=monta_funcao_js_normaliza_numero_pedido_e_sufixo%>

function restauraVisibility(nome_controle) {
	var c;
	c = document.getElementById(nome_controle);
	if (c) c.style.visibility = "";
}

function trataCliqueBotao(id_botao) {
	var c;
	c = document.getElementById(id_botao);
	c.style.visibility = "hidden";
	setTimeout("restauraVisibility('" + id_botao + "')", 20000);
}

function fPEDPESQConclui() {
	var c;
	if (trim(fPEDPESQ.pedido_selecionado.value) == "") return;
	if (normaliza_numero_pedido_e_sufixo(fPEDPESQ.pedido_selecionado.value) != '') {
		fPEDPESQ.pedido_selecionado.value = normaliza_numero_pedido_e_sufixo(fPEDPESQ.pedido_selecionado.value);
	}

	if (isNumeroOrcamento(fPEDPESQ.pedido_selecionado.value)) {
		fPEDPESQ.orcamento_selecionado.value = fPEDPESQ.pedido_selecionado.value;
		fPEDPESQ.action = "orcamento.asp";
	}
	else {
		fPEDPESQ.action = "pedido.asp";
	}

	trataCliqueBotao("imgPedPesq");

	fPEDPESQ.submit();
}

function fPEDOcorrenciaAlteraImpressao(f) {
	if (document.getElementById("tableOcorrencia").className == "notPrint") {
		document.getElementById("tableOcorrencia").className = "";
		document.getElementById("imgPrinterOcorrencia").src = document.getElementById("imgPrinterOcorrencia").src.replace("PrinterError.png", "Printer.png");
	}
	else {
		document.getElementById("tableOcorrencia").className = "notPrint";
		document.getElementById("imgPrinterOcorrencia").src = document.getElementById("imgPrinterOcorrencia").src.replace("Printer.png", "PrinterError.png");
	}
}

function fPEDBlocoNotasAlteraImpressao(f) {
	if (document.getElementById("tableBlocoNotas").className == "notPrint") {
		document.getElementById("tableBlocoNotas").className = "";
		document.getElementById("imgPrinterBlocoNotas").src = document.getElementById("imgPrinterBlocoNotas").src.replace("PrinterError.png", "Printer.png");
	}
	else {
		document.getElementById("tableBlocoNotas").className = "notPrint";
		document.getElementById("imgPrinterBlocoNotas").src = document.getElementById("imgPrinterBlocoNotas").src.replace("Printer.png", "PrinterError.png");
	}
}

function fPEDBlocoNotasItemDevolvidoAlteraImpressao(f) {
	if (document.getElementById("tableBlocoNotasItemDevolvido").className == "notPrint") {
		document.getElementById("tableBlocoNotasItemDevolvido").className = "";
		document.getElementById("imgPrinterBlocoNotasItemDevolvido").src = document.getElementById("imgPrinterBlocoNotasItemDevolvido").src.replace("PrinterError.png", "Printer.png");
	}
	else {
		document.getElementById("tableBlocoNotasItemDevolvido").className = "notPrint";
		document.getElementById("imgPrinterBlocoNotasItemDevolvido").src = document.getElementById("imgPrinterBlocoNotasItemDevolvido").src.replace("Printer.png", "PrinterError.png");
	}
}

function fCLIConsulta() {
	window.status = "Aguarde ...";
	fCLI.edicao_bloqueada.value = 'S';
	fCLI.submit();
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
<link href="<%=URL_FILE__ESCREEN_CSS%>" rel="stylesheet" type="text/css" media="screen">

<style type="text/css">
#rb_etg_imediata {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#divConsultaPedidoWrapper
{
	left:1px;
	position:absolute;
	margin-left:1px;
	width:110px;
	z-index:0;
}
#divConsultaPedido
{
	margin-top:60px;
	border: 1px solid #A9A9A9;
	padding-top: 4px;
	padding-bottom: 4px;
	padding-left: 6px;
	padding-right: 6px;
	position: absolute;
	background-color: #F5F5F5;
	top:0;
	z-index:0;
}
#divConsultaPedido.divFixo
{
	position:fixed;
	top:0;
}
#divClienteConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoClienteConsultaView
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoClienteConsultaView.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivClienteConsultaView
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframeClienteConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
</style>


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ********************************************************** -->
<!-- **********  P�GINA PARA EXIBIR O PEDIDO  ***************** -->
<!-- ********************************************************** -->
<body onload="fPEDPESQ.pedido_selecionado.focus();" link="#ffffff" alink="#ffffff" vlink="#ffffff">

<div id="divConsultaPedidoWrapper" class="notPrint" style="z-index:1000;">
	<div id="divConsultaPedido" class="notPrint">
	<form action="pedido.asp" id="fPEDPESQ" name="fPEDPESQ" method="post" onsubmit="if (trim(fPEDPESQ.pedido_selecionado.value)=='')return false;">
	<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
	<span class="Rf">N� Pedido</span><br />
	<span class="Rf">ou Pr�-Pedido</span><br />
	<input maxlength="10" name="pedido_selecionado" class="C" style="width:75px;margin-left:0px;margin-right:0px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {fPEDPESQConclui();} filtra_pedido();" onblur="if (normaliza_numero_pedido_e_sufixo(this.value)!='') {this.value=normaliza_numero_pedido_e_sufixo(this.value);}">
	<input type="hidden" name="orcamento_selecionado" value="" />
	<br />
	<center>
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="p�gina inicial"><img src="../imagem/home_22x22.png" id="imgPagInicial" alt="p�gina inicial" title="p�gina inicial" style="border:0;margin-top:3px;" onclick="trataCliqueBotao('imgPagInicial');" /></a>
	<input type="image" id="imgPedPesq" src="../imagem/ok_24x24.png" alt="Submit" style="vertical-align:bottom;margin-left:15px;margin-right:0px;" onclick="fPEDPESQConclui();">
	</center>
	</form>
	</div>
</div>

<center>

<form id="fPED" name="fPED" method="post">
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>

<!--  I D E N T I F I C A � � O   D O   P E D I D O -->
<%=MontaHeaderIdentificacaoPedido(pedido_selecionado, r_pedido, 649)%>
<br>

<!--  EXIBE LINKS PARA A FAM�LIA DE PEDIDOS?   -->
<%=exibe_familia_pedido(pedido_selecionado, v_pedido)%>


<!--  L O J A   -->
<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	set r_loja = New cl_LOJA
	if x_loja_bd(r_pedido.loja, r_loja) then
		with r_loja
			if Trim(.razao_social) <> "" then
				s = Trim(.razao_social)
			else
				s = Trim(.nome)
				end if
			end with
		end if
%>
	<td class="MD" align="left"><p class="Rf">LOJA</p><p class="C"><%=s%>&nbsp;</p></td>
	<td width="145" class="MD" align="left"><p class="Rf">INDICADOR</p><p class="C"><%=r_pedido.indicador%>&nbsp;</p></td>
	<td width="145" align="left"><p class="Rf">VENDEDOR</p><p class="C"><%=r_pedido.vendedor%>&nbsp;</p></td>
	</tr>
	</table>

<br>

<!--  CLIENTE   -->
<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	set r_cliente = New cl_CLIENTE
	if x_cliente_bd(r_pedido.id_cliente, r_cliente) then
	
	
%>
<%	if r_cliente.tipo = ID_PF then s_aux="CPF" else s_aux="CNPJ"
	s = cnpj_cpf_formata(r_cliente.cnpj_cpf) 
%>
		<td align="left" width="50%" class="MD"><p class="Rf"><%=s_aux%></p>
		
			<a href='javascript:fCLIConsulta();' title='clique para consultar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
		
		</td>
		<%
		with r_cliente
		if .tipo = ID_PF then s = Trim(.rg) else s = Trim(.ie)
	end with
			if r_cliente.tipo = ID_PF then 
%>
	<td align="left" class="MD"><p class="Rf">RG</p><p class="C"><%=s%>&nbsp;</p></td>
<% else %>
	<td align="left" class="MD"><p class="Rf">IE</p><p class="C"><%=s%>&nbsp;</p></td>
<% end if %>
<td align="center" valign="middle" style="width:15px"><a href='javascript:fCLIConsultaView(<%=chr(34) & r_cliente.id & chr(34) & "," & chr(34) & usuario & chr(34)%>);' title="clique para visualizar o cadastro do cliente"><img id="imgClienteConsultaView" src="../imagem/doc_preview_22.png" /></a></td>
		</tr>
<%
		
		with r_cliente
			if Trim(.nome) <> "" then
				s = Trim(.nome)
				end if
			end with
		end if
	
	if r_cliente.tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZ�O SOCIAL DO CLIENTE"
%>
    <tr>
	<td class="MC" align="left" colspan="3"><p class="Rf"><%=s_aux%></p>
	
		<a href='javascript:fCLIConsulta();' title='clique para consultar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
	
		</td>
	</tr>
	</table>

<!--  ENDERE�O DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%	with r_cliente
		s = formata_endereco(.endereco, .endereco_numero, .endereco_complemento, .bairro, .cidade, .uf, .cep)
		end with
%>		
		<td align="left"><p class="Rf">ENDERE�O</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
</table>

<!--  TELEFONE DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%	s = ""
	with r_cliente
		if Trim(.tel_res) <> "" then
			s = telefone_formata(Trim(.tel_res))
			s_aux=Trim(.ddd_res)
			if s_aux<>"" then s = "(" & s_aux & ") " & s
			end if
		end with
	
	s2 = ""
	with r_cliente
		if Trim(.tel_com) <> "" then
			s2 = telefone_formata(Trim(.tel_com))
			s_aux = Trim(.ddd_com)
			if s_aux<>"" then s2 = "(" & s_aux & ") " & s2
			s_aux = Trim(.ramal_com)
			if s_aux<>"" then s2 = s2 & "  (R. " & s_aux & ")"
			end if
		end with
	with r_cliente
		if Trim(.tel_cel) <> "" then
			s3 = telefone_formata(Trim(.tel_cel))
			s_aux = Trim(.ddd_cel)
			if s_aux<>"" then s3 = "(" & s_aux & ") " & s3
			end if
		end with
	with r_cliente
		if Trim(.tel_com_2) <> "" then
			s4 = telefone_formata(Trim(.tel_com_2))
			s_aux = Trim(.ddd_com_2)
			if s_aux<>"" then s4 = "(" & s_aux & ") " & s4
			s_aux = Trim(.ramal_com_2)
			if s_aux<>"" then s4 = s4 & "  (R. " & s_aux & ")"
			end if
		end with
	
%>

<% if r_cliente.tipo = ID_PF then %>
	<td class="MD" width="33%" align="left"><p class="Rf">TELEFONE RESIDENCIAL</p><p class="C"><%=s%>&nbsp;</p></td>
	<td class="MD" width="33%" align="left"><p class="Rf">TELEFONE COMERCIAL</p><p class="C"><%=s2%>&nbsp;</p></td>
		<td align="left"><p class="Rf">CELULAR</p><p class="C"><%=s3%>&nbsp;</p></td>

<% else %>
	<td class="MD" width="50%" align="left"><p class="Rf">TELEFONE</p><p class="C"><%=s2%>&nbsp;</p></td>
	<td width="50%" align="left"><p class="Rf">TELEFONE</p><p class="C"><%=s4%>&nbsp;</p></td>

<% end if %>

	</tr>
</table>

<!--  E-MAIL DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left"><p class="Rf">E-MAIL</p><p class="C"><%=Trim(r_cliente.email)%>&nbsp;</p></td>
	</tr>
</table>

<!--  ENDERE�O DE ENTREGA  -->
<%	with r_pedido
		s = formata_endereco(.EndEtg_endereco, .EndEtg_endereco_numero, .EndEtg_endereco_complemento, .EndEtg_bairro, .EndEtg_cidade, .EndEtg_uf, .EndEtg_cep)
		end with
%>		
<table width="649" class="QS" cellspacing="0" style="table-layout:fixed">
	<tr>
		<td align="left"><p class="Rf">ENDERE�O DE ENTREGA</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
    <%	if r_pedido.EndEtg_cod_justificativa <> "" then %>		
	<tr>
		<td align="left" style="word-wrap:break-word"><p class="C"><%=obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA,r_pedido.EndEtg_cod_justificativa)%>&nbsp;</p></td>
	</tr>
    <%end if %>
</table>


<!--  R E L A � � O   D E   P R O D U T O S  -->
<br>
<br>
<table class="Qx" cellspacing="0">
	<tr bgcolor="#FFFFFF">
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Produto</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Descri��o</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Qtd</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Falt</span></td>
	<% if blnTemRA Or ((r_pedido.permite_RA_status = 1) And (r_pedido.opcao_possui_RA = "S")) then %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Pre�o</span></td>
	<% end if %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Lista</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Desc</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Venda</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Total</span></td>
	</tr>

<% m_TotalDestePedido=0
   m_TotalDestePedidoComRA=0
   n = Lbound(v_item)-1
   for i=1 to MAX_ITENS 
	 n = n+1
	 s_cor = "black"
	 if n <= Ubound(v_item) then
		with v_item(n)
			s_fabricante=.fabricante
			s_produto=.produto
			s_descricao=.descricao
			s_descricao_html=produto_formata_descricao_em_html(.descricao_html)
			s_qtde=.qtde
			s_preco_lista=formata_moeda(.preco_lista)
			if .desc_dado=0 then s_desc_dado="" else s_desc_dado=formata_perc_desc(.desc_dado)
			s_vl_unitario=formata_moeda(.preco_venda)
			if .preco_NF <> 0 then s_preco_NF=formata_moeda(.preco_NF) else s_preco_NF=""
			m_TotalItem=.qtde * .preco_venda
			m_TotalItemComRA=.qtde * .preco_NF
			s_vl_TotalItem=formata_moeda(m_TotalItem)
			m_TotalDestePedido=m_TotalDestePedido + m_TotalItem
			m_TotalDestePedidoComRA=m_TotalDestePedidoComRA + m_TotalItemComRA
			end with
		s_falta=""
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then
			with v_disp(n)
				if .qtde_estoque_sem_presenca<>0 then s_falta=Cstr(.qtde_estoque_sem_presenca)
				s_cor = x_cor_item(.qtde, .qtde_estoque_vendido, .qtde_estoque_sem_presenca)
				end with
			end if
			
	 else
		s_fabricante=""
		s_produto=""
		s_descricao=""
		s_descricao_html=""
		s_qtde=""
		s_falta=""
		s_preco_lista=""
		s_desc_dado=""
		s_vl_unitario=""
		s_preco_NF=""
		s_vl_TotalItem=""
		end if
%>
	<% if (i > MIN_LINHAS_ITENS_IMPRESSAO_PEDIDO) And (s_produto = "") then %>
	<tr class="notPrint">
	<% else %>
	<tr>
	<% end if %>
	<td class="MDBE" align="left"><input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:25px; color:<%=s_cor%>"
		value='<%=s_fabricante%>' readonly tabindex=-1></td>
	<td class="MDB" align="left"><input name="c_produto" id="c_produto" class="PLLe" style="width:54px; color:<%=s_cor%>"
		value='<%=s_produto%>' readonly tabindex=-1></td>
	<td class="MDB" align="left" style="width:277px;">
		<span class="PLLe" style="color:<%=s_cor%>"><%=s_descricao_html%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value='<%=s_descricao%>'>
	</td>
	<td class="MDB" align="right"><input name="c_qtde" id="c_qtde" class="PLLd" style="width:21px; color:<%=s_cor%>"
		value='<%=s_qtde%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_qtde_falta" id="c_qtde_falta" class="PLLd" style="width:20px; color:<%=s_cor%>"
		value='<%=s_falta%>' readonly tabindex=-1></td>
	<% if blnTemRA Or ((r_pedido.permite_RA_status = 1) And (r_pedido.opcao_possui_RA = "S")) then %>
	<td class="MDB" align="right"><input name="c_vl_NF" id="c_vl_NF" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_preco_NF%>' readonly tabindex=-1></td>
	<% end if %>
	<td class="MDB" align="right"><input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_preco_lista%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_desc" id="c_desc" class="PLLd" style="width:28px; color:<%=s_cor%>"
		value='<%=s_desc_dado%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_vl_unitario%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px; color:<%=s_cor%>" 
		value='<%=s_vl_TotalItem%>' readonly tabindex=-1></td>
	</tr>
<% next %>

	<tr>
	<td colspan="4" align="left">
		<table cellspacing="0" cellpadding="0" width='100%' style="margin-top:4px;">
			<tr>
			<td width="50%" align="left">&nbsp;</td>
			<% if blnTemRA Or ((r_pedido.permite_RA_status = 1) And (r_pedido.opcao_possui_RA = "S")) then %>
			<td align="right">
				<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
					<tr>
					<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;RA Bruto</span></td>
					<td class="MTBD" align="right"><input name="c_total_RA" id="c_total_RA" class="PLLd" style="width:70px;color:<%if m_TotalFamiliaParcelaRA >=0 then Response.Write " green" else Response.Write " red"%>;" 
						value='<%=formata_moeda(m_TotalFamiliaParcelaRA)%>' readonly tabindex=-1></td>
					</tr>
				</table>
			</td>
			<% end if %>
			<td align="right">
				<table cellspacing="0" cellpadding="0">
				<tr>
				<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;COM(%)</span></td>
				<td class="MTBD" align="right"><input name="c_perc_RT" id="c_perc_RT" class="PLLd" style="width:30px;color:blue;" 
					value='<%=formata_perc_RT(r_pedido.perc_RT)%>' readonly tabindex=-1></td>
				</tr>
			</table>
			</td>
			</tr>
		</table>
	</td>
	<% if blnTemRA Or ((r_pedido.permite_RA_status = 1) And (r_pedido.opcao_possui_RA = "S")) then %>
	<td class="MD" align="left">&nbsp;</td>
	<td class="MDB" align="right">
		<input name="c_total_NF" id="c_total_NF" class="PLLd" style="width:70px;color:blue;" 
				value='<%=formata_moeda(m_TotalDestePedidoComRA)%>' readonly tabindex=-1>
	</td>
	<td colspan="3" class="MD" align="left">&nbsp;</td>
	<% else %>
	<td colspan="4" class="MD" align="left">&nbsp;</td>
	<% end if %>
	<td class="MDB" align="right"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:blue;" 
		value='<%=formata_moeda(m_TotalDestePedido)%>' readonly tabindex=-1></td>
	</tr>
</table>

<!--  NOVA VERS�O DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" style="width:649px;" cellspacing="0">
	<tr>
		<td class="MB" colspan="6" align="left"><p class="Rf">Observa��es </p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:99%;margin-left:2pt;" 
				readonly tabindex=-1><%=r_pedido.obs_1%></textarea>
			<span class="PLLe notVisible"><%
				s = substitui_caracteres(r_pedido.obs_1,chr(13),"<br>")
				if s = "" then s = "&nbsp;"
				Response.Write s %></span>
		</td>
	</tr>
    <tr>
		<td class="MB" colspan="7" align="left"><p class="Rf">Constar na NF</p>
			<textarea name="c_nf_texto" id="c_nf_texto" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_NF_TEXTO_CONSTAR)%>" 
				style="width:641px;margin-left:2pt;"
				readonly tabindex=-1><%=r_pedido.NFe_texto_constar%></textarea>
            <span class="PLLe notVisible"><%
				s = substitui_caracteres(r_pedido.NFe_texto_constar,chr(13),"<br>")
				if s = "" then s = "&nbsp;"
				Response.Write s %></span>
		</td>
	</tr>
    <tr>
        <td class="MB" align="left" colspan="7" nowrap><p class="Rf">xPed</p>
			<input name="c_num_pedido_compra" id="c_num_pedido_compra" class="PLLe" maxlength="15" style="width:100px;margin-left:2pt;" onkeypress="filtra_nome_identificador();" onblur="this.value=trim(this.value);"
				value='<%=r_pedido.NFe_xPed%>' readonly tabindex=-1>
		</td>
    </tr>
	<tr>
		<td class="MD" align="left" nowrap><p class="Rf">N� Nota Fiscal</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" style="width:75px;margin-left:2pt;" 
				readonly tabindex=-1 value='<%=r_pedido.obs_2%>'>
		</td>
		<td class="MD" align="left" nowrap><p class="Rf">NF Simples Remessa</p>
			<input name="c_obs3" id="c_obs3" class="PLLe" style="width:75px;margin-left:2pt;" 
				readonly tabindex=-1 value='<%=r_pedido.obs_3%>'>
		</td>
		<td class="MD" nowrap align="left" valign="top"><p class="Rf">Entrega Imediata</p>
		<% 	if Cstr(r_pedido.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_NAO) then
				s = "N�O"
			elseif Cstr(r_pedido.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			 
			if s <> "" then
				s_aux=formata_data_e_talvez_hora_hhmm(r_pedido.etg_imediata_data)
				if s_aux <> "" then s = s & " &nbsp; (" & iniciais_em_maiusculas(r_pedido.etg_imediata_usuario) & " - " & s_aux & ")"
				end if
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MD" nowrap align="left" valign="top"><p class="Rf">Bem Uso/Consumo</p>
		<% 	if Cstr(r_pedido.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then
				s = "N�O"
			elseif Cstr(r_pedido.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MD" nowrap align="left" valign="top"><p class="Rf">Instalador Instala</p>
		<% 	if Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_NAO) then
				s = "N�O"
			elseif Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="tdGarInd" nowrap align="left" valign="top"><p class="Rf">Garantia Indicador</p>
		<% 	if Cstr(r_pedido.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then
				s = "N�O"
			elseif Cstr(r_pedido.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
	</tr>
</table>
<br>
<table class="Q" style="width:649px;" cellspacing="0">
  <tr>
	<td align="left"><span class="Rf">Forma de Pagamento</span></td>
  </tr>
  <tr>
	<td align="left">
	  <table width="100%" cellspacing="0" cellpadding="0" border="0">
		<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then %>
		<!--  � VISTA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">� Vista&nbsp&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.av_forma_pagto)%>)</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
		<!--  PARCELA �NICA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">Parcela �nica:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pu_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pu_forma_pagto)%>)&nbsp;&nbsp;vencendo ap�s&nbsp;<%=formata_inteiro(r_pedido.pu_vencto_apos)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
		<!--  PARCELADO NO CART�O (INTERNET)  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">Parcelado no Cart�o (internet) em&nbsp;&nbsp;<%=Cstr(r_pedido.pc_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_pedido.pc_valor_parcela)%></span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
		<!--  PARCELADO NO CART�O (MAQUINETA)  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">Parcelado no Cart�o (maquineta) em&nbsp;&nbsp;<%=Cstr(r_pedido.pc_maquineta_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_pedido.pc_maquineta_valor_parcela)%></span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
		<!--  PARCELADO COM ENTRADA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">Entrada:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pce_entrada_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pce_forma_pagto_entrada)%>)</span></td>
			  </tr>
			  <tr>
				<td align="left"><span class="C">Presta��es:&nbsp;&nbsp;<%=formata_inteiro(r_pedido.pce_prestacao_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pce_prestacao_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pce_forma_pagto_prestacao)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=formata_inteiro(r_pedido.pce_prestacao_periodo)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
		<!--  PARCELADO SEM ENTRADA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">1� Presta��o:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pse_prim_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pse_forma_pagto_prim_prest)%>)&nbsp;&nbsp;vencendo ap�s&nbsp;<%=formata_inteiro(r_pedido.pse_prim_prest_apos)%>&nbsp;dias</span></td>
			  </tr>
			  <tr>
				<td align="left"><span class="C">Demais Presta��es:&nbsp;&nbsp;<%=Cstr(r_pedido.pse_demais_prest_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pse_demais_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pse_forma_pagto_demais_prest)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=Cstr(r_pedido.pse_demais_prest_periodo)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% end if %>
	  </table>
	</td>
  </tr>
  <tr>
	<td class="MC" align="left"><p class="Rf">Informa��es Sobre An�lise de Cr�dito</p>
	  <textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
				style="width:642px;margin-left:2pt;"
				readonly tabindex=-1><%=r_pedido.forma_pagto%></textarea>
	  <span class="PLLe notVisible"><%
			s = substitui_caracteres(r_pedido.forma_pagto,chr(13),"<br>")
			if s = "" then s = "&nbsp;"
			Response.Write s %></span>
	</td>
  </tr>
</table>


<!--  STATUS DE PAGAMENTO   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td width="16.67%" class="MD" align="left" valign="bottom"><span class="Rf">Status de Pagto</span></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><span class="Rf">VL Total&nbsp;&nbsp;(Fam�lia)&nbsp;</span></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><span class="Rf">VL Pago&nbsp;</span></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><span class="Rf">VL Devolu��es&nbsp;</span></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><span class="Rf">VL Perdas&nbsp;</span></td>
	<td width="16.65%" align="right" valign="bottom"><span class="Rf">Saldo a Pagar&nbsp;</span></td>
</tr>
<tr>
	<% s_aux = x_status_pagto_cor(st_pagto) 
	   s = Ucase(x_status_pagto(st_pagto)) %>
	<td width="16.67%" class="MD" align="left"><span class="C" style="color:<%=s_aux%>;"><%=s%>&nbsp;</span></td>
	<% s = formata_moeda(vl_TotalFamiliaPrecoNF) %>
	<td width="16.67%" align="right" class="MD"><span class="Cd"><%=s%></span></td>
	<% s = formata_moeda(vl_TotalFamiliaPago) %>
	<td width="16.67%" align="right" class="MD"><span class="Cd" style="color:<%
		if vl_TotalFamiliaPago >= 0 then Response.Write "black" else Response.Write "red" 
		%>;"><%=s%></span></td>
	<% s = formata_moeda(vl_TotalFamiliaDevolucaoPrecoNF) %>
	<td width="16.67%" align="right" class="MD"><span class="Cd"><%=s%></span></td>
	<% s = formata_moeda(vl_total_perdas) %>
	<td width="16.67%" align="right" class="MD"><span class="Cd"><%=s%></span></td>
	<td width="16.65%" align="right"><span class="Cd" style="color:<% 
		if vl_saldo_a_pagar >= 0 then Response.Write "black" else Response.Write "red" 
		%>;"><%=s_vl_saldo_a_pagar%></span></td>
</tr>
</table>


<!--  AN�LISE DE CR�DITO   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<%	s=x_analise_credito(r_pedido.analise_credito)
		if s <> "" then
			s_aux=formata_data_e_talvez_hora_hhmm(r_pedido.analise_credito_data)
			if Trim(r_pedido.analise_credito_usuario) <> "" then
				if s_aux <> "" then s_aux = s_aux & " - "
				s_aux = s_aux & iniciais_em_maiusculas(Trim(r_pedido.analise_credito_usuario))
				end if
			if s_aux <> "" then s = s & " &nbsp; (" & s_aux & ")"
			end if
		if s="" then s="&nbsp;"
	%>
	<td align="left"><p class="Rf">AN�LISE DE CR�DITO</p><p class="C" style="color:<%=x_analise_credito_cor(r_pedido.analise_credito)%>;"><%=s%></p></td>
</tr>
</table>


<% if s_devolucoes <> "" then %>
<!--  DEVOLU��ES   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td align="left"><p class="Rf" style="color:red;">DEVOLU��O DE MERCADORIAS</p><p class="C"><%=s_devolucoes%></p></td>
</tr>
</table>
<% end if %>


<% if s_perdas <> "" then %>
<!--  PERDAS   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td align="left"><p class="Rf" style="color:red;">PERDAS</p><p class="C"><%=s_perdas%></p></td>
</tr>
</table>
<% end if %>


<% if IsEntregaAgendavel(r_pedido.st_entrega) then %>
<!--  DATA DE COLETA   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<%	s=formata_data(r_pedido.a_entregar_data_marcada)
		if s="" then s="&nbsp;"
	%>
	<td align="left"><p class="Rf">DATA DE COLETA</p><p class="C"><%=s%></p></td>
</tr>
</table>
<% end if %>


<% if r_pedido.transportadora_id <> "" then %>
<!--  TRANSPORTADORA   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<%	s = r_pedido.transportadora_id & " (" & x_transportadora(r_pedido.transportadora_id) & ")"
		if s="" then s="&nbsp;"
	%>
	<td class="MD" align="left"><p class="Rf">TRANSPORTADORA</p><p class="C"><%=s%></p></td>

	<%	if r_pedido.frete_status <> 0 then
			s = formata_moeda(r_pedido.frete_valor)
		else
			s = ""
			end if
		if s = "" then s = "&nbsp;"
	%>
	<td align="right" style="width:65px;"><p class="Rf">FRETE (<%=SIMBOLO_MONETARIO%>)&nbsp;</p><p class="Cd"><%=s%></p></td>
</tr>
</table>
<% end if %>


<br>
<table id="tableOcorrencia" class="notPrint" width="649" cellspacing="0" cellpadding="0" border="0">
<tr>
	<td colspan="3" class="ME MD MC MB" align="left"><span class="Rf">OCORR�NCIAS</span></td>
</tr>
<% s = "SELECT " & _
			"*" & _
			"," & _
			" (" & _
				"SELECT" & _
					" Count(*)" & _
				" FROM t_PEDIDO_OCORRENCIA_MENSAGEM" & _
				" WHERE" & _
					" (id_ocorrencia=t_PEDIDO_OCORRENCIA.id)" & _
					" AND (fluxo_mensagem='" & COD_FLUXO_MENSAGEM_OCORRENCIAS_EM_PEDIDOS__CENTRAL_PARA_LOJA & "')" & _
			") AS qtde_msg_central" & _
		" FROM t_PEDIDO_OCORRENCIA" & _
		" WHERE" & _
			" (pedido = '" & pedido_selecionado & "')" & _
		" ORDER BY" & _
			" dt_hr_cadastro," & _
			" id"
	set rs = cn.execute(s)
	if rs.Eof then %>
		<tr class="notVisible">
			<td colspan="3" class="ME MD MB" align="left">&nbsp;</td>
		</tr>
<%		end if
	blnHaOcorrenciaEmAberto=False
	n_offset_tabela_ocorrencia = 24
	do while Not rs.Eof
		if CInt(rs("finalizado_status"))=0 then blnHaOcorrenciaEmAberto=True
%>
	<tr>
		<td class="ME" style="width:<%=n_offset_tabela_ocorrencia%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" style="width:<%=649-3-n_offset_tabela_ocorrencia%>px;" align="left">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td class="C MD MB tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">ABERTA POR:</span>&nbsp;<%=Trim("" & rs("usuario_cadastro"))%></td>
			<td class="C MD MB tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">EM:</span>&nbsp;<%=formata_data_hora_sem_seg(rs("dt_hr_cadastro"))%></td>
			<%	if CInt(rs("finalizado_status")) <> 0 then
					s_cor = "green"
					s = "Finalizada"
				else
					s_cor = "red"
					if CInt(rs("qtde_msg_central")) > 0 then
						s = "Em Andamento"
					else
						s = "Aberta"
						end if
					end if
			%>
			<td class="C MB tdWithPadding" align="left" valign="top" style="color:<%=s_cor%>"><span class="Rf" style="margin-left:0px;">SITUA��O:</span>&nbsp;<%=UCase(s)%></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="ME" style="width:<%=n_offset_tabela_ocorrencia%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" style="width:<%=649-3-n_offset_tabela_ocorrencia%>px;" align="left">
			<table width="100%" cellspacing="0" cellpadding="1">
			<tr>
			<%	s = Trim("" & rs("contato"))
				s2 = Trim("" & rs("tel_1"))
				if s2 <> "" then
					s2 = telefone_formata(s2)
					s_aux = Trim("" & rs("ddd_1"))
					if s_aux <> "" then s2 = "(" & s_aux & ") " & s2
					if s <> "" then s = s & " &nbsp; "
					s = s & s2
					end if
				s2 = Trim("" & rs("tel_2"))
				if s2 <> "" then
					s2 = telefone_formata(s2)
					s_aux = Trim("" & rs("ddd_2"))
					if s_aux <> "" then s2 = "(" & s_aux & ") " & s2
					if s <> "" then s = s & " &nbsp; "
					s = s & s2
					end if
			%>
			<td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">CONTATO:</span>&nbsp;<%=s%></td>
			</tr>
			<tr>
			<td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">OCORR�NCIA:</span>&nbsp;<%=substitui_caracteres(Trim("" & rs("texto_ocorrencia")), chr(13), "<br>")%></td>
			</tr>
			</table>
		</td>
	</tr>
	
<% s = "SELECT " & _
			"*" & _
	   " FROM t_PEDIDO_OCORRENCIA_MENSAGEM" & _
	   " WHERE" & _
			" (id_ocorrencia = " & Trim("" & rs("id")) & ")" & _
	   " ORDER BY" & _
			" dt_hr_cadastro," & _
			" id"
	set rs2 = cn.execute(s)
%>
	<tr>
		<%	if CInt(rs("finalizado_status"))=0 then s="ME MB" else s="ME" %>
		<td class="<%=s%>" style="width:<%=n_offset_tabela_ocorrencia%>px;" align="left">&nbsp;</td>
		<td class="<%=s%>" style="width:<%=n_offset_tabela_ocorrencia%>px;" align="left">&nbsp;</td>
		<td class="ME MD" style="width:<%=649-3-2*n_offset_tabela_ocorrencia%>px;" align="left">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td class="Rf tdWithPadding" align="left">MENSAGENS</td>
			</tr>
			<% if rs2.Eof then %>
			<tr>
				<td align="left">&nbsp;</td>
			</tr>
			<% end if %>

			<%	do while not rs2.Eof %>
			<tr>
				<td align="left">
					<table width="100%" cellspacing="0" cellpadding="0">
					<tr>
					<td class="C MD MC tdWithPadding" style="width:60px;" align="center" valign="top"><%=formata_data_hora_sem_seg(rs2("dt_hr_cadastro"))%></td>
					<td class="C MD MC tdWithPadding" style="width:80px;" align="center" valign="top"><%
						s = rs2("usuario_cadastro")
						if Trim("" & rs2("loja")) <> "" then s = s & " (Loja&nbsp;" & Trim("" & rs2("loja")) & ")"
						Response.Write s
						%></td>
					<td class="C MC tdWithPadding" align="left" valign="top"><%=substitui_caracteres(Trim("" & rs2("texto_mensagem")), chr(13), "<br>")%></td>
					</tr>
					</table>
				</td>
			</tr>
			<%		rs2.MoveNext
					loop 
			%>

			<% if CInt(rs("finalizado_status"))=0 then %>
			<tr class="notPrint">
				<td class="MC" style="padding:0px;" align="left">
					<table width="100%" cellpadding="0" cellspacing="0">
					<tr>
					<td align="left">&nbsp;</td>
					</tr>
					</table>
				</td>
			</tr>

			<tr class="notPrint">
				<td class="MB" align="left"><span style='font-family: Arial, Helvetica, sans-serif;color:white;font-size:6pt;font-style:normal;'>&nbsp;</span></td>
			</tr>
			<tr class="notVisible">
				<td class="MB MC" align="left"><span style='font-family: Arial, Helvetica, sans-serif;color:white;font-size:6pt;font-style:normal;'>&nbsp;</span></td>
			</tr>
			<% end if %>
			
			</table>
		</td>
	</tr>

	<% if CInt(rs("finalizado_status")) <> 0 then %>
	<tr>
		<td class="ME MB" style="width:<%=n_offset_tabela_ocorrencia%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="MC ME MD" align="left">
			<table width="100%" cellspacing="0" cellpadding="1">
			<tr>
			<td class="C MB" width="100%" align="left" valign="top">
				<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
				<td class="C MD tdWithPadding" width="50%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">FINALIZADA POR:</span>&nbsp;<%=Trim("" & rs("finalizado_usuario"))%></td>
				<td class="C tdWithPadding" width="50%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">EM:</span>&nbsp;<%=formata_data_hora_sem_seg(rs("finalizado_data_hora"))%></td>
				</tr>
				</table>
			</td>
			</tr>
			<tr>
			<% s = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__TIPO_OCORRENCIA, Trim("" & rs("tipo_ocorrencia"))) %>
			<td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">TIPO DE OCORR�NCIA:</span>&nbsp;<%=s%></td>
			</tr>
			<tr>
			<td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">SOLU��O:</span>&nbsp;<%=substitui_caracteres(Trim("" & rs("texto_finalizacao")), chr(13), "<br>")%></td>
			</tr>
			</table>
		</td>
	</tr>
	<% end if %>
<%
		rs.MoveNext
		loop
%>
	<tr class="notPrint">
		<td colspan="3" style="padding:0px;" align="left">
			<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
			<td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="bOcorrenciaAlteraImpressao" id="bOcorrenciaAlteraImpressao" href="javascript:fPEDOcorrenciaAlteraImpressao(fPED)" title="configura as informa��es sobre ocorr�ncias para serem impressas ou n�o"><img name="imgPrinterOcorrencia" id="imgPrinterOcorrencia" src="../botao/PrinterError.png" border="0"></a></td>
			<td align="left">&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>

</table>


<br>
<table id="tableBlocoNotas" class="notPrint" width="649" cellspacing="0" cellpadding="1">
<tr>
	<td colspan="4" class="ME MD MC MB" align="left"><span class="Rf">BLOCO DE NOTAS</span></td>
</tr>
<% s = "SELECT " & _
			"*" & _
		" FROM t_PEDIDO_BLOCO_NOTAS" & _
		" WHERE" & _
			" (pedido = '" & pedido_selecionado & "')" & _
			" AND (nivel_acesso <= " & Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO) & ")" & _
			" AND (anulado_status = 0)" &_
		" ORDER BY" & _
			" dt_hr_cadastro"
	set rs = cn.execute(s)
	if rs.Eof then %>
		<tr class="notVisible">
			<td colspan="4" class="ME MD MB" align="left">&nbsp;</td>
		</tr>
<%		end if
		
	do while Not rs.Eof
%>
	<tr>
		<td class="C ME MD MB" style="width:60px;" align="center" valign="top"><%=formata_data_hora(rs("dt_hr_cadastro"))%></td>
		<td class="C MD MB" style="width:80px;" align="center" valign="top"><%
			s = rs("usuario")
			if Trim("" & rs("loja")) <> "" then s = s & " (Loja&nbsp;" & Trim("" & rs("loja")) & ")"
			Response.Write s
			%></td>
		<td colspan="2" class="C MD MB" align="left" valign="top"><%=substitui_caracteres(Trim("" & rs("mensagem")), chr(13), "<br>")%></td>
	</tr>
<%
		rs.MoveNext
		loop
%>

	<tr class="notPrint">
		<td colspan="4" style="padding:0px;" align="left">
			<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
			<td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="bBlocoNotasAlteraImpressao" id="bBlocoNotasAlteraImpressao" href="javascript:fPEDBlocoNotasAlteraImpressao(fPED)" title="configura as mensagens do bloco de notas para serem impressas ou n�o"><img name="imgPrinterBlocoNotas" id="imgPrinterBlocoNotas" src="../botao/PrinterError.png" border="0"></a></td>
			<td align="left">&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>
</table>


<% if s_devolucoes <> "" then %>
<br>
<table id="tableBlocoNotasItemDevolvido" class="notPrint" width="649" cellspacing="0" cellpadding="1" border="0">
<tr>
	<td colspan="3" class="ME MD MC MB" align="left"><span class="Rf">BLOCO DE NOTAS (DEVOLU��O DE MERCADORIAS)</span></td>
</tr>
<%
'	Obs: devido a algum bug do IE (verificado nas vers�es 8 e 9), quando h� apenas 1 linha de dados, o t�tulo maior
'	desta se��o faz c/ que as colunas n�o apare�am na largura esperada. Por este motivo, foi necess�rio definir
'	explicitamente a largura da coluna "mensagem".
	s = "SELECT " & _
			"tPIDBN.*" & _
		" FROM t_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS tPIDBN" & _
			" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO tPID ON (tPIDBN.id_item_devolvido=tPID.id)" & _
		" WHERE" & _
			" (tPID.pedido = '" & pedido_selecionado & "')" & _
			" AND (tPIDBN.anulado_status = 0)" & _
		" ORDER BY" & _
			" tPIDBN.dt_hr_cadastro," & _
			" tPIDBN.id"
	set rs = cn.execute(s)
	if rs.Eof then %>
		<tr class="notVisible">
			<td colspan="3" class="ME MD MB" align="left">&nbsp;</td>
		</tr>
<%		end if
	
	do while Not rs.Eof
%>
	<tr>
		<td class="C ME MD MB" style="width:60px;" align="center" valign="top"><%=formata_data_hora(rs("dt_hr_cadastro"))%></td>
		<td class="C MD MB" style="width:80px;" align="center" valign="top"><%
			s = rs("usuario")
			if Trim("" & rs("loja")) <> "" then s = s & " (Loja&nbsp;" & Trim("" & rs("loja")) & ")"
			Response.Write s
			%></td>
		<td class="C MD MB" style="width:499px;" align="left" valign="top"><%=substitui_caracteres(Trim("" & rs("mensagem")), chr(13), "<br>")%></td>
	</tr>
<%
		rs.MoveNext
		loop
%>

	<tr class="notPrint">
		<td colspan="3" style="padding:0px;" align="left">
			<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
			<td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="bBlocoNotasItemDevolvidoAlteraImpressao" id="bBlocoNotasItemDevolvidoAlteraImpressao" href="javascript:fPEDBlocoNotasItemDevolvidoAlteraImpressao(fPED)" title="configura as mensagens do bloco de notas de itens devolvidos para serem impressas ou n�o"><img name="imgPrinterBlocoNotasItemDevolvido" id="imgPrinterBlocoNotasItemDevolvido" src="../botao/PrinterError.png" border="0"></a></td>
			<td align="left">&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>

</table>
<% end if %>


<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOT�ES   ************ -->
<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr>
	<% if url_back <> "" then s="Resumo.asp" else s="javascript:history.back()" %>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>" title="volta para p�gina anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>

</form>


<!-- ************   DIRECIONA PARA CADASTRO DE CLIENTES   ************ -->
<form method="post" action="ClienteEdita.asp" id="fCLI" name="fCLI">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=r_pedido.id_cliente%>'>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<input type="hidden" name="edicao_bloqueada" id="edicao_bloqueada" />
<input type="hidden" name="pagina_retorno" id="pagina_retorno" value='Pedido.asp?pedido_selecionado=<%=pedido_selecionado%>&url_back=X'>
</form>


</center>
<div id="divClienteConsultaView"><center><div id="divInternoClienteConsultaView"><img id="imgFechaDivClienteConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframeClienteConsultaView"></iframe></div></center></div>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>