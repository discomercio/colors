<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  P E D I D O C O N S U L T A . A S P
'     ===========================================
'	  P�GINA EXCLUSIVAMENTE P/ VISUALIZAR OS DADOS DO PEDIDO
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

	dim s, usuario, loja, pedido_selecionado,s_sql
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim url_back
	url_back = Trim(request("url_back"))

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then 
		if url_back <> "" then
			Response.Redirect("resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		else
			Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
			end if
		end if
		
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	if Len(pedido_selecionado) > TAM_MAX_ID_PEDIDO then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_INVALIDO)
	
	dim i, n, x, s_fabricante, s_produto, s_descricao, s_descricao_html, s_qtde, s_preco_lista, s_desc_dado
	dim s_vl_unitario, s_vl_TotalItem, m_TotalItem, m_TotalDestePedido, m_TotalItemComRA, m_TotalDestePedidoComRA
	dim s_preco_NF, m_TotalFamiliaParcelaRA, intQtdeFrete, notPrint
    dim blnIsUsuarioResponsavelDepto, blnIsUsuarioCadastroChamado
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim nivel_acesso_bloco_notas
	nivel_acesso_bloco_notas = Session("nivel_acesso_bloco_notas")
	if Trim(nivel_acesso_bloco_notas) = "" then 
		nivel_acesso_bloco_notas = obtem_nivel_acesso_bloco_notas_pedido(cn, usuario)
		Session("nivel_acesso_bloco_notas") = nivel_acesso_bloco_notas
		end if

    dim nivel_acesso_chamado
	nivel_acesso_chamado = Session("nivel_acesso_chamado")
	if Trim(nivel_acesso_chamado) = "" then
		nivel_acesso_chamado = obtem_nivel_acesso_chamado_pedido(cn, usuario)
		Session("nivel_acesso_chamado") = nivel_acesso_chamado
		end if

	dim blnAcessoLojaOk
	blnAcessoLojaOk = False
	
	dim r_pedido, v_item, alerta
	alerta=""
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
	else
		if Trim(r_pedido.loja) = loja then blnAcessoLojaOk = True
		if Not blnAcessoLojaOk then
			if PossuiAcessoLoja(usuario, r_pedido.loja) then blnAcessoLojaOk = True
			end if
		if Not blnAcessoLojaOk then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_ACESSO_NEGADO)
		if Not le_pedido_item(pedido_selecionado, v_item, msg_erro) then alerta = msg_erro
		if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
			if Trim(r_pedido.vendedor <> usuario) then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_ACESSO_NEGADO)
			end if
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
	
	dim n_offset_tabela_bloco_notas_item_devolvido
	dim n_offset_tabela_ocorrencia, n_offset_tabela_chamado, blnHaOcorrenciaEmAberto
    dim n_offset_tabela_devolucao
	dim s_aux, s2, s3, s4, r_loja, r_cliente, s_cor, s_falta, v_pedido
	dim v_disp
	dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF
	dim vl_saldo_a_pagar, s_vl_saldo_a_pagar, st_pagto
	dim v_item_devolvido, s_devolucoes
	dim pedido_splitado
	dim v_pedido_perda, s_perdas, vl_total_perdas, vl_total_frete, frete_transportadora_id, frete_numero_NF, frete_serie_NF
    dim vlTotalItemDevolucao, vlTotalDevolucao
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
			
			if Not estoque_verifica_status_item(v_disp, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
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

	dim s_link_rastreio

	dim strHistPagtoModulo
	dim strHistPagtoStatusDescricao
	dim strHistPagtoStatusImg
	dim strHistPagtoCor
	dim strHistPagtoValorPago
	dim strHistPagtoDescricao
	dim strHistPagtoDtVencto
	dim strHistPagtoVlParcela
	dim strHistPagtoDtPagto
	dim strHistPagtoCorParcelaEmAtraso
	dim dtReferenciaLimitePagamentoEmAtraso

	dim blnPossuiFormaPagtoProporcional, sDescricaoFormaPagtoProporcional, blnFormaPagtoProporcionalNaoSeAplica, blnFormaPagtoProporcionalFalhaCalculo, msgFormaPagtoProporcionalFalhaCalculo
	if alerta = "" then
		blnPossuiFormaPagtoProporcional = monta_descricao_forma_pagto_proporcional(r_pedido, sDescricaoFormaPagtoProporcional, blnFormaPagtoProporcionalNaoSeAplica, blnFormaPagtoProporcionalFalhaCalculo, msgFormaPagtoProporcionalFalhaCalculo, msg_erro)
		if blnPossuiFormaPagtoProporcional then
			if blnFormaPagtoProporcionalNaoSeAplica then blnPossuiFormaPagtoProporcional = False
			end if
		if blnPossuiFormaPagtoProporcional then
			'Se houve falha no c�lculo e foi retornada uma mensagem da falha, pode-se exibi-la ou n�o para o usu�rio
			if blnFormaPagtoProporcionalFalhaCalculo then sDescricaoFormaPagtoProporcional = msgFormaPagtoProporcionalFalhaCalculo
			end if
		if blnPossuiFormaPagtoProporcional then
			'Situa��o inesperada: n�o h� descri��o para ser exibida
			if Trim(sDescricaoFormaPagtoProporcional) = "" then blnPossuiFormaPagtoProporcional = False
			end if
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
				x = x & "<a href='PedidoConsulta.asp?pedido_selecionado=" & Trim(v_pedido(i)) & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & _
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
	<title>LOJA<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	$(document).ready(function() {
		var topo = $('#divConsultaPedido').offset().top - parseFloat($('#divConsultaPedido').css('margin-top').replace(/auto/, 0)) - parseFloat($('#divConsultaPedido').css('padding-top').replace(/auto/, 0));
		$('#divConsultaPedido').addClass('divFixo');
		$("#divOrcamentistaEIndicadorConsultaView").hide();
		$("#divClienteConsultaView").hide();
		$("#divRastreioConsultaView").hide();
		$('#divInternoClienteConsultaView').addClass('divFixo');
		$('#divInternoOrcamentistaEIndicadorConsultaView').addClass('divFixo');
		$('#divInternoRastreioConsultaView').addClass('divFixo');
		sizeDivClienteConsultaView();
		sizeDivOrcamentistaEIndicadorConsultaView();
		sizeDivRastreioConsultaView();

		$(document).keyup(function(e) {
			if (e.keyCode == 27) {
				fechaDivClienteConsultaView();
				fechaDivRastreioConsultaView();
				fechaDivOrcamentistaEIndicadorConsultaView();
            }
		});

		$("#divClienteConsultaView").click(function() {
		    fechaDivClienteConsultaView();
		});

		$("#divOrcamentistaEIndicadorConsultaView").click(function() {
		    fechaDivOrcamentistaEIndicadorConsultaView();
		});

		$("#divRastreioConsultaView").click(function () {
			fechaDivRastreioConsultaView();
		});

		$("#imgFechaDivClienteConsultaView").click(function() {
		    fechaDivClienteConsultaView();
		});

		$("#imgFechaDivOrcamentistaEIndicadorConsultaView").click(function() {
		    fechaDivOrcamentistaEIndicadorConsultaView();
		});

		$("#imgFechaDivRastreioConsultaView").click(function () {
			fechaDivRastreioConsultaView();
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
	sizeDivRastreioConsultaView();
	sizeDivOrcamentistaEIndicadorConsultaView();
});

function sizeDivClienteConsultaView() {
    var newHeight = $(document).height() + "px";
    $("#divClienteConsultaView").css("height", newHeight);
}

function sizeDivOrcamentistaEIndicadorConsultaView() {
    var newHeight = $(document).height() + "px";
    $("#divOrcamentistaEIndicadorConsultaView").css("height", newHeight);
}

function sizeDivRastreioConsultaView() {
	var newHeight = $(document).height() + "px";
	$("#divRastreioConsultaView").css("height", newHeight);
}

function fechaDivClienteConsultaView() {
    $("#divClienteConsultaView").fadeOut();
    $("#iframeClienteConsultaView").attr("src", "");
}
function fechaDivOrcamentistaEIndicadorConsultaView() {
    $("#divOrcamentistaEIndicadorConsultaView").fadeOut();
    $("#iframeOrcamentistaEIndicadorConsultaView").attr("src", "");
}

function fechaDivRastreioConsultaView() {
	$("#divRastreioConsultaView").fadeOut();
	$("#iframeRastreioConsultaView").attr("src", "");
}

function fCLIConsultaView(id_cliente, usuario) {
    sizeDivClienteConsultaView();
    $("#iframeClienteConsultaView").attr("src", "ClienteConsultaView.asp?cliente_selecionado=" + id_cliente + "&usuario=" + usuario + "&ocultar_botoes=S");
    $("#divClienteConsultaView").fadeIn();
}
function fOrcamentistaEIndicadorConsultaView(apelido) {
    sizeDivOrcamentistaEIndicadorConsultaView();
    $("#iframeOrcamentistaEIndicadorConsultaView").attr("src", "OrcamentistaEIndicadorConsultaView.asp?id_selecionado=" + encodeURIComponent(apelido));
    $("#divOrcamentistaEIndicadorConsultaView").fadeIn();
}

function fRastreioConsultaView(url) {
	sizeDivRastreioConsultaView();
	$("#iframeRastreioConsultaView").attr("src", url);
	$("#divRastreioConsultaView").fadeIn();
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

function fCLIEdita( ){
	window.status = "Aguarde ...";
	fCLI.edicao_bloqueada.value = 'N';
	fCLI.submit(); 
}

function fCLIConsulta() {
	window.status = "Aguarde ...";
	fCLI.edicao_bloqueada.value = 'S';
	fCLI.submit();
}

function fPEDPESQConclui() {
	var c;

	if (trim(fPEDPESQ.pedido_selecionado.value) == "") return;

	if (trim(fPEDPESQ.pedido_selecionado.value).toUpperCase().substring(0, 2) == "NF") {
		c = trim(fPEDPESQ.pedido_selecionado.value);
		c = c.substring(2, c.length);
		fPEDPESQ.c_tipo_num_pedido.value = "NF";
		fPEDPESQ.c_nf.value = c;
		fPEDPESQ.action = "RelPesquisaPedidoNF.asp";
	}
	else
	{
		if ((normaliza_numero_pedido_e_sufixo(fPEDPESQ.pedido_selecionado.value) != '')&&(retorna_so_digitos(fPEDPESQ.pedido_selecionado.value).length < 10)&&(retorna_so_digitos(fPEDPESQ.pedido_selecionado.value).length!=8)) {
			fPEDPESQ.pedido_selecionado.value = normaliza_numero_pedido_e_sufixo(fPEDPESQ.pedido_selecionado.value);
		}

		if (isNumeroOrcamento(fPEDPESQ.pedido_selecionado.value)) {
			fPEDPESQ.orcamento_selecionado.value = fPEDPESQ.pedido_selecionado.value;
			fPEDPESQ.action = "orcamento.asp";
		}
		else {
			if ((retorna_so_digitos(fPEDPESQ.pedido_selecionado.value).length == 8)||(retorna_so_digitos(fPEDPESQ.pedido_selecionado.value).length >= 10)||(trim((fPEDPESQ.pedido_selecionado.value).toUpperCase()).substring(0, 1) == "M")) {
				if (trim((fPEDPESQ.pedido_selecionado.value).toUpperCase()).substring(0, 1) == "M") {
					c = trim(fPEDPESQ.pedido_selecionado.value);
					c = c.substring(1, c.length);
					fPEDPESQ.c_num_pedido_aux.value = c;
					fPEDPESQ.c_tipo_num_pedido.value = "<%=OP_PESQ_PEDIDO_MARKETPLACE_AR_CLUBE%>";
					fPEDPESQ.action = "RelPesquisaPedidoEcommerce.asp";
				}
				else if (trim((fPEDPESQ.pedido_selecionado.value).toUpperCase()).substring(0, 1) == "B") {
					c = trim(fPEDPESQ.pedido_selecionado.value);
					c = c.substring(1, c.length);
					fPEDPESQ.c_num_pedido_aux.value = c;
					fPEDPESQ.c_tipo_num_pedido.value = "<%=OP_PESQ_PEDIDO_MAGENTO_BONSHOP%>";
					fPEDPESQ.action = "RelPesquisaPedidoEcommerce.asp";
				}
				else if (trim((fPEDPESQ.pedido_selecionado.value).toUpperCase()).substring(0, 1) == "A") {
					c = trim(fPEDPESQ.pedido_selecionado.value);
					c = c.substring(1, c.length);
					fPEDPESQ.c_num_pedido_aux.value = c;
					fPEDPESQ.c_tipo_num_pedido.value = "<%=OP_PESQ_PEDIDO_MAGENTO_AR_CLUBE%>";
					fPEDPESQ.action = "RelPesquisaPedidoEcommerce.asp";
				}
				else {
					c = trim(fPEDPESQ.pedido_selecionado.value);
					c = c.substring(0, c.length);
					fPEDPESQ.c_num_pedido_aux.value = c;
					fPEDPESQ.c_tipo_num_pedido.value = "<%=OP_PESQ_PEDIDO_MARKETPLACE_AR_CLUBE%>";
					fPEDPESQ.action = "RelPesquisaPedidoEcommerce.asp";
				}
	        
			}
			else if (trim((fPEDPESQ.pedido_selecionado.value).toUpperCase()).substring(0, 1) == "A") {
				c = trim(fPEDPESQ.pedido_selecionado.value);
				c = c.substring(1, c.length);
				fPEDPESQ.c_num_pedido_aux.value = c;
				fPEDPESQ.c_tipo_num_pedido.value = "<%=OP_PESQ_PEDIDO_MAGENTO_AR_CLUBE%>";
				fPEDPESQ.action = "RelPesquisaPedidoEcommerce.asp";
			}
			else if (trim((fPEDPESQ.pedido_selecionado.value).toUpperCase()).substring(0, 1) == "B") {
				c = trim(fPEDPESQ.pedido_selecionado.value);
				c = c.substring(1, c.length);
				fPEDPESQ.c_num_pedido_aux.value = c;
				fPEDPESQ.c_tipo_num_pedido.value = "<%=OP_PESQ_PEDIDO_MAGENTO_BONSHOP%>";
				fPEDPESQ.action = "RelPesquisaPedidoEcommerce.asp";
			}
			else {
				fPEDPESQ.action = "pedido.asp";
			}
		}
	}

	trataCliqueBotao("imgPedPesq");

	fPEDPESQ.submit();
}

function fPEDOcorrenciaAlteraImpressao(f) {
    if (document.getElementById("tableOcorrencia").className == "notPrint") {
        document.getElementById("brOcorrencia").className = "";
        document.getElementById("tableOcorrencia").className = "";
        document.getElementById("imgPrinterOcorrencia").src = document.getElementById("imgPrinterOcorrencia").src.replace("PrinterError.png", "Printer.png");
    }
    else {
        document.getElementById("brOcorrencia").className = "notPrint";
        document.getElementById("tableOcorrencia").className = "notPrint";
        document.getElementById("imgPrinterOcorrencia").src = document.getElementById("imgPrinterOcorrencia").src.replace("Printer.png", "PrinterError.png");
    }
}

function fPEDDevolucaoAlteraImpressao(f) {
    if (document.getElementById("tableDevolucao").className == "notPrint") {
        document.getElementById("brDevolucao").className = "";
        document.getElementById("tableDevolucao").className = "";
        document.getElementById("imgPrinterDevolucao").src = document.getElementById("imgPrinterDevolucao").src.replace("PrinterError.png", "Printer.png");
    }
    else {
        document.getElementById("brDevolucao").className = "notPrint";
        document.getElementById("tableDevolucao").className = "notPrint";
        document.getElementById("imgPrinterDevolucao").src = document.getElementById("imgPrinterDevolucao").src.replace("Printer.png", "PrinterError.png");
    }
}

function fPEDChamadoAlteraImpressao(f) {
    if (document.getElementById("tableChamado").className == "notPrint") {
        document.getElementById("brChamado").className = "";
        document.getElementById("tableChamado").className = "";
        document.getElementById("imgPrinterChamado").src = document.getElementById("imgPrinterChamado").src.replace("PrinterError.png", "Printer.png");
    }
    else {
        document.getElementById("brChamado").className = "notPrint";
        document.getElementById("tableChamado").className = "notPrint";
        document.getElementById("imgPrinterChamado").src = document.getElementById("imgPrinterChamado").src.replace("Printer.png", "PrinterError.png");
    }
}

function fPEDBlocoNotasAlteraImpressao(f) {
    if (document.getElementById("tableBlocoNotas").className == "notPrint") {
        document.getElementById("brBlocoNotas").className = "";
        document.getElementById("tableBlocoNotas").className = "";
        document.getElementById("imgPrinterBlocoNotas").src = document.getElementById("imgPrinterBlocoNotas").src.replace("PrinterError.png", "Printer.png");
    }
    else {
        document.getElementById("brBlocoNotas").className = "notPrint";
        document.getElementById("tableBlocoNotas").className = "notPrint";
        document.getElementById("imgPrinterBlocoNotas").src = document.getElementById("imgPrinterBlocoNotas").src.replace("Printer.png", "PrinterError.png");
    }
}
function fPEDBlocoNotasAT(f) {
    if (document.getElementById("tableBlocoNotasAT").className == "notPrint") {
        document.getElementById("brBlocoNotasAT").className = "";
        document.getElementById("tableBlocoNotasAT").className = "";
        document.getElementById("imgPrinterBlocoNotasAT").src = document.getElementById("imgPrinterBlocoNotasAT").src.replace("PrinterError.png", "Printer.png");
    }
    else {
        document.getElementById("brBlocoNotasAT").className = "notPrint";
        document.getElementById("tableBlocoNotasAT").className = "notPrint";
        document.getElementById("imgPrinterBlocoNotasAT").src = document.getElementById("imgPrinterBlocoNotasAT").src.replace("Printer.png", "PrinterError.png");
    }
}

function fPEDHistPagtoAlteraImpressao(f) {
    if (document.getElementById("tableHistPagto").className == "notPrint") {
        document.getElementById("brHistPagto").className = "";
        document.getElementById("tableHistPagto").className = "";
        document.getElementById("imgPrinterHistPagto").src = document.getElementById("imgPrinterHistPagto").src.replace("PrinterError.png", "Printer.png");
    }
    else {
        document.getElementById("brHistPagto").className = "notPrint";
        document.getElementById("tableHistPagto").className = "notPrint";
        document.getElementById("imgPrinterHistPagto").src = document.getElementById("imgPrinterHistPagto").src.replace("Printer.png", "PrinterError.png");
    }
}

function fPEDDetalhesPagtoCartaoAlteraImpressao(f) {
    if (document.getElementById("tableDetalhesPagtoCartao").className == "notPrint") {
        document.getElementById("brDetalhesPagtoCartao").className = "";
        document.getElementById("tableDetalhesPagtoCartao").className = "";
        document.getElementById("imgPrinterDetalhesPagtoCartao").src = document.getElementById("imgPrinterDetalhesPagtoCartao").src.replace("PrinterError.png", "Printer.png");
    }
    else {
        document.getElementById("brDetalhesPagtoCartao").className = "notPrint";
        document.getElementById("tableDetalhesPagtoCartao").className = "notPrint";
        document.getElementById("imgPrinterDetalhesPagtoCartao").src = document.getElementById("imgPrinterDetalhesPagtoCartao").src.replace("Printer.png", "PrinterError.png");
    }
}

function fPEDBlocoNotasItemDevolvidoAlteraImpressao(f) {
    if (document.getElementById("tableBlocoNotasItemDevolvido").className == "notPrint") {
        document.getElementById("brBlocoNotasItemDevolvido").className = "";
        document.getElementById("tableBlocoNotasItemDevolvido").className = "";
        document.getElementById("imgPrinterBlocoNotasItemDevolvido").src = document.getElementById("imgPrinterBlocoNotasItemDevolvido").src.replace("PrinterError.png", "Printer.png");
    }
    else {
        document.getElementById("brBlocoNotasItemDevolvido").className = "notPrint";
        document.getElementById("tableBlocoNotasItemDevolvido").className = "notPrint";
        document.getElementById("imgPrinterBlocoNotasItemDevolvido").src = document.getElementById("imgPrinterBlocoNotasItemDevolvido").src.replace("Printer.png", "PrinterError.png");
    }
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
<link href="<%=URL_FILE__ESCREEN_CSS%>" Rel="stylesheet" Type="text/css" media="screen">

<style type="text/css">
#rb_etg_imediata {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
.tdWithPadding
{
	padding:1px;
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
#divOrcamentistaEIndicadorConsultaView
{
    position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divRastreioConsultaView
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
#divInternoOrcamentistaEIndicadorConsultaView
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
#divInternoRastreioConsultaView
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#fff;
	opacity: 1;
}
#divInternoClienteConsultaView.divFixo
{
	position:fixed;
	top:6%;
}
#divInternoOrcamentistaEIndicadorConsultaView.divFixo
{
    position:fixed;
	top:6%;
}
#divInternoRastreioConsultaView.divFixo
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
#imgFechaDivOrcamentistaEIndicadorConsultaView
{
    position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#imgFechaDivRastreioConsultaView
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
#iframeOrcamentistaEIndicadorConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
#iframeRastreioConsultaView
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
	<input maxlength="21" name="pedido_selecionado" class="C" style="width:75px;margin-left:0px;margin-right:0px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {fPEDPESQConclui();} filtra_pedido();" onblur="if (this.value.length < 10) { if (normaliza_numero_pedido_e_sufixo(this.value)!='') {this.value=normaliza_numero_pedido_e_sufixo(this.value);}}" autocomplete="off" />
	<input type="hidden" name="orcamento_selecionado" value="" />
    <input type="hidden" name="c_num_pedido_aux" id="c_num_pedido_aux" value="" />
    <input type="hidden" name="c_tipo_num_pedido" id="c_tipo_num_pedido" value="" />
	<input type="hidden" name="c_nf" id="c_nf" value="" />
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
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="id_devolucao" id="id_devolucao" />

<!--  I D E N T I F I C A � � O   D O   P E D I D O -->  
<%=MontaHeaderIdentificacaoPedido(pedido_selecionado, r_pedido, 649)%>
<br>

<!-- EXIBE ALERTA SOBRE NF CANCELADA -->
<%=exibe_alerta_nf_cancelada(pedido_selecionado, r_pedido.obs_1)%>

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
	<td width="90" class="MD" align="left"><p class="Rf">CD</p><p class="C"><%=obtem_apelido_empresa_NFe_emitente(r_pedido.id_nfe_emitente)%>&nbsp;</p></td>
	<td class="MD" align="left"><p class="Rf">LOJA</p><p class="C"><%=s%>&nbsp;</p></td>
	<td width="145" class="MD" align="left"><p class="Rf">INDICADOR</p><a href='javascript:fOrcamentistaEIndicadorConsultaView(<%=chr(34) & r_pedido.indicador & chr(34)%>)' title="clique para consultar o cadastro do indicador"><p class="C"><%=r_pedido.indicador%>&nbsp;</p></a></td>
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
	
    'le as vari�veis da origem certa: ou do pedido ou do cliente, todas comecam com cliente__
    dim cliente__tipo, cliente__cnpj_cpf, cliente__rg, cliente__ie, cliente__nome
    dim cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep
    dim cliente__tel_res, cliente__ddd_res, cliente__tel_com, cliente__ddd_com, cliente__ramal_com, cliente__tel_cel, cliente__ddd_cel
    dim cliente__tel_com_2, cliente__ddd_com_2, cliente__ramal_com_2, cliente__email

    cliente__tipo = r_cliente.tipo
    cliente__cnpj_cpf = r_cliente.cnpj_cpf
    cliente__rg = r_cliente.rg
    cliente__ie = r_cliente.ie
    cliente__nome = r_cliente.nome
    cliente__endereco = r_cliente.endereco
    cliente__endereco_numero = r_cliente.endereco_numero
    cliente__endereco_complemento = r_cliente.endereco_complemento
    cliente__bairro = r_cliente.bairro
    cliente__cidade = r_cliente.cidade
    cliente__uf = r_cliente.uf
    cliente__cep = r_cliente.cep
    cliente__tel_res = r_cliente.tel_res
    cliente__ddd_res = r_cliente.ddd_res
    cliente__tel_com = r_cliente.tel_com
    cliente__ddd_com = r_cliente.ddd_com
    cliente__ramal_com = r_cliente.ramal_com
    cliente__tel_cel = r_cliente.tel_cel
    cliente__ddd_cel = r_cliente.ddd_cel
    cliente__tel_com_2 = r_cliente.tel_com_2
    cliente__ddd_com_2 = r_cliente.ddd_com_2
    cliente__ramal_com_2 = r_cliente.ramal_com_2
    cliente__email = r_cliente.email

    if isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos and r_pedido.st_memorizacao_completa_enderecos <> 0 then 
        cliente__tipo = r_pedido.endereco_tipo_pessoa
        cliente__cnpj_cpf = r_pedido.endereco_cnpj_cpf
	    cliente__rg = r_pedido.endereco_rg
        cliente__ie = r_pedido.endereco_ie
        cliente__nome = r_pedido.endereco_nome
        cliente__endereco = r_pedido.endereco_logradouro
        cliente__endereco_numero = r_pedido.endereco_numero
        cliente__endereco_complemento = r_pedido.endereco_complemento
        cliente__bairro = r_pedido.endereco_bairro
        cliente__cidade = r_pedido.endereco_cidade
        cliente__uf = r_pedido.endereco_uf
        cliente__cep = r_pedido.endereco_cep
        cliente__tel_res = r_pedido.endereco_tel_res
        cliente__ddd_res = r_pedido.endereco_ddd_res
        cliente__tel_com = r_pedido.endereco_tel_com
        cliente__ddd_com = r_pedido.endereco_ddd_com
        cliente__ramal_com = r_pedido.endereco_ramal_com
        cliente__tel_cel = r_pedido.endereco_tel_cel
        cliente__ddd_cel = r_pedido.endereco_ddd_cel
        cliente__tel_com_2 = r_pedido.endereco_tel_com_2
        cliente__ddd_com_2 = r_pedido.endereco_ddd_com_2
        cliente__ramal_com_2 = r_pedido.endereco_ramal_com_2
        cliente__email = r_pedido.endereco_email
        end if


	
%>
<%	if cliente__tipo = ID_PF then s_aux="CPF" else s_aux="CNPJ"
	s = cnpj_cpf_formata(cliente__cnpj_cpf) 
%>
		<td align="left" width="50%" class="MD"><p class="Rf"><%=s_aux%></p>
		<% if operacao_permitida(OP_LJA_EDITA_CLIENTE_DADOS_CADASTRAIS, s_lista_operacoes_permitidas)then %>
			<a href='javascript:fCLIEdita();' title='clique para editar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
		<% else %>
			<a href='javascript:fCLIConsulta();' title='clique para consultar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
		<% end if %>
		</td>
		<%
		if cliente__tipo = ID_PF then s = Trim(cliente__rg) else s = Trim(cliente__ie)
			if cliente__tipo = ID_PF then 
%>
	<td align="left" class="MD"><p class="Rf">RG</p><p class="C"><%=s%>&nbsp;</p></td>
<% else %>
	<td align="left" class="MD"><p class="Rf">IE</p><p class="C"><%=s%>&nbsp;</p></td>
<% end if %>
    <td align="center" valign="middle" style="width:15px"><a href='javascript:fCLIConsultaView(<%=chr(34) & r_cliente.id & chr(34) & "," & chr(34) & usuario & chr(34)%>);' title="clique para visualizar o cadastro do cliente"><img id="imgClienteConsultaView" src="../imagem/doc_preview_22.png" /></a></td>
		</tr>
<%
		
		if Trim(cliente__nome) <> "" then
			s = Trim(cliente__nome)
			end if
		end if
	
	if cliente__tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZ�O SOCIAL DO CLIENTE"
%>
    <tr>
	<td class="MC" align="left" colspan="3"><p class="Rf"><%=s_aux%></p>
	<% if operacao_permitida(OP_LJA_EDITA_CLIENTE_DADOS_CADASTRAIS, s_lista_operacoes_permitidas)then %>
		<a href='javascript:fCLIEdita();' title='clique para editar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
	<% else %>
		<a href='javascript:fCLIConsulta();' title='clique para consultar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
	<% end if %>
		</td>
	</tr>
	</table>

<!--  ENDERE�O DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%	
	s = formata_endereco(cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep)
%>		
		<td align="left"><p class="Rf">ENDERE�O</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
</table>

<!--  TELEFONE DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%	s = ""
	if Trim(cliente__tel_res) <> "" then
		s = telefone_formata(Trim(cliente__tel_res))
		s_aux=Trim(cliente__ddd_res)
		if s_aux<>"" then s = "(" & s_aux & ") " & s
		end if
	
	s2 = ""
	if Trim(cliente__tel_com) <> "" then
		s2 = telefone_formata(Trim(cliente__tel_com))
		s_aux = Trim(cliente__ddd_com)
		if s_aux<>"" then s2 = "(" & s_aux & ") " & s2
		s_aux = Trim(cliente__ramal_com)
		if s_aux<>"" then s2 = s2 & "  (R. " & s_aux & ")"
		end if
	if Trim(cliente__tel_cel) <> "" then
		s3 = telefone_formata(Trim(cliente__tel_cel))
		s_aux = Trim(cliente__ddd_cel)
		if s_aux<>"" then s3 = "(" & s_aux & ") " & s3
		end if
	if Trim(cliente__tel_com_2) <> "" then
		s4 = telefone_formata(Trim(cliente__tel_com_2))
		s_aux = Trim(cliente__ddd_com_2)
		if s_aux<>"" then s4 = "(" & s_aux & ") " & s4
		s_aux = Trim(cliente__ramal_com_2)
		if s_aux<>"" then s4 = s4 & "  (R. " & s_aux & ")"
		end if
	
%>

<% if cliente__tipo = ID_PF then %>
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
<%  notPrint = ""
    if Trim(cliente__email) = "" then notPrint=" notPrint" %>
<table width="649" class="QS<%=notPrint%>" cellspacing="0">
	<tr>
		<td align="left"><p class="Rf">E-MAIL</p><p class="C"><%=Trim(cliente__email)%>&nbsp;</p></td>
	</tr>
</table>

<!--  ENDERE�O DE ENTREGA  -->
<%  notPrint = ""
    if Trim(r_pedido.EndEtg_endereco) = "" then notPrint=" notPrint" %>
<%	
	s = pedido_formata_endereco_entrega(r_pedido, r_cliente)
%>		
<table width="649" class="QS<%=notPrint%>" cellspacing="0" style="table-layout:fixed">
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
	<tr bgColor="#FFFFFF">
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
	<% if (i > Lbound(v_item)) And (s_produto = "") then %>
	<tr class="notPrint">
	<% else %>
	<tr>
	<% end if %>
	<td class="MDBE" align="left"><input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:25px; color:<%=s_cor%>"
		value='<%=s_fabricante%>' readonly tabindex=-1></td>
	<td class="MDB" align="left"><input name="c_produto" id="c_produto" class="PLLe" style="width:54px; color:<%=s_cor%>"
		value='<%=s_produto%>' readonly tabindex=-1></td>
	<td class="MDB"  style="width:269px;" align="left">
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
			<td width="20%" align="left">&nbsp;</td>
			<% if blnTemRA Or ((r_pedido.permite_RA_status = 1) And (r_pedido.opcao_possui_RA = "S")) then %>
			<td align="right">
				<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
					<tr>
					<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;RA L�quido</span></td>
					<td class="MTBD" align="right"><input name="c_total_RA_Liquido" id="c_total_RA_Liquido" class="PLLd" style="width:70px;color:<%if r_pedido.vl_total_RA_liquido >=0 then Response.Write " green" else Response.Write " red"%>;" 
						value='<%=formata_moeda(r_pedido.vl_total_RA_liquido)%>' readonly tabindex=-1></td>
					</tr>
				</table>
			</td>
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

<% if r_pedido.tipo_parcelamento = 0 then %>
<!--  TRATA VERS�O ANTIGA DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" style="width:649px;" cellspacing="0">
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observa��es </p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:642px;margin-left:2pt;" 
				readonly tabindex=-1><%=r_pedido.obs_1%></textarea>
			<span class="PLLe notVisible"><%
				s = substitui_caracteres(r_pedido.obs_1,chr(13),"<br>")
				if s = "" then s = "&nbsp;"
				Response.Write s %></span>
		</td>
	</tr>
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">N� Nota Fiscal</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" style="width:85px;margin-left:2pt;" 
				readonly tabindex=-1 value='<%=r_pedido.obs_2%>'>
		</td>
	</tr>
	<tr>
		<td class="MDB" nowrap width="10%" align="left"><p class="Rf">Parcelas</p>
			<input name="c_qtde_parcelas" id="c_qtde_parcelas" class="PLLc" style="width:60px;"
				readonly tabindex=-1 value='<%if (r_pedido.qtde_parcelas<>0) Or (r_pedido.forma_pagto<>"") then Response.write Cstr(r_pedido.qtde_parcelas)%>'>
		</td>
		<td class="MDB" nowrap align="left" valign="top"><p class="Rf">Entrega Imediata</p>
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
		<td class="MDB" nowrap align="left" valign="top"><p class="Rf">Bem de Uso/Consumo</p>
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
		<td class="MDB" nowrap align="left" valign="top"><p class="Rf">Instalador Instala</p>
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
		<td class="MB tdGarInd" nowrap align="left" valign="top"><p class="Rf">Garantia Indicador</p>
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
	<tr>
		<td colspan="5" align="left"><p class="Rf">Forma de Pagamento</p>
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
<% else %>
<!--  TRATA NOVA VERS�O DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" style="width:649px;" cellspacing="0">
	<tr>
		<td class="MB" align="left"><p class="Rf">Observa��es </p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:99%;margin-left:2pt;" 
				readonly tabindex=-1><%=r_pedido.obs_1%></textarea>
			<span class="PLLe notVisible"><%
				s = substitui_caracteres(r_pedido.obs_1,chr(13),"<br>")
				if s = "" then s = "&nbsp;"
				Response.Write s %></span>
		</td>
	</tr>
<%  notPrint = ""
    if Trim(r_pedido.NFe_texto_constar) = "" then notPrint="class='notPrint'" %>
    <tr <%=notPrint%>>
		<td class="MB" align="left"><p class="Rf">Constar na NF</p>
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
		<td width="100%">
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td class="MB MD" align="left" nowrap width="40%"><p class="Rf">xPed</p>
						<input name="c_num_pedido_compra" id="c_num_pedido_compra" class="PLLe" maxlength="15" style="width:100px;margin-left:2pt;" onkeypress="filtra_nome_identificador();" onblur="this.value=trim(this.value);"
							value='<%=r_pedido.NFe_xPed%>' readonly tabindex=-1>
					</td>
					<td class="MB" align="left">
						<p class="Rf">Previs�o de Entrega</p>
						<% s = formata_data_e_talvez_hora_hhmm(r_pedido.PrevisaoEntregaData)
							if s <> "" then s = s & " &nbsp; (" & iniciais_em_maiusculas(r_pedido.PrevisaoEntregaUsuarioUltAtualiz) & " - " & formata_data_e_talvez_hora_hhmm(r_pedido.PrevisaoEntregaDtHrUltAtualiz) & ")"
							if s="" then s="&nbsp;"
						%>
						<p class="C"><%=s%></p>
					</td>
				</tr>
			</table>
		</td>
    </tr>
	<tr>
		<td width="100%">
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td class="MB MD" nowrap align="left" valign="top" width="40%"><p class="Rf">Entrega Imediata</p>
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
					<span class="C" style="margin-top:3px;"><%=s%></span>
					</td>
					<td class="MB MD" nowrap align="left" valign="top" width="20%"><p class="Rf">Bem Uso/Consumo</p>
					<% 	if Cstr(r_pedido.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then
							s = "N�O"
						elseif Cstr(r_pedido.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then
							s = "SIM"
						else
							s = ""
							end if
			
						if s="" then s="&nbsp;"
					%>
					<span class="C" style="margin-top:3px;"><%=s%></span>
					</td>
					<td class="MB MD" nowrap align="left" valign="top" width="20%"><p class="Rf">Instalador Instala</p>
					<% 	if Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_NAO) then
							s = "N�O"
						elseif Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
							s = "SIM"
						else
							s = ""
							end if
			
						if s="" then s="&nbsp;"
					%>
					<span class="C" style="margin-top:3px;"><%=s%></span>
					</td>
					<td class="MB tdGarInd" nowrap align="left" valign="top" width="20%"><p class="Rf">Garantia Indicador</p>
					<% 	if Cstr(r_pedido.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then
							s = "N�O"
						elseif Cstr(r_pedido.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then
							s = "SIM"
						else
							s = ""
							end if
			
						if s="" then s="&nbsp;"
					%>
					<span class="C" style="margin-top:3px;"><%=s%></span>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%">
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td class="MD" align="left" nowrap width="33.3%"><p class="Rf">N� Nota Fiscal</p>
						<% s_link_rastreio = monta_link_rastreio(pedido_selecionado, r_pedido.obs_2, r_pedido.transportadora_id, r_pedido.loja) %>
						<input name="c_obs2" id="c_obs2" class="PLLe" style="width:67px;margin-left:2pt;" 
							readonly tabindex=-1 value='<%=r_pedido.obs_2%>'><%=s_link_rastreio%>
					</td>
					<td class="MD" align="left" nowrap width="33.3%"><p class="Rf">NF Simples Remessa</p>
						<% s_link_rastreio = monta_link_rastreio(pedido_selecionado, r_pedido.obs_3, r_pedido.transportadora_id, r_pedido.loja) %>
						<input name="c_obs3" id="c_obs3" class="PLLe" style="width:67px;margin-left:2pt;" 
							readonly tabindex=-1 value='<%=r_pedido.obs_3%>'><%=s_link_rastreio%>
					</td>
					<td nowrap align="left" width="33.3%"><p class="Rf">NF Entrega Futura</p>
						<input name="c_obs4" id="c_obs4" class="PLLe" style="width:75px;margin-left:2pt;" 
							readonly tabindex=-1 value='<%=r_pedido.obs_4%>'>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<% if ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then %>
	<tr>
		<td class="MC" align="left" nowrap><p class="Rf">Referente Pedido Bonshop</p>
			<input name="c_ped_bonshop" id="c_ped_bonshop" class="PLLe" style="width:100px;margin-left:2pt;height:20px" 
				readonly tabindex=-1 value='<%=r_pedido.pedido_bs_x_at%>'>
		</td>
	</tr>
	<% end if %>
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
  <% if blnPossuiFormaPagtoProporcional then %>
  <tr>
	<td align="left" class="MC"><span class="Rf">Forma de Pagamento Proporcional Deste Pedido</span></td>
  </tr>
  <tr>
	<td align="left">
	  <table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C" style="display:inline-block;"><%=sDescricaoFormaPagtoProporcional%></span></td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
  <% end if %>
<%  notPrint = ""
    if Trim(r_pedido.forma_pagto) = "" then notPrint="class='notPrint'" %>
    <tr <%=notPrint%>>
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
<% end if %>


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
<% if r_pedido.PagtoAntecipadoStatus <> 0 then %>
<tr>
	<td colspan="3" class="MC MD" align="left" valign="bottom"><span class="Rf">Condi��o Pagto</span></td>
	<td colspan="3" class="MC" align="left" valign="bottom"><span class="Rf">Status Pagto Antecipado</span></td>
</tr>
<tr>
	<td colspan="3" class="MD" align="left"><span class="C"><%=pagto_antecipado_descricao(r_pedido.PagtoAntecipadoStatus)%></span></td>
	<td colspan="3" align="left"><span class="C" style="color:<%=pagto_antecipado_quitado_cor(r_pedido.PagtoAntecipadoStatus, r_pedido.PagtoAntecipadoQuitadoStatus)%>;"><%=pagto_antecipado_quitado_descricao(r_pedido.PagtoAntecipadoStatus, r_pedido.PagtoAntecipadoQuitadoStatus)%></span></td>
</tr>
<% end if %>
</table>


<!--  AN�LISE DE CR�DITO   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<%	s=x_analise_credito(r_pedido.analise_credito)
		if s <> "" then
            if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE_VENDAS) then 
                if r_pedido.analise_credito_pendente_vendas_motivo <> "" then s = s & " (" & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__AC_PENDENTE_VENDAS_MOTIVO, r_pedido.analise_credito_pendente_vendas_motivo) & ")"  
                end if 
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
<%  notPrint = ""
    if Trim(formata_data(r_pedido.a_entregar_data_marcada)) = "" then notPrint=" notPrint" %>
<br>
<table width="649" class="Q<%=notPrint%>" cellspacing="0">
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
	<td align="left"><p class="Rf">TRANSPORTADORA</p><p class="C"><%=s%></p></td>
	
<!--   FRETES   -->

    <%  s = "SELECT * FROM t_PEDIDO_FRETE WHERE pedido='" & r_pedido.pedido & "' ORDER BY dt_cadastro" 
        x = ""
        intQtdeFrete = 0
        vl_total_frete = 0
        set rs = cn.execute(s)

        do while Not rs.Eof
            frete_transportadora_id = Trim("" & rs("transportadora_id"))
            frete_numero_NF = Trim("" & rs("numero_NF"))
            frete_serie_NF = Trim("" & rs("serie_NF"))
            if frete_numero_NF = "0" then frete_numero_NF = ""
            if frete_serie_NF = "0" then 
                frete_serie_NF = ""
            else
                frete_serie_NF = NFeFormataSerieNF(frete_serie_NF)
            end if
            if intQtdeFrete > 0 then x = x & "</tr><tr>" & chr(13)
            
            x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & UCase(rs("transportadora_id")) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:150px;'><span class='C'>" & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE, rs("codigo_tipo_frete")) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & obtem_apelido_empresa_NFe_emitente(rs("id_nfe_emitente")) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:80px;'><span class='C'>" & frete_numero_NF & "</td>" & chr(13)
                x = x & "<td class='MD MB' align='center' style='width:50px;'><span class='C'>" & frete_serie_NF & "</td>" & chr(13)
                x = x & "<td class='MB' align='right' style='width:97px;padding-right: 5px'><span class='C'>" & formata_moeda(rs("vl_frete")) & "</td>" & chr(13)
            
            
            intQtdeFrete = intQtdeFrete + 1
            vl_total_frete = vl_total_frete + rs("vl_frete")
        rs.MoveNext
        loop
        s = formata_moeda(vl_total_frete) 
    %>

	
	

</tr>
</table>
<br />
<%  notPrint = ""
    if intQtdeFrete = 0 then notPrint=" notPrint" %>
<table id="tFretes" width="649" class="Q<%=notPrint%>" cellspacing="0" style="border-bottom:0">
    <tr>
        <td class="MB" align="left" style="width:130px;" colspan="6"><p class="Rf">FRETES</p></td>

    </tr>
    <tr>
        <td class="MD MB" align="center" style="width:130px;"><p class="Rf">TRANSPORTADORA</p></td>
        <td class="MD MB" align="center" style="width:150px;"><p class="Rf">TIPO DE FRETE</p></td>
        <td class="MD MB" align="center" style="width:130px;"><p class="Rf">EMITENTE</p></td>
        <td class="MD MB" align="center" style="width:80px;"><p class="Rf">N�MERO NF</p></td>
        <td class="MD MB" align="center" style="width:80px;"><p class="Rf">S�RIE NF</p></td>
        <td class="MB" align="right" style="width:50px;padding-right: 5px"><p class="Rf">VALOR</p></td>

    </tr>
    <tr>
        <%=x%>
    </tr>
    <tr>
        <td class="MB MD" colspan="5" align="right" valign="bottom"><p class="Cd">TOTAL</p></td>
        <td class="MB" align="right" style="width:65px;padding-right: 5px">
            <p class="Cd"><%=s%></p>
	</td>
    </tr>
</table>
<% end if %>
<%if r_pedido.st_entrega = ST_ENTREGA_CANCELADO then%>
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td width="50%" class="MD" align="left" valign="bottom"><span class="Rf">Data do Cancelamento</span></td>
	<td width="50%"  align="left" valign="bottom"><span class="Rf">Usu�rio</span></td>   
</tr>
<tr>
	<% s = formata_data_hora(r_pedido.cancelado_data_hora) %>
	<td width="324px" class="MD" align="left"><span class="C"><%=s%></span></td>
	<% s = r_pedido.cancelado_usuario %>
	<td width="324px"  align="left"><span class="C"><%=s%>&nbsp;</span></td>		
</tr>
<tr>
    <td width="60%" align="left"  class="MC" valign="bottom" colspan="2"><span class="Rf">Causa</span></td>
</tr>
<tr>
    <%if r_pedido.cancelado_codigo_motivo <> "" then %>
    <% s = "SELECT descricao from t_CODIGO_DESCRICAO where codigo = " & r_pedido.cancelado_codigo_motivo & " AND grupo = '" & GRUPO_T_CODIGO_DESCRICAO__CANCELAMENTOPEDIDO_MOTIVO & "'" 
       set rs = cn.execute(s)
        s = rs("descricao")%>
    <%else
        s = ""%>
    <%end if%>
    <% if r_pedido.cancelado_codigo_sub_motivo <> "" then
        s_sql = "SELECT descricao FROM t_CODIGO_DESCRICAO where grupo = 'CancelamentoPedido_Motivo_Sub' AND codigo = " & r_pedido.cancelado_codigo_sub_motivo & "" 
        set rs = cn.execute(s_sql)
        s = s & " (" & rs("descricao") & ") "
        end if  %>
	<td width="270px" align="left" ><span class="C"><%=s%>&nbsp;</span></td>
</tr>
<tr>
    <td width="100%" class='MC' align="left" valign="bottom" colspan="3"><span class="Rf">Descri��o/Motivo</span> 
	    <textarea name="c_motivo" id="c_motivo" class="PLLe notPrint" rows="5"
				    style="width:642px;margin-left:2pt;"
				    readonly tabindex=-1><%=r_pedido.cancelado_motivo%></textarea>
	    
	    <span class="PLLe notVisible"><%
			    s = substitui_caracteres(r_pedido.cancelado_motivo,chr(13),"<br>")
			    if s = "" then s = "&nbsp;"
			    Response.Write s %></span>	    
    </td>
</tr>
</table>
<%end if %>

<!-- DEVOLU��O -->

<% if operacao_permitida(OP_LJA_PRE_DEVOLUCAO_LEITURA, s_lista_operacoes_permitidas) then %>
<br id="brDevolucao" class="notPrint">
<a name="aPedidoDevolucao"></a>
<table id="tableDevolucao" class="notPrint" width="649" cellspacing="0" cellpadding="0" border="0">
<tr>
	<td colspan="4" class="ME MD MC MB" align="left"><span class="Rf">DEVOLU��ES</span></td>
</tr>
<% s = "SELECT" & _
			" t_PEDIDO_DEVOLUCAO.usuario_cadastro AS devolucao_usuario," & _
            " t_PEDIDO_DEVOLUCAO.dt_hr_cadastro AS devolucao_dt_hr_cadastro," & _
            " t_PEDIDO_DEVOLUCAO.id AS devolucao_id," & _
			"*" & _
	   " FROM t_PEDIDO_DEVOLUCAO INNER JOIN t_PEDIDO ON (t_PEDIDO.pedido = t_PEDIDO_DEVOLUCAO.pedido)" & _  
	   " WHERE" & _
			" (t_PEDIDO_DEVOLUCAO.pedido = '" & pedido_selecionado & "')" & _
	   " ORDER BY" & _
			" t_PEDIDO_DEVOLUCAO.dt_hr_cadastro," & _
			" id"
	set rs = cn.execute(s)
	if rs.Eof then %>
		<tr class="notVisible">
			<td colspan="3" class="ME MD MB" align="left">&nbsp;</td>
		</tr>
<%		end if
	n_offset_tabela_devolucao = 24
	do while Not rs.Eof
		if CInt(rs("st_finalizado"))=0 then blnHaOcorrenciaEmAberto=True
%>
	<tr>
		<td class="ME" style="width:<%=n_offset_tabela_devolucao%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" style="width:<%=649-3-n_offset_tabela_devolucao%>px;" align="left">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td class="C MD MB tdWithPadding" width="11%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">ID:</span>&nbsp;<%=Trim("" & rs("devolucao_id"))%></td>
			<td class="C MD MB tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">CADASTRADA POR:</span>&nbsp;<%=Trim("" & rs("devolucao_usuario"))%></td>
			<td class="C MD MB tdWithPadding" width="20%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">EM:</span>&nbsp;<%=formata_data_hora_sem_seg(rs("devolucao_dt_hr_cadastro"))%></td>
			<%	
                obtem_descricao_status_devolucao rs("status"), s, s_cor
			%>
			<td class="C MB tdWithPadding" align="left" valign="top" style="color:<%=s_cor%>"><span class="Rf" style="margin-left:0px;">SITUA��O:</span>&nbsp;<%=UCase(s)%></td>
			</tr>
			</table>
		</td>

	</tr>

    <% if CStr(rs("status")) = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA then %>
    <tr>
		<td class="ME MB" style="width:<%=n_offset_tabela_devolucao%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" align="left">
			<table width="100%" cellspacing="0" cellpadding="1">
			<tr>
			<td class="C MB" width="100%" align="left" valign="top">
				<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
				<td class="C tdWithPadding" width="33%" align="left" valign="top"><span class="C" style="margin-left:0px;color: darkgoldenrod">Aguardando aprova��o</td>
				</tr>
				</table>
			</td>
			</tr>
            </table>
		</td>
	</tr>
    <% elseif CStr(rs("status")) = COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO Or CStr(rs("status")) = COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA then %>
    <tr>
		<td class="ME MB" style="width:<%=n_offset_tabela_devolucao%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" align="left">
			<table width="100%" cellspacing="0" cellpadding="1">
			<tr>
			<td class="C MB" width="100%" align="left" valign="top">
				<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
				<td class="C MD tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">APROVADA POR:</span>&nbsp;<%=Trim("" & rs("usuario_aprovado"))%></td>
				<td class="C tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">EM:</span>&nbsp;<%=formata_data_hora_sem_seg(rs("dt_hr_aprovado"))%></td>
                <td>&nbsp;</td>
				</tr>
				</table>
			</td>
			</tr>
            </table>
		</td>
	</tr>
    <% elseif CStr(rs("status")) = COD_ST_PEDIDO_DEVOLUCAO__FINALIZADA then %>
    <tr>
		<td class="ME MB" style="width:<%=n_offset_tabela_devolucao%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" align="left">
			<table width="100%" cellspacing="0" cellpadding="1">
			<tr>
			<td class="C MB" width="100%" align="left" valign="top">
				<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
				<td class="C MD tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">FINALIZADA POR:</span>&nbsp;<%=Trim("" & rs("usuario_finalizado"))%></td>
				<td class="C tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">EM:</span>&nbsp;<%=formata_data_hora_sem_seg(rs("dt_hr_finalizado"))%></td>
                <td>&nbsp;</td>
				</tr>
				</table>
			</td>
			</tr>
            </table>
		</td>
	</tr>
    <% elseif CStr(rs("status")) = COD_ST_PEDIDO_DEVOLUCAO__REPROVADA then %>
    <tr>
		<td class="ME MB" style="width:<%=n_offset_tabela_devolucao%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" align="left">
			<table width="100%" cellspacing="0" cellpadding="1">
			<tr>
			<td class="C MB" width="100%" align="left" valign="top">
				<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
				<td class="C MD tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">REPROVADA POR:</span>&nbsp;<%=Trim("" & rs("usuario_reprovado"))%></td>
				<td class="C tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">EM:</span>&nbsp;<%=formata_data_hora_sem_seg(rs("dt_hr_reprovado"))%></td>
                <td>&nbsp;</td>
				</tr>
				</table>
			</td>
			</tr>
            </table>
		</td>
	</tr>
    <% elseif CStr(rs("status")) = COD_ST_PEDIDO_DEVOLUCAO__CANCELADA then %>
    <tr>
		<td class="ME MB" style="width:<%=n_offset_tabela_devolucao%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" align="left">
			<table width="100%" cellspacing="0" cellpadding="1">
			<tr>
			<td class="C MB" width="100%" align="left" valign="top">
				<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
				<td class="C MD tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">CANCELADA POR:</span>&nbsp;<%=Trim("" & rs("usuario_cancelado"))%></td>
				<td class="C tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">EM:</span>&nbsp;<%=formata_data_hora_sem_seg(rs("dt_hr_cancelado"))%></td>
                <td>&nbsp;</td>
				</tr>
				</table>
			</td>
			</tr>
            </table>
		</td>
	</tr>

    <% end if %>
	<tr>
		<td class="ME" style="width:<%=n_offset_tabela_devolucao%>px;" align="left">&nbsp;</td>
		<td colspan="3" class="ME MD" style="width:<%=649-3-n_offset_tabela_devolucao%>px;" align="left">
			<table width="100%" cellspacing="0" cellpadding="1">
			<tr>
			<%	s = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__PROCEDIMENTO, Trim("" & rs("cod_procedimento"))) %>
			<td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">PROCEDIMENTO:</span>&nbsp;<%=s%></td>
			</tr>
			<tr>
            <% s = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__MOTIVO, Trim("" & rs("cod_devolucao_motivo"))) %>
			    <td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">MOTIVO:</span>&nbsp;<%=s%></td>
            </tr>
            <tr>
			    <td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">DESCRI��O:</span>&nbsp;<%=substitui_caracteres(Trim("" & rs("motivo_observacao")), chr(13), "<br>")%></td>
            </tr>
			</table>
		</td>
	</tr>

<% s = "SELECT " & _
			"t_PEDIDO_DEVOLUCAO_ITEM.fabricante," & _
			"t_PEDIDO_DEVOLUCAO_ITEM.produto," & _
			"t_PEDIDO_DEVOLUCAO_ITEM.qtde," & _
			"t_PEDIDO_DEVOLUCAO_ITEM.qtde_estoque_venda," & _
			"t_PEDIDO_DEVOLUCAO_ITEM.qtde_estoque_danificado," & _
			"t_PEDIDO_DEVOLUCAO_ITEM.vl_unitario," & _
            "t_PRODUTO.descricao," & _
            "t_PRODUTO.descricao_html" & _
	   " FROM t_PEDIDO_DEVOLUCAO_ITEM" & _
       " INNER JOIN t_PRODUTO ON ((t_PEDIDO_DEVOLUCAO_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_PEDIDO_DEVOLUCAO_ITEM.produto=t_PRODUTO.produto))" & _
	   " WHERE" & _
			" (id_pedido_devolucao = " & Trim("" & rs("id")) & ")" & _
	   " ORDER BY" & _
			" fabricante," & _
			" produto"
	set rs2 = cn.execute(s)

    vlTotalDevolucao = 0
%>
	<tr>
		<%	if CInt(rs("st_finalizado"))=0 Or CInt(rs("st_reprovado"))=0 Or CInt(rs("st_cancelado"))=0 then s="ME MB" else s="ME" %>
		<td class="<%=s%>" style="width:<%=n_offset_tabela_devolucao%>px;" align="left">&nbsp;</td>
		<td class="<%=s%>" style="width:<%=n_offset_tabela_devolucao%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" style="width:<%=649-3-2*n_offset_tabela_devolucao%>px;" align="left">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td colspan="8" class="Rf tdWithPadding" align="left">ITENS DEVOLVIDOS</td>
			</tr>
            <tr>
                 <TD class='MTD MB' style='vertical-align:bottom;width: 50px;'><P class='Rc'>Fabricante</P></TD>
                <TD class='MTD MB' style='vertical-align:bottom;width: 60px;'><P class='Rc'>Produto</P></TD>
                <TD class='MTD MB' style='vertical-align:bottom;width: 240px;padding-left: 3px;'><P class='R'>Descri��o</P></TD>
                <TD class='MTD MB' style='vertical-align:bottom;width: 35px;' align='right'><P class='R'>Qtde</P></TD>
                <TD class='MTD MB' style='vertical-align:bottom;width: 35px;' align='right'><P class='R'>Estoque Venda</P></TD>
                <TD class='MTD MB' style='vertical-align:bottom;width: 35px;' align='right'><P class='R'>Estoque Danif</P></TD>
                <TD class='MTD MB' style='vertical-align:bottom;width: 50px;'  align='right'><P class='R'>VL Unit�rio</P></TD>
                <TD class='MC MB' style='vertical-align:bottom;width: 50px;'  align='right'><P class='R'>VL Total Devol</P></TD>
            </tr>
			<%	do while not rs2.Eof 
            vlTotalItemDevolucao = converte_numero(rs2("vl_unitario"))*converte_numero(rs2("qtde"))
            vlTotalDevolucao = vlTotalDevolucao+vlTotalItemDevolucao
            %>
			<tr>
				<td class="C MD MB" style="width:60px;" align="center" valign="top"><%=rs2("fabricante")%></td>
				<td class="C MD MB tdWithPadding" align="center" valign="top"><%=rs2("produto")%></td>
				<td class="C MD MB tdWithPadding" style="padding-left: 3px;" align="left" valign="top"><%=Trim("" & rs2("descricao_html"))%></td>
				<td class="C MD MB tdWithPadding" align="right" valign="top"><%=rs2("qtde")%></td>
				<td class="C MD MB tdWithPadding" align="right" valign="top"><%=Trim("" & rs2("qtde_estoque_venda"))%></td>
				<td class="C MD MB tdWithPadding" align="right" valign="top"><%=Trim("" & rs2("qtde_estoque_danificado"))%></td>
				<td class="C MD MB tdWithPadding" align="right" valign="top"><%=formata_moeda(Trim("" & rs2("vl_unitario")))%></td>
				<td class="C MB tdWithPadding" align="right" valign="top"><%=formata_moeda(vlTotalItemDevolucao)%></td>
			</tr>
			<%		rs2.MoveNext
					loop 
			%>
			
            <tr>
                <td colspan="7" class="C MB tdWithPadding" align="right" valign="top">Total:</td>
                <td class="C MB tdWithPadding" align="right" valign="top"><%=formata_moeda(vlTotalDevolucao)%></td>
            </tr>

			</table>
		</td>
	</tr>

<% s = "SELECT " & _
			"*" & _
	   " FROM t_PEDIDO_DEVOLUCAO_MENSAGEM" & _
	   " WHERE" & _
			" (id_pedido_devolucao = " & Trim("" & rs("id")) & ")" & _
	   " ORDER BY" & _
			" dt_hr_cadastro DESC," & _
			" id"
	set rs2 = cn.execute(s)
%>
	<tr>
		<%	if CInt(rs("st_finalizado"))=0 Or CInt(rs("st_reprovado"))=0 Or CInt(rs("st_cancelado"))=0 then s="ME MB" else s="ME" %>
		<td class="<%=s%>" style="width:<%=n_offset_tabela_devolucao%>px;" align="left">&nbsp;</td>
		<td class="<%=s%>" style="width:<%=n_offset_tabela_devolucao%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" style="width:<%=649-3-2*n_offset_tabela_devolucao%>px;" align="left">
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

			<% if CInt(rs("st_finalizado"))=0 And CInt(rs("st_reprovado"))=0 And CInt(rs("st_cancelado"))=0 then %>
			<tr class="notPrint">
				<td class="MC" style="padding:0px;" align="left">
					<table width="100%" cellpadding="0" cellspacing="0">
					<tr>
					<td align="left">&nbsp;</td>

					</tr>
					</table>
				</td>
			</tr>
            <% end if %>
			<tr class="notPrint">
				<td class="MB" align="left"><span style='font-family: Arial, Helvetica, sans-serif;color:white;font-size:6pt;font-style:normal;'>&nbsp;</span></td>
			</tr>
			<tr class="notVisible">
				<td class="MB MC" align="left"><span style='font-family: Arial, Helvetica, sans-serif;color:white;font-size:6pt;font-style:normal;'>&nbsp;</span></td>
			</tr>

			
			</table>
		</td>
	</tr>
<%
		rs.MoveNext
		loop
%>
	<tr class="notPrint">
		<td colspan="3" style="padding:0px;" align="left">
			<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
			    <td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="bDevolucaoAlteraImpressao" id="bDevolucaoAlteraImpressao" href="javascript:fPEDDevolucaoAlteraImpressao(fPED)" title="configura as informa��es sobre devolu��es para serem impressas ou n�o"><img id="imgPrinterDevolucao" src="../botao/PrinterError.png" border="0"></a></td>
			    <td align="left">&nbsp;</td>
                <td align="left">&nbsp;</td>
                <td align="left">&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>

</table>
<% end if %>

<% if operacao_permitida(OP_LJA_OCORRENCIAS_EM_PEDIDOS_LEITURA, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_LJA_OCORRENCIAS_EM_PEDIDOS_CADASTRAMENTO, s_lista_operacoes_permitidas) then %>
<br id="brOcorrencia" class="notPrint">
<table id="tableOcorrencia" class="notPrint" width="649" cellspacing="0" cellpadding="0" border="0">
<tr>
	<td colspan="3" class="ME MD MC MB" align="left"><span class="Rf">OCORR�NCIAS</span></td>
</tr>
<% s = "SELECT" & _
			" t_PEDIDO_OCORRENCIA.usuario_cadastro AS ocorrencia_usuario," & _
            " t_PEDIDO_OCORRENCIA.dt_hr_cadastro AS ocorrencia_dt_hr_cadastro," & _
			"*" & _
			"," & _
            " t_PEDIDO.loja AS pedido_loja," & _
			" (" & _
				"SELECT" & _
					" Count(*)" & _
				" FROM t_PEDIDO_OCORRENCIA_MENSAGEM" & _
				" WHERE" & _
					" (id_ocorrencia=t_PEDIDO_OCORRENCIA.id)" & _
					" AND (fluxo_mensagem='" & COD_FLUXO_MENSAGEM_OCORRENCIAS_EM_PEDIDOS__CENTRAL_PARA_LOJA & "')" & _
			") AS qtde_msg_central," & _
            " (" & _
                " SELECT Count(*)" & _
		           " FROM t_PEDIDO_OCORRENCIA_MENSAGEM" & _
		           " WHERE (id_ocorrencia = t_PEDIDO_OCORRENCIA.id)" & _
            ") AS qtde_msg" & _
	   " FROM t_PEDIDO_OCORRENCIA t_PEDIDO_OCORRENCIA LEFT JOIN t_CODIGO_DESCRICAO ON (t_PEDIDO_OCORRENCIA.cod_motivo_abertura=t_CODIGO_DESCRICAO.codigo) AND (t_CODIGO_DESCRICAO.grupo='" & GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__MOTIVO_ABERTURA & "')" & _
       " INNER JOIN t_PEDIDO ON (t_PEDIDO.pedido = t_PEDIDO_OCORRENCIA.pedido)" & _ 
	   " WHERE" & _
			" (t_PEDIDO_OCORRENCIA.pedido = '" & pedido_selecionado & "')" & _
	   " ORDER BY" & _
			" t_PEDIDO_OCORRENCIA.dt_hr_cadastro," & _
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
			<td class="C MD MB tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">ABERTA POR:</span>&nbsp;<%=Trim("" & rs("ocorrencia_usuario"))%></td>
			<td class="C MD MB tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">EM:</span>&nbsp;<%=formata_data_hora_sem_seg(rs("ocorrencia_dt_hr_cadastro"))%></td>
			<%	if CInt(rs("finalizado_status")) <> 0 then
					s_cor = "green"
					s = "Finalizada"
				else
					s_cor = "red"
					if CInt(rs("qtde_msg_central")) > 0 Or _
                        (Trim("" & rs("pedido_loja")) = NUMERO_LOJA_ECOMMERCE_AR_CLUBE And CInt(rs("qtde_msg")) > 0) then
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
                <% if Trim("" & rs("cod_motivo_abertura")) = "" then %>
			<td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">OCORR�NCIA:</span>&nbsp;<%=substitui_caracteres(Trim("" & rs("texto_ocorrencia")), chr(13), "<br>")%></td>
			    <% else %>
			<td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">OCORR�NCIA:</span>&nbsp;<%=Trim("" & rs("descricao"))%>
                <% if Trim("" & rs("texto_ocorrencia")) <> "" then %>
                <br /><br />
                <% Response.Write substitui_caracteres(Trim("" & rs("texto_ocorrencia")), chr(13), "<br>") 
                    end if%>
			</td>
                <% end if %>
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
<% end if %>


<!-- CHAMADOS //-->

<% s = "select * from t_pedido_chamado_depto where usuario_responsavel = '" & usuario & "' or usuario_gestor = '" & usuario & "' " 
    blnIsUsuarioResponsavelDepto = false

    if rs.State <> 0 then rs.Close
    rs.open s, cn
    if Not rs.Eof then
        blnIsUsuarioResponsavelDepto = true
    end if        

    s = "select count(*) as numero_chamados from t_pedido_chamado where pedido = '" & pedido_selecionado & "' and usuario_cadastro='" & usuario & "'"
    blnIsUsuarioCadastroChamado = false

    if rs.State <> 0 then rs.Close
    rs.open s, cn
    if Not rs.Eof then
        if CInt(rs("numero_chamados")) > 0 then blnIsUsuarioCadastroChamado=true
    end if    

%>

<% if operacao_permitida(OP_LJA_PEDIDO_CHAMADO_LEITURA_QUALQUER_CHAMADO, s_lista_operacoes_permitidas) Or _
    operacao_permitida(OP_LJA_PEDIDO_CHAMADO_ESCREVER_MSG_QUALQUER_CHAMADO, s_lista_operacoes_permitidas) Or _
    operacao_permitida(OP_LJA_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) Or _
    blnIsUsuarioResponsavelDepto Or _
    blnIsUsuarioCadastroChamado then %>

<br id="brChamado" class="notPrint">
<table id="tableChamado" class="notPrint" width="649" cellspacing="0" cellpadding="0" border="0">
<tr>
	<td colspan="3" class="ME MD MC MB" align="left"><span class="Rf">CHAMADOS</span></td>
</tr>
<% s = "SELECT t_PEDIDO_CHAMADO.usuario_cadastro AS chamado_usuario," & _
            "t_PEDIDO_CHAMADO.dt_hr_cadastro AS chamado_dt_hr_cadastro," & _
            "t_PEDIDO_CHAMADO.id AS chamado_id," & _
            "t_PEDIDO_CHAMADO_DEPTO.descricao AS depto," & _
			"*" & _
			"," & _
			" (" & _
				"SELECT" & _
					" Count(*)" & _
				" FROM t_PEDIDO_CHAMADO_MENSAGEM" & _
				" WHERE" & _
					" (id_chamado=t_PEDIDO_CHAMADO.id)" & _
					" AND (fluxo_mensagem='" & COD_FLUXO_MENSAGEM_CHAMADOS_EM_PEDIDOS__RX & "')" & _
			") AS qtde_msg_rx" & _
	   " FROM t_PEDIDO_CHAMADO" & _
       " LEFT JOIN t_PEDIDO_CHAMADO_DEPTO ON (t_PEDIDO_CHAMADO.id_depto=t_PEDIDO_CHAMADO_DEPTO.id)" & _
       " LEFT JOIN t_CODIGO_DESCRICAO ON (t_PEDIDO_CHAMADO.cod_motivo_abertura=t_CODIGO_DESCRICAO.codigo) AND (t_CODIGO_DESCRICAO.grupo='" & GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA & "')" & _
	   " WHERE" & _
			" (t_PEDIDO_CHAMADO.pedido = '" & pedido_selecionado & "')"
            
    if Not blnIsUsuarioResponsavelDepto And Not blnIsUsuarioCadastroChamado then    
        s = s & " AND (nivel_acesso <= '" & CStr(nivel_acesso_chamado) & "')"
    end if

    if Not operacao_permitida(OP_LJA_PEDIDO_CHAMADO_LEITURA_QUALQUER_CHAMADO, s_lista_operacoes_permitidas) And _
    Not operacao_permitida(OP_LJA_PEDIDO_CHAMADO_ESCREVER_MSG_QUALQUER_CHAMADO, s_lista_operacoes_permitidas) then
        s = s & " AND ((t_PEDIDO_CHAMADO_DEPTO.usuario_responsavel = '" & usuario & "')" & _
                " OR (t_PEDIDO_CHAMADO_DEPTO.usuario_gestor = '" & usuario & "')" & _
                " OR (t_PEDIDO_CHAMADO.usuario_cadastro = '" & usuario & "'))"
    end if

	  s = s & " ORDER BY" & _
			" t_PEDIDO_CHAMADO.dt_hr_cadastro," & _
			" t_PEDIDO_CHAMADO.id"

	set rs = cn.execute(s)
	if rs.Eof then %>
		<tr class="notVisible">
			<td colspan="3" class="ME MD MB" align="left">&nbsp;</td>
		</tr>
<%		end if
	n_offset_tabela_chamado = 24
	do while Not rs.Eof

    blnIsUsuarioCadastroChamado=false
    if Trim("" & rs("chamado_usuario")) = usuario then blnIsUsuarioCadastroChamado=true
%>
	<tr>
		<td class="ME" style="width:<%=n_offset_tabela_chamado%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" style="width:<%=649-3-n_offset_tabela_chamado%>px;" align="left">
			<table width="100%" cellspacing="0" cellpadding="0">
            <tr>
			    <td colspan="4" class="C MD MB tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">DEPARTAMENTO RESPONS�VEL:</span>&nbsp;<%=Trim("" & rs("depto"))%></td>
            </tr>
			<tr>
			<td class="C MD MB tdWithPadding" width="33%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">ABERTO POR:</span>&nbsp;<%=Trim("" & rs("chamado_usuario"))%></td>
			<td class="C MD MB tdWithPadding" width="20%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">EM:</span>&nbsp;<%=formata_data_hora_sem_seg(rs("chamado_dt_hr_cadastro"))%></td>
			<td class="C MD MB tdWithPadding" width="20%" align="left" valign="top"><span class="Rf" style="margin-left:0px;"></span>&nbsp;<%=nivel_acesso_chamado_pedido_descricao(rs("nivel_acesso"))%></td>
			<%	if CInt(rs("finalizado_status")) <> 0 then
					s_cor = "green"
					s = "Finalizado"
				else
					s_cor = "red"
					if CInt(rs("qtde_msg_rx")) > 0 then
						s = "Em Andamento"
					else
						s = "Aberto"
						end if
					end if
			%>
			<td class="C MB tdWithPadding" align="left" valign="top" style="color:<%=s_cor%>"><span class="Rf" style="margin-left:0px;">SITUA��O:</span>&nbsp;<%=UCase(s)%></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="ME" style="width:<%=n_offset_tabela_chamado%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="ME MD" style="width:<%=649-3-n_offset_tabela_chamado%>px;" align="left">
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
            <% s = ""
               if Trim("" & rs("cod_motivo_abertura")) <> "" then s = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA, Trim("" & rs("cod_motivo_abertura"))) %>
			<tr>
			    <td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">MOTIVO DA ABERTURA:</span>&nbsp;<%=s%></td>
            </tr>
            <tr>
			    <td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">DESCRI��O:</span>&nbsp;<%=substitui_caracteres(Trim("" & rs("texto_chamado")), chr(13), "<br>") %></td>
            </tr>
			</table>
		</td>
	</tr>
	
<% s = "SELECT " & _
			"*" & _
	   " FROM t_PEDIDO_CHAMADO_MENSAGEM" & _
	   " WHERE" & _
			" (id_chamado = " & Trim("" & rs("chamado_id")) & ")" & _
            " AND (nivel_acesso <= '" & CStr(nivel_acesso_chamado) & "')" & _
	   " ORDER BY" & _
			" dt_hr_cadastro," & _
			" id"
	set rs2 = cn.execute(s)
%>
	<tr>
		<%	if CInt(rs("finalizado_status"))=0 then s="ME MB" else s="ME" %>
		<td class="<%=s%>" style="width:<%=n_offset_tabela_chamado%>px;" align="left">&nbsp;</td>
		<td class="<%=s%>" style="width:<%=n_offset_tabela_chamado%>px;" align="left">&nbsp;</td>
		<td class="ME MD" style="width:<%=649-3-2*n_offset_tabela_chamado%>px;" align="left">
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
		            <td class="C MC MD" align="center" valign="top" style="width:50px;"><%=nivel_acesso_chamado_pedido_descricao(rs2("nivel_acesso"))%></td>
                    
					<td class="C MC tdWithPadding" align="left" valign="top"><%=substitui_caracteres(Trim("" & rs2("texto_mensagem")), chr(13), "<br>")%></td>
					</tr>
					</table>
				</td>
			</tr>
			<%		rs2.MoveNext
					loop 
			%>

			<%	if CInt(rs("finalizado_status"))=0 then s="MB" else s="" %>
                <tr class="notPrint">
				<td class="<%=s%>" colspan="3" style="padding:0px;" align="left">
					
				</td>
			</tr>

			
			</table>
		</td>
	</tr>

	<% if CInt(rs("finalizado_status")) <> 0 then %>
	<tr>
		<td class="ME MB" style="width:<%=n_offset_tabela_chamado%>px;" align="left">&nbsp;</td>
		<td colspan="2" class="MC ME MD" align="left">
			<table width="100%" cellspacing="0" cellpadding="1">
			<tr>
			<td class="C MB" width="100%" align="left" valign="top">
				<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
				<td class="C MD tdWithPadding" width="50%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">FINALIZADO POR:</span>&nbsp;<%=Trim("" & rs("finalizado_usuario"))%></td>
				<td class="C tdWithPadding" width="50%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">EM:</span>&nbsp;<%=formata_data_hora_sem_seg(rs("finalizado_data_hora"))%></td>
				</tr>
				</table>
			</td>
			</tr>
			<tr>
			<% s = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_FINALIZACAO, Trim("" & rs("cod_motivo_finalizacao"))) %>
			<td class="C MB" width="100%" align="left" valign="top"><span class="Rf" style="margin-left:0px;">MOTIVO DA FINALIZA��O:</span>&nbsp;<%=s%></td>
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
			<td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="bChamadoAlteraImpressao" id="bChamadoAlteraImpressao" href="javascript:fPEDChamadoAlteraImpressao(fPED)" title="configura as informa��es sobre chamados para serem impressos ou n�o"><img name="imgPrinterChamado" id="imgPrinterChamado" src="../botao/PrinterError.png" border="0"></a></td>
			<td align="left">&nbsp;</td>

			<td align="center" style="width:45px;padding:2px;">
				&nbsp;
			</td>

			</tr>
			</table>
		</td>
	</tr>

</table>
<%end if %>


<% if operacao_permitida(OP_LJA_BLOCO_NOTAS_PEDIDO_LEITURA, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_LJA_BLOCO_NOTAS_PEDIDO_CADASTRAMENTO, s_lista_operacoes_permitidas) then %>
<br id="brBlocoNotas" class="notPrint">
<table id="tableBlocoNotas" class="notPrint" width="649" cellspacing="0" cellpadding="1">
<tr>
	<td colspan="4" class="ME MD MC MB" align="left"><span class="Rf">BLOCO DE NOTAS</span></td>
</tr>
<% s = "SELECT " & _
			"*" & _
	   " FROM t_PEDIDO_BLOCO_NOTAS" & _
	   " WHERE" & _
			" (pedido = '" & pedido_selecionado & "')" & _
			" AND (nivel_acesso <= " & Cstr(nivel_acesso_bloco_notas) & ")" & _
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
		<% if converte_numero(nivel_acesso_bloco_notas) = converte_numero(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO) then %>
		<td colspan="2" class="C MD MB" align="left" valign="top"><%=substitui_caracteres(Trim("" & rs("mensagem")), chr(13), "<br>")%></td>
		<% else %>
		<td class="C MD MB" align="center" valign="top" style="width:50px;color:<%=nivel_acesso_bloco_notas_pedido_cor(rs("nivel_acesso"))%>;"><%=nivel_acesso_bloco_notas_pedido_descricao(rs("nivel_acesso"))%></td>
		<td class="C MD MB" align="left" valign="top"><%=substitui_caracteres(Trim("" & rs("mensagem")), chr(13), "<br>")%></td>
		<% end if %>
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
<!---------- BLOCO DE NOTAS ASSISTENCIA T�CNICA ----------------->
    <% if ID_PARAM_SITE = COD_SITE_ARTVEN_BONSHOP then
    dim cn2, pedido_bs_x_at
    pedido_bs_x_at = ""
    If Not bdd_AT_conecta(cn2) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO) 
    
    %>
    <br id="brBlocoNotasAT" class="notPrint">
    <table id="tableBlocoNotasAT" class="notPrint" width="649" cellspacing="0" cellpadding="1">
    <tr>
	    <td colspan="4" class="ME MD MC MB" align="left"><span class="Rf">BLOCO DE NOTAS (ASSIST�NCIA T�CNICA)</span></td>
    </tr>
    <%       
        s = "SELECT " & _
			    "*" & _
	       " FROM t_PEDIDO_BLOCO_NOTAS" & _
	       " INNER JOIN t_PEDIDO ON (t_PEDIDO.pedido = t_PEDIDO_BLOCO_NOTAS.pedido)" & _
	       " WHERE" & _
			    " (t_PEDIDO.pedido_bs_x_at='" & pedido_selecionado & "')" & _
			    " AND (nivel_acesso <= " & Cstr(nivel_acesso_bloco_notas) & ")" & _
			    " AND (anulado_status = 0)" &_
	       " ORDER BY" & _
			    " t_PEDIDO.data_hora, t_PEDIDO_BLOCO_NOTAS.dt_hr_cadastro"
	    set rs = cn2.execute(s)
	    if rs.Eof then %>
		    <tr class="notVisible">
			    <td colspan="4" class="ME MD MB" align="left">&nbsp;</td>
		    </tr>
    <%		end if
    		
	    do while Not rs.Eof
    %>
        <% 
            dim ultima_os
            s = rs("pedido") 
            if (s <> ultima_os) then %>
        
        <tr>
            <td colspan="4" class="C ME MD MB" align="left"><span class="Rf"> 
            <% 
            Response.Write s
            ultima_os = s
            
             %>
            </span></td>
        </tr>
        <% end if %>
	    <tr>
		    <td class="C ME MD MB" style="width:60px;" align="center" valign="top"><%=formata_data_hora(rs("dt_hr_cadastro"))%></td>
		    <td class="C MD MB" style="width:80px;" align="center" valign="top"><%
			    s = rs("usuario")
			    if Trim("" & rs("loja")) <> "" then s = s & " (Loja&nbsp;" & Trim("" & rs("loja")) & ")"
			    Response.Write s
			    %></td>
		    <% if converte_numero(nivel_acesso_bloco_notas) = converte_numero(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO) then %>
		    <td colspan="2" class="C MD MB" align="left" valign="top"><%=substitui_caracteres(Trim("" & rs("mensagem")), chr(13), "<br>")%></td>
		    <% else %>
		    <td class="C MD MB" align="center" valign="top" style="width:50px;color:<%=nivel_acesso_bloco_notas_pedido_cor(rs("nivel_acesso"))%>;"><%=nivel_acesso_bloco_notas_pedido_descricao(rs("nivel_acesso"))%></td>
		    <td class="C MD MB" align="left" valign="top"><%=substitui_caracteres(Trim("" & rs("mensagem")), chr(13), "<br>")%></td>
		    <% end if %>
	    </tr>
    <%  
		    rs.MoveNext
		    loop
    cn2.Close
	set cn2 = nothing
    %>

	    <tr class="notPrint">
		    <td colspan="4" style="padding:0px;" align="left">
			    <table width="100%" cellpadding="0" cellspacing="0">
			    <tr>
			    <td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="bBlocoNotasAT" id="bBlocoNotasAT" href="javascript:fPEDBlocoNotasAT(fPED)" title="configura as mensagens do bloco de notas para serem impressas ou n�o"><img name="imgPrinterBlocoNotasAT" id="imgPrinterBlocoNotasAT" src="../botao/PrinterError.png" border="0"></a></td>
			    <td align="left">&nbsp;</td>
			    </tr>
			    </table>
		    </td>
	    </tr>

    </table>
    <% end if %>
<% end if %>



<% if s_devolucoes <> "" then %>
<%		if operacao_permitida(OP_LJA_BLOCO_NOTAS_ITEM_DEVOLVIDO_LEITURA, s_lista_operacoes_permitidas) Or _
		   operacao_permitida(OP_LJA_BLOCO_NOTAS_ITEM_DEVOLVIDO_CADASTRAMENTO, s_lista_operacoes_permitidas) then %>
<br id="brBlocoNotasItemDevolvido" class="notPrint">
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
<%		end if %>
<% end if %>



<% if operacao_permitida(OP_LJA_PEDIDO_HISTORICO_PAGAMENTO_EXIBE, s_lista_operacoes_permitidas) then %>
<br id="brHistPagto" class="notPrint">
<table id="tableHistPagto" class="notPrint" width="649" cellspacing="0" cellpadding="1">
<tr>
	<td colspan="8" class="ME MD MC MB" align="left"><span class="Rf">HIST�RICO DE PAGAMENTO</span></td>
</tr>
<%	dtReferenciaLimitePagamentoEmAtraso = obtemDataReferenciaLimitePagamentoEmAtraso
	
	s = "SELECT" & _
			" tFPHP.*," & _
			" Coalesce(tFFC.valor, 0) AS vl_pago_FC," & _
			" tFFC.descricao AS tFFC_descricao"
	if USAR_BRASPAG_CLEARSALE then
		s = s & _
			", t_PAG_PAYMENT.ult_GlobalStatus" & _
			", t_PAG_PAYMENT.captura_confirmada_status"
	else
		s = s & ", '' AS ult_GlobalStatus"
		end if
	s = s & _
	   " FROM t_FIN_PEDIDO_HIST_PAGTO tFPHP" & _
			" LEFT JOIN t_FIN_FLUXO_CAIXA tFFC ON (tFPHP.id_fluxo_caixa = tFFC.id)"

	if USAR_BRASPAG_CLEARSALE then
		s =s & _
			" LEFT JOIN t_PAGTO_GW_PAG_PAYMENT t_PAG_PAYMENT ON ((tFPHP.ctrl_pagto_id_parcela = t_PAG_PAYMENT.id) AND (tFPHP.ctrl_pagto_modulo = " & CTRL_PAGTO_MODULO__BRASPAG_CLEARSALE & "))"
		end if
	
	s = s & _
	   " WHERE" & _
			" (tFPHP.pedido = '" & pedido_selecionado & "')"

'	NA LOJA, SOMENTE OS DADOS DE PAGAMENTO POR CART�O PODEM SER CONSULTADOS
'	EM 22/08/2018, A LILIAN SOLICITOU A LIBERA��O DA EXIBI��O DAS PARCELAS DE BOLETO C/ AUTORIZA��O DO ROG�RIO RASGA
'	if r_pedido.loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
'		s = s & _
'				" AND ((tFPHP.ctrl_pagto_modulo = " & CTRL_PAGTO_MODULO__BRASPAG_CLEARSALE & ") OR (tFPHP.ctrl_pagto_modulo = " & CTRL_PAGTO_MODULO__BRASPAG_WEBHOOK & "))"
'	else
'		s = s & _
'				" AND (tFPHP.ctrl_pagto_modulo = " & CTRL_PAGTO_MODULO__BRASPAG_CLEARSALE & ")"
'		end if

	s = s & _
	   " ORDER BY" & _
			" tFPHP.id"

	set rs = cn.execute(s)
	if rs.Eof then %>
		<tr class="notVisible">
			<td colspan="8" class="ME MD MB" align="left">&nbsp;</td>
		</tr>
<%	else %>
	<tr>
		<td class="Rf ME MD MB" style="width:70px;" align="center" valign="top">Meio Pagto</td>
		<td class="Rf MB" style="width:20px;" align="right" valign="middle"><%=strHistPagtoStatusImg%></td>
		<td class="Rf MD MB" style="width:60px;" align="left" valign="top">Status</td>
		<td class="Rf MD MB" style="width:70px;" align="center" valign="top">DT Vencto</td>
		<td class="Rf MD MB" style="width:80px;" align="right" valign="top">VL Parcela</td>
		<td class="Rf MD MB" style="width:70px;" align="center" valign="top">DT Pagto</td>
		<td class="Rf MD MB" style="width:80px;" align="right" valign="top">VL Pago</td>
		<td class="Rf MD MB" style="padding-left:2px;" align="left" valign="top">Descri��o</td>
	</tr>
<%		end if
		
	do while Not rs.Eof
		strHistPagtoModulo = obtemDescricaoCtrlPagtoModulo(rs("ctrl_pagto_modulo"))
		strHistPagtoDescricao = Trim("" & rs("tFFC_descricao"))
		
		strHistPagtoValorPago = ""
		strHistPagtoStatusDescricao = ""
		strHistPagtoDtVencto = ""
		strHistPagtoVlParcela = ""
		strHistPagtoDtPagto = ""
		strHistPagtoStatusImg = ""
		strHistPagtoCor = "black"
		
		if Trim("" & rs("ctrl_pagto_modulo")) = CTRL_PAGTO_MODULO__BOLETO then
			strHistPagtoDtVencto = formata_data(rs("dt_vencto"))
			strHistPagtoVlParcela = formata_moeda(rs("valor_total"))
			strHistPagtoDtPagto = formata_data(rs("dt_credito"))
		elseif Trim("" & rs("ctrl_pagto_modulo")) = CTRL_PAGTO_MODULO__BRASPAG_WEBHOOK then
			strHistPagtoDtVencto = formata_data(rs("dt_vencto"))
			strHistPagtoVlParcela = formata_moeda(rs("valor_total"))
			strHistPagtoDtPagto = formata_data(rs("dt_credito"))
			if strHistPagtoDescricao = "" then strHistPagtoDescricao = Trim("" & rs("descricao"))
		elseif (Trim("" & rs("ctrl_pagto_modulo")) = CTRL_PAGTO_MODULO__BRASPAG_CARTAO) Or (Trim("" & rs("ctrl_pagto_modulo")) = CTRL_PAGTO_MODULO__BRASPAG_CLEARSALE) then
			strHistPagtoDtPagto = formata_data(rs("dt_operacao"))
			if strHistPagtoDescricao = "" then strHistPagtoDescricao = Trim("" & rs("descricao"))
			end if
		
		if (Trim("" & rs("ctrl_pagto_modulo")) = CTRL_PAGTO_MODULO__BOLETO) And (rs("st_boleto_baixado") = 1) then
			strHistPagtoStatusDescricao = "Baixado"
			strHistPagtoStatusImg = "<img src='../imagem/error_14x14.png' border='0' />"
			strHistPagtoCor = "red"
		elseif (Trim("" & rs("ctrl_pagto_modulo")) = CTRL_PAGTO_MODULO__BRASPAG_CLEARSALE) And (Trim("" & rs("ult_GlobalStatus")) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA) then
			strHistPagtoStatusDescricao = "Processando"
			strHistPagtoStatusImg = "<img src='../imagem/exclamacao_14x14.png' border='0' />"
			strHistPagtoCor = "black"
		elseif Trim("" & rs("status")) = Trim("" & ST_T_FIN_PEDIDO_HIST_PAGTO__QUITADO) Or _
			   ((Trim("" & rs("ctrl_pagto_modulo")) = CTRL_PAGTO_MODULO__BRASPAG_CLEARSALE) And (Trim("" & rs("ult_GlobalStatus")) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA)) then
			strHistPagtoStatusDescricao = "Pago"
			strHistPagtoStatusImg = "<img src='../imagem/ok_14x14.png' border='0' />"
			strHistPagtoCor = "green"
			if Trim("" & rs("ctrl_pagto_modulo")) = CTRL_PAGTO_MODULO__BOLETO then
				strHistPagtoValorPago = formata_moeda(rs("vl_pago_FC"))
			elseif Trim("" & rs("ctrl_pagto_modulo")) = CTRL_PAGTO_MODULO__BRASPAG_WEBHOOK then
				strHistPagtoValorPago = formata_moeda(rs("valor_pago"))
			elseif (Trim("" & rs("ctrl_pagto_modulo")) = CTRL_PAGTO_MODULO__BRASPAG_CARTAO) Or (Trim("" & rs("ctrl_pagto_modulo")) = CTRL_PAGTO_MODULO__BRASPAG_CLEARSALE) then
				strHistPagtoValorPago = formata_moeda(rs("valor_total"))
				end if
		elseif Trim("" & rs("status")) = Trim("" & ST_T_FIN_PEDIDO_HIST_PAGTO__CANCELADO) then
			strHistPagtoStatusDescricao = "Cancelado"
			strHistPagtoCor = "red"
			if (Trim("" & rs("ult_GlobalStatus")) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNADA) Or ((Trim("" & rs("ult_GlobalStatus")) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURA_CANCELADA) And (rs("captura_confirmada_status") = 1)) then
				strHistPagtoValorPago = formata_moeda(-1 * rs("valor_total"))
				end if
			end if
		
	'	PARCELA DE BOLETO EM ATRASO?
		strHistPagtoCorParcelaEmAtraso = "black"
		if Trim("" & rs("ctrl_pagto_modulo")) = Trim("" & CTRL_PAGTO_MODULO__BOLETO) then
			if (Trim("" & rs("status")) = Trim("" & ST_T_FIN_PEDIDO_HIST_PAGTO__PREVISAO)) And (rs("st_boleto_baixado") = 0) And (rs("st_boleto_pago_cheque") = 0) then
				if rs("dt_vencto") <= dtReferenciaLimitePagamentoEmAtraso then
					strHistPagtoCorParcelaEmAtraso = "red"
					end if
				end if
			end if
%>
	<tr>
		<td class="C ME MD MB" style="width:70px;" align="center" valign="middle"><%=strHistPagtoModulo%></td>
		<td class="Rf MB" style="width:20px;vertical-align:middle;" align="right" valign="middle"><%=strHistPagtoStatusImg%></td>
		<td class="C MD MB" style="width:60px;color:<%=strHistPagtoCor%>;" align="left" valign="middle"><%=strHistPagtoStatusDescricao%></td>
		<td class="C MD MB" style="width:70px;color:<%=strHistPagtoCorParcelaEmAtraso%>;" align="center" valign="middle"><%=strHistPagtoDtVencto%></td>
		<td class="C MD MB" style="width:80px;" align="right" valign="middle"><%=strHistPagtoVlParcela%></td>
		<td class="C MD MB" style="width:70px;" align="center" valign="middle"><%=strHistPagtoDtPagto%></td>
		<td class="C MD MB" style="width:80px;" align="right" valign="middle"><%=strHistPagtoValorPago%></td>
		<td class="C MD MB" style="padding-left:2px;" align="left" valign="middle"><%=strHistPagtoDescricao%></td>
	</tr>
<%
		rs.MoveNext
		loop
%>

	<tr class="notPrint">
		<td colspan="8" style="padding:0px;" align="left">
			<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
			<td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="bHistPagtoAlteraImpressao" id="bHistPagtoAlteraImpressao" href="javascript:fPEDHistPagtoAlteraImpressao(fPED)" title="configura o hist�rico de pagamento para ser impresso ou n�o"><img name="imgPrinterHistPagto" id="imgPrinterHistPagto" src="../botao/PrinterError.png" border="0"></a></td>
			<td align="left">&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>

</table>
<% end if %>


<% if operacao_permitida(OP_LJA_PEDIDO_EXIBE_DETALHES_HISTORICO_PAGTO_CARTAO, s_lista_operacoes_permitidas) then %>
<br id="brDetalhesPagtoCartao" class="notPrint">
<table id="tableDetalhesPagtoCartao" class="notPrint" width="649" cellspacing="0" cellpadding="1">
<tr>
	<td colspan="8" class="ME MD MC MB" align="left"><span class="Rf">DETALHES DO HIST�RICO DE PAGAMENTO POR CART�O</span></td>
</tr>
<%	dim strMsgRetorno, strMsgRetornoCor, strPrim_GlobalStatus
    	
	s = "SELECT" & _
	        " pag.data_hora," & _
	        " pag.usuario," & _
	        " payment.bandeira," & _
	        " LEFT(payment.checkout_cartao_numero, 6) + '-****-' + RIGHT(payment.checkout_cartao_numero, 4) AS numero_cartao," & _
	        " payment.checkout_titular_nome," & _
	        " payment.valor_transacao," & _
	        " payment.req_PaymentDataRequest_NumberOfPayments AS num_parcelas," & _
	        " payment.prim_GlobalStatus," & _
	        " payment.resp_PaymentDataResponse_ReturnCode," & _
	        " payment.resp_PaymentDataResponse_ReturnMessage," & _
            " pag.trx_erro_codigo," & _
	        " pag.trx_erro_mensagem" & _
        " FROM t_PAGTO_GW_PAG pag" & _
        " INNER JOIN t_PAGTO_GW_PAG_PAYMENT payment ON (pag.id = payment.id_pagto_gw_pag)" & _
        " WHERE pag.pedido = '" & pedido_selecionado & "'" & _
        " ORDER BY pag.data_hora," & _
	        " payment.id"


	set rs = cn.execute(s)
	if rs.Eof then %>
		<tr class="notVisible">
			<td colspan="8" class="ME MD MB" align="left">&nbsp;</td>
		</tr>
<%	else %>
	<tr>
		<td class="Rf ME MD MB" style="width:55px;" align="center" valign="top">Data/Hora</td>
		<td class="Rf MB MD" style="width:50px;" align="center" valign="middle">Usu�rio</td>
		<td class="Rf MD MB" style="width:50px;" align="center" valign="top">Bandeira</td>
		<td class="Rf MD MB" style="width:85px;" align="center" valign="top">N� Cart�o</td>
		<td class="Rf MD MB" style="width:120px;" align="left" valign="top">Nome do Titular</td>
		<td class="Rf MD MB" style="width:55px;" align="right" valign="top">Valor</td>
		<td class="Rf MD MB" style="width:25px;" align="center" valign="top">Parc</td>
		<td class="Rf MD MB" style="padding-left:2px;" align="left" valign="top">Mensagem de Retorno</td>
	</tr>
<%		end if
		
	do while Not rs.Eof
        strPrim_GlobalStatus = Trim("" & rs("prim_GlobalStatus"))
        
		if CStr(strPrim_GlobalStatus) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA Or _
            CStr(strPrim_GlobalStatus) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA then
            strMsgRetornoCor = "green"
            strMsgRetorno = DecodeUTF8(Trim("" & rs("resp_PaymentDataResponse_ReturnMessage")))
        else
            if Trim("" & rs("resp_PaymentDataResponse_ReturnCode")) <> "" then
                strMsgRetornoCor = "red"
                strMsgRetorno = DecodeUTF8(Trim("" & rs("resp_PaymentDataResponse_ReturnMessage")))
                strMsgRetorno = "(" & rs("resp_PaymentDataResponse_ReturnCode") & ") " & strMsgRetorno
            else
                strMsgRetornoCor = "red"
                strMsgRetorno = Trim("" & rs("trx_erro_mensagem"))
                strMsgRetorno = "(" & rs("trx_erro_codigo") & ") " & strMsgRetorno
            end if
        end if
        
%>
	<tr>
		<td class="C ME MD MB"" align="center" valign="middle"><%=formata_data_hora(rs("data_hora"))%></td>
		<td class="C MB MD" align="center" valign="middle"><%=Trim("" & rs("usuario"))%></td>
		<td class="C MD MB" align="center" valign="middle"><%=iniciais_em_maiusculas(Trim("" & rs("bandeira")))%></td>
		<td class="C MD MB" align="center" valign="middle"><%=Trim("" & rs("numero_cartao"))%></td>
		<td class="C MD MB" align="left" valign="middle"><%=Ucase(Trim("" & rs("checkout_titular_nome")))%></td>
		<td class="C MD MB" align="right" valign="middle"><%=formata_moeda(Trim("" & rs("valor_transacao")))%></td>
		<td class="C MD MB" align="center" valign="middle"><%=Trim("" & rs("num_parcelas"))%></td>
		<td class="C MD MB" style="padding-left:2px; color:<%=strMsgRetornoCor%>" align="left" valign="middle"><%=strMsgRetorno%></td>
	</tr>
<%
		rs.MoveNext
		loop
%>

	<tr class="notPrint">
		<td colspan="8" style="padding:0px;" align="left">
			<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
			<td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="bDetalhesPagtoCartaoAlteraImpressao" id="bDetalhesPagtoCartaoAlteraImpressao" href="javascript:fPEDDetalhesPagtoCartaoAlteraImpressao(fPED)" title="configura o hist�rico de pagamento para ser impresso ou n�o"><img name="imgPrinterDetalhesPagtoCartao" id="imgPrinterDetalhesPagtoCartao" src="../botao/PrinterError.png" border="0"></a></td>
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
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOT�ES   ************ -->
<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr>
	<%	if url_back <> "" then 
			s="resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
		else 
			s="javascript:history.back()"
			end if
	%>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>" title="volta para p�gina anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>

</form>


<!-- ************   DIRECIONA PARA CADASTRO DE CLIENTES   ************ -->
<form method="post" action="clienteedita.asp" id="fCLI" name="fCLI">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=r_pedido.id_cliente%>'>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<input type="hidden" name="edicao_bloqueada" id="edicao_bloqueada" />
<input type="hidden" name="pagina_retorno" id="pagina_retorno" value='pedido.asp?pedido_selecionado=<%=pedido_selecionado%>&url_back=X'>
</form>


</center>
<div id="divClienteConsultaView"><center><div id="divInternoClienteConsultaView"><img id="imgFechaDivClienteConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframeClienteConsultaView"></iframe></div></center></div>
<div id="divRastreioConsultaView"><center><div id="divInternoRastreioConsultaView"><img id="imgFechaDivRastreioConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframeRastreioConsultaView"></iframe></div></center></div>
<div id="divOrcamentistaEIndicadorConsultaView"><center><div id="divInternoOrcamentistaEIndicadorConsultaView"><img id="imgFechaDivOrcamentistaEIndicadorConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframeOrcamentistaEIndicadorConsultaView"></iframe></div></center></div>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>