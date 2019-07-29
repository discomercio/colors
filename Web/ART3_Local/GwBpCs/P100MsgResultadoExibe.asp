<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Braspag.asp"    -->
<!-- #include file = "../global/BraspagCS.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  P100MsgResultadoExibe.asp
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

'	OBS: A PÁGINA QUE EXIBE A MENSAGEM SOBRE O RESULTADO DA TRANSAÇÃO É ACIONADA
'	~~~~ ATRAVÉS DOS SEGUINTES PASSOS:
'	1) A PÁGINA QUE EXECUTA A TRANSAÇÃO C/ A BRASPAG VIA WEB SERVICE ARMAZENA OS DADOS
'		RECEBIDOS NO BD E, EM SEGUIDA, ENCAMINHA P/ ESTA PÁGINA INTERMEDIÁRIA INFORMANDO
'		O ID DO REGISTRO.
'	2) A PÁGINA INTERMEDIÁRIA PREPARA OS DADOS EM CAMPOS HIDDEN DE UM FORM, LÊ E APAGA OS
'		DADOS ARMAZENADOS ATRAVÉS DA SESSION E, POR FIM, FAZ UM SUBMIT() P/ A PÁGINA
'		FINAL DE EXIBIÇÃO.
'	3) COM ESTE MECANISMO, SE O USUÁRIO ACIONAR O REFRESH NA PÁGINA DE EXIBIÇÃO, EVITAM-SE
'		OS SEGUINTES PROBLEMAS:
'		A) REEXECUTAR O PROCESSAMENTO DA TRANSAÇÃO.
'		B) PARA OS DADOS ARMAZENADOS NA SESSION, A PARTIR DA 2ª EXECUÇÃO OS DADOS JÁ TERIAM
'			SIDO APAGADOS.

	On Error GoTo 0
	Err.Clear

	dim s, usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))

	dim pedido_selecionado, id_pedido_base
	pedido_selecionado = Trim(Request("pedido_selecionado"))
	id_pedido_base = retorna_num_pedido_base(pedido_selecionado)

	dim strIdPagtoGwPag
	strIdPagtoGwPag = Trim(Request("idPagtoGwPag"))

	dim alerta
	alerta = Trim(Request("alerta"))

	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, t_PAG, t_PAG_PAYMENT, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(t_PAG, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(t_PAG_PAYMENT, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim strMsgErroRecordsetNotFound
	dim s_cor
	dim blnTransacaoOk
	dim intQtdeTransacaoOk
	dim intQtdeNaoAutorizada
	dim intQtdeCartoes
	intQtdeTransacaoOk = 0
	intQtdeNaoAutorizada = 0
	intQtdeCartoes = 0
	
	dim i
	dim r_pedido, v_item, r_cliente
	dim m_vl_total, m_total_geral
	
	if alerta = "" then
		if Not le_pedido(id_pedido_base, r_pedido, msg_erro) then
			alerta = msg_erro
		else
			if Not le_pedido_item_consolidado_familia(id_pedido_base, v_item, msg_erro) then alerta = msg_erro
			end if
		
		if alerta = "" then
			set r_cliente = New cl_CLIENTE
			if Not x_cliente_bd(r_pedido.id_cliente, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_CADASTRADO)
			end if
		end if
	
	dim s_endereco
	s_endereco = ""
	if alerta = "" then
		s_endereco = iniciais_em_maiusculas(r_cliente.endereco)
		if Trim(r_cliente.endereco_numero) <> "" then s_endereco = s_endereco & ", " & Trim(r_cliente.endereco_numero)
		if Trim(r_cliente.endereco_complemento) <> "" then s_endereco = s_endereco & " " & Trim(r_cliente.endereco_complemento)
		end if
	
	dim strMsgResultado, strRecibo, strReciboTela, strReciboBd
	dim s_class_box, s_id, s_ordem
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
	<title>LOJA</title>
	</head>


<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>



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
<link href="<%=URL_FILE__BRASPAG_CARTAO_RECIBO_CSS%>" rel="stylesheet" type="text/css">



<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();" link=#000000 alink=#000000 vlink=#000000>
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" href="<%="../Loja/pedido.asp?pedido_selecionado=" & pedido_selecionado & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>



<% else %>
<!-- ****************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR DADOS DO RETORNO  ***************** -->
<!-- ****************************************************************** -->
<body>
<center>

<%		strMsgErroRecordsetNotFound = ""
		s = "SELECT * FROM t_PAGTO_GW_PAG WHERE (id = " & strIdPagtoGwPag & ")"
		if t_PAG.State <> 0 then t_PAG.Close
		t_PAG.Open s, cn
		if t_PAG.Eof then
			strMsgErroRecordsetNotFound = "Falha ao tentar localizar o registro principal da transação!"
			end if
		
		if strMsgErroRecordsetNotFound = "" then
			s = "SELECT * FROM t_PAGTO_GW_PAG_PAYMENT WHERE (id_pagto_gw_pag = " & strIdPagtoGwPag & ") ORDER BY ordem"
			if t_PAG_PAYMENT.State <> 0 then t_PAG_PAYMENT.Close
			t_PAG_PAYMENT.Open s, cn
			if t_PAG_PAYMENT.Eof then
				strMsgErroRecordsetNotFound = "Falha ao tentar localizar o registro dos dados da consulta de autorização!"
				end if
			end if
		
		'	NÃO ENCONTROU O REGISTRO: EXIBE MENSAGEM DE ERRO!!
		if strMsgErroRecordsetNotFound <> "" then
%>
	<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><span style='margin:5px 2px 5px 2px;'><%=strMsgErroRecordsetNotFound%></span></div>
<%
		else ' if (strMsgErroRecordsetNotFound <> "")
%>

	<table cellspacing="0" width="649" style="border-bottom:1px solid black">
	<tr>
		<td align="center" valign="bottom"><img src="../imagem/<%=BRASPAG_LOGOTIPO_LOJA%>"></td>
	</tr>
	</table>

	<br />
	<br />
	
<%
	strRecibo = _
			"<table id='tabRecibo' border='0' cellpadding='2' cellspacing='1' width='649'>" & chr(13) & _
			"	<tr id='trTitRec'>" & chr(13) & _
			"		<td class='tdTitMain prnBorderAll' colspan='2'>Transação</td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"	<tr id='trPedido'>" & chr(13) & _
			"		<td class='tdTitR'>Nº Pedido:&nbsp;</td>" & chr(13) & _
			"		<td class='tdDadosL'>" & id_pedido_base & "</td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"	<tr id='trDtOp'>" & chr(13) & _
			"		<td class='tdTitR'>Data:&nbsp;</td>" & chr(13) & _
			"		<td class='tdDadosL'>" & formata_data(t_PAG("trx_RX_data")) & "</td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"	<tr id='trHrOp'>" & chr(13) & _
			"		<td class='tdTitR'>Hora:&nbsp;</td>" & chr(13) & _
			"		<td class='tdDadosL'>" & formata_hora_hhmm(t_PAG("trx_RX_data_hora")) & "</td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"	<tr id='trSpTabTr'>" & chr(13) & _
			"		<td colspan='2' style='height:15px'></td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"	<tr id='trTitRec'>" & chr(13) & _
			"		<td class='tdTitMain prnBorderAll' colspan='2'>Resultado da Transação</td>" & chr(13) & _
			"	</tr>" & chr(13)
%>
	<%
	do while Not t_PAG_PAYMENT.Eof
		intQtdeCartoes = intQtdeCartoes + 1
		s_id = Trim("" & t_PAG_PAYMENT("id"))
		s_ordem = Trim("" & t_PAG_PAYMENT("ordem"))
		s_class_box = "trBoxTrx trBoxTrxOrd_" & s_ordem & " trBoxTrxId_" & s_id

		if (Trim("" & t_PAG_PAYMENT("resp_PaymentDataResponse_Status")) = BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__CAPTURADA) then
			blnTransacaoOk = True
			intQtdeTransacaoOk = intQtdeTransacaoOk + 1
			s_cor = "Green"
			strMsgResultado = "Pagamento Autorizado (em processamento)"
		elseif (Trim("" & t_PAG_PAYMENT("resp_PaymentDataResponse_Status")) = BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__AUTORIZADA) then
			blnTransacaoOk = True
			intQtdeTransacaoOk = intQtdeTransacaoOk + 1
			s_cor = "Green"
			strMsgResultado = "Pagamento Autorizado (em processamento)"
		else
			blnTransacaoOk = False
			intQtdeNaoAutorizada = intQtdeNaoAutorizada + 1
			s_cor = "Red"
			strMsgResultado = "Pagamento Não Autorizado"
			end if
	%>
	
<%
		strRecibo = strRecibo & _
				"	<tr id='trBand_" & s_ordem & "' class='" & s_class_box & " trBand trBandOrd_" & s_ordem & " trBandId_" & s_id & "'>" & chr(13) & _
				"		<td class='tdTitR'>Bandeira:&nbsp;</td>" & chr(13) & _
				"		<td class='tdDadosL'>" & BraspagDescricaoBandeira(Trim("" & t_PAG_PAYMENT("bandeira"))) & "</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr id='trCard_" & s_ordem & "' class='" & s_class_box & " trCard trCardOrd_" & s_ordem & " trCardId_" & s_id & "'>" & chr(13) & _
				"		<td class='tdTitR'>Cartão:&nbsp;</td>" & chr(13) & _
				"		<td class='tdDadosL'>" & _
							"<span class='tdDadosL spnCard spnCardOrd_" & s_ordem & " spnCardId_" & s_id & "' id='spnCard_" & s_id & "'>" & BraspagCSProtegeNumeroCartao(t_PAG_PAYMENT("checkout_cartao_numero")) & "</span>" & _
						"</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr id='trStTr_" & s_ordem & "' class='" & s_class_box & " trStTr trStTrOrd_" & s_ordem & " trStTrId_" & s_id & "'>" & chr(13) & _
				"		<td class='tdTitR'>Status da Transação:&nbsp;</td>" & chr(13) & _
				"		<td class='tdDadosL'>" & _
							"<span class='tdDadosL spnStTrx spnStTrxOrd_" & s_ordem & " spnStTrxId_" & s_id & "' id='spnStTrx_" & s_id & "' style='color:" & s_cor & ";'>" & strMsgResultado & "</span>" & _
						"</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr id='trCodAut_" & s_ordem & "' class='" & s_class_box & " trCodAut trCodAutOrd_" & s_ordem & " trCodAutId_" & s_id & "'>" & chr(13) & _
				"		<td class='tdTitR'>Código de Autorização:&nbsp;</td>" & chr(13) & _
				"		<td class='tdDadosL'>" & _
							"<span class='tdDadosL spnCodAutTrx spnCodAutTrxOrd_" & s_ordem & " spnCodAutTrxId_" & s_id & "' id='spnCodAutTrx_" & s_id & "'>" & Trim("" & t_PAG_PAYMENT("resp_PaymentDataResponse_AuthorizationCode")) & "</span>" & _
						"</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr id='trComprov_" & s_ordem & "' class='" & s_class_box & " trComprov trComprovOrd_" & s_ordem & " trComprovId_" & s_id & "'>" & chr(13) & _
				"		<td class='tdTitR'>Comprovante de Venda:&nbsp;</td>" & chr(13) & _
				"		<td class='tdDadosL'>" & _
							"<span class='tdDadosL spnComprovTrx spnComprovTrxOrd_" & s_ordem & " spnComprovTrxId_" & s_id & "' id='spnComprovTrx_" & s_id & "'>" & Trim("" & t_PAG_PAYMENT("resp_PaymentDataResponse_ProofOfSale")) & "</span>" & _
						"</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr id='trVlTr_" & s_ordem & "' class='" & s_class_box & " trVlTr trVlTrOrd_" & s_ordem & " trVlTrId_" & s_id & "'>" & chr(13) & _
				"		<td class='tdTitR'>Valor:&nbsp;</td>" & chr(13) & _
				"		<td class='tdDadosL'>" & _
							"<span class='tdDadosL spnVlTrx spnVlTrxOrd_" & s_ordem & " spnVlTrxId_" & s_id & "' id='spnVlTrx_" & s_id & "'>" & formata_moeda(t_PAG_PAYMENT("valor_transacao")) & "</span>" & _
						"</td>" & chr(13) & _
				"	</tr>" & chr(13)
%>
	
<%
		if blnTransacaoOk then
			strRecibo = strRecibo & _
				"	<tr id='trOpPg_" & s_ordem & "' class='" & s_class_box & " trOpPg trOpPgOrd_" & s_ordem & " trOpPgId_" & s_id & "'>" & chr(13) & _
				"		<td class='tdTitR'>Opção de Pagamento:&nbsp;</td>" & chr(13) & _
				"		<td class='tdDadosL'>" & _
							"<span class='tdDadosL spnOpPgTrx spnOpPgTrxOrd_" & s_ordem & " spnOpPgTrxId_" & s_id & "' id='spnOpPgTrx_" & s_id & "'>" & BraspagDescricaoParcelamento(Trim("" & t_PAG_PAYMENT("req_PaymentDataRequest_PaymentPlan")), Trim("" & t_PAG_PAYMENT("req_PaymentDataRequest_NumberOfPayments")), t_PAG_PAYMENT("valor_transacao")) & "</span>" & _
						"</td>" & chr(13) & _
				"	</tr>" & chr(13)
			end if

		strRecibo = strRecibo & _
			"	<tr id='trSpTabResTr_" & s_ordem & "' class='" & s_class_box & " trSpTabRes trSpTabResOrd_" & s_ordem & " trSpTabResId_" & s_id & "'>" & chr(13) & _
			"		<td colspan='2' style='height:15px'></td>" & chr(13) & _
			"	</tr>" & chr(13)

		t_PAG_PAYMENT.MoveNext
		loop
%>

<%
	strRecibo = strRecibo & _
			"	<tr id='trTitCli'>" & chr(13) & _
			"		<td colspan='2' bgcolor='#eeeeee' class='tdTitC prnBorderAll'><b>Dados do Cliente</b></td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"	<tr id='trNomeCli'>" & chr(13) & _
			"		<td class='tdTitR'>Cliente:&nbsp;</td>" & chr(13) & _
			"		<td class='tdDadosL'>" & iniciais_em_maiusculas(r_cliente.nome) & "</td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"	<tr id='trEmailCli'>" & chr(13) & _
			"		<td class='tdTitR'>E-Mail:&nbsp;</td>" & chr(13) & _
			"		<td class='tdDadosL'>" & LCase(r_cliente.email) & "</td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"	<tr id='trEndCli'>" & chr(13) & _
			"		<td class='tdTitR' valign='top'>Endereço:&nbsp;</td>" & chr(13) & _
			"		<td class='tdDadosL'>" & _
						s_endereco & "<br />" & _
						iniciais_em_maiusculas(r_cliente.bairro) & "<br />" & _
						iniciais_em_maiusculas(r_cliente.cidade) & " - " & UCase(r_cliente.uf) & "<br />" & _
						cep_formata(r_cliente.cep) & _
					"</td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"	<tr id='trSpTabProd'>" & chr(13) & _
			"		<td colspan='2' style='height:15px'></td>" & chr(13) & _
			"	</tr>" & chr(13)
%>

<%
	strRecibo = strRecibo & _
			"	<tr id='trTabProd'>" & chr(13) & _
			"		<td colspan='2' align='center'>" & chr(13) & _
			"			<table id='tabProd' width='100%' border='0' cellspacing='0' cellpadding='3'>" & chr(13) & _
			"				<tr bgcolor='#eeeeee'>" & chr(13) & _
			"					<td width='350' class='tdTitL prnBorderLTB'><b>Item</b></td>" & chr(13) & _
			"					<td width='100' class='tdTitR prnBorderTB'><b>Preço</b></td>" & chr(13) & _
			"					<td width='80' class='tdTitR prnBorderTB'><b>Qtde</b></td>" & chr(13) & _
			"					<td width='90' class='tdTitR prnBorderRTB'><b>Total</b></td>" & chr(13) & _
			"				</tr>" & chr(13)
%>

<%
	for i=Lbound(v_item) to Ubound(v_item)
		if Trim(v_item(i).produto) <> "" then
			with v_item(i)
				m_vl_total = .qtde * .preco_NF
				m_total_geral = m_total_geral + m_vl_total
				
				strRecibo = strRecibo & _
			"				<tr>" & chr(13) & _
			"					<td class='tdDadosL'>" & Trim(.descricao) & "</td>" & chr(13) & _
			"					<td class='tdDadosR'>" & formata_moeda(.preco_NF) & "</td>" & chr(13) & _
			"					<td class='tdDadosR'>" & Cstr(.qtde) & "</td>" & chr(13) & _
			"					<td class='tdDadosR'>" & formata_moeda(m_vl_total) & "</td>" & chr(13) & _
			"				</tr>" & chr(13)
				end with
			end if
		next
%>

<%
	strRecibo = strRecibo & _
			"				<tr>" & chr(13) & _
			"					<td colspan='4' style='height:5px'></td></tr>" & chr(13) & _
			"				<tr id='trVlTot'>" & chr(13) & _
			"					<td colspan='3' class='tdDadosR'>" & _
									"<b>Subtotal: " & SIMBOLO_MONETARIO & "</b>" & _
								"</td>" & chr(13) & _
			"					<td class='tdDadosR'>" & _
									"<b>" & formata_moeda(m_total_geral) & "</b>" & _
								"</td>" & chr(13) & _
			"				</tr>" & chr(13) & _
			"				<tr id='trVlFrete'>" & chr(13) & _
			"					<td colspan='3' height='2' class='tdDadosR'>" & _
									"<b>Frete: " & SIMBOLO_MONETARIO & "</b>" & _
								"</td>" & chr(13) & _
			"					<td class='tdDadosR'>" & _
									"<b>" & formata_moeda(r_pedido.vl_frete) & "</b>" & _
								"</td>" & chr(13) & _
			"				</tr>" & chr(13) & _
			"				<tr id='trVlTotGeral'>" & chr(13) & _
			"					<td colspan='3' height='2' class='tdDadosR'>" & _
									"<b>Total: " & SIMBOLO_MONETARIO & "</b>" & _
								"</td>" & chr(13) & _
			"					<td class='tdDadosR'>" & _
									"<b>" & formata_moeda(m_total_geral+r_pedido.vl_frete) & "</b>" & _
								"</td>" & chr(13) & _
			"				</tr>" & chr(13) & _
			"			</table>" & chr(13) & _
			"		</td>" & chr(13) & _
			"	</tr>" & chr(13)
%>

<%
	strReciboTela = strRecibo
	strReciboBd = strRecibo
%>

<%	if intQtdeTransacaoOk > 0 then
		strReciboTela = strReciboTela & _
			"	<tr>" & chr(13) & _
			"		<td colspan='2' style='height:15px'></td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"	<tr id='trBtnPrn' class='notPrint'>" & chr(13) & _
			"		<td colspan='2' align='center'>" & chr(13) & _
			"			<table width='500'>" & chr(13) & _
			"				<tr>" & chr(13) & _
			"					<td class='tdDadosC'>" & _
									"<center>" & _
									"Recomendamos imprimir este recibo.&nbsp;&nbsp;<a href='javascript:window.print();'><img src='../botao/Printer.png' border='0' /></a>" & _
									"<br />" & _
									"</center>" & _
								"</td>" & chr(13) & _
			"				</tr>" & chr(13) & _
			"			</table>" & chr(13) & _
			"		</td>" & chr(13) & _
			"	</tr>" & chr(13)
		end if
%>

<%	if intQtdeNaoAutorizada > 0 then
		strReciboTela = strReciboTela & _
			"	<tr>" & chr(13) & _
			"		<td colspan='2' style='height:15px'></td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"	<tr id='trAvisoNOK'>" & chr(13) & _
			"		<td colspan='2' align='center'>" & chr(13) & _
			"			<table width='550'>" & chr(13) & _
			"				<tr>" & chr(13) & _
			"					<td class='tdDadosL' style='color:red;'>" & _
									"ATENÇÃO!<br />"
		
		if intQtdeNaoAutorizada = 1 then
			if intQtdeCartoes = 1 then
				strReciboTela = strReciboTela & _
					"A transação não foi autorizada."
			else
				strReciboTela = strReciboTela & _
					"1 transação não foi autorizada."
				end if
		else
			strReciboTela = strReciboTela & _
				Cstr(intQtdeNaoAutorizada) & " transações não foram autorizadas."
			end if

		strReciboTela = strReciboTela & _
									"<br />Por favor, refaça o pagamento que falhou preenchendo atentamente os dados solicitados ou tente utilizar outro cartão." & _
									"<br />" & _
								"</td>" & chr(13) & _
			"				</tr>" & chr(13) & _
			"			</table>" & chr(13) & _
			"		</td>" & chr(13) & _
			"	</tr>" & chr(13)
		end if
%>

<%
	strReciboTela = strReciboTela & _
			"</table>" & chr(13)
	
	strReciboBd = strReciboBd & _
			"</table>" & chr(13)
	
	Response.Write strReciboTela
	
	if Not t_PAG.Eof then
		t_PAG("recibo_html") = strReciboBd
		t_PAG("recibo_url_css") = URL_FILE__BRASPAG_CARTAO_RECIBO_CSS
		t_PAG.Update
		end if
%>

<%
			end if ' if (strMsgErroRecordsetNotFound <> "")
%>

<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<br />

<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" href="<%="../Loja/pedido.asp?pedido_selecionado=" & pedido_selecionado & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>

</center>
</body>
<% end if %>

</html>


<%

'	FECHA CONEXAO COM O BANCO DE DADOS
	if t_PAG.State <> 0 then t_PAG.Close
	set t_PAG=nothing

	if t_PAG_PAYMENT.State <> 0 then t_PAG_PAYMENT.Close
	set t_PAG_PAYMENT=nothing

	cn.Close
	set cn = nothing
	
%>