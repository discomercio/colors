<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Braspag.asp"    -->
<!-- #include file = "../global/BraspagCS.asp"    -->

<%
'     =============================================================
'	  AjaxBraspagClearsalePagCancelOuEstorno.asp
'     =============================================================
'
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


' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_BRASPAG_EM_SEG
	
	dim usuario, id_pagto_gw_pag, id_pagto_gw_pag_payment, msg_erro
	usuario = Trim(Request("usuario"))
	id_pagto_gw_pag = Trim(Request("id_pagto_gw_pag"))
	id_pagto_gw_pag_payment = Trim(Request("id_pagto_gw_pag_payment"))
	
	if id_pagto_gw_pag = "" then
		Response.Write "Identificador não informado!!"
		Response.End
		end if
	
	if id_pagto_gw_pag_payment = "" then
		Response.Write "Identificador da operação com o Pagador não informado!!"
		Response.End
		end if
	
	if converte_numero(id_pagto_gw_pag) = 0 then
		Response.Write "Identificador informado é inválido (" & id_pagto_gw_pag & ")!!"
		Response.End
		end if
	
	if converte_numero(id_pagto_gw_pag_payment) = 0 then
		Response.Write "Identificador da operação com o Pagador é inválido (" & id_pagto_gw_pag_payment & ")!!"
		Response.End
		end if
	
	dim s
	dim cn, rs
	dim strResp
'	CONECTA COM O BANCO DE DADOS
	If Not bdd_conecta(cn) then
		msg_erro = erro_descricao(ERR_CONEXAO)
		Response.Write msg_erro
		Response.End
		end if
	
	dim blnEstornoPendente
	blnEstornoPendente = False
	s = "SELECT" & _
			" refund_pending_status," & _
			" refund_pending_confirmado_status," & _
			" refund_pending_falha_status" & _
		" FROM t_PAGTO_GW_PAG_PAYMENT" & _
		" WHERE" & _
			" (id = " & id_pagto_gw_pag_payment & ")"
	set rs = cn.Execute(s)
	if Not rs.Eof then
		if (rs("refund_pending_status")=1) And (rs("refund_pending_confirmado_status")=0) And (rs("refund_pending_falha_status")=0) then blnEstornoPendente = True
		rs.Close
		set rs = nothing
		end if

	if blnEstornoPendente then
		Response.Write "A transação encontra-se com status '" & BraspagPagadorDescricaoGlobalStatus(BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNO_PENDENTE) & "', portanto, não é possível fazer uma nova requisição de estorno!!"
		Response.End
		end if

'	EXECUTA A CONSULTA P/ ATUALIZAR O STATUS E VERIFICAR SE A TRANSAÇÃO ESTÁ CAPTURADA E QUAL É A DATA DA CAPTURA
	dim trx_Get_1, r_rx_Get_1
	call BraspagClearsaleProcessaConsulta_GetTransactionData(id_pagto_gw_pag, id_pagto_gw_pag_payment, usuario, trx_Get_1, r_rx_Get_1, msg_erro)
	if msg_erro <> "" then
		Response.Write msg_erro
		Response.End
		end if
	
	dim stGlobalStatus_1
	stGlobalStatus_1 = decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(r_rx_Get_1.PAG_Status)
	if (Trim("" & stGlobalStatus_1) <> Trim("" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA)) And (Trim("" & stGlobalStatus_1) <> Trim("" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA)) then
		Response.Write "Não é possível realizar o cancelamento/estorno porque a transação não está com status 'Autorizada' ou 'Capturada'"
		Response.End
		end if
	
	if Trim("" & stGlobalStatus_1) = Trim("" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA) then
		if Trim("" & r_rx_Get_1.PAG_CapturedDate) = "" then
			Response.Write "Não é possível realizar o cancelamento/estorno porque a data de captura da transação não foi informada pela Braspag"
			Response.End
			end if
		end if
	
	dim stGlobalStatusAux, stDescricaoGlobalStatusAux
	dim blnRequestVoid, blnRequestRefund, stGlobalStatus_2, stSucessoOperacao
	dim dtCapturedDate
	
	if Trim("" & stGlobalStatus_1) = Trim("" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA) then
		blnRequestVoid = True
		blnRequestRefund = False
	else
		dtCapturedDate = converte_para_datetime_from_mmddyyyy_hhmmss_am_pm(r_rx_Get_1.PAG_CapturedDate)
	'	O CANCELAMENTO (VOID) PODE SER FEITO ATÉ AS 23:59:59 DO MESMO DIA DA CAPTURA, APÓS ISSO DEVE SER FEITO UM ESTORNO (REFUND)
		if formata_data_yyyymmdd(Date) = formata_data_yyyymmdd(dtCapturedDate) then
			blnRequestVoid = True
			blnRequestRefund = False
		else
			blnRequestVoid = False
			blnRequestRefund = True
			end if
		end if
	
	if blnRequestVoid then
	'	CANCELAMENTO (VOID)
		dim trx_Void, r_rx_Void
		call BraspagClearsaleProcessaRequisicao_VoidCreditCardTransaction(id_pagto_gw_pag, id_pagto_gw_pag_payment, usuario, trx_Void, r_rx_Void, msg_erro)
		if (r_rx_Void.PAG_ReturnCode = "090") And Instr(r_rx_Void.PAG_ReturnMessage, "90-ESTORNO SOMENTE P/TRAN DO DIA") <> 0 then
			'TRATAMENTO P/ O CASO EM QUE A GETNET NÃO ACEITA A REQUISIÇÃO VOID DE TRANSAÇÃO AUTORIZADA EM DATA ANTERIOR E CAPTURADA HOJE
			blnRequestRefund = True
		elseif msg_erro <> "" then
			Response.Write msg_erro
			Response.End
			end if
		end if

	if blnRequestRefund then
	'	ESTORNO (REFUND)
		dim trx_Refund, r_rx_Refund
		call BraspagClearsaleProcessaRequisicao_RefundCreditCardTransaction(id_pagto_gw_pag, id_pagto_gw_pag_payment, usuario, trx_Refund, r_rx_Refund, msg_erro)
		if msg_erro <> "" then
			Response.Write msg_erro
			Response.End
		else
			if r_rx_Refund.PAG_Status = BRASPAG_PAGADOR_CARTAO_REFUNDCREDITCARDTRANSACTIONRESPONSE_STATUS__REFUND_ACCEPTED then
				blnEstornoPendente = True
				end if
			end if
		end if
	
'	REALIZA NOVA CONSULTA P/ OBTER O STATUS ATUALIZADO
	dim trx_Get_2, r_rx_Get_2
	call BraspagClearsaleProcessaConsulta_GetTransactionData(id_pagto_gw_pag, id_pagto_gw_pag_payment, usuario, trx_Get_2, r_rx_Get_2, msg_erro)
	if msg_erro <> "" then
		Response.Write msg_erro
		Response.End
		end if
	
	stGlobalStatus_2 = decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(r_rx_Get_2.PAG_Status)
	
	if blnRequestRefund then
		if (r_rx_Refund.PAG_Status = BRASPAG_PAGADOR_CARTAO_REFUNDCREDITCARDTRANSACTIONRESPONSE_STATUS__REFUND_CONFIRMED) Or blnEstornoPendente then
			stSucessoOperacao = "1"
		else
			stSucessoOperacao = "0"
			end if
		
		if blnEstornoPendente then
			stGlobalStatusAux = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNO_PENDENTE
			stDescricaoGlobalStatusAux = BraspagPagadorDescricaoGlobalStatus(BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNO_PENDENTE)
		else
			stGlobalStatusAux = stGlobalStatus_2
			stDescricaoGlobalStatusAux = BraspagPagadorDescricaoGlobalStatus(stGlobalStatus_2)
			end if

	'	MONTA RESPOSTA
		strResp = "{ " & _
					"""id_pagto_gw_pag"" : """ & id_pagto_gw_pag & """," & _
					"""id_pagto_gw_pag_payment"" : """ & id_pagto_gw_pag_payment & """," & _
					"""PAG_TipoOperacao"" : """ & "REFUND" & """," & _
					"""PAG_SucessoOperacao"" : """ & stSucessoOperacao & """," & _
					"""PAG_CorrelationId"" : """ & r_rx_Refund.PAG_CorrelationId & """," & _
					"""PAG_BraspagTransactionId"" : """ & r_rx_Refund.PAG_BraspagTransactionId & """," & _
					"""PAG_AcquirerTransactionId"" : """ & r_rx_Refund.PAG_AcquirerTransactionId & """," & _
					"""PAG_Amount"" : """ & r_rx_Refund.PAG_Amount & """," & _
					"""PAG_AuthorizationCode"" : """ & r_rx_Refund.PAG_AuthorizationCode & """," & _
					"""PAG_ReturnCode"" : """ & r_rx_Refund.PAG_ReturnCode & """," & _
					"""PAG_ReturnMessage"" : """ & r_rx_Refund.PAG_ReturnMessage & """," & _
					"""PAG_GlobalStatus"" : """ & stGlobalStatusAux & """," & _
					"""PAG_DescricaoGlobalStatus"" : """ & stDescricaoGlobalStatusAux & """," & _
					"""PAG_GlobalStatus_atualizacao_data_hora"" : """ & formata_data_hora(Now) & """," & _
					"""PAG_ProofOfSale"" : """ & r_rx_Refund.PAG_ProofOfSale & """," & _
					"""PAG_ServiceTaxAmount"" : """ & r_rx_Refund.PAG_ServiceTaxAmount & """," & _
					"""PAG_ErrorCode"" : """ & r_rx_Refund.PAG_ErrorCode & """," & _
					"""PAG_ErrorMessage"" : """ & r_rx_Refund.PAG_ErrorMessage & """," & _
					"""PAG_faultcode"" : """ & r_rx_Refund.PAG_faultcode & """," & _
					"""PAG_faultstring"" : """ & r_rx_Refund.PAG_faultstring & """" & _
				" }"
	elseif blnRequestVoid then
		if r_rx_Void.PAG_Status = BRASPAG_PAGADOR_CARTAO_VOIDCREDITCARDTRANSACTIONRESPONSE_STATUS__VOID_CONFIRMED then
			stSucessoOperacao = "1"
		else
			stSucessoOperacao = "0"
			end if
		
	'	MONTA RESPOSTA
		strResp = "{ " & _
					"""id_pagto_gw_pag"" : """ & id_pagto_gw_pag & """," & _
					"""id_pagto_gw_pag_payment"" : """ & id_pagto_gw_pag_payment & """," & _
					"""PAG_TipoOperacao"" : """ & "VOID" & """," & _
					"""PAG_SucessoOperacao"" : """ & stSucessoOperacao & """," & _
					"""PAG_CorrelationId"" : """ & r_rx_Void.PAG_CorrelationId & """," & _
					"""PAG_BraspagTransactionId"" : """ & r_rx_Void.PAG_BraspagTransactionId & """," & _
					"""PAG_AcquirerTransactionId"" : """ & r_rx_Void.PAG_AcquirerTransactionId & """," & _
					"""PAG_Amount"" : """ & r_rx_Void.PAG_Amount & """," & _
					"""PAG_AuthorizationCode"" : """ & r_rx_Void.PAG_AuthorizationCode & """," & _
					"""PAG_ReturnCode"" : """ & r_rx_Void.PAG_ReturnCode & """," & _
					"""PAG_ReturnMessage"" : """ & r_rx_Void.PAG_ReturnMessage & """," & _
					"""PAG_GlobalStatus"" : """ & stGlobalStatus_2 & """," & _
					"""PAG_DescricaoGlobalStatus"" : """ & BraspagPagadorDescricaoGlobalStatus(stGlobalStatus_2) & """," & _
					"""PAG_GlobalStatus_atualizacao_data_hora"" : """ & formata_data_hora(Now) & """," & _
					"""PAG_ProofOfSale"" : """ & r_rx_Void.PAG_ProofOfSale & """," & _
					"""PAG_ServiceTaxAmount"" : """ & r_rx_Void.PAG_ServiceTaxAmount & """," & _
					"""PAG_ErrorCode"" : """ & r_rx_Void.PAG_ErrorCode & """," & _
					"""PAG_ErrorMessage"" : """ & r_rx_Void.PAG_ErrorMessage & """," & _
					"""PAG_faultcode"" : """ & r_rx_Void.PAG_faultcode & """," & _
					"""PAG_faultstring"" : """ & r_rx_Void.PAG_faultstring & """" & _
				" }"
	else
		Response.Write "Erro desconhecido: não foi processada a requisição Void e nem a Refund!"
		end if
	
	cn.Close
	set cn = nothing
	
'	ENVIA RESPOSTA
	Response.Write strResp
	Response.End
%>
