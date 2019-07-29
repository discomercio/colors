<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Braspag.asp"    -->

<%
'     =============================================================
'	  AjaxBraspagAfConsultaStatus.asp
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
	
	dim usuario, id_pedido_pagto_braspag, id_pedido_pagto_braspag_af, id_pedido_pagto_braspag_pag, msg_erro
	usuario = Trim(Request("usuario"))
	id_pedido_pagto_braspag = Trim(Request("id_pedido_pagto_braspag"))
	id_pedido_pagto_braspag_af = Trim(Request("id_pedido_pagto_braspag_af"))
	id_pedido_pagto_braspag_pag = Trim(Request("id_pedido_pagto_braspag_pag"))
	
	if id_pedido_pagto_braspag = "" then
		Response.Write "Identificador não informado!!"
		Response.End
		end if
	
	if id_pedido_pagto_braspag_af = "" then
		Response.Write "Identificador da operação com o Antifraude não informado!!"
		Response.End
		end if
	
	if id_pedido_pagto_braspag_pag = "" then
		Response.Write "Identificador da operação com o Pagador não informado!!"
		Response.End
		end if
	
	if converte_numero(id_pedido_pagto_braspag) = 0 then
		Response.Write "Identificador informado é inválido (" & id_pedido_pagto_braspag & ")!!"
		Response.End
		end if
	
	if converte_numero(id_pedido_pagto_braspag_af) = 0 then
		Response.Write "Identificador da operação com o Antifraude é inválido (" & id_pedido_pagto_braspag_af & ")!!"
		Response.End
		end if
	
	if converte_numero(id_pedido_pagto_braspag_pag) = 0 then
		Response.Write "Identificador da operação com o Pagador é inválido (" & id_pedido_pagto_braspag_pag & ")!!"
		Response.End
		end if
	
	dim cn
	dim strResp
	dim trx_AF, trx_PAG, r_rx_AF, r_rx_PAG
	
'	CONECTA COM O BANCO DE DADOS
	If Not bdd_conecta(cn) then
		msg_erro = erro_descricao(ERR_CONEXAO)
		Response.Write msg_erro
		Response.End
		end if
	
'	CONSULTA STATUS DA TRANSAÇÃO NO ANTIFRAUDE
	call BraspagProcessaConsulta_FraudAnalysisTransactionDetails(id_pedido_pagto_braspag, id_pedido_pagto_braspag_af, usuario, trx_AF, r_rx_AF, msg_erro)
	if msg_erro <> "" then
		Response.Write msg_erro
		Response.End
		end if
	
'	CONSULTA STATUS DA TRANSAÇÃO NO PAGADOR
	call BraspagVerificaPreRequisito_BraspagTransactionId(id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, usuario, msg_erro)
	
	call BraspagProcessaConsulta_GetTransactionData(id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, usuario, trx_PAG, r_rx_PAG, msg_erro)
	if msg_erro <> "" then
		Response.Write msg_erro
		Response.End
		end if
	
	cn.Close
	set cn = nothing
	
	dim stPagGlobalStatus
	stPagGlobalStatus = decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(r_rx_PAG.PAG_Status)
	
	dim stAfGlobalStatus
	stAfGlobalStatus = decodifica_FraudAnalysisTransactionDetailsResponseAntiFraudTransactionStatusCode_para_GlobalStatus(r_rx_AF.AF_AntiFraudTransactionStatusCode)
	
'	ENVIA RESPOSTA
	strResp = "{ " & _
				"""AF"" : {" & _
					"""id_pedido_pagto_braspag"" : """ & id_pedido_pagto_braspag & """," & _
					"""id_pedido_pagto_braspag_af"" : """ & id_pedido_pagto_braspag_af & """," & _
					"""AF_CorrelatedId"" : """ & r_rx_AF.AF_CorrelatedId & """," & _
					"""AF_AntiFraudMerchantId"" : """ & r_rx_AF.AF_AntiFraudMerchantId & """," & _
					"""AF_AntiFraudTransactionId"" : """ & r_rx_AF.AF_AntiFraudTransactionId & """," & _
					"""AF_GlobalStatus"" : """ & stAfGlobalStatus & """," & _
					"""AF_DescricaoGlobalStatus"" : """ & BraspagAntiFraudeDescricaoGlobalStatus(stAfGlobalStatus) & """," & _
					"""AF_GlobalStatus_atualizacao_data_hora"" : """ & formata_data_hora(Now) & """," & _
					"""AF_AntiFraudReceiveDate"" : """ & r_rx_AF.AF_AntiFraudReceiveDate & """," & _
					"""AF_AntiFraudStatusLastUpdateDate"" : """ & r_rx_AF.AF_AntiFraudStatusLastUpdateDate & """," & _
					"""AF_AntiFraudAnalysisScore"" : """ & r_rx_AF.AF_AntiFraudAnalysisScore & """," & _
					"""AF_BraspagTransactionId"" : """ & r_rx_AF.AF_BraspagTransactionId & """," & _
					"""AF_MerchantOrderId"" : """ & r_rx_AF.AF_MerchantOrderId & """," & _
					"""AF_AntiFraudAcquirerConversionDate"" : """ & r_rx_AF.AF_AntiFraudAcquirerConversionDate & """," & _
					"""AF_AntiFraudTransactionOriginalStatusCode"" : """ & r_rx_AF.AF_AntiFraudTransactionOriginalStatusCode & """," & _
					"""AF_ErrorCode"" : """ & r_rx_AF.AF_ErrorCode & """," & _
					"""AF_ErrorMessage"" : """ & r_rx_AF.AF_ErrorMessage & """," & _
					"""AF_faultcode"" : """ & r_rx_AF.AF_faultcode & """," & _
					"""AF_faultstring"" : """ & r_rx_AF.AF_faultstring & """" & _
				"}," & _
				"""PAG"" : {" & _
					"""id_pedido_pagto_braspag"" : """ & id_pedido_pagto_braspag & """," & _
					"""id_pedido_pagto_braspag_pag"" : """ & id_pedido_pagto_braspag_pag & """," & _
					"""PAG_CorrelationId"" : """ & r_rx_PAG.PAG_CorrelationId & """," & _
					"""PAG_Success"" : """ & r_rx_PAG.PAG_Success & """," & _
					"""PAG_BraspagTransactionId"" : """ & r_rx_PAG.PAG_BraspagTransactionId & """," & _
					"""PAG_OrderId"" : """ & r_rx_PAG.PAG_OrderId & """," & _
					"""PAG_AcquirerTransactionId"" : """ & r_rx_PAG.PAG_AcquirerTransactionId & """," & _
					"""PAG_PaymentMethod"" : """ & r_rx_PAG.PAG_PaymentMethod & """," & _
					"""PAG_PaymentMethodName"" : """ & r_rx_PAG.PAG_PaymentMethodName & """," & _
					"""PAG_Amount"" : """ & r_rx_PAG.PAG_Amount & """," & _
					"""PAG_AuthorizationCode"" : """ & r_rx_PAG.PAG_AuthorizationCode & """," & _
					"""PAG_NumberOfPayments"" : """ & r_rx_PAG.PAG_NumberOfPayments & """," & _
					"""PAG_Currency"" : """ & r_rx_PAG.PAG_Currency & """," & _
					"""PAG_Country"" : """ & r_rx_PAG.PAG_Country & """," & _
					"""PAG_TransactionType"" : """ & r_rx_PAG.PAG_TransactionType & """," & _
					"""PAG_GlobalStatus"" : """ & stPagGlobalStatus & """," & _
					"""PAG_DescricaoGlobalStatus"" : """ & BraspagPagadorDescricaoGlobalStatus(stPagGlobalStatus) & """," & _
					"""PAG_GlobalStatus_atualizacao_data_hora"" : """ & formata_data_hora(Now) & """," & _
					"""PAG_ReceivedDate"" : """ & r_rx_PAG.PAG_ReceivedDate & """," & _
					"""PAG_CapturedDate"" : """ & r_rx_PAG.PAG_CapturedDate & """," & _
					"""PAG_VoidedDate"" : """ & r_rx_PAG.PAG_VoidedDate & """," & _
					"""PAG_CreditCardToken"" : """ & r_rx_PAG.PAG_CreditCardToken & """," & _
					"""PAG_ProofOfSale"" : """ & r_rx_PAG.PAG_ProofOfSale & """," & _
					"""PAG_ErrorCode"" : """ & r_rx_PAG.PAG_ErrorCode & """," & _
					"""PAG_ErrorMessage"" : """ & r_rx_PAG.PAG_ErrorMessage & """," & _
					"""PAG_faultcode"" : """ & r_rx_PAG.PAG_faultcode & """," & _
					"""PAG_faultstring"" : """ & r_rx_PAG.PAG_faultstring & """" & _
				"}" & _
			" }"
	Response.Write strResp
	Response.End
%>
