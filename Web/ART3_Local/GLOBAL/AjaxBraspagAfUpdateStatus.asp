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
'	  AjaxBraspagAfUpdateStatus.asp
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
	
	dim s, usuario, af_decision, af_comentario, id_pedido_pagto_braspag, id_pedido_pagto_braspag_af, id_pedido_pagto_braspag_pag, msg_erro
	usuario = Trim(Request("usuario"))
	af_decision = Trim(Request("af_decision"))
'	The jquery doc says: "Data will always be transmitted to the server using UTF-8 charset; you must decode this appropriately on the server side."
	af_comentario = retira_acentuacao(Trim(DecodeUTF8(Request("af_comentario"))))
	id_pedido_pagto_braspag = Trim(Request("id_pedido_pagto_braspag"))
	id_pedido_pagto_braspag_af = Trim(Request("id_pedido_pagto_braspag_af"))
	id_pedido_pagto_braspag_pag = Trim(Request("id_pedido_pagto_braspag_pag"))
	
	if af_decision = "" then
		Response.Write "Decisão sobre a análise de fraude não informada!!"
		Response.End
		end if
	
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
	
	dim cn, tPPB
	dim strResp, strRespAf, strRespPag
	strRespAf = ""
	strRespPag = ""

	dim stAfSucessoOperacao, stPagSucessoOperacao
	stAfSucessoOperacao = "0"
	stPagSucessoOperacao = "0"
	
	dim blnErroPag, blnErroAf
	blnErroPag = False
	blnErroAf = False
	
	dim strMsgErroPag, strMsgErroAf
	strMsgErroPag = ""
	strMsgErroAf = ""
	
'	CONECTA COM O BANCO DE DADOS
	If Not bdd_conecta(cn) then
		msg_erro = erro_descricao(ERR_CONEXAO)
		Response.Write msg_erro
		Response.End
		end if
	
'	CONSULTA STATUS DA TRANSAÇÃO NO ANTIFRAUDE
	dim trx_AF, r_rx_AF
	dim trx_AF_TD, r_rx_AF_TD

	dim stAfGlobalStatus
	stAfGlobalStatus = ""
	
	dim blnAfStatusJaRevisado
	blnAfStatusJaRevisado = False
	
	call BraspagProcessaConsulta_FraudAnalysisTransactionDetails(id_pedido_pagto_braspag, id_pedido_pagto_braspag_af, usuario, trx_AF_TD, r_rx_AF_TD, msg_erro)
	if msg_erro = "" then
		stAfGlobalStatus = decodifica_FraudAnalysisTransactionDetailsResponseAntiFraudTransactionStatusCode_para_GlobalStatus(r_rx_AF_TD.AF_AntiFraudTransactionStatusCode)
		if stAfGlobalStatus <> BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW then blnAfStatusJaRevisado = True
		end if
	
	if Not blnAfStatusJaRevisado then
		call BraspagProcessaRequisicao_AF_UpdateStatus(af_decision, id_pedido_pagto_braspag, id_pedido_pagto_braspag_af, usuario, af_comentario, trx_AF, r_rx_AF, msg_erro)
		if msg_erro <> "" then
			blnErroAf = True
			if strMsgErroAf <> "" then strMsgErroAf = strMsgErroAf & chr(13)
			strMsgErroAf = strMsgErroAf & msg_erro
			end if
		end if
	
	if blnAfStatusJaRevisado then
		strRespAf = "{ " & _
					"""id_pedido_pagto_braspag"" : """ & id_pedido_pagto_braspag & """," & _
					"""id_pedido_pagto_braspag_af"" : """ & id_pedido_pagto_braspag_af & """," & _
					"""AF_usuario"" : """ & usuario & """," & _
					"""AF_NewDecision"" : """ & af_decision & """," & _
					"""AF_SucessoOperacao"" : """ & "1" & """," & _
					"""AF_CorrelatedId"" : """ & "" & """," & _
					"""AF_AntiFraudTransactionId"" : """ & "" & """," & _
					"""AF_GlobalStatus"" : """ & stAfGlobalStatus & """," & _
					"""AF_DescricaoGlobalStatus"" : """ & BraspagAntiFraudeDescricaoGlobalStatus(stAfGlobalStatus) & """," & _
					"""AF_GlobalStatus_atualizacao_data_hora"" : """ & "" & """," & _
					"""AF_ErrorCode"" : """ & "" & """," & _
					"""AF_ErrorMessage"" : """ & "" & """," & _
					"""AF_faultcode"" : """ & "" & """," & _
					"""AF_faultstring"" : """ & "" & """" & _
				" }"
	elseif blnErroAf then
		strRespAf = "{ " & _
					"""id_pedido_pagto_braspag"" : """ & id_pedido_pagto_braspag & """," & _
					"""id_pedido_pagto_braspag_af"" : """ & id_pedido_pagto_braspag_af & """," & _
					"""AF_usuario"" : """ & usuario & """," & _
					"""AF_NewDecision"" : """ & af_decision & """," & _
					"""AF_SucessoOperacao"" : """ & stAfSucessoOperacao & """," & _
					"""AF_CorrelatedId"" : """ & "" & """," & _
					"""AF_AntiFraudTransactionId"" : """ & "" & """," & _
					"""AF_GlobalStatus"" : """ & "" & """," & _
					"""AF_DescricaoGlobalStatus"" : """ & "" & """," & _
					"""AF_GlobalStatus_atualizacao_data_hora"" : """ & "" & """," & _
					"""AF_ErrorCode"" : """ & "" & """," & _
					"""AF_ErrorMessage"" : """ & strMsgErroAf & """," & _
					"""AF_faultcode"" : """ & "" & """," & _
					"""AF_faultstring"" : """ & "" & """" & _
				" }"
	else
		if r_rx_AF.AF_RequestStatusCode = BRASPAG_ANTIFRAUDE_CARTAO_UPDATESTATUSRESPONSE_REQUESTSTATUSCODE__SUCCESS then
			stAfSucessoOperacao = "1"
			if af_decision = BRASPAG_AF_DECISION__ACCEPT then
				stAfGlobalStatus = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ACCEPT
			elseif af_decision = BRASPAG_AF_DECISION__REJECT then
				stAfGlobalStatus = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REJECT
				end if
			end if
		
		strRespAf = "{ " & _
					"""id_pedido_pagto_braspag"" : """ & id_pedido_pagto_braspag & """," & _
					"""id_pedido_pagto_braspag_af"" : """ & id_pedido_pagto_braspag_af & """," & _
					"""AF_usuario"" : """ & usuario & """," & _
					"""AF_NewDecision"" : """ & af_decision & """," & _
					"""AF_SucessoOperacao"" : """ & stAfSucessoOperacao & """," & _
					"""AF_CorrelatedId"" : """ & r_rx_AF.AF_CorrelatedId & """," & _
					"""AF_AntiFraudTransactionId"" : """ & r_rx_AF.AF_AntiFraudTransactionId & """," & _
					"""AF_GlobalStatus"" : """ & stAfGlobalStatus & """," & _
					"""AF_DescricaoGlobalStatus"" : """ & BraspagAntiFraudeDescricaoGlobalStatus(stAfGlobalStatus) & """," & _
					"""AF_GlobalStatus_atualizacao_data_hora"" : """ & formata_data_hora(Now) & """," & _
					"""AF_ErrorCode"" : """ & r_rx_AF.AF_ErrorCode & """," & _
					"""AF_ErrorMessage"" : """ & r_rx_AF.AF_ErrorMessage & """," & _
					"""AF_faultcode"" : """ & r_rx_AF.AF_faultcode & """," & _
					"""AF_faultstring"" : """ & r_rx_AF.AF_faultstring & """" & _
				" }"
		end if 'if blnAfStatusJaRevisado
	
	
'	EXECUTA A CONSULTA P/ ATUALIZAR O STATUS E VERIFICAR SE A TRANSAÇÃO ESTÁ CAPTURADA E QUAL É A DATA DA CAPTURA
	dim trx_Get_1, r_rx_Get_1
	dim trx_Get_2, r_rx_Get_2
	dim trx_Void, r_rx_Void
	dim trx_Refund, r_rx_Refund
	dim stPagGlobalStatus_1, stPagGlobalStatus_2
	dim blnRequestVoid, blnTransacaoJaCanceladaEstornada
	dim dtCapturedDate

	blnTransacaoJaCanceladaEstornada = False
	
	call BraspagProcessaConsulta_GetTransactionData(id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, usuario, trx_Get_1, r_rx_Get_1, msg_erro)
	if msg_erro <> "" then
		blnErroPag = True
		if strMsgErroPag <> "" then strMsgErroPag = strMsgErroPag & chr(13)
		strMsgErroPag = strMsgErroPag & msg_erro
		end if
	
	if Not blnErroPag then
		if Trim("" & r_rx_Get_1.PAG_CapturedDate) <> "" then
			dtCapturedDate = converte_para_datetime_from_mmddyyyy_hhmmss_am_pm(r_rx_Get_1.PAG_CapturedDate)
			end if
		end if
	
	if af_decision = BRASPAG_AF_DECISION__REJECT then
		if Not blnErroPag then
			stPagGlobalStatus_1 = decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(r_rx_Get_1.PAG_Status)
			if (Trim("" & stPagGlobalStatus_1) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURA_CANCELADA) Or _
				(Trim("" & stPagGlobalStatus_1) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNADA) then
				blnTransacaoJaCanceladaEstornada = True
			else
				if Trim("" & stPagGlobalStatus_1) <> Trim("" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA) then
					blnErroPag = True
					if strMsgErroPag <> "" then strMsgErroPag = strMsgErroPag & chr(13)
					strMsgErroPag = strMsgErroPag & "Não é possível realizar o cancelamento/estorno porque a transação não está com status 'Capturada'"
					end if
				end if
			end if
		
		if Not blnTransacaoJaCanceladaEstornada then
			if Not blnErroPag then
				if Trim("" & r_rx_Get_1.PAG_CapturedDate) = "" then
					blnErroPag = True
					if strMsgErroPag <> "" then strMsgErroPag = strMsgErroPag & chr(13)
					strMsgErroPag = strMsgErroPag & "Não é possível realizar o cancelamento/estorno porque a data de captura da transação não foi informada pela Braspag"
					end if
				end if
			
			if Not blnErroPag then
			'	O CANCELAMENTO (VOID) PODE SER FEITO ATÉ AS 23:59:59 DO MESMO DIA DA CAPTURA, APÓS ISSO DEVE SER FEITO UM ESTORNO (REFUND)
				if formata_data_yyyymmdd(Date) = formata_data_yyyymmdd(dtCapturedDate) then
					blnRequestVoid = True
				else
					blnRequestVoid = False
					end if
				
				if blnRequestVoid then
				'	CANCELAMENTO (VOID)
					call BraspagProcessaRequisicao_VoidCreditCardTransaction(id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, usuario, trx_Void, r_rx_Void, msg_erro)
					if msg_erro <> "" then
						blnErroPag = True
						if strMsgErroPag <> "" then strMsgErroPag = strMsgErroPag & chr(13)
						strMsgErroPag = strMsgErroPag & msg_erro
						end if
				else
				'	ESTORNO (REFUND)
					call BraspagProcessaRequisicao_RefundCreditCardTransaction(id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, usuario, trx_Refund, r_rx_Refund, msg_erro)
					if msg_erro <> "" then
						blnErroPag = True
						if strMsgErroPag <> "" then strMsgErroPag = strMsgErroPag & chr(13)
						strMsgErroPag = strMsgErroPag & msg_erro
						end if
					end if
				end if 'if Not blnErroPag
			
		'	REALIZA NOVA CONSULTA P/ OBTER O STATUS ATUALIZADO
			if Not blnErroPag then
				call BraspagProcessaConsulta_GetTransactionData(id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, usuario, trx_Get_2, r_rx_Get_2, msg_erro)
				if msg_erro <> "" then
					blnErroPag = True
					if strMsgErroPag <> "" then strMsgErroPag = strMsgErroPag & chr(13)
					strMsgErroPag = strMsgErroPag & msg_erro
				else
					stPagGlobalStatus_2 = decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(r_rx_Get_2.PAG_Status)
					end if
				end if 'if Not blnErroPag
			end if 'if Not blnTransacaoJaCanceladaEstornada
		end if 'if af_decision = BRASPAG_AF_DECISION__REJECT
	
	
	if blnErroPag then
	'	MONTA RESPOSTA DE ERRO
		strRespPag = "{ " & _
					"""id_pedido_pagto_braspag"" : """ & id_pedido_pagto_braspag & """," & _
					"""id_pedido_pagto_braspag_pag"" : """ & id_pedido_pagto_braspag_pag & """," & _
					"""PAG_TipoOperacao"" : """ & "" & """," & _
					"""PAG_SucessoOperacao"" : """ & stPagSucessoOperacao & """," & _
					"""PAG_CorrelationId"" : """ & "" & """," & _
					"""PAG_BraspagTransactionId"" : """ & "" & """," & _
					"""PAG_AcquirerTransactionId"" : """ & "" & """," & _
					"""PAG_Amount"" : """ & "" & """," & _
					"""PAG_AuthorizationCode"" : """ & "" & """," & _
					"""PAG_ReturnCode"" : """ & "" & """," & _
					"""PAG_ReturnMessage"" : """ & "" & """," & _
					"""PAG_GlobalStatus"" : """ & "" & """," & _
					"""PAG_DescricaoGlobalStatus"" : """ & "" & """," & _
					"""PAG_GlobalStatus_atualizacao_data_hora"" : """ & "" & """," & _
					"""PAG_ProofOfSale"" : """ & "" & """," & _
					"""PAG_ServiceTaxAmount"" : """ & "" & """," & _
					"""PAG_ErrorCode"" : """ & "" & """," & _
					"""PAG_ErrorMessage"" : """ & strMsgErroPag & """," & _
					"""PAG_faultcode"" : """ & "" & """," & _
					"""PAG_faultstring"" : """ & "" & """" & _
				" }"
	elseif af_decision = BRASPAG_AF_DECISION__ACCEPT then
		dim stPagGlobalStatusBd, dtHrPagGlobalStatusBd
		s = "SELECT ult_PAG_GlobalStatus, ult_PAG_atualizacao_data_hora FROM t_PEDIDO_PAGTO_BRASPAG WHERE (id = " & id_pedido_pagto_braspag & ")"
		set tPPB = cn.Execute(s)
		if Not tPPB.Eof then
			stPagGlobalStatusBd = tPPB("ult_PAG_GlobalStatus")
			dtHrPagGlobalStatusBd = tPPB("ult_PAG_atualizacao_data_hora")
			tPPB.Close
			end if
		set tPPB = nothing
		strRespPag = "{ " & _
					"""id_pedido_pagto_braspag"" : """ & id_pedido_pagto_braspag & """," & _
					"""id_pedido_pagto_braspag_pag"" : """ & id_pedido_pagto_braspag_pag & """," & _
					"""PAG_TipoOperacao"" : """ & "" & """," & _
					"""PAG_SucessoOperacao"" : """ & "1" & """," & _
					"""PAG_CorrelationId"" : """ & "" & """," & _
					"""PAG_BraspagTransactionId"" : """ & "" & """," & _
					"""PAG_AcquirerTransactionId"" : """ & "" & """," & _
					"""PAG_Amount"" : """ & "" & """," & _
					"""PAG_AuthorizationCode"" : """ & "" & """," & _
					"""PAG_ReturnCode"" : """ & "" & """," & _
					"""PAG_ReturnMessage"" : """ & "" & """," & _
					"""PAG_GlobalStatus"" : """ & stPagGlobalStatusBd & """," & _
					"""PAG_DescricaoGlobalStatus"" : """ & BraspagPagadorDescricaoGlobalStatus(stPagGlobalStatusBd) & """," & _
					"""PAG_GlobalStatus_atualizacao_data_hora"" : """ & formata_data_hora(dtHrPagGlobalStatusBd) & """," & _
					"""PAG_ProofOfSale"" : """ & "" & """," & _
					"""PAG_ServiceTaxAmount"" : """ & "" & """," & _
					"""PAG_ErrorCode"" : """ & "" & """," & _
					"""PAG_ErrorMessage"" : """ & "" & """," & _
					"""PAG_faultcode"" : """ & "" & """," & _
					"""PAG_faultstring"" : """ & "" & """" & _
				" }"
	elseif blnTransacaoJaCanceladaEstornada then
		strRespPag = "{ " & _
					"""id_pedido_pagto_braspag"" : """ & id_pedido_pagto_braspag & """," & _
					"""id_pedido_pagto_braspag_pag"" : """ & id_pedido_pagto_braspag_pag & """," & _
					"""PAG_TipoOperacao"" : """ & "" & """," & _
					"""PAG_SucessoOperacao"" : """ & "1" & """," & _
					"""PAG_CorrelationId"" : """ & "" & """," & _
					"""PAG_BraspagTransactionId"" : """ & "" & """," & _
					"""PAG_AcquirerTransactionId"" : """ & "" & """," & _
					"""PAG_Amount"" : """ & "" & """," & _
					"""PAG_AuthorizationCode"" : """ & "" & """," & _
					"""PAG_ReturnCode"" : """ & "" & """," & _
					"""PAG_ReturnMessage"" : """ & "" & """," & _
					"""PAG_GlobalStatus"" : """ & stPagGlobalStatus_1 & """," & _
					"""PAG_DescricaoGlobalStatus"" : """ & BraspagPagadorDescricaoGlobalStatus(stPagGlobalStatus_1) & """," & _
					"""PAG_GlobalStatus_atualizacao_data_hora"" : """ & "" & """," & _
					"""PAG_ProofOfSale"" : """ & "" & """," & _
					"""PAG_ServiceTaxAmount"" : """ & "" & """," & _
					"""PAG_ErrorCode"" : """ & "" & """," & _
					"""PAG_ErrorMessage"" : """ & "" & """," & _
					"""PAG_faultcode"" : """ & "" & """," & _
					"""PAG_faultstring"" : """ & "" & """" & _
				" }"
	else
		if blnRequestVoid then
			if r_rx_Void.PAG_Status = BRASPAG_PAGADOR_CARTAO_VOIDCREDITCARDTRANSACTIONRESPONSE_STATUS__VOID_CONFIRMED then stPagSucessoOperacao = "1"
			
		'	MONTA RESPOSTA
			strRespPag = "{ " & _
						"""id_pedido_pagto_braspag"" : """ & id_pedido_pagto_braspag & """," & _
						"""id_pedido_pagto_braspag_pag"" : """ & id_pedido_pagto_braspag_pag & """," & _
						"""PAG_TipoOperacao"" : """ & "VOID" & """," & _
						"""PAG_SucessoOperacao"" : """ & stPagSucessoOperacao & """," & _
						"""PAG_CorrelationId"" : """ & r_rx_Void.PAG_CorrelationId & """," & _
						"""PAG_BraspagTransactionId"" : """ & r_rx_Void.PAG_BraspagTransactionId & """," & _
						"""PAG_AcquirerTransactionId"" : """ & r_rx_Void.PAG_AcquirerTransactionId & """," & _
						"""PAG_Amount"" : """ & r_rx_Void.PAG_Amount & """," & _
						"""PAG_AuthorizationCode"" : """ & r_rx_Void.PAG_AuthorizationCode & """," & _
						"""PAG_ReturnCode"" : """ & r_rx_Void.PAG_ReturnCode & """," & _
						"""PAG_ReturnMessage"" : """ & r_rx_Void.PAG_ReturnMessage & """," & _
						"""PAG_GlobalStatus"" : """ & stPagGlobalStatus_2 & """," & _
						"""PAG_DescricaoGlobalStatus"" : """ & BraspagPagadorDescricaoGlobalStatus(stPagGlobalStatus_2) & """," & _
						"""PAG_GlobalStatus_atualizacao_data_hora"" : """ & formata_data_hora(Now) & """," & _
						"""PAG_ProofOfSale"" : """ & r_rx_Void.PAG_ProofOfSale & """," & _
						"""PAG_ServiceTaxAmount"" : """ & r_rx_Void.PAG_ServiceTaxAmount & """," & _
						"""PAG_ErrorCode"" : """ & r_rx_Void.PAG_ErrorCode & """," & _
						"""PAG_ErrorMessage"" : """ & r_rx_Void.PAG_ErrorMessage & """," & _
						"""PAG_faultcode"" : """ & r_rx_Void.PAG_faultcode & """," & _
						"""PAG_faultstring"" : """ & r_rx_Void.PAG_faultstring & """" & _
					" }"
		else 'if blnRequestVoid then
			if r_rx_Refund.PAG_Status = BRASPAG_PAGADOR_CARTAO_REFUNDCREDITCARDTRANSACTIONRESPONSE_STATUS__REFUND_CONFIRMED then stPagSucessoOperacao = "1"
			
		'	MONTA RESPOSTA
			strRespPag = "{ " & _
						"""id_pedido_pagto_braspag"" : """ & id_pedido_pagto_braspag & """," & _
						"""id_pedido_pagto_braspag_pag"" : """ & id_pedido_pagto_braspag_pag & """," & _
						"""PAG_TipoOperacao"" : """ & "REFUND" & """," & _
						"""PAG_SucessoOperacao"" : """ & stPagSucessoOperacao & """," & _
						"""PAG_CorrelationId"" : """ & r_rx_Refund.PAG_CorrelationId & """," & _
						"""PAG_BraspagTransactionId"" : """ & r_rx_Refund.PAG_BraspagTransactionId & """," & _
						"""PAG_AcquirerTransactionId"" : """ & r_rx_Refund.PAG_AcquirerTransactionId & """," & _
						"""PAG_Amount"" : """ & r_rx_Refund.PAG_Amount & """," & _
						"""PAG_AuthorizationCode"" : """ & r_rx_Refund.PAG_AuthorizationCode & """," & _
						"""PAG_ReturnCode"" : """ & r_rx_Refund.PAG_ReturnCode & """," & _
						"""PAG_ReturnMessage"" : """ & r_rx_Refund.PAG_ReturnMessage & """," & _
						"""PAG_GlobalStatus"" : """ & stPagGlobalStatus_2 & """," & _
						"""PAG_DescricaoGlobalStatus"" : """ & BraspagPagadorDescricaoGlobalStatus(stPagGlobalStatus_2) & """," & _
						"""PAG_GlobalStatus_atualizacao_data_hora"" : """ & formata_data_hora(Now) & """," & _
						"""PAG_ProofOfSale"" : """ & r_rx_Refund.PAG_ProofOfSale & """," & _
						"""PAG_ServiceTaxAmount"" : """ & r_rx_Refund.PAG_ServiceTaxAmount & """," & _
						"""PAG_ErrorCode"" : """ & r_rx_Refund.PAG_ErrorCode & """," & _
						"""PAG_ErrorMessage"" : """ & r_rx_Refund.PAG_ErrorMessage & """," & _
						"""PAG_faultcode"" : """ & r_rx_Refund.PAG_faultcode & """," & _
						"""PAG_faultstring"" : """ & r_rx_Refund.PAG_faultstring & """" & _
					" }"
			end if 'if blnRequestVoid then
		end if 'if blnTransacaoJaCanceladaEstornada
	
	cn.Close
	set cn = nothing
	
	strResp = "{ " & _
				"""AF"" : " & strRespAf & _
				"," & _
				"""PAG"" : " & strRespPag & _
			" }"
	
'	ENVIA RESPOSTA
	Response.Write strResp
	Response.End
%>
