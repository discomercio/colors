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
'	  AjaxBraspagClearsalePagConsultaStatus.asp
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
	dim trx, r_rx
	
'	CONECTA COM O BANCO DE DADOS
	If Not bdd_conecta(cn) then
		msg_erro = erro_descricao(ERR_CONEXAO)
		Response.Write msg_erro
		Response.End
		end if
	
	call BraspagClearsaleVerificaPreRequisito_BraspagTransactionId(id_pagto_gw_pag, id_pagto_gw_pag_payment, usuario, msg_erro)
	
	call BraspagClearsaleProcessaConsulta_GetTransactionData(id_pagto_gw_pag, id_pagto_gw_pag_payment, usuario, trx, r_rx, msg_erro)
	if msg_erro <> "" then
		Response.Write msg_erro
		Response.End
		end if
	
	dim stGlobalStatus
	stGlobalStatus = decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(r_rx.PAG_Status)

	dim blnEstornoPendente
	blnEstornoPendente = False

	if stGlobalStatus = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA then
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
		end if

	cn.Close
	set cn = nothing
	
	dim stGlobalStatusAux, stDescricaoGlobalStatusAux
	if blnEstornoPendente then
		stGlobalStatusAux = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNO_PENDENTE
		stDescricaoGlobalStatusAux = BraspagPagadorDescricaoGlobalStatus(BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNO_PENDENTE)
	else
		stGlobalStatusAux = stGlobalStatus
		stDescricaoGlobalStatusAux = BraspagPagadorDescricaoGlobalStatus(stGlobalStatus)
		end if

'	ENVIA RESPOSTA
	strResp = "{ " & _
				"""id_pagto_gw_pag"" : """ & id_pagto_gw_pag & """," & _
				"""id_pagto_gw_pag_payment"" : """ & id_pagto_gw_pag_payment & """," & _
				"""PAG_CorrelationId"" : """ & r_rx.PAG_CorrelationId & """," & _
				"""PAG_Success"" : """ & r_rx.PAG_Success & """," & _
				"""PAG_BraspagTransactionId"" : """ & r_rx.PAG_BraspagTransactionId & """," & _
				"""PAG_OrderId"" : """ & r_rx.PAG_OrderId & """," & _
				"""PAG_AcquirerTransactionId"" : """ & r_rx.PAG_AcquirerTransactionId & """," & _
				"""PAG_PaymentMethod"" : """ & r_rx.PAG_PaymentMethod & """," & _
				"""PAG_PaymentMethodName"" : """ & r_rx.PAG_PaymentMethodName & """," & _
				"""PAG_Amount"" : """ & r_rx.PAG_Amount & """," & _
				"""PAG_AuthorizationCode"" : """ & r_rx.PAG_AuthorizationCode & """," & _
				"""PAG_NumberOfPayments"" : """ & r_rx.PAG_NumberOfPayments & """," & _
				"""PAG_Currency"" : """ & r_rx.PAG_Currency & """," & _
				"""PAG_Country"" : """ & r_rx.PAG_Country & """," & _
				"""PAG_TransactionType"" : """ & r_rx.PAG_TransactionType & """," & _
				"""PAG_GlobalStatus"" : """ & stGlobalStatusAux & """," & _
				"""PAG_DescricaoGlobalStatus"" : """ & stDescricaoGlobalStatusAux & """," & _
				"""PAG_GlobalStatus_atualizacao_data_hora"" : """ & formata_data_hora(Now) & """," & _
				"""PAG_ReceivedDate"" : """ & r_rx.PAG_ReceivedDate & """," & _
				"""PAG_CapturedDate"" : """ & r_rx.PAG_CapturedDate & """," & _
				"""PAG_VoidedDate"" : """ & r_rx.PAG_VoidedDate & """," & _
				"""PAG_CreditCardToken"" : """ & r_rx.PAG_CreditCardToken & """," & _
				"""PAG_ProofOfSale"" : """ & r_rx.PAG_ProofOfSale & """," & _
				"""PAG_ErrorCode"" : """ & r_rx.PAG_ErrorCode & """," & _
				"""PAG_ErrorMessage"" : """ & r_rx.PAG_ErrorMessage & """," & _
				"""PAG_faultcode"" : """ & r_rx.PAG_faultcode & """," & _
				"""PAG_faultstring"" : """ & r_rx.PAG_faultstring & """" & _
			" }"
	Response.Write strResp
	Response.End
%>
