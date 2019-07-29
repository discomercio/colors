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
'	  AjaxBraspagPagConsultaStatus.asp
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
	
	dim usuario, id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, msg_erro
	usuario = Trim(Request("usuario"))
	id_pedido_pagto_braspag = Trim(Request("id_pedido_pagto_braspag"))
	id_pedido_pagto_braspag_pag = Trim(Request("id_pedido_pagto_braspag_pag"))
	
	if id_pedido_pagto_braspag = "" then
		Response.Write "Identificador não informado!!"
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
	
	if converte_numero(id_pedido_pagto_braspag_pag) = 0 then
		Response.Write "Identificador da operação com o Pagador é inválido (" & id_pedido_pagto_braspag_pag & ")!!"
		Response.End
		end if
	
	dim cn
	dim strResp
	dim trx, r_rx
	
'	CONECTA COM O BANCO DE DADOS
	If Not bdd_conecta(cn) then
		msg_erro = erro_descricao(ERR_CONEXAO)
		Response.Write msg_erro
		Response.End
		end if
	
	call BraspagVerificaPreRequisito_BraspagTransactionId(id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, usuario, msg_erro)
	
	call BraspagProcessaConsulta_GetTransactionData(id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, usuario, trx, r_rx, msg_erro)
	if msg_erro <> "" then
		Response.Write msg_erro
		Response.End
		end if
	
	cn.Close
	set cn = nothing
	
	dim stGlobalStatus
	stGlobalStatus = decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(r_rx.PAG_Status)
	
'	ENVIA RESPOSTA
	strResp = "{ " & _
				"""id_pedido_pagto_braspag"" : """ & id_pedido_pagto_braspag & """," & _
				"""id_pedido_pagto_braspag_pag"" : """ & id_pedido_pagto_braspag_pag & """," & _
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
				"""PAG_GlobalStatus"" : """ & stGlobalStatus & """," & _
				"""PAG_DescricaoGlobalStatus"" : """ & BraspagPagadorDescricaoGlobalStatus(stGlobalStatus) & """," & _
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
