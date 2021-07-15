<%
' =========================================
'          C O N S T A N T E S
' =========================================

	const AF_END_ETG_DUMMY__STREET1 = "1295 Charleston Road"
	const AF_END_ETG_DUMMY__CITY = "Mountain View"
	const AF_END_ETG_DUMMY__STATE = "CA"
	const AF_END_ETG_DUMMY__COUNTRY = "US"
	const AF_END_ETG_DUMMY__POSTALCODE = "94043"
	
	
'	Authorize
'	=========
	class cl_BRASPAG_Authorize_TX
		dim RequestId
		dim Version
		dim OrderData_MerchantId
		dim OrderData_OrderId
		dim CustomerData_CustomerIdentity
		dim CustomerData_CustomerIdentityType
		dim CustomerData_CustomerName
		dim CustomerData_CustomerEmail
		end class

	class cl_BRASPAG_Authorize_PaymentDataRequest_TX
		dim bandeira
		dim PAG_PaymentMethod
		dim PAG_Amount
		dim PAG_Currency
		dim PAG_Country
		dim PAG_ServiceTaxAmount
		dim PAG_NumberOfPayments
		dim PAG_PaymentPlan
		dim PAG_TransactionType
		dim PAG_CardHolder
		dim PAG_CardNumber
		dim PAG_CardSecurityCode
		dim PAG_CardExpirationDate
		end class

	class cl_BRASPAG_Authorize_RX_PAG_ERROR
		dim ErrorCode
		dim ErrorMessage
		end class

	class cl_BRASPAG_Authorize_RX
		dim PAG_CorrelationId
		dim PAG_Success
		dim PAG_OrderData_OrderId
		dim PAG_OrderData_BraspagOrderId
		end class

	class cl_BRASPAG_Authorize_PaymentDataRequest_RX
		dim bandeira
		dim PAG_BraspagTransactionId
		dim PAG_PaymentMethod
		dim PAG_Amount
		dim PAG_AcquirerTransactionId
		dim PAG_AuthorizationCode
		dim PAG_ReturnCode
		dim PAG_ReturnMessage
		dim PAG_Status
		dim PAG_CreditCardToken
		dim PAG_ProofOfSale
		end class


'	AnalyseAndAuthorizeOnSuccess
'	============================
	class cl_BRASPAG_AnalyseAndAuthorizeOnSuccess_TX
		dim AF_RequestId
		dim AF_Version
		dim AF_MerchantId
		dim AF_AntiFraudSequenceType
		
		dim AF_DocumentData_Cpf
		dim AF_DocumentData_Cnpj
		
		dim AF_AntiFraudRequest_BillToData_CustomerId
		dim AF_AntiFraudRequest_BillToData_City
		dim AF_AntiFraudRequest_BillToData_Country
		dim AF_AntiFraudRequest_BillToData_Email
		dim AF_AntiFraudRequest_BillToData_FirstName
		dim AF_AntiFraudRequest_BillToData_LastName
		dim AF_AntiFraudRequest_BillToData_State
		dim AF_AntiFraudRequest_BillToData_Street1
		dim AF_AntiFraudRequest_BillToData_Street2
		dim AF_AntiFraudRequest_BillToData_PostalCode
		dim AF_AntiFraudRequest_BillToData_PhoneNumber
		dim AF_AntiFraudRequest_BillToData_IpAddress
		
		dim AF_AntiFraudRequest_ShipToData_City
		dim AF_AntiFraudRequest_ShipToData_Country
		dim AF_AntiFraudRequest_ShipToData_FirstName
		dim AF_AntiFraudRequest_ShipToData_LastName
		dim AF_AntiFraudRequest_ShipToData_PhoneNumber
		dim AF_AntiFraudRequest_ShipToData_PostalCode
		dim AF_AntiFraudRequest_ShipToData_ShippingMethod
		dim AF_AntiFraudRequest_ShipToData_State
		dim AF_AntiFraudRequest_ShipToData_Street1
		dim AF_AntiFraudRequest_ShipToData_Street2
		
		dim AF_AntiFraudRequest_DeviceFingerPrintId
		
		dim AF_AntiFraudRequest_CardData_AccountNumber
		dim AF_AntiFraudRequest_CardData_Card
		dim AF_AntiFraudRequest_CardData_ExpirationMonth
		dim AF_AntiFraudRequest_CardData_ExpirationYear
		
		dim AF_AntiFraudRequest_PurchaseTotalsData_Currency
		dim AF_AntiFraudRequest_PurchaseTotalsData_GrandTotalAmount
		
		dim AF_AntiFraudRequest_MerchantReferenceCode
		
		dim PAG_RequestId
		dim PAG_Version
		dim PAG_OrderData_MerchantId
		dim PAG_OrderData_OrderId
		dim PAG_CustomerData_CustomerIdentity
		dim PAG_CustomerData_CustomerName
		dim PAG_PaymentDataCollection_PaymentMethod
		dim PAG_PaymentDataCollection_Amount
		dim PAG_PaymentDataCollection_Currency
		dim PAG_PaymentDataCollection_Country
		dim PAG_PaymentDataCollection_ServiceTaxAmount
		dim PAG_PaymentDataCollection_NumberOfPayments
		dim PAG_PaymentDataCollection_PaymentPlan
		dim PAG_PaymentDataCollection_TransactionType
		dim PAG_PaymentDataCollection_CardHolder
		dim PAG_PaymentDataCollection_CardNumber
		dim PAG_PaymentDataCollection_CardSecurityCode
		dim PAG_PaymentDataCollection_CardExpirationDate
		end class
	
	class cl_BRASPAG_AnalyseAndAuthorizeOnSuccess_Item_TX
		dim ProductData_Name
		dim ProductData_Sku
		dim ProductData_Quantity
		dim ProductData_UnitPrice
		end class
	
'	O AdditionalDataCollection é usado para enviar os dados de MerchantDefinedData!!
	class cl_BRASPAG_AnalyseAndAuthorizeOnSuccess_AdditionalData_TX
		dim Id
		dim Value
		end class
	
	class cl_BRASPAG_AnalyseAndAuthorizeOnSuccess_RX
		dim AF_CorrelatedId
		dim AF_Success
		dim AF_AntiFraudTransactionId
		dim AF_TransactionStatusCode
		dim AF_TransactionStatusDescription
		dim AF_AFResp_AfsReply_AddressInfoCode
		dim AF_AFResp_AfsReply_AfsFactorCode
		dim AF_AFResp_AfsReply_AfsResult
		dim AF_AFResp_AfsReply_BinCountry
		dim AF_AFResp_AfsReply_CardAccount
		dim AF_AFResp_AfsReply_CardIssuer
		dim AF_AFResp_AfsReply_CardScheme
		dim AF_AFResp_AfsReply_ConsumerLocalTime
		dim AF_AFResp_AfsReply_DF_Data_CookiesEnabled
		dim AF_AFResp_AfsReply_DF_Data_FlashEnabled
		dim AF_AFResp_AfsReply_DF_Data_Hash
		dim AF_AFResp_AfsReply_DF_Data_ImagesEnabled
		dim AF_AFResp_AfsReply_DF_Data_JavascriptEnabled
		dim AF_AFResp_AfsReply_DF_Data_TrueIPAddress
		dim AF_AFResp_AfsReply_DF_Data_TrueIPAddressCity
		dim AF_AFResp_AfsReply_DF_Data_TrueIPAddressCountry
		dim AF_AFResp_AfsReply_DF_Data_SmartID
		dim AF_AFResp_AfsReply_DF_Data_SmartIDConfidenceLevel
		dim AF_AFResp_AfsReply_DF_Data_ScreenResolution
		dim AF_AFResp_AfsReply_DF_Data_BrowserLanguage
		dim AF_AFResp_AfsReply_HostSeverity
		dim AF_AFResp_AfsReply_HotlistInfoCode
		dim AF_AFResp_AfsReply_IdentityInfoCode
		dim AF_AFResp_AfsReply_InternetInfoCode
		dim AF_AFResp_AfsReply_IpRoutingMethod
		dim AF_AFResp_AfsReply_PhoneInfoCode
		dim AF_AFResp_AfsReply_ReasonCode
		dim AF_AFResp_AfsReply_ScoreModelUsed
		dim AF_AFResp_AfsReply_SuspiciousInfoCode
		dim AF_AFResp_AfsReply_VelocityInfoCode
		dim AF_AFResp_Decision
		dim AF_AFResp_DecisionReply_CasePriority
		dim AF_AFResp_MerchantReferenceCode
		dim AF_AFResp_ReasonCode
		dim AF_AFResp_RequestId
		dim AF_AFResp_RequestToken
		dim PAG_CorrelationId
		dim PAG_Success
		dim PAG_OrderData_OrderId
		dim PAG_OrderData_BraspagOrderId
		dim PAG_PaymentDataResponse_BraspagTransactionId
		dim PAG_PaymentDataResponse_PaymentMethod
		dim PAG_PaymentDataResponse_Amount
		dim PAG_PaymentDataResponse_AcquirerTransactionId
		dim PAG_PaymentDataResponse_AuthorizationCode
		dim PAG_PaymentDataResponse_ReturnCode
		dim PAG_PaymentDataResponse_ReturnMessage
		dim PAG_PaymentDataResponse_Status
		dim PAG_PaymentDataResponse_CreditCardToken
		dim PAG_PaymentDataResponse_ProofOfSale
		end class
	
	class cl_BRASPAG_AnalyseAndAuthorizeOnSuccess_RX_AF_ERROR
		dim ErrorCode
		dim ErrorMessage
		end class
	
	class cl_BRASPAG_AnalyseAndAuthorizeOnSuccess_RX_PAG_ERROR
		dim ErrorCode
		dim ErrorMessage
		end class
	
	
'	GetTransactionData
'	==================
	class cl_BRASPAG_GetTransactionData_TX
		dim PAG_Version
		dim PAG_RequestId
		dim PAG_MerchantId
		dim PAG_BraspagTransactionId
		end class
	
	class cl_BRASPAG_GetTransactionDataResponse_RX
		dim PAG_CorrelationId
		dim PAG_Success
		dim PAG_BraspagTransactionId
		dim PAG_OrderId
		dim PAG_AcquirerTransactionId
		dim PAG_PaymentMethod
		dim PAG_PaymentMethodName
		dim PAG_Amount
		dim PAG_AuthorizationCode
		dim PAG_NumberOfPayments
		dim PAG_Currency
		dim PAG_Country
		dim PAG_TransactionType
		dim PAG_Status
		dim PAG_ReceivedDate
		dim PAG_CapturedDate
		dim PAG_VoidedDate
		dim PAG_CreditCardToken
		dim PAG_ProofOfSale
		dim PAG_ErrorCode
		dim PAG_ErrorMessage
		dim PAG_faultcode
		dim PAG_faultstring
		end class
	
	
'	GetOrderIdData
'	==============
	class cl_BRASPAG_GetOrderIdData_TX
		dim PAG_Version
		dim PAG_RequestId
		dim PAG_MerchantId
		dim PAG_OrderId
		end class
	
	class cl_BRASPAG_GetOrderIdDataResponse_RX
		dim PAG_CorrelationId
		dim PAG_Success
		dim PAG_ErrorCode
		dim PAG_ErrorMessage
		dim PAG_faultcode
		dim PAG_faultstring
		end class
	
	class cl_BRASPAG_OrderIdDataCollection_RX
		dim blnHaDados
		dim PAG_BraspagOrderId
		dim PAG_BraspagTransactionId
		end class
	
	
'	VoidCreditCardTransaction
'	=========================
	class cl_BRASPAG_VoidCreditCardTransaction_TX
		dim PAG_Version
		dim PAG_RequestId
		dim PAG_MerchantId
		dim PAG_BraspagTransactionId
		dim PAG_Amount
		dim PAG_ServiceTaxAmount
		end class
	
	class cl_BRASPAG_VoidCreditCardTransactionResponse_RX
		dim PAG_CorrelationId
		dim PAG_Success
		dim PAG_BraspagTransactionId
		dim PAG_AcquirerTransactionId
		dim PAG_Amount
		dim PAG_AuthorizationCode
		dim PAG_ReturnCode
		dim PAG_ReturnMessage
		dim PAG_Status
		dim PAG_ProofOfSale
		dim PAG_ServiceTaxAmount
		dim PAG_ErrorCode
		dim PAG_ErrorMessage
		dim PAG_faultcode
		dim PAG_faultstring
		end class
	
	
'	RefundCreditCardTransaction
'	===========================
	class cl_BRASPAG_RefundCreditCardTransaction_TX
		dim PAG_Version
		dim PAG_RequestId
		dim PAG_MerchantId
		dim PAG_BraspagTransactionId
		dim PAG_Amount
		dim PAG_ServiceTaxAmount
		end class
	
	class cl_BRASPAG_RefundCreditCardTransactionResponse_RX
		dim PAG_CorrelationId
		dim PAG_Success
		dim PAG_BraspagTransactionId
		dim PAG_AcquirerTransactionId
		dim PAG_Amount
		dim PAG_AuthorizationCode
		dim PAG_ReturnCode
		dim PAG_ReturnMessage
		dim PAG_Status
		dim PAG_ProofOfSale
		dim PAG_ServiceTaxAmount
		dim PAG_ErrorCode
		dim PAG_ErrorMessage
		dim PAG_faultcode
		dim PAG_faultstring
		end class
	
'	CaptureCreditCardTransaction
'	============================
	class cl_BRASPAG_CaptureCreditCardTransaction_TX
		dim PAG_Version
		dim PAG_RequestId
		dim PAG_MerchantId
		dim PAG_BraspagTransactionId
		dim PAG_Amount
		dim PAG_ServiceTaxAmount
		end class
	
	class cl_BRASPAG_CaptureCreditCardTransactionResponse_RX
		dim PAG_CorrelationId
		dim PAG_Success
		dim PAG_BraspagTransactionId
		dim PAG_AcquirerTransactionId
		dim PAG_Amount
		dim PAG_AuthorizationCode
		dim PAG_ReturnCode
		dim PAG_ReturnMessage
		dim PAG_Status
		dim PAG_ProofOfSale
		dim PAG_ServiceTaxAmount
		dim PAG_ErrorCode
		dim PAG_ErrorMessage
		dim PAG_faultcode
		dim PAG_faultstring
		end class
	
'	FraudAnalysisTransactionDetails
'	===============================
	class cl_BRASPAG_FraudAnalysisTransactionDetails_TX
		dim AF_Version
		dim AF_RequestId
		dim AF_MerchantId
		dim AF_AntiFraudTransactionId
		end class
	
	class cl_BRASPAG_FraudAnalysisTransactionDetailsResponse_RX
		dim AF_CorrelatedId
		dim AF_Success
		dim AF_AntiFraudMerchantId
		dim AF_AntiFraudTransactionId
		dim AF_AntiFraudTransactionStatusCode
		dim AF_AntiFraudReceiveDate
		dim AF_AntiFraudStatusLastUpdateDate
		dim AF_AntiFraudAnalysisScore
		dim AF_BraspagTransactionId
		dim AF_MerchantOrderId
		dim AF_AntiFraudAcquirerConversionDate
		dim AF_AntiFraudTransactionOriginalStatusCode
		dim AF_ErrorCode
		dim AF_ErrorMessage
		dim AF_faultcode
		dim AF_faultstring
		end class
	
'	AF - UpdateStatus
'	=================
	class cl_BRASPAG_AF_UpdateStatus_TX
		dim AF_Version
		dim AF_RequestId
		dim AF_MerchantId
		dim AF_AntiFraudTransactionId
		dim AF_NewStatus
		dim AF_Comment
		end class
	
	class cl_BRASPAG_AF_UpdateStatusResponse_RX
		dim AF_CorrelatedId
		dim AF_Success
		dim AF_AntiFraudTransactionId
		dim AF_RequestStatusCode
		dim AF_RequestStatusDescription
		dim AF_ErrorCode
		dim AF_ErrorMessage
		dim AF_faultcode
		dim AF_faultstring
		end class







' =========================================
'          F  U  N  Ç  Õ  E  S
' =========================================

' ------------------------------------------------------------------------
'   xmlReadNode
function xmlReadNode(ByRef objXml, Byval node_path, Byref blnNodeNotFound)
dim oNode
	blnNodeNotFound = False

	set oNode = objXML.documentElement.selectSingleNode(node_path)
	if oNode is nothing then
		blnNodeNotFound = True
		xmlReadNode = ""
		exit function
		end if
	
	xmlReadNode = oNode.text
end function



' ------------------------------------------------------------------------
'   xmlReadSubNode
function xmlReadSubNode(ByRef objNode, Byval node_path, Byref blnNodeNotFound)
dim oNode
	blnNodeNotFound = False

	set oNode = objNode.selectSingleNode(node_path)
	if oNode is nothing then
		blnNodeNotFound = True
		xmlReadSubNode = ""
		exit function
		end if
	
	xmlReadSubNode = oNode.text
end function



' ------------------------------------------------------------------------
'   BraspagDescricaoParcelamento
'   Retorna a descrição para a forma de pagamento selecionada.
function BraspagDescricaoParcelamento(byval cod_produto, byval qtde_parcelas, byval valor_total)
dim s_resp
dim vl_parcela
dim vl_total

	cod_produto = Trim("" & cod_produto)
	vl_total = converte_numero(valor_total)
	if qtde_parcelas <> 0 then vl_parcela = vl_total / qtde_parcelas

	select case cod_produto
	'	CRÉDITO À VISTA
		case "0"
			s_resp = SIMBOLO_MONETARIO & " " & formata_moeda(valor_total) & " À Vista (no Crédito)"
	'	PARCELADO LOJA
		case "1"
			s_resp = formata_inteiro(qtde_parcelas) & "x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " iguais"
	'	PARCELADO ADMINISTRADORA
		case "2"
			s_resp = formata_inteiro(qtde_parcelas) & "x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " mais juros"
	'	DÉBITO
		case "A"
			s_resp = SIMBOLO_MONETARIO & " " & formata_moeda(valor_total) & " À Vista (no Débito)"
		case else
			s_resp = ""
	end select

	BraspagDescricaoParcelamento = s_resp
end function



' ------------------------------------------------------------------------
'   BraspagDescricaoBandeira
function BraspagDescricaoBandeira(Byval bandeira)
dim s_resp
	bandeira = Lcase(Trim("" & bandeira))
	if bandeira = "visa" then
		s_resp = "Visa"
	elseif bandeira = "mastercard" then
		s_resp = "Mastercard"
	elseif bandeira = "amex" then
		s_resp = "Amex"
	elseif bandeira = "elo" then
		s_resp = "Elo"
	elseif bandeira = "hipercard" then
		s_resp = "Hipercard"
	elseif bandeira = "diners" then
		s_resp = "Diners"
	elseif bandeira = "discover" then
		s_resp = "Discover"
	elseif bandeira = "aura" then
		s_resp = "Aura"
	elseif bandeira = "jcb" then
		s_resp = "JCB"
	elseif bandeira <> "" then
		s_resp = "Bandeira desconhecida (" & bandeira & ")"
	else
		s_resp = ""
		end if
		
	BraspagDescricaoBandeira = s_resp
end function



' ------------------------------------------------------------------------
'	BraspagObtemIdRegistroBdPrazoPagtoLoja
'	Dada a bandeira do cartão, retorna o ID do registro da tabela
'	t_PRAZO_PAGTO_VISANET que contém os dados do parcelamento pela loja.
function BraspagObtemIdRegistroBdPrazoPagtoLoja(Byval bandeira)
dim s_resp
	s_resp = ""
	bandeira = Ucase(Trim("" & bandeira))
	if bandeira = Ucase(BRASPAG_BANDEIRA__VISA) then
		s_resp = COD_VISANET_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__MASTERCARD) then
		s_resp = COD_MASTERCARD_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__AMEX) then
		s_resp = COD_AMEX_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__ELO) then
		s_resp = COD_ELO_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__HIPERCARD) then
		s_resp = COD_HIPERCARD_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__DINERS) then
		s_resp = COD_DINERS_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__DISCOVER) then
		s_resp = COD_DISCOVER_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__AURA) then
		s_resp = COD_AURA_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__JCB) then
		s_resp = COD_JCB_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__CELULAR) then
		s_resp = COD_CELULAR_PRAZO_PAGTO_LOJA
		end if
	BraspagObtemIdRegistroBdPrazoPagtoLoja = s_resp
end function



' ------------------------------------------------------------------------
'	BraspagObtemIdRegistroBdPrazoPagtoEmissor
'	Dada a bandeira do cartão, retorna o ID do registro da tabela
'	t_PRAZO_PAGTO_VISANET que contém os dados do parcelamento pelo
'	emissor do cartão.
function BraspagObtemIdRegistroBdPrazoPagtoEmissor(Byval bandeira)
dim s_resp
	s_resp = ""
	bandeira = Ucase(Trim("" & bandeira))
	if bandeira = Ucase(BRASPAG_BANDEIRA__VISA) then
		s_resp = COD_VISANET_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__MASTERCARD) then
		s_resp = COD_MASTERCARD_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__AMEX) then
		s_resp = COD_AMEX_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__ELO) then
		s_resp = COD_ELO_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__HIPERCARD) then
		s_resp = COD_HIPERCARD_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__DINERS) then
		s_resp = COD_DINERS_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__DISCOVER) then
		s_resp = COD_DISCOVER_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__AURA) then
		s_resp = COD_AURA_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__JCB) then
		s_resp = COD_JCB_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__CELULAR) then
		s_resp = COD_CELULAR_PRAZO_PAGTO_EMISSOR
		end if
	BraspagObtemIdRegistroBdPrazoPagtoEmissor = s_resp
end function



' ------------------------------------------------------------------------
'	BraspagObtemNomeArquivoLogo
'	Dada a bandeira do cartão, retorna o nome do arquivo que contém o
'	logotipo.
function BraspagObtemNomeArquivoLogo(Byval bandeira)
dim s_resp
	s_resp = ""
	bandeira = Ucase(Trim("" & bandeira))
	if bandeira = Ucase(BRASPAG_BANDEIRA__VISA) then
		s_resp = "LogoVisa.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__MASTERCARD) then
		s_resp = "mastercard.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__AMEX) then
		s_resp = "Amex.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__ELO) then
		s_resp = "Elo.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__HIPERCARD) then
		s_resp = "Hipercard.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__DINERS) then
		s_resp = "Diners.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__DISCOVER) then
		s_resp = "Discover.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__AURA) then
		s_resp = "Aura.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__JCB) then
		s_resp = "JCB.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__CELULAR) then
		s_resp = "Celular.gif"
	else
		s_resp = "Unknown.gif"
		end if
	BraspagObtemNomeArquivoLogo = s_resp
end function



' ------------------------------------------------------------------------
'	BraspagObtemNomeArquivoLogoOpcao
'	Dada a bandeira do cartão, retorna o nome do arquivo que contém o
'	logotipo.
function BraspagObtemNomeArquivoLogoOpcao(Byval bandeira)
dim s_resp
	s_resp = ""
	bandeira = Ucase(Trim("" & bandeira))
	if bandeira = Ucase(BRASPAG_BANDEIRA__VISA) then
		s_resp = "opt_LogoVisa.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__MASTERCARD) then
		s_resp = "opt_mastercard.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__AMEX) then
		s_resp = "opt_Amex.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__ELO) then
		s_resp = "opt_Elo.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__HIPERCARD) then
		s_resp = "opt_Hipercard.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__DINERS) then
		s_resp = "opt_Diners.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__DISCOVER) then
		s_resp = "opt_Discover.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__AURA) then
		s_resp = "opt_Aura.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__JCB) then
		s_resp = "opt_JCB.gif"
	elseif bandeira = Ucase(BRASPAG_BANDEIRA__CELULAR) then
		s_resp = "opt_Celular.gif"
	else
		s_resp = "Unknown.gif"
		end if
	BraspagObtemNomeArquivoLogoOpcao = s_resp
end function



' ------------------------------------------------------------------------
'	BraspagQtdeBandeirasHabilitadas
'	Calcula a quantidade de bandeiras ativas que estão disponíveis para
'	serem usadas nas transações.
function BraspagQtdeBandeirasHabilitadas(ByVal owner)
dim qtdeBandeiras
	qtdeBandeiras = 0
	
	if Cstr(owner) = Cstr(BRASPAG_OWNER_OLD01) then
		if BRASPAG_OLD01_BANDEIRA_HABILITADA__VISA then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD01_BANDEIRA_HABILITADA__MASTERCARD then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD01_BANDEIRA_HABILITADA__AMEX then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD01_BANDEIRA_HABILITADA__ELO then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD01_BANDEIRA_HABILITADA__HIPERCARD then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD01_BANDEIRA_HABILITADA__DINERS then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD01_BANDEIRA_HABILITADA__DISCOVER then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD01_BANDEIRA_HABILITADA__AURA then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD01_BANDEIRA_HABILITADA__JCB then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD01_BANDEIRA_HABILITADA__CELULAR then qtdeBandeiras = qtdeBandeiras + 1
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_OLD02) then
		if BRASPAG_OLD02_BANDEIRA_HABILITADA__VISA then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD02_BANDEIRA_HABILITADA__MASTERCARD then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD02_BANDEIRA_HABILITADA__AMEX then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD02_BANDEIRA_HABILITADA__ELO then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD02_BANDEIRA_HABILITADA__HIPERCARD then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD02_BANDEIRA_HABILITADA__DINERS then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD02_BANDEIRA_HABILITADA__DISCOVER then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD02_BANDEIRA_HABILITADA__AURA then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD02_BANDEIRA_HABILITADA__JCB then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_OLD02_BANDEIRA_HABILITADA__CELULAR then qtdeBandeiras = qtdeBandeiras + 1
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_DIS) then
		if BRASPAG_DIS_BANDEIRA_HABILITADA__VISA then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_DIS_BANDEIRA_HABILITADA__MASTERCARD then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_DIS_BANDEIRA_HABILITADA__AMEX then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_DIS_BANDEIRA_HABILITADA__ELO then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_DIS_BANDEIRA_HABILITADA__HIPERCARD then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_DIS_BANDEIRA_HABILITADA__DINERS then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_DIS_BANDEIRA_HABILITADA__DISCOVER then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_DIS_BANDEIRA_HABILITADA__AURA then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_DIS_BANDEIRA_HABILITADA__JCB then qtdeBandeiras = qtdeBandeiras + 1
		if BRASPAG_DIS_BANDEIRA_HABILITADA__CELULAR then qtdeBandeiras = qtdeBandeiras + 1
		end if
	
	BraspagQtdeBandeirasHabilitadas = qtdeBandeiras
end function



' ------------------------------------------------------------------------
'	BraspagArrayBandeiras
'	Cria e retorna um array contendo as bandeiras existentes, ou seja,
'	independentemente da bandeira estar habilitada ou não.
function BraspagArrayBandeiras
	BraspagArrayBandeiras = Array(BRASPAG_BANDEIRA__VISA, _
							BRASPAG_BANDEIRA__MASTERCARD, _
							BRASPAG_BANDEIRA__AMEX, _
							BRASPAG_BANDEIRA__ELO, _
							BRASPAG_BANDEIRA__HIPERCARD, _
							BRASPAG_BANDEIRA__DINERS, _
							BRASPAG_BANDEIRA__DISCOVER, _
							BRASPAG_BANDEIRA__AURA, _
							BRASPAG_BANDEIRA__JCB)
end function



' ------------------------------------------------------------------------
'	BraspagSelecaoBandeiraQtdePorLinha
'	Calcula quantas bandeiras devem ser exibidas por linha na tela de
'	escolha da bandeira a ser usada no pagamento.
function BraspagSelecaoBandeiraQtdePorLinha(ByVal owner)
dim qtdeBandeiras
dim qtdePorLinha
	qtdeBandeiras=BraspagQtdeBandeirasHabilitadas(owner)
	select case qtdeBandeiras
		case 1, 2, 3, 4
			qtdePorLinha = qtdeBandeiras
		case 5
			qtdePorLinha = 3	' L1 = 3, L2 = 2
		case 6
			qtdePorLinha = 3	' L1 = 3, L2 = 3
		case 7
			qtdePorLinha = 4	' L1 = 4, L2 = 3
		case 8
			qtdePorLinha = 4	' L1 = 4, L2 = 4
		case 9
			qtdePorLinha = 3	' L1 = 3, L2 = 3, L3 = 3
		case 10
			qtdePorLinha = 4	' L1 = 4, L2 = 4, L3 = 2
		case 11
			qtdePorLinha = 4	' L1 = 4, L2 = 4, L3 = 3
		case 12
			qtdePorLinha = 4	' L1 = 4, L2 = 4, L3 = 4
		case else
			qtdePorLinha = 4
	end select
	
	BraspagSelecaoBandeiraQtdePorLinha = qtdePorLinha
end function



' ------------------------------------------------------------------------
'	BraspagAntiFraudeDecodificaBandeira
'	Converte a codificação que identifica a bandeira no sistema da Artven
'	p/ a codificação usada no Anti Fraude da Braspag.
function BraspagAntiFraudeDecodificaBandeira(ByVal bandeira)
dim strResposta, strBandeira
	BraspagAntiFraudeDecodificaBandeira = ""
	strBandeira = UCase(Trim("" & bandeira))
	if strBandeira = UCase(BRASPAG_BANDEIRA__VISA) then
		strResposta = "Visa"
	elseif strBandeira = UCase(BRASPAG_BANDEIRA__MASTERCARD) then
		strResposta = "Mastercard"
	elseif strBandeira = UCase(BRASPAG_BANDEIRA__AMEX) then
		strResposta = "AmericanExpress"
	elseif strBandeira = UCase(BRASPAG_BANDEIRA__DINERS) then
		strResposta = "DinersClub"
	elseif strBandeira = UCase(BRASPAG_BANDEIRA__ELO) then
		strResposta = "Elo"
	elseif strBandeira = UCase(BRASPAG_BANDEIRA__HIPERCARD) then
		strResposta = "Hipercard"
	else
		strResposta = ""
		end if
	BraspagAntiFraudeDecodificaBandeira = strResposta
end function



' ------------------------------------------------------------------------
'   BraspagEnviaTransacaoComRetry
'   Método que executa o BraspagEnviaTransacao() dentro de um laço de tentativas até que a execução seja bem sucedida ou a quantidade máxima de tentativas seja atingida.
'   Importante: este método pode ser utilizado livremente para requisições de consulta, entretanto, para requisições que alteram dados é importante avaliar antes
'   as possíveis consequências que podem ocorrer no caso da requisição ter sido processada no web service e o erro ter ocorrido em algum estágio posterior durante
'   o recebimento da resposta. Nesse caso, o uso deste método pode causar múltiplas execuções da requisição.
function BraspagEnviaTransacaoComRetry(Byval xml, Byval WS_ENDERECO)
const MAX_TENTATIVAS = 3
dim qtdeTentativasRealizadas
dim xmlResp
dim err_number
dim err_description
dim blnErroTimeout

	On Error Resume Next
	
	qtdeTentativasRealizadas = 0
	do while qtdeTentativasRealizadas < MAX_TENTATIVAS
		qtdeTentativasRealizadas = qtdeTentativasRealizadas + 1
		
		Err.Clear
		xmlResp = BraspagEnviaTransacao(xml, WS_ENDERECO)
		if Err.number = 0 then
		'	EXECUÇÃO FOI BEM SUCEDIDA
			BraspagEnviaTransacaoComRetry = xmlResp
			exit do
		else
			err_number = Err.number
			err_description = Trim("" & Err.Description)
		'	SE OCORREU UM ERRO DE TIMEOUT, CONTINUA TENTANDO, CASO CONTRÁRIO, DESISTE
			blnErroTimeout = False
			if err_number = -2147012894 then blnErroTimeout = True
			if InStr(UCase(err_description), UCase(" timed out")) <> 0 then blnErroTimeout = True
			if InStr(UCase(err_description), UCase("tempo limite da operação foi atingido")) <> 0 then blnErroTimeout = True
			if Not blnErroTimeout then
				exit do
				end if
			end if
		loop

end function



' ------------------------------------------------------------------------
'   BraspagEnviaTransacao
'	Option: 2 = SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
'	The SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS option is a DWORD mask of various flags that can be set to change this default behavior.
'	The default value is to ignore all problems. You must set this option before calling the send method. The flags are as follows:
'		SXH_SERVER_CERT_IGNORE_UNKNOWN_CA = 256
'		Unknown certificate authority
'		SXH_SERVER_CERT_IGNORE_WRONG_USAGE = 512
'		Malformed certificate such as a certificate with no subject name.
'		SXH_SERVER_CERT_IGNORE_CERT_CN_INVALID = 4096
'		Mismatch between the visited hostname and the certificate name being used on the server.
'		SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID = 8192
'		The date in the certificate is invalid or has expired.
'		SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
'		All certificate errors.
'	To turn off a flag, you subtract it from the default value, which is the sum of all flags.
'	For example, to catch an invalid date in a certificate, you turn off the SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID flag as follows:
'	shx.setOption(2) = (shx.getOption(2) - SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID)
'
'	oServerXMLHTTPRequest.setTimeouts(resolveTimeout, connectTimeout, sendTimeout, receiveTimeout);
'	The timeout parameters of the setTimeouts method are specified in milliseconds
'		resolveTimeout: the default value is infinite
'		connectTimeout: default timeout value of 60 seconds
'		sendTimeout: default value is 30 seconds
'		receiveTimeout: default value is 30 seconds
function BraspagEnviaTransacao(Byval xml, Byval WS_ENDERECO)
dim xmlhttp
const RESOLVE_TIMEOUT_MS = 30000
const CONNECT_TIMEOUT_MS = 30000
const SEND_TIMEOUT_MS = 60000
const RECEIVE_TIMEOUT_MS = 180000

'	set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
	set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	xmlhttp.setTimeouts RESOLVE_TIMEOUT_MS, CONNECT_TIMEOUT_MS, SEND_TIMEOUT_MS, RECEIVE_TIMEOUT_MS
	xmlhttp.open "POST", WS_ENDERECO, False
	xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	xmlhttp.setOption 2, 13056
	xmlhttp.send xml
	BraspagEnviaTransacao = xmlhttp.responseText
end function



' ------------------------------------------------------------------------
'   cria_instancia_cl_BRASPAG_AnalyseAndAuthorizeOnSuccess_TX
function cria_instancia_cl_BRASPAG_AnalyseAndAuthorizeOnSuccess_TX(Byval owner, Byval bandeira)
dim trx
	set trx = new cl_BRASPAG_AnalyseAndAuthorizeOnSuccess_TX
	
	bandeira = Lcase(Trim("" & bandeira))
	
	trx.AF_Version = BRASPAG_ANTIFRAUDE_VERSION
	trx.AF_AntiFraudSequenceType = "AnalyseAndAuthorizeOnSuccess"
	
	trx.AF_MerchantId = BraspagObtem_AF_MERCHANT_ID(owner)
	trx.PAG_OrderData_MerchantId = BraspagObtem_PAG_MERCHANT_ID(owner)
	
	trx.AF_AntiFraudRequest_PurchaseTotalsData_Currency = "BRL"
	
	trx.PAG_Version = BRASPAG_PAGADOR_VERSION
	trx.PAG_PaymentDataCollection_Currency = "BRL"
	trx.PAG_PaymentDataCollection_Country = "BRA"
	trx.PAG_PaymentDataCollection_TransactionType = "2" '1=Pré-Autorização / 2=Captura Automática
	
	'Observações:
	'	1) O meio de pagamento 599 (Getnet WebService) independe da bandeira, basta a bandeira estar habilitada na plataforma da adquirente (Getnet WebService)
	'	2) O meio de pagamento 612 (SafraPay) independe da bandeira, basta a bandeira estar habilitada na plataforma da adquirente (SafraPay)
	if (bandeira = Lcase(BRASPAG_BANDEIRA__VISA)) then
		trx.PAG_PaymentDataCollection_PaymentMethod = "612" 'Cielo VISA = 500 / SiTef Santander VISA (Getnet) = 531 / Getnet WebService = 599 / SafraPay = 612
	elseif (bandeira = Lcase(BRASPAG_BANDEIRA__MASTERCARD)) then
		trx.PAG_PaymentDataCollection_PaymentMethod = "612" 'Cielo MASTERCARD = 501 / SiTef Santander MASTERCARD (Getnet) = 532 / Getnet WebService = 599 / SafraPay = 612
	elseif (bandeira = Lcase(BRASPAG_BANDEIRA__AMEX)) then
		trx.PAG_PaymentDataCollection_PaymentMethod = "612" 'Cielo AMEX = 502 / Getnet WebService = 599 / SafraPay = 612
	elseif (bandeira = Lcase(BRASPAG_BANDEIRA__ELO)) then
		trx.PAG_PaymentDataCollection_PaymentMethod = "612" 'Cielo ELO = 504 / Getnet WebService = 599 / SafraPay = 612
	elseif (bandeira = Lcase(BRASPAG_BANDEIRA__HIPERCARD)) then
		trx.PAG_PaymentDataCollection_PaymentMethod = "612" 'Cielo ELO = 504 / Getnet WebService = 599 / SafraPay = 612
	elseif (bandeira = Lcase(BRASPAG_BANDEIRA__DINERS)) then
		trx.PAG_PaymentDataCollection_PaymentMethod = "503"
	elseif (bandeira = Lcase(BRASPAG_BANDEIRA__DISCOVER)) then
		trx.PAG_PaymentDataCollection_PaymentMethod = "543"
	elseif (bandeira = Lcase(BRASPAG_BANDEIRA__AURA)) then
		trx.PAG_PaymentDataCollection_PaymentMethod = "545"
	elseif (bandeira = Lcase(BRASPAG_BANDEIRA__JCB)) then
		trx.PAG_PaymentDataCollection_PaymentMethod = "544"
		end if
	
	if BRASPAG_AMBIENTE_HOMOLOGACAO then trx.PAG_PaymentDataCollection_PaymentMethod = "997"
	
	set cria_instancia_cl_BRASPAG_AnalyseAndAuthorizeOnSuccess_TX = trx
end function



' ------------------------------------------------------------------------
'   cria_instancia_cl_BRASPAG_Authorize_TX
function cria_instancia_cl_BRASPAG_Authorize_TX(Byval owner)
dim trx
	set trx = new cl_BRASPAG_Authorize_TX
	
	trx.OrderData_MerchantId = BraspagObtem_PAG_MERCHANT_ID(owner)
	trx.Version = BRASPAG_PAGADOR_VERSION

	set cria_instancia_cl_BRASPAG_Authorize_TX = trx
end function



' ------------------------------------------------------------------------
'   cria_instancia_cl_BRASPAG_Authorize_PaymentDataRequest_TX
function cria_instancia_cl_BRASPAG_Authorize_PaymentDataRequest_TX(Byval owner, Byref vDadosCartao, Byref v_trx)
dim i, bandeira
	redim v_trx(0)
	set v_trx(Ubound(v_trx)) = new cl_BRASPAG_Authorize_PaymentDataRequest_TX

	for i=Lbound(vDadosCartao) to Ubound(vDadosCartao)
		if Not IsEmpty(vDadosCartao(i)) then
			bandeira = Lcase(Trim("" & vDadosCartao(i).bandeira))

			if Trim("" & v_trx(Ubound(v_trx)).PAG_PaymentMethod) <> "" then
				redim preserve v_trx(Ubound(v_trx)+1)
				set v_trx(Ubound(v_trx)) = new cl_BRASPAG_Authorize_PaymentDataRequest_TX
				end if
			
		'	BANDEIRA
			v_trx(Ubound(v_trx)).bandeira = bandeira

		'	PAYMENT METHOD
			'Observações: 
			'	1) O meio de pagamento 599 (Getnet WebService) independe da bandeira, basta a bandeira estar habilitada na plataforma da adquirente (Getnet WebService)
			'	2) O meio de pagamento 612 (SafraPay) independe da bandeira, basta a bandeira estar habilitada na plataforma da adquirente (SafraPay)
			if (bandeira = Lcase(BRASPAG_BANDEIRA__VISA)) then
				v_trx(Ubound(v_trx)).PAG_PaymentMethod = "612" 'Cielo VISA = 500 / SiTef Santander VISA (Getnet) = 531 / Getnet WebService = 599 / SafraPay = 612
			elseif (bandeira = Lcase(BRASPAG_BANDEIRA__MASTERCARD)) then
				v_trx(Ubound(v_trx)).PAG_PaymentMethod = "612" 'Cielo MASTERCARD = 501 / SiTef Santander MASTERCARD (Getnet) = 532 / Getnet WebService = 599 / SafraPay = 612
			elseif (bandeira = Lcase(BRASPAG_BANDEIRA__AMEX)) then
				v_trx(Ubound(v_trx)).PAG_PaymentMethod = "612" 'Cielo AMEX = 502 / Getnet WebService = 599 / SafraPay = 612
			elseif (bandeira = Lcase(BRASPAG_BANDEIRA__ELO)) then
				v_trx(Ubound(v_trx)).PAG_PaymentMethod = "612" 'Cielo ELO = 504 / Getnet WebService = 599 / SafraPay = 612
			elseif (bandeira = Lcase(BRASPAG_BANDEIRA__HIPERCARD)) then
				v_trx(Ubound(v_trx)).PAG_PaymentMethod = "612" 'Cielo ELO = 504 / Getnet WebService = 599 / SafraPay = 612
			elseif (bandeira = Lcase(BRASPAG_BANDEIRA__DINERS)) then
				v_trx(Ubound(v_trx)).PAG_PaymentMethod = "503"
			elseif (bandeira = Lcase(BRASPAG_BANDEIRA__DISCOVER)) then
				v_trx(Ubound(v_trx)).PAG_PaymentMethod = "543"
			elseif (bandeira = Lcase(BRASPAG_BANDEIRA__AURA)) then
				v_trx(Ubound(v_trx)).PAG_PaymentMethod = "545"
			elseif (bandeira = Lcase(BRASPAG_BANDEIRA__JCB)) then
				v_trx(Ubound(v_trx)).PAG_PaymentMethod = "544"
				end if
	
			if BRASPAG_AMBIENTE_HOMOLOGACAO then v_trx(Ubound(v_trx)).PAG_PaymentMethod = "997"
		
		'	AMOUNT
			v_trx(Ubound(v_trx)).PAG_Amount = retorna_so_digitos(vDadosCartao(i).valor_pagamento)
		'	CURRENCY
			v_trx(Ubound(v_trx)).PAG_Currency = "BRL"
		'	COUNTRY
			v_trx(Ubound(v_trx)).PAG_Country = "BRA"
		'	SERVICE TAX AMOUNT
			v_trx(Ubound(v_trx)).PAG_ServiceTaxAmount = ""
		'	NUMBER OF PAYMENTS
			v_trx(Ubound(v_trx)).PAG_NumberOfPayments = vDadosCartao(i).qtde_parcelas
		'	PAYMENT PLAN
			v_trx(Ubound(v_trx)).PAG_PaymentPlan = vDadosCartao(i).codigo_produto
		'	TRANSACTION TYPE
			v_trx(Ubound(v_trx)).PAG_TransactionType = "1" 'Pré-Autorização
		'	CARD HOLDER
			v_trx(Ubound(v_trx)).PAG_CardHolder = substitui_caracteres(vDadosCartao(i).titular_nome, "&", " E ")
		'	CARD NUMBER
			v_trx(Ubound(v_trx)).PAG_CardNumber = vDadosCartao(i).cartao_numero
		'	CARD SECURITY CODE
			v_trx(Ubound(v_trx)).PAG_CardSecurityCode = vDadosCartao(i).cartao_codigo_seguranca
		'	CARD EXPIRATION DATE
			if (vDadosCartao(i).cartao_validade_mes <> "") And (vDadosCartao(i).cartao_validade_ano <> "") then
				v_trx(Ubound(v_trx)).PAG_CardExpirationDate = vDadosCartao(i).cartao_validade_mes & "/" & vDadosCartao(i).cartao_validade_ano
			else
				v_trx(Ubound(v_trx)).PAG_CardExpirationDate = ""
				end if
			end if
		next
end function



' ------------------------------------------------------------------------
'   BraspagPagadorDescricaoPaymentDataResponseStatus
function BraspagPagadorDescricaoPaymentDataResponseStatus(Byval codigoStatus)
dim s_resp
	codigoStatus = Trim("" & codigoStatus)
		
	select case codigoStatus
		case BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__CAPTURADA
			s_resp = "Capturada"
		case BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__AUTORIZADA
			s_resp = "Autorizada"
		case BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__NAO_AUTORIZADA
			s_resp = "Não Autorizada"
		case BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__ERRO_DESQUALIFICANTE
			s_resp = "Transação Com Erro Desqualificante"
		case BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__AGUARDANDO_RESPOSTA
			s_resp = "Transação Aguardando Resposta"
		case ""
			s_resp = ""
		case else
			s_resp = "Código Desconhecido: " & codigoStatus
	end select
	
	BraspagPagadorDescricaoPaymentDataResponseStatus = s_resp
end function



' ------------------------------------------------------------------------
'   BraspagPagadorDescricaoGlobalStatus
function BraspagPagadorDescricaoGlobalStatus(Byval codigoGlobalStatus)
dim s_resp
	codigoGlobalStatus = Trim("" & codigoGlobalStatus)
		
	select case codigoGlobalStatus
		case BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__INDEFINIDA
			s_resp = "Status Indefinido"
		case BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA
			s_resp = "Capturada"
		case BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA
			s_resp = "Autorizada"
		case BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__NAO_AUTORIZADA
			s_resp = "Não Autorizada"
		case BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURA_CANCELADA
			s_resp = "Captura Cancelada"
		case BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNADA
			s_resp = "Estornada"
		case BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNO_PENDENTE
			s_resp = "Estorno Pendente"
		case BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AGUARDANDO_RESPOSTA
			s_resp = "Transação Aguardando Resposta"
		case BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ERRO_DESQUALIFICANTE
			s_resp = "Transação Com Erro Desqualificante"
		case ""
			s_resp = ""
		case else
			s_resp = "Código Desconhecido: " & codigoGlobalStatus
	end select
	
	BraspagPagadorDescricaoGlobalStatus = s_resp
end function



' ------------------------------------------------------------------------
'   BraspagAntiFraudeDescricaoFraudAnalysisResponseTransactionStatusCode
function BraspagAntiFraudeDescricaoFraudAnalysisResponseTransactionStatusCode(Byval codigoStatus)
dim s_resp
	codigoStatus = Trim("" & codigoStatus)
		
	select case codigoStatus
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__STARTED
			s_resp = "Started"
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__ACCEPT
			s_resp = "Accept"
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__REVIEW
			s_resp = "Review"
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__REJECT
			s_resp = "Reject"
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__PENDENT
			s_resp = "Pendent"
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__UNFINISHED
			s_resp = "Unfinished"
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__ABORTED
			s_resp = "Aborted"
		case ""
			s_resp = ""
		case else
			s_resp = "Código Desconhecido: " & codigoStatus
	end select
	
	BraspagAntiFraudeDescricaoFraudAnalysisResponseTransactionStatusCode = s_resp
end function



' ------------------------------------------------------------------------
'   BraspagAntiFraudeDescricaoGlobalStatus
function BraspagAntiFraudeDescricaoGlobalStatus(Byval codigoGlobalStatus)
dim s_resp
	codigoGlobalStatus = Trim("" & codigoGlobalStatus)
		
	select case codigoGlobalStatus
		case BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__STARTED
			s_resp = "Started"
		case BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ACCEPT
			s_resp = "Accept"
		case BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW
			s_resp = "Review"
		case BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REJECT
			s_resp = "Reject"
		case BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__PENDENT
			s_resp = "Pendent"
		case BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__UNFINISHED
			s_resp = "Unfinished"
		case BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ABORTED
			s_resp = "Aborted"
		case ""
			s_resp = ""
		case else
			s_resp = "Código Desconhecido: " & codigoGlobalStatus
	end select
	
	BraspagAntiFraudeDescricaoGlobalStatus = s_resp
end function



' ------------------------------------------------------------------------
'   BraspagXmlMontaRequisicaoAuthorize
function BraspagXmlMontaRequisicaoAuthorize(ByRef trx, ByRef v_trx_payment, ByRef xmlMontadoMasked)
dim xml, xml_aux, xmlPayment
dim i, iTab

	xmlMontadoMasked = ""

	if Trim("" & trx.RequestId) = "" then trx.RequestId = Lcase(gera_uid)
	
	xml =	"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & chr(13) & _
			"	<soap:Body>" & chr(13) & _
			"		<AuthorizeTransaction xmlns=""https://www.pagador.com.br/webservice/pagador"">" & chr(13) & _
			"			<request>" & chr(13) & _
							xml_monta_campo(trx.RequestId, "RequestId", 4) & _
							xml_monta_campo(trx.Version, "Version", 4)
	
	xml = xml & _
			"				<OrderData>" & chr(13) & _
								xml_monta_campo(trx.OrderData_MerchantId, "MerchantId", 5) & _
								xml_monta_campo(trx.OrderData_OrderId, "OrderId", 5) & _
			"				</OrderData>" & chr(13)
	
	xml = xml & _
			"				<CustomerData>" & chr(13) & _
								xml_monta_campo(trx.CustomerData_CustomerIdentity, "CustomerIdentity", 5) & _
								xml_monta_campo(trx.CustomerData_CustomerIdentityType, "CustomerIdentityType", 5) & _
								xml_monta_campo(trx.CustomerData_CustomerName, "CustomerName", 5) & _
								xml_monta_campo(trx.CustomerData_CustomerEmail, "CustomerEmail", 5) & _
			"				</CustomerData>" & chr(13)
	
	xml = xml & _
			"				<PaymentDataCollection>" & chr(13)
	
	xmlMontadoMasked = xml

	iTab = 6
	for i=Lbound(v_trx_payment) to Ubound(v_trx_payment)
		xml_aux = _
			"					<PaymentDataRequest xsi:type=""CreditCardDataRequest"">" & chr(13) & _
									xml_monta_campo(v_trx_payment(i).PAG_PaymentMethod, "PaymentMethod", iTab) & _
									xml_monta_campo(v_trx_payment(i).PAG_Amount, "Amount", iTab) & _
									xml_monta_campo(v_trx_payment(i).PAG_Currency, "Currency", iTab) & _
									xml_monta_campo(v_trx_payment(i).PAG_Country, "Country", iTab) & _
									xml_monta_campo(v_trx_payment(i).PAG_ServiceTaxAmount, "ServiceTaxAmount", iTab) & _
									xml_monta_campo(v_trx_payment(i).PAG_NumberOfPayments, "NumberOfPayments", iTab) & _
									xml_monta_campo(v_trx_payment(i).PAG_PaymentPlan, "PaymentPlan", iTab) & _
									xml_monta_campo(v_trx_payment(i).PAG_TransactionType, "TransactionType", iTab) & _
									xml_monta_campo(v_trx_payment(i).PAG_CardHolder, "CardHolder", iTab)

		xml = xml & xml_aux
		xmlMontadoMasked = xmlMontadoMasked & xml_aux

		xml = xml & _
									xml_monta_campo(v_trx_payment(i).PAG_CardNumber, "CardNumber", iTab) & _
									xml_monta_campo(v_trx_payment(i).PAG_CardSecurityCode, "CardSecurityCode", iTab)

		xmlMontadoMasked = xmlMontadoMasked & _
									xml_monta_campo(BraspagCSProtegeNumeroCartao(v_trx_payment(i).PAG_CardNumber), "CardNumber", iTab) & _
									xml_monta_campo(String(Len(v_trx_payment(i).PAG_CardSecurityCode),"*"), "CardSecurityCode", iTab)

		xml_aux = _
									xml_monta_campo(v_trx_payment(i).PAG_CardExpirationDate, "CardExpirationDate", iTab) & _
		"						</PaymentDataRequest>" & chr(13)

		xml = xml & xml_aux
		xmlMontadoMasked = xmlMontadoMasked & xml_aux
		next

	xml_aux = _
			"				</PaymentDataCollection>" & chr(13) & _
			"			</request>" & chr(13) & _
			"		</AuthorizeTransaction>" & chr(13) & _
			"	</soap:Body>" & chr(13) & _
			"</soap:Envelope>" & chr(13)
	
	xml = xml & xml_aux
	xmlMontadoMasked = xmlMontadoMasked & xml_aux

	BraspagXmlMontaRequisicaoAuthorize = xml
end function



' ------------------------------------------------------------------------
'   BraspagXmlMontaRequisicaoAnalyseAndAuthorizeOnSuccess
function BraspagXmlMontaRequisicaoAnalyseAndAuthorizeOnSuccess(ByRef trx, ByRef vItem, ByRef vAdditionalData)
dim xml, xmlItemData, xmlAdditionalData
dim strAF_DocumentData_CnpjCpf
dim i

	if Trim("" & trx.AF_RequestId) = "" then trx.AF_RequestId = Lcase(gera_uid)
	if Trim("" & trx.PAG_RequestId) = "" then trx.PAG_RequestId = Lcase(gera_uid)
	
	if Trim(trx.AF_DocumentData_Cpf) <> "" then
		strAF_DocumentData_CnpjCpf = xml_monta_campo(trx.AF_DocumentData_Cpf, "ant:Cpf", 5)
	elseif Trim(trx.AF_DocumentData_Cnpj) <> "" then
		strAF_DocumentData_CnpjCpf = xml_monta_campo(trx.AF_DocumentData_Cnpj, "ant:Cnpj", 5)
		end if
	
	xml =	"<soapenv:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ant=""http://www.braspag.com.br/antifraud/"" xmlns:pag=""https://www.pagador.com.br/webservice/pagador"">" & chr(13) & _
			"	<soapenv:Header/>" & chr(13) & _
			"	<soapenv:Body>" & chr(13) & _
			"		<ant:FraudAnalysis>" & chr(13) & _
			"			<ant:request>" & chr(13) & _
							xml_monta_campo(trx.AF_RequestId, "ant:RequestId", 4) & _
							xml_monta_campo(trx.AF_Version, "ant:Version", 4) & _
							xml_monta_campo("AnalyseAndAuthorizeOnSuccess", "ant:AntiFraudSequenceType", 4) & _
			"				<ant:DocumentData>" & chr(13) & _
								strAF_DocumentData_CnpjCpf & _
			"				</ant:DocumentData>" & chr(13) & _
			"				<ant:AntiFraudRequest>" & chr(13)
	
	xml = xml & _
			"					<ant:BillToData>" & chr(13) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_CustomerId, "ant:CustomerId", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_City, "ant:City", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_Country, "ant:Country", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_Email, "ant:Email", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_FirstName, "ant:FirstName", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_LastName, "ant:LastName", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_State, "ant:State", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_Street1, "ant:Street1", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_Street2, "ant:Street2", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_PostalCode, "ant:PostalCode", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_PhoneNumber, "ant:PhoneNumber", 6)

	if Trim(trx.AF_AntiFraudRequest_BillToData_IpAddress) <> "" then
		xml = xml & _
									xml_monta_campo(trx.AF_AntiFraudRequest_BillToData_IpAddress, "ant:IpAddress", 6)
		end if
	
	xml = xml & _
			"					</ant:BillToData>" & chr(13)
	
	xml = xml & _
			"					<ant:ShipToData>" & chr(13) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_ShipToData_City, "ant:City", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_ShipToData_Country, "ant:Country", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_ShipToData_FirstName, "ant:FirstName", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_ShipToData_LastName, "ant:LastName", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_ShipToData_PhoneNumber, "ant:PhoneNumber", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_ShipToData_PostalCode, "ant:PostalCode", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_ShipToData_ShippingMethod, "ant:ShippingMethod", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_ShipToData_State, "ant:State", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_ShipToData_Street1, "ant:Street1", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_ShipToData_Street2, "ant:Street2", 6) & _
			"					</ant:ShipToData> " & chr(13)
	
	if Trim(trx.AF_AntiFraudRequest_DeviceFingerPrintId) <> "" then
		xml = xml & _
									xml_monta_campo(trx.AF_AntiFraudRequest_DeviceFingerPrintId, "ant:DeviceFingerPrintId", 5)
		end if
	
	xml = xml & _
			"					<ant:CardData>" & chr(13) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_CardData_AccountNumber, "ant:AccountNumber", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_CardData_Card, "ant:Card", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_CardData_ExpirationMonth, "ant:ExpirationMonth", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_CardData_ExpirationYear, "ant:ExpirationYear", 6) & _
			"					</ant:CardData>" & chr(13)
	
	xmlItemData = ""
	for i=LBound(vItem) to UBound(vItem)
		if Trim("" & vItem(i).ProductData_Sku) <> "" then
			xmlItemData = xmlItemData & _
				"					<ant:ItemData>" & chr(13) & _
				"						<ant:ProductData>" & chr(13) & _
											xml_monta_campo(vItem(i).ProductData_Name, "ant:Name", 7) & _
											xml_monta_campo(vItem(i).ProductData_Sku, "ant:Sku", 7) & _
											xml_monta_campo(vItem(i).ProductData_Quantity, "ant:Quantity", 7) & _
											xml_monta_campo(vItem(i).ProductData_UnitPrice, "ant:UnitPrice", 7) & _
				"						</ant:ProductData>" & chr(13) & _
				"					</ant:ItemData>" & chr(13)
			end if
		next
	
	if xmlItemData <> "" then
		xml = xml & _
				"					<ant:ItemDataCollection>" & chr(13) & _
										xmlItemData & _
				"					</ant:ItemDataCollection>" & chr(13)
		end if
	
'	O AdditionalDataCollection é usado para enviar os dados de MerchantDefinedData!!
'	Se o parâmetro não estiver preenchido, não deve ser enviado!!
	xmlAdditionalData = ""
	for i=LBound(vAdditionalData) to UBound(vAdditionalData)
		if (Trim(vAdditionalData(i).Id) <> "") And (Trim(vAdditionalData(i).Value) <> "") then
			xmlAdditionalData = xmlAdditionalData & _
				"						<ant:AdditionalData>" & chr(13) & _
											xml_monta_campo(vAdditionalData(i).Id, "ant:Id", 7) & _
											xml_monta_campo(vAdditionalData(i).Value, "ant:Value", 7) & _
				"						</ant:AdditionalData>" & chr(13)
			end if
		next
	
	if xmlAdditionalData <> "" then
		xml = xml & _
				"					<ant:AdditionalDataCollection>" & chr(13) & _
										xmlAdditionalData & _
				"					</ant:AdditionalDataCollection>" & chr(13)
		end if
	
	xml = xml & _
			"					<ant:PurchaseTotalsData>" & chr(13) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_PurchaseTotalsData_Currency, "ant:Currency", 6) & _
									xml_monta_campo(trx.AF_AntiFraudRequest_PurchaseTotalsData_GrandTotalAmount, "ant:GrandTotalAmount", 6) & _
			"					</ant:PurchaseTotalsData>" & chr(13)
	
	xml = xml & _
								xml_monta_campo(trx.AF_AntiFraudRequest_MerchantReferenceCode, "ant:MerchantReferenceCode", 5) & _
			"				</ant:AntiFraudRequest>" & chr(13) & _
							xml_monta_campo(trx.AF_MerchantId, "ant:MerchantId", 4)
	
	xml = xml & _
		"				<ant:AuthorizeTransactionRequest xmlns=""https://www.pagador.com.br/webservice/pagador"">" & chr(13)
	
	xml = xml & _
									xml_monta_campo(trx.PAG_RequestId, "pag:RequestId", 5) & _
									xml_monta_campo(trx.PAG_Version, "pag:Version", 5) & _
		"					<pag:OrderData>" & chr(13) & _
										xml_monta_campo(trx.PAG_OrderData_MerchantId, "pag:MerchantId", 6) & _
										xml_monta_campo(trx.PAG_OrderData_OrderId, "pag:OrderId", 6) & _
		"					</pag:OrderData>" & chr(13)
	
	xml = xml & _
		"					<pag:CustomerData>" & chr(13) & _
								xml_monta_campo(trx.PAG_CustomerData_CustomerIdentity, "pag:CustomerIdentity", 6) & _
								xml_monta_campo(trx.PAG_CustomerData_CustomerName, "pag:CustomerName", 6) & _
		"					</pag:CustomerData>" & chr(13)
	
	xml = xml & _
		"					<pag:PaymentDataCollection>" & chr(13) & _
		"						<pag:PaymentDataRequest xsi:type=""CreditCardDataRequest"">" & chr(13) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_PaymentMethod, "pag:PaymentMethod", 7) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_Amount, "pag:Amount", 7) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_Currency, "pag:Currency", 7) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_Country, "pag:Country", 7) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_ServiceTaxAmount, "pag:ServiceTaxAmount", 7) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_NumberOfPayments, "pag:NumberOfPayments", 7) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_PaymentPlan, "pag:PaymentPlan", 7) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_TransactionType, "pag:TransactionType", 7) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_CardHolder, "pag:CardHolder", 7) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_CardNumber, "pag:CardNumber", 7) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_CardSecurityCode, "pag:CardSecurityCode", 7) & _
									xml_monta_campo(trx.PAG_PaymentDataCollection_CardExpirationDate, "pag:CardExpirationDate", 7) & _
		"						</pag:PaymentDataRequest>" & chr(13) & _
		"					</pag:PaymentDataCollection>" & chr(13)
	
	xml = xml & _
		"				</ant:AuthorizeTransactionRequest>" & chr(13) & _
		"			</ant:request>" & chr(13) & _
		"		</ant:FraudAnalysis>" & chr(13) & _
		"	</soapenv:Body>" & chr(13) & _
		"</soapenv:Envelope>" & chr(13)
	
	BraspagXmlMontaRequisicaoAnalyseAndAuthorizeOnSuccess = xml
end function



' ------------------------------------------------------------------------
'   BraspagXmlMontaRequisicaoGetOrderIdData
function BraspagXmlMontaRequisicaoGetOrderIdData(ByRef trx)
dim xml
	xml =	"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & chr(13) & _
			"	<soap:Body>" & chr(13) & _
			"		<GetOrderIdData xmlns=""https://www.pagador.com.br/query/pagadorquery"">" & chr(13) & _
			"			<orderIdDataRequest>" & chr(13) & _
			"				<Version>" & trx.PAG_Version & "</Version>" & chr(13) & _
			"				<RequestId>" & trx.PAG_RequestId & "</RequestId>" & chr(13) & _
			"				<MerchantId>" & trx.PAG_MerchantId & "</MerchantId>" & chr(13) & _
			"				<OrderId>" & trx.PAG_OrderId & "</OrderId>" & chr(13) & _
			"			</orderIdDataRequest>" & chr(13) & _
			"		</GetOrderIdData>" & chr(13) & _
			"	</soap:Body>" & chr(13) & _
			"</soap:Envelope>" & chr(13)
	BraspagXmlMontaRequisicaoGetOrderIdData = xml
end function



' ------------------------------------------------------------------------
'   BraspagXmlMontaRequisicaoGetTransactionData
function BraspagXmlMontaRequisicaoGetTransactionData(ByRef trx)
dim xml
	xml =	"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & chr(13) & _
			"	<soap:Body>" & chr(13) & _
			"		<GetTransactionData xmlns=""https://www.pagador.com.br/query/pagadorquery"">" & chr(13) & _
			"			<transactionDataRequest>" & chr(13) & _
			"				<Version>" & trx.PAG_Version & "</Version>" & chr(13) & _
			"				<RequestId>" & trx.PAG_RequestId & "</RequestId>" & chr(13) & _
			"				<MerchantId>" & trx.PAG_MerchantId & "</MerchantId>" & chr(13) & _
			"				<BraspagTransactionId>" & trx.PAG_BraspagTransactionId & "</BraspagTransactionId>" & chr(13) & _
			"			</transactionDataRequest>" & chr(13) & _
			"		</GetTransactionData>" & chr(13) & _
			"	</soap:Body>" & chr(13) & _
			"</soap:Envelope>" & chr(13)
	BraspagXmlMontaRequisicaoGetTransactionData = xml
end function



' ------------------------------------------------------------------------
'   BraspagXmlMontaRequisicaoFraudAnalysisTransactionDetails
function BraspagXmlMontaRequisicaoFraudAnalysisTransactionDetails(ByRef trx)
dim xml
	xml =	"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & chr(13) & _
			"	<soap:Body>" & chr(13) & _
			"		<FraudAnalysisTransactionDetails xmlns=""http://www.braspag.com.br/antifraud/"">" & chr(13) & _
			"			<request>" & chr(13) & _
			"				<Version>" & trx.AF_Version & "</Version>" & chr(13) & _
			"				<RequestId>" & trx.AF_RequestId & "</RequestId>" & chr(13) & _
			"				<MerchantId>" & trx.AF_MerchantId & "</MerchantId>" & chr(13) & _
			"				<AntiFraudTransactionId>" & trx.AF_AntiFraudTransactionId & "</AntiFraudTransactionId>" & chr(13) & _
			"			</request>" & chr(13) & _
			"		</FraudAnalysisTransactionDetails>" & chr(13) & _
			"	</soap:Body>" & chr(13) & _
			"</soap:Envelope>" & chr(13)
	BraspagXmlMontaRequisicaoFraudAnalysisTransactionDetails = xml
end function



' ------------------------------------------------------------------------
'   BraspagXmlMontaRequisicaoAfUpdateStatus
function BraspagXmlMontaRequisicaoAfUpdateStatus(ByRef trx)
dim xml
	xml =	"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & chr(13) & _
			"	<soap:Body>" & chr(13) & _
			"		<UpdateStatus xmlns=""http://www.braspag.com.br/antifraud/"">" & chr(13) & _
			"			<updateStatusRequest>" & chr(13) & _
			"				<Version>" & trx.AF_Version & "</Version>" & chr(13) & _
			"				<RequestId>" & trx.AF_RequestId & "</RequestId>" & chr(13) & _
			"				<MerchantId>" & trx.AF_MerchantId & "</MerchantId>" & chr(13) & _
			"				<AntiFraudTransactionId>" & trx.AF_AntiFraudTransactionId & "</AntiFraudTransactionId>" & chr(13) & _
			"				<NewStatus>" & trx.AF_NewStatus & "</NewStatus>" & chr(13)
	if Trim(trx.AF_Comment) <> "" then
		xml = xml & _
			"				<Comment>" & trx.AF_Comment & "</Comment>" & chr(13)
		end if
	xml = xml & _
			"			</updateStatusRequest>" & chr(13) & _
			"		</UpdateStatus>" & chr(13) & _
			"	</soap:Body>" & chr(13) & _
			"</soap:Envelope>" & chr(13)
	BraspagXmlMontaRequisicaoAfUpdateStatus = xml
end function



' ------------------------------------------------------------------------
'   BraspagFormataNumero2Dec
function BraspagFormataNumero2Dec(ByVal numero)
dim strSeparadorDecimal
dim strValorFormatado
dim i
dim c
dim s
	strSeparadorDecimal = ""
	s = CStr(0.5)
	For i = Len(s) To 1 Step -1
		c = Mid(s, i, 1)
		If Not IsNumeric(c) Then
			strSeparadorDecimal = c
			Exit For
			End If
		Next

	If strSeparadorDecimal = "" Then strSeparadorDecimal = ","
	
'	FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
'	Lembrando que IncludeLeadingDigit indica se valores como .5 serão exibidos como .5 ou 0.5
'	A função FormatCurrency sempre inclui o símbolo monetário.
	strValorFormatado = FormatNumber(numero, 2, -1, 0, 0)
	strValorFormatado = substitui_caracteres(strValorFormatado, strSeparadorDecimal, "V")
	strValorFormatado = substitui_caracteres(strValorFormatado, ".", "")
	strValorFormatado = substitui_caracteres(strValorFormatado, ",", "")
	strValorFormatado = substitui_caracteres(strValorFormatado, "V", ".")
	
	BraspagFormataNumero2Dec = strValorFormatado
end function



' --------------------------------------------------------------------------------
'   BraspagRegistraPagtoNoPedido
'   Registra o pagamento no pedido em decorrência de uma transação na Braspag
'   É necessário que a chamada desta função esteja dentro de uma transação,
'   a qual deve ser iniciada e finalizada pela rotina chamadora.
function BraspagRegistraPagtoNoPedido(byval tipo_operacao, byval pedido, byval idPedidoPagtoBraspag, byval idPedidoPagtoBraspagPAG, byval vl_transacao, byval usuario, byref mensagem_erro)
dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF
dim s, st_pagto_original, st_pagto_novo, st_pagto, s_id_pedido_pagto, msg_erro, s_log, id_pedido_base, loja, s_descricao_tipo_operacao
dim bandeira, pag_PaymentPlan, pag_NumberOfPayments
dim idFinPedidoHistPagto, s_hist_pagto_status, s_hist_pagto_descricao
dim lngRecordsAffected
dim rs, tPPB
dim s_ult_AF_GlobalStatus

	BraspagRegistraPagtoNoPedido = False
	
	s_log = ""
	mensagem_erro = ""

	id_pedido_base = retorna_num_pedido_base(pedido)
	
	if Not cria_recordset_pessimista(rs, msg_erro) then
		mensagem_erro = "Falha ao tentar abrir o recordset em modo de gravação: " & msg_erro
		exit function
		end if
	
	if Not calcula_pagamentos(pedido, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then
		mensagem_erro = "Falha ao tentar calcular os pagamentos anteriores do pedido: " & msg_erro
		exit function
		end if
	
'	REGISTRA O PAGAMENTO NO PEDIDO
	if Not gera_nsu(NSU_PEDIDO_PAGAMENTO, s_id_pedido_pagto, msg_erro) then
		mensagem_erro = "Falha ao tentar gerar o NSU para o novo registro de pagamento no pedido: " & msg_erro
		exit function
		end if
	
	if (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CANCELAMENTO) Or (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_ESTORNO) then
		if vl_transacao > 0 then vl_transacao = -1 * vl_transacao
		end if
	
	s = "INSERT INTO t_PEDIDO_PAGAMENTO (" & _
			"id, " & _
			"pedido, " & _
			"data, " & _
			"hora, " & _
			"valor, " & _
			"usuario, " & _
			"tipo_pagto, " & _
			"id_pedido_pagto_braspag" & _
		") VALUES (" & _
			"'" & s_id_pedido_pagto & "', " & _
			"'" & pedido & "', " & _
			bd_formata_data(Date) & ", " & _
			"'" & retorna_so_digitos(formata_hora(Now)) & "', " & _
			bd_formata_numero(vl_transacao) & ", " & _
			"'" & usuario & "', " & _
			"'" & COD_PAGTO_BRASPAG & "', " & _
			idPedidoPagtoBraspag & _
		")"
	cn.Execute s, lngRecordsAffected
	if lngRecordsAffected <> 1 then
		mensagem_erro = "Falha ao tentar gravar o novo registro de pagamento no pedido!!"
		exit function
		end if
	
'	PROCESSA A SITUAÇÃO DO PEDIDO C/ RELAÇÃO AOS PAGAMENTOS (QUITADO, PAGO PARCIAL, NÃO-PAGO)
	s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & id_pedido_base & "')"
	if rs.State <> 0 then rs.Close
	rs.Open s, cn
	if rs.Eof then
		mensagem_erro = "Pedido-base " & id_pedido_base & " não foi encontrado!"
		exit function
		end if
	
	loja = rs("loja")
	st_pagto_original = Trim("" & rs("st_pagto"))
	
'	PAGO (QUITADO)
'	~~~~~~~~~~~~~~
	if (vl_TotalFamiliaDevolucaoPrecoNF + vl_TotalFamiliaPago + vl_transacao) >= (vl_TotalFamiliaPrecoNF - MAX_VALOR_MARGEM_ERRO_PAGAMENTO) then
		if Trim("" & rs("st_pagto")) <> ST_PAGTO_PAGO then
			rs("dt_st_pagto") = Date
			rs("dt_hr_st_pagto") = Now
			rs("usuario_st_pagto") = usuario
			end if
		rs("st_pagto") = ST_PAGTO_PAGO
		s_log = "quitado"
		if (vl_TotalFamiliaDevolucaoPrecoNF + vl_TotalFamiliaPago + vl_transacao) > vl_TotalFamiliaPrecoNF then
			s_log = s_log & " (excedeu " & SIMBOLO_MONETARIO & " " & _
					formata_moeda((vl_TotalFamiliaDevolucaoPrecoNF+vl_TotalFamiliaPago+vl_transacao)-vl_TotalFamiliaPrecoNF) & ")"
		elseif (vl_TotalFamiliaDevolucaoPrecoNF + vl_TotalFamiliaPago + vl_transacao) < vl_TotalFamiliaPrecoNF then
			s_log = s_log & " (faltou " & SIMBOLO_MONETARIO & " " & _
					formata_moeda(vl_TotalFamiliaPrecoNF-(vl_TotalFamiliaDevolucaoPrecoNF+vl_TotalFamiliaPago+vl_transacao)) & ")"
			end if
	'	ANÁLISE DE CRÉDITO
		s_ult_AF_GlobalStatus = ""
		s = "SELECT ult_AF_GlobalStatus FROM t_PEDIDO_PAGTO_BRASPAG WHERE (id = " & idPedidoPagtoBraspag & ")"
		set tPPB = cn.Execute(s)
		if Not tPPB.Eof then
			s_ult_AF_GlobalStatus = Trim("" & tPPB("ult_AF_GlobalStatus"))
			tPPB.Close
			set tPPB = nothing
			end if
		
		dim blnCreditoOkAutomaticoDesativado
		blnCreditoOkAutomaticoDesativado = True
		if blnCreditoOkAutomaticoDesativado then
			if CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_VENDAS) And _
				(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)) And _
				(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)) And _
				(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)) And _
				(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) then
				if s_log <> "" then s_log = s_log & "; "
				s_log = s_log & " Análise de crédito: " & descricao_analise_credito(rs("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS)
				rs("analise_credito") = CLng(COD_AN_CREDITO_PENDENTE_VENDAS)
				rs("analise_credito_data") = Now
				rs("analise_credito_usuario") = ID_USUARIO_SISTEMA
				end if
		else
		'	TRANSAÇÕES INDICADAS P/ REVISÃO MANUAL DE PEDIDOS A PARTIR DE 5.000,00 SÃO COLOCADOS NO STATUS 'PENDENTE VENDAS'
		'	04/05/2015: TEMPORARIAMENTE, DEVIDO AO ELEVADO NÚMERO DE FRAUDES, O LIMITE DE 5.000,00 SERÁ ZERADO
			if (s_ult_AF_GlobalStatus = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW) And (vl_TotalFamiliaPrecoNF >= 0) then
				if CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_VENDAS) And _
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)) And _
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)) And _
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)) And _
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) then
					if s_log <> "" then s_log = s_log & "; "
					s_log = s_log & " Análise de crédito: " & descricao_analise_credito(rs("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS)
					rs("analise_credito") = CLng(COD_AN_CREDITO_PENDENTE_VENDAS)
					rs("analise_credito_data") = Now
					rs("analise_credito_usuario") = ID_USUARIO_SISTEMA
					end if
			else
				if (Trim("" & rs("indicador")) = "") And _
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)) And _
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)) And _
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)) And _
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) then
				'	TODO PEDIDO SEM INDICADOR DEVE PASSAR PELA ANÁLISE MANUAL
					if s_log <> "" then s_log = s_log & "; "
					s_log = s_log & " Análise de crédito: " & descricao_analise_credito(rs("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS) & " (motivo: pedido não possui indicador)"
					rs("analise_credito") = CLng(COD_AN_CREDITO_PENDENTE_VENDAS)
					rs("analise_credito_data") = Now
					rs("analise_credito_usuario") = ID_USUARIO_SISTEMA
				elseif (CLng(rs("analise_credito")) = CLng(COD_AN_CREDITO_ST_INICIAL)) Or (CLng(rs("analise_credito")) = CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)) And _
					(CLng(rs("st_forma_pagto_somente_cartao")) = 1) And _ 
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)) And _
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)) And _
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)) And _
					(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) then
					if s_log <> "" then s_log = s_log & "; "
					s_log = s_log & " Análise de crédito: " & descricao_analise_credito(rs("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_OK)
					rs("analise_credito") = CLng(COD_AN_CREDITO_OK)
					rs("analise_credito_data") = Now
					rs("analise_credito_usuario") = ID_USUARIO_SISTEMA
					end if
				end if
			end if 'if blnCreditoOkAutomaticoDesativado then-else
'	PAGAMENTO PARCIAL
'	~~~~~~~~~~~~~~~~~
	elseif (vl_TotalFamiliaPago + vl_transacao) > 0 then
		if Trim("" & rs("st_pagto")) <> ST_PAGTO_PARCIAL then
			rs("dt_st_pagto") = Date
			rs("dt_hr_st_pagto") = Now
			rs("usuario_st_pagto") = usuario
			end if
		rs("st_pagto") = ST_PAGTO_PARCIAL
		s_log = "pago parcial"
		if (CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_VENDAS)) And _
			(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)) And _
			(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)) And _
			(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)) And _
			(CLng(rs("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) then
			if s_log <> "" then s_log = s_log & "; "
			s_log = s_log & " Análise de crédito: " & descricao_analise_credito(rs("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS)
			rs("analise_credito") = CLng(COD_AN_CREDITO_PENDENTE_VENDAS)
			rs("analise_credito_data") = Now
			rs("analise_credito_usuario") = ID_USUARIO_SISTEMA
			end if
'	NÃO PAGO
'	~~~~~~~~
	else
		if Trim("" & rs("st_pagto")) <> ST_PAGTO_NAO_PAGO then
			rs("dt_st_pagto") = Date
			rs("dt_hr_st_pagto") = Now
			rs("usuario_st_pagto") = usuario
			end if
		rs("st_pagto") = ST_PAGTO_NAO_PAGO
		s_log = "não-pago"
		end if
	
	rs("vl_pago_familia") = vl_TotalFamiliaPago + vl_transacao
	st_pagto_novo = Trim("" & rs("st_pagto"))
	s_log = "Status do pedido: " & s_log & " (st_pagto: " & st_pagto_original & " => " & st_pagto_novo & ")"
	rs.Update
	if Err <> 0 then
		mensagem_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	s = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG WHERE (id = " & idPedidoPagtoBraspagPAG & ")"
	if rs.State <> 0 then rs.Close
	rs.Open s, cn
	if rs.Eof then
		mensagem_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag (id=" & idPedidoPagtoBraspagPAG & ")"
		exit function
		end if
	
	pag_PaymentPlan = Trim("" & rs("Req_PaymentDataCollection_PaymentPlan"))
	pag_NumberOfPayments = rs("Req_PaymentDataCollection_NumberOfPayments")
	
'	ANOTA NO REGISTRO DA TRANSAÇÃO QUE O PAGAMENTO JÁ FOI REGISTRADO NO PEDIDO
	s = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG WHERE (id = " & idPedidoPagtoBraspag & ")"
	if rs.State <> 0 then rs.Close
	rs.Open s, cn
	if rs.Eof then
		mensagem_erro = "Falha ao tentar localizar o registro da transação com a Braspag (id=" & idPedidoPagtoBraspag & ")"
		exit function
		end if
	
	bandeira = Trim("" & rs("bandeira"))
	rs("pagto_registrado_no_pedido_status") = 1
	rs("pagto_registrado_no_pedido_tipo_operacao") = tipo_operacao
	rs("pagto_registrado_no_pedido_data") = Date
	rs("pagto_registrado_no_pedido_data_hora") = Now
	rs("pagto_registrado_no_pedido_usuario") = usuario
	rs("pagto_registrado_no_pedido_id_pedido_pagamento") = s_id_pedido_pagto
	rs("pagto_registrado_no_pedido_st_pagto_anterior") = st_pagto_original
	rs("pagto_registrado_no_pedido_st_pagto_novo") = st_pagto_novo
	rs.Update
	
	s_descricao_tipo_operacao = BraspagDescricaoOperacaoRegistraPagto(tipo_operacao)
	s_log = "Registro automático de pagamento decorrente de operação de '" & s_descricao_tipo_operacao & "' na Braspag no valor de " & formata_moeda(vl_transacao) & " foi registrado com sucesso no pedido (t_PEDIDO_PAGTO_BRASPAG.id=" & Cstr(idPedidoPagtoBraspag) & ", t_PEDIDO_PAGAMENTO.id=" & s_id_pedido_pagto & "): " & s_log & ", Bandeira: " & BraspagDescricaoBandeira(bandeira) & ", Valor: " & formata_moeda(Abs(vl_transacao)) & ", Opção Pagamento: " & BraspagDescricaoParcelamento(pag_PaymentPlan, pag_NumberOfPayments, Abs(vl_transacao))
	grava_log usuario, loja, pedido, "", OP_LOG_PEDIDO_PAGTO_CONTABILIZADO_BRASPAG, s_log
	
'	REGISTRA NO HISTÓRICO DE PAGAMENTOS DO PEDIDO
	if Not fin_gera_nsu(T_FIN_PEDIDO_HIST_PAGTO, idFinPedidoHistPagto, msg_erro) then
		mensagem_erro = "Falha ao tentar gerar o NSU para o novo registro do histórico de pagamentos do pedido: " & msg_erro
		exit function
		end if
	
	s_hist_pagto_descricao = s_descricao_tipo_operacao & " (" & BraspagDescricaoBandeira(bandeira) & "): " & formata_moeda(Abs(vl_transacao))
	if (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CAPTURA) Or (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_AUTORIZACAO) then s_hist_pagto_descricao = s_hist_pagto_descricao & " em " & Cstr(pag_NumberOfPayments) & "x"
	s_hist_pagto_descricao = Left(s_hist_pagto_descricao, 60)
	if tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CAPTURA then
		s_hist_pagto_status = ST_T_FIN_PEDIDO_HIST_PAGTO__QUITADO
	elseif tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_AUTORIZACAO then
		s_hist_pagto_status = ST_T_FIN_PEDIDO_HIST_PAGTO__PREVISAO
	elseif (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CANCELAMENTO) Or (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_ESTORNO) then
		s_hist_pagto_status = ST_T_FIN_PEDIDO_HIST_PAGTO__CANCELADO
	else
		s_hist_pagto_status = "0"
		end if
	
	s = "INSERT INTO t_FIN_PEDIDO_HIST_PAGTO (" & _
			"id, " & _
			"pedido, " & _
			"status, " & _
			"ctrl_pagto_id_parcela, " & _
			"ctrl_pagto_modulo, " & _
			"dt_operacao, " & _
			"valor_total, " & _
			"valor_rateado, " & _
			"descricao, " & _
			"usuario_cadastro, " & _
			"usuario_ult_atualizacao" & _
		") VALUES (" & _
			Cstr(idFinPedidoHistPagto) & ", " & _
			"'" & pedido & "', " & _
			s_hist_pagto_status & ", " & _
			idPedidoPagtoBraspag & ", " & _
			CTRL_PAGTO_MODULO__BRASPAG_CARTAO & ", " & _
			bd_formata_data(Date) & ", " & _
			bd_formata_numero(Abs(vl_transacao)) & ", " & _
			bd_formata_numero(Abs(vl_transacao)) & ", " & _
			"'" & s_hist_pagto_descricao & "'" & ", " & _
			"'" & usuario & "', " & _
			"'" & usuario & "'" & _
		")"
	cn.Execute s, lngRecordsAffected
	if lngRecordsAffected <> 1 then
		mensagem_erro = "Falha ao tentar gravar o novo registro no histórico de pagamentos do pedido!!"
		exit function
		end if
	
	BraspagRegistraPagtoNoPedido = True
end function



' --------------------------------------------------------------------------------
'   cria_instancia_cl_BRASPAG_GetOrderIdData_TX
function cria_instancia_cl_BRASPAG_GetOrderIdData_TX(byval strMerchantId, byval strOrderId)
dim trx
	msg_erro = ""
	set trx = new cl_BRASPAG_GetOrderIdData_TX
	trx.PAG_Version = BRASPAG_PAGADOR_VERSION
	trx.PAG_RequestId = Lcase(gera_uid)
	trx.PAG_MerchantId = strMerchantId
	trx.PAG_OrderId = strOrderId
	set cria_instancia_cl_BRASPAG_GetOrderIdData_TX = trx
end function



' --------------------------------------------------------------------------------
'   BraspagCarregaDados_GetOrderIdDataResponse
'   Processa o xml de resposta e carrega os dados na estrutura.
function BraspagCarregaDados_GetOrderIdDataResponse(byval rxXml, byref r_rx, byref v_rx_item, byref msg_erro)
dim objXML
dim blnNodeNotFound, strNodeName, strNodeValue
dim oNode, oNodeErrorList, oNodeSet, oOrderIdDataCollection, oOrderIdDataItem
dim strTipoRetorno
dim strBraspagOrderId, strBraspagTransactionId
dim strErrorCode, strErrorMessage
	
	msg_erro = ""
	
	set r_rx = new cl_BRASPAG_GetOrderIdDataResponse_RX
	redim v_rx_item(0)
	set v_rx_item(UBound(v_rx_item)) = new cl_BRASPAG_OrderIdDataCollection_RX
	v_rx_item(UBound(v_rx_item)).blnHaDados = False
	
	Set objXML = Server.CreateObject("MSXML2.DOMDocument.3.0")
	objXML.Async = False
	objXML.LoadXml(rxXml)
	
	strTipoRetorno = objXML.documentElement.baseName
	if Ucase(strTipoRetorno) <> "ENVELOPE" then
		msg_erro = "Resposta recebida é inválida!" & chr(13)& rxXml
		exit function
		end if
	
'	CorrelationId
	strNodeName = "//GetOrderIdDataResponse/GetOrderIdDataResult/CorrelationId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_CorrelationId = strNodeValue
	
'	Success
	strNodeName = "//GetOrderIdDataResponse/GetOrderIdDataResult/Success"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_Success = strNodeValue
	
'	OrderIdDataCollection
	set oOrderIdDataCollection=objXML.documentElement.selectNodes("//GetOrderIdDataResponse/GetOrderIdDataResult/OrderIdDataCollection")
	if Not oOrderIdDataCollection is nothing then
		for each oOrderIdDataItem in oOrderIdDataCollection
		'	BraspagOrderId
			strBraspagOrderId = ""
			strNodeName = "//OrderIdTransactionResponse/BraspagOrderId"
			strBraspagOrderId = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strBraspagOrderId = ""
		'	BraspagTransactionId
			strBraspagTransactionId = ""
			strNodeName = "//OrderIdTransactionResponse/BraspagTransactionId"
			strBraspagTransactionId = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strBraspagTransactionId = ""
			if strBraspagTransactionId = "" then
				strNodeName = "//OrderIdTransactionResponse/BraspagTransactionId/guid"
				strBraspagTransactionId = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
				if blnNodeNotFound then strBraspagTransactionId = ""
				end if
			
		'	HÁ DADOS?
			if (strBraspagOrderId <> "") Or (strBraspagTransactionId <> "") then
				if v_rx_item(UBound(v_rx_item)).blnHaDados then
					redim preserve v_rx_item(UBound(v_rx_item)+1)
					set v_rx_item(UBound(v_rx_item)) = new cl_BRASPAG_OrderIdDataCollection_RX
					end if
				v_rx_item(UBound(v_rx_item)).blnHaDados = True
				v_rx_item(UBound(v_rx_item)).PAG_BraspagOrderId = strBraspagOrderId
				v_rx_item(UBound(v_rx_item)).PAG_BraspagTransactionId = strBraspagTransactionId
				end if
			next
		end if
	
'	ErrorCode/ErrorMessage
	set oNodeErrorList=objXML.documentElement.selectNodes("//GetOrderIdDataResponse/GetOrderIdDataResult/ErrorReportDataCollection")
	if Not oNodeErrorList is nothing then
		for each oNodeSet in oNodeErrorList
		'	OBTÉM OS DADOS DO ERRO P/ VERIFICAR SE HÁ CONTEÚDO
		'	ErrorCode
			strNodeName = "//ErrorReportDataResponse/ErrorCode"
			strErrorCode = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorCode = ""
			r_rx.PAG_ErrorCode = strErrorCode
			
		'	ErrorMessage
			strNodeName = "//ErrorReportDataResponse/ErrorMessage"
			strErrorMessage = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorMessage = ""
			r_rx.PAG_ErrorMessage = strErrorMessage
			next
		end if
	
'	faultcode
	strNodeName = "//Fault/faultcode"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_faultcode = strNodeValue
	
'	faultstring
	strNodeName = "//Fault/faultstring"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_faultstring = strNodeValue
end function



' --------------------------------------------------------------------------------
'   cria_instancia_cl_BRASPAG_GetTransactionData_TX
function cria_instancia_cl_BRASPAG_GetTransactionData_TX(byval strMerchantId, byval strBraspagTransactionId)
dim trx
	msg_erro = ""
	set trx = new cl_BRASPAG_GetTransactionData_TX
	trx.PAG_Version = BRASPAG_PAGADOR_VERSION
	trx.PAG_RequestId = Lcase(gera_uid)
	trx.PAG_MerchantId = strMerchantId
	trx.PAG_BraspagTransactionId = strBraspagTransactionId
	set cria_instancia_cl_BRASPAG_GetTransactionData_TX = trx
end function



' --------------------------------------------------------------------------------
'   BraspagCarregaDados_GetTransactionDataResponse
'   Processa o xml de resposta e carrega os dados na estrutura.
function BraspagCarregaDados_GetTransactionDataResponse(byval rxXml, byref msg_erro)
dim r_rx, objXML
dim blnNodeNotFound, strNodeName, strNodeValue
dim oNode, oNodeErrorList, oNodeSet
dim strTipoRetorno
dim strErrorCode, strErrorMessage
	
	msg_erro = ""
	
	set r_rx = new cl_BRASPAG_GetTransactionDataResponse_RX
	
	Set objXML = Server.CreateObject("MSXML2.DOMDocument.3.0")
	objXML.Async = False
	objXML.LoadXml(rxXml)
	
	strTipoRetorno = objXML.documentElement.baseName
	if Ucase(strTipoRetorno) <> "ENVELOPE" then
		msg_erro = "Resposta recebida é inválida!" & chr(13)& rxXml
		exit function
		end if
	
'	CorrelationId
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/CorrelationId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_CorrelationId = strNodeValue
	
'	Success
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/Success"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_Success = strNodeValue
	
'	BraspagTransactionId
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/BraspagTransactionId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_BraspagTransactionId = strNodeValue
	
'	OrderId
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/OrderId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_OrderId = strNodeValue
	
'	AcquirerTransactionId
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/AcquirerTransactionId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_AcquirerTransactionId = strNodeValue
	
'	PaymentMethod
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/PaymentMethod"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_PaymentMethod = strNodeValue
	
'	PaymentMethodName
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/PaymentMethodName"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_PaymentMethodName = strNodeValue
	
'	Amount
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/Amount"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_Amount = strNodeValue
	
'	AuthorizationCode
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/AuthorizationCode"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_AuthorizationCode = strNodeValue
	
'	NumberOfPayments
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/NumberOfPayments"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_NumberOfPayments = strNodeValue
	
'	Currency
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/Currency"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_Currency = strNodeValue
	
'	Country
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/Country"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_Country = strNodeValue
	
'	TransactionType
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/TransactionType"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_TransactionType = strNodeValue
	
'	Status
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/Status"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_Status = strNodeValue
	
'	ReceivedDate
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/ReceivedDate"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_ReceivedDate = strNodeValue
	
'	CapturedDate
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/CapturedDate"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_CapturedDate = strNodeValue
	
'	VoidedDate
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/VoidedDate"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_VoidedDate = strNodeValue
	
'	CreditCardToken
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/CreditCardToken"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_CreditCardToken = strNodeValue
	
'	ProofOfSale
	strNodeName = "//GetTransactionDataResponse/GetTransactionDataResult/ProofOfSale"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_ProofOfSale = strNodeValue
	
'	ErrorCode/ErrorMessage
	set oNodeErrorList=objXML.documentElement.selectNodes("//GetTransactionDataResponse/GetTransactionDataResult/ErrorReportDataCollection")
	if Not oNodeErrorList is nothing then
		for each oNodeSet in oNodeErrorList
		'	OBTÉM OS DADOS DO ERRO P/ VERIFICAR SE HÁ CONTEÚDO
		'	ErrorCode
			strNodeName = "//ErrorReportDataResponse/ErrorCode"
			strErrorCode = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorCode = ""
			r_rx.PAG_ErrorCode = strErrorCode
			
		'	ErrorMessage
			strNodeName = "//ErrorReportDataResponse/ErrorMessage"
			strErrorMessage = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorMessage = ""
			r_rx.PAG_ErrorMessage = strErrorMessage
			next
		end if
	
'	faultcode
	strNodeName = "//Fault/faultcode"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_faultcode = strNodeValue
	
'	faultstring
	strNodeName = "//Fault/faultstring"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_faultstring = strNodeValue
	
	set BraspagCarregaDados_GetTransactionDataResponse = r_rx
end function



' --------------------------------------------------------------------------------
'   cria_instancia_cl_BRASPAG_AF_UpdateStatus_TX
function cria_instancia_cl_BRASPAG_AF_UpdateStatus_TX(byval strMerchantId, byval strAntiFraudTransactionId, byval af_decision, af_comentario)
dim trx
	msg_erro = ""
	set trx = new cl_BRASPAG_AF_UpdateStatus_TX
	trx.AF_Version = BRASPAG_ANTIFRAUDE_VERSION
	trx.AF_RequestId = Lcase(gera_uid)
	trx.AF_MerchantId = strMerchantId
	trx.AF_AntiFraudTransactionId = strAntiFraudTransactionId
	trx.AF_NewStatus = af_decision
	trx.AF_Comment = af_comentario
	set cria_instancia_cl_BRASPAG_AF_UpdateStatus_TX = trx
end function



' --------------------------------------------------------------------------------
'   cria_instancia_cl_BRASPAG_FraudAnalysisTransactionDetails_TX
function cria_instancia_cl_BRASPAG_FraudAnalysisTransactionDetails_TX(byval strMerchantId, byval strAntiFraudTransactionId)
dim trx
	msg_erro = ""
	set trx = new cl_BRASPAG_FraudAnalysisTransactionDetails_TX
	trx.AF_Version = BRASPAG_ANTIFRAUDE_VERSION
	trx.AF_RequestId = Lcase(gera_uid)
	trx.AF_MerchantId = strMerchantId
	trx.AF_AntiFraudTransactionId = strAntiFraudTransactionId
	set cria_instancia_cl_BRASPAG_FraudAnalysisTransactionDetails_TX = trx
end function



' --------------------------------------------------------------------------------
'   BraspagCarregaDados_FraudAnalysisTransactionDetailsResponse
'   Processa o xml de resposta e carrega os dados na estrutura.
function BraspagCarregaDados_FraudAnalysisTransactionDetailsResponse(byval rxXml, byref msg_erro)
dim r_rx, objXML
dim blnNodeNotFound, strNodeName, strNodeValue
dim oNode, oNodeErrorList, oNodeSet
dim strTipoRetorno
dim strErrorCode, strErrorMessage
	
	msg_erro = ""
	
	set r_rx = new cl_BRASPAG_FraudAnalysisTransactionDetailsResponse_RX
	
	Set objXML = Server.CreateObject("MSXML2.DOMDocument.3.0")
	objXML.Async = False
	objXML.LoadXml(rxXml)
	
	strTipoRetorno = objXML.documentElement.baseName
	if Ucase(strTipoRetorno) <> "ENVELOPE" then
		msg_erro = "Resposta recebida é inválida!" & chr(13)& rxXml
		exit function
		end if
	
'	CorrelatedId
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/CorrelatedId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_CorrelatedId = strNodeValue
	
'	Success
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/Success"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_Success = strNodeValue
	
'	AntiFraudMerchantId
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/AntiFraudMerchantId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_AntiFraudMerchantId = strNodeValue
	
'	AntiFraudTransactionId
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/AntiFraudTransactionId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_AntiFraudTransactionId = strNodeValue
	
'	AntiFraudTransactionStatusCode
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/AntiFraudTransactionStatusCode"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_AntiFraudTransactionStatusCode = strNodeValue
	
'	AntiFraudReceiveDate
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/AntiFraudReceiveDate"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_AntiFraudReceiveDate = strNodeValue
	
'	AntiFraudStatusLastUpdateDate
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/AntiFraudStatusLastUpdateDate"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_AntiFraudStatusLastUpdateDate = strNodeValue
	
'	AntiFraudAnalysisScore
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/AntiFraudAnalysisScore"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_AntiFraudAnalysisScore = strNodeValue
	
'	BraspagTransactionId
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/BraspagTransactionId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_BraspagTransactionId = strNodeValue
	
'	MerchantOrderId
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/MerchantOrderId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_MerchantOrderId = strNodeValue
	
'	AntiFraudAcquirerConversionDate
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/AntiFraudAcquirerConversionDate"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_AntiFraudAcquirerConversionDate = strNodeValue
	
'	AntiFraudTransactionOriginalStatusCode
	strNodeName = "//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/AntiFraudTransactionOriginalStatusCode"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_AntiFraudTransactionOriginalStatusCode = strNodeValue
	
'	ErrorCode/ErrorMessage
	set oNodeErrorList=objXML.documentElement.selectNodes("//FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/ErrorReportCollection")
	if Not oNodeErrorList is nothing then
		for each oNodeSet in oNodeErrorList
		'	OBTÉM OS DADOS DO ERRO P/ VERIFICAR SE HÁ CONTEÚDO
		'	ErrorCode
			strNodeName = "//ErrorReportResponse/ErrorCode"
			strErrorCode = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorCode = ""
			r_rx.AF_ErrorCode = strErrorCode
			
		'	ErrorMessage
			strNodeName = "//ErrorReportResponse/ErrorMessage"
			strErrorMessage = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorMessage = ""
			r_rx.AF_ErrorMessage = strErrorMessage
			next
		end if
	
'	faultcode
	strNodeName = "//Fault/faultcode"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_faultcode = strNodeValue
	
'	faultstring
	strNodeName = "//Fault/faultstring"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_faultstring = strNodeValue
	
	set BraspagCarregaDados_FraudAnalysisTransactionDetailsResponse = r_rx
end function



' --------------------------------------------------------------------------------
'   BraspagCarregaDados_AF_UpdateStatusResponse
'   Processa o xml de resposta e carrega os dados na estrutura.
function BraspagCarregaDados_AF_UpdateStatusResponse(byval rxXml, byref msg_erro)
dim r_rx, objXML
dim blnNodeNotFound, strNodeName, strNodeValue
dim oNode, oNodeErrorList, oNodeSet
dim strTipoRetorno
dim strErrorCode, strErrorMessage
	
	msg_erro = ""
	
	set r_rx = new cl_BRASPAG_AF_UpdateStatusResponse_RX
	
	Set objXML = Server.CreateObject("MSXML2.DOMDocument.3.0")
	objXML.Async = False
	objXML.LoadXml(rxXml)
	
	strTipoRetorno = objXML.documentElement.baseName
	if Ucase(strTipoRetorno) <> "ENVELOPE" then
		msg_erro = "Resposta recebida é inválida!" & chr(13)& rxXml
		exit function
		end if
	
'	CorrelatedId
	strNodeName = "//UpdateStatusResponse/UpdateStatusResult/CorrelatedId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_CorrelatedId = strNodeValue
	
'	Success
	strNodeName = "//UpdateStatusResponse/UpdateStatusResult/Success"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_Success = strNodeValue
	
'	AntiFraudTransactionId
	strNodeName = "//UpdateStatusResponse/UpdateStatusResult/AntiFraudTransactionId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_AntiFraudTransactionId = strNodeValue
	
'	RequestStatusCode
	strNodeName = "//UpdateStatusResponse/UpdateStatusResult/RequestStatusCode"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_RequestStatusCode = strNodeValue
	
'	RequestStatusDescription
	strNodeName = "//UpdateStatusResponse/UpdateStatusResult/RequestStatusDescription"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_RequestStatusDescription = strNodeValue
	
'	ErrorCode/ErrorMessage
	set oNodeErrorList=objXML.documentElement.selectNodes("//UpdateStatusResponse/UpdateStatusResult/ErrorReportCollection")
	if Not oNodeErrorList is nothing then
		for each oNodeSet in oNodeErrorList
		'	OBTÉM OS DADOS DO ERRO P/ VERIFICAR SE HÁ CONTEÚDO
		'	ErrorCode
			strNodeName = "//ErrorReportResponse/ErrorCode"
			strErrorCode = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorCode = ""
			r_rx.AF_ErrorCode = strErrorCode
			
		'	ErrorMessage
			strNodeName = "//ErrorReportResponse/ErrorMessage"
			strErrorMessage = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorMessage = ""
			r_rx.AF_ErrorMessage = strErrorMessage
			next
		end if
	
'	faultcode
	strNodeName = "//Fault/faultcode"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_faultcode = strNodeValue
	
'	faultstring
	strNodeName = "//Fault/faultstring"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.AF_faultstring = strNodeValue
	
	set BraspagCarregaDados_AF_UpdateStatusResponse = r_rx
end function



' --------------------------------------------------------------------------------
'   BraspagVerificaPreRequisito_BraspagTransactionId
'   Verifica se há a informação de 'BraspagTransactionId'. Caso não,
'   executa a consulta 'GetOrderIdData' usando o campo 'OrderId' p/
'   tentar obter o 'BraspagTransactionId', que é necessário p/ a maioria
'   das requisicoes.
function BraspagVerificaPreRequisito_BraspagTransactionId(byval id_pedido_pagto_braspag, byval id_pedido_pagto_braspag_pag, byval usuario, byref msg_erro)
dim t, t_PP_BRASPAG, t_PP_BRASPAG_PAG, t_PP_BRASPAG_PAG_OP_COMPLEMENTAR, t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML
dim i, lngRecordsAffected, intQtdeRespostas
dim idPedidoPagtoBraspagPagOpComplementar, idPedidoPagtoBraspagPagOpComplXmlTx, idPedidoPagtoBraspagPagOpComplXmlRx
dim strMerchantId, strBraspagTransactionId, strOrderId
dim strSql
dim txXml, rxXml
dim r_rx, v_rx_item()

	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(t_PP_BRASPAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag & ")"
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	t_PP_BRASPAG.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if t_PP_BRASPAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG_PAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag_pag & ")"
	if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
	t_PP_BRASPAG_PAG.open strSql, cn
	if t_PP_BRASPAG_PAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag!!"
		exit function
		end if
	
	strMerchantId = Trim("" & t_PP_BRASPAG_PAG("Req_OrderData_MerchantId"))
	strOrderId = Trim("" & t_PP_BRASPAG_PAG("Req_OrderData_OrderId"))
	strBraspagTransactionId = Trim("" & t_PP_BRASPAG_PAG("Resp_PaymentDataResponse_BraspagTransactionId"))
	
'	A INFORMAÇÃO 'BraspagTransactionId' ESTÁ DISPONÍVEL?
'	SE O CAMPO 'Req_OrderData_OrderId' ESTIVER VAZIO NÃO SERÁ POSSÍVEL REALIZAR A CONSULTA 'GetOrderIdData'
	if (strBraspagTransactionId <> "") Or (strOrderId = "") then
	'	FECHA TABELAS
		if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
		set t_PP_BRASPAG = nothing
		
		if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
		set t_PP_BRASPAG_PAG = nothing
		
		exit function
		end if
	
	strSql = "SELECT" & _
				" Count(*) AS qtde" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG_PAG" & _
			" WHERE" & _
				" (Req_OrderData_OrderId = '" & strOrderId & "')"
	set t = cn.Execute(strSql)
	if Not t.Eof then
	'	SE HOUVER MAIS DO QUE UMA TRANSAÇÃO C/ O MESMO VALOR DE 'OrderId'
	'	NÃO SERÁ POSSÍVEL DETERMINAR A QUAL DELAS SE REFERE A RESPOSTA
	'	RETORNADA PELA CONSULTA 'GetOrderIdData'.
	'	PORTANTO, NESTE CASO OPTOU-SE POR NÃO FAZER A CONSULTA AO INVÉS
	'	DE CORRER O RISCO DE EXIBIR UMA INFORMAÇÃO INCONSISTENTE.
	'	EX: A PRIMEIRA TENTATIVA DE PAGAMENTO FALHOU DE FORMA QUE O CAMPO 'BraspagTransactionId' NÃO RETORNOU DA BRASPAG OU NÃO FOI GRAVADO CORRETAMENTE NO BD.
	'		A SEGUNDA TENTATIVA TAMBÉM FALHOU DA MESMA MANEIRA.
	'		A TERCEIRA TENTATIVA FOI BEM-SUCEDIDA.
	'		SE AS 3 TRANSAÇÕES POSSUÍREM O MESMO VALOR DE 'OrderId', A CONSULTA 'GetOrderIdData'
	'		FEITA P/ A TENTATIVA 1 OU 2 PODERÁ RETORNAR O 'BraspagTransactionId' DA TENTATIVA 3.
	'		O USO DESSE 'BraspagTransactionId' POSTERIORMENTE NA CONSULTA 'GetTransactionData'
	'		CAUSARIA UM ENTENDIMENTO ERRADO DE QUE HOUVE MAIS DO QUE UMA TRANSAÇÃO BEM-SUCEDIDA.
		if CLng(t("qtde")) > 1 then
		'	FECHA TABELAS
			t.Close
			set t = nothing
			
			if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
			set t_PP_BRASPAG = nothing
			
			if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
			set t_PP_BRASPAG_PAG = nothing
			
			exit function
			end if
		end if
	
	t.Close
	set t = nothing
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG_OP_COMPLEMENTAR, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_GetOrderIdData_TX(strMerchantId, strOrderId)
	txXml = BraspagXmlMontaRequisicaoGetOrderIdData(trx)
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR, idPedidoPagtoBraspagPagOpComplementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagPagOpComplXmlTx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplXmlTx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplXmlTx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagPagOpComplXmlRx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplXmlRx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplXmlRx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível consultar a Braspag devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id_pedido_pagto_braspag") = CLng(id_pedido_pagto_braspag)
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id_pedido_pagto_braspag_pag") = CLng(id_pedido_pagto_braspag_pag)
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("operacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_ORDERID_DATA
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("usuario") = usuario
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("request_data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("request_data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_RequestId") = trx.PAG_RequestId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_Version") = trx.PAG_Version
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_MerchantId") = trx.PAG_MerchantId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_OrderId") = trx.PAG_OrderId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Update
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagPagOpComplXmlTx
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_pag_op_complementar") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_ORDERID_DATA
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("xml") = txXml
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Update
	
	rxXml = BraspagEnviaTransacao(txXml, BRASPAG_WS_ENDERECO_PAGADOR_QUERY)
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR WHERE (id = " & idPedidoPagtoBraspagPagOpComplementar & ")"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Open strSql, cn
	if Not t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Eof then
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("response_data") = Date
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("response_data_hora") = Now
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("st_resposta_recebida") = 1
		if Trim(rxXml) = "" then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("st_resposta_vazia") = 1
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Update
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagPagOpComplXmlRx
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_pag_op_complementar") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_ORDERID_DATA
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("xml") = rxXml
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Update
	
	call BraspagCarregaDados_GetOrderIdDataResponse(rxXml, r_rx, v_rx_item, msg_erro)
	if msg_erro <> "" then exit function
	
'	SE OBTEVE UM VALOR ÚNICO DE 'BraspagTransactionId', ATUALIZA A INFORMAÇÃO NO BD
	strBraspagTransactionId = ""
	intQtdeRespostas = 0
	for i = LBound(v_rx_item) to UBound(v_rx_item)
		if Trim("" & v_rx_item(i).PAG_BraspagTransactionId) <> "" then
			intQtdeRespostas = intQtdeRespostas + 1
			strBraspagTransactionId = Trim("" & v_rx_item(i).PAG_BraspagTransactionId)
			end if
		next
	
	if (intQtdeRespostas = 1) And (strBraspagTransactionId <> "") then
		strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG_PAG SET" & _
					" Resp_PaymentDataResponse_BraspagTransactionId = '" & strBraspagTransactionId & "'" & _
				" WHERE" & _
					" (id = " & id_pedido_pagto_braspag_pag & ")"
		cn.Execute strSql, lngRecordsAffected
		
		strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR SET" & _
					" st_sucesso = 1," & _
					" Resp_BraspagTransactionId = '" & strBraspagTransactionId & "'" & _
				" WHERE" & _
					" (id = " & idPedidoPagtoBraspagPagOpComplementar & ")"
		cn.Execute strSql, lngRecordsAffected
		end if
	
'	FECHA TABELAS
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	set t_PP_BRASPAG = nothing

	if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
	set t_PP_BRASPAG_PAG = nothing

	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	set t_PP_BRASPAG_PAG_OP_COMPLEMENTAR = nothing

	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	set t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML = nothing
end function



' --------------------------------------------------------------------------------
'   BraspagProcessaConsulta_GetTransactionData
'   Executa a consulta e realiza o processamento relacionado ao BD.
function BraspagProcessaConsulta_GetTransactionData(byval id_pedido_pagto_braspag, byval id_pedido_pagto_braspag_pag, byval usuario, byref trx, byref r_rx, byref msg_erro)
dim t_PP_BRASPAG, t_PP_BRASPAG_PAG, t_PP_BRASPAG_PAG_OP_COMPLEMENTAR, t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML
dim lngRecordsAffected
dim idPedidoPagtoBraspagPagOpComplementar, idPedidoPagtoBraspagPagOpComplXmlTx, idPedidoPagtoBraspagPagOpComplXmlRx
dim strCapturedDate, strVoidedDate
dim strMerchantId, strBraspagTransactionId
dim strSql
dim txXml, rxXml
	
	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(t_PP_BRASPAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG_OP_COMPLEMENTAR, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag & ")"
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	t_PP_BRASPAG.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if t_PP_BRASPAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG_PAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag_pag & ")"
	if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
	t_PP_BRASPAG_PAG.open strSql, cn
	if t_PP_BRASPAG_PAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag!!"
		exit function
		end if
	
	strMerchantId = Trim("" & t_PP_BRASPAG_PAG("Req_OrderData_MerchantId"))
	strBraspagTransactionId = Trim("" & t_PP_BRASPAG_PAG("Resp_PaymentDataResponse_BraspagTransactionId"))
	
	if strBraspagTransactionId = "" then
		msg_erro = "Não é possível consultar a Braspag porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!"
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_GetTransactionData_TX(strMerchantId, strBraspagTransactionId)
	txXml = BraspagXmlMontaRequisicaoGetTransactionData(trx)
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR, idPedidoPagtoBraspagPagOpComplementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagPagOpComplXmlTx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplXmlTx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplXmlTx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagPagOpComplXmlRx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplXmlRx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplXmlRx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível consultar a Braspag devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id_pedido_pagto_braspag") = CLng(id_pedido_pagto_braspag)
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id_pedido_pagto_braspag_pag") = CLng(id_pedido_pagto_braspag_pag)
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("operacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_TRANSACTION_DATA
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("usuario") = usuario
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("request_data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("request_data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_RequestId") = trx.PAG_RequestId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_Version") = trx.PAG_Version
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_MerchantId") = trx.PAG_MerchantId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_BraspagTransactionId") = trx.PAG_BraspagTransactionId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Update
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagPagOpComplXmlTx
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_pag_op_complementar") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_TRANSACTION_DATA
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("xml") = txXml
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Update
	
	rxXml = BraspagEnviaTransacao(txXml, BRASPAG_WS_ENDERECO_PAGADOR_QUERY)
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR WHERE (id = " & idPedidoPagtoBraspagPagOpComplementar & ")"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Open strSql, cn
	if Not t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Eof then
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("response_data") = Date
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("response_data_hora") = Now
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("st_resposta_recebida") = 1
		if Trim(rxXml) = "" then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("st_resposta_vazia") = 1
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Update
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagPagOpComplXmlRx
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_pag_op_complementar") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_TRANSACTION_DATA
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("xml") = rxXml
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Update
	
	set r_rx = BraspagCarregaDados_GetTransactionDataResponse(rxXml, msg_erro)
	if msg_erro <> "" then exit function
	
'	ATUALIZA O ÚLTIMO STATUS DA TRANSAÇÃO
	strCapturedDate = "NULL"
	if r_rx.PAG_CapturedDate <> "" then
	'	DATA/HORA ESTÁ NO FORMATO AM/PM
		strCapturedDate = bd_monta_data_hora(converte_para_datetime_from_mmddyyyy_hhmmss_am_pm(r_rx.PAG_CapturedDate))
		end if
	
	strVoidedDate = "NULL"
	if r_rx.PAG_VoidedDate <> "" then
	'	DATA/HORA ESTÁ NO FORMATO AM/PM
		strVoidedDate = bd_monta_data_hora(converte_para_datetime_from_mmddyyyy_hhmmss_am_pm(r_rx.PAG_VoidedDate))
		end if
	
	strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG SET" & _
				" ult_PAG_GlobalStatus = '" & decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(r_rx.PAG_Status) & "'," & _
				" ult_PAG_atualizacao_data_hora = getdate()," & _
				" ult_PAG_atualizacao_usuario = '" & usuario & "'," & _
				" ult_id_pedido_pagto_braspag_pag_op_complementar = " & idPedidoPagtoBraspagPagOpComplementar & "," & _
				" ult_PAG_CapturedDate = " & strCapturedDate & "," & _
				" ult_PAG_VoidedDate = " & strVoidedDate & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag & ")"
	cn.Execute strSql, lngRecordsAffected
	
	strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR SET" & _
				" Resp_AuthorizationCode = '" & r_rx.PAG_AuthorizationCode & "'," & _
				" Resp_ProofOfSale = '" & r_rx.PAG_ProofOfSale & "'," & _
				" st_sucesso = 1, " & _
				" Resp_GetTransactionDataResponse_Status = '" & r_rx.PAG_Status & "'" & _
			" WHERE" & _
				" (id = " & idPedidoPagtoBraspagPagOpComplementar & ")"
	cn.Execute strSql, lngRecordsAffected
	
'	FECHA TABELAS
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	set t_PP_BRASPAG = nothing

	if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
	set t_PP_BRASPAG_PAG = nothing

	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	set t_PP_BRASPAG_PAG_OP_COMPLEMENTAR = nothing

	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	set t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML = nothing
end function



' --------------------------------------------------------------------------------
'   BraspagProcessaConsulta_FraudAnalysisTransactionDetails
'   Executa a consulta e realiza o processamento relacionado ao BD.
function BraspagProcessaConsulta_FraudAnalysisTransactionDetails(byval id_pedido_pagto_braspag, byval id_pedido_pagto_braspag_af, byval usuario, byref trx, byref r_rx, byref msg_erro)
dim t_PP_BRASPAG, t_PP_BRASPAG_AF, t_PP_BRASPAG_AF_OP_COMPLEMENTAR, t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML
dim lngRecordsAffected
dim idPedidoPagtoBraspagAfOpComplementar, idPedidoPagtoBraspagAfOpComplXmlTx, idPedidoPagtoBraspagAfOpComplXmlRx
dim strMerchantId, strAntiFraudTransactionId
dim strSql
dim txXml, rxXml
	
	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(t_PP_BRASPAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_AF, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_AF_OP_COMPLEMENTAR, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag & ")"
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	t_PP_BRASPAG.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if t_PP_BRASPAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG_AF" & _
			" WHERE" & _
				" (id_pedido_pagto_braspag = " & id_pedido_pagto_braspag_af & ")"
	if t_PP_BRASPAG_AF.State <> 0 then t_PP_BRASPAG_AF.Close
	t_PP_BRASPAG_AF.open strSql, cn
	if t_PP_BRASPAG_AF.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Antifraude!!"
		exit function
		end if
	
	strMerchantId = Trim("" & t_PP_BRASPAG_AF("Req_MerchantId"))
	strAntiFraudTransactionId = Trim("" & t_PP_BRASPAG_AF("Resp_AntiFraudTransactionId"))
	
	if strAntiFraudTransactionId = "" then
		msg_erro = "Não é possível consultar os dados no Antifraude porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!"
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_FraudAnalysisTransactionDetails_TX(strMerchantId, strAntiFraudTransactionId)
	txXml = BraspagXmlMontaRequisicaoFraudAnalysisTransactionDetails(trx)
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR, idPedidoPagtoBraspagAfOpComplementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagAfOpComplementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagAfOpComplementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagAfOpComplXmlTx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagAfOpComplXmlTx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagAfOpComplXmlTx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagAfOpComplXmlRx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagAfOpComplXmlRx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagAfOpComplXmlRx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível consultar o Antifraude devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR WHERE (id = -1)"
	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Open strSql, cn
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR.AddNew
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("id") = idPedidoPagtoBraspagAfOpComplementar
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("id_pedido_pagto_braspag") = CLng(id_pedido_pagto_braspag)
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("id_pedido_pagto_braspag_af") = CLng(id_pedido_pagto_braspag_af)
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("operacao") = BRASPAG_TIPO_TRANSACAO__AF_FRAUD_ANALYSIS_TRANSACTION_DETAILS
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("usuario") = usuario
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("request_data") = Date
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("request_data_hora") = Now
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("Req_RequestId") = trx.AF_RequestId
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("Req_Version") = trx.AF_Version
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("Req_MerchantId") = trx.AF_MerchantId
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("Req_AntiFraudTransactionId") = trx.AF_AntiFraudTransactionId
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Update
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagAfOpComplXmlTx
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_af_op_complementar") = idPedidoPagtoBraspagAfOpComplementar
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__AF_FRAUD_ANALYSIS_TRANSACTION_DETAILS
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("xml") = txXml
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Update
	
	rxXml = BraspagEnviaTransacao(txXml, BRASPAG_WS_ENDERECO_ANTIFRAUDE_QUERY)
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR WHERE (id = " & idPedidoPagtoBraspagAfOpComplementar & ")"
	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Open strSql, cn
	if Not t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Eof then
		t_PP_BRASPAG_AF_OP_COMPLEMENTAR("response_data") = Date
		t_PP_BRASPAG_AF_OP_COMPLEMENTAR("response_data_hora") = Now
		t_PP_BRASPAG_AF_OP_COMPLEMENTAR("st_resposta_recebida") = 1
		if Trim(rxXml) = "" then t_PP_BRASPAG_AF_OP_COMPLEMENTAR("st_resposta_vazia") = 1
		t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Update
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagAfOpComplXmlRx
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_af_op_complementar") = idPedidoPagtoBraspagAfOpComplementar
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__AF_FRAUD_ANALYSIS_TRANSACTION_DETAILS
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("xml") = rxXml
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Update
	
	set r_rx = BraspagCarregaDados_FraudAnalysisTransactionDetailsResponse(rxXml, msg_erro)
	if msg_erro <> "" then exit function
	
'	ATUALIZA O ÚLTIMO STATUS DA TRANSAÇÃO
	strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG SET" & _
				" ult_AF_GlobalStatus = '" & decodifica_FraudAnalysisTransactionDetailsResponseAntiFraudTransactionStatusCode_para_GlobalStatus(r_rx.AF_AntiFraudTransactionStatusCode) & "'," & _
				" ult_AF_atualizacao_data_hora = getdate()," & _
				" ult_AF_atualizacao_usuario = '" & usuario & "'," & _
				" ult_id_pedido_pagto_braspag_af_op_complementar = " & idPedidoPagtoBraspagAfOpComplementar & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag & ")"
	cn.Execute strSql, lngRecordsAffected
	
	strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR SET" & _
				" st_sucesso = 1, " & _
				" Resp_FraudAnalysisTransactionDetailsResponse_AntiFraudTransactionStatusCode = '" & r_rx.AF_AntiFraudTransactionStatusCode & "'" & _
			" WHERE" & _
				" (id = " & idPedidoPagtoBraspagAfOpComplementar & ")"
	cn.Execute strSql, lngRecordsAffected
	
'	FECHA TABELAS
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	set t_PP_BRASPAG = nothing

	if t_PP_BRASPAG_AF.State <> 0 then t_PP_BRASPAG_AF.Close
	set t_PP_BRASPAG_AF = nothing

	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Close
	set t_PP_BRASPAG_AF_OP_COMPLEMENTAR = nothing

	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Close
	set t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML = nothing
end function



' ------------------------------------------------------------------------
'   BraspagXmlMontaRequisicaoCaptureCreditCardTransaction
function BraspagXmlMontaRequisicaoCaptureCreditCardTransaction(ByRef trx)
dim xml
	xml =	"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & chr(13) & _
			"	<soap:Body>" & chr(13) & _
			"		<CaptureCreditCardTransaction xmlns=""https://www.pagador.com.br/webservice/pagador"">" & chr(13) & _
			"			<request>" & chr(13) & _
			"				<Version>" & trx.PAG_Version & "</Version>" & chr(13) & _
			"				<RequestId>" & trx.PAG_RequestId & "</RequestId>" & chr(13) & _
			"				<MerchantId>" & trx.PAG_MerchantId & "</MerchantId>" & chr(13) & _
			"				<TransactionDataCollection>" & chr(13) & _
			"					<TransactionDataRequest>" & chr(13) & _
			"						<BraspagTransactionId>" & trx.PAG_BraspagTransactionId & "</BraspagTransactionId>" & chr(13) & _
			"						<Amount>" & trx.PAG_Amount & "</Amount>" & chr(13) & _
			"						<ServiceTaxAmount>" & trx.PAG_ServiceTaxAmount & "</ServiceTaxAmount>" & chr(13) & _
			"					</TransactionDataRequest>" & chr(13) & _
			"				</TransactionDataCollection>" & chr(13) & _
			"			</request>" & chr(13) & _
			"		</CaptureCreditCardTransaction>" & chr(13) & _
			"	</soap:Body>" & chr(13) & _
			"</soap:Envelope>" & chr(13)
	BraspagXmlMontaRequisicaoCaptureCreditCardTransaction = xml
end function



' ------------------------------------------------------------------------
'   BraspagXmlMontaRequisicaoVoidCreditCardTransaction
function BraspagXmlMontaRequisicaoVoidCreditCardTransaction(ByRef trx)
dim xml
	xml =	"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & chr(13) & _
			"	<soap:Body>" & chr(13) & _
			"		<VoidCreditCardTransaction xmlns=""https://www.pagador.com.br/webservice/pagador"">" & chr(13) & _
			"			<request>" & chr(13) & _
			"				<Version>" & trx.PAG_Version & "</Version>" & chr(13) & _
			"				<RequestId>" & trx.PAG_RequestId & "</RequestId>" & chr(13) & _
			"				<MerchantId>" & trx.PAG_MerchantId & "</MerchantId>" & chr(13) & _
			"				<TransactionDataCollection>" & chr(13) & _
			"					<TransactionDataRequest>" & chr(13) & _
			"						<BraspagTransactionId>" & trx.PAG_BraspagTransactionId & "</BraspagTransactionId>" & chr(13) & _
			"						<Amount>" & trx.PAG_Amount & "</Amount>" & chr(13) & _
			"						<ServiceTaxAmount>" & trx.PAG_ServiceTaxAmount & "</ServiceTaxAmount>" & chr(13) & _
			"					</TransactionDataRequest>" & chr(13) & _
			"				</TransactionDataCollection>" & chr(13) & _
			"			</request>" & chr(13) & _
			"		</VoidCreditCardTransaction>" & chr(13) & _
			"	</soap:Body>" & chr(13) & _
			"</soap:Envelope>" & chr(13)
	BraspagXmlMontaRequisicaoVoidCreditCardTransaction = xml
end function



' --------------------------------------------------------------------------------
'   cria_instancia_cl_BRASPAG_CaptureCreditCardTransaction_TX
function cria_instancia_cl_BRASPAG_CaptureCreditCardTransaction_TX(byval strMerchantId, byval strBraspagTransactionId, byval strAmount, byval strServiceTaxAmount)
dim trx
	msg_erro = ""
	set trx = new cl_BRASPAG_CaptureCreditCardTransaction_TX
	trx.PAG_Version = BRASPAG_PAGADOR_VERSION
	trx.PAG_RequestId = Lcase(gera_uid)
	trx.PAG_MerchantId = strMerchantId
	trx.PAG_BraspagTransactionId = strBraspagTransactionId
	trx.PAG_Amount = Trim("" & strAmount)
	trx.PAG_ServiceTaxAmount = Trim("" & strServiceTaxAmount)
	set cria_instancia_cl_BRASPAG_CaptureCreditCardTransaction_TX = trx
end function



' --------------------------------------------------------------------------------
'   cria_instancia_cl_BRASPAG_VoidCreditCardTransaction_TX
function cria_instancia_cl_BRASPAG_VoidCreditCardTransaction_TX(byval strMerchantId, byval strBraspagTransactionId, byval strAmount, byval strServiceTaxAmount)
dim trx
	msg_erro = ""
	set trx = new cl_BRASPAG_VoidCreditCardTransaction_TX
	trx.PAG_Version = BRASPAG_PAGADOR_VERSION
	trx.PAG_RequestId = Lcase(gera_uid)
	trx.PAG_MerchantId = strMerchantId
	trx.PAG_BraspagTransactionId = strBraspagTransactionId
	trx.PAG_Amount = Trim("" & strAmount)
	trx.PAG_ServiceTaxAmount = Trim("" & strServiceTaxAmount)
	set cria_instancia_cl_BRASPAG_VoidCreditCardTransaction_TX = trx
end function



' --------------------------------------------------------------------------------
'   BraspagCarregaDados_CaptureCreditCardTransactionResponse
'   Processa o xml de resposta e carrega os dados na estrutura.
function BraspagCarregaDados_CaptureCreditCardTransactionResponse(byval rxXml, byref msg_erro)
dim r_rx, objXML
dim blnNodeNotFound, strNodeName, strNodeValue
dim oNode, oNodeErrorList, oNodeSet
dim oTransactionDataCollection, oTransactionSet
dim strTipoRetorno
dim strErrorCode, strErrorMessage
	
	msg_erro = ""
	
	set r_rx = new cl_BRASPAG_CaptureCreditCardTransactionResponse_RX
	
	Set objXML = Server.CreateObject("MSXML2.DOMDocument.3.0")
	objXML.Async = False
	objXML.LoadXml(rxXml)
	
	strTipoRetorno = objXML.documentElement.baseName
	if Ucase(strTipoRetorno) <> "ENVELOPE" then
		msg_erro = "Resposta recebida é inválida!" & chr(13)& rxXml
		exit function
		end if
	
	set oTransactionDataCollection=objXML.documentElement.selectNodes("//CaptureCreditCardTransactionResponse/CaptureCreditCardTransactionResult/TransactionDataCollection")
	if Not oTransactionDataCollection is nothing then
		for each oTransactionSet in oTransactionDataCollection
		'	BraspagTransactionId
			strNodeName = "//TransactionDataResponse/BraspagTransactionId"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_BraspagTransactionId = strNodeValue
			
		'	AcquirerTransactionId
			strNodeName = "//TransactionDataResponse/AcquirerTransactionId"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_AcquirerTransactionId = strNodeValue
			
		'	Amount
			strNodeName = "//TransactionDataResponse/Amount"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_Amount = strNodeValue
			
		'	AuthorizationCode
			strNodeName = "//TransactionDataResponse/AuthorizationCode"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_AuthorizationCode = strNodeValue
			
		'	ReturnCode
			strNodeName = "//TransactionDataResponse/ReturnCode"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ReturnCode = strNodeValue
			
		'	ReturnMessage
			strNodeName = "//TransactionDataResponse/ReturnMessage"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ReturnMessage = strNodeValue
			
		'	Status
			strNodeName = "//TransactionDataResponse/Status"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_Status = strNodeValue
			
		'	ProofOfSale
			strNodeName = "//TransactionDataResponse/ProofOfSale"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ProofOfSale = strNodeValue
			
		'	ServiceTaxAmount
			strNodeName = "//TransactionDataResponse/ServiceTaxAmount"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ServiceTaxAmount = strNodeValue
			next
		end if
	
'	CorrelationId
	strNodeName = "//CaptureCreditCardTransactionResponse/CaptureCreditCardTransactionResult/CorrelationId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_CorrelationId = strNodeValue
	
'	Success
	strNodeName = "//CaptureCreditCardTransactionResponse/CaptureCreditCardTransactionResult/Success"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_Success = strNodeValue

'	ErrorCode/ErrorMessage
	set oNodeErrorList=objXML.documentElement.selectNodes("//CaptureCreditCardTransactionResponse/CaptureCreditCardTransactionResult/ErrorReportDataCollection")
	if Not oNodeErrorList is nothing then
		for each oNodeSet in oNodeErrorList
		'	OBTÉM OS DADOS DO ERRO P/ VERIFICAR SE HÁ CONTEÚDO
		'	ErrorCode
			strNodeName = "//ErrorReportDataResponse/ErrorCode"
			strErrorCode = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorCode = ""
			r_rx.PAG_ErrorCode = strErrorCode
			
		'	ErrorMessage
			strNodeName = "//ErrorReportDataResponse/ErrorMessage"
			strErrorMessage = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorMessage = ""
			r_rx.PAG_ErrorMessage = strErrorMessage
			next
		end if
	
'	faultcode
	strNodeName = "//Fault/faultcode"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_faultcode = strNodeValue
	
'	faultstring
	strNodeName = "//Fault/faultstring"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_faultstring = strNodeValue
	
	set BraspagCarregaDados_CaptureCreditCardTransactionResponse = r_rx
end function



' --------------------------------------------------------------------------------
'   BraspagCarregaDados_VoidCreditCardTransactionResponse
'   Processa o xml de resposta e carrega os dados na estrutura.
function BraspagCarregaDados_VoidCreditCardTransactionResponse(byval rxXml, byref msg_erro)
dim r_rx, objXML
dim blnNodeNotFound, strNodeName, strNodeValue
dim oNode, oNodeErrorList, oNodeSet
dim oTransactionDataCollection, oTransactionSet
dim strTipoRetorno
dim strErrorCode, strErrorMessage
	
	msg_erro = ""
	
	set r_rx = new cl_BRASPAG_VoidCreditCardTransactionResponse_RX
	
	Set objXML = Server.CreateObject("MSXML2.DOMDocument.3.0")
	objXML.Async = False
	objXML.LoadXml(rxXml)
	
	strTipoRetorno = objXML.documentElement.baseName
	if Ucase(strTipoRetorno) <> "ENVELOPE" then
		msg_erro = "Resposta recebida é inválida!" & chr(13)& rxXml
		exit function
		end if
	
	set oTransactionDataCollection=objXML.documentElement.selectNodes("//VoidCreditCardTransactionResponse/VoidCreditCardTransactionResult/TransactionDataCollection")
	if Not oTransactionDataCollection is nothing then
		for each oTransactionSet in oTransactionDataCollection
		'	BraspagTransactionId
			strNodeName = "//TransactionDataResponse/BraspagTransactionId"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_BraspagTransactionId = strNodeValue
			
		'	AcquirerTransactionId
			strNodeName = "//TransactionDataResponse/AcquirerTransactionId"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_AcquirerTransactionId = strNodeValue
			
		'	Amount
			strNodeName = "//TransactionDataResponse/Amount"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_Amount = strNodeValue
			
		'	AuthorizationCode
			strNodeName = "//TransactionDataResponse/AuthorizationCode"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_AuthorizationCode = strNodeValue
			
		'	ReturnCode
			strNodeName = "//TransactionDataResponse/ReturnCode"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ReturnCode = strNodeValue
			
		'	ReturnMessage
			strNodeName = "//TransactionDataResponse/ReturnMessage"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ReturnMessage = strNodeValue
			
		'	Status
			strNodeName = "//TransactionDataResponse/Status"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_Status = strNodeValue
			
		'	ProofOfSale
			strNodeName = "//TransactionDataResponse/ProofOfSale"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ProofOfSale = strNodeValue
			
		'	ServiceTaxAmount
			strNodeName = "//TransactionDataResponse/ServiceTaxAmount"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ServiceTaxAmount = strNodeValue
			next
		end if
	
'	CorrelationId
	strNodeName = "//VoidCreditCardTransactionResponse/VoidCreditCardTransactionResult/CorrelationId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_CorrelationId = strNodeValue
	
'	Success
	strNodeName = "//VoidCreditCardTransactionResponse/VoidCreditCardTransactionResult/Success"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_Success = strNodeValue
	
'	ErrorCode/ErrorMessage
	set oNodeErrorList=objXML.documentElement.selectNodes("//VoidCreditCardTransactionResponse/VoidCreditCardTransactionResult/ErrorReportDataCollection")
	if Not oNodeErrorList is nothing then
		for each oNodeSet in oNodeErrorList
		'	OBTÉM OS DADOS DO ERRO P/ VERIFICAR SE HÁ CONTEÚDO
		'	ErrorCode
			strNodeName = "//ErrorReportDataResponse/ErrorCode"
			strErrorCode = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorCode = ""
			r_rx.PAG_ErrorCode = strErrorCode
			
		'	ErrorMessage
			strNodeName = "//ErrorReportDataResponse/ErrorMessage"
			strErrorMessage = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorMessage = ""
			r_rx.PAG_ErrorMessage = strErrorMessage
			next
		end if
	
'	faultcode
	strNodeName = "//Fault/faultcode"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_faultcode = strNodeValue
	
'	faultstring
	strNodeName = "//Fault/faultstring"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_faultstring = strNodeValue
	
	set BraspagCarregaDados_VoidCreditCardTransactionResponse = r_rx
end function



' --------------------------------------------------------------------------------
'   BraspagProcessaRequisicao_CaptureCreditCardTransaction
'   Executa a requisição e realiza o processamento relacionado ao BD.
function BraspagProcessaRequisicao_CaptureCreditCardTransaction(byval id_pedido_pagto_braspag, byval id_pedido_pagto_braspag_pag, byval usuario, byref trx, byref r_rx, byref msg_erro)
dim t_PP_BRASPAG, t_PP_BRASPAG_PAG, t_PP_BRASPAG_PAG_OP_COMPLEMENTAR, t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML
dim pedido, vl_transacao
dim lngRecordsAffected
dim idPedidoPagtoBraspagPagOpComplementar, idPedidoPagtoBraspagPagOpComplXmlTx, idPedidoPagtoBraspagPagOpComplXmlRx
dim strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount
dim strCapturedDate
dim strSql
dim txXml, rxXml
dim st_sucesso
	
	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(t_PP_BRASPAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG_OP_COMPLEMENTAR, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag & ")"
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	t_PP_BRASPAG.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if t_PP_BRASPAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	pedido = Trim("" & t_PP_BRASPAG("pedido"))
	vl_transacao = t_PP_BRASPAG("valor_transacao")
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG_PAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag_pag & ")"
	if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
	t_PP_BRASPAG_PAG.open strSql, cn
	if t_PP_BRASPAG_PAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag!!"
		exit function
		end if
	
	strMerchantId = Trim("" & t_PP_BRASPAG_PAG("Req_OrderData_MerchantId"))
	strBraspagTransactionId = Trim("" & t_PP_BRASPAG_PAG("Resp_PaymentDataResponse_BraspagTransactionId"))
	strAmount = Trim("" & t_PP_BRASPAG_PAG("Req_PaymentDataCollection_Amount"))
	strServiceTaxAmount = Trim("" & t_PP_BRASPAG_PAG("Req_PaymentDataCollection_ServiceTaxAmount"))
	
	if strBraspagTransactionId = "" then
		msg_erro = "Não é possível consultar a Braspag porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!"
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_CaptureCreditCardTransaction_TX(strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount)
	txXml = BraspagXmlMontaRequisicaoCaptureCreditCardTransaction(trx)
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR, idPedidoPagtoBraspagPagOpComplementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagPagOpComplXmlTx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplXmlTx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplXmlTx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagPagOpComplXmlRx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplXmlRx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplXmlRx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível enviar a solicitação à Braspag devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id_pedido_pagto_braspag") = CLng(id_pedido_pagto_braspag)
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id_pedido_pagto_braspag_pag") = CLng(id_pedido_pagto_braspag_pag)
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("operacao") = BRASPAG_TIPO_TRANSACAO__PAG_CAPTURECREDITCARDTRANSACTION
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("usuario") = usuario
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("request_data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("request_data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_RequestId") = trx.PAG_RequestId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_Version") = trx.PAG_Version
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_MerchantId") = trx.PAG_MerchantId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_BraspagTransactionId") = trx.PAG_BraspagTransactionId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_Amount") = trx.PAG_Amount
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_ServiceTaxAmount") = trx.PAG_ServiceTaxAmount
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Update
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagPagOpComplXmlTx
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_pag_op_complementar") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_CAPTURECREDITCARDTRANSACTION
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("xml") = txXml
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Update
	
	rxXml = BraspagEnviaTransacao(txXml, BRASPAG_WS_ENDERECO_PAGADOR_TRANSACTION)
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR WHERE (id = " & idPedidoPagtoBraspagPagOpComplementar & ")"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Open strSql, cn
	if Not t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Eof then
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("response_data") = Date
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("response_data_hora") = Now
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("st_resposta_recebida") = 1
		if Trim(rxXml) = "" then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("st_resposta_vazia") = 1
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Update
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagPagOpComplXmlRx
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_pag_op_complementar") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_CAPTURECREDITCARDTRANSACTION
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("xml") = rxXml
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Update
	
	set r_rx = BraspagCarregaDados_CaptureCreditCardTransactionResponse(rxXml, msg_erro)
	if msg_erro <> "" then exit function
	
'	ATUALIZA O ÚLTIMO STATUS DA TRANSAÇÃO
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_CAPTURECREDITCARDTRANSACTIONRESPONSE_STATUS__CAPTURE_CONFIRMED then
		strCapturedDate = bd_monta_data_hora(Now)
		strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG SET" & _
					" ult_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "'," & _
					" ult_PAG_CapturedDate = " & strCapturedDate & "," & _
					" ult_PAG_atualizacao_data_hora = getdate()," & _
					" ult_PAG_atualizacao_usuario = '" & usuario & "'," & _
					" ult_id_pedido_pagto_braspag_pag_op_complementar = " & idPedidoPagtoBraspagPagOpComplementar & _
				" WHERE" & _
					" (id = " & id_pedido_pagto_braspag & ")"
	else
		strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG SET" & _
					" ult_id_pedido_pagto_braspag_pag_op_complementar = " & idPedidoPagtoBraspagPagOpComplementar & _
				" WHERE" & _
					" (id = " & id_pedido_pagto_braspag & ")"
		end if
	cn.Execute strSql, lngRecordsAffected
	
	st_sucesso = 0
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_CAPTURECREDITCARDTRANSACTIONRESPONSE_STATUS__CAPTURE_CONFIRMED then st_sucesso = 1
	strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR SET" & _
				" Resp_AuthorizationCode = '" & r_rx.PAG_AuthorizationCode & "'," & _
				" Resp_ProofOfSale = '" & r_rx.PAG_ProofOfSale & "'," & _
				" st_sucesso = " & CStr(st_sucesso) & "," & _
				" Resp_CaptureCreditCardTransactionResponse_Status = '" & r_rx.PAG_Status & "'" & _
			" WHERE" & _
				" (id = " & idPedidoPagtoBraspagPagOpComplementar & ")"
	cn.Execute strSql, lngRecordsAffected
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	set t_PP_BRASPAG = nothing

	if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
	set t_PP_BRASPAG_PAG = nothing

	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	set t_PP_BRASPAG_PAG_OP_COMPLEMENTAR = nothing

	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	set t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML = nothing
	
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_CAPTURECREDITCARDTRANSACTIONRESPONSE_STATUS__CAPTURE_CONFIRMED then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if BraspagRegistraPagtoNoPedido(BRASPAG_REGISTRA_PAGTO__OP_CAPTURA, pedido, id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, converte_numero(vl_transacao), usuario, msg_erro) then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			msg_erro = "Falha ao tentar registrar automaticamente o pagamento no pedido!!" & chr(13) & msg_erro
			exit function
			end if
		end if
end function



' --------------------------------------------------------------------------------
'   BraspagProcessaRequisicao_VoidCreditCardTransaction
'   Executa a requisição e realiza o processamento relacionado ao BD.
function BraspagProcessaRequisicao_VoidCreditCardTransaction(byval id_pedido_pagto_braspag, byval id_pedido_pagto_braspag_pag, byval usuario, byref trx, byref r_rx, byref msg_erro)
dim t_PP_BRASPAG, t_PP_BRASPAG_PAG, t_PP_BRASPAG_PAG_OP_COMPLEMENTAR, t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML
dim pedido, vl_transacao
dim lngRecordsAffected
dim idPedidoPagtoBraspagPagOpComplementar, idPedidoPagtoBraspagPagOpComplXmlTx, idPedidoPagtoBraspagPagOpComplXmlRx
dim strVoidedDate
dim strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount
dim strSql
dim txXml, rxXml
dim st_sucesso
	
	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(t_PP_BRASPAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG_OP_COMPLEMENTAR, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag & ")"
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	t_PP_BRASPAG.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if t_PP_BRASPAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	pedido = Trim("" & t_PP_BRASPAG("pedido"))
	vl_transacao = t_PP_BRASPAG("valor_transacao")
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG_PAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag_pag & ")"
	if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
	t_PP_BRASPAG_PAG.open strSql, cn
	if t_PP_BRASPAG_PAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag!!"
		exit function
		end if
	
	strMerchantId = Trim("" & t_PP_BRASPAG_PAG("Req_OrderData_MerchantId"))
	strBraspagTransactionId = Trim("" & t_PP_BRASPAG_PAG("Resp_PaymentDataResponse_BraspagTransactionId"))
	strAmount = Trim("" & t_PP_BRASPAG_PAG("Req_PaymentDataCollection_Amount"))
	strServiceTaxAmount = Trim("" & t_PP_BRASPAG_PAG("Req_PaymentDataCollection_ServiceTaxAmount"))
	
	if strBraspagTransactionId = "" then
		msg_erro = "Não é possível consultar a Braspag porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!"
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_VoidCreditCardTransaction_TX(strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount)
	txXml = BraspagXmlMontaRequisicaoVoidCreditCardTransaction(trx)
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR, idPedidoPagtoBraspagPagOpComplementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagPagOpComplXmlTx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplXmlTx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplXmlTx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagPagOpComplXmlRx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplXmlRx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplXmlRx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível enviar a solicitação à Braspag devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id_pedido_pagto_braspag") = CLng(id_pedido_pagto_braspag)
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id_pedido_pagto_braspag_pag") = CLng(id_pedido_pagto_braspag_pag)
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("operacao") = BRASPAG_TIPO_TRANSACAO__PAG_VOIDCREDITCARDTRANSACTION
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("usuario") = usuario
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("request_data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("request_data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_RequestId") = trx.PAG_RequestId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_Version") = trx.PAG_Version
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_MerchantId") = trx.PAG_MerchantId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_BraspagTransactionId") = trx.PAG_BraspagTransactionId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_Amount") = trx.PAG_Amount
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_ServiceTaxAmount") = trx.PAG_ServiceTaxAmount
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Update
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagPagOpComplXmlTx
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_pag_op_complementar") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_VOIDCREDITCARDTRANSACTION
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("xml") = txXml
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Update
	
	rxXml = BraspagEnviaTransacao(txXml, BRASPAG_WS_ENDERECO_PAGADOR_TRANSACTION)
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR WHERE (id = " & idPedidoPagtoBraspagPagOpComplementar & ")"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Open strSql, cn
	if Not t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Eof then
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("response_data") = Date
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("response_data_hora") = Now
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("st_resposta_recebida") = 1
		if Trim(rxXml) = "" then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("st_resposta_vazia") = 1
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Update
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagPagOpComplXmlRx
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_pag_op_complementar") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_VOIDCREDITCARDTRANSACTION
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("xml") = rxXml
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Update
	
	set r_rx = BraspagCarregaDados_VoidCreditCardTransactionResponse(rxXml, msg_erro)
	if msg_erro <> "" then exit function
	
'	ATUALIZA O ÚLTIMO STATUS DA TRANSAÇÃO
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_VOIDCREDITCARDTRANSACTIONRESPONSE_STATUS__VOID_CONFIRMED then
		strVoidedDate = bd_monta_data_hora(Now)
		strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG SET" & _
					" ult_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURA_CANCELADA & "'," & _
					" ult_PAG_VoidedDate = " & strVoidedDate & "," & _
					" ult_PAG_atualizacao_data_hora = getdate()," & _
					" ult_PAG_atualizacao_usuario = '" & usuario & "'," & _
					" ult_id_pedido_pagto_braspag_pag_op_complementar = " & idPedidoPagtoBraspagPagOpComplementar & _
				" WHERE" & _
					" (id = " & id_pedido_pagto_braspag & ")"
	else
		strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG SET" & _
					" ult_id_pedido_pagto_braspag_pag_op_complementar = " & idPedidoPagtoBraspagPagOpComplementar & _
				" WHERE" & _
					" (id = " & id_pedido_pagto_braspag & ")"
		end if
	cn.Execute strSql, lngRecordsAffected
	
	st_sucesso = 0
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_VOIDCREDITCARDTRANSACTIONRESPONSE_STATUS__VOID_CONFIRMED then st_sucesso = 1
	strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR SET" & _
				" Resp_AuthorizationCode = '" & r_rx.PAG_AuthorizationCode & "'," & _
				" Resp_ProofOfSale = '" & r_rx.PAG_ProofOfSale & "'," & _
				" st_sucesso = " & CStr(st_sucesso) & "," & _
				" Resp_VoidCreditCardTransactionResponse_Status = '" & r_rx.PAG_Status & "'" & _
			" WHERE" & _
				" (id = " & idPedidoPagtoBraspagPagOpComplementar & ")"
	cn.Execute strSql, lngRecordsAffected
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	set t_PP_BRASPAG = nothing

	if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
	set t_PP_BRASPAG_PAG = nothing

	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	set t_PP_BRASPAG_PAG_OP_COMPLEMENTAR = nothing

	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	set t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML = nothing
	
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_VOIDCREDITCARDTRANSACTIONRESPONSE_STATUS__VOID_CONFIRMED then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if BraspagRegistraPagtoNoPedido(BRASPAG_REGISTRA_PAGTO__OP_CANCELAMENTO, pedido, id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, converte_numero(vl_transacao), usuario, msg_erro) then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			msg_erro = "Falha ao tentar registrar automaticamente o cancelamento do pagamento no pedido!!" & chr(13) & msg_erro
			exit function
			end if
		end if
end function



' --------------------------------------------------------------------------------
'   decodifica_PaymentDataResponseStatus_para_GlobalStatus
'   Decodifica o código de status retornado em PaymentDataResponse para um
'   código global.
'   Caso seja informado um código de status desconhecido, o mesmo será retornado
'   com a seguinte formatação: 'PGnnn'
'        'PG' = Payment
'        'nnn' = código do status desconhecido formatado c/ zeros à esquerda
function decodifica_PaymentDataResponseStatus_para_GlobalStatus(byval codigoStatus)
dim strResp
	strResp = ""
	codigoStatus = Trim("" & codigoStatus)
	
	select case codigoStatus
		case BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__CAPTURADA
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA
		
		case BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__AUTORIZADA
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA
		
		case BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__NAO_AUTORIZADA
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__NAO_AUTORIZADA
		
		case BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__ERRO_DESQUALIFICANTE
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ERRO_DESQUALIFICANTE
		
		case BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__AGUARDANDO_RESPOSTA
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AGUARDANDO_RESPOSTA
		
		case else
		'	CÓDIGO DESCONHECIDO
			strResp = codigoStatus
			do while Len(strResp) < 3: strResp = "0" & strResp: loop
			strResp = "PG" & strResp
		end select
	
	decodifica_PaymentDataResponseStatus_para_GlobalStatus = strResp
end function



' --------------------------------------------------------------------------------
'   decodifica_GetTransactionDataResponseStatus_para_GlobalStatus
'   Decodifica o código de status retornado em GetTransactionDataResponse para um
'   código global.
'   Caso seja informado um código de status desconhecido, o mesmo será retornado
'   com a seguinte formatação: 'QYnnn'
'        'QY' = Query
'        'nnn' = código do status desconhecido formatado c/ zeros à esquerda
function decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(byval codigoStatus)
dim strResp
	strResp = ""
	codigoStatus = Trim("" & codigoStatus)
	
	select case codigoStatus
		case BRASPAG_PAGADOR_CARTAO_GETTRANSACTIONDATARESPONSE_STATUS__INDEFINIDA
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__INDEFINIDA
		
		case BRASPAG_PAGADOR_CARTAO_GETTRANSACTIONDATARESPONSE_STATUS__CAPTURADA
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA
		
		case BRASPAG_PAGADOR_CARTAO_GETTRANSACTIONDATARESPONSE_STATUS__AUTORIZADA
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA
		
		case BRASPAG_PAGADOR_CARTAO_GETTRANSACTIONDATARESPONSE_STATUS__NAO_AUTORIZADA
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__NAO_AUTORIZADA
		
		case BRASPAG_PAGADOR_CARTAO_GETTRANSACTIONDATARESPONSE_STATUS__CAPTURA_CANCELADA
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURA_CANCELADA
		
		case BRASPAG_PAGADOR_CARTAO_GETTRANSACTIONDATARESPONSE_STATUS__ESTORNADA
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNADA
		
		case BRASPAG_PAGADOR_CARTAO_GETTRANSACTIONDATARESPONSE_STATUS__AGUARDANDO_RESPOSTA
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AGUARDANDO_RESPOSTA
		
		case BRASPAG_PAGADOR_CARTAO_GETTRANSACTIONDATARESPONSE_STATUS__ERRO_DESQUALIFICANTE
			strResp = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ERRO_DESQUALIFICANTE
		
		case else
		'	CÓDIGO DESCONHECIDO
			strResp = codigoStatus
			do while Len(strResp) < 3: strResp = "0" & strResp: loop
			strResp = "QY" & strResp
		end select
	
	decodifica_GetTransactionDataResponseStatus_para_GlobalStatus = strResp
end function



' --------------------------------------------------------------------------------
'   decodifica_FraudAnalysisTransactionDetailsResponseAntiFraudTransactionStatusCode_para_GlobalStatus
'   Decodifica o código de status retornado em FraudAnalysisTransactionDetailsResponse/FraudAnalysisTransactionDetailsResult/AntiFraudTransactionStatusCode
'   para um código global.
'   A requisição 'FraudAnalysisTransactionDetails' é para consultar uma transação de análise de fraude.
'   Caso seja informado um código de status desconhecido, o mesmo será retornado
'   com a seguinte formatação: 'FDnnn'
'        'FD' = Fraud Analysis Transaction Details
'        'nnn' = código do status desconhecido formatado c/ zeros à esquerda
function decodifica_FraudAnalysisTransactionDetailsResponseAntiFraudTransactionStatusCode_para_GlobalStatus(byval codigoStatus)
dim strResp
	strResp = ""
	codigoStatus = Trim("" & codigoStatus)
	
	select case codigoStatus
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISTRANSACTIONDETAILSRESPONSE_ANTIFRAUDTRANSACTIONSTATUSCODE__STARTED
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__STARTED
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISTRANSACTIONDETAILSRESPONSE_ANTIFRAUDTRANSACTIONSTATUSCODE__ACCEPT
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ACCEPT
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISTRANSACTIONDETAILSRESPONSE_ANTIFRAUDTRANSACTIONSTATUSCODE__REVIEW
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISTRANSACTIONDETAILSRESPONSE_ANTIFRAUDTRANSACTIONSTATUSCODE__REJECT
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REJECT
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISTRANSACTIONDETAILSRESPONSE_ANTIFRAUDTRANSACTIONSTATUSCODE__PENDENT
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__PENDENT
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISTRANSACTIONDETAILSRESPONSE_ANTIFRAUDTRANSACTIONSTATUSCODE__UNFINISHED
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__UNFINISHED
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISTRANSACTIONDETAILSRESPONSE_ANTIFRAUDTRANSACTIONSTATUSCODE__ABORTED
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ABORTED
		
		case else
		'	CÓDIGO DESCONHECIDO
			strResp = codigoStatus
			do while Len(strResp) < 3: strResp = "0" & strResp: loop
			strResp = "FD" & strResp
		end select
	
	decodifica_FraudAnalysisTransactionDetailsResponseAntiFraudTransactionStatusCode_para_GlobalStatus = strResp
end function



' --------------------------------------------------------------------------------
'   decodifica_FraudAnalysisResponseTransactionStatusCode_para_GlobalStatus
'   Decodifica o código de status retornado em FraudAnalysisResponse/FraudAnalysisResult/TransactionStatusCode
'   para um código global.
'   A requisição 'FraudAnalysis' é de solicitação de análise de fraude.
'   Caso seja informado um código de status desconhecido, o mesmo será retornado
'   com a seguinte formatação: 'FAnnn'
'        'FA' = Fraud Analysis
'        'nnn' = código do status desconhecido formatado c/ zeros à esquerda
function decodifica_FraudAnalysisResponseTransactionStatusCode_para_GlobalStatus(byval codigoStatus)
dim strResp
	strResp = ""
	codigoStatus = Trim("" & codigoStatus)
	
	select case codigoStatus
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__STARTED
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__STARTED
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__ACCEPT
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ACCEPT
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__REVIEW
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__REJECT
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REJECT
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__PENDENT
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__PENDENT
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__UNFINISHED
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__UNFINISHED
		
		case BRASPAG_ANTIFRAUDE_CARTAO_FRAUDANALYSISRESPONSE_TRANSACTIONSTATUSCODE__ABORTED
			strResp = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ABORTED
		
		case else
		'	CÓDIGO DESCONHECIDO
			strResp = codigoStatus
			do while Len(strResp) < 3: strResp = "0" & strResp: loop
			strResp = "FA" & strResp
		end select
	
	decodifica_FraudAnalysisResponseTransactionStatusCode_para_GlobalStatus = strResp
end function



' ------------------------------------------------------------------------
'   BraspagXmlMontaRequisicaoRefundCreditCardTransaction
function BraspagXmlMontaRequisicaoRefundCreditCardTransaction(ByRef trx)
dim xml
	xml =	"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & chr(13) & _
			"	<soap:Body>" & chr(13) & _
			"		<RefundCreditCardTransaction xmlns=""https://www.pagador.com.br/webservice/pagador"">" & chr(13) & _
			"			<request>" & chr(13) & _
			"				<Version>" & trx.PAG_Version & "</Version>" & chr(13) & _
			"				<RequestId>" & trx.PAG_RequestId & "</RequestId>" & chr(13) & _
			"				<MerchantId>" & trx.PAG_MerchantId & "</MerchantId>" & chr(13) & _
			"				<TransactionDataCollection>" & chr(13) & _
			"					<TransactionDataRequest>" & chr(13) & _
			"						<BraspagTransactionId>" & trx.PAG_BraspagTransactionId & "</BraspagTransactionId>" & chr(13) & _
			"						<Amount>" & trx.PAG_Amount & "</Amount>" & chr(13) & _
			"						<ServiceTaxAmount>" & trx.PAG_ServiceTaxAmount & "</ServiceTaxAmount>" & chr(13) & _
			"					</TransactionDataRequest>" & chr(13) & _
			"				</TransactionDataCollection>" & chr(13) & _
			"			</request>" & chr(13) & _
			"		</RefundCreditCardTransaction>" & chr(13) & _
			"	</soap:Body>" & chr(13) & _
			"</soap:Envelope>" & chr(13)
	BraspagXmlMontaRequisicaoRefundCreditCardTransaction = xml
end function



' --------------------------------------------------------------------------------
'   cria_instancia_cl_BRASPAG_RefundCreditCardTransaction_TX
function cria_instancia_cl_BRASPAG_RefundCreditCardTransaction_TX(byval strMerchantId, byval strBraspagTransactionId, byval strAmount, byval strServiceTaxAmount)
dim trx
	msg_erro = ""
	set trx = new cl_BRASPAG_RefundCreditCardTransaction_TX
	trx.PAG_Version = BRASPAG_PAGADOR_VERSION
	trx.PAG_RequestId = Lcase(gera_uid)
	trx.PAG_MerchantId = strMerchantId
	trx.PAG_BraspagTransactionId = strBraspagTransactionId
	trx.PAG_Amount = Trim("" & strAmount)
	trx.PAG_ServiceTaxAmount = Trim("" & strServiceTaxAmount)
	set cria_instancia_cl_BRASPAG_RefundCreditCardTransaction_TX = trx
end function



' --------------------------------------------------------------------------------
'   BraspagCarregaDados_RefundCreditCardTransactionResponse
'   Processa o xml de resposta e carrega os dados na estrutura.
function BraspagCarregaDados_RefundCreditCardTransactionResponse(byval rxXml, byref msg_erro)
dim r_rx, objXML
dim blnNodeNotFound, strNodeName, strNodeValue
dim oNode, oNodeErrorList, oNodeSet
dim oTransactionDataCollection, oTransactionSet
dim strTipoRetorno
dim strErrorCode, strErrorMessage
	
	msg_erro = ""
	
	set r_rx = new cl_BRASPAG_RefundCreditCardTransactionResponse_RX
	
	Set objXML = Server.CreateObject("MSXML2.DOMDocument.3.0")
	objXML.Async = False
	objXML.LoadXml(rxXml)
	
	strTipoRetorno = objXML.documentElement.baseName
	if Ucase(strTipoRetorno) <> "ENVELOPE" then
		msg_erro = "Resposta recebida é inválida!" & chr(13)& rxXml
		exit function
		end if
	
	set oTransactionDataCollection=objXML.documentElement.selectNodes("//RefundCreditCardTransactionResponse/RefundCreditCardTransactionResult/TransactionDataCollection")
	if Not oTransactionDataCollection is nothing then
		for each oTransactionSet in oTransactionDataCollection
		'	BraspagTransactionId
			strNodeName = "//TransactionDataResponse/BraspagTransactionId"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_BraspagTransactionId = strNodeValue
			
		'	AcquirerTransactionId
			strNodeName = "//TransactionDataResponse/AcquirerTransactionId"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_AcquirerTransactionId = strNodeValue
			
		'	Amount
			strNodeName = "//TransactionDataResponse/Amount"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_Amount = strNodeValue
			
		'	AuthorizationCode
			strNodeName = "//TransactionDataResponse/AuthorizationCode"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_AuthorizationCode = strNodeValue
			
		'	ReturnCode
			strNodeName = "//TransactionDataResponse/ReturnCode"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ReturnCode = strNodeValue
			
		'	ReturnMessage
			strNodeName = "//TransactionDataResponse/ReturnMessage"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ReturnMessage = strNodeValue
			
		'	Status
			strNodeName = "//TransactionDataResponse/Status"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_Status = strNodeValue
			
		'	ProofOfSale
			strNodeName = "//TransactionDataResponse/ProofOfSale"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ProofOfSale = strNodeValue
			
		'	ServiceTaxAmount
			strNodeName = "//TransactionDataResponse/ServiceTaxAmount"
			strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			r_rx.PAG_ServiceTaxAmount = strNodeValue
			next
		end if
	
'	CorrelationId
	strNodeName = "//RefundCreditCardTransactionResponse/RefundCreditCardTransactionResult/CorrelationId"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_CorrelationId = strNodeValue
	
'	Success
	strNodeName = "//RefundCreditCardTransactionResponse/RefundCreditCardTransactionResult/Success"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_Success = strNodeValue
	
'	ErrorCode/ErrorMessage
	set oNodeErrorList=objXML.documentElement.selectNodes("//RefundCreditCardTransactionResponse/RefundCreditCardTransactionResult/ErrorReportDataCollection")
	if Not oNodeErrorList is nothing then
		for each oNodeSet in oNodeErrorList
		'	OBTÉM OS DADOS DO ERRO P/ VERIFICAR SE HÁ CONTEÚDO
		'	ErrorCode
			strNodeName = "//ErrorReportDataResponse/ErrorCode"
			strErrorCode = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorCode = ""
			r_rx.PAG_ErrorCode = strErrorCode
			
		'	ErrorMessage
			strNodeName = "//ErrorReportDataResponse/ErrorMessage"
			strErrorMessage = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
			if blnNodeNotFound then strErrorMessage = ""
			r_rx.PAG_ErrorMessage = strErrorMessage
			next
		end if
	
'	faultcode
	strNodeName = "//Fault/faultcode"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_faultcode = strNodeValue
	
'	faultstring
	strNodeName = "//Fault/faultstring"
	strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
	r_rx.PAG_faultstring = strNodeValue
	
	set BraspagCarregaDados_RefundCreditCardTransactionResponse = r_rx
end function



' --------------------------------------------------------------------------------
'   BraspagProcessaRequisicao_RefundCreditCardTransaction
'   Executa a requisição e realiza o processamento relacionado ao BD.
function BraspagProcessaRequisicao_RefundCreditCardTransaction(byval id_pedido_pagto_braspag, byval id_pedido_pagto_braspag_pag, byval usuario, byref trx, byref r_rx, byref msg_erro)
dim t_PP_BRASPAG, t_PP_BRASPAG_PAG, t_PP_BRASPAG_PAG_OP_COMPLEMENTAR, t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML
dim pedido, vl_transacao
dim lngRecordsAffected
dim idPedidoPagtoBraspagPagOpComplementar, idPedidoPagtoBraspagPagOpComplXmlTx, idPedidoPagtoBraspagPagOpComplXmlRx
dim strVoidedDate
dim strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount
dim strSql
dim txXml, rxXml
dim st_sucesso
	
	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(t_PP_BRASPAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG_OP_COMPLEMENTAR, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag & ")"
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	t_PP_BRASPAG.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if t_PP_BRASPAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	pedido = Trim("" & t_PP_BRASPAG("pedido"))
	vl_transacao = t_PP_BRASPAG("valor_transacao")
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG_PAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag_pag & ")"
	if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
	t_PP_BRASPAG_PAG.open strSql, cn
	if t_PP_BRASPAG_PAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag!!"
		exit function
		end if
	
	strMerchantId = Trim("" & t_PP_BRASPAG_PAG("Req_OrderData_MerchantId"))
	strBraspagTransactionId = Trim("" & t_PP_BRASPAG_PAG("Resp_PaymentDataResponse_BraspagTransactionId"))
	strAmount = Trim("" & t_PP_BRASPAG_PAG("Req_PaymentDataCollection_Amount"))
	strServiceTaxAmount = Trim("" & t_PP_BRASPAG_PAG("Req_PaymentDataCollection_ServiceTaxAmount"))
	
	if strBraspagTransactionId = "" then
		msg_erro = "Não é possível consultar a Braspag porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!"
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_RefundCreditCardTransaction_TX(strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount)
	txXml = BraspagXmlMontaRequisicaoRefundCreditCardTransaction(trx)
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR, idPedidoPagtoBraspagPagOpComplementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagPagOpComplXmlTx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplXmlTx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplXmlTx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagPagOpComplXmlRx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagPagOpComplXmlRx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagPagOpComplXmlRx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível enviar a solicitação à Braspag devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id_pedido_pagto_braspag") = CLng(id_pedido_pagto_braspag)
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("id_pedido_pagto_braspag_pag") = CLng(id_pedido_pagto_braspag_pag)
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("operacao") = BRASPAG_TIPO_TRANSACAO__PAG_REFUNDCREDITCARDTRANSACTION
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("usuario") = usuario
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("request_data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("request_data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_RequestId") = trx.PAG_RequestId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_Version") = trx.PAG_Version
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_MerchantId") = trx.PAG_MerchantId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_BraspagTransactionId") = trx.PAG_BraspagTransactionId
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_Amount") = trx.PAG_Amount
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("Req_ServiceTaxAmount") = trx.PAG_ServiceTaxAmount
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Update
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagPagOpComplXmlTx
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_pag_op_complementar") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_REFUNDCREDITCARDTRANSACTION
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("xml") = txXml
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Update
	
	rxXml = BraspagEnviaTransacao(txXml, BRASPAG_WS_ENDERECO_PAGADOR_TRANSACTION)
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR WHERE (id = " & idPedidoPagtoBraspagPagOpComplementar & ")"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Open strSql, cn
	if Not t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Eof then
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("response_data") = Date
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("response_data_hora") = Now
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("st_resposta_recebida") = 1
		if Trim(rxXml) = "" then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR("st_resposta_vazia") = 1
		t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Update
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagPagOpComplXmlRx
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_pag_op_complementar") = idPedidoPagtoBraspagPagOpComplementar
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_REFUNDCREDITCARDTRANSACTION
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML("xml") = rxXml
	t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Update
	
	set r_rx = BraspagCarregaDados_RefundCreditCardTransactionResponse(rxXml, msg_erro)
	if msg_erro <> "" then exit function
	
'	ATUALIZA O ÚLTIMO STATUS DA TRANSAÇÃO
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_REFUNDCREDITCARDTRANSACTIONRESPONSE_STATUS__REFUND_CONFIRMED then
		strVoidedDate = bd_monta_data_hora(Now)
		strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG SET" & _
					" ult_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNADA & "'," & _
					" ult_PAG_VoidedDate = " & strVoidedDate & "," & _
					" ult_PAG_atualizacao_data_hora = getdate()," & _
					" ult_PAG_atualizacao_usuario = '" & usuario & "'," & _
					" ult_id_pedido_pagto_braspag_pag_op_complementar = " & idPedidoPagtoBraspagPagOpComplementar & _
				" WHERE" & _
					" (id = " & id_pedido_pagto_braspag & ")"
	else
		strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG SET" & _
					" ult_id_pedido_pagto_braspag_pag_op_complementar = " & idPedidoPagtoBraspagPagOpComplementar & _
				" WHERE" & _
					" (id = " & id_pedido_pagto_braspag & ")"
		end if
	cn.Execute strSql, lngRecordsAffected
	
	st_sucesso = 0
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_REFUNDCREDITCARDTRANSACTIONRESPONSE_STATUS__REFUND_CONFIRMED then st_sucesso = 1
	strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG_PAG_OP_COMPLEMENTAR SET" & _
				" Resp_AuthorizationCode = '" & r_rx.PAG_AuthorizationCode & "'," & _
				" Resp_ProofOfSale = '" & r_rx.PAG_ProofOfSale & "'," & _
				" st_sucesso = " & CStr(st_sucesso) & "," & _
				" Resp_RefundCreditCardTransactionResponse_Status = '" & r_rx.PAG_Status & "'" & _
			" WHERE" & _
				" (id = " & idPedidoPagtoBraspagPagOpComplementar & ")"
	cn.Execute strSql, lngRecordsAffected
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	set t_PP_BRASPAG = nothing

	if t_PP_BRASPAG_PAG.State <> 0 then t_PP_BRASPAG_PAG.Close
	set t_PP_BRASPAG_PAG = nothing

	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR.Close
	set t_PP_BRASPAG_PAG_OP_COMPLEMENTAR = nothing

	if t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML.Close
	set t_PP_BRASPAG_PAG_OP_COMPLEMENTAR_XML = nothing
	
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_REFUNDCREDITCARDTRANSACTIONRESPONSE_STATUS__REFUND_CONFIRMED then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if BraspagRegistraPagtoNoPedido(BRASPAG_REGISTRA_PAGTO__OP_ESTORNO, pedido, id_pedido_pagto_braspag, id_pedido_pagto_braspag_pag, converte_numero(vl_transacao), usuario, msg_erro) then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			msg_erro = "Falha ao tentar registrar automaticamente o estorno do pagamento no pedido!!" & chr(13) & msg_erro
			exit function
			end if
		end if
end function



' --------------------------------------------------------------------------------
'   BraspagProcessaRequisicao_AF_UpdateStatus
'   Executa a requisição e realiza o processamento relacionado ao BD.
function BraspagProcessaRequisicao_AF_UpdateStatus(byval af_decision, byval id_pedido_pagto_braspag, byval id_pedido_pagto_braspag_af, byval usuario, byval af_comentario, byref trx, byref r_rx, byref msg_erro)
dim t_PP_BRASPAG, t_PP_BRASPAG_AF, t_PP_BRASPAG_AF_OP_COMPLEMENTAR, t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML
dim lngRecordsAffected
dim idPedidoPagtoBraspagAfOpComplementar, idPedidoPagtoBraspagAfOpComplXmlTx, idPedidoPagtoBraspagAfOpComplXmlRx
dim strMerchantId, strAntiFraudTransactionId
dim strSql
dim txXml, rxXml
	
	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(t_PP_BRASPAG, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_AF, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_AF_OP_COMPLEMENTAR, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG" & _
			" WHERE" & _
				" (id = " & id_pedido_pagto_braspag & ")"
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	t_PP_BRASPAG.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if t_PP_BRASPAG.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG_AF" & _
			" WHERE" & _
				" (id_pedido_pagto_braspag = " & id_pedido_pagto_braspag_af & ")"
	if t_PP_BRASPAG_AF.State <> 0 then t_PP_BRASPAG_AF.Close
	t_PP_BRASPAG_AF.open strSql, cn
	if t_PP_BRASPAG_AF.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Antifraude!!"
		exit function
		end if
	
	strMerchantId = Trim("" & t_PP_BRASPAG_AF("Req_MerchantId"))
	strAntiFraudTransactionId = Trim("" & t_PP_BRASPAG_AF("Resp_AntiFraudTransactionId"))
	
	if strAntiFraudTransactionId = "" then
		msg_erro = "Não é possível consultar os dados no Antifraude porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!"
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_AF_UpdateStatus_TX(strMerchantId, strAntiFraudTransactionId, af_decision, af_comentario)
	txXml = BraspagXmlMontaRequisicaoAfUpdateStatus(trx)
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR, idPedidoPagtoBraspagAfOpComplementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagAfOpComplementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagAfOpComplementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagAfOpComplXmlTx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagAfOpComplXmlTx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagAfOpComplXmlTx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR_XML, idPedidoPagtoBraspagAfOpComplXmlRx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif idPedidoPagtoBraspagAfOpComplXmlRx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & idPedidoPagtoBraspagAfOpComplXmlRx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível enviar a requisição para o Antifraude devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR WHERE (id = -1)"
	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Open strSql, cn
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR.AddNew
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("id") = idPedidoPagtoBraspagAfOpComplementar
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("id_pedido_pagto_braspag") = CLng(id_pedido_pagto_braspag)
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("id_pedido_pagto_braspag_af") = CLng(id_pedido_pagto_braspag_af)
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("operacao") = BRASPAG_TIPO_TRANSACAO__AF_UPDATE_STATUS
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("usuario") = usuario
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("request_data") = Date
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("request_data_hora") = Now
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("Req_RequestId") = trx.AF_RequestId
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("Req_Version") = trx.AF_Version
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("Req_MerchantId") = trx.AF_MerchantId
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("Req_AntiFraudTransactionId") = trx.AF_AntiFraudTransactionId
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("Req_NewStatus") = trx.AF_NewStatus
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR("Req_Comment") = trx.AF_Comment
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Update
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagAfOpComplXmlTx
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_af_op_complementar") = idPedidoPagtoBraspagAfOpComplementar
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__AF_UPDATE_STATUS
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("xml") = txXml
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Update
	
	rxXml = BraspagEnviaTransacao(txXml, BRASPAG_WS_ENDERECO_ANTIFRAUDE_TRANSACTION)
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR WHERE (id = " & idPedidoPagtoBraspagAfOpComplementar & ")"
	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Close
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Open strSql, cn
	if Not t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Eof then
		t_PP_BRASPAG_AF_OP_COMPLEMENTAR("response_data") = Date
		t_PP_BRASPAG_AF_OP_COMPLEMENTAR("response_data_hora") = Now
		t_PP_BRASPAG_AF_OP_COMPLEMENTAR("st_resposta_recebida") = 1
		if Trim(rxXml) = "" then t_PP_BRASPAG_AF_OP_COMPLEMENTAR("st_resposta_vazia") = 1
		t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Update
		end if
	
	strSql = "SELECT * FROM t_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Close
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Open strSql, cn
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.AddNew
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("id") = idPedidoPagtoBraspagAfOpComplXmlRx
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("id_pedido_pagto_braspag_af_op_complementar") = idPedidoPagtoBraspagAfOpComplementar
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("data") = Date
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("data_hora") = Now
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__AF_UPDATE_STATUS
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML("xml") = rxXml
	t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Update
	
	set r_rx = BraspagCarregaDados_AF_UpdateStatusResponse(rxXml, msg_erro)
	if msg_erro <> "" then exit function
	
'	ATUALIZA O ÚLTIMO STATUS DA TRANSAÇÃO
	if r_rx.AF_RequestStatusCode = BRASPAG_ANTIFRAUDE_CARTAO_UPDATESTATUSRESPONSE_REQUESTSTATUSCODE__SUCCESS then
		strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG SET"
		
		if af_decision = BRASPAG_AF_DECISION__ACCEPT then
			strSql = strSql & _
					" ult_AF_GlobalStatus = '" & BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ACCEPT & "',"
		elseif af_decision = BRASPAG_AF_DECISION__REJECT then
			strSql = strSql & _
					" ult_AF_GlobalStatus = '" & BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REJECT & "',"
			end if
		
		strSql = strSql & _
					" ult_AF_atualizacao_data_hora = getdate()," & _
					" ult_AF_atualizacao_usuario = '" & usuario & "'," & _
					" AF_review_tratado_status = 1," & _
					" AF_review_tratado_usuario = '" & usuario & "'," & _
					" AF_review_tratado_data = " & bd_monta_data(Date) & "," & _
					" AF_review_tratado_data_hora = " & bd_monta_data_hora(Now) & "," & _
					" AF_review_tratado_decision = '" & af_decision & "'," & _
					" AF_review_tratado_comentario = '" & QuotedStr(af_comentario) & "'," & _
					" ult_id_pedido_pagto_braspag_af_op_complementar = " & idPedidoPagtoBraspagAfOpComplementar & _
				" WHERE" & _
					" (id = " & id_pedido_pagto_braspag & ")"
	else
		strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG SET" & _
					" ult_id_pedido_pagto_braspag_af_op_complementar = " & idPedidoPagtoBraspagAfOpComplementar & _
				" WHERE" & _
					" (id = " & id_pedido_pagto_braspag & ")"
		end if
	cn.Execute strSql, lngRecordsAffected
	
	strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG_AF_OP_COMPLEMENTAR SET" & _
				" st_sucesso = 1, " & _
				" Resp_UpdateStatusResponse_RequestStatusCode = '" & r_rx.AF_RequestStatusCode & "'" & _
			" WHERE" & _
				" (id = " & idPedidoPagtoBraspagAfOpComplementar & ")"
	cn.Execute strSql, lngRecordsAffected
	
'	FECHA TABELAS
	if t_PP_BRASPAG.State <> 0 then t_PP_BRASPAG.Close
	set t_PP_BRASPAG = nothing

	if t_PP_BRASPAG_AF.State <> 0 then t_PP_BRASPAG_AF.Close
	set t_PP_BRASPAG_AF = nothing

	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR.Close
	set t_PP_BRASPAG_AF_OP_COMPLEMENTAR = nothing

	if t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.State <> 0 then t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML.Close
	set t_PP_BRASPAG_AF_OP_COMPLEMENTAR_XML = nothing
end function



' --------------------------------------------------------------------------------
'   calculaQtdeDiasClienteCadastrado
'   Calcula há quanto tempo (em dias) o cliente já está cadastrado.
function calculaQtdeDiasClienteCadastrado(Byval cnpj_cpf)
dim strSql, dt_cliente_desde, qtdeDiasClienteCadastrado
dim t
	calculaQtdeDiasClienteCadastrado = 0
	strSql = "SELECT" & _
				" Min(data) AS dt_cliente_desde" & _
			" FROM t_PEDIDO tP" & _
				" INNER JOIN t_CLIENTE tC ON (tP.id_cliente = tC.id)" & _
			" WHERE" & _
				" (tC.cnpj_cpf = '" & retorna_so_digitos(cnpj_cpf) & "')"
	set t = cn.Execute(strSql)
	if t.Eof then exit function
	dt_cliente_desde = t("dt_cliente_desde")
	t.Close
	set t = nothing
	qtdeDiasClienteCadastrado = DateDiff("d", dt_cliente_desde, Date)
	if qtdeDiasClienteCadastrado < 0 then qtdeDiasClienteCadastrado = 0
	calculaQtdeDiasClienteCadastrado = qtdeDiasClienteCadastrado
end function



' --------------------------------------------------------------------------------
'   obtemDataUltCompra
'   Obtém a data da última compra.
function obtemDataUltCompra(Byval cnpj_cpf, Byval pedido)
dim strSql, dt_ult_compra
dim t
	obtemDataUltCompra = Null
	strSql = "SELECT TOP 1" & _
				" tP.data" & _
			" FROM t_PEDIDO tP" & _
				" INNER JOIN t_CLIENTE tC ON (tP.id_cliente = tC.id)" & _
			" WHERE" & _
				" (tC.cnpj_cpf = '" & retorna_so_digitos(cnpj_cpf) & "')" & _
				" AND (tP.pedido <> '" & pedido & "')" & _
				" AND (tP.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
				" AND (tP.pedido = tP.pedido_base)" & _
			" ORDER BY" & _
				" data DESC"
	set t = cn.Execute(strSql)
	if t.Eof then exit function
	dt_ult_compra = t("data")
	t.Close
	set t = nothing
	obtemDataUltCompra = dt_ult_compra
end function



' --------------------------------------------------------------------------------
'   obtemQtdeTentativasCompra
'   Nº de vezes que tentou fazer o pagamento do pedido. Cartões de crédito diferentes tentados e/ou outros meios de pagamento tentados. Para o mesmo pedido.
'   Lembrando que uma transação pode estar na situação 'captura cancelada' ou 'estornada' somente se ela foi capturada anteriormente.
function obtemQtdeTentativasCompra(Byval cnpj_cpf, Byval pedido, Byval id_pedido_pagto_braspag_atual)
dim strSql, intQtdeTentativasCompra
dim t
	obtemQtdeTentativasCompra = 0
	strSql = "SELECT" & _
				" Count(*) AS qtde" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG tPPB" & _
				" INNER JOIN t_CLIENTE tC ON (tPPB.id_cliente = tC.id)" & _
			" WHERE" & _
				" (tC.cnpj_cpf = '" & retorna_so_digitos(cnpj_cpf) & "')" & _
				" AND (tPPB.pedido = '" & pedido & "')"
	
	if id_pedido_pagto_braspag_atual > 0 then
		strSql = strSql & _
				" AND (tPPB.id <> " & id_pedido_pagto_braspag_atual & ")"
		end if
	
	strSql = strSql & _
				" AND (tPPB.ult_PAG_GlobalStatus <> '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "')" & _
				" AND (tPPB.ult_PAG_GlobalStatus <> '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA & "')" & _
				" AND (tPPB.ult_PAG_GlobalStatus <> '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURA_CANCELADA & "')" & _
				" AND (tPPB.ult_PAG_GlobalStatus <> '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNADA & "')"
	set t = cn.Execute(strSql)
	if t.Eof then exit function
	intQtdeTentativasCompra = t("qtde")
	t.Close
	set t = nothing
	obtemQtdeTentativasCompra = intQtdeTentativasCompra
end function



' --------------------------------------------------------------------------------
'   BraspagDescricaoOperacaoRegistraPagto
'   Retorna a descrição para os códigos usados no processamento do registro
'   automático do pagamento no pedido.
function BraspagDescricaoOperacaoRegistraPagto(byval codigo_operacao)
dim strResp
	if codigo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CAPTURA then
		strResp = "Captura"
	elseif codigo_operacao = BRASPAG_REGISTRA_PAGTO__OP_AUTORIZACAO then
		strResp = "Autorização"
	elseif codigo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CANCELAMENTO then
		strResp = "Cancelamento"
	elseif codigo_operacao = BRASPAG_REGISTRA_PAGTO__OP_ESTORNO then
		strResp = "Estorno"
	elseif Trim(codigo_operacao) <> "" then
		strResp = "Código desconhecido (" & codigo_operacao & ")"
	else
		strResp = ""
		end if
	BraspagDescricaoOperacaoRegistraPagto = strResp
end function



' --------------------------------------------------------------------------------
'   geraBraspagPagtoSufixoPedidoNsu
'   Gera um sufixo do tipo NSU para o pedido de forma a poder identificar na
'   Braspag de maneira inequívoca uma transação enviada através do nº do pedido.
function geraBraspagPagtoSufixoPedidoNsu(Byval pedido, Byval usuario)
dim strSql, intNsu, lngRecordsAffected, s_log
dim t
	intNsu = 0
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG_NSU" & _
			" WHERE" & _
				" (pedido = '" & pedido & "')"
	set t = cn.Execute(strSql)
	if t.Eof then
		intNsu = 0
		t.Close
		set t = nothing
		strSql = "INSERT INTO t_PEDIDO_PAGTO_BRASPAG_NSU (" & _
					"pedido," & _
					"nsu," & _
					"dt_hr_atualizacao," & _
					"usuario_atualizacao" & _
				") VALUES (" & _
					"'" & pedido & "'," & _
					"0," & _
					"getdate()," & _
					"'" & usuario & "'" & _
				")"
		cn.Execute strSql, lngRecordsAffected
	else
		intNsu = t("nsu")
		t.Close
		set t = nothing
		end if
	
	intNsu = intNsu + 1
	strSql = "UPDATE t_PEDIDO_PAGTO_BRASPAG_NSU SET " & _
				"nsu = " & Cstr(intNsu) & "," & _
				"dt_hr_atualizacao = getdate()," & _
				"usuario_atualizacao = '" & usuario & "'" & _
			" WHERE" & _
				" (pedido = '" & pedido & "')"
	cn.Execute strSql, lngRecordsAffected
	
	s_log = "Gerado NSU=" & Cstr(intNsu) & " para o sufixo do pedido " & pedido & " na transação de pagamento da Braspag"
	grava_log usuario, "", pedido, "", OP_LOG_BRASPAG_PEDIDO_NSU_GERADO, s_log
	
	geraBraspagPagtoSufixoPedidoNsu = intNsu
end function



' --------------------------------------------------------------------------------
'   BraspagAfDescricaoReasonCode
'   Retorna a descrição para o ReasonCode informado pelo Antifraude.
function BraspagAfDescricaoReasonCode(byval codigoReasonCode)
dim strResp
	strResp = ""
	codigoReasonCode = Trim("" & codigoReasonCode)
	select case codigoReasonCode
		case "100"
			strResp = "Operação bem sucedida."
		case "101"
			strResp = "O pedido está faltando um ou mais campos necessários. Possível ação: Veja os campos que estão faltando na lista AntiFraudResponse.MissingFieldCollection. Reenviar o pedido com a informação completa."
		case "102"
			strResp = "Um ou mais campos do pedido contêm dados inválidos. Possível ação: Veja os campos inválidos na lista AntiFraudResponse.InvalidFieldCollection. Reenviar o pedido com as informações corretas."
		case "150"
			strResp = "Falha no sistema geral. Possível ação: Aguarde alguns minutos e tente reenviar o pedido."
		case "151"
			strResp = "O pedido foi recebido, mas ocorreu time-out no servidor. Este erro não inclui time-out entre o cliente e o servidor. Possível ação: Aguarde alguns minutos e tente reenviar o pedido."
		case "152"
			strResp = "O pedido foi recebido, mas ocorreu time-out. Possível ação: Aguarde alguns minutos e reenviar o pedido."
		case "202"
			strResp = "CyberSource recusou o pedido porque o cartão expirou. Você também pode receber este código se a data de validade não coincidir com a data em arquivo do banco emissor. Se o processador de pagamento permite a emissão de créditos para cartões expirados, a CyberSource não limita essa funcionalidade. Possível ação: Solicite um cartão ou outra forma de pagamento."
		case "231"
			strResp = "O número da conta é inválido. Possível ação: Solicite um cartão ou outra forma de pagamento."
		case "234"
			strResp = "Há um problema com a configuração do comerciante na CyberSource. Possível ação: Não envie o pedido. Entre em contato com o Suporte ao Cliente para corrigir o problema de configuração."
		case "400"
			strResp = "A pontuação de fraude ultrapassa o seu limite. Possível ação: Reveja o pedido do cliente."
		case "480"
			strResp = "O pedido foi marcado para revisão pelo Gerenciador de Decisão."
		case "481"
			strResp = "O pedido foi rejeitado pelo Gerenciador de Decisão."
		case else
			if codigoReasonCode <> "" then strResp = "Código Desconhecido (" & codigoReasonCode & ")"
		end select
	BraspagAfDescricaoReasonCode = strResp
end function



' --------------------------------------------------------------------------------
'   BraspagAfDescricaoAddressInfoCode
'   Retorna a descrição para o AddressInfoCode informado pelo Antifraude.
function BraspagAfDescricaoAddressInfoCode(byval codigo)
dim strResp
	strResp = ""
	codigo = Trim("" & codigo)
	select case codigo
		case "COR-BA"
			strResp = "Endereço de cobrança corrigido ou corrigível."
		case "COR-SA"
			strResp = "Endereço de entrega corrigido ou corrigível."
		case "INTL-BA"
			strResp = "O país de cobrança é fora dos U.S."
		case "INTL-SA"
			strResp = "O país de entrega é fora dos U.S."
		case "MIL-USA"
			strResp = "Este é um endereço militar nos U.S."
		case "MM-A"
			strResp = "Endereços diferentes de cobrança e envio."
		case "MM-BIN"
			strResp = "O BIN do cartão (os seis primeiros dígitos do número) não corresponde ao país."
		case "MM-C"
			strResp = "Os endereços de cobrança e entrega usam cidades diferentes."
		case "MM-CO"
			strResp = "Os endereços de cobrança e entrega usam países diferentes."
		case "MM-ST"
			strResp = "Os endereços de cobrança e entrega usam estados diferentes."
		case "MM-Z"
			strResp = "Os endereços de cobrança e entrega usam códidos postais diferentes."
		case "UNV-ADDR"
			strResp = "O endereço é inverificável."
		case else
			if codigo <> "" then strResp = "Código Desconhecido (" & codigo & ")"
		end select
	BraspagAfDescricaoAddressInfoCode = strResp
end function



' --------------------------------------------------------------------------------
'   BraspagAfDescricaoAfsFactorCode
'   Retorna a descrição para o AfsFactorCode informado pelo Antifraude.
function BraspagAfDescricaoAfsFactorCode(byval codigo)
dim strResp
	strResp = ""
	codigo = Trim("" & codigo)
	select case codigo
		case "A"
			strResp = "Mudança de endereço excessiva. O cliente mudou o endereço de cobrança duas ou mais vezes nos últimos seis meses."
		case "B"
			strResp = "BIN do cartão ou autorização de risco. Os fatores de risco estão relacionados com BIN de cartão de crédito e/ou verificações de autorização do cartão."
		case "C"
			strResp = "Elevado números de cartões de créditos. O cliente tem usado mais de seis números de cartões de créditos nos últimos seis meses."
		case "D"
			strResp = "Impacto do endereço de e-mail. O cliente usa um provedor de e-mail gratuito ou o endereço de email é arriscado."
		case "E"
			strResp = "Lista positiva. O cliente está na sua lista positiva."
		case "F"
			strResp = "Lista negativa. O número da conta, endereço, endereço de e-mail ou endereço IP para este fim aparece na sua lista negativa."
		case "G"
			strResp = "Inconsistências de geolocalização. O domínio do cliente de e-mail, número de telefone, endereço de cobrança, endereço de envio ou endereço IP é suspeito."
		case "H"
			strResp = "Excessivas mudanças de nome. O cliente mudou o nome duas ou mais vezes nos últimos seis meses."
		case "I"
			strResp = "Inconsistências de internet. O endereço IP e de domínio de e-mail não são consistentes com o endereço de cobrança."
		case "N"
			strResp = "Entrada sem sentido. O nome do cliente e os campos de endereço contém palavras ou linguagem sem sentido."
		case "O"
			strResp = "Obscenidades. Dados do cliente contém palavras obscenas."
		case "P"
			strResp = "Identidade morphing. Vários valores de um elemento de identidade estão ligados a um valor de um elemento de identidade diferentes. Por exemplo, vários números de telefone estão ligados a um número de conta única."
		case "Q"
			strResp = "Inconsistências do telefone. O número de telefone do cliente é suspeito."
		case "R"
			strResp = "Ordem arriscada. A transação, o cliente e o lojista mostram informações correlacionadas de alto risco."
		case "T"
			strResp = "Time hedge. O cliente está tentando uma compra fora do horário esperado."
		case "U"
			strResp = "Endereço não verificável. O endereço de cobrança ou de entrega não pode ser verificado."
		case "V"
			strResp = "Velocidade. O número da conta foi usado muitas vezes nos últimos 15 minutos."
		case "W"
			strResp = "Marcado como suspeito. O endereço de cobrança ou de entrega é semelhante a um endereço previamente marcado como suspeito."
		case "Y"
			strResp = "O endereço, cidade, estado ou país dos endereços de cobrança e entrega não se correlacionam."
		case "Z"
			strResp = "Valor inválido. Como a solicitação contém um valor inesperado, um valor padrão foi substituído. Embora a transação ainda possa ser processada, examinar o pedido com cuidado para detectar anomalias."
		case else
			if codigo <> "" then strResp = "Código Desconhecido (" & codigo & ")"
		end select
	BraspagAfDescricaoAfsFactorCode = strResp
end function



' --------------------------------------------------------------------------------
'   BraspagAfDescricaoHotlistInfoCode
'   Retorna a descrição para o HotlistInfoCode informado pelo Antifraude.
function BraspagAfDescricaoHotlistInfoCode(byval codigo)
dim strResp
	strResp = ""
	codigo = Trim("" & codigo)
	select case codigo
		case "CON-POSNEG"
			strResp = "A ordem disparada bate tanto com a lista positiva e negativa. O resultado da lista positiva sobrescreve a lista negativa."
		case "NEG-BA"
			strResp = "O endereço de cobrança está na lista negativa."
		case "NEG-BCO"
			strResp = "O país de cobrança está na lista negativa."
		case "NEG-BIN"
			strResp = "O BIN do cartão de crédito (os seis primeiros dígitos do número do cartão) está na lista negativa."
		case "NEG-BINCO"
			strResp = "O país em que o cartão de crédito foi emitido está na lista negativa."
		case "NEG-BZC"
			strResp = "O código postal de cobrança está na lista negativa."
		case "NEG-CC"
			strResp = "O número de cartão de crédito está na lista negativa."
		case "NEG-EM"
			strResp = "O endereço de e-mail está na lista negativa."
		case "NEG-EMCO"
			strResp = "O país em que o endereço de e-mail está localizado consta na lista negativa."
		case "NEG-EMDOM"
			strResp = "O domínio de e-mail (por exemplo, mail.example.com) está na lista negativa."
		case "NEG-FP"
			strResp = "O device fingerprint está na lista negativa"
		case "NEG-HIST"
			strResp = "A transação foi encontrada na lista negativa."
		case "NEG-ID"
			strResp = "ID da conta do cliente está na lista negativa."
		case "NEG-IP"
			strResp = "O endereço IP (por exemplo, 10.1.27.63) está na lista negativa."
		case "NEG-IP3"
			strResp = "O endereço IP de rede (por exemplo, 10.1.27) está na lista negativa. Um endereço de IP da rede inclui até 256 endereços IP."
		case "NEG-IPCO"
			strResp = "O país em que o endereço IP está localizado está na lista negativa."
		case "NEG-PEM"
			strResp = "Um endereço de e-mail do passageiro está na lista negativa."
		case "NEG-PH"
			strResp = "O número do telefone está na lista negativa."
		case "NEG-PID"
			strResp = "ID da conta do passageiro está na lista negativa."
		case "NEG-PPH"
			strResp = "O número do telefone do passageiro está na lista negativa."
		case "NEG-SA"
			strResp = "O endereço de entrega está na lista negativa."
		case "NEG-SCO"
			strResp = "O país de entrega está na lista negativa."
		case "NEG-SZC"
			strResp = "O código postal de entrega está na lista negativa."
		case "POS-TEMP"
			strResp = "O cliente está na lista positiva temporária."
		case "POS-PERM"
			strResp = "O cliente está na lista positiva permanente."
		case "REV-BA"
			strResp = "O endereço de cobrança esta na lista de revisão."
		case "REV-BCO"
			strResp = "O país de cobrança está na lista de revisão."
		case "REV-BIN"
			strResp = "O BIN do cartão de crédito (os seis primeiros dígitos do número do cartão) está na lista de revisão."
		case "REV-BINCO"
			strResp = "O país em que o cartão de crédito foi emitido está na lista de revisão."
		case "REV-BZC"
			strResp = "O código postal de cobrança está na lista de revisão."
		case "REV-CC"
			strResp = "O número do cartão de crédito está na lista de revisão."
		case "REV-EM"
			strResp = "O endereço de e-mail está na lista de revisão."
		case "REV-EMCO"
			strResp = "O país em que o endereço de e-mail está localizado está na lista de revisão."
		case "REV-EMDOM"
			strResp = "O domínio de e-mail (por exemplo, mail.example.com) está na lista de revisão."
		case "REV-FP"
			strResp = "O device fingerprint está na lista de revisão"
		case "REV-ID"
			strResp = "ID da conta do cliente está na lista de revisão."
		case "REV-IP"
			strResp = "O endereço IP (por exemplo, 10.1.27.63) está na lista de revisão."
		case "REV-IP3"
			strResp = "O endereço IP de rede (por exemplo, 10.1.27) está na lista de revisão. Um endereço de IP da rede inclui até 256 endereços IP."
		case "REV-IPCO"
			strResp = "O país em que o endereço IP está localizado está na lista de revisão."
		case "REV-PEM"
			strResp = "Um endereço de e-mail do passageiro está na lista de revisão."
		case "REV-PH"
			strResp = "O número do telefone está na lista de revisão."
		case "REV-PID"
			strResp = "ID da conta do passageiro está na lista de revisão."
		case "REV-PPH"
			strResp = "O número do telefone do passageiro está na lista de revisão."
		case "REV-SA"
			strResp = "O endereço de entrega está na lista de revisão."
		case "REV-SCO"
			strResp = "O país de entrega está na lista de revisão."
		case "REV-SZC"
			strResp = "O código postal de entrega está na lista de revisão."
		case else
			if codigo <> "" then strResp = "Código Desconhecido (" & codigo & ")"
		end select
	BraspagAfDescricaoHotlistInfoCode = strResp
end function



' --------------------------------------------------------------------------------
'   BraspagAfDescricaoIdentityInfoCode
'   Retorna a descrição para o IdentityInfoCode informado pelo Antifraude.
function BraspagAfDescricaoIdentityInfoCode(byval codigo)
dim strResp
	strResp = ""
	codigo = Trim("" & codigo)
	select case codigo
		case "MORPH-B"
			strResp = "O mesmo endereço de cobrança tem sido utilizado várias vezes com identidades de clientes múltiplos."
		case "MORPH-C"
			strResp = "O mesmo número de conta tem sido utilizado várias vezes com identidades de clientes múltiplos."
		case "MORPH-E"
			strResp = "O mesmo endereço de e-mail tem sido utilizado várias vezes com identidades de clientes múltiplos."
		case "MORPH-I"
			strResp = "O mesmo endereço IP tem sido utilizado várias vezes com identidades de clientes múltiplos."
		case "MORPH-P"
			strResp = "O mesmo número de telefone tem sido usado várias vezes com identidades de clientes múltiplos."
		case "MORPH-S"
			strResp = "O mesmo endereço de entrega tem sido utilizado várias vezes com identidades de clientes múltiplos."
		case else
			if codigo <> "" then strResp = "Código Desconhecido (" & codigo & ")"
		end select
	BraspagAfDescricaoIdentityInfoCode = strResp
end function



' --------------------------------------------------------------------------------
'   BraspagAfDescricaoInternetInfoCode
'   Retorna a descrição para o InternetInfoCode informado pelo Antifraude.
function BraspagAfDescricaoInternetInfoCode(byval codigo)
dim strResp
	strResp = ""
	codigo = Trim("" & codigo)
	select case codigo
		case "FREE-EM"
			strResp = "O endereço de e-mail do cliente é de um provedor de e-mail gratuito."
		case "INTL-IPCO"
			strResp = "O país do endereço de e-mail do cliente é fora do U.S."
		case "INV-EM"
			strResp = "O endereço de e-mail do cliente é inválido."
		case "MM-EMBCO"
			strResp = "O domínio do endereço de e-mail do cliente não é consistente com o país do endereço de cobrança."
		case "MM-IPBC"
			strResp = "O endereço IP do cliente não é consistente com a cidade do endereço de cobrança."
		case "MM-IPBCO"
			strResp = "O endereço IP do cliente não é consistente com a país do endereço de cobrança."
		case "MM-IPBST"
			strResp = "O endereço IP do cliente não é consistente com o estado no endereço de cobrança. No entanto, este código de informação não pode ser devolvido quando a inconsistência é entre estados imediatamente adjacentes."
		case "MM-IPEM"
			strResp = "O endereço de e-mail do cliente não é consistente com o endereço IP."
		case "RISK-EM"
			strResp = "O domínio do e-mail do cliente (por exemplo, mail.example.com) está associada com alto risco."
		case "UNV-NID"
			strResp = "O endereço IP do cliente é de um proxy anônimo. Estas entidades escondem completamente informações sobre o endereço de IP."
		case "UNV-RI400SK"
			strResp = "O endereço IP é de origem de risco."
		case "UNV-EMBCO"
			strResp = "O país do endereço do cliente de e-mail não corresponde ao país do endereço de cobrança."
		case else
			if codigo <> "" then strResp = "Código Desconhecido (" & codigo & ")"
		end select
	BraspagAfDescricaoInternetInfoCode = strResp
end function



' --------------------------------------------------------------------------------
'   BraspagAfDescricaoPhoneInfoCode
'   Retorna a descrição para o PhoneInfoCode informado pelo Antifraude.
function BraspagAfDescricaoPhoneInfoCode(byval codigo)
dim strResp
	strResp = ""
	codigo = Trim("" & codigo)
	select case codigo
		case "MM-ACBST"
			strResp = "O número de telefone do cliente não é consistente com o estado no endereço de cobrança."
		case "RISK-AC"
			strResp = "O código de área do cliente está associado com risco alto."
		case "RISK-PH"
			strResp = "O número de telefone dos U.S. ou do Canadá é incompleta, ou uma ou mais partes do número são arriscadas."
		case "TF-AC"
			strResp = "O número do telefone utiliza um código de área toll-free."
		case "UNV-AC"
			strResp = "O código de área é inválido."
		case "UNV-OC"
			strResp = "O código de área e/ou o prefixo de telefone são/é inválido."
		case "UNV-PH"
			strResp = "O número do telefone é inválido."
		case else
			if codigo <> "" then strResp = "Código Desconhecido (" & codigo & ")"
		end select
	BraspagAfDescricaoPhoneInfoCode = strResp
end function



' --------------------------------------------------------------------------------
'   BraspagAfDescricaoSuspiciousInfoCode
'   Retorna a descrição para o SuspiciousInfoCode informado pelo Antifraude.
function BraspagAfDescricaoSuspiciousInfoCode(byval codigo)
dim strResp
	strResp = ""
	codigo = Trim("" & codigo)
	select case codigo
		case "BAD-FP"
			strResp = "O dispositivo é arriscado."
		case "INTL-BIN"
			strResp = "O cartão de crédito foi emitido fora dos U.S."
		case "MM-TZTLO"
			strResp = "Fuso horário do dispositivo é incompatível com os fusos horários do país."
		case "MUL-EM"
			strResp = "O cliente tem usado mais de quatro endereços de email diferentes."
		case "NON-BC"
			strResp = "A cidade de cobrança é uma desconhecida."
		case "NON-FN"
			strResp = "O primeiro nome do cliente é desconhecido."
		case "NON-LN"
			strResp = "O último nome do cliente é desconhecido."
		case "OBS-BC"
			strResp = "A cidade de cobrança contem obscenidades."
		case "OBS-EM"
			strResp = "O endereço de e-mail contem obscenidades."
		case "RISK-AVS"
			strResp = "O resultado do teste combinado AVS e endereço de cobrança normalizado são arriscados, o resultado AVS indica uma correspondência exata, mas o endereço de cobrança não é entrega normalizada."
		case "RISK-BC"
			strResp = "A cidade de cobrança possui caracteres repetidos."
		case "RISK-BIN"
			strResp = "No passado, este BIN do cartão de crédito (os seis primeiros dígitos do número do cartão) mostrou uma elevada incidência de fraude."
		case "RISK-DEV"
			strResp = "Algumas das características do dispositivo são arriscadas."
		case "RISK-FN"
			strResp = "Nome e sobrenome do cliente contêm combinações de letras improváveis."
		case "RISK-LN"
			strResp = "Nome do meio ou o sobrenome do cliente contém combinações de letras improváveis."
		case "RISK-PIP"
			strResp = "O endereço IP do proxy é arriscado."
		case "RISK-SD"
			strResp = "A inconsistência nos países de cobrança e entrega é arriscado."
		case "RISK-TB"
			strResp = "O dia e a hora da ordem associada ao endereço de cobrança é arriscado."
		case "RISK-TIP"
			strResp = "O verdadeiro endereço IP é arriscado."
		case "RISK-TS"
			strResp = "O dia e a hora da ordem associada ao endereço de entrega é arriscado."
		case else
			if codigo <> "" then strResp = "Código Desconhecido (" & codigo & ")"
		end select
	BraspagAfDescricaoSuspiciousInfoCode = strResp
end function



' --------------------------------------------------------------------------------
'   BraspagAfDescricaoVelocityInfoCode
'   Retorna a descrição para o VelocityInfoCode informado pelo Antifraude.
function BraspagAfDescricaoVelocityInfoCode(byval codigo)
dim strResp
	strResp = ""
	codigo = Trim("" & codigo)
	select case codigo
		case "VEL-ADDR"
			strResp = "Diferentes estados de faturamento e/ou o envio (EUA e Canadá apenas) têm sido usadas várias vezes com o número do cartão de crédito e/ou endereço de email."
		case "VEL-CC"
			strResp = "Diferentes números de contas foram usados várias vezes com o mesmo nome ou endereço de email."
		case "VEL-NAME"
			strResp = "Diferentes nomes foram usados várias vezes com o número do cartão de crédito e/ou endereço de email."
		case "VELS-CC"
			strResp = "O número de conta tem sido utilizado várias vezes durante o intervalo de controle curto."
		case "VELI-CC"
			strResp = "O número de conta tem sido utilizado várias vezes durante o intervalo de controle médio."
		case "VELL-CC"
			strResp = "O número de conta tem sido utilizado várias vezes durante o intervalo de controle longo."
		case "VELV-CC"
			strResp = "O número de conta tem sido utilizado várias vezes durante o intervalo de controle muito longo."
		case "VELS-EM"
			strResp = "O endereço de e-mail tem sido utilizado várias vezes durante o intervalo de controle curto."
		case "VELI-EM"
			strResp = "O endereço de e-mail tem sido utilizado várias vezes durante o intervalo de controle médio."
		case "VELL-EM"
			strResp = "O endereço de e-mail tem sido utilizado várias vezes durante o intervalo de controle longo."
		case "VELV-EM"
			strResp = "O endereço de e-mail tem sido utilizado várias vezes durante o intervalo de controle muito longo."
		case "VELS-FP"
			strResp = "O device fingerprint tem sido utilizado várias vezes durante um intervalo curto"
		case "VELI-FP"
			strResp = "O device fingerprint tem sido utilizado várias vezes durante um intervalo médio"
		case "VELL-FP"
			strResp = "O device fingerprint tem sido utilizado várias vezes durante um intervalo longo"
		case "VELV-FP"
			strResp = "O device fingerprint tem sido utilizado várias vezes durante um intervalo muito longo"
		case "VELS-IP"
			strResp = "O endereço IP tem sido utilizado várias vezes durante o intervalo de controle curto."
		case "VELI-IP"
			strResp = "O endereço IP tem sido utilizado várias vezes durante o intervalo de controle médio."
		case "VELL-IP"
			strResp = "O endereço IP tem sido utilizado várias vezes durante o intervalo de controle longo."
		case "VELV-IP"
			strResp = "O endereço IP tem sido utilizado várias vezes durante o intervalo de controle muito longo."
		case "VELS-SA"
			strResp = "O endereço de entrega tem sido utilizado várias vezes durante o intervalo de controle curto."
		case "VELI-SA"
			strResp = "O endereço de entrega tem sido utilizado várias vezes durante o intervalo de controle médio."
		case "VELL-SA"
			strResp = "O endereço de entrega tem sido utilizado várias vezes durante o intervalo de controle longo."
		case "VELV-SA"
			strResp = "O endereço de entrega tem sido utilizado várias vezes durante o intervalo de controle muito longo."
		case "VELS-TIP"
			strResp = "O endereço IP verdadeiro tem sido utilizado várias vezes durante o intervalo de controle curto."
		case "VELI-TIP"
			strResp = "O endereço IP verdadeiro tem sido utilizado várias vezes durante o intervalo de controle médio."
		case "VELL-TIP"
			strResp = "O endereço IP verdadeiro tem sido utilizado várias vezes durante o intervalo de controle longo."
		case else
			if codigo <> "" then strResp = "Código Desconhecido (" & codigo & ")"
		end select
	BraspagAfDescricaoVelocityInfoCode = strResp
end function



' --------------------------------------------------------------------------------
'   BraspagObtemOwnerPeloPedido
'   A partir do nº do pedido, identifica e retorna a empresa usada na transação
'   com a Braspag (OLD01, OLD02, etc)
function BraspagObtemOwnerPeloPedido(ByVal pedido)
dim r, s
	pedido = Trim("" & pedido)
'	VERIFICA SE O PRÓPRIO PEDIDO JÁ POSSUI A INFORMAÇÃO DE QUAL É A EMPRESA RESPONSÁVEL PELA EMISSÃO DA NFe E, ATRAVÉS DELA, QUAL É A EMPRESA DEFINIDA COMO RESPONSÁVEL PELAS TRANSAÇÕES BRASPAG
	s = "SELECT " & _
			"braspag_id_boleto_cedente" & _
		" FROM t_PEDIDO" & _
			" INNER JOIN t_NFe_EMITENTE ON (t_PEDIDO.id_nfe_emitente = t_NFe_EMITENTE.id)" & _
		" WHERE" & _
			" (pedido = '" & pedido & "')"
	set r = cn.Execute(s)
	if Not r.Eof then BraspagObtemOwnerPeloPedido = CLng(r("braspag_id_boleto_cedente"))
	if r.State <> 0 then r.Close
	set r = Nothing
end function



' --------------------------------------------------------------------------------
'   BraspagIsBandeiraHabilitada
'   Para a empresa indicada, informa se a bandeira está habilitada p/ transacionar.
function BraspagIsBandeiraHabilitada(ByVal owner, ByVal bandeira)
dim blnResp
	blnResp = False
	if Cstr(owner) = Cstr(BRASPAG_OWNER_OLD01) then
		if Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__VISA)) then
			blnResp = BRASPAG_OLD01_BANDEIRA_HABILITADA__VISA
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__MASTERCARD)) then
			blnResp = BRASPAG_OLD01_BANDEIRA_HABILITADA__MASTERCARD
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__AMEX)) then
			blnResp = BRASPAG_OLD01_BANDEIRA_HABILITADA__AMEX
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__ELO)) then
			blnResp = BRASPAG_OLD01_BANDEIRA_HABILITADA__ELO
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__HIPERCARD)) then
			blnResp = BRASPAG_OLD01_BANDEIRA_HABILITADA__HIPERCARD
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__DINERS)) then
			blnResp = BRASPAG_OLD01_BANDEIRA_HABILITADA__DINERS
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__DISCOVER)) then
			blnResp = BRASPAG_OLD01_BANDEIRA_HABILITADA__DISCOVER
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__AURA)) then
			blnResp = BRASPAG_OLD01_BANDEIRA_HABILITADA__AURA
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__JCB)) then
			blnResp = BRASPAG_OLD01_BANDEIRA_HABILITADA__JCB
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__CELULAR)) then
			blnResp = BRASPAG_OLD01_BANDEIRA_HABILITADA__CELULAR
			end if
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_OLD02) then
		if Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__VISA)) then
			blnResp = BRASPAG_OLD02_BANDEIRA_HABILITADA__VISA
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__MASTERCARD)) then
			blnResp = BRASPAG_OLD02_BANDEIRA_HABILITADA__MASTERCARD
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__AMEX)) then
			blnResp = BRASPAG_OLD02_BANDEIRA_HABILITADA__AMEX
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__ELO)) then
			blnResp = BRASPAG_OLD02_BANDEIRA_HABILITADA__ELO
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__HIPERCARD)) then
			blnResp = BRASPAG_OLD02_BANDEIRA_HABILITADA__HIPERCARD
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__DINERS)) then
			blnResp = BRASPAG_OLD02_BANDEIRA_HABILITADA__DINERS
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__DISCOVER)) then
			blnResp = BRASPAG_OLD02_BANDEIRA_HABILITADA__DISCOVER
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__AURA)) then
			blnResp = BRASPAG_OLD02_BANDEIRA_HABILITADA__AURA
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__JCB)) then
			blnResp = BRASPAG_OLD02_BANDEIRA_HABILITADA__JCB
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__CELULAR)) then
			blnResp = BRASPAG_OLD02_BANDEIRA_HABILITADA__CELULAR
			end if
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_DIS) then
		if Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__VISA)) then
			blnResp = BRASPAG_DIS_BANDEIRA_HABILITADA__VISA
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__MASTERCARD)) then
			blnResp = BRASPAG_DIS_BANDEIRA_HABILITADA__MASTERCARD
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__AMEX)) then
			blnResp = BRASPAG_DIS_BANDEIRA_HABILITADA__AMEX
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__ELO)) then
			blnResp = BRASPAG_DIS_BANDEIRA_HABILITADA__ELO
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__HIPERCARD)) then
			blnResp = BRASPAG_DIS_BANDEIRA_HABILITADA__HIPERCARD
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__DINERS)) then
			blnResp = BRASPAG_DIS_BANDEIRA_HABILITADA__DINERS
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__DISCOVER)) then
			blnResp = BRASPAG_DIS_BANDEIRA_HABILITADA__DISCOVER
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__AURA)) then
			blnResp = BRASPAG_DIS_BANDEIRA_HABILITADA__AURA
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__JCB)) then
			blnResp = BRASPAG_DIS_BANDEIRA_HABILITADA__JCB
		elseif Ucase(Cstr(bandeira)) = Ucase(Cstr(BRASPAG_BANDEIRA__CELULAR)) then
			blnResp = BRASPAG_DIS_BANDEIRA_HABILITADA__CELULAR
			end if
		end if

	BraspagIsBandeiraHabilitada = blnResp
end function



' --------------------------------------------------------------------------------
'   BraspagObtem_AF_MERCHANT_ID
function BraspagObtem_AF_MERCHANT_ID(Byval owner)
dim resp
	if Cstr(owner) = Cstr(BRASPAG_OWNER_OLD01) then
		resp = BRASPAG_OLD01_AF_MERCHANT_ID
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_OLD02) then
		resp = BRASPAG_OLD02_AF_MERCHANT_ID
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_DIS) then
		resp = BRASPAG_DIS_AF_MERCHANT_ID
		end if
	BraspagObtem_AF_MERCHANT_ID = resp
end function



' --------------------------------------------------------------------------------
'   BraspagObtem_PAG_MERCHANT_ID
function BraspagObtem_PAG_MERCHANT_ID(Byval owner)
dim resp
	if Cstr(owner) = Cstr(BRASPAG_OWNER_OLD01) then
		resp = BRASPAG_OLD01_PAG_MERCHANT_ID
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_OLD02) then
		resp = BRASPAG_OLD02_PAG_MERCHANT_ID
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_DIS) then
		resp = BRASPAG_DIS_PAG_MERCHANT_ID
		end if
	BraspagObtem_PAG_MERCHANT_ID = resp
end function



' --------------------------------------------------------------------------------
'   BraspagObtem_DF_MERCHANT_ID
function BraspagObtem_DF_MERCHANT_ID(Byval owner)
dim resp
	if Cstr(owner) = Cstr(BRASPAG_OWNER_OLD01) then
		resp = BRASPAG_OLD01_DF_MERCHANT_ID
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_OLD02) then
		resp = BRASPAG_OLD02_DF_MERCHANT_ID
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_DIS) then
		resp = BRASPAG_DIS_DF_MERCHANT_ID
		end if
	BraspagObtem_DF_MERCHANT_ID = resp
end function



' --------------------------------------------------------------------------------
'   BraspagObtem_ORG_ID
function BraspagObtem_ORG_ID(Byval owner)
dim resp
	if Cstr(owner) = Cstr(BRASPAG_OWNER_OLD01) then
		resp = BRASPAG_OLD01_ORG_ID
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_OLD02) then
		resp = BRASPAG_OLD02_ORG_ID
	elseif Cstr(owner) = Cstr(BRASPAG_OWNER_DIS) then
		resp = BRASPAG_DIS_ORG_ID
		end if
	BraspagObtem_ORG_ID = resp
end function
%>