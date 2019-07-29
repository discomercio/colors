<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Braspag.asp"    -->
<!-- #include file = "../global/BraspagCS.asp"    -->

<%
'     ===========================================
'	  P080PagtoExec.asp
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


	On Error GoTo 0
	Err.Clear

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_BRASPAG_EM_SEG
	
	Const XML_PATH_AuthorizeTransactionResult = "//AuthorizeTransactionResponse/AuthorizeTransactionResult"

	dim alerta
	alerta = ""

	dim err_number, err_description
	dim s_log, s_log_aux, s_log_dados_cartao
	dim s, usuario, loja, pedido_selecionado, id_pedido_base

	usuario = BRASPAG_USUARIO_CLIENTE

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	id_pedido_base = retorna_num_pedido_base(pedido_selecionado)

	dim pedido_com_sufixo_nsu
	pedido_com_sufixo_nsu = Trim(Request("pedido_com_sufixo_nsu"))
	if pedido_com_sufixo_nsu = "" then pedido_com_sufixo_nsu = pedido_selecionado

	dim cnpj_cpf_selecionado
	cnpj_cpf_selecionado = retorna_so_digitos(Request("cnpj_cpf_selecionado"))

	dim FingerPrint_SessionID
	FingerPrint_SessionID = Trim(Request("FingerPrint_SessionID"))

	dim blnStAnaliseCreditoAlterado
	blnStAnaliseCreditoAlterado = False

	dim c_qtde_cartoes, qtde_cartoes, blnAmbiguidade
	c_qtde_cartoes = Trim(Request("c_qtde_cartoes"))
	qtde_cartoes = converte_numero(c_qtde_cartoes)
	if qtde_cartoes = 0 then Response.Redirect("aviso.asp?id=" & ERR_QTDE_CARTOES_INVALIDA)

	dim c_fatura_telefone_pais
	c_fatura_telefone_pais = Trim(Request("c_fatura_telefone_pais"))

	dim i, vDadosCartao
	redim vDadosCartao(qtde_cartoes)
	for i = 1 to qtde_cartoes
		set vDadosCartao(i) = new cl_BraspagCS_DadosCartao_Checkout
		next

'	RECUPERA DADOS DO FORMULÁRIO
	dim s_name
	for i = 1 to qtde_cartoes
		s_name = "c_cartao_bandeira_" & i
		vDadosCartao(i).bandeira = Trim(Request(s_name))
		s_name = "c_cartao_valor_" & i
		vDadosCartao(i).valor_pagamento = Trim(Request(s_name))
		s_name = "c_opcao_parcelamento_" & i
		vDadosCartao(i).opcao_parcelamento = Trim(Request(s_name))
		s_name = "c_cartao_nome_" & i
		vDadosCartao(i).titular_nome = Trim(Request(s_name))
		s_name = "c_cartao_cpf_cnpj_" & i
		vDadosCartao(i).titular_cpf_cnpj = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_cartao_numero_" & i
		vDadosCartao(i).cartao_numero = decriptografa(Trim(Request(s_name)))
		s_name = "c_cartao_validade_mes_" & i
		vDadosCartao(i).cartao_validade_mes = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_cartao_validade_ano_" & i
		vDadosCartao(i).cartao_validade_ano = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_cartao_codigo_seguranca_" & i
		vDadosCartao(i).cartao_codigo_seguranca = decriptografa(Trim(Request(s_name)))
		s_name = "c_cartao_proprio_" & i
		vDadosCartao(i).cartao_proprio = Trim(Request(s_name))
		s_name = "c_fatura_end_logradouro_" & i
		vDadosCartao(i).fatura_end_logradouro = Trim(Request(s_name))
		s_name = "c_fatura_end_numero_" & i
		vDadosCartao(i).fatura_end_numero = Trim(Request(s_name))
		s_name = "c_fatura_end_complemento_" & i
		vDadosCartao(i).fatura_end_complemento = Trim(Request(s_name))
		s_name = "c_fatura_end_bairro_" & i
		vDadosCartao(i).fatura_end_bairro = Trim(Request(s_name))
		s_name = "c_fatura_end_cidade_" & i
		vDadosCartao(i).fatura_end_cidade = Trim(Request(s_name))
		s_name = "c_fatura_end_uf_" & i
		vDadosCartao(i).fatura_end_uf = Trim(Request(s_name))
		s_name = "c_fatura_end_cep_" & i
		vDadosCartao(i).fatura_end_cep = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_fatura_telefone_ddd_" & i
		vDadosCartao(i).fatura_tel_ddd = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_fatura_telefone_numero_" & i
		vDadosCartao(i).fatura_tel_numero = retorna_so_digitos(Trim(Request(s_name)))
		next

'	PROCESSAMENTO PARA OBTER CAMPOS AUXILIARES
	dim v
	for i = 1 to qtde_cartoes
	'	VALOR DO PAGAMENTO NESTE CARTÃO
		vDadosCartao(i).vl_pagamento = converte_numero(vDadosCartao(i).valor_pagamento)
	'	OPÇÃO DE PARCELAMENTO
		if vDadosCartao(i).opcao_parcelamento = "0" then
		'	À VISTA
			vDadosCartao(i).codigo_produto = "0"
			vDadosCartao(i).qtde_parcelas = 1
		elseif InStr(vDadosCartao(i).opcao_parcelamento, "PL|") <> 0 then
		'	PARCELADO ESTABELECIMENTO (LOJA)
			vDadosCartao(i).codigo_produto = "1"
			v = Split(vDadosCartao(i).opcao_parcelamento, "|")
			vDadosCartao(i).qtde_parcelas = converte_numero(v(Ubound(v)))
			if vDadosCartao(i).qtde_parcelas <= 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Quantidade de parcelas inválida no cartão de bandeira " & BraspagDescricaoBandeira(vDadosCartao(i).bandeira)
				end if
		elseif InStr(vDadosCartao(i).opcao_parcelamento, "PC|") <> 0 then
		'	PARCELADO CARTÃO (PELO EMISSOR DO CARTÃO)
			vDadosCartao(i).codigo_produto = "2"
			v = Split(vDadosCartao(i).opcao_parcelamento, "|")
			vDadosCartao(i).qtde_parcelas = converte_numero(v(Ubound(v)))
			if vDadosCartao(i).qtde_parcelas <= 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Quantidade de parcelas inválida no cartão de bandeira " & BraspagDescricaoBandeira(vDadosCartao(i).bandeira)
				end if
		else
			alerta=texto_add_br(alerta)
			alerta=alerta & "Opção de parcelamento inválida no cartão de bandeira " & BraspagDescricaoBandeira(vDadosCartao(i).bandeira)
			end if
	'	DESCRIÇÃO DO PARCELAMENTO
		vDadosCartao(i).descricao_parcelamento = BraspagCSDescricaoParcelamento(vDadosCartao(i).codigo_produto, vDadosCartao(i).qtde_parcelas, vDadosCartao(i).vl_pagamento)
		next

	if alerta <> "" then
		exibe_erro_e_encerra(alerta)
		end if

	dim strScriptWindowName
	strScriptWindowName = _
				"<script language='JavaScript'>" & chr(13) & _
				"	window.name = '" & SITE_CLIENTE_TITULO_JANELA & "';" & chr(13) & _
				"</script>" & chr(13)


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim s_sql
	dim cn, msg_erro
	dim t_PAG, t_PAG_PAYMENT, t_PAG_XML, t_PAG_ERROR, t_PEDIDO
	dim lngRecordsAffected
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(t_PAG, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(t_PAG_PAYMENT, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(t_PAG_XML, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(t_PAG_ERROR, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(t_PEDIDO, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim r_pedido, v_item
	if Not le_pedido(id_pedido_base, r_pedido, msg_erro) then 
		alerta=msg_erro
	else
		loja = Trim(r_pedido.loja)
		if Not le_pedido_item_consolidado_familia(id_pedido_base, v_item, msg_erro) then alerta=msg_erro
		end if

	if alerta <> "" then
		exibe_erro_e_encerra(alerta)
		end if

	dim r_cliente
	set r_cliente = New cl_CLIENTE
	if Not x_cliente_bd(r_pedido.id_cliente, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_CADASTRADO)


'	VERIFICA SE ESTA REQUISIÇÃO JÁ FOI PROCESSADA
'	SUBSÍDIOS: ISSO É UMA PROTEÇÃO P/ QUANDO O CLIENTE CLICAR NO BOTÃO VOLTAR, PRINCIPALMENTE QUANDO SE RETORNA 'N' PÁGINAS ATRÁS,
'	========== JÁ QUE FOI IMPLEMENTADO UM MECANISMO DE REDIRECT P/ PREVENIR O VOLTAR P/ A PÁGINA ANTERIOR.
'	REDIRECIONA P/ A PÁGINA P030PagtoVerificaStatus.asp PARA QUE O CLIENTE VEJA AS TRANSAÇÕES BEM SUCEDIDAS
	s = "SELECT " & _
			"*" & _
		" FROM t_PAGTO_GW_PAG" & _
		" WHERE" & _
			" (pedido_com_sufixo_nsu = '" & pedido_com_sufixo_nsu & "')"
	if t_PAG.State <> 0 then t_PAG.Close
	t_PAG.Open s, cn
	if Not t_PAG.Eof then
		On Error Resume Next

	'	FECHA CONEXAO COM O BANCO DE DADOS
		if t_PAG.State <> 0 then t_PAG.Close
		set t_PAG=nothing

		if t_PAG_PAYMENT.State <> 0 then t_PAG_PAYMENT.Close
		set t_PAG_PAYMENT=nothing

		if t_PAG_XML.State <> 0 then t_PAG_XML.Close
		set t_PAG_XML=nothing

		if t_PAG_ERROR.State <> 0 then t_PAG_ERROR.Close
		set t_PAG_ERROR=nothing

		if t_PEDIDO.State <> 0 then t_PEDIDO.Close
		set t_PEDIDO=nothing

	'	FECHA CONEXÃO
		cn.Close
		set cn = nothing
		
		Response.Redirect("P030PagtoVerificaStatus.asp?pedido_selecionado=" & pedido_selecionado & "&cnpj_cpf_selecionado=" & cnpj_cpf_selecionado)
		end if


'	CRIA E MONTA A TRANSAÇÃO
'	========================
	dim trx, v_trx_payment, v_trx_payment_RX
	dim owner
	owner = BraspagObtemOwnerPeloPedido(id_pedido_base)
	set trx = cria_instancia_cl_BRASPAG_Authorize_TX(owner)
	call cria_instancia_cl_BRASPAG_Authorize_PaymentDataRequest_TX(owner, vDadosCartao, v_trx_payment)
	redim v_trx_payment_RX(0)
	set v_trx_payment_RX(0) = new cl_BRASPAG_Authorize_PaymentDataRequest_RX

	trx.CustomerData_CustomerIdentity = retorna_so_digitos(r_cliente.cnpj_cpf)
	trx.CustomerData_CustomerName = substitui_caracteres(Ucase(r_cliente.nome), "&", " E ")
	trx.OrderData_OrderId = pedido_com_sufixo_nsu

	dim idPagtoGwPag, idPagtoGwPagPayment, idPagtoGwPagXmlTx, idPagtoGwPagXmlRx, idPagtoGwPagError
	idPagtoGwPag = 0

	dim intPaymentOrdem
	dim intSequenciaErrorList, intSequenciaPaymentList

'	MONTA O CAMPO C/ OS DETALHES DO PEDIDO
	dim m_vl_total, m_total_geral
	m_total_geral = 0
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			if Trim(v_item(i).produto) <> "" then
				with v_item(i)
					m_vl_total=.qtde * .preco_NF
					m_total_geral=m_total_geral + m_vl_total
					end with
				end if
			next
		m_total_geral = m_total_geral + r_pedido.vl_frete
		end if
	
	
	if alerta = "" then
	'	GERA O NSU P/ GRAVAR NO BD O REGISTRO C/ AS INFORMAÇÕES DA TRANSAÇÃO
		if Not fin_gera_nsu(T_PAGTO_GW_PAG, idPagtoGwPag, msg_erro) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "FALHA AO GERAR NSU PARA O NOVO REGISTRO DE TRANSAÇÃO (" & msg_erro & ")"
		elseif idPagtoGwPag <= 0 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "NSU GERADO É INVÁLIDO (" & idPagtoGwPag & ")"
			end if
		
		if Not fin_gera_nsu(T_PAGTO_GW_PAG_XML, idPagtoGwPagXmlTx, msg_erro) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		elseif idPagtoGwPagXmlTx <= 0 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "NSU GERADO É INVÁLIDO (" & idPagtoGwPagXmlTx & ")"
			end if
		end if
	
	dim vl_transacao
	dim strBraspagTransactionId, strPaymentMethod, strAmount, strStatus
	dim oNodeErrorList, oNodeSet, oNodePaymentList
	dim txXml, txXmlMasked
	if alerta = "" then
		txXml = BraspagXmlMontaRequisicaoAuthorize(trx, v_trx_payment, txXmlMasked)
		txXml = retira_acentuacao(txXml)
		txXmlMasked = retira_acentuacao(txXmlMasked)
		end if
	
'	GRAVA AS INFORMAÇÕES NO BANCO DE DADOS
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		s = "SELECT * FROM t_PAGTO_GW_PAG WHERE (id = -1)"
		if t_PAG.State <> 0 then t_PAG.Close
		t_PAG.Open s, cn
		t_PAG.AddNew
		t_PAG("id") = idPagtoGwPag
		t_PAG("usuario") = usuario
		t_PAG("owner") = owner
		t_PAG("loja") = loja
		t_PAG("id_cliente") = r_pedido.id_cliente
		t_PAG("pedido") = id_pedido_base
		t_PAG("pedido_com_sufixo_nsu") = pedido_com_sufixo_nsu
		t_PAG("valor_pedido") = m_total_geral
		t_PAG("operacao") = OP_BRASPAG_OPERACAO__AUTHORIZE
		t_PAG("origem_endereco_IP") = Trim(Request.ServerVariables("REMOTE_ADDR"))
		t_PAG("FingerPrint_SessionID") = FingerPrint_SessionID
		t_PAG("trx_TX_data") = Date
		t_PAG("trx_TX_data_hora") = Now
		t_PAG("trx_TX_id_pagto_gw_pag_xml") = idPagtoGwPagXmlTx
	'	INDICA QUE A TRANSAÇÃO FOI REALIZADA DIRETAMENTE PELO CLIENTE E NÃO PELO VENDEDOR NO TELEVENDAS
		t_PAG("executado_pelo_cliente_status") = 1
		t_PAG("req_RequestId") = trx.RequestId
		t_PAG("req_Version") = trx.Version
		t_PAG("req_OrderData_MerchantId") = trx.OrderData_MerchantId
		t_PAG("req_OrderData_OrderId") = trx.OrderData_OrderId
		t_PAG("req_CustomerData_CustomerIdentity") = trx.CustomerData_CustomerIdentity
		t_PAG("req_CustomerData_CustomerName") = trx.CustomerData_CustomerName
		t_PAG.Update

		intPaymentOrdem = 0
		for i=Lbound(v_trx_payment) to Ubound(v_trx_payment)
			if Trim(v_trx_payment(i).PAG_PaymentMethod) <> "" then
				intPaymentOrdem = intPaymentOrdem + 1
				if alerta = "" then
				'	GERA O NSU P/ GRAVAR NO BD O REGISTRO C/ AS INFORMAÇÕES DA TRANSAÇÃO
					if Not fin_gera_nsu(T_PAGTO_GW_PAG_PAYMENT, idPagtoGwPagPayment, msg_erro) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "FALHA AO GERAR NSU PARA O NOVO REGISTRO DE TRANSAÇÃO (" & msg_erro & ")"
					elseif idPagtoGwPagPayment <= 0 then
						alerta=texto_add_br(alerta)
						alerta=alerta & "NSU GERADO É INVÁLIDO (" & idPagtoGwPagPayment & ")"
						end if
					end if

				if alerta = "" then
					s = "SELECT * FROM t_PAGTO_GW_PAG_PAYMENT WHERE (id = -1)"
					if t_PAG_PAYMENT.State <> 0 then t_PAG_PAYMENT.Close
					t_PAG_PAYMENT.Open s, cn
					t_PAG_PAYMENT.AddNew
					t_PAG_PAYMENT("id") = idPagtoGwPagPayment
					t_PAG_PAYMENT("id_pagto_gw_pag") = idPagtoGwPag
					t_PAG_PAYMENT("ordem") = intPaymentOrdem
					t_PAG_PAYMENT("bandeira") = v_trx_payment(i).bandeira
				'	ANOTA FIELMENTE OS DADOS DO CARTÃO DIGITADOS PELO CLIENTE, INCLUSIVE P/ USO POSTERIOR P/ REQUISIÇÃO DO ANTIFRAUDE
				'	O ARRAY vDadosCartao COMEÇA DA POSIÇÃO 1
					t_PAG_PAYMENT("valor_transacao") = vDadosCartao(intPaymentOrdem).vl_pagamento
					t_PAG_PAYMENT("checkout_opcao_parcelamento") = vDadosCartao(intPaymentOrdem).opcao_parcelamento
					t_PAG_PAYMENT("checkout_titular_nome") = vDadosCartao(intPaymentOrdem).titular_nome
					t_PAG_PAYMENT("checkout_titular_cpf_cnpj") = vDadosCartao(intPaymentOrdem).titular_cpf_cnpj
					t_PAG_PAYMENT("checkout_cartao_numero") = BraspagCSProtegeNumeroCartao(vDadosCartao(intPaymentOrdem).cartao_numero)
					t_PAG_PAYMENT("checkout_cartao_validade_mes") = vDadosCartao(intPaymentOrdem).cartao_validade_mes
					t_PAG_PAYMENT("checkout_cartao_validade_ano") = vDadosCartao(intPaymentOrdem).cartao_validade_ano
					t_PAG_PAYMENT("checkout_cartao_codigo_seguranca") = String(Len(vDadosCartao(intPaymentOrdem).cartao_codigo_seguranca), "*")
					if vDadosCartao(intPaymentOrdem).cartao_proprio = "PROPRIO" then
						t_PAG_PAYMENT("checkout_cartao_proprio") = "S"
					else
						t_PAG_PAYMENT("checkout_cartao_proprio") = "N"
						end if
					t_PAG_PAYMENT("checkout_fatura_end_logradouro") = vDadosCartao(intPaymentOrdem).fatura_end_logradouro
					t_PAG_PAYMENT("checkout_fatura_end_numero") = vDadosCartao(intPaymentOrdem).fatura_end_numero
					t_PAG_PAYMENT("checkout_fatura_end_complemento") = vDadosCartao(intPaymentOrdem).fatura_end_complemento
					t_PAG_PAYMENT("checkout_fatura_end_bairro") = vDadosCartao(intPaymentOrdem).fatura_end_bairro
					t_PAG_PAYMENT("checkout_fatura_end_cidade") = vDadosCartao(intPaymentOrdem).fatura_end_cidade
					t_PAG_PAYMENT("checkout_fatura_end_uf") = vDadosCartao(intPaymentOrdem).fatura_end_uf
					t_PAG_PAYMENT("checkout_fatura_end_cep") = vDadosCartao(intPaymentOrdem).fatura_end_cep
					t_PAG_PAYMENT("checkout_fatura_tel_pais") = c_fatura_telefone_pais
					t_PAG_PAYMENT("checkout_fatura_tel_ddd") = vDadosCartao(intPaymentOrdem).fatura_tel_ddd
					t_PAG_PAYMENT("checkout_fatura_tel_numero") = vDadosCartao(intPaymentOrdem).fatura_tel_numero
					t_PAG_PAYMENT("checkout_email") = r_cliente.email
				'	CAMPOS COM DADOS A SEREM ENVIADOS AO PAGADOR
					t_PAG_PAYMENT("req_PaymentDataRequest_PaymentMethod") = v_trx_payment(i).PAG_PaymentMethod
					t_PAG_PAYMENT("req_PaymentDataRequest_Amount") = v_trx_payment(i).PAG_Amount
					t_PAG_PAYMENT("req_PaymentDataRequest_Currency") = v_trx_payment(i).PAG_Currency
					t_PAG_PAYMENT("req_PaymentDataRequest_Country") = v_trx_payment(i).PAG_Country
					t_PAG_PAYMENT("req_PaymentDataRequest_ServiceTaxAmount") = v_trx_payment(i).PAG_ServiceTaxAmount
					t_PAG_PAYMENT("req_PaymentDataRequest_NumberOfPayments") = v_trx_payment(i).PAG_NumberOfPayments
					t_PAG_PAYMENT("req_PaymentDataRequest_PaymentPlan") = v_trx_payment(i).PAG_PaymentPlan
					t_PAG_PAYMENT("req_PaymentDataRequest_TransactionType") = v_trx_payment(i).PAG_TransactionType
					t_PAG_PAYMENT("req_PaymentDataRequest_CardHolder") = v_trx_payment(i).PAG_CardHolder
					t_PAG_PAYMENT("req_PaymentDataRequest_CardNumber") = BraspagCSProtegeNumeroCartao(v_trx_payment(i).PAG_CardNumber)
					t_PAG_PAYMENT("req_PaymentDataRequest_CardSecurityCode") = String(Len(v_trx_payment(i).PAG_CardSecurityCode), "*")
					t_PAG_PAYMENT("req_PaymentDataRequest_CardExpirationDate") = v_trx_payment(i).PAG_CardExpirationDate
					t_PAG_PAYMENT.Update
					end if
				end if
			next

		if alerta = "" then
			s = "SELECT * FROM t_PAGTO_GW_PAG_XML WHERE (id = -1)"
			if t_PAG_XML.State <> 0 then t_PAG_XML.Close
			t_PAG_XML.Open s, cn
			t_PAG_XML.AddNew
			t_PAG_XML("id") = idPagtoGwPagXmlTx
			t_PAG_XML("id_pagto_gw_pag") = idPagtoGwPag
			t_PAG_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__REQ_AUTHORIZE
			t_PAG_XML("fluxo_xml") = BRASPAG_FLUXO_XML__TX
			t_PAG_XML("xml") = txXmlMasked
			t_PAG_XML.Update
			end if
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			end if
		end if

	dim r_rx
	set r_rx = new cl_BRASPAG_Authorize_RX

	dim v_rx_PAG_Error
	redim v_rx_PAG_Error(0)
	set v_rx_PAG_Error(UBound(v_rx_PAG_Error)) = new cl_BRASPAG_Authorize_RX_PAG_ERROR
	
	dim blnRequisicaoErroStatus
	dim strRequisicaoMsgErro
	blnRequisicaoErroStatus = False
	strRequisicaoMsgErro = ""
	
	dim rxXml
	dim strTipoRetorno
	dim objXML, blnNodeNotFound, strNodeName, strNodeValue
	dim strErrorCode, strErrorMessage
	strErrorCode = ""
	strErrorMessage = ""
	
	if alerta = "" then
	'	ATUALIZA A DATA/HORA USANDO O RELÓGIO DO SQL SERVER
		s = "UPDATE t_PAGTO_GW_PAG SET" & _
				" trx_TX_data = Convert(varchar(10),getdate(), 121)," & _
				" trx_TX_data_hora = getdate()" & _
			" WHERE" & _
				" (id = " & Cstr(idPagtoGwPag) & ")"
		cn.Execute s, lngRecordsAffected


	'	TRANSMITE A TRANSAÇÃO
	'	~~~~~~~~~~~~~~~~~~~~~
	'	LOG
		s_log_dados_cartao = ""
		for i = 1 to qtde_cartoes
			if s_log_dados_cartao <> "" then s_log_dados_cartao = s_log_dados_cartao & chr(13)
			s_log_dados_cartao = s_log_dados_cartao & _
						"Cartão " & Cstr(i) & ": " & _
						"bandeira: " & BraspagDescricaoBandeira(vDadosCartao(i).bandeira) & _
						", valor: " & formata_moeda(vDadosCartao(i).vl_pagamento) & _
						", opção pagamento: " & vDadosCartao(i).descricao_parcelamento
			next
		s_log = "Preparando requisição 'Authorize' (t_PAGTO_GW_PAG.id=" & Cstr(idPagtoGwPag) & "): Qtde cartões usada no pagamento: " & Cstr(qtde_cartoes) & chr(13) & s_log_dados_cartao
		grava_log usuario, loja, pedido_selecionado, "", OP_LOG_PEDIDO_TRX_BRASPAG, s_log

		On Error Resume Next
		rxXml = BraspagEnviaTransacaoComRetry(txXml, BRASPAG_WS_ENDERECO_PAGADOR_TRANSACTION)
		if Err.number <> 0 then
			err_number = Err.number
			err_description = Err.Description
			alerta = "Falha ao tentar comunicar com o gateway de pagamentos!" & chr(13) & _
					"<br />" & chr(13) & _
					"<br />" & chr(13) & _
					"Código: " & Cstr(err_number) & chr(13) & _
					"<br />" & chr(13) & _
					"Descrição: " & err_description & chr(13) & _
					"<br />" & chr(13) & _
					"<br />" & chr(13) & _
					"Por favor, tente novamente em alguns instantes." & chr(13)
		'	ARMAZENA A MENSAGEM QUE DEVE SER EXIBIDA NA PÁGINA INFORMATIVA
		'	LEMBRANDO QUE NAS PÁGINAS ACESSADAS DIRETAMENTE PELOS CLIENTES PARA FAZER O PAGAMENTO NÃO SE DEVE USAR 'SESSION',
		'	JÁ QUE OCORRERAM VÁRIOS CASOS DE CLIENTES QUE NÃO CONSEGUIRAM INICIAR A SESSÃO (PROVAVELMENTE POR PROBLEMAS NA CONFIGURAÇÃO DE COOKIES OU MESMO DEVIDO A ANTIVIRUS)
			s = "SELECT * FROM t_PAGTO_GW_PAG WHERE (id = " & idPagtoGwPag & ")"
			if t_PAG.State <> 0 then t_PAG.Close
			t_PAG.Open s, cn
			if Not t_PAG.Eof then
				t_PAG("msg_alerta_tela") = alerta
				t_PAG.Update
				end if
			
		'	LOG
			s_log = "Erro na transação (t_PAGTO_GW_PAG.id=" & Cstr(idPagtoGwPag) & "): " & Cstr(err_number) & " - " & err_description
			grava_log usuario, loja, pedido_selecionado, "", OP_LOG_PEDIDO_TRX_BRASPAG, s_log
			
		'	REDIRECIONAMENTO P/ PÁGINA DE EXIBIÇÃO DE MENSAGEM DE ERRO (EXIBIR A MENSAGEM NESTA MESMA PÁGINA AUMENTA O RISCO DO USUÁRIO ATUALIZAR A PÁGINA E EXECUTAR NOVAMENTE A TRANSAÇÃO)
			Response.Redirect("P090MsgErroPrepara.asp?idPagtoGwPag=" & criptografa(Cstr(idPagtoGwPag)) & "&pedido=" & pedido_selecionado)
			end if
		
		On Error GoTo 0
		
		if Not fin_gera_nsu(T_PAGTO_GW_PAG_XML, idPagtoGwPagXmlRx, msg_erro) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		elseif idPagtoGwPagXmlRx <= 0 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "NSU GERADO É INVÁLIDO (" & idPagtoGwPagXmlRx & ")"
			end if
		
	'	ATUALIZA A DATA/HORA USANDO O RELÓGIO DO SQL SERVER
		s = "UPDATE t_PAGTO_GW_PAG SET" & _
				" trx_RX_status = 1," & _
				" trx_RX_id_pagto_gw_pag_xml = " & Cstr(idPagtoGwPagXmlRx) & "," & _
				" trx_RX_data = Convert(varchar(10),getdate(), 121)," & _
				" trx_RX_data_hora = getdate()" & _
			" WHERE" & _
				" (id = " & Cstr(idPagtoGwPag) & ")"
		cn.Execute s, lngRecordsAffected
		
	'	GRAVA NO HISTÓRICO O XML COMPLETO RECEBIDO
		s = "SELECT * FROM t_PAGTO_GW_PAG_XML WHERE (id = -1)"
		if t_PAG_XML.State <> 0 then t_PAG_XML.Close
		t_PAG_XML.Open s, cn
		t_PAG_XML.AddNew
		t_PAG_XML("id") = idPagtoGwPagXmlRx
		t_PAG_XML("id_pagto_gw_pag") = idPagtoGwPag
		t_PAG_XML("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__REQ_AUTHORIZE
		t_PAG_XML("fluxo_xml") = BRASPAG_FLUXO_XML__RX
		t_PAG_XML("xml") = rxXml
		t_PAG_XML.Update
		
		if Trim(rxXml) = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Ocorreu um erro inesperado: a resposta da transação está com conteúdo vazio.<br />Favor reportar o problema com a seguinte informação: ID = " & idPagtoGwPag
		'	ANOTA NO REGISTRO A OCORRÊNCIA DE RECEBIMENTO DE RESPOSTA VAZIA
			s = "UPDATE t_PAGTO_GW_PAG SET" & _
					" trx_erro_status = 1," & _
					" trx_RX_vazio_status = 1," & _
					" trx_erro_mensagem = 'RESPOSTA DA TRANSAÇÃO ESTÁ COM CONTEÚDO VAZIO'" & _
				" WHERE" & _
					" (id = " & Cstr(idPagtoGwPag) & ")"
			cn.Execute s, lngRecordsAffected
			end if
		
		Set objXML = Server.CreateObject("MSXML2.DOMDocument.3.0")
		objXML.Async = False
		objXML.LoadXml(rxXml)
		
	'	REGISTRA O QUANTO ANTES O 'TRANSACTION ID' P/ POSSIBILITAR CONSULTAS POSTERIORES MESMO NO CASO DE OCORRER UM ERRO QUALQUER NESTA PÁGINA (NÃO SALVANDO TODAS AS INFORMAÇÕES CORRETAMENTE)
		intSequenciaPaymentList = 0
		set oNodePaymentList=objXML.documentElement.selectNodes(XML_PATH_AuthorizeTransactionResult & "/PaymentDataCollection/PaymentDataResponse")
		if Not oNodePaymentList is nothing then
			for each oNodeSet in oNodePaymentList
				intSequenciaPaymentList = intSequenciaPaymentList + 1
				strBraspagTransactionId = ""
				strPaymentMethod = ""
				strAmount = ""
			'	BraspagTransactionId
				strNodeName = "BraspagTransactionId"
				strBraspagTransactionId = Trim(xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound))
			'	PaymentMethod
				strNodeName = "PaymentMethod"
				strPaymentMethod = Trim(xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound))
			'	Amount
				strNodeName = "Amount"
				strAmount = Trim(xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound))
			'	GRAVA O 'TRANSACTION ID'
				if qtde_cartoes = 1 then
					if strBraspagTransactionId<>"" then
					'	APENAS POR SEGURANÇA ADICIONAL, RESTRINGE TAMBÉM PELO 'PaymentMethod' E 'Amount'
						s = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" & _
								" resp_PaymentDataResponse_BraspagTransactionId = '" & strBraspagTransactionId & "'" & _
							" WHERE" & _
								" (id_pagto_gw_pag = " & idPagtoGwPag & ")" & _
								" AND (req_PaymentDataRequest_PaymentMethod = '" & strPaymentMethod & "')" & _
								" AND (req_PaymentDataRequest_Amount = '" & strAmount & "')"
						cn.Execute s, lngRecordsAffected
						if lngRecordsAffected <> 1 then
							s_log = "Falha ao anotar o campo BraspagTransactionId (Payment nº " & Cstr(intSequenciaPaymentList) & "): foram encontrados " & Cstr(lngRecordsAffected) & " registros com os seguintes dados: id_pagto_gw_pag=" & Cstr(idPagtoGwPag) & " (transação usando " & Cstr(qtde_cartoes) & " cartões)"
							grava_log usuario, loja, pedido_selecionado, "", OP_LOG_PEDIDO_TRX_BRASPAG_FALHA, s_log
							end if
						end if
				else
				'	MAIS DO QUE 1 CARTÃO: SERÁ NECESSÁRIO IDENTIFICAR QUAL É O CARTÃO
				'	IMPORTANTE: ASSUME-SE QUE FOI FEITA A VERIFICAÇÃO QUE IMPEDE O CLIENTE DE USAR 2 CARTÕES DA MESMA BANDEIRA P/ PAGAR VALORES IDÊNTICOS
					if (strBraspagTransactionId<>"") And (strPaymentMethod <> "") And (strAmount <> "") then
						s = "SELECT" & _
								" COUNT(*) AS qtde" & _
							" FROM t_PAGTO_GW_PAG_PAYMENT" & _
							" WHERE" & _
								" (id_pagto_gw_pag = " & idPagtoGwPag & ")" & _
								" AND (req_PaymentDataRequest_PaymentMethod = '" & strPaymentMethod & "')" & _
								" AND (req_PaymentDataRequest_Amount = '" & strAmount & "')"
						if t_PAG_PAYMENT.State <> 0 then t_PAG_PAYMENT.Close
						t_PAG_PAYMENT.Open s, cn
						if Not t_PAG_PAYMENT.Eof then
							if CLng(t_PAG_PAYMENT("qtde")) > 1 then
								blnAmbiguidade = True
							else
								blnAmbiguidade = False
								end if
							s = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" & _
									" resp_PaymentDataResponse_BraspagTransactionId = '" & strBraspagTransactionId & "'" & _
								" WHERE" & _
									" (id_pagto_gw_pag = " & idPagtoGwPag & ")" & _
									" AND (req_PaymentDataRequest_PaymentMethod = '" & strPaymentMethod & "')" & _
									" AND (req_PaymentDataRequest_Amount = '" & strAmount & "')"
						'	SE HOUVER AMBIGUIDADE, TENTA RESOLVER PELA POSIÇÃO EM QUE A TRANSAÇÃO FOI ENVIADA E RETORNADA DENTRO DA COLEÇÃO
							if blnAmbiguidade then
								s = s & _
									" AND (ordem = " & CStr(intSequenciaPaymentList) & ")"
								end if
							cn.Execute s, lngRecordsAffected
							if lngRecordsAffected <> 1 then
								s_log = "Falha ao anotar o campo BraspagTransactionId (Payment nº " & Cstr(intSequenciaPaymentList) & "): foram encontrados " & Cstr(lngRecordsAffected) & " registros com os seguintes dados: id_pagto_gw_pag=" & Cstr(idPagtoGwPag) & ", PaymentMethod=" & strPaymentMethod & ", Amount=" & strAmount & " (transação usando " & Cstr(qtde_cartoes) & " cartões)"
								grava_log usuario, loja, pedido_selecionado, "", OP_LOG_PEDIDO_TRX_BRASPAG_FALHA, s_log
								end if
							end if
						
						if t_PAG_PAYMENT.State <> 0 then t_PAG_PAYMENT.Close
						end if
					end if
				next
			end if
		end if

	if alerta = "" then
		s = "SELECT * FROM t_PAGTO_GW_PAG WHERE (id = " & idPagtoGwPag & ")"
		if t_PAG.State <> 0 then t_PAG.Close
		t_PAG.Open s, cn
		if t_PAG.Eof then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Falha ao localizar no BD o registro da transação de pagamento (t_PAGTO_GW_PAG.Id=" & idPagtoGwPag & ")."
			end if
		end if
	
	dim blnErroFatalTransacaoBD
	blnErroFatalTransacaoBD = False
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if Instr(Ucase(rxXml), Ucase("SSL certificate problem")) <> 0 then
			msg_erro="CERTIFICADO INVÁLIDO - O certificado da transação não foi aprovado!"
			t_PAG("trx_erro_status") = 1
			t_PAG("trx_erro_mensagem") = msg_erro
			alerta=texto_add_br(alerta)
			alerta=alerta & msg_erro
			end if

		if alerta = "" then
			strTipoRetorno = objXML.documentElement.baseName
			if Ucase(strTipoRetorno) = "ERRO" then
			'	~~~~~~~~~~~~~~
			'	RESPOSTA: ERRO
			'	~~~~~~~~~~~~~~
				t_PAG("trx_erro_status") = 1
			'	CÓDIGO DO ERRO
				strNodeValue = xmlReadNode(objXml, "codigo", blnNodeNotFound)
				t_PAG("trx_erro_codigo") = strNodeValue
			'	MENSAGEM DE ERRO
				strNodeValue = xmlReadNode(objXml, "mensagem", blnNodeNotFound)
				t_PAG("trx_erro_mensagem") = strNodeValue
				alerta=texto_add_br(alerta)
				alerta=alerta & "Resposta do servidor da Braspag:" & "<br>" & Trim("" & t_PAG("trx_erro_codigo")) & " - " & Trim("" & t_PAG("trx_erro_mensagem"))
			'	LOG
				s_log = "Transação retornou erro (t_PAGTO_GW_PAG.id=" & Cstr(idPagtoGwPag) & "): " & Trim("" & t_PAG("trx_erro_codigo")) & " - " & Trim("" & t_PAG("trx_erro_mensagem"))
				grava_log usuario, loja, pedido_selecionado, "", OP_LOG_PEDIDO_TRX_BRASPAG, s_log
			elseif Ucase(strTipoRetorno) = "ENVELOPE" then
			'	~~~~~~~~~~~~~~~~~~~
			'	RESPOSTA: ENVELOPE
			'	~~~~~~~~~~~~~~~~~~~
			'	Houve falha?
				strNodeName = "//faultcode"
				strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
				if Not blnNodeNotFound then
					blnRequisicaoErroStatus = True
					strRequisicaoMsgErro=texto_add_cr(strRequisicaoMsgErro)
					strRequisicaoMsgErro=strRequisicaoMsgErro & texto_add_cr("Falha na requisição!!")
					strRequisicaoMsgErro=strRequisicaoMsgErro & "Código: " & strNodeValue
					end if
				
				strNodeName = "//faultstring"
				strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
				if Not blnNodeNotFound then
					if Not blnRequisicaoErroStatus then
						blnRequisicaoErroStatus = True
						strRequisicaoMsgErro=texto_add_cr(strRequisicaoMsgErro)
						strRequisicaoMsgErro=strRequisicaoMsgErro & texto_add_cr("Falha na requisição!!")
						end if
					strRequisicaoMsgErro=texto_add_cr(strRequisicaoMsgErro)
					strRequisicaoMsgErro=strRequisicaoMsgErro & "Descrição: " & strNodeValue
					end if
				
			'	CorrelationId
				strNodeName = XML_PATH_AuthorizeTransactionResult & "/CorrelationId"
				strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
				r_rx.PAG_CorrelationId = strNodeValue
				if Not blnNodeNotFound then
					t_PAG("resp_CorrelationId") = strNodeValue
					end if
				
			'	Success
				strNodeName = XML_PATH_AuthorizeTransactionResult & "/Success"
				strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
				r_rx.PAG_Success = strNodeValue
				if Not blnNodeNotFound then
					t_PAG("resp_Success") = strNodeValue
					end if
				
			'	OrderId
				strNodeName = XML_PATH_AuthorizeTransactionResult & "/OrderData/OrderId"
				strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
				r_rx.PAG_OrderData_OrderId = strNodeValue
				if Not blnNodeNotFound then
					t_PAG("resp_OrderData_OrderId") = strNodeValue
					end if
				
			'	BraspagOrderId
				strNodeName = XML_PATH_AuthorizeTransactionResult & "/OrderData/BraspagOrderId"
				strNodeValue = xmlReadNode(objXml, strNodeName, blnNodeNotFound)
				r_rx.PAG_OrderData_BraspagOrderId = strNodeValue
				if Not blnNodeNotFound then
					t_PAG("resp_OrderData_BraspagOrderId") = strNodeValue
					end if
				
			'	PaymentDataCollection
				intSequenciaPaymentList = 0
				set oNodePaymentList=objXML.documentElement.selectNodes(XML_PATH_AuthorizeTransactionResult & "/PaymentDataCollection/PaymentDataResponse")
				if Not oNodePaymentList is nothing then
					for each oNodeSet in oNodePaymentList
						intSequenciaPaymentList = intSequenciaPaymentList + 1
						
						if Trim(v_trx_payment_RX(UBound(v_trx_payment_RX)).PAG_PaymentMethod) <> "" then
							redim preserve v_trx_payment_RX(UBound(v_trx_payment_RX)+1)
							set v_trx_payment_RX(UBound(v_trx_payment_RX)) = new cl_BRASPAG_Authorize_PaymentDataRequest_RX
							end if

					'	BraspagTransactionId
						strNodeName = "BraspagTransactionId"
						strBraspagTransactionId = Trim(xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound))
						v_trx_payment_RX(UBound(v_trx_payment_RX)).PAG_BraspagTransactionId = strBraspagTransactionId

					'	PaymentMethod
						strNodeName = "PaymentMethod"
						strPaymentMethod = Trim(xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound))
						v_trx_payment_RX(UBound(v_trx_payment_RX)).PAG_PaymentMethod = strPaymentMethod

					'	Amount
						strNodeName = "Amount"
						strAmount = Trim(xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound))
						v_trx_payment_RX(UBound(v_trx_payment_RX)).PAG_Amount = strAmount

						if qtde_cartoes = 1 then
							s_sql = "SELECT * FROM t_PAGTO_GW_PAG_PAYMENT WHERE (id_pagto_gw_pag = " & Cstr(idPagtoGwPag) & ")"
						else
						'	MAIS DO QUE 1 CARTÃO: SERÁ NECESSÁRIO IDENTIFICAR QUAL É O CARTÃO
						'	IMPORTANTE: ASSUME-SE QUE FOI FEITA A VERIFICAÇÃO QUE IMPEDE O CLIENTE DE USAR 2 CARTÕES DA MESMA BANDEIRA P/ PAGAR VALORES IDÊNTICOS
							s_sql = "SELECT " & _
										"*" & _
									" FROM t_PAGTO_GW_PAG_PAYMENT" & _
									" WHERE" & _
										" (id_pagto_gw_pag = " & Cstr(idPagtoGwPag) & ")" & _
										" AND (req_PaymentDataRequest_PaymentMethod = '" & strPaymentMethod & "')" & _
										" AND (req_PaymentDataRequest_Amount = '" & strAmount & "')"
							end if

						if t_PAG_PAYMENT.State <> 0 then t_PAG_PAYMENT.Close
						t_PAG_PAYMENT.Open s_sql, cn
						if t_PAG_PAYMENT.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Falha ao tentar localizar o registro da transação de pagamento (id_pagto_gw_pag=" & Cstr(idPagtoGwPag) & ", PaymentMethod=" & strPaymentMethod & ", Amount=" & strAmount & ")"
						else
							if t_PAG_PAYMENT.RecordCount > 1 then
								blnAmbiguidade = True
							else
								blnAmbiguidade = False
								end if

						'	SE HOUVER AMBIGUIDADE, TENTA RESOLVER PELA POSIÇÃO EM QUE A TRANSAÇÃO FOI ENVIADA E RETORNADA DENTRO DA COLEÇÃO
							if blnAmbiguidade then
								s_sql = s_sql & _
											" AND (ordem = " & CStr(intSequenciaPaymentList) & ")"
								if t_PAG_PAYMENT.State <> 0 then t_PAG_PAYMENT.Close
								t_PAG_PAYMENT.Open s_sql, cn
								end if
							end if
						
						if Not t_PAG_PAYMENT.Eof then
							idPagtoGwPagPayment = t_PAG_PAYMENT("id")
							vl_transacao = t_PAG_PAYMENT("valor_transacao")

						'	BraspagTransactionId / PaymentMethod / Amount
							t_PAG_PAYMENT("resp_PaymentDataResponse_BraspagTransactionId") = strBraspagTransactionId
							t_PAG_PAYMENT("resp_PaymentDataResponse_PaymentMethod") = strPaymentMethod
							t_PAG_PAYMENT("resp_PaymentDataResponse_Amount") = strAmount
							
						'	Bandeira
						'	Obs: recupera a informação da bandeira a partir do BD porque na transação é usado um código da Braspag, sendo que no ambiente de homologação o código é sempre "997",
						'	o que dificulta a depuração por não ser possível identificar a bandeira nos casos de pagamento com 2 ou mais cartões.
							v_trx_payment_RX(UBound(v_trx_payment_RX)).bandeira = Trim("" & t_PAG_PAYMENT("bandeira"))

						'	AcquirerTransactionId
							strNodeName = "AcquirerTransactionId"
							strNodeValue = xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound)
							if Not blnNodeNotFound then
								t_PAG_PAYMENT("resp_PaymentDataResponse_AcquirerTransactionId") = strNodeValue
								v_trx_payment_RX(UBound(v_trx_payment_RX)).PAG_AcquirerTransactionId = strNodeValue
								end if
						
						'	AuthorizationCode
							strNodeName = "AuthorizationCode"
							strNodeValue = xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound)
							if Not blnNodeNotFound then
								t_PAG_PAYMENT("resp_PaymentDataResponse_AuthorizationCode") = strNodeValue
								v_trx_payment_RX(UBound(v_trx_payment_RX)).PAG_AuthorizationCode = strNodeValue
								end if
						
						'	CreditCardToken
							strNodeName = "CreditCardToken"
							strNodeValue = xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound)
							if Not blnNodeNotFound then
								t_PAG_PAYMENT("resp_PaymentDataResponse_CreditCardToken") = strNodeValue
								v_trx_payment_RX(UBound(v_trx_payment_RX)).PAG_CreditCardToken = strNodeValue
								end if
						
						'	ProofOfSale
							strNodeName = "ProofOfSale"
							strNodeValue = xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound)
							if Not blnNodeNotFound then
								t_PAG_PAYMENT("resp_PaymentDataResponse_ProofOfSale") = strNodeValue
								v_trx_payment_RX(UBound(v_trx_payment_RX)).PAG_ProofOfSale = strNodeValue
								end if
						
						'	ReturnCode
							strNodeName = "ReturnCode"
							strNodeValue = xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound)
							if Not blnNodeNotFound then
								t_PAG_PAYMENT("resp_PaymentDataResponse_ReturnCode") = strNodeValue
								v_trx_payment_RX(UBound(v_trx_payment_RX)).PAG_ReturnCode = strNodeValue
								end if
						
						'	ReturnMessage
							strNodeName = "ReturnMessage"
							strNodeValue = xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound)
							if Not blnNodeNotFound then
								t_PAG_PAYMENT("resp_PaymentDataResponse_ReturnMessage") = strNodeValue
								v_trx_payment_RX(UBound(v_trx_payment_RX)).PAG_ReturnMessage = strNodeValue
								end if
						
						'	Status
							strStatus = ""
							strNodeName = "Status"
							strNodeValue = xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound)
							if Not blnNodeNotFound then
								strStatus = strNodeValue
								v_trx_payment_RX(UBound(v_trx_payment_RX)).PAG_Status = strNodeValue
								t_PAG_PAYMENT("resp_PaymentDataResponse_Status") = strNodeValue
								t_PAG_PAYMENT("prim_GlobalStatus") = decodifica_PaymentDataResponseStatus_para_GlobalStatus(strNodeValue)
								t_PAG_PAYMENT("prim_atualizacao_data_hora") = Now
								t_PAG_PAYMENT("prim_atualizacao_usuario") = usuario
								t_PAG_PAYMENT("ult_GlobalStatus") = decodifica_PaymentDataResponseStatus_para_GlobalStatus(strNodeValue)
								t_PAG_PAYMENT("ult_atualizacao_data_hora") = Now
								t_PAG_PAYMENT("ult_atualizacao_usuario") = usuario
								if strNodeValue = BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__AUTORIZADA then
									t_PAG_PAYMENT("resp_AuthorizedDate") = Date
								elseif strNodeValue = BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__CAPTURADA then
								'	SE A TRANSAÇÃO JÁ FOI CAPTURADA, ISSO IMPLICA EM TER SIDO AUTORIZADA TAMBÉM
									t_PAG_PAYMENT("resp_AuthorizedDate") = Date
									t_PAG_PAYMENT("resp_CapturedDate") = Date
									t_PAG_PAYMENT("captura_confirmada_status") = 1
									t_PAG_PAYMENT("captura_confirmada_data") = Date
									t_PAG_PAYMENT("captura_confirmada_data_hora") = Now
									end if
								end if
						
							t_PAG_PAYMENT.Update

							if strStatus = BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__AUTORIZADA then
							'	GERA OS REGISTROS EM t_FIN_PEDIDO_HIST_PAGTO
								if Not BraspagClearsaleGeraPedidoHistPagto(BRASPAG_REGISTRA_PAGTO__OP_AUTORIZACAO, id_pedido_base, idPagtoGwPagPayment, vl_transacao, usuario, msg_erro) then
									alerta=texto_add_br(alerta)
									alerta=alerta & "Falha ao tentar gravar o histórico de pagamento no pedido " & id_pedido_base & "<br />" & msg_erro
									end if
							'	SE O STATUS DA ANÁLISE DE CRÉDITO DO PEDIDO ESTIVER COM O STATUS INICIAL, COLOCA EM 'PENDENTE VENDAS'
								if Not blnStAnaliseCreditoAlterado then
									if Trim(id_pedido_base) <> "" then
										s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & id_pedido_base & "')"
										if t_PEDIDO.State <> 0 then t_PEDIDO.Close
										t_PEDIDO.Open s, cn
										if Not t_PEDIDO.Eof then
											if CLng(t_PEDIDO("analise_credito")) = CLng(COD_AN_CREDITO_ST_INICIAL) then
												t_PEDIDO("analise_credito") = CLng(COD_AN_CREDITO_PENDENTE_VENDAS)
												t_PEDIDO("analise_credito_data") = Now
												t_PEDIDO("analise_credito_usuario") = ID_USUARIO_SISTEMA
												t_PEDIDO.Update
												blnStAnaliseCreditoAlterado = True
												end if
											end if
										end if
									end if
							elseif strStatus = BRASPAG_PAGADOR_CARTAO_PAYMENTDATARESPONSE_STATUS__CAPTURADA then
							'	FAZ A CHAMADA A BraspagClearsaleRegistraPagtoNoPedido() SOMENTE APÓS A EXECUÇÃO DE t_PAG_PAYMENT.Update PORQUE A ROTINA ACESSA E ALTERA OS DADOS DESSA TABELA
								if Not BraspagClearsaleRegistraPagtoNoPedido(BRASPAG_REGISTRA_PAGTO__OP_CAPTURA, id_pedido_base, idPagtoGwPagPayment, vl_transacao, usuario, msg_erro) then
									alerta=texto_add_br(alerta)
									alerta=alerta & "Falha ao tentar registrar o pagamento no pedido " & id_pedido_base & "<br />" & msg_erro
									end if
								end if
							end if 'if Not t_PAG_PAYMENT.Eof
						next
					end if
				
			'	ErrorCode/ErrorMessage
				intSequenciaErrorList = 0
				set oNodeErrorList=objXML.documentElement.selectNodes(XML_PATH_AuthorizeTransactionResult & "/ErrorReportDataCollection/ErrorReportDataResponse")
				if Not oNodeErrorList is nothing then
					for each oNodeSet in oNodeErrorList
					'	OBTÉM OS DADOS DO ERRO P/ VERIFICAR SE HÁ CONTEÚDO
					'	PAG: ErrorCode
						strNodeName = "ErrorCode"
						strErrorCode = Trim(xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound))
						if blnNodeNotFound then strErrorCode = ""
						
					'	PAG: ErrorMessage
						strNodeName = "ErrorMessage"
						strErrorMessage = Trim(xmlReadSubNode(oNodeSet, strNodeName, blnNodeNotFound))
						if blnNodeNotFound then strErrorMessage = ""
						
						if (strErrorCode <> "") Or (strErrorMessage <> "") then
							intSequenciaErrorList = intSequenciaErrorList + 1
							
						'	GERA O NSU
							if Not fin_gera_nsu(T_PAGTO_GW_PAG_ERROR, idPagtoGwPagError, msg_erro) then
								blnErroFatalTransacaoBD = True
								alerta=texto_add_br(alerta)
								alerta=alerta & "FALHA AO GERAR NSU PARA O NOVO REGISTRO DE 't_PAGTO_GW_PAG_ERROR' (" & msg_erro & ")"
							elseif idPagtoGwPagError <= 0 then
								blnErroFatalTransacaoBD = True
								alerta=texto_add_br(alerta)
								alerta=alerta & "NSU GERADO É INVÁLIDO (" & idPagtoGwPagError & ")"
								end if
							
							if alerta <> "" then exit for
							
							redim preserve v_rx_PAG_Error(UBound(v_rx_PAG_Error)+1)
							set v_rx_PAG_Error(UBound(v_rx_PAG_Error)) = new cl_BRASPAG_Authorize_RX_PAG_ERROR
							
							s = "SELECT * FROM t_PAGTO_GW_PAG_ERROR WHERE (id = -1)"
							if t_PAG_ERROR.State <> 0 then t_PAG_ERROR.Close
							t_PAG_ERROR.Open s, cn
							t_PAG_ERROR.AddNew
							t_PAG_ERROR("id") = idPagtoGwPagError
							t_PAG_ERROR("id_pagto_gw_pag") = idPagtoGwPag
							t_PAG_ERROR("sequencia") = intSequenciaErrorList
							
						'	ErrorCode
							v_rx_PAG_Error(UBound(v_rx_PAG_Error)).ErrorCode = strErrorCode
							t_PAG_ERROR("ErrorCode") = strErrorCode
							
						'	ErrorMessage
							v_rx_PAG_Error(UBound(v_rx_PAG_Error)).ErrorMessage = strErrorMessage
							t_PAG_ERROR("ErrorMessage") = strErrorMessage
							
							t_PAG_ERROR.Update

							if intSequenciaPaymentList = 0 then
							'	NÃO RETORNOU NENHUMA INFORMAÇÃO DE PaymentDataResponse, APENAS ErrorReportDataResponse
								blnRequisicaoErroStatus = True
								strRequisicaoMsgErro = strErrorCode & " - " & strErrorMessage
								end if
							end if
						next
					end if
				
				if blnRequisicaoErroStatus then
					t_PAG("trx_erro_status") = 1
					if (strErrorCode <> "") And (strErrorMessage <> "") then
						t_PAG("trx_erro_codigo") = strErrorCode
						t_PAG("trx_erro_mensagem") = strErrorMessage
					else
						t_PAG("trx_erro_mensagem") = strRequisicaoMsgErro
						end if
					alerta=texto_add_br(alerta)
					alerta=alerta & substitui_caracteres(strRequisicaoMsgErro, chr(13), "<br>")
				'	LOG
					s_log = "Falha na transação (t_PAGTO_GW_PAG.id=" & Cstr(idPagtoGwPag) & "): " & strRequisicaoMsgErro
					grava_log usuario, loja, pedido_selecionado, "", OP_LOG_PEDIDO_TRX_BRASPAG, s_log
				else
				'	LOG
					s_log_aux = ""
					for i=LBound(v_trx_payment_RX) to UBound(v_trx_payment_RX)
						if Trim(v_trx_payment_RX(i).PAG_PaymentMethod) <> "" then
							if s_log_aux <> "" then s_log_aux = s_log_aux & " " & chr(13)
							s_log_aux = s_log_aux & _
										" Bandeira: " & v_trx_payment_RX(i).bandeira & _
										", PaymentMethod: " & v_trx_payment_RX(i).PAG_PaymentMethod & _
										", Status (Pagador): " & v_trx_payment_RX(i).PAG_Status & " - " & BraspagPagadorDescricaoPaymentDataResponseStatus(v_trx_payment_RX(i).PAG_Status) & _
										", Amount: " & v_trx_payment_RX(i).PAG_Amount & _
										", BraspagTransactionId: " & v_trx_payment_RX(i).PAG_BraspagTransactionId & _
										", AcquirerTransactionId: " & v_trx_payment_RX(i).PAG_AcquirerTransactionId & _
										", AuthorizationCode: " & v_trx_payment_RX(i).PAG_AuthorizationCode & _
										", ProofOfSale: " & v_trx_payment_RX(i).PAG_ProofOfSale & _
										", ReturnCode: " & v_trx_payment_RX(i).PAG_ReturnCode & _
										", ReturnMessage: " & v_trx_payment_RX(i).PAG_ReturnMessage & _
										", CreditCardToken: " & v_trx_payment_RX(i).PAG_CreditCardToken
							end if
						next
					s_log = "Retorno da transação (t_PAGTO_GW_PAG.id=" & Cstr(idPagtoGwPag) & "):" & s_log_aux
					grava_log usuario, loja, pedido_selecionado, "", OP_LOG_PEDIDO_TRX_BRASPAG, s_log
					end if
			else
			'	~~~~~~~~~~~~~~~~~~~~~~
			'	RESPOSTA: DESCONHECIDA
			'	~~~~~~~~~~~~~~~~~~~~~~
				t_PAG("trx_erro_status") = 1
				t_PAG("trx_erro_mensagem") = "TIPO DE RESPOSTA DESCONHECIDO: " & strTipoRetorno
				alerta=texto_add_br(alerta)
				alerta=alerta & "RESPOSTA RECEBIDA ESTÁ COM FORMATO INVÁLIDO"
			'	LOG
				s_log = "Erro na transação (t_PAGTO_GW_PAG.id=" & Cstr(idPagtoGwPag) & "): " & Trim("" & t_PAG("trx_erro_mensagem"))
				grava_log usuario, loja, pedido_selecionado, "", OP_LOG_PEDIDO_TRX_BRASPAG, s_log
				end if ' if Ucase(strTipoRetorno) = "ENVELOPE"
			end if ' if alerta = ""
		
		t_PAG.Update
		
		if Not blnErroFatalTransacaoBD then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			end if
		end if ' if alerta = "" then
	
	if alerta <> "" then
		s_log = "Transação de pagamento com a Braspag (t_PAGTO_GW_PAG.id=" & Cstr(idPagtoGwPag) & ") resultou em falha: " & chr(13) & alerta
		grava_log usuario, loja, pedido_selecionado, "", OP_LOG_PEDIDO_TRX_BRASPAG_FALHA, s_log
		end if
		
	if alerta <> "" then
	'	ARMAZENA A MENSAGEM QUE DEVE SER EXIBIDA NA PÁGINA INFORMATIVA
	'	LEMBRANDO QUE NAS PÁGINAS ACESSADAS DIRETAMENTE PELOS CLIENTES PARA FAZER O PAGAMENTO NÃO SE DEVE USAR 'SESSION',
	'	JÁ QUE OCORRERAM VÁRIOS CASOS DE CLIENTES QUE NÃO CONSEGUIRAM INICIAR A SESSÃO (PROVAVELMENTE POR PROBLEMAS NA CONFIGURAÇÃO DE COOKIES OU MESMO DEVIDO A ANTIVIRUS)
		s = "SELECT * FROM t_PAGTO_GW_PAG WHERE (id = " & idPagtoGwPag & ")"
		if t_PAG.State <> 0 then t_PAG.Close
		t_PAG.Open s, cn
		if Not t_PAG.Eof then
			t_PAG("msg_alerta_tela") = alerta
			t_PAG.Update
			end if
	else
	'	ASSEGURA QUE O CONTEÚDO DO CAMPO ESTEJA VAZIO
		s = "UPDATE t_PAGTO_GW_PAG SET msg_alerta_tela = NULL WHERE (id = " & idPagtoGwPag & ")"
		cn.Execute s, lngRecordsAffected
		end if

'	FECHA CONEXAO COM O BANCO DE DADOS
	if t_PAG.State <> 0 then t_PAG.Close
	set t_PAG=nothing

	if t_PAG_PAYMENT.State <> 0 then t_PAG_PAYMENT.Close
	set t_PAG_PAYMENT=nothing

	if t_PAG_XML.State <> 0 then t_PAG_XML.Close
	set t_PAG_XML=nothing

	if t_PAG_ERROR.State <> 0 then t_PAG_ERROR.Close
	set t_PAG_ERROR=nothing

	if t_PEDIDO.State <> 0 then t_PEDIDO.Close
	set t_PEDIDO=nothing

'	FECHA CONEXÃO
	cn.Close
	set cn = nothing

	if alerta = "" then
	'	REDIRECIONAMENTO P/ PÁGINA DE EXIBIÇÃO DE MENSAGEM COM O STATUS DA TRANSAÇÃO
		Response.Redirect("P090MsgResultadoPrepara.asp?idPagtoGwPag=" & criptografa(Cstr(idPagtoGwPag)) & "&pedido=" & pedido_selecionado)
	else
	'	REDIRECIONAMENTO P/ PÁGINA DE EXIBIÇÃO DE MENSAGEM DE ERRO (EXIBIR A MENSAGEM NESTA MESMA PÁGINA AUMENTA O RISCO DO USUÁRIO ATUALIZAR A PÁGINA E EXECUTAR NOVAMENTE A TRANSAÇÃO)
	'	A MENSAGEM DE ERRO ESTÁ GRAVADA EM t_PAGTO_GW_PAG.msg_alerta_tela
		Response.Redirect("P090MsgErroPrepara.asp?idPagtoGwPag=" & criptografa(Cstr(idPagtoGwPag)) & "&pedido=" & pedido_selecionado)
		end if






' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

sub exibe_erro_e_encerra(Byval mensagem)
dim x
	x = DOCTYPE_LEGADO & chr(13) & _
		"<html>" & chr(13) & _
		"	<head>" & chr(13) & _
		"		<title>" & SITE_CLIENTE_TITULO_JANELA & "</title>" & chr(13) & _
		"	</head>" & chr(13) & _
		"<script src=" & chr(34) & URL_FILE__GLOBAL_JS & chr(34) & " language=" & chr(34) & "JavaScript" & chr(34) & " type=" & chr(34)& "text/javascript" & chr(34) & "></script>" & chr(13) & _
		"<link href=" & chr(34) & URL_FILE__E_CSS & chr(34) & " rel=" & chr(34) & "stylesheet" & chr(34) & " type=" & chr(34) & "text/css" & chr(34) & ">" & chr(13) & _
		"<link href=" & chr(34) & URL_FILE__EPRINTER_CSS & chr(34) & " rel=" & chr(34) & "stylesheet" & chr(34) & " type=" & chr(34) & "text/css" & chr(34) & " media=" & chr(34) & "print" & chr(34) & ">" & chr(13) & _
		"<body onload=" & chr(34) & "bVOLTAR.focus();" & chr(34) & ">" & chr(13) & _
		"	<center>" & chr(13) & _
		"	<br />" & chr(13) & _
		"	<p class=" & chr(34) & "T" & chr(34) & ">A V I S O</p>" & chr(13) & _
		"	<div class=" & chr(34) & "MtAlerta" & chr(34) & " style=" & chr(34) & "width:600px;font-weight:bold;" & chr(34) & " align=" & chr(34) & "center" & chr(34) & ">" & chr(13) & _
				"<p style=" & chr(34) & "margin:5px 2px 5px 2px;" & chr(34) & ">" & _
					mensagem & _
				"</p>" & chr(13) & _
			"</div>" & chr(13) & _
		"	<br /><br />" & chr(13) & _
		"	<p class=" & chr(34) & "TracoBottom" & chr(34) & "></p>" & chr(13) & _
		"	<table cellspacing=" & chr(34) & "0" & chr(34) & ">" & chr(13) & _
		"		<tr>" & chr(13) & _
		"			<td align=" & chr(34) & "center" & chr(34) & ">" & chr(13) & _
						"<a name=" & chr(34) & "bVOLTAR" & chr(34) & " href=" & chr(34) & "javascript:Navega('../ClienteCartao/Id.asp')" & chr(34) & ">" & chr(13) & _
							"<img src=" & chr(34) & "../botao/voltar.gif" & chr(34) & " width=" & chr(34) & "176" & chr(34) & " height=" & chr(34) & "55" & chr(34) & " border=" & chr(34) & "0" & chr(34) & ">" & chr(13) & _
						"</a>" & chr(13) & _
					"</td>" & chr(13) & _
		"		</tr>" & chr(13) & _
		"	</table>" & chr(13) & _
		"	</center>" & chr(13) & _
		"</body>" & chr(13) & _
		"</html>" & chr(13)

	Response.Write x
	Response.End
end sub
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


<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__SSL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<% =strScriptWindowName %>

<script language="JavaScript" type="text/javascript">
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style type="text/css">
body::before
{
	content: '';
	border: none;
	margin-top: 0px;
	margin-bottom: 0px;
	padding: 0px;
}
</style>



<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
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
	<td align="center"><a name="bVOLTAR" href="javascript:Navega('../ClienteCartao/Id.asp')"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>

<% if SITE_CLIENTE_EXIBIR_LOGO_SSL then %>
<script language="JavaScript" type="text/javascript">
	logo_ssl_corner("../imagem/ssl/ssl_corner.gif");
</script>
<% end if %>

</body>




<% else %>
<!-- ****************************************************************** -->
<!-- **********          FALHA  DESCONHECIDA          ***************** -->
<!-- ****************************************************************** -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'>OCORREU UMA FALHA DESCONHECIDA!</p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" href="javascript:Navega('../ClienteCartao/Id.asp')"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>

<% if SITE_CLIENTE_EXIBIR_LOGO_SSL then %>
<script language="JavaScript" type="text/javascript">
	logo_ssl_corner("../imagem/ssl/ssl_corner.gif");
</script>
<% end if %>

</body>
<% end if %>

</html>