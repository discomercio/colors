<%
' =========================================
'          C O N S T A N T E S
' =========================================

const MAX_QTDE_CARTOES_POR_TRANSACAO = 4

const STATUS_AF_APROVACAO_AUTOMATICA = "APA"
const STATUS_AF_APROVACAO_MANUAL = "APM"
const STATUS_AF_REPROVADO_SEM_SUSPEITA = "RPM"
const STATUS_AF_ANALISE_MANUAL = "AMA"
const STATUS_AF_ERRO = "ERR"
const STATUS_AF_NOVO = "NVO"
const STATUS_AF_SUSPENSAO_MANUAL = "SUS"
const STATUS_AF_CANCELADO_PELO_CLIENTE = "CAN"
const STATUS_AF_FRAUDE_CONFIRMADA = "FRD"
const STATUS_AF_REPROVACAO_AUTOMATICA = "RPA"
const STATUS_AF_REPROVACAO_POR_POLITICA = "RPP"

class cl_BraspagCS_DadosCartao_Checkout
	dim bandeira
	dim valor_pagamento
	dim vl_pagamento
	dim opcao_parcelamento
	dim codigo_produto
	dim qtde_parcelas
	dim descricao_parcelamento
	dim titular_nome
	dim titular_cpf_cnpj
	dim cartao_numero
	dim cartao_validade_mes
	dim cartao_validade_ano
	dim cartao_codigo_seguranca
	dim cartao_proprio
	dim fatura_end_logradouro
	dim fatura_end_numero
	dim fatura_end_complemento
	dim fatura_end_bairro
	dim fatura_end_cidade
	dim fatura_end_uf
	dim fatura_end_cep
	dim fatura_tel_ddd
	dim fatura_tel_numero
	end class



' _________________
' gera_FingerPrint_SessionID
'
function gera_FingerPrint_SessionID
dim r
	set r = cn.Execute("SELECT convert(varchar(128), convert(varchar(36), NEWID()) + '-' + convert(varchar(36), NEWID()) + '-' + convert(varchar(36), NEWID()) + '-' + convert(varchar(36), NEWID())) AS SessionID")
	if Not r.Eof then gera_FingerPrint_SessionID = Trim("" & r("SessionID"))
	if r.State <> 0 then r.Close
	set r = nothing
end function




' ------------------------------------------------------------------------
'   BraspagCSProtegeNumeroCartao
'   Mascara os dígitos centrais do número do cartão.
function BraspagCSProtegeNumeroCartao(byval numero_cartao)
dim i, s_resp
	s_resp = ""
	numero_cartao = Trim("" & numero_cartao)
	for i = 1 to len(numero_cartao)
		if (i <= 6) Or (i > (len(numero_cartao)-4)) then
			s_resp = s_resp & mid(numero_cartao, i, 1)
		else
			s_resp = s_resp & "*"
			end if
		next
	BraspagCSProtegeNumeroCartao = s_resp
end function



' ------------------------------------------------------------------------
'   BraspagCSDescricaoParcelamento
'   Retorna a descrição para a forma de pagamento selecionada.
function BraspagCSDescricaoParcelamento(byval cod_produto, byval qtde_parcelas, byval valor_total)
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
			s_resp = SIMBOLO_MONETARIO & " " & formata_moeda(valor_total) & " em " & formata_inteiro(qtde_parcelas) & "x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " iguais"
	'	PARCELADO ADMINISTRADORA
		case "2"
			s_resp = SIMBOLO_MONETARIO & " " & formata_moeda(valor_total) & " em " & formata_inteiro(qtde_parcelas) & "x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " mais juros"
	'	DÉBITO
		case "A"
			s_resp = SIMBOLO_MONETARIO & " " & formata_moeda(valor_total) & " À Vista (no Débito)"
		case else
			s_resp = ""
	end select

	BraspagCSDescricaoParcelamento = s_resp
end function



' --------------------------------------------------------------------------------
' BraspagCS_monta_select_qtde_cartoes
function BraspagCS_monta_select_qtde_cartoes
dim i, strResp
	strResp = ""
	strResp = strResp & "<option selected value='1'>&nbsp;&nbsp;1&nbsp;</option>" & chr(13)
	for i = 2 to MAX_QTDE_CARTOES_POR_TRANSACAO
		strResp = strResp & "<option value='" & Cstr(i) & "'>&nbsp;&nbsp;" & Cstr(i) & "&nbsp;</option>" & chr(13)
		next
	BraspagCS_monta_select_qtde_cartoes = strResp
end function


' --------------------------------------------------------------------------------
'   BraspagCSGeraSufixoPedidoNsuPag
'   Gera um sufixo do tipo NSU para o pedido de forma a poder identificar na
'   Braspag de maneira inequívoca uma transação enviada através do nº do pedido.
function BraspagCSGeraSufixoPedidoNsuPag(Byval pedido, Byval usuario)
dim strSql, intNsu, lngRecordsAffected, s_log
dim t
	intNsu = 0
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PAGTO_GW_PAG_NSU" & _
			" WHERE" & _
				" (pedido = '" & pedido & "')"
	set t = cn.Execute(strSql)
	if t.Eof then
		intNsu = 0
		t.Close
		set t = nothing
		strSql = "INSERT INTO t_PAGTO_GW_PAG_NSU (" & _
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
	strSql = "UPDATE t_PAGTO_GW_PAG_NSU SET " & _
				"nsu = " & Cstr(intNsu) & "," & _
				"dt_hr_atualizacao = getdate()," & _
				"usuario_atualizacao = '" & usuario & "'" & _
			" WHERE" & _
				" (pedido = '" & pedido & "')"
	cn.Execute strSql, lngRecordsAffected
	
	s_log = "Gerado NSU=" & Cstr(intNsu) & " para o sufixo do pedido " & pedido & " na transação de pagamento da Braspag"
	grava_log usuario, "", pedido, "", OP_LOG_BRASPAG_PEDIDO_NSU_GERADO, s_log
	
	BraspagCSGeraSufixoPedidoNsuPag = intNsu
end function



function BraspagCS_monta_select_bandeiras(byval owner, byval bandeira_default)
dim v, i, bandeira, strResp, ha_default
	strResp = ""
	ha_default=False
	bandeira_default = Trim("" & bandeira_default)
	v = BraspagArrayBandeiras
	for i = Lbound(v) to Ubound(v)
		bandeira = Trim("" & v(i))
		if BraspagIsBandeiraHabilitada(owner, bandeira) then
			if (bandeira_default<>"") And (bandeira_default=bandeira) then
				strResp = strResp & "<option selected"
				ha_default=True
			else
				strResp = strResp & "<option"
				end if
			strResp = strResp & " value='" & bandeira & "'>"
			strResp = strResp & BraspagDescricaoBandeira(bandeira)
			strResp = strResp & "</option>" & chr(13)
			end if
		next

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if

	BraspagCS_monta_select_bandeiras = strResp
end function



function BraspagCSCalculaValorPagadorAutorizadoCapturado(Byval pedido, Byref msg_erro)
dim s, vl_pagador
dim r
	BraspagCSCalculaValorPagadorAutorizadoCapturado = 0
	vl_pagador = 0
	msg_erro = ""

	s = "SELECT" & _
			" Coalesce(Sum(valor_transacao),0) AS valor_total_transacao" & _
		" FROM t_PAGTO_GW_PAG t_PAG" & _
			" INNER JOIN t_PAGTO_GW_PAG_PAYMENT t_PAYMENT ON (t_PAG.id = t_PAYMENT.id_pagto_gw_pag)" & _
		" WHERE" & _
			" (pedido = '" & pedido & "')" & _
			" AND (ult_GlobalStatus IN ('" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA & "','" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "'))"
	set r = cn.Execute(s)
	if Not r.Eof then
		vl_pagador = r("valor_total_transacao")
		end if
	
	BraspagCSCalculaValorPagadorAutorizadoCapturado = vl_pagador
end function


function BraspagCSCalculaValorPagadorAutorizadoCapturadoFamilia(Byval pedido, Byref msg_erro)
dim s, vl_pagador
dim r
dim id_pedido_base
	BraspagCSCalculaValorPagadorAutorizadoCapturadoFamilia = 0
	vl_pagador = 0
	msg_erro = ""
	id_pedido_base = retorna_num_pedido_base(pedido)

	s = "SELECT" & _
			" Coalesce(Sum(valor_transacao),0) AS valor_total_transacao" & _
		" FROM t_PAGTO_GW_PAG t_PAG" & _
			" INNER JOIN t_PAGTO_GW_PAG_PAYMENT t_PAYMENT ON (t_PAG.id = t_PAYMENT.id_pagto_gw_pag)" & _
		" WHERE" & _
			" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')" & _
			" AND (ult_GlobalStatus IN ('" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA & "','" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "'))"
	set r = cn.Execute(s)
	if Not r.Eof then
		vl_pagador = r("valor_total_transacao")
		end if
	
	BraspagCSCalculaValorPagadorAutorizadoCapturadoFamilia = vl_pagador
end function


function isAFStatusAprovado(byval status)
dim blnResp
	isAFStatusAprovado = False
	status = Trim("" & status)
	if status = "" then exit function

	blnResp = False
	if status = STATUS_AF_APROVACAO_AUTOMATICA then blnResp = True
	if status = STATUS_AF_APROVACAO_MANUAL then blnResp = True
	isAFStatusAprovado=blnResp
end function


function isAFStatusReprovado(byval status)
dim blnResp
	isAFStatusReprovado = False
	status = Trim("" & status)
	if status = "" then exit function

	blnResp = False
	if status = STATUS_AF_REPROVADO_SEM_SUSPEITA then blnResp = True
	if status = STATUS_AF_FRAUDE_CONFIRMADA then blnResp = True
	if status = STATUS_AF_REPROVACAO_AUTOMATICA then blnResp = True
	if status = STATUS_AF_REPROVACAO_POR_POLITICA then blnResp = True
	if status = STATUS_AF_CANCELADO_PELO_CLIENTE then blnResp = True
	isAFStatusReprovado = blnResp
end function


function ClearsaleDescricaoAFStatus(byval status)
dim s
	status = Trim("" & status)
	select case status
		case STATUS_AF_APROVACAO_AUTOMATICA: s="Aprovação Automática"
		case STATUS_AF_APROVACAO_MANUAL: s="Aprovação Manual"
		case STATUS_AF_REPROVADO_SEM_SUSPEITA: s="Reprovado Sem Suspeita"
		case STATUS_AF_ANALISE_MANUAL: s="Análise Manual"
		case STATUS_AF_ERRO: s="Erro"
		case STATUS_AF_NOVO: s="Novo"
		case STATUS_AF_SUSPENSAO_MANUAL: s="Suspensão Manual"
		case STATUS_AF_CANCELADO_PELO_CLIENTE: s="Cancelado pelo Cliente"
		case STATUS_AF_FRAUDE_CONFIRMADA: s="Fraude Confirmada"
		case STATUS_AF_REPROVACAO_AUTOMATICA: s="Reprovação Automática"
		case STATUS_AF_REPROVACAO_POR_POLITICA: s="Reprovação por Política"
		case else s=""
		end select
	ClearsaleDescricaoAFStatus=s
end function



' --------------------------------------------------------------------------------
'   BraspagClearsaleVerificaPreRequisito_BraspagTransactionId
'   Verifica se há a informação de 'BraspagTransactionId'. Caso não,
'   executa a consulta 'GetOrderIdData' usando o campo 'OrderId' p/
'   tentar obter o 'BraspagTransactionId', que é necessário p/ a maioria
'   das requisicoes.
function BraspagClearsaleVerificaPreRequisito_BraspagTransactionId(byval id_pagto_gw_pag, byval id_pagto_gw_pag_payment, byval usuario, byref msg_erro)
dim t, tPagtoGwPag, tPagtoGwPagPayment, tPagtoGwPagOpComplementar, tPagtoGwPagOpComplementarXml
dim i, lngRecordsAffected, intQtdeRespostas
dim id_pagto_gw_pag_op_complementar, id_pagto_gw_pag_op_compl_xml_tx, id_pagto_gw_pag_op_compl_xml_rx
dim strMerchantId, strBraspagTransactionId, strOrderId
dim strSql
dim txXml, rxXml
dim r_rx, v_rx_item()

	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(tPagtoGwPag, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagPayment, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PAGTO_GW_PAG" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag & ")"
	if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
	tPagtoGwPag.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if tPagtoGwPag.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PAGTO_GW_PAG_PAYMENT" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag_payment & ")"
	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	tPagtoGwPagPayment.open strSql, cn
	if tPagtoGwPagPayment.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag!!"
		exit function
		end if
	
	strMerchantId = Trim("" & tPagtoGwPag("req_OrderData_MerchantId"))
	strOrderId = Trim("" & tPagtoGwPag("req_OrderData_OrderId"))
	strBraspagTransactionId = Trim("" & tPagtoGwPagPayment("resp_PaymentDataResponse_BraspagTransactionId"))
	
'	A INFORMAÇÃO 'BraspagTransactionId' ESTÁ DISPONÍVEL?
'	SE O CAMPO 'req_OrderData_OrderId' ESTIVER VAZIO NÃO SERÁ POSSÍVEL REALIZAR A CONSULTA 'GetOrderIdData'
	if (strBraspagTransactionId <> "") Or (strOrderId = "") then
	'	FECHA TABELAS
		if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
		set tPagtoGwPag = nothing
		
		if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
		set tPagtoGwPagPayment = nothing
		
		exit function
		end if
	
	strSql = "SELECT" & _
				" Count(*) AS qtde" & _
			" FROM t_PAGTO_GW_PAG" & _
				" INNER JOIN t_PAGTO_GW_PAG_PAYMENT ON (t_PAGTO_GW_PAG.id = t_PAGTO_GW_PAG_PAYMENT.id_pagto_gw_pag)" & _
			" WHERE" & _
				" (req_OrderData_OrderId = '" & strOrderId & "')"
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
			
			if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
			set tPagtoGwPag = nothing
			
			if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
			set tPagtoGwPagPayment = nothing
			
			exit function
			end if
		end if
	
	t.Close
	set t = nothing
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(tPagtoGwPagOpComplementar, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagOpComplementarXml, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_GetOrderIdData_TX(strMerchantId, strOrderId)
	txXml = BraspagXmlMontaRequisicaoGetOrderIdData(trx)
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR, id_pagto_gw_pag_op_complementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_complementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_complementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, id_pagto_gw_pag_op_compl_xml_tx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_compl_xml_tx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_compl_xml_tx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, id_pagto_gw_pag_op_compl_xml_rx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_compl_xml_rx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_compl_xml_rx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível consultar a Braspag devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR WHERE (id = -1)"
	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	tPagtoGwPagOpComplementar.Open strSql, cn
	tPagtoGwPagOpComplementar.AddNew
	tPagtoGwPagOpComplementar("id") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementar("id_pagto_gw_pag") = CLng(id_pagto_gw_pag)
	tPagtoGwPagOpComplementar("id_pagto_gw_pag_payment") = CLng(id_pagto_gw_pag_payment)
	tPagtoGwPagOpComplementar("operacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_ORDERID_DATA
	tPagtoGwPagOpComplementar("usuario") = usuario
	tPagtoGwPagOpComplementar("trx_TX_data") = Date
	tPagtoGwPagOpComplementar("trx_TX_data_hora") = Now
	tPagtoGwPagOpComplementar("req_RequestId") = trx.PAG_RequestId
	tPagtoGwPagOpComplementar("req_Version") = trx.PAG_Version
	tPagtoGwPagOpComplementar("req_MerchantId") = trx.PAG_MerchantId
	tPagtoGwPagOpComplementar("req_OrderId") = trx.PAG_OrderId
	tPagtoGwPagOpComplementar.Update
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	tPagtoGwPagOpComplementarXml.Open strSql, cn
	tPagtoGwPagOpComplementarXml.AddNew
	tPagtoGwPagOpComplementarXml("id") = id_pagto_gw_pag_op_compl_xml_tx
	tPagtoGwPagOpComplementarXml("id_pagto_gw_pag_op_complementar") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementarXml("data") = Date
	tPagtoGwPagOpComplementarXml("data_hora") = Now
	tPagtoGwPagOpComplementarXml("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_ORDERID_DATA
	tPagtoGwPagOpComplementarXml("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	tPagtoGwPagOpComplementarXml("xml") = txXml
	tPagtoGwPagOpComplementarXml.Update
	
	rxXml = BraspagEnviaTransacaoComRetry(txXml, BRASPAG_WS_ENDERECO_PAGADOR_QUERY)
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR WHERE (id = " & id_pagto_gw_pag_op_complementar & ")"
	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	tPagtoGwPagOpComplementar.Open strSql, cn
	if Not tPagtoGwPagOpComplementar.Eof then
		tPagtoGwPagOpComplementar("trx_RX_data") = Date
		tPagtoGwPagOpComplementar("trx_RX_data_hora") = Now
		tPagtoGwPagOpComplementar("trx_RX_status") = 1
		if Trim(rxXml) = "" then tPagtoGwPagOpComplementar("trx_RX_vazio_status") = 1
		tPagtoGwPagOpComplementar.Update
		end if
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	tPagtoGwPagOpComplementarXml.Open strSql, cn
	tPagtoGwPagOpComplementarXml.AddNew
	tPagtoGwPagOpComplementarXml("id") = id_pagto_gw_pag_op_compl_xml_rx
	tPagtoGwPagOpComplementarXml("id_pagto_gw_pag_op_complementar") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementarXml("data") = Date
	tPagtoGwPagOpComplementarXml("data_hora") = Now
	tPagtoGwPagOpComplementarXml("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_ORDERID_DATA
	tPagtoGwPagOpComplementarXml("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	tPagtoGwPagOpComplementarXml("xml") = rxXml
	tPagtoGwPagOpComplementarXml.Update
	
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
		strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" & _
					" resp_PaymentDataResponse_BraspagTransactionId = '" & strBraspagTransactionId & "'" & _
				" WHERE" & _
					" (id = " & id_pagto_gw_pag_payment & ")"
		cn.Execute strSql, lngRecordsAffected
		
		strSql = "UPDATE t_PAGTO_GW_PAG_OP_COMPLEMENTAR SET" & _
					" st_sucesso = 1," & _
					" resp_BraspagTransactionId = '" & strBraspagTransactionId & "'" & _
				" WHERE" & _
					" (id = " & id_pagto_gw_pag_op_complementar & ")"
		cn.Execute strSql, lngRecordsAffected
		end if
	
'	FECHA TABELAS
	if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
	set tPagtoGwPag = nothing

	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	set tPagtoGwPagPayment = nothing

	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	set tPagtoGwPagOpComplementar = nothing

	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	set tPagtoGwPagOpComplementarXml = nothing
end function



' --------------------------------------------------------------------------------
'   BraspagClearsaleProcessaConsulta_GetTransactionData
'   Executa a consulta e realiza o processamento relacionado ao BD.
function BraspagClearsaleProcessaConsulta_GetTransactionData(byval id_pagto_gw_pag, byval id_pagto_gw_pag_payment, byval usuario, byref trx, byref r_rx, byref msg_erro)
dim tPagtoGwPag, tPagtoGwPagPayment, tPagtoGwPagOpComplementar, tPagtoGwPagOpComplementarXml
dim lngRecordsAffected
dim id_pagto_gw_pag_op_complementar, id_pagto_gw_pag_op_compl_xml_tx, id_pagto_gw_pag_op_compl_xml_rx
dim strCapturedDate, strVoidedDate
dim strMerchantId, strBraspagTransactionId
dim strSql
dim txXml, rxXml
	
	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(tPagtoGwPag, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagPayment, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagOpComplementar, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagOpComplementarXml, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PAGTO_GW_PAG" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag & ")"
	if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
	tPagtoGwPag.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if tPagtoGwPag.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PAGTO_GW_PAG_PAYMENT" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag_payment & ")"
	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	tPagtoGwPagPayment.open strSql, cn
	if tPagtoGwPagPayment.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag!!"
		exit function
		end if
	
	strMerchantId = Trim("" & tPagtoGwPag("req_OrderData_MerchantId"))
	strBraspagTransactionId = Trim("" & tPagtoGwPagPayment("resp_PaymentDataResponse_BraspagTransactionId"))
	
	if strBraspagTransactionId = "" then
		msg_erro = "Não é possível consultar a Braspag porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!"
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_GetTransactionData_TX(strMerchantId, strBraspagTransactionId)
	txXml = BraspagXmlMontaRequisicaoGetTransactionData(trx)
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR, id_pagto_gw_pag_op_complementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_complementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_complementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, id_pagto_gw_pag_op_compl_xml_tx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_compl_xml_tx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_compl_xml_tx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, id_pagto_gw_pag_op_compl_xml_rx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_compl_xml_rx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_compl_xml_rx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível consultar a Braspag devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR WHERE (id = -1)"
	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	tPagtoGwPagOpComplementar.Open strSql, cn
	tPagtoGwPagOpComplementar.AddNew
	tPagtoGwPagOpComplementar("id") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementar("id_pagto_gw_pag") = CLng(id_pagto_gw_pag)
	tPagtoGwPagOpComplementar("id_pagto_gw_pag_payment") = CLng(id_pagto_gw_pag_payment)
	tPagtoGwPagOpComplementar("operacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_TRANSACTION_DATA
	tPagtoGwPagOpComplementar("usuario") = usuario
	tPagtoGwPagOpComplementar("trx_TX_data") = Date
	tPagtoGwPagOpComplementar("trx_TX_data_hora") = Now
	tPagtoGwPagOpComplementar("req_RequestId") = trx.PAG_RequestId
	tPagtoGwPagOpComplementar("req_Version") = trx.PAG_Version
	tPagtoGwPagOpComplementar("req_MerchantId") = trx.PAG_MerchantId
	tPagtoGwPagOpComplementar("req_BraspagTransactionId") = trx.PAG_BraspagTransactionId
	tPagtoGwPagOpComplementar.Update
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	tPagtoGwPagOpComplementarXml.Open strSql, cn
	tPagtoGwPagOpComplementarXml.AddNew
	tPagtoGwPagOpComplementarXml("id") = id_pagto_gw_pag_op_compl_xml_tx
	tPagtoGwPagOpComplementarXml("id_pagto_gw_pag_op_complementar") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementarXml("data") = Date
	tPagtoGwPagOpComplementarXml("data_hora") = Now
	tPagtoGwPagOpComplementarXml("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_TRANSACTION_DATA
	tPagtoGwPagOpComplementarXml("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	tPagtoGwPagOpComplementarXml("xml") = txXml
	tPagtoGwPagOpComplementarXml.Update
	
	rxXml = BraspagEnviaTransacaoComRetry(txXml, BRASPAG_WS_ENDERECO_PAGADOR_QUERY)
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR WHERE (id = " & id_pagto_gw_pag_op_complementar & ")"
	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	tPagtoGwPagOpComplementar.Open strSql, cn
	if Not tPagtoGwPagOpComplementar.Eof then
		tPagtoGwPagOpComplementar("trx_RX_data") = Date
		tPagtoGwPagOpComplementar("trx_RX_data_hora") = Now
		tPagtoGwPagOpComplementar("trx_RX_status") = 1
		if Trim(rxXml) = "" then tPagtoGwPagOpComplementar("trx_RX_vazio_status") = 1
		tPagtoGwPagOpComplementar.Update
		end if
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	tPagtoGwPagOpComplementarXml.Open strSql, cn
	tPagtoGwPagOpComplementarXml.AddNew
	tPagtoGwPagOpComplementarXml("id") = id_pagto_gw_pag_op_compl_xml_rx
	tPagtoGwPagOpComplementarXml("id_pagto_gw_pag_op_complementar") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementarXml("data") = Date
	tPagtoGwPagOpComplementarXml("data_hora") = Now
	tPagtoGwPagOpComplementarXml("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_GET_TRANSACTION_DATA
	tPagtoGwPagOpComplementarXml("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	tPagtoGwPagOpComplementarXml("xml") = rxXml
	tPagtoGwPagOpComplementarXml.Update
	
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
	
	strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" & _
				" ult_GlobalStatus = '" & decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(r_rx.PAG_Status) & "'," & _
				" ult_atualizacao_data_hora = getdate()," & _
				" ult_atualizacao_usuario = '" & usuario & "'," & _
				" ult_id_pagto_gw_pag_payment_op_complementar = " & id_pagto_gw_pag_op_complementar & "," & _
				" resp_CapturedDate = " & strCapturedDate & "," & _
				" resp_VoidedDate = " & strVoidedDate & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag_payment & ")"
	cn.Execute strSql, lngRecordsAffected
	
	strSql = "UPDATE t_PAGTO_GW_PAG_OP_COMPLEMENTAR SET" & _
				" resp_AuthorizationCode = '" & r_rx.PAG_AuthorizationCode & "'," & _
				" resp_ProofOfSale = '" & r_rx.PAG_ProofOfSale & "'," & _
				" st_sucesso = 1, " & _
				" resp_Status = '" & r_rx.PAG_Status & "'" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag_op_complementar & ")"
	cn.Execute strSql, lngRecordsAffected
	
'	FECHA TABELAS
	if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
	set tPagtoGwPag = nothing

	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	set tPagtoGwPagPayment = nothing

	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	set tPagtoGwPagOpComplementar = nothing

	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	set tPagtoGwPagOpComplementarXml = nothing
end function



' --------------------------------------------------------------------------------
'   BraspagClearsaleProcessaRequisicao_CaptureCreditCardTransaction
'   Executa a requisição e realiza o processamento relacionado ao BD.
function BraspagClearsaleProcessaRequisicao_CaptureCreditCardTransaction(byval id_pagto_gw_pag, byval id_pagto_gw_pag_payment, byval usuario, byref trx, byref r_rx, byref msg_erro)
dim tPagtoGwPag, tPagtoGwPagPayment, tPagtoGwPagOpComplementar, tPagtoGwPagOpComplementarXml
dim pedido, vl_transacao
dim lngRecordsAffected
dim id_pagto_gw_pag_op_complementar, id_pagto_gw_pag_op_compl_xml_tx, id_pagto_gw_pag_op_compl_xml_rx
dim strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount
dim strCapturedDate
dim strSql
dim txXml, rxXml
dim st_sucesso
	
	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(tPagtoGwPag, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagPayment, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagOpComplementar, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagOpComplementarXml, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PAGTO_GW_PAG" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag & ")"
	if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
	tPagtoGwPag.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if tPagtoGwPag.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PAGTO_GW_PAG_PAYMENT" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag_payment & ")"
	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	tPagtoGwPagPayment.open strSql, cn
	if tPagtoGwPagPayment.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag!!"
		exit function
		end if
	
	pedido = Trim("" & tPagtoGwPag("pedido"))
	vl_transacao = tPagtoGwPagPayment("valor_transacao")
	strMerchantId = Trim("" & tPagtoGwPag("req_OrderData_MerchantId"))
	strBraspagTransactionId = Trim("" & tPagtoGwPagPayment("resp_PaymentDataResponse_BraspagTransactionId"))
	strAmount = Trim("" & tPagtoGwPagPayment("req_PaymentDataRequest_Amount"))
	strServiceTaxAmount = Trim("" & tPagtoGwPagPayment("req_PaymentDataRequest_ServiceTaxAmount"))
	
	if strBraspagTransactionId = "" then
		msg_erro = "Não é possível consultar a Braspag porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!"
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_CaptureCreditCardTransaction_TX(strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount)
	txXml = BraspagXmlMontaRequisicaoCaptureCreditCardTransaction(trx)
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR, id_pagto_gw_pag_op_complementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_complementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_complementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, id_pagto_gw_pag_op_compl_xml_tx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_compl_xml_tx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_compl_xml_tx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, id_pagto_gw_pag_op_compl_xml_rx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_compl_xml_rx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_compl_xml_rx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível enviar a solicitação à Braspag devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR WHERE (id = -1)"
	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	tPagtoGwPagOpComplementar.Open strSql, cn
	tPagtoGwPagOpComplementar.AddNew
	tPagtoGwPagOpComplementar("id") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementar("id_pagto_gw_pag") = CLng(id_pagto_gw_pag)
	tPagtoGwPagOpComplementar("id_pagto_gw_pag_payment") = CLng(id_pagto_gw_pag_payment)
	tPagtoGwPagOpComplementar("operacao") = BRASPAG_TIPO_TRANSACAO__PAG_CAPTURECREDITCARDTRANSACTION
	tPagtoGwPagOpComplementar("usuario") = usuario
	tPagtoGwPagOpComplementar("trx_TX_data") = Date
	tPagtoGwPagOpComplementar("trx_TX_data_hora") = Now
	tPagtoGwPagOpComplementar("req_RequestId") = trx.PAG_RequestId
	tPagtoGwPagOpComplementar("req_Version") = trx.PAG_Version
	tPagtoGwPagOpComplementar("req_MerchantId") = trx.PAG_MerchantId
	tPagtoGwPagOpComplementar("req_BraspagTransactionId") = trx.PAG_BraspagTransactionId
	tPagtoGwPagOpComplementar("req_Amount") = trx.PAG_Amount
	tPagtoGwPagOpComplementar("req_ServiceTaxAmount") = trx.PAG_ServiceTaxAmount
	tPagtoGwPagOpComplementar.Update
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	tPagtoGwPagOpComplementarXml.Open strSql, cn
	tPagtoGwPagOpComplementarXml.AddNew
	tPagtoGwPagOpComplementarXml("id") = id_pagto_gw_pag_op_compl_xml_tx
	tPagtoGwPagOpComplementarXml("id_pagto_gw_pag_op_complementar") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementarXml("data") = Date
	tPagtoGwPagOpComplementarXml("data_hora") = Now
	tPagtoGwPagOpComplementarXml("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_CAPTURECREDITCARDTRANSACTION
	tPagtoGwPagOpComplementarXml("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	tPagtoGwPagOpComplementarXml("xml") = txXml
	tPagtoGwPagOpComplementarXml.Update
	
	rxXml = BraspagEnviaTransacaoComRetry(txXml, BRASPAG_WS_ENDERECO_PAGADOR_TRANSACTION)
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR WHERE (id = " & id_pagto_gw_pag_op_complementar & ")"
	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	tPagtoGwPagOpComplementar.Open strSql, cn
	if Not tPagtoGwPagOpComplementar.Eof then
		tPagtoGwPagOpComplementar("trx_RX_data") = Date
		tPagtoGwPagOpComplementar("trx_RX_data_hora") = Now
		tPagtoGwPagOpComplementar("trx_RX_status") = 1
		if Trim(rxXml) = "" then tPagtoGwPagOpComplementar("trx_RX_vazio_status") = 1
		tPagtoGwPagOpComplementar.Update
		end if
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	tPagtoGwPagOpComplementarXml.Open strSql, cn
	tPagtoGwPagOpComplementarXml.AddNew
	tPagtoGwPagOpComplementarXml("id") = id_pagto_gw_pag_op_compl_xml_rx
	tPagtoGwPagOpComplementarXml("id_pagto_gw_pag_op_complementar") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementarXml("data") = Date
	tPagtoGwPagOpComplementarXml("data_hora") = Now
	tPagtoGwPagOpComplementarXml("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_CAPTURECREDITCARDTRANSACTION
	tPagtoGwPagOpComplementarXml("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	tPagtoGwPagOpComplementarXml("xml") = rxXml
	tPagtoGwPagOpComplementarXml.Update
	
	set r_rx = BraspagCarregaDados_CaptureCreditCardTransactionResponse(rxXml, msg_erro)
	if msg_erro <> "" then exit function
	
'	ATUALIZA O ÚLTIMO STATUS DA TRANSAÇÃO
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_CAPTURECREDITCARDTRANSACTIONRESPONSE_STATUS__CAPTURE_CONFIRMED then
		strCapturedDate = bd_monta_data_hora(Now)
		strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" & _
					" captura_confirmada_status = 1," & _
					" captura_confirmada_data = Convert(datetime, Convert(varchar(10),getdate(), 121), 121)," & _
					" captura_confirmada_data_hora = getdate()," & _
					" captura_confirmada_usuario = '" & usuario & "'," & _
					" ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "'," & _
					" resp_CapturedDate = " & strCapturedDate & "," & _
					" ult_atualizacao_data_hora = getdate()," & _
					" ult_atualizacao_usuario = '" & usuario & "'," & _
					" ult_id_pagto_gw_pag_payment_op_complementar = " & id_pagto_gw_pag_op_complementar & _
				" WHERE" & _
					" (id = " & id_pagto_gw_pag_payment & ")"
	else
		strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" & _
					" ult_id_pagto_gw_pag_payment_op_complementar = " & id_pagto_gw_pag_op_complementar & _
				" WHERE" & _
					" (id = " & id_pagto_gw_pag_payment & ")"
		end if
	cn.Execute strSql, lngRecordsAffected
	
	st_sucesso = 0
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_CAPTURECREDITCARDTRANSACTIONRESPONSE_STATUS__CAPTURE_CONFIRMED then st_sucesso = 1
	strSql = "UPDATE t_PAGTO_GW_PAG_OP_COMPLEMENTAR SET" & _
				" resp_AuthorizationCode = '" & r_rx.PAG_AuthorizationCode & "'," & _
				" resp_ProofOfSale = '" & r_rx.PAG_ProofOfSale & "'," & _
				" st_sucesso = " & CStr(st_sucesso) & "," & _
				" resp_Status = '" & r_rx.PAG_Status & "'" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag_op_complementar & ")"
	cn.Execute strSql, lngRecordsAffected
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
	set tPagtoGwPag = nothing

	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	set tPagtoGwPagPayment = nothing

	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	set tPagtoGwPagOpComplementar = nothing

	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	set tPagtoGwPagOpComplementarXml = nothing
	
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_CAPTURECREDITCARDTRANSACTIONRESPONSE_STATUS__CAPTURE_CONFIRMED then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if BraspagClearsaleRegistraPagtoNoPedido(BRASPAG_REGISTRA_PAGTO__OP_CAPTURA, pedido, id_pagto_gw_pag, id_pagto_gw_pag_payment, converte_numero(vl_transacao), usuario, msg_erro) then
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
'   BraspagClearsaleProcessaRequisicao_VoidCreditCardTransaction
'   Executa a requisição e realiza o processamento relacionado ao BD.
'   Transações autorizadas são canceladas através do método Void.
'   Transações capturadas são canceladas através do método Void até a meia-noite do mesmo dia, após isso deve-se usar o método Refund.
function BraspagClearsaleProcessaRequisicao_VoidCreditCardTransaction(byval id_pagto_gw_pag, byval id_pagto_gw_pag_payment, byval usuario, byref trx, byref r_rx, byref msg_erro)
dim tPagtoGwPag, tPagtoGwPagPayment, tPagtoGwPagOpComplementar, tPagtoGwPagOpComplementarXml
dim pedido, vl_transacao
dim lngRecordsAffected
dim id_pagto_gw_pag_op_complementar, id_pagto_gw_pag_op_compl_xml_tx, id_pagto_gw_pag_op_compl_xml_rx
dim strVoidedDate
dim strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount
dim strSql
dim txXml, rxXml
dim st_sucesso
	
	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(tPagtoGwPag, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagPayment, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagOpComplementar, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagOpComplementarXml, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PAGTO_GW_PAG" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag & ")"
	if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
	tPagtoGwPag.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if tPagtoGwPag.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PAGTO_GW_PAG_PAYMENT" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag_payment & ")"
	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	tPagtoGwPagPayment.open strSql, cn
	if tPagtoGwPagPayment.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag!!"
		exit function
		end if
	
	pedido = Trim("" & tPagtoGwPag("pedido"))
	vl_transacao = tPagtoGwPagPayment("valor_transacao")
	strMerchantId = Trim("" & tPagtoGwPag("req_OrderData_MerchantId"))
	strBraspagTransactionId = Trim("" & tPagtoGwPagPayment("resp_PaymentDataResponse_BraspagTransactionId"))
	strAmount = Trim("" & tPagtoGwPagPayment("req_PaymentDataRequest_Amount"))
	strServiceTaxAmount = Trim("" & tPagtoGwPagPayment("req_PaymentDataRequest_ServiceTaxAmount"))
	
	if strBraspagTransactionId = "" then
		msg_erro = "Não é possível consultar a Braspag porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!"
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_VoidCreditCardTransaction_TX(strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount)
	txXml = BraspagXmlMontaRequisicaoVoidCreditCardTransaction(trx)
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR, id_pagto_gw_pag_op_complementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_complementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_complementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, id_pagto_gw_pag_op_compl_xml_tx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_compl_xml_tx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_compl_xml_tx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, id_pagto_gw_pag_op_compl_xml_rx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_compl_xml_rx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_compl_xml_rx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível enviar a solicitação à Braspag devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR WHERE (id = -1)"
	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	tPagtoGwPagOpComplementar.Open strSql, cn
	tPagtoGwPagOpComplementar.AddNew
	tPagtoGwPagOpComplementar("id") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementar("id_pagto_gw_pag") = CLng(id_pagto_gw_pag)
	tPagtoGwPagOpComplementar("id_pagto_gw_pag_payment") = CLng(id_pagto_gw_pag_payment)
	tPagtoGwPagOpComplementar("operacao") = BRASPAG_TIPO_TRANSACAO__PAG_VOIDCREDITCARDTRANSACTION
	tPagtoGwPagOpComplementar("usuario") = usuario
	tPagtoGwPagOpComplementar("trx_TX_data") = Date
	tPagtoGwPagOpComplementar("trx_TX_data_hora") = Now
	tPagtoGwPagOpComplementar("req_RequestId") = trx.PAG_RequestId
	tPagtoGwPagOpComplementar("req_Version") = trx.PAG_Version
	tPagtoGwPagOpComplementar("req_MerchantId") = trx.PAG_MerchantId
	tPagtoGwPagOpComplementar("req_BraspagTransactionId") = trx.PAG_BraspagTransactionId
	tPagtoGwPagOpComplementar("req_Amount") = trx.PAG_Amount
	tPagtoGwPagOpComplementar("req_ServiceTaxAmount") = trx.PAG_ServiceTaxAmount
	tPagtoGwPagOpComplementar.Update
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	tPagtoGwPagOpComplementarXml.Open strSql, cn
	tPagtoGwPagOpComplementarXml.AddNew
	tPagtoGwPagOpComplementarXml("id") = id_pagto_gw_pag_op_compl_xml_tx
	tPagtoGwPagOpComplementarXml("id_pagto_gw_pag_op_complementar") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementarXml("data") = Date
	tPagtoGwPagOpComplementarXml("data_hora") = Now
	tPagtoGwPagOpComplementarXml("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_VOIDCREDITCARDTRANSACTION
	tPagtoGwPagOpComplementarXml("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	tPagtoGwPagOpComplementarXml("xml") = txXml
	tPagtoGwPagOpComplementarXml.Update
	
	rxXml = BraspagEnviaTransacaoComRetry(txXml, BRASPAG_WS_ENDERECO_PAGADOR_TRANSACTION)
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR WHERE (id = " & id_pagto_gw_pag_op_complementar & ")"
	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	tPagtoGwPagOpComplementar.Open strSql, cn
	if Not tPagtoGwPagOpComplementar.Eof then
		tPagtoGwPagOpComplementar("trx_RX_data") = Date
		tPagtoGwPagOpComplementar("trx_RX_data_hora") = Now
		tPagtoGwPagOpComplementar("trx_RX_status") = 1
		if Trim(rxXml) = "" then tPagtoGwPagOpComplementar("trx_RX_vazio_status") = 1
		tPagtoGwPagOpComplementar.Update
		end if
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	tPagtoGwPagOpComplementarXml.Open strSql, cn
	tPagtoGwPagOpComplementarXml.AddNew
	tPagtoGwPagOpComplementarXml("id") = id_pagto_gw_pag_op_compl_xml_rx
	tPagtoGwPagOpComplementarXml("id_pagto_gw_pag_op_complementar") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementarXml("data") = Date
	tPagtoGwPagOpComplementarXml("data_hora") = Now
	tPagtoGwPagOpComplementarXml("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_VOIDCREDITCARDTRANSACTION
	tPagtoGwPagOpComplementarXml("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	tPagtoGwPagOpComplementarXml("xml") = rxXml
	tPagtoGwPagOpComplementarXml.Update
	
	set r_rx = BraspagCarregaDados_VoidCreditCardTransactionResponse(rxXml, msg_erro)
	if msg_erro <> "" then exit function
	
'	ATUALIZA O ÚLTIMO STATUS DA TRANSAÇÃO
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_VOIDCREDITCARDTRANSACTIONRESPONSE_STATUS__VOID_CONFIRMED then
		strVoidedDate = bd_monta_data_hora(Now)
		strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" & _
					" voided_status = 1," & _
					" voided_data = Convert(datetime, Convert(varchar(10),getdate(), 121), 121)," & _
					" voided_data_hora = getdate()," & _
					" voided_usuario = '" & usuario & "'," & _
					" ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURA_CANCELADA & "'," & _
					" resp_VoidedDate = " & strVoidedDate & "," & _
					" ult_atualizacao_data_hora = getdate()," & _
					" ult_atualizacao_usuario = '" & usuario & "'," & _
					" ult_id_pagto_gw_pag_payment_op_complementar = " & id_pagto_gw_pag_op_complementar & _
				" WHERE" & _
					" (id = " & id_pagto_gw_pag_payment & ")"
	else
		strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" & _
					" ult_id_pagto_gw_pag_payment_op_complementar = " & id_pagto_gw_pag_op_complementar & _
				" WHERE" & _
					" (id = " & id_pagto_gw_pag_payment & ")"
		end if
	cn.Execute strSql, lngRecordsAffected
	
	st_sucesso = 0
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_VOIDCREDITCARDTRANSACTIONRESPONSE_STATUS__VOID_CONFIRMED then st_sucesso = 1
	strSql = "UPDATE t_PAGTO_GW_PAG_OP_COMPLEMENTAR SET" & _
				" resp_AuthorizationCode = '" & r_rx.PAG_AuthorizationCode & "'," & _
				" resp_ProofOfSale = '" & r_rx.PAG_ProofOfSale & "'," & _
				" st_sucesso = " & CStr(st_sucesso) & "," & _
				" resp_Status = '" & r_rx.PAG_Status & "'" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag_op_complementar & ")"
	cn.Execute strSql, lngRecordsAffected
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
	set tPagtoGwPag = nothing

	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	set tPagtoGwPagPayment = nothing

	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	set tPagtoGwPagOpComplementar = nothing

	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	set tPagtoGwPagOpComplementarXml = nothing
	
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_VOIDCREDITCARDTRANSACTIONRESPONSE_STATUS__VOID_CONFIRMED then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if BraspagClearsaleRegistraPagtoNoPedido(BRASPAG_REGISTRA_PAGTO__OP_CANCELAMENTO, pedido, id_pagto_gw_pag, id_pagto_gw_pag_payment, converte_numero(vl_transacao), usuario, msg_erro) then
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
'   BraspagClearsaleProcessaRequisicao_RefundCreditCardTransaction
'   Executa a requisição e realiza o processamento relacionado ao BD.
'   Transações autorizadas são canceladas através do método Void.
'   Transações capturadas são canceladas através do método Void até a meia-noite do mesmo dia, após isso deve-se usar o método Refund.
function BraspagClearsaleProcessaRequisicao_RefundCreditCardTransaction(byval id_pagto_gw_pag, byval id_pagto_gw_pag_payment, byval usuario, byref trx, byref r_rx, byref msg_erro)
dim tPagtoGwPag, tPagtoGwPagPayment, tPagtoGwPagOpComplementar, tPagtoGwPagOpComplementarXml
dim pedido, vl_transacao
dim lngRecordsAffected
dim id_pagto_gw_pag_op_complementar, id_pagto_gw_pag_op_compl_xml_tx, id_pagto_gw_pag_op_compl_xml_rx
dim strVoidedDate
dim strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount
dim strSql
dim txXml, rxXml
dim st_sucesso
	
	msg_erro = ""
	
'	TABELAS DO BD
	If Not cria_recordset_otimista(tPagtoGwPag, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagPayment, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagOpComplementar, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	If Not cria_recordset_otimista(tPagtoGwPagOpComplementarXml, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO) & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PAGTO_GW_PAG" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag & ")"
	if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
	tPagtoGwPag.open strSql, cn
'	NÃO ENCONTROU O REGISTRO?
	if tPagtoGwPag.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com a Braspag!!"
		exit function
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_PAGTO_GW_PAG_PAYMENT" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag_payment & ")"
	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	tPagtoGwPagPayment.open strSql, cn
	if tPagtoGwPagPayment.Eof then
		msg_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag!!"
		exit function
		end if
	
	pedido = Trim("" & tPagtoGwPag("pedido"))
	vl_transacao = tPagtoGwPagPayment("valor_transacao")
	strMerchantId = Trim("" & tPagtoGwPag("req_OrderData_MerchantId"))
	strBraspagTransactionId = Trim("" & tPagtoGwPagPayment("resp_PaymentDataResponse_BraspagTransactionId"))
	strAmount = Trim("" & tPagtoGwPagPayment("req_PaymentDataRequest_Amount"))
	strServiceTaxAmount = Trim("" & tPagtoGwPagPayment("req_PaymentDataRequest_ServiceTaxAmount"))
	
	if strBraspagTransactionId = "" then
		msg_erro = "Não é possível consultar a Braspag porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!"
		exit function
		end if
	
	set trx = cria_instancia_cl_BRASPAG_RefundCreditCardTransaction_TX(strMerchantId, strBraspagTransactionId, strAmount, strServiceTaxAmount)
	txXml = BraspagXmlMontaRequisicaoRefundCreditCardTransaction(trx)
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR, id_pagto_gw_pag_op_complementar, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DA OPERAÇÃO COMPLEMENTAR (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_complementar <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_complementar & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, id_pagto_gw_pag_op_compl_xml_tx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_compl_xml_tx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_compl_xml_tx & ")"
		exit function
		end if
	
	if Not fin_gera_nsu(T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, id_pagto_gw_pag_op_compl_xml_rx, msg_erro) then
		msg_erro = "FALHA AO GERAR NSU PARA O REGISTRO DO XML DE RETORNO DA TRANSAÇÃO (" & msg_erro & ")"
		exit function
	elseif id_pagto_gw_pag_op_compl_xml_rx <= 0 then
		msg_erro = "NSU GERADO É INVÁLIDO (" & id_pagto_gw_pag_op_compl_xml_rx & ")"
		exit function
		end if
	
	if msg_erro <> "" then
		msg_erro = "Não é possível enviar a solicitação à Braspag devido a uma falha:" & chr(13) & msg_erro
		exit function
		end if
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR WHERE (id = -1)"
	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	tPagtoGwPagOpComplementar.Open strSql, cn
	tPagtoGwPagOpComplementar.AddNew
	tPagtoGwPagOpComplementar("id") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementar("id_pagto_gw_pag") = CLng(id_pagto_gw_pag)
	tPagtoGwPagOpComplementar("id_pagto_gw_pag_payment") = CLng(id_pagto_gw_pag_payment)
	tPagtoGwPagOpComplementar("operacao") = BRASPAG_TIPO_TRANSACAO__PAG_REFUNDCREDITCARDTRANSACTION
	tPagtoGwPagOpComplementar("usuario") = usuario
	tPagtoGwPagOpComplementar("trx_TX_data") = Date
	tPagtoGwPagOpComplementar("trx_TX_data_hora") = Now
	tPagtoGwPagOpComplementar("req_RequestId") = trx.PAG_RequestId
	tPagtoGwPagOpComplementar("req_Version") = trx.PAG_Version
	tPagtoGwPagOpComplementar("req_MerchantId") = trx.PAG_MerchantId
	tPagtoGwPagOpComplementar("req_BraspagTransactionId") = trx.PAG_BraspagTransactionId
	tPagtoGwPagOpComplementar("req_Amount") = trx.PAG_Amount
	tPagtoGwPagOpComplementar("req_ServiceTaxAmount") = trx.PAG_ServiceTaxAmount
	tPagtoGwPagOpComplementar.Update
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	tPagtoGwPagOpComplementarXml.Open strSql, cn
	tPagtoGwPagOpComplementarXml.AddNew
	tPagtoGwPagOpComplementarXml("id") = id_pagto_gw_pag_op_compl_xml_tx
	tPagtoGwPagOpComplementarXml("id_pagto_gw_pag_op_complementar") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementarXml("data") = Date
	tPagtoGwPagOpComplementarXml("data_hora") = Now
	tPagtoGwPagOpComplementarXml("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_REFUNDCREDITCARDTRANSACTION
	tPagtoGwPagOpComplementarXml("fluxo_xml") = BRASPAG_FLUXO_XML__TX
	tPagtoGwPagOpComplementarXml("xml") = txXml
	tPagtoGwPagOpComplementarXml.Update
	
	rxXml = BraspagEnviaTransacaoComRetry(txXml, BRASPAG_WS_ENDERECO_PAGADOR_TRANSACTION)
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR WHERE (id = " & id_pagto_gw_pag_op_complementar & ")"
	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	tPagtoGwPagOpComplementar.Open strSql, cn
	if Not tPagtoGwPagOpComplementar.Eof then
		tPagtoGwPagOpComplementar("trx_RX_data") = Date
		tPagtoGwPagOpComplementar("trx_RX_data_hora") = Now
		tPagtoGwPagOpComplementar("trx_RX_status") = 1
		if Trim(rxXml) = "" then tPagtoGwPagOpComplementar("trx_RX_vazio_status") = 1
		tPagtoGwPagOpComplementar.Update
		end if
	
	strSql = "SELECT * FROM t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML WHERE (id = -1)"
	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	tPagtoGwPagOpComplementarXml.Open strSql, cn
	tPagtoGwPagOpComplementarXml.AddNew
	tPagtoGwPagOpComplementarXml("id") = id_pagto_gw_pag_op_compl_xml_rx
	tPagtoGwPagOpComplementarXml("id_pagto_gw_pag_op_complementar") = id_pagto_gw_pag_op_complementar
	tPagtoGwPagOpComplementarXml("data") = Date
	tPagtoGwPagOpComplementarXml("data_hora") = Now
	tPagtoGwPagOpComplementarXml("tipo_transacao") = BRASPAG_TIPO_TRANSACAO__PAG_REFUNDCREDITCARDTRANSACTION
	tPagtoGwPagOpComplementarXml("fluxo_xml") = BRASPAG_FLUXO_XML__RX
	tPagtoGwPagOpComplementarXml("xml") = rxXml
	tPagtoGwPagOpComplementarXml.Update
	
	set r_rx = BraspagCarregaDados_RefundCreditCardTransactionResponse(rxXml, msg_erro)
	if msg_erro <> "" then exit function
	
'	ATUALIZA O ÚLTIMO STATUS DA TRANSAÇÃO
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_REFUNDCREDITCARDTRANSACTIONRESPONSE_STATUS__REFUND_CONFIRMED then
		strVoidedDate = bd_monta_data_hora(Now)
		strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" & _
					" refunded_status = 1," & _
					" refunded_data = Convert(datetime, Convert(varchar(10),getdate(), 121), 121)," & _
					" refunded_data_hora = getdate()," & _
					" refunded_usuario = '" & usuario & "'," & _
					" ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNADA & "'," & _
					" resp_VoidedDate = " & strVoidedDate & "," & _
					" ult_atualizacao_data_hora = getdate()," & _
					" ult_atualizacao_usuario = '" & usuario & "'," & _
					" ult_id_pagto_gw_pag_payment_op_complementar = " & id_pagto_gw_pag_op_complementar & _
				" WHERE" & _
					" (id = " & id_pagto_gw_pag_payment & ")"
	elseif r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_REFUNDCREDITCARDTRANSACTIONRESPONSE_STATUS__REFUND_ACCEPTED then
		' OBSERVAÇÕES: a Cielo retorna na própria requisição se o estorno foi realizado ou não, mas a Getnet e Redecard informam
		' inicialmente apenas que a requisição foi recebida e o processamento é realizado em até D+1 ou D+2, dependendo do horário
		' em que a requisição foi realizada.
		' O controle do estorno pendente é feito através dos campos de status 'refund_pending_status', 'refund_pending_confirmado_status' e 'refund_pending_falha_status', pois
		' o campo 'ult_GlobalStatus' pode ser alterado em várias rotinas diferentes de atualização de status.
		strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" & _
					" refund_pending_status = 1," & _
					" refund_pending_data = Convert(datetime, Convert(varchar(10),getdate(), 121), 121)," & _
					" refund_pending_data_hora = getdate()," & _
					" refund_pending_usuario = '" & usuario & "'," & _
					" ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNO_PENDENTE & "'," & _
					" ult_atualizacao_data_hora = getdate()," & _
					" ult_atualizacao_usuario = '" & usuario & "'," & _
					" ult_id_pagto_gw_pag_payment_op_complementar = " & id_pagto_gw_pag_op_complementar & _
				" WHERE" & _
					" (id = " & id_pagto_gw_pag_payment & ")"
	else
		strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" & _
					" ult_id_pagto_gw_pag_payment_op_complementar = " & id_pagto_gw_pag_op_complementar & _
				" WHERE" & _
					" (id = " & id_pagto_gw_pag_payment & ")"
		end if
	cn.Execute strSql, lngRecordsAffected
	
	st_sucesso = 0
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_REFUNDCREDITCARDTRANSACTIONRESPONSE_STATUS__REFUND_CONFIRMED then st_sucesso = 1
	strSql = "UPDATE t_PAGTO_GW_PAG_OP_COMPLEMENTAR SET" & _
				" resp_AuthorizationCode = '" & r_rx.PAG_AuthorizationCode & "'," & _
				" resp_ProofOfSale = '" & r_rx.PAG_ProofOfSale & "'," & _
				" st_sucesso = " & CStr(st_sucesso) & "," & _
				" resp_Status = '" & r_rx.PAG_Status & "'" & _
			" WHERE" & _
				" (id = " & id_pagto_gw_pag_op_complementar & ")"
	cn.Execute strSql, lngRecordsAffected
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	if tPagtoGwPag.State <> 0 then tPagtoGwPag.Close
	set tPagtoGwPag = nothing

	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	set tPagtoGwPagPayment = nothing

	if tPagtoGwPagOpComplementar.State <> 0 then tPagtoGwPagOpComplementar.Close
	set tPagtoGwPagOpComplementar = nothing

	if tPagtoGwPagOpComplementarXml.State <> 0 then tPagtoGwPagOpComplementarXml.Close
	set tPagtoGwPagOpComplementarXml = nothing
	
	if r_rx.PAG_Status = BRASPAG_PAGADOR_CARTAO_REFUNDCREDITCARDTRANSACTIONRESPONSE_STATUS__REFUND_CONFIRMED then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if BraspagClearsaleRegistraPagtoNoPedido(BRASPAG_REGISTRA_PAGTO__OP_ESTORNO, pedido, id_pagto_gw_pag, id_pagto_gw_pag_payment, converte_numero(vl_transacao), usuario, msg_erro) then
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



function BraspagClearsaleGeraPedidoHistPagto(byval tipo_operacao, byval pedido, byval idPagtoGwPagPayment, byval vl_transacao, byval usuario, byref mensagem_erro)
dim idFinPedidoHistPagto, msg_erro
dim s, s_hist_pagto_status, s_hist_pagto_descricao, s_descricao_tipo_operacao
dim bandeira, pag_PaymentPlan, pag_NumberOfPayments
dim lngRecordsAffected
dim rs

	BraspagClearsaleGeraPedidoHistPagto = False

	if Not cria_recordset_pessimista(rs, msg_erro) then
		mensagem_erro = "Falha ao tentar abrir o recordset em modo de gravação: " & msg_erro
		exit function
		end if

	s = "SELECT * FROM t_PAGTO_GW_PAG_PAYMENT WHERE (id = " & idPagtoGwPagPayment & ")"
	if rs.State <> 0 then rs.Close
	rs.Open s, cn
	if rs.Eof then
		mensagem_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag (t_PAGTO_GW_PAG_PAYMENT.id=" & idPagtoGwPagPayment & ")"
		exit function
		end if

	pag_PaymentPlan = Trim("" & rs("req_PaymentDataRequest_PaymentPlan"))
	pag_NumberOfPayments = rs("req_PaymentDataRequest_NumberOfPayments")
	bandeira = Trim("" & rs("bandeira"))
	if rs.State <> 0 then rs.Close

'	REGISTRA NO HISTÓRICO DE PAGAMENTOS DO PEDIDO
	if Not fin_gera_nsu(T_FIN_PEDIDO_HIST_PAGTO, idFinPedidoHistPagto, msg_erro) then
		mensagem_erro = "Falha ao tentar gerar o NSU para o novo registro do histórico de pagamentos do pedido: " & msg_erro
		exit function
		end if
	
	s_descricao_tipo_operacao = BraspagDescricaoOperacaoRegistraPagto(tipo_operacao)
	s_hist_pagto_descricao = BraspagDescricaoBandeira(bandeira) & ": " & formata_moeda(Abs(vl_transacao))
	if (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CAPTURA) Or (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_AUTORIZACAO) then s_hist_pagto_descricao = s_hist_pagto_descricao & " em " & Cstr(pag_NumberOfPayments) & "x"
	s_hist_pagto_descricao = Left(s_hist_pagto_descricao, 60)
	if tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CAPTURA then
		s_hist_pagto_status = ST_T_FIN_PEDIDO_HIST_PAGTO__QUITADO
	elseif tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_AUTORIZACAO then
		s_hist_pagto_status = ST_T_FIN_PEDIDO_HIST_PAGTO__PREVISAO
	elseif (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CANCELAMENTO) Or (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_ESTORNO) then
		s_hist_pagto_status = ST_T_FIN_PEDIDO_HIST_PAGTO__CANCELADO
		s_hist_pagto_descricao = "(" & s_descricao_tipo_operacao & ") " & s_hist_pagto_descricao
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
			idPagtoGwPagPayment & ", " & _
			CTRL_PAGTO_MODULO__BRASPAG_CLEARSALE & ", " & _
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

	BraspagClearsaleGeraPedidoHistPagto = True
end function



' --------------------------------------------------------------------------------
'   BraspagClearsaleRegistraPagtoNoPedido
'   Registra o pagamento no pedido em decorrência de uma transação na Braspag
'   É necessário que a chamada desta função esteja dentro de uma transação,
'   a qual deve ser iniciada e finalizada pela rotina chamadora.
'   IMPORTANTE: AS ALTERAÇÕES NAS REGRAS DEVEM ESTAR SINCRONIZADAS ENTRE AS ROTINAS BraspagClearsaleRegistraPagtoNoPedido() de BraspagCS.asp E registraPagamentoNoPedido() DE BraspagDAO.cs (FinanceiroService)
function BraspagClearsaleRegistraPagtoNoPedido(byval tipo_operacao, byval pedido, byval id_pagto_gw_pag, byval id_pagto_gw_pag_payment, byval vl_transacao, byval usuario, byref mensagem_erro)
dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF
dim s, st_pagto_original, st_pagto_novo, st_pagto, s_id_pedido_pagto, msg_erro, s_log, id_pedido_base, loja, s_descricao_tipo_operacao
dim bandeira, pag_PaymentPlan, pag_NumberOfPayments
dim idFinPedidoHistPagto, s_hist_pagto_status, s_hist_pagto_descricao
dim lngRecordsAffected
dim tPagtoGwPagPayment, tPagtoGwAf, tPedidoBase, tFinPedidoHistPagto
dim s_ult_AF_GlobalStatus, s_descricao_ult_AF_GlobalStatus
dim blnTransacaoCapturada, blnRegistrarValorTransacao, blnUpdateHistPagto
dim id_fin_pedido_hist_pagto

	BraspagClearsaleRegistraPagtoNoPedido = False
	
	s_log = ""
	mensagem_erro = ""
	blnTransacaoCapturada = False
	blnRegistrarValorTransacao = False
	
	id_pedido_base = retorna_num_pedido_base(pedido)
	
	if Not cria_recordset_pessimista(tPedidoBase, msg_erro) then
		mensagem_erro = "Falha ao tentar abrir o recordset em modo de gravação: " & msg_erro
		exit function
		end if
	
	if Not cria_recordset_pessimista(tFinPedidoHistPagto, msg_erro) then
		mensagem_erro = "Falha ao tentar abrir o recordset em modo de gravação: " & msg_erro
		exit function
		end if
	
	if Not cria_recordset_pessimista(tPagtoGwAf, msg_erro) then
		mensagem_erro = "Falha ao tentar abrir o recordset em modo de gravação: " & msg_erro
		exit function
		end if
	
	if Not cria_recordset_pessimista(tPagtoGwPagPayment, msg_erro) then
		mensagem_erro = "Falha ao tentar abrir o recordset em modo de gravação: " & msg_erro
		exit function
		end if
	
	s = "SELECT * FROM t_PAGTO_GW_PAG_PAYMENT WHERE (id = " & id_pagto_gw_pag_payment & ")"
	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	tPagtoGwPagPayment.Open s, cn
	if tPagtoGwPagPayment.Eof then
		mensagem_erro = "Falha ao tentar localizar o registro da transação com o Pagador da Braspag (t_PAGTO_GW_PAG_PAYMENT.id=" & id_pagto_gw_pag_payment & ")"
		exit function
		end if
	
	if CLng(tPagtoGwPagPayment("captura_confirmada_status")) <> 0 then blnTransacaoCapturada = True
	
	if Not calcula_pagamentos(pedido, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then
		mensagem_erro = "Falha ao tentar calcular os pagamentos anteriores do pedido: " & msg_erro
		exit function
		end if
	
	if (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CANCELAMENTO) Or (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_ESTORNO) then
		if vl_transacao > 0 then vl_transacao = -1 * vl_transacao
		end if
	
	if (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CAPTURA) Or _
	   (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_ESTORNO) Or _
	   ((tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CANCELAMENTO) And blnTransacaoCapturada) then
	   blnRegistrarValorTransacao = True
	   end if
	
	if blnRegistrarValorTransacao then
	'	REGISTRA O PAGAMENTO NO PEDIDO
		if Not gera_nsu(NSU_PEDIDO_PAGAMENTO, s_id_pedido_pagto, msg_erro) then
			mensagem_erro = "Falha ao tentar gerar o NSU para o novo registro de pagamento no pedido: " & msg_erro
			exit function
			end if
		
		s = "INSERT INTO t_PEDIDO_PAGAMENTO (" & _
				"id, " & _
				"pedido, " & _
				"data, " & _
				"hora, " & _
				"valor, " & _
				"usuario, " & _
				"tipo_pagto, " & _
				"id_pagto_gw_pag_payment" & _
			") VALUES (" & _
				"'" & s_id_pedido_pagto & "', " & _
				"'" & pedido & "', " & _
				bd_formata_data(Date) & ", " & _
				"'" & retorna_so_digitos(formata_hora(Now)) & "', " & _
				bd_formata_numero(vl_transacao) & ", " & _
				"'" & usuario & "', " & _
				"'" & COD_PAGTO_GW_BRASPAG_CLEARSALE & "', " & _
				id_pagto_gw_pag_payment & _
			")"
		cn.Execute s, lngRecordsAffected
		if lngRecordsAffected <> 1 then
			mensagem_erro = "Falha ao tentar gravar o novo registro de pagamento no pedido!!"
			exit function
			end if
		end if
	
'	PROCESSA A SITUAÇÃO DO PEDIDO C/ RELAÇÃO AOS PAGAMENTOS (QUITADO, PAGO PARCIAL, NÃO-PAGO)
	s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & id_pedido_base & "')"
	if tPedidoBase.State <> 0 then tPedidoBase.Close
	tPedidoBase.Open s, cn
	if tPedidoBase.Eof then
		mensagem_erro = "Pedido-base " & id_pedido_base & " não foi encontrado!"
		exit function
		end if
	
	loja = tPedidoBase("loja")
	st_pagto_original = Trim("" & tPedidoBase("st_pagto"))
	
	if blnRegistrarValorTransacao then
	'	OBTÉM STATUS DO RESULTADO DA ANÁLISE ANTIFRAUDE P/ USAR DURANTE PROCESSAMENTO DO STATUS DA ANÁLISE DE CRÉDITO
		s_ult_AF_GlobalStatus = ""
		s_descricao_ult_AF_GlobalStatus = "(sem status)"
		s = "SELECT" & _
				" t_PAGTO_GW_AF.ult_Status" & _
			" FROM t_PAGTO_GW_PAG_PAYMENT" & _
				" INNER JOIN t_PAGTO_GW_AF ON (t_PAGTO_GW_PAG_PAYMENT.id_pagto_gw_af = t_PAGTO_GW_AF.id)" & _
			" WHERE" & _
				" (t_PAGTO_GW_PAG_PAYMENT.id = " & id_pagto_gw_pag_payment & ")"
		if tPagtoGwAf.State <> 0 then tPagtoGwAf.Close
		tPagtoGwAf.Open s, cn
		if Not tPagtoGwAf.Eof then
			s_ult_AF_GlobalStatus = Trim("" & tPagtoGwAf("ult_Status"))
			if s_ult_AF_GlobalStatus <> "" then
				s_descricao_ult_AF_GlobalStatus = s_ult_AF_GlobalStatus & " - " & ClearsaleDescricaoAFStatus(s_ult_AF_GlobalStatus)
				end if
			tPagtoGwAf.Close
			set tPagtoGwAf = nothing
			end if
		
	'	PAGO (QUITADO)
	'	~~~~~~~~~~~~~~
		if (vl_TotalFamiliaDevolucaoPrecoNF + vl_TotalFamiliaPago + vl_transacao) >= (vl_TotalFamiliaPrecoNF - MAX_VALOR_MARGEM_ERRO_PAGAMENTO) then
			st_pagto_novo = ST_PAGTO_PAGO
			if Trim("" & tPedidoBase("st_pagto")) <> st_pagto_novo then
				tPedidoBase("dt_st_pagto") = Date
				tPedidoBase("dt_hr_st_pagto") = Now
				tPedidoBase("usuario_st_pagto") = usuario
				end if
			tPedidoBase("st_pagto") = st_pagto_novo
			s_log = "Status de pagamento do pedido: quitado (st_pagto: " & st_pagto_original & " => " & st_pagto_novo & ")"
			if (vl_TotalFamiliaDevolucaoPrecoNF + vl_TotalFamiliaPago + vl_transacao) > vl_TotalFamiliaPrecoNF then
				s_log = s_log & " (excedeu " & SIMBOLO_MONETARIO & " " & _
						formata_moeda((vl_TotalFamiliaDevolucaoPrecoNF+vl_TotalFamiliaPago+vl_transacao)-vl_TotalFamiliaPrecoNF) & ")"
			elseif (vl_TotalFamiliaDevolucaoPrecoNF + vl_TotalFamiliaPago + vl_transacao) < vl_TotalFamiliaPrecoNF then
				s_log = s_log & " (faltou " & SIMBOLO_MONETARIO & " " & _
						formata_moeda(vl_TotalFamiliaPrecoNF-(vl_TotalFamiliaDevolucaoPrecoNF+vl_TotalFamiliaPago+vl_transacao)) & ")"
				end if
		'	ANÁLISE DE CRÉDITO
		'	IMPORTANTE: AS ALTERAÇÕES NAS REGRAS DEVEM ESTAR SINCRONIZADAS ENTRE AS ROTINAS BraspagClearsaleRegistraPagtoNoPedido() de BraspagCS.asp E registraPagamentoNoPedido() DE BraspagDAO.cs (FinanceiroService)
		'	===========
		'	OBSERVAÇÃO: NO FLUXO COM A CLEARSALE, O FLUXO SEGUE A SEGUINTE SEQUENCIA:
		'		1) ENVIO DA TRANSAÇÃO DE AUTORIZAÇÃO DO PAGAMENTO PARA A BRASPAG
		'		2) CASO O PAGAMENTO TENHA SIDO AUTORIZADO, É ENVIADA A TRANSAÇÃO PARA ANÁLISE ANTIFRAUDE (CLEARSALE)
		'		3-A) SE A ANÁLISE ANTIFRAUDE APROVOU A TRANSAÇÃO, É FEITA A CAPTURA DA TRANSAÇÃO
		'		3-B) SE A ANÁLISE ANTIFRAUDE REPROVOU A TRANSAÇÃO, É FEITO O CANCELAMENTO/ESTORNO
		'	ESTA ROTINA TENTA ESTAR MELHOR PREPARADA P/ A EVENTUAL SITUAÇÃO EM QUE O FLUXO SOFRA ALTERAÇÕES (EX: CAPTURA AUTOMÁTICA AO INVÉS DE PRÉ-AUTORIZAÇÃO, ANÁLISE AF ANTES DO PAGADOR, ETC),
		'	MAS NÃO GARANTE 100% DAS FUNCIONALIDADES, PRINCIPALMENTE NO TRATAMENTO DO REGISTRO AUTOMÁTICO DE PAGAMENTOS E ALTERAÇÕES NO STATUS DA ANÁLISE DE CRÉDITO.
			dim blnCreditoOkAutomaticoDesativado
			blnCreditoOkAutomaticoDesativado = False
			if blnCreditoOkAutomaticoDesativado then
				if (CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_VENDAS)) And (CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK)) And _
					(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)) And _
					(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)) And _
					(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)) And _
					(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) then
					if s_log <> "" then s_log = s_log & "; "
					s_log = s_log & " Análise de crédito: " & descricao_analise_credito(tPedidoBase("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS) & " (motivo: crédito Ok automático está desativado)"
					tPedidoBase("analise_credito") = CLng(COD_AN_CREDITO_PENDENTE_VENDAS)
					tPedidoBase("analise_credito_data") = Now
					tPedidoBase("analise_credito_usuario") = ID_USUARIO_SISTEMA
				else
					s_log = s_log & " Análise de crédito: status não foi alterado porque pedido encontra-se em " & descricao_analise_credito(tPedidoBase("analise_credito")) & " (motivo: crédito Ok automático está desativado)"
					end if
			else
				if (s_ult_AF_GlobalStatus = CLEARSALE_ANTIFRAUDE_STATUS__APROVACAO_MANUAL) then
				'	A ANÁLISE DE ANTIFRAUDE É FEITA POR EQUIPE PRÓPRIA E ASSUME-SE QUE DURANTE A ANÁLISE ANTIFRAUDE FORAM FEITAS TODAS AS VERIFICAÇÕES NECESSÁRIAS
				'	PORTANTO, A APROVAÇÃO POR PARTE DO ANALISTA INTERNO SIGNIFICA QUE A ANÁLISE DE CRÉDITO ESTÁ OK
				'	IMPORTANTE: O CLIENTE PODE REALIZAR O PAGAMENTO UTILIZANDO VÁRIOS CARTÕES DE CRÉDITO. ESSAS N TRANSAÇÕES SERÃO ENVIADAS JUNTAS P/ UMA ÚNICA REQUISIÇÃO DE
				'	==========  ANÁLISE ANTIFRAUDE P/ A CLEARSALE. QUANDO A ANÁLISE AF ESTIVER CONCLUÍDA, O PROCESSAMENTO FINAL DE CAPTURA OU CANCELAMENTO/ESTORNO DAS TRANSAÇÕES
				'	SERÁ FEITA NO FINANCEIROSERVICE. CADA TRANSAÇÃO DE CARTÃO ENVOLVIDA NO PAGAMENTO DO PEDIDO IRÁ ACIONAR UMA VEZ A ROTINA BraspagDAO.registraPagamentoNoPedido()
				'	PORTANTO, A 1ª TRANSAÇÃO IRÁ ALTERAR O PEDIDO PARA O STATUS DE PAGAMENTO 'PARCIAL' E APENAS QUANDO A ÚLTIMA TRANSAÇÃO FOR PROCESSADA, O PEDIDO FICARÁ
				'	COM O STATUS 'PAGO' E A ANÁLISE DE CRÉDITO DEVE FICAR C/ O STATUS 'OK'.
				'	ENTRETANTO, O SISTEMA NÃO SABE DE ANTEMÃO SE O PAGAMENTO SERÁ INTEGRALIZADO OU NÃO, PORTANTO, ENQUANTO O STATUS DE PAGAMENTO ESTIVER PARCIAL, O STATUS
				'	DA ANÁLISE DE CRÉDITO DEVE FICAR COMO 'PENDENTE VENDAS', POIS CASO O VALOR NÃO SEJA INTEGRALIZADO, ESSE SERÁ O STATUS FINAL.
					if ( _
							(CLng(tPedidoBase("analise_credito")) = CLng(COD_AN_CREDITO_ST_INICIAL)) _
							Or _
							(CLng(tPedidoBase("analise_credito")) = CLng(COD_AN_CREDITO_PENDENTE_VENDAS)) _
						) _
						And _
						( CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK) And _
							(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)) And _
							(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)) And _
							(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)) And _
							(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) ) then
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & " Análise de crédito: " & descricao_analise_credito(tPedidoBase("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_OK) & " (status AF Clearsale: '" & s_descricao_ult_AF_GlobalStatus & "')"
						tPedidoBase("analise_credito") = CLng(COD_AN_CREDITO_OK)
						tPedidoBase("analise_credito_data") = Now
						tPedidoBase("analise_credito_usuario") = ID_USUARIO_SISTEMA
					else
						s_log = s_log & " Análise de crédito: status não foi alterado porque pedido encontra-se em " & descricao_analise_credito(tPedidoBase("analise_credito")) & " (status AF Clearsale: '" & s_descricao_ult_AF_GlobalStatus & "')"
						end if
				else
					if (CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_VENDAS)) And (CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK)) And _
						(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)) And _
						(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)) And _
						(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)) And _
						(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) then
					'	EM CASO DE APROVAÇÃO AUTOMÁTICA, COLOCA-SE EM 'PENDENTE VENDAS' PARA DAR OPORTUNIDADE AO ANALISTA CONFERIR A TITULARIDADE DO CARTÃO
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & " Análise de crédito: " & descricao_analise_credito(tPedidoBase("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS) & " (status AF Clearsale: '" & s_descricao_ult_AF_GlobalStatus & "')"
						tPedidoBase("analise_credito") = CLng(COD_AN_CREDITO_PENDENTE_VENDAS)
						tPedidoBase("analise_credito_data") = Now
						tPedidoBase("analise_credito_usuario") = ID_USUARIO_SISTEMA
					else
						s_log = s_log & " Análise de crédito: status não foi alterado porque pedido encontra-se em " & descricao_analise_credito(tPedidoBase("analise_credito")) & " (status AF Clearsale: '" & s_descricao_ult_AF_GlobalStatus & "')"
						end if
					end if
				end if 'if blnCreditoOkAutomaticoDesativado then-else
	'	PAGAMENTO PARCIAL
	'	~~~~~~~~~~~~~~~~~
		elseif (vl_TotalFamiliaPago + vl_transacao) > 0 then
			st_pagto_novo = ST_PAGTO_PARCIAL
			if Trim("" & tPedidoBase("st_pagto")) <> st_pagto_novo then
				tPedidoBase("dt_st_pagto") = Date
				tPedidoBase("dt_hr_st_pagto") = Now
				tPedidoBase("usuario_st_pagto") = usuario
				end if
			tPedidoBase("st_pagto") = st_pagto_novo
			s_log = "Status de pagamento do pedido: pago parcial (st_pagto: " & st_pagto_original & " => " & st_pagto_novo & ")"
		'	SE O STATUS É 'PAGO PARCIAL', PODE TER HAVIDO UMA OPERAÇÃO DE CAPTURA OU DE CANCELAMENTO/ESTORNO. NESTE CASO, AS SEGUINTES PREMISSAS SÃO SEGUIDAS:
		'		1) SE O PEDIDO ESTIVER COM 'CRÉDITO OK', NÃO SERÁ ALTERADO DEVIDO A CANCELAMENTO/ESTORNO (DEFINIDO PELA ROSE EM 22/06/2016)
		'		2) SE O PEDIDO ESTIVER COM 'PENDENTE VENDAS', NÃO SERÁ ALTERADO. NÃO HÁ NECESSIDADE DE ATUALIZAR A DATA DA ÚLTIMA ALTERAÇÃO DE STATUS, POIS PEDIDOS C/ STATUS DE PAGTO 'PAGO PARCIAL' NÃO SÃO CANCELADOS AUTOMATICAMENTE
			if (CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_VENDAS)) And (CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK)) And _
				(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)) And _
				(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)) And _
				(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)) And _
				(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) then
				if s_log <> "" then s_log = s_log & "; "
				s_log = s_log & " Análise de crédito: " & descricao_analise_credito(tPedidoBase("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS) & " (status AF Clearsale: '" & s_descricao_ult_AF_GlobalStatus & "')"
				tPedidoBase("analise_credito") = CLng(COD_AN_CREDITO_PENDENTE_VENDAS)
				tPedidoBase("analise_credito_data") = Now
				tPedidoBase("analise_credito_usuario") = ID_USUARIO_SISTEMA
			else
				s_log = s_log & " Análise de crédito: status não foi alterado porque pedido encontra-se em " & descricao_analise_credito(tPedidoBase("analise_credito")) & " (status AF Clearsale: '" & s_descricao_ult_AF_GlobalStatus & "')"
				end if
	'	NÃO PAGO
	'	~~~~~~~~
		else
			st_pagto_novo = ST_PAGTO_NAO_PAGO
			if Trim("" & tPedidoBase("st_pagto")) <> st_pagto_novo then
				tPedidoBase("dt_st_pagto") = Date
				tPedidoBase("dt_hr_st_pagto") = Now
				tPedidoBase("usuario_st_pagto") = usuario
				end if
			tPedidoBase("st_pagto") = st_pagto_novo
			s_log = "Status de pagamento do pedido: não-pago (st_pagto: " & st_pagto_original & " => " & st_pagto_novo & ")"
		'	SE O STATUS É 'NÃO PAGO', ENTÃO OCORREU UMA OPERAÇÃO DE CANCELAMENTO/ESTORNO. NESTE CASO, AS SEGUINTES PREMISSAS SÃO SEGUIDAS:
		'		1) SE O PEDIDO ESTIVER COM 'CRÉDITO OK', NÃO SERÁ ALTERADO DEVIDO A CANCELAMENTO/ESTORNO (DEFINIDO PELA ROSE EM 22/06/2016)
		'		2) SE O PEDIDO ESTIVER COMO 'PENDENTE VENDAS', CONTINUA COMO ESTÁ E A DATA DA ÚLTIMA ALTERAÇÃO DE STATUS NÃO É ALTERADA, MANTENDO A CONTAGEM ORIGINAL DO PERÍODO DE CANCELAMENTO AUTOMÁTICO DE PEDIDOS
			if (CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_VENDAS)) And (CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK)) And _
				(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)) And _
				(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)) And _
				(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)) And _
				(CLng(tPedidoBase("analise_credito")) <> CLng(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) then
				if s_log <> "" then s_log = s_log & "; "
				s_log = s_log & " Análise de crédito: " & descricao_analise_credito(tPedidoBase("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS) & " (status AF Clearsale: '" & s_descricao_ult_AF_GlobalStatus & "')"
				tPedidoBase("analise_credito") = CLng(COD_AN_CREDITO_PENDENTE_VENDAS)
				tPedidoBase("analise_credito_data") = Now
				tPedidoBase("analise_credito_usuario") = ID_USUARIO_SISTEMA
			else
				s_log = s_log & " Análise de crédito: status não foi alterado porque pedido encontra-se em " & descricao_analise_credito(tPedidoBase("analise_credito")) & " (status AF Clearsale: '" & s_descricao_ult_AF_GlobalStatus & "')"
				end if
			end if
		
		tPedidoBase("vl_pago_familia") = vl_TotalFamiliaPago + vl_transacao
		tPedidoBase.Update
		if Err <> 0 then
			mensagem_erro = Cstr(Err) & ": " & Err.Description
			exit function
			end if
		end if 'if blnRegistrarValorTransacao then
	
	
	pag_PaymentPlan = Trim("" & tPagtoGwPagPayment("req_PaymentDataRequest_PaymentPlan"))
	pag_NumberOfPayments = tPagtoGwPagPayment("req_PaymentDataRequest_NumberOfPayments")
	
'	ANOTA NO REGISTRO DA TRANSAÇÃO QUE O PAGAMENTO JÁ FOI REGISTRADO NO PEDIDO
	bandeira = Trim("" & tPagtoGwPagPayment("bandeira"))
	if blnRegistrarValorTransacao then
		if (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CAPTURA) then
			tPagtoGwPagPayment("pagto_registrado_no_pedido_status") = 1
			tPagtoGwPagPayment("pagto_registrado_no_pedido_tipo_operacao") = tipo_operacao
			tPagtoGwPagPayment("pagto_registrado_no_pedido_data") = Date
			tPagtoGwPagPayment("pagto_registrado_no_pedido_data_hora") = Now
			tPagtoGwPagPayment("pagto_registrado_no_pedido_usuario") = usuario
			tPagtoGwPagPayment("pagto_registrado_no_pedido_id_pedido_pagamento") = s_id_pedido_pagto
			tPagtoGwPagPayment("pagto_registrado_no_pedido_st_pagto_anterior") = st_pagto_original
			tPagtoGwPagPayment("pagto_registrado_no_pedido_st_pagto_novo") = st_pagto_novo
		else
			tPagtoGwPagPayment("estorno_registrado_no_pedido_status") = 1
			tPagtoGwPagPayment("estorno_registrado_no_pedido_tipo_operacao") = tipo_operacao
			tPagtoGwPagPayment("estorno_registrado_no_pedido_data") = Date
			tPagtoGwPagPayment("estorno_registrado_no_pedido_data_hora") = Now
			tPagtoGwPagPayment("estorno_registrado_no_pedido_usuario") = usuario
			tPagtoGwPagPayment("estorno_registrado_no_pedido_id_pedido_pagamento") = s_id_pedido_pagto
			tPagtoGwPagPayment("estorno_registrado_no_pedido_st_pagto_anterior") = st_pagto_original
			tPagtoGwPagPayment("estorno_registrado_no_pedido_st_pagto_novo") = st_pagto_novo
			end if
		tPagtoGwPagPayment.Update
		
		s_descricao_tipo_operacao = BraspagDescricaoOperacaoRegistraPagto(tipo_operacao)
		s_log = "Registro automático de pagamento decorrente de operação de '" & s_descricao_tipo_operacao & "' na Braspag no valor de " & formata_moeda(vl_transacao) & " foi registrado com sucesso no pedido (t_PAGTO_GW_PAG_PAYMENT.id=" & Cstr(id_pagto_gw_pag_payment) & ", t_PEDIDO_PAGAMENTO.id=" & s_id_pedido_pagto & "): " & s_log & ", Bandeira: " & BraspagDescricaoBandeira(bandeira) & ", Valor: " & formata_moeda(Abs(vl_transacao)) & ", Opção Pagamento: " & BraspagDescricaoParcelamento(pag_PaymentPlan, pag_NumberOfPayments, Abs(vl_transacao))
		grava_log usuario, loja, pedido, "", OP_LOG_PEDIDO_PAGTO_CONTABILIZADO_BRASPAG_CLEARSALE, s_log
		end if 'if blnRegistrarValorTransacao then
	
'	Registra no histórico de pagamentos do pedido
'	Lembrando que no modelo em que se usa a Clearsale, a transação é primeiro autorizada pelo Pagador e somente em caso de sucesso é enviada p/ a Clearsale.
'	Após a análise da Clearsale, a transação é capturada ou cancelada. Entretanto, é possível que a transação seja capturada ou cancelada diretamente, de forma manual, pelo
'	painel de controle do Pagador, Cielo, Magento, etc.
	blnUpdateHistPagto = False
	s_hist_pagto_status = "0"
	if (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CAPTURA) Or _
	   ((tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CANCELAMENTO) And (Not blnTransacaoCapturada)) then
		s = "SELECT " & _
				"*" & _
			" FROM t_FIN_PEDIDO_HIST_PAGTO" & _
			" WHERE" & _
				" (ctrl_pagto_modulo = " & CTRL_PAGTO_MODULO__BRASPAG_CLEARSALE & ")" & _
				" AND (ctrl_pagto_id_parcela = " & id_pagto_gw_pag_payment & ")"
		if tFinPedidoHistPagto.State <> 0 then tFinPedidoHistPagto.Close
		tFinPedidoHistPagto.Open s, cn
		if Not tFinPedidoHistPagto.Eof then
			id_fin_pedido_hist_pagto = tFinPedidoHistPagto("id")
			if Trim("" & tFinPedidoHistPagto("status")) = Trim("" & ST_T_FIN_PEDIDO_HIST_PAGTO__PREVISAO) then
				blnUpdateHistPagto = True
				if (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CAPTURA) then
					s_hist_pagto_status = ST_T_FIN_PEDIDO_HIST_PAGTO__QUITADO
				elseif (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CANCELAMENTO) And (Not blnTransacaoCapturada) then
					s_hist_pagto_status = ST_T_FIN_PEDIDO_HIST_PAGTO__CANCELADO
					end if
				end if
			end if 'if Not tFinPedidoHistPagto.Eof
		end if
	
	if blnUpdateHistPagto then
		s = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET" & _
				" status = " & s_hist_pagto_status & "," & _
				" dt_ult_atualizacao = Convert(datetime, Convert(varchar(10),getdate(), 121), 121)," & _
				" usuario_ult_atualizacao = '" & usuario & "'" & _
			" WHERE" & _
				" (id = " & id_fin_pedido_hist_pagto & ")"
		cn.Execute s, lngRecordsAffected
		if lngRecordsAffected <> 1 then
			mensagem_erro = "Falha ao tentar atualizar dados no registro de histórico de pagamentos do pedido (t_FIN_PEDIDO_HIST_PAGTO.id=" & id_fin_pedido_hist_pagto & ")!!"
			exit function
			end if
	
	else
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
		elseif (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_ESTORNO) then
			s_hist_pagto_status = ST_T_FIN_PEDIDO_HIST_PAGTO__CANCELADO
			s_hist_pagto_descricao = "(Estorno) " & s_hist_pagto_descricao
		elseif (tipo_operacao = BRASPAG_REGISTRA_PAGTO__OP_CANCELAMENTO) And blnTransacaoCapturada then
			s_hist_pagto_status = ST_T_FIN_PEDIDO_HIST_PAGTO__CANCELADO
			s_hist_pagto_descricao = "(Cancelamento) " & s_hist_pagto_descricao
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
				id_pagto_gw_pag_payment & ", " & _
				CTRL_PAGTO_MODULO__BRASPAG_CLEARSALE  & ", " & _
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
		end if 'if blnUpdateHistPagto
	
	if tPagtoGwPagPayment.State <> 0 then tPagtoGwPagPayment.Close
	set tPagtoGwPagPayment = nothing

	if tPedidoBase.State <> 0 then tPedidoBase.Close
	set tPedidoBase = nothing

	if tFinPedidoHistPagto.State <> 0 then tFinPedidoHistPagto.Close
	set tFinPedidoHistPagto = nothing
	
	BraspagClearsaleRegistraPagtoNoPedido = True
end function
%>
