<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->
<!-- #include file = "../global/Braspag.asp"    -->
<!-- #include file = "../global/BraspagCS.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelBraspagClearsaleTransacoesExec.asp
'     ========================================================
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
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	Const EXIBIR_BOTAO_CAPTURA = True
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_BRASPAG_TRANSACOES, s_lista_operacoes_permitidas) then
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim s_filtro, intQtdeTransacoes
	intQtdeTransacoes = 0

	dim alerta
	dim s, s_aux
	dim c_dt_inicio, c_dt_termino
	dim c_resultado_transacao, c_bandeira, c_pedido, c_cliente_cnpj_cpf, c_loja, rb_ordenacao_saida, rb_tratadas
	dim s_nome_cliente, s_nome_loja

	alerta = ""

	c_dt_inicio = Trim(Request("c_dt_inicio"))
	c_dt_termino = Trim(Request("c_dt_termino"))
	c_resultado_transacao = Trim(Request("c_resultado_transacao"))
	c_bandeira = Trim(Request("c_bandeira"))
	c_pedido = Trim(Request("c_pedido"))
	c_cliente_cnpj_cpf = retorna_so_digitos(Trim(Request("c_cliente_cnpj_cpf")))
	c_loja = retorna_so_digitos(Trim(Request("c_loja")))
	rb_ordenacao_saida = Trim(Request("rb_ordenacao_saida"))
	rb_tratadas = Trim(Request("rb_tratadas"))
	
	s = normaliza_num_pedido(c_pedido)
	if s <> "" then c_pedido = s
	
	if alerta = "" then
		if c_pedido <> "" then
			s = "SELECT pedido FROM t_PEDIDO WHERE (pedido = '" & c_pedido & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta = "PEDIDO " & c_pedido & " NÃO ESTÁ CADASTRADO."
				end if
			end if
		end if
	
	if alerta = "" then
		s_nome_cliente = ""
		if c_cliente_cnpj_cpf <> "" then
			if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
				s = "SELECT TOP 1 dbo.SqlClrUtilIniciaisEmMaiusculas(endereco_nome) AS nome_iniciais_em_maiusculas FROM t_PEDIDO WHERE (endereco_cnpj_cpf = '" & c_cliente_cnpj_cpf & "') ORDER BY data_hora DESC"
			else
				s = "SELECT nome_iniciais_em_maiusculas FROM t_CLIENTE WHERE (cnpj_cpf = '" & c_cliente_cnpj_cpf & "')"
				end if
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "CLIENTE " & cnpj_cpf_formata(c_cliente_cnpj_cpf) & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_cliente = Trim("" & rs("nome_iniciais_em_maiusculas"))
				end if
			end if
		end if

	if alerta = "" then
		if c_loja <> "" then
			s = "SELECT nome, razao_social FROM t_LOJA WHERE (CONVERT(smallint,loja) = " & c_loja & ")"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta = "LOJA " & c_loja & " NÃO ESTÁ CADASTRADA."
			else
				s_nome_loja = iniciais_em_maiusculas(Trim("" & rs("nome")))
				if s_nome_loja = "" then s_nome_loja = iniciais_em_maiusculas(Trim("" & rs("razao_social")))
				end if
			end if
		end if


	dim qtde_transacoes
	qtde_transacoes = 0





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function monta_link_pedido(byval id_pedido, byval usuario)
dim strLink
	monta_link_pedido = ""
	id_pedido = Trim("" & id_pedido)
	if id_pedido = "" then exit function
	strLink = "<a href='javascript:fPEDConsulta(" & _
				chr(34) & id_pedido & chr(34) & _
				"," & _
				chr(34) & usuario & chr(34) & _
				")' title='clique para consultar o pedido " & id_pedido & "'>" & _
				id_pedido & "</a>"
	monta_link_pedido=strLink
end function



' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim s, s_sql, x
dim s_where
dim s_color
dim r
dim cab_table, cab
dim vl_total_geral
dim s_end_fatura, s_tel_pais, s_tel_ddd, s_tel_numero
dim s_class_alerta_titular_divergente
dim blnTitularCartaoDivergente

'	MONTAGEM DAS RESTRIÇÕES
	s_where = ""
	
	if c_bandeira <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (bandeira = '" & c_bandeira & "')"
		end if
	
	if c_dt_inicio <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if

	if c_dt_termino <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
	if c_pedido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (pedido = '" & c_pedido & "')"
		end if
	
	if c_cliente_cnpj_cpf <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
			'TRATA-SE DE UM WHERE EXTERNO, SENDO QUE O SELECT INTERNO NORMALIZOU O NOME DO CAMPO
			s_where = s_where & " (cnpj_cpf = '" & retorna_so_digitos(c_cliente_cnpj_cpf) & "')"
		else
			s_where = s_where & " (cnpj_cpf = '" & retorna_so_digitos(c_cliente_cnpj_cpf) & "')"
			end if
		end if
	
	if c_loja <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (numero_loja = " & c_loja & ")"
		end if
	
	if rb_tratadas = "SOMENTE_JA_TRATADAS" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tratado_manual_status = 1)"
	elseif rb_tratadas = "SOMENTE_NAO_TRATADAS" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tratado_manual_status = 0)"
		end if
	
'	MONTAGEM DA CONSULTA
	s_sql = "SELECT"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_cnpj_cpf AS cnpj_cpf," & _
				" dbo.SqlClrUtilIniciaisEmMaiusculas(t_PEDIDO.endereco_nome) AS cliente_nome," & _
				" t_PEDIDO.endereco_logradouro AS cliente_endereco," & _
				" t_PEDIDO.endereco_numero AS cliente_endereco_numero," & _
				" t_PEDIDO.endereco_complemento AS cliente_endereco_complemento," & _
				" t_PEDIDO.endereco_bairro AS cliente_bairro," & _
				" t_PEDIDO.endereco_cidade AS cliente_cidade," & _
				" t_PEDIDO.endereco_uf AS cliente_uf," & _
				" t_PEDIDO.endereco_cep AS cliente_cep,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.cnpj_cpf," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas AS cliente_nome," & _
				" t_CLIENTE.endereco AS cliente_endereco," & _
				" t_CLIENTE.endereco_numero AS cliente_endereco_numero," & _
				" t_CLIENTE.endereco_complemento AS cliente_endereco_complemento," & _
				" t_CLIENTE.bairro AS cliente_bairro," & _
				" t_CLIENTE.cidade AS cliente_cidade," & _
				" t_CLIENTE.uf AS cliente_uf," & _
				" t_CLIENTE.cep AS cliente_cep,"
		end if

	s_sql = s_sql & _
				" t_PEDIDO.numero_loja," & _
				" t_PEDIDO.st_end_entrega," & _
				" t_PEDIDO.EndEtg_endereco AS EndEtg_endereco," & _
				" t_PEDIDO.EndEtg_endereco_numero AS EndEtg_endereco_numero," & _
				" t_PEDIDO.EndEtg_endereco_complemento AS EndEtg_endereco_complemento," & _
				" t_PEDIDO.EndEtg_bairro AS EndEtg_bairro," & _
				" t_PEDIDO.EndEtg_cidade AS EndEtg_cidade," & _
				" t_PEDIDO.EndEtg_uf AS EndEtg_uf," & _
				" t_PEDIDO.EndEtg_cep AS EndEtg_cep," & _
				" tPAG.data," & _
				" tPAG.data_hora," & _
				" tPAG.usuario," & _
				" tPAG.pedido," & _
				" tPAG.valor_pedido," & _
				" tPAG.resp_OrderData_BraspagOrderId AS PAG_Resp_OrderData_BraspagOrderId," & _
				" tPAG.resp_Success AS PAG_Resp_Success," & _
				" tPAYMENT.id_pagto_gw_pag," & _
				" tPAYMENT.id AS id_pagto_gw_pag_payment," & _
				" tPAYMENT.valor_transacao," & _
				" tPAYMENT.bandeira," & _
				" tPAYMENT.ult_GlobalStatus," & _
				" tPAYMENT.ult_atualizacao_data_hora," & _
				" tPAYMENT.tratado_manual_status," & _
				" tPAYMENT.tratado_manual_obs," & _
				" tPAYMENT.req_PaymentDataRequest_PaymentPlan AS PAG_Req_PaymentDataRequest_PaymentPlan," & _
				" tPAYMENT.req_PaymentDataRequest_NumberOfPayments AS PAG_Req_PaymentDataRequest_NumberOfPayments," & _
				" tPAYMENT.req_PaymentDataRequest_CardHolder," & _
				" tPAYMENT.checkout_fatura_end_logradouro," & _
				" tPAYMENT.checkout_fatura_end_numero," & _
				" tPAYMENT.checkout_fatura_end_complemento," & _
				" tPAYMENT.checkout_fatura_end_bairro," & _
				" tPAYMENT.checkout_fatura_end_cidade," & _
				" tPAYMENT.checkout_fatura_end_uf," & _
				" tPAYMENT.checkout_fatura_end_cep," & _
				" tPAYMENT.checkout_fatura_tel_pais," & _
				" tPAYMENT.checkout_fatura_tel_ddd," & _
				" tPAYMENT.checkout_fatura_tel_numero," & _
				" tPAYMENT.resp_PaymentDataResponse_BraspagTransactionId AS PAG_Resp_PaymentDataResponse_BraspagTransactionId," & _
				" tPAYMENT.resp_PaymentDataResponse_AuthorizationCode AS PAG_Resp_PaymentDataResponse_AuthorizationCode," & _
				" tPAYMENT.resp_PaymentDataResponse_ProofOfSale AS PAG_Resp_PaymentDataResponse_ProofOfSale," & _
				" tPAYMENT.resp_PaymentDataResponse_ReturnMessage AS PAG_Resp_PaymentDataResponse_ReturnMessage," & _
				" tPAYMENT.refund_pending_status," & _
				" tPAYMENT.refund_pending_confirmado_status," & _
				" tPAYMENT.refund_pending_falha_status," & _
				" (SELECT TOP 1 ult_Status FROM t_PAGTO_GW_AF INNER JOIN t_PAGTO_GW_PAG_PAYMENT ON (t_PAGTO_GW_AF.id = t_PAGTO_GW_PAG_PAYMENT.id_pagto_gw_af) WHERE (t_PAGTO_GW_PAG_PAYMENT.id = tPAYMENT.id) AND (anulado_status = 0) AND (trx_RX_vazio_status = 0) AND (trx_RX_status = 1) AND (trx_erro_status = 0) ORDER BY t_PAGTO_GW_PAG_PAYMENT.id DESC) AS AF_Resp_Status," & _
				" (SELECT TOP 1 resp_Score FROM t_PAGTO_GW_AF INNER JOIN t_PAGTO_GW_PAG_PAYMENT ON (t_PAGTO_GW_AF.id = t_PAGTO_GW_PAG_PAYMENT.id_pagto_gw_af) WHERE (t_PAGTO_GW_PAG_PAYMENT.id = tPAYMENT.id) AND (anulado_status = 0) AND (trx_RX_vazio_status = 0) AND (trx_RX_status = 1) AND (trx_erro_status = 0) ORDER BY t_PAGTO_GW_PAG_PAYMENT.id DESC) AS AF_Resp_Score," & _
				" (SELECT TOP 1 ErrorCode + ' - ' + ErrorMessage FROM t_PAGTO_GW_PAG_ERROR tPAG_ERR WHERE (tPAG_ERR.id_pagto_gw_pag = tPAG.id) ORDER BY tPAG_ERR.id) AS PAG_ErrorMessage" & _
			" FROM t_PAGTO_GW_PAG tPAG" & _
				" INNER JOIN t_PEDIDO ON (tPAG.pedido = t_PEDIDO.pedido)" & _
				" INNER JOIN t_CLIENTE ON (tPAG.id_cliente = t_CLIENTE.id)" & _
				" INNER JOIN t_PAGTO_GW_PAG_PAYMENT tPAYMENT ON (tPAG.id = tPAYMENT.id_pagto_gw_pag)" & _
			" WHERE" & _
				" (operacao = '" & OP_BRASPAG_OPERACAO__AUTHORIZE & "')"
	
	if c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_CAPTURADA then
		s_sql = s_sql & _
				" AND (ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_AUTORIZADA then
		s_sql = s_sql & _
				" AND (ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_NAO_AUTORIZADA then
		s_sql = s_sql & _
				" AND (ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__NAO_AUTORIZADA & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_CAPTURA_CANCELADA then
		s_sql = s_sql & _
				" AND (ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURA_CANCELADA & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_ESTORNADA then
		s_sql = s_sql & _
				" AND (ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNADA & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_ESTORNO_PENDENTE then
		s_sql = s_sql & _
				" AND ((ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNO_PENDENTE & "') OR ((refund_pending_status=1) AND (refund_pending_confirmado_status=0) AND (refund_pending_falha_status=0)))"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_COM_ERRO_DESQUALIFICANTE then
		s_sql = s_sql & _
				" AND (ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ERRO_DESQUALIFICANTE & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_AGUARDANDO_RESPOSTA then
		s_sql = s_sql & _
				" AND (ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AGUARDANDO_RESPOSTA & "')"
		end if
	
	if s_sql <> "" then
		if s_where <> "" then s_where = " WHERE " & s_where
		s_sql = "SELECT " & _
					"*" & _
				" FROM (" & s_sql & ") t" & _
				s_where & _
				" ORDER BY"
		
		if rb_ordenacao_saida = "ORD_POR_PEDIDO" then
			s_sql = s_sql & _
						" pedido, id_pagto_gw_pag_payment"
		else
			s_sql = s_sql & _
						" id_pagto_gw_pag_payment"
			end if
		end if
	
	if s_sql = "" then
		Response.Write "Falha ao elaborar a consulta SQL: a consulta não possui conteúdo!"
		Response.End
		end if
	
	cab_table = "<table cellspacing='0' cellpadding='0'>" & chr(13)
	cab = "	<tr style='background:#F0FFFF;' nowrap>" & chr(13) & _
		"		<td class='MDTE tdDataHora' style='vertical-align:bottom'><span class='Rc'>Data</span></td>" & chr(13) & _
		"		<td class='MTD tdUsuario' style='vertical-align:bottom'><span class='Rc'>Usuário</span></td>" & chr(13) & _
		"		<td class='MTD tdPedido' style='vertical-align:bottom'><span class='R'>Pedido</span></td>" & chr(13) & _
		"		<td class='MTD tdVlPedido' style='vertical-align:bottom;padding-right:0px;'><span class='Rd'>Valor</span><br /><span class='Rd'>Pedido</span></td>" & chr(13) & _
		"		<td class='MTD tdVlTransacao' style='vertical-align:bottom;padding-right:0px;'><span class='Rd'>Valor</span><br /><span class='Rd'>Transação</span></td>" & chr(13) & _
		"		<td class='MTD tdBandeira' style='vertical-align:bottom'><span class='Rc'>Bandeira</span></td>" & chr(13) & _
		"		<td class='MTD tdStTransacao' style='vertical-align:bottom'><span class='Rc'>Status da</span><br /><span class='Rc'>Transação</span></td>" & chr(13) & _
		"		<td class='MTD tdDtHrStTransacao' style='vertical-align:bottom'><span class='Rc'>Data/Hora</span><br /><span class='Rc'>Status</span></td>" & chr(13) & _
		"		<td class='MTD tdCliente' style='vertical-align:bottom'><span class='R'>Cliente</span></td>" & chr(13) & _
		"		<td class='MTD tdFinalizado' style='vertical-align:bottom'><span class='Rc'>Tratado</span></td>" & chr(13) & _
		"		<td style='background:#FFFFFF;'>&nbsp;</TD>" & chr(13) & _
		"	</tr>" & chr(13)
	
	x = cab_table & cab
	intQtdeTransacoes = 0
	vl_total_geral = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	 ' CONTAGEM
		intQtdeTransacoes = intQtdeTransacoes + 1
		
		vl_total_geral = vl_total_geral + r("valor_transacao")
		
		x = x & "	<tr nowrap>" & chr(13)

	'> DATA DA TRANSAÇÃO
		s = formata_data_hora_sem_seg(r("data_hora"))
		x = x & "		<td class='MDTE tdDataHora'><span class='Cnc'>" & s & "</span></td>" & chr(13)

	'> USUÁRIO
		s = Trim("" & r("usuario"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdUsuario'><span class='Cnc'>" & s & "</span></td>" & chr(13)

	'> PEDIDO
		s = monta_link_pedido(Trim("" & r("pedido")), usuario)
		x = x & "		<td class='MTD tdPedido'><span class='Cn'>" & s & "</span></td>" & chr(13)

	'> VALOR DO PEDIDO
		s = formata_moeda(r("valor_pedido"))
		x = x & "		<td class='MTD tdVlPedido'><span class='Cnd'>" & s & "</span></td>" & chr(13)

	'> VALOR DA TRANSAÇÃO
		s = formata_moeda(r("valor_transacao"))
		x = x & "		<td class='MTD tdVlTransacao'><span class='Cnd'>" & s & "</span></td>" & chr(13)

	'> BANDEIRA DO CARTÃO
		s = BraspagDescricaoBandeira(Trim("" & r("bandeira")))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdBandeira'><span class='Cnc'>" & s & "</span></td>" & chr(13)

	'> STATUS DA TRANSAÇÃO
		s = Trim("" & r("ult_GlobalStatus"))
		if s <> "" then s = BraspagPagadorDescricaoGlobalStatus(s)
		if s = "" then s = "&nbsp;"
		if Trim("" & r("ult_GlobalStatus")) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA then
			s_color = "green"
		elseif Trim("" & r("ult_GlobalStatus")) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA then
			s_color = "black"
		else
			s_color = "red"
			end if
		
		if (Trim("" & r("ult_GlobalStatus")) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA) And (r("refund_pending_status")=1) And (r("refund_pending_confirmado_status")=0) And (r("refund_pending_falha_status")=0) then
			s = BraspagPagadorDescricaoGlobalStatus(BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNO_PENDENTE)
			s_color = "red"
			end if

		x = x & "		<td class='MTD tdStTransacao'><span id='spnStTransacao_" & Trim("" & r("id_pagto_gw_pag_payment")) & "' class='Cnc' style='color:" & s_color & ";'>" & s & "</span></td>" & chr(13)

	'> DATA/HORA STATUS
		s = formata_data_hora(Trim("" & r("ult_atualizacao_data_hora")))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdDtHrStTransacao'><span id='spnDtHrStTransacao_" & Trim("" & r("id_pagto_gw_pag_payment")) & "' class='Cnc'>" & s & "</span></td>" & chr(13)

	'> CLIENTE
		s = cnpj_cpf_formata(Trim("" & r("cnpj_cpf"))) & " - " & Trim("" & r("cliente_nome"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdCliente'><span class='Cn'>" & s & "</span></td>" & chr(13)

	'> TRATADO
	' PARA TRANSAÇÕES QUE FORAM MARCADAS COMO 'JÁ TRATADAS' ANTERIORMENTE PERMITE EDIÇÃO NO TEXTO DA OBSERVAÇÃO
		if r("tratado_manual_status") = 0 then
			x = x & _
					"		<td class='MTD tdFinalizado'>" & chr(13) & _
					"			<input type='checkbox' name='ckb_tratado' id='ckb_tratado' class='CheckTratado'" & _
									" value='" & Trim("" & r("id_pagto_gw_pag_payment")) & "|" & Trim("" & r("pedido")) & "|" & Trim("" & r("PAG_Resp_PaymentDataResponse_BraspagTransactionId")) & "'" & _
									" onclick=""configuraEdicaoCampoObs(this, '" & Trim("" & r("id_pagto_gw_pag_payment")) & "');""" & _
									">" & chr(13) & _
							"</td>" & chr(13)
		else
			x = x & _
					"		<td class='MTD tdFinalizado'>" & chr(13) & _
					"			<input type='checkbox' name='ckb_ja_tratado_readonly' class='CheckTratado CheckTratadoReadOnly'" & _
									" value='" & Trim("" & r("id_pagto_gw_pag_payment")) & "|" & Trim("" & r("pedido")) & "|" & Trim("" & r("PAG_Resp_PaymentDataResponse_BraspagTransactionId")) & "'" & _
									" checked='checked' disabled='disabled' />" & chr(13) & _
					"			<input type='hidden' name='c_flag_ja_tratado_" & Trim("" & r("id_pagto_gw_pag_payment")) & "' id='c_flag_ja_tratado_" & Trim("" & r("id_pagto_gw_pag_payment")) & "'" & _
									" value='S'" & _
									" />" & chr(13) & _
							"</td>" & chr(13)
			end if
		
	'> BOTÃO P/ EXIBIR DEMAIS CAMPOS
		x = x & "		<td valign='bottom' class='notPrint' align='left'>" & _
							"&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(intQtdeTransacoes) & chr(34) & ");' title='exibe ou oculta os campos adicionais'><img src='../botao/view_bottom.png' border='0'></a>" & _
						"</td>" & chr(13)
		
		x = x & "	</tr>" & chr(13)

	'> OUTRAS INFORMAÇÕES
		x = x & "	<tr style='display:none;' id='TR_MORE_INFO_" & Cstr(intQtdeTransacoes) & "'>" & chr(13) & _
				"		<td class='ME MD' align='left'>&nbsp;</td>" & chr(13) & _
				"		<td colspan='9' class='MC MD' align='left'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<tr>" & chr(13) & _
				"					<td class='Rf tdWithPadding' align='left'>OUTRAS INFORMAÇÕES</td>" & chr(13) & _
				"				</tr>" & chr(13)
		
		x = x & _
			"				<tr>" & chr(13) & _
			"					<td align='left'>" & chr(13) & _
			"						<table width='100%' cellspacing='0' cellpadding='0' border=0>" & chr(13)
		
		blnTitularCartaoDivergente = False
		if Not is_nome_e_sobrenome_iguais(Trim("" & r("cliente_nome")), Trim("" & r("req_PaymentDataRequest_CardHolder"))) then blnTitularCartaoDivergente = True
		s_class_alerta_titular_divergente = ""
		if blnTitularCartaoDivergente then s_class_alerta_titular_divergente = " colorRed"
		
	'	NOME DO CLIENTE
		s_aux = Trim("" & r("cliente_nome"))
		s_aux = Ucase(s_aux)
		if s_aux = "" then s_aux = "&nbsp;"
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco' rowspan='3'>&nbsp;</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Nome do Cliente:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo" & s_class_alerta_titular_divergente & "'>" & chr(13) & _
												s_aux & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
	
	'	ENDEREÇO NO CADASTRO DO CLIENTE
		s_aux = formata_endereco(Trim("" & r("cliente_endereco")), Trim("" & r("cliente_endereco_numero")), Trim("" & r("cliente_endereco_complemento")), Trim("" & r("cliente_bairro")), Trim("" & r("cliente_cidade")), Trim("" & r("cliente_uf")), Trim("" & r("cliente_cep")))
		s_aux = Ucase(s_aux)
		if s_aux = "" then s_end_fatura = "&nbsp;"
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Endereço (Cadastro do Cliente):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												s_aux & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)

	'	ENDEREÇO DE ENTREGA
		s_aux = ""
		if CLng(r("st_end_entrega")) <> 0 then
			s_aux = formata_endereco(Trim("" & r("EndEtg_endereco")), Trim("" & r("EndEtg_endereco_numero")), Trim("" & r("EndEtg_endereco_complemento")), Trim("" & r("EndEtg_bairro")), Trim("" & r("EndEtg_cidade")), Trim("" & r("EndEtg_uf")), Trim("" & r("EndEtg_cep")))
			end if
		s_aux = Ucase(s_aux)
		if s_aux = "" then s_end_fatura = "&nbsp;"
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Endereço de Entrega:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												s_aux & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
	
	'	TITULAR DO CARTÃO
		s_aux = Trim("" & r("req_PaymentDataRequest_CardHolder"))
		s_aux = Ucase(s_aux)
		if s_aux = "" then s_aux = "&nbsp;"
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco' rowspan='3'>&nbsp;</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Titular do Cartão:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo" & s_class_alerta_titular_divergente & "'>" & chr(13) & _
												s_aux & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	ENDEREÇO DA FATURA
		s_end_fatura = Trim("" & r("checkout_fatura_end_logradouro"))
		if s_end_fatura <> "" then
			s_aux = Trim("" & r("checkout_fatura_end_numero"))
			if s_aux <> "" then s_end_fatura = s_end_fatura & ", " & s_aux
			s_aux = Trim("" & r("checkout_fatura_end_complemento"))
			if s_aux <> "" then s_end_fatura = s_end_fatura & " " & s_aux
			s_aux = Trim("" & r("checkout_fatura_end_bairro"))
			if s_aux <> "" then s_end_fatura = s_end_fatura & " - " & s_aux
			s_aux = Trim("" & r("checkout_fatura_end_cidade"))
			if s_aux <> "" then s_end_fatura = s_end_fatura & " - " & s_aux
			s_aux = Trim("" & r("checkout_fatura_end_uf"))
			if s_aux <> "" then s_end_fatura = s_end_fatura & " - " & s_aux
			s_aux = Trim("" & r("checkout_fatura_end_cep"))
			if s_aux <> "" then s_end_fatura = s_end_fatura & " - " & cep_formata(s_aux)
			end if
		s_end_fatura = Ucase(s_end_fatura)
		if s_end_fatura = "" then s_end_fatura = "&nbsp;"
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Endereço (Fatura do Cartão):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												s_end_fatura & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	TELEFONE CADASTRADO NA FATURA
		s_aux = Trim("" & r("checkout_fatura_tel_numero"))
		if s_aux <> "" then
			s_tel_numero = telefone_formata(s_aux)
			s_tel_pais = Trim("" & r("checkout_fatura_tel_pais"))
			s_tel_ddd = Trim("" & r("checkout_fatura_tel_ddd"))
			s_aux = "(" & s_tel_ddd & ") " & s_tel_numero
			end if
		if s_aux = "" then s_end_fatura = "&nbsp;"
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Telefone (Fatura do Cartão):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												s_aux & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	OPÇÃO DE PAGAMENTO
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco' rowspan='2'>&nbsp;</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Opção de Pagamento:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(BraspagCSDescricaoParcelamento(Trim("" & r("PAG_Req_PaymentDataRequest_PaymentPlan")), Trim("" & r("PAG_Req_PaymentDataRequest_NumberOfPayments")), r("valor_transacao")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>NSU da Transação:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												Trim("" & r("id_pagto_gw_pag_payment")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	BLOCO DE DADOS: PAGADOR
		if UCase(Trim("" & r("PAG_Resp_Success"))) = "TRUE" then
			s = "Sucesso"
			s_color = "black"
		else
			s = "Falha"
			s_color = "red"
			end if
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco' rowspan='5' align='center' valign='middle'><img src='../Imagem/braspag-PAG-vert.png' width='12' height='46' /></td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Envio da Transação:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo' style='font-weight:bold;color:" & s_color & ";'>" & chr(13) & _
												s & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Código de Autorização:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("PAG_Resp_PaymentDataResponse_AuthorizationCode")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Comprovante de Venda:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("PAG_Resp_PaymentDataResponse_ProofOfSale")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Mensagem de Retorno:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("PAG_Resp_PaymentDataResponse_ReturnMessage")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Mensagem de Erro:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("PAG_ErrorMessage")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	BLOCO DE DADOS: ANTIFRAUDE
		if isAFStatusAprovado(Trim("" & r("AF_Resp_Status"))) then
			s_color = "green"
		elseif isAFStatusReprovado(Trim("" & r("AF_Resp_Status"))) then
			s_color ="red"
		else
			s_color ="black"
			end if
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco' rowspan='2' align='center' valign='middle'><img src='../Imagem/braspag-AF-vert.png' width='24' height='38' /></td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Status do Antifraude:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo' style='font-weight:bold;color:" & s_color & ";'>" & chr(13) & _
												"<span id='spnStatusAntifraude_" & Trim("" & r("id_pagto_gw_pag_payment")) & "'>" & ClearsaleDescricaoAFStatus(Trim("" & r("AF_Resp_Status"))) & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Score:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span>" & Trim("" & r("AF_Resp_Score")) & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco'>&nbsp;</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Observação:<br />" & _
												"<span class='lblTitTamRestante'>(tamanho restante:&nbsp;</span>" & _
												"<span class='lblTamRestante' name='lblTamRestante_" & Trim("" & r("id_pagto_gw_pag_payment")) & "' id='lblTamRestante_" & Trim("" & r("id_pagto_gw_pag_payment")) & "'></span>" & _
												"<span class='lblTitTamRestante'>)</span>" & _
											"</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
											"	<textarea style='display:none;'" & _
													" name='c_tratado_manual_obs_original_" & Trim("" & r("id_pagto_gw_pag_payment")) & "'" & _
													" id='c_tratado_manual_obs_original_" & Trim("" & r("id_pagto_gw_pag_payment")) & "'" & _
													">" & Trim("" & r("tratado_manual_obs")) & "</textarea>"  & chr(13) & _
											"	<textarea class='TextAreaObsEdit' maxlength='400' style='width:100%;height:100px;border:0px;'" & _
													" name='c_tratado_manual_obs_" & Trim("" & r("id_pagto_gw_pag_payment")) & "'" & _
													" id='c_tratado_manual_obs_" & Trim("" & r("id_pagto_gw_pag_payment")) & "'" & _
													">" & Trim("" & r("tratado_manual_obs")) & "</textarea>"  & chr(13) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	BOTÕES P/ ACIONAMENTO DE CONSULTAS/OPERAÇÕES
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco'>&nbsp;</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo' colspan='2' align='center' valign='middle'>" & chr(13) & _
			"									<table cellspacing='0' cellpadding='2' border='0'>" & chr(13) & _
			"										<tr>" & chr(13) & _
			"											<td align='left' style='width:10px;'>&nbsp;</td>" & chr(13) & _
			"											<td align='center'><a href='javascript:braspagClearsalePagConsultaStatus(" & Trim("" & r("id_pagto_gw_pag")) & "," & Trim("" & r("id_pagto_gw_pag_payment")) & ",""" & usuario & """);'><p class='Button BtnBraspag'>&nbsp;Consulta Status Atualizado&nbsp;</p></a></td>" & chr(13) & _
			"											<td align='left' style='width:15px;'>&nbsp;</td>" & chr(13) & _
			"											<td align='center'><a href='javascript:braspagClearsalePagExecutaVoidOrRefund(" & Trim("" & r("id_pagto_gw_pag")) & "," & Trim("" & r("id_pagto_gw_pag_payment")) & ",""" & Trim("" & r("pedido")) & """,""" & usuario & """);'><p class='Button BtnBraspag'>&nbsp;Cancelamento/Estorno&nbsp;</p></a></td>" & chr(13)
		
		if EXIBIR_BOTAO_CAPTURA then
			x = x & _
			"											<td align='left' style='width:15px;'>&nbsp;</td>" & chr(13) & _
			"											<td align='center'><a href='javascript:braspagClearsalePagExecutaCapture(" & Trim("" & r("id_pagto_gw_pag")) & "," & Trim("" & r("id_pagto_gw_pag_payment")) & ",""" & Trim("" & r("pedido")) & """,""" & usuario & """);'><p class='Button BtnBraspag'>&nbsp;Confirma Captura&nbsp;</p></a></td>" & chr(13)
			end if
		
		x = x & _
			"										</tr>" & chr(13) & _
			"									</table>" & chr(13) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		x = x & _
			"						</table>" & chr(13) & _
			"					</td>" & chr(13) & _
			"				</tr>" & chr(13)
		
		x = x & _
				"			</table>" & chr(13) & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13)

		if (intQtdeTransacoes mod 20) = 0 then
			Response.Write x
			x = ""
			end if

		r.MoveNext
		loop
	
	
'	TOTAL GERAL
	if intQtdeTransacoes > 0 then
		x = x & "	<tr>" & chr(13) & _
				"		<td colspan='4' align='right' class='MC ME'><span class='C'>TOTAL GERAL (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
				"		<td class='MC' align='right'><span class='Cd'>" & formata_moeda(vl_total_geral) & "</span></td>" & chr(13) & _
				"		<td colspan='5' class='MC MD' align='left'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr>" & chr(13) & _
				"		<td colspan='10' class='MC' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr nowrap style='background:#F0FFF0;'>" & chr(13) & _
				"		<td colspan='10' class='MT' align='left'><span class='C'>TOTAL: &nbsp; " & Cstr(intQtdeTransacoes) & iif((intQtdeTransacoes=1), " transação", " transações") & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if
	
'	MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeTransacoes = 0 then
		x = cab_table & cab
		x = x & "	<tr nowrap>" & chr(13) & _
				"		<td class='MT ALERTA' align='center' colspan='10'><span class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

'	FECHA TABELA
	x = x & "</table>" & chr(13)
	
	x = x & "<input type='hidden' name='c_qtde_transacoes' id='c_qtde_transacoes' value='" & Cstr(intQtdeTransacoes) & "'>" & chr(13)

	Response.write x

	qtde_transacoes = intQtdeTransacoes

	if r.State <> 0 then r.Close
	set r=nothing
	
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
	<title>CENTRAL</title>
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
		$("#divAjaxRunning").hide(); // Mantém oculto inicialmente
		$("#divPedidoConsulta").hide();

		sizeDivPedidoConsulta();

		$(document).ajaxStart(function() {
			$("#divAjaxRunning").show();
		})
		.ajaxStop(function() {
			$("#divAjaxRunning").hide();
		});

		$('#divInternoPedidoConsulta').addClass('divFixo');

		$(document).keyup(function(e) {
			if (e.keyCode == 27) fechaDivPedidoConsulta();
		});

		$("#divPedidoConsulta").click(function() {
			fechaDivPedidoConsulta();
		});

		$("#imgFechaDivPedidoConsulta").click(function() {
			fechaDivPedidoConsulta();
		});

		$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8

		$("textarea").attr("readonly", "readonly");

		$("textarea").focus(function() {
			if ($(this).attr("readonly") == "readonly") {
				$(this).css('background-color', '#FF6347');
				alert("Para poder escrever uma observação, é necessário antes assinalar esta transação como já tratada!!");
			}
			else {
				$(this).css('background-color', '#98FB98');
			}
		});

		$("textarea").blur(function() {
			$(this).css('background-color', '#FFFFFF');
			$(this).val($.trim($(this).val()));
		});

		$(":checkbox").each(function() {
			var s_value = $(this).val();
			var v = s_value.split("|");
			var s_id_registro = v[0];
			var s_campo = "#c_tratado_manual_obs_" + s_id_registro;
			if ($(this).is(":checked")) {
				$(s_campo).removeAttr("readonly");
			}
			else {
				$(s_campo).attr("readonly", "readonly");
			}
		});

		// PARA TRANSAÇÕES QUE FORAM MARCADAS COMO 'JÁ TRATADAS' ANTERIORMENTE PERMITE EDIÇÃO NO TEXTO DA OBSERVAÇÃO
		$(".CheckTratadoReadOnly").each(function() {
			var s_value = $(this).val();
			var v = s_value.split("|");
			var s_id_registro = v[0];
			var s_campo = "#c_tratado_manual_obs_" + s_id_registro;
			$(s_campo).removeAttr("readonly");
		});

		$("span.lblTamRestante").each(function() {
			var s_id = $(this).attr("id");
			s_id = "#" + s_id.replace("lblTamRestante_", "c_tratado_manual_obs_");
			var c = $(s_id);
			var n = MAX_TAM_T_PEDIDO_PAGTO_BRASPAG__TRATADO_MANUAL_OBS - $(c).val().length;
			$(this).html(n.toString());
		});

		$("textarea").bind("input keyup paste", function() {
			var s_id = $(this).attr("id");
			s_id = "#" + s_id.replace("c_tratado_manual_obs_", "lblTamRestante_");
			var c = $(s_id);
			var n = MAX_TAM_T_PEDIDO_PAGTO_BRASPAG__TRATADO_MANUAL_OBS - $(this).val().length;
			$(c).html(n.toString());
		});
	});

	//Every resize of window
	$(window).resize(function() {
		sizeDivAjaxRunning();
		sizeDivPedidoConsulta();
	});

	//Every scroll of window
	$(window).scroll(function() {
		sizeDivAjaxRunning();
	});
	
	//Dynamically assign height
	function sizeDivAjaxRunning() {
		var newTop = $(window).scrollTop() + "px";
		$("#divAjaxRunning").css("top", newTop);
	}

	function sizeDivPedidoConsulta() {
		var newHeight = $(document).height() + "px";
		$("#divPedidoConsulta").css("height", newHeight);
	}

	function fechaDivPedidoConsulta() {
		$(window).scrollTop(windowScrollTopAnterior);
		$("#divPedidoConsulta").fadeOut();
		$("#iframePedidoConsulta").attr("src", "");
	}
</script>

<script language="JavaScript" type="text/javascript">
var MAX_TAM_T_PEDIDO_PAGTO_BRASPAG__TRATADO_MANUAL_OBS=<%=MAX_TAM_T_PEDIDO_PAGTO_BRASPAG__TRATADO_MANUAL_OBS%>;
var windowScrollTopAnterior;

window.status = 'Aguarde, executando a consulta ...';

function braspagClearsalePagExecutaVoidOrRefund(_id_pagto_gw_pag, _id_pagto_gw_pag_payment, _pedido, _usuario){
var s_id, s_color;
var strMsg = "";
var strMsgErro = "";
var blnShowJsonFields = false;

if (!confirm("Confirma o cancelamento/estorno da transação do pedido " + _pedido + "?")) return;

var jqxhr = $.ajax({
	url: "../Global/AjaxBraspagClearsalePagCancelOuEstorno.asp",
	type: "POST",
	dataType: 'json',
	data: {
		id_pagto_gw_pag : _id_pagto_gw_pag.toString(),
		id_pagto_gw_pag_payment : _id_pagto_gw_pag_payment.toString(),
		usuario : _usuario
		}
	})
	.done(function(response){
		if ((response.PAG_ErrorCode.length > 0)||(response.PAG_ErrorMessage.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.PAG_ErrorCode + ": " + response.PAG_ErrorMessage;
		}
		if ((response.PAG_faultcode.length > 0)||(response.PAG_faultstring.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.PAG_faultcode + ": " + response.PAG_faultstring;
		}
		if (strMsgErro.length > 0) {
			strMsgErro = "Falha na requisição!!\n\n" + strMsgErro;
			alert(strMsgErro);
			return;
		}
		
		if (response.PAG_GlobalStatus == "<%=BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA%>") {
			s_color = "green";
		}
		else if (response.PAG_GlobalStatus == "<%=BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA%>") {
			s_color = "black";
		}
		else {
			s_color = "red";
		}
		s_id = "#spnStTransacao_" + response.id_pagto_gw_pag_payment;
		$(s_id).text(response.PAG_DescricaoGlobalStatus);
		$(s_id).css("color", s_color);
		
		s_id = "#spnDtHrStTransacao_" + response.id_pagto_gw_pag_payment;
		$(s_id).text(response.PAG_GlobalStatus_atualizacao_data_hora);
		
		if (blnShowJsonFields){
			$.each(response, function(key, val){
				if (strMsg.length > 0) strMsg += "\n";
				strMsg += key + ": " + val;
			});
			alert("Dados da resposta:\n" + strMsg);
		}
	})
	.fail(function(jqXHR, textStatus){
		var msgErro = "";
		if (textStatus.toString().length > 0) msgErro = "Mensagem de Status: " + textStatus.toString();
		try {
			if (jqXHR.status.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Status: " + jqXHR.status.toString();}
		} catch (e) { }

		try {
			if (jqXHR.statusText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Descrição do Status: " + jqXHR.statusText.toString();}
		} catch (e) { }
		
		try {
			if (jqXHR.responseText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Mensagem de Resposta: " + jqXHR.responseText.toString();}
		} catch (e) { }
		
		alert("Falha ao tentar processar a requisição!!\n\n" + msgErro);
	});
}

function braspagClearsalePagExecutaCapture(_id_pagto_gw_pag, _id_pagto_gw_pag_payment, _pedido, _usuario){
var s_id, s_color;
var strMsg = "";
var strMsgErro = "";
var blnShowJsonFields = false;

if (!confirm("Confirma a captura da transação do pedido " + _pedido + "?")) return;

var jqxhr = $.ajax({
	url: "../Global/AjaxBraspagClearsalePagCaptura.asp",
	type: "POST",
	dataType: 'json',
	data: {
		id_pagto_gw_pag : _id_pagto_gw_pag.toString(),
		id_pagto_gw_pag_payment : _id_pagto_gw_pag_payment.toString(),
		usuario : _usuario
		}
	})
	.done(function(response){
		if ((response.PAG_ErrorCode.length > 0)||(response.PAG_ErrorMessage.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.PAG_ErrorCode + ": " + response.PAG_ErrorMessage;
		}
		if ((response.PAG_faultcode.length > 0)||(response.PAG_faultstring.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.PAG_faultcode + ": " + response.PAG_faultstring;
		}
		if (strMsgErro.length > 0) {
			strMsgErro = "Falha na requisição!!\n\n" + strMsgErro;
			alert(strMsgErro);
			return;
		}
		
		if (response.PAG_GlobalStatus == "<%=BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA%>") {
			s_color = "green";
		}
		else if (response.PAG_GlobalStatus == "<%=BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA%>") {
			s_color = "black";
		}
		else {
			s_color = "red";
		}
		s_id = "#spnStTransacao_" + response.id_pagto_gw_pag_payment;
		$(s_id).text(response.PAG_DescricaoGlobalStatus);
		$(s_id).css("color", s_color);
		
		s_id = "#spnDtHrStTransacao_" + response.id_pagto_gw_pag_payment;
		$(s_id).text(response.PAG_GlobalStatus_atualizacao_data_hora);
		
		if (blnShowJsonFields){
			$.each(response, function(key, val){
				if (strMsg.length > 0) strMsg += "\n";
				strMsg += key + ": " + val;
			});
			alert("Dados da resposta:\n" + strMsg);
		}
	})
	.fail(function(jqXHR, textStatus){
		var msgErro = "";
		if (textStatus.toString().length > 0) msgErro = "Mensagem de Status: " + textStatus.toString();
		try {
			if (jqXHR.status.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Status: " + jqXHR.status.toString();}
		} catch (e) { }

		try {
			if (jqXHR.statusText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Descrição do Status: " + jqXHR.statusText.toString();}
		} catch (e) { }
		
		try {
			if (jqXHR.responseText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Mensagem de Resposta: " + jqXHR.responseText.toString();}
		} catch (e) { }
		
		alert("Falha ao tentar processar a requisição!!\n\n" + msgErro);
	});
}

function braspagClearsalePagConsultaStatus(_id_pagto_gw_pag, _id_pagto_gw_pag_payment, _usuario){
var s_id, s_color;
var strMsg = "";
var strMsgErro = "";
var blnShowJsonFields = false;
var jqxhr = $.ajax({
	url: "../Global/AjaxBraspagClearsalePagConsultaStatus.asp",
	type: "POST",
	dataType: 'json',
	data: {
		id_pagto_gw_pag : _id_pagto_gw_pag.toString(),
		id_pagto_gw_pag_payment : _id_pagto_gw_pag_payment.toString(),
		usuario : _usuario
		}
	})
	.done(function(response){
		if ((response.PAG_ErrorCode.length > 0)||(response.PAG_ErrorMessage.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.PAG_ErrorCode + ": " + response.PAG_ErrorMessage;
		}
		if ((response.PAG_faultcode.length > 0)||(response.PAG_faultstring.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.PAG_faultcode + ": " + response.PAG_faultstring;
		}
		if (strMsgErro.length > 0) {
			strMsgErro = "Falha na consulta!!\n\n" + strMsgErro;
			alert(strMsgErro);
			return;
		}
		
		if (response.PAG_GlobalStatus == "<%=BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA%>") {
			s_color = "green";
		}
		else if (response.PAG_GlobalStatus == "<%=BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA%>") {
			s_color = "black";
		}
		else {
			s_color = "red";
		}
		s_id = "#spnStTransacao_" + response.id_pagto_gw_pag_payment;
		$(s_id).text(response.PAG_DescricaoGlobalStatus);
		$(s_id).css("color", s_color);
		
		s_id = "#spnDtHrStTransacao_" + response.id_pagto_gw_pag_payment;
		$(s_id).text(response.PAG_GlobalStatus_atualizacao_data_hora);
		
		if (blnShowJsonFields){
			$.each(response, function(key, val){
				if (strMsg.length > 0) strMsg += "\n";
				strMsg += key + ": " + val;
			});
			alert("Dados da resposta:\n" + strMsg);
		}
	})
	.fail(function(jqXHR, textStatus){
		var msgErro = "";
		if (textStatus.toString().length > 0) msgErro = "Mensagem de Status: " + textStatus.toString();
		try {
			if (jqXHR.status.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Status: " + jqXHR.status.toString();}
		} catch (e) { }

		try {
			if (jqXHR.statusText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Descrição do Status: " + jqXHR.statusText.toString();}
		} catch (e) { }
		
		try {
			if (jqXHR.responseText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Mensagem de Resposta: " + jqXHR.responseText.toString();}
		} catch (e) { }
		
		alert("Falha ao tentar processar a consulta!!\n\n" + msgErro);
	});
}

function expandirTudo() {
var i;
var row_MORE_INFO;
	for (i = 1; i <= intQtdeTransacoes; i++) {
		row_MORE_INFO = document.getElementById("TR_MORE_INFO_" + i);
		row_MORE_INFO.style.display = "";
	}
}

function recolherTudo() {
var i;
var row_MORE_INFO;
	for (i = 1; i <= intQtdeTransacoes; i++) {
		row_MORE_INFO = document.getElementById("TR_MORE_INFO_" + i);
		row_MORE_INFO.style.display = "none";
	}
}

function marcarTodas(){
	$(":checkbox").each(function() {
		if (!$(this).is(":checked")) {
			$(this).trigger('click');
		}
	});
}

function desmarcarTodas(){
	$(":checkbox").each(function() {
		if ($(this).is(":checked")) {
			$(this).trigger('click');
		}
	});
}

function fExibeOcultaCampos(indice_row) {
var row_MORE_INFO;

	row_MORE_INFO = document.getElementById("TR_MORE_INFO_" + indice_row);
	if (row_MORE_INFO.style.display.toString() == "none") {
		row_MORE_INFO.style.display = "";
	}
	else {
		row_MORE_INFO.style.display = "none";
	}
}

function configuraEdicaoCampoObs(ckb, id_registro) {
var s_id;
	s_id = "#c_tratado_manual_obs_" + id_registro;
	if ($(ckb).is(':checked')) {
		$(s_id).removeAttr("readonly");
	}
	else {
		$(s_id).attr("readonly", "readonly");
	}
}

function fPEDConsulta(id_pedido, usuario) {
	windowScrollTopAnterior = $(window).scrollTop();
	sizeDivPedidoConsulta();
	$("#iframePedidoConsulta").attr("src","PedidoConsultaView.asp?pedido_selecionado="+id_pedido+"&pedido_selecionado_inicial="+id_pedido+"&usuario="+usuario);
	$("#divPedidoConsulta").fadeIn();
}

function fFiltroVoltar(){
	fPED.action = "RelBraspagClearsaleTransacoesFiltro.asp";
	fPED.submit();
}

function fPEDGravaDados(f) {
var i, intQtdeTratados, c, blnExcedeu;

// O TEXTO DO CAMPO OBS É GRAVADO DE DUAS MANEIRAS: JUNTO COM UMA TRANSAÇÃO QUE ESTÁ SENDO MARCADA COMO 'JÁ TRATADA'
// OU ATRAVÉS DA EDIÇÃO DE UMA TRANSAÇÃO MARCADA COMO 'JÁ TRATADA' ANTERIORMENTE
	f.c_lista_transacoes_obs_editado.value = "";
	$(".TextAreaObsEdit").each(function(){
		var sNome = $(this).attr("name");
		var vSplit = sNome.split("_");
		var s_idReg = vSplit[vSplit.length - 1];
		if ($("#c_flag_ja_tratado_"+s_idReg).val() == "S"){
			if ($("#c_tratado_manual_obs_"+s_idReg).val() != $("#c_tratado_manual_obs_original_"+s_idReg).val()){
				if (f.c_lista_transacoes_obs_editado.value.length > 0) f.c_lista_transacoes_obs_editado.value += "|";
				f.c_lista_transacoes_obs_editado.value += s_idReg;
			}
		}
	});
	
	intQtdeTratados = 0;
	for (i = 0; i < f.ckb_tratado.length; i++) {
		if (f.ckb_tratado[i].checked) intQtdeTratados++;
	}

	if ((intQtdeTratados == 0)&&(f.c_lista_transacoes_obs_editado.value.length == 0)) {
		alert('Nenhuma transação foi assinalada para ser marcada como já tratada ou teve edição no texto da observação!!');
		return;
	}
	
	blnExcedeu = false;
	$("textarea").each(function(){
		if (blnExcedeu) return;
		if ($(this).val().length > MAX_TAM_T_PEDIDO_PAGTO_BRASPAG__TRATADO_MANUAL_OBS){
			blnExcedeu = true;
			c = $(this);
			return;
		}
	});
	
	if (blnExcedeu){
		alert("O texto de observação excedeu o tamanho máximo!!");
		c.focus();
		return;
	}
	
	window.status = "Aguarde ...";
	f.action = "RelBraspagClearsaleTransacoesGravaDados.asp";
	f.submit();
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__ESCREEN_CSS%>" rel="stylesheet" type="text/css" media="screen">

<style type="text/css">
html
{
	overflow-y: scroll;
	height:100%;
	margin:0px;
}
body
{
	height:100%;
	margin:0px;
}
.colorRed {
	color: red;
}
.tdWithPadding
{
	padding:1px;
}
.tdDataHora
{
	text-align: center;
	vertical-align: middle;
	width: 65px;
}
.tdUsuario{
	text-align: center;
	vertical-align: middle;
	width: 85px;
}
.tdPedido{
	text-align: left;
	vertical-align: middle;
	font-weight: bold;
	width: 65px;
}
.tdVlPedido
{
	text-align: right;
	vertical-align: middle;
	font-weight: bold;
	width: 70px;
}
.tdVlTransacao
{
	text-align: right;
	vertical-align: middle;
	font-weight: bold;
	width: 70px;
}
.tdBandeira
{
	text-align:center;
	vertical-align: middle;
	width: 75px;
}
.tdFinalizado
{
	text-align:center;
	vertical-align: middle;
	font-weight: bold;
	width: 60px;
}
.tdStTransacao
{
	text-align: center;
	vertical-align: middle;
	font-weight: bold;
	width: 85px;
}
.tdDtHrStTransacao
{
	text-align: center;
	vertical-align: middle;
	width: 65px;
}
.tdCliente{
	text-align: left;
	vertical-align: middle;
	width: 245px;
}
.tdTitMoreInfo{
	text-align: right;
	vertical-align: top;
	padding-right: 2px;
	width: 200px;
}
.tdMoreInfo
{
	text-align: left;
	vertical-align: top;
	padding-left: 2px;
	width: 580px;
}
.lblTitTamRestante
{
	color: #696969;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 8pt;
	font-style: normal;
	text-align: right;
}
.lblTamRestante
{
	color: #696969;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 8pt;
	font-style: normal;
	text-align: right;
}
.tdTitBloco
{
	padding: 2px;
	width: 40px;
}
.BtnBraspag
{
	margin-top:1px;
	margin-bottom:1px;
}
.BtnAll
{
	margin-top:0px;
	margin-bottom:0px;
}
#divAjaxRunning
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	z-index:1001;
	background-color:grey;
	opacity: .6;
}
.AjaxImgLoader
{
	position: absolute;
	left: 50%;
	top: 50%;
	margin-left: -128px; /* -1 * image width / 2 */
	margin-top: -128px;  /* -1 * image height / 2 */
	display: block;
}
#divPedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsulta
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
#divInternoPedidoConsulta.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivPedidoConsulta
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframePedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
</style>


<body onload="window.status='Concluído';focus();" link=#000000 alink=#000000 vlink=#000000>
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>" />
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>" />
<input type="hidden" name="c_resultado_transacao" id="c_resultado_transacao" value="<%=c_resultado_transacao%>" />
<input type="hidden" name="c_bandeira" id="c_bandeira" value="<%=c_bandeira%>" />
<input type="hidden" name="c_pedido" id="c_pedido" value="<%=c_pedido%>" />
<input type="hidden" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" value="<%=c_cliente_cnpj_cpf%>" />
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>" />
<input type="hidden" name="rb_ordenacao_saida" id="rb_ordenacao_saida" value="<%=rb_ordenacao_saida%>" />
<input type="hidden" name="rb_tratadas" id="rb_tratadas" value="<%=rb_tratadas%>" />

<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<input type="hidden" name="c_usuario_sessao" id="c_usuario_sessao" value="<%=usuario%>" />
<input type="hidden" name="c_lista_transacoes_obs_editado" id="c_lista_transacoes_obs_editado" value="" />

<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="ckb_tratado" id="ckb_tratado" value="">
<input type="hidden" name="c_tratado_manual_obs_0" id="c_tratado_manual_obs_0" value="">



<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="920" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Transações Braspag/Clearsale</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='920' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black;' border='0'>" & chr(13)

'	PERÍODO
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Período:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	RESULTADO DA TRANSAÇÃO
	s = c_resultado_transacao
	if s = "" then
		s = "N.I."
	else
		s = iniciais_em_maiusculas(descricao_cod_rel_transacoes_braspag(c_resultado_transacao))
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Resultado da Transação:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	BANDEIRA
	s = c_bandeira
	if s = "" then
		s = "N.I."
	else
		s = BraspagDescricaoBandeira(c_bandeira)
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Bandeira:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

'	PEDIDO
	s = c_pedido
	if s = "" then s = "N.I."
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Pedido:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

'	CLIENTE
	s = c_cliente_cnpj_cpf
	if s = "" then 
		s = "N.I."
	else
		s = cnpj_cpf_formata(c_cliente_cnpj_cpf)
		if s_nome_cliente <> "" then s = s & " - " & s_nome_cliente
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Cliente:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

'	LOJA
	s = c_loja
	if s = "" then 
		s = "N.I."
	else
		if s_nome_loja <> "" then s = s & " - " & s_nome_loja
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Loja:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

'	ORDENAÇÃO
	s = rb_ordenacao_saida
	if s = "ORD_POR_PEDIDO" then
		s = "Pedido"
	else
		s = "Data"
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Ordenação do resultado:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	TRANSAÇÕES TRATADAS
	s = rb_tratadas
	if s = "SOMENTE_JA_TRATADAS" then
		s = "Somente Já Tratadas"
	elseif s = "SOMENTE_NAO_TRATADAS" then
		s = "Somente Não Tratadas"
	else
		s = "Todas"
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Transações Tratadas:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	EMISSÃO
	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap>" & _
					"<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
					"<span class='N'>" & formata_data_hora(Now) & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
	
	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>

<% consulta_executa %>

<script language="JavaScript" type="text/javascript">
var intQtdeTransacoes=<%=Cstr(intQtdeTransacoes)%>;
</script>


<!-- ************   SEPARADOR   ************ -->
<table width="920" cellpadding="0" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width='920' cellpadding='0' cellspacing='0' border='0' style="margin-top:5px;">
<tr>
	<td width="25%" align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkMarcarTudo" href="javascript:marcarTodas();"><p class="Button BtnAll" style="margin-bottom:0px;">Marcar Todas</p></a></td>
	<td align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkDesmarcarTudo" href="javascript:desmarcarTodas();"><p class="Button BtnAll" style="margin-bottom:0px;">Desmarcar Todas</p></a></td>
	<td align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkImprimir" href="javascript:window.print();"><p class="Button BtnAll" style="margin-bottom:0px;">Imprimir...</p></a></td>
	<td align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkExpandirTudo" href="javascript:expandirTudo();"><p class="Button BtnAll" style="margin-bottom:0px;">Expandir Tudo</p></a></td>
	<td align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkRecolherTudo" href="javascript:recolherTudo();"><p class="Button BtnAll" style="margin-bottom:0px;">Recolher Tudo</p></a></td>
</tr>
</table>

<br />
<table class="notPrint" width="920" cellspacing="0" border="0">
<tr>
	<% if qtde_transacoes > 0 then %>
	<td align="left">
		<a name="bVOLTAR" id="bVOLTAR" href="javascript:fFiltroVoltar();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="left">&nbsp;</td>
	<td align="right">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDGravaDados(fPED)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0"></a>
	</td>
	<% else %>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fFiltroVoltar();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	<% end if %>
	</td>
</tr>
</table>

</form>

</center>

<div id="divAjaxRunning"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>
<div id="divPedidoConsulta"><center><div id="divInternoPedidoConsulta"><img id="imgFechaDivPedidoConsulta" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframePedidoConsulta"></iframe></div></center></div>

</body>
</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
