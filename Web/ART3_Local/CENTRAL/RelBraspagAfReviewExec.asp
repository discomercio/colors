<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->
<!-- #include file = "../global/Braspag.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelBraspagAfReviewExec.asp
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
	if Not operacao_permitida(OP_CEN_REL_BRASPAG_AF_REVIEW, s_lista_operacoes_permitidas) then
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim s_filtro, intQtdeTransacoes
	intQtdeTransacoes = 0

	dim alerta
	dim s, s_aux
	dim c_dt_transacao_inicio, c_dt_transacao_termino, c_dt_tratado_inicio, c_dt_tratado_termino
	dim c_resultado_transacao, c_bandeira, c_pedido, c_cliente_cnpj_cpf, c_loja, rb_ordenacao_saida, rb_tratadas
	dim s_nome_cliente, s_nome_loja

	alerta = ""

	c_dt_transacao_inicio = Trim(Request("c_dt_transacao_inicio"))
	c_dt_transacao_termino = Trim(Request("c_dt_transacao_termino"))
	c_dt_tratado_inicio = Trim(Request("c_dt_tratado_inicio"))
	c_dt_tratado_termino = Trim(Request("c_dt_tratado_termino"))
	c_resultado_transacao = Trim(Request("c_resultado_transacao"))
	c_bandeira = Trim(Request("c_bandeira"))
	c_pedido = Trim(Request("c_pedido"))
	c_cliente_cnpj_cpf = retorna_so_digitos(Trim(Request("c_cliente_cnpj_cpf")))
	c_loja = retorna_so_digitos(Trim(Request("c_loja")))
	rb_ordenacao_saida = Trim(Request("rb_ordenacao_saida"))
	rb_tratadas = Trim(Request("rb_tratadas"))
	
	if (c_resultado_transacao = COD_REL_BRASPAG_AF_REVIEW__APROVADO_AUTOMATICAMENTE) Or _
		(c_resultado_transacao = COD_REL_BRASPAG_AF_REVIEW__REJEITADO_AUTOMATICAMENTE) then
		rb_tratadas = ""
		end if

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

function monta_link_pedido(byval id_pedido, byval usuario, byval color)
dim strLink, strColor
	monta_link_pedido = ""
	id_pedido = Trim("" & id_pedido)
	if id_pedido = "" then exit function
	
	strColor = Trim("" & color)
	if strColor <> "" then strColor = "color:" & strColor & ";"
	
	strLink = "<a href='javascript:fPEDConsulta(" & _
				chr(34) & id_pedido & chr(34) & _
				"," & _
				chr(34) & usuario & chr(34) & _
				")' title='clique para consultar o pedido " & id_pedido & "'" & _
				" style='" & strColor & "'" & _
				">" & _
				id_pedido & "</a>"
	monta_link_pedido=strLink
end function



' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const PEDIDOS_POR_LINHA = 8
const MAX_PEDIDOS_POR_LINHA_ANALISE_ENDERECO = 8
dim s, s_aux, s_dados, s_sql, x, x2
dim s_where
dim s_color, s_class
dim r, r2, r3
dim cab_table, cab
dim vl_total_geral
dim v, i
dim strPedidosAnteriores, strColor
dim n_ped_linha
dim strInfoAnEnd, blnAnEnderecoUsaEndParceiro
dim intQtdePedidoAnEndereco, intQtdeTotalPedidosAnEndereco, intQtdeLinhasPedidoAnEndereco
dim intResto
dim s_end_fatura, s_tel_pais, s_tel_ddd, s_tel_numero
dim s_prim_AF_GlobalStatus, s_class_alerta_titular_divergente
dim blnTitularCartaoDivergente

	If Not cria_recordset_otimista(r2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(r3, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	MONTAGEM DAS RESTRIÇÕES
	s_where = ""
	
	if c_bandeira <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (bandeira = '" & c_bandeira & "')"
		end if
	
'	PERÍODO: TRANSAÇÃO EFETUADA
	if c_dt_transacao_inicio <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (data >= " & bd_formata_data(StrToDate(c_dt_transacao_inicio)) & ")"
		end if
	
	if c_dt_transacao_termino <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (data < " & bd_formata_data(StrToDate(c_dt_transacao_termino)+1) & ")"
		end if
	
'	PERÍODO: TRATAMENTO DA REVISÃO
	if c_dt_tratado_inicio <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (AF_review_tratado_data >= " & bd_formata_data(StrToDate(c_dt_tratado_inicio)) & ")"
		end if
	
	if c_dt_tratado_termino <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (AF_review_tratado_data < " & bd_formata_data(StrToDate(c_dt_tratado_termino)+1) & ")"
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
	
	if c_resultado_transacao = COD_REL_BRASPAG_AF_REVIEW__APROVADO_AUTOMATICAMENTE then
		s_prim_AF_GlobalStatus = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ACCEPT
	elseif c_resultado_transacao = COD_REL_BRASPAG_AF_REVIEW__REJEITADO_AUTOMATICAMENTE then
		s_prim_AF_GlobalStatus = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REJECT
	else
		s_prim_AF_GlobalStatus = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW
		end if
	
	if rb_tratadas = "SOMENTE_JA_TRATADAS" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (AF_review_tratado_status = 1)"
	elseif rb_tratadas = "SOMENTE_NAO_TRATADAS" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (AF_review_tratado_status = 0)"
		end if
	
'	MONTAGEM DA CONSULTA
	s_sql = "SELECT" & _
				" t_PEDIDO.data AS data_pedido," & _
				" t_PEDIDO.hora AS hora_pedido," & _
				" t_PEDIDO__BASE.analise_credito," & _
				" t_PEDIDO__BASE.analise_endereco_tratar_status," & _
				" t_PEDIDO.st_end_entrega," & _
				" t_PEDIDO.EndEtg_endereco AS EndEtg_endereco," & _
				" t_PEDIDO.EndEtg_endereco_numero AS EndEtg_endereco_numero," & _
				" t_PEDIDO.EndEtg_endereco_complemento AS EndEtg_endereco_complemento," & _
				" t_PEDIDO.EndEtg_bairro AS EndEtg_bairro," & _
				" t_PEDIDO.EndEtg_cidade AS EndEtg_cidade," & _
				" t_PEDIDO.EndEtg_uf AS EndEtg_uf," & _
				" t_PEDIDO.EndEtg_cep AS EndEtg_cep," & _
				" t_PEDIDO_PAGTO_BRASPAG.*," & _
				" tPPB_PAG.id AS id_pedido_pagto_braspag_pag," & _
				" tPPB_AF.id AS id_pedido_pagto_braspag_af,"

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
				" tPPB_PAG.Req_PaymentDataCollection_CardHolder," & _
				" tPPB_PAG.Req_PaymentDataCollection_PaymentPlan AS PAG_Req_PaymentDataCollection_PaymentPlan," & _
				" tPPB_PAG.Req_PaymentDataCollection_NumberOfPayments AS PAG_Req_PaymentDataCollection_NumberOfPayments," & _
				" tPPB_PAG.Resp_OrderData_BraspagOrderId AS PAG_Resp_OrderData_BraspagOrderId," & _
				" tPPB_PAG.Resp_Success AS PAG_Resp_Success," & _
				" tPPB_PAG.Resp_PaymentDataResponse_BraspagTransactionId AS PAG_Resp_PaymentDataResponse_BraspagTransactionId," & _
				" tPPB_PAG.Resp_PaymentDataResponse_AuthorizationCode AS PAG_Resp_PaymentDataResponse_AuthorizationCode," & _
				" tPPB_PAG.Resp_PaymentDataResponse_ProofOfSale AS PAG_Resp_PaymentDataResponse_ProofOfSale," & _
				" tPPB_PAG.Resp_PaymentDataResponse_ReturnMessage AS PAG_ReturnMessage," & _
				" tPPB_AF.Req_AFReq_BillToData_Street1," & _
				" tPPB_AF.Req_AFReq_BillToData_Street2," & _
				" tPPB_AF.Req_AFReq_BillToData_City," & _
				" tPPB_AF.Req_AFReq_BillToData_State," & _
				" tPPB_AF.Req_AFReq_BillToData_PostalCode," & _
				" tPPB_AF.Req_AFReq_BillToData_PhoneNumber," & _
				" tPPB_AF.Resp_AntiFraudTransactionId AS AF_Resp_AntiFraudTransactionId," & _
				" tPPB_AF.Resp_AFResp_AfsReply_AfsResult AS AF_AfsReply_AfsResult," & _
				" tPPB_AF.Resp_AFResp_AfsReply_ReasonCode AS AF_AfsReply_ReasonCode," & _
				" tPPB_AF.Resp_AFResp_AfsReply_AddressInfoCode AS AF_AfsReply_AddressInfoCode," & _
				" tPPB_AF.Resp_AFResp_AfsReply_AfsFactorCode AS AF_AfsReply_AfsFactorCode," & _
				" tPPB_AF.Resp_AFResp_AfsReply_HotlistInfoCode AS AF_AfsReply_HotlistInfoCode," & _
				" tPPB_AF.Resp_AFResp_AfsReply_IdentityInfoCode AS AF_AfsReply_IdentityInfoCode," & _
				" tPPB_AF.Resp_AFResp_AfsReply_InternetInfoCode AS AF_AfsReply_InternetInfoCode," & _
				" tPPB_AF.Resp_AFResp_AfsReply_PhoneInfoCode AS AF_AfsReply_PhoneInfoCode," & _
				" tPPB_AF.Resp_AFResp_AfsReply_SuspiciousInfoCode AS AF_AfsReply_SuspiciousInfoCode," & _
				" tPPB_AF.Resp_AFResp_AfsReply_VelocityInfoCode AS AF_AfsReply_VelocityInfoCode," & _
				" (SELECT TOP 1 ErrorCode + ' - ' + ErrorMessage FROM t_PEDIDO_PAGTO_BRASPAG_PAG_ERROR tPPB_PAG_E WHERE (tPPB_PAG_E.id_pedido_pagto_braspag_pag = tPPB_PAG.id) ORDER BY tPPB_PAG_E.id) AS PAG_ErrorMessage," & _
				" (SELECT TOP 1 ErrorCode + ' - ' + ErrorMessage FROM t_PEDIDO_PAGTO_BRASPAG_AF_ERROR tPPB_AF_E WHERE (tPPB_AF_E.id_pedido_pagto_braspag_af = tPPB_AF.id) ORDER BY tPPB_AF_E.id) AS AF_ErrorMessage" & _
			" FROM t_PEDIDO_PAGTO_BRASPAG" & _
				" INNER JOIN t_PEDIDO ON (t_PEDIDO_PAGTO_BRASPAG.pedido = t_PEDIDO.pedido)" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" INNER JOIN t_CLIENTE ON (t_PEDIDO_PAGTO_BRASPAG.id_cliente = t_CLIENTE.id)" & _
				" INNER JOIN t_PEDIDO_PAGTO_BRASPAG_PAG tPPB_PAG ON (t_PEDIDO_PAGTO_BRASPAG.id = tPPB_PAG.id_pedido_pagto_braspag)" & _
				" INNER JOIN t_PEDIDO_PAGTO_BRASPAG_AF tPPB_AF ON (t_PEDIDO_PAGTO_BRASPAG.id = tPPB_AF.id_pedido_pagto_braspag)" & _
			" WHERE" & _
				" (operacao = '" & OP_BRASPAG_OPERACAO__AF_PAG & "')" & _
				" AND (prim_AF_GlobalStatus = '" & s_prim_AF_GlobalStatus & "')"

	if s_prim_AF_GlobalStatus = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW then
		s_sql = s_sql & _
				" AND (" & _
					"(prim_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "')" & _
					" OR " & _
					"(prim_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA & "')" & _
					" OR " & _
					"(prim_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AGUARDANDO_RESPOSTA & "')" & _
					")"
		end if
	
	if c_resultado_transacao = COD_REL_BRASPAG_AF_REVIEW__REVISAO_MANUAL_PENDENTE then
		s_sql = s_sql & _
				" AND (AF_review_tratado_status = 0)"
	elseif c_resultado_transacao = COD_REL_BRASPAG_AF_REVIEW__REVISAO_MANUAL_TRATADA_ACCEPT then
		s_sql = s_sql & _
				" AND (AF_review_tratado_status = 1)" & _
				" AND (AF_review_tratado_decision = '" & BRASPAG_AF_DECISION__ACCEPT & "')"
	elseif c_resultado_transacao = COD_REL_BRASPAG_AF_REVIEW__REVISAO_MANUAL_TRATADA_REJECT then
		s_sql = s_sql & _
				" AND (AF_review_tratado_status = 1)" & _
				" AND (AF_review_tratado_decision = '" & BRASPAG_AF_DECISION__REJECT & "')"
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
						" pedido, id"
		else
			s_sql = s_sql & _
						" id"
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
		"		<td class='MTD tdStAF' style='vertical-align:bottom'><span class='Rc'>Status do</span><br /><span class='Rc'>Antifraude</span></td>" & chr(13) & _
		"		<td class='MTD tdDtHrStAF' style='vertical-align:bottom'><span class='Rc'>Data/Hora</span><br /><span class='Rc'>Status AF</span></td>" & chr(13) & _
		"		<td class='MTD tdCliente' style='vertical-align:bottom'><span class='R'>Cliente</span></td>" & chr(13) & _
		"		<td class='MTD tdRevisado' style='vertical-align:bottom'><span class='Rc'>Revisado</span></td>" & chr(13) & _
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
		s = monta_link_pedido(Trim("" & r("pedido")), usuario, "black")
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
		s = Trim("" & r("ult_PAG_GlobalStatus"))
		if s <> "" then s = BraspagPagadorDescricaoGlobalStatus(s)
		if s = "" then s = "&nbsp;"
		if Trim("" & r("ult_PAG_GlobalStatus")) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA then
			s_color = "green"
		elseif Trim("" & r("ult_PAG_GlobalStatus")) = BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA then
			s_color = "black"
		else
			s_color = "red"
			end if
		x = x & "		<td class='MTD tdStTransacao'><span id='spnStTransacao_" & Trim("" & r("id")) & "' class='Cnc' style='color:" & s_color & ";'>" & s & "</span></td>" & chr(13)

	'> DATA/HORA STATUS
		s = formata_data_hora(Trim("" & r("ult_PAG_atualizacao_data_hora")))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdDtHrStTransacao'><span id='spnDtHrStTransacao_" & Trim("" & r("id")) & "' class='Cnc'>" & s & "</span></td>" & chr(13)

	'> STATUS DO ANTIFRAUDE
		s = Trim("" & r("ult_AF_GlobalStatus"))
		if s <> "" then s = BraspagAntiFraudeDescricaoGlobalStatus(s)
		if s = "" then s = "&nbsp;"
		if Trim("" & r("ult_AF_GlobalStatus")) = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ACCEPT then
			s_color = "green"
		elseif Trim("" & r("ult_AF_GlobalStatus")) = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REJECT then
			s_color = "red"
		elseif Trim("" & r("ult_AF_GlobalStatus")) = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW then
			s_color = "#FF8C00"
		else
			s_color = "black"
			end if
		x = x & "		<td class='MTD tdStAF'><span id='spnStAF_" & Trim("" & r("id")) & "' class='Cnc' style='color:" & s_color & ";'>" & s & "</span></td>" & chr(13)

	'> DATA/HORA STATUS DO ANTIFRAUDE
		s = formata_data_hora(Trim("" & r("ult_AF_atualizacao_data_hora")))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdDtHrStAF'><span id='spnDtHrStAF_" & Trim("" & r("id")) & "' class='Cnc'>" & s & "</span></td>" & chr(13)

	'> CLIENTE
		s = cnpj_cpf_formata(Trim("" & r("cnpj_cpf"))) & " - " & Trim("" & r("cliente_nome"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdCliente'><span class='Cn'>" & s & "</span></td>" & chr(13)

	'> TRANSAÇÃO JÁ REVISADA?
		if r("AF_review_tratado_status") = 0 then
			s = "Não"
			s_color = "red"
			s_class = "Cnc"
		else
			s = "Sim"
			s_color = "green"
			s_class = "Cnc REVIEWED"
			end if
		
		x = x & _
			"		<td class='MTD tdRevisado'>" & chr(13) & _
						"<span id='spnAfRevisado_" & Trim("" & r("id")) & "' class='" & s_class & "' style='color:" & s_color & ";'>" & s & "</span>" & _
			"		</td>" & chr(13)
		
	'> BOTÃO P/ EXIBIR DEMAIS CAMPOS
		x = x & "		<td valign='bottom' class='notPrint' align='left'>" & _
							"&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(intQtdeTransacoes) & chr(34) & ");' title='exibe ou oculta os campos adicionais'><img src='../botao/view_bottom.png' border='0'></a>" & _
						"</td>" & chr(13)
		
		x = x & "	</tr>" & chr(13)
		
	'> OUTRAS INFORMAÇÕES
		x = x & "	<tr style='display:none;' id='TR_MORE_INFO_" & Cstr(intQtdeTransacoes) & "'>" & chr(13) & _
				"		<td class='ME MD' align='left'>&nbsp;</td>" & chr(13) & _
				"		<td colspan='11' class='MC MD' align='left'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<tr>" & chr(13) & _
				"					<td class='Rf tdWithPadding' align='left'>OUTRAS INFORMAÇÕES</td>" & chr(13) & _
				"				</tr>" & chr(13)
		
		x = x & _
			"				<tr>" & chr(13) & _
			"					<td align='left'>" & chr(13) & _
			"						<table width='100%' cellspacing='0' cellpadding='0' border=0>" & chr(13)
		
		blnTitularCartaoDivergente = False
		if Not is_nome_e_sobrenome_iguais(Trim("" & r("cliente_nome")), Trim("" & r("Req_PaymentDataCollection_CardHolder"))) then blnTitularCartaoDivergente = True
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
		s_aux = Trim("" & r("Req_PaymentDataCollection_CardHolder"))
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
		s_end_fatura = Trim("" & r("Req_AFReq_BillToData_Street1"))
		if s_end_fatura <> "" then
			s_aux = Trim("" & r("Req_AFReq_BillToData_Street2"))
			if s_aux <> "" then s_end_fatura = s_end_fatura & " " & s_aux
			s_aux = Trim("" & r("Req_AFReq_BillToData_City"))
			if s_aux <> "" then s_end_fatura = s_end_fatura & " - " & s_aux
			s_aux = Trim("" & r("Req_AFReq_BillToData_State"))
			if s_aux <> "" then s_end_fatura = s_end_fatura & " - " & s_aux
			s_aux = Trim("" & r("Req_AFReq_BillToData_PostalCode"))
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
		s_aux = Trim("" & r("Req_AFReq_BillToData_PhoneNumber"))
		if s_aux <> "" then
			s_tel_pais = Mid(s_aux, 1, 2)
			s_tel_ddd = Mid(s_aux, 3, 2)
			s_tel_numero = telefone_formata(Mid(s_aux, 5))
			s_aux = "+" & s_tel_pais & " (" & s_tel_ddd & ") " & s_tel_numero
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
												primeiroNaoVazio(Array(BraspagDescricaoParcelamento(Trim("" & r("PAG_Req_PaymentDataCollection_PaymentPlan")), Trim("" & r("PAG_Req_PaymentDataCollection_NumberOfPayments")), r("valor_transacao")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>NSU da Transação:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												Trim("" & r("id")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
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
												primeiroNaoVazio(Array(Trim("" & r("PAG_ReturnMessage")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Mensagem de Erro:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("PAG_ErrorMessage")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco' rowspan='14' align='center' valign='middle'><img src='../Imagem/braspag-AF-vert.png' width='24' height='38' /></td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Revisão Manual Realizada em:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span id='spnDtHrRevisaoManual_" & Trim("" & r("id")) & "'>" & formata_data_hora(r("AF_review_tratado_data_hora")) & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Revisão Manual Realizada por:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span id='spnUsuarioRevisaoManual_" & Trim("" & r("id")) & "'>" & Trim("" & r("AF_review_tratado_usuario")) & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		if Trim("" & r("AF_review_tratado_decision")) = BRASPAG_AF_DECISION__ACCEPT then
			s_color = "green"
		elseif Trim("" & r("AF_review_tratado_decision")) = BRASPAG_AF_DECISION__REJECT then
			s_color ="red"
		else
			s_color ="black"
			end if
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Decisão da Revisão Manual:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo' style='color:" & s_color & ";'>" & chr(13) & _
												"<span id='spnDecisaoRevisaoManual_" & Trim("" & r("id")) & "'>" & Trim("" & r("AF_review_tratado_decision")) & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Score:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span>" & Trim("" & r("AF_AfsReply_AfsResult")) & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	ReasonCode
		s = Trim("" & r("AF_AfsReply_ReasonCode"))
		if s <> "" then s = s & " - " & BraspagAfDescricaoReasonCode(s)
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Código do Motivo:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span>" & s & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	AddressInfoCode
		s_dados = ""
		s = Trim("" & r("AF_AfsReply_AddressInfoCode"))
		v = Split(s, "^")
		for i=LBound(v) to UBound(v)
			if Trim("" & v(i)) <> "" then
				s = Trim("" & v(i))
				s = s & " - " & BraspagAfDescricaoAddressInfoCode(s)
				if s_dados <> "" then s_dados = s_dados & "<br />"
				s_dados = s_dados & s
				end if
			next
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Inf. Endereço:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span>" & s_dados & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	AfsFactorCode
		s_dados = ""
		s = Trim("" & r("AF_AfsReply_AfsFactorCode"))
		v = Split(s, "^")
		for i=LBound(v) to UBound(v)
			if Trim("" & v(i)) <> "" then
				s = Trim("" & v(i))
				s = s & " - " & BraspagAfDescricaoAfsFactorCode(s)
				if s_dados <> "" then s_dados = s_dados & "<br />"
				s_dados = s_dados & s
				end if
			next
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Fatores de Risco:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span>" & s_dados & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	HotlistInfoCode
		s_dados = ""
		s = Trim("" & r("AF_AfsReply_HotlistInfoCode"))
		v = Split(s, "^")
		for i=LBound(v) to UBound(v)
			if Trim("" & v(i)) <> "" then
				s = Trim("" & v(i))
				s = s & " - " & BraspagAfDescricaoHotlistInfoCode(s)
				if s_dados <> "" then s_dados = s_dados & "<br />"
				s_dados = s_dados & s
				end if
			next
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Inf. Listas de Clientes:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span>" & s_dados & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	IdentityInfoCode
		s_dados = ""
		s = Trim("" & r("AF_AfsReply_IdentityInfoCode"))
		v = Split(s, "^")
		for i=LBound(v) to UBound(v)
			if Trim("" & v(i)) <> "" then
				s = Trim("" & v(i))
				s = s & " - " & BraspagAfDescricaoIdentityInfoCode(s)
				if s_dados <> "" then s_dados = s_dados & "<br />"
				s_dados = s_dados & s
				end if
			next
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Mudanças de Identidade Excessivas:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span>" & s_dados & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	InternetInfoCode
		s_dados = ""
		s = Trim("" & r("AF_AfsReply_InternetInfoCode"))
		v = Split(s, "^")
		for i=LBound(v) to UBound(v)
			if Trim("" & v(i)) <> "" then
				s = Trim("" & v(i))
				s = s & " - " & BraspagAfDescricaoInternetInfoCode(s)
				if s_dados <> "" then s_dados = s_dados & "<br />"
				s_dados = s_dados & s
				end if
			next
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Inf. Internet:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span>" & s_dados & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	PhoneInfoCode
		s_dados = ""
		s = Trim("" & r("AF_AfsReply_PhoneInfoCode"))
		v = Split(s, "^")
		for i=LBound(v) to UBound(v)
			if Trim("" & v(i)) <> "" then
				s = Trim("" & v(i))
				s = s & " - " & BraspagAfDescricaoPhoneInfoCode(s)
				if s_dados <> "" then s_dados = s_dados & "<br />"
				s_dados = s_dados & s
				end if
			next
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Inf. Telefone:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span>" & s_dados & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	SuspiciousInfoCode
		s_dados = ""
		s = Trim("" & r("AF_AfsReply_SuspiciousInfoCode"))
		v = Split(s, "^")
		for i=LBound(v) to UBound(v)
			if Trim("" & v(i)) <> "" then
				s = Trim("" & v(i))
				s = s & " - " & BraspagAfDescricaoSuspiciousInfoCode(s)
				if s_dados <> "" then s_dados = s_dados & "<br />"
				s_dados = s_dados & s
				end if
			next
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Inf. Dados Suspeitos:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span>" & s_dados & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	VelocityInfoCode
		s_dados = ""
		s = Trim("" & r("AF_AfsReply_VelocityInfoCode"))
		v = Split(s, "^")
		for i=LBound(v) to UBound(v)
			if Trim("" & v(i)) <> "" then
				s = Trim("" & v(i))
				s = s & " - " & BraspagAfDescricaoVelocityInfoCode(s)
				if s_dados <> "" then s_dados = s_dados & "<br />"
				s_dados = s_dados & s
				end if
			next
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Inf. Velocidade Global:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												"<span>" & s_dados & "</span>" & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	Mensagem de Erro
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Mensagem de Erro:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("AF_ErrorMessage")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco'>&nbsp;</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Comentário:<br />" & _
												"<span class='lblTitTamRestante'>(tamanho restante:&nbsp;</span>" & _
												"<span class='lblTamRestante' name='lblTamRestante_" & Trim("" & r("id")) & "' id='lblTamRestante_" & Trim("" & r("id")) & "'></span>" & _
												"<span class='lblTitTamRestante'>)</span>" & _
											"</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
											"	<textarea maxlength='400' style='width:100%;height:100px;border:0px;'" & _
													" name='c_AF_review_tratado_comentario_" & Trim("" & r("id")) & "'" & _
													" id='c_AF_review_tratado_comentario_" & Trim("" & r("id")) & "'"
			if r("AF_review_tratado_status") <> 0 then x = x & " readonly"
			x = x & _
													">" & Trim("" & r("AF_review_tratado_comentario")) & "</textarea>"  & chr(13) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	ANÁLISE DO ENDEREÇO
		x2 = "<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
			 "	<tr>" & chr(13)
		
		if CLng(r("analise_endereco_tratar_status")) = 0 then
			x2 = x2 & _
				"<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"	<tr>" & chr(13) & _
				"		<td align='left'><span class='Cn' style='text-align:left;'>(Nenhum)</span></td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"</table>" & chr(13)
		else
			x2 = x2 & _
				"		<td align='left'>" & _
				"<a id='hrefTitAnEnd_" & Trim("" & r("id")) & "' href='javascript:exibeOcultaTodosInfoAnEnd(" & Trim("" & r("id")) & ");' title='clique para exibir mais detalhes'>" & _
				"<span id='spanTitAnEnd_" & Trim("" & r("id")) & "' class='Cn TIT_INFO_AN_END_BLOCO'>Exibir/ocultar todos</span>" & _
				"&nbsp;<img id='imgPlusMinusTitAnEnd_" & Trim("" & r("id")) & "' style='vertical-align:bottom;margin-bottom:2px;' src='../imagem/plus.gif' />" & _
				"</a>" & _
				"</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr>" & chr(13) & _
				"		<td align='left'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13)
			
			strInfoAnEnd = ""
			intQtdePedidoAnEndereco = 0
			intQtdeLinhasPedidoAnEndereco = 0
			intQtdeTotalPedidosAnEndereco = 0
		'	VERIFICA SE HÁ COINCIDÊNCIA C/ ENDEREÇO DO PARCEIRO
			blnAnEnderecoUsaEndParceiro = False
			
			s_sql = "SELECT" & _
						" tP.indicador," & _
						" tOI.razao_social_nome_iniciais_em_maiusculas AS nome_indicador," & _
						" tOI.cnpj_cpf AS cnpj_cpf_indicador," & _
						" tPAEC.*" & _
					" FROM t_PEDIDO_ANALISE_ENDERECO tPAE" & _
						" INNER JOIN t_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO tPAEC ON (tPAE.id = tPAEC.id_pedido_analise_endereco)" & _
						" LEFT JOIN t_PEDIDO tP ON (tPAE.pedido = tP.pedido)" & _
						" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR tOI ON (tP.indicador = tOI.apelido)" & _
					" WHERE" & _
						" (tPAE.pedido = '" & Trim("" & r("pedido")) & "')" & _
						" AND (tPAEC.tipo_endereco = '" & COD_PEDIDO_AN_ENDERECO__END_PARCEIRO & "')" & _
					" ORDER BY" & _
						" tPAE.id," & _
						" tPAEC.id"
			if r2.State <> 0 then r2.Close
			r2.open s_sql, cn
			do while Not r2.Eof
				blnAnEnderecoUsaEndParceiro = True
				intQtdeTotalPedidosAnEndereco = intQtdeTotalPedidosAnEndereco + 1
				if intQtdeTotalPedidosAnEndereco > MAX_AN_ENDERECO_QTDE_PEDIDOS_EXIBICAO then exit do
				intResto = intQtdePedidoAnEndereco Mod MAX_PEDIDOS_POR_LINHA_ANALISE_ENDERECO
				if (intQtdePedidoAnEndereco = 0) Or (intResto = 0) then
					intQtdePedidoAnEndereco = 0
					if intQtdeLinhasPedidoAnEndereco > 0 then
						x2 = x2 & "				</tr>" & chr(13)
						end if
					x2 = x2 & "				<tr>" & chr(13)
					intQtdeLinhasPedidoAnEndereco = intQtdeLinhasPedidoAnEndereco + 1
					end if
				
			'	LEMBRANDO QUE UM MESMO PEDIDO PODE GERAR MAIS DE UMA TRANSAÇÃO DE PAGAMENTO, PORTANTO, OS MESMOS DADOS DE ANÁLISE DE ENDEREÇO
			'	PODEM SER EXIBIDOS MAIS DE UMA VEZ.
			'	POR ESSE MOTIVO, O IDENTIFICADOR DOS CAMPOS DEVE INCLUIR O IDENTIFICADOR DA TRANSAÇÃO PARA QUE O ID DO ELEMENTO SEJA ÚNICO
				x2 = x2 & _
					"					<td align='left' valign='bottom'>" & chr(13) & _
					"<a id='hrefPedAnEnd_" & Trim("" & r("id")) & "_" & Trim("" & r2("id")) & "' class='hrefAnEndBloco_" & Trim("" & r("id")) & "' href='javascript:exibeOcultaInfoAnEnd(" & chr(34) & Trim("" & r("id")) & "_" & Trim("" & r2("id")) & chr(34) & ");' title='clique para exibir mais detalhes'>" & _
						"<span class='C' id='spanPedidoAnEnd_" & Trim("" & r("id")) & "_" & Trim("" & r2("id")) & "'>Indicador</span>" & _
						"&nbsp;<img id='imgPlusMinusPedAnEnd_" & Trim("" & r("id")) & "_" & Trim("" & r2("id")) & "' class='imgPlusMinusAnEndBloco_" & Trim("" & r("id")) & "' style='vertical-align:bottom;margin-bottom:0px;' src='../imagem/plus.gif' />" & _
					"</a>" & _
					"					</td>" & chr(13)
				
				strInfoAnEnd = strInfoAnEnd & _
					"	<tr id='TR_INFO_AN_END_LN1_" & Trim("" & r("id")) & "_" & Trim("" & r2("id")) & "' class='TR_INFO_AN_END TR_INFO_AN_END_BLOCO_" & Trim("" & r("id")) & "'>" & chr(13) & _
					"		<td align='left' valign='bottom' class='MC tdAnEndPed'>" & chr(13) & _
							"<a href='javascript:ocultaInfoAnEnd(" & chr(34) & Trim("" & r("id")) & "_" & Trim("" & r2("id")) & chr(34) & ");' title='clique para ocultar os detalhes'>" & _
								"<img id='imgMinusPedAnEnd_" & Trim("" & r("id")) & "_" & Trim("" & r2("id")) & "' style='vertical-align:bottom;margin-left:2px;margin-bottom:1px;' src='../imagem/minus.gif' />" & chr(13) & _
							"</a>" & _
								"<span class='Cn'>Indicador</span>" & _
					"		</td>" & chr(13) & _
					"		<td align='left' class='MC'>" & chr(13) & _
								"<span class='Cn'>" & _
								Trim("" & r2("indicador")) & " - " & Trim("" & r2("nome_indicador")) & " ("
				
				s_aux = retorna_so_digitos(Trim("" & r2("cnpj_cpf_indicador")))
				if Len(s_aux) = 11 then
					strInfoAnEnd = strInfoAnEnd & "CPF: " & s_aux & ")"
				else
					strInfoAnEnd = strInfoAnEnd & "CNPJ: " & s_aux & ")"
					end if
				
				strInfoAnEnd = strInfoAnEnd & _
								"</span>" & _
					"		</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<tr id='TR_INFO_AN_END_LN2_" & Trim("" & r("id")) & "_" & Trim("" & r2("id")) & "' class='TR_INFO_AN_END TR_INFO_AN_END_BLOCO_" & Trim("" & r("id")) & "'>" & chr(13) & _
					"		<td align='left'>&nbsp;</td>" & chr(13) & _
					"		<td align='left'>" & chr(13)
				
				s_aux = "End. do Indicador: "
				s = formata_endereco(iniciais_em_maiusculas(Trim("" & r2("endereco_logradouro"))), Trim("" & r2("endereco_numero")), Trim("" & r2("endereco_complemento")), iniciais_em_maiusculas(Trim("" & r2("endereco_bairro"))), iniciais_em_maiusculas(Trim("" & r2("endereco_cidade"))), Ucase(Trim("" & r2("endereco_uf"))), retorna_so_digitos(Trim("" & r2("endereco_cep"))))
				strInfoAnEnd = strInfoAnEnd & _
								"<span class='Cni'>" & _
								s_aux & _
								"</span>" & _
								"<span class='Cn'>" & _
								s & _
								"</span>"
								
				strInfoAnEnd = strInfoAnEnd & _
					"		</td>" & chr(13) & _
					"	</tr>" & chr(13)
				
				intQtdePedidoAnEndereco = intQtdePedidoAnEndereco + 1
				
				r2.MoveNext
				loop
			
		'	VERIFICA SE HÁ COINCIDÊNCIA C/ ENDEREÇO DE OUTROS CLIENTES
			s_sql = "SELECT " & _
						"*" & _
					" FROM t_PEDIDO_ANALISE_ENDERECO" & _
					" WHERE" & _
						" (pedido = '" & Trim("" & r("pedido")) & "')" & _
					" ORDER BY" & _
						" id"
			if r2.State <> 0 then r2.Close
			r2.open s_sql, cn
			if r2.Eof then
				if Not blnAnEnderecoUsaEndParceiro then
					x2 = "				<tr>" & chr(13) & _
						"					<td align='left'>" & chr(13) & _
											"&nbsp;" & _
						"					</td>" & chr(13) & _
						"				</tr>" & chr(13)
					end if
			else
				do while Not r2.Eof
					if intQtdeTotalPedidosAnEndereco > MAX_AN_ENDERECO_QTDE_PEDIDOS_EXIBICAO then exit do
					s_sql = "SELECT" & _
								" tPAEC.*," & _
								" tP.st_entrega," & _
								" tC.nome_iniciais_em_maiusculas," & _
								" tC.cnpj_cpf" & _
							" FROM t_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO tPAEC" & _
								" INNER JOIN t_PEDIDO tP ON (tPAEC.pedido=tP.pedido)" & _
								" LEFT JOIN t_CLIENTE tC ON (tPAEC.id_cliente=tC.id)" & _
							" WHERE" & _
								" (tPAEC.id_pedido_analise_endereco = " & Trim("" & r2("id")) & ")" & _
								" AND (tPAEC.tipo_endereco <> '" & COD_PEDIDO_AN_ENDERECO__END_PARCEIRO & "')" & _
							" ORDER BY" & _
								" tPAEC.id"
					if r3.State <> 0 then r3.Close
					r3.open s_sql, cn
					do while Not r3.Eof
						intQtdeTotalPedidosAnEndereco = intQtdeTotalPedidosAnEndereco + 1
						if intQtdeTotalPedidosAnEndereco > MAX_AN_ENDERECO_QTDE_PEDIDOS_EXIBICAO then exit do
						intResto = intQtdePedidoAnEndereco Mod MAX_PEDIDOS_POR_LINHA_ANALISE_ENDERECO
						if (intQtdePedidoAnEndereco = 0) Or (intResto = 0) then
							intQtdePedidoAnEndereco = 0
							if intQtdeLinhasPedidoAnEndereco > 0 then
								x2 = x2 & "				</tr>" & chr(13)
								end if
							x2 = x2 & "				<tr>" & chr(13)
							intQtdeLinhasPedidoAnEndereco = intQtdeLinhasPedidoAnEndereco + 1
							end if
						
						if Trim("" & r3("st_entrega")) = ST_ENTREGA_ENTREGUE then
							strColor = "green"
						elseif Trim("" & r3("st_entrega")) = ST_ENTREGA_CANCELADO then
							strColor = "red"
						else
							strColor = "black"
							end if
						x2 = x2 & _
							"					<td class='C' align='left' valign='bottom'>" & chr(13) & _
							monta_link_pedido(Trim("" & r3("pedido")), usuario, strColor) & _
							"&nbsp;<a id='hrefPedAnEnd_" & Trim("" & r("id")) & "_" & Trim("" & r3("id")) & "' class='hrefAnEndBloco_" & Trim("" & r("id")) & "' href='javascript:exibeOcultaInfoAnEnd(" & chr(34) & Trim("" & r("id")) & "_" & Trim("" & r3("id")) & chr(34) & ");' title='clique para exibir mais detalhes'>" & _
								"<img id='imgPlusMinusPedAnEnd_" & Trim("" & r("id")) & "_" & Trim("" & r3("id")) & "' class='imgPlusMinusAnEndBloco_" & Trim("" & r("id")) & "' style='vertical-align:bottom;margin-bottom:0px;' src='../imagem/plus.gif' />" & _
							"</a>" & _
							"					</td>" & chr(13)
						
						strInfoAnEnd = strInfoAnEnd & _
							"	<tr id='TR_INFO_AN_END_LN1_" & Trim("" & r("id")) & "_" & Trim("" & r3("id")) & "' class='TR_INFO_AN_END TR_INFO_AN_END_BLOCO_" & Trim("" & r("id")) & "'>" & chr(13) & _
							"		<td align='left' valign='bottom' class='MC tdAnEndPed'>" & chr(13) & _
									"<a href='javascript:ocultaInfoAnEnd(" & chr(34) & Trim("" & r("id")) & "_" & Trim("" & r3("id")) & chr(34) & ");' title='clique para ocultar os detalhes'>" & _
										"<img id='imgMinusPedAnEnd_" & Trim("" & r("id")) & "_" & Trim("" & r3("id")) & "' style='vertical-align:bottom;margin-left:2px;margin-bottom:1px;' src='../imagem/minus.gif' />" & chr(13) & _
									"</a>" & _
										"<span class='Cn spnLnkPedAnEnd' style='color:" & strColor & ";' onclick='fPEDConsulta(" & chr(34) & Trim("" & r3("pedido")) & chr(34) & ");'>" & Trim("" & r3("pedido")) & "</span>" & _
							"		</td>" & chr(13) & _
							"		<td align='left' class='MC'>" & chr(13) & _
										"<span class='Cn'>" & _
										Trim("" & r3("nome_iniciais_em_maiusculas")) & " ("
						
						s_aux = retorna_so_digitos(Trim("" & r3("cnpj_cpf")))
						if Len(s_aux) = 11 then
							strInfoAnEnd = strInfoAnEnd & "CPF: " & s_aux & ")"
						else
							strInfoAnEnd = strInfoAnEnd & "CNPJ: " & s_aux & ")"
							end if
						
						strInfoAnEnd = strInfoAnEnd & _
										"</span>" & _
							"		</td>" & chr(13) & _
							"	</tr>" & chr(13) & _
							"	<tr id='TR_INFO_AN_END_LN2_" & Trim("" & r("id")) & "_" & Trim("" & r3("id")) & "' class='TR_INFO_AN_END TR_INFO_AN_END_BLOCO_" & Trim("" & r("id")) & "'>" & chr(13) & _
							"		<td align='left'>&nbsp;</td>" & chr(13) & _
							"		<td align='left'>" & chr(13)
						
						if Trim("" & r3("tipo_endereco")) = COD_PEDIDO_AN_ENDERECO__END_ENTREGA then
							s_aux = "End. Entrega: "
						else
							s_aux = "End. Cadastro: "
							end if
						s = formata_endereco(iniciais_em_maiusculas(Trim("" & r3("endereco_logradouro"))), Trim("" & r3("endereco_numero")), Trim("" & r3("endereco_complemento")), iniciais_em_maiusculas(Trim("" & r3("endereco_bairro"))), iniciais_em_maiusculas(Trim("" & r3("endereco_cidade"))), Ucase(Trim("" & r3("endereco_uf"))), retorna_so_digitos(Trim("" & r3("endereco_cep"))))
						strInfoAnEnd = strInfoAnEnd & _
										"<span class='Cni'>" & _
										s_aux & _
										"</span>" & _
										"<span class='Cn'>" & _
										s & _
										"</span>"
										
						strInfoAnEnd = strInfoAnEnd & _
							"		</td>" & chr(13) & _
							"	</tr>" & chr(13)
						
						intQtdePedidoAnEndereco = intQtdePedidoAnEndereco + 1
						r3.MoveNext
						loop
					
					if intQtdeTotalPedidosAnEndereco > MAX_AN_ENDERECO_QTDE_PEDIDOS_EXIBICAO then exit do
					r2.MoveNext
					loop
				end if
			
			x2 = x2 & _
				"				</tr>" & chr(13)
			
			if strInfoAnEnd <> "" then
				x2 = x2 & _
					"	<tr>" & chr(13) & _
					"		<td colspan='" & MAX_PEDIDOS_POR_LINHA_ANALISE_ENDERECO & "' align='left'>" & chr(13) & _
					"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
								strInfoAnEnd & _
					"			</table>" & chr(13) & _
					"		</td>" & chr(13) & _
					"	</tr>" & chr(13)
				end if
			
			x2 = x2 & _
				"			</table>" & chr(13) & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"</table>" & chr(13)
			end if ' if CLng(r("analise_endereco_tratar_status")) <> 0
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco'>&nbsp;</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Análise do Endereço:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
			x2 & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	PEDIDOS ANTERIORES
		strPedidosAnteriores = "<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
								"	<tr>" & chr(13)
		s_sql = "SELECT" & _
					" pedido," & _
					" st_entrega" & _
				" FROM t_PEDIDO" & _
				" WHERE" & _
					" (id_cliente = '" & Trim("" & r("id_cliente")) & "')" & _
					" AND (pedido <> '" & Trim("" & r("pedido")) & "')" & _
					" AND (data_hora < CONVERT(datetime, '" & formata_data_com_separador_yyyymmdd(r("data_pedido"),"-") & " " & formata_hhnnss_para_hh_nn_ss(Trim("" & r("hora_pedido"))) & "', 120))" & _
				" ORDER BY" & _
					" data_hora DESC," & _
					" pedido DESC"
		if r2.State <> 0 then r2.Close
		r2.open s_sql, cn
		if r2.Eof then
			strPedidosAnteriores = strPedidosAnteriores & _
									"		<td align='left'><span class='Cn' style='text-align:left;'>(Nenhum)</span></td>" & chr(13)
		else
			n_ped_linha = 0
			do while Not r2.Eof
				n_ped_linha = n_ped_linha+1
				if n_ped_linha > PEDIDOS_POR_LINHA then
					n_ped_linha = 1
					strPedidosAnteriores = strPedidosAnteriores & _
											"	</tr>" & chr(13) & "	<tr>" & chr(13)
					end if
				if Trim("" & r2("st_entrega")) = ST_ENTREGA_ENTREGUE then
					strColor = "green"
				elseif Trim("" & r2("st_entrega")) = ST_ENTREGA_CANCELADO then
					strColor = "red"
				else
					strColor = "black"
					end if
				strPedidosAnteriores = strPedidosAnteriores & _
										"		<td width='12.5%' class='L' style='text-align:left;color:black;' align='left'>" & monta_link_pedido(Trim("" & r2("pedido")), usuario, strColor) & "</td>" & chr(13)
				r2.MoveNext
				loop
			
			if (n_ped_linha Mod PEDIDOS_POR_LINHA)<> 0 then
				for i = ((n_ped_linha Mod PEDIDOS_POR_LINHA)+1) to PEDIDOS_POR_LINHA
					strPedidosAnteriores = strPedidosAnteriores & _
											"		<td align='left'>&nbsp;</td>" & chr(13)
					next
				end if
			end if
		strPedidosAnteriores = strPedidosAnteriores & _
								"	</tr>" & chr(13) & _
								"</table>" & chr(13)
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco'>&nbsp;</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Pedidos Anteriores:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
			strPedidosAnteriores & _
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
			"											<td style='width:180px;' align='center'><a href='javascript:braspagAfConsultaStatus(" & Trim("" & r("id")) & "," & Trim("" & r("id_pedido_pagto_braspag_af")) & "," & Trim("" & r("id_pedido_pagto_braspag_pag")) & ",""" & usuario & """);'><p class='Button BtnBraspag'  style='color:black;'>&nbsp;Consulta Status Atualizado&nbsp;</p></a></td>" & chr(13) & _
			"											<td align='left' style='width:15px;'>&nbsp;</td>" & chr(13) & _
			"											<td style='width:120px;' align='center'><a id='lnkAccept_" & Trim("" & r("id")) & "' href='javascript:braspagAfAprovar(" & Trim("" & r("id")) & "," & Trim("" & r("id_pedido_pagto_braspag_af")) & "," & Trim("" & r("id_pedido_pagto_braspag_pag")) & ",""" & Trim("" & r("pedido")) & """,""" & usuario & """, $(""#c_AF_review_tratado_comentario_" & Trim("" & r("id")) & """).val());'><p class='Button BtnBraspag' style='color:green;'>&nbsp;Aprovar&nbsp;</p></a></td>" & chr(13) & _
			"											<td align='left' style='width:15px;'>&nbsp;</td>" & chr(13) & _
			"											<td style='width:120px;' align='center'><a id='lnkReject_" & Trim("" & r("id")) & "' href='javascript:braspagAfRejeitar(" & Trim("" & r("id")) & "," & Trim("" & r("id_pedido_pagto_braspag_af")) & "," & Trim("" & r("id_pedido_pagto_braspag_pag")) & ",""" & Trim("" & r("pedido")) & """,""" & usuario & """, $(""#c_AF_review_tratado_comentario_" & Trim("" & r("id")) & """).val());'><p class='Button BtnBraspag' style='color:red;'>&nbsp;Rejeitar&nbsp;</p></a></td>" & chr(13)
		
		if operacao_permitida(OP_CEN_EDITA_ANALISE_CREDITO, s_lista_operacoes_permitidas) Or _
			operacao_permitida(OP_CEN_REL_ANALISE_CREDITO, s_lista_operacoes_permitidas) then
			if CStr(r("analise_credito")) = CStr(COD_AN_CREDITO_OK) then
				x = x & _
					"											<td align='left' style='width:15px;'>&nbsp;</td>" & chr(13) & _
					"											<td style='width:120px;' align='center'><a class='LnkCredOk LnkCredOkDisabled' id='lnkAnaliseCreditoOk_" & Trim("" & r("id")) & "' href='javascript:gravarAnaliseCreditoOk(""" & Trim("" & r("pedido")) & """, " & Trim("" & r("id")) & ",""" & usuario & """);'><p class='Button BtnBraspag'>&nbsp;Gravar Crédito Ok&nbsp;</p></a></td>" & chr(13)
			else
				x = x & _
					"											<td align='left' style='width:15px;'>&nbsp;</td>" & chr(13) & _
					"											<td style='width:120px;' align='center'><a class='LnkCredOk LnkCredOkEnabled' id='lnkAnaliseCreditoOk_" & Trim("" & r("id")) & "' href='javascript:gravarAnaliseCreditoOk(""" & Trim("" & r("pedido")) & """, " & Trim("" & r("id")) & ",""" & usuario & """);'><p class='Button BtnBraspag'>&nbsp;Gravar Crédito Ok&nbsp;</p></a></td>" & chr(13)
				end if
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
				"		<td colspan='7' class='MC MD' align='left'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr>" & chr(13) & _
				"		<td colspan='12' class='MC' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr nowrap style='background:#F0FFF0;'>" & chr(13) & _
				"		<td colspan='12' class='MT' align='left'><span class='C'>TOTAL: &nbsp; " & Cstr(intQtdeTransacoes) & iif((intQtdeTransacoes=1), " transação", " transações") & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if
	
'	MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeTransacoes = 0 then
		x = cab_table & cab
		x = x & "	<tr nowrap>" & chr(13) & _
				"		<td class='MT ALERTA' align='center' colspan='12'><span class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

'	FECHA TABELA
	x = x & "</table>" & chr(13)
	
	x = x & "<input type='hidden' name='c_qtde_transacoes' id='c_qtde_transacoes' value='" & Cstr(intQtdeTransacoes) & "'>" & chr(13)

	Response.write x

	qtde_transacoes = intQtdeTransacoes

	if r.State <> 0 then r.Close
	set r=nothing
	
	if r2.State <> 0 then r2.Close
	set r2=nothing
	
	if r3.State <> 0 then r3.Close
	set r3=nothing
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

		$("textarea").focus(function() {
			if ($(this).is("[readonly]")) {
				$(this).css('background-color', '#FF6347');
			}
			else {
				$(this).css('background-color', '#98FB98');
			}
		});

		$("textarea").blur(function() {
			$(this).css('background-color', '#FFFFFF');
			$(this).val($.trim($(this).val()));
		});

		$("span.lblTamRestante").each(function() {
			var s_id = $(this).attr("id");
			s_id = "#" + s_id.replace("lblTamRestante_", "c_AF_review_tratado_comentario_");
			var c = $(s_id);
			var n = MAX_TAM_T_PEDIDO_PAGTO_BRASPAG__AF_REVIEW_COMENTARIO - $(c).val().length;
			$(this).html(n.toString());
		});

		$("textarea").bind("input keyup paste", function() {
			var s_id = $(this).attr("id");
			s_id = "#" + s_id.replace("c_AF_review_tratado_comentario_", "lblTamRestante_");
			var c = $(s_id);
			var n = MAX_TAM_T_PEDIDO_PAGTO_BRASPAG__AF_REVIEW_COMENTARIO - $(this).val().length;
			$(c).html(n.toString());
		});

		$(".LnkCredOk").click(function(e) {
			if ($(this).hasClass("LnkCredOkDisabled")) {
				e.preventDefault();
			}
		});

		$(".TR_INFO_AN_END").hide().addClass("TR_INFO_AN_END_HIDDEN");
		$(".TIT_INFO_AN_END_BLOCO").addClass("TR_INFO_AN_END_HIDDEN");
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
var MAX_TAM_T_PEDIDO_PAGTO_BRASPAG__AF_REVIEW_COMENTARIO=<%=MAX_TAM_T_PEDIDO_PAGTO_BRASPAG__AF_REVIEW_COMENTARIO%>;
var BRASPAG_AF_DECISION__ACCEPT="<%=BRASPAG_AF_DECISION__ACCEPT%>";
var BRASPAG_AF_DECISION__REJECT="<%=BRASPAG_AF_DECISION__REJECT%>";
var windowScrollTopAnterior;

window.status = 'Aguarde, executando a consulta ...';

function braspagAfAprovar(_id_pedido_pagto_braspag, _id_pedido_pagto_braspag_af, _id_pedido_pagto_braspag_pag, _pedido, _usuario, _af_comentario){
	braspagAfExecutaUpdateStatus(BRASPAG_AF_DECISION__ACCEPT, _id_pedido_pagto_braspag, _id_pedido_pagto_braspag_af, _id_pedido_pagto_braspag_pag, _pedido, _usuario, _af_comentario);
}

function braspagAfRejeitar(_id_pedido_pagto_braspag, _id_pedido_pagto_braspag_af, _id_pedido_pagto_braspag_pag, _pedido, _usuario, _af_comentario){
	if (!confirm("Rejeitar a transação do pedido " + _pedido + "?\nIsso causará o cancelamento/estorno automático da transação!")) return;
	braspagAfExecutaUpdateStatus(BRASPAG_AF_DECISION__REJECT, _id_pedido_pagto_braspag, _id_pedido_pagto_braspag_af, _id_pedido_pagto_braspag_pag, _pedido, _usuario, _af_comentario);
}

function braspagAfExecutaUpdateStatus(_af_decision, _id_pedido_pagto_braspag, _id_pedido_pagto_braspag_af, _id_pedido_pagto_braspag_pag, _pedido, _usuario, _af_comentario){
var s, s_id, s_color;
var strMsg = "";
var strMsgErro = "";
var blnShowJsonFields = false;

	s_id = "#spnAfRevisado_" + _id_pedido_pagto_braspag;
	if ( ($(s_id).text().toUpperCase() == "SIM") || ($(s_id).hasClass("REVIEWED")) ) {
		if (!confirm("Esta transação já consta como tratada!\nTem certeza de que deseja enviar novamente a requisição da revisão manual?")) return;
		}

var jqxhr = $.ajax({
	url: "../Global/AjaxBraspagAfUpdateStatus.asp",
	type: "POST",
	dataType: 'json',
	data: {
		af_decision : _af_decision,
		id_pedido_pagto_braspag : _id_pedido_pagto_braspag.toString(),
		id_pedido_pagto_braspag_af : _id_pedido_pagto_braspag_af.toString(),
		id_pedido_pagto_braspag_pag : _id_pedido_pagto_braspag_pag.toString(),
		usuario : _usuario,
		af_comentario : _af_comentario
		}
	})
	.done(function(response){
	// Dados de resposta do Antifraude
		if ((response.AF.AF_ErrorCode.length > 0)||(response.AF.AF_ErrorMessage.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.AF.AF_ErrorCode + ": " + response.AF.AF_ErrorMessage;
		}
		if ((response.AF.AF_faultcode.length > 0)||(response.AF.AF_faultstring.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.AF.AF_faultcode + ": " + response.AF.AF_faultstring;
		}
		if (strMsgErro.length > 0) {
			strMsgErro = "Falha no processamento da requisição!!\nIMPORTANTE: consulte o status atualizado para verificar a necessidade de refazer esta operação ou de realizar o cancelamento/estorno manual da transação!!\n\n" + strMsgErro;
			alert(strMsgErro);
			return;
		}
		
		if (response.AF.AF_GlobalStatus == "<%=BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ACCEPT%>") {
			s_color = "green";
		}
		else if (response.AF.AF_GlobalStatus == "<%=BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REJECT%>") {
			s_color = "red";
		}
		else if (response.AF.AF_GlobalStatus == "<%=BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW%>") {
			s_color = "#FF8C00";
		}
		else {
			s_color = "black";
		}
		s_id = "#spnStAF_" + response.AF.id_pedido_pagto_braspag;
		$(s_id).text(response.AF.AF_DescricaoGlobalStatus);
		$(s_id).css("color", s_color);
		
		s_id = "#spnDtHrStAF_" + response.AF.id_pedido_pagto_braspag;
		$(s_id).text(response.AF.AF_GlobalStatus_atualizacao_data_hora);
		
		s_id = "#spnAfRevisado_" + response.AF.id_pedido_pagto_braspag;
		if (response.AF.AF_SucessoOperacao == "1"){
			s = "Sim";
			s_color = "green";
			$(s_id).addClass("REVIEWED");
		}
		else {
			s = "Não";
			s_color = "red";
		}
		$(s_id).text(s);
		$(s_id).css("color", s_color);
		
		if (response.AF.AF_SucessoOperacao == "1"){
			s_id = "#c_AF_review_tratado_comentario_" + response.AF.id_pedido_pagto_braspag;
			$(s_id).attr("readonly", "readonly");
			s_id = "#spnDtHrRevisaoManual_" + response.AF.id_pedido_pagto_braspag;
			$(s_id).text(response.AF.AF_GlobalStatus_atualizacao_data_hora);
			s_id = "#spnUsuarioRevisaoManual_" + response.AF.id_pedido_pagto_braspag;
			$(s_id).text(response.AF.AF_usuario);
			if (response.AF.AF_NewDecision == "<%=BRASPAG_AF_DECISION__ACCEPT%>") {s_color="green";} else {s_color="red";}
			s_id = "#spnDecisaoRevisaoManual_" + response.AF.id_pedido_pagto_braspag;
			$(s_id).text(response.AF.AF_NewDecision);
			$(s_id).css("color", s_color);
		}
		
	// Dados de resposta do Pagador
		if ((response.PAG.PAG_ErrorCode.length > 0)||(response.PAG.PAG_ErrorMessage.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.PAG.PAG_ErrorCode + ": " + response.PAG.PAG_ErrorMessage;
		}
		if ((response.PAG.PAG_faultcode.length > 0)||(response.PAG.PAG_faultstring.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.PAG.PAG_faultcode + ": " + response.PAG.PAG_faultstring;
		}
		if (strMsgErro.length > 0) {
			strMsgErro = "Falha no processamento da requisição!!\nIMPORTANTE: consulte o status atualizado para verificar a necessidade de refazer esta operação ou de realizar o cancelamento/estorno manual da transação!!\n\n" + strMsgErro;
			alert(strMsgErro);
			return;
		}
		
		if (response.PAG.PAG_GlobalStatus == "<%=BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA%>") {
			s_color = "green";
		}
		else if (response.PAG.PAG_GlobalStatus == "<%=BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA%>") {
			s_color = "black";
		}
		else {
			s_color = "red";
		}
		s_id = "#spnStTransacao_" + response.PAG.id_pedido_pagto_braspag;
		$(s_id).text(response.PAG.PAG_DescricaoGlobalStatus);
		$(s_id).css("color", s_color);
		
		s_id = "#spnDtHrStTransacao_" + response.PAG.id_pedido_pagto_braspag;
		$(s_id).text(response.PAG.PAG_GlobalStatus_atualizacao_data_hora);
		
	// Mostra todos os campos? (para fins de debug)
		if (blnShowJsonFields){
			$.each(response.AF, function(key, val){
				if (strMsg.length > 0) strMsg += "\n";
				strMsg += key + ": " + val;
			});
			$.each(response.PAG, function(key, val){
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

function braspagAfConsultaStatus(_id_pedido_pagto_braspag, _id_pedido_pagto_braspag_af, _id_pedido_pagto_braspag_pag, _usuario){
var s_id, s_color;
var strMsg = "";
var strMsgErro = "";
var blnShowJsonFields = false;
var jqxhr = $.ajax({
	url: "../Global/AjaxBraspagAfConsultaStatus.asp",
	type: "POST",
	dataType: 'json',
	data: {
		id_pedido_pagto_braspag : _id_pedido_pagto_braspag.toString(),
		id_pedido_pagto_braspag_af : _id_pedido_pagto_braspag_af.toString(),
		id_pedido_pagto_braspag_pag : _id_pedido_pagto_braspag_pag.toString(),
		usuario : _usuario
		}
	})
	.done(function(response){
	// Dados de resposta do Antifraude
		if ((response.AF.AF_ErrorCode.length > 0)||(response.AF.AF_ErrorMessage.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.AF.AF_ErrorCode + ": " + response.AF.AF_ErrorMessage;
		}
		if ((response.AF.AF_faultcode.length > 0)||(response.AF.AF_faultstring.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.AF.AF_faultcode + ": " + response.AF.AF_faultstring;
		}
		if (strMsgErro.length > 0) {
			strMsgErro = "Falha na consulta!!\n\n" + strMsgErro;
			alert(strMsgErro);
			return;
		}
		
		if (response.AF.AF_GlobalStatus == "<%=BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ACCEPT%>") {
			s_color = "green";
		}
		else if (response.AF.AF_GlobalStatus == "<%=BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REJECT%>") {
			s_color = "red";
		}
		else if (response.AF.AF_GlobalStatus == "<%=BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW%>") {
			s_color = "#FF8C00";
		}
		else {
			s_color = "black";
		}
		s_id = "#spnStAF_" + response.AF.id_pedido_pagto_braspag;
		$(s_id).text(response.AF.AF_DescricaoGlobalStatus);
		$(s_id).css("color", s_color);
		
		s_id = "#spnDtHrStAF_" + response.AF.id_pedido_pagto_braspag;
		$(s_id).text(response.AF.AF_GlobalStatus_atualizacao_data_hora);
		
	// Dados de resposta do Pagador
		if ((response.PAG.PAG_ErrorCode.length > 0)||(response.PAG.PAG_ErrorMessage.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.PAG.PAG_ErrorCode + ": " + response.PAG.PAG_ErrorMessage;
		}
		if ((response.PAG.PAG_faultcode.length > 0)||(response.PAG.PAG_faultstring.length > 0)) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.PAG.PAG_faultcode + ": " + response.PAG.PAG_faultstring;
		}
		if (strMsgErro.length > 0) {
			strMsgErro = "Falha na consulta!!\n\n" + strMsgErro;
			alert(strMsgErro);
			return;
		}
		
		if (response.PAG.PAG_GlobalStatus == "<%=BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA%>") {
			s_color = "green";
		}
		else if (response.PAG.PAG_GlobalStatus == "<%=BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA%>") {
			s_color = "black";
		}
		else {
			s_color = "red";
		}
		s_id = "#spnStTransacao_" + response.PAG.id_pedido_pagto_braspag;
		$(s_id).text(response.PAG.PAG_DescricaoGlobalStatus);
		$(s_id).css("color", s_color);
		
		s_id = "#spnDtHrStTransacao_" + response.PAG.id_pedido_pagto_braspag;
		$(s_id).text(response.PAG.PAG_GlobalStatus_atualizacao_data_hora);
		
	// Mostra todos os campos? (para fins de debug)
		if (blnShowJsonFields){
			$.each(response.AF, function(key, val){
				if (strMsg.length > 0) strMsg += "\n";
				strMsg += key + ": " + val;
			});
			$.each(response.PAG, function(key, val){
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

function gravarAnaliseCreditoOk(pedido, id_operacao_origem, usuario){
	if (!confirm("Confirma a atualização do status da análise de crédito para 'Crédito OK'?")) return;
	executaUpdateStatusAnaliseCredito(<%=COD_AN_CREDITO_OK%>, pedido, id_operacao_origem, usuario, "RelBraspagAfReviewExec.asp");
}

function executaUpdateStatusAnaliseCredito(_novo_status_analise_credito, _pedido, _id_operacao_origem, _usuario, _pagina_origem){
var s_id;
var strMsgErro = "";

var jqxhr = $.ajax({
	url: "../Global/AjaxAnaliseCreditoUpdateStatus.asp",
	type: "POST",
	dataType: 'json',
	data: {
		novo_status_analise_credito : _novo_status_analise_credito,
		pedido : _pedido,
		id_operacao_origem : _id_operacao_origem,
		usuario : _usuario,
		pagina_origem : _pagina_origem
		}
	})
	.done(function(response){
	// Dados de resposta
		if (response.msg_erro.length > 0) {
			if (strMsgErro.length > 0) strMsgErro += "\n";
			strMsgErro += response.msg_erro;
		}
		if (strMsgErro.length > 0) {
			strMsgErro = "Falha ao tentar atualizar o status da análise de crédito!!\n\n" + strMsgErro;
			alert(strMsgErro);
			return;
		}
		
		s_id = "#lnkAnaliseCreditoOk_" + response.id_operacao_origem;

		if (response.resultado_operacao == "OK") {
			$(s_id).addClass("LnkCredOkDisabled");
			$(s_id).removeClass("LnkCredOkEnabled");
			alert("Status atualizado com sucesso!");
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

function ocultaInfoAnEnd(id_row) {
	var s_id_ln1, s_id_ln2, s_id_img, s_id_href;
	s_id_ln1 = "#TR_INFO_AN_END_LN1_" + id_row;
	s_id_ln2 = "#TR_INFO_AN_END_LN2_" + id_row;
	s_id_img = "#imgPlusMinusPedAnEnd_" + id_row;
	s_id_href = "#hrefPedAnEnd_" + id_row;
	$(s_id_ln1).hide();
	$(s_id_ln1).addClass("TR_INFO_AN_END_HIDDEN");
	$(s_id_ln2).hide();
	$(s_id_ln2).addClass("TR_INFO_AN_END_HIDDEN");
	$(s_id_img).attr({ src: '../imagem/plus.gif' });
	$(s_id_href).attr({ title: 'clique para exibir mais detalhes' });
}

function exibeOcultaInfoAnEnd(id_row) {
	var s_id_ln1, s_id_ln2, s_id_img, s_id_href;
	s_id_ln1 = "#TR_INFO_AN_END_LN1_" + id_row;
	s_id_ln2 = "#TR_INFO_AN_END_LN2_" + id_row;
	s_id_img = "#imgPlusMinusPedAnEnd_" + id_row;
	s_id_href = "#hrefPedAnEnd_" + id_row;
	if ($(s_id_ln1).hasClass("TR_INFO_AN_END_HIDDEN")) {
		$(s_id_ln1).show();
		$(s_id_ln1).removeClass("TR_INFO_AN_END_HIDDEN");
		$(s_id_ln2).show();
		$(s_id_ln2).removeClass("TR_INFO_AN_END_HIDDEN");
		$(s_id_img).attr({ src: '../imagem/minus.gif' });
		$(s_id_href).attr({ title: 'clique para ocultar os detalhes' });
	}
	else {
		$(s_id_ln1).hide();
		$(s_id_ln1).addClass("TR_INFO_AN_END_HIDDEN");
		$(s_id_ln2).hide();
		$(s_id_ln2).addClass("TR_INFO_AN_END_HIDDEN");
		$(s_id_img).attr({ src: '../imagem/plus.gif' });
		$(s_id_href).attr({ title: 'clique para exibir mais detalhes' });
	}
}

function exibeOcultaTodosInfoAnEnd(id_bloco) {
var s_tit_id_img, s_tit_id_href, s_tit_id_span;
var s_item_img_classe, s_item_href_classe;
var s_classe;
	s_classe = ".TR_INFO_AN_END_BLOCO_" + id_bloco;
	s_tit_id_img = "#imgPlusMinusTitAnEnd_" + id_bloco;
	s_tit_id_href = "#hrefTitAnEnd_" + id_bloco;
	s_tit_id_span = "#spanTitAnEnd_" + id_bloco;
	s_item_img_classe = ".imgPlusMinusAnEndBloco_" + id_bloco;
	s_item_href_classe = ".hrefAnEndBloco_" + id_bloco;
	if ($(s_tit_id_span).hasClass("TR_INFO_AN_END_HIDDEN")) {
		$(s_classe).show();
		$(s_classe).removeClass("TR_INFO_AN_END_HIDDEN");
		$(s_tit_id_span).removeClass("TR_INFO_AN_END_HIDDEN");
		$(s_tit_id_img).attr({ src: '../imagem/minus.gif' });
		$(s_tit_id_href).attr({ title: 'clique para ocultar os detalhes' });
		$(s_item_img_classe).attr({ src: '../imagem/minus.gif' });
		$(s_item_href_classe).attr({ title: 'clique para ocultar os detalhes' });
	}
	else {
		$(s_classe).hide();
		$(s_classe).addClass("TR_INFO_AN_END_HIDDEN");
		$(s_tit_id_span).addClass("TR_INFO_AN_END_HIDDEN");
		$(s_tit_id_img).attr({ src: '../imagem/plus.gif' });
		$(s_tit_id_href).attr({ title: 'clique para exibir mais detalhes' });
		$(s_item_img_classe).attr({ src: '../imagem/plus.gif' });
		$(s_item_href_classe).attr({ title: 'clique para exibir mais detalhes' });
	}
}

function fPEDConsulta(id_pedido, usuario) {
	windowScrollTopAnterior = $(window).scrollTop();
	sizeDivPedidoConsulta();
	$("#iframePedidoConsulta").attr("src","PedidoConsultaView.asp?pedido_selecionado="+id_pedido+"&pedido_selecionado_inicial="+id_pedido+"&usuario="+usuario);
	$("#divPedidoConsulta").fadeIn();
}

function fFiltroVoltar(){
	fPED.action = "RelBraspagAfReviewFiltro.asp";
	fPED.submit();
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
	width: 60px;
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
	width: 65px;
}
.tdRevisado
{
	text-align:center;
	vertical-align: middle;
	font-weight: bold;
	width: 50px;
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
.tdStAF
{
	text-align: center;
	vertical-align: middle;
	font-weight: bold;
	width: 75px;
}
.tdDtHrStAF
{
	text-align: center;
	vertical-align: middle;
	width: 65px;
}
.tdCliente{
	text-align: left;
	vertical-align: middle;
	width: 190px;
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
	width: 642px;
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
.LnkCredOkEnabled
{
	cursor:pointer;
}
.LnkCredOkEnabled p
{
	color:green;
	cursor:pointer;
}
.LnkCredOkDisabled
{
	cursor:default;
}
.LnkCredOkDisabled p
{
	color:gray;
	cursor:default;
}
.spnLnkPedAnEnd
{
	cursor:pointer;
}
.tdAnEndPed
{
	width:80px;
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
<input type="hidden" name="c_dt_transacao_inicio" id="c_dt_transacao_inicio" value="<%=c_dt_transacao_inicio%>" />
<input type="hidden" name="c_dt_transacao_termino" id="c_dt_transacao_termino" value="<%=c_dt_transacao_termino%>" />
<input type="hidden" name="c_dt_tratado_inicio" id="c_dt_tratado_inicio" value="<%=c_dt_tratado_inicio%>" />
<input type="hidden" name="c_dt_tratado_termino" id="c_dt_tratado_termino" value="<%=c_dt_tratado_termino%>" />
<input type="hidden" name="c_resultado_transacao" id="c_resultado_transacao" value="<%=c_resultado_transacao%>" />
<input type="hidden" name="c_bandeira" id="c_bandeira" value="<%=c_bandeira%>" />
<input type="hidden" name="c_pedido" id="c_pedido" value="<%=c_pedido%>" />
<input type="hidden" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" value="<%=c_cliente_cnpj_cpf%>" />
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>" />
<input type="hidden" name="rb_ordenacao_saida" id="rb_ordenacao_saida" value="<%=rb_ordenacao_saida%>" />
<input type="hidden" name="rb_tratadas" id="rb_tratadas" value="<%=rb_tratadas%>" />

<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<input type="hidden" name="c_usuario_sessao" id="c_usuario_sessao" value="<%=usuario%>" />



<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="982" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Revisão Manual Antifraude Braspag</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='982' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black;' border='0'>" & chr(13)

'	PERÍODO TRANSAÇÃO EFETUADA
	s = ""
	s_aux = c_dt_transacao_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_transacao_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Período Transação Efetuada:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	PERÍODO REVISÃO TRATADA
	s = ""
	s_aux = c_dt_tratado_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_tratado_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Período Revisão Tratada:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	RESULTADO DA TRANSAÇÃO
	s = c_resultado_transacao
	if s = "" then
		s = "N.I."
	else
		s = iniciais_em_maiusculas(descricao_cod_rel_braspag_af_review(c_resultado_transacao))
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
<table width="982" cellpadding="0" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width='982' cellpadding='0' cellspacing='0' border='0' style="margin-top:5px;">
<tr>
	<td width="50%" align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkImprimir" href="javascript:window.print();"><p class="Button BtnAll" style="margin-bottom:0px;">Imprimir...</p></a></td>
	<td align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkExpandirTudo" href="javascript:expandirTudo();"><p class="Button BtnAll" style="margin-bottom:0px;">Expandir Tudo</p></a></td>
	<td align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkRecolherTudo" href="javascript:recolherTudo();"><p class="Button BtnAll" style="margin-bottom:0px;">Recolher Tudo</p></a></td>
</tr>
</table>

<br />
<table class="notPrint" width="982" cellspacing="0" border="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fFiltroVoltar();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
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
