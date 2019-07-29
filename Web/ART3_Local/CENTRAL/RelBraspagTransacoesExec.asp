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
'	  RelBraspagTransacoesExec.asp
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
	
	Const EXIBIR_BOTAO_CAPTURA = False
	
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
			s = "SELECT nome_iniciais_em_maiusculas FROM t_CLIENTE WHERE (cnpj_cpf = '" & c_cliente_cnpj_cpf & "')"
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
dim s, s_dados, s_sql, x
dim s_where
dim s_color, strRevisaoManual
dim r
dim cab_table, cab
dim vl_total_geral
dim v, i

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
		s_where = s_where & " (cnpj_cpf = '" & retorna_so_digitos(c_cliente_cnpj_cpf) & "')"
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
	s_sql = "SELECT" & _
				" t_PEDIDO_PAGTO_BRASPAG.*," & _
				" tPPB_PAG.id AS id_pedido_pagto_braspag_pag," & _
				" tPPB_AF.id AS id_pedido_pagto_braspag_af," & _
				" t_CLIENTE.cnpj_cpf," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas AS cliente_nome," & _
				" t_PEDIDO.numero_loja," & _
				" tPPB_PAG.Req_PaymentDataCollection_PaymentPlan AS PAG_Req_PaymentDataCollection_PaymentPlan," & _
				" tPPB_PAG.Req_PaymentDataCollection_NumberOfPayments AS PAG_Req_PaymentDataCollection_NumberOfPayments," & _
				" tPPB_PAG.Resp_OrderData_BraspagOrderId AS PAG_Resp_OrderData_BraspagOrderId," & _
				" tPPB_PAG.Resp_Success AS PAG_Resp_Success," & _
				" tPPB_PAG.Resp_PaymentDataResponse_BraspagTransactionId AS PAG_Resp_PaymentDataResponse_BraspagTransactionId," & _
				" tPPB_PAG.Resp_PaymentDataResponse_AuthorizationCode AS PAG_Resp_PaymentDataResponse_AuthorizationCode," & _
				" tPPB_PAG.Resp_PaymentDataResponse_ProofOfSale AS PAG_Resp_PaymentDataResponse_ProofOfSale," & _
				" tPPB_PAG.Resp_PaymentDataResponse_ReturnMessage AS PAG_ReturnMessage," & _
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
				" INNER JOIN t_CLIENTE ON (t_PEDIDO_PAGTO_BRASPAG.id_cliente = t_CLIENTE.id)" & _
				" INNER JOIN t_PEDIDO_PAGTO_BRASPAG_PAG tPPB_PAG ON (t_PEDIDO_PAGTO_BRASPAG.id = tPPB_PAG.id_pedido_pagto_braspag)" & _
				" INNER JOIN t_PEDIDO_PAGTO_BRASPAG_AF tPPB_AF ON (t_PEDIDO_PAGTO_BRASPAG.id = tPPB_AF.id_pedido_pagto_braspag)" & _
			" WHERE" & _
				" (operacao = '" & OP_BRASPAG_OPERACAO__AF_PAG & "')"
	
	if c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_CAPTURADA then
		s_sql = s_sql & _
				" AND (ult_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_AUTORIZADA then
		s_sql = s_sql & _
				" AND (ult_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_NAO_AUTORIZADA then
		s_sql = s_sql & _
				" AND (ult_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__NAO_AUTORIZADA & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_CAPTURA_CANCELADA then
		s_sql = s_sql & _
				" AND (ult_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURA_CANCELADA & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_ESTORNADA then
		s_sql = s_sql & _
				" AND (ult_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ESTORNADA & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_COM_ERRO_DESQUALIFICANTE then
		s_sql = s_sql & _
				" AND (ult_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__ERRO_DESQUALIFICANTE & "')"
	elseif c_resultado_transacao = COD_REL_TRANSACOES_BRASPAG__TRANSACAO_AGUARDANDO_RESPOSTA then
		s_sql = s_sql & _
				" AND (ult_PAG_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AGUARDANDO_RESPOSTA & "')"
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
									" value='" & Trim("" & r("id")) & "|" & Trim("" & r("pedido")) & "|" & Trim("" & r("PAG_Resp_PaymentDataResponse_BraspagTransactionId")) & "'" & _
									" onclick=""configuraEdicaoCampoObs(this, '" & Trim("" & r("id")) & "');""" & _
									">" & chr(13) & _
							"</td>" & chr(13)
		else
			x = x & _
					"		<td class='MTD tdFinalizado'>" & chr(13) & _
					"			<input type='checkbox' name='ckb_ja_tratado_readonly' class='CheckTratado CheckTratadoReadOnly'" & _
									" value='" & Trim("" & r("id")) & "|" & Trim("" & r("pedido")) & "|" & Trim("" & r("PAG_Resp_PaymentDataResponse_BraspagTransactionId")) & "'" & _
									" checked='checked' disabled='disabled' />" & chr(13) & _
					"			<input type='hidden' name='c_flag_ja_tratado_" & Trim("" & r("id")) & "' id='c_flag_ja_tratado_" & Trim("" & r("id")) & "'" & _
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
												primeiroNaoVazio(Array(Trim("" & r("PAG_ReturnMessage")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Mensagem de Erro:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("PAG_ErrorMessage")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
	'	BLOCO DE DADOS: ANTIFRAUDE
		if UCase(Trim("" & r("ult_AF_GlobalStatus"))) = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__ACCEPT then
			s_color = "black"
		elseif UCase(Trim("" & r("ult_AF_GlobalStatus"))) = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW then
			s_color = "#FF8C00"
		else
			s_color ="red"
			end if
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco' rowspan='16' align='center' valign='middle'><img src='../Imagem/braspag-AF-vert.png' width='24' height='38' /></td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Status:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo' style='font-weight:bold;color:" & s_color & ";'>" & chr(13) & _
												primeiroNaoVazio(Array(BraspagAntiFraudeDescricaoGlobalStatus(Trim("" & r("ult_AF_GlobalStatus"))), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		if UCase(Trim("" & r("prim_AF_GlobalStatus"))) = BRASPAG_ANTIFRAUDE_CARTAO_GLOBAL_STATUS__REVIEW then
			strRevisaoManual = "Sim"
		else
			strRevisaoManual = "Não"
			end if
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Indicação de Revisão Manual:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												strRevisaoManual & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		x = x & _
			"							<tr>" & chr(13) & _
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
			s_color = "black"
		elseif Trim("" & r("AF_review_tratado_decision")) = BRASPAG_AF_DECISION__REJECT then
			s_color ="black"
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
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Mensagem de Erro:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("AF_ErrorMessage")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='MC MD tdTitBloco'>&nbsp;</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Observação:<br />" & _
												"<span class='lblTitTamRestante'>(tamanho restante:&nbsp;</span>" & _
												"<span class='lblTamRestante' name='lblTamRestante_" & Trim("" & r("id")) & "' id='lblTamRestante_" & Trim("" & r("id")) & "'></span>" & _
												"<span class='lblTitTamRestante'>)</span>" & _
											"</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
											"	<textarea style='display:none;'" & _
													" name='c_tratado_manual_obs_original_" & Trim("" & r("id")) & "'" & _
													" id='c_tratado_manual_obs_original_" & Trim("" & r("id")) & "'" & _
													">" & Trim("" & r("tratado_manual_obs")) & "</textarea>"  & chr(13) & _
											"	<textarea class='TextAreaObsEdit' maxlength='400' style='width:100%;height:100px;border:0px;'" & _
													" name='c_tratado_manual_obs_" & Trim("" & r("id")) & "'" & _
													" id='c_tratado_manual_obs_" & Trim("" & r("id")) & "'" & _
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
			"											<td align='center'><a href='javascript:braspagPagConsultaStatus(" & Trim("" & r("id")) & "," & Trim("" & r("id_pedido_pagto_braspag_pag")) & ",""" & usuario & """);'><p class='Button BtnBraspag'>&nbsp;Consulta Status Atualizado&nbsp;</p></a></td>" & chr(13) & _
			"											<td align='left' style='width:15px;'>&nbsp;</td>" & chr(13) & _
			"											<td align='center'><a href='javascript:braspagPagExecutaVoidOrRefund(" & Trim("" & r("id")) & "," & Trim("" & r("id_pedido_pagto_braspag_pag")) & ",""" & Trim("" & r("pedido")) & """,""" & usuario & """);'><p class='Button BtnBraspag'>&nbsp;Cancelamento/Estorno&nbsp;</p></a></td>" & chr(13)
		
		if EXIBIR_BOTAO_CAPTURA then
			x = x & _
			"											<td align='left' style='width:15px;'>&nbsp;</td>" & chr(13) & _
			"											<td align='center'><a href='javascript:braspagPagExecutaCapture(" & Trim("" & r("id")) & "," & Trim("" & r("id_pedido_pagto_braspag_pag")) & ",""" & Trim("" & r("pedido")) & """,""" & usuario & """);'><p class='Button BtnBraspag'>&nbsp;Captura&nbsp;</p></a></td>" & chr(13)
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

function braspagPagExecutaVoidOrRefund(_id_pedido_pagto_braspag, _id_pedido_pagto_braspag_pag, _pedido, _usuario){
var s_id, s_color;
var strMsg = "";
var strMsgErro = "";
var blnShowJsonFields = false;

if (!confirm("Confirma o cancelamento/estorno da transação do pedido " + _pedido + "?")) return;

var jqxhr = $.ajax({
	url: "../Global/AjaxBraspagPagCancelOuEstorno.asp",
	type: "POST",
	dataType: 'json',
	data: {
		id_pedido_pagto_braspag : _id_pedido_pagto_braspag.toString(),
		id_pedido_pagto_braspag_pag : _id_pedido_pagto_braspag_pag.toString(),
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
		s_id = "#spnStTransacao_" + response.id_pedido_pagto_braspag;
		$(s_id).text(response.PAG_DescricaoGlobalStatus);
		$(s_id).css("color", s_color);
		
		s_id = "#spnDtHrStTransacao_" + response.id_pedido_pagto_braspag;
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

function braspagPagExecutaCapture(_id_pedido_pagto_braspag, _id_pedido_pagto_braspag_pag, _pedido, _usuario){
var s_id, s_color;
var strMsg = "";
var strMsgErro = "";
var blnShowJsonFields = false;

if (!confirm("Confirma a captura da transação do pedido " + _pedido + "?")) return;

var jqxhr = $.ajax({
	url: "../Global/AjaxBraspagPagCaptura.asp",
	type: "POST",
	dataType: 'json',
	data: {
		id_pedido_pagto_braspag : _id_pedido_pagto_braspag.toString(),
		id_pedido_pagto_braspag_pag : _id_pedido_pagto_braspag_pag.toString(),
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
		s_id = "#spnStTransacao_" + response.id_pedido_pagto_braspag;
		$(s_id).text(response.PAG_DescricaoGlobalStatus);
		$(s_id).css("color", s_color);
		
		s_id = "#spnDtHrStTransacao_" + response.id_pedido_pagto_braspag;
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

function braspagPagConsultaStatus(_id_pedido_pagto_braspag, _id_pedido_pagto_braspag_pag, _usuario){
var s_id, s_color;
var strMsg = "";
var strMsgErro = "";
var blnShowJsonFields = false;
var jqxhr = $.ajax({
	url: "../Global/AjaxBraspagPagConsultaStatus.asp",
	type: "POST",
	dataType: 'json',
	data: {
		id_pedido_pagto_braspag : _id_pedido_pagto_braspag.toString(),
		id_pedido_pagto_braspag_pag : _id_pedido_pagto_braspag_pag.toString(),
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
		s_id = "#spnStTransacao_" + response.id_pedido_pagto_braspag;
		$(s_id).text(response.PAG_DescricaoGlobalStatus);
		$(s_id).css("color", s_color);
		
		s_id = "#spnDtHrStTransacao_" + response.id_pedido_pagto_braspag;
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
	fPED.action = "RelBraspagTransacoesFiltro.asp";
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
	f.action = "RelBraspagTransacoesGravaDados.asp";
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
	<td align="right" valign="bottom"><span class="PEDIDO">Transações Braspag</span>
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
