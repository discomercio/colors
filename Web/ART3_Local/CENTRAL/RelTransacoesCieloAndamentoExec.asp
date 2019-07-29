<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelTransacoesCieloAndamentoExec.asp
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
	if Not operacao_permitida(OP_CEN_REL_TRANSACOES_CIELO_ANDAMENTO, s_lista_operacoes_permitidas) then
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim url_back, strUrlBotaoVoltar
	url_back = Trim(Request("url_back"))
	if url_back <> "" then
		strUrlBotaoVoltar = "RelTransacoesCieloAndamentoFiltro.asp?url_back=X&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	else
		strUrlBotaoVoltar = "javascript:history.back()"
		end if

	dim s_filtro, intQtdeTransacoes
	intQtdeTransacoes = 0

	dim alerta
	dim s, s_aux
	dim c_dt_inicio, c_dt_termino
	dim c_resultado_transacao, c_bandeira, c_pedido, c_cliente_cnpj_cpf, c_loja, rb_ordenacao_saida
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

function monta_link_pedido(byval id_pedido)
dim strLink
	monta_link_pedido = ""
	id_pedido = Trim("" & id_pedido)
	if id_pedido = "" then exit function
	strLink = "<a href='javascript:fPEDConsulta(" & _
				chr(34) & id_pedido & chr(34) & _
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
		s_where = s_where & " (CONVERT(smallint, pedido_loja) = " & c_loja & ")"
		end if
	
'	MONTAGEM DA CONSULTA
	s_sql = ""
	
	if (c_resultado_transacao = "") Or (c_resultado_transacao = COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_ANDAMENTO) then
		if s_sql <> "" then s_sql = s_sql & " UNION "
		s_sql = s_sql & _
				"SELECT " & _
					"t_PEDIDO_PAGTO_CIELO.*, " & _
					"'" & COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_ANDAMENTO & "' AS situacao_transacao, " & _
					"t_CLIENTE.cnpj_cpf, " & _
					"t_CLIENTE.nome_iniciais_em_maiusculas AS cliente_nome, " & _
					"t_PEDIDO.loja AS pedido_loja" & _
				" FROM t_PEDIDO_PAGTO_CIELO" & _
					" INNER JOIN t_PEDIDO ON (t_PEDIDO_PAGTO_CIELO.pedido = t_PEDIDO.pedido)" & _
					" INNER JOIN t_CLIENTE ON (t_PEDIDO_PAGTO_CIELO.id_cliente = t_CLIENTE.id)" & _
				" WHERE" & _
					" (operacao = '" & OP_CIELO_OPERACAO__PAGAMENTO & "')" & _
					" AND (requisicao_transacao_concluido_status <> 0)" & _
					" AND (requisicao_consulta_concluido_status <> 0)" & _
					" AND (requisicao_consulta_status = '" & CIELO_TRANSACAO_STATUS__EM_ANDAMENTO & "')" & _
					" AND (sucesso_final_status = 0)" & _
					" AND (cancelado_status = 0)" & _
					" AND (tratado_manual_status = 0)"
		end if
	
	if (c_resultado_transacao = "") Or (c_resultado_transacao = COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_AUTENTICACAO) then
		if s_sql <> "" then s_sql = s_sql & " UNION "
		s_sql = s_sql & _
				"SELECT " & _
					"t_PEDIDO_PAGTO_CIELO.*, " & _
					"'" & COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_AUTENTICACAO & "' AS situacao_transacao, " & _
					"t_CLIENTE.cnpj_cpf, " & _
					"t_CLIENTE.nome_iniciais_em_maiusculas AS cliente_nome, " & _
					"t_PEDIDO.loja AS pedido_loja" & _
				" FROM t_PEDIDO_PAGTO_CIELO" & _
					" INNER JOIN t_PEDIDO ON (t_PEDIDO_PAGTO_CIELO.pedido = t_PEDIDO.pedido)" & _
					" INNER JOIN t_CLIENTE ON (t_PEDIDO_PAGTO_CIELO.id_cliente = t_CLIENTE.id)" & _
				" WHERE" & _
					" (operacao = '" & OP_CIELO_OPERACAO__PAGAMENTO & "')" & _
					" AND (requisicao_transacao_concluido_status <> 0)" & _
					" AND (requisicao_consulta_concluido_status <> 0)" & _
					" AND (requisicao_consulta_status = '" & CIELO_TRANSACAO_STATUS__EM_AUTENTICACAO & "')" & _
					" AND (sucesso_final_status = 0)" & _
					" AND (cancelado_status = 0)" & _
					" AND (tratado_manual_status = 0)"
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
		s = monta_link_pedido(Trim("" & r("pedido")))
		x = x & "		<td class='MTD tdPedido'><span class='Cn'>" & s & "</span></td>" & chr(13)

	'> VALOR DO PEDIDO
		s = formata_moeda(r("valor_pedido"))
		x = x & "		<td class='MTD tdVlPedido'><span class='Cnd'>" & s & "</span></td>" & chr(13)

	'> VALOR DA TRANSAÇÃO
		s = formata_moeda(r("valor_transacao"))
		x = x & "		<td class='MTD tdVlTransacao'><span class='Cnd'>" & s & "</span></td>" & chr(13)

	'> BANDEIRA DO CARTÃO
		s = CieloDescricaoBandeira(Trim("" & r("bandeira")))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdBandeira'><span class='Cnc'>" & s & "</span></td>" & chr(13)

	'> STATUS DA TRANSAÇÃO
		s = Trim("" & r("requisicao_consulta_status"))
		if s <> "" then s = CieloDescricaoStatus(s)
		if s = "" then s = "&nbsp;"
		if Trim("" & r("requisicao_consulta_status")) = CIELO_TRANSACAO_STATUS__EM_ANDAMENTO then
			s_color = "#FF8C00"
		else
			s_color = "#B22222"
			end if
		x = x & "		<td class='MTD tdStTransacao'><span class='Cnc' style='color:" & s_color & ";'>" & s & "</span></td>" & chr(13)

	'> CLIENTE
		s = cnpj_cpf_formata(Trim("" & r("cnpj_cpf"))) & " - " & Trim("" & r("cliente_nome"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdCliente'><span class='Cn'>" & s & "</span></td>" & chr(13)

	'> TRATADO
		x = x & _
				"		<td class='MTD tdFinalizado'>" & chr(13) & _
				"			<input type='checkbox' name='ckb_tratado' id='ckb_tratado' class='CheckTratado'" & _
								" value='" & Trim("" & r("id")) & "|" & Trim("" & r("pedido")) & "|" & Trim("" & r("requisicao_transacao_tid")) & "'" & _
								" onclick=""configuraEdicaoCampoObs(this, '" & Trim("" & r("id")) & "');""" & _
								">" & chr(13) & _
						"</td>" & chr(13)

	'> BOTÃO P/ EXIBIR DEMAIS CAMPOS
		x = x & "		<td valign='bottom' class='notPrint' align='left'>" & _
							"&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(intQtdeTransacoes) & chr(34) & ");' title='exibe ou oculta os campos adicionais'><img src='../botao/view_bottom.png' border='0'></a>" & _
						"</td>" & chr(13)
		
		x = x & "	</tr>" & chr(13)

	'> OUTRAS INFORMAÇÕES
		x = x & "	<tr style='display:none;' id='TR_MORE_INFO_" & Cstr(intQtdeTransacoes) & "'>" & chr(13) & _
				"		<td class='ME MD'>&nbsp;</td>" & chr(13) & _
				"		<td colspan='8' class='MC MD' align='left'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<tr>" & chr(13) & _
				"					<td class='Rf tdWithPadding' align='left'>OUTRAS INFORMAÇÕES</td>" & chr(13) & _
				"				</tr>" & chr(13)
		
		x = x & _
			"				<tr>" & chr(13) & _
			"					<td align='left'>" & chr(13) & _
			"						<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Transação (TID):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_transacao_tid")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13)
		
		x = x & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Autenticação (Código):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autenticacao_codigo")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Autenticação (Mensagem):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autenticacao_mensagem")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Autenticação (ECI):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autenticacao_eci")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Autorização (Código):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autorizacao_codigo")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Autorização (Mensagem):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autorizacao_mensagem")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Autorização (LR):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autorizacao_lr")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Autorização (ARP):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autorizacao_arp")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Autorização (NSU):</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autorizacao_nsu")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Opção de Pagamento:</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
												primeiroNaoVazio(Array(CieloDescricaoParcelamento(Trim("" & r("forma_pagamento_produto")), Trim("" & r("forma_pagamento_parcelas")), r("valor_transacao")), "&nbsp;")) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"							<tr>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdTitMoreInfo'>Observação:<br />" & _
												"<span class='lblTitTamRestante'>(tamanho restante:&nbsp;</span>" & _
												"<span class='lblTamRestante' name='lblTamRestante_" & Trim("" & r("id")) & "' id='lblTamRestante_" & Trim("" & r("id")) & "'></span>" & _
												"<span class='lblTitTamRestante'>)</span>" & _
											"</td>" & chr(13) & _
			"								<td class='Cn MC tdWithPadding tdMoreInfo'>" & chr(13) & _
											"	<textarea maxlength='400' style='width:100%;height:100px;border:0px;'" & _
													" name='c_tratado_manual_obs_" & Trim("" & r("id")) & "'" & _
													" id='c_tratado_manual_obs_" & Trim("" & r("id")) & "'" & _
													">" & Trim("" & r("tratado_manual_obs")) & "</textarea>"  & chr(13) & _
			"								</td>" & chr(13) & _
			"							</tr>" & chr(13) & _
			"						</table>" & chr(13) & _
			"					</td>" & chr(13) & _
			"				</tr>" & chr(13)
		
		x = x & _
				"			</table>" & chr(13) & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13)

		if (intQtdeTransacoes mod 100) = 0 then
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
				"		<td colspan='4' class='MC MD'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr>" & chr(13) & _
				"		<td colspan='9' class='MC' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr nowrap style='background:#F0FFF0;'>" & chr(13) & _
				"		<td colspan='9' class='MT' align='left'><span class='C'>TOTAL: &nbsp; " & Cstr(intQtdeTransacoes) & iif((intQtdeTransacoes=1), " transação", " transações") & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if
	
'	MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeTransacoes = 0 then
		x = cab_table & cab
		x = x & "	<tr nowrap>" & chr(13) & _
				"		<td class='MT ALERTA' align='center' colspan='9'><span class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</span></td>" & chr(13) & _
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

		$("span.lblTamRestante").each(function() {
			var s_id = $(this).attr("id");
			s_id = "#" + s_id.replace("lblTamRestante_", "c_tratado_manual_obs_");
			var c = $(s_id);
			var n = MAX_TAM_T_PEDIDO_PAGTO_CIELO__TRATADO_MANUAL_OBS - $(c).val().length;
			$(this).html(n.toString());
		});

		$("textarea").bind("input keyup paste", function() {
			var s_id = $(this).attr("id");
			s_id = "#" + s_id.replace("c_tratado_manual_obs_", "lblTamRestante_");
			var c = $(s_id);
			var n = MAX_TAM_T_PEDIDO_PAGTO_CIELO__TRATADO_MANUAL_OBS - $(this).val().length;
			$(c).html(n.toString());
		});
	});
</script>

<script language="JavaScript" type="text/javascript">
var MAX_TAM_T_PEDIDO_PAGTO_CIELO__TRATADO_MANUAL_OBS=<%=MAX_TAM_T_PEDIDO_PAGTO_CIELO__TRATADO_MANUAL_OBS%>;
window.status = 'Aguarde, executando a consulta ...';

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

function fPEDConsulta(id_pedido) {
	window.status = "Aguarde ...";
	fPED.pedido_selecionado.value = id_pedido;
	fPED.action = "pedido.asp"
	fPED.submit();
}

function fPEDGravaDados(f) {
var i, intQtdeTratados, c, blnExcedeu;

	intQtdeTratados = 0;
	for (i = 0; i < f.ckb_tratado.length; i++) {
		if (f.ckb_tratado[i].checked) intQtdeTratados++;
	}

	if (intQtdeTratados == 0) {
		alert('Nenhuma transação foi assinalada para ser marcada como já tratada!!');
		return;
	}
	
	blnExcedeu = false;
	$("textarea").each(function(){
		if (blnExcedeu) return;
		if ($(this).val().length > MAX_TAM_T_PEDIDO_PAGTO_CIELO__TRATADO_MANUAL_OBS){
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
	f.action = "RelTransacoesCieloAndamentoGravaDados.asp";
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

<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="ckb_tratado" id="ckb_tratado" value="">
<input type="hidden" name="c_tratado_manual_obs_0" id="c_tratado_manual_obs_0" value="">



<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="853" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Transações Cielo</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='853' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black;' border='0'>" & chr(13)

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
		s = iniciais_em_maiusculas(descricao_cod_rel_transacoes_cielo(c_resultado_transacao))
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
		s = CieloDescricaoBandeira(c_bandeira)
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
<table width="853" cellpadding="0" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width='853' cellpadding='0' cellspacing='0' border='0' style="margin-top:5px;">
<tr>
	<td width="25%" align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkMarcarTudo" href="javascript:marcarTodas();"><p class="Button" style="margin-bottom:0px;">Marcar Todas</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkDesmarcarTudo" href="javascript:desmarcarTodas();"><p class="Button" style="margin-bottom:0px;">Desmarcar Todas</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkImprimir" href="javascript:window.print();"><p class="Button" style="margin-bottom:0px;">Imprimir...</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkExpandirTudo" href="javascript:expandirTudo();"><p class="Button" style="margin-bottom:0px;">Expandir Tudo</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkRecolherTudo" href="javascript:recolherTudo();"><p class="Button" style="margin-bottom:0px;">Recolher Tudo</p></a></td>
</tr>
</table>

<br />
<table class="notPrint" width="853" cellspacing="0" border="0">
<tr>
	<% if qtde_transacoes > 0 then %>
	<td align="left">
		<a name="bVOLTAR" id="bVOLTAR" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td>&nbsp;</td>
	<td align="right">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDGravaDados(fPED)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0"></a>
	</td>
	<% else %>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	<% end if %>
	</td>
</tr>
</table>

</form>

</center>
</body>
</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
