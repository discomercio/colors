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
'	  RelTransacoesCieloExec.asp
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
	if Not operacao_permitida(OP_CEN_REL_TRANSACOES_CIELO, s_lista_operacoes_permitidas) then
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim s_filtro, intQtdeTransacoes
	intQtdeTransacoes = 0

	dim alerta
	dim s, s_aux
	dim c_dt_inicio, c_dt_termino
	dim c_resultado_transacao, c_bandeira, c_pedido, c_cliente_cnpj_cpf, c_loja
	dim s_nome_cliente, s_nome_loja

	alerta = ""

	c_dt_inicio = Trim(Request("c_dt_inicio"))
	c_dt_termino = Trim(Request("c_dt_termino"))
	c_resultado_transacao = Trim(Request("c_resultado_transacao"))
	c_bandeira = Trim(Request("c_bandeira"))
	c_pedido = Trim(Request("c_pedido"))
	c_cliente_cnpj_cpf = retorna_so_digitos(Trim(Request("c_cliente_cnpj_cpf")))
	c_loja = retorna_so_digitos(Trim(Request("c_loja")))
	
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
			s = "SELECT nome_iniciais_em_maiusculas FROM t_CLIENTE WHERE (cnpj_cpf='" & c_cliente_cnpj_cpf & "')"
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
dim qtde_transacao_autorizada, qtde_transacao_nao_autorizada, qtde_transacao_cancelada_pelo_usuario, qtde_transacao_situacao_desconhecida
dim vl_total_geral, vl_total_transacao_autorizada, vl_total_transacao_nao_autorizada, vl_total_transacao_cancelada_pelo_usuario, vl_total_transacao_situacao_desconhecida

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
	
	if (c_resultado_transacao = "") Or (c_resultado_transacao = COD_REL_TRANSACOES_CIELO__TRANSACAO_AUTORIZADA) then
		if s_sql <> "" then s_sql = s_sql & " UNION "
		s_sql = s_sql & _
				"SELECT " & _
					"t_PEDIDO_PAGTO_CIELO.*, " & _
					"'" & COD_REL_TRANSACOES_CIELO__TRANSACAO_AUTORIZADA & "' AS situacao_transacao, " & _
					"t_CLIENTE.cnpj_cpf, " & _
					"t_CLIENTE.nome_iniciais_em_maiusculas AS cliente_nome, " & _
					"t_PEDIDO.loja AS pedido_loja" & _
				" FROM t_PEDIDO_PAGTO_CIELO" & _
					" INNER JOIN t_PEDIDO ON (t_PEDIDO_PAGTO_CIELO.pedido = t_PEDIDO.pedido)" & _
					" INNER JOIN t_CLIENTE ON (t_PEDIDO_PAGTO_CIELO.id_cliente = t_CLIENTE.id)" & _
				" WHERE" & _
					" (operacao = '" & OP_CIELO_OPERACAO__PAGAMENTO & "')" & _
					" AND (sucesso_final_status <> 0)" & _
					" AND (cancelado_status = 0)"
		end if
	
	if (c_resultado_transacao = "") Or (c_resultado_transacao = COD_REL_TRANSACOES_CIELO__TRANSACAO_NAO_AUTORIZADA) then
		if s_sql <> "" then s_sql = s_sql & " UNION "
		s_sql = s_sql & _
				"SELECT " & _
					"t_PEDIDO_PAGTO_CIELO.*, " & _
					"'" & COD_REL_TRANSACOES_CIELO__TRANSACAO_NAO_AUTORIZADA & "' AS situacao_transacao, " & _
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
					" AND (requisicao_consulta_status = '" & CIELO_TRANSACAO_STATUS__NAO_AUTORIZADA & "')" & _
					" AND (sucesso_final_status = 0)" & _
					" AND (cancelado_status = 0)"
		end if
	
	if (c_resultado_transacao = "") Or (c_resultado_transacao = COD_REL_TRANSACOES_CIELO__TRANSACAO_CANCELADA_PELO_USUARIO) then
		if s_sql <> "" then s_sql = s_sql & " UNION "
		s_sql = s_sql & _
				"SELECT " & _
					"t_PEDIDO_PAGTO_CIELO.*, " & _
					"'" & COD_REL_TRANSACOES_CIELO__TRANSACAO_CANCELADA_PELO_USUARIO & "' AS situacao_transacao, " & _
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
					" AND (requisicao_consulta_status = '" & CIELO_TRANSACAO_STATUS__CANCELADA & "')" & _
					" AND (sucesso_final_status = 0)" & _
					" AND (cancelado_status = 0)"
		end if
	
	if (c_resultado_transacao = "") Or (c_resultado_transacao = COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_SITUACAO_DESCONHECIDA) then
		if s_sql <> "" then s_sql = s_sql & " UNION "
		s_sql = s_sql & _
				"SELECT " & _
					"t_PEDIDO_PAGTO_CIELO.*, " & _
					"'" & COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_SITUACAO_DESCONHECIDA & "' AS situacao_transacao, " & _
					"t_CLIENTE.cnpj_cpf, " & _
					"t_CLIENTE.nome_iniciais_em_maiusculas AS cliente_nome, " & _
					"t_PEDIDO.loja AS pedido_loja" & _
				" FROM t_PEDIDO_PAGTO_CIELO" & _
					" INNER JOIN t_PEDIDO ON (t_PEDIDO_PAGTO_CIELO.pedido = t_PEDIDO.pedido)" & _
					" INNER JOIN t_CLIENTE ON (t_PEDIDO_PAGTO_CIELO.id_cliente = t_CLIENTE.id)" & _
				" WHERE" & _
				" (operacao = '" & OP_CIELO_OPERACAO__PAGAMENTO & "')" & _
				" AND (requisicao_transacao_sucesso_status <> 0)" & _
				" AND " & _
					"(" & _
						"(" & _
							"(ambiente_cielo_redirecionado_status <> 0)" & _
							" AND " & _
							"(requisicao_consulta_concluido_status = 0)" & _
						")" & _
						" OR " & _
						"(" & _
							"(requisicao_consulta_status = " & CIELO_TRANSACAO_STATUS__EM_AUTENTICACAO & ")" & _
						")" & _
					")" & _
				" AND (cancelado_status = 0)"
		end if
	
	if s_sql <> "" then
		if s_where <> "" then s_where = " WHERE " & s_where
		s_sql = "SELECT " & _
					"*" & _
				" FROM (" & s_sql & ") t" & _
				s_where & _
				" ORDER BY" & _
					" id"
		end if
	
	if s_sql = "" then
		Response.Write "Falha ao elaborar a consulta SQL: a consulta não possui conteúdo!"
		Response.End
		end if

	cab_table = "<TABLE cellSpacing=0 cellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		"		<TD class='MDTE tdDataHora' style='vertical-align:bottom'><P class='Rc'>Data</P></TD>" & chr(13) & _
		"		<TD class='MTD tdUsuario' style='vertical-align:bottom'><P class='Rc'>Usuário</P></TD>" & chr(13) & _
		"		<TD class='MTD tdPedido' style='vertical-align:bottom'><P class='R'>Pedido</P></TD>" & chr(13) & _
		"		<TD class='MTD tdVlPedido' style='vertical-align:bottom;padding-right:4px;'><P class='R'>Valor Pedido</P></TD>" & chr(13) & _
		"		<TD class='MTD tdVlTransacao' style='vertical-align:bottom;padding-right:4px;'><P class='R'>Valor Transação</P></TD>" & chr(13) & _
		"		<TD class='MTD tdBandeira' style='vertical-align:bottom'><P class='Rc'>Bandeira</P></TD>" & chr(13) & _
		"		<TD class='MTD tdFinalizado' style='vertical-align:bottom'><P class='Rc'>Finalizado com Sucesso</P></TD>" & chr(13) & _
		"		<TD class='MTD tdStTransacao' style='vertical-align:bottom'><P class='Rc'>Status da Transação</P></TD>" & chr(13) & _
		"		<TD class='MTD tdCliente' style='vertical-align:bottom'><P class='R'>Cliente</P></TD>" & chr(13) & _
		"		<TD style='background:white;'>&nbsp;</TD>" & chr(13) & _
		"	</TR>" & chr(13)
	
	x = cab_table & cab
	intQtdeTransacoes = 0
	qtde_transacao_autorizada = 0
	qtde_transacao_nao_autorizada = 0
	qtde_transacao_cancelada_pelo_usuario = 0
	qtde_transacao_situacao_desconhecida = 0
	vl_total_geral = 0
	vl_total_transacao_autorizada = 0
	vl_total_transacao_nao_autorizada = 0
	vl_total_transacao_cancelada_pelo_usuario = 0
	vl_total_transacao_situacao_desconhecida = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	 ' CONTAGEM
		intQtdeTransacoes = intQtdeTransacoes + 1
		
		vl_total_geral = vl_total_geral + r("valor_transacao")
		
		if Trim("" & r("situacao_transacao")) = COD_REL_TRANSACOES_CIELO__TRANSACAO_AUTORIZADA then
			qtde_transacao_autorizada = qtde_transacao_autorizada + 1
			vl_total_transacao_autorizada = vl_total_transacao_autorizada + r("valor_transacao")
		elseif Trim("" & r("situacao_transacao")) = COD_REL_TRANSACOES_CIELO__TRANSACAO_NAO_AUTORIZADA then
			qtde_transacao_nao_autorizada = qtde_transacao_nao_autorizada + 1
			vl_total_transacao_nao_autorizada = vl_total_transacao_nao_autorizada + r("valor_transacao")
		elseif Trim("" & r("situacao_transacao")) = COD_REL_TRANSACOES_CIELO__TRANSACAO_CANCELADA_PELO_USUARIO then
			qtde_transacao_cancelada_pelo_usuario = qtde_transacao_cancelada_pelo_usuario + 1
			vl_total_transacao_cancelada_pelo_usuario = vl_total_transacao_cancelada_pelo_usuario + r("valor_transacao")
		elseif Trim("" & r("situacao_transacao")) = COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_SITUACAO_DESCONHECIDA then
			qtde_transacao_situacao_desconhecida = qtde_transacao_situacao_desconhecida + 1
			vl_total_transacao_situacao_desconhecida = vl_total_transacao_situacao_desconhecida + r("valor_transacao")
			end if
		
		x = x & "	<TR NOWRAP>" & chr(13)

	'> DATA DA TRANSAÇÃO
		s = formata_data_hora_sem_seg(r("data_hora"))
		x = x & "		<TD class='MDTE tdDataHora'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> USUÁRIO
		s = Trim("" & r("usuario"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdUsuario'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> PEDIDO
		s = monta_link_pedido(Trim("" & r("pedido")))
		x = x & "		<TD class='MTD tdPedido'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> VALOR DO PEDIDO
		s = formata_moeda(r("valor_pedido"))
		x = x & "		<TD class='MTD tdVlPedido'><P class='Cnd'>" & s & "</P></TD>" & chr(13)

	'> VALOR DA TRANSAÇÃO
		s = formata_moeda(r("valor_transacao"))
		x = x & "		<TD class='MTD tdVlTransacao'><P class='Cnd'>" & s & "</P></TD>" & chr(13)

	'> BANDEIRA DO CARTÃO
		s = CieloDescricaoBandeira(Trim("" & r("bandeira")))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdBandeira'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> FINALIZADO COM SUCESSO
		if r("sucesso_final_status") = 0 then
			s = "Não"
			s_color = "red"
		else
			s = "Sim"
			s_color = "green"
			end if
		x = x & "		<TD class='MTD tdFinalizado'><P class='Cn' style='color:" & s_color & ";'>" & s & "</P></TD>" & chr(13)

	'> STATUS DA TRANSAÇÃO
		s = Trim("" & r("requisicao_consulta_status"))
		if s <> "" then s = CieloDescricaoStatus(s)
		if s = "" then s = "&nbsp;"
		if Trim("" & r("requisicao_consulta_status")) = CIELO_TRANSACAO_STATUS__AUTORIZADA then
			s_color = "green"
		else
			s_color = "red"
			end if
		x = x & "		<TD class='MTD tdStTransacao'><P class='Cn' style='color:" & s_color & ";'>" & s & "</P></TD>" & chr(13)

	'> CLIENTE
		s = cnpj_cpf_formata(Trim("" & r("cnpj_cpf"))) & " - " & Trim("" & r("cliente_nome"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdCliente'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> BOTÃO P/ EXIBIR DEMAIS CAMPOS
		x = x & "		<TD valign='bottom' class='notPrint'>" & _
							"&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(intQtdeTransacoes) & chr(34) & ")' title='exibe ou oculta os campos adicionais'><img src='../botao/view_bottom.png' border='0'></a>" & _
						"</TD>" & chr(13)
		
		x = x & "	</TR>" & chr(13)

	'> OUTRAS INFORMAÇÕES
		x = x & "	<TR style='display:none;' id='TR_MORE_INFO_" & Cstr(intQtdeTransacoes) & "'>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='8' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>OUTRAS INFORMAÇÕES</td>" & chr(13) & _
				"				</TR>" & chr(13)
		
		x = x & _
			"				<TR>" & chr(13) & _
			"					<TD>" & chr(13) & _
			"						<table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Transação (TID):</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_transacao_tid")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13)
		
		if Trim("" & r("requisicao_consulta_status")) = CIELO_TRANSACAO_STATUS__CANCELADA then
			x = x & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Cancelamento (Código):</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_cancelamento_codigo")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13) & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Cancelamento (Mensagem):</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_cancelamento_mensagem")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13)
			end if
		
		x = x & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Autenticação (Código):</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autenticacao_codigo")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13) & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Autenticação (Mensagem):</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autenticacao_mensagem")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13) & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Autenticação (ECI):</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autenticacao_eci")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13) & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Autorização (Código):</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autorizacao_codigo")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13) & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Autorização (Mensagem):</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autorizacao_mensagem")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13) & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Autorização (LR):</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autorizacao_lr")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13) & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Autorização (ARP):</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autorizacao_arp")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13) & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Autorização (NSU):</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(Trim("" & r("requisicao_consulta_autorizacao_nsu")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13) & _
			"							<TR>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdTitMoreInfo' align='right'>Opção de Pagamento:</TD>" & chr(13) & _
			"								<TD class='Cn MC tdWithPadding tdMoreInfo' align='left'>" & chr(13) & _
												primeiroNaoVazio(Array(CieloDescricaoParcelamento(Trim("" & r("forma_pagamento_produto")), Trim("" & r("forma_pagamento_parcelas")), r("valor_transacao")), "&nbsp;")) & _
			"								</TD>" & chr(13) & _
			"							</TR>" & chr(13) & _
			"						</table>" & chr(13) & _
			"					</TD>" & chr(13) & _
			"				</TR>" & chr(13)
		
		x = x & _
				"			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)

		if (intQtdeTransacoes mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
		
'	TOTAL GERAL
	if intQtdeTransacoes > 0 then
		x = x & "	<TR>" & chr(13) & _
				"		<TD COLSPAN='4' align='right' class='MC ME'><p class='C'>TOTAL GERAL (" & SIMBOLO_MONETARIO & ")</p></TD>" & chr(13) & _
				"		<TD class='MC' align='right'><p class='Cd'>" & formata_moeda(vl_total_geral) & "</p></TD>" & chr(13) & _
				"		<TD COLSPAN='4' class='MC MD'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR>" & chr(13) & _
				"		<TD COLSPAN='9' class='MC' style='border-left:0px;border-right:0px;'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD COLSPAN='9' class='MT'><p class='C'>TOTAL: &nbsp; " & Cstr(intQtdeTransacoes) & iif((intQtdeTransacoes=1), " transação", " transações") & "</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if
	
'	MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeTransacoes = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='9'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

'	FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

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



<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
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

function fPEDConsulta(id_pedido) {
	window.status = "Aguarde ...";
	fPED.pedido_selecionado.value = id_pedido;
	fPED.action = "pedido.asp"
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">
<link href="<%=URL_FILE__ESCREEN_CSS%>" Rel="stylesheet" Type="text/css" media="screen">

<style type="text/css">
.tdWithPadding
{
	padding:1px;
}
.tdDataHora{
	vertical-align: middle;
	width: 65px;
	}
.tdUsuario{
	vertical-align: middle;
	width: 85px;
	}
.tdPedido{
	vertical-align: middle;
	font-weight: bold;
	width: 65px;
	}
.tdVlPedido
{
	vertical-align: middle;
	text-align: right;
	font-weight: bold;
	width: 70px;
}
.tdVlTransacao
{
	vertical-align: middle;
	text-align: right;
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
	text-align:center;
	vertical-align: middle;
	font-weight: bold;
	width: 85px;
	}
.tdCliente{
	vertical-align: middle;
	width: 240px;
	}
.tdTitMoreInfo{
	vertical-align: top;
	padding-right: 2px;
	width: 200px;
	}
.tdMoreInfo{
	vertical-align: top;
	padding-left: 2px;
	}
</style>


<body onload="window.status='Concluído';focus();" link=#000000 alink=#000000 vlink=#000000>
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
</form>


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="853" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
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
	s_filtro = "<table width='853' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

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
				"		<td align='right' valign='top' NOWRAP><p class='N'>Período:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
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
				"		<td align='right' valign='top' NOWRAP><p class='N'>Resultado da Transação:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	BANDEIRA
	s = c_bandeira
	if s = "" then
		s = "N.I."
	else
		s = CieloDescricaoBandeira(c_bandeira)
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Bandeira:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

'	PEDIDO
	s = c_pedido
	if s = "" then s = "N.I."
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Pedido:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

'	CLIENTE
	s = c_cliente_cnpj_cpf
	if s = "" then 
		s = "N.I."
	else
		s = cnpj_cpf_formata(c_cliente_cnpj_cpf)
		if s_nome_cliente <> "" then s = s & " - " & s_nome_cliente
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Cliente:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

'	LOJA
	s = c_loja
	if s = "" then 
		s = "N.I."
	else
		if s_nome_loja <> "" then s = s & " - " & s_nome_loja
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Loja:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

'	EMISSÃO
	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP>" & _
					"<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
					"<p class='N'>" & formata_data_hora(Now) & "</p></td>" & chr(13) & _
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
<table width="853" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>

<table class="notPrint" width='853' cellPadding='0' CellSpacing='0' border='0' style="margin-top:5px;">
<tr>
	<td width="50%" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="bImprimir" href="javascript:window.print();"><p class="Button" style="margin-bottom:0px;">Imprimir...</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkExpandirTudo" href="javascript:expandirTudo();"><p class="Button" style="margin-bottom:0px;">Expandir Tudo</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkRecolherTudo" href="javascript:recolherTudo();"><p class="Button" style="margin-bottom:0px;">Recolher Tudo</p></a></td>
</tr>
</table>


<table class="notPrint" width="853" cellSpacing="0" border="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>

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
