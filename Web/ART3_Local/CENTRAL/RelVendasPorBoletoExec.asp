<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelVendasPorBoletoExec.asp
'     ======================================================
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
	
	Const COD_SAIDA_REL_VENDEDOR = "VENDEDOR"
	Const COD_SAIDA_REL_INDICADOR = "INDICADOR"
	Const COD_SAIDA_REL_INDICADORES_DO_VENDEDOR = "INDICADORES_DO_VENDEDOR"
	Const COD_SAIDA_REL_UF = "UF"
	Const COD_SAIDA_REL_ANALISTA_CREDITO = "ANALISTA_CREDITO"
	
	Const COD_ORDENACAO_VL_BOLETO = "ORD_POR_VL_BOLETO"
	Const COD_ORDENACAO_PERC_ATRASO = "ORD_POR_PERC_ATRASO"
	
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
	if Not operacao_permitida(OP_CEN_VENDAS_POR_BOLETO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	alerta = ""

	dim s, s_aux, s_filtro
	dim c_dt_entregue_inicio, c_dt_entregue_termino, c_loja
	dim rb_saida
	dim c_vendedor, c_indicador, c_indicadores_do_vendedor, c_analista
	dim s_nome_vendedor, s_nome_indicador, s_nome_analista, s_nome_loja
	dim c_uf, rb_tipo_cliente, rb_ordenacao
	s_nome_vendedor = ""
	s_nome_indicador = ""
	s_nome_analista = ""
	s_nome_loja = ""

	rb_saida = Trim(Request.Form("rb_saida"))
	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))
	c_loja = retorna_so_digitos(Trim(Request.Form("c_loja")))
	c_vendedor = UCase(Trim(Request.Form("c_vendedor")))
	c_indicador = UCase(Trim(Request.Form("c_indicador")))
	c_indicadores_do_vendedor = UCase(Trim(Request.Form("c_indicadores_do_vendedor")))
	c_analista = UCase(Trim(Request.Form("c_analista")))
	rb_tipo_cliente = Trim(Request.Form("rb_tipo_cliente"))
	rb_ordenacao = Trim(Request.Form("rb_ordenacao"))
	c_uf = Trim(Request.Form("c_uf"))
	
	if alerta = "" then
		if (rb_saida <> COD_SAIDA_REL_VENDEDOR) And _
			(rb_saida <> COD_SAIDA_REL_INDICADOR) And _
			(rb_saida <> COD_SAIDA_REL_INDICADORES_DO_VENDEDOR) And _
			(rb_saida <> COD_SAIDA_REL_UF) And _
			(rb_saida <> COD_SAIDA_REL_ANALISTA_CREDITO) then
			alerta = "O TIPO DE SAÍDA SELECIONADO PARA O RELATÓRIO É INVÁLIDO"
			end if
		end if
	
	if alerta = "" then
		if rb_saida = COD_SAIDA_REL_VENDEDOR then
			if c_vendedor <> "" then
				s = "SELECT nome FROM t_USUARIO WHERE (usuario='" & c_vendedor & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then 
					alerta=texto_add_br(alerta)
					alerta = alerta & "VENDEDOR " & c_vendedor & " NÃO ESTÁ CADASTRADO."
				else
					s_nome_vendedor = Ucase(Trim("" & rs("nome")))
					end if
				end if
			end if
		end if
	
	if alerta = "" then
		if rb_saida = COD_SAIDA_REL_INDICADOR then
			if c_indicador <> "" then
				s = "SELECT razao_social_nome FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido='" & c_indicador & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then 
					alerta=texto_add_br(alerta)
					alerta = alerta & "INDICADOR " & c_indicador & " NÃO ESTÁ CADASTRADO."
				else
					s_nome_indicador = Ucase(Trim("" & rs("razao_social_nome")))
					end if
				end if
			end if
		end if
	
	if alerta = "" then
		if rb_saida = COD_SAIDA_REL_INDICADORES_DO_VENDEDOR then
			if c_indicadores_do_vendedor = "" then
				alerta = "É NECESSÁRIO ESPECIFICAR UM VENDEDOR"
				end if
			
			if c_indicadores_do_vendedor <> "" then
				s = "SELECT nome FROM t_USUARIO WHERE (usuario='" & c_indicadores_do_vendedor & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then 
					alerta=texto_add_br(alerta)
					alerta = alerta & "VENDEDOR " & c_indicadores_do_vendedor & " NÃO ESTÁ CADASTRADO."
				else
					s_nome_vendedor = Ucase(Trim("" & rs("nome")))
					end if
				end if
			end if
		end if
	
	if alerta = "" then
		if rb_saida = COD_SAIDA_REL_ANALISTA_CREDITO then
			if c_analista <> "" then
				s = "SELECT nome FROM t_USUARIO WHERE (usuario='" & c_analista & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then 
					alerta=texto_add_br(alerta)
					alerta = alerta & "ANALISTA " & c_analista & " NÃO ESTÁ CADASTRADO."
				else
					s_nome_analista = Ucase(Trim("" & rs("nome")))
					end if
				end if
			end if
		end if
	
	if alerta = "" then
		if c_loja <> "" then
			s = "SELECT loja, nome, razao_social FROM t_LOJA WHERE (loja = '" & c_loja & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "LOJA " & c_loja & " NÃO ESTÁ CADASTRADA."
			else
				s_nome_loja = Trim("" & rs("nome"))
				end if
			end if
		end if

'	Período de consulta está restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
		
	'	PERÍODO DE ENTREGA
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_entregue_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_entregue_termino
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			if c_dt_entregue_inicio = "" then c_dt_entregue_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
			end if
		
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
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
dim MAX_PEDIDOS_ATRASADOS_POR_LINHA
dim blnAchou
dim s, s_aux, s_cor, s_campo_saida
dim r
dim dt_referencia
dim strSqlCampoSaida
dim s_sql, s_sql_boleto, s_sql_NF
dim s_sql_calc_vl_boleto, s_sql_calc_vl_atrasado, s_sql_vl_NF
dim s_sql_pedidos_atrasados
dim s_sql_devolucoes
dim s_where, s_where_boleto, s_where_NF
dim strColSpanTodasColunas
dim vl_boleto, vl_atrasado, vl_NF
dim vl_total_final_boleto, vl_total_final_NF
dim vl_total_final_atrasado
dim perc_vl_boleto, perc_vl_atrasado
dim qtde_pedidos_boleto, qtde_pedidos_atrasados
dim lista_pedidos_atrasados
dim qtde_total_final_pedidos_boleto, qtde_total_final_pedidos_atrasados
dim cab_table, cab, x
dim vRelat()
dim intIdxVetor, intIdxVetorSelecionado
dim n_linhas_relatorio, n_reg_vetor
dim bln_t_CLIENTE, bln_t_ORCAMENTISTA_E_INDICADOR
dim vPedAtrasado, iPedAtrasado, s_row_pedidos_atrasados, qtde_pedidos_row, s_perc_larg_cel

	bln_t_CLIENTE = False
	bln_t_ORCAMENTISTA_E_INDICADOR = False
	
	if rb_saida = COD_SAIDA_REL_UF then
		MAX_PEDIDOS_ATRASADOS_POR_LINHA = 8
	else
		MAX_PEDIDOS_ATRASADOS_POR_LINHA = 12
		end if
	s_perc_larg_cel = formata_numero(100/MAX_PEDIDOS_ATRASADOS_POR_LINHA, 2)
	s_perc_larg_cel = substitui_caracteres(s_perc_larg_cel, ",", ".")
	
'	OBTÉM A DATA DA CARGA DO ÚLTIMO ARQUIVO DE RETORNO, POIS É ESSA A DATA USADA COMO REFERÊNCIA P/ DETERMINAR OS PAGAMENTOS EM ATRASO
	s_sql = "SELECT TOP 1" & _
				" dt_gravacao_arquivo" & _
			" FROM t_FIN_BOLETO_ARQ_RETORNO" & _
			" WHERE" & _
				" (st_processamento = " & COD_BOLETO_ARQ_RETORNO_ST_PROCESSAMENTO__SUCESSO & ")" & _
				" AND (dt_gravacao_arquivo IS NOT NULL)" & _
			" ORDER BY" & _
				" dt_gravacao_arquivo DESC"
	set r = cn.execute(s_sql)
	if r.Eof then
		dt_referencia = DateAdd("d", -1, Date)
	else
		dt_referencia = r("dt_gravacao_arquivo")
		end if
	
'	QUAL É O CAMPO DE SAÍDA SELECIONADO?
	if rb_saida = COD_SAIDA_REL_VENDEDOR then
		strSqlCampoSaida = "t_PEDIDO__BASE__X.vendedor"
	elseif rb_saida = COD_SAIDA_REL_INDICADOR then
		strSqlCampoSaida = "t_PEDIDO__BASE__X.indicador"
	elseif rb_saida = COD_SAIDA_REL_INDICADORES_DO_VENDEDOR then
		strSqlCampoSaida = "t_PEDIDO__BASE__X.indicador"
	elseif rb_saida = COD_SAIDA_REL_UF then
		bln_t_CLIENTE = True
		strSqlCampoSaida = "t_CLIENTE__X.uf"
	elseif rb_saida = COD_SAIDA_REL_ANALISTA_CREDITO then
		strSqlCampoSaida = "t_PEDIDO__BASE__X.analise_credito_usuario"
		end if
	
'	CRITÉRIOS COMUNS
	s_where = ""
	
'	SOMENTE PEDIDOS ENTREGUES
	s = " (t_PEDIDO__X.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"
	
'	PERÍODO DE ENTREGA
	if IsDate(c_dt_entregue_inicio) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__X.entregue_data >= " & bd_formata_data(StrToDate(c_dt_entregue_inicio)) & ")"
		end if
		
	if IsDate(c_dt_entregue_termino) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__X.entregue_data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		end if
	
'	LOJA
	if c_loja <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__X.loja = '" & c_loja & "')"
		end if

'	TIPO DE CLIENTE
	if rb_tipo_cliente <> "" then
		bln_t_CLIENTE = True
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_CLIENTE__X.tipo = '" & rb_tipo_cliente & "')"
		end if
	
'	UF
	if c_uf <> "" then
	'	QUAL É O CAMPO DE SAÍDA SELECIONADO?
		if rb_saida = COD_SAIDA_REL_VENDEDOR then
			bln_t_CLIENTE = True
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE__X.uf = '" & c_uf & "')"
		elseif rb_saida = COD_SAIDA_REL_INDICADOR then
			bln_t_CLIENTE = True
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE__X.uf = '" & c_uf & "')"
		elseif rb_saida = COD_SAIDA_REL_INDICADORES_DO_VENDEDOR then
			bln_t_CLIENTE = True
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE__X.uf = '" & c_uf & "')"
		elseif rb_saida = COD_SAIDA_REL_UF then
			bln_t_CLIENTE = True
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE__X.uf = '" & c_uf & "')"
		elseif rb_saida = COD_SAIDA_REL_ANALISTA_CREDITO then
			bln_t_CLIENTE = True
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE__X.uf = '" & c_uf & "')"
			end if
		end if
	
'	SAÍDA: VENDEDOR
	if rb_saida = COD_SAIDA_REL_VENDEDOR then
		if c_vendedor <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_PEDIDO__BASE__X.vendedor = '" & c_vendedor & "')"
			end if
		end if
	
'	SAÍDA: INDICADOR
	if rb_saida = COD_SAIDA_REL_INDICADOR then
		if c_indicador <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_PEDIDO__BASE__X.indicador = '" & c_indicador & "')"
			end if
		end if
	
'	SAÍDA: INDICADORES DO VENDEDOR
	if rb_saida = COD_SAIDA_REL_INDICADORES_DO_VENDEDOR then
		if c_indicadores_do_vendedor <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_PEDIDO__BASE__X.indicador IN (SELECT apelido FROM t_ORCAMENTISTA_E_INDICADOR WHERE (vendedor = '" & c_indicadores_do_vendedor & "')))"
			end if
		end if
	
'	SAÍDA: ANALISTA DE CRÉDITO
	if rb_saida = COD_SAIDA_REL_ANALISTA_CREDITO then
		if c_analista <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_PEDIDO__BASE__X.analise_credito_usuario = '" & c_analista & "')"
			end if
		end if
	
'	HAVERÁ UMA CONSULTA APENAS P/ PEDIDOS QUE TENHAM ALGUMA PARCELA EM BOLETO
'	E OUTRA CONSULTA SEM ESSA LIMITAÇÃO
	s_where_boleto = s_where
	s_where_NF = s_where
	
'	PEDIDOS QUE CONTENHAM QUALQUER PARCELA EM BOLETO
	s = " (t_PEDIDO__BASE__X.av_forma_pagto = " & ID_FORMA_PAGTO_BOLETO & ")" & _
		" OR (t_PEDIDO__BASE__X.pu_forma_pagto = " & ID_FORMA_PAGTO_BOLETO & ")" & _
		" OR (t_PEDIDO__BASE__X.pce_forma_pagto_entrada = " & ID_FORMA_PAGTO_BOLETO & ")" & _
		" OR (t_PEDIDO__BASE__X.pce_forma_pagto_prestacao = " & ID_FORMA_PAGTO_BOLETO & ")" & _
		" OR (t_PEDIDO__BASE__X.pse_forma_pagto_prim_prest = " & ID_FORMA_PAGTO_BOLETO & ")" & _
		" OR (t_PEDIDO__BASE__X.pse_forma_pagto_demais_prest = " & ID_FORMA_PAGTO_BOLETO & ")"
	if s_where_boleto <> "" then s_where_boleto = s_where_boleto & " AND"
	s_where_boleto = s_where_boleto & " (" & s & ")"
	
'	MONTA SQL DOS CÁLCULOS DE VALOR
	s_sql_calc_vl_boleto = _
		"SELECT " & _
			SCHEMA_BD & ".SqlClrUtilCalculaValorMeioPagtoEspecificadoFormaPagtoPedido(" & _
			ID_FORMA_PAGTO_BOLETO & "," & _
			" t_PEDIDO__BASE.vl_total_NF," & _
			" t_PEDIDO__BASE.tipo_parcelamento," & _
			" t_PEDIDO__BASE.av_forma_pagto," & _
			" t_PEDIDO__BASE.pu_forma_pagto," & _
			" t_PEDIDO__BASE.pu_valor," & _
			" t_PEDIDO__BASE.pc_qtde_parcelas," & _
			" t_PEDIDO__BASE.pc_valor_parcela," & _
			" t_PEDIDO__BASE.pc_maquineta_qtde_parcelas," & _
			" t_PEDIDO__BASE.pc_maquineta_valor_parcela," & _
			" t_PEDIDO__BASE.pce_forma_pagto_entrada," & _
			" t_PEDIDO__BASE.pce_forma_pagto_prestacao," & _
			" t_PEDIDO__BASE.pce_entrada_valor," & _
			" t_PEDIDO__BASE.pce_prestacao_qtde," & _
			" t_PEDIDO__BASE.pce_prestacao_valor," & _
			" t_PEDIDO__BASE.pse_forma_pagto_prim_prest," & _
			" t_PEDIDO__BASE.pse_forma_pagto_demais_prest," & _
			" t_PEDIDO__BASE.pse_prim_prest_valor," & _
			" t_PEDIDO__BASE.pse_demais_prest_qtde," & _
			" t_PEDIDO__BASE.pse_demais_prest_valor" & _
			") AS vl_boleto" & _
		" FROM t_PEDIDO" & _
			" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
		" WHERE" & _
			" (t_PEDIDO.pedido = t_PEDIDO__X.pedido)"
	
'	O campo t_FIN_FLUXO_CAIXA.st_boleto_baixado não está sendo considerado na consulta porque os boletos atrasados são baixados pelo banco por decurso de prazo.
	s_sql_calc_vl_atrasado = _
		"SELECT" & _
			" Sum(tFC.valor) AS vl_acumulado" & _
		" FROM t_FIN_FLUXO_CAIXA tFC" & _
			" INNER JOIN t_FIN_BOLETO_ITEM_RATEIO tBIR ON (tFC.ctrl_pagto_id_parcela = tBIR.id_boleto_item)" & _
			" INNER JOIN t_PEDIDO tP ON (tBIR.pedido = tP.pedido)" & _
		" WHERE" & _
			" (tFC.dt_competencia <= " & bd_monta_data(dt_referencia) & ")" & _
			" AND (tFC.ctrl_pagto_modulo = " & CTRL_PAGTO_MODULO__BOLETO & ")" & _
			" AND (tFC.st_confirmacao_pendente = 1)" & _
			" AND (tFC.st_sem_efeito = 0)" & _
			" AND (tFC.ctrl_pagto_status <> " & CTRL_PAGTO_STATUS__BOLETO_PAGO_CHEQUE_VINCULADO & ")" & _
			" AND (tBIR.pedido = t_PEDIDO__X.pedido)"
	
	s_sql_pedidos_atrasados = _
		"SELECT DISTINCT" & _
			" 'S'" & _
		" FROM t_FIN_FLUXO_CAIXA tFC" & _
			" INNER JOIN t_FIN_BOLETO_ITEM_RATEIO tBIR ON (tFC.ctrl_pagto_id_parcela = tBIR.id_boleto_item)" & _
			" INNER JOIN t_PEDIDO tP ON (tBIR.pedido = tP.pedido)" & _
		" WHERE" & _
			" (tFC.dt_competencia <= " & bd_monta_data(dt_referencia) & ")" & _
			" AND (tFC.ctrl_pagto_modulo = " & CTRL_PAGTO_MODULO__BOLETO & ")" & _
			" AND (tFC.st_confirmacao_pendente = 1)" & _
			" AND (tFC.st_sem_efeito = 0)" & _
			" AND (tFC.ctrl_pagto_status <> " & CTRL_PAGTO_STATUS__BOLETO_PAGO_CHEQUE_VINCULADO & ")" & _
			" AND (tBIR.pedido = t_PEDIDO__X.pedido)"
	
'	VALOR EM DEVOLUÇÕES
	s_sql_devolucoes = _
		"SELECT" & _
			" Sum(qtde*preco_NF) AS vl_devolucao" & _
		" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
		" WHERE" & _
			" (t_PEDIDO_ITEM_DEVOLVIDO.pedido = t_PEDIDO__X.pedido)"
	
'	VALOR NF
	s_sql_vl_NF = _
		"SELECT" & _
			" Sum(qtde*preco_NF) AS vl_total_NF" & _
		" FROM t_PEDIDO_ITEM tPI" & _
			" INNER JOIN t_PEDIDO tP ON (tPI.pedido=tP.pedido)" & _
		" WHERE" & _
			" (tP.pedido = t_PEDIDO__X.pedido)"
	
'	MONTA CONSULTA COMPLETA
	s_sql_boleto = _
		"SELECT " & _
			strSqlCampoSaida & "," & _
			" t_PEDIDO__X.pedido," & _
			" Coalesce((" & s_sql_calc_vl_boleto & "), 0) AS vl_boleto," & _
			" Coalesce((" & s_sql_calc_vl_atrasado & "), 0) AS vl_atrasado," & _
			" Coalesce((" & s_sql_pedidos_atrasados & "), 'N') AS flag_pedido_atrasado," & _
			" Coalesce((" & s_sql_devolucoes & "), 0) AS vl_devolucao" & _
		" FROM t_PEDIDO AS t_PEDIDO__X" & _
			" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE__X ON (t_PEDIDO__X.pedido_base=t_PEDIDO__BASE__X.pedido)"
	
	s_sql_NF = _
		"SELECT " & _
			strSqlCampoSaida & "," & _
			" t_PEDIDO__X.pedido," & _
			" Coalesce((" & s_sql_devolucoes & "), 0) AS vl_devolucao," & _
			" Coalesce((" & s_sql_vl_NF & "), 0) AS vl_NF" & _
		" FROM t_PEDIDO AS t_PEDIDO__X" & _
			" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE__X ON (t_PEDIDO__X.pedido_base=t_PEDIDO__BASE__X.pedido)"
	
	if bln_t_CLIENTE then
		s_sql_boleto = s_sql_boleto & _
			" INNER JOIN t_CLIENTE AS t_CLIENTE__X ON (t_PEDIDO__BASE__X.id_cliente=t_CLIENTE__X.id)"
		s_sql_NF = s_sql_NF & _
			" INNER JOIN t_CLIENTE AS t_CLIENTE__X ON (t_PEDIDO__BASE__X.id_cliente=t_CLIENTE__X.id)"
		end if
	
	if bln_t_ORCAMENTISTA_E_INDICADOR then
		s_sql_boleto = s_sql_boleto & _
			" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR AS t_ORCAMENTISTA_E_INDICADOR__X ON (t_PEDIDO__BASE__X.indicador=t_ORCAMENTISTA_E_INDICADOR__X.apelido)"
		s_sql_NF = s_sql_NF & _
			" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR AS t_ORCAMENTISTA_E_INDICADOR__X ON (t_PEDIDO__BASE__X.indicador=t_ORCAMENTISTA_E_INDICADOR__X.apelido)"
		end if
	
	s_sql_boleto = s_sql_boleto & _
			" WHERE" & _
				"(" & s_where_boleto & ")" & _
			" GROUP BY " & _
				strSqlCampoSaida & "," & _
				" t_PEDIDO__X.pedido" & _
			" ORDER BY " & _
				strSqlCampoSaida & "," & _
				" t_PEDIDO__X.pedido"
	
	s_sql_NF = s_sql_NF & _
			" WHERE" & _
				"(" & s_where_NF & ")" & _
			" GROUP BY " & _
				strSqlCampoSaida & "," & _
				" t_PEDIDO__X.pedido" & _
			" ORDER BY " & _
				strSqlCampoSaida & "," & _
				" t_PEDIDO__X.pedido"
	
'	CABEÇALHO
	cab_table = "<table cellspacing='0' cellpadding='0'>" & chr(13)
	cab = "	<tr style='background:azure' nowrap>" & chr(13)
	
	cab = cab & _
			"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13)
	
'	QUAL É O CAMPO DE SAÍDA SELECIONADO?
	strColSpanTodasColunas = "colspan='8'"
	if rb_saida = COD_SAIDA_REL_VENDEDOR then
		cab = cab & _
			"		<td class='MDTE tdColSaida' align='left' valign='bottom' nowrap><span class='R'>Vendedor</span></td>" & chr(13)
	elseif rb_saida = COD_SAIDA_REL_INDICADOR then
		cab = cab & _
			"		<td class='MDTE tdColSaida' align='left' valign='bottom' nowrap><span class='R'>Indicador</span></td>" & chr(13)
	elseif rb_saida = COD_SAIDA_REL_INDICADORES_DO_VENDEDOR then
		cab = cab & _
			"		<td class='MDTE tdColSaida' align='left' valign='bottom' nowrap><span class='R'>Indicador</span></td>" & chr(13)
	elseif rb_saida = COD_SAIDA_REL_UF then
		cab = cab & _
			"		<td class='MDTE tdColSaidaUF' align='center' valign='bottom' nowrap><span class='Rc'>UF</span></td>" & chr(13)
	elseif rb_saida = COD_SAIDA_REL_ANALISTA_CREDITO then
		cab = cab & _
			"		<td class='MDTE tdColSaida' align='left' valign='bottom' nowrap><span class='R'>Analista</span></td>" & chr(13)
		end if
	
	cab = cab & _
			"		<td class='MTD tdColValor' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>&nbsp;</span><br /><span class='Rd' style='font-weight:bold;'>VL (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
			"		<td class='MTD tdColValor' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL</span><br /><span class='Rd' style='font-weight:bold;'>Boleto (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
			"		<td class='MTD tdColPerc' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>%</span><br /><span class='Rd' style='font-weight:bold;'>Boleto</span></td>" & chr(13) & _
			"		<td class='MTD tdColQtde' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>Qtde</span><br /><span class='Rd' style='font-weight:bold;'>Pedido</span></td>" & chr(13) & _
			"		<td class='MTD tdColValor' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL</span><br /><span class='Rd' style='font-weight:bold;'>Atraso (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
			"		<td class='MTD tdColPerc' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>%</span><br /><span class='Rd' style='font-weight:bold;'>Atraso</span></td>" & chr(13) & _
			"		<td class='MTD tdColQtde' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>Qtde</span><br /><span class='Rd' style='font-weight:bold;'>Pedido</span></td>" & chr(13) & _
			"		<td align='left' style='background:#FFFFFF;'>&nbsp;</td>" & chr(13) & _
			"	</tr>" & chr(13)
	
	redim vRelat(1)
	for intIdxVetor = Lbound(vRelat) to Ubound(vRelat)
		set vRelat(intIdxVetor) = New cl_VINTE_COLUNAS
		vRelat(intIdxVetor).CampoOrdenacao = ""
		vRelat(intIdxVetor).c1 = ""
		next
	
'	INICIALIZAÇÃO
	vl_total_final_boleto = 0
	vl_total_final_NF = 0
	vl_total_final_atrasado = 0
	qtde_total_final_pedidos_boleto = 0
	qtde_total_final_pedidos_atrasados = 0
	
'	ARMAZENA OS DADOS EM VETOR PARA FAZER A CONSOLIDAÇÃO DOS VALORES
'	DESCRIÇÃO DO VETOR:
'		c1: indica se há dados na linha ('S': há dados, '': não há dados)
'		c2: campo de saída do relatório (vendedor, indicador, uf, etc)
'		c3: vl_boleto
'		c4: perc_vl_boleto
'		c5: qtde_pedidos_boleto
'		c6: vl_atrasado
'		c7: perc_vl_atrasado
'		c8: qtde_pedidos_atrasados
'		c9: lista dos pedidos atrasados
'		c10: vl_NF
	set r = cn.execute(s_sql_boleto)
	do while Not r.Eof
		
		if rb_saida = COD_SAIDA_REL_VENDEDOR then
			s_campo_saida = Trim("" & r("vendedor"))
		elseif rb_saida = COD_SAIDA_REL_INDICADOR then
			s_campo_saida = Trim("" & r("indicador"))
		elseif rb_saida = COD_SAIDA_REL_INDICADORES_DO_VENDEDOR then
			s_campo_saida = Trim("" & r("indicador"))
		elseif rb_saida = COD_SAIDA_REL_UF then
			s_campo_saida = Trim("" & r("uf"))
		elseif rb_saida = COD_SAIDA_REL_ANALISTA_CREDITO then
			s_campo_saida = Trim("" & r("analise_credito_usuario"))
			end if
		
		blnAchou = False
		for intIdxVetor = Lbound(vRelat) to Ubound(vRelat)
			if vRelat(intIdxVetor).c1 = "S" then
				if UCase(vRelat(intIdxVetor).c2) = UCase(s_campo_saida) then
					intIdxVetorSelecionado = intIdxVetor
					blnAchou = True
					exit for
					end if
				end if
			next
		
		if Not blnAchou then
			if vRelat(UBound(vRelat)).c1 = "S" then
				redim preserve vRelat(Ubound(vRelat)+1)
				set vRelat(Ubound(vRelat)) = New cl_VINTE_COLUNAS
				end if
			intIdxVetorSelecionado = Ubound(vRelat)
			with vRelat(intIdxVetorSelecionado)
				.CampoOrdenacao = ""
				.c1 = "S"
				.c2 = s_campo_saida
				.c3 = 0
				.c4 = 0
				.c5 = 0
				.c6 = 0
				.c7 = 0
				.c8 = 0
				.c9 = ""
				.c10 = 0
				end with
			end if
		
		with vRelat(intIdxVetorSelecionado)
		'	VALOR BOLETO
			vl_boleto = r("vl_boleto")-r("vl_devolucao")
			if vl_boleto < 0 then vl_boleto = 0
			.c3 = .c3 + vl_boleto
			vl_total_final_boleto = vl_total_final_boleto + vl_boleto
			
		'	QTDE DE PEDIDOS (BOLETO)
			.c5 = .c5 + 1
			qtde_total_final_pedidos_boleto = qtde_total_final_pedidos_boleto + 1
			
		'	EM ATRASO?
			if Trim("" & r("flag_pedido_atrasado")) = "S" then
			'	VALOR EM ATRASO
				vl_atrasado = r("vl_atrasado")-r("vl_devolucao")
				if vl_atrasado < 0 then vl_atrasado = 0
				.c6 = .c6 + vl_atrasado
				vl_total_final_atrasado = vl_total_final_atrasado + vl_atrasado
			'	QTDE PEDIDOS EM ATRASO
				.c8 = .c8 + 1
				qtde_total_final_pedidos_atrasados = qtde_total_final_pedidos_atrasados + 1
			'	LISTA DOS PEDIDOS ATRASADOS
				if Trim(.c9) <> "" then .c9 = .c9 & "|"
				.c9 = .c9 & Trim("" & r("pedido"))
				end if
			end with
		
		r.MoveNext
		loop
	
	if r.State <> 0 then r.Close
	set r=nothing
	
	set r = cn.execute(s_sql_NF)
	do while Not r.Eof
		if rb_saida = COD_SAIDA_REL_VENDEDOR then
			s_campo_saida = Trim("" & r("vendedor"))
		elseif rb_saida = COD_SAIDA_REL_INDICADOR then
			s_campo_saida = Trim("" & r("indicador"))
		elseif rb_saida = COD_SAIDA_REL_INDICADORES_DO_VENDEDOR then
			s_campo_saida = Trim("" & r("indicador"))
		elseif rb_saida = COD_SAIDA_REL_UF then
			s_campo_saida = Trim("" & r("uf"))
		elseif rb_saida = COD_SAIDA_REL_ANALISTA_CREDITO then
			s_campo_saida = Trim("" & r("analise_credito_usuario"))
			end if
		
		blnAchou = False
		for intIdxVetor = Lbound(vRelat) to Ubound(vRelat)
			if vRelat(intIdxVetor).c1 = "S" then
				if UCase(vRelat(intIdxVetor).c2) = UCase(s_campo_saida) then
					intIdxVetorSelecionado = intIdxVetor
					blnAchou = True
					exit for
					end if
				end if
			next
		
		if blnAchou then
			with vRelat(intIdxVetorSelecionado)
			'	VALOR NF
				vl_NF = r("vl_NF")-r("vl_devolucao")
				if vl_NF < 0 then vl_NF = 0
				.c10 = .c10 + vl_NF
				vl_total_final_NF = vl_total_final_NF + vl_NF
				end with
			end if
		
		r.MoveNext
		loop
	
	if r.State <> 0 then r.Close
	set r=nothing
	
	n_linhas_relatorio = 0
	for intIdxVetor = LBound(vRelat) to UBound(vRelat)
		with vRelat(intIdxVetor)
			if Trim("" & .c1) = "S" then
				n_linhas_relatorio = n_linhas_relatorio + 1
				
				vl_boleto = .c3
				vl_atrasado = .c6
				
				if vl_total_final_boleto = 0 then
					perc_vl_boleto = 0
				else
					perc_vl_boleto = 100 * (vl_boleto / vl_total_final_boleto)
					end if
				.c4 = perc_vl_boleto
				
			'	O PERCENTUAL EM ATRASO É CALCULADO SOBRE O VALOR DO BOLETO DA RESPECTIVA LINHA
				if vl_boleto = 0 then
					perc_vl_atrasado = 0
				else
					perc_vl_atrasado = 100 * (vl_atrasado / vl_boleto)
					end if
				.c7 = perc_vl_atrasado
				
				if rb_ordenacao = COD_ORDENACAO_PERC_ATRASO then
					.CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_perc(perc_vl_atrasado)),10) & "|" & normaliza_codigo(retorna_so_digitos(formata_moeda(vl_boleto)), 20) & "|" & .c2
				else
					.CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_boleto)), 20) & "|" & .c2
					end if
				end if
			end with
		next
	
'	ORDENA O VETOR COM RESULTADOS
	ordena_cl_vinte_colunas vRelat, 1, Ubound(vRelat)
	
	n_reg_vetor = 0
	if n_linhas_relatorio > 0 then
		x = cab_table
		x = x & cab
		
		for intIdxVetor = Ubound(vRelat) to 1 step -1
			if Trim("" & vRelat(intIdxVetor).c1) = "S" then
			'	CONTAGEM
				n_reg_vetor = n_reg_vetor + 1
				
				x = x & "	<tr nowrap>" & chr(13)
				
				with vRelat(intIdxVetor)
					s_campo_saida = .c2
					vl_boleto = .c3
					perc_vl_boleto = .c4
					qtde_pedidos_boleto = .c5
					vl_atrasado = .c6
					perc_vl_atrasado = .c7
					qtde_pedidos_atrasados = .c8
					lista_pedidos_atrasados = .c9
					vl_NF = .c10
					end with
				
				s_cor="black"
				if vl_boleto < 0 then s_cor="red"
				
				x = x & "		<td align='right' valign='bottom' nowrap><span class='Rd' style='margin-right:2px;'>" & Cstr(n_reg_vetor) & ".</span></td>" & chr(13)
				
				if rb_saida = COD_SAIDA_REL_VENDEDOR then
					s = UCase(s_campo_saida)
					s_aux = x_usuario(s)
					if (s <> "") And (s_aux <> "") then s = s & " - " & s_aux
					if s = "" then s = "&nbsp;"
					x = x & "		<td class='MDTE tdColSaida' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
				elseif rb_saida = COD_SAIDA_REL_INDICADOR then
					s = UCase(s_campo_saida)
					s_aux = x_orcamentista_e_indicador(s)
					if (s <> "") And (s_aux <> "") then s = s & " - " & s_aux
					if s = "" then s = "&nbsp;"
					x = x & "		<td class='MDTE tdColSaida' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
				elseif rb_saida = COD_SAIDA_REL_INDICADORES_DO_VENDEDOR then
					s = UCase(s_campo_saida)
					s_aux = x_orcamentista_e_indicador(s)
					if (s <> "") And (s_aux <> "") then s = s & " - " & s_aux
					if s = "" then s = "&nbsp;"
					x = x & "		<td class='MDTE tdColSaida' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
				elseif rb_saida = COD_SAIDA_REL_UF then
					s = UCase(s_campo_saida)
					if s = "" then s = "&nbsp;"
					x = x & "		<td class='MDTE tdColSaidaUF' align='center' valign='bottom'><span class='Cnc' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
				elseif rb_saida = COD_SAIDA_REL_ANALISTA_CREDITO then
					s = UCase(s_campo_saida)
					s_aux = x_usuario(s)
					if (s <> "") And (s_aux <> "") then s = s & " - " & s_aux
					if s = "" then s = "&nbsp;"
					x = x & "		<td class='MDTE tdColSaida' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
					end if
				
			 '> VALOR NF
				x = x & "		<td class='MTD tdColValor' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF) & "</span></td>" & chr(13)
				
			 '> VALOR BOLETO
				x = x & "		<td class='MTD tdColValor' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_boleto) & "</span></td>" & chr(13)
				
			 '> PERCENTUAL RELATIVO AO VALOR TOTAL
				x = x & "		<td class='MTD tdColPerc' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_vl_boleto) & "%" & "</span></td>" & chr(13)
				
			 '> QTDE PEDIDOS
				x = x & "		<td class='MTD tdColQtde' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_pedidos_boleto) & "</span></td>" & chr(13)
				
			 '> VALOR ATRASADO
				x = x & "		<td class='MTD tdColValor' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_atrasado) & "</span></td>" & chr(13)
				
			 '> PERCENTUAL RELATIVO AO VALOR TOTAL
				x = x & "		<td class='MTD tdColPerc' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_vl_atrasado) & "%" & "</span></td>" & chr(13)
				
			 '> QTDE PEDIDOS
				x = x & "		<td class='MTD tdColQtde' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_pedidos_atrasados) & "</span></td>" & chr(13)
				
			'	HAVERÁ LINHA LISTANDO OS PEDIDOS ATRASADOS?
				if qtde_pedidos_atrasados > 0 then
					x = x & "<td align='left' valign='bottom' class='notPrint'>&nbsp;<a name='bExibeOcultaPedidosAtrasados_" & Cstr(n_reg_vetor) & "' id='bExibeOcultaPedidosAtrasados_" & Cstr(n_reg_vetor) & "' href='javascript:fExibeOcultaLinhaPedidosAtrasados(" & chr(34) & Cstr(n_reg_vetor) & chr(34) & ");' title='exibe ou oculta a relação dos pedidos atrasados'><img src='../botao/view_bottom.png' border='0'></a></td>" & chr(13)
				else
					x = x & "<td align='left' valign='bottom' class='notPrint'>&nbsp;</td>" & chr(13)
					end if
				
				x = x & "	</tr>" & chr(13)
				
			 '> LISTA DE PEDIDOS ATRASADOS?
				if Trim(lista_pedidos_atrasados) <> "" then
					s_row_pedidos_atrasados = "				<tr>" & chr(13)
					qtde_pedidos_row = 0
					vPedAtrasado = Split(lista_pedidos_atrasados, "|")
					for iPedAtrasado = LBound(vPedAtrasado) to UBound(vPedAtrasado)
						if Trim(vPedAtrasado(iPedAtrasado)) <> "" then
							qtde_pedidos_row = qtde_pedidos_row + 1
							if qtde_pedidos_row > MAX_PEDIDOS_ATRASADOS_POR_LINHA then
								qtde_pedidos_row = 1
								s_row_pedidos_atrasados = s_row_pedidos_atrasados & _
									"				</tr>" & chr(13) & _
									"				<tr>" & chr(13)
								end if
							s_row_pedidos_atrasados = s_row_pedidos_atrasados & _
								"					<td class='C' align='left' width='" & s_perc_larg_cel & "%'>" & monta_link_pedido(Trim(vPedAtrasado(iPedAtrasado))) & "</td>"
							end if
						next
				'	COMPLETA O RESTANTE DA LINHA
					for iPedAtrasado=(qtde_pedidos_row+1) to MAX_PEDIDOS_ATRASADOS_POR_LINHA
						s_row_pedidos_atrasados = s_row_pedidos_atrasados & _
							"					<td align='left' width='" & s_perc_larg_cel & "%'>&nbsp;</td>" & chr(13)
						next
					s_row_pedidos_atrasados = s_row_pedidos_atrasados & "				</tr>" & chr(13)
					
					x = x & "	<tr class='TR_EXPANSIVEL' style='display:none;' id='TR_PEDIDOS_ATRASADOS_" & CStr(n_reg_vetor) & "'>" & chr(13) & _
							"		<td>&nbsp;</td>" & chr(13) & _
							"		<td " & strColSpanTodasColunas & " class='MC MD' align='left'>" & chr(13) & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
							"				<tr>" & chr(13) & _
							"					<td style='width:20px;' class='MD' align='left'>&nbsp;</td>" & chr(13) & _
							"					<td align='right'>" & chr(13) & _
							"						<table width='100%' style='margin-left:3px;' cellspacing='0' cellpadding='0'>" & chr(13) & _
							"							<tr>" & chr(13) & _
							"								<td colspan='" & MAX_PEDIDOS_ATRASADOS_POR_LINHA & "' class='Rf' align='left'>PEDIDOS ATRASADOS</td>" & chr(13) & _
							"							</tr>" & chr(13) & _
														s_row_pedidos_atrasados & _
							"						</table>" & chr(13) & _
							"					</td>" & chr(13) & _
							"				</tr>" & chr(13) & _
							"			</table>" & chr(13) & _
							"		</td>" & chr(13) & _
							"	</tr>" & chr(13)
					end if
				
				if (n_reg_vetor mod 100) = 0 then
					Response.Write x
					x = ""
					end if
				end if ' if Trim("" & vRelat(intIdxVetor).c1) = "S"
			next
		
	'	MOSTRA TOTAL
		if vl_total_final_boleto = 0 then
			perc_vl_atrasado = 0
		else
			perc_vl_atrasado = 100 * (vl_total_final_atrasado / vl_total_final_boleto)
			end if
		s_cor="black"
		if vl_total_final_boleto < 0 then s_cor="red"
		x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
				"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
				"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
				"TOTAL:</span></td>" & chr(13) & _
				"		<td class='MTB tdColValor' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_final_NF) & "</span></td>" & chr(13) & _
				"		<td class='MTB tdColValor' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_final_boleto) & "</span></td>" & chr(13) & _
				"		<td class='MTB tdColPerc' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & "&nbsp;" & "</span></td>" & chr(13) & _
				"		<td class='MTB tdColQtde' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_total_final_pedidos_boleto) & "</span></td>" & chr(13) & _
				"		<td class='MTB tdColValor' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_final_atrasado) & "</span></td>" & chr(13) & _
				"		<td class='MTB tdColPerc' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_vl_atrasado) & "%" & "</span></td>" & chr(13) & _
				"		<td class='MTBD tdColQtde' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_total_final_pedidos_atrasados) & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if


  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_linhas_relatorio = 0 then
		x = cab_table & cab
		x = x & "	<tr nowrap>" & chr(13) & _
				"		<td align='left'>&nbsp;</td>" & chr(13) & _
				"		<td class='MT ALERTA' " & strColSpanTodasColunas & " align='center'><span class='ALERTA'>&nbsp;NENHUM REGISTRO SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</table>" & chr(13)
	
	Response.write x
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


<style type="text/css">
html
{
	overflow-y: scroll;
}
.style1
{
	height: 46px;
}
.tdColSaida
{
	width: 400px;
}
.tdColSaidaUF
{
	width: 60px;
}
.tdColValor
{
	width: 90px;
}
.tdColPerc
{
	width: 60px;
}
.tdColQtde
{
	width:60px;
}
</style>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function expandirTudo() {
	$('.TR_EXPANSIVEL').each(function() {
		$(this).css('display', '');
		$(this).addClass('EXPANDIDO');
	});
}

function recolherTudo() {
	$('.TR_EXPANSIVEL').each(function() {
		$(this).css('display', 'none');
		$(this).removeClass('EXPANDIDO');
	});
}

function fExibeOcultaLinhaPedidosAtrasados(indice_row) {
	var row_MORE_INFO;

	row_MORE_INFO = document.getElementById("TR_PEDIDOS_ATRASADOS_" + indice_row);
	if (row_MORE_INFO.style.display.toString() == "none") {
		row_MORE_INFO.style.display = "";
		$(row_MORE_INFO).addClass('EXPANDIDO');
	}
	else {
		row_MORE_INFO.style.display = "none";
		$(row_MORE_INFO).removeClass('EXPANDIDO');
	}
}

function restauraExpandidos() {
	var i, listaExpandidos, vExpandidos;
	listaExpandidos = fPED.c_row_expandidas.value;
	if (listaExpandidos != "") {
		vExpandidos = listaExpandidos.split("|");
		for (i = 0; i < vExpandidos.length; i++) {
			$('#' + vExpandidos[i]).css('display', '');
			$('#' + vExpandidos[i]).addClass('EXPANDIDO');
		}
	}
}

function fPEDConsulta(id_pedido) {
	window.status = "Aguarde ...";
	fPED.c_row_expandidas.value = "";
	$('.EXPANDIDO').each(function() {
		if (fPED.c_row_expandidas.value != "") fPED.c_row_expandidas.value += "|";
		fPED.c_row_expandidas.value += $(this).attr('id');
	});
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">


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
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>



<% else %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';restauraExpandidos();" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<input type="hidden" name="c_row_expandidas" id="c_row_expandidas" value="">
</form>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="rb_saida" id="rb_saida" value="<%=rb_saida%>">
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="<%=c_dt_entregue_inicio%>">
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">
<input type="hidden" name="c_indicadores_do_vendedor" id="c_indicadores_do_vendedor" value="<%=c_indicadores_do_vendedor%>">
<input type="hidden" name="c_analista" id="c_analista" value="<%=c_analista%>">
<input type="hidden" name="rb_tipo_cliente" id="rb_tipo_cliente" value="<%=rb_tipo_cliente%>">
<input type="hidden" name="rb_ordenacao" id="rb_ordenacao" value="<%=rb_ordenacao%>">
<input type="hidden" name="c_uf" id="c_uf" value="<%=c_uf%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="960" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" class="style1"><span class="PEDIDO">Relatório de Vendas por Boleto</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<%
	s_filtro = "<table width='960' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>"
	
	s = ""
	s_aux = c_dt_entregue_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_entregue_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Período de Entrega:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	s = c_loja
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Loja:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"

	if rb_saida = COD_SAIDA_REL_VENDEDOR then
		s = c_vendedor
		if s = "" then 
			s = "todos"
		else
			if (s_nome_vendedor <> "") And (s_nome_vendedor <> c_vendedor) then s = s & " (" & s_nome_vendedor & ")"
			end if
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>Vendedor:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	if rb_saida = COD_SAIDA_REL_INDICADOR then
		s = c_indicador
		if s = "" then 
			s = "todos"
		else
			s = s & " (" & x_orcamentista_e_indicador(c_indicador) & ")"
			end if
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>Indicador:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	if rb_saida = COD_SAIDA_REL_INDICADORES_DO_VENDEDOR then
		s = c_indicadores_do_vendedor
		if s = "" then 
			s = "todos"
		else
			if (s_nome_vendedor <> "") And (s_nome_vendedor <> c_vendedor) then s = s & " (" & s_nome_vendedor & ")"
			end if
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>Indicadores do Vendedor:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	if rb_saida = COD_SAIDA_REL_UF then
		s = "todas"
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>UF:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	if rb_saida = COD_SAIDA_REL_ANALISTA_CREDITO then
		s = c_analista
		if s = "" then
			s = "todos"
		else
			s = s & " (" & x_usuario(c_analista) & ")"
			end if
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>Analista:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	if rb_tipo_cliente = ID_PF then
		s = "PF"
	elseif rb_tipo_cliente = ID_PJ then
		s = "PJ"
	else
		s = "todos"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Tipo de Cliente:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	s = c_uf
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>UF:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	if rb_ordenacao = COD_ORDENACAO_PERC_ATRASO then
		s = "coluna [% Atraso]"
	else
		s = "coluna [VL Boleto (" & SIMBOLO_MONETARIO & ")]"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Ordenação:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="960" cellpadding="0" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width='960' cellpadding='0' cellspacing='0' border='0' style="margin-top:5px;">
<tr>
	<td width="50%" align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkImprimir" href="javascript:window.print();"><p class="Button" style="margin-bottom:0px;">Imprimir...</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkExpandirTudo" href="javascript:expandirTudo();"><p class="Button" style="margin-bottom:0px;">Expandir Tudo</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkRecolherTudo" href="javascript:recolherTudo();"><p class="Button" style="margin-bottom:0px;">Recolher Tudo</p></a></td>
</tr>
</table>

<br />

<table class="notPrint" width="960" cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
