<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L P E D I D O S M C R I T E X E C . A S P
'     =================================================================
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
	
	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if (Not operacao_permitida(OP_LJA_REL_MULTICRITERIO_PEDIDOS_ANALITICO, s_lista_operacoes_permitidas)) And _
	   (Not operacao_permitida(OP_LJA_REL_MULTICRITERIO_PEDIDOS_SINTETICO, s_lista_operacoes_permitidas)) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	dim s, s_aux, s_aux_dti, s_aux_dtf, s_filtro, flag_ok, cadastrado
	dim ckb_st_entrega_esperar, ckb_st_entrega_split, ckb_st_entrega_exceto_cancelados
	dim ckb_st_entrega_separar_sem_marc, ckb_st_entrega_separar_com_marc
	dim ckb_st_entrega_a_entregar_sem_marc, ckb_st_entrega_a_entregar_com_marc, ckb_pedido_nao_recebido_pelo_cliente, ckb_pedido_recebido_pelo_cliente
	dim c_dt_coleta_a_separar_inicio, c_dt_coleta_a_separar_termino, c_dt_coleta_st_a_entregar_inicio, c_dt_coleta_st_a_entregar_termino
	dim ckb_st_entrega_entregue, c_dt_entregue_inicio, c_dt_entregue_termino
	dim ckb_st_entrega_cancelado, c_dt_cancelado_inicio, c_dt_cancelado_termino
	dim ckb_st_pagto_pago, ckb_st_pagto_nao_pago, ckb_st_pagto_pago_parcial
	dim ckb_periodo_cadastro, c_dt_cadastro_inicio, c_dt_cadastro_termino
	dim ckb_entrega_marcada_para, c_dt_entrega_inicio, c_dt_entrega_termino
	dim ckb_produto, c_fabricante, c_produto, c_grupo, v_grupos
	dim c_loja
	dim c_cliente_cnpj_cpf, c_cliente_uf
	dim c_transportadora
	dim ckb_visanet
	dim ckb_perc_RT, c_perc_RT
	dim ckb_analise_credito_st_inicial, ckb_analise_credito_pendente_vendas, ckb_analise_credito_pendente_endereco, ckb_analise_credito_pendente, ckb_analise_credito_pendente_cartao
	dim ckb_analise_credito_ok, ckb_analise_credito_ok_aguardando_deposito, ckb_analise_credito_ok_deposito_aguardando_desbloqueio
	dim ckb_analise_credito_pendente_pagto_antecipado_boleto, ckb_analise_credito_ok_aguardando_pagto_boleto_av
	dim ckb_entrega_imediata_sim, ckb_entrega_imediata_nao, c_dt_previsao_entrega_inicio, c_dt_previsao_entrega_termino
	dim op_forma_pagto, c_forma_pagto_qtde_parc
	dim c_vendedor, c_indicador
	dim ckb_indicador_preenchido, ckb_indicador_nao_preenchido
	dim data_pedido
    dim c_pedido_origem, c_grupo_pedido_origem, c_empresa
	dim c_FormFieldValues
    dim blnMostraMotivoCancelado, c_cancelados_ordena
	
	alerta = ""

	ckb_st_entrega_exceto_cancelados = Trim(Request.Form("ckb_st_entrega_exceto_cancelados"))
	ckb_st_entrega_esperar = Trim(Request.Form("ckb_st_entrega_esperar"))
	ckb_st_entrega_split = Trim(Request.Form("ckb_st_entrega_split"))
	ckb_st_entrega_separar_sem_marc = Trim(Request.Form("ckb_st_entrega_separar_sem_marc"))
	ckb_st_entrega_separar_com_marc = Trim(Request.Form("ckb_st_entrega_separar_com_marc"))
	c_dt_coleta_a_separar_inicio = Trim(Request.Form("c_dt_coleta_a_separar_inicio"))
	c_dt_coleta_a_separar_termino = Trim(Request.Form("c_dt_coleta_a_separar_termino"))
	ckb_st_entrega_a_entregar_sem_marc = Trim(Request.Form("ckb_st_entrega_a_entregar_sem_marc"))
	ckb_st_entrega_a_entregar_com_marc = Trim(Request.Form("ckb_st_entrega_a_entregar_com_marc"))
	c_dt_coleta_st_a_entregar_inicio = Trim(Request.Form("c_dt_coleta_st_a_entregar_inicio"))
	c_dt_coleta_st_a_entregar_termino = Trim(Request.Form("c_dt_coleta_st_a_entregar_termino"))
	ckb_pedido_nao_recebido_pelo_cliente = Trim(Request.Form("ckb_pedido_nao_recebido_pelo_cliente"))
	ckb_pedido_recebido_pelo_cliente = Trim(Request.Form("ckb_pedido_recebido_pelo_cliente"))
	ckb_st_entrega_entregue = Trim(Request.Form("ckb_st_entrega_entregue"))
	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))
	ckb_st_entrega_cancelado = Trim(Request.Form("ckb_st_entrega_cancelado"))
	c_dt_cancelado_inicio = Trim(Request.Form("c_dt_cancelado_inicio"))
	c_dt_cancelado_termino = Trim(Request.Form("c_dt_cancelado_termino"))
	ckb_st_pagto_pago = Trim(Request.Form("ckb_st_pagto_pago"))
	ckb_st_pagto_nao_pago = Trim(Request.Form("ckb_st_pagto_nao_pago"))
	ckb_st_pagto_pago_parcial = Trim(Request.Form("ckb_st_pagto_pago_parcial"))
	ckb_periodo_cadastro = Trim(Request.Form("ckb_periodo_cadastro"))
	c_dt_cadastro_inicio = Trim(Request.Form("c_dt_cadastro_inicio"))
	c_dt_cadastro_termino = Trim(Request.Form("c_dt_cadastro_termino"))
	ckb_entrega_marcada_para = Trim(Request.Form("ckb_entrega_marcada_para"))
	c_dt_entrega_inicio = Trim(Request.Form("c_dt_entrega_inicio"))
	c_dt_entrega_termino = Trim(Request.Form("c_dt_entrega_termino"))
	ckb_produto = Trim(Request.Form("ckb_produto"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_loja = Trim(Request.Form("c_loja"))
	c_cliente_cnpj_cpf=retorna_so_digitos(trim(request("c_cliente_cnpj_cpf")))
    c_cliente_uf=trim(request("c_cliente_uf"))
	c_transportadora = filtra_nome_identificador(UCase(Trim(Request.Form("c_transportadora"))))
	ckb_visanet = Trim(Request.Form("ckb_visanet"))
	ckb_perc_RT = Trim(Request.Form("ckb_perc_RT"))
	c_perc_RT = Trim(Request.Form("c_perc_RT"))
	ckb_analise_credito_st_inicial = Trim(Request.Form("ckb_analise_credito_st_inicial"))
	ckb_analise_credito_pendente_vendas = Trim(Request.Form("ckb_analise_credito_pendente_vendas"))
	ckb_analise_credito_pendente_endereco = Trim(Request.Form("ckb_analise_credito_pendente_endereco"))
	ckb_analise_credito_pendente = Trim(Request.Form("ckb_analise_credito_pendente"))
	ckb_analise_credito_pendente_cartao = Trim(Request.Form("ckb_analise_credito_pendente_cartao"))
	ckb_analise_credito_pendente_pagto_antecipado_boleto = Trim(Request.Form("ckb_analise_credito_pendente_pagto_antecipado_boleto"))
	ckb_analise_credito_ok = Trim(Request.Form("ckb_analise_credito_ok"))
	ckb_analise_credito_ok_aguardando_deposito = Trim(Request.Form("ckb_analise_credito_ok_aguardando_deposito"))
	ckb_analise_credito_ok_deposito_aguardando_desbloqueio = Trim(Request.Form("ckb_analise_credito_ok_deposito_aguardando_desbloqueio"))
	ckb_analise_credito_ok_aguardando_pagto_boleto_av = Trim(Request.Form("ckb_analise_credito_ok_aguardando_pagto_boleto_av"))
	ckb_entrega_imediata_sim = Trim(Request.Form("ckb_entrega_imediata_sim"))
	ckb_entrega_imediata_nao = Trim(Request.Form("ckb_entrega_imediata_nao"))
	c_dt_previsao_entrega_inicio = Trim(Request.Form("c_dt_previsao_entrega_inicio"))
	c_dt_previsao_entrega_termino = Trim(Request.Form("c_dt_previsao_entrega_termino"))
	op_forma_pagto = Trim(Request.Form("op_forma_pagto"))
	c_forma_pagto_qtde_parc = retorna_so_digitos(Trim(Request.Form("c_forma_pagto_qtde_parc")))
	c_vendedor = Trim(Request.Form("c_vendedor"))
	c_indicador = Trim(Request.Form("c_indicador"))
	ckb_indicador_preenchido = Trim(Request.Form("ckb_indicador_preenchido"))
	ckb_indicador_nao_preenchido = Trim(Request.Form("ckb_indicador_nao_preenchido"))
    c_pedido_origem = Trim(Request.Form("c_pedido_origem"))
    c_grupo_pedido_origem = Trim(Request.Form("c_grupo_pedido_origem"))
    c_empresa = Trim(Request.Form("c_empresa"))
	c_FormFieldValues = Trim(Request.Form("c_FormFieldValues"))
    c_grupo = Trim(Request.Form("c_grupo"))
    c_cancelados_ordena = Trim(Request.Form("c_cancelados_ordena"))
	
	call set_default_valor_texto_bd(usuario, "LOJA/RelPedidosMCrit|FormFields", c_FormFieldValues)
	
'	TRATA-SE DE UM VENDEDOR NORMAL?
	if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
		c_vendedor = usuario
		end if
	
	if alerta = "" then
		if c_fabricante <> "" then
			s = "SELECT fabricante FROM t_FABRICANTE WHERE (fabricante='" & c_fabricante & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "FABRICANTE " & c_fabricante & " NÃO ESTÁ CADASTRADO."
				end if
			end if
		end if
	
	if alerta = "" then
		if c_produto <> "" then
			if (Not IsEAN(c_produto)) And (c_fabricante="") then
				alerta=texto_add_br(alerta)
				alerta=alerta & "NÃO FOI ESPECIFICADO O FABRICANTE DO PRODUTO A SER CONSULTADO."
			else
				s = "SELECT * FROM t_PRODUTO WHERE"
				if IsEAN(c_produto) then
					s = s & " (ean='" & c_produto & "')"
				else
					s = s & " (fabricante='" & c_fabricante & "') AND (produto='" & c_produto & "')"
					end if
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if Not rs.Eof then
					flag_ok = True
					if IsEAN(c_produto) And (c_fabricante<>"") then
						if (c_fabricante<>Trim("" & rs("fabricante"))) then
							flag_ok = False
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto a ser consultado " & c_produto & " NÃO pertence ao fabricante " & c_fabricante & "."
							end if
						end if
					if flag_ok then
					'	CARREGA CÓDIGO INTERNO DO PRODUTO
						c_fabricante = Trim("" & rs("fabricante"))
						c_produto = Trim("" & rs("produto"))
						end if
					end if
				end if
			end if
		end if
	
	if alerta = "" then
		if c_loja = "" then
			alerta = "Especifique o número da loja."
		else
			s = "SELECT loja FROM t_LOJA WHERE (loja='" & c_loja & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta = "Loja " & c_loja & " não está cadastrada."
				end if
			end if
		end if
	
	if alerta = "" then
		if c_cliente_cnpj_cpf <> "" then
			if Not cnpj_cpf_ok(c_cliente_cnpj_cpf) then
				alerta=texto_add_br(alerta)
				alerta = alerta & "CNPJ/CPF do cliente é inválido."
				end if
			end if
		end if
	
	if alerta = "" then
		if c_transportadora <> "" then
			if Trim(x_transportadora(c_transportadora)) = "" then
				alerta=texto_add_br(alerta)
				alerta = alerta & "Transportadora '" & c_transportadora & "' NÃO está cadastrada."
				end if
			end if
		end if
	
	
'	Período de consulta está restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	if operacao_permitida(OP_LJA_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)

		if alerta = "" then
			if ckb_st_entrega_separar_com_marc <> "" then
				if (c_dt_coleta_a_separar_inicio <> "") And (c_dt_coleta_a_separar_termino <> "") then
					if StrToDate(c_dt_coleta_a_separar_termino) < StrToDate(c_dt_coleta_a_separar_inicio) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "A data de término (" & c_dt_coleta_a_separar_termino & ") do período de coleta é anterior à data de início (" & c_dt_coleta_a_separar_inicio & ")!"
						end if
					end if
				end if
			end if

		if alerta = "" then
			if ckb_st_entrega_a_entregar_com_marc <> "" then
				if (c_dt_coleta_st_a_entregar_inicio <> "") And (c_dt_coleta_st_a_entregar_termino <> "") then
					if StrToDate(c_dt_coleta_st_a_entregar_termino) < StrToDate(c_dt_coleta_st_a_entregar_inicio) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "A data de término (" & c_dt_coleta_st_a_entregar_termino & ") do período de coleta é anterior à data de início (" & c_dt_coleta_st_a_entregar_inicio & ")!"
						end if
					end if
				end if
			end if

	'	COLOCADOS ENTRE
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_cadastro_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_cadastro_termino
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			if c_dt_cadastro_inicio = "" then c_dt_cadastro_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
			end if
		
	'	ENTREGUE ENTRE
		if ckb_st_entrega_entregue <> "" then
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
			end if

	'	CANCELADO ENTRE
		if ckb_st_entrega_cancelado <> "" then
			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_cancelado_inicio
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_cancelado_termino
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				if c_dt_cancelado_inicio = "" then c_dt_cancelado_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
				end if
			end if
		
	'	DATA DE COLETA (RÓTULO ANTIGO: ENTREGA MARCADA ENTRE)
		if ckb_entrega_marcada_para <> "" then
			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_entrega_inicio
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_entrega_termino
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				if c_dt_entrega_inicio = "" then c_dt_entrega_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
				end if
			end if
		
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if


'   MOSTRA COLUNAS DE MOTIVO CANCELAMENTO E VALOR ORIGINAL DO PEDIDO CANCELADO?
    blnMostraMotivoCancelado = False
    if ckb_st_entrega_cancelado <> "" then blnMostraMotivoCancelado = True


' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim blnPorFornecedor
dim s, s_aux, s_periodo_aux, s_cor, s_sql, cab_table, cab, n_reg, n_reg_total, s_colspan
dim s_where, s_where_aux, s_from, cont
dim vl_total_faturamento, vl_sub_total_faturamento, vl_total_pago, vl_sub_total_pago
dim vl_total_faturamento_NF, vl_sub_total_faturamento_NF
dim vl_a_pagar, vl_sub_total_a_pagar, vl_total_a_pagar
dim vl_total_fornecedor, vl_sub_total_fornecedor
dim vl_total_fornecedor_NF, vl_sub_total_fornecedor_NF
dim vl_total_pedido_original, vl_sub_total_pedido_original, vl_pedido_original, s_class
dim x, loja_a, qtde_lojas
dim w_cliente, w_st_entrega, w_valor
dim blnRelAnalitico
dim intNumLinha
dim s_grupo_origem

'	RELATÓRIO SINTÉTICO OU ANALÍTICO?
	blnRelAnalitico=False
	if operacao_permitida(OP_LJA_REL_MULTICRITERIO_PEDIDOS_ANALITICO, s_lista_operacoes_permitidas) then blnRelAnalitico=True

'	MONTA CLÁUSULA WHERE
	s_where = ""

'	CRITÉRIO: STATUS DE ENTREGA
	s = ""
	s_aux = ckb_st_entrega_esperar
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.st_entrega = '" & s_aux & "')"
		end if

	s_aux = ckb_st_entrega_split
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.st_entrega = '" & s_aux & "')"
		end if

	s_aux = ckb_st_entrega_separar_sem_marc
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO.st_entrega = '" & s_aux & "')AND(t_PEDIDO.a_entregar_status=0))"
		end if

	s_aux = ckb_st_entrega_separar_com_marc
	if s_aux <> "" then
		s_where_aux = ""
		if c_dt_coleta_a_separar_inicio <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (t_PEDIDO.a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_coleta_a_separar_inicio)) & ")"
			end if
		if c_dt_coleta_a_separar_termino <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (t_PEDIDO.a_entregar_data_marcada < " & bd_formata_data(StrToDate(c_dt_coleta_a_separar_termino)+1) & ")"
			end if
		if s_where_aux <> "" then s_where_aux = " AND" & s_where_aux
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO.st_entrega = '" & s_aux & "')AND(t_PEDIDO.a_entregar_status<>0)" & s_where_aux & ")"
		end if

	s_aux = ckb_st_entrega_a_entregar_sem_marc
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO.st_entrega = '" & s_aux & "')AND(t_PEDIDO.a_entregar_status=0))"
		end if

	s_aux = ckb_st_entrega_a_entregar_com_marc
	if s_aux <> "" then
		s_where_aux = ""
		if c_dt_coleta_st_a_entregar_inicio <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (t_PEDIDO.a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_coleta_st_a_entregar_inicio)) & ")"
			end if
		if c_dt_coleta_st_a_entregar_termino <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (t_PEDIDO.a_entregar_data_marcada < " & bd_formata_data(StrToDate(c_dt_coleta_st_a_entregar_termino)+1) & ")"
			end if
		if s_where_aux <> "" then s_where_aux = " AND" & s_where_aux
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO.st_entrega = '" & s_aux & "')AND(t_PEDIDO.a_entregar_status<>0)" & s_where_aux & ")"
		end if

'	ENTREGUE ENTRE
	if ckb_st_entrega_entregue <> "" then
		s_aux = ""
		if c_dt_entregue_inicio <> "" then 
			if s_aux <> "" then s_aux = s_aux & " AND"
			s_aux = s_aux & " (t_PEDIDO.entregue_data >= " & bd_formata_data(StrToDate(c_dt_entregue_inicio)) & ")"
			end if
		if c_dt_entregue_termino <> "" then 
			if s_aux <> "" then s_aux = s_aux & " AND"
			s_aux = s_aux & " (t_PEDIDO.entregue_data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
			end if
		
		if s_aux <> "" then s_aux = s_aux & " AND"
		s_aux = s_aux & " (t_PEDIDO.st_entrega = '" & ckb_st_entrega_entregue & "')"
		
		if s_aux <> "" then s_aux = " (" & s_aux & ")"
		if s <> "" then s = s & " OR"
		s = s & s_aux
		end if

'	CANCELADO ENTRE
	if ckb_st_entrega_cancelado <> "" then
		s_aux = ""
		if c_dt_cancelado_inicio <> "" then 
			if s_aux <> "" then s_aux = s_aux & " AND"
			s_aux = s_aux & " (t_PEDIDO.cancelado_data >= " & bd_formata_data(StrToDate(c_dt_cancelado_inicio)) & ")"
			end if
		if c_dt_cancelado_termino <> "" then 
			if s_aux <> "" then s_aux = s_aux & " AND"
			s_aux = s_aux & " (t_PEDIDO.cancelado_data < " & bd_formata_data(StrToDate(c_dt_cancelado_termino)+1) & ")"
			end if
		
		if s_aux <> "" then s_aux = s_aux & " AND"
		s_aux = s_aux & " (t_PEDIDO.st_entrega = '" & ckb_st_entrega_cancelado & "')"
		
		if s_aux <> "" then s_aux = " (" & s_aux & ")"
		if s <> "" then s = s & " OR"
		s = s & s_aux
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	EXCETO CANCELADOS
	if ckb_st_entrega_exceto_cancelados <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')"
		end if

'	CRITÉRIO: PEDIDOS RECEBIDOS PELO CLIENTE
	s = ""
    s_aux = ckb_pedido_nao_recebido_pelo_cliente
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.PedidoRecebidoStatus = 0)"
		end if

	s_aux = ckb_pedido_recebido_pelo_cliente
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.PedidoRecebidoStatus = 1)"
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRITÉRIO: STATUS DE PAGAMENTO
	s = ""
	s_aux = ckb_st_pagto_pago
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.st_pagto = '" & s_aux & "')"
		end if

	s_aux = ckb_st_pagto_nao_pago
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.st_pagto = '" & s_aux & "')"
		end if
	
	s_aux = ckb_st_pagto_pago_parcial
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.st_pagto = '" & s_aux & "')"
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRITÉRIO: ANÁLISE DE CRÉDITO
	s = ""

	s_aux = ckb_analise_credito_st_inicial
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_pendente_vendas
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_pendente_endereco
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_pendente
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if
	
	s_aux = ckb_analise_credito_pendente_pagto_antecipado_boleto
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_ok
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_ok_aguardando_deposito
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_ok_deposito_aguardando_desbloqueio
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_ok_aguardando_pagto_boleto_av
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

'	O STATUS "PENDENTE CARTÃO DE CRÉDITO" NÃO EXISTE NO BD, É UMA SITUAÇÃO DEFINIDA
'	PELA COMBINAÇÃO DO STATUS COD_AN_CREDITO_ST_INICIAL + FORMA DE PAGTO USANDO SOMENTE PAGAMENTO POR CARTÃO
	s_aux = ckb_analise_credito_pendente_cartao
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO__BASE.analise_credito = " & COD_AN_CREDITO_ST_INICIAL & ") AND (t_PEDIDO__BASE.st_forma_pagto_somente_cartao = 1))"
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRITÉRIO: ENTREGA IMEDIATA
	s = ""
	if ckb_entrega_imediata_sim <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.st_etg_imediata = " & COD_ETG_IMEDIATA_SIM & ")"
		end if
	
	if ckb_entrega_imediata_nao <> "" then
		s_periodo_aux = ""
		if c_dt_previsao_entrega_inicio <> "" then
			s_periodo_aux = " (t_PEDIDO.PrevisaoEntregaData >= " & bd_formata_data(StrToDate(c_dt_previsao_entrega_inicio)) & ")"
			end if
		if c_dt_previsao_entrega_termino <> "" then
			if s_periodo_aux <> "" then s_periodo_aux = s_periodo_aux & " AND"
			s_periodo_aux = s_periodo_aux & " (t_PEDIDO.PrevisaoEntregaData < " & bd_formata_data(StrToDate(c_dt_previsao_entrega_termino)+1) & ")"
			end if
		if s_periodo_aux <> "" then s_periodo_aux = " AND" & s_periodo_aux
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO.st_etg_imediata = " & COD_ETG_IMEDIATA_NAO & ")" & s_periodo_aux & ")"
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CRITÉRIO: FORMA DE PAGAMENTO (NOVA VERSÃO)
	s = ""
	if op_forma_pagto <> "" then
		s = " (t_PEDIDO__BASE.av_forma_pagto = " & op_forma_pagto & ")" & _
			" OR (t_PEDIDO__BASE.pu_forma_pagto = " & op_forma_pagto & ")" & _
			" OR (t_PEDIDO__BASE.pce_forma_pagto_entrada = " & op_forma_pagto & ")" & _
			" OR (t_PEDIDO__BASE.pce_forma_pagto_prestacao = " & op_forma_pagto & ")" & _
			" OR (t_PEDIDO__BASE.pse_forma_pagto_prim_prest = " & op_forma_pagto & ")" & _
			" OR (t_PEDIDO__BASE.pse_forma_pagto_demais_prest = " & op_forma_pagto & ")"
		if op_forma_pagto = ID_FORMA_PAGTO_CARTAO then
			s = s & " OR (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_CARTAO & ")"
		elseif op_forma_pagto = ID_FORMA_PAGTO_CARTAO_MAQUINETA then
			s = s & " OR (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA & ")"
			end if
		end if
	
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRITÉRIO: QUANTIDADE DE PARCELAS
	s = ""
	if c_forma_pagto_qtde_parc <> "" then
		s = " (t_PEDIDO__BASE.qtde_parcelas = " & c_forma_pagto_qtde_parc & ")"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CRITÉRIO: PERÍODO DE CADASTRAMENTO DO PEDIDO
	s = ""
	if c_dt_cadastro_inicio <> "" then
		if s <> "" then s = s & " AND"
		s = s & " (t_PEDIDO.data >= " & bd_formata_data(StrToDate(c_dt_cadastro_inicio)) & ")"
		end if
		
	if c_dt_cadastro_termino <> "" then
		if s <> "" then s = s & " AND"
		s = s & " (t_PEDIDO.data < " & bd_formata_data(StrToDate(c_dt_cadastro_termino)+1) & ")"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CRITÉRIO: DATA DE COLETA (RÓTULO ANTIGO: ENTREGA MARCADA PARA)
	if ckb_entrega_marcada_para <> "" then
		s = ""
		if c_dt_entrega_inicio <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_entrega_inicio)) & ")"
			end if
		
		if c_dt_entrega_termino <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.a_entregar_data_marcada < " & bd_formata_data(StrToDate(c_dt_entrega_termino)+1) & ")"
			end if
		
		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
		end if
	
'	CRITÉRIO: PRODUTO
	blnPorFornecedor = False
	if ckb_produto <> "" then
		s = ""
		if c_fabricante <> "" then
			blnPorFornecedor = True
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO_ITEM.fabricante = '" & c_fabricante & "')"
			end if
		
		if c_produto <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO_ITEM.produto = '" & c_produto & "')"
			end if
		
		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
		end if

    s = ""
	if c_grupo <> "" then
	    v_grupos = split(c_grupo, ", ")
	    for cont = Lbound(v_grupos) to Ubound(v_grupos)
	        if s <> "" then s = s & " OR"
		    s = s & _
			    " (t_PRODUTO.grupo = '" & v_grupos(cont) & "')"
	    next
	    if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
		end if
    end if

' CRITÉRIO: ORIGEM DO PEDIDO (GRUPO)
    s = ""
    if c_grupo_pedido_origem <> "" then
        s_grupo_origem = "SELECT codigo FROM t_CODIGO_DESCRICAO WHERE (codigo_pai = '" & c_grupo_pedido_origem & "') AND grupo='PedidoECommerce_Origem'"
        if rs.State <> 0 then rs.Close
	    rs.open s_grupo_origem, cn
		if rs.Eof then
            alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_grupo_pedido_origem & " NÃO EXISTE."
        else
            do while Not rs.Eof
                if s <> "" then s = s & ", "
                s = s & "'" & rs("codigo") & "'"      
                rs.MoveNext
            loop
            s = " t_PEDIDO.marketplace_codigo_origem IN (" & s & ")"
        end if
        if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
		end if
    end if

' CRITÉRIO: ORIGEM DO PEDIDO
    s = ""
    if c_pedido_origem <> "" then
        s = s & " t_PEDIDO.marketplace_codigo_origem = '" & c_pedido_origem & "'"

        if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
		end if
    end if

' CRITÉRIO: EMPRESA
    s = ""
    if c_empresa <> "" then
        s = s & " t_PEDIDO.id_nfe_emitente = '" & c_empresa & "'"

        if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
		end if
    end if

'	CRITÉRIO: LOJA
	s = " (t_PEDIDO.numero_loja = " & c_loja & ")"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	CRITÉRIO: TRANSPORTADORA
	if c_transportadora <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.transportadora_id = '" & c_transportadora & "')"
		end if
	
'	CRITÉRIO: CLIENTE
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		if c_cliente_cnpj_cpf <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_PEDIDO.endereco_cnpj_cpf = '" & retorna_so_digitos(c_cliente_cnpj_cpf) & "')"
			end if
		if c_cliente_uf <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_PEDIDO.endereco_uf = '" & c_cliente_uf & "')"
		end if
	else
		if c_cliente_cnpj_cpf <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE.cnpj_cpf = '" & retorna_so_digitos(c_cliente_cnpj_cpf) & "')"
			end if
		if c_cliente_uf <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE.uf = '" & c_cliente_uf & "')"
		end if
	end if

'	CRITÉRIO: CARTÃO DE CRÉDITO (ANTIGAMENTE PELA VISANET E AGORA PELA CIELO)
	if ckb_visanet <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & _
					" (" & _
						"(" & _
							"(t_PEDIDO_PAGTO_VISANET__SUCESSO.operacao = '" & OP_VISANET_PAGAMENTO & "')" & _
							" AND (t_PEDIDO_PAGTO_VISANET__SUCESSO.concluido_status<>0)" & _
							" AND (t_PEDIDO_PAGTO_VISANET__SUCESSO.sucesso_status<>0)" & _
							" AND (t_PEDIDO_PAGTO_VISANET__SUCESSO.cancelado_status=0)" & _
						")" & _
						" OR " & _
						"(" & _
							"(t_PEDIDO_PAGTO_CIELO__SUCESSO.operacao = '" & OP_CIELO_OPERACAO__PAGAMENTO & "')" & _
							" AND (t_PEDIDO_PAGTO_CIELO__SUCESSO.sucesso_final_status<>0)" & _
							" AND (t_PEDIDO_PAGTO_CIELO__SUCESSO.cancelado_status=0)" & _
						")" & _
					")"
		end if
	
'	CRITÉRIO: COMISSÃO (RT) ABAIXO DO PERCENTUAL INFORMADO
	if ckb_perc_RT <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & _
					" (t_PEDIDO__BASE.perc_RT < " & bd_formata_numero(converte_numero(c_perc_RT)) & ")"
		end if

'	CRITÉRIO: VENDEDOR
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.vendedor = '" & c_vendedor & "')"
		end if

'	CRITÉRIO: INDICADOR (LEMBRE-SE: O ORÇAMENTISTA DE UM ORÇAMENTO É USADO AUTOMATICAMENTE COMO O INDICADOR DO PEDIDO QUANDO O ORÇAMENTO VIRA PEDIDO)
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if

'	CRITÉRIO: CAMPO INDICADOR PREENCHIDO
	if ckb_indicador_preenchido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (RTrim(Coalesce(t_PEDIDO.indicador,'')) <> '')"
		end if
		
'	CRITÉRIO: CAMPO INDICADOR NÃO PREENCHIDO
	if ckb_indicador_nao_preenchido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (RTrim(Coalesce(t_PEDIDO.indicador,'')) = '')"
		end if
	
'	CLÁUSULA WHERE
	if s_where <> "" then s_where = " WHERE" & s_where
	
	
'	MONTA CLÁUSULA FROM
	s_from = " FROM t_PEDIDO" & _
			 " INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)"
	
	if ckb_produto <> "" then
		s_from = s_from & " INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)"
		end if

    if c_grupo <> "" then
		if ckb_produto = "" then s_from = s_from & " INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)"
        s_from = s_from & " INNER JOIN t_PRODUTO ON (t_PEDIDO_ITEM.produto = t_PRODUTO.produto)"
        end if
	
	if c_cliente_cnpj_cpf <> "" Or c_cliente_uf <> "" then
		s_from = s_from & " INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)"
	else
		s_from = s_from & " LEFT JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)"
		end if
	
'	PAGAMENTO POR CARTÃO (ANTIGAMENTE PELA VISANET E AGORA PELA CIELO)
	if ckb_visanet <> "" then
		s_from = s_from & " LEFT JOIN (" & _
								"SELECT " & _
									"*" & _
								" FROM t_PEDIDO_PAGTO_VISANET" & _
								" WHERE" & _
									" (t_PEDIDO_PAGTO_VISANET.operacao = '" & OP_VISANET_PAGAMENTO & "')" & _
									" AND (t_PEDIDO_PAGTO_VISANET.concluido_status<>0)" & _
									" AND (t_PEDIDO_PAGTO_VISANET.sucesso_status<>0)" & _
									" AND (t_PEDIDO_PAGTO_VISANET.cancelado_status=0)" & _
								") AS t_PEDIDO_PAGTO_VISANET__SUCESSO ON (t_PEDIDO.pedido=t_PEDIDO_PAGTO_VISANET__SUCESSO.pedido)"
		
		s_from = s_from & " LEFT JOIN (" & _
								"SELECT " & _
									"*" & _
								" FROM t_PEDIDO_PAGTO_CIELO" & _
								" WHERE" & _
									" (t_PEDIDO_PAGTO_CIELO.operacao = '" & OP_CIELO_OPERACAO__PAGAMENTO & "')" & _
									" AND (t_PEDIDO_PAGTO_CIELO.sucesso_final_status<>0)" & _
									" AND (t_PEDIDO_PAGTO_CIELO.cancelado_status=0)" & _
								") AS t_PEDIDO_PAGTO_CIELO__SUCESSO ON (t_PEDIDO.pedido=t_PEDIDO_PAGTO_CIELO__SUCESSO.pedido)"
		end if

'	CRIA UMA "DERIVED TABLE" PARA OBTER O TOTAL EM DEVOLUÇÕES DO PEDIDO
	s_from = s_from & _
			" LEFT JOIN (" & _
				"SELECT pedido," & _
				" Sum(qtde) AS qtde_produtos_devolvidos," & _
				" Sum(qtde*preco_venda) AS vl_devolucao_pedido," & _
				" Sum(qtde*preco_NF) AS vl_devolucao_pedido_NF" & _
				" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
				" GROUP BY pedido" & _
				") AS t_PEDIDO_ITEM_DEVOLVIDO__AUX" & _
			" ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO__AUX.pedido)"

'	CRIA UMA "DERIVED TABLE" PARA OBTER O VALOR TOTAL DO PEDIDO
	s_from = s_from & _
			" LEFT JOIN (" & _
				"SELECT t_PEDIDO_ITEM.pedido AS pedido," & _
				" Sum(qtde*preco_venda) AS vl_total_pedido," & _
				" Sum(qtde*preco_NF) AS vl_total_pedido_NF" & _
				" FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO" & _
				" ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
				" WHERE (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')" & _
				" GROUP BY t_PEDIDO_ITEM.pedido" & _
				") AS t_PEDIDO__VL_TOTAL" & _
			" ON (t_PEDIDO.pedido=t_PEDIDO__VL_TOTAL.pedido)"

'	CRIA UMA "DERIVED TABLE" PARA OBTER O TOTAL EM PAGAMENTOS DO PEDIDO
	s_from = s_from & _
			" LEFT JOIN (" & _
				"SELECT pedido," & _
				" Sum(valor) AS vl_pago_pedido" & _
				" FROM t_PEDIDO_PAGAMENTO" & _
				" GROUP BY pedido" & _
				") AS t_PEDIDO__VL_PAGO" & _
			" ON (t_PEDIDO.pedido=t_PEDIDO__VL_PAGO.pedido)"

'	CRIA UMA "DERIVED TABLE" PARA OBTER O VALOR TOTAL RELATIVO AO FORNECEDOR
	if blnPorFornecedor then
		s_from = s_from & _
				" LEFT JOIN (" & _
					"SELECT t_PEDIDO_ITEM.pedido AS pedido," & _
					" Sum(qtde*preco_venda) AS vl_total_fornecedor," & _
					" Sum(qtde*preco_NF) AS vl_total_fornecedor_NF" & _
					" FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO" & _
					" ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
					" WHERE (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')" & _
					" AND (fabricante = '" & c_fabricante & "')" & _
					" GROUP BY t_PEDIDO_ITEM.pedido" & _
					") AS t_PEDIDO__VL_FORNECEDOR" & _
				" ON (t_PEDIDO.pedido=t_PEDIDO__VL_FORNECEDOR.pedido)"
		end if

'	OBS: SINTAXE DA FUNÇÃO ISNULL():
'		 ISNULL ( check_expression , replacement_value )
'		 SE "check_expression" FOR NULL, RETORNA "replacement_value"
	s_sql = "SELECT DISTINCT t_PEDIDO.loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO.data, t_PEDIDO.pedido, t_PEDIDO.pedido_bs_x_ac," & _
			" t_PEDIDO.st_entrega,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO__BASE.st_pagto," & _
            " t_PEDIDO__BASE.vendedor," & _
            " t_PEDIDO__BASE.indicador," & _
            " t_PEDIDO.cancelado_codigo_motivo," & _
            " t_PEDIDO.cancelado_codigo_sub_motivo," & _
			" ISNULL(t_PEDIDO__VL_TOTAL.vl_total_pedido,0) AS vl_total_pedido," & _
			" ISNULL(t_PEDIDO__VL_TOTAL.vl_total_pedido_NF,0) AS vl_total_pedido_NF," & _
			" ISNULL(t_PEDIDO__VL_PAGO.vl_pago_pedido,0) AS vl_pago_pedido," & _
			" ISNULL(t_PEDIDO_ITEM_DEVOLVIDO__AUX.vl_devolucao_pedido,0) AS vl_devolucao_pedido," & _
			" ISNULL(t_PEDIDO_ITEM_DEVOLVIDO__AUX.vl_devolucao_pedido_NF,0) AS vl_devolucao_pedido_NF," & _
			" ISNULL(t_PEDIDO_ITEM_DEVOLVIDO__AUX.qtde_produtos_devolvidos,0) AS qtde_produtos_devolvidos"

    if blnMostraMotivoCancelado then
        s_sql = s_sql & _            
            ", Coalesce((SELECT Sum(qtde * preco_venda) FROM t_PEDIDO_ITEM WHERE (pedido = t_PEDIDO.pedido)), 0) AS vl_total_original"
    end if
	
	if blnPorFornecedor then
		s_sql = s_sql & _
				", ISNULL(t_PEDIDO__VL_FORNECEDOR.vl_total_fornecedor,0) AS vl_total_fornecedor" & _
				", ISNULL(t_PEDIDO__VL_FORNECEDOR.vl_total_fornecedor_NF,0) AS vl_total_fornecedor_NF"
		end if
	
	s_sql = s_sql & _
			s_from & _
			s_where

	if ckb_st_entrega_cancelado <> "" then
        if c_cancelados_ordena = "VENDEDOR" then
            s_sql = s_sql & " ORDER BY numero_loja, vendedor, indicador, t_PEDIDO.data, t_PEDIDO.pedido"
        else
	        s_sql = s_sql & " ORDER BY numero_loja, t_PEDIDO.data, t_PEDIDO.pedido"
        end if
    else
	    s_sql = s_sql & " ORDER BY numero_loja, t_PEDIDO.data, t_PEDIDO.pedido"
    end if

  ' CABEÇALHO
	if blnPorFornecedor then
		if blnRelAnalitico then
			w_cliente = 201
		else
			w_cliente = 400
			end if
		w_st_entrega = 74
		w_valor = 70
	else
		if blnRelAnalitico then
			w_cliente = 250
		else
			w_cliente = 400
			end if
		w_st_entrega = 70
		w_valor = 80
		end if
		
	cab_table = "<table cellspacing=0>" & chr(13)
	cab = "	<tr style='background:azure'>" & chr(13) & _
		  "		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
		  "		<td class='MT' style='width:70px' align='left' valign='bottom'><p class='R'>Nº Pedido</p></td>" & chr(13)
          if loja=NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
		    cab = cab & "		<td class='MTBD' style='width:70px' align='left' valign='bottom'><p class='R'>Número Magento</p></td>" & chr(13)
          end if
		  cab = cab & "     <td class='MTBD' style='width:70px' align='center' valign='bottom' nowrap><p class='R'>Data</p></td>" & chr(13) & _
		  "		<td class='MTBD' style='width:" & Cstr(w_cliente) & "px' align='left' valign='bottom'><p class='R'>Cliente</p></td>" & chr(13)
	
	if blnPorFornecedor then
		if blnRelAnalitico then
			cab = cab & _ 
				  "		<td class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><p class='Rd' style='font-weight:bold;'>VL Fornec</p></td>" & chr(13) & _
				  "		<td class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><p class='Rd' style='font-weight:bold;'>VL Fornec (RA)</p></td>" & chr(13) & _
				  "		<td class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><p class='Rd' style='font-weight:bold;'>VL Pedido</p></td>" & chr(13) & _
				  "		<td class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><p class='Rd' style='font-weight:bold;'>VL Pedido (RA)</p></td>" & chr(13)
			end if
	else
		if blnRelAnalitico then
			cab = cab & _ 
				  "		<td class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><p class='Rd' style='font-weight:bold;'>Valor</p></td>" & chr(13) & _
				  "		<td class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><p class='Rd' style='font-weight:bold;'>Valor (RA)</p></td>" & chr(13)
			end if
		end if
	
	if blnRelAnalitico then
		cab = cab & _ 
			  "		<td class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><p class='Rd' style='font-weight:bold;'>VL Pago</p></td>" & chr(13) & _
			  "		<td class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><p class='Rd' style='font-weight:bold;'>VL A Pagar</p></td>" & chr(13)
		end if
		
	cab = cab & _ 
		  "		<td class='MTBD' style='width:" & Cstr(w_st_entrega) & "px' align='left' valign='bottom' nowrap><p class='R'>Status de Entrega</p></td>" & chr(13)

	if blnMostraMotivoCancelado then
        cab = cab & _
                "		<td class='MTBD' style='width:" & Cstr(w_valor) & "px' valign='bottom' NOWRAP><p class='R' style='font-weight:bold;'>Vendedor</p></td>" & chr(13) & _
                "		<td class='MTBD' style='width:" & Cstr(w_valor) & "px' valign='bottom' NOWRAP><p class='R' style='font-weight:bold;'>Indicador</p></td>" & chr(13)
        if blnRelAnalitico then
            cab = cab & _
                "		<td class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom' NOWRAP><p class='Rd' style='font-weight:bold;'>VL Original</p></td>" & chr(13)
        end if
        cab = cab & _
            "		<td class='MTBD' style='width:200px' valign='bottom' NOWRAP><p class='R' style='font-weight:bold;'>Motivo Cancelamento</p></td>" & chr(13)
    end if

    cab = cab & _
		  "	</tr>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_lojas = 0
	vl_total_faturamento = 0
	vl_total_faturamento_NF = 0
	vl_total_pago = 0
	vl_total_a_pagar = 0
	vl_total_fornecedor = 0
	vl_total_fornecedor_NF = 0
    vl_total_pedido_original = 0
	intNumLinha = 0
	
	loja_a = "XXXXX"

    s_class = ""
    if blnMostraMotivoCancelado then 
        s_class = "MB"
    else
        s_class = "MDB"
    end if
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE LOJA?
		if Trim("" & r("loja"))<>loja_a then
			loja_a = Trim("" & r("loja"))
			qtde_lojas = qtde_lojas + 1
		  ' FECHA TABELA DA LOJA ANTERIOR
			if n_reg > 0 then 
				if blnRelAnalitico then
					s_cor = ""
					if vl_sub_total_a_pagar < 0 then s_cor = " style='color:red;'"
					x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
							"		<td style='background:white;' align='left'>&nbsp;</td>" & chr(13)
                            if loja=NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
							   x = x & "		<td class='MEB' colspan='4' align='right' nowrap><p class='Cd'>TOTAL:</p></td>" & chr(13)
                            else
                               x = x & "		<td class='MEB' colspan='3' align='right' nowrap><p class='Cd'>TOTAL:</p></td>" & chr(13)
                            end if
				
					if blnPorFornecedor then
						x = x & _
							"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_fornecedor) & "</p></td>" & chr(13) & _
							"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_fornecedor_NF) & "</p></td>" & chr(13)
						end if
						
					x = x & _
							"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_faturamento) & "</p></td>" & chr(13) & _
							"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_faturamento_NF) & "</p></td>" & chr(13) & _
							"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_pago) & "</p></td>" & chr(13) & _
							"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'" & s_cor & ">" & formata_moeda(vl_sub_total_a_pagar) & "</p></td>" & chr(13)
							
					x = x & _
							"		<td class='" & s_class & "' align='left'><p class='C'>&nbsp;</p></td>" & chr(13)

                    if blnMostraMotivoCancelado then
                        x = x & _
                            "		<TD class='MB'><p class='C'>&nbsp;</p></td>" & chr(13) & _
                            "		<TD class='MB'><p class='C'>&nbsp;</p></td>" & chr(13) & _
							"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_pedido_original) & "</p></td>" & chr(13) & _
							"		<TD class='MDB'><p class='C'>&nbsp;</p></td>" & chr(13)
                        end if

                    x = x & _
						"	</tr>" & chr(13)
                        
					end if
                    
                    x = x & _
							"</table>" & chr(13)
					Response.Write x
					x="<BR>" & chr(13)
				end if

			n_reg = 0
			vl_sub_total_faturamento = 0
			vl_sub_total_faturamento_NF = 0
			vl_sub_total_pago = 0
			vl_sub_total_a_pagar = 0
			vl_sub_total_fornecedor = 0
			vl_sub_total_fornecedor_NF = 0
            vl_sub_total_pedido_original = 0

			if n_reg_total > 0 then x = x & "<BR>" & chr(13)
			s = Trim("" & r("loja"))
			s_aux = x_loja(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then
				if blnPorFornecedor then 
					if blnRelAnalitico then 
						if blnMostraMotivoCancelado then
                            s_colspan = "14"
                        else
						    s_colspan = "10" 
                        end if
					else
						if blnMostraMotivoCancelado then
                            s_colspan = "8"
                        else
						    s_colspan = "4" 
                            end if
						end if
				else 
					if blnRelAnalitico then
                        if loja=NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
						    if blnMostraMotivoCancelado then
                                s_colspan = "13"
                            else
						        s_colspan = "9" 
                                end if 
                        else
                            if blnMostraMotivoCancelado then
                                s_colspan = "12"
                            else
						        s_colspan = "8" 
                            end if
                        end if
					else
						if blnMostraMotivoCancelado then
                            s_colspan = "8"
                        else
						    s_colspan = "5" 
                            end if
						end if
					end if
				
				x = x & _
					"	<tr>" & chr(13) & _
					"		<td style='background:white;' align='left'>&nbsp;</td>" & chr(13) & _
					"		<td class='MDTE' colspan='" & s_colspan & "' align='left' valign='bottom' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
				end if
			x = x & cab
			end if

	 ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1
		intNumLinha = intNumLinha + 1

		x = x & "	<tr>" & chr(13)

	'> Nº DA LINHA
		x = x & "		<td align='right' valign='top' nowrap><p class='Rd' style='margin-right:2px;'>" & Cstr(intNumLinha) & ".</p></td>" & chr(13)
		


	'> Nº PEDIDO
		x = x & "		<td align='left' valign='top' class='MDBE'><p class='C'>&nbsp;<a href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & _
				")' title='clique para consultar o pedido'>" & Trim("" & r("pedido")) & "</a></p></td>" & chr(13)

    '> PEDIDO MAGENTO
        if loja=NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
		    x = x & "		<td align='left' valign='top' class='MDB'><p class='C'>&nbsp;" & Trim("" & r("pedido_bs_x_ac")) & "</p></td>" & chr(13)
        end if
				
	'> DATA DO PEDIDO
	    x = x & "       <td align='center' valign='top' class='MDB'><p class='Cn'>&nbsp;" & r("data") & "</p></td>" & chr(13)

	'> CLIENTE
		s = Trim("" & r("nome_iniciais_em_maiusculas"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td align='left' valign='top' style='width:" & Cstr(w_cliente) & "px' class='MDB'><p class='Cn'>" & s & "</p></td>" & chr(13)

	'> VALOR DO FORNECEDOR
		if blnPorFornecedor then
			if blnRelAnalitico then
				s = formata_moeda(r("vl_total_fornecedor"))
				x = x & "		<td align='right' valign='top' style='width:" & Cstr(w_valor) & "px' class='MDB'><p class='Cnd'>" & s & "</p></td>" & chr(13)
				end if
			end if
		
	'> VALOR DO FORNECEDOR COM RA
		if blnPorFornecedor then
			if blnRelAnalitico then
				s = formata_moeda(r("vl_total_fornecedor_NF"))
				x = x & "		<td align='right' valign='top' style='width:" & Cstr(w_valor) & "px' class='MDB'><p class='Cnd'>" & s & "</p></td>" & chr(13)
				end if
			end if
		
	'> VALOR DO PEDIDO
		if blnRelAnalitico then
			s = formata_moeda(r("vl_total_pedido"))
			x = x & "		<td align='right' valign='top' style='width:" & Cstr(w_valor) & "px' class='MDB'><p class='Cnd'>" & s & "</p></td>" & chr(13)
			end if
			
	'> VALOR DO PEDIDO COM RA
		if blnRelAnalitico then
			s = formata_moeda(r("vl_total_pedido_NF"))
			x = x & "		<td align='right' valign='top' style='width:" & Cstr(w_valor) & "px' class='MDB'><p class='Cnd'>" & s & "</p></td>" & chr(13)
			end if
		
	'> VALOR JÁ PAGO
		if blnRelAnalitico then
			s = formata_moeda(r("vl_pago_pedido"))
			x = x & "		<td align='right' valign='top' style='width:" & Cstr(w_valor) & "px' class='MDB'><p class='Cnd'>" & s & "</p></td>" & chr(13)
			end if
		
	'> VALOR A PAGAR
		vl_a_pagar = 0
		s_cor = ""
		vl_a_pagar = r("vl_total_pedido_NF")-r("vl_pago_pedido")-r("vl_devolucao_pedido_NF")
	'	VALORES NEGATIVOS REPRESENTAM O 'CRÉDITO' QUE O CLIENTE POSSUI EM CASO DE PEDIDOS CANCELADOS QUE HAVIAM SIDO PAGOS
		if (Trim("" & r("st_pagto")) = ST_PAGTO_PAGO) And (vl_a_pagar > 0)  then vl_a_pagar = 0
		s = formata_moeda(vl_a_pagar)
		if blnRelAnalitico then
			if vl_a_pagar < 0 then s_cor = " style='color:red;'"
			x = x & "		<td align='right' valign='top' style='width:" & Cstr(w_valor) & "px' class='MDB'><p class='Cnd'" & s_cor & ">" & s & "</p></td>" & chr(13)
			end if
			
	'> STATUS DE ENTREGA
		s = Trim("" & r("st_entrega"))
		if s <> "" then 
			s = x_status_entrega(s)
			if (Trim("" & r("st_entrega"))=ST_ENTREGA_ENTREGUE) And (converte_numero(r("qtde_produtos_devolvidos"))>0) then s = s & " (*)"
			end if
		if s = "" then s = "&nbsp;"
		x = x & "		<td align='left' valign='top' style='width:" & Cstr(w_st_entrega) & "px' class='MDB'><p class='Cn'>" & s & "</p></td>" & chr(13)

    '> VENDEDOR
        if blnMostraMotivoCancelado then
            s = Trim("" & r("vendedor"))
            x = x & "		<TD valign='top' style='width:" & Cstr(w_valor) & "px' class='MDB'><P class='Cn'>" & s & "</P></TD>" & chr(13)
        end if

    '> INDICADOR
        if blnMostraMotivoCancelado then
            s = Trim("" & r("indicador"))
            x = x & "		<TD valign='top' style='width:" & Cstr(w_valor) & "px' class='MDB'><P class='Cn'>" & s & "</P></TD>" & chr(13)
        end if

    '> VALOR ORIGINAL DO PEDIDO
        if blnMostraMotivoCancelado then
            if blnRelAnalitico then
                s = formata_moeda(r("vl_total_original"))
                x = x & "		<TD valign='top' align='right' style='width:" & Cstr(w_valor) & "px' class='MDB'><P class='Cnd'>" & s & "</P></TD>" & chr(13)
            end if
        end if

    '> MOTIVO CANCELAMENTO
        if blnMostraMotivoCancelado then
            if Trim("" & r("cancelado_codigo_motivo")) <> "" then
                s = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CANCELAMENTOPEDIDO_MOTIVO, Trim("" & r("cancelado_codigo_motivo")))
            end if
            if Trim("" & r("cancelado_codigo_sub_motivo")) <> "" then
                s = s & " (" & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CANCELAMENTOPEDIDO_MOTIVO_SUB, Trim("" & r("cancelado_codigo_sub_motivo"))) & ")"
            end if
		    x = x & "		<TD valign='top' style='width:200px' class='MDB'><P class='Cn'>" & s & "</P></TD>" & chr(13)            
        end if
	
	'> TOTALIZAÇÃO DE VALORES
		vl_sub_total_faturamento = vl_sub_total_faturamento + r("vl_total_pedido")
		vl_sub_total_faturamento_NF = vl_sub_total_faturamento_NF + r("vl_total_pedido_NF")
		vl_sub_total_pago = vl_sub_total_pago + r("vl_pago_pedido")
		vl_sub_total_a_pagar = vl_sub_total_a_pagar + vl_a_pagar
		if blnPorFornecedor then 
			vl_sub_total_fornecedor = vl_sub_total_fornecedor + r("vl_total_fornecedor")
			vl_sub_total_fornecedor_NF = vl_sub_total_fornecedor_NF + r("vl_total_fornecedor_NF")
			end if
        if blnMostraMotivoCancelado then
            vl_sub_total_pedido_original = vl_sub_total_pedido_original + r("vl_total_original")
            end if
		
		vl_total_faturamento = vl_total_faturamento + r("vl_total_pedido")
		vl_total_faturamento_NF = vl_total_faturamento_NF + r("vl_total_pedido_NF")
		vl_total_pago = vl_total_pago + r("vl_pago_pedido")
		vl_total_a_pagar = vl_total_a_pagar + vl_a_pagar
		if blnPorFornecedor then 
			vl_total_fornecedor = vl_total_fornecedor + r("vl_total_fornecedor")
			vl_total_fornecedor_NF = vl_total_fornecedor_NF + r("vl_total_fornecedor_NF")
			end if
        if blnMostraMotivoCancelado then
            vl_total_pedido_original = vl_total_pedido_original + r("vl_total_original")
            end if
			
		x = x & "	</tr>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
		
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if blnRelAnalitico then
		if n_reg <> 0 then 
			s_cor = ""
			if vl_sub_total_a_pagar < 0 then s_cor = " style='color:red;'"
			x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
					"		<td style='background:white;' align='left'>&nbsp;</td>" & chr(13)
					if loja=NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
							   x = x & "		<td class='MEB' colspan='4' align='right' nowrap><p class='Cd'>TOTAL:</p></td>" & chr(13)
                            else
                               x = x & "		<td class='MEB' colspan='3' align='right' nowrap><p class='Cd'>TOTAL:</p></td>" & chr(13)
                            end if
					
			if blnPorFornecedor then
				x = x & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_fornecedor) & "</p></td>" & chr(13) & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_fornecedor_NF) & "</p></td>" & chr(13)
				end if
				
			x = x & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_faturamento) & "</p></td>" & chr(13) & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_faturamento_NF) & "</p></td>" & chr(13) & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_pago) & "</p></td>" & chr(13) & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'" & s_cor & ">" & formata_moeda(vl_sub_total_a_pagar) & "</p></td>" & chr(13) & _
					"		<td align='left' class='" & s_class & "'><p class='C'>&nbsp;</p></td>" & chr(13)

			if blnMostraMotivoCancelado then
                x = x & _
                    "		<TD class='MB'><p class='C'>&nbsp;</p></td>" & chr(13) & _
                    "		<TD class='MB'><p class='C'>&nbsp;</p></td>" & chr(13) & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><p class='Cd'>" & formata_moeda(vl_sub_total_pedido_original) & "</p></td>" & chr(13) & _
					"		<td class='MDB'><p class='C'>&nbsp;</p></td>" & chr(13)
                end if

            x = x & _
					"	</tr>" & chr(13)

		'>	TOTAL GERAL
			if qtde_lojas > 1 then
				s_cor = ""
				if vl_total_a_pagar < 0 then s_cor = " style='color:red;'"
				if blnPorFornecedor then s_colspan = "10" else s_colspan = "8" 
				x = x & _
					"	<tr>" & chr(13) & _
					"		<td colspan='" & s_colspan & "' align='left' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<tr>" & chr(13) & _
					"		<td colspan='" & s_colspan & "' align='left' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<tr style='background:honeydew'>" & chr(13) & _
					"		<td style='background:white;' align='left'>&nbsp;</td>" & chr(13) & _
					"		<td class='MTBE' colspan='3' align='right' nowrap><p class='Cd'>TOTAL GERAL:</p></td>" & chr(13)
					
				if blnPorFornecedor then
					x = x & _
						"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><p class='Cd'>" & formata_moeda(vl_total_fornecedor) & "</p></td>" & chr(13) & _
						"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><p class='Cd'>" & formata_moeda(vl_total_fornecedor_NF) & "</p></td>" & chr(13)
					end if

                s_class = ""
                if blnMostraMotivoCancelado then 
                    s_class = "MTB"
                else 
                    s_class = "MTBD"
                end if
					
				x = x & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><p class='Cd'>" & formata_moeda(vl_total_faturamento) & "</p></td>" & chr(13) & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><p class='Cd'>" & formata_moeda(vl_total_faturamento_NF) & "</p></td>" & chr(13) & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><p class='Cd'>" & formata_moeda(vl_total_pago) & "</p></td>" & chr(13) & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><p class='Cd'" & s_cor & ">" & formata_moeda(vl_total_a_pagar) & "</p></td>" & chr(13) & _
					"		<td align='left' class='MTBD'><p class='C'>&nbsp;</p></td>" & chr(13)

					if blnMostraMotivoCancelado then
                    x = x & _
                    "		<td class='MTB'><p class='C'>&nbsp;</p></td>" & chr(13) & _
                    "		<td class='MTB'><p class='C'>&nbsp;</p></td>" & chr(13) & _
					"		<td align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><p class='Cd'>" & formata_moeda(vl_total_pedido_original) & "</p></td>" & chr(13) & _
					"		<td class='MTBD'><p class='C'>&nbsp;</p></td>" & chr(13)
                end if

                x = x & _
                    "	</tr>" & chr(13)
				end if
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		if blnPorFornecedor then 
			if blnRelAnalitico then
				if blnMostraMotivoCancelado then
                    s_colspan = "13"
                else
					s_colspan = "11" 
                end if 
			else
				if blnMostraMotivoCancelado then
                    s_colspan = "6"
                else
					s_colspan = "5" 
                    end if
				end if
		else 
			if blnRelAnalitico then
				if loja=NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
					if blnMostraMotivoCancelado then
                    s_colspan = "11"
                else
					s_colspan = "9" 
                    end if
				else
					if blnMostraMotivoCancelado then
                        s_colspan = "10"
                    else
					    s_colspan = "8" 
                        end if
					end if
			else
				if blnMostraMotivoCancelado then
                    s_colspan = "6"
                else
					s_colspan = "5" 
                    end if
				end if
			end if
			
		x = cab_table & cab
		x = x & "	<tr>" & chr(13) & _
				"		<td style='background:white;' align='left'>&nbsp;</td>" & chr(13) & _
				"		<td class='MDBE ALERTA' colspan='" & s_colspan & "' align='center'><p class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</p></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' FECHA TABELA DA ÚLTIMA LOJA
	x = x & "</table>" & chr(13)
	
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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir( id_pedido ){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=id_pedido;
	fREL.action = "pedido.asp"
	fREL.submit();
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
<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<input type="hidden" name="url_origem" id="url_origem" value="RelPedidosMCrit.asp" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="849" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório Multicritério de Pedidos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='849' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
	s = ""
	s_aux = Lcase(x_status_entrega(ckb_st_entrega_esperar))
	if s_aux<>"" then
	'	DEVIDO AO WORD WRAP: SÓ FAZ WORD WRAP QUANDO ENCONTRA CHR(32), OU SEJA, MANTÉM AGRUPADO TEXTO COM &nbsp;
		if s <> "" then s = s & ",&nbsp; "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_status_entrega(ckb_st_entrega_split))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_separar_sem_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		s_aux = s_aux & " (sem data de coleta)"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_separar_com_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		if c_dt_coleta_a_separar_inicio <> "" then s_aux_dti = c_dt_coleta_a_separar_inicio else s_aux_dti = "N.I."
		if c_dt_coleta_a_separar_termino <> "" then s_aux_dtf = c_dt_coleta_a_separar_termino else s_aux_dtf = "N.I."
		s_aux = s_aux & " (com data de coleta: " & s_aux_dti & " a " & s_aux_dtf & ")"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_status_entrega(ckb_st_entrega_a_entregar_sem_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		s_aux = s_aux & " (sem data de coleta)"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_a_entregar_com_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		if c_dt_coleta_st_a_entregar_inicio <> "" then s_aux_dti = c_dt_coleta_st_a_entregar_inicio else s_aux_dti = "N.I."
		if c_dt_coleta_st_a_entregar_termino <> "" then s_aux_dtf = c_dt_coleta_st_a_entregar_termino else s_aux_dtf = "N.I."
		s_aux = s_aux & " (com data de coleta: " & s_aux_dti & " a " & s_aux_dtf & ")"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_entregue))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		s = s & s_aux
		s_aux = c_dt_entregue_inicio
		if s_aux = "" then s_aux = "N.I."
		s_aux = " (" & s_aux & " a "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		s_aux = c_dt_entregue_termino
		if s_aux = "" then s_aux = "N.I."
		s_aux = s_aux & ")"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_cancelado))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		s = s & s_aux
		s_aux = c_dt_cancelado_inicio
		if s_aux = "" then s_aux = "N.I."
		s_aux = " (" & s_aux & " a "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		s_aux = c_dt_cancelado_termino
		if s_aux = "" then s_aux = "N.I."
		s_aux = s_aux & ")"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	if ckb_st_entrega_exceto_cancelados <> "" then
		s_aux = "exceto cancelados"
		if s <> "" then s = s & ",&nbsp; "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Status de Entrega:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = ""
    if ckb_pedido_nao_recebido_pelo_cliente <> "" then
		s_aux = "não recebidos"
		if s <> "" then s = s & ",&nbsp; "
		s = s & s_aux
		end if

	if ckb_pedido_recebido_pelo_cliente <> "" then
		s_aux = "recebidos"
		if s <> "" then s = s & ",&nbsp; "
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Pedidos Recebidos pelo Cliente:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = ""
	s_aux = Lcase(x_status_pagto(ckb_st_pagto_pago))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_status_pagto(ckb_st_pagto_nao_pago))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_pagto(ckb_st_pagto_pago_parcial))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Status de Pagamento:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = ""

	if ckb_analise_credito_st_inicial <> "" then
		s_aux = "status inicial"
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_vendas))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_endereco))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_cartao))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_pagto_antecipado_boleto))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok_aguardando_deposito))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok_deposito_aguardando_desbloqueio))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok_aguardando_pagto_boleto_av))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Análise de Crédito:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = ""
	s_aux = ""
	if CStr(ckb_entrega_imediata_sim) = CStr(COD_ETG_IMEDIATA_SIM) then s_aux = "sim"
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = ""
	if CStr(ckb_entrega_imediata_nao) = CStr(COD_ETG_IMEDIATA_NAO) then s_aux = "não"
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		s = s & " (previsão de entrega: "
		s_aux = c_dt_previsao_entrega_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s = s & " a "
		s_aux = c_dt_previsao_entrega_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s = s & ")"
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Entrega Imediata:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if
	
	'Indicador preenchido
	s = ""
	s_aux = ""
	if ckb_indicador_preenchido <> "" then s_aux = "Indicador preenchido"
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = ""
	if ckb_indicador_nao_preenchido <> "" then s_aux = "Indicador não preenchido"
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Indicador:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if (c_dt_cadastro_inicio <> "") Or (c_dt_cadastro_termino <> "") then
		s = ""
		s_aux = c_dt_cadastro_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " e "
		s_aux = c_dt_cadastro_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Pedidos colocados entre:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if ckb_entrega_marcada_para <> "" then
		s = ""
		s_aux = c_dt_entrega_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " a "
		s_aux = c_dt_entrega_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Data de coleta:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if ckb_produto <> "" then 
		s_aux = c_fabricante
		if s_aux = "" then s_aux = "todos"
		s = "fabricante: " & s_aux
		s_aux = c_produto
		if s_aux = "" then s_aux = "todos"
		s = s & ",&nbsp;&nbsp;produto: " & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Somente pedidos que incluam:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

    s = c_grupo
	if s = "" then 
		s = "todos"
	else
        s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Grupo(s) de Produtos:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	end if

    s = c_grupo_pedido_origem
	if s = "" then 
		s = "todos"
	else
		s = obtem_descricao_tabela_t_codigo_descricao("PedidoECommerce_Origem_Grupo", c_grupo_pedido_origem)
        s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Origem Pedido (Grupo):&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	end if

    s = c_pedido_origem
	if s = "" then 
		s = "todos"
	else
		s = obtem_descricao_tabela_t_codigo_descricao("PedidoECommerce_Origem", c_pedido_origem)
        s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Origem do Pedido:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
    end if

    s = c_empresa
	if s = "" then 
		s = "todas"
	else
		s =  obtem_apelido_empresa_NFe_emitente(c_empresa)
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			"<span class='N'>Empresa:&nbsp;</span></td><td align='left' valign='top'>" & _
			"<span class='N'>" & s & "</span></td></tr>"

	if op_forma_pagto <> "" then
		s = x_opcao_forma_pagamento(op_forma_pagto)
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Forma Pagto:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if c_forma_pagto_qtde_parc <> "" then
		s = c_forma_pagto_qtde_parc
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Nº Parcelas:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if c_cliente_cnpj_cpf <> "" then
		s = cnpj_cpf_formata(c_cliente_cnpj_cpf)
		s_aux = x_cliente_por_cnpj_cpf(c_cliente_cnpj_cpf, cadastrado)
		if Not cadastrado then s_aux = "Não Cadastrado"
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Cliente:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

    if c_cliente_uf <> "" then
		s = c_cliente_uf
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>UF do Cliente:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if ckb_visanet <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Cartão de Crédito:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><p class='N'>somente pedidos pagos usando cartão de crédito</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if
	
	if ckb_perc_RT <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Comissão:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><p class='N'>somente pedidos com comissão abaixo de " & formata_perc_RT(converte_numero(c_perc_RT)) & "%</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if c_transportadora <> "" then
		s = c_transportadora
		s_aux = iniciais_em_maiusculas(x_transportadora(c_transportadora))
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Transportadora:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if c_vendedor <> "" then
		s = c_vendedor
		s_aux = x_usuario(c_vendedor)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Vendedor:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if c_indicador <> "" then
		s = c_indicador
		s_aux = x_orcamentista_e_indicador(c_indicador)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><p class='N'>Indicador:&nbsp;</p></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if
	
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><p class='N'>Emissão:&nbsp;</p></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><p class='N'>" & formata_data_hora(Now) & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="849" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align='left'>&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="849" cellspacing="0">
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
