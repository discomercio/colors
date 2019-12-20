<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================================
'	  O R C A M E N T O V I R A R P E D I D O C O N F I R M A . A S P
'     ===============================================================
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

	dim msg_erro
	dim usuario, loja, orcamento_selecionado, tipo_cliente
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	orcamento_selecionado = Trim(request("orcamento_selecionado"))
	if (orcamento_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_NAO_ESPECIFICADO)

	dim rb_selecao_cd, c_id_nfe_emitente_selecao_manual
	rb_selecao_cd = Trim(Request("rb_selecao_cd"))
	c_id_nfe_emitente_selecao_manual = Trim(Request("c_id_nfe_emitente_selecao_manual"))

	dim alerta, alerta_aux
	alerta=""

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim blnAnalisarEndereco, blnGravouRegPai, intNsu, intNsuPai, vAnEndConfrontacao
	dim intQtdeTotalPedidosAnEndereco
	dim blnAnEnderecoCadClienteUsaEndParceiro, blnAnEnderecoEndEntregaUsaEndParceiro
	blnAnalisarEndereco = False
	blnAnEnderecoCadClienteUsaEndParceiro = False
	blnAnEnderecoEndEntregaUsaEndParceiro = False

'	FORMA DE PAGAMENTO (NOVA VERSÃO)
	dim rb_forma_pagto, op_av_forma_pagto, c_pc_qtde, c_pc_valor, c_pc_maquineta_qtde, c_pc_maquineta_valor
	dim op_pu_forma_pagto, c_pu_valor, c_pu_vencto_apos
	dim op_pce_entrada_forma_pagto, c_pce_entrada_valor, op_pce_prestacao_forma_pagto, c_pce_prestacao_qtde, c_pce_prestacao_valor, c_pce_prestacao_periodo
	dim op_pse_prim_prest_forma_pagto, c_pse_prim_prest_valor, c_pse_prim_prest_apos, op_pse_demais_prest_forma_pagto, c_pse_demais_prest_qtde, c_pse_demais_prest_valor, c_pse_demais_prest_periodo
	dim vlTotalFormaPagto
    dim s_nf_texto, s_num_pedido_compra

	vlTotalFormaPagto = 0
	if alerta = "" then
		rb_forma_pagto = Trim(Request.Form("rb_forma_pagto"))
		if rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA then
			op_av_forma_pagto = Trim(Request.Form("op_av_forma_pagto"))
			if op_av_forma_pagto = "" then alerta = "Indique a forma de pagamento (à vista)."
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELA_UNICA then
			op_pu_forma_pagto = Trim(Request.Form("op_pu_forma_pagto"))
			c_pu_valor = Trim(Request.Form("c_pu_valor"))
			c_pu_vencto_apos = Trim(Request.Form("c_pu_vencto_apos"))
			if op_pu_forma_pagto = "" then
				alerta = "Indique a forma de pagamento da parcela única."
			elseif c_pu_valor = "" then
				alerta = "Indique o valor da parcela única."
			elseif converte_numero(c_pu_valor) <= 0 then
				alerta = "Valor da parcela única é inválido."
			elseif c_pu_vencto_apos = "" then
				alerta = "Indique o intervalo de vencimento da parcela única."
			elseif converte_numero(c_pu_vencto_apos) <= 0 then
				alerta = "Intervalo de vencimento da parcela única é inválido."
				end if
			if alerta = "" then vlTotalFormaPagto = converte_numero(c_pu_valor)
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO then
			c_pc_qtde = Trim(Request.Form("c_pc_qtde"))
			c_pc_valor = Trim(Request.Form("c_pc_valor"))
			if c_pc_qtde = "" then
				alerta = "Indique a quantidade de parcelas (parcelado no cartão [internet])."
			elseif c_pc_valor = "" then
				alerta = "Indique o valor da parcela (parcelado no cartão [internet])."
			elseif converte_numero(c_pc_qtde) < 1 then
				alerta = "Quantidade de parcelas inválida (parcelado no cartão [internet])."
			elseif converte_numero(c_pc_valor) <= 0 then
				alerta = "Valor de parcela inválido (parcelado no cartão [internet])."
				end if
			if alerta = "" then vlTotalFormaPagto = converte_numero(c_pc_qtde) * converte_numero(c_pc_valor)
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
			c_pc_maquineta_qtde = Trim(Request.Form("c_pc_maquineta_qtde"))
			c_pc_maquineta_valor = Trim(Request.Form("c_pc_maquineta_valor"))
			if c_pc_maquineta_qtde = "" then
				alerta = "Indique a quantidade de parcelas (parcelado no cartão [maquineta])."
			elseif c_pc_maquineta_valor = "" then
				alerta = "Indique o valor da parcela (parcelado no cartão [maquineta])."
			elseif converte_numero(c_pc_maquineta_qtde) < 1 then
				alerta = "Quantidade de parcelas inválida (parcelado no cartão [maquineta])."
			elseif converte_numero(c_pc_maquineta_valor) <= 0 then
				alerta = "Valor de parcela inválido (parcelado no cartão [maquineta])."
				end if
			if alerta = "" then vlTotalFormaPagto = converte_numero(c_pc_maquineta_qtde) * converte_numero(c_pc_maquineta_valor)
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
			op_pce_entrada_forma_pagto = Trim(Request.Form("op_pce_entrada_forma_pagto"))
			c_pce_entrada_valor = Trim(Request.Form("c_pce_entrada_valor"))
			op_pce_prestacao_forma_pagto = Trim(Request.Form("op_pce_prestacao_forma_pagto"))
			c_pce_prestacao_qtde = Trim(Request.Form("c_pce_prestacao_qtde"))
			c_pce_prestacao_valor = Trim(Request.Form("c_pce_prestacao_valor"))
			c_pce_prestacao_periodo = Trim(Request.Form("c_pce_prestacao_periodo"))
			if op_pce_entrada_forma_pagto = "" then
				alerta = "Indique a forma de pagamento da entrada (parcelado com entrada)."
			elseif c_pce_entrada_valor = "" then
				alerta = "Indique o valor da entrada (parcelado com entrada)."
			elseif converte_numero(c_pce_entrada_valor) <= 0 then
				alerta = "Valor da entrada inválido (parcelado com entrada)."
			elseif op_pce_prestacao_forma_pagto = "" then
				alerta = "Indique a forma de pagamento das prestações (parcelado com entrada)."
			elseif c_pce_prestacao_qtde = "" then
				alerta = "Indique a quantidade de prestações (parcelado com entrada)."
			elseif converte_numero(c_pce_prestacao_qtde) <= 0 then
				alerta = "Quantidade de prestações inválida (parcelado com entrada)."
			elseif c_pce_prestacao_valor = "" then
				alerta = "Indique o valor da prestação (parcelado com entrada)."
			elseif converte_numero(c_pce_prestacao_valor) <= 0 then
				alerta = "Valor de prestação inválido (parcelado com entrada)."
			elseif c_pce_prestacao_periodo = "" then
				alerta = "Indique o intervalo de vencimento entre as parcelas (parcelado com entrada)."
			elseif converte_numero(c_pce_prestacao_periodo) <= 0 then
				alerta = "Intervalo de vencimento inválido (parcelado com entrada)."
				end if
			if alerta = "" then
				vlTotalFormaPagto = converte_numero(c_pce_entrada_valor) + (converte_numero(c_pce_prestacao_qtde) * converte_numero(c_pce_prestacao_valor))
				end if
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
			op_pse_prim_prest_forma_pagto = Trim(Request.Form("op_pse_prim_prest_forma_pagto"))
			c_pse_prim_prest_valor = Trim(Request.Form("c_pse_prim_prest_valor"))	
			c_pse_prim_prest_apos = Trim(Request.Form("c_pse_prim_prest_apos"))	
			op_pse_demais_prest_forma_pagto = Trim(Request.Form("op_pse_demais_prest_forma_pagto"))
			c_pse_demais_prest_qtde = Trim(Request.Form("c_pse_demais_prest_qtde"))
			c_pse_demais_prest_valor = Trim(Request.Form("c_pse_demais_prest_valor"))
			c_pse_demais_prest_periodo = Trim(Request.Form("c_pse_demais_prest_periodo"))
			if op_pse_prim_prest_forma_pagto = "" then
				alerta = "Indique a forma de pagamento da 1ª prestação (parcelado sem entrada)."
			elseif c_pse_prim_prest_valor = "" then
				alerta = "Indique o valor da 1ª prestação (parcelado sem entrada)."
			elseif converte_numero(c_pse_prim_prest_valor) <= 0 then
				alerta = "Valor da 1ª prestação inválido (parcelado sem entrada)."
			elseif c_pse_prim_prest_apos = "" then
				alerta = "Indique o intervalo de vencimento da 1ª parcela (parcelado sem entrada)."
			elseif converte_numero(c_pse_prim_prest_apos) <= 0 then
				alerta = "Intervalo de vencimento da 1ª parcela é inválido (parcelado sem entrada)."
			elseif op_pse_demais_prest_forma_pagto = "" then
				alerta = "Indique a forma de pagamento das demais prestações (parcelado sem entrada)."
			elseif c_pse_demais_prest_qtde = "" then
				alerta = "Indique a quantidade das demais prestações (parcelado sem entrada)."
			elseif converte_numero(c_pse_demais_prest_qtde) <= 0 then
				alerta = "Quantidade de prestações inválida (parcelado sem entrada)."
			elseif c_pse_demais_prest_valor = "" then
				alerta = "Indique o valor das demais prestações (parcelado sem entrada)."
			elseif converte_numero(c_pse_demais_prest_valor) <= 0 then
				alerta = "Valor de prestação inválido (parcelado sem entrada)."
			elseif c_pse_demais_prest_periodo = "" then
				alerta = "Indique o intervalo de vencimento entre as parcelas (parcelado sem entrada)."
			elseif converte_numero(c_pse_demais_prest_periodo) <= 0 then
				alerta = "Intervalo de vencimento inválido (parcelado sem entrada)."
				end if
			if alerta = "" then
				vlTotalFormaPagto = converte_numero(c_pse_prim_prest_valor) + (converte_numero(c_pse_demais_prest_qtde) * converte_numero(c_pse_demais_prest_valor))
				end if
		else
			alerta = "É obrigatório especificar a forma de pagamento"
			end if
		end if
	
	dim c_custoFinancFornecTipoParcelamento, c_custoFinancFornecQtdeParcelas, coeficiente
	dim c_custoFinancFornecTipoParcelamentoConferencia, c_custoFinancFornecQtdeParcelasConferencia
	c_custoFinancFornecTipoParcelamento = Trim(Request.Form("c_custoFinancFornecTipoParcelamento"))
	c_custoFinancFornecQtdeParcelas = Trim(Request.Form("c_custoFinancFornecQtdeParcelas"))
	if rb_forma_pagto=COD_FORMA_PAGTO_A_VISTA then
		c_custoFinancFornecTipoParcelamentoConferencia=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA
		c_custoFinancFornecQtdeParcelasConferencia="0"
	elseif rb_forma_pagto=COD_FORMA_PAGTO_PARCELA_UNICA then
		c_custoFinancFornecTipoParcelamentoConferencia=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA
		c_custoFinancFornecQtdeParcelasConferencia="1"
	elseif rb_forma_pagto=COD_FORMA_PAGTO_PARCELADO_CARTAO then
		c_custoFinancFornecTipoParcelamentoConferencia=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA
		c_custoFinancFornecQtdeParcelasConferencia=c_pc_qtde
	elseif rb_forma_pagto=COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
		c_custoFinancFornecTipoParcelamentoConferencia=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA
		c_custoFinancFornecQtdeParcelasConferencia=c_pc_maquineta_qtde
	elseif rb_forma_pagto=COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
		c_custoFinancFornecTipoParcelamentoConferencia=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA
		c_custoFinancFornecQtdeParcelasConferencia=c_pce_prestacao_qtde
	elseif rb_forma_pagto=COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
		c_custoFinancFornecTipoParcelamentoConferencia=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA
		c_custoFinancFornecQtdeParcelasConferencia=Cstr(converte_numero(c_pse_demais_prest_qtde)+1)
	else
		c_custoFinancFornecTipoParcelamentoConferencia=""
		c_custoFinancFornecQtdeParcelasConferencia="0"
		end if

	if alerta = "" then
		if c_custoFinancFornecTipoParcelamentoConferencia<>c_custoFinancFornecTipoParcelamento then
			alerta="Foi detectada uma inconsistência no tipo de parcelamento do pagamento (código esperado=" & c_custoFinancFornecTipoParcelamentoConferencia & ", código lido=" & c_custoFinancFornecTipoParcelamento & ")"
		elseif converte_numero(c_custoFinancFornecQtdeParcelasConferencia)<>converte_numero(c_custoFinancFornecQtdeParcelas) then
			alerta="Foi detectada uma inconsistência na quantidade de parcelas de pagamento (qtde esperada=" & c_custoFinancFornecQtdeParcelasConferencia & ", qtde lida=" & c_custoFinancFornecQtdeParcelas & ")"
			end if
		end if


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim r_orcamento, v_orcamento_item
	if Not le_orcamento(orcamento_selecionado, r_orcamento, msg_erro) then 
		alerta = msg_erro
	else
		if Trim(r_orcamento.loja) <> loja then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_INVALIDO)
		if r_orcamento.st_orc_virou_pedido = 1 then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_INVALIDO)
	'	TEM ACESSO A ESTE ORÇAMENTO?
		if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then 
			if r_orcamento.vendedor <> usuario then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_INVALIDO)
			end if
		if Not le_orcamento_item(orcamento_selecionado, v_orcamento_item, msg_erro) then alerta = msg_erro
		end if

	dim r_cliente
	set r_cliente = New cl_CLIENTE
	if alerta = "" then
		if Not x_cliente_bd(r_orcamento.id_cliente, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		end if
	tipo_cliente = r_cliente.tipo
	
	dim r_orcamentista_e_indicador
	if alerta = "" then
		if Not le_orcamentista_e_indicador(r_orcamento.orcamentista, r_orcamentista_e_indicador, msg_erro) then
			alerta = "Falha ao recuperar os dados do indicador!!<br>" & msg_erro
			end if
		end if
	
	dim vl_aprov_auto_analise_credito
	vl_aprov_auto_analise_credito = 0
	
	dim vl_total_RA_liquido
	dim s, i, iv, j, k, n, opcao_venda_sem_estoque, vl_total, vl_total_NF, vl_total_RA, qtde_estoque_total_disponivel, blnAchou, blnDesativado
	dim v_desconto()
	ReDim v_desconto(0)
	v_desconto(UBound(v_desconto)) = ""

	opcao_venda_sem_estoque = Trim(request("opcao_venda_sem_estoque"))
	
	dim s_forma_pagto, s_obs1, s_obs2, s_recebido, c_perc_RT, s_etg_imediata, s_bem_uso_consumo
	s_obs1=Trim(request("c_obs1"))
	s_obs2=Trim(request("c_obs2"))
	s_recebido=Trim(request("rb_recebido"))
	s_etg_imediata=Trim(request("rb_etg_imediata"))
	s_bem_uso_consumo=Trim(request("rb_bem_uso_consumo"))
	s_forma_pagto=Trim(request("c_forma_pagto"))
	c_perc_RT = Trim(request("c_perc_RT"))
    s_nf_texto = Trim(request("c_nf_texto"))
    s_num_pedido_compra = Trim(request("c_num_pedido_compra"))

	dim perc_RT
	perc_RT = converte_numero(c_perc_RT)

	if alerta = "" then
		if (perc_RT < 0) Or (perc_RT > 100) then
			alerta = "Percentual de comissão inválido."
			end if
		end if

	dim rCD
	set rCD = obtem_perc_max_comissao_e_desconto_por_loja(r_orcamento.loja)

'	OBTÉM A RELAÇÃO DE MEIOS DE PAGAMENTO PREFERENCIAIS (QUE FAZEM USO O PERCENTUAL DE COMISSÃO+DESCONTO NÍVEL 2)
	dim rP, vMPN2
	set rP = get_registro_t_parametro(ID_PARAMETRO_PercMaxComissaoEDesconto_Nivel2_MeiosPagto)
	if Trim("" & rP.id) <> "" then
		vMPN2 = Split(rP.campo_texto, ",")
		for i=Lbound(vMPN2) to Ubound(vMPN2)
			vMPN2(i) = Trim("" & vMPN2(i))
			next
	else
		redim vMPN2(0)
		vMPN2(0) = ""
		end if

	if alerta = "" then
		if perc_RT > rCD.perc_max_comissao then
			alerta = "Percentual de comissão excede o máximo permitido."
			end if
		end if

	dim s_instalador_instala
	s_instalador_instala = Trim(Request.Form("rb_instalador_instala"))
	
	dim rb_garantia_indicador
	rb_garantia_indicador = Trim(Request.Form("rb_garantia_indicador"))
	
	dim v_item
	redim v_item(0)
	set v_item(Ubound(v_item)) = New cl_ITEM_PEDIDO_NOVO
	v_item(Ubound(v_item)).produto = ""
	
	n = Request.Form("c_produto").Count
	for i = 1 to n
		s=Trim(Request.Form("c_produto")(i))
		if s <> "" then
			if Trim(v_item(ubound(v_item)).produto) <> "" then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ITEM_PEDIDO_NOVO
				end if
			with v_item(ubound(v_item))
				.produto=Ucase(Trim(Request.Form("c_produto")(i)))
				s=retorna_so_digitos(Request.Form("c_fabricante")(i))
				.fabricante=normaliza_codigo(s, TAM_MIN_FABRICANTE)
				s = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
				s=Trim(Request.Form("c_vl_unitario")(i))
				.preco_venda=converte_numero(s)
				if (r_orcamento.permite_RA_status = 1) Or (r_orcamento.st_violado_permite_RA_status = 1) then
					s=Trim(Request.Form("c_vl_NF")(i))
					.preco_NF=converte_numero(s)
				else
					.preco_NF = .preco_venda
					end if
				.qtde_estoque_total_disponivel = 0
				.qtde_estoque_vendido = 0
				.qtde_estoque_sem_presenca = 0
				end with
			end if
		next

	dim blnTemRA
	blnTemRA = False
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			if Trim("" & v_item(i).produto) <> "" then
				if v_item(i).preco_NF <> v_item(i).preco_venda then
					blnTemRA = True
					exit for
					end if
				end if
			next
		end if

'	CUSTO FINANCEIRO FORNECEDOR
	if alerta = "" then
		if (c_custoFinancFornecTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA) And _
		   (c_custoFinancFornecTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) And _
		   (c_custoFinancFornecTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) then
			alerta = "A forma de pagamento não foi informada (à vista, com entrada, sem entrada)."
			end if
		end if
	
	if alerta = "" then
		if (c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) Or _
		   (c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) then
			if converte_numero(c_custoFinancFornecQtdeParcelas) <= 0 then
				alerta = "Não foi informada a quantidade de parcelas para a forma de pagamento selecionada (" & descricaoCustoFinancFornecTipoParcelamento(c_custoFinancFornecTipoParcelamento) &  ")"
				end if
			end if
		end if
	
'	CALCULA O VALOR TOTAL DO PEDIDO
	if alerta = "" then
		vl_total = 0
		vl_total_NF = 0
		vl_total_RA = 0
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .produto <> "" then 
					vl_total = vl_total + (.qtde * .preco_venda)
					vl_total_NF = vl_total_NF + (.qtde * .preco_NF)
					end if
				end with
			next
		vl_total_RA = vl_total_NF - vl_total
		end if
	
'	ANALISA O PERCENTUAL DE COMISSÃO+DESCONTO
	dim perc_comissao_e_desconto_a_utilizar
	dim s_pg, blnPreferencial
	dim vlNivel1, vlNivel2
	if tipo_cliente = ID_PJ then
		perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_pj
	else
		perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto
		end if

	if alerta="" then
		if rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA then
			s_pg = Trim(op_av_forma_pagto)
			if s_pg <> "" then
				for i=Lbound(vMPN2) to Ubound(vMPN2)
				'	O meio de pagamento selecionado é um dos preferenciais
					if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
						if tipo_cliente = ID_PJ then
							perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
						else
							perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
							end if
						exit for
						end if
					next
				end if
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELA_UNICA then
			s_pg = Trim(op_pu_forma_pagto)
			if s_pg <> "" then
				for i=Lbound(vMPN2) to Ubound(vMPN2)
				'	O meio de pagamento selecionado é um dos preferenciais
					if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
						if tipo_cliente = ID_PJ then
							perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
						else
							perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
							end if
						exit for
						end if
					next
				end if
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO then
			s_pg = Trim(ID_FORMA_PAGTO_CARTAO)
			if s_pg <> "" then
				for i=Lbound(vMPN2) to Ubound(vMPN2)
				'	O meio de pagamento selecionado é um dos preferenciais
					if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
						if tipo_cliente = ID_PJ then
							perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
						else
							perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
							end if
						exit for
						end if
					next
				end if
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
			s_pg = Trim(ID_FORMA_PAGTO_CARTAO_MAQUINETA)
			if s_pg <> "" then
				for i=Lbound(vMPN2) to Ubound(vMPN2)
				'	O meio de pagamento selecionado é um dos preferenciais
					if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
						if tipo_cliente = ID_PJ then
							perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
						else
							perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
							end if
						exit for
						end if
					next
				end if
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
		'	Identifica e contabiliza o valor da entrada
			blnPreferencial = False
			s_pg = Trim(op_pce_entrada_forma_pagto)
			if s_pg <> "" then
				for i=Lbound(vMPN2) to Ubound(vMPN2)
				'	O meio de pagamento selecionado é um dos preferenciais
					if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
						blnPreferencial = True
						exit for
						end if
					next
				end if
			
			if blnPreferencial then
				vlNivel2 = converte_numero(c_pce_entrada_valor)
			else
				vlNivel1 = converte_numero(c_pce_entrada_valor)
				end if
			
		'	Identifica e contabiliza o valor das parcelas
			blnPreferencial = False
			s_pg = Trim(op_pce_prestacao_forma_pagto)
			if s_pg <> "" then
				for i=Lbound(vMPN2) to Ubound(vMPN2)
				'	O meio de pagamento selecionado é um dos preferenciais
					if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
						blnPreferencial = True
						exit for
						end if
					next
				end if
			
			if blnPreferencial then
				vlNivel2 = vlNivel2 + converte_numero(c_pce_prestacao_qtde) * converte_numero(c_pce_prestacao_valor)
			else
				vlNivel1 = vlNivel1 + converte_numero(c_pce_prestacao_qtde) * converte_numero(c_pce_prestacao_valor)
				end if
		
		'	O montante a pagar por meio de pagamento preferencial é maior que 50% do total?
			if vlNivel2 > (vl_total/2) then
				if tipo_cliente = ID_PJ then
					perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
				else
					perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
					end if
				end if
			
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
		'	Identifica e contabiliza o valor da 1ª parcela
			blnPreferencial = False
			s_pg = Trim(op_pse_prim_prest_forma_pagto)
			if s_pg <> "" then
				for i=Lbound(vMPN2) to Ubound(vMPN2)
				'	O meio de pagamento selecionado é um dos preferenciais
					if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
						blnPreferencial = True
						exit for
						end if
					next
				end if
			
			if blnPreferencial then
				vlNivel2 = converte_numero(c_pse_prim_prest_valor)
			else
				vlNivel1 = converte_numero(c_pse_prim_prest_valor)
				end if
			
		'	Identifica e contabiliza o valor das parcelas
			blnPreferencial = False
			s_pg = Trim(op_pse_demais_prest_forma_pagto)
			if s_pg <> "" then
				for i=Lbound(vMPN2) to Ubound(vMPN2)
				'	O meio de pagamento selecionado é um dos preferenciais
					if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
						blnPreferencial = True
						exit for
						end if
					next
				end if
			
			if blnPreferencial then
				vlNivel2 = vlNivel2 + converte_numero(c_pse_demais_prest_qtde) * converte_numero(c_pse_demais_prest_valor)
			else
				vlNivel1 = vlNivel1 + converte_numero(c_pse_demais_prest_qtde) * converte_numero(c_pse_demais_prest_valor)
				end if
			
		'	O montante a pagar por meio de pagamento preferencial é maior que 50% do total?
			if vlNivel2 > (vl_total/2) then
				if tipo_cliente = ID_PJ then
					perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
				else
					perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
					end if
				end if
			end if
		end if
	
'	CONSISTÊNCIA PARA VALOR ZERADO
	if alerta="" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .preco_venda <= 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto '" & .produto & "' está com valor de venda zerado!"
				elseif ((r_orcamento.permite_RA_status = 1) Or (r_orcamento.st_violado_permite_RA_status = 1)) And (.preco_NF <= 0) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto '" & .produto & "' está com preço zerado!"
					end if
				end with
			next
		end if

'	VERIFICA CADA UM DOS PRODUTOS SELECIONADOS
	dim desc_dado_arredondado
	if alerta="" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				s = "SELECT " & _
						"*" & _
					" FROM t_PRODUTO" & _
						" INNER JOIN t_PRODUTO_LOJA" & _
							" ON (t_PRODUTO.fabricante=t_PRODUTO_LOJA.fabricante) AND (t_PRODUTO.produto=t_PRODUTO_LOJA.produto)" & _
						" INNER JOIN t_FABRICANTE" & _
							" ON (t_PRODUTO.fabricante=t_FABRICANTE.fabricante)" & _
					" WHERE" & _
						" (t_PRODUTO.fabricante='" & .fabricante & "')" & _
						" AND (t_PRODUTO.produto='" & .produto & "')" & _
						" AND (loja='" & loja & "')"
				set rs = cn.execute(s)
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " NÃO está cadastrado para a loja " & loja
				else
					.preco_lista = rs("preco_lista")
					.margem = rs("margem")
					.desc_max = rs("desc_max")
					.comissao = rs("comissao")
					.preco_fabricante = rs("preco_fabricante")
					.vl_custo2 = rs("vl_custo2")
					.descricao = Trim("" & rs("descricao"))
					.descricao_html = Trim("" & rs("descricao_html"))
					.ean = Trim("" & rs("ean"))
					.grupo = Trim("" & rs("grupo"))
					.peso = rs("peso")
					.qtde_volumes = rs("qtde_volumes")
					.markup_fabricante = rs("markup")
					.cubagem = rs("cubagem")
					.ncm = Trim("" & rs("ncm"))
					.cst = Trim("" & rs("cst"))
					.descontinuado = Trim("" & rs("descontinuado"))

					.custoFinancFornecPrecoListaBase = .preco_lista
					if c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA then
						coeficiente = 1
					else
						s = "SELECT " & _
								"*" & _
							" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
							" WHERE" & _
								" (fabricante = '" & .fabricante & "')" & _
								" AND (tipo_parcelamento = '" & c_custoFinancFornecTipoParcelamento & "')" & _
								" AND (qtde_parcelas = " & c_custoFinancFornecQtdeParcelas & ")"
						set rs2 = cn.execute(s)
						if rs2.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Opção de parcelamento não disponível para fornecedor " & .fabricante & ": " & decodificaCustoFinancFornecQtdeParcelas(c_custoFinancFornecTipoParcelamento, c_custoFinancFornecQtdeParcelas) & " parcela(s)"
						else
							coeficiente = converte_numero(rs2("coeficiente"))
							.preco_lista=converte_numero(formata_moeda(coeficiente*.preco_lista))
							end if
						end if
					.custoFinancFornecCoeficiente = coeficiente
					
					if .preco_lista = 0 then 
						.desc_dado = 0
						desc_dado_arredondado = 0
					else
						.desc_dado = 100*(.preco_lista-.preco_venda)/.preco_lista
						desc_dado_arredondado = converte_numero(formata_perc_desc(.desc_dado))
						end if
					
					if desc_dado_arredondado > perc_comissao_e_desconto_a_utilizar then
						if rs.State <> 0 then rs.Close
						s = "SELECT " & _
								"*" & _
							" FROM t_DESCONTO" & _
							" WHERE" & _
								" (usado_status=0)" & _
								" AND (cancelado_status=0)" & _
								" AND (id_cliente='" & r_orcamento.id_cliente & "')" & _
								" AND (fabricante='" & .fabricante & "')" & _
								" AND (produto='" & .produto & "')" & _
								" AND (loja='" & loja & "')" & _
								" AND (data >= " & bd_formata_data_hora(Now-converte_min_to_dec(TIMEOUT_DESCONTO_EM_MIN)) & ")" & _
							" ORDER BY" & _
								" data DESC"
						set rs=cn.execute(s)
						if rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": desconto de " & formata_perc_desc(.desc_dado) & "% excede o máximo permitido."
						else
							if .desc_dado > rs("desc_max") then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": desconto de " & formata_perc_desc(.desc_dado) & "% excede o máximo autorizado."
							else
								.abaixo_min_status=1
								.abaixo_min_autorizacao=Trim("" & rs("id"))
								.abaixo_min_autorizador=Trim("" & rs("autorizador"))
								.abaixo_min_superv_autorizador=Trim("" & rs("supervisor_autorizador"))
								If v_desconto(UBound(v_desconto)) <> "" Then
									ReDim Preserve v_desconto(UBound(v_desconto) + 1)
									v_desconto(UBound(v_desconto)) = ""
									End If
								v_desconto(UBound(v_desconto)) = Trim("" & rs("id"))
								end if
							end if
						end if
					end if
				rs.Close
				end with
			next
		end if
	
'	RECUPERA OS PRODUTOS QUE O CLIENTE CONCORDOU EM COMPRAR MESMO SEM PRESENÇA NO ESTOQUE.
	dim v_spe
	redim v_spe(0)
	set v_spe(0) = New cl_CTRL_ESTOQUE_PEDIDO_ITEM_NOVO
	if (alerta="") And (opcao_venda_sem_estoque<>"") then
		n=Request.Form("c_spe_produto").Count
		for i=1 to n
			s=Trim(Request.Form("c_spe_produto")(i))
			if s<>"" then
				if Trim(v_spe(ubound(v_spe)).produto) <> "" then
					redim preserve v_spe(ubound(v_spe)+1)
					set v_spe(ubound(v_spe)) = New cl_CTRL_ESTOQUE_PEDIDO_ITEM_NOVO
					end if
				with v_spe(ubound(v_spe))
					.produto=Ucase(Trim(Request.Form("c_spe_produto")(i)))
					s=retorna_so_digitos(Request.Form("c_spe_fabricante")(i))
					.fabricante=normaliza_codigo(s, TAM_MIN_FABRICANTE)
					s = Trim(Request.Form("c_spe_qtde_solicitada")(i))
					if IsNumeric(s) then .qtde_solicitada = CLng(s) else .qtde_solicitada = 0
					s = Trim(Request.Form("c_spe_qtde_estoque")(i))
					if IsNumeric(s) then .qtde_estoque = CLng(s) else .qtde_estoque = 0
					end with
				end if
			next
		end if
	
	
'	LÓGICA P/ CONSUMO DO ESTOQUE
	dim tipo_pessoa
	dim descricao_tipo_pessoa
	tipo_pessoa = multi_cd_regra_determina_tipo_pessoa(r_cliente.tipo, r_cliente.contribuinte_icms_status, r_cliente.produtor_rural_status)
	descricao_tipo_pessoa = descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa)

	dim id_nfe_emitente_selecao_manual
	dim vProdRegra, iRegra, iCD, iItem, idxItem, qtde_CD_ativo
	dim qtde_spe, qtde_estoque_vendido_aux, qtde_estoque_sem_presenca_aux, total_estoque_vendido, total_estoque_sem_presenca

	if alerta="" then
		if rb_selecao_cd = MODO_SELECAO_CD__MANUAL then
			id_nfe_emitente_selecao_manual = converte_numero(c_id_nfe_emitente_selecao_manual)
			if id_nfe_emitente_selecao_manual = 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O CD selecionado manualmente é inválido"
				end if
		else
			id_nfe_emitente_selecao_manual = 0
			end if
		end if
	
	if alerta="" then
		'PREPARA O VETOR PARA RECUPERAR AS REGRAS DE CONSUMO DO ESTOQUE ASSOCIADAS AOS PRODUTOS
		redim vProdRegra(0)
		inicializa_cl_PEDIDO_SELECAO_PRODUTO_REGRA vProdRegra(UBound(vProdRegra))
		for i=LBound(v_item) to UBound(v_item)
			if vProdRegra(UBound(vProdRegra)).produto <> "" then
				redim preserve vProdRegra(UBound(vProdRegra)+1)
				inicializa_cl_PEDIDO_SELECAO_PRODUTO_REGRA vProdRegra(UBound(vProdRegra))
				end if
			vProdRegra(UBound(vProdRegra)).fabricante = v_item(i).fabricante
			vProdRegra(UBound(vProdRegra)).produto =v_item(i).produto
			next
		
		'RECUPERA AS REGRAS DE CONSUMO DO ESTOQUE ASSOCIADAS AOS PRODUTOS
		if Not obtemCtrlEstoqueProdutoRegra(r_cliente.uf, r_cliente.tipo, r_cliente.contribuinte_icms_status, r_cliente.produtor_rural_status, vProdRegra, msg_erro) then
			alerta = "Falha ao tentar obter a(s) regra(s) de consumo do estoque"
			if msg_erro <> "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & msg_erro
				end if
			end if
		end if 'if alerta=""
	
	if alerta="" then
		'VERIFICA SE HOUVE ERRO NA LEITURA DAS REGRAS DE CONSUMO DO ESTOQUE ASSOCIADAS AOS PRODUTOS
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			if Trim(vProdRegra(iRegra).produto) <> "" then
				if Not vProdRegra(iRegra).st_regra_ok then
					if Trim(vProdRegra(iRegra).msg_erro) <> "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & vProdRegra(iRegra).msg_erro
					else
						alerta=texto_add_br(alerta)
						alerta=alerta & "Falha desconhecida na leitura da regra de consumo do estoque para o produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " (UF: '" & r_cliente.uf & "', tipo de pessoa: '" & descricao_tipo_pessoa & "')"
						end if
					end if
				end if
			next
		end if 'if alerta=""
	
	if alerta="" then
		'VERIFICA SE AS REGRAS ASSOCIADAS AOS PRODUTOS ESTÃO OK
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			if Trim(vProdRegra(iRegra).produto) <> "" then
				if converte_numero(vProdRegra(iRegra).regra.id) = 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " não possui regra de consumo do estoque associada"
				elseif vProdRegra(iRegra).regra.st_inativo = 1 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " está desativada"
				elseif vProdRegra(iRegra).regra.regraUF.st_inativo = 1 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " está bloqueada para a UF '" & r_cliente.uf & "'"
				elseif vProdRegra(iRegra).regra.regraUF.regraPessoa.st_inativo = 1 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " está bloqueada para clientes '" & descricao_tipo_pessoa & "' da UF '" & r_cliente.uf & "'"
				elseif converte_numero(vProdRegra(iRegra).regra.regraUF.regraPessoa.spe_id_nfe_emitente) = 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " não especifica nenhum CD para aguardar produtos sem presença no estoque para clientes '" & descricao_tipo_pessoa & "' da UF '" & r_cliente.uf & "'"
				else
					qtde_CD_ativo = 0
					for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
						if converte_numero(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente) > 0 then
							if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0 then
								qtde_CD_ativo = qtde_CD_ativo + 1
								end if
							end if
						next
					if qtde_CD_ativo = 0 then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " não especifica nenhum CD ativo para clientes '" & descricao_tipo_pessoa & "' da UF '" & r_cliente.uf & "'"
						end if
					end if
				end if
			next
		end if 'if alerta=""
	
	'NO CASO DE SELEÇÃO MANUAL DO CD, VERIFICA SE O CD SELECIONADO ESTÁ HABILITADO EM TODAS AS REGRAS
	if alerta="" then
		if id_nfe_emitente_selecao_manual <> 0 then
			alerta_aux = ""
			for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
				blnAchou = False
				blnDesativado = False
				if Trim(vProdRegra(iRegra).produto) <> "" then
					for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
						if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual then
							blnAchou = True
							if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 1 then blnDesativado = True
							exit for
							end if
						next
					end if

				if Not blnAchou then
					alerta_aux=texto_add_br(alerta_aux)
					alerta_aux=alerta_aux & "Produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & ": regra '" & vProdRegra(iRegra).regra.apelido & "' (Id=" & vProdRegra(iRegra).regra.id & ") não permite o CD '" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_selecao_manual) & "'"
				elseif blnDesativado then
					alerta_aux=texto_add_br(alerta_aux)
					alerta_aux=alerta_aux & "Produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & ": regra '" & vProdRegra(iRegra).regra.apelido & "' (Id=" & vProdRegra(iRegra).regra.id & ") define o CD '" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_selecao_manual) & "' como 'desativado'"
					end if
				next

			if alerta_aux <> "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O CD selecionado manualmente não pode ser usado devido aos seguintes motivos:"
				alerta=texto_add_br(alerta)
				alerta=alerta & alerta_aux
				end if
			end if
		end if
	
	dim erro_produto_indisponivel
	if alerta="" then
		'OBTÉM DISPONIBILIDADE DO PRODUTO NO ESTOQUE
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			if Trim(vProdRegra(iRegra).produto) <> "" then
				for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
					if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
						( (id_nfe_emitente_selecao_manual = 0) Or (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) ) then
						'VERIFICA SE O CD ESTÁ HABILITADO
						if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0 then
							idxItem = Lbound(v_item) - 1
							for iItem=Lbound(v_item) to Ubound(v_item)
								if (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) then
									idxItem = iItem
									exit for
									end if
								next
							if idxItem < Lbound(v_item) then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Falha ao localizar o produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " na lista de produtos a ser processada"
							else
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.fabricante = v_item(idxItem).fabricante
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.produto = v_item(idxItem).produto
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.descricao = v_item(idxItem).descricao
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.descricao_html = v_item(idxItem).descricao_html
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = v_item(idxItem).qtde
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque = 0
								if Not estoque_verifica_disponibilidade_integral_v2(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente, vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque) then
									alerta=texto_add_br(alerta)
									alerta=alerta & "Falha ao tentar consultar disponibilidade no estoque do produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto
									end if
								end if
							end if
						end if

					if alerta <> "" then exit for
					next
				end if

			if alerta <> "" then exit for
			next
		end if 'if alerta=""
	
'	HÁ PRODUTO C/ ESTOQUE INSUFICIENTE (SOMANDO-SE O ESTOQUE DE TODAS AS EMPRESAS CANDIDATAS)
	erro_produto_indisponivel = False
	if alerta="" then
		for iItem=Lbound(v_item) to Ubound(v_item)
			if Trim(v_item(iItem).produto) <> "" then
				qtde_estoque_total_disponivel = 0
				for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
					if Trim(vProdRegra(iRegra).produto) <> "" then
						for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
							if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
								( (id_nfe_emitente_selecao_manual = 0) Or (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) ) then
								'VERIFICA SE O CD ESTÁ HABILITADO
								if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0 then
									if (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) then
										qtde_estoque_total_disponivel = qtde_estoque_total_disponivel + vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque
										end if
									end if
								end if
							next
						end if
					next

				v_item(iItem).qtde_estoque_total_disponivel = qtde_estoque_total_disponivel

				if v_item(iItem).qtde > qtde_estoque_total_disponivel then
					erro_produto_indisponivel = True
					end if
				end if
			next
		end if 'if alerta=""
	
	if alerta = "" then
		if erro_produto_indisponivel then
			for i=Lbound(v_item) to Ubound(v_item)
				if v_item(i).qtde > v_item(i).qtde_estoque_total_disponivel then
					if (opcao_venda_sem_estoque="") then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & v_item(i).produto & " do fabricante " & v_item(i).fabricante & ": falta(m) " & Cstr(Abs(v_item(i).qtde_estoque_total_disponivel-v_item(i).qtde)) & " unidade(s) no estoque."
					else
						qtde_spe = -1
						for j=Lbound(v_spe) to Ubound(v_spe)
							if (v_item(i).fabricante=v_spe(j).fabricante) And (v_item(i).produto=v_spe(j).produto) then
								qtde_spe = v_spe(j).qtde_estoque
								exit for
								end if
							next
						if qtde_spe <> v_item(i).qtde_estoque_total_disponivel then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto " & v_item(i).produto & " do fabricante " & v_item(i).fabricante & ": disponibilidade do estoque foi alterada."
							end if
						end if
					end if
				next
			end if
		end if
	
'	ANALISA A QUANTIDADE DE PEDIDOS QUE SERÃO CADASTRADOS (AUTO-SPLIT)
'	INICIALIZA O CAMPO 'qtde_solicitada', POIS ELE IRÁ CONTROLAR A QUANTIDADE A SER ALOCADA NO ESTOQUE DE CADA EMPRESA
	if alerta = "" then
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
				vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = 0
				next
			next
		end if 'if alerta=""
	
'	REALIZA A ANÁLISE DA QUANTIDADE DE PEDIDOS NECESSÁRIA (AUTO-SPLIT)
	dim qtde_a_alocar
	if alerta = "" then
		for iItem=Lbound(v_item) to Ubound(v_item)
			if Trim(v_item(iItem).produto) <> "" then
			'	OS CD'S ESTÃO ORDENADOS DE ACORDO C/ A PRIORIZAÇÃO DEFINIDA PELA REGRA DE CONSUMO DO ESTOQUE
			'	SE O PRIMEIRO CD HABILITADO NÃO PUDER ATENDER INTEGRALMENTE A QUANTIDADE SOLICITADA DO PRODUTO,
			'	A QUANTIDADE RESTANTE SERÁ CONSUMIDA DOS DEMAIS CD'S.
			'	SE HOUVER ALGUMA QUANTIDADE RESIDUAL P/ FICAR NA LISTA DE PRODUTOS SEM PRESENÇA NO ESTOQUE:
			'		1) SELEÇÃO AUTOMÁTICA DE CD: A QUANTIDADE PENDENTE FICARÁ ALOCADA NO CD DEFINIDO P/ TAL
			'		2) SELEÇÃO MANUAL DE CD: A QUANTIDADE PENDENTE FICARÁ ALOCADA NO CD SELECIONADO MANUALMENTE
				qtde_a_alocar = v_item(iItem).qtde
				for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
					if qtde_a_alocar = 0 then exit for

					if Trim(vProdRegra(iRegra).produto) <> "" then
						for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
							if qtde_a_alocar = 0 then exit for

							if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
								( (id_nfe_emitente_selecao_manual = 0) Or (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) ) then
								'VERIFICA SE O CD ESTÁ HABILITADO
								if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0 then
									if (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) then
										if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque >= qtde_a_alocar then
										'	HÁ QUANTIDADE DISPONÍVEL SUFICIENTE PARA INTEGRALMENTE
											vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = qtde_a_alocar
											qtde_a_alocar = 0
										elseif vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque > 0 then
										'	A QUANTIDADE DISPONÍVEL NO ESTOQUE É INSUFICIENTE P/ ATENDER INTEGRALMENTE À QUANTIDADE SOLICITADA,
										'	PORTANTO, A QUANTIDADE DISPONÍVEL NESTE CD SERÁ CONSUMIDA P/ ATENDER PARCIALMENTE À REQUISIÇÃO E A
										'	QUANTIDADE REMANESCENTE SERÁ ATENDIDA PELO PRÓXIMO CD DA LISTA OU ENTÃO SERÁ COLOCADA NA LISTA DE
										'	PRODUTOS SEM PRESENÇA NO ESTOQUE DO CD SELECIONADO P/ TAL NA REGRA.
											vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque
											qtde_a_alocar = qtde_a_alocar - vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque
											end if
										end if
									end if
								end if
							next
						end if
					next

			'	RESTOU SALDO A ALOCAR NA LISTA DE PRODUTOS SEM PRESENÇA NO ESTOQUE?
				if qtde_a_alocar > 0 then
				'	LOCALIZA E ALOCA A QUANTIDADE PENDENTE:
				'		1) SELEÇÃO AUTOMÁTICA DE CD: A QUANTIDADE PENDENTE FICARÁ ALOCADA NO CD DEFINIDO P/ TAL
				'		2) SELEÇÃO MANUAL DE CD: A QUANTIDADE PENDENTE FICARÁ ALOCADA NO CD SELECIONADO MANUALMENTE
					for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
						if qtde_a_alocar = 0 then exit for

						if Trim(vProdRegra(iRegra).produto) <> "" then
							for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
								if qtde_a_alocar = 0 then exit for

								if id_nfe_emitente_selecao_manual = 0 then
									'MODO DE SELEÇÃO AUTOMÁTICO
									if ( (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) ) And _
										(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
										(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = vProdRegra(iRegra).regra.regraUF.regraPessoa.spe_id_nfe_emitente) then
										vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada + qtde_a_alocar
										qtde_a_alocar = 0
										exit for
										end if
								else
									'MODO DE SELEÇÃO MANUAL
									if ( (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) ) And _
										(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
										(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) then
										vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada + qtde_a_alocar
										qtde_a_alocar = 0
										exit for
										end if
									end if
								next
							end if
						next
					end if

			'	HOUVE FALHA EM ALOCAR A QUANTIDADE REMANESCENTE?
				if qtde_a_alocar > 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Falha ao processar a alocação de produtos no estoque: restaram " & qtde_a_alocar & " unidades do produto (" & v_item(iItem).fabricante & ")" & v_item(iItem).produto & " que não puderam ser alocados na lista de produtos sem presença no estoque de nenhum CD"
					end if
				end if
			next
		end if 'if alerta=""
	
'	CONTAGEM DE EMPRESAS QUE SERÃO USADAS NO AUTO-SPLIT, OU SEJA, A QUANTIDADE DE PEDIDOS QUE SERÁ CADASTRADA, JÁ QUE CADA PEDIDO SE REFERE AO ESTOQUE DE UMA EMPRESA
	dim vEmpresaAutoSplit
	redim vEmpresaAutoSplit(0)
	vEmpresaAutoSplit(UBound(vEmpresaAutoSplit)) = 0

	dim qtde_empresa_selecionada, lista_empresa_selecionada
	qtde_empresa_selecionada = 0
	lista_empresa_selecionada = ""
	if alerta = "" then
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			if Trim(vProdRegra(iRegra).produto) <> "" then
				for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
					if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
						( (id_nfe_emitente_selecao_manual = 0) Or (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) ) then
						if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada > 0 then
							s = "|" & vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente & "|"
							if Instr(lista_empresa_selecionada, s) = 0 then
							'	SE O CD AINDA NÃO CONSTA DA LISTA, INCLUI
								qtde_empresa_selecionada = qtde_empresa_selecionada + 1
								lista_empresa_selecionada = lista_empresa_selecionada & s
								if vEmpresaAutoSplit(UBound(vEmpresaAutoSplit)) <> 0 then
									redim preserve vEmpresaAutoSplit(UBound(vEmpresaAutoSplit) + 1)
									vEmpresaAutoSplit(UBound(vEmpresaAutoSplit)) = 0
									end if
								vEmpresaAutoSplit(UBound(vEmpresaAutoSplit)) = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente
								end if
							end if
						end if
					next
				end if
			next
		end if 'if alerta=""
	
	
'	OBTÉM O VALOR LIMITE P/ APROVAÇÃO AUTOMÁTICA DA ANÁLISE DE CRÉDITO
	if alerta = "" then
		s = "SELECT nsu FROM t_CONTROLE WHERE (id_nsu = '" & ID_PARAM_CAD_VL_APROV_AUTO_ANALISE_CREDITO & "')"
		set rs = cn.execute(s)
		if Not rs.Eof then
			vl_aprov_auto_analise_credito = converte_numero(rs("nsu"))
			end if
		if rs.State <> 0 then rs.Close
		end if
	
'	RA Líquido
	dim perc_desagio_RA, perc_limite_RA_sem_desagio
	dim vl_limite_mensal, vl_limite_mensal_consumido, vl_limite_mensal_disponivel
	if alerta = "" then
		perc_desagio_RA = obtem_perc_desagio_RA_do_indicador(r_orcamento.orcamentista)
		perc_limite_RA_sem_desagio = obtem_perc_limite_RA_sem_desagio()
		vl_limite_mensal = obtem_limite_mensal_compras_do_indicador(r_orcamento.orcamentista)
		vl_limite_mensal_consumido = calcula_limite_mensal_consumido_do_indicador(r_orcamento.orcamentista, Date)
		vl_limite_mensal_disponivel = vl_limite_mensal - vl_limite_mensal_consumido
	'	POR SOLICITAÇÃO DO ROGÉRIO, A CONSISTÊNCIA DO LIMITE DE COMPRAS FOI DESATIVADA (NOV/2008)
'		if (vl_limite_mensal_disponivel - vl_total) <= 0 then
'			alerta = "Não é possível cadastrar este pedido porque excede o valor do limite mensal estabelecido para o indicador (" & r_orcamento.orcamentista & ")"
'			end if
		if rb_garantia_indicador = "" then
			alerta = "Informe se o pedido é garantido pelo indicador ou não."
			end if
		end if
	
	if alerta = "" then
		if s_etg_imediata = "" then
			alerta = "É necessário selecionar uma opção para o campo 'Entrega Imediata'."
			end if
		end if

	if alerta = "" then
		if s_bem_uso_consumo = "" then
			alerta = "É necessário informar se é Bem de Uso/Consumo."
			end if
		end if

	if alerta = "" then
		if s_instalador_instala = "" then
			alerta = "É necessário preencher o campo 'Instalador Instala'."
			end if
		end if
	
'	CONSISTÊNCIA DO VALOR TOTAL DA FORMA DE PAGAMENTO
	if alerta = "" then
		if rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA then vlTotalFormaPagto = vl_total_NF
		if Abs(vlTotalFormaPagto-vl_total_NF) > 0.1 then
			alerta = "Há divergência entre o valor total do pedido (" & SIMBOLO_MONETARIO & " " & formata_moeda(vl_total_NF) & ") e o valor total descrito através da forma de pagamento (" & SIMBOLO_MONETARIO & " " & formata_moeda(vlTotalFormaPagto) & ")!!"
			end if
		end if
	
	if alerta = "" then
		if CLng(r_orcamento.st_end_entrega) <> 0 then
			if Len(r_orcamento.EndEtg_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
				alerta = "Endereço de entrega excede o tamanho máximo permitido:<br>Tamanho atual: " & Cstr(Len(r_orcamento.EndEtg_endereco)) & " caracteres<br>Tamanho máximo: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " caracteres"
			elseif Trim(r_orcamento.EndEtg_endereco_numero) = "" then
				alerta = "O endereço de entrega deve ser corrigido, separando as informações do número e complemento nos campos adequados."
				end if
			end if
		end if
	
'	OBTENÇÃO DE TRANSPORTADORA QUE ATENDA AO CEP INFORMADO, SE HOUVER
	dim sTranspSelAutoTransportadoraId, sTranspSelAutoCep, iTranspSelAutoTipoEndereco, iTranspSelAutoStatus
	sTranspSelAutoTransportadoraId = ""
	if alerta = "" then
		if CLng(r_orcamento.st_end_entrega) <> 0 then
			if r_orcamento.EndEtg_cep <> "" then
				sTranspSelAutoTransportadoraId = obtem_transportadora_pelo_cep(retorna_so_digitos(r_orcamento.EndEtg_cep))
				if sTranspSelAutoTransportadoraId <> "" then
					sTranspSelAutoCep = retorna_so_digitos(r_orcamento.EndEtg_cep)
					iTranspSelAutoTipoEndereco = TRANSPORTADORA_SELECAO_AUTO_TIPO_ENDERECO_ENTREGA
					iTranspSelAutoStatus = TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_S
					end if
				end if
		else
			if r_cliente.cep <> "" then
				sTranspSelAutoTransportadoraId = obtem_transportadora_pelo_cep(retorna_so_digitos(r_cliente.cep))
				if sTranspSelAutoTransportadoraId <> "" then
					sTranspSelAutoCep = retorna_so_digitos(r_cliente.cep)
					iTranspSelAutoTipoEndereco = TRANSPORTADORA_SELECAO_AUTO_TIPO_ENDERECO_CLIENTE
					iTranspSelAutoStatus = TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_S
					end if
				end if
			end if
		end if
	
	
'	CADASTRA O PEDIDO E PROCESSA A MOVIMENTAÇÃO NO ESTOQUE
	if alerta="" then
		dim id_pedido, id_pedido_base, id_pedido_temp, id_pedido_temp_base, indice_pedido, indice_item, sequencia_item, s_hora_pedido, s_log, s_log_cliente_indicador, vLogAutoSplit, s_log_item_autosplit
		indice_pedido = 0
		id_pedido_base = ""
		id_pedido_temp_base = ""
		s_log=""
		s_log_cliente_indicador=""
		redim vLogAutoSplit(0)
		vLogAutoSplit(UBound(vLogAutoSplit)) = ""
		s_hora_pedido = retorna_so_digitos(formata_hora(Now))
		if Not gera_num_pedido_temp(id_pedido_temp_base, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		for iv = LBound(vEmpresaAutoSplit) to UBound(vEmpresaAutoSplit)
			if (vEmpresaAutoSplit(iv) <> 0) then
				if Not (rs Is nothing) then
					if rs.State <> 0 then rs.Close
					set rs=nothing
					end if
		
				if Not cria_recordset_pessimista(rs, msg_erro) then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
					end if

			'	Controla a quantidade de pedidos no auto-split
			'	pedido-base: indice_pedido=1
			'	pedido-filhote 'A' => indice_pedido=2
			'	pedido-filhote 'B' => indice_pedido=3
			'	etc
				indice_pedido = indice_pedido + 1
				if indice_pedido = 1 then
					id_pedido_temp = id_pedido_temp_base
				else
					id_pedido_temp = id_pedido_temp_base & gera_letra_pedido_filhote(indice_pedido-1)
					end if

				s = "SELECT * FROM t_PEDIDO WHERE pedido='X'"
				rs.Open s, cn
				rs.AddNew
				rs("pedido")=id_pedido_temp
				rs("loja")=loja
				rs("data")=Date
				rs("hora")=s_hora_pedido
				if indice_pedido = 1 then
				'	PEDIDO BASE
				'	===========
					if qtde_empresa_selecionada > 1 then rs("st_auto_split") = 1
					if Trim("" & rs("st_pagto")) <> ST_PAGTO_NAO_PAGO then
						rs("dt_st_pagto") = Date
						rs("dt_hr_st_pagto") = Now
						rs("usuario_st_pagto") = usuario
						end if
					rs("st_pagto")=ST_PAGTO_NAO_PAGO
					rs("st_recebido")=s_recebido
					rs("obs_1")=s_obs1
					rs("obs_2")=s_obs2
				'	Forma de Pagamento (nova versão)
					rs("tipo_parcelamento")=CLng(rb_forma_pagto)
					if rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA then
						rs("av_forma_pagto") = CLng(op_av_forma_pagto)
						rs("qtde_parcelas")=1
					elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELA_UNICA then
						rs("pu_forma_pagto") = CLng(op_pu_forma_pagto)
						rs("pu_valor") = converte_numero(c_pu_valor)
						rs("pu_vencto_apos") = CLng(c_pu_vencto_apos)
						rs("qtde_parcelas")=1
					elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO then
						rs("pc_qtde_parcelas") = CLng(c_pc_qtde)
						rs("pc_valor_parcela") = converte_numero(c_pc_valor)
						rs("qtde_parcelas")=CLng(c_pc_qtde)
					elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
						rs("pc_maquineta_qtde_parcelas") = CLng(c_pc_maquineta_qtde)
						rs("pc_maquineta_valor_parcela") = converte_numero(c_pc_maquineta_valor)
						rs("qtde_parcelas")=CLng(c_pc_maquineta_qtde)
					elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
						rs("pce_forma_pagto_entrada") = CLng(op_pce_entrada_forma_pagto)
						rs("pce_forma_pagto_prestacao") = CLng(op_pce_prestacao_forma_pagto)
						rs("pce_entrada_valor") = converte_numero(c_pce_entrada_valor)
						rs("pce_prestacao_qtde") = CLng(c_pce_prestacao_qtde)
						rs("pce_prestacao_valor") = converte_numero(c_pce_prestacao_valor)
						rs("pce_prestacao_periodo") = CLng(c_pce_prestacao_periodo)
					'	Entrada + Prestações
						rs("qtde_parcelas")=CLng(c_pce_prestacao_qtde)+1
					elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
						rs("pse_forma_pagto_prim_prest") = CLng(op_pse_prim_prest_forma_pagto)
						rs("pse_forma_pagto_demais_prest") = CLng(op_pse_demais_prest_forma_pagto)
						rs("pse_prim_prest_valor") = converte_numero(c_pse_prim_prest_valor)
						rs("pse_prim_prest_apos") = CLng(c_pse_prim_prest_apos)
						rs("pse_demais_prest_qtde") = CLng(c_pse_demais_prest_qtde)
						rs("pse_demais_prest_valor") = converte_numero(c_pse_demais_prest_valor)
						rs("pse_demais_prest_periodo") = CLng(c_pse_demais_prest_periodo)
					'	1ª prestação + Demais prestações
						rs("qtde_parcelas")=CLng(c_pse_demais_prest_qtde)+1
						end if
					rs("forma_pagto")=s_forma_pagto
					rs("vl_total_familia")=vl_total
					if vl_total <= vl_aprov_auto_analise_credito then
						rs("analise_credito")=Clng(COD_AN_CREDITO_OK)
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")="AUTOMÁTICO"
					elseif (rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA) And (CStr(op_av_forma_pagto) = CStr(ID_FORMA_PAGTO_DEPOSITO)) then
						rs("analise_credito")=Clng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")="AUTOMÁTICO"
					elseif Cstr(loja) = Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) And (rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA) And (CStr(op_av_forma_pagto) = Cstr(ID_FORMA_PAGTO_DINHEIRO)) then
						rs("analise_credito")=Clng(COD_AN_CREDITO_PENDENTE_VENDAS)
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")="AUTOMÁTICO"
					elseif (rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) then
						rs("analise_credito")=Clng(COD_AN_CREDITO_PENDENTE_VENDAS)
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")="AUTOMÁTICO"
						end if

			'	CUSTO FINANCEIRO FORNECEDOR
				rs("custoFinancFornecTipoParcelamento") = c_custoFinancFornecTipoParcelamento
				rs("custoFinancFornecQtdeParcelas") = c_custoFinancFornecQtdeParcelas
				rs("vl_total_NF") = vl_total_NF
				rs("vl_total_RA") = vl_total_RA
				rs("perc_RT") = perc_RT
				rs("perc_desagio_RA") = perc_desagio_RA
				rs("perc_limite_RA_sem_desagio") = perc_limite_RA_sem_desagio

				rs("endereco_memorizado_status") = 1
				rs("endereco_logradouro") = r_cliente.endereco
				rs("endereco_bairro") = r_cliente.bairro
				rs("endereco_cidade") = r_cliente.cidade
				rs("endereco_uf") = r_cliente.uf
				rs("endereco_cep") = r_cliente.cep
				rs("endereco_numero") = r_cliente.endereco_numero
				rs("endereco_complemento") = r_cliente.endereco_complemento

				else
				'	PEDIDO FILHOTE
				'	==============
					rs("st_auto_split") = 1
					rs("split_status") = 1
					rs("split_data") = Date
					rs("split_hora") = s_hora_pedido
					rs("split_usuario") = ID_USUARIO_SISTEMA
					rs("st_pagto")=""
					rs("usuario_st_pagto")=""
					rs("st_recebido")=""
					rs("obs_1")=""
					rs("obs_2")=""
					rs("qtde_parcelas")=0
					rs("forma_pagto")=""
					end if

				rs("st_orc_virou_pedido")=1
				rs("orcamento")=orcamento_selecionado
				rs("orcamentista")=r_orcamento.orcamentista

				rs("id_cliente")=r_orcamento.id_cliente
				rs("midia")=r_orcamento.midia
				rs("servicos")=""
				rs("vendedor")=r_orcamento.vendedor
				rs("usuario_cadastro")=usuario
				rs("st_entrega")=""
				if Trim("" & r_orcamento.st_etg_imediata) <> Trim(s_etg_imediata) then
					rs("st_etg_imediata")=CLng(s_etg_imediata)
					rs("etg_imediata_data")=Now
					rs("etg_imediata_usuario")=usuario
				else
					rs("st_etg_imediata")=r_orcamento.st_etg_imediata
					rs("etg_imediata_data")=r_orcamento.etg_imediata_data
					rs("etg_imediata_usuario")=r_orcamento.etg_imediata_usuario
					end if

				if Trim("" & r_orcamento.StBemUsoConsumo) <> Trim(s_bem_uso_consumo) then
					rs("StBemUsoConsumo")=CLng(s_bem_uso_consumo)
				else
					rs("StBemUsoConsumo")=r_orcamento.StBemUsoConsumo
					end if

				if s_instalador_instala <> "" then
					rs("InstaladorInstalaStatus")=CLng(s_instalador_instala)
					rs("InstaladorInstalaUsuarioUltAtualiz")=usuario
					rs("InstaladorInstalaDtHrUltAtualiz")=Now
					end if

				rs("NFe_texto_constar")=s_nf_texto
				rs("NFe_xPed")=s_num_pedido_compra

				rs("indicador") = r_orcamento.orcamentista

				rs("GarantiaIndicadorStatus") = CLng(rb_garantia_indicador)
				rs("GarantiaIndicadorUsuarioUltAtualiz") = usuario
				rs("GarantiaIndicadorDtHrUltAtualiz") = Now

				rs("st_end_entrega") = r_orcamento.st_end_entrega
				if CLng(r_orcamento.st_end_entrega) <> 0 then
					rs("EndEtg_endereco") = r_orcamento.EndEtg_endereco
					rs("EndEtg_endereco_numero") = r_orcamento.EndEtg_endereco_numero
					rs("EndEtg_endereco_complemento") = r_orcamento.EndEtg_endereco_complemento
					rs("EndEtg_bairro") = r_orcamento.EndEtg_bairro
					rs("EndEtg_cidade") = r_orcamento.EndEtg_cidade
					rs("EndEtg_uf") = r_orcamento.EndEtg_uf
					rs("EndEtg_cep") = r_orcamento.EndEtg_cep
					rs("EndEtg_cod_justificativa") = r_orcamento.EndEtg_cod_justificativa
					end if

				'OBTENÇÃO DE TRANSPORTADORA QUE ATENDA AO CEP INFORMADO, SE HOUVER
				if sTranspSelAutoTransportadoraId <> "" then
					rs("transportadora_id") = sTranspSelAutoTransportadoraId
					rs("transportadora_data") = Now
					rs("transportadora_usuario") = usuario
					rs("transportadora_selecao_auto_status") = iTranspSelAutoStatus
					rs("transportadora_selecao_auto_cep") = sTranspSelAutoCep
					rs("transportadora_selecao_auto_transportadora") = sTranspSelAutoTransportadoraId
					rs("transportadora_selecao_auto_tipo_endereco") = iTranspSelAutoTipoEndereco
					rs("transportadora_selecao_auto_data_hora") = Now
					end if

				'01/02/2018: os pedidos do Arclube usam o RA para incluir o valor do frete e, portanto, não devem ter deságio do RA
				if Cstr(loja) <> Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then rs("perc_desagio_RA_liquida") = r_orcamento.perc_desagio_RA_liquida

				rs("permite_RA_status") = r_orcamento.permite_RA_status

				if (r_orcamento.permite_RA_status = 1) then
					if blnTemRA then
						rs("opcao_possui_RA") = "S"
					else
						rs("opcao_possui_RA") = "N"
						end if
				else
					rs("opcao_possui_RA") = "-" ' Não se aplica
					end if

				rs("st_violado_permite_RA_status") = r_orcamento.st_violado_permite_RA_status
				rs("dt_hr_violado_permite_RA_status") = r_orcamento.dt_hr_violado_permite_RA_status
				rs("usuario_violado_permite_RA_status") = r_orcamento.usuario_violado_permite_RA_status

				rs("plataforma_origem_pedido") = COD_PLATAFORMA_ORIGEM_PEDIDO__ERP

				rs("id_nfe_emitente") = vEmpresaAutoSplit(iv)

				rs.Update
				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if

				if rs.State <> 0 then rs.Close

				sequencia_item = 0
				total_estoque_vendido=0
				total_estoque_sem_presenca=0
				s_log_item_autosplit = ""
				for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
					if Trim(vProdRegra(iRegra).produto) <> "" then
						for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
							if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = vEmpresaAutoSplit(iv)) And (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada > 0) then
							'	LOCALIZA O PRODUTO EM V_ITEM
								indice_item = -1
								for j=LBound(v_item) to UBound(v_item)
									if (Trim("" & v_item(j).fabricante) = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.fabricante) And _
										(Trim("" & v_item(j).produto) = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.produto) then
										indice_item = j
										exit for
										end if
									next

								if indice_item > -1 then
									sequencia_item = sequencia_item + 1
									with v_item(indice_item)
										s="SELECT * FROM t_PEDIDO_ITEM WHERE pedido='X'"
										rs.Open s, cn
										rs.AddNew
										rs("pedido") = id_pedido_temp
										rs("fabricante") = .fabricante
										rs("produto") = .produto
										rs("qtde") = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada
										rs("desc_dado") = .desc_dado
										rs("preco_venda") = .preco_venda
										rs("preco_NF") = .preco_NF
										rs("preco_fabricante") = .preco_fabricante
										rs("vl_custo2") = .vl_custo2
										rs("preco_lista") = .preco_lista
										rs("margem") = .margem
										rs("desc_max") = .desc_max
										rs("comissao") = .comissao
										rs("descricao") = .descricao
										rs("descricao_html") = .descricao_html
										rs("ean") = .ean
										rs("grupo") = .grupo
										rs("peso") = .peso
										rs("qtde_volumes") = .qtde_volumes
										rs("abaixo_min_status") = .abaixo_min_status
										rs("abaixo_min_autorizacao") = .abaixo_min_autorizacao
										rs("abaixo_min_autorizador") = .abaixo_min_autorizador
										rs("abaixo_min_superv_autorizador") = .abaixo_min_superv_autorizador
										rs("sequencia") = sequencia_item
										rs("markup_fabricante") = .markup_fabricante
										rs("custoFinancFornecCoeficiente") = .custoFinancFornecCoeficiente
										rs("custoFinancFornecPrecoListaBase") = .custoFinancFornecPrecoListaBase
										rs("cubagem") = .cubagem
										rs("ncm") = .ncm
										rs("cst") = .cst
										rs("descontinuado") = .descontinuado
										rs.Update
										if Err <> 0 then
										'	~~~~~~~~~~~~~~~~
											cn.RollbackTrans
										'	~~~~~~~~~~~~~~~~
											Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
											end if
										if rs.State <> 0 then rs.Close

										if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada > vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque then
											qtde_spe = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada - vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque
										else
											qtde_spe = 0
											end if

										if Not ESTOQUE_produto_saida_v2(usuario, id_pedido_temp, vEmpresaAutoSplit(iv), .fabricante, .produto, vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada, qtde_spe, qtde_estoque_vendido_aux, qtde_estoque_sem_presenca_aux, msg_erro) then
										'	~~~~~~~~~~~~~~~~
											cn.RollbackTrans
										'	~~~~~~~~~~~~~~~~
											Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
											end if

										.qtde_estoque_vendido = .qtde_estoque_vendido + qtde_estoque_vendido_aux
										.qtde_estoque_sem_presenca = .qtde_estoque_sem_presenca + qtde_estoque_sem_presenca_aux

										total_estoque_vendido = total_estoque_vendido + qtde_estoque_vendido_aux
										total_estoque_sem_presenca = total_estoque_sem_presenca + qtde_estoque_sem_presenca_aux

									'	LOG
										if s_log_item_autosplit <> "" then s_log_item_autosplit = s_log_item_autosplit & chr(13)
										s_log_item_autosplit = s_log_item_autosplit & "(" & .fabricante & ")" & .produto & ":" & _
													" Qtde Solicitada = " & vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada & "," & _
													" Qtde Sem Presença Autorizada = " & Cstr(qtde_spe) & "," & _
													" Qtde Estoque Vendido = " & Cstr(qtde_estoque_vendido_aux) & "," & _
													" Qtde Sem Presença = " & Cstr(qtde_estoque_sem_presenca_aux)
										end with
									end if 'if indice_item > -1
								end if 'if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = vEmpresaAutoSplit(iv)) And (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada > 0)
							next 'for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
						end if 'if Trim(vProdRegra(iRegra).produto) <> ""
					next 'for iRegra=LBound(vProdRegra) to UBound(vProdRegra)

				if indice_pedido = 1 then
					if Not gera_num_pedido(id_pedido_base, msg_erro) then
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
						end if
					id_pedido = id_pedido_base
				else
					id_pedido = id_pedido_base & COD_SEPARADOR_FILHOTE & gera_letra_pedido_filhote(indice_pedido-1)
					end if

			'	LOG
				if Trim("" & vLogAutoSplit(UBound(vLogAutoSplit))) <> "" then redim preserve vLogAutoSplit(UBound(vLogAutoSplit)+1)
				vLogAutoSplit(UBound(vLogAutoSplit)) = id_pedido & " (" & obtem_apelido_empresa_NFe_emitente(vEmpresaAutoSplit(iv)) & ")" & chr(13) & _
														s_log_item_autosplit

				s="UPDATE t_PEDIDO SET pedido='" & id_pedido & "' WHERE pedido='" & id_pedido_temp & "'"
				cn.Execute(s)

				s="UPDATE t_PEDIDO_ITEM SET pedido='" & id_pedido & "' WHERE pedido='" & id_pedido_temp & "'"
				cn.Execute(s)

				s="UPDATE t_ESTOQUE_MOVIMENTO SET pedido='" & id_pedido & "' WHERE pedido='" & id_pedido_temp & "'"
				cn.Execute(s)

				s="UPDATE t_ORCAMENTO SET st_orc_virou_pedido=1, pedido='" & id_pedido & "' WHERE (orcamento='" & orcamento_selecionado & "') AND (st_orc_virou_pedido=0)"
				cn.Execute(s)

				s="UPDATE t_ESTOQUE_LOG SET pedido_estoque_origem='" & id_pedido & "' WHERE pedido_estoque_origem='" & id_pedido_temp & "'"
				cn.Execute(s)

				s="UPDATE t_ESTOQUE_LOG SET pedido_estoque_destino='" & id_pedido & "' WHERE pedido_estoque_destino='" & id_pedido_temp & "'"
				cn.Execute(s)

				if indice_pedido = 1 then
				'	INDICADOR: SE ESTE PEDIDO É COM INDICADOR E O CLIENTE AINDA NÃO TEM UM INDICADOR NO CADASTRO, ENTÃO CADASTRA ESTE.
					if Trim(r_orcamento.orcamentista) <> "" then
						if Trim(r_cliente.indicador) = "" then
							s="UPDATE t_CLIENTE SET indicador='" & Trim(r_orcamento.orcamentista) & "' WHERE (id='" & r_orcamento.id_cliente & "')"
							cn.Execute(s)
							s_log_cliente_indicador = "Cadastrado o indicador '" & Trim(r_orcamento.orcamentista) & "' no cliente id=" & r_orcamento.id_cliente
							end if
						end if
					end if

		'		STATUS DE ENTREGA
				if total_estoque_vendido = 0 then
					s = ST_ENTREGA_ESPERAR
				elseif total_estoque_sem_presenca = 0 then
					s = ST_ENTREGA_SEPARAR
				else
					s = ST_ENTREGA_SPLIT_POSSIVEL
					end if
		
				s = "UPDATE t_PEDIDO SET st_entrega='" & s & "' WHERE pedido='" & id_pedido & "'"
				cn.Execute(s)

				if Not calcula_total_RA_liquido_BD(id_pedido, vl_total_RA_liquido, msg_erro) then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if

				if indice_pedido = 1 then
					s = "SELECT * FROM t_PEDIDO WHERE (pedido='" & id_pedido & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta = "Falha ao consultar o registro do novo pedido (" & id_pedido & ")"
					else
						rs("vl_total_RA_liquido") = vl_total_RA_liquido
						rs("qtde_parcelas_desagio_RA") = 0
						if vl_total_RA <> 0 then
							rs("st_tem_desagio_RA") = 1
						else
							rs("st_tem_desagio_RA") = 0
							end if
						rs.Update
						end if
					end if

				if indice_pedido = 1 then
			'		SENHAS DE AUTORIZAÇÃO PARA DESCONTO SUPERIOR
					for k = Lbound(v_desconto) to Ubound(v_desconto)
						if Trim(v_desconto(k)) <> "" then
							s = "SELECT * FROM t_DESCONTO" & _
								" WHERE (usado_status=0)" & _
								" AND (cancelado_status=0)" & _
								" AND (id='" & Trim(v_desconto(k)) & "')"
							if rs.State <> 0 then rs.Close
							rs.open s, cn
							if rs.Eof then
								alerta = "Senha de autorização para desconto superior não encontrado."
								exit for
							else
								rs("usado_status") = 1
								rs("usado_data") = Now
								rs("vendedor") = r_orcamento.vendedor
								rs("usado_usuario") = usuario
								rs.Update
								if Err <> 0 then
								'	~~~~~~~~~~~~~~~~
									cn.RollbackTrans
								'	~~~~~~~~~~~~~~~~
									Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
									end if
								end if
							end if
						next
					end if

				if indice_pedido = 1 then
					if alerta = "" then
					'	VERIFICA SE O ENDEREÇO JÁ FOI USADO ANTERIORMENTE POR OUTRO CLIENTE (POSSÍVEL FRAUDE)
					'	ENDEREÇO DO CADASTRO
					'	====================
					'	1) VERIFICA SE O ENDEREÇO USADO É O DO PARCEIRO
						if r_orcamento.orcamentista <> "" then
							if isEnderecoIgual(r_cliente.endereco, r_cliente.endereco_numero, r_cliente.cep, r_orcamentista_e_indicador.endereco, r_orcamentista_e_indicador.endereco_numero, r_orcamentista_e_indicador.cep) then
								blnAnEnderecoCadClienteUsaEndParceiro = True
								blnAnalisarEndereco = True
					
								if Not fin_gera_nsu(T_PEDIDO_ANALISE_ENDERECO, intNsuPai, msg_erro) then
									alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
								else
									s = "SELECT * FROM t_PEDIDO_ANALISE_ENDERECO WHERE (id = -1)"
									if rs.State <> 0 then rs.Close
									rs.Open s, cn
									rs.AddNew
									rs("id") = intNsuPai
									rs("pedido") = id_pedido
									rs("id_cliente") = r_orcamento.id_cliente
									rs("tipo_endereco") = COD_PEDIDO_AN_ENDERECO__CAD_CLIENTE
									rs("endereco_logradouro") = r_cliente.endereco
									rs("endereco_bairro") = r_cliente.bairro
									rs("endereco_cidade") = r_cliente.cidade
									rs("endereco_uf") = r_cliente.uf
									rs("endereco_cep") = r_cliente.cep
									rs("endereco_numero") = r_cliente.endereco_numero
									rs("endereco_complemento") = r_cliente.endereco_complemento
									rs("usuario_cadastro") = usuario
									rs.Update
									end if ' if Not fin_gera_nsu()
					
								if alerta = "" then
									if Not fin_gera_nsu(T_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO, intNsu, msg_erro) then
										alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
									else
										s = "SELECT * FROM t_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO WHERE (id = -1)"
										if rs.State <> 0 then rs.Close
										rs.Open s, cn
										rs.AddNew
										with r_orcamentista_e_indicador
											rs("id") = intNsu
											rs("id_pedido_analise_endereco") = intNsuPai
											rs("pedido") = ""
											rs("id_cliente") = ""
											rs("tipo_endereco") = COD_PEDIDO_AN_ENDERECO__END_PARCEIRO
											rs("endereco_logradouro") = .endereco
											rs("endereco_bairro") = .bairro
											rs("endereco_cidade") = .cidade
											rs("endereco_uf") = .uf
											rs("endereco_cep") = .cep
											rs("endereco_numero") = .endereco_numero
											rs("endereco_complemento") = .endereco_complemento
											end with
										rs.Update
										end if ' if Not fin_gera_nsu()
									end if ' if alerta = ""
								end if ' if isEnderecoIgual()
							end if ' if r_orcamento.orcamentista <> ""
						end if 'if alerta = ""
			
					if alerta = "" then
					'	2)VERIFICA PEDIDOS DE OUTROS CLIENTES
						if Not blnAnEnderecoCadClienteUsaEndParceiro then
							redim vAnEndConfrontacao(0)
							set vAnEndConfrontacao(Ubound(vAnEndConfrontacao)) = new cl_ANALISE_ENDERECO_CONFRONTACAO
							intQtdeTotalPedidosAnEndereco = 0
				
							s = "SELECT DISTINCT * FROM " & _
									"(" & _
										"SELECT" & _
											" '" & COD_PEDIDO_AN_ENDERECO__CAD_CLIENTE & "' AS tipo_endereco," & _
											" p.pedido," & _
											" p.data_hora," & _
											" p.id_cliente," & _
											" c.endereco AS endereco_logradouro," & _
											" c.endereco_numero," & _
											" c.endereco_complemento," & _
											" c.bairro AS endereco_bairro," & _
											" c.cidade AS endereco_cidade," & _
											" c.uf AS endereco_uf," & _
											" c.cep AS endereco_cep" & _
										" FROM t_PEDIDO p" & _
											" INNER JOIN t_CLIENTE c ON (p.id_cliente = c.id)" & _
										" WHERE" & _
											" (endereco_memorizado_status = 0)" & _
											" AND (c.id <> '" & r_orcamento.id_cliente & "')" & _
											" AND (c.cep = '" & retorna_so_digitos(r_cliente.cep) & "')" & _
										" UNION " & _
										"SELECT" & _
											" '" & COD_PEDIDO_AN_ENDERECO__CAD_CLIENTE_MEMORIZADO & "' AS tipo_endereco," & _
											" pedido," & _
											" data_hora," & _
											" id_cliente," & _
											" endereco_logradouro," & _
											" endereco_numero," & _
											" endereco_complemento," & _
											" endereco_bairro," & _
											" endereco_cidade," & _
											" endereco_uf," & _
											" endereco_cep" & _
										" FROM t_PEDIDO" & _
										" WHERE" & _
											" (endereco_memorizado_status = 1)" & _
											" AND (id_cliente <> '" & r_orcamento.id_cliente & "')" & _
											" AND (endereco_cep = '" & retorna_so_digitos(r_cliente.cep) & "')" & _
										" UNION " & _
										"SELECT" & _
											" '" & COD_PEDIDO_AN_ENDERECO__END_ENTREGA & "' AS tipo_endereco," & _
											" pedido," & _
											" data_hora," & _
											" id_cliente," & _
											" EndEtg_endereco AS endereco_logradouro," & _
											" EndEtg_endereco_numero AS endereco_numero," & _
											" EndEtg_endereco_complemento AS endereco_complemento," & _
											" EndEtg_bairro AS endereco_bairro," & _
											" EndEtg_cidade AS endereco_cidade," & _
											" EndEtg_uf AS endereco_uf," & _
											" EndEtg_cep AS endereco_cep" & _
										" FROM t_PEDIDO" & _
										" WHERE" & _
											" (st_end_entrega = 1)" & _
											" AND (id_cliente <> '" & r_orcamento.id_cliente & "')" & _
											" AND (EndEtg_cep = '" & retorna_so_digitos(r_cliente.cep) & "')" & _
									") t" & _
								" ORDER BY" & _
									" data_hora DESC"
							if rs.State <> 0 then rs.Close
							rs.Open s, cn
							do while Not rs.Eof
								if isEnderecoIgual(r_cliente.endereco, r_cliente.endereco_numero, r_cliente.cep, Trim("" & rs("endereco_logradouro")), Trim("" & rs("endereco_numero")), Trim("" & rs("endereco_cep"))) then
									if Trim("" & vAnEndConfrontacao(Ubound(vAnEndConfrontacao)).pedido) <> "" then
										redim preserve vAnEndConfrontacao(UBound(vAnEndConfrontacao)+1)
										set vAnEndConfrontacao(UBound(vAnEndConfrontacao)) = new cl_ANALISE_ENDERECO_CONFRONTACAO
										end if
						
									with vAnEndConfrontacao(UBound(vAnEndConfrontacao))
										.pedido = Trim("" & rs("pedido"))
										.id_cliente = Trim("" & rs("id_cliente"))
										.tipo_endereco = Trim("" & rs("tipo_endereco"))
										.endereco_logradouro = Trim("" & rs("endereco_logradouro"))
										.endereco_bairro = Trim("" & rs("endereco_bairro"))
										.endereco_cidade = Trim("" & rs("endereco_cidade"))
										.endereco_uf = Trim("" & rs("endereco_uf"))
										.endereco_cep = Trim("" & rs("endereco_cep"))
										.endereco_numero = Trim("" & rs("endereco_numero"))
										.endereco_complemento = Trim("" & rs("endereco_complemento"))
										end with
						
									intQtdeTotalPedidosAnEndereco = intQtdeTotalPedidosAnEndereco + 1
									if intQtdeTotalPedidosAnEndereco >= MAX_AN_ENDERECO_QTDE_PEDIDOS_CADASTRAMENTO then exit do
									end if 'if isEnderecoIgual()
					
								rs.MoveNext
								loop
							if rs.State <> 0 then rs.Close
				
							blnGravouRegPai = False
							for i=LBound(vAnEndConfrontacao) to UBound(vAnEndConfrontacao)
								with vAnEndConfrontacao(i)
									if Trim("" & .pedido) <> "" then
										blnAnalisarEndereco = True
									'	JÁ GRAVOU O REGISTRO PAI?
										if Not blnGravouRegPai then
											blnGravouRegPai = True
											if Not fin_gera_nsu(T_PEDIDO_ANALISE_ENDERECO, intNsuPai, msg_erro) then
												alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
												exit for
												end if
								
											s = "SELECT * FROM t_PEDIDO_ANALISE_ENDERECO WHERE (id = -1)"
											if rs.State <> 0 then rs.Close
											rs.Open s, cn
											rs.AddNew
											rs("id") = intNsuPai
											rs("pedido") = id_pedido
											rs("id_cliente") = r_orcamento.id_cliente
											rs("tipo_endereco") = COD_PEDIDO_AN_ENDERECO__CAD_CLIENTE
											rs("endereco_logradouro") = r_cliente.endereco
											rs("endereco_bairro") = r_cliente.bairro
											rs("endereco_cidade") = r_cliente.cidade
											rs("endereco_uf") = r_cliente.uf
											rs("endereco_cep") = r_cliente.cep
											rs("endereco_numero") = r_cliente.endereco_numero
											rs("endereco_complemento") = r_cliente.endereco_complemento
											rs("usuario_cadastro") = usuario
											rs.Update
											end if 'if Not blnGravouRegPai
							
										if Not fin_gera_nsu(T_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO, intNsu, msg_erro) then
											alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
											exit for
											end if
							
										s = "SELECT * FROM t_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO WHERE (id = -1)"
										if rs.State <> 0 then rs.Close
										rs.Open s, cn
										rs.AddNew
										rs("id") = intNsu
										rs("id_pedido_analise_endereco") = intNsuPai
										rs("pedido") = .pedido
										rs("id_cliente") = .id_cliente
										rs("tipo_endereco") = .tipo_endereco
										rs("endereco_logradouro") = .endereco_logradouro
										rs("endereco_bairro") = .endereco_bairro
										rs("endereco_cidade") = .endereco_cidade
										rs("endereco_uf") = .endereco_uf
										rs("endereco_cep") = .endereco_cep
										rs("endereco_numero") = .endereco_numero
										rs("endereco_complemento") = .endereco_complemento
										rs.Update
										end if 'if Trim("" & .pedido) <> ""
									end with
								next
							end if ' if Not blnAnEnderecoCadClienteUsaEndParceiro
						end if 'if alerta = ""
		
					if alerta = "" then
						if CLng(r_orcamento.st_end_entrega) <> 0 then
						'	ENDEREÇO DE ENTREGA (SE HOUVER)
						'	===============================
						'	1) VERIFICA SE O ENDEREÇO USADO É O DO PARCEIRO
							if r_orcamento.orcamentista <> "" then
								if isEnderecoIgual(r_orcamento.EndEtg_endereco, r_orcamento.EndEtg_endereco_numero, r_orcamento.EndEtg_cep, r_orcamentista_e_indicador.endereco, r_orcamentista_e_indicador.endereco_numero, r_orcamentista_e_indicador.cep) then
									blnAnEnderecoEndEntregaUsaEndParceiro = True
									blnAnalisarEndereco = True
						
									if Not fin_gera_nsu(T_PEDIDO_ANALISE_ENDERECO, intNsuPai, msg_erro) then
										alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
									else
										s = "SELECT * FROM t_PEDIDO_ANALISE_ENDERECO WHERE (id = -1)"
										if rs.State <> 0 then rs.Close
										rs.Open s, cn
										rs.AddNew
										rs("id") = intNsuPai
										rs("pedido") = id_pedido
										rs("id_cliente") = r_orcamento.id_cliente
										rs("tipo_endereco") = COD_PEDIDO_AN_ENDERECO__END_ENTREGA
										rs("endereco_logradouro") = r_orcamento.EndEtg_endereco
										rs("endereco_bairro") = r_orcamento.EndEtg_bairro
										rs("endereco_cidade") = r_orcamento.EndEtg_cidade
										rs("endereco_uf") = r_orcamento.EndEtg_uf
										rs("endereco_cep") = r_orcamento.EndEtg_cep
										rs("endereco_numero") = r_orcamento.EndEtg_endereco_numero
										rs("endereco_complemento") = r_orcamento.EndEtg_endereco_complemento
										rs("usuario_cadastro") = usuario
										rs.Update
										end if ' if Not fin_gera_nsu()
						
									if alerta = "" then
										if Not fin_gera_nsu(T_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO, intNsu, msg_erro) then
											alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
										else
											s = "SELECT * FROM t_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO WHERE (id = -1)"
											if rs.State <> 0 then rs.Close
											rs.Open s, cn
											rs.AddNew
											with r_orcamentista_e_indicador
												rs("id") = intNsu
												rs("id_pedido_analise_endereco") = intNsuPai
												rs("pedido") = ""
												rs("id_cliente") = ""
												rs("tipo_endereco") = COD_PEDIDO_AN_ENDERECO__END_PARCEIRO
												rs("endereco_logradouro") = .endereco
												rs("endereco_bairro") = .bairro
												rs("endereco_cidade") = .cidade
												rs("endereco_uf") = .uf
												rs("endereco_cep") = .cep
												rs("endereco_numero") = .endereco_numero
												rs("endereco_complemento") = .endereco_complemento
												end with
											rs.Update
											end if ' if Not fin_gera_nsu()
										end if ' if alerta = ""
									end if ' if isEnderecoIgual()
								end if ' if r_orcamento.orcamentista <> ""
				
						'	2)VERIFICA PEDIDOS DE OUTROS CLIENTES
							if alerta = "" then
								if Not blnAnEnderecoEndEntregaUsaEndParceiro then
									redim vAnEndConfrontacao(0)
									set vAnEndConfrontacao(Ubound(vAnEndConfrontacao)) = new cl_ANALISE_ENDERECO_CONFRONTACAO
									intQtdeTotalPedidosAnEndereco = 0
						
									s = "SELECT DISTINCT * FROM " & _
											"(" & _
												"SELECT" & _
													" '" & COD_PEDIDO_AN_ENDERECO__CAD_CLIENTE & "' AS tipo_endereco," & _
													" p.pedido," & _
													" p.data_hora," & _
													" p.id_cliente," & _
													" c.endereco AS endereco_logradouro," & _
													" c.endereco_numero," & _
													" c.endereco_complemento," & _
													" c.bairro AS endereco_bairro," & _
													" c.cidade AS endereco_cidade," & _
													" c.uf AS endereco_uf," & _
													" c.cep AS endereco_cep" & _
												" FROM t_PEDIDO p" & _
													" INNER JOIN t_CLIENTE c ON (p.id_cliente = c.id)" & _
												" WHERE" & _
													" (endereco_memorizado_status = 0)" & _
													" AND (c.id <> '" & r_orcamento.id_cliente & "')" & _
													" AND (c.cep = '" & retorna_so_digitos(r_orcamento.EndEtg_cep) & "')" & _
												" UNION " & _
												"SELECT" & _
													" '" & COD_PEDIDO_AN_ENDERECO__CAD_CLIENTE_MEMORIZADO & "' AS tipo_endereco," & _
													" pedido," & _
													" data_hora," & _
													" id_cliente," & _
													" endereco_logradouro," & _
													" endereco_numero," & _
													" endereco_complemento," & _
													" endereco_bairro," & _
													" endereco_cidade," & _
													" endereco_uf," & _
													" endereco_cep" & _
												" FROM t_PEDIDO" & _
												" WHERE" & _
													" (endereco_memorizado_status = 1)" & _
													" AND (id_cliente <> '" & r_orcamento.id_cliente & "')" & _
													" AND (endereco_cep = '" & retorna_so_digitos(r_orcamento.EndEtg_cep) & "')" & _
												" UNION " & _
												"SELECT" & _
													" '" & COD_PEDIDO_AN_ENDERECO__END_ENTREGA & "' AS tipo_endereco," & _
													" pedido," & _
													" data_hora," & _
													" id_cliente," & _
													" EndEtg_endereco AS endereco_logradouro," & _
													" EndEtg_endereco_numero AS endereco_numero," & _
													" EndEtg_endereco_complemento AS endereco_complemento," & _
													" EndEtg_bairro AS endereco_bairro," & _
													" EndEtg_cidade AS endereco_cidade," & _
													" EndEtg_uf AS endereco_uf," & _
													" EndEtg_cep AS endereco_cep" & _
												" FROM t_PEDIDO" & _
												" WHERE" & _
													" (st_end_entrega = 1)" & _
													" AND (id_cliente <> '" & r_orcamento.id_cliente & "')" & _
													" AND (EndEtg_cep = '" & retorna_so_digitos(r_orcamento.EndEtg_cep) & "')" & _
											") t" & _
										" ORDER BY" & _
											" data_hora DESC"
									if rs.State <> 0 then rs.Close
									rs.Open s, cn
									do while Not rs.Eof
										if isEnderecoIgual(r_orcamento.EndEtg_endereco, r_orcamento.EndEtg_endereco_numero, r_orcamento.EndEtg_cep, Trim("" & rs("endereco_logradouro")), Trim("" & rs("endereco_numero")), Trim("" & rs("endereco_cep"))) then
											if Trim("" & vAnEndConfrontacao(Ubound(vAnEndConfrontacao)).pedido) <> "" then
												redim preserve vAnEndConfrontacao(UBound(vAnEndConfrontacao)+1)
												set vAnEndConfrontacao(UBound(vAnEndConfrontacao)) = new cl_ANALISE_ENDERECO_CONFRONTACAO
												end if
								
											with vAnEndConfrontacao(UBound(vAnEndConfrontacao))
												.pedido = Trim("" & rs("pedido"))
												.id_cliente = Trim("" & rs("id_cliente"))
												.tipo_endereco = Trim("" & rs("tipo_endereco"))
												.endereco_logradouro = Trim("" & rs("endereco_logradouro"))
												.endereco_bairro = Trim("" & rs("endereco_bairro"))
												.endereco_cidade = Trim("" & rs("endereco_cidade"))
												.endereco_uf = Trim("" & rs("endereco_uf"))
												.endereco_cep = Trim("" & rs("endereco_cep"))
												.endereco_numero = Trim("" & rs("endereco_numero"))
												.endereco_complemento = Trim("" & rs("endereco_complemento"))
												end with
								
											intQtdeTotalPedidosAnEndereco = intQtdeTotalPedidosAnEndereco + 1
											if intQtdeTotalPedidosAnEndereco >= MAX_AN_ENDERECO_QTDE_PEDIDOS_CADASTRAMENTO then exit do
											end if 'if isEnderecoIgual()
							
										rs.MoveNext
										loop
									if rs.State <> 0 then rs.Close
						
									blnGravouRegPai = False
									for i=LBound(vAnEndConfrontacao) to UBound(vAnEndConfrontacao)
										with vAnEndConfrontacao(i)
											if Trim("" & .pedido) <> "" then
												blnAnalisarEndereco = True
											'	JÁ GRAVOU O REGISTRO PAI?
												if Not blnGravouRegPai then
													blnGravouRegPai = True
													if Not fin_gera_nsu(T_PEDIDO_ANALISE_ENDERECO, intNsuPai, msg_erro) then
														alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
														exit for
														end if
										
													s = "SELECT * FROM t_PEDIDO_ANALISE_ENDERECO WHERE (id = -1)"
													if rs.State <> 0 then rs.Close
													rs.Open s, cn
													rs.AddNew
													rs("id") = intNsuPai
													rs("pedido") = id_pedido
													rs("id_cliente") = r_orcamento.id_cliente
													rs("tipo_endereco") = COD_PEDIDO_AN_ENDERECO__END_ENTREGA
													rs("endereco_logradouro") = r_orcamento.EndEtg_endereco
													rs("endereco_bairro") = r_orcamento.EndEtg_bairro
													rs("endereco_cidade") = r_orcamento.EndEtg_cidade
													rs("endereco_uf") = r_orcamento.EndEtg_uf
													rs("endereco_cep") = r_orcamento.EndEtg_cep
													rs("endereco_numero") = r_orcamento.EndEtg_endereco_numero
													rs("endereco_complemento") = r_orcamento.EndEtg_endereco_complemento
													rs("usuario_cadastro") = usuario
													rs.Update
													end if 'if Not blnGravouRegPai
							
												if Not fin_gera_nsu(T_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO, intNsu, msg_erro) then
													alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
													exit for
													end if
									
												s = "SELECT * FROM t_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO WHERE (id = -1)"
												if rs.State <> 0 then rs.Close
												rs.Open s, cn
												rs.AddNew
												rs("id") = intNsu
												rs("id_pedido_analise_endereco") = intNsuPai
												rs("pedido") = .pedido
												rs("id_cliente") = .id_cliente
												rs("tipo_endereco") = .tipo_endereco
												rs("endereco_logradouro") = .endereco_logradouro
												rs("endereco_bairro") = .endereco_bairro
												rs("endereco_cidade") = .endereco_cidade
												rs("endereco_uf") = .endereco_uf
												rs("endereco_cep") = .endereco_cep
												rs("endereco_numero") = .endereco_numero
												rs("endereco_complemento") = .endereco_complemento
												rs.Update
												end if 'if Trim("" & .pedido) <> ""
											end with
										next
									end if ' if Not blnAnEnderecoEndEntregaUsaEndParceiro
								end if ' if alerta = ""
							end if 'if CLng(r_orcamento.st_end_entrega) <> 0
						end if 'if alerta = ""
		
					if alerta = "" then
						if blnAnalisarEndereco then
							s = "UPDATE t_PEDIDO SET analise_endereco_tratar_status = 1 WHERE (pedido = '" & id_pedido & "')"
							cn.Execute(s)
							end if
						end if
					end if 'if indice_pedido = 1
				end if ' if (vEmpresaAutoSplit(iv) <> 0) then
			
			if alerta <> "" then exit for
			next ' for iv = LBound(vEmpresaAutoSplit) to UBound(vEmpresaAutoSplit)

	'	LOG
		if alerta = "" then
			s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & id_pedido_base & "')"
			set rs = cn.execute(s)
			if Not rs.Eof then
				s_log = "Nº Pré-Pedido=" & orcamento_selecionado
				s_log = s_log & "; vl total=" & formata_moeda(vl_total)
				s_log = s_log & "; indicador=" & formata_texto_log(rs("indicador"))
				s_log = s_log & "; vl_total_NF=" & formata_moeda(rs("vl_total_NF"))
				s_log = s_log & "; vl_total_RA=" & formata_moeda(rs("vl_total_RA"))
				s_log = s_log & "; perc_RT=" & formata_texto_log(rs("perc_RT"))
				s_log = s_log & "; qtde_parcelas=" & formata_texto_log(rs("qtde_parcelas"))
				if Trim("" & rs("forma_pagto"))<>"" then s_log = s_log & "; forma_pagto=" & formata_texto_log(rs("forma_pagto"))
				if Trim("" & rs("servicos"))<>"" then s_log = s_log & "; servicos=" & formata_texto_log(rs("servicos")) 
				if (Trim("" & rs("vl_servicos"))<>"") And (Trim("" & rs("vl_servicos"))<>"0") then s_log = s_log & "; vl_servicos=" & formata_texto_log(rs("vl_servicos")) 
				if Trim("" & rs("st_recebido"))<>"" then s_log = s_log & "; st_recebido=" & formata_texto_log(rs("st_recebido")) 
				if Trim("" & rs("st_etg_imediata"))<> "" then s_log = s_log & "; st_etg_imediata=" & formata_texto_log(rs("st_etg_imediata")) 
				if Trim("" & rs("StBemUsoConsumo"))<> "" then s_log = s_log & "; StBemUsoConsumo=" & formata_texto_log(rs("StBemUsoConsumo")) 
				if Trim("" & rs("InstaladorInstalaStatus"))<> "" then s_log = s_log & "; InstaladorInstalaStatus=" & formata_texto_log(rs("InstaladorInstalaStatus")) 
				if Trim("" & rs("obs_1"))<>"" then s_log = s_log & "; obs_1=" & formata_texto_log(rs("obs_1")) 
				if Trim("" & rs("obs_2"))<>"" then s_log = s_log & "; obs_2=" & formata_texto_log(rs("obs_2"))
				if Cstr(rs("analise_credito"))=Cstr(COD_AN_CREDITO_OK) then
					s_log = s_log & "; análise crédito OK (<=" & formata_moeda(vl_aprov_auto_analise_credito) & ")"
				else
					s_log = s_log & "; status da análise crédito: " & Cstr(rs("analise_credito")) & " - " & descricao_analise_credito(Cstr(rs("analise_credito")))
					end if
			'	Forma de Pagamento (nova versão)
				s_log = s_log & "; tipo_parcelamento=" & formata_texto_log(rs("tipo_parcelamento"))
				if rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA then
					s_log = s_log & "; av_forma_pagto=" & formata_texto_log(rs("av_forma_pagto"))
				elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELA_UNICA then
					s_log = s_log & "; pu_forma_pagto=" & formata_texto_log(rs("pu_forma_pagto"))
					s_log = s_log & "; pu_valor=" & formata_texto_log(rs("pu_valor"))
					s_log = s_log & "; pu_vencto_apos=" & formata_texto_log(rs("pu_vencto_apos"))
				elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO then
					s_log = s_log & "; pc_qtde_parcelas=" & formata_texto_log(rs("pc_qtde_parcelas"))
					s_log = s_log & "; pc_valor_parcela=" & formata_texto_log(rs("pc_valor_parcela"))
				elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
					s_log = s_log & "; pc_maquineta_qtde_parcelas=" & formata_texto_log(rs("pc_maquineta_qtde_parcelas"))
					s_log = s_log & "; pc_maquineta_valor_parcela=" & formata_texto_log(rs("pc_maquineta_valor_parcela"))
				elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
					s_log = s_log & "; pce_forma_pagto_entrada=" & formata_texto_log(rs("pce_forma_pagto_entrada"))
					s_log = s_log & "; pce_forma_pagto_prestacao=" & formata_texto_log(rs("pce_forma_pagto_prestacao"))
					s_log = s_log & "; pce_entrada_valor=" & formata_texto_log(rs("pce_entrada_valor"))
					s_log = s_log & "; pce_prestacao_qtde=" & formata_texto_log(rs("pce_prestacao_qtde"))
					s_log = s_log & "; pce_prestacao_valor=" & formata_texto_log(rs("pce_prestacao_valor"))
					s_log = s_log & "; pce_prestacao_periodo=" & formata_texto_log(rs("pce_prestacao_periodo"))
				elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
					s_log = s_log & "; pse_forma_pagto_prim_prest=" & formata_texto_log(rs("pse_forma_pagto_prim_prest"))
					s_log = s_log & "; pse_forma_pagto_demais_prest=" & formata_texto_log(rs("pse_forma_pagto_demais_prest"))
					s_log = s_log & "; pse_prim_prest_valor=" & formata_texto_log(rs("pse_prim_prest_valor"))
					s_log = s_log & "; pse_prim_prest_apos=" & formata_texto_log(rs("pse_prim_prest_apos"))
					s_log = s_log & "; pse_demais_prest_qtde=" & formata_texto_log(rs("pse_demais_prest_qtde"))
					s_log = s_log & "; pse_demais_prest_valor=" & formata_texto_log(rs("pse_demais_prest_valor"))
					s_log = s_log & "; pse_demais_prest_periodo=" & formata_texto_log(rs("pse_demais_prest_periodo"))
					end if
		
				s_log = s_log & "; custoFinancFornecTipoParcelamento=" & formata_texto_log(rs("custoFinancFornecTipoParcelamento"))
				s_log = s_log & "; custoFinancFornecQtdeParcelas=" & formata_texto_log(rs("custoFinancFornecQtdeParcelas"))
	
				if CLng(r_orcamento.st_end_entrega) <> 0 then
					s_log = s_log & "; Endereço entrega=" & r_orcamento.EndEtg_endereco
					if Trim(r_orcamento.EndEtg_endereco_numero) <> "" then s_log = s_log & ", " & r_orcamento.EndEtg_endereco_numero
					if Trim(r_orcamento.EndEtg_endereco_complemento) <> "" then s_log = s_log & " " & r_orcamento.EndEtg_endereco_complemento
					s_log = s_log & " - " & r_orcamento.EndEtg_bairro & " - " & r_orcamento.EndEtg_cidade & " - " & r_orcamento.EndEtg_uf
					if r_orcamento.EndEtg_cep <> "" then s_log = s_log & " - " & cep_formata(r_orcamento.EndEtg_cep)
					if r_orcamento.EndEtg_obs <> "" then s_log = s_log & " - " & r_orcamento.EndEtg_obs
				else
					s_log = s_log & "; Endereço entrega=mesmo do cadastro"
					end if

				if sTranspSelAutoTransportadoraId = "" then
					s_log = s_log & "; Escolha automática de transportadora=N"
				else
					s_log = s_log & "; Escolha automática de transportadora=S"
					s_log = s_log & "; Transportadora=" & sTranspSelAutoTransportadoraId
					s_log = s_log & "; CEP relacionado=" & cep_formata(sTranspSelAutoCep)
					end if

				s_log = s_log & "; GarantiaIndicadorStatus=" & rb_garantia_indicador
				s_log = s_log & "; perc_desagio_RA_liquida=" & r_orcamento.perc_desagio_RA_liquida
				end if 'if Not rs.Eof

			if s_log_cliente_indicador <> "" then
				if s_log <> "" then s_log = s_log & "; "
				s_log = s_log & s_log_cliente_indicador
				end if

		'	MONTA LOG DOS ITENS
			for i=Lbound(v_item) to Ubound(v_item)
				with v_item(i)
					if s_log <> "" then s_log=s_log & ";" & chr(13)
					s_log = s_log & _
							log_produto_monta(.qtde, .fabricante, .produto) & _
							"; preco_lista=" & formata_texto_log(.preco_lista) & _
							"; desc_dado=" & formata_texto_log(.desc_dado) & _
							"; preco_venda=" & formata_texto_log(.preco_venda) & _
							"; preco_NF=" & formata_texto_log(.preco_NF) & _
							"; custoFinancFornecCoeficiente=" & formata_texto_log(.custoFinancFornecCoeficiente) & _
							"; custoFinancFornecPrecoListaBase=" & formata_texto_log(.custoFinancFornecPrecoListaBase)
					if .qtde_estoque_vendido<>0 then s_log = s_log & "; estoque_vendido=" & formata_texto_log(.qtde_estoque_vendido)
					if .qtde_estoque_sem_presenca<>0 then s_log = s_log & "; estoque_sem_presenca=" & formata_texto_log(.qtde_estoque_sem_presenca)
				
					if converte_numero(.abaixo_min_status) <> 0 then
						s_log = s_log & _
								"; abaixo_min_status=" & formata_texto_log(.abaixo_min_status) & _
								"; abaixo_min_autorizacao=" & formata_texto_log(.abaixo_min_autorizacao) & _
								"; abaixo_min_autorizador=" & formata_texto_log(.abaixo_min_autorizador) & _
								"; abaixo_min_superv_autorizador=" & formata_texto_log(.abaixo_min_superv_autorizador)
						end if
					end with
				next

		'	ADICIONA DETALHES SOBRE O AUTO-SPLIT
			blnAchou=False
			for i=LBound(vLogAutoSplit) to UBound(vLogAutoSplit)
				if Trim("" & vLogAutoSplit(i)) <> "" then
					if s_log <> "" then s_log = s_log & chr(13)
					if Not blnAchou then
						s_log = s_log & "Detalhes do auto-split: Modo de seleção do CD = " & rb_selecao_cd
						if rb_selecao_cd = MODO_SELECAO_CD__MANUAL then s_log = s_log & "; id_nfe_emitente = " & c_id_nfe_emitente_selecao_manual
						s_log = s_log & chr(13)
						blnAchou = True
						end if
					s_log = s_log & vLogAutoSplit(i)
					end if
				next

			if s_log <> "" then
				grava_log usuario, loja, id_pedido, r_orcamento.id_cliente, OP_LOG_ORCAMENTO_VIROU_PEDIDO, s_log
				end if
			end if 'if alerta = ""

		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("pedido.asp?pedido_selecionado=" & id_pedido_base & "&url_back=X" & "&AutoSplitQty=" & Cstr(indice_pedido) & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			end if
		end if

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



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>



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
	<% 	if erro_produto_indisponivel then 
			s="Resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
		else
			s="javascript:history.back()"
			end if	%>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>
<% end if %>

</html>


<%
	on error resume next
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>