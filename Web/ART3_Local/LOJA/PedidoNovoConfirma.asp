<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  P E D I D O N O V O C O N F I R M A . A S P
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
'
'
'	REVISADO P/ IE10


	On Error GoTo 0
	Err.Clear

	dim msg_erro
	dim usuario, loja, cliente_selecionado, midia_selecionada, indicador_original, tipo_cliente
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim alerta, alerta_aux
	alerta=""
	
	cliente_selecionado = Trim(request("cliente_selecionado"))
	if (cliente_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2, t_CLIENTE, tMAP_XML, tOI
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim operacao_origem, c_numero_magento, operationControlTicket, sessionToken, id_magento_api_pedido_xml
	operacao_origem = Trim(Request("operacao_origem"))
	c_numero_magento = ""
	operationControlTicket = ""
	sessionToken = ""
	id_magento_api_pedido_xml = ""
	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		c_numero_magento = Trim(Request("c_numero_magento"))
		operationControlTicket = Trim(Request("operationControlTicket"))
		sessionToken = Trim(Request("sessionToken"))
		id_magento_api_pedido_xml = Trim(Request("id_magento_api_pedido_xml"))
		end if

	dim rb_selecao_cd, c_id_nfe_emitente_selecao_manual
	rb_selecao_cd = Trim(Request("rb_selecao_cd"))
	c_id_nfe_emitente_selecao_manual = Trim(Request("c_id_nfe_emitente_selecao_manual"))

	dim blnAnalisarEndereco, blnGravouRegPai, intNsu, intNsuPai, vAnEndConfrontacao
	dim intQtdeTotalPedidosAnEndereco
	dim blnAnEnderecoCadClienteUsaEndParceiro, blnAnEnderecoEndEntregaUsaEndParceiro
	blnAnalisarEndereco = False
	blnAnEnderecoCadClienteUsaEndParceiro = False
	blnAnEnderecoEndEntregaUsaEndParceiro = False
	
	dim rb_indicacao, rb_RA, c_indicador, c_perc_RT, rb_garantia_indicador, c_ped_bonshop
	c_ped_bonshop = Trim(Request.Form("c_ped_bonshop"))
	rb_indicacao = Trim(Request.Form("rb_indicacao"))
	if rb_indicacao = "S" then
		c_indicador = Trim(Request.Form("c_indicador"))
		c_perc_RT = Trim(Request.Form("c_perc_RT"))
		rb_RA = Trim(Request.Form("rb_RA"))
		rb_garantia_indicador = Trim(Request.Form("rb_garantia_indicador"))
	else
		c_indicador = ""
		c_perc_RT = ""
		rb_RA = ""
		rb_garantia_indicador = COD_GARANTIA_INDICADOR_STATUS__NAO
		end if
	
	dim perc_RT
	perc_RT = converte_numero(c_perc_RT)

	if alerta = "" then
		if (perc_RT < 0) Or (perc_RT > 100) then
			alerta = "Percentual de comiss�o inv�lido."
			end if
		end if
	
	'TRATAMENTO PARA CADASTRAMENTO DE PEDIDOS DO SITE MAGENTO DA BONSHOP
	dim blnMagentoPedidoComIndicador, sListaLojaMagentoPedidoComIndicador, vLoja, rParametro
	blnMagentoPedidoComIndicador = False
	sListaLojaMagentoPedidoComIndicador = ""
	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		If Not cria_recordset_otimista(tMAP_XML, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

		set rParametro = get_registro_t_parametro(ID_PARAMETRO_MagentoPedidoComIndicadorListaLojaErp)
		sListaLojaMagentoPedidoComIndicador = Trim("" & rParametro.campo_texto)
		if sListaLojaMagentoPedidoComIndicador <> "" then
			vLoja = Split(sListaLojaMagentoPedidoComIndicador, ",")
			for i=LBound(vLoja) to UBound(vLoja)
				if Trim("" & vLoja(i)) = loja then
					blnMagentoPedidoComIndicador = True
					exit for
					end if
				next
			end if
		end if

'	FORMA DE PAGAMENTO (NOVA VERS�O)
	dim rb_forma_pagto, op_av_forma_pagto, c_pc_qtde, c_pc_valor, c_pc_maquineta_qtde, c_pc_maquineta_valor
	dim op_pu_forma_pagto, c_pu_valor, c_pu_vencto_apos
	dim op_pce_entrada_forma_pagto, c_pce_entrada_valor, op_pce_prestacao_forma_pagto, c_pce_prestacao_qtde, c_pce_prestacao_valor, c_pce_prestacao_periodo
	dim op_pse_prim_prest_forma_pagto, c_pse_prim_prest_valor, c_pse_prim_prest_apos, op_pse_demais_prest_forma_pagto, c_pse_demais_prest_qtde, c_pse_demais_prest_valor, c_pse_demais_prest_periodo
	dim vlTotalFormaPagto

	vlTotalFormaPagto = 0
	if alerta = "" then
		rb_forma_pagto = Trim(Request.Form("rb_forma_pagto"))
		if rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA then
			op_av_forma_pagto = Trim(Request.Form("op_av_forma_pagto"))
			if op_av_forma_pagto = "" then alerta = "Indique a forma de pagamento (� vista)."
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELA_UNICA then
			op_pu_forma_pagto = Trim(Request.Form("op_pu_forma_pagto"))
			c_pu_valor = Trim(Request.Form("c_pu_valor"))
			c_pu_vencto_apos = Trim(Request.Form("c_pu_vencto_apos"))
			if op_pu_forma_pagto = "" then
				alerta = "Indique a forma de pagamento da parcela �nica."
			elseif c_pu_valor = "" then
				alerta = "Indique o valor da parcela �nica."
			elseif converte_numero(c_pu_valor) <= 0 then
				alerta = "Valor da parcela �nica � inv�lido."
			elseif c_pu_vencto_apos = "" then
				alerta = "Indique o intervalo de vencimento da parcela �nica."
			elseif converte_numero(c_pu_vencto_apos) <= 0 then
				alerta = "Intervalo de vencimento da parcela �nica � inv�lido."
				end if
			if alerta = "" then vlTotalFormaPagto = converte_numero(c_pu_valor)
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO then
			c_pc_qtde = Trim(Request.Form("c_pc_qtde"))
			c_pc_valor = Trim(Request.Form("c_pc_valor"))
			if c_pc_qtde = "" then
				alerta = "Indique a quantidade de parcelas (parcelado no cart�o [internet])."
			elseif c_pc_valor = "" then
				alerta = "Indique o valor da parcela (parcelado no cart�o [internet])."
			elseif converte_numero(c_pc_qtde) < 1 then
				alerta = "Quantidade de parcelas inv�lida (parcelado no cart�o [internet])."
			elseif converte_numero(c_pc_valor) <= 0 then
				alerta = "Valor de parcela inv�lido (parcelado no cart�o [internet])."
				end if
			if alerta = "" then vlTotalFormaPagto = converte_numero(c_pc_qtde) * converte_numero(c_pc_valor)
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
			c_pc_maquineta_qtde = Trim(Request.Form("c_pc_maquineta_qtde"))
			c_pc_maquineta_valor = Trim(Request.Form("c_pc_maquineta_valor"))
			if c_pc_maquineta_qtde = "" then
				alerta = "Indique a quantidade de parcelas (parcelado no cart�o [maquineta])."
			elseif c_pc_maquineta_valor = "" then
				alerta = "Indique o valor da parcela (parcelado no cart�o [maquineta])."
			elseif converte_numero(c_pc_maquineta_qtde) < 1 then
				alerta = "Quantidade de parcelas inv�lida (parcelado no cart�o [maquineta])."
			elseif converte_numero(c_pc_maquineta_valor) <= 0 then
				alerta = "Valor de parcela inv�lido (parcelado no cart�o [maquineta])."
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
				alerta = "Valor da entrada inv�lido (parcelado com entrada)."
			elseif op_pce_prestacao_forma_pagto = "" then
				alerta = "Indique a forma de pagamento das presta��es (parcelado com entrada)."
			elseif c_pce_prestacao_qtde = "" then
				alerta = "Indique a quantidade de presta��es (parcelado com entrada)."
			elseif converte_numero(c_pce_prestacao_qtde) <= 0 then
				alerta = "Quantidade de presta��es inv�lida (parcelado com entrada)."
			elseif c_pce_prestacao_valor = "" then
				alerta = "Indique o valor da presta��o (parcelado com entrada)."
			elseif converte_numero(c_pce_prestacao_valor) <= 0 then
				alerta = "Valor de presta��o inv�lido (parcelado com entrada)."
			elseif c_pce_prestacao_periodo = "" then
				alerta = "Indique o intervalo de vencimento entre as parcelas (parcelado com entrada)."
			elseif converte_numero(c_pce_prestacao_periodo) <= 0 then
				alerta = "Intervalo de vencimento inv�lido (parcelado com entrada)."
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
				alerta = "Indique a forma de pagamento da 1� presta��o (parcelado sem entrada)."
			elseif c_pse_prim_prest_valor = "" then
				alerta = "Indique o valor da 1� presta��o (parcelado sem entrada)."
			elseif converte_numero(c_pse_prim_prest_valor) <= 0 then
				alerta = "Valor da 1� presta��o inv�lido (parcelado sem entrada)."
			elseif c_pse_prim_prest_apos = "" then
				alerta = "Indique o intervalo de vencimento da 1� parcela (parcelado sem entrada)."
			elseif converte_numero(c_pse_prim_prest_apos) <= 0 then
				alerta = "Intervalo de vencimento da 1� parcela � inv�lido (parcelado sem entrada)."
			elseif op_pse_demais_prest_forma_pagto = "" then
				alerta = "Indique a forma de pagamento das demais presta��es (parcelado sem entrada)."
			elseif c_pse_demais_prest_qtde = "" then
				alerta = "Indique a quantidade das demais presta��es (parcelado sem entrada)."
			elseif converte_numero(c_pse_demais_prest_qtde) <= 0 then
				alerta = "Quantidade de presta��es inv�lida (parcelado sem entrada)."
			elseif c_pse_demais_prest_valor = "" then
				alerta = "Indique o valor das demais presta��es (parcelado sem entrada)."
			elseif converte_numero(c_pse_demais_prest_valor) <= 0 then
				alerta = "Valor de presta��o inv�lido (parcelado sem entrada)."
			elseif c_pse_demais_prest_periodo = "" then
				alerta = "Indique o intervalo de vencimento entre as parcelas (parcelado sem entrada)."
			elseif converte_numero(c_pse_demais_prest_periodo) <= 0 then
				alerta = "Intervalo de vencimento inv�lido (parcelado sem entrada)."
				end if
			if alerta = "" then
				vlTotalFormaPagto = converte_numero(c_pse_prim_prest_valor) + (converte_numero(c_pse_demais_prest_qtde) * converte_numero(c_pse_demais_prest_valor))
				end if
		else
			alerta = "� obrigat�rio especificar a forma de pagamento"
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
			alerta="Foi detectada uma inconsist�ncia no tipo de parcelamento do pagamento (c�digo esperado=" & c_custoFinancFornecTipoParcelamentoConferencia & ", c�digo lido=" & c_custoFinancFornecTipoParcelamento & ")"
		elseif converte_numero(c_custoFinancFornecQtdeParcelasConferencia)<>converte_numero(c_custoFinancFornecQtdeParcelas) then
			alerta="Foi detectada uma inconsist�ncia na quantidade de parcelas de pagamento (qtde esperada=" & c_custoFinancFornecQtdeParcelasConferencia & ", qtde lida=" & c_custoFinancFornecQtdeParcelas & ")"
			end if
		end if

	dim rb_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento
	dim EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs
	dim EndEtg_email, EndEtg_email_xml, EndEtg_nome, EndEtg_ddd_res, EndEtg_tel_res, EndEtg_ddd_com, EndEtg_tel_com, EndEtg_ramal_com
	dim EndEtg_ddd_cel, EndEtg_tel_cel, EndEtg_ddd_com_2, EndEtg_tel_com_2, EndEtg_ramal_com_2
	dim EndEtg_tipo_pessoa, EndEtg_cnpj_cpf, EndEtg_contribuinte_icms_status, EndEtg_produtor_rural_status
	dim EndEtg_ie, EndEtg_rg
	rb_end_entrega = Trim(Request.Form("rb_end_entrega"))
	EndEtg_endereco = Trim(Request.Form("EndEtg_endereco"))
	EndEtg_endereco_numero = Trim(Request.Form("EndEtg_endereco_numero"))
	EndEtg_endereco_complemento = Trim(Request.Form("EndEtg_endereco_complemento"))
	EndEtg_bairro = Trim(Request.Form("EndEtg_bairro"))
	EndEtg_cidade = Trim(Request.Form("EndEtg_cidade"))
	EndEtg_uf = Trim(Request.Form("EndEtg_uf"))
	EndEtg_cep = retorna_so_digitos(Trim(Request.Form("EndEtg_cep")))
    EndEtg_obs = Trim(Request.Form("EndEtg_obs"))
	EndEtg_email = Trim(Request.Form("EndEtg_email"))
	EndEtg_email_xml = Trim(Request.Form("EndEtg_email_xml"))
	EndEtg_nome = Trim(Request.Form("EndEtg_nome"))
	EndEtg_ddd_res = Trim(Request.Form("EndEtg_ddd_res"))
	EndEtg_tel_res = Trim(Request.Form("EndEtg_tel_res"))
	EndEtg_ddd_com = Trim(Request.Form("EndEtg_ddd_com"))
	EndEtg_tel_com = Trim(Request.Form("EndEtg_tel_com"))
	EndEtg_ramal_com = Trim(Request.Form("EndEtg_ramal_com"))
	EndEtg_ddd_cel = Trim(Request.Form("EndEtg_ddd_cel"))
	EndEtg_tel_cel = Trim(Request.Form("EndEtg_tel_cel"))
	EndEtg_ddd_com_2 = Trim(Request.Form("EndEtg_ddd_com_2"))
	EndEtg_tel_com_2 = Trim(Request.Form("EndEtg_tel_com_2"))
	EndEtg_ramal_com_2 = Trim(Request.Form("EndEtg_ramal_com_2"))
	EndEtg_tipo_pessoa = Trim(Request.Form("EndEtg_tipo_pessoa"))
	EndEtg_cnpj_cpf = Trim(Request.Form("EndEtg_cnpj_cpf"))
	EndEtg_contribuinte_icms_status = Trim(Request.Form("EndEtg_contribuinte_icms_status"))
	EndEtg_produtor_rural_status = Trim(Request.Form("EndEtg_produtor_rural_status"))
	EndEtg_ie = Trim(Request.Form("EndEtg_ie"))
	EndEtg_rg = Trim(Request.Form("EndEtg_rg"))

	dim vl_aprov_auto_analise_credito
	vl_aprov_auto_analise_credito = 0

	dim vl_total_RA_liquido
	dim s, c, i, iv, j, k, n, opcao_venda_sem_estoque, vl_total, vl_total_NF, vl_total_RA, qtde_estoque_total_disponivel, blnAchou, blnDesativado
	dim v_desconto()
	ReDim v_desconto(0)
	v_desconto(UBound(v_desconto)) = ""

	opcao_venda_sem_estoque = Trim(request("opcao_venda_sem_estoque"))
	
	dim s_forma_pagto, s_obs1, s_obs2, s_recebido, s_etg_imediata, s_bem_uso_consumo, s_pedido_ac, s_numero_mktplace, s_origem_pedido
    dim s_nf_texto, s_num_pedido_compra
	s_obs1=Trim(request("c_obs1"))
	s_obs2=Trim(request("c_obs2"))
    s_pedido_ac=Trim(request("c_pedido_ac"))
    s_numero_mktplace = Trim(Request("c_numero_mktplace"))
    s_origem_pedido = Trim(Request("c_origem_pedido"))
	s_recebido=Trim(request("rb_recebido"))
	s_etg_imediata=Trim(request("rb_etg_imediata"))
	s_bem_uso_consumo=Trim(request("rb_bem_uso_consumo"))
	s_forma_pagto=Trim(request("c_forma_pagto"))
    s_nf_texto = Trim(request("c_nf_texto"))
    s_num_pedido_compra = Trim(request("c_num_pedido_compra"))

	dim c_exibir_campo_instalador_instala, s_instalador_instala
	c_exibir_campo_instalador_instala = Trim(Request.Form("c_exibir_campo_instalador_instala"))
	s_instalador_instala = Trim(Request.Form("rb_instalador_instala"))
	
	dim s_loja_indicou, comissao_loja_indicou, venda_externa
	s_loja_indicou=""
	comissao_loja_indicou=0
	venda_externa=0
	if Session("vendedor_externo") then
		s_loja_indicou=retorna_so_digitos(Trim(request("loja_indicou")))
		venda_externa=1
		end if
	
'	LOCALIZA DADOS DO CLIENTE
	midia_selecionada = ""
	indicador_original = ""
	tipo_cliente = ""
	if cliente_selecionado <> "" then
		s = "SELECT * FROM t_CLIENTE WHERE (id='" & cliente_selecionado & "')"
		set t_CLIENTE = cn.execute(s)
		if Not t_CLIENTE.Eof then
			midia_selecionada = Trim("" & t_CLIENTE("midia"))
			indicador_original = Trim("" & t_CLIENTE("indicador"))
			tipo_cliente = Trim("" & t_CLIENTE("tipo"))
			if Trim("" & t_CLIENTE("cep")) = "" then alerta = "� necess�rio preencher o CEP no cadastro do cliente."
			end if
		end if
	
	dim rCD
	set rCD = obtem_perc_max_comissao_e_desconto_por_loja(loja)

'	OBT�M A RELA��O DE MEIOS DE PAGAMENTO PREFERENCIAIS (QUE FAZEM USO O PERCENTUAL DE COMISS�O+DESCONTO N�VEL 2)
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
			alerta = "Percentual de comiss�o excede o m�ximo permitido."
			end if
		end if
	
	dim r_orcamentista_e_indicador
	dim permite_RA_status
	permite_RA_status = 0
	if alerta = "" then
		if c_indicador <> "" then
			if Not le_orcamentista_e_indicador(c_indicador, r_orcamentista_e_indicador, msg_erro) then
				alerta = "Falha ao recuperar os dados do indicador!!<br>" & msg_erro
			else
				permite_RA_status = r_orcamentista_e_indicador.permite_RA_status
				end if
			end if
		end if
	
	if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) And blnMagentoPedidoComIndicador then permite_RA_status = 1

'	RA L�quido
	dim perc_desagio_RA, perc_limite_RA_sem_desagio
	dim vl_limite_mensal, vl_limite_mensal_consumido, vl_limite_mensal_disponivel

	if rb_indicacao = "S" then
		perc_desagio_RA = obtem_perc_desagio_RA_do_indicador(c_indicador)
		perc_limite_RA_sem_desagio = obtem_perc_limite_RA_sem_desagio()
		vl_limite_mensal = obtem_limite_mensal_compras_do_indicador(c_indicador)
		vl_limite_mensal_consumido = calcula_limite_mensal_consumido_do_indicador(c_indicador, Date)
		vl_limite_mensal_disponivel = vl_limite_mensal - vl_limite_mensal_consumido
	else
		perc_desagio_RA = 0
		perc_limite_RA_sem_desagio = 0
		vl_limite_mensal = 0
		vl_limite_mensal_consumido = 0
		end if

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
				if (rb_RA = "S") And (permite_RA_status = 1) then
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
	
'	VERIFICA SE ESTE PEDIDO J� FOI GRAVADO!!
	dim pedido_a, vjg
	if Cstr(loja) <> NUMERO_LOJA_OLD03 then
		s = "SELECT t_PEDIDO.pedido, fabricante, produto, qtde, preco_venda FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			" WHERE (id_cliente='" & cliente_selecionado & "') AND (data=" & bd_formata_data(Date) & ")" & _
			" AND (loja='" & loja & "') AND (vendedor='" & usuario & "')" & _
			" AND (data >= " & bd_monta_data(Date) & ")" & _
			" AND (hora >= '" & formata_hora_hhnnss(Now-converte_min_to_dec(10))& "')" & _
			" AND (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')" & _
			" ORDER BY t_PEDIDO_ITEM.pedido, sequencia"
		set rs = cn.execute(s)
		redim vjg(0)
		set vjg(ubound(vjg)) = New cl_DUAS_COLUNAS
		vjg(ubound(vjg)).c1=""
		pedido_a="--XX--"
		do while Not rs.EOF 
			if pedido_a<>Trim("" & rs("pedido")) then
				pedido_a=Trim("" & rs("pedido"))
				if vjg(ubound(vjg)).c1 <> "" then 
					redim preserve vjg(ubound(vjg)+1)
					set vjg(ubound(vjg)) = New cl_DUAS_COLUNAS
					vjg(ubound(vjg)).c1=""
					end if
				vjg(ubound(vjg)).c2=pedido_a
				end if
		
			vjg(ubound(vjg)).c1=vjg(ubound(vjg)).c1 & Trim("" & rs("fabricante")) & "|" & Trim("" & rs("produto")) & "|" & Trim("" & rs("qtde")) & "|" & formata_moeda(rs("preco_venda")) & "|"
			rs.MoveNext 
			Loop

		if rs.State <> 0 then rs.Close
	
		s=""
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .produto<>"" then
					s=s & .fabricante & "|" & .produto & "|" & Cstr(.qtde) & "|" & formata_moeda(.preco_venda) & "|"
					end if
				end with
			next

		for i=Lbound(vjg) to Ubound(vjg)
			if s=vjg(i).c1 then
				alerta="Este pedido j� foi gravado com o n�mero " & vjg(i).c2
				exit for
				end if
			next
		end if
	
'	CUSTO FINANCEIRO FORNECEDOR
	if alerta = "" then
		if (c_custoFinancFornecTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA) And _
		   (c_custoFinancFornecTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) And _
		   (c_custoFinancFornecTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) then
			alerta = "A forma de pagamento n�o foi informada (� vista, com entrada, sem entrada)."
			end if
		end if
		
	if alerta = "" then
		if (c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) Or _
		   (c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) then
			if converte_numero(c_custoFinancFornecQtdeParcelas) <= 0 then
				alerta = "N�o foi informada a quantidade de parcelas para a forma de pagamento selecionada (" & descricaoCustoFinancFornecTipoParcelamento(c_custoFinancFornecTipoParcelamento) &  ")"
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
	
'	ANALISA O PERCENTUAL DE COMISS�O+DESCONTO
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
				'	O meio de pagamento selecionado � um dos preferenciais
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
				'	O meio de pagamento selecionado � um dos preferenciais
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
				'	O meio de pagamento selecionado � um dos preferenciais
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
				'	O meio de pagamento selecionado � um dos preferenciais
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
				'	O meio de pagamento selecionado � um dos preferenciais
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
				'	O meio de pagamento selecionado � um dos preferenciais
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
		
		'	O montante a pagar por meio de pagamento preferencial � maior que 50% do total?
			if vlNivel2 > (vl_total/2) then
				if tipo_cliente = ID_PJ then
					perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
				else
					perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
					end if
				end if
			
		elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
		'	Identifica e contabiliza o valor da 1� parcela
			blnPreferencial = False
			s_pg = Trim(op_pse_prim_prest_forma_pagto)
			if s_pg <> "" then
				for i=Lbound(vMPN2) to Ubound(vMPN2)
				'	O meio de pagamento selecionado � um dos preferenciais
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
				'	O meio de pagamento selecionado � um dos preferenciais
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
			
		'	O montante a pagar por meio de pagamento preferencial � maior que 50% do total?
			if vlNivel2 > (vl_total/2) then
				if tipo_cliente = ID_PJ then
					perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
				else
					perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
					end if
				end if
			end if
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
							" ON ((t_PRODUTO.fabricante=t_PRODUTO_LOJA.fabricante) AND (t_PRODUTO.produto=t_PRODUTO_LOJA.produto))" & _
						" INNER JOIN t_FABRICANTE" & _
							" ON (t_PRODUTO.fabricante=t_FABRICANTE.fabricante)" & _
					" WHERE" & _
						" (t_PRODUTO.fabricante='" & .fabricante & "')" & _
						" AND (t_PRODUTO.produto='" & .produto & "')" & _
						" AND (loja='" & loja & "')"
				set rs = cn.execute(s)
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " N�O est� cadastrado para a loja " & loja
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
							alerta=alerta & "Op��o de parcelamento n�o dispon�vel para fornecedor " & .fabricante & ": " & decodificaCustoFinancFornecQtdeParcelas(c_custoFinancFornecTipoParcelamento, c_custoFinancFornecQtdeParcelas) & " parcela(s)"
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
								" AND (id_cliente='" & cliente_selecionado & "')" & _
								" AND (fabricante='" & .fabricante & "')" & _
								" AND (produto='" & .produto & "')" & _
								" AND (loja='" & loja & "')" & _
								" AND (data >= " & bd_formata_data_hora(Now-converte_min_to_dec(TIMEOUT_DESCONTO_EM_MIN)) & ")" & _
							" ORDER BY" & _
								" data DESC"
						set rs=cn.execute(s)
						if rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": desconto de " & formata_perc_desc(.desc_dado) & "% excede o m�ximo permitido."
						else
							if .desc_dado > rs("desc_max") then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": desconto de " & formata_perc_desc(.desc_dado) & "% excede o m�ximo autorizado."
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
	
'	RECUPERA OS PRODUTOS QUE O CLIENTE CONCORDOU EM COMPRAR MESMO SEM PRESEN�A NO ESTOQUE.
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
		
	dim r_cliente
	set r_cliente = New cl_CLIENTE
	call x_cliente_bd(cliente_selecionado, r_cliente)

	'TRATAMENTO PARA CADASTRAMENTO DE PEDIDOS DO SITE MAGENTO DA BONSHOP
	dim c_mag_installer_document, percCommissionValue, percCommissionDiscount, vlMagentoShippingAmount
	dim sIdIndicador, sNomeIndicador, sIdVendedor, sNomeVendedor
	c_mag_installer_document = ""
	sIdIndicador = ""
	sNomeIndicador = ""
	sIdVendedor = ""
	sNomeVendedor = ""
	percCommissionValue = 0
	percCommissionDiscount = 0
	vlMagentoShippingAmount = 0

	if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) And blnMagentoPedidoComIndicador then
		if alerta = "" then
			s = "SELECT " & _
					"*" & _
				" FROM t_MAGENTO_API_PEDIDO_XML" & _
				" WHERE" & _
					" (id = " & id_magento_api_pedido_xml & ")"
			if tMAP_XML.State <> 0 then tMAP_XML.Close
			tMAP_XML.open s, cn
			if tMAP_XML.Eof then
				alerta = "Falha ao tentar localizar no banco de dados o registro com os dados do pedido Magento consultados via API (id = " & id_magento_api_pedido_xml & ")"
			else
				c_mag_installer_document = retorna_so_digitos(Trim("" & tMAP_XML("installer_document")))
				percCommissionValue = tMAP_XML("commission_value")
				percCommissionDiscount = tMAP_XML("commission_discount")
				vlMagentoShippingAmount = tMAP_XML("shipping_amount")

				if c_mag_installer_document = "" then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O pedido Magento n� " & c_numero_magento & " n�o informa o CPF/CNPJ do indicador!"
				else
					If Not cria_recordset_otimista(tOI, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
					s = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (cnpj_cpf = '" & retorna_so_digitos(c_mag_installer_document) & "') AND (Convert(smallint, loja) = " & loja & ")"
					if tOI.State <> 0 then tOI.Close
					tOI.open s, cn
					if tOI.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "O pedido Magento n� " & c_numero_magento & " especifica o indicador com CPF/CNPJ " & cnpj_cpf_formata(c_mag_installer_document) & " que n�o foi localizado no banco de dados (loja: " & loja & ")!"
					else
						sIdIndicador = Trim("" & tOI("apelido"))
						sNomeIndicador = Trim("" & tOI("razao_social_nome"))
						sIdVendedor = Trim("" & tOI("vendedor"))
						sNomeVendedor = Trim("" & x_usuario (sIdVendedor))
						end if
					if tOI.State <> 0 then tOI.Close
					set tOI = nothing
					end if
				end if 'if tMAP_XML.Eof
			end if 'if alerta = ""

		if alerta = "" then
			if sIdIndicador = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "N�o foi poss�vel localizar no banco de dados o indicador com CPF/CNPJ " & cnpj_cpf_formata(c_mag_installer_document)
				end if
			end if 'if alerta = ""

		if alerta = "" then
			if UCase(sIdIndicador) <> UCase(c_indicador) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Inconsist�ncia encontrada na identifica��o do indicador: '" & sIdIndicador & "' e '" & c_indicador & "'"
				end if
			end if 'if alerta = ""

		if alerta = "" then
			if sIdVendedor = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "N�o foi poss�vel determinar o vendedor associado ao indicador '" & sIdIndicador & "'"
				end if
			end if 'if alerta = ""
		end if 'if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO

'	L�GICA P/ CONSUMO DO ESTOQUE
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
				alerta=alerta & "O CD selecionado manualmente � inv�lido"
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
		'VERIFICA SE AS REGRAS ASSOCIADAS AOS PRODUTOS EST�O OK
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			if Trim(vProdRegra(iRegra).produto) <> "" then
				if converte_numero(vProdRegra(iRegra).regra.id) = 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " n�o possui regra de consumo do estoque associada"
				elseif vProdRegra(iRegra).regra.st_inativo = 1 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " est� desativada"
				elseif vProdRegra(iRegra).regra.regraUF.st_inativo = 1 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " est� bloqueada para a UF '" & r_cliente.uf & "'"
				elseif vProdRegra(iRegra).regra.regraUF.regraPessoa.st_inativo = 1 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " est� bloqueada para clientes '" & descricao_tipo_pessoa & "' da UF '" & r_cliente.uf & "'"
				elseif converte_numero(vProdRegra(iRegra).regra.regraUF.regraPessoa.spe_id_nfe_emitente) = 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " n�o especifica nenhum CD para aguardar produtos sem presen�a no estoque para clientes '" & descricao_tipo_pessoa & "' da UF '" & r_cliente.uf & "'"
				else
					qtde_CD_ativo = 0
					for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
						if converte_numero(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente) > 0 then
							if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0 then
								qtde_CD_ativo = qtde_CD_ativo + 1
								end if
							end if
						next
					'A SELE��O MANUAL DE CD PERMITE O USO DE CD DESATIVADO
					if (qtde_CD_ativo = 0) And (id_nfe_emitente_selecao_manual = 0) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " n�o especifica nenhum CD ativo para clientes '" & descricao_tipo_pessoa & "' da UF '" & r_cliente.uf & "'"
						end if
					end if
				end if
			next
		end if 'if alerta=""
	
	'NO CASO DE SELE��O MANUAL DO CD, VERIFICA SE O CD SELECIONADO EST� HABILITADO EM TODAS AS REGRAS
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
					alerta_aux=alerta_aux & "Produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & ": regra '" & vProdRegra(iRegra).regra.apelido & "' (Id=" & vProdRegra(iRegra).regra.id & ") n�o permite o CD '" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_selecao_manual) & "'"
				elseif blnDesativado then
					'16/09/2017: FOI REALIZADA UMA ALTERA��O P/ QUE A SELE��O MANUAL DE CD PERMITA O USO DE CD DESATIVADO
					'alerta_aux=texto_add_br(alerta_aux)
					'alerta_aux=alerta_aux & "Produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & ": regra '" & vProdRegra(iRegra).regra.apelido & "' (Id=" & vProdRegra(iRegra).regra.id & ") define o CD '" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_selecao_manual) & "' como 'desativado'"
					end if
				next

			if alerta_aux <> "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O CD selecionado manualmente n�o pode ser usado devido aos seguintes motivos:"
				alerta=texto_add_br(alerta)
				alerta=alerta & alerta_aux
				end if
			end if
		end if
	
	dim erro_produto_indisponivel
	if alerta="" then
		'OBT�M DISPONIBILIDADE DO PRODUTO NO ESTOQUE
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			if Trim(vProdRegra(iRegra).produto) <> "" then
				for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
					if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
						( (id_nfe_emitente_selecao_manual = 0) Or (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) ) then
						'VERIFICA SE O CD EST� HABILITADO
						'IMPORTANTE: A SELE��O MANUAL DE CD PERMITE O USO DE CD DESATIVADO
						if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0) Or (id_nfe_emitente_selecao_manual <> 0) then
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
	
'	H� PRODUTO C/ ESTOQUE INSUFICIENTE (SOMANDO-SE O ESTOQUE DE TODAS AS EMPRESAS CANDIDATAS)
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
								'VERIFICA SE O CD EST� HABILITADO
								'IMPORTANTE: A SELE��O MANUAL DE CD PERMITE O USO DE CD DESATIVADO
								if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0) Or (id_nfe_emitente_selecao_manual <> 0) then
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
	
'	ANALISA A QUANTIDADE DE PEDIDOS QUE SER�O CADASTRADOS (AUTO-SPLIT)
'	INICIALIZA O CAMPO 'qtde_solicitada', POIS ELE IR� CONTROLAR A QUANTIDADE A SER ALOCADA NO ESTOQUE DE CADA EMPRESA
	if alerta = "" then
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
				vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = 0
				next
			next
		end if 'if alerta=""
	
'	REALIZA A AN�LISE DA QUANTIDADE DE PEDIDOS NECESS�RIA (AUTO-SPLIT)
	dim qtde_a_alocar
	if alerta = "" then
		for iItem=Lbound(v_item) to Ubound(v_item)
			if Trim(v_item(iItem).produto) <> "" then
			'	OS CD'S EST�O ORDENADOS DE ACORDO C/ A PRIORIZA��O DEFINIDA PELA REGRA DE CONSUMO DO ESTOQUE
			'	SE O PRIMEIRO CD HABILITADO N�O PUDER ATENDER INTEGRALMENTE A QUANTIDADE SOLICITADA DO PRODUTO,
			'	A QUANTIDADE RESTANTE SER� CONSUMIDA DOS DEMAIS CD'S.
			'	SE HOUVER ALGUMA QUANTIDADE RESIDUAL P/ FICAR NA LISTA DE PRODUTOS SEM PRESEN�A NO ESTOQUE:
			'		1) SELE��O AUTOM�TICA DE CD: A QUANTIDADE PENDENTE FICAR� ALOCADA NO CD DEFINIDO P/ TAL
			'		2) SELE��O MANUAL DE CD: A QUANTIDADE PENDENTE FICAR� ALOCADA NO CD SELECIONADO MANUALMENTE
				qtde_a_alocar = v_item(iItem).qtde
				for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
					if qtde_a_alocar = 0 then exit for

					if Trim(vProdRegra(iRegra).produto) <> "" then
						for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
							if qtde_a_alocar = 0 then exit for

							if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
								( (id_nfe_emitente_selecao_manual = 0) Or (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) ) then
								'VERIFICA SE O CD EST� HABILITADO
								'IMPORTANTE: A SELE��O MANUAL DE CD PERMITE O USO DE CD DESATIVADO
								if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0) Or (id_nfe_emitente_selecao_manual <> 0) then
									if (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) then
										if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque >= qtde_a_alocar then
										'	H� QUANTIDADE DISPON�VEL SUFICIENTE PARA INTEGRALMENTE
											vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = qtde_a_alocar
											qtde_a_alocar = 0
										elseif vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque > 0 then
										'	A QUANTIDADE DISPON�VEL NO ESTOQUE � INSUFICIENTE P/ ATENDER INTEGRALMENTE � QUANTIDADE SOLICITADA,
										'	PORTANTO, A QUANTIDADE DISPON�VEL NESTE CD SER� CONSUMIDA P/ ATENDER PARCIALMENTE � REQUISI��O E A
										'	QUANTIDADE REMANESCENTE SER� ATENDIDA PELO PR�XIMO CD DA LISTA OU ENT�O SER� COLOCADA NA LISTA DE
										'	PRODUTOS SEM PRESEN�A NO ESTOQUE DO CD SELECIONADO P/ TAL NA REGRA.
											vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque
											qtde_a_alocar = qtde_a_alocar - vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque
											end if
										end if
									end if
								end if
							next
						end if
					next

			'	RESTOU SALDO A ALOCAR NA LISTA DE PRODUTOS SEM PRESEN�A NO ESTOQUE?
				if qtde_a_alocar > 0 then
				'	LOCALIZA E ALOCA A QUANTIDADE PENDENTE:
				'		1) SELE��O AUTOM�TICA DE CD: A QUANTIDADE PENDENTE FICAR� ALOCADA NO CD DEFINIDO P/ TAL
				'		2) SELE��O MANUAL DE CD: A QUANTIDADE PENDENTE FICAR� ALOCADA NO CD SELECIONADO MANUALMENTE
					for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
						if qtde_a_alocar = 0 then exit for

						if Trim(vProdRegra(iRegra).produto) <> "" then
							for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
								if qtde_a_alocar = 0 then exit for

								if id_nfe_emitente_selecao_manual = 0 then
									'MODO DE SELE��O AUTOM�TICO
									if ( (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) ) And _
										(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
										(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = vProdRegra(iRegra).regra.regraUF.regraPessoa.spe_id_nfe_emitente) then
										vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada + qtde_a_alocar
										qtde_a_alocar = 0
										exit for
										end if
								else
									'MODO DE SELE��O MANUAL
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
					alerta=alerta & "Falha ao processar a aloca��o de produtos no estoque: restaram " & qtde_a_alocar & " unidades do produto (" & v_item(iItem).fabricante & ")" & v_item(iItem).produto & " que n�o puderam ser alocados na lista de produtos sem presen�a no estoque de nenhum CD"
					end if
				end if
			next
		end if 'if alerta=""
	
'	CONTAGEM DE EMPRESAS QUE SER�O USADAS NO AUTO-SPLIT, OU SEJA, A QUANTIDADE DE PEDIDOS QUE SER� CADASTRADA, J� QUE CADA PEDIDO SE REFERE AO ESTOQUE DE UMA EMPRESA
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
							'	SE O CD AINDA N�O CONSTA DA LISTA, INCLUI
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
	
	
'	OBT�M O VALOR LIMITE P/ APROVA��O AUTOM�TICA DA AN�LISE DE CR�DITO
	if alerta = "" then
		s = "SELECT nsu FROM t_CONTROLE WHERE (id_nsu = '" & ID_PARAM_CAD_VL_APROV_AUTO_ANALISE_CREDITO & "')"
		set rs = cn.execute(s)
		if Not rs.Eof then
			vl_aprov_auto_analise_credito = converte_numero(rs("nsu"))
			end if
		if rs.State <> 0 then rs.Close
		end if
	
'	OBT�M O PERCENTUAL DA COMISS�O
	if alerta = "" then
		if s_loja_indicou<>"" then
			s = "SELECT loja, comissao_indicacao FROM t_LOJA WHERE (loja='" & s_loja_indicou & "')"
			set rs = cn.execute(s)
			if Not rs.Eof then
				comissao_loja_indicou = rs("comissao_indicacao")
			else
				alerta = "Loja " & s_loja_indicou & " n�o est� cadastrada."
				end if
			end if
		end if
	
	if alerta="" then
		if rb_indicacao = "" then
			alerta = "Informe se o pedido � com indica��o ou n�o."
		elseif rb_indicacao = "S" then
			if c_indicador = "" then
				alerta = "Informe quem � o indicador."
			elseif rb_RA = "" then
				alerta = "Informe se o pedido possui RA ou n�o."
		'	POR SOLICITA��O DO ROG�RIO, A CONSIST�NCIA DO LIMITE DE COMPRAS FOI DESATIVADA (NOV/2008)
'			elseif (vl_limite_mensal_disponivel - vl_total) <= 0 then
'				alerta = "N�o � poss�vel cadastrar este pedido porque excede o valor do limite mensal estabelecido para o indicador (" & c_indicador & ")"
			elseif rb_garantia_indicador = "" then
				alerta = "Informe se o pedido � garantido pelo indicador ou n�o."
				end if
			end if
		end if
	
	if alerta = "" then
		if s_etg_imediata = "" then
			alerta = "� necess�rio selecionar uma op��o para o campo 'Entrega Imediata'."
			end if
		end if

	if alerta = "" then
		if s_bem_uso_consumo = "" then
			alerta = "� necess�rio informar se � 'Bem de Uso/Consumo'."
			end if
		end if
	
	if alerta = "" then
		if c_exibir_campo_instalador_instala = "S" then
			if s_instalador_instala = "" then
				alerta = "� necess�rio preencher o campo 'Instalador Instala'."
				end if
			end if
		end if
	
'	CEP
	if alerta = "" then
		if rb_end_entrega = "S" then
			if EndEtg_cep = "" then
				alerta = "Informe o CEP do endere�o de entrega."
				end if
			end if
		end if
	
'	CONSIST�NCIA DO VALOR TOTAL DA FORMA DE PAGAMENTO
	if alerta = "" then
		if rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA then vlTotalFormaPagto = vl_total_NF
		if Abs(vlTotalFormaPagto-vl_total_NF) > 0.1 then
			alerta = "H� diverg�ncia entre o valor total do pedido (" & SIMBOLO_MONETARIO & " " & formata_moeda(vl_total_NF) & ") e o valor total descrito atrav�s da forma de pagamento (" & SIMBOLO_MONETARIO & " " & formata_moeda(vlTotalFormaPagto) & ")!!"
			end if
		end if
	
'	OBTEN��O DE TRANSPORTADORA QUE ATENDA AO CEP INFORMADO, SE HOUVER
	dim sTranspSelAutoTransportadoraId, sTranspSelAutoCep, iTranspSelAutoTipoEndereco, iTranspSelAutoStatus
	sTranspSelAutoTransportadoraId = ""
	if alerta = "" then
		if rb_end_entrega = "S" then
			if EndEtg_cep <> "" then
				sTranspSelAutoTransportadoraId = obtem_transportadora_pelo_cep(retorna_so_digitos(EndEtg_cep))
				if sTranspSelAutoTransportadoraId <> "" then
					sTranspSelAutoCep = retorna_so_digitos(EndEtg_cep)
					iTranspSelAutoTipoEndereco = TRANSPORTADORA_SELECAO_AUTO_TIPO_ENDERECO_ENTREGA
					iTranspSelAutoStatus = TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_S
					end if
				end if
		else
			if Trim("" & t_CLIENTE("cep")) <> "" then
				sTranspSelAutoTransportadoraId = obtem_transportadora_pelo_cep(retorna_so_digitos(Trim("" & t_CLIENTE("cep"))))
				if sTranspSelAutoTransportadoraId <> "" then
					sTranspSelAutoCep = retorna_so_digitos(Trim("" & t_CLIENTE("cep")))
					iTranspSelAutoTipoEndereco = TRANSPORTADORA_SELECAO_AUTO_TIPO_ENDERECO_CLIENTE
					iTranspSelAutoStatus = TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_S
					end if
				end if
			end if
		end if

	if alerta = "" then
		if rb_end_entrega = "S" then 
			if (EndEtg_endereco<>"") Or (EndEtg_bairro<>"") Or (EndEtg_cidade<>"") Or (EndEtg_uf<>"") Or (EndEtg_cep<>"") then
				if EndEtg_endereco="" then
					alerta="PREENCHA O ENDERE�O DE ENTREGA."
				elseif Len(EndEtg_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
					alerta="ENDERE�O DE ENTREGA EXCEDE O TAMANHO M�XIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(EndEtg_endereco)) & " CARACTERES<br>TAMANHO M�XIMO: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " CARACTERES"
				elseif EndEtg_endereco_numero="" then
					alerta="PREENCHA O N�MERO DO ENDERE�O DE ENTREGA."
				elseif EndEtg_cidade="" then
					alerta="PREENCHA A CIDADE DO ENDERE�O DE ENTREGA."
				elseif EndEtg_uf="" then
					alerta="PREENCHA A UF DO ENDERE�O DE ENTREGA."
				elseif EndEtg_cep="" then
					alerta="PREENCHA O CEP DO ENDERE�O DE ENTREGA."
					end if
				end if
			end if
		end if
	
	'TRATAMENTO PARA CADASTRAMENTO DE PEDIDOS DO SITE MAGENTO DA BONSHOP
	if (loja = NUMERO_LOJA_BONSHOP) And (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) then
		if alerta = "" then
			if s_pedido_ac = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Informe o n� pedido Magento"
				end if
			
			if s_pedido_ac <> "" then
				if s_pedido_ac <> retorna_so_digitos(s_pedido_ac) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O n�mero do pedido Magento deve conter apenas d�gitos"
					end if

				do while Len(s_pedido_ac) < 9
					if Len(s_pedido_ac) = 8 then
						s_pedido_ac = "2" & s_pedido_ac
					else
						s_pedido_ac = "0" & s_pedido_ac
						end if
					Loop
				
				if Left(s_pedido_ac, 1) <> "2" then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O n�mero do pedido Magento inicia com d�gito inv�lido para a loja " & loja
					end if
				end if 'if s_pedido_ac <> ""
			end if 'if alerta = ""

	'	VERIFICA SE H� PEDIDO J� CADASTRADO COM O MESMO N� PEDIDO MAGENTO (POSS�VEL CADASTRO EM DUPLICIDADE)
		if alerta = "" then
			if s_pedido_ac <> "" then
				s = "SELECT" & _
						" tP.pedido," & _
						" tP.pedido_bs_x_ac," & _
						" tP.data_hora," & _
						" tP.vendedor," & _
						" tU.nome AS nome_vendedor," & _
						" tP.usuario_cadastro," & _
						" tUC.nome AS nome_usuario_cadastro," & _
						" tC.cnpj_cpf," & _
						" tC.nome AS nome_cliente" & _
					" FROM t_PEDIDO tP" & _
						" INNER JOIN t_CLIENTE tC ON (tP.id_cliente = tC.id)" & _
						" LEFT JOIN t_USUARIO tU ON (tP.vendedor = tU.usuario)" & _
						" LEFT JOIN t_USUARIO tUC ON (tP.usuario_cadastro = tUC.usuario)" & _
					" WHERE" & _
						" (tP.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
						" AND (pedido_bs_x_ac = '" & s_pedido_ac & "')" & _
						" AND (" & _
							"tP.pedido NOT IN (" & _
								"SELECT DISTINCT" & _
									" pedido" & _
								" FROM t_PEDIDO_DEVOLUCAO tPD" & _
								" WHERE" & _
									" (tPD.pedido = tP.pedido)" & _
									" AND (status IN (" & _
										COD_ST_PEDIDO_DEVOLUCAO__FINALIZADA & "," & _
										COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA & "," & _
										COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO & "," & _
										COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA & _
										")" & _
									")" & _
								")" & _
							")" & _
						" AND (" & _
							"tP.pedido NOT IN (" & _
								"SELECT DISTINCT" & _
									" pedido" & _
								" FROM t_PEDIDO_ITEM_DEVOLVIDO tPID" & _
								" WHERE" & _
									" (tPID.pedido = tP.pedido)" & _
								")" & _
							")"
				set rs = cn.execute(s)
				if Not rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O n� pedido Magento " & Trim("" & rs("pedido_bs_x_ac")) & " j� est� cadastrado no pedido " & Trim("" & rs("pedido"))
					alerta=texto_add_br(alerta)
					alerta=alerta & "Data de cadastramento do pedido: " & formata_data_hora_sem_seg(rs("data_hora"))
					alerta=texto_add_br(alerta)
					if UCase(Trim("" & rs("vendedor"))) = UCase(Trim("" & rs("usuario_cadastro"))) then
						alerta=alerta & "Cadastrado por: " & Trim("" & rs("vendedor"))
						if Ucase(Trim("" & rs("vendedor"))) <> Ucase(Trim("" & rs("nome_vendedor"))) then alerta=alerta & " (" & Trim("" & rs("nome_vendedor")) & ")"
					else
						alerta=alerta & "Cadastrado por: " & Trim("" & rs("usuario_cadastro"))
						if Ucase(Trim("" & rs("usuario_cadastro"))) <> Ucase(Trim("" & rs("nome_usuario_cadastro"))) then alerta=alerta & " (" & Trim("" & rs("nome_usuario_cadastro")) & ")"
						alerta=texto_add_br(alerta)
						alerta=alerta & "Vendedor: " & Trim("" & rs("vendedor"))
						if Ucase(Trim("" & rs("vendedor"))) <> Ucase(Trim("" & rs("nome_vendedor"))) then alerta=alerta & " (" & Trim("" & rs("nome_vendedor")) & ")"
						end if
					alerta=texto_add_br(alerta)
					alerta=alerta & "Cliente: " & cnpj_cpf_formata(Trim("" & rs("cnpj_cpf"))) & " - " & Trim("" & rs("nome_cliente"))
					end if 'if Not rs.Eof
				end if 'if s_pedido_ac <> "" then
			end if 'if alerta = "" then
		end if 'if (loja = NUMERO_LOJA_BONSHOP) And (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO)

	'TRATAMENTO PARA CADASTRAMENTO DE PEDIDOS DO SITE MAGENTO DO ARCLUBE
	dim blnPedidoECommerceOrigemMarketplaceCreditoOkAutomatico
	blnPedidoECommerceOrigemMarketplaceCreditoOkAutomatico = False
	if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
		if alerta = "" then
			if s_origem_pedido = "" then
				alerta = "Informe a origem do pedido"
				end if
			end if
		
		if alerta = "" then
		'	SOMENTE PEDIDO ORIGINADO PELO TELEVENDAS DO ARCLUBE PODE FICAR SEM O N� PEDIDO MAGENTO
			if Trim(s_origem_pedido) <> "002" then
				if s_pedido_ac = "" then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Informe o n� Magento"
					end if
				end if

            if s_pedido_ac <> "" then
                if s_pedido_ac <> retorna_so_digitos(s_pedido_ac) then
                    alerta=texto_add_br(alerta)
	                alerta=alerta & "O n�mero Magento deve conter apenas d�gitos"
                end if

                do while Len(s_pedido_ac) < 9
                    if Len(s_pedido_ac) = 8 then
                        s_pedido_ac = "1" & s_pedido_ac
                    else
                        s_pedido_ac = "0" & s_pedido_ac
						end if
					Loop

				if Left(s_pedido_ac, 1) <> "1" then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O n�mero do pedido Magento inicia com d�gito inv�lido para a loja " & loja
					end if
				end if 'if s_pedido_ac <> ""
			
			s = "SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo = 'PedidoECommerce_Origem') AND (codigo = '" & s_origem_pedido & "')"
			set rs = cn.execute(s)
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "C�digo de origem do pedido (marketplace) n�o cadastrado: " & s_origem_pedido
			else
			'	PROCESSA OS PAR�METROS DEFINIDOS PARA A ORIGEM (GRUPO)
				s = "SELECT * FROM T_CODIGO_DESCRICAO WHERE (grupo = 'PedidoECommerce_Origem_Grupo') AND (codigo = '" & Trim("" & rs("codigo_pai")) & "')"
				set rs2 = cn.execute(s)
				if Not rs2.Eof then
				'	OBT�M O PERCENTUAL DE COMISS�O DO MARKETPLACE
					perc_RT = rs2("parametro_campo_real")
				'	DEVE COLOCAR AUTOMATICAMENTE COM 'CR�DITO OK'?
					if rs2("parametro_1_campo_flag") = 1 then blnPedidoECommerceOrigemMarketplaceCreditoOkAutomatico = True
				'	N� PEDIDO MARKETPLACE � OBRIGAT�RIO?
					if rs2("parametro_2_campo_flag") = 1 then
						if s_numero_mktplace = "" then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Informe o n� do pedido do marketplace (" & Trim("" & rs("descricao")) & ")"
							end if
						end if
					end if 'if Not rs2.Eof then
				end if 'if rs.Eof
			if rs.State <> 0 then rs.Close
			end if 'if alerta = "" then

            if s_numero_mktplace <> "" then
                s = ""
                For i = 1 To Len(s_numero_mktplace)
                    c = Mid(s_numero_mktplace, i, 1)
                    If IsNumeric(c) Or c = chr(45) Then s = s & c
                    Next
                if s_numero_mktplace <> s then
                    alerta=texto_add_br(alerta)
					alerta=alerta & "O n�mero Marketplace deve conter apenas d�gitos e h�fen"
					end if
				end if

	'	VERIFICA SE H� PEDIDO J� CADASTRADO COM O MESMO N� PEDIDO MAGENTO (POSS�VEL CADASTRO EM DUPLICIDADE)
		if alerta = "" then
			if s_pedido_ac <> "" then
				s = "SELECT" & _
						" tP.pedido," & _
						" tP.pedido_bs_x_ac," & _
						" tP.data_hora," & _
						" tP.vendedor," & _
						" tU.nome AS nome_vendedor," & _
						" tC.cnpj_cpf," & _
						" tC.nome AS nome_cliente" & _
					" FROM t_PEDIDO tP" & _
						" INNER JOIN t_CLIENTE tC ON (tP.id_cliente = tC.id)" & _
						" LEFT JOIN t_USUARIO tU ON (tP.vendedor = tU.usuario)" & _
					" WHERE" & _
						" (tP.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
						" AND (pedido_bs_x_ac = '" & s_pedido_ac & "')" & _
						" AND (" & _
							"tP.pedido NOT IN (" & _
								"SELECT DISTINCT" & _
									" pedido" & _
								" FROM t_PEDIDO_DEVOLUCAO tPD" & _
								" WHERE" & _
									" (tPD.pedido = tP.pedido)" & _
									" AND (tPD.status IN (" & _
										COD_ST_PEDIDO_DEVOLUCAO__FINALIZADA & "," & _
										COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA & "," & _
										COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO & "," & _
										COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA & _
										")" & _
									")" & _
								")" & _
							")" & _
						" AND (" & _
							"tP.pedido NOT IN (" & _
								"SELECT DISTINCT" & _
									" pedido" & _
								" FROM t_PEDIDO_ITEM_DEVOLVIDO tPID" & _
								" WHERE" & _
									" (tPID.pedido = tP.pedido)" & _
								")" & _
							")"
				set rs = cn.execute(s)
				if Not rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O n� pedido Magento " & Trim("" & rs("pedido_bs_x_ac")) & " j� est� cadastrado no pedido " & Trim("" & rs("pedido"))
					alerta=texto_add_br(alerta)
					alerta=alerta & "Data de cadastramento do pedido: " & formata_data_hora_sem_seg(rs("data_hora"))
					alerta=texto_add_br(alerta)
					alerta=alerta & "Cadastrado por: " & Trim("" & rs("vendedor"))
					if Ucase(Trim("" & rs("vendedor"))) <> Ucase(Trim("" & rs("nome_vendedor"))) then alerta=alerta & " (" & Trim("" & rs("nome_vendedor")) & ")"
					alerta=texto_add_br(alerta)
					alerta=alerta & "Cliente: " & cnpj_cpf_formata(Trim("" & rs("cnpj_cpf"))) & " - " & Trim("" & rs("nome_cliente"))
					end if 'if Not rs.Eof
				end if 'if s_pedido_ac <> ""
			end if 'if alerta = ""

	'	VERIFICA SE H� PEDIDO J� CADASTRADO COM O MESMO N� PEDIDO MAGENTO (POSS�VEL CADASTRO EM DUPLICIDADE)
		if alerta = "" then
			if s_numero_mktplace <> "" then
				s = "SELECT" & _
						" tP.pedido," & _
						" tP.pedido_bs_x_ac," & _
						" tP.pedido_bs_x_marketplace," & _
						" tP.data_hora," & _
						" tP.vendedor," & _
						" tU.nome AS nome_vendedor," & _
						" tC.cnpj_cpf," & _
						" tC.nome AS nome_cliente" & _
					" FROM t_PEDIDO tP" & _
						" INNER JOIN t_CLIENTE tC ON (tP.id_cliente = tC.id)" & _
						" LEFT JOIN t_USUARIO tU ON (tP.vendedor = tU.usuario)" & _
					" WHERE" & _
						" (tP.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
						" AND (pedido_bs_x_marketplace = '" & s_numero_mktplace & "')" & _
						" AND (" & _
							"tP.pedido NOT IN (" & _
								"SELECT DISTINCT" & _
									" pedido" & _
								" FROM t_PEDIDO_DEVOLUCAO tPD" & _
								" WHERE" & _
									" (tPD.pedido = tP.pedido)" & _
									" AND (tPD.status IN (" & _
										COD_ST_PEDIDO_DEVOLUCAO__FINALIZADA & "," & _
										COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA & "," & _
										COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO & "," & _
										COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA & _
										")" & _
									")" & _
								")" & _
							")" & _
						" AND (" & _
							"tP.pedido NOT IN (" & _
								"SELECT DISTINCT" & _
									" pedido" & _
								" FROM t_PEDIDO_ITEM_DEVOLVIDO tPID" & _
								" WHERE" & _
									" (tPID.pedido = tP.pedido)" & _
								")" & _
							")"
				set rs = cn.execute(s)
				if Not rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O n� pedido marketplace " & Trim("" & rs("pedido_bs_x_marketplace")) & " j� est� cadastrado no pedido " & Trim("" & rs("pedido"))
					alerta=texto_add_br(alerta)
					alerta=alerta & "N� pedido Magento: " & Trim("" & rs("pedido_bs_x_ac"))
					alerta=texto_add_br(alerta)
					alerta=alerta & "Data de cadastramento do pedido: " & formata_data_hora_sem_seg(rs("data_hora"))
					alerta=texto_add_br(alerta)
					alerta=alerta & "Cadastrado por: " & Trim("" & rs("vendedor"))
					if Ucase(Trim("" & rs("vendedor"))) <> Ucase(Trim("" & rs("nome_vendedor"))) then alerta=alerta & " (" & Trim("" & rs("nome_vendedor")) & ")"
					alerta=texto_add_br(alerta)
					alerta=alerta & "Cliente: " & cnpj_cpf_formata(Trim("" & rs("cnpj_cpf"))) & " - " & Trim("" & rs("nome_cliente"))
					end if 'if Not rs.Eof then
				end if 'if s_numero_mktplace <> "" then
			end if 'if alerta = ""
		end if 'if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE

'	CADASTRA O PEDIDO E PROCESSA A MOVIMENTA��O NO ESTOQUE
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
					if s_recebido <> "" then rs("st_recebido")=s_recebido
					rs("obs_1")=s_obs1
					rs("obs_2")=s_obs2
				'	Forma de Pagamento (nova vers�o)
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
					'	Entrada + Presta��es
						rs("qtde_parcelas")=CLng(c_pce_prestacao_qtde)+1
					elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
						rs("pse_forma_pagto_prim_prest") = CLng(op_pse_prim_prest_forma_pagto)
						rs("pse_forma_pagto_demais_prest") = CLng(op_pse_demais_prest_forma_pagto)
						rs("pse_prim_prest_valor") = converte_numero(c_pse_prim_prest_valor)
						rs("pse_prim_prest_apos") = CLng(c_pse_prim_prest_apos)
						rs("pse_demais_prest_qtde") = CLng(c_pse_demais_prest_qtde)
						rs("pse_demais_prest_valor") = converte_numero(c_pse_demais_prest_valor)
						rs("pse_demais_prest_periodo") = CLng(c_pse_demais_prest_periodo)
					'	1� presta��o + Demais presta��es
						rs("qtde_parcelas")=CLng(c_pse_demais_prest_qtde)+1
						end if
					rs("forma_pagto")=s_forma_pagto
					rs("vl_total_familia")=vl_total
					if blnPedidoECommerceOrigemMarketplaceCreditoOkAutomatico then
						rs("analise_credito")=Clng(COD_AN_CREDITO_OK)
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")="AUTOM�TICO"
					elseif vl_total <= vl_aprov_auto_analise_credito then
						rs("analise_credito")=Clng(COD_AN_CREDITO_OK)
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")="AUTOM�TICO"
					elseif Cstr(loja) = Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) And (rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA) And (CStr(op_av_forma_pagto) = Cstr(ID_FORMA_PAGTO_DINHEIRO)) then
						rs("analise_credito")=Clng(COD_AN_CREDITO_PENDENTE_VENDAS)
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")="AUTOM�TICO"
					elseif Cstr(loja) = Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) And (rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA) And (CStr(op_av_forma_pagto) = Cstr(ID_FORMA_PAGTO_BOLETO_AV)) then
						rs("analise_credito")=Clng(COD_AN_CREDITO_PENDENTE_VENDAS)
						rs("analise_credito_pendente_vendas_motivo")="006" 'Aguardando Emiss�o do Boleto Avulso
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")="AUTOM�TICO"
					elseif (rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA) And ( (CStr(op_av_forma_pagto) = CStr(ID_FORMA_PAGTO_DEPOSITO)) Or (CStr(op_av_forma_pagto) = Cstr(ID_FORMA_PAGTO_BOLETO_AV)) ) then
						rs("analise_credito")=Clng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")="AUTOM�TICO"
					elseif (Cstr(loja) = Cstr(NUMERO_LOJA_TRANSFERENCIA)) Or (Cstr(loja) = Cstr(NUMERO_LOJA_KITS)) then
						'Lojas usadas para pedidos de opera��es internas
						rs("analise_credito")=Clng(COD_AN_CREDITO_OK)
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")="AUTOM�TICO"
					elseif (rb_forma_pagto = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) then
						rs("analise_credito")=Clng(COD_AN_CREDITO_PENDENTE_VENDAS)
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")="AUTOM�TICO"
						end if

				'	CUSTO FINANCEIRO FORNECEDOR
					rs("custoFinancFornecTipoParcelamento") = c_custoFinancFornecTipoParcelamento
					rs("custoFinancFornecQtdeParcelas") = c_custoFinancFornecQtdeParcelas
					rs("vl_total_NF") = vl_total_NF
					rs("vl_total_RA") = vl_total_RA
					rs("perc_RT") = perc_RT
					rs("perc_desagio_RA") = perc_desagio_RA
					rs("perc_limite_RA_sem_desagio") = perc_limite_RA_sem_desagio

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

			'	CAMPOS ARMAZENADOS TANTO NO PEDIDO-PAI QUANTO NO PEDIDO-FILHOTE
				rs("id_cliente")=cliente_selecionado
				rs("midia")=midia_selecionada
				rs("servicos")=""
				if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) And blnMagentoPedidoComIndicador then
					rs("vendedor")=sIdVendedor
				else
					rs("vendedor")=usuario
					end if
				rs("usuario_cadastro")=usuario
				rs("st_entrega")=""
				rs("pedido_bs_x_at")=c_ped_bonshop
				if s_etg_imediata <> "" then 
					rs("st_etg_imediata")=CLng(s_etg_imediata)
					rs("etg_imediata_data")=Now
					rs("etg_imediata_usuario")=usuario
					end if
				if s_bem_uso_consumo <> "" then 
					rs("StBemUsoConsumo")=CLng(s_bem_uso_consumo)
					end if
				if s_instalador_instala <> "" then
					rs("InstaladorInstalaStatus")=CLng(s_instalador_instala)
					rs("InstaladorInstalaUsuarioUltAtualiz")=usuario
					rs("InstaladorInstalaDtHrUltAtualiz")=Now
					end if
				rs("pedido_bs_x_ac")=s_pedido_ac
				rs("pedido_bs_x_marketplace")=s_numero_mktplace
				rs("marketplace_codigo_origem")=s_origem_pedido
				rs("NFe_texto_constar")=s_nf_texto
				rs("NFe_xPed")=s_num_pedido_compra
				rs("loja_indicou")=s_loja_indicou
				rs("comissao_loja_indicou")=comissao_loja_indicou
				rs("venda_externa")=venda_externa
		
				rs("indicador") = c_indicador
		
				rs("GarantiaIndicadorStatus") = CLng(rb_garantia_indicador)
				rs("GarantiaIndicadorUsuarioUltAtualiz") = usuario
				rs("GarantiaIndicadorDtHrUltAtualiz") = Now

				if rb_end_entrega = "S" then
					rs("st_end_entrega") = 1
					rs("EndEtg_endereco") = EndEtg_endereco
					rs("EndEtg_endereco_numero") = EndEtg_endereco_numero
					rs("EndEtg_endereco_complemento") = EndEtg_endereco_complemento
					rs("EndEtg_bairro") = EndEtg_bairro
					rs("EndEtg_cidade") = EndEtg_cidade
					rs("EndEtg_uf") = EndEtg_uf
					rs("EndEtg_cep") = EndEtg_cep
					rs("EndEtg_cod_justificativa") = EndEtg_obs
					if blnUsarMemorizacaoCompletaEnderecos then
						rs("EndEtg_email") = EndEtg_email
						rs("EndEtg_email_xml") = EndEtg_email_xml
						rs("EndEtg_nome") = EndEtg_nome
						rs("EndEtg_ddd_res") = EndEtg_ddd_res
						rs("EndEtg_tel_res") = EndEtg_tel_res
						rs("EndEtg_ddd_com") = EndEtg_ddd_com
						rs("EndEtg_tel_com") = EndEtg_tel_com
						rs("EndEtg_ramal_com") = EndEtg_ramal_com
						rs("EndEtg_ddd_cel") = EndEtg_ddd_cel
						rs("EndEtg_tel_cel") = EndEtg_tel_cel
						rs("EndEtg_ddd_com_2") = EndEtg_ddd_com_2
						rs("EndEtg_tel_com_2") = EndEtg_tel_com_2
						rs("EndEtg_ramal_com_2") = EndEtg_ramal_com_2
						rs("EndEtg_tipo_pessoa") = EndEtg_tipo_pessoa
						rs("EndEtg_cnpj_cpf") = retorna_so_digitos(EndEtg_cnpj_cpf)
						rs("EndEtg_contribuinte_icms_status") = converte_numero(EndEtg_contribuinte_icms_status)
						rs("EndEtg_produtor_rural_status") = converte_numero(EndEtg_produtor_rural_status)
						rs("EndEtg_ie") = EndEtg_ie
						rs("EndEtg_rg") = EndEtg_rg
						end if
					end if
		
				'OBTEN��O DE TRANSPORTADORA QUE ATENDA AO CEP INFORMADO, SE HOUVER
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
		
				'01/02/2018: os pedidos do Arclube usam o RA para incluir o valor do frete e, portanto, n�o devem ter des�gio do RA
				if (Cstr(loja) <> Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE)) And (Not blnMagentoPedidoComIndicador) then rs("perc_desagio_RA_liquida") = PERC_DESAGIO_RA_LIQUIDA

				if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) And blnMagentoPedidoComIndicador then
					rs("magento_installer_commission_value") = percCommissionValue
					rs("magento_installer_commission_discount") = percCommissionDiscount
					rs("magento_shipping_amount") = vlMagentoShippingAmount
					end if

				rs("permite_RA_status") = permite_RA_status
		
				if permite_RA_status = 1 then
					rs("opcao_possui_RA") = rb_RA
				else
					rs("opcao_possui_RA") = "-" ' N�o se aplica
					end if
		
				rs("endereco_memorizado_status") = 1
				rs("endereco_logradouro") = Trim("" & t_CLIENTE("endereco"))
				rs("endereco_bairro") = Trim("" & t_CLIENTE("bairro"))
				rs("endereco_cidade") = Trim("" & t_CLIENTE("cidade"))
				rs("endereco_uf") = Trim("" & t_CLIENTE("uf"))
				rs("endereco_cep") = Trim("" & t_CLIENTE("cep"))
				rs("endereco_numero") = Trim("" & t_CLIENTE("endereco_numero"))
				rs("endereco_complemento") = Trim("" & t_CLIENTE("endereco_complemento"))

				if blnUsarMemorizacaoCompletaEnderecos then
					rs("st_memorizacao_completa_enderecos") = 1
					rs("endereco_email") = Trim("" & t_CLIENTE("email"))
					rs("endereco_email_xml") = Trim("" & t_CLIENTE("email_xml"))
					rs("endereco_nome") = Trim("" & t_CLIENTE("nome"))
					rs("endereco_ddd_res") = Trim("" & t_CLIENTE("ddd_res"))
					rs("endereco_tel_res") = Trim("" & t_CLIENTE("tel_res"))
					rs("endereco_ddd_com") = Trim("" & t_CLIENTE("ddd_com"))
					rs("endereco_tel_com") = Trim("" & t_CLIENTE("tel_com"))
					rs("endereco_ramal_com") = Trim("" & t_CLIENTE("ramal_com"))
					rs("endereco_ddd_cel") = Trim("" & t_CLIENTE("ddd_cel"))
					rs("endereco_tel_cel") = Trim("" & t_CLIENTE("tel_cel"))
					rs("endereco_ddd_com_2") = Trim("" & t_CLIENTE("ddd_com_2"))
					rs("endereco_tel_com_2") = Trim("" & t_CLIENTE("tel_com_2"))
					rs("endereco_ramal_com_2") = Trim("" & t_CLIENTE("ramal_com_2"))
					rs("endereco_tipo_pessoa") = Trim("" & t_CLIENTE("tipo"))
					rs("endereco_cnpj_cpf") = Trim("" & t_CLIENTE("cnpj_cpf"))
					rs("endereco_contribuinte_icms_status") = t_CLIENTE("contribuinte_icms_status")
					rs("endereco_produtor_rural_status") = t_CLIENTE("produtor_rural_status")
					rs("endereco_ie") = Trim("" & t_CLIENTE("ie"))
					rs("endereco_rg") = Trim("" & t_CLIENTE("rg"))
					end if

				if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) OR ( (Cstr(loja) = Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE)) And (Trim(s_pedido_ac) <> "") ) then
					rs("plataforma_origem_pedido") = COD_PLATAFORMA_ORIGEM_PEDIDO__MAGENTO
				else
					rs("plataforma_origem_pedido") = COD_PLATAFORMA_ORIGEM_PEDIDO__ERP
					end if

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
													" Qtde Sem Presen�a Autorizada = " & Cstr(qtde_spe) & "," & _
													" Qtde Estoque Vendido = " & Cstr(qtde_estoque_vendido_aux) & "," & _
													" Qtde Sem Presen�a = " & Cstr(qtde_estoque_sem_presenca_aux)
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
		
				s="UPDATE t_ESTOQUE_LOG SET pedido_estoque_origem='" & id_pedido & "' WHERE pedido_estoque_origem='" & id_pedido_temp & "'"
				cn.Execute(s)

				s="UPDATE t_ESTOQUE_LOG SET pedido_estoque_destino='" & id_pedido & "' WHERE pedido_estoque_destino='" & id_pedido_temp & "'"
				cn.Execute(s)
		
				if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
					s="UPDATE t_MAGENTO_API_PEDIDO_XML SET st_usado_cadastramento_pedido_erp = 1, dt_hr_usado_cadastramento_pedido_erp = getdate(), pedido_erp = '" & id_pedido_base & "' WHERE (id = " & id_magento_api_pedido_xml & ")"
					cn.Execute(s)
					end if

				if indice_pedido = 1 then
				'	INDICADOR: SE ESTE PEDIDO � COM INDICADOR E O CLIENTE AINDA N�O TEM UM INDICADOR NO CADASTRO, ENT�O CADASTRA ESTE.
					if rb_indicacao = "S" then
						if indicador_original = "" then
							s="UPDATE t_CLIENTE SET indicador='" & c_indicador & "' WHERE (id='" & cliente_selecionado & "')"
							cn.Execute(s)
							s_log_cliente_indicador = "Cadastrado o indicador '" & c_indicador & "' no cliente id=" & cliente_selecionado
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
			'		SENHAS DE AUTORIZA��O PARA DESCONTO SUPERIOR
					for k = Lbound(v_desconto) to Ubound(v_desconto)
						if Trim(v_desconto(k)) <> "" then
							s = "SELECT * FROM t_DESCONTO" & _
								" WHERE (usado_status=0)" & _
								" AND (cancelado_status=0)" & _
								" AND (id='" & Trim(v_desconto(k)) & "')"
							if rs.State <> 0 then rs.Close
							rs.open s, cn
							if rs.Eof then
								alerta = "Senha de autoriza��o para desconto superior n�o encontrado."
								exit for
							else
								rs("usado_status") = 1
								rs("usado_data") = Now
								if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) And blnMagentoPedidoComIndicador then
									rs("vendedor") = sIdVendedor
								else
									rs("vendedor") = usuario
									end if
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
					'	VERIFICA SE O ENDERE�O J� FOI USADO ANTERIORMENTE POR OUTRO CLIENTE (POSS�VEL FRAUDE)
					'	ENDERE�O DO CADASTRO
					'	====================
					'	1) VERIFICA SE O ENDERE�O USADO � O DO PARCEIRO
						if c_indicador <> "" then
							if isEnderecoIgual(Trim("" & t_CLIENTE("endereco")), Trim("" & t_CLIENTE("endereco_numero")), Trim("" & t_CLIENTE("cep")), r_orcamentista_e_indicador.endereco, r_orcamentista_e_indicador.endereco_numero, r_orcamentista_e_indicador.cep) then
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
									rs("id_cliente") = cliente_selecionado
									rs("tipo_endereco") = COD_PEDIDO_AN_ENDERECO__CAD_CLIENTE
									rs("endereco_logradouro") = Trim("" & t_CLIENTE("endereco"))
									rs("endereco_bairro") = Trim("" & t_CLIENTE("bairro"))
									rs("endereco_cidade") = Trim("" & t_CLIENTE("cidade"))
									rs("endereco_uf") = Trim("" & t_CLIENTE("uf"))
									rs("endereco_cep") = Trim("" & t_CLIENTE("cep"))
									rs("endereco_numero") = Trim("" & t_CLIENTE("endereco_numero"))
									rs("endereco_complemento") = Trim("" & t_CLIENTE("endereco_complemento"))
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
							end if ' if c_indicador <> ""
						end if ' if alerta = ""
		
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
											" AND (c.id <> '" & cliente_selecionado & "')" & _
											" AND (c.cep = '" & retorna_so_digitos(Trim("" & t_CLIENTE("cep"))) & "')" & _
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
											" AND (id_cliente <> '" & cliente_selecionado & "')" & _
											" AND (endereco_cep = '" & retorna_so_digitos(Trim("" & t_CLIENTE("cep"))) & "')" & _
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
											" AND (id_cliente <> '" & cliente_selecionado & "')" & _
											" AND (EndEtg_cep = '" & retorna_so_digitos(Trim("" & t_CLIENTE("cep"))) & "')" & _
									") t" & _
								" ORDER BY" & _
									" data_hora DESC"
							if rs.State <> 0 then rs.Close
							rs.Open s, cn
							do while Not rs.Eof
								if isEnderecoIgual(Trim("" & t_CLIENTE("endereco")), Trim("" & t_CLIENTE("endereco_numero")), Trim("" & t_CLIENTE("cep")), Trim("" & rs("endereco_logradouro")), Trim("" & rs("endereco_numero")), Trim("" & rs("endereco_cep"))) then
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
									'	J� GRAVOU O REGISTRO PAI?
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
											rs("id_cliente") = cliente_selecionado
											rs("tipo_endereco") = COD_PEDIDO_AN_ENDERECO__CAD_CLIENTE
											rs("endereco_logradouro") = Trim("" & t_CLIENTE("endereco"))
											rs("endereco_bairro") = Trim("" & t_CLIENTE("bairro"))
											rs("endereco_cidade") = Trim("" & t_CLIENTE("cidade"))
											rs("endereco_uf") = Trim("" & t_CLIENTE("uf"))
											rs("endereco_cep") = Trim("" & t_CLIENTE("cep"))
											rs("endereco_numero") = Trim("" & t_CLIENTE("endereco_numero"))
											rs("endereco_complemento") = Trim("" & t_CLIENTE("endereco_complemento"))
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
						if rb_end_entrega = "S" then 
						'	ENDERE�O DE ENTREGA (SE HOUVER)
						'	===============================
						'	1) VERIFICA SE O ENDERE�O USADO � O DO PARCEIRO
							if c_indicador <> "" then
								if isEnderecoIgual(EndEtg_endereco, EndEtg_endereco_numero, EndEtg_cep, r_orcamentista_e_indicador.endereco, r_orcamentista_e_indicador.endereco_numero, r_orcamentista_e_indicador.cep) then
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
										rs("id_cliente") = cliente_selecionado
										rs("tipo_endereco") = COD_PEDIDO_AN_ENDERECO__END_ENTREGA
										rs("endereco_logradouro") = EndEtg_endereco
										rs("endereco_bairro") = EndEtg_bairro
										rs("endereco_cidade") = EndEtg_cidade
										rs("endereco_uf") = EndEtg_uf
										rs("endereco_cep") = EndEtg_cep
										rs("endereco_numero") = EndEtg_endereco_numero
										rs("endereco_complemento") = EndEtg_endereco_complemento
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
								end if ' if c_indicador <> ""
				
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
													" AND (c.id <> '" & cliente_selecionado & "')" & _
													" AND (c.cep = '" & retorna_so_digitos(EndEtg_cep) & "')" & _
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
													" AND (id_cliente <> '" & cliente_selecionado & "')" & _
													" AND (endereco_cep = '" & retorna_so_digitos(EndEtg_cep) & "')" & _
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
													" AND (id_cliente <> '" & cliente_selecionado & "')" & _
													" AND (EndEtg_cep = '" & retorna_so_digitos(EndEtg_cep) & "')" & _
											") t" & _
										" ORDER BY" & _
											" data_hora DESC"
									if rs.State <> 0 then rs.Close
									rs.Open s, cn
									do while Not rs.Eof
										if isEnderecoIgual(EndEtg_endereco, EndEtg_endereco_numero, EndEtg_cep, Trim("" & rs("endereco_logradouro")), Trim("" & rs("endereco_numero")), Trim("" & rs("endereco_cep"))) then
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
											'	J� GRAVOU O REGISTRO PAI?
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
													rs("id_cliente") = cliente_selecionado
													rs("tipo_endereco") = COD_PEDIDO_AN_ENDERECO__END_ENTREGA
													rs("endereco_logradouro") = EndEtg_endereco
													rs("endereco_bairro") = EndEtg_bairro
													rs("endereco_cidade") = EndEtg_cidade
													rs("endereco_uf") = EndEtg_uf
													rs("endereco_cep") = EndEtg_cep
													rs("endereco_numero") = EndEtg_endereco_numero
													rs("endereco_complemento") = EndEtg_endereco_complemento
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
							end if 'if rb_end_entrega = "S"
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
				s_log = "vl total=" & formata_moeda(vl_total)
				s_log = s_log & "; RA=" & formata_texto_log(rb_RA)
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
				if Trim("" & rs("NFe_texto_constar"))<>"" then s_log = s_log & "; NFe_texto_constar=" & formata_texto_log(rs("NFe_texto_constar"))
				IF Trim("" & rs("NFe_xPed"))<>"" then s_log = s_log & "; NFe_xPed=" & formata_texto_log(rs("NFe_xPed"))        
				if Trim("" & rs("obs_2"))<>"" then s_log = s_log & "; obs_2=" & formata_texto_log(rs("obs_2"))
				if Trim("" & rs("pedido_bs_x_ac"))<>"" then s_log = s_log & "; pedido_bs_x_ac=" & formata_texto_log(rs("pedido_bs_x_ac"))
				if Trim("" & rs("pedido_bs_x_marketplace"))<>"" then s_log = s_log & "; pedido_bs_x_marketplace=" & formata_texto_log(rs("pedido_bs_x_marketplace"))
				if Trim("" & rs("marketplace_codigo_origem"))<>"" then s_log = s_log & "; marketplace_codigo_origem=" & formata_texto_log(rs("marketplace_codigo_origem"))
				if Trim("" & rs("loja_indicou"))<>"" then
					s_log = s_log & "; loja_indicou=" & formata_texto_log(rs("loja_indicou"))
					s_log = s_log & "; comissao_loja_indicou=" & formata_perc_comissao(rs("comissao_loja_indicou")) & "%"
					end if
				if Cstr(rs("analise_credito"))=Cstr(COD_AN_CREDITO_OK) then
					s_log = s_log & "; an�lise cr�dito OK (<=" & formata_moeda(vl_aprov_auto_analise_credito) & ")"
				else
					s_log = s_log & "; status da an�lise cr�dito: " & Cstr(rs("analise_credito")) & " - " & descricao_analise_credito(Cstr(rs("analise_credito")))
					end if
			'	Forma de Pagamento (nova vers�o)
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
		
				s_log = s_log & "; Endere�o cobran�a=" & formata_endereco(Trim("" & t_CLIENTE("endereco")), Trim("" & t_CLIENTE("endereco_numero")), Trim("" & t_CLIENTE("endereco_complemento")), Trim("" & t_CLIENTE("bairro")), Trim("" & t_CLIENTE("cidade")), Trim("" & t_CLIENTE("uf")), Trim("" & t_CLIENTE("cep")))
				if blnUsarMemorizacaoCompletaEnderecos then
					s_log = s_log & _
							" (" & _
							"email=" & Trim("" & t_CLIENTE("email")) & _
							", email_xml=" & Trim("" & t_CLIENTE("email_xml")) & _
							", nome=" & Trim("" & t_CLIENTE("nome")) & _
							", ddd_res=" & Trim("" & t_CLIENTE("ddd_res")) & _
							", tel_res=" & Trim("" & t_CLIENTE("tel_res")) & _
							", ddd_com=" & Trim("" & t_CLIENTE("ddd_com")) & _
							", tel_com=" & Trim("" & t_CLIENTE("tel_com")) & _
							", ramal_com=" & Trim("" & t_CLIENTE("ramal_com")) & _
							", ddd_cel=" & Trim("" & t_CLIENTE("ddd_cel")) & _
							", tel_cel=" & Trim("" & t_CLIENTE("tel_cel")) & _
							", ddd_com_2=" & Trim("" & t_CLIENTE("ddd_com_2")) & _
							", tel_com_2=" & Trim("" & t_CLIENTE("tel_com_2")) & _
							", ramal_com_2=" & Trim("" & t_CLIENTE("ramal_com_2")) & _
							", tipo_pessoa=" & Trim("" & t_CLIENTE("tipo")) & _
							", cnpj_cpf=" & Trim("" & t_CLIENTE("cnpj_cpf")) & _
							", contribuinte_icms_status=" & t_CLIENTE("contribuinte_icms_status") & _
							", produtor_rural_status=" & t_CLIENTE("produtor_rural_status") & _
							", ie=" & Trim("" & t_CLIENTE("ie")) & _
							", rg=" & Trim("" & t_CLIENTE("rg")) & _
							")"
					end if

				if rb_end_entrega = "S" then
					s_log = s_log & "; Endere�o entrega=" & formata_endereco(EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep) & " [EndEtg_cod_justificativa=" & EndEtg_obs & "]"
					if blnUsarMemorizacaoCompletaEnderecos then
						s_log = s_log & _
								" (" & _
								"email=" & EndEtg_email & _
								", email_xml=" & EndEtg_email_xml & _
								", nome=" & EndEtg_nome & _
								", ddd_res=" & EndEtg_ddd_res & _
								", tel_res=" & EndEtg_tel_res & _
								", ddd_com=" & EndEtg_ddd_com & _
								", tel_com=" & EndEtg_tel_com & _
								", ramal_com=" & EndEtg_ramal_com & _
								", ddd_cel=" & EndEtg_ddd_cel & _
								", tel_cel=" & EndEtg_tel_cel & _
								", ddd_com_2=" & EndEtg_ddd_com_2 & _
								", tel_com_2=" & EndEtg_tel_com_2 & _
								", ramal_com_2=" & EndEtg_ramal_com_2 & _
								", tipo_pessoa=" & EndEtg_tipo_pessoa & _
								", cnpj_cpf=" & EndEtg_cnpj_cpf & _
								", contribuinte_icms_status=" & EndEtg_contribuinte_icms_status & _
								", produtor_rural_status=" & EndEtg_produtor_rural_status & _
								", ie=" & EndEtg_ie & _
								", rg=" & EndEtg_rg & _
								")"
						end if
				else
					s_log = s_log & "; Endere�o entrega=mesmo do cadastro"
					end if
		
				if sTranspSelAutoTransportadoraId = "" then
					s_log = s_log & "; Escolha autom�tica de transportadora=N"
				else
					s_log = s_log & "; Escolha autom�tica de transportadora=S"
					s_log = s_log & "; Transportadora=" & sTranspSelAutoTransportadoraId
					s_log = s_log & "; CEP relacionado=" & cep_formata(sTranspSelAutoCep)
					end if
		
				s_log = s_log & "; GarantiaIndicadorStatus=" & rb_garantia_indicador
				s_log = s_log & "; perc_desagio_RA_liquida=" & rs("perc_desagio_RA_liquida")
				s_log = s_log & "; pedido_bs_x_at=" & c_ped_bonshop

				if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
					if Trim("" & rs("pedido_bs_x_marketplace"))<>"" then s_log = s_log & "; numero_pedido_marketplace=" & s_numero_mktplace
					s_log = s_log & "; cod_origem_pedido=" & s_origem_pedido
					end if

				if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
					if s_log <> "" then s_log = s_log & ";"
					s_log = s_log & " Opera��o de origem: cadastramento semi-autom�tico de pedido do e-commerce (n� Magento=" & c_numero_magento & ", t_MAGENTO_API_PEDIDO_XML.id=" & id_magento_api_pedido_xml & ")"
					end if
				end if ' if Not rs.Eof

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
						s_log = s_log & "Detalhes do auto-split: Modo de sele��o do CD = " & rb_selecao_cd
						if rb_selecao_cd = MODO_SELECAO_CD__MANUAL then s_log = s_log & "; id_nfe_emitente = " & c_id_nfe_emitente_selecao_manual
						s_log = s_log & chr(13)
						blnAchou = True
						end if
					s_log = s_log & vLogAutoSplit(i)
					end if
				next

			if s_log <> "" then
				grava_log usuario, loja, id_pedido, cliente_selecionado, OP_LOG_PEDIDO_NOVO, s_log
				end if
			end if

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
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<BR><BR>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<% 	if erro_produto_indisponivel then 
		'	VOLTA PARA A TELA QUE CADASTRA A QUANTIDADE DE PRODUTOS
			s="javascript:history.go(-2)"
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
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
	if Not (t_CLIENTE is nothing) then
		if t_CLIENTE.State <> 0 then t_CLIENTE.Close
		set t_CLIENTE = nothing
		end if
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>