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

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim alerta, alerta_aux
	alerta=""
	
	cliente_selecionado = Trim(request("cliente_selecionado"))
	if (cliente_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2, t_CLIENTE, tMAP_XML, tMAP_ITEM, tITEM_SVC, tOI
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim r_cliente
	set r_cliente = New cl_CLIENTE
	call x_cliente_bd(cliente_selecionado, r_cliente)
	tipo_cliente = r_cliente.tipo

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim sBlocoNotasEndCob, sBlocoNotasEndEtg, sBlocoNotasMsg
	sBlocoNotasEndCob = ""
	sBlocoNotasEndEtg = ""
	sBlocoNotasMsg = ""

	dim insert_request_guid
	insert_request_guid = Trim(Request.Form("insert_request_guid"))

	dim s_mktp_payment
	dim percDescServico, vl_servico_original_price, vl_servico_price
	dim operacao_origem, c_numero_magento, operationControlTicket, sessionToken, id_magento_api_pedido_xml
	dim c_vl_total_produto_magento, vl_total_produto_magento
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
		c_vl_total_produto_magento = Trim(Request.Form("c_vl_total_produto_magento"))
		vl_total_produto_magento = converte_numero(c_vl_total_produto_magento)
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

'	LOCALIZA DADOS DO CLIENTE
	midia_selecionada = ""
	indicador_original = ""
	if cliente_selecionado <> "" then
		s = "SELECT * FROM t_CLIENTE WHERE (id='" & cliente_selecionado & "')"
		set t_CLIENTE = cn.execute(s)
		if Not t_CLIENTE.Eof then
			midia_selecionada = Trim("" & t_CLIENTE("midia"))
			indicador_original = Trim("" & t_CLIENTE("indicador"))
			end if
		end if

	dim c_FlagCadSemiAutoPedMagento_FluxoOtimizado
	c_FlagCadSemiAutoPedMagento_FluxoOtimizado = Trim(Request.Form("c_FlagCadSemiAutoPedMagento_FluxoOtimizado"))

	dim EndCob_endereco, EndCob_endereco_numero, EndCob_endereco_complemento, EndCob_endereco_ponto_referencia, EndCob_bairro, EndCob_cidade, EndCob_uf, EndCob_cep
	dim EndCob_email, EndCob_email_xml, EndCob_nome, EndCob_tipo_pessoa
	dim EndCob_ddd_res, EndCob_tel_res, EndCob_ddd_com, EndCob_tel_com, EndCob_ramal_com, EndCob_ddd_com_2, EndCob_tel_com_2, EndCob_ramal_com_2, EndCob_ddd_cel, EndCob_tel_cel
	dim EndCob_cnpj_cpf, EndCob_contribuinte_icms_status, EndCob_produtor_rural_status, EndCob_ie, EndCob_rg, EndCob_contato
	dim rb_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_endereco_ponto_referencia
	dim EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs
	dim EndEtg_email, EndEtg_email_xml, EndEtg_nome, EndEtg_ddd_res, EndEtg_tel_res, EndEtg_ddd_com, EndEtg_tel_com, EndEtg_ramal_com
	dim EndEtg_ddd_cel, EndEtg_tel_cel, EndEtg_ddd_com_2, EndEtg_tel_com_2, EndEtg_ramal_com_2
	dim EndEtg_tipo_pessoa, EndEtg_cnpj_cpf, EndEtg_contribuinte_icms_status, EndEtg_produtor_rural_status
	dim EndEtg_ie, EndEtg_rg

	if c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "1" then
		EndCob_endereco = Trim(Request.Form("EndCob_endereco"))
		EndCob_endereco_numero = Trim(Request.Form("EndCob_endereco_numero"))
		EndCob_endereco_complemento = Trim(Request.Form("EndCob_endereco_complemento"))
		EndCob_endereco_ponto_referencia = Trim(Request.Form("EndCob_endereco_ponto_referencia"))
		EndCob_bairro = Trim(Request.Form("EndCob_bairro"))
		EndCob_cidade = Trim(Request.Form("EndCob_cidade"))
		EndCob_uf = Trim(Request.Form("EndCob_uf"))
		EndCob_cep = Trim(Request.Form("EndCob_cep"))
		EndCob_email = Trim(Request.Form("EndCob_email"))
		EndCob_email_xml = Trim(Request.Form("EndCob_email_xml"))
		EndCob_nome = Trim(Request.Form("EndCob_nome"))
		EndCob_tipo_pessoa = Trim(Request.Form("EndCob_tipo_pessoa"))
		EndCob_ddd_res = Trim(Request.Form("EndCob_ddd_res"))
		EndCob_tel_res = Trim(Request.Form("EndCob_tel_res"))
		EndCob_ddd_com = Trim(Request.Form("EndCob_ddd_com"))
		EndCob_tel_com = Trim(Request.Form("EndCob_tel_com"))
		EndCob_ramal_com = Trim(Request.Form("EndCob_ramal_com"))
		EndCob_ddd_com_2 = Trim(Request.Form("EndCob_ddd_com_2"))
		EndCob_tel_com_2 = Trim(Request.Form("EndCob_tel_com_2"))
		EndCob_ramal_com_2 = Trim(Request.Form("EndCob_ramal_com_2"))
		EndCob_ddd_cel = Trim(Request.Form("EndCob_ddd_cel"))
		EndCob_tel_cel = Trim(Request.Form("EndCob_tel_cel"))
		EndCob_cnpj_cpf = Trim(Request.Form("EndCob_cnpj_cpf"))
		EndCob_contribuinte_icms_status = Trim(Request.Form("EndCob_contribuinte_icms_status"))
		EndCob_produtor_rural_status = Trim(Request.Form("EndCob_produtor_rural_status"))
		EndCob_ie = Trim(Request.Form("EndCob_ie"))
		EndCob_rg = Trim(Request.Form("EndCob_rg"))
		EndCob_contato = Trim("" & t_CLIENTE("contato"))
	else
		EndCob_endereco = Trim("" & t_CLIENTE("endereco"))
		EndCob_endereco_numero = Trim("" & t_CLIENTE("endereco_numero"))
		EndCob_endereco_complemento = Trim("" & t_CLIENTE("endereco_complemento"))
		EndCob_endereco_ponto_referencia = ""
		EndCob_bairro = Trim("" & t_CLIENTE("bairro"))
		EndCob_cidade = Trim("" & t_CLIENTE("cidade"))
		EndCob_uf = Trim("" & t_CLIENTE("uf"))
		EndCob_cep = Trim("" & t_CLIENTE("cep"))
		EndCob_email = Trim("" & t_CLIENTE("email"))
		EndCob_email_xml = Trim("" & t_CLIENTE("email_xml"))
		EndCob_nome = Trim("" & t_CLIENTE("nome"))
		EndCob_tipo_pessoa = Trim("" & t_CLIENTE("tipo"))
		EndCob_ddd_res = Trim("" & t_CLIENTE("ddd_res"))
		EndCob_tel_res = Trim("" & t_CLIENTE("tel_res"))
		EndCob_ddd_com = Trim("" & t_CLIENTE("ddd_com"))
		EndCob_tel_com = Trim("" & t_CLIENTE("tel_com"))
		EndCob_ramal_com = Trim("" & t_CLIENTE("ramal_com"))
		EndCob_ddd_com_2 = Trim("" & t_CLIENTE("ddd_com_2"))
		EndCob_tel_com_2 = Trim("" & t_CLIENTE("tel_com_2"))
		EndCob_ramal_com_2 = Trim("" & t_CLIENTE("ramal_com_2"))
		EndCob_ddd_cel = Trim("" & t_CLIENTE("ddd_cel"))
		EndCob_tel_cel = Trim("" & t_CLIENTE("tel_cel"))
		EndCob_cnpj_cpf = Trim("" & t_CLIENTE("cnpj_cpf"))
		EndCob_contribuinte_icms_status = t_CLIENTE("contribuinte_icms_status")
		EndCob_produtor_rural_status = t_CLIENTE("produtor_rural_status")
		EndCob_ie = Trim("" & t_CLIENTE("ie"))
		EndCob_rg = Trim("" & t_CLIENTE("rg"))
		EndCob_contato = Trim("" & t_CLIENTE("contato"))
		end if

	rb_end_entrega = Trim(Request.Form("rb_end_entrega"))
	EndEtg_endereco = Trim(Request.Form("EndEtg_endereco"))
	EndEtg_endereco_numero = Trim(Request.Form("EndEtg_endereco_numero"))
	EndEtg_endereco_complemento = Trim(Request.Form("EndEtg_endereco_complemento"))
	EndEtg_endereco_ponto_referencia = Trim(Request.Form("EndEtg_endereco_ponto_referencia"))
	EndEtg_bairro = Trim(Request.Form("EndEtg_bairro"))
	EndEtg_cidade = Trim(Request.Form("EndEtg_cidade"))
	EndEtg_uf = Trim(Request.Form("EndEtg_uf"))
	EndEtg_cep = retorna_so_digitos(Trim(Request.Form("EndEtg_cep")))
    EndEtg_obs = Trim(Request.Form("EndEtg_obs"))
	EndEtg_email = Trim(Request.Form("EndEtg_email"))
	EndEtg_email_xml = Trim(Request.Form("EndEtg_email_xml"))
	EndEtg_nome = Trim(Request.Form("EndEtg_nome"))
	EndEtg_ddd_res = retorna_so_digitos(Trim(Request.Form("EndEtg_ddd_res")))
	EndEtg_tel_res = retorna_so_digitos(Trim(Request.Form("EndEtg_tel_res")))
	EndEtg_ddd_com = retorna_so_digitos(Trim(Request.Form("EndEtg_ddd_com")))
	EndEtg_tel_com = retorna_so_digitos(Trim(Request.Form("EndEtg_tel_com")))
	EndEtg_ramal_com = retorna_so_digitos(Trim(Request.Form("EndEtg_ramal_com")))
	EndEtg_ddd_cel = retorna_so_digitos(Trim(Request.Form("EndEtg_ddd_cel")))
	EndEtg_tel_cel = retorna_so_digitos(Trim(Request.Form("EndEtg_tel_cel")))
	EndEtg_ddd_com_2 = retorna_so_digitos(Trim(Request.Form("EndEtg_ddd_com_2")))
	EndEtg_tel_com_2 = retorna_so_digitos(Trim(Request.Form("EndEtg_tel_com_2")))
	EndEtg_ramal_com_2 = retorna_so_digitos(Trim(Request.Form("EndEtg_ramal_com_2")))
	EndEtg_tipo_pessoa = Trim(Request.Form("EndEtg_tipo_pessoa"))
	EndEtg_cnpj_cpf = Trim(Request.Form("EndEtg_cnpj_cpf"))
	EndEtg_contribuinte_icms_status = Trim(Request.Form("EndEtg_contribuinte_icms_status"))
	EndEtg_produtor_rural_status = Trim(Request.Form("EndEtg_produtor_rural_status"))
	EndEtg_ie = Trim(Request.Form("EndEtg_ie"))
	EndEtg_rg = Trim(Request.Form("EndEtg_rg"))

	dim vl_aprov_auto_analise_credito
	vl_aprov_auto_analise_credito = 0

	dim vl_total_RA_liquido
	dim s, c, i, iv, j, k, n, opcao_venda_sem_estoque, vl_total_preco_lista, vl_total, vl_total_NF, vl_total_RA, qtde_estoque_total_disponivel, blnAchou, blnDesativado
	dim v_desconto()
	ReDim v_desconto(0)
	v_desconto(UBound(v_desconto)) = ""

	opcao_venda_sem_estoque = Trim(request("opcao_venda_sem_estoque"))
	
	dim s_forma_pagto, s_obs1, s_obs2, s_recebido, s_etg_imediata, s_bem_uso_consumo, s_pedido_ac, s_numero_mktplace, s_origem_pedido, c_data_previsao_entrega
    dim s_nf_texto, s_num_pedido_compra
	s_obs1=Trim(request("c_obs1"))
	s_obs2=Trim(request("c_obs2"))
    s_pedido_ac=Trim(request("c_pedido_ac"))
    s_numero_mktplace = Trim(Request("c_numero_mktplace"))
    s_origem_pedido = Trim(Request("c_origem_pedido"))
	s_recebido=Trim(request("rb_recebido"))
	s_etg_imediata=Trim(request("rb_etg_imediata"))
	c_data_previsao_entrega = Trim(Request("c_data_previsao_entrega"))
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
	
'	PAR�METROS DE COMISS�O E DESCONTO
	dim rCD
	set rCD = obtem_perc_max_comissao_e_desconto_por_loja(loja)

	dim r_loja
	set r_loja = New cl_LOJA
	call x_loja_bd(loja, r_loja)

	dim r_usuario
	if alerta = "" then
		call le_usuario(usuario, r_usuario, msg_erro)
		end if

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
	
	vlMagentoShippingAmount = 0
	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
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
				vlMagentoShippingAmount = converte_numero(tMAP_XML("shipping_amount")) - converte_numero(tMAP_XML("shipping_discount_amount"))
				end if
			end if 'if alerta = ""
		end if 'if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO

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
	
	if alerta = "" then
		if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) then
			if blnMagentoPedidoComIndicador then permite_RA_status = 1
		
			if (Trim("" & tMAP_XML("b2b_type_order")) = COD_MAGENTO_TYPE_ORDER__INSTALLER) And (Trim("" & tMAP_XML("magento_api_versao")) = CStr(VERSAO_API_MAGENTO_V2_REST_JSON)) then
				if vlMagentoShippingAmount > 0 then permite_RA_status = 1
				end if
			end if 'if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO)
		end if 'if alerta = ""

'	RA L�quido
	dim perc_desagio_RA, perc_limite_RA_sem_desagio, perc_desagio_RA_liquida
	dim vl_limite_mensal, vl_limite_mensal_consumido, vl_limite_mensal_disponivel

	if rb_indicacao = "S" then
		perc_desagio_RA = obtem_perc_desagio_RA_do_indicador(c_indicador)
		perc_limite_RA_sem_desagio = obtem_perc_limite_RA_sem_desagio()
		vl_limite_mensal = obtem_limite_mensal_compras_do_indicador(c_indicador)
		vl_limite_mensal_consumido = calcula_limite_mensal_consumido_do_indicador(c_indicador, Date)
		vl_limite_mensal_disponivel = vl_limite_mensal - vl_limite_mensal_consumido
		'01/02/2018: os pedidos do Arclube usam o RA para incluir o valor do frete e, portanto, n�o devem ter des�gio do RA
		if (Cstr(loja) <> Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE)) And (Not blnMagentoPedidoComIndicador) then
			perc_desagio_RA_liquida = getParametroPercDesagioRALiquida
		else
			perc_desagio_RA_liquida = 0
			end if
	else
		perc_desagio_RA = 0
		perc_limite_RA_sem_desagio = 0
		vl_limite_mensal = 0
		vl_limite_mensal_consumido = 0
		perc_desagio_RA_liquida = 0
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
				s=Trim(Request.Form("c_preco_lista")(i))
				.preco_lista=converte_numero(s)
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
'	TRATAMENTO P/ OS CASOS EM QUE: USU�RIO EST� TENTANDO USAR O BOT�O VOLTAR, OCORREU DUPLO CLIQUE OU USU�RIO ATUALIZOU A P�GINA ENQUANTO AINDA ESTAVA PROCESSANDO (DUPLO ACIONAMENTO)
'	Esse tratamento est� sendo feito atrav�s do campo insert_request_guid (t_PEDIDO.InsertRequestGuid), mas se esse campo estiver vazio, continua realizando o controle pelo m�todo antigo
'	O tratamento usando o campo t_PEDIDO.InsertRequestGuid � feito dentro da transa��o para assegurar que n�o haver� problemas de acesso concorrente
	dim pedido_a, vjg
	if alerta = "" then
		if insert_request_guid = "" then
			if Cstr(loja) <> NUMERO_LOJA_OLD03 then
				s = "SELECT t_PEDIDO.pedido, fabricante, produto, qtde, preco_venda FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
					" WHERE (id_cliente='" & cliente_selecionado & "') AND (data=" & bd_formata_data(Date) & ")" & _
					" AND (loja='" & loja & "') AND (vendedor='" & usuario & "')" & _
					" AND (data >= " & bd_monta_data(Date) & ")" & _
					" AND (hora >= '" & formata_hora_hhnnss(Now-converte_min_to_dec(5))& "')" & _
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
				end if 'if Cstr(loja) <> NUMERO_LOJA_OLD03
			end if 'if insert_request_guid = ""
		end if 'if alerta = ""
	
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
		vl_total_preco_lista = 0
		vl_total = 0
		vl_total_NF = 0
		vl_total_RA = 0
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .produto <> "" then 
					vl_total_preco_lista = vl_total_preco_lista + (.qtde * .preco_lista)
					vl_total = vl_total + (.qtde * .preco_venda)
					vl_total_NF = vl_total_NF + (.qtde * .preco_NF)
					end if
				end with
			next
		vl_total_RA = vl_total_NF - vl_total
		end if
	
	dim desc_dado_medio
	if vl_total_preco_lista = 0 then
		desc_dado_medio = 0
	else
		desc_dado_medio = 100 * (vl_total_preco_lista - vl_total) / vl_total_preco_lista
		end if

	if alerta = "" then
		if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) then
			if (Trim("" & tMAP_XML("b2b_type_order")) = COD_MAGENTO_TYPE_ORDER__INSTALLER) And (Trim("" & tMAP_XML("magento_api_versao")) = CStr(VERSAO_API_MAGENTO_V2_REST_JSON)) then
				'Recalcula percentual de comiss�o p/ o caso do pre�o de venda ter sido editado
				if vl_total <> 0 then
					perc_RT = 100 * (converte_numero(tMAP_XML("b2b_installer_commission_value")) / vl_total)
				else
					perc_RT = 0
					end if
				end if
			end if 'if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO)
		end if 'if alerta = ""

'	ANALISA O PERCENTUAL DE COMISS�O+DESCONTO
	dim perc_max_RT_a_utilizar, perc_max_RT_padrao
	dim perc_comissao_e_desconto_a_utilizar, perc_comissao_e_desconto_padrao
	dim s_pg, blnPreferencial
	dim vlNivel1, vlNivel2
	perc_max_RT_padrao = rCD.perc_max_comissao
	perc_max_RT_a_utilizar = perc_max_RT_padrao
	if EndCob_tipo_pessoa = ID_PJ then
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
						if EndCob_tipo_pessoa = ID_PJ then
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
						if EndCob_tipo_pessoa = ID_PJ then
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
						if EndCob_tipo_pessoa = ID_PJ then
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
						if EndCob_tipo_pessoa = ID_PJ then
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
				if EndCob_tipo_pessoa = ID_PJ then
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
				if EndCob_tipo_pessoa = ID_PJ then
					perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
				else
					perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
					end if
				end if
			end if
		end if
	
	' Verifica se o usu�rio tem permiss�o de desconto por al�ada
	if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_1, s_lista_operacoes_permitidas) then
		if rCD.perc_max_comissao_alcada1 > perc_max_RT_a_utilizar then perc_max_RT_a_utilizar = rCD.perc_max_comissao_alcada1
		end if
	if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_2, s_lista_operacoes_permitidas) then
		if rCD.perc_max_comissao_alcada2 > perc_max_RT_a_utilizar then perc_max_RT_a_utilizar = rCD.perc_max_comissao_alcada2
		end if
	if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_3, s_lista_operacoes_permitidas) then
		if rCD.perc_max_comissao_alcada3 > perc_max_RT_a_utilizar then perc_max_RT_a_utilizar = rCD.perc_max_comissao_alcada3
		end if

	perc_comissao_e_desconto_padrao = perc_comissao_e_desconto_a_utilizar
	if tipo_cliente = ID_PF then
		if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_1, s_lista_operacoes_permitidas) then
			if rCD.perc_max_comissao_e_desconto_alcada1_pf > perc_comissao_e_desconto_a_utilizar then perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_alcada1_pf
			end if
		if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_2, s_lista_operacoes_permitidas) then
			if rCD.perc_max_comissao_e_desconto_alcada2_pf > perc_comissao_e_desconto_a_utilizar then perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_alcada2_pf
			end if
		if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_3, s_lista_operacoes_permitidas) then
			if rCD.perc_max_comissao_e_desconto_alcada3_pf > perc_comissao_e_desconto_a_utilizar then perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_alcada3_pf
			end if
	else
		if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_1, s_lista_operacoes_permitidas) then
			if rCD.perc_max_comissao_e_desconto_alcada1_pj > perc_comissao_e_desconto_a_utilizar then perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_alcada1_pj
			end if
		if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_2, s_lista_operacoes_permitidas) then
			if rCD.perc_max_comissao_e_desconto_alcada2_pj > perc_comissao_e_desconto_a_utilizar then perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_alcada2_pj
			end if
		if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_3, s_lista_operacoes_permitidas) then
			if rCD.perc_max_comissao_e_desconto_alcada3_pj > perc_comissao_e_desconto_a_utilizar then perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_alcada3_pj
			end if
		end if
	
	if alerta = "" then
		'Devido a arredondamentos no front, aceita margem de erro
		if (desc_dado_medio + perc_RT) > (perc_comissao_e_desconto_a_utilizar + MAX_MARGEM_ERRO_PERC_DESC_E_RT) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "A soma dos percentuais de comiss�o (" & formata_perc_RT(perc_RT) & "%) e de desconto m�dio do(s) produto(s) (" & formata_perc(desc_dado_medio) & "%) totaliza " & _
							formata_perc(perc_RT + desc_dado_medio) & "% e excede o m�ximo permitido!"
			end if
		
		if (Cstr(loja) <> Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE)) And (Not blnMagentoPedidoComIndicador) then
			if perc_RT > perc_max_RT_a_utilizar then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O percentual de comiss�o (" & formata_perc_RT(perc_RT) & "%) excede o m�ximo permitido!"
				end if
			end if
		end if 'if alerta = "" then

'	CONSIST�NCIA PARA VALOR ZERADO
	if alerta="" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .preco_venda <= 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto '" & .produto & "' est� com valor de venda zerado!"
				elseif ((rb_RA = "S") And (permite_RA_status = 1)) And (.preco_NF <= 0) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto '" & .produto & "' est� com pre�o zerado!"
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
                    .subgrupo = Trim("" & rs("subgrupo"))
					.peso = rs("peso")
					.qtde_volumes = rs("qtde_volumes")
					.markup_fabricante = rs("markup")
					.cubagem = rs("cubagem")
					.ncm = Trim("" & rs("ncm"))
					.cst = Trim("" & rs("cst"))
					.descontinuado = Trim("" & rs("descontinuado"))
					.cod_produto_xml_fabricante = Trim("" & rs("cod_produto_xml_fabricante"))
					.cod_produto_alfanum_fabricante = Trim("" & rs("cod_produto_alfanum_fabricante"))
					.potencia_valor = rs("potencia_valor")
					.id_unidade_potencia = rs("id_unidade_potencia")

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
					
					'Se houve edi��o no pre�o de venda, verifica se h� necessidade de registrar o ID do usu�rio que fez uso da al�ada
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'Soma do desconto e comiss�o est� abaixo do limite padr�o, portanto, assegura que os dados do uso do desconto por al�ada est�o vazios
					if (.desc_dado + perc_RT) <= (perc_comissao_e_desconto_padrao + MAX_MARGEM_ERRO_PERC_DESC_E_RT) then
						.StatusDescontoSuperior = 0
						.IdUsuarioDescontoSuperior = 0
						.DataHoraDescontoSuperior = Null
					else
						'Soma do desconto e comiss�o excede limite padr�o
						'Registra o uso da al�ada somente se o usu�rio possuir essa permiss�o, ou seja, se o limite m�ximo que ele pode utilizar � acima do padr�o
						if perc_comissao_e_desconto_a_utilizar > perc_comissao_e_desconto_padrao then
							.StatusDescontoSuperior = 1
							.IdUsuarioDescontoSuperior = r_usuario.Id
							.DataHoraDescontoSuperior = Now
							end if 'if perc_comissao_e_desconto_a_utilizar > perc_comissao_e_desconto_padrao
						end if 'if (.desc_dado + perc_RT) <= (perc_comissao_e_desconto_padrao + MAX_MARGEM_ERRO_PERC_DESC_E_RT) then-else

					'Verifica necessidade de senha de autoriza��o de desconto superior e se essa autoriza��o foi cadastrada
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

	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		if alerta = "" then
			if blnMagentoPedidoComIndicador then
				c_mag_installer_document = retorna_so_digitos(Trim("" & tMAP_XML("installer_document")))
				percCommissionValue = tMAP_XML("commission_value")
				percCommissionDiscount = tMAP_XML("commission_discount")

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
				end if 'if blnMagentoPedidoComIndicador
			end if 'if alerta = ""

		if blnMagentoPedidoComIndicador then
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
			end if 'if blnMagentoPedidoComIndicador
		end if 'if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO

'	L�GICA P/ CONSUMO DO ESTOQUE
	dim tipo_pessoa
	dim descricao_tipo_pessoa
	tipo_pessoa = multi_cd_regra_determina_tipo_pessoa(EndCob_tipo_pessoa, EndCob_contribuinte_icms_status, EndCob_produtor_rural_status)
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
		if Not obtemCtrlEstoqueProdutoRegra(EndCob_uf, EndCob_tipo_pessoa, EndCob_contribuinte_icms_status, EndCob_produtor_rural_status, vProdRegra, msg_erro) then
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
						alerta=alerta & "Falha desconhecida na leitura da regra de consumo do estoque para o produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " (UF: '" & EndCob_uf & "', tipo de pessoa: '" & descricao_tipo_pessoa & "')"
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
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " est� bloqueada para a UF '" & EndCob_uf & "'"
				elseif vProdRegra(iRegra).regra.regraUF.regraPessoa.st_inativo = 1 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " est� bloqueada para clientes '" & descricao_tipo_pessoa & "' da UF '" & EndCob_uf & "'"
				elseif converte_numero(vProdRegra(iRegra).regra.regraUF.regraPessoa.spe_id_nfe_emitente) = 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " n�o especifica nenhum CD para aguardar produtos sem presen�a no estoque para clientes '" & descricao_tipo_pessoa & "' da UF '" & EndCob_uf & "'"
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
						alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " n�o especifica nenhum CD ativo para clientes '" & descricao_tipo_pessoa & "' da UF '" & EndCob_uf & "'"
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
	
	dim erro_produto_sem_estoque
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
	erro_produto_sem_estoque = False
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
					erro_produto_sem_estoque = True
					end if
				end if
			next
		end if 'if alerta=""
	
	if alerta = "" then
		if erro_produto_sem_estoque then
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

		if CLng(s_etg_imediata) = CLng(COD_ETG_IMEDIATA_NAO) then
			if c_data_previsao_entrega = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "� necess�rio informar a data de previs�o de entrega"
			elseif Not IsDate(c_data_previsao_entrega) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Data de previs�o de entrega informada � inv�lida"
			elseif StrToDate(c_data_previsao_entrega) <= Date then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Data de previs�o de entrega deve ser uma data futura"
				end if
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
	
'	CONSIST�NCIAS
	if alerta = "" then
		if c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "1" then
			'CPF/CNPJ
			if Not cnpj_cpf_ok(EndCob_cnpj_cpf) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "CPF/CNPJ informado � inv�lido!"
				end if
			'NOME
			if EndCob_nome = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Nome do cliente n�o informado!"
				end if
			'TIPO DE PESSOA
			if EndCob_tipo_pessoa = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Tipo de pessoa n�o informado!"
				end if
			'INSCRI��O ESTADUAL
			if ((EndCob_tipo_pessoa = ID_PJ) And (EndCob_contribuinte_icms_status = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM)) _
				Or _
				((EndCob_tipo_pessoa = ID_PJ) And ((EndCob_contribuinte_icms_status = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) And (EndCob_ie <> ""))) _
				Or _
				((EndCob_tipo_pessoa = ID_PF) And (EndCob_produtor_rural_status = COD_ST_CLIENTE_PRODUTOR_RURAL_SIM)) then
				if Not isInscricaoEstadualValida(EndCob_ie, EndCob_uf) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "N�mero de IE (inscri��o estadual) informado � inv�lido (IE: " & EndCob_ie & ", UF: " & EndCob_uf & ")!" & _
									"<br />" & "Certifique-se de que a UF do endere�o corresponde � UF respons�vel pelo registro da IE."
					end if
				end if
			'ENDERE�O
			if EndCob_endereco = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Endere�o (logradouro) n�o informado!"
				end if
			if Len(EndCob_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Endere�o (logradouro) excede o tamanho m�ximo permitido:<br />Tamanho atual: " & Cstr(Len(EndCob_endereco)) & " caracteres<br />Tamanho m�ximo: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " caracteres"
				end if
			'N�MERO DO ENDERE�O
			if EndCob_endereco_numero = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "N�mero do endere�o n�o informado!"
				end if
			'CIDADE
			if EndCob_cidade = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Cidade do endere�o n�o informada!"
				end if
			'UF
			if EndCob_uf = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "UF do endere�o n�o informada!"
				end if
			'CEP
			if EndCob_cep = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "CEP n�o informado!"
				end if
		else 'bloco else: if c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "1"
			'CPF/CNPJ
			if Not cnpj_cpf_ok(EndCob_cnpj_cpf) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "CPF/CNPJ do cadastro do cliente � inv�lido!"
				end if
			'NOME
			if EndCob_nome = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "� necess�rio preencher o nome no cadastro do cliente!"
				end if
			'TIPO DE PESSOA
			if EndCob_tipo_pessoa = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "� necess�rio preencher o tipo de pessoa no cadastro do cliente!"
				end if
			'INSCRI��O ESTADUAL
			if ((EndCob_tipo_pessoa = ID_PJ) And (EndCob_contribuinte_icms_status = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM)) _
				Or _
				((EndCob_tipo_pessoa = ID_PJ) And ((EndCob_contribuinte_icms_status = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) And (EndCob_ie <> ""))) _
				Or _
				((EndCob_tipo_pessoa = ID_PF) And (EndCob_produtor_rural_status = COD_ST_CLIENTE_PRODUTOR_RURAL_SIM)) then
				if Not isInscricaoEstadualValida(EndCob_ie, EndCob_uf) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "N�mero de IE (inscri��o estadual) no cadastro do cliente � inv�lido (IE: " & EndCob_ie & ", UF: " & EndCob_uf & ")!" & _
									"<br />" & "Certifique-se de que a UF do endere�o corresponde � UF respons�vel pelo registro da IE."
					end if
				end if
			'ENDERE�O
			if EndCob_endereco = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "� necess�rio preencher o endere�o (logradouro) no cadastro do cliente!"
				end if
			if Len(EndCob_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Endere�o (logradouro) do cadastro do cliente excede o tamanho m�ximo permitido:<br />Tamanho atual: " & Cstr(Len(EndCob_endereco)) & " caracteres<br />Tamanho m�ximo: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " caracteres"
				end if
			'N�MERO DO ENDERE�O
			if EndCob_endereco_numero = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "� necess�rio preencher o n�mero do endere�o no cadastro do cliente!"
				end if
			'CIDADE
			if EndCob_cidade = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "� necess�rio preencher a cidade no cadastro do cliente!"
				end if
			'UF
			if EndCob_uf = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "� necess�rio preencher a UF no cadastro do cliente!"
				end if
			'CEP
			if EndCob_cep = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "� necess�rio preencher o CEP no cadastro do cliente!"
				end if
			end if 'if c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "1"
		end if

'	ENDERE�O DE ENTREGA
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
			
			if ( (EndEtg_tipo_pessoa = ID_PF) And (Cstr(EndEtg_produtor_rural_status) = Cstr(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM)) ) _
				Or _
			   ( (EndEtg_tipo_pessoa = ID_PJ) And (Cstr(EndEtg_contribuinte_icms_status) = Cstr(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM)) ) _
			   Or _
			   ( (EndEtg_tipo_pessoa = ID_PJ) And (Cstr(EndEtg_contribuinte_icms_status) = Cstr(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO)) And (Trim(EndEtg_ie) <> "") ) then
				if Not isInscricaoEstadualValida(EndEtg_ie, EndEtg_uf) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Endere�o de entrega: preencha a IE (Inscri��o Estadual) com um n�mero v�lido!" & _
							"<br>" & "Certifique-se de que a UF do endere�o de entrega corresponde � UF respons�vel pelo registro da IE."
					end if
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
			if EndCob_cep <> "" then
				sTranspSelAutoTransportadoraId = obtem_transportadora_pelo_cep(retorna_so_digitos(EndCob_cep))
				if sTranspSelAutoTransportadoraId <> "" then
					sTranspSelAutoCep = retorna_so_digitos(EndCob_cep)
					iTranspSelAutoTipoEndereco = TRANSPORTADORA_SELECAO_AUTO_TIPO_ENDERECO_CLIENTE
					iTranspSelAutoStatus = TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_S
					end if
				end if
			end if
		end if
	
	'TRATAMENTO PARA CADASTRAMENTO DE PEDIDOS DO SITE MAGENTO DA BONSHOP
	if isLojaBonshop(loja) And (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) then
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
						" tP.endereco_cnpj_cpf AS cnpj_cpf," & _
						" tP.endereco_nome AS nome_cliente" & _
					" FROM t_PEDIDO tP" & _
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
		end if 'if isLojaBonshop(loja) And (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO)

	'TRATAMENTO PARA CADASTRAMENTO DE PEDIDOS DO SITE MAGENTO DO ARCLUBE
	dim blnPedidoECommerceCreditoOkAutomatico
	blnPedidoECommerceCreditoOkAutomatico = False
	if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
		if alerta = "" then
			if s_origem_pedido = "" then
				alerta = "Informe a origem do pedido"
				end if
			end if
		
		if alerta = "" then
		'	PARA PEDIDOS DO ARCLUBE, � PERMITIDO FICAR SEM O N� MAGENTO SOMENTE NOS SEGUINTES CASOS:
		'		1) PEDIDO ORIGINADO PELO TELEVENDAS
		'		2) PEDIDO GERADO CONTRA A TRANSPORTADORA (EM CASOS QUE A TRANSPORTADORA SE RESPONSABILIZA PELA REPOSI��O DE MERCADORIA EXTRAVIADA)
			if (Trim(s_origem_pedido) <> "002") And (Trim(s_origem_pedido) <> "019") then
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

				if r_loja.magento_api_versao <> VERSAO_API_MAGENTO_V2_REST_JSON then
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
					if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
						if Trim("" & tMAP_XML("b2b_type_order")) <> COD_MAGENTO_TYPE_ORDER__INSTALLER then perc_RT = rs2("parametro_campo_real")
					else
						perc_RT = rs2("parametro_campo_real")
						end if
				'	DEVE COLOCAR AUTOMATICAMENTE COM 'CR�DITO OK'?
					if rs2("parametro_1_campo_flag") = 1 then blnPedidoECommerceCreditoOkAutomatico = True
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
						" tP.endereco_cnpj_cpf AS cnpj_cpf," & _
						" tP.endereco_nome AS nome_cliente" & _
					" FROM t_PEDIDO tP" & _
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
						" tP.endereco_cnpj_cpf AS cnpj_cpf," & _
						" tP.endereco_nome AS nome_cliente" & _
					" FROM t_PEDIDO tP" & _
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
		dim s_log_item_servico
		indice_pedido = 0
		id_pedido_base = ""
		id_pedido_temp_base = ""
		s_log=""
		s_log_cliente_indicador=""
		s_log_item_servico = ""
		redim vLogAutoSplit(0)
		vLogAutoSplit(UBound(vLogAutoSplit)) = ""
		s_hora_pedido = retorna_so_digitos(formata_hora(Now))
		if Not gera_num_pedido_temp(id_pedido_temp_base, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO then
		'	BLOQUEIA REGISTRO PARA EVITAR ACESSO CONCORRENTE (REALIZA O FLIP EM UM CAMPO BIT APENAS P/ ADQUIRIR O LOCK EXCLUSIVO)
		'	OBS: TODOS OS M�DULOS DO SISTEMA QUE REALIZEM ESTA OPERA��O DE CADASTRAMENTO DEVEM SINCRONIZAR O ACESSO OBTENDO O LOCK EXCLUSIVO DO REGISTRO DE CONTROLE DESIGNADO
			s = "UPDATE t_CONTROLE SET" & _
					" dummy = ~dummy" & _
				" WHERE" & _
					" id_nsu = '" & ID_XLOCK_SYNC_PEDIDO & "'"
			cn.Execute(s)
			end if

	'	VERIFICA SE ESTE PEDIDO J� FOI GRAVADO!!
	'	TRATAMENTO P/ OS CASOS EM QUE: USU�RIO EST� TENTANDO USAR O BOT�O VOLTAR, OCORREU DUPLO CLIQUE OU USU�RIO ATUALIZOU A P�GINA ENQUANTO AINDA ESTAVA PROCESSANDO (DUPLO ACIONAMENTO)
	'	Esse tratamento est� sendo feito atrav�s do campo insert_request_guid (t_PEDIDO.InsertRequestGuid), mas se esse campo estiver vazio, continua realizando o controle pelo m�todo antigo
	'	Realiza a verifica��o ap�s obter o lock do registro de controle p/ assegurar que n�o haver� problemas de acesso concorrente
		if insert_request_guid <> "" then
			s = "SELECT pedido, vendedor, data_hora FROM t_PEDIDO WHERE (InsertRequestGuid = '" & insert_request_guid & "')"
			set rs = cn.execute(s)
			if Not rs.Eof then
				alerta = "Este pedido j� foi gravado com o n� " & Trim("" & rs("pedido")) & " em " & formata_data_hora(rs("data_hora")) & " por " & Trim("" & rs("vendedor"))
				end if
			if rs.State <> 0 then rs.Close
			end if 'if insert_request_guid <> ""

		if alerta = "" then
			if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				If Not cria_recordset_pessimista(tMAP_ITEM, msg_erro) then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
					end if
				
				If Not cria_recordset_pessimista(tITEM_SVC, msg_erro) then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
					end if
				end if 'if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO
			
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
						if insert_request_guid <> "" then rs("InsertRequestGuid") = "{" & insert_request_guid & "}"
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
						if blnPedidoECommerceCreditoOkAutomatico then
							rs("analise_credito")=Clng(COD_AN_CREDITO_OK)
							rs("analise_credito_data")=Now
							rs("analise_credito_usuario")="AUTOM�TICO"
						elseif vl_total <= vl_aprov_auto_analise_credito then
							rs("analise_credito")=Clng(COD_AN_CREDITO_OK)
							rs("analise_credito_data")=Now
							rs("analise_credito_usuario")="AUTOM�TICO"
						elseif (Cstr(loja) = Cstr(NUMERO_LOJA_TRANSFERENCIA)) Or (Cstr(loja) = Cstr(NUMERO_LOJA_KITS)) Or isLojaGarantia(loja) then
							'Lojas usadas para pedidos de opera��es internas
							rs("analise_credito")=Clng(COD_AN_CREDITO_OK)
							rs("analise_credito_data")=Now
							rs("analise_credito_usuario")="AUTOM�TICO"
						elseif (Cstr(loja) = NUMERO_LOJA_INDICACAO_OBRA) And (rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA) And ((CStr(op_av_forma_pagto) = Cstr(ID_FORMA_PAGTO_DINHEIRO)) Or (CStr(op_av_forma_pagto) = Cstr(ID_FORMA_PAGTO_DEPOSITO))) then
							'Loja Indica��o de Obra: Cr�dito Ok autom�tico p/ pagto � vista em dinheiro ou dep�sito
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
						elseif (rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA) And (CStr(op_av_forma_pagto) = CStr(ID_FORMA_PAGTO_DEPOSITO)) then
							rs("analise_credito")=Clng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)
							rs("analise_credito_data")=Now
							rs("analise_credito_usuario")="AUTOM�TICO"
						elseif (rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA) And (CStr(op_av_forma_pagto) = Cstr(ID_FORMA_PAGTO_BOLETO_AV)) then
							rs("analise_credito")=Clng(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)
							rs("analise_credito_data")=Now
							rs("analise_credito_usuario")="AUTOM�TICO"
							'OBSERVA��O: no caso do 'parcelado com entrada' quando a entrada � 'Boleto AV', o pedido deve continuar sendo cadastrado com o status de an�lise de cr�dito
							'seguindo a l�gica j� existente. Quando o depto de an�lise de cr�dito aprovar o pedido, ir� se encarregar de alterar manualmente o pedido para
							'"Cr�dito OK (aguardando pagto boleto AV)"
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

						vl_total_RA_liquido = CCur(vl_total_RA - (perc_desagio_RA_liquida/100)*vl_total_RA)
						vl_total_RA_liquido = converte_numero(formata_moeda(vl_total_RA_liquido))

						rs("vl_total_RA_liquido") = vl_total_RA_liquido
						rs("qtde_parcelas_desagio_RA") = 0
						if vl_total_RA <> 0 then
							rs("st_tem_desagio_RA") = 1
						else
							rs("st_tem_desagio_RA") = 0
							end if

						if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) then
							rs("vl_frete_total_cobrado_cliente") = vlMagentoShippingAmount
							rs("vl_base_calculo_frete_total_cobrado_cliente") = vl_total_produto_magento
							end if
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
						rs("EtgImediataIdTipoUsuarioContexto") = COD_USUARIO_CONTEXTO__USUARIO_INTERNO
						rs("EtgImediataIdUsuarioUltAtualiz") = r_usuario.Id
						end if
					if CLng(s_etg_imediata) = CLng(COD_ETG_IMEDIATA_NAO) then
						rs("PrevisaoEntregaData") = StrToDate(c_data_previsao_entrega)
						rs("PrevisaoEntregaUsuarioUltAtualiz") = usuario
						rs("PrevisaoEntregaDtHrUltAtualiz") = Now
						rs("PrevisaoEntregaIdTipoUsuarioContexto") = COD_USUARIO_CONTEXTO__USUARIO_INTERNO
						rs("PrevisaoEntregaIdUsuarioUltAtualiz") = r_usuario.Id
						end if
					if s_bem_uso_consumo <> "" then 
						rs("StBemUsoConsumo")=CLng(s_bem_uso_consumo)
						end if
					if s_instalador_instala <> "" then
						rs("InstaladorInstalaStatus")=CLng(s_instalador_instala)
						rs("InstaladorInstalaUsuarioUltAtualiz")=usuario
						rs("InstaladorInstalaDtHrUltAtualiz")=Now
						rs("InstaladorInstalaIdTipoUsuarioContexto") = COD_USUARIO_CONTEXTO__USUARIO_INTERNO
						rs("InstaladorInstalaIdUsuarioUltAtualiz") = r_usuario.Id
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
					rs("GarantiaIndicadorIdTipoUsuarioContexto") = COD_USUARIO_CONTEXTO__USUARIO_INTERNO
					rs("GarantiaIndicadorIdUsuarioUltAtualiz") = r_usuario.Id

					if rb_end_entrega = "S" then
						rs("st_end_entrega") = 1
						if (c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "1") Or (c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "9") then
							rs("EndEtg_endereco") = Left(EndEtg_endereco, MAX_TAMANHO_CAMPO_ENDERECO)
							rs("EndEtg_endereco_numero") = Left(EndEtg_endereco_numero, MAX_TAMANHO_CAMPO_ENDERECO_NUMERO)
							rs("EndEtg_endereco_complemento") = Left(EndEtg_endereco_complemento, MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO)
							rs("EndEtg_bairro") = Left(EndEtg_bairro, MAX_TAMANHO_CAMPO_ENDERECO_BAIRRO)
							rs("EndEtg_cidade") = Left(EndEtg_cidade, MAX_TAMANHO_CAMPO_ENDERECO_CIDADE)
						else
							rs("EndEtg_endereco") = EndEtg_endereco
							rs("EndEtg_endereco_numero") = EndEtg_endereco_numero
							rs("EndEtg_endereco_complemento") = EndEtg_endereco_complemento
							rs("EndEtg_bairro") = EndEtg_bairro
							rs("EndEtg_cidade") = EndEtg_cidade
							end if

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

						sBlocoNotasEndEtg = formata_endereco(Trim("" & rs("EndEtg_endereco")), Trim("" & rs("EndEtg_endereco_numero")), Trim("" & rs("EndEtg_endereco_complemento")), Trim("" & rs("EndEtg_bairro")), Trim("" & rs("EndEtg_cidade")), Trim("" & rs("EndEtg_uf")), Trim("" & rs("EndEtg_cep")))
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
					if (Cstr(loja) <> Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE)) And (Not blnMagentoPedidoComIndicador) then rs("perc_desagio_RA_liquida") = perc_desagio_RA_liquida

					if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) then
						rs("magento_shipping_amount") = vlMagentoShippingAmount
						if blnMagentoPedidoComIndicador then
							rs("magento_installer_commission_value") = percCommissionValue
							rs("magento_installer_commission_discount") = percCommissionDiscount
							end if
						end if

					rs("permite_RA_status") = permite_RA_status
		
					if permite_RA_status = 1 then
						rs("opcao_possui_RA") = rb_RA
					else
						rs("opcao_possui_RA") = "-" ' N�o se aplica
						end if
		
					rs("endereco_memorizado_status") = 1
				
					if (c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "1") Or (c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "9") then
						rs("endereco_logradouro") = Left(EndCob_endereco, MAX_TAMANHO_CAMPO_ENDERECO)
						rs("endereco_numero") = Left(EndCob_endereco_numero, MAX_TAMANHO_CAMPO_ENDERECO_NUMERO)
						rs("endereco_complemento") = Left(EndCob_endereco_complemento, MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO)
						rs("endereco_bairro") = Left(EndCob_bairro, MAX_TAMANHO_CAMPO_ENDERECO_BAIRRO)
						rs("endereco_cidade") = Left(EndCob_cidade, MAX_TAMANHO_CAMPO_ENDERECO_CIDADE)
					else
						rs("endereco_logradouro") = EndCob_endereco
						rs("endereco_numero") = EndCob_endereco_numero
						rs("endereco_complemento") = EndCob_endereco_complemento
						rs("endereco_bairro") = EndCob_bairro
						rs("endereco_cidade") = EndCob_cidade
						end if

					rs("endereco_uf") = EndCob_uf
					rs("endereco_cep") = EndCob_cep

					if blnUsarMemorizacaoCompletaEnderecos then
						rs("st_memorizacao_completa_enderecos") = 1
						rs("endereco_email") = EndCob_email
						rs("endereco_email_xml") = EndCob_email_xml
						rs("endereco_nome") = EndCob_nome
						rs("endereco_ddd_res") = EndCob_ddd_res
						rs("endereco_tel_res") = EndCob_tel_res
						rs("endereco_ddd_com") = EndCob_ddd_com
						rs("endereco_tel_com") = EndCob_tel_com
						rs("endereco_ramal_com") = EndCob_ramal_com
						rs("endereco_ddd_cel") = EndCob_ddd_cel
						rs("endereco_tel_cel") = EndCob_tel_cel
						rs("endereco_ddd_com_2") = EndCob_ddd_com_2
						rs("endereco_tel_com_2") = EndCob_tel_com_2
						rs("endereco_ramal_com_2") = EndCob_ramal_com_2
						rs("endereco_tipo_pessoa") = EndCob_tipo_pessoa
						rs("endereco_cnpj_cpf") = EndCob_cnpj_cpf
						rs("endereco_contribuinte_icms_status") = converte_numero(EndCob_contribuinte_icms_status)
						rs("endereco_produtor_rural_status") = converte_numero(EndCob_produtor_rural_status)
						rs("endereco_ie") = EndCob_ie
						rs("endereco_rg") = EndCob_rg
						rs("endereco_contato") = EndCob_contato
						end if

					sBlocoNotasEndCob = formata_endereco(Trim("" & rs("endereco_logradouro")), Trim("" & rs("endereco_numero")), Trim("" & rs("endereco_complemento")), Trim("" & rs("endereco_bairro")), Trim("" & rs("endereco_cidade")), Trim("" & rs("endereco_uf")), Trim("" & rs("endereco_cep")))

					if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) OR ( (Cstr(loja) = Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE)) And (Trim(s_pedido_ac) <> "") ) then
						rs("plataforma_origem_pedido") = COD_PLATAFORMA_ORIGEM_PEDIDO__MAGENTO
					else
						rs("plataforma_origem_pedido") = COD_PLATAFORMA_ORIGEM_PEDIDO__ERP
						end if

					rs("sistema_responsavel_cadastro") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
					rs("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP

					rs("id_nfe_emitente") = vEmpresaAutoSplit(iv)

					if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
						rs("magento_quote_id") = tMAP_XML("quote_id")
						rs("magento_installer_name") = Trim("" & tMAP_XML("b2b_installer_name"))
						rs("magento_installer_id") = tMAP_XML("b2b_installer_id")
						rs("magento_installer_commission_value") = tMAP_XML("b2b_installer_commission_value")
						rs("magento_installer_commission_percentage") = tMAP_XML("b2b_installer_commission_percentage")
						rs("magento_type_order") = tMAP_XML("b2b_type_order")
						end if

					rs("perc_max_comissao_padrao") = rCD.perc_max_comissao
					rs("perc_max_comissao_e_desconto_padrao") = perc_comissao_e_desconto_padrao
					rs("UsuarioCadastroIdTipoUsuarioContexto") = COD_USUARIO_CONTEXTO__USUARIO_INTERNO
					rs("UsuarioCadastroId") = r_usuario.Id
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
											rs("subgrupo") = .subgrupo
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
											rs("cod_produto_xml_fabricante") = .cod_produto_xml_fabricante
											rs("cod_produto_alfanum_fabricante") = .cod_produto_alfanum_fabricante
											rs("potencia_valor") = .potencia_valor
											rs("id_unidade_potencia") = .id_unidade_potencia
											rs("StatusDescontoSuperior") = .StatusDescontoSuperior
											if .StatusDescontoSuperior <> 0 then
												rs("IdUsuarioDescontoSuperior") = CLng(.IdUsuarioDescontoSuperior)
												rs("DataHoraDescontoSuperior") = .DataHoraDescontoSuperior
												end if
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
				
					'Cadastramento semi-autom�tico de pedidos do Magento: se houver servi�os no pedido, cadastra no pedido-base
					if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
						if indice_pedido = 1 then
							sequencia_item = 0
							s = "SELECT " & _
									"*" & _
								" FROM t_MAGENTO_API_PEDIDO_XML_DECODE_ITEM" & _
								" WHERE" & _
									" (id_magento_api_pedido_xml = " & id_magento_api_pedido_xml & ")" & _
									" AND (product_type = '" & COD_MAGENTO_PRODUCT_TYPE__VIRTUAL & "')" & _
								" ORDER BY" & _
									" id"
							if tMAP_ITEM.State <> 0 then tMAP_ITEM.Close
							tMAP_ITEM.open s, cn
							do while Not tMAP_ITEM.Eof
								sequencia_item = sequencia_item + 1
								s = "SELECT * FROM t_PEDIDO_ITEM_SERVICO WHERE pedido='X'"
								if tITEM_SVC.State <> 0 then tITEM_SVC.Close
								tITEM_SVC.Open s, cn
								tITEM_SVC.AddNew
								tITEM_SVC("pedido") = id_pedido_temp
								tITEM_SVC("fabricante") = ""
								tITEM_SVC("produto") = Trim("" & tMAP_ITEM("sku"))
								tITEM_SVC("qtde") = CLng(tMAP_ITEM("qty_ordered"))
								vl_servico_original_price = converte_numero(tMAP_ITEM("original_price"))
								'O campo discount_amount informa o valor total do desconto j� multiplicado pela quantidade, ou seja, n�o h� campo com o valor unit�rio do desconto aplicado e
								'nem o valor unit�rio de venda do item j� com o desconto aplicado.
								'J� o campo row_total informa o valor total do item j� calculado com os descontos e multiplicado pela quantidade
								vl_servico_price = 0
								if converte_numero(tMAP_ITEM("qty_ordered")) > 0 then
									vl_servico_price = converte_numero(tMAP_ITEM("row_total")) / converte_numero(tMAP_ITEM("qty_ordered"))
									end if
								'Calcula percentual de desconto
								percDescServico = 0
								if vl_servico_original_price <> 0 then
									percDescServico = 100*((vl_servico_original_price - vl_servico_price)/vl_servico_original_price)
									end if
								tITEM_SVC("desc_dado") = percDescServico
								tITEM_SVC("preco_venda") = vl_servico_price
								tITEM_SVC("preco_NF") = vl_servico_price
								tITEM_SVC("preco_lista") = vl_servico_original_price
								tITEM_SVC("descricao") = Trim("" & tMAP_ITEM("name"))
								tITEM_SVC("descricao_html") = Trim("" & tMAP_ITEM("name"))
								tITEM_SVC("sequencia") = sequencia_item
							
								if s_log_item_servico <> "" then s_log_item_servico = s_log_item_servico & chr(13)
								s_log_item_servico = s_log_item_servico & Trim("" & tITEM_SVC("produto")) & ":" & _
													" qtde = " & Trim("" & tITEM_SVC("qtde")) & ", " & _
													" preco_venda = " & formata_moeda(tITEM_SVC("preco_venda")) & ", " & _
													" preco_NF = " & formata_moeda(tITEM_SVC("preco_NF")) & ", " & _
													" preco_lista = " & formata_moeda(tITEM_SVC("preco_lista")) & ", " & _
													" desc_dado = " & formata_perc(tITEM_SVC("desc_dado")) & ", " & _
													" descricao = " & Trim("" & tITEM_SVC("descricao"))

								tITEM_SVC.Update
								if Err <> 0 then
								'	~~~~~~~~~~~~~~~~
									cn.RollbackTrans
								'	~~~~~~~~~~~~~~~~
									Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
									end if

								if tITEM_SVC.State <> 0 then tITEM_SVC.Close

								tMAP_ITEM.MoveNext
								loop
							
							'Grava no pedido os dados de pagamento do marketplace, se houver
							'Esses dados s�o usados para informar os dados do intermediador da transa��o na NFe
							s = "INSERT INTO t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT (" & _
									"pedido" & _
									", value" & _
									", type" & _
									", transaction_date" & _
									", status" & _
									", parcels" & _
									", method" & _
									", description" & _
									", card_issuer" & _
									", autorization_id" & _
									", sefaz_type_integration" & _
									", sefaz_payment_indicator" & _
									", sefaz_name_payment" & _
									", sefaz_name_card_issuer" & _
									", sefaz_id_payment" & _
									", sefaz_id_card_issuer" & _
								") SELECT " & _
									"'" & id_pedido_temp & "'" & _
									", value" & _
									", type" & _
									", transaction_date" & _
									", status" & _
									", parcels" & _
									", method" & _
									", description" & _
									", card_issuer" & _
									", autorization_id" & _
									", sefaz_type_integration" & _
									", sefaz_payment_indicator" & _
									", sefaz_name_payment" & _
									", sefaz_name_card_issuer" & _
									", sefaz_id_payment" & _
									", sefaz_id_card_issuer" & _
								" FROM t_MAGENTO_API_PEDIDO_XML_DECODE_SKYHUB_MKTP_PAYMENT" & _
								" WHERE" & _
									" (id_magento_api_pedido_xml = " & id_magento_api_pedido_xml & ")"
							cn.Execute(s)
							end if 'if indice_pedido = 1
						end if 'if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO

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
		
					if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
						s="UPDATE t_PEDIDO_ITEM_SERVICO SET pedido='" & id_pedido & "' WHERE pedido='" & id_pedido_temp & "'"
						cn.Execute(s)

						s="UPDATE t_PEDIDO_MAGENTO_SKYHUB_MKTP_PAYMENT SET pedido='" & id_pedido & "' WHERE pedido='" & id_pedido_temp & "'"
						cn.Execute(s)
						end if

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
					
					'Atualiza o campo 'nsu_pedido_base' e se refere ao valor salvo em t_PEDIDO.nsu do registro do pedido-base.
					'O valor t_PEDIDO.nsu � gerado por uma CONSTRAINT DEFAULT em conjunto c/ a SEQUENCE 'seq_T_PEDIDO' e s�
					'� gerada de fato ap�s o UPDATE.
					'Os campos 'nsu_pedido_base' e 'nsu' s�o usados principalmente p/ ordena��o dos pedidos, principalmente
					'nos ambientes em que a letra do sufixo seguem a ordem decrescente.
					'O campo 'nsu_pedido_base' tem por finalidade facilitar o agrupamento da fam�lia de pedidos na ordena��o,
					'j� que no SPLIT MANUAL, o campo 'nsu' n�o ser� consecutivo em rela��o aos demais pedidos da fam�lia.
					s = "UPDATE tPed SET" & _
							" tPed.nsu_pedido_base = tPedBase.nsu" & _
						" FROM t_PEDIDO tPed" & _
							" INNER JOIN t_PEDIDO tPedBase ON (tPed.pedido_base = tPedBase.pedido)" & _
						" WHERE" & _
							" (tPed.pedido_base = '" & id_pedido_base & "')" & _
							" AND (tPed.nsu_pedido_base = 0)"
					cn.Execute(s)
					
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
								if isEnderecoIgual(EndCob_endereco, EndCob_endereco_numero, EndCob_cep, r_orcamentista_e_indicador.endereco, r_orcamentista_e_indicador.endereco_numero, r_orcamentista_e_indicador.cep) then
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
										rs("endereco_logradouro") = Left(EndCob_endereco, MAX_TAMANHO_CAMPO_ENDERECO)
										rs("endereco_numero") = Left(EndCob_endereco_numero, MAX_TAMANHO_CAMPO_ENDERECO_NUMERO)
										rs("endereco_complemento") = Left(EndCob_endereco_complemento, MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO)
										rs("endereco_bairro") = Left(EndCob_bairro, MAX_TAMANHO_CAMPO_ENDERECO_BAIRRO)
										rs("endereco_cidade") = Left(EndCob_cidade, MAX_TAMANHO_CAMPO_ENDERECO_CIDADE)
										rs("endereco_uf") = EndCob_uf
										rs("endereco_cep") = EndCob_cep
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
												rs("endereco_logradouro") = Left(.endereco, MAX_TAMANHO_CAMPO_ENDERECO)
												rs("endereco_numero") = Left(.endereco_numero, MAX_TAMANHO_CAMPO_ENDERECO_NUMERO)
												rs("endereco_complemento") = Left(.endereco_complemento, MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO)
												rs("endereco_bairro") = Left(.bairro, MAX_TAMANHO_CAMPO_ENDERECO_BAIRRO)
												rs("endereco_cidade") = Left(.cidade, MAX_TAMANHO_CAMPO_ENDERECO_CIDADE)
												rs("endereco_uf") = .uf
												rs("endereco_cep") = .cep
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
				
								'em 2020-04-16 n�o temos mais registros com endereco_memorizado_status = 0
								s = "SELECT DISTINCT * FROM " & _
										"(" & _
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
												" (id_cliente <> '" & cliente_selecionado & "')" & _
												" AND (endereco_cep = '" & retorna_so_digitos(EndCob_cep) & "')" & _
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
												" AND (EndEtg_cep = '" & retorna_so_digitos(EndCob_cep) & "')" & _
										") t" & _
									" ORDER BY" & _
										" data_hora DESC"
								if rs.State <> 0 then rs.Close
								rs.Open s, cn
								do while Not rs.Eof
									if isEnderecoIgual(EndCob_endereco, EndCob_endereco_numero, EndCob_cep, Trim("" & rs("endereco_logradouro")), Trim("" & rs("endereco_numero")), Trim("" & rs("endereco_cep"))) then
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
												rs("endereco_logradouro") = Left(EndCob_endereco, MAX_TAMANHO_CAMPO_ENDERECO)
												rs("endereco_numero") = Left(EndCob_endereco_numero, MAX_TAMANHO_CAMPO_ENDERECO_NUMERO)
												rs("endereco_complemento") = Left(EndCob_endereco_complemento, MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO)
												rs("endereco_bairro") = Left(EndCob_bairro, MAX_TAMANHO_CAMPO_ENDERECO_BAIRRO)
												rs("endereco_cidade") = Left(EndCob_cidade, MAX_TAMANHO_CAMPO_ENDERECO_CIDADE)
												rs("endereco_uf") = EndCob_uf
												rs("endereco_cep") = retorna_so_digitos(EndCob_cep)
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
											rs("endereco_logradouro") = Left(.endereco_logradouro, MAX_TAMANHO_CAMPO_ENDERECO)
											rs("endereco_numero") = Left(.endereco_numero, MAX_TAMANHO_CAMPO_ENDERECO_NUMERO)
											rs("endereco_complemento") = Left(.endereco_complemento, MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO)
											rs("endereco_bairro") = Left(.endereco_bairro, MAX_TAMANHO_CAMPO_ENDERECO_BAIRRO)
											rs("endereco_cidade") = Left(.endereco_cidade, MAX_TAMANHO_CAMPO_ENDERECO_CIDADE)
											rs("endereco_uf") = .endereco_uf
											rs("endereco_cep") = .endereco_cep
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
											rs("endereco_logradouro") = Left(EndEtg_endereco, MAX_TAMANHO_CAMPO_ENDERECO)
											rs("endereco_numero") = Left(EndEtg_endereco_numero, MAX_TAMANHO_CAMPO_ENDERECO_NUMERO)
											rs("endereco_complemento") = Left(EndEtg_endereco_complemento, MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO)
											rs("endereco_bairro") = Left(EndEtg_bairro, MAX_TAMANHO_CAMPO_ENDERECO_BAIRRO)
											rs("endereco_cidade") = Left(EndEtg_cidade, MAX_TAMANHO_CAMPO_ENDERECO_CIDADE)
											rs("endereco_uf") = EndEtg_uf
											rs("endereco_cep") = EndEtg_cep
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
													rs("endereco_logradouro") = Left(.endereco, MAX_TAMANHO_CAMPO_ENDERECO)
													rs("endereco_numero") = Left(.endereco_numero, MAX_TAMANHO_CAMPO_ENDERECO_NUMERO)
													rs("endereco_complemento") = Left(.endereco_complemento, MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO)
													rs("endereco_bairro") = Left(.bairro, MAX_TAMANHO_CAMPO_ENDERECO_BAIRRO)
													rs("endereco_cidade") = Left(.cidade, MAX_TAMANHO_CAMPO_ENDERECO_CIDADE)
													rs("endereco_uf") = .uf
													rs("endereco_cep") = .cep
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
						
										'em 2020-04-16 n�o temos mais registros com endereco_memorizado_status = 0
										s = "SELECT DISTINCT * FROM " & _
												"(" & _
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
														" (id_cliente <> '" & cliente_selecionado & "')" & _
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
														rs("endereco_logradouro") = Left(EndEtg_endereco, MAX_TAMANHO_CAMPO_ENDERECO)
														rs("endereco_numero") = Left(EndEtg_endereco_numero, MAX_TAMANHO_CAMPO_ENDERECO_NUMERO)
														rs("endereco_complemento") = Left(EndEtg_endereco_complemento, MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO)
														rs("endereco_bairro") = Left(EndEtg_bairro, MAX_TAMANHO_CAMPO_ENDERECO_BAIRRO)
														rs("endereco_cidade") = Left(EndEtg_cidade, MAX_TAMANHO_CAMPO_ENDERECO_CIDADE)
														rs("endereco_uf") = EndEtg_uf
														rs("endereco_cep") = EndEtg_cep
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
													rs("endereco_logradouro") = Left(.endereco_logradouro, MAX_TAMANHO_CAMPO_ENDERECO)
													rs("endereco_numero") = Left(.endereco_numero, MAX_TAMANHO_CAMPO_ENDERECO_NUMERO)
													rs("endereco_complemento") = Left(.endereco_complemento, MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO)
													rs("endereco_bairro") = Left(.endereco_bairro, MAX_TAMANHO_CAMPO_ENDERECO_BAIRRO)
													rs("endereco_cidade") = Left(.endereco_cidade, MAX_TAMANHO_CAMPO_ENDERECO_CIDADE)
													rs("endereco_uf") = .endereco_uf
													rs("endereco_cep") = .endereco_cep
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

					'Registra no bloco de notas que o pedido-filhote foi gerado por split autom�tico
					if alerta = "" then
						if indice_pedido > 1 then
							sBlocoNotasMsg = "Pedido gerado atrav�s de split autom�tico durante o cadastramento inicial"
							if Not grava_bloco_notas_pedido(id_pedido, ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_SPLIT_AUTOMATICO, msg_erro) then
								alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & id_pedido & ")"
								end if
							end if
						end if

					'Registra no bloco de notas os dados do endere�o inicial
					if alerta = "" then
						sBlocoNotasMsg = "Endere�o de cobran�a inicial: " & vbCrLf & _
										sBlocoNotasEndCob

						sBlocoNotasMsg = sBlocoNotasMsg & vbCrLf & _
										vbCrLf & _
										"Endere�o de entrega inicial: " & vbCrLf
						if rb_end_entrega = "S" then
							sBlocoNotasMsg = sBlocoNotasMsg & sBlocoNotasEndEtg
						else
							sBlocoNotasMsg = sBlocoNotasMsg & "(N.I.)"
							end if

						if Not grava_bloco_notas_pedido(id_pedido, ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_ENDERECO, msg_erro) then
							alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & id_pedido & ")"
							end if
						end if

					'Registra no bloco de notas o meio de pagamento informado pelo marketplace, se for o caso
					if alerta = "" then
						if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
							if indice_pedido = 1 then
								s_mktp_payment = ""
								s = "SELECT * FROM t_MAGENTO_API_PEDIDO_XML_DECODE_SKYHUB_MKTP_PAYMENT WHERE (id_magento_api_pedido_xml = " & id_magento_api_pedido_xml & ") ORDER BY id"
								if rs.State <> 0 then rs.Close
								rs.Open s, cn
								do while Not rs.Eof
									if Trim("" & rs("method")) <> "" then
										if s_mktp_payment <> "" then s_mktp_payment = s_mktp_payment & ", "
										s_mktp_payment = s_mktp_payment & Trim("" & rs("method"))
										if Trim("" & rs("value")) <> "" then
											s_mktp_payment = s_mktp_payment & " (" & formata_moeda(rs("value")) & ")"
											end if
										end if
									rs.MoveNext
									loop

								if s_mktp_payment <> "" then
									sBlocoNotasMsg = "M�todo de pagamento usado pelo cliente no marketplace: " & s_mktp_payment
									if Not grava_bloco_notas_pedido(id_pedido, ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_SKYHUB_MKTP_PAYMENT, msg_erro) then
										alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & id_pedido & ")"
										end if
									end if
								end if 'if indice_pedido = 1
							end if 'if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO
						end if 'if alerta = ""
					end if ' if (vEmpresaAutoSplit(iv) <> 0) then
			
				if alerta <> "" then exit for
				next ' for iv = LBound(vEmpresaAutoSplit) to UBound(vEmpresaAutoSplit)
			end if 'if alerta = ""
		
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
				if Trim("" & rs("st_etg_imediata")) = Trim(COD_ETG_IMEDIATA_NAO) then s_log = s_log & " (previs�o de entrega: " & formata_data(rs("PrevisaoEntregaData")) & ")"
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
		
				s_log = s_log & "; Endere�o cobran�a=" & formata_endereco(EndCob_endereco, EndCob_endereco_numero, EndCob_endereco_complemento, EndCob_bairro, EndCob_cidade, EndCob_uf, EndCob_cep)
				if blnUsarMemorizacaoCompletaEnderecos then
					s_log = s_log & _
							" (" & _
							"email=" & EndCob_email & _
							", email_xml=" & EndCob_email_xml & _
							", nome=" & EndCob_nome & _
							", ddd_res=" & EndCob_ddd_res & _
							", tel_res=" & EndCob_tel_res & _
							", ddd_com=" & EndCob_ddd_com & _
							", tel_com=" & EndCob_tel_com & _
							", ramal_com=" & EndCob_ramal_com & _
							", ddd_cel=" & EndCob_ddd_cel & _
							", tel_cel=" & EndCob_tel_cel & _
							", ddd_com_2=" & EndCob_ddd_com_2 & _
							", tel_com_2=" & EndCob_tel_com_2 & _
							", ramal_com_2=" & EndCob_ramal_com_2 & _
							", tipo_pessoa=" & EndCob_tipo_pessoa & _
							", cnpj_cpf=" & EndCob_cnpj_cpf & _
							", contribuinte_icms_status=" & EndCob_contribuinte_icms_status & _
							", produtor_rural_status=" & EndCob_produtor_rural_status & _
							", ie=" & EndCob_ie & _
							", rg=" & EndCob_rg & _
							", contato=" & EndCob_contato & _
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

					'DESCONTO SUPERIOR POR AL�ADA
					s_log = s_log & _
							"; StatusDescontoSuperior=" & CStr(.StatusDescontoSuperior)
					if .StatusDescontoSuperior <> 0 then
						s_log = s_log & _
								"; IdUsuarioDescontoSuperior=" & CStr(.IdUsuarioDescontoSuperior) & _
								"; DataHoraDescontoSuperior=" & formata_data_hora(.DataHoraDescontoSuperior)
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

			if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				if s_log_item_servico <> "" then
					if s_log <> "" then s_log = s_log & chr(13)
					s_log = s_log & "Servi�o(s):"
					s_log = s_log & chr(13)
					s_log = s_log & s_log_item_servico
					end if
				end if

			if s_log <> "" then
				grava_log usuario, loja, id_pedido_base, cliente_selecionado, OP_LOG_PEDIDO_NOVO, s_log
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
<br />
<table class="notPrint" width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>
<table cellSpacing="0">
<tr>
	<% 	if erro_produto_sem_estoque then 
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
	on error resume next

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