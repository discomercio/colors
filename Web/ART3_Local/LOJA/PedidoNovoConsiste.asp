<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  P E D I D O N O V O C O N S I S T E . A S P
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

	class cl_ITEM_CAD_SEMI_AUTO_PED_MAGE_RATEIO_FRETE_ANALISE
		dim sku
		dim qtde_vendida_sku
		dim isProdutoComposto
		dim fabricante_composto
		dim produto_composto
		dim fabricante_item
		dim produto_item
		dim qtde_produto_item
		dim preco_lista_sku
		dim preco_venda_sku
		dim preco_lista_produto_item
		dim preco_venda_produto_item
		dim vl_frete_rateado_produto_item
		dim preco_nf_produto_item
		end class

	class cl_ITEM_CAD_SEMI_AUTO_PED_MAGE_RATEIO_FRETE_CONSOLIDADO
		dim fabricante
		dim produto
		dim qtde_totalizada
		dim preco_lista_totalizado
		dim preco_venda_totalizado
		dim preco_nf_totalizado
		dim preco_lista_medio
		dim preco_venda_medio
		dim preco_nf_medio
		end class

	dim msg_erro
	dim usuario, loja, cliente_selecionado
	dim s, s_value, i, j, n, nColSpan, idx, intColSpan, qtde_estoque_total_disponivel, blnAchou, blnDesativado
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	cliente_selecionado = Trim(request("cliente_selecionado"))
	if (cliente_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)
	
	dim alerta, alerta_aux, alerta_informativo, alerta_informativo_aux
	alerta=""
	alerta_informativo=""

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, tMAP_XML, tMAP_ITEM, tMAP_END_COB, tMAP_END_ETG, tOI, t_CLIENTE, tPL, tPCI
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(t_CLIENTE, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tPCI, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tPL, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim insert_request_guid
	insert_request_guid = Trim(Request.Form("insert_request_guid"))
	if insert_request_guid = "" then insert_request_guid = gera_uid

	dim blnLojaHabilitadaProdCompostoECommerce
	blnLojaHabilitadaProdCompostoECommerce = isLojaHabilitadaProdCompostoECommerce(loja)

	dim blnFlagCadSemiAutoPedMagentoUsarCamposValorMktpDataSource
	blnFlagCadSemiAutoPedMagentoUsarCamposValorMktpDataSource = isActivatedFlagCadSemiAutoPedMagentoUsarCamposValorMktpDataSource

	dim blnFlagCadSemiAutoPedMagentoRateioFreteAutomatico, blnExecutarCadSemiAutoPedMagentoRateioFreteAutomatico
	blnFlagCadSemiAutoPedMagentoRateioFreteAutomatico = isActivatedFlagCadSemiAutoPedMagentoRateioFreteAutomatico
	blnExecutarCadSemiAutoPedMagentoRateioFreteAutomatico = False

	dim vItemCadSemiAutoPedMageRateioFreteAnalise, vItemCadSemiAutoPedMageRateioFreteConsolidado, vlRateioFreteProdCompPrecoListaTotal
	dim vlRateioFretePrecoVendaTotal, vlRateioFretePrecoNfTotal, vlRateioFretePrecoVendaDif, vlRateioFretePrecoNfDif, sinalAjuste, blnAjusteRateioOk
	dim vlRateioFretePrecoAux, vlRateioFretePrecoAtualAux, vlRateioFretePrecoNovoAux, vlRateioFretePrecoMenorDif, vlRateioFretePrecoMenorDifAux

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	s = "SELECT * FROM t_CLIENTE WHERE (id='" & cliente_selecionado & "')"
	if t_CLIENTE.State <> 0 then t_CLIENTE.Close
	t_CLIENTE.open s, cn
	if t_CLIENTE.Eof then
		Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_CADASTRADO)
		end if

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim rb_indicacao, rb_RA, c_indicador, c_perc_RT, c_ped_bonshop
	rb_indicacao = Trim(Request.Form("rb_indicacao"))
	rb_RA = Trim(Request.Form("rb_RA"))
	c_indicador = Trim(Request.Form("c_indicador"))
	c_perc_RT = Trim(Request.Form("c_perc_RT"))
	c_ped_bonshop = Trim(Request.Form("pedBonshop"))

	if c_indicador = "" then c_perc_RT = ""
	
	dim blnTemRA
	blnTemRA = False
	if rb_RA = "S" then blnTemRA = True

	dim rb_selecao_cd, c_id_nfe_emitente_selecao_manual
	rb_selecao_cd = Trim(Request("rb_selecao_cd"))
	c_id_nfe_emitente_selecao_manual = Trim(Request("c_id_nfe_emitente_selecao_manual"))

	dim c_custoFinancFornecTipoParcelamento, c_custoFinancFornecQtdeParcelas, coeficiente
	c_custoFinancFornecTipoParcelamento = Trim(Request.Form("c_custoFinancFornecTipoParcelamento"))
	c_custoFinancFornecQtdeParcelas = Trim(Request.Form("c_custoFinancFornecQtdeParcelas"))
	
	dim c_FlagCadSemiAutoPedMagento_FluxoOtimizado, s_checked
	c_FlagCadSemiAutoPedMagento_FluxoOtimizado = Trim(Request.Form("c_FlagCadSemiAutoPedMagento_FluxoOtimizado"))

	dim EndCob_endereco, EndCob_endereco_numero, EndCob_endereco_complemento, EndCob_endereco_complemento_original_magento, EndCob_endereco_ponto_referencia, EndCob_bairro, EndCob_cidade, EndCob_uf, EndCob_cep
	dim EndCob_email, EndCob_email_xml, EndCob_email_boleto, EndCob_nome, EndCob_tipo_pessoa
	dim EndCob_ddd_res, EndCob_tel_res, EndCob_ddd_com, EndCob_tel_com, EndCob_ramal_com, EndCob_ddd_com_2, EndCob_tel_com_2, EndCob_ramal_com_2, EndCob_ddd_cel, EndCob_tel_cel
	dim EndCob_cnpj_cpf, EndCob_contribuinte_icms_status, EndCob_produtor_rural_status, EndCob_ie, EndCob_rg, EndCob_contato
	dim rb_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_endereco_ponto_referencia
	dim EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs
	dim EndEtg_email, EndEtg_email_xml, EndEtg_nome, EndEtg_ddd_res, EndEtg_tel_res, EndEtg_ddd_com, EndEtg_tel_com, EndEtg_ramal_com
	dim EndEtg_ddd_cel, EndEtg_tel_cel, EndEtg_ddd_com_2, EndEtg_tel_com_2, EndEtg_ramal_com_2
	dim EndEtg_tipo_pessoa, EndEtg_cnpj_cpf, EndEtg_contribuinte_icms_status, EndEtg_produtor_rural_status
	dim EndEtg_ie, EndEtg_rg

	dim s_fabricante, s_produto, s_descricao, s_descricao_html, s_qtde, s_readonly, s_vl_NF_readonly, s_vl_NF
	dim s_preco_lista, s_preco_venda, s_vl_TotalItem, m_TotalItem, m_TotalDestePedido, m_TotalItemComRA, m_TotalServicos
	dim s_campo_focus
	dim m_TotalDestePedidoComRA, s_TotalDestePedidoComRA
	dim intIdx, qtdeColProd, percDescServico, sPercDescServico, sColorPercDescServico, vl_servico_original_price, vl_servico_price, vl_total_produto_magento, vl_total_servico_magento, vl_frete_magento
	dim s_qtde_dias
	
'	OBTÉM PARÂMETROS DE COMISSÃO E DESCONTO
	dim rCD
	set rCD = obtem_perc_max_comissao_e_desconto_por_loja(loja)

'	OBTÉM A RELAÇÃO DE MEIOS DE PAGAMENTO PREFERENCIAIS (QUE FAZEM USO O PERCENTUAL DE COMISSÃO+DESCONTO NÍVEL 2)
	dim rP, vMPN2, strScriptJS_MPN2
	set rP = get_registro_t_parametro(ID_PARAMETRO_PercMaxComissaoEDesconto_Nivel2_MeiosPagto)
	
	strScriptJS_MPN2 = "<script type='text/javascript'>" & chr(13) & _
						"var vMPN2 = new Array();" & chr(13) & _
						"vMPN2[0] = 0;" & chr(13)
	if Trim("" & rP.id) <> "" then
		vMPN2 = Split(rP.campo_texto, ",")
		for i=Lbound(vMPN2) to Ubound(vMPN2)
			vMPN2(i) = Trim("" & vMPN2(i))
			if vMPN2(i) <> "" then
				strScriptJS_MPN2 = strScriptJS_MPN2 & _
									"vMPN2[vMPN2.length] = " & vMPN2(i) & ";" & chr(13)
				end if
			next
		end if
	strScriptJS_MPN2 = strScriptJS_MPN2 & _
						"</script>" & chr(13)
	
	dim perc_max_RT_a_utilizar, perc_max_RT_padrao
	perc_max_RT_padrao = rCD.perc_max_comissao
	perc_max_RT_a_utilizar = perc_max_RT_padrao

	dim strPercMaxRT, strPercMaxRTAlcada1, strPercMaxRTAlcada2, strPercMaxRTAlcada3
	dim strPercMaxComissaoEDesconto, strPercMaxComissaoEDescontoPj, strPercMaxComissaoEDescontoNivel2, strPercMaxComissaoEDescontoNivel2Pj
	dim strPercMaxDescAlcada1Pf, strPercMaxDescAlcada1Pj, strPercMaxDescAlcada2Pf, strPercMaxDescAlcada2Pj, strPercMaxDescAlcada3Pf, strPercMaxDescAlcada3Pj
	strPercMaxRT = formata_perc(rCD.perc_max_comissao)
	strPercMaxComissaoEDesconto = formata_perc(rCD.perc_max_comissao_e_desconto)
	strPercMaxComissaoEDescontoPj = formata_perc(rCD.perc_max_comissao_e_desconto_pj)
	strPercMaxComissaoEDescontoNivel2 = formata_perc(rCD.perc_max_comissao_e_desconto_nivel2)
	strPercMaxComissaoEDescontoNivel2Pj = formata_perc(rCD.perc_max_comissao_e_desconto_nivel2_pj)
	strPercMaxRTAlcada1 = "0"
	strPercMaxDescAlcada1Pf = "0"
	strPercMaxDescAlcada1Pj = "0"
	strPercMaxRTAlcada2 = "0"
	strPercMaxDescAlcada2Pf = "0"
	strPercMaxDescAlcada2Pj = "0"
	strPercMaxRTAlcada3 = "0"
	strPercMaxDescAlcada3Pf = "0"
	strPercMaxDescAlcada3Pj = "0"
	
	if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_1, s_lista_operacoes_permitidas) then
		if rCD.perc_max_comissao_alcada1 > perc_max_RT_a_utilizar then perc_max_RT_a_utilizar = rCD.perc_max_comissao_alcada1
		strPercMaxRTAlcada1 = formata_perc(rCD.perc_max_comissao_alcada1)
		strPercMaxDescAlcada1Pf = formata_perc(rCD.perc_max_comissao_e_desconto_alcada1_pf)
		strPercMaxDescAlcada1Pj = formata_perc(rCD.perc_max_comissao_e_desconto_alcada1_pj)
		end if

	if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_2, s_lista_operacoes_permitidas) then
		if rCD.perc_max_comissao_alcada2 > perc_max_RT_a_utilizar then perc_max_RT_a_utilizar = rCD.perc_max_comissao_alcada2
		strPercMaxRTAlcada2 = formata_perc(rCD.perc_max_comissao_alcada2)
		strPercMaxDescAlcada2Pf = formata_perc(rCD.perc_max_comissao_e_desconto_alcada2_pf)
		strPercMaxDescAlcada2Pj = formata_perc(rCD.perc_max_comissao_e_desconto_alcada2_pj)
		end if

	if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_3, s_lista_operacoes_permitidas) then
		if rCD.perc_max_comissao_alcada3 > perc_max_RT_a_utilizar then perc_max_RT_a_utilizar = rCD.perc_max_comissao_alcada3
		strPercMaxRTAlcada3 = formata_perc(rCD.perc_max_comissao_alcada3)
		strPercMaxDescAlcada3Pf = formata_perc(rCD.perc_max_comissao_e_desconto_alcada3_pf)
		strPercMaxDescAlcada3Pj = formata_perc(rCD.perc_max_comissao_e_desconto_alcada3_pj)
		end if

	dim strPercVlPedidoLimiteRA, percPercVlPedidoLimiteRA
	percPercVlPedidoLimiteRA = obtem_PercVlPedidoLimiteRA()
	strPercVlPedidoLimiteRA = formata_perc(percPercVlPedidoLimiteRA)
	
	dim v_item
	redim v_item(0)
	set v_item(0) = New cl_ITEM_PEDIDO_NOVO
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
				end with
			end if
		next
	
'	CONSISTÊNCIAS
	dim s_nome_cliente, c_mag_cpf_cnpj_identificado, c_mag_installer_document
	dim operacao_origem, c_numero_magento, c_numero_marketplace, c_marketplace_codigo_origem, operationControlTicket, sessionToken, id_magento_api_pedido_xml
	operacao_origem = Trim(Request("operacao_origem"))
	c_mag_installer_document = ""
	c_numero_magento = ""
	c_numero_marketplace = ""
	c_marketplace_codigo_origem = ""
	operationControlTicket = ""
	sessionToken = ""
	id_magento_api_pedido_xml = ""
	s_nome_cliente = ""
	c_mag_cpf_cnpj_identificado = ""

	dim blnMagentoPedidoComIndicador, sListaLojaMagentoPedidoComIndicador, vLoja, rParametro
	dim percCommissionValue, percCommissionDiscount, percComissionPercentage
	dim sIdIndicador, sNomeIndicador, sIdVendedor, sNomeVendedor
	blnMagentoPedidoComIndicador = False
	sListaLojaMagentoPedidoComIndicador = ""
	sIdIndicador = ""
	sNomeIndicador = ""
	sIdVendedor = ""
	sNomeVendedor = ""
	percCommissionValue = 0
	percCommissionDiscount = 0
	percComissionPercentage = 0

	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		c_numero_magento = Trim(Request("c_numero_magento"))
		operationControlTicket = Trim(Request("operationControlTicket"))
		sessionToken = Trim(Request("sessionToken"))
		id_magento_api_pedido_xml = Trim(Request("id_magento_api_pedido_xml"))
		
		If Not cria_recordset_otimista(tMAP_XML, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
		If Not cria_recordset_otimista(tMAP_ITEM, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
		If Not cria_recordset_otimista(tMAP_END_COB, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
		If Not cria_recordset_otimista(tMAP_END_ETG, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

		set rParametro = get_registro_t_parametro(ID_PARAMETRO_MagentoPedidoComIndicadorListaLojaErp)
		sListaLojaMagentoPedidoComIndicador = Trim("" & rParametro.campo_texto)
		if sListaLojaMagentoPedidoComIndicador <> "" then
			vLoja = Split(sListaLojaMagentoPedidoComIndicador, ",")
			for i=LBound(vLoja) to UBound(vLoja)
				if Trim("" & vLoja(i)) = loja then
					'Esta implementação do pedido Magento com indicador é referente ao projeto em Magento 1 da Bonshop
					blnMagentoPedidoComIndicador = True
					exit for
					end if
				next
			end if
		end if

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
				c_numero_marketplace = Trim("" & tMAP_XML("pedido_marketplace"))
				c_marketplace_codigo_origem = Trim("" & tMAP_XML("marketplace_codigo_origem"))
				s_nome_cliente = UCase(ec_dados_formata_nome(tMAP_XML("customer_firstname"), tMAP_XML("customer_middlename"), tMAP_XML("customer_lastname"), 60))
				c_mag_cpf_cnpj_identificado = retorna_so_digitos(Trim("" & tMAP_XML("cpfCnpjIdentificado")))
				c_mag_installer_document = retorna_so_digitos(Trim("" & tMAP_XML("installer_document")))
				percCommissionValue = tMAP_XML("commission_value")
				percCommissionDiscount = tMAP_XML("commission_discount")
				vl_frete_magento = converte_numero(tMAP_XML("shipping_amount")) - converte_numero(tMAP_XML("shipping_discount_amount"))

				if blnMagentoPedidoComIndicador then
					if c_mag_installer_document = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "O pedido Magento nº " & c_numero_magento & " não informa o CPF/CNPJ do indicador!"
					else
						If Not cria_recordset_otimista(tOI, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
						s = "SELECT " & _
								"*" & _
							" FROM t_ORCAMENTISTA_E_INDICADOR" & _
							" WHERE" & _
								" (cnpj_cpf = '" & retorna_so_digitos(c_mag_installer_document) & "')" & _
								" AND (Convert(smallint, loja) = " & loja & ")" & _
								" AND (status = 'A')"
						if tOI.State <> 0 then tOI.Close
						tOI.open s, cn
						if tOI.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "O pedido Magento nº " & c_numero_magento & " especifica o indicador com CPF/CNPJ " & cnpj_cpf_formata(c_mag_installer_document) & " que não foi localizado no banco de dados (loja: " & loja & ")!"
						else
							sIdIndicador = Trim("" & tOI("apelido"))
							sNomeIndicador = Trim("" & tOI("razao_social_nome"))
							sIdVendedor = Trim("" & tOI("vendedor"))
							sNomeVendedor = Trim("" & x_usuario (sIdVendedor))
							
							'VERIFICA SE HÁ MAIS DE UM INDICADOR CADASTRADO
							tOI.MoveNext
							if Not tOI.Eof then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Há mais de um indicador cadastrado com o CPF/CNPJ " & cnpj_cpf_formata(c_mag_installer_document) & " para a loja " & loja
								end if
							end if
						if tOI.State <> 0 then tOI.Close
						set tOI = nothing
						end if
					end if
				end if 'if tMAP_XML.Eof
			
			if alerta = "" then
				s = "SELECT " & _
						"*" & _
					" FROM t_MAGENTO_API_PEDIDO_XML_DECODE_ENDERECO" & _
					" WHERE" & _
						" (id_magento_api_pedido_xml = " & tMAP_XML("id") & ")" & _
						" AND (tipo_endereco = 'COB')"
				if tMAP_END_COB.State <> 0 then tMAP_END_COB.Close
				tMAP_END_COB.open s, cn
				if tMAP_END_COB.Eof then
					alerta = "Falha ao tentar localizar no banco de dados o registro do endereço de cobrança do pedido Magento nº " & c_numero_magento & " (operationControlTicket = " & operationControlTicket & ")"
					end if
				end if

			if alerta = "" then
				s = "SELECT " & _
						"*" & _
					" FROM t_MAGENTO_API_PEDIDO_XML_DECODE_ENDERECO" & _
					" WHERE" & _
						" (id_magento_api_pedido_xml = " & tMAP_XML("id") & ")" & _
						" AND (tipo_endereco = 'ETG')"
				if tMAP_END_ETG.State <> 0 then tMAP_END_ETG.Close
				tMAP_END_ETG.open s, cn
				if tMAP_END_ETG.Eof then
					alerta = "Falha ao tentar localizar no banco de dados o registro do endereço de entrega do pedido Magento nº " & c_numero_magento & " (operationControlTicket = " & operationControlTicket & ")"
					end if
				end if
			end if 'if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO
		end if 'if alerta = ""

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
		EndCob_email_boleto = Trim(Request.Form("EndCob_email_boleto"))
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
		EndCob_email_boleto = Trim("" & t_CLIENTE("email_boleto"))
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
		
		'QUANDO O FLUXO PASSA PELA TELA DE CADASTRO DO CLIENTE, REALIZA TRATAMENTO ADICIONAL P/ CONSIDERAR OS CAMPOS ORIGINAIS DO MAGENTO DE COMPLEMENTO E PONTO DE REFERÊNCIA
		if c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "9" then
			if EndCob_tipo_pessoa = ID_PF then
				EndCob_endereco_ponto_referencia = Trim("" & tMAP_END_ETG("street_detail"))
				EndCob_endereco_complemento_original_magento = Trim("" & tMAP_END_ETG("endereco_complemento"))
				'O COMPLEMENTO DO ENDEREÇO FOI TRUNCADO?
				if (Len(EndCob_endereco_complemento) < Len(EndCob_endereco_complemento_original_magento)) And _
					(Ucase(EndCob_endereco_complemento) = Ucase(Left(EndCob_endereco_complemento_original_magento, Len(EndCob_endereco_complemento)))) then
					EndCob_endereco_complemento = EndCob_endereco_complemento_original_magento
					end if
				end if
			end if
		end if

	rb_end_entrega = Trim(Request.Form("rb_end_entrega"))
	EndEtg_endereco = Trim(Request.Form("EndEtg_endereco"))
	EndEtg_endereco_numero = Trim(Request.Form("EndEtg_endereco_numero"))
	EndEtg_endereco_complemento = Trim(Request.Form("EndEtg_endereco_complemento"))
	EndEtg_endereco_ponto_referencia = Trim(Request.Form("EndEtg_endereco_ponto_referencia"))
	EndEtg_bairro = Trim(Request.Form("EndEtg_bairro"))
	EndEtg_cidade = Trim(Request.Form("EndEtg_cidade"))
	EndEtg_uf = Trim(Request.Form("EndEtg_uf"))
	EndEtg_cep = Trim(Request.Form("EndEtg_cep"))
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

	'QUANDO O FLUXO PASSA PELA TELA DE CADASTRO DO CLIENTE, REALIZA TRATAMENTO ADICIONAL P/ CONSIDERAR OS CAMPOS ORIGINAIS DO MAGENTO DE COMPLEMENTO E PONTO DE REFERÊNCIA
	if c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "9" then
		if (EndCob_tipo_pessoa = ID_PJ) And (rb_end_entrega = "S") then
			if (EndEtg_endereco_ponto_referencia = "") And (Trim("" & tMAP_END_ETG("street_detail")) <> "") then
				EndEtg_endereco_ponto_referencia = Trim("" & tMAP_END_ETG("street_detail"))
				end if
			end if
		end if

	dim s_loja_indicou, s_nome_loja_indicou
	if Session("vendedor_externo") then
		s_loja_indicou=retorna_so_digitos(Trim(request("loja_indicou")))
		s_nome_loja_indicou = ""
		if s_loja_indicou = "" then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Não foi especificada a loja que fez a indicação."
		else
			s = "SELECT * FROM t_LOJA WHERE (loja='" & s_loja_indicou & "')"
			set rs = cn.execute(s)
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = alerta & "Loja " & s_loja_indicou & " não está cadastrada."
			else
				s_nome_loja_indicou = Trim("" & rs("nome"))
				if s_nome_loja_indicou = "" then s_nome_loja_indicou = Trim("" & rs("razao_social"))
				end if
			if rs.State <> 0 then rs.Close
			end if
		end if

	dim r_orcamentista_e_indicador
	dim permite_RA_status
	permite_RA_status = 0
	if alerta = "" then
		if c_indicador <> "" then
			if Not le_orcamentista_e_indicador(c_indicador, r_orcamentista_e_indicador, msg_erro) then
				alerta = "Falha ao recuperar os dados do indicador!<br>" & msg_erro
			else
				if blnMagentoPedidoComIndicador then
					rb_RA = "S"
					permite_RA_status = 1
				else
					if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
						if (Trim("" & tMAP_XML("b2b_type_order")) = COD_MAGENTO_TYPE_ORDER__INSTALLER) And (Trim("" & tMAP_XML("magento_api_versao")) = CStr(VERSAO_API_MAGENTO_V2_REST_JSON)) then
							if vl_frete_magento > 0 then
								rb_RA = "S"
								permite_RA_status = 1
								end if
						else
							if r_orcamentista_e_indicador.permite_RA_status = 0 then rb_RA = "N"
							permite_RA_status = r_orcamentista_e_indicador.permite_RA_status
							end if
					else
						if r_orcamentista_e_indicador.permite_RA_status = 0 then rb_RA = "N"
						permite_RA_status = r_orcamentista_e_indicador.permite_RA_status
						end if
					end if
				end if
			end if
		end if

	if alerta = "" then
		if rb_end_entrega = "" then
			alerta = "Não foi informado se o endereço de entrega é o mesmo do cadastro ou não."
		elseif rb_end_entrega = "S" then
			if EndEtg_endereco = "" then
				alerta="PREENCHA O ENDEREÇO DE ENTREGA."
			elseif Len(EndEtg_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
				alerta="ENDEREÇO DE ENTREGA EXCEDE O TAMANHO MÁXIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(EndEtg_endereco)) & " CARACTERES<br>TAMANHO MÁXIMO: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " CARACTERES"
			elseif EndEtg_endereco_numero = "" then
				alerta="PREENCHA O NÚMERO DO ENDEREÇO DE ENTREGA."
			elseif EndEtg_bairro = "" then
				alerta="PREENCHA O BAIRRO DO ENDEREÇO DE ENTREGA."
			elseif EndEtg_cidade = "" then
				alerta="PREENCHA A CIDADE DO ENDEREÇO DE ENTREGA."
			elseif (EndEtg_uf="") Or (Not uf_ok(EndEtg_uf)) then
				alerta="UF INVÁLIDA NO ENDEREÇO DE ENTREGA."
			elseif Not cep_ok(EndEtg_cep) then
				alerta="CEP INVÁLIDO NO ENDEREÇO DE ENTREGA."
				end if


            if (alerta = "") And (EndCob_tipo_pessoa = ID_PJ) and blnUsarMemorizacaoCompletaEnderecos then
                if EndEtg_tipo_pessoa <> "PJ" and EndEtg_tipo_pessoa <> "PF" then
                    alerta = "Necessário escolher Pessoa Jurídica ou Pessoa Física no Endereço de entrega!"
    			elseif EndEtg_nome = "" then
                    alerta = "Preencha o nome/razão social no endereço de entrega!"
                    end if 
	
                if alerta = "" and EndEtg_tipo_pessoa = "PJ" then
                    '//Campos PJ: 
                    if EndEtg_cnpj_cpf = "" or not cnpj_ok(EndEtg_cnpj_cpf) then
                        alerta = "Endereço de entrega: CNPJ inválido!"
                    elseif EndEtg_contribuinte_icms_status = "" then
                        alerta = "Endereço de entrega: selecione o tipo de contribuinte de ICMS!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and EndEtg_ie = "" then
                        alerta = "Endereço de entrega: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) and InStr(EndEtg_ie, "ISEN") > 0 then 
                        alerta = "Endereço de entrega: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and InStr(EndEtg_ie, "ISEN") > 0 then 
                        alerta = "Endereço de entrega: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!"
                    'telefones PJ:
                    'EndEtg_ddd_com
                    'EndEtg_tel_com
                    'EndEtg_ramal_com
                    'EndEtg_ddd_com_2
                    'EndEtg_tel_com_2
                    'EndEtg_ramal_com_2
                    elseif not ddd_ok(EndEtg_ddd_com) then
                        alerta = "Endereço de entrega: DDD inválido!"
                    elseif not telefone_ok(EndEtg_tel_com) then
                        alerta = "Endereço de entrega: telefone inválido!"
                    elseif EndEtg_ddd_com = "" and EndEtg_tel_com <> "" then
                        alerta = "Endereço de entrega: preencha o DDD do telefone."
                    elseif EndEtg_tel_com = "" and EndEtg_ddd_com <> "" then
                        alerta = "Endereço de entrega: preencha o telefone."

                    elseif not ddd_ok(EndEtg_ddd_com_2) then
                        alerta = "Endereço de entrega: DDD inválido!"
                    elseif not telefone_ok(EndEtg_tel_com_2) then
                        alerta = "Endereço de entrega: telefone inválido!"
                    elseif EndEtg_ddd_com_2 = "" and EndEtg_tel_com_2 <> "" then
                        alerta = "Endereço de entrega: preencha o DDD do telefone."
                    elseif EndEtg_tel_com_2 = "" and EndEtg_ddd_com_2 <> "" then
                        alerta = "Endereço de entrega: preencha o telefone."
                        end if 
                    end if 

                if alerta = "" and EndEtg_tipo_pessoa <> "PJ" then
                    '//campos PF
                    if EndEtg_cnpj_cpf = "" or not cpf_ok(EndEtg_cnpj_cpf) then
                        alerta = "Endereço de entrega: CPF inválido!"
                    elseif EndEtg_produtor_rural_status = "" then
                        alerta = "Endereço de entrega: informe se o cliente é produtor rural ou não!"
                    elseif converte_numero(EndEtg_produtor_rural_status) <> converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
                        if converte_numero(EndEtg_contribuinte_icms_status) <> converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                            alerta = "Endereço de entrega: para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!"
                        elseif EndEtg_contribuinte_icms_status = "" then
                            alerta = "Endereço de entrega: informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!"
                        elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and EndEtg_ie = "" then
                            alerta = "Endereço de entrega: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!"
                        elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) and InStr(EndEtg_ie, "ISEN") > 0 then 
                            alerta = "Endereço de entrega: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!"
                        elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and InStr(EndEtg_ie, "ISEN") > 0 then 
                            alerta = "Endereço de entrega: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!"
                        elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) and EndEtg_ie <> "" then 
                            alerta = "Endereço de entrega: se o Contribuinte ICMS é isento, o campo IE deve ser vazio!"
                            end if
                        end if

                    if alerta = "" then
                        'telefones PF:
                        'EndEtg_ddd_res
                        'EndEtg_tel_res
                        'EndEtg_ddd_cel
                        'EndEtg_tel_cel
                        if not ddd_ok(retorna_so_digitos(EndEtg_ddd_res)) then
                            alerta = "Endereço de entrega: DDD inválido!"
                        elseif not telefone_ok(retorna_so_digitos(EndEtg_tel_res)) then
                            alerta = "Endereço de entrega: telefone inválido!"
                        elseif EndEtg_ddd_res <> "" or EndEtg_tel_res <> "" then
                            if EndEtg_ddd_res = "" then
                                alerta = "Endereço de entrega: preencha o DDD!"
                            elseif EndEtg_tel_res = "" then
                                alerta = "Endereço de entrega: preencha o telefone!"
                                end if
                            end if
                        end if

                    if alerta = "" then
                        if not ddd_ok(retorna_so_digitos(EndEtg_ddd_cel)) then
                            alerta = "Endereço de entrega: DDD inválido!"
                        elseif not telefone_ok(retorna_so_digitos(EndEtg_tel_cel)) then
                            alerta = "Endereço de entrega: telefone inválido!"
                        elseif EndEtg_ddd_cel = "" and EndEtg_tel_cel <> "" then
                            alerta = "Endereço de entrega: preencha o DDD do celular."
                        elseif EndEtg_tel_cel = "" and EndEtg_ddd_cel <> "" then
                            alerta = "Endereço de entrega: preencha o número do celular."
                            end if
                        end if
                    end if
                end if

		    end if
	    end if

	if alerta = "" then
		if ( (EndCob_tipo_pessoa = ID_PF) And (Cstr(EndCob_produtor_rural_status) = Cstr(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM)) ) _
			Or _
			( (EndCob_tipo_pessoa = ID_PJ) And (Cstr(EndCob_contribuinte_icms_status) = Cstr(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM)) ) _
			Or _
			( (EndCob_tipo_pessoa = ID_PJ) And (Cstr(EndCob_contribuinte_icms_status) = Cstr(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO)) And (EndCob_ie <> "") ) then
			if Not isInscricaoEstadualValida(EndCob_ie, EndCob_uf) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Preencha a IE (Inscrição Estadual) com um número válido!" & _
						"<br>" & "Certifique-se de que a UF informada corresponde à UF responsável pelo registro da IE."
				end if
			end if

		if rb_end_entrega = "S" then
			if ( (EndEtg_tipo_pessoa = ID_PF) And (Cstr(EndEtg_produtor_rural_status) = Cstr(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM)) ) _
				Or _
			   ( (EndEtg_tipo_pessoa = ID_PJ) And (Cstr(EndEtg_contribuinte_icms_status) = Cstr(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM)) ) _
			   Or _
			   ( (EndEtg_tipo_pessoa = ID_PJ) And (Cstr(EndEtg_contribuinte_icms_status) = Cstr(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO)) And (Trim(EndEtg_ie) <> "") ) then
				if Not isInscricaoEstadualValida(EndEtg_ie, EndEtg_uf) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Endereço de entrega: preencha a IE (Inscrição Estadual) com um número válido!" & _
							"<br>" & "Certifique-se de que a UF do endereço de entrega corresponde à UF responsável pelo registro da IE."
					end if
				end if
			end if
		end if

	if alerta="" then
	'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
		dim s_lista_sugerida_municipios
		if Not consiste_municipio_IBGE_ok(EndCob_cidade, EndCob_uf, s_lista_sugerida_municipios, msg_erro) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			if msg_erro <> "" then
				alerta = alerta & msg_erro
			else
				alerta = alerta & "Município '" & EndCob_cidade & "' não consta na relação de municípios do IBGE para a UF de '" & EndCob_uf & "'!!"
				if s_lista_sugerida_municipios <> "" then
					alerta = alerta & "<br>" & _
										"Localize o município na lista abaixo e verifique se a grafia está correta!!"
					end if
				end if
			end if
		
		if rb_end_entrega = "S" then
			if Not consiste_municipio_IBGE_ok(EndEtg_cidade, EndEtg_uf, s_lista_sugerida_municipios, msg_erro) then
				if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				if msg_erro <> "" then
					alerta = alerta & msg_erro
				else
					alerta = alerta & "Endereço de entrega: município '" & EndEtg_cidade & "' não consta na relação de municípios do IBGE para a UF de '" & EndEtg_uf & "'!!"
					if s_lista_sugerida_municipios <> "" then
						alerta = alerta & "<br>" & _
											"Localize o município na lista abaixo e verifique se a grafia está correta!!"
						end if
					end if
				end if
			end if
		end if

	if alerta="" then
		if rb_indicacao = "" then
			alerta = "Informe se o pedido é com indicação ou não."
		elseif rb_indicacao = "S" then
			if c_indicador = "" then
				alerta = "Informe quem é o indicador."
			elseif rb_RA = "" then
				alerta = "Informe se o pedido possui RA ou não."
				end if
			end if
		end if
	
	if alerta = "" then
		if c_perc_RT <> "" then
			if (converte_numero(c_perc_RT) < 0) Or (converte_numero(c_perc_RT) > 100) then
				alerta = "Percentual de comissão inválido."
			elseif converte_numero(c_perc_RT) > perc_max_RT_a_utilizar then
				alerta = "O percentual de comissão excede o máximo permitido."
				end if
			end if
		end if
	
'	VERIFICA CADA UM DOS PRODUTOS SELECIONADOS
	if alerta="" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .qtde <= 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": quantidade " & cstr(.qtde) & " é inválida."
					end if

				for j=Lbound(v_item) to (i-1)
					if (.produto = v_item(j).produto) And (.fabricante = v_item(j).fabricante) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": linha " & renumera_com_base1(Lbound(v_item),i) & " repete o mesmo produto da linha " & renumera_com_base1(Lbound(v_item),j) & "."
						exit for
						end if
					next

				s = "SELECT " & _
						"*" & _
					" FROM t_PRODUTO" & _
						" INNER JOIN t_PRODUTO_LOJA" & _
							" ON (t_PRODUTO.fabricante=t_PRODUTO_LOJA.fabricante) AND (t_PRODUTO.produto=t_PRODUTO_LOJA.produto)" & _
					" WHERE" & _
						" (t_PRODUTO.fabricante='" & .fabricante & "')" & _
						" AND (t_PRODUTO.produto='" & .produto & "')" & _
						" AND (loja='" & loja & "')"
				if tPL.State <> 0 then tPL.Close
				tPL.open s, cn
				if tPL.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " NÃO está cadastrado."
				else
					if Ucase(Trim("" & tPL("vendavel"))) <> "S" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " NÃO está disponível para venda."
					elseif .qtde > tPL("qtde_max_venda") then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": quantidade " & cstr(.qtde) & " excede o máximo permitido."
					else
						.preco_lista = tPL("preco_lista")
						.margem = tPL("margem")
						.desc_max = tPL("desc_max")
						.comissao = tPL("comissao")
						.preco_fabricante = tPL("preco_fabricante")
						.vl_custo2 = tPL("vl_custo2")
						.descricao = Trim("" & tPL("descricao"))
						.descricao_html = Trim("" & tPL("descricao_html"))
						.ean = Trim("" & tPL("ean"))
						.grupo = Trim("" & tPL("grupo"))
						.subgrupo = Trim("" & tPL("subgrupo"))
						.peso = tPL("peso")
						.qtde_volumes = Trim("" & tPL("qtde_volumes"))
						.cubagem = tPL("cubagem")
						.ncm = Trim("" & tPL("ncm"))
						.cst = Trim("" & tPL("cst"))
						.descontinuado = Trim("" & tPL("descontinuado"))
						end if
					end if
				tPL.Close
				end with
			next
		end if

	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			s = "SELECT " & _
					"*" & _
				" FROM t_EC_PRODUTO_COMPOSTO_ITEM" & _
				" WHERE" & _
					" (fabricante_composto = '" & v_item(i).fabricante & "')" & _
					" AND (produto_composto = '" & v_item(i).produto & "')" & _
				" ORDER BY" & _
					" fabricante_item," & _
					" produto_item"
			if tPCI.State <> 0 then tPCI.Close
			tPCI.open s, cn
			if Not tPCI.Eof then
				s = ""
				do while Not tPCI.Eof
					if s <> "" then s = s & ", "
					s = s & Trim("" & tPCI("produto_item"))
					tPCI.MoveNext
					loop
				alerta=texto_add_br(alerta)
				alerta=alerta & "O código de produto " & v_item(i).produto & " do fabricante " & v_item(i).fabricante & " é somente um código auxiliar para agrupar os produtos " & s & " e não pode ser usado diretamente no pedido!"
				end if
			next
		end if

	'Regra de negócio: permite vender determinados produtos somente se não excederem determinado limite percentual em relação ao total do pedido
	dim rVCLstProd, rVCPercMax
	dim sVCLstProd, sProdutosVendaCondicionada, qtdeProdutosVendaCondicionada
	dim vl_total_preco_lista, vl_total_venda_condicionada
	if alerta = "" then
		set rVCLstProd = get_registro_t_parametro(ID_PARAMETRO_VendaCondicionada_RegraProporcao_ListaProdutos)
		set rVCPercMax = get_registro_t_parametro(ID_PARAMETRO_VendaCondicionada_RegraProporcao_PercentualMaximoPedido)
		if Trim("" & rVCLstProd.campo_texto) <> "" then
			sVCLstProd = "|" & Trim("" & rVCLstProd.campo_texto) & "|"
			sProdutosVendaCondicionada = ""
			qtdeProdutosVendaCondicionada = 0
			vl_total_preco_lista = 0
			vl_total_venda_condicionada = 0
			for i=Lbound(v_item) to Ubound(v_item)
				with v_item(i)
					if Trim("" & .produto) <> "" then
						vl_total_preco_lista = vl_total_preco_lista + (.qtde * .preco_lista)
						s = "|" & Trim("" & .produto) & "|"
						if Instr(sVCLstProd, s) <> 0 then
							qtdeProdutosVendaCondicionada = qtdeProdutosVendaCondicionada + 1
							vl_total_venda_condicionada = vl_total_venda_condicionada + (.qtde * .preco_lista)
							if sProdutosVendaCondicionada <> "" then sProdutosVendaCondicionada = sProdutosVendaCondicionada & ", "
							sProdutosVendaCondicionada = sProdutosVendaCondicionada & Trim("" & .produto)
							end if
						end if
					end with
				next
			
			if vl_total_preco_lista <> 0 then
				if (qtdeProdutosVendaCondicionada > 0) And ((vl_total_venda_condicionada / vl_total_preco_lista) > (rVCPercMax.campo_real / 100)) then
					alerta=texto_add_br(alerta)
					if qtdeProdutosVendaCondicionada > 1 then
						alerta=alerta & "Os produtos " & sProdutosVendaCondicionada & " não podem ser vendidos neste pedido!"
					else
						alerta=alerta & "O produto " & sProdutosVendaCondicionada & " não pode ser vendido neste pedido!"
						end if
					end if
				end if
			end if
		end if 'if alerta = ""

	if alerta = "" then
		'Calcula o rateio do frete automaticamente?
		'Importante: o frete grátis é um caso específico de valor de frete, pois essa rotina irá
		'=========== calcular 2 valores:
		' 1) Preço de venda do item: irá preencher automaticamente o preço de venda usando o valor
		'    informado pelo Magento/marketplace já considerando os descontos. Se o produto for um
		'    produto composto, irá fazer o rateio do preço de venda com base na proporção do preço
		'    de lista.
		' 2) Preço NF: irá calcular o preço de NF de forma que o valor do frete seja distribuído
		'    entre os itens. No caso de ser um produto composto, também irá fazer o rateio com base
		'    na proporção do preço de lista.
		if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) then
			if blnFlagCadSemiAutoPedMagentoRateioFreteAutomatico _
				And (Trim("" & tMAP_XML("magento_api_versao")) = CStr(VERSAO_API_MAGENTO_V2_REST_JSON)) then
				blnExecutarCadSemiAutoPedMagentoRateioFreteAutomatico = True
				vl_total_produto_magento = 0
				redim vItemCadSemiAutoPedMageRateioFreteAnalise(0)
				set vItemCadSemiAutoPedMageRateioFreteAnalise(0) = New cl_ITEM_CAD_SEMI_AUTO_PED_MAGE_RATEIO_FRETE_ANALISE
				vItemCadSemiAutoPedMageRateioFreteAnalise(Ubound(vItemCadSemiAutoPedMageRateioFreteAnalise)).produto_item = ""
				'Para cada item, verifica se faz parte de um produto composto
				'	1) Se for um produto composto:
				'		1a) Calcula o preço de lista do conjunto
				'		1b) Localiza o preço de venda nos dados do Magento (corresponde ao preço de venda do produto composto)
				'		1c) Calcula o preço de venda de cada parte da composição com base na proporção do preço de lista
				'		1d) Calcula o rateio do frete com base na proporção do preço de venda do item em relação ao preço total de venda do pedido
				'	2) Se for um produto normal
				'		2a) Armazena o preço de lista do produto
				'		2b) Localiza o preço de venda nos dados do Magento
				'		2c) Calcula o rateio do frete com base na proporção do preço de venda do item em relação ao preço total de venda do pedido

				s = "SELECT " & _
						"tMAP_ITEM.*" & _
					" FROM t_MAGENTO_API_PEDIDO_XML tMAP" & _
						" INNER JOIN t_MAGENTO_API_PEDIDO_XML_DECODE_ITEM tMAP_ITEM ON (tMAP.id = tMAP_ITEM.id_magento_api_pedido_xml)" & _
					" WHERE" & _
						" (tMAP_ITEM.id_magento_api_pedido_xml = " & id_magento_api_pedido_xml & ")" & _
						" AND (" & _
							"(tMAP.magento_api_versao = " & VERSAO_API_MAGENTO_V2_REST_JSON & ") AND (tMAP_ITEM.product_type IN ('" & COD_MAGENTO_PRODUCT_TYPE__SIMPLE & "'))" & _
							")" & _
					" ORDER BY" & _
						" tMAP_ITEM.id"
				if tMAP_ITEM.State <> 0 then tMAP_ITEM.Close
				tMAP_ITEM.open s, cn
				do while Not tMAP_ITEM.Eof
					'Verifica se é um produto composto
					s = "SELECT" & _
							" t_EC_PRODUTO_COMPOSTO_ITEM.*" & _
							", Coalesce((SELECT TOP 1 preco_lista FROM t_PRODUTO_LOJA WHERE (fabricante=t_EC_PRODUTO_COMPOSTO_ITEM.fabricante_item) AND (produto=t_EC_PRODUTO_COMPOSTO_ITEM.produto_item) AND (loja='" & loja & "')), 0) AS preco_lista" & _
						" FROM t_EC_PRODUTO_COMPOSTO_ITEM" & _
						" WHERE" & _
							" (produto_composto = '" & normaliza_codigo(Trim("" & tMAP_ITEM("sku")), TAM_MIN_PRODUTO) & "')" & _
							" AND (excluido_status = 0)" & _
						" ORDER BY" & _
							" fabricante_item," & _
							" produto_item"
					if tPCI.State <> 0 then tPCI.Close
					tPCI.open s, cn
					if Not tPCI.Eof then
						'É produto composto!
						'Passo 1) Calcula o preço de lista total
						vlRateioFreteProdCompPrecoListaTotal = 0
						do while Not tPCI.Eof
							vlRateioFreteProdCompPrecoListaTotal = vlRateioFreteProdCompPrecoListaTotal + (converte_numero(tPCI("qtde")) * converte_numero(tPCI("preco_lista")))
							tPCI.MoveNext
							loop

						'Passo 2) Calcula o preço de venda unitário rateado de cada item da composição com base na proporção do preço de lista
						if vlRateioFreteProdCompPrecoListaTotal > 0 then
							tPCI.MoveFirst
							do while Not tPCI.Eof
								if vItemCadSemiAutoPedMageRateioFreteAnalise(Ubound(vItemCadSemiAutoPedMageRateioFreteAnalise)).produto_item <> "" then
									redim preserve vItemCadSemiAutoPedMageRateioFreteAnalise(Ubound(vItemCadSemiAutoPedMageRateioFreteAnalise)+1)
									set vItemCadSemiAutoPedMageRateioFreteAnalise(Ubound(vItemCadSemiAutoPedMageRateioFreteAnalise)) = New cl_ITEM_CAD_SEMI_AUTO_PED_MAGE_RATEIO_FRETE_ANALISE
									end if

								with vItemCadSemiAutoPedMageRateioFreteAnalise(Ubound(vItemCadSemiAutoPedMageRateioFreteAnalise))
									.sku = normaliza_codigo(Trim("" & tMAP_ITEM("sku")), TAM_MIN_PRODUTO)
									.qtde_vendida_sku = converte_numero(tMAP_ITEM("qty_ordered"))
									.isProdutoComposto = True
									.fabricante_composto = Trim("" & tPCI("fabricante_composto"))
									.produto_composto = Trim("" & tPCI("produto_composto"))
									.fabricante_item = Trim("" & tPCI("fabricante_item"))
									.produto_item = Trim("" & tPCI("produto_item"))
									.qtde_produto_item = converte_numero(tPCI("qtde"))
									.preco_lista_sku = vlRateioFreteProdCompPrecoListaTotal
									if blnFlagCadSemiAutoPedMagentoUsarCamposValorMktpDataSource And (tMAP_XML("mktp_datasource_status") = 1) then
										'O campo mktp_datasource_special_price informa o valor unitário do item já contabilizando o desconto
										.preco_venda_sku = converte_numero(tMAP_ITEM("mktp_datasource_special_price"))
									else
										'O campo row_total informa o valor total do item já calculado com os descontos e multiplicado pela quantidade
										.preco_venda_sku = converte_numero(tMAP_ITEM("row_total")) / .qtde_vendida_sku
										end if
									.preco_lista_produto_item = converte_numero(tPCI("preco_lista"))
									.preco_venda_produto_item = .preco_venda_sku * (.preco_lista_produto_item / .preco_lista_sku)
									end with

								tPCI.MoveNext
								loop
							end if 'if vlRateioFreteProdCompPrecoListaTotal > 0

					else
						'É produto normal!
						s = "SELECT" & _
								" t_PRODUTO.*" & _
								", Coalesce((SELECT TOP 1 preco_lista FROM t_PRODUTO_LOJA WHERE (fabricante=t_PRODUTO.fabricante) AND (produto=t_PRODUTO.produto) AND (loja='" & loja & "')), 0) AS preco_lista" & _
							" FROM t_PRODUTO" & _
							" WHERE" & _
								" (produto = '" & normaliza_codigo(Trim("" & tMAP_ITEM("sku")), TAM_MIN_PRODUTO) & "')" & _
								" AND (excluido_status = 0)"
						if tPL.State <> 0 then tPL.Close
						tPL.open s, cn
						if Not tPL.Eof then
							if vItemCadSemiAutoPedMageRateioFreteAnalise(Ubound(vItemCadSemiAutoPedMageRateioFreteAnalise)).produto_item <> "" then
								redim preserve vItemCadSemiAutoPedMageRateioFreteAnalise(Ubound(vItemCadSemiAutoPedMageRateioFreteAnalise)+1)
								set vItemCadSemiAutoPedMageRateioFreteAnalise(Ubound(vItemCadSemiAutoPedMageRateioFreteAnalise)) = New cl_ITEM_CAD_SEMI_AUTO_PED_MAGE_RATEIO_FRETE_ANALISE
								end if

							with vItemCadSemiAutoPedMageRateioFreteAnalise(Ubound(vItemCadSemiAutoPedMageRateioFreteAnalise))
								.sku = normaliza_codigo(Trim("" & tMAP_ITEM("sku")), TAM_MIN_PRODUTO)
								.qtde_vendida_sku = converte_numero(tMAP_ITEM("qty_ordered"))
								.isProdutoComposto = False
								.fabricante_composto = ""
								.produto_composto = ""
								.fabricante_item = Trim("" & tPL("fabricante"))
								.produto_item = Trim("" & tPL("produto"))
								.qtde_produto_item = 1
								.preco_lista_sku = converte_numero(tPL("preco_lista"))
								if blnFlagCadSemiAutoPedMagentoUsarCamposValorMktpDataSource And (tMAP_XML("mktp_datasource_status") = 1) then
									'O campo mktp_datasource_special_price informa o valor unitário do item já contabilizando o desconto
									.preco_venda_sku = converte_numero(tMAP_ITEM("mktp_datasource_special_price"))
								else
									'O campo row_total informa o valor total do item já calculado com os descontos e multiplicado pela quantidade
									.preco_venda_sku = converte_numero(tMAP_ITEM("row_total")) / .qtde_vendida_sku
									end if
								.preco_lista_produto_item = converte_numero(tPL("preco_lista"))
								.preco_venda_produto_item = .preco_venda_sku
								end with
							end if 'if Not tPL.Eof
						end if 'if Not tPCI.Eof

					if blnFlagCadSemiAutoPedMagentoUsarCamposValorMktpDataSource And (tMAP_XML("mktp_datasource_status") = 1) then
						'O campo mktp_datasource_special_price informa o valor unitário do item já contabilizando o desconto
						vl_total_produto_magento = vl_total_produto_magento + (converte_numero(tMAP_ITEM("qty_ordered")) * converte_numero(tMAP_ITEM("mktp_datasource_special_price")))
					else
						'O campo row_total informa o valor total do item já calculado com os descontos e multiplicado pela quantidade
						vl_total_produto_magento = vl_total_produto_magento + converte_numero(tMAP_ITEM("row_total"))
						end if

					tMAP_ITEM.MoveNext
					loop

				for i=LBound(vItemCadSemiAutoPedMageRateioFreteAnalise) to UBound(vItemCadSemiAutoPedMageRateioFreteAnalise)
					if vItemCadSemiAutoPedMageRateioFreteAnalise(i).produto_item <> "" then
						with vItemCadSemiAutoPedMageRateioFreteAnalise(i)
							.vl_frete_rateado_produto_item = vl_frete_magento * (.preco_venda_produto_item / vl_total_produto_magento)
							.preco_nf_produto_item = .preco_venda_produto_item + .vl_frete_rateado_produto_item
							end with
						end if
					next

				'Como no produto composto podem haver itens que são compartilhados entre vários outros produtos compostos, o processamento
				'até agora criou uma entrada no vetor p/ cada item da composição, mesmo que ele já pudesse existir devido a outro produto composto
				'já processado anteriormente.
				'Agora realiza uma consolidação de forma que cada código de produto seja único e os valores representem a média aritmética
				redim vItemCadSemiAutoPedMageRateioFreteConsolidado(0)
				set vItemCadSemiAutoPedMageRateioFreteConsolidado(0) = New cl_ITEM_CAD_SEMI_AUTO_PED_MAGE_RATEIO_FRETE_CONSOLIDADO
				vItemCadSemiAutoPedMageRateioFreteConsolidado(UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)).produto = ""
				for i=LBound(vItemCadSemiAutoPedMageRateioFreteAnalise) to UBound(vItemCadSemiAutoPedMageRateioFreteAnalise)
					if vItemCadSemiAutoPedMageRateioFreteAnalise(i).produto_item <> "" then
						blnAchou = False
						for j=LBound(vItemCadSemiAutoPedMageRateioFreteConsolidado) to UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)
							if (vItemCadSemiAutoPedMageRateioFreteAnalise(i).fabricante_item = vItemCadSemiAutoPedMageRateioFreteConsolidado(j).fabricante) _
								And (vItemCadSemiAutoPedMageRateioFreteAnalise(i).produto_item = vItemCadSemiAutoPedMageRateioFreteConsolidado(j).produto) then
									blnAchou = True
									idx = j
									exit for
									end if
							next

						if Not blnAchou then
							if vItemCadSemiAutoPedMageRateioFreteConsolidado(UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)).produto <> "" then
								redim preserve vItemCadSemiAutoPedMageRateioFreteConsolidado(UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)+1)
								set vItemCadSemiAutoPedMageRateioFreteConsolidado(UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)) = New cl_ITEM_CAD_SEMI_AUTO_PED_MAGE_RATEIO_FRETE_CONSOLIDADO
								end if
							idx = UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)
							vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).fabricante = vItemCadSemiAutoPedMageRateioFreteAnalise(i).fabricante_item
							vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).produto = vItemCadSemiAutoPedMageRateioFreteAnalise(i).produto_item
							vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).qtde_totalizada = 0
							vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).preco_lista_totalizado = 0
							vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).preco_venda_totalizado = 0
							vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).preco_nf_totalizado = 0
							end if

						n = vItemCadSemiAutoPedMageRateioFreteAnalise(i).qtde_vendida_sku * vItemCadSemiAutoPedMageRateioFreteAnalise(i).qtde_produto_item
						vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).qtde_totalizada = vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).qtde_totalizada + n
						vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).preco_lista_totalizado = vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).preco_lista_totalizado + (n * vItemCadSemiAutoPedMageRateioFreteAnalise(i).preco_lista_produto_item)
						vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).preco_venda_totalizado = vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).preco_venda_totalizado + (n * vItemCadSemiAutoPedMageRateioFreteAnalise(i).preco_venda_produto_item)
						vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).preco_nf_totalizado = vItemCadSemiAutoPedMageRateioFreteConsolidado(idx).preco_nf_totalizado + (n * vItemCadSemiAutoPedMageRateioFreteAnalise(i).preco_nf_produto_item)
						end if
					next

				'Calcula valores médios
				for i=LBound(vItemCadSemiAutoPedMageRateioFreteConsolidado) to UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)
					with vItemCadSemiAutoPedMageRateioFreteConsolidado(i)
						if (.produto <> "") And (.qtde_totalizada <> 0) then
							.preco_lista_medio = converte_numero(formata_moeda(.preco_lista_totalizado / .qtde_totalizada))
							.preco_venda_medio = converte_numero(formata_moeda(.preco_venda_totalizado / .qtde_totalizada))
							.preco_nf_medio = converte_numero(formata_moeda(.preco_nf_totalizado / .qtde_totalizada))
							end if
						end with
					next

				'Verifica se há necessidade de ajustes para chegar no valor total exato
				vlRateioFretePrecoVendaTotal = 0
				vlRateioFretePrecoNfTotal = 0
				for i=LBound(vItemCadSemiAutoPedMageRateioFreteConsolidado) to UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)
					with vItemCadSemiAutoPedMageRateioFreteConsolidado(i)
						if (.produto <> "") And (.qtde_totalizada <> 0) then
							vlRateioFretePrecoVendaTotal = vlRateioFretePrecoVendaTotal + (.qtde_totalizada * .preco_venda_medio)
							vlRateioFretePrecoNfTotal = vlRateioFretePrecoNfTotal + (.qtde_totalizada * .preco_nf_medio)
							end if
						end with
					next

				vlRateioFretePrecoVendaDif = vl_total_produto_magento - vlRateioFretePrecoVendaTotal
				vlRateioFretePrecoNfDif = (vl_total_produto_magento + vl_frete_magento) - vlRateioFretePrecoNfTotal

				'Ajustar preço de venda?
				blnAjusteRateioOk = False
				if vlRateioFretePrecoVendaDif <> 0 then
					if vlRateioFretePrecoVendaDif > 0 then
						sinalAjuste = 1
					else
						sinalAjuste = -1
						end if

					'Tenta localizar um item em que a quantidade seja divisível pelo valor da diferença (obs: para o Mod funcionar neste caso, o dividendo não pode conter a parte dos centavos, por isso a multiplicação por 100)
					for i=LBound(vItemCadSemiAutoPedMageRateioFreteConsolidado) to UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)
						with vItemCadSemiAutoPedMageRateioFreteConsolidado(i)
							if (.produto <> "") And (.qtde_totalizada <> 0) then
								if ((100 * Abs(vlRateioFretePrecoVendaDif)) Mod .qtde_totalizada) = 0 then
									blnAjusteRateioOk = True
									.preco_venda_medio = .preco_venda_medio + (sinalAjuste * (Abs(vlRateioFretePrecoVendaDif) / .qtde_totalizada))
									exit for
									end if
								end if
							end with
						next
					
					'Se não conseguiu ajustar, tenta localizar o item que irá ter a menor diferença possível
					'Lembrando que há casos em que pode ser impossível zerar a diferença devido à forma como o desconto foi aplicado e as quantidades dos itens
					if Not blnAjusteRateioOk then
						idx = -1
						vlRateioFretePrecoMenorDif = Abs(vlRateioFretePrecoVendaDif)
						for i=LBound(vItemCadSemiAutoPedMageRateioFreteConsolidado) to UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)
							with vItemCadSemiAutoPedMageRateioFreteConsolidado(i)
								if (.produto <> "") And (.qtde_totalizada <> 0) then
									vlRateioFretePrecoAux = converte_numero(formata_moeda(.preco_venda_medio + (sinalAjuste * (Abs(vlRateioFretePrecoVendaDif) / .qtde_totalizada))))
									vlRateioFretePrecoAtualAux = .qtde_totalizada * .preco_venda_medio
									vlRateioFretePrecoNovoAux = .qtde_totalizada * vlRateioFretePrecoAux
									vlRateioFretePrecoMenorDifAux = Abs(vlRateioFretePrecoAtualAux - vlRateioFretePrecoNovoAux)
									if vlRateioFretePrecoMenorDifAux < vlRateioFretePrecoMenorDif then
										vlRateioFretePrecoMenorDif = vlRateioFretePrecoMenorDifAux
										idx = i
										end if
									end if
								end with
							next
						
						if idx > -1 then
							with vItemCadSemiAutoPedMageRateioFreteConsolidado(idx)
								.preco_venda_medio = .preco_venda_medio + (sinalAjuste * (Abs(vlRateioFretePrecoVendaDif) / .qtde_totalizada))
								end with
							end if
						end if
					end if 'if vlRateioFretePrecoVendaDif <> 0

				'Ajustar preço de NF?
				blnAjusteRateioOk = False
				if vlRateioFretePrecoNfDif <> 0 then
					if vlRateioFretePrecoNfDif > 0 then
						sinalAjuste = 1
					else
						sinalAjuste = -1
						end if

					'Tenta localizar um item em que a quantidade seja divisível pelo valor da diferença (obs: para o Mod funcionar neste caso, o dividendo não pode conter a parte dos centavos, por isso a multiplicação por 100)
					for i=LBound(vItemCadSemiAutoPedMageRateioFreteConsolidado) to UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)
						with vItemCadSemiAutoPedMageRateioFreteConsolidado(i)
							if (.produto <> "") And (.qtde_totalizada <> 0) then
								if ((100 * Abs(vlRateioFretePrecoNfDif)) Mod .qtde_totalizada) = 0 then
									blnAjusteRateioOk = True
									.preco_nf_medio = .preco_nf_medio + (sinalAjuste * (Abs(vlRateioFretePrecoNfDif) / .qtde_totalizada))
									exit for
									end if
								end if
							end with
						next
					
					'Se não conseguiu ajustar, tenta localizar o item que irá ter a menor diferença possível
					'Lembrando que há casos em que pode ser impossível zerar a diferença devido à forma como o desconto foi aplicado e as quantidades dos itens
					if Not blnAjusteRateioOk then
						idx = -1
						vlRateioFretePrecoMenorDif = Abs(vlRateioFretePrecoNfDif)
						for i=LBound(vItemCadSemiAutoPedMageRateioFreteConsolidado) to UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)
							with vItemCadSemiAutoPedMageRateioFreteConsolidado(i)
								if (.produto <> "") And (.qtde_totalizada <> 0) then
									vlRateioFretePrecoAux = converte_numero(formata_moeda(.preco_nf_medio + (sinalAjuste * (Abs(vlRateioFretePrecoNfDif) / .qtde_totalizada))))
									vlRateioFretePrecoAtualAux = .qtde_totalizada * .preco_nf_medio
									vlRateioFretePrecoNovoAux = .qtde_totalizada * vlRateioFretePrecoAux
									vlRateioFretePrecoMenorDifAux = Abs(vlRateioFretePrecoAtualAux - vlRateioFretePrecoNovoAux)
									if vlRateioFretePrecoMenorDifAux < vlRateioFretePrecoMenorDif then
										vlRateioFretePrecoMenorDif = vlRateioFretePrecoMenorDifAux
										idx = i
										end if
									end if
								end with
							next
						
						if idx > -1 then
							with vItemCadSemiAutoPedMageRateioFreteConsolidado(idx)
								.preco_nf_medio = .preco_nf_medio + (sinalAjuste * (Abs(vlRateioFretePrecoNfDif) / .qtde_totalizada))
								end with
							end if
						end if
					end if 'if vlRateioFretePrecoNfDif <> 0
				end if 'if blnFlagCadSemiAutoPedMagentoRateioFreteAutomatico And (Trim("" & tMAP_XML("magento_api_versao")) = CStr(VERSAO_API_MAGENTO_V2_REST_JSON))
			end if 'if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO)
		end if 'if alerta = ""


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
		
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
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
					set rs = cn.execute(s)
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Opção de parcelamento não disponível para fornecedor " & .fabricante & ": " & decodificaCustoFinancFornecQtdeParcelas(c_custoFinancFornecTipoParcelamento, c_custoFinancFornecQtdeParcelas) & " parcela(s)"
					else
						coeficiente = converte_numero(rs("coeficiente"))
						.preco_lista=converte_numero(formata_moeda(coeficiente*.preco_lista))
						end if
					end if
				end with
			next
		end if
		
	dim s_caracteres_invalidos
	if alerta = "" then
		if Not isTextoValido(EndEtg_endereco, s_caracteres_invalidos) then
			alerta="O CAMPO 'ENDEREÇO DE ENTREGA' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_endereco_numero, s_caracteres_invalidos) then
			alerta="O CAMPO NÚMERO DO ENDEREÇO DE ENTREGA POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_endereco_complemento, s_caracteres_invalidos) then
			alerta="O CAMPO COMPLEMENTO DO ENDEREÇO DE ENTREGA POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_bairro, s_caracteres_invalidos) then
			alerta="O CAMPO BAIRRO DO ENDEREÇO DE ENTREGA POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_cidade, s_caracteres_invalidos) then
			alerta="O CAMPO CIDADE DO ENDEREÇO DE ENTREGA POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_nome, s_caracteres_invalidos) then
			alerta="O CAMPO NOME DO ENDEREÇO DE ENTREGA POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
			end if
		end if
	
	
'	LÓGICA P/ CONSUMO DO ESTOQUE (REGRA DEFINIDA POR PRODUTO)
	dim tipo_pessoa
	dim descricao_tipo_pessoa
	tipo_pessoa = multi_cd_regra_determina_tipo_pessoa(EndCob_tipo_pessoa, EndCob_contribuinte_icms_status, EndCob_produtor_rural_status)
	descricao_tipo_pessoa = descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa)

	dim id_nfe_emitente_selecao_manual
	dim vProdRegra, iRegra, iCD, iItem, idxItem, qtde_CD_ativo
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
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " está bloqueada para a UF '" & EndCob_uf & "'"
				elseif vProdRegra(iRegra).regra.regraUF.regraPessoa.st_inativo = 1 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " está bloqueada para clientes '" & descricao_tipo_pessoa & "' da UF '" & EndCob_uf & "'"
				elseif converte_numero(vProdRegra(iRegra).regra.regraUF.regraPessoa.spe_id_nfe_emitente) = 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " não especifica nenhum CD para aguardar produtos sem presença no estoque para clientes '" & descricao_tipo_pessoa & "' da UF '" & EndCob_uf & "'"
				else
					qtde_CD_ativo = 0
					for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
						if converte_numero(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente) > 0 then
							if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0 then
								qtde_CD_ativo = qtde_CD_ativo + 1
								end if
							end if
						next
					'A SELEÇÃO MANUAL DE CD PERMITE O USO DE CD DESATIVADO
					if (qtde_CD_ativo = 0) And (id_nfe_emitente_selecao_manual = 0) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " não especifica nenhum CD ativo para clientes '" & descricao_tipo_pessoa & "' da UF '" & EndCob_uf & "'"
						end if
					end if
				end if
			next
		end if 'if alerta=""
	
	'NO CASO DE SELEÇÃO MANUAL DO CD, VERIFICA SE O CD SELECIONADO ESTÁ HABILITADO EM TODAS AS REGRAS
	if alerta="" then
		if id_nfe_emitente_selecao_manual <> 0 then
			alerta_aux = ""
			alerta_informativo_aux = ""
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
					'16/09/2017: FOI REALIZADA UMA ALTERAÇÃO P/ QUE A SELEÇÃO MANUAL DE CD PERMITA O USO DE CD DESATIVADO
					alerta_informativo_aux = "Regra '" & vProdRegra(iRegra).regra.apelido & "' (Id=" & vProdRegra(iRegra).regra.id & ") define o CD '" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_selecao_manual) & "' como 'desativado'"
					if Instr(alerta_informativo, alerta_informativo_aux) = 0 then
						alerta_informativo=texto_add_br(alerta_informativo)
						alerta_informativo=alerta_informativo & alerta_informativo_aux
						end if
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
						'IMPORTANTE: A SELEÇÃO MANUAL DE CD PERMITE O USO DE CD DESATIVADO
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
								'IMPORTANTE: A SELEÇÃO MANUAL DE CD PERMITE O USO DE CD DESATIVADO
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
								'IMPORTANTE: A SELEÇÃO MANUAL DE CD PERMITE O USO DE CD DESATIVADO
								if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0) Or (id_nfe_emitente_selecao_manual <> 0) then
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
								qtde_empresa_selecionada = qtde_empresa_selecionada + 1
								lista_empresa_selecionada = lista_empresa_selecionada & s
								end if
							end if
						end if
					next
				end if
			next
		end if 'if alerta=""


'	HÁ ALGUM PRODUTO DESCONTINUADO?
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			if Trim(v_item(i).produto) <> "" then
				if Ucase(Trim(v_item(i).descontinuado)) = "S" then
					if v_item(i).qtde > v_item(i).qtde_estoque_total_disponivel then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto (" & v_item(i).fabricante & ")" & v_item(i).produto & " consta como 'descontinuado' e não há mais saldo suficiente no estoque para atender à quantidade solicitada."
						end if
					end if
				end if
			next
		end if
	
	'TEM RA?
	dim strPercLimiteRASemDesagio, strPercDesagio
	strPercLimiteRASemDesagio = "0"
	strPercDesagio = "0"
	if alerta = "" then
		if rb_RA = "S" then
			strPercLimiteRASemDesagio = formata_perc(obtem_perc_limite_RA_sem_desagio())
			strPercDesagio = formata_perc(obtem_perc_desagio_RA_do_indicador(c_indicador))
			end if
		end if
		
'	HÁ MENSAGENS DE ALERTA SOBRE OS PRODUTOS P/ SEREM EXIBIDAS?
	dim strScriptMsgAlerta
	dim strMensagem
	strScriptMsgAlerta = _
		"<script language='JavaScript'>" & chr(13) & _
		"var Pd = new Array();" & chr(13) & _
		"Pd[0] = new oPd('','','','');" & chr(13)

	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			s = "SELECT" & _
					" tpa.fabricante," & _
					" tpa.produto," & _
					" mensagem," & _
					" descricao" & _
				" FROM t_PRODUTO_X_ALERTA tpa" & _
					" INNER JOIN t_ALERTA_PRODUTO tap ON (tpa.id_alerta=tap.apelido)" & _
					" INNER JOIN t_PRODUTO tp ON (tpa.fabricante = tp.fabricante) AND (tpa.produto = tp.produto)" & _
				" WHERE" & _
					" (tpa.fabricante = '" & v_item(i).fabricante & "')" & _
					" AND (tpa.produto = '" & v_item(i).produto & "')" & _
					" AND (tap.ativo = 'S')" & _
				" ORDER BY" & _
					" tpa.dt_cadastro," & _
					" tpa.id_alerta"
			set rs = cn.execute(s)
			do while Not rs.Eof
				strMensagem=Trim("" & rs("mensagem"))
				strMensagem=Replace(strMensagem, chr(10), "")
				strMensagem=Replace(strMensagem, chr(13), "\n")
				strScriptMsgAlerta = strScriptMsgAlerta & _
					"Pd[Pd.length]=new oPd('" & Trim("" & rs("fabricante")) & "'" & _
					",'" & Trim("" & rs("produto")) & "'" & _
					",'" & filtra_nome_identificador(Trim("" & rs("descricao"))) & "'" & _
					",'" & filtra_nome_identificador(strMensagem) & "'" & _
					");" & chr(13)
				rs.MoveNext
				loop
			next
		end if
		
	strScriptMsgAlerta = strScriptMsgAlerta & _
		"</script>" & chr(13)

	dim bloquear_cadastramento_quando_produto_indiponivel
	bloquear_cadastramento_quando_produto_indiponivel = False
	if ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then bloquear_cadastramento_quando_produto_indiponivel = False
	
	dim strScriptJS
	if (Cstr(loja) = Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE)) Or blnMagentoPedidoComIndicador then
		strScriptJS = "<script language='JavaScript'>" & chr(13) & _
					  "var PERC_DESAGIO_RA_LIQUIDA_PEDIDO = " & js_formata_numero(0) & ";" & chr(13)
	else
		strScriptJS = "<script language='JavaScript'>" & chr(13) & _
					  "var PERC_DESAGIO_RA_LIQUIDA_PEDIDO = " & js_formata_numero(getParametroPercDesagioRALiquida) & ";" & chr(13)
		end if

	if erro_produto_indisponivel then
		strScriptJS = strScriptJS & "var erro_produto_indisponivel = true;" & chr(13)
	else
		strScriptJS = strScriptJS & "var erro_produto_indisponivel = false;" & chr(13)
		end if
	if bloquear_cadastramento_quando_produto_indiponivel then
		strScriptJS = strScriptJS & "var bloquear_cadastramento_quando_produto_indiponivel = true;" & chr(13)
	else
		strScriptJS = strScriptJS & "var bloquear_cadastramento_quando_produto_indiponivel = false;" & chr(13)
		end if

	if blnLojaHabilitadaProdCompostoECommerce then
		strScriptJS = strScriptJS & "var formata_perc_desconto = formata_perc_2dec;" & chr(13)
	else
		'Devido à implementação do campo "Desc Linear (%)", a precisão do campo desconto foi alterada p/ 2 decimais
		strScriptJS = strScriptJS & "var formata_perc_desconto = formata_perc_2dec;" & chr(13)
		end if
	
	if blnTemRA then s = "true" else s = "false"
	strScriptJS = strScriptJS & _
				  "var formata_perc_desc_linear = formata_perc_2dec;" & chr(13) & _
				  "var blnTemRA = " & s & ";" & chr(13)

	strScriptJS = strScriptJS & _
				  "</script>" & chr(13)



' FUNÇÕES
' _____________________________________________
' ORIGEM_PEDIDO_MONTA_ITENS_SELECT
'
function origem_pedido_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo='PedidoECommerce_Origem') AND (st_inativo=0) ORDER BY ordenacao")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("codigo"))
		if (id_default=x) then
			strResp = strResp & "<option selected"
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
	
    strResp = "<option value=''>&nbsp;</option>" & strResp

	origem_pedido_monta_itens_select = strResp
	r.close
	set r=nothing
end function

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
	<title>LOJA</title>
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<%=strScriptJS%>
<%=strScriptJS_MPN2%>

<script type="text/javascript">
	$(function() {
		$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8
		<% if EndCob_tipo_pessoa = ID_PF then %>
		if (($("#c_loja").val() != NUMERO_LOJA_ECOMMERCE_AR_CLUBE) && (!FLAG_MAGENTO_PEDIDO_COM_INDICADOR)) $(".TR_FP_PU").hide();
		$(".TR_FP_PSE").hide();
		<% end if %>
		<% if EndCob_tipo_pessoa = ID_PJ then %>
		$(".TR_FP_PSE").hide();
		<% end if %>
		$(".tdGarInd").hide();
		$(".rbGarIndNao").attr('checked', 'checked');
		$("#c_data_previsao_entrega").hUtilUI('datepicker_padrao');

        $("input[name = 'rb_etg_imediata']").change(function () {
			if ($("input[name='rb_etg_imediata']:checked").val() == '<%=COD_ETG_IMEDIATA_NAO%>') {
				$("#c_data_previsao_entrega").prop("readonly", false);
				$("#c_data_previsao_entrega").prop("disabled", false);
                $("#c_data_previsao_entrega").datepicker("enable");
			}
			else {
				$("#c_data_previsao_entrega").val("");
                $("#c_data_previsao_entrega").prop("readonly", true);
				$("#c_data_previsao_entrega").prop("disabled", true);
				$("#c_data_previsao_entrega").datepicker("disable");
            }
		});

		recalcula_total_todas_linhas();
		recalcula_RA();
		recalcula_RA_Liquido();
		recalcula_parcelas();
	});

	//Every resize of window
	$(window).resize(function() {
		sizeDivAjaxRunning();
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
</script>

<script language="JavaScript" type="text/javascript">
var objAjaxCustoFinancFornecConsultaPreco;
var blnConfirmaDifRAeValores=false;
var objSenhaDesconto;
var NUMERO_LOJA_ECOMMERCE_AR_CLUBE = "<%=NUMERO_LOJA_ECOMMERCE_AR_CLUBE%>";
<% if blnMagentoPedidoComIndicador then %>
var FLAG_MAGENTO_PEDIDO_COM_INDICADOR = true;
<% else %>
var FLAG_MAGENTO_PEDIDO_COM_INDICADOR = false;
<% end if %>

function processaFormaPagtoDefault() {
var f, i;
	f=fPED;
	for (i=0; i<fPED.rb_forma_pagto.length; i++) {
		if (fPED.rb_forma_pagto[i].checked) {
			fPED.rb_forma_pagto[i].click();
			break;
			}
		}

	f.c_custoFinancFornecParcelamentoDescricao.value=descricaoCustoFinancFornecTipoParcelamento(f.c_custoFinancFornecTipoParcelamento.value);
	if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (1+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}
	else if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (0+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}
}

function trataRespostaAjaxCustoFinancFornecSincronizaPrecos() {
var f, strResp, i, j, xmlDoc, oNodes;
var strFabricante,strProduto, strStatus, strPrecoLista, strMsgErro, strCodigoErro;
var percDesc,vlLista,vlVenda,strMsgErroAlert;
	f=fPED;
	strMsgErroAlert="";
	if (objAjaxCustoFinancFornecConsultaPreco.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxCustoFinancFornecConsultaPreco.responseText;
		if (strResp=="") {
			alert("Falha ao consultar o preço!");
			window.status="Concluído";
			$("#divAjaxRunning").hide();
			return;
			}

		if (strResp!="") {
			try
				{
				xmlDoc=objAjaxCustoFinancFornecConsultaPreco.responseXML.documentElement;
				for (i=0; i < xmlDoc.getElementsByTagName("ItemConsulta").length; i++) {
				//  Fabricante
					oNodes=xmlDoc.getElementsByTagName("fabricante")[i];
					if (oNodes.childNodes.length > 0) strFabricante=oNodes.childNodes[0].nodeValue; else strFabricante="";
					if (strFabricante==null) strFabricante="";
				//  Produto
					oNodes=xmlDoc.getElementsByTagName("produto")[i];
					if (oNodes.childNodes.length > 0) strProduto=oNodes.childNodes[0].nodeValue; else strProduto="";
					if (strProduto==null) strProduto="";
				//  Status
					oNodes=xmlDoc.getElementsByTagName("status")[i];
					if (oNodes.childNodes.length > 0) strStatus=oNodes.childNodes[0].nodeValue; else strStatus="";
					if (strStatus==null) strStatus="";
					if (strStatus=="OK") {
					//  Preço
						oNodes=xmlDoc.getElementsByTagName("precoLista")[i];
						if (oNodes.childNodes.length > 0) strPrecoLista=oNodes.childNodes[0].nodeValue; else strPrecoLista="";
						if (strPrecoLista==null) strPrecoLista="";
					//  Atualiza o preço
						if (strPrecoLista=="") {
							alert("Falha na consulta do preço do produto " + strProduto + "!\n" + strMsgErro);
							}
						else {
							for (j=0; j<f.c_fabricante.length; j++) {
								if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
								    //  (apesar de que isso não será aceito pelas consistências que serão feitas).
								    if (f.c_preco_lista[j].value == f.c_vl_unitario[j].value) f.c_vl_unitario[j].value=strPrecoLista;
								    if (f.c_preco_lista[j].value == f.c_vl_NF[j].value) f.c_vl_NF[j].value=strPrecoLista;
								    f.c_preco_lista[j].value=strPrecoLista;
								    f.c_preco_lista[j].style.color="black"; 
									}
								}
							}
						}
					else {
					//  Código do Erro
						oNodes=xmlDoc.getElementsByTagName("codigo_erro")[i];
						if (oNodes.childNodes.length > 0) strCodigoErro=oNodes.childNodes[0].nodeValue; else strCodigoErro="";
						if (strCodigoErro==null) strCodigoErro="";
					//  Mensagem de Erro
						oNodes=xmlDoc.getElementsByTagName("msg_erro")[i];
						if (oNodes.childNodes.length > 0) strMsgErro=oNodes.childNodes[0].nodeValue; else strMsgErro="";
						if (strMsgErro==null) strMsgErro="";
						for (j=0; j<f.c_fabricante.length; j++) {
						//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
						//  (apesar de que isso não será aceito pelas consistências que serão feitas).
							if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								f.c_preco_lista[j].style.color=COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE;
								}
							}
						if (strMsgErroAlert!="") strMsgErroAlert+="\n\n";
						strMsgErroAlert+="Falha ao consultar o preço do produto " + strProduto + "!\n" + strMsgErro;
						}
					}
				}
			catch (e)
				{
				alert("Falha na consulta do preço!\n"+e.message);
				}
			}
			
		if (strMsgErroAlert!="") alert(strMsgErroAlert);
		
		recalcula_total_todas_linhas(); 
		recalcula_RA();
		recalcula_RA_Liquido();
		recalcula_parcelas();

		window.status="Concluído";
		$("#divAjaxRunning").hide();
		}
}

function recalculaCustoFinanceiroPrecoLista() {
var f, i, strListaProdutos, strUrl, strOpcaoFormaPagto;
	f=fPED;
	objAjaxCustoFinancFornecConsultaPreco=GetXmlHttpObject();
	if (objAjaxCustoFinancFornecConsultaPreco==null) {
		alert("O browser NÃO possui suporte ao AJAX!");
		return;
		}
		
	strListaProdutos="";
	for (i=0; i<f.c_fabricante.length; i++) {
		if ((trim(f.c_fabricante[i].value)!="")&&(trim(f.c_produto[i].value)!="")) {
			if (strListaProdutos!="") strListaProdutos+=";";
			strListaProdutos += f.c_fabricante[i].value + "|" + f.c_produto[i].value;
			}
		}
	if (strListaProdutos=="") return;
	
//  Converte as opções de forma de pagamento do pedido em uma opção que possa tratada pela tabela de custo financeiro
	strOpcaoFormaPagto="";
	for (i=0; i<fPED.rb_forma_pagto.length; i++) {
		if (fPED.rb_forma_pagto[i].checked) {
			strOpcaoFormaPagto=f.rb_forma_pagto[i].value;
			break;
			}
		}
	if (strOpcaoFormaPagto=="") return;
	
	if (strOpcaoFormaPagto==COD_FORMA_PAGTO_A_VISTA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA;
		f.c_custoFinancFornecQtdeParcelas.value='0';
		}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELA_UNICA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value='1';
		}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_CARTAO) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value=f.c_pc_qtde.value;
	}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value=f.c_pc_maquineta_qtde.value;
	}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value=f.c_pce_prestacao_qtde.value;
		}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value=(converte_numero(f.c_pse_demais_prest_qtde.value)+1).toString();
		}
	else {
		f.c_custoFinancFornecTipoParcelamento.value="";
		f.c_custoFinancFornecQtdeParcelas.value="";
		}
		
	if (trim(f.c_custoFinancFornecQtdeParcelas.value)=="") return;

//  Não consulta novamente se for a mesma consulta anterior
	if ((f.c_custoFinancFornecTipoParcelamento.value==f.c_custoFinancFornecTipoParcelamentoUltConsulta.value)&&
		(f.c_custoFinancFornecQtdeParcelas.value==f.c_custoFinancFornecQtdeParcelasUltConsulta.value)) return;
	
	f.c_custoFinancFornecTipoParcelamentoUltConsulta.value=f.c_custoFinancFornecTipoParcelamento.value;
	f.c_custoFinancFornecQtdeParcelasUltConsulta.value=f.c_custoFinancFornecQtdeParcelas.value;

	f.c_custoFinancFornecParcelamentoDescricao.value=descricaoCustoFinancFornecTipoParcelamento(f.c_custoFinancFornecTipoParcelamento.value);
	if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (1+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}
	else if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (0+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}

	window.status="Aguarde, consultando preços ...";
	$("#divAjaxRunning").show();
	
	strUrl = "../Global/AjaxCustoFinancFornecConsultaPrecoBD.asp";
	strUrl+="?tipoParcelamento="+f.c_custoFinancFornecTipoParcelamento.value;
	strUrl+="&qtdeParcelas="+f.c_custoFinancFornecQtdeParcelas.value;
	strUrl+="&loja="+f.c_loja.value;
	strUrl+="&listaProdutos="+strListaProdutos;
//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxCustoFinancFornecConsultaPreco.onreadystatechange=trataRespostaAjaxCustoFinancFornecSincronizaPrecos;
	objAjaxCustoFinancFornecConsultaPreco.open("GET",strUrl,true);
	objAjaxCustoFinancFornecConsultaPreco.send(null);
}

function executa_consulta_senha_desconto(id_cliente, loja) {
	var postData = "id_cliente=" + id_cliente + "&loja=" + loja;
//	Prevents server from using a cached file
	var url = "../Global/JsonConsultaSenhaDescontoBD.asp" + "?anticache=" + Math.random() + Math.random();
	window.status = "Consultando banco de dados...";
	var responseText = synchronous_ajax(url, postData);
	objSenhaDesconto = eval("(" + responseText + ")");
	window.status = "Concluído";
}

function oPd(fabricante, produto, descricao, mensagem) {
	this.fabricante = fabricante;
	this.produto = produto;
	this.descricao = descricao;
	this.mensagem = mensagem;
}

function obtem_perc_comissao_e_desconto_a_utilizar(f, vl_total_pedido, perc_comissao_e_desconto_nivel1, perc_comissao_e_desconto_nivel1_pj, perc_comissao_e_desconto_nivel2, perc_comissao_e_desconto_nivel2_pj) {
var i, idx, s_pg, blnPreferencial;
var vlNivel1 = 0;
var vlNivel2 = 0;

	// ANALISA QUAL É O MEIO DE PAGAMENTO PREDOMINANTE
	idx = -1;
	//	À Vista
	//	=======
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_pg = trim(f.op_av_forma_pagto.value);
		if (s_pg == '') return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
		for (i = 0; i < vMPN2.length; i++) {
		//	O meio de pagamento selecionado é um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);
		}
		//	O meio de pagamento não é preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}

	//	Parcela Única
	//	=============
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_pg = trim(f.op_pu_forma_pagto.value);
		if (s_pg == '') return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado é um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);
		}
		//	O meio de pagamento não é preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}

	//	Parcelado no Cartão (internet)
	//	==============================
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_pg = ID_FORMA_PAGTO_CARTAO;
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado é um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);
		}
		//	O meio de pagamento não é preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}

	//	Parcelado no Cartão (maquineta)
	//	===============================
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_pg = ID_FORMA_PAGTO_CARTAO_MAQUINETA;
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado é um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);
		}
		//	O meio de pagamento não é preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}

	//	Parcelado Com Entrada
	//	=====================
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		//	Identifica e contabiliza o valor da entrada
		blnPreferencial = false;
		s_pg = trim(f.op_pce_entrada_forma_pagto.value);
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado é um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) {
				blnPreferencial = true;
				break;
			}
		}

		if (blnPreferencial) {
			vlNivel2 = converte_numero(trim(f.c_pce_entrada_valor.value));
		}
		else {
			vlNivel1 = converte_numero(trim(f.c_pce_entrada_valor.value));
		}

		//	Identifica e contabiliza o valor das parcelas
		blnPreferencial = false;
		s_pg = trim(f.op_pce_prestacao_forma_pagto.value);
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado é um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) {
				blnPreferencial = true;
				break;
			}
		}

		if (blnPreferencial) {
			vlNivel2 += converte_numero(f.c_pce_prestacao_qtde.value) * converte_numero(f.c_pce_prestacao_valor.value);
		}
		else {
			vlNivel1 += converte_numero(f.c_pce_prestacao_qtde.value) * converte_numero(f.c_pce_prestacao_valor.value);
		}

		//	O montante a pagar por meio de pagamento preferencial é maior que 50% do total?
		if (vlNivel2 > (vl_total_pedido / 2)) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);
		
		//	O meio de pagamento não é preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}
	
	//	Parcelado Sem Entrada
	//	=====================
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		//	Identifica e contabiliza o valor da 1ª parcela
		blnPreferencial = false;
		s_pg = trim(f.op_pse_prim_prest_forma_pagto.value);
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado é um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) {
				blnPreferencial = true;
				break;
			}
		}

		if (blnPreferencial) {
			vlNivel2 = converte_numero(trim(f.c_pse_prim_prest_valor.value));
		}
		else {
			vlNivel1 = converte_numero(trim(f.c_pse_prim_prest_valor.value));
		}

		//	Identifica e contabiliza o valor das parcelas
		blnPreferencial = false;
		s_pg = trim(f.op_pse_demais_prest_forma_pagto.value);
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado é um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) {
				blnPreferencial = true;
				break;
			}
		}

		if (blnPreferencial) {
			vlNivel2 += converte_numero(trim(f.c_pse_demais_prest_qtde.value)) * converte_numero(trim(f.c_pse_demais_prest_valor.value));
		}
		else {
			vlNivel1 += converte_numero(trim(f.c_pse_demais_prest_qtde.value)) * converte_numero(trim(f.c_pse_demais_prest_valor.value));
		}

		//	O montante a pagar por meio de pagamento preferencial é maior que 50% do total?
		if (vlNivel2 > (vl_total_pedido / 2)) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);
		
		//	O meio de pagamento não é preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}
	
	//	O meio de pagamento não é preferencial
	return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
}

function calcula_vl_total_preco_venda(f) {
var mTotVenda;
	mTotVenda = 0;
	for (i = 0; i < f.c_qtde.length; i++) mTotVenda = mTotVenda + converte_numero(f.c_qtde[i].value) * converte_numero(f.c_vl_unitario[i].value);
	return mTotVenda;
}

// RETORNA O VALOR TOTAL DO PEDIDO A SER USADO P/ CALCULAR A FORMA DE PAGAMENTO
function fp_vl_total_pedido( ) {
var f,i,mTotVenda,mTotNF;
	f=fPED;
	mTotVenda=0;
	for (i=0; i<f.c_qtde.length; i++) mTotVenda=mTotVenda+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_unitario[i].value);
	mTotNF=0;
	for (i=0; i<f.c_qtde.length; i++) mTotNF=mTotNF+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_NF[i].value);
//  Retorna total de preço NF (tem valor de NF, ou seja, pedido c/ RA)?
	if (f.rb_RA.value=='S') {
		return mTotNF;
		}
//  Retorna total de preço de venda
	else {
		return mTotVenda;
		}
}

// PARCELA ÚNICA
function pu_atualiza_valor( ){
var f,vt;
	f=fPED;
	vt=fp_vl_total_pedido();
	f.c_pu_valor.value=formata_moeda(vt);
}

// PARCELADO NO CARTÃO (INTERNET)
function pc_calcula_valor_parcela( ){
var f,n,t;
	f=fPED;
	if (trim(f.c_pc_qtde.value)=='') return;
	n=converte_numero(f.c_pc_qtde.value);
	if (n<=0) return;
	t=fp_vl_total_pedido();
	p=t/n;
	f.c_pc_valor.value=formata_moeda(p);
}

// PARCELADO NO CARTÃO (MAQUINETA)
function pc_maquineta_calcula_valor_parcela( ){
	var f,n,t;
	f=fPED;
	if (trim(f.c_pc_maquineta_qtde.value)=='') return;
	n=converte_numero(f.c_pc_maquineta_qtde.value);
	if (n<=0) return;
	t=fp_vl_total_pedido();
	p=t/n;
	f.c_pc_maquineta_valor.value=formata_moeda(p);
}

// PARCELADO COM ENTRADA
function pce_preenche_sugestao_intervalo() {
var f;
	f=fPED;
	if (converte_numero(trim(f.c_pce_prestacao_periodo.value))>0) return;
	f.c_pce_prestacao_periodo.value='30';
}

function pce_calcula_valor_parcela( ){
var f,n,e,t;
	f=fPED;
	t=fp_vl_total_pedido();
	if (trim(f.c_pce_entrada_valor.value)=='') return;
	e=converte_numero(f.c_pce_entrada_valor.value);
	if (e<=0) return;
	if (trim(f.c_pce_prestacao_qtde.value)=='') return;
	n=converte_numero(f.c_pce_prestacao_qtde.value);
	if (n<=0) return;
	p=(t-e)/n;
	f.c_pce_prestacao_valor.value=formata_moeda(p);
}

// PARCELADO SEM ENTRADA
function pse_preenche_sugestao_intervalo() {
var f;
	f=fPED;
	if (converte_numero(trim(f.c_pse_demais_prest_periodo.value))>0) return;
	f.c_pse_demais_prest_periodo.value='30';
}

function pse_calcula_valor_parcela( ){
var f,n,e,t;
	f=fPED;
	t=fp_vl_total_pedido();
	if (trim(f.c_pse_prim_prest_valor.value)=='') return;
	e=converte_numero(f.c_pse_prim_prest_valor.value);
	if (e<=0) return;
	if (trim(f.c_pse_demais_prest_qtde.value)=='') return;
	n=converte_numero(f.c_pse_demais_prest_qtde.value);
	if (n<=0) return;
	p=(t-e)/n;
	f.c_pse_demais_prest_valor.value=formata_moeda(p);
}

function pce_sugestao_forma_pagto( ) {
var f, p, s, i, n;
	f=fPED;
	f.c_forma_pagto.value="";
	p=converte_numero(f.c_pce_prestacao_periodo.value);
	if (p<=0) return;
	n=converte_numero(f.c_pce_prestacao_qtde.value);
	if (n<=0) return;
	s='0';
	for (i=1; i<=n; i++) {
		s=s+'/';
		s=s+formata_inteiro(i*p);
		}
	f.c_forma_pagto.value=s;
}

function pse_sugestao_forma_pagto( ) {
var f, p1, p2, s, i, n;
	f=fPED;
	f.c_forma_pagto.value="";
	p1=converte_numero(f.c_pse_prim_prest_apos.value);
	if (p1<=0) return;
	p2=converte_numero(f.c_pse_demais_prest_periodo.value);
	if (p2<=0) return;
	n=converte_numero(f.c_pse_demais_prest_qtde.value);
	if (n<=0) return;
	s=formata_inteiro(p1);
	for (i=1; i<=n; i++) {
		s=s+'/';
		s=s+formata_inteiro(i*p2);
		}
	f.c_forma_pagto.value=s;
}

function restaura_cor_desconto( ) {
var f,i;
	f=fPED;
	for (i=0; i < f.c_desc.length; i++) {
		if (converte_numero(f.c_desc[i].value)<0) f.c_desc[i].style.color="red"; else f.c_desc[i].style.color="green";
		}
}

function calcula_desconto(idx) {
	var f, s, i, m, d, m_lista, m_unit;
	f = fPED;
	if (f.c_produto[idx].value == "") return;
	d = converte_numero(f.c_desc[idx].value);
	m_lista = converte_numero(f.c_preco_lista[idx].value);
	m_unit = m_lista - (m_lista * d / 100);
	f.c_vl_unitario[idx].value = formata_moeda(m_unit);
	s = formata_moeda(parseInt(f.c_qtde[idx].value) * m_unit);
	if (f.c_vl_total[idx].value != s) f.c_vl_total[idx].value = s;
	m = 0;
	for (i = 0; i < f.c_vl_total.length; i++) m = m + converte_numero(f.c_vl_total[i].value);
	s = formata_moeda(m);
	if (f.c_total_geral.value != s) f.c_total_geral.value = s;
}

function atualiza_itens_com_desc_linear() {
	var f;
	f = fPED;
	if (trim(f.c_desc_linear.value) == "") return;
	f.c_desc_linear.value = formata_perc_desc_linear(f.c_desc_linear.value);
	if (trim(f.c_desc_linear.value) == "") return;

	for (i = 0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value) != "") {
			f.c_desc[i].value = f.c_desc_linear.value;
			calcula_desconto(i);
			if (blnTemRA) {
				f.c_vl_NF[i].value = f.c_vl_unitario[i].value;
			}
		}
	}
	recalcula_total_todas_linhas();
	recalcula_RA();
	recalcula_RA_Liquido();
}

function recalcula_total_linha( id ) {
var idx, m, m_lista, m_unit, d, f, i, s;
	f=fPED;
	idx=parseInt(id)-1;
	if (f.c_produto[idx].value=="") return;
	m_lista=converte_numero(f.c_preco_lista[idx].value);
	m_unit=converte_numero(f.c_vl_unitario[idx].value);
	if (m_lista==0) d=0; else d=100*(m_lista-m_unit)/m_lista;
	if (d<0) f.c_desc[idx].style.color="red"; else f.c_desc[idx].style.color="green";
	if (d == 0) s = ""; else s = formata_perc_desconto(d);
	if (f.c_desc[idx].value!=s) f.c_desc[idx].value=s;
	s=formata_moeda(parseInt(f.c_qtde[idx].value)*m_unit);
	if (f.c_vl_total[idx].value!=s) f.c_vl_total[idx].value=s;
	m=0;
	for (i=0; i<f.c_vl_total.length; i++) m=m+converte_numero(f.c_vl_total[i].value);
	s=formata_moeda(m);
	if (f.c_total_geral.value!=s) f.c_total_geral.value=s;
	f.c_desc_medio_total.value = formata_perc_desc_linear(calcula_desconto_medio());
}

function recalcula_total_todas_linhas() {
var f,i,vt,m_lista,m_unit,d,m,s;
	f = fPED;
	vt=0;
	for (i=0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value)!="") {
			m_lista=converte_numero(f.c_preco_lista[i].value);
			m_unit=converte_numero(f.c_vl_unitario[i].value);
			if (m_lista==0) d=0; else d=100*(m_lista-m_unit)/m_lista;
			if (d<0) f.c_desc[i].style.color="red"; else f.c_desc[i].style.color="green";
			if (d == 0) s = ""; else s = formata_perc_desconto(d);
			if (f.c_desc[i].value!=s) f.c_desc[i].value=s;
			m=parseInt(f.c_qtde[i].value)*m_unit;
			f.c_vl_total[i].value=formata_moeda(m);
			vt=vt+m;
			}
		}
	f.c_total_geral.value=formata_moeda(vt);
	f.c_desc_medio_total.value = formata_perc_desc_linear(calcula_desconto_medio());
}

function recalcula_RA( ) {
var f,i,mTotVenda,mTotNF;
	f=fPED;
	if (f.rb_RA.value!='S') return;
	mTotVenda=0;
	for (i=0; i<f.c_vl_total.length; i++) mTotVenda=mTotVenda+converte_numero(f.c_vl_total[i].value);
	mTotNF=0;
	for (i=0; i<f.c_qtde.length; i++) mTotNF=mTotNF+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_NF[i].value);
	f.c_total_NF.value = formata_moeda(mTotNF);
	f.c_total_RA.value = formata_moeda(mTotNF-mTotVenda);
	if (mTotNF >=mTotVenda) f.c_total_RA.style.color="green"; else f.c_total_RA.style.color="red";
}

function recalcula_RA_Liquido( ) {
var f,i,mTotVenda,mTotNF,vl_RA,vl_RA_liquido;
var r_RA_liquido;
	f=fPED;
	if (f.rb_RA.value!='S') return;

	recalcula_total_todas_linhas();
	
	mTotVenda=0;
	for (i=0; i<f.c_vl_total.length; i++) mTotVenda=mTotVenda+converte_numero(f.c_vl_total[i].value);
	mTotNF=0;
	for (i=0; i<f.c_qtde.length; i++) mTotNF=mTotNF+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_NF[i].value);
	vl_RA=mTotNF-mTotVenda;

	r_RA_liquido = new calcula_total_RA_liquido(PERC_DESAGIO_RA_LIQUIDA_PEDIDO, vl_RA);
	vl_RA_liquido = r_RA_liquido.vl_total_RA_liquido;
	f.c_total_RA_Liquido.value = formata_moeda(vl_RA_liquido);
	if (vl_RA_liquido>=0) f.c_total_RA_Liquido.style.color="green"; else f.c_total_RA_Liquido.style.color="red";
	if (r_RA_liquido.blnAplicouDesagioRA) f.c_aplicou_desagio_RA.value = "S"; else f.c_aplicou_desagio_RA.value = "N";
}

function consiste_forma_pagto( blnComAvisos ) {
var f,idx,vtNF,vtFP,ve,ni,nip,n,vp;
var MAX_ERRO_ARREDONDAMENTO = 0.1;
	f=fPED;
	vtNF=fp_vl_total_pedido();
	vtFP=0;
	idx=-1;
	
//	À Vista
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.op_av_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento!');
				f.op_av_forma_pagto.focus();
				}
			return false;
			}
		return true;
		}

//	Parcela Única
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.op_pu_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento da parcela única!');
				f.op_pu_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pu_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da parcela única!');
				f.c_pu_valor.focus();
				}
			return false;
			}
		ve=converte_numero(f.c_pu_valor.value);
		vtFP=ve;
		if (ve<=0) {
			if (blnComAvisos) {
				alert('Valor da parcela única é inválido!');
				f.c_pu_valor.focus();
				}
			return false;
			}
		if (trim(f.c_pu_vencto_apos.value)=='') {
			if (blnComAvisos) {
				alert('Indique o intervalo de vencimento da parcela única!');
				f.c_pu_vencto_apos.focus();
				}
			return false;
			}
		nip=converte_numero(f.c_pu_vencto_apos.value);
		if (nip<=0) {
			if (blnComAvisos) {
				alert('Intervalo de vencimento da parcela única é inválido!');
				f.c_pu_vencto_apos.focus();
				}
			return false;
			}
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
				alert('Há divergência entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!');
				f.c_pu_valor.focus();
				}
			return false;
			}
		return true;
		}

//	Parcelado no cartão (internet)
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.c_pc_qtde.value)=='') {
			if (blnComAvisos) {
				alert('Indique a quantidade de parcelas!');
				f.c_pc_qtde.focus();
				}
			return false;
			}
		n=converte_numero(f.c_pc_qtde.value);
		if (n < 1) {
			if (blnComAvisos) {
				alert('Quantidade de parcelas inválida!');
				f.c_pc_qtde.focus();
				}
			return false;
			}
		if (trim(f.c_pc_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da parcela!');
				f.c_pc_valor.focus();
				}
			return false;
			}
		vp=converte_numero(f.c_pc_valor.value);
		if (vp<=0) {
			if (blnComAvisos) {
				alert('Valor de parcela inválido!');
				f.c_pc_valor.focus();
				}
			return false;
			}
		vtFP=n*vp;
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
				alert('Há divergência entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!');
				f.c_pc_valor.focus();
				}
			return false;
			}
		return true;
		}

	//	Parcelado no cartão (maquineta)
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.c_pc_maquineta_qtde.value)=='') {
			if (blnComAvisos) {
				alert('Indique a quantidade de parcelas!');
				f.c_pc_maquineta_qtde.focus();
			}
			return false;
		}
		n=converte_numero(f.c_pc_maquineta_qtde.value);
		if (n < 1) {
			if (blnComAvisos) {
				alert('Quantidade de parcelas inválida!');
				f.c_pc_maquineta_qtde.focus();
			}
			return false;
		}
		if (trim(f.c_pc_maquineta_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da parcela!');
				f.c_pc_maquineta_valor.focus();
			}
			return false;
		}
		vp=converte_numero(f.c_pc_maquineta_valor.value);
		if (vp<=0) {
			if (blnComAvisos) {
				alert('Valor de parcela inválido!');
				f.c_pc_maquineta_valor.focus();
			}
			return false;
		}
		vtFP=n*vp;
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
				alert('Há divergência entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!');
				f.c_pc_maquineta_valor.focus();
			}
			return false;
		}
		return true;
	}

//	Parcelado com entrada
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.op_pce_entrada_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento da entrada!');
				f.op_pce_entrada_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pce_entrada_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da entrada!');
				f.c_pce_entrada_valor.focus();
				}
			return false;
			}
		ve=converte_numero(f.c_pce_entrada_valor.value);
		if (ve<=0) {
			if (blnComAvisos) {
				alert('Valor da entrada inválido!');
				f.c_pce_entrada_valor.focus();
				}
			return false;
			}
		if (trim(f.op_pce_prestacao_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento das prestações!');
				f.op_pce_prestacao_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pce_prestacao_qtde.value)=='') {
			if (blnComAvisos) {
				alert('Indique a quantidade de prestações!');
				f.c_pce_prestacao_qtde.focus();
				}
			return false;
			}
		n=converte_numero(f.c_pce_prestacao_qtde.value);
		if (n<=0) {
			if (blnComAvisos) {
				alert('Quantidade de prestações inválida!');
				f.c_pce_prestacao_qtde.focus();
				}
			return false;
			}
		if (trim(f.c_pce_prestacao_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da prestação!');
				f.c_pce_prestacao_valor.focus();
				}
			return false;
			}
		vp=converte_numero(f.c_pce_prestacao_valor.value);
		if (vp<=0) {
			if (blnComAvisos) {
				alert('Valor de prestação inválido!');
				f.c_pce_prestacao_valor.focus();
				}
			return false;
			}
		vtFP=ve+(n*vp);
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
				alert('Há divergência entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!');
				f.c_pce_prestacao_valor.focus();
				}
			return false;
			}
		if (trim(f.c_pce_prestacao_periodo.value)=='') {
			if (blnComAvisos) {
				alert('Indique o intervalo de vencimento entre as parcelas!');
				f.c_pce_prestacao_periodo.focus();
				}
			return false;
			}
		ni=converte_numero(f.c_pce_prestacao_periodo.value);
		if (ni<=0) {
			if (blnComAvisos) {
				alert('Intervalo de vencimento inválido!');
				f.c_pce_prestacao_periodo.focus();
				}
			return false;
			}
		return true;
		}

//	Parcelado sem entrada
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.op_pse_prim_prest_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento da 1ª prestação!');
				f.op_pse_prim_prest_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pse_prim_prest_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da 1ª prestação!');
				f.c_pse_prim_prest_valor.focus();
				}
			return false;
			}
		ve=converte_numero(f.c_pse_prim_prest_valor.value);
		if (ve<=0) {
			if (blnComAvisos) {
				alert('Valor da 1ª prestação inválido!');
				f.c_pse_prim_prest_valor.focus();
				}
			return false;
			}
		if (trim(f.c_pse_prim_prest_apos.value)=='') {
			if (blnComAvisos) {
				alert('Indique o intervalo de vencimento da 1ª parcela!');
				f.c_pse_prim_prest_apos.focus();
				}
			return false;
			}
		nip=converte_numero(f.c_pse_prim_prest_apos.value);
		if (nip<=0) {
			if (blnComAvisos) {
				alert('Intervalo de vencimento da 1ª parcela é inválido!');
				f.c_pse_prim_prest_apos.focus();
				}
			return false;
			}
		if (trim(f.op_pse_demais_prest_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento das demais prestações!');
				f.op_pse_demais_prest_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pse_demais_prest_qtde.value)=='') {
			if (blnComAvisos) {
				alert('Indique a quantidade das demais prestações!');
				f.c_pse_demais_prest_qtde.focus();
				}
			return false;
			}
		n=converte_numero(f.c_pse_demais_prest_qtde.value);
		if (n<=0) {
			if (blnComAvisos) {
				alert('Quantidade de prestações inválida!');
				f.c_pse_demais_prest_qtde.focus();
				}
			return false;
			}
		if (trim(f.c_pse_demais_prest_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor das demais prestações!');
				f.c_pse_demais_prest_valor.focus();
				}
			return false;
			}
		vp=converte_numero(f.c_pse_demais_prest_valor.value);
		if (vp<=0) {
			if (blnComAvisos) {
				alert('Valor de prestação inválido!');
				f.c_pse_demais_prest_valor.focus();
				}
			return false;
			}
		vtFP=ve+(n*vp);
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
				alert('Há divergência entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!');
				f.c_pse_demais_prest_valor.focus();
				}
			return false;
			}
		if (trim(f.c_pse_demais_prest_periodo.value)=='') {
			if (blnComAvisos) {
				alert('Indique o intervalo de vencimento entre as parcelas!');
				f.c_pse_demais_prest_periodo.focus();
				}
			return false;
			}
		ni=converte_numero(f.c_pse_demais_prest_periodo.value);
		if (ni<=0) {
			if (blnComAvisos) {
				alert('Intervalo de vencimento inválido!');
				f.c_pse_demais_prest_periodo.focus();
				}
			return false;
			}
		return true;
		}
		
	if (blnComAvisos) {
		// Nenhuma forma de pagamento foi escolhida
		alert('Indique a forma de pagamento!');
		}
		
	return false;
}

function recalcula_parcelas() {
    var f, idx;
    f = fPED;
    idx=-1;

    idx++;
    idx++;
    if (f.rb_forma_pagto[idx].checked) {
        pu_atualiza_valor();
        return;
    }

    idx++;
    if (f.rb_forma_pagto[idx].checked) {
        pc_calcula_valor_parcela();
        return;
    }

    idx++;
    if (f.rb_forma_pagto[idx].checked) {
        pce_calcula_valor_parcela();
        return;
    }

    idx++;
    if (f.rb_forma_pagto[idx].checked) {
        pse_calcula_valor_parcela();
        return;
    }
  
}

function calcula_desconto_medio() {
	var f, i, vl_total_preco_lista, vl_total_preco_venda, perc_desc_medio;
	
	f = fPED;
	vl_total_preco_lista = 0;
	vl_total_preco_venda = 0;
	
	// Laço p/ produtos
	for (i = 0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value) != "") {
			vl_total_preco_lista += converte_numero(f.c_qtde[i].value) * converte_numero(f.c_preco_lista[i].value);
			vl_total_preco_venda += converte_numero(f.c_qtde[i].value) * converte_numero(f.c_vl_unitario[i].value);
		}
	}

	if (vl_total_preco_lista == 0) {
		perc_desc_medio = 0;
	}
	else {
		perc_desc_medio = 100 * (vl_total_preco_lista - vl_total_preco_venda) / vl_total_preco_lista;
	}
	return perc_desc_medio;
}

function trata_edicao_RA(index) {
var f;
	f = fPED;
	if ((f.c_permite_RA_status.value != '1') || (f.rb_RA.value != 'S')) f.c_vl_NF[index].value = f.c_vl_unitario[index].value;
}

function fOpCancela( f )
{
	f.submit();
}

function fPEDConfirma( f ) {
var s,i,j,vl_preco_lista,vl_preco_venda,vl_NF,perc_desc,blnFlag,strProduto,strLinha,strMsgAlerta,vlAux,strMsgErro;
var perc_RT, perc_RT_novo, perc_max_RT_padrao, perc_max_comissao_e_desconto, perc_max_comissao_e_desconto_pj, perc_max_comissao_e_desconto_nivel2, perc_max_comissao_e_desconto_nivel2_pj, perc_senha_desconto, perc_desc_medio;
var perc_max_RT_a_utilizar, perc_max_comissao_e_desconto_a_utilizar;
var perc_max_desc_alcada_1_pf, perc_max_desc_alcada_1_pj, perc_max_desc_alcada_2_pf, perc_max_desc_alcada_2_pj, perc_max_desc_alcada_3_pf, perc_max_desc_alcada_3_pj;
var perc_max_comissao_alcada1, perc_max_comissao_alcada2, perc_max_comissao_alcada3;

	recalcula_total_todas_linhas();

	recalcula_RA();
	
	recalcula_RA_Liquido();
	
	s = "" + f.c_obs1.value;
	if (s.length > MAX_TAM_OBS1) {
		alert('Conteúdo de "Observações " excede em ' + (s.length-MAX_TAM_OBS1) + ' caracteres o tamanho máximo de ' + MAX_TAM_OBS1 + '!');
		f.c_obs1.focus();
		return;
	}

	s = "" + f.c_nf_texto.value;
	if (s.length > MAX_TAM_NF_TEXTO) {
	    alert('Conteúdo de "Constar na NF" excede em ' + (s.length-MAX_TAM_NF_TEXTO) + ' caracteres o tamanho máximo de ' + MAX_TAM_NF_TEXTO + '!');
	    f.c_nf_texto.focus();
	    return;
	}
	
	s = "" + f.c_forma_pagto.value;
	if (s.length > MAX_TAM_FORMA_PAGTO) {
		alert('Conteúdo de "Forma de Pagamento" excede em ' + (s.length-MAX_TAM_FORMA_PAGTO) + ' caracteres o tamanho máximo de ' + MAX_TAM_FORMA_PAGTO + '!');
		f.c_forma_pagto.focus();
		return;
		}
	
//  Consiste a nova versão da forma de pagamento
	if (!consiste_forma_pagto(true)) return;
	
	if (f.rb_RA.value=='S') {
	//	Limita o RA a um percentual do valor do pedido
		if (converte_numero(f.c_PercVlPedidoLimiteRA.value)!=0) {
			vlAux = (converte_numero(f.c_PercVlPedidoLimiteRA.value)/100) * converte_numero(f.c_total_geral.value);
			if (($("#c_loja").val()!=NUMERO_LOJA_ECOMMERCE_AR_CLUBE)&&(!FLAG_MAGENTO_PEDIDO_COM_INDICADOR)){
				if (converte_numero(f.c_total_RA.value) > vlAux) {
					alert('O valor total de RA excede o limite permitido para este pedido!');
					return;
				}
			}
		}

		if (blnConfirmaDifRAeValores) {
			if (converte_numero(f.c_total_RA.value) != 0) {
				if (!confirm("O valor do RA é de " + SIMBOLO_MONETARIO + " " + formata_moeda(converte_numero(f.c_total_RA.value))+"\nContinua?")) return;
				}
			}
		}

	// Consiste percentual máximo de comissão e desconto
	objSenhaDesconto = null;
	perc_RT = converte_numero(f.c_perc_RT.value);
	perc_max_RT_padrao = converte_numero(f.c_PercMaxRT.value);
	perc_max_RT_a_utilizar = perc_max_RT_padrao;

	perc_max_comissao_e_desconto = converte_numero(f.c_PercMaxComissaoEDesconto.value);
	perc_max_comissao_e_desconto_pj = converte_numero(f.c_PercMaxComissaoEDescontoPj.value);
	perc_max_comissao_e_desconto_nivel2 = converte_numero(f.c_PercMaxComissaoEDescontoNivel2.value);
	perc_max_comissao_e_desconto_nivel2_pj = converte_numero(f.c_PercMaxComissaoEDescontoNivel2Pj.value);
	perc_max_comissao_e_desconto_a_utilizar = obtem_perc_comissao_e_desconto_a_utilizar(f, calcula_vl_total_preco_venda(f), perc_max_comissao_e_desconto, perc_max_comissao_e_desconto_pj, perc_max_comissao_e_desconto_nivel2, perc_max_comissao_e_desconto_nivel2_pj);

	perc_desc_medio = calcula_desconto_medio();

	// Verifica se o usuário tem permissão de desconto por alçada
	perc_max_comissao_alcada1 = converte_numero(f.c_PercMaxRTAlcada1.value);
	perc_max_desc_alcada_1_pf = converte_numero(f.c_PercMaxDescAlcada1Pf.value);
	perc_max_desc_alcada_1_pj = converte_numero(f.c_PercMaxDescAlcada1Pj.value);
	perc_max_comissao_alcada2 = converte_numero(f.c_PercMaxRTAlcada2.value);
	perc_max_desc_alcada_2_pf = converte_numero(f.c_PercMaxDescAlcada2Pf.value);
	perc_max_desc_alcada_2_pj = converte_numero(f.c_PercMaxDescAlcada2Pj.value);
	perc_max_comissao_alcada3 = converte_numero(f.c_PercMaxRTAlcada3.value);
	perc_max_desc_alcada_3_pf = converte_numero(f.c_PercMaxDescAlcada3Pf.value);
	perc_max_desc_alcada_3_pj = converte_numero(f.c_PercMaxDescAlcada3Pj.value);

	if (perc_max_comissao_alcada1 > perc_max_RT_a_utilizar) perc_max_RT_a_utilizar = perc_max_comissao_alcada1;
	if (perc_max_comissao_alcada2 > perc_max_RT_a_utilizar) perc_max_RT_a_utilizar = perc_max_comissao_alcada2;
	if (perc_max_comissao_alcada3 > perc_max_RT_a_utilizar) perc_max_RT_a_utilizar = perc_max_comissao_alcada3;

	if (f.c_tipo_cliente.value == ID_PF) {
		if (perc_max_desc_alcada_1_pf > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_1_pf;
		if (perc_max_desc_alcada_2_pf > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_2_pf;
		if (perc_max_desc_alcada_3_pf > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_3_pf;
	}
	else {
		if (perc_max_desc_alcada_1_pj > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_1_pj;
		if (perc_max_desc_alcada_2_pj > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_2_pj;
		if (perc_max_desc_alcada_3_pj > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_3_pj;
	}

	// Verifica se todos os produtos cujo desconto excedem o máximo permitido possuem senha de desconto disponível
	// Laço p/ produtos
	strMsgErro = "";
	for (i = 0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value) != "") {
			perc_senha_desconto = 0;
			vl_preco_lista = converte_numero(f.c_preco_lista[i].value);
			vl_preco_venda = converte_numero(f.c_vl_unitario[i].value);
			if (vl_preco_lista == 0) {
				perc_desc = 0;
			}
			else {
				perc_desc = 100 * (vl_preco_lista - vl_preco_venda) / vl_preco_lista;
			}

			// Tem desconto: sim
			if (perc_desc != 0) {
				// Desconto excede limite máximo: sim
				if (perc_desc > perc_max_comissao_e_desconto_a_utilizar) {
					// Tem senha de desconto?
					if (objSenhaDesconto == null) {
						executa_consulta_senha_desconto(f.cliente_selecionado.value, f.c_loja.value);
					}
					for (j = 0; j < objSenhaDesconto.item.length; j++) {
						if ((objSenhaDesconto.item[j].fabricante == f.c_fabricante[i].value) && (objSenhaDesconto.item[j].produto == f.c_produto[i].value)) {
							perc_senha_desconto = converte_numero(objSenhaDesconto.item[j].desc_max);
							break;
						}
					}
					// Tem senha de desconto: sim
					if (perc_senha_desconto != 0) {
						// Senha de desconto NÃO cobre desconto
						if (perc_senha_desconto < perc_desc) {
							if (strMsgErro != "") strMsgErro += "\n";
							strMsgErro += "O desconto do produto '" + f.c_descricao[i].value + "' (" + formata_numero(perc_desc, 2) + "%) excede o máximo autorizado!";
						}
					}
					// Não tem senha de desconto
					else {
						if (strMsgErro != "") strMsgErro += "\n";
						strMsgErro += "O desconto do produto '" + f.c_descricao[i].value + "' (" + formata_numero(perc_desc, 2) + "%) excede o máximo permitido!";
					}
				} // if (perc_desc > perc_max_comissao_e_desconto_a_utilizar)
			} // if (perc_desc != 0)
		} // if (trim(f.c_produto[i].value) != "")
	} // for (laço produtos)

	if (strMsgErro != "") {
		strMsgErro += "\n\nNão é possível continuar!";
		alert(strMsgErro);
		return;
	}
	
	// Tem RT: sim
	if (f.operacao_origem.value != "<%=OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO%>") {
		if (perc_RT != 0) {
			// RT excede limite máximo?
			if (perc_RT > perc_max_RT_a_utilizar) {
				alert("Percentual de comissão excede o máximo permitido!");
				return;
			}

			// Neste ponto, é certo que todos os produtos que possuem desconto estão dentro do máximo permitido
			// ou possuem senha de desconto autorizando.
			// Verifica-se agora se é necessário reduzir automaticamente o percentual da RT usando p/ o cálculo
			// o percentual de desconto médio.
			perc_RT_novo = Math.min(perc_RT, (perc_max_comissao_e_desconto_a_utilizar - perc_desc_medio));
			if (perc_RT_novo < 0) perc_RT_novo = 0;

			// O percentual de RT será alterado automaticamente, solicita confirmação
			if (perc_RT_novo != perc_RT) {
				s = "A soma dos percentuais de comissão (" + formata_numero(perc_RT, 2) + "%) e de desconto médio do(s) produto(s) (" + formata_numero(perc_desc_medio, 2) + "%) totaliza " + formata_numero(perc_desc_medio + perc_RT, 2) + "% e excede o máximo permitido!" +
					"\nA comissão será reduzida automaticamente para " + formata_numero(perc_RT_novo, 2) + "%!" +
					"\nContinua?";
				if (!confirm(s)) {
					s = "Operação cancelada!";
					alert(s);
					return;
				}
				else {
					// Verifica se o novo percentual de RT está dentro do limite definido p/ o perfil do usuário que está editando o pedido
					if (perc_RT_novo > perc_max_RT_a_utilizar) {
						s = "O percentual de comissão (" + formata_numero(perc_RT_novo, 2) + "%) excede o máximo permitido!!" +
							"\nA comissão será reduzida automaticamente para " + formata_numero(perc_max_RT_a_utilizar, 2) + "%!!" +
							"\nContinua?";
						if (!confirm(s)) {
							s = "Operação cancelada!!";
							alert(s);
							return;
						}
						else {
							// Novo percentual de RT
							perc_RT_novo = perc_max_RT_a_utilizar;
						}
					}

					// Novo percentual de RT
					f.c_perc_RT.value = formata_perc_RT(perc_RT_novo);
					perc_RT = perc_RT_novo;
				}
			}
		} // if (perc_RT != 0)
	} // if (f.operacao_origem.value != "PED_NOVO_EC_SEMI_AUTO")
	
	blnFlag=false;
	for (i=0; i < f.rb_etg_imediata.length; i++) {
		if (f.rb_etg_imediata[i].checked) blnFlag=true;
		}
	if (!blnFlag) {
		alert('Selecione uma opção para o campo "Entrega Imediata"');
		return;
		}

	if (f.rb_etg_imediata[0].checked)
	{
		if (trim(f.c_data_previsao_entrega.value) == "") {
			alert("Informe a data de previsão de entrega!");
			f.c_data_previsao_entrega.focus();
			return;
		}

		if (!isDate(f.c_data_previsao_entrega)) {
            alert("Data de previsão de entrega é inválida!");
            f.c_data_previsao_entrega.focus();
			return;
		}

		if (retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(f.c_data_previsao_entrega.value)) <= retorna_so_digitos(formata_ddmmyyyy_yyyymmdd('<%=formata_data(Date)%>'))) {
			alert("Data de previsão de entrega deve ser uma data futura!");
            f.c_data_previsao_entrega.focus();
			return;
        }
	}

	blnFlag=false;
	for (i=0; i < f.rb_bem_uso_consumo.length; i++) {
		if (f.rb_bem_uso_consumo[i].checked) blnFlag=true;
		}
	if (!blnFlag) {
		alert('Informe se é "Bem de Uso/Consumo"');
		return;
		}
		
	if (f.c_exibir_campo_instalador_instala.value=="S") {
		blnFlag=false;
		for (i=0; i < f.rb_instalador_instala.length; i++) {
			if (f.rb_instalador_instala[i].checked) blnFlag=true;
			}
		if (!blnFlag) {
			alert('Preencha o campo "Instalador Instala"');
			return;
			}
		}
		
	if (f.rb_indicacao.value=="S") {
		blnFlag=false;
		for (i=0; i < f.rb_garantia_indicador.length; i++) {
			if (f.rb_garantia_indicador[i].checked) blnFlag=true;
			}
		if (!blnFlag) {
			alert('Preencha o campo "Garantia Indicador"');
			return;
			}
		}
	
//  Há mensagens de alerta para os produtos?
//  A primeira posição do vetor é vazia, apenas p/ garantir que o vetor existe mesmo quando não há mensagens
	for (i=1; i < Pd.length; i++) {
		if (trim(Pd[i].mensagem)!="") {
			strProduto="Produto: " + trim(Pd[i].produto) + " - " + trim(Pd[i].descricao);
			strLinha=new Array(strProduto.length).join("=");
			strMsgAlerta=strLinha + "\n" + strProduto + "\n" + strLinha + "\n\n" + trim(Pd[i].mensagem) + "\n";
			if (!confirm(strMsgAlerta)) return;
			}
		}
	
	strMsgErro="";
	for (i=0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value)!="") {
			if (f.c_preco_lista[i].style.color.toLowerCase()==COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE.toLowerCase()) {
				strMsgErro+="\n" + f.c_produto[i].value + " - " + f.c_descricao[i].value;
				}
			}
		}
	if (strMsgErro!="") {
		strMsgErro="A forma de pagamento " + KEY_ASPAS + f.c_custoFinancFornecParcelamentoDescricao.value.toLowerCase() + KEY_ASPAS + " não está disponível para o(s) produto(s):"+strMsgErro;
		alert(strMsgErro);
		return;
		}
	
	if ($("#c_loja").val()==NUMERO_LOJA_ECOMMERCE_AR_CLUBE){
		if ($("#c_origem_pedido").val() == ""){
			alert("Selecione a origem do pedido (marketplace)!");
			$("#c_origem_pedido").focus();
			return;
		}

		if ($("#c_pedido_ac").val() != "") {
		    if(retorna_so_digitos($("#c_pedido_ac").val()) != $("#c_pedido_ac").val()) {
		        alert("O número Magento deve conter apenas dígitos!");
		        $("#c_pedido_ac").focus();
		        return;
		    }
		}
	}

	if (FLAG_MAGENTO_PEDIDO_COM_INDICADOR)
	{
		if ($("#c_pedido_ac").val() != "") {
			if(retorna_so_digitos($("#c_pedido_ac").val()) != $("#c_pedido_ac").val()) {
				alert("O número Magento deve conter apenas dígitos!");
				$("#c_pedido_ac").focus();
				return;
			}
		}
	}

	// CONSISTÊNCIA PARA VALOR ZERADO
    strMsgErro = "";
    for (i = 0; i < f.c_produto.length; i++) {
        if (trim(f.c_produto[i].value) != "") {
            vl_preco_venda = converte_numero(f.c_vl_unitario[i].value);
            if (vl_preco_venda <= 0) {
                if (strMsgErro != "") strMsgErro += "\n";
                strMsgErro += "O produto '" + f.c_descricao[i].value + "' está com valor de venda zerado!";
            }
            else if ((f.c_permite_RA_status.value == '1') && (f.rb_RA.value == 'S')) {
                vl_NF = converte_numero(f.c_vl_NF[i].value);
                if (vl_NF <= 0) {
                    if (strMsgErro != "") strMsgErro += "\n";
                    strMsgErro += "O produto '" + f.c_descricao[i].value + "' está com o preço zerado!";
                }
            }
        }
    }

    if (strMsgErro != "") {
        strMsgErro += "\n\nNão é possível continuar!";
        alert(strMsgErro);
        return;
    }

	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}

</script>
<script language='JavaScript'>
    function SomenteNumero(e){
        var tecla=(window.event)?event.keyCode:e.which;   
        if((tecla>47 && tecla<58)) return true;
        else{
            if (tecla==8 || tecla==0) return true;
            else  return false;
        }
    }
    function SomenteNumeroHifen(e){
        var tecla=(window.event)?event.keyCode:e.which;   
        if((tecla>47 && tecla<58)) return true;
        else{
            if (tecla==8 || tecla==0 || tecla==45) return true;
            else  return false;
        }
    }
</script>

<%
	Response.Write strScriptMsgAlerta
%>




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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style type="text/css">
#rb_etg_imediata, #rb_bem_uso_consumo, #rb_instalador_instala {
	margin: 0pt 2pt 1pt 3pt;
	vertical-align: top;
	}
#rb_forma_pagto {
	margin: 0pt 2pt 1pt 10pt;
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
.TdCliLbl
{
	width:200px;
	text-align:right;
}
.TdCliCel
{
	width:450px;
	text-align:left;
}
.TdCliBtn
{
	width:30px;
	text-align:center;
}
</style>

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
<table cellspacing="0" width="649">
<% if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then %>
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCANCELA" id="dCANCELA">
		<a name="bCANCELA" id="bCANCELA" href="javascript:fOpCancela(fCANCEL)" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% else %>
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
<% end if %>
</table>
</center>
</body>




<% else %>
<!-- ********************************************************** -->
<!-- **********  PÁGINA PARA EDITAR ITENS DO PEDIDO  ********** -->
<!-- ********************************************************** -->
<body onload="if (!(erro_produto_indisponivel&&bloquear_cadastramento_quando_produto_indiponivel)) {processaFormaPagtoDefault();restaura_cor_desconto();fPED.c_obs1.focus();}">
<center>

<form id="fPED" name="fPED" method="post" action="PedidoNovoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=cliente_selecionado%>'>
<input type="hidden" name="c_cnpj_cpf" id="c_cnpj_cpf" value='<%=EndCob_cnpj_cpf%>'>
<input type="hidden" name="c_tipo_cliente" id="c_tipo_cliente" value='<%=EndCob_tipo_pessoa%>'>
<% if erro_produto_indisponivel then s="S" else s="" %>
<input type="hidden" name="opcao_venda_sem_estoque" id="opcao_venda_sem_estoque" value='<%=s%>'>
<input type="hidden" name="insert_request_guid" id="insert_request_guid" value="<%=insert_request_guid%>" />

<input type="hidden" name="rb_indicacao" id="rb_indicacao" value='<%=rb_indicacao%>'>
<input type="hidden" name="rb_RA" id="rb_RA" value='<%=rb_RA%>'>
<input type="hidden" name="c_indicador" id="c_indicador" value='<%=c_indicador%>'>
<input type="hidden" name="c_ped_bonshop" id="c_ped_bonshop" value='<%=c_ped_bonshop %>' />

<input type="hidden" name="c_PercLimiteRASemDesagio" id="c_PercLimiteRASemDesagio" value='<%=strPercLimiteRASemDesagio%>'>
<input type="hidden" name="c_PercDesagio" id="c_PercDesagio" value='<%=strPercDesagio%>'>
<input type="hidden" name="c_aplicou_desagio_RA" id="c_aplicou_desagio_RA" value=''>

<input type="hidden" name="c_PercMaxRT" id="c_PercMaxRT" value='<%=strPercMaxRT%>'>
<input type="hidden" name="c_PercMaxComissaoEDesconto" id="c_PercMaxComissaoEDesconto" value='<%=strPercMaxComissaoEDesconto%>'>
<input type="hidden" name="c_PercMaxComissaoEDescontoPj" id="c_PercMaxComissaoEDescontoPj" value='<%=strPercMaxComissaoEDescontoPj%>'>
<input type="hidden" name="c_PercMaxComissaoEDescontoNivel2" id="c_PercMaxComissaoEDescontoNivel2" value='<%=strPercMaxComissaoEDescontoNivel2%>'>
<input type="hidden" name="c_PercMaxComissaoEDescontoNivel2Pj" id="c_PercMaxComissaoEDescontoNivel2Pj" value='<%=strPercMaxComissaoEDescontoNivel2Pj%>'>
<input type="hidden" name="c_PercMaxRTAlcada1" id="c_PercMaxRTAlcada1" value="<%=strPercMaxRTAlcada1%>" />
<input type="hidden" name="c_PercMaxDescAlcada1Pf" id="c_PercMaxDescAlcada1Pf" value="<%=strPercMaxDescAlcada1Pf%>" />
<input type="hidden" name="c_PercMaxDescAlcada1Pj" id="c_PercMaxDescAlcada1Pj" value="<%=strPercMaxDescAlcada1Pj%>" />
<input type="hidden" name="c_PercMaxRTAlcada2" id="c_PercMaxRTAlcada2" value="<%=strPercMaxRTAlcada2%>" />
<input type="hidden" name="c_PercMaxDescAlcada2Pf" id="c_PercMaxDescAlcada2Pf" value="<%=strPercMaxDescAlcada2Pf%>" />
<input type="hidden" name="c_PercMaxDescAlcada2Pj" id="c_PercMaxDescAlcada2Pj" value="<%=strPercMaxDescAlcada2Pj%>" />
<input type="hidden" name="c_PercMaxRTAlcada3" id="c_PercMaxRTAlcada3" value="<%=strPercMaxRTAlcada3%>" />
<input type="hidden" name="c_PercMaxDescAlcada3Pf" id="c_PercMaxDescAlcada3Pf" value="<%=strPercMaxDescAlcada3Pf%>" />
<input type="hidden" name="c_PercMaxDescAlcada3Pj" id="c_PercMaxDescAlcada3Pj" value="<%=strPercMaxDescAlcada3Pj%>" />
<input type="hidden" name="c_PercVlPedidoLimiteRA" id="c_PercVlPedidoLimiteRA" value='<%=strPercVlPedidoLimiteRA%>'>
<input type="hidden" name="c_permite_RA_status" id="c_permite_RA_status" value='<%=permite_RA_status%>' />

<input type="hidden" name="rb_end_entrega" id="rb_end_entrega" value='<%=rb_end_entrega%>'>
<input type="hidden" name="EndEtg_endereco" id="EndEtg_endereco" value="<%=EndEtg_endereco%>">
<input type="hidden" name="EndEtg_endereco_numero" id="EndEtg_endereco_numero" value="<%=EndEtg_endereco_numero%>">
<input type="hidden" name="EndEtg_endereco_complemento" id="EndEtg_endereco_complemento" value="<%=EndEtg_endereco_complemento%>">
<input type="hidden" name="EndEtg_bairro" id="EndEtg_bairro" value="<%=EndEtg_bairro%>">
<input type="hidden" name="EndEtg_cidade" id="EndEtg_cidade" value="<%=EndEtg_cidade%>">
<input type="hidden" name="EndEtg_uf" id="EndEtg_uf" value="<%=EndEtg_uf%>">
<input type="hidden" name="EndEtg_cep" id="EndEtg_cep" value="<%=EndEtg_cep%>">
<input type="hidden" name="EndEtg_obs" id="EndEtg_obs" value='<%=EndEtg_obs%>'>
<% if operacao_permitida(OP_LJA_EXIBIR_CAMPO_INSTALADOR_INSTALA_AO_CADASTRAR_NOVO_PEDIDO, s_lista_operacoes_permitidas) then s="S" else s="" %>
<input type="hidden" name="c_exibir_campo_instalador_instala" id="c_exibir_campo_instalador_instala" value='<%=s%>'>
<input type="hidden" name="c_loja" id="c_loja" value='<%=loja%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamento" id="c_custoFinancFornecTipoParcelamento" value='<%=c_custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelas" id="c_custoFinancFornecQtdeParcelas" value='<%=c_custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamentoUltConsulta" id="c_custoFinancFornecTipoParcelamentoUltConsulta" value='<%=c_custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelasUltConsulta" id="c_custoFinancFornecQtdeParcelasUltConsulta" value='<%=c_custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecParcelamentoDescricao" id="c_custoFinancFornecParcelamentoDescricao" value=''>
<input type="hidden" name="rb_selecao_cd" id="rb_selecao_cd" value="<%=rb_selecao_cd%>" />
<input type="hidden" name="c_id_nfe_emitente_selecao_manual" id="c_id_nfe_emitente_selecao_manual" value="<%=c_id_nfe_emitente_selecao_manual%>" />
<input type="hidden" name="operacao_origem" id="operacao_origem" value="<%=operacao_origem%>" />
<input type="hidden" name="id_magento_api_pedido_xml" id="id_magento_api_pedido_xml" value="<%=id_magento_api_pedido_xml%>" />
<input type="hidden" name="c_numero_magento" id="c_numero_magento" value="<%=c_numero_magento%>" />
<input type="hidden" name="operationControlTicket" id="operationControlTicket" value="<%=operationControlTicket%>" />
<input type="hidden" name="sessionToken" id="sessionToken" value="<%=sessionToken%>" />

<!--  CAMPOS ADICIONAIS DO ENDERECO DE ENTREGA  -->
<input type="hidden" name="EndEtg_endereco_ponto_referencia" id="EndEtg_endereco_ponto_referencia" value="<%=EndEtg_endereco_ponto_referencia%>" />
<input type="hidden" name="EndEtg_email" id="EndEtg_email" value="<%=EndEtg_email%>" />
<input type="hidden" name="EndEtg_email_xml" id="EndEtg_email_xml" value="<%=EndEtg_email_xml%>" />
<input type="hidden" name="EndEtg_nome" id="EndEtg_nome" value="<%=EndEtg_nome%>" />
<input type="hidden" name="EndEtg_ddd_res" id="EndEtg_ddd_res" value="<%=EndEtg_ddd_res%>" />
<input type="hidden" name="EndEtg_tel_res" id="EndEtg_tel_res" value="<%=EndEtg_tel_res%>" />
<input type="hidden" name="EndEtg_ddd_com" id="EndEtg_ddd_com" value="<%=EndEtg_ddd_com%>" />
<input type="hidden" name="EndEtg_tel_com" id="EndEtg_tel_com" value="<%=EndEtg_tel_com%>" />
<input type="hidden" name="EndEtg_ramal_com" id="EndEtg_ramal_com" value="<%=EndEtg_ramal_com%>" />
<input type="hidden" name="EndEtg_ddd_cel" id="EndEtg_ddd_cel" value="<%=EndEtg_ddd_cel%>" />
<input type="hidden" name="EndEtg_tel_cel" id="EndEtg_tel_cel" value="<%=EndEtg_tel_cel%>" />
<input type="hidden" name="EndEtg_ddd_com_2" id="EndEtg_ddd_com_2" value="<%=EndEtg_ddd_com_2%>" />
<input type="hidden" name="EndEtg_tel_com_2" id="EndEtg_tel_com_2" value="<%=EndEtg_tel_com_2%>" />
<input type="hidden" name="EndEtg_ramal_com_2" id="EndEtg_ramal_com_2" value="<%=EndEtg_ramal_com_2%>" />
<input type="hidden" name="EndEtg_tipo_pessoa" id="EndEtg_tipo_pessoa" value="<%=EndEtg_tipo_pessoa%>" />
<input type="hidden" name="EndEtg_cnpj_cpf" id="EndEtg_cnpj_cpf" value="<%=EndEtg_cnpj_cpf%>" />
<input type="hidden" name="EndEtg_contribuinte_icms_status" id="EndEtg_contribuinte_icms_status" value="<%=EndEtg_contribuinte_icms_status%>" />
<input type="hidden" name="EndEtg_produtor_rural_status" id="EndEtg_produtor_rural_status" value="<%=EndEtg_produtor_rural_status%>" />
<input type="hidden" name="EndEtg_ie" id="EndEtg_ie" value="<%=EndEtg_ie%>" />
<input type="hidden" name="EndEtg_rg" id="EndEtg_rg" value="<%=EndEtg_rg%>" />
<input type="hidden" name="c_FlagCadSemiAutoPedMagento_FluxoOtimizado" id="c_FlagCadSemiAutoPedMagento_FluxoOtimizado" value="<%=c_FlagCadSemiAutoPedMagento_FluxoOtimizado%>" />

<% if c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "1" then %>
<input type="hidden" name="EndCob_endereco" id="EndCob_endereco" value="<%=EndCob_endereco%>" />
<input type="hidden" name="EndCob_endereco_numero" id="EndCob_endereco_numero" value="<%=EndCob_endereco_numero%>" />
<input type="hidden" name="EndCob_endereco_complemento" id="EndCob_endereco_complemento" value="<%=EndCob_endereco_complemento%>" />
<input type="hidden" name="EndCob_endereco_ponto_referencia" id="EndCob_endereco_ponto_referencia" value="<%=EndCob_endereco_ponto_referencia%>" />
<input type="hidden" name="EndCob_bairro" id="EndCob_bairro" value="<%=EndCob_bairro%>" />
<input type="hidden" name="EndCob_cidade" id="EndCob_cidade" value="<%=EndCob_cidade%>" />
<input type="hidden" name="EndCob_uf" id="EndCob_uf" value="<%=EndCob_uf %>" />
<input type="hidden" name="EndCob_cep" id="EndCob_cep" value="<%=EndCob_cep%>" />
<input type="hidden" name="EndCob_email" id="EndCob_email" value="<%=EndCob_email%>" />
<input type="hidden" name="EndCob_email_xml" id="EndCob_email_xml" value="<%=EndCob_email_xml%>" />
<input type="hidden" name="EndCob_email_boleto" id="EndCob_email_boleto" value="<%=EndCob_email_boleto%>" />
<input type="hidden" name="EndCob_nome" id="EndCob_nome" value="<%=EndCob_nome%>" />
<input type="hidden" name="EndCob_tipo_pessoa" id="EndCob_tipo_pessoa" value="<%=EndCob_tipo_pessoa%>" />
<input type="hidden" name="EndCob_ddd_res" id="EndCob_ddd_res" value="<%=EndCob_ddd_res%>" />
<input type="hidden" name="EndCob_tel_res" id="EndCob_tel_res" value="<%=EndCob_tel_res%>" />
<input type="hidden" name="EndCob_ddd_com" id="EndCob_ddd_com" value="<%=EndCob_ddd_com%>" />
<input type="hidden" name="EndCob_tel_com" id="EndCob_tel_com" value="<%=EndCob_tel_com%>" />
<input type="hidden" name="EndCob_ramal_com" id="EndCob_ramal_com" value="<%=EndCob_ramal_com%>" />
<input type="hidden" name="EndCob_ddd_com_2" id="EndCob_ddd_com_2" value="<%=EndCob_ddd_com_2%>" />
<input type="hidden" name="EndCob_tel_com_2" id="EndCob_tel_com_2" value="<%=EndCob_tel_com_2%>" />
<input type="hidden" name="EndCob_ramal_com_2" id="EndCob_ramal_com_2" value="<%=EndCob_ramal_com_2%>" />
<input type="hidden" name="EndCob_ddd_cel" id="EndCob_ddd_cel" value="<%=EndCob_ddd_cel%>" />
<input type="hidden" name="EndCob_tel_cel" id="EndCob_tel_cel" value="<%=EndCob_tel_cel%>" />
<input type="hidden" name="EndCob_cnpj_cpf" id="EndCob_cnpj_cpf" value="<%=EndCob_cnpj_cpf%>" />
<input type="hidden" name="EndCob_contribuinte_icms_status" id="EndCob_contribuinte_icms_status" value="<%=EndCob_contribuinte_icms_status%>" />
<input type="hidden" name="EndCob_produtor_rural_status" id="EndCob_produtor_rural_status" value="<%=EndCob_produtor_rural_status%>" />
<input type="hidden" name="EndCob_ie" id="EndCob_ie" value="<%=EndCob_ie%>" />
<input type="hidden" name="EndCob_rg" id="EndCob_rg" value="<%=EndCob_rg%>" />
<input type="hidden" name="EndCob_contato" id="EndCob_contato" value="<%=EndCob_contato%>" />
<% end if %>


<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>


<% if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then %>
<!--  DADOS DO MAGENTO  -->
<table class="Qx" cellspacing="0">
	<tr style="background-color:azure;">
		<td colspan="3" class="MC MB ME MD" align="center"><span class="N">Dados do Magento (pedido nº <%=c_numero_magento%>)</span></td>
	</tr>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">Cliente</span></td>
		<td class="MB MD TdCliCel">
			<span class="C"><%=s_nome_cliente%></span>
			<% if c_mag_cpf_cnpj_identificado <> "" then %>
			<br /><span class="C"><%=cnpj_cpf_formata(c_mag_cpf_cnpj_identificado)%></span>
			<% end if %>
		</td>
	</tr>

	<% if (c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "1") Or (c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "9") then 
			vl_total_produto_magento = 0
			vl_total_servico_magento = 0
			if Trim("" & tMAP_XML("magento_api_versao")) = CStr(VERSAO_API_MAGENTO_V1_SOAP_XML) then
				vl_total_produto_magento = converte_numero(tMAP_XML("grand_total")) - vl_frete_magento
			elseif Trim("" & tMAP_XML("magento_api_versao")) = CStr(VERSAO_API_MAGENTO_V2_REST_JSON) then
				s = "SELECT " & _
						"tMAP_ITEM.*" & _
					" FROM t_MAGENTO_API_PEDIDO_XML tMAP" & _
						" INNER JOIN t_MAGENTO_API_PEDIDO_XML_DECODE_ITEM tMAP_ITEM ON (tMAP.id = tMAP_ITEM.id_magento_api_pedido_xml)" & _
					" WHERE" & _
						" (tMAP_ITEM.id_magento_api_pedido_xml = " & id_magento_api_pedido_xml & ")" & _
						" AND (" & _
							"(tMAP.magento_api_versao = " & VERSAO_API_MAGENTO_V2_REST_JSON & ") AND (tMAP_ITEM.product_type IN ('" & COD_MAGENTO_PRODUCT_TYPE__SIMPLE & "', '" & COD_MAGENTO_PRODUCT_TYPE__VIRTUAL & "'))" & _
							")" & _
					" ORDER BY" & _
						" tMAP_ITEM.id"
				if tMAP_ITEM.State <> 0 then tMAP_ITEM.Close
				tMAP_ITEM.open s, cn
				do while Not tMAP_ITEM.Eof
					if (UCase(Trim("" & tMAP_ITEM("product_type"))) = UCase(COD_MAGENTO_PRODUCT_TYPE__SIMPLE)) OR (Trim("" & tMAP_ITEM("product_type")) = "") then
						if blnFlagCadSemiAutoPedMagentoUsarCamposValorMktpDataSource And (tMAP_XML("mktp_datasource_status") = 1) then
							'O campo mktp_datasource_special_price informa o valor unitário do item já contabilizando o desconto
							vl_total_produto_magento = vl_total_produto_magento + (converte_numero(tMAP_ITEM("qty_ordered")) * converte_numero(tMAP_ITEM("mktp_datasource_special_price")))
						else
							'O campo row_total informa o valor total do item já calculado com os descontos e multiplicado pela quantidade
							vl_total_produto_magento = vl_total_produto_magento + converte_numero(tMAP_ITEM("row_total"))
							end if
					elseif UCase(Trim("" & tMAP_ITEM("product_type"))) = UCase(COD_MAGENTO_PRODUCT_TYPE__VIRTUAL) then
						vl_total_servico_magento = vl_total_servico_magento + converte_numero(tMAP_ITEM("row_total"))
						end if
					tMAP_ITEM.MoveNext
					loop
				end if 'elseif Trim("" & tMAP_XML("magento_api_versao")) = CStr(VERSAO_API_MAGENTO_V2_REST_JSON)
	%>

	<% if blnMagentoPedidoComIndicador then %>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">Indicador</span></td>
		<td class="MB MD TdCliCel">
			<span class="C"><%=cnpj_cpf_formata(c_mag_installer_document)%></span>
			<br /><span class="C"><%=sIdIndicador & " - " & sNomeIndicador%></span>
		</td>
	</tr>
	<% end if %>
	
	<% if Trim("" & tMAP_XML("b2b_type_order")) = COD_MAGENTO_TYPE_ORDER__INSTALLER then
			'Esta implementação do pedido Magento com indicador é referente ao projeto em Magento 2 (B2B) do Arclube
			if vl_total_produto_magento <> 0 then
				percComissionPercentage = 100 * (converte_numero(tMAP_XML("b2b_installer_commission_value")) / vl_total_produto_magento)
			else
				percComissionPercentage = 0
				end if
			c_perc_RT = formata_perc_RT(percComissionPercentage)
	%>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">Parceiro</span></td>
		<td class="MB MD TdCliCel">
			<span class="C"><%="ID: " & Trim("" & tMAP_XML("b2b_installer_id"))%></span>
			<br /><span class="C"><%=Trim("" & tMAP_XML("b2b_installer_name"))%></span>
			<br /><span class="C"><%="Comissão (" & formata_perc_RT(tMAP_XML("b2b_installer_commission_percentage")) & "%): " & SIMBOLO_MONETARIO & " " & formata_moeda(tMAP_XML("b2b_installer_commission_value"))%></span>
		</td>
	</tr>
	<% end if %>

	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">Indicador</span></td>
		<td class="MB MD TdCliCel">
			<span class="C"><%=c_indicador%></span>
		</td>
	</tr>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">VL Frete</span></td>
		<td class="MB MD TdCliCel">
			<span class="C"><%=formata_moeda(vl_frete_magento)%></span>
		</td>
	</tr>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">VL Produtos c/ Desc</span></td>
		<td class="MB MD TdCliCel">
			<span class="C"><%=formata_moeda(vl_total_produto_magento)%></span>
			<input type="hidden" name="c_vl_total_produto_magento" id="c_vl_total_produto_magento" value="<%=formata_moeda(vl_total_produto_magento)%>" />
		</td>
	</tr>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">VL Serviços c/ Desc</span></td>
		<td class="MB MD TdCliCel">
			<span class="C"><%=formata_moeda(vl_total_servico_magento)%></span>
		</td>
	</tr>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">VL Total (produtos e frete)</span></td>
		<td class="MB MD TdCliCel">
			<span class="C" style="color:blue;"><%=formata_moeda(vl_total_produto_magento+vl_frete_magento)%></span>
		</td>
	</tr>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">VL Total (produtos, serviços e frete)</span></td>
		<td class="MB MD TdCliCel">
			<span class="C"><%=formata_moeda(vl_total_produto_magento+vl_total_servico_magento+vl_frete_magento)%></span>
		</td>
	</tr>

	<% end if %>
</table>
<% end if %>

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->  
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pedido Novo</span></td>
</tr>
</table>
<br>


<% if alerta_informativo <> "" then %>
<table class="Qx" cellspacing="0" style="width:649px;border:solid 2px black;">
	<tr>
		<td class="ALERTA" align="center"><span style="font-size:9pt;">AVISO IMPORTANTE!</span><br /><span style="font-size:9pt;"><%=alerta_informativo %></span></td>
	</tr>
</table>
<br />
<% end if %>

<% if qtde_empresa_selecionada > 1 then %>
<table class="Qx" cellspacing="0" style="width:649px;border:solid 2px black;">
	<tr>
		<td class="ALERTA" align="center"><span style="font-size:9pt;">ATENÇÃO!</span><br /><span style="font-size:9pt;">Ao cadastrar este pedido irá ocorrer um split automático!</span></td>
	</tr>
</table>
<br />
<% end if %>

<% if erro_produto_indisponivel then %>
<!--  RELAÇÃO DE PRODUTOS SEM PRESENÇA NO ESTOQUE -->
<table class="Qx" cellspacing="0" style="width:649px;">
	<tr><td class="MB ALERTA" colspan="6" align="center"><span class="ALERTA" style="font-size:9pt;">PRODUTOS SEM PRESENÇA NO ESTOQUE</span></td></tr>
	<tr bgColor="#FFFFFF">
	<td class="MDBE" align="left"><span class="PLTe">Fabr</span></td>
	<td class="MDB" align="left"><span class="PLTe">Produto</span></td>
	<td class="MDB" align="left"><span class="PLTe">Descrição</span></td>
	<td class="MDB" align="right"><span class="PLTd">Solicitado</span></td>
	<td class="MDB" align="right"><span class="PLTd">Disponível</span></td>
	<td class="MDB" align="right"><span class="PLTd">Faltam</span></td>
	</tr>

<%
	for i=Lbound(v_item) to Ubound(v_item)
		if Trim(v_item(i).produto) <> "" then
			with v_item(i)
				if .qtde > .qtde_estoque_total_disponivel then
%>
			<tr>
			<td class="MDBE" align="left"><input name="c_spe_fabricante" id="c_spe_fabricante" class="PLLe" style="width:26px;"
				value='<%=.fabricante%>' readonly tabindex=-1></td>
			<td class="MDB" align="left"><input name="c_spe_produto" id="c_spe_produto" class="PLLe" style="width:55px;"
				value='<%=.produto%>' readonly tabindex=-1></td>
			<td class="MDB" align="left">
				<span class="PLLe" style="width:333px;"><%=produto_formata_descricao_em_html(.descricao_html)%></span>
				<input type="hidden" name="c_spe_descricao" id="c_spe_descricao" value='<%=.descricao%>'>
			</td>
			<td class="MDB" align="right"><input name="c_spe_qtde_solicitada" id="c_spe_qtde_solicitada" class="PLLd" style="width:70px;"
				value='<%=Cstr(.qtde)%>' readonly tabindex=-1></td>
			<td class="MDB" align="right"><input name="c_spe_qtde_estoque" id="c_spe_qtde_estoque" class="PLLd" style="width:70px;"
				value='<%=Cstr(.qtde_estoque_total_disponivel)%>' readonly tabindex=-1></td>
			<td class="MDB" align="right"><input name="c_spe_saldo" id="c_spe_saldo" class="PLLd" style="width:70px;color:red;"
				value='<%=Cstr(Abs(.qtde_estoque_total_disponivel - .qtde))%>' readonly tabindex=-1></td>
			</tr>
<%
					end if
				end with
			end if
		next
%>
</table>
<% end if %>

<% if Not (erro_produto_indisponivel And bloquear_cadastramento_quando_produto_indiponivel) then %>
<br>
<br>
<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0">
	<tr bgColor="#FFFFFF">
	<% if (permite_RA_status = 1) And (rb_RA = "S") then nColSpan=5 else nColSpan=4 %>
	<td colspan="<%=CStr(nColSpan)%>" align="left">&nbsp;</td>
	<td colspan="2" align="right"><span class="PLTe">Desc Linear (%)&nbsp;<input name="c_desc_linear" id="c_desc_linear" class="Cd" style="width:36px;" 
		onkeypress="if (digitou_enter(true)){this.value=formata_perc_desc_linear(this.value);fPED.btnDescLinear.focus();} filtra_percentual();"
		onblur="this.value=formata_perc_desc_linear(this.value);"
		/></span></td>
	<td colspan="2" align="left"><input type="button" name="btnDescLinear" id="btnDescLinear" class="Button" onclick="atualiza_itens_com_desc_linear();" value="Aplicar" title="aplicar o desconto em todos os itens" style="margin-left:1px;margin-bottom:2px;" /></td>
	</tr>
	<tr bgColor="#FFFFFF">
	<% if (permite_RA_status = 1) And (rb_RA = "S") then nColSpan=9 else nColSpan=8 %>
	<td colspan="<%=CStr(nColSpan)%>" align="left" style="height:6px;"></td>
	</tr>
	<tr bgColor="#FFFFFF">
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Produto</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Descrição</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Qtde</span></td>
	<% if (permite_RA_status = 1) And (rb_RA = "S") then %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Preço</span></td>
	<% end if %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Lista</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Desc%</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Venda</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Total</span></td>
	</tr>

<%	qtdeColProd = 8
	if (permite_RA_status = 1) And (rb_RA = "S") then qtdeColProd = qtdeColProd + 1
	m_TotalDestePedido=0
	m_TotalDestePedidoComRA=0
	n = Lbound(v_item)-1
	for i=1 to MAX_ITENS 
		s_readonly = "readonly tabindex=-1"
		s_vl_NF_readonly = "readonly tabindex=-1"
		n = n+1
		if n <= Ubound(v_item) then
			with v_item(n)
				s_fabricante=.fabricante
				s_produto=.produto
				s_descricao=.descricao
				s_descricao_html=produto_formata_descricao_em_html(.descricao_html)
				s_qtde=.qtde
				s_preco_lista=formata_moeda(.preco_lista)
				s_preco_venda = s_preco_lista
				
				s_readonly = ""
				if (permite_RA_status = 1) And (rb_RA = "S") then
					s_vl_NF_readonly = ""
					s_vl_NF = s_preco_venda
				else
					s_vl_NF = ""
					end if

				if blnExecutarCadSemiAutoPedMagentoRateioFreteAutomatico then
					for j=LBound(vItemCadSemiAutoPedMageRateioFreteConsolidado) to UBound(vItemCadSemiAutoPedMageRateioFreteConsolidado)
						if (.fabricante = vItemCadSemiAutoPedMageRateioFreteConsolidado(j).fabricante) And (.produto = vItemCadSemiAutoPedMageRateioFreteConsolidado(j).produto) then
							s_preco_venda = formata_moeda(vItemCadSemiAutoPedMageRateioFreteConsolidado(j).preco_venda_medio)
							if (permite_RA_status = 1) And (rb_RA = "S") then
								s_vl_NF = formata_moeda(vItemCadSemiAutoPedMageRateioFreteConsolidado(j).preco_nf_medio)
							else
								s_vl_NF = ""
								end if
							exit for
							end if
						next
					end if 'if blnExecutarCadSemiAutoPedMagentoRateioFreteAutomatico

				m_TotalItem=.qtde * converte_numero(s_preco_venda)
			'	INICIALMENTE, O PRECO_NF É O MESMO VALOR DO PRECO_LISTA, FICANDO DIFERENTE APENAS SE FOR EDITADO
				if converte_numero(s_vl_NF) > 0 then
					m_TotalItemComRA=.qtde * converte_numero(s_vl_NF)
				else
					m_TotalItemComRA=.qtde * .preco_lista
					end if
				s_vl_TotalItem=formata_moeda(m_TotalItem)
				m_TotalDestePedido=m_TotalDestePedido + m_TotalItem
				m_TotalDestePedidoComRA=m_TotalDestePedidoComRA + m_TotalItemComRA
				end with
		else
			s_fabricante=""
			s_produto=""
			s_descricao=""
			s_descricao_html=""
			s_qtde=""
			s_preco_lista=""
			s_preco_venda=""
			s_vl_NF=""
			s_vl_TotalItem=""
			end if
%>
	<tr>
	<td class="MDBE" align="left">
		<input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:26px;"
			value='<%=s_fabricante%>' readonly tabindex=-1 />
	</td>
	<td class="MDB" align="left">
		<input name="c_produto" id="c_produto" class="PLLe" style="width:55px;"
			value='<%=s_produto%>' readonly tabindex=-1 />
	</td>
	<td class="MDB" align="left" style="width:277px;">
		<span class="PLLe"><%=s_descricao_html%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value='<%=s_descricao%>' />
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde" class="PLLd" style="width:27px;"
			value='<%=s_qtde%>' readonly tabindex=-1 />
	</td>
	<% if (permite_RA_status = 1) And (rb_RA = "S") then %>
	<td class="MDB" align="right">
		<input name="c_vl_NF" id="c_vl_NF" class="PLLd" style="width:62px;"
			onkeypress="if (digitou_enter(true)) fPED.c_vl_unitario[<%=Cstr(i-1)%>].focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); trata_edicao_RA(<%=Cstr(i-1)%>); recalcula_RA(); recalcula_RA_Liquido(); recalcula_parcelas();"
			value='<%=s_vl_NF%>'
			<%=s_vl_NF_readonly%>
			/>
	</td>
	<% else %>
	<input type="hidden" name="c_vl_NF" id="c_vl_NF" value='<%=s_vl_NF%>'>
	<% end if %>
	<td class="MDB" align="right">
		<input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px;"
			value='<%=s_preco_lista%>' readonly tabindex=-1 />
	</td>
	<td class="MDB" align="right">
		<input name="c_desc" id="c_desc" class="PLLd" style="width:36px;" value=""
		<% if blnLojaHabilitadaProdCompostoECommerce then %>
			<%=s_readonly%>
			onkeypress="if (digitou_enter(true)){fPED.c_vl_unitario[<%=Cstr(i-1)%>].focus();} filtra_percentual();"
			onblur="this.value=formata_perc_desconto(this.value); calcula_desconto(<%=Cstr(i-1)%>); trata_edicao_RA(<%=Cstr(i-1)%>); recalcula_total_linha(<%=Cstr(i)%>); recalcula_RA(); recalcula_RA_Liquido();"
		<% else %>
			readonly tabindex=-1
		<% end if %>
		/>
	</td>
	<td class="MDB" align="right">
		<% if blnLojaHabilitadaProdCompostoECommerce then s_campo_focus="c_desc" else s_campo_focus="c_vl_unitario"%>
		<input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px;"
			onkeypress="if (digitou_enter(true)) {if ((<%=Cstr(i)%>==fPED.c_vl_unitario.length)||(trim(fPED.c_produto[<%=Cstr(i)%>].value)=='')) fPED.c_obs1.focus(); else <% if (permite_RA_status = 1) And (rb_RA = "S") then Response.Write "fPED.c_vl_NF" else Response.Write "fPED." & s_campo_focus%>[<%=Cstr(i)%>].focus();} filtra_moeda_positivo();"
			onblur="this.value=formata_moeda(this.value); trata_edicao_RA(<%=Cstr(i-1)%>); recalcula_total_linha(<%=Cstr(i)%>); recalcula_RA(); recalcula_RA_Liquido(); recalcula_parcelas();"
			value='<%=s_preco_venda%>'
			<%=s_readonly%>
			/>
	</td>
	<td class="MDB" align="right">
		<input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px;" 
		value='<%=s_vl_TotalItem%>' readonly tabindex=-1 />
	</td>
	</tr>
<% next %>
	<tr>
	<td colspan="3" align="left">
		<table cellspacing="0" cellpadding="0" width='100%' style="margin-top:4px;">
		<tr>
			<td width="20%" align="left">&nbsp;</td>
			<% if (permite_RA_status = 1) And (rb_RA = "S") then %>
			<td align="right">
			<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
				<tr>
				<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;RA Líquido</span></td>
				<td class="MTBD" align="right">
					<input name="c_total_RA_Liquido" id="c_total_RA_Liquido" class="PLLd" style="width:70px;color:blue;" 
						value="" readonly tabindex=-1 />
				</td>
				</tr>
			</table>
			</td>
			<td align="right">
			<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
				<tr>
				<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;RA Bruto</span></td>
				<td class="MTBD" align="right">
					<input name="c_total_RA" id="c_total_RA" class="PLLd" style="width:70px;color:blue;"
						value="" readonly tabindex=-1 />
				</td>
				</tr>
			</table>
			</td>
			<% else %>
			<input type="hidden" name="c_total_RA_Liquido" id="c_total_RA_Liquido" value=''>
			<input type="hidden" name="c_total_RA" id="c_total_RA" value=''>
			<% end if %>
			
			<td align="right">
				<table cellspacing="0" cellpadding="0">
					<tr>
					<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;COM(%)</span></td>
					<td class="MTBD" align="right">
						<input name="c_perc_RT" id="c_perc_RT" class="PLLd" style="width:30px;color:blue;"
							value='<%=c_perc_RT%>' readonly tabindex=-1 />
					</td>
					</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
	<% if (permite_RA_status = 1) And (rb_RA = "S") then %>
	<td class="MD" align="left">&nbsp;</td>
	<td class="MDB" align="right">
		<%	if rb_RA = "S" then 
				s_TotalDestePedidoComRA=formata_moeda(m_TotalDestePedidoComRA)
			else
				s_TotalDestePedidoComRA=""
			end if
			%>
		<input name="c_total_NF" id="c_total_NF" class="PLLd" style="width:70px;color:blue;" 
				value='<%=s_TotalDestePedidoComRA%>' readonly tabindex=-1 />
	</td>
	<% else %>
	<td align="left">&nbsp;</td>
	<input type="hidden" name="c_total_NF" id="c_total_NF" value='<%=s_TotalDestePedidoComRA%>'>
	<% end if %>

	<td class="MD" align="left">&nbsp;</td>
	<td class="MDB" align="right"><input name="c_desc_medio_total" id="c_desc_medio_total" class="PLLd" style="width:36px;color:blue;" readonly tabindex=-1 /></td>
	<td class="MD" align="left">&nbsp;</td>

	<td class="MDB" align="right">
		<input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:blue;" 
			value='<%=formata_moeda(m_TotalDestePedido)%>' readonly tabindex=-1 />
	</td>
	</tr>

	<%
		if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
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
			if Not tMAP_ITEM.Eof then
				m_TotalServicos = 0
	%>
	<tr><td colspan="<%=CStr(qtdeColProd)%>">&nbsp;</td></tr>
	<tr><td class="MB" colspan="<%=CStr(qtdeColProd)%>" align="left"><span class="PLTe">Serviços</span></td></tr>
	<%
				do while Not tMAP_ITEM.Eof
					vl_servico_original_price = converte_numero(tMAP_ITEM("original_price"))
					'O campo discount_amount informa o valor total do desconto já multiplicado pela quantidade, ou seja, não há campo com o valor unitário do desconto aplicado e
					'nem o valor unitário de venda do item já com o desconto aplicado.
					'Já o campo row_total informa o valor total do item já calculado com os descontos e multiplicado pela quantidade
					vl_servico_price = 0
					if converte_numero(tMAP_ITEM("qty_ordered")) > 0 then
						vl_servico_price = converte_numero(tMAP_ITEM("row_total")) / converte_numero(tMAP_ITEM("qty_ordered"))
						end if
	%>
	<tr>
		<td class="MB ME" align="left">
			&nbsp;
		</td>
		<td class="MDB" align="left">
			<input name="c_servico_sku" class="PLLe" style="width:55px;" value="<%=Trim("" & tMAP_ITEM("sku"))%>" readonly tabindex="-1" />
		</td>
		<td class="MDB" align="left" style="width:277px;">
			<input name="c_servico_descricao" class="PLLe" style="width:277px;" value="<%=Trim("" & tMAP_ITEM("name"))%>" readonly tabindex="-1" />
		</td>
		<td class="MDB" align="right">
			<input name="c_servico_qtde" class="PLLd" style="width:27px;" value="<%=Trim("" & tMAP_ITEM("qty_ordered"))%>" readonly tabindex="-1" />
		</td>
		<% if (permite_RA_status = 1) And (rb_RA = "S") then %>
		<td class="MDB" align="right">
			<input name="c_servico_vl_NF" class="PLLd" style="width:62px;" value="<%=formata_moeda(vl_servico_price)%>" readonly tabindex="-1" />
		</td>
		<% end if %>
		<td class="MDB" align="right">
			<input name="c_servico_preco_lista" class="PLLd" style="width:62px;" value="<%=formata_moeda(vl_servico_original_price)%>" readonly tabindex="-1" />
		</td>
		<td class="MDB" align="right">
		<% percDescServico = 0
			sPercDescServico = ""
			sColorPercDescServico = "green"
			if vl_servico_original_price <> 0 then
				percDescServico = 100*((vl_servico_original_price - vl_servico_price)/vl_servico_original_price)
				if percDescServico <> 0 then sPercDescServico = formata_perc(percDescServico)
				if percDescServico < 0 then sColorPercDescServico = "red"
				end if%>
			<input name="c_servico_desc" class="PLLd" style="width:36px;color:<%=sColorPercDescServico%>;" value="<%=sPercDescServico%>" readonly tabindex="-1" />
		</td>
		<td class="MDB" align="right">
			<input name="c_servico_vl_unitario" class="PLLd" style="width:62px;" value="<%=formata_moeda(vl_servico_price)%>" readonly tabindex="-1" />
		</td>
		<td class="MDB" align="right">
			<input name="c_servico_vl_total" class="PLLd" style="width:70px;" value="<%=formata_moeda(tMAP_ITEM("row_total"))%>" readonly tabindex="-1" />
		</td>
	</tr>
	<%
				m_TotalServicos = m_TotalServicos + converte_numero(tMAP_ITEM("row_total"))
				tMAP_ITEM.MoveNext
				loop %>
	<tr>
		<td class="MD" colspan="<%=CStr(qtdeColProd-1)%>">&nbsp;</td>
		<td class="MDB" align="right"><input name="c_total_servicos" id="c_total_servicos" class="PLLd" style="width:70px;color:blue;" value="<%=formata_moeda(m_TotalServicos)%>" readonly tabindex="-1" /></td>
	</tr>
	<%
			end if 'if Not tMAP_ITEM.Eof
		end if %>
</table>

<%	intColSpan=3
	if operacao_permitida(OP_LJA_EXIBIR_CAMPO_INSTALADOR_INSTALA_AO_CADASTRAR_NOVO_PEDIDO, s_lista_operacoes_permitidas) then intColSpan = intColSpan + 1
    if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then intColSpan = intColSpan + 1
	if rb_indicacao = "S" then intColSpan = intColSpan + 1
%>
<br>
<table class="Q" style="width:649px;" cellspacing="0">
	<tr>
		<td class="MB" colspan="<%=Cstr(intColSpan)%>" align="left"><p class="Rf">Observações</p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:641px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS1);" onblur="this.value=trim(this.value);"
				></textarea>
		</td>
	</tr>
    <tr>
		<%
			s_value = ""
			if (c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "1") Or (c_FlagCadSemiAutoPedMagento_FluxoOtimizado = "9") then
				'Colocar a informação do ponto de referência no campo 'Constar na NF'.
				'Comparar o conteúdo do ponto de referência com o campo complemento. Se forem iguais, não colocar em 'Constar na NF'.
				'Se o campo complemento exceder o tamanho do BD e precisar ser truncado, copiá-lo no campo 'Constar na NF', junto com o ponto de referência.
				if rb_end_entrega = "S" then
					'Texto do complemento do endereço será truncado
					if Len(EndEtg_endereco_complemento) > MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO then
						if s_value <> "" then s_value = s_value & vbCrLf
						s_value = s_value & "Complemento do endereço: " & EndEtg_endereco_complemento
						end if
					'Texto do ponto de referência é diferente do texto do complemento do endereço
					if (Ucase(Trim(EndEtg_endereco_complemento)) <> Ucase(Trim(EndEtg_endereco_ponto_referencia))) And (Trim(EndEtg_endereco_ponto_referencia) <> "") then
						if s_value <> "" then s_value = s_value & vbCrLf
						s_value = s_value & "Ponto de referência: " & EndEtg_endereco_ponto_referencia
						end if
				else
					'Texto do complemento do endereço será truncado
					if Len(EndCob_endereco_complemento) > MAX_TAMANHO_CAMPO_ENDERECO_COMPLEMENTO then
						if s_value <> "" then s_value = s_value & vbCrLf
						s_value = s_value & "Complemento do endereço: " & EndCob_endereco_complemento
						end if
					'Texto do ponto de referência é diferente do texto do complemento do endereço
					if (Ucase(Trim(EndCob_endereco_complemento)) <> Ucase(Trim(EndCob_endereco_ponto_referencia))) And (Trim(EndCob_endereco_ponto_referencia) <> "") then
						if s_value <> "" then s_value = s_value & vbCrLf
						s_value = s_value & "Ponto de referência: " & EndCob_endereco_ponto_referencia
						end if
					end if
				end if
		%>
		<td class="MB" colspan="<%=Cstr(intColSpan)%>" align="left"><p class="Rf">Constar na NF</p>
			<textarea name="c_nf_texto" id="c_nf_texto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_NF_TEXTO_CONSTAR)%>" 
				style="width:641px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_NF_TEXTO);" onblur="this.value=trim(this.value);"
				><%=s_value%></textarea>
		</td>
	</tr>
	<tr>
		<td class="MB MD" align="left" nowrap><p class="Rf">Nº Nota Fiscal</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" maxlength="10" style="width:100px;margin-left:2pt;" onkeypress="filtra_nome_identificador();" onblur="this.value=trim(this.value);"
				value='' readonly />
		</td>
        <%if (loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE) Or blnMagentoPedidoComIndicador then
				s_value = ""
				if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
					s_value = c_numero_magento
					end if
		%>
        <td class="MB MD" align="left" nowrap><p class="Rf">Número Magento</p>
			<input name="c_pedido_ac" id="c_pedido_ac" class="PLLe" maxlength="9" style="width:100px;margin-left:2pt;" onkeypress="filtra_nome_identificador();return SomenteNumero(event)" onblur="this.value=trim(this.value);"
				value='<%=s_value%>'>
		</td>
        <%end if %>
		<td class="MB MD" align="left" nowrap><p class="Rf">Entrega Imediata</p>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata"
				value="<%=COD_ETG_IMEDIATA_NAO%>" /><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[0].click();">Não</span>
			<% s_checked = ""
				if Cstr(loja)=NUMERO_LOJA_ECOMMERCE_AR_CLUBE then s_checked = " checked" %>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
				value="<%=COD_ETG_IMEDIATA_SIM%>" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[1].click();">Sim</span>
		</td>
		<td class="MB" align="left" nowrap><p class="Rf">Bem de Uso/Consumo&nbsp;</p>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				value="<%=COD_ST_BEM_USO_CONSUMO_NAO%>"><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[0].click();">Não</span>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				value="<%=COD_ST_BEM_USO_CONSUMO_SIM%>" <%if Cstr(loja)=NUMERO_LOJA_ECOMMERCE_AR_CLUBE then Response.write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[1].click();">Sim</span>
		</td>
		<% if operacao_permitida(OP_LJA_EXIBIR_CAMPO_INSTALADOR_INSTALA_AO_CADASTRAR_NOVO_PEDIDO, s_lista_operacoes_permitidas) then %>
		<td class="MB ME" align="left" nowrap><p class="Rf">Instalador Instala</p>
			<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
				value="<%=COD_INSTALADOR_INSTALA_NAO%>" <%if Cstr(loja)=NUMERO_LOJA_ECOMMERCE_AR_CLUBE then Response.write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_instalador_instala[0].click();">Não</span>
			<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
				value="<%=COD_INSTALADOR_INSTALA_SIM%>"><span class="C" style="cursor:default" onclick="fPED.rb_instalador_instala[1].click();">Sim</span>
		</td>
		<% end if %>
	<% if rb_indicacao = "S" then %>
		<td class="MB ME tdGarInd" align="left" nowrap><p class="Rf">Garantia Indicador</p>
			<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" class="rbGarIndNao"
				value="<%=COD_GARANTIA_INDICADOR_STATUS__NAO%>" <%if Cstr(loja)=NUMERO_LOJA_ECOMMERCE_AR_CLUBE then Response.write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[0].click();">Não</span>
			<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador"
				value="<%=COD_GARANTIA_INDICADOR_STATUS__SIM%>"><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[1].click();">Sim</span>
		</td>
	<% end if %>
	</tr>
    <tr>
        <td class="MD" align="left" valign="top" nowrap>
			<p class="Rf">xPed</p>
			<input name="c_num_pedido_compra" id="c_num_pedido_compra" class="PLLe" maxlength="15" style="width:100px;padding-top:10px;margin-left:2pt;" onkeypress="filtra_nome_identificador();" onblur="this.value=trim(this.value);"
				value=''>
		</td>
		<td align="left" colspan="4">
			<p class="Rf">Previsão de Entrega</p>
			<input type="text" class="PLLc" name="c_data_previsao_entrega" id="c_data_previsao_entrega" maxlength="10" style="width:90px;" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="filtra_data();" />
		</td>
    </tr>
    <% if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
			s_value = ""
			if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s_value = c_numero_marketplace
				end if
	%>
    <tr>
        <td class="MC MD" align="left" nowrap valign="top"><p class="Rf">Nº Pedido Marketplace</p>
			<input name="c_numero_mktplace" id="c_numero_mktplace" class="PLLe" maxlength="20" style="width:135px;margin-left:2pt;margin-top:5px;" onkeypress="filtra_nome_identificador();return SomenteNumeroHifen(event)" onblur="this.value=trim(this.value);"
				value='<%=s_value%>'>
		</td>
        <td class="MC" colspan="4" align="left" nowrap valign="top"><p class="Rf">Origem do Pedido</p>
			<select name="c_origem_pedido" id="c_origem_pedido" style="margin: 3px; 3px; 3px">
			<% if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then %>
				<%=origem_pedido_monta_itens_select(c_marketplace_codigo_origem) %>
			<% else %>
                <%=origem_pedido_monta_itens_select(Null) %>
			<% end if %>
			</select>
		</td>
    </tr>
    <% end if %>
</table>

<!--  NOVA VERSÃO DA FORMA DE PAGAMENTO  -->
<br>
<table class="Q" style="width:649px;" cellspacing="0">
  <tr>
	<td align="left">
	  <p class="Rf">Forma de Pagamento</p>
	</td>
  </tr>  
  <tr>
	<td align="left">
	  <table width="100%" cellspacing="0" cellpadding="4" border="0">
		<!--  À VISTA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left">
				  <% intIdx = 0 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_A_VISTA%>" 
						<%if c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA then Response.Write " checked"%>
						onclick="recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
						><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">À Vista</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <select id="op_av_forma_pagto" name="op_av_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
					<%	if (rb_indicacao = "S") And (loja <> NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then
							Response.Write forma_pagto_liberada_av_monta_itens_select(Null, c_indicador, EndCob_tipo_pessoa)
						else
							Response.Write forma_pagto_av_monta_itens_select(Null)
							end if %>
				  </select>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELA ÚNICA  -->
		<tr class="TR_FP_PU">
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto"
						value="<%=COD_FORMA_PAGTO_PARCELA_UNICA%>"
						<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) And _
							 (converte_numero(c_custoFinancFornecQtdeParcelas)=1) then Response.Write " checked"%>
						onclick="recalculaCustoFinanceiroPrecoLista();pu_atualiza_valor();recalcula_RA_Liquido();"
						><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcela Única</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <select id="op_pu_forma_pagto" name="op_pu_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
					<%	if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
							s_qtde_dias = "30"
							Response.Write forma_pagto_da_parcela_unica_monta_itens_select_EC(ID_FORMA_PAGTO_DEPOSITO)
						else
							s_qtde_dias = ""
							if rb_indicacao = "S" then
								Response.Write forma_pagto_liberada_da_parcela_unica_monta_itens_select(Null, c_indicador, EndCob_tipo_pessoa)
							else
								Response.Write forma_pagto_da_parcela_unica_monta_itens_select(Null)
								end if
							end if%>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pu_valor" id="c_pu_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pu_vencto_apos.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);recalcula_RA_Liquido();" value=''
				  ><span style="width:10px;">&nbsp;</span
				  ><span class="C">vencendo após</span
				  ><input name="c_pu_vencto_apos" id="c_pu_vencto_apos" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();" value='<%=s_qtde_dias%>'
				  ><span class="C">dias</span>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO NO CARTÃO (INTERNET)  -->
		<% if (rb_indicacao = "S") And is_restricao_ativa_forma_pagto(c_indicador, ID_FORMA_PAGTO_CARTAO, EndCob_tipo_pessoa) then %>
		<tr style="display:none;">
		<% else %>
		<tr>
		<% end if %>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELADO_CARTAO%>" 
						onclick="recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
						><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado no Cartão (internet)</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <input name="c_pc_qtde" id="c_pc_qtde" class="Cc" maxlength="2" style="width:30px;"  onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pc_valor.focus(); filtra_numerico();" onblur="pc_calcula_valor_parcela();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();" value=''>
				</td>
				<td align="left"><span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span></td>
				<td align="left">
				  <input name="c_pc_valor" id="c_pc_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);recalcula_RA_Liquido();" value=''>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO NO CARTÃO (MAQUINETA)  -->
		<% if (rb_indicacao = "S") And is_restricao_ativa_forma_pagto(c_indicador, ID_FORMA_PAGTO_CARTAO_MAQUINETA, EndCob_tipo_pessoa) then %>
		<tr style="display:none;">
		<% else %>
		<tr>
		<% end if %>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA%>" 
						onclick="recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
						><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado no Cartão (maquineta)</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <input name="c_pc_maquineta_qtde" id="c_pc_maquineta_qtde" class="Cc" maxlength="2" style="width:30px;"  onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pc_maquineta_valor.focus(); filtra_numerico();" onblur="pc_maquineta_calcula_valor_parcela();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();" value=''>
				</td>
				<td align="left"><span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span></td>
				<td align="left">
				  <input name="c_pc_maquineta_valor" id="c_pc_maquineta_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);recalcula_RA_Liquido();" value=''>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO COM ENTRADA  -->
		<tr>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td colspan="3" align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA%>" 
						<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) then Response.Write " checked"%>
						onclick="recalculaCustoFinanceiroPrecoLista();pce_preenche_sugestao_intervalo();recalcula_RA_Liquido();"
						><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado com Entrada</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">Entrada&nbsp;</span></td>
				<td align="left">
				  <select id="op_pce_entrada_forma_pagto" name="op_pce_entrada_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
					<%	if rb_indicacao = "S" then
							Response.Write forma_pagto_liberada_da_entrada_monta_itens_select(Null, c_indicador, EndCob_tipo_pessoa)
						else
							Response.Write forma_pagto_da_entrada_monta_itens_select(Null)
							end if%>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pce_entrada_valor" id="c_pce_entrada_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.op_pce_prestacao_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); recalcula_RA_Liquido();" value=''>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">Prestações&nbsp;</span></td>
				<td align="left">
				  <select id="op_pce_prestacao_forma_pagto" name="op_pce_prestacao_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
					<%	if rb_indicacao = "S" then
							Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select(Null, c_indicador, EndCob_tipo_pessoa)
						else
							Response.Write forma_pagto_da_prestacao_monta_itens_select(Null)
							end if%>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <input name="c_pce_prestacao_qtde" id="c_pce_prestacao_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pce_prestacao_valor.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();pce_calcula_valor_parcela();" 
					value='<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) then Response.Write c_custoFinancFornecQtdeParcelas%>'
					>
				  <span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pce_prestacao_valor" id="c_pce_prestacao_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pce_prestacao_periodo.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);recalcula_RA_Liquido();" value=''>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td colspan="2" align="left"><span class="C">Parcelas vencendo a cada</span
				><input name="c_pce_prestacao_periodo" id="c_pce_prestacao_periodo" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();" 
					value='<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) then Response.Write "30"%>'
				><span class="C">dias</span
				><span style="width:10px;">&nbsp;</span
				><span class="notPrint"><input name="b_pce_SugereFormaPagto" id="b_pce_SugereFormaPagto" type="button" class="Button" style="visibility:hidden;" onclick="pce_sugestao_forma_pagto();" value="sugestão automática" title="preenche o campo 'Forma de Pagamento' com uma sugestão de texto"></span
				></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO SEM ENTRADA  -->
		<tr class="TR_FP_PSE">
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td colspan="3" align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
					value="<%=COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA%>" 
					<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) And _
						 (converte_numero(c_custoFinancFornecQtdeParcelas)>1) then Response.Write " checked"%>
					onclick="pse_preenche_sugestao_intervalo();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
					><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado sem Entrada</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">1ª Prestação&nbsp;</span></td>
				<td align="left">
				  <select id="op_pse_prim_prest_forma_pagto" name="op_pse_prim_prest_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
					<%	if rb_indicacao = "S" then
							Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select(Null, c_indicador, EndCob_tipo_pessoa)
						else
							Response.Write forma_pagto_da_prestacao_monta_itens_select(Null)
							end if%>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pse_prim_prest_valor" id="c_pse_prim_prest_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pse_prim_prest_apos.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); pse_calcula_valor_parcela(); recalcula_RA_Liquido();" value=''
				  ><span style="width:10px;">&nbsp;</span
				  ><span class="C">vencendo após</span
				  ><input name="c_pse_prim_prest_apos" id="c_pse_prim_prest_apos" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.op_pse_demais_prest_forma_pagto.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();" value=''
				  ><span class="C">dias</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">Demais Prestações&nbsp;</span></td>
				<td align="left">
				  <select id="op_pse_demais_prest_forma_pagto" name="op_pse_demais_prest_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
					<%	if rb_indicacao = "S" then
							Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select(Null, c_indicador, EndCob_tipo_pessoa)
						else
							Response.Write forma_pagto_da_prestacao_monta_itens_select(Null)
							end if%>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <input name="c_pse_demais_prest_qtde" id="c_pse_demais_prest_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pse_demais_prest_valor.focus(); filtra_numerico();" onblur="pse_calcula_valor_parcela();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();" 
						value='<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) And (converte_numero(c_custoFinancFornecQtdeParcelas)>1) then Response.Write Cstr(converte_numero(c_custoFinancFornecQtdeParcelas)-1)%>'
						>
				  <span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pse_demais_prest_valor" id="c_pse_demais_prest_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pse_demais_prest_periodo.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); recalcula_RA_Liquido();" value=''>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td colspan="2" align="left"><span class="C">Parcelas vencendo a cada</span
				><input name="c_pse_demais_prest_periodo" id="c_pse_demais_prest_periodo" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();" 
						value='<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) And (converte_numero(c_custoFinancFornecQtdeParcelas)>1) then Response.Write "30"%>'
				><span class="C">dias</span
				><span style="width:10px;">&nbsp;</span
				><span class="notPrint"><input name="b_pse_SugereFormaPagto" id="b_pse_SugereFormaPagto" type="button" class="Button" style="visibility:hidden;" onclick="pse_sugestao_forma_pagto();" value="sugestão automática" title="preenche o campo 'Forma de Pagamento' com uma sugestão de texto"></span
				></td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
  <tr>
	<td class="MC" align="left">
	  <p class="Rf">Informações Sobre Análise de Crédito</p>
		<textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
			style="width:641px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_FORMA_PAGTO);" onblur="this.value=trim(this.value);"
			></textarea>
	</td>
  </tr>  
</table>


<!--  VENDEDOR EXTERNO: LOJA QUE INDICOU  -->
<% IF Session("vendedor_externo") THEN %>
	<br>
	<table class="Q" style="width:649px;" cellspacing="0">
		<tr><td align="left"><p class="Rf">Loja que fez a indicação</p>
			<input name="loja_indicou" id="loja_indicou" class="PLLd" style="width:30px;" readonly tabindex=-1 
				value='<%=s_loja_indicou%>'>&nbsp;-
			<input name="nome_loja_indicou" id="nome_loja_indicou" class="PLLe" style="width:300px;" readonly tabindex=-1 
				value='<%=s_nome_loja_indicou%>'></td>
		</tr>
	</table>
<% END IF %>

<% end if 'if (Not (erro_produto_indisponivel And bloquear_cadastramento_quando_produto_indiponivel)) %>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<% if erro_produto_indisponivel And bloquear_cadastramento_quando_produto_indiponivel then %>
	<% if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then %>
	<tr>
		<td align="left"><a name="bVOLTAR" id="A1" href="javascript:history.back()" title="volta para página anterior">
			<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
		<td align="right"><div name="dCANCELA" id="dCANCELA">
			<a name="bCANCELA" id="bCANCELA" href="javascript:fOpCancela(fCANCEL)" title="cancela a operação">
			<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></div>
		</td>
	</tr>
	<% else %>
	<tr>
		<td align="center"><a name="bVOLTAR" id="A1" href="javascript:history.back()" title="volta para página anterior">
			<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	</tr>
	<% end if %>
<% else %>
	<tr>
		<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
			<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
		<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
			<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDConfirma(fPED)" title="confirma o novo pedido">
			<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
		</td>
	</tr>
<% end if %>
</table>
</form>

</center>
</body>

<% end if %>

<form id="fCANCEL" name="fCANCEL" method="post" action="resumo.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
</form>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		if tMAP_END_COB.State <> 0 then tMAP_END_COB.Close
		set tMAP_END_COB = nothing

		if tMAP_END_ETG.State <> 0 then tMAP_END_ETG.Close
		set tMAP_END_ETG = nothing
		
		if tMAP_ITEM.State <> 0 then tMAP_ITEM.Close
		set tMAP_ITEM = nothing

		if tMAP_XML.State <> 0 then tMAP_XML.Close
		set tMAP_XML = nothing
		end if

	if tPL.State <> 0 then tPL.Close
	set tPL = nothing

	if tPCI.State <> 0 then tPCI.Close
	set tPCI = nothing

	if t_CLIENTE.State <> 0 then t_CLIENTE.Close
	set t_CLIENTE = nothing

	cn.Close
	set cn = nothing
%>