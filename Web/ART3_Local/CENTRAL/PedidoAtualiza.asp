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
'	  P E D I D O A T U A L I Z A . A S P
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

	dim s, usuario, pedido_selecionado, pedido_base
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

    dim url_origem
    url_origem = Trim(Request("url_origem"))

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s

	dim alerta, blnErroConsistencia
	alerta=""
	blnErroConsistencia=False

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2
	dim msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim r_pedido
	if alerta = "" then
		if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then alerta = msg_erro
		end if
	
	dim rEmailDestinatario
	dim id_email, corpo_mensagem, msg_erro_grava_email, emailSndSvcRemetenteMensagemSistema
	dim r_vendedor, blnEnviarEmailVendedorStatusAnaliseCredito
	blnEnviarEmailVendedorStatusAnaliseCredito = False
	emailSndSvcRemetenteMensagemSistema = getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__MENSAGEM_SISTEMA)

'	FORMA DE PAGAMENTO (NOVA VERS�O)
	dim versao_forma_pagamento 
	dim rb_forma_pagto, op_av_forma_pagto, c_pc_qtde, c_pc_valor, c_pc_maquineta_qtde, c_pc_maquineta_valor
	dim op_pu_forma_pagto, c_pu_valor, c_pu_vencto_apos
	dim op_pce_entrada_forma_pagto, c_pce_entrada_valor, op_pce_prestacao_forma_pagto, c_pce_prestacao_qtde, c_pce_prestacao_valor, c_pce_prestacao_periodo
	dim op_pse_prim_prest_forma_pagto, c_pse_prim_prest_valor, c_pse_prim_prest_apos, op_pse_demais_prest_forma_pagto, c_pse_demais_prest_qtde, c_pse_demais_prest_valor, c_pse_demais_prest_periodo
	dim vlTotalFormaPagto
	dim s_perc_RT, vl_total_RA, vl_total_RA_liquido
	versao_forma_pagamento = Trim(Request.Form("versao_forma_pagamento"))
	vlTotalFormaPagto = 0
	
    dim blnIndicadorEdicaoLiberada
    s = Trim(Request.Form("blnIndicadorEdicaoLiberada"))
    blnIndicadorEdicaoLiberada = CBool(s)

	dim blnObs1EdicaoLiberada
	s = Trim(Request.Form("blnObs1EdicaoLiberada"))
	blnObs1EdicaoLiberada = CBool(s)

	dim blnObs2EdicaoLiberada
	s = Trim(Request.Form("blnObs2EdicaoLiberada"))
	blnObs2EdicaoLiberada = CBool(s)

	dim blnObs3EdicaoLiberada
	s = Trim(Request.Form("blnObs3EdicaoLiberada"))
	blnObs3EdicaoLiberada = CBool(s)

	dim blnFormaPagtoEdicaoLiberada
	s = Trim(Request.Form("blnFormaPagtoEdicaoLiberada"))
	blnFormaPagtoEdicaoLiberada = CBool(s)
	
	dim blnEntregaImediataEdicaoLiberada
	s = Trim(Request.Form("blnEntregaImediataEdicaoLiberada"))
	blnEntregaImediataEdicaoLiberada = CBool(s)
	
	dim blnEntregaImediataNaoSemDataPrevisao
	s = Trim(Request.Form("blnEntregaImediataNaoSemDataPrevisao"))
	blnEntregaImediataNaoSemDataPrevisao = CBool(s)

	dim blnInstaladorInstalaEdicaoLiberada
	s = Trim(Request.Form("blnInstaladorInstalaEdicaoLiberada"))
	blnInstaladorInstalaEdicaoLiberada = CBool(s)
	
	dim blnBemUsoConsumoEdicaoLiberada
	s = Trim(Request.Form("blnBemUsoConsumoEdicaoLiberada"))
	blnBemUsoConsumoEdicaoLiberada = CBool(s)
	
	dim bln_RT_e_RA_EdicaoLiberada
	s = Trim(Request.Form("bln_RT_e_RA_EdicaoLiberada"))
	bln_RT_e_RA_EdicaoLiberada = CBool(s)
	
	dim blnItemPedidoEdicaoLiberada
	s = Trim(Request.Form("blnItemPedidoEdicaoLiberada"))
	blnItemPedidoEdicaoLiberada = CBool(s)
	
	dim blnAnaliseCreditoEdicaoLiberada
	s = Trim(Request.Form("blnAnaliseCreditoEdicaoLiberada"))
	blnAnaliseCreditoEdicaoLiberada = CBool(s)
	
	dim blnPedidoRecebidoStatusEdicaoLiberada
	s = Trim(Request.Form("blnPedidoRecebidoStatusEdicaoLiberada"))
	blnPedidoRecebidoStatusEdicaoLiberada = CBool(s)
	
	dim blnDadosNFeMercadoriasDevolvidasEdicaoLiberada
	s = Trim(Request.Form("blnDadosNFeMercadoriasDevolvidasEdicaoLiberada"))
	blnDadosNFeMercadoriasDevolvidasEdicaoLiberada = CBool(s)
	
	dim blnMarketplaceCodigoOrigemAlterado
	blnMarketplaceCodigoOrigemAlterado = False

	dim s_qtde_parcelas, s_forma_pagto, s_obs1, s_obs2, s_obs2_original, s_obs3, s_obs3_original, s_ped_bonshop, s_indicador, s_pedido_ac, s_pedido_mktplace, s_pedido_origem
    dim s_nf_texto, s_num_pedido_compra
	dim blnAEntregarStatusEdicaoLiberada, c_a_entregar_data_marcada, c_a_entregar_data_marcada_original
	dim s_analise_credito, s_analise_credito_a, s_ac_pendente_vendas_motivo
	dim s_etg_imediata, s_bem_uso_consumo, s_etg_imediata_original, c_data_previsao_entrega
	dim blnUpdate, blnFlag, blnEditou
	dim blnEditouTransp, blnProcessaSelecaoAutoTransp
    dim transportadora_cnpj, blnEditouFrete
	transportadora_cnpj = ""
    blnEditouFrete = False

	s_obs1=Trim(request("c_obs1"))
	s_obs2=Trim(request("c_obs2"))
	s_obs2_original=Trim(request("c_obs2_original"))
	s_obs3=Trim(request("c_obs3"))
	s_obs3_original=Trim(request("c_obs3_original"))
	s_ped_bonshop=Trim(request("pedBonshop"))
	c_a_entregar_data_marcada=Trim(request("c_a_entregar_data_marcada"))
	c_a_entregar_data_marcada_original=Trim(request("c_a_entregar_data_marcada_original"))
	s = Trim(Request.Form("blnAEntregarStatusEdicaoLiberada"))
	blnAEntregarStatusEdicaoLiberada = CBool(s)
	s_analise_credito=Trim(request("rb_analise_credito"))
	s_etg_imediata=Trim(request("rb_etg_imediata"))
	c_data_previsao_entrega = Trim(Request("c_data_previsao_entrega"))
	s_bem_uso_consumo=Trim(request("rb_bem_uso_consumo"))
	s_forma_pagto=Trim(request("c_forma_pagto"))
    s_indicador = Trim(Request("c_indicador"))
    s_pedido_ac = Trim(Request("c_pedido_ac"))
    s_pedido_mktplace = Trim(Request("c_numero_mktplace"))
    s_pedido_origem = Trim(Request("c_origem_pedido"))
    s_nf_texto = Trim(Request("c_nf_texto"))
    s_num_pedido_compra = Trim(Request("c_num_pedido_compra"))
    s_ac_pendente_vendas_motivo = Trim(Request("c_pendente_vendas_motivo"))

' BUG:	if s_pedido_mktplace = "" then s_pedido_origem = ""

'	PARA PEDIDOS DO ARCLUBE, � PERMITIDO FICAR SEM O N� MAGENTO SOMENTE NOS SEGUINTES CASOS:
'		1) PEDIDO ORIGINADO PELO TELEVENDAS
'		2) PEDIDO GERADO CONTRA A TRANSPORTADORA (EM CASOS QUE A TRANSPORTADORA SE RESPONSABILIZA PELA REPOSI��O DE MERCADORIA EXTRAVIADA)
	if r_pedido.loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
		if (Trim(s_pedido_origem) <> "002") And (Trim(s_pedido_origem) <> "019") then
			if s_pedido_ac = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Informe o n� Magento"
				end if
			end if
		end if

    dim c_loja
	c_loja = Trim(Request.Form("c_loja"))

	if r_pedido.loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
		if s_pedido_ac <> "" then
			do while Len(s_pedido_ac) < 9
				if Len(s_pedido_ac) = 8 then
					s_pedido_ac = "1" & s_pedido_ac
				else
					s_pedido_ac = "0" & s_pedido_ac
					end if
				Loop
			
			if Left(s_pedido_ac, 1) <> "1" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O n�mero do pedido Magento inicia com d�gito inv�lido para a loja " & r_pedido.loja
				end if
			end if
		end if
		
	if isLojaBonshop(r_pedido.loja) And (r_pedido.plataforma_origem_pedido = COD_PLATAFORMA_ORIGEM_PEDIDO__MAGENTO) then
		if s_pedido_ac <> "" then
			do while Len(s_pedido_ac) < 9
				if Len(s_pedido_ac) = 8 then
					s_pedido_ac = "2" & s_pedido_ac
				else
					s_pedido_ac = "0" & s_pedido_ac
					end if
				Loop
			
			if Left(s_pedido_ac, 1) <> "2" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O n�mero do pedido Magento inicia com d�gito inv�lido para a loja " & r_pedido.loja
				end if
			end if
		end if

	if versao_forma_pagamento = "1" then
		s_qtde_parcelas=retorna_so_digitos(request("c_qtde_parcelas"))
		end if
	s_perc_RT = Trim(request("c_perc_RT"))

	dim c_exibir_campo_instalador_instala, s_instalador_instala
	c_exibir_campo_instalador_instala = Trim(Request.Form("c_exibir_campo_instalador_instala"))
	s_instalador_instala = Trim(Request.Form("rb_instalador_instala"))

	dim r_cliente
	set r_cliente = New cl_CLIENTE
	call x_cliente_bd(r_pedido.id_cliente, r_cliente)

	dim eh_cpf
	eh_cpf=(len(r_cliente.cnpj_cpf)=11)

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim endereco__bairro, endereco__endereco, endereco__numero, endereco__complemento, endereco__cidade, endereco__uf, endereco__cep
	dim cliente__ddd_res, cliente__tel_res, cliente__ddd_cel, cliente__tel_cel, cliente__ddd_com, cliente__tel_com, cliente__ramal_com,cliente__ddd_com_2, cliente__tel_com_2, cliente__ramal_com_2
	dim cliente__email, cliente__email_xml , cliente__nome, cliente__ie, cliente__rg, cliente__contribuinte_icms_status, cliente__produtor_rural
	
	if r_pedido.st_memorizacao_completa_enderecos = 1 or r_pedido.st_memorizacao_completa_enderecos = 9 then
		endereco__bairro = Trim(Request.Form("endereco__bairro"))
		endereco__endereco = Trim(Request.Form("endereco__endereco"))
		endereco__numero = Trim(Request.Form("endereco__numero"))
		endereco__complemento = Trim(Request.Form("endereco__complemento"))
		endereco__cidade = Trim(Request.Form("endereco__cidade"))
		endereco__uf = Trim(Request.Form("endereco__uf"))
		endereco__cep = retorna_so_digitos(Trim(Request.Form("endereco__cep"))) 	
		cliente__ddd_res = retorna_so_digitos(Trim(Request.Form("cliente__ddd_res"))) 
		cliente__tel_res = retorna_so_digitos(Trim(Request.Form("cliente__tel_res"))) 		
		cliente__ddd_cel = retorna_so_digitos(Trim(Request.Form("cliente__ddd_cel"))) 
		cliente__tel_cel = retorna_so_digitos(Trim(Request.Form("cliente__tel_cel"))) 		
		cliente__ddd_com = retorna_so_digitos(Trim(Request.Form("cliente__ddd_com"))) 
		cliente__tel_com = retorna_so_digitos(Trim(Request.Form("cliente__tel_com"))) 
		cliente__ramal_com = retorna_so_digitos(Trim(Request.Form("cliente__ramal_com")))		
		cliente__ddd_com_2 = retorna_so_digitos(Trim(Request.Form("cliente__ddd_com_2"))) 
		cliente__tel_com_2 = retorna_so_digitos(Trim(Request.Form("cliente__tel_com_2"))) 
		cliente__ramal_com_2 = retorna_so_digitos(Trim(Request.Form("cliente__ramal_com_2")))		
		cliente__email = Trim(Request.Form("cliente__email"))
		cliente__email_xml = Trim(Request.Form("cliente__email_xml"))
		cliente__nome = Trim(Request.Form("cliente__nome"))
		cliente__rg = Trim(Request.Form("cliente__rg"))
		cliente__produtor_rural = Trim(request("rb_produtor_rural"))
		cliente__contribuinte_icms_status = Trim(request("rb_contribuinte_icms"))
		cliente__ie = Trim(Request.Form("cliente__ie")) 		
	end if



	dim blnEndEntregaEdicaoLiberada, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs,blnEndEtg_obs
	dim EndEtg_email, EndEtg_email_xml, EndEtg_nome, EndEtg_ddd_res, EndEtg_tel_res, EndEtg_ddd_com, EndEtg_tel_com, EndEtg_ramal_com
	dim EndEtg_ddd_cel, EndEtg_tel_cel, EndEtg_ddd_com_2, EndEtg_tel_com_2, EndEtg_ramal_com_2
	dim EndEtg_tipo_pessoa, EndEtg_cnpj_cpf, EndEtg_contribuinte_icms_status, EndEtg_produtor_rural_status
	dim EndEtg_ie, EndEtg_rg
	dim blnEndEtgComDados
    blnEndEtg_obs = false
    blnEndEtgComDados = false
	s = Trim(Request.Form("blnEndEntregaEdicaoLiberada"))
	blnEndEntregaEdicaoLiberada = CBool(s)
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

	dim blnTransportadoraEdicaoLiberada, c_transportadora_id, c_transportadora_id_original, c_transportadora_num_coleta, c_transportadora_contato
	s = Trim(Request.Form("blnTransportadoraEdicaoLiberada"))
	blnTransportadoraEdicaoLiberada = CBool(s)
	c_transportadora_id = Trim(Request.Form("c_transportadora_id"))
	c_transportadora_id_original = Trim(Request.Form("c_transportadora_id_original"))
	c_transportadora_num_coleta = Trim(Request.Form("c_transportadora_num_coleta"))
	c_transportadora_contato = Trim(Request.Form("c_transportadora_contato"))
	
    dim blnNumPedidoECommerceEdicaoLiberada
    s = Trim(Request.Form("blnNumPedidoECommerceEdicaoLiberada"))
    blnNumPedidoECommerceEdicaoLiberada = CBool(s)

	dim blnValorFreteEdicaoLiberada, c_valor_frete, frete_transportadora_id, frete_tipo, frete_id, ckb_frete_exclui, c_frete_serie_NF, c_frete_numero_NF, c_frete_emitente
	s = Trim(Request.Form("blnValorFreteEdicaoLiberada"))
	blnValorFreteEdicaoLiberada = CBool(s)
	
	dim rb_PedidoRecebidoStatus, c_PedidoRecebidoData
	rb_PedidoRecebidoStatus = Trim(Request.Form("rb_PedidoRecebidoStatus"))
	c_PedidoRecebidoData = Trim(Request.Form("c_PedidoRecebidoData"))

	dim rb_garantia_indicador, GarantiaIndicadorStatusOriginal
	dim blnGarantiaIndicadorEdicaoLiberada
	GarantiaIndicadorStatusOriginal = Trim(Request.Form("GarantiaIndicadorStatusOriginal"))
	rb_garantia_indicador = Trim(Request.Form("rb_garantia_indicador"))
	s = Trim(Request.Form("blnGarantiaIndicadorEdicaoLiberada"))
	blnGarantiaIndicadorEdicaoLiberada = CBool(s)
        
    if c_loja <> NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
        if s_indicador = "" then
            if s_perc_RT <> "" then
                if converte_numero(s_perc_RT) > 0 then
                    alerta = "N�o � poss�vel gravar o pedido com o campo ""Indicador"" vazio e ""COM(%)"" maior do que zero."
                end if
            end if
        end if
    end if
	
	if blnGarantiaIndicadorEdicaoLiberada then
		if alerta = "" then
			if rb_garantia_indicador = "" then
				if GarantiaIndicadorStatusOriginal = "" then
					alerta = "Falha ao obter o campo 'Garantia Indicador'"
				else
				'	Lembrando que pedidos antigos est�o com o status COD_GARANTIA_INDICADOR_STATUS__NAO_DEFINIDO
				'	e que os radio buttons de edi��o ficam ambos desmarcados
					rb_garantia_indicador = GarantiaIndicadorStatusOriginal
					end if
				end if
			end if
		end if

	dim v_item, i, j, n, intQtdeFretes
	dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_totalFamiliaPrecoNFLiquido, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, id_pedido_base
	redim v_item(0)
	set v_item(Ubound(v_item)) = New cl_ITEM_PEDIDO
	v_item(Ubound(v_item)).produto = ""
	n = Request.Form("c_produto").Count
	for i = 1 to n
		s=Trim(Request.Form("c_produto")(i))
		if s <> "" then
			if Trim(v_item(ubound(v_item)).produto) <> "" then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ITEM_PEDIDO
				end if
			with v_item(ubound(v_item))
				.produto=Ucase(Trim(Request.Form("c_produto")(i)))
				s=retorna_so_digitos(Request.Form("c_fabricante")(i))
				.fabricante=normaliza_codigo(s, TAM_MIN_FABRICANTE)
				s=Trim(Request.Form("c_vl_unitario")(i))
				.preco_venda=converte_numero(s)
				s=Trim(Request.Form("c_vl_NF")(i))
				.preco_NF=converte_numero(s)
				end with
			end if
		next

	dim blnAtivarFlag_st_violado_permite_RA_status
	blnAtivarFlag_st_violado_permite_RA_status = False
	if alerta = "" then
		if (r_pedido.permite_RA_status = 0) And (r_pedido.st_violado_permite_RA_status = 0) then
			for i=Lbound(v_item) to Ubound(v_item)
				if Trim("" & v_item(i).produto) <> "" then
					if v_item(i).preco_NF <> v_item(i).preco_venda then
						blnAtivarFlag_st_violado_permite_RA_status = True
						exit for
						end if
					end if
				next
			end if
		end if
	
	dim blnNFEmitida
	blnNFEmitida = False
    if alerta = "" then
	    if Trim("" & r_pedido.obs_2) <> "" then blnNFEmitida = True
    end if
	
	dim idxIndiceBd
	dim v_nfe_item_devolvido, v_item_devolvido_bd
	redim v_nfe_item_devolvido(0)
	set v_nfe_item_devolvido(0) = New cl_PEDIDO_ITEM_DEVOLVIDO
	if blnDadosNFeMercadoriasDevolvidasEdicaoLiberada then
		n = Request.Form("c_item_devolvido_id").Count
		for i = 1 to n
			s=Trim(Request.Form("c_item_devolvido_id")(i))
			if s <> "" then
				if Trim(v_nfe_item_devolvido(Ubound(v_nfe_item_devolvido)).id) <> "" then
					redim preserve v_nfe_item_devolvido(Ubound(v_nfe_item_devolvido)+1)
					set v_nfe_item_devolvido(Ubound(v_nfe_item_devolvido)) = New cl_PEDIDO_ITEM_DEVOLVIDO
					end if
				with v_nfe_item_devolvido(Ubound(v_nfe_item_devolvido))
					.id = Trim(Request.Form("c_item_devolvido_id")(i))
					.fabricante = Trim(Request.Form("c_item_devolvido_fabricante")(i))
					.produto = Trim(Request.Form("c_item_devolvido_produto")(i))
					.id_nfe_emitente = Trim(Request.Form("c_item_devolvido_nfe_emitente")(i))
					.NFe_serie_NF = Trim(Request.Form("c_item_devolvido_nfe_serie")(i))
					.NFe_numero_NF = Trim(Request.Form("c_item_devolvido_nfe_numero")(i))
					end with
				end if
			next
		
		if alerta = "" then
			for i=Lbound(v_nfe_item_devolvido) to Ubound(v_nfe_item_devolvido)
				with v_nfe_item_devolvido(i)
					if Trim(.id) <> "" then
						if Trim(.fabricante)="" then
							alerta=texto_add_br(alerta)
							alerta=alerta & "A devolu��o de mercadorias id=" & Trim(.id) & " n�o informa o c�digo do fabricante."
							end if
						if Trim(.produto)="" then
							alerta=texto_add_br(alerta)
							alerta=alerta & "A devolu��o de mercadorias id=" & Trim(.id) & " n�o informa o c�digo do produto."
							end if
						
						blnFlag=False
						if (Trim(.id_nfe_emitente)<>"") And (Trim(.id_nfe_emitente)<>"0") then blnFlag=True
						if (Trim(.NFe_serie_NF)<>"") And (Trim(.NFe_serie_NF)<>"0") then blnFlag=True
						if (Trim(.NFe_numero_NF)<>"") And (Trim(.NFe_numero_NF)<>"0") then blnFlag=True
						
						if blnFlag then
							if converte_numero(.id_nfe_emitente) = 0 then
								alerta=texto_add_br(alerta)
								alerta=alerta & "� necess�rio selecionar um emitente de NFe v�lido na devolu��o do produto " & Trim(.produto) & "!!"
								end if
							if converte_numero(.NFe_serie_NF) = 0 then
								alerta=texto_add_br(alerta)
								alerta=alerta & "� necess�rio informar a s�rie da NFe na devolu��o do produto " & Trim(.produto) & "!!"
								end if
							if converte_numero(.NFe_numero_NF) = 0 then
								alerta=texto_add_br(alerta)
								alerta=alerta & "� necess�rio informar o n�mero da NFe na devolu��o do produto " & Trim(.produto) & "!!"
								end if
							end if
						end if
					end with
				next
			end if
			
		if alerta = "" then
			if Not le_pedido_item_devolvido(pedido_selecionado, v_item_devolvido_bd, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			for i=Lbound(v_nfe_item_devolvido) to Ubound(v_nfe_item_devolvido)
				blnFlag=False
				idxIndiceBd = -1
				if Trim(v_nfe_item_devolvido(i).id) <> "" then
				'	LOCALIZA REGISTRO DE DEVOLU��O E VERIFICA SE HOUVE ALTERA��O NOS DADOS DA NFe
					for j=Lbound(v_item_devolvido_bd) to Ubound(v_item_devolvido_bd)
						if Trim(v_nfe_item_devolvido(i).id) = Trim(v_item_devolvido_bd(j).id) then
							blnFlag=True
							idxIndiceBd = j
							exit for
							end if
						next
						
					if Not blnFlag then
						alerta = "N�o foi localizado o registro referente � devolu��o de mercadorias id=" & Trim(v_nfe_item_devolvido(i).id)
						exit for
						end if
					
					if alerta = "" then
					'	SE OS DADOS DA NFe FORAM EDITADOS, CONSISTE SE A NFe � DO CLIENTE CORRETO E SE DECLARA O REFERIDO PRODUTO
						blnFlag=False
						if converte_numero(v_nfe_item_devolvido(i).id_nfe_emitente)<>converte_numero(v_item_devolvido_bd(idxIndiceBd).id_nfe_emitente) then blnFlag=True
						if converte_numero(v_nfe_item_devolvido(i).NFe_serie_NF)<>converte_numero(v_item_devolvido_bd(idxIndiceBd).NFe_serie_NF) then blnFlag=True
						if converte_numero(v_nfe_item_devolvido(i).NFe_numero_NF)<>converte_numero(v_item_devolvido_bd(idxIndiceBd).NFe_numero_NF) then blnFlag=True
						if blnFlag then
						'	USA UM DOS CAMPOS N�O UTILIZADOS P/ SINALIZAR QUE O ITEM FOI EDITADO
							v_nfe_item_devolvido(i).descricao = "**-#-EDITADO-#-**"
                            if v_nfe_item_devolvido(i).id_nfe_emitente <> "" And v_nfe_item_devolvido(i).NFe_serie_NF <> "" And v_nfe_item_devolvido(i).NFe_numero_NF <> "" then

							    if CLng(v_nfe_item_devolvido(i).id_nfe_emitente) <> CLng(COD_NFE_EMITENTE__CLIENTE) then
								    s = "SELECT" & _
										    " id," & _
										    " ide__tpNF," & _
										    " dest__CNPJ," & _
										    " dest__CPF" & _
									    " FROM t_NFe_IMAGEM" & _
									    " WHERE" & _
										    " (id_nfe_emitente = " & v_nfe_item_devolvido(i).id_nfe_emitente & ")" & _
										    " AND (NFe_serie_NF = " & v_nfe_item_devolvido(i).NFe_serie_NF & ")" & _
										    " AND (NFe_numero_NF = " & v_nfe_item_devolvido(i).NFe_numero_NF & ")" & _
										    " AND (codigo_retorno_NFe_T1 = '1')" & _
										    " AND (st_anulado = 0)" & _
									    " ORDER BY" & _
										    " id DESC"
								    set rs = cn.Execute(s)
								    if rs.Eof then
									    alerta = "N�o consta no banco de dados a NFe com n� " & Cstr(v_nfe_item_devolvido(i).NFe_numero_NF) & " da s�rie n� " & Cstr(v_nfe_item_devolvido(i).NFe_serie_NF)
									    exit for
								    else
									    if Trim("" & rs("ide__tpNF")) <> "0" then
										    alerta = "A NFe n� " & Cstr(v_nfe_item_devolvido(i).NFe_numero_NF) & " da s�rie n� " & Cstr(v_nfe_item_devolvido(i).NFe_serie_NF) & " N�O � uma nota de entrada."
										    exit for
										    end if
									    if Len(retorna_so_digitos(r_cliente.cnpj_cpf)) = 14 then
										    if retorna_so_digitos(r_cliente.cnpj_cpf)<>retorna_so_digitos(Trim("" & rs("dest__CNPJ"))) then
											    alerta = "A NFe n� " & Cstr(v_nfe_item_devolvido(i).NFe_numero_NF) & " da s�rie n� " & Cstr(v_nfe_item_devolvido(i).NFe_serie_NF) & " N�O foi emitida para o cliente deste pedido."
											    exit for
											    end if
									    else
										    if retorna_so_digitos(r_cliente.cnpj_cpf)<>retorna_so_digitos(Trim("" & rs("dest__CPF"))) then
											    alerta = "A NFe n� " & Cstr(v_nfe_item_devolvido(i).NFe_numero_NF) & " da s�rie n� " & Cstr(v_nfe_item_devolvido(i).NFe_serie_NF) & " N�O foi emitida para o cliente deste pedido."
											    exit for
											    end if
										    end if

								    '	VERIFICA PRODUTO
									    s = "SELECT" & _
											    " id," & _
											    " fabricante," & _
											    " produto," & _
											    " det__qCom" & _
										    " FROM t_NFe_IMAGEM_ITEM" & _
										    " WHERE" & _
											    " (id_nfe_imagem = " & Trim("" & rs("id")) & ")" & _
											    " AND (fabricante = '" & v_nfe_item_devolvido(i).fabricante & "')" & _
											    " AND (produto = '" & v_nfe_item_devolvido(i).produto & "')"
									    set rs2 = cn.Execute(s)
									    if rs2.Eof then
										    alerta = "O produto " & v_nfe_item_devolvido(i).produto & " N�O consta na NFe n� " & Cstr(v_nfe_item_devolvido(i).NFe_numero_NF) & " da s�rie n� " & Cstr(v_nfe_item_devolvido(i).NFe_serie_NF)
										    exit for
									    else
										    if converte_numero(v_item_devolvido_bd(idxIndiceBd).qtde) > converte_numero(Trim("" & rs2("det__qCom"))) then
											    alerta = "A NFe n� " & Cstr(v_nfe_item_devolvido(i).NFe_numero_NF) & " da s�rie n� " & Cstr(v_nfe_item_devolvido(i).NFe_serie_NF) & " declara apenas " & converte_numero(Trim("" & rs2("det__qCom"))) & " unidade(s) do produto " & v_item_devolvido_bd(idxIndiceBd).produto & " ao inv�s de " & v_item_devolvido_bd(idxIndiceBd).qtde
											    exit for
											    end if
										    end if
									    end if  'if rs.Eof
								    end if  'if CLng(v_nfe_item_devolvido(i).id_nfe_emitente) <> CLng(COD_NFE_EMITENTE__CLIENTE)
                                end if 'if v_nfe_item_devolvido(i).id_nfe_emitente <> "" And v_nfe_item_devolvido(i).NFe_serie_NF <> "" And v_nfe_item_devolvido(i).NFe_numero_NF <> ""'
							end if  'if blnFlag
						end if
					end if
				next
			end if
		end if  'if blnDadosNFeMercadoriasDevolvidasEdicaoLiberada
		
'	FORMA DE PAGAMENTO (NOVA VERS�O)
	if alerta = "" then
		if (versao_forma_pagamento = "2") And blnFormaPagtoEdicaoLiberada then
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
		end if

	dim c_custoFinancFornecTipoParcelamento, c_custoFinancFornecQtdeParcelas
	dim c_custoFinancFornecTipoParcelamentoOriginal, c_custoFinancFornecQtdeParcelasOriginal
	dim c_custoFinancFornecTipoParcelamentoConferencia, c_custoFinancFornecQtdeParcelasConferencia
	dim coeficiente, vlCustoFinancFornecPrecoLista, vlCustoFinancFornecPrecoListaBase
	c_custoFinancFornecTipoParcelamentoOriginal = Trim(Request.Form("c_custoFinancFornecTipoParcelamentoOriginal"))
	c_custoFinancFornecQtdeParcelasOriginal = Trim(Request.Form("c_custoFinancFornecQtdeParcelasOriginal"))
	c_custoFinancFornecTipoParcelamento = Trim(Request.Form("c_custoFinancFornecTipoParcelamento"))
	c_custoFinancFornecQtdeParcelas = Trim(Request.Form("c_custoFinancFornecQtdeParcelas"))
	
'	O PEDIDO FOI CADASTRADO J� DENTRO DA POL�TICA DE PERCENTUAL DE CUSTO FINANCEIRO POR FORNECEDOR?
	if versao_forma_pagamento = "2" then
		if (c_custoFinancFornecTipoParcelamentoOriginal <> "") And blnFormaPagtoEdicaoLiberada then
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
			end if
		end if

	'Edit�vel?
	if blnEndEntregaEdicaoLiberada then
		if alerta = "" then
            if (EndEtg_endereco<>r_pedido.EndEtg_endereco) Or (EndEtg_bairro<>r_pedido.EndEtg_bairro) Or (EndEtg_cidade<>r_pedido.EndEtg_cidade) Or (EndEtg_uf<>r_pedido.EndEtg_uf) Or (EndEtg_cep<>r_pedido.EndEtg_cep) Or (EndEtg_obs<>r_pedido.EndEtg_obs) Or (EndEtg_obs<>r_pedido.EndEtg_cod_justificativa) then
                  blnEndEtg_obs = true 
                end if

            'na memorizacao de endere�os ligada, sempre verificamos
            if r_pedido.st_memorizacao_completa_enderecos <> 0 and blnUsarMemorizacaoCompletaEnderecos then
                blnEndEtg_obs = true 
                end if

            blnEndEtgComDados = false
			if (EndEtg_endereco<>"") Or (EndEtg_endereco_numero<>"") Or (EndEtg_endereco_complemento<>"") Or (EndEtg_bairro<>"") Or (EndEtg_cidade<>"") Or (EndEtg_uf<>"") Or (EndEtg_cep<>"") Or (EndEtg_obs<>"") then
                blnEndEtgComDados = true
                end if
            if r_pedido.st_memorizacao_completa_enderecos <> 0 and blnUsarMemorizacaoCompletaEnderecos then

				if (EndEtg_email<>"") Or (EndEtg_email_xml<>"") then
					blnEndEtgComDados = true
					end if

                if not eh_cpf then
			        if (EndEtg_ddd_res<>"") Or (EndEtg_tel_res<>"") Or (EndEtg_ddd_com<>"") Or (EndEtg_tel_com<>"") Or (EndEtg_ramal_com<>"") then
                        blnEndEtgComDados = true
                        end if
			        if (EndEtg_ddd_cel<>"") Or (EndEtg_tel_cel<>"") Or (EndEtg_ddd_com_2<>"") Or (EndEtg_tel_com_2<>"") Or (EndEtg_ramal_com_2<>"") Or (EndEtg_tipo_pessoa<>"") then
                        blnEndEtgComDados = true
                        end if
			        if (EndEtg_cnpj_cpf<>"") Or (EndEtg_contribuinte_icms_status<>"") Or (EndEtg_produtor_rural_status<>"") Or (EndEtg_ie<>"") Or (EndEtg_rg<>"") then
                        blnEndEtgComDados = true
                        end if

                    end if

                if eh_cpf and not blnEndEtgComDados then
                    'nenhum campo deve ser preenchido pelo usu�rio
                    'todos possuem prenchimento autom�tico
                    EndEtg_ddd_res = ""
                    EndEtg_tel_res = ""
                    EndEtg_ddd_cel = ""
                    EndEtg_tel_cel = ""
                    EndEtg_ddd_com = ""
                    EndEtg_tel_com = ""
                    EndEtg_ramal_com = ""
                    EndEtg_ddd_com_2 = ""
                    EndEtg_tel_com_2 = ""
                    EndEtg_ramal_com_2 = ""
                    EndEtg_tipo_pessoa = ""
                    EndEtg_cnpj_cpf = ""
                    EndEtg_ie = ""
                    EndEtg_contribuinte_icms_status = ""
                    EndEtg_rg = ""
                    EndEtg_produtor_rural_status = ""
                    EndEtg_email = ""
                    EndEtg_email_xml = ""
                    EndEtg_nome = ""
                    end if

				'limpeza de campos EndEtg
				if blnEndEtgComDados and EndEtg_tipo_pessoa = "PJ" then
					EndEtg_produtor_rural_status = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_INICIAL)
					end if
				if blnEndEtgComDados and EndEtg_tipo_pessoa <> "PJ" then
					if converte_numero(EndEtg_produtor_rural_status) = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
						EndEtg_contribuinte_icms_status = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_INICIAL
						EndEtg_ie = ""
						end if
					end if

                end if


			if blnEndEtgComDados then
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
                 elseif (EndEtg_obs="" AND blnEndEtg_obs= true) then
                    alerta="PREENCHA A JUSTIFICATIVA DO ENDERE�O DE ENTREGA."
                    end if
				end if
			end if

			if r_pedido.st_memorizacao_completa_enderecos = 1 or r_pedido.st_memorizacao_completa_enderecos = 9 then
			
				if endereco__endereco="" then
					alerta="PREENCHA O ENDERE�O."
				elseif Len(endereco__endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
					alerta="ENDERE�O EXCEDE O TAMANHO M�XIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(endereco__endereco)) & " CARACTERES<br>TAMANHO M�XIMO: " & Cstr(endereco__endereco) & " CARACTERES"
				elseif endereco__numero="" then
					alerta="PREENCHA O N�MERO DO ENDERE�O."
				elseif endereco__cidade="" then
					alerta="PREENCHA A CIDADE DO ENDERE�O."
				elseif endereco__uf="" then
					alerta="PREENCHA A UF DO ENDERE�O."
				elseif endereco__cep="" then
					alerta="PREENCHA O CEP DO ENDERE�O."	
		        elseif Not cep_ok(endereco__cep) then
			        alerta="CEP INV�LIDO."
		        elseif Not ddd_ok(cliente__ddd_res) then
			        alerta="DADOS CADASTRAIS: DDD INV�LIDO."
		        elseif Not telefone_ok(cliente__tel_res) then
			        alerta="DADOS CADASTRAIS: TELEFONE RESIDENCIAL INV�LIDO."
		        elseif (cliente__ddd_res <> "") And ((cliente__tel_res = "")) then
			        alerta="DADOS CADASTRAIS: PREENCHA O TELEFONE RESIDENCIAL."
		        elseif (cliente__ddd_res = "") And ((cliente__tel_res <> "")) then
			        alerta="DADOS CADASTRAIS: PREENCHA O DDD."
		        elseif Not ddd_ok(cliente__ddd_com) then
			        alerta="DADOS CADASTRAIS: DDD INV�LIDO."
		        elseif Not telefone_ok(cliente__tel_com) then
			        alerta="DADOS CADASTRAIS: TELEFONE COMERCIAL INV�LIDO."
		        elseif (cliente__ddd_com <> "") And ((cliente__tel_com = "")) then
			        alerta="DADOS CADASTRAIS: PREENCHA O TELEFONE COMERCIAL."
		        elseif (cliente__ddd_com = "") And ((cliente__tel_com <> "")) then
			        alerta="DADOS CADASTRAIS: PREENCHA O DDD."
		        elseif eh_cpf And (cliente__tel_res="") And (cliente__tel_com="") And (cliente__tel_cel="") then
			        alerta="DADOS CADASTRAIS: PREENCHA PELO MENOS UM TELEFONE."
		        elseif (Not eh_cpf) And (cliente__tel_com="") And (cliente__tel_com_2="") then
			        alerta="DADOS CADASTRAIS: PREENCHA O TELEFONE."
				end if

				if  eh_cpf then
                    if cliente__produtor_rural = "" then
                        alerta = "Dados cadastrais: informe se o cliente � produtor rural ou n�o!!"
                    elseif converte_numero(cliente__produtor_rural) = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
                        cliente__contribuinte_icms_status = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_INICIAL
                        cliente__ie = ""
                    elseif converte_numero(cliente__produtor_rural) <> converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
                        if converte_numero(cliente__contribuinte_icms_status) <> converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                            alerta = "Dados cadastrais: para ser cadastrado como Produtor Rural, � necess�rio ser contribuinte do ICMS e possuir n� de IE!!"
                        elseif cliente__contribuinte_icms_status = "" then
                            alerta = "Dados cadastrais: informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!"
                        elseif converte_numero(cliente__contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and cliente__ie = "" then
                            alerta = "Dados cadastrais: se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!"
                        elseif converte_numero(cliente__contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) and InStr(cliente__ie, "ISEN") > 0 then 
                            alerta = "Dados cadastrais: se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!"
                        elseif converte_numero(cliente__contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and InStr(cliente__ie, "ISEN") > 0 then 
                            alerta = "Dados cadastrais: se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!"
                        elseif converte_numero(cliente__contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) and cliente__ie <> "" then 
                            alerta = "Dados cadastrais: se o Contribuinte ICMS � isento, o campo IE deve ser vazio!"
                            end if
                        end if
					end if
			
				if	cliente__ie <> "" then 
					if Not isInscricaoEstadualValida(cliente__ie, endereco__uf) then
						alerta="Preencha a IE (Inscri��o Estadual) com um n�mero v�lido!!" & _
							"<br>" & "Certifique-se de que a UF do endere�o corresponde � UF respons�vel pelo registro da IE."
						end if
				end if
			end if


        if alerta = "" and blnEndEtgComDados and r_pedido.st_memorizacao_completa_enderecos <> 0 and blnUsarMemorizacaoCompletaEnderecos and Not eh_cpf then
            if EndEtg_tipo_pessoa <> "PJ" and EndEtg_tipo_pessoa <> "PF" then
                alerta = "Necess�rio escolher Pessoa Jur�dica ou Pessoa F�sica no Endere�o de entrega!!"
    		elseif EndEtg_nome = "" then
                alerta = "Preencha o nome/raz�o social no endere�o de entrega!!"
                end if 
	
            if alerta = "" and EndEtg_tipo_pessoa = "PJ" then

                'limpa os n�meros de telefone que n�o foram informados
                EndEtg_ddd_res = ""
                EndEtg_tel_res = ""
                EndEtg_ddd_cel = ""
                EndEtg_tel_cel = ""

                '//Campos PJ: 
                if EndEtg_cnpj_cpf = "" or not cnpj_ok(EndEtg_cnpj_cpf) then
                    alerta = "Endere�o de entrega: CNPJ inv�lido!!"
                elseif EndEtg_contribuinte_icms_status = "" then
                    alerta = "Endere�o de entrega: selecione o tipo de contribuinte de ICMS!!"
                elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and EndEtg_ie = "" then
                    alerta = "Endere�o de entrega: se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!"
                elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) and InStr(EndEtg_ie, "ISEN") > 0 then 
                    alerta = "Endere�o de entrega: se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!"
                elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and InStr(EndEtg_ie, "ISEN") > 0 then 
                    alerta = "Endere�o de entrega: se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!"
                'telefones PJ:
                'EndEtg_ddd_com
                'EndEtg_tel_com
                'EndEtg_ramal_com
                'EndEtg_ddd_com_2
                'EndEtg_tel_com_2
                'EndEtg_ramal_com_2
                elseif not ddd_ok(EndEtg_ddd_com) then
                    alerta = "Endere�o de entrega: DDD inv�lido!!"
                elseif not telefone_ok(EndEtg_tel_com) then
                    alerta = "Endere�o de entrega: telefone inv�lido!!"
                elseif EndEtg_ddd_com = "" and EndEtg_tel_com <> "" then
                    alerta = "Endere�o de entrega: preencha o DDD do telefone."
                elseif EndEtg_tel_com = "" and EndEtg_ddd_com <> "" then
                    alerta = "Endere�o de entrega: preencha o telefone."

                elseif not ddd_ok(EndEtg_ddd_com_2) then
                    alerta = "Endere�o de entrega: DDD inv�lido!!"
                elseif not telefone_ok(EndEtg_tel_com_2) then
                    alerta = "Endere�o de entrega: telefone inv�lido!!"
                elseif EndEtg_ddd_com_2 = "" and EndEtg_tel_com_2 <> "" then
                    alerta = "Endere�o de entrega: preencha o DDD do telefone."
                elseif EndEtg_tel_com_2 = "" and EndEtg_ddd_com_2 <> "" then
                    alerta = "Endere�o de entrega: preencha o telefone."
                    end if 
                end if 

            if alerta = "" and EndEtg_tipo_pessoa <> "PJ" then

                'limpa os n�meros de telefone que n�o foram informados
                EndEtg_ddd_com = ""
                EndEtg_tel_com = ""
                EndEtg_ramal_com = ""
                EndEtg_ddd_com_2 = ""
                EndEtg_tel_com_2 = ""
                EndEtg_ramal_com_2 = ""

                '//campos PF
                if EndEtg_cnpj_cpf = "" or not cpf_ok(EndEtg_cnpj_cpf) then
                    alerta = "Endere�o de entrega: CPF inv�lido!!"
                elseif EndEtg_produtor_rural_status = "" then
                    alerta = "Endere�o de entrega: informe se o cliente � produtor rural ou n�o!!"
                elseif converte_numero(EndEtg_produtor_rural_status) <> converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
                    if converte_numero(EndEtg_contribuinte_icms_status) <> converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                        alerta = "Endere�o de entrega: para ser cadastrado como Produtor Rural, � necess�rio ser contribuinte do ICMS e possuir n� de IE!!"
                    elseif EndEtg_contribuinte_icms_status = "" then
                        alerta = "Endere�o de entrega: informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and EndEtg_ie = "" then
                        alerta = "Endere�o de entrega: se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) and InStr(EndEtg_ie, "ISEN") > 0 then 
                        alerta = "Endere�o de entrega: se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and InStr(EndEtg_ie, "ISEN") > 0 then 
                        alerta = "Endere�o de entrega: se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) and EndEtg_ie <> "" then 
                        alerta = "Endere�o de entrega: se o Contribuinte ICMS � isento, o campo IE deve ser vazio!"
                        end if
                    end if

                if alerta = "" then
                    'telefones PF:
                    'EndEtg_ddd_res
                    'EndEtg_tel_res
                    'EndEtg_ddd_cel
                    'EndEtg_tel_cel
                    if not ddd_ok(retorna_so_digitos(EndEtg_ddd_res)) then
                        alerta = "Endere�o de entrega: DDD inv�lido!!"
                    elseif not telefone_ok(retorna_so_digitos(EndEtg_tel_res)) then
                        alerta = "Endere�o de entrega: telefone inv�lido!!"
                    elseif EndEtg_ddd_res <> "" or EndEtg_tel_res <> "" then
                        if EndEtg_ddd_res = "" then
                            alerta = "Endere�o de entrega: preencha o DDD!!"
                        elseif EndEtg_tel_res = "" then
                            alerta = "Endere�o de entrega: preencha o telefone!!"
                            end if
                        end if
                    end if

                if alerta = "" then
                    if not ddd_ok(retorna_so_digitos(EndEtg_ddd_cel)) then
                        alerta = "Endere�o de entrega: DDD inv�lido!!"
                    elseif not telefone_ok(retorna_so_digitos(EndEtg_tel_cel)) then
                        alerta = "Endere�o de entrega: telefone inv�lido!!"
                    elseif EndEtg_ddd_cel = "" and EndEtg_tel_cel <> "" then
                        alerta = "Endere�o de entrega: preencha o DDD do celular."
                    elseif EndEtg_tel_cel = "" and EndEtg_ddd_cel <> "" then
                        alerta = "Endere�o de entrega: preencha o n�mero do celular."
                        end if
                    end if

                end if

		    if alerta = "" and EndEtg_ie <> "" then
			    if Not isInscricaoEstadualValida(EndEtg_ie, EndEtg_uf) then
				    alerta="Endere�o de entrega: preencha a IE (Inscri��o Estadual) com um n�mero v�lido!!" & _
						    "<br>" & "Certifique-se de que a UF do endere�o de entrega corresponde � UF respons�vel pelo registro da IE."
				    end if
			    end if

            end if
		end if

	'Verifica se est� havendo edi��o no cadastro de cliente que possui pedido com status de an�lise de cr�dito 'cr�dito ok' e com entrega pendente
    function log_endereco_um_vetor(byref v1, nome, valor)
		redim preserve v1(ubound(v1)+1)
        set v1(ubound(v1)) = new cl_LOG_VIA_VETOR
    	v1(ubound(v1)).nome = nome
    	v1(ubound(v1)).valor = valor
    end function

    function log_endereco(byref v1, byref v2, nome, c1, c2)
		log_endereco_um_vetor v1, nome, c1
		log_endereco_um_vetor v2, nome, c2
    end function

	dim blnHaPedidoAprovadoComEntregaPendente				
	dim sLogEmail
    dim sLogVetor1
    dim sLogVetor2
	blnHaPedidoAprovadoComEntregaPendente = False
    sLogEmail = ""
	if alerta = "" then
		if r_pedido.st_entrega <> "ETG" and r_pedido.st_entrega <> "CAN" and CLng(r_pedido.analise_credito) = CLng(COD_AN_CREDITO_OK) then
			if r_pedido.endereco_logradouro  <> endereco__endereco or r_pedido.endereco_bairro  <> endereco__bairro or r_pedido.endereco_numero  <> endereco__numero or r_pedido.endereco_complemento  <> endereco__complemento or r_pedido.endereco_cidade  <> endereco__cidade or r_pedido.endereco_uf  <> endereco__uf or r_pedido.endereco_cep  <> endereco__cep then 
				blnHaPedidoAprovadoComEntregaPendente = true							
	            redim sLogVetor1(0)
	            set sLogVetor1(0) = new cl_LOG_VIA_VETOR
	            redim sLogVetor2(0)
	            set sLogVetor2(0) = new cl_LOG_VIA_VETOR

                log_endereco sLogVetor1, sLogVetor2, "Endere�o", r_pedido.endereco_logradouro, endereco__endereco
                log_endereco sLogVetor1, sLogVetor2, "Bairro", r_pedido.endereco_bairro, endereco__bairro
                log_endereco sLogVetor1, sLogVetor2, "N�mero", r_pedido.endereco_numero, endereco__numero
                log_endereco sLogVetor1, sLogVetor2, "Complemento", r_pedido.endereco_complemento, endereco__complemento
                log_endereco sLogVetor1, sLogVetor2, "Cidade", r_pedido.endereco_cidade, endereco__cidade
                log_endereco sLogVetor1, sLogVetor2, "UF", r_pedido.endereco_uf, endereco__uf
                log_endereco sLogVetor1, sLogVetor2, "CEP", r_pedido.endereco_cep, endereco__cep

                sLogEmail = sLogEmail & log_via_vetor_monta_alteracao(sLogVetor1, sLogVetor2)
    			sLogEmail = sLogEmail & ";;Endere�o novo: " & endereco__endereco & ", " & endereco__numero & " " & endereco__complemento & " - " & endereco__bairro  & " - " & endereco__cidade & "/" & endereco__uf & " " & cep_formata(endereco__cep)
    			sLogEmail = sLogEmail & ";Endere�o anterior: " & r_pedido.endereco_logradouro & ", " & r_pedido.endereco_numero & " " & r_pedido.endereco_complemento & " - " & r_pedido.endereco_bairro  & " - " & r_pedido.endereco_cidade & "/" & r_pedido.endereco_uf & " " & cep_formata(r_pedido.endereco_cep)

			end if
		end if
	end if

	
'	CONSIST�NCIAS P/ EMISS�O DE NFe
	dim s_tabela_municipios_IBGE
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
		if blnEndEntregaEdicaoLiberada And (EndEtg_cidade <> "") then
		'	MUNIC�PIO DE ACORDO C/ TABELA DO IBGE?
			dim s_lista_sugerida_municipios
			dim v_lista_sugerida_municipios
			dim iCounterLista, iNumeracaoLista
			if Not consiste_municipio_IBGE_ok(EndEtg_cidade, EndEtg_uf, s_lista_sugerida_municipios, msg_erro) then
				if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				if msg_erro <> "" then
					alerta = alerta & msg_erro
				else
					alerta = alerta & "Munic�pio '" & EndEtg_cidade & "' n�o consta na rela��o de munic�pios do IBGE para a UF de '" & EndEtg_uf & "'!!"
					if s_lista_sugerida_municipios <> "" then
						alerta = alerta & "<br>" & _
										  "Localize o munic�pio na lista abaixo e verifique se a grafia est� correta!!"
						v_lista_sugerida_municipios = Split(s_lista_sugerida_municipios, chr(13))
						iNumeracaoLista=0
						for iCounterLista=LBound(v_lista_sugerida_municipios) to UBound(v_lista_sugerida_municipios)
							if Trim("" & v_lista_sugerida_municipios(iCounterLista)) <> "" then
								iNumeracaoLista=iNumeracaoLista+1
								s_tabela_municipios_IBGE = s_tabela_municipios_IBGE & _
													"	<tr>" & chr(13) & _
													"		<td align='right'>" & chr(13) & _
													"			<span class='N'>&nbsp;" & Cstr(iNumeracaoLista) & "." & "</span>" & chr(13) & _
													"		</td>" & chr(13) & _
													"		<td>" & chr(13) & _
													"			<span class='N'>" & Trim("" & v_lista_sugerida_municipios(iCounterLista)) & "</span>" & chr(13) & _
													"		</td>" & chr(13) & _
													"	</tr>" & chr(13)
								end if
							next

						if s_tabela_municipios_IBGE <> "" then
							s_tabela_municipios_IBGE = _
									"<table cellspacing='0' cellpadding='1'>" & chr(13) & _
									"	<tr>" & chr(13) & _
									"		<td align='center'>" & chr(13) & _
									"			<p class='N'>" & "Rela��o de munic�pios de '" & EndEtg_uf & "' que se iniciam com a letra '" & Ucase(left(EndEtg_cidade,1)) & "'" & "</p>" & chr(13) & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"	<tr>" & chr(13) & _
									"		<td align='center'>" & chr(13) &_
									"			<table cellspacing='0' border='1'>" & chr(13) & _
													s_tabela_municipios_IBGE & _
									"			</table>" & chr(13) & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"</table>" & chr(13)
							end if
						end if
					end if
				end if 'if Not consiste_municipio_IBGE_ok()
			end if 'if blnEndEntregaEdicaoLiberada And (EndEtg_cidade <> "")
		end if 'if alerta = ""
	
'	CONSIST�NCIAS P/ EMISS�O DE NFe (DADOS CADASTRAIS)
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
		if r_pedido.st_memorizacao_completa_enderecos <> 0 then
		'	MUNIC�PIO DE ACORDO C/ TABELA DO IBGE?
			if Not consiste_municipio_IBGE_ok(endereco__cidade, endereco__uf, s_lista_sugerida_municipios, msg_erro) then
				if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				if msg_erro <> "" then
					alerta = alerta & msg_erro
				else
					alerta = alerta & "Munic�pio '" & endereco__cidade & "' n�o consta na rela��o de munic�pios do IBGE para a UF de '" & endereco__uf & "'!!"
					if s_lista_sugerida_municipios <> "" then
						alerta = alerta & "<br>" & _
										  "Localize o munic�pio na lista abaixo e verifique se a grafia est� correta!!"
						v_lista_sugerida_municipios = Split(s_lista_sugerida_municipios, chr(13))
						iNumeracaoLista=0
						for iCounterLista=LBound(v_lista_sugerida_municipios) to UBound(v_lista_sugerida_municipios)
							if Trim("" & v_lista_sugerida_municipios(iCounterLista)) <> "" then
								iNumeracaoLista=iNumeracaoLista+1
								s_tabela_municipios_IBGE = s_tabela_municipios_IBGE & _
													"	<tr>" & chr(13) & _
													"		<td align='right'>" & chr(13) & _
													"			<span class='N'>&nbsp;" & Cstr(iNumeracaoLista) & "." & "</span>" & chr(13) & _
													"		</td>" & chr(13) & _
													"		<td>" & chr(13) & _
													"			<span class='N'>" & Trim("" & v_lista_sugerida_municipios(iCounterLista)) & "</span>" & chr(13) & _
													"		</td>" & chr(13) & _
													"	</tr>" & chr(13)
								end if
							next

						if s_tabela_municipios_IBGE <> "" then
							s_tabela_municipios_IBGE = _
									"<table cellspacing='0' cellpadding='1'>" & chr(13) & _
									"	<tr>" & chr(13) & _
									"		<td align='center'>" & chr(13) & _
									"			<p class='N'>" & "Rela��o de munic�pios de '" & endereco__uf & "' que se iniciam com a letra '" & Ucase(left(endereco__cidade,1)) & "'" & "</p>" & chr(13) & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"	<tr>" & chr(13) & _
									"		<td align='center'>" & chr(13) &_
									"			<table cellspacing='0' border='1'>" & chr(13) & _
													s_tabela_municipios_IBGE & _
									"			</table>" & chr(13) & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"</table>" & chr(13)
							end if
						end if
					end if
				end if 'if Not consiste_municipio_IBGE_ok()
			end if 'if r_pedido.st_memorizacao_completa_enderecos <> 0 
		end if 'if alerta = ""

	dim s_caracteres_invalidos
	if alerta = "" then
		if Not isTextoValido(EndEtg_endereco, s_caracteres_invalidos) then
			alerta="O CAMPO 'ENDERE�O DE ENTREGA' POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_endereco_numero, s_caracteres_invalidos) then
			alerta="O CAMPO N�MERO DO ENDERE�O DE ENTREGA POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_endereco_complemento, s_caracteres_invalidos) then
			alerta="O CAMPO COMPLEMENTO DO ENDERE�O DE ENTREGA POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_bairro, s_caracteres_invalidos) then
			alerta="O CAMPO BAIRRO DO ENDERE�O DE ENTREGA POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_cidade, s_caracteres_invalidos) then
			alerta="O CAMPO CIDADE DO ENDERE�O DE ENTREGA POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_nome, s_caracteres_invalidos) then
			alerta="O CAMPO NOME DO ENDERE�O DE ENTREGA POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
			end if
		end if

	if alerta = "" then
		if blnEntregaImediataEdicaoLiberada then
			if CLng(s_etg_imediata) = CLng(COD_ETG_IMEDIATA_NAO) then
				if Not blnEntregaImediataNaoSemDataPrevisao then
					if c_data_previsao_entrega = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "� necess�rio informar a data de previs�o de entrega"
						end if
					end if

				if c_data_previsao_entrega <> "" then
					if Not IsDate(c_data_previsao_entrega) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Data de previs�o de entrega informada � inv�lida"
					elseif StrToDate(c_data_previsao_entrega) <= Date then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Data de previs�o de entrega deve ser uma data futura"
						end if
					end if
				end if
			end if
		end if

	if alerta <> "" then blnErroConsistencia=True
	
	
'	MENSAGEM DE ALERTA
	dim s_descricao_forma_pagto, s_descricao_forma_pagto_anterior, quebraLinhaFormaPagto
	dim s_indicador_anterior
	s_descricao_forma_pagto = ""
	s_descricao_forma_pagto_anterior = ""
	quebraLinhaFormaPagto = ",  "


'	BANCO DE DADOS
'	==============
	dim vLogFP1()
	dim vLogFP2()
	dim s_logFP
	dim campos_a_omitir_FP
	dim s_log_FP
	s_log_FP = ""
	dim vLogItemCFF1()  'Custo Financeiro por Fornecedor
	dim vLogItemCFF2()
	dim campos_a_omitir_ItemCFF
	dim s_log_ItemCFF
	s_log_ItemCFF = ""
	dim vLogPedCFF1()  'Custo Financeiro por Fornecedor
	dim vLogPedCFF2()
	dim campos_a_omitir_PedCFF
	dim s_log_PedCFF
	s_log_PedCFF = ""
	dim vLogItemDevolvido1()  'Devolu��o de mercadorias
	dim vLogItemDevolvido2()
	dim campos_a_omitir_ItemDevolvido
	dim s_log_ItemDevolvido
	s_log_ItemDevolvido = ""
    dim s_log_Frete_Altera
    s_log_Frete_Altera = ""
    dim s_log_Frete_Exclui
    s_log_Frete_Exclui = ""
	dim vLog1()
	dim vLog2()
	dim s_log, s_log_manual
	dim campos_a_omitir
	s_log = ""
	s_log_manual = ""
	campos_a_omitir = "|a_entregar_data|a_entregar_hora|etg_imediata_data|etg_imediata_usuario|PedidoRecebidoDtHrUltAtualiz|PedidoRecebidoUsuarioUltAtualiz|InstaladorInstalaUsuarioUltAtualiz|InstaladorInstalaDtHrUltAtualiz|"
	campos_a_omitir_FP = campos_a_omitir & "|analise_credito|st_recebido|"
	campos_a_omitir_ItemCFF = ""
	campos_a_omitir_PedCFF = ""

	pedido_base = retorna_num_pedido_base(pedido_selecionado)
	
	if alerta = "" then
	'	ATUALIZA O PEDIDO
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if
		
		s_analise_credito_a = ""
		if IsPedidoFilhote(pedido_selecionado) then
			s = "SELECT * FROM t_PEDIDO WHERE pedido='" & pedido_base & "'"
			rs.Open s, cn
			if Err <> 0 then
				alerta = Cstr(Err) & ": " & Err.Description
			elseif rs.EOF then
				alerta = "Pedido base " & pedido_base & " n�o foi encontrado."
			else
				log_via_vetor_carrega_do_recordset rs, vLogFP1, campos_a_omitir_FP
				s_analise_credito_a = Trim("" & rs("analise_credito"))
				if blnAnaliseCreditoEdicaoLiberada then
					if s_analise_credito <> "" then 
						if CLng(rs("analise_credito")) <> CLng(s_analise_credito) then
                            if s_analise_credito = COD_AN_CREDITO_PENDENTE_VENDAS then
								blnEnviarEmailVendedorStatusAnaliseCredito = True
                                if s_ac_pendente_vendas_motivo = "" then
                                    alerta = "N�o foi informado o motivo do status 'Pendente Vendas'."
                                end if
                            end if
							rs("analise_credito")=CLng(s_analise_credito)
							rs("analise_credito_data")=Now
							rs("analise_credito_usuario")=usuario
							end if
                        if s_analise_credito <> COD_AN_CREDITO_PENDENTE_VENDAS then
                            if CStr(Trim("" & rs("analise_credito_pendente_vendas_motivo"))) <> "" then rs("analise_credito_pendente_vendas_motivo") = ""
                            end if
                        if Trim("" & rs("analise_credito_pendente_vendas_motivo")) <> Trim(s_ac_pendente_vendas_motivo) then
                                if s_analise_credito = COD_AN_CREDITO_PENDENTE_VENDAS then
                                    rs("analise_credito_pendente_vendas_motivo") = s_ac_pendente_vendas_motivo
							        rs("analise_credito_data")=Now
							        rs("analise_credito_usuario")=usuario
                                    end if
                                end if
						end if
					end if
					
			'	Forma de Pagamento (nova vers�o)
				if (versao_forma_pagamento = "2") And blnFormaPagtoEdicaoLiberada then
					s_descricao_forma_pagto_anterior = monta_descricao_forma_pagto_com_quebra_linha(rs, quebraLinhaFormaPagto)

					rs("tipo_parcelamento")=CLng(rb_forma_pagto)
				'	Limpa os campos n�o usados p/ facilitar a consulta multicrit�rio e p/ que na pr�xima altera��o, se houver, o log perceba a altera��o (ex: parcelado c/ entrada mudou p/ parcelado no cart�o; ao alterar novamente p/ parcelado c/ entrada e forem preenchidos os mesmos valores, ser� percebida e registrada apenas a mudan�a da op��o "parcelado c/ entrada", j� os demais campos ficaram c/ os mesmos valores).
					rs("av_forma_pagto")=0
					rs("pu_forma_pagto") = 0
					rs("pu_valor") = 0
					rs("pu_vencto_apos") = 0
					rs("pc_qtde_parcelas")=0
					rs("pc_valor_parcela")=0
					rs("pc_maquineta_qtde_parcelas")=0
					rs("pc_maquineta_valor_parcela")=0
					rs("pce_forma_pagto_entrada")=0
					rs("pce_forma_pagto_prestacao")=0
					rs("pce_entrada_valor")=0
					rs("pce_prestacao_qtde")=0
					rs("pce_prestacao_valor")=0
					rs("pce_prestacao_periodo")=0
					rs("pse_forma_pagto_prim_prest")=0
					rs("pse_forma_pagto_demais_prest")=0
					rs("pse_prim_prest_valor")=0
					rs("pse_prim_prest_apos")=0
					rs("pse_demais_prest_qtde")=0
					rs("pse_demais_prest_valor")=0
					rs("pse_demais_prest_periodo")=0
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

					s_descricao_forma_pagto = monta_descricao_forma_pagto_com_quebra_linha(rs, quebraLinhaFormaPagto)
					if (s_descricao_forma_pagto <> s_descricao_forma_pagto_anterior) And (Trim("" & rs("analise_credito")) = COD_AN_CREDITO_OK) then
						'Envia mensagem de alerta sobre edi��o na forma de pagamento em pedido com status "cr�dito ok"
						set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoFormaPagtoEmPedidoCreditoOk)
						if Trim("" & rEmailDestinatario.campo_texto) <> "" then
							corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na Central a forma de pagamento do pedido " & pedido_selecionado & vbCrLf & _
												vbCrLf & _
												"Forma de pagamento anterior:" & vbCrLf & _
												s_descricao_forma_pagto_anterior & vbCrLf & _
												vbCrLf & _
												"Forma de pagamento atual:" & vbCrLf & _
												s_descricao_forma_pagto

							EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
															"", _
															rEmailDestinatario.campo_texto, _
															"", _
															"", _
															"Edi��o da forma de pagamento em pedido com status 'Cr�dito OK'", _
															corpo_mensagem, _
															Now, _
															id_email, _
															msg_erro_grava_email
							end if
						end if
					end if
					
				if bln_RT_e_RA_EdicaoLiberada then rs("perc_RT") = converte_numero(s_perc_RT)
				
				if blnIndicadorEdicaoLiberada then
					s_indicador_anterior = Trim("" & rs("indicador"))
					rs("indicador") = s_indicador

					if (Ucase(Trim(s_indicador_anterior)) <> Ucase(Trim(s_indicador))) And (Trim("" & rs("analise_credito")) = COD_AN_CREDITO_OK) then
						'Envia mensagem de alerta sobre altera��o do indicador em pedido com status "cr�dito ok"
						set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaAlteracaoIndicadorEmPedidoCreditoOk)
						if Trim("" & rEmailDestinatario.campo_texto) <> "" then
							corpo_mensagem = "O usu�rio '" & usuario & "' alterou em " & formata_data_hora_sem_seg(Now) & " na Central o indicador do pedido " & pedido_selecionado & vbCrLf & _
												vbCrLf & _
												"Indicador anterior:" & vbCrLf & _
												Ucase(Trim(s_indicador_anterior)) & vbCrLf & _
												vbCrLf & _
												"Indicador atual:" & vbCrLf & _
												Ucase(Trim(s_indicador))

							EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
															"", _
															rEmailDestinatario.campo_texto, _
															"", _
															"", _
															"Altera��o do indicador em pedido com status 'Cr�dito OK'", _
															corpo_mensagem, _
															Now, _
															id_email, _
															msg_erro_grava_email
							end if
						end if
					end if

				if blnNumPedidoECommerceEdicaoLiberada then rs("pedido_bs_x_ac") = CStr(s_pedido_ac)

				if blnNumPedidoECommerceEdicaoLiberada And (c_loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then
					rs("pedido_bs_x_marketplace")=s_pedido_mktplace
					if Trim("" & rs("marketplace_codigo_origem")) <> s_pedido_origem then
						s = "SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo = 'PedidoECommerce_Origem') AND (codigo = '" & s_pedido_origem & "')"
						set rs2 = cn.execute(s)
						if Not rs2.Eof then
						'	OBT�M O PERCENTUAL DE COMISS�O DO MARKETPLACE E SE DEVE COLOCAR AUTOMATICAMENTE COM 'CR�DITO OK'
							s = "SELECT * FROM T_CODIGO_DESCRICAO WHERE (grupo = 'PedidoECommerce_Origem_Grupo') AND (codigo = '" & Trim("" & rs2("codigo_pai")) & "')"
							set rs2 = cn.execute(s)
							if Not rs2.Eof then
								if rs("perc_RT") <> rs2("parametro_campo_real") then
									if s_log_manual <> "" then s_log_manual = s_log_manual & "; "
									s_log_manual = s_log_manual & "perc_RT: " & rs("perc_RT") & " => " & rs2("parametro_campo_real") & " (taxa marketplace)"
									rs("perc_RT") = rs2("parametro_campo_real")
									end if
								if rs2("parametro_1_campo_flag") = 1 then
									if rs("analise_credito") <> Clng(COD_AN_CREDITO_OK) then
										if s_log_manual <> "" then s_log_manual = s_log_manual & "; "
										s_log_manual = s_log_manual & "analise_credito: " & rs("analise_credito") & " => " & COD_AN_CREDITO_OK & " (Cr�dito Ok autom�tico para pedido de marketplace)"
										rs("analise_credito")=Clng(COD_AN_CREDITO_OK)
										rs("analise_credito_data")=Now
										rs("analise_credito_usuario")="AUTOM�TICO"
										end if
									end if
								end if
							end if
						
						blnMarketplaceCodigoOrigemAlterado = True
						rs("marketplace_codigo_origem")=s_pedido_origem
						end if
					end if

				rs("pedido_bs_x_at")=s_ped_bonshop

				rs("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP

				rs.Update
				log_via_vetor_carrega_do_recordset rs, vLogFP2, campos_a_omitir_FP
				s_log_FP = log_via_vetor_monta_alteracao(vLogFP1, vLogFP2)
				if Err <> 0 then
					alerta = Cstr(Err) & ": " & Err.Description
					end if
				end if
			end if

		if alerta = "" then
			s = "SELECT * FROM t_PEDIDO WHERE pedido='" & pedido_selecionado & "'"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Err <> 0 then
				alerta = Cstr(Err) & ": " & Err.Description
			elseif rs.EOF then
				alerta = "Pedido " & pedido_selecionado & " n�o foi encontrado."
			else
				log_via_vetor_carrega_do_recordset rs, vLog1, campos_a_omitir
				if Not IsPedidoFilhote(pedido_selecionado) then 
					s_analise_credito_a = Trim("" & rs("analise_credito"))
					if blnAnaliseCreditoEdicaoLiberada then
						if s_analise_credito <> "" then 
							if CLng(rs("analise_credito")) <> CLng(s_analise_credito) then
                                if s_analise_credito = COD_AN_CREDITO_PENDENTE_VENDAS then
									blnEnviarEmailVendedorStatusAnaliseCredito = True
                                    if s_ac_pendente_vendas_motivo = "" then
                                        alerta = "N�o foi informado o motivo do status 'Pendente Vendas'."
                                    end if
                                end if
								rs("analise_credito")=CLng(s_analise_credito)
								rs("analise_credito_data")=Now
								rs("analise_credito_usuario")=usuario
								end if
                            if s_analise_credito <> COD_AN_CREDITO_PENDENTE_VENDAS then
                                if CStr(Trim("" & rs("analise_credito_pendente_vendas_motivo"))) <> "" then rs("analise_credito_pendente_vendas_motivo") = ""
                                end if
                            if Cstr(Trim("" & rs("analise_credito_pendente_vendas_motivo"))) <> Cstr(s_ac_pendente_vendas_motivo) then
                                if s_analise_credito = COD_AN_CREDITO_PENDENTE_VENDAS then
                                    rs("analise_credito_pendente_vendas_motivo") = s_ac_pendente_vendas_motivo
								    rs("analise_credito_data")=Now
								    rs("analise_credito_usuario")=usuario
                                    end if
                                end if
							end if
						end if

				'	Forma de Pagamento (nova vers�o)
					if (versao_forma_pagamento = "2") And blnFormaPagtoEdicaoLiberada then
						s_descricao_forma_pagto_anterior = monta_descricao_forma_pagto_com_quebra_linha(rs, quebraLinhaFormaPagto)

						rs("tipo_parcelamento")=CLng(rb_forma_pagto)
					'	Limpa os campos n�o usados p/ facilitar a consulta multicrit�rio e p/ que na pr�xima altera��o, se houver, o log perceba a altera��o (ex: parcelado c/ entrada mudou p/ parcelado no cart�o; ao alterar novamente p/ parcelado c/ entrada e forem preenchidos os mesmos valores, ser� percebida e registrada apenas a mudan�a da op��o "parcelado c/ entrada", j� os demais campos ficaram c/ os mesmos valores).
						rs("av_forma_pagto")=0
						rs("pu_forma_pagto") = 0
						rs("pu_valor") = 0
						rs("pu_vencto_apos") = 0
						rs("pc_qtde_parcelas")=0
						rs("pc_valor_parcela")=0
						rs("pc_maquineta_qtde_parcelas")=0
						rs("pc_maquineta_valor_parcela")=0
						rs("pce_forma_pagto_entrada")=0
						rs("pce_forma_pagto_prestacao")=0
						rs("pce_entrada_valor")=0
						rs("pce_prestacao_qtde")=0
						rs("pce_prestacao_valor")=0
						rs("pce_prestacao_periodo")=0
						rs("pse_forma_pagto_prim_prest")=0
						rs("pse_forma_pagto_demais_prest")=0
						rs("pse_prim_prest_valor")=0
						rs("pse_prim_prest_apos")=0
						rs("pse_demais_prest_qtde")=0
						rs("pse_demais_prest_valor")=0
						rs("pse_demais_prest_periodo")=0
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

						s_descricao_forma_pagto = monta_descricao_forma_pagto_com_quebra_linha(rs, quebraLinhaFormaPagto)
						if (s_descricao_forma_pagto <> s_descricao_forma_pagto_anterior) And (Trim("" & rs("analise_credito")) = COD_AN_CREDITO_OK) then
							'Envia mensagem de alerta sobre edi��o na forma de pagamento em pedido com status "cr�dito ok"
							set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoFormaPagtoEmPedidoCreditoOk)
							if Trim("" & rEmailDestinatario.campo_texto) <> "" then
								corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na Central a forma de pagamento do pedido " & pedido_selecionado & vbCrLf & _
												vbCrLf & _
												"Forma de pagamento anterior:" & vbCrLf & _
												s_descricao_forma_pagto_anterior & vbCrLf & _
												vbCrLf & _
												"Forma de pagamento atual:" & vbCrLf & _
												s_descricao_forma_pagto

								EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
																"", _
																rEmailDestinatario.campo_texto, _
																"", _
																"", _
																"Edi��o da forma de pagamento em pedido com status 'Cr�dito OK'", _
																corpo_mensagem, _
																Now, _
																id_email, _
																msg_erro_grava_email
								end if
							end if
						end if
					end if
				
				if bln_RT_e_RA_EdicaoLiberada then rs("perc_RT") = converte_numero(s_perc_RT)

                if blnIndicadorEdicaoLiberada then
					s_indicador_anterior = Trim("" & rs("indicador"))
					rs("indicador") = s_indicador

					if (Ucase(Trim(s_indicador_anterior)) <> Ucase(Trim(s_indicador))) And (Trim("" & rs("analise_credito")) = COD_AN_CREDITO_OK) then
						'Envia mensagem de alerta sobre altera��o do indicador em pedido com status "cr�dito ok"
						set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaAlteracaoIndicadorEmPedidoCreditoOk)
						if Trim("" & rEmailDestinatario.campo_texto) <> "" then
							corpo_mensagem = "O usu�rio '" & usuario & "' alterou em " & formata_data_hora_sem_seg(Now) & " na Central o indicador do pedido " & pedido_selecionado & vbCrLf & _
												vbCrLf & _
												"Indicador anterior:" & vbCrLf & _
												Ucase(Trim(s_indicador_anterior)) & vbCrLf & _
												vbCrLf & _
												"Indicador atual:" & vbCrLf & _
												Ucase(Trim(s_indicador))

							EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
															"", _
															rEmailDestinatario.campo_texto, _
															"", _
															"", _
															"Altera��o do indicador em pedido com status 'Cr�dito OK'", _
															corpo_mensagem, _
															Now, _
															id_email, _
															msg_erro_grava_email
							end if
						end if
					end if

				if blnObs1EdicaoLiberada then
                     rs("obs_1") = s_obs1
                     rs("NFe_texto_constar") = s_nf_texto
                     rs("NFe_xPed") = s_num_pedido_compra
                end if

				if blnObs2EdicaoLiberada then
					'Usu�rio fez altera��o do campo na p�gina de edi��o?
					'Obs: controle feito com o objetivo de evitar que altera��es realizadas por outros processos enquanto o usu�rio estava na p�gina de edi��o sejam sobrescritas
					if s_obs2 <> s_obs2_original then
						rs("obs_2") = s_obs2
						end if
					end if
				
				if blnObs3EdicaoLiberada then
					'Usu�rio fez altera��o do campo na p�gina de edi��o?
					'Obs: controle feito com o objetivo de evitar que altera��es realizadas por outros processos enquanto o usu�rio estava na p�gina de edi��o sejam sobrescritas
					if s_obs3 <> s_obs3_original then
						rs("obs_3") = s_obs3
						end if
					end if
				
				if blnFormaPagtoEdicaoLiberada then rs("forma_pagto") = s_forma_pagto
				
				if (versao_forma_pagamento = "1") And blnFormaPagtoEdicaoLiberada then
					if IsNumeric(s_qtde_parcelas) then 
						rs("qtde_parcelas") = CLng(s_qtde_parcelas)
					else
						rs("qtde_parcelas") = 0
						end if
					end if
				
				if blnNumPedidoECommerceEdicaoLiberada then rs("pedido_bs_x_ac") = CStr(s_pedido_ac)

				if blnNumPedidoECommerceEdicaoLiberada And (c_loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then
					rs("pedido_bs_x_marketplace")=s_pedido_mktplace
					if Trim("" & rs("marketplace_codigo_origem")) <> s_pedido_origem then
						s = "SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo = 'PedidoECommerce_Origem') AND (codigo = '" & s_pedido_origem & "')"
						set rs2 = cn.execute(s)
						if Not rs2.Eof then
						'	OBT�M O PERCENTUAL DE COMISS�O DO MARKETPLACE E SE DEVE COLOCAR AUTOMATICAMENTE COM 'CR�DITO OK'
							s = "SELECT * FROM T_CODIGO_DESCRICAO WHERE (grupo = 'PedidoECommerce_Origem_Grupo') AND (codigo = '" & Trim("" & rs2("codigo_pai")) & "')"
							set rs2 = cn.execute(s)
							if Not rs2.Eof then
								if rs("perc_RT") <> rs2("parametro_campo_real") then
									if s_log_manual <> "" then s_log_manual = s_log_manual & "; "
									s_log_manual = s_log_manual & "perc_RT: " & rs("perc_RT") & " => " & rs2("parametro_campo_real") & " (taxa marketplace)"
									rs("perc_RT") = rs2("parametro_campo_real")
									end if
								if rs2("parametro_1_campo_flag") = 1 then
									if rs("analise_credito") <> Clng(COD_AN_CREDITO_OK) then
										if s_log_manual <> "" then s_log_manual = s_log_manual & "; "
										s_log_manual = s_log_manual & "analise_credito: " & rs("analise_credito") & " => " & COD_AN_CREDITO_OK & " (Cr�dito Ok autom�tico para pedido de marketplace)"
										rs("analise_credito")=Clng(COD_AN_CREDITO_OK)
										rs("analise_credito_data")=Now
										rs("analise_credito_usuario")="AUTOM�TICO"
										end if
									end if
								end if
							end if
						
						blnMarketplaceCodigoOrigemAlterado = True
						rs("marketplace_codigo_origem")=s_pedido_origem
						end if
					end if
				
				if blnAEntregarStatusEdicaoLiberada then
					'Usu�rio fez altera��o da data na p�gina de edi��o?
					'Obs: controle feito com o objetivo de evitar que altera��es realizadas por outros processos enquanto o usu�rio estava na p�gina de edi��o sejam sobrescritas
					if c_a_entregar_data_marcada <> c_a_entregar_data_marcada_original then
						if IsDate(c_a_entregar_data_marcada) then
							blnFlag = False
							if Trim("" & rs("a_entregar_status")) <> "1" then blnFlag = True
							rs("a_entregar_status")=1
							if formata_data(rs("a_entregar_data_marcada")) <> formata_data(StrToDate(c_a_entregar_data_marcada)) then blnFlag = True
							rs("a_entregar_data_marcada")=StrToDate(c_a_entregar_data_marcada)
							if blnFlag then
								rs("a_entregar_data")=Date
								rs("a_entregar_hora")=retorna_so_digitos(formata_hora(Now))
								rs("a_entregar_usuario")=usuario
								end if
						elseif rs("a_entregar_status")<>0 then
							blnFlag = False
							if Trim("" & rs("a_entregar_status")) <> "0" then blnFlag = True
							rs("a_entregar_status")=0
							if Trim("" & rs("a_entregar_data_marcada")) <> "" then blnFlag = True
							rs("a_entregar_data_marcada")=Null
							if blnFlag then
								rs("a_entregar_data")=Date
								rs("a_entregar_hora")=retorna_so_digitos(formata_hora(Now))
								rs("a_entregar_usuario")=usuario
								end if
							end if
						end if
					end if
				
				'Guardar informa��es de endere�o presentes no pedido (consist�ncia para verificar se mudou de endere�o)
				dim st_end_entrega_anterior, EndEtg_cep_anterior, blnEndereco_cep_alterado
				st_end_entrega_anterior = rs("st_end_entrega")
				EndEtg_cep_anterior = rs("EndEtg_cep")
				blnEndereco_cep_alterado = False
	
				if r_pedido.st_memorizacao_completa_enderecos = 1 or r_pedido.st_memorizacao_completa_enderecos = 9 then
					if rs("endereco_cep") <> endereco__cep then blnEndereco_cep_alterado = True    
					rs("st_memorizacao_completa_enderecos") = 1
					rs("endereco_bairro") = endereco__bairro
					rs("endereco_numero") = endereco__numero
					rs("endereco_complemento") = endereco__complemento
					rs("endereco_cidade") = endereco__cidade
					rs("endereco_uf") = endereco__uf
					rs("endereco_cep") = endereco__cep
					rs("endereco_logradouro") = endereco__endereco
					rs("endereco_ddd_res") = cliente__ddd_res
					rs("endereco_tel_res") = cliente__tel_res				
					rs("endereco_ddd_cel") = cliente__ddd_cel
					rs("endereco_tel_cel") = cliente__tel_cel				
					rs("endereco_ddd_com") = cliente__ddd_com
					rs("endereco_tel_com") = cliente__tel_com
					rs("endereco_ramal_com") = cliente__ramal_com	
					rs("endereco_ddd_com_2") = cliente__ddd_com_2
					rs("endereco_tel_com_2") = cliente__tel_com_2
					rs("endereco_ramal_com_2") = cliente__ramal_com_2
					rs("endereco_email") = cliente__email
					rs("endereco_email_xml") = cliente__email_xml
					rs("endereco_nome") = cliente__nome 
					rs("endereco_ie") = cliente__ie
					rs("endereco_contribuinte_icms_status") = cliente__contribuinte_icms_status

					if eh_cpf then
						rs("endereco_rg") = cliente__rg					
						rs("endereco_produtor_rural_status") = cliente__produtor_rural
					end if
				end if


				'Edit�vel?
				if blnEndEntregaEdicaoLiberada then
					if EndEtg_endereco <> "" then 
						rs("st_end_entrega") = 1
					else
						rs("st_end_entrega") = 0
						end if
					
					rs("EndEtg_endereco") = EndEtg_endereco
					rs("EndEtg_endereco_numero") = EndEtg_endereco_numero
					rs("EndEtg_endereco_complemento") = EndEtg_endereco_complemento
					rs("EndEtg_bairro") = EndEtg_bairro
					rs("EndEtg_cidade") = EndEtg_cidade
					rs("EndEtg_uf") = EndEtg_uf
					rs("EndEtg_cep") = EndEtg_cep
                    rs("EndEtg_cod_justificativa") = EndEtg_obs
                	if not blnUsarMemorizacaoCompletaEnderecos then
                        rs("st_memorizacao_completa_enderecos") = 0
                        end if
                	if r_pedido.st_memorizacao_completa_enderecos <> 0 and blnUsarMemorizacaoCompletaEnderecos then
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
				
				'Edit�vel?
				if blnTransportadoraEdicaoLiberada then
					'Usu�rio fez altera��o de transportadora na p�gina de edi��o?
					'Obs: controle feito com o objetivo de evitar que altera��es realizadas por outros processos enquanto o usu�rio estava na p�gina de edi��o sejam sobrescritas
					if c_transportadora_id <> c_transportadora_id_original then
						blnEditouTransp = False
						if UCase(Trim("" & rs("transportadora_id"))) <> UCase(c_transportadora_id) then blnEditouTransp = True
						if UCase(Trim("" & rs("transportadora_num_coleta"))) <> UCase(c_transportadora_num_coleta) then blnEditouTransp = True
						if UCase(Trim("" & rs("transportadora_contato"))) <> UCase(c_transportadora_contato) then blnEditouTransp = True
						if blnEditouTransp then
							if UCase(Trim("" & rs("transportadora_id"))) <> UCase(c_transportadora_id) then
							'	LIMPA DADOS DA SELE��O AUTOM�TICA DE TRANSPORTADORA BASEADO NO CEP
							'	MANT�M OS DADOS ANTERIORES (SE HOUVER) P/ FINS DE HIST�RICO/LOG DOS SEGUINTES CAMPOS:
							'	transportadora_selecao_auto_cep, transportadora_selecao_auto_tipo_endereco e transportadora_selecao_auto_transportadora
								rs("transportadora_selecao_auto_status") = TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_N
								rs("transportadora_selecao_auto_data_hora") = Now
								end if
							rs("transportadora_id") = c_transportadora_id
							rs("transportadora_num_coleta") = c_transportadora_num_coleta
							rs("transportadora_contato") = c_transportadora_contato
							rs("transportadora_data")=Now
							rs("transportadora_usuario")=usuario
							end if
						end if
					end if
				
			'	SELE��O AUTOM�TICA DA TRANSPORTADORA COM BASE NO CEP
				blnProcessaSelecaoAutoTransp = False
				if (Not blnEditouTransp) And (Not IsPedidoEncerrado(r_pedido.st_entrega)) And _
					(Not blnNFEmitida) And (Trim("" & rs("analise_credito")) <> Trim("" & COD_AN_CREDITO_OK)) then
					if CInt(st_end_entrega_anterior) <> CInt(rs("st_end_entrega")) then
					'	HOUVE ALTERA��O ENTRE USAR O ENDERE�O DE ENTREGA E O ENDERE�O DO CADASTRO (OU VICE-VERSA)
						blnProcessaSelecaoAutoTransp = True
					else
						if CInt(rs("st_end_entrega")) <> 0 then
						'	OBS: ALTERA��ES NO ENDERE�O DO CADASTRO S�O PROCESSADAS NA P�GINA CLIENTEATUALIZA.ASP
						'	H� ENDERE�O DE ENTREGA: O CEP MUDOU?
							if Trim("" & EndEtg_cep_anterior) <> Trim("" & rs("EndEtg_cep")) then blnProcessaSelecaoAutoTransp = True
						else
							'   Se teve altera��o no endereco_cep vamos recalcular a transportadora
							if blnEndereco_cep_alterado then blnProcessaSelecaoAutoTransp = True
							end if
						end if
					end if
				
				if blnProcessaSelecaoAutoTransp then
					dim sTranspSelAutoTransportadoraId, sTranspSelAutoCep, iTranspSelAutoTipoEndereco, iTranspSelAutoStatus
					sTranspSelAutoTransportadoraId = ""
					if CInt(rs("st_end_entrega")) <> 0 then
						if EndEtg_cep <> "" then sTranspSelAutoTransportadoraId = obtem_transportadora_pelo_cep(retorna_so_digitos(EndEtg_cep))
						sTranspSelAutoCep = retorna_so_digitos(EndEtg_cep)
						iTranspSelAutoTipoEndereco = TRANSPORTADORA_SELECAO_AUTO_TIPO_ENDERECO_ENTREGA
						iTranspSelAutoStatus = TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_S
					else
						if endereco__cep <> "" then sTranspSelAutoTransportadoraId = obtem_transportadora_pelo_cep(retorna_so_digitos(endereco__cep))
						sTranspSelAutoCep = retorna_so_digitos(endereco__cep)
						iTranspSelAutoTipoEndereco = TRANSPORTADORA_SELECAO_AUTO_TIPO_ENDERECO_CLIENTE
						iTranspSelAutoStatus = TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_S
						end if
					
				'	ALTERAR SE A TRANSPORTADORA FOR DIFERENTE DA QUE EST� GRAVADA
					if Ucase(sTranspSelAutoTransportadoraId) <> Ucase(Trim("" & rs("transportadora_id"))) then
						rs("transportadora_id") = sTranspSelAutoTransportadoraId
						rs("transportadora_data") = Now
						rs("transportadora_usuario") = usuario
						rs("transportadora_selecao_auto_status") = iTranspSelAutoStatus
						rs("transportadora_selecao_auto_cep") = sTranspSelAutoCep
						rs("transportadora_selecao_auto_transportadora") = sTranspSelAutoTransportadoraId
						rs("transportadora_selecao_auto_tipo_endereco") = iTranspSelAutoTipoEndereco
						rs("transportadora_selecao_auto_data_hora") = Now
						end if
					end if
				

				if blnPedidoRecebidoStatusEdicaoLiberada then
					blnEditou = False
					if CStr(rs("PedidoRecebidoStatus")) <> Cstr(rb_PedidoRecebidoStatus) then 
					'	ALTEROU STATUS DE RECEBIMENTO DO PEDIDO?
						blnEditou = True
					else
					'	STATUS PERMANECE IGUAL, COMO "RECEBIDO", MAS ALTEROU A DATA DO RECEBIMENTO?
						if Cstr(rb_PedidoRecebidoStatus) = Cstr(COD_ST_PEDIDO_RECEBIDO_SIM) then
							if formata_data(rs("PedidoRecebidoData")) <> formata_data(StrToDate(c_PedidoRecebidoData)) then blnEditou = True
							end if
						end if
					
					if blnEditou then
						rs("PedidoRecebidoStatus") = CLng(rb_PedidoRecebidoStatus)
						if Cstr(rb_PedidoRecebidoStatus) = Cstr(COD_ST_PEDIDO_RECEBIDO_SIM) then
							rs("PedidoRecebidoData") = StrToDate(c_PedidoRecebidoData)
						else
							rs("PedidoRecebidoData") = Null
							end if
						rs("PedidoRecebidoUsuarioUltAtualiz") = usuario
						rs("PedidoRecebidoDtHrUltAtualiz") = Now
						end if
					end if
					
				if blnEntregaImediataEdicaoLiberada then
					s_etg_imediata_original = Trim("" & rs("st_etg_imediata"))
					if s_etg_imediata <> "" then 
						if CLng(rs("st_etg_imediata")) <> CLng(s_etg_imediata) then
							rs("st_etg_imediata")=CLng(s_etg_imediata)
							rs("etg_imediata_data")=Now
							rs("etg_imediata_usuario")=usuario
							end if
						end if
					
					if CLng(s_etg_imediata) = CLng(COD_ETG_IMEDIATA_NAO) then
						if s_etg_imediata_original <> Trim(s_etg_imediata) then
							rs("PrevisaoEntregaData") = StrToDate(c_data_previsao_entrega)
							rs("PrevisaoEntregaUsuarioUltAtualiz") = usuario
							rs("PrevisaoEntregaDtHrUltAtualiz") = Now
						elseif blnEntregaImediataNaoSemDataPrevisao And (c_data_previsao_entrega = "") then
							'NOP
							'O STATUS DA ENTREGA IMEDIATA J� ESTAVA COMO N�O E SEM DATA DE PREVIS�O, PORTANTO, N�O FAZ NADA
						elseif (s_etg_imediata_original <> Trim(s_etg_imediata)) Or (formata_data(rs("PrevisaoEntregaData")) <> formata_data(StrToDate(c_data_previsao_entrega))) then
							rs("PrevisaoEntregaData") = StrToDate(c_data_previsao_entrega)
							rs("PrevisaoEntregaUsuarioUltAtualiz") = usuario
							rs("PrevisaoEntregaDtHrUltAtualiz") = Now
							end if
					else
						if (s_etg_imediata_original <> Trim(s_etg_imediata)) then
							rs("PrevisaoEntregaData") = Null
							rs("PrevisaoEntregaUsuarioUltAtualiz") = usuario
							rs("PrevisaoEntregaDtHrUltAtualiz") = Now
							end if
						end if
					end if

				if blnBemUsoConsumoEdicaoLiberada then
					if s_bem_uso_consumo <> "" then 
						if CLng(rs("StBemUsoConsumo")) <> CLng(s_bem_uso_consumo) then
							rs("StBemUsoConsumo")=CLng(s_bem_uso_consumo)
							end if
						end if
					end if
					
				if blnInstaladorInstalaEdicaoLiberada then
					if s_instalador_instala <> "" then
						if CLng(rs("InstaladorInstalaStatus")) <> CLng(s_instalador_instala) then
							rs("InstaladorInstalaStatus")=CLng(s_instalador_instala)
							rs("InstaladorInstalaUsuarioUltAtualiz")=usuario
							rs("InstaladorInstalaDtHrUltAtualiz")=Now
							end if
						end if
					end if
					
				if blnGarantiaIndicadorEdicaoLiberada then
					if CLng(rs("GarantiaIndicadorStatus")) <> CLng(rb_garantia_indicador) then
						rs("GarantiaIndicadorStatus")=CLng(rb_garantia_indicador)
						rs("GarantiaIndicadorUsuarioUltAtualiz")=usuario
						rs("GarantiaIndicadorDtHrUltAtualiz")=Now
						end if
					end if
				rs("pedido_bs_x_at")=s_ped_bonshop
				if blnAtivarFlag_st_violado_permite_RA_status then
					rs("st_violado_permite_RA_status") = 1
					rs("dt_hr_violado_permite_RA_status") = Now
					rs("usuario_violado_permite_RA_status") = usuario
					end if
				
				rs("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP

				rs.Update 
				if Err <> 0 then
					alerta = Cstr(Err) & ": " & Err.Description
				else
					log_via_vetor_carrega_do_recordset rs, vLog2, campos_a_omitir
					s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
					if blnAnaliseCreditoEdicaoLiberada then
						if (s_analise_credito<>"") And (s_analise_credito<>s_analise_credito_a) And (Instr(s_log,"analise_credito")=0) then
							if s_log <> "" then s_log = s_log & "; "
							s_log = s_log & "analise_credito: " & formata_texto_log(s_analise_credito_a) & " => " & formata_texto_log(s_analise_credito)
							end if
						end if
					if s_log_manual <> "" then
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & s_log_manual
						end if
					end if
				end if
			end if

		if alerta = "" then
		'	O PEDIDO FOI CADASTRADO J� DENTRO DA POL�TICA DE PERCENTUAL DE CUSTO FINANCEIRO POR FORNECEDOR?
			if c_custoFinancFornecTipoParcelamentoOriginal <> "" then
				if blnFormaPagtoEdicaoLiberada then
					if (c_custoFinancFornecTipoParcelamentoOriginal <> c_custoFinancFornecTipoParcelamento) Or _
					   (c_custoFinancFornecQtdeParcelasOriginal <> c_custoFinancFornecQtdeParcelas) then
						for i=Lbound(v_item) to Ubound(v_item)
							with v_item(i)
								if Trim(.produto)<>"" then
								'	Inicializa��o
									vlCustoFinancFornecPrecoListaBase = 0
									coeficiente = 0
									
								'	Obt�m Pre�o de Lista Base
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
											" AND (loja='" & c_loja & "')"
									set rs2 = cn.execute(s)
									if rs2.Eof then
										alerta=texto_add_br(alerta)
										alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " N�O est� cadastrado para a loja " & c_loja
									else
										vlCustoFinancFornecPrecoListaBase = rs2("preco_lista")
										end if
									
								'	Obt�m coeficiente do custo financeiro
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
											end if
										end if
									
									if alerta = "" then
										s = "SELECT " & _
												"*" & _
											" FROM t_PEDIDO_ITEM" & _
											" WHERE" & _
												" (pedido='" & pedido_selecionado & "') AND" & _
												" (fabricante='" & Trim(.fabricante) & "') AND" & _
												" (produto='" & Trim(.produto) & "')"
										if rs.State <> 0 then rs.Close
										rs.Open s, cn
										if rs.Eof then
											alerta=texto_add_br(alerta)
											alerta=alerta & "Falha ao localizar o registro do produto " & .produto
										else
											log_via_vetor_carrega_do_recordset rs, vLogItemCFF1, campos_a_omitir_ItemCFF
											rs("custoFinancFornecCoeficiente")=coeficiente
											rs("custoFinancFornecPrecoListaBase")=vlCustoFinancFornecPrecoListaBase
											vlCustoFinancFornecPrecoLista=converte_numero(formata_moeda(coeficiente*vlCustoFinancFornecPrecoListaBase))
											rs("preco_lista")=vlCustoFinancFornecPrecoLista
											if vlCustoFinancFornecPrecoLista = 0 then 
												rs("desc_dado") = 0 
											else
												rs("desc_dado") = 100*(vlCustoFinancFornecPrecoLista-rs("preco_venda"))/vlCustoFinancFornecPrecoLista
												end if
											rs.Update
											log_via_vetor_carrega_do_recordset rs, vLogItemCFF2, campos_a_omitir_ItemCFF
											s = log_via_vetor_monta_alteracao(vLogItemCFF1, vLogItemCFF2)
											if s <> "" then
												if s_log_ItemCFF <> "" then s_log_ItemCFF=s_log_ItemCFF & "; "
												s_log_ItemCFF=s_log_ItemCFF & "produto (" & .fabricante & ")" & .produto & ": " & s
												end if
											end if
										end if
									end if
								end with
							next
						
						if alerta = "" then
							s = "SELECT " & _
									"*" & _
								" FROM t_PEDIDO" & _
								" WHERE" & _
								" pedido='" & pedido_base & "'"
							if rs.State <> 0 then rs.Close
							rs.Open s, cn
							if rs.Eof then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Falha ao localizar o registro do pedido " & pedido_base
							else
								log_via_vetor_carrega_do_recordset rs, vLogPedCFF1, campos_a_omitir_PedCFF
								rs("custoFinancFornecTipoParcelamento") = c_custoFinancFornecTipoParcelamento
								rs("custoFinancFornecQtdeParcelas") = c_custoFinancFornecQtdeParcelas
								rs.Update
								log_via_vetor_carrega_do_recordset rs, vLogPedCFF2, campos_a_omitir_PedCFF
								s_log_PedCFF = log_via_vetor_monta_alteracao(vLogPedCFF1, vLogPedCFF2)
								end if
							end if
						end if
					end if
				end if
			end if

		if alerta = "" then
		'	ATUALIZA DADOS DOS PRODUTOS
			for i=Lbound(v_item) to Ubound(v_item)
				with v_item(i)
					if Trim(.produto)<>"" then
						s = "SELECT " & _
								"*" & _
							" FROM t_PEDIDO_ITEM" & _
							" WHERE" & _
								" (pedido='" & pedido_selecionado & "') AND" & _
								" (fabricante='" & Trim(.fabricante) & "') AND" & _
								" (produto='" & Trim(.produto) & "')"
						if rs.State <> 0 then rs.Close
						rs.Open s, cn
						blnUpdate = False
						if Err <> 0 then
							alerta=texto_add_br(alerta)
							alerta=alerta & Cstr(Err) & ": " & Err.Description
						elseif rs.EOF then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Item do pedido do produto " & Trim(.produto) & " (fabricante " & Trim(.fabricante) & ") n�o foi encontrado."
						else
						'	EDITOU PRE�O DE VENDA?
							if blnItemPedidoEdicaoLiberada then
								if rs("preco_venda") <> .preco_venda then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "pre�o de venda do produto " & Trim(.produto) & " (" & Trim(.fabricante) & "): " & formata_moeda(rs("preco_venda")) & " => " & formata_moeda(.preco_venda)
									.preco_lista=rs("preco_lista")
									if .preco_lista = 0 then
										.desc_dado=0
									else
										.desc_dado=100*(.preco_lista-.preco_venda)/.preco_lista
										end if
									rs("preco_venda")=.preco_venda
									rs("desc_dado")=.desc_dado
									blnUpdate = True
									end if
								end if
							
						'	EDITOU PRE�O DE NF?
							if bln_RT_e_RA_EdicaoLiberada then
								if rs("preco_NF") <> .preco_NF then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "pre�o de NF do produto " & Trim(.produto) & " (" & Trim(.fabricante) & "): " & formata_moeda(rs("preco_NF")) & " => " & formata_moeda(.preco_NF)
									rs("preco_NF")=.preco_NF
									blnUpdate = True
									end if
								end if
							end if
						
						if blnUpdate then rs.Update
						end if
					end with
				next
			
		'	ATUALIZA O VALOR TOTAL DA FAM�LIA DE PEDIDOS
		'	OBT�M OS VALORES A PAGAR, J� PAGO E O STATUS DE PAGAMENTO (PARA TODA A FAM�LIA DE PEDIDOS)
		'	*** OBSERVA��O: A PEDIDO DO ROG�RIO DA ARTVEN, O STATUS DE PAGAMENTO N�O DEVE SER ATUALIZADO
			if Not calcula_pagamentos(pedido_selecionado, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then 
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if
				
			vl_total_RA = vl_TotalFamiliaPrecoNF - vl_TotalFamiliaPrecoVenda
			
			if Not calcula_total_RA_liquido_BD(pedido_selecionado, vl_total_RA_liquido, msg_erro) then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if
			
			id_pedido_base = retorna_num_pedido_base(pedido_selecionado)
			s = "SELECT * FROM t_PEDIDO WHERE (pedido='" & id_pedido_base & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Not rs.Eof then
				rs("vl_total_familia") = vl_TotalFamiliaPrecoVenda
				rs("vl_total_NF") = vl_TotalFamiliaPrecoNF
				rs("vl_total_RA") = vl_total_RA
				rs("vl_total_RA_liquido") = vl_total_RA_liquido
				rs("qtde_parcelas_desagio_RA") = 0
				if vl_total_RA <> 0 then
					rs("st_tem_desagio_RA") = 1
				else
					rs("st_tem_desagio_RA") = 0
					end if
				rs.Update
				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
				rs.Close 
				end if
			
		'	CONSIST�NCIA DO VALOR TOTAL DA FORMA DE PAGAMENTO
			if alerta = "" then
				if (versao_forma_pagamento = "2") And blnFormaPagtoEdicaoLiberada then
					vl_totalFamiliaPrecoNFLiquido = vl_TotalFamiliaPrecoNF - vl_TotalFamiliaDevolucaoPrecoNF
					if rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA then vlTotalFormaPagto = vl_totalFamiliaPrecoNFLiquido
					if Abs(vlTotalFormaPagto-vl_totalFamiliaPrecoNFLiquido) > 0.1 then
						alerta = "H� diverg�ncia entre o valor total do pedido (" & SIMBOLO_MONETARIO & " " & formata_moeda(vl_totalFamiliaPrecoNFLiquido) & ") e o valor total descrito atrav�s da forma de pagamento (" & SIMBOLO_MONETARIO & " " & formata_moeda(vlTotalFormaPagto) & ")!!"
						end if
					end if
				end if
				
			end if
		
		if alerta = "" then
		'	ANOTA��O DO N� NFe NOS ITENS DEVOLVIDOS
			if blnDadosNFeMercadoriasDevolvidasEdicaoLiberada then
				for i=LBound(v_nfe_item_devolvido) to Ubound(v_nfe_item_devolvido)
					if Trim("" & v_nfe_item_devolvido(i).descricao) = "**-#-EDITADO-#-**" then
						s = "SELECT " & _
								"*" & _
							" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
							" WHERE" & _
								" (id = '" & v_nfe_item_devolvido(i).id & "')"
						if rs.State <> 0 then rs.Close
						rs.Open s, cn
						if Err <> 0 then
							alerta=texto_add_br(alerta)
							alerta=alerta & Cstr(Err) & ": " & Err.Description
						elseif rs.EOF then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Registro da devolu��o de mercadorias (id=" & v_nfe_item_devolvido(i).id & ") n�o foi encontrado."
						else
							log_via_vetor_carrega_do_recordset rs, vLogItemDevolvido1, campos_a_omitir_ItemDevolvido
                            if v_nfe_item_devolvido(i).id_nfe_emitente <> "" And v_nfe_item_devolvido(i).NFe_serie_NF <> "" And v_nfe_item_devolvido(i).NFe_numero_NF <> "" then
							    rs("id_nfe_emitente") = CInt(v_nfe_item_devolvido(i).id_nfe_emitente)
							    rs("NFe_serie_NF") = CLng(v_nfe_item_devolvido(i).NFe_serie_NF)
							    rs("NFe_numero_NF") = CLng(v_nfe_item_devolvido(i).NFe_numero_NF)
                            else
                                rs("id_nfe_emitente") = 0
							    rs("NFe_serie_NF") = 0
							    rs("NFe_numero_NF") = 0
                            end if
							rs("dt_hr_anotacao_numero_NF") = Now
							rs("usuario_anotacao_numero_NF") = usuario
							rs.Update
							log_via_vetor_carrega_do_recordset rs, vLogItemDevolvido2, campos_a_omitir_ItemDevolvido
							s = log_via_vetor_monta_alteracao(vLogItemDevolvido1, vLogItemDevolvido2)
							if s <> "" then
								if s_log_ItemDevolvido <> "" then s_log_ItemDevolvido=s_log_ItemDevolvido & "; "
								s_log_ItemDevolvido=s_log_ItemDevolvido & "Devolu��o ID=" & v_nfe_item_devolvido(i).id & ", produto (" & v_nfe_item_devolvido(i).fabricante & ")" & v_nfe_item_devolvido(i).produto & ": " & s
								end if
							end if
						end if
					next
				end if
			end if
        
        ' ATUALIZA TABELA DE FRETES
        s_log_Frete_Exclui = ""
		if alerta = "" then        
            if blnValorFreteEdicaoLiberada then
                    intQtdeFretes = Request.Form("c_valor_frete").Count
                    for i = 1 to intQtdeFretes - 1
                        blnEditouFrete = False
                        ckb_frete_exclui = Request.Form("ckb_exclui_frete_" & i)
                        frete_id = Trim(Request.Form("frete_id")(i))
                        c_valor_frete = Trim(Request.Form("c_valor_frete")(i + 1))
                        frete_tipo = Request.Form("c_tipo_frete")(i)
                        frete_transportadora_id = Trim(Request.Form("c_frete_transportadora_id")(i))
                        c_frete_serie_NF = Trim(Request.Form("c_frete_serie_NF")(i)) 
                        c_frete_numero_NF = Trim(Request.Form("c_frete_numero_NF")(i)) 
                        c_frete_emitente = Trim(Request.Form("c_frete_emitente")(i))
                        if c_frete_emitente = "" then c_frete_emitente = "0"
                        s = "SELECT * FROM t_PEDIDO_FRETE WHERE (id = '" & frete_id & "')"
		                if frete_id <> "" then
				            if rs.State <> 0 then rs.Close
                            rs.Open s, cn
                            if CCur(c_valor_frete) <> CCur(Trim("" & rs("vl_frete"))) then
                                s_log_Frete_Altera = s_log_Frete_Altera & "frete_valor: " & formata_moeda(Trim("" & rs("vl_frete"))) & " => " & c_valor_frete & "; "
                                blnEditouFrete = True
                            end if
                            if frete_tipo <> Trim("" & rs("codigo_tipo_frete")) then
                                s_log_Frete_Altera = s_log_Frete_Altera & "tipo_frete: " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE,Trim("" & rs("codigo_tipo_frete"))) & " => " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE,frete_tipo) & "; "
                                blnEditouFrete = True   
                            end if                     
                            if frete_transportadora_id <> Trim("" & rs("transportadora_id")) then
                                s_log_Frete_Altera = s_log_Frete_Altera & "transportadora_id: " & Trim("" & rs("transportadora_id")) & " => " & frete_transportadora_id & "; "
                                s = "SELECT * FROM t_TRANSPORTADORA WHERE (id='" & frete_transportadora_id & "')"
                                set rs2 = cn.Execute(s)
                                transportadora_cnpj = rs2("cnpj")
                                blnEditouFrete = True 
                            end if
                            if c_frete_emitente <> Trim("" & rs("id_nfe_emitente")) then 
                                s_log_Frete_Altera = s_log_Frete_Altera & "id_nfe_emitente: " & Trim("" & rs("id_nfe_emitente")) & " => " & c_frete_emitente & "; "
                                blnEditouFrete = True 
                            end if
                            if c_frete_numero_NF <> Trim("" & rs("numero_NF")) then
                                s_log_Frete_Altera = s_log_Frete_Altera & "numero_NF: " & Trim("" & rs("numero_NF")) & " => " & c_frete_numero_NF & "; "
                                blnEditouFrete = True 
                            end if
                            if c_frete_serie_NF <> Trim("" & rs("serie_NF")) then
                                s_log_Frete_Altera = s_log_Frete_Altera & "serie_NF: " & Trim("" & rs("serie_NF")) & " => " & c_frete_serie_NF & "; "
                                blnEditouFrete = True 
                            end if
                            

                            if frete_tipo <> "" then rs("codigo_tipo_frete") = frete_tipo
                            rs("vl_frete") = converte_numero(c_valor_frete)
                            if frete_transportadora_id <> "" then rs("transportadora_id") = frete_transportadora_id
                            if c_frete_serie_NF <> "" then rs("serie_NF") = c_frete_serie_NF
                            if c_frete_numero_NF <> "" then rs("numero_NF") = c_frete_numero_NF
                            if c_frete_emitente <> "" then rs("id_nfe_emitente") = c_frete_emitente
                            if transportadora_cnpj <> "" then rs("transportadora_cnpj") = transportadora_cnpj
                            if blnEditouFrete then
                                rs("dt_ult_atualizacao") = Now
                                rs("dt_hr_ult_atualizacao") = Now
                                rs("usuario_ult_atualizacao") = usuario
                            end if
                            rs.Update
                            
                            if ckb_frete_exclui <> "" then
                                s = "DELETE FROM t_PEDIDO_FRETE WHERE (id = '" & frete_id & "')"
			                    cn.Execute(s)
			                    If Err = 0 then
                                        if s_log_Frete_Exclui <> "" then s_log_Frete_Exclui = s_log_Frete_Exclui & chr(13) 
				                        s_log_Frete_Exclui = s_log_Frete_Exclui & "ID_frete: " & rs("id") & "; data_cadastro: " & rs("dt_cadastro") & "; valor_frete: " & c_valor_frete & "; transportadora_id: " & frete_transportadora_id & "; tipo_frete: " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE,frete_tipo) & _
                                                             "; id_nfe_emitente: " & c_frete_emitente & "; numero_NF: " & c_frete_numero_NF & "; serie_NF: " & c_frete_serie_NF
                               else
				                    alerta = "FALHA AO REMOVER O FRETE (" & Cstr(Err) & ": " & Err.Description & ")."
				                end if                        
                            end if
			            end if
                        if s_log_Frete_Altera <> "" then
                            s_log_Frete_Altera = "ID_frete: " & rs("id") & "; data_cadastro: " & rs("dt_cadastro") & "; " & s_log_Frete_Altera
                            grava_log usuario, "", pedido_selecionado, "", "PED EDITA FRETE", s_log_Frete_Altera
                            s_log_Frete_Altera = ""
                        end if
		            next
			    end if
        end if
		
		'Sincroniza o campo 'marketplace_codigo_origem' dos pedidos-filhote, se existirem
		if alerta = "" then
			if blnMarketplaceCodigoOrigemAlterado then
				s = "UPDATE t_PED__FILHOTE" & _
					" SET" & _
						" t_PED__FILHOTE.marketplace_codigo_origem = t_PED__BASE.marketplace_codigo_origem" & _
					" FROM t_PEDIDO AS t_PED__FILHOTE" & _
						" INNER JOIN t_PEDIDO AS t_PED__BASE ON (t_PED__FILHOTE.pedido_base = t_PED__BASE.pedido)" & _
					" WHERE" & _
						" (t_PED__FILHOTE.pedido_base = '" & retorna_num_pedido_base(pedido_selecionado) & "')" & _
						" AND (t_PED__FILHOTE.pedido <> t_PED__FILHOTE.pedido_base)"
				cn.Execute(s)
				If Err <> 0 then
					alerta = "FALHA AO SINCRONIZAR O CAMPO 'marketplace_codigo_origem' (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if
			end if

		if alerta = "" then
			if blnEnviarEmailVendedorStatusAnaliseCredito then
				call le_usuario(r_pedido.vendedor, r_vendedor, msg_erro)
				if Trim("" & r_vendedor.email) = "" then
					if s_log <> "" then s_log = s_log & "; "
					s_log = s_log & "Vendedor n�o possui e-mail cadastrado para ser notificado sobre altera��o do status da an�lise de cr�dito"
				else
					corpo_mensagem = "Pedido " & pedido_base & " foi alterado para '" & x_analise_credito(s_analise_credito) & "' �s " & formata_data_hora_sem_seg(Now) & _
									vbCrLf & _
									"Motivo: " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__AC_PENDENTE_VENDAS_MOTIVO, s_ac_pendente_vendas_motivo) & _
									vbCrLf & vbCrLf & _
									String(60, "-") & _
									vbCrLf & _
									"Aten��o: esta � uma mensagem autom�tica, N�O responda a este e-mail!"
					if EmailSndSvcGravaMensagemParaEnvio(emailSndSvcRemetenteMensagemSistema, _
													"", _
													r_vendedor.email, _
													"", _
													"", _
													"Pedido " & pedido_base & " alterado para '" & x_analise_credito(s_analise_credito) & "'", _
													corpo_mensagem, _
													Now, _
													id_email, _
													msg_erro_grava_email) then
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & "Enviado e-mail para o vendedor avisando sobre altera��o do status da an�lise de cr�dito: " & r_vendedor.email & " (mensagem id: " & id_email & ")"
					else
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & "Falha ao tentar enviar e-mail para o vendedor avisando sobre altera��o do status da an�lise de cr�dito: " & msg_erro_grava_email
						end if
					end if
				end if
			end if

		if alerta = "" then
			if blnHaPedidoAprovadoComEntregaPendente then
				''Envia alerta de que houve edi��o no cadastro de cliente que possui pedido com status de an�lise de cr�dito 'cr�dito ok' e com entrega pendente
				corpo_mensagem = "" 
				id_email = ""
				msg_erro_grava_email = ""
				set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoCadastroClienteComPedidoCreditoOkEntregaPendente)
				if Trim("" & rEmailDestinatario.campo_texto) <> "" then
					
					corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " no m�dulo Central o endere�o do cliente:" & _
									vbCrLf & _
										cnpj_cpf_formata(r_cliente.cnpj_cpf) & " - " & r_cliente.nome  & _
										vbCrLf & _
										"A altera��o foi realizada para o pedido: '" & r_pedido.pedido & "' que possui o status de an�lise de cr�dito 'Cr�dito OK' e com entrega pendente." & _
										vbCrLf & _
									
										"Informa��es detalhadas sobre as altera��es:" & vbCrLf & _
                                        substitui_caracteres(sLogEmail, ";", vbCrLf)
									
										EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
																		"", _
																		rEmailDestinatario.campo_texto, _
																		"", _
																		"", _
																		"Edi��o no endere�o de cliente que possui pedido com status 'Cr�dito OK' e entrega pendente", _
																		corpo_mensagem, _
																		Now, _
																		id_email, _
																		msg_erro_grava_email
					'alerta = msg_erro_grava_email   & " | " & id_email  & " | " & getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA)
				end if
			end if
		end if

		if alerta = "" then
			if (s_log <> "") And (s_log_FP <> "") then s_log = s_log & "; "
			s_log = s_log & s_log_FP
			
        if s_log_Frete_Exclui <> "" then grava_log usuario, "", pedido_selecionado, "", "PED EXCLUI FRETE", s_log_Frete_Exclui
            
		'	CUSTO FINANCEIRO POR FORNECEDOR
			if (s_log <> "") And (s_log_PedCFF <> "") then s_log = s_log & "; "
			s_log = s_log & s_log_PedCFF
			if (s_log <> "") And (s_log_ItemCFF <> "") then s_log = s_log & "; "
			s_log = s_log & s_log_ItemCFF
			
		'	DEVOLU��O DE MERCADORIAS
			if (s_log <> "") And (s_log_ItemDevolvido <> "") then s_log = s_log & "; "
			s_log = s_log & s_log_ItemDevolvido
			
		'	GRAVA O LOG!!
			if s_log <> "" then grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_ALTERACAO, s_log
		
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("pedido.asp?pedido_selecionado=" & pedido_selecionado & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))) & "&url_origem=" & url_origem
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
	<title>CENTRAL</title>
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
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<% if s_tabela_municipios_IBGE <> "" then %>
	<br /><br />
	<%=s_tabela_municipios_IBGE%>
<% end if %>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>
<% end if %>

</html>


<%
	if Not blnErroConsistencia then
		if Not rs is nothing then
			if rs.State <> 0 then rs.Close
			set rs = nothing
			end if
		
	'	FECHA CONEXAO COM O BANCO DE DADOS
		cn.Close
		set cn = nothing
		end if
%>