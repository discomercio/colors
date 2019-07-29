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

	class cl_ITEM_PEDIDO_EDICAO
		dim pedido
		dim fabricante
		dim produto
		dim qtde
		dim desc_dado
		dim preco_venda
		dim preco_venda_original
		dim preco_NF
		dim preco_fabricante
		dim preco_lista
		dim margem
		dim desc_max
		dim comissao
		dim descricao
		dim descricao_html
		dim ean
		dim grupo
		dim peso
		dim qtde_volumes
		dim abaixo_min_status
		dim abaixo_min_autorizacao
		dim abaixo_min_autorizador
		dim sequencia
		dim markup_fabricante
		dim abaixo_min_superv_autorizador
		dim vl_custo2
		dim custoFinancFornecCoeficiente
		dim custoFinancFornecPrecoListaBase
		dim cubagem
		dim ncm
		dim cst
		end class

	dim s, usuario, loja, pedido_selecionado, pedido_base, tipo_cliente
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s

    dim url_origem
    url_origem = Trim(Request("url_origem"))

	dim cliente_selecionado
	cliente_selecionado = Trim(Request.Form("cliente_selecionado"))

	dim alerta, blnErroConsistencia
	alerta=""
	blnErroConsistencia=False

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim r_cliente
	set r_cliente = New cl_CLIENTE
	call x_cliente_bd(cliente_selecionado, r_cliente)
	tipo_cliente = r_cliente.tipo
	
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

	dim r_pedido, v_item_bd
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
	else
		if Not le_pedido_item(pedido_selecionado, v_item_bd, msg_erro) then alerta = msg_erro
		end if

	dim blnTemRA
	blnTemRA = False
	if alerta = "" then
		for i=Lbound(v_item_bd) to Ubound(v_item_bd)
			if Trim("" & v_item_bd(i).produto) <> "" then
				if v_item_bd(i).preco_NF <> v_item_bd(i).preco_venda then
					blnTemRA = True
					exit for
					end if
				end if
			next
		end if
	
	dim blnProcessaSelecaoAutoTransp
	dim blnNFEmitida
	blnNFEmitida = False
	if Trim("" & r_pedido.obs_2) <> "" then blnNFEmitida = True
	
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

    dim blnNumPedidoECommerceEdicaoLiberada
    s = Trim(Request.Form("blnNumPedidoECommerceEdicaoLiberada"))
    blnNumPedidoECommerceEdicaoLiberada = CBool(s)

	dim blnObs1EdicaoLiberada
	s = Trim(Request.Form("blnObs1EdicaoLiberada"))
	blnObs1EdicaoLiberada = CBool(s)

	dim blnFormaPagtoEdicaoLiberada
	s = Trim(Request.Form("blnFormaPagtoEdicaoLiberada"))
	blnFormaPagtoEdicaoLiberada = CBool(s)

	dim blnFormaPagtoEditada
	s = Trim(Request.Form("blnFormaPagtoEditada"))
	blnFormaPagtoEditada = CBool(s)
	
	dim bln_RA_EdicaoLiberada
	s = Trim(Request.Form("bln_RA_EdicaoLiberada"))
	bln_RA_EdicaoLiberada = CBool(s)
	
	dim bln_RT_EdicaoLiberada
	s = Trim(Request.Form("bln_RT_EdicaoLiberada"))
	bln_RT_EdicaoLiberada = CBool(s)
	
	dim blnItemPedidoEdicaoLiberada
	s = Trim(Request.Form("blnItemPedidoEdicaoLiberada"))
	blnItemPedidoEdicaoLiberada = CBool(s)

	dim blnEtgImediataEdicaoLiberada
	s = Trim(Request.Form("blnEtgImediataEdicaoLiberada"))
	blnEtgImediataEdicaoLiberada = CBool(s)
	
	dim blnAnaliseCreditoEdicaoLiberada
	s = Trim(Request.Form("blnAnaliseCreditoEdicaoLiberada"))
	blnAnaliseCreditoEdicaoLiberada = CBool(s)

	dim blnBemUsoConsumoEdicaoLiberada
	s = Trim(Request.Form("blnBemUsoConsumoEdicaoLiberada"))
	blnBemUsoConsumoEdicaoLiberada = CBool(s)

	dim rb_garantia_indicador, GarantiaIndicadorStatusOriginal
	dim blnGarantiaIndicadorEdicaoLiberada
	GarantiaIndicadorStatusOriginal = Trim(Request.Form("GarantiaIndicadorStatusOriginal"))
	rb_garantia_indicador = Trim(Request.Form("rb_garantia_indicador"))
	s = Trim(Request.Form("blnGarantiaIndicadorEdicaoLiberada"))
	blnGarantiaIndicadorEdicaoLiberada = CBool(s)

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
	
	dim s_qtde_parcelas, s_forma_pagto, s_obs1, s_obs2, s_obs3, s_ped_bonshop, s_indicador, s_pedido_ac, s_pedido_mktplace, s_pedido_origem
	dim s_analise_credito, s_analise_credito_a, s_nf_texto, s_num_pedido_compra
	dim s_etg_imediata, s_bem_uso_consumo
	dim blnUpdate
	s_obs1=Trim(request("c_obs1"))
	s_obs2=Trim(request("c_obs2"))
	s_obs3=Trim(request("c_obs3"))
	s_ped_bonshop=Trim(request("pedBonshop"))
	s_analise_credito=Trim(request("rb_analise_credito"))
	s_etg_imediata=Trim(request("rb_etg_imediata"))
	s_bem_uso_consumo=Trim(request("rb_bem_uso_consumo"))
	s_forma_pagto=Trim(request("c_forma_pagto"))
    s_indicador = Trim(Request("c_indicador"))
    s_pedido_ac = Trim(request("c_pedido_ac"))
    s_pedido_mktplace = Trim(Request("c_numero_mktplace"))
    s_pedido_origem = Trim(Request("c_origem_pedido"))
    s_nf_texto = Trim(Request("c_nf_texto"))
    s_num_pedido_compra = Trim(Request("c_num_pedido_compra"))

	if s_pedido_mktplace = "" then s_pedido_origem = ""

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

	if (r_pedido.loja = NUMERO_LOJA_BONSHOP) And (r_pedido.plataforma_origem_pedido = COD_PLATAFORMA_ORIGEM_PEDIDO__MAGENTO) then
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

	dim perc_RT
	perc_RT = converte_numero(s_perc_RT)

    if c_loja <> NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
        if blnIndicadorEdicaoLiberada then
            if s_indicador = "" then
                if s_perc_RT <> "" then
                    if converte_numero(s_perc_RT) > 0 then
                        alerta = "N�o � poss�vel gravar o pedido com o campo ""Indicador"" vazio e ""COM(%)"" maior do que zero."
                    end if
                end if
            end if
        end if
    end if
	
	dim c_gravar_perc_RT_novo
	c_gravar_perc_RT_novo = Trim(Request("c_gravar_perc_RT_novo"))
	
	dim c_consiste_perc_max_comissao_e_desconto
	c_consiste_perc_max_comissao_e_desconto = Trim(Request("c_consiste_perc_max_comissao_e_desconto"))
	
	dim blnEndEntregaEdicaoLiberada, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs,blnEndEtg_obs
    blnEndEtg_obs = false
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
	
	dim v_item, i, n, k, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_totalFamiliaPrecoNFLiquido, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, id_pedido_base
	redim v_item(0)
	set v_item(Ubound(v_item)) = New cl_ITEM_PEDIDO_EDICAO
	v_item(Ubound(v_item)).produto = ""
	n = Request.Form("c_produto").Count
	for i = 1 to n
		s=Trim(Request.Form("c_produto")(i))
		if s <> "" then
			if Trim(v_item(ubound(v_item)).produto) <> "" then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ITEM_PEDIDO_EDICAO
				end if
			with v_item(ubound(v_item))
				.produto=Ucase(Trim(Request.Form("c_produto")(i)))
				s=retorna_so_digitos(Request.Form("c_fabricante")(i))
				.fabricante=normaliza_codigo(s, TAM_MIN_FABRICANTE)
				s=Trim(Request.Form("c_vl_unitario")(i))
				.preco_venda=converte_numero(s)
				s=Trim(Request.Form("c_vl_unitario_original")(i))
				.preco_venda_original=converte_numero(s)
				if (r_pedido.permite_RA_status = 1) Or (blnTemRA And (r_pedido.st_violado_permite_RA_status = 1)) then
					s=Trim(Request.Form("c_vl_NF")(i))
					.preco_NF=converte_numero(s)
				else
					.preco_NF = .preco_venda
					end if
				s=Trim(Request.Form("c_preco_lista")(i))
				.preco_lista=converte_numero(s)
				end with
			end if
		next

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
	
'	CALCULA O VALOR TOTAL DO PEDIDO
	dim vl_total
	if alerta = "" then
		vl_total = 0
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .produto <> "" then 
					vl_total = vl_total + (.qtde * .preco_venda)
					end if
				end with
			next
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
	
	'Edit�vel?
	if blnEndEntregaEdicaoLiberada then
		if alerta = "" then
            if (EndEtg_endereco<>r_pedido.EndEtg_endereco) Or (EndEtg_bairro<>r_pedido.EndEtg_bairro) Or (EndEtg_cidade<>r_pedido.EndEtg_cidade) Or (EndEtg_uf<>r_pedido.EndEtg_uf) Or (EndEtg_cep<>r_pedido.EndEtg_cep) Or (EndEtg_obs<>r_pedido.EndEtg_cod_justificativa) then
                  blnEndEtg_obs = true 
            end if
			if (EndEtg_endereco<>"") Or (EndEtg_bairro<>"") Or (EndEtg_cidade<>"") Or (EndEtg_uf<>"") Or (EndEtg_cep<>"") Or (EndEtg_obs<>"") then
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
			end if
		end if
	
	dim v_desconto()
	ReDim v_desconto(0)
	v_desconto(UBound(v_desconto)) = ""

	dim desc_dado_arredondado
	
	if alerta="" then
		if c_consiste_perc_max_comissao_e_desconto = "S" then
			for i=Lbound(v_item) to Ubound(v_item)
				with v_item(i)
					if (.preco_venda <> .preco_venda_original) Or blnFormaPagtoEditada then
						if .preco_lista = 0 then 
							.desc_dado = 0
							desc_dado_arredondado = 0
						else
							.desc_dado = 100*(.preco_lista-.preco_venda)/.preco_lista
							desc_dado_arredondado = converte_numero(formata_perc_desc(.desc_dado))
							end if
						
						if desc_dado_arredondado > perc_comissao_e_desconto_a_utilizar then
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
					end with
				next
			end if
		end if
	
	
	if alerta <> "" then blnErroConsistencia=True
	
	
'	GRAVA NO BANCO DE DADOS
'	=======================
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
	dim vLog1()
	dim vLog2()
	dim s_log, s_log_manual
	dim campos_a_omitir
	s_log = ""
	s_log_manual = ""
	campos_a_omitir = "|etg_imediata_data|etg_imediata_usuario|"
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
							rs("analise_credito")=CLng(s_analise_credito)
							rs("analise_credito_data")=Now
							rs("analise_credito_usuario")=usuario
							end if
						end if
					end if
					
			'	Forma de Pagamento (nova vers�o)
				if (versao_forma_pagamento = "2") And blnFormaPagtoEdicaoLiberada then
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
					end if
				
				if bln_RT_EdicaoLiberada Or (c_gravar_perc_RT_novo = "S") then rs("perc_RT") = converte_numero(s_perc_RT)

				if blnIndicadorEdicaoLiberada then rs("indicador") = s_indicador
				
				if blnNumPedidoECommerceEdicaoLiberada then rs("pedido_bs_x_ac")=s_pedido_ac

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
						
						rs("marketplace_codigo_origem")=s_pedido_origem
						end if
					end if

				rs("pedido_bs_x_at")=s_ped_bonshop

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
								rs("analise_credito")=CLng(s_analise_credito)
								rs("analise_credito_data")=Now
								rs("analise_credito_usuario")=usuario
								end if
							end if
						end if
					
				'	Forma de Pagamento (nova vers�o)
					if (versao_forma_pagamento = "2") And blnFormaPagtoEdicaoLiberada then
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
						end if
					end if
				
				if blnObs1EdicaoLiberada then
                     rs("obs_1") = s_obs1
                     rs("NFe_texto_constar") = s_nf_texto
                     rs("NFe_xPed") = s_num_pedido_compra
                end if
				
				if blnFormaPagtoEdicaoLiberada then rs("forma_pagto") = s_forma_pagto

				if bln_RT_EdicaoLiberada Or (c_gravar_perc_RT_novo = "S") then rs("perc_RT") = converte_numero(s_perc_RT)

				if blnIndicadorEdicaoLiberada then rs("indicador") = s_indicador
				
				if (versao_forma_pagamento = "1") And blnFormaPagtoEdicaoLiberada then
					if IsNumeric(s_qtde_parcelas) then 
						rs("qtde_parcelas") = CLng(s_qtde_parcelas)
					else
						rs("qtde_parcelas") = 0
						end if
					end if
				
				if blnNumPedidoECommerceEdicaoLiberada then rs("pedido_bs_x_ac")=s_pedido_ac

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
						
						rs("marketplace_codigo_origem")=s_pedido_origem
						end if
					end if

				'Guardar informa��es de endere�o presentes no pedido (consist�ncia para verificar se mudou de endere�o)
				dim st_end_entrega_anterior, EndEtg_cep_anterior
				st_end_entrega_anterior = rs("st_end_entrega")
				EndEtg_cep_anterior = rs("EndEtg_cep")
				
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
					end if
				
			'	SELE��O AUTOM�TICA DA TRANSPORTADORA COM BASE NO CEP
				blnProcessaSelecaoAutoTransp = False
				if (Not IsPedidoEncerrado(r_pedido.st_entrega)) And _
					(Not blnNFEmitida) And (Trim("" & rs("analise_credito")) <> Trim("" & COD_AN_CREDITO_OK)) then
					if CInt(st_end_entrega_anterior) <> CInt(rs("st_end_entrega")) then
					'	HOUVE ALTERA��O ENTRE USAR O ENDERE�O DE ENTREGA E O ENDERE�O DO CADASTRO (OU VICE-VERSA)
						blnProcessaSelecaoAutoTransp = True
					else
						if CInt(rs("st_end_entrega")) <> 0 then
						'	OBS: ALTERA��ES NO ENDERE�O DO CADASTRO S�O PROCESSADAS NA P�GINA CLIENTEATUALIZA.ASP
						'	H� ENDERE�O DE ENTREGA: O CEP MUDOU?
							if Trim("" & EndEtg_cep_anterior) <> Trim("" & rs("EndEtg_cep")) then blnProcessaSelecaoAutoTransp = True
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
						if r_cliente.cep <> "" then sTranspSelAutoTransportadoraId = obtem_transportadora_pelo_cep(retorna_so_digitos(r_cliente.cep))
						sTranspSelAutoCep = retorna_so_digitos(r_cliente.cep)
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
				
				
				if blnEtgImediataEdicaoLiberada then
					if s_etg_imediata <> "" then 
						if CLng(rs("st_etg_imediata")) <> CLng(s_etg_imediata) then
							rs("st_etg_imediata")=CLng(s_etg_imediata)
							rs("etg_imediata_data")=Now
							rs("etg_imediata_usuario")=usuario
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

				if blnGarantiaIndicadorEdicaoLiberada then
					if CLng(rs("GarantiaIndicadorStatus")) <> CLng(rb_garantia_indicador) then
						rs("GarantiaIndicadorStatus")=CLng(rb_garantia_indicador)
						rs("GarantiaIndicadorUsuarioUltAtualiz")=usuario
						rs("GarantiaIndicadorDtHrUltAtualiz")=Now
						end if
					end if
				rs("pedido_bs_x_at")=s_ped_bonshop
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
							if bln_RA_EdicaoLiberada then
								if rs("preco_NF") <> .preco_NF then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "pre�o de NF do produto " & Trim(.produto) & " (" & Trim(.fabricante) & "): " & formata_moeda(rs("preco_NF")) & " => " & formata_moeda(.preco_NF)
									rs("preco_NF")=.preco_NF
									blnUpdate = True
									end if
								end if
							if converte_numero(.abaixo_min_status) <> 0 then
								blnUpdate = True
								rs("abaixo_min_status")=.abaixo_min_status
								rs("abaixo_min_autorizacao")=.abaixo_min_autorizacao
								rs("abaixo_min_autorizador")=.abaixo_min_autorizador
								rs("abaixo_min_superv_autorizador")=.abaixo_min_superv_autorizador
								
								if s_log <> "" then s_log = s_log & "; "
								s_log = s_log & _
										"Senha de desconto p/ produto (" & .fabricante & ")" & .produto & ": " & _
										"abaixo_min_status=" & formata_texto_log(.abaixo_min_status) & "; " & _
										"abaixo_min_autorizacao=" & formata_texto_log(.abaixo_min_autorizacao) & "; " & _
										"abaixo_min_autorizador=" & formata_texto_log(.abaixo_min_autorizador) & "; " & _
										"abaixo_min_superv_autorizador=" & formata_texto_log(.abaixo_min_superv_autorizador)
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
						rs("vendedor") = usuario
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
		
		if alerta = "" then
			if (s_log <> "") And (s_log_FP <> "") then s_log = s_log & "; "
			s_log = s_log & s_log_FP
			
		'	CUSTO FINANCEIRO POR FORNECEDOR
			if (s_log <> "") And (s_log_PedCFF <> "") then s_log = s_log & "; "
			s_log = s_log & s_log_PedCFF
			if (s_log <> "") And (s_log_ItemCFF <> "") then s_log = s_log & "; "
			s_log = s_log & s_log_ItemCFF
			
		'	GRAVA O LOG!!
			if s_log <> "" then grava_log usuario, loja, pedido_selecionado, "", OP_LOG_PEDIDO_ALTERACAO, s_log
		
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
<% if s_tabela_municipios_IBGE <> "" then %>
	<br /><br />
	<%=s_tabela_municipios_IBGE%>
<% end if %>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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