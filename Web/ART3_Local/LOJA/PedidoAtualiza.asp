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
		dim preco_lista_original
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
		dim StatusDescontoSuperior
		dim IdUsuarioDescontoSuperior
		dim DataHoraDescontoSuperior
		end class

	dim s, usuario, loja, pedido_selecionado, pedido_base, tipo_cliente
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s

    dim url_origem
    url_origem = Trim(Request("url_origem"))

	dim cliente_selecionado
	cliente_selecionado = Trim(Request.Form("cliente_selecionado"))

	dim i, j, n, k
	dim blnAchou
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
	
	dim eh_cpf
	eh_cpf=(len(r_cliente.cnpj_cpf)=11)

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim sBlocoNotasEndCob, sBlocoNotasEndEtg, sBlocoNotasMsg, sEnderecoOriginal, sEnderecoNovo
	sBlocoNotasEndCob = ""
	sBlocoNotasEndEtg = ""
	sBlocoNotasMsg = ""

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

	dim r_pedido, r_pedido_atualizado, v_item_bd
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
	else
		if Not le_pedido_item(pedido_selecionado, v_item_bd, msg_erro) then alerta = msg_erro
		end if

	dim r_loja
	set r_loja = New cl_LOJA
	call x_loja_bd(r_pedido.loja, r_loja)

	dim r_usuario
	if alerta = "" then
		call le_usuario(usuario, r_usuario, msg_erro)
		end if

	dim blnUsuarioDeptoFinanceiro, vDeptoSetorUsuario
	blnUsuarioDeptoFinanceiro = False
	
	if alerta = "" then
		if Not obtem_Usuario_x_DeptoSetor(usuario, vDeptoSetorUsuario, msg_erro) then
			alerta=texto_add_br(alerta)
			alerta = alerta & msg_erro
		else
			for i=LBound(vDeptoSetorUsuario) to UBound(vDeptoSetorUsuario)
				if (vDeptoSetorUsuario(i).StInativo = 0) then
					if (vDeptoSetorUsuario(i).Id = ID_DEPTO_SETOR__FIN_FINANCEIRO) Or (vDeptoSetorUsuario(i).Id = ID_DEPTO_SETOR__FIN_CREDITO) then
						blnUsuarioDeptoFinanceiro = True
						exit for
						end if
					end if
				next
			end if
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
	dim s_perc_RT, perc_RT, s_perc_RT_original, vl_total_RA, vl_total_RA_liquido
	dim idMeioPagtoMonitorado, sMeioPagtoMonitoradoIdentificado
	dim vMeioPagtoMonitorado, iMeioPagtoMonitorado

	versao_forma_pagamento = Trim(Request.Form("versao_forma_pagamento"))
	vlTotalFormaPagto = 0
	
	dim blnEditouIndicador
	blnEditouIndicador = False
    
	dim blnIndicadorEdicaoLiberada
    s = Trim(Request.Form("blnIndicadorEdicaoLiberada"))
    blnIndicadorEdicaoLiberada = CBool(s)

    dim blnNumPedidoECommerceEdicaoLiberada
    s = Trim(Request.Form("blnNumPedidoECommerceEdicaoLiberada"))
    blnNumPedidoECommerceEdicaoLiberada = CBool(s)

	dim blnObs1EdicaoLiberada
	s = Trim(Request.Form("blnObs1EdicaoLiberada"))
	blnObs1EdicaoLiberada = CBool(s)

	dim nivelEdicaoFormaPagto
	s = Trim(Request.Form("nivelEdicaoFormaPagto"))
	nivelEdicaoFormaPagto = CLng(s)

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
	
	dim blnMarketplaceCodigoOrigemAlterado
	blnMarketplaceCodigoOrigemAlterado = False

	dim s_qtde_parcelas, s_forma_pagto, s_obs1, s_obs2, s_obs3, s_ped_bonshop, s_indicador, s_pedido_ac, s_pedido_mktplace, s_pedido_origem
	dim s_analise_credito, s_analise_credito_a, s_nf_texto, s_num_pedido_compra
	dim s_etg_imediata, s_bem_uso_consumo, s_etg_imediata_original, c_data_previsao_entrega
	dim blnUpdate
	s_obs1=Trim(request("c_obs1"))
	s_obs2=Trim(request("c_obs2"))
	s_obs3=Trim(request("c_obs3"))
	s_ped_bonshop=Trim(request("pedBonshop"))
	s_analise_credito=Trim(request("rb_analise_credito"))
	s_etg_imediata=Trim(request("rb_etg_imediata"))
	c_data_previsao_entrega = Trim(Request("c_data_previsao_entrega"))
	s_bem_uso_consumo=Trim(request("rb_bem_uso_consumo"))
	s_forma_pagto=Trim(request("c_forma_pagto"))
    s_indicador = Trim(Request("c_indicador"))
    s_pedido_ac = Trim(request("c_pedido_ac"))
    s_pedido_mktplace = Trim(Request("c_numero_mktplace"))
    s_pedido_origem = Trim(Request("c_origem_pedido"))
    s_nf_texto = Trim(Request("c_nf_texto"))
    s_num_pedido_compra = Trim(Request("c_num_pedido_compra"))

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
		if (s_pedido_ac <> "") And (r_loja.magento_api_versao <> VERSAO_API_MAGENTO_V2_REST_JSON) then
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
	s_perc_RT_original = Trim(request("c_perc_RT_original"))
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

	'Confere se o n�vel de permiss�o de edi��o da forma de pagamento obtido pelo form est� correto
	dim nivelEdicaoFormaPagtoConferencia
	nivelEdicaoFormaPagtoConferencia = COD_NIVEL_EDICAO_BLOQUEADA
	if operacao_permitida(OP_LJA_EDITA_PEDIDO_FORMA_PAGTO, s_lista_operacoes_permitidas) Or operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
		nivelEdicaoFormaPagtoConferencia = COD_NIVEL_EDICAO_LIBERADA_TOTAL

		' Analisa situa��es que liberam apenas parcialmente a edi��o da forma de pagamento, ou seja,
		' pode-se alterar os valores da forma de pagamento atualmente selecionada, mas n�o se pode
		' alterar a forma de pagamento e nem os meios de pagamento (ex: de '� Vista' para 
		' 'Parcelado com Entrada' ou de 'Dep�sito' para 'Boleto').
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'Se o status da an�lise de cr�dito est� em uma situa��o que demanda uma confirma��o manual do depto de an�lise de cr�dito, bloqueia
		'a edi��o da forma de pagamento para n�o haver o risco de uma altera��o ser feita sem o conhecimento do depto de an�lise de cr�dito.
		'Qualquer altera��o necess�ria na forma de pagamento deve ser solicitada ao depto de an�lise de cr�dito.
		if (nivelEdicaoFormaPagtoConferencia > COD_NIVEL_EDICAO_LIBERADA_PARCIAL) _
			AND _
			(Cstr(r_pedido.loja) <> Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE)) _
			AND _
			( _
				(Trim("" & r_pedido.analise_credito) = Cstr(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) _
				OR (Trim("" & r_pedido.analise_credito) = Cstr(COD_AN_CREDITO_OK)) _
			) then
			nivelEdicaoFormaPagtoConferencia = COD_NIVEL_EDICAO_LIBERADA_PARCIAL
			end if
		
		' Analisa situa��es em que a edi��o da forma de pagamento deve ser bloqueada totalmente
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if Trim("" & r_pedido.st_entrega) = ST_ENTREGA_ENTREGUE then
			nivelEdicaoFormaPagtoConferencia = COD_NIVEL_EDICAO_BLOQUEADA
			end if

		if Trim("" & r_pedido.st_entrega) = ST_ENTREGA_CANCELADO then
			nivelEdicaoFormaPagtoConferencia = COD_NIVEL_EDICAO_BLOQUEADA
			end if
		end if 'if operacao_permitida(OP_LJA_EDITA_PEDIDO_FORMA_PAGTO, s_lista_operacoes_permitidas) Or operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas)

	if Cstr(nivelEdicaoFormaPagto) <> Cstr(nivelEdicaoFormaPagtoConferencia) then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Foi encontrada uma inconsist�ncia na verifica��o do n�vel de permiss�o de edi��o da forma de pagamento (" & Cstr(nivelEdicaoFormaPagto) & " <> " & Cstr(nivelEdicaoFormaPagtoConferencia) & ")"
		end if


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

	dim blnHouvePrecoVendaEditado, blnHouveAlteracaoPrecoLista
	blnHouvePrecoVendaEditado = False
	blnHouveAlteracaoPrecoLista = False

	dim v_item, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_totalFamiliaPrecoNFLiquido, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, id_pedido_base
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
				s=Trim(Request.Form("c_preco_lista_original")(i))
				.preco_lista_original=converte_numero(s)
				if .preco_venda <> .preco_venda_original then blnHouvePrecoVendaEditado = True
				if .preco_lista <> .preco_lista_original then blnHouveAlteracaoPrecoLista = True
				end with
			
			for j=LBound(v_item_bd) to UBound(v_item_bd)
				if Trim("" & v_item_bd(j).produto) <> "" then
					if (v_item(ubound(v_item)).fabricante = Trim("" & v_item_bd(j).fabricante)) And (v_item(ubound(v_item)).produto = Trim("" & v_item_bd(j).produto)) then
						v_item(ubound(v_item)).StatusDescontoSuperior = v_item_bd(j).StatusDescontoSuperior
						v_item(ubound(v_item)).IdUsuarioDescontoSuperior = v_item_bd(j).IdUsuarioDescontoSuperior
						v_item(ubound(v_item)).DataHoraDescontoSuperior = v_item_bd(j).DataHoraDescontoSuperior
						exit for
						end if
					end if 'if Trim("" & v_item_bd(j).produto) <> ""
				next
			end if
		next

'	FORMA DE PAGAMENTO (NOVA VERS�O)
'	PARA OS CAMPOS BLOQUEADOS P/ EDI��O, ASSUME O VALOR CADASTRADO ATUALMENTE
	if alerta = "" then
		if (versao_forma_pagamento = "2") And (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then
			rb_forma_pagto = Trim(Request.Form("rb_forma_pagto"))
			if rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA then
				if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then
					op_av_forma_pagto = Trim(Request.Form("op_av_forma_pagto"))
				else
					op_av_forma_pagto = Cstr(r_pedido.av_forma_pagto)
					end if
				if op_av_forma_pagto = "" then alerta = "Indique a forma de pagamento (� vista)."
			elseif rb_forma_pagto = COD_FORMA_PAGTO_PARCELA_UNICA then
				if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then
					op_pu_forma_pagto = Trim(Request.Form("op_pu_forma_pagto"))
				else
					op_pu_forma_pagto = Cstr(r_pedido.pu_forma_pagto)
					end if
				c_pu_valor = Trim(Request.Form("c_pu_valor"))
				if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then
					c_pu_vencto_apos = Trim(Request.Form("c_pu_vencto_apos"))
				else
					c_pu_vencto_apos = Cstr(r_pedido.pu_vencto_apos)
					end if
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
				if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then
					c_pc_qtde = Trim(Request.Form("c_pc_qtde"))
				else
					c_pc_qtde = Cstr(r_pedido.pc_qtde_parcelas)
					end if
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
				if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then
					c_pc_maquineta_qtde = Trim(Request.Form("c_pc_maquineta_qtde"))
				else
					c_pc_maquineta_qtde = Cstr(r_pedido.pc_maquineta_qtde_parcelas)
					end if
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
				if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then
					op_pce_entrada_forma_pagto = Trim(Request.Form("op_pce_entrada_forma_pagto"))
				else
					op_pce_entrada_forma_pagto = Cstr(r_pedido.pce_forma_pagto_entrada)
					end if
				c_pce_entrada_valor = Trim(Request.Form("c_pce_entrada_valor"))
				if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then
					op_pce_prestacao_forma_pagto = Trim(Request.Form("op_pce_prestacao_forma_pagto"))
					c_pce_prestacao_qtde = Trim(Request.Form("c_pce_prestacao_qtde"))
				else
					op_pce_prestacao_forma_pagto = Cstr(r_pedido.pce_forma_pagto_prestacao)
					c_pce_prestacao_qtde = Cstr(r_pedido.pce_prestacao_qtde)
					end if
				c_pce_prestacao_valor = Trim(Request.Form("c_pce_prestacao_valor"))
				if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then
					c_pce_prestacao_periodo = Trim(Request.Form("c_pce_prestacao_periodo"))
				else
					c_pce_prestacao_periodo = Cstr(r_pedido.pce_prestacao_periodo)
					end if
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
				if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then
					op_pse_prim_prest_forma_pagto = Trim(Request.Form("op_pse_prim_prest_forma_pagto"))
				else
					op_pse_prim_prest_forma_pagto = Cstr(r_pedido.pse_forma_pagto_prim_prest)
					end if
				c_pse_prim_prest_valor = Trim(Request.Form("c_pse_prim_prest_valor"))
				if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then
					c_pse_prim_prest_apos = Trim(Request.Form("c_pse_prim_prest_apos"))
					op_pse_demais_prest_forma_pagto = Trim(Request.Form("op_pse_demais_prest_forma_pagto"))
					c_pse_demais_prest_qtde = Trim(Request.Form("c_pse_demais_prest_qtde"))
				else
					c_pse_prim_prest_apos = Cstr(r_pedido.pse_prim_prest_apos)
					op_pse_demais_prest_forma_pagto = Cstr(r_pedido.pse_forma_pagto_demais_prest)
					c_pse_demais_prest_qtde = Cstr(r_pedido.pse_demais_prest_qtde)
					end if
				c_pse_demais_prest_valor = Trim(Request.Form("c_pse_demais_prest_valor"))
				if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then
					c_pse_demais_prest_periodo = Trim(Request.Form("c_pse_demais_prest_periodo"))
				else
					c_pse_demais_prest_periodo = Cstr(r_pedido.pse_demais_prest_periodo)
					end if
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
	dim coeficiente, vlCustoFinancFornecPrecoLista, vlCustoFinancFornecPrecoListaBase, dtCriacaoPedido
	c_custoFinancFornecTipoParcelamentoOriginal = Trim(Request.Form("c_custoFinancFornecTipoParcelamentoOriginal"))
	c_custoFinancFornecQtdeParcelasOriginal = Trim(Request.Form("c_custoFinancFornecQtdeParcelasOriginal"))
	c_custoFinancFornecTipoParcelamento = Trim(Request.Form("c_custoFinancFornecTipoParcelamento"))
	c_custoFinancFornecQtdeParcelas = Trim(Request.Form("c_custoFinancFornecQtdeParcelas"))
	
'	O PEDIDO FOI CADASTRADO J� DENTRO DA POL�TICA DE PERCENTUAL DE CUSTO FINANCEIRO POR FORNECEDOR?
	if versao_forma_pagamento = "2" then
		if (c_custoFinancFornecTipoParcelamentoOriginal <> "") And (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then
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
	dim vl_total_preco_lista, vl_total
	if alerta = "" then
		vl_total_preco_lista = 0
		vl_total = 0
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .produto <> "" then 
					vl_total_preco_lista = vl_total_preco_lista + (.qtde * .preco_lista)
					vl_total = vl_total + (.qtde * .preco_venda)
					end if
				end with
			next
		end if
	
	dim desc_dado_medio
	if vl_total_preco_lista = 0 then
		desc_dado_medio = 0
	else
		desc_dado_medio = 100 * (vl_total_preco_lista - vl_total) / vl_total_preco_lista
		end if

'	ANALISA O PERCENTUAL DE COMISS�O+DESCONTO
	dim perc_max_RT_a_utilizar, perc_max_RT_padrao
	dim perc_comissao_e_desconto_a_utilizar, perc_comissao_e_desconto_padrao, StatusDescontoSuperiorBD, blnAtualizarDadosDescontoSuperior
	dim s_pg, blnPreferencial
	dim vlNivel1, vlNivel2
	perc_max_RT_padrao = rCD.perc_max_comissao
	perc_max_RT_a_utilizar = perc_max_RT_padrao
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
		if blnHouvePrecoVendaEditado Or blnFormaPagtoEditada Or (s_perc_RT <> s_perc_RT_original) then
			'Devido a arredondamentos no front, aceita margem de erro
			if (desc_dado_medio + perc_RT) > (perc_comissao_e_desconto_a_utilizar + MAX_MARGEM_ERRO_PERC_DESC_E_RT) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "A soma dos percentuais de comiss�o (" & formata_perc_RT(perc_RT) & "%) e de desconto m�dio do(s) produto(s) (" & formata_perc(desc_dado_medio) & "%) totaliza " & _
								formata_perc(perc_RT + desc_dado_medio) & "% e excede o m�ximo permitido!"
				end if
			end if
		end if

	if alerta = "" then
		if s_perc_RT <> s_perc_RT_original then
			if perc_RT > perc_max_RT_a_utilizar then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O percentual de comiss�o (" & formata_perc_RT(perc_RT) & "%) excede o m�ximo permitido!"
				end if
			end if
		end if

	'Edit�vel?
	if blnEndEntregaEdicaoLiberada then
		if alerta = "" then
            if (EndEtg_endereco<>r_pedido.EndEtg_endereco) Or (EndEtg_bairro<>r_pedido.EndEtg_bairro) Or (EndEtg_cidade<>r_pedido.EndEtg_cidade) Or (EndEtg_uf<>r_pedido.EndEtg_uf) Or (EndEtg_cep<>r_pedido.EndEtg_cep) Or (EndEtg_obs<>r_pedido.EndEtg_cod_justificativa) then
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

	dim blnHaPedidoAprovadoComEntregaPendente, blnEditouEndEtgPedidoAprovadoComEntregaPendente
	dim sLogEmail, sLogEndEtgEmail
    dim sLogVetor1
    dim sLogVetor2
	blnHaPedidoAprovadoComEntregaPendente = False
	blnEditouEndEtgPedidoAprovadoComEntregaPendente = False
    sLogEmail = ""
	sLogEndEtgEmail = ""
	if alerta = "" then
		if r_pedido.st_entrega <> "ETG" and r_pedido.st_entrega <> "CAN" and CLng(r_pedido.analise_credito) = CLng(COD_AN_CREDITO_OK) then
			'Monitora edi��o no endere�o de cobran�a
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

			'Monitora edi��o no endere�o de entrega: coleta de informa��es p/ mensagem de alerta feita no trecho que monta texto p/ registrar no bloco de notas
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
	
	dim v_desconto()
	ReDim v_desconto(0)
	v_desconto(UBound(v_desconto)) = ""

	dim desc_dado_arredondado
	
	if alerta="" then
		if c_consiste_perc_max_comissao_e_desconto = "S" then
			for i=Lbound(v_item) to Ubound(v_item)
				with v_item(i)
					if .preco_lista = 0 then 
						.desc_dado = 0
						desc_dado_arredondado = 0
					else
						.desc_dado = 100*(.preco_lista-.preco_venda)/.preco_lista
						desc_dado_arredondado = converte_numero(formata_perc_desc(.desc_dado))
						end if
						
					'Se houve edi��o no pre�o de venda, verifica se h� necessidade de atualizar o ID do usu�rio que fez uso da al�ada
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'Soma do desconto e comiss�o est� abaixo do limite padr�o, portanto, assegura que os dados do uso do desconto por al�ada est�o vazios
					if (.desc_dado + perc_RT) <= (perc_comissao_e_desconto_padrao + MAX_MARGEM_ERRO_PERC_DESC_E_RT) then
						.StatusDescontoSuperior = 0
						.IdUsuarioDescontoSuperior = 0
						.DataHoraDescontoSuperior = Null
					else
						'Soma do desconto e comiss�o excede limite padr�o
						'1) Se o usu�rio possui n�vel de al�ada, ou seja, se o limite m�ximo que ele pode utilizar � acima do padr�o,
						'   verifica se houve edi��o que demande o registro ou a atualiza��o do respons�vel pelo desconto superior.
						'2) Se o usu�rio n�o possui n�vel de al�ada, n�o faz nada, pois caso j� existam dados do respons�vel da edi��o anterior que permitiu o desconto superior,
						'   esses dados n�o devem ser sobreescritos
						if perc_comissao_e_desconto_a_utilizar > perc_comissao_e_desconto_padrao then
							'Altera a identifica��o do usu�rio que concedeu o desconto superior nos seguintes casos:
							'	1) Houve edi��o no pre�o de venda
							'	2) Houve aumento no percentual de comiss�o (lembrando que um usu�rio c/ al�ada inferior pode ter sido obrigado a reduzir automaticamente o percentual de comiss�o, situa��o em que n�o faria sentido atualizar o ID do respons�vel pelo desconto)
							'	3) Houve altera��o na forma de pagamento que tenha acarretado altera��o do pre�o de lista, o que implica na altera��o do percentual de desconto
							'H� um tratamento espec�fico para identificar se as altera��es foram realizadas pelo depto financeiro (an�lise de cr�dito), situa��o em que N�O se
							'deve registrar o ID do usu�rio respons�vel pelo desconto.
							'Obs: o depto financeiro edita apenas a forma de pagamento e n�o altera pre�o de venda ou percentual de comiss�o (lembrando que o desconto
							'pode sofrer altera��o caso a forma de pagamento seja alterada entre � Vista e A Prazo e/ou quantidade de parcelas).
							'Caso um usu�rio do depto financeiro edite pre�o de venda e/ou percentual de comiss�o, seu ID ser� registrado normalmente.
							blnAtualizarDadosDescontoSuperior = False
							if (Abs(.preco_venda - .preco_venda_original) > MAX_VALOR_MARGEM_ERRO_PAGAMENTO) _
								OR ( (s_perc_RT <> s_perc_RT_original) And (converte_numero(s_perc_RT) > converte_numero(s_perc_RT_original)) ) _
								OR (blnFormaPagtoEditada And blnHouveAlteracaoPrecoLista) then
								if blnUsuarioDeptoFinanceiro then
									'Se editou pre�o de venda ou aumentou o percentual de comiss�o, registra o ID, mesmo sendo do depto financeiro
									if (Abs(.preco_venda - .preco_venda_original) > MAX_VALOR_MARGEM_ERRO_PAGAMENTO) _
										OR ( (s_perc_RT <> s_perc_RT_original) And (converte_numero(s_perc_RT) > converte_numero(s_perc_RT_original)) ) then
										blnAtualizarDadosDescontoSuperior = True
										end if
								else
									blnAtualizarDadosDescontoSuperior = True
									end if
								end if

							if blnAtualizarDadosDescontoSuperior then
								.StatusDescontoSuperior = 1
								.IdUsuarioDescontoSuperior = r_usuario.Id
								.DataHoraDescontoSuperior = Now
								end if
							end if 'if perc_comissao_e_desconto_a_utilizar > perc_comissao_e_desconto_padrao
						end if 'if ((.desc_dado + perc_RT) <= (perc_comissao_e_desconto_padrao + MAX_MARGEM_ERRO_PERC_DESC_E_RT)) then-else

					'Verifica necessidade de senha de autoriza��o de desconto superior e se essa autoriza��o foi cadastrada
					if (.preco_venda <> .preco_venda_original) Or (blnFormaPagtoEditada And (Not blnUsuarioDeptoFinanceiro)) then
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
	
	if alerta = "" then
		if blnEtgImediataEdicaoLiberada then
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
		end if
	
	dim bln_RT_EdicaoLiberada_Conferencia
	dim blnFamiliaPedidosPossuiPedidoEntregueMesAnterior, blnFamiliaPedidosPossuiPedidoComissaoPaga, blnFamiliaPedidosPossuiPedidoComissaoDescontada
	dim rEdicaoRTMaxPrazo
	set rEdicaoRTMaxPrazo = get_registro_t_parametro(ID_PARAMETRO_Pedido_RT_Edicao_MaxPrazo)
	bln_RT_EdicaoLiberada_Conferencia = False
	blnFamiliaPedidosPossuiPedidoEntregueMesAnterior = False
	blnFamiliaPedidosPossuiPedidoComissaoPaga = False
	blnFamiliaPedidosPossuiPedidoComissaoDescontada = False
	if alerta = "" then
		'Confere se edi��o da RT est� liberada
		'A regra de edi��o do percentual de RT leva em considera��o que o percentual � �nico p/ toda a fam�lia de pedidos
		s = "SELECT" & _
				" pedido" & _
				", comissao_descontada" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
			" WHERE" & _
				" (pedido LIKE '" & retorna_num_pedido_base(pedido_selecionado) & BD_CURINGA_TODOS & "')" & _
				" AND (comissao_descontada = " & COD_COMISSAO_DESCONTADA & ")"
		set rs = cn.Execute(s)
		if Not rs.Eof then blnFamiliaPedidosPossuiPedidoComissaoDescontada = True
		if rs.State <> 0 then rs.Close

		s = "SELECT" & _
				" pedido" & _
				", comissao_descontada" & _
			" FROM t_PEDIDO_PERDA" & _
			" WHERE" & _
				" (pedido LIKE '" & retorna_num_pedido_base(pedido_selecionado) & BD_CURINGA_TODOS & "')" & _
				" AND (comissao_descontada = " & COD_COMISSAO_DESCONTADA & ")"
		set rs = cn.Execute(s)
		if Not rs.Eof then blnFamiliaPedidosPossuiPedidoComissaoDescontada = True
		if rs.State <> 0 then rs.Close

		s = "SELECT" & _
				" pedido" & _
				", st_entrega" & _
				", entregue_data" & _
				", comissao_paga" & _
			" FROM t_PEDIDO" & _
			" WHERE" & _
				" (pedido LIKE '" & retorna_num_pedido_base(pedido_selecionado) & BD_CURINGA_TODOS & "')"
		set rs = cn.Execute(s)
		do while Not rs.Eof
			if (Trim("" & rs("st_entrega")) = ST_ENTREGA_ENTREGUE) And (Not IsMesmoAnoEMes(rs("entregue_data"), Date)) then blnFamiliaPedidosPossuiPedidoEntregueMesAnterior = True
			if CLng(rs("comissao_paga")) = CLng(COD_COMISSAO_PAGA) then blnFamiliaPedidosPossuiPedidoComissaoPaga = True
			rs.MoveNext
			loop
		if rs.State <> 0 then rs.Close

		if operacao_permitida(OP_LJA_EDITA_RT, s_lista_operacoes_permitidas) _
			And ( (rEdicaoRTMaxPrazo.campo_inteiro = 0) Or (Abs(DateDiff("d", r_pedido.data, Date)) <= rEdicaoRTMaxPrazo.campo_inteiro) ) then
			if (Not blnFamiliaPedidosPossuiPedidoComissaoPaga) _
				And (Not blnFamiliaPedidosPossuiPedidoComissaoDescontada) _
				And (Not blnFamiliaPedidosPossuiPedidoEntregueMesAnterior) then
				bln_RT_EdicaoLiberada_Conferencia = True
				end if
			end if
		
		if bln_RT_EdicaoLiberada_Conferencia <> bln_RT_EdicaoLiberada then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Inconsist�ncia encontrada na valida��o da regra de libera��o da edi��o da RT"
			end if
		end if 'if alerta = ""
	
	if alerta <> "" then blnErroConsistencia=True
	
	
'	MENSAGEM DE ALERTA
	dim rEmailDestinatario
	dim corpo_mensagem, id_email, msg_erro_grava_email
	dim s_descricao_forma_pagto, s_descricao_forma_pagto_anterior, quebraLinhaFormaPagto
	dim s_indicador_anterior
	s_descricao_forma_pagto = ""
	s_descricao_forma_pagto_anterior = ""
	quebraLinhaFormaPagto = ",  "


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
	campos_a_omitir = ""
	campos_a_omitir_FP = campos_a_omitir & "|analise_credito|st_recebido|"
	campos_a_omitir_ItemCFF = ""
	campos_a_omitir_PedCFF = ""
	
	pedido_base = retorna_num_pedido_base(pedido_selecionado)

	if alerta = "" then
	'	ATUALIZA O PEDIDO
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
				if (versao_forma_pagamento = "2") And (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then
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
							corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " a forma de pagamento do pedido " & pedido_selecionado & vbCrLf & _
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
															"Edi��o da forma de pagamento em pedido com status 'Cr�dito OK' (pedido " & pedido_selecionado & ")", _
															corpo_mensagem, _
															Now, _
															id_email, _
															msg_erro_grava_email
							end if
						end if

					'Registra edi��o no bloco de notas
					if s_descricao_forma_pagto <> s_descricao_forma_pagto_anterior then
						sBlocoNotasMsg = "Edi��o da forma de pagamento realizada por '" & usuario & "' (status da an�lise de cr�dito: " & descricao_analise_credito(s_analise_credito_a) & ")" & vbCrLf & _
										"Anterior: " & s_descricao_forma_pagto_anterior & vbCrLf & _
										"Nova: " & s_descricao_forma_pagto
						if Not grava_bloco_notas_pedido(pedido_selecionado, ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_FORMA_PAGTO, msg_erro) then
							alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & pedido_selecionado & ")"
							end if
						'Assegura de gravar tamb�m no pedido-base pois trata-se de informa��o controlada atrav�s do pedido-base
						if IsPedidoFilhote(pedido_selecionado) then
							if Not grava_bloco_notas_pedido(retorna_num_pedido_base(pedido_selecionado), ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_FORMA_PAGTO, msg_erro) then
								alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & retorna_num_pedido_base(pedido_selecionado) & ")"
								end if
							end if
						end if
					end if 'if (versao_forma_pagamento = "2") And (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL)
				
				'No caso de pedidos Arclube c/ instalador pode haver uma grande precis�o no percentual de comiss�o, portanto, evita de sobreescrever se n�o tiver havido edi��o
				if (bln_RT_EdicaoLiberada And (s_perc_RT <> s_perc_RT_original)) Or (c_gravar_perc_RT_novo = "S") then rs("perc_RT") = converte_numero(s_perc_RT)

				if blnIndicadorEdicaoLiberada then
					s_indicador_anterior = Trim("" & rs("indicador"))
					rs("indicador") = s_indicador
					if Ucase(Trim(s_indicador_anterior)) <> Ucase(Trim(s_indicador)) then blnEditouIndicador = True

					if blnEditouIndicador And (Trim("" & rs("analise_credito")) = COD_AN_CREDITO_OK) then
						'Envia mensagem de alerta sobre altera��o do indicador em pedido com status "cr�dito ok"
						set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaAlteracaoIndicadorEmPedidoCreditoOk)
						if Trim("" & rEmailDestinatario.campo_texto) <> "" then
							corpo_mensagem = "O usu�rio '" & usuario & "' alterou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " o indicador do pedido " & pedido_selecionado & vbCrLf & _
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
															"Altera��o do indicador em pedido com status 'Cr�dito OK' (pedido " & pedido_selecionado & ")", _
															corpo_mensagem, _
															Now, _
															id_email, _
															msg_erro_grava_email
							end if
						end if

					'Registra edi��o no bloco de notas
					if blnEditouIndicador then
						sBlocoNotasMsg = "Edi��o do indicador realizada por '" & usuario & "' (status da an�lise de cr�dito: " & descricao_analise_credito(s_analise_credito_a) & ")" & vbCrLf & _
										"Anterior: " & s_indicador_anterior & vbCrLf & _
										"Novo: " & s_indicador
						if Not grava_bloco_notas_pedido(pedido_selecionado, ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_INDICADOR, msg_erro) then
							alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & pedido_selecionado & ")"
							end if
						'Assegura de gravar tamb�m no pedido-base pois trata-se de informa��o controlada atrav�s do pedido-base
						if IsPedidoFilhote(pedido_selecionado) then
							if Not grava_bloco_notas_pedido(retorna_num_pedido_base(pedido_selecionado), ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_INDICADOR, msg_erro) then
								alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & retorna_num_pedido_base(pedido_selecionado) & ")"
								end if
							end if
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
			end if 'if IsPedidoFilhote(pedido_selecionado)


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
					if (versao_forma_pagamento = "2") And (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then
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
								corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " a forma de pagamento do pedido " & pedido_selecionado & vbCrLf & _
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
																"Edi��o da forma de pagamento em pedido com status 'Cr�dito OK' (pedido " & pedido_selecionado & ")", _
																corpo_mensagem, _
																Now, _
																id_email, _
																msg_erro_grava_email
								end if
							end if

						'Registra edi��o no bloco de notas
						if s_descricao_forma_pagto <> s_descricao_forma_pagto_anterior then
							sBlocoNotasMsg = "Edi��o da forma de pagamento realizada por '" & usuario & "' (status da an�lise de cr�dito: " & descricao_analise_credito(s_analise_credito_a) & ")" & vbCrLf & _
											"Anterior: " & s_descricao_forma_pagto_anterior & vbCrLf & _
											"Nova: " & s_descricao_forma_pagto
							if Not grava_bloco_notas_pedido(pedido_selecionado, ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_FORMA_PAGTO, msg_erro) then
								alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & pedido_selecionado & ")"
								end if
							'Assegura de gravar tamb�m no pedido-base pois trata-se de informa��o controlada atrav�s do pedido-base
							if IsPedidoFilhote(pedido_selecionado) then
								if Not grava_bloco_notas_pedido(retorna_num_pedido_base(pedido_selecionado), ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_FORMA_PAGTO, msg_erro) then
									alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & retorna_num_pedido_base(pedido_selecionado) & ")"
									end if
								end if
							end if
						end if 'if (versao_forma_pagamento = "2") And (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL)

					if blnIndicadorEdicaoLiberada then
						s_indicador_anterior = Trim("" & rs("indicador"))
						rs("indicador") = s_indicador
						if Ucase(Trim(s_indicador_anterior)) <> Ucase(Trim(s_indicador)) then blnEditouIndicador = True

						if blnEditouIndicador And (Trim("" & rs("analise_credito")) = COD_AN_CREDITO_OK) then
							'Envia mensagem de alerta sobre altera��o do indicador em pedido com status "cr�dito ok"
							set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaAlteracaoIndicadorEmPedidoCreditoOk)
							if Trim("" & rEmailDestinatario.campo_texto) <> "" then
								corpo_mensagem = "O usu�rio '" & usuario & "' alterou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " o indicador do pedido " & pedido_selecionado & vbCrLf & _
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
																"Altera��o do indicador em pedido com status 'Cr�dito OK' (pedido " & pedido_selecionado & ")", _
																corpo_mensagem, _
																Now, _
																id_email, _
																msg_erro_grava_email
								end if
							end if
					
						'Registra edi��o no bloco de notas
						if blnEditouIndicador then
							sBlocoNotasMsg = "Edi��o do indicador realizada por '" & usuario & "' (status da an�lise de cr�dito: " & descricao_analise_credito(s_analise_credito_a) & ")" & vbCrLf & _
											"Anterior: " & s_indicador_anterior & vbCrLf & _
											"Novo: " & s_indicador
							if Not grava_bloco_notas_pedido(pedido_selecionado, ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_INDICADOR, msg_erro) then
								alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & pedido_selecionado & ")"
								end if
							'Assegura de gravar tamb�m no pedido-base pois trata-se de informa��o controlada atrav�s do pedido-base
							if IsPedidoFilhote(pedido_selecionado) then
								if Not grava_bloco_notas_pedido(retorna_num_pedido_base(pedido_selecionado), ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_INDICADOR, msg_erro) then
									alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & retorna_num_pedido_base(pedido_selecionado) & ")"
									end if
								end if
							end if
						end if
					end if 'if Not IsPedidoFilhote(pedido_selecionado)
				
				if blnObs1EdicaoLiberada then
                     rs("obs_1") = s_obs1
                     rs("NFe_texto_constar") = s_nf_texto
                     rs("NFe_xPed") = s_num_pedido_compra
                end if
				
				if (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then rs("forma_pagto") = s_forma_pagto

				'No caso de pedidos Arclube c/ instalador pode haver uma grande precis�o no percentual de comiss�o, portanto, evita de sobreescrever se n�o tiver havido edi��o
				if (bln_RT_EdicaoLiberada And (s_perc_RT <> s_perc_RT_original)) Or (c_gravar_perc_RT_novo = "S") then rs("perc_RT") = converte_numero(s_perc_RT)
				
				if (versao_forma_pagamento = "1") And (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then
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
						
						blnMarketplaceCodigoOrigemAlterado = True
						rs("marketplace_codigo_origem")=s_pedido_origem
						end if
					end if

				'Guardar informa��es de endere�o presentes no pedido (consist�ncia para verificar se mudou de endere�o)
				dim st_end_entrega_anterior, EndEtg_cep_anterior, blnEndereco_cep_alterado
				st_end_entrega_anterior = rs("st_end_entrega")
				EndEtg_cep_anterior = rs("EndEtg_cep")
                blnEndereco_cep_alterado = False


				if r_pedido.st_memorizacao_completa_enderecos = 1 or r_pedido.st_memorizacao_completa_enderecos = 9 then
					sEnderecoOriginal = formata_endereco(Trim("" & rs("endereco_logradouro")), Trim("" & rs("endereco_numero")), Trim("" & rs("endereco_complemento")), Trim("" & rs("endereco_bairro")), Trim("" & rs("endereco_cidade")), Trim("" & rs("endereco_uf")), Trim("" & rs("endereco_cep")))
					sEnderecoNovo = formata_endereco(endereco__endereco, endereco__numero, endereco__complemento, endereco__bairro, endereco__cidade, endereco__uf, endereco__cep)
					if UCase(sEnderecoOriginal) <> UCase(sEnderecoNovo) then
						sBlocoNotasEndCob = "Endere�o de cobran�a: " & vbCrLf & _
											String(4, " ") & "Anterior: " & sEnderecoOriginal & vbCrLf & _
											String(4, " ") & "Novo: " & sEnderecoNovo
						end if

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
					if rs("st_end_entrega") = 0 then
						sEnderecoOriginal = "(N.I.)"
					else
						sEnderecoOriginal = formata_endereco(Trim("" & rs("EndEtg_endereco")), Trim("" & rs("EndEtg_endereco_numero")), Trim("" & rs("EndEtg_endereco_complemento")), Trim("" & rs("EndEtg_bairro")), Trim("" & rs("EndEtg_cidade")), Trim("" & rs("EndEtg_uf")), Trim("" & rs("EndEtg_cep")))
						end if

					if EndEtg_endereco <> "" then 
						rs("st_end_entrega") = 1
						sEnderecoNovo = formata_endereco(EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep)
					else
						rs("st_end_entrega") = 0
						sEnderecoNovo = "(N.I.)"
						end if
					
					if UCase(sEnderecoOriginal) <> UCase(sEnderecoNovo) then
						sBlocoNotasEndEtg = "Endere�o de entrega: " & vbCrLf & _
											String(4, " ") & "Anterior: " & sEnderecoOriginal & vbCrLf & _
											String(4, " ") & "Novo: " & sEnderecoNovo
						
						'Monitora edi��o no endere�o de entrega
						if r_pedido.st_entrega <> "ETG" and r_pedido.st_entrega <> "CAN" and CLng(r_pedido.analise_credito) = CLng(COD_AN_CREDITO_OK) then
							blnEditouEndEtgPedidoAprovadoComEntregaPendente = True
							sLogEndEtgEmail = vbTab & "Endere�o de entrega anterior: " & sEnderecoOriginal & vbCrLf & _
												vbTab & "Endere�o de entrega novo: " & sEnderecoNovo
							end if
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
				
				
				if blnEtgImediataEdicaoLiberada then
					s_etg_imediata_original = Trim("" & rs("st_etg_imediata"))
					if s_etg_imediata <> "" then 
						if CLng(rs("st_etg_imediata")) <> CLng(s_etg_imediata) then
							rs("st_etg_imediata")=CLng(s_etg_imediata)
							rs("etg_imediata_data")=Now
							rs("etg_imediata_usuario")=usuario
							end if
						end if
					
					if CLng(s_etg_imediata) = CLng(COD_ETG_IMEDIATA_NAO) then
						if (s_etg_imediata_original <> Trim(s_etg_imediata)) Or (formata_data(rs("PrevisaoEntregaData")) <> formata_data(StrToDate(c_data_previsao_entrega))) then
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

				if blnGarantiaIndicadorEdicaoLiberada then
					if CLng(rs("GarantiaIndicadorStatus")) <> CLng(rb_garantia_indicador) then
						rs("GarantiaIndicadorStatus")=CLng(rb_garantia_indicador)
						rs("GarantiaIndicadorUsuarioUltAtualiz")=usuario
						rs("GarantiaIndicadorDtHrUltAtualiz")=Now
						end if
					end if
				rs("pedido_bs_x_at")=s_ped_bonshop
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

				if alerta = "" then
					if (sBlocoNotasEndCob <> "") Or (sBlocoNotasEndEtg <> "") then
						sBlocoNotasMsg = sBlocoNotasEndCob
						if (sBlocoNotasMsg <> "") And (sBlocoNotasEndEtg <> "") then sBlocoNotasMsg = sBlocoNotasMsg & vbCrLf & vbCrLf
						sBlocoNotasMsg = sBlocoNotasMsg & sBlocoNotasEndEtg
						sBlocoNotasMsg = "Edi��o de endere�o realizada por '" & usuario & "' (status da an�lise de cr�dito: " & descricao_analise_credito(s_analise_credito_a) & ")" & vbCrLf & _
										vbCrLf & _
										sBlocoNotasMsg
						if Not grava_bloco_notas_pedido(pedido_selecionado, ID_USUARIO_SISTEMA, loja, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_ENDERECO, msg_erro) then
							alerta = "Falha ao gravar bloco de notas com mensagem autom�tica no pedido (" & pedido_selecionado & ")"
							end if
						end if
					end if
				end if
			end if
		
		if alerta = "" then
		'	O PEDIDO FOI CADASTRADO J� DENTRO DA POL�TICA DE PERCENTUAL DE CUSTO FINANCEIRO POR FORNECEDOR?
			if c_custoFinancFornecTipoParcelamentoOriginal <> "" then
				if (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then
					if (c_custoFinancFornecTipoParcelamentoOriginal <> c_custoFinancFornecTipoParcelamento) Or _
					   (c_custoFinancFornecQtdeParcelasOriginal <> c_custoFinancFornecQtdeParcelas) then
						for i=Lbound(v_item) to Ubound(v_item)
							with v_item(i)
								if Trim(.produto)<>"" then
								'	Inicializa��o
									vlCustoFinancFornecPrecoListaBase = 0
									coeficiente = 0
									
								'	Obt�m Pre�o de Lista Base
									'Localiza os dados do momento em que o pedido foi criado
									blnAchou = False
									for j=LBound(v_item_bd) to UBound(v_item_bd)
										if Trim("" & v_item_bd(j).produto) <> "" then
											if (Trim("" & v_item_bd(j).fabricante) = Trim("" & v_item(i).fabricante)) And (Trim("" & v_item_bd(j).produto) = Trim("" & v_item(i).produto)) then
												vlCustoFinancFornecPrecoListaBase = v_item_bd(j).custoFinancFornecPrecoListaBase
												blnAchou = True
												exit for
												end if
											end if 'if Trim("" & v_item_bd(j).produto) <> ""
										next

									if Not blnAchou then
										alerta=texto_add_br(alerta)
										alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " N�O foi encontrado no pedido " & pedido_selecionado
										end if
									
								'	Obt�m coeficiente do custo financeiro
									if c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA then
										coeficiente = 1
									else
										'Inicializa��o do coeficiente
										coeficiente = 0
										dtCriacaoPedido = Null

										'Tenta ocalizar os dados do momento em que o pedido foi criado
										s = "SELECT data FROM t_PEDIDO WHERE (pedido = '" & retorna_num_pedido_base(pedido_selecionado) & "')"
										set rs2 = cn.execute(s)
										if Not rs2.Eof then
											dtCriacaoPedido = rs2("data")
											end if

										if Not Isnull(dtCriacaoPedido) then
											s = "SELECT " & _
													"*" & _
												" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR_HISTORICO" & _
												" WHERE" & _
													" (data = " & bd_formata_data(dtCriacaoPedido) & ")" & _
													" AND (fabricante = '" & .fabricante & "')" & _
													" AND (tipo_parcelamento = '" & c_custoFinancFornecTipoParcelamento & "')" & _
													" AND (qtde_parcelas = " & c_custoFinancFornecQtdeParcelas & ")"
											set rs2 = cn.execute(s)
											if Not rs2.Eof then
												coeficiente = converte_numero(rs2("coeficiente"))
												end if
											end if 'if Not Isnull(dtCriacaoPedido)

										'Se n�o encontrou dados originais, pesquisa pelo coeficiente registrado no hist�rico p/ a data de hoje
										if coeficiente = 0 then
											s = "SELECT " & _
													"*" & _
												" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR_HISTORICO" & _
												" WHERE" & _
													" (data = " & bd_formata_data(Now) & ")" & _
													" AND (fabricante = '" & .fabricante & "')" & _
													" AND (tipo_parcelamento = '" & c_custoFinancFornecTipoParcelamento & "')" & _
													" AND (qtde_parcelas = " & c_custoFinancFornecQtdeParcelas & ")"
											set rs2 = cn.execute(s)
											if Not rs2.Eof then
												coeficiente = converte_numero(rs2("coeficiente"))
												end if
											end if 'if coeficiente = 0

										'Em �ltimo caso, pesquisa pelo coeficiente atual
										if coeficiente = 0 then
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
											end if 'if coeficiente = 0
										end if 'if c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA then-else
									
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
											vlCustoFinancFornecPrecoLista=converte_numero(formata_moeda(coeficiente*vlCustoFinancFornecPrecoListaBase))
											rs("preco_lista")=vlCustoFinancFornecPrecoLista
											if vlCustoFinancFornecPrecoLista = 0 then 
												rs("desc_dado") = 0 
											else
												rs("desc_dado") = 100*(vlCustoFinancFornecPrecoLista-rs("preco_venda"))/vlCustoFinancFornecPrecoLista
												end if

											'Verifica se h� necessidade de atualizar o ID do usu�rio que fez uso da al�ada (lembrando que a necessidade de atualizar esses dados pode decorrer da edi��o da RT)
											'O campo 'StatusDescontoSuperior' � do tipo 'bit' e o ASP converte como se fosse um boolean, portanto, � feito um tratamento manual
											if rs("StatusDescontoSuperior") = 0 then StatusDescontoSuperiorBD = 0 else StatusDescontoSuperiorBD = 1
											if (StatusDescontoSuperiorBD <> .StatusDescontoSuperior) Or (Trim("" & rs("IdUsuarioDescontoSuperior")) <> Trim("" & .IdUsuarioDescontoSuperior)) then
												if s_log <> "" then s_log = s_log & "; "
												s_log = s_log & _
														"Desconto por al�ada no item (" & .fabricante & ")" & .produto & ": " & _
														"StatusDescontoSuperior: " & monta_campo_log(StatusDescontoSuperiorBD) & " => " & monta_campo_log(.StatusDescontoSuperior) & "; " & _
														"IdUsuarioDescontoSuperior: " & monta_campo_log(rs("IdUsuarioDescontoSuperior")) & " => " & monta_campo_log(.IdUsuarioDescontoSuperior) & "; " & _
														"DataHoraDescontoSuperior: " & monta_campo_log(rs("DataHoraDescontoSuperior")) & " => " & monta_campo_log(.DataHoraDescontoSuperior)
												
												rs("StatusDescontoSuperior") = .StatusDescontoSuperior
												if .StatusDescontoSuperior <> 0 then
													rs("IdUsuarioDescontoSuperior") = CLng(.IdUsuarioDescontoSuperior)
													rs("DataHoraDescontoSuperior") = .DataHoraDescontoSuperior
												else
													rs("IdUsuarioDescontoSuperior") = Null
													rs("DataHoraDescontoSuperior") = Null
													end if
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
							
							'Verifica se h� necessidade de atualizar o ID do usu�rio que fez uso da al�ada (lembrando que a necessidade de atualizar esses dados pode decorrer da edi��o da RT)
							'O campo 'StatusDescontoSuperior' � do tipo 'bit' e o ASP converte como se fosse um boolean, portanto, � feito um tratamento manual
							if rs("StatusDescontoSuperior") = 0 then StatusDescontoSuperiorBD = 0 else StatusDescontoSuperiorBD = 1
							if (StatusDescontoSuperiorBD <> .StatusDescontoSuperior) Or (Trim("" & rs("IdUsuarioDescontoSuperior")) <> Trim("" & .IdUsuarioDescontoSuperior)) then
								if s_log <> "" then s_log = s_log & "; "
								s_log = s_log & _
										"Desconto por al�ada no item (" & .fabricante & ")" & .produto & ": " & _
										"StatusDescontoSuperior: " & monta_campo_log(StatusDescontoSuperiorBD) & " => " & monta_campo_log(.StatusDescontoSuperior) & "; " & _
										"IdUsuarioDescontoSuperior: " & monta_campo_log(rs("IdUsuarioDescontoSuperior")) & " => " & monta_campo_log(.IdUsuarioDescontoSuperior) & "; " & _
										"DataHoraDescontoSuperior: " & monta_campo_log(rs("DataHoraDescontoSuperior")) & " => " & monta_campo_log(.DataHoraDescontoSuperior)
								
								rs("StatusDescontoSuperior") = .StatusDescontoSuperior
								if .StatusDescontoSuperior <> 0 then
									rs("IdUsuarioDescontoSuperior") = CLng(.IdUsuarioDescontoSuperior)
									rs("DataHoraDescontoSuperior") = .DataHoraDescontoSuperior
								else
									rs("IdUsuarioDescontoSuperior") = Null
									rs("DataHoraDescontoSuperior") = Null
									end if
								blnUpdate = True
								end if
							end if 'if rs.EOF then-else
						
						if blnUpdate then rs.Update
						end if 'if Trim(.produto)<>""
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
				if (versao_forma_pagamento = "2") And (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then
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
						rs("vendedor") = r_pedido.vendedor
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

		'Ajusta o indicador de todos os pedidos da fam�lia
		if alerta = "" then
			if blnEditouIndicador then
				s = "UPDATE t_PED__FILHOTE" & _
					" SET" & _
						" t_PED__FILHOTE.indicador = t_PED__BASE.indicador" & _
					" FROM t_PEDIDO AS t_PED__FILHOTE" & _
						" INNER JOIN t_PEDIDO AS t_PED__BASE ON (t_PED__FILHOTE.pedido_base = t_PED__BASE.pedido)" & _
					" WHERE" & _
						" (t_PED__FILHOTE.pedido_base = '" & retorna_num_pedido_base(pedido_selecionado) & "')" & _
						" AND (t_PED__FILHOTE.pedido <> t_PED__FILHOTE.pedido_base)"
				cn.Execute(s)
				If Err <> 0 then
					alerta = "FALHA AO SINCRONIZAR O CAMPO 'indicador' (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if
			end if

		if alerta = "" then
			if blnHaPedidoAprovadoComEntregaPendente then
				''Envia alerta de que houve edi��o no cadastro de cliente que possui pedido com status de an�lise de cr�dito 'cr�dito ok' e com entrega pendente
				set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoCadastroClienteComPedidoCreditoOkEntregaPendente)
				if Trim("" & rEmailDestinatario.campo_texto) <> "" then
					
					corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " o endere�o do cliente:" & _
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
																		"Edi��o no endere�o de cliente que possui pedido com status 'Cr�dito OK' e entrega pendente (pedido " & r_pedido.pedido & ")", _
																		corpo_mensagem, _
																		Now, _
																		id_email, _
																		msg_erro_grava_email
				end if
			end if
		end if

		if alerta = "" then
			if blnEditouEndEtgPedidoAprovadoComEntregaPendente then
				''Envia alerta de que houve edi��o no endere�o de entrega em pedido com status de an�lise de cr�dito 'cr�dito ok' e com entrega pendente
				set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoCadastroClienteComPedidoCreditoOkEntregaPendente)
				if Trim("" & rEmailDestinatario.campo_texto) <> "" then
					if r_pedido.st_memorizacao_completa_enderecos <> 0 And blnUsarMemorizacaoCompletaEnderecos then
						s = cnpj_cpf_formata(r_pedido.endereco_cnpj_cpf) & " - " & r_pedido.endereco_nome
					else
						s = cnpj_cpf_formata(r_cliente.cnpj_cpf) & " - " & r_cliente.nome
						end if
					
					corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " o endere�o de entrega do pedido " & r_pedido.pedido & ":" & _
									vbCrLf & _
										"Cliente: " & s & _
										vbCrLf & _
										"A altera��o foi realizada para o pedido: '" & r_pedido.pedido & "' que possui o status de an�lise de cr�dito 'Cr�dito OK' e com entrega pendente." & _
										vbCrLf & _
										vbCrLf & _
										"Informa��es detalhadas sobre as altera��es:" & vbCrLf & _
										substitui_caracteres(sLogEndEtgEmail, vbTab, String(4, " "))
									
										EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
																		"", _
																		rEmailDestinatario.campo_texto, _
																		"", _
																		"", _
																		"Edi��o no endere�o de entrega em pedido com status 'Cr�dito OK' e entrega pendente (pedido " & r_pedido.pedido & ")", _
																		corpo_mensagem, _
																		Now, _
																		id_email, _
																		msg_erro_grava_email
				end if
			end if
		end if
		
		if alerta = "" then
			'Prepara dados que ser�o usados no processamento a seguir
			if Not le_pedido(pedido_selecionado, r_pedido_atualizado, msg_erro) then
				alerta = msg_erro
				end if
			end if
		
		if alerta = "" then
			'Email de alerta p/ a equipe do financeiro caso tenha havido edi��o na forma de pagamento alterando para algum meio de pagamento monitorado
			s = ID_FORMA_PAGTO_BOLETO & "|" & ID_FORMA_PAGTO_CARTAO & "|" & ID_FORMA_PAGTO_CARTAO_MAQUINETA
			vMeioPagtoMonitorado = Split(s, "|")
			
			sMeioPagtoMonitoradoIdentificado = ""
			for iMeioPagtoMonitorado = LBound(vMeioPagtoMonitorado) to UBound(vMeioPagtoMonitorado)
				idMeioPagtoMonitorado = Trim("" & vMeioPagtoMonitorado(iMeioPagtoMonitorado))
				if idMeioPagtoMonitorado <> "" then
					if parcelamentoPassouPossuirMeioPagamento(r_pedido, r_pedido_atualizado, idMeioPagtoMonitorado, False) then
						if sMeioPagtoMonitoradoIdentificado <> "" then sMeioPagtoMonitoradoIdentificado = sMeioPagtoMonitoradoIdentificado & ", "
						sMeioPagtoMonitoradoIdentificado = sMeioPagtoMonitoradoIdentificado & "'" & x_opcao_forma_pagamento(idMeioPagtoMonitorado) & "'"
						end if
					end if
				next
			
			if sMeioPagtoMonitoradoIdentificado <> "" then
				set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoFormaPagtoPassouPossuirMeioPagtoMonitorado)
				if Trim("" & rEmailDestinatario.campo_texto) <> "" then
					corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " a forma de pagamento do pedido " & pedido_selecionado & " incluindo o meio de pagamento: " & sMeioPagtoMonitoradoIdentificado & vbCrLf & _
									vbCrLf & _
									"Forma de pagamento anterior:" & vbCrLf & _
									monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido, quebraLinhaFormaPagto) & vbCrLf & _
									vbCrLf & _
									"Forma de pagamento atual:" & vbCrLf & _
									monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido_atualizado, quebraLinhaFormaPagto) & vbCrLf & _
									vbCrLf & _
									"Informa��es adicionais:" & vbCrLf & _
									"Status da an�lise de cr�dito: " & x_analise_credito(r_pedido.analise_credito) & vbCrLf & _
									"Status de pagamento: " & Ucase(x_status_pagto(r_pedido.st_pagto))
					
					EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
													"", _
													rEmailDestinatario.campo_texto, _
													"", _
													"", _
													"Edi��o da forma de pagamento incluindo o meio de pagamento: " & sMeioPagtoMonitoradoIdentificado & " (pedido " & pedido_selecionado & ")", _
													corpo_mensagem, _
													Now, _
													id_email, _
													msg_erro_grava_email
					end if 'if Trim("" & rEmailDestinatario.campo_texto) <> ""
				end if 'if sMeioPagtoMonitoradoIdentificado <> ""
			end if 'if alerta = ""
		
		if alerta = "" then
			'Email de alerta p/ a equipe do financeiro caso tenha havido edi��o na forma de pagamento de modo que uma parcela a prazo com meio de pagamento monitorado
			'tenha sido alterada para outro meio de pagamento tamb�m monitorado
			s = ID_FORMA_PAGTO_BOLETO & "|" & ID_FORMA_PAGTO_CARTAO & "|" & ID_FORMA_PAGTO_CARTAO_MAQUINETA
			vMeioPagtoMonitorado = Split(s, "|")
			
			if houveEdicaoParcelaAPrazoEntreMeiosPagtoMonitorados(r_pedido, r_pedido_atualizado, vMeioPagtoMonitorado, sMeioPagtoMonitoradoIdentificado) then
				set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoFormaPagtoMonitoramentoParcelaAPrazo)
				if Trim("" & rEmailDestinatario.campo_texto) <> "" then
					corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " a forma de pagamento do pedido " & pedido_selecionado & " alterando meio de pagamento de parcela a prazo incluindo: " & sMeioPagtoMonitoradoIdentificado & vbCrLf & _
									vbCrLf & _
									"Forma de pagamento anterior:" & vbCrLf & _
									monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido, quebraLinhaFormaPagto) & vbCrLf & _
									vbCrLf & _
									"Forma de pagamento atual:" & vbCrLf & _
									monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido_atualizado, quebraLinhaFormaPagto) & vbCrLf & _
									vbCrLf & _
									"Informa��es adicionais:" & vbCrLf & _
									"Status da an�lise de cr�dito: " & x_analise_credito(r_pedido.analise_credito) & vbCrLf & _
									"Status de pagamento: " & Ucase(x_status_pagto(r_pedido.st_pagto))
					
					EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
													"", _
													rEmailDestinatario.campo_texto, _
													"", _
													"", _
													"Edi��o da forma de pagamento alterando meio de pagamento de parcela a prazo incluindo: " & sMeioPagtoMonitoradoIdentificado & " (pedido " & pedido_selecionado & ")", _
													corpo_mensagem, _
													Now, _
													id_email, _
													msg_erro_grava_email
					end if 'if Trim("" & rEmailDestinatario.campo_texto) <> ""
				end if 'if houveEdicaoParcelaAPrazoEntreMeiosPagtoMonitorados(r_pedido, r_pedido_atualizado, vMeioPagtoMonitorado)
			end if 'if alerta = ""
		
		if alerta = "" then
			s = ID_FORMA_PAGTO_DINHEIRO & "|" & ID_FORMA_PAGTO_DEPOSITO & "|" & ID_FORMA_PAGTO_CHEQUE
			vMeioPagtoMonitorado = Split(s, "|")
			
			'Email de alerta p/ a equipe do financeiro caso tenha havido edi��o na forma de pagamento alterando parcela a prazo para algum meio de pagamento monitorado
			sMeioPagtoMonitoradoIdentificado = ""
			for iMeioPagtoMonitorado = LBound(vMeioPagtoMonitorado) to UBound(vMeioPagtoMonitorado)
				idMeioPagtoMonitorado = Trim("" & vMeioPagtoMonitorado(iMeioPagtoMonitorado))
				if idMeioPagtoMonitorado <> "" then
					if parcelamentoPassouPossuirParcelaAPrazoComMeioPagto(r_pedido, r_pedido_atualizado, idMeioPagtoMonitorado, False) then
						if sMeioPagtoMonitoradoIdentificado <> "" then sMeioPagtoMonitoradoIdentificado = sMeioPagtoMonitoradoIdentificado & ", "
						sMeioPagtoMonitoradoIdentificado = sMeioPagtoMonitoradoIdentificado & "'" & x_opcao_forma_pagamento(idMeioPagtoMonitorado) & "'"
						end if
					end if
				next
			
			if sMeioPagtoMonitoradoIdentificado <> "" then
				set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoFormaPagtoMonitoramentoParcelaAPrazo)
				if Trim("" & rEmailDestinatario.campo_texto) <> "" then
					corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " a forma de pagamento do pedido " & pedido_selecionado & " incluindo meio de pagamento em parcela a prazo: " & sMeioPagtoMonitoradoIdentificado & vbCrLf & _
									vbCrLf & _
									"Forma de pagamento anterior:" & vbCrLf & _
									monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido, quebraLinhaFormaPagto) & vbCrLf & _
									vbCrLf & _
									"Forma de pagamento atual:" & vbCrLf & _
									monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido_atualizado, quebraLinhaFormaPagto) & vbCrLf & _
									vbCrLf & _
									"Informa��es adicionais:" & vbCrLf & _
									"Status da an�lise de cr�dito: " & x_analise_credito(r_pedido.analise_credito) & vbCrLf & _
									"Status de pagamento: " & Ucase(x_status_pagto(r_pedido.st_pagto))
					
					EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
													"", _
													rEmailDestinatario.campo_texto, _
													"", _
													"", _
													"Edi��o da forma de pagamento incluindo meio de pagamento em parcela a prazo: " & sMeioPagtoMonitoradoIdentificado & " (pedido " & pedido_selecionado & ")", _
													corpo_mensagem, _
													Now, _
													id_email, _
													msg_erro_grava_email
					end if 'if Trim("" & rEmailDestinatario.campo_texto) <> ""
				end if 'if sMeioPagtoMonitoradoIdentificado <> ""
			end if 'if alerta = ""
		
		if alerta = "" then
			'Monitoramento da forma de pagamento "� vista"
			if houve_edicao_forma_pagto_pedido(r_pedido, r_pedido_atualizado) And (CStr(r_pedido_atualizado.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA) then
				set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoFormaPagtoAVista)
				if Trim("" & rEmailDestinatario.campo_texto) <> "" then
					corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " a forma de pagamento do pedido " & pedido_selecionado & " para: " & monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido_atualizado, quebraLinhaFormaPagto) & vbCrLf & _
									vbCrLf & _
									"Forma de pagamento anterior:" & vbCrLf & _
									monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido, quebraLinhaFormaPagto) & vbCrLf & _
									vbCrLf & _
									"Forma de pagamento atual:" & vbCrLf & _
									monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido_atualizado, quebraLinhaFormaPagto) & vbCrLf & _
									vbCrLf & _
									"Informa��es adicionais:" & vbCrLf & _
									"Status da an�lise de cr�dito: " & x_analise_credito(r_pedido.analise_credito) & vbCrLf & _
									"Status de pagamento: " & Ucase(x_status_pagto(r_pedido.st_pagto))

					EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
													"", _
													rEmailDestinatario.campo_texto, _
													"", _
													"", _
													"Edi��o da forma de pagamento para: '" & monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido_atualizado, quebraLinhaFormaPagto) & "' (pedido " & pedido_selecionado & ")", _
													corpo_mensagem, _
													Now, _
													id_email, _
													msg_erro_grava_email
					end if 'if Trim("" & rEmailDestinatario.campo_texto) <> ""
				end if 'if houve_edicao_forma_pagto_pedido(r_pedido, r_pedido_atualizado) And (CStr(r_pedido_atualizado.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA)
			end if 'if alerta = ""

		if alerta = "" then
			'Email de alerta p/ a equipe do financeiro caso tenha havido edi��o em pedido que possua parcela em 'Boleto AV' e que esteja com status de pagamento 'PAGO'
			'Obs: analisa a situa��o do pedido antes e depois da altera��o
			idMeioPagtoMonitorado = ID_FORMA_PAGTO_BOLETO_AV
			if houve_edicao_forma_pagto_pedido(r_pedido, r_pedido_atualizado) _
				And (parcelamentoPossuiMeioPagamento(r_pedido, idMeioPagtoMonitorado) OR parcelamentoPossuiMeioPagamento(r_pedido_atualizado, idMeioPagtoMonitorado)) _
				And ((r_pedido.st_pagto = ST_PAGTO_PAGO) Or (CLng(r_pedido.analise_credito) = CLng(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO))) then
					set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoFormaPagtoComBoletoAV)
					if Trim("" & rEmailDestinatario.campo_texto) <> "" then
						corpo_mensagem = "O usu�rio '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " a forma de pagamento do pedido " & pedido_selecionado & " que possui o meio de pagamento: '" & x_opcao_forma_pagamento(idMeioPagtoMonitorado) & "'" & vbCrLf & _
										vbCrLf & _
										"Forma de pagamento anterior:" & vbCrLf & _
										monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido, quebraLinhaFormaPagto) & vbCrLf & _
										vbCrLf & _
										"Forma de pagamento atual:" & vbCrLf & _
										monta_descricao_forma_pagto_pedido_com_quebra_linha(r_pedido_atualizado, quebraLinhaFormaPagto) & vbCrLf & _
										vbCrLf & _
										"Informa��es adicionais:" & vbCrLf & _
										"Status da an�lise de cr�dito: " & x_analise_credito(r_pedido.analise_credito) & vbCrLf & _
										"Status de pagamento: " & Ucase(x_status_pagto(r_pedido.st_pagto))

						EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
														"", _
														rEmailDestinatario.campo_texto, _
														"", _
														"", _
														"Edi��o da forma de pagamento que possui o meio de pagamento: '" & x_opcao_forma_pagamento(idMeioPagtoMonitorado) & "' (pedido " & pedido_selecionado & ")", _
														corpo_mensagem, _
														Now, _
														id_email, _
														msg_erro_grava_email
						end if 'if Trim("" & rEmailDestinatario.campo_texto) <> ""
				end if 'if (crit�rios da situa��o do pedido)
			end if 'if alerta = ""


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