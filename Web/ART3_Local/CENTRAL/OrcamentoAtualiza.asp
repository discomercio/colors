<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  O R C A M E N T O A T U A L I Z A . A S P
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

	dim s, usuario, orcamento_selecionado
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	orcamento_selecionado = ucase(Trim(request("orcamento_selecionado")))
	if (orcamento_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_NAO_ESPECIFICADO)
	s = normaliza_num_orcamento(orcamento_selecionado)
	if s <> "" then orcamento_selecionado = s
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

'	FORMA DE PAGAMENTO (NOVA VERSÃO)
	dim versao_forma_pagamento
	dim rb_forma_pagto, op_av_forma_pagto, c_pc_qtde, c_pc_valor, c_pc_maquineta_qtde, c_pc_maquineta_valor
	dim op_pu_forma_pagto, c_pu_valor, c_pu_vencto_apos
	dim op_pce_entrada_forma_pagto, c_pce_entrada_valor, op_pce_prestacao_forma_pagto, c_pce_prestacao_qtde, c_pce_prestacao_valor, c_pce_prestacao_periodo
	dim op_pse_prim_prest_forma_pagto, c_pse_prim_prest_valor, c_pse_prim_prest_apos, op_pse_demais_prest_forma_pagto, c_pse_demais_prest_qtde, c_pse_demais_prest_valor, c_pse_demais_prest_periodo
	dim s_perc_RT
	versao_forma_pagamento = Trim(Request.Form("versao_forma_pagamento"))

	dim FormaPagtoBloqueado, blnFormaPagtoBloqueado
	FormaPagtoBloqueado = Trim(Request.Form("FormaPagtoBloqueado"))
	blnFormaPagtoBloqueado = CBool(FormaPagtoBloqueado)
	
	dim s_qtde_parcelas, s_forma_pagto, s_obs1, s_obs2
	dim s_etg_imediata, s_etg_imediata_original, s_bem_uso_consumo, c_data_previsao_entrega
	s_obs1=Trim(request("c_obs1"))
	s_obs2=Trim(request("c_obs2"))
	s_etg_imediata=Trim(request("rb_etg_imediata"))
	c_data_previsao_entrega = Trim(Request("c_data_previsao_entrega"))
	s_bem_uso_consumo=Trim(request("rb_bem_uso_consumo"))
	s_forma_pagto=Trim(request("c_forma_pagto"))
	if versao_forma_pagamento = "1" then
		s_qtde_parcelas=retorna_so_digitos(request("c_qtde_parcelas"))
		end if
	s_perc_RT = Trim(request("c_perc_RT"))

	dim c_FlagEndEntregaEditavel, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs,blnEndEtg_obs
	dim EndEtg_email, EndEtg_email_xml, EndEtg_nome, EndEtg_ddd_res, EndEtg_tel_res, EndEtg_ddd_com, EndEtg_tel_com, EndEtg_ramal_com
	dim EndEtg_ddd_cel, EndEtg_tel_cel, EndEtg_ddd_com_2, EndEtg_tel_com_2, EndEtg_ramal_com_2
	dim EndEtg_tipo_pessoa, EndEtg_cnpj_cpf, EndEtg_contribuinte_icms_status, EndEtg_produtor_rural_status
	dim EndEtg_ie, EndEtg_rg
	dim blnEndEtgComDados
    blnEndEtg_obs = false
    blnEndEtgComDados = false
	c_FlagEndEntregaEditavel = Trim(Request.Form("c_FlagEndEntregaEditavel"))
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

	dim alerta, blnErroConsistencia
	alerta=""
	blnErroConsistencia=False
		
	dim vl_total, vl_total_NF, vl_total_RA
	dim v_item, i, n, editou_item
	redim v_item(0)
	set v_item(Ubound(v_item)) = New cl_ITEM_ORCAMENTO
	v_item(Ubound(v_item)).produto = ""
	n = Request.Form("c_produto").Count
	for i = 1 to n
		s=Trim(Request.Form("c_produto")(i))
		if s <> "" then
			if Trim(v_item(ubound(v_item)).produto) <> "" then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ITEM_ORCAMENTO
				end if
			with v_item(ubound(v_item))
				.produto=Ucase(Trim(Request.Form("c_produto")(i)))
				s=retorna_so_digitos(Request.Form("c_fabricante")(i))
				.fabricante=normaliza_codigo(s, TAM_MIN_FABRICANTE)
				.obs=Trim(Request.Form("c_obs")(i))
				s=Trim(Request.Form("c_vl_unitario")(i))
				.preco_venda=converte_numero(s)
				s=Trim(Request.Form("c_vl_NF")(i))
				.preco_NF=converte_numero(s)
				s = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
				end with
			end if
		next


'	CALCULA O VALOR TOTAL
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


	dim r_cliente
	set r_cliente = New cl_CLIENTE

	dim eh_cpf
    eh_cpf = false

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim r_orcamento
	if alerta = "" then
		if Not le_orcamento(orcamento_selecionado, r_orcamento, msg_erro) then 
            alerta = msg_erro
        else
	        call x_cliente_bd(r_orcamento.id_cliente, r_cliente)
	        eh_cpf=(len(r_cliente.cnpj_cpf)=11)
		    end if
		end if

	dim blnAtivarFlag_st_violado_permite_RA_status
	blnAtivarFlag_st_violado_permite_RA_status = False
	if alerta = "" then
		if (r_orcamento.permite_RA_status = 0) And (r_orcamento.st_violado_permite_RA_status = 0) then
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

	dim rb_instalador_instala
	rb_instalador_instala = Trim(Request.Form("rb_instalador_instala"))




	dim endereco__bairro, endereco__endereco, endereco__numero, endereco__complemento, endereco__cidade, endereco__uf, endereco__cep
	dim cliente__ddd_res, cliente__tel_res, cliente__ddd_cel, cliente__tel_cel, cliente__ddd_com, cliente__tel_com, cliente__ramal_com,cliente__ddd_com_2, cliente__tel_com_2, cliente__ramal_com_2
	dim cliente__email, cliente__email_xml , cliente__nome, cliente__ie, cliente__rg, cliente__contribuinte_icms_status, cliente__produtor_rural
	
	if r_orcamento.st_memorizacao_completa_enderecos = 1 or r_orcamento.st_memorizacao_completa_enderecos = 9 then
	
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





	dim rb_garantia_indicador, GarantiaIndicadorStatusOriginal
	dim strBlnGarantiaIndicadorBloqueado, blnGarantiaIndicadorBloqueado
	GarantiaIndicadorStatusOriginal = Trim(Request.Form("GarantiaIndicadorStatusOriginal"))
	rb_garantia_indicador = Trim(Request.Form("rb_garantia_indicador"))
	strBlnGarantiaIndicadorBloqueado = Trim(Request.Form("blnGarantiaIndicadorBloqueado"))
	blnGarantiaIndicadorBloqueado = CBool(strBlnGarantiaIndicadorBloqueado)
	
	if alerta = "" then
		if rb_garantia_indicador = "" then
			if GarantiaIndicadorStatusOriginal = "" then
				alerta = "Falha ao obter o campo 'Garantia Indicador'"
			else
			'	Lembrando que pedidos antigos estão com o status COD_GARANTIA_INDICADOR_STATUS__NAO_DEFINIDO
			'	e que os radio buttons de edição ficam ambos desmarcados
				rb_garantia_indicador = GarantiaIndicadorStatusOriginal
				end if
			end if
		end if

	if alerta = "" then
		if CLng(s_etg_imediata) = CLng(COD_ETG_IMEDIATA_NAO) then
			if c_data_previsao_entrega = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "É necessário informar a data de previsão de entrega"
			elseif Not IsDate(c_data_previsao_entrega) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Data de previsão de entrega informada é inválida"
			elseif StrToDate(c_data_previsao_entrega) <= Date then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Data de previsão de entrega deve ser uma data futura"
				end if
			end if
		end if

'	FORMA DE PAGAMENTO (NOVA VERSÃO)
	if alerta = "" then
		if versao_forma_pagamento = "2" then
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
			else
				alerta = "É obrigatório especificar a forma de pagamento"
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

	dim c_loja
	c_loja = Trim(Request.Form("c_loja"))
	
'	O ORÇAMENTO FOI CADASTRADO JÁ DENTRO DA POLÍTICA DE PERCENTUAL DE CUSTO FINANCEIRO POR FORNECEDOR?
	if c_custoFinancFornecTipoParcelamentoOriginal <> "" then
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
			c_custoFinancFornecQtdeParcelasConferencia=0
			end if

		if alerta = "" then
			if c_custoFinancFornecTipoParcelamentoConferencia<>c_custoFinancFornecTipoParcelamento then
				alerta="Foi detectada uma inconsistência no tipo de parcelamento do pagamento (código esperado=" & c_custoFinancFornecTipoParcelamentoConferencia & ", código lido=" & c_custoFinancFornecTipoParcelamento & ")"
			elseif c_custoFinancFornecQtdeParcelasConferencia<>c_custoFinancFornecQtdeParcelas then
				alerta="Foi detectada uma inconsistência na quantidade de parcelas de pagamento (qtde esperada=" & c_custoFinancFornecQtdeParcelasConferencia & ", qtde lida=" & c_custoFinancFornecQtdeParcelas & ")"
				end if
			end if

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
		end if

	'Editável?
	if c_FlagEndEntregaEditavel = "S" then
		if alerta = "" then
            if (EndEtg_endereco<>r_orcamento.EndEtg_endereco) Or (EndEtg_bairro<>r_orcamento.EndEtg_bairro) Or (EndEtg_cidade<>r_orcamento.EndEtg_cidade) Or (EndEtg_uf<>r_orcamento.EndEtg_uf) Or (EndEtg_cep<>r_orcamento.EndEtg_cep) Or (EndEtg_obs<>r_orcamento.EndEtg_obs) then
                blnEndEtg_obs = true 
                end if

            'na memorizacao de endereços ligada, sempre verificamos
            if r_orcamento.st_memorizacao_completa_enderecos <> 0 and blnUsarMemorizacaoCompletaEnderecos then
                blnEndEtg_obs = true 
                end if

            blnEndEtgComDados = false
			if (EndEtg_endereco<>"") Or (EndEtg_bairro<>"") Or (EndEtg_cidade<>"") Or (EndEtg_uf<>"") Or (EndEtg_cep<>"") Or (EndEtg_obs<>"") then
                blnEndEtgComDados = true
                end if

            if r_orcamento.st_memorizacao_completa_enderecos <> 0 and blnUsarMemorizacaoCompletaEnderecos then
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
                    'nenhum campo deve ser preenchido pelo usuário
                    'todos possuem prenchimento automático
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
					alerta="PREENCHA O ENDEREÇO DE ENTREGA."
				elseif Len(EndEtg_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
					alerta="ENDEREÇO DE ENTREGA EXCEDE O TAMANHO MÁXIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(EndEtg_endereco)) & " CARACTERES<br>TAMANHO MÁXIMO: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " CARACTERES"
				elseif EndEtg_endereco_numero="" then
					alerta="PREENCHA O NÚMERO DO ENDEREÇO DE ENTREGA."
				elseif EndEtg_cidade="" then
					alerta="PREENCHA A CIDADE DO ENDEREÇO DE ENTREGA."
				elseif EndEtg_uf="" then
					alerta="PREENCHA A UF DO ENDEREÇO DE ENTREGA."
				elseif EndEtg_cep="" then
					alerta="PREENCHA O CEP DO ENDEREÇO DE ENTREGA."
                elseif (EndEtg_obs="" And blnEndEtg_obs = true)  then
                    alerta="PREENCHA A JUSTIFICATIVA DO ENDEREÇO DE ENTREGA."
					end if
				end if
			end if




			if r_orcamento.st_memorizacao_completa_enderecos = 1 or r_orcamento.st_memorizacao_completa_enderecos = 9 then
			
				if endereco__endereco="" then
					alerta="PREENCHA O ENDEREÇO."
				elseif Len(endereco__endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
					alerta="ENDEREÇO EXCEDE O TAMANHO MÁXIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(endereco__endereco)) & " CARACTERES<br>TAMANHO MÁXIMO: " & Cstr(endereco__endereco) & " CARACTERES"
				elseif endereco__numero="" then
					alerta="PREENCHA O NÚMERO DO ENDEREÇO."
				elseif endereco__cidade="" then
					alerta="PREENCHA A CIDADE DO ENDEREÇO."
				elseif endereco__uf="" then
					alerta="PREENCHA A UF DO ENDEREÇO."
				elseif endereco__cep="" then
					alerta="PREENCHA O CEP DO ENDEREÇO."	
		        elseif Not cep_ok(endereco__cep) then
			        alerta="CEP INVÁLIDO."
		        elseif Not ddd_ok(cliente__ddd_res) then
			        alerta="DADOS CADASTRAIS: DDD INVÁLIDO."
		        elseif Not telefone_ok(cliente__tel_res) then
			        alerta="DADOS CADASTRAIS: TELEFONE RESIDENCIAL INVÁLIDO."
		        elseif (cliente__ddd_res <> "") And ((cliente__tel_res = "")) then
			        alerta="DADOS CADASTRAIS: PREENCHA O TELEFONE RESIDENCIAL."
		        elseif (cliente__ddd_res = "") And ((cliente__tel_res <> "")) then
			        alerta="DADOS CADASTRAIS: PREENCHA O DDD."
		        elseif Not ddd_ok(cliente__ddd_com) then
			        alerta="DADOS CADASTRAIS: DDD INVÁLIDO."
		        elseif Not telefone_ok(cliente__tel_com) then
			        alerta="DADOS CADASTRAIS: TELEFONE COMERCIAL INVÁLIDO."
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
                        alerta = "Dados cadastrais: informe se o cliente é produtor rural ou não!!"
                    elseif converte_numero(cliente__produtor_rural) = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
                        cliente__contribuinte_icms_status = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_INICIAL
                        cliente__ie = ""
                    elseif converte_numero(cliente__produtor_rural) <> converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
                        if converte_numero(cliente__contribuinte_icms_status) <> converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                            alerta = "Dados cadastrais: para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!!"
                        elseif cliente__contribuinte_icms_status = "" then
                            alerta = "Dados cadastrais: informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!"
                        elseif converte_numero(cliente__contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and cliente__ie = "" then
                            alerta = "Dados cadastrais: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!"
                        elseif converte_numero(cliente__contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) and InStr(cliente__ie, "ISEN") > 0 then 
                            alerta = "Dados cadastrais: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!"
                        elseif converte_numero(cliente__contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and InStr(cliente__ie, "ISEN") > 0 then 
                            alerta = "Dados cadastrais: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!"
                        elseif converte_numero(cliente__contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) and cliente__ie <> "" then 
                            alerta = "Dados cadastrais: se o Contribuinte ICMS é isento, o campo IE deve ser vazio!"
                            end if
                        end if
					end if
			
				if	cliente__ie <> "" then 
					if Not isInscricaoEstadualValida(cliente__ie, endereco__uf) then
						alerta="Preencha a IE (Inscrição Estadual) com um número válido!!" & _
							"<br>" & "Certifique-se de que a UF do endereço corresponde à UF responsável pelo registro da IE."
						end if
				end if
			end if


        if alerta = "" and blnEndEtgComDados and r_orcamento.st_memorizacao_completa_enderecos <> 0 and blnUsarMemorizacaoCompletaEnderecos and Not eh_cpf then
            if EndEtg_tipo_pessoa <> "PJ" and EndEtg_tipo_pessoa <> "PF" then
                alerta = "Necessário escolher Pessoa Jurídica ou Pessoa Física no Endereço de entrega!!"
    		elseif EndEtg_nome = "" then
                alerta = "Preencha o nome/razão social no endereço de entrega!!"
                end if 
	
            if alerta = "" and EndEtg_tipo_pessoa = "PJ" then

                'limpa os números de telefone que não foram informados
                EndEtg_ddd_res = ""
                EndEtg_tel_res = ""
                EndEtg_ddd_cel = ""
                EndEtg_tel_cel = ""

                '//Campos PJ: 
                if EndEtg_cnpj_cpf = "" or not cnpj_ok(EndEtg_cnpj_cpf) then
                    alerta = "Endereço de entrega: CNPJ inválido!!"
                elseif EndEtg_contribuinte_icms_status = "" then
                    alerta = "Endereço de entrega: selecione o tipo de contribuinte de ICMS!!"
                elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and EndEtg_ie = "" then
                    alerta = "Endereço de entrega: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!"
                elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) and InStr(EndEtg_ie, "ISEN") > 0 then 
                    alerta = "Endereço de entrega: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!"
                elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and InStr(EndEtg_ie, "ISEN") > 0 then 
                    alerta = "Endereço de entrega: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!"
                'telefones PJ:
                'EndEtg_ddd_com
                'EndEtg_tel_com
                'EndEtg_ramal_com
                'EndEtg_ddd_com_2
                'EndEtg_tel_com_2
                'EndEtg_ramal_com_2
                elseif not ddd_ok(EndEtg_ddd_com) then
                    alerta = "Endereço de entrega: DDD inválido!!"
                elseif not telefone_ok(EndEtg_tel_com) then
                    alerta = "Endereço de entrega: telefone inválido!!"
                elseif EndEtg_ddd_com = "" and EndEtg_tel_com <> "" then
                    alerta = "Endereço de entrega: preencha o DDD do telefone."
                elseif EndEtg_tel_com = "" and EndEtg_ddd_com <> "" then
                    alerta = "Endereço de entrega: preencha o telefone."

                elseif not ddd_ok(EndEtg_ddd_com_2) then
                    alerta = "Endereço de entrega: DDD inválido!!"
                elseif not telefone_ok(EndEtg_tel_com_2) then
                    alerta = "Endereço de entrega: telefone inválido!!"
                elseif EndEtg_ddd_com_2 = "" and EndEtg_tel_com_2 <> "" then
                    alerta = "Endereço de entrega: preencha o DDD do telefone."
                elseif EndEtg_tel_com_2 = "" and EndEtg_ddd_com_2 <> "" then
                    alerta = "Endereço de entrega: preencha o telefone."
                    end if 
                end if 

            if alerta = "" and EndEtg_tipo_pessoa <> "PJ" then

                'limpa os números de telefone que não foram informados
                EndEtg_ddd_com = ""
                EndEtg_tel_com = ""
                EndEtg_ramal_com = ""
                EndEtg_ddd_com_2 = ""
                EndEtg_tel_com_2 = ""
                EndEtg_ramal_com_2 = ""

                '//campos PF
                if EndEtg_cnpj_cpf = "" or not cpf_ok(EndEtg_cnpj_cpf) then
                    alerta = "Endereço de entrega: CPF inválido!!"
                elseif EndEtg_produtor_rural_status = "" then
                    alerta = "Endereço de entrega: informe se o cliente é produtor rural ou não!!"
                elseif converte_numero(EndEtg_produtor_rural_status) <> converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
                    if converte_numero(EndEtg_contribuinte_icms_status) <> converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                        alerta = "Endereço de entrega: para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!!"
                    elseif EndEtg_contribuinte_icms_status = "" then
                        alerta = "Endereço de entrega: informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and EndEtg_ie = "" then
                        alerta = "Endereço de entrega: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) and InStr(EndEtg_ie, "ISEN") > 0 then 
                        alerta = "Endereço de entrega: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and InStr(EndEtg_ie, "ISEN") > 0 then 
                        alerta = "Endereço de entrega: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!"
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
                        alerta = "Endereço de entrega: DDD inválido!!"
                    elseif not telefone_ok(retorna_so_digitos(EndEtg_tel_res)) then
                        alerta = "Endereço de entrega: telefone inválido!!"
                    elseif EndEtg_ddd_res <> "" or EndEtg_tel_res <> "" then
                        if EndEtg_ddd_res = "" then
                            alerta = "Endereço de entrega: preencha o DDD!!"
                        elseif EndEtg_tel_res = "" then
                            alerta = "Endereço de entrega: preencha o telefone!!"
                            end if
                        end if
                    end if

                if alerta = "" then
                    if not ddd_ok(retorna_so_digitos(EndEtg_ddd_cel)) then
                        alerta = "Endereço de entrega: DDD inválido!!"
                    elseif not telefone_ok(retorna_so_digitos(EndEtg_tel_cel)) then
                        alerta = "Endereço de entrega: telefone inválido!!"
                    elseif EndEtg_ddd_cel = "" and EndEtg_tel_cel <> "" then
                        alerta = "Endereço de entrega: preencha o DDD do celular."
                    elseif EndEtg_tel_cel = "" and EndEtg_ddd_cel <> "" then
                        alerta = "Endereço de entrega: preencha o número do celular."
                        end if
                    end if

                end if

		    if alerta = "" and EndEtg_ie <> "" then
			    if Not isInscricaoEstadualValida(EndEtg_ie, EndEtg_uf) then
				    alerta="Endereço de entrega: preencha a IE (Inscrição Estadual) com um número válido!!" & _
						    "<br>" & "Certifique-se de que a UF do endereço de entrega corresponde à UF responsável pelo registro da IE."
				    end if
			    end if

            end if
		end if

'	CONSISTÊNCIAS P/ EMISSÃO DE NFe
	dim s_tabela_municipios_IBGE
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
		if (c_FlagEndEntregaEditavel = "S") And (EndEtg_cidade <> "") then
		'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
			dim s_lista_sugerida_municipios
			dim v_lista_sugerida_municipios
			dim iCounterLista, iNumeracaoLista
			if Not consiste_municipio_IBGE_ok(EndEtg_cidade, EndEtg_uf, s_lista_sugerida_municipios, msg_erro) then
				if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				if msg_erro <> "" then
					alerta = alerta & msg_erro
				else
					alerta = alerta & "Município '" & EndEtg_cidade & "' não consta na relação de municípios do IBGE para a UF de '" & EndEtg_uf & "'!!"
					if s_lista_sugerida_municipios <> "" then
						alerta = alerta & "<br>" & _
										  "Localize o município na lista abaixo e verifique se a grafia está correta!!"
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
									"			<p class='N'>" & "Relação de municípios de '" & EndEtg_uf & "' que se iniciam com a letra '" & Ucase(left(EndEtg_cidade,1)) & "'" & "</p>" & chr(13) & _
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
			end if 'if (c_FlagEndEntregaEditavel = "S") And (EndEtg_cidade <> "")
		end if 'if alerta = ""
	
'	CONSISTÊNCIAS P/ EMISSÃO DE NFe (DADOS CADASTRAIS)
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
		if r_orcamento.st_memorizacao_completa_enderecos <> 0 then
		'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
			if Not consiste_municipio_IBGE_ok(endereco__cidade, endereco__uf, s_lista_sugerida_municipios, msg_erro) then
				if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				if msg_erro <> "" then
					alerta = alerta & msg_erro
				else
					alerta = alerta & "Município '" & endereco__cidade & "' não consta na relação de municípios do IBGE para a UF de '" & endereco__uf & "'!!"
					if s_lista_sugerida_municipios <> "" then
						alerta = alerta & "<br>" & _
										  "Localize o município na lista abaixo e verifique se a grafia está correta!!"
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
									"			<p class='N'>" & "Relação de municípios de '" & endereco__uf & "' que se iniciam com a letra '" & Ucase(left(endereco__cidade,1)) & "'" & "</p>" & chr(13) & _
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
			end if 'if r_orcamento.st_memorizacao_completa_enderecos <> 0 
		end if 'if alerta = ""


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

	if alerta <> "" then blnErroConsistencia=True
	
	
'	GRAVA NO BANCO DE DADOS
'	=======================
	dim rs, rs2
	
	dim msg_erro
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
	dim s_log
	dim campos_a_omitir
	s_log = ""
	campos_a_omitir = "|etg_imediata_data|etg_imediata_usuario|"
	campos_a_omitir_ItemCFF = ""
	campos_a_omitir_PedCFF = ""
	
	if alerta = "" then
	'	ATUALIZA O ORÇAMENTO
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO then
		'	BLOQUEIA REGISTRO PARA EVITAR ACESSO CONCORRENTE (REALIZA O FLIP EM UM CAMPO BIT APENAS P/ ADQUIRIR O LOCK EXCLUSIVO)
		'	OBS: TODOS OS MÓDULOS DO SISTEMA QUE REALIZEM ESTA OPERAÇÃO DE CADASTRAMENTO DEVEM SINCRONIZAR O ACESSO OBTENDO O LOCK EXCLUSIVO DO REGISTRO DE CONTROLE DESIGNADO
			s = "UPDATE t_CONTROLE SET" & _
					" dummy = ~dummy" & _
				" WHERE" & _
					" id_nsu = '" & ID_XLOCK_SYNC_ORCAMENTO & "'"
			cn.Execute(s)
			end if

		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

		s_etg_imediata_original = ""
		if alerta = "" then
			s = "SELECT * FROM t_ORCAMENTO WHERE orcamento='" & orcamento_selecionado & "'"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Err <> 0 then
				alerta = Cstr(Err) & ": " & Err.Description
			elseif rs.EOF then
				alerta = "Orçamento " & orcamento_selecionado & " não foi encontrado."
			else
				log_via_vetor_carrega_do_recordset rs, vLog1, campos_a_omitir

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

				if s_bem_uso_consumo <> "" then 
					if CLng(rs("StBemUsoConsumo")) <> CLng(s_bem_uso_consumo) then
						rs("StBemUsoConsumo")=CLng(s_bem_uso_consumo)
						end if
					end if

				if rb_instalador_instala <> "" then
					if CLng(rs("InstaladorInstalaStatus")) <> CLng(rb_instalador_instala) then
						rs("InstaladorInstalaStatus")=CLng(rb_instalador_instala)
						rs("InstaladorInstalaUsuarioUltAtualiz")=usuario
						rs("InstaladorInstalaDtHrUltAtualiz")=Now
						end if
					end if
					
				if Not blnGarantiaIndicadorBloqueado then
					if CLng(rs("GarantiaIndicadorStatus")) <> CLng(rb_garantia_indicador) then
						rs("GarantiaIndicadorStatus")=CLng(rb_garantia_indicador)
						rs("GarantiaIndicadorUsuarioUltAtualiz")=usuario
						rs("GarantiaIndicadorDtHrUltAtualiz")=Now
						end if
					end if

			'	Forma de Pagamento (nova versão)
				if versao_forma_pagamento = "2" then
					rs("tipo_parcelamento")=CLng(rb_forma_pagto)
				'	Limpa os campos não usados p/ facilitar a consulta multicritério e p/ que na próxima alteração, se houver, o log perceba a alteração (ex: parcelado c/ entrada mudou p/ parcelado no cartão; ao alterar novamente p/ parcelado c/ entrada e forem preenchidos os mesmos valores, será percebida e registrada apenas a mudança da opção "parcelado c/ entrada", já os demais campos ficaram c/ os mesmos valores).
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
					end if
				
				rs("obs_1") = s_obs1
				rs("obs_2") = s_obs2
				rs("forma_pagto") = s_forma_pagto
				
				if versao_forma_pagamento = "1" then
					if IsNumeric(s_qtde_parcelas) then 
						rs("qtde_parcelas") = CLng(s_qtde_parcelas)
					else
						rs("qtde_parcelas") = 0
						end if
					end if
					
				rs("vl_total") = vl_total
				rs("vl_total_NF") = vl_total_NF
				rs("vl_total_RA") = vl_total_RA
				rs("perc_RT") = converte_numero(s_perc_RT)





				if r_orcamento.st_memorizacao_completa_enderecos = 1 or r_orcamento.st_memorizacao_completa_enderecos = 9 then
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








				
				'Editável?
				if c_FlagEndEntregaEditavel = "S" then
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
                	if r_orcamento.st_memorizacao_completa_enderecos <> 0 and blnUsarMemorizacaoCompletaEnderecos then
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
					end if
				end if
			end if
			
		if alerta = "" then
		'	O ORÇAMENTO FOI CADASTRADO JÁ DENTRO DA POLÍTICA DE PERCENTUAL DE CUSTO FINANCEIRO POR FORNECEDOR?
			if c_custoFinancFornecTipoParcelamentoOriginal <> "" then
				if Not blnFormaPagtoBloqueado then
					if (c_custoFinancFornecTipoParcelamentoOriginal <> c_custoFinancFornecTipoParcelamento) Or _
					   (c_custoFinancFornecQtdeParcelasOriginal <> c_custoFinancFornecQtdeParcelas) then
						for i=Lbound(v_item) to Ubound(v_item)
							with v_item(i)
								if Trim(.produto)<>"" then
								'	Inicialização
									vlCustoFinancFornecPrecoListaBase = 0
									coeficiente = 0
									
								'	Obtém Preço de Lista Base
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
										alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " NÃO está cadastrado para a loja " & c_loja
									else
										vlCustoFinancFornecPrecoListaBase = rs2("preco_lista")
										end if
								
								'	Obtém coeficiente do custo financeiro
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
											end if
										end if
									
									if alerta = "" then
										s = "SELECT " & _
												"*" & _
											" FROM t_ORCAMENTO_ITEM" & _
											" WHERE" & _
												" (orcamento='" & orcamento_selecionado & "') AND" & _
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
								" FROM t_ORCAMENTO" & _
								" WHERE" & _
								" orcamento='" & orcamento_selecionado & "'"
							if rs.State <> 0 then rs.Close
							rs.Open s, cn
							if rs.Eof then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Falha ao localizar o registro do orçamento " & orcamento_selecionado
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
							" FROM t_ORCAMENTO_ITEM" & _
							" WHERE" & _
								" (orcamento='" & orcamento_selecionado & "') AND" & _
								" (fabricante='" & Trim(.fabricante) & "') AND" & _
								" (produto='" & Trim(.produto) & "')"
						if rs.State <> 0 then rs.Close
						rs.Open s, cn
						if Err <> 0 then
							alerta=texto_add_br(alerta)
							alerta=alerta & Cstr(Err) & ": " & Err.Description
						elseif rs.EOF then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Item do orçamento do produto " & Trim(.produto) & " (fabricante " & Trim(.fabricante) & ") não foi encontrado."
						else
							editou_item = False
						'	EDITOU "OBSERVAÇÕES"?
							if Trim("" & rs("obs")) <> .obs then
								if s_log <> "" then s_log = s_log & "; "
								s_log = s_log & "observações do produto " & Trim(.produto) & " (" & Trim(.fabricante) & "): " & formata_texto_log(Trim("" & rs("obs"))) & " => " & formata_texto_log(.obs)
								rs("obs")=.obs
								editou_item = True
								end if
							
						'	EDITOU PREÇO DE VENDA?
							if rs("preco_venda") <> .preco_venda then
								if s_log <> "" then s_log = s_log & "; "
								s_log = s_log & "preço de venda do produto " & Trim(.produto) & " (" & Trim(.fabricante) & "): " & formata_moeda(rs("preco_venda")) & " => " & formata_moeda(.preco_venda)
								.preco_lista=rs("preco_lista")
								if .preco_lista = 0 then
									.desc_dado=0
								else
									.desc_dado=100*(.preco_lista-.preco_venda)/.preco_lista
									end if
								rs("preco_venda")=.preco_venda
								rs("desc_dado")=.desc_dado
								editou_item = True
								end if
							
						'	EDITOU PREÇO DE NF?
							if rs("preco_NF") <> .preco_NF then
								if s_log <> "" then s_log = s_log & "; "
								s_log = s_log & "preço de NF do produto " & Trim(.produto) & " (" & Trim(.fabricante) & "): " & formata_moeda(rs("preco_NF")) & " => " & formata_moeda(.preco_NF)
								rs("preco_NF")=.preco_NF
								editou_item = True
								end if
								
							if editou_item then rs.Update
							end if
						end if
					end with
				next
			end if
			
		if alerta = "" then
		'	CUSTO FINANCEIRO POR FORNECEDOR
			if (s_log <> "") And (s_log_PedCFF <> "") then s_log = s_log & "; "
			s_log = s_log & s_log_PedCFF
			if (s_log <> "") And (s_log_ItemCFF <> "") then s_log = s_log & "; "
			s_log = s_log & s_log_ItemCFF

		'	GRAVA O LOG!!
			if s_log <> "" then grava_log usuario, "", orcamento_selecionado, "", OP_LOG_ORCAMENTO_ALTERACAO, s_log
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("orcamento.asp?orcamento_selecionado=" & orcamento_selecionado & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
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
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
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