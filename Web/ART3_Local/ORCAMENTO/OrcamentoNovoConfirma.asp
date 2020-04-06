<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<%
'     =================================================
'	  O R C A M E N T O N O V O C O N F I R M A . A S P
'     =================================================
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

	const blnConsisteDescontoMax = False
	
	dim usuario, loja, cliente_selecionado, midia_selecionada, vendedor_selecionado, s_perc_RT
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim alerta
	alerta=""

	s_perc_RT = Trim(request("c_perc_RT"))
	midia_selecionada = Trim(request("midia"))
	vendedor_selecionado = Trim(request("vendedor"))
	cliente_selecionado = Trim(request("cliente_selecionado"))
	if (cliente_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)

	dim rb_instalador_instala, rb_garantia_indicador
	rb_instalador_instala = Trim(Request.Form("rb_instalador_instala"))
	rb_garantia_indicador = Trim(Request.Form("rb_garantia_indicador"))
	
'	FORMA DE PAGAMENTO (NOVA VERSÃO)
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
		c_custoFinancFornecQtdeParcelasConferencia=0
		end if

	if alerta = "" then
		if c_custoFinancFornecTipoParcelamentoConferencia<>c_custoFinancFornecTipoParcelamento then
			alerta="Foi detectada uma inconsistência no tipo de parcelamento do pagamento (código esperado=" & c_custoFinancFornecTipoParcelamentoConferencia & ", código lido=" & c_custoFinancFornecTipoParcelamento & ")"
		elseif converte_numero(c_custoFinancFornecQtdeParcelasConferencia)<>converte_numero(c_custoFinancFornecQtdeParcelas) then
			alerta="Foi detectada uma inconsistência na quantidade de parcelas de pagamento (qtde esperada=" & c_custoFinancFornecQtdeParcelasConferencia & ", qtde lida=" & c_custoFinancFornecQtdeParcelas & ")"
			end if
		end if

	dim rb_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs
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

	dim s, i, k, n, opcao_venda_sem_estoque, qtde_spe, vl_total, vl_total_NF, vl_total_RA
	dim v_desconto()
	ReDim v_desconto(0)
	v_desconto(UBound(v_desconto)) = ""

	opcao_venda_sem_estoque = Trim(request("opcao_venda_sem_estoque"))
	
	dim s_forma_pagto, s_obs1, s_obs2, s_etg_imediata, s_bem_uso_consumo
	s_obs1=Trim(request("c_obs1"))
	s_obs2=Trim(request("c_obs2"))
	s_etg_imediata=Trim(request("rb_etg_imediata"))
	s_bem_uso_consumo=Trim(request("rb_bem_uso_consumo"))
	s_forma_pagto=Trim(request("c_forma_pagto"))

	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim r_orcamentista_e_indicador
	if alerta = "" then
		if Not le_orcamentista_e_indicador(usuario, r_orcamentista_e_indicador, msg_erro) then
			alerta = "Falha ao recuperar os dados cadastrais!!"
			end if
		end if

	dim v_item
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
				s = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
				s=Trim(Request.Form("c_vl_unitario")(i))
				.preco_venda=converte_numero(s)
				if r_orcamentista_e_indicador.permite_RA_status = 1 then
					s=Trim(Request.Form("c_vl_NF")(i))
					.preco_NF=converte_numero(s)
				else
					.preco_NF = .preco_venda
					end if
				.obs=Trim(Request.Form("c_obs")(i))
				end with
			end if
		next
	
'	VERIFICA SE ESTE ORÇAMENTO JÁ FOI GRAVADO!!
	dim orcamento_a, vjg
	s = "SELECT t_ORCAMENTO.orcamento, fabricante, produto, qtde, preco_venda FROM t_ORCAMENTO INNER JOIN t_ORCAMENTO_ITEM ON (t_ORCAMENTO.orcamento=t_ORCAMENTO_ITEM.orcamento)" & _
		" WHERE (id_cliente='" & cliente_selecionado & "') AND (data=" & bd_formata_data(Date) & ")" & _
		" AND (loja='" & loja & "') AND (orcamentista='" & usuario & "')" & _
		" AND (hora>='" & formata_hora_hhnnss(Now-converte_min_to_dec(10))& "')" & _
		" ORDER BY t_ORCAMENTO_ITEM.orcamento, sequencia"
	set rs = cn.execute(s)
	redim vjg(0)
	set vjg(ubound(vjg)) = New cl_DUAS_COLUNAS
	vjg(ubound(vjg)).c1=""
	orcamento_a="--XX--"
	do while Not rs.EOF 
		if orcamento_a<>Trim("" & rs("orcamento")) then
			orcamento_a=Trim("" & rs("orcamento"))
			if vjg(ubound(vjg)).c1 <> "" then 
				redim preserve vjg(ubound(vjg)+1)
				set vjg(ubound(vjg)) = New cl_DUAS_COLUNAS
				vjg(ubound(vjg)).c1=""
				end if
			vjg(ubound(vjg)).c2=orcamento_a
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
			alerta="Este pré-pedido já foi gravado com o número " & vjg(i).c2
			exit for
			end if
		next

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
	

'	VERIFICA CADA UM DOS PRODUTOS SELECIONADOS
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
                    .subgrupo = Trim("" & rs("subgrupo"))
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
					else
						.desc_dado = 100*(.preco_lista-.preco_venda)/.preco_lista
						end if
					
					if blnConsisteDescontoMax then
						if .desc_dado > .desc_max then
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
								alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": desconto de " & formata_perc_desc(.desc_dado) & "% excede o máximo permitido."
							else
								if .desc_dado > rs("desc_max") then
									alerta=texto_add_br(alerta)
									alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": desconto de " & formata_perc_desc(.desc_dado) & "% excede o máximo permitido."
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
					end if
				rs.Close
				end with
			next
		end if

'	RECUPERA OS PRODUTOS QUE O CLIENTE CONCORDOU EM FAZER ORÇAMENTO MESMO SEM PRESENÇA NO ESTOQUE.
	dim v_spe
	redim v_spe(0)
	set v_spe(0) = New cl_ESTOQUE_VERIFICA_DISPONIBILIDADE_PRODUTO
	if (alerta="") And (opcao_venda_sem_estoque<>"") then
		n=Request.Form("c_spe_produto").Count
		for i=1 to n
			s=Trim(Request.Form("c_spe_produto")(i))
			if s<>"" then
				if Trim(v_spe(ubound(v_spe)).produto) <> "" then
					redim preserve v_spe(ubound(v_spe)+1)
					set v_spe(ubound(v_spe)) = New cl_ESTOQUE_VERIFICA_DISPONIBILIDADE_PRODUTO
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
	
	
'	OBTÉM O VEÍCULO DE MÍDIA (SE NENHUM FOI INFORMADO) / CONSISTE CEP
	if alerta = "" then
		s = "SELECT id, midia, cep FROM t_CLIENTE WHERE (id='" & cliente_selecionado & "')"
		set rs = cn.execute(s)
		if Not rs.Eof then
			if midia_selecionada = "" then midia_selecionada = Trim("" & rs("midia"))
			if Trim("" & rs("cep")) = "" then alerta = "É necessário preencher o CEP no cadastro do cliente."
			end if
		if rs.State <> 0 then rs.Close
		end if
	
	
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

	if alerta = "" then
		if s_etg_imediata = "" then
			alerta = "É necessário selecionar uma opção para o campo 'Entrega Imediata'."
			end if
		end if

	if alerta = "" then
		if s_bem_uso_consumo = "" then
			alerta = "É necessário informar se é 'Bem de Uso/Consumo'."
			end if
		end if

	if alerta = "" then
		if rb_instalador_instala = "" then
			alerta = "É necessário preencher o campo 'Instalador Instala'."
			end if
		end if

	if alerta = "" then
		if rb_garantia_indicador = "" then
			alerta = "Informe se o pedido será garantido pelo indicador ou não."
			end if
		end if
	
'	CEP
	if alerta = "" then
		if rb_end_entrega = "S" then
			if EndEtg_cep = "" then
				alerta = "Informe o CEP do endereço de entrega."
				end if
			end if
		end if
	
'	CONSISTÊNCIA DO VALOR TOTAL DA FORMA DE PAGAMENTO
	if alerta = "" then
		if rb_forma_pagto = COD_FORMA_PAGTO_A_VISTA then vlTotalFormaPagto = vl_total_NF
		if Abs(vlTotalFormaPagto-vl_total_NF) > 0.1 then
			alerta = "Há divergência entre o valor total do pré-pedido (" & SIMBOLO_MONETARIO & " " & formata_moeda(vl_total_NF) & ") e o valor total descrito através da forma de pagamento (" & SIMBOLO_MONETARIO & " " & formata_moeda(vlTotalFormaPagto) & ")!!"
			end if
		end if
	
	if alerta = "" then
		if rb_end_entrega = "S" then
			if (EndEtg_endereco<>"") Or (EndEtg_bairro<>"") Or (EndEtg_cidade<>"") Or (EndEtg_uf<>"") Or (EndEtg_cep<>"") Or (EndEtg_obs<>"") then
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
                elseif EndEtg_obs="" then
					alerta="PREENCHA A JUSTIFICATIVA DO ENDEREÇO DE ENTREGA."
					end if
				end if
			end if
		end if
		

'	CADASTRA O ORÇAMENTO
	if alerta="" then
		dim id_orcamento, id_orcamento_temp, s_log, msg_erro
		s_log=""
		if Not gera_num_orcamento_temp(id_orcamento_temp, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
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

		s = "SELECT * FROM t_ORCAMENTO WHERE orcamento='X'"
		rs.Open s, cn
		rs.AddNew 
		rs("orcamento")=id_orcamento_temp
		rs("loja")=loja
		rs("data")=Date
		rs("hora")=retorna_so_digitos(formata_hora(Now))
		rs("id_cliente")=cliente_selecionado
		rs("midia")=midia_selecionada
		rs("servicos")=""
		rs("orcamentista")=usuario
		rs("vendedor")=vendedor_selecionado
		rs("st_orcamento")=""
		rs("st_fechamento")=""
		if s_etg_imediata <> "" then 
			rs("st_etg_imediata")=CLng(s_etg_imediata)
			rs("etg_imediata_data")=Now
			rs("etg_imediata_usuario")=usuario
			end if
		if s_bem_uso_consumo <> "" then 
			rs("StBemUsoConsumo")=CLng(s_bem_uso_consumo)
			end if
		rs("obs_1")=s_obs1
		rs("obs_2")=s_obs2
		rs("forma_pagto")=s_forma_pagto
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

	'	CUSTO FINANCEIRO FORNECEDOR
		rs("custoFinancFornecTipoParcelamento") = c_custoFinancFornecTipoParcelamento
		rs("custoFinancFornecQtdeParcelas") = c_custoFinancFornecQtdeParcelas
		
		rs("vl_total") = vl_total
		rs("vl_total_NF") = vl_total_NF
		rs("vl_total_RA") = vl_total_RA
		rs("perc_RT") = converte_numero(s_perc_RT)
		
		rs("InstaladorInstalaStatus")=CLng(rb_instalador_instala)
		rs("InstaladorInstalaUsuarioUltAtualiz")=usuario
		rs("InstaladorInstalaDtHrUltAtualiz")=Now
		
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
		
		rs("perc_desagio_RA_liquida") = getParametroPercDesagioRALiquida
		rs("permite_RA_status") = r_orcamentista_e_indicador.permite_RA_status

		rs("sistema_responsavel_cadastro") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
		rs("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
		
		rs.Update 
		if Err <> 0 then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			end if
	'	Valor Total
		s_log = "vl total=" & formata_moeda(vl_total)
		s_log = s_log & "; vl_total_NF=" & formata_moeda(rs("vl_total_NF"))
		s_log = s_log & "; vl_total_RA=" & formata_moeda(rs("vl_total_RA"))
		s_log = s_log & "; qtde_parcelas=" & formata_texto_log(rs("qtde_parcelas"))
		s_log = s_log & "; perc_RT=" & formata_texto_log(rs("perc_RT"))
		s_log = s_log & "; midia=" & formata_texto_log(rs("midia"))
		if Trim("" & rs("forma_pagto"))<>"" then s_log = s_log & "; forma_pagto=" & formata_texto_log(rs("forma_pagto"))
		if Trim("" & rs("servicos"))<>"" then s_log = s_log & "; servicos=" & formata_texto_log(rs("servicos")) 
		if (Trim("" & rs("vl_servicos"))<>"") And (Trim("" & rs("vl_servicos"))<>"0") then s_log = s_log & "; vl_servicos=" & formata_texto_log(rs("vl_servicos")) 
		if Trim("" & rs("st_etg_imediata"))<> "" then s_log = s_log & "; st_etg_imediata=" & formata_texto_log(rs("st_etg_imediata")) 
		if Trim("" & rs("StBemUsoConsumo"))<> "" then s_log = s_log & "; StBemUsoConsumo=" & formata_texto_log(rs("StBemUsoConsumo")) 
		if Trim("" & rs("obs_1"))<>"" then s_log = s_log & "; obs_1=" & formata_texto_log(rs("obs_1")) 
		if Trim("" & rs("obs_2"))<>"" then s_log = s_log & "; obs_2=" & formata_texto_log(rs("obs_2"))
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
		
		if rb_end_entrega = "S" then
			s_log = s_log & "; Endereço entrega=" & formata_endereco(EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep) & " [EndEtg_cod_justificativa=" & EndEtg_obs & "]"
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
			s_log = s_log & "; Endereço entrega=mesmo do cadastro"
			end if
		
		s_log = s_log & "; InstaladorInstalaStatus=" & rb_instalador_instala
		s_log = s_log & "; GarantiaIndicadorStatus=" & rb_garantia_indicador
		s_log = s_log & "; perc_desagio_RA_liquida=" & rs("perc_desagio_RA_liquida")
		
		if rs.State <> 0 then rs.Close
		
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				qtde_spe = 0
				for k=Lbound(v_spe) to Ubound(v_spe)
					if (v_spe(k).fabricante=.fabricante) And (v_spe(k).produto=.produto) then
						if v_spe(k).qtde_solicitada > v_spe(k).qtde_estoque then qtde_spe = v_spe(k).qtde_solicitada - v_spe(k).qtde_estoque
						exit for
						end if
					next

				s="SELECT * FROM t_ORCAMENTO_ITEM WHERE orcamento='X'"
				rs.Open s, cn
				rs.AddNew 
				rs("orcamento")=id_orcamento_temp
				rs("fabricante")=.fabricante
				rs("produto")=.produto
				rs("qtde")=.qtde
				rs("qtde_spe")=qtde_spe
				rs("desc_dado")=.desc_dado
				rs("preco_venda")=.preco_venda
				rs("preco_NF")=.preco_NF
				rs("preco_fabricante")=.preco_fabricante
				rs("vl_custo2")=.vl_custo2
				rs("preco_lista")=.preco_lista
				rs("margem")=.margem
				rs("desc_max")=.desc_max
				rs("comissao")=.comissao
				rs("descricao")=.descricao
				rs("descricao_html")=.descricao_html
				rs("obs")=.obs
				rs("ean")=.ean
				rs("grupo")=.grupo
                rs("subgrupo")=.subgrupo
				rs("peso")=.peso
				rs("qtde_volumes")=.qtde_volumes
				rs("abaixo_min_status")=.abaixo_min_status
				rs("abaixo_min_autorizacao")=.abaixo_min_autorizacao
				rs("abaixo_min_autorizador")=.abaixo_min_autorizador
				rs("abaixo_min_superv_autorizador")=.abaixo_min_superv_autorizador
				rs("sequencia")=renumera_com_base1(Lbound(v_item), i)
				rs("markup_fabricante")=.markup_fabricante
				rs("custoFinancFornecCoeficiente")=.custoFinancFornecCoeficiente
				rs("custoFinancFornecPrecoListaBase")=.custoFinancFornecPrecoListaBase
				rs("cubagem")=.cubagem
				rs("ncm")=.ncm
				rs("cst")=.cst
				rs("descontinuado")=.descontinuado
				rs.Update
				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
				if rs.State <> 0 then rs.Close

				if s_log <> "" then s_log=s_log & ";" & chr(13)
				s_log = s_log & _
						log_produto_monta(.qtde, .fabricante, .produto) & _
						"; preco_lista=" & formata_texto_log(.preco_lista) & _
						"; desc_dado=" & formata_texto_log(.desc_dado) & _
						"; preco_venda=" & formata_texto_log(.preco_venda) & _
						"; preco_NF=" & formata_texto_log(.preco_NF) & _
						"; obs=" & formata_texto_log(.obs) & _
						"; custoFinancFornecCoeficiente=" & formata_texto_log(.custoFinancFornecCoeficiente) & _
						"; custoFinancFornecPrecoListaBase=" & formata_texto_log(.custoFinancFornecPrecoListaBase)
				
				if qtde_spe > 0 then s_log = s_log & "; spe=" & Cstr(qtde_spe)
					
				if converte_numero(.abaixo_min_status) <> 0 then
					s_log = s_log & _
							"; abaixo_min_status=" & formata_texto_log(.abaixo_min_status) & _
							"; abaixo_min_autorizacao=" & formata_texto_log(.abaixo_min_autorizacao) & _
							"; abaixo_min_autorizador=" & formata_texto_log(.abaixo_min_autorizador) & _
							"; abaixo_min_superv_autorizador=" & formata_texto_log(.abaixo_min_superv_autorizador)
					end if
				end with
			next
		
		if Not gera_num_orcamento(id_orcamento, msg_erro) then 
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
			end if
		
		s="UPDATE t_ORCAMENTO SET orcamento='" & id_orcamento & "' WHERE orcamento='" & id_orcamento_temp & "'"
		cn.Execute(s)
		
		s="UPDATE t_ORCAMENTO_ITEM SET orcamento='" & id_orcamento & "' WHERE orcamento='" & id_orcamento_temp & "'"
		cn.Execute(s)
		
		grava_log usuario, loja, id_orcamento, cliente_selecionado, OP_LOG_ORCAMENTO_NOVO, s_log		

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
		
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("orcamento.asp?orcamento_selecionado=" & id_orcamento & "&url_back=X")
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
	<title><%=TITULO_JANELA_MODULO_ORCAMENTO%></title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>



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
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>