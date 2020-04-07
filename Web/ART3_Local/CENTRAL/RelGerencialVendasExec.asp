<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L G E R E N C I A L V E N D A S E X E C . A S P
'     ======================================================
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
	
	Const COD_CONSULTA_POR_PERIODO_CADASTRO = "CADASTRO"
	Const COD_CONSULTA_POR_PERIODO_ENTREGA = "ENTREGA"
	Const COD_SAIDA_REL_FABRICANTE = "FABRICANTE"
	Const COD_SAIDA_REL_PRODUTO = "PRODUTO"
	Const COD_SAIDA_REL_VENDEDOR = "VENDEDOR"
	Const COD_SAIDA_REL_INDICADOR = "INDICADOR"
	Const COD_SAIDA_REL_UF = "UF"
	Const COD_SAIDA_REL_INDICADOR_UF = "INDICADOR_UF"
	Const COD_SAIDA_REL_CIDADE_UF = "CIDADE_UF"
    Const COD_SAIDA_REL_ORIGEM_PEDIDO = "ORIGEM_PEDIDO"
    Const COD_SAIDA_REL_LOJA = "LOJA"
    Const COD_SAIDA_REL_EMPRESA = "EMPRESA"
    Const COD_SAIDA_REL_GRUPO_PRODUTO = "GRUPO_PRODUTO"

    class cl_TOT_GRUPO
		dim ha_dados
		dim codigo_grupo
		dim descricao_grupo
		dim qtde
		dim vl_saida
		dim vl_entrada
		dim vl_lucro_bruto
		dim perc_lucro_bruto_total
		dim vl_RT
		dim perc_comissao_fat
		dim vl_lucro_liq
		dim perc_lucro_liq_fat
		end class
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_GERENCIAL_DE_VENDAS, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, lista_loja, s_filtro_loja, v_loja, v, i, flag_ok
	dim c_dt_entregue_inicio, c_dt_entregue_termino, c_dt_cadastro_inicio, c_dt_cadastro_termino
	dim c_loja, c_fabricante, c_produto, c_vendedor, c_indicador, c_captador, c_cnpj_cpf, rb_tipo_cliente, ckb_contribuinte_icms_nao, ckb_contribuinte_icms_sim, ckb_contribuinte_icms_isento, c_grupo, c_subgrupo
    dim c_pedido_origem, c_grupo_pedido_origem
	dim c_loc_uf, c_loc_digitada, c_loc_escolhidas, s_where_loc, v_loc_escolhidas,c_uf_saida
    dim c_empresa
	dim s_nome_vendedor
	dim op_forma_pagto, c_forma_pagto_qtde_parc
	dim rb_periodo, rb_saida
    dim v_grupo_pedido_origem, v_pedido_origem
    dim ckb_ordenar_marg_contrib
	dim c_cst, v_cst, v_cst_aux

	alerta = ""

	rb_periodo = Trim(Request.Form("rb_periodo"))
	rb_saida = Trim(Request.Form("rb_saida"))

	c_dt_cadastro_inicio = ""
	c_dt_cadastro_termino = ""
	c_dt_entregue_inicio = ""
	c_dt_entregue_termino = ""

	if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
		c_dt_cadastro_inicio = Trim(Request.Form("c_dt_cadastro_inicio"))
		c_dt_cadastro_termino = Trim(Request.Form("c_dt_cadastro_termino"))
	elseif rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
		c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
		c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))
		end if
		
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	c_indicador = Ucase(Trim(Request.Form("c_indicador")))
	c_captador = Ucase(Trim(Request.Form("c_captador")))
	c_cnpj_cpf = retorna_so_digitos(Trim(Request.Form("c_cnpj_cpf")))
    c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_subgrupo = Ucase(Trim(Request.Form("c_subgrupo")))
	rb_tipo_cliente = Trim(Request.Form("rb_tipo_cliente"))
	ckb_contribuinte_icms_nao = Trim(Request.Form("ckb_contribuinte_icms_nao"))
	ckb_contribuinte_icms_sim = Trim(Request.Form("ckb_contribuinte_icms_sim"))
	ckb_contribuinte_icms_isento = Trim(Request.Form("ckb_contribuinte_icms_isento"))
	c_forma_pagto_qtde_parc = retorna_so_digitos(Trim(Request.Form("c_forma_pagto_qtde_parc")))
    c_empresa = Trim(Request.Form("c_empresa"))
    c_pedido_origem = Trim(Request.Form("c_pedido_origem"))
    c_grupo_pedido_origem = Trim(Request.Form("c_grupo_pedido_origem"))
    
	c_loja = Trim(Request.Form("c_loja"))
	lista_loja = substitui_caracteres(c_loja,chr(10),"")
	v_loja = split(lista_loja,chr(13),-1)
    ckb_ordenar_marg_contrib = Trim(Request.Form("ckb_ordenar_marg_contrib"))
	c_cst = Trim(Request.Form("c_cst"))

'	NORMALIZA A LISTA DE CST (SEPARADOR ADOTADO É O ESPAÇO EM BRANCO)
	redim v_cst(0)
	v_cst(Ubound(v_cst)) = ""
	if c_cst <> "" then
		c_cst = Replace(c_cst, ",", " ")
		c_cst = Replace(c_cst, ";", " ")
		c_cst = Replace(c_cst, "-", " ")
		c_cst = Replace(c_cst, "/", " ")
	'	REMOVE ESPAÇOS EM BRANCO DUPLICADOS
		do while Instr(c_cst, "  ") <> 0
			c_cst = Replace(c_cst, "  ", " ")
			loop
		
		v_cst_aux = Split(c_cst, " ")
		for i=Lbound(v_cst_aux) to Ubound(v_cst_aux)
			if Trim("" & v_cst_aux(i)) <> "" then
				if Trim("" & v_cst(Ubound(v_cst))) <> "" then
					redim preserve v_cst(Ubound(v_cst)+1)
					end if
				v_cst(Ubound(v_cst)) = Trim("" & v_cst_aux(i))
				end if
			next
		end if
	
	op_forma_pagto = ""
	if (True) Or (rb_saida = COD_SAIDA_REL_UF) Or (rb_saida = COD_SAIDA_REL_VENDEDOR) then
		op_forma_pagto = Trim(Request.Form("op_forma_pagto"))
		end if
	
	if (rb_saida = COD_SAIDA_REL_CIDADE_UF) then
		c_loc_uf = Ucase(Trim(Request.Form("c_loc_uf")))
		c_loc_digitada = Ucase(Trim(Request.Form("c_loc_digitada")))
		c_loc_escolhidas = Ucase(Trim(Request.Form("c_loc_escolhidas")))
    else
        c_uf_saida= Trim(Request.Form("c_uf_saida"))
	end if
	
	if alerta = "" then
		if rb_periodo = "" then
			alerta = "Selecione o tipo de consulta: 'por pedidos cadastrados' ou 'por pedidos entregues'"
		elseif (rb_periodo <> COD_CONSULTA_POR_PERIODO_CADASTRO) and (rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA) then
			alerta = "Opção inválida para tipo de período de consulta."
			end if
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
		s_nome_vendedor = ""
		if c_vendedor <> "" then
			s = "SELECT nome FROM t_USUARIO WHERE (usuario='" & c_vendedor & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "VENDEDOR " & c_vendedor & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_vendedor = Ucase(Trim("" & rs("nome")))
				end if
			end if
		end if
		
		
'	Período de consulta está restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
	'	PERÍODO DE CADASTRO
		if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
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
			end if
		
	'	PERÍODO DE ENTREGA
		if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
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
		
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if

' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' inicializa_cl_TOT_GRUPO
'
sub inicializa_cl_TOT_GRUPO(byref rv)
	rv.ha_dados = False
	rv.codigo_grupo = ""
	rv.descricao_grupo = ""
	rv.qtde = 0
	rv.vl_saida = 0
	rv.vl_entrada = 0
	rv.vl_lucro_bruto = 0
	rv.perc_lucro_bruto_total = 0
	rv.vl_RT = 0
	rv.perc_comissao_fat = 0
	rv.vl_lucro_liq = 0
	rv.perc_lucro_liq_fat = 0
end sub


' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim intLargCodProduto, intLargDescrProduto, intLargQtdeProduto, intLargVendedor, intLargCidade , intLargCodFabricante ,intLargNomeFabricante
dim intLargIndicador, intLargUF, intLargColPerc, intLargColMonetario, intLargLoja
dim r
dim s, s_aux, s_where, s_where_aux, s_where_venda, s_where_devolucao, s_where_loja, s_where_cst, s_cor, s_nome_loja
dim s_sql, cab_table, cab, n_reg_BD, n_reg_vetor, x
dim vl_total_final_venda, vl_total_venda
dim percFatVenda, percFinalFatVenda, percRTFatVenda, percRTEDesc
dim percFinalDesc, percFinalRTFatVenda, percFinalRTEDescFatVenda
dim i, v, intQtdeTotalProdutos
dim strSqlCampoSaida, strSqlCampoSaidaNameOnly, strSqlCampoGroupByOrderBy, strSqlCampoGroupByOrderByNameOnly, strColSpanTodasColunas
dim vl_desconto, vl_total_desconto
dim vl_total_final_lista, vl_lista, vl_total_lista
dim vl_venda, vl_RT, vl_total_RT
dim vl_RA_bruto, vl_RA_liquido, vl_total_RA_liquido
dim perc_desconto
dim vRelat(), intIdxVetor
dim strItemAtual, strItemAnterior
dim intQtdeSubTotalProdutos, intQtdeProdutos
dim vl_sub_total_venda, vl_sub_total_lista, vl_sub_total_desconto
dim vl_sub_total_RT, vl_sub_total_RA_liquido
dim strCampoOrdenacao
dim strAuxUF, strAuxUFAnterior, strAuxIndicador, strAuxIndicadorAnterior, strAuxCodigoPai, strAuxCodigoPaiAnterior, strAuxCodigo, strAuxCodigoAnterior
dim strAuxCidade, strAuxCidadeAnterior
dim strAuxFabricante,strAuxFabricanteAnterior,strAuxNomeFabricante,strAuxNomeFabricanteAnterior, strAuxProdutoGrupo,strAuxProdutoGrupoAnterior,strAuxNomeProdutoGrupo,strAuxNomeProdutoGrupoAnterior
dim strAuxEmpresa,strAuxEmpresaAnterior,strAuxNomeEmpresa,strAuxNomeEmpresaAnterior
dim intIdxUF,intIdxEmpresa
dim vl_venda_UF, percFatVenda_UF, vl_lista_UF, vl_desconto_UF, perc_desconto_UF, vl_RT_UF, percRTFatVenda_UF, percRTEDesc_UF, vl_RA_liquido_UF
dim v_total_UF, blnAchou, intIdxSelecionado, vl_aux, vl_aux_ProdutoGrupo, vl_marg_contrib_ProdutoGrupo
dim v_total_FABRICANTE,vl_total_FABRICANTE, v_total_origem_pedido,v_total_empresa, v_total_GRUPO_PRODUTO
dim intIdxFabricante
dim vl_desconto_Fabricante,vl_RA_liquido_Fabricante,percFatVenda_Fabricante,perc_desconto_Fabricante,vl_lista_Fabricante,vl_venda_Fabricante
dim percRTFatVenda_Fabricante,vl_RT_Fabricante,percRTEDesc_Fabricante,marg_contrib_Fabricante
dim vl_desconto_Empresa,vl_RA_liquido_Empresa,percFatVenda_Empresa,perc_desconto_Empresa,vl_lista_Empresa,vl_venda_Empresa
dim percRTFatVenda_Empresa,vl_RT_Empresa,percRTEDesc_Empresa,marg_contrib_Empresa,lucro_liquido_Empresa,vl_ent_Empresa
dim vl_ent_Fabricante,lucro_liquido_Fabricante
dim strAuxVlFabricante,strAuxVlFabricanteAnterior,intQtdeTotalFabricante,intQtdeTotalIndicadorUf,intQtdeTotalEmpresa, strAuxVlProdutoGrupo
dim intQtdeTotalProdutoGrupo, vl_total_ProdutoGrupo, vl_venda_ProdutoGrupo, percFatVenda_ProdutoGrupo, vl_lista_ProdutoGrupo, vl_desconto_ProdutoGrupo, perc_desconto_ProdutoGrupo, vl_RT_ProdutoGrupo
dim percRTFatVenda_ProdutoGrupo, percRTEDesc_ProdutoGrupo, vl_RA_liquido_ProdutoGrupo, lucro_liquido_ProdutoGrupo, marg_contrib_ProdutoGrupo, vl_ent_ProdutoGrupo
dim strAuxVlEmpresa,strAuxVlEmpresaAnterior
dim lucro_liquido, marg_contrib, vl_ent, lucro_liquido_total, lucro_liquido_subtotal, marg_contrib_ordenacao
dim vl_NF, lucro_bruto, lucro_bruto_UF
dim vl_total_NF, vl_sub_total_NF, lucro_bruto_total, lucro_bruto_subtotal, marg_contrib_bruta, marg_contrib_bruta_ordenacao, marg_contrib_bruta_UF, vl_NF_UF, marg_contrib_bruta_Fabricante, vl_NF_Fabricante, lucro_bruto_Fabricante
dim marg_contrib_bruta_ProdutoGrupo, vl_NF_ProdutoGrupo, lucro_bruto_ProdutoGrupo, vl_NF_Empresa, marg_contrib_bruta_Empresa, lucro_bruto_Empresa, marg_contrib_bruta_final, vl_total_final_NF, vl_marg_contrib_bruta_ProdutoGrupo
dim lucro_liquido_UF, marg_contrib_UF, marg_contrib_final, subtotal_entrada, total_entrada, vl_ent_UF
dim iEspacos, md
dim strEspacos,v_grupos,v_subgrupos,cont,s_where_temp,intQtdeTotalIndicadorCidUf,intQtdeTotal
dim vTotGrupo, vl_aux_RA_liquido


	intQtdeTotal = 0

'	CRITÉRIOS COMUNS
	s_where = ""

	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.vendedor = '" & c_vendedor & "')"
		end if
		
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if

	if c_captador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.captador = '" & c_captador & "')"
		end if

    s_where_temp = ""
    v_pedido_origem = split(c_pedido_origem, ", ")
    if c_pedido_origem <> "" then
        for i = LBound(v_pedido_origem) to UBound(v_pedido_origem)
            if s_where_temp <> "" then s_where_temp = s_where_temp & ", "
            s_where_temp = s_where_temp & "'" & v_pedido_origem(i) & "'"
        next
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_PEDIDO.marketplace_codigo_origem IN (" & s_where_temp & "))"
    end if
    
    s_where_temp = ""
    if c_grupo_pedido_origem <> "" then
        v_grupo_pedido_origem = split(c_grupo_pedido_origem, ", ")
        for i = LBound(v_grupo_pedido_origem) to UBound(v_grupo_pedido_origem)
            s = "SELECT codigo FROM t_CODIGO_DESCRICAO WHERE (codigo_pai = '" & v_grupo_pedido_origem(i) & "') AND grupo='PedidoECommerce_Origem'"
            if rs.State <> 0 then rs.Close
	        rs.open s, cn
		    if rs.Eof then
                alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_grupo_pedido_origem & " NÃO EXISTE."
                exit for
            else
                do while Not rs.Eof
                    if s_where_temp <> "" then s_where_temp = s_where_temp & ", "
                    s_where_temp = s_where_temp & "'" & rs("codigo") & "'"
                    rs.MoveNext
                loop
            end if
        next
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_PEDIDO.marketplace_codigo_origem IN (" & s_where_temp & "))"
    end if

	s_where_temp = ""
	if c_grupo <> "" then
		v_grupos = split(c_grupo, ", ")
		for cont = LBound(v_grupos) to UBound(v_grupos)
			if Trim(v_grupos(cont)) <> "" then
				if s_where_temp <> "" then s_where_temp = s_where_temp & ", "
				s_where_temp = s_where_temp & "'" & Trim(v_grupos(cont)) & "'"
				end if
			next
		
		if s_where_temp <> "" then
			s_where_temp = " (t_PRODUTO.grupo IN (" & s_where_temp & "))"
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s_where_temp & ")"
			end if
		end if

	s_where_temp = ""
	if c_subgrupo <> "" then
		v_subgrupos = split(c_subgrupo, ", ")
		for cont = LBound(v_subgrupos) to UBound(v_subgrupos)
			if Trim(v_subgrupos(cont)) <> "" then
				if s_where_temp <> "" then s_where_temp = s_where_temp & ", "
				s_where_temp = s_where_temp & "'" & Trim(v_subgrupos(cont)) & "'"
				end if
			next
		
		if s_where_temp <> "" then
			s_where_temp = " (t_PRODUTO.subgrupo IN (" & s_where_temp & "))"
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s_where_temp & ")"
			end if
		end if

	if c_cnpj_cpf <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_CLIENTE.cnpj_cpf = '" & c_cnpj_cpf & "')"
		end if
	
	if rb_tipo_cliente <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_CLIENTE.tipo = '" & rb_tipo_cliente & "')"
		end if
	
	if rb_tipo_cliente <> ID_PF then
		if (ckb_contribuinte_icms_nao <> "") Or (ckb_contribuinte_icms_sim <> "") Or (ckb_contribuinte_icms_isento <> "") then
			s_where_aux = ""
			if ckb_contribuinte_icms_nao <> "" then
				if s_where_aux <> "" then s_where_aux = s_where_aux & ", "
				s_where_aux = s_where_aux & ckb_contribuinte_icms_nao
				end if

			if ckb_contribuinte_icms_sim <> "" then
				if s_where_aux <> "" then s_where_aux = s_where_aux & ", "
				s_where_aux = s_where_aux & ckb_contribuinte_icms_sim
				end if

			if ckb_contribuinte_icms_isento <> "" then
				if s_where_aux <> "" then s_where_aux = s_where_aux & ", "
				s_where_aux = s_where_aux & ckb_contribuinte_icms_isento
				end if

			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " ((t_CLIENTE.tipo = '" & ID_PJ & "') AND (t_CLIENTE.contribuinte_icms_status IN (" & s_where_aux & ")))"
			end if
		end if

    if c_uf_saida <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE.uf = '" + c_uf_saida + "')"
			end if

     if c_empresa <> "" then
        if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.id_nfe_emitente = " & c_empresa & ")"			
	 end if

	if (rb_saida = COD_SAIDA_REL_CIDADE_UF) then
		if c_loc_uf <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE.uf = '" + c_loc_uf + "')"
			end if
		
		if c_loc_digitada <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE.cidade LIKE '" + c_loc_digitada + "%')"
		else
			s_where_loc = ""
			if c_loc_escolhidas <> "" then
				v_loc_escolhidas =split(c_loc_escolhidas, ", ")
				for i = LBound(v_loc_escolhidas) to UBound(v_loc_escolhidas)
					if s_where_loc <> "" then s_where_loc = s_where_loc & " OR"
					s_where_loc = s_where_loc & " (t_CLIENTE.cidade = '" & trim(replace(v_loc_escolhidas(i), "'", "''")) & "' COLLATE Latin1_General_CI_AI)"
					next
				if s_where_loc <> "" then
						s_where_loc = " AND (" & s_where_loc & ") "
						s_where = s_where & s_where_loc
					end if
				end if
			end if
		end if
	
	if (True) Or (rb_saida = COD_SAIDA_REL_UF) Or (rb_saida = COD_SAIDA_REL_VENDEDOR) then
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
	
	s_where_loja = ""
	for i=Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
				s_where_loja = s_where_loja & " (t_PEDIDO.numero_loja = " & v_loja(i) & ")"
			else
				s = ""
				if v(Lbound(v))<>"" then 
					if s <> "" then s = s & " AND"
					s = s & " (t_PEDIDO.numero_loja >= " & v(Lbound(v)) & ")"
					end if
				if v(Ubound(v))<>"" then
					if s <> "" then s = s & " AND"
					s = s & " (t_PEDIDO.numero_loja <= " & v(Ubound(v)) & ")"
					end if
				if s <> "" then 
					if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
					s_where_loja = s_where_loja & " (" & s & ")"
					end if
				end if
			end if
		next
		
	if s_where_loja <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s_where_loja & ")"
		end if

'	CST
	if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
		s_where_cst = ""
		if c_cst <> "" then
			for i=Lbound(v_cst) to Ubound(v_cst)
				if v_cst(i) <> "" then
					if s_where_cst <> "" then s_where_cst = s_where_cst & ", "
					s_where_cst = s_where_cst & "'" & v_cst(i) & "'"
					end if
				next
		
			if s_where_cst <> "" then
				s_where_cst = " t_ESTOQUE_ITEM.cst IN (" & s_where_cst & ")"
				if s_where <> "" then s_where = s_where & " AND"
				s_where = s_where & " (" & s_where_cst & ")"
				end if
			end if
		end if

'	CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	s_where_venda = ""

	if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
		if IsDate(c_dt_cadastro_inicio) then
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (t_PEDIDO.data >= " & bd_formata_data(StrToDate(c_dt_cadastro_inicio)) & ")"
			end if
			
		if IsDate(c_dt_cadastro_termino) then
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (t_PEDIDO.data < " & bd_formata_data(StrToDate(c_dt_cadastro_termino)+1) & ")"
			end if
		end if

	if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
		if IsDate(c_dt_entregue_inicio) then
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data >= " & bd_formata_data(StrToDate(c_dt_entregue_inicio)) & ")"
			end if
			
		if IsDate(c_dt_entregue_termino) then
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
			end if
		end if

	if c_fabricante <> "" then
		if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (t_ESTOQUE_MOVIMENTO.fabricante = '" & c_fabricante & "')"
		else
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (t_PEDIDO_ITEM.fabricante = '" & c_fabricante & "')"
			end if
		end if
	
	if c_produto <> "" then
		if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (t_ESTOQUE_MOVIMENTO.produto = '" & c_produto & "')"
		else
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (t_PEDIDO_ITEM.produto = '" & c_produto & "')"
			end if
		end if
	
'	CRITÉRIOS PARA DEVOLUÇÕES
	s_where_devolucao = ""
	
	if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
		if IsDate(c_dt_entregue_inicio) then
			if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
			s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(StrToDate(c_dt_entregue_inicio)) & ")"
			end if
			
		if IsDate(c_dt_entregue_termino) then
			if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
			s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
			end if
		end if
	
	if c_fabricante <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.produto = '" & c_produto & "')"
		end if

'	MONTA SQL DE CONSULTA
	if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
	'	QUAL É O CAMPO DE SAÍDA SELECIONADO?
		strSqlCampoGroupByOrderBy = ""
		strSqlCampoGroupByOrderByNameOnly = ""
		if rb_saida = COD_SAIDA_REL_PRODUTO then
			strSqlCampoSaida =	"t_PEDIDO_ITEM.fabricante, " & _
								"t_PEDIDO_ITEM.produto, " & _
								"t_PRODUTO.descricao, " & _
								"t_PRODUTO.descricao_html"
			strSqlCampoSaidaNameOnly = "fabricante, " & _
										"produto, " & _
										"descricao, " & _
										"descricao_html"
		elseif rb_saida = COD_SAIDA_REL_FABRICANTE then
			strSqlCampoSaida =	"t_PEDIDO_ITEM.fabricante, " & _
                                "t_PEDIDO_ITEM.produto," & _
								"t_PRODUTO.descricao, " & _
								"t_PRODUTO.descricao_html"
			strSqlCampoSaidaNameOnly = "fabricante, " & _
										"produto," & _
										"descricao, " & _
										"descricao_html"
        elseif rb_saida = COD_SAIDA_REL_GRUPO_PRODUTO then
            strSqlCampoSaida = "t_PEDIDO_ITEM.fabricante, " & _
                                "t_PRODUTO.grupo, " & _
                                "t_PRODUTO_GRUPO.descricao AS grupo_descricao, " & _
                                "t_PEDIDO_ITEM.produto," & _
								"t_PRODUTO.descricao, " & _
								"t_PRODUTO.descricao_html"
            strSqlCampoSaidaNameOnly = "fabricante, " & _
										"grupo, " & _
										"grupo_descricao, " & _
										"produto," & _
										"descricao, " & _
										"descricao_html"
		elseif rb_saida = COD_SAIDA_REL_VENDEDOR then
			strSqlCampoSaida = "t_PEDIDO__BASE.vendedor"
			strSqlCampoSaidaNameOnly = "vendedor"
		elseif rb_saida = COD_SAIDA_REL_INDICADOR then
			strSqlCampoSaida =	"t_PEDIDO__BASE.indicador, " & _
								"t_ORCAMENTISTA_E_INDICADOR.uf"
			strSqlCampoSaidaNameOnly = "indicador, " & _
										"uf"
		elseif rb_saida = COD_SAIDA_REL_UF then
			strSqlCampoSaida = "t_CLIENTE.uf"
			strSqlCampoSaidaNameOnly = "uf"
		elseif rb_saida = COD_SAIDA_REL_INDICADOR_UF then
			strSqlCampoSaida =	"t_ORCAMENTISTA_E_INDICADOR.uf, " & _
								"t_PEDIDO__BASE.indicador"
			strSqlCampoSaidaNameOnly = "uf, " & _
										"indicador"
		elseif rb_saida = COD_SAIDA_REL_CIDADE_UF then
			strSqlCampoSaida =	"t_CLIENTE.uf, " & _
								"t_CLIENTE.cidade COLLATE Latin1_General_CI_AI AS cidade"
			strSqlCampoSaidaNameOnly = "uf, " & _
										"cidade COLLATE Latin1_General_CI_AI  AS cidade"
        elseif rb_saida = COD_SAIDA_REL_ORIGEM_PEDIDO then
		'	QUANDO NÃO HOUVER CÓDIGO DA ORIGEM DA PEDIDO, AGRUPA PELA LOJA
            strSqlCampoSaida = "t_CODIGO_DESCRICAO.codigo_pai, " & _
                               "CASE WHEN t_CODIGO_DESCRICAO.codigo IS NOT NULL THEN t_CODIGO_DESCRICAO.codigo ELSE t_PEDIDO__BASE.loja END AS codigo_origem, " & _
                               "CASE WHEN t_CODIGO_DESCRICAO.codigo IS NOT NULL THEN t_CODIGO_DESCRICAO.descricao ELSE 'Loja ' + t_PEDIDO__BASE.loja END AS descricao_origem"
            strSqlCampoSaidaNameOnly = "codigo_pai, " & _
									   "codigo_origem, " & _
									   "descricao_origem"
			strSqlCampoGroupByOrderBy = "t_CODIGO_DESCRICAO.codigo_pai, " & _
                               "CASE WHEN t_CODIGO_DESCRICAO.codigo IS NOT NULL THEN t_CODIGO_DESCRICAO.codigo ELSE t_PEDIDO__BASE.loja END, " & _
                               "CASE WHEN t_CODIGO_DESCRICAO.codigo IS NOT NULL THEN t_CODIGO_DESCRICAO.descricao ELSE 'Loja ' + t_PEDIDO__BASE.loja END"
			strSqlCampoGroupByOrderByNameOnly = "codigo_pai, " & _
												"codigo_origem, " & _
												"descricao_origem"
        elseif rb_saida = COD_SAIDA_REL_LOJA then
            strSqlCampoSaida = "t_PEDIDO__BASE.loja"
            strSqlCampoSaidaNameOnly = "loja"
        elseif rb_saida = COD_SAIDA_REL_EMPRESA then
            strSqlCampoSaida =  "t_PEDIDO.id_nfe_emitente, " & _
                                "t_PEDIDO_ITEM.fabricante, " & _
                                "t_PEDIDO_ITEM.produto, " & _
								"t_PRODUTO.descricao, " & _
								"t_PRODUTO.descricao_html"
            strSqlCampoSaidaNameOnly =  "id_nfe_emitente, " & _
										"fabricante, " & _
										"produto, " & _
										"descricao, " & _
										"descricao_html"
		else
			strSqlCampoSaida = ""
			strSqlCampoSaidaNameOnly = ""
			end if
		
		if strSqlCampoGroupByOrderBy = "" then strSqlCampoGroupByOrderBy = replace(replace(strSqlCampoSaida, "AS grupo_descricao", ""), "AS cidade", "")
		if strSqlCampoGroupByOrderByNameOnly = "" then strSqlCampoGroupByOrderByNameOnly = replace(strSqlCampoSaidaNameOnly, "AS cidade", "")

		s = s_where
		if (s <> "") And (s_where_venda <> "") then s = s & " AND"
		s = s & s_where_venda
		if s <> "" then s = " AND" & s
		s_sql = "SELECT * FROM (" & _
				"SELECT " & _
					strSqlCampoSaida & "," & _
					" t_PEDIDO__BASE.perc_RT," & _
					" t_PEDIDO__BASE.st_tem_desagio_RA," & _
					" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
					" t_PEDIDO_ITEM.pedido," & _
					" t_PEDIDO_ITEM.fabricante As CodFabricante," & _
					" t_PEDIDO_ITEM.produto As CodProduto," & _
					" Sum(t_ESTOQUE_MOVIMENTO.qtde) AS qtde," & _
					" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda) AS valor_venda," & _
					" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_NF) AS valor_NF," & _
					" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_lista) AS valor_lista," & _
					" Sum(t_ESTOQUE_MOVIMENTO.qtde * t_ESTOQUE_ITEM.vl_custo2) AS valor_entrada" & _
				" FROM t_PEDIDO" & _
					" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
						" ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
					" INNER JOIN t_PEDIDO_ITEM" & _
						" ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
					" INNER JOIN t_ESTOQUE_MOVIMENTO" & _
						" ON ((t_ESTOQUE_MOVIMENTO.pedido = t_PEDIDO_ITEM.pedido)" & _
							" AND (t_ESTOQUE_MOVIMENTO.fabricante = t_PEDIDO_ITEM.fabricante)" & _
							" AND (t_ESTOQUE_MOVIMENTO.produto = t_PEDIDO_ITEM.produto))" & _
					" INNER JOIN t_ESTOQUE_ITEM" & _
						" ON ((t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE_ITEM.id_estoque)" & _
							" AND (t_ESTOQUE_MOVIMENTO.fabricante = t_ESTOQUE_ITEM.fabricante)" & _
							" AND (t_ESTOQUE_MOVIMENTO.produto = t_ESTOQUE_ITEM.produto))" & _
					" LEFT JOIN t_PRODUTO" & _
						" ON ((t_PEDIDO_ITEM.fabricante=t_PRODUTO.fabricante)AND(t_PEDIDO_ITEM.produto=t_PRODUTO.produto))" & _
					" LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
					" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR" & _
						" ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
					" LEFT JOIN t_CODIGO_DESCRICAO ON ((t_CODIGO_DESCRICAO.codigo=t_PEDIDO.marketplace_codigo_origem) AND (t_CODIGO_DESCRICAO.grupo='PedidoECommerce_Origem'))" & _
					" INNER JOIN t_CLIENTE" & _
						" ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
				" WHERE" & _
					" ((t_ESTOQUE_MOVIMENTO.estoque <> '" & ID_ESTOQUE_SEM_PRESENCA & "') AND (t_ESTOQUE_MOVIMENTO.anulado_status = 0))" & _
					" AND (t_PEDIDO.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
					s & _
				" GROUP BY " & _
					strSqlCampoGroupByOrderBy & "," & _
					" t_PEDIDO__BASE.perc_RT," & _
					" t_PEDIDO__BASE.st_tem_desagio_RA," & _
					" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
					" t_PEDIDO_ITEM.pedido," & _
					" t_PEDIDO_ITEM.fabricante," & _
					" t_PEDIDO_ITEM.produto"
			
			s_sql = s_sql & _
				" UNION ALL " & _
				" SELECT " & _
					strSqlCampoSaidaNameOnly & "," & _
					" perc_RT," & _
					" st_tem_desagio_RA," & _
					" perc_desagio_RA_liquida," & _
					" pedido," & _
					" CodFabricante," & _
					" CodProduto," & _
					" Sum(qtde) AS qtde," & _
					" Sum(qtde * preco_venda) AS valor_venda," & _
					" Sum(qtde * preco_NF) AS valor_NF," & _
					" Sum(qtde * preco_lista) AS valor_lista," & _
					" Sum(qtde * (Coalesce(Coalesce(valor_entrada, valor_entrada_ult_estoque), preco_venda))) AS valor_entrada" & _
				" FROM (" & _
					"SELECT " & _
						strSqlCampoSaida & "," & _
						" t_PEDIDO__BASE.perc_RT," & _
						" t_PEDIDO__BASE.st_tem_desagio_RA," & _
						" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
						" t_PEDIDO_ITEM.pedido," & _
						" t_PEDIDO_ITEM.fabricante As CodFabricante," & _
						" t_PEDIDO_ITEM.produto As CodProduto," & _
						" t_ESTOQUE_MOVIMENTO.qtde," & _
						" t_PEDIDO_ITEM.preco_venda," & _
						" t_PEDIDO_ITEM.preco_NF," & _
						" t_PEDIDO_ITEM.preco_lista," & _
						" (SELECT TOP 1" & _
								" (vl_custo2_total / qtde)" & _
							" FROM t_ESTOQUE_VENDA_SALDO_DIARIO" & _
							" WHERE" & _
								" (id_nfe_emitente = - 1)" & _
								" AND (fabricante = t_ESTOQUE_MOVIMENTO.fabricante) " & _
								" AND (produto = t_ESTOQUE_MOVIMENTO.produto)" & _
								" AND (data <= t_PEDIDO__BASE.data)" & _
							" ORDER BY" & _
								" data DESC" & _
							") AS valor_entrada," & _
						" tEI_ULT.vl_custo2 AS valor_entrada_ult_estoque" & _
					" FROM t_PEDIDO" & _
						" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base = t_PEDIDO__BASE.pedido)" & _
						" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido = t_PEDIDO_ITEM.pedido)" & _
						" INNER JOIN t_ESTOQUE_MOVIMENTO" & _
							" ON ((t_ESTOQUE_MOVIMENTO.pedido = t_PEDIDO_ITEM.pedido)" & _
							" AND (t_ESTOQUE_MOVIMENTO.fabricante = t_PEDIDO_ITEM.fabricante)" & _
							" AND (t_ESTOQUE_MOVIMENTO.produto = t_PEDIDO_ITEM.produto))" & _
						" LEFT JOIN t_PRODUTO" & _
							" ON ((t_PEDIDO_ITEM.fabricante = t_PRODUTO.fabricante)" & _
							" AND (t_PEDIDO_ITEM.produto = t_PRODUTO.produto))" & _
						" LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo = t_PRODUTO_GRUPO.codigo)" & _
						" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador = t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
						" LEFT JOIN t_CODIGO_DESCRICAO" & _
							" ON ((t_CODIGO_DESCRICAO.codigo = t_PEDIDO.marketplace_codigo_origem)" & _
							" AND (t_CODIGO_DESCRICAO.grupo = 'PedidoECommerce_Origem'))" & _
						" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente = t_CLIENTE.id)" & _
						" LEFT JOIN (" & _
							"SELECT" & _
								" t_ESTOQUE_ITEM.*" & _
							" FROM (" & _
								"SELECT" & _
									" t_ESTOQUE_ITEM.fabricante," & _
									" produto," & _
									" Max(t_ESTOQUE.id_estoque) AS id_estoque" & _
								" FROM t_ESTOQUE_ITEM INNER JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" & _
								" WHERE" & _
									" (entrada_especial = 0)" & _
								"GROUP BY" & _
									" t_ESTOQUE_ITEM.fabricante," & _
									" produto" & _
								") t" & _
							" INNER JOIN t_ESTOQUE_ITEM ON (t.id_estoque = t_ESTOQUE_ITEM.id_estoque)" & _
								" AND (t.fabricante = t_ESTOQUE_ITEM.fabricante)" & _
								" AND (t.produto = t_ESTOQUE_ITEM.produto)" & _
							") tEI_ULT ON (tEI_ULT.fabricante=t_ESTOQUE_MOVIMENTO.fabricante) AND (tEI_ULT.produto=t_ESTOQUE_MOVIMENTO.produto)" & _
					" WHERE" & _
						" ((t_ESTOQUE_MOVIMENTO.estoque = '" & ID_ESTOQUE_SEM_PRESENCA & "') AND (t_ESTOQUE_MOVIMENTO.anulado_status = 0))" & _
						" AND (t_PEDIDO.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
						s & _
					") tt" & _
				" GROUP BY " & _
					strSqlCampoGroupByOrderByNameOnly & "," & _
					" perc_RT," & _
					" st_tem_desagio_RA," & _
					" perc_desagio_RA_liquida," & _
					" pedido," & _
					" CodFabricante," & _
					" CodProduto"

			s_sql = s_sql & _
				") ttt" & _
				" ORDER BY " & _
					strSqlCampoGroupByOrderByNameOnly
		
	elseif rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
	'	QUAL É O CAMPO DE SAÍDA SELECIONADO?
		strSqlCampoGroupByOrderBy = ""
		if rb_saida = COD_SAIDA_REL_PRODUTO then
			strSqlCampoSaida =	"t_PRODUTO.fabricante, " & _
								"t_PRODUTO.produto, " & _
								"t_PRODUTO.descricao, " & _
								"t_PRODUTO.descricao_html"
		elseif rb_saida = COD_SAIDA_REL_FABRICANTE then
			strSqlCampoSaida =	"t_FABRICANTE.fabricante, " & _
			                    "t_FABRICANTE.nome, " & _
								"t_PRODUTO.produto," & _
								"t_PRODUTO.descricao, " & _
								"t_PRODUTO.descricao_html"
        elseif rb_saida = COD_SAIDA_REL_GRUPO_PRODUTO then
			strSqlCampoSaida =	"t_FABRICANTE.fabricante, " & _
			                    "t_FABRICANTE.nome, " & _
                                "t_PRODUTO.grupo, " & _
                                "t_PRODUTO_GRUPO.descricao AS grupo_descricao, " & _
								"t_PRODUTO.produto," & _
								"t_PRODUTO.descricao, " & _
								"t_PRODUTO.descricao_html"
		elseif rb_saida = COD_SAIDA_REL_VENDEDOR then
			strSqlCampoSaida = "t_PEDIDO__BASE.vendedor"
		elseif rb_saida = COD_SAIDA_REL_INDICADOR then
			strSqlCampoSaida =	"t_PEDIDO__BASE.indicador, " & _
								"t_ORCAMENTISTA_E_INDICADOR.uf"
		elseif rb_saida = COD_SAIDA_REL_UF then
			strSqlCampoSaida = "t_CLIENTE.uf"
		elseif rb_saida = COD_SAIDA_REL_INDICADOR_UF then
			strSqlCampoSaida =	"t_ORCAMENTISTA_E_INDICADOR.uf, " & _
								"t_PEDIDO__BASE.indicador"
		elseif rb_saida = COD_SAIDA_REL_CIDADE_UF then
			strSqlCampoSaida =	"t_CLIENTE.uf, " & _
								"t_CLIENTE.cidade COLLATE Latin1_General_CI_AI"
        elseif rb_saida = COD_SAIDA_REL_ORIGEM_PEDIDO then
		'	QUANDO NÃO HOUVER CÓDIGO DA ORIGEM DA PEDIDO, AGRUPA PELA LOJA
            strSqlCampoSaida = "t_CODIGO_DESCRICAO.codigo_pai, " & _
                               "CASE WHEN t_CODIGO_DESCRICAO.codigo IS NOT NULL THEN t_CODIGO_DESCRICAO.codigo ELSE t_PEDIDO__BASE.loja END AS codigo_origem, " & _
                               "CASE WHEN t_CODIGO_DESCRICAO.codigo IS NOT NULL THEN t_CODIGO_DESCRICAO.descricao ELSE 'Loja ' + t_PEDIDO__BASE.loja END AS descricao_origem"
			strSqlCampoGroupByOrderBy = "t_CODIGO_DESCRICAO.codigo_pai, " & _
                               "CASE WHEN t_CODIGO_DESCRICAO.codigo IS NOT NULL THEN t_CODIGO_DESCRICAO.codigo ELSE t_PEDIDO__BASE.loja END, " & _
                               "CASE WHEN t_CODIGO_DESCRICAO.codigo IS NOT NULL THEN t_CODIGO_DESCRICAO.descricao ELSE 'Loja ' + t_PEDIDO__BASE.loja END"
        elseif rb_saida = COD_SAIDA_REL_LOJA then
            strSqlCampoSaida = "t_PEDIDO__BASE.loja"
        elseif rb_saida = COD_SAIDA_REL_EMPRESA then
            strSqlCampoSaida = "t_PEDIDO.id_nfe_emitente," & _
                               "t_PRODUTO.fabricante, " & _
                               "t_PRODUTO.produto," & _
							   "t_PRODUTO.descricao, " & _
							   "t_PRODUTO.descricao_html"
		else
			strSqlCampoSaida = ""
			end if
		
		if strSqlCampoGroupByOrderBy = "" then strSqlCampoGroupByOrderBy = replace(strSqlCampoSaida, "AS grupo_descricao", "")

	'	A) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
	'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
	'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	'	B) O CAMPO 'QTDE' A SER USADO DEVE SER DA TABELA T_ESTOQUE_MOVIMENTO, JÁ
	'	QUE UM PEDIDO PODE TER USADO DIVERSOS LOTES DO ESTOQUE PARA ATENDER A
	'	UM ÚNICO PRODUTO.  NESSE CASO, HAVERÁ MAIS DE UM REGISTRO EM 
	'	T_ESTOQUE_MOVIMENTO SE RELACIONANDO COM O MESMO REGISTRO DE T_PEDIDO_ITEM.
	'	A SOMA DE 'QTDE' DOS REGISTROS DE T_ESTOQUE_MOVIMENTO RESULTAM NO VALOR
	'	DE 'QTDE' DE T_PEDIDO_ITEM.
		s = s_where
		if (s <> "") And (s_where_venda <> "") then s = s & " AND"
		s = s & s_where_venda
		if s <> "" then s = " AND" & s
		s_sql = "SELECT " & _ 
					strSqlCampoSaida & "," & _
					" t_PEDIDO__BASE.perc_RT," & _
					" t_PEDIDO__BASE.st_tem_desagio_RA," & _
					" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
					" t_ESTOQUE_MOVIMENTO.pedido," & _
					" t_ESTOQUE_MOVIMENTO.fabricante As CodFabricante," & _
					" t_ESTOQUE_MOVIMENTO.produto As CodProduto," & _
					" Sum(t_ESTOQUE_MOVIMENTO.qtde) AS qtde," & _
					" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda) AS valor_venda," & _
					" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_NF) AS valor_NF," & _
					" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_lista) AS valor_lista," & _
                    " Sum(t_ESTOQUE_MOVIMENTO.qtde*t_ESTOQUE_ITEM.vl_custo2) AS valor_entrada" & _
				" FROM t_ESTOQUE_MOVIMENTO" & _
					" INNER JOIN t_ESTOQUE_ITEM" & _
						" ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
					" INNER JOIN t_PEDIDO" & _
						" ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
					" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
						" ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
					" INNER JOIN t_PEDIDO_ITEM" & _
						" ON ((t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto))" & _
					"LEFT JOIN t_FABRICANTE " & _
					    " ON (t_ESTOQUE_MOVIMENTO.fabricante=t_FABRICANTE.fabricante)" & _
					" LEFT JOIN t_PRODUTO" & _
						" ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
                    " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
					" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR" & _
						" ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
                    " LEFT JOIN t_CODIGO_DESCRICAO ON ((t_CODIGO_DESCRICAO.codigo=t_PEDIDO.marketplace_codigo_origem) AND (t_CODIGO_DESCRICAO.grupo='PedidoECommerce_Origem'))" & _ 
					" INNER JOIN t_CLIENTE" & _
						" ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
				" WHERE" & _
					" (anulado_status=0)" & _
					" AND (estoque='" & ID_ESTOQUE_ENTREGUE & "')" & _
					s & _
				" GROUP BY " & _
					strSqlCampoGroupByOrderBy & "," & _
					" t_PEDIDO__BASE.perc_RT," & _
					" t_PEDIDO__BASE.st_tem_desagio_RA," & _
					" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
					" t_ESTOQUE_MOVIMENTO.pedido," & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
					" t_ESTOQUE_MOVIMENTO.produto"
	
	'	LEMBRE-SE: NA DEVOLUÇÃO DO PRODUTO, É CRIADA UMA ENTRADA NO ESTOQUE DE VENDA P/
	'	REPRESENTAR A ENTRADA DA MERCADORIA NO ESTOQUE. ENTRETANTO, A QUANTIDADE
	'	DEVOLVIDA FICA INICIALMENTE TODA ALOCADA P/ O ESTOQUE DE DEVOLUÇÃO, DEVIDO
	'	À NECESSIDADE DE TRATAR A MERCADORIA ANTES DE DISPONIBILIZA-LA P/ VENDA.
	'	IMPORTANTE: NO CASO DE OCORRER A DEVOLUÇÃO DE VÁRIAS UNIDADES, PODEM SER
	'	CRIADOS VÁRIOS REGISTROS DE ESTOQUE A DIFERENTES CUSTOS DE AQUISIÇÃO.
		s = s_where
		if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
		s = s & s_where_devolucao
		if s <> "" then s = " WHERE" & s
		s_sql = s_sql & _
				" UNION ALL" & _
				" SELECT " & _
					strSqlCampoSaida & "," & _
					" t_PEDIDO__BASE.perc_RT," & _
					" t_PEDIDO__BASE.st_tem_desagio_RA," & _
					" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.pedido," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.fabricante As CodFabricante," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.produto As CodProduto," & _
					" Sum(-t_ESTOQUE_ITEM.qtde) AS qtde," & _
					" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS valor_venda," & _
					" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_NF) AS valor_NF," & _
					" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_lista) AS valor_lista," & _
					" Sum(-t_ESTOQUE_ITEM.qtde*t_ESTOQUE_ITEM.vl_custo2) AS valor_entrada" & _
				" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
					" INNER JOIN t_PEDIDO" & _
						" ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)" & _
					" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
						" ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
					" INNER JOIN t_ESTOQUE" & _
						" ON (t_PEDIDO_ITEM_DEVOLVIDO.id=t_ESTOQUE.devolucao_id_item_devolvido)" & _
					" INNER JOIN t_ESTOQUE_ITEM" & _
						" ON ((t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_ESTOQUE_ITEM.produto))" & _ 
					" INNER JOIN t_PEDIDO_ITEM " & _
                        " ON ((t_PEDIDO_ITEM_DEVOLVIDO.pedido = t_PEDIDO_ITEM.pedido) AND (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = t_PEDIDO_ITEM.fabricante) AND (t_PEDIDO_ITEM_DEVOLVIDO.produto = t_PEDIDO_ITEM.produto))" & _
                    " LEFT JOIN t_FABRICANTE " & _
					    " ON (t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_FABRICANTE.fabricante)" & _
					" LEFT JOIN t_PRODUTO" & _
						" ON ((t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_PRODUTO.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_PRODUTO.produto))" & _
                    " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
					" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR" & _
						" ON (t_PEDIDO.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
                    " LEFT JOIN t_CODIGO_DESCRICAO ON ((t_CODIGO_DESCRICAO.codigo=t_PEDIDO.marketplace_codigo_origem) AND (t_CODIGO_DESCRICAO.grupo='PedidoECommerce_Origem'))" & _                    
					" INNER JOIN t_CLIENTE" & _
						" ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
					s & _
				" GROUP BY " & _
					strSqlCampoGroupByOrderBy & "," & _
					" t_PEDIDO__BASE.perc_RT," & _
					" t_PEDIDO__BASE.st_tem_desagio_RA," & _
					" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.pedido," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.fabricante," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.produto"
		
		s_sql = s_sql & _
				" ORDER BY " & _
					strSqlCampoGroupByOrderBy
		end if
			

  ' CABEÇALHO
	intLargCodProduto = 45
	intLargQtdeProduto = 35
	intLargDescrProduto = 220
	intLargCodFabricante = 45
	intLargNomeFabricante = 220
	intLargVendedor = 270
	intLargIndicador = 270
	intLargUF = 30
	intLargColMonetario = 70
	intLargColPerc = 42
	intLargCidade = 270
    intLargLoja = 90
	cab_table = "<table cellspacing=0>" & chr(13)
	cab = "	<tr style='background:azure' nowrap>" & chr(13)
	
	cab = cab & _
			"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13)
	
	if rb_saida = COD_SAIDA_REL_PRODUTO then
		strColSpanTodasColunas = "colspan='17'" '14+3
		cab = cab & _
			"		<td class='MDTE' style='width:" & CStr(intLargCodProduto) & "px' align='left' valign='bottom' nowrap><span class='R'>Código</span></td>" & chr(13) & _
			"		<td class='MTD' style='width:" & CStr(intLargDescrProduto) & "px' align='left' valign='bottom' nowrap><span class='R'>Descrição</span></td>" & chr(13) 
			
	elseif rb_saida = COD_SAIDA_REL_FABRICANTE then
	    strColSpanTodasColunas = "colspan='17'" '14+3
	    cab = cab & _
	        "		<td class='MDTE' style='width:" & CStr(intLargCodFabricante) & "px' align='left' valign='bottom' nowrap><span class='R'> Código </span></td>" & chr(13) & _
	        "		<td class='MTD' style='width:" & CStr(intLargNomeFabricante) & "px' align='left' valign='bottom' nowrap><span class='R'>Descrição</span></td>" & chr(13) 
	       
    elseif rb_saida = COD_SAIDA_REL_GRUPO_PRODUTO then
	    strColSpanTodasColunas = "colspan='17'" '14+3
	    cab = cab & _
	        "		<td class='MDTE' style='width:" & CStr(intLargCodFabricante) & "px' align='left' valign='bottom' nowrap><span class='R'>Código</span></td>" & chr(13) & _
	        "		<td class='MTD' style='width:" & CStr(intLargNomeFabricante) & "px' align='left' valign='bottom' nowrap><span class='R'>Descrição</span></td>" & chr(13) 
	elseif rb_saida = COD_SAIDA_REL_VENDEDOR then
		strColSpanTodasColunas = "colspan='16'" '13+3
		cab = cab & _
			"		<td class='MDTE' style='width:" & CStr(intLargVendedor) & "px' align='left' valign='bottom' nowrap><span class='R'>Vendedor</span></td>" & chr(13)            
	elseif rb_saida = COD_SAIDA_REL_INDICADOR then
		strColSpanTodasColunas = "colspan='17'" '14+3
		cab = cab & _
			"		<td class='MDTE' style='width:" & CStr(intLargIndicador) & "px' align='left' valign='bottom' nowrap><span class='R'>Indicador</span></td>" & chr(13) & _
			"		<td class='MTD' style='width:" & CStr(intLargUF) & "px' align='center' valign='bottom' nowrap><span class='Rc'>UF</span></td>" & chr(13)
	elseif rb_saida = COD_SAIDA_REL_UF then
		strColSpanTodasColunas = "colspan='16'" '13+3
		cab = cab & _
			"		<td class='MDTE' style='width:" & CStr(intLargUF) & "px' align='center' valign='bottom' nowrap><span class='Rc'>UF</span></td>" & chr(13)
	elseif rb_saida = COD_SAIDA_REL_INDICADOR_UF then
		strColSpanTodasColunas = "colspan='17'" '14+3
		cab = cab & _
			"		<td class='MDTE' style='width:" & CStr(intLargIndicador) & "px' align='left' valign='bottom' nowrap><span class='R'>Indicador</span></td>" & chr(13)
	elseif rb_saida = COD_SAIDA_REL_CIDADE_UF then
		strColSpanTodasColunas = "colspan='17'" '14+3
		cab = cab & _
			"		<td class='MDTE' style='width:" & CStr(intLargCidade) & "px' align='left' valign='bottom' nowrap><span class='R'>Cidade</span></td>" & chr(13)
    elseif rb_saida = COD_SAIDA_REL_ORIGEM_PEDIDO then
        strColSpanTodasColunas = "colspan='17'" '14+3
        cab = cab & _
            "       <td class='MDTE' style='width:220px' align='left' valign='bottom' nowrap><span class='R'>Origem Pedido</span></td>" & chr(13) 
    elseif rb_saida = COD_SAIDA_REL_LOJA then
        strColSpanTodasColunas = "colspan='17'" '14+3
        cab = cab & _
            "       <td class='MDTE' style='width::" & CStr(intLargLoja) & "px' align='center' valign='bottom' nowrap><span class='R'>Loja</span></td>" & chr(13) 
    elseif rb_saida = COD_SAIDA_REL_EMPRESA then
	    strColSpanTodasColunas = "colspan='17'" '14+3
	    cab = cab & _
	        "		<td class='MDTE' style='width:" & CStr(intLargCodProduto) & "px' align='left' valign='bottom' nowrap><span class='R'> Código </span></td>" & chr(13) & _
	        "		<td class='MTD' style='width:" & CStr(intLargDescrProduto) & "px' align='left' valign='bottom' nowrap><span class='R'>Descrição</span></td>" & chr(13)
	end if
		
	cab = cab & _
            "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px' align='right' valign='bottom' nowrap><span class='Rd' style='font-weight:bold;'>Qtde</span></td>" & chr(13) & _
			"		<td class='MTD' style='width:" & CStr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Venda (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
			"		<td class='MTD' style='width:" & CStr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Fat Venda Total</span></td>" & chr(13) & _
			"		<td class='MTD' style='width:" & CStr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Lista (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
			"		<td class='MTD' style='width:" & CStr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Desc (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
			"		<td class='MTD' style='width:" & CStr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Desc</span></td>" & chr(13) & _
			"		<td class='MTD' style='width:" & CStr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>COM (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
			"		<td class='MTD' style='width:" & CStr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% COM</span></td>" & chr(13) & _
			"		<td class='MTD' style='width:" & CStr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>%COM<br>+<br>%Desc</span></td>" & chr(13) & _
			"		<td class='MTD' style='width:" & CStr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>RA Líquido (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) 
			if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
			    cab = cab & _
			        "		<td class='MTD' style='width:" & CStr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto -RA -COM (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
			        "		<td class='MTD' style='width:" & CStr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib (%)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & CStr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & CStr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & CStr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta (%)</span></td>" & chr(13)
			else
			    cab = cab & _
			        "		<td class='MTD' style='width:" & CStr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado -RA -COM (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
			        "		<td class='MTD' style='width:" & CStr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib Estimada (%)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & CStr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & CStr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & CStr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta Estimada (%)</span></td>" & chr(13)
			end if
			cab = cab & _
			    "	</tr>" & chr(13)
	
    redim vTotGrupo(0)
	set vTotGrupo(UBound(vTotGrupo)) = new cl_VINTE_COLUNAS
	vTotGrupo(UBound(vTotGrupo)).CampoOrdenacao = ""
    vTotGrupo(UBound(vTotGrupo)).c1 = "-x-x-x-" 'código grupo
    vTotGrupo(UBound(vTotGrupo)).c2 = "" 'descrição grupo
    vTotGrupo(UBound(vTotGrupo)).c3 = 0 'qtde
    vTotGrupo(UBound(vTotGrupo)).c4 = 0 'valor venda
    vTotGrupo(UBound(vTotGrupo)).c5 = 0 '%fat venda total
    vTotGrupo(UBound(vTotGrupo)).c6 = 0 'valor lista
    vTotGrupo(UBound(vTotGrupo)).c7 = 0 'desc R$
    vTotGrupo(UBound(vTotGrupo)).c8 = 0 'desc %
    vTotGrupo(UBound(vTotGrupo)).c9 = 0 'COM R$
    vTotGrupo(UBound(vTotGrupo)).c10 = 0 'COM %
    vTotGrupo(UBound(vTotGrupo)).c11 = 0 '%COM + %Desc
    vTotGrupo(UBound(vTotGrupo)).c12 = 0 'RA líquido
    if rb_saida = COD_SAIDA_REL_GRUPO_PRODUTO then
        vTotGrupo(UBound(vTotGrupo)).c13 = 0 'Lucro Líquido
        vTotGrupo(UBound(vTotGrupo)).c14 = 0 'Margem Contrib %
        vTotGrupo(UBound(vTotGrupo)).c15 = 0 'valor entrada
		vTotGrupo(UBound(vTotGrupo)).c16 = 0 'Valor NF
		vTotGrupo(UBound(vTotGrupo)).c17 = 0 'Lucro Bruto (R$)
		vTotGrupo(UBound(vTotGrupo)).c18 = 0 'Margem Contrib Bruta (%)
    end if

	redim v_total_UF(0)
	set v_total_UF(UBound(v_total_UF)) = new cl_TRES_COLUNAS
	v_total_UF(UBound(v_total_UF)).c1 = ""
	v_total_UF(UBound(v_total_UF)).c2 = 0
	v_total_UF(UBound(v_total_UF)).c3 = 0
	vl_total_final_venda = 0
	vl_total_final_NF = 0
	vl_total_final_lista = 0
	
	' TOTAL FABRICANTE
	redim v_total_FABRICANTE(0)
	set v_total_FABRICANTE(UBound(v_total_FABRICANTE)) = new cl_TRES_COLUNAS
	v_total_FABRICANTE(UBound(v_total_FABRICANTE)).c1 = ""
	v_total_FABRICANTE(UBound(v_total_FABRICANTE)).c2 = 0
	v_total_FABRICANTE(UBound(v_total_FABRICANTE)).c3 = 0

    ' TOTAL GRUPO DE PRODUTOS
    redim v_total_GRUPO_PRODUTO(0)
    set v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)) = new cl_DEZ_COLUNAS
    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c1 = ""
    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c2 = 0
    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c3 = 0
    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c4 = 0
    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c5 = 0
    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c6 = 0
    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c7 = 0

    ' TOTAL EMPRESA
	redim v_total_Empresa(0)
	set v_total_Empresa(UBound(v_total_Empresa)) = new cl_TRES_COLUNAS
	v_total_Empresa(UBound(v_total_Empresa)).c1 = ""
	v_total_Empresa(UBound(v_total_Empresa)).c2 = 0
	v_total_Empresa(UBound(v_total_Empresa)).c3 = 0

    ' TOTAL GRUPO ORIGEM DE PEDIDO
    redim v_total_origem_pedido(0)
    set v_total_origem_pedido(UBound(v_total_origem_pedido)) = new cl_TRES_COLUNAS
    v_total_origem_pedido(UBound(v_total_origem_pedido)).c1 = ""
    v_total_origem_pedido(UBound(v_total_origem_pedido)).c2 = 0
	v_total_origem_pedido(UBound(v_total_origem_pedido)).c3 = 0
	
	set r = cn.execute(s_sql)
	n_reg_BD = 0
	do while Not r.Eof
		n_reg_BD = n_reg_BD + 1
		vl_total_final_venda = vl_total_final_venda + r("valor_venda")
		vl_total_final_NF = vl_total_final_NF + r("valor_NF")
		vl_total_final_lista = vl_total_final_lista + r("valor_lista")
		if rb_saida = COD_SAIDA_REL_CIDADE_UF then
			blnAchou = False
			for i=LBound(v_total_UF) to UBound(v_total_UF)
				if v_total_UF(i).c1 = Trim("" & r("uf")) then
					blnAchou = True
					intIdxSelecionado = i
					exit for
					end if
				next
			if Not blnAchou then
				redim preserve v_total_UF(UBound(v_total_UF)+1)
				set v_total_UF(UBound(v_total_UF)) = new cl_TRES_COLUNAS
				v_total_UF(UBound(v_total_UF)).c1 = Trim("" & r("uf"))
				v_total_UF(UBound(v_total_UF)).c2 = 0
				v_total_UF(UBound(v_total_UF)).c3 = 0
				intIdxSelecionado = UBound(v_total_UF)
				end if
			v_total_UF(intIdxSelecionado).c2 = v_total_UF(intIdxSelecionado).c2 + r("valor_venda")
			v_total_UF(intIdxSelecionado).c3 = v_total_UF(intIdxSelecionado).c3 + r("valor_NF")
			end if 
		'ENCONTRAR OS FABRICANTES
		if rb_saida = COD_SAIDA_REL_FABRICANTE then
		blnAchou = False
			for i=LBound(v_total_FABRICANTE) to UBound(v_total_FABRICANTE)
				if v_total_FABRICANTE(i).c1 = Trim("" & r("fabricante")) then
					blnAchou = True
					strAuxVlFabricante = i
					exit for
					end if
				next
		 if Not blnAchou then
		   redim preserve v_total_FABRICANTE(UBound(v_total_FABRICANTE)+1)
				set v_total_FABRICANTE(UBound(v_total_FABRICANTE)) = new cl_TRES_COLUNAS
				v_total_FABRICANTE(UBound(v_total_FABRICANTE)).c1 = Trim("" & r("fabricante"))
				v_total_FABRICANTE(UBound(v_total_FABRICANTE)).c2 = 0
				v_total_FABRICANTE(UBound(v_total_FABRICANTE)).c3 = 0
				strAuxVlFabricante = UBound(v_total_FABRICANTE)
				 end if
			v_total_FABRICANTE(strAuxVlFabricante).c2 = v_total_FABRICANTE(strAuxVlFabricante).c2 + r("valor_venda")
			v_total_FABRICANTE(strAuxVlFabricante).c3 = v_total_FABRICANTE(strAuxVlFabricante).c3 + r("valor_NF")
		 end if

        ' ENCONTRAR FABRICANTES E GRUPOS DE PRODUTOS
        if rb_saida = COD_SAIDA_REL_GRUPO_PRODUTO then
        blnAchou = False
			for i=LBound(v_total_FABRICANTE) to UBound(v_total_FABRICANTE)
				if v_total_FABRICANTE(i).c1 = Trim("" & r("fabricante")) then
					blnAchou = True
					strAuxVlFabricante = i
					exit for
					end if
				next
		 if Not blnAchou then
		   redim preserve v_total_FABRICANTE(UBound(v_total_FABRICANTE)+1)
				set v_total_FABRICANTE(UBound(v_total_FABRICANTE)) = new cl_TRES_COLUNAS
				v_total_FABRICANTE(UBound(v_total_FABRICANTE)).c1 = Trim("" & r("fabricante"))
				v_total_FABRICANTE(UBound(v_total_FABRICANTE)).c2 = 0
				v_total_FABRICANTE(UBound(v_total_FABRICANTE)).c3 = 0
				strAuxVlFabricante = UBound(v_total_FABRICANTE)
				 end if
			v_total_FABRICANTE(strAuxVlFabricante).c2 = v_total_FABRICANTE(strAuxVlFabricante).c2 + r("valor_venda")
			v_total_FABRICANTE(strAuxVlFabricante).c3 = v_total_FABRICANTE(strAuxVlFabricante).c3 + r("valor_NF")
		blnAchou = False
			for i=LBound(v_total_GRUPO_PRODUTO) to UBound(v_total_GRUPO_PRODUTO)
				if v_total_GRUPO_PRODUTO(i).c1 = Trim("" & r("fabricante")) & Trim("" & r("grupo")) then
					blnAchou = True
					strAuxVlProdutoGrupo = i
					exit for
					end if
				next
		     if Not blnAchou then
		       redim preserve v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)+1)
				    set v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)) = new cl_DEZ_COLUNAS
				    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c1 = Trim("" & r("fabricante")) & Trim("" & r("grupo"))
                    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c2 = 0 'Valor Venda
                    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c3 = 0 'Valor Entrada
                    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c4 = 0 'Comissão (R$)
                    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c5 = 0 'Margem Contrib (%)
                    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c6 = 0 'Valor NF
                    v_total_GRUPO_PRODUTO(UBound(v_total_GRUPO_PRODUTO)).c7 = 0 'Margem Contrib Bruta (%)
				    strAuxVlProdutoGrupo = UBound(v_total_GRUPO_PRODUTO)
            end if
            
			v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c2 = v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c2 + r("valor_venda")
			v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c6 = v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c6 + r("valor_NF")
			if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
                v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3 = v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3 + r("valor_entrada")
			    v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c4 = v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c4 + ((r("perc_RT")/100) * r("valor_venda"))
                'Margem Contrib (%)
				if v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3 = 0 then
                    v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c5 = 0
                else
                    if v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c2 = 0 then
                        v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c5 = 0
                    else
                        v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c5 = 100 * (v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c2 - v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3 - (v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c4)) / v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c2
                    end if
                end if
				'Margem Contrib Bruta (%)
				if v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3 = 0 then
					v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c7 = 0
				else
					if v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c6 = 0 then
						v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c7 = 0
					else
						v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c7 = 100 * ((v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c6 - v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3) / v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c6)
						end if
					end if
			else 'if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA
				v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3 = v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3 + r("valor_entrada")
				v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c4 = v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c4 + ((r("perc_RT")/100) * r("valor_venda"))
				'Margem Contrib (%)
				if v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3 = 0 then
					v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c5 = 0
				else
					if v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c2 = 0 then
						v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c5 = 0
					else
						v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c5 = 100 * (v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c2 - v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3 - (v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c4)) / v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c2
						end if
					end if
				'Margem Contrib Bruta (%)
				if v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3 = 0 then
					v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c7 = 0
				else
					if v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c6 = 0 then
						v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c7 = 0
					else
						v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c7 = 100 * ((v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c6 - v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c3) / v_total_GRUPO_PRODUTO(strAuxVlProdutoGrupo).c6)
						end if
					end if
				end if
		 

        ' FAZER TOTALIZAÇÃO GERAL POR GRUPO DE PRODUTO
        strAuxVlProdutoGrupo = 0
        blnAchou = False
        for i=LBound(vTotGrupo) to UBound(vTotGrupo)
            if vTotGrupo(i).c1 = Trim("" & r("grupo")) then
                blnAchou = True
                strAuxVlProdutoGrupo = i
                exit for
            end if
        next 
        if Not blnAchou then
            redim preserve vTotGrupo(UBound(vTotGrupo)+1)
            set vTotGrupo(UBound(vTotGrupo)) = New cl_VINTE_COLUNAS
            vTotGrupo(UBound(vTotGrupo)).CampoOrdenacao = ""
            vTotGrupo(UBound(vTotGrupo)).c1 = Trim("" & r("grupo")) 'código grupo
            vTotGrupo(UBound(vTotGrupo)).c2 = Trim("" & r("grupo_descricao")) 'descrição grupo
            vTotGrupo(UBound(vTotGrupo)).c3 = 0 'qtde
            vTotGrupo(UBound(vTotGrupo)).c4 = 0 'valor venda
            vTotGrupo(UBound(vTotGrupo)).c5 = 0 '%fat venda total
            vTotGrupo(UBound(vTotGrupo)).c6 = 0 'valor lista
            vTotGrupo(UBound(vTotGrupo)).c7 = 0 'desc R$
            vTotGrupo(UBound(vTotGrupo)).c8 = 0 'desc %
            vTotGrupo(UBound(vTotGrupo)).c9 = 0 'COM R$
            vTotGrupo(UBound(vTotGrupo)).c10 = 0 'COM %
            vTotGrupo(UBound(vTotGrupo)).c11 = 0 '%COM + %Desc
            vTotGrupo(UBound(vTotGrupo)).c12 = 0 'RA líquido
            if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
                vTotGrupo(UBound(vTotGrupo)).c13 = 0 'Lucro Líquido
                vTotGrupo(UBound(vTotGrupo)).c14 = 0 'Margem Contrib %
                vTotGrupo(UBound(vTotGrupo)).c15 = 0 'Vl entrada
				vTotGrupo(UBound(vTotGrupo)).c16 = 0 'VL NF
				vTotGrupo(UBound(vTotGrupo)).c17 = 0 'Lucro Bruto (R$)
				vTotGrupo(UBound(vTotGrupo)).c18 = 0 'Margem Contrib Bruta (%)
            else
                vTotGrupo(UBound(vTotGrupo)).c13 = 0 'Lucro Líquido
                vTotGrupo(UBound(vTotGrupo)).c14 = 0 'Margem Contrib %
                vTotGrupo(UBound(vTotGrupo)).c15 = 0 'Vl entrada
				vTotGrupo(UBound(vTotGrupo)).c16 = 0 'VL NF
				vTotGrupo(UBound(vTotGrupo)).c17 = 0 'Lucro Bruto (R$)
				vTotGrupo(UBound(vTotGrupo)).c18 = 0 'Margem Contrib Bruta (%)
			end if
            strAuxVlProdutoGrupo = UBound(vTotGrupo)
        end if
            
        '> RA LÍQUIDO (R$)
        vl_aux_RA_liquido = 0
		vl_RA_bruto = r("valor_NF")-r("valor_venda")
		if Not calcula_total_RA_liquido(r("perc_desagio_RA_liquida"), vl_RA_bruto, vl_aux_RA_liquido) then
			Response.Write "FALHA AO CALCULAR O RA LÍQUIDO"
			Response.End
		end if
        
        vTotGrupo(strAuxVlProdutoGrupo).c3 = vTotGrupo(strAuxVlProdutoGrupo).c3 + r("qtde")
        vTotGrupo(strAuxVlProdutoGrupo).c4 = vTotGrupo(strAuxVlProdutoGrupo).c4 + r("valor_venda")
        vTotGrupo(strAuxVlProdutoGrupo).c5 = 0           
        vTotGrupo(strAuxVlProdutoGrupo).c6 = vTotGrupo(strAuxVlProdutoGrupo).c6 + r("valor_lista")
        vTotGrupo(strAuxVlProdutoGrupo).c7 = vTotGrupo(strAuxVlProdutoGrupo).c7 + (r("valor_lista") - r("valor_venda"))
        vTotGrupo(strAuxVlProdutoGrupo).c8 = 0
        vTotGrupo(strAuxVlProdutoGrupo).c9 = vTotGrupo(strAuxVlProdutoGrupo).c9 + ((r("perc_RT")/100) * r("valor_venda"))
        vTotGrupo(strAuxVlProdutoGrupo).c10 = 0
        vTotGrupo(strAuxVlProdutoGrupo).c11 = 0
        vTotGrupo(strAuxVlProdutoGrupo).c12 = vTotGrupo(strAuxVlProdutoGrupo).c12 + vl_aux_RA_liquido
        if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then    
			'Lucro Líquido
            vTotGrupo(strAuxVlProdutoGrupo).c13 = vTotGrupo(strAuxVlProdutoGrupo).c13 + (r("valor_venda")-((r("perc_RT")/100) * r("valor_venda"))-r("valor_entrada"))
            vTotGrupo(strAuxVlProdutoGrupo).c14 = 0
            vTotGrupo(strAuxVlProdutoGrupo).c15 = vTotGrupo(strAuxVlProdutoGrupo).c15 + r("valor_entrada")
			vTotGrupo(strAuxVlProdutoGrupo).c16 = vTotGrupo(strAuxVlProdutoGrupo).c16 + r("valor_NF")
			'Margem Contrib %
			if vTotGrupo(strAuxVlProdutoGrupo).c4 = 0 then
                    vTotGrupo(strAuxVlProdutoGrupo).c14 = 0
                else
                    vTotGrupo(strAuxVlProdutoGrupo).c14 = 100 * ((vTotGrupo(strAuxVlProdutoGrupo).c4 - vTotGrupo(strAuxVlProdutoGrupo).c15 - vTotGrupo(strAuxVlProdutoGrupo).c9) / vTotGrupo(strAuxVlProdutoGrupo).c4)
                end if
			'Lucro Bruto (R$)
			vTotGrupo(strAuxVlProdutoGrupo).c17 = vTotGrupo(strAuxVlProdutoGrupo).c17 + (r("valor_NF") - r("valor_entrada"))
			'Margem Contrib Bruta (%)
			if vTotGrupo(strAuxVlProdutoGrupo).c16 = 0 then
				vTotGrupo(strAuxVlProdutoGrupo).c18 = 0
			else
				vTotGrupo(strAuxVlProdutoGrupo).c18 = 100 * ((vTotGrupo(strAuxVlProdutoGrupo).c16 - vTotGrupo(strAuxVlProdutoGrupo).c15) / vTotGrupo(strAuxVlProdutoGrupo).c16)
				end if
            if ckb_ordenar_marg_contrib = "1" then
                ' ordena pela marg contrib bruta
                vTotGrupo(strAuxVlProdutoGrupo).CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vTotGrupo(strAuxVlProdutoGrupo).c18)), 20)
            else
                ' ordena pelo valor NF
                vTotGrupo(strAuxVlProdutoGrupo).CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vTotGrupo(strAuxVlProdutoGrupo).c16)), 20)
            end if
        else
            vTotGrupo(strAuxVlProdutoGrupo).c13 = vTotGrupo(strAuxVlProdutoGrupo).c13 + (r("valor_venda")-((r("perc_RT")/100) * r("valor_venda"))-r("valor_entrada"))
            vTotGrupo(strAuxVlProdutoGrupo).c14 = 0
			vTotGrupo(strAuxVlProdutoGrupo).c15 = vTotGrupo(strAuxVlProdutoGrupo).c15 + r("valor_entrada")
            if vTotGrupo(strAuxVlProdutoGrupo).c4 = 0 then
                vTotGrupo(strAuxVlProdutoGrupo).c14 = 0
            else
                vTotGrupo(strAuxVlProdutoGrupo).c14 = 100 * ((vTotGrupo(strAuxVlProdutoGrupo).c4 - vTotGrupo(strAuxVlProdutoGrupo).c15 - vTotGrupo(strAuxVlProdutoGrupo).c9) / vTotGrupo(strAuxVlProdutoGrupo).c4)
            end if
			vTotGrupo(strAuxVlProdutoGrupo).c16 = vTotGrupo(strAuxVlProdutoGrupo).c16 + r("valor_NF")
			'Lucro Bruto
			vTotGrupo(strAuxVlProdutoGrupo).c17 = vTotGrupo(strAuxVlProdutoGrupo).c17 + (r("valor_NF") - r("valor_entrada"))
			'Margem Contrib Bruta (%)
			if vTotGrupo(strAuxVlProdutoGrupo).c16 = 0 then
				vTotGrupo(strAuxVlProdutoGrupo).c18 = 0
			else
				vTotGrupo(strAuxVlProdutoGrupo).c18 = 100 * ((vTotGrupo(strAuxVlProdutoGrupo).c16 - vTotGrupo(strAuxVlProdutoGrupo).c15) / vTotGrupo(strAuxVlProdutoGrupo).c16)
				end if

            if ckb_ordenar_marg_contrib = "1" then
                ' ordena pela marg contrib bruta
                vTotGrupo(strAuxVlProdutoGrupo).CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vTotGrupo(strAuxVlProdutoGrupo).c18)), 20)
            else
				' ordena pelo valor NF
				vTotGrupo(strAuxVlProdutoGrupo).CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vTotGrupo(strAuxVlProdutoGrupo).c16)), 20)
				end if
        end if
        ordena_cl_vinte_colunas vTotGrupo, 1, Ubound(vTotGrupo)

    end if

        ' ENCONTRAR OS GRUPOS DE ORIGEM DO PEDIDO
        if rb_saida = COD_SAIDA_REL_ORIGEM_PEDIDO then
		blnAchou = False
			for i=LBound(v_total_origem_pedido) to UBound(v_total_origem_pedido)
				if v_total_origem_pedido(i).c1 = Trim("" & r("codigo_pai")) then
					blnAchou = True
					strAuxCodigoPai = i
					exit for
					end if
				next
		 if Not blnAchou then
		   redim preserve v_total_origem_pedido(UBound(v_total_origem_pedido)+1)
				set v_total_origem_pedido(UBound(v_total_origem_pedido)) = new cl_TRES_COLUNAS
				v_total_origem_pedido(UBound(v_total_origem_pedido)).c1 = Trim("" & r("codigo_pai"))
				v_total_origem_pedido(UBound(v_total_origem_pedido)).c2 = 0
				v_total_origem_pedido(UBound(v_total_origem_pedido)).c3 = 0
				strAuxCodigoPai = UBound(v_total_origem_pedido)
				 end if
			v_total_origem_pedido(strAuxCodigoPai).c2 = v_total_origem_pedido(strAuxCodigoPai).c2 + r("valor_venda")
			v_total_origem_pedido(strAuxCodigoPai).c3 = v_total_origem_pedido(strAuxCodigoPai).c3 + r("valor_NF")
		 end if

        'ENCONTRAR AS EMPRESAS
		if rb_saida = COD_SAIDA_REL_EMPRESA then
		blnAchou = False
			for i=LBound(v_total_Empresa) to UBound(v_total_Empresa)
				if v_total_Empresa(i).c1 = Trim("" & r("id_nfe_emitente")) then
					blnAchou = True
					strAuxVlEmpresa = i
					exit for
					end if
				next
		 if Not blnAchou then
		   redim preserve v_total_Empresa(UBound(v_total_Empresa)+1)
				set v_total_Empresa(UBound(v_total_Empresa)) = new cl_TRES_COLUNAS
				v_total_Empresa(UBound(v_total_Empresa)).c1 = Trim("" & r("id_nfe_emitente"))
				v_total_Empresa(UBound(v_total_Empresa)).c2 = 0
				v_total_Empresa(UBound(v_total_Empresa)).c3 = 0
				strAuxVlEmpresa = UBound(v_total_Empresa)
				 end if
			v_total_Empresa(strAuxVlEmpresa).c2 = v_total_Empresa(strAuxVlEmpresa).c2 + r("valor_venda")
			v_total_Empresa(strAuxVlEmpresa).c3 = v_total_Empresa(strAuxVlEmpresa).c3 + r("valor_NF")
		 end if

		r.MoveNext
		loop

			
	if n_reg_BD > 0 then r.MoveFirst
	
	n_reg_BD = 0
	intQtdeTotalProdutos = 0
	vl_total_venda = 0
	vl_total_NF = 0
	vl_total_lista = 0
	vl_total_desconto = 0
	vl_total_RT = 0
	vl_total_RA_liquido = 0
	lucro_liquido_total = 0
	lucro_bruto_total = 0
	total_entrada = 0

	redim vRelat(1)
	for intIdxVetor = Lbound(vRelat) to Ubound(vRelat)
		set vRelat(intIdxVetor) = New cl_VINTE_COLUNAS
		vRelat(intIdxVetor).CampoOrdenacao = ""
		next

'	EM UMA ÚNICA LINHA POR PRODUTO/VENDEDOR/INDICADOR/UF
	strItemAnterior = "--XX--XX--XX--XX--XX--XX--XX--XX--XX--XX--XX--XX--XX--XX--XX--"
	strAuxUFAnterior = "XX"
	strAuxIndicadorAnterior = "XXXXXXXXXXXX"
	strAuxCidadeAnterior = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
	strAuxFabricanteAnterior="XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    strAuxCodigoPaiAnterior = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    strAuxEmpresaAnterior="XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'	A VARIÁVEL ABAIXO É UTILIZADA PARA A ORDENAÇÃO POR ESTADO FICAR EM ORDEM CRESCENTE POR SIGLA
	intIdxUF = 99
'   ORDENAÇÃO POR FABRICANTE
    intIdxFabricante = 99
'   ORDENAÇÃO POR EMPRESA
    intIdxEmpresa = 99

'	PERCORRE OS REGISTROS OBTIDOS DO BD E CONSOLIDA OS VALORES
	do while Not r.Eof

		if rb_saida = COD_SAIDA_REL_PRODUTO then
			strItemAtual = Trim("" & r("fabricante")) & "|" & Trim("" & r("produto")) & "|" & Trim("" & r("descricao")) & "|" & Trim("" & r("descricao_html"))
		
		elseif rb_saida = COD_SAIDA_REL_FABRICANTE then
		    strItemAtual = Trim("" & r("fabricante"))& "|" & Trim("" & r("descricao"))& "|" & Trim("" & r("produto"))& "|" & Trim("" & r("descricao_html"))
		    strAuxNomeFabricante = Ucase(Trim("" & r("fabricante")))
        elseif rb_saida = COD_SAIDA_REL_GRUPO_PRODUTO then
		    strItemAtual = Trim("" & r("fabricante")) & "|" & Trim("" & r("descricao")) & "|" & Trim("" & r("grupo")) & "|" & Trim("" & r("grupo_descricao")) & "|" & Trim("" & r("produto"))& "|" & Trim("" & r("descricao_html"))
		    strAuxNomeFabricante = Ucase(Trim("" & r("fabricante")))
            strAuxNomeProdutoGrupo = Trim("" & r("grupo"))
		elseif rb_saida = COD_SAIDA_REL_VENDEDOR then
			strItemAtual = Trim("" & r("vendedor"))
		elseif rb_saida = COD_SAIDA_REL_INDICADOR then
			strItemAtual = Trim("" & r("indicador")) & "|" & Trim("" & r("uf"))
		elseif rb_saida = COD_SAIDA_REL_UF then
			strItemAtual = Trim("" & r("uf"))
		elseif rb_saida = COD_SAIDA_REL_INDICADOR_UF then
			strItemAtual = Trim("" & r("uf")) & "|" & Trim("" & r("indicador"))
			strAuxIndicador = UCase(Trim("" & r("indicador")))
			strAuxUF = UCase(Trim("" & r("uf")))
		elseif rb_saida = COD_SAIDA_REL_CIDADE_UF then
			strItemAtual = UCase(Trim("" & r(0))) & "|" & iniciais_em_maiusculas(retira_acentuacao(Trim("" & r(1))))
			strAuxCidade = iniciais_em_maiusculas(retira_acentuacao(Trim("" & r(1))))
			strAuxUF = UCase(Trim("" & r(0)))
        elseif rb_saida = COD_SAIDA_REL_ORIGEM_PEDIDO then
            strItemAtual = Trim("" & r("codigo_pai")) & "|" & Trim("" & r("descricao_origem"))
            strAuxCodigoPai = Trim("" & r("codigo_pai"))
            strAuxCodigo = Trim("" & r("codigo_origem"))
        elseif rb_saida = COD_SAIDA_REL_LOJA then
            strItemAtual = Trim("" & r("loja"))
       elseif rb_saida = COD_SAIDA_REL_EMPRESA then
		    strItemAtual = Trim("" & r("id_nfe_emitente"))& "|" & Trim("" & r("fabricante"))& "|" & Trim("" & r("descricao"))& "|" & Trim("" & r("produto"))& "|" & Trim("" & r("descricao_html"))
		    strAuxNomeEmpresa = Ucase(Trim("" & r("id_nfe_emitente")))
		else
			strItemAtual = ""
			end if
		
	'	MUDOU DE ITEM?
		if strItemAtual <> strItemAnterior then
			
			if vRelat(Ubound(vRelat)).CampoOrdenacao <> "" then
				redim preserve vRelat(Ubound(vRelat)+1)
				set vRelat(Ubound(vRelat)) = New cl_VINTE_COLUNAS
				vRelat(Ubound(vRelat)).CampoOrdenacao = ""
				end if
				
			if n_reg_BD > 0 then
				if rb_saida = COD_SAIDA_REL_INDICADOR_UF then
					strCampoOrdenacao = Trim(CStr(intIdxUF)) & "|" & strAuxUFAnterior & normaliza_codigo(retorna_so_digitos(formata_moeda(vl_sub_total_NF)), 20) & "|" & _
										strAuxIndicadorAnterior
					if strAuxUF <> strAuxUFAnterior then 
						intIdxUF = intIdxUF - 1
						end if
				elseif rb_saida = COD_SAIDA_REL_CIDADE_UF then
					vl_aux = 0
					for i=LBound(v_total_UF) to UBound(v_total_UF)
						if v_total_UF(i).c1 = strAuxUFAnterior then
							vl_aux = v_total_UF(i).c3
							exit for
							end if
						next
					strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & strAuxUFAnterior & normaliza_codigo(retorna_so_digitos(formata_moeda(vl_sub_total_NF)), 20) & "|" & _
										strAuxCidadeAnterior
					if strAuxUF <> strAuxUFAnterior then 
						intIdxUF = intIdxUF - 1
						end if

                elseif rb_saida = COD_SAIDA_REL_ORIGEM_PEDIDO then
					vl_aux = 0
					for i=LBound(v_total_origem_pedido) to UBound(v_total_origem_pedido)
						if v_total_origem_pedido(i).c1 = strauxCodigoPaiAnterior then
							vl_aux = v_total_origem_pedido(i).c3
							exit for
							end if
						next
					strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & strauxCodigoPaiAnterior & normaliza_codigo(retorna_so_digitos(formata_moeda(vl_sub_total_NF)), 20) & "|" & _
										strauxCodigoPaiAnterior
					if strauxCodigoPai <> strauxCodigoPaiAnterior then 
						intIdxUF = intIdxUF - 1
						end if

		'ORDENA OS ITEMS FABRICANTE
			elseif rb_saida = COD_SAIDA_REL_FABRICANTE  then
			        vl_aux = 0
					for i=LBound(v_total_FABRICANTE) to UBound(v_total_FABRICANTE)
						if v_total_FABRICANTE(i).c1 = strAuxNomeFabricanteAnterior then
							vl_aux = v_total_FABRICANTE(i).c3
							exit for
							end if
						next
                    if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
                        strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & strAuxNomeFabricanteAnterior & normaliza_codigo(retorna_so_digitos(intQtdeSubTotalProdutos), 20) & "|" &  normaliza_codigo(retorna_so_digitos(codigoProdutoComplemento(Trim("" & r("produto")))), 20) & "|" & _
										strAuxFabricanteAnterior
                    else
					    strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & strAuxNomeFabricanteAnterior & normaliza_codigo(retorna_so_digitos(intQtdeSubTotalProdutos), 20) & "|" &  normaliza_codigo(retorna_so_digitos(codigoProdutoComplemento(Trim("" & r("produto")))), 20) & "|" & _
										strAuxFabricanteAnterior
                     end if
					if strAuxNomeFabricante <> strAuxNomeFabricanteAnterior then 
						intIdxFABRICANTE = intIdxFABRICANTE - 1
						end if

        'ORDENA OS ITENS FABRICANTE/GRUPO PRODUTO
			elseif rb_saida = COD_SAIDA_REL_GRUPO_PRODUTO then
                    if vl_sub_total_venda = 0 then
				        marg_contrib_ordenacao = 0
				    else
				        marg_contrib_ordenacao = 100 * ((vl_sub_total_venda - subtotal_entrada - vl_sub_total_RT) / vl_sub_total_venda)
				    end if

					if vl_sub_total_NF = 0 then
						marg_contrib_bruta_ordenacao = 0
					else
						marg_contrib_bruta_ordenacao = 100 * ((vl_sub_total_NF - subtotal_entrada) / vl_sub_total_NF)
						end if

			        vl_aux = 0
                    vl_aux_ProdutoGrupo = 0
                    for i=LBound(v_total_FABRICANTE) to UBound(v_total_FABRICANTE)
						if v_total_FABRICANTE(i).c1 = strAuxNomeFabricanteAnterior then
							vl_aux = v_total_FABRICANTE(i).c3
							exit for
							end if
						next
					for i=LBound(v_total_GRUPO_PRODUTO) to UBound(v_total_GRUPO_PRODUTO)
						if v_total_GRUPO_PRODUTO(i).c1 = strAuxNomeFabricanteAnterior & strAuxNomeProdutoGrupoAnterior then
							vl_aux_ProdutoGrupo = v_total_GRUPO_PRODUTO(i).c6
                            vl_marg_contrib_bruta_ProdutoGrupo = v_total_GRUPO_PRODUTO(i).c7
							exit for
							end if
						next
                    
                    if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
                        if ckb_ordenar_marg_contrib = "1" then
					        strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & iif(CInt(vl_marg_contrib_bruta_ProdutoGrupo)<0, "1", "2") & normaliza_codigo(retorna_so_digitos(formata_moeda(vl_marg_contrib_bruta_ProdutoGrupo)), 20) & "|" & iif(CInt(marg_contrib_bruta_ordenacao)<0, "1", "2") & normaliza_codigo(retorna_so_digitos(formata_moeda(marg_contrib_bruta_ordenacao)), 20)
                        else
							strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux_ProdutoGrupo)), 20) & "|" & strAuxNomeFabricanteAnterior & normaliza_codigo(retorna_so_digitos(intQtdeSubTotalProdutos), 20) & "|" & _
											strAuxFabricanteAnterior
							end if
                    else
					    if ckb_ordenar_marg_contrib = "1" then
					        strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & iif(CInt(vl_marg_contrib_bruta_ProdutoGrupo)<0, "1", "2") & normaliza_codigo(retorna_so_digitos(formata_moeda(vl_marg_contrib_bruta_ProdutoGrupo)), 20) & "|" & iif(CInt(marg_contrib_bruta_ordenacao)<0, "1", "2") & normaliza_codigo(retorna_so_digitos(formata_moeda(marg_contrib_bruta_ordenacao)), 20)
                        else
                            strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux_ProdutoGrupo)), 20) & "|" & strAuxNomeFabricanteAnterior & normaliza_codigo(retorna_so_digitos(intQtdeSubTotalProdutos), 20) & "|" & _
										strAuxFabricanteAnterior
                        end if
					 end if		
                                    
         'ORDENA OS ITEMS EMPRESA
			    elseif rb_saida = COD_SAIDA_REL_EMPRESA  then
			        vl_aux = 0
					    for i=LBound(v_total_Empresa) to UBound(v_total_Empresa)
						    if v_total_Empresa(i).c1 = strAuxNomeEmpresaAnterior then
							    vl_aux = v_total_Empresa(i).c3
							    exit for
						    end if
					    next
					    if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
                        strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & strAuxNomeEmpresaAnterior & normaliza_codigo(retorna_so_digitos(intQtdeSubTotalProdutos), 20) & "|" &  normaliza_codigo(retorna_so_digitos(codigoProdutoComplemento(Trim("" & r("produto")))), 20) & "|" & _
										    strAuxEmpresaAnterior
                        else
					    strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & strAuxNomeEmpresaAnterior & normaliza_codigo(retorna_so_digitos(intQtdeSubTotalProdutos), 20) & "|" & normaliza_codigo(retorna_so_digitos(codigoProdutoComplemento(Trim("" & r("produto")))), 20) & "|" & _
										    strAuxEmpresaAnterior
                        end if
					    if strAuxNomeEmpresa <> strAuxNomeEmpresaAnterior then 
						    intIdxEmpresa = intIdxEmpresa - 1
					    end if
            
	            
               elseif rb_saida = COD_SAIDA_REL_PRODUTO then
               strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(intQtdeSubTotalProdutos), 20) & "|" &  normaliza_codigo(retorna_so_digitos(codigoProdutoComplemento(Trim("" & r(1)))), 20) & "|" & _
										strItemAnterior

				else
					strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_sub_total_NF)), 20) & _
										strItemAnterior
					end if
				
				with vRelat(Ubound(vRelat))
				'	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
					if vl_total_final_venda = 0 then
						percFatVenda = 0
					else
						percFatVenda = (vl_sub_total_venda/vl_total_final_venda)*100
						end if
					
				'	% DESCONTO
					if vl_sub_total_lista = 0 then
						perc_desconto = 0
					else
						perc_desconto = 100 * (vl_sub_total_lista-vl_sub_total_venda) / vl_sub_total_lista
						end if
				
				'	% COMISSÃO SOBRE FATURAMENTO VENDA
					if vl_sub_total_venda = 0 then
						percRTFatVenda = 0
					else
						percRTFatVenda = (vl_sub_total_RT/vl_sub_total_venda)*100
						end if

				'	%COMISSÃO + %DESCONTO
				'	O CARLOS SOLICITOU EM 12/MAR/2008 QUE SEJA FEITA A SOMA SIMPLES DOS 2 PERCENTUAIS
					percRTEDesc = perc_desconto + percRTFatVenda
					
				'   % MARGEM CONTRIBUIÇÃO
				    if vl_sub_total_venda = 0 then
				        marg_contrib = 0
				    else
				        marg_contrib = 100 * ((vl_sub_total_venda - subtotal_entrada - vl_sub_total_RT) / vl_sub_total_venda)
				    end if

				'	% MARGEM CONTRIBUIÇÃO BRUTA
					if vl_sub_total_NF = 0 then
						marg_contrib_bruta = 0
					else
						marg_contrib_bruta = 100 * ((vl_sub_total_NF - subtotal_entrada) / vl_sub_total_NF)
						end if

					.CampoOrdenacao = strCampoOrdenacao
					.c1 = strItemAnterior
					.c2 = intQtdeSubTotalProdutos
					.c3 = vl_sub_total_venda
					.c4 = percFatVenda
					.c5 = vl_sub_total_lista
					.c6 = vl_sub_total_desconto
					.c7 = perc_desconto
					.c8 = vl_sub_total_RT
					.c9 = percRTFatVenda
					.c10 = percRTEDesc
					.c11 = vl_sub_total_RA_liquido
					.c12 = lucro_liquido_subtotal
					.c13 = marg_contrib
					.c14 = subtotal_entrada
					.c15 = vl_sub_total_NF
					.c16 = lucro_bruto_subtotal
					.c17 = marg_contrib_bruta
					end with
				end if
				
			intQtdeSubTotalProdutos = 0
			vl_sub_total_venda = 0
			vl_sub_total_NF = 0
			vl_sub_total_lista = 0
			vl_sub_total_desconto = 0
			vl_sub_total_RT = 0
			vl_sub_total_RA_liquido = 0
			lucro_liquido_subtotal = 0
			lucro_bruto_subtotal = 0
			subtotal_entrada = 0
			strItemAnterior = strItemAtual
			strAuxUFAnterior = strAuxUF
			strAuxIndicadorAnterior = strAuxIndicador
			strAuxNomeFabricanteAnterior = strAuxNomeFabricante
			strAuxFabricanteAnterior = strAuxFabricante
            strAuxNomeProdutoGrupoAnterior = strAuxNomeProdutoGrupo
            strAuxCodigoPaiAnterior = strAuxCodigoPai
            strAuxNomeEmpresaAnterior = strAuxNomeEmpresa
			strAuxEmpresaAnterior = strAuxEmpresa           
			end if
			
		n_reg_BD = n_reg_BD + 1
		
	 '> VALOR DE VENDA
		vl_venda = r("valor_venda")

	 '> VALOR NF
		vl_NF = r("valor_NF")

	 '> VALOR DE LISTA
		vl_lista = r("valor_lista")

	 '> DESCONTO (R$)
		vl_desconto = vl_lista - vl_venda

	 '> COMISSÃO (R$) (ANTERIORMENTE CHAMADO DE RT)
		vl_RT = (r("perc_RT")/100) * vl_venda

	 '> RA LÍQUIDO (R$)
		vl_RA_bruto = r("valor_NF")-r("valor_venda")
		if Not calcula_total_RA_liquido(r("perc_desagio_RA_liquida"), vl_RA_bruto, vl_RA_liquido) then
			Response.Write "FALHA AO CALCULAR O RA LÍQUIDO"
			Response.End
			end if
			
	 '> VALOR DE ENTRADA (R$)
		if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
			vl_ent = r("valor_entrada")
			
		 '> LUCRO_LÍQUIDO (R$)
			lucro_liquido = vl_venda - vl_RT - vl_ent
		else
			vl_ent = r("valor_entrada")
			
			'> LUCRO_LÍQUIDO (R$)
			lucro_liquido = vl_venda - vl_RT - vl_ent
			end if

	 '> LUCRO BRUTO (R$)
		lucro_bruto = vl_NF - vl_ent

		vl_sub_total_venda = vl_sub_total_venda + vl_venda
		vl_sub_total_NF = vl_sub_total_NF + vl_NF
		vl_sub_total_lista = vl_sub_total_lista + vl_lista
		vl_sub_total_desconto = vl_sub_total_desconto + vl_desconto
		vl_sub_total_RT = vl_sub_total_RT + vl_RT
		vl_sub_total_RA_liquido = vl_sub_total_RA_liquido + vl_RA_liquido
		lucro_liquido_subtotal = lucro_liquido_subtotal + lucro_liquido
		lucro_bruto_subtotal = lucro_bruto_subtotal + lucro_bruto
		subtotal_entrada = subtotal_entrada + vl_ent
		intQtdeSubTotalProdutos = intQtdeSubTotalProdutos + r("qtde")
		
		vl_total_venda = vl_total_venda + vl_venda
		vl_total_NF = vl_total_NF + vl_NF
		vl_total_lista = vl_total_lista + vl_lista
		vl_total_desconto = vl_total_desconto + vl_desconto
		vl_total_RT = vl_total_RT + vl_RT
		vl_total_RA_liquido = vl_total_RA_liquido + vl_RA_liquido
		lucro_liquido_total = lucro_liquido_total + lucro_liquido
		lucro_bruto_total = lucro_bruto_total + lucro_bruto
		total_entrada = total_entrada + vl_ent
		intQtdeTotalProdutos = intQtdeTotalProdutos + r("qtde")
		
		r.MoveNext
		loop


'	TOTAL DO ÚLTIMO ITEM
	if n_reg_BD > 0 then
		if vRelat(Ubound(vRelat)).CampoOrdenacao <> "" then
			redim preserve vRelat(Ubound(vRelat)+1)
			set vRelat(Ubound(vRelat)) = New cl_VINTE_COLUNAS
			vRelat(Ubound(vRelat)).CampoOrdenacao = ""
			end if
	
		if rb_saida = COD_SAIDA_REL_INDICADOR_UF then
			strCampoOrdenacao = Trim(CStr(intIdxUF)) & "|" & strAuxUFAnterior & normaliza_codigo(retorna_so_digitos(formata_moeda(vl_sub_total_NF)), 20) & "|" & _
								strAuxIndicadorAnterior
		elseif rb_saida = COD_SAIDA_REL_CIDADE_UF then
			vl_aux = 0
			for i=LBound(v_total_UF) to UBound(v_total_UF)
				if v_total_UF(i).c1 = strAuxUFAnterior then
					vl_aux = v_total_UF(i).c3
					exit for
					end if
				next
			strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & strAuxUFAnterior & normaliza_codigo(retorna_so_digitos(formata_moeda(vl_sub_total_NF)), 20) & "|" & _
								strAuxCidadeAnterior

        elseif rb_saida = COD_SAIDA_REL_ORIGEM_PEDIDO then
			vl_aux = 0
			for i=LBound(v_total_origem_pedido) to UBound(v_total_origem_pedido)
				if v_total_origem_pedido(i).c1 = strauxCodigoPaiAnterior then
					vl_aux = v_total_origem_pedido(i).c3
					exit for
					end if
				next
			strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & strauxCodigoPaiAnterior & normaliza_codigo(retorna_so_digitos(formata_moeda(intQtdeSubTotalProdutos)), 20) & "|" & _
								strauxCodigoAnterior
    

		elseif rb_saida = COD_SAIDA_REL_FABRICANTE  then
	        vl_aux = 0
			for i=LBound(v_total_FABRICANTE) to UBound(v_total_FABRICANTE)
				if v_total_FABRICANTE(i).c1 = strAuxNomeFabricanteAnterior then
					vl_aux = v_total_FABRICANTE(i).c3
					exit for
					end if
				next
			strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & strAuxNomeFabricanteAnterior & normaliza_codigo(retorna_so_digitos(intQtdeSubTotalProdutos), 20) & "|" & _
								strAuxFabricanteAnterior
                
        
        elseif rb_saida = COD_SAIDA_REL_EMPRESA  then
	        vl_aux = 0
			for i=LBound(v_total_Empresa) to UBound(v_total_Empresa)
				if v_total_Empresa(i).c1 = strAuxNomeEmpresaAnterior then
					vl_aux = v_total_Empresa(i).c3
					exit for
					end if
				next
			strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_aux)), 20) & "|" & strAuxNomeEmpresaAnterior & normaliza_codigo(retorna_so_digitos(intQtdeSubTotalProdutos), 20) & "|"  & _
										strAuxEmpresaAnterior

            elseif rb_saida = COD_SAIDA_REL_PRODUTO then
               strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(intQtdeSubTotalProdutos), 20) & "|"  & _
										strItemAnterior
    		else
			strCampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_sub_total_NF)), 20) & _
								strItemAnterior
			end if
		
		with vRelat(Ubound(vRelat))
		'	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
			if vl_total_final_venda = 0 then
				percFatVenda = 0
			else
				percFatVenda = (vl_sub_total_venda/vl_total_final_venda)*100
				end if
				
		'	% DESCONTO
			if vl_sub_total_lista = 0 then
				perc_desconto = 0
			else
				perc_desconto = 100 * (vl_sub_total_lista-vl_sub_total_venda) / vl_sub_total_lista
				end if
				
		'	% COMISSÃO SOBRE FATURAMENTO VENDA
			if vl_sub_total_venda = 0 then
				percRTFatVenda = 0
			else
				percRTFatVenda = (vl_sub_total_RT/vl_sub_total_venda)*100
				end if

		'	%COMISSÃO + %DESCONTO
		'	O CARLOS SOLICITOU EM 12/MAR/2008 QUE SEJA FEITA A SOMA SIMPLES DOS 2 PERCENTUAIS
			percRTEDesc = perc_desconto + percRTFatVenda
			
	    '    MARGEM CONTRIBUIÇÃO (%)
	        if vl_sub_total_venda = 0 then
	             marg_contrib = 0
	        else
	             marg_contrib = 100 * ((vl_sub_total_venda - subtotal_entrada - vl_sub_total_RT) / vl_sub_total_venda)
	        end if  

		'	MARGEM CONTRIBUIÇÃO BRUTA (%)
			if vl_sub_total_NF = 0 then
				marg_contrib_bruta = 0
			else
				marg_contrib_bruta = 100 * ((vl_sub_total_NF - subtotal_entrada) / vl_sub_total_NF)
				end if

			.CampoOrdenacao = strCampoOrdenacao
			.c1 = strItemAnterior
			.c2 = intQtdeSubTotalProdutos
			.c3 = vl_sub_total_venda
			.c4 = percFatVenda
			.c5 = vl_sub_total_lista
			.c6 = vl_sub_total_desconto
			.c7 = perc_desconto
			.c8 = vl_sub_total_RT
			.c9 = percRTFatVenda
			.c10 = percRTEDesc
			.c11 = vl_sub_total_RA_liquido
			.c12 = lucro_liquido_subtotal
			.c13 = marg_contrib
			.c14 = subtotal_entrada
			.c15 = vl_sub_total_NF
			.c16 = lucro_bruto_subtotal
			.c17 = marg_contrib_bruta
			end with
		end if


'	ORDENA O VETOR COM RESULTADOS
	ordena_cl_vinte_colunas vRelat, 1, Ubound(vRelat)


	n_reg_vetor = 0
	if n_reg_BD > 0 then
		if (rb_saida <> COD_SAIDA_REL_INDICADOR_UF) and (rb_saida <> COD_SAIDA_REL_CIDADE_UF)and(rb_saida <> COD_SAIDA_REL_FABRICANTE)and(rb_saida <> COD_SAIDA_REL_GRUPO_PRODUTO) and (rb_saida <> COD_SAIDA_REL_ORIGEM_PEDIDO) and (rb_saida <> COD_SAIDA_REL_EMPRESA) then
			x = cab_table
			x = x & cab
			end if
		
		strAuxUF = "-XX-XX-"
		strAuxUFAnterior = "-XX-XX-"
	    strAuxNomeFabricante = "-XX-XX-"
	    strAuxNomeFabricanteAnterior = "-XX-XX-"
        strAuxNomeProdutoGrupo = "-XX-XX-"
	    strAuxNomeProdutoGrupoAnterior = "-XX-XX-"
        strAuxCodigoPaiAnterior = "-XX-XX-"
        strAuxCodigoPai = "-XX-XX-"
	    strAuxNomeEmpresa = "-XX-XX-"
	    strAuxNomeEmpresaAnterior = "-XX-XX-"

		for intIdxVetor = Ubound(vRelat) to 1 step -1
		  ' CONTAGEM
			n_reg_vetor = n_reg_vetor + 1

			if (rb_saida <> COD_SAIDA_REL_INDICADOR_UF) and (rb_saida <> COD_SAIDA_REL_CIDADE_UF)and(rb_saida <> COD_SAIDA_REL_FABRICANTE)and(rb_saida <> COD_SAIDA_REL_GRUPO_PRODUTO) and (rb_saida <> COD_SAIDA_REL_ORIGEM_PEDIDO) and (rb_saida <> COD_SAIDA_REL_EMPRESA) then
				'	NUMERAÇÃO DA LINHA
				x = x & "	<tr nowrap>" & chr(13)
				x = x & "		<td valign='bottom' align='right' nowrap><span class='Rd' style='margin-right:2px;'>" & Cstr(n_reg_vetor) & ".</span></td>" & chr(13)
				end if
			
               

			with vRelat(intIdxVetor)
				intQtdeProdutos = .c2
				vl_venda = .c3
				percFatVenda = .c4
				vl_lista = .c5
				vl_desconto = .c6
				perc_desconto = .c7
				vl_RT = .c8
				percRTFatVenda = .c9
				percRTEDesc = .c10
				vl_RA_liquido = .c11
				lucro_liquido = .c12
				marg_contrib = .c13
				vl_ent = .c14
				vl_NF = .c15
				lucro_bruto = .c16
				marg_contrib_bruta = .c17
				end with
				
			s_cor="black"
			if vl_venda < 0 then s_cor="red"

		 '> CAMPO DE SAÍDA
			if rb_saida = COD_SAIDA_REL_PRODUTO then
				v = Split(vRelat(intIdxVetor).c1, "|")
			'> CÓDIGO DO PRODUTO
				if intQtdeProdutos < 0 then s_cor="red"
				x = x & "		<td class='MDTE' style='width:" & CStr(intLargCodProduto) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & Trim("" & v(1)) & "</span></td>" & chr(13)
			'> DESCRIÇÃO DO PRODUTO
				s = Trim("" & v(3))
				if s <> "" then s = produto_formata_descricao_em_html(s) else s = "&nbsp;"
				x = x & "		<td class='MTD' style='width:" & CStr(intLargDescrProduto) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
			'> QUANTIDADE
				x = x & "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px;' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeProdutos) & "</span></td>" & chr(13)
			elseif rb_saida = COD_SAIDA_REL_VENDEDOR then
			'> VENDEDOR
				s = Trim(vRelat(intIdxVetor).c1)
				s_aux = x_usuario(s)
				if (s <> "") And (s_aux <> "") then s = s & " - " & s_aux
				if s = "" then s = "&nbsp;"
				x = x & "		<td class='MDTE' style='width:" & CStr(intLargVendedor) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
            '> QUANTIDADE
				x = x & "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px;' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeProdutos) & "</span></td>" & chr(13)
			elseif rb_saida = COD_SAIDA_REL_INDICADOR then
				v = Split(vRelat(intIdxVetor).c1, "|")
			'> INDICADOR
				s = Trim("" & v(0))
				s_aux = x_orcamentista_e_indicador(v(0))
				if (s <> "") And (s_aux <> "") then s = s & " - " & s_aux
				if s = "" then s = "&nbsp;"
				x = x & "		<td class='MDTE' style='width:" & CStr(intLargIndicador) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)           
			'> UF
				s = UCase(Trim("" & v(1)))
				if s = "" then s = "&nbsp;"
				x = x & "		<td class='MTD' style='width:" & CStr(intLargUF) & "px;' align='center' valign='bottom'><span class='Cnc' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
            '> QUANTIDADE
				x = x & "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px;' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeProdutos) & "</span></td>" & chr(13)
			elseif rb_saida = COD_SAIDA_REL_UF then
			'> UF
				s = UCase(Trim(vRelat(intIdxVetor).c1))
				if s = "" then s = "&nbsp;"
				x = x & "		<td class='MDTE' style='width:" & CStr(intLargUF) & "px;' align='center' valign='bottom'><span class='Cnc' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
            '> QUANTIDADE
				x = x & "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px;' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeProdutos) & "</span></td>" & chr(13)
            elseif rb_saida = COD_SAIDA_REL_LOJA then
			'> LOJA
				s = UCase(Trim(vRelat(intIdxVetor).c1))
				if s = "" then
					s = "&nbsp;"
					s_nome_loja = ""
				else
					s_nome_loja = "<br /><span class='Cnc' style='color:" & s_cor & ";'>(" & x_loja(s) & ")</span>"
					end if
				x = x & "		<td class='MDTE' style='width:" & CStr(intLargLoja) & "px;' align='center' valign='bottom'><span class='Cnc' style='color:" & s_cor & ";'>" & s & "</span>" & s_nome_loja & "</td>" & chr(13)
            '> QUANTIDADE
				x = x & "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px;' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeProdutos) & "</span></td>" & chr(13)
			elseif rb_saida = COD_SAIDA_REL_INDICADOR_UF then
			'	CRIAR QUEBRA PARA A MUDANÇA DE UF
				v = Split(vRelat(intIdxVetor).c1, "|")
				strAuxUFAnterior = strAuxUF
				strAuxIndicadorAnterior = strAuxIndicador
				strAuxUF = UCase(Trim("" & v(0)))
				if strAuxUFAnterior <> strAuxUF then
					if strAuxUFAnterior <> "-XX-XX-" then
					'	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
						if vl_total_final_venda = 0 then
							percFatVenda_UF = 0
						else
							percFatVenda_UF = (vl_venda_UF/vl_total_final_venda)*100
							end if
					
					'	% DESCONTO
						if vl_lista_UF = 0 then
							perc_desconto_UF = 0
						else
							perc_desconto_UF = 100 * (vl_lista_UF-vl_venda_UF) / vl_lista_UF
							end if
					
					'	%COM SOBRE FATURAMENTO VENDA
						if vl_venda_UF = 0 then
							percRTFatVenda_UF = 0
						else
							percRTFatVenda_UF = (vl_RT_UF/vl_venda_UF)*100
							end if
							
					'	%COM + %DESCONTO
						percRTEDesc_UF = perc_desconto_UF + percRTFatVenda_UF
						
					'   %MARGEM CONTRIBUIÇÃO
					    if vl_venda_UF = 0 then
					        marg_contrib_UF = 0
					    else
					        marg_contrib_UF = 100 * ((vl_venda_UF - vl_RT_UF - vl_ent_UF) / vl_venda_UF)
					    end if

						if vl_NF_UF = 0 then
							marg_contrib_bruta_UF = 0
						else
							marg_contrib_bruta_UF = 100 * ((vl_NF_UF - vl_ent_UF) / vl_NF_UF)
							end if

					    md = ""
					    if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
						s_cor="black"
						if vl_total_venda < 0 then s_cor="red"
						x = x & "	</tr>" & chr(13)
						x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
								"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
								"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
								"TOTAL:</span></td>" & chr(13) & _
                                "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalIndicadorUf) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_UF) & "</span></td>" & chr(13) 
								if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_UF) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_UF) & "%" & "</span></td>" & chr(13) & _
										"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_UF) & "%" & "</span></td>" & chr(13) & _
										"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_UF) & "%" & "</span></td>" & chr(13) & _
										"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_UF) & "%" & "</span></td>" & chr(13)
								else
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_UF) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_UF) & "%" & "</span></td>" & chr(13) & _
										"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_UF) & "%" & "</span></td>" & chr(13) & _
										"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_UF) & "%" & "</span></td>" & chr(13) & _
										"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_UF) & "%" & "</span></td>" & chr(13)
								end if
								x = x & _
								"	</tr>" & chr(13) 
						
						x = x & "</table>" & chr(13) 
						x = x & "<br>" & chr(13) 
						end if
					
					x = x & "<table cellspacing=0>" & chr(13)
					x = x & _
						"	<tr>" & chr(13) & _
						"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
						"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;" & strAuxUF & iif(strAuxUF <> "", " &nbsp;-&nbsp; " & UF_descricao(strAuxUF), "") & "</span></td>" & chr(13) & _
						"	</tr>" & chr(13)
					x = x & "<br>" & chr(13)
					x = x & cab
					
					vl_venda_UF = 0
					vl_NF_UF = 0
					percFatVenda_UF = 0
					vl_lista_UF = 0
					vl_desconto_UF = 0
					perc_desconto_UF = 0
					vl_RT_UF = 0
					percRTFatVenda_UF = 0
					percRTEDesc_UF = 0
					vl_RA_liquido_UF = 0
					lucro_liquido_UF = 0
					lucro_bruto_UF = 0
					marg_contrib_UF = 0
					marg_contrib_bruta_UF = 0
					vl_ent_UF = 0
                    intQtdeTotalIndicadorUf=0
					end if
				
			'	FAZENDO A SUBTOTALIZAÇÃO POR UF
                intQtdeTotalIndicadorUf = intQtdeTotalIndicadorUf + intQtdeProdutos
				vl_venda_UF = vl_venda_UF + vl_venda
				vl_NF_UF = vl_NF_UF + vl_NF
				vl_lista_UF = vl_lista_UF +vl_lista
				vl_desconto_UF = vl_desconto_UF + vl_desconto
				vl_RT_UF = vl_RT_UF + vl_RT
				vl_RA_liquido_UF = vl_RA_liquido_UF + vl_RA_liquido
				lucro_liquido_UF = lucro_liquido_UF + lucro_liquido
				lucro_bruto_UF = lucro_bruto_UF + lucro_bruto
				vl_ent_UF = vl_ent_UF + vl_ent
				
			'	NUMERAÇÃO DA LINHA
				x = x & "	<tr nowrap>" & chr(13)
				x = x & "		<td valign='bottom' align='right' nowrap><span class='Rd' style='margin-right:2px;'>" & Cstr(n_reg_vetor) & ".</span></td>" & chr(13)

			'> UF
'				s = UCase(Trim("" & v(0)))
'				if s = "" then s = "&nbsp;"
'				s = s & " <> " & vRelat(intIdxVetor).CampoOrdenacao
            
			'> INDICADOR
				s = Trim("" & v(1))
				s_aux = x_orcamentista_e_indicador(v(1))
				if (s <> "") And (s_aux <> "") then s = s & " - " & s_aux
				if s = "" then s = "&nbsp;"
				x = x & "		<td class='MDTE' style='width:" & CStr(intLargIndicador) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
			'> QUANTIDADE
				x = x & "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px;' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeProdutos) & "</span></td>" & chr(13)
			elseif rb_saida = COD_SAIDA_REL_CIDADE_UF then
			'	CRIAR QUEBRA PARA A MUDANÇA DE UF
				v = Split(vRelat(intIdxVetor).c1, "|")
				strAuxUFAnterior = strAuxUF
				strAuxCidadeAnterior = strAuxCidade
				strAuxUF = UCase(Trim("" & v(0)))
				if strAuxUFAnterior <> strAuxUF then
					if strAuxUFAnterior <> "-XX-XX-" then
					'	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
						if vl_total_final_venda = 0 then
							percFatVenda_UF = 0
						else
							percFatVenda_UF = (vl_venda_UF/vl_total_final_venda)*100
							end if
					
					'	% DESCONTO
						if vl_lista_UF = 0 then
							perc_desconto_UF = 0
						else
							perc_desconto_UF = 100 * (vl_lista_UF-vl_venda_UF) / vl_lista_UF
							end if
					
					'	%COM SOBRE FATURAMENTO VENDA
						if vl_venda_UF = 0 then
							percRTFatVenda_UF = 0
						else
							percRTFatVenda_UF = (vl_RT_UF/vl_venda_UF)*100
							end if

					'	%COM + %DESCONTO
						percRTEDesc_UF = perc_desconto_UF + percRTFatVenda_UF
						
				    '   %MARGEM CONTRIB
				        if vl_venda_UF = 0 then
				            marg_contrib_UF = 0
				        else
                            marg_contrib_UF = 100 * ((vl_venda_UF - vl_RT_UF - vl_ent_UF) / vl_venda_UF)
                        end if

					'	% MARGEM CONTRIB BRUTA
						if vl_NF_UF = 0 then
							marg_contrib_bruta_UF = 0
						else
							marg_contrib_bruta_UF = 100 * ((vl_NF_UF - vl_ent_UF) / vl_NF_UF)
							end if

                        md = ""
					    if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
						s_cor="black"
						if vl_total_venda < 0 then s_cor="red"
						x = x & "	</tr>" & chr(13)
						x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
								"		<td valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
								"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
								"TOTAL:</span></td>" & chr(13) & _
                                 "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalIndicadorCidUf) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_UF) & "</span></td>" & chr(13) 
								if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_UF) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_UF) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_UF) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_UF) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_UF) & "%" & "</span></td>" & chr(13)
								else
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_UF) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_UF) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_UF) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_UF) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_UF) & "%" & "</span></td>" & chr(13)
								end if
								x = x & _
								"	</tr>" & chr(13) 
						
						x = x & "</table>" & chr(13)
						x = x & "<br>" & chr(13)
						end if
					
					x = x & "<table cellspacing=0>" & chr(13)
					x = x & _
						"	<tr>" & chr(13) & _
						"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
						"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;" & strAuxUF & iif(strAuxUF <> "", " &nbsp;-&nbsp; " & UF_descricao(strAuxUF), "") & "</span></td>" & chr(13) & _
						"	</tr>" & chr(13)
					x = x & "<br>" & chr(13)
					x = x & cab
					
					vl_venda_UF = 0
					vl_NF_UF = 0
					percFatVenda_UF = 0
					vl_lista_UF = 0
					vl_desconto_UF = 0
					perc_desconto_UF = 0
					vl_RT_UF = 0
					percRTFatVenda_UF = 0
					percRTEDesc_UF = 0
					vl_RA_liquido_UF = 0
					lucro_liquido_UF = 0
					lucro_bruto_UF = 0
					marg_contrib_UF = 0
					marg_contrib_bruta_UF = 0
					vl_ent_UF = 0
                    intQtdeTotalIndicadorCidUf = 0
					end if
				
			'	FAZENDO A SUBTOTALIZAÇÃO POR UF
                intQtdeTotalIndicadorCidUf = intQtdeTotalIndicadorCidUf + intQtdeProdutos
				vl_venda_UF = vl_venda_UF + vl_venda
				vl_NF_UF = vl_NF_UF + vl_NF
				vl_lista_UF = vl_lista_UF +vl_lista
				vl_desconto_UF = vl_desconto_UF + vl_desconto
				vl_RT_UF = vl_RT_UF + vl_RT
				vl_RA_liquido_UF = vl_RA_liquido_UF + vl_RA_liquido
				lucro_liquido_UF = lucro_liquido_UF + lucro_liquido
				lucro_bruto_UF = lucro_bruto_UF + lucro_bruto
				vl_ent_UF = vl_ent_UF + vl_ent
				
			'	NUMERAÇÃO DA LINHA
				x = x & "	<tr nowrap>" & chr(13)
				x = x & "		<td valign='bottom' align='right' nowrap><span class='Rd' style='margin-right:2px;'>" & Cstr(n_reg_vetor) & ".</span></td>" & chr(13)

			'> UF
'				s = UCase(Trim("" & v(0)))
'				if s = "" then s = "&nbsp;"
'				s = s & " <> " & vRelat(intIdxVetor).CampoOrdenacao
			'> CIDADE
				s = Trim("" & v(1))
				if s = "" then s = "&nbsp;"
				x = x & "		<td class='MDTE' style='width:" & CStr(intLargCidade) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
            '> QUANTIDADE
				x = x & "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px;' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeProdutos) & "</span></td>" & chr(13)
        
        elseif rb_saida = COD_SAIDA_REL_ORIGEM_PEDIDO then
			'	CRIAR QUEBRA PARA A MUDANÇA DE PEDIDO ORIGEM
				v = Split(vRelat(intIdxVetor).c1, "|")
				strAuxCodigoPaiAnterior = strAuxCodigoPai
				strAuxCodigoAnterior = strAuxCodigo
				strAuxCodigoPai = UCase(Trim("" & v(0)))
				if strAuxCodigoPaiAnterior <> strAuxCodigoPai then
					if strAuxCodigoPaiAnterior <> "-XX-XX-" then
					'	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
						if vl_total_final_venda = 0 then
							percFatVenda_UF = 0
						else
							percFatVenda_UF = (vl_venda_UF/vl_total_final_venda)*100
							end if
					
					'	% DESCONTO
						if vl_lista_UF = 0 then
							perc_desconto_UF = 0
						else
							perc_desconto_UF = 100 * (vl_lista_UF-vl_venda_UF) / vl_lista_UF
							end if
					
					'	%COM SOBRE FATURAMENTO VENDA
						if vl_venda_UF = 0 then
							percRTFatVenda_UF = 0
						else
							percRTFatVenda_UF = (vl_RT_UF/vl_venda_UF)*100
							end if

					'	%COM + %DESCONTO
						percRTEDesc_UF = perc_desconto_UF + percRTFatVenda_UF
						
					'   %MARGEM CONTRIB
				        if vl_venda_UF = 0 then
				            marg_contrib_UF = 0
				        else
                            marg_contrib_UF = 100 * ((vl_venda_UF - vl_RT_UF - vl_ent_UF) / vl_venda_UF)
                        end if

					'	% MARG CONTRIB BRUTA
						if vl_NF_UF = 0 then
							marg_contrib_bruta_UF = 0
						else
							marg_contrib_bruta_UF = 100 * ((vl_NF_UF - vl_ent_UF) / vl_NF_UF)
							end if

						s_cor="black"
						if vl_total_venda < 0 then s_cor="red"
						x = x & "	</tr>" & chr(13)
						x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
								"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
								"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
								"TOTAL:</span></td>" & chr(13) & _
                                "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalIndicadorCidUf) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_UF) & "</span></td>" & chr(13)
								if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_UF) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_UF) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_UF) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_UF) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_UF) & "%" & "</span></td>" & chr(13)
								else
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_UF) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_UF) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_UF) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_UF) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_UF) & "%" & "</span></td>" & chr(13)
								end if
								x = x & _
								"	</tr>" & chr(13)
						
						x = x & "</table>" & chr(13)
						x = x & "<br>" & chr(13)
						end if
					
					x = x & "<table cellspacing=0>" & chr(13)
					x = x & _
						"	<tr>" & chr(13) & _
						"		<td style='background:white;' align='left'>&nbsp;</td>" & chr(13) & _
						"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;" & iif(strAuxCodigoPai <> "", obtem_descricao_tabela_t_codigo_descricao("PedidoECommerce_Origem_Grupo", strAuxCodigoPai), "") & "</span></td>" & chr(13) & _
						"	</tr>" & chr(13)
					x = x & "<br>" & chr(13)
					x = x & cab
					
					vl_venda_UF = 0
					vl_NF_UF = 0
					percFatVenda_UF = 0
					vl_lista_UF = 0
					vl_desconto_UF = 0
					perc_desconto_UF = 0
					vl_RT_UF = 0
					percRTFatVenda_UF = 0
					percRTEDesc_UF = 0
					vl_RA_liquido_UF = 0
					lucro_liquido_UF = 0
					lucro_bruto_UF = 0
					marg_contrib_UF = 0
					marg_contrib_bruta_UF = 0
					vl_ent_UF = 0
                    intQtdeTotalIndicadorCidUf = 0
					end if
				
			'	FAZENDO A SUBTOTALIZAÇÃO POR PEDIDO ORIGEM
                intQtdeTotalIndicadorCidUf = intQtdeTotalIndicadorCidUf + intQtdeProdutos
				vl_venda_UF = vl_venda_UF + vl_venda
				vl_NF_UF = vl_NF_UF + vl_NF
				vl_lista_UF = vl_lista_UF +vl_lista
				vl_desconto_UF = vl_desconto_UF + vl_desconto
				vl_RT_UF = vl_RT_UF + vl_RT
				vl_RA_liquido_UF = vl_RA_liquido_UF + vl_RA_liquido
				lucro_liquido_UF = lucro_liquido_UF + lucro_liquido
				lucro_bruto_UF = lucro_bruto_UF + lucro_bruto
				vl_ent_UF = vl_ent_UF + vl_ent
				
			'	NUMERAÇÃO DA LINHA
				x = x & "	<tr nowrap>" & chr(13)
				x = x & "		<td valign='bottom' align='right' nowrap><span class='Rd' style='margin-right:2px;'>" & Cstr(n_reg_vetor) & ".</span></td>" & chr(13)

			'> UF
'				s = UCase(Trim("" & v(0)))
'				if s = "" then s = "&nbsp;"
'				s = s & " <> " & vRelat(intIdxVetor).CampoOrdenacao
			'> CIDADE
				s = Trim("" & v(1))
				if s = "" then s = "&nbsp;"
				x = x & "		<td class='MDTE' style='width:" & CStr(intLargCidade) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
            '> QUANTIDADE
				x = x & "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px;' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeProdutos) & "</span></td>" & chr(13)

        elseif rb_saida = COD_SAIDA_REL_FABRICANTE then
			'	CRIAR QUEBRA PARA A MUDANÇA FABRICANTE
				v = Split(vRelat(intIdxVetor).c1, "|")
				strAuxFabricanteAnterior = strAuxFabricante
				strAuxNomeFabricanteAnterior = strAuxNomeFabricante
				strAuxNomeFabricante = UCase(Trim("" & v(0)))
				if strAuxNomeFabricanteAnterior <> strAuxNomeFabricante then
					if strAuxNomeFabricanteAnterior <> "-XX-XX-" then
					
				 '	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
						if vl_total_final_venda = 0 then
							percFatVenda_Fabricante = 0
						else
							percFatVenda_Fabricante = (vl_venda_Fabricante/vl_total_final_venda)*100
							end if
					
					'	% DESCONTO
						if vl_lista_Fabricante = 0 then
							perc_desconto_Fabricante = 0
						else
							perc_desconto_Fabricante = 100 * (vl_lista_Fabricante-vl_venda_Fabricante) / vl_lista_Fabricante
							end if
					
					'	%COM SOBRE FATURAMENTO VENDA
						if vl_venda_Fabricante = 0 then
							percRTFatVenda_Fabricante = 0
						else
							percRTFatVenda_Fabricante = (vl_RT_Fabricante/vl_venda_Fabricante)*100
							end if
							
					'	%COM + %DESCONTO
						percRTEDesc_Fabricante = perc_desconto_Fabricante + percRTFatVenda_Fabricante
						
					'   %MARGEM CONTRIBUIÇÃO
					    if vl_venda_Fabricante = 0 then
					        marg_contrib_Fabricante = 0
					    else
					        marg_contrib_Fabricante = 100 * ((vl_venda_Fabricante - vl_RT_Fabricante - vl_ent_Fabricante) / vl_venda_Fabricante)
					    end if

						if vl_NF_Fabricante = 0 then
							marg_contrib_bruta_Fabricante = 0
						else
							marg_contrib_bruta_Fabricante = 100 * ((vl_NF_Fabricante - vl_ent_Fabricante) / vl_NF_Fabricante)
							end if

					    md = ""
					    if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
						s_cor="black"
						if vl_total_venda < 0 then s_cor="red"
						x = x & "	</tr>" & chr(13)
						x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
								"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
								"		<td class='MTBE' align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
								"TOTAL:</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalFabricante) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_Fabricante) & "</span></td>" & chr(13) & _
								"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_Fabricante) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_Fabricante) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_Fabricante) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_Fabricante) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_Fabricante) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_Fabricante) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_Fabricante) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_Fabricante) & "</span></td>" & chr(13) 
								if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_Fabricante) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_Fabricante) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Fabricante) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Fabricante) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Fabricante) & "%" & "</span></td>" & chr(13)
								else
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_Fabricante) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_Fabricante) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Fabricante) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Fabricante) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Fabricante) & "%" & "</span></td>" & chr(13)
								end if
								x = x & _
								"	</tr>" & chr(13) 
						
						x = x & "</table>" & chr(13) 
						x = x & "<br>" & chr(13) 
						end if
					
					x = x & "<table cellspacing=0>" & chr(13)
					x = x & _
						"	<tr>" & chr(13) & _
						"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
						"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;" & strAuxNomeFabricante & iif(strAuxNomeFabricante <> "", " &nbsp;-&nbsp; " & Fabricante_descricao(strAuxNomeFabricante), "") & "</span></td>" & chr(13) & _
						"	</tr>" & chr(13)
					x = x & "<br>" & chr(13)
					x = x & cab
					
					vl_venda_Fabricante = 0
					vl_NF_Fabricante = 0
					percFatVenda_Fabricante = 0
					vl_lista_Fabricante = 0
					vl_desconto_Fabricante = 0
					perc_desconto_Fabricante = 0
					vl_RT_Fabricante = 0
					percRTFatVenda_Fabricante = 0
					percRTEDesc_Fabricante = 0
					vl_RA_liquido_Fabricante = 0
					lucro_liquido_Fabricante = 0
					lucro_bruto_Fabricante = 0
					marg_contrib_Fabricante = 0
					marg_contrib_bruta_Fabricante = 0
					vl_ent_Fabricante = 0
					intQtdeTotalFabricante=0
					end if
				
			'	FAZENDO A SUBTOTALIZAÇÃO POR FABRICANTE
			    intQtdeTotalFabricante= intQtdeTotalFabricante + intQtdeProdutos
				vl_venda_Fabricante = vl_venda_Fabricante + vl_venda
				vl_NF_Fabricante = vl_NF_Fabricante + vl_NF
				vl_lista_Fabricante = vl_lista_Fabricante +vl_lista
				vl_desconto_Fabricante = vl_desconto_Fabricante + vl_desconto
				vl_RT_Fabricante = vl_RT_Fabricante + vl_RT
				vl_RA_liquido_Fabricante = vl_RA_liquido_Fabricante + vl_RA_liquido
				lucro_liquido_Fabricante = lucro_liquido_Fabricante + lucro_liquido
				lucro_bruto_Fabricante = lucro_bruto_Fabricante + lucro_bruto
				vl_ent_Fabricante = vl_ent_Fabricante + vl_ent
				
			'	NUMERAÇÃO DA LINHA
				x = x & "	<tr nowrap>" & chr(13)
				x = x & "		<td valign='bottom' align='right' nowrap><span class='Rd' style='margin-right:2px;'>" & Cstr(n_reg_vetor) & ".</span></td>" & chr(13)
            
              '> NOME FABRICANTE
'				s = Trim("" & v(2))
'				if s = "" then s = "&nbsp;"
'				s = s & " <> " & vRelat(intIdxVetor).CampoOrdenacao
                
              '> CODIGO FABRICANTE E PRODUTO
				 s = Trim("" & v(2))
				if s = "" then s = "&nbsp;"
				x = x & "		<td class='MDTE' style='width:" & CStr(intLargCodFabricante) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
              
               '> Descricao
				s = Trim("" & v(3))
				if s <> "" then s = produto_formata_descricao_em_html(s) else s = "&nbsp;"
				x = x & "		<td class='MTD' style='width:" & CStr(intLargDescrProduto) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
                '> QUANTIDADE
				x = x & "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px;' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeProdutos) & "</span></td>" & chr(13)	     
            
           elseif rb_saida = COD_SAIDA_REL_GRUPO_PRODUTO then
			'	CRIAR QUEBRA PARA A MUDANÇA FABRICANTE
				v = Split(vRelat(intIdxVetor).c1, "|")
				strAuxFabricanteAnterior = strAuxFabricante
				strAuxNomeFabricanteAnterior = strAuxNomeFabricante
				strAuxNomeFabricante = UCase(Trim("" & v(0)))
				if strAuxNomeFabricanteAnterior <> strAuxNomeFabricante then
					if strAuxNomeFabricanteAnterior <> "-XX-XX-" then
					
					    md = ""
					    if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
						s_cor="black"
						if vl_total_venda < 0 then s_cor="red"
                    
                    ' PERCENTUAIS RELATIVOS AOS GRUPOS DE PRODUTO
                    '	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
						if vl_total_final_venda = 0 then
							percFatVenda_ProdutoGrupo = 0
						else
							percFatVenda_ProdutoGrupo = (vl_venda_ProdutoGrupo/vl_total_final_venda)*100
							end if
					
					'	% DESCONTO
						if vl_lista_ProdutoGrupo = 0 then
							perc_desconto_ProdutoGrupo = 0
						else
							perc_desconto_ProdutoGrupo = 100 * (vl_lista_ProdutoGrupo-vl_venda_ProdutoGrupo) / vl_lista_ProdutoGrupo
							end if
					
					'	%COM SOBRE FATURAMENTO VENDA
						if vl_venda_ProdutoGrupo = 0 then
							percRTFatVenda_ProdutoGrupo = 0
						else
							percRTFatVenda_ProdutoGrupo = (vl_RT_ProdutoGrupo/vl_venda_ProdutoGrupo)*100
							end if
							
					'	%COM + %DESCONTO
						percRTEDesc_ProdutoGrupo = perc_desconto_ProdutoGrupo + percRTFatVenda_ProdutoGrupo
						
					'   %MARGEM CONTRIBUIÇÃO
					    if vl_venda_ProdutoGrupo = 0 then
					        marg_contrib_ProdutoGrupo = 0
					    else
					        marg_contrib_ProdutoGrupo = 100 * ((vl_venda_ProdutoGrupo - vl_RT_ProdutoGrupo - vl_ent_ProdutoGrupo) / vl_venda_ProdutoGrupo)
					    end if

						if vl_NF_ProdutoGrupo = 0 then
							marg_contrib_bruta_ProdutoGrupo = 0
						else
							marg_contrib_bruta_ProdutoGrupo = 100 * ((vl_NF_ProdutoGrupo - vl_ent_ProdutoGrupo) / vl_NF_ProdutoGrupo)
							end if

                        x = x & "	</tr>" & chr(13)
						x = x & "	<tr nowrap>" & chr(13) & _
								"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
								"		<td class='ME MC' align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
								"Total do Grupo:</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalProdutoGrupo) & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_ProdutoGrupo) & "</span></td>" & chr(13) & _
								"		<td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_ProdutoGrupo) & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_ProdutoGrupo) & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_ProdutoGrupo) & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MC " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_ProdutoGrupo) & "</span></td>" & chr(13) 
								if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
								    x = x & _
								        "		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_ProdutoGrupo) & "</span></td>" & chr(13) & _
								        "       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_ProdutoGrupo) & "</span></td>" & chr(13) & _
										"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_ProdutoGrupo) & "</span></td>" & chr(13) & _
										"       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_ProdutoGrupo) & "%" & "</span></td>" & chr(13)
								else
								    x = x & _
								        "		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_ProdutoGrupo) & "</span></td>" & chr(13) & _
								        "       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_ProdutoGrupo) & "</span></td>" & chr(13) & _
										"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_ProdutoGrupo) & "</span></td>" & chr(13) & _
										"       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_ProdutoGrupo) & "%" & "</span></td>" & chr(13)
								end if         
    
                    '	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
						if vl_total_final_venda = 0 then
							percFatVenda_Fabricante = 0
						else
							percFatVenda_Fabricante = (vl_venda_Fabricante/vl_total_final_venda)*100
							end if
					
					'	% DESCONTO
						if vl_lista_Fabricante = 0 then
							perc_desconto_Fabricante = 0
						else
							perc_desconto_Fabricante = 100 * (vl_lista_Fabricante-vl_venda_Fabricante) / vl_lista_Fabricante
							end if
					
					'	%COM SOBRE FATURAMENTO VENDA
						if vl_venda_Fabricante = 0 then
							percRTFatVenda_Fabricante = 0
						else
							percRTFatVenda_Fabricante = (vl_RT_Fabricante/vl_venda_Fabricante)*100
							end if
							
					'	%COM + %DESCONTO
						percRTEDesc_Fabricante = perc_desconto_Fabricante + percRTFatVenda_Fabricante
						
					'   %MARGEM CONTRIBUIÇÃO
					    if vl_venda_Fabricante = 0 then
					        marg_contrib_Fabricante = 0
					    else
					        marg_contrib_Fabricante = 100 * ((vl_venda_Fabricante - vl_RT_Fabricante - vl_ent_Fabricante) / vl_venda_Fabricante)
					    end if

						if vl_NF_Fabricante = 0 then
							marg_contrib_bruta_Fabricante = 0
						else
							marg_contrib_bruta_Fabricante = 100 * ((vl_NF_Fabricante - vl_ent_Fabricante) / vl_NF_Fabricante)
							end if
                                  

						x = x & "	</tr>" & chr(13)
						x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
								"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
								"		<td class='MTBE' align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
								"TOTAL:</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalFabricante) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_Fabricante) & "</span></td>" & chr(13) & _
								"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_Fabricante) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_Fabricante) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_Fabricante) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_Fabricante) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_Fabricante) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_Fabricante) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_Fabricante) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_Fabricante) & "</span></td>" & chr(13) 
								if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_Fabricante) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_Fabricante) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Fabricante) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Fabricante) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Fabricante) & "%" & "</span></td>" & chr(13)
								else
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_Fabricante) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_Fabricante) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Fabricante) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Fabricante) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Fabricante) & "%" & "</span></td>" & chr(13)
								end if
								x = x & _
								"	</tr>" & chr(13) 
						
						x = x & "</table>" & chr(13) 
						x = x & "<br>" & chr(13) 
						end if
					
					x = x & "<table cellspacing=0>" & chr(13)
					x = x & _
						"	<tr>" & chr(13) & _
						"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
						"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;" & strAuxNomeFabricante & iif(strAuxNomeFabricante <> "", " &nbsp;-&nbsp; " & Fabricante_descricao(strAuxNomeFabricante), "") & "</span></td>" & chr(13) & _
						"	</tr>" & chr(13)
					x = x & "<br>" & chr(13)
					x = x & cab
        
                    x = x & _
					"	<tr>" & chr(13) & _
					"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:#EEE;'><span class='N'>&nbsp;" & v(2) & " - " & v(3) & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
					x = x & "<br>" & chr(13)
                    
                    vl_venda_ProdutoGrupo = 0
					vl_NF_ProdutoGrupo = 0
					percFatVenda_ProdutoGrupo = 0
					vl_lista_ProdutoGrupo = 0
					vl_desconto_ProdutoGrupo = 0
					perc_desconto_ProdutoGrupo = 0
					vl_RT_ProdutoGrupo = 0
					percRTFatVenda_ProdutoGrupo = 0
					percRTEDesc_ProdutoGrupo = 0
					vl_RA_liquido_ProdutoGrupo = 0
					lucro_liquido_ProdutoGrupo = 0
					lucro_bruto_ProdutoGrupo = 0
					marg_contrib_ProdutoGrupo = 0
					marg_contrib_bruta_ProdutoGrupo = 0
					vl_ent_ProdutoGrupo = 0
					intQtdeTotalProdutoGrupo=0
					
					vl_venda_Fabricante = 0
					vl_NF_Fabricante = 0
					percFatVenda_Fabricante = 0
					vl_lista_Fabricante = 0
					vl_desconto_Fabricante = 0
					perc_desconto_Fabricante = 0
					vl_RT_Fabricante = 0
					percRTFatVenda_Fabricante = 0
					percRTEDesc_Fabricante = 0
					vl_RA_liquido_Fabricante = 0
					lucro_liquido_Fabricante = 0
					lucro_bruto_Fabricante = 0
					marg_contrib_Fabricante = 0
					marg_contrib_bruta_Fabricante = 0
					vl_ent_Fabricante = 0
					intQtdeTotalFabricante=0
                end if
				
		'	FAZENDO A SUBTOTALIZAÇÃO POR FABRICANTE
			intQtdeTotalFabricante= intQtdeTotalFabricante + intQtdeProdutos
			vl_venda_Fabricante = vl_venda_Fabricante + vl_venda
			vl_NF_Fabricante = vl_NF_Fabricante + vl_NF
			vl_lista_Fabricante = vl_lista_Fabricante +vl_lista
			vl_desconto_Fabricante = vl_desconto_Fabricante + vl_desconto
			vl_RT_Fabricante = vl_RT_Fabricante + vl_RT
			vl_RA_liquido_Fabricante = vl_RA_liquido_Fabricante + vl_RA_liquido
			lucro_liquido_Fabricante = lucro_liquido_Fabricante + lucro_liquido
			lucro_bruto_Fabricante = lucro_bruto_Fabricante + lucro_bruto
			vl_ent_Fabricante = vl_ent_Fabricante + vl_ent               

            ' FAZ A QUEBRA DO GRUPO DE PRODUTOS
            v = Split(vRelat(intIdxVetor).c1, "|")
			strAuxProdutoGrupoAnterior = strAuxProdutoGrupo
			strAuxNomeProdutoGrupoAnterior = strAuxNomeProdutoGrupo
			strAuxNomeProdutoGrupo = UCase(Trim("" & v(2)))
            if strAuxNomeProdutoGrupoAnterior <> strAuxNomeProdutoGrupo then
				if strAuxNomeProdutoGrupoAnterior <> "-XX-XX-" then
					
				'	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
					if vl_total_final_venda = 0 then
						percFatVenda_ProdutoGrupo = 0
					else
						percFatVenda_ProdutoGrupo = (vl_venda_ProdutoGrupo/vl_total_final_venda)*100
						end if
					
				'	% DESCONTO
					if vl_lista_ProdutoGrupo = 0 then
						perc_desconto_ProdutoGrupo = 0
					else
						perc_desconto_ProdutoGrupo = 100 * (vl_lista_ProdutoGrupo-vl_venda_ProdutoGrupo) / vl_lista_ProdutoGrupo
						end if
					
				'	%COM SOBRE FATURAMENTO VENDA
					if vl_venda_ProdutoGrupo = 0 then
						percRTFatVenda_ProdutoGrupo = 0
					else
						percRTFatVenda_ProdutoGrupo = (vl_RT_ProdutoGrupo/vl_venda_ProdutoGrupo)*100
						end if
							
				'	%COM + %DESCONTO
					percRTEDesc_ProdutoGrupo = perc_desconto_ProdutoGrupo + percRTFatVenda_ProdutoGrupo
						
				'   %MARGEM CONTRIBUIÇÃO
					if vl_venda_ProdutoGrupo = 0 then
					    marg_contrib_ProdutoGrupo = 0
					else
					    marg_contrib_ProdutoGrupo = 100 * ((vl_venda_ProdutoGrupo - vl_RT_ProdutoGrupo - vl_ent_ProdutoGrupo) / vl_venda_ProdutoGrupo)
					end if

					if vl_NF_ProdutoGrupo = 0 then
						marg_contrib_bruta_ProdutoGrupo = 0
					else
						marg_contrib_bruta_ProdutoGrupo = 100 * ((vl_NF_ProdutoGrupo - vl_ent_ProdutoGrupo) / vl_NF_ProdutoGrupo)
						end if

					md = ""
					if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
					s_cor="black"
					if vl_total_venda < 0 then s_cor="red"
                    
                    if strAuxNomeFabricanteAnterior = strAuxNomeFabricante then

                        x = x & "	</tr>" & chr(13)
						x = x & "	<tr nowrap>" & chr(13) & _
								"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
								"		<td class='MC ME' align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
								"Total do Grupo:</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalProdutoGrupo) & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_ProdutoGrupo) & "</span></td>" & chr(13) & _
								"		<td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_ProdutoGrupo) & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_ProdutoGrupo) & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_ProdutoGrupo) & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MC " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_ProdutoGrupo) & "</span></td>" & chr(13) 
								if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
								    x = x & _
								        "		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_ProdutoGrupo) & "</span></td>" & chr(13) & _
								        "       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_ProdutoGrupo) & "</span></td>" & chr(13) & _
										"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_ProdutoGrupo) & "</span></td>" & chr(13) & _
										"       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_ProdutoGrupo) & "%" & "</span></td>" & chr(13)
								else
								    x = x & _
								        "		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_ProdutoGrupo) & "</span></td>" & chr(13) & _
								        "       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_ProdutoGrupo) & "</span></td>" & chr(13) & _
										"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_ProdutoGrupo) & "</span></td>" & chr(13) & _
										"       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_ProdutoGrupo) & "%" & "</span></td>" & chr(13)
								end if
								x = x & _
								"	</tr>" & chr(13)

                        x = x & _
						"	<tr>" & chr(13) & _
						"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
						"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:#EEE;'><span class='N'>&nbsp;" & v(2) & " - " & v(3) & "</span></td>" & chr(13) & _
						"	</tr>" & chr(13)
					    x = x & "<br>" & chr(13)

                        vl_venda_ProdutoGrupo = 0
						vl_NF_ProdutoGrupo = 0
					    percFatVenda_ProdutoGrupo = 0
					    vl_lista_ProdutoGrupo = 0
					    vl_desconto_ProdutoGrupo = 0
					    perc_desconto_ProdutoGrupo = 0
					    vl_RT_ProdutoGrupo = 0
					    percRTFatVenda_ProdutoGrupo = 0
					    percRTEDesc_ProdutoGrupo = 0
					    vl_RA_liquido_ProdutoGrupo = 0
					    lucro_liquido_ProdutoGrupo = 0
						lucro_bruto_ProdutoGrupo = 0
					    marg_contrib_ProdutoGrupo = 0
						marg_contrib_bruta_ProdutoGrupo = 0
					    vl_ent_ProdutoGrupo = 0
					    intQtdeTotalProdutoGrupo=0
                    end if
                end if
			end if
				
			'	FAZENDO A SUBTOTALIZAÇÃO POR GRUPO DE PRODUTOS
			    intQtdeTotalProdutoGrupo= intQtdeTotalProdutoGrupo + intQtdeProdutos
				vl_venda_ProdutoGrupo = vl_venda_ProdutoGrupo + vl_venda
				vl_NF_ProdutoGrupo = vl_NF_ProdutoGrupo + vl_NF
				vl_lista_ProdutoGrupo = vl_lista_ProdutoGrupo +vl_lista
				vl_desconto_ProdutoGrupo = vl_desconto_ProdutoGrupo + vl_desconto
				vl_RT_ProdutoGrupo = vl_RT_ProdutoGrupo + vl_RT
				vl_RA_liquido_ProdutoGrupo = vl_RA_liquido_ProdutoGrupo + vl_RA_liquido
				lucro_liquido_ProdutoGrupo = lucro_liquido_ProdutoGrupo + lucro_liquido
				lucro_bruto_ProdutoGrupo = lucro_bruto_ProdutoGrupo + lucro_bruto
				vl_ent_ProdutoGrupo = vl_ent_ProdutoGrupo + vl_ent 

			'	NUMERAÇÃO DA LINHA
				x = x & "	<tr nowrap>" & chr(13)
				x = x & "		<td valign='bottom' align='right' nowrap><span class='Rd' style='margin-right:2px;'>" & Cstr(n_reg_vetor) & ".</span></td>" & chr(13)
            
              '> NOME FABRICANTE
'				s = Trim("" & v(2))
'				if s = "" then s = "&nbsp;"
'				s = s & " <> " & vRelat(intIdxVetor).CampoOrdenacao
                
              '> CODIGO FABRICANTE E PRODUTO
				 s = Trim("" & v(4))
				if s = "" then s = "&nbsp;"
				x = x & "		<td class='MDTE' style='width:" & CStr(intLargCodFabricante) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
              
               '> Descricao
				s = Trim("" & v(5))
				if s <> "" then s = produto_formata_descricao_em_html(s) else s = "&nbsp;"
				x = x & "		<td class='MTD' style='width:" & CStr(intLargDescrProduto) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
                '> QUANTIDADE
				x = x & "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px;' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeProdutos) & "</span></td>" & chr(13)	  
    

            elseif rb_saida = COD_SAIDA_REL_EMPRESA then
			'	CRIAR QUEBRA PARA A MUDANÇA DE EMPRESA
				v = Split(vRelat(intIdxVetor).c1, "|")
				strAuxEmpresaAnterior = strAuxEmpresa
				strAuxNomeEmpresaAnterior = strAuxNomeEmpresa
				strAuxNomeEmpresa = UCase(Trim("" & v(0)))
				if strAuxNomeEmpresaAnterior <> strAuxNomeEmpresa then
					if strAuxNomeEmpresaAnterior <> "-XX-XX-" then
					
				 '	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
						if vl_total_final_venda = 0 then
							percFatVenda_Empresa = 0
						else
							percFatVenda_Empresa = (vl_venda_Empresa/vl_total_final_venda)*100
							end if
					
					'	% DESCONTO
						if vl_lista_Empresa = 0 then
							perc_desconto_Empresa = 0
						else
							perc_desconto_Empresa = 100 * (vl_lista_Empresa-vl_venda_Empresa) / vl_lista_Empresa
							end if
					
					'	%COM SOBRE FATURAMENTO VENDA
						if vl_venda_Empresa = 0 then
							percRTFatVenda_Empresa = 0
						else
							percRTFatVenda_Empresa = (vl_RT_Empresa/vl_venda_Empresa)*100
							end if
							
					'	%COM + %DESCONTO
						percRTEDesc_Empresa = perc_desconto_Empresa + percRTFatVenda_Empresa
						
					'   %MARGEM CONTRIBUIÇÃO
					    if vl_venda_Empresa = 0 then
					        marg_contrib_Empresa = 0
					    else
					        marg_contrib_Empresa = 100 * ((vl_venda_Empresa - vl_RT_Empresa - vl_ent_Empresa) / vl_venda_Empresa)
					    end if

						if vl_NF_Empresa = 0 then
							marg_contrib_bruta_Empresa = 0
						else
							marg_contrib_bruta_Empresa = 100 * ((vl_NF_Empresa - vl_ent_Empresa) / vl_NF_Empresa)
							end if

                        md = ""
					    if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
						s_cor="black"
						if vl_total_venda < 0 then s_cor="red"
						x = x & "	</tr>" & chr(13)
						x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
								"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
								"		<td class='MTBE' align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
								"TOTAL:</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalEmpresa) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_Empresa) & "</span></td>" & chr(13) & _
								"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_Empresa) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_Empresa) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_Empresa) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_Empresa) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_Empresa) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_Empresa) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_Empresa) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_Empresa) & "</span></td>" & chr(13) 
								if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_Empresa) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_Empresa) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Empresa) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Empresa) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Empresa) & "%" & "</span></td>" & chr(13)
								else
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_Empresa) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_Empresa) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Empresa) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Empresa) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Empresa) & "%" & "</span></td>" & chr(13)
								end if
								x = x & _
								"	</tr>" & chr(13) 
						
						x = x & "</table>" & chr(13) 
						x = x & "<br>" & chr(13) 
						end if
					
					x = x & "<table cellspacing=0>" & chr(13)
					x = x & _
						"	<tr>" & chr(13) & _
						"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
						"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;" & strAuxNomeEmpresa & iif(strAuxNomeEmpresa <> "", " &nbsp;-&nbsp; " & obtem_apelido_empresa_NFe_emitente(strAuxNomeEmpresa), "") & "</span></td>" & chr(13) & _
						"	</tr>" & chr(13)
					x = x & "<br>" & chr(13)
					x = x & cab
					
					vl_venda_Empresa = 0
					vl_NF_Empresa = 0
					percFatVenda_Empresa = 0
					vl_lista_Empresa = 0
					vl_desconto_Empresa = 0
					perc_desconto_Empresa = 0
					vl_RT_Empresa = 0
					percRTFatVenda_Empresa = 0
					percRTEDesc_Empresa = 0
					vl_RA_liquido_Empresa = 0
					lucro_liquido_Empresa = 0
					lucro_bruto_Empresa = 0
					marg_contrib_Empresa = 0
					marg_contrib_bruta_Empresa = 0
					vl_ent_Empresa = 0
					intQtdeTotalEmpresa = 0
					end if
				
			'	FAZENDO A SUBTOTALIZAÇÃO POR EMPRESA
			    intQtdeTotalEmpresa = intQtdeTotalEmpresa + intQtdeProdutos
				vl_venda_Empresa = vl_venda_Empresa + vl_venda
				vl_NF_Empresa = vl_NF_Empresa + vl_NF
				vl_lista_Empresa = vl_lista_Empresa +vl_lista
				vl_desconto_Empresa = vl_desconto_Empresa + vl_desconto
				vl_RT_Empresa = vl_RT_Empresa + vl_RT
				vl_RA_liquido_Empresa = vl_RA_liquido_Empresa + vl_RA_liquido
				lucro_liquido_Empresa = lucro_liquido_Empresa + lucro_liquido
				lucro_bruto_Empresa = lucro_bruto_Empresa + lucro_bruto
				vl_ent_Empresa = vl_ent_Empresa + vl_ent
				
				
			'	NUMERAÇÃO DA LINHA
				x = x & "	<tr nowrap>" & chr(13)
				x = x & "		<td valign='bottom' align='right' nowrap><span class='Rd' style='margin-right:2px;'>" & Cstr(n_reg_vetor) & ".</span></td>" & chr(13)
            
              '> NOME EMPRESA
'				s = Trim("" & v(0))
'				if s = "" then s = "&nbsp;"
'				s = s & " <> " & vRelat(intIdxVetor).CampoOrdenacao
                
              '> CODIGO FABRICANTE E PRODUTO
				 s = Trim("" & v(3))
				if s = "" then s = "&nbsp;"
				x = x & "		<td class='MDTE' style='width:" & CStr(intLargCodFabricante) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
              
               '> Descricao
				s = Trim("" & v(4))
				if s <> "" then s = produto_formata_descricao_em_html(s) else s = "&nbsp;"
				x = x & "		<td class='MTD' style='width:" & CStr(intLargDescrProduto) & "px;' align='left' valign='bottom'><span class='Cn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)
                '> QUANTIDADE
				x = x & "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px;' align='right' valign='bottom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeProdutos) & "</span></td>" & chr(13)
	      
	      
			else ' if (rb_saida = ...)
				s = ""
				end if

		 '> VALOR DE VENDA
			x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda) & "</span></td>" & chr(13)

		 '>	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
			x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda) & "%" & "</span></td>" & chr(13)
			
		 '> VALOR DE LISTA
			x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista) & "</span></td>" & chr(13)

		 '> DESCONTO (R$)
			x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto) & "</span></td>" & chr(13)

		 '> % DESCONTO
			x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto) & "%" & "</span></td>" & chr(13)

		 '> COMISSÃO (R$)
			x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT) & "</span></td>" & chr(13)

		 '> % COMISSÃO SOBRE FATURAMENTO VENDA
			x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda) & "%" & "</span></td>" & chr(13)

		 '> %COMISSÃO + %DESCONTO
			x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc) & "%" & "</span></td>" & chr(13)

		 '> RA LÍQUIDO (R$)
			x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido) & "</span></td>" & chr(13)
			
			if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
			    '> LUCRO LIQUIDO (R$)
			        x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido) & "</span></td>" & chr(13)
			
			    '> Margem Contrib (%)
			        x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib) & "%" & "</span></td>" & chr(13)
			
				 '> VALOR NF
					x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF) & "</span></td>" & chr(13)

			    '> LUCRO BRUTO (R$)
			        x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto) & "</span></td>" & chr(13)

			    '> Margem Contrib Bruta (%)
			        x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta) & "%" & "</span></td>" & chr(13)
			else
			    '> LUCRO LIQUIDO (R$)
			        x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido) & "</span></td>" & chr(13)
			
			    '> Margem Contrib (%)
			        x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib) & "%" & "</span></td>" & chr(13)

				 '> VALOR NF
					x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF) & "</span></td>" & chr(13)

			    '> LUCRO BRUTO (R$)
			        x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto) & "</span></td>" & chr(13)

			    '> Margem Contrib Bruta (%)
			        x = x & "		<td align='right' valign='bottom' class='MTD'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta) & "%" & "</span></td>" & chr(13)
			end if
			
			x = x & "	<input type='hidden' value='" & formata_moeda(vl_ent) & "' />" & chr(13)

			x = x & "	</tr>" & chr(13)
				
			if (n_reg_vetor mod 100) = 0 then
				Response.Write x
				x = ""
				end if
				
			next


	'	ÚLTIMO SUBTOTAL NO RELATÓRIO INDICADOR/UF E CIDADE/UF
		if (rb_saida = COD_SAIDA_REL_INDICADOR_UF) or (rb_saida = COD_SAIDA_REL_CIDADE_UF) then


		'	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
			    if vl_total_final_venda = 0 then
				percFatVenda_UF = 0
			    else
				percFatVenda_UF = (vl_venda_UF/vl_total_final_venda)*100
				end if
		  
		
		'	% DESCONTO
			if vl_lista_UF = 0 then
				perc_desconto_UF = 0
			else
				perc_desconto_UF = 100 * (vl_lista_UF-vl_venda_UF) / vl_lista_UF
				end if
				
		
		
		'	%COM SOBRE FATURAMENTO VENDA
			if vl_venda_UF = 0 then
				percRTFatVenda_UF = 0
			else
				percRTFatVenda_UF = (vl_RT_UF/vl_venda_UF)*100
				end if
		'  QTDE
        if rb_saida = COD_SAIDA_REL_INDICADOR_UF then
           intQtdeTotal = intQtdeTotalIndicadorUf       
        else
            intQtdeTotal = intQtdeTotalIndicadorCidUf
        end if
		'	%COM + %DESCONTO
			percRTEDesc_UF = perc_desconto_UF + percRTFatVenda_UF
			
			
		'   %MARGEM CONTRIB
		    if vl_venda_UF = 0 then
		        marg_contrib_UF = 0
		    else
		        marg_contrib_UF = 100 * ((vl_venda_UF - vl_RT_UF - vl_ent_UF) / vl_venda_UF)
		    end if

			if vl_NF_UF = 0 then
				marg_contrib_bruta_UF = 0
			else
				marg_contrib_bruta_UF = 100 * ((vl_NF_UF - vl_ent_UF) / vl_NF_UF)
				end if

		    md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			s_cor="black"
			if vl_total_venda < 0 then s_cor="red"
			x = x & "	</tr>" & chr(13)
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
					"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL:</span></td>" & chr(13) & _
                    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotal) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_UF) & "</span></td>" & chr(13) & _
					"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_UF) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_UF) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_UF) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_UF) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_UF) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_UF) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_UF) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_UF) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_UF) & "</span></td>" & chr(13) & _
					        "       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_UF) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_UF) & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_UF) & "</span></td>" & chr(13) & _
							"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_UF) & "%" & "</span></td>" & chr(13)
					else
					    x = x & _
					        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_UF) & "</span></td>" & chr(13) & _
					        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_UF) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_UF) & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_UF) & "</span></td>" & chr(13) & _
							"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_UF) & "%" & "</span></td>" & chr(13)
					end if
					x = x & _
					"	</tr>" & chr(13)
		 
	
			
		'	CRIANDO ESPAÇOS EM BRANCO PARA ALINHAR A TABELA COM O TOTAL DE TODOS OS ESTADOS

			
			strEspacos = ""
			for iEspacos = 0 to Len(Cstr(n_reg_vetor)) + 1
				strEspacos = strEspacos & "&nbsp;"
				next
			
			x = x & "</table>" & chr(13)
			x = x & "<br>" & chr(13)
			x = x & "<table cellspacing=0>" & chr(13)
			x = x & _
					"	<tr>" & chr(13) & _
					"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;TOTAL</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
			x = x & "<br>" & chr(13)
			x = x & "	<tr nowrap style='background:azure'>" & chr(13) & _
					"		<td align='left' valign='bottom' style='background:white;' nowrap>" & strEspacos & "</td>" & chr(13) & _
					"		<td class='MDTE' style='width:" & Cstr(intLargIndicador) & "px' align='left' valign='bottom' nowrap><span class='R'>&nbsp;</span></td>" & chr(13) & _
                    "		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px' align='right' valign='bottom' nowrap><span class='Rd' style='font-weight:bold;'>Qtde</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Venda (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Fat Venda Total</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Lista (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Desc (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Desc</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>COM (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% COM</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px;' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>%COM<br>+<br>%Desc</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>RA Líquido (R$)</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto -RA -COM (R$)</span></td>" & chr(13) & _
					        "       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib (%)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (R$)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto (R$)</span></td>" & chr(13) & _
							"       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta (%)</span></td>" & chr(13)
					else
					    x = x & _
					        "		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado -RA -COM (R$)</span></td>" & chr(13) & _
					        "       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib Estimada (%)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (R$)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado (R$)</span></td>" & chr(13) & _
							"       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta Estimada (%)</span></td>" & chr(13)
					end if
					x = x & _
					"	</tr>" & chr(13)
			end if
	
     '	ÚLTIMO SUBTOTAL NO RELATÓRIO PEDIDO ORIGEM
		     if (rb_saida = COD_SAIDA_REL_ORIGEM_PEDIDO)  then


					'	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
						if vl_total_final_venda = 0 then
							percFatVenda_UF = 0
						else
							percFatVenda_UF = (vl_venda_UF/vl_total_final_venda)*100
							end if
					
					'	% DESCONTO
						if vl_lista_UF = 0 then
							perc_desconto_UF = 0
						else
							perc_desconto_UF = 100 * (vl_lista_UF-vl_venda_UF) / vl_lista_UF
							end if
					
					'	%COM SOBRE FATURAMENTO VENDA
						if vl_venda_UF = 0 then
							percRTFatVenda_UF = 0
						else
							percRTFatVenda_UF = (vl_RT_UF/vl_venda_UF)*100
							end if

					'	%COM + %DESCONTO
						percRTEDesc_UF = perc_desconto_UF + percRTFatVenda_UF
						
					'   %MARGEM CONTRIB
				        if vl_venda_UF = 0 then
				            marg_contrib_UF = 0
				        else
                            marg_contrib_UF = 100 * ((vl_venda_UF - vl_RT_UF - vl_ent_UF) / vl_venda_UF)
                        end if

						if vl_NF_UF = 0 then
							marg_contrib_bruta_UF = 0
						else
							marg_contrib_bruta_UF = 100 * ((vl_NF_UF - vl_ent_UF) / vl_NF_UF)
							end if
		    
			s_cor="black"
			if vl_total_venda < 0 then s_cor="red"
			x = x & "	</tr>" & chr(13)
						x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
								"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
								"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
								"TOTAL:</span></td>" & chr(13) & _
                                "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalIndicadorCidUf) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_UF) & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_UF) & "%" & "</span></td>" & chr(13) & _
								"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_UF) & "</span></td>" & chr(13)
								if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_UF) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_UF) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_UF) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_UF) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_UF) & "%" & "</span></td>" & chr(13)
								else
								    x = x & _
								        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_UF) & "</span></td>" & chr(13) & _
								        "       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_UF) & "%" & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_UF) & "</span></td>" & chr(13) & _
										"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_UF) & "</span></td>" & chr(13) & _
										"       <td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_UF) & "%" & "</span></td>" & chr(13)
								end if
								x = x & _
								"	</tr>" & chr(13)
			'	CRIANDO ESPAÇOS EM BRANCO PARA ALINHAR A TABELA COM O TOTAL DE TODOS OS GRUPOS DE ORIGEM DE PEDIDO

			
			strEspacos = ""
			for iEspacos = 0 to Len(Cstr(n_reg_vetor)) + 1
				strEspacos = strEspacos & "&nbsp;"
				next
			
			x = x & "</table>" & chr(13)
			x = x & "<br>" & chr(13)
			x = x & "<table cellspacing=0>" & chr(13)
			x = x & _
					"	<tr>" & chr(13) & _
					"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;TOTAL</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
			x = x & "<br>" & chr(13)
			x = x & "	<tr nowrap style='background:azure'>" & chr(13) & _
					"		<td align='left' valign='bottom' style='background:white;' nowrap>" & strEspacos & "</td>" & chr(13) & _
					"		<td class='MDTE' style='width:" & Cstr(intLargIndicador) & "px' align='left' valign='bottom' nowrap><span class='R'>&nbsp;</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px' align='right' valign='bottom' nowrap><span class='Rd' style='font-weight:bold;'>Qtde</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Venda (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Fat Venda Total</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Lista (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Desc (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Desc</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>COM (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% COM</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px;' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>%COM<br>+<br>%Desc</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>RA Líquido (R$)</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto -RA -COM (R$)</span></td>" & chr(13) & _
					        "       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib (%)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (R$)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto (R$)</span></td>" & chr(13) & _
							"       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta (%)</span></td>" & chr(13)
					else
					    x = x & _
					        "		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado -RA -COM (R$)</span></td>" & chr(13) & _
					        "       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib Estimada (%)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (R$)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado (R$)</span></td>" & chr(13) & _
							"       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta Estimada (%)</span></td>" & chr(13)
					end if
					x = x & _
					"	</tr>" & chr(13)
			end if
	 
	 '	ÚLTIMO SUBTOTAL NO RELATÓRIO FABRICANTE
		if (rb_saida = COD_SAIDA_REL_FABRICANTE)  then


		'	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
			    if vl_total_final_venda = 0 then
				percFatVenda_FABRICANTE = 0
			    else
				percFatVenda_FABRICANTE = (vl_venda_FABRICANTE/vl_total_final_venda)*100
				end if
		  
		
		'	% DESCONTO
			if vl_lista_FABRICANTE = 0 then
				perc_desconto_UF = 0
			else
				perc_desconto_FABRICANTE = 100 * (vl_lista_FABRICANTE-vl_venda_FABRICANTE) / vl_lista_FABRICANTE
				end if
				
		
		
		'	%COM SOBRE FATURAMENTO VENDA
			if vl_venda_FABRICANTE = 0 then
				percRTFatVenda_FABRICANTE = 0
			else
				percRTFatVenda_FABRICANTE = (vl_RT_FABRICANTE/vl_venda_FABRICANTE)*100
				end if
			
       
            
		'	%COM + %DESCONTO
			percRTEDesc_FABRICANTE = perc_desconto_FABRICANTE + percRTFatVenda_FABRICANTE
			
			
		'   %MARGEM CONTRIB
		    if vl_venda_FABRICANTE = 0 then
		        marg_contrib_FABRICANTE = 0
		    else
		        marg_contrib_FABRICANTE = 100 * ((vl_venda_FABRICANTE - vl_RT_FABRICANTE - vl_ent_FABRICANTE) / vl_venda_FABRICANTE)
		    end if

			if vl_NF_Fabricante = 0 then
				marg_contrib_bruta_Fabricante = 0
			else
				marg_contrib_bruta_Fabricante = 100 * ((vl_NF_Fabricante - vl_ent_Fabricante) / vl_NF_Fabricante)
				end if

		    md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			s_cor="black"
			if vl_total_venda < 0 then s_cor="red"
			x = x & "	</tr>" & chr(13)
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
					"		<td class='MTBE' align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL:</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalFabricante) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_FABRICANTE) & "</span></td>" & chr(13) & _
					"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_FABRICANTE) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_FABRICANTE) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_FABRICANTE) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_FABRICANTE) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_FABRICANTE) & "</span></td>" & chr(13) & _
					        "       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Fabricante) & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Fabricante) & "</span></td>" & chr(13) & _
							"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Fabricante) & "%" & "</span></td>" & chr(13)
					else
					    x = x & _
					        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_FABRICANTE) & "</span></td>" & chr(13) & _
					        "       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Fabricante) & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Fabricante) & "</span></td>" & chr(13) & _
							"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Fabricante) & "%" & "</span></td>" & chr(13)
					end if
					x = x & _
					"	</tr>" & chr(13)
			'	CRIANDO ESPAÇOS EM BRANCO PARA ALINHAR A TABELA COM O TOTAL DE TODOS OS ESTADOS

			
			strEspacos = ""
			for iEspacos = 0 to Len(Cstr(n_reg_vetor)) + 1
				strEspacos = strEspacos & "&nbsp;"
				next
			
			x = x & "</table>" & chr(13)
			x = x & "<br>" & chr(13)
			x = x & "<table cellspacing=0>" & chr(13)
			x = x & _
					"	<tr>" & chr(13) & _
					"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;TOTAL</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
			x = x & "<br>" & chr(13)
			x = x & "	<tr nowrap style='background:azure'>" & chr(13) & _
					"		<td align='left' valign='bottom' style='background:white;' nowrap>" & strEspacos & "</td>" & chr(13) & _
					"		<td class='MDTE' style='width:" & Cstr(intLargIndicador) & "px' align='left' valign='bottom' nowrap><span class='R'>&nbsp;</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px' align='right' valign='bottom' nowrap><span class='Rd' style='font-weight:bold;'>Qtde</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Venda (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Fat Venda Total</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Lista (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Desc (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Desc</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>COM (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% COM</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px;' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>%COM<br>+<br>%Desc</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>RA Líquido (R$)</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto -RA -COM (R$)</span></td>" & chr(13) & _
					        "       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib (%)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (R$)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto (R$)</span></td>" & chr(13) & _
							"       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta (%)</span></td>" & chr(13)
					else
					    x = x & _
					        "		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado -RA -COM (R$)</span></td>" & chr(13) & _
					        "       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib Estimada (%)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (R$)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado (R$)</span></td>" & chr(13) & _
							"       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta Estimada (%)</span></td>" & chr(13)
					end if
					x = x & _
					"	</tr>" & chr(13)
			end if

    '	ÚLTIMO SUBTOTAL NO RELATÓRIO GRUPO PRODUTO
		if (rb_saida = COD_SAIDA_REL_GRUPO_PRODUTO)  then

            '	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
						if vl_total_final_venda = 0 then
							percFatVenda_ProdutoGrupo = 0
						else
							percFatVenda_ProdutoGrupo = (vl_venda_ProdutoGrupo/vl_total_final_venda)*100
							end if
					
					'	% DESCONTO
						if vl_lista_ProdutoGrupo = 0 then
							perc_desconto_ProdutoGrupo = 0
						else
							perc_desconto_ProdutoGrupo = 100 * (vl_lista_ProdutoGrupo-vl_venda_ProdutoGrupo) / vl_lista_ProdutoGrupo
							end if
					
					'	%COM SOBRE FATURAMENTO VENDA
						if vl_venda_ProdutoGrupo = 0 then
							percRTFatVenda_ProdutoGrupo = 0
						else
							percRTFatVenda_ProdutoGrupo = (vl_RT_ProdutoGrupo/vl_venda_ProdutoGrupo)*100
							end if
							
					'	%COM + %DESCONTO
						percRTEDesc_ProdutoGrupo = perc_desconto_ProdutoGrupo + percRTFatVenda_ProdutoGrupo
						
					'   %MARGEM CONTRIBUIÇÃO
					    if vl_venda_ProdutoGrupo = 0 then
					        marg_contrib_ProdutoGrupo = 0
					    else
					        marg_contrib_ProdutoGrupo = 100 * ((vl_venda_ProdutoGrupo - vl_RT_ProdutoGrupo - vl_ent_ProdutoGrupo) / vl_venda_ProdutoGrupo)
					    end if

						if vl_NF_ProdutoGrupo = 0 then
							marg_contrib_bruta_ProdutoGrupo = 0
						else
							marg_contrib_bruta_ProdutoGrupo = 100 * ((vl_NF_ProdutoGrupo - vl_ent_ProdutoGrupo) / vl_NF_ProdutoGrupo)
							end if

            x = x & "	</tr>" & chr(13)
			x = x & "	<tr nowrap>" & chr(13) & _
					"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
					"		<td class='MC ME' align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"Total do Grupo:</span></td>" & chr(13) & _
					"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalProdutoGrupo) & "</span></td>" & chr(13) & _
					"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_ProdutoGrupo) & "</span></td>" & chr(13) & _
					"		<td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_ProdutoGrupo) & "</span></td>" & chr(13) & _
					"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_ProdutoGrupo) & "</span></td>" & chr(13) & _
					"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_ProdutoGrupo) & "</span></td>" & chr(13) & _
					"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MC " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_ProdutoGrupo) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
						x = x & _
							"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_ProdutoGrupo) & "</span></td>" & chr(13) & _
							"       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_ProdutoGrupo) & "</span></td>" & chr(13) & _
							"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_ProdutoGrupo) & "</span></td>" & chr(13) & _
							"       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_ProdutoGrupo) & "%" & "</span></td>" & chr(13)
					else
						x = x & _
							"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_ProdutoGrupo) & "</span></td>" & chr(13) & _
							"       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_ProdutoGrupo) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_ProdutoGrupo) & "</span></td>" & chr(13) & _
							"		<td class='MC' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_ProdutoGrupo) & "</span></td>" & chr(13) & _
							"       <td class='MC MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_ProdutoGrupo) & "%" & "</span></td>" & chr(13)
					end if


		'	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
			    if vl_total_final_venda = 0 then
				percFatVenda_FABRICANTE = 0
			    else
				percFatVenda_FABRICANTE = (vl_venda_FABRICANTE/vl_total_final_venda)*100
				end if
		  
		
		'	% DESCONTO
			if vl_lista_FABRICANTE = 0 then
				perc_desconto_UF = 0
			else
				perc_desconto_FABRICANTE = 100 * (vl_lista_FABRICANTE-vl_venda_FABRICANTE) / vl_lista_FABRICANTE
				end if
				
		
		
		'	%COM SOBRE FATURAMENTO VENDA
			if vl_venda_FABRICANTE = 0 then
				percRTFatVenda_FABRICANTE = 0
			else
				percRTFatVenda_FABRICANTE = (vl_RT_FABRICANTE/vl_venda_FABRICANTE)*100
				end if
			
       
            
		'	%COM + %DESCONTO
			percRTEDesc_FABRICANTE = perc_desconto_FABRICANTE + percRTFatVenda_FABRICANTE
			
			
		'   %MARGEM CONTRIB
		    if vl_venda_FABRICANTE = 0 then
		        marg_contrib_FABRICANTE = 0
		    else
		        marg_contrib_FABRICANTE = 100 * ((vl_venda_FABRICANTE - vl_RT_FABRICANTE - vl_ent_FABRICANTE) / vl_venda_FABRICANTE)
		    end if

			if vl_NF_Fabricante = 0 then
				marg_contrib_bruta_Fabricante = 0
			else
				marg_contrib_bruta_Fabricante = 100 * ((vl_NF_Fabricante - vl_ent_Fabricante) / vl_NF_Fabricante)
				end if

		    md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			s_cor="black"
			if vl_total_venda < 0 then s_cor="red"
			x = x & "	</tr>" & chr(13)
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
					"		<td class='MTBE' align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL:</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalFabricante) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_FABRICANTE) & "</span></td>" & chr(13) & _
					"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_FABRICANTE) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_FABRICANTE) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_FABRICANTE) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_FABRICANTE) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_FABRICANTE) & "</span></td>" & chr(13) & _
					        "       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Fabricante) & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Fabricante) & "</span></td>" & chr(13) & _
							"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Fabricante) & "%" & "</span></td>" & chr(13)
					else
					    x = x & _
					        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_FABRICANTE) & "</span></td>" & chr(13) & _
					        "       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_FABRICANTE) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Fabricante) & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Fabricante) & "</span></td>" & chr(13) & _
							"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Fabricante) & "%" & "</span></td>" & chr(13)
					end if
					x = x & _
					"	</tr>" & chr(13)
			'	CRIANDO ESPAÇOS EM BRANCO PARA ALINHAR A TABELA COM O TOTAL DE TODOS OS ESTADOS

			
			strEspacos = ""
			for iEspacos = 0 to Len(Cstr(n_reg_vetor)) + 1
				strEspacos = strEspacos & "&nbsp;"
				next
			
			x = x & "</table>" & chr(13)
			x = x & "<br>" & chr(13)
			x = x & "<table cellspacing=0>" & chr(13)
			x = x & _
					"	<tr>" & chr(13) & _
					"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;TOTAL</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
			x = x & "<br>" & chr(13)
			x = x & "	<tr nowrap style='background:azure'>" & chr(13) & _
					"		<td align='left' valign='bottom' style='background:white;' nowrap>" & strEspacos & "</td>" & chr(13) & _
					"		<td class='MDTE' style='width:" & Cstr(intLargIndicador) & "px' align='left' valign='bottom' nowrap><span class='R'>&nbsp;</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px' align='right' valign='bottom' nowrap><span class='Rd' style='font-weight:bold;'>Qtde</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Venda (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Fat Venda Total</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Lista (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Desc (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Desc</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>COM (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% COM</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px;' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>%COM<br>+<br>%Desc</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>RA Líquido (R$)</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto -RA -COM (R$)</span></td>" & chr(13) & _
					        "       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib (%)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (R$)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto (R$)</span></td>" & chr(13) & _
							"       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta (%)</span></td>" & chr(13)
					else
					    x = x & _
					        "		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado -RA -COM (R$)</span></td>" & chr(13) & _
					        "       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib Estimada (%)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (R$)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado (R$)</span></td>" & chr(13) & _
							"       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta Estimada (%)</span></td>" & chr(13)
					end if
					x = x & _
					"	</tr>" & chr(13)
			end if
        
     '	ÚLTIMO SUBTOTAL NO RELATÓRIO EMPRESA
		if (rb_saida = COD_SAIDA_REL_EMPRESA)  then

		'	PERCENTUAL RELATIVO AO FATURAMENTO VENDA TOTAL
			    if vl_total_final_venda = 0 then
				percFatVenda_Empresa = 0
			    else
				percFatVenda_Empresa = (vl_venda_Empresa/vl_total_final_venda)*100
				end if		  
		
		'	% DESCONTO
			if vl_lista_Empresa = 0 then
				perc_desconto_Empresa = 0
			else
				perc_desconto_Empresa = 100 * (vl_lista_Empresa-vl_venda_Empresa) / vl_lista_Empresa
				end if			
		
		'	%COM SOBRE FATURAMENTO VENDA
			if vl_venda_Empresa = 0 then
				percRTFatVenda_Empresa = 0
			else
				percRTFatVenda_Empresa = (vl_RT_Empresa/vl_venda_Empresa)*100
				end if	   
            
		'	%COM + %DESCONTO
			percRTEDesc_Empresa = perc_desconto_Empresa + percRTFatVenda_Empresa
			
			
		'   %MARGEM CONTRIB
		    if vl_venda_Empresa = 0 then
		        marg_contrib_Empresa = 0
		    else
		        marg_contrib_Empresa = 100 * ((vl_venda_Empresa - vl_RT_Empresa - vl_ent_Empresa) / vl_venda_Empresa)
		    end if

			if vl_NF_Empresa = 0 then
				marg_contrib_bruta_Empresa = 0
			else
				marg_contrib_bruta_Empresa = 100 * ((vl_NF_Empresa - vl_ent_Empresa) / vl_NF_Empresa)
				end if

		    md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			s_cor="black"        
			if vl_total_venda < 0 then s_cor="red"            
			x = x & "	</tr>" & chr(13)
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
					"		<td class='MTBE' align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL:</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalEmpresa) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_venda_Empresa) & "</span></td>" & chr(13) & _
					"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFatVenda_Empresa) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lista_Empresa) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_desconto_Empresa) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desconto_Empresa) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT_Empresa) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTFatVenda_Empresa) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percRTEDesc_Empresa) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido_Empresa) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_Empresa) & "</span></td>" & chr(13) & _
					        "       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_Empresa) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Empresa) & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Empresa) & "</span></td>" & chr(13) & _
							"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Empresa) & "%" & "</span></td>" & chr(13)
					else
					    x = x & _
					        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_Empresa) & "</span></td>" & chr(13) & _
					        "       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_Empresa) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_NF_Empresa) & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_Empresa) & "</span></td>" & chr(13) & _
							"       <td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_Empresa) & "%" & "</span></td>" & chr(13)
					end if
					x = x & _
					"	</tr>" & chr(13)

			'	CRIANDO ESPAÇOS EM BRANCO PARA ALINHAR A TABELA COM O TOTAL DE TODOS OS ESTADOS
	

			strEspacos = ""
			for iEspacos = 0 to Len(Cstr(n_reg_vetor)) + 1
				strEspacos = strEspacos & "&nbsp;"
				next
			
			x = x & "</table>" & chr(13)
			x = x & "<br>" & chr(13)
			x = x & "<table cellspacing=0>" & chr(13)
			x = x & _
					"	<tr>" & chr(13) & _
					"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MDTE' " & strColSpanTodasColunas & " align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;TOTAL</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
			x = x & "<br>" & chr(13)
			x = x & "	<tr nowrap style='background:azure'>" & chr(13) & _
					"		<td align='left' valign='bottom' style='background:white;' nowrap>" & strEspacos & "</td>" & chr(13) & _
					"		<td class='MDTE' style='width:" & Cstr(intLargIndicador) & "px' align='left' valign='bottom' nowrap><span class='R'>&nbsp;</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & CStr(intLargQtdeProduto) & "px' align='right' valign='bottom' nowrap><span class='Rd' style='font-weight:bold;'>Qtde</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Venda (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Fat Venda Total</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Lista (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Desc (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% Desc</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>COM (R$)</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>% COM</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColPerc) & "px;' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>%COM<br>+<br>%Desc</span></td>" & chr(13) & _
					"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>RA Líquido (R$)</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto -RA -COM (R$)</span></td>" & chr(13) & _
					        "       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib (%)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (R$)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto (R$)</span></td>" & chr(13) & _
							"       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta (%)</span></td>" & chr(13)
					else
					    x = x & _
					        "		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado -RA -COM (R$)</span></td>" & chr(13) & _
					        "       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Contrib Estimada (%)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>VL Fat (R$)</span></td>" & chr(13) & _
							"		<td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Lucro Bruto Estimado (R$)</span></td>" & chr(13) & _
							"       <td class='MTD' style='width:" & Cstr(intLargColMonetario) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;margin-right:0px;'>Margem Bruta Estimada (%)</span></td>" & chr(13)
					end if
					x = x & _
					"	</tr>" & chr(13)
			end if



	' MOSTRA TOTAL
	'	% DO FATURAMENTO VENDA TOTAL
		if vl_total_final_venda = 0 then
			percFinalFatVenda = 0
		else
			percFinalFatVenda = (vl_total_venda/vl_total_final_venda)*100
			end if
			
	'	% DESCONTO
		if vl_total_lista = 0 then
			percFinalDesc = 0
		else
			percFinalDesc = 100 * (vl_total_lista-vl_total_venda) / vl_total_lista
			end if
			
	'	% COMISSÃO SOBRE FATURAMENTO VENDA
		if vl_total_final_venda = 0 then
			percFinalRTFatVenda = 0
		else
			percFinalRTFatVenda = (vl_total_RT/vl_total_venda)*100
			end if
			
	'	%COMISSÃO + %DESCONTO
	'	O CARLOS SOLICITOU EM 12/MAR/2008 QUE SEJA FEITA A SOMA SIMPLES DOS 2 PERCENTUAIS
		percFinalRTEDescFatVenda = percFinalDesc + percFinalRTFatVenda
		
	'   %FINAL DE MARGEM CONTRIBUIÇÃO
	    if vl_total_venda = 0 then
	        marg_contrib_final = 0
	    else
	        marg_contrib_final = 100 * ((vl_total_venda - total_entrada - vl_total_RT) / vl_total_venda)
	    end if

		if vl_total_NF = 0 then
			marg_contrib_bruta_final = 0
		else
			marg_contrib_bruta_final = 100 * ((vl_total_NF - total_entrada) / vl_total_NF)
			end if

		if rb_saida = COD_SAIDA_REL_PRODUTO then
			s_cor="black"
            md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			if intQtdeTotalProdutos < 0 then s_cor="red"
			if vl_total_venda < 0 then s_cor="red"
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td align='left' valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
					"		<td colspan='2' class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL:</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalProdutos) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_venda) & "</span></td>" & chr(13) & _
					"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_lista) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_desconto) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalDesc) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTEDescFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					else
					    x = x & _
					        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					end if
					x = x & _
					"	</tr>" & chr(13)

		elseif rb_saida = COD_SAIDA_REL_INDICADOR then
			s_cor="black"
            md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			if vl_total_venda < 0 then s_cor="red"
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
					"		<td colspan='2' class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL:</span></td>" & chr(13) & _
                    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalProdutos) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_venda) & "</span></td>" & chr(13) & _
					"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_lista) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_desconto) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalDesc) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTEDescFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
				        	"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					else
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
				        	"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					end if
					x = x  & _
					"	</tr>" & chr(13)
		elseif rb_saida = COD_SAIDA_REL_INDICADOR_UF then
			if vl_total_venda < 0 then s_cor="red"
            md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL:</span></td>" & chr(13) & _
                    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalProdutos) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_venda) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_lista) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_desconto) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalDesc) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTEDescFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					else
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					end if
					x = x & _
					"	</tr>" & chr(13)
		elseif rb_saida = COD_SAIDA_REL_CIDADE_UF then
			if vl_total_venda < 0 then s_cor="red"
            md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL:</span></td>" & chr(13) & _
                    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalProdutos) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_venda) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_lista) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_desconto) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalDesc) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTEDescFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					else
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					end if 
					x = x & _
					"	</tr>" & chr(13)
		elseif rb_saida = COD_SAIDA_REL_FABRICANTE then
			if vl_total_venda < 0 then s_cor="red"
            md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL:</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalProdutos) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_venda) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_lista) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_desconto) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalDesc) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTEDescFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					else
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					end if 
					x = x & _
					"	</tr>" & chr(13)

        elseif rb_saida = COD_SAIDA_REL_GRUPO_PRODUTO then
            
            for i=UBound(vTotGrupo) to 1 step -1
                if vl_total_final_venda = 0 then
                    vTotGrupo(i).c5 = 0
                else
                    vTotGrupo(i).c5 = (vTotGrupo(i).c4/vl_total_final_venda)*100 '%Fat Venda Total
                end if
                if vTotGrupo(i).c6 = 0 then
                    vTotGrupo(i).c8 = 0
                else
                    vTotGrupo(i).c8 = 100 * (vTotGrupo(i).c6-vTotGrupo(i).c4) / vTotGrupo(i).c6 '%Desconto
                end if
                if vTotGrupo(i).c4 = 0 then
                    vTotGrupo(i).c10 = 0
                else
                    vTotGrupo(i).c10 = (vTotGrupo(i).c9/vTotGrupo(i).c4)*100 '%comissão
                end if
                vTotGrupo(i).c11 = vTotGrupo(i).c10 + vTotGrupo(i).c8 '%comissão + %desconto
                'Margem Contrib %
				if vTotGrupo(i).c4 = 0 then
                    vTotGrupo(i).c14 = 0
                else
                    vTotGrupo(i).c14 = 100 * ((vTotGrupo(i).c4 - vTotGrupo(i).c15 - vTotGrupo(i).c9) / vTotGrupo(i).c4)
                end if
				'Margem Contrib Bruta (%)
				if vTotGrupo(i).c16 = 0 then
					vTotGrupo(i).c18 = 0
				else
					vTotGrupo(i).c18 = 100 * ((vTotGrupo(i).c16 - vTotGrupo(i).c15) / vTotGrupo(i).c16)
					end if
                

    if vl_total_venda = 0 then
	        marg_contrib_final = 0
	    else
	        marg_contrib_final = 100 * ((vl_total_venda - total_entrada - vl_total_RT) / vl_total_venda)
	    end if

		if vl_total_NF = 0 then
			marg_contrib_bruta_final = 0
		else
			marg_contrib_bruta_final = 100 * ((vl_total_NF - total_entrada) / vl_total_NF)
			end if

                s_cor="black"
				if vTotGrupo(i).c3 < 0 then s_cor="red"
				if vTotGrupo(i).c4 < 0 then s_cor="red"
                x = x & "	<tr nowrap style='background:honeydew'>" & chr(13) & _
							"		<td style='background:white'>&nbsp;</td>" & chr(13) & _
							"		<td class='MC MD ME'><p class='C' style='color:" & s_cor & ";'>" & vTotGrupo(i).c1 & " - " & vTotGrupo(i).c2 & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(vTotGrupo(i).c3) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(i).c4) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(vTotGrupo(i).c5) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(i).c6) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(i).c7) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(vTotGrupo(i).c8) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(i).c9) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(vTotGrupo(i).c10) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(vTotGrupo(i).c11) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(i).c12) & "</p></td>" & chr(13)
                if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
                    x = x & "		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(i).c13) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(vTotGrupo(i).c14) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(i).c16) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(i).c17) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(vTotGrupo(i).c18) & "%" & "</p></td>" & chr(13)
                else
                    x = x & "		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(i).c13) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(vTotGrupo(i).c14) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(i).c16) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(i).c17) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(vTotGrupo(i).c18) & "%" & "</p></td>" & chr(13)
                end if

                x = x & "	</tr>" & chr(13)
                
            next


			if vl_total_venda < 0 then s_cor="red"
            md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL:</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalProdutos) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_venda) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_lista) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_desconto) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalDesc) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTEDescFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					else
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					end if 
					x = x & _
					"	</tr>" & chr(13)

        elseif rb_saida = COD_SAIDA_REL_EMPRESA then
			if vl_total_venda < 0 then s_cor="red"
            md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL:</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalProdutos) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_venda) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_lista) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_desconto) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalDesc) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTEDescFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					else
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
					end if 
					x = x & _
					"	</tr>" & chr(13)
		else
			s_cor="black"
            md = ""
			if rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA then md = "MD"
			if vl_total_venda < 0 then s_cor="red"
			x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
					"		<td valign='bottom' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
					"		<td class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL:</span></td>" & chr(13) & _
                    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(intQtdeTotalProdutos) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_venda) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_lista) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_desconto) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalDesc) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(percFinalRTEDescFatVenda) & "%" & "</span></td>" & chr(13) & _
					"		<td class='MTB " & md & "' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) 
					if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
				    else
					    x = x & _
					        "       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_liquido_total) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_final) & "%" & "</span></td>" & chr(13) & _
							"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_NF) & "</span></td>" & chr(13) & _
							"       <td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(lucro_bruto_total) & "</span></td>" & chr(13) & _
							"		<td class='MTBD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_perc(marg_contrib_bruta_final) & "%" & "</span></td>" & chr(13)
				    end if
				    x = x & _
					"	</tr>" & chr(13)
			end if
		end if


  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!!
	if n_reg_BD = 0 then
		x = cab_table & cab
		x = x & "	<tr nowrap>" & chr(13) & _
				"		<td>&nbsp;</td>" & chr(13) & _
				"		<td class='MT ALERTA' " & strColSpanTodasColunas & " align='center'><span class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' FECHA TABELA
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL</title>
	</head>


<style type="text/css">
	.style1
	{
		height: 46px;
	}
</style>


<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';
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
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_cadastro_inicio" id="c_dt_cadastro_inicio" value="<%=c_dt_cadastro_inicio%>">
<input type="hidden" name="c_dt_cadastro_termino" id="c_dt_cadastro_termino" value="<%=c_dt_cadastro_termino%>">
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="<%=c_dt_entregue_inicio%>">
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">
<input type="hidden" name="c_captador" id="c_captador" value="<%=c_captador%>">
<input type="hidden" name="c_cnpj_cpf" id="c_cnpj_cpf" value="<%=c_cnpj_cpf%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">
<input type="hidden" name="rb_periodo" id="rb_periodo" value="<%=rb_periodo%>">
<input type="hidden" name="rb_saida" id="rb_saida" value="<%=rb_saida%>">
<input type="hidden" name="op_forma_pagto" id="op_forma_pagto" value="<%=op_forma_pagto%>">
<input type="hidden" name="c_forma_pagto_qtde_parc" id="c_forma_pagto_qtde_parc" value="<%=c_forma_pagto_qtde_parc%>">
<input type="hidden" name="c_cst" id="c_cst" value="<%=c_cst%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="1008" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" class="style1"><span class="PEDIDO">Relatório Gerencial de Vendas</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<%
	s_filtro = "<table width='1008' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>"
	
	if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
		s = ""
		s_aux = c_dt_cadastro_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " a "
		s_aux = c_dt_cadastro_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>Período de Cadastro:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
		s = ""
		s_aux = c_dt_entregue_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " a "
		s_aux = c_dt_entregue_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>Período de Entrega:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	s_aux = ""
	if c_fabricante <> "" then s_aux = x_fabricante(c_fabricante)
	s = c_fabricante
	if (s <> "") And (s_aux <> "") then s = s & " - " & s_aux
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Fabricante:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	s = c_produto
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Produto:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
		s = ""
		for i=Lbound(v_cst) to Ubound(v_cst)
			if v_cst(i) <> "" then
				if s <> "" then s = s & ", "
				s = s & v_cst(i)
				end if
			next
		if s = "" then s = "todos"
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>CST:&nbsp;</span></td><td valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if

	s = c_vendedor
	if s = "" then 
		s = "todos"
	else
		if (s_nome_vendedor <> "") And (s_nome_vendedor <> c_vendedor) then s = s & " (" & s_nome_vendedor & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Vendedor:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	s = c_indicador
	if s = "" then 
		s = "todos"
	else
		s = s & " (" & x_orcamentista_e_indicador(c_indicador) & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Indicador:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	s = c_captador
	if s = "" then 
		s = "todos"
	else
		s = s & " (" & x_usuario(c_captador) & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Captador:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = c_grupo_pedido_origem
	if s = "" then 
		s = "todos"
	else
        v_grupo_pedido_origem = split(c_grupo_pedido_origem, ", ")
        s = ""
        for i = Lbound(v_grupo_pedido_origem) to Ubound(v_grupo_pedido_origem)
            if s <> "" then s = s & ", "
		    s = s & obtem_descricao_tabela_t_codigo_descricao("PedidoECommerce_Origem_Grupo", v_grupo_pedido_origem(i))
        next
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Origem Pedido (Grupo):&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = c_pedido_origem
	if s = "" then 
		s = "todos"
	else
		v_pedido_origem = split(c_pedido_origem, ", ")
        s = ""
        for i = Lbound(v_pedido_origem) to Ubound(v_pedido_origem)
            if s <> "" then s = s & ", "
		    s = s & obtem_descricao_tabela_t_codigo_descricao("PedidoECommerce_Origem", v_pedido_origem(i))
        next
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Origem do Pedido:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	if False then
		s = c_cnpj_cpf
		if s = "" then 
			s = "todos"
		else
			s = cnpj_cpf_formata(c_cnpj_cpf)
			end if
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>CNPJ/CPF:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
    s = c_grupo
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Grupo(s):&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = c_subgrupo
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Subgrupo(s):&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	s = rb_tipo_cliente
	if s = "" then
		s = "todos"
	elseif s = ID_PF then
		s = "Pessoa Física"
	elseif s = ID_PJ then
		s = "Pessoa Jurídica"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Tipo de Cliente:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	if rb_tipo_cliente <> ID_PF then
		s = ""
		if ckb_contribuinte_icms_sim <> "" then
			if s <> "" then s = s & ", "
			s = s & "Contribuinte"
			end if

		if ckb_contribuinte_icms_nao <> "" then
			if s <> "" then s = s & ", "
			s = s & "Não Contribuinte"
			end if

		if ckb_contribuinte_icms_isento <> "" then
			if s <> "" then s = s & ", "
			s = s & "Isento"
			end if

		if s = "" then s = "todos"

		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>Contribuinte ICMS:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if

    if (rb_saida <> COD_SAIDA_REL_CIDADE_UF) then 
        if c_uf_saida <> "" then
			    s = c_uf_saida
		    else
			    s = "N.I."
			    end if
		    s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				       "<span class='N'>UF:&nbsp;</span></td><td align='left' valign='top'>" & _
				       "<span class='N'>" & s & "</span></td></tr>"
    end if

    s = c_empresa
	if s = "" then 
		s = "todas"
	else
		s = obtem_apelido_empresa_NFe_emitente(c_empresa)
		end if
	 s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				       "<span class='N'>Empresa:&nbsp;</span></td><td align='left' valign='top'>" & _
				       "<span class='N'>" & s & "</span></td></tr>"

	s_filtro_loja = ""
	for i = Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
				s_filtro_loja = s_filtro_loja & v_loja(i)
			else
				if (v(Lbound(v))<>"") And (v(Ubound(v))<>"") then 
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Lbound(v)) & " a " & v(Ubound(v))
				elseif (v(Lbound(v))<>"") And (v(Ubound(v))="") then
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Lbound(v)) & " e acima"
				elseif (v(Lbound(v))="") And (v(Ubound(v))<>"") then
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Ubound(v)) & " e abaixo"
					end if
				end if
			end if
		next
	s = s_filtro_loja
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Loja(s):&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	if (True) Or (rb_saida = COD_SAIDA_REL_UF) Or (rb_saida = COD_SAIDA_REL_VENDEDOR) then
		if op_forma_pagto = "" then
			s = "N.I."
		else
			s = x_opcao_forma_pagamento(op_forma_pagto)
			end if
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>Forma Pagto:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	if c_forma_pagto_qtde_parc <> "" then
		s = c_forma_pagto_qtde_parc
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>Nº Parcelas:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	if (rb_saida = COD_SAIDA_REL_CIDADE_UF) then
		if c_loc_uf <> "" then
			s = c_loc_uf
		else
			s = "N.I."
			end if
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>UF:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"

		if c_loc_escolhidas <> "" then
			s = c_loc_escolhidas
		elseif c_loc_digitada <> "" then
			s = "iniciadas com " & c_loc_digitada
		else
			s = "N.I."
			end if
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				   "<span class='N'>Cidade(s):&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="1008" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
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
