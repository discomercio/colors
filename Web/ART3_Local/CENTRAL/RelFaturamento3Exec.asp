<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelFaturamento3Exec.asp
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
	
	Const COD_TIPO_AGRUPAMENTO__GRUPO = "Grupo"
	Const COD_TIPO_AGRUPAMENTO__PRODUTO = "Produto"
	
	class cl_REL_FAT3_TOT_GRUPO
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
	if Not operacao_permitida(OP_CEN_REL_FATURAMENTO3, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, lista_loja, s_filtro_loja, v_loja, v, i, flag_ok
	dim c_dt_inicio, c_dt_termino, c_loja, c_fabricante, c_produto, c_grupo, c_vendedor, c_indicador, c_pedido
    dim v_grupo_pedido_origem, v_pedido_origem, c_grupo_pedido_origem, c_pedido_origem, c_empresa
	dim s_nome_vendedor
	dim op_forma_pagto, c_forma_pagto_qtde_parc
	dim rb_tipo_cliente, c_cliente_cnpj_cpf
	dim c_uf_pesq, v_uf_pesq
	dim rb_tipo_agrupamento
	dim c_cst, v_cst, v_cst_aux

	alerta = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	c_indicador = Ucase(Trim(Request.Form("c_indicador")))
	c_pedido = Ucase(Trim(Request.Form("c_pedido")))
    c_grupo_pedido_origem = Trim(Request.Form("c_grupo_pedido_origem"))
    c_pedido_origem = Trim(Request.Form("c_pedido_origem"))
    c_empresa = Trim(Request.Form("c_empresa"))
	c_cliente_cnpj_cpf=Trim(Request.Form("c_cliente_cnpj_cpf"))
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

	if c_pedido <> "" then
		if normaliza_num_pedido(c_pedido) <> "" then c_pedido = normaliza_num_pedido(c_pedido)
		end if

	c_loja = Trim(Request.Form("c_loja"))
	lista_loja = substitui_caracteres(c_loja,chr(10),"")
	v_loja = split(lista_loja,chr(13),-1)
	
	op_forma_pagto = Trim(Request.Form("op_forma_pagto"))
	c_forma_pagto_qtde_parc = retorna_so_digitos(Trim(Request.Form("c_forma_pagto_qtde_parc")))
	rb_tipo_cliente = Trim(Request.Form("rb_tipo_cliente"))
	c_uf_pesq = Ucase(Trim(Request.Form("c_uf_pesq")))
	rb_tipo_agrupamento = Trim(Request.Form("rb_tipo_agrupamento"))
	
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
		if c_pedido <> "" then
			s = "SELECT pedido FROM t_PEDIDO WHERE (pedido='" & c_pedido & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "PEDIDO " & c_pedido & " NÃO ESTÁ CADASTRADO."
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

    if alerta = "" then
		if c_cliente_cnpj_cpf <> "" then
			if Not cnpj_cpf_ok(retorna_so_digitos(c_cliente_cnpj_cpf)) then
				alerta=texto_add_br(alerta)
				alerta = alerta & "CNPJ/CPF do cliente é inválido."
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
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_termino
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			if c_dt_inicio = "" then c_dt_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
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
' inicializa_cl_REL_FAT3_TOT_GRUPO
'
sub inicializa_cl_REL_FAT3_TOT_GRUPO(byref rv)
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
dim r
dim s, s_where, s_where_venda, s_where_devolucao, s_where_loja, s_where_cst, s_cor, s_where_temp
dim s_aux, s_sql, cab_table, cab, n_reg, n_reg_grupo, n_reg_total, x, fabricante_a, grupo_a
dim perc, perc2, perc3, vl_total_saida, vl_total_entrada, vl_sub_total_saida, vl_sub_total_entrada, vl_sub_total_saida_grupo, vl_sub_total_entrada_grupo
dim vl_total_valor_RT, vl_total_lucro_liquido, vl_sub_total_valor_RT, vl_sub_total_lucro_liquido, vl_sub_total_valor_RT_grupo, vl_sub_total_lucro_liquido_grupo
dim vl_lucro_liquido
dim i, j, v, qtde_total, qtde_sub_total, qtde_sub_total_grupo, qtde_fabricantes
dim vTotGrupo, idxSelecionado, linha_tot_grupo, s_codigo_grupo, s_descricao_grupo, s_class
dim vOrdGrupo

'	CRITÉRIOS COMUNS
	s_where = ""
	if c_grupo <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.grupo = '" & c_grupo & "')"
		
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.grupo IS NOT NULL)"
		end if
	
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.vendedor = '" & c_vendedor & "')"
		end if

	if c_pedido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.pedido = '" & c_pedido & "')"
		end if

    if c_empresa <> "" then
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_PEDIDO.id_nfe_emitente = '" & c_empresa & "')"
    end if
		
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if
		
'	CRITÉRIO: FORMA DE PAGAMENTO
	s = ""
	if op_forma_pagto <> "" then
		s = " ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_A_VISTA & ") AND (t_PEDIDO__BASE.av_forma_pagto = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELA_UNICA & ") AND (t_PEDIDO__BASE.pu_forma_pagto = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA & ") AND (t_PEDIDO__BASE.pce_forma_pagto_entrada = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA & ") AND (t_PEDIDO__BASE.pce_forma_pagto_prestacao = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA & ") AND (t_PEDIDO__BASE.pse_forma_pagto_prim_prest = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA & ") AND (t_PEDIDO__BASE.pse_forma_pagto_demais_prest = " & op_forma_pagto & ") )"
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

'	CRITÉRIO: QUANTIDADE DE PARCELAS
	s = ""
	if c_forma_pagto_qtde_parc <> "" then
		s = " (t_PEDIDO__BASE.qtde_parcelas = " & c_forma_pagto_qtde_parc & ")"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
		
'	CRITÉRIO: TIPO DE CLIENTE
	s = ""
	if rb_tipo_cliente <> "" then
		s = " (t_CLIENTE.tipo = '" & rb_tipo_cliente & "')"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRITÉRIO: UF DO CLIENTE
	s = ""
	if c_uf_pesq <> "" then
		v_uf_pesq = Split(c_uf_pesq, ", ")
		s = ""
		for i=LBound(v_uf_pesq) to UBound(v_uf_pesq)
			if Trim("" & v_uf_pesq(i)) <> "" then
				 if s <> "" then s = s & ","
				 s = s & "'" & Trim("" & v_uf_pesq(i)) & "'"
				 end if
			next
		s = " (t_CLIENTE.uf IN (" & s & "))"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRITÉRIO: CNPJ/CPF CLIENTE
    s = ""
    if c_cliente_cnpj_cpf <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_CLIENTE.cnpj_cpf = '" & retorna_so_digitos(c_cliente_cnpj_cpf) & "')"
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
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
    v_pedido_origem = split(c_pedido_origem, ", ")
    if c_pedido_origem <> "" then
        for i = LBound(v_pedido_origem) to UBound(v_pedido_origem)
            if s_where_temp <> "" then s_where_temp = s_where_temp & ", "
            s_where_temp = s_where_temp & "'" & v_pedido_origem(i) & "'"
        next
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_PEDIDO.marketplace_codigo_origem IN (" & s_where_temp & "))"
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

'	CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	s_where_venda = ""
	if IsDate(c_dt_inicio) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if

	if c_fabricante <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_ESTOQUE_MOVIMENTO.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_ESTOQUE_MOVIMENTO.produto = '" & c_produto & "')"
		end if
	
'	CRITÉRIOS PARA DEVOLUÇÕES
	s_where_devolucao = ""
	if IsDate(c_dt_inicio) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
	if c_fabricante <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.produto = '" & c_produto & "')"
		end if

' 	A) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
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
	s_sql = "SELECT t_ESTOQUE_MOVIMENTO.fabricante AS fabricante, t_ESTOQUE_MOVIMENTO.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" t_PRODUTO.grupo," & _
			" t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
			" Sum((t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda)*(t_PEDIDO__BASE.perc_RT/100)) AS valor_RT," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_NF) AS valor_saida," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde) AS qtde," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_ESTOQUE_ITEM.vl_custo2) AS valor_entrada" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
			" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
			" INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
			" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_PEDIDO_ITEM ON ((t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto))" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
			" LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque='" & ID_ESTOQUE_ENTREGUE & "')" & _
			s & _
			" GROUP BY t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, t_PRODUTO.grupo, t_PRODUTO_GRUPO.descricao"
	
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
	s_sql = s_sql & " UNION " & _
			"SELECT t_PEDIDO_ITEM_DEVOLVIDO.fabricante AS fabricante, t_PEDIDO_ITEM_DEVOLVIDO.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" t_PRODUTO.grupo," & _
			" t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
			" Sum(-(t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda)*(t_PEDIDO__BASE.perc_RT/100)) AS valor_RT," & _
			" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_NF) AS valor_saida," & _
			" Sum(-t_ESTOQUE_ITEM.qtde) AS qtde," & _
			" Sum(-t_ESTOQUE_ITEM.qtde*t_ESTOQUE_ITEM.vl_custo2) AS valor_entrada" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
			" INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)" & _
			" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_ESTOQUE ON (t_PEDIDO_ITEM_DEVOLVIDO.id=t_ESTOQUE.devolucao_id_item_devolvido)" & _
			" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_ESTOQUE_ITEM.produto))" & _ 
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			" LEFT JOIN t_PRODUTO ON ((t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_PRODUTO.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_PRODUTO.produto))" & _
			" LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
			s & _
			" GROUP BY t_PEDIDO_ITEM_DEVOLVIDO.fabricante, t_PEDIDO_ITEM_DEVOLVIDO.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, t_PRODUTO.grupo, t_PRODUTO_GRUPO.descricao"
	
	if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then
		s_sql = s_sql & " ORDER BY fabricante, produto, descricao, descricao_html, qtde DESC"
	elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
		s_sql = s_sql & " ORDER BY fabricante, t_PRODUTO.grupo, produto, descricao, descricao_html, qtde DESC"
		end if

  ' CABEÇALHO
	cab_table = "<table cellspacing=0>" & chr(13)
	cab = "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='MDTE tdCodProd' style='vertical-align:bottom' nowrap><p class='R'>Código</p></td>" & chr(13) & _
		  "		<td class='MTD tdDescrProd' style='vertical-align:bottom' nowrap><p class='R'>Descrição</p></td>" & chr(13) & _
		  "		<td class='MTD tdQtde' align='right' style='vertical-align:bottom' nowrap><p class='Rd' style='font-weight:bold;'>Qtde</p></td>" & chr(13) & _
		  "		<td class='MTD tdFatTot' align='right' style='vertical-align:bottom'><p class='Rd' style='font-weight:bold;'>Faturamento Total (" & SIMBOLO_MONETARIO & ")</p></td>" & chr(13) & _
		  "		<td class='MTD tdCmvTot' align='right' style='vertical-align:bottom'><p class='Rd' style='font-weight:bold;'>CMV Total (" & SIMBOLO_MONETARIO & ")</p></td>" & chr(13) & _
		  "		<td class='MTD tdLucro' align='right' style='vertical-align:bottom'><p class='Rd' style='font-weight:bold;'>Lucro (" & SIMBOLO_MONETARIO & ")</p></td>" & chr(13) & _
		  "		<td class='MTD tdPercLucroTot' align='right' style='vertical-align:bottom'><p class='Rd' style='font-weight:bold;'>% do Lucro Total</p></td>" & chr(13) & _
		  "		<td class='MTD tdVlRt' align='right' style='vertical-align:bottom'><p class='Rd' style='font-weight:bold;'>COM (" & SIMBOLO_MONETARIO & ")</p></td>" & chr(13) & _
		  "		<td class='MTD tdRtPercFat' align='right' style='vertical-align:bottom'><p class='Rd' style='font-weight:bold;'>COM (% do Fat)</p></td>" & chr(13) & _
		  "		<td class='MTD tdLucroLiq' align='right' style='vertical-align:bottom'><p class='Rd' style='font-weight:bold;'>Lucro Bruto (" & SIMBOLO_MONETARIO & ")</p></td>" & chr(13) & _
		  "		<td class='MTD tdLucroLiqPercFat' align='right' style='vertical-align:bottom'><p class='Rd' style='font-weight:bold;'>Lucro Bruto (% do Fat)</p></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	redim vTotGrupo(0)
	set vTotGrupo(UBound(vTotGrupo)) = new cl_REL_FAT3_TOT_GRUPO
	inicializa_cl_REL_FAT3_TOT_GRUPO vTotGrupo(UBound(vTotGrupo))

	vl_total_saida = 0
	vl_total_entrada = 0
	vl_total_valor_RT = 0
	vl_total_lucro_liquido = 0
	set r = cn.execute(s_sql)
	n_reg = 0
	do while Not r.Eof
		n_reg = n_reg + 1
		vl_total_saida = vl_total_saida + r("valor_saida")
		vl_total_entrada = vl_total_entrada + r("valor_entrada")
		
		vl_total_valor_RT = vl_total_valor_RT + r("valor_RT")
		vl_total_lucro_liquido = vl_total_lucro_liquido + (r("valor_saida")-r("valor_entrada")) - r("valor_RT")
		r.MoveNext
		loop

	if n_reg > 0 then r.MoveFirst
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_total = 0
	qtde_sub_total = 0
	qtde_sub_total_grupo = 0
	qtde_fabricantes = 0
	
	fabricante_a = "XXXXX"
	grupo_a = "XXXXX"
	do while Not r.Eof
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante"))<>fabricante_a then
			fabricante_a = Trim("" & r("fabricante"))
			qtde_fabricantes = qtde_fabricantes + 1
			if n_reg_total > 0 then
			  ' Fecha o grupo de produtos anterior
				if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
					if vl_sub_total_saida_grupo = 0 then
						perc = 0
					else
						perc = ((vl_sub_total_saida_grupo-vl_sub_total_entrada_grupo)/vl_sub_total_saida_grupo)*100
						end if
					
					if vl_sub_total_saida_grupo = 0 then
						perc2 = 0
					else
						perc2 = (vl_sub_total_valor_RT_grupo/vl_sub_total_saida_grupo)*100
						end if

					if vl_sub_total_saida_grupo = 0 then
						perc3 = 0
					else
						perc3 = (vl_sub_total_lucro_liquido_grupo/vl_sub_total_saida_grupo)*100
						end if

					s_cor="black"
					if qtde_sub_total_grupo < 0 then s_cor="red"
					if vl_sub_total_saida_grupo < 0 then s_cor="red"
					if vl_sub_total_entrada_grupo < 0 then s_cor="red"
					if (vl_sub_total_saida_grupo-vl_sub_total_entrada_grupo) < 0 then s_cor="red"
					x = x & "	<tr nowrap>" & chr(13) & _
							"		<td class='MTBE' colspan='2' nowrap><p class='Cd' style='color:" & s_cor & ";'>" & _
							"Total do Grupo:</p></td>" & chr(13) & _
							"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_sub_total_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_entrada_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida_grupo-vl_sub_total_entrada_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_valor_RT_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc2) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_lucro_liquido_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc3) & "%" & "</p></td>" & chr(13) & _
							"	</tr>" & chr(13)
					Response.Write x
					x="<br />" & chr(13)
					end if
				
			  ' FECHA TABELA DO FABRICANTE ANTERIOR
				if vl_sub_total_saida = 0 then
					perc = 0
				else
					perc = ((vl_sub_total_saida-vl_sub_total_entrada)/vl_sub_total_saida)*100
					end if
				
				if vl_sub_total_saida = 0 then
					perc2 = 0
				else
					perc2 = (vl_sub_total_valor_RT/vl_sub_total_saida)*100
					end if
				
				if vl_sub_total_saida = 0 then
					perc3 = 0
				else
					perc3 = (vl_sub_total_lucro_liquido/vl_sub_total_saida)*100
					end if
				
				s_cor="black"
				if qtde_sub_total < 0 then s_cor="red"
				if vl_sub_total_saida < 0 then s_cor="red"
				if vl_sub_total_entrada < 0 then s_cor="red"
				if (vl_sub_total_saida-vl_sub_total_entrada) < 0 then s_cor="red"
				s_class = ""
				if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then s_class = " MC"
				x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
						"		<td class='MB ME" & s_class & "' colspan='2' nowrap><p class='Cd' style='color:" & s_cor & ";'>" & _
						"TOTAL:</p></td>" & chr(13) & _
						"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_sub_total) & "</p></td>" & chr(13) & _
						"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida) & "</p></td>" & chr(13) & _
						"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_entrada) & "</p></td>" & chr(13) & _
						"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida-vl_sub_total_entrada) & "</p></td>" & chr(13) & _
						"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</p></td>" & chr(13) & _
						"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_valor_RT) & "</p></td>" & chr(13) & _
						"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc2) & "%" & "</p></td>" & chr(13) & _
						"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_lucro_liquido) & "</p></td>" & chr(13) & _
						"		<td class='MB MD" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc3) & "%" & "</p></td>" & chr(13) & _
						"	</tr>" & chr(13) & _
						"</table>" & chr(13)
				Response.Write x
				x="<br />" & chr(13)
				end if

			n_reg = 0
			vl_sub_total_saida = 0
			vl_sub_total_entrada = 0
			qtde_sub_total = 0
			vl_sub_total_valor_RT = 0
			vl_sub_total_lucro_liquido = 0
			
			grupo_a = "XXXXXXX"
			n_reg_grupo = 0
			vl_sub_total_saida_grupo = 0
			vl_sub_total_entrada_grupo = 0
			qtde_sub_total_grupo = 0
			vl_sub_total_valor_RT_grupo = 0
			vl_sub_total_lucro_liquido_grupo = 0

			if n_reg_total > 0 then x = x & "<br />" & chr(13)
			s = Trim("" & r("fabricante"))
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then x = x & "	<tr><td class='MDTE' colspan='11' valign='bottom' class='MB' style='background:azure;'><span class='N'>&nbsp;" & s & "</span></td></tr>" & chr(13)
			x = x & cab
			end if
		

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then
			x = x & "	<tr nowrap>" & chr(13)

			s_cor="black"
			if IsNumeric(r("qtde")) then if Clng(r("qtde")) < 0 then s_cor="red"

		 '> CÓDIGO DO PRODUTO
			x = x & "		<td class='MDTE tdCodProd' valign='bottom'><p class='Cn' style='color:" & s_cor & ";'>" & Trim("" & r("produto")) & "</p></td>" & chr(13)

		 '> DESCRIÇÃO DO PRODUTO
			s = Trim("" & r("descricao_html"))
			if s <> "" then s = produto_formata_descricao_em_html(s) else s = "&nbsp;"
			x = x & "		<td class='MTD tdDescrProd' valign='bottom'><p class='Cn' style='color:" & s_cor & ";'>" & s & "</p></td>" & chr(13)

		 '> QUANTIDADE
			x = x & "		<td align='right' valign='bottom' class='MTD tdQtde'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(r("qtde")) & "</p></td>" & chr(13)

		 '> VALOR SAÍDA
			x = x & "		<td align='right' valign='bottom' class='MTD tdFatTot'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_saida")) & "</p></td>" & chr(13)

		 '> VALOR ENTRADA
			x = x & "		<td align='right' valign='bottom' class='MTD tdCmvTot'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_entrada")) & "</p></td>" & chr(13)

		 '> LUCRO BRUTO
			x = x & "		<td align='right' valign='bottom' class='MTD tdLucro'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_saida")-r("valor_entrada")) & "</p></td>" & chr(13)

		 '> PERCENTUAL DO LUCRO BRUTO TOTAL
			if CCur(r("valor_saida")) = CCur(0) then
				perc = 0
			else
				perc = ((r("valor_saida")-r("valor_entrada"))/r("valor_saida"))*100
				end if
			x = x & "		<td align='right' valign='bottom' class='MTD tdPercLucroTot'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</p></td>" & chr(13)
		
		 '> VALOR COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
			x = x & "		<td align='right' class='MTD tdVlRt'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_RT")) & "</p></td>" & chr(13)
		
		 '> COMISSÃO (% DO FAT)
			if CCur(r("valor_saida")) = CCur(0) then
				perc2 = 0
			else
				perc2 = (r("valor_RT")/r("valor_saida"))*100
				end if
			x = x & "		<td align='right' class='MTD tdRtPercFat'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc2) & "%" & "</p></td>" & chr(13)
		
		 '> LUCRO LÍQUIDO (DESCONTADA A COMISSÃO)
			vl_lucro_liquido = r("valor_saida")-r("valor_entrada")-r("valor_RT")
			x = x & "		<td align='right' class='MTD tdLucroLiq'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lucro_liquido) & "</p></td>" & chr(13)
	 
		 '> PERCENTUAL DO LUCRO LÍQUIDO SOBRE O FATURAMENTO
			if CCur(r("valor_saida")) = CCur(0) then
				perc3 = 0
			else
				perc3 = (vl_lucro_liquido/r("valor_saida"))*100
				end if
			x = x & "		<td align='right' class='MTD tdLucroLiqPercFat'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc3) & "%" & "</p></td>" & chr(13)
		
			x = x & "	</tr>" & chr(13)

		elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
			if (Trim("" & r("grupo")) <> grupo_a) then
				if n_reg_grupo > 0 then
				  ' Fecha o grupo de produtos anterior
					if vl_sub_total_saida_grupo = 0 then
						perc = 0
					else
						perc = ((vl_sub_total_saida_grupo-vl_sub_total_entrada_grupo)/vl_sub_total_saida_grupo)*100
						end if
					
					if vl_sub_total_saida_grupo = 0 then
						perc2 = 0
					else
						perc2 = (vl_sub_total_valor_RT_grupo/vl_sub_total_saida_grupo)*100
						end if

					if vl_sub_total_saida_grupo = 0 then
						perc3 = 0
					else
						perc3 = (vl_sub_total_lucro_liquido_grupo/vl_sub_total_saida_grupo)*100
						end if

					s_cor="black"
					if qtde_sub_total_grupo < 0 then s_cor="red"
					if vl_sub_total_saida_grupo < 0 then s_cor="red"
					if vl_sub_total_entrada_grupo < 0 then s_cor="red"
					if (vl_sub_total_saida_grupo-vl_sub_total_entrada_grupo) < 0 then s_cor="red"
					x = x & "	<tr nowrap>" & chr(13) & _
							"		<td class='MC ME' colspan='2' nowrap><p class='Cd' style='color:" & s_cor & ";'>" & _
							"Total do Grupo:</p></td>" & chr(13) & _
							"		<td class='MC'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_sub_total_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MC'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MC'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_entrada_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MC'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida_grupo-vl_sub_total_entrada_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MC'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_valor_RT_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MC'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc2) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MC'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_lucro_liquido_grupo) & "</p></td>" & chr(13) & _
							"		<td class='MC MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc3) & "%" & "</p></td>" & chr(13) & _
							"	</tr>" & chr(13)
					Response.Write x
					x="<br />" & chr(13)
					end if
				
				grupo_a = Trim("" & r("grupo"))
				n_reg_grupo = 0
				vl_sub_total_saida_grupo = 0
				vl_sub_total_entrada_grupo = 0
				qtde_sub_total_grupo = 0
				vl_sub_total_valor_RT_grupo = 0
				vl_sub_total_lucro_liquido_grupo = 0
				
				x = x & "	<tr><td class='MDTE' colspan='11' valign='bottom' class='MB' style='background:#EEE;'><span class='N'>&nbsp;" & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</span></td></tr>" & chr(13)
				end if

			n_reg_grupo = n_reg_grupo + 1

			x = x & "	<tr nowrap>" & chr(13)

			s_cor="black"
			if IsNumeric(r("qtde")) then if Clng(r("qtde")) < 0 then s_cor="red"

		 '> CÓDIGO DO PRODUTO
			x = x & "		<td class='MDTE tdCodProd' valign='bottom'><p class='Cn' style='color:" & s_cor & ";'>" & Trim("" & r("produto")) & "</p></td>" & chr(13)

		 '> DESCRIÇÃO DO PRODUTO
			s = Trim("" & r("descricao_html"))
			if s <> "" then s = produto_formata_descricao_em_html(s) else s = "&nbsp;"
			x = x & "		<td class='MTD tdDescrProd' valign='bottom'><p class='Cn' style='color:" & s_cor & ";'>" & s & "</p></td>" & chr(13)

		 '> QUANTIDADE
			x = x & "		<td align='right' valign='bottom' class='MTD tdQtde'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(r("qtde")) & "</p></td>" & chr(13)

		 '> VALOR SAÍDA
			x = x & "		<td align='right' valign='bottom' class='MTD tdFatTot'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_saida")) & "</p></td>" & chr(13)

		 '> VALOR ENTRADA
			x = x & "		<td align='right' valign='bottom' class='MTD tdCmvTot'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_entrada")) & "</p></td>" & chr(13)

		 '> LUCRO BRUTO
			x = x & "		<td align='right' valign='bottom' class='MTD tdLucro'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_saida")-r("valor_entrada")) & "</p></td>" & chr(13)

		 '> PERCENTUAL DO LUCRO BRUTO TOTAL
			if Ccur(r("valor_saida")) = Ccur(0) then
				perc = 0
			else
				perc = ((r("valor_saida")-r("valor_entrada"))/r("valor_saida"))*100
				end if
			x = x & "		<td align='right' valign='bottom' class='MTD tdPercLucroTot'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</p></td>" & chr(13)
		
		 '> VALOR COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
			x = x & "		<td align='right' class='MTD tdVlRt'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_RT")) & "</p></td>" & chr(13)
		
		 '> COMISSÃO (% DO FAT)
			if Ccur(r("valor_saida")) = Ccur(0) then
				perc2 = 0
			else
				perc2 = (r("valor_RT")/r("valor_saida"))*100
				end if
			x = x & "		<td align='right' class='MTD tdRtPercFat'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc2) & "%" & "</p></td>" & chr(13)
		
		 '> LUCRO LÍQUIDO (DESCONTADA A COMISSÃO)
			vl_lucro_liquido = r("valor_saida")-r("valor_entrada")-r("valor_RT")
			x = x & "		<td align='right' class='MTD tdLucroLiq'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lucro_liquido) & "</p></td>" & chr(13)
	 
		 '> PERCENTUAL DO LUCRO LÍQUIDO SOBRE O FATURAMENTO
			if CCur(r("valor_saida")) = CCur(0) then
				perc3 = 0
			else
				perc3 = (vl_lucro_liquido/r("valor_saida"))*100
				end if
			x = x & "		<td align='right' class='MTD tdLucroLiqPercFat'><p class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc3) & "%" & "</p></td>" & chr(13)
		
			x = x & "	</tr>" & chr(13)

		'	TOTALIZAÇÃO GERAL POR GRUPO
			idxSelecionado = LBound(vTotGrupo) - 1
			for j=LBound(vTotGrupo) to UBound(vTotGrupo)
				if Trim("" & vTotGrupo(j).codigo_grupo) = Trim("" & r("grupo")) then
					idxSelecionado = j
					exit for
					end if
				next

			if idxSelecionado = (LBound(vTotGrupo) - 1) then
				redim preserve vTotGrupo(UBound(vTotGrupo)+1)
				set vTotGrupo(UBound(vTotGrupo)) = new cl_REL_FAT3_TOT_GRUPO
				inicializa_cl_REL_FAT3_TOT_GRUPO vTotGrupo(UBound(vTotGrupo))
				idxSelecionado=UBound(vTotGrupo)
				vTotGrupo(idxSelecionado).codigo_grupo = Trim("" & r("grupo"))
				vTotGrupo(idxSelecionado).descricao_grupo = Trim("" & r("grupo_descricao"))
				end if

			vTotGrupo(idxSelecionado).ha_dados = True
			vTotGrupo(idxSelecionado).qtde = vTotGrupo(idxSelecionado).qtde + r("qtde")
			vTotGrupo(idxSelecionado).vl_saida = vTotGrupo(idxSelecionado).vl_saida + r("valor_saida")
			vTotGrupo(idxSelecionado).vl_entrada = vTotGrupo(idxSelecionado).vl_entrada + r("valor_entrada")
			vTotGrupo(idxSelecionado).vl_lucro_bruto = vTotGrupo(idxSelecionado).vl_lucro_bruto + (r("valor_saida")-r("valor_entrada"))
			vTotGrupo(idxSelecionado).vl_RT = vTotGrupo(idxSelecionado).vl_RT + r("valor_RT")
			vTotGrupo(idxSelecionado).vl_lucro_liq = vTotGrupo(idxSelecionado).vl_lucro_liq + (r("valor_saida")-r("valor_entrada")-r("valor_RT"))
			end if

		qtde_sub_total = qtde_sub_total + r("qtde")
		qtde_sub_total_grupo = qtde_sub_total_grupo + r("qtde")
		qtde_total = qtde_total + r("qtde")
		vl_sub_total_saida = vl_sub_total_saida + r("valor_saida")
		vl_sub_total_saida_grupo = vl_sub_total_saida_grupo + r("valor_saida")
		vl_sub_total_entrada = vl_sub_total_entrada + r("valor_entrada")
		vl_sub_total_entrada_grupo = vl_sub_total_entrada_grupo + r("valor_entrada")
		vl_sub_total_valor_RT = vl_sub_total_valor_RT + r("valor_RT")
		vl_sub_total_valor_RT_grupo = vl_sub_total_valor_RT_grupo + r("valor_RT")
		vl_sub_total_lucro_liquido = vl_sub_total_lucro_liquido + vl_lucro_liquido
		vl_sub_total_lucro_liquido_grupo = vl_sub_total_lucro_liquido_grupo + vl_lucro_liquido
		
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
	
'	TOTALIZAÇÃO GERAL POR GRUPO DE PRODUTOS: CALCULA DADOS FINAIS
	if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
		for j=LBound(vTotGrupo) to UBound(vTotGrupo)
			if vTotGrupo(j).ha_dados then
			'	PERCENTUAL DO LUCRO BRUTO TOTAL
				if CCur(vTotGrupo(j).vl_saida) = CCur(0) then
					perc = 0
				else
					perc = ((vTotGrupo(j).vl_saida - vTotGrupo(j).vl_entrada)/vTotGrupo(j).vl_saida)*100
					end if
				vTotGrupo(j).perc_lucro_bruto_total = perc
				
			'	COMISSÃO (% DO FAT)
				if CCur(vTotGrupo(j).vl_saida) = Ccur(0) then
					perc = 0
				else
					perc = (vTotGrupo(j).vl_RT/vTotGrupo(j).vl_saida)*100
					end if
				vTotGrupo(j).perc_comissao_fat = perc
				
			'	PERCENTUAL DO LUCRO LÍQUIDO SOBRE O FATURAMENTO
				if CCur(vTotGrupo(j).vl_saida) = CCur(0) then
					perc = 0
				else
					perc = (vTotGrupo(j).vl_lucro_liq/vTotGrupo(j).vl_saida)*100
					end if
				vTotGrupo(j).perc_lucro_liq_fat = perc
				end if
			next

		redim vOrdGrupo(UBound(vTotGrupo))
		for j=LBound(vOrdGrupo) to UBound(vOrdGrupo)
			set vOrdGrupo(j) = new cl_VINTE_COLUNAS
			vOrdGrupo(j).CampoOrdenacao = vTotGrupo(j).codigo_grupo
			vOrdGrupo(j).c1 = vTotGrupo(j).ha_dados
			vOrdGrupo(j).c2 = vTotGrupo(j).codigo_grupo
			vOrdGrupo(j).c3 = vTotGrupo(j).descricao_grupo
			vOrdGrupo(j).c4 = vTotGrupo(j).qtde
			vOrdGrupo(j).c5 = vTotGrupo(j).vl_saida
			vOrdGrupo(j).c6 = vTotGrupo(j).vl_entrada
			vOrdGrupo(j).c7 = vTotGrupo(j).vl_lucro_bruto
			vOrdGrupo(j).c8 = vTotGrupo(j).perc_lucro_bruto_total
			vOrdGrupo(j).c9 = vTotGrupo(j).vl_RT
			vOrdGrupo(j).c10 = vTotGrupo(j).perc_comissao_fat
			vOrdGrupo(j).c11 = vTotGrupo(j).vl_lucro_liq
			vOrdGrupo(j).c12 = vTotGrupo(j).perc_lucro_liq_fat
			next

		ordena_cl_vinte_colunas vOrdGrupo, 0, UBound(vOrdGrupo)

		redim vTotGrupo(0)
		redim vTotGrupo(ubound(vOrdGrupo))
		for j=LBound(vTotGrupo) to UBound(vTotGrupo)
			set vTotGrupo(j) = new cl_REL_FAT3_TOT_GRUPO
			vTotGrupo(j).ha_dados = vOrdGrupo(j).c1
			vTotGrupo(j).codigo_grupo = vOrdGrupo(j).c2
			vTotGrupo(j).descricao_grupo = vOrdGrupo(j).c3
			vTotGrupo(j).qtde = vOrdGrupo(j).c4
			vTotGrupo(j).vl_saida = vOrdGrupo(j).c5
			vTotGrupo(j).vl_entrada = vOrdGrupo(j).c6
			vTotGrupo(j).vl_lucro_bruto = vOrdGrupo(j).c7
			vTotGrupo(j).perc_lucro_bruto_total = vOrdGrupo(j).c8
			vTotGrupo(j).vl_RT = vOrdGrupo(j).c9
			vTotGrupo(j).perc_comissao_fat = vOrdGrupo(j).c10
			vTotGrupo(j).vl_lucro_liq = vOrdGrupo(j).c11
			vTotGrupo(j).perc_lucro_liq_fat = vOrdGrupo(j).c12
			next
		end if

  ' MOSTRA TOTAL DO ÚLTIMO GRUPO E FABRICANTE
	if n_reg <> 0 then 
	  ' MOSTRA TOTAL DO ÚLTIMO GRUPO
		if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
			if vl_sub_total_saida_grupo = 0 then
				perc = 0
			else
				perc = ((vl_sub_total_saida_grupo-vl_sub_total_entrada_grupo)/vl_sub_total_saida_grupo)*100
				end if
					
			if vl_sub_total_saida_grupo = 0 then
				perc2 = 0
			else
				perc2 = (vl_sub_total_valor_RT_grupo/vl_sub_total_saida_grupo)*100
				end if

			if vl_sub_total_saida_grupo = 0 then
				perc3 = 0
			else
				perc3 = (vl_sub_total_lucro_liquido_grupo/vl_sub_total_saida_grupo)*100
				end if

			s_cor="black"
			if qtde_sub_total_grupo < 0 then s_cor="red"
			if vl_sub_total_saida_grupo < 0 then s_cor="red"
			if vl_sub_total_entrada_grupo < 0 then s_cor="red"
			if (vl_sub_total_saida_grupo-vl_sub_total_entrada_grupo) < 0 then s_cor="red"
			x = x & "	<tr nowrap>" & chr(13) & _
					"		<td class='MTBE' colspan='2' nowrap><p class='Cd' style='color:" & s_cor & ";'>" & _
					"Total do Grupo:</p></td>" & chr(13) & _
					"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_sub_total_grupo) & "</p></td>" & chr(13) & _
					"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida_grupo) & "</p></td>" & chr(13) & _
					"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_entrada_grupo) & "</p></td>" & chr(13) & _
					"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida_grupo-vl_sub_total_entrada_grupo) & "</p></td>" & chr(13) & _
					"		<td class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</p></td>" & chr(13) & _
					"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_valor_RT_grupo) & "</p></td>" & chr(13) & _
					"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc2) & "%" & "</p></td>" & chr(13) & _
					"		<td class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_lucro_liquido_grupo) & "</p></td>" & chr(13) & _
					"		<td class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc3) & "%" & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
			Response.Write x
			x="<br />" & chr(13)
			end if
		
	  ' MOSTRA TOTAL DO ÚLTIMO FABRICANTE
		if vl_sub_total_saida = 0 then
			perc = 0
		else
			perc = ((vl_sub_total_saida-vl_sub_total_entrada)/vl_sub_total_saida)*100
			end if
		
		if vl_sub_total_saida = 0 then
			perc2 = 0
		else
			perc2 = (vl_sub_total_valor_RT/vl_sub_total_saida)*100
			end if
		
		if vl_sub_total_saida = 0 then
			perc3 = 0
		else
			perc3 = (vl_sub_total_lucro_liquido/vl_sub_total_saida)*100
			end if
			
		s_cor="black"
		if qtde_sub_total < 0 then s_cor="red"
		if vl_sub_total_saida < 0 then s_cor="red"
		if vl_sub_total_entrada < 0 then s_cor="red"
		if (vl_sub_total_saida-vl_sub_total_entrada) < 0 then s_cor="red"
		s_class = ""
		if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then s_class = " MC"
		x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
				"		<td colspan='2' class='MB ME" & s_class & "' nowrap><p class='Cd' style='color:" & s_cor & ";'>" & _
				"TOTAL:</p></td>" & chr(13) & _
				"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_sub_total) & "</p></td>" & chr(13) & _
				"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida) & "</p></td>" & chr(13) & _
				"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_entrada) & "</p></td>" & chr(13) & _
				"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida-vl_sub_total_entrada) & "</p></td>" & chr(13) & _
				"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</p></td>" & chr(13) & _
				"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_valor_RT) & "</p></td>" & chr(13) & _
				"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc2) & "%" & "</p></td>" & chr(13) & _
				"		<td class='MB" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_lucro_liquido) & "</p></td>" & chr(13) & _
				"		<td class='MB MD" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc3) & "%" & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)
		
	'>	TOTAL GERAL
		if qtde_fabricantes > 1 then
			if vl_total_saida = 0 then
				perc = 0
			else
				perc = ((vl_total_saida-vl_total_entrada)/vl_total_saida)*100
				end if
			
			if vl_total_saida = 0 then
				perc2 = 0
				perc3 = 0
			else
				perc2 = (vl_total_valor_RT/vl_total_saida)*100
				perc3 = (vl_total_lucro_liquido/vl_total_saida)*100
				end if
			
			s_cor="black"
			if qtde_total < 0 then s_cor="red"
			if vl_total_saida < 0 then s_cor="red"
			if vl_total_entrada < 0 then s_cor="red"
			if (vl_total_saida-vl_total_entrada) < 0 then s_cor="red"
			x = x & "	<tr><td colspan='11' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
					"	<tr><td colspan='11' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
					"	<tr nowrap style='background:honeydew'>" & chr(13) & _
					"		<td class='MC MB ME' colspan='2' nowrap><p class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL:</p></td>" & chr(13) & _
					"		<td class='MC MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_total) & "</p></td>" & chr(13) & _
					"		<td class='MC MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_saida) & "</p></td>" & chr(13) & _
					"		<td class='MC MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_entrada) & "</p></td>" & chr(13) & _
					"		<td class='MC MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_saida-vl_total_entrada) & "</p></td>" & chr(13) & _
					"		<td class='MC MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</p></td>" & chr(13) & _
					"		<td class='MC MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_valor_RT) & "</p></td>" & chr(13) & _
					"		<td class='MC MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc2) & "%" & "</p></td>" & chr(13) & _
					"		<td class='MC MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_lucro_liquido) & "</p></td>" & chr(13) & _
					"		<td class='MC MB MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc3) & "%" & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
			end if
		end if

	if n_reg_total <> 0 then
		if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
			linha_tot_grupo = 0
			for j=LBound(vTotGrupo) to UBound(vTotGrupo)
				if vTotGrupo(j).ha_dados then
					linha_tot_grupo = linha_tot_grupo + 1
					s_cor="black"
					if vTotGrupo(j).qtde < 0 then s_cor="red"
					if vTotGrupo(j).vl_saida < 0 then s_cor="red"
					if vTotGrupo(j).vl_entrada < 0 then s_cor="red"
					if (vTotGrupo(j).vl_saida-vTotGrupo(j).vl_entrada) < 0 then s_cor="red"
					if linha_tot_grupo = 1 then
						x = x & "	<tr><td colspan='11' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
								"	<tr><td colspan='11' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
								"	<tr style='background:azure'><td class='MC MD ME' colspan='11'><p class='R' style='font-weight:bold;'>TOTAL GERAL POR GRUPO DE PRODUTOS</p></td></tr>" & chr(13)
						end if
					s_class = ""
					if j = UBound(vTotGrupo) then s_class = " MB"
					s_codigo_grupo = Trim("" & vTotGrupo(j).codigo_grupo)
					s_descricao_grupo = Trim("" & vTotGrupo(j).descricao_grupo)
					if s_codigo_grupo = "" then
						s_codigo_grupo = "-"
						if s_descricao_grupo = "" then s_descricao_grupo = "-"
						end if
					x = x & "	<tr nowrap style='background:honeydew'>" & chr(13) & _
							"		<td class='MC MD ME" & s_class & "' nowrap><p class='Cd' style='color:" & s_cor & ";'>" & s_codigo_grupo & "</p></td>" & chr(13) & _
							"		<td class='MC MD" & s_class & "'><p class='C' style='color:" & s_cor & ";'>" & s_descricao_grupo & "</p></td>" & chr(13) & _
							"		<td class='MC MD" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(vTotGrupo(j).qtde) & "</p></td>" & chr(13) & _
							"		<td class='MC MD" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(j).vl_saida) & "</p></td>" & chr(13) & _
							"		<td class='MC MD" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(j).vl_entrada) & "</p></td>" & chr(13) & _
							"		<td class='MC MD" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(j).vl_saida-vTotGrupo(j).vl_entrada) & "</p></td>" & chr(13) & _
							"		<td class='MC MD" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(vTotGrupo(j).perc_lucro_bruto_total) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MC MD" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(j).vl_RT) & "</p></td>" & chr(13) & _
							"		<td class='MC MD" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(vTotGrupo(j).perc_comissao_fat) & "%" & "</p></td>" & chr(13) & _
							"		<td class='MC MD" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vTotGrupo(j).vl_lucro_liq) & "</p></td>" & chr(13) & _
							"		<td class='MC MD" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(vTotGrupo(j).perc_lucro_liq_fat) & "%" & "</p></td>" & chr(13) & _
							"	</tr>" & chr(13)
					end if
				next
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = ""
		if c_fabricante <> "" then
			s = c_fabricante
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			if s <> "" then x = x & cab_table & "	<tr><td class='MDTE' colspan='11' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td></tr>" & chr(13) & cab
		else
			x = x & cab_table & cab
			end if

		x = x & "	<tr nowrap>" & chr(13) & _
				"		<td class='MT ALERTA' colspan='11' align='center'><span class='ALERTA'>&nbsp;NENHUM PRODUTO SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO FABRICANTE
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



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style type="text/css">
.tdCodProd{
	vertical-align: top;
	width: 50px;
	}
.tdDescrProd{
	vertical-align: top;
	width: 190px;
	}
.tdQtde{
	vertical-align: top;
	width: 40px;
	}
.tdFatTot{
	vertical-align: top;
	width: 80px;
	}
.tdCmvTot{
	vertical-align: top;
	width: 80px;
	}
.tdLucro{
	vertical-align: top;
	width: 80px;
	}
.tdPercLucroTot{
	vertical-align: top;
	width: 50px;
	}
.tdVlRt{
	vertical-align: top;
	width: 80px;
	}
.tdRtPercFat{
	vertical-align: top;
	width: 50px;
	}
.tdLucroLiq{
	vertical-align: top;
	width: 80px;
	}
.tdLucroLiqPercFat{
	vertical-align: top;
	width: 50px;
	}
</style>


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br />
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br /><br />
<p class="TracoBottom"></p>
<table cellSpacing="0">
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
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>" />
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>" />
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>" />
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>" />
<input type="hidden" name="c_grupo" id="c_grupo" value="<%=c_grupo%>" />
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>" />
<input type="hidden" name="c_pedido" id="c_pedido" value="<%=c_pedido%>" />
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>" />
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>" />
<input type="hidden" name="rb_tipo_agrupamento" id="rb_tipo_agrupamento" value="<%=rb_tipo_agrupamento%>" />
<input type="hidden" name="c_cst" id="c_cst" value="<%=c_cst%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="864" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Faturamento III</span>
	<br /><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='864' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>"
	
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Período:&nbsp;</span></td><td valign='top' width='99%'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	s_aux = ""
	if c_fabricante <> "" then s_aux = x_fabricante(c_fabricante)
	s = c_fabricante
	if (s <> "") And (s_aux <> "") then s = s & " - " & s_aux
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Fabricante:&nbsp;</span></td><td valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	s = c_produto
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Produto:&nbsp;</span></td><td valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	s = c_grupo
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Grupo de Produtos:&nbsp;</span></td><td valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

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

	s = c_vendedor
	if s = "" then 
		s = "todos"
	else
		if (s_nome_vendedor <> "") And (s_nome_vendedor <> c_vendedor) then s = s & " (" & s_nome_vendedor & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Vendedor:&nbsp;</span></td><td valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	s = c_indicador
	if s = "" then 
		s = "todos"
	else
		s = s & " (" & x_orcamentista_e_indicador(c_indicador) & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Indicador:&nbsp;</span></td><td valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	s = c_pedido
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Pedido:&nbsp;</span></td><td valign='top'>" & _
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

	s = x_opcao_forma_pagamento(op_forma_pagto)
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Forma de Pagamento:&nbsp;</span></td><td valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	if c_forma_pagto_qtde_parc <> "" then
		s = c_forma_pagto_qtde_parc
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Nº Parcelas:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = rb_tipo_cliente
	if s = "" then 
		s = "todos"
	elseif s = ID_PF then
		s = "Pessoa Física"
	elseif s = ID_PJ then
		s = "Pessoa Jurídica"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Tipo de Cliente:&nbsp;</span></td><td valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = c_cliente_cnpj_cpf
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>CNPJ/CPF Cliente:&nbsp;</span></td><td valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	s = c_uf_pesq
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>UF:&nbsp;</span></td><td valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = c_empresa
	if s = "" then 
		s = "todas"
	else
		s =  obtem_apelido_empresa_NFe_emitente(c_empresa)
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
			   "<span class='N'>Loja(s):&nbsp;</span></td><td valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
		s = "Grupo de Produtos"
	elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then
		s = "Produto"
	else
		s = "N.I."
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Tipo de Agrupamento:&nbsp;</span></td><td valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br />
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="864" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br />


<table class="notPrint" width="864" cellspacing="0">
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
