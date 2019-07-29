<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelCompras2Exec.asp
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
	if (Not operacao_permitida(OP_CEN_REL_COMPRAS2, s_lista_operacoes_permitidas)) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, c_fabricante, c_produto, c_dt_inicio, c_dt_termino, rb_detalhe
	dim cod_fabricante, cod_produto
	dim s_nome_fabricante, s_nome_produto, s_nome_produto_html,c_grupo,c_potencia_BTU,c_ciclo,c_posicao_mercado,v_fabricantes,cont
    dim s_where_temp,v_grupos

	c_produto = UCase(Trim(Request.Form("c_produto")))
	rb_detalhe = Trim(Request.Form("rb_detalhe"))
	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_fabricante = Trim(Request.Form("c_fabricante"))
	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_potencia_BTU = Trim(Request.Form("c_potencia_BTU"))
	c_ciclo = Trim(Request.Form("c_ciclo"))
	c_posicao_mercado = Trim(Request.Form("c_posicao_mercado"))

	alerta = ""
	if (c_produto<>"") And (Not IsEAN(c_produto)) then
		if c_fabricante = "" then alerta = "PARA CONSULTAR PELO CÓDIGO INTERNO DE PRODUTO É NECESSÁRIO ESPECIFICAR O FABRICANTE."
		end if
		
	if alerta = "" then
	'	DEFAULT
		cod_produto = c_produto
		cod_fabricante = c_fabricante
		
		if IsEAN(c_produto) then
			s = "SELECT fabricante, produto, ean FROM t_PRODUTO WHERE (ean='" & c_produto & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta = "Produto com código EAN " & c_produto & " não está cadastrado."
			else
				if c_fabricante <> "" then
					if c_fabricante <> Trim("" & rs("fabricante")) then 
						alerta = "Produto " & Trim("" & rs("produto")) & " (EAN: " & _
								 Trim("" & rs("ean")) & ") não pertence ao fabricante " & c_fabricante & "."
						end if
					end if
				
				if alerta = "" then
				'	OBTÉM O CÓDIGO INTERNO DE PRODUTO
					cod_fabricante = Trim("" & rs("fabricante"))
					cod_produto = Trim("" & rs("produto"))
					end if
				end if
			end if
		end if

    if alerta = "" then
		if c_fabricante <> "" then
		    s = "SELECT nome, razao_social FROM t_FABRICANTE WHERE "
		    v_fabricantes = split(c_fabricante, ", ")
		    for cont = LBound(v_fabricantes) to UBound(v_fabricantes)
                if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		        s_where_temp = s_where_temp & " (fabricante = '" & v_fabricantes(cont) & "')"
            next
            s = s & s_where_temp
			set rs = cn.Execute(s)
			if rs.Eof then
				alerta = "FABRICANTE '" & c_fabricante & "' NÃO ESTÁ CADASTRADO"
			else
				s_nome_fabricante = Trim("" & rs("razao_social"))
				if s_nome_fabricante = "" then s_nome_fabricante = Trim("" & rs("nome"))
				end if
			end if
		end if
	if alerta = "" then
		if cod_fabricante <> "" then
			s_nome_fabricante = fabricante_descricao(cod_fabricante)
		else
			s_nome_fabricante = ""
			end if
				
		if cod_produto <> "" then
			s_nome_produto = produto_descricao(cod_fabricante, cod_produto)
			s_nome_produto_html = produto_formata_descricao_em_html(produto_descricao_html(cod_fabricante, cod_produto))
		else
			s_nome_produto = ""
			s_nome_produto_html = ""
			end if
		end if
		
	if alerta = "" then
		if c_dt_inicio = "" then
			alerta="A data de início do período não foi informada."
		elseif Not IsDate(c_dt_inicio) then
			alerta="A data de início do período é inválida (" & c_dt_inicio & ")."
		elseif c_dt_termino = "" then
			alerta="A data de término do período não foi informada."
		elseif Not IsDate(c_dt_termino) then
			alerta="A data de término do período é inválida (" & c_dt_termino & ")."
		elseif CDate(c_dt_inicio) > CDate(c_dt_termino) then
			alerta="A data de início é posterior à data de término."
			end if
		end if

	if alerta = "" then
	'	Período de consulta está restrito por perfil de acesso?
		dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
		dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
		dim strDtRefDDMMYYYY
		if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
			intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
			dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
			strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
			strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
		'	PERÍODO
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
		end if

    if alerta = "" then
		call set_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_dt_inicio", c_dt_inicio)
		call set_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_dt_termino", c_dt_termino)
		call set_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_fabricante", c_fabricante)
		call set_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_grupo", c_grupo)
		call set_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_potencia_BTU", c_potencia_BTU)
		call set_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_ciclo", c_ciclo)
		call set_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_posicao_mercado", c_posicao_mercado)
        call set_default_valor_texto_bd(usuario, "RelCompras2Filtro|rb_detalhe", rb_detalhe)
		end if



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA_DETALHE_SINTETICO_FABRICANTE
'
sub consulta_detalhe_sintetico_fabricante
dim r
dim valor_total, s_sql, cab, n_reg, x

	s_sql = "SELECT" & _
				" e.fabricante," & _
				" Sum(qtde* i.vl_custo2) AS valor" & _
			" FROM t_ESTOQUE e " & _
            "INNER JOIN t_ESTOQUE_ITEM i ON (e.id_estoque=i.id_estoque)" & _
            "INNER JOIN t_PRODUTO p ON (  i.fabricante = p.fabricante ) AND (i.produto= p.produto)" & _
			" WHERE" & _
				" (kit=0)" & _
				" AND (entrada_especial=0)" & _
				" AND (devolucao_status=0)"
                
	if IsDate(c_dt_inicio) then
		s_sql = s_sql & " AND (data_entrada >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		s_sql = s_sql & " AND (data_entrada < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if

    s_where_temp = ""
	if c_fabricante <> "" then
	for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	    if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		s_where_temp = s_where_temp & _
			   "  (i.fabricante = '" & v_fabricantes(cont) & "')"
	next
	s_sql = s_sql & "AND"
	s_sql =  s_sql & "(" & s_where_temp & ")"
    end if

    s_where_temp = ""
	if c_grupo <> "" then
	v_grupos = split(c_grupo, ", ")
	for cont = Lbound(v_grupos) to Ubound(v_grupos)
	    if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		s_where_temp = s_where_temp & _
			" (grupo = '" & v_grupos(cont) & "')"
	next
	s_sql = s_sql & "AND "
	s_sql = s_sql & "(" & s_where_temp & ")"
		end if

	if Trim(cod_produto) <> "" then
		s_sql = s_sql & " AND (i.produto = '" & cod_produto & "')"
		end if

    if Trim(c_potencia_BTU) <> "" then
        s_sql = s_sql & " AND (potencia_BTU = '" & c_potencia_BTU & "')"
        end if

    if Trim(c_ciclo) <> "" then
        s_sql = s_sql & " AND (ciclo = '" & c_ciclo & "')"
        end if

    if Trim(c_posicao_mercado) <> "" then
        s_sql = s_sql & " AND (posicao_mercado = '" & c_posicao_mercado & "')"
        end if
		
	s_sql = s_sql & " GROUP BY e.fabricante" & _
					" ORDER BY e.fabricante"

  ' CABEÇALHO
	cab = "<TABLE class='Q' cellSpacing=0>" & chr(13) & _
		  "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD valign='bottom' NOWRAP class='MD MB'><P style='width:380px' class='R'>FORNECEDOR</P></TD>" & chr(13) & _
		  "		<TD align='right' valign='bottom' NOWRAP class='MB'><P style='width:110px' class='Rd' style='font-weight:bold;'>VALOR</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab
	n_reg = 0
	valor_total = 0
		
	set r = cn.execute(s_sql)
	do while Not r.Eof
	  ' CONTAGEM
		n_reg = n_reg + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> FORNECEDOR
		x = x & "		<TD class='MDB'><P class='C'>&nbsp;" & Trim("" & r("fabricante")) & " - " & iniciais_em_maiusculas(x_fabricante(Trim("" & r("fabricante")))) & "</P></TD>" & chr(13)

	 '> VALOR
		x = x & "		<TD align='right' class='MB'><P class='Cd'>&nbsp;" & formata_moeda(r("valor")) & "</P></TD>" & chr(13)

		valor_total = valor_total + r("valor")
		
		x = x & "	</TR>" & chr(13)
			
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
	
  ' MOSTRA VALOR TOTAL
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP style='background:ivory;'>" & chr(13) & _
				"		<TD COLSPAN='2' NOWRAP><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & formata_moeda(valor_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='2'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	
end sub



' _____________________________________
' CONSULTA_DETALHE_SINTETICO_PRODUTO
'
sub consulta_detalhe_sintetico_produto
dim r
dim valor_total, vl_sub_total, s_sql, cab_table, cab, n_reg, x
dim intQtde, intQtdeTotal, intQtdeSubTotal
dim strFabricanteAnterior, strFabricante, strProduto, intQtdeFabricantes

	s_sql = "SELECT" & _
				" i.fabricante," & _
				" i.produto," & _
				" Coalesce(Sum(qtde),0) AS qtde," & _
				" Coalesce(Sum(qtde* i.vl_custo2),0) AS valor" & _
			" FROM t_ESTOQUE e INNER JOIN t_ESTOQUE_ITEM i ON (e.id_estoque=i.id_estoque)" & _
            " INNER JOIN t_PRODUTO p ON (  i.fabricante = p.fabricante ) AND (i.produto= p.produto)" & _
			" WHERE" & _
				" (kit=0)" & _
				" AND (entrada_especial=0)" & _
				" AND (devolucao_status=0)"

	if IsDate(c_dt_inicio) then
		s_sql = s_sql & " AND (data_entrada >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
	
	if IsDate(c_dt_termino) then
		s_sql = s_sql & " AND (data_entrada < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
	s_where_temp = ""
	if c_fabricante <> "" then
	for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	    if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		s_where_temp = s_where_temp & _
			   "  (i.fabricante = '" & v_fabricantes(cont) & "')"
	next
	s_sql = s_sql & "AND"
	s_sql =  s_sql & "(" & s_where_temp & ")"
    end if

     s_where_temp = ""
	if c_grupo <> "" then
	v_grupos = split(c_grupo, ", ")
	for cont = Lbound(v_grupos) to Ubound(v_grupos)
	    if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		s_where_temp = s_where_temp & _
			" (grupo = '" & v_grupos(cont) & "')"
	next
	s_sql = s_sql & "AND "
	s_sql = s_sql & "(" & s_where_temp & ")"
		end if

	if Trim(cod_produto) <> "" then
		s_sql = s_sql & " AND (i.produto = '" & cod_produto & "')"
		end if

    if Trim(c_potencia_BTU) <> "" then
        s_sql = s_sql & " AND (potencia_BTU = '" & c_potencia_BTU & "')"
        end if

    if Trim(c_ciclo) <> "" then
        s_sql = s_sql & " AND (ciclo = '" & c_ciclo & "')"
        end if

    if Trim(c_posicao_mercado) <> "" then
        s_sql = s_sql & " AND (posicao_mercado = '" & c_posicao_mercado & "')"
        end if
		
	s_sql = s_sql & " GROUP BY i.fabricante, i.produto" & _
					" ORDER BY i.fabricante, i.produto"

  ' CABEÇALHO
	cab_table = "<TABLE class='MB' cellSpacing=0>" & chr(13)
	
	cab = _
		  "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD valign='bottom' class='MDTE' NOWRAP><P style='width:300px' class='R'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD valign='bottom' class='MTD' NOWRAP ><P style='width:60px' class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table
	n_reg = 0
	valor_total = 0
	intQtdeTotal = 0
	intQtdeFabricantes = 0
	
	strFabricanteAnterior = "XXXXXXXXXXXXXXXXXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

		strFabricante = Trim("" & r("fabricante"))
		strProduto = Trim("" & r("produto"))

		if strFabricante <> strFabricanteAnterior then
			intQtdeFabricantes = intQtdeFabricantes + 1
		'	SUB-TOTAIS POR FABRICANTE
		'	EXIBE SUB-TOTAL DO FABRICANTE ANTERIOR?
			if n_reg > 0 then
				x = x & _
					"	<TR NOWRAP style='background:ivory;'>" & chr(13) & _
					"		<TD class='MDTE' NOWRAP><P class='Cd'>" & "TOTAL:" & "</P></TD>" & chr(13) & _
					"		<TD class='MTD'><P class='Cd'>" & formata_inteiro(intQtdeSubTotal) & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD colspan='2' class='MC'>&nbsp;</TD>" & chr(13) & _
					"	</TR>" & chr(13)
				end if

			x = x & _
				"	<TR style='background:azure'>" & chr(13) & _
				"		<TD colspan='2' class='MC ME MD'><P class='C'>" & strFabricante & " - " & iniciais_em_maiusculas(x_fabricante(strFabricante)) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			
		'	TÍTULO DAS COLUNAS
			x = x & _
				cab

			vl_sub_total = 0
			intQtdeSubTotal = 0
			strFabricanteAnterior = strFabricante
			end if
			
			
	  ' CONTAGEM
		n_reg = n_reg + 1
		
		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDTE' valign='bottom'><P class='C'>" & strProduto & " - " & produto_formata_descricao_em_html(produto_descricao_html(strFabricante, strProduto)) & "</P></TD>" & chr(13)

	 '> QTDE
		intQtde = r("qtde")
		intQtdeTotal = intQtdeTotal + intQtde
		intQtdeSubTotal = intQtdeSubTotal + intQtde
		x = x & "		<TD class='MTD' valign='bottom'><P class='Cd'>&nbsp;" & formata_inteiro(intQtde) & "</P></TD>" & chr(13)

	 '> VALOR
	 '	NÃO EXIBIR O VALOR
		valor_total = valor_total + r("valor")
		vl_sub_total = vl_sub_total + r("valor")
		
		x = x & "	</TR>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		r.movenext
		loop
		
  ' MOSTRA VALOR TOTAL
	if n_reg <> 0 then
	'	SUB-TOTAL DO ÚLTIMO FABRICANTE 
		x = x & "	<TR NOWRAP style='background:ivory;'>" & chr(13) & _
				"		<TD class='MDTE' NOWRAP><P class='Cd'>" & "TOTAL:" & "</P></TD>" & chr(13) & _
				"		<TD class='MTD'><P class='Cd'>" & formata_inteiro(intQtdeSubTotal) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
	
	'	TOTAL GERAL
		if intQtdeFabricantes > 1 then
			x = x & _
				"	<TR>" & chr(13) & _
				"		<TD colspan='2' class='MC'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR>" & chr(13) & _
				"		<TD colspan='2'><P class='C'>TOTAL GERAL</P></TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:ivory;'>" & chr(13) & _
				"		<TD class='MDTE' NOWRAP><P class='Cd'>" & "TOTAL GERAL:" & "</P></TD>" & chr(13) & _
				"		<TD class='MTD'><P class='Cd'>" & formata_inteiro(intQtdeTotal) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & _
			cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD class='MC MD ME' colspan='2'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	
end sub



' _____________________________________
' CONSULTA_ESTOQUE_DETALHE_CUSTO_MEDIO
'
sub consulta_estoque_detalhe_custo_medio
dim r
dim valor_total, vl_sub_total, vlAux, s_sql, cab_table, cab, n_reg, x
dim intQtde, intQtdeTotal, intQtdeSubTotal
dim strFabricanteAnterior, strFabricante, strProduto, intQtdeFabricantes

	s_sql = "SELECT" & _
				" i.fabricante," & _
				" i.produto," & _
				" Coalesce(Sum(qtde),0) AS qtde," & _
				" Coalesce(Sum(qtde* i.vl_custo2),0) AS valor" & _
			" FROM t_ESTOQUE e INNER JOIN t_ESTOQUE_ITEM i ON (e.id_estoque=i.id_estoque)" & _
            " INNER JOIN t_PRODUTO p ON (  i.fabricante = p.fabricante ) AND (i.produto= p.produto)" & _
			" WHERE" & _
				" (kit=0)" & _
				" AND (entrada_especial=0)" & _
				" AND (devolucao_status=0)"

	if IsDate(c_dt_inicio) then
		s_sql = s_sql & " AND (data_entrada >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		s_sql = s_sql & " AND (data_entrada < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
	s_where_temp = ""
	if c_fabricante <> "" then
	for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	    if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		s_where_temp = s_where_temp & _
			   "  (i.fabricante = '" & v_fabricantes(cont) & "')"
	next
	s_sql = s_sql & "AND"
	s_sql =  s_sql & "(" & s_where_temp & ")"
    end if

     s_where_temp = ""
	if c_grupo <> "" then
	v_grupos = split(c_grupo, ", ")
	for cont = Lbound(v_grupos) to Ubound(v_grupos)
	    if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		s_where_temp = s_where_temp & _
			" (grupo = '" & v_grupos(cont) & "')"
	next
	s_sql = s_sql & "AND "
	s_sql = s_sql & "(" & s_where_temp & ")"
		end if

	if Trim(cod_produto) <> "" then
		s_sql = s_sql & " AND (i.produto = '" & cod_produto & "')"
		end if

    if Trim(c_potencia_BTU) <> "" then
        s_sql = s_sql & " AND (potencia_BTU = '" & c_potencia_BTU & "')"
        end if

    if Trim(c_ciclo) <> "" then
        s_sql = s_sql & " AND (ciclo = '" & c_ciclo & "')"
        end if

    if Trim(c_posicao_mercado) <> "" then
        s_sql = s_sql & " AND (posicao_mercado = '" & c_posicao_mercado & "')"
        end if
		
	s_sql = s_sql & " GROUP BY i.fabricante, i.produto" & _
					" ORDER BY i.fabricante, i.produto"

  ' CABEÇALHO
	cab_table = "<TABLE class='MB' cellSpacing=0>" & chr(13)
	
	cab = _
		  "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD valign='bottom' class='MDTE' NOWRAP><P style='width:300px' class='R'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD valign='bottom' class='MTD' NOWRAP ><P style='width:60px' class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD valign='bottom' class='MTD' NOWRAP ><P style='width:110px' class='Rd' style='font-weight:bold;'>REFERÊNCIA MÉDIO</P></TD>" & chr(13) & _
		  "		<TD align='right' class='MTD' valign='bottom' NOWRAP ><P style='width:110px' class='Rd' style='font-weight:bold;'>TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table
	n_reg = 0
	valor_total = 0
	intQtdeTotal = 0
	intQtdeFabricantes = 0
	
	strFabricanteAnterior = "XXXXXXXXXXXXXXXXXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

		strFabricante = Trim("" & r("fabricante"))
		strProduto = Trim("" & r("produto"))

		if strFabricante <> strFabricanteAnterior then
			intQtdeFabricantes = intQtdeFabricantes + 1
		'	SUB-TOTAIS POR FABRICANTE
		'	EXIBE SUB-TOTAL DO FABRICANTE ANTERIOR?
			if n_reg > 0 then
				x = x & _
					"	<TR NOWRAP style='background:ivory;'>" & chr(13) & _
					"		<TD class='MDTE' NOWRAP><P class='Cd'>" & "TOTAL:" & "</P></TD>" & chr(13) & _
					"		<TD class='MTD'><P class='Cd'>" & formata_inteiro(intQtdeSubTotal) & "</P></TD>" & chr(13) & _
					"		<TD class='MTD'>&nbsp;</TD>" & chr(13) & _
					"		<TD class='MTD'><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD colspan='4' class='MC'>&nbsp;</TD>" & chr(13) & _
					"	</TR>" & chr(13)
				end if

			x = x & _
				"	<TR style='background:azure'>" & chr(13) & _
				"		<TD colspan='4' class='MC ME MD'><P class='C'>" & strFabricante & " - " & iniciais_em_maiusculas(x_fabricante(strFabricante)) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			
		'	TÍTULO DAS COLUNAS
			x = x & _
				cab

			vl_sub_total = 0
			intQtdeSubTotal = 0
			strFabricanteAnterior = strFabricante
			end if
			
			
	  ' CONTAGEM
		n_reg = n_reg + 1
		
		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDTE' valign='bottom'><P class='C'>" & strProduto & " - " & produto_formata_descricao_em_html(produto_descricao_html(strFabricante, strProduto)) & "</P></TD>" & chr(13)

	 '> QTDE
		intQtde = r("qtde")
		intQtdeTotal = intQtdeTotal + intQtde
		intQtdeSubTotal = intQtdeSubTotal + intQtde
		x = x & "		<TD class='MTD' valign='bottom'><P class='Cd'>&nbsp;" & formata_inteiro(intQtde) & "</P></TD>" & chr(13)

	 '> CUSTO MÉDIO
		if intQtde = 0 then vlAux = 0 else vlAux = r("valor")/intQtde
		x = x & "		<TD class='MTD' valign='bottom'><P class='Cd'>&nbsp;" & formata_moeda(vlAux) & "</P></TD>" & chr(13)

	 '> VALOR
		x = x & "		<TD align='right' valign='bottom' class='MTD'><P class='Cd'>&nbsp;" & formata_moeda(r("valor")) & "</P></TD>" & chr(13)

		valor_total = valor_total + r("valor")
		vl_sub_total = vl_sub_total + r("valor")
		
		x = x & "	</TR>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		r.movenext
		loop
		
  ' MOSTRA VALOR TOTAL
	if n_reg <> 0 then
	'	SUB-TOTAL DO ÚLTIMO FABRICANTE 
		x = x & "	<TR NOWRAP style='background:ivory;'>" & chr(13) & _
				"		<TD class='MDTE' NOWRAP><P class='Cd'>" & "TOTAL:" & "</P></TD>" & chr(13) & _
				"		<TD class='MTD'><P class='Cd'>" & formata_inteiro(intQtdeSubTotal) & "</P></TD>" & chr(13) & _
				"		<TD class='MTD'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MTD'><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
	
	'	TOTAL GERAL
		if intQtdeFabricantes > 1 then
			x = x & _
				"	<TR>" & chr(13) & _
				"		<TD colspan='4' class='MC'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR>" & chr(13) & _
				"		<TD colspan='4'><P class='C'>TOTAL GERAL</P></TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:ivory;'>" & chr(13) & _
				"		<TD class='MDTE' NOWRAP><P class='Cd'>" & "TOTAL GERAL:" & "</P></TD>" & chr(13) & _
				"		<TD class='MTD'><P class='Cd'>" & formata_inteiro(intQtdeTotal) & "</P></TD>" & chr(13) & _
				"		<TD class='MTD'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MTD'><P class='Cd'>" & formata_moeda(valor_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & _
			cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD class='MC MD ME' colspan='4'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	
end sub



' ____________________________________________
' CONSULTA_ESTOQUE_DETALHE_CUSTO_INDIVIDUAL
'
sub consulta_estoque_detalhe_custo_individual
dim r
dim valor_total, vl_sub_total, vlAux, s_sql, cab_table, cab, n_reg, x
dim intQtde, intQtdeTotal, intQtdeSubTotal
dim strFabricanteAnterior, strFabricante, strProduto, intQtdeFabricantes

	s_sql = "SELECT" & _
				" i.fabricante," & _
				" i.produto," & _
				" i.vl_custo2," & _
				" Coalesce(Sum(qtde),0) AS qtde" & _
			" FROM t_ESTOQUE e INNER JOIN t_ESTOQUE_ITEM i ON (e.id_estoque=i.id_estoque)" & _
            " INNER JOIN t_PRODUTO p ON (  i.fabricante = p.fabricante ) AND (i.produto= p.produto)" & _
			" WHERE" & _
				" (kit=0)" & _
				" AND (entrada_especial=0)" & _
				" AND (devolucao_status=0)"

	if IsDate(c_dt_inicio) then
		s_sql = s_sql & " AND (data_entrada >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		s_sql = s_sql & " AND (data_entrada < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
	s_where_temp = ""
	if c_fabricante <> "" then
	for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	    if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		s_where_temp = s_where_temp & _
			   "  (i.fabricante = '" & v_fabricantes(cont) & "')"
	next
	s_sql = s_sql & "AND"
	s_sql =  s_sql & "(" & s_where_temp & ")"
    end if

     s_where_temp = ""
	if c_grupo <> "" then
	v_grupos = split(c_grupo, ", ")
	for cont = Lbound(v_grupos) to Ubound(v_grupos)
	    if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		s_where_temp = s_where_temp & _
			" (grupo = '" & v_grupos(cont) & "')"
	next
	s_sql = s_sql & "AND "
	s_sql = s_sql & "(" & s_where_temp & ")"
		end if

	if Trim(cod_produto) <> "" then
		s_sql = s_sql & " AND (i.produto = '" & cod_produto & "')"
		end if

    if Trim(c_potencia_BTU) <> "" then
        s_sql = s_sql & " AND (potencia_BTU = '" & c_potencia_BTU & "')"
        end if

    if Trim(c_ciclo) <> "" then
        s_sql = s_sql & " AND (ciclo = '" & c_ciclo & "')"
        end if

    if Trim(c_posicao_mercado) <> "" then
        s_sql = s_sql & " AND (posicao_mercado = '" & c_posicao_mercado & "')"
        end if
		
	s_sql = s_sql & " GROUP BY i.fabricante, i.produto, i.vl_custo2" & _
					" ORDER BY i.fabricante, i.produto, i.vl_custo2"

  ' CABEÇALHO
	cab_table = "<TABLE class='MB' cellSpacing=0>" & chr(13)
	
	cab = _
		  "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD valign='bottom' class='MDTE' NOWRAP><P style='width:300px' class='R'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD valign='bottom' class='MTD' NOWRAP><P style='width:60px' class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD valign='bottom' class='MTD'><P style='width:110px' class='Rd' style='font-weight:bold;'>REFERÊNCIA INDIVIDUAL</P></TD>" & chr(13) & _
		  "		<TD align='right' class='MTD' valign='bottom' NOWRAP ><P style='width:110px' class='Rd' style='font-weight:bold;'>TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table
	n_reg = 0
	valor_total = 0
	intQtdeTotal = 0
	intQtdeFabricantes = 0
	
	strFabricanteAnterior = "XXXXXXXXXXXXXXXXXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

		strFabricante = Trim("" & r("fabricante"))
		strProduto = Trim("" & r("produto"))

		if strFabricante <> strFabricanteAnterior then
			intQtdeFabricantes = intQtdeFabricantes + 1
		'	SUB-TOTAIS POR FABRICANTE
		'	EXIBE SUB-TOTAL DO FABRICANTE ANTERIOR?
			if n_reg > 0 then
				x = x & _
					"	<TR NOWRAP style='background:ivory;'>" & chr(13) & _
					"		<TD class='MDTE' NOWRAP><P class='Cd'>" & "TOTAL:" & "</P></TD>" & chr(13) & _
					"		<TD class='MTD'><P class='Cd'>" & formata_inteiro(intQtdeSubTotal) & "</P></TD>" & chr(13) & _
					"		<TD class='MTD'>&nbsp;</TD>" & chr(13) & _
					"		<TD class='MTD'><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD colspan='4' class='MC'>&nbsp;</TD>" & chr(13) & _
					"	</TR>" & chr(13)
				end if

			x = x & _
				"	<TR style='background:azure'>" & chr(13) & _
				"		<TD colspan='4' class='MC ME MD'><P class='C'>" & strFabricante & " - " & iniciais_em_maiusculas(x_fabricante(strFabricante)) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			
		'	TÍTULO DAS COLUNAS
			x = x & _
				cab

			vl_sub_total = 0
			intQtdeSubTotal = 0
			strFabricanteAnterior = strFabricante
			end if
			
			
	  ' CONTAGEM
		n_reg = n_reg + 1
		
		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDTE' valign='bottom'><P class='C'>" & strProduto & " - " & produto_formata_descricao_em_html(produto_descricao_html(strFabricante, strProduto)) & "</P></TD>" & chr(13)

	 '> QTDE
		intQtde = r("qtde")
		intQtdeTotal = intQtdeTotal + intQtde
		intQtdeSubTotal = intQtdeSubTotal + intQtde
		x = x & "		<TD class='MTD' valign='bottom'><P class='Cd'>&nbsp;" & formata_inteiro(intQtde) & "</P></TD>" & chr(13)

	 '> CUSTO INDIVIDUAL
		vlAux = r("vl_custo2")
		x = x & "		<TD class='MTD' valign='bottom'><P class='Cd'>&nbsp;" & formata_moeda(vlAux) & "</P></TD>" & chr(13)

	 '> VALOR
		vlAux = r("vl_custo2")*r("qtde")
		x = x & "		<TD align='right' valign='bottom' class='MTD'><P class='Cd'>&nbsp;" & formata_moeda(vlAux) & "</P></TD>" & chr(13)

		valor_total = valor_total + vlAux
		vl_sub_total = vl_sub_total + vlAux
		
		x = x & "	</TR>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA VALOR TOTAL
	if n_reg <> 0 then
	'	SUB-TOTAL DO ÚLTIMO FABRICANTE 
		x = x & "	<TR NOWRAP style='background:ivory;'>" & chr(13) & _
				"		<TD class='MDTE' NOWRAP><P class='Cd'>" & "TOTAL:" & "</P></TD>" & chr(13) & _
				"		<TD class='MTD'><P class='Cd'>" & formata_inteiro(intQtdeSubTotal) & "</P></TD>" & chr(13) & _
				"		<TD class='MTD'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MTD'><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
	
	'	TOTAL GERAL
		if intQtdeFabricantes > 1 then
			x = x & _
				"	<TR>" & chr(13) & _
				"		<TD colspan='4' class='MC'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR>" & chr(13) & _
				"		<TD colspan='4'><P class='C'>TOTAL GERAL</P></TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:ivory;'>" & chr(13) & _
				"		<TD class='MDTE' NOWRAP><P class='Cd'>" & "TOTAL GERAL:" & "</P></TD>" & chr(13) & _
				"		<TD class='MTD'><P class='Cd'>" & formata_inteiro(intQtdeTotal) & "</P></TD>" & chr(13) & _
				"		<TD class='MTD'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MTD'><P class='Cd'>" & formata_moeda(valor_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & _
			cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD class='MC MD ME' colspan='4'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
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
#rb_detalhe_aux {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
</style>

<% if rb_detalhe = "SINTETICO" then %>
<style TYPE="text/css">
P.C { font-size:10pt; }
P.Cc { font-size:10pt; }
P.Cd { font-size:10pt; }
P.F { font-size:11pt; }
</style>
<% end if %>


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



<% else %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="rb_detalhe" id="rb_detalhe" value="<%=rb_detalhe%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=cod_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=cod_produto%>">
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Compras II</span>
	<br>
	<%	s = "<span class='N'>Emissão:&nbsp;" & formata_data_hora(Now) & "</span>"
		Response.Write s
	%>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS DA CONSULTA  -->
<table class="Qx" cellSpacing="0">
<!--  PERÍODO  -->
	<tr bgColor="#FFFFFF">
	<td class="MT" NOWRAP><span class="PLTe">Período</span>
		<br><p class="C" style="width:230px;cursor:default;"><%=c_dt_inicio & " a " & c_dt_termino %></p></td>
	</tr>

<!--  FABRICANTE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">Fabricante</span>
		<%	s = cod_fabricante
			if (s<>"") And (s_nome_fabricante<>"") then s = s & " - " & s_nome_fabricante 
			if s = "" then s = "N.I."
		%>
		<br><input name="c_fabricante_aux" id="c_fabricante_aux" readonly tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=s%>"></td>
	</tr>
	
<!--  PRODUTO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">Produto</span>
		<%	s = cod_produto
			if (s<>"") And (s_nome_produto_html<>"") then s = s & " - " & s_nome_produto_html
			if s = "" then s = "N.I."
		%>
		<br>
			<span class="PLLe" style="width:460px;margin-left:2pt;"><%=s%></span>
		<%	s = cod_produto
			if (s<>"") And (s_nome_produto<>"") then s = s & " - " & s_nome_produto 
			if s = "" then s = "N.I."
		%>
			<input type="hidden" name="c_produto_aux" id="c_produto_aux" value="<%=s%>">
		</td>
	</tr>

<!--  GRUPO  -->
<%if c_grupo <> "" then %>
    <tr bgColor="#FFFFFF">
	    <td class="MDBE" NOWRAP><span class="PLTe">Grupo</span>
		    <br><p class="C" style="width:230px;cursor:default;"><%=c_grupo%></p></td>
    </tr>
<%end if %>

<%if c_potencia_BTU <> "" then %>
    <tr bgColor="#FFFFFF">
	    <td class="MDBE" NOWRAP><span class="PLTe">Potência BTU/H</span>
		    <br><p class="C" style="width:230px;cursor:default;"><%=c_potencia_BTU%></p></td>
    </tr>
<%end if %>

<%if c_ciclo <> "" then %>
    <tr bgColor="#FFFFFF">
	    <td class="MDBE" NOWRAP><span class="PLTe">Ciclo</span>
		    <br><p class="C" style="width:230px;cursor:default;"><%=c_ciclo%></p></td>
    </tr>
<%end if %>

<%if c_posicao_mercado <> "" then %>
    <tr bgColor="#FFFFFF">
	    <td class="MDBE" NOWRAP><span class="PLTe">Posição Mercado</span>
		    <br><p class="C" style="width:230px;cursor:default;"><%=c_posicao_mercado%></p></td>
    </tr>
<%end if %>

<!--  TIPO DE DETALHAMENTO  -->
	<tr bgColor="#FFFFFF">
		<% select case rb_detalhe
			case "SINTETICO_FABR":		s = "Sintético por Fabricante"
			case "SINTETICO_PROD":		s = "Sintético por Produto"
			case "CUSTO_MEDIO":			s = "Valor Referência Médio"
			case "CUSTO_INDIVIDUAL":	s = "Valor Referência Individual"
			case else					s = ""
			end select
		%>
	<td class="MDBE" NOWRAP><span class="PLTe">Tipo de Detalhamento</span>
		<br><p class="C" style="width:230px;cursor:default;"><%=s%></p></td>
	</tr>

</table>

<!--  RELATÓRIO  -->
<br>
<%	
	select case rb_detalhe
		case "SINTETICO_FABR"
			consulta_detalhe_sintetico_fabricante
		case "SINTETICO_PROD"
			consulta_detalhe_sintetico_produto
		case "CUSTO_MEDIO"
			consulta_estoque_detalhe_custo_medio
		case "CUSTO_INDIVIDUAL"
			consulta_estoque_detalhe_custo_individual
		end select
%>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
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
