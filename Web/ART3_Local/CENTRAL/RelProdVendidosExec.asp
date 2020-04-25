<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L P R O D V E N D I D O S E X E C . A S P
'     =================================================================
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
	
	const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
	const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
	const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
	const MSO_NUMBER_FORMAT_TEXTO = "\@"
    const COD_CONSULTA_POR_PERIODO_CADASTRO = "CADASTRO"
	const COD_CONSULTA_POR_PERIODO_ENTREGA = "ENTREGA"
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_PRODUTOS_VENDIDOS, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	const OPCAO_UM_CODIGO = "UM"
	const OPCAO_FAIXA_CODIGOS = "FAIXA"
	
	dim alerta
	dim s, s_aux, s_filtro, lista_loja, s_filtro_loja, v_loja, v, i
	dim rb_periodo, c_dt_entregue_inicio, c_dt_entregue_termino, c_dt_cadastro_inicio, c_dt_cadastro_termino, c_loja
	dim rb_fabricante, c_fabricante, c_fabricante_de, c_fabricante_ate
	dim rb_produto, c_produto, c_produto_de, c_produto_ate
	dim c_grupo, c_subgrupo
    dim c_empresa
	dim rb_saida

	alerta = ""

    rb_periodo = Trim(Request.Form("rb_periodo"))

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
	
	rb_fabricante = Ucase(Trim(Request.Form("rb_fabricante")))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_fabricante_de = retorna_so_digitos(Trim(Request.Form("c_fabricante_de")))
	c_fabricante_ate = retorna_so_digitos(Trim(Request.Form("c_fabricante_ate")))
	
	rb_produto = Ucase(Trim(Request.Form("rb_produto")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_produto_de = Ucase(Trim(Request.Form("c_produto_de")))
	c_produto_ate = Ucase(Trim(Request.Form("c_produto_ate")))
	
	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_subgrupo = Ucase(Trim(Request.Form("c_subgrupo")))

	c_loja = Trim(Request.Form("c_loja"))
	lista_loja = substitui_caracteres(c_loja,chr(10),"")
	v_loja = split(lista_loja,chr(13),-1)
    c_empresa = Trim(Request.Form("c_empresa"))
	rb_saida = Ucase(Trim(Request.Form("rb_saida")))

	if alerta = "" then
		call set_default_valor_texto_bd(usuario, "RelProdVendidos|c_grupo", c_grupo)
		call set_default_valor_texto_bd(usuario, "RelProdVendidos|c_subgrupo", c_subgrupo)
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

	dim blnSaidaExcel
	blnSaidaExcel = False
	if alerta = "" then
		if rb_saida = "XLS" then
			blnSaidaExcel = True
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=RelProdVendidos_" & formata_data_yyyymmdd(Now) & "_" & formata_hora_hhnnss(Now) & ".xls"
			Response.Write "<h2>Relatório de  Produtos Vendidos</h2>"
			Response.Write excel_monta_texto_filtro
			Response.Write "<br><br>"
			consulta_executa
			Response.End
			end if
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' EXCEL MONTA TEXTO FILTRO
'
function excel_monta_texto_filtro
dim s, s_aux, s_filtro

'	PERÍODO
    s = ""
    if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
        s_aux = c_dt_cadastro_inicio
        if s_aux = "" then s_aux = "N.I."
	    s = s & s_aux & " e "
        s_aux = c_dt_cadastro_termino
        if s_aux = "" then s_aux = "N.I."
	    s = s & s_aux
        s_filtro = s_filtro & "Cadastrados entre: " & s & "<br>"
    elseif rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
        s_aux = c_dt_entregue_inicio
        if s_aux = "" then s_aux = "N.I."
	    s = s & s_aux & " e "
        s_aux = c_dt_entregue_termino
        if s_aux = "" then s_aux = "N.I."
	    s = s & s_aux
        s_filtro = s_filtro & "Entregues entre: " & s & "<br>"
        end if

'	FABRICANTE
	s = ""
	if rb_fabricante = OPCAO_UM_CODIGO then
		s = c_fabricante
	elseif rb_fabricante = OPCAO_FAIXA_CODIGOS then
		if (c_fabricante_de<>"") Or (c_fabricante_ate<>"") then
			s_aux = c_fabricante_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " a "
			s_aux = c_fabricante_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
			end if
		end if
	
	if s = "" then s = "todos"
	s_filtro = s_filtro & "Fabricante: " & s & "<br>"

'	PRODUTO
	s = ""
	if rb_produto = OPCAO_UM_CODIGO then
		s = c_produto
	elseif rb_produto = OPCAO_FAIXA_CODIGOS then
		if (c_produto_de<>"") Or (c_produto_ate<>"") then
			s_aux = c_produto_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " a "
			s_aux = c_produto_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
			end if
		end if

	if s = "" then s = "todos"
	s_filtro = s_filtro & "Produto: " & s & "<br>"
	
'	GRUPO DE PRODUTOS
	s = c_grupo
	if s = "" then s = "N.I."
	s_filtro = s_filtro & "Grupo de Produtos: " & s & "<br>"

'	SUBGRUPO DE PRODUTOS
	s = c_subgrupo
	if s = "" then s = "N.I."
	s_filtro = s_filtro & "Subgrupo de Produtos: " & s & "<br>"

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
	s_filtro = s_filtro & "Loja(s): " & s & "<br>"

	s_filtro = s_filtro & "Emissão: " & formata_data_hora(Now) & "<br>"
	
	excel_monta_texto_filtro = s_filtro
end function



' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const NOME_CAMPO_DATA = "§DATA§"
const NOME_CAMPO_FABRICANTE = "§FABR§"
const NOME_CAMPO_PRODUTO = "§PROD§"
dim r
dim s, s_aux, s_where_loja, s_where_temp, s_sql, cab_table, cab, n_reg, n_reg_total, x, fabricante_a
dim qtde_sub_total, qtde_total
dim vl_sub_total_faturamento, vl_total_faturamento
dim i, v, qtde_fabricante
dim s_where, s_where_aux, s_where_venda, s_where_devolucao, s_cor, s_nbsp, s_bkg_color, s_align, s_nowrap
dim cont, v_grupos, v_subgrupos

'	CRITÉRIOS COMUNS
'	================
	s_where = ""
	
'	GRUPO DE PRODUTOS
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

'	SUBGRUPO DE PRODUTOS
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
	
'	EMPRESA
    if c_empresa <> "" then
        if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.id_nfe_emitente = " & c_empresa & ")"
	end if

'	LOJA(S)
	s_where_loja = ""
	for i=Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
				s_where_loja = s_where_loja & " (CONVERT(smallint, t_PEDIDO.loja) = " & v_loja(i) & ")"
			else
				s = ""
				if v(Lbound(v))<>"" then 
					if s <> "" then s = s & " AND"
					s = s & " (CONVERT(smallint, t_PEDIDO.loja) >= " & v(Lbound(v)) & ")"
					end if
				if v(Ubound(v))<>"" then
					if s <> "" then s = s & " AND"
					s = s & " (CONVERT(smallint, t_PEDIDO.loja) <= " & v(Ubound(v)) & ")"
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


'	CRITÉRIOS P/ PEDIDOS DE VENDA NORMAIS E P/ DEVOLUÇÕES
'	=====================================================
	s_where_aux = ""

'	PERÍODO
    if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
        if IsDate(c_dt_cadastro_inicio) then
		    if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
		    s_where_aux = s_where_aux & " (" & NOME_CAMPO_DATA & " >= " & bd_formata_data(StrToDate(c_dt_cadastro_inicio)) & ")"
		    end if		
	    if IsDate(c_dt_cadastro_termino) then
		    if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
		    s_where_aux = s_where_aux & " (" & NOME_CAMPO_DATA & " < " & bd_formata_data(StrToDate(c_dt_cadastro_termino)+1) & ")"
		    end if
    elseif rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
        if IsDate(c_dt_entregue_inicio) then
		    if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
		    s_where_aux = s_where_aux & " (" & NOME_CAMPO_DATA & " >= " & bd_formata_data(StrToDate(c_dt_entregue_inicio)) & ")"
		    end if		
	    if IsDate(c_dt_entregue_termino) then
		    if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
		    s_where_aux = s_where_aux & " (" & NOME_CAMPO_DATA & " < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		    end if
        end if

'	FABRICANTE
	if rb_fabricante = OPCAO_UM_CODIGO then
		if c_fabricante <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (" & NOME_CAMPO_FABRICANTE & " = '" & c_fabricante & "')"
			end if
	elseif rb_fabricante = OPCAO_FAIXA_CODIGOS then
		if (c_fabricante_de<>"") Or (c_fabricante_ate<>"") then
			s = ""
			if c_fabricante_de <> "" then
				if s <> "" then s = s & " AND"
				s = s & " (" & NOME_CAMPO_FABRICANTE & " >= '" & c_fabricante_de & "')"
				end if
			if c_fabricante_ate <> "" then
				if s <> "" then s = s & " AND"
				s = s & " (" & NOME_CAMPO_FABRICANTE & " <= '" & c_fabricante_ate & "')"
				end if
			if s <> "" then 
				s = " (" & s & ")"
				if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
				s_where_aux = s_where_aux & s
				end if
			end if
		end if
	
'	PRODUTO
	if rb_produto = OPCAO_UM_CODIGO then
		if c_produto <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (" & NOME_CAMPO_PRODUTO & " = '" & c_produto & "')"
			end if
	elseif rb_produto = OPCAO_FAIXA_CODIGOS then
		if (c_produto_de<>"") Or (c_produto_ate<>"") then
			s = ""
			if c_produto_de <> "" then
				if s <> "" then s = s & " AND"
				s = s & " (" & NOME_CAMPO_PRODUTO & " >= '" & c_produto_de & "')"
				end if
			if c_produto_ate <> "" then
				if s <> "" then s = s & " AND"
				s = s & " (" & NOME_CAMPO_PRODUTO & " <= '" & c_produto_ate & "')"
				end if
			if s <> "" then 
				s = " (" & s & ")"
				if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
				s_where_aux = s_where_aux & s
				end if
			end if
		end if


	s_where_venda = s_where_aux
	s_where_devolucao = s_where_aux
	if s_where_aux <> "" then
        if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
		    s_where_venda = replace(s_where_venda, NOME_CAMPO_DATA, "t_PEDIDO.data")
        elseif rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then 
            s_where_venda = replace(s_where_venda, NOME_CAMPO_DATA, "t_PEDIDO.entregue_data")
            end if
		s_where_devolucao = replace(s_where_devolucao, NOME_CAMPO_DATA, "t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data")
		
		s_where_venda = replace(s_where_venda, NOME_CAMPO_FABRICANTE, "t_PEDIDO_ITEM.fabricante")
		s_where_devolucao = replace(s_where_devolucao, NOME_CAMPO_FABRICANTE, "t_PEDIDO_ITEM_DEVOLVIDO.fabricante")

		s_where_venda = replace(s_where_venda, NOME_CAMPO_PRODUTO, "t_PEDIDO_ITEM.produto")
		s_where_devolucao = replace(s_where_devolucao, NOME_CAMPO_PRODUTO, "t_PEDIDO_ITEM_DEVOLVIDO.produto")
		end if


'	MONTA CONSULTA
'	==============
	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then s = " AND" & s
	s_sql = "SELECT t_PEDIDO_ITEM.fabricante AS fabricante," & _
			" t_PEDIDO_ITEM.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" Sum(qtde) AS qtde_total," & _
			" Sum(qtde*preco_venda) AS valor_total" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			" LEFT JOIN t_PRODUTO ON ((t_PEDIDO_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_PEDIDO_ITEM.produto=t_PRODUTO.produto))" & _
			" WHERE (st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
			s & _
			" GROUP BY t_PEDIDO_ITEM.fabricante, t_PEDIDO_ITEM.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html"

	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT t_PEDIDO_ITEM_DEVOLVIDO.fabricante AS fabricante," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" -Sum(qtde) AS qtde_total," & _
			" -Sum(qtde*preco_venda) AS valor_total" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)" & _
			" LEFT JOIN t_PRODUTO ON ((t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_PRODUTO.fabricante) AND (t_PEDIDO_ITEM_DEVOLVIDO.produto=t_PRODUTO.produto))" & _
			s & _
			" GROUP BY t_PEDIDO_ITEM_DEVOLVIDO.fabricante, t_PEDIDO_ITEM_DEVOLVIDO.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html"

	s_sql = s_sql & " ORDER BY fabricante, produto, descricao, descricao_html, qtde_total DESC"
	
  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE' style='width:60px' valign='bottom' NOWRAP><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:220px' valign='bottom' NOWRAP><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:50px' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:120px' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>FATURAMENTO (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_fabricante = 0
	qtde_sub_total = 0
	qtde_total = 0
	vl_sub_total_faturamento = 0
	vl_total_faturamento = 0
	
	fabricante_a = "XXXXX"
	if blnSaidaExcel then s_nbsp = "" else s_nbsp = "&nbsp;"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
			fabricante_a = Trim("" & r("fabricante"))
			qtde_fabricante = qtde_fabricante + 1
		  ' FECHA TABELA DO FABRICANTE ANTERIOR
			if n_reg > 0 then 
				s_cor="black"
				if qtde_sub_total < 0 then s_cor="red"
				if vl_sub_total_faturamento < 0 then s_cor="red"
				x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
						"		<TD class='MTBE' colspan='2' align='right' NOWRAP><p class='Cd' style='font-weight:bold;'>TOTAL:</p></td>" & chr(13) & _
						"		<TD class='MTB'><p class='Cd' style='font-weight:bold;color:" & s_cor & ";mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(qtde_sub_total)  & "</p></td>" & chr(13) & _
						"		<TD class='MTBD'><p class='Cd' style='font-weight:bold;color:" & s_cor & ";mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_faturamento) & "</p></td>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			n_reg = 0
			qtde_sub_total = 0
			vl_sub_total_faturamento = 0

			if n_reg_total > 0 then x = x & "<BR>" & chr(13)
			
			s = Trim("" & r("fabricante"))
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			if s = "" then s = "&nbsp;"
			x = x & cab_table
			
			if blnSaidaExcel then 
				s_bkg_color = "tomato"
				s_align = " align='center'"
			else
				s_bkg_color = "azure"
				s_align = ""
				end if
			
			if s <> "" then x = x & "	<TR>" & chr(13) & _
									"		<TD class='MDTE' colspan='4'" & s_align & " valign='bottom' class='MB' style='background:" & s_bkg_color & ";'><p class='N' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</p></td>" & chr(13) & _
									"	</TR>" & chr(13)
			x = x & cab
			end if
		

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)
		
		s_cor=""
		if IsNumeric(r("qtde_total")) then if CLng(r("qtde_total")) < 0 then s_cor="red"
		if IsNumeric(r("valor_total")) then if Ccur(r("valor_total")) < 0 then s_cor="red"
		if s_cor <> "" then s_cor="color:" & s_cor & ";"
		
	 '> CÓDIGO DO PRODUTO
		x = x & "		<TD class='MDTE' valign='bottom'><P class='Cn' style='" & s_cor & "mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		if blnSaidaExcel then s_nowrap = " NOWRAP" else s_nowrap = ""
		s = produto_formata_descricao_em_html(Trim("" & r("descricao_html")))
		if (s = "") And (Not blnSaidaExcel) then s = "&nbsp;"
		x = x & "		<TD class='MTD' valign='bottom'" & s_nowrap & "><P class='Cn' style='" & s_cor & "mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s & "</P></TD>" & chr(13)

	 '> QTDE
		s = formata_inteiro(converte_numero(r("qtde_total")))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD align='right' valign='bottom' class='MTD'><P class='Cnd' style='" & s_cor & "mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & s & "</P></TD>" & chr(13)

	 '> VALOR FATURAMENTO
		x = x & "		<TD align='right' valign='bottom' class='MTD'><P class='Cnd' style='" & s_cor & "mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(r("valor_total")) & "</P></TD>" & chr(13)

		qtde_sub_total = qtde_sub_total + r("qtde_total")
		qtde_total = qtde_total + r("qtde_total")

		vl_sub_total_faturamento = vl_sub_total_faturamento + r("valor_total")
		vl_total_faturamento = vl_total_faturamento + r("valor_total")
		
		x = x & "	</TR>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		

  ' MOSTRA TOTAL DO ÚLTIMO FABRICANTE
	if n_reg <> 0 then 
		s_cor="black"
		if qtde_sub_total < 0 then s_cor="red"
		if vl_sub_total_faturamento < 0 then s_cor="red"
		x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
				"		<TD colspan='2' class='MTBE' align='right' NOWRAP><p class='Cd' style='font-weight:bold;'>TOTAL:</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd' style='font-weight:bold;color:" & s_cor & ";mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(qtde_sub_total) & "</p></td>" & chr(13) & _
				"		<TD class='MTBD'><p class='Cd' style='font-weight:bold;color:" & s_cor & ";mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_faturamento) & "</p></td>" & chr(13) & _
				"	</TR>" & chr(13)
		end if
		
'>	TOTAL GERAL
	if qtde_fabricante > 1 then
		s_cor="black"
		if qtde_total < 0 then s_cor="red"
		if vl_total_faturamento < 0 then s_cor="red"
		x = x & "	<TR>" & chr(13) & _
				"		<TD colspan='4' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR>" & chr(13) & _
				"		<TD colspan='4' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:honeydew'>" & _
				"		<TD class='MTBE' colspan='2' align='right' NOWRAP><p class='Cd' style='font-weight:bold;'>TOTAL GERAL:</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd' style='font-weight:bold;color:" & s_cor & ";mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(qtde_total) & "</p></td>" & chr(13) & _
				"		<TD class='MTBD'><p class='Cd' style='font-weight:bold;color:" & s_cor & ";mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_faturamento) & "</p></td>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='4' align='center'><P class='ALERTA'>NENHUM PEDIDO ENCONTRADO</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO FABRICANTE
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
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_cadastro_inicio" id="c_dt_cadastro_inicio" value="<%=c_dt_cadastro_inicio%>" />
<input type="hidden" name="c_dt_cadastro_termino" id="c_dt_cadastro_termino" value="<%=c_dt_cadastro_termino%>" />
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="<%=c_dt_entregue_inicio%>" />
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>" />
<input type="hidden" name="rb_periodo" id="rb_periodo" value="<%=rb_periodo%>" />
<input type="hidden" name="rb_fabricante" id="rb_fabricante" value="<%=rb_fabricante%>" />
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>" />
<input type="hidden" name="c_fabricante_de" id="c_fabricante_de" value="<%=c_fabricante_de%>" />
<input type="hidden" name="c_fabricante_ate" id="c_fabricante_ate" value="<%=c_fabricante_ate%>" />
<input type="hidden" name="rb_produto" id="rb_produto" value="<%=rb_produto%>" />
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>" />
<input type="hidden" name="c_produto_de" id="c_produto_de" value="<%=c_produto_de%>" />
<input type="hidden" name="c_produto_ate" id="c_produto_ate" value="<%=c_produto_ate%>" />
<input type="hidden" name="c_grupo" id="c_grupo" value="<%=c_grupo%>" />
<input type="hidden" name="c_subgrupo" id="c_subgrupo" value="<%=c_subgrupo%>" />
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Produtos Vendidos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
'	PERÍODO
    s = ""
    if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
        s_aux = c_dt_cadastro_inicio
        if s_aux = "" then s_aux = "N.I."
	    s = s & s_aux & " e "
        s_aux = c_dt_cadastro_termino
        if s_aux = "" then s_aux = "N.I."
	    s = s & s_aux
        s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Cadastrados entre:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)
    elseif rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
        s_aux = c_dt_entregue_inicio
        if s_aux = "" then s_aux = "N.I."
	    s = s & s_aux & " e "
        s_aux = c_dt_entregue_termino
        if s_aux = "" then s_aux = "N.I."
	    s = s & s_aux
        s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Entregues entre:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)
        end if

'	FABRICANTE
	s = ""
	if rb_fabricante = OPCAO_UM_CODIGO then
		s = c_fabricante
	elseif rb_fabricante = OPCAO_FAIXA_CODIGOS then
		if (c_fabricante_de<>"") Or (c_fabricante_ate<>"") then
			s_aux = c_fabricante_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " a "
			s_aux = c_fabricante_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
			end if
		end if
		
	if s = "" then s = "todos"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Fabricante:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	PRODUTO
	s = ""
	if rb_produto = OPCAO_UM_CODIGO then
		s = c_produto
	elseif rb_produto = OPCAO_FAIXA_CODIGOS then
		if (c_produto_de<>"") Or (c_produto_ate<>"") then
			s_aux = c_produto_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " a "
			s_aux = c_produto_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
			end if
		end if

	if s = "" then s = "todos"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Produto:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)
	
'	GRUPO DE PRODUTOS
	s = c_grupo
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Grupo de Produtos:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	SUBGRUPO DE PRODUTOS
	s = c_subgrupo
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Subgrupo de Produtos:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	EMPRESA
    s = c_empresa
	if s = "" then 
		s = "todas"
	else
		s = obtem_apelido_empresa_NFe_emitente(c_empresa)
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Empresa:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

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
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Loja(s):&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Emissão:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top' width='99%'><p class='N'>" & formata_data_hora(Now) & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
