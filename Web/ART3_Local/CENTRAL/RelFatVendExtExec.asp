<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L F A T V E N D E X T E X E C . A S P
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
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_FATURAMENTO_VENDEDORES_EXT, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, lista_loja, s_filtro_loja, v_loja, v, i
	dim c_dt_inicio, c_dt_termino, c_loja

	alerta = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))

	c_loja = Trim(Request.Form("c_loja"))
	lista_loja = substitui_caracteres(c_loja,chr(10),"")
	v_loja = split(lista_loja,chr(13),-1)
	

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
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim s, s_aux, s_where, s_where_loja, s_where_devolucao, s_where_venda, s_from, s_cor
dim x, s_sql, cab_table, cab, loja_a, qtde_lojas
dim i, v, n_reg, n_reg_total, qtde_sub_total, qtde_total
dim vl_faturamento, vl_cmv, vl_lucro, margem
dim vl_sub_total_faturamento, vl_sub_total_cmv, vl_total_faturamento, vl_total_cmv

'	CLÁUSULA FROM COMUM
	s_from = " FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO_BASE__AUX ON" & _
			 " (t_PEDIDO.pedido_base=t_PEDIDO_BASE__AUX.pedido)"

'	CRITÉRIOS COMUNS
	s_where = " (t_PEDIDO_BASE__AUX.venda_externa<>0)"
	
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
	
'	IMPORTANTE: SEGUNDO O ROGÉRIO (ARTVEN), OS CÁLCULOS SÃO BASEADOS NO PREÇO DE
'	==========	LISTA DO PRODUTO NO ATO DA VENDA, OU SEJA, NÃO IMPORTA O CUSTO DE 
'				AQUISIÇÃO DO PRODUTO QUANDO ELE DEU ENTRADA NO ESTOQUE.
	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then s = " AND" & s
	s_sql = "SELECT t_PEDIDO.loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO_ITEM.fabricante AS fabricante, t_PEDIDO_ITEM.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" t_PEDIDO_ITEM.preco_lista AS preco_lista, t_PEDIDO_ITEM.preco_venda," & _
			" t_PEDIDO_ITEM.markup_fabricante, Sum(t_PEDIDO_ITEM.qtde) AS qtde_total" & _
			s_from & _
			" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			" LEFT JOIN t_PRODUTO ON (t_PEDIDO_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_PEDIDO_ITEM.produto=t_PRODUTO.produto)" & _
			" WHERE (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
			s & _
			" GROUP BY t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_ITEM.fabricante, t_PEDIDO_ITEM.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html," & _
			" t_PEDIDO_ITEM.preco_lista, t_PEDIDO_ITEM.preco_venda, t_PEDIDO_ITEM.markup_fabricante"

	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT t_PEDIDO.loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.fabricante AS fabricante, t_PEDIDO_ITEM_DEVOLVIDO.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.preco_lista AS preco_lista, t_PEDIDO_ITEM_DEVOLVIDO.preco_venda," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.markup_fabricante, -Sum(t_PEDIDO_ITEM_DEVOLVIDO.qtde) AS qtde_total" & _
			s_from & _
			" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido)" & _
			" LEFT JOIN t_PRODUTO ON (t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_PRODUTO.fabricante) AND (t_PEDIDO_ITEM_DEVOLVIDO.produto=t_PRODUTO.produto)" & _
			s & _
			" GROUP BY t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_ITEM_DEVOLVIDO.fabricante, t_PEDIDO_ITEM_DEVOLVIDO.produto," & _
			" t_PRODUTO.descricao, t_PRODUTO.descricao_html, t_PEDIDO_ITEM_DEVOLVIDO.preco_lista, t_PEDIDO_ITEM_DEVOLVIDO.preco_venda, t_PEDIDO_ITEM_DEVOLVIDO.markup_fabricante"

	s_sql = s_sql & " ORDER BY numero_loja, fabricante, produto, descricao, descricao_html, qtde_total DESC, preco_lista"


  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE' valign='bottom' NOWRAP><P style='width:31px' class='R'>Fabr</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:56px' class='R'>Produto</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom'><P style='width:200px' class='R'>Descrição</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='right' valign='bottom' NOWRAP><P style='width:31px' class='Rd' style='font-weight:bold;'>Qtde</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:71px' class='Rd' style='font-weight:bold;'>Faturamento Total (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:70px' class='Rd' style='font-weight:bold;'>CMV (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:70px' class='Rd' style='font-weight:bold;'>Lucro (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:41px' class='Rd' style='font-weight:bold;'>Margem</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)

	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_sub_total = 0
	qtde_total = 0
	vl_sub_total_faturamento = 0
	vl_total_faturamento = 0
	vl_sub_total_cmv = 0
	vl_total_cmv = 0
	qtde_lojas = 0
	
	loja_a = "XXXXX"

	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE LOJA?
		if Trim("" & r("loja"))<>loja_a then
			loja_a = Trim("" & r("loja"))
			qtde_lojas = qtde_lojas + 1
		  ' FECHA TABELA DA LOJA ANTERIOR
			if n_reg > 0 then 
				if vl_sub_total_faturamento = 0 then
					margem = 0
				else
					margem = ((vl_sub_total_faturamento-vl_sub_total_cmv)/vl_sub_total_faturamento)*100
					end if
				s_cor="black"
				if qtde_sub_total < 0 then s_cor="red"
				if vl_sub_total_faturamento < 0 then s_cor="red"
				if vl_sub_total_cmv < 0 then s_cor="red"
				if (vl_sub_total_faturamento-vl_sub_total_cmv) < 0 then s_cor="red"
				x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
						"		<TD class='MTBE' colspan='3' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
						"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_sub_total) & "</p></TD>" & chr(13) & _
						"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_faturamento) & "</p></td>" & chr(13) & _
						"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_cmv) & "</p></td>" & chr(13) & _
						"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_faturamento-vl_sub_total_cmv) & "</p></td>" & chr(13) & _
						"		<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(margem) & "%" & "</p></td>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x = "<BR>" & chr(13)
				end if

			n_reg = 0
			qtde_sub_total = 0
			vl_sub_total_faturamento = 0
			vl_sub_total_cmv = 0

			if n_reg_total > 0 then x = x & "<BR>" & chr(13)
			
			s = Trim("" & r("loja"))
			s_aux = x_loja(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			if s = "" then s = "&nbsp;"
			x = x & cab_table
			x = x & "	<TR>" & chr(13) & _
					"		<TD class='MDTE' colspan='8' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			x = x & cab
			end if
		

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)

		vl_faturamento = r("qtde_total")*r("preco_venda")
	'	ARREDONDAMENTO
		vl_cmv = converte_numero(formata_moeda(r("qtde_total")*(r("preco_lista")-r("preco_lista")*r("markup_fabricante")/100)))
		vl_lucro = vl_faturamento-vl_cmv

		s_cor="black"
		if converte_numero(r("qtde_total")) < 0 then 
			s_cor="red"
		elseif vl_lucro < 0 then 
			s_cor = "maroon"
			end if

	 '> FABRICANTE
		x = x & "		<TD class='MDTE' valign='bottom'><P class='Cn' style='color:" & s_cor & ";'>" & Trim("" & r("fabricante")) & "</P></TD>" & chr(13)

	 '> CÓDIGO DO PRODUTO
		x = x & "		<TD class='MTD' valign='bottom'><P class='Cn' style='color:" & s_cor & ";'>" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO DO PRODUTO
		s = Trim("" & r("descricao_html"))
		if s <> "" then s = produto_formata_descricao_em_html(s) else s = "&nbsp;"
		x = x & "		<TD class='MTD' valign='bottom'><P class='Cn' style='color:" & s_cor & ";'>" & s & "</P></TD>" & chr(13)

	 '> QUANTIDADE
		x = x & "		<TD align='right' class='MTD' valign='bottom'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(r("qtde_total")) & "</P></TD>" & chr(13)

	 '> FATURAMENTO
		x = x & "		<TD align='right' class='MTD' valign='bottom'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_faturamento) & "</P></TD>" & chr(13)

	 '> CMV
		x = x & "		<TD align='right' class='MTD' valign='bottom'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_cmv) & "</P></TD>" & chr(13)

	 '> LUCRO
		x = x & "		<TD align='right' class='MTD' valign='bottom'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lucro) & "</P></TD>" & chr(13)

	 '> MARGEM
		if vl_faturamento = 0 then
			margem = 0
		else
			margem = ((vl_faturamento-vl_cmv)/vl_faturamento)*100
			end if
		x = x & "		<TD align='right' class='MTD' valign='bottom'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(margem) & "%" & "</P></TD>" & chr(13)
		
		qtde_sub_total = qtde_sub_total + r("qtde_total")
		qtde_total = qtde_total + r("qtde_total")
		
		vl_sub_total_faturamento = vl_sub_total_faturamento + vl_faturamento
		vl_sub_total_cmv = vl_sub_total_cmv + vl_cmv
		
		vl_total_faturamento = vl_total_faturamento + vl_faturamento
		vl_total_cmv = vl_total_cmv + vl_cmv
		
		x = x & "	</TR>" & chr(13)
	
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		if vl_sub_total_faturamento = 0 then
			margem = 0
		else
			margem = ((vl_sub_total_faturamento-vl_sub_total_cmv)/vl_sub_total_faturamento)*100
			end if
		
		s_cor="black"
		if qtde_sub_total < 0 then s_cor="red"
		if vl_sub_total_faturamento < 0 then s_cor="red"
		if vl_sub_total_cmv < 0 then s_cor="red"
		if (vl_sub_total_faturamento-vl_sub_total_cmv) < 0 then s_cor="red"
		x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
				"		<TD colspan='3' class='MTBE' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_sub_total) & "</p></TD>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_faturamento) & "</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_cmv) & "</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_faturamento-vl_sub_total_cmv) & "</p></td>" & chr(13) & _
				"		<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(margem) & "%" & "</p></td>" & chr(13) & _
				"	</TR>" & chr(13)
		
	'>	TOTAL GERAL
		if qtde_lojas > 1 then
			if vl_total_faturamento = 0 then
				margem = 0
			else
				margem = ((vl_total_faturamento-vl_total_cmv)/vl_total_faturamento)*100
				end if
			s_cor="black"
			if qtde_total < 0 then s_cor="red"
			if vl_total_faturamento < 0 then s_cor="red"
			if vl_total_cmv < 0 then s_cor="red"
			if (vl_total_faturamento-vl_total_cmv) < 0 then s_cor="red"
			x = x & "	<TR>" & chr(13) & _
					"		<TD colspan='8' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD colspan='8' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
					"		<TD class='MTBE' colspan='3' NOWRAP><p class='Cd'>TOTAL GERAL:</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_total) & "</p></TD>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_faturamento) & "</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_cmv) & "</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_faturamento-vl_total_cmv) & "</p></td>" & chr(13) & _
					"		<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(margem) & "%" & "</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='8'><P class='ALERTA'>&nbsp;&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DA ÚLTIMA LOJA
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
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Faturamento Vendedores Externos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Período:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
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
