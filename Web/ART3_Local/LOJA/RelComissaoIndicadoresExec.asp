<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================
'	  R E L C O M I S S A O I N D I C A D O R E S E X E C . A S P
'     ===========================================================
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
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro
	dim ckb_st_entrega_entregue, c_dt_entregue_inicio, c_dt_entregue_termino
	dim ckb_comissao_paga_sim, ckb_comissao_paga_nao
	dim c_vendedor, c_indicador
	dim c_loja, lista_loja, s_filtro_loja, v_loja, v, i
	dim rb_visao

	alerta = ""

	ckb_st_entrega_entregue = Trim(Request.Form("ckb_st_entrega_entregue"))
	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))

	if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
		c_vendedor = Trim(Request.Form("c_vendedor"))
	else
	'	CONSULTA APENAS AOS SEUS PRÓPRIOS PEDIDOS
		c_vendedor = usuario
		end if
	
	c_indicador = Trim(Request.Form("c_indicador"))

	ckb_comissao_paga_sim = Trim(Request.Form("ckb_comissao_paga_sim"))
	ckb_comissao_paga_nao = Trim(Request.Form("ckb_comissao_paga_nao"))
	rb_visao = Trim(Request.Form("rb_visao"))

'	APENAS PEDIDOS DESTA LOJA
	c_loja = loja
	lista_loja = substitui_caracteres(c_loja,chr(10),"")
	v_loja = split(lista_loja,chr(13),-1)

	if alerta = "" then
		if c_dt_entregue_inicio <> "" then
			if Not IsDate(StrToDate(c_dt_entregue_inicio)) then
				alerta = "DATA DE INÍCIO DO PERÍODO É INVÁLIDA."
				end if
			end if
		end if
	
	if alerta = "" then
		if c_dt_entregue_termino <> "" then
			if Not IsDate(StrToDate(c_dt_entregue_termino)) then
				alerta = "DATA DE TÉRMINO DO PERÍODO É INVÁLIDA."
				end if
			end if
		end if


'	Período de consulta está restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	if operacao_permitida(OP_LJA_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
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
dim s, s_aux, s_sql, x, cab_table, cab, indicador_a, n_reg, n_reg_total, qtde_indicadores
dim vl_preco_venda, vl_sub_total_preco_venda, vl_total_preco_venda
dim vl_preco_NF, vl_sub_total_preco_NF, vl_total_preco_NF
dim vl_RT, vl_sub_total_RT, vl_total_RT
dim vl_RA, vl_sub_total_RA, vl_total_RA
dim perc_RT
dim s_where, s_where_venda, s_where_devolucao, s_where_perdas, s_where_loja, s_cor, s_sinal, s_cor_sinal
dim nome_cliente

'	CRITÉRIOS COMUNS
	s_where = "(LEN(Coalesce(t_PEDIDO__BASE.indicador, '')) > 0)"

	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.vendedor = '" & c_vendedor & "')"
		end if

	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if

'	LOJA(S)
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

'	CRITÉRIO: COMISSÃO PAGA
	s = ""
	s_aux = ckb_comissao_paga_sim
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.comissao_paga = 1)"
		end if

	s_aux = ckb_comissao_paga_nao
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.comissao_paga = 0)"
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	s_where_venda = ""
	if IsDate(c_dt_entregue_inicio) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data >= " & bd_formata_data(StrToDate(c_dt_entregue_inicio)) & ")"
		end if
		
	if IsDate(c_dt_entregue_termino) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		end if

'	CRITÉRIOS PARA DEVOLUÇÕES
	s_where_devolucao = ""
	if IsDate(c_dt_entregue_inicio) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(StrToDate(c_dt_entregue_inicio)) & ")"
		end if
		
	if IsDate(c_dt_entregue_termino) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		end if

'	CRITÉRIOS PARA PERDAS
	s_where_perdas = ""
	if IsDate(c_dt_entregue_inicio) then
		if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		s_where_perdas = s_where_perdas & " (t_PEDIDO_PERDA.data >= " & bd_formata_data(StrToDate(c_dt_entregue_inicio)) & ")"
		end if
		
	if IsDate(c_dt_entregue_termino) then
		if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		s_where_perdas = s_where_perdas & " (t_PEDIDO_PERDA.data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		end if
		
		
	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then s = " AND" & s
	s_sql = "SELECT t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor," & _
			" t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO.entregue_data AS data," & _
			" t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			" t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			" t_CLIENTE.nome AS nome_cliente," & _
			" Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.preco_venda) AS total_preco_venda," & _
			" Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.preco_NF) AS total_preco_NF" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente = t_CLIENTE.id)" & _
			" WHERE (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
			s & _
			" GROUP BY t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor, t_PEDIDO.loja, t_PEDIDO.numero_loja, t_CLIENTE.nome, t_PEDIDO.entregue_data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto"

'	ITENS DEVOLVIDOS
	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION ALL " & _
			"SELECT t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor," & _
			" t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data," & _
			" t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			" t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			" t_CLIENTE.nome AS nome_cliente," & _
			" Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS total_preco_venda," & _
			" Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_NF) AS total_preco_NF" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido)" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente = t_CLIENTE.id)" & _
			s & _
			" GROUP BY t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor, t_PEDIDO.loja, t_PEDIDO.numero_loja, t_CLIENTE.nome, t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto"

'	PERDAS
	s = s_where
	if (s <> "") And (s_where_perdas <> "") then s = s & " AND"
	s = s & s_where_perdas
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION ALL " & _
			"SELECT t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor," & _
			" t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO_PERDA.data AS data," & _
			" t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			" t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			" t_CLIENTE.nome AS nome_cliente," & _
			" Sum(-t_PEDIDO_PERDA.valor) AS total_preco_venda," & _
			" Sum(-t_PEDIDO_PERDA.valor) AS total_preco_NF" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente = t_CLIENTE.id)" & _
			" INNER JOIN t_PEDIDO_PERDA ON (t_PEDIDO.pedido=t_PEDIDO_PERDA.pedido)" & _
			s & _
			" GROUP BY t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor, t_PEDIDO.loja, t_PEDIDO.numero_loja, t_CLIENTE.nome, t_PEDIDO_PERDA.data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto"
	
	s_sql = "SELECT " & _
				"*" & _
			" FROM (" & _
				s_sql & _
				") t" & _
			" ORDER BY t.indicador, t.numero_loja, t.data, t.pedido, t.total_preco_venda DESC"

  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0 style='width: 600px'>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE' valign='bottom' NOWRAP><P style='width:65px' class='R'>Nº ORÇAM</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:68px' class='R'>Nº PEDIDO</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='center' valign='bottom' NOWRAP><P style='width:70px' class='Rc'>DATA</P></TD>" & chr(13)
		  if (rb_visao = "SINTETICA") then
		  cab = cab & "		<TD class='MTD' valign='bottom'><P style='width:80px' class='R' style='font-weight:bold;'>CLIENTE</P></TD>" & chr(13)
		  elseif (rb_visao = "ANALITICA") then
		      cab = cab & "		<TD class='MTD' align='right' valign='bottom'><P style='width:80px' class='Rd' style='font-weight:bold;'>VL PEDIDO</P></TD>" & chr(13) & _
		      "		<TD class='MTD' align='right' valign='bottom'><P style='width:80px' class='Rd' style='font-weight:bold;'>COM (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		      "		<TD class='MTD' align='right' valign='bottom'><P style='width:80px' class='Rd' style='font-weight:bold;'>RA (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		      "		<TD class='MTD' valign='bottom'><P style='width:70px' class='R' style='font-weight:bold;'>ST PAGTO</P></TD>" & chr(13) & _
		      "		<TD class='MTD' align='center' valign='bottom'><P style='width:20px' class='Rc' style='font-weight:bold;'>+/-</P></TD>" & chr(13)
		  end if
		  cab = cab & "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_indicadores = 0
	vl_sub_total_preco_venda = 0
	vl_total_preco_venda = 0
	vl_sub_total_preco_NF = 0
	vl_total_preco_NF = 0
	vl_sub_total_RT = 0
	vl_total_RT = 0
	vl_sub_total_RA = 0
	vl_total_RA = 0

	indicador_a = "XXXXXXXXXXXX"
	set r = cn.execute(s_sql)
	do while Not r.Eof
	nome_cliente = r("nome_cliente")
	'	MUDOU DE INDICADOR?
		if Trim("" & r("indicador"))<>indicador_a then
			indicador_a = Trim("" & r("indicador"))
			qtde_indicadores = qtde_indicadores + 1
		  ' FECHA TABELA DO INDICADOR ANTERIOR
			if x <> "" then 
				s_cor="black"
				if (vl_sub_total_preco_venda < 0 And rb_visao = "ANALITICA") then s_cor="red"
				if (vl_sub_total_RT < 0 And rb_visao = "ANALITICA") then s_cor="red"
				if (vl_sub_total_RA < 0 And rb_visao = "ANALITICA") then s_cor="red"
				x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13)
				    if (rb_visao = "ANALITICA") then
						x = x & "		<TD class='MTBE' COLSPAN='3' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
						"TOTAL:</p></td>" & chr(13)  & _
						"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</p></TD>" & chr(13) & _
						"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</p></td>" & chr(13) & _
						"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</p></td>" & chr(13) & _
						"		<TD class='MTBD' colspan='2'><p class='Cd' style='color:" & s_cor & ";'>&nbsp;</p></td>" & chr(13)
						else
						x = x & "		<TD class='MTBE' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
						"TOTAL DE PEDIDOS:</p></td>" & chr(13)  & _
						"<TD class='MTBD' colspan='3'><p class='C' style='color:" & s_cor & ";'>" & n_reg & "</p></td>" & chr(13)
						end if
						x = x & "	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			n_reg = 0
			vl_sub_total_preco_venda = 0
			vl_sub_total_preco_NF = 0
			vl_sub_total_RT = 0
			vl_sub_total_RA = 0

			if n_reg_total > 0 then x = x & "<BR>"
			s = Trim("" & r("indicador"))
			s_aux = x_orcamentista_e_indicador(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then x = x & "	<TR>" & chr(13) & _
									"		<TD class='MDTE' COLSPAN='8' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
									"	</tr>" & chr(13)
			x = x & cab
			end if
		

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>"  & chr(13)

	'	EVITA DIFERENÇAS DE ARREDONDAMENTO
		vl_preco_venda = converte_numero(formata_moeda(r("total_preco_venda")))
		vl_preco_NF = converte_numero(formata_moeda(r("total_preco_NF")))
		perc_RT = r("perc_RT")
		vl_RT = (perc_RT/100) * vl_preco_venda
		vl_RA = vl_preco_NF - vl_preco_venda
		if (vl_preco_venda < 0) Or (vl_RT < 0) Or (vl_RA < 0) then
		    if (rb_visao = "ANALITICA") then
			    s_cor = "red"
			    s_cor_sinal = "red"
			    s_sinal = "-"
			end if
		else
			s_cor = "black"
			s_cor_sinal = "green"
			s_sinal = "+"
			end if

	 '> Nº ORÇAMENTO
		s = Trim("" & r("orcamento"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MDTE'><P class='Cn'><a style='color:" & s_cor & ";' href='javascript:fORCConcluir(" & _
				chr(34) & s & chr(34) & ")' title='clique para consultar o orçamento'>" & _
				s & "</a></P></TD>" & chr(13)

	 '> Nº PEDIDO
		x = x & "		<TD class='MTD'><P class='Cn'><a style='color:" & s_cor & ";' href='javascript:fPEDConcluir(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & _
				Trim("" & r("pedido")) & "</a></P></TD>" & chr(13)

	 '> DATA
		s = formata_data(r("data"))
		x = x & "		<TD align='center' class='MTD'><P class='Cnc' style='color:" & s_cor & ";'>" & s & "</P></TD>" & chr(13)
		
     if (rb_visao = "SINTETICA") then
        '> NOME DO CLIENTE  
            x = x & "		<TD align='left' class='MTD' style='width:350px'><P class='Cn' style='color:" & s_cor & ";'>" & nome_cliente & "</P></TD>" & chr(13)

     elseif (rb_visao = "ANALITICA") then
	     '> VALOR DO PEDIDO (PREÇO DE VENDA)
		    x = x & "		<TD align='right' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_preco_venda) & "</P></TD>" & chr(13)

	     '> COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
		    x = x & "		<TD align='right' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT) & "</P></TD>" & chr(13)

	     '> RA
		    x = x & "		<TD align='right' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA) & "</P></TD>" & chr(13)

	     '> STATUS DE PAGAMENTO
		    x = x & "		<TD class='MTD'><P class='Cn' style='color:" & s_cor & ";'>" & x_status_pagto(Trim("" & r("st_pagto"))) & "</P></TD>" & chr(13)

	     '> +/-
		    x = x & "		<TD align='center' class='MTD'><P class='C' style='font-family:Courier,Arial;color:" & s_cor_sinal & "'>" & s_sinal & "</P></TD>" & chr(13)
    end if
		vl_sub_total_preco_venda = vl_sub_total_preco_venda + r("total_preco_venda")
		vl_total_preco_venda = vl_total_preco_venda + r("total_preco_venda")
		vl_sub_total_preco_NF = vl_sub_total_preco_NF + r("total_preco_NF")
		vl_total_preco_NF = vl_total_preco_NF + r("total_preco_NF")
		vl_sub_total_RT = vl_sub_total_RT + vl_RT
		vl_total_RT = vl_total_RT + vl_RT
		vl_sub_total_RA = vl_sub_total_RA + vl_RA
		vl_total_RA = vl_total_RA + vl_RA
		
		x = x & "	</TR>" & chr(13)
			
		r.MoveNext
		loop
		
  ' MOSTRA TOTAL DO ÚLTIMO INDICADOR
	if n_reg <> 0 then 
		s_cor="black"
		if vl_sub_total_preco_venda < 0 then s_cor="red"
		if vl_sub_total_RT < 0 then s_cor="red"
		if vl_sub_total_RA < 0 then s_cor="red"
		x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13)
		     if (rb_visao = "ANALITICA") then
				x = x & "		<TD COLSPAN='3' class='MTBE' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
				"TOTAL:</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</p></td>" & chr(13) & _
				"		<TD class='MTBD' colspan='2'><p class='Cd' style='color:" & s_cor & ";'>&nbsp;</p></td>" & chr(13)
			else
			    x = x & "		<TD class='MTBE' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
				"TOTAL DE PEDIDOS:</p></td>" & chr(13) & _
				"		<TD class='MTBD' colspan='4'><p class='C' style='color:" & s_cor & ";'>" & n_reg & "</p></td>" & chr(13)
			end if
				x = x & "	</TR>" & chr(13)
				
	'>	TOTAL GERAL
		if qtde_indicadores > 1 then
			s_cor="black"
			if vl_total_preco_venda < 0 then s_cor="red"
			if vl_total_RT < 0 then s_cor="red"
			if vl_total_RA < 0 then s_cor="red"
			if (rb_visao = "SINTETICA") then
			x = x & "	<TR>" & chr(13) & _
					"		<TD COLSPAN='8' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD COLSPAN='8' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
					"		<TD class='MTBE' COLSPAN='3' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL DE PEDIDOS:</p></td>" & chr(13) & _
					"		<TD class='MTBD' colspan='3'><p class='C' style='color:" & s_cor & ";'>" & n_reg_total & "</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
		    else 
		    x = x & "	<TR>" & chr(13) & _
					"		<TD COLSPAN='8' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD COLSPAN='8' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
					"		<TD class='MTBE' COLSPAN='3' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL:</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_preco_venda) & "</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA) & "</p></td>" & chr(13) & _
					"		<TD class='MTBD' colspan='2'><p class='C' style='color:" & s_cor & ";'>&nbsp;</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='8'><P class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS NO PERÍODO ESPECIFICADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO INDICADOR
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	end if
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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fPEDConcluir( id_pedido ) {
	fREL.action = "Pedido.asp";
	fREL.pedido_selecionado.value = id_pedido;
	fREL.submit();
}

function fORCConcluir( id_orcamento ) {
	if (trim(id_orcamento) == '') return;
	fREL.action = "Orcamento.asp";
	fREL.orcamento_selecionado.value = id_orcamento;
	fREL.submit();
}
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
<body onload="bVOLTAR.focus();" link="black" vlink="black" alink="black">
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

<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<input type="hidden" name="orcamento_selecionado" id="orcamento_selecionado" value="">

<input type="hidden" name="ckb_st_entrega_entregue" id="ckb_st_entrega_entregue" value="<%=ckb_st_entrega_entregue%>">
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="<%=c_dt_entregue_inicio%>">
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>">
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">
<input type="hidden" name="ckb_comissao_paga_sim" id="ckb_comissao_paga_sim" value="<%=ckb_comissao_paga_sim%>">
<input type="hidden" name="ckb_comissao_paga_nao" id="ckb_comissao_paga_nao" value="<%=ckb_comissao_paga_nao%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	PERÍODO: PEDIDOS ENTREGUES ENTRE
	s = ""
	if (c_dt_entregue_inicio <> "") Or (c_dt_entregue_termino <> "") then
	'	DEVIDO AO WORD WRAP: SÓ FAZ WORD WRAP QUANDO ENCONTRA CHR(32), OU SEJA, MANTÉM AGRUPADO TEXTO COM &nbsp;
		if s <> "" then s = s & ",&nbsp; "
		s_aux = c_dt_entregue_inicio
		if s_aux = "" then s_aux = "N.I."
		s_aux = " " & s_aux & " a "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		s_aux = c_dt_entregue_termino
		if s_aux = "" then s_aux = "N.I."
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Período Entrega:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	COMISSÃO PAGA
	s = ""
	s_aux = ckb_comissao_paga_sim
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & "paga"
		end if
	
	s_aux = ckb_comissao_paga_nao
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & "não-paga"
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Comissão:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	INDICADOR
	if c_indicador <> "" then
		s = c_indicador
		s_aux = x_orcamentista_e_indicador(c_indicador)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Indicador:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	VENDEDOR
	if c_vendedor <> "" then
		s = c_vendedor
		s_aux = x_usuario(c_vendedor)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Vendedor:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	LOJA
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
				"		<td align='right' valign='top' NOWRAP><p class='N'>Loja:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	EMISSÃO
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & formata_data_hora(Now) & "</p></td></tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
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
