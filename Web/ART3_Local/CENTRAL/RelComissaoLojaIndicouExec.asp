<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L C O M I S S A O L O J A I N D I C O U E X E C . A S P
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
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_COMISSAO_LOJA_POR_INDICACAO, s_lista_operacoes_permitidas) then 
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
const TAMANHO_NUM_LOJA_ORDENACAO = 5
dim r
dim s, s_aux, s_sql, cab_table, cab, n_reg, n_reg_total, x, loja_a, loja_indicou_a
dim s_where, s_where_loja, s_where_venda, s_where_devolucao, s_where_perdas, s_from, s_cor
dim vl_subtotal_pedido, vl_subtotal_comissao
dim vl_total_pedido, vl_total_comissao
dim vl_total_geral_pedido, vl_total_geral_comissao
dim vl_comissao
dim i, v, qtde_loja, novo_bloco_loja_indicou
dim vL(), idx

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
	
'	CRITÉRIOS PARA PERDAS
	s_where_perdas = ""
	if IsDate(c_dt_inicio) then
		if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		s_where_perdas = s_where_perdas & " (t_PEDIDO_PERDA.data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		s_where_perdas = s_where_perdas & " (t_PEDIDO_PERDA.data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if


	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then s = " AND" & s
	s_sql = "SELECT t_PEDIDO.loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO_BASE__AUX.loja_indicou, CONVERT(smallint,t_PEDIDO_BASE__AUX.loja_indicou) AS numero_loja_indicou," & _
			" t_PEDIDO.data, t_PEDIDO.pedido," & _
			" t_PEDIDO_BASE__AUX.comissao_loja_indicou," & _
			" Sum(qtde*preco_venda) AS valor_total" & _
			s_from & _
			" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			" WHERE (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
			s & _
			" GROUP BY t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_BASE__AUX.loja_indicou, t_PEDIDO.data, t_PEDIDO.pedido, t_PEDIDO_BASE__AUX.comissao_loja_indicou"

'	ITENS DEVOLVIDOS
	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT t_PEDIDO.loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO_BASE__AUX.loja_indicou, CONVERT(smallint,t_PEDIDO_BASE__AUX.loja_indicou) AS numero_loja_indicou," & _
			" t_PEDIDO.data, t_PEDIDO.pedido," & _
			" t_PEDIDO_BASE__AUX.comissao_loja_indicou," & _
			" Sum(-qtde*preco_venda) AS valor_total" & _
			s_from & _
			" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido)" & _
			s & _
			" GROUP BY t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_BASE__AUX.loja_indicou, t_PEDIDO.data, t_PEDIDO.pedido, t_PEDIDO_BASE__AUX.comissao_loja_indicou"

'	PERDAS
	s = s_where
	if (s <> "") And (s_where_perdas <> "") then s = s & " AND"
	s = s & s_where_perdas
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT t_PEDIDO.loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO_BASE__AUX.loja_indicou, CONVERT(smallint,t_PEDIDO_BASE__AUX.loja_indicou) AS numero_loja_indicou," & _
			" t_PEDIDO.data, t_PEDIDO.pedido," & _
			" t_PEDIDO_BASE__AUX.comissao_loja_indicou," & _
			" Sum(-valor) AS valor_total" & _
			s_from & _
			" INNER JOIN t_PEDIDO_PERDA ON (t_PEDIDO.pedido=t_PEDIDO_PERDA.pedido)" & _
			s & _
			" GROUP BY t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_BASE__AUX.loja_indicou, t_PEDIDO.data, t_PEDIDO.pedido, t_PEDIDO_BASE__AUX.comissao_loja_indicou"

	s_sql = s_sql & " ORDER BY numero_loja, numero_loja_indicou, t_PEDIDO.data, t_PEDIDO.pedido, valor_total DESC"
	
  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDE' style='width:50px;background:white;' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:80px' valign='bottom' NOWRAP><P class='R'>PEDIDO</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:120px' align='right' valign='bottom' NOWRAP><P class='Rd' style='font-weight:bold;'>VALOR</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:120px' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>COMISSÃO</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)

'	AS ROTINAS DE ORDENAÇÃO USAM VETORES QUE SE INICIAM NA POSIÇÃO 1
	redim vL(1)
	for i = Lbound(vL) to Ubound(vL)
		set vL(i) = New cl_TRES_COLUNAS
		with vL(i)
			.c1 = ""
			.c2 = 0
			.c3 = 0
			end with
		next
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_loja = 0
	vl_subtotal_pedido = 0
	vl_subtotal_comissao = 0
	vl_total_pedido = 0
	vl_total_comissao = 0
	vl_total_geral_pedido = 0
	vl_total_geral_comissao = 0
			
	loja_a = "XXXXX"
	loja_indicou_a = "XXXXX"
	novo_bloco_loja_indicou = False
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	MUDOU "LOJA QUE INDICOU"?
	'	ATENÇÃO PARA O CASO EM QUE MUDOU DE LOJA, MAS POR COINCIDÊNCIA A "LOJA QUE INDICOU" É A MESMA DO ÚLTIMO REGISTRO DA LOJA ANTERIOR
		if (Trim("" & r("loja")) <> loja_a) Or (Trim("" & r("loja_indicou"))<> loja_indicou_a) then
			novo_bloco_loja_indicou = True
			loja_indicou_a = Trim("" & r("loja_indicou"))
		  ' FECHA BLOCO DA "LOJA QUE INDICOU" ANTERIOR?
			if n_reg > 0 then
				s_cor="black"
				if (vl_subtotal_pedido<0) Or (vl_subtotal_comissao<0) then s_cor="red"
				x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
						"		<TD class='MDE' style='background:white;'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MC' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
						"		<TD class='MTD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_subtotal_pedido) & "</p></td>" & chr(13) & _
						"		<TD class='MTD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_subtotal_comissao) & "</p></td>" & chr(13) & _
						"	</TR>" & chr(13)
				Response.Write x
				x = ""
				end if
			
			vl_subtotal_pedido = 0
			vl_subtotal_comissao = 0
			end if
			
	'	MUDOU DE LOJA?
		if Trim("" & r("loja")) <> loja_a then
			novo_bloco_loja_indicou = True
			loja_a = Trim("" & r("loja"))
			qtde_loja = qtde_loja + 1
		  ' FECHA TABELA DA LOJA ANTERIOR
			if n_reg > 0 then
				s_cor="black"
				if (vl_total_pedido<0) Or (vl_total_comissao<0) then s_cor="red"
				x = x & "	<TR NOWRAP style='background: #FFF0E0'>" & chr(13) & _
						"		<TD class='MTBE' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
						"		<TD class='MTB' colspan='2'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_pedido) & "</p></td>" & chr(13) & _
						"		<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_comissao) & "</p></td>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x = "<BR>" & chr(13)
				end if

			n_reg = 0
			vl_total_pedido = 0
			vl_total_comissao = 0

			if n_reg_total > 0 then x = x & "<BR>" & chr(13)
			s = Trim("" & r("loja"))
			s_aux = x_loja(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s = "" then s = "&nbsp;"
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MDTE' colspan='4' valign='bottom' class='MB' style='background:powderblue;'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		
	'	ABRE NOVO BLOCO "LOJA QUE INDICOU"?
		if novo_bloco_loja_indicou then
			novo_bloco_loja_indicou = False
			s = Trim("" & r("loja_indicou"))
			s_aux = x_loja(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			if s = "" then s = "&nbsp;"
			x = x & "	<TR>" & chr(13) & _
					"		<TD class='MDTE'>&nbsp;</TD>" & chr(13) & _
					"		<TD class='MTD' colspan='3' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			x = x & cab
			end if


	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	'	DEVIDO A ARREDONDAMENTOS
		vl_comissao = converte_numero(formata_moeda(r("valor_total") * (r("comissao_loja_indicou")/100)))

		s_cor="black"
		if IsNumeric(r("valor_total")) then if Ccur(r("valor_total")) < 0 then s_cor="red"
		if (vl_comissao < 0) then s_cor="red"

	'> IDENTAÇÃO
		x = x & "		<TD class='MDE'>&nbsp;</TD>" & chr(13)

	 '> PEDIDO
		x = x & "		<TD class='MTD'><P class='Cn' style='color:" & s_cor & ";'><a href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & _
				Trim("" & r("pedido")) & "</a></P></TD>" & chr(13)

	 '> VALOR DO PEDIDO
		x = x & "		<TD align='right' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_total")) & "</P></TD>" & chr(13)

	 '>	PERCENTUAL DA COMISSÃO
		x = x & "		<TD align='right' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_comissao) & "</P></TD>" & chr(13)
		
		vl_subtotal_pedido = vl_subtotal_pedido + r("valor_total")
		vl_total_pedido = vl_total_pedido + r("valor_total")
		vl_total_geral_pedido = vl_total_geral_pedido + r("valor_total")
		
		vl_subtotal_comissao = vl_subtotal_comissao + vl_comissao
		vl_total_comissao = vl_total_comissao + vl_comissao
		vl_total_geral_comissao = vl_total_geral_comissao + vl_comissao

	'	TOTALIZAÇÃO POR "LOJA QUE INDICOU"
		s = normaliza_codigo(Trim("" & r("loja_indicou")), TAMANHO_NUM_LOJA_ORDENACAO)
		if localiza_cl_tres_colunas(vL, s, idx) then
			with vL(idx)
				.c2 = .c2 + r("valor_total")
				.c3 = .c3 + vl_comissao
				end with
		else
			if (vL(Ubound(vL)).c1<>"") then
				redim preserve vL(Ubound(vL)+1)
				set vL(Ubound(vL)) = New cl_TRES_COLUNAS
				end if
			with vL(Ubound(vL))
				.c1 = normaliza_codigo(Trim("" & r("loja_indicou")), TAMANHO_NUM_LOJA_ORDENACAO)
				.c2 = .c2 + r("valor_total")
				.c3 = .c3 + vl_comissao
				end with
			ordena_cl_tres_colunas vL, 1, Ubound(vL)
			end if
		
		x = x & "	</TR>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
	  ' FECHA BLOCO DA "LOJA QUE INDICOU" DA ÚLTIMA LOJA
		s_cor="black"
		if (vl_subtotal_pedido<0) Or (vl_subtotal_comissao<0) then s_cor="red"
		x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
				"		<TD class='MDE' style='background:white;'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MC' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
				"		<TD class='MTD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_subtotal_pedido) & "</p></td>" & chr(13) & _
				"		<TD class='MTD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_subtotal_comissao) & "</p></td>" & chr(13) & _
				"	</TR>" & chr(13)

		s_cor="black"
		if (vl_total_pedido<0) Or (vl_total_comissao<0) then s_cor="red"
		x = x & "	<TR NOWRAP style='background: #FFF0E0'>" & chr(13) & _
				"		<TD class='MTBE' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
				"		<TD class='MTB' colspan='2'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_pedido) & "</p></td>" & chr(13) & _
				"		<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_comissao) & "</p></td>" & chr(13) & _
				"	</TR>" & chr(13)
				
	'>	TOTAL GERAL
		if qtde_loja > 1 then
			s_cor="black"
			if (vl_total_geral_pedido<0) Or (vl_total_geral_comissao<0) then s_cor="red"
			x = x & "	<TR>" & chr(13) & _
					"		<TD colspan='4' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD colspan='4' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
					"		<TD class='MTBE' colspan='2' NOWRAP><p class='Cd'>TOTAL GERAL:</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_geral_pedido) & "</p></td>" & chr(13) & _
					"		<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_geral_comissao) & "</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			
		'	TOTAL POR "LOJA QUE INDICOU"
			x = x & "	<TR>" & chr(13) & _
					"		<TD colspan='4' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD colspan='4' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</TR>" & chr(13)
					
			for i=Lbound(vL) to Ubound(vL)
				with vL(i)
					if Trim("" & .c1) <> "" then
						s_cor="black"
						if (.c2<0) Or (.c3<0) then s_cor="red"
						s = Trim("" & .c1)
						if IsNumeric(s) then s = normaliza_codigo(converte_numero(s), TAM_MIN_LOJA)
						s_aux = x_loja(s)
						if (s<>"") And (s_aux<>"") then s = s & " - "
						s = s & s_aux
						if s = "" then s = "&nbsp;"
						x = x & "	<TR>" & chr(13) & _
								"		<TD colspan='4' style='border-left:0px;border-right:0px;font-size:8pt;'>&nbsp;</td>" & chr(13) & _
								"	</TR>" & chr(13) & _
								"	<TR NOWRAP style='background:azure;'>" & chr(13) & _
								"		<TD class='MDTE' colspan='4' NOWRAP><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
								"	</TR>" & chr(13) & _
								"	<TR NOWRAP>" & chr(13) & _
								"		<TD class='MTBE' colspan='2' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
								"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.c2) & "</p></td>" & chr(13) & _
								"		<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.c3) & "</p></td>" & chr(13) & _
								"	</TR>" & chr(13)
						end if
					end with
				next
			
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT'><P class='ALERTA'>&nbsp;&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;&nbsp;</P></TD>" & chr(13) & _
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

function fRELConcluir( id_pedido ) {
	fREL.action = "pedido.asp";
	fREL.pedido_selecionado.value = id_pedido;
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
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Comissão de Loja por Indicação</span>
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
