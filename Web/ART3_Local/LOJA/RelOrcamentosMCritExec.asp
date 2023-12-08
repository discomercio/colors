<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L O R C A M E N T O S M C R I T E X E C . A S P
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
	if (Not operacao_permitida(OP_LJA_CONSULTA_ORCAMENTO, s_lista_operacoes_permitidas)) And _
	   (Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas)) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	dim s, s_aux, s_filtro, flag_ok, cadastrado
	dim ckb_orcamento_em_aberto, ckb_orcamento_virou_pedido, ckb_orcamento_cancelado
	dim ckb_periodo_cadastro, c_dt_cadastro_inicio, c_dt_cadastro_termino
	dim ckb_produto, c_fabricante, c_produto
	dim c_orcamento
	dim c_cliente_cnpj_cpf

	alerta = ""

	ckb_orcamento_em_aberto = Trim(Request.Form("ckb_orcamento_em_aberto"))
	ckb_orcamento_virou_pedido = Trim(Request.Form("ckb_orcamento_virou_pedido"))
	ckb_orcamento_cancelado = Trim(Request.Form("ckb_orcamento_cancelado"))
	ckb_periodo_cadastro = Trim(Request.Form("ckb_periodo_cadastro"))
	c_dt_cadastro_inicio = Trim(Request.Form("c_dt_cadastro_inicio"))
	c_dt_cadastro_termino = Trim(Request.Form("c_dt_cadastro_termino"))
	ckb_produto = Trim(Request.Form("ckb_produto"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_orcamento = normaliza_num_orcamento(ucase(Trim(Request.Form("c_orcamento"))))
	c_cliente_cnpj_cpf=retorna_so_digitos(trim(request("c_cliente_cnpj_cpf")))
	
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
		if c_cliente_cnpj_cpf <> "" then
			if Not cnpj_cpf_ok(c_cliente_cnpj_cpf) then
				alerta=texto_add_br(alerta)
				alerta = alerta & "CNPJ/CPF do cliente é inválido."
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
dim s, s_aux, s_sql, cab_table, cab, n_reg, n_reg_total
dim s_where, s_from
dim vl_total, vl_sub_total, vl_total_NF, vl_sub_total_NF
dim x, loja_a, qtde_lojas
dim w_cliente, w_valor
dim intNumLinha

'	MONTA CLÁUSULA WHERE
	s_where = ""

'	TEM ACESSO A TODOS OS ORÇAMENTOS?
	if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
		s = "t_ORCAMENTO.vendedor = '" & usuario & "'"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRITÉRIO: STATUS DO ORÇAMENTO
	s = ""
	s_aux = ckb_orcamento_em_aberto
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & "(" & _
					" ((t_ORCAMENTO.st_orcamento='') OR (t_ORCAMENTO.st_orcamento IS NULL))" & _
					" AND" & _
					"(t_ORCAMENTO.st_orc_virou_pedido = 0)" & _
				")"
		end if

	s_aux = ckb_orcamento_virou_pedido
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_ORCAMENTO.st_orc_virou_pedido = 1)"
		end if

	s_aux = ckb_orcamento_cancelado
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_ORCAMENTO.st_orcamento = '" & ST_ORCAMENTO_CANCELADO & "')"
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
		
'	CRITÉRIO: PERÍODO DE CADASTRAMENTO DO ORÇAMENTO
	s = ""
	if c_dt_cadastro_inicio <> "" then
		if s <> "" then s = s & " AND"
		s = s & " (t_ORCAMENTO.data >= " & bd_formata_data(StrToDate(c_dt_cadastro_inicio)) & ")"
		end if
		
	if c_dt_cadastro_termino <> "" then
		if s <> "" then s = s & " AND"
		s = s & " (t_ORCAMENTO.data < " & bd_formata_data(StrToDate(c_dt_cadastro_termino)+1) & ")"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CRITÉRIO: PRODUTO
	if ckb_produto <> "" then
		s = ""
		if c_fabricante <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_ORCAMENTO_ITEM.fabricante = '" & c_fabricante & "')"
			end if
		
		if c_produto <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_ORCAMENTO_ITEM.produto = '" & c_produto & "')"
			end if
		
		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
		end if

'	CRITÉRIO: LOJA (CADA LOJA SÓ PODE CONSULTAR ORÇAMENTOS DA PRÓPRIA LOJA)
	s = " (CONVERT(smallint, t_ORCAMENTO.loja) = " & loja & ")"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	CRITÉRIO: Nº ORÇAMENTO
	if c_orcamento <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ORCAMENTO.orcamento = '" & c_orcamento & "')"
		end if

'	CRITÉRIO: CLIENTE
	if c_cliente_cnpj_cpf <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
			s_where = s_where & " (t_ORCAMENTO.endereco_cnpj_cpf = '" & retorna_so_digitos(c_cliente_cnpj_cpf) & "')"
		else
			s_where = s_where & " (t_CLIENTE.cnpj_cpf = '" & retorna_so_digitos(c_cliente_cnpj_cpf) & "')"
			end if
		end if
		
	
'	CLÁUSULA WHERE
	if s_where <> "" then s_where = " WHERE" & s_where
	
	
'	MONTA CLÁUSULA FROM
	s_from = " FROM t_ORCAMENTO INNER JOIN t_ORCAMENTO_ITEM ON (t_ORCAMENTO.orcamento=t_ORCAMENTO_ITEM.orcamento)"
	
	if c_cliente_cnpj_cpf <> "" then
		s_from = s_from & " INNER JOIN t_CLIENTE ON (t_ORCAMENTO.id_cliente=t_CLIENTE.id)"
	else
		s_from = s_from & " LEFT JOIN t_CLIENTE ON (t_ORCAMENTO.id_cliente=t_CLIENTE.id)"
		end if

	s_sql = "SELECT DISTINCT t_ORCAMENTO.loja, CONVERT(smallint,t_ORCAMENTO.loja) AS numero_loja," & _
			" t_ORCAMENTO.data, t_ORCAMENTO.nsu, t_ORCAMENTO.orcamento,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_ORCAMENTO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_ORCAMENTO.st_orcamento," & _
			" t_ORCAMENTO.st_orc_virou_pedido, t_ORCAMENTO.pedido," & _
			" Sum(t_ORCAMENTO_ITEM.qtde*t_ORCAMENTO_ITEM.preco_venda) AS valor_total," & _
			" Sum(t_ORCAMENTO_ITEM.qtde*t_ORCAMENTO_ITEM.preco_NF) AS valor_total_NF" & _
			s_from & _
			s_where

	s_sql = s_sql & " GROUP BY t_ORCAMENTO.loja, t_ORCAMENTO.data, t_ORCAMENTO.nsu, t_ORCAMENTO.orcamento,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_ORCAMENTO.endereco_nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_ORCAMENTO.st_orcamento," & _
			" t_ORCAMENTO.st_orc_virou_pedido, t_ORCAMENTO.pedido"

	s_sql = s_sql & " ORDER BY numero_loja, t_ORCAMENTO.data, t_ORCAMENTO.nsu, t_ORCAMENTO.orcamento"

  ' CABEÇALHO
	w_cliente = 250
	w_valor = 80
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure'>" & chr(13) & _
		  "		<TD valign='bottom' style='background:white;' NOWRAP>&nbsp;</TD>" & chr(13) & _
		  "		<TD class='MDTE' style='width:70px' valign='bottom' NOWRAP><P class='R'>Pré-Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:" & Cstr(w_cliente) & "px' valign='bottom'><P class='R'>Cliente</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>VL Total</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>VL Total (RA)</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:70px' valign='bottom'><P class='R'>Status do Pré-Pedido</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_lojas = 0
	vl_sub_total = 0
	vl_sub_total_NF = 0
	vl_total = 0
	vl_total_NF = 0
	intNumLinha = 0
	
	loja_a = "XXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE LOJA?
		if Trim("" & r("loja"))<>loja_a then
			loja_a = Trim("" & r("loja"))
			qtde_lojas = qtde_lojas + 1
		  ' FECHA TABELA DA LOJA ANTERIOR
			if n_reg > 0 then 
				x = x & "	<TR style='background: #FFFFDD'>" & chr(13) & _
						"		<TD valign='bottom' style='background:white;' NOWRAP>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MTBE' colspan='2' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
						"		<TD class='MTB'><p class='Cd'>" & formata_moeda(vl_sub_total) & "</p></td>" & chr(13) & _
						"		<TD class='MTB'><p class='Cd'>" & formata_moeda(vl_sub_total_NF) & "</p></td>" & chr(13) & _
						"		<TD class='MTBD'><p class='C'>&nbsp;</p></td>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			n_reg = 0
			vl_sub_total = 0
			vl_sub_total_NF = 0

			if n_reg_total > 0 then x = x & "<BR>" & chr(13)
			s = Trim("" & r("loja"))
			s_aux = x_loja(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then 
				x = x & _
					"	<TR>" & chr(13) & _
					"		<TD valign='bottom' style='background:white;' NOWRAP>&nbsp;</TD>" & chr(13) & _
					"		<TD class='MDTE' colspan='5' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
				end if
			x = x & cab
			end if

	 ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1
		intNumLinha = intNumLinha + 1

		x = x & "	<TR>" & chr(13)

	'> Nº DA LINHA
		x = x & "		<TD valign='top' align='right' NOWRAP><P class='Rd' style='margin-right:2px;'>" & Cstr(intNumLinha) & ".</P></TD>" & chr(13)

	'> Nº ORÇAMENTO
		x = x & "		<TD valign='top' class='MDTE'><P class='C'>&nbsp;<a href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("orcamento")) & chr(34) & _
				")' title='clique para consultar o pré-pedido'>" & Trim("" & r("orcamento")) & "</a></P></TD>" & chr(13)

	'> CLIENTE
		s = Trim("" & r("nome_iniciais_em_maiusculas"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='top' style='width:" & Cstr(w_cliente) & "px' class='MTD'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> VALOR DO ORÇAMENTO
		s = formata_moeda(r("valor_total"))
		x = x & "		<TD valign='top' align='right' style='width:" & Cstr(w_valor) & "px' class='MTD'><P class='Cnd'>" & s & "</P></TD>" & chr(13)

	'> VALOR DO ORÇAMENTO COM RA
		s = formata_moeda(r("valor_total_NF"))
		x = x & "		<TD valign='top' align='right' style='width:" & Cstr(w_valor) & "px' class='MTD'><P class='Cnd'>" & s & "</P></TD>" & chr(13)

	'> STATUS DO ORÇAMENTO
		s = Trim("" & r("st_orcamento"))
		if s <> "" then s = x_st_orcamento(s)
		if s = "" then
			if r("st_orc_virou_pedido") = 1 then s = "Virou pedido"
			end if
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='top' class='MTD'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> TOTALIZAÇÃO DE VALORES
		vl_sub_total = vl_sub_total + r("valor_total")
		vl_total = vl_total + r("valor_total")

		vl_sub_total_NF = vl_sub_total_NF + r("valor_total_NF")
		vl_total_NF = vl_total_NF + r("valor_total_NF")
			
		x = x & "	</TR>" & chr(13)
		
		r.MoveNext
		loop

	
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR style='background: #FFFFDD'>" & chr(13) & _
				"		<TD valign='bottom' style='background:white;' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='2' class='MTBE' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd'>" & formata_moeda(vl_sub_total) & "</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd'>" & formata_moeda(vl_sub_total_NF) & "</p></td>" & chr(13) & _
				"		<TD class='MTBD'><p class='C'>&nbsp;</p></td>" & chr(13) & _
				"	</TR>" & chr(13)
		
	'>	TOTAL GERAL
		if qtde_lojas > 1 then
			x = x & _
				"	<TR>" & chr(13) & _
				"		<TD colspan='6' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR>" & chr(13) & _
				"		<TD colspan='6' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:honeydew'>" & chr(13) & _
				"		<TD valign='bottom' style='background:white;' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MTBE' colspan='2' NOWRAP><p class='Cd'>TOTAL GERAL:</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd'>" & formata_moeda(vl_total) & "</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd'>" & formata_moeda(vl_total_NF) & "</p></td>" & chr(13) & _
				"		<TD class='MTBD'><p class='C'>&nbsp;</p></td>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR>" & chr(13) & _
				"		<TD valign='bottom' style='background:white;' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MT' colspan='5'><P class='ALERTA'>&nbsp;NENHUM PRÉ-PEDIDO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir( id_orcamento ){
	window.status = "Aguarde ...";
	fREL.orcamento_selecionado.value=id_orcamento;
	fREL.action = "orcamento.asp"
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
<input type="hidden" name="orcamento_selecionado" id="orcamento_selecionado" value="">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório Multicritério de Pré-Pedidos</span>
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
	s_aux = ckb_orcamento_em_aberto
	if s_aux<>"" then
	'	DEVIDO AO WORD WRAP: SÓ FAZ WORD WRAP QUANDO ENCONTRA CHR(32), OU SEJA, MANTÉM AGRUPADO TEXTO COM &nbsp;
		s_aux = "em aberto"
		if s <> "" then s = s & ",&nbsp; "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = ckb_orcamento_virou_pedido
	if s_aux<>"" then
		s_aux = "virou pedido"
		if s <> "" then s = s & ",&nbsp; "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = ckb_orcamento_cancelado
	if s_aux<>"" then
		s_aux = "cancelado"
		if s <> "" then s = s & ",&nbsp; "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Status do Pré-Pedido:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if
	
	if (c_dt_cadastro_inicio <> "") Or (c_dt_cadastro_termino <> "") then
		s = ""
		s_aux = c_dt_cadastro_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " e "
		s_aux = c_dt_cadastro_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Pré-Pedidos cadastrados entre:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if ckb_produto <> "" then 
		s_aux = c_fabricante
		if s_aux = "" then s_aux = "todos"
		s = "fabricante: " & s_aux
		s_aux = c_produto
		if s_aux = "" then s_aux = "todos"
		s = s & ",&nbsp;&nbsp;produto: " & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Somente Pré-Pedidos que incluam:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if c_orcamento <> "" then
		s = c_orcamento
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Nº Pré-Pedido:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if
	
	if c_cliente_cnpj_cpf <> "" then
		s = cnpj_cpf_formata(c_cliente_cnpj_cpf)
		s_aux = x_cliente_por_cnpj_cpf(c_cliente_cnpj_cpf, cadastrado)
		if Not cadastrado then s_aux = "Não Cadastrado"
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Cliente:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if
	
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
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
