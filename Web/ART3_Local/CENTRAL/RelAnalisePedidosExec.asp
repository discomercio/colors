<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L A N A L I S E P E D I D O S E X E C . A S P
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
	if Not operacao_permitida(OP_CEN_REL_ANALISE_PEDIDOS, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, s_filtro_fabricante, s_filtro_loja, flag_ok
	dim ckb_periodo_cadastro, c_dt_cadastro_inicio, c_dt_cadastro_termino
	dim ckb_periodo_entrega, c_dt_entrega_inicio, c_dt_entrega_termino
	dim ckb_produto, c_fabricante, c_produto
	dim c_lista_fabricante, s_lista_fabricante, v_fabricante
	dim c_lista_loja, s_lista_loja, v_loja
	dim c_cliente_cnpj_cpf
	dim rb_PF_PJ
	dim i, v, cadastrado
	
	alerta = ""

	ckb_periodo_cadastro = Trim(Request.Form("ckb_periodo_cadastro"))
	c_dt_cadastro_inicio = Trim(Request.Form("c_dt_cadastro_inicio"))
	c_dt_cadastro_termino = Trim(Request.Form("c_dt_cadastro_termino"))
	ckb_periodo_entrega = Trim(Request.Form("ckb_periodo_entrega"))
	c_dt_entrega_inicio = Trim(Request.Form("c_dt_entrega_inicio"))
	c_dt_entrega_termino = Trim(Request.Form("c_dt_entrega_termino"))
	ckb_produto = Trim(Request.Form("ckb_produto"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))

	c_lista_fabricante = Trim(Request.Form("c_lista_fabricante"))
	s_lista_fabricante = substitui_caracteres(c_lista_fabricante,chr(10),"")
	v_fabricante = split(s_lista_fabricante,chr(13),-1)

	c_lista_loja = Trim(Request.Form("c_lista_loja"))
	s_lista_loja = substitui_caracteres(c_lista_loja,chr(10),"")
	v_loja = split(s_lista_loja,chr(13),-1)

	c_cliente_cnpj_cpf=retorna_so_digitos(trim(request("c_cliente_cnpj_cpf")))
	
	rb_PF_PJ = Ucase(Trim(Request.Form("rb_PF_PJ")))
	
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
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
	'	PERÍODO DE CADASTRO
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
			
	'	PERÍODO DE ENTREGA
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_entrega_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_entrega_termino
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			if c_dt_entrega_inicio = "" then c_dt_entrega_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
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
dim s_where, s_from, s_where_loja, s_where_fornecedor
dim s_nome, s_cnpj_cpf, s_endereco, s_tel_res, s_tel_com, s_rg, s_email, s_email_xml
dim s_indicador, s_desempenho_nota
dim x, loja_a, qtde_lojas, pedido_a
dim w_fabricante, w_produto, w_descricao, w_qtde
dim w_preco_lista, w_desconto, w_vl_unitario, w_vl_total
dim st_pagto, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, s_cor, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, vl_saldo_a_pagar
dim vl_item, vl_total_item
dim vl_sub_total_familia, vl_total_geral_familia
dim vl_sub_total_pedido, vl_total_geral_pedido
dim mudou_loja, mudou_pedido
dim s_desc, perc_desc, vl_unitario, vl_lista

'	MONTA CLÁUSULA WHERE
	s_where = ""

'	CRITÉRIO: PERÍODO DE CADASTRAMENTO DO PEDIDO
	if ckb_periodo_cadastro <> "" then
		s = ""
		if c_dt_cadastro_inicio <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.data >= " & bd_formata_data(StrToDate(c_dt_cadastro_inicio)) & ")"
			end if
		
		if c_dt_cadastro_termino <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.data < " & bd_formata_data(StrToDate(c_dt_cadastro_termino)+1) & ")"
			end if
		
		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
		end if

'	CRITÉRIO: PERÍODO DE ENTREGA DO PEDIDO
	if ckb_periodo_entrega <> "" then
		s = ""
		if c_dt_entrega_inicio <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.entregue_data >= " & bd_formata_data(StrToDate(c_dt_entrega_inicio)) & ")"
			end if
		
		if c_dt_entrega_termino <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.entregue_data < " & bd_formata_data(StrToDate(c_dt_entrega_termino)+1) & ")"
			end if
		
		if s <> "" then s = s & " AND"
		s = s & " (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')"

		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
		end if
		
'	CRITÉRIO: PRODUTO
	if ckb_produto <> "" then
		s = ""
		if c_fabricante <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO_ITEM.fabricante = '" & c_fabricante & "')"
			end if
		
		if c_produto <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO_ITEM.produto = '" & c_produto & "')"
			end if
		
		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
		end if

'	CRITÉRIO: FORNECEDORES
	s_where_fornecedor = ""
	for i=Lbound(v_fabricante) to Ubound(v_fabricante)
		if v_fabricante(i) <> "" then
			v = split(v_fabricante(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_where_fornecedor <> "" then s_where_fornecedor = s_where_fornecedor & " OR"
				s_where_fornecedor = s_where_fornecedor & " (CONVERT(smallint, t_PEDIDO_ITEM.fabricante) = " & v_fabricante(i) & ")"
			else
				s = ""
				if v(Lbound(v))<>"" then 
					if s <> "" then s = s & " AND"
					s = s & " (CONVERT(smallint, t_PEDIDO_ITEM.fabricante) >= " & v(Lbound(v)) & ")"
					end if
				if v(Ubound(v))<>"" then
					if s <> "" then s = s & " AND"
					s = s & " (CONVERT(smallint, t_PEDIDO_ITEM.fabricante) <= " & v(Ubound(v)) & ")"
					end if
				if s <> "" then 
					if s_where_fornecedor <> "" then s_where_fornecedor = s_where_fornecedor & " OR"
					s_where_fornecedor = s_where_fornecedor & " (" & s & ")"
					end if
				end if
			end if
		next
		
	if s_where_fornecedor <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s_where_fornecedor & ")"
		end if

'	CRITÉRIO: LOJAS
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

'	CRITÉRIO: CLIENTE - CPF/CNPJ
	if c_cliente_cnpj_cpf <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_CLIENTE.cnpj_cpf = '" & retorna_so_digitos(c_cliente_cnpj_cpf) & "')"
		end if

'	CRITÉRIO: CLIENTE - PF/PJ
	if rb_PF_PJ = "PF_ON" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_CLIENTE.tipo = '" & ID_PF & "')"
		end if
		
	if rb_PF_PJ = "PJ_ON" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_CLIENTE.tipo = '" & ID_PJ & "')"
		end if

	if s_where <> "" then s_where = " WHERE" & s_where

	
'	MONTA CLÁUSULA FROM
	s_from = " FROM t_PEDIDO" & _
			 " INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			 " INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
			 " LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO.indicador = t_ORCAMENTISTA_E_INDICADOR.apelido)"

	s_sql = "SELECT t_PEDIDO.loja, CONVERT(smallint,t_PEDIDO.loja) AS numero_loja," & _
			" t_PEDIDO.data, t_PEDIDO.pedido, t_PEDIDO.st_entrega, t_PEDIDO.entregue_data, t_PEDIDO.cancelado_data," & _
			" t_CLIENTE.nome, t_CLIENTE.nome_iniciais_em_maiusculas, t_CLIENTE.tipo, t_CLIENTE.cnpj_cpf, t_CLIENTE.rg, t_CLIENTE.ie," & _
			" t_CLIENTE.ddd_res, t_CLIENTE.tel_res, t_CLIENTE.ddd_com, t_CLIENTE.tel_com, t_CLIENTE.ramal_com," & _
			" t_CLIENTE.endereco, t_CLIENTE.endereco_numero, t_CLIENTE.endereco_complemento, t_CLIENTE.bairro, t_CLIENTE.cidade, t_CLIENTE.uf, t_CLIENTE.cep, t_CLIENTE.email," & _
			" t_ORCAMENTISTA_E_INDICADOR.apelido AS indicador, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota," & _
			" t_PEDIDO_ITEM.fabricante, t_PEDIDO_ITEM.produto, t_PEDIDO_ITEM.descricao, t_PEDIDO_ITEM.descricao_html," & _
			" t_PEDIDO_ITEM.qtde, t_PEDIDO_ITEM.preco_lista, t_PEDIDO_ITEM.desc_dado, t_PEDIDO_ITEM.preco_NF, " & _
			" t_PEDIDO.st_memorizacao_completa_enderecos, t_CLIENTE.email_xml, t_CLIENTE.produtor_rural_status, t_CLIENTE.contribuinte_icms_status, " & _
			" t_PEDIDO.endereco_rg, t_PEDIDO.endereco_ie, t_PEDIDO.endereco_nome, t_PEDIDO.endereco_logradouro as pedido_endereco_logradouro, " & _
			" t_PEDIDO.endereco_numero as pedido_endereco_numero, t_PEDIDO.endereco_complemento as pedido_endereco_complemento, " & _
			" t_PEDIDO.endereco_bairro as pedido_endereco_bairro, t_PEDIDO.endereco_cidade as pedido_endereco_cidade, " & _
			" t_PEDIDO.endereco_uf as pedido_endereco_uf, t_PEDIDO.endereco_cep as pedido_endereco_cep, " & _
			" t_PEDIDO.endereco_tel_res, t_PEDIDO.endereco_ddd_res, t_PEDIDO.endereco_tel_com, t_PEDIDO.endereco_ddd_com, t_PEDIDO.endereco_ramal_com, " & _
			" t_PEDIDO.endereco_tel_cel, t_PEDIDO.endereco_ddd_cel, t_PEDIDO.endereco_tel_com_2, t_PEDIDO.endereco_ddd_com_2, t_PEDIDO.endereco_ramal_com_2, " & _
			" t_PEDIDO.endereco_email, t_PEDIDO.endereco_email_xml, t_PEDIDO.endereco_produtor_rural_status, t_PEDIDO.endereco_contribuinte_icms_status " & _
			s_from & _
			s_where
			
	s_sql = s_sql & " ORDER BY numero_loja, t_PEDIDO.data, t_PEDIDO.pedido"

  ' CABEÇALHO
	w_fabricante = 29
	w_produto = 54
	w_descricao = 279
	w_qtde = 26
	w_preco_lista = 69
	w_desconto = 40
	w_vl_unitario = 64
	w_vl_total = 79
	
	cab_table = "<TABLE cellSpacing=0 cellPadding=0>" & chr(13)
	cab = "	<tr style='background:mintcream;'>" & chr(13) & _
		  "		<td class='MDTE' style='width:" & cstr(w_fabricante) & "px' valign='bottom' NOWRAP><P class='PLTe'>Fabr</P></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_produto) & "px' valign='bottom'><P class='PLTe'>Produto</P></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_descricao) & "px' valign='bottom'><P class='PLTe'>Descrição</P></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_qtde) & "px' align='right' valign='bottom'><P class='PLTd'>Qtd</P></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_preco_lista) & "px' align='right' valign='bottom'><P class='PLTd'>Preço Lista</P></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_desconto) & "px' align='right' valign='bottom'><P class='PLTd'>Desc</P></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_vl_unitario) & "px' align='right' valign='bottom'><P class='PLTd'>Valor Unit</P></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_vl_total) & "px' align='right' valign='bottom'><P class='PLTd'>Valor Total</P></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_lojas = 0
	vl_sub_total_familia = 0
	vl_total_geral_familia = 0
	vl_sub_total_pedido = 0
	vl_total_geral_pedido = 0
	loja_a = "XXXXX"
	pedido_a = "XXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
		mudou_pedido = (Trim("" & r("pedido")) <> pedido_a)
		mudou_loja = (Trim("" & r("loja")) <> loja_a)
		
	'	IMPRIME TOTAL DO PEDIDO ANTERIOR?
		if mudou_pedido And (n_reg_total > 0) then
			x = x & "	<TR>" & chr(13) & _
					"		<TD COLSPAN='8' style='border-top:1px solid #C0C0C0;'>" & chr(13)& _
					"			<table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
					"				<tr>" & chr(13) & _
					"					<td width='20%' class='MDE' valign='bottom'><p class='Rf'>Status de Pagto</p></td>" & chr(13) & _
					"					<td width='20%' class='MD' align='right' valign='bottom'><p class='Rf'>VL Total&nbsp;&nbsp;(Família)&nbsp;</p></td>" & chr(13) & _
					"					<td width='20%' class='MD' align='right' valign='bottom'><p class='Rf'>VL Pago&nbsp;</p></td>" & chr(13) & _
					"					<td width='20%' class='MD' align='right' valign='bottom'><p class='Rf'>VL Devoluções&nbsp;</p></td>" & chr(13) & _
					"					<td width='20%' class='MD' align='right' valign='bottom'><p class='Rf'>Total dos Itens&nbsp;</p></td>" & chr(13) & _
					"				</tr>" & chr(13) & _
					"				<tr>" & chr(13) & _
					"					<td width='20%' class='MDE'><p class='C'>" & Ucase(x_status_pagto(st_pagto)) & "&nbsp;</p></td>" & chr(13) & _
					"					<td width='20%' class='MD' align='right'><p class='Cd'>" & formata_moeda(vl_TotalFamiliaPrecoNF) & "</p></td>" & chr(13)
				
			if vl_TotalFamiliaPago >= 0 then s_cor = "black" else s_cor = "red"
			x = x & "					<td width='20%' class='MD' align='right'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_TotalFamiliaPago) & "</p></td>" & chr(13) & _
					"					<td width='20%' class='MD' align='right'><p class='Cd'>" & formata_moeda(vl_TotalFamiliaDevolucaoPrecoNF) & "</p></td>" & chr(13) & _
					"					<td width='20%' class='MD' align='right'><p class='Cd'>" & formata_moeda(vl_total_item) & "</p></td>" & chr(13) & _
					"				</tr>" & chr(13) 
						
		'	IMPRIME TOTAL DA LOJA ANTERIOR?
			if mudou_loja then
				x = x & "			<tr>" & chr(13) & _
						"				<td ColSpan='5' class='MDTE'>&nbsp;</td>" & chr(13) & _
						"			</tr>" & chr(13) & _
						"			<tr style='background:azure;'>" & chr(13) & _
						"				<td class='MTBE'><p class='Cd'>Total da Loja " & loja_a & "</p></td>" & chr(13) & _
						"				<td class='MTB'><p class='Cd'>" & formata_moeda(vl_sub_total_familia) & "</p></td>" & chr(13) & _
						"				<td ColSpan='3' class='MTBD'><p class='Cd'>" & formata_moeda(vl_sub_total_pedido) & "</p></td>" & chr(13) & _
						"			</tr>" & chr(13) 
				end if
						
			x = x & "			</table>" & chr(13) & _
					"		</TD>" & chr(13) & _
					"	</TR>" & chr(13)
			Response.Write x
			x = ""
			end if


	'	IMPRIME CABEÇALHO DA LOJA?
		if mudou_loja then
			n_reg = 0
			vl_sub_total_familia = 0
			vl_sub_total_pedido = 0
			qtde_lojas = qtde_lojas + 1

			if n_reg_total = 0 then 
				x = cab_table
			elseif n_reg_total > 0 then
				x = x & "	<tr>" & chr(13) & _
						"		<td colspan='8'>&nbsp;</td>" & chr(13) & _
						"	</tr>" & chr(13) & _
						"	<tr>" & chr(13) & _
						"		<td colspan='8'>&nbsp;</td>" & chr(13) & _
						"	</tr>" & chr(13)
				end if
				
			s = Trim("" & r("loja"))
			s_aux = x_loja(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			if s <> "" then x = x & "	<TR>" & chr(13) & _
									"		<TD class='MDTE' COLSPAN='8' valign='bottom' class='MB' style='background:azure;'>" & chr(13) & _
									"			<p class='N' style='margin-left:8px;font-size:12pt;'>&nbsp;" & s & "</p>" & chr(13) & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13)
			end if
		
	'	IMPRIME DADOS DO CLIENTE?
		if mudou_pedido then
			vl_total_item = 0

		'	PULA LINHA ENTRE PEDIDOS
			if n_reg > 0 then
				x = x & "	<TR>" & chr(13) & _
						"		<TD class='MDTE' COLSPAN='8'>&nbsp;</td>" & chr(13) & _
						"	</tr>" & chr(13)
				end if
				
		'	OBTÉM OUTROS DADOS DO PEDIDO
			if Not calcula_pagamentos(Trim("" & r("pedido")), vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			vl_saldo_a_pagar = vl_TotalFamiliaPrecoNF-vl_TotalFamiliaPago-vl_TotalFamiliaDevolucaoPrecoNF
			vl_total_geral_familia = vl_total_geral_familia + vl_TotalFamiliaPrecoNF
			vl_sub_total_familia = vl_sub_total_familia + vl_TotalFamiliaPrecoNF
			
		'	Nº PEDIDO
			s = ""
			if Trim("" & r("st_entrega")) = ST_ENTREGA_ENTREGUE then
				s = formata_data(r("entregue_data"))
			elseif Trim("" & r("st_entrega")) = ST_ENTREGA_CANCELADO then
				s = formata_data(r("cancelado_data"))
				end if
			if s<>"" then s="  (" & s & ")"
			
			s_indicador = Trim("" & r("indicador"))
			if s_indicador = "" then s_indicador = "&nbsp;"
			
			s_desempenho_nota = Trim("" & r("desempenho_nota"))
			if s_desempenho_nota = "" then 
				s_desempenho_nota = "&nbsp;"
			else
				s_desempenho_nota = "(" & s_desempenho_nota & ")"
				end if
			
			x = x & "	<TR style='background:whitesmoke;'>" & chr(13) & _
					"		<TD class='MDTE' COLSPAN='8'>" & chr(13) & _
					"			<table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
					"				<tr>" & chr(13) & _
					"					<td style='width:80px;'>" & chr(13) & _
					"						<p class='Cn' style='font-weight:bold;margin-left:12px;'>" & _
											"<a href='javascript:fRELConcluir(" & chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & _
											Trim("" & r("pedido")) & "</a></p>" & chr(13) &_
					"					</td>" & chr(13) & _
					"					<td style='width:80px;'>" & chr(13) & _
					"						<p class='Cn'>" & formata_data(r("data")) & "</p>" & chr(13) & _
					"					</td>" & chr(13) & _
					"					<td style='width:200px;' align='left' class='MD'>" & chr(13) & _
					"						<p class='Cn'>" & x_status_entrega(r("st_entrega")) & s & "</p>" & chr(13) & _
					"					</td>" & chr(13) & _
					"					<td align='left'>" & chr(13) & _
					"						<p class='Cn'>" & s_indicador & " <span style='font-weight:bold;'>" & s_desempenho_nota & "</span></p>" & chr(13) & _
					"					</td>" & chr(13) & _
					"				</tr>" & chr(13) & _
					"			</table>" & chr(13) & _
					"		</td>" & chr(13) & _
					"	</tr>" & chr(13)
			
		'	DADOS DO CLIENTE
			dim cliente__nome_iniciais_em_maiusculas, cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento
			dim cliente__bairro, cliente__cidade, cliente__uf, cliente__cep, cliente__cnpj_cpf
			dim cliente__tel_res, cliente__ddd_res, cliente__tel_com, cliente__ddd_com, cliente__ramal_com
			dim cliente__ie, cliente__rg, cliente__email, cliente__email_xml
			cliente__nome_iniciais_em_maiusculas = Trim("" & r("nome_iniciais_em_maiusculas"))
			cliente__endereco = Trim("" & r("endereco"))
			cliente__endereco_numero = Trim("" & r("endereco_numero"))
			cliente__endereco_complemento = Trim("" & r("endereco_complemento"))
			cliente__bairro = Trim("" & r("bairro"))
			cliente__cidade = Trim("" & r("cidade"))
			cliente__uf = Trim("" & r("uf"))
			cliente__cep = Trim("" & r("cep"))
			cliente__cnpj_cpf = Trim("" & r("cnpj_cpf"))
			cliente__tel_res = Trim("" & r("tel_res"))
			cliente__ddd_res = Trim("" & r("ddd_res"))
			cliente__tel_com = Trim("" & r("tel_com"))
			cliente__ddd_com = Trim("" & r("ddd_com"))
			cliente__ramal_com = Trim("" & r("ramal_com"))
			cliente__ie = Trim("" & r("ie"))
			cliente__rg = Trim("" & r("rg"))
			cliente__email = Trim("" & r("email"))
			cliente__email_xml = Trim("" & r("email_xml"))
			if Trim("" & r("st_memorizacao_completa_enderecos")) <> 0 then
				cliente__nome_iniciais_em_maiusculas = iniciais_em_maiusculas(Trim("" & r("endereco_nome")))
				cliente__endereco = Trim("" & r("pedido_endereco_logradouro"))
				cliente__endereco_numero = Trim("" & r("pedido_endereco_numero"))
				cliente__endereco_complemento = Trim("" & r("pedido_endereco_complemento"))
				cliente__bairro = Trim("" & r("pedido_endereco_bairro"))
				cliente__cidade = Trim("" & r("pedido_endereco_cidade"))
				cliente__uf = Trim("" & r("pedido_endereco_uf"))
				cliente__cep = Trim("" & r("pedido_endereco_cep"))
				cliente__cnpj_cpf = Trim("" & r("cnpj_cpf")) ' este usamos o principal porque sempre é igual
				cliente__tel_res = Trim("" & r("endereco_tel_res"))
				cliente__ddd_res = Trim("" & r("endereco_ddd_res"))
				cliente__tel_com = Trim("" & r("endereco_tel_com"))
				cliente__ddd_com = Trim("" & r("endereco_ddd_com"))
				cliente__ramal_com = Trim("" & r("endereco_ramal_com"))
				cliente__ie = Trim("" & r("endereco_ie"))
				cliente__rg = Trim("" & r("endereco_rg"))
				cliente__email = Trim("" & r("endereco_email"))
				cliente__email_xml = Trim("" & r("endereco_email_xml"))
				end if

		'	nome
			s_nome = "&nbsp;"
			if cliente__nome_iniciais_em_maiusculas <> "" then s_nome = cliente__nome_iniciais_em_maiusculas
		'	endereço
			s_endereco = "&nbsp;"
			if cliente__endereco <> "" then
				s_endereco = iniciais_em_maiusculas(cliente__endereco)
				s = cliente__endereco_numero
				if s<>"" then s_endereco=s_endereco & ", " & s
				s = cliente__endereco_complemento
				if s<>"" then s_endereco=s_endereco & " " & s
				s = iniciais_em_maiusculas(cliente__bairro)
				if s<>"" then s_endereco=s_endereco & " - " & s
				s = iniciais_em_maiusculas(cliente__cidade)
				if s<>"" then s_endereco=s_endereco & " - " & s
				s=UCase(cliente__uf)
				if s<>"" then s_endereco=s_endereco & " - " & s
				s=cliente__cep
				if s<>"" then s_endereco=s_endereco & " - " & cep_formata(s)
				end if
		'	cnpj/cpf
			s_cnpj_cpf = "CPF: "
			if cliente__cnpj_cpf <> "" then
				s_cnpj_cpf = cnpj_cpf_formata(cliente__cnpj_cpf)
				if Len(cliente__cnpj_cpf) = 14 then
					s_cnpj_cpf = "CNPJ: " & s_cnpj_cpf
				else
					s_cnpj_cpf = "CPF: " & s_cnpj_cpf
					end if
				end if
		'	telefone residencial
			s_tel_res = ""
			if cliente__tel_res <> "" then
				s = cliente__tel_res
				s_tel_res = telefone_formata(s)
				s = cliente__ddd_res
				if s <> "" then s_tel_res = "(" & s & ") " & s_tel_res
				end if
			s_tel_res = "Tel Res: " & s_tel_res
		'	telefone comercial
			s_tel_com = ""
			if cliente__tel_com <> "" then
				s = cliente__tel_com
				s_tel_com = telefone_formata(s)
				s = cliente__ddd_com
				if s <> "" then s_tel_com = "(" & s & ") " & s_tel_com
				s = cliente__ramal_com
				if s<>"" then s_tel_com = s_tel_com & "  (R." & s & ")"
				end if
			s_tel_com = "Tel Com: " & s_tel_com
		'	rg
			s_rg = "&nbsp;"
			if Trim("" & r("tipo")) = ID_PJ then
				if cliente__ie <> "" then s_rg = "IE: " & cliente__ie
			else
				if cliente__rg <> "" then s_rg = "RG: " & cliente__rg
				end if
		'	e-mail
			s_email = ""
			if cliente__email <> "" then
				s_email = "E-mail: " & cliente__email
				end if
		'	e-mail-xml
			s_email_xml = ""
			if cliente__email_xml <> "" then
				s_email_xml = "E-mail (XML): " & cliente__email_xml
				end if
		'	concatena e-mail e e-mail-xml
			if s_email = "" then
				s_email = s_email_xml
			else
				if s_email_xml <> "" then
					s_email = s_email  & " - " & s_email_xml
					end if
				end if
			if s_email = "" then
				s_email = "E-mail: "
				end if

			x = x & "	<TR>" & chr(13) & _
					"		<TD class='MDTE' COLSPAN='8'>" & chr(13) & _
					"			<table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
					"				<tr>" & chr(13) & _
					"					<td colspan='2' class='MD' valign='bottom'><P class='Cn'>" & s_nome & "</p></td>" & chr(13) & _
					"					<td valign='bottom'><P class='Cn'>" & s_cnpj_cpf & "</p></td>" & chr(13) & _
					"				</tr>" & chr(13) & _
					"				<tr>" & chr(13) & _
					"					<td colspan='3' class='MC' valign='bottom'><P class='Cn'>" & s_endereco & "</p></td>" & chr(13) & _
					"				</tr>" & chr(13) & _
					"				<tr>" & chr(13) & _
					"					<td width='34%' class='MTD' valign='bottom'><P class='Cn'>" & s_tel_res & "</p></td>" & chr(13) & _
					"					<td width='33%' class='MTD' valign='bottom'><P class='Cn'>" & s_tel_com & "</p></td>" & chr(13) & _
					"					<td width='33%' class='MC' valign='bottom'><P class='Cn'>" & s_rg & "</p></td>" & chr(13) & _
					"				</tr>" & chr(13) & _
					"				<tr>" & chr(13) & _
					"					<td colspan='3' class='MC' valign='bottom'><P class='Cn'>" & s_email & "</p></td>" & chr(13) & _
					"				</tr>" & chr(13) & _
					"			</table>" & chr(13) & _
					"		</td>" & chr(13) & _
					"	</tr>" & chr(13)
		
		'	CABEÇALHO C/ TÍTULOS DOS ITENS DO PEDIDO
			x = x & cab
			end if
		
	'	MEMORIZA PEDIDO ATUAL
		if mudou_pedido then
			mudou_pedido = False
			pedido_a = Trim("" & r("pedido"))
			end if
		
	'	MEMORIZA LOJA ATUAL
		if mudou_loja then
			mudou_loja = False
			loja_a = Trim("" & r("loja"))
			end if
		
	 ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR>" & chr(13)


	'> FABRICANTE
		s = Trim("" & r("fabricante"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='bottom' class='MDTE'>" & chr(13) & _
				"			<P class='Cn'>" & s & "</P>" & chr(13) & _
				"		</TD>" & chr(13)

	'> PRODUTO
		s = Trim("" & r("produto"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='bottom' class='MTD'>" & chr(13) & _
				"			<P class='Cn'>" & s & "</P>" & chr(13) & _
				"		</TD>" & chr(13)

	'> DESCRIÇÃO DO PRODUTO
		s = produto_formata_descricao_em_html(Trim("" & r("descricao_html")))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='bottom' class='MTD'>" & chr(13) & _
				"			<P class='Cn'>" & s & "</P>" & chr(13) & _
				"		</TD>" & chr(13)

	'> QUANTIDADE
		s = formata_inteiro(r("qtde"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='bottom' class='MTD'>" & chr(13) & _
				"			<P class='Cnd'>" & s & "</P>" & chr(13) & _
				"		</TD>" & chr(13)

	'> PREÇO DE LISTA
		vl_lista=r("preco_lista")
		s = formata_moeda(vl_lista)
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='bottom' class='MTD'>" & chr(13) & _
				"			<P class='Cnd'>" & s & "</P>" & chr(13) & _
				"		</TD>" & chr(13)

	'> DESCONTO
		vl_unitario=r("preco_NF")
			
		if vl_unitario < vl_lista then
			if vl_lista > 0 then
				perc_desc=100*((vl_lista-vl_unitario)/vl_lista)
			else
				perc_desc=0
			end if
			s_desc = formata_perc(perc_desc)
		else
			perc_desc=0
			s_desc=""
			end if
			
		s = s_desc
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='bottom' class='MTD'>" & chr(13) & _
				"			<P class='Cnd'>" & s & "</P>" & chr(13) & _
				"		</TD>" & chr(13)

	'> VALOR UNITÁRIO
		s = formata_moeda(vl_unitario)
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='bottom' class='MTD'>" & chr(13) &_
				"			<P class='Cnd'>" & s & "</P>" & chr(13) &_
				"		</TD>" & chr(13)

	'> VALOR TOTAL DO ITEM
		vl_item = r("qtde") * vl_unitario
		s = formata_moeda(vl_item)
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='bottom' class='MTD'>" & chr(13) &_
				"			<P class='Cnd'>" & s & "</P>" & chr(13) & _
				"		</TD>" & chr(13)

	'> TOTALIZAÇÃO DE VALORES
		vl_total_item = vl_total_item + vl_item
		vl_sub_total_pedido = vl_sub_total_pedido + vl_item
		vl_total_geral_pedido = vl_total_geral_pedido + vl_item
			
		x = x & "	</TR>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
		
  ' MOSTRA TOTAL DO ÚLTIMO BLOCO
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO PEDIDO
		x = x & "	<TR>" & chr(13) & _
				"		<TD COLSPAN='8' style='border-top:1px solid #C0C0C0;'>" & chr(13)& _
				"			<table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
				"				<tr>" & chr(13) & _
				"					<td width='20%' class='MDE' valign='bottom'><p class='Rf'>Status de Pagto</p></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right' valign='bottom'><p class='Rf'>VL Total&nbsp;&nbsp;(Família)&nbsp;</p></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right' valign='bottom'><p class='Rf'>VL Pago&nbsp;</p></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right' valign='bottom'><p class='Rf'>VL Devoluções&nbsp;</p></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right' valign='bottom'><p class='Rf'>Total dos Itens&nbsp;</p></td>" & chr(13) & _
				"				</tr>" & chr(13) & _
				"				<tr>" & chr(13) & _
				"					<td width='20%' class='MDE'><p class='C'>" & Ucase(x_status_pagto(st_pagto)) & "&nbsp;</p></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right'><p class='Cd'>" & formata_moeda(vl_TotalFamiliaPrecoNF) & "</p></td>" & chr(13)
		
		if vl_TotalFamiliaPago >= 0 then s_cor = "black" else s_cor = "red"
		x = x & "					<td width='20%' class='MD' align='right'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_TotalFamiliaPago) & "</p></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right'><p class='Cd'>" & formata_moeda(vl_TotalFamiliaDevolucaoPrecoNF) & "</p></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right'><p class='Cd'>" & formata_moeda(vl_total_item) & "</p></td>" & chr(13) & _
				"				</tr>" & chr(13) 

	'	TOTAL DA ÚLTIMA LOJA
		x = x & "				<tr>" & chr(13) & _
				"					<td ColSpan='5' class='MDTE'>&nbsp;</td>" & chr(13) & _
				"				</tr>" & chr(13) & _
				"				<tr style='background:azure;'>" & chr(13) & _
				"					<td class='MTBE'><p class='Cd'>Total da Loja " & loja_a & "</p></td>" & chr(13) & _
				"					<td class='MTB'><p class='Cd'>" & formata_moeda(vl_sub_total_familia) & "</p></td>" & chr(13) & _
				"					<td ColSpan='3' class='MTBD'><p class='Cd'>" & formata_moeda(vl_sub_total_pedido) & "</p></td>" & chr(13) & _
				"				</tr>" & chr(13) 
		
	'>	TOTAL GERAL
		if qtde_lojas > 1 then
			x = x & "			<tr>" & chr(13) & _
					"				<td ColSpan='5' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"			</tr>" & chr(13) & _
					"			<tr>" & chr(13) & _
					"				<td ColSpan='5' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"			</tr>" & chr(13) & _
					"			<tr style='background:honeydew'>" & chr(13) & _
					"				<td class='MTBE'><p class='Cd'>TOTAL GERAL:</p></td>" & chr(13) & _
					"				<td class='MTB'><p class='Cd'>" & formata_moeda(vl_total_geral_familia) & "</p></td>" & chr(13) & _
					"				<td ColSpan='3' class='MTBD'><p class='Cd'>" & formata_moeda(vl_total_geral_pedido) & "</p></td>" & chr(13) & _
					"			</tr>" & chr(13)
			end if

		x = x & "			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if


  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR>" & chr(13) & _
				"		<TD class='MT' colspan='8'><P class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
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

function fRELConcluir( id_pedido ){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=id_pedido;
	fREL.action = "pedido.asp"
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
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Análise de Pedidos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

	if ckb_periodo_cadastro <> "" then
		s = ""
		s_aux = c_dt_cadastro_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " e "
		s_aux = c_dt_cadastro_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Pedidos colocados entre:&nbsp;</p></td><td valign='top' width='99%'>" & _
				   "<p class='N'>" & s & "</p></td></tr>" & chr(13)
		end if

	if ckb_periodo_entrega <> "" then
		s = ""
		s_aux = c_dt_entrega_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " e "
		s_aux = c_dt_entrega_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Pedidos entregues entre:&nbsp;</p></td><td valign='top' width='99%'>" & _
				   "<p class='N'>" & s & "</p></td></tr>" & chr(13)
		end if

	if ckb_produto <> "" then 
		s_aux = c_fabricante
		if s_aux = "" then s_aux = "todos"
		s = "fabricante: " & s_aux
		s_aux = c_produto
		if s_aux = "" then s_aux = "todos"
		s = s & ",&nbsp;&nbsp;produto: " & s_aux
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Somente produto:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>" & chr(13)
		end if

'	LISTA DE FORNECEDORES
	s_filtro_fabricante = ""
	for i = Lbound(v_fabricante) to Ubound(v_fabricante)
		if v_fabricante(i) <> "" then
			v = split(v_fabricante(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_filtro_fabricante <> "" then s_filtro_fabricante = s_filtro_fabricante & ", "
				s_filtro_fabricante = s_filtro_fabricante & v_fabricante(i)
			else
				if (v(Lbound(v))<>"") And (v(Ubound(v))<>"") then 
					if s_filtro_fabricante <> "" then s_filtro_fabricante = s_filtro_fabricante & ", "
					s_filtro_fabricante = s_filtro_fabricante & v(Lbound(v)) & " a " & v(Ubound(v))
				elseif (v(Lbound(v))<>"") And (v(Ubound(v))="") then
					if s_filtro_fabricante <> "" then s_filtro_fabricante = s_filtro_fabricante & ", "
					s_filtro_fabricante = s_filtro_fabricante & v(Lbound(v)) & " e acima"
				elseif (v(Lbound(v))="") And (v(Ubound(v))<>"") then
					if s_filtro_fabricante <> "" then s_filtro_fabricante = s_filtro_fabricante & ", "
					s_filtro_fabricante = s_filtro_fabricante & v(Ubound(v)) & " e abaixo"
					end if
				end if
			end if
		next
	s = s_filtro_fabricante
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Fornecedor(es):&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>" & chr(13)

'	LISTA DE LOJAS
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
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Loja(s):&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>" & chr(13)

'	CLIENTE (CNPJ OU CPF)
	if c_cliente_cnpj_cpf <> "" then
		s = cnpj_cpf_formata(c_cliente_cnpj_cpf)
		s_aux = x_cliente_por_cnpj_cpf(c_cliente_cnpj_cpf, cadastrado)
		if Not cadastrado then s_aux = "Não Cadastrado"
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
					"<p class='N'>Cliente:&nbsp;</p></td><td valign='top'>" & _
					"<p class='N'>" & s & "</p></td></tr>" & chr(13)
		end if
		
'	PF OU PJ
	select case rb_PF_PJ
		case "PF_PJ_ON": s = "todos"
		case "PF_ON": s = "apenas PF"
		case "PJ_ON": s = "apenas PJ"
		case else: s = ""
		end select
		
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>PF/PJ:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>" & chr(13)
		
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
