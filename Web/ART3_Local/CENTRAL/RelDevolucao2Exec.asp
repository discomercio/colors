<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelDevolucao2Exec.asp
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
	if Not operacao_permitida(OP_CEN_REL_DEVOLUCAO_PRODUTOS2, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	dim s, s_aux, s_filtro, s_filtro_loja, flag_ok, s_filtro_operacao
	dim ckb_periodo_devolucao, c_dt_devolucao_inicio, c_dt_devolucao_termino
	dim c_fabricante, c_produto, c_pedido
	dim c_vendedor, c_indicador, c_captador, s_nome_vendedor
	dim c_lista_loja, s_lista_loja, v_loja, v, i

	alerta = ""

	dim origem
	origem = ucase(Trim(request("origem")))

	if origem="A" then
	'	PARÂMETROS INFORMADOS PELA QUERYSTRING
		ckb_periodo_devolucao = Request("ckb_periodo_devolucao")
		c_dt_devolucao_inicio = Request("c_dt_devolucao_inicio")
		c_dt_devolucao_termino = Request("c_dt_devolucao_termino")
		c_fabricante = Request("c_fabricante")
		c_produto = Request("c_produto")
		c_pedido = Request("c_pedido")
		c_vendedor = Request("c_vendedor")
		c_indicador = Request("c_indicador")
		c_captador = Request("c_captador")
		c_lista_loja = Request("c_lista_loja")
	else
		ckb_periodo_devolucao = Trim(Request.Form("ckb_periodo_devolucao"))
		c_dt_devolucao_inicio = Trim(Request.Form("c_dt_devolucao_inicio"))
		c_dt_devolucao_termino = Trim(Request.Form("c_dt_devolucao_termino"))
		c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
		c_produto = Ucase(Trim(Request.Form("c_produto")))
		c_pedido = Ucase(Trim(Request.Form("c_pedido")))
		c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
		c_indicador = Ucase(Trim(Request.Form("c_indicador")))
		c_captador = Ucase(Trim(Request.Form("c_captador")))
		c_lista_loja = Trim(Request.Form("c_lista_loja"))
		end if
		
	s_lista_loja = substitui_caracteres(c_lista_loja,chr(10),"")
	v_loja = split(s_lista_loja,chr(13),-1)

	if alerta = "" then
		if ckb_periodo_devolucao = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "É obrigatório informar um período para a consulta."
			end if
		end if
	
	if alerta = "" then
		if ckb_periodo_devolucao <> "" then
			if Not isDate(c_dt_devolucao_inicio) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Informe uma data inicial válida para o período de devolução."
				end if
			if Not isDate(c_dt_devolucao_termino) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Informe uma data final válida para o período de devolução."
				end if
			if isDate(c_dt_devolucao_inicio) And isDate(c_dt_devolucao_termino) then
				if strToDate(c_dt_devolucao_inicio) > strToDate(c_dt_devolucao_termino) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "A data inicial do período de devolução é posterior à data final."
					end if
				end if
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
					'   CARREGA CÓDIGO INTERNO DO PRODUTO
						c_fabricante = Trim("" & rs("fabricante"))
						c_produto = Trim("" & rs("produto"))
						end if
					end if
				end if
			end if
		end if


'	Pedido cadastrado?
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
		
'	Vendedor
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






' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function monta_link_pedido(byval id_pedido)
dim strLink
	monta_link_pedido = ""
	id_pedido = Trim("" & id_pedido)
	if id_pedido = "" then exit function
	strLink = "<a href='javascript:fRELConcluir(" & _
				chr(34) & id_pedido & chr(34) & _
				")' title='clique para consultar o pedido " & id_pedido & "'>" & _
				id_pedido & "</a>"
	monta_link_pedido=strLink
end function


' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim cab, cab_table
dim n_reg_total
dim x, s, s_aux, s_sql, msg_erro
dim s_where, s_where_loja
dim intQtdeTotal

'	CRITÉRIOS COMUNS
	s_where = ""

'	FILTROS
'	~~~~~~~
'	PERÍODO DE DEVOLUÇÃO
	if ckb_periodo_devolucao <> "" then
		s = ""
		if c_dt_devolucao_inicio <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(StrToDate(c_dt_devolucao_inicio)) & ")"
			end if
			
		if c_dt_devolucao_termino <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(StrToDate(c_dt_devolucao_termino)+1) & ")"
			end if
			
		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
		end if
		
'	FABRICANTE
	if c_fabricante <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = '" & c_fabricante & "')"
		end if

'	PRODUTO
	if c_produto <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO_ITEM_DEVOLVIDO.produto = '" & c_produto & "')"
		end if

'	PEDIDO
	if c_pedido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO_ITEM_DEVOLVIDO.pedido = '" & c_pedido & "')"
		end if

'	VENDEDOR
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.vendedor = '" & c_vendedor & "')"
		end if

'	INDICADOR
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if

'	CAPTADOR
	if c_captador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.captador = '" & c_captador & "')"
		end if

'	LOJAS
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
	
'	MONTA A CONSULTA
	if s_where <> "" then s_where = " WHERE" & s_where
	
	s_sql = "SELECT " & _
				"*" & _
			" FROM (" & _
				"SELECT" & _
					" t_PEDIDO.data AS data_pedido," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data_devolucao," & _
					" t_PEDIDO.loja," & _
					" t_PEDIDO.numero_loja," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.id AS id_item_devolvido," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.fabricante," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.produto," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.descricao," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.descricao_html," & _
					" t_PEDIDO.pedido," & _
					" t_PEDIDO.obs_2," & _
					" t_PEDIDO__BASE.vendedor," & _
					" t_PEDIDO__BASE.indicador,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_cliente,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome_iniciais_em_maiusculas AS nome_cliente,"
		end if

	s_sql = s_sql & _
					" t_PEDIDO_ITEM_DEVOLVIDO.motivo," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.qtde," & _
					" (SELECT Count(*) FROM t_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS tAuxPIDBN INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO tAuxPID ON (tAuxPIDBN.id_item_devolvido=tAuxPID.id) WHERE (tAuxPID.pedido=t_PEDIDO.pedido) AND (anulado_status = 0)) AS qtde_msgs," & _
					" (" & _
						"SELECT" & _
							" TOP 1 t_ESTOQUE_MOVIMENTO.anulado_data" & _
						" FROM t_ESTOQUE INNER JOIN t_ESTOQUE_MOVIMENTO ON (t_ESTOQUE.id_estoque=t_ESTOQUE_MOVIMENTO.id_estoque)" & _
						" WHERE" & _
							" (t_ESTOQUE.devolucao_id_item_devolvido=t_PEDIDO_ITEM_DEVOLVIDO.id)" & _
							" AND (t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM_DEVOLVIDO.fabricante)" & _
							" AND (t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM_DEVOLVIDO.produto)" & _
							" AND (t_ESTOQUE_MOVIMENTO.estoque='" & ID_ESTOQUE_DEVOLUCAO & "')" & _
							" AND (operacao='" & OP_ESTOQUE_DEVOLUCAO & "')" & _
						" ORDER BY" & _
							" id_movimento" & _
					") AS data_baixa" & _
				" FROM t_PEDIDO" & _
					" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
						" ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
					" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido)" & _
					" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR" & _
						" ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
					" INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
				s_where & _
				") t" & _
			" ORDER BY" & _
				" data_pedido," & _
				" data_devolucao," & _
				" data_baixa," & _
				" fabricante," & _
				" produto," & _
				" descricao," & _
				" pedido"
	
  ' CABEÇALHO
	cab_table = "<TABLE CellSpacing=0 CellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='tdMargemEsq'>&nbsp;</TD>" & chr(13) & _
		  "		<TD class='MDTE tdData' style='vertical-align:bottom'><P class='Rc'>DT Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdData' style='vertical-align:bottom'><P class='Rc'>DT Devol</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdData' style='vertical-align:bottom'><P class='Rc'>DT Baixa</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdQtd' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Qtde</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdProduto' style='vertical-align:bottom'><P class='R'>Produto</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdPedido' style='vertical-align:bottom'><P class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdCliente' style='vertical-align:bottom'><P class='R'>Cliente</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdVendedor' style='vertical-align:bottom'><P class='R'>Vendedor</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdIndicador' style='vertical-align:bottom'><P class='R'>Parceiro</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdMotivo' style='vertical-align:bottom'><P class='R'>Motivo</P></TD>" & chr(13) & _
		  "		<TD class='tdBotao'>&nbsp;</TD>" & chr(13) & _
		  "	</TR>" & chr(13)

	n_reg_total = 0
	x = cab_table & cab

	If Not cria_recordset_otimista(r, msg_erro) then 
		Response.Write msg_erro
		exit sub
		end if

	r.open s_sql, cn
	do while Not r.Eof
	
	 ' CONTAGEM
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	'> T_PEDIDO_ITEM_DEVOLVIDO.ID (CAMPO HIDDEN)
		x = x & "		<input type=hidden name='c_id_item_devolvido_" & Cstr(n_reg_total) & "' id='c_id_item_devolvido_" & Cstr(n_reg_total) & "' value='" & Trim("" & r("id_item_devolvido")) & "'>" & chr(13)
		
	'> Nº PEDIDO (HIDDEN)
		x = x & "		<input type=hidden name='c_pedido_" & Cstr(n_reg_total) & "' id='c_pedido_" & Cstr(n_reg_total) & "' value='" & Trim("" & r("pedido")) & "'>" & chr(13)

	'> ESPAÇAMENTO À ESQUERDA P/ TENTAR MELHORAR A CENTRALIZAÇÃO DEVIDO À COLUNA DO BOTÃO À DIREITA (QUE NÃO APARECE NA IMPRESSÃO)
		x = x & "		<TD class='tdMargemEsq'>&nbsp;</TD>" & chr(13)
		
	'> DATA DO PEDIDO
		s = formata_data(r("data_pedido"))
		x = x & "		<TD class='MDTE tdData'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> DATA DA DEVOLUÇÃO
		s = formata_data(r("data_devolucao"))
		x = x & "		<TD class='MTD tdData'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> DATA DA BAIXA
		s = formata_data(r("data_baixa"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdData'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> QUANTIDADE
		s = formata_inteiro(r("qtde"))
		x = x & "		<TD class='MTD tdQtd'><P class='Cd'>" & s & "</P></TD>" & chr(13)

		intQtdeTotal = intQtdeTotal + r("qtde")

	'> PRODUTO
		s = "(" & Trim("" & r("fabricante")) & ")<b>" & Trim("" & r("produto")) & "</b> - " & Trim("" & r("descricao_html"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdProduto'><P class='Cn'>" & s & "</P></TD>" & chr(13)
	
	'> PEDIDO
		s = monta_link_pedido(Trim("" & r("pedido")))
		x = x & "		<TD class='MTD tdPedido'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> CLIENTE
		s = Trim("" & r("nome_cliente"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdCliente'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> VENDEDOR
		s = iniciais_em_maiusculas(Trim("" & r("vendedor")))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdVendedor'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> INDICADOR/PARCEIRO
		s = iniciais_em_maiusculas(Trim("" & r("indicador")))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdIndicador'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> MOTIVO
		s = Trim("" & r("motivo"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdMotivo'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> BOTÃO P/ EXIBIR DEMAIS CAMPOS
		x = x & "		<TD valign='bottom' class='notPrint tdBotao'>" & _
							"&nbsp;" & _
							"<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(n_reg_total) & chr(34) & ")' title='exibe ou oculta os campos adicionais'>" & _
							"<img src='../botao/view_bottom.png' border='0'>"
		if CLng(r("qtde_msgs")) > 0 then
			x = x & _
							"<span class='lblQtdeMsgs'> (" & Cstr(r("qtde_msgs")) & ")</span>"
			end if
			
		x = x & _
							"</a>" & _
						"</TD>" & chr(13)

		x = x & "	</TR>" & chr(13)

	'> MENSAGENS
		s_sql = "SELECT " & _
					"*" & _
				" FROM t_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS tPIDBN" & _
					" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO tPID ON (tPIDBN.id_item_devolvido=tPID.id)" & _
				" WHERE" & _
					" (tPID.pedido = '" & r("pedido") & "')" & _
					" AND (anulado_status = 0)" & _
				" ORDER BY" & _
					" tPIDBN.dt_hr_cadastro," & _
					" tPIDBN.id"
		if rs.State <> 0 then rs.Close
		rs.open s_sql, cn
		x = x & "	<TR style='display:none;' id='TR_MSGS_" & Cstr(n_reg_total) & "'>" & chr(13) & _
				"		<TD class='tdMargemEsq'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='9' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>MENSAGENS</td>" & chr(13) & _
				"				</TR>" & chr(13)
		if rs.Eof then
			x = x & _
				"				<TR>" & chr(13) & _
				"					<td>&nbsp;</td>" & chr(13) & _
				"				</TR>" & chr(13)
			end if

		do while Not rs.Eof
			x = x & _
				"				<TR>" & chr(13) & _
				"					<TD>" & chr(13) & _
				"						<table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
				"							<TR>" & chr(13) & _
				"								<TD class='Cn MD MC tdWithPadding tdDataHoraMsg' align='center'>" & chr(13) & _
													formata_data_hora_sem_seg(rs("dt_hr_cadastro")) & _
				"								</TD>" & chr(13) & _
				"								<TD class='Cn MD MC tdWithPadding tdUsuarioMsg' align='center'>" & chr(13) & _
													rs("usuario")
			if Trim("" & rs("loja")) <> "" then x = x & " (Loja&nbsp;" & Trim("" & rs("loja")) & ")"
			x = x & _
				"								</TD>" & chr(13) & _
				"								<TD class='Cn MC tdWithPadding tdTextoMensagem' align='left' valign='top'>" & chr(13) & _
													substitui_caracteres(Trim("" & rs("mensagem")), chr(13), "<br>") & _
												"</TD>" & chr(13) & _
				"							</TR>" & chr(13) & _
				"						</table>" & chr(13) & _
				"					</TD>" & chr(13) & _
				"				</TR>" & chr(13)
			rs.MoveNext
			loop
		
		x = x & _
				"			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)

	'> NOVA MENSAGEM
		x = x & "	<TR style='display:none;' id='TR_NEW_MSG_" & Cstr(n_reg_total) & "'>" & chr(13) & _
				"		<TD class='tdMargemEsq'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='9' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>NOVA MENSAGEM</td>" & chr(13) & _
				"					<td>&nbsp;</td>" & chr(13) & _
				"					<td align='right' valign='bottom'>" & chr(13) & _
										"<span class='PLLd' style='font-weight:normal;'>Tamanho restante:</span><input name='c_tamanho_restante_nova_msg_" & Cstr(n_reg_total) & "' id='c_tamanho_restante_nova_msg_" & Cstr(n_reg_total) & "' tabindex=-1 readonly class='TA' style='width:35px;text-align:right;font-size:8pt;font-weight:normal;' value='" & Cstr(MAX_TAM_MENSAGEM_BLOCO_NOTAS_EM_ITEM_DEVOLVIDO) & "' />" & chr(13) & _
				"					</td>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<TD colspan='3'>" & chr(13) & _
				"						<table width='100%' cellSpacing='0'>" & chr(13) & _
				"							<TR>" & chr(13) & _
				"								<td>" & chr(13) & _
													"<textarea name='c_nova_msg_" & Cstr(n_reg_total) & "' id='c_nova_msg_" & Cstr(n_reg_total) & "' class='PLLe' rows='3' style='width:100%;margin-left:2pt;' onkeypress='limita_tamanho(this,MAX_TAM_MENSAGEM_BLOCO_NOTAS_EM_ITEM_DEVOLVIDO);' onblur='this.value=trim(this.value);calcula_tamanho_restante_nova_msg(" & chr(34) & Cstr(n_reg_total) & chr(34) & ");' onkeyup='calcula_tamanho_restante_nova_msg(" & chr(34) & Cstr(n_reg_total) & chr(34) & ");'></textarea>" & _
				"								</td>" & chr(13) & _
				"							</TR>" & chr(13) & _
				"						</table>" & chr(13) & _
				"					</TD>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop


'	TOTAL GERAL
	if n_reg_total > 0 then
		x = x & "<TR><TD class='tdMargemEsq'>&nbsp;</TD><TD COLSPAN='10' class='MC' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
				"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='tdMargemEsq'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MTBE' colspan='3' NOWRAP><p class='Cd'>TOTAL GERAL:</p></td>" & chr(13) & _
				"		<TD class='MTB' NOWRAP><p class='Cd'>" & formata_inteiro(intQtdeTotal) & "</p></td>" & chr(13) & _
				"		<TD class='MTBD' colspan='6'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if
	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='tdMargemEsq'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MT' colspan='10'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)

	x = x & "<input type=hidden name='c_qtde_registros' id='c_qtde_registros' value='" & Cstr(n_reg_total) & "'>" & chr(13)
	
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

function calcula_tamanho_restante_nova_msg(indice_row) {
	var ctr, cnm, s;
	ctr = document.getElementById("c_tamanho_restante_nova_msg_" + indice_row);
	cnm = document.getElementById("c_nova_msg_" + indice_row);
	s = "" + cnm.value;
	ctr.value = MAX_TAM_MENSAGEM_BLOCO_NOTAS_EM_ITEM_DEVOLVIDO - s.length;
}

function fExibeOcultaCampos(indice_row) {
	var row_MSGS, row_NEW_MSG;

	row_MSGS = document.getElementById("TR_MSGS_" + indice_row);
	row_NEW_MSG = document.getElementById("TR_NEW_MSG_" + indice_row);

	if (row_MSGS.style.display.toString() == "none") {
		row_MSGS.style.display = "";
		row_NEW_MSG.style.display = "";
	}
	else {
		row_MSGS.style.display = "none";
		row_NEW_MSG.style.display = "none";
	}
}

function fRELConcluir( id_pedido ){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=id_pedido;
	fREL.action = "pedido.asp"
	fREL.submit();
}

function fRELGravaDados(f) {
var c, i, n, blnAchou;
	c = document.getElementById("c_qtde_registros");
	n = parseInt(c.value);
	blnAchou = false;
	for (i = 1; i <= n; i++) {
		c = document.getElementById("c_nova_msg_" + i.toString());
		if (c.value != "") {
			blnAchou = true;
			break;
		}
	}

	if (!blnAchou) {
		alert("Não há nenhuma mensagem para gravar!!");
		return;
	}

	f.action = "RelDevolucao2GravaDados.asp"
	dCONFIRMA.style.visibility = "hidden";
	window.status = "Aguarde ...";
	f.submit();
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
<link href="<%=URL_FILE__ESCREEN_CSS%>" Rel="stylesheet" Type="text/css" media="screen">

<style type="text/css">
.tdData{
	vertical-align: top;
	width: 65px;
	}
.tdPedido{
	vertical-align: top;
	font-weight: bold;
	width: 65px;
	}
.tdQtd{
	vertical-align: top;
	width: 35px;
	}
.tdProduto{
	vertical-align: top;
	width: 160px;
	}
.tdCliente{
	vertical-align: top;
	width: 130px;
	}
.tdIndicador{
	vertical-align: top;
	width: 65px;
	}
.tdVendedor{
	vertical-align: top;
	width: 65px;
	}
.tdMotivo{
	vertical-align: top;
	width: 160px;
	}
.tdWithPadding
{
	padding:1px;
}
.tdDataHoraMsg{
	vertical-align: top;
	width: 63px;
	}
.tdUsuarioMsg{
	vertical-align: top;
	width: 80px;
	}
.tdTextoMensagem{
	vertical-align: top;
	width: 665px;
	}
.lblQtdeMsgs
{
	font-family: Arial, Helvetica, sans-serif;
	color: #000000;
	font-size: 8pt;
	font-style: normal;
	position:relative;
	bottom: 3px;
}
.tdMargemEsq
{
	width:18px;
	background:white;
}
.tdBotao
{
	width:46px;
	background:white;
}
</style>


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
<input type="hidden" name="ckb_periodo_devolucao" id="ckb_periodo_devolucao" value="<%=ckb_periodo_devolucao%>">
<input type="hidden" name="c_dt_devolucao_inicio" id="c_dt_devolucao_inicio" value="<%=c_dt_devolucao_inicio%>">
<input type="hidden" name="c_dt_devolucao_termino" id="c_dt_devolucao_termino" value="<%=c_dt_devolucao_termino%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">
<input type="hidden" name="c_pedido" id="c_pedido" value="<%=c_pedido%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">
<input type="hidden" name="c_captador" id="c_captador" value="<%=c_captador%>">
<input type="HIDDEN" name="c_lista_loja" id="c_lista_loja" value="<%=c_lista_loja%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="926" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Devolução de Produtos II</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='926' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

	if ckb_periodo_devolucao <> "" then
		s = ""
		s_aux = c_dt_devolucao_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " e "
		s_aux = c_dt_devolucao_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Devolvido entre:&nbsp;</p></td><td valign='top' width='99%'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if

	s = c_fabricante
	if s <> "" then
		s_aux = x_fabricante(s)
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		end if
	if s = "" then s = "todos"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Fabricante:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)
	
	s = c_produto
	if s <> "" then
		s_aux = produto_formata_descricao_em_html(produto_descricao_html(c_fabricante, s))
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		end if
	if s = "" then s = "todos"
	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='baseline' NOWRAP>" & _
					"<p class='N'>Produto:&nbsp;</p></td><td valign='baseline'>" & _
					"<p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)

	s = c_pedido
	if s = "" then s = "todos"
	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP>" & _
					"<p class='N'>Pedido:&nbsp;</p></td><td valign='top'>" & _
					"<p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)

	s = c_vendedor
	if s = "" then 
		s = "todos"
	else
		if (s_nome_vendedor <> "") And (s_nome_vendedor <> c_vendedor) then s = s & " (" & s_nome_vendedor & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Vendedor:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s = c_indicador
	if s = "" then 
		s = "todos"
	else
		s = s & " (" & x_orcamentista_e_indicador(c_indicador) & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Indicador:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s = c_captador
	if s = "" then 
		s = "todos"
	else
		s = s & " (" & x_usuario(c_captador) & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Captador:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"
	
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

	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP>" & _
					"<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
					"<p class='N'>" & formata_data_hora(Now) & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
	
	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="926" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="926" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" 
		<% if origem="A" then %>
			href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>"
		<% else %>
			href="javascript:history.back()"
		<% end if %>
	title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava as mensagens">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
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
