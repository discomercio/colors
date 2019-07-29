<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L L O G E S T O Q U E E X E C . A S P
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
	if Not operacao_permitida(OP_CEN_REL_LOG_ESTOQUE, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, flag_ok
	dim c_dt_inicio, c_dt_termino, c_fabricante, c_produto, c_id_nfe_emitente

	alerta = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_id_nfe_emitente = Trim(Request.Form("c_id_nfe_emitente"))

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
dim s, s_aux, s_where, s_where_movto, s_where_entrada, s_where_movto_ativo, s_where_movto_anulado
dim s_sql, cab_table, cab, n_reg, n_reg_total, x, msg_erro

'	CRITÉRIOS COMUNS
	s_where = ""

'	CRITÉRIOS PARA REGISTROS DE MOVIMENTO (COMUM)
	s_where_movto = ""
	if c_fabricante <> "" then
		if s_where_movto <> "" then s_where_movto = s_where_movto & " AND"
		s_where_movto = s_where_movto & " (t_ESTOQUE_MOVIMENTO.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where_movto <> "" then s_where_movto = s_where_movto & " AND"
		s_where_movto = s_where_movto & " (t_ESTOQUE_MOVIMENTO.produto = '" & c_produto & "')"
		end if

	if c_id_nfe_emitente <> "" then
		if s_where_movto <> "" then s_where_movto = s_where_movto & " AND"
		s_where_movto = s_where_movto & " ((t_ESTOQUE.id_nfe_emitente = " & c_id_nfe_emitente & ") OR (t_PEDIDO.id_nfe_emitente = " & c_id_nfe_emitente & "))"
		end if

	if s_where_movto <> "" then s_where_movto = s_where_movto & " AND"
	s_where_movto = s_where_movto & _
					" ((operacao<>'" & OP_ESTOQUE_CONVERSAO_KIT & "')OR(operacao='')OR(operacao IS NULL))"

	if s_where_movto <> "" then s_where_movto = s_where_movto & " AND"
	s_where_movto = s_where_movto & _
					" ((operacao<>'" & OP_ESTOQUE_DEVOLUCAO & "')OR(operacao='')OR(operacao IS NULL))"

'	CRITÉRIOS PARA REGISTROS DE MOVIMENTO ATIVOS
	s_where_movto_ativo = ""
	if IsDate(c_dt_inicio) then
		if s_where_movto_ativo <> "" then s_where_movto_ativo = s_where_movto_ativo & " AND"
		s_where_movto_ativo = s_where_movto_ativo & " (t_ESTOQUE_MOVIMENTO.data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where_movto_ativo <> "" then s_where_movto_ativo = s_where_movto_ativo & " AND"
		s_where_movto_ativo = s_where_movto_ativo & " (t_ESTOQUE_MOVIMENTO.data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if

'	CRITÉRIOS PARA REGISTROS DE MOVIMENTO ANULADOS
	s_where_movto_anulado = ""
	if IsDate(c_dt_inicio) then
		if s_where_movto_anulado <> "" then s_where_movto_anulado = s_where_movto_anulado & " AND"
		s_where_movto_anulado = s_where_movto_anulado & " (t_ESTOQUE_MOVIMENTO.anulado_data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where_movto_anulado <> "" then s_where_movto_anulado = s_where_movto_anulado & " AND"
		s_where_movto_anulado = s_where_movto_anulado & " (t_ESTOQUE_MOVIMENTO.anulado_data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
'	CRITÉRIOS PARA REGISTROS DE ENTRADA
	s_where_entrada = ""
	if IsDate(c_dt_inicio) then
		if s_where_entrada <> "" then s_where_entrada = s_where_entrada & " AND"
		s_where_entrada = s_where_entrada & " (t_ESTOQUE.data_entrada >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where_entrada <> "" then s_where_entrada = s_where_entrada & " AND"
		s_where_entrada = s_where_entrada & " (t_ESTOQUE.data_entrada < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
	if c_fabricante <> "" then
		if s_where_entrada <> "" then s_where_entrada = s_where_entrada & " AND"
		s_where_entrada = s_where_entrada & " (t_ESTOQUE_ITEM.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where_entrada <> "" then s_where_entrada = s_where_entrada & " AND"
		s_where_entrada = s_where_entrada & " (t_ESTOQUE_ITEM.produto = '" & c_produto & "')"
		end if

	if c_id_nfe_emitente <> "" then
		if s_where_entrada <> "" then s_where_entrada = s_where_entrada & " AND"
		s_where_entrada = s_where_entrada & " (t_ESTOQUE.id_nfe_emitente = " & c_id_nfe_emitente & ")"
		end if

'	A) IMPORTANTE: COMO ESTE RELATÓRIO LISTA O HISTÓRICO DE MOVIMENTAÇÕES NO ESTOQUE,
'		EXCEPCIONALMENTE NÃO SE DEVE FAZER A RESTRIÇÃO "anulado_status=0".
'		LEMBRANDO QUE "anulado_status=0" É USADO NORMALMENTE P/ SELECIONAR APENAS 
'		OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'		FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s = s_where
	if (s <> "") And (s_where_movto <> "") then s = s & " AND"
	s = s & s_where_movto
	if (s <> "") And (s_where_movto_ativo <> "") then s = s & " AND"
	s = s & s_where_movto_ativo
'	HÁ RESTRIÇÕES?
	if s <> "" then s = " WHERE" & s
	s_sql = "SELECT t_ESTOQUE_MOVIMENTO.data, t_ESTOQUE_MOVIMENTO.hora," & _
			" t_ESTOQUE.id_nfe_emitente AS id_nfe_emitente_estoque," & _
			" t_PEDIDO.id_nfe_emitente AS id_nfe_emitente_pedido," & _
			" t_ESTOQUE_MOVIMENTO.fabricante AS fabricante, t_ESTOQUE_MOVIMENTO.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" t_ESTOQUE.documento, t_ESTOQUE_MOVIMENTO.qtde," & _
			" t_ESTOQUE_MOVIMENTO.operacao, t_ESTOQUE_MOVIMENTO.estoque," & _
			" t_ESTOQUE_ITEM.preco_fabricante, t_PEDIDO_ITEM.preco_venda," & _
			" t_ESTOQUE_MOVIMENTO.pedido AS pedido, t_ESTOQUE_MOVIMENTO.id_movimento AS nsu_seq" & _
			" FROM t_ESTOQUE_MOVIMENTO LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE.id_estoque)" & _
			" LEFT JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
			" LEFT JOIN t_PEDIDO ON (t_PEDIDO.pedido = t_ESTOQUE_MOVIMENTO.pedido)" & _
			" LEFT JOIN t_PEDIDO_ITEM ON ((t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto))" & _
			s

	s = s_where
	if (s <> "") And (s_where_movto <> "") then s = s & " AND"
	s = s & s_where_movto
	if (s <> "") And (s_where_movto_anulado <> "") then s = s & " AND"
	s = s & s_where_movto_anulado
'	RESTRIÇÕES
	if s <> "" then s = s & " AND"
	s = s & " (anulado_status<>0)"
'	HÁ RESTRIÇÕES?
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT t_ESTOQUE_MOVIMENTO.anulado_data AS data, t_ESTOQUE_MOVIMENTO.anulado_hora AS hora," & _
			" t_ESTOQUE.id_nfe_emitente AS id_nfe_emitente_estoque," & _
			" t_PEDIDO.id_nfe_emitente AS id_nfe_emitente_pedido," & _
			" t_ESTOQUE_MOVIMENTO.fabricante AS fabricante, t_ESTOQUE_MOVIMENTO.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" t_ESTOQUE.documento, t_ESTOQUE_MOVIMENTO.qtde," & _
			" '~' + t_ESTOQUE_MOVIMENTO.operacao AS operacao, t_ESTOQUE_MOVIMENTO.estoque," & _
			" t_ESTOQUE_ITEM.preco_fabricante, t_PEDIDO_ITEM.preco_venda," & _
			" t_ESTOQUE_MOVIMENTO.pedido AS pedido, t_ESTOQUE_MOVIMENTO.id_movimento AS nsu_seq" & _
			" FROM t_ESTOQUE_MOVIMENTO LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE.id_estoque)" & _
			" LEFT JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
			" LEFT JOIN t_PEDIDO ON (t_PEDIDO.pedido = t_ESTOQUE_MOVIMENTO.pedido)" & _
			" LEFT JOIN t_PEDIDO_ITEM ON ((t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto))" & _
			s
	
	s = s_where
	if (s <> "") And (s_where_entrada <> "") then s = s & " AND"
	s = s & s_where_entrada
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT t_ESTOQUE.data_entrada AS data, t_ESTOQUE.hora_entrada AS hora," & _
			" t_ESTOQUE.id_nfe_emitente AS id_nfe_emitente_estoque," & _
			" NULL AS id_nfe_emitente_pedido," & _
			" t_ESTOQUE_ITEM.fabricante AS fabricante, t_ESTOQUE_ITEM.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" t_ESTOQUE.documento, t_ESTOQUE_ITEM.qtde," & _
			" '" & OP_ESTOQUE_ENTRADA & "' AS operacao, '" & ID_ESTOQUE_VENDA & "' AS estoque," & _
			" t_ESTOQUE_ITEM.preco_fabricante, NULL AS preco_venda," & _
			" '' AS pedido, t_ESTOQUE.id_estoque AS nsu_seq" & _
			" FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante)AND(t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
			s
	
	s_sql = "SELECT" & _
				" tNFeEmit1.apelido AS id_nfe_emitente_estoque_descricao," & _
				" tNFeEmit2.apelido AS id_nfe_emitente_pedido_descricao," & _
				" *" & _
			" FROM" & _
				" (" & s_sql & ") t" & _
				" LEFT JOIN t_NFe_EMITENTE tNFeEmit1 ON (t.id_nfe_emitente_estoque = tNFeEmit1.id)" & _
				" LEFT JOIN t_NFe_EMITENTE tNFeEmit2 ON (t.id_nfe_emitente_pedido = tNFeEmit2.id)"

	s_sql = s_sql & " ORDER BY data, hora, nsu_seq"

  ' CABEÇALHO
	cab_table = "<TABLE class='Q' style='border-bottom:0px;' CellSpacing=0 CellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDB tdData' style='vertical-align:bottom'><P class='R'>Data</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdProd' style='vertical-align:bottom'><P class='R'>Produto</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdCD' style='vertical-align:bottom'><P class='R'>CD</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdPed' style='vertical-align:bottom'><P class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdDoc' style='vertical-align:bottom'><P class='R'>Documento</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdOp' style='vertical-align:bottom'><P class='R'>Operação</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdEst' style='vertical-align:bottom'><P class='R'>Tipo de Estoque</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdQtd' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Qtde</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdVle' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Valor Unit Entrada</P></TD>" & chr(13) & _
		  "		<TD class='MB tdVls' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Valor Unit Saída</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)

	n_reg = 0
	n_reg_total = 0

	x = cab_table & cab

	If Not cria_recordset_otimista(r, msg_erro) then 
		Response.Write msg_erro
		exit sub
		end if

	r.open s_sql, cn
	do while Not r.Eof

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> DATA
		s = formata_data(r("data"))
		s_aux = formata_hhnnss_para_hh_nn_ss(r("hora"))
		if (s<>"") And (s_aux<>"") then s = s & "<br>"
		s = s & s_aux
		x = x & "		<TD class='MDB tdData'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	 '> PRODUTO
		s = Trim("" & r("produto"))
		s_aux = Trim("" & r("descricao_html"))
		if s_aux <> "" then s_aux = produto_formata_descricao_em_html(s_aux)
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
	'	FABRICANTE
		s_aux = Trim("" & r("fabricante"))
		if s_aux <> "" then s_aux = "(" & s_aux & ") "
		s = s_aux & s
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MDB tdProd'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	 '> CD
		s = ""
		if converte_numero(Trim("" & r("id_nfe_emitente_estoque"))) <> 0 then
			s = Trim("" & r("id_nfe_emitente_estoque_descricao"))
		elseif converte_numero(Trim("" & r("id_nfe_emitente_pedido"))) <> 0 then
			s = Trim("" & r("id_nfe_emitente_pedido_descricao"))
			end if

		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MDB tdCD'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	 '> PEDIDO
		s = Trim("" & r("pedido"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MDB tdPed'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	 '> DOCUMENTO
		s = Trim("" & r("documento"))
		if s <> "" then s = iniciais_em_maiusculas(s)
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MDB tdDoc'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	 '> OPERAÇÃO
		s = Trim("" & r("operacao"))
		if left(s, 1) = "~" then
			s_aux = "CANCELADO: "
			s = mid(s, 2)
		else
			s_aux = ""
			end if
		if s <> "" then s = x_operacao_estoque(s)
		if s <> "" then s = s_aux & s
		if s = "" then s = "&nbsp;" 
		x = x & "		<TD class='MDB tdOp'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	 '> ESTOQUE
		s = Trim("" & r("estoque"))
		if s <> "" then s = x_estoque(s)
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MDB tdEst'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	 '> QUANTIDADE
		x = x & "		<TD class='MDB tdQtd'><P class='Cnd'>" & formata_inteiro(r("qtde")) & "</P></TD>" & chr(13)

	 '> VALOR ENTRADA
		if IsNull(r("preco_fabricante")) then
			s = "&nbsp;"
		else
			s = formata_moeda(r("preco_fabricante"))
			end if
		x = x & "		<TD class='MDB tdVle'><P class='Cnd'>" & s & "</P></TD>" & chr(13)

	 '> VALOR SAÍDA
		if IsNull(r("preco_venda")) then
			s = "&nbsp;"
		else
			s = formata_moeda(r("preco_venda"))
			end if
		x = x & "		<TD class='MB tdVls'><P class='Cnd'>" & s & "</P></TD>" & chr(13)

		x = x & "	</TR>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop

	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = ""
		if c_fabricante <> "" then
			s = c_fabricante
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			if s <> "" then x = x & cab_table & _
									"	<TR>" & chr(13) & _
									"		<TD class='MB' COLSPAN='10' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
									"	</tr>" & chr(13) & cab
		else
			x = x & cab_table & cab
			end if

		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MB' colspan='10'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
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
.tdData{
	vertical-align: top;
	width: 60px;
	}
.tdProd{
	vertical-align: top;
	width: 128px;
	}
.tdCD{
	vertical-align: top;
	width: 50px;
	}
.tdPed{
	vertical-align: top;
	width: 65px;
	}
.tdDoc{
	vertical-align: top;
	width: 85px;
	}
.tdOp{
	vertical-align: top;
	width: 74px;
	}
.tdEst{
	vertical-align: top;
	width: 65px;
	}
.tdQtd{
	vertical-align: top;
	width: 36px;
	}
.tdVle{
	vertical-align: top;
	width: 63px;
	}
.tdVls{
	vertical-align: top;
	width: 63px;
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
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="708" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Log Estoque</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='708' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Período:&nbsp;</p></td><td valign='top' width='99%'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s = c_id_nfe_emitente
	if s <> "" then
		s_aux = obtem_apelido_empresa_NFe_emitente(s)
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		end if
	if s = "" then s = "todos"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Empresa (CD):&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

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
<table width="708" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="708" cellSpacing="0">
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
