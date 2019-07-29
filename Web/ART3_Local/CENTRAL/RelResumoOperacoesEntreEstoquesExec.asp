<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelResumoOperacoesEntreEstoquesExec.asp
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
	if Not operacao_permitida(OP_CEN_REL_RESUMO_OPERACOES_ENTRE_ESTOQUES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, s_filtro_loja, flag_ok, s_filtro_operacao
	dim c_dt_inicio, c_dt_termino, c_fabricante, c_produto, c_id_nfe_emitente
	dim c_lista_loja, s_lista_loja, v_loja, v, i

	dim ckb_OP_ESTOQUE_LOG_ENTRADA
	dim ckb_OP_ESTOQUE_LOG_VENDA
	dim ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA
	dim ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT
	dim ckb_OP_ESTOQUE_LOG_TRANSFERENCIA
	dim ckb_OP_ESTOQUE_LOG_ENTREGA
	dim ckb_OP_ESTOQUE_LOG_DEVOLUCAO
	dim ckb_OP_ESTOQUE_LOG_ESTORNO
	dim ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA
	dim ckb_OP_ESTOQUE_LOG_SPLIT
	dim ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA
	dim ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE
	dim ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM
	dim ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA
	dim ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA
	dim ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA
	dim ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS

	alerta = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_lista_loja = Trim(Request.Form("c_lista_loja"))
	s_lista_loja = substitui_caracteres(c_lista_loja,chr(10),"")
	v_loja = split(s_lista_loja,chr(13),-1)
	c_id_nfe_emitente = Trim(Request.Form("c_id_nfe_emitente"))

	ckb_OP_ESTOQUE_LOG_ENTRADA = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_ENTRADA"))
	ckb_OP_ESTOQUE_LOG_VENDA = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_VENDA"))
	ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA"))
	ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT"))
	ckb_OP_ESTOQUE_LOG_TRANSFERENCIA = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_TRANSFERENCIA"))
	ckb_OP_ESTOQUE_LOG_ENTREGA = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_ENTREGA"))
	ckb_OP_ESTOQUE_LOG_DEVOLUCAO = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_DEVOLUCAO"))
	ckb_OP_ESTOQUE_LOG_ESTORNO = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_ESTORNO"))
	ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA"))
	ckb_OP_ESTOQUE_LOG_SPLIT = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_SPLIT"))
	ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA"))
	ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE"))
	ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM"))
	ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA"))
	ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA"))
	ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA"))
	ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS = Trim(Request.Form("ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS"))

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

function monta_link_pedido(byval id_pedido)
dim strLink
	monta_link_pedido = ""
	id_pedido = Trim("" & id_pedido)
	if id_pedido = "" then exit function
	strLink = "<a href='javascript:fRELConcluir(" & _
				chr(34) & id_pedido & chr(34) & _
				")' title='clique para consultar o pedido'>" & _
				id_pedido & "</a>"
	monta_link_pedido=strLink
end function


' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const TEXTO_INFO_NAO_DISPONIVEL = "N.D."
dim r
dim cab, cab_table
dim n_reg, n_reg_total
dim x, s, s2, s_op, s_estoque, s_sql, msg_erro
dim s_where, s_where_loja, s_where_loja_origem, s_where_loja_destino, s_where_operacao
dim s_loja_origem, s_loja_destino
dim s_pedido_origem, s_pedido_destino
dim intQtdeSolicitada, intQtdeAtendida, intQtde
dim blnPulaRegistro
dim strQtdeInicioEstoqueVenda
dim strQtdeInicioEstoqueShowRoom
dim strQtdeInicioEstoqueDevolucao
dim strQtdeInicioEstoqueSemPresenca
dim strQtdeInicioEstoqueVendido
dim strQtdeInicioEstoqueDanificado
dim strQtdeInicioEstoqueRouboPerda
dim strQtdeFimEstoqueVenda
dim strQtdeFimEstoqueShowRoom
dim strQtdeFimEstoqueDevolucao
dim strQtdeFimEstoqueSemPresenca
dim strQtdeFimEstoqueVendido
dim strQtdeFimEstoqueDanificado
dim strQtdeFimEstoqueRouboPerda
dim blnHaInfoSaldoInicio, blnHaInfoSaldoFim
dim dtSaldoInicio, dtSaldoFim

'	CRITÉRIOS COMUNS
	s_where = ""

'	FILTROS
'	~~~~~~~
'	PERÍODO
	if IsDate(c_dt_inicio) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tELog.data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tELog.data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
'	FABRICANTE
	if c_fabricante <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tELog.fabricante = '" & c_fabricante & "')"
		end if

'	PRODUTO
	if c_produto <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tELog.produto = '" & c_produto & "')"
		end if

'	EMPRESA (CD)
	if converte_numero(c_id_nfe_emitente) <> 0 then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " ((tELog.id_nfe_emitente = " & c_id_nfe_emitente & ") OR (tPOrig.id_nfe_emitente = " & c_id_nfe_emitente & ") OR (tPDest.id_nfe_emitente = " & c_id_nfe_emitente & "))"
	end if



'	LOJAS
	s_where_loja = ""
	s_where_loja_origem = ""
	s_where_loja_destino = ""
	for i=Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_where_loja_origem <> "" then s_where_loja_origem = s_where_loja_origem & " OR"
				s_where_loja_origem = s_where_loja_origem & " (CONVERT(smallint, tELog.loja_estoque_origem) = " & v_loja(i) & ")"
				if s_where_loja_destino <> "" then s_where_loja_destino = s_where_loja_destino & " OR"
				s_where_loja_destino = s_where_loja_destino & " (CONVERT(smallint, tELog.loja_estoque_destino) = " & v_loja(i) & ")"
			else
				s = ""
				s2 = ""
				if v(Lbound(v))<>"" then 
					if s <> "" then s = s & " AND"
					s = s & " (CONVERT(smallint, tELog.loja_estoque_origem) >= " & v(Lbound(v)) & ")"
					if s2 <> "" then s2 = s2 & " AND"
					s2 = s2 & " (CONVERT(smallint, tELog.loja_estoque_destino) >= " & v(Lbound(v)) & ")"
					end if
				if v(Ubound(v))<>"" then
					if s <> "" then s = s & " AND"
					s = s & " (CONVERT(smallint, tELog.loja_estoque_origem) <= " & v(Ubound(v)) & ")"
					if s2 <> "" then s2 = s2 & " AND"
					s2 = s2 & " (CONVERT(smallint, tELog.loja_estoque_destino) <= " & v(Ubound(v)) & ")"
					end if
				if s <> "" then 
					if s_where_loja_origem <> "" then s_where_loja_origem = s_where_loja_origem & " OR"
					s_where_loja_origem = s_where_loja_origem & " (" & s & ")"
					end if
				if s2 <> "" then 
					if s_where_loja_destino <> "" then s_where_loja_destino = s_where_loja_destino & " OR"
					s_where_loja_destino = s_where_loja_destino & " (" & s2 & ")"
					end if
				end if
			end if
		next
	
	if (s_where_loja_origem <> "") And (s_where_loja_destino <> "") then
		'Ambos possuem conteúdo
		s_where_loja = " (" & s_where_loja_origem & ") OR (" & s_where_loja_destino & ")"
	else
		'Um dos dois ou ambos estão vazios
		s_where_loja = s_where_loja_origem & s_where_loja_destino
		end if
	
	if s_where_loja <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s_where_loja & ")"
		end if

	'OPERAÇÕES
	s_where_operacao = ""
	if ckb_OP_ESTOQUE_LOG_ENTRADA <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_ENTRADA & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_VENDA <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_VENDA & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & _
							" (" & _
								"(tELog.operacao = '" & OP_ESTOQUE_LOG_TRANSFERENCIA & "')" & _
								" AND " & _
								"(" & _
									"(tELog.cod_estoque_origem = '" & ID_ESTOQUE_ROUBO & "')" & _
									" OR " & _
									"(tELog.cod_estoque_destino = '" & ID_ESTOQUE_ROUBO & "')" & _
								")" & _
							")"
		end if
		
	if ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_TRANSFERENCIA <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_TRANSFERENCIA & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_ENTREGA <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_ENTREGA & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_DEVOLUCAO <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_DEVOLUCAO & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_ESTORNO <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_ESTORNO & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_SPLIT <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_SPLIT & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA & "')"
		end if

	if ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS <> "" then
		if s_where_operacao <> "" then s_where_operacao = s_where_operacao & " OR"
		s_where_operacao = s_where_operacao & " (tELog.operacao = '" & ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS & "')"
		end if

	if s_where_operacao <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s_where_operacao & ")"
	else
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tELog.operacao IS NULL)"
		end if
	
	'MONTA A CONSULTA P/ OBTER DADOS DO LOG DE MOVIMENTAÇÃO NO ESTOQUE
	if s_where <> "" then s_where = " WHERE" & s_where
	
	s_sql = "SELECT " & _
				"tELog.data, " & _
				"tELog.data_hora, " & _
				"tELog.usuario, " & _
				"tELog.id_nfe_emitente, " & _
				"tNFeEmit1.apelido AS id_nfe_emitente_descricao, " & _
				"tELog.fabricante, " & _
				"tELog.produto, " & _
				"tELog.qtde_solicitada, " & _
				"tELog.qtde_atendida, " & _
				"tELog.operacao, " & _
				"tELog.cod_estoque_origem, " & _
				"tELog.cod_estoque_destino, " & _
				"tELog.loja_estoque_origem, " & _
				"tELog.loja_estoque_destino, " & _
				"tELog.pedido_estoque_origem, " & _
				"tPOrig.id_nfe_emitente AS id_nfe_emitente_pedido_origem, " & _
				"tNFeEmit2.apelido AS id_nfe_emitente_pedido_origem_descricao, " & _
				"tELog.pedido_estoque_destino, " & _
				"tPDest.id_nfe_emitente AS id_nfe_emitente_pedido_destino, " & _
				"tNFeEmit3.apelido AS id_nfe_emitente_pedido_destino_descricao, " & _
				"tELog.documento, " & _
				"tELog.complemento " & _
			"FROM t_ESTOQUE_LOG tELog" & _
				" LEFT JOIN t_NFe_EMITENTE tNFeEmit1 ON (tELog.id_nfe_emitente = tNFeEmit1.id)" & _
				" LEFT JOIN t_PEDIDO tPOrig ON (tELog.pedido_estoque_origem = tPOrig.pedido)" & _
				" LEFT JOIN t_NFe_EMITENTE tNFeEmit2 ON (tPOrig.id_nfe_emitente = tNFeEmit2.id)" & _
				" LEFT JOIN t_PEDIDO tPDest ON (tELog.pedido_estoque_destino = tPDest.pedido)" & _
				" LEFT JOIN t_NFe_EMITENTE tNFeEmit3 ON (tPDest.id_nfe_emitente = tNFeEmit3.id)" & _
			s_where & _
			" ORDER BY tELog.data_hora"
	
  ' CABEÇALHO
	cab_table = "<TABLE class='Q' style='border-bottom:0px;' CellSpacing=0 CellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDB tdData' style='vertical-align:bottom'><P class='R'>Data</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdCD' style='vertical-align:bottom'><P class='R'>CD</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdProd' style='vertical-align:bottom'><P class='R'>Produto</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdQtd' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Qtde</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdOp' style='vertical-align:bottom'><P class='R'>Operação</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdEstOrigem' style='vertical-align:bottom'><P class='R'>Estoque Origem</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdEstDestino' style='vertical-align:bottom'><P class='R'>Estoque Destino</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdLoja' style='vertical-align:bottom'><P class='R'>Loja</P></TD>" & chr(13) & _
		  "		<TD class='MDB tdPedido' style='vertical-align:bottom'><P class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MB tdOperador' style='vertical-align:bottom'><P class='R'>Operador</P></TD>" & chr(13) & _
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

		s_op = Trim("" & r("operacao"))
		intQtdeSolicitada = r("qtde_solicitada")
		intQtdeAtendida = r("qtde_atendida")

		blnPulaRegistro = False
		if (s_op=OP_ESTOQUE_LOG_VENDA) And (intQtdeAtendida=0) then blnPulaRegistro=True

		if Not blnPulaRegistro then
		
		 ' CONTAGEM
			n_reg = n_reg + 1
			n_reg_total = n_reg_total + 1

			x = x & "	<TR NOWRAP>" & chr(13)

		'> DATA
			s = formata_data(r("data_hora"))
			s_aux = formata_hora(r("data_hora"))
			if (s<>"") And (s_aux<>"") then s = s & "<br>"
			s = s & s_aux
			x = x & "		<TD class='MDB tdData'><P class='Cn'>" & s & "</P></TD>" & chr(13)

		'> CD
			s = ""
			if converte_numero(Trim("" & r("id_nfe_emitente"))) <> 0 then
				s = Trim("" & r("id_nfe_emitente_descricao"))
			elseif converte_numero(Trim("" & r("id_nfe_emitente_pedido_origem"))) <> 0 then
				s = Trim("" & r("id_nfe_emitente_pedido_origem_descricao"))
			elseif converte_numero(Trim("" & r("id_nfe_emitente_pedido_destino"))) <> 0 then
				s = Trim("" & r("id_nfe_emitente_pedido_destino_descricao"))
				end if
			if s = "" then s = "&nbsp;"
			x = x & "		<TD class='MDB tdCD'><P class='Cn'>" & s & "</P></TD>" & chr(13)

		'> PRODUTO
			s_aux = Trim("" & r("fabricante"))
			if s_aux <> "" then s_aux = "(" & s_aux & ") "
			s = Trim("" & r("produto"))
			s = s_aux & s
			if s = "" then s = "&nbsp;"
			x = x & "		<TD class='MDB tdProd'><P class='Cn'>" & s & "</P></TD>" & chr(13)

		'> QUANTIDADE
			if s_op = OP_ESTOQUE_LOG_VENDA then
				intQtde = intQtdeAtendida
			else
			'	O VALOR -1 INDICA QUE FOI SOLICITADO P/ PROCESSAR TUDO QUE FOSSE POSSÍVEL
				if intQtdeSolicitada=-1 then
					intQtde=intQtdeAtendida
				else
					intQtde=intQtdeSolicitada
					end if
				end if
		
			x = x & "		<TD class='MDB tdQtd'><P class='Cd'>" & formata_inteiro(intQtde) & "</P></TD>" & chr(13)

		'> OPERAÇÃO
			s = x_operacao_log_estoque(Trim("" & r("operacao")))
			if s = "" then s = "&nbsp;"
			x = x & "		<TD class='MDB tdOp'><P class='Cn'>" & s & "</P></TD>" & chr(13)

		'> ESTOQUE ORIGEM
			s = x_estoque(Trim("" & r("cod_estoque_origem")))
			if s = "" then s = "&nbsp;"
			x = x & "		<TD class='MDB tdEstOrigem'><P class='Cn'>" & s & "</P></TD>" & chr(13)

		'> ESTOQUE DESTINO
			s = x_estoque(Trim("" & r("cod_estoque_destino")))
			if s = "" then s = "&nbsp;"
			x = x & "		<TD class='MDB tdEstDestino'><P class='Cn'>" & s & "</P></TD>" & chr(13)

		'> LOJA
			s_loja_origem = Trim("" & r("loja_estoque_origem"))
			s_loja_destino = Trim("" & r("loja_estoque_destino"))
			if (s_loja_origem <> "") And (s_loja_destino <> "") then
				if s_loja_origem = s_loja_destino then
					s = s_loja_origem
				else
					s = s_loja_origem & " => " & s_loja_destino
					end if
			else
			'	'Um dos dois ou ambos estão vazios
				s = s_loja_origem & s_loja_destino
				end if
			if s = "" then s = "&nbsp;"
			x = x & "		<TD class='MDB tdLoja'><P class='Cn'>" & s & "</P></TD>" & chr(13)

		'> PEDIDO
			s_pedido_origem = Trim("" & r("pedido_estoque_origem"))
			s_pedido_destino = Trim("" & r("pedido_estoque_destino"))
			if (s_pedido_origem <> "") And (s_pedido_destino <> "") then
				if s_pedido_origem = s_pedido_destino then
					s = monta_link_pedido(s_pedido_origem)
				else
					s = monta_link_pedido(s_pedido_origem) & " => " & monta_link_pedido(s_pedido_destino)
					end if
			else
			'	'Um dos dois ou ambos estão vazios
				s = monta_link_pedido(s_pedido_origem) & monta_link_pedido(s_pedido_destino)
				end if
			if s = "" then s = "&nbsp;"
			x = x & "		<TD class='MDB tdPedido'><P class='Cn'>" & s & "</P></TD>" & chr(13)

		'> OPERADOR
			s = Trim("" & r("usuario"))
			if s <> "" then s = iniciais_em_maiusculas(s)
			if s = "" then s = "&nbsp;"
			x = x & "		<TD class='MB tdOperador'><P class='Cn'>" & s & "</P></TD>" & chr(13)

			x = x & "	</TR>" & chr(13)

			if (n_reg mod 100) = 0 then
				Response.Write x
				x = ""
				end if
			
			end if  ' if (blnPulaRegistro)
			
		r.MoveNext
		loop

	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MB' colspan='9'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)

'	EXIBE O SALDO DO ESTOQUE (APENAS NO CASO DE TER SIDO CONSULTADO UM PRODUTO ESPECÍFICO)
	if (c_fabricante <> "") And (c_produto <> "") then
	
		strQtdeInicioEstoqueVenda="0"
		strQtdeInicioEstoqueShowRoom="0"
		strQtdeInicioEstoqueDevolucao="0"
		strQtdeInicioEstoqueSemPresenca="0"
		strQtdeInicioEstoqueVendido="0"
		strQtdeInicioEstoqueDanificado="0"
		strQtdeInicioEstoqueRouboPerda="0"
		strQtdeFimEstoqueVenda="0"
		strQtdeFimEstoqueShowRoom="0"
		strQtdeFimEstoqueDevolucao="0"
		strQtdeFimEstoqueSemPresenca="0"
		strQtdeFimEstoqueVendido="0"
		strQtdeFimEstoqueDanificado="0"
		strQtdeFimEstoqueRouboPerda="0"

		blnHaInfoSaldoInicio = False
		blnHaInfoSaldoFim = False
		
		if IsDate(c_dt_inicio) then
			dtSaldoInicio=StrToDate(c_dt_inicio)
			
		'	JOB QUE COLETA OS DADOS FOI EXECUTADO COM SUCESSO NESSE DIA?
			s_sql = "SELECT" & _
						" *" & _
					" FROM t_ESTOQUE_SALDO_DIARIO" & _                    
					" WHERE" & _
						" (data = " & bd_formata_data(dtSaldoInicio) & ")" & _
						" AND (fabricante = '----')" & _
						" AND (produto = '--------')"
			if r.State <> 0 then r.Close
			r.open s_sql, cn
			if Not r.Eof then 
				blnHaInfoSaldoInicio = True
			else
				strQtdeInicioEstoqueVenda=TEXTO_INFO_NAO_DISPONIVEL
				strQtdeInicioEstoqueShowRoom=TEXTO_INFO_NAO_DISPONIVEL
				strQtdeInicioEstoqueDevolucao=TEXTO_INFO_NAO_DISPONIVEL
				strQtdeInicioEstoqueSemPresenca=TEXTO_INFO_NAO_DISPONIVEL
				strQtdeInicioEstoqueVendido=TEXTO_INFO_NAO_DISPONIVEL
				strQtdeInicioEstoqueDanificado=TEXTO_INFO_NAO_DISPONIVEL
				strQtdeInicioEstoqueRouboPerda=TEXTO_INFO_NAO_DISPONIVEL
				end if
			
			if blnHaInfoSaldoInicio then
				s_sql = "SELECT" & _
							" fabricante, produto, estoque, Coalesce(Sum(qtde), 0) As qtde" & _
						" FROM t_ESTOQUE_SALDO_DIARIO" & _
						" WHERE" & _
							" (data = " & bd_formata_data(dtSaldoInicio) & ")" & _
							" AND (fabricante = '" & c_fabricante & "')" & _
							" AND (produto = '" & c_produto & "')"
				
				if converte_numero(c_id_nfe_emitente) <> 0 then
					s_sql = s_sql & " AND (id_nfe_emitente = " & c_id_nfe_emitente & ")"
					end if

				s_sql = s_sql & _
						" GROUP BY fabricante, produto, estoque"

				if r.State <> 0 then r.Close
				r.open s_sql, cn
				do while Not r.Eof
					s_estoque=Trim("" & r("estoque"))
					
					if s_estoque = ID_ESTOQUE_VENDA then
						strQtdeInicioEstoqueVenda=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_VENDIDO then
						strQtdeInicioEstoqueVendido=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_SEM_PRESENCA then
						strQtdeInicioEstoqueSemPresenca=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_SHOW_ROOM then
						strQtdeInicioEstoqueShowRoom=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_DANIFICADOS then
						strQtdeInicioEstoqueDanificado=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_DEVOLUCAO then
						strQtdeInicioEstoqueDevolucao=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_ROUBO then
						strQtdeInicioEstoqueRouboPerda=formata_inteiro(r("qtde"))
						end if
					
					r.MoveNext
					loop
				end if
			end if  ' if (tem data início)

		if IsDate(c_dt_termino) then
			dtSaldoFim=StrToDate(c_dt_termino)
			if dtSaldoFim > Date then dtSaldoFim=Date
			
			if (dtSaldoFim = Date) then
			'	OBTÉM A POSIÇÃO ATUAL DOS ESTOQUES
				s_sql = "SELECT" & _
							" t_ESTOQUE_ITEM.fabricante, produto, 'VDA' As estoque, Sum(qtde-qtde_utilizada) As qtde" & _
						" FROM t_ESTOQUE_ITEM" & _
							" INNER JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque = t_ESTOQUE.id_estoque)" & _
						" WHERE" & _
							" ((qtde-qtde_utilizada) > 0)" & _
							" AND (t_ESTOQUE_ITEM.fabricante = '" & c_fabricante & "')" & _
							" AND (produto = '" & c_produto & "')" 
                        
                if converte_numero(c_id_nfe_emitente) <> 0 then
	                s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente = " & c_id_nfe_emitente & ")"
	            end if
        
				s_sql = s_sql & _
                        " GROUP BY t_ESTOQUE_ITEM.fabricante, produto" & vbcrlf & _
						"UNION" & vbcrlf & _
						"SELECT" & _
							" t_ESTOQUE_MOVIMENTO.fabricante, produto, estoque, Sum(qtde) As qtde" & _
						" FROM t_ESTOQUE_MOVIMENTO" & _
							" LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
						" WHERE" & _
							" (anulado_status=0)" & _
							" AND (estoque <> 'ETG')" & _
							" AND (estoque <> 'KIT')" & _
							" AND (t_ESTOQUE_MOVIMENTO.fabricante = '" & c_fabricante & "')" & _
							" AND (produto = '" & c_produto & "')" 

                if converte_numero(c_id_nfe_emitente) <> 0 then
	                s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente = " & c_id_nfe_emitente & ")"
	            end if

				s_sql = s_sql & _
                		" GROUP BY t_ESTOQUE_MOVIMENTO.fabricante, produto, estoque"
			else
			'	JOB QUE COLETA OS DADOS FOI EXECUTADO COM SUCESSO NESSE DIA?
				s_sql = "SELECT" & _
							" *" & _
						" FROM t_ESTOQUE_SALDO_DIARIO" & _
						" WHERE" & _
							" (data = " & bd_formata_data(dtSaldoFim+1) & ")" & _
							" AND (fabricante = '----')" & _
							" AND (produto = '--------')"
					
				if r.State <> 0 then r.Close
				r.open s_sql, cn
				if Not r.Eof then 
					blnHaInfoSaldoFim = True
				else
					strQtdeFimEstoqueVenda=TEXTO_INFO_NAO_DISPONIVEL
					strQtdeFimEstoqueShowRoom=TEXTO_INFO_NAO_DISPONIVEL
					strQtdeFimEstoqueDevolucao=TEXTO_INFO_NAO_DISPONIVEL
					strQtdeFimEstoqueSemPresenca=TEXTO_INFO_NAO_DISPONIVEL
					strQtdeFimEstoqueVendido=TEXTO_INFO_NAO_DISPONIVEL
					strQtdeFimEstoqueDanificado=TEXTO_INFO_NAO_DISPONIVEL
					strQtdeFimEstoqueRouboPerda=TEXTO_INFO_NAO_DISPONIVEL
					end if
			
				if Not blnHaInfoSaldoFim then
					s_sql = ""
				else
				'	IMPORTANTE: A GERAÇÃO DOS DADOS QUE ALIMENTA A T_ESTOQUE_SALDO_DIARIO É
				'			    FEITA DURANTE A MADRUGADA
					s_sql = "SELECT" & _
								" fabricante, produto, estoque, Coalesce(Sum(qtde), 0) As qtde" & _
							" FROM t_ESTOQUE_SALDO_DIARIO" & _                           
							" WHERE" & _
								" (data = " & bd_formata_data(dtSaldoFim+1) & ")" & _
								" AND (fabricante = '" & c_fabricante & "')" & _
								" AND (produto = '" & c_produto & "')" & _                 
                    		" GROUP BY fabricante, produto, estoque"

					end if
				end if
				
			if s_sql <> "" then
				if r.State <> 0 then r.Close
				r.open s_sql, cn
				do while Not r.Eof
					s_estoque=Trim("" & r("estoque"))
					
					if s_estoque = ID_ESTOQUE_VENDA then
						strQtdeFimEstoqueVenda=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_VENDIDO then
						strQtdeFimEstoqueVendido=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_SEM_PRESENCA then
						strQtdeFimEstoqueSemPresenca=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_SHOW_ROOM then
						strQtdeFimEstoqueShowRoom=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_DANIFICADOS then
						strQtdeFimEstoqueDanificado=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_DEVOLUCAO then
						strQtdeFimEstoqueDevolucao=formata_inteiro(r("qtde"))
					elseif s_estoque = ID_ESTOQUE_ROUBO then
						strQtdeFimEstoqueRouboPerda=formata_inteiro(r("qtde"))
						end if
					
					r.MoveNext
					loop
				end if ' if (tem sql)
			end if  ' if (tem data término)

		cab_table = "<TABLE class='Q' style='border-bottom:0px;' CellSpacing=0 CellPadding=0>" & chr(13)
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD class='MDB tdSaldoEstoque' style='vertical-align:bottom'><P class='R'>Estoque</P></TD>" & chr(13) & _
			  "		<TD class='MDB tdSaldoQtdInicio' style='vertical-align:bottom'><P class='Rd'>Saldo Inicial em " & c_dt_inicio & "</P></TD>" & chr(13)
		if (dtSaldoFim = Date) then
			cab = cab & _
			  "		<TD class='MB tdSaldoQtdFim' style='vertical-align:bottom'><P class='Rd'>Saldo Atual em " & formata_data_hora(Now) & "</P></TD>" & chr(13)
		else
			cab = cab & _
			  "		<TD class='MB tdSaldoQtdFim' style='vertical-align:bottom'><P class='Rd'>Saldo Final em " & c_dt_termino & "</P></TD>" & chr(13)
			end if
			
		cab = cab & _
			  "	</TR>" & chr(13)

		x = x & _
			"<br>" & chr(13) & _
			cab_table & chr(13) & _
			cab & chr(13)
			
		'Estoque de Venda
		x = x & _
			"	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MDB tdSaldoEstoque'><P class='Cn'>" & x_estoque(ID_ESTOQUE_VENDA) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB tdSaldoQtdInicio'><P class='Cd'>" & strQtdeInicioEstoqueVenda & "</P></TD>" & chr(13) & _
				"		<TD class='MB tdSaldoQtdFim'><P class='Cd'>" & strQtdeFimEstoqueVenda & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
				
		'Estoque Vendido
		x = x & _
			"	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MDB tdSaldoEstoque'><P class='Cn'>" & x_estoque(ID_ESTOQUE_VENDIDO) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB tdSaldoQtdInicio'><P class='Cd'>" & strQtdeInicioEstoqueVendido & "</P></TD>" & chr(13) & _
				"		<TD class='MB tdSaldoQtdFim'><P class='Cd'>" & strQtdeFimEstoqueVendido & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		
		'Show-room
		x = x & _
			"	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MDB tdSaldoEstoque'><P class='Cn'>" & x_estoque(ID_ESTOQUE_SHOW_ROOM) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB tdSaldoQtdInicio'><P class='Cd'>" & strQtdeInicioEstoqueShowRoom & "</P></TD>" & chr(13) & _
				"		<TD class='MB tdSaldoQtdFim'><P class='Cd'>" & strQtdeFimEstoqueShowRoom & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			
		'Danificados
		x = x & _
			"	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MDB tdSaldoEstoque'><P class='Cn'>" & x_estoque(ID_ESTOQUE_DANIFICADOS) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB tdSaldoQtdInicio'><P class='Cd'>" & strQtdeInicioEstoqueDanificado & "</P></TD>" & chr(13) & _
				"		<TD class='MB tdSaldoQtdFim'><P class='Cd'>" & strQtdeFimEstoqueDanificado & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

		'Devolução
		x = x & _
			"	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MDB tdSaldoEstoque'><P class='Cn'>" & x_estoque(ID_ESTOQUE_DEVOLUCAO) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB tdSaldoQtdInicio'><P class='Cd'>" & strQtdeInicioEstoqueDevolucao & "</P></TD>" & chr(13) & _
				"		<TD class='MB tdSaldoQtdFim'><P class='Cd'>" & strQtdeFimEstoqueDevolucao & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

		'Roubo/Perda
		x = x & _
			"	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MDB tdSaldoEstoque'><P class='Cn'>" & x_estoque(ID_ESTOQUE_ROUBO) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB tdSaldoQtdInicio'><P class='Cd'>" & strQtdeInicioEstoqueRouboPerda & "</P></TD>" & chr(13) & _
				"		<TD class='MB tdSaldoQtdFim'><P class='Cd'>" & strQtdeFimEstoqueRouboPerda & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

		'Sem Presença
		x = x & _
			"	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MDB tdSaldoEstoque'><P class='Cn'>" & x_estoque(ID_ESTOQUE_SEM_PRESENCA) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB tdSaldoQtdInicio'><P class='Cd'>" & strQtdeInicioEstoqueSemPresenca & "</P></TD>" & chr(13) & _
				"		<TD class='MB tdSaldoQtdFim'><P class='Cd'>" & strQtdeFimEstoqueSemPresenca & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

		x = x & _
			"</table>" & chr(13)
			
		end if  ' if (produto específico)
	
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

<style type="text/css">
.tdData{
	vertical-align: top;
	width: 60px;
	}
.tdCD{
	vertical-align: top;
	width: 50px;
	}
.tdProd{
	vertical-align: top;
	width: 50px;
	}
.tdLoja{
	vertical-align: top;
	width: 50px;
	}
.tdPedido{
	vertical-align: top;
	width: 70px;
	}
.tdOperador{
	vertical-align: top;
	width: 80px;
	}
.tdOp{
	vertical-align: top;
	width: 130px;
	}
.tdEstOrigem{
	vertical-align: top;
	width: 80px;
	}
.tdEstDestino{
	vertical-align: top;
	width: 80px;
	}
.tdQtd{
	vertical-align: top;
	width: 36px;
	}
.tdSaldoEstoque{
	vertical-align: top;
	width: 100px;
	}
.tdSaldoQtdInicio{
	vertical-align: top;
	width: 80px;
	}
.tdSaldoQtdFim{
	vertical-align: top;
	width: 80px;
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
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">
<input type="hidden" name="c_lista_loja" id="c_lista_loja" value="<%=c_lista_loja%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_ENTRADA" id="ckb_OP_ESTOQUE_LOG_ENTRADA" value="<%=ckb_OP_ESTOQUE_LOG_ENTRADA%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_VENDA" id="ckb_OP_ESTOQUE_LOG_VENDA" value="<%=ckb_OP_ESTOQUE_LOG_VENDA%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA" id="ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA" value="<%=ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT" id="ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT" value="<%=ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_TRANSFERENCIA" id="ckb_OP_ESTOQUE_LOG_TRANSFERENCIA" value="<%=ckb_OP_ESTOQUE_LOG_TRANSFERENCIA%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_ENTREGA" id="ckb_OP_ESTOQUE_LOG_ENTREGA" value="<%=ckb_OP_ESTOQUE_LOG_ENTREGA%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_DEVOLUCAO" id="ckb_OP_ESTOQUE_LOG_DEVOLUCAO" value="<%=ckb_OP_ESTOQUE_LOG_DEVOLUCAO%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_ESTORNO" id="ckb_OP_ESTOQUE_LOG_ESTORNO" value="<%=ckb_OP_ESTOQUE_LOG_ESTORNO%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA" id="ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA" value="<%=ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_SPLIT" id="ckb_OP_ESTOQUE_LOG_SPLIT" value="<%=ckb_OP_ESTOQUE_LOG_SPLIT%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA" id="ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA" value="<%=ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE" id="ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE" value="<%=ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM" id="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM" value="<%=ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA" id="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA" value="<%=ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA" id="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA" value="<%=ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA" id="ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA" value="<%=ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA%>">
<input type="hidden" name="ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS" id="ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS" value="<%=ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="704" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Resumo de Operações Entre Estoques</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='704' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
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

' EMPRESA
	s = ""
	if converte_numero(c_id_nfe_emitente) <> 0 then
		s = obtem_apelido_empresa_NFe_emitente(c_id_nfe_emitente)
		end if
	if s = "" then s = "todas"
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

'	OPERAÇÕES
	s_filtro_operacao = ""

	s = ckb_OP_ESTOQUE_LOG_ENTRADA
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_VENDA
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_TRANSFERENCIA
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_DEVOLUCAO
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_SPLIT
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_ESTORNO
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_ENTREGA
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA
	if s <> "" then
		s = "Roubo ou Perda Total"
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	s = ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA
	if s <> "" then
		s = x_operacao_log_estoque(s)
		s = substitui_caracteres(s, " ", "&nbsp;")
		if s_filtro_operacao <> "" then s_filtro_operacao = s_filtro_operacao & ",&nbsp; "
		s_filtro_operacao = s_filtro_operacao & s
		end if

	if s_filtro_operacao = "" then s_filtro_operacao = "nenhuma"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Operações:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s_filtro_operacao & "</p></td></tr>" & chr(13)

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
<table width="704" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="704" cellSpacing="0">
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
