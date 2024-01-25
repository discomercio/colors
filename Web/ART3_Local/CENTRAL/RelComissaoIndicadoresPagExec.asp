<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L C O M I S S A O I N D I C A D O R E S P A G E X E C . A S P
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

	' CONECTA COM O BANCO DE DADOS
	dim cn, rs, rs2, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	' VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	dim s, s_aux, s_filtro, sJsRetornaTotalBloco
	dim c_dt_entregue_mes, c_dt_entregue_ano, dtDataRef, mes, ano
	dim ckb_comissao_paga_sim, ckb_comissao_paga_nao
	dim ckb_st_pagto_pago, ckb_st_pagto_nao_pago, ckb_st_pagto_pago_parcial
	dim c_vendedor, v_vendedor
	dim rb_visao, blnVisaoSintetica
	dim aviso

	alerta = ""
	sJsRetornaTotalBloco = ""

	c_dt_entregue_mes = Trim(Request.Form("c_dt_entregue_mes"))
	c_dt_entregue_ano = Trim(Request.Form("c_dt_entregue_ano"))
	c_vendedor = Trim(Request.Form("c_vendedor"))
	ckb_comissao_paga_sim = Trim(Request.Form("ckb_comissao_paga_sim"))
	ckb_comissao_paga_nao = Trim(Request.Form("ckb_comissao_paga_nao"))
	ckb_st_pagto_pago = Trim(Request.Form("ckb_st_pagto_pago"))
	ckb_st_pagto_nao_pago = Trim(Request.Form("ckb_st_pagto_nao_pago"))
	ckb_st_pagto_pago_parcial = Trim(Request.Form("ckb_st_pagto_pago_parcial"))
	rb_visao = Trim(Request.Form("rb_visao"))
	
	blnVisaoSintetica = False
	if rb_visao = "SINTETICA" then blnVisaoSintetica = True

	if c_dt_entregue_mes = "" then
		alerta = texto_add_br(alerta)
		alerta = alerta & "Mês de competência não informado"
	elseif converte_numero(c_dt_entregue_mes) = 0 then
		alerta = texto_add_br(alerta)
		alerta = alerta & "Mês de competência é inválido"
		end if
	
	if c_dt_entregue_ano = "" then
		alerta = texto_add_br(alerta)
		alerta = alerta & "Ano de competência não informado"
	elseif converte_numero(c_dt_entregue_ano) = 0 then
		alerta = texto_add_br(alerta)
		alerta = alerta & "Ano de competência é inválido"
		end if

	if c_vendedor = "" then
		alerta = texto_add_br(alerta)
		alerta = alerta & "Nenhum vendedor foi selecionado"
		end if

	if alerta = "" then
		mes = c_dt_entregue_mes
		ano = c_dt_entregue_ano

		if len(Cstr(mes)) = 1 then mes =  "0" & Cstr(mes)
		s = "01/" & Cstr(mes) & "/" & Cstr(ano)
		dtDataRef = strToDate(s)
		dtDataRef = DateAdd("m", 1, dtDataRef)

		v_vendedor = Split(c_vendedor, ", ")
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'

sub consulta_executa
const VENDA_NORMAL = "VEN"
const DEVOLUCAO = "DEV"
const PERDA = "PER"
const TABLE_MARGIN_N1 = ""
const TABLE_MARGIN_N2 = "			"
dim r
dim i, j
dim s, s_aux, s_sql, x, cab_table, cab, indicador_a, n_reg, n_reg_total, qtde_indicadores, indicador_bloco_anterior
dim vl_preco_venda, vl_sub_total_preco_venda, vl_total_preco_venda
dim vl_preco_NF, vl_sub_total_preco_NF, vl_total_preco_NF
dim vl_RT, vl_sub_total_RT, vl_total_RT
dim vl_RA, vl_sub_total_RA, vl_total_RA
dim vl_RA_liquido, vl_sub_total_RA_liquido, vl_total_RA_liquido
dim vl_RA_diferenca, vl_sub_total_RA_diferenca, vl_total_RA_diferenca
dim perc_RT, lista_operacao
dim s_where_base, s_where_aux, s_where_venda, s_where_devolucao, s_where_perdas, s_where_vendedor
dim s_where_comissao_paga, s_where_comissao_descontada, s_where_st_pagto
dim s_cor, s_sinal, s_cor_sinal
dim s_banco, s_banco_nome, s_banco_codigo_descricao, s_agencia, s_conta, s_favorecido
dim s_nome_cliente
dim s_checked, idx_bloco, s_new_cab
dim lista_completa_pedidos, lista_qtde_reg_descontos
dim lista_vl_comissao, lista_vl_RA_bruto, lista_vl_RA_liquido, lista_vl_total_comissao, lista_vl_total_comissao_arredondado, s_lista_meio_pagto, lista_vl_pedido, lista_vl_total_RA, lista_vl_total_RA_arredondado
dim sub_total_comissao, s_disabled, msg_desconto
dim lista_indicador_com_desconto, lista_indicador_negativo, lista_vl_total_desconto_planilha
dim vl_sub_total_RT_arredondado, vl_sub_total_RA_arredondado
dim qtde_rel_vendedor, vendedor_processado, lista_vl_total_pagto, sub_total_com, v_desconto_descricao(), v_desconto_valor(), contador, vl_total_desc_planilha, qtde_registro_desc
dim vIndicador, vPedido, s_dados, s_erro

	redim vIndicador(0)
	set vIndicador(UBound(vIndicador)) = new cl_REL_PEDIDOS_INDICADORES_INFO_INDICADOR
	inicializa_cl_REL_PEDIDOS_INDICADORES_INFO_INDICADOR vIndicador(UBound(vIndicador))

	redim vPedido(0)
	set vPedido(UBound(vPedido)) = new cl_REL_PEDIDOS_INDICADORES_INFO_PEDIDO
	inicializa_cl_REL_PEDIDOS_INDICADORES_INFO_PEDIDO vPedido(UBound(vPedido))

	vl_total_desc_planilha = 0
	lista_indicador_com_desconto = ""
	lista_indicador_negativo = ""
	s_disabled = ""
	lista_vl_total_desconto_planilha = ""
	qtde_rel_vendedor = 0
	vendedor_processado = ""
	aviso = ""
	lista_operacao = ""

	s = "SELECT DISTINCT" & _
			" vendedor" & _
		" FROM t_COMISSAO_INDICADOR_N2" & _
		" WHERE" & _
			" (competencia_ano = " & c_dt_entregue_ano & ")" & _
			" AND (competencia_mes = " & c_dt_entregue_mes & ")"
	rs.Open s, cn
	do while not rs.Eof
		for i=Lbound(v_vendedor) to Ubound(v_vendedor)
			if Trim("" & v_vendedor(i)) = Trim("" & rs("vendedor")) then
				if vendedor_processado <> "" then vendedor_processado = vendedor_processado & ", "
				vendedor_processado = vendedor_processado & Trim("" & rs("vendedor"))
				qtde_rel_vendedor = qtde_rel_vendedor + 1
				end if
			next
		rs.MoveNext
		loop

	if rs.State <> 0 then rs.Close

	if qtde_rel_vendedor = 1 then
		aviso = "<div class='MtAlerta' style='width:600px;font-weight:bold;' align='center'>" & chr(13) & _
				"<p style='margin:5px 2px 5px 2px;'>" & _
				"O vendedor " & vendedor_processado & " já foi processado no mês de competência informado" & _
				"<a href='RelComissaoIndicadoresConsultaPedidoExec.asp?vendedor=" & vendedor_processado & "&ano_competencia=" & c_dt_entregue_ano & "&mes_competencia=" & c_dt_entregue_mes & "' style='text-decoration:underline;color:#CCC'>" & _
				"<br />Clique aqui para consultar o relatório gerado.</a>" & chr(13) & _
				"</p>" & chr(13)& _
				"</div><br />"
	elseif qtde_rel_vendedor > 1 then
		aviso =  "<div class='MtAlerta' style='width:600px;font-weight:bold;' align='center'>" & chr(13) & _
				"<p style='margin:5px 2px 5px 2px;'>" & _
				"Os vendedores " & vendedor_processado & " já foram processados no mês de competência informado" & _
				"</p>" & chr(13) & _
				"</div><br />"
	end if

	if aviso <> "" then
		Response.Write aviso
		exit sub
		end if


	' CRITÉRIOS COMUNS
	' PROCESSA SOMENTE OS PARCEIROS HABILITADOS PARA RECEBEREM A COMISSÃO VIA CARTÃO
	s_where_base = " (LEN(Coalesce(t_PEDIDO__BASE.indicador, '')) > 0)" & _
					" AND (t_ORCAMENTISTA_E_INDICADOR.comissao_cartao_status = 1)"

	if c_vendedor <> "" then
		s_where_vendedor = ""
		for j = LBound(v_vendedor) to UBound(v_vendedor)
			if s_where_vendedor <> "" then s_where_vendedor = s_where_vendedor & ", "
			s_where_vendedor = s_where_vendedor & "'" & Trim(replace(v_vendedor(j), "'", "''")) & "'"
			next
		if s_where_vendedor <> "" then
			s_where_vendedor = " (t_PEDIDO__BASE.vendedor IN (" & s_where_vendedor & ")) "
			if s_where_base <> "" then s_where_base = s_where_base & " AND"
			s_where_base = s_where_base & s_where_vendedor
			end if
		end if

	' CRITÉRIO: COMISSÃO PAGA
	' A) VENDAS
	s_where_comissao_paga = " (t_PEDIDO.comissao_paga = 0)"
		
	' B) PERDAS/DEVOLUÇÕES
	s_where_comissao_descontada = " (comissao_descontada = 0)"

	' CRITÉRIO: STATUS DE PAGAMENTO
	s_where_st_pagto = " ((t_PEDIDO__BASE.st_pagto = 'S') AND (t_PEDIDO__BASE.dt_st_pagto < " & bd_formata_data(dtDataRef) & "))"


	' CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	s_where_venda = ""
	if (c_dt_entregue_mes <> "") And (c_dt_entregue_ano <> "") then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data < " & bd_formata_data(dtDataRef) & ")"
		end if

	if s_where_comissao_paga <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (" & s_where_comissao_paga & ")"
		end if

	if s_where_st_pagto <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (" & s_where_st_pagto & ")"
		end if

	' CRITÉRIOS PARA DEVOLUÇÕES
	s_where_devolucao = ""
	if (c_dt_entregue_mes <> "") And (c_dt_entregue_ano <> "") then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(dtDataRef) & ")"
		end if

	if s_where_comissao_descontada <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (" & s_where_comissao_descontada & ")"
		end if

	' CRITÉRIOS PARA PERDAS
	s_where_perdas = ""
	if (c_dt_entregue_mes <> "") And (c_dt_entregue_ano <> "") then
		if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		s_where_perdas = s_where_perdas & " (t_PEDIDO_PERDA.data < " & bd_formata_data(dtDataRef) & ")"
		end if

	if s_where_comissao_descontada <> "" then
		if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		s_where_perdas = s_where_perdas & " (" & s_where_comissao_descontada & ")"
		end if

	' VENDAS NORMAIS
	s_where_aux = s_where_base
	if (s_where_aux <> "") And (s_where_venda <> "") then s_where_aux = s_where_aux & " AND"
	s_where_aux = s_where_aux & s_where_venda
	if s_where_aux <> "" then s_where_aux = " AND" & s_where_aux

	s_sql = "SELECT" & _
				" '" & VENDA_NORMAL & "' AS operacao" & _
				", t_PEDIDO.pedido AS id_registro" & _
				", t_PEDIDO.comissao_paga AS status_comissao" & _
				", Coalesce(t_ORCAMENTISTA_E_INDICADOR.desempenho_nota,'') AS desempenho_nota" & _
				", t_ORCAMENTISTA_E_INDICADOR.banco" & _
				", t_PEDIDO__BASE.indicador" & _
				", t_ORCAMENTISTA_E_INDICADOR.Id AS IdIndicador" & _
				", t_PEDIDO__BASE.vendedor" & _
				", t_USUARIO.Id AS IdVendedor"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				", t_PEDIDO.endereco_nome AS nome" & _
				", t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas"
	else
		s_sql = s_sql & _
				", t_CLIENTE.nome" & _
				", t_CLIENTE.nome_iniciais_em_maiusculas"
		end if

	s_sql = s_sql & _
				", t_PEDIDO.loja AS loja" & _
				", t_PEDIDO.numero_loja" & _
				", t_PEDIDO.entregue_data AS data" & _
				", t_PEDIDO.pedido AS pedido" & _
				", t_PEDIDO.orcamento AS orcamento" & _
				", t_PEDIDO__BASE.perc_RT" & _
				", t_PEDIDO__BASE.st_pagto" & _
				", t_PEDIDO__BASE.vl_total_RA_liquido" & _
				", t_PEDIDO__BASE.st_tem_desagio_RA" & _
				", t_PEDIDO__BASE.perc_desagio_RA_liquida" & _
				", Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.preco_venda) AS total_preco_venda" & _
				", Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.preco_NF) AS total_preco_NF" & _
			" FROM t_PEDIDO" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
				" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
				" INNER JOIN t_USUARIO ON (t_USUARIO.usuario = t_PEDIDO__BASE.vendedor)" & _
				" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			" WHERE" & _
				" (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
				s_where_aux & _
			" GROUP BY" & _
				" t_PEDIDO.pedido" & _
				", t_PEDIDO.comissao_paga" & _
				", t_ORCAMENTISTA_E_INDICADOR.desempenho_nota" & _
				", t_ORCAMENTISTA_E_INDICADOR.banco" & _
				", t_PEDIDO__BASE.indicador" & _
				", t_ORCAMENTISTA_E_INDICADOR.Id" & _
				", t_PEDIDO__BASE.vendedor" & _
				", t_USUARIO.Id"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				", t_PEDIDO.endereco_nome" & _
				", t_PEDIDO.endereco_nome_iniciais_em_maiusculas"
	else
		s_sql = s_sql & _
				", t_CLIENTE.nome" & _
				", t_CLIENTE.nome_iniciais_em_maiusculas"
		end if

	s_sql = s_sql & _
				", t_PEDIDO.loja" & _
				", t_PEDIDO.numero_loja" & _
				", t_PEDIDO.entregue_data" & _
				", t_PEDIDO.pedido" & _
				", t_PEDIDO.orcamento" & _
				", t_PEDIDO__BASE.perc_RT" & _
				", t_PEDIDO__BASE.st_pagto" & _
				", t_PEDIDO__BASE.vl_total_RA_liquido" & _
				", t_PEDIDO__BASE.st_tem_desagio_RA" & _
				", t_PEDIDO__BASE.perc_desagio_RA_liquida"

	' ITENS DEVOLVIDOS
	s_where_aux = s_where_base
	if (s_where_aux <> "") And (s_where_devolucao <> "") then s_where_aux = s_where_aux & " AND"
	s_where_aux = s_where_aux & s_where_devolucao
	if s_where_aux <> "" then s_where_aux = " WHERE " & s_where_aux

	s_sql = s_sql & _
			" UNION ALL " & _
			"SELECT" & _
				" '" & DEVOLUCAO & "' AS operacao" & _
				", t_PEDIDO_ITEM_DEVOLVIDO.id AS id_registro" & _
				", t_PEDIDO_ITEM_DEVOLVIDO.comissao_descontada AS status_comissao" & _
				", Coalesce(t_ORCAMENTISTA_E_INDICADOR.desempenho_nota,'') AS desempenho_nota" & _
				", t_ORCAMENTISTA_E_INDICADOR.banco" & _
				", t_PEDIDO__BASE.indicador" & _
				", t_ORCAMENTISTA_E_INDICADOR.Id AS IdIndicador" & _
				", t_PEDIDO__BASE.vendedor" & _
				", t_USUARIO.Id AS IdVendedor"
			
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				", t_PEDIDO.endereco_nome AS nome" & _
				", t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas"
	else
		s_sql = s_sql & _
				", t_CLIENTE.nome" & _
				", t_CLIENTE.nome_iniciais_em_maiusculas"
		end if

	s_sql = s_sql & _
				", t_PEDIDO.loja AS loja" & _
				", t_PEDIDO.numero_loja" & _
				", t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data" & _
				", t_PEDIDO.pedido AS pedido" & _
				", t_PEDIDO.orcamento AS orcamento" & _
				", t_PEDIDO__BASE.perc_RT" & _
				", t_PEDIDO__BASE.st_pagto" & _
				", t_PEDIDO__BASE.vl_total_RA_liquido" & _
				", t_PEDIDO__BASE.st_tem_desagio_RA" & _
				", t_PEDIDO__BASE.perc_desagio_RA_liquida" & _
				", Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS total_preco_venda" & _
				", Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_NF) AS total_preco_NF" & _
			" FROM t_PEDIDO" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido)" & _
				" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
				" INNER JOIN t_USUARIO ON (t_USUARIO.usuario = t_PEDIDO__BASE.vendedor)" & _
				" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			s_where_aux & _
			" GROUP BY" & _
				" t_PEDIDO_ITEM_DEVOLVIDO.id" & _
				", t_PEDIDO_ITEM_DEVOLVIDO.comissao_descontada" & _
				", t_ORCAMENTISTA_E_INDICADOR.desempenho_nota" & _
				", t_ORCAMENTISTA_E_INDICADOR.banco" & _
				", t_PEDIDO__BASE.indicador" & _
				", t_ORCAMENTISTA_E_INDICADOR.Id" & _
				", t_PEDIDO__BASE.vendedor" & _
				", t_USUARIO.Id"
			
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				", t_PEDIDO.endereco_nome" & _
				", t_PEDIDO.endereco_nome_iniciais_em_maiusculas"
	else
		s_sql = s_sql & _
				", t_CLIENTE.nome" & _
				", t_CLIENTE.nome_iniciais_em_maiusculas"
		end if

	s_sql = s_sql & _
				", t_PEDIDO.loja" & _
				", t_PEDIDO.numero_loja" & _
				", t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data" & _
				", t_PEDIDO.pedido" & _
				", t_PEDIDO.orcamento" & _
				", t_PEDIDO__BASE.perc_RT" & _
				", t_PEDIDO__BASE.st_pagto" & _
				", t_PEDIDO__BASE.vl_total_RA_liquido" & _
				", t_PEDIDO__BASE.st_tem_desagio_RA" & _
				", t_PEDIDO__BASE.perc_desagio_RA_liquida"

	' PERDAS
	s_where_aux = s_where_base
	if (s_where_aux <> "") And (s_where_perdas <> "") then s_where_aux = s_where_aux & " AND"
	s_where_aux = s_where_aux & s_where_perdas
	if s_where_aux <> "" then s_where_aux = " WHERE " & s_where_aux
		
	s_sql = s_sql & _
			" UNION ALL " & _
			"SELECT" & _
				" '" & PERDA & "' AS operacao" & _
				", t_PEDIDO_PERDA.id AS id_registro" & _
				", t_PEDIDO_PERDA.comissao_descontada AS status_comissao" & _
				", Coalesce(t_ORCAMENTISTA_E_INDICADOR.desempenho_nota,'') AS desempenho_nota" & _
				", t_ORCAMENTISTA_E_INDICADOR.banco" & _
				", t_PEDIDO__BASE.indicador" & _
				", t_ORCAMENTISTA_E_INDICADOR.Id AS IdIndicador" & _
				", t_PEDIDO__BASE.vendedor" & _
				", t_USUARIO.Id AS IdVendedor"
			
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				", t_PEDIDO.endereco_nome AS nome" & _
				", t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas"
	else
		s_sql = s_sql & _
				", t_CLIENTE.nome" & _
				", t_CLIENTE.nome_iniciais_em_maiusculas"
		end if

	s_sql = s_sql & _
				", t_PEDIDO.loja AS loja" & _
				", t_PEDIDO.numero_loja" & _
				", t_PEDIDO_PERDA.data AS data" & _
				", t_PEDIDO.pedido AS pedido" & _
				", t_PEDIDO.orcamento AS orcamento" & _
				", t_PEDIDO__BASE.perc_RT" & _
				", t_PEDIDO__BASE.st_pagto" & _
				", t_PEDIDO__BASE.vl_total_RA_liquido" & _
				", t_PEDIDO__BASE.st_tem_desagio_RA" & _
				", t_PEDIDO__BASE.perc_desagio_RA_liquida" & _
				", Sum(-t_PEDIDO_PERDA.valor) AS total_preco_venda" & _
				", Sum(-t_PEDIDO_PERDA.valor) AS total_preco_NF" & _
			" FROM t_PEDIDO" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" INNER JOIN t_PEDIDO_PERDA ON (t_PEDIDO.pedido=t_PEDIDO_PERDA.pedido)" & _
				" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
				" INNER JOIN t_USUARIO ON (t_USUARIO.usuario = t_PEDIDO__BASE.vendedor)" & _
				" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			s_where_aux & _
			" GROUP BY" & _
				" t_PEDIDO_PERDA.id" & _
				", t_PEDIDO_PERDA.comissao_descontada" & _
				", t_ORCAMENTISTA_E_INDICADOR.desempenho_nota" & _
				", t_ORCAMENTISTA_E_INDICADOR.banco" & _
				", t_PEDIDO__BASE.indicador" & _
				", t_ORCAMENTISTA_E_INDICADOR.Id" & _
				", t_PEDIDO__BASE.vendedor" & _
				", t_USUARIO.Id"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				", t_PEDIDO.endereco_nome" & _
				", t_PEDIDO.endereco_nome_iniciais_em_maiusculas"
	else
		s_sql = s_sql & _
				", t_CLIENTE.nome" & _
				", t_CLIENTE.nome_iniciais_em_maiusculas"
		end if

	s_sql = s_sql & _
				", t_PEDIDO.loja" & _
				", t_PEDIDO.numero_loja" & _
				", t_PEDIDO_PERDA.data" & _
				", t_PEDIDO.pedido" & _
				", t_PEDIDO.orcamento" & _
				", t_PEDIDO__BASE.perc_RT" & _
				", t_PEDIDO__BASE.st_pagto" & _
				", t_PEDIDO__BASE.vl_total_RA_liquido" & _
				", t_PEDIDO__BASE.st_tem_desagio_RA" & _
				", t_PEDIDO__BASE.perc_desagio_RA_liquida"

	s_sql = "SELECT " & _
				"*" & _
			" FROM (" & _
				s_sql & _
				") t" & _
			" ORDER BY" & _
				" t.vendedor" & _
				", t.desempenho_nota" & _
				", t.indicador" & _
				", t.numero_loja" & _
				", t.data" & _
				", t.pedido" & _
				", t.total_preco_venda DESC"

	' CABEÇALHO
	cab_table = TABLE_MARGIN_N2 & "<table border='0' cellspacing='0' id='tableDados'>" & chr(13)
	cab = TABLE_MARGIN_N2 & "	<tr style='background:azure' nowrap>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td class='MDTE tdLoja' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>Loja</span></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td class='MTD tdOrcamento' align='left' valign='bottom' nowrap><span class='R VISAO_ANALIT'>Nº Orçam</span></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td class='MTD tdPedido' align='left' valign='bottom' nowrap><span class='R VISAO_ANALIT'>Nº Pedido</span></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td class='MTD tdData' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>Data</span></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td class='MTD tdVlPedido' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL Pedido</span></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td class='MTD tdVlRT' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>COM (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td class='MTD tdVlRABruto' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Bruto (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td class='MTD tdVlRALiq' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Líq (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td class='MTD tdVlRADif' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Dif (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td class='MTD tdStPagto' align='left' valign='bottom'><span class='R VISAO_ANALIT' style='font-weight:bold;'>St Pagto</span></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td class='MTD tdSinal' align='center' valign='bottom'><span class='Rc VISAO_ANALIT' style='font-weight:bold;'>+/-</span></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "		<td valign='bottom' class='notPrint BkgWhite' align='left'>&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & "_NNNNN_" & chr(34) & ");' title='exibe ou oculta os dados'><img src='../botao/view_bottom.png' border='0'></a></td>" & chr(13) & _
		TABLE_MARGIN_N2 & "	</tr>" & chr(13)
	
	x = TABLE_MARGIN_N2 & "<br />" & chr(13)
	n_reg = 0
	n_reg_total = 0
	idx_bloco = 0
	qtde_indicadores = 0
	vl_sub_total_preco_venda = 0
	vl_total_preco_venda = 0
	vl_sub_total_preco_NF = 0
	vl_total_preco_NF = 0
	vl_sub_total_RT = 0
	vl_total_RT = 0
	vl_sub_total_RA = 0
	vl_total_RA = 0
	vl_sub_total_RA_liquido = 0
	vl_total_RA_liquido = 0
	vl_sub_total_RA_diferenca = 0
	vl_total_RA_diferenca = 0
	lista_completa_pedidos = ""
	lista_vl_comissao = ""
	lista_vl_RA_bruto = ""
	lista_vl_RA_liquido = ""
	lista_vl_pedido = ""
	lista_vl_total_comissao = ""
	lista_vl_total_comissao_arredondado = ""
	lista_vl_total_RA = ""
	lista_vl_total_RA_arredondado = ""
	lista_qtde_reg_descontos = ""
	sub_total_comissao = 0
	vl_sub_total_RT_arredondado = 0
	indicador_a = "XXXXXXXXXXXX"
	sub_total_com = 0
	lista_vl_total_pagto = ""

	set r = cn.execute(s_sql)
	do while Not r.Eof
		' MUDOU DE INDICADOR?
		if Trim("" & r("indicador")) <> indicador_a then
			' FECHA TABELA DO INDICADOR ANTERIOR
			if n_reg_total > 0 then
				s_checked = ""
				s_disabled = ""
				lista_vl_total_comissao = lista_vl_total_comissao & formata_moeda(vl_sub_total_RT) & ";"
				lista_vl_total_RA = lista_vl_total_RA & formata_moeda(vl_sub_total_RA_liquido) & ";"
				if sub_total_com >= 0 then
					s_checked = " checked"
				else
					s_disabled = " disabled"
					lista_indicador_negativo = lista_indicador_negativo & indicador_bloco_anterior & ", "
					vIndicador(UBound(vIndicador)).indicador_negativo = True
					end if

				msg_desconto = ""
				vl_total_desc_planilha = 0
				contador = 0
				qtde_registro_desc = 0
				Erase v_desconto_descricao
				Erase v_desconto_valor

				s = "SELECT" & _
						" descricao" & _
						", valor" & _
						", ordenacao" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO" & _
					" WHERE" & _
						" (apelido = '" & indicador_bloco_anterior & "')" & _
					" ORDER BY" & _
						" ordenacao"
				if rs2.State <> 0 then rs2.Close
				rs2.Open s, cn
				if Not rs2.Eof then
					do while Not rs2.Eof
						redim preserve v_desconto_descricao(contador)
						redim preserve  v_desconto_valor(contador)
						v_desconto_descricao(contador) = Trim("" & rs2("descricao"))
						v_desconto_valor(contador) = rs2("valor")
						vl_total_desc_planilha = vl_total_desc_planilha + v_desconto_valor(contador)
						qtde_registro_desc = qtde_registro_desc + 1
						contador = contador + 1
						rs2.MoveNext
						loop

					s_checked = ""
					s_disabled = " disabled"
					msg_desconto = "Planilha de Desconto:&nbsp; " & Cstr(qtde_registro_desc) & " registro(s) no valor total de&nbsp;" & formata_moeda(vl_total_desc_planilha)
					lista_indicador_com_desconto = lista_indicador_com_desconto & indicador_bloco_anterior & ", "
					lista_vl_total_desconto_planilha = lista_vl_total_desconto_planilha & formata_moeda(vl_total_desc_planilha) & ";"
					lista_qtde_reg_descontos = lista_qtde_reg_descontos & qtde_registro_desc & ";"
					with vIndicador(UBound(vIndicador))
						.indicador_com_desconto = True
						.vl_total_desc_planilha = vl_total_desc_planilha
						.qtde_reg_descontos = qtde_registro_desc
						end with
					end if

				' TOTAL DO INDICADOR
				s_cor = "black"
				if vl_sub_total_preco_venda < 0 then s_cor = "red"
				if vl_sub_total_RT < 0 then s_cor = "red"
				if vl_sub_total_RA < 0 then s_cor = "red"
				if vl_sub_total_RA_liquido < 0 then s_cor = "red"

				x = x & TABLE_MARGIN_N2 & "	<tr style='background: #FFFFDD'>" & chr(13) & _
						TABLE_MARGIN_N2 & "		<td class='MTBE' colspan='4' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
									"TOTAL:</span></td>" & chr(13) & _
						TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _
						TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_preco_venda_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_preco_venda) & "'>" & chr(13) & _
						TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _
						TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissao_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RT) & "'>" & chr(13)

				if (vl_sub_total_RA_arredondado > 0) And (vl_sub_total_RT_arredondado < 0) then
					x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRT_" & CStr(idx_bloco) & "' value='" & CStr(0) & "'>" & chr(13)
				else
					if (vl_sub_total_RA_arredondado >= 0) And (vl_sub_total_RT_arredondado >= 0) then
						x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRT_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RT_arredondado) & "'>" & chr(13)
					else
						x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRT_" & CStr(idx_bloco) & "' value='" & formata_moeda(sub_total_com) & "'>" & chr(13)
						end if
					end if

				x = x &_
					TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</span></td>" & chr(13) & _
					TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_RA_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RA) & "'>" & chr(13)
				
				if (vl_sub_total_RA_arredondado < 0) And (vl_sub_total_RT_arredondado > 0) then
					x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRA_" & CStr(idx_bloco) & "' value='" & CStr(0) & "'>" & chr(13)
				else
					if (vl_sub_total_RA_arredondado >= 0) And (vl_sub_total_RT_arredondado >= 0) then
						x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRA_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RA_arredondado) & "'>" & chr(13)
					else
						x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRA_" & CStr(idx_bloco) & "' value='" & formata_moeda(sub_total_com) & "'>" & chr(13)
						end if
					end if

				x = x & _
					TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _
					TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_RA_liquido_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RA_liquido) & "'>" & chr(13) & _
					TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _
					TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_RA_diferenca_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RA_diferenca) & "'>" & chr(13) & _
					TABLE_MARGIN_N2 & "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
					TABLE_MARGIN_N2 & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					TABLE_MARGIN_N2 & "	</tr>" & chr(13) & _
					TABLE_MARGIN_N2 & "	<tr>" & chr(13)

				' MENSAGEM DE DESCONTO E VALORES DE RT E RA
				x = x & TABLE_MARGIN_N2 & "		<td align='left' colspan='5' nowrap>"
				if msg_desconto <> "" then x = x & "<span class='Cd'><a href='javascript:abreDesconto(" & CStr(idx_bloco) & ")' title='Exibe ou oculta os registros de descontos' style='color:red;'>" & msg_desconto & "</a></span>"
				x = x & TABLE_MARGIN_N2 & "</td>" & chr(13)

				' COMISSAO RT E RA
				if vl_sub_total_RT_arredondado >= 0 then
					s_cor = "black"
				else
					s_cor = "red"
					end if
				x = x & TABLE_MARGIN_N2 & "		<td align='right' nowrap><span class='Rd' style='color:" & s_cor & ";margin-right:0.3;'>" & formata_moeda(vl_sub_total_RT_arredondado) & "</span></td>"& chr(13)
					
				if vl_sub_total_RA_arredondado >= 0 then
					s_cor = "black"
				else
					s_cor = "red"
					end if
				x = x & TABLE_MARGIN_N2 & "		<td align='right' colspan='2' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RA_arredondado) & "</span></td>"& chr(13)

				if sub_total_com >= 0 then
					s_cor = "black"
				else
					s_cor = "red"
					end if
				x = x & TABLE_MARGIN_N2 & "		<td align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>"

				' MEIO DE PAGAMENTO
				if sub_total_com <= 0 then
					s_lista_meio_pagto = s_lista_meio_pagto & " " & ";"
					x = x & "&nbsp;</span></td>" & chr(13)
					if lista_vl_total_pagto <> "" then lista_vl_total_pagto = lista_vl_total_pagto & ";"
					lista_vl_total_pagto = lista_vl_total_pagto & formata_moeda(0)
				else
					vIndicador(UBound(vIndicador)).meio_pagto = "CARD"
					s_lista_meio_pagto = s_lista_meio_pagto & "CARD" & ";"
					x = x & "CARTÃO:</span></td>" & chr(13) &_
							TABLE_MARGIN_N2 & "		<input type='hidden' id='forma_pag_" & CStr(idx_bloco) & "' value='CARD'>" & chr(13)
					x = x & TABLE_MARGIN_N2 & "		<input type='hidden' id='forma_pag_valor_" & CStr(idx_bloco) & "' value='" & formata_moeda(sub_total_com) & "'>" & chr(13)
					if lista_vl_total_pagto <> "" then lista_vl_total_pagto = lista_vl_total_pagto & ";"
					lista_vl_total_pagto = lista_vl_total_pagto & formata_moeda(sub_total_com)
					vIndicador(UBound(vIndicador)).vl_total_pagto = sub_total_com
					end if

				' SUB TOTAL DA COMISSAO
				x = x & TABLE_MARGIN_N2 & "		<td align='right'>"
				if sub_total_com >= 0 then
					x = x & "<span class='Cd' style='color:black;'>" & formata_moeda(sub_total_com) & "</span>"
				else
					x = x & "&nbsp;"
					end if
				x = x & _
					"</td>" & chr(13) & _
					TABLE_MARGIN_N2 & "	</tr>" & chr(13)

				if msg_desconto <> "" then
					x = x & TABLE_MARGIN_N2 & "	<tr>" & chr(13) & _
							TABLE_MARGIN_N2 & "		<td class='table_Desconto' id='table_Desconto_" & CStr(idx_bloco) & "' colspan='15'>" & chr(13) & _
							TABLE_MARGIN_N2 & "			<table border='0'>" & chr(13)
					for contador = 0 to Ubound(v_desconto_descricao)
						x = x & _
							TABLE_MARGIN_N2 & "				<tr>" & chr(13) & _
							TABLE_MARGIN_N2 & "					<td width='15'>&nbsp;</td>" & chr(13) & _
							TABLE_MARGIN_N2 & "					<td align='left' width='400'><span class='Cd' style='color:red;'>" & v_desconto_descricao(contador) & "</span></td>" & chr(13) & _
							TABLE_MARGIN_N2 & "					<td align='left'><span class='Cd' style='color:red;'> R$ " & formata_moeda(v_desconto_valor(contador)) & "</span></td>" & chr(13) & _
							TABLE_MARGIN_N2 & "				</tr>" & chr(13)
						next

					x = x & TABLE_MARGIN_N2 & "			</table>" & chr(13) & _
							TABLE_MARGIN_N2 & "		</td>" & chr(13) & _
							TABLE_MARGIN_N2 & "	</tr>" & chr(13)
					end if 'if msg_desconto <> ""

				x = x & TABLE_MARGIN_N2 & "</table>" & chr(13)

				x = TABLE_MARGIN_N1 & "<table cellpadding='0' cellspacing='0'>" & chr(13) & _
					TABLE_MARGIN_N1 & "	<tr>" & chr(13) & _
					TABLE_MARGIN_N1 & "		<td valign='top'><br />" & chr(13) & _
					TABLE_MARGIN_N1 & "			<input type='checkbox' name='ckb_bloco_indicador' class='CKB_COM' id='ckb_bloco_indicador_" & CStr(idx_bloco) & "' onclick='trata_ckb_onclick();calculaTotalComissao();' value='" & Cstr(vIndicador(Ubound(vIndicador)).IdVendedor) & "|" & Cstr(vIndicador(Ubound(vIndicador)).IdIndicador) & "' " & s_checked & s_disabled & " />" & chr(13) & _
					TABLE_MARGIN_N1 & "		</td>" & chr(13) & _
					TABLE_MARGIN_N1 & "		<td valign='top'>" & chr(13) & _
									x & _
									"		</td>" & chr(13) & _
					TABLE_MARGIN_N1 & "	</tr>" & chr(13) & _
					TABLE_MARGIN_N1 & "</table>" & chr(13)

				lista_vl_total_comissao_arredondado = lista_vl_total_comissao_arredondado & formata_moeda(vl_sub_total_RT_arredondado) & ";"
				lista_vl_total_RA_arredondado = lista_vl_total_RA_arredondado & formata_moeda(vl_sub_total_RA_arredondado) & ";"

				with vIndicador(UBound(vIndicador))
					.vl_total_comissao = vl_sub_total_RT
					.vl_total_RA = vl_sub_total_RA_liquido
					.vl_total_comissao_arredondado = vl_sub_total_RT_arredondado
					.vl_total_RA_arredondado = vl_sub_total_RA_arredondado
					end with

				Response.Write x
				x = TABLE_MARGIN_N2 & "<br />" & chr(13)
				end if 'if n_reg_total > 0


			if vIndicador(UBound(vIndicador)).IdIndicador <> 0 then
				redim preserve vIndicador(UBound(vIndicador)+1)
				set vIndicador(UBound(vIndicador)) = new cl_REL_PEDIDOS_INDICADORES_INFO_INDICADOR
				inicializa_cl_REL_PEDIDOS_INDICADORES_INFO_INDICADOR vIndicador(UBound(vIndicador))
				end if

' TODO			vIndicador(UBound(vIndicador)).vendedor = Trim("" & r("vendedor"))
			vIndicador(UBound(vIndicador)).IdVendedor = r("IdVendedor")
' TODO			vIndicador(UBound(vIndicador)).indicador = Trim("" & r("indicador"))
			vIndicador(UBound(vIndicador)).IdIndicador = r("IdIndicador")

			indicador_a = Trim("" & r("indicador"))
			idx_bloco = idx_bloco + 1
			qtde_indicadores = qtde_indicadores + 1

			s_banco = ""
			s_banco_nome = ""
			s_banco_codigo_descricao = ""
			s_agencia = ""
			s_conta = ""
			s_favorecido = ""

			s_sql = "SELECT" & _
						" *" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (apelido = '" & Trim("" & r("indicador")) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Not rs.Eof then
				s_banco = Trim("" & rs("banco"))
				s_agencia = Trim("" & rs("agencia"))
				if s_agencia <> "" then
					if Trim("" & rs("agencia_dv")) <> "" then s_agencia = s_agencia & "-" & Trim("" & rs("agencia_dv"))
					end if
				s_conta = Trim("" & rs("conta"))
				if s_conta <> "" then
					if Trim("" & rs("conta_operacao")) <> "" then s_conta = Trim("" & rs("conta_operacao")) & "-" & s_conta
					if Trim("" & rs("conta_dv")) <> "" then s_conta = s_conta & Trim("" & rs("conta_dv"))
					end if
				s_favorecido = Trim("" & rs("favorecido"))
				if s_banco <> "" then
					s_banco_nome = x_banco(s_banco)
					if s_banco_nome <> "" then s_banco_codigo_descricao = s_banco & " - " & s_banco_nome
					end if
				end if 'if Not rs.Eof

			x = x & Replace(cab_table, "tableDados", "tableDados_" & CStr(idx_bloco))
			x = x & TABLE_MARGIN_N2 & "	<tr>" & chr(13)
			
			s = Trim("" & r("indicador"))
			s_aux = x_orcamentista_e_indicador(s)
			if (s <> "") And (s_aux <> "") then s = s & " - "
			s = s & s_aux
			if s <> "" then
				x = x & TABLE_MARGIN_N2 & "		<td class='MDTE' colspan='11' align='left' valign='bottom' class='MB' style='background:azure;'><span class='N'>&nbsp;" & s & "</span></td>" & chr(13) & _
						TABLE_MARGIN_N2 & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
						TABLE_MARGIN_N2 & "	</tr>" & chr(13)
				end if

			x = x & _
				TABLE_MARGIN_N2 & "	<tr>" & chr(13) & _
				TABLE_MARGIN_N2 & "		<td class='MDTE' colspan='11' align='left' valign='bottom' class='MB' style='background:white;'>" & chr(13) & _
				TABLE_MARGIN_N2 & "			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				TABLE_MARGIN_N2 & "				<tr>" & chr(13) & _
				TABLE_MARGIN_N2 & "					<td align='right' valign='bottom' nowrap><span class='Cn'>Pagamento da Comissão via Cartão:</span></td>" & chr(13) & _
				TABLE_MARGIN_N2 & "					<td width='90%' align='left' valign='bottom' nowrap><span class='Cn'>"
			if rs("comissao_cartao_status") = 1 then
				x = x & "Sim" & " &nbsp; " & cnpj_cpf_formata(Trim("" & rs("comissao_cartao_cpf"))) & " - " & Trim("" & rs("comissao_cartao_titular"))
			else
				x = x & "Não"
				end if

			x = x & _
				"</span></td>" & chr(13) & _
				TABLE_MARGIN_N2 & "				</tr>" & chr(13) & _
				TABLE_MARGIN_N2 & "			</table>" & chr(13) & _
				TABLE_MARGIN_N2 & "		</td>" & chr(13) & _
				TABLE_MARGIN_N2 & "	</tr>" & chr(13)

			x = x & _
				TABLE_MARGIN_N2 & "	<tr>" & chr(13) & _
				TABLE_MARGIN_N2 & "		<td class='MDTE' colspan='11' align='left' valign='bottom' class='MB' style='background:whitesmoke;'>" & chr(13) & _
				TABLE_MARGIN_N2 & "			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				TABLE_MARGIN_N2 & "				<tr>" & chr(13) & _
				TABLE_MARGIN_N2 & "					<td colspan='3' align='left' valign='bottom' style='vertical-align:middle'><div valign='bottom' style='height:14px;max-height:14px;overflow:hidden;vertical-align:middle'><span class='Cn'>Banco: " & s_banco_codigo_descricao & "</span></div></td>" & chr(13) & _
				TABLE_MARGIN_N2 & "				</tr>" & chr(13) & _
				TABLE_MARGIN_N2 & "				<tr>" & chr(13) & _
				TABLE_MARGIN_N2 & "					<td class='MTD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>Agência: " & s_agencia & "</span></td>" & chr(13) & _
				TABLE_MARGIN_N2 & "					<td class='MC MD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>"

			if Trim("" & rs("tipo_conta")) <> "" then
				if Trim("" & rs("tipo_conta")) = "P" then
					x = x & "C/P: "
				elseif Trim("" & rs("tipo_conta")) = "C" then
					x = x & "C/C: "
					end if
			else
				x = x & "Conta: "
				end if

			x = x & s_conta & "</span></td>" & chr(13)

			x = x & _
				TABLE_MARGIN_N2 & "					<td class='MC' width='60%' align='left' valign='bottom'><span class='Cn'>Favorecido: " & s_favorecido & "</span></td>" & chr(13) & _
				TABLE_MARGIN_N2 & "				</tr>" & chr(13) & _
				TABLE_MARGIN_N2 & "			</table>" & chr(13) & _
				TABLE_MARGIN_N2 & "		</td>" & chr(13) & _
				TABLE_MARGIN_N2 & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				TABLE_MARGIN_N2 & "	</tr>" & chr(13)

			
			' TODO - s_new_cab = Replace(cab, "ckb_comissao_paga_tit_bloco", "ckb_bloco_indicador_" & CStr(idx_bloco))
			' TODO - s_new_cab = Replace(s_new_cab, "trata_ckb_onclick();", "trata_ckb_onclick(" & chr(34) & CStr(idx_bloco) & chr(34) & ");")
			s_new_cab = Replace(cab, "trata_ckb_onclick();", "trata_ckb_onclick(" & chr(34) & CStr(idx_bloco) & chr(34) & ");")
			s_new_cab = Replace(s_new_cab, "_NNNNN_", CStr(idx_bloco))
			x = x & s_new_cab

			n_reg = 0
			vl_sub_total_preco_venda = 0
			vl_sub_total_preco_NF = 0
			vl_sub_total_RT = 0
			vl_sub_total_RA = 0
			vl_sub_total_RA_liquido = 0
			vl_sub_total_RA_diferenca = 0

			end if 'if Trim("" & r("indicador")) <> indicador_a


		if vPedido(UBound(vPedido)).pedido <> "" then
			redim preserve vPedido(UBound(vPedido)+1)
			set vPedido(UBound(vPedido)) = new cl_REL_PEDIDOS_INDICADORES_INFO_PEDIDO
			inicializa_cl_REL_PEDIDOS_INDICADORES_INFO_PEDIDO vPedido(UBound(vPedido))
			end if

		with vPedido(Ubound(vPedido))
			.pedido = Trim("" & r("pedido"))
			.IdIndicador = r("IdIndicador")
			.IdVendedor = r("IdVendedor")
			.operacao = Trim("" & r("operacao"))
			.id_registro_operacao = Trim("" & r("id_registro"))
			end with

		' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1
		
		' CÁLCULOS

		' EVITA DIFERENÇAS DE ARREDONDAMENTO
		vl_preco_venda = converte_numero(formata_moeda(r("total_preco_venda")))
		vl_preco_NF = converte_numero(formata_moeda(r("total_preco_NF")))
		perc_RT = r("perc_RT")
		vl_RT = (perc_RT/100) * vl_preco_venda
		vl_RA = vl_preco_NF - vl_preco_venda

		if Not calcula_total_RA_liquido(r("perc_desagio_RA_liquida"), vl_RA, vl_RA_liquido) then
			Response.Write "FALHA AO CALCULAR O RA LÍQUIDO."
			Response.End
			end if
		
		vl_RA_diferenca = vl_RA - vl_RA_liquido

		' CÁLCULOS DE SUB TOTAL
		vl_sub_total_preco_venda = vl_sub_total_preco_venda + r("total_preco_venda")
		vl_total_preco_venda = vl_total_preco_venda + r("total_preco_venda")
		vl_sub_total_preco_NF = vl_sub_total_preco_NF + r("total_preco_NF")
		vl_total_preco_NF = vl_total_preco_NF + r("total_preco_NF")
		vl_sub_total_RT = vl_sub_total_RT + vl_RT
		vl_total_RT = vl_total_RT + vl_RT
		vl_sub_total_RA = vl_sub_total_RA + vl_RA
		vl_total_RA = vl_total_RA + vl_RA
		vl_sub_total_RA_liquido = vl_sub_total_RA_liquido + vl_RA_liquido
		vl_total_RA_liquido = vl_total_RA_liquido + vl_RA_liquido
		vl_sub_total_RA_diferenca = vl_sub_total_RA_diferenca + vl_RA_diferenca
		vl_total_RA_diferenca = vl_total_RA_diferenca + vl_RA_diferenca
		sub_total_comissao = vl_sub_total_RT + vl_sub_total_RA_liquido

		vl_sub_total_RT_arredondado = vl_sub_total_RT
		vl_sub_total_RA_arredondado = vl_sub_total_RA_liquido
		sub_total_com = vl_sub_total_RT_arredondado + vl_sub_total_RA_arredondado

		if lista_operacao <> "" then lista_operacao = lista_operacao & ", "
		lista_operacao = lista_operacao & r("operacao")
		if lista_completa_pedidos <> "" then lista_completa_pedidos = lista_completa_pedidos & ";"
		lista_completa_pedidos = lista_completa_pedidos & Trim("" & r("pedido"))

		with vPedido(Ubound(vPedido))
			.vl_pedido = vl_preco_venda
			.vl_comissao = vl_RT
			.vl_RA_bruto = vl_RA
			.vl_RA_liquido = vl_RA_liquido
			end with

		'> CHECK BOX
		' É USADO O CÓDIGO DA OPERAÇÃO (VENDA NORMAL, DEVOLUÇÃO, PERDA) P/ NÃO CORRER O RISCO DE HAVER CONFLITO DEVIDO A ID'S REPETIDOS ENTRE AS OPERAÇÕES
		x = x & TABLE_MARGIN_N2 & "	<tr class='VISAO_ANALIT'>" & chr(13)
		
		if (vl_preco_venda < 0) Or (vl_RT < 0) Or (vl_RA < 0) Or (vl_RA_liquido < 0) then
			s_cor = "red"
			s_cor_sinal = "red"
			s_sinal = "-"
		else
			s_cor = "black"
			s_cor_sinal = "green"
			s_sinal = "+"
			end if

		'> LOJA
		x = x & TABLE_MARGIN_N2 & "		<td class='MDTE tdLoja' align='center'><span class='Cnc' style='color:" & s_cor & ";'>" & Trim("" & r("loja")) & "</span></td>" & chr(13)

		'> Nº ORÇAMENTO
		s = Trim("" & r("orcamento"))
		if s = "" then s = "&nbsp;"
		x = x & TABLE_MARGIN_N2 & "		<td class='MTD tdOrcamento' align='left'><span class='Cn'><a style='color:" & s_cor & ";' href='javascript:fORCConsulta(" & _
				chr(34) & s & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o orçamento'>" & _
				s & "</a></span></td>" & chr(13)

		'> Nº PEDIDO
		s_nome_cliente = Trim("" & r("nome_iniciais_em_maiusculas"))
		s_nome_cliente = Left(s_nome_cliente, 15)
		x = x & TABLE_MARGIN_N2 & "		<td class='MTD tdPedido' align='left'><span class='Cn'><a style='color:" & s_cor & ";' href='javascript:fPEDConsulta(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o pedido'>" & _
				Trim("" & r("pedido")) & "<br />" & s_nome_cliente & "</a></span></td>" & chr(13)

		'> DATA
		s = formata_data(r("data"))
		x = x & TABLE_MARGIN_N2 & "		<td align='center' class='MTD tdData'><span class='Cnc' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)

		'> VALOR DO PEDIDO (PREÇO DE VENDA)
		x = x & TABLE_MARGIN_N2 & "		<td align='right' class='MTD tdVlPedido'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_preco_venda) & "</span></td>" & chr(13)
		lista_vl_pedido = lista_vl_pedido & formata_moeda(vl_preco_venda) & ";"

		'> COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
		x = x & TABLE_MARGIN_N2 & "		<td align='right' class='MTD tdVlRT'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT) & "</span></td>" & chr(13)
		lista_vl_comissao = lista_vl_comissao & formata_moeda(vl_RT) & ";"

		'> RA BRUTO
		x = x & TABLE_MARGIN_N2 & "		<td align='right' class='MTD tdVlRABruto'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA) & "</span></td>" & chr(13)
		lista_vl_RA_bruto = lista_vl_RA_bruto & formata_moeda(vl_RA) & ";"

		'> RA LÍQUIDO
		x = x & TABLE_MARGIN_N2 & "		<td align='right' class='MTD tdVlRALiq'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido) & "</span></td>" & chr(13)
		lista_vl_RA_liquido = lista_vl_RA_liquido & formata_moeda(vl_RA_liquido) & ";"

		'> RA DIFERENÇA
		x = x & TABLE_MARGIN_N2 & "		<td align='right' class='MTD tdVlRADif'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_diferenca) & "</span></td>" & chr(13)

		'> STATUS DE PAGAMENTO
		x = x & TABLE_MARGIN_N2 & "		<td class='MTD tdStPagto' align='left'><span class='Cn' style='color:" & s_cor & ";'>" & x_status_pagto(Trim("" & r("st_pagto"))) & "</span></td>" & chr(13)

		'> +/-
		x = x & TABLE_MARGIN_N2 & "		<td align='center' class='MTD tdSinal'><span class='C' style='font-family:Courier,Arial;color:" & s_cor_sinal & "'>" & s_sinal & "</span></td>" & chr(13)
		
		'> COLUNA DO ÍCONE (EXPANDE/RECOLHE)
		x = x & TABLE_MARGIN_N2 & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13)
		
		x = x & TABLE_MARGIN_N2 & "	</tr>" & chr(13)

		indicador_bloco_anterior = Trim("" & r("indicador"))

		r.MoveNext
		loop


	' MOSTRA TOTAL DO ÚLTIMO INDICADOR
	if n_reg <> 0 then
		s_checked = ""
		s_disabled = ""
' TODO - REMOVER ?		sub_total_comissao = vl_sub_total_RT + vl_sub_total_RA_liquido
		
		lista_vl_total_comissao = lista_vl_total_comissao & formata_moeda(vl_sub_total_RT) & ";"
		lista_vl_total_RA = lista_vl_total_RA & formata_moeda(vl_sub_total_RA_liquido) & ";"
		if sub_total_com >= 0 then
			s_checked = " checked"
		else
			s_disabled = " disabled"
			lista_indicador_negativo = lista_indicador_negativo & indicador_bloco_anterior & ", "
			vIndicador(UBound(vIndicador)).indicador_negativo = True
			end if

		msg_desconto = ""
		vl_total_desc_planilha = 0
		qtde_registro_desc = 0
		contador = 0

		s = "SELECT" & _
				" descricao" & _
				", valor" & _
				", ordenacao" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO" & _
			" WHERE" & _
				" (apelido = '" & indicador_bloco_anterior & "')" & _
			" ORDER BY" & _
				" ordenacao"
		if rs2.State <> 0 then rs2.Close
		rs2.Open s, cn
		if Not rs2.Eof then
			do while Not rs2.Eof
				redim preserve v_desconto_descricao(contador)
				redim preserve v_desconto_valor(contador)
				v_desconto_descricao(contador) = Trim("" & rs2("descricao"))
				v_desconto_valor(contador) = rs2("valor")
				vl_total_desc_planilha = vl_total_desc_planilha + v_desconto_valor(contador)
				qtde_registro_desc = qtde_registro_desc + 1
				contador = contador + 1
				rs2.MoveNext
				loop

			s_checked = ""
			s_disabled = " disabled"
			msg_desconto = "Planilha de Desconto:&nbsp; " & Cstr(qtde_registro_desc) & " registro(s) no valor total de&nbsp;" & formata_moeda(vl_total_desc_planilha)
			lista_indicador_com_desconto = lista_indicador_com_desconto & indicador_bloco_anterior & ", "
			lista_vl_total_desconto_planilha = lista_vl_total_desconto_planilha & formata_moeda(vl_total_desc_planilha) & ";"
			lista_qtde_reg_descontos = lista_qtde_reg_descontos & qtde_registro_desc & ";"
			with vIndicador(UBound(vIndicador))
				.indicador_com_desconto = True
				.vl_total_desc_planilha = vl_total_desc_planilha
				.qtde_reg_descontos = qtde_registro_desc
				end with
			end if 'if Not rs2.Eof

		s_cor = "black"
		if vl_sub_total_preco_venda < 0 then s_cor = "red"
		if vl_sub_total_RT < 0 then s_cor = "red"
		if vl_sub_total_RA < 0 then s_cor = "red"
		if vl_sub_total_RA_liquido < 0 then s_cor = "red"

		x = x & TABLE_MARGIN_N2 & "	<tr style='background: #FFFFDD'>" & chr(13) & _
				TABLE_MARGIN_N2 & "		<td colspan='4' class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
							"TOTAL:</span></td>" & chr(13) & _
						TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _
						TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_preco_venda_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_preco_venda) & "'>" & chr(13) & _
						TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _
						TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissao_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RT) & "'>" & chr(13)

		if (vl_sub_total_RA_arredondado >= 0) And (vl_sub_total_RT_arredondado < 0) then
			x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRT_" & CStr(idx_bloco) & "' value='" & CStr(0) & "'>" & chr(13)
		else
			if (vl_sub_total_RA_arredondado >= 0) And (vl_sub_total_RT_arredondado >= 0) then
				x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRT_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RT_arredondado) & "'>" & chr(13)
			else
				x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRT_" & CStr(idx_bloco) & "' value='" & formata_moeda(sub_total_com) & "'>" & chr(13)
				end if
			end if

		x = x &_
			TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_RA_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RA) & "'>" & chr(13)

		if (vl_sub_total_RA_arredondado < 0) And (vl_sub_total_RT_arredondado > 0) then
			x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRA_" & CStr(idx_bloco) & "' value='" & CStr(0) & "'>" & chr(13)
		else
			if (vl_sub_total_RA_arredondado >= 0) And (vl_sub_total_RT_arredondado >= 0) then
				x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRA_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RA_arredondado) & "'>" & chr(13)
			else
				x = x & TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_comissaoRA_" & CStr(idx_bloco) & "' value='" & formata_moeda(sub_total_com) & "'>" & chr(13)
				end if
			end if

		x = x & _
			TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_RA_liquido_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RA_liquido) & "'>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "			<input type='hidden' id='sub_total_RA_diferenca_" & CStr(idx_bloco) & "' value='" & formata_moeda(vl_sub_total_RA_diferenca) & "'>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
			TABLE_MARGIN_N2 & "	</tr>" & chr(13) & _
			TABLE_MARGIN_N2 & "	<tr>" & chr(13)

		lista_vl_total_comissao_arredondado = lista_vl_total_comissao_arredondado & formata_moeda(vl_sub_total_RT_arredondado) & ";"
		lista_vl_total_RA_arredondado = lista_vl_total_RA_arredondado & formata_moeda(vl_sub_total_RA_arredondado) & ";"

		with vIndicador(UBound(vIndicador))
			.vl_total_comissao = vl_sub_total_RT
			.vl_total_RA = vl_sub_total_RA_liquido
			.vl_total_comissao_arredondado = vl_sub_total_RT_arredondado
			.vl_total_RA_arredondado = vl_sub_total_RA_arredondado
			end with

		' MENSAGEM DE DESCONTO E VALORES DE RT E RA
		x = x & TABLE_MARGIN_N2 & "		<td align='left' colspan='5' nowrap>"
		if msg_desconto <> "" then x = x & "<span class='Cd'><a href='javascript:abreDesconto(" & CStr(idx_bloco) & ")' title='Exibe ou oculta os registros de descontos' style='color:red;'>" & msg_desconto & "</a></span>"
		x = x & TABLE_MARGIN_N2 & "</td>" & chr(13)

		' COMISSAO RT E RA
		if vl_sub_total_RT_arredondado >= 0 then
			s_cor = "black"
		else
			s_cor = "red"
			end if
		x = x & TABLE_MARGIN_N2 & "		<td align='right' nowrap><span class='Rd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT_arredondado) & "</span></td>" & chr(13)

		if vl_sub_total_RA_arredondado >= 0 then
			s_cor = "black"
		else
			s_cor = "red"
			end if
		x = x & TABLE_MARGIN_N2 & "		<td align='right' colspan='2' nowrap><span class='Rd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_arredondado) & "</span></td>" & chr(13)

		if sub_total_com >= 0 then
			s_cor = "black"
		else
			s_cor = "red"
			end if
		x = x & TABLE_MARGIN_N2 & "		<td align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>"

		' MEIO DE PAGAMENTO
		if sub_total_com <= 0 then
			x = x & "&nbsp;</span></td>" & chr(13)
			if lista_vl_total_pagto <> "" then lista_vl_total_pagto = lista_vl_total_pagto & ";"
			lista_vl_total_pagto = lista_vl_total_pagto & formata_moeda(0)
		else
			vIndicador(UBound(vIndicador)).meio_pagto = "CARD"
			s_lista_meio_pagto = s_lista_meio_pagto & "CARD" & ";"
			x = x & "CARTÃO:</span></td>" & chr(13) &_
					TABLE_MARGIN_N2 & "		<input type='hidden' id='forma_pag_" & CStr(idx_bloco) & "' value='CARD'>" & chr(13)
			x = x & TABLE_MARGIN_N2 & "		<input type='hidden' id='forma_pag_valor_" & CStr(idx_bloco) & "' value='" & formata_moeda(sub_total_com) & "'>" & chr(13)
			if lista_vl_total_pagto <> "" then lista_vl_total_pagto = lista_vl_total_pagto & ";"
			lista_vl_total_pagto = lista_vl_total_pagto & formata_moeda(sub_total_com)
			vIndicador(UBound(vIndicador)).vl_total_pagto = sub_total_com
			end if

		' SUB TOTAL DA COMISSAO
		x = x & TABLE_MARGIN_N2 & "		<td align='right'>"
		if sub_total_com >= 0 then
			x = x & "<span class='Cd' style='color:black;'>" & formata_moeda(sub_total_com) & "</span>"
		else
			x = x & "&nbsp;"
			end if
		x = x & _
			"</td>" & chr(13) & _
			TABLE_MARGIN_N2 & "	</tr>" & chr(13)

		if msg_desconto <> "" then
			x = x & TABLE_MARGIN_N2 & "	<tr>" & chr(13) & _
					TABLE_MARGIN_N2 & "		<td class='table_Desconto' id='table_Desconto_" & CStr(idx_bloco) & "' colspan='15'>" & chr(13) & _
					TABLE_MARGIN_N2 & "			<table border='0'>" & chr(13)
			for contador = 0 to Ubound(v_desconto_descricao)
				x = x & _
					TABLE_MARGIN_N2 & "				<tr>" & chr(13) & _
					TABLE_MARGIN_N2 & TABLE_MARGIN_N2 & "					<td width='15'>&nbsp;</td>" & chr(13) & _
					TABLE_MARGIN_N2 & "					<td align='left' width='400'><span class='Cd' style='color:red;'>" & v_desconto_descricao(contador) & "</span></td>" & chr(13) & _
					TABLE_MARGIN_N2 & "					<td align='left'><span class='Cd' style='color:red;'> R$ " & formata_moeda(v_desconto_valor(contador)) & "</span></td>" & chr(13) & _
					TABLE_MARGIN_N2 & "				</tr>" & chr(13)
				next

			x = x & TABLE_MARGIN_N2 & "			</table>" & chr(13) & _
					TABLE_MARGIN_N2 & "		</td>" & chr(13) & _
					TABLE_MARGIN_N2 & "	</tr>" & chr(13)
			end if 'if msg_desconto <> ""

' TODO - REMOVER?		x = x & "</table>" & chr(13)
		end if 'if n_reg <> 0


	'>	TOTAL GERAL
	if qtde_indicadores >= 1 then
		s_cor = "black"

		x = x & _
			TABLE_MARGIN_N2 & "	<tr>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td colspan='12' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
			TABLE_MARGIN_N2 & "	</tr>" & chr(13) & _
			TABLE_MARGIN_N2 & "	<tr>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td colspan='12' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
			TABLE_MARGIN_N2 & "	</tr>" & chr(13) & _
			TABLE_MARGIN_N2 & "	<tr style='background:honeydew'>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='MTBE' colspan='4' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
						"TOTAL GERAL:</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span id='total_VlPedido' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_preco_venda) & "</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span id='totalComissao' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span id='total_RA' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA) & "</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span id='total_RAliq' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='MTB' align='right'><span id='total_RAdif' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_diferenca) & "</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
			TABLE_MARGIN_N2 & "	</tr>" & chr(13) & _

			TABLE_MARGIN_N2 & "	<tr style='background:honeydew'>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='ME MB' colspan='4' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
						"TOTAL COMISSÃO:</span></td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='MD MB' align='left' colspan='7'><span id='spnTotalGeralResumidoComissao' class='Cd'>" + "COM: " & formata_moeda(vl_total_RT) & "&nbsp;+&nbsp;RA: " & formata_moeda(vl_total_RA_liquido) & "&nbsp;=&nbsp;" & formata_moeda(vl_total_RT+vl_total_RA_liquido) & "</td>" & chr(13) & _
			TABLE_MARGIN_N2 & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
			TABLE_MARGIN_N2 & "	</tr>" & chr(13) & _


			TABLE_MARGIN_N2 & "</table>" & chr(13)

		'Lembrando que dentro do laço a variável x é limpa a cada mudança de indicador
		x = TABLE_MARGIN_N1 & "<table cellpadding='0' cellspacing='0'>" & chr(13) & _
			TABLE_MARGIN_N1 & "	<tr>" & chr(13) & _
			TABLE_MARGIN_N1 & "		<td valign='top'><br />" & chr(13) & _
			TABLE_MARGIN_N1 & "			<input type='checkbox' name='ckb_bloco_indicador' class='CKB_COM' id='ckb_bloco_indicador_" & CStr(idx_bloco) & "' onclick='trata_ckb_onclick();calculaTotalComissao();' value='" & Cstr(vIndicador(Ubound(vIndicador)).IdVendedor) & "|" & Cstr(vIndicador(Ubound(vIndicador)).IdIndicador) & "' " & s_checked & s_disabled & " />" & chr(13) & _
			TABLE_MARGIN_N1 & "		</td>" & chr(13) & _
			TABLE_MARGIN_N1 & "		<td valign='top'>" & chr(13) & _
							x & _
							"		</td>" & chr(13) & _
			TABLE_MARGIN_N1 & "	</tr>" & chr(13) & _
			TABLE_MARGIN_N1 & "</table>" & chr(13)

		x = x & "<br />"& chr(13)

		sJsRetornaTotalBloco = _
			chr(13) & _
			"<script type='text/javascript'>" & chr(13) & _
			"function retornaQtdeTotalBlocos() {" & chr(13) & _
			"var qtdeTotalBlocos;" & chr(13) & _
			"qtdeTotalBlocos = " & CStr(idx_bloco) & ";" & chr(13) & _
			"return qtdeTotalBlocos;" & chr(13) & _
			"}" & chr(13) & _
			"</script>" & chr(13)
		end if 'if qtde_indicadores >= 1

	' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & _
			"	<tr>" & chr(13) & _
			"		<td class='MT ALERTA' colspan='12' align='center'><span class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS NO PERÍODO ESPECIFICADO&nbsp;</span></td>" & chr(13) & _
			"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"</table>" & chr(13)
		end if

	' FECHA TABELA DO ÚLTIMO INDICADOR
' TODO - EXCLUIR?	x = x & "</table>" & chr(13)

	Response.write x

	x = "<input type='hidden' name='c_lista_completa_pedidos' id='c_lista_completa_pedidos' value='" & lista_completa_pedidos & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_vl_pedido' id='c_lista_vl_pedido' value='" & lista_vl_pedido & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_vl_total_comissao' id='c_lista_vl_total_comissao' value='" & lista_vl_total_comissao & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_vl_total_comissao_arredondado' id='c_lista_vl_total_comissao_arredondado' value='" & lista_vl_total_comissao_arredondado & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_vl_total_RA' id='c_lista_vl_total_RA' value='" & lista_vl_total_RA & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_vl_total_RA_arredondado' id='c_lista_vl_total_RA_arredondado' value='" & lista_vl_total_RA_arredondado & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_meio_pagto' id='c_lista_meio_pagto' value='" & s_lista_meio_pagto & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_indicador_com_desconto' id='c_lista_indicador_com_desconto' value='" & lista_indicador_com_desconto & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_indicador_negativo' id='c_lista_indicador_negativo' value='" & lista_indicador_negativo & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_vl_total_desc_planilha' id='c_lista_vl_total_desc_planilha' value='" & lista_vl_total_desconto_planilha & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_vl_comissao' id='c_lista_vl_comissao' value='" & lista_vl_comissao & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_vl_RA_bruto' id='c_lista_vl_RA_bruto' value='" & lista_vl_RA_bruto & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_vl_RA_liquido' id='c_lista_vl_RA_liquido' value='" & lista_vl_RA_liquido & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_qtde_reg_descontos' id='c_lista_qtde_reg_descontos' value='" & lista_qtde_reg_descontos & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_operacao' id='c_lista_operacao' value='" & lista_operacao & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_vl_total_pagto' id='c_lista_vl_total_pagto' value='" & lista_vl_total_pagto & "' />" & chr(13)

	Response.write x

	if Not serializa_cl_REL_PEDIDOS_INDICADORES_INFO_INDICADOR(vIndicador, s_dados, s_erro) then
		s = "Falha ao tentar serializar os dados de indicadores"
		if s_erro <> "" then s = s & " (" & s_erro & ")"
		Session(SESSION_CLIPBOARD) = s
		Response.Redirect("mensagem.asp" & "?url_back=X&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		end if

	x = "<input type='hidden' name='c_lista_info_indicador' id='c_lista_info_indicador' value='" & s_dados & "' />" & chr(13)
	Response.write x

	if Not serializa_cl_REL_PEDIDOS_INDICADORES_INFO_PEDIDO(vPedido, s_dados, s_erro) then
		s = "Falha ao tentar serializar os dados de pedidos"
		if s_erro <> "" then s = s & " (" & s_erro & ")"
		Session(SESSION_CLIPBOARD) = s
		Response.Redirect("mensagem.asp" & "?url_back=X&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		end if

	x = "<input type='hidden' name='c_lista_info_pedido' id='c_lista_info_pedido' value='" & s_dados & "' />" & chr(13)
	Response.write x

	if r.State <> 0 then r.Close
	if rs2.State <> 0 then rs2.Close
	set r=nothing
	set rs2=nothing
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

<%=DOCTYPE_LEGADO%>

<html>

<head>
	<title>CENTRAL</title>
	</head>

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var windowScrollTopAnterior;
window.status = 'Aguarde, executando a consulta ...';

$(function() {
	$("#divPedidoConsulta").hide();

	sizeDivPedidoConsulta();

	$('#divInternoPedidoConsulta').addClass('divFixo');

	$(document).keyup(function(e) {
		if (e.keyCode == 27) fechaDivPedidoConsulta();
	});

	$("#divPedidoConsulta").click(function() {
		fechaDivPedidoConsulta();
	});

	$("#imgFechaDivPedidoConsulta").click(function() {
		fechaDivPedidoConsulta();
	});

	// EXIBE O REALCE NOS CHECKBOXES QUE SÃO EXIBIDOS INICIALMENTE ASSINALADOS
	$(".CKB_COM:enabled").each(function() {
		if ($(this).is(":checked")) {
			$(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
		}
		else {
			$(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
		}
	})

	// EVENTO P/ REALÇAR OU NÃO CONFORME SE MARCA/DESMARCA O CHECKBOX
	$(".CKB_COM:enabled").click(function() {
		if ($(this).is(":checked")) {
			$(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
		}
		else {
			$(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
		}
	})

	$(".table_Desconto").hide();
});

function abreDesconto(idx_bloco) {
    var s_seletor = "#table_Desconto_" + idx_bloco;

    $(s_seletor).toggle();
}
//Every resize of window
$(window).resize(function() {
	sizeDivPedidoConsulta();
});

function sizeDivPedidoConsulta() {
	var newHeight = $(document).height() + "px";
	$("#divPedidoConsulta").css("height", newHeight);
}

function fechaDivPedidoConsulta() {
	$(window).scrollTop(windowScrollTopAnterior);
	$("#divPedidoConsulta").fadeOut();
	$("#iframePedidoConsulta").attr("src", "");
}

function marcar_todos() {
	$(".CKB_COM:enabled")
		.prop("checked", true)
		.parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
}

function desmarcar_todos() {
	$(".CKB_COM:enabled")
		.prop("checked", false)
		.parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
}

function trata_ckb_onclick(idx_bloco) {
var s_id;
	s_id = "#ckb_bloco_indicador_" + idx_bloco;
	if ($(s_id).is(":checked")) {
		// NOP
	}
	else {
		// NOP
	}
}

function fExibeOcultaCampos(indice_bloco) {
	var s_seletor = "#tableDados_" + indice_bloco + " .VISAO_ANALIT";
	$(s_seletor).toggle();
}

function expandir_todos() {
	$(".VISAO_ANALIT").show();
}

function recolher_todos() {
	$(".VISAO_ANALIT").hide();
}

function fPEDConsulta(id_pedido, usuario) {
	windowScrollTopAnterior = $(window).scrollTop();
	sizeDivPedidoConsulta();
	$("#iframePedidoConsulta").attr("src", "PedidoConsultaView.asp?pedido_selecionado=" + id_pedido + "&pedido_selecionado_inicial=" + id_pedido + "&usuario=" + usuario);
	$("#divPedidoConsulta").fadeIn();
}

function fORCConsulta(id_orcamento, usuario) {
	windowScrollTopAnterior = $(window).scrollTop();
	sizeDivPedidoConsulta();
	$("#iframePedidoConsulta").attr("src", "OrcamentoConsultaView.asp?orcamento_selecionado=" + id_orcamento + "&orcamento_selecionado_inicial=" + id_orcamento + "&usuario=" + usuario);
	$("#divPedidoConsulta").fadeIn();
}

function fRELGravaDados(f) {
	window.status = "Aguarde ...";
	bCONFIRMA.style.visibility = "hidden";
	f.action = "RelComissaoIndicadoresPagExecConfirma.asp";

	f.submit();
}

	function calculaTotalComissao() {
		var vlTotalComissao, qtdeTotalBlocos, i, vlSubTotalComissaoAux, vlTotalPedido, vlTotalRA, vlTotalRALiq, vlTotalRADif, vlTotalRTA, vlTotalRAA;
		var textComissao;
		var vlSubTotalPrecoVendaAux, vlSubTotalRAAux, vlSubTotalRALiqAux, vlSubTotalRADifAux, vlSubTotalComissaoRTAux, vlSubTotalComissaoRAAux;
		qtdeTotalBlocos = retornaQtdeTotalBlocos();
		vlTotalComissao = 0; textComissao = "";
		vlTotalPedido = 0; vlTotalRA = 0; vlTotalRALiq = 0; vlTotalRADif = 0; vlTotalRTA = 0; vlTotalRAA = 0; vlSubTotalComissaoRTAux = 0; vlSubTotalComissaoRAAux = 0;
		vlSubTotalPrecoVendaAux = 0;
		for (i = 1; i <= qtdeTotalBlocos; i++) {
			if ($("#ckb_bloco_indicador_" + i).is(':checked')) {
				vlSubTotalComissaoAux = converte_numero($("#sub_total_comissao_" + i).val());
				vlSubTotalPrecoVendaAux = converte_numero($("#sub_total_preco_venda_" + i).val());
				vlSubTotalRAAux = converte_numero($("#sub_total_RA_" + i).val());
				vlSubTotalRALiqAux = converte_numero($("#sub_total_RA_liquido_" + i).val());
				vlSubTotalRADifAux = converte_numero($("#sub_total_RA_diferenca_" + i).val());
				vlSubTotalComissaoRTAux = converte_numero($("#sub_total_comissaoRT_" + i).val());
				vlSubTotalComissaoRAAux = converte_numero($("#sub_total_comissaoRA_" + i).val());
			}
			else {
				vlSubTotalComissaoAux = 0;
				vlSubTotalPrecoVendaAux = 0;
				vlSubTotalRAAux = 0;
				vlSubTotalRALiqAux = 0;
				vlSubTotalRADifAux = 0;
				vlSubTotalComissaoRTAux = 0;
				vlSubTotalComissaoRAAux = 0;
			}

			// Total
			vlTotalComissao += vlSubTotalComissaoAux;
			vlTotalPedido += vlSubTotalPrecoVendaAux;
			vlTotalRA += vlSubTotalRAAux;
			vlTotalRALiq += vlSubTotalRALiqAux;
			vlTotalRADif += vlSubTotalRADifAux;
			vlTotalRAA += vlSubTotalComissaoRAAux;
			vlTotalRTA += vlSubTotalComissaoRTAux;
		}

		$("#totalComissao").text(formata_moeda(vlTotalComissao));
		$("#total_VlPedido").text(formata_moeda(vlTotalPedido));
		$("#total_RA").text(formata_moeda(vlTotalRA));
		$("#total_RAliq").text(formata_moeda(vlTotalRALiq));
		$("#total_RAdif").text(formata_moeda(vlTotalRADif));

		if (vlTotalRTA != 0) {
			textComissao = textComissao + "COM: " + formata_moeda(String(vlTotalRTA))
			if (vlTotalRAA != 0) { textComissao = textComissao + "&nbsp;+&nbsp;" }
		}
		if (vlTotalRAA != 0) {
			textComissao = textComissao + "RA: " + formata_moeda(String(vlTotalRAA))
			if (vlTotalRTA != 0) { textComissao = textComissao + "&nbsp;=&nbsp;" + formata_moeda((vlTotalRTA + vlTotalRAA)) }
		}

		$("#spnTotalGeralResumidoComissao").html(textComissao);
	}
</script>

<script type="text/javascript">
	$(function () {
		calculaTotalComissao();
	});
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
.tdCkb
{
	width: 20px;
}
.tdLoja{
	width: 28px;
	}
.tdOrcamento{
	width: 58px;
	}
.tdPedido{
	width: 83px;
	}
.tdData{
	width: 60px;
	}
.tdVlPedido{
	width: 70px;
	}
.tdVlRT{
	width: 60px;
	}
.tdVlRABruto{
	width: 60px;
	}
.tdVlRALiq{
	width: 60px;
	}
.tdVlRADif{
	width: 60px;
	}
.tdStPagto{
	width: 60px;
	}
.tdSinal{
	width: 18px;
	}
.BTN_LNK
{
	min-width:140px;
}
.CKB_HIGHLIGHT
{
	background-color:#90EE90;
}
#divPedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsulta
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsulta.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivPedidoConsulta
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframePedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
.BkgWhite
{
	background-color:#ffffff;
}
.VISAO_ANALIT
{
	<%if blnVisaoSintetica then Response.Write "display:none;"%>
}

</style>

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br />
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br /><br />
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>

<% else
     %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<input type="hidden" name="mes" id="mes" value="<%=c_dt_entregue_mes%>">
<input type="hidden" name="ano" id="ano" value="<%=c_dt_entregue_ano%>">
<input type="hidden" name="c_usuario_sessao" id="c_usuario_sessao" value="<%=usuario%>" />
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (Processamento)</span>
	<br /><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<%
	s_filtro = "<table width='709' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	PERÍODO: MÊS DE COMPETÊNCIA
	s = ""
	if (c_dt_entregue_mes <> "") Or (c_dt_entregue_ano <> "") then
	'	DEVIDO AO WORD WRAP: SÓ FAZ WORD WRAP QUANDO ENCONTRA CHR(32), OU SEJA, MANTÉM AGRUPADO TEXTO COM &nbsp;
		if s <> "" then s = s & ",&nbsp; "
		s_aux = c_dt_entregue_mes
		if s_aux = "" then 
            s_aux = "N.I."
        else
            if c_dt_entregue_ano = "" then 
                s_aux = "N.I."
            else
	    	    s_aux = " " & s_aux & "/"
		        s_aux = replace(s_aux, " ", "&nbsp;")
		        s = s & s_aux
		        s_aux = c_dt_entregue_ano
		        s_aux = replace(s_aux, " ", "&nbsp;")
            end if
        end if
		s = s & s_aux  
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Mês de competência:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
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
					"		<td align='right' valign='top' nowrap><span class='N'>Comissão:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	STATUS DE PAGAMENTO
	s = ""
	s_aux = Lcase(x_status_pagto(ckb_st_pagto_pago))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_status_pagto(ckb_st_pagto_nao_pago))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_pagto(ckb_st_pagto_pago_parcial))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Status de Pagamento:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
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
					"		<td align='right' valign='top' nowrap><span class='N'>Vendedor(es):&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	EMISSÃO
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br />
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td class="Rc" align="left">&nbsp;</td>
</tr>
</table>


    <%if aviso="" then%>
<table class="notPrint" width="709" cellpadding="0" cellspacing="0" style="margin-top:5px;">
<tr>
	<td align="right">
		<button type="button" name="bExpandirTodos" id="bExpandirTodos" class="Button BTN_LNK" onclick="expandir_todos();" title="expandir todas as linhas de dados" style="margin-left:6px;margin-bottom:2px">Expandir Tudo</button>
		&nbsp;
		<button type="button" name="bRecolherTodos" id="bRecolherTodos" class="Button BTN_LNK" onclick="recolher_todos();" title="recolher todas as linhas de dados" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Recolher Tudo</button>
		&nbsp;
		<button type="button" name="bMarcarTodos" id="bMarcarTodos" class="Button BTN_LNK" onclick="marcar_todos();calculaTotalComissao();" title="assinala todos os pedidos para gravar o status da comissão como paga" style="margin-left:6px;margin-bottom:2px">Marcar todos</button>
		&nbsp;
		<button type="button" name="bDesmarcarTodos" id="bDesmarcarTodos" class="Button BTN_LNK" onclick="desmarcar_todos();calculaTotalComissao();" title="desmarca todos os pedidos para gravar o status da comissão como não-paga" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Desmarcar todos</button>
	</td>
</tr>
</table>
    <%end if%>

<br />
<table class="notPrint" width="709" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="left">&nbsp;</td>
	<td align="right">
        <%if aviso = "" then %>
		<div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	    <%end if %>
    </td>
</tr>
</table>


</form>

</center>

<div id="divPedidoConsulta"><center><div id="divInternoPedidoConsulta"><img id="imgFechaDivPedidoConsulta" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframePedidoConsulta"></iframe></div></center></div>

</body>

<%=sJsRetornaTotalBloco%>

<% end if %>

</html>

<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
