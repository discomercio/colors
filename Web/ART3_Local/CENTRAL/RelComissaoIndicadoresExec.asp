<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
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
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro,rs2
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
    If Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	dim s, s_aux, s_filtro
	dim ckb_st_entrega_entregue, c_dt_entregue_inicio, c_dt_entregue_termino
	dim ckb_comissao_paga_sim, ckb_comissao_paga_nao
	dim ckb_st_pagto_pago, ckb_st_pagto_nao_pago, ckb_st_pagto_pago_parcial
	dim rb_pagto_comissao_via_cartao
	dim c_vendedor, c_indicador
	dim c_loja, lista_loja, s_filtro_loja, v_loja, v, i
	dim rb_visao, blnVisaoSintetica
    

	alerta = ""

	ckb_st_entrega_entregue = Trim(Request.Form("ckb_st_entrega_entregue"))
	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))

	c_vendedor = Trim(Request.Form("c_vendedor"))
	c_indicador = Trim(Request.Form("c_indicador"))

	ckb_comissao_paga_sim = Trim(Request.Form("ckb_comissao_paga_sim"))
	ckb_comissao_paga_nao = Trim(Request.Form("ckb_comissao_paga_nao"))

	ckb_st_pagto_pago = Trim(Request.Form("ckb_st_pagto_pago"))
	ckb_st_pagto_nao_pago = Trim(Request.Form("ckb_st_pagto_nao_pago"))
	ckb_st_pagto_pago_parcial = Trim(Request.Form("ckb_st_pagto_pago_parcial"))
	
	rb_pagto_comissao_via_cartao = Trim(Request.Form("rb_pagto_comissao_via_cartao"))
	rb_visao = Trim(Request.Form("rb_visao"))
	
	blnVisaoSintetica = False
	if rb_visao = "SINTETICA" then blnVisaoSintetica = True
	
	c_loja = Trim(Request.Form("c_loja"))
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
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
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
	
dim o
dim strMsg
dim resultadoCalculo,resultadoDigito,QtdeCedulas,TotalCedula,limitador(5)
dim dadosCalculo
set o = createobject( "ComPlusCalcCedulas.ComPlusCalcCedulas" )
dim v_cedulas,aux(5),y,totalArredondado, cont
dim cedulas()
dim qtdeCedula(), limitador_fixo(5)







' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const VENDA_NORMAL = "VENDA_NORMAL"
const DEVOLUCAO = "DEVOLUCAO"
const PERDA = "PERDA"
dim r
dim s, s_aux, s_sql, x, cab_table, cab, indicador_a, n_reg, n_reg_total, qtde_indicadores
dim vl_preco_venda, vl_sub_total_preco_venda, vl_total_preco_venda, cont2
dim vl_preco_NF, vl_sub_total_preco_NF, vl_total_preco_NF
dim vl_RT, vl_sub_total_RT, vl_total_RT
dim vl_RA, vl_sub_total_RA, vl_total_RA
dim vl_RA_liquido, vl_sub_total_RA_liquido, vl_total_RA_liquido
dim vl_RA_diferenca, vl_sub_total_RA_diferenca, vl_total_RA_diferenca
dim perc_RT
dim s_where, s_where_venda, s_where_devolucao, s_where_perdas, s_where_loja
dim s_where_comissao_paga, s_where_comissao_descontada, s_where_st_pagto, s_where_dt_st_pagto
dim s_cor, s_sinal, s_cor_sinal
dim s_banco, s_banco_nome, s_agencia, s_conta, s_favorecido, s_favorecido_cnpj_cpf
dim s_nome_cliente, s_desempenho_nota
dim s_id, s_checked, s_class, s_class_td, idx_bloco, s_new_cab
dim s_lista_completa_venda_normal, s_lista_completa_devolucao, s_lista_completa_perda
dim s_lista_ja_marcado_venda_normal, s_lista_ja_marcado_devolucao, s_lista_ja_marcado_perda,sub_total_com
dim sub_total_comissao, banco, limita_cedula,qtdeChq ,textDin,textChq,sub_total_com_RA,sub_total_com_RT,vl_sub_total_RT_arredondado,vl_sub_total_RA_arredondado
dim v_desconto_descricao(),v_desconto_valor(),contador,valor_desconto,qtde_registro_desc,msg_desconto,ind_anterior
    limita_cedula = Request.Form("limita_cedula")
    textDin = ""
    textChq = ""
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
'	A) VENDAS
	s_where_comissao_paga = ""
	s = ""
	s_aux = ckb_comissao_paga_sim
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.comissao_paga = " & COD_COMISSAO_PAGA & ")"
		end if

	s_aux = ckb_comissao_paga_nao
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.comissao_paga = " & COD_COMISSAO_NAO_PAGA & ")"
		end if

	if s <> "" then 
		if s_where_comissao_paga <> "" then s_where_comissao_paga = s_where_comissao_paga & " AND"
		s_where_comissao_paga = s_where_comissao_paga & " (" & s & ")"
		end if
		
'	B) PERDAS/DEVOLUÇÕES
	s_where_comissao_descontada = ""
	s = ""
	s_aux = ckb_comissao_paga_sim
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (comissao_descontada = " & COD_COMISSAO_DESCONTADA & ")"
		end if
	
	s_aux = ckb_comissao_paga_nao
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (comissao_descontada = " & COD_COMISSAO_NAO_DESCONTADA & ")"
		end if

	if s <> "" then 
		if s_where_comissao_descontada <> "" then s_where_comissao_descontada = s_where_comissao_descontada & " AND"
		s_where_comissao_descontada = s_where_comissao_descontada & " (" & s & ")"
		end if

'	CRITÉRIO: STATUS DE PAGAMENTO
	s_where_dt_st_pagto = ""
	if IsDate(c_dt_entregue_termino) then
		s_where_dt_st_pagto = s_where_dt_st_pagto & " AND (t_PEDIDO__BASE.dt_st_pagto < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		end if

	s_where_st_pagto = ""
	s = ""
	s_aux = ckb_st_pagto_pago
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO__BASE.st_pagto = '" & s_aux & "')" & s_where_dt_st_pagto & ")"
		end if

	s_aux = ckb_st_pagto_nao_pago
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO__BASE.st_pagto = '" & s_aux & "')" & s_where_dt_st_pagto & ")"
		end if
	
	s_aux = ckb_st_pagto_pago_parcial
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO__BASE.st_pagto = '" & s_aux & "')" & s_where_dt_st_pagto & ")"
		end if

	if s <> "" then 
		if s_where_st_pagto <> "" then s_where_st_pagto = s_where_st_pagto & " AND"
		s_where_st_pagto = s_where_st_pagto & " (" & s & ")"
		end if
	
'	CRITÉRIO: PAGAMENTO DA COMISSÃO VIA CARTÃO
	if rb_pagto_comissao_via_cartao <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.comissao_cartao_status = " & rb_pagto_comissao_via_cartao & ")"
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
		
	if s_where_comissao_paga <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (" & s_where_comissao_paga & ")"
		end if

	if s_where_st_pagto <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (" & s_where_st_pagto & ")"
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

	if s_where_comissao_descontada <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (" & s_where_comissao_descontada & ")"
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
		
	if s_where_comissao_descontada <> "" then
		if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		s_where_perdas = s_where_perdas & " (" & s_where_comissao_descontada & ")"
		end if
		
		
'	VENDAS NORMAIS
	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then s = " AND" & s
	s_sql = "SELECT" & _
			" '" & VENDA_NORMAL & "' AS operacao," & _
			" t_PEDIDO.pedido AS id_registro," & _
			" t_PEDIDO.comissao_paga AS status_comissao," & _
			" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota," & _
            " t_ORCAMENTISTA_E_INDICADOR.banco," & _
			" t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
			" t_PEDIDO.endereco_nome AS nome," & _
			" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
			" t_CLIENTE.nome," & _
			" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO.entregue_data AS data," & _
			" t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			" t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			" t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA," & _
			" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
			" Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.preco_venda) AS total_preco_venda," & _
			" Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.preco_NF) AS total_preco_NF" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			" WHERE (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
			s & _
			" GROUP BY t_PEDIDO.pedido, t_PEDIDO.comissao_paga, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome," & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO.entregue_data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"

'	ITENS DEVOLVIDOS
	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then s = " WHERE " & s
	s_sql = s_sql & " UNION ALL " & _
			"SELECT" & _
			" '" & DEVOLUCAO & "' AS operacao," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.id AS id_registro," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.comissao_descontada AS status_comissao," & _
			" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota," & _
            " t_ORCAMENTISTA_E_INDICADOR.banco," & _
			" t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome AS nome," & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data," & _
			" t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			" t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			" t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA," & _
			" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
			" Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS total_preco_venda," & _
			" Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_NF) AS total_preco_NF" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido)" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			s & _
			" GROUP BY t_PEDIDO_ITEM_DEVOLVIDO.id, t_PEDIDO_ITEM_DEVOLVIDO.comissao_descontada, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome," & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"

'	PERDAS
	s = s_where
	if (s <> "") And (s_where_perdas <> "") then s = s & " AND"
	s = s & s_where_perdas
	if s <> "" then s = " WHERE " & s
	s_sql = s_sql & " UNION ALL " & _
			"SELECT" & _
			" '" & PERDA & "' AS operacao," & _
			" t_PEDIDO_PERDA.id AS id_registro," & _
			" t_PEDIDO_PERDA.comissao_descontada AS status_comissao," & _
			" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota," & _
            " t_ORCAMENTISTA_E_INDICADOR.banco," & _
			" t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
			" t_PEDIDO.endereco_nome AS nome," & _
			" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
			" t_CLIENTE.nome," & _
			" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO_PERDA.data AS data," & _
			" t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			" t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			" t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA," & _
			" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
			" Sum(-t_PEDIDO_PERDA.valor) AS total_preco_venda," & _
			" Sum(-t_PEDIDO_PERDA.valor) AS total_preco_NF" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_PEDIDO_PERDA ON (t_PEDIDO.pedido=t_PEDIDO_PERDA.pedido)" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			s & _
			" GROUP BY t_PEDIDO_PERDA.id, t_PEDIDO_PERDA.comissao_descontada, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome," & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if
		
	s_sql = s_sql & _
			" t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_PERDA.data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"
	
	s_sql = "SELECT " & _
				"*" & _
			" FROM (" & _
				s_sql & _
				") t" & _
			" ORDER BY t.desempenho_nota, t.indicador, t.numero_loja, t.data, t.pedido, t.total_preco_venda DESC"

  ' CABEÇALHO
	cab_table = "<table cellspacing='0' id='tableDados'>" & chr(13)
	cab = "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='MDTE tdCkb' align='center' valign='bottom' nowrap><input type='checkbox' name='ckb_comissao_paga_tit_bloco' id='ckb_comissao_paga_tit_bloco' class='CKB_COM VISAO_ANALIT' onclick='trata_ckb_onclick();' /></td>" & chr(13) & _
		  "		<td class='MTD tdLoja' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>Loja</span></td>" & chr(13) & _
		  "		<td class='MTD tdOrcamento' align='left' valign='bottom' nowrap><span class='R VISAO_ANALIT'>Nº Orçam</span></td>" & chr(13) & _
		  "		<td class='MTD tdPedido' align='left' valign='bottom' nowrap><span class='R VISAO_ANALIT'>Nº Pedido</span></td>" & chr(13) & _
		  "		<td class='MTD tdData' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>Data</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlPedido' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL Pedido</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRT' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>COM (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRABruto' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Bruto (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRALiq' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Líq (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRADif' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Dif (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdStPagto' align='left' valign='bottom'><span class='R VISAO_ANALIT' style='font-weight:bold;'>St Pagto</span></td>" & chr(13) & _
		  "		<td class='MTD tdSinal' align='center' valign='bottom'><span class='Rc VISAO_ANALIT' style='font-weight:bold;'>+/-</span></td>" & chr(13) & _
		  "		<td valign='bottom' class='notPrint BkgWhite' align='left'>&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & "_NNNNN_" & chr(34) & ");' title='exibe ou oculta os dados'><img src='../botao/view_bottom.png' border='0'></a></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
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
	s_lista_completa_venda_normal = ""
	s_lista_completa_devolucao = ""
	s_lista_completa_perda = ""
	s_lista_ja_marcado_venda_normal = ""
	s_lista_ja_marcado_devolucao = ""
	s_lista_ja_marcado_perda = ""
    sub_total_comissao = 0
    sub_total_com = 0
	indicador_a = "XXXXXXXXXXXX"
	set r = cn.execute(s_sql)

dim totalComDin,totalComChqOutros,totalComchqBradesco
    totalComDin= 0
    totalComChqOutros=0
    totalComchqBradesco=0
     
if limita_cedula = "" then
limitador(0)= "10"
limitador(1)= "10"
limitador(2)= "10"
limitador(3)= "10"
limitador(4)= "10"
limitador(5)= "10"
else 
for cont = 0 to Ubound(limitador)
        limitador(cont) = Request.Form("txtcedula")(cont+1)
        limitador_fixo(cont) = Request.Form("txtcedula")(cont+1)

        if (limitador(cont) = "") then limitador(cont) = 0
        if (limitador_fixo(cont) = "") then limitador_fixo(cont) = 0
next
end if

	do while Not r.Eof
	'	MUDOU DE INDICADOR?
		if Trim("" & r("indicador"))<>indicador_a then
			indicador_a = Trim("" & r("indicador"))
			idx_bloco = idx_bloco + 1
			qtde_indicadores = qtde_indicadores + 1
            sub_total_comissao = vl_sub_total_RT + vl_sub_total_RA_liquido


    '---------------------------------------   calculo de cédulas
    dim j, z,totalcomissao
    z = 0
    j = 0

        if sub_total_com <= 0 Or sub_total_com > 300 Or banco = "237" then
                    if (sub_total_com  >0 Or sub_total_com > 300) And banco <> "237" then                             
                            totalComChqOutros = totalComChqOutros + Cstr(sub_total_com)   
                            qtdeChq = qtdeChq +1
                            elseif sub_total_com <0 then 
                
                            else
                                totalComchqBradesco = totalComchqBradesco + Cstr(sub_total_com)
                                                            
                            end if
                    else
                    dadosCalculo = Clng(sub_total_com)  
                    totalComDin =  totalComDin + dadosCalculo
                    dadosCalculo = o.CalculaCedulas(dadosCalculo,"2#"& limitador(5) &"|5#"&limitador(4)&"|10#"&limitador(3)&"|20#"&limitador(2)&"|50#"&limitador(1)&"|100#"&limitador(0)&"",resultadoCalculo)                    
                    dadosCalculo = resultadocalculo

                    v_cedulas = Split(dadosCalculo,"|")   
                         for cont=0 to Ubound(v_cedulas)
                               if cont mod 2 = 0 then
                                    redim  preserve cedulas(j)
                                    ' cedulas(j)=  cint(v_cedulas(cont))
                                     j = j +1
                                else
                                     redim preserve qtdeCedula(z)
                                     qtdeCedula(z) = cint(v_cedulas(cont))
                                     z = z + 1                   
                                end if
                         next 
          end if
   
		  ' FECHA TABELA DO INDICADOR ANTERIOR
			if n_reg_total > 0 then 
				s_cor="black"
				if vl_sub_total_preco_venda < 0 then s_cor="red"
				if vl_sub_total_RT < 0 then s_cor="red"
				if vl_sub_total_RA < 0 then s_cor="red"
				if vl_sub_total_RA_liquido < 0 then s_cor="red"

                if rs2.State <> 0 then rs2.Close 
                rs2.Open "SELECT COUNT(*) qtde_Desconto, descricao,valor,ordenacao  FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (apelido = '" & ind_anterior & "') GROUP BY  descricao,valor,ordenacao ORDER BY ordenacao", cn
                msg_desconto = ""
                if Not rs2.Eof then
                    
                        valor_desconto = 0
                        contador = 0
                        qtde_registro_desc = 0
                        Erase v_desconto_descricao
                        Erase v_desconto_valor
                        do while Not rs2.EoF         
                            redim preserve v_desconto_descricao(contador) 
                            redim preserve  v_desconto_valor(contador)                 
                            v_desconto_descricao(contador) = rs2("descricao")
                            v_desconto_valor(contador) = rs2("valor")
                            valor_desconto = valor_desconto + v_desconto_valor(contador)
                            qtde_registro_desc = qtde_registro_desc + 1
                            contador = contador + 1
                            rs2.MoveNext
		                loop
                        msg_desconto =  "Planilha de Desconto:&nbsp; " & Cstr(qtde_registro_desc) & " registro(s) no valor total de&nbsp;" & formata_moeda(valor_desconto)                       
                    else
                        msg_desconto= ""
                    
                end if
				x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
						"		<td class='MTBE' colspan='5' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
						"TOTAL:</span></td>" & chr(13) & _
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</span></td>" & chr(13) & _
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _
						"		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
						"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
						"	</tr>" & chr(13) & _ 
                        "   <tr>" & chr(13) 
                   if sub_total_com <= 0 Or o.DigitoFinal(sub_total_com) > 300 Or banco = "237" then
                       x = x & "<td align='left' colspan='10' nowrap><span class='Cd' style='color:" & s_cor & ";'></span>" 
                                    if msg_desconto <> "" then x = x & "<span class='Cd'><a href='javascript:abreDesconto(" & idx_bloco -1 & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"
                   else if sub_total_com >= 0 then
                        x = x & "       <td align='left' colspan='10' nowrap>"
                        if msg_desconto <> "" then 
                            x = x & "<span class='Cd' ><a href='javascript:abreDesconto(" & idx_bloco -1  & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"
                        else   
                            x = x & "<span class='Rd' style='color: black;'>"          
                            for cont = 0 to UBound(qtdeCedula)
                                if (cont = 0 And qtdeCedula(cont) <> 0) then
                                                aux(cont) = aux(cont) + qtdeCedula(cont)                                               
                                elseif (cont = 1 And qtdeCedula(cont) <> 0) then                         
                                                aux(cont) = aux(cont) + qtdeCedula(cont)                              
                                elseif (cont = 2 And qtdeCedula(cont) <> 0) then                 
                                                aux(cont) = aux(cont) + qtdeCedula(cont)           
                                elseif (cont = 3 And qtdeCedula(cont) <> 0) then                  
                                                aux(cont) = aux(cont) + qtdeCedula(cont)
                                elseif (cont = 4 And qtdeCedula(cont) <> 0) then
                                                aux(cont) = aux(cont) + qtdeCedula(cont)                      
                                elseif (cont = 5 And qtdeCedula(cont) <> 0) then                      
                                                aux(cont) = aux(cont) + qtdeCedula(cont) 
                                end if
                        
                            next
                      x = x & "</span></td>" & chr(13) 
                      end if
                   end if
                   end if
                   if sub_total_com_RT + sub_total_com_RA >= 0 then 
                    s_cor = "black"
                   else 
                    s_cor = "red"
                   end if
                     x = x & "       <td align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>"
                   if sub_total_com <= 0 then
                   x = x & "&nbsp;</span></td>" & chr(13)
                   elseif o.DigitoFinal(sub_total_com) > 300 And banco <> "237" then
                   x = x & "CHQ:</span></td>" & chr(13)
                   elseif banco = "237" then
                   x = x & "DEP:</span></td>" & chr(13)  
                   else 
                   x = x & "DIN:</span></td>" & chr(13)
                   end if
                    
                    if sub_total_com >=0  then                      
                       x = x & "       <td align='right' ><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(sub_total_com) & "</span></td>" & chr(13)      
                    end if
                  x = x &"	</tr>" & chr(13) 
                  if msg_desconto <> "" then
                            x = x &"   <tr>" & chr(13) & _
                                   "          <td  class='VISAO_ANALIT' id='table_Desconto_"& idx_bloco -1   &"' colspan='15' >" & chr(13)& _
                                   "          <table colspan='2' align='left' >"& chr(13)
                                   for contador = 0 to Ubound(v_desconto_descricao)                                 
                                       x = x & "   <tr>" & chr(13)& _
                                                "       <td width='15'>&nbsp;</td>" & chr(13)& _
                                                "       <td  align='left' width='400' ><span class='Cd'style='color: red;' >"& v_desconto_descricao(contador)& "</span></td>"& _
                                                "       <td align='left' ><span class='Cd'style='color: red;' > R$ "& formata_moeda(v_desconto_valor(contador))& "</span></td>"& _
                                                "   </tr>"                                     
                                   next          
                                   x = x & "   </table>"& chr(13)& _
                                           "   </td>"& chr(13)& _
                                           "</tr>"
                   end if         
						    x = x &"</table>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			n_reg = 0
			vl_sub_total_preco_venda = 0
			vl_sub_total_preco_NF = 0
			vl_sub_total_RT = 0
			vl_sub_total_RA = 0
			vl_sub_total_RA_liquido = 0
			vl_sub_total_RA_diferenca = 0

			s_sql = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido = '" & Trim("" & r("indicador")) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Not rs.Eof then
				s_banco = Trim("" & rs("banco"))
				s_agencia = Trim("" & rs("agencia"))
				s_conta = Trim("" & rs("conta"))
				s_favorecido = Trim("" & rs("favorecido"))
				s_favorecido_cnpj_cpf = Trim("" & rs("favorecido_cnpj_cpf"))
				if s_favorecido_cnpj_cpf <> "" then s_favorecido_cnpj_cpf = cnpj_cpf_formata(s_favorecido_cnpj_cpf)
				s_banco_nome = x_banco(s_banco)
				if (s_banco <> "") And (s_banco_nome <> "") then s_banco = s_banco & " - " & s_banco_nome
			else
				s_banco = ""
				s_banco_nome = ""
				s_agencia = ""
				s_conta = ""
				s_favorecido = ""
				s_favorecido_cnpj_cpf = ""
				end if
			
			if n_reg_total > 0 then x = x & "<BR>"
			s_desempenho_nota = Trim("" & r("desempenho_nota"))
			if s_desempenho_nota = "" then 
				s_desempenho_nota = "&nbsp;"
			else
				s_desempenho_nota = "(" & s_desempenho_nota & ") "
				end if
			s = Trim("" & r("indicador"))
			s_aux = x_orcamentista_e_indicador(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & Replace(cab_table, "tableDados", "tableDados_" & idx_bloco)
			if s <> "" then x = x & "	<tr>" & chr(13) & _
									"		<td class='MDTE' colspan='12' align='left' valign='bottom' class='MB' style='background:azure;'><span class='N'>&nbsp;" & s_desempenho_nota & s & "</span></td>" & chr(13) & _
									"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"	<tr>" & chr(13) & _
									"		<td class='MDTE' colspan='12' align='left' valign='bottom' class='MB' style='background:white;'>" & chr(13) & _
									"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
									"				<tr>" & chr(13) & _
									"					<td align='right' valign='bottom' nowrap><span class='Cn'>Pagamento da Comissão via Cartão:</span></td>" & chr(13) & _
									"					<td width='90%' align='left' valign='bottom' nowrap><span class='Cn'>"
				if rs("comissao_cartao_status") = 1 then
					x = x & "Sim" & " &nbsp; " & cnpj_cpf_formata(Trim("" & rs("comissao_cartao_cpf"))) & " - " & Trim("" & rs("comissao_cartao_titular"))
				else
					x = x & "Não"
					end if

				x = x & _
									"</span></td>" & chr(13) & _
									"				</tr>" & chr(13) & _
									"				<tr>" & chr(13) & _
									"					<td class='MC' align='right' valign='bottom' nowrap><span class='Cn'>Pagamento da Comissão via NFSe:</span></td>" & chr(13) & _
									"					<td class='MC' width='90%' align='left' valign='bottom' nowrap><span class='Cn'>" & chr(13)

				if Trim("" & rs("comissao_NFSe_cnpj")) <> "" then
					x = x & cnpj_cpf_formata(Trim("" & rs("comissao_NFSe_cnpj"))) & " - " & Trim("" & rs("comissao_NFSe_razao_social"))
				else
					x = x & "N.I."
					end if

				x = x & _
									"</span></td>" & chr(13) & _
									"				</tr>" & chr(13) & _
									"			</table>" & chr(13) & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"	<tr>" & chr(13) & _
									"		<td class='MDTE' colspan='12' align='left' valign='bottom' class='MB' style='background:whitesmoke;'>" & chr(13) & _
									"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
									"				<tr>" & chr(13) & _
									"					<td colspan='3' align='left' valign='bottom'><span class='Cn'>Banco: " & s_banco & "</span></td>" & chr(13) & _
									"				</tr>" & chr(13) & _
									"				<tr>" & chr(13) & _
									"					<td class='MTD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>Agência: " & rs("agencia")
			            if Trim("" & rs("agencia_dv")) <> "" then
                            x = x & "-" & rs("agencia_dv") & chr(13)
                        end if
    
                        x = x & "</span></td>" & chr(13) & _
									"					<td class='MC MD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>"

                        if Trim("" & rs("tipo_conta")) <> "" then
                            if rs("tipo_conta") = "P" then
                                x = x & "C/P: "
                            elseif rs("tipo_conta") = "C" then
                                x = x & "C/C: "
                            end if
                        else
                            x = x & "Conta: "
                        end if

                        if Trim("" & rs("conta_operacao")) <> "" then
                            x = x & rs("conta_operacao") & "-"
                        end if               
    
                        x = x & rs("conta")
    
                        if Trim("" & rs("conta_dv")) <> "" then
                            x = x & "-" & rs("conta_dv") & chr(13)
                        end if
						x = x &		"					<td class='MC' width='60%' align='left' valign='bottom'><span class='Cn'>Favorecido: " & s_favorecido & "</span></td>" & chr(13) & _
									"				</tr>" & chr(13)

						if Len(retorna_so_digitos(s_favorecido_cnpj_cpf)) = 11 then
							s_aux = "CPF"
						else
							s_aux = "CNPJ"
							end if

						x = x & _
									"				<tr>" & chr(13) & _
									"					<td colspan='3' class='MC' align='left' valign='bottom'><span class='Cn'>" & s_aux & ": " & s_favorecido_cnpj_cpf & "</span></td>" & chr(13) & _
									"				</tr>" & chr(13)

						x = x & _
									"			</table>" & chr(13) & _
									"		</td>" & chr(13) & _
									"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
									"	</tr>" & chr(13)
			s_new_cab = Replace(cab, "ckb_comissao_paga_tit_bloco", "ckb_comissao_paga_tit_bloco_" & idx_bloco)
			s_new_cab = Replace(s_new_cab, "trata_ckb_onclick();", "trata_ckb_onclick(" & chr(34) & idx_bloco & chr(34) & ");")
			s_new_cab = Replace(s_new_cab, "_NNNNN_", CStr(idx_bloco))
			x = x & s_new_cab
			end if
		

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<tr nowrap class='VISAO_ANALIT'>"  & chr(13)

	'	EVITA DIFERENÇAS DE ARREDONDAMENTO
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
		
		if (vl_preco_venda < 0) Or (vl_RT < 0) Or (vl_RA < 0) Or (vl_RA_liquido < 0) then
			s_cor = "red"
			s_cor_sinal = "red"
			s_sinal = "-"
		else
			s_cor = "black"
			s_cor_sinal = "green"
			s_sinal = "+"
			end if

	 '> CHECK BOX
	 '	É USADO O CÓDIGO DA OPERAÇÃO (VENDA NORMAL, DEVOLUÇÃO, PERDA) P/ NÃO CORRER O RISCO DE HAVER CONFLITO DEVIDO A ID'S REPETIDOS ENTRE AS OPERAÇÕES
		s_id = "ckb_comissao_paga_" & Trim("" & r("operacao")) & "_" & Trim("" & r("id_registro"))
		s_checked = ""
		s_class = " CKB_COM_BL_" & idx_bloco
		s_class_td = ""
		if Trim("" & r("operacao")) = VENDA_NORMAL then
			if s_lista_completa_venda_normal <> "" then s_lista_completa_venda_normal = s_lista_completa_venda_normal & ";"
			s_lista_completa_venda_normal = s_lista_completa_venda_normal & Trim("" & r("id_registro"))
			s_class = s_class & " CKB_COM_VDNORM"
			if CLng(r("status_comissao")) = CLng(COD_COMISSAO_PAGA) then
				s_checked = " checked"
				s_class_td = s_class_td & " CKB_HIGHLIGHT"
				if s_lista_ja_marcado_venda_normal <> "" then s_lista_ja_marcado_venda_normal = s_lista_ja_marcado_venda_normal & ";"
				s_lista_ja_marcado_venda_normal = s_lista_ja_marcado_venda_normal & Trim("" & r("id_registro"))
				end if
		elseif Trim("" & r("operacao")) = DEVOLUCAO then
			if s_lista_completa_devolucao <> "" then s_lista_completa_devolucao = s_lista_completa_devolucao & ";"
			s_lista_completa_devolucao = s_lista_completa_devolucao & Trim("" & r("id_registro"))
			s_class = s_class & " CKB_COM_DEV"
			if CLng(r("status_comissao")) = CLng(COD_COMISSAO_DESCONTADA) then
				s_checked = " checked"
				s_class_td = s_class_td & " CKB_HIGHLIGHT"
				if s_lista_ja_marcado_devolucao <> "" then s_lista_ja_marcado_devolucao = s_lista_ja_marcado_devolucao & ";"
				s_lista_ja_marcado_devolucao = s_lista_ja_marcado_devolucao & Trim("" & r("id_registro"))
				end if
		elseif Trim("" & r("operacao")) = PERDA then
			if s_lista_completa_perda <> "" then s_lista_completa_perda = s_lista_completa_perda & ";"
			s_lista_completa_perda = s_lista_completa_perda & Trim("" & r("id_registro"))
			s_class = s_class & " CKB_COM_PERDA"
			if CLng(r("status_comissao")) = CLng(COD_COMISSAO_DESCONTADA) then
				s_checked = " checked"
				s_class_td = s_class_td & " CKB_HIGHLIGHT"
				if s_lista_ja_marcado_perda <> "" then s_lista_ja_marcado_perda = s_lista_ja_marcado_perda & ";"
				s_lista_ja_marcado_perda = s_lista_ja_marcado_perda & Trim("" & r("id_registro"))
				end if
			end if
		x = x & "		<td class='MDTE tdCkb" & s_class_td & "' align='center'>" & _
                            "<input type='checkbox' class='CKB_COM " & s_class & "' name='" & s_id & "' id='" & s_id & "' value='" & Trim("" & r("id_registro")) & "|" & Trim("" & r("operacao")) & "'" & s_checked & " />" & _
                            "<input type='hidden' name='" & s_id & "_original" & "' id='" & s_id & "_original" & "' value='" & Trim("" & r("status_comissao")) & "' />" & _
                "</td>" & chr(13)
		
	 '> LOJA
		x = x & "		<td class='MTD tdLoja' align='center'><span class='Cnc' style='color:" & s_cor & ";'>" & Trim("" & r("loja")) & "</span></td>" & chr(13)

	 '> Nº ORÇAMENTO
		s = Trim("" & r("orcamento"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdOrcamento' align='left'><span class='Cn'><a style='color:" & s_cor & ";' href='javascript:fORCConsulta(" & _
				chr(34) & s & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o orçamento'>" & _
				s & "</a></span></td>" & chr(13)

	 '> Nº PEDIDO
		s_nome_cliente = Trim("" & r("nome_iniciais_em_maiusculas"))
		s_nome_cliente = Left(s_nome_cliente, 15)
		
		x = x & "		<td class='MTD tdPedido' align='left'><span class='Cn'><a style='color:" & s_cor & ";' href='javascript:fPEDConsulta(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o pedido'>" & _
				Trim("" & r("pedido")) & "<br>" & s_nome_cliente & "</a></span></td>" & chr(13)

	 '> DATA
		s = formata_data(r("data"))
		x = x & "		<td align='center' class='MTD tdData'><span class='Cnc' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)

	 '> VALOR DO PEDIDO (PREÇO DE VENDA)
		x = x & "		<td align='right' class='MTD tdVlPedido'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_preco_venda) & "</span></td>" & chr(13)

	 '> COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
		x = x & "		<td align='right' class='MTD tdVlRT'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT) & "</span></td>" & chr(13)

	 '> RA BRUTO
		x = x & "		<td align='right' class='MTD tdVlRABruto'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA) & "</span></td>" & chr(13)

	 '> RA LÍQUIDO
		x = x & "		<td align='right' class='MTD tdVlRALiq'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido) & "</span></td>" & chr(13)

	 '> RA DIFERENÇA
		x = x & "		<td align='right' class='MTD tdVlRADif'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_diferenca) & "</span></td>" & chr(13)

	 '> STATUS DE PAGAMENTO
		x = x & "		<td class='MTD tdStPagto' align='left'><span class='Cn' style='color:" & s_cor & ";'>" & x_status_pagto(Trim("" & r("st_pagto"))) & "</span></td>" & chr(13)

	 '> +/-
		x = x & "		<td align='center' class='MTD tdSinal'><span class='C' style='font-family:Courier,Arial;color:" & s_cor_sinal & "'>" & s_sinal & "</span></td>" & chr(13)
		
	 '> COLUNA DA FIGURA (EXPANDE/RECOLHE)
		x = x & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13)
		
		vl_sub_total_preco_venda = vl_sub_total_preco_venda + r("total_preco_venda")
		        vl_total_preco_venda = vl_total_preco_venda + r("total_preco_venda")
		        vl_sub_total_preco_NF = vl_sub_total_preco_NF + r("total_preco_NF")
		        vl_total_preco_NF = vl_total_preco_NF + r("total_preco_NF")
		        vl_sub_total_RT = vl_sub_total_RT +  vl_RT
		        vl_total_RT = vl_total_RT + vl_RT        
		        vl_sub_total_RA = vl_sub_total_RA +  vl_RA       
		        vl_total_RA = vl_total_RA + vl_RA
		        vl_sub_total_RA_liquido = vl_sub_total_RA_liquido + vl_RA_liquido
		        vl_total_RA_liquido = vl_total_RA_liquido + vl_RA_liquido
		        vl_sub_total_RA_diferenca = vl_sub_total_RA_diferenca + vl_RA_diferenca
		        vl_total_RA_diferenca = vl_total_RA_diferenca + vl_RA_diferenca
                sub_total_comissao = vl_sub_total_RT + vl_sub_total_RA_liquido


                if vl_sub_total_RT >= 0 and vl_sub_total_RA_liquido >= 0 then 
                    vl_sub_total_RT_arredondado = Clng(o.digitoFinal(formata_moeda(vl_sub_total_RT)))
                    vl_sub_total_RA_arredondado = Clng(o.digitoFinal(formata_moeda(vl_sub_total_RA_liquido)))
                else
                     vl_sub_total_RT_arredondado =vl_sub_total_RT
                     vl_sub_total_RA_arredondado =vl_sub_total_RA_liquido
                end if

               if vl_sub_total_RT_arredondado >=0 or vl_sub_total_RA_arredondado >=0 then
                    if vl_sub_total_RT_arredondado < 0 or vl_sub_total_RA_arredondado <0 then
                        sub_total_com = o.digitoFinal(vl_sub_total_RT+vl_sub_total_RA_liquido)
                    else 
                        sub_total_com = vl_sub_total_RT_arredondado + vl_sub_total_RA_arredondado
                    end if
                end if
                        
                if vl_sub_total_RT <= 0 and vl_sub_total_RA_liquido <= 0 then 
                    sub_total_com = vl_sub_total_RT_arredondado + vl_sub_total_RA_arredondado
                end if 
		
		x = x & "	</tr>" & chr(13)
		
		if (n_reg_total mod 50) = 0 then
			Response.Write x
			x = ""
			end if
			banco = rs("banco")
        ind_anterior = r("indicador")
		r.MoveNext
		loop
		
  ' MOSTRA TOTAL DO ÚLTIMO INDICADOR

	if n_reg <> 0 then 
		s_cor="black"
		if vl_sub_total_preco_venda < 0 then s_cor="red"
		if vl_sub_total_RT < 0 then s_cor="red"
		if vl_sub_total_RA < 0 then s_cor="red"
		if vl_sub_total_RA_liquido < 0 then s_cor="red"
        sub_total_comissao = vl_sub_total_RT + vl_sub_total_RA_liquido
 
        z=0

        if sub_total_com <= 0 Or sub_total_com > 300 Or banco = "237" then
            if (sub_total_com =0 Or sub_total_com > 300) And banco <> "237" then                   
                   totalComChqOutros = totalComChqOutros + sub_total_com
                   qtdeChq = qtdeChq +1  
            elseif sub_total_com <0 then  
                   
            else                    
                   totalComchqBradesco = totalComchqBradesco + sub_total_com           
            end if
        else
            dadosCalculo = sub_total_com
            dadosCalculo = o.CalculaCedulas(dadosCalculo,"2#"& limitador(5) &"|5#"&limitador(4)&"|10#"&limitador(3)&"|20#"&limitador(2)&"|50#"&limitador(1)&"|100#"&limitador(0)&"",resultadoCalculo)
            dadosCalculo = resultadocalculo
            v_cedulas = Split(dadosCalculo,"|")
            for cont=0 to Ubound(v_cedulas)
                if cont mod 2 = 0 then
                    redim  preserve cedulas(j)
                    ' cedulas(j)=  cint(v_cedulas(cont))            
                    j = j +1
                else
                        redim preserve qtdeCedula(z)
                        qtdeCedula(z) = cint(v_cedulas(cont)) 
                        z = z + 1                   
                end if
            next 
        end if
        if rs2.State <> 0 then rs2.Close 
        rs2.Open "SELECT COUNT(*) qtde_Desconto, descricao,valor,ordenacao  FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (apelido = '" & ind_anterior & "') GROUP BY  descricao,valor,ordenacao ORDER BY ordenacao", cn
        msg_desconto = ""
        if Not rs2.Eof then
            
                valor_desconto = 0
                contador = 0
                qtde_registro_desc = 0
                Erase v_desconto_descricao
                Erase v_desconto_valor
                do while Not rs2.EoF         
                    redim preserve v_desconto_descricao(contador) 
                    redim preserve  v_desconto_valor(contador)                 
                    v_desconto_descricao(contador) = rs2("descricao")
                    v_desconto_valor(contador) = rs2("valor")
                    valor_desconto = valor_desconto + v_desconto_valor(contador)
                    qtde_registro_desc = qtde_registro_desc + 1
                    contador = contador + 1
                    rs2.MoveNext
		        loop
                msg_desconto =  "Planilha de Desconto:&nbsp; " & Cstr(qtde_registro_desc) & " registro(s) no valor total de&nbsp;" & formata_moeda(valor_desconto)                       
            else
                msg_desconto= ""
            
        end if
		x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
				"		<td colspan='5' class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
				"TOTAL:</span></td>" & chr(13) & _
				"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _
				"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _
				"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</span></td>" & chr(13) & _
				"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _
				"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _
				"		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
				"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
                   "   <tr>" & chr(13) 
                   if sub_total_com <= 0 Or o.DigitoFinal(sub_total_com) > 300 Or banco = "237" then
                        x = x & "<td align='left' colspan='10' nowrap><span class='Cd' style='color:" & s_cor & ";'></span>" 
                                    if msg_desconto <> "" then x = x & "<span class='Cd'><a href='javascript:abreDesconto(" & idx_bloco  & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"
                   else if sub_total_com >= 0 then
                        x = x & "       <td align='left' colspan='10' nowrap>"
                        if msg_desconto <> "" then 
                            x = x & "<span class='Cd' ><a href='javascript:abreDesconto(" & idx_bloco  & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"
                        else   
                            x = x & "<span class='Rd' style='color: black;'>" 
                            for cont = 0 to UBound(qtdeCedula)
                               if (cont = 0 And qtdeCedula(cont) <> 0) then                 
                                            aux(cont) = aux(cont) + qtdeCedula(cont)
                                elseif (cont = 1 And qtdeCedula(cont) <> 0) then
                                             aux(cont) = aux(cont) + qtdeCedula(cont)
                                elseif (cont = 2 And qtdeCedula(cont) <> 0) then
                                             aux(cont) = aux(cont) + qtdeCedula(cont)
                                elseif (cont = 3 And qtdeCedula(cont) <> 0) then
                                             aux(cont) = aux(cont) + qtdeCedula(cont)
                                elseif (cont = 4 And qtdeCedula(cont) <> 0) then
                                             aux(cont) = aux(cont) + qtdeCedula(cont)
                                elseif (cont = 5 And qtdeCedula(cont) <> 0) then
                                             aux(cont) = aux(cont) + qtdeCedula(cont)
                                end if
                        
                            next                    
                            x = x & "</span></td>" 
                        end if
                   end if
                   end if
                   if sub_total_com_RT + sub_total_com_RA >= 0 then 
                    s_cor = "black"
                   else 
                    s_cor = "red"
                   end if
                     x = x & "       <td align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>"
                   if sub_total_com <= 0 then
                   x = x & "&nbsp;</span></td>" & chr(13)
                    elseif sub_total_com > 300 And banco <> "237" then
                   x = x & "CHQ:</span></td>" & chr(13)
                    elseif banco = "237" then
                   x = x & "DEP:</span></td>" & chr(13)      
                    else
                   x = x & "DIN:</span></td>" & chr(13)
                    end if

                   if sub_total_com >=0  then                        
                      x = x & "       <td align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(sub_total_com) & "</span></td>" & chr(13)
                   end if
                    x = x & "	    </tr>" & chr(13) 
                   if msg_desconto <> "" then
                        x = x &"   <tr>" & chr(13) & _
                                "          <td  class='VISAO_ANALIT' id='table_Desconto_"& idx_bloco   &"'"" colspan='15' >" & chr(13)& _
                                "          <table colspan='2' align='left' >"& chr(13)
                                for contador = 0 to Ubound(v_desconto_descricao)                                 
                                    x = x & "   <tr>" & chr(13)& _
                                            "       <td width='15'>&nbsp;</td>" & chr(13)& _
                                            "       <td  align='left' width='400' ><span class='Cd'style='color: red;' >"& v_desconto_descricao(contador)& "</span></td>"& _
                                            "       <td align='left' ><span class='Cd'style='color: red;' > R$ "& formata_moeda(v_desconto_valor(contador))& "</span></td>"& _
                                            "   </tr>"                                     
                                next          
                                x = x & "   </table>"& chr(13)& _
                                        "   </td>"& chr(13)& _
                                        "</tr>"
                    end if
		
	'>	TOTAL GERAL
		if qtde_indicadores > 1 then
			s_cor="black"
			if vl_total_preco_venda < 0 then s_cor="red"
			if vl_total_RT < 0 then s_cor="red"
			if vl_total_RA < 0 then s_cor="red"
			if vl_total_RA_liquido < 0 then s_cor="red"
			x = x & "	<tr>" & chr(13) & _
					"		<td colspan='12' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<tr>" & chr(13) & _
					"		<td colspan='12' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<tr nowrap style='background:honeydew'>" & chr(13) & _
					"		<td class='MTBE' colspan='5' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL:</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_preco_venda) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_diferenca) & "</span></td>" & chr(13) & _
					"		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
                    " </table>" & chr(13)


                    x = x & "<br>"& chr(13)
            '    TOTAL DE CEDULAS
                if aux(0) <> 0 then                 
                    textDin = textDin + Cstr(aux(0)) + "&times;100,00 "
                    if ((aux(1) <> 0) OR (aux(2) <> 0) OR (aux(3) <> 0) OR (aux(4) <> 0) OR (aux(5) <> 0)) then textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;" 
                end if
                if aux(1) <> 0 then
                    textDin = textDin + Cstr(aux(1)) + "&times;50,00 "
                    if ((aux(2) <> 0) OR (aux(3) <> 0) OR (aux(4) <> 0) OR (aux(5) <> 0)) then textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"
                end if
                if aux(2) <> 0 then
                    textDin = textDin + Cstr(aux(2)) + "&times;20,00 "
                    if ((aux(3) <> 0) OR (aux(4) <> 0) OR (aux(5) <> 0)) then textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"
                end if
                if aux(3) <> 0 then
                    textDin = textDin + Cstr(aux(3)) + "&times;10,00 "
                    if ((aux(4) <> 0) OR (aux(5) <> 0)) then textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"
                end if
                if aux(4) <> 0 then
                    textDin = textDin + Cstr(aux(4)) + "&times;5,00 "
                    if  (aux(5) <> 0) then textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"
                end if
                if aux(5) <> 0 then
                    textDin = textDin + Cstr(aux(5)) + "&times;2,00 "
                end if    
           '    TOTAL CHQ textChq
                if qtdeChq <> 0 then                 
                    textChq = textChq + formata_moeda(totalComChqOutros) + "&nbsp;(Nº cheques " +  Cstr(qtdeChq) + ")"                   
                else
                    textChq = textChq + formata_moeda(totalComChqOutros)
                end if
        
                    x = x & " <table  cellspacing='0'  width='700px'> " & chr(13) & _
                    "   <tr  nowrap style='background:honeydew'>"& chr(13) & _
                    "       <td width='30%' class='MDTE' align='left'><span class='Cd' style='color:black;' > TOTAL COMISSÃO</td> "& chr(13) & _
                    "       <td class='MD MC'  style='background:honeydew'><span class='Cd'>" & formata_moeda(totalComDin+totalComChqOutros+totalComChqBradesco) &" </td> "& chr(13) & _
                    "   </tr>"& chr(13) & _
                    "   <tr nowrap>"& chr(13) & _
                            "<td  style='background:honeydew;' width='30%' class='MDTE' align='left'><span class='Cd' style='color:black;' valign='bottom'>Comissão em CHQ</td>" & chr(13) & _
                    "       <td class='MD MC'  style='background:honeydew'><span class='Cd'>" & textChq & "</td>" & chr(13) & _
                    "   </tr>"& chr(13) & _
                    "    <tr nowrap >"& chr(13) & _
                    "       <td  style='background:honeydew' width='30%' class='MDTE' align='left'><span class='Cd' style='color:black;' >Comissão em DEP</td> "& chr(13) & _
                    "       <td class='MTD' style='background:honeydew'><span class='Cd'>"& formata_moeda(totalComChqBradesco)&" </td> "& chr(13) & _
                    "   </tr>"& chr(13) & _
                    "   <tr nowrap >"& chr(13) & _
                    "       <td style='background:honeydew' width='30%' class='MDTE' align='left'><span class='Cd' style='color:black;' >Comissão em DIN:</td>" & chr(13) & _
                    "       <td class='MD MC' style='background:honeydew'><span class='Cd'>"& formata_moeda(totalComDin)&"</td>"& chr(13) & _
                    "   </tr>"& chr(13) & _
                    "   <tr nowrap >"& chr(13) & _
                    "       <td  style='background:honeydew' width='30%' class='MTBE MD' align='left'><span class='Cd' style='color:black;' >Qtde Cedulas para Comissão em DIN</td>"& chr(13) & _
                    "       <td class='MTB MD' align='left'  style='background:honeydew'><span class='Cd'>" & textDin &" </td>"& chr(13) & _
                    "   </tr>"& chr(13) & _
                    "</table>"& chr(13) 

                    
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<tr nowrap>" & chr(13) & _
				"		<td class='MT ALERTA' colspan='12' align='center'><span class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS NO PERÍODO ESPECIFICADO&nbsp;</span></td>" & chr(13) & _
				"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO INDICADOR
	x = x & "</table>" & chr(13)
	
	Response.write x

	x = "<input type='hidden' name='c_lista_completa_venda_normal' id='c_lista_completa_venda_normal' value='" & s_lista_completa_venda_normal & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_completa_devolucao' id ='c_lista_completa_devolucao' value='" & s_lista_completa_devolucao & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_completa_perda' id='c_lista_completa_perda' value='" & s_lista_completa_perda & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_ja_marcado_venda_normal' id='c_lista_ja_marcado_venda_normal' value='" & s_lista_ja_marcado_venda_normal & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_ja_marcado_devolucao' id='c_lista_ja_marcado_devolucao' value='" & s_lista_ja_marcado_devolucao & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_ja_marcado_perda' id='c_lista_ja_marcado_perda' value='" & s_lista_ja_marcado_perda & "' />" & chr(13)
	Response.write x
	
	if r.State <> 0 then r.Close
    if rs2.State <> 0 then rs2.Close 
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
	
	// VISÃO SINTÉTICA?
	if ($("#rb_visao").val() == "SINTETICA") {
		$(".CKB_COM").attr("disabled", true);
	}
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
var s_id, s_class;
	s_id = "#ckb_comissao_paga_tit_bloco_" + idx_bloco;
	s_class = ".CKB_COM_BL_" + idx_bloco;
	if ($(s_id).is(":checked")) {
		$(s_class).prop("checked", true);
		$(s_class).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
	}
	else {
		$(s_class).prop("checked", false);
		$(s_class).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
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
	dCONFIRMA.style.visibility = "hidden";
	f.action = "RelComissaoIndicadoresGravaDados.asp";
	f.submit();
}
</script>

<script type="text/javascript">
function limitaCedulas() {
    fREL.limita_cedula.value="1";
    fREL.action = "RelComissaoIndicadoresExec.asp";
    fREL.submit();
}
</script>
<script type="text/javascript">
    function abreLimitaCedulas() {
        $("#div_cedulas").css('display', 'block');
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
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
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
<input type="hidden" name="ckb_st_pagto_pago" id="ckb_st_pagto_pago" value="<%=ckb_st_pagto_pago%>">
<input type="hidden" name="ckb_st_pagto_nao_pago" id="ckb_st_pagto_nao_pago" value="<%=ckb_st_pagto_nao_pago%>">
<input type="hidden" name="ckb_st_pagto_pago_parcial" id="ckb_st_pagto_pago_parcial" value="<%=ckb_st_pagto_pago_parcial%>">
<input type="hidden" name="rb_pagto_comissao_via_cartao" id="rb_pagto_comissao_via_cartao" value="<%=rb_pagto_comissao_via_cartao%>" />
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>">
<input type="hidden" name="c_usuario_sessao" id="c_usuario_sessao" value="<%=usuario%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
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
	s_filtro = "<table width='709' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

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
					"		<td align='right' valign='top' nowrap><span class='N'>Período Entrega:&nbsp;</span></td>" & chr(13) & _
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

'	PAGAMENTO DA COMISSÃO VIA CARTÃO
	s = ""
	if rb_pagto_comissao_via_cartao = "0" then
		s = "não"
	elseif rb_pagto_comissao_via_cartao = "1" then
		s = "sim"
	else
		s = "todos"
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Comissão via Cartão:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
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
					"		<td align='right' valign='top' nowrap><span class='N'>Indicador:&nbsp;</span></td>" & chr(13) & _
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
					"		<td align='right' valign='top' nowrap><span class='N'>Vendedor:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	LOJA(S)
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
				"		<td align='right' valign='top' nowrap><span class='N'>Loja(s):&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	EMISSÃO
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td class="Rc" align="left">&nbsp;</td>
</tr>
</table>

<% if blnVisaoSintetica then %>
<table class="notPrint" width="709" cellpadding="0" cellspacing="0" style="margin-top:5px;">
<tr>
	<td align="right">
		<button type="button" name="bExpandirTodos" id="bExpandirTodos" class="Button BTN_LNK" onclick="expandir_todos();" title="expandir todas as linhas de dados" style="margin-left:6px;margin-bottom:2px">Expandir Tudo</button>
		&nbsp;
		<button type="button" name="bRecolherTodos" id="bRecolherTodos" class="Button BTN_LNK" onclick="recolher_todos();" title="recolher todas as linhas de dados" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Recolher Tudo</button>
	</td>
</tr>
</table>
<br />
<table class="notPrint" width="709" cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="A1" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>
<% else %>
<table class="notPrint" width="709" cellpadding="0" cellspacing="0" style="margin-top:5px;">
<tr>
	<td align="right">
		<button type="button" name="bExpandirTodos" id="bExpandirTodos" class="Button BTN_LNK" onclick="expandir_todos();" title="expandir todas as linhas de dados" style="margin-left:6px;margin-bottom:2px">Expandir Tudo</button>
		&nbsp;
		<button type="button" name="bRecolherTodos" id="bRecolherTodos" class="Button BTN_LNK" onclick="recolher_todos();" title="recolher todas as linhas de dados" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Recolher Tudo</button>
		&nbsp;
		<button type="button" name="bMarcarTodos" id="bMarcarTodos" class="Button BTN_LNK" onclick="marcar_todos();" title="assinala todos os pedidos para gravar o status da comissão como paga" style="margin-left:6px;margin-bottom:2px">Marcar todos</button>
		&nbsp;
		<button type="button" name="bDesmarcarTodos" id="bDesmarcarTodos" class="Button BTN_LNK" onclick="desmarcar_todos();" title="desmarca todos os pedidos para gravar o status da comissão como não-paga" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Desmarcar todos</button>
	</td>
</tr>
</table>

<br />
<table class="notPrint" width="709" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="left">&nbsp;</td>
	<td align="right">
		<div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
<% end if %>

</form>

</center>

<div id="divPedidoConsulta"><center><div id="divInternoPedidoConsulta"><img id="imgFechaDivPedidoConsulta" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframePedidoConsulta"></iframe></div></center></div>

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
