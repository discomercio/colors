<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L C O M I S S A O E X E C . A S P
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
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_COMISSAO_VENDEDORES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro
	dim c_dt_inicio, c_dt_termino, c_vendedor, c_lojas
	dim rb_visao, blnVisaoSintetica
	
	alerta = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	c_lojas = Trim(Request.Form("c_lojas"))
	rb_visao = Trim(Request.Form("rb_visao"))
	
	blnVisaoSintetica = False
	if rb_visao = "SINTETICA" then blnVisaoSintetica = True

	if c_dt_inicio = "" then
		alerta = "DATA DE INÍCIO DO PERÍODO NÃO FOI PREENCHIDA."
	elseif c_dt_termino = "" then
		alerta = "DATA DE TÉRMINO DO PERÍODO NÃO FOI PREENCHIDA."
	elseif Not IsDate(StrToDate(c_dt_inicio)) then
		alerta = "DATA DE INÍCIO DO PERÍODO É INVÁLIDA."
	elseif Not IsDate(StrToDate(c_dt_termino)) then
		alerta = "DATA DE TÉRMINO DO PERÍODO É INVÁLIDA."
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
dim s, s_aux, s_sql, x, cab_table, cab, vendedor_a, n_reg, n_reg_total, qtde_vendedores
dim vl_comissao, vl_saida, vl_total_saida, vl_total_comissao, vl_sub_total_saida, vl_sub_total_comissao
dim s_where, s_where_venda, s_where_devolucao, s_where_perdas, s_where_lista_lojas, s_cor, s_sinal, s_cor_sinal
dim s_new_cab, idx_bloco
dim v_lojas, i

'	CRITÉRIOS COMUNS
	s_where = ""
	if c_vendedor <> "" then
		s = substitui_caracteres(c_vendedor, "*", BD_CURINGA_TODOS)
		s_aux = "="
		if Instr(1, s, BD_CURINGA_TODOS) <> 0 then s_aux = "LIKE"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.vendedor " & s_aux & " '" & s & "'" & SQL_COLLATE_CASE_ACCENT & ")"
		end if
	
	if c_lojas <> "" then
		v_lojas = Split(c_lojas, ", ")
		s_where_lista_lojas = ""
		for i=LBound(v_lojas) to UBound(v_lojas)
			if Trim("" & v_lojas(i)) <> "" then
				if s_where_lista_lojas <> "" then s_where_lista_lojas = s_where_lista_lojas & ", "
				s_where_lista_lojas = s_where_lista_lojas & "'" & Trim("" & v_lojas(i)) & "'"
				end if
			next
		if s_where_lista_lojas <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_PEDIDO.loja IN (" & s_where_lista_lojas & "))"
			end if
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
	
	
' 	A) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
'	B) O CAMPO 'QTDE' A SER USADO DEVE SER DA TABELA T_ESTOQUE_MOVIMENTO, JÁ
'	QUE UM PEDIDO PODE TER USADO DIVERSOS LOTES DO ESTOQUE PARA ATENDER A
'	UM ÚNICO PRODUTO.  NESSE CASO, HAVERÁ MAIS DE UM REGISTRO EM 
'	T_ESTOQUE_MOVIMENTO SE RELACIONANDO COM O MESMO REGISTRO DE T_PEDIDO_ITEM.
'	A SOMA DE 'QTDE' DOS REGISTROS DE T_ESTOQUE_MOVIMENTO RESULTAM NO VALOR
'	DE 'QTDE' DE T_PEDIDO_ITEM.
	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then s = " AND" & s
	s_sql = "SELECT t_PEDIDO.vendedor, t_PEDIDO.loja AS loja, CONVERT(smallint,t_PEDIDO.loja) AS numero_loja," & _
			" t_PEDIDO.entregue_data AS data," & _
			" t_PEDIDO.pedido AS pedido," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda) AS valor_saida," & _
			" Sum((t_PEDIDO_ITEM.comissao/100)*(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda)) AS valor_comissao" & _
			" FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
			" INNER JOIN t_PEDIDO_ITEM ON ((t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto))" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque='" & ID_ESTOQUE_ENTREGUE & "')" & _
			s & _
			" GROUP BY t_PEDIDO.vendedor, t_PEDIDO.loja, t_PEDIDO.entregue_data, t_PEDIDO.pedido"

'	ITENS DEVOLVIDOS
'	LEMBRE-SE: NA DEVOLUÇÃO DO PRODUTO, É CRIADA UMA ENTRADA NO ESTOQUE DE VENDA P/
'	REPRESENTAR A ENTRADA DA MERCADORIA NO ESTOQUE. ENTRETANTO, A QUANTIDADE
'	DEVOLVIDA FICA INICIALMENTE TODA ALOCADA P/ O ESTOQUE DE DEVOLUÇÃO, DEVIDO
'	À NECESSIDADE DE TRATAR A MERCADORIA ANTES DE DISPONIBILIZA-LA P/ VENDA.
'	IMPORTANTE: NO CASO DE OCORRER A DEVOLUÇÃO DE VÁRIAS UNIDADES, PODEM SER
'	CRIADOS VÁRIOS REGISTROS DE ESTOQUE A DIFERENTES CUSTOS DE AQUISIÇÃO.
	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT t_PEDIDO.vendedor, t_PEDIDO.loja AS loja, CONVERT(smallint,t_PEDIDO.loja) AS numero_loja," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data," & _
			" t_PEDIDO.pedido AS pedido," & _
			" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS valor_saida," & _
			" Sum((t_PEDIDO_ITEM_DEVOLVIDO.comissao/100)*(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda)) AS valor_comissao" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)" & _
			" INNER JOIN t_ESTOQUE ON (t_PEDIDO_ITEM_DEVOLVIDO.id=t_ESTOQUE.devolucao_id_item_devolvido)" & _
			" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_ESTOQUE_ITEM.produto))" & _ 
			s & _
			" GROUP BY t_PEDIDO.vendedor, t_PEDIDO.loja, t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data, t_PEDIDO.pedido"

'	PERDAS
	s = s_where
	if (s <> "") And (s_where_perdas <> "") then s = s & " AND"
	s = s & s_where_perdas
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT t_PEDIDO.vendedor, t_PEDIDO.loja AS loja, CONVERT(smallint,t_PEDIDO.loja) AS numero_loja," & _
			" t_PEDIDO_PERDA.data AS data," & _
			" t_PEDIDO.pedido AS pedido," & _
			" Sum(-t_PEDIDO_PERDA.valor) AS valor_saida," & _
			" 0 AS valor_comissao" & _
			" FROM t_PEDIDO_PERDA INNER JOIN t_PEDIDO ON (t_PEDIDO_PERDA.pedido=t_PEDIDO.pedido)" & _
			s & _
			" GROUP BY t_PEDIDO.vendedor, t_PEDIDO.loja, t_PEDIDO_PERDA.data, t_PEDIDO.pedido"
	
	s_sql = s_sql & " ORDER BY vendedor, numero_loja, data, pedido, valor_saida DESC"

  ' CABEÇALHO
	cab_table = "<table cellspacing='0' id='tableDados'>" & chr(13)
	cab = "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='MDTE tdLoja' align='left' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>LOJA</span></td>" & chr(13) & _
		  "		<td class='MTD tdPedido' align='left' valign='bottom' nowrap><span class='R VISAO_ANALIT'>Nº PEDIDO</span></td>" & chr(13) & _
		  "		<td class='MTD tdData' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>DATA</span></td>" & chr(13) & _
		  "		<td class='MTD tdValor' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VALOR</span></td>" & chr(13) & _
		  "		<td class='MTD tdCom' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>COMISSÃO</span></td>" & chr(13) & _
		  "		<td class='MTD tdSinal' align='center' valign='bottom'><span class='Rc' style='font-weight:bold;'>+/-</span></td>" & chr(13) & _
		  "		<td valign='bottom' class='notPrint BkgWhite' align='left'>&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & "_NNNNN_" & chr(34) & ");' title='exibe ou oculta os dados'><img src='../botao/view_bottom.png' border='0'></a></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	idx_bloco = 0
	qtde_vendedores = 0
	vl_total_saida = 0
	vl_total_comissao = 0
	vl_sub_total_saida = 0
	vl_sub_total_comissao = 0

	vendedor_a = "XXXXXXXXXXXX"
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE VENDEDOR?
		if Trim("" & r("vendedor"))<>vendedor_a then
			vendedor_a = Trim("" & r("vendedor"))
			idx_bloco = idx_bloco + 1
			qtde_vendedores = qtde_vendedores + 1
		  ' FECHA TABELA DO VENDEDOR ANTERIOR
			if n_reg_total > 0 then 
				s_cor="black"
				if vl_sub_total_saida < 0 then s_cor="red"
				if vl_sub_total_comissao < 0 then s_cor="red"
				x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
						"		<td class='MTBE' align='right' colspan='3' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
						"TOTAL:</span></td>" & chr(13) & _
						"		<td class='MTB tdValor' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida) & "</span></td>" & chr(13) & _
						"		<td class='MTB tdCom' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_comissao) & "</span></td>" & chr(13) & _
						"		<td class='MTBD tdSinal' align='right'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
						"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
						"	</tr>" & chr(13) & _
						"</table>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			n_reg = 0
			vl_sub_total_saida = 0
			vl_sub_total_comissao = 0

			if n_reg_total > 0 then x = x & "<BR>"
			s = Trim("" & r("vendedor"))
			s_aux = x_usuario(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & Replace(cab_table, "tableDados", "tableDados_" & idx_bloco)
			if s <> "" then x = x & "	<tr>" & chr(13) & _
									"		<td class='MDTE' colspan='6' align='left' valign='bottom' class='MB' style='background:azure;'><span class='N'>&nbsp;" & s & "</span></td>" & chr(13) & _
									"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
									"	</tr>" & chr(13)
			s_new_cab = cab
			s_new_cab = Replace(s_new_cab, "_NNNNN_", CStr(idx_bloco))
			x = x & s_new_cab
			end if
		
		
	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<tr nowrap class='VISAO_ANALIT'>"  & chr(13)

	'   EVITA DIFERENÇAS DE ARREDONDAMENTO
		vl_saida = converte_numero(formata_moeda(r("valor_saida")))
		vl_comissao = converte_numero(formata_moeda(r("valor_comissao")))
		if (vl_saida < 0) Or (vl_comissao < 0) then
			s_cor = "red"
			s_cor_sinal = "red"
			s_sinal = "-"
		else
			s_cor = "black"
			s_cor_sinal = "green"
			s_sinal = "+"
			end if

	 '> LOJA
		x = x & "		<td class='MDTE tdLoja' align='center'><span class='Cnc' style='color:" & s_cor & ";'>" & Trim("" & r("loja")) & "</span></td>" & chr(13)

	 '> Nº PEDIDO
		x = x & "		<td class='MTD tdPedido' align='left'><span class='Cn'><a style='color:" & s_cor & ";' href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & _
				Trim("" & r("pedido")) & "</a></span></td>" & chr(13)

	 '> DATA DO PEDIDO
		s = formata_data(r("data"))
		x = x & "		<td align='center' class='MTD tdData'><span class='Cnc' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)

	 '> VALOR DO PEDIDO
		x = x & "		<td align='right' class='MTD tdValor'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_saida) & "</span></td>" & chr(13)

	 '> VALOR DA COMISSÃO
		x = x & "		<td align='right' class='MTD tdCom'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_comissao) & "</span></td>" & chr(13)

	 '> +/-
		x = x & "		<td align='center' class='MTD tdSinal'><span class='C' style='font-family:Courier,Arial;color:" & s_cor_sinal & "'>" & s_sinal & "</span></td>" & chr(13)

	 '> COLUNA DA FIGURA (EXPANDE/RECOLHE)
		x = x & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13)

		vl_sub_total_saida = vl_sub_total_saida + r("valor_saida")
		vl_total_saida = vl_total_saida + r("valor_saida")
		vl_sub_total_comissao = vl_sub_total_comissao + vl_comissao
		vl_total_comissao = vl_total_comissao + vl_comissao
		
		x = x & "	</tr>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		r.MoveNext
		loop
	
  ' MOSTRA TOTAL DO ÚLTIMO VENDEDOR
	if n_reg <> 0 then 
		s_cor="black"
		if vl_sub_total_saida < 0 then s_cor="red"
		if vl_sub_total_comissao < 0 then s_cor="red"
		x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
				"		<td colspan='3' class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
				"TOTAL:</span></td>" & chr(13) & _
				"		<td class='MTB tdValor' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida) & "</span></td>" & chr(13) & _
				"		<td class='MTB tdCom' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_comissao) & "</span></td>" & chr(13) & _
				"		<td class='MTBD tdSinal' align='right'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
				"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13)
		
	'>	TOTAL GERAL
		if qtde_vendedores > 1 then
			s_cor="black"
			if vl_total_saida < 0 then s_cor="red"
			if vl_total_comissao < 0 then s_cor="red"
			x = x & "	<tr>" & chr(13) & _
					"		<td colspan='6' align='left' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<tr>" & chr(13) & _
					"		<td colspan='6' align='left' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<tr nowrap style='background:honeydew'>" & chr(13) & _
					"		<td class='MTBE' colspan='3' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL:</span></td>" & chr(13) & _
					"		<td class='MTB tdValor' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_saida) & "</span></td>" & chr(13) & _
					"		<td class='MTB tdCom' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_comissao) & "</span></td>" & chr(13) & _
					"		<td class='MTBD tdSinal' align='right'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<tr nowrap>" & chr(13) & _
				"		<td class='MT ALERTA' colspan='6' align='center'><span class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS NO PERÍODO ESPECIFICADO&nbsp;</span></td>" & chr(13) & _
				"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO VENDEDOR
	x = x & "</table>" & chr(13)
	
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
window.status='Aguarde, executando a consulta ...';

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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
.tdLoja{
	width: 35px;
	}
.tdPedido{
	width: 80px;
	}
.tdData{
	width: 70px;
	}
.tdValor{
	width: 100px;
	}
.tdCom{
	width: 100px;
	}
.tdSinal{
	width: 20px;
	}
.BTN_LNK
{
	min-width:140px;
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



<% else %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Comissão aos Vendedores</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Período:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & s & "</span></td></tr>" & chr(13)

	s = c_vendedor
	if s = "" then 
		s = "todos"
	elseif (Instr(1,s,"*")=0) And (Instr(1,s,BD_CURINGA_TODOS)=0) then
		s_aux = x_usuario(c_vendedor)
		if s_aux <> "" then s = s & " (" & s_aux & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Vendedor:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>" & chr(13)

	s = c_lojas
	if s = "" then s = "N.I."
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Loja(s):&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>" & chr(13)

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
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width="649" cellpadding="0" cellspacing="0" style="margin-top:5px;">
<tr>
	<td align="right">
		<button type="button" name="bExpandirTodos" id="bExpandirTodos" class="Button BTN_LNK" onclick="expandir_todos();" title="expandir todas as linhas de dados" style="margin-left:6px;margin-bottom:2px">Expandir Tudo</button>
		&nbsp;
		<button type="button" name="bRecolherTodos" id="bRecolherTodos" class="Button BTN_LNK" onclick="recolher_todos();" title="recolher todas as linhas de dados" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Recolher Tudo</button>
	</td>
</tr>
</table>

<br>

<table class="notPrint" width="649" cellspacing="0">
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
