<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R e l V e n d a s C m v P v E x e c . a s p
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
	if Not operacao_permitida(OP_LJA_REL_FATURAMENTO_CMVPV, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, flag_ok
	dim c_dt_inicio, c_dt_termino, c_loja, c_fabricante, c_produto, c_grupo, c_vendedor, c_indicador, c_pedido, c_empresa
	dim s_nome_vendedor

	alerta = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	c_indicador = Ucase(Trim(Request.Form("c_indicador")))
	c_loja = Trim(Request.Form("c_loja"))
	c_pedido = Ucase(Trim(Request.Form("c_pedido")))
    c_empresa = Trim(Request.Form("c_empresa"))

	if c_pedido <> "" then
		if normaliza_num_pedido(c_pedido) <> "" then c_pedido = normaliza_num_pedido(c_pedido)
		end if
	
	if alerta = "" then
		if c_loja = "" then
			alerta=texto_add_br(alerta)
			alerta = "NÃO FOI INFORMADO O Nº DA LOJA."
			end if
		
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
dim s, s_where, s_where_venda, s_where_devolucao, s_where_loja, s_cor
dim s_aux, s_sql, cab_table, cab, n_reg, n_reg_total, x, fabricante_a
dim perc, vl_total_saida, vl_total_entrada, vl_sub_total_saida, vl_sub_total_entrada
dim qtde_total, qtde_sub_total, qtde_fabricantes

'	CRITÉRIOS COMUNS
	s_where = ""
	if c_grupo <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.grupo = '" & c_grupo & "')"
		
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.grupo IS NOT NULL)"
		end if
	
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.vendedor = '" & c_vendedor & "')"
		end if

	if c_pedido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.pedido = '" & c_pedido & "')"
		end if

    if c_empresa <> "" then
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_PEDIDO.id_nfe_emitente = '" & c_empresa & "')"
    end if
		
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if
		
	s_where_loja = " (t_PEDIDO.numero_loja = " & c_loja & ")"
	
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

	if c_fabricante <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_ESTOQUE_MOVIMENTO.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_ESTOQUE_MOVIMENTO.produto = '" & c_produto & "')"
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
	
	if c_fabricante <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.produto = '" & c_produto & "')"
		end if

'	IMPORTANTE: NESTE RELATÓRIO É USADO O CAMPO 'PRECO_FABRICANTE' DE T_PEDIDO_ITEM
'	==========  E T_PEDIDO_ITEM_DEVOLVIDO, POIS ELES ARMAZENAM O VALOR DE CMV_PV
'				NO MOMENTO EM QUE O PEDIDO FOI CADASTRADO OU EM QUE A MERCADORIA FOI
'				DEVOLVIDA. PORTANTO, O HISTÓRICO DO CMV_PV ESTÁ ARMAZENADO NESSES 
'				CAMPOS.
'				O HISTÓRICO DO CUSTO REAL PAGO AO FABRICANTE ESTÁ ARMAZENADO EM
'				T_ESTOQUE_ITEM.VL_CUSTO2

'	A) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
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
	s_sql = "SELECT" & _
				" t_ESTOQUE_MOVIMENTO.fabricante AS fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto AS produto," & _
				" t_PRODUTO.descricao AS descricao," & _
				" t_PRODUTO.descricao_html AS descricao_html," & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda) AS valor_saida," & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde) AS qtde," & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_fabricante) AS valor_entrada" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
				" INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" INNER JOIN t_PEDIDO_ITEM ON ((t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto))" & _
				" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
			" WHERE" & _
				" (anulado_status=0)" & _
				" AND (estoque='" & ID_ESTOQUE_ENTREGUE & "')" & _
			s & _
			" GROUP BY" & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html"
	
'	LEMBRE-SE: NA DEVOLUÇÃO DO PRODUTO, É CRIADA UMA ENTRADA NO ESTOQUE DE VENDA P/
'	REPRESENTAR A ENTRADA DA MERCADORIA NO ESTOQUE. ENTRETANTO, A QUANTIDADE
'	DEVOLVIDA FICA INICIALMENTE TODA ALOCADA P/ O ESTOQUE DE DEVOLUÇÃO, DEVIDO
'	À NECESSIDADE DE TRATAR A MERCADORIA ANTES DE DISPONIBILIZA-LA P/ VENDA.
'	IMPORTANTE: NO CASO DE OCORRER A DEVOLUÇÃO DE VÁRIAS UNIDADES, PODEM SER
'	CRIADOS VÁRIOS REGISTROS DE ESTOQUE A DIFERENTES CUSTOS DE AQUISIÇÃO.
	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then s = " WHERE " & s
	s_sql = s_sql & " UNION " & _
			"SELECT" & _
				" t_PEDIDO_ITEM_DEVOLVIDO.fabricante AS fabricante," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.produto AS produto," & _
				" t_PRODUTO.descricao AS descricao," & _
				" t_PRODUTO.descricao_html AS descricao_html," & _
				" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS valor_saida," & _
				" Sum(-t_ESTOQUE_ITEM.qtde) AS qtde," & _
				" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_fabricante) AS valor_entrada" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
				" INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" INNER JOIN t_ESTOQUE ON (t_PEDIDO_ITEM_DEVOLVIDO.id=t_ESTOQUE.devolucao_id_item_devolvido)" & _
				" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_ESTOQUE_ITEM.produto))" & _ 
				" LEFT JOIN t_PRODUTO ON ((t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_PRODUTO.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_PRODUTO.produto))" & _
			s & _
			" GROUP BY" & _
				" t_PEDIDO_ITEM_DEVOLVIDO.fabricante," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.produto," & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html"
			
	s_sql = s_sql & _
			" ORDER BY" & _
				" fabricante," & _
				" produto," & _
				" descricao," & _
				" descricao_html," & _
				" qtde DESC"

  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & _
		  "		<TD class='MDTE' valign='bottom' NOWRAP><P style='width:60px' class='R'>Código</P></TD>" & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:201px' class='R'>Descrição</P></TD>" & _
		  "		<TD class='MTD' align='right' valign='bottom' NOWRAP><P style='width:40px' class='Rd' style='font-weight:bold;'>Qtde</P></TD>" & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:80px' class='Rd' style='font-weight:bold;'>Faturamento Total (" & SIMBOLO_MONETARIO & ")</P></TD>" & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:80px' class='Rd' style='font-weight:bold;'>CMV Total (" & SIMBOLO_MONETARIO & ")</P></TD>" & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:80px' class='Rd' style='font-weight:bold;'>Lucro (" & SIMBOLO_MONETARIO & ")</P></TD>" & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:40px' class='Rd' style='font-weight:bold;'>% do Lucro Total</P></TD>" & _
		  "	</TR>" & chr(13)
	
	vl_total_saida = 0
	vl_total_entrada = 0
	set r = cn.execute(s_sql)
	n_reg = 0
	do while Not r.Eof
		n_reg = n_reg + 1
		vl_total_saida = vl_total_saida + r("valor_saida")
		vl_total_entrada = vl_total_entrada + r("valor_entrada")
		r.MoveNext
		loop

	if n_reg > 0 then r.MoveFirst	
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_total = 0
	qtde_sub_total = 0
	qtde_fabricantes = 0
	
	fabricante_a = "XXXXX"
	do while Not r.Eof
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante"))<>fabricante_a then
			fabricante_a = Trim("" & r("fabricante"))
			qtde_fabricantes = qtde_fabricantes + 1
		  ' FECHA TABELA DO FABRICANTE ANTERIOR
			if n_reg_total > 0 then 
				if (vl_total_saida-vl_total_entrada) = 0 then
					perc = 0
				else
					perc = ((vl_sub_total_saida-vl_sub_total_entrada)/(vl_total_saida-vl_total_entrada))*100
					end if
				s_cor="black"
				if qtde_sub_total < 0 then s_cor="red"
				if vl_sub_total_saida < 0 then s_cor="red"
				if vl_sub_total_entrada < 0 then s_cor="red"
				if (vl_sub_total_saida-vl_sub_total_entrada) < 0 then s_cor="red"
				x = x & "<TR NOWRAP style='background: #FFFFDD'>" & _
						"<TD class='MTBE' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
						"TOTAL:</p></td>" & _
						"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_sub_total) & "</p></TD>" & _
						"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida) & "</p></td>" & _
						"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_entrada) & "</p></td>" & _
						"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida-vl_sub_total_entrada) & "</p></td>" & _
						"<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</p></td></TR>" & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>"
				end if

			n_reg = 0
			vl_sub_total_saida = 0
			vl_sub_total_entrada = 0
			qtde_sub_total = 0

			if n_reg_total > 0 then x = x & "<BR>"
			s = Trim("" & r("fabricante"))
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then x = x & "<TR><TD class='MDTE' COLSPAN='7' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td></tr>" & chr(13)
			x = x & cab
			end if
		

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>"

		s_cor="black"
		if IsNumeric(r("qtde")) then if Clng(r("qtde")) < 0 then s_cor="red"

	 '> CÓDIGO DO PRODUTO
		x = x & "		<TD class='MDTE' valign='bottom'><P class='Cn' style='color:" & s_cor & ";'>" & Trim("" & r("produto")) & "</P></TD>"

	 '> DESCRIÇÃO DO PRODUTO
		s = Trim("" & r("descricao_html"))
		if s <> "" then s = produto_formata_descricao_em_html(s) else s = "&nbsp;"
		x = x & "		<TD class='MTD' valign='bottom'><P class='Cn' style='color:" & s_cor & ";'>" & s & "</P></TD>"

	 '> QUANTIDADE
		x = x & "		<TD align='right' valign='bottom' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(r("qtde")) & "</P></TD>"

	 '> VALOR SAÍDA
		x = x & "		<TD align='right' valign='bottom' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_saida")) & "</P></TD>"

	 '> VALOR ENTRADA
		x = x & "		<TD align='right' valign='bottom' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_entrada")) & "</P></TD>"

	 '> LUCRO BRUTO
		x = x & "		<TD align='right' valign='bottom' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_saida")-r("valor_entrada")) & "</P></TD>"

	 '>	PERCENTUAL DO LUCRO BRUTO TOTAL
		if (vl_total_saida-vl_total_entrada) = 0 then
			perc = 0
		else
			perc = ((r("valor_saida")-r("valor_entrada"))/(vl_total_saida-vl_total_entrada))*100
			end if
		x = x & "		<TD align='right' valign='bottom' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</P></TD>"
		
		qtde_sub_total = qtde_sub_total + r("qtde")
		qtde_total = qtde_total + r("qtde")
		vl_sub_total_saida = vl_sub_total_saida + r("valor_saida")
		vl_sub_total_entrada = vl_sub_total_entrada + r("valor_entrada")
		
		x = x & "</TR>" & chr(13)
			
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
  ' MOSTRA TOTAL DO ÚLTIMO FABRICANTE
	if n_reg <> 0 then 
		if (vl_total_saida-vl_total_entrada) = 0 then
			perc = 0
		else
			perc = ((vl_sub_total_saida-vl_sub_total_entrada)/(vl_total_saida-vl_total_entrada))*100
			end if
		
		s_cor="black"
		if qtde_sub_total < 0 then s_cor="red"
		if vl_sub_total_saida < 0 then s_cor="red"
		if vl_sub_total_entrada < 0 then s_cor="red"
		if (vl_sub_total_saida-vl_sub_total_entrada) < 0 then s_cor="red"
		x = x & "<TR NOWRAP style='background: #FFFFDD'>" & _
				"<TD COLSPAN='2' class='MTBE' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
				"TOTAL:</p></td>" & _
				"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_sub_total) & "</p></TD>" & _
				"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida) & "</p></td>" & _
				"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_entrada) & "</p></td>" & _
				"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida-vl_sub_total_entrada) & "</p></td>" & _
				"<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</p></td></TR>"
		
	'>	TOTAL GERAL
		if qtde_fabricantes > 1 then
			if (vl_total_saida-vl_total_entrada) = 0 then
				perc = 0
			else
				perc = ((vl_total_saida-vl_total_entrada)/(vl_total_saida-vl_total_entrada))*100
				end if
			s_cor="black"
			if qtde_total < 0 then s_cor="red"
			if vl_total_saida < 0 then s_cor="red"
			if vl_total_entrada < 0 then s_cor="red"
			if (vl_total_saida-vl_total_entrada) < 0 then s_cor="red"
			x = x & "<TR><TD COLSPAN='7' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
					"<TR><TD COLSPAN='7' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
					"<TR NOWRAP style='background:honeydew'>" & _
					"<TD class='MTBE' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL:</p></td>" & _
					"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(qtde_total) & "</p></TD>" & _
					"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_saida) & "</p></td>" & _
					"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_entrada) & "</p></td>" & _
					"<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_saida-vl_total_entrada) & "</p></td>" & _
					"<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc) & "%" & "</p></td></TR>"
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = ""
		if c_fabricante <> "" then
			s = c_fabricante
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			if s <> "" then x = x & cab_table & "<TR><TD class='MDTE' COLSPAN='7' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td></tr>" & chr(13) & cab
		else
			x = x & cab_table & cab
			end if

		x = x & "	<TR NOWRAP>" & _
				"		<TD class='MT' colspan='7'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & _
				"	</TR>"
		end if

  ' FECHA TABELA DO ÚLTIMO FABRICANTE
	x = x & "</TABLE>"
	
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
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">
<input type="hidden" name="c_grupo" id="c_grupo" value="<%=c_grupo%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="c_pedido" id="c_pedido" value="<%=c_pedido%>">
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Faturamento</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>"
	
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Período:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s_aux = ""
	if c_fabricante <> "" then s_aux = x_fabricante(c_fabricante)
	s = c_fabricante
	if (s <> "") And (s_aux <> "") then s = s & " - " & s_aux
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Fabricante:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"
	
	s = c_produto
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Produto:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"
	
	s = c_grupo
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Grupo de Produtos:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

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

    s = c_empresa
	if s = "" then 
		s = "todas"
	else
		s =  obtem_apelido_empresa_NFe_emitente(c_empresa)
    end if
        s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Empresa:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"    

	s = c_pedido
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Pedido:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & formata_data_hora(Now) & "</p></td></tr>"

	s_filtro = s_filtro & "</table>"
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
