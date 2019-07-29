<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelComissaoTabelaProgressivaAnaliticoExec.asp
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

	dim dtInicioProcessamento, dtTerminoProcessamento, dtDuracaoProcessamento
	dtInicioProcessamento = Now
	
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
	if Not operacao_permitida(OP_LJA_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_ANALITICO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro
	dim c_dt_inicio, c_dt_termino, c_vendedor

	alerta = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))

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

	if alerta = "" then
		if c_vendedor <> "" then
			if (Instr(1, c_vendedor, BD_CURINGA_TODOS) = 0) And (Instr(1, c_vendedor, "*") = 0) then 
				s = "SELECT * FROM t_USUARIO WHERE usuario = '" & c_vendedor & "'" & SQL_COLLATE_CASE_ACCENT
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then alerta = "VENDEDOR '" & c_vendedor & "' NÃO ESTÁ CADASTRADO."
				end if
			end if
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
dim perc_comissao, perc_desconto
dim vl_comissao, vl_saida, vl_lista, vl_total_saida, vl_total_comissao, vl_sub_total_saida, vl_sub_total_comissao
dim s_where, s_where_venda, s_where_devolucao, s_where_perdas, s_cor, s_sinal, s_cor_sinal, s_vl_comissao
dim s_perc_desconto, s_perc_comissao
dim vTabela, intUltPosicaoTabela, intQtdeTabelas, rsTabela, rsPerfil, intCounter, blnOpPerda

'	CRITÉRIOS COMUNS
	s_where = " (CONVERT(smallint, t_PEDIDO.loja) = " & loja & ")"
	if c_vendedor <> "" then
		s = substitui_caracteres(c_vendedor, "*", BD_CURINGA_TODOS)
		s_aux = "="
		if Instr(1, s, BD_CURINGA_TODOS) <> 0 then s_aux = "LIKE"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.vendedor " & s_aux & " '" & s & "'" & SQL_COLLATE_CASE_ACCENT & ")"
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
	s_sql = "SELECT" & _
				" 'VENDA' AS operacao," & _
				" t_PEDIDO.vendedor," & _
				" t_PEDIDO.loja AS loja," & _
				" CONVERT(smallint,t_PEDIDO.loja) AS numero_loja," & _
				" t_PEDIDO.entregue_data AS data," & _
				" t_PEDIDO.pedido AS pedido," & _
				" t_PEDIDO_ITEM.fabricante AS fabricante," & _
				" t_PEDIDO_ITEM.produto AS produto," & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda) AS valor_saida," & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_lista) AS valor_lista" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
				" INNER JOIN t_PEDIDO_ITEM ON ((t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto))" & _
			" WHERE" & _
				" (anulado_status=0)" & _
				" AND (estoque='" & ID_ESTOQUE_ENTREGUE & "')" & _
			s & _
			" GROUP BY" & _
				" t_PEDIDO.vendedor," & _
				" t_PEDIDO.loja," & _
				" t_PEDIDO.entregue_data," & _
				" t_PEDIDO.pedido," & _
				" t_PEDIDO_ITEM.fabricante," & _
				" t_PEDIDO_ITEM.produto"

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
			"SELECT" & _
				" 'DEVOLUCAO' AS operacao," & _
				" t_PEDIDO.vendedor," & _
				" t_PEDIDO.loja AS loja," & _
				" CONVERT(smallint,t_PEDIDO.loja) AS numero_loja," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data," & _
				" t_PEDIDO.pedido AS pedido," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.fabricante AS fabricante," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.produto AS produto," & _
				" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS valor_saida," & _
				" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_lista) AS valor_lista" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
				" INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)" & _
				" INNER JOIN t_ESTOQUE ON (t_PEDIDO_ITEM_DEVOLVIDO.id=t_ESTOQUE.devolucao_id_item_devolvido)" & _
				" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_ESTOQUE_ITEM.produto))" & _ 
			s & _
			" GROUP BY" & _
				" t_PEDIDO.vendedor," & _
				" t_PEDIDO.loja," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data," & _
				" t_PEDIDO.pedido," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.fabricante," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.produto"

'	PERDAS
	s = s_where
	if (s <> "") And (s_where_perdas <> "") then s = s & " AND"
	s = s & s_where_perdas
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT" & _
				" 'PERDA' AS operacao," & _
				" t_PEDIDO.vendedor," & _
				" t_PEDIDO.loja AS loja," & _
				" CONVERT(smallint,t_PEDIDO.loja) AS numero_loja," & _
				" t_PEDIDO_PERDA.data AS data," & _
				" t_PEDIDO.pedido AS pedido," & _
				" '' AS fabricante," & _
				" '' AS produto," & _
				" Sum(-t_PEDIDO_PERDA.valor) AS valor_saida," & _
				" Sum(-t_PEDIDO_PERDA.valor) AS valor_lista" & _
			" FROM t_PEDIDO_PERDA" & _
				" INNER JOIN t_PEDIDO ON (t_PEDIDO_PERDA.pedido=t_PEDIDO.pedido)" & _
			s & _
			" GROUP BY" & _
				" t_PEDIDO.vendedor," & _
				" t_PEDIDO.loja," & _
				" t_PEDIDO_PERDA.data," & _
				" t_PEDIDO.pedido"
	
	s_sql = s_sql & _
			" ORDER BY" & _
				" vendedor," & _
				" numero_loja," & _
				" data," & _
				" pedido," & _
				" fabricante," & _
				" produto," & _
				" valor_saida DESC"

  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE' valign='bottom' NOWRAP><P style='width:35px' class='Rc'>Loja</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:75px' class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='center' valign='bottom' NOWRAP><P style='width:70px' class='Rc'>Data</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:35px' class='R'>Fabr</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:55px' class='R'>Produto</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:80px' class='Rd' style='font-weight:bold;'>Valor<br>(" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:30px' class='Rd' style='font-weight:bold;'>Desc<br>(%)</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:35px' class='Rd' style='font-weight:bold;'>Comis<br>(%)</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:80px' class='Rd' style='font-weight:bold;'>Comissão<br>(" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD' align='center' valign='bottom'><P style='width:20px' class='Rc' style='font-weight:bold;'>+/-</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_vendedores = 0
	vl_total_saida = 0
	vl_total_comissao = 0
	vl_sub_total_saida = 0
	vl_sub_total_comissao = 0
	intQtdeTabelas = 0

	vendedor_a = "XXXXXXXXXXXX"
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE VENDEDOR?
		if Trim("" & r("vendedor"))<>vendedor_a then
			vendedor_a = Trim("" & r("vendedor"))
			qtde_vendedores = qtde_vendedores + 1
		  ' FECHA TABELA DO VENDEDOR ANTERIOR
			if n_reg_total > 0 then 
				s_cor="black"
				if vl_sub_total_saida < 0 then s_cor="red"
				if vl_sub_total_comissao < 0 then s_cor="red"
				if intQtdeTabelas = 1 then
					x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
							"		<TD class='MTBE' COLSPAN='5' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
							"TOTAL:</p></td>" & chr(13) & _
							"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida) & "</p></TD>" & chr(13) & _
							"		<TD class='MTB' colspan='2'><p class='Cd' style='color:" & s_cor & ";'>&nbsp;</p></td>" & chr(13) & _
							"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_comissao) & "</p></td>" & chr(13) & _
							"		<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>&nbsp;</p></td>" & chr(13) & _
							"	</TR>" & chr(13)
					end if
				
				x = x & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			n_reg = 0
			vl_sub_total_saida = 0
			vl_sub_total_comissao = 0

		'	OBTÉM TABELA DE COMISSÃO
			intQtdeTabelas = 0
			s_sql = "SELECT " & _
						"*" & _
					" FROM t_PERFIL_X_USUARIO" & _
					" WHERE" & _
						" (usuario = '" & Trim("" & r("vendedor")) & "')" & _
					" ORDER BY" & _
						" id_perfil"
			set rsPerfil = cn.Execute(s_sql)
			do while Not rsPerfil.Eof
				s_sql = "SELECT " & _
							"*" & _
						" FROM t_PERCENTUAL_COMISSAO_VENDEDOR" & _
						" WHERE" & _
							" (id_perfil = '" & Trim("" & rsPerfil("id_perfil")) & "')" & _
						" ORDER BY" & _
							" perc_desconto"
				set rsTabela = cn.Execute(s_sql)
			'	A tabela vai de 0 a intUltPosicaoTabela
				if Not rsTabela.Eof then 
					intUltPosicaoTabela = -1
					redim vTabela(0)
					intQtdeTabelas = intQtdeTabelas + 1
					end if
				
				do while Not rsTabela.Eof
					intUltPosicaoTabela = intUltPosicaoTabela + 1
					redim preserve vTabela(intUltPosicaoTabela)
					set vTabela(intUltPosicaoTabela) = new cl_DUAS_COLUNAS
					vTabela(intUltPosicaoTabela).c1 = rsTabela("perc_desconto")
					vTabela(intUltPosicaoTabela).c2 = rsTabela("perc_comissao")
					rsTabela.MoveNext
					loop
					
				rsPerfil.MoveNext
				loop

			if n_reg_total > 0 then x = x & "<BR>"
			s = Trim("" & r("vendedor"))
			s_aux = x_usuario(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then x = x & "	<TR>" & chr(13) & _
									"		<TD class='MDTE' COLSPAN='10' valign='bottom' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
									"	</tr>" & chr(13)
			x = x & cab
			
			if intQtdeTabelas = 0 then
				x = x & "	<TR>" & chr(13) & _
						"		<TD class='MDTE MB' COLSPAN='10' align='center' valign='bottom' style='background:red;'><p class='N' style='color:yellow'>&nbsp;Nenhuma tabela de comissão associada ao vendedor&nbsp;</p></td>" & chr(13) & _
						"	</tr>" & chr(13)
			elseif intQtdeTabelas > 1 then
				x = x & "	<TR>" & chr(13) & _
						"		<TD class='MDTE MB' COLSPAN='10' align='center' valign='bottom' style='background:red;'><p class='N' style='color:yellow'>&nbsp;Há " & Cstr(intQtdeTabelas) & " tabelas de comissão associadas ao vendedor&nbsp;</p></td>" & chr(13) & _
						"	</tr>" & chr(13)
				end if
			end if


	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		if intQtdeTabelas = 1 then
			x = x & "	<TR NOWRAP>"  & chr(13)

		'	EVITA DIFERENÇAS DE ARREDONDAMENTO
			vl_saida = converte_numero(formata_moeda(r("valor_saida")))
			vl_lista = converte_numero(formata_moeda(r("valor_lista")))
	
		'	CALCULA COMISSÃO
			if vl_lista = 0 then
				perc_desconto = 0
			else
				perc_desconto = 100*((vl_lista - vl_saida) / vl_lista)
				end if
				
			blnOpPerda = False
			if Trim("" & r("operacao")) = "PERDA" then blnOpPerda = True
			
			if blnOpPerda then
				perc_comissao = 0
				vl_comissao = 0
			else
				perc_comissao = 0
				for intCounter = 0 to intUltPosicaoTabela
					if converte_numero(perc_desconto) <= converte_numero(vTabela(intCounter).c1) then
						perc_comissao = converte_numero(vTabela(intCounter).c2)
						exit for
						end if
					next
				vl_comissao = converte_numero(formata_moeda((perc_comissao/100)*vl_saida))
				end if
			
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
			x = x & "		<TD class='MDTE'><P class='Cnc' style='color:" & s_cor & ";'>" & Trim("" & r("loja")) & "</P></TD>" & chr(13)

		 '> Nº PEDIDO
			x = x & "		<TD class='MTD'><P class='Cn'><a style='color:" & s_cor & ";' href='javascript:fRELConcluir(" & _
					chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & _
					Trim("" & r("pedido")) & "</a></P></TD>" & chr(13)

		 '> DATA DO PEDIDO
			s = formata_data(r("data"))
			x = x & "		<TD align='center' class='MTD'><P class='Cnc' style='color:" & s_cor & ";'>" & s & "</P></TD>" & chr(13)

		 '> FABRICANTE
			s = Trim("" & r("fabricante"))
			if s = "" then s = "&nbsp;"
			x = x & "		<TD class='MTD'><P class='Cn' style='color:" & s_cor & ";'>" & s & "</P></TD>" & chr(13)

		 '> PRODUTO
			s = Trim("" & r("produto"))
			if s = "" then s = "&nbsp;"
			x = x & "		<TD class='MTD'><P class='Cn' style='color:" & s_cor & ";'>" & s & "</P></TD>" & chr(13)
		 
		 '> VALOR DO ITEM
			x = x & "		<TD align='right' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_saida) & "</P></TD>" & chr(13)

		'> PERCENTUAL DO DESCONTO
			if blnOpPerda then
				s_perc_desconto = "&nbsp;"
			else
				s_perc_desconto = formata_perc_desc(perc_desconto)
				end if
			x = x & "		<TD align='right' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & s_perc_desconto & "</P></TD>" & chr(13)
			
		'> PERCENTUAL DA COMISSÃO
			if blnOpPerda then
				s_perc_comissao = "&nbsp;"
			else
				s_perc_comissao = formata_perc_desc(perc_comissao)
				end if
			x = x & "		<TD align='right' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & s_perc_comissao & "</P></TD>" & chr(13)

		 '> VALOR DA COMISSÃO
			if blnOpPerda then 
				s_vl_comissao = "&nbsp;"
			else
				s_vl_comissao = formata_moeda(vl_comissao)
				end if
			x = x & "		<TD align='right' class='MTD'><P class='Cnd' style='color:" & s_cor & ";'>" & s_vl_comissao & "</P></TD>" & chr(13)

		 '> +/-
			x = x & "		<TD align='center' class='MTD'><P class='C' style='font-family:Courier,Arial;color:" & s_cor_sinal & "'>" & s_sinal & "</P></TD>" & chr(13)

			vl_sub_total_saida = vl_sub_total_saida + r("valor_saida")
			vl_total_saida = vl_total_saida + r("valor_saida")
			vl_sub_total_comissao = vl_sub_total_comissao + vl_comissao
			vl_total_comissao = vl_total_comissao + vl_comissao
			
			x = x & "	</TR>" & chr(13)
			end if
			
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
		if intQtdeTabelas = 1 then
			x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='5' class='MTBE' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL:</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida) & "</p></td>" & chr(13) & _
					"		<TD class='MTB' colspan='2'><p class='Cd' style='color:" & s_cor & ";'>&nbsp;</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_comissao) & "</p></td>" & chr(13) & _
					"		<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>&nbsp;</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		
	'>	TOTAL GERAL
		if qtde_vendedores > 1 then
			s_cor="black"
			if vl_total_saida < 0 then s_cor="red"
			if vl_total_comissao < 0 then s_cor="red"
			x = x & "	<TR>" & chr(13) & _
					"		<TD COLSPAN='10' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD COLSPAN='10' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
					"		<TD class='MTBE' COLSPAN='5' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL GERAL:</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_saida) & "</p></td>" & chr(13) & _
					"		<TD class='MTB' colspan='2'><p class='Cd' style='color:" & s_cor & ";'>&nbsp;</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_comissao) & "</p></td>" & chr(13) & _
					"		<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>&nbsp;</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='10'><P class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS NO PERÍODO ESPECIFICADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO VENDEDOR
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
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Comissão aos Vendedores Analítico (Tabela Progressiva)</span>
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
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Período:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & s & "</p></td></tr>" & chr(13)

	s = c_vendedor
	if s = "" then 
		s = "todos"
	elseif (Instr(1,s,"*")=0) And (Instr(1,s,BD_CURINGA_TODOS)=0) then
		s_aux = x_usuario(c_vendedor)
		if s_aux <> "" then s = s & " (" & s_aux & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Vendedor:&nbsp;</p></td><td valign='top'>" & _
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

<%	dtTerminoProcessamento = Now 
	dtDuracaoProcessamento = dtTerminoProcessamento - dtInicioProcessamento
%>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" class="notPrint">
<tr>
	<td class="Rc" style="border-bottom:1px solid black">&nbsp;</td>
</tr>
<tr>
	<td class="Rd">Processado no servidor em: <%=formata_duracao_hms(dtDuracaoProcessamento)%></td>
</tr>
</table>

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
