<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L C O N T A G E M E S T O Q U E E X E C . A S P
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

	const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
	const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
	const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
	const MSO_NUMBER_FORMAT_TEXTO = "\@"

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_CONTAGEM_ESTOQUE, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_filtro, c_fabricante
	dim s_nome_fabricante
	dim rb_saida
    dim c_empresa
	dim c_grupo, c_subgrupo

	c_fabricante = retorna_so_digitos(Request.Form("c_fabricante"))
	if c_fabricante <> "" then c_fabricante = normaliza_codigo(c_fabricante, TAM_MIN_FABRICANTE)
	
	alerta = ""

	rb_saida = Ucase(Trim(Request.Form("rb_saida")))
    c_empresa = Trim(Request.Form("c_empresa"))

	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_subgrupo = Ucase(Trim(Request.Form("c_subgrupo")))

	if c_fabricante <> "" then
		s_nome_fabricante = fabricante_descricao(c_fabricante)
	else
		s_nome_fabricante = ""
		end if
	
	if alerta = "" then
		call set_default_valor_texto_bd(usuario, "RelContagemEstoque|c_fabricante", c_fabricante)
		call set_default_valor_texto_bd(usuario, "RelContagemEstoque|c_grupo", c_grupo)
		call set_default_valor_texto_bd(usuario, "RelContagemEstoque|c_subgrupo", c_subgrupo)
		call set_default_valor_texto_bd(usuario, "RelContagemEstoque|c_empresa", c_empresa)
		end if

	dim blnSaidaExcel
	blnSaidaExcel = False
	if alerta = "" then
		if rb_saida = "XLS" then
			blnSaidaExcel = True
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=ContagemEstoque_" & formata_data_yyyymmdd(Now) & "_" & formata_hora_hhnnss(Now) & ".xls"
			Response.Write "<h2>Relatório de Contagem de Estoque</h2>"
		'	FABRICANTE
			s = Trim(c_fabricante)
			if s = "" then
				s = "todos"
			else
				if s_nome_fabricante <> "" then s = s & " - " & s_nome_fabricante 
				end if
			Response.Write "Fabricante: " & s
			Response.Write "<br>"
		'	EMPRESA
			s = c_empresa
			if s = "" then 
				s = "todas"
			else
				s = obtem_apelido_empresa_NFe_emitente(c_empresa)
				end if
			Response.Write "Empresa: " & s
			Response.Write "<br>"
		'	GRUPO DE PRODUTOS
			s = c_grupo
			if s = "" then s = "N.I."
			Response.Write "Grupo de Produtos: " & s
			Response.Write "<br>"
		'	SUBGRUPO DE PRODUTOS
			s = c_subgrupo
			if s = "" then s = "N.I."
			Response.Write "Subgrupo de Produtos: " & s
			Response.Write "<br>"
		'	DATA EMISSÃO
			s = "Emissão: " & formata_data_hora(Now)
			Response.Write s
			Response.Write "<br><br>"
			
			contagem_estoque_detalhe_sintetico
			Response.End
			end if
		end if



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONTAGEM ESTOQUE DETALHE SINTETICO
'
sub contagem_estoque_detalhe_sintetico
const w_codigo = 70
const w_descricao = 315
const w_qtd_estoque_venda = 60
const w_qtd_split_possivel = 60
const w_qtd_a_separar = 60
const w_qtd_todos = 60
dim r
dim s, s_aux, s_bkg_color, s_nbsp, s_sql,s_sql_aux, x, cab_table, cab, fabricante_a, msg_erro
dim n, n_total_linha, n_reg, qtde_fabricantes
dim n_sub_total_estoque_venda, n_sub_total_split_possivel, n_sub_total_a_separar, n_sub_total_todos
dim n_total_estoque_venda, n_total_split_possivel, n_total_a_separar, n_total_todos
dim cont, v_grupos, v_subgrupos, s_where_grupo, s_where_subgrupo

'	SELECIONA TODOS OS PRODUTOS QUE POSSUEM ALGUM ITEM NA SITUAÇÃO DESEJADA
'	OBS: O USO DE 'UNION' SIMPLES ELIMINA AS LINHAS DUPLICADAS DOS RESULTADOS
'		 O USO DE 'UNION ALL' RETORNARIA TODAS AS LINHAS, INCLUSIVE AS DUPLICADAS
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.

'	GRUPO DE PRODUTOS
	s_where_grupo = ""
	if c_grupo <> "" then
		v_grupos = split(c_grupo, ", ")
		for cont = LBound(v_grupos) to UBound(v_grupos)
			if Trim(v_grupos(cont)) <> "" then
				if s_where_grupo <> "" then s_where_grupo = s_where_grupo & ", "
				s_where_grupo = s_where_grupo & "'" & Trim(v_grupos(cont)) & "'"
				end if
			next
		
		if s_where_grupo <> "" then
			s_where_grupo = " (t_PRODUTO.grupo IN (" & s_where_grupo & "))"
			end if
		end if

'	SUBGRUPO DE PRODUTOS
	s_where_subgrupo = ""
	if c_subgrupo <> "" then
		v_subgrupos = split(c_subgrupo, ", ")
		for cont = LBound(v_subgrupos) to UBound(v_subgrupos)
			if Trim(v_subgrupos(cont)) <> "" then
				if s_where_subgrupo <> "" then s_where_subgrupo = s_where_subgrupo & ", "
				s_where_subgrupo = s_where_subgrupo & "'" & Trim(v_subgrupos(cont)) & "'"
				end if
			next
		
		if s_where_subgrupo <> "" then
			s_where_subgrupo = " (t_PRODUTO.subgrupo IN (" & s_where_subgrupo & "))"
			end if
		end if

'	PRODUTOS NO ESTOQUE DE VENDA
	s_sql_aux = "SELECT DISTINCT t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto" & _
			" FROM t_ESTOQUE_ITEM" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque = t_ESTOQUE.id_estoque)"
	
	if (s_where_grupo <> "") Or (s_where_subgrupo <> "") then
		s_sql_aux = s_sql_aux & _
				" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))"
		end if

	s_sql_aux = s_sql_aux & _
			" WHERE ((qtde-qtde_utilizada) > 0)"

	if c_fabricante <> "" then
		s_sql_aux = s_sql_aux & " AND (t_ESTOQUE_ITEM.fabricante='" & c_fabricante & "')" 
		end if

    if c_empresa <> "" then
        s_sql_aux = s_sql_aux & " AND (t_ESTOQUE.id_nfe_emitente= " & c_empresa & ")"
		end if

	if s_where_grupo <> "" then
		s_sql_aux = s_sql_aux & " AND" & s_where_grupo
		end if

	if s_where_subgrupo <> "" then
		s_sql_aux = s_sql_aux & " AND" & s_where_subgrupo
		end if

'	PRODUTOS DO ESTOQUE DE PRODUTOS VENDIDOS (PEDIDOS EM STATUS SPLIT POSSÍVEL)
	s_sql_aux = s_sql_aux & _
			" UNION " & _
			"SELECT DISTINCT t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto" & _
			" FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)"

	if (s_where_grupo <> "") Or (s_where_subgrupo <> "") then
		s_sql_aux = s_sql_aux & _
				" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))"
		end if

	s_sql_aux = s_sql_aux & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque='" & ID_ESTOQUE_VENDIDO & "')" & _
			" AND (qtde > 0)" & _
			" AND (t_PEDIDO.st_entrega = '" & ST_ENTREGA_SPLIT_POSSIVEL & "')"

	if c_fabricante <> "" then
		s_sql_aux = s_sql_aux & " AND (t_ESTOQUE_MOVIMENTO.fabricante='" & c_fabricante & "')" 
		end if

    if c_empresa <> "" then
        s_sql_aux = s_sql_aux & " AND (t_ESTOQUE.id_nfe_emitente= " & c_empresa & ")"
		end if

	if s_where_grupo <> "" then
		s_sql_aux = s_sql_aux & " AND" & s_where_grupo
		end if

	if s_where_subgrupo <> "" then
		s_sql_aux = s_sql_aux & " AND" & s_where_subgrupo
		end if

'	PRODUTOS DO ESTOQUE DE PRODUTOS VENDIDOS (PEDIDOS EM STATUS 'A SEPARAR')
	s_sql_aux = s_sql_aux & _
			" UNION " & _
			"SELECT DISTINCT t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto" & _
			" FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)"

	if (s_where_grupo <> "") Or (s_where_subgrupo <> "") then
		s_sql_aux = s_sql_aux & _
				" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))"
		end if

	s_sql_aux = s_sql_aux & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque='" & ID_ESTOQUE_VENDIDO & "')" & _
			" AND (qtde > 0)" & _
			" AND (t_PEDIDO.st_entrega = '" & ST_ENTREGA_SEPARAR & "')"

	if c_fabricante <> "" then
		s_sql_aux = s_sql_aux & " AND (t_ESTOQUE_MOVIMENTO.fabricante='" & c_fabricante & "')" 
		end if

    if c_empresa <> "" then
        s_sql_aux = s_sql_aux & " AND (t_ESTOQUE.id_nfe_emitente= " & c_empresa & ")"
		end if

	if s_where_grupo <> "" then
		s_sql_aux = s_sql_aux & " AND" & s_where_grupo
		end if

	if s_where_subgrupo <> "" then
		s_sql_aux = s_sql_aux & " AND" & s_where_subgrupo
		end if

'	A PARTIR DA CONSULTA QUE OBTÉM TODA A RELAÇÃO DE PRODUTOS A SER LISTADA,
'	REALIZA A CONSULTA P/ CALCULAR AS QUANTIDADES
	s_sql = "SELECT" & _
				" tAuxBase.fabricante," & _
				" tAuxBase.produto," & _
				" tProd.descricao," & _
				" tProd.descricao_html," & _
				"(" & _
					"SELECT" & _
						" Sum(qtde-qtde_utilizada)" & _
					" FROM t_ESTOQUE_ITEM tEI" & _
                    " INNER JOIN t_ESTOQUE ON (tEI.id_estoque = t_ESTOQUE.id_estoque)"

	if (s_where_grupo <> "") Or (s_where_subgrupo <> "") then
		s_sql = s_sql & _
				" LEFT JOIN t_PRODUTO ON ((tEI.fabricante=t_PRODUTO.fabricante) AND (tEI.produto=t_PRODUTO.produto))"
		end if

	s_sql = s_sql & _
					" WHERE" & _
						" ((qtde-qtde_utilizada) > 0)" & _
						" AND (tEI.fabricante=tAuxBase.fabricante)" & _
						" AND (tEI.produto=tAuxBase.produto)" 

    if c_empresa <> "" then
        s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente= " & c_empresa & ")"
    end if

	if s_where_grupo <> "" then
		s_sql = s_sql & " AND" & s_where_grupo
		end if

	if s_where_subgrupo <> "" then
		s_sql = s_sql & " AND" & s_where_subgrupo
		end if

	s_sql = s_sql & ") AS qtde_estoque_venda," & _
				"(" & _
					"SELECT" & _
						" Sum(qtde)" & _
					" FROM t_ESTOQUE_MOVIMENTO tEM" & _
						" INNER JOIN t_PEDIDO tP ON (tEM.pedido=tP.pedido)" & _
                        " INNER JOIN t_ESTOQUE ON (tEM.id_estoque = t_ESTOQUE.id_estoque)"

	if (s_where_grupo <> "") Or (s_where_subgrupo <> "") then
		s_sql = s_sql & _
				" LEFT JOIN t_PRODUTO ON ((tEM.fabricante=t_PRODUTO.fabricante) AND (tEM.produto=t_PRODUTO.produto))"
		end if

	s_sql = s_sql & _
					" WHERE" & _
						" (anulado_status=0)" & _
						" AND (estoque='" & ID_ESTOQUE_VENDIDO & "')" & _
						" AND (qtde > 0)" & _
						" AND (tP.st_entrega = '" & ST_ENTREGA_SPLIT_POSSIVEL & "')" & _
						" AND (tEM.fabricante=tAuxBase.fabricante)" & _
						" AND (tEM.produto=tAuxBase.produto)" 

    if c_empresa <> "" then
        s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente= " & c_empresa & ")"
    end if

	if s_where_grupo <> "" then
		s_sql = s_sql & " AND" & s_where_grupo
		end if

	if s_where_subgrupo <> "" then
		s_sql = s_sql & " AND" & s_where_subgrupo
		end if

	s_sql = s_sql & ") AS qtde_split_possivel," & _
				"(" & _
					"SELECT" & _
						" Sum(qtde)" & _
					" FROM t_ESTOQUE_MOVIMENTO tEM" & _
						" INNER JOIN t_PEDIDO tP ON (tEM.pedido=tP.pedido)" & _
                        " INNER JOIN t_ESTOQUE ON (tEM.id_estoque = t_ESTOQUE.id_estoque)"

	if (s_where_grupo <> "") Or (s_where_subgrupo <> "") then
		s_sql = s_sql & _
				" LEFT JOIN t_PRODUTO ON ((tEM.fabricante=t_PRODUTO.fabricante) AND (tEM.produto=t_PRODUTO.produto))"
		end if

	s_sql = s_sql & _
					" WHERE" & _
						" (anulado_status=0)" & _
						" AND (estoque='" & ID_ESTOQUE_VENDIDO & "')" & _
						" AND (qtde > 0)" & _
						" AND (tP.st_entrega = '" & ST_ENTREGA_SEPARAR & "')" & _
						" AND (tEM.fabricante=tAuxBase.fabricante)" & _
						" AND (tEM.produto=tAuxBase.produto)" 

    if c_empresa <> "" then
        s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente= " & c_empresa & ")"
    end if

	if s_where_grupo <> "" then
		s_sql = s_sql & " AND" & s_where_grupo
		end if

	if s_where_subgrupo <> "" then
		s_sql = s_sql & " AND" & s_where_subgrupo
		end if

	s_sql = s_sql & ") AS qtde_a_separar" & _
			" FROM (" & s_sql_aux & ") tAuxBase" & _
				" LEFT JOIN t_PRODUTO tProd ON (tAuxBase.fabricante=tProd.fabricante) AND (tAuxBase.produto=tProd.produto)" & _               
			" ORDER BY" & _
				" tAuxBase.fabricante," & _
				" tAuxBase.produto"

	
  ' CABEÇALHO
	cab_table = "<TABLE class='Qt' cellSpacing=0>" & chr(13)
	if blnSaidaExcel then
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD valign='bottom' NOWRAP class='MDBE' style='width:" & Cstr(w_codigo) & "px'><P class='R' style='font-weight:bold;'>Código</P></TD>" & chr(13) & _
			  "		<TD valign='bottom' NOWRAP class='MDB' style='width:" & Cstr(w_descricao) & "px'><P class='R' style='font-weight:bold;'>Produto</P></TD>" & chr(13) & _
			  "		<TD align='right' valign='bottom' class='MDB' style='width:" & Cstr(w_qtd_estoque_venda) & "px'><P class='Rd' style='font-weight:bold;'>Estoque<br style='mso-data-placement:same-cell;' />Venda</P></TD>" & chr(13) & _
			  "		<TD align='right' valign='bottom' class='MDB' style='width:" & Cstr(w_qtd_split_possivel) & "px'><P class='Rd' style='font-weight:bold;'>Split<br style='mso-data-placement:same-cell;' />Possível</P></TD>" & chr(13) & _
			  "		<TD align='right' valign='bottom' class='MDB' style='width:" & Cstr(w_qtd_a_separar) & "px'><P class='Rd' style='font-weight:bold;'>A Separar</P></TD>" & chr(13) & _
			  "		<TD align='right' valign='bottom' class='MDB' style='width:" & Cstr(w_qtd_todos) & "px'><P class='Rd' style='font-weight:bold;'>Total</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
	else
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD valign='bottom' NOWRAP class='MDBE' style='width:" & Cstr(w_codigo) & "px'><P class='R' style='font-weight:bold;'>Código</P></TD>" & chr(13) & _
			  "		<TD valign='bottom' NOWRAP class='MDB' style='width:" & Cstr(w_descricao) & "px'><P class='R' style='font-weight:bold;'>Produto</P></TD>" & chr(13) & _
			  "		<TD align='right' valign='bottom' NOWRAP class='MDB' style='width:" & Cstr(w_qtd_estoque_venda) & "px'><P class='Rd' style='font-weight:bold;'>Estoque Venda</P></TD>" & chr(13) & _
			  "		<TD align='right' valign='bottom' NOWRAP class='MDB' style='width:" & Cstr(w_qtd_split_possivel) & "px'><P class='Rd' style='font-weight:bold;'>Split Possível</P></TD>" & chr(13) & _
			  "		<TD align='right' valign='bottom' NOWRAP class='MDB' style='width:" & Cstr(w_qtd_a_separar) & "px'><P class='Rd' style='font-weight:bold;'>A Separar</P></TD>" & chr(13) & _
			  "		<TD align='right' valign='bottom' NOWRAP class='MDB' style='width:" & Cstr(w_qtd_todos) & "px'><P class='Rd' style='font-weight:bold;'>Total</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
		end if
	
	x = ""
	n_reg = 0
	qtde_fabricantes = 0
	fabricante_a = "XXXXX"

	n_sub_total_estoque_venda = 0
	n_sub_total_split_possivel = 0
	n_sub_total_a_separar = 0
	n_sub_total_todos = 0
	n_total_estoque_venda = 0
	n_total_split_possivel = 0
	n_total_a_separar = 0
	n_total_todos = 0
	
	if blnSaidaExcel then s_nbsp = "" else s_nbsp = "&nbsp;"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then
				x = x & "	<TR style='background: #FFFFDD' NOWRAP>" & chr(13) & _
						"		<TD colspan='2' class='MEB' align='right'><P class='Cd' style='font-weight:bold;'>Total:</P></TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_sub_total_estoque_venda) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_sub_total_split_possivel) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_sub_total_a_separar) & "</P></TD>" & chr(13) & _
						"		<TD class='MDB'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_sub_total_todos) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x = "<BR>" & chr(13) & _
					"<BR>" & chr(13)
				end if

			x = x & cab_table
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			if blnSaidaExcel then s_bkg_color = "tomato" else s_bkg_color = "azure"
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MDBE' align='center' colspan='6' style='background:" & s_bkg_color & ";' style='font-weight:bold;'><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13) & _
					cab
			n_sub_total_estoque_venda = 0
			n_sub_total_split_possivel = 0
			n_sub_total_a_separar = 0
			n_sub_total_todos = 0
			end if
		
	  ' CONTAGEM
		n_reg = n_reg + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> CÓDIGO DO PRODUTO
		x = x & "		<TD class='MDBE' valign='middle' style='width:" & Cstr(w_codigo) & "px;' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	 '> DESCRIÇÃO DO PRODUTO
		x = x & "		<TD class='MDB' valign='middle' style='width:" & Cstr(w_descricao) & "px' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

		n_total_linha = 0
		
	 '> ESTOQUE DE VENDA
		n = 0 
		if Not IsNull(r("qtde_estoque_venda")) then n = r("qtde_estoque_venda")
		n_sub_total_estoque_venda = n_sub_total_estoque_venda + n
		n_total_estoque_venda = n_total_estoque_venda + n
		n_total_linha = n_total_linha + n
		x = x & "		<TD class='MDB' valign='middle' style='width:" & Cstr(w_qtd_estoque_venda) & "px' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n) & "</P></TD>" & chr(13)

	 '> SPLIT POSSÍVEL
		n = 0 
		if Not IsNull(r("qtde_split_possivel")) then n = r("qtde_split_possivel")
		n_sub_total_split_possivel = n_sub_total_split_possivel + n
		n_total_split_possivel = n_total_split_possivel + n
		n_total_linha = n_total_linha + n
		x = x & "		<TD class='MDB' valign='middle' style='width:" & Cstr(w_qtd_split_possivel) & "px' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n) & "</P></TD>" & chr(13)

	 '> A SEPARAR
		n = 0 
		if Not IsNull(r("qtde_a_separar")) then n = r("qtde_a_separar")
		n_sub_total_a_separar = n_sub_total_a_separar + n
		n_total_a_separar = n_total_a_separar + n
		n_total_linha = n_total_linha + n
		x = x & "		<TD class='MDB' valign='middle' style='width:" & Cstr(w_qtd_a_separar) & "px' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n) & "</P></TD>" & chr(13)

		n_sub_total_todos = n_sub_total_todos + n_total_linha
		n_total_todos = n_total_todos + n_total_linha

	 '> TOTAL
		x = x & "		<TD class='MDB' valign='middle' style='width:" & Cstr(w_qtd_todos) & "px;' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_total_linha) & "</P></TD>" & chr(13)
		
		x = x & "	</TR>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA TOTAL DO ÚLTIMO FABRICANTE
	if n_reg <> 0 then 
		x = x & "	<TR style='background: #FFFFDD' NOWRAP>"  & chr(13) & _
				"		<TD class='MEB' COLSPAN='2' align='right' NOWRAP><P class='Cd' style='font-weight:bold;'>" & "Total:" & "</P></TD>" & chr(13) & _
				"		<TD class='MB' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_sub_total_estoque_venda) & "</P></TD>" & chr(13) & _
				"		<TD class='MB' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_sub_total_split_possivel) & "</P></TD>" & chr(13) & _
				"		<TD class='MB' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_sub_total_a_separar) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_sub_total_todos) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
	'>	TOTAL GERAL
		if qtde_fabricantes > 1 then
			x = x & _
				"<TR><TD COLSPAN='6'>&nbsp;</TD></TR>" & chr(13) & _
				"<TR><TD COLSPAN='6' class='MB'>&nbsp;</TD></TR>" & chr(13) & _
				"	<TR style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' style='width:" & Cstr(w_codigo) & "px;' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MB' align='right' style='width:" & Cstr(w_descricao) & "px' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;'>TOTAL GERAL:</p></TD>" & chr(13) & _
				"		<TD class='MB' style='width:" & Cstr(w_qtd_estoque_venda) & "px' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_total_estoque_venda) & "</p></TD>" & chr(13) & _
				"		<TD class='MB' style='width:" & Cstr(w_qtd_split_possivel) & "px' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_total_split_possivel) & "</p></TD>" & chr(13) & _
				"		<TD class='MB' style='width:" & Cstr(w_qtd_a_separar) & "px' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_total_a_separar) & "</p></TD>" & chr(13) & _
				"		<TD class='MDB' style='width:" & Cstr(w_qtd_todos) & "px;' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_total_todos) & "</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='6' class='MDBE' align='center'><P class='ALERTA'>Nenhum produto do estoque satisfaz as condições especificadas</P></TD>" & chr(13) & _
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
P.C { font-size:10pt; }
P.Cc { font-size:10pt; }
P.Cd { font-size:10pt; }
P.F { font-size:11pt; }
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

<form id="fESTOQ" name="fESTOQ" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Contagem de Estoque</span>
	<br>
	<%	s = "<span class='N'>Emissão:&nbsp;" & formata_data_hora(Now) & "</span>"
		Response.Write s
	%>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='4' style='border-bottom:1px solid black' border='0'>" & chr(13)
	s = Trim(c_fabricante)
	if s = "" then
		s = "todos"
	else
		if s_nome_fabricante <> "" then s = s & " - " & s_nome_fabricante 
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Fabricante:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

     s = c_empresa
	if s = "" then 
		s = "todas"
	else
		s = obtem_apelido_empresa_NFe_emitente(c_empresa)
		end if
	 s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				       "<span class='N'>Empresa:&nbsp;</span></td><td align='left' valign='top'>" & _
				       "<span class='N'>" & s & "</span></td></tr>"

'	GRUPO DE PRODUTOS
	s = c_grupo
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Grupo de Produtos:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	SUBGRUPO DE PRODUTOS
	s = c_subgrupo
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Subgrupo de Produtos:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

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
<%
	contagem_estoque_detalhe_sintetico
%>

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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
