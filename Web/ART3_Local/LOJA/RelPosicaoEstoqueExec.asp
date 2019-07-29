<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L P O S I C A O E S T O Q U E E X E C . A S P
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
	if Not ( _
			operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_SINTETICO, s_lista_operacoes_permitidas) Or _
			operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO, s_lista_operacoes_permitidas) _
			) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, c_fabricante, c_produto, rb_estoque, rb_detalhe
	dim cod_fabricante, cod_produto
	dim s_nome_fabricante, s_nome_produto, s_nome_produto_html

	c_fabricante = retorna_so_digitos(Request.Form("c_fabricante"))
	if c_fabricante <> "" then c_fabricante = normaliza_codigo(c_fabricante, TAM_MIN_FABRICANTE)
	c_produto = UCase(Trim(Request.Form("c_produto")))
	rb_estoque = Trim(Request.Form("rb_estoque"))
	rb_detalhe = Trim(Request.Form("rb_detalhe"))
	
	alerta = ""
	if (c_produto<>"") And (Not IsEAN(c_produto)) then
		if c_fabricante = "" then alerta = "PARA CONSULTAR PELO CÓDIGO INTERNO DE PRODUTO É NECESSÁRIO ESPECIFICAR O FABRICANTE."
		end if
		
	if alerta = "" then
	'	DEFAULT
		cod_produto = c_produto
		cod_fabricante = c_fabricante
		
		if IsEAN(c_produto) then
			s = "SELECT fabricante, produto, ean FROM t_PRODUTO WHERE (ean='" & c_produto & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta = "Produto com código EAN " & c_produto & " não está cadastrado."
			else
				if c_fabricante <> "" then
					if c_fabricante <> Trim("" & rs("fabricante")) then 
						alerta = "Produto " & Trim("" & rs("produto")) & " (EAN: " & _
								 Trim("" & rs("ean")) & ") não pertence ao fabricante " & c_fabricante & "."
						end if
					end if
				
				if alerta = "" then
				'	OBTÉM O CÓDIGO INTERNO DE PRODUTO
					cod_fabricante = Trim("" & rs("fabricante"))
					cod_produto = Trim("" & rs("produto"))
					end if
				end if
			end if
		end if

	if alerta = "" then
		if cod_fabricante <> "" then
			s_nome_fabricante = fabricante_descricao(cod_fabricante)
		else
			s_nome_fabricante = ""
			end if
				
		if cod_produto <> "" then
			s_nome_produto = produto_descricao(cod_fabricante, cod_produto)
			s_nome_produto_html = produto_formata_descricao_em_html(produto_descricao_html(cod_fabricante, cod_produto))
		else
			s_nome_produto = ""
			s_nome_produto_html = ""
			end if
		end if






' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA ESTOQUE DETALHE SINTETICO
'
sub consulta_estoque_detalhe_sintetico
dim r
dim s, s_aux, s_sql, loja_a, x, cab_table, cab, fabricante_a
dim n_reg, n_reg_total, n_saldo_parcial, n_saldo_total, qtde_lojas, qtde_fabricantes

'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT loja, CONVERT(smallint,loja) AS numero_loja" & _
			", t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto" & _
			", descricao" & _
			", descricao_html" & _
			", Sum(qtde) AS saldo" & _
			" FROM t_ESTOQUE_MOVIMENTO LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque='" & rb_estoque & "')"
	
	if cod_fabricante <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.fabricante='" & cod_fabricante & "')" 
		end if

	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.produto='" & cod_produto & "')"
		end if
	
	s_sql = s_sql & " GROUP BY loja, t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, descricao, descricao_html" & _
					" ORDER BY numero_loja, t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, descricao, descricao_html"

  ' CABEÇALHO
	cab_table = "<TABLE class='Q' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' nowrap>" & chr(13) & _
		  "		<TD width='75' valign='bottom' nowrap class='MD MB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='480' valign='bottom' nowrap class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
		  "		<TD width='60' valign='bottom' nowrap class='MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	n_saldo_parcial = 0
	n_saldo_total = 0
	qtde_lojas = 0
	qtde_fabricantes = 0
	loja_a = "XXX"
	fabricante_a = "XXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	'	MUDOU LOJA?
		if Trim("" & r("loja")) <> loja_a then
			if n_reg_total > 0 then 
			  ' FECHA TABELA DA LOJA ANTERIOR
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD COLSPAN='2' NOWRAP><P class='Cd'>" & "Total:" & "</P></TD>" & chr(13) & _
						"		<TD NOWRAP><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			x = x & cab_table

		  ' INICIA NOVA TABELA P/ A NOVA LOJA
			if Trim("" & r("loja")) <> "" then
			'	QUEBRA POR LOJA APENAS SE HOUVER LOJA
				x = x & "	<TR NOWRAP style='background:azure'>" & chr(13) & _
						"		<TD class='MB' align='center' colspan='3'><P class='F'>" & r("loja") & " - " & ucase(x_loja(r("loja"))) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				end if

			x = x & cab

			n_reg = 0
			n_saldo_parcial = 0
			loja_a = Trim("" & r("loja"))
			qtde_lojas = qtde_lojas + 1
			fabricante_a = "XXXXX"
			end if
		
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MB' colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='3' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MB' align='center' colspan='3' style='background: honeydew'><P class='C'>&nbsp;" & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
			end if
			
	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='C'>&nbsp;" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='bottom'><P class='C' NOWRAP>&nbsp;" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

 	 '> SALDO
		x = x & "		<TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

		n_saldo_parcial = n_saldo_parcial + r("saldo")
		n_saldo_total = n_saldo_total + r("saldo")
		
		x = x & "	</TR>" & chr(13)
		
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP>"  & chr(13) & _
				"		<TD COLSPAN='2' NOWRAP><P class='Cd'>" & "Total:" & "</P></TD>" & chr(13) & _
				"		<TD NOWRAP><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
	'>	TOTAL GERAL
		if (qtde_lojas > 1) Or (qtde_fabricantes > 1) then
			x = x & "</TABLE>" & chr(13) & _
				"<BR>" & chr(13) & _
				"<BR>" & chr(13) & _
				cab_table & _
				"	<TR style='background:honeydew'>" & chr(13) & _
				"		<TD width='75' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD width='480' valign='bottom' NOWRAP><p class='Cd'>TOTAL GERAL:</p></TD>" & chr(13) & _
				"		<TD width='60' valign='bottom' NOWRAP><p class='Cd'>" & formata_inteiro(n_saldo_total) & "</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='3'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DA ÚLTIMA LOJA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing

end sub




' ____________________________________________
' CONSULTA ESTOQUE DETALHE INTERMEDIARIO
'
sub consulta_estoque_detalhe_intermediario
const LargColOrdemServico = 110
dim LargColDescricao
dim r
dim s, s_aux, s_sql, s_lista_OS, s_chave_OS, s_num_OS_tela, loja_a, x, cab_table, cab, fabricante_a
dim n_reg, n_reg_total, n_saldo_parcial, n_saldo_total, qtde_lojas, qtde_fabricantes
dim vl, vl_sub_total, vl_total_geral
dim intQtdeColunasColSpanSubTotal, intQtdeColunasColSpanTotalGeral, intQtdeTotalColunasColSpan

'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT loja, CONVERT(smallint,loja) AS numero_loja" & _
			", t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto" & _
			", descricao" & _
			", descricao_html" & _
			", Sum(t_ESTOQUE_MOVIMENTO.qtde) AS saldo" & _
			", Sum(t_ESTOQUE_MOVIMENTO.qtde*t_ESTOQUE_ITEM.vl_custo2) AS preco_total" & _
			" FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque='" & rb_estoque & "')"

	if cod_fabricante <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.fabricante='" & cod_fabricante & "')" 
		end if

	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.produto='" & cod_produto & "')"
		end if

	s_sql = s_sql & " GROUP BY loja, t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, descricao, descricao_html" & _
					" ORDER BY numero_loja, t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, descricao, descricao_html"

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD width='60' valign='bottom' NOWRAP class='MDBE'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='" & CStr(LargColDescricao) & "' valign='bottom' NOWRAP class='MDB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13)

	if rb_estoque = ID_ESTOQUE_DANIFICADOS then
		cab = cab & _
			  "		<TD width='" & CStr(LargColOrdemServico) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>ORDEM SERVIÇO</P></TD>" & chr(13)
		end if
		  
	cab = cab & _
		  "		<TD width='60' valign='bottom' NOWRAP class='MDB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD width='100' valign='bottom' NOWRAP class='MDB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
		  "		<TD width='100' valign='bottom' NOWRAP class='MDB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)

	intQtdeTotalColunasColSpan = 5
	intQtdeColunasColSpanSubTotal = 2
	intQtdeColunasColSpanTotalGeral = 1
	LargColDescricao = 270
	if rb_estoque = ID_ESTOQUE_DANIFICADOS then
		intQtdeTotalColunasColSpan = 6
		intQtdeColunasColSpanSubTotal = 3
		intQtdeColunasColSpanTotalGeral = 2
		LargColDescricao = 200
		end if
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	n_saldo_parcial = 0
	n_saldo_total = 0
	vl_sub_total = 0
	vl_total_geral = 0
	qtde_lojas = 0
	loja_a = "XXX"
	fabricante_a = "XXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU LOJA?
		if Trim("" & r("loja")) <> loja_a then
			if n_reg_total > 0 then 
			  ' FECHA TABELA DA LOJA ANTERIOR
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD COLSPAN='" & Cstr(intQtdeColunasColSpanSubTotal) & "' class='MEB' NOWRAP><P class='Cd'>" & _
						"Total:</P></TD>" & chr(13) & _
						"		<TD NOWRAP class='MB'><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & _
						"</P></TD>" & chr(13) & _
						"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						"		<TD NOWRAP class='MDB'><P class='Cd'>" & _
						formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			x = x & cab_table

		  ' INICIA NOVA TABELA P/ A NOVA LOJA
			if Trim("" & r("loja")) <> "" then
			'   QUEBRA POR LOJA APENAS SE HOUVER LOJA
				x = x & "	<TR NOWRAP style='background:azure'>" & chr(13) & _
						"		<TD class='MDBE' align='center' colspan='" & CStr(intQtdeTotalColunasColSpan) & "'><P class='F'>" & r("loja") & " - " & ucase(x_loja(r("loja"))) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				end if
			
			x = x & cab
			
			n_reg = 0
			n_saldo_parcial = 0
			vl_sub_total = 0
			loja_a = Trim("" & r("loja"))
			qtde_lojas = qtde_lojas + 1
			fabricante_a = "XXXXX"
			end if
			
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='" & CStr(intQtdeColunasColSpanSubTotal) & "'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MDB'><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "' class='MDBE'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MDBE' align='center' colspan='" & CStr(intQtdeTotalColunasColSpan) & "' style='background: honeydew'><P class='C'>&nbsp;" & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
			vl_sub_total = 0
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDBE' valign='bottom' NOWRAP><P class='C'>&nbsp;" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='bottom' width='" & CStr(LargColDescricao) & "'><P class='C'>&nbsp;" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

	'> ORDEM DE SERVIÇO
		if rb_estoque = ID_ESTOQUE_DANIFICADOS then
			s_lista_OS = ""
			s = "SELECT" & _
					" id_ordem_servico" & _
				" FROM t_ESTOQUE_MOVIMENTO" & _
				" WHERE" & _
					" (anulado_status=0)" & _
					" AND (estoque='" & ID_ESTOQUE_DANIFICADOS & "')" & _
					" AND (fabricante='" & Trim("" & r("fabricante")) & "')" & _
					" AND (produto='" & Trim("" & r("produto")) & "')" & _
					" AND (id_ordem_servico IS NOT NULL)" & _
				" ORDER BY" & _
					" id_ordem_servico"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			do while Not rs.Eof
				s_chave_OS = Trim("" & rs("id_ordem_servico"))
				s_num_OS_tela = formata_num_OS_tela(s_chave_OS)
				if s_lista_OS <> "" then s_lista_OS = s_lista_OS & ", "
				s_lista_OS = s_lista_OS & "<a href='OrdemServico.asp?num_OS=" & s_chave_OS & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "' title='Clique para consultar a Ordem de Serviço'>" & s_num_OS_tela & "</a>"
				rs.MoveNext
				loop
			
			if s_lista_OS = "" then s_lista_OS = "&nbsp;"
			x = x & "		<TD class='MDB' valign='bottom' width='" & CStr(LargColOrdemServico) & "'><P class='C'>" & s_lista_OS & "</P></TD>" & chr(13)
			end if

	 '> SALDO
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA UNITÁRIO (MÉDIA)
		if r("saldo") = 0 then vl = 0 else vl = r("preco_total")/r("saldo")
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA TOTAL
		vl = r("preco_total")
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

		vl_sub_total = vl_sub_total + vl
		vl_total_geral = vl_total_geral + vl
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		n_saldo_total = n_saldo_total + r("saldo")

		x = x & "	</TR>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MEB' COLSPAN='" & CStr(intQtdeColunasColSpanSubTotal) & "' NOWRAP><P class='Cd'>" & "Total:</P></TD>" & chr(13) & _
				"		<TD class='MB' NOWRAP><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MDB' NOWRAP><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
	'>	TOTAL GERAL
		if (qtde_lojas > 1) Or (qtde_fabricantes > 1) then
			x = x & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:honeydew'>" & chr(13)
			
			if rb_estoque = ID_ESTOQUE_DANIFICADOS then
				x = x & _
					"		<TD class='MTBE' valign='bottom' NOWRAP colspan='" & CStr(intQtdeColunasColSpanTotalGeral) & "'>&nbsp;</TD>" & chr(13)
			else
				x = x & _
					"		<TD class='MTBE' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13)
				end if
				
			x = x & _
				"		<TD class='MTB' valign='bottom' NOWRAP><p class='Cd'>TOTAL GERAL:</p></TD>" & chr(13) & _
				"		<TD class='MTB' valign='bottom' NOWRAP><p class='Cd'>" & formata_inteiro(n_saldo_total) & "</p></TD>" & chr(13) & _
				"		<TD class='MTB' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MTBD' valign='bottom' NOWRAP><p class='Cd'>" & formata_moeda(vl_total_geral) & "</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "' class='MB'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DA ÚLTIMA LOJA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing

end sub




' ____________________________________
' CONSULTA ESTOQUE DETALHE COMPLETO
'
sub consulta_estoque_detalhe_completo
dim r
dim s, s_aux, s_sql, loja_a, x, cab_table, cab, fabricante_a
dim n_reg, n_reg_total, n_saldo_parcial, n_saldo_total, qtde_lojas, qtde_fabricantes
dim vl, vl_sub_total, vl_total_geral

' 	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT loja, CONVERT(smallint,loja) AS numero_loja" & _
			", t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto" & _
			", descricao" & _
			", descricao_html" & _
			", Sum(t_ESTOQUE_MOVIMENTO.qtde) AS saldo" & _
			", t_ESTOQUE_ITEM.vl_custo2" & _
			" FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque='" & rb_estoque & "')"
	
	if cod_fabricante <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.fabricante='" & cod_fabricante & "')" 
		end if

	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.produto='" & cod_produto & "')"
		end if
	
	s_sql = s_sql & " GROUP BY loja, t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, descricao, descricao_html, t_ESTOQUE_ITEM.vl_custo2" & _
					" ORDER BY numero_loja, t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, descricao, descricao_html, t_ESTOQUE_ITEM.vl_custo2"

  ' CABEÇALHO
	cab_table = "<TABLE class='Q' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' nowrap>" & chr(13) & _
		  "		<TD width='60' valign='bottom' nowrap class='MD MB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='274' valign='bottom' nowrap class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
		  "		<TD width='60' valign='bottom' nowrap class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD width='100' valign='bottom' nowrap class='MD MB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA UNITÁRIO</P></TD>" & chr(13) & _
		  "		<TD width='100' valign='bottom' nowrap class='MB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	n_saldo_parcial = 0
	n_saldo_total = 0
	vl_sub_total = 0
	vl_total_geral = 0
	qtde_lojas = 0
	qtde_fabricantes = 0
	loja_a = "XXX"
	fabricante_a = "XXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU LOJA?
		if Trim("" & r("loja")) <> loja_a then
			if n_reg_total > 0 then 
			  ' FECHA TABELA DA LOJA ANTERIOR
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD COLSPAN='2' NOWRAP><P class='Cd'>Total:</P></TD>" & chr(13) & _
						"		<TD NOWRAP><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD>&nbsp;</TD>" & chr(13) & _
						"		<TD NOWRAP><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			x = x & cab_table

		  ' INICIA NOVA TABELA P/ A NOVA LOJA
			if Trim("" & r("loja")) <> "" then
			'	QUEBRA POR LOJA APENAS SE HOUVER LOJA
				x = x & "	<TR NOWRAP style='background:azure'>" & chr(13) & _
						"		<TD class='MB' align='center' colspan='5'><P class='F'>" & r("loja") & " - " & ucase(x_loja(r("loja"))) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				end if

			x = x & cab

			n_reg = 0
			n_saldo_parcial = 0
			vl_sub_total = 0
			qtde_lojas = qtde_lojas + 1
			loja_a = Trim("" & r("loja"))
			fabricante_a = "XXXXX"
			end if
		
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MB' colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='5' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MB' align='center' colspan='5' style='background: honeydew'><P class='C'>&nbsp;" & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
			vl_sub_total = 0
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='C'>&nbsp;" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='bottom'><P class='C' NOWRAP>&nbsp;" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

 	 '> SALDO
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA UNITÁRIO
		vl = r("vl_custo2")
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA TOTAL
		vl = r("vl_custo2")*r("saldo")
		x = x & "		<TD class='MB' valign='bottom' NOWRAP><P class='Cd'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

		vl_sub_total = vl_sub_total + vl
		vl_total_geral = vl_total_geral + vl
		
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		n_saldo_total = n_saldo_total + r("saldo")
		
		x = x & "	</TR>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD COLSPAN='2' NOWRAP><P class='Cd'>" & "Total:</P></TD>" & chr(13) & _
				"		<TD NOWRAP><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"		<TD>&nbsp;</TD>" & chr(13) & _
				"		<TD NOWRAP><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
	'>	TOTAL GERAL
		if (qtde_lojas > 1) Or (qtde_fabricantes > 1) then
			x = x & "</TABLE>" & chr(13) & _
				"<BR>" & chr(13) & _
				"<BR>" & chr(13) & _
				cab_table & _
				"	<TR style='background:honeydew'>" & chr(13) & _
				"		<TD width='60' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD width='274' valign='bottom' NOWRAP><p class='Cd'>TOTAL GERAL:</p></TD>" & chr(13) & _
				"		<TD width='60' valign='bottom' NOWRAP><p class='Cd'>" & formata_inteiro(n_saldo_total) & "</p></TD>" & chr(13) & _
				"		<TD width='100' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD width='100' valign='bottom' NOWRAP><p class='Cd'>" & formata_moeda(vl_total_geral) & "</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='5'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DA ÚLTIMA LOJA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing

end sub



' ___________________________________________
' CONSULTA ESTOQUE VENDA DETALHE SINTETICO
'
sub consulta_estoque_venda_detalhe_sintetico
dim r
dim s, s_aux, s_sql, x, cab_table, cab, fabricante_a
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes

	s_sql = "SELECT t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto," & _
			" descricao, descricao_html, Sum(qtde-qtde_utilizada) AS saldo" & _
			" FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON" & _
			" ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
			" WHERE ((qtde-qtde_utilizada) > 0)"

	if cod_fabricante <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_ITEM.fabricante='" & cod_fabricante & "')" 
		end if

	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
		end if
	
	s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html" & _
					" ORDER BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html"

  ' CABEÇALHO
	cab_table = "<TABLE class='Q' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' nowrap>" & chr(13) & _
		  "		<TD width='75' valign='bottom' nowrap class='MD MB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='480' valign='bottom' nowrap class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
		  "		<TD width='60' valign='bottom' nowrap class='MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table & cab
	n_reg = 0
	n_saldo_total = 0
	n_saldo_parcial = 0
	qtde_fabricantes = 0
	fabricante_a = "XXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MB' colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='3' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MB' align='center' colspan='3' style='background: honeydew'><P class='C'>&nbsp;" & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='C'>&nbsp;" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='bottom'><P class='C' NOWRAP>&nbsp;" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

 	 '> SALDO
		x = x & "		<TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

		n_saldo_total = n_saldo_total + r("saldo")
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		
		x = x & "	</TR>" & chr(13)

		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA TOTAL
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO FORNECEDOR
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
				"		<TD><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		
		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='3' class='MC'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='2' NOWRAP class='MC'><P class='Cd'>" & "TOTAL GERAL:" & "</P></TD>" & chr(13) & _
					"		<TD NOWRAP class='MC'><P class='Cd'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='3'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing

end sub





' ________________________________________________
' CONSULTA ESTOQUE VENDA DETALHE INTERMEDIARIO
'
sub consulta_estoque_venda_detalhe_intermediario
dim r
dim s, s_aux, s_sql, x, cab_table, cab, fabricante_a
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes
dim vl, vl_total_geral, vl_sub_total

	s_sql = "SELECT t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html" & _
			", Sum(qtde-qtde_utilizada) AS saldo, Sum((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total" & _
			" FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON" & _
			" ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
			" WHERE ((qtde-qtde_utilizada) > 0)"
	
	if cod_fabricante <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_ITEM.fabricante='" & cod_fabricante & "')" 
		end if

	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
		end if

	s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html" & _
					" ORDER BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html"

  ' CABEÇALHO
	cab_table = "<TABLE class='Q' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' nowrap>" & chr(13) & _
		  "		<TD width='60' valign='bottom' nowrap class='MD MB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='274' valign='bottom' class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
		  "		<TD width='60' valign='bottom' nowrap class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD width='100' valign='bottom' nowrap class='MD MB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
		  "		<TD width='100' valign='bottom' nowrap class='MB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table & cab
	n_reg = 0
	n_saldo_total = 0
	n_saldo_parcial = 0
	vl_total_geral = 0
	vl_sub_total = 0
	qtde_fabricantes = 0
	fabricante_a = "XXXXX"
		
	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MB' colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='5' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MB' align='center' colspan='5' style='background: honeydew'><P class='C'>&nbsp;" & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
			vl_sub_total = 0
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='C'>&nbsp;" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='bottom'><P class='C'>&nbsp;" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

 	 '> SALDO
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)
	
	 '> CUSTO DE ENTRADA UNITÁRIO (MÉDIA)
		if r("saldo") = 0 then vl = 0 else vl = r("preco_total")/r("saldo")
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA TOTAL
		vl = r("preco_total")
		x = x & "		<TD class='MB' valign='bottom' NOWRAP><P class='Cd'>" & formata_moeda(vl) & "</P></TD>" & chr(13)
		
		vl_total_geral = vl_total_geral + vl
		vl_sub_total = vl_sub_total + vl
		
		n_saldo_total = n_saldo_total + r("saldo")
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		
		x = x & "	</TR>" & chr(13)

		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA TOTAL
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO FORNECEDOR
		x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						"		<TD><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD>&nbsp;</TD>" & chr(13) & _
						"		<TD><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='5' class='MC'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='2' NOWRAP class='MC'><P class='Cd'>TOTAL GERAL:</P></TD>" & chr(13) & _
					"		<TD NOWRAP class='MC'><P class='Cd'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
					"		<TD class='MC'>&nbsp;</TD>" & chr(13) & _
					"		<TD NOWRAP class='MC'><P class='Cd'>" & formata_moeda(vl_total_geral) & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='5'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing

end sub





' _________________________________________
' CONSULTA ESTOQUE VENDA DETALHE COMPLETO
'
sub consulta_estoque_venda_detalhe_completo
dim r
dim s, s_aux, s_sql, x, cab_table, cab, fabricante_a
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes
dim vl, vl_total_geral, vl_sub_total

	s_sql = "SELECT t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html" & _
			", Sum(qtde-qtde_utilizada) AS saldo, t_ESTOQUE_ITEM.vl_custo2" & _
			" FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON" & _
			" ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
			" WHERE ((qtde-qtde_utilizada) > 0)"
	
	if cod_fabricante <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_ITEM.fabricante='" & cod_fabricante & "')" 
		end if

	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
		end if

	s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html, t_ESTOQUE_ITEM.vl_custo2" & _
					" ORDER BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html, t_ESTOQUE_ITEM.vl_custo2"

  ' CABEÇALHO
	cab_table = "<TABLE class='Q' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' nowrap>" & chr(13) & _
		  "		<TD width='60' valign='bottom' nowrap class='MD MB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='274' valign='bottom' nowrap class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
		  "		<TD width='60' valign='bottom' nowrap class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD width='100' valign='bottom' nowrap class='MD MB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA UNITÁRIO</P></TD>" & chr(13) & _
		  "		<TD width='100' valign='bottom' nowrap class='MB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table & cab
	n_reg = 0
	n_saldo_total = 0
	n_saldo_parcial = 0
	vl_total_geral = 0
	vl_sub_total = 0
	qtde_fabricantes = 0
	fabricante_a = "XXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MB' colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='5' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MB' align='center' colspan='5' style='background: honeydew'><P class='C'>&nbsp;" & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
			vl_sub_total = 0
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='C'>&nbsp;" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='bottom'><P class='C' NOWRAP>&nbsp;" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

	 '> SALDO
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)
	
	 '> CUSTO DE ENTRADA UNITÁRIO
		vl = r("vl_custo2")
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA TOTAL
		vl = r("vl_custo2")*r("saldo")
		x = x & "		<TD class='MB' valign='bottom' NOWRAP><P class='Cd'>" & formata_moeda(vl) & "</P></TD>" & chr(13)
		
		vl_total_geral = vl_total_geral + vl
		vl_sub_total = vl_sub_total + vl

		n_saldo_total = n_saldo_total + r("saldo")
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		
		x = x & "	</TR>" & chr(13)

		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA TOTAL
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO FORNECEDOR
		x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						"		<TD><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD>&nbsp;</TD>" & chr(13) & _
						"		<TD><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
	
		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='5' class='MC'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='2' NOWRAP class='MC'><P class='Cd'>TOTAL GERAL:</P></TD>" & chr(13) & _
					"		<TD NOWRAP class='MC'><P class='Cd'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
					"		<TD class='MC'>&nbsp;</TD>" & chr(13) & _
					"		<TD NOWRAP class='MC'><P class='Cd'>" & formata_moeda(vl_total_geral) & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='5'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
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

<style TYPE="text/css">
#rb_estoque_aux {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#rb_detalhe_aux {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
</style>

<% if rb_detalhe = "SINTETICO" then %>
<style TYPE="text/css">
P.C { font-size:10pt; }
P.Cc { font-size:10pt; }
P.Cd { font-size:10pt; }
P.F { font-size:11pt; }
</style>
<% end if %>


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

<form id="fESTOQ" name="fESTOQ" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="rb_estoque" id="rb_estoque" value="<%=rb_estoque%>">
<input type="hidden" name="rb_detalhe" id="rb_detalhe" value="<%=rb_detalhe%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque (Antigo)</span>
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
<br>

<!--  PARÂMETROS DA CONSULTA  -->
<table class="Qx" cellSpacing="0">
<!--  ESTOQUE  -->
	<tr bgColor="#FFFFFF">
		<% select case rb_estoque
				case ID_ESTOQUE_VENDA:			s = "VENDA"
				case ID_ESTOQUE_VENDIDO:		s = "VENDIDO"
				case ID_ESTOQUE_SHOW_ROOM:		s = "SHOW-ROOM"
				case ID_ESTOQUE_DANIFICADOS:	s = "PRODUTOS DANIFICADOS"
				case ID_ESTOQUE_DEVOLUCAO:		s = "DEVOLUÇÃO"
				case else						s = ""
				end select
		%>
	<td class="MT" NOWRAP><span class="PLTe">Estoque de Interesse</span>
		<br><p class="C" style="width:230px;cursor:default;"><%=s%></p></td>

<!--  TIPO DE DETALHAMENTO  -->
		<% select case rb_detalhe
			case "SINTETICO":		s = "SINTÉTICO (SEM CUSTOS)"
			case "INTERMEDIARIO":	s = "INTERMEDIÁRIO (CUSTOS MÉDIOS)"
			case "COMPLETO":		s = "COMPLETO (CUSTOS DIFERENCIADOS)"
			case else				s = ""
			end select
		%>
	<td class="MT" style="border-left:0px;" NOWRAP><span class="PLTe">Tipo de Detalhamento</span>
		<br><p class="C" style="width:230px;cursor:default;"><%=s%></p></td>
	</tr>

<!--  FABRICANTE  -->
	<% if cod_fabricante <> "" then %>
		<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP colspan="2"><span class="PLTe">Fabricante</span>
			<%	s = cod_fabricante
				if (s<>"") And (s_nome_fabricante<>"") then s = s & " - " & s_nome_fabricante %>
			<br><input name="c_fabricante_aux" id="c_fabricante_aux" readonly tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
					value="<%=s%>"></td>
		</tr>
	<% end if %>
	
<!--  PRODUTO  -->
	<% if cod_produto <> "" then %>
		<tr bgColor="#FFFFFF">
		<td class="MDBE" nowrap colspan="2"><span class="PLTe">Produto</span>
			<%	s = cod_produto
				if (s<>"") And (s_nome_produto_html<>"") then s = s & " - " & s_nome_produto_html %>
			<br>
				<span class="PLLe" style="width:460px;margin-left:2pt;"><%=s%></span>
			<%	s = cod_produto
				if (s<>"") And (s_nome_produto<>"") then s = s & " - " & s_nome_produto %>
				<input type="hidden" name="c_produto_aux" id="c_produto_aux" value="<%=s%>">
			</td>
		</tr>
	<% end if %>
	
</table>

<!--  RELATÓRIO  -->
<br>
<%	
	if rb_estoque = ID_ESTOQUE_VENDA then
		select case rb_detalhe
			case "SINTETICO"
				consulta_estoque_venda_detalhe_sintetico
			case "INTERMEDIARIO"
				consulta_estoque_venda_detalhe_intermediario
			case "COMPLETO"
				consulta_estoque_venda_detalhe_completo
			end select
	else
		select case rb_detalhe
			case "SINTETICO"
				consulta_estoque_detalhe_sintetico
			case "INTERMEDIARIO"
				consulta_estoque_detalhe_intermediario
			case "COMPLETO"
				consulta_estoque_detalhe_completo
			end select
		end if
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
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
