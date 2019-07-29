<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================
'	  R e l P o s i c a o E s t o q u e C m v P v E x e c . a s p
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

	const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
	const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
	const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
	const MSO_NUMBER_FORMAT_TEXTO = "\@"

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro, tPCI, t
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tPCI, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(t, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if (Not operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_BASICA_CMVPV, s_lista_operacoes_permitidas)) And (Not operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA_CMVPV, s_lista_operacoes_permitidas)) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, c_fabricante, c_produto, rb_estoque, rb_detalhe, rb_saida
	dim cod_fabricante, cod_produto
	dim s_nome_fabricante, s_nome_produto, s_nome_produto_html
    dim c_fabricante_multiplo, c_grupo, c_potencia_BTU, c_ciclo, c_posicao_mercado, v_fabricantes, v_grupos, rb_tipo_agrupamento, rb_tipo_produto
	dim blnSaidaExcel
    dim c_empresa

	c_fabricante = retorna_so_digitos(Request.Form("c_fabricante"))
	if c_fabricante <> "" then c_fabricante = normaliza_codigo(c_fabricante, TAM_MIN_FABRICANTE)
	c_produto = UCase(Trim(Request.Form("c_produto")))
	rb_estoque = Trim(Request.Form("rb_estoque"))
	rb_detalhe = Trim(Request.Form("rb_detalhe"))
	rb_saida = Trim(Request.Form("rb_saida"))
    rb_tipo_agrupamento = Trim(Request.Form("rb_tipo_agrupamento"))
    rb_tipo_produto = Trim(Request.Form("rb_tipo_produto"))

    c_fabricante_multiplo = Trim(Request.Form("c_fabricante_multiplo"))
	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_potencia_BTU = Trim(Request.Form("c_potencia_BTU"))
	c_ciclo = Trim(Request.Form("c_ciclo"))
	c_posicao_mercado = Trim(Request.Form("c_posicao_mercado"))
    c_empresa = Trim(Request.Form("c_empresa"))

	blnSaidaExcel = False
	if rb_saida = "XLS" then
		blnSaidaExcel = True
		end if
	
	alerta = ""
	if (c_produto<>"") And (Not IsEAN(c_produto)) then
		if c_fabricante = "" then alerta = "PARA CONSULTAR PELO CÓDIGO INTERNO DE PRODUTO É NECESSÁRIO ESPECIFICAR O FABRICANTE."
		end if
		
	if alerta = "" then
	'	DEFAULT
		cod_produto = c_produto
		cod_fabricante = c_fabricante

        call set_default_valor_texto_bd(usuario, "RelPosicaoEstoqueCmvPv|c_fabricante", c_fabricante_multiplo)
		call set_default_valor_texto_bd(usuario, "RelPosicaoEstoqueCmvPv|c_grupo", c_grupo)
		call set_default_valor_texto_bd(usuario, "RelPosicaoEstoqueCmvPv|c_potencia_BTU", c_potencia_BTU)
		call set_default_valor_texto_bd(usuario, "RelPosicaoEstoqueCmvPv|c_ciclo", c_ciclo)
		call set_default_valor_texto_bd(usuario, "RelPosicaoEstoqueCmvPv|c_posicao_mercado", c_posicao_mercado)
		
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

		if blnSaidaExcel then
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=Estoque_" & formata_data_yyyymmdd(Now) & "_" & formata_hora_hhnnss(Now) & ".xls"
			Response.Write "<h2>Relatório de Estoque</h2>"

			select case rb_estoque
				case ID_ESTOQUE_VENDA:			s = "VENDA"
				case ID_ESTOQUE_VENDIDO:		s = "VENDIDO"
				case ID_ESTOQUE_SHOW_ROOM:		s = "SHOW-ROOM"
				case ID_ESTOQUE_DANIFICADOS:	s = "PRODUTOS DANIFICADOS"
				case ID_ESTOQUE_DEVOLUCAO:		s = "DEVOLUÇÃO"
				case else						s = ""
				end select
			Response.Write "Estoque de Interesse: " & s
			Response.Write "<br>"

			select case rb_detalhe
				case "SINTETICO":		s = "SINTÉTICO (SEM CUSTOS)"
				case "INTERMEDIARIO":	s = "INTERMEDIÁRIO (CUSTOS MÉDIOS)"
				case "COMPLETO":		s = "COMPLETO (CUSTOS DIFERENCIADOS)"
				case else				s = ""
				end select
			Response.Write "Tipo de Detalhamento: " & s
			Response.Write "<br>"

			s = ""
			if cod_fabricante <> "" then 
				s = cod_fabricante
				if (s<>"") And (s_nome_fabricante<>"") then s = s & " - " & s_nome_fabricante
				end if
			if s <> "" then
				Response.Write "Fabricante: " & s
				Response.Write "<br>"
				end if

			s = ""
			if cod_produto <> "" then
				s = cod_produto
				if (s<>"") And (s_nome_produto_html<>"") then s = s & " - " & s_nome_produto_html
				end if
			if s <> "" then
				Response.Write "Produto: " & s
				Response.Write "<br>"
				end if
			
			s = "Emissão: " & formata_data_hora(Now)
			Response.Write s
			Response.Write "<br><br>"
			
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

			Response.End
			end if




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA ESTOQUE DETALHE SINTETICO
'
sub consulta_estoque_detalhe_sintetico
const LargColOrdemServico = 110
dim LargColProduto, LargColDescricao
dim r
dim s, s_aux, s_bkg_color, s_nbsp, s_sql, s_lista_OS, s_chave_OS, s_num_OS_tela, loja_a, x, cab_table, cab, fabricante_a, grupo_a, saldo_grupo, n_ctrl_reg
dim n_reg, n_reg_total, n_saldo_parcial, n_saldo_total, qtde_lojas, qtde_fabricantes
dim intQtdeColunasColSpanSubTotal, intQtdeColunasColSpanTotalGeral, intQtdeTotalColunasColSpan, s_where_temp, cont
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_TotalGrupo()

	intQtdeTotalColunasColSpan = 3
	intQtdeColunasColSpanSubTotal = 2
	intQtdeColunasColSpanTotalGeral = 1
	LargColProduto = 75
	LargColDescricao = 480
	if rb_estoque = ID_ESTOQUE_DANIFICADOS and rb_tipo_agrupamento <> "Grupo" then
		intQtdeTotalColunasColSpan = 4
		intQtdeColunasColSpanSubTotal = 3
		intQtdeColunasColSpanTotalGeral = 2
		LargColDescricao = 370
    end if

'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT" & _
                " grupo," & _
                " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
				" loja," & _
				" CONVERT(smallint,loja) AS numero_loja," & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html," & _
				" Sum(qtde) AS saldo" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
                " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
			" WHERE" & _
				" (anulado_status=0)" & _
				" AND (estoque='" & rb_estoque & "')"
	
	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.produto='" & cod_produto & "')"
		end if

    s_where_temp = ""
    if cod_fabricante <> "" then
		s_where_temp = s_where_temp & "(t_ESTOQUE_MOVIMENTO.fabricante='" & cod_fabricante & "')"
		end if

    v_fabricantes = split(c_fabricante_multiplo, ", ")
    if c_fabricante_multiplo <> "" then
	    for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			       " (t_ESTOQUE_MOVIMENTO.fabricante = '" & v_fabricantes(cont) & "')"
	    next
    	
    end if

    if s_where_temp <> "" then
        s_sql = s_sql & "AND"  
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    s_where_temp = ""
	if c_grupo <> "" then
	    v_grupos = split(c_grupo, ", ")
	    for cont = Lbound(v_grupos) to Ubound(v_grupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.grupo = '" & v_grupos(cont) & "')"
	    next
	    s_sql = s_sql & "AND "
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    if c_potencia_BTU <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.potencia_BTU = " & c_potencia_BTU & ")"
		end if

    if c_ciclo <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.ciclo = '" & c_ciclo & "')"
		end if
	
	if c_posicao_mercado <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.posicao_mercado = '" & c_posicao_mercado & "')"
		end if

    if c_empresa <> "" then
        s_sql = s_sql & _
            " AND (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
	end if

	s_sql = s_sql & _
			" GROUP BY" & _
				" loja," & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
                " grupo," & _
				" t_PRODUTO_GRUPO.descricao," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html"

    if rb_tipo_agrupamento = "Produto" then
		s_sql = s_sql & _ 
        " ORDER BY" & _
			" numero_loja," & _
			" t_ESTOQUE_MOVIMENTO.fabricante," & _
			" t_ESTOQUE_MOVIMENTO.produto," & _
			" t_PRODUTO.descricao," & _
			" descricao_html"
    elseif rb_tipo_agrupamento = "Grupo" then
        s_sql = s_sql & _ 
            " ORDER BY" & _
				" numero_loja," & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
                " grupo," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html"
    end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)

	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD width='" & Cstr(LargColProduto) & "' valign='bottom' class='MDBE'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='" & CStr(LargColDescricao) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13)

	if rb_estoque = ID_ESTOQUE_DANIFICADOS and rb_tipo_agrupamento <> "Grupo" then
		cab = cab & _
			  "		<TD width='" & CStr(LargColOrdemServico) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>ORDEM SERVIÇO</P></TD>" & chr(13)
		end if
	
	cab = cab & _
		  "		<TD width='60' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
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
    grupo_a = "XXXXX"
    saldo_grupo = 0
    n_ctrl_reg = 0

	if blnSaidaExcel then s_nbsp = "" else s_nbsp = "&nbsp;"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	'	MUDOU LOJA?
		if Trim("" & r("loja")) <> loja_a then
			if n_reg_total > 0 then 

              ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = "Grupo" then
		            x = x & "		<TD class='MB ME' colspan='2' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

			  ' FECHA TABELA DA LOJA ANTERIOR
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD COLSPAN='" & Cstr(intQtdeColunasColSpanSubTotal) & "' align='right' valign='bottom' class='MEB' NOWRAP><P class='Cd' style='font-weight:bold;'>" & _
						"TOTAL:" & "</P></TD>" & chr(13) & _
						"		<TD NOWRAP class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & _
						"</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				if blnSaidaExcel then
					x="<BR>" & chr(13) & "<BR>" & chr(13)
				else
					x="<BR>" & chr(13)
					end if
				end if

			x = x & cab_table

		  ' INICIA NOVA TABELA P/ A NOVA LOJA
			if Trim("" & r("loja")) <> "" then
			'	QUEBRA POR LOJA APENAS SE HOUVER LOJA
				if blnSaidaExcel then s_bkg_color = "tomato" else s_bkg_color = "azure"
				x = x & "	<TR NOWRAP style='background:" & s_bkg_color & "'>" & chr(13) & _
						"		<TD class='MDBE' align='center' valign='bottom' colspan='" & CStr(intQtdeTotalColunasColSpan) & "'><P class='F' style='font-weight:bold;'>" & r("loja") & " - " & ucase(x_loja(r("loja"))) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				end if

			x = x & cab

			n_reg = 0
			n_saldo_parcial = 0
			loja_a = Trim("" & r("loja"))
			qtde_lojas = qtde_lojas + 1
			fabricante_a = "XXXXX"
            grupo_a = "XXXXX"
			end if
			
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then
                ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = "Grupo" then
		            x = x & "		<TD class='MB ME' colspan='2' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' align='right' valign='bottom' colspan='" & CStr(intQtdeColunasColSpanSubTotal) & "'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "' class='MB' valign='bottom'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MDBE' align='center' valign='bottom' colspan='" & CStr(intQtdeTotalColunasColSpan) & "' style='background: honeydew'><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
            grupo_a = "XXXXXXX"
			end if
			
	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

        if rb_tipo_agrupamento = "Produto" then

		    x = x & "	<TR NOWRAP>" & chr(13)

	     '> PRODUTO
		    x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

	    '> ORDEM DE SERVIÇO
		    if rb_estoque = ID_ESTOQUE_DANIFICADOS then
			    s_lista_OS = ""
			    s = "SELECT" & _
					    " id_ordem_servico" & _
				    " FROM t_ESTOQUE_MOVIMENTO" & _
                    " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
				    " WHERE" & _
					    " (anulado_status=0)" & _
					    " AND (estoque='" & ID_ESTOQUE_DANIFICADOS & "')" & _
					    " AND (fabricante='" & Trim("" & r("fabricante")) & "')" & _
					    " AND (produto='" & Trim("" & r("produto")) & "')" & _
					    " AND (id_ordem_servico IS NOT NULL)" 
            
                if c_empresa <> "" then
                    s = s & _
                    " AND (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
	            end if    

				s = s & " ORDER BY" & _
					    " id_ordem_servico"
			    if rs.State <> 0 then rs.Close
			    rs.open s, cn
			    do while Not rs.Eof
				    s_chave_OS = Trim("" & rs("id_ordem_servico"))
				    s_num_OS_tela = formata_num_OS_tela(s_chave_OS)
				    if s_lista_OS <> "" then s_lista_OS = s_lista_OS & ", "
				    if blnSaidaExcel then
					    s_lista_OS = s_lista_OS & s_num_OS_tela
				    else
					    s_lista_OS = s_lista_OS & "<a href='OrdemServico.asp?num_OS=" & s_chave_OS & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "' title='Clique para consultar a Ordem de Serviço'>" & s_num_OS_tela & "</a>"
					    end if
				    rs.MoveNext
				    loop
			
			    if (s_lista_OS = "") And (Not blnSaidaExcel) then s_lista_OS = "&nbsp;"
			    x = x & "		<TD class='MDB' align='left' valign='middle' width='" & CStr(LargColOrdemServico) & "'><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_lista_OS & "</P></TD>" & chr(13)
			    end if
			
	     '> SALDO
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

		    n_saldo_parcial = n_saldo_parcial + r("saldo")
		    n_saldo_total = n_saldo_total + r("saldo")
		
		    x = x & "	</TR>" & chr(13)

        elseif rb_tipo_agrupamento = "Grupo" then
            if Trim("" & r("grupo")) <> grupo_a then
                if n_ctrl_reg > 0 then
		            x = x & "		<TD class='MB ME' colspan='2' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    cont = cont + 1
                end if
                x = x & "	<TR NOWRAP>" & chr(13)
		        x = x & "		<TD class='MDB ME' colspan='3' align='left' style='background-color: #EEE' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)

                redim preserve vt_CodigoGrupo(cont)
                redim preserve vt_DescricaoGrupo(cont)     
                if Trim("" & r("grupo")) = "" then
                    vt_CodigoGrupo(cont) = "-"
                else
                    vt_CodigoGrupo(cont) = Trim("" & r("grupo"))
                end if
                vt_DescricaoGrupo(cont) = Trim("" & r("grupo_descricao"))           

                grupo_a = Trim("" & r("grupo"))    
                saldo_grupo = 0            
                n_ctrl_reg = n_ctrl_reg + 1
            end if

            x = x & "   <TR NOWRAP>" & chr(13)

         '> PRODUTO
		    x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

        '> SALDO
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

            x = x & "   </TR>" & chr(13)

            saldo_grupo = saldo_grupo + CInt(r("saldo"))
            n_saldo_parcial = n_saldo_parcial + r("saldo")
		    n_saldo_total = n_saldo_total + r("saldo")

        end if
		
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop

    ' MOSTRA TOTAL DO ÚLTIMO GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        if n_reg <> 0 then
            x = x & "		<TD class='MB ME' colspan='2' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)                    
            vt_QtdeGrupo(cont) = saldo_grupo
        end if
    end if
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP>"  & chr(13) & _
				"		<TD class='MEB' COLSPAN='" & CStr(intQtdeColunasColSpanSubTotal) & "' align='right' NOWRAP><P class='Cd' style='font-weight:bold;'>" & "TOTAL:" & "</P></TD>" & chr(13) & _
				"		<TD class='MDB' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

    ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        redim vt_TotalGrupo(0)
        set vt_TotalGrupo(0) = New cl_TRES_COLUNAS
        with vt_TotalGrupo(0)
            .c1 = ""
            .c2 = ""
            .c3 = ""
        end with

        if vt_CodigoGrupo(UBound(vt_CodigoGrupo)) <> "" then
            for cont = 0 to UBound(vt_CodigoGrupo) 
                if Trim(vt_TotalGrupo(UBound(vt_TotalGrupo)).c1) <> "" then
                    redim preserve vt_TotalGrupo(UBound(vt_TotalGrupo)+1)
                    set vt_TotalGrupo(UBound(vt_TotalGrupo)) = New cl_TRES_COLUNAS
                end if
                with vt_TotalGrupo(UBound(vt_TotalGrupo))
                    .c1 = vt_CodigoGrupo(cont)
                    .c2 = vt_DescricaoGrupo(cont)
                    .c3 = vt_QtdeGrupo(cont)
                end with
            next
        end if
        ordena_cl_tres_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        n_ctrl_reg = 0
        grupo_a = "XXXXXXX"

        if (qtde_fabricantes > 1) Or (qtde_lojas > 1) then
            x = x & _
                "	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:azure'>" & chr(13) & _
                "		<TD class='MTBD ME' valign='bottom' colspan='" & CStr(intQtdeTotalColunasColSpan) & "' NOWRAP><P class='R' style='font-weight:bold;'>TOTAL GERAL POR GRUPO DE PRODUTOS</P></TD>" & chr(13) & _
                "   </TR>" & CHR(13)

        for cont = 0 to UBound(vt_TotalGrupo)
            if vt_TotalGrupo(cont).c1 <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & _
                    "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
				    "	</TR>" & chr(13)
                    saldo_grupo = 0
                end if
                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MDB ME' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
			"	</TR>" & chr(13)

        end if
    end if




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
				"		<TD class='MTB' align='right' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;'>TOTAL GERAL:</p></TD>" & chr(13) & _
				"		<TD class='MTBD' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_total) & "</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "' align='center' class='MDBE'><P class='ALERTA'>NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS</P></TD>" & chr(13) & _
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
dim LargColProduto, LargColDescricao
dim r
dim s, s_aux, s_bkg_color, s_nbsp, s_sql, s_lista_OS, s_chave_OS, s_num_OS_tela, loja_a, x, cab_table, cab, fabricante_a, grupo_a, saldo_grupo, vl_custo_entrada_grupo, n_ctrl_reg
dim n_reg, n_reg_total, n_saldo_parcial, n_saldo_total, qtde_lojas, qtde_fabricantes
dim vl, vl_sub_total, vl_total_geral
dim intQtdeColunasColSpanSubTotal, intQtdeColunasColSpanTotalGeral, intQtdeTotalColunasColSpan, s_where_temp, cont
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_CustoEntrada(), vt_TotalGrupo()

	intQtdeTotalColunasColSpan = 5
	intQtdeColunasColSpanSubTotal = 2
	intQtdeColunasColSpanTotalGeral = 1
	LargColDescricao = 270
	if rb_estoque = ID_ESTOQUE_DANIFICADOS and rb_tipo_agrupamento <> "Grupo" then
		intQtdeTotalColunasColSpan = 6
		intQtdeColunasColSpanSubTotal = 3
		intQtdeColunasColSpanTotalGeral = 2
		LargColDescricao = 200
    end if

	LargColProduto = 60
	if blnSaidaExcel then
		LargColProduto = 75
		LargColDescricao = 350
		end if

'	IMPORTANTE: O VALOR ATUAL DE CMV_PV ESTÁ EM T_PRODUTO.PRECO_FABRICANTE
'	==========  O HISTÓRICO DO VALOR DE CMV_PV ESTÁ EM T_PEDIDO_ITEM.PRECO_FABRICANTE (E T_PEDIDO_ITEM_DEVOLVIDO.PRECO_FABRICANTE)
'				O HISTÓRICO DO CUSTO REAL PAGO AO FABRICANTE ESTÁ EM T_ESTOQUE_ITEM.VL_CUSTO2

'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT" & _
                " t_PRODUTO.grupo AS grupo," & _
                " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
				" loja," & _
				" CONVERT(smallint,loja) AS numero_loja," & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html," & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde) AS saldo,"
				
	if rb_estoque = ID_ESTOQUE_VENDIDO then
		s_sql = s_sql & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_fabricante) AS preco_total" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
				" INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
                " LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
				" INNER JOIN t_PEDIDO_ITEM ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto)"
	elseif rb_estoque = ID_ESTOQUE_DEVOLUCAO then
		s_sql = s_sql & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_fabricante) AS preco_total" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
				" INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
                " LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
				" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM_DEVOLVIDO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM_DEVOLVIDO.produto)"
	else
		s_sql = s_sql & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PRODUTO.preco_fabricante) AS preco_total" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
				" INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
                " LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)"
		end if

	s_sql = s_sql & _
			" WHERE" & _
				" (anulado_status=0)" & _
				" AND (estoque='" & rb_estoque & "')"

	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.produto='" & cod_produto & "')"
		end if

    s_where_temp = ""
    if cod_fabricante <> "" then
		s_where_temp = s_where_temp & "(t_ESTOQUE_MOVIMENTO.fabricante='" & cod_fabricante & "')"
		end if

    v_fabricantes = split(c_fabricante_multiplo, ", ")
    if c_fabricante_multiplo <> "" then
	    for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			       " (t_ESTOQUE_MOVIMENTO.fabricante = '" & v_fabricantes(cont) & "')"
	    next
    	
    end if

    if s_where_temp <> "" then
        s_sql = s_sql & "AND"  
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    s_where_temp = ""
	if c_grupo <> "" then
	    v_grupos = split(c_grupo, ", ")
	    for cont = Lbound(v_grupos) to Ubound(v_grupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.grupo = '" & v_grupos(cont) & "')"
	    next
	    s_sql = s_sql & "AND "
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    if c_potencia_BTU <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.potencia_BTU = " & c_potencia_BTU & ")"
		end if

    if c_ciclo <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.ciclo = '" & c_ciclo & "')"
		end if
	
	if c_posicao_mercado <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.posicao_mercado = '" & c_posicao_mercado & "')"
		end if

    if c_empresa <> "" then
        s_sql = s_sql & _
        " AND (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
	end if

	s_sql = s_sql & _
			" GROUP BY" & _
				" loja," & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
                " t_PRODUTO.grupo," & _
				" t_PRODUTO_GRUPO.descricao," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html"
			if rb_tipo_agrupamento = "Produto" then
		        s_sql = s_sql & _ 
                " ORDER BY" & _
			        " numero_loja," & _
			        " t_ESTOQUE_MOVIMENTO.fabricante," & _
			        " t_ESTOQUE_MOVIMENTO.produto," & _
			        " t_PRODUTO.descricao," & _
			        " descricao_html"
            elseif rb_tipo_agrupamento = "Grupo" then
                s_sql = s_sql & _ 
                    " ORDER BY" & _
				        " numero_loja," & _
				        " t_ESTOQUE_MOVIMENTO.fabricante," & _
                        " t_PRODUTO.grupo," & _
				        " t_ESTOQUE_MOVIMENTO.produto," & _
				        " t_PRODUTO.descricao," & _
				        " descricao_html"
            end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure'>" & chr(13) & _
		  "		<TD width='" & Cstr(LargColProduto) & "' valign='bottom' class='MDBE'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='" & CStr(LargColDescricao) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13)

	if rb_estoque = ID_ESTOQUE_DANIFICADOS then
		cab = cab & _
			  "		<TD width='" & CStr(LargColOrdemServico) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>ORDEM SERVIÇO</P></TD>" & chr(13)
		end if
	
	cab = cab & _
		  "		<TD width='60' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13)
	
	if blnSaidaExcel then
		cab = cab & _
		  "		<TD style='width:100px' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;width:100px;'>CUSTO ENTRADA<br style='mso-data-placement:same-cell;' />UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
		  "		<TD style='width:100px' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;width:100px;'>CUSTO ENTRADA<br style='mso-data-placement:same-cell;' />TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	else
		cab = cab & _
		  "		<TD width='100' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
		  "		<TD width='100' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
		end if

	x = ""
	n_reg = 0
	n_reg_total = 0
	n_saldo_parcial = 0
	n_saldo_total = 0
	vl_sub_total = 0
	vl_total_geral = 0
	qtde_lojas = 0
    saldo_grupo = 0
    n_ctrl_reg = 0
	loja_a = "XXX"
	fabricante_a = "XXXXX"
    grupo_a = "XXXXX"
	
	if blnSaidaExcel then s_nbsp = "" else s_nbsp = "&nbsp;"

	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU LOJA?
		if Trim("" & r("loja")) <> loja_a then
			if n_reg_total > 0 then 

            ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = "Grupo" then
		            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

			  ' FECHA TABELA DA LOJA ANTERIOR
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD COLSPAN='" & Cstr(intQtdeColunasColSpanSubTotal) & "' align='right' valign='bottom' class='MEB' NOWRAP><P class='Cd' style='font-weight:bold;'>" & _
						"TOTAL:</P></TD>" & chr(13) & _
						"		<TD NOWRAP class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & _
						"</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'>&nbsp;</TD>" & chr(13) & _
						"		<TD NOWRAP class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & _
						formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				if blnSaidaExcel then
					x="<BR>" & chr(13) & "<BR>" & chr(13)
				else
					x="<BR>" & chr(13)
					end if
				end if

			x = x & cab_table

		  ' INICIA NOVA TABELA P/ A NOVA LOJA
			if Trim("" & r("loja")) <> "" then
			'   QUEBRA POR LOJA APENAS SE HOUVER LOJA
				if blnSaidaExcel then s_bkg_color = "tomato" else s_bkg_color = "azure"
				x = x & "	<TR style='background:" & s_bkg_color & "'>" & chr(13) & _
						"		<TD class='MDBE' align='center' valign='bottom' colspan='" & CStr(intQtdeTotalColunasColSpan) & "'><P class='F' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & r("loja") & " - " & ucase(x_loja(r("loja"))) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				end if
			
			x = x & cab
			
			n_reg = 0
			n_saldo_parcial = 0
			vl_sub_total = 0
			loja_a = Trim("" & r("loja"))
			qtde_lojas = qtde_lojas + 1
			fabricante_a = "XXXXX"
            grupo_a = "XXXXX"
			end if
			
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then

                ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = "Grupo" then
		            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='" & CStr(intQtdeColunasColSpanSubTotal) & "' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MDBE' align='center' valign='bottom' colspan='" & CStr(intQtdeTotalColunasColSpan) & "' style='background: honeydew'><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
			vl_sub_total = 0
            grupo_a = "XXXXXXX"
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

    if rb_tipo_agrupamento = "Produto" then

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

	'> ORDEM DE SERVIÇO
		if rb_estoque = ID_ESTOQUE_DANIFICADOS then
			s_lista_OS = ""
			s = "SELECT" & _
					" id_ordem_servico" & _
				" FROM t_ESTOQUE_MOVIMENTO" & _
                " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
				" WHERE" & _
					" (anulado_status=0)" & _
					" AND (estoque='" & ID_ESTOQUE_DANIFICADOS & "')" & _
					" AND (fabricante='" & Trim("" & r("fabricante")) & "')" & _
					" AND (produto='" & Trim("" & r("produto")) & "')" & _
					" AND (id_ordem_servico IS NOT NULL)" 

            if c_empresa <> "" then
                s = s & _
                " AND (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
	        end if

			s = s & " ORDER BY" & _
					" id_ordem_servico"

			if rs.State <> 0 then rs.Close
			rs.open s, cn
			do while Not rs.Eof
				s_chave_OS = Trim("" & rs("id_ordem_servico"))
				s_num_OS_tela = formata_num_OS_tela(s_chave_OS)
				if s_lista_OS <> "" then s_lista_OS = s_lista_OS & ", "
				if blnSaidaExcel then
					s_lista_OS = s_lista_OS & s_num_OS_tela
				else
					s_lista_OS = s_lista_OS & "<a href='OrdemServico.asp?num_OS=" & s_chave_OS & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "' title='Clique para consultar a Ordem de Serviço'>" & s_num_OS_tela & "</a>"
					end if
				rs.MoveNext
				loop
			
			if (s_lista_OS = "") And (Not blnSaidaExcel) then s_lista_OS = "&nbsp;"
			x = x & "		<TD class='MDB' align='left' valign='middle' width='" & CStr(LargColOrdemServico) & "'><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_lista_OS & "</P></TD>" & chr(13)
			end if

	 '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA UNITÁRIO (MÉDIA)
		if r("saldo") = 0 then vl = 0 else vl = r("preco_total")/r("saldo")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA TOTAL
		vl = r("preco_total")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

		vl_sub_total = vl_sub_total + vl
		vl_total_geral = vl_total_geral + vl
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		n_saldo_total = n_saldo_total + r("saldo")

		x = x & "	</TR>" & chr(13)

    elseif rb_tipo_agrupamento = "Grupo" then
            if Trim("" & r("grupo")) <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                end if
                x = x & "	<TR NOWRAP>" & chr(13)
		        x = x & "		<TD class='MDB ME' colspan='5' align='left' style='background-color: #EEE' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)


                redim preserve vt_CodigoGrupo(cont)
                redim preserve vt_DescricaoGrupo(cont)     
                if Trim("" & r("grupo")) = "" then
                    vt_CodigoGrupo(cont) = "-"
                else
                    vt_CodigoGrupo(cont) = Trim("" & r("grupo"))
                end if
                vt_DescricaoGrupo(cont) = Trim("" & r("grupo_descricao"))           

                grupo_a = Trim("" & r("grupo"))    
                saldo_grupo = 0          
                vl_custo_entrada_grupo = 0  
                n_ctrl_reg = n_ctrl_reg + 1
            end if

        x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

     '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA UNITÁRIO (MÉDIA)
		if r("saldo") = 0 then vl = 0 else vl = r("preco_total")/r("saldo")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA TOTAL
		vl = r("preco_total")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)


            saldo_grupo = saldo_grupo + CInt(r("saldo"))
            vl_custo_entrada_grupo = vl_custo_entrada_grupo + r("preco_total")

            n_saldo_parcial = n_saldo_parcial + r("saldo")
		    n_saldo_total = n_saldo_total + r("saldo")
            vl_sub_total = vl_sub_total + vl
		    vl_total_geral = vl_total_geral + vl

            x = x & "	</TR>" & chr(13)
        end if

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop

  ' MOSTRA TOTAL DO ÚLTIMO GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        if n_reg <> 0 then
            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)         
            redim preserve vt_CustoEntrada(cont)           
            vt_QtdeGrupo(cont) = saldo_grupo
            vt_CustoEntrada(cont) = vl_custo_entrada_grupo
        end if
    end if
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MEB' COLSPAN='" & CStr(intQtdeColunasColSpanSubTotal) & "' align='right' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;'>" & "TOTAL:</P></TD>" & chr(13) & _
				"		<TD class='MB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"		<TD class='MB' valign='bottom'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

   ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        redim vt_TotalGrupo(0)
        set vt_TotalGrupo(0) = New cl_QUATRO_COLUNAS
        with vt_TotalGrupo(0)
            .c1 = ""
            .c2 = ""
            .c3 = ""
            .c4 = ""
        end with

        if vt_CodigoGrupo(UBound(vt_CodigoGrupo)) <> "" then
            for cont = 0 to UBound(vt_CodigoGrupo) 
                if Trim(vt_TotalGrupo(UBound(vt_TotalGrupo)).c1) <> "" then
                    redim preserve vt_TotalGrupo(UBound(vt_TotalGrupo)+1)
                    set vt_TotalGrupo(UBound(vt_TotalGrupo)) = New cl_QUATRO_COLUNAS
                end if
                with vt_TotalGrupo(UBound(vt_TotalGrupo))
                    .c1 = vt_CodigoGrupo(cont)
                    .c2 = vt_DescricaoGrupo(cont)
                    .c3 = vt_QtdeGrupo(cont)
                    .c4 = vt_CustoEntrada(cont)
                end with
            next
        end if
        ordena_cl_quatro_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        vl_custo_entrada_grupo = 0
        n_ctrl_reg = 0
        grupo_a = "XXXXXXX"

        if (qtde_fabricantes > 1) Or (qtde_lojas > 1) then
            x = x & _
                "	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:azure'>" & chr(13) & _
                "		<TD class='MTBD ME' valign='bottom' colspan='" & CStr(intQtdeTotalColunasColSpan) & "' NOWRAP><P class='R' style='font-weight:bold;'>TOTAL GERAL POR GRUPO DE PRODUTOS</P></TD>" & chr(13) & _
                "   </TR>" & CHR(13)

        for cont = 0 to UBound(vt_TotalGrupo)
            if vt_TotalGrupo(cont).c1 <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & _
                    "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
                    "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    				"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
                    "	</TR>" & chr(13)
                    saldo_grupo = 0
                    vl_custo_entrada_grupo = 0
                end if
                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MB ME' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            vl_custo_entrada_grupo = vl_custo_entrada_grupo + vt_TotalGrupo(cont).c4
            
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
            "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    		"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
			"	</TR>" & chr(13)

        end if
    end if


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
				"		<TD class='MTB' align='right' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;'>TOTAL GERAL:</p></TD>" & chr(13) & _
				"		<TD class='MTB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_total) & "</p></TD>" & chr(13) & _
				"		<TD class='MTB' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MTBD' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_geral) & "</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "' align='center' class='MDBE'><P class='ALERTA'>NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS</P></TD>" & chr(13) & _
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
dim s, s_aux, s_bkg_color, s_nbsp, s_sql, loja_a, x, cab_table, cab, fabricante_a, grupo_a, saldo_grupo, vl_custo_entrada_grupo, n_ctrl_reg
dim n_reg, n_reg_total, n_saldo_parcial, n_saldo_total, qtde_lojas, qtde_fabricantes
dim vl, vl_sub_total, vl_total_geral
dim largColProduto, largColDescricao, s_where_temp, cont
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_CustoEntrada(), vt_TotalGrupo()

'	IMPORTANTE: O VALOR ATUAL DE CMV_PV ESTÁ EM T_PRODUTO.PRECO_FABRICANTE
'	==========  O HISTÓRICO DO VALOR DE CMV_PV ESTÁ EM T_PEDIDO_ITEM.PRECO_FABRICANTE (E T_PEDIDO_ITEM_DEVOLVIDO.PRECO_FABRICANTE)
'				O HISTÓRICO DO CUSTO REAL PAGO AO FABRICANTE ESTÁ EM T_ESTOQUE_ITEM.VL_CUSTO2

'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT" & _
                " t_PRODUTO.grupo AS grupo," & _
                " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
				" loja," & _
				" CONVERT(smallint,loja) AS numero_loja," & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html," & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde) AS saldo,"

	if rb_estoque = ID_ESTOQUE_VENDIDO then
		s_sql = s_sql & _
				" t_PEDIDO_ITEM.preco_fabricante" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
				" INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
                " LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
				" INNER JOIN t_PEDIDO_ITEM ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto)"
	elseif rb_estoque = ID_ESTOQUE_DEVOLUCAO then
		s_sql = s_sql & _
				" t_PEDIDO_ITEM_DEVOLVIDO.preco_fabricante" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
				" INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
                " LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
				" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM_DEVOLVIDO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM_DEVOLVIDO.produto)"
	else
		s_sql = s_sql & _
				" t_PRODUTO.preco_fabricante" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
				" INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
                " LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)"
		end if
		
	s_sql = s_sql & _
			" WHERE" & _
				" (anulado_status=0)" & _
				" AND (estoque='" & rb_estoque & "')"
	
	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.produto='" & cod_produto & "')"
		end if

    s_where_temp = ""
    if cod_fabricante <> "" then
		s_where_temp = s_where_temp & "(t_ESTOQUE_MOVIMENTO.fabricante='" & cod_fabricante & "')"
		end if

    v_fabricantes = split(c_fabricante_multiplo, ", ")
    if c_fabricante_multiplo <> "" then
	    for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			       " (t_ESTOQUE_MOVIMENTO.fabricante = '" & v_fabricantes(cont) & "')"
	    next
    	
    end if

    if s_where_temp <> "" then
        s_sql = s_sql & "AND"  
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if
	
    s_where_temp = ""
	if c_grupo <> "" then
	    v_grupos = split(c_grupo, ", ")
	    for cont = Lbound(v_grupos) to Ubound(v_grupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.grupo = '" & v_grupos(cont) & "')"
	    next
	    s_sql = s_sql & "AND "
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    if c_potencia_BTU <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.potencia_BTU = " & c_potencia_BTU & ")"
		end if

    if c_ciclo <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.ciclo = '" & c_ciclo & "')"
		end if
	
	if c_posicao_mercado <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.posicao_mercado = '" & c_posicao_mercado & "')"
		end if

    if c_empresa <> "" then
        s_sql = s_sql & _
            " AND (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
	end if

	if rb_estoque = ID_ESTOQUE_VENDIDO then
		s_sql = s_sql & _
				" GROUP BY" & _
					" loja," & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
                    " t_PRODUTO.grupo," & _
				    " t_PRODUTO_GRUPO.descricao," & _
					" t_ESTOQUE_MOVIMENTO.produto," & _
					" t_PRODUTO.descricao," & _
					" t_PRODUTO.descricao_html," & _
					" t_PEDIDO_ITEM.preco_fabricante"
        if rb_tipo_agrupamento = "Produto" then
		        s_sql = s_sql & _
				    " ORDER BY" & _
					    " numero_loja," & _
					    " t_ESTOQUE_MOVIMENTO.fabricante," & _
					    " t_ESTOQUE_MOVIMENTO.produto," & _
					    " t_PRODUTO.descricao," & _
					    " t_PRODUTO.descricao_html," & _
					    " t_PEDIDO_ITEM.preco_fabricante"
        elseif rb_tipo_agrupamento = "Grupo" then
                s_sql = s_sql & _
                    " ORDER BY" & _
					    " numero_loja," & _
					    " t_ESTOQUE_MOVIMENTO.fabricante," & _
                        " t_PRODUTO.grupo," & _
					    " t_ESTOQUE_MOVIMENTO.produto," & _
					    " t_PRODUTO.descricao," & _
					    " t_PRODUTO.descricao_html," & _
					    " t_PEDIDO_ITEM.preco_fabricante"
        end if
	elseif rb_estoque = ID_ESTOQUE_DEVOLUCAO then
		s_sql = s_sql & _
				" GROUP BY" & _
					" loja," & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
                    " t_PRODUTO.grupo," & _
				    " t_PRODUTO_GRUPO.descricao," & _
					" t_ESTOQUE_MOVIMENTO.produto," & _
					" t_PRODUTO.descricao," & _
					" t_PRODUTO.descricao_html," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.preco_fabricante"
        if rb_tipo_agrupamento = "Produto" then
		        s_sql = s_sql & _ 
				" ORDER BY" & _
					" numero_loja," & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
					" t_ESTOQUE_MOVIMENTO.produto," & _
					" t_PRODUTO.descricao," & _
					" t_PRODUTO.descricao_html," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.preco_fabricante"
        elseif rb_tipo_agrupamento = "Grupo" then
                s_sql = s_sql & _ 
                " ORDER BY" & _
					" numero_loja," & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
                    " t_PRODUTO.grupo," & _
					" t_ESTOQUE_MOVIMENTO.produto," & _
					" t_PRODUTO.descricao," & _
					" t_PRODUTO.descricao_html," & _
					" t_PEDIDO_ITEM_DEVOLVIDO.preco_fabricante"
        end if
	else
		s_sql = s_sql & _
				" GROUP BY" & _
					" loja," & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
                    " t_PRODUTO.grupo," & _
				    " t_PRODUTO_GRUPO.descricao," & _
					" t_ESTOQUE_MOVIMENTO.produto," & _
					" t_PRODUTO.descricao," & _
					" t_PRODUTO.descricao_html," & _
					" t_PRODUTO.preco_fabricante"
        if rb_tipo_agrupamento = "Produto" then
		        s_sql = s_sql & _ 
				" ORDER BY" & _
					" numero_loja," & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
					" t_ESTOQUE_MOVIMENTO.produto," & _
					" t_PRODUTO.descricao," & _
					" t_PRODUTO.descricao_html," & _
					" t_PRODUTO.preco_fabricante"
        elseif rb_tipo_agrupamento = "Grupo" then
                s_sql = s_sql & _ 
                " ORDER BY" & _
					" numero_loja," & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
                    " t_PRODUTO.grupo," & _
					" t_ESTOQUE_MOVIMENTO.produto," & _
					" t_PRODUTO.descricao," & _
					" t_PRODUTO.descricao_html," & _
					" t_PRODUTO.preco_fabricante"
        end if
	end if

	if blnSaidaExcel then
		largColProduto = 75
		largColDescricao = 350
	else
		largColProduto = 60
		largColDescricao = 274
		end if
	
  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13)
	
	if blnSaidaExcel then
		cab = cab & _
		  "		<TD class='ME' width='" & Cstr(largColProduto) & "' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='" & Cstr(largColDescricao) & "' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
		  "		<TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA<br style='mso-data-placement:same-cell;' />UNITÁRIO</P></TD>" & chr(13) & _
		  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA<br style='mso-data-placement:same-cell;' />TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	else
		cab = cab & _
		  "		<TD width='" & Cstr(largColProduto) & "' valign='bottom' class='MD ME MB'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='" & Cstr(largColDescricao) & "' valign='bottom' class='MD MB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
		  "		<TD width='60' align='right' valign='bottom' class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD width='100' align='right' valign='bottom' class='MD MB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA UNITÁRIO</P></TD>" & chr(13) & _
		  "		<TD width='100' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
		end if
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	n_saldo_parcial = 0
	n_saldo_total = 0
	vl_sub_total = 0
	vl_total_geral = 0
	qtde_lojas = 0
    saldo_grupo = 0
    n_ctrl_reg = 0
	qtde_fabricantes = 0
	loja_a = "XXX"
	fabricante_a = "XXXXX"
    grupo_a = "XXXXX"
	
	if blnSaidaExcel then s_nbsp = "" else s_nbsp = "&nbsp;"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU LOJA?
		if Trim("" & r("loja")) <> loja_a then
			if n_reg_total > 0 then 

            ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = "Grupo" then
		            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

			  ' FECHA TABELA DA LOJA ANTERIOR
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD COLSPAN='2' class='MB ME' align='right' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
						"		<TD class='MB' NOWRAP valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MDB' NOWRAP valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				if blnSaidaExcel then
					x="<BR>" & chr(13) & "<BR>" & chr(13)
				else
					x="<BR>" & chr(13)
					end if
				end if

			x = x & cab_table

		  ' INICIA NOVA TABELA P/ A NOVA LOJA
			if Trim("" & r("loja")) <> "" then
			'   QUEBRA POR LOJA APENAS SE HOUVER LOJA
				if blnSaidaExcel then s_bkg_color = "tomato" else s_bkg_color = "azure"
				x = x & "	<TR NOWRAP style='background:" & s_bkg_color & "'>" & chr(13) & _
						"		<TD class='MDB ME' align='center' valign='bottom' colspan='5'><P class='F' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & r("loja") & " - " & ucase(x_loja(r("loja"))) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				end if

			x = x & cab

			n_reg = 0
			n_saldo_parcial = 0
			vl_sub_total = 0
			qtde_lojas = qtde_lojas + 1
			loja_a = Trim("" & r("loja"))
			fabricante_a = "XXXXX"
            grupo_a = "XXXXX"
			end if
			
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then

            ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = "Grupo" then
		            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MB ME' colspan='2' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
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
					"		<TD class='MDB ME' align='center' valign='bottom' colspan='5' style='background: honeydew'><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
			vl_sub_total = 0
            grupo_a = "XXXXXXX"
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

    if rb_tipo_agrupamento = "Produto" then

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDB ME' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

	 '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

	 '> CMV_PV
		vl = r("preco_fabricante")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CMV_PV TOTAL
		vl = r("preco_fabricante")*r("saldo")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

		vl_sub_total = vl_sub_total + vl
		vl_total_geral = vl_total_geral + vl
		
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		n_saldo_total = n_saldo_total + r("saldo")
		
		x = x & "	</TR>" & chr(13)

    elseif rb_tipo_agrupamento = "Grupo" then
            if Trim("" & r("grupo")) <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                end if
                x = x & "	<TR NOWRAP>" & chr(13)
		        x = x & "		<TD class='MDB ME' colspan='5' align='left' style='background-color: #EEE' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)


                redim preserve vt_CodigoGrupo(cont)
                redim preserve vt_DescricaoGrupo(cont)     
                if Trim("" & r("grupo")) = "" then
                    vt_CodigoGrupo(cont) = "-"
                else
                    vt_CodigoGrupo(cont) = Trim("" & r("grupo"))
                end if
                vt_DescricaoGrupo(cont) = Trim("" & r("grupo_descricao"))           

                grupo_a = Trim("" & r("grupo"))    
                saldo_grupo = 0          
                vl_custo_entrada_grupo = 0  
                n_ctrl_reg = n_ctrl_reg + 1
            end if

        x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDB ME' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

	 '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

	 '> CMV_PV
		vl = r("preco_fabricante")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CMV_PV TOTAL
		vl = r("preco_fabricante")*r("saldo")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)


            vl_sub_total = vl_sub_total + vl
		    vl_total_geral = vl_total_geral + vl
		
		    n_saldo_parcial = n_saldo_parcial + r("saldo")
		    n_saldo_total = n_saldo_total + r("saldo")

            saldo_grupo = saldo_grupo + CInt(r("saldo"))
            vl_custo_entrada_grupo = vl_custo_entrada_grupo + vl

        x = x & "	</TR>" & chr(13)

        end if

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop

  ' MOSTRA TOTAL DO ÚLTIMO GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        if n_reg <> 0 then
            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)         
            redim preserve vt_CustoEntrada(cont)           
            vt_QtdeGrupo(cont) = saldo_grupo
            vt_CustoEntrada(cont) = vl_custo_entrada_grupo
        end if
    end if
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MB ME' COLSPAN='2' align='right' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;'>" & "Total:</P></TD>" & chr(13) & _
				"		<TD class='MB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

   ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        redim vt_TotalGrupo(0)
        set vt_TotalGrupo(0) = New cl_QUATRO_COLUNAS
        with vt_TotalGrupo(0)
            .c1 = ""
            .c2 = ""
            .c3 = ""
            .c4 = ""
        end with

        if vt_CodigoGrupo(UBound(vt_CodigoGrupo)) <> "" then
            for cont = 0 to UBound(vt_CodigoGrupo) 
                if Trim(vt_TotalGrupo(UBound(vt_TotalGrupo)).c1) <> "" then
                    redim preserve vt_TotalGrupo(UBound(vt_TotalGrupo)+1)
                    set vt_TotalGrupo(UBound(vt_TotalGrupo)) = New cl_QUATRO_COLUNAS
                end if
                with vt_TotalGrupo(UBound(vt_TotalGrupo))
                    .c1 = vt_CodigoGrupo(cont)
                    .c2 = vt_DescricaoGrupo(cont)
                    .c3 = vt_QtdeGrupo(cont)
                    .c4 = vt_CustoEntrada(cont)
                end with
            next
        end if
        ordena_cl_quatro_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        vl_custo_entrada_grupo = 0
        n_ctrl_reg = 0
        grupo_a = "XXXXXXX"

        if (qtde_fabricantes > 1) Or (qtde_lojas > 1) then
            x = x & _
                "	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='5'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='5'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:azure'>" & chr(13) & _
                "		<TD class='MTBD ME' valign='bottom' colspan='5' NOWRAP><P class='R' style='font-weight:bold;'>TOTAL GERAL POR GRUPO DE PRODUTOS</P></TD>" & chr(13) & _
                "   </TR>" & CHR(13)

        for cont = 0 to UBound(vt_TotalGrupo)
            if vt_TotalGrupo(cont).c1 <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & _
                    "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
                    "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    				"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
                    "	</TR>" & chr(13)
                    saldo_grupo = 0
                    vl_custo_entrada_grupo = 0
                end if
                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MB ME' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            vl_custo_entrada_grupo = vl_custo_entrada_grupo + vt_TotalGrupo(cont).c4
            
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
            "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    		"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
			"	</TR>" & chr(13)

        end if
    end if

	'>	TOTAL GERAL
		if (qtde_lojas > 1) Or (qtde_fabricantes > 1) then
			x = x & "</TABLE>" & chr(13) & _
				"<BR>" & chr(13) & _
				"<BR>" & chr(13) & _
				cab_table & _
				"	<TR style='background:honeydew'>" & chr(13) & _
				"		<TD class='MB ME' width='60' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MB' width='274' align='right' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;'>TOTAL GERAL:</p></TD>" & chr(13) & _
				"		<TD class='MB' width='60' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_total) & "</p></TD>" & chr(13) & _
				"		<TD class='MB' width='100' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MDB' width='100' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_geral) & "</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='5' align='center'><P class='ALERTA'>NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS</P></TD>" & chr(13) & _
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
dim s, s_aux, s_sql, s_nbsp, x, cab_table, cab, fabricante_a, grupo_a, saldo_grupo, n_ctrl_reg
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes, s_where_temp, cont
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_TotalGrupo()


	s_sql = "SELECT" & _
                " t_PRODUTO.grupo," & _
                " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
				" t_ESTOQUE_ITEM.fabricante," & _
				" t_ESTOQUE_ITEM.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html," & _
				" Sum(qtde-qtde_utilizada) AS saldo" & _
			" FROM t_ESTOQUE_ITEM" & _
				" LEFT JOIN t_PRODUTO ON" & _
					" ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
                " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque = t_ESTOQUE.id_estoque)" & _
			" WHERE" & _
				" ((qtde-qtde_utilizada) > 0)"

	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
		end if

    s_where_temp = ""
    if cod_fabricante <> "" then
		s_where_temp = s_where_temp & "(t_ESTOQUE_ITEM.fabricante='" & cod_fabricante & "')"
		end if

    v_fabricantes = split(c_fabricante_multiplo, ", ")
    if c_fabricante_multiplo <> "" then
	    for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			       " (t_ESTOQUE_ITEM.fabricante = '" & v_fabricantes(cont) & "')"
	    next
    	
    end if

    if s_where_temp <> "" then
        s_sql = s_sql & "AND"  
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    s_where_temp = ""
	if c_grupo <> "" then
	    v_grupos = split(c_grupo, ", ")
	    for cont = Lbound(v_grupos) to Ubound(v_grupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.grupo = '" & v_grupos(cont) & "')"
	    next
	    s_sql = s_sql & "AND "
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    if c_potencia_BTU <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.potencia_BTU = " & c_potencia_BTU & ")"
		end if

    if c_ciclo <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.ciclo = '" & c_ciclo & "')"
		end if
	
	if c_posicao_mercado <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.posicao_mercado = '" & c_posicao_mercado & "')"
		end if

    if c_empresa <> "" then
        s_sql = s_sql & _
            " AND (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
	end if

	s_sql = s_sql & _
			" GROUP BY" & _
				" t_ESTOQUE_ITEM.fabricante," & _
                " t_PRODUTO.grupo," & _
				" t_PRODUTO_GRUPO.descricao," & _
				" t_ESTOQUE_ITEM.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html"
    if rb_tipo_agrupamento = "Produto" then
		s_sql = s_sql & _ 
			" ORDER BY" & _
				" t_ESTOQUE_ITEM.fabricante," & _
				" t_ESTOQUE_ITEM.produto," & _
				" descricao," & _
				" descricao_html"
    elseif rb_tipo_agrupamento = "Grupo" then
        s_sql = s_sql & _ 
            " ORDER BY" & _
				" t_ESTOQUE_ITEM.fabricante," & _
                " grupo," & _
				" t_ESTOQUE_ITEM.produto," & _
				" descricao," & _
				" descricao_html"
    end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)

'	Mantendo a padronização com as demais opções de emissão do relatório em Excel (sem bordas no cabeçalho)
	if blnSaidaExcel then
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD width='75' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='480' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
			  "		<TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
	else
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD width='75' valign='bottom' NOWRAP class='MD ME MB'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='480' valign='bottom' NOWRAP class='MD MB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
			  "		<TD width='60' align='right' valign='bottom' NOWRAP class='MDB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
		end if
	
	x = cab_table & cab
	n_reg = 0
	n_saldo_total = 0
	n_saldo_parcial = 0
	qtde_fabricantes = 0
	fabricante_a = "XXXXX"
    grupo_a = "XXXXX"
    saldo_grupo = 0
    n_ctrl_reg = 0
	
	if blnSaidaExcel then s_nbsp = "" else s_nbsp = "&nbsp;"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then

                ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = "Grupo" then
		            x = x & "		<TD class='MB ME' colspan='2' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='2' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				
				if blnSaidaExcel then
					x = x & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='3' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='3'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				else
					x = x & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='3' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
					end if
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MDBE' align='center' valign='bottom' colspan='3' style='background: honeydew'><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
            grupo_a = "XXXXXXX"
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1

    if rb_tipo_agrupamento = "Produto" then

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

 	 '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

		n_saldo_total = n_saldo_total + r("saldo")
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		
		x = x & "	</TR>" & chr(13)

            elseif rb_tipo_agrupamento = "Grupo" then
            if Trim("" & r("grupo")) <> grupo_a then
                if n_ctrl_reg > 0 then
		            x = x & "		<TD class='MB ME' colspan='2' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    cont = cont + 1
                end if
                x = x & "	<TR NOWRAP>" & chr(13)
		        x = x & "		<TD class='MDB ME' colspan='3' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)

                redim preserve vt_CodigoGrupo(cont)
                redim preserve vt_DescricaoGrupo(cont)     
                if Trim("" & r("grupo")) = "" then
                    vt_CodigoGrupo(cont) = "-"
                else
                    vt_CodigoGrupo(cont) = Trim("" & r("grupo"))
                end if
                vt_DescricaoGrupo(cont) = Trim("" & r("grupo_descricao"))           

                grupo_a = Trim("" & r("grupo"))    
                saldo_grupo = 0            
                n_ctrl_reg = n_ctrl_reg + 1
            end if

            x = x & "   <TR NOWRAP>" & chr(13)

    	 '> PRODUTO
		    x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

 	     '> SALDO
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)


            x = x & "   </TR>" & chr(13)

            saldo_grupo = saldo_grupo + CInt(r("saldo"))
            n_saldo_parcial = n_saldo_parcial + r("saldo")
		    n_saldo_total = n_saldo_total + r("saldo")

        end if


		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop

    ' MOSTRA TOTAL DO ÚLTIMO GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        if n_reg <> 0 then
            x = x & "		<TD class='MB ME' colspan='2' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)                    
            vt_QtdeGrupo(cont) = saldo_grupo
        end if
    end if
		
  ' MOSTRA TOTAL
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO FORNECEDOR
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='2' align='right' valign='bottom' class='MEB'><P class='Cd' style='font-weight:bold;'>Total:</P></TD>" & chr(13) & _
				"		<TD valign='bottom' class='MDB'><P class='Cd'  style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

       ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        redim vt_TotalGrupo(0)
        set vt_TotalGrupo(0) = New cl_TRES_COLUNAS
        with vt_TotalGrupo(0)
            .c1 = ""
            .c2 = ""
            .c3 = ""
        end with

        if vt_CodigoGrupo(UBound(vt_CodigoGrupo)) <> "" then
            for cont = 0 to UBound(vt_CodigoGrupo) 
                if Trim(vt_TotalGrupo(UBound(vt_TotalGrupo)).c1) <> "" then
                    redim preserve vt_TotalGrupo(UBound(vt_TotalGrupo)+1)
                    set vt_TotalGrupo(UBound(vt_TotalGrupo)) = New cl_TRES_COLUNAS
                end if
                with vt_TotalGrupo(UBound(vt_TotalGrupo))
                    .c1 = vt_CodigoGrupo(cont)
                    .c2 = vt_DescricaoGrupo(cont)
                    .c3 = vt_QtdeGrupo(cont)
                end with
            next
        end if
        ordena_cl_tres_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        n_ctrl_reg = 0
        grupo_a = "XXXXXXX"

        if (qtde_fabricantes > 1) then
            x = x & _
                "	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='3'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='3'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:azure'>" & chr(13) & _
                "		<TD class='MTBD ME' valign='bottom' colspan='3' NOWRAP><P class='R' style='font-weight:bold;'>TOTAL GERAL POR GRUPO DE PRODUTOS</P></TD>" & chr(13) & _
                "   </TR>" & CHR(13)

        for cont = 0 to UBound(vt_TotalGrupo)
            if vt_TotalGrupo(cont).c1 <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & _
                    "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
				    "	</TR>" & chr(13)
                    saldo_grupo = 0
                end if
                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MDB ME' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
			"	</TR>" & chr(13)

        end if
    end if
		
		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='3' class='MC'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='2' NOWRAP class='MC MEB' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>" & "TOTAL GERAL:" & "</P></TD>" & chr(13) & _
					"		<TD NOWRAP class='MC MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='3' align='center'><P class='ALERTA'>NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS</P></TD>" & chr(13) & _
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
dim s, s_aux, s_sql, s_nbsp, x, cab_table, cab, fabricante_a, grupo_a, saldo_grupo, vl_custo_entrada_grupo, n_ctrl_reg
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes, s_where_temp, cont
dim vl, vl_total_geral, vl_sub_total
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_CustoEntrada(), vt_TotalGrupo()

'	IMPORTANTE: O VALOR ATUAL DE CMV_PV ESTÁ EM T_PRODUTO.PRECO_FABRICANTE
'	==========  O HISTÓRICO DO VALOR DE CMV_PV ESTÁ EM T_PEDIDO_ITEM.PRECO_FABRICANTE (E T_PEDIDO_ITEM_DEVOLVIDO.PRECO_FABRICANTE)
'				O HISTÓRICO DO CUSTO REAL PAGO AO FABRICANTE ESTÁ EM T_ESTOQUE_ITEM.VL_CUSTO2

	s_sql = "SELECT" & _
                " t_PRODUTO.grupo AS grupo," & _
                " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
				" t_ESTOQUE_ITEM.fabricante," & _
				" t_ESTOQUE_ITEM.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html," & _
				" Sum(qtde-qtde_utilizada) AS saldo," & _
				" Sum((qtde-qtde_utilizada)*t_PRODUTO.preco_fabricante) AS preco_total" & _
			" FROM t_ESTOQUE_ITEM" & _
				" LEFT JOIN t_PRODUTO ON" & _
					" ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
                " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque = t_ESTOQUE.id_estoque)" & _
			" WHERE" & _
				" ((qtde-qtde_utilizada) > 0)"
	
	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
		end if

    s_where_temp = ""
    if cod_fabricante <> "" then
		s_where_temp = s_where_temp & "(t_ESTOQUE_ITEM.fabricante='" & cod_fabricante & "')"
		end if

    v_fabricantes = split(c_fabricante_multiplo, ", ")
    if c_fabricante_multiplo <> "" then
	    for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			       " (t_ESTOQUE_ITEM.fabricante = '" & v_fabricantes(cont) & "')"
	    next
    	
    end if

    if s_where_temp <> "" then
        s_sql = s_sql & "AND"  
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    s_where_temp = ""
	if c_grupo <> "" then
	    v_grupos = split(c_grupo, ", ")
	    for cont = Lbound(v_grupos) to Ubound(v_grupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.grupo = '" & v_grupos(cont) & "')"
	    next
	    s_sql = s_sql & "AND "
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    if c_potencia_BTU <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.potencia_BTU = " & c_potencia_BTU & ")"
		end if

    if c_ciclo <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.ciclo = '" & c_ciclo & "')"
		end if
	
	if c_posicao_mercado <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.posicao_mercado = '" & c_posicao_mercado & "')"
		end if

    if c_empresa <> "" then
        s_sql = s_sql & _
            " AND (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
	end if

	s_sql = s_sql & _
			" GROUP BY" & _
				" t_ESTOQUE_ITEM.fabricante," & _
                " t_PRODUTO.grupo," & _
                " t_PRODUTO_GRUPO.descricao," & _
				" t_ESTOQUE_ITEM.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html"
    if rb_tipo_agrupamento = "Produto" then
		        s_sql = s_sql & _ 
			    " ORDER BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " t_PRODUTO.descricao," & _
				    " descricao_html"
    elseif rb_tipo_agrupamento = "Grupo" then
                s_sql = s_sql & _ 
                " ORDER BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
                    " t_PRODUTO.grupo," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " t_PRODUTO.descricao," & _
				    " descricao_html"
    end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
'	Mantendo a padronização com as demais opções de emissão do relatório em Excel (sem bordas no cabeçalho)
	if blnSaidaExcel then
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD width='75' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='350' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
			  "		<TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA<br style='mso-data-placement:same-cell;' />UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA<br style='mso-data-placement:same-cell;' />TOTAL</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
	else
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD width='60' valign='bottom' NOWRAP class='MDBE'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='274' valign='bottom' class='MD MB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
			  "		<TD width='60' align='right' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom' NOWRAP class='MDB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA TOTAL</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
		end if
	
	x = cab_table & cab
	n_reg = 0
	n_saldo_total = 0
	n_saldo_parcial = 0
	vl_total_geral = 0
	vl_sub_total = 0
    saldo_grupo = 0
    n_ctrl_reg = 0
	qtde_fabricantes = 0
	fabricante_a = "XXXXX"
    grupo_a = "XXXXX"
	
	if blnSaidaExcel then s_nbsp = "" else s_nbsp = "&nbsp;"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then

                ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = "Grupo" then
		            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='2' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				
				if blnSaidaExcel then
					x = x & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='5' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='5'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				else
					x = x & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='5' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
					end if
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MDBE' align='center' valign='bottom' colspan='5' style='background: honeydew'><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
			vl_sub_total = 0
            grupo_a = "XXXXXXX"
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1

    if rb_tipo_agrupamento = "Produto" then

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

 	 '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)
	
	 '> CUSTO DE ENTRADA UNITÁRIO (MÉDIA)
		if r("saldo") = 0 then vl = 0 else vl = r("preco_total")/r("saldo")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA TOTAL
		vl = r("preco_total")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)
		
		vl_total_geral = vl_total_geral + vl
		vl_sub_total = vl_sub_total + vl
		
		n_saldo_total = n_saldo_total + r("saldo")
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		
		x = x & "	</TR>" & chr(13)

    elseif rb_tipo_agrupamento = "Grupo" then
            if Trim("" & r("grupo")) <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                end if
                x = x & "	<TR NOWRAP>" & chr(13)
		        x = x & "		<TD class='MDB ME' colspan='5' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)


                redim preserve vt_CodigoGrupo(cont)
                redim preserve vt_DescricaoGrupo(cont)     
                if Trim("" & r("grupo")) = "" then
                    vt_CodigoGrupo(cont) = "-"
                else
                    vt_CodigoGrupo(cont) = Trim("" & r("grupo"))
                end if
                vt_DescricaoGrupo(cont) = Trim("" & r("grupo_descricao"))           

                grupo_a = Trim("" & r("grupo"))    
                saldo_grupo = 0          
                vl_custo_entrada_grupo = 0  
                n_ctrl_reg = n_ctrl_reg + 1
            end if

            x = x & "	<TR NOWRAP>" & chr(13)

	     '> PRODUTO
		    x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

 	     '> SALDO
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)
	
	     '> CUSTO DE ENTRADA UNITÁRIO (MÉDIA)
		    if r("saldo") = 0 then vl = 0 else vl = r("preco_total")/r("saldo")
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	     '> CUSTO DE ENTRADA TOTAL
		    vl = r("preco_total")
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)
		

            saldo_grupo = saldo_grupo + CInt(r("saldo"))
            vl_custo_entrada_grupo = vl_custo_entrada_grupo + r("preco_total")

            n_saldo_parcial = n_saldo_parcial + r("saldo")
		    n_saldo_total = n_saldo_total + r("saldo")
            vl_sub_total = vl_sub_total + vl
		    vl_total_geral = vl_total_geral + vl

            x = x & "	</TR>" & chr(13)
        end if


		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop

   ' MOSTRA TOTAL DO ÚLTIMO GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        if n_reg <> 0 then
            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)         
            redim preserve vt_CustoEntrada(cont)           
            vt_QtdeGrupo(cont) = saldo_grupo
            vt_CustoEntrada(cont) = vl_custo_entrada_grupo
        end if
    end if
		
  ' MOSTRA TOTAL
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO FORNECEDOR
		x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='2' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MdB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

       ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        redim vt_TotalGrupo(0)
        set vt_TotalGrupo(0) = New cl_QUATRO_COLUNAS
        with vt_TotalGrupo(0)
            .c1 = ""
            .c2 = ""
            .c3 = ""
            .c4 = ""
        end with

        if vt_CodigoGrupo(UBound(vt_CodigoGrupo)) <> "" then
            for cont = 0 to UBound(vt_CodigoGrupo) 
                if Trim(vt_TotalGrupo(UBound(vt_TotalGrupo)).c1) <> "" then
                    redim preserve vt_TotalGrupo(UBound(vt_TotalGrupo)+1)
                    set vt_TotalGrupo(UBound(vt_TotalGrupo)) = New cl_QUATRO_COLUNAS
                end if
                with vt_TotalGrupo(UBound(vt_TotalGrupo))
                    .c1 = vt_CodigoGrupo(cont)
                    .c2 = vt_DescricaoGrupo(cont)
                    .c3 = vt_QtdeGrupo(cont)
                    .c4 = vt_CustoEntrada(cont)
                end with
            next
        end if
        ordena_cl_quatro_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        vl_custo_entrada_grupo = 0
        n_ctrl_reg = 0
        grupo_a = "XXXXXXX"

        if (qtde_fabricantes > 1) then
            x = x & _
                "	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='5'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='5'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:azure'>" & chr(13) & _
                "		<TD class='MTBD ME' valign='bottom' colspan='5' NOWRAP><P class='R' style='font-weight:bold;'>TOTAL GERAL POR GRUPO DE PRODUTOS</P></TD>" & chr(13) & _
                "   </TR>" & CHR(13)

        for cont = 0 to UBound(vt_TotalGrupo)
            if vt_TotalGrupo(cont).c1 <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & _
                    "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
                    "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    				"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
                    "	</TR>" & chr(13)
                    saldo_grupo = 0
                    vl_custo_entrada_grupo = 0
                end if
                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MB ME' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            vl_custo_entrada_grupo = vl_custo_entrada_grupo + vt_TotalGrupo(cont).c4
            
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
            "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    		"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
			"	</TR>" & chr(13)

        end if
    end if

		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='5' class='MC'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='2' align='right' valign='bottom' NOWRAP class='MC MEB'><P class='Cd' style='font-weight:bold;'>TOTAL GERAL:</P></TD>" & chr(13) & _
					"		<TD NOWRAP class='MC MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
					"		<TD class='MC MB'>&nbsp;</TD>" & chr(13) & _
					"		<TD NOWRAP class='MC MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_geral) & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='5' align='center'><P class='ALERTA'>NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS</P></TD>" & chr(13) & _
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
dim s, s_aux, s_sql, s_nbsp, x, cab_table, cab, fabricante_a, grupo_a, saldo_grupo, vl_custo_entrada_grupo, n_ctrl_reg
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes, s_where_temp, cont
dim vl, vl_total_geral, vl_sub_total
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_CustoEntrada(), vt_TotalGrupo()

'	IMPORTANTE: O VALOR ATUAL DE CMV_PV ESTÁ EM T_PRODUTO.PRECO_FABRICANTE
'	==========  O HISTÓRICO DO VALOR DE CMV_PV ESTÁ EM T_PEDIDO_ITEM.PRECO_FABRICANTE (E T_PEDIDO_ITEM_DEVOLVIDO.PRECO_FABRICANTE)
'				O HISTÓRICO DO CUSTO REAL PAGO AO FABRICANTE ESTÁ EM T_ESTOQUE_ITEM.VL_CUSTO2

	s_sql = "SELECT" & _
                " t_PRODUTO.grupo AS grupo," & _
                " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
				" t_ESTOQUE_ITEM.fabricante," & _
				" t_ESTOQUE_ITEM.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html," & _
				" Sum(qtde-qtde_utilizada) AS saldo," & _
				" t_PRODUTO.preco_fabricante" & _
			" FROM t_ESTOQUE_ITEM" & _
				" LEFT JOIN t_PRODUTO ON" & _
					" ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
                " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque = t_ESTOQUE.id_estoque)" & _
			" WHERE" & _
				" ((qtde-qtde_utilizada) > 0)"
	
	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
		end if

    s_where_temp = ""
    if cod_fabricante <> "" then
		s_where_temp = s_where_temp & "(t_ESTOQUE_ITEM.fabricante='" & cod_fabricante & "')"
		end if

    v_fabricantes = split(c_fabricante_multiplo, ", ")
    if c_fabricante_multiplo <> "" then
	    for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			       " (t_ESTOQUE_ITEM.fabricante = '" & v_fabricantes(cont) & "')"
	    next
    	
    end if

    if s_where_temp <> "" then
        s_sql = s_sql & "AND"  
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    s_where_temp = ""
	if c_grupo <> "" then
	    v_grupos = split(c_grupo, ", ")
	    for cont = Lbound(v_grupos) to Ubound(v_grupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.grupo = '" & v_grupos(cont) & "')"
	    next
	    s_sql = s_sql & "AND "
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    if c_potencia_BTU <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.potencia_BTU = " & c_potencia_BTU & ")"
		end if

    if c_ciclo <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.ciclo = '" & c_ciclo & "')"
		end if
	
	if c_posicao_mercado <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.posicao_mercado = '" & c_posicao_mercado & "')"
		end if

    if c_empresa <> "" then
        s_sql = s_sql & _
            " AND (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
	end if

	s_sql = s_sql & _
			" GROUP BY" & _
				" t_ESTOQUE_ITEM.fabricante," & _
                " t_PRODUTO.grupo," & _
                " t_PRODUTO_GRUPO.descricao," & _
				" t_ESTOQUE_ITEM.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html," & _
				" t_PRODUTO.preco_fabricante"
    if rb_tipo_agrupamento = "Produto" then
		        s_sql = s_sql & _
			    " ORDER BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " t_PRODUTO.descricao," & _
				    " descricao_html," & _
				    " t_PRODUTO.preco_fabricante"
    elseif rb_tipo_agrupamento = "Grupo" then
                s_sql = s_sql & _
                " ORDER BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
                    " t_PRODUTO.grupo," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " t_PRODUTO.descricao," & _
				    " descricao_html," & _
				    " t_PRODUTO.preco_fabricante"
    end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
	'mantendo a padronização com as demais opções de emissão do relatório em Excel (sem bordas no cabeçalho)
	if blnSaidaExcel then
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD width='75' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='350' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
			  "		<TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA<br style='mso-data-placement:same-cell;' />UNITÁRIO</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA<br style='mso-data-placement:same-cell;' />TOTAL</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
	else
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD width='60' valign='bottom' NOWRAP class='MDBE'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='274' valign='bottom' NOWRAP class='MD MB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
			  "		<TD width='60' align='right' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA UNITÁRIO</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom' NOWRAP class='MDB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA TOTAL</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
		end if
	
	x = cab_table & cab
	n_reg = 0
	n_saldo_total = 0
	n_saldo_parcial = 0
	vl_total_geral = 0
	vl_sub_total = 0
    saldo_grupo = 0
    n_ctrl_reg = 0
	qtde_fabricantes = 0
	fabricante_a = "XXXXX"
    grupo_a = "XXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then

                ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = "Grupo" then
		            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='2' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				
				if blnSaidaExcel then
					x = x & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='5' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='5'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				else
					x = x & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='5' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
					end if
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MDBE' align='center' valign='bottom' colspan='5' style='background: honeydew'><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
			vl_sub_total = 0
            grupo_a = "XXXXXXX"
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1

    if rb_tipo_agrupamento = "Produto" then

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

	 '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)
	
	 '> CMV_PV
		vl = r("preco_fabricante")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CMV_PV TOTAL
		vl = r("preco_fabricante")*r("saldo")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)
		
		vl_total_geral = vl_total_geral + vl
		vl_sub_total = vl_sub_total + vl

		n_saldo_total = n_saldo_total + r("saldo")
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		
		x = x & "	</TR>" & chr(13)

        elseif rb_tipo_agrupamento = "Grupo" then
            if Trim("" & r("grupo")) <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                end if
                x = x & "	<TR NOWRAP>" & chr(13)
		        x = x & "		<TD class='MDB ME' colspan='5' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)


                redim preserve vt_CodigoGrupo(cont)
                redim preserve vt_DescricaoGrupo(cont)     
                if Trim("" & r("grupo")) = "" then
                    vt_CodigoGrupo(cont) = "-"
                else
                    vt_CodigoGrupo(cont) = Trim("" & r("grupo"))
                end if
                vt_DescricaoGrupo(cont) = Trim("" & r("grupo_descricao"))           

                grupo_a = Trim("" & r("grupo"))    
                saldo_grupo = 0          
                vl_custo_entrada_grupo = 0  
                n_ctrl_reg = n_ctrl_reg + 1
            end if

            x = x & "	<TR NOWRAP>" & chr(13)

	      '> PRODUTO
		    x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

	     '> SALDO
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)
	
	     '> CMV_PV
		    vl = r("preco_fabricante")
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	     '> CMV_PV TOTAL
		    vl = r("preco_fabricante")*r("saldo")
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)
		

            vl_sub_total = vl_sub_total + vl
		    vl_total_geral = vl_total_geral + vl
		
		    n_saldo_parcial = n_saldo_parcial + r("saldo")
		    n_saldo_total = n_saldo_total + r("saldo")

            saldo_grupo = saldo_grupo + CInt(r("saldo"))
            vl_custo_entrada_grupo = vl_custo_entrada_grupo + vl

            x = x & "	</TR>" & chr(13)

        end if

		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop

   ' MOSTRA TOTAL DO ÚLTIMO GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        if n_reg <> 0 then
            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='2' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)         
            redim preserve vt_CustoEntrada(cont)           
            vt_QtdeGrupo(cont) = saldo_grupo
            vt_CustoEntrada(cont) = vl_custo_entrada_grupo
        end if
    end if
		
  ' MOSTRA TOTAL
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO FORNECEDOR
		x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='2' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
	
   ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        redim vt_TotalGrupo(0)
        set vt_TotalGrupo(0) = New cl_QUATRO_COLUNAS
        with vt_TotalGrupo(0)
            .c1 = ""
            .c2 = ""
            .c3 = ""
            .c4 = ""
        end with

        if vt_CodigoGrupo(UBound(vt_CodigoGrupo)) <> "" then
            for cont = 0 to UBound(vt_CodigoGrupo) 
                if Trim(vt_TotalGrupo(UBound(vt_TotalGrupo)).c1) <> "" then
                    redim preserve vt_TotalGrupo(UBound(vt_TotalGrupo)+1)
                    set vt_TotalGrupo(UBound(vt_TotalGrupo)) = New cl_QUATRO_COLUNAS
                end if
                with vt_TotalGrupo(UBound(vt_TotalGrupo))
                    .c1 = vt_CodigoGrupo(cont)
                    .c2 = vt_DescricaoGrupo(cont)
                    .c3 = vt_QtdeGrupo(cont)
                    .c4 = vt_CustoEntrada(cont)
                end with
            next
        end if
        ordena_cl_quatro_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        vl_custo_entrada_grupo = 0
        n_ctrl_reg = 0
        grupo_a = "XXXXXXX"

        if (qtde_fabricantes > 1) then
            x = x & _
                "	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='5'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='5'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:azure'>" & chr(13) & _
                "		<TD class='MTBD ME' valign='bottom' colspan='5' NOWRAP><P class='R' style='font-weight:bold;'>TOTAL GERAL POR GRUPO DE PRODUTOS</P></TD>" & chr(13) & _
                "   </TR>" & CHR(13)

        for cont = 0 to UBound(vt_TotalGrupo)
            if vt_TotalGrupo(cont).c1 <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & _
                    "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
                    "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    				"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
                    "	</TR>" & chr(13)
                    saldo_grupo = 0
                    vl_custo_entrada_grupo = 0
                end if
                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MB ME' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            vl_custo_entrada_grupo = vl_custo_entrada_grupo + vt_TotalGrupo(cont).c4
            
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
            "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    		"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
			"	</TR>" & chr(13)

        end if
    end if

		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='5' class='MC'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='2' align='right' valign='bottom' NOWRAP class='MC MEB'><P class='Cd' style='font-weight:bold;'>TOTAL GERAL:</P></TD>" & chr(13) & _
					"		<TD NOWRAP class='MC MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
					"		<TD class='MC MB'>&nbsp;</TD>" & chr(13) & _
					"		<TD NOWRAP class='MC MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_geral) & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='5' align='center'><P class='ALERTA'>NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	
end sub

' ________________________________________________________
' CONSULTA ESTOQUE DETALHE SINTETICO COM CÓDIGO UNIFICADO
'
sub consulta_estoque_detalhe_sintetico_codigo_unificado
const LargColOrdemServico = 110
dim LargColProduto, LargColDescricao
dim r
dim s, i, s_aux, s_bkg_color, s_nbsp, s_sql, s_lista_OS, s_chave_OS, s_num_OS_tela, loja_a, x, cab_table, cab, fabricante_a, grupo_a, saldo_grupo, n_ctrl_reg
dim n_reg, n_reg_total, n_saldo_parcial, n_saldo_total, qtde_lojas, qtde_fabricantes
dim intQtdeColunasColSpanSubTotal, intQtdeColunasColSpanTotalGeral, intQtdeTotalColunasColSpan, s_where_temp, cont, blnPularProdutoComposto, qtde_estoque_composto, qtde_estoque_aux
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_TotalGrupo()
dim vt_Loja(), vt_GrupoCodigo(), vt_GrupoDescricao(), vt_Fabricante(), vt_Produto(), vt_ProdutoDescricao(), vt_ProdutoDescricaoHtml(), vt_Saldo(), vRelat()
dim loja_composto, grupo_codigo_composto, grupo_descricao_composto, strCampoOrdenacao

	intQtdeTotalColunasColSpan = 3
	intQtdeColunasColSpanSubTotal = 2
	intQtdeColunasColSpanTotalGeral = 1
	LargColProduto = 75
	LargColDescricao = 480
	if rb_estoque = ID_ESTOQUE_DANIFICADOS and rb_tipo_agrupamento <> "Grupo" then
		intQtdeTotalColunasColSpan = 4
		intQtdeColunasColSpanSubTotal = 3
		intQtdeColunasColSpanTotalGeral = 2
		LargColDescricao = 370
    end if
    
'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT" & _
                " grupo," & _
                " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
				" loja," & _
				" CONVERT(smallint,loja) AS numero_loja," & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html," & _
				" Sum(qtde) AS saldo" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
                " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
			" WHERE" & _
				" (anulado_status=0)" & _
                " AND (t_ESTOQUE_MOVIMENTO.produto NOT IN (SELECT DISTINCT produto_composto AS produto FROM t_EC_PRODUTO_COMPOSTO))" & _
			    " AND (t_ESTOQUE_MOVIMENTO.produto NOT IN (SELECT DISTINCT produto_item AS produto FROM t_EC_PRODUTO_COMPOSTO_ITEM))" & _  
				" AND (estoque='" & rb_estoque & "')"
	
	if cod_produto <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.produto='" & cod_produto & "')"
		end if

    s_where_temp = ""
    if cod_fabricante <> "" then
		s_where_temp = s_where_temp & "(t_ESTOQUE_MOVIMENTO.fabricante='" & cod_fabricante & "')"
		end if

    v_fabricantes = split(c_fabricante_multiplo, ", ")
    if c_fabricante_multiplo <> "" then
	    for i = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			       " (t_ESTOQUE_MOVIMENTO.fabricante = '" & v_fabricantes(i) & "')"
	    next
    	
    end if

    if s_where_temp <> "" then
        s_sql = s_sql & "AND"  
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    s_where_temp = ""
	if c_grupo <> "" then
	    v_grupos = split(c_grupo, ", ")
	    for i = Lbound(v_grupos) to Ubound(v_grupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.grupo = '" & v_grupos(i) & "')"
	    next
	    s_sql = s_sql & "AND "
	    s_sql = s_sql & "(" & s_where_temp & ")"
    end if

    if c_potencia_BTU <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.potencia_BTU = " & c_potencia_BTU & ")"
		end if

    if c_ciclo <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.ciclo = '" & c_ciclo & "')"
		end if
	
	if c_posicao_mercado <> "" then
		s_sql = s_sql & _
			" AND (t_PRODUTO.posicao_mercado = '" & c_posicao_mercado & "')"
		end if
	
    if c_empresa <> "" then
        s_sql = s_sql & _
            " AND (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
	end if

	s_sql = s_sql & _
			" GROUP BY" & _
				" loja," & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
                " grupo," & _
				" t_PRODUTO_GRUPO.descricao," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html"

    set r = cn.execute(s_sql)
	do while Not r.Eof
        redim preserve vt_Loja(cont)
        redim preserve vt_GrupoCodigo(cont)
        redim preserve vt_GrupoDescricao(cont)
        redim preserve vt_Fabricante(cont)
        redim preserve vt_Produto(cont)
        redim preserve vt_ProdutoDescricao(cont)
        redim preserve vt_ProdutoDescricaoHtml(cont)
        redim preserve vt_Saldo(cont)

        vt_Loja(cont) = r("loja")
        vt_GrupoCodigo(cont) = r("grupo")
        vt_GrupoDescricao(cont) = r("grupo_descricao")
        vt_Fabricante(cont) = r("fabricante")
        vt_Produto(cont) = r("produto")
        vt_ProdutoDescricao(cont) = r("descricao")
        vt_ProdutoDescricaoHtml(cont) = r("descricao_html")
        vt_Saldo(cont) = r("saldo")

        
        cont = cont + 1
        r.MoveNext
    loop
    if r.State <> 0 then r.Close
    set r=nothing

    s_sql = "SELECT" & _
					" tECPC.fabricante_composto," & _
					" tECPC.produto_composto," & _
					" tECPC.descricao," & _
					" tF.nome AS nome_fabricante" & _
				" FROM t_EC_PRODUTO_COMPOSTO tECPC" & _
					" LEFT JOIN t_FABRICANTE tF ON (tECPC.fabricante_composto = tF.fabricante)" & _
				" ORDER BY" & _
					" tECPC.fabricante_composto," & _
					" tECPC.produto_composto"

    if cod_produto <> "" then
		s_sql = s_sql & " AND (t_EC_PRODUTO_COMPOSTO.produto_composto='" & cod_produto & "')"
		end if

    if cod_fabricante <> "" then
		s_where_temp = s_where_temp & "(t_EC_PRODUTO_COMPOSTO.fabricante_composto='" & cod_fabricante & "')"
		end if

    set r = cn.execute(s_sql)
	do while Not r.Eof
		blnPularProdutoComposto = False
        qtde_estoque_composto = -1

        s_sql = "SELECT " & _
					" fabricante_item," & _
					" produto_item," & _
					" qtde" & _
			        " FROM t_EC_PRODUTO_COMPOSTO_ITEM"  & _
                    " LEFT JOIN t_PRODUTO tP ON ((tP.fabricante=t_EC_PRODUTO_COMPOSTO_ITEM.fabricante_item) AND (tP.produto=t_EC_PRODUTO_COMPOSTO_ITEM.produto_item))" & _
                    " WHERE" & _
					" (fabricante_composto = '" & Trim("" & r("fabricante_composto")) & "')" & _
					" AND (produto_composto = '" & Trim("" & r("produto_composto")) & "')" 

        if cod_fabricante <> "" then
	            s_sql = s_sql & " AND "
	            s_sql = s_sql & " (" & s_where_compostos & ") "
            end if

            if cod_produto <> "" then
		        s_sql = s_sql & " AND (produto_composto='" & cod_produto & "')"
	        end if 

        s_where_temp = ""
        v_fabricantes = split(c_fabricante_multiplo, ", ")
        if c_fabricante_multiplo <> "" then
	        for i = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	            if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		        s_where_temp = s_where_temp & _
			           " (tP.fabricante = '" & v_fabricantes(i) & "')"
	        next
    	
        end if

        if s_where_temp <> "" then
            s_sql = s_sql & "AND"  
	        s_sql = s_sql & "(" & s_where_temp & ")"
        end if

        s_where_temp = ""
	    if c_grupo <> "" then
	        v_grupos = split(c_grupo, ", ")
	        for i = Lbound(v_grupos) to Ubound(v_grupos)
	            if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		        s_where_temp = s_where_temp & _
			        " (tP.grupo = '" & v_grupos(i) & "')"
	        next
	        s_sql = s_sql & "AND "
	        s_sql = s_sql & "(" & s_where_temp & ")"
        end if

        if c_potencia_BTU <> "" then
		    s_sql = s_sql & _
			    " AND (tP.potencia_BTU = " & c_potencia_BTU & ")"
		    end if

        if c_ciclo <> "" then
		    s_sql = s_sql & _
			    " AND (tP.ciclo = '" & c_ciclo & "')"
		    end if
	
	    if c_posicao_mercado <> "" then
		    s_sql = s_sql & _
			    " AND (tP.posicao_mercado = '" & c_posicao_mercado & "')"
		    end if        

	    s_sql = s_sql & " ORDER BY" & _
					    " fabricante_item," & _
					    " produto_item"    
    
	    if tPCI.State <> 0 then tPCI.Close
        tPCI.Open s_sql, cn  
        do while not tPCI.Eof

            s_sql = " SELECT" & _
				" grupo," & _
                " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
				" loja," & _
				" CONVERT(smallint,loja) AS numero_loja," & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" descricao_html," & _
				" Sum(qtde) AS saldo" & _
						" FROM t_ESTOQUE_MOVIMENTO" & _
							" INNER JOIN t_PRODUTO ON (t_ESTOQUE_MOVIMENTO.fabricante = t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto = t_PRODUTO.produto)" & _
                            " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
                            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
						" WHERE " & _
                        " (anulado_status=0)" & _
                        " AND (estoque='" & rb_estoque & "')" & _
                        " AND (t_ESTOQUE_MOVIMENTO.fabricante = '" & Trim("" & tPCI("fabricante_item")) & "')" & _
                       	" AND (t_ESTOQUE_MOVIMENTO.produto = '" & Trim("" & tPCI("produto_item")) & "') " 

            if c_empresa <> "" then
                s_sql = s_sql & _
                    " AND (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
	        end if

            s_sql = s_sql & _
			    " GROUP BY" & _
				    " loja," & _
				    " t_ESTOQUE_MOVIMENTO.fabricante," & _
                    " grupo," & _
				    " t_PRODUTO_GRUPO.descricao," & _
				    " t_ESTOQUE_MOVIMENTO.produto," & _
				    " t_PRODUTO.descricao," & _
				    " descricao_html"

            s_sql = " SELECT " & _
							"*" & _
						" FROM (" & s_sql & ") t" & _
						" ORDER BY" & _
							" fabricante," & _
							" produto"    

            if t.State <> 0 then t.Close
				t.Open s_sql, cn
				if t.Eof then
				    blnPularProdutoComposto = true                                  
				else
                    qtde_estoque_aux = t("saldo") / tPCI("qtde")
                    loja_composto = Trim("" & t("loja"))
                    grupo_codigo_composto = Trim("" & t("grupo"))
                    grupo_descricao_composto = Trim("" & t("grupo_descricao"))

                    if qtde_estoque_composto = -1 then
                        qtde_estoque_composto = qtde_estoque_aux
                    else 
                        if qtde_estoque_aux < qtde_estoque_composto then
                        qtde_estoque_composto = qtde_estoque_aux                            
                        end if
                    end if 

                end if
            if blnPularProdutoComposto then exit do
            tPCI.MoveNext
		loop
            
            if qtde_estoque_composto > 0 then
                if Not blnPularProdutoComposto then
                    redim preserve vt_Loja(cont)
                    redim preserve vt_GrupoCodigo(cont)
                    redim preserve vt_GrupoDescricao(cont)
                    redim preserve vt_Fabricante(cont)
                    redim preserve vt_Produto(cont)
                    redim preserve vt_ProdutoDescricao(cont)
                    redim preserve vt_ProdutoDescricaoHtml(cont)
                    redim preserve vt_Saldo(cont)

                    vt_Loja(cont) = loja_composto
                    vt_GrupoCodigo(cont) = grupo_codigo_composto
                    vt_GrupoDescricao(cont) = grupo_descricao_composto
                    vt_Fabricante(cont) = Trim("" & r("fabricante_composto"))	
                    vt_Produto(cont) = Trim("" & r("produto_composto")) 
                    vt_ProdutoDescricao(cont) = Trim("" & r("descricao"))
                    vt_ProdutoDescricaoHtml(cont) = Trim("" & r("descricao"))
                    vt_Saldo(cont) = qtde_estoque_composto


                    cont = cont + 1
                end if        
            end if

        r.MoveNext
    loop


    redim vRelat(0)
    set vRelat(0)=New cl_VINTE_COLUNAS
    with vRelat(0)
        .CampoOrdenacao = ""
        .c1 = ""
		.c2 = ""
		.c3 = ""
		.c4 = ""
		.c5 = ""
        .c6 = ""
        .c7 = ""
        .c8 = ""
        .c9 = ""
    end with

    if vt_Produto(UBound(vt_Produto)) <> "" then
        for cont = 0 to UBound(vt_Produto)
            if Trim(vRelat(UBound(vRelat)).c5) <> "" then
                redim preserve vRelat(UBound(vRelat)+1)
				set vRelat(UBound(vRelat)) = New cl_VINTE_COLUNAS
            end if
            with vRelat(UBound(vRelat))
                .c1 = vt_Loja(cont)
                .c2 = vt_GrupoCodigo(cont)
                .c3 = vt_GrupoDescricao(cont)
                .c4 = vt_Fabricante(cont)
                .c5 = vt_Produto(cont)
                .c6 = vt_ProdutoDescricao(cont)
                .c7 = vt_ProdutoDescricaoHtml(cont)
                .c8 = vt_Saldo(cont)

                if rb_tipo_agrupamento = "Produto" then
                    .CampoOrdenacao = normaliza_codigo(.c1, 3) & "|" & normaliza_codigo(.c4, 3) & "|" & normaliza_codigo(.c5, 6) & "|" & .c6
                elseif rb_tipo_agrupamento = "Grupo" then
                    .CampoOrdenacao = normaliza_codigo(.c1, 3) & "|" & normaliza_codigo(.c4, 3) & "|" & normaliza_codigo(.c2, 2) & "|" & normaliza_codigo(.c5, 6) & "|" & .c6
                end if
            end with
        next
    end if

  '	ORDENA O VETOR COM RESULTADOS
	ordena_cl_vinte_colunas vRelat, 0, Ubound(vRelat)

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)

	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD width='" & Cstr(LargColProduto) & "' valign='bottom' class='MDBE'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='" & CStr(LargColDescricao) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13)

	if rb_estoque = ID_ESTOQUE_DANIFICADOS and rb_tipo_agrupamento <> "Grupo" then
		cab = cab & _
			  "		<TD width='" & CStr(LargColOrdemServico) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>ORDEM SERVIÇO</P></TD>" & chr(13)
		end if
	
	cab = cab & _
		  "		<TD width='60' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
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
    grupo_a = "XXXXX"
    saldo_grupo = 0
    n_ctrl_reg = 0

	if blnSaidaExcel then s_nbsp = "" else s_nbsp = "&nbsp;"
	
	for cont=Lbound(vRelat) to Ubound(vRelat)
	
	'	MUDOU LOJA?
		if vRelat(cont).c1 <> loja_a then
			if n_reg_total > 0 then 

              ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = "Grupo" then
		            x = x & "		<TD class='MB ME' colspan='2' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

			  ' FECHA TABELA DA LOJA ANTERIOR
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD COLSPAN='" & Cstr(intQtdeColunasColSpanSubTotal) & "' align='right' valign='bottom' class='MEB' NOWRAP><P class='Cd' style='font-weight:bold;'>" & _
						"TOTAL:" & "</P></TD>" & chr(13) & _
						"		<TD NOWRAP class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & _
						"</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				if blnSaidaExcel then
					x="<BR>" & chr(13) & "<BR>" & chr(13)
				else
					x="<BR>" & chr(13)
					end if
				end if

			x = x & cab_table

		  ' INICIA NOVA TABELA P/ A NOVA LOJA
			if vRelat(cont).c1 <> "" then
			'	QUEBRA POR LOJA APENAS SE HOUVER LOJA
				if blnSaidaExcel then s_bkg_color = "tomato" else s_bkg_color = "azure"
				x = x & "	<TR NOWRAP style='background:" & s_bkg_color & "'>" & chr(13) & _
						"		<TD class='MDBE' align='center' valign='bottom' colspan='" & CStr(intQtdeTotalColunasColSpan) & "'><P class='F' style='font-weight:bold;'>" & vRelat(cont).c1 & " - " & ucase(x_loja(vRelat(cont).c1)) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				end if

			x = x & cab

			n_reg = 0
			n_saldo_parcial = 0
			loja_a = vRelat(cont).c1
			qtde_lojas = qtde_lojas + 1
			fabricante_a = "XXXXX"
            grupo_a = "XXXXX"
			end if
			
	'	MUDOU DE FABRICANTE?
		if vRelat(cont).c4 <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then
                ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = "Grupo" then
		            x = x & "		<TD class='MB ME' colspan='2' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' align='right' valign='bottom' colspan='" & CStr(intQtdeColunasColSpanSubTotal) & "'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "' class='MB' valign='bottom'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				end if
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = vRelat(cont).c4
			s = vRelat(cont).c4
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MDBE' align='center' valign='bottom' colspan='" & CStr(intQtdeTotalColunasColSpan) & "' style='background: honeydew'><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
            grupo_a = "XXXXXXX"
			end if
			
	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

        if rb_tipo_agrupamento = "Produto" then

		    x = x & "	<TR NOWRAP>" & chr(13)

	     '> PRODUTO
		    x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & vRelat(cont).c5 & "</P></TD>" & chr(13)

 	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(vRelat(cont).c7) & "</P></TD>" & chr(13)

	    '> ORDEM DE SERVIÇO
		    if rb_estoque = ID_ESTOQUE_DANIFICADOS then
			    s_lista_OS = ""
			    s = "SELECT" & _
					    " id_ordem_servico" & _
				    " FROM t_ESTOQUE_MOVIMENTO" & _
                    " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
				    " WHERE" & _
					    " (anulado_status=0)" & _
					    " AND (estoque='" & ID_ESTOQUE_DANIFICADOS & "')" & _
					    " AND (fabricante='" & vRelat(cont).c4 & "')" & _
					    " AND (produto='" & vRelat(cont).c5 & "')" & _
					    " AND (id_ordem_servico IS NOT NULL)" 

                if c_empresa <> "" then
                    s = s & _
                        " AND (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
	            end if

	            s = s & " ORDER BY" & _
					    " id_ordem_servico"
			    if rs.State <> 0 then rs.Close
			    rs.open s, cn
			    do while Not rs.Eof
				    s_chave_OS = Trim("" & rs("id_ordem_servico"))
				    s_num_OS_tela = formata_num_OS_tela(s_chave_OS)
				    if s_lista_OS <> "" then s_lista_OS = s_lista_OS & ", "
				    if blnSaidaExcel then
					    s_lista_OS = s_lista_OS & s_num_OS_tela
				    else
					    s_lista_OS = s_lista_OS & "<a href='OrdemServico.asp?num_OS=" & s_chave_OS & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "' title='Clique para consultar a Ordem de Serviço'>" & s_num_OS_tela & "</a>"
					    end if
				    rs.MoveNext
				    loop
			
			    if (s_lista_OS = "") And (Not blnSaidaExcel) then s_lista_OS = "&nbsp;"
			    x = x & "		<TD class='MDB' align='left' valign='middle' width='" & CStr(LargColOrdemServico) & "'><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_lista_OS & "</P></TD>" & chr(13)
			    end if
			
	     '> SALDO
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(vRelat(cont).c8) & "</P></TD>" & chr(13)

		    n_saldo_parcial = n_saldo_parcial + vRelat(cont).c8
		    n_saldo_total = n_saldo_total + vRelat(cont).c8
		
		    x = x & "	</TR>" & chr(13)

        elseif rb_tipo_agrupamento = "Grupo" then
            if vRelat(cont).c2 <> grupo_a then
                if n_ctrl_reg > 0 then
		            x = x & "		<TD class='MB ME' colspan='2' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    cont = cont + 1
                end if
                x = x & "	<TR NOWRAP>" & chr(13)
		        x = x & "		<TD class='MDB ME' colspan='3' align='left' style='background-color: #EEE' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & vRelat(cont).c2 & " - " & vRelat(cont).c3 & "</P></TD>" & chr(13)

                redim preserve vt_CodigoGrupo(cont)
                redim preserve vt_DescricaoGrupo(cont)     
                vt_CodigoGrupo(cont) = vRelat(cont).c2
                vt_DescricaoGrupo(cont) = vRelat(cont).c3       

                grupo_a = vRelat(cont).c2  
                saldo_grupo = 0            
                n_ctrl_reg = n_ctrl_reg + 1
            end if

            x = x & "   <TR NOWRAP>" & chr(13)

         '> PRODUTO
		    x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & vRelat(cont).c5 & "</P></TD>" & chr(13)

 	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(vRelat(cont).c7) & "</P></TD>" & chr(13)

        '> SALDO
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(vRelat(cont).c8) & "</P></TD>" & chr(13)

            x = x & "   </TR>" & chr(13)

            saldo_grupo = saldo_grupo + CInt(vRelat(cont).c8)
            n_saldo_parcial = n_saldo_parcial + vRelat(cont).c8
		    n_saldo_total = n_saldo_total + vRelat(cont).c8

        end if
		
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
	next

    ' MOSTRA TOTAL DO ÚLTIMO GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        if n_reg <> 0 then
            x = x & "		<TD class='MB ME' colspan='2' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)                    
            vt_QtdeGrupo(cont) = saldo_grupo
        end if
    end if
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP>"  & chr(13) & _
				"		<TD class='MEB' COLSPAN='" & CStr(intQtdeColunasColSpanSubTotal) & "' align='right' NOWRAP><P class='Cd' style='font-weight:bold;'>" & "TOTAL:" & "</P></TD>" & chr(13) & _
				"		<TD class='MDB' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

    ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = "Grupo" then
        redim vt_TotalGrupo(0)
        set vt_TotalGrupo(0) = New cl_TRES_COLUNAS
        with vt_TotalGrupo(0)
            .c1 = ""
            .c2 = ""
            .c3 = ""
        end with

        if vt_CodigoGrupo(UBound(vt_CodigoGrupo)) <> "" then
            for cont = 0 to UBound(vt_CodigoGrupo) 
                if Trim(vt_TotalGrupo(UBound(vt_TotalGrupo)).c1) <> "" then
                    redim preserve vt_TotalGrupo(UBound(vt_TotalGrupo)+1)
                    set vt_TotalGrupo(UBound(vt_TotalGrupo)) = New cl_TRES_COLUNAS
                end if
                with vt_TotalGrupo(UBound(vt_TotalGrupo))
                    .c1 = vt_CodigoGrupo(cont)
                    .c2 = vt_DescricaoGrupo(cont)
                    .c3 = vt_QtdeGrupo(cont)
                end with
            next
        end if
        ordena_cl_tres_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        n_ctrl_reg = 0
        grupo_a = "XXXXXXX"

        if (qtde_fabricantes > 1) Or (qtde_lojas > 1) then
            x = x & _
                "	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:azure'>" & chr(13) & _
                "		<TD class='MTBD ME' valign='bottom' colspan='" & CStr(intQtdeTotalColunasColSpan) & "' NOWRAP><P class='R' style='font-weight:bold;'>TOTAL GERAL POR GRUPO DE PRODUTOS</P></TD>" & chr(13) & _
                "   </TR>" & CHR(13)

        for cont = 0 to UBound(vt_TotalGrupo)
            if vt_TotalGrupo(cont).c1 <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & _
                    "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
				    "	</TR>" & chr(13)
                    saldo_grupo = 0
                end if
                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MDB ME' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
			"	</TR>" & chr(13)

        end if
    end if




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
				"		<TD class='MTB' align='right' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;'>TOTAL GERAL:</p></TD>" & chr(13) & _
				"		<TD class='MTBD' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_total) & "</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='" & CStr(intQtdeTotalColunasColSpan) & "' align='center' class='MDBE'><P class='ALERTA'>NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS</P></TD>" & chr(13) & _
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



<html dir="ltr">


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
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque</span>
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

<!--  EMPRESA  -->
	<% if c_empresa = "" then
			s = "Todas"
		else
			s = obtem_apelido_empresa_NFe_emitente(c_empresa)
			end if
	 %>
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" nowrap colspan="2"><span class="PLTe">Empresa</span>			
		<br><span class="C">
            <%=s%>
			</span></td>
	</tr>

<!--  FABRICANTE  -->
	<% if cod_fabricante <> "" Or c_fabricante_multiplo <> "" then %>
		<tr bgcolor="#FFFFFF">
		<td class="MDBE" nowrap colspan="2"><span class="PLTe">Fabricante(s)</span>			
			<br><span class="C">
                <%if cod_fabricante <> "" then Response.Write cod_fabricante %>
                <%if cod_fabricante <> "" And c_fabricante_multiplo <> "" then Response.Write ", " %>
                <%if c_fabricante_multiplo <> "" then Response.Write c_fabricante_multiplo %>
			    </span></td>
		</tr>
	<% end if %>
	
<!--  PRODUTO  -->
	<% if cod_produto <> "" then %>
		<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP colspan="2"><span class="PLTe">Produto</span>
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
