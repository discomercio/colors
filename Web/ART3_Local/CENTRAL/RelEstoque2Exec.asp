<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelEstoque2Exec.asp
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

	const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
	const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
	const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
	const MSO_NUMBER_FORMAT_TEXTO = "\@"
    Const VENDA_SHOW_ROOM = "VENDA_SHOW_ROOM"
	Const COD_TIPO_AGRUPAMENTO__GRUPO = "Grupo"
	Const COD_TIPO_AGRUPAMENTO__PRODUTO = "Produto"

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
	if (Not operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas)) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, c_fabricante, c_produto, rb_estoque, rb_detalhe, rb_saida
	dim cod_fabricante, cod_produto
	dim s_nome_fabricante, s_nome_produto, s_nome_produto_html
    dim c_fabricante_multiplo, c_grupo, c_subgrupo, c_potencia_BTU, c_ciclo, c_posicao_mercado, v_fabricantes, v_grupos, v_subgrupos, rb_tipo_agrupamento
    dim c_empresa
	dim blnSaidaExcel
    dim Relatorio_Url, Relatorio_Id

	c_fabricante = retorna_so_digitos(Request.Form("c_fabricante"))
	if c_fabricante <> "" then c_fabricante = normaliza_codigo(c_fabricante, TAM_MIN_FABRICANTE)
	c_produto = UCase(Trim(Request.Form("c_produto")))
	rb_estoque = Trim(Request.Form("rb_estoque"))
	rb_detalhe = Trim(Request.Form("rb_detalhe"))
	rb_saida = Trim(Request.Form("rb_saida"))
    rb_tipo_agrupamento = Trim(Request.Form("rb_tipo_agrupamento"))

    c_fabricante_multiplo = Trim(Request.Form("c_fabricante_multiplo"))
	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_subgrupo = Ucase(Trim(Request.Form("c_subgrupo")))
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

        call set_default_valor_texto_bd(usuario, "RelEstoque2Filtro|c_fabricante_multiplo", c_fabricante_multiplo)
		call set_default_valor_texto_bd(usuario, "RelEstoque2Filtro|c_grupo", c_grupo)
		call set_default_valor_texto_bd(usuario, "RelEstoque2Filtro|c_subgrupo", c_subgrupo)
		call set_default_valor_texto_bd(usuario, "RelEstoque2Filtro|c_potencia_BTU", c_potencia_BTU)
		call set_default_valor_texto_bd(usuario, "RelEstoque2Filtro|c_ciclo", c_ciclo)
		call set_default_valor_texto_bd(usuario, "RelEstoque2Filtro|c_posicao_mercado", c_posicao_mercado)
		
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

'   Recupera o ID da página do relatório
    Relatorio_Id = 0
    Relatorio_Url = LCase(Trim("" & Request.ServerVariables("SCRIPT_NAME")))
    s = "SELECT * FROM t_RELATORIO_PRODUTO_PAGINA WHERE (Pagina = '" & QuotedStr(Relatorio_Url) & "')"
    rs.Open s, cn
    if Not rs.Eof then
        Relatorio_Id = CLng(rs("Id"))
    else
        s = "SET NOCOUNT ON; INSERT INTO t_RELATORIO_PRODUTO_PAGINA (Pagina) VALUES ('" & QuotedStr(Relatorio_Url) & "'); SELECT SCOPE_IDENTITY() AS Id"
        set rs = cn.Execute(s)
        if Not rs.Eof then Relatorio_Id = CLng(rs("Id"))
    end if
    
    if Relatorio_Id = 0 then 
        alerta = "Não foi possível obter o ID da página do relatório"
    end if

		if blnSaidaExcel then
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=Estoque2_" & formata_data_yyyymmdd(Now) & "_" & formata_hora_hhnnss(Now) & ".xls"
			Response.Write "<h2>Relatório de Estoque II</h2>"

			select case rb_estoque
				case ID_ESTOQUE_VENDA:			s = "VENDA"
				case ID_ESTOQUE_VENDIDO:		s = "VENDIDO"
				case ID_ESTOQUE_SHOW_ROOM:		s = "SHOW-ROOM"
				case ID_ESTOQUE_DANIFICADOS:	s = "PRODUTOS DANIFICADOS"
				case ID_ESTOQUE_DEVOLUCAO:		s = "DEVOLUÇÃO"
                case VENDA_SHOW_ROOM:           s = "VENDA + SHOW-ROOM"
				case else						s = ""
				end select
			Response.Write "Estoque de Interesse: " & s
			Response.Write "<br>"

			select case rb_detalhe
				case "SINTETICO":		s = "SINTÉTICO (SEM VALOR)"
				case "INTERMEDIARIO":	s = "INTERMEDIÁRIO (VALOR MÉDIO)"
				case "COMPLETO":		s = "COMPLETO (VALOR DIFERENCIADO)"
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
            elseif rb_estoque = VENDA_SHOW_ROOM then
                select case rb_detalhe
                    case "SINTETICO"
                        consulta_estoque_venda_show_room_detalhe_sintetico
                    case "INTERMEDIARIO"
                        consulta_estoque_venda_show_room_detalhe_intermediario
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
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_TotalGrupo(), vt_QtdeSaldoDinamicoGrupo()
dim saldo_dinamico, saldo_dinamico_parcial, saldo_dinamico_grupo, s_checked, s_bg_realce, saldo_dinamico_total
dim input_Id_Subtotal_Grupo, input_Id_Subtotal_Fabricante, input_Id_Total_Grupo

	intQtdeTotalColunasColSpan = 5
	intQtdeColunasColSpanSubTotal = 3
	intQtdeColunasColSpanTotalGeral = 2
	LargColProduto = 75
	LargColDescricao = 480
	if rb_estoque = ID_ESTOQUE_DANIFICADOS and rb_tipo_agrupamento <> COD_TIPO_AGRUPAMENTO__GRUPO then
		intQtdeTotalColunasColSpan = 6
		intQtdeColunasColSpanSubTotal = 4
		intQtdeColunasColSpanTotalGeral = 2
		LargColDescricao = 370
		end if

' 	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT" & _
            " grupo," & _
            " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
            " loja, CONVERT(smallint,loja) AS numero_loja" & _
			", t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto" & _
			", t_PRODUTO.descricao, descricao_html, Sum(qtde) AS saldo," & _
            " (SELECT TOP 1 Flag FROM t_RELATORIO_PRODUTO_FLAG" & _
	        "   WHERE (Usuario = '" & usuario & "' AND Fabricante = t_ESTOQUE_MOVIMENTO.fabricante AND Produto = t_ESTOQUE_MOVIMENTO.produto AND IdPagina = " & Relatorio_Id & ")" & _
            " ) AS saldo_flag" & _
			" FROM t_ESTOQUE_MOVIMENTO LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
            " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE.id_estoque = t_ESTOQUE_MOVIMENTO.id_estoque)" & _
			" WHERE (anulado_status=0)" & _
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

    s_where_temp = ""
	if c_subgrupo <> "" then
	    v_subgrupos = split(c_subgrupo, ", ")
	    for cont = Lbound(v_subgrupos) to Ubound(v_subgrupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.subgrupo = '" & v_subgrupos(cont) & "')"
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
			" AND (t_ESTOQUE.id_nfe_emitente = '" & c_empresa & "')"
		end if
	
	s_sql = s_sql & " GROUP BY loja, t_ESTOQUE_MOVIMENTO.fabricante, grupo, t_PRODUTO_GRUPO.descricao, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, descricao_html"

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then
		s_sql = s_sql & " ORDER BY numero_loja, t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, descricao_html"
    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        s_sql = s_sql & " ORDER BY numero_loja, t_ESTOQUE_MOVIMENTO.fabricante, grupo, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, descricao_html"
    end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
          "		<TD width='30' valign='bottom' NOWRAP class='MD ME MB'><P class='R' style='font-weight:bold;'></P></TD>" & chr(13) & _
		  "		<TD width='" & Cstr(LargColProduto) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='" & CStr(LargColDescricao) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13)

	if rb_estoque = ID_ESTOQUE_DANIFICADOS then
		cab = cab & _
			  "		<TD width='" & CStr(LargColOrdemServico) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>ORDEM SERVIÇO</P></TD>" & chr(13)
		end if		  
		  
	cab = cab & _
          "     <TD width='60' align='right' valign=bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'></P></TD>" & chr(13) & _  
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
    saldo_dinamico = 0
    saldo_dinamico_parcial = 0
    saldo_dinamico_grupo = 0
    saldo_dinamico_total = 0
	
	if blnSaidaExcel then s_nbsp = "" else s_nbsp = "&nbsp;"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	'	MUDOU LOJA?
		if Trim("" & r("loja")) <> loja_a then
			if n_reg_total > 0 then 

              ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
		            x = x & "		<TD class='MB ME' colspan='3' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

			  ' FECHA TABELA DA LOJA ANTERIOR
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD COLSPAN='" & Cstr(intQtdeColunasColSpanSubTotal) & "' align='right' valign='bottom' class='MEB' NOWRAP><P class='Cd' style='font-weight:bold;'>" & _
						"TOTAL:" & "</P></TD>" & chr(13) & _
                        "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
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
			'   QUEBRA POR LOJA APENAS SE HOUVER LOJA
				if blnSaidaExcel then s_bkg_color = "tomato" else s_bkg_color = "azure"
				x = x & "	<TR NOWRAP style='background:" & s_bkg_color & "'>" & chr(13) & _
						"		<TD class='MDBE' align='center' valign='bottom' colspan='" & CStr(intQtdeTotalColunasColSpan) & "'><P class='F' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & r("loja") & " - " & ucase(x_loja(r("loja"))) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				end if

			x = x & cab

			n_reg = 0
			n_saldo_parcial = 0
            saldo_dinamico_parcial = 0
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
                if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
		            x = x & "		<TD class='MB ME' colspan='3' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' align='right' valign='bottom' colspan='" & CStr(intQtdeColunasColSpanSubTotal) & "'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
                        "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
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
            saldo_dinamico_parcial = 0
            grupo_a = "XXXXXXX"
			end if
			
	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then

		s_checked = ""
        s_bg_realce = ""
        saldo_dinamico = 0
        if Trim("" & r("saldo_flag")) = "1" then 
            s_checked = " checked"
            s_bg_realce = "style='background-color: lightgreen'"
            saldo_dinamico = formata_inteiro(r("saldo"))
        end if

        input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))
        input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("grupo")) & "_" & Trim("" & r("fabricante"))
        input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante")) & "_" & Trim("" & r("loja"))

		x = x & "	<TR NOWRAP>" & chr(13)
    
    '> FLAG
        if Not blnSaidaExcel then
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">" & chr(13) & _
                    "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """, null, null, ""col_saldo_" & n_reg_total & """, """ & r("saldo") & """)' />" & chr(13) & _
                    "       </TD>" & chr(13)
        else
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
        end if

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

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
		
     '> SALDO (DINÂMICO)
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg_total & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

	 '> SALDO
		x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

		n_saldo_parcial = n_saldo_parcial + r("saldo")
		n_saldo_total = n_saldo_total + r("saldo")

        saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
        saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico
		
		x = x & "	</TR>" & chr(13)

    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
            if Trim("" & r("grupo")) <> grupo_a then
                if n_ctrl_reg > 0 then
		            x = x & "		<TD class='MB ME' colspan='3' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
                    cont = cont + 1
                end if
                x = x & "	<TR NOWRAP>" & chr(13)
                x = x & "       <TD class='MDBE' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C'>&nbsp;</P></TD>" & chr(13)
		        x = x & "		<TD class='MDB' colspan='5' align='left' style='background-color: #EEE' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)

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
                saldo_dinamico_grupo = 0
                n_ctrl_reg = n_ctrl_reg + 1
            end if

            s_checked = ""
            s_bg_realce = ""
            saldo_dinamico = 0
            if Trim("" & r("saldo_flag")) = "1" then 
                s_checked = " checked"
                s_bg_realce = "style='background-color: lightgreen'"
                saldo_dinamico = formata_inteiro(r("saldo"))
            end if

            input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))        
            input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("fabricante")) & "_" & Trim("" & r("grupo"))
            input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante"))

            x = x & "   <TR NOWRAP>" & chr(13)

            '> FLAG
            if Not blnSaidaExcel then
                x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">"  & chr(13) & _
                        "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """,""" & input_Id_Subtotal_Grupo & """, """ & input_Id_Total_Grupo & """, ""col_saldo_" & n_reg_total & """, """ & r("saldo") & """)' />" & chr(13) & _
                        "       </TD>" & chr(13)
            else
                x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
            end if

         '> PRODUTO
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

        '> SALDO (DINÂMICO)
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg_total & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

        '> SALDO
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

            x = x & "   </TR>" & chr(13)

            saldo_grupo = saldo_grupo + CInt(r("saldo"))
            n_saldo_parcial = n_saldo_parcial + r("saldo")
		    n_saldo_total = n_saldo_total + r("saldo")

            saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
            saldo_dinamico_grupo = saldo_dinamico_grupo + saldo_dinamico
            saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico

        end if
		
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop

  ' MOSTRA TOTAL DO ÚLTIMO GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        if n_reg <> 0 then
            x = x & "		<TD class='MB ME' colspan='3' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)                    
            vt_QtdeGrupo(cont) = saldo_grupo
            redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
            vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
        end if
    end if
	
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP>"  & chr(13) & _
				"		<TD class='MEB' COLSPAN='" & CStr(intQtdeColunasColSpanSubTotal) & "' align='right' NOWRAP><P class='Cd' style='font-weight:bold;'>" & "TOTAL:" & "</P></TD>" & chr(13) & _
                "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

    ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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
                    .c4 = vt_QtdeSaldoDinamicoGrupo(cont)
                end with
            next
        end if
        ordena_cl_quatro_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        saldo_dinamico_grupo = 0
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
                    "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
                    "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
				    "	</TR>" & chr(13)
                    saldo_grupo = 0
                    saldo_dinamico_grupo = 0
                end if

                if vt_TotalGrupo(cont).c1 <> "-" then 
                    input_Id_Total_Grupo = "TotalGeral_Grupo__" & vt_TotalGrupo(cont).c1
                else
                    input_Id_Total_Grupo = "TotalGeral_Grupo__"
                end if

                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MDB ME' colspan='2' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            saldo_dinamico_grupo = saldo_dinamico_grupo + vt_TotalGrupo(cont).c4
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
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
				"		<TD class='MTB' align='right' colspan='" & CStr(intQtdeColunasColSpanTotalGeral) & "' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;'>TOTAL GERAL:</p></TD>" & chr(13) & _
                "		<TD NOWRAP class='MC MB' valign='bottom'><P class='Cd' id='input_Id_Total_Saldo' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_total)=0,"",formata_inteiro(saldo_dinamico_total)) & "</P></TD>" & chr(13) & _
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

    x = x & "<input type='hidden' id='qtde_fabricantes' value='" & qtde_fabricantes & "' />" & chr(13)
	
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
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_CustoEntrada(), vt_TotalGrupo(), vt_QtdeSaldoDinamicoGrupo()
dim saldo_dinamico, saldo_dinamico_parcial, saldo_dinamico_grupo, s_checked, s_bg_realce, saldo_dinamico_total
dim input_Id_Subtotal_Grupo, input_Id_Subtotal_Fabricante, input_Id_Total_Grupo

	intQtdeTotalColunasColSpan = 7
	intQtdeColunasColSpanSubTotal = 3
	intQtdeColunasColSpanTotalGeral = 2
	LargColDescricao = 270
	if rb_estoque = ID_ESTOQUE_DANIFICADOS and rb_tipo_agrupamento <> COD_TIPO_AGRUPAMENTO__GRUPO then
		intQtdeTotalColunasColSpan = 8
		intQtdeColunasColSpanSubTotal = 4
		intQtdeColunasColSpanTotalGeral = 2
		LargColDescricao = 200
		end if

	LargColProduto = 60
	if blnSaidaExcel then
		LargColProduto = 75
		LargColDescricao = 350
		end if

' 	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT" & _
            " t_PRODUTO.grupo AS grupo," & _
            " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
            " loja, CONVERT(smallint,loja) AS numero_loja" & _
			", t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto" & _
			", t_PRODUTO.descricao, descricao_html, Sum(t_ESTOQUE_MOVIMENTO.qtde) AS saldo" & _
			", Sum(t_ESTOQUE_MOVIMENTO.qtde*t_ESTOQUE_ITEM.vl_custo2) AS preco_total," & _
            " (SELECT TOP 1 Flag FROM t_RELATORIO_PRODUTO_FLAG" & _
	        "   WHERE (Usuario = '" & usuario & "' AND Fabricante = t_ESTOQUE_MOVIMENTO.fabricante AND Produto = t_ESTOQUE_MOVIMENTO.produto AND IdPagina = " & Relatorio_Id & ")" & _
            " ) AS saldo_flag" & _
			" FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
            " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE.id_estoque = t_ESTOQUE_MOVIMENTO.id_estoque)" & _
			" WHERE (anulado_status=0)" & _
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

    s_where_temp = ""
	if c_subgrupo <> "" then
	    v_subgrupos = split(c_subgrupo, ", ")
	    for cont = Lbound(v_subgrupos) to Ubound(v_subgrupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.subgrupo = '" & v_subgrupos(cont) & "')"
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
			" AND (t_ESTOQUE.id_nfe_emitente = '" & c_empresa & "')"
		end if

	s_sql = s_sql & " GROUP BY loja, t_ESTOQUE_MOVIMENTO.fabricante, t_PRODUTO.grupo, t_PRODUTO_GRUPO.descricao, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, descricao_html"
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then
		        s_sql = s_sql &	" ORDER BY numero_loja, t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, descricao_html"
    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
                s_sql = s_sql & " ORDER BY numero_loja, t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, descricao_html"
    end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure'>" & chr(13) & _
          "		<TD width='30' valign='bottom' NOWRAP class='MD ME MB'><P class='R' style='font-weight:bold;'></P></TD>" & chr(13) & _
		  "		<TD width='" & Cstr(LargColProduto) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='" & CStr(LargColDescricao) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13)

	if rb_estoque = ID_ESTOQUE_DANIFICADOS then
		cab = cab & _
			  "		<TD width='" & CStr(LargColOrdemServico) & "' valign='bottom' class='MDB'><P class='R' style='font-weight:bold;'>ORDEM SERVIÇO</P></TD>" & chr(13)
		end if
		
	cab = cab & _
          "     <TD width='60' align='right' valign=bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'></P></TD>" & chr(13) & _  
		  "		<TD width='60' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13)
	
	if blnSaidaExcel then
		cab = cab & _
		  "		<TD style='width:100px' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;width:100px;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
		  "		<TD style='width:100px' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;width:100px;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	else
		cab = cab & _
		  "		<TD width='100' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
		  "		<TD width='100' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA TOTAL</P></TD>" & chr(13) & _
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
    saldo_dinamico = 0
    saldo_dinamico_parcial = 0
    saldo_dinamico_grupo = 0
    saldo_dinamico_total = 0
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
                if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
		            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='3' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont) 
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo          
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

			  ' FECHA TABELA DA LOJA ANTERIOR
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD COLSPAN='" & Cstr(intQtdeColunasColSpanSubTotal) & "' align='right' valign='bottom' class='MEB' NOWRAP><P class='Cd' style='font-weight:bold;'>" & _
						"TOTAL:</P></TD>" & chr(13) & _
                        "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
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
            saldo_dinamico_parcial = 0
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
                if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
		            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='3' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)      
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo     
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='" & CStr(intQtdeColunasColSpanSubTotal) & "' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
                        "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
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
            saldo_dinamico_parcial = 0
			vl_sub_total = 0
            grupo_a = "XXXXXXX"
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then

		s_checked = ""
        s_bg_realce = ""
        saldo_dinamico = 0
        if Trim("" & r("saldo_flag")) = "1" then 
            s_checked = " checked"
            s_bg_realce = "style='background-color: lightgreen'"
            saldo_dinamico = formata_inteiro(r("saldo"))
        end if

        input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))
        input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("grupo")) & "_" & Trim("" & r("fabricante"))
        input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante")) & "_" & Trim("" & r("loja"))

		x = x & "	<TR NOWRAP>" & chr(13)
    
    '> FLAG
        if Not blnSaidaExcel then
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">" & chr(13) & _
                    "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """, null, null, ""col_saldo_" & n_reg_total & """, """ & r("saldo") & """)' />" & chr(13) & _
                    "       </TD>" & chr(13)
        else
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
        end if

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

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

     '> SALDO (DINÂMICO)
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg_total & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

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
        saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
        saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico

		x = x & "	</TR>" & chr(13)

    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
            if Trim("" & r("grupo")) <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='3' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)           
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                end if
                x = x & "	<TR NOWRAP>" & chr(13)
                x = x & "       <TD class='MDBE' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C'>&nbsp;</P></TD>" & chr(13)
		        x = x & "		<TD class='MDB' colspan='6' align='left' style='background-color: #EEE' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)


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
                saldo_dinamico_grupo = 0
                vl_custo_entrada_grupo = 0  
                n_ctrl_reg = n_ctrl_reg + 1
            end if

        s_checked = ""
        s_bg_realce = ""
        saldo_dinamico = 0
        if Trim("" & r("saldo_flag")) = "1" then 
            s_checked = " checked"
            s_bg_realce = "style='background-color: lightgreen'"
            saldo_dinamico = formata_inteiro(r("saldo"))
        end if

        input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))        
        input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("fabricante")) & "_" & Trim("" & r("grupo"))
        input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante"))

        x = x & "   <TR NOWRAP>" & chr(13)

        '> FLAG
        if Not blnSaidaExcel then
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">"  & chr(13) & _
                    "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """,""" & input_Id_Subtotal_Grupo & """, """ & input_Id_Total_Grupo & """, ""col_saldo_" & n_reg_total & """, """ & r("saldo") & """)' />" & chr(13) & _
                    "       </TD>" & chr(13)
        else
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
        end if

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' width='" & CStr(LargColDescricao) & "' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

     '> SALDO (DINÂMICO)
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg_total & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

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
            saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
            saldo_dinamico_grupo = saldo_dinamico_grupo + saldo_dinamico
            saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico            
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
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        if n_reg <> 0 then
            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='3' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)         
            redim preserve vt_CustoEntrada(cont)           
            redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
            vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
            vt_QtdeGrupo(cont) = saldo_grupo
            vt_CustoEntrada(cont) = vl_custo_entrada_grupo
        end if
    end if
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MEB' COLSPAN='" & CStr(intQtdeColunasColSpanSubTotal) & "' align='right' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;'>" & "TOTAL:</P></TD>" & chr(13) & _
                "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
				"		<TD class='MB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"		<TD class='MB' valign='bottom'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

    ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        redim vt_TotalGrupo(0)
        set vt_TotalGrupo(0) = New cl_CINCO_COLUNAS
        with vt_TotalGrupo(0)
            .c1 = ""
            .c2 = ""
            .c3 = ""
            .c4 = ""
            .c5 = ""
        end with

        if vt_CodigoGrupo(UBound(vt_CodigoGrupo)) <> "" then
            for cont = 0 to UBound(vt_CodigoGrupo) 
                if Trim(vt_TotalGrupo(UBound(vt_TotalGrupo)).c1) <> "" then
                    redim preserve vt_TotalGrupo(UBound(vt_TotalGrupo)+1)
                    set vt_TotalGrupo(UBound(vt_TotalGrupo)) = New cl_CINCO_COLUNAS
                end if
                with vt_TotalGrupo(UBound(vt_TotalGrupo))
                    .c1 = vt_CodigoGrupo(cont)
                    .c2 = vt_DescricaoGrupo(cont)
                    .c3 = vt_QtdeGrupo(cont)
                    .c4 = vt_CustoEntrada(cont)
                    .c5 = vt_QtdeSaldoDinamicoGrupo(cont)
                end with
            next
        end if
        ordena_cl_cinco_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        saldo_dinamico_grupo = 0
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
                    "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
                    "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
                    "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    				"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
                    "	</TR>" & chr(13)

                    saldo_grupo = 0
                    saldo_dinamico_grupo = 0
                    vl_custo_entrada_grupo = 0
                end if

                if vt_TotalGrupo(cont).c1 <> "-" then 
                    input_Id_Total_Grupo = "TotalGeral_Grupo__" & vt_TotalGrupo(cont).c1
                else
                    input_Id_Total_Grupo = "TotalGeral_Grupo__"
                end if

                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MB ME' colspan='2' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            saldo_dinamico_grupo = saldo_dinamico_grupo + vt_TotalGrupo(cont).c5
            vl_custo_entrada_grupo = vl_custo_entrada_grupo + vt_TotalGrupo(cont).c4
            
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
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
				"		<TD class='MTB' align='right' valign='bottom' colspan='" & CStr(intQtdeColunasColSpanTotalGeral) & "' NOWRAP><p class='Cd' style='font-weight:bold;'>TOTAL GERAL:</p></TD>" & chr(13) & _
                "		<TD NOWRAP class='MC MB' valign='bottom'><P class='Cd' id='input_Id_Total_Saldo' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_total)=0,"",formata_inteiro(saldo_dinamico_total)) & "</P></TD>" & chr(13) & _
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

    x = x & "<input type='hidden' id='qtde_fabricantes' value='" & qtde_fabricantes & "' />" & chr(13)
	
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

'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT" & _
            " t_PRODUTO.grupo AS grupo," & _
            " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
            " loja, CONVERT(smallint,loja) AS numero_loja" & _
			", t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto" & _
			", t_PRODUTO.descricao, descricao_html, Sum(t_ESTOQUE_MOVIMENTO.qtde) AS saldo" & _
			", t_ESTOQUE_ITEM.vl_custo2" & _
			" FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
            " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE.id_estoque = t_ESTOQUE_MOVIMENTO.id_estoque)" & _
			" WHERE (anulado_status=0)" & _
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

    s_where_temp = ""
	if c_subgrupo <> "" then
	    v_subgrupos = split(c_subgrupo, ", ")
	    for cont = Lbound(v_subgrupos) to Ubound(v_subgrupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.subgrupo = '" & v_subgrupos(cont) & "')"
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
			" AND (t_ESTOQUE.id_nfe_emitente = '" & c_empresa & "')"
		end if

	s_sql = s_sql & " GROUP BY loja, t_ESTOQUE_MOVIMENTO.fabricante, t_PRODUTO.grupo, t_PRODUTO_GRUPO.descricao, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, descricao_html, t_ESTOQUE_ITEM.vl_custo2"

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then
		        s_sql = s_sql &	" ORDER BY numero_loja, t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, descricao_html, t_ESTOQUE_ITEM.vl_custo2"
    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
                s_sql = s_sql & " ORDER BY numero_loja, t_ESTOQUE_MOVIMENTO.fabricante, t_PRODUTO.grupo, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, descricao_html, t_ESTOQUE_ITEM.vl_custo2"
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
		  "		<TD width='" & Cstr(largColProduto) & "' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='" & Cstr(largColDescricao) & "' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
		  "		<TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />UNITÁRIO</P></TD>" & chr(13) & _
		  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />TOTAL</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	else
		cab = cab & _
		  "		<TD width='" & Cstr(largColProduto) & "' valign='bottom' class='MD ME MB'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='" & Cstr(largColDescricao) & "' valign='bottom' class='MD MB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
		  "		<TD width='60' align='right' valign='bottom' class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD width='100' align='right' valign='bottom' class='MD MB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA UNITÁRIO</P></TD>" & chr(13) & _
		  "		<TD width='100' align='right' valign='bottom' class='MDB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA TOTAL</P></TD>" & chr(13) & _
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
	qtde_fabricantes = 0
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
                if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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
                if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDB ME' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

	 '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA UNITÁRIO
		vl = r("vl_custo2")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA TOTAL
		vl = r("vl_custo2")*r("saldo")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

		vl_sub_total = vl_sub_total + vl
		vl_total_geral = vl_total_geral + vl
		
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		n_saldo_total = n_saldo_total + r("saldo")
		
		x = x & "	</TR>" & chr(13)

    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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

	 '> CUSTO DE ENTRADA UNITÁRIO
		vl = r("vl_custo2")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA TOTAL
		vl = r("vl_custo2")*r("saldo")
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
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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
				"		<TD class='MB ME' COLSPAN='2' align='right' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;'>" & "TOTAL:</P></TD>" & chr(13) & _
				"		<TD class='MB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

       ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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
			x = x & _
                "	<TR NOWRAP><TD COLSPAN='5' class='MB'>&nbsp;</TD></TR>" & chr(13) & _
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
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_TotalGrupo(), vt_QtdeSaldoDinamicoGrupo()
dim saldo_dinamico, saldo_dinamico_parcial, saldo_dinamico_grupo, s_checked, s_bg_realce, saldo_dinamico_total
dim input_Id_Subtotal_Grupo, input_Id_Subtotal_Fabricante, input_Id_Total_Grupo

	s_sql = "SELECT" & _
            " t_PRODUTO.grupo," & _
            " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
            " t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html, Sum(qtde-qtde_utilizada) AS saldo," & _
            " (SELECT TOP 1 Flag FROM t_RELATORIO_PRODUTO_FLAG" & _
	        "   WHERE (Usuario = '" & usuario & "' AND Fabricante = t_ESTOQUE_ITEM.fabricante AND Produto = t_ESTOQUE_ITEM.produto AND IdPagina = " & Relatorio_Id & ")" & _
            " ) AS saldo_flag" & _
			" FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON" & _
			" ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
            " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE.id_estoque = t_ESTOQUE_ITEM.id_estoque)" & _
			" WHERE ((qtde-qtde_utilizada) > 0)"

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

    s_where_temp = ""
	if c_subgrupo <> "" then
	    v_subgrupos = split(c_subgrupo, ", ")
	    for cont = Lbound(v_subgrupos) to Ubound(v_subgrupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.subgrupo = '" & v_subgrupos(cont) & "')"
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
			" AND (t_ESTOQUE.id_nfe_emitente = '" & c_empresa & "')"
		end if
	
	s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_PRODUTO.grupo, t_PRODUTO_GRUPO.descricao, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html"

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then
		s_sql = s_sql &	" ORDER BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html"
    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        s_sql = s_sql & " ORDER BY t_ESTOQUE_ITEM.fabricante, t_PRODUTO.grupo, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html"
        
    end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
	
'	Mantendo a padronização com as demais opções de emissão do relatório em Excel (sem bordas no cabeçalho)
	if blnSaidaExcel then
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD width='30' valign='bottom' NOWRAP><P class='R' style='font-weight:bold;'></P></TD>" & chr(13) & _
			  "		<TD width='75' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='480' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
              "     <TD width='60' align='right' valign=bottom' NOWRAP><P class='Rd' style='font-weight:bold;'></P></TD>" & chr(13) & _  
			  "		<TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
	else
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD width='30' valign='bottom' NOWRAP class='MD ME MB'><P class='R' style='font-weight:bold;'></P></TD>" & chr(13) & _
			  "		<TD width='75' valign='bottom' NOWRAP class='MD MB'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='480' valign='bottom' NOWRAP class='MB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
              "     <TD width='60' align='right' valign=bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'></P></TD>" & chr(13) & _  
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
    saldo_dinamico = 0
    saldo_dinamico_parcial = 0
    saldo_dinamico_grupo = 0
    saldo_dinamico_total = 0
	
	if blnSaidaExcel then s_nbsp = "" else s_nbsp = "&nbsp;"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then

            ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
		            x = x & "		<TD class='MB ME' colspan='3' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if
    
				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='3' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
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
            saldo_dinamico_parcial = 0
            grupo_a = "XXXXXXX"
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then

        s_checked = ""
        s_bg_realce = ""
        saldo_dinamico = 0
        if Trim("" & r("saldo_flag")) = "1" then 
            s_checked = " checked"
            s_bg_realce = "style='background-color: lightgreen'"
            saldo_dinamico = formata_inteiro(r("saldo"))
        end if

        input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))
        input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("grupo")) & "_" & Trim("" & r("fabricante"))
        input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante"))

		x = x & "	<TR NOWRAP>" & chr(13)
    
    '> FLAG
        if Not blnSaidaExcel then
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">" & chr(13) & _
                    "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """, null, null, ""col_saldo_" & n_reg & """, """ & r("saldo") & """)' />" & chr(13) & _
                    "       </TD>" & chr(13)
        else
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
        end if

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

    '> SALDO (DINÂMICO)
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

	 '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

		n_saldo_total = n_saldo_total + r("saldo")
		n_saldo_parcial = n_saldo_parcial + r("saldo")
        saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
        saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico
		
		x = x & "	</TR>" & chr(13)

    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        if Trim("" & r("grupo")) <> grupo_a then
            if n_ctrl_reg > 0 then
		        x = x & "		<TD class='MB ME' colspan='3' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		        x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		        x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                x = x & "   </TR>" & chr(13)

                redim preserve vt_QtdeGrupo(cont)                    
                vt_QtdeGrupo(cont) = saldo_grupo
                redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
                cont = cont + 1
            end if
            x = x & "	<TR NOWRAP>" & chr(13)
            x = x & "       <TD class='MDBE' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C'>&nbsp;</P></TD>" & chr(13)
		    x = x & "		<TD class='MDB' colspan='4' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)

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
            saldo_dinamico_grupo = 0
            n_ctrl_reg = n_ctrl_reg + 1
        end if

        s_checked = ""
        s_bg_realce = ""
        saldo_dinamico = 0
        if Trim("" & r("saldo_flag")) = "1" then 
            s_checked = " checked"
            s_bg_realce = "style='background-color: lightgreen'"
            saldo_dinamico = formata_inteiro(r("saldo"))
        end if

        input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))        
        input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("fabricante")) & "_" & Trim("" & r("grupo"))
        input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante"))

        x = x & "   <TR NOWRAP>" & chr(13)

        '> FLAG
        if Not blnSaidaExcel then
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">"  & chr(13) & _
                    "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """,""" & input_Id_Subtotal_Grupo & """, """ & input_Id_Total_Grupo & """, ""col_saldo_" & n_reg & """, """ & r("saldo") & """)' />" & chr(13) & _
                    "       </TD>" & chr(13)
        else
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
        end if

    	'> PRODUTO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	    '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

        '> SALDO (DINÂMICO)
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

 	    '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

        x = x & "   </TR>" & chr(13)

        saldo_grupo = saldo_grupo + CInt(r("saldo"))
        n_saldo_parcial = n_saldo_parcial + r("saldo")
		n_saldo_total = n_saldo_total + r("saldo")
        saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
        saldo_dinamico_grupo = saldo_dinamico_grupo + saldo_dinamico
        saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico

        end if

		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop

  ' MOSTRA TOTAL DO ÚLTIMO GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        if n_reg <> 0 then
            x = x & "		<TD class='MB ME' colspan='3' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)                    
            vt_QtdeGrupo(cont) = saldo_grupo
            redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
            vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
        end if
    end if
		
  ' MOSTRA TOTAL
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO FORNECEDOR
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MEB' colspan='3' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
				"		<TD class='MDB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

  ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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
                    .c4 = vt_QtdeSaldoDinamicoGrupo(cont)
                end with
            next
        end if
        ordena_cl_quatro_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        saldo_dinamico_grupo = 0
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
                    "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
                    "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
				    "	</TR>" & chr(13)
                    saldo_grupo = 0
                    saldo_dinamico_grupo = 0
                end if
                if vt_TotalGrupo(cont).c1 <> "-" then 
                    input_Id_Total_Grupo = "TotalGeral_Grupo__" & vt_TotalGrupo(cont).c1
                else
                    input_Id_Total_Grupo = "TotalGeral_Grupo__"
                end if
                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MDB ME' colspan='2' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            saldo_dinamico_grupo = saldo_dinamico_grupo + vt_TotalGrupo(cont).c4
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
            "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
			"	</TR>" & chr(13)

        end if
    end if
		
		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='5'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='3' NOWRAP class='MC MEB' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>" & "TOTAL GERAL:" & "</P></TD>" & chr(13) & _
					"		<TD NOWRAP class='MC MDB' valign='bottom'><P class='Cd' id='input_Id_Total_Saldo' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_total)=0,"",formata_inteiro(saldo_dinamico_total)) & "</P></TD>" & chr(13) & _
					"		<TD NOWRAP class='MC MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
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
	
    x = x & "<input type='hidden' id='qtde_fabricantes' value='" & qtde_fabricantes & "' />" & chr(13)

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
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_CustoEntrada(), vt_TotalGrupo(), vt_QtdeSaldoDinamicoGrupo()
dim saldo_dinamico, saldo_dinamico_parcial, saldo_dinamico_grupo, s_checked, s_bg_realce, saldo_dinamico_total
dim input_Id_Subtotal_Grupo, input_Id_Subtotal_Fabricante, input_Id_Total_Grupo

	s_sql = "SELECT" & _
            " t_PRODUTO.grupo AS grupo," & _
            " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
            " t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html" & _
			", Sum(qtde-qtde_utilizada) AS saldo, Sum((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total," & _
            " (SELECT TOP 1 Flag FROM t_RELATORIO_PRODUTO_FLAG" & _
	        "   WHERE (Usuario = '" & usuario & "' AND Fabricante = t_ESTOQUE_ITEM.fabricante AND Produto = t_ESTOQUE_ITEM.produto AND IdPagina = " & Relatorio_Id & ")" & _
            " ) AS saldo_flag" & _
			" FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON" & _
			" ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
            " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE.id_estoque = t_ESTOQUE_ITEM.id_estoque)" & _
			" WHERE ((qtde-qtde_utilizada) > 0)"
	
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

    s_where_temp = ""
	if c_subgrupo <> "" then
	    v_subgrupos = split(c_subgrupo, ", ")
	    for cont = Lbound(v_subgrupos) to Ubound(v_subgrupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.subgrupo = '" & v_subgrupos(cont) & "')"
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
			" AND (t_ESTOQUE.id_nfe_emitente = '" & c_empresa & "')"
		end if

	s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_PRODUTO.grupo, t_PRODUTO_GRUPO.descricao, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html"

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then
		        s_sql = s_sql &	" ORDER BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html"
    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
                s_sql = s_sql & " ORDER BY t_ESTOQUE_ITEM.fabricante, t_PRODUTO.grupo, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html"
    end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
'	Mantendo a padronização com as demais opções de emissão do relatório em Excel (sem bordas no cabeçalho)
	if blnSaidaExcel then
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
              "		<TD width='30' valign='bottom' NOWRAP><P class='R' style='font-weight:bold;'></P></TD>" & chr(13) & _
			  "		<TD width='75' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='350' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
              "     <TD width='60' align='right' valign=bottom' NOWRAP><P class='Rd' style='font-weight:bold;'></P></TD>" & chr(13) & _  
			  "		<TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />TOTAL</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
	else
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
              "		<TD width='30' valign='bottom' NOWRAP class='MD ME MB'><P class='R' style='font-weight:bold;'></P></TD>" & chr(13) & _
			  "		<TD width='60' valign='bottom' NOWRAP class='MDB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='274' valign='bottom' class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
              "     <TD width='60' align='right' valign=bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'></P></TD>" & chr(13) & _  
			  "		<TD width='60' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "		<TD width='100' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
			  "		<TD width='100' valign='bottom' NOWRAP class='MDB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA TOTAL</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
		end if
	
	x = cab_table & cab
	n_reg = 0
	n_saldo_total = 0
	n_saldo_parcial = 0
	vl_total_geral = 0
	vl_sub_total = 0
	qtde_fabricantes = 0
    saldo_grupo = 0
    n_ctrl_reg = 0
	fabricante_a = "XXXXX"
    grupo_a = "XXXXX"
    saldo_dinamico = 0
    saldo_dinamico_parcial = 0
    saldo_dinamico_grupo = 0
    saldo_dinamico_total = 0
		
	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then

            ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
		            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='3' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
    		        x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if


				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='3' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
                        "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
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
						"		<TD colspan='7' class='MB'>&nbsp;</TD>" & _
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
					"		<TD class='MDBE' align='center' valign='bottom' colspan='7' style='background: honeydew'><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
            saldo_dinamico_parcial = 0
			vl_sub_total = 0
            grupo_a = "XXXXXXX"
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then

		s_checked = ""
        s_bg_realce = ""
        saldo_dinamico = 0
        if Trim("" & r("saldo_flag")) = "1" then 
            s_checked = " checked"
            s_bg_realce = "style='background-color: lightgreen'"
            saldo_dinamico = formata_inteiro(r("saldo"))
        end if

        input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))
        input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("grupo")) & "_" & Trim("" & r("fabricante"))
        input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante"))

		x = x & "	<TR NOWRAP>" & chr(13)
    
    '> FLAG
        if Not blnSaidaExcel then
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">" & chr(13) & _
                    "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """, null, null, ""col_saldo_" & n_reg & """, """ & r("saldo") & """)' />" & chr(13) & _
                    "       </TD>" & chr(13)
        else
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
        end if

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

    '> SALDO (DINÂMICO)
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

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
        saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
        saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico
		
		x = x & "	</TR>" & chr(13)

    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
            if Trim("" & r("grupo")) <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='3' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont) 
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo          
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                end if
                x = x & "	<TR NOWRAP>" & chr(13)
                x = x & "       <TD class='MDBE' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C'>&nbsp;</P></TD>" & chr(13)
		        x = x & "		<TD class='MDB' colspan='6' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)


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
                saldo_dinamico_grupo = 0
                vl_custo_entrada_grupo = 0  
                n_ctrl_reg = n_ctrl_reg + 1
            end if

        s_checked = ""
        s_bg_realce = ""
        saldo_dinamico = 0
        if Trim("" & r("saldo_flag")) = "1" then 
            s_checked = " checked"
            s_bg_realce = "style='background-color: lightgreen'"
            saldo_dinamico = formata_inteiro(r("saldo"))
        end if

        input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))        
        input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("fabricante")) & "_" & Trim("" & r("grupo"))
        input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante"))

        x = x & "   <TR NOWRAP>" & chr(13)

        '> FLAG
        if Not blnSaidaExcel then
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">"  & chr(13) & _
                    "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """,""" & input_Id_Subtotal_Grupo & """, """ & input_Id_Total_Grupo & """, ""col_saldo_" & n_reg & """, """ & r("saldo") & """)' />" & chr(13) & _
                    "       </TD>" & chr(13)
        else
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
        end if


	     '> PRODUTO
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

        '> SALDO (DINÂMICO)
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

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
            saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
            saldo_dinamico_grupo = saldo_dinamico_grupo + saldo_dinamico
            saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico
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
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        if n_reg <> 0 then
            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='3' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)         
            redim preserve vt_CustoEntrada(cont)       
            redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
            vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo    
            vt_QtdeGrupo(cont) = saldo_grupo
            vt_CustoEntrada(cont) = vl_custo_entrada_grupo
        end if
    end if
		
  ' MOSTRA TOTAL
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO FORNECEDOR
		x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='3' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
                        "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

  ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        redim vt_TotalGrupo(0)
        set vt_TotalGrupo(0) = New cl_CINCO_COLUNAS
        with vt_TotalGrupo(0)
            .c1 = ""
            .c2 = ""
            .c3 = ""
            .c4 = ""
            .c5 = ""
        end with

        if vt_CodigoGrupo(UBound(vt_CodigoGrupo)) <> "" then
            for cont = 0 to UBound(vt_CodigoGrupo) 
                if Trim(vt_TotalGrupo(UBound(vt_TotalGrupo)).c1) <> "" then
                    redim preserve vt_TotalGrupo(UBound(vt_TotalGrupo)+1)
                    set vt_TotalGrupo(UBound(vt_TotalGrupo)) = New cl_CINCO_COLUNAS
                end if
                with vt_TotalGrupo(UBound(vt_TotalGrupo))
                    .c1 = vt_CodigoGrupo(cont)
                    .c2 = vt_DescricaoGrupo(cont)
                    .c3 = vt_QtdeGrupo(cont)
                    .c4 = vt_CustoEntrada(cont)
                    .c5 = vt_QtdeSaldoDinamicoGrupo(cont)
                end with
            next
        end if
        ordena_cl_cinco_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        saldo_dinamico_grupo = 0
        vl_custo_entrada_grupo = 0
        n_ctrl_reg = 0
        grupo_a = "XXXXXXX"

        if (qtde_fabricantes > 1) then
            x = x & _
                "	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='6'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='6'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:azure'>" & chr(13) & _
                "		<TD class='MTBD ME' valign='bottom' colspan='7' NOWRAP><P class='R' style='font-weight:bold;'>TOTAL GERAL POR GRUPO DE PRODUTOS</P></TD>" & chr(13) & _
                "   </TR>" & CHR(13)

        for cont = 0 to UBound(vt_TotalGrupo)
            if vt_TotalGrupo(cont).c1 <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & _
                    "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
                    "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
                    "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    				"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
                    "	</TR>" & chr(13)
                    saldo_grupo = 0
                    saldo_dinamico_grupo = 0
                    vl_custo_entrada_grupo = 0
                end if
                if vt_TotalGrupo(cont).c1 <> "-" then 
                    input_Id_Total_Grupo = "TotalGeral_Grupo__" & vt_TotalGrupo(cont).c1
                else
                    input_Id_Total_Grupo = "TotalGeral_Grupo__"
                end if
                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MB ME' colspan='2' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            vl_custo_entrada_grupo = vl_custo_entrada_grupo + vt_TotalGrupo(cont).c4
            saldo_dinamico_grupo = saldo_dinamico_grupo + vt_TotalGrupo(cont).c5
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
            "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
            "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    		"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
			"	</TR>" & chr(13)

        end if
    end if

		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='7'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='3' align='right' valign='bottom' NOWRAP class='MC MEB'><P class='Cd' style='font-weight:bold;'>TOTAL GERAL:</P></TD>" & chr(13) & _
                    "		<TD NOWRAP class='MC MB' valign='bottom'><P class='Cd' id='input_Id_Total_Saldo' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_total)=0,"",formata_inteiro(saldo_dinamico_total)) & "</P></TD>" & chr(13) & _
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
			"		<TD colspan='6' align='center'><P class='ALERTA'>NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)

    x = x & "<input type='hidden' id='qtde_fabricantes' value='" & qtde_fabricantes & "' />" & chr(13)
	
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

	s_sql = "SELECT" & _
            " t_PRODUTO.grupo AS grupo," & _
            " t_PRODUTO_GRUPO.descricao AS grupo_descricao," & _
            " t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html" & _
			", Sum(qtde-qtde_utilizada) AS saldo, t_ESTOQUE_ITEM.vl_custo2" & _
			" FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON" & _
			" ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
            " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE.id_estoque = t_ESTOQUE_ITEM.id_estoque)" & _
			" WHERE ((qtde-qtde_utilizada) > 0)"
	
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

    s_where_temp = ""
	if c_subgrupo <> "" then
	    v_subgrupos = split(c_subgrupo, ", ")
	    for cont = Lbound(v_subgrupos) to Ubound(v_subgrupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.subgrupo = '" & v_subgrupos(cont) & "')"
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
			" AND (t_ESTOQUE.id_nfe_emitente = '" & c_empresa & "')"
		end if

	s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_PRODUTO.grupo, t_PRODUTO_GRUPO.descricao, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html, t_ESTOQUE_ITEM.vl_custo2"

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then
		        s_sql = s_sql &	" ORDER BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html, t_ESTOQUE_ITEM.vl_custo2"
    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
                s_sql = s_sql & " ORDER BY t_ESTOQUE_ITEM.fabricante, t_PRODUTO.grupo, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html, t_ESTOQUE_ITEM.vl_custo2"
    end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
	'mantendo a padronização com as demais opções de emissão do relatório em Excel (sem bordas no cabeçalho)
	if blnSaidaExcel then
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD width='75' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='350' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
			  "		<TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />UNITÁRIO</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />TOTAL</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
	else
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
			  "		<TD width='60' valign='bottom' NOWRAP class='MDBE'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='274' valign='bottom' NOWRAP class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
			  "		<TD width='60' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "		<TD width='100' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA UNITÁRIO</P></TD>" & chr(13) & _
			  "		<TD width='100' valign='bottom' NOWRAP class='MDB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA TOTAL</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
		end if
	
	x = cab_table & cab
	n_reg = 0
	n_saldo_total = 0
	n_saldo_parcial = 0
	vl_total_geral = 0
	vl_sub_total = 0
	qtde_fabricantes = 0
    saldo_grupo = 0
    n_ctrl_reg = 0
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
                if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDBE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

	 '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)
	
	 '> CUSTO DE ENTRADA UNITÁRIO
		vl = r("vl_custo2")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	 '> CUSTO DE ENTRADA TOTAL
		vl = r("vl_custo2")*r("saldo")
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)
		
		vl_total_geral = vl_total_geral + vl
		vl_sub_total = vl_sub_total + vl

		n_saldo_total = n_saldo_total + r("saldo")
		n_saldo_parcial = n_saldo_parcial + r("saldo")
		
		x = x & "	</TR>" & chr(13)

          elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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
	
	     '> CUSTO DE ENTRADA UNITÁRIO
		    vl = r("vl_custo2")
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	     '> CUSTO DE ENTRADA TOTAL
		    vl = r("vl_custo2")*r("saldo")
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
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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
			x = x & "	<TR NOWRAP><TD COLSPAN='5'>&nbsp;</TD></TR>" & chr(13) & _
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

' ____________________________________________________
' CONSULTA ESTOQUE VENDA + SHOW-ROOM DETALHE SINTETICO
'
sub consulta_estoque_venda_show_room_detalhe_sintetico
dim r
dim s, s_aux, s_sql, s_nbsp, x, cab_table, cab, fabricante_a, grupo_a, saldo_grupo, n_ctrl_reg
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes, s_where_temp, cont
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_TotalGrupo(), vt_QtdeSaldoDinamicoGrupo()
dim saldo_dinamico, saldo_dinamico_parcial, saldo_dinamico_grupo, s_checked, s_bg_realce, saldo_dinamico_total
dim input_Id_Subtotal_Grupo, input_Id_Subtotal_Fabricante, input_Id_Total_Grupo


	s_sql = "SELECT grupo, grupo_descricao, fabricante, produto, descricao, descricao_html, SUM(saldo) AS saldo, saldo_flag FROM (" & _
            " SELECT grupo, grupo_descricao, fabricante, produto, descricao , descricao_html, SUM(saldo) AS saldo, saldo_flag FROM (SELECT t_PRODUTO.grupo AS grupo, t_PRODUTO_GRUPO.descricao AS grupo_descricao, SUM(t_ESTOQUE_ITEM.qtde-t_ESTOQUE_ITEM.qtde_utilizada) AS saldo," & _
            " t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html," & _
            " (SELECT TOP 1 Flag FROM t_RELATORIO_PRODUTO_FLAG" & _
	        "   WHERE (Usuario = '" & usuario & "' AND Fabricante = t_ESTOQUE_ITEM.fabricante AND Produto = t_ESTOQUE_ITEM.produto AND IdPagina = " & Relatorio_Id & ")" & _
            " ) AS saldo_flag" & _
            " FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
            " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE.id_estoque = t_ESTOQUE_ITEM.id_estoque)" & _
            " WHERE ((t_ESTOQUE_ITEM.qtde-t_ESTOQUE_ITEM.qtde_utilizada) > 0)"
    
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

    s_where_temp = ""
	if c_subgrupo <> "" then
	    v_subgrupos = split(c_subgrupo, ", ")
	    for cont = Lbound(v_subgrupos) to Ubound(v_subgrupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.subgrupo = '" & v_subgrupos(cont) & "')"
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
			" AND (t_ESTOQUE.id_nfe_emitente = '" & c_empresa & "')"
		end if

            s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_PRODUTO.grupo, t_PRODUTO_GRUPO.descricao, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html" & _
            " UNION ALL SELECT t_PRODUTO.grupo AS grupo, t_PRODUTO_GRUPO.descricao AS grupo_descricao, SUM(t_ESTOQUE_MOVIMENTO.qtde) AS saldo, " & _
            " t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html," & _
            " (SELECT TOP 1 Flag FROM t_RELATORIO_PRODUTO_FLAG" & _
	        "   WHERE (Usuario = '" & usuario & "' AND Fabricante = t_ESTOQUE_MOVIMENTO.fabricante AND Produto = t_ESTOQUE_MOVIMENTO.produto AND IdPagina = " & Relatorio_Id & ")" & _
            " ) AS saldo_flag" & _
            " FROM t_ESTOQUE_MOVIMENTO" & _
            " LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
            " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
            " WHERE (anulado_status=0 AND estoque='SHR')"

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

    s_where_temp = ""
	if c_subgrupo <> "" then
	    v_subgrupos = split(c_subgrupo, ", ")
	    for cont = Lbound(v_subgrupos) to Ubound(v_subgrupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.subgrupo = '" & v_subgrupos(cont) & "')"
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

            s_sql = s_sql & " GROUP BY t_ESTOQUE_MOVIMENTO.fabricante, t_PRODUTO.grupo, t_PRODUTO_GRUPO.descricao, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, descricao_html) tbl"
            
    s_sql = s_sql & " GROUP BY tbl.fabricante, tbl.grupo, tbl.grupo_descricao, tbl.produto, tbl.descricao, tbl.descricao_html, tbl.saldo_flag" & _
                    ") tbl2" & _
                    " GROUP BY tbl2.fabricante, tbl2.grupo, tbl2.grupo_descricao, tbl2.produto, tbl2.descricao, tbl2.descricao_html, tbl2.saldo_flag"

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then
		s_sql = s_sql & " ORDER BY tbl2.fabricante, tbl2.produto, tbl2.descricao, tbl2.descricao_html"
    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        s_sql = s_sql & " ORDER BY tbl2.fabricante, tbl2.grupo, tbl2.produto, tbl2.descricao, tbl2.descricao_html"
    end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
	
'	Mantendo a padronização com as demais opções de emissão do relatório em Excel (sem bordas no cabeçalho)
	if blnSaidaExcel then
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
              "		<TD width='30' valign='bottom' NOWRAP><P class='R' style='font-weight:bold;'></P></TD>" & chr(13) & _
			  "		<TD width='75' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='480' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
              "     <TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'></P></TD>" & chr(13) & _  
			  "		<TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
	else
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
              "		<TD width='30' valign='bottom' NOWRAP class='MD ME MB'><P class='R' style='font-weight:bold;'></P></TD>" & chr(13) & _
			  "		<TD width='75' valign='bottom' NOWRAP class='MD MB'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='480' valign='bottom' NOWRAP class='MD MB'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
              "     <TD width='60' align='right' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'></P></TD>" & chr(13) & _  
			  "		<TD width='60' align='right' valign='bottom' NOWRAP class='MDB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
		end if
	
	x = cab_table & cab
	n_reg = 0
	n_saldo_total = 0
	n_saldo_parcial = 0
    saldo_dinamico = 0
    saldo_dinamico_parcial = 0
    saldo_dinamico_grupo = 0
    saldo_dinamico_total = 0
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
                if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
		            x = x & "		<TD class='MB ME' colspan='3' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)                    
                    vt_QtdeGrupo(cont) = saldo_grupo
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='3' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
                        "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
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
            saldo_dinamico_parcial = 0
            grupo_a = "XXXXXXX"
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then

		s_checked = ""
        s_bg_realce = ""
        saldo_dinamico = 0
        if Trim("" & r("saldo_flag")) = "1" then 
            s_checked = " checked"
            s_bg_realce = "style='background-color: lightgreen'"
            saldo_dinamico = formata_inteiro(r("saldo"))
        end if

        input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))
        input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("grupo")) & "_" & Trim("" & r("fabricante"))
        input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante"))

		x = x & "	<TR NOWRAP>" & chr(13)
    
    '> FLAG
        if Not blnSaidaExcel then
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">" & chr(13) & _
                    "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """, null, null, ""col_saldo_" & n_reg & """, """ & r("saldo") & """)' />" & chr(13) & _
                    "       </TD>" & chr(13)
        else
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
        end if

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

     '> SALDO (DINÂMICO)
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

	 '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

		n_saldo_total = n_saldo_total + r("saldo")
		n_saldo_parcial = n_saldo_parcial + r("saldo")
        saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
        saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico
		
		x = x & "	</TR>" & chr(13)

        elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        if Trim("" & r("grupo")) <> grupo_a then
            if n_ctrl_reg > 0 then
		        x = x & "		<TD class='MB ME' colspan='3' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
                x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		        x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                x = x & "   </TR>" & chr(13)

                redim preserve vt_QtdeGrupo(cont)                    
                vt_QtdeGrupo(cont) = saldo_grupo
                redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
                cont = cont + 1
            end if
            x = x & "	<TR NOWRAP>" & chr(13)
            x = x & "       <TD class='MDBE' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C'>&nbsp;</P></TD>" & chr(13)
		    x = x & "		<TD class='MDB' colspan='5' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)

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
            saldo_dinamico_grupo = 0 
            n_ctrl_reg = n_ctrl_reg + 1
        end if

        s_checked = ""
        s_bg_realce = ""
        saldo_dinamico = 0
        if Trim("" & r("saldo_flag")) = "1" then 
            s_checked = " checked"
            s_bg_realce = "style='background-color: lightgreen'"
            saldo_dinamico = formata_inteiro(r("saldo"))
        end if

        input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))        
        input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("fabricante")) & "_" & Trim("" & r("grupo"))
        input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante"))

        x = x & "   <TR NOWRAP>" & chr(13)

        '> FLAG
        if Not blnSaidaExcel then
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">"  & chr(13) & _
                    "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """,""" & input_Id_Subtotal_Grupo & """, """ & input_Id_Total_Grupo & """, ""col_saldo_" & n_reg & """, """ & r("saldo") & """)' />" & chr(13) & _
                    "       </TD>" & chr(13)
        else
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
        end if

    	'> PRODUTO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	    '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

        '> SALDO (DINÂMICO)
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

 	    '> SALDO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

        x = x & "   </TR>" & chr(13)

        saldo_grupo = saldo_grupo + CInt(r("saldo"))
        n_saldo_parcial = n_saldo_parcial + r("saldo")
		n_saldo_total = n_saldo_total + r("saldo")
        saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
        saldo_dinamico_grupo = saldo_dinamico_grupo + saldo_dinamico
        saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico

        end if

		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop

   ' MOSTRA TOTAL DO ÚLTIMO GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        if n_reg <> 0 then
            x = x & "		<TD class='MB ME' colspan='3' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>Total do Grupo:</P></TD>" & chr(13)
            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		    x = x & "		<TD class='MDB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)                    
            vt_QtdeGrupo(cont) = saldo_grupo
            redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
            vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo
        end if
    end if
		
  ' MOSTRA TOTAL
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO FORNECEDOR
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MEB' colspan='3' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
                "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

  ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
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
                    .c4 = vt_QtdeSaldoDinamicoGrupo(cont)
                end with
            next
        end if
        ordena_cl_quatro_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        saldo_dinamico_grupo = 0
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
                "		<TD class='MTBD ME' valign='bottom' colspan='5' NOWRAP><P class='R' style='font-weight:bold;'>TOTAL GERAL POR GRUPO DE PRODUTOS</P></TD>" & chr(13) & _
                "   </TR>" & CHR(13)

        for cont = 0 to UBound(vt_TotalGrupo)
            if vt_TotalGrupo(cont).c1 <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & _
                    "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
                    "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
				    "	</TR>" & chr(13)
                    saldo_grupo = 0
                    saldo_dinamico_grupo = 0
                end if

                if vt_TotalGrupo(cont).c1 <> "-" then 
                    input_Id_Total_Grupo = "TotalGeral_Grupo__" & vt_TotalGrupo(cont).c1
                else
                    input_Id_Total_Grupo = "TotalGeral_Grupo__"
                end if

                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MDB ME' colspan='2' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            saldo_dinamico_grupo = saldo_dinamico_grupo + vt_TotalGrupo(cont).c4
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
            "		<TD class='MDB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
			"	</TR>" & chr(13)

        end if
    end if
		
		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='5'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='3' NOWRAP class='MC MEB' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>" & "TOTAL GERAL:" & "</P></TD>" & chr(13) & _
                    "		<TD NOWRAP class='MC MB' valign='bottom'><P class='Cd' id='input_Id_Total_Saldo' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_total)=0,"",formata_inteiro(saldo_dinamico_total)) & "</P></TD>" & chr(13) & _
					"		<TD NOWRAP class='MC MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
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

    x = x & "<input type='hidden' id='qtde_fabricantes' value='" & qtde_fabricantes & "' />" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	
end sub

' ________________________________________________________
' CONSULTA ESTOQUE VENDA + SHOW-ROOM DETALHE INTERMEDIARIO
'
sub consulta_estoque_venda_show_room_detalhe_intermediario
dim r
dim s, s_aux, s_sql, s_nbsp, x, cab_table, cab, fabricante_a, grupo_a, saldo_grupo, vl_custo_entrada_grupo, n_ctrl_reg
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes, s_where_temp, cont
dim vl, vl_total_geral, vl_sub_total
dim vt_CodigoGrupo(), vt_DescricaoGrupo(), vt_QtdeGrupo(), vt_TotalGrupo(), vt_QtdeSaldoDinamicoGrupo()
dim saldo_dinamico, saldo_dinamico_parcial, saldo_dinamico_grupo, s_checked, s_bg_realce, saldo_dinamico_total
dim input_Id_Subtotal_Grupo, input_Id_Subtotal_Fabricante, input_Id_Total_Grupo

	s_sql = "SELECT grupo, grupo_descricao, fabricante, produto, descricao, descricao_html, SUM(saldo) AS saldo, SUM(preco_total) AS preco_total, saldo_flag FROM (" & _
                " SELECT grupo, grupo_descricao, fabricante, produto, descricao, descricao_html, SUM(saldo) AS saldo, SUM(preco_total) AS preco_total, saldo_flag FROM (SELECT t_PRODUTO.grupo AS grupo, t_PRODUTO_GRUPO.descricao AS grupo_descricao, SUM(t_ESTOQUE_ITEM.qtde-t_ESTOQUE_ITEM.qtde_utilizada) AS saldo," & _
                " t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, Sum((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total," & _
                " (SELECT TOP 1 Flag FROM t_RELATORIO_PRODUTO_FLAG" & _
	            "   WHERE (Usuario = '" & usuario & "' AND Fabricante = t_ESTOQUE_ITEM.fabricante AND Produto = t_ESTOQUE_ITEM.produto AND IdPagina = " & Relatorio_Id & ")" & _
                " ) AS saldo_flag" & _
                " FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
                " INNER JOIN t_ESTOQUE ON (t_ESTOQUE.id_estoque = t_ESTOQUE_ITEM.id_estoque)" & _
                " WHERE ((t_ESTOQUE_ITEM.qtde-t_ESTOQUE_ITEM.qtde_utilizada) > 0)"

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

    s_where_temp = ""
	if c_subgrupo <> "" then
	    v_subgrupos = split(c_subgrupo, ", ")
	    for cont = Lbound(v_subgrupos) to Ubound(v_subgrupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.subgrupo = '" & v_subgrupos(cont) & "')"
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
			" AND (t_ESTOQUE.id_nfe_emitente = '" & c_empresa & "')"
		end if

    s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_PRODUTO.grupo, t_PRODUTO_GRUPO.descricao, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, descricao_html" & _
                " UNION ALL SELECT t_PRODUTO.grupo AS grupo, t_PRODUTO_GRUPO.descricao AS grupo_descricao, SUM(t_ESTOQUE_MOVIMENTO.qtde) AS saldo," & _
                " t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, Sum(t_ESTOQUE_MOVIMENTO.qtde*t_ESTOQUE_ITEM.vl_custo2) AS preco_total," & _
                " (SELECT TOP 1 Flag FROM t_RELATORIO_PRODUTO_FLAG" & _
	            "   WHERE (Usuario = '" & usuario & "' AND Fabricante = t_ESTOQUE_MOVIMENTO.fabricante AND Produto = t_ESTOQUE_MOVIMENTO.produto AND IdPagina = " & Relatorio_Id & ")" & _
                " ) AS saldo_flag" & _
                " FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
                " AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
                " LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
                " LEFT JOIN t_PRODUTO_GRUPO ON (t_PRODUTO.grupo=t_PRODUTO_GRUPO.codigo)" & _
                " WHERE (anulado_status=0 AND estoque='SHR')"
                
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

    s_where_temp = ""
	if c_subgrupo <> "" then
	    v_subgrupos = split(c_subgrupo, ", ")
	    for cont = Lbound(v_subgrupos) to Ubound(v_subgrupos)
	        if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.subgrupo = '" & v_subgrupos(cont) & "')"
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

    s_sql = s_sql & " GROUP BY t_ESTOQUE_MOVIMENTO.fabricante, t_PRODUTO.grupo, t_PRODUTO_GRUPO.descricao, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, descricao_html) tbl" & _
                    " GROUP BY tbl.fabricante, tbl.grupo, tbl.grupo_descricao, tbl.produto, tbl.descricao, tbl.descricao_html, tbl.saldo_flag" & _
                    ") tbl2"

    s_sql = s_sql & " GROUP BY tbl2.fabricante, tbl2.grupo, tbl2.grupo_descricao, tbl2.produto, tbl2.descricao, tbl2.descricao_html, tbl2.saldo_flag"
     
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then
		s_sql = s_sql & " ORDER BY tbl2.fabricante, tbl2.produto, tbl2.descricao, tbl2.descricao_html"
    elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        s_sql = s_sql & " ORDER BY tbl2.fabricante, tbl2.grupo, tbl2.produto, tbl2.descricao, tbl2.descricao_html"
    end if

  ' CABEÇALHO
	cab_table = "<TABLE class='MC' cellSpacing=0>" & chr(13)
'	Mantendo a padronização com as demais opções de emissão do relatório em Excel (sem bordas no cabeçalho)
	if blnSaidaExcel then
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
              "		<TD width='30' valign='bottom' NOWRAP><P class='R' style='font-weight:bold;'></P></TD>" & chr(13) & _
			  "		<TD width='75' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='350' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
              "     <TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'></P></TD>" & chr(13) & _  
			  "		<TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
			  "		<TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />TOTAL</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
	else
		cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
              "		<TD width='30' valign='bottom' NOWRAP class='MD ME MB'><P class='R' style='font-weight:bold;'></P></TD>" & chr(13) & _
			  "		<TD width='60' valign='bottom' NOWRAP class='MDB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
			  "		<TD width='274' valign='bottom' class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
              "     <TD width='60' align='right' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'></P></TD>" & chr(13) & _  
			  "		<TD width='60' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			  "		<TD width='100' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
			  "		<TD width='100' valign='bottom' NOWRAP class='MDB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA TOTAL</P></TD>" & chr(13) & _
			  "	</TR>" & chr(13)
		end if
	
	x = cab_table & cab
	n_reg = 0
	n_saldo_total = 0
	n_saldo_parcial = 0
    saldo_dinamico = 0
    saldo_dinamico_parcial = 0
    saldo_dinamico_grupo = 0
    saldo_dinamico_total = 0
	vl_total_geral = 0
	vl_sub_total = 0
	qtde_fabricantes = 0
	saldo_grupo = 0
    n_ctrl_reg = 0
	fabricante_a = "XXXXX"
    grupo_a = "XXXXX"
		
	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then

            ' força o fechamento do último grupo de produtos
                if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
		            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='3' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont)
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo           
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                    n_ctrl_reg = 0
                end if

				x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='3' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
                        "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)
				
				if blnSaidaExcel then
					x = x & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='7' class='MB'>&nbsp;</TD>" & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='7'>&nbsp;</TD>" & _
						"	</TR>" & chr(13)
				else
					x = x & _
						"	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='7' class='MB'>&nbsp;</TD>" & _
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
					"		<TD class='MDBE' align='center' valign='bottom' colspan='7' style='background: honeydew'><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			n_saldo_parcial = 0
            saldo_dinamico_parcial = 0
			vl_sub_total = 0
            grupo_a = "XXXXXXX"
			end if

	  ' CONTAGEM
		n_reg = n_reg + 1

    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__PRODUTO then

		s_checked = ""
        s_bg_realce = ""
        saldo_dinamico = 0
        if Trim("" & r("saldo_flag")) = "1" then 
            s_checked = " checked"
            s_bg_realce = "style='background-color: lightgreen'"
            saldo_dinamico = formata_inteiro(r("saldo"))
        end if

        input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))
        input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("grupo")) & "_" & Trim("" & r("fabricante"))
        input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante"))

		x = x & "	<TR NOWRAP>" & chr(13)
    
    '> FLAG
        if Not blnSaidaExcel then
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">" & chr(13) & _
                    "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """, null, null, ""col_saldo_" & n_reg & """, """ & r("saldo") & """)' />" & chr(13) & _
                    "       </TD>" & chr(13)
        else
            x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
        end if

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

     '> SALDO (DINÂMICO)
		x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

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
        saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
        saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico
		
		x = x & "	</TR>" & chr(13)

       elseif rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
            if Trim("" & r("grupo")) <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='3' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)

                    redim preserve vt_QtdeGrupo(cont)         
                    redim preserve vt_CustoEntrada(cont) 
                    redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
                    vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo          
                    vt_QtdeGrupo(cont) = saldo_grupo
                    vt_CustoEntrada(cont) = vl_custo_entrada_grupo
                    cont = cont + 1
                end if
                x = x & "	<TR NOWRAP>" & chr(13)
                x = x & "       <TD class='MDBE' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C'>&nbsp;</P></TD>" & chr(13)
		        x = x & "		<TD class='MDB ME' colspan='6' align='left' style='background-color: #EEE' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("grupo")) & " - " & Trim("" & r("grupo_descricao")) & "</P></TD>" & chr(13)


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
                saldo_dinamico_grupo = 0 
                vl_custo_entrada_grupo = 0  
                n_ctrl_reg = n_ctrl_reg + 1
            end if

            s_checked = ""
            s_bg_realce = ""
            saldo_dinamico = 0
            if Trim("" & r("saldo_flag")) = "1" then 
                s_checked = " checked"
                s_bg_realce = "style='background-color: lightgreen'"
                saldo_dinamico = formata_inteiro(r("saldo"))
            end if

            input_Id_Total_Grupo = "TotalGeral_Grupo__" & Trim("" & r("grupo"))        
            input_Id_Subtotal_Grupo = "Subtotal_Grupo__" & Trim("" & r("fabricante")) & "_" & Trim("" & r("grupo"))
            input_Id_Subtotal_Fabricante = "Subtotal_Fabricante__" & Trim("" & r("fabricante"))

            x = x & "   <TR NOWRAP>" & chr(13)

            '> FLAG
            if Not blnSaidaExcel then
                x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP " & s_bg_realce & ">"  & chr(13) & _
                        "           <input name='ckb_saldo_flag' id='ckb_saldo_flag' type='checkbox'" & s_checked & " onchange='enviaRequisicaoAjaxRelatorioProdutoFlag(this, """ & Trim("" & r("fabricante")) & """, """ & Trim("" & r("produto")) & """, """ & input_Id_Subtotal_Fabricante & """,""" & input_Id_Subtotal_Grupo & """, """ & input_Id_Total_Grupo & """, ""col_saldo_" & n_reg & """, """ & r("saldo") & """)' />" & chr(13) & _
                        "       </TD>" & chr(13)
            else
                x = x & "		<TD class='MDBE' align='center' valign='middle' NOWRAP></TD>" & chr(13)
            end if

	     '> PRODUTO
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & Trim("" & r("produto")) & "</P></TD>" & chr(13)

 	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

         '> SALDO (DINÂMICO)
		    x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' id='col_saldo_" & n_reg & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico)=0, "", formata_inteiro(saldo_dinamico)) & "</P></TD>" & chr(13)

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
            saldo_dinamico_parcial = saldo_dinamico_parcial + saldo_dinamico
            saldo_dinamico_grupo = saldo_dinamico_grupo + saldo_dinamico
            saldo_dinamico_total = saldo_dinamico_total + saldo_dinamico

            x = x & "	</TR>" & chr(13)
        end if

		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop

  ' MOSTRA TOTAL DO ÚLTIMO GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        if n_reg <> 0 then
            x = x & "   <TR NOWRAP>" & chr(13)
                    x = x & "       <TD class='MB ME' align='right' valign='middle' colspan='3' NOWRAP><P class='Cd'>Total do Grupo:</P></TD>" & chr(13)
                    x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' id='" & input_Id_Subtotal_Grupo & "' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</P></TD>" & chr(13)
		            x = x & "		<TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</P></TD>" & chr(13)
                    x = x & "       <TD class='MB' align='right' valign='middle' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13)
		            x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13)
                    x = x & "   </TR>" & chr(13)
            x = x & "   </TR>" & chr(13)

            redim preserve vt_QtdeGrupo(cont)         
            redim preserve vt_CustoEntrada(cont) 
            redim preserve vt_QtdeSaldoDinamicoGrupo(cont)
            vt_QtdeSaldoDinamicoGrupo(cont) = saldo_dinamico_grupo          
            vt_QtdeGrupo(cont) = saldo_grupo
            vt_CustoEntrada(cont) = vl_custo_entrada_grupo
        end if
    end if
		
  ' MOSTRA TOTAL
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO FORNECEDOR
		x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MEB' colspan='3' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>TOTAL:</P></TD>" & chr(13) & _
                        "		<TD class='MB' valign='bottom'><P class='Cd' id='" & input_Id_Subtotal_Fabricante & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_parcial)=0,"",formata_inteiro(saldo_dinamico_parcial)) & "</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						"		<TD class='MDB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

  ' TOTAL GERAL POR GRUPO DE PRODUTOS
    if rb_tipo_agrupamento = COD_TIPO_AGRUPAMENTO__GRUPO then
        redim vt_TotalGrupo(0)
        set vt_TotalGrupo(0) = New cl_CINCO_COLUNAS
        with vt_TotalGrupo(0)
            .c1 = ""
            .c2 = ""
            .c3 = ""
            .c4 = ""
            .c5 = ""
        end with

        if vt_CodigoGrupo(UBound(vt_CodigoGrupo)) <> "" then
            for cont = 0 to UBound(vt_CodigoGrupo) 
                if Trim(vt_TotalGrupo(UBound(vt_TotalGrupo)).c1) <> "" then
                    redim preserve vt_TotalGrupo(UBound(vt_TotalGrupo)+1)
                    set vt_TotalGrupo(UBound(vt_TotalGrupo)) = New cl_CINCO_COLUNAS
                end if
                with vt_TotalGrupo(UBound(vt_TotalGrupo))
                    .c1 = vt_CodigoGrupo(cont)
                    .c2 = vt_DescricaoGrupo(cont)
                    .c3 = vt_QtdeGrupo(cont)
                    .c4 = vt_CustoEntrada(cont)
                    .c5 = vt_QtdeSaldoDinamicoGrupo(cont)
                end with
            next
        end if
        ordena_cl_cinco_colunas vt_TotalGrupo, 0, UBound(vt_TotalGrupo)

        saldo_grupo = 0
        vl_custo_entrada_grupo = 0
        n_ctrl_reg = 0
        grupo_a = "XXXXXXX"

        if (qtde_fabricantes > 1) then
            x = x & _
                "	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='7'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD colspan='7'>&nbsp;</TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:azure'>" & chr(13) & _
                "		<TD class='MTBD ME' valign='bottom' colspan='7' NOWRAP><P class='R' style='font-weight:bold;'>TOTAL GERAL POR GRUPO DE PRODUTOS</P></TD>" & chr(13) & _
                "   </TR>" & CHR(13)

        for cont = 0 to UBound(vt_TotalGrupo)
            if vt_TotalGrupo(cont).c1 <> grupo_a then
                if n_ctrl_reg > 0 then
                    x = x & _
                    "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
                    "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
                    "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    				"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
                    "	</TR>" & chr(13)
                    saldo_grupo = 0
                    saldo_dinamico_grupo = 0
                    vl_custo_entrada_grupo = 0
                end if

                if vt_TotalGrupo(cont).c1 <> "-" then 
                    input_Id_Total_Grupo = "TotalGeral_Grupo__" & vt_TotalGrupo(cont).c1
                else
                    input_Id_Total_Grupo = "TotalGeral_Grupo__"
                end if


                x = x & _
                "   <TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD class='MEB' align='left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left;margin-left: 10px'>" & vt_TotalGrupo(cont).c1 & "</p></TD>" & chr(13) & _
                "		<TD class='MB ME' colspan='2' style='text-align: left' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;text-align: left'>" & vt_TotalGrupo(cont).c2 & "</p></TD>" & chr(13)
				
            end if
            saldo_grupo = saldo_grupo + vt_TotalGrupo(cont).c3
            vl_custo_entrada_grupo = vl_custo_entrada_grupo + vt_TotalGrupo(cont).c4
            saldo_dinamico_grupo = saldo_dinamico_grupo + vt_TotalGrupo(cont).c5            
            grupo_a = vt_TotalGrupo(cont).c1     
            n_ctrl_reg = n_ctrl_reg + 1  
        next

        ' total geral último grupo
        x = x & _
            "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' id='" & input_Id_Total_Grupo & "' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_grupo)=0,"",formata_inteiro(saldo_dinamico_grupo)) & "</p></TD>" & chr(13) & _
            "		<TD class='MB' valign='bottom' NOWRAP><p class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(saldo_grupo) & "</p></TD>" & chr(13) & _
            "       <TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;</P></TD>" & chr(13) & _
    		"		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_custo_entrada_grupo) & "</P></TD>" & chr(13) & _    
			"	</TR>" & chr(13)

        end if
    end if

		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='7'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='3' align='right' valign='bottom' NOWRAP class='MC MEB'><P class='Cd' style='font-weight:bold;'>TOTAL GERAL:</P></TD>" & chr(13) & _
                    "		<TD NOWRAP class='MC MB' valign='bottom'><P class='Cd' id='input_Id_Total_Saldo' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & iif(CInt(saldo_dinamico_total)=0,"",formata_inteiro(saldo_dinamico_total)) & "</P></TD>" & chr(13) & _
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
			"		<TD colspan='7' align='center'><P class='ALERTA'>NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)

    x = x & "<input type='hidden' id='qtde_fabricantes' value='" & qtde_fabricantes & "' />" & chr(13)
	
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



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
    window.status = 'Aguarde, executando a consulta ...';
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
<style type="text/css">
    P.C {
        font-size: 10pt;
    }

    P.Cc {
        font-size: 10pt;
    }

    P.Cd {
        font-size: 10pt;
    }

    P.F {
        font-size: 11pt;
    }
</style>
<% end if %>

<script type="text/javascript">
    function recalculaSubTotalFabricanteGrupo(checkboxElement, input_Id_Subtotal_Fabricante, input_Id_Subtotal_Grupo, input_Id_Total_Grupo, col_saldo, saldo) {
        var qtde_fabricantes;
        var Subtotal_Fabricante, Subtotal_Grupo, Total_Grupo, Total_Geral;

        Total_Geral = document.getElementById("input_Id_Total_Saldo");
        Subtotal_Fabricante = document.getElementById(input_Id_Subtotal_Fabricante);
        Subtotal_Grupo = input_Id_Subtotal_Grupo != null ? document.getElementById(input_Id_Subtotal_Grupo) : null;
        Total_Grupo = input_Id_Total_Grupo != null ? document.getElementById(input_Id_Total_Grupo) : null;

        qtde_fabricantes = parseInt(document.getElementById("qtde_fabricantes").value);

        if (checkboxElement.checked) {
            checkboxElement.parentElement.style.background = "lightgreen";
            document.getElementById(col_saldo).innerText = formata_numero(saldo);
            Subtotal_Fabricante.innerText = formata_numero((parseInt(retorna_so_digitos(Subtotal_Fabricante.innerText)) || 0) + parseInt(saldo));
            if (qtde_fabricantes > 1) Total_Geral.innerText = formata_numero((parseInt(retorna_so_digitos(Total_Geral.innerText)) || 0) + parseInt(saldo));
            if (input_Id_Subtotal_Grupo != null) {
                Subtotal_Grupo.innerText = formata_numero((parseInt(retorna_so_digitos(Subtotal_Grupo.innerText)) || 0) + parseInt(saldo));
                if (qtde_fabricantes > 1) Total_Grupo.innerText = formata_numero((parseInt(retorna_so_digitos(Total_Grupo.innerText)) || 0) + parseInt(saldo));
            }
        }
        else {
            checkboxElement.parentElement.style.background = "white";
            document.getElementById(col_saldo).innerText = "";
            Subtotal_Fabricante.innerText = formata_numero(parseInt(retorna_so_digitos(Subtotal_Fabricante.innerText)) - parseInt(saldo));
            Subtotal_Fabricante.innerText = Subtotal_Fabricante.innerText == "0" ? "" : Subtotal_Fabricante.innerText;
            if (qtde_fabricantes > 1) {
                Total_Geral.innerText = formata_numero(parseInt(retorna_so_digitos(Total_Geral.innerText)) - parseInt(saldo));
                Total_Geral.innerText = Total_Geral.innerText == "0" ? "" : Total_Geral.innerText;
            }
            if (input_Id_Subtotal_Grupo != null) {
                Subtotal_Grupo.innerText = formata_numero(parseInt(retorna_so_digitos(Subtotal_Grupo.innerText)) - parseInt(saldo));
                Subtotal_Grupo.innerText = Subtotal_Grupo.innerText == "0" ? "" : Subtotal_Grupo.innerText;
                if (qtde_fabricantes > 1) {
                    Total_Grupo.innerText = formata_numero(parseInt(retorna_so_digitos(Total_Grupo.innerText)) - parseInt(saldo));
                    Total_Grupo.innerText = Total_Grupo.innerText == "0" ? "" : Total_Grupo.innerText;
                }
            }
        }
    }
</script>
<script type="text/javascript">
    function enviaRequisicaoAjaxRelatorioProdutoFlag(checkboxElement, fabricante, produto, input_Id_Subtotal_Fabricante, input_Id_Subtotal_Grupo, input_Id_Total_Grupo, col_saldo, saldo) {
        var serverVariableUrl;
        var relatorioId;
        var usuarioRel;
        var Flag;
        var parameters;
        serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
        serverVariableUrl = serverVariableUrl.toUpperCase();
        serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));
        relatorioId = "<%=Relatorio_Id%>";
        usuarioRel = "<%=usuario%>";
        Flag = 0;

        if (checkboxElement.checked)
            Flag = 1;

        parameters = "?paginaId=" + relatorioId + "&usuario=" + usuarioRel + "&codFabricante=" + fabricante + "&codProduto=" + produto + "&flag=" + Flag;

        $.ajax({
        	url: 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/PostData/RelatorioProdutoFlagPost' + parameters,
            type: "POST",
            async: true
        })
        .success(function (response) {
            recalculaSubTotalFabricanteGrupo(checkboxElement, input_Id_Subtotal_Fabricante, input_Id_Subtotal_Grupo, input_Id_Total_Grupo, col_saldo, saldo);
        })
        .fail(function (jqXHR, textStatus) {
            if (checkboxElement.checked) {
                checkboxElement.checked = false;
            }
            else {
                checkboxElement.checked = true;
            }
            alert("Falha ao gravar o flag do produto checado\nMarque o produto novamente!!");
        });
    }
</script>



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
<body onload="window.status='Concluído';" link="#000000" alink="#000000" vlink="#000000">

    <center>

<form id="fESTOQ" name="fESTOQ" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="rb_estoque" id="rb_estoque" value="<%=rb_estoque%>">
<input type="hidden" name="rb_detalhe" id="rb_detalhe" value="<%=rb_detalhe%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque II</span>
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
<table class="Qx" cellspacing="0">
<!--  ESTOQUE  -->
	<tr bgcolor="#FFFFFF">
		<% select case rb_estoque
				case ID_ESTOQUE_VENDA:			s = "VENDA"
				case ID_ESTOQUE_VENDIDO:		s = "VENDIDO"
				case ID_ESTOQUE_SHOW_ROOM:		s = "SHOW-ROOM"
				case ID_ESTOQUE_DANIFICADOS:	s = "PRODUTOS DANIFICADOS"
				case ID_ESTOQUE_DEVOLUCAO:		s = "DEVOLUÇÃO"
                case VENDA_SHOW_ROOM:           s = "VENDA + SHOW-ROOM"
				case else						s = ""
				end select
		%>
	<td class="MT" nowrap><span class="PLTe">Estoque de Interesse</span>
		<br><p class="C" style="width:230px;cursor:default;"><%=s%></p></td>

<!--  TIPO DE DETALHAMENTO  -->
		<% select case rb_detalhe
			case "SINTETICO":		s = "SINTÉTICO (SEM VALOR)"
			case "INTERMEDIARIO":	s = "INTERMEDIÁRIO (VALOR MÉDIO)"
			case "COMPLETO":		s = "COMPLETO (VALOR DIFERENCIADO)"
			case else				s = ""
			end select
		%>
	<td class="MT" style="border-left:0px;" nowrap><span class="PLTe">Tipo de Detalhamento</span>
		<br><p class="C" style="width:230px;cursor:default;"><%=s%></p></td>
	</tr>

<!--  EMPRESA  -->
	<% if c_empresa <> "" then %>
		<tr bgcolor="#FFFFFF">
		<td class="MDBE" nowrap colspan="2"><span class="PLTe">Empresa</span>			
			<br><span class="C">
                <%=obtem_apelido_empresa_NFe_emitente(c_empresa)%>          
			    </span></td>
		</tr>
	<% end if %>

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
		<tr bgcolor="#FFFFFF">
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
    elseif rb_estoque = VENDA_SHOW_ROOM then
        select case rb_detalhe
            case "SINTETICO"
                consulta_estoque_venda_show_room_detalhe_sintetico
            case "INTERMEDIARIO"
                consulta_estoque_venda_show_room_detalhe_intermediario
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
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
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
