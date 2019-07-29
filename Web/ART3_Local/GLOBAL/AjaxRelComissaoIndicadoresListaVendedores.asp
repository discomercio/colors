<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     ==========================================================================================
'	   A J A X R E L C O M I S S Ã O I N D I C A D O R E S L I S T A V E N D E D O R E S . A S P
'     ==========================================================================================
'
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

' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
'	OBTEM O ID
	dim strSql, strResp, msg_erro

    const VENDA_NORMAL = "VEN"
    const DEVOLUCAO = "DEV"
    const PERDA = "PER"
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s, mes, ano, id_a, vendedor, data

	ano = Trim(Request("ano"))
	mes = Trim(Request("mes"))
    data = "1/" & mes & "/" & ano
    data = DateAdd("m", 1, data)
    data = Cstr(data)
    data = strToDate(data)

	strSql = "SELECT DISTINCT usuario, " & _
	        "nome, " & _
	        "nome_iniciais_em_maiusculas " & _
        "FROM (" & _
	        "SELECT usuario, " & _
		        "nome, " & _
		        "nome_iniciais_em_maiusculas " & _
	        "FROM t_USUARIO " & _
	        "WHERE (vendedor_loja <> 0) " & _
	        "UNION " & _
	        "SELECT t_USUARIO.usuario AS usuario, " & _
		    "    t_USUARIO.nome AS nome, " & _
		    "    t_USUARIO.nome_iniciais_em_maiusculas " & _
	        "FROM t_USUARIO " & _
	        "INNER JOIN t_PERFIL_X_USUARIO " & _
		    "    ON (t_USUARIO.usuario = t_PERFIL_X_USUARIO.usuario) " & _
	        "INNER JOIN t_PERFIL " & _
		    "    ON (t_PERFIL_X_USUARIO.id_perfil = t_PERFIL.id) " & _
	        "INNER JOIN t_PERFIL_ITEM " & _
		    "    ON (t_PERFIL.id = t_PERFIL_ITEM.id_perfil) " & _
	        " WHERE (t_PERFIL_ITEM.id_operacao = 10100) " & _
	        ") AS tU " & _
        "WHERE usuario IN ( " & _
		    "    SELECT DISTINCT vendedor " & _
		    "    FROM ( " & _
			"        SELECT '" & VENDA_NORMAL & "' AS operacao, " & _
			"	        t_PEDIDO.vendedor, " & _
			"	        Sum(t_PEDIDO_ITEM.qtde * t_PEDIDO_ITEM.preco_venda) AS total_preco_venda, " & _
			"	        Sum(t_PEDIDO_ITEM.qtde * t_PEDIDO_ITEM.preco_NF) AS total_preco_NF " & _
			"        FROM t_PEDIDO " & _
			"        INNER JOIN t_PEDIDO AS t_PEDIDO__BASE " & _
			"	        ON (t_PEDIDO.pedido_base = t_PEDIDO__BASE.pedido) " & _
			"        INNER JOIN t_PEDIDO_ITEM " & _
			"	        ON (t_PEDIDO.pedido = t_PEDIDO_ITEM.pedido) " & _
			"        INNER JOIN t_CLIENTE " & _
			"	        ON (t_PEDIDO__BASE.id_cliente = t_CLIENTE.id) " & _
			"        LEFT JOIN t_ORCAMENTISTA_E_INDICADOR " & _
			"	        ON (t_PEDIDO.indicador = t_ORCAMENTISTA_E_INDICADOR.apelido) " & _
			"        WHERE (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "') " & _
			"	        AND (Coalesce(t_PEDIDO.indicador, '') <> '') " & _
			"	        AND (t_PEDIDO.entregue_data < " & bd_formata_data(data) & ") " & _
			"	        AND (((t_PEDIDO.comissao_paga = 0))) " & _
			"	        AND (t_PEDIDO.pedido = t_PEDIDO.pedido_base) " & _
			"	        AND (((t_PEDIDO__BASE.st_pagto = 'S') AND (t_PEDIDO__BASE.dt_st_pagto < " & bd_formata_data(data) & "))) " & _
			"        GROUP BY t_PEDIDO.vendedor " & _
			"        UNION " & _
			"        SELECT '" & DEVOLUCAO & "' AS operacao, " & _
			"	        t_PEDIDO.vendedor, " & _
			"	        Sum(- t_PEDIDO_ITEM_DEVOLVIDO.qtde * t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS total_preco_venda, " & _
			"	        Sum(- t_PEDIDO_ITEM_DEVOLVIDO.qtde * t_PEDIDO_ITEM_DEVOLVIDO.preco_NF) AS total_preco_NF " & _
			"        FROM t_PEDIDO " & _
			"        INNER JOIN t_PEDIDO AS t_PEDIDO__BASE " & _
			"	        ON (t_PEDIDO.pedido_base = t_PEDIDO__BASE.pedido) " & _
			"        INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO " & _
			"	        ON (t_PEDIDO.pedido = t_PEDIDO_ITEM_DEVOLVIDO.pedido) " & _
			"        INNER JOIN t_CLIENTE " & _
			"	        ON (t_PEDIDO__BASE.id_cliente = t_CLIENTE.id) " & _
			"        LEFT JOIN t_ORCAMENTISTA_E_INDICADOR " & _
			"	        ON (t_PEDIDO.indicador = t_ORCAMENTISTA_E_INDICADOR.apelido) " & _
			"       WHERE (Coalesce(t_PEDIDO.indicador, '') <> '') " & _
			"	        AND (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(data) & ") " & _
			"	        AND (((comissao_descontada = 0))) " & _
			"        GROUP BY t_PEDIDO.vendedor " & _
			"        UNION		 " & _
			"        SELECT '" & PERDA & "' AS operacao, " & _
			"	        t_PEDIDO.vendedor, " & _
			"	        Sum(- t_PEDIDO_PERDA.valor) AS total_preco_venda, " & _
			"	        Sum(- t_PEDIDO_PERDA.valor) AS total_preco_NF " & _
			"        FROM t_PEDIDO " & _
			"        INNER JOIN t_PEDIDO AS t_PEDIDO__BASE " & _
			"	        ON (t_PEDIDO.pedido_base = t_PEDIDO__BASE.pedido) " & _
			"        INNER JOIN t_PEDIDO_PERDA " & _
			"	        ON (t_PEDIDO.pedido = t_PEDIDO_PERDA.pedido) " & _
			"       INNER JOIN t_CLIENTE " & _
			"	        ON (t_PEDIDO__BASE.id_cliente = t_CLIENTE.id) " & _
			"        LEFT JOIN t_ORCAMENTISTA_E_INDICADOR " & _
			"	        ON (t_PEDIDO.indicador = t_ORCAMENTISTA_E_INDICADOR.apelido) " & _
			"        WHERE (Coalesce(t_PEDIDO.indicador, '') <> '') " & _
			"	        AND (t_PEDIDO_PERDA.data < " & bd_formata_data(data) & ") " & _
			"	        AND (((comissao_descontada = 0))) " & _
			"        GROUP BY t_PEDIDO.vendedor " & _
			"        ) t " & _
		    "    WHERE (" & _
			"	        (total_preco_venda <> 0) " & _
			"	        OR (total_preco_NF <> 0) " & _
			"	        ) " & _
			"        AND ( " & _
			"	        vendedor NOT IN (" & _
			"		        SELECT vendedor " & _
			"		        FROM t_COMISSAO_INDICADOR_N2 " & _
			"		        WHERE competencia_ano = " & ano & _
			"			        AND competencia_mes = " & mes & _
			"		        )" & _
			"	        )" & _
		     "   ) " & _
        "ORDER BY usuario"


'	EXECUTA A CONSULTA
	strResp = ""
	
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
'	ENCONTROU DADOS
	do while Not rs.Eof

        if strResp <> "" then strResp = strResp & "</option>"
            
		    strResp = strResp & _
				  "<option value='" & UCase(Trim("" & rs("usuario"))) & "'>"
             strResp = strResp & UCase(Trim("" & rs("usuario"))) & " - " & Trim("" & rs("nome_iniciais_em_maiusculas"))

		rs.MoveNext
		loop
	
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
