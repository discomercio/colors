<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  ECProdutoCompostoEditaConfirma.asp
'     ===========================================
'
'
'      SSSSSSS   EEEEEEEEE  RRRRRRRR   VVV   VVV  IIIII  DDDDDDDD    OOOOOOO   RRRRRRRR
'     SSS   SSS  EEE        RRR   RRR  VVV   VVV   III   DDD   DDD  OOO   OOO  RRR   RRR
'      SSS       EEE        RRR   RRR  VVV   VVV   III   DDD   DDD  OOO   OOO  RRR   RRR
'       SSSS     EEEEEE     RRRRRRRR   VVV   VVV   III   DDD   DDD  OOO   OOO  RRRRRRRR
'          SSS   EEE        RRR RRR     VVV VVV    III   DDD   DDD  OOO   OOO  RRR RRR
'     SSS   SSS  EEE        RRR  RRR     VVVVV     III   DDD   DDD  OOO   OOO  RRR  RRR
'      SSSSSSS   EEEEEEEEE  RRR   RRR     VVV     IIIII  DDDDDDDD    OOOOOOO   RRR   RRR



' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P     N O     S E R V I D O R
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
	dim s, s_aux, usuario
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, r, strSql, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim vl_preco_fabricante, vl_custo2, vl_preco_lista

	dim s_log, s_log_aux
	s_log = ""

	dim alerta
	alerta = ""

	dim erro_consistencia, erro_fatal
	erro_consistencia=false
	erro_fatal=false
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, fabricante_selecionado, produto_selecionado, descricao_fornecida
	operacao_selecionada=request("operacao_selecionada")
	fabricante_selecionado = trim(request("fabricante_selecionado"))
	produto_selecionado = trim(request("produto_selecionado"))
	descricao_fornecida = trim(request("descricao_fornecida"))
	
	fabricante_selecionado=retorna_so_digitos(fabricante_selecionado)
	produto_selecionado=retorna_so_digitos(produto_selecionado)

	fabricante_selecionado=normaliza_codigo(fabricante_selecionado, TAM_MIN_FABRICANTE)
	produto_selecionado=normaliza_produto(produto_selecionado)
	
	if (fabricante_selecionado="") Or (fabricante_selecionado="000") then Response.Redirect("aviso.asp?id=" & ERR_FABRICANTE_NAO_ESPECIFICADO) 
	if (produto_selecionado="") Or (produto_selecionado="000000") then Response.Redirect("aviso.asp?id=" & ERR_EC_PRODUTO_COMPOSTO_NAO_ESPECIFICADO) 

	dim i, j, n, intIdx
	dim v_item, v_item_bd
	redim v_item(0)
	set v_item(0) = new cl_EC_ITEM_PRODUTO_COMPOSTO
	n = Request.Form("c_produto_item").Count

	for i = 1 to n
		s=Trim(Request.Form("c_produto_item")(i))
		if s <> "" then
			if Trim(v_item(ubound(v_item)).produto_item) <> "" then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_EC_ITEM_PRODUTO_COMPOSTO
				end if
			with v_item(ubound(v_item))
				.produto_item=Ucase(Trim(Request.Form("c_produto_item")(i)))
				s=retorna_so_digitos(Request.Form("c_fabricante_item")(i))
				.fabricante_item=normaliza_codigo(s, TAM_MIN_FABRICANTE)
				s = Trim(Request.Form("c_qtde_item")(i))
				if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
				end with
			end if
		next

'	PRODUTO COMPOSTO ESTÁ CADASTRADO?
	strSql = "SELECT " & _
				"*" & _
			" FROM t_EC_PRODUTO_COMPOSTO" & _
			" WHERE" & _
				" (fabricante_composto = '" & fabricante_selecionado & "')" & _
				" AND (produto_composto='" & produto_selecionado & "')"
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_EC_PRODUTO_COMPOSTO_NAO_CADASTRADO)
	
	if alerta = "" then
		if Not le_EC_produto_composto_item(fabricante_selecionado, produto_selecionado, v_item_bd, msg_erro) then
			alerta = "Falha ao ler os itens do produto composto."
			end if
		end if

	if operacao_selecionada <> OP_EXCLUI then
	'	CONSISTE ITENS DO PRODUTO COMPOSTO
		if alerta = "" then
			for i=Lbound(v_item) to Ubound(v_item)
				with v_item(i)
					if .qtde <= 0 then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto_item & " do fabricante " & .fabricante_item & ": quantidade " & cstr(.qtde) & " é inválida."
						end if
					
					for j=Lbound(v_item) to (i-1)
						if (.produto_item = v_item(j).produto_item) And (.fabricante_item = v_item(j).fabricante_item) then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto " & .produto_item & " do fabricante " & .fabricante_item & ": linha " & renumera_com_base1(Lbound(v_item),i) & " repete o mesmo produto da linha " & renumera_com_base1(Lbound(v_item),j) & "."
							exit for
							end if
						next
					end with
				next
			end if

		if alerta = "" then
			for i=Lbound(v_item) to Ubound(v_item)
				with v_item(i)
					if Trim(.produto_item) <> "" then
						strSql = "SELECT " & _
									"*" & _
								" FROM t_PRODUTO" & _
								" WHERE" & _
									" (fabricante = '" & Trim(.fabricante_item) & "')" & _
									" AND (produto = '" & Trim(.produto_item) & "')"
						if rs.State <> 0 then rs.Close
						rs.open strSql, cn
						if Not rs.Eof then
							.descricao = Trim("" & rs("descricao"))
						else
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto " & .produto_item & " do fabricante " & .fabricante_item & " NÃO está cadastrado na tabela de produtos."
							end if
						end if
					end with
				next
			end if
		end if


	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'    =========
			if Not erro_fatal then
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				s_log = "Exclusão do produto composto " & produto_selecionado & " (" & fabricante_selecionado & ") formado por:"
				for i=Lbound(v_item_bd) to Ubound(v_item_bd)
					with v_item_bd(i)
						if Trim(.produto_item) <> "" then
							s_log = s_log & " " & formata_inteiro(.qtde) & "x" & Trim(.produto_item) & "(" & Trim(.fabricante_item) & ")"
							end if
						end with
					next

			'	APAGA!!
				strSql = "DELETE" & _
						 " FROM t_EC_PRODUTO_COMPOSTO_ITEM" & _
						 " WHERE" & _
							" (fabricante_composto = '" & fabricante_selecionado & "')" & _
							" AND (produto_composto = '" & produto_selecionado & "')"
				cn.Execute(strSql)
				
				strSql = "DELETE" & _
						 " FROM t_EC_PRODUTO_COMPOSTO" & _
						 " WHERE" & _
							" (fabricante_composto = '" & fabricante_selecionado & "')" & _
							" AND (produto_composto = '" & produto_selecionado & "')"
				cn.Execute(strSql)
				
				If Err = 0 then 
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_EC_PRODUTO_COMPOSTO_EXCLUSAO, s_log
				else
					erro_fatal=True
					alerta = "FALHA AO REMOVER O PRODUTO COMPOSTO (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				
				if alerta = "" then
				'	~~~~~~~~~~~~~~
					cn.CommitTrans
				'	~~~~~~~~~~~~~~
					if Err<>0 then 
						alerta=Cstr(Err) & ": " & Err.Description
						end if
				else
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					end if
				end if


		case else
		'    ====
		'	 ATUALIZA!!
			if alerta = "" then
			'	LOG DA EDIÇÃO
				s_log = "Produto composto " & produto_selecionado & " (" & fabricante_selecionado & ") alterado de ["
				s_log_aux = ""
				for i=Lbound(v_item_bd) to Ubound(v_item_bd)
					with v_item_bd(i)
						if Trim(.produto_item) <> "" then
							if s_log_aux <> "" then s_log_aux = s_log_aux & " "
							s_log_aux = s_log_aux & formata_inteiro(.qtde) & "x" & Trim(.produto_item) & "(" & Trim(.fabricante_item) & ")"
							end if
						end with
					next

				s_log = s_log & s_log_aux & "] para ["
				
				s_log_aux = ""
				for i=Lbound(v_item) to Ubound(v_item)
					with v_item(i)
						if Trim(.produto_item) <> "" then
							if s_log_aux <> "" then s_log_aux = s_log_aux & " "
							s_log_aux = s_log_aux & formata_inteiro(.qtde) & "x" & Trim(.produto_item) & "(" & Trim(.fabricante_item) & ")"
							end if
						end with
					next
				
				s_log = s_log & s_log_aux & "]"

			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				strSql = "SELECT " & _
							"*" & _
						" FROM t_EC_PRODUTO_COMPOSTO" & _
						" WHERE" & _
							" (fabricante_composto = '" & fabricante_selecionado & "')" & _
							" AND (produto_composto = '" & produto_selecionado & "')"
				if r.State <> 0 then r.Close
				r.Open strSql, cn
				if Not r.Eof then
					if UCase(Trim("" & r("descricao"))) <> UCase(descricao_fornecida) then
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & " descrição: " & chr(34) & Trim("" & r("descricao")) & chr(34) & " => " & chr(34) & descricao_fornecida & chr(34)
						r("descricao") = descricao_fornecida
						r("dt_ult_atualizacao") = Now
						r("usuario_ult_atualizacao") = usuario
						r.Update
						end if
					end if

				strSql = "UPDATE t_EC_PRODUTO_COMPOSTO_ITEM" & _
						 " SET" & _
							" excluido_status = 1" & _
						 " WHERE" & _
							" (fabricante_composto = '" & fabricante_selecionado & "')" & _
							" AND (produto_composto = '" & produto_selecionado & "')"
				cn.Execute(strSql)
				
				intIdx = 0
				for i=Lbound(v_item) to Ubound(v_item)
					with v_item(i)
						if Trim(.produto_item) <> "" then
							intIdx = intIdx + 1
							strSql = "SELECT " & _
										"*" & _
									" FROM t_EC_PRODUTO_COMPOSTO_ITEM" & _
									" WHERE" & _
										" (fabricante_composto = '" & fabricante_selecionado & "')" & _
										" AND (produto_composto = '" & produto_selecionado & "')" & _
										" AND (fabricante_item = '" & Trim(.fabricante_item) & "')" & _
										" AND (produto_item = '" & Trim(.produto_item) & "')"
							if r.State <> 0 then r.Close
							r.Open strSql, cn
							if r.Eof then
								r.AddNew
								r("fabricante_composto") = fabricante_selecionado
								r("produto_composto") = produto_selecionado
								r("fabricante_item") = Trim(.fabricante_item)
								r("produto_item") = Trim(.produto_item)
								r("dt_cadastro") = Now
								r("usuario_cadastro") = usuario
								end if
								
							r("dt_ult_atualizacao") = Now
							r("usuario_ult_atualizacao") = usuario
							r("qtde") = CLng(.qtde)
							r("sequencia") = intIdx
							r("excluido_status") = 0
							r.Update
							end if
						end with
					next

			'	EXCLUI OS ITENS RETIRADOS DA COMPOSIÇÃO
				strSql = "DELETE" & _
						 " FROM t_EC_PRODUTO_COMPOSTO_ITEM" & _
						 " WHERE" & _
							" (fabricante_composto = '" & fabricante_selecionado & "')" & _
							" AND (produto_composto = '" & produto_selecionado & "')" & _
							" AND (excluido_status = 1)"
				cn.Execute(strSql)
				
			'	FINALIZA TRANSAÇÃO
				If Err = 0 then 
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_EC_PRODUTO_COMPOSTO_ALTERACAO, s_log
				else
					erro_fatal=True
					alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
					end if

				if alerta = "" then
				'	~~~~~~~~~~~~~~
					cn.CommitTrans
				'	~~~~~~~~~~~~~~
					if Err<>0 then 
						alerta=Cstr(Err) & ": " & Err.Description
						end if
				else
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					end if

				if r.State <> 0 then r.Close
				set r = nothing
				end if
		
		end select
		
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
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>



<!-- C A S C A D I N G   S T Y L E   S H E E T

	 CCCCCCC    SSSSSSS    SSSSSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	CCC        SSS        SSS
	CCC         SSSS       SSSS
	CCC            SSSS       SSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	 CCCCCCC    SSSSSSS    SSSSSSS
-->
<link href="../global/e.css" rel="stylesheet" type="text/css">


<body onload="bVOLTAR.focus();">
<center>
<br>

<!--  T E L A  -->

<p class="T">A V I S O</p>

<% 
	s = ""
	s_aux="'MtAviso'"
	if alerta <> "" then
		s = "<p style='margin:5px 2px 5px 2px;'>" & alerta & "</p>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_EXCLUI
				s = "PRODUTO COMPOSTO " & chr(34) & produto_selecionado & " (" & fabricante_selecionado & ")" & " - " & descricao_fornecida & chr(34) & " EXCLUÍDO COM SUCESSO."
			case else
				s = "PRODUTO COMPOSTO " & chr(34) & produto_selecionado & " (" & fabricante_selecionado & ")" & " - " & descricao_fornecida & chr(34) & " ATUALIZADO COM SUCESSO."
			end select
		if s <> "" then s="<p style='margin:5px 2px 5px 2px;'>" & s & "</p>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<br><br>

<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
<%
	s="ECProdutoCompostoMenu.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<td align="center"><a name="bVOLTAR" href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>
</body>
</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>