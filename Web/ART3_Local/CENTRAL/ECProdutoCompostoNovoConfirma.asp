<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  ECProdutoCompostoNovoConfirma.asp
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
	
	dim s, s_aux, usuario, alerta, strSql
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_log
	s_log = ""

	dim vl_preco_fabricante, vl_custo2, vl_preco_lista
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim fabricante_selecionado, produto_selecionado, descricao_original, descricao_fornecida
	fabricante_selecionado = trim(request("fabricante_selecionado"))
	produto_selecionado = trim(request("produto_selecionado"))
	descricao_fornecida = trim(request("descricao_fornecida"))

	fabricante_selecionado=retorna_so_digitos(fabricante_selecionado)
	produto_selecionado=retorna_so_digitos(produto_selecionado)

	fabricante_selecionado=normaliza_codigo(fabricante_selecionado, TAM_MIN_FABRICANTE)
	produto_selecionado=normaliza_produto(produto_selecionado)

	if (fabricante_selecionado="") Or (fabricante_selecionado="000") then Response.Redirect("aviso.asp?id=" & ERR_FABRICANTE_NAO_ESPECIFICADO)
	if (produto_selecionado="") Or (produto_selecionado="000000") then Response.Redirect("aviso.asp?id=" & ERR_EC_PRODUTO_COMPOSTO_NAO_ESPECIFICADO)

	dim i, j, n, intQtdeUnidades
	dim v_item
	redim v_item(0)
	set v_item(0) = new cl_EC_ITEM_PRODUTO_COMPOSTO
	n = Request.Form("c_produto_item").Count

	intQtdeUnidades=0
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
				intQtdeUnidades=intQtdeUnidades+.qtde
				end with
			end if
		next

	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if fabricante_selecionado = "" then
		alerta="INFORME O CÓDIGO DO FABRICANTE DO PRODUTO COMPOSTO."
	elseif produto_selecionado = "" then
		alerta="INFORME O CÓDIGO DO PRODUTO COMPOSTO."
	elseif intQtdeUnidades < 2 then
		alerta="UM PRODUTO COMPOSTO DEVE CONTER 2 OU MAIS UNIDADES DE PRODUTOS."
		end if

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

	strSql = "SELECT " & _
				"*" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (fabricante = '" & fabricante_selecionado & "')" & _
				" AND (produto = '" & produto_selecionado & "')"
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
	if Not rs.Eof then
		descricao_original = Trim("" & rs("descricao"))
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
	
	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
'	EXECUTA OPERAÇÃO NO BD
	if alerta = "" then 
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		s_log="Cadastramento de produto composto: " & produto_selecionado & " (" & fabricante_selecionado & ") formado por:"
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if Trim(.produto_item) <> "" then
					s_log = s_log & " " & formata_inteiro(.qtde) & "x" & Trim(.produto_item) & "(" & .fabricante_item & ")"
					end if
				end with
			next
			
		strSql = "INSERT INTO t_EC_PRODUTO_COMPOSTO (" & _
					"fabricante_composto, " & _
					"produto_composto, " & _
					"descricao, " & _
					"dt_cadastro, " & _
					"usuario_cadastro, " & _
					"dt_ult_atualizacao, " & _
					"usuario_ult_atualizacao" & _
				") VALUES (" & _
					"'" & fabricante_selecionado & "', " & _
					"'" & produto_selecionado & "', " & _
					"'" & descricao_fornecida & "', " & _
					"getdate(), " & _
					"'" & usuario & "', " & _
					"getdate(), " & _
					"'" & usuario & "'" & _
				")"
		cn.Execute(strSql)

		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				strSql = "INSERT INTO t_EC_PRODUTO_COMPOSTO_ITEM (" & _
							"fabricante_composto, " & _
							"produto_composto, " & _
							"fabricante_item, " & _
							"produto_item, " & _
							"qtde, " & _
							"sequencia, " & _
							"dt_cadastro, " & _
							"usuario_cadastro, " & _
							"dt_ult_atualizacao, " & _
							"usuario_ult_atualizacao" & _
						") VALUES (" & _
							"'" & fabricante_selecionado & "', " & _
							"'" & produto_selecionado & "', " & _
							"'" & Trim(.fabricante_item) & "', " & _
							"'" & Trim(.produto_item) & "', " & _
							Cstr(.qtde) & ", " &_ 
							Cstr(i+1) & ", " & _
							"getdate(), " &_ 
							"'" & usuario & "', " &_
							"getdate(), " & _
							"'" & usuario & "'" & _
						")"
				cn.Execute(strSql)
				end with
			next
		
	'	FINALIZA TRANSAÇÃO
		If Err = 0 then
			if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_EC_PRODUTO_COMPOSTO_INCLUSAO, s_log
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

		end if
		
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
		s = "PRODUTO COMPOSTO " & chr(34) & produto_selecionado & " (" & fabricante_selecionado & ")" & " - " & descricao_fornecida & chr(34) & " FOI CADASTRADO COM SUCESSO."
		s="<p style='margin:5px 2px 5px 2px;'>" & s & "</p>"
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