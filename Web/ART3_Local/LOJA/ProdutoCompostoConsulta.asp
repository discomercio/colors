<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  ProdutoCompostoConsulta.asp
'     ========================================================
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

	dim fabricante_selecionado, produto_selecionado, descricao_fornecida

'	PRODUTO COMPOSTO
	fabricante_selecionado = trim(request("c_prod_comp_fabricante"))
	produto_selecionado = trim(request("c_prod_comp_produto"))
	
	fabricante_selecionado=retorna_so_digitos(fabricante_selecionado)
	produto_selecionado=retorna_so_digitos(produto_selecionado)

	fabricante_selecionado=normaliza_codigo(fabricante_selecionado, TAM_MIN_FABRICANTE)
	produto_selecionado=normaliza_produto(produto_selecionado)
	
	if (fabricante_selecionado="") Or (fabricante_selecionado="000") then Response.Redirect("aviso.asp?id=" & ERR_FABRICANTE_NAO_ESPECIFICADO)
	if (produto_selecionado="") Or (produto_selecionado="000000") then Response.Redirect("aviso.asp?id=" & ERR_EC_PRODUTO_COMPOSTO_NAO_ESPECIFICADO)

	dim v_item
	dim strFabricante, strProduto, strQtde, strDescricao

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim max_qtde_itens
	max_qtde_itens = obtem_parametro_ProdutoComposto_MaxQtdeItens

	dim alerta
	alerta = ""

	dim strSql
	strSql = "SELECT " & _
				"*" & _
			" FROM t_EC_PRODUTO_COMPOSTO" & _
			" WHERE" & _
				" (fabricante_composto = '" & fabricante_selecionado & "')" & _
				" AND (produto_composto='" & produto_selecionado & "')"
	rs.Open strSql, cn
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if Not rs.EOF then
		descricao_fornecida = rs("descricao")
	else
		'Verifica se é um produto normal
		strSql = "SELECT " & _
					"*" & _
				" FROM t_PRODUTO" & _
				" WHERE" & _
					" (fabricante = '" & fabricante_selecionado & "')" & _
					" AND (produto = '" & produto_selecionado & "')"
		if rs.State <> 0 then rs.Close
		rs.Open strSql, cn
		if rs.Eof then
			alerta = "Produto (" & fabricante_selecionado & ") " & produto_selecionado & " não está cadastrado!"
		else
			alerta = "Código informado é de um produto normal:<br />(" & fabricante_selecionado & ") " & produto_selecionado & " - " & Trim("" & rs("descricao"))
			end if
		end if

	dim i, n, msg_erro
	
	if alerta = "" then
		if Not le_EC_produto_composto_item(fabricante_selecionado, produto_selecionado, v_item, msg_erro) then
			alerta = "Falha ao ler os itens do produto composto."
			if msg_erro <> "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & msg_erro
				end if
			end if
		end if
	
	if alerta = "" then
		'Assegura que dados cadastrados anteriormente sejam exibidos corretamente, mesmo se o parâmetro da quantidade máxima de itens tiver sido reduzido
		if VectorLength(v_item) > max_qtde_itens then max_qtde_itens = VectorLength(v_item)
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
					rs.Open strSql, cn
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
	<title>LOJA</title>
	</head>

<script src="../GLOBAL/global.js" language="JavaScript" type="text/javascript"></script>


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
<link href="../global/eprinter.css" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
	.TdFabr {
		width: 40px;
	}
	.TdProd {
		width: 60px;
	}
	.TdQtde {
		width: 50px;
	}
	.TdDescr {
		width: 400px;
	}
</style>


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Consulta de Produto Composto</span></td>
</tr>
</table>
<br />
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>

<% else %>
<!-- ************************************************************ -->
<!-- **********      PÁGINA PARA CONSULTAR      ********** -->
<!-- ************************************************************ -->

<body onload="focus();">
<center>

<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Consulta de Produto Composto</span></td>
</tr>
</table>
<br />


<!-- ************   CÓDIGO / DESCRIÇÃO   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td class="MD" width="15%" align="left"><p class="R">FABRICANTE</p><p class="C"><input id="fabricante_selecionado" name="fabricante_selecionado" class="TA" value="<%=fabricante_selecionado%>" readonly size="6" style="text-align:center; color=#0000ff"></p></td>
		<td class="MD" width="15%" align="left"><p class="R">PRODUTO</p><p class="C"><input id="produto_selecionado" name="produto_selecionado" class="TA" value="<%=produto_selecionado%>" readonly size="10" style="text-align:center; color=#0000ff"></p></td>
		<td width="70%" align="left"><p class="R">DESCRIÇÃO</p><p class="C"><input id="descricao_fornecida" name="descricao_fornecida" class="TA" type="text" maxlength="80" size="60" value="<%=descricao_fornecida%>" /></p></td>
	</tr>
</table>

<br><br><p class="F" style="margin-bottom:5px;">Composição de 1 unidade do produto composto</p>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0" cellpadding="4">
	<tr bgcolor="#FFFFFF">
	<td class="MB TdFabr" align="left"><span class="PLTe">Fabr</span></td>
	<td class="MB TdProd" align="left"><span class="PLTe">Produto</span></td>
	<td class="MB TdQtde" align="right"><span class="PLTd">Qtde</span></td>
	<td class="MB TdDescr" align="left"><span class="PLTe">Descrição</span></td>
	</tr>
<% 
	n = Lbound(v_item)-1
	for i=1 to max_qtde_itens
		n = n+1
		if n <= Ubound(v_item) then
			with v_item(n) 
				strFabricante = Trim(.fabricante_item)
				strProduto = Trim(.produto_item)
				strQtde = Cstr(.qtde)
				strDescricao = Trim(.descricao)
				end with
		else
			strFabricante = ""
			strProduto = ""
			strQtde = ""
			strDescricao = ""
			end if

		if strProduto <> "" then
%>
	<tr>
		<td class="MDBE TdFabr" align="left"><span class="C"><%=strFabricante%></span></td>
		<td class="MDB TdProd" align="left"><span class="C"><%=strProduto%></span></td>
		<td class="MDB TdQtde" align="right"><span class="C"><%=strQtde%></span></td>
		<td class="MDB TdDescr" align="left"><span class="C"><%=strDescricao%></span></td>
	</tr>
<%
		end if
	next 
%>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center"><a href="javascript:history.back()" title="Voltar">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>

</center>
</body>
<% end if %>

</html>


<%

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
	cn.Close
	set cn = nothing

%>
