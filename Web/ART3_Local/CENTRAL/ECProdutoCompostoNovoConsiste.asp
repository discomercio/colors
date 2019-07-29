<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  ECProdutoCompostoNovoConsiste.asp
'     =====================================
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
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
	dim usuario, fabricante_selecionado, produto_selecionado, descricao_original, descricao_fornecida
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim alerta
	alerta = ""
	
'	PRODUTO COMPOSTO A EDITAR
	fabricante_selecionado = trim(request("fabricante_selecionado"))
	produto_selecionado = trim(request("produto_selecionado"))
	descricao_fornecida = trim(request("descricao_fornecida"))
	
	fabricante_selecionado=retorna_so_digitos(fabricante_selecionado)
	produto_selecionado=retorna_so_digitos(produto_selecionado)

	fabricante_selecionado=normaliza_codigo(fabricante_selecionado, TAM_MIN_FABRICANTE)
	produto_selecionado=normaliza_produto(produto_selecionado)
	
	if (fabricante_selecionado="") Or (fabricante_selecionado="000") then Response.Redirect("aviso.asp?id=" & ERR_FABRICANTE_NAO_ESPECIFICADO)
	if (produto_selecionado="") Or (produto_selecionado="000000") then Response.Redirect("aviso.asp?id=" & ERR_EC_PRODUTO_COMPOSTO_NAO_ESPECIFICADO)
	
	dim s, i, j, n
	dim v_item
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

	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .qtde <= 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto_item & " do fabricante " & .fabricante_item & ": quantidade " & cstr(.qtde) & " é inválida."
					end if
				
				if (.fabricante_item <> fabricante_selecionado) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Linha " & renumera_com_base1(Lbound(v_item),i) & " informa um fabricante diferente do selecionado para o produto composto."
					exit for
					end if
				
				if (.produto_item = produto_selecionado) And (.fabricante_item = fabricante_selecionado) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Linha " & renumera_com_base1(Lbound(v_item),i) & " informa o código do próprio produto composto como sendo um de seus componentes."
					exit for
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
    
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, strSql
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	strSql = "SELECT " & _
				"*" & _
			" FROM t_EC_PRODUTO_COMPOSTO" & _
			" WHERE" & _
				" (fabricante_composto = '" & fabricante_selecionado & "')" & _
				" AND (produto_composto='" & produto_selecionado & "')"
	set rs = cn.Execute(strSql)
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_EC_PRODUTO_COMPOSTO_JA_CADASTRADO)

	strSql = "SELECT " & _
				"*" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (fabricante = '" & fabricante_selecionado & "')" & _
				" AND (produto = '" & produto_selecionado & "')"
	set rs = cn.Execute(strSql)
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
					set rs = cn.Execute(strSql)
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
    if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if Trim(.produto_item) <> "" then
					strSql = "SELECT " & _
								"*" & _
							" FROM t_EC_PRODUTO_COMPOSTO" & _
							" WHERE" & _
								" (fabricante_composto = '" & Trim(.fabricante_item) & "')" & _
								" AND (produto_composto = '" & Trim(.produto_item) & "')"
					set rs = cn.Execute(strSql)
					if Not rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Não é possivel cadastrar o produto " & .produto_item & " do fabricante " & .fabricante_item & ", pois ele está cadastrado como produto composto"						
						end if
					end if
				end with
			next
		end if
%>


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>


<%
'		C L I E N T   S I D E   S C R I P T I N G
'
'      CCCCCCC   LLL        IIIII  EEEEEEEEE  NNN   NNN  TTTTTTTTT EEEEEEEEE
'     CCC   CCC  LLL         III   EEE        NNNN  NNN     TTT    EEE
'     CCC        LLL         III   EEE        NNNNN NNN     TTT    EEE
'     CCC        LLL         III   EEEEEE     NNN NNNNN     TTT    EEEEEE
'     CCC        LLL         III   EEE        NNN  NNNN     TTT    EEE
'     CCC   CCC  LLL   LLL   III   EEE        NNN   NNN     TTT    EEE
'      CCCCCCC   LLLLLLLLL  IIIII  EEEEEEEEE  NNN   NNN     TTT    EEEEEEEEE
'
%>

<script src="../GLOBAL/global.js" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript">
function AtualizaProdutoComposto( f ) {
var i,blnTemDado,intQtdeUnidades;
	intQtdeUnidades=0;
	for (i=0; i<f.c_produto_item.length; i++) {
		blnTemDado=false;
		if (trim(f.c_fabricante_item[i].value)!="") blnTemDado=true;
		if (trim(f.c_produto_item[i].value)!="") blnTemDado =true;
		if (converte_numero(f.c_qtde_item[i].value)>0) blnTemDado=true;
		if (blnTemDado) {
			if (trim(f.c_fabricante_item[i].value)=="") {
				alert('Informe o fabricante do item do produto composto!!');
				f.c_fabricante_item[i].focus();
				return;
				}
			if (trim(f.c_produto_item[i].value)=="") {
				alert('Informe o produto do item do produto composto!!');
				f.c_produto_item[i].focus();
				return;
				}
			if (converte_numero(f.c_qtde_item[i].value)==0) {
				alert('Informe a quantidade do item do produto composto!!');
				f.c_qtde_item[i].focus();
				return;
				}
			intQtdeUnidades=intQtdeUnidades+converte_numero(f.c_qtde_item[i].value);
			}
		}
		
	if (intQtdeUnidades < 2) {
		alert('Um produto composto deve conter 2 ou mais unidades de produtos!!');
		return;
		}

	f.descricao_fornecida.value = f.c_produto_composto_descricao.value;

	dATUALIZA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
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
<link href="../global/e.css" rel="stylesheet" type="text/css">
<link href="../global/eprinter.css" rel="stylesheet" type="text/css" media="print">


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
	<td align="center"><a name="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>

<% else %>
<!-- ************************************************************ -->
<!-- **********      PÁGINA PARA CADASTRAR/EDITAR      ********** -->
<!-- ************************************************************ -->

<body onload="fCAD.c_fabricante_item[0].focus();">
<center>


<!--  CADASTRO DO PRODUTO COMPOSTO -->

<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="center" valign="bottom"><span class="PEDIDO">E-Commerce: Cadastro de Novo Produto Composto</span></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="ECProdutoCompostoNovoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='descricao_fornecida' value=''>

<!-- ************   CÓDIGO / DESCRIÇÃO   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td class="MD" width="15%" align="left"><p class="R">FABRICANTE</p><p class="C"><input id="fabricante_selecionado" name="fabricante_selecionado" class="TA" value="<%=fabricante_selecionado%>" readonly size="6" style="text-align:center; color=#0000ff"></p></td>
		<td class="MD" width="15%" align="left"><p class="R">PRODUTO</p><p class="C"><input id="produto_selecionado" name="produto_selecionado" class="TA" value="<%=produto_selecionado%>" readonly size="10" style="text-align:center; color=#0000ff"></p></td>
		<td width="70%" align="left"><p class="R">DESCRIÇÃO</p><p class="C"><input id="c_produto_composto_descricao" name="c_produto_composto_descricao" class="TA" type="text" maxlength="40" size="60" value="<%=descricao_fornecida%>" readonly tabindex=-1></p></td>
	</tr>
</table>

<br><br><p class="F" style="margin-bottom:8px;">Composição de 1 unidade do produto composto</p>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0">
	<tr bgcolor="#FFFFFF">
	<td class="MB" align="left"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left"><span class="PLTe">Produto</span></td>
	<td class="MB" align="right"><span class="PLTd">Qtde</span></td>
	<td class="MB" align="left"><span class="PLTe">Descrição</span></td>
	</tr>
<% for i=Lbound(v_item) to Ubound(v_item) 
		with v_item(i) %>
	
	<tr>
	<td class="MDBE" align="left"><input name="c_fabricante_item" id=<%=cstr(i)%> class="PLLe" maxlength="4" style="width:30px;" 
		value="<%=Trim("" & .fabricante_item)%>" readonly tabindex=-1></td>
	<td class="MDB" align="left"><input name="c_produto_item" id=<%=cstr(i)%> class="PLLe" maxlength="8" style="width:60px;" 
		value="<%=Trim("" & .produto_item)%>" readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_qtde_item" id=<%=cstr(i)%> class="PLLd" maxlength="4" style="width:30px;" 
		value="<%=formata_inteiro(.qtde)%>" readonly tabindex=-1></td>
	<td class="MDB" align="left"><input name="c_descricao_item" id=<%=cstr(i)%> class="PLLe" maxlength="40" style="width:277px;" 
		value="<%=Trim("" & .descricao)%>" readonly tabindex=-1></td>
	</tr>
<%		end with
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
	<td align="left"><a href="javascript:history.back()" title="cancela as alterações no cadastro">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" href="javascript:AtualizaProdutoComposto(fCAD)" title="grava o novo cadastro">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
<% end if %>

</html>


<%

'	FECHA CONEXAO COM O BANCO DE DADOS
	rs.Close
	set rs = nothing
	
	cn.Close
	set cn = nothing

%>