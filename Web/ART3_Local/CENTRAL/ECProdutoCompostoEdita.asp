<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  ECProdutoCompostoEdita.asp
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
	
'	OBTEM O ID
	dim s, usuario, fabricante_selecionado, produto_selecionado, descricao_fornecida
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	PRODUTO COMPOSTO A EDITAR
	fabricante_selecionado = trim(request("fabricante_selecionado"))
	produto_selecionado = trim(request("produto_selecionado"))
	
	fabricante_selecionado=retorna_so_digitos(fabricante_selecionado)
	produto_selecionado=retorna_so_digitos(produto_selecionado)

	fabricante_selecionado=normaliza_codigo(fabricante_selecionado, TAM_MIN_FABRICANTE)
	produto_selecionado=normaliza_produto(produto_selecionado)
	
	if (fabricante_selecionado="") Or (fabricante_selecionado="000") then Response.Redirect("aviso.asp?id=" & ERR_FABRICANTE_NAO_ESPECIFICADO) 
	if (produto_selecionado="") Or (produto_selecionado="000000") then Response.Redirect("aviso.asp?id=" & ERR_EC_PRODUTO_COMPOSTO_NAO_ESPECIFICADO) 

	dim v_item
	dim strFabricante, strProduto, strQtde, strDescricao, blnExibeBotaoLimpaLinha
	
'	CONECTA COM O BANCO DE DADOS
	dim cn,rs,r,strSql
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim max_qtde_itens
	max_qtde_itens = obtem_parametro_ProdutoComposto_MaxQtdeItens

	strSql = "SELECT " & _
				"*" & _
			" FROM t_EC_PRODUTO_COMPOSTO" & _
			" WHERE" & _
				" (fabricante_composto = '" & fabricante_selecionado & "')" & _
				" AND (produto_composto='" & produto_selecionado & "')"
	set rs = cn.Execute(strSql)
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_EC_PRODUTO_COMPOSTO_NAO_CADASTRADO)

	descricao_fornecida = rs("descricao")

	dim i, n, msg_erro
	dim alerta
	alerta = ""
	
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
function LimparLinha(f, intIdx) {
	f.c_fabricante_item[intIdx].value="";
	f.c_produto_item[intIdx].value="";
	f.c_qtde_item[intIdx].value="";
	f.c_descricao_item[intIdx].value="";
}

function RemoveProdutoComposto( f ) {
var b;
	b=window.confirm('Confirma a exclusão deste produto composto?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.action="ECProdutoCompostoEditaConfirma.asp";
		f.submit();
		}
}

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

	dATUALIZA.style.visibility="hidden";
	f.action="ECProdutoCompostoEditaConsiste.asp";
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

<body onload="focus();">
<center>



<!--  CADASTRO DO PRODUTO COMPOSTO -->

<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">E-Commerce: Consulta/Edição de Produto Composto</span></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='operacao_selecionada' value=''>

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
<table class="Qx" cellspacing="0">
	<tr bgcolor="#FFFFFF">
	<td class="MB" align="left"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left"><span class="PLTe">Produto</span></td>
	<td class="MB" align="left"><span class="PLTd">Qtde</span></td>
	<td class="MB" align="left"><span class="PLTe">Descrição</span></td>
	<td align="left">&nbsp;</td>
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
				blnExibeBotaoLimpaLinha = True
				end with
		else
			strFabricante = ""
			strProduto = ""
			strQtde = ""
			strDescricao = ""
			blnExibeBotaoLimpaLinha = False
			end if
%>

	<tr>
	<td class="MDBE" align="left"><input name="c_fabricante_item" id=<%=cstr(i)%> class="PLLe" maxlength="3" style="width:30px;" onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(this.id!=1))) if (trim(this.value)=='') bATUALIZA.focus(); else fCAD.c_produto_item[parseInt(this.id)-1].focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"
		value="<%=strFabricante%>"></td>
	<td class="MDB" align="left"><input name="c_produto_item" id=<%=cstr(i)%> class="PLLe" maxlength="6" style="width:60px;" onkeypress="if (digitou_enter(true)) fCAD.c_qtde_item[parseInt(this.id)-1].focus(); filtra_produto();" onblur="this.value=normaliza_produto(this.value);"
		value="<%=strProduto%>"></td>
	<td class="MDB" align="right"><input name="c_qtde_item" id=<%=cstr(i)%> class="PLLd" maxlength="4" style="width:30px;" onkeypress="if (digitou_enter(true)) {if (parseInt(this.id)==fCAD.c_qtde_item.length) bATUALIZA.focus(); else fCAD.c_fabricante_item[parseInt(this.id)].focus();} filtra_numerico();"
		value="<%=strQtde%>"></td>
	<td class="MDB" align="left"><input name="c_descricao_item" id=<%=cstr(i)%> class="PLLe" maxlength="40" style="width:277px;"
		value="<%=strDescricao%>" readonly tabindex=-1></td>
	<td align="left">
	<% if blnExibeBotaoLimpaLinha then %>
		<a name="bLimparLinha" href="javascript:LimparLinha(fCAD,<%=Cstr(i-1)%>)" title="limpa o conteúdo desta linha"><img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
	<% else %>
		&nbsp;
	<% end if %>
	</td>
	</tr>

<%
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
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align='center'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveProdutoComposto(fCAD)'
		title='remove o produto composto cadastrado'><img src='../botao/remover.gif' width=176 height=55 border=0></a>
		</div>
	</td>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" href="javascript:AtualizaProdutoComposto(fCAD)" title="segue para a página de consistência da edição">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
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