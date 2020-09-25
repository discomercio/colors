<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  ProdutoLista.asp
'     =====================================
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


' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
'	OBTEM USUÁRIO
	dim usuario
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim ordenacao_selecionada
	ordenacao_selecionada=Trim(request("ord"))




' ________________________________
' E X E C U T A _ C O N S U L T A
'
Sub executa_consulta
dim strSql, s, i, x, cab, s_link_produto, s_link_produto_placeholder
dim r
dim intLargFabricante, intLargProduto, intLargDescricao, intLargDescricaoHtml

	intLargFabricante = 35
	intLargProduto = 80
	intLargDescricao = 300
	intLargDescricaoHtml = 300
	
  ' CABEÇALHO
	cab="<table class='Q' cellspacing=0>" & chr(13)
	cab=cab & "<tr style='background:azure;' nowrap>"
	cab=cab & "<td align='left' style='width:" & Cstr(intLargFabricante) & "px;border-right: 1px solid;border-bottom:1px solid'><span class='R' style='cursor: hand;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='ECProdutoCompostoLista.asp?ord=1" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">FABR</span></td>"
	cab=cab & "<td align='left' style='width:" & Cstr(intLargProduto) & "px;border-right: 1px solid;border-bottom:1px solid'><span class='R' style='cursor: hand;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='ECProdutoCompostoLista.asp?ord=2" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">PRODUTO</span></TD>"
	cab=cab & "<td align='left' style='width:" & Cstr(intLargDescricao) & "px;border-right: 1px solid;border-bottom:1px solid'><span class='R' style='cursor: hand;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='ECProdutoCompostoLista.asp?ord=3" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">DESCRIÇÃO</span></TD>"
	cab=cab & "<td align='left' style='width:" & Cstr(intLargDescricaoHtml) & "px;border-bottom:1px solid'><span class='R'>DESCRIÇÃO (HTML)</span></td>"
	cab=cab & "</tr>" & chr(13)

	strSql= "SELECT " & _
				"*" & _
			" FROM t_PRODUTO" & _
			" ORDER BY "

	select case ordenacao_selecionada
		case "1": strSql = strSql & "fabricante, LEN(produto), produto"
		case "2": strSql = strSql & "LEN(produto), produto, fabricante"
		case "3": strSql = strSql & "descricao, LEN(produto), produto, fabricante"
		case else: strSql = strSql & "fabricante, LEN(produto), produto"
		end select

  ' EXECUTA CONSULTA
	x=cab
	i=0
	
	set r = cn.Execute( strSql )

	do while Not r.Eof 
	  ' CONTAGEM
		i = i + 1

	  ' ALTERNÂNCIA NAS CORES DAS LINHAS
		if (i AND 1)=0 then
			x=x & "<tr nowrap style='background: #FFF0E0' onmouseover='realca_cor_mouse_over(this);' onmouseout='realca_cor_mouse_out(this);'>"
		else
			x=x & "<tr nowrap onmouseover='realca_cor_mouse_over(this);' onmouseout='realca_cor_mouse_out(this);'>"
			end if

		s_link_produto_placeholder = "___LINK_PRODUTO_PLACEHOLDER___"
		s_link_produto = "<a href='javascript:fOPConcluir(" & _
						chr(34) & r("fabricante") & chr(34) & _
						"," & chr(34) & r("produto") & chr(34) & _
						")' title='clique para consultar o cadastro deste produto'>" & _
						s_link_produto_placeholder & _
						"</a>"

	 '> FABRICANTE
		x=x & " <td class='MDB' align='left' valign='top'><span class='C'>" & _
			replace(s_link_produto, s_link_produto_placeholder, r("fabricante")) & _
			"</span></td>"

	 '> PRODUTO
		x=x & " <td class='MDB' align='left' valign='top' nowrap><span class='C' nowrap>" & _
			replace(s_link_produto, s_link_produto_placeholder, r("produto")) & _
			"</span></td>"

	 '> DESCRIÇÃO
		s=Trim("" & r("descricao"))
		if s="" then s="&nbsp;"
		x=x & " <td class='MDB' align='left' valign='top'><span class='C'>" & _
			replace(s_link_produto, s_link_produto_placeholder, s) & _
			"</span></td>"

	 '> DESCRIÇÃO (HTML)
		s=Trim("" & r("descricao_html"))
		if s="" then s="&nbsp;"
		x=x & " <td class='MB' align='left' valign='top'><span class='C'>" & _
			replace(s_link_produto, s_link_produto_placeholder, s) & _
			"</span></td>"

		x=x & "</tr>" & chr(13)

		if (i mod 100) = 0 then
			Response.Write x
			x = ""
			end if

		r.MoveNext
		loop


  ' MOSTRA TOTAL DE PRODUTOS
	x=x & "<tr nowrap style='background: #FFFFDD'><td align='right' colspan='4' nowrap><span class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & cstr(i) & "&nbsp;&nbsp;produtos" & "</span></td></tr>"

  ' FECHA TABELA
	x=x & "</table>"
	

	Response.write x

	r.close
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

<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<script language="JavaScript">
var backgroundColorAnterior = "___UNDEFINED___";
window.status='Aguarde, executando consulta ...';

function realca_cor_mouse_over(c) {
    backgroundColorAnterior = c.style.backgroundColor;
    c.style.backgroundColor = 'palegreen';
}

function realca_cor_mouse_out(c) {
    if (backgroundColorAnterior != "___UNDEFINED___") c.style.backgroundColor = backgroundColorAnterior;
}

function fOPConcluir(s_fabricante, s_produto){
	fOP.fabricante_selecionado.value=s_fabricante;
	fOP.produto_selecionado.value=s_produto;
	window.status = "Aguarde ...";
	fOP.submit(); 
}
</script>

<link href="../global/e.css" rel="stylesheet" type="text/css">
<link href="../global/eprinter.css" rel="stylesheet" type="text/css" media="print">



<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Relação de Produtos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  RELAÇÃO DE PRODUTOS COMPOSTOS  -->
<br>
<center>
<form method="post" action="ProdutoEdita.asp" id="fOP" name="fOP">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='fabricante_selecionado' value=''>
<input type="hidden" name='produto_selecionado' value=''>
<input type="hidden" name='operacao_selecionada' value='<%=OP_CONSULTA%>'>
<% executa_consulta %>
</form>

<br>

<p class="TracoBottom"></p>

<table class="notPrint" cellspacing="0">
<tr>
	<td align="center"><a href="ProdutoMenu.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>

</body>
</html>


<%

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing

%>