<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  ECProdutoCompostoLista.asp
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
'			I N I C I A L I Z A     P � G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
'	OBTEM USU�RIO
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
dim strSql, s, i, x, cab, strComposto
dim r, r2
dim intLargFabricante, intLargProduto, intLargDescricao, intLargComposicao

	intLargFabricante = 35
	intLargProduto = 80
	intLargDescricao = 200
	intLargComposicao = 310
	
  ' CABE�ALHO
	cab="<table class='Q' cellspacing=0>" & chr(13)
	cab=cab & "<tr style='background: #FFF0E0' nowrap>"
	cab=cab & "<td align='left' style='width:" & Cstr(intLargFabricante) & ";border-right: 1px solid;border-bottom:1px solid'><span class='R' style='cursor: hand;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='ECProdutoCompostoLista.asp?ord=1" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">FABR</span></td>"
	cab=cab & "<td align='left' style='width:" & Cstr(intLargProduto) & ";border-right: 1px solid;border-bottom:1px solid'><span class='R' style='cursor: hand;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='ECProdutoCompostoLista.asp?ord=2" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">PRODUTO</span></TD>"
	cab=cab & "<td align='left' style='width:" & Cstr(intLargDescricao) & ";border-right: 1px solid;border-bottom:1px solid'><span class='R' style='cursor: hand;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='ECProdutoCompostoLista.asp?ord=3" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">DESCRI��O</span></TD>"
	cab=cab & "<td align='left' style='width:" & Cstr(intLargComposicao) & ";border-bottom:1px solid'><span class='R'>COMPOSI��O</span></td>"
	cab=cab & "</tr>" & chr(13)

	strSql= "SELECT " & _
				"*" & _
			" FROM t_EC_PRODUTO_COMPOSTO" & _
			" ORDER BY "

	select case ordenacao_selecionada
		case "1": strSql = strSql & "fabricante_composto, produto_composto"
		case "2": strSql = strSql & "produto_composto, fabricante_composto"
		case "3": strSql = strSql & "descricao, produto_composto, fabricante_composto"
		case else: strSql = strSql & "fabricante_composto, produto_composto"
		end select

  ' EXECUTA CONSULTA
	x=cab
	i=0
	
	set r = cn.Execute( strSql )

	do while Not r.Eof 
	  ' CONTAGEM
		i = i + 1

	  ' ALTERN�NCIA NAS CORES DAS LINHAS
		if (i AND 1)=0 then
			x=x & "<tr nowrap style='background: #FFF0E0'>"
		else
			x=x & "<tr nowrap>"
			end if

	
	 '> FABRICANTE
		x=x & " <td style='width:" & Cstr(intLargFabricante) & ";' class='MDB' align='left' valign='top'><span class='C'>"
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("fabricante_composto") & chr(34) & "," & chr(34) & r("produto_composto") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste produto composto'>"
		x=x & r("fabricante_composto")
		x=x & "</a>"
		x=x & "</span></td>"

	 '> PRODUTO
		x=x & " <td style='width:" & Cstr(intLargProduto) & ";' class='MDB' align='left' valign='top' nowrap><span class='C' nowrap>" 
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("fabricante_composto") & chr(34) & "," & chr(34) & r("produto_composto") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste produto composto'>"
		x=x & r("produto_composto")
		x=x & "</a>"
		x=x & "</span></td>"

	 '> DESCRI��O
		s=Trim("" & r("descricao"))
		if s="" then s="&nbsp;"
		x=x & " <td style='width:" & Cstr(intLargDescricao) & ";' class='MDB' align='left' valign='top'><span class='C'>"
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("fabricante_composto") & chr(34) & "," & chr(34) & r("produto_composto") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste produto composto'>"
		x=x & s
		x=x & "</a>"
		x=x & "</span></td>"

	 '> COMPOSI��O
		strSql = "SELECT " & _
					"*" & _
				" FROM t_EC_PRODUTO_COMPOSTO_ITEM a" & _
					" LEFT JOIN t_PRODUTO b ON (a.fabricante_item=b.fabricante) AND (a.produto_item=b.produto)" & _
				" WHERE" & _
					" (fabricante_composto = '" & Trim("" & r("fabricante_composto")) & "')" & _
					" AND (produto_composto = '" & Trim("" & r("produto_composto")) & "')" & _
				" ORDER BY" & _
					" sequencia"
		set r2 = cn.Execute(strSql)
		strComposto = ""
		do while Not r2.Eof
			if strComposto <> "" then strComposto = strComposto & "<br />"
			strComposto = strComposto & "<span class='C'>" & formata_inteiro(r2("qtde")) & " x " & Trim("" & r2("produto_item")) & " (" & Trim("" & r2("fabricante_item")) & ") - " & Trim("" & r2("descricao")) & "</span>"
			r2.MoveNext
			loop
		
		x=x & " <td style='width:" & Cstr(intLargComposicao) & ";' class='MB' align='left' valign='top' nowrap>"
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("fabricante_composto") & chr(34) & "," & chr(34) & r("produto_composto") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste produto composto'>"
		x=x & strComposto
		x=x & "</a>"
		x=x & "</td>"

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
window.status='Aguarde, executando consulta ...';

function fOPConcluir(s_fabricante, s_produto){
	fOP.fabricante_selecionado.value=s_fabricante;
	fOP.produto_selecionado.value=s_produto;
	window.status = "Aguarde ...";
	fOP.submit(); 
}
</script>

<link href="../global/e.css" rel="stylesheet" type="text/css">
<link href="../global/eprinter.css" rel="stylesheet" type="text/css" media="print">



<body onload="window.status='Conclu�do';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A � � O  -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">E-Commerce: Rela��o de Produtos Compostos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  RELA��O DE PRODUTOS COMPOSTOS  -->
<br>
<center>
<form method="post" action="ECProdutoCompostoEdita.asp" id="fOP" name="fOP">
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
	<td align="center"><a href="ECProdutoCompostoMenu.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a p�gina anterior">
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