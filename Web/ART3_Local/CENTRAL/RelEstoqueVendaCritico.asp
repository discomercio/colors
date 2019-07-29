<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L E S T O Q U E V E N D A C R I T I C O . A S P
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_ESTOQUE_VENDA_CRITICO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta, s
	alerta = ""



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA ESTOQUE VENDA CRITICO
'
sub consulta_estoque_venda_critico
dim r
dim s_sql, x, cab
dim n, n_reg, n_saldo

'	O FLAG "EXCLUIDO_STATUS" INDICA SE O PRODUTO ESTÁ EXCLUÍDO LOGICAMENTE DO SISTEMA!!
'	A TABELA BÁSICA DE PRODUTOS MANTÉM INFORMAÇÕES DE PRODUTOS EXCLUÍDOS LOGICAMENTE 
'	P/ MANTER A REFERÊNCIA COM OUTRAS TABELAS QUE NECESSITEM DE DADOS COMO DESCRIÇÃO, ETC.
	s_sql = "SELECT t_PRODUTO.fabricante, t_PRODUTO.produto, descricao, descricao_html, estoque_critico, Sum(qtde-qtde_utilizada) AS saldo" & _
			" FROM t_PRODUTO LEFT JOIN t_ESTOQUE_ITEM ON ((t_PRODUTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_PRODUTO.produto=t_ESTOQUE_ITEM.produto))" & _
			" WHERE (excluido_status=0)" & _
			" AND (estoque_critico > 0)" & _
			" GROUP BY t_PRODUTO.fabricante, t_PRODUTO.produto, descricao, descricao_html, estoque_critico" & _
			" HAVING (Sum(qtde-qtde_utilizada) = 0)" & _
				" OR (Sum(qtde-qtde_utilizada) IS NULL)" & _
				" OR (Sum(qtde-qtde_utilizada) < estoque_critico)" & _
			" ORDER BY t_PRODUTO.fabricante, t_PRODUTO.produto, descricao, descricao_html, estoque_critico"

  ' CABEÇALHO
	cab = "<TABLE class='Q' cellSpacing=0>" & chr(13) & _
		  "<TR style='background:azure' NOWRAP>" & _
		  "<TD width='36' valign='bottom' NOWRAP class='MD MB'><P class='R'>FABR</P></TD>" & _
		  "<TD width='60' valign='bottom' NOWRAP class='MD MB'><P class='R'>PRODUTO</P></TD>" & _
		  "<TD width='417' valign='bottom' NOWRAP class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & _
		  "<TD width='60' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE CRÍTICA</P></TD>" & _
		  "<TD width='60' valign='bottom' NOWRAP class='MB'><P class='Rd' style='font-weight:bold;'>QTDE DISPON</P></TD>" & _
		  "</TR>" & chr(13)
	
	x = cab
	n_reg = 0
	n_saldo = 0

	set r = cn.execute(s_sql)
	do while Not r.Eof
	  ' CONTAGEM
		n_reg = n_reg + 1

		x = x & "<TR NOWRAP>"

	 '> FABRICANTE
		x = x & "	<TD class='MDB' valign='bottom'><P class='C'>&nbsp;" & Trim("" & r("fabricante")) & "</P></TD>"

	 '> PRODUTO
		x = x & "	<TD class='MDB' valign='bottom'><P class='C'>&nbsp;" & Trim("" & r("produto")) & "</P></TD>"

	 '> DESCRIÇÃO
		x = x & "	<TD class='MDB' valign='bottom' NOWRAP><P class='C' NOWRAP>&nbsp;" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>"

	 '> QTDE CRÍTICA
		x = x & "	<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(r("estoque_critico")) & "</P></TD>"

	 '> QTDE DISPONÍVEL
		if IsNumeric(r("saldo")) then n=r("saldo") else n=0
		x = x & "	<TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(n) & "</P></TD>"

		n_saldo = n_saldo + n
		
		x = x & "</TR>" & chr(13)
			
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA TOTAL
	if n_reg <> 0 then 
		x = x & "<TR NOWRAP style='background: #FFFFDD'><TD COLSPAN='5' NOWRAP><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & formata_inteiro(n_saldo) & "</P></TD></TR>"
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab & _
			"<TR NOWRAP>" & _
			"	<TD colspan='5'><P class='ALERTA'>&nbsp;NENHUM PRODUTO ESTÁ COM A QUANTIDADE EM NÍVEL CRÍTICO&nbsp;</P></TD>" & _
			"</TR>"
		end if

  ' FECHA TABELA
	x = x & "</TABLE>"
	
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
<body onload="window.status='Concluído';">

<center>

<form id="fESTOQ" name="fESTOQ" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque de Venda Crítico</span>
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

	
<!--  RELATÓRIO  -->
<% consulta_estoque_venda_critico %>

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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
