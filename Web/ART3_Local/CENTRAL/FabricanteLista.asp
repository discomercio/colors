<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  F A B R I C A N T E L I S T A . A S P
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
'
'
'	REVISADO P/ IE10


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
	Dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim ordenacao_selecionada
	ordenacao_selecionada=Trim(request("ord"))




' ________________________________
' E X E C U T A _ C O N S U L T A
'
Sub executa_consulta
dim consulta, s, i, x, cab
dim r

  ' CABEÇALHO
	cab="<table class='Q' cellspacing=0>" & chr(13)
	cab=cab & "<tr style='background: #FFF0E0' nowrap>"
	cab=cab & "<td width='35' align='right' nowrap class='MD MB'><P class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='fabricantelista.asp?ord=1" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Nº&nbsp;</P></TD>"
	cab=cab & "<td width='200' nowrap class='MD MB'><P class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='fabricantelista.asp?ord=2" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">&nbsp;NOME (apelido)</P></TD>"
	cab=cab & "<td width='250' class='MD MB'><P class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='fabricantelista.asp?ord=3" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">&nbsp;RAZÃO SOCIAL</P></TD>"
	cab=cab & "<td width='80' nowrap class='MB'><P class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='fabricantelista.asp?ord=4" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">&nbsp;CNPJ</P></TD>"
	cab=cab & "</tr>" & chr(13)

	consulta= "SELECT * FROM t_FABRICANTE ORDER BY "
	select case ordenacao_selecionada
		case "1": consulta = consulta & "fabricante"
		case "2": consulta = consulta & "nome, fabricante"
		case "3": consulta = consulta & "razao_social, fabricante"
		case "4": consulta = consulta & "cnpj, fabricante"
		case else: consulta = consulta & "fabricante"
		end select

  ' EXECUTA CONSULTA
	x=cab
	i=0
	
	set r = cn.Execute( consulta )

	while not r.eof 
	  ' CONTAGEM
		i = i + 1

	  ' ALTERNÂNCIA NAS CORES DAS LINHAS
		if (i AND 1)=0 then
			x=x & "<TR NOWRAP style='background: #FFF0E0'>"
		else
			x=x & "<TR NOWRAP >"
			end if

	 '> Nº
		x=x & " <TD class='MDB' valign='top'><P class='Cd'>&nbsp;"
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("fabricante") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste fabricante'>"
		x=x & r("fabricante") & "</a></P></TD>"

	 '> NOME (APELIDO)
		x=x & " <TD class='MDB' valign='top' NOWRAP><P class='C' NOWRAP>&nbsp;" 
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("fabricante") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste fabricante'>"
		x=x & r("nome") & "</a></P></TD>"

	 '> RAZÃO SOCIAL
		s=Trim("" & r("razao_social"))
		x=x & " <TD class='MDB' valign='top'><P class='C'>&nbsp;" & s & "</P></TD>"

	 '> CNPJ
		x=x & " <TD class='MB' valign='top' NOWRAP><P class='C'>&nbsp;"
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("fabricante") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste fabricante'>"
		x=x & cnpj_cpf_formata(r("cnpj")) & "</a></P></TD>"

		x=x & "</TR>" & chr(13)

		if (i mod 100) = 0 then
			Response.Write x
			x = ""
			end if

		r.MoveNext
		wend


  ' MOSTRA TOTAL DE FABRICANTES
	x=x & "<TR NOWRAP style='background: #FFFFDD'><TD COLSPAN='4' NOWRAP><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & cstr(i) & "&nbsp;&nbsp;fabricantes" & "</P></TD></TR>"

  ' FECHA TABELA
	x=x & "</TABLE>"
	

	Response.write x

	r.close
	set r=nothing

End sub

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
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando consulta ...';

function fOPConcluir(s_fabricante){
	fOP.fabricante_selecionado.value=s_fabricante;
	window.status = "Aguarde ...";
	fOP.submit(); 
}

</script>

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">



<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Relação de Fabricantes Cadastrados</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  RELAÇÃO DE FABRICANTES  -->
<br>
<center>
<form method="post" action="fabricanteedita.asp" id="fOP" name="fOP">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='fabricante_selecionado' id="fabricante_selecionado" value=''>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<% executa_consulta %>
</form>

<br>

<p class="TracoBottom"></p>

<table class="notPrint" cellspacing="0">
<tr>
	<td align="center"><a href="fabricante.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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