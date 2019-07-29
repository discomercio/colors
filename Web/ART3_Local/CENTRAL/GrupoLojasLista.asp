<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  G R U P O L O J A S L I S T A . A S P
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
const w_lojas = 250
const w_descricao = 220
dim consulta, s, n_reg, x, cab, s_lojas
dim r, sx

  ' CABEÇALHO
	cab =	"<TABLE class='Q' cellSpacing=0>" & chr(13) & _
			"	<TR style='background: #FFF0E0'>" & chr(13) & _
			"		<TD align='right' class='MD MB' style='width:50px;'><P class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='GrupoLojasLista.asp?ord=1" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">GRUPO&nbsp;</P></TD>" & chr(13) & _
			"		<TD class='MD MB' style='width:" & Cstr(w_descricao) & "px;'><P class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='GrupoLojasLista.asp?ord=2" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">&nbsp;DESCRIÇÃO</P></TD>" & chr(13) & _
			"		<TD class='MB' style='width:" & Cstr(w_lojas) & "px;'><P class='R'>&nbsp;LOJAS</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)

	consulta= "SELECT * FROM t_LOJA_GRUPO ORDER BY "
	select case ordenacao_selecionada
		case "1": consulta = consulta & "CONVERT(smallint,grupo)"
		case "2": consulta = consulta & "descricao, CONVERT(smallint,grupo)"
		case else: consulta = consulta & "CONVERT(smallint,grupo)"
		end select

  ' EXECUTA CONSULTA
	x = cab
	n_reg = 0
	
	set r = cn.Execute( consulta )

	while not r.eof 
	  ' CONTAGEM
		n_reg = n_reg + 1

	  ' ALTERNÂNCIA NAS CORES DAS LINHAS
		if (n_reg AND 1)=0 then
			x = x & "	<TR style='background: #FFF0E0'>" & chr(13)
		else
			x = x & "	<TR>" & chr(13)
			end if

	 '> Nº DO GRUPO DE LOJAS
		s = Trim("" & r("grupo"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MDB' valign='top'><P class='Cd'>&nbsp;" & _
				"<a href='javascript:fOPConcluir(" & chr(34) & r("grupo") & chr(34) & _
				")' title='clique para consultar o cadastro deste grupo de lojas'>" & _
				s & "</a></P></TD>" & chr(13)

 	 '> DESCRIÇÃO
		s = Trim("" & r("descricao"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MDB' valign='top' style='width:" & Cstr(w_descricao) & "px;'>" & _
				"<P class='C' style='margin-left:6px;'>"  & _
				"<a href='javascript:fOPConcluir(" & chr(34) & r("grupo") & chr(34) & _
				")' title='clique para consultar o cadastro deste grupo de lojas'>" & _
				s & "</a></P></TD>" & chr(13)

 	 '> LOJAS
 		s = "SELECT loja FROM t_LOJA_GRUPO_ITEM WHERE (grupo = '" & Trim("" & r("grupo")) & "') ORDER BY CONVERT(smallint,loja)"
 		s_lojas = ""
 		set sx = cn.Execute(s)
 		do while Not sx.Eof
 		'	IMPORTANTE: É PRECISO UM ESPAÇO EM BRANCO (CHR(32)) P/ QUE SEJA FEITO O WORD-WRAP
 			if s_lojas <> "" then s_lojas = s_lojas & ",&nbsp;&nbsp;" & " "
 			s_lojas = s_lojas & Trim("" & sx("loja"))
 			sx.movenext
 			loop
 		sx.Close
		set sx=nothing
 		
		if s_lojas = "" then s_lojas = "&nbsp;"
		x = x & "		<TD class='MB' valign='top' style='width:" & Cstr(w_lojas) & "px;'>" & _
				"<P class='C' style='margin-left:6px;'>" & s_lojas & "</P></TD>" & chr(13)

		x = x & "	</TR>" & chr(13)

		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if

		r.MoveNext
		wend


  ' MOSTRA TOTAL DE GRUPOS DE LOJAS
	if n_reg = 1 then
		s = "grupo de lojas"
	else
		s = "grupos de lojas"
		end if
		
	x = x & "	<TR style='background: #FFFFDD'>" & chr(13) & _
			"		<TD COLSPAN='3'><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & _
					cstr(n_reg) & "&nbsp;&nbsp;" & s & "</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)

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

function fOPConcluir(s_grupo_lojas){
	fOP.grupo_lojas_selecionado.value=s_grupo_lojas;
	window.status = "Aguarde ...";
	fOP.submit(); 
}

</script>

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">


<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->  
<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="RIGHT" vAlign="BOTTOM" NOWRAP><span class="PEDIDO">Relação dos Grupos de Lojas Cadastrados</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  RELAÇÃO DOS GRUPOS DE LOJAS  -->
<br>
<center>
<form method="post" action="GrupoLojasEdita.asp" id="fOP" name="fOP">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<INPUT type=hidden name='grupo_lojas_selecionado' id="grupo_lojas_selecionado" value=''>
<INPUT type=hidden name='operacao_selecionada' id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<% executa_consulta %>
</form>

<br>

<p class="TracoBottom"></p>

<table class="notPrint" cellSpacing="0">
<tr>
	<td align="CENTER"><a href="GrupoLojas.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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