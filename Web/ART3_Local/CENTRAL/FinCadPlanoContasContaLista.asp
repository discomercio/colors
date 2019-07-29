<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  FinCadPlanoContasContaLista.asp
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
dim w_descricao, w_grupo
dim strTextoNaturezaDebito, strTextoNaturezaCredito
dim strTextoStatusInativo, strTextoStatusAtivo

	strTextoNaturezaDebito="<span style='color:" & finNaturezaCor(COD_FIN_NATUREZA__DEBITO) & ";'>" & finNaturezaDescricao(COD_FIN_NATUREZA__DEBITO) & "</span>"
	strTextoNaturezaCredito="<span style='color:" & finNaturezaCor(COD_FIN_NATUREZA__CREDITO) & ";'>" & finNaturezaDescricao(COD_FIN_NATUREZA__CREDITO) & "</span>"
	
	strTextoStatusInativo = "<span style='color:" & finStAtivoCor(COD_FIN_ST_ATIVO__INATIVO) & ";'>" & finStAtivoDescricao(COD_FIN_ST_ATIVO__INATIVO) & "</span>"
	strTextoStatusAtivo = "<span style='color:" & finStAtivoCor(COD_FIN_ST_ATIVO__ATIVO) & ";'>" & finStAtivoDescricao(COD_FIN_ST_ATIVO__ATIVO) & "</span>"

	w_descricao=200
	w_grupo=200
	
  ' CABEÇALHO
	cab = _
			"<TABLE class='Q' cellSpacing=0>" & chr(13) & _
			"	<TR style='background:azure;' NOWRAP>" & chr(13) & _
			"		<TD width='35' align='center' NOWRAP class='MD MB'><P class='R' style='cursor:pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadPlanoContasContaLista.asp?ord=1" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Id</P></TD>" & chr(13) & _
			"		<TD width='" & w_descricao & "' NOWRAP class='MD MB'><P class='R' style='cursor:pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadPlanoContasContaLista.asp?ord=2" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">&nbsp;Descrição</P></TD>" & chr(13) & _
			"		<TD width='" & w_grupo & "' NOWRAP class='MD MB'><P class='R' style='cursor:pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadPlanoContasContaLista.asp?ord=3" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">&nbsp;Grupo</P></TD>" & chr(13) & _
			"		<TD width='70' NOWRAP class='MD MB'><P class='R' style='cursor:pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadPlanoContasContaLista.asp?ord=4" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">&nbsp;Natureza</P></TD>" & chr(13) & _
			"		<TD width='70' align='center' NOWRAP class='MB'><P class='R' style='cursor:pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadPlanoContasContaLista.asp?ord=5" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Status</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)

	consulta = _
			"SELECT " & _
				"c.*, g.descricao AS descricao_grupo" & _
			" FROM t_FIN_PLANO_CONTAS_CONTA c" & _
				" INNER JOIN t_FIN_PLANO_CONTAS_GRUPO g" & _
					" ON c.id_plano_contas_grupo=g.id" & _
			" ORDER BY "
	select case ordenacao_selecionada
		case "1": consulta = consulta & "c.id"
		case "2": consulta = consulta & "c.descricao, c.id"
		case "3": consulta = consulta & "c.id_plano_contas_grupo, c.id"
		case "4": consulta = consulta & "c.natureza, c.id"
		case "5": consulta = consulta & "c.st_ativo, c.id"
		case else: consulta = consulta & "c.id"
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
			x=x & "	<TR NOWRAP style='background:#FFFFF0;'>"
		else
			x=x & "	<TR NOWRAP>"
			end if

	 '> ID
		x = x & _
			"		<TD class='MDB' align='center' valign='top'><P class='Cc'>&nbsp;" & _
			"<a href='javascript:fOPConcluir(" & chr(34) & r("id") & chr(34) & _
			"," & chr(34) & r("natureza") & chr(34) & _
			")' title='clique para consultar o cadastro deste registro'>" & _
			normaliza_codigo(r("id"),TAM_PLANO_CONTAS__CONTA) & "</a></P></TD>" & chr(13)

 	 '> DESCRIÇÃO
		x = x & _
			"		<TD class='MDB' width='" & w_descricao & "' valign='top'><P class='C' NOWRAP>&nbsp;" & _
			"<a href='javascript:fOPConcluir(" & chr(34) & r("id") & chr(34) & _
			"," & chr(34) & r("natureza") & chr(34) & _
			")' title='clique para consultar o cadastro deste registro'>" & _
			r("descricao") & "</a></P></TD>" & chr(13)

 	 '> GRUPO DE CONTAS
		x = x & _
			"		<TD class='MDB' width='" & w_grupo & "' valign='top'><P class='C' NOWRAP>&nbsp;" & _
			"<a href='javascript:fOPConcluir(" & chr(34) & r("id") & chr(34) & _
			"," & chr(34) & r("natureza") & chr(34) & _
			")' title='clique para consultar o cadastro deste registro'>" & _
			normaliza_codigo(r("id_plano_contas_grupo"),TAM_PLANO_CONTAS__GRUPO) & " - " & r("descricao_grupo") & "</a></P></TD>" & chr(13)

 	 '> NATUREZA
 		if Cstr(r("natureza"))=Cstr(COD_FIN_NATUREZA__DEBITO) then 
 			s=strTextoNaturezaDebito
 		else 
 			s=strTextoNaturezaCredito
 			end if
		x = x & _
			"		<TD class='MDB' align='center' valign='top' NOWRAP><P class='Cc'>&nbsp;" & s & "</P></TD>" & chr(13)

 	 '> STATUS
 		if Cstr(r("st_ativo"))=Cstr(COD_FIN_ST_ATIVO__INATIVO) then 
 			s=strTextoStatusInativo
 		else 
 			s=strTextoStatusAtivo
 			end if
		x = x & _
			"		<TD class='MB' align='center' valign='top' NOWRAP><P class='Cc'>&nbsp;" & s & "</P></TD>" & chr(13)

		x = x & _
			"	</TR>" & chr(13)

		if (i mod 100) = 0 then
			Response.Write x
			x = ""
			end if

		r.MoveNext
		wend


  ' MOSTRA TOTAL DE REGISTROS
	x = x & _
		"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
		"		<TD COLSPAN='5' NOWRAP><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & cstr(i) & "&nbsp;&nbsp;registros" & "</P></TD>" & chr(13) & _
		"	</TR>" & chr(13)

  ' FECHA TABELA
	x = x & _
		"</TABLE>" & chr(13)
	
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

function fOPConcluir(s_id, s_natureza){
	window.status = "Aguarde ...";
	fOP.id_selecionado.value=s_id;
	fOP.rb_natureza.value=s_natureza;
	fOP.submit(); 
}

</script>

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">


<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Plano de Contas: Relação de Contas Cadastradas</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  RELAÇÃO DE REGISTROS CADASTRADOS  -->
<br>
<center>
<form method="post" action="FinCadPlanoContasContaEdita.asp" id="fOP" name="fOP">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='id_selecionado' id="id_selecionado" value=''>
<input type="hidden" name='rb_natureza' id="rb_natureza" value=''>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<% executa_consulta %>
</form>

<br>

<p class="TracoBottom"></p>

<table class="notPrint" cellspacing="0">
<tr>
	<td align="center"><a href="FinCadPlanoContasContaMenu.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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