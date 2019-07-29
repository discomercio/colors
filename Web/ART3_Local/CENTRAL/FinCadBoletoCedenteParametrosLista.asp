<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =========================================
'	  FinCadBoletoCedenteParametrosLista.asp
'     =========================================
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
dim w_descricao
dim strTextoStatusInativo, strTextoStatusAtivo

	strTextoStatusInativo = "<span style='color:" & finStAtivoCor(COD_FIN_ST_ATIVO__INATIVO) & ";'>" & finStAtivoDescricao(COD_FIN_ST_ATIVO__INATIVO) & "</span>"
	strTextoStatusAtivo = "<span style='color:" & finStAtivoCor(COD_FIN_ST_ATIVO__ATIVO) & ";'>" & finStAtivoDescricao(COD_FIN_ST_ATIVO__ATIVO) & "</span>"
	
	w_descricao = 200
	
  ' CABEÇALHO
	cab = _
			"<TABLE class='Q' cellSpacing=0>" & chr(13) & _
			"	<TR style='background:azure;' NOWRAP>" & chr(13) & _
			"		<TD width='30' align='center' valign='bottom' NOWRAP class='MD MB'><P class='R' style='cursor:pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadBoletoCedenteParametrosLista.asp?ord=1" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Id</P></TD>" & chr(13) & _
			"		<TD width='" & w_descricao & "' valign='bottom' NOWRAP class='MD MB'><P class='R' style='cursor:pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadBoletoCedenteParametrosLista.asp?ord=2" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Nome da Empresa</P></TD>" & chr(13) & _
			"		<TD width='50' align='center' valign='bottom' NOWRAP class='MD MB'><P class='Rc' style='cursor:pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadBoletoCedenteParametrosLista.asp?ord=3" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Banco</P></TD>" & chr(13) & _
			"		<TD width='60' align='center' valign='bottom'  NOWRAP class='MD MB'><P class='Rc' style='cursor:pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadBoletoCedenteParametrosLista.asp?ord=4" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Agência</P></TD>" & chr(13) & _
			"		<TD width='70' align='center' valign='bottom'  NOWRAP class='MD MB'><P class='Rc' style='cursor:pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadBoletoCedenteParametrosLista.asp?ord=5" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Nº Conta</P></TD>" & chr(13) & _
			"		<TD width='65' align='right' valign='bottom'  NOWRAP class='MD MB'><P class='Rd' style='cursor:pointer;font-weight:bold;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadBoletoCedenteParametrosLista.asp?ord=6" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Nº Seq Remessa</P></TD>" & chr(13) & _
			"		<TD width='75' align='right' valign='bottom'  NOWRAP class='MD MB'><P class='Rd' style='cursor:pointer;font-weight:bold;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadBoletoCedenteParametrosLista.asp?ord=7" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Juros Mora (%)</P></TD>" & chr(13) & _
			"		<TD width='50' align='center' valign='bottom' NOWRAP class='MB'><P class='R' style='cursor:pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='FinCadBoletoCedenteParametrosLista.asp?ord=8" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Status</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)

	consulta = "SELECT * FROM t_FIN_BOLETO_CEDENTE ORDER BY "
	select case ordenacao_selecionada
		case "1": consulta = consulta & "id"
		case "2": consulta = consulta & "nome_empresa, id"
		case "3": consulta = consulta & "num_banco, agencia, conta, id"
		case "4": consulta = consulta & "agencia, conta, id"
		case "5": consulta = consulta & "conta, id"
		case "6": consulta = consulta & "nsu_arq_remessa, id"
		case "7": consulta = consulta & "juros_mora, id"
		case "8": consulta = consulta & "st_ativo, id"
		case else: consulta = consulta & "id"
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
			")' title='clique para consultar o cadastro deste registro'>" & _
			r("id") & "</a></P></TD>" & chr(13)

	 '> NOME DA EMPRESA
		x = x & _
			"		<TD class='MDB' width='" & w_descricao & "' valign='top'><P class='C' NOWRAP>&nbsp;" & _
			"<a href='javascript:fOPConcluir(" & chr(34) & r("id") & chr(34) & _
			")' title='clique para consultar o cadastro deste registro'>" & _
			r("nome_empresa") & "</a></P></TD>" & chr(13)

	 '> BANCO
		x = x & _
			"		<TD class='MDB' align='center' valign='top' NOWRAP><P class='Cc'>" & Trim("" & r("num_banco")) & "</p></td>" & chr(13)

	 '> AGÊNCIA
		x = x & _
			"		<TD class='MDB' align='center' valign='top' NOWRAP><P class='Cc'>" & Trim("" & r("agencia")) & "</p></td>" & chr(13)

	 '> CONTA
		x = x & _
			"		<TD class='MDB' align='center' valign='top' NOWRAP><P class='Cc'>" & Trim("" & r("conta")) & "</p></td>" & chr(13)

	 '> NSU ARQ REMESSA
		x = x & _
			"		<TD class='MDB' align='center' valign='top' NOWRAP><P class='Cc'>" & r("nsu_arq_remessa") & "</p></td>" & chr(13)

	 '> JUROS MORA
		x = x & _
			"		<TD class='MDB' align='right' valign='top' NOWRAP><P class='Cd'>" & formata_perc(r("juros_mora")) & "</p></td>" & chr(13)
			
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
		"		<TD COLSPAN='8' NOWRAP><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & cstr(i) & "&nbsp;&nbsp;registros" & "</P></TD>" & chr(13) & _
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

function fOPConcluir(s_id){
	window.status = "Aguarde ...";
	fOP.id_selecionado.value=s_id;
	fOP.submit(); 
}

</script>

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">


<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->  
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Boleto - Relação de Contas do Cedente</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  RELAÇÃO DE REGISTROS CADASTRADOS  -->
<br>
<center>
<form method="post" action="FinCadBoletoCedenteParametrosEdita.asp" id="fOP" name="fOP">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='id_selecionado' id="id_selecionado" value=''>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<% executa_consulta %>
</form>

<br>

<p class="TracoBottom"></p>

<table class="notPrint" cellspacing="0">
<tr>
	<td align="center"><a href="FinCadBoletoCedenteParametrosMenu.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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