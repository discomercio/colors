<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================
'	  U S U A R I O L I S T A . A S P
'     ===============================
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

	dim opcao_consulta
	opcao_consulta=UCase(Trim(request("op")))



' ________________________________
' E X E C U T A _ C O N S U L T A
'
Sub executa_consulta
dim consulta, s_where, s, i, x, cab, s_op
dim r

	s_op = ""
	if opcao_consulta <> "" then s_op = "&op=" & opcao_consulta
	
  ' CABEÇALHO
	cab="<table class='Q' cellspacing=0>" & chr(13)
	cab=cab & "<tr style='background: #FFF0E0'>"
	cab=cab & "<td width='100' nowrap class='MD MB' align='left' valign='bottom'><p class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='usuariolista.asp?ord=1" & s_op & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Identificação</p></td>"
	cab=cab & "<td width='150' class='MD MB' align='left' valign='bottom'><p class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='usuariolista.asp?ord=2" & s_op & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Nome</p></td>"
	cab=cab & "<td width='200' class='MD MB' align='left' valign='bottom'><p class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='usuariolista.asp?ord=3" & s_op & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Perfil</p></td>"
	cab=cab & "<td width='35' nowrap class='MD MB' align='right' valign='bottom'><p class='Rd' style='font-weight:bold; cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='usuariolista.asp?ord=4" & s_op & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Vend<br />Loja</p></td>"
	cab=cab & "<td width='20' nowrap class='MD MB' align='center' valign='bottom'><p class='Rd' style='font-weight:bold; cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='usuariolista.asp?ord=5" & s_op & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Vend<br />Ext</p></td>"
	cab=cab & "<td width='20' nowrap class='MD MB' align='center' valign='bottom'><p class='Rc' style='font-weight:bold; cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='usuariolista.asp?ord=5" & s_op & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">CD</p></td>"
	cab=cab & "<td width='20' nowrap class='MD MB' align='center' valign='bottom'><p class='Rc' style='font-weight:bold; cursor: pointer;'>Telefone</p></td>"
	cab=cab & "<td width='50' nowrap class='MB' align='left' valign='bottom'><p class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='usuariolista.asp?ord=6" & s_op & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Acesso</p></td>"
	cab=cab & "</tr>" & chr(13)

	consulta = "SELECT t_USUARIO.*, " & _
				SCHEMA_BD & ".ConcatenaPerfisDoUsuario(t_USUARIO.usuario, '<br />') AS perfil," & _
				SCHEMA_BD & ".ConcatenaLojasDoUsuario(t_USUARIO.usuario, '<br />') AS lista_lojas," & _
				SCHEMA_BD & ".ConcatenaSiglaWmsCdDoUsuario(t_USUARIO.usuario, '<br />') AS lista_wms_cd" & _
				" FROM t_USUARIO"

	s_where = ""
	if opcao_consulta = "A" then
		s_where = "bloqueado = 0"
	elseif opcao_consulta = "I" then
		s_where = "bloqueado = 1"
		end if
		
	if s_where <> "" then s_where = " WHERE " & s_where
	consulta = consulta & s_where
	
	consulta = consulta & " ORDER BY "
	select case ordenacao_selecionada
		case "1": consulta = consulta & "usuario"
		case "2": consulta = consulta & "nome, usuario"
		case "3": consulta = consulta & "perfil, usuario"
		case "4": consulta = consulta & "lista_lojas, usuario"
		case "5": consulta = consulta & "vendedor_externo, usuario"
		case "6": consulta = consulta & "bloqueado, usuario"
		case else: consulta = consulta & "usuario"
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
			x=x & "<tr style='background: #FFF0E0'>"
		else
			x=x & "<tr>"
			end if

	 '> APELIDO
		x=x & " <td class='MDB' align='left' valign='top'><p class='C'>"
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("usuario") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste usuário'>"
		x=x & r("usuario") & "</a></p></td>"

	 '> NOME
		x=x & " <td class='MDB' style='width:150px;' align='left' valign='top'><p class='C'>" 
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("usuario") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste usuário'>"
		x=x & r("nome_iniciais_em_maiusculas") & "</a></p></td>"

	 '> PERFIL
		s = Trim("" & r("perfil"))
		if s = "" then s = "&nbsp;"
		x=x & " <td class='MDB' style='width:200px;' align='left' valign='top'><p class='C'>" & s & "</p></td>"

	 '> VENDEDOR DA LOJA
		s=""
		if r("vendedor_loja") <> 0 then s = Trim("" & r("lista_lojas"))
		if s="" then s="&nbsp;"
		x=x & " <td class='MDB' align='right' valign='top' nowrap><p class='Cd'>" & s & "</p></td>"

	 '> VENDEDOR EXTERNO
		s=""
		if r("vendedor_externo") = 0 then 
			s="<span>Não</span>"
		else 
			s="<span style='color:#006600'>Sim</span>"
			end if
		x=x & " <td class='MDB' align='center' valign='top' nowrap><p class='Cc'>" & s & "</p></td>"

	 '> CD
		s = Trim("" & r("lista_wms_cd"))
		if s="" then s="&nbsp;"
		x=x & " <td class='MDB' align='center' valign='top' nowrap><p class='Cc'>" & s & "</p></td>"

	 '> TELEFONE
		s = formata_ddd_telefone_ramal(Trim("" & r("ddd")), Trim("" & r("telefone")), "")
		if s="" then s="&nbsp;"
		x=x & " <td class='MDB' align='center' valign='top' nowrap><p class='Cc'>" & s & "</p></td>"

	 '> ACESSO
		if r("bloqueado")=0 then 
			s="<span style='color:#006600'>Liberado</span>"
		else 
			s="<span style='color:#ff0000'>Bloqueado</span>"
			end if
		x=x & " <td class='MB' align='left' valign='top' nowrap><p class='C'>" & s & "</p></td>"

		x=x & "</tr>" & chr(13)

		if (i mod 100) = 0 then
			Response.Write x
			x = ""
			end if

		r.MoveNext
		wend


  ' MOSTRA TOTAL DE USUÁRIOS
	x=x & "<tr nowrap style='background: #FFFFDD'><td colspan='8' align='right' nowrap><p class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & cstr(i) & "&nbsp;&nbsp;usuários" & "</p></td></tr>"

  ' FECHA TABELA
	x=x & "</table>"
	

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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando consulta ...';

function fOPConcluir(s_user){
	window.status = "Aguarde ...";
	fOP.usuario_selecionado.value=s_user;
	fOP.submit(); 
}

</script>

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">


<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Relação de Usuários Cadastrados</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  RELAÇÃO DE USUÁRIOS  -->
<br>
<center>
<form method="post" action="usuarioedita.asp" id="fOP" name="fOP">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="usuario_selecionado" id="usuario_selecionado" value=''>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<% executa_consulta %>
</form>

<br>

<p class="TracoBottom"></p>

<table class="notPrint" cellspacing="0">
<tr>
	<td align="center"><a href="usuario.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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