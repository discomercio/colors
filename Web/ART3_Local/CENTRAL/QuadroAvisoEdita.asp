<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================
'	  Q U A D R O A V I S O E D I T A . A S P
'     =======================================
'
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
	
'	OBTEM O ID
	dim s, usuario, operacao_selecionada
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
	operacao_selecionada = trim(request("operacao_selecionada"))

'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim id_aviso, msg_erro
	if operacao_selecionada=OP_INCLUI then
		if Not gera_nsu(NSU_QUADRO_AVISO, id_aviso, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
	else
		id_aviso = trim(request("aviso_selecionado"))
		if id_aviso = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
		end if

	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	
	set rs = cn.Execute("select * from t_AVISO where (id='" & id_aviso & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)

	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_NSU_JA_EM_USO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_NSU_NAO_LOCALIZADO)
		end if
	
%>

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

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function RemoveAviso( f ) {
var b;
	b=window.confirm('Confirma a exclusão deste aviso?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaAviso( f ) {
	if (trim(f.mensagem.value)=="") {
		alert("Não há texto na mensagem de aviso!!");
		f.mensagem.focus();
		return;
		}
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.mensagem.focus();"
	else
		s = "focus();"
		end if
%>
<body onload="<%=s%>">
<center>



<!--  QUADRO DE AVISOS -->

<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Novo Aviso"
	else
		s = "Consulta/Edição de Aviso Cadastrado"
		end if
%>
	<td align="center" valign="bottom"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS  -->
<form id="fCAD" name="fCAD" method="post" action="QuadroAvisoAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=operacao_selecionada%>'>
<input type="hidden" name="aviso_selecionado" id="aviso_selecionado" value='<%=id_aviso%>'>

<!-- ************   MENSAGEM   ************ -->
<table width="649" cellSpacing="0">
	<tr><td width="100%">
	<span class="PLTe">TEXTO DA MENSAGEM</span>
	<br>
	<%	s=""
		if Not rs.EOF then 
			if Not IsNull(rs("mensagem")) then s = rs("mensagem")
			end if	%>
	<textarea id="mensagem" name="mensagem" class="QuadroAviso" onkeypress="filtra_nome_identificador();"><%=s%></textarea>
	</td></tr>
</table>

<!-- ************   DESTINATÁRIO DA MENSAGEM   ************ -->
<br>
<table width="649" cellSpacing="0">
	<tr><td width="100%">
	<span class="PLTe">LOJA DESTINATÁRIA</span>
	<br>
	<%	s=""
		if Not rs.EOF then s = Trim("" & rs("destinatario"))
	%>
	<input maxlength="3" class="Cc" style="width:50px;" name="c_destinatario" id="c_destinatario" 
		value='<%=s%>' onblur="this.value=normaliza_codigo(this.value,TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true)) bATUALIZA.focus(); filtra_numerico();">
	</td></tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a href="javascript:history.back()" title="cancela as alterações no cadastro">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		s = "<td align='center'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveAviso(fCAD)' "
		s =s + "title='remove o aviso cadastrado'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaAviso(fCAD)" title="atualiza o quadro de avisos">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	rs.Close
	set rs = nothing
	
	cn.Close
	set cn = nothing
%>