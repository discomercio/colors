<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================
'	  MensagemAlertaProdutoEdita.asp
'     ===============================
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
	dim s, usuario, alerta_selecionado, operacao_selecionada
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	MENSAGEM A EDITAR
	alerta_selecionado = trim(request("alerta_selecionado"))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if (alerta_selecionado="") then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICADOR_NAO_FORNECIDO) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	
'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	set rs = cn.Execute("SELECT * FROM t_ALERTA_PRODUTO WHERE (apelido='" & alerta_selecionado & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICADOR_JA_CADASTRADO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_REGISTRO_NAO_CADASTRADO)
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
function RemoveCadastro( f ) {
var b;
	b=window.confirm('Confirma a exclusão desta mensagem de alerta?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaCadastro( f ) {
var blnOk;

	blnOk=false;
	if (f.rb_ativo[0].checked) blnOk=true;
	if (f.rb_ativo[1].checked) blnOk=true;
	if (!blnOk) {
		alert('Informe se a mensagem deve ser exibida ou não!!');
		return;
		}
	
	if (trim(f.c_mensagem.value)=="") {
		alert('Preencha a mensagem de alerta!!');
		f.c_mensagem.focus();
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
		s = "fCAD.c_mensagem.focus()"
	else
		s = "focus()"
		end if
%>
<body onLoad="<%=s%>">
<center>



<!--  CADASTRO DA MENSAGEM DE ALERTA -->

<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Nova Mensagem de Alerta para Produtos"
	else
		s = "Consulta/Edição de Mensagem de Alerta para Produtos"
		end if
%>
	<td align="CENTER" vAlign="BOTTOM"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="MensagemAlertaProdutoAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>

<!-- ************   APELIDO   ************ -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
		<td class="MD" width="50%">
			<p class="R">ID</p>
			<p class="C">
				<input id="alerta_selecionado" name="alerta_selecionado" class="TA" value="<%=alerta_selecionado%>" readonly tabindex=-1 size="14" style="text-align:left; color:#0000ff">
			</p>
		</td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ativo")) else s=""%>
		<td width="50%">
			<p class="R">EXIBIR ALERTA</p>
			<p class="C">
				<span style="width:20px;"></span>
				<input type="radio" id="rb_ativo" name="rb_ativo" value="S"
					<% if s = "S" then Response.Write " checked"%>>
				<span class="C" style="cursor:default;" onclick="fCAD.rb_ativo[0].click();">Sim</span>
				&nbsp;<span style="width:20px;"></span>&nbsp;
				<input type="radio" id="rb_ativo" name="rb_ativo" value="N"
					<% if s = "N" then Response.Write " checked"%>>
				<span class="C" style="cursor:default;" onclick="fCAD.rb_ativo[1].click();">Não</span>
			</p>
		</td>
	</tr>
</table>
<br>

<!-- ************   MENSAGEM   ************ -->
<table width="649" cellSpacing="0">
	<tr><td width="100%">
	<span class="PLTe">TEXTO DA MENSAGEM DE ALERTA</span>
	<br>
	<%	if operacao_selecionada=OP_CONSULTA then s = Trim("" & rs("mensagem")) else s="" %>
	<textarea id="c_mensagem" name="c_mensagem" TYPE="textarea" class="QuadroAviso" onkeypress="filtra_nome_identificador();"><%=s%></textarea>
	</td></tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
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
		s = "<td align='CENTER'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveCadastro(fCAD)' "
		s =s + "title='remove esta mensagem de alerta'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="RIGHT"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaCadastro(fCAD)" title="atualiza o cadastro">
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