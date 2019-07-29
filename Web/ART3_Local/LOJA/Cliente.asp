<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%

'     ===========================
'	  C L I E N T E . A S P
'     ===========================
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
'						I N I C I A L I Z A     P Á G I N A     A S P
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
'	OBTEM USUÁRIO
	dim s, usuario, usuario_nome, loja, loja_nome
	usuario = trim(Session("usuario_atual"))
	loja = Session("loja_atual")
	usuario_nome = Trim(Session("usuario_nome_atual"))
	loja_nome = Trim(Session("loja_nome_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

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
	<title>LOJA</title>
	</head>

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function restauraVisibility(nome_controle) {
	var c;
	c = document.getElementById(nome_controle);
	if (c) c.style.visibility = "";
}

function fPESQConcluir( f ){
var s, c;
	s=f.cnpj_cpf_selecionado.value;
	s=retorna_so_digitos(s);
	if (!cnpj_cpf_ok(s)) {
		alert("CNPJ/CPF inválido!!");
		f.cnpj_cpf_selecionado.focus();
		return false;
		}

	c = document.getElementById("bPesqCliEXECUTAR");
	c.style.visibility = "hidden";
	setTimeout("restauraVisibility('bPesqCliEXECUTAR')", 30000);

	window.status = "Aguarde ...";
	f.submit();
}

function fOPConcluir( f ){
var s_dest, s_op, s_cnpj_cpf, c;
	
	s_dest="";
	s_op="";
	s_cnpj_cpf="";
	
	if (f.rb_op[0].checked) {
		s_dest="clienteedita.asp";
		s_op=OP_INCLUI;
		s_cnpj_cpf=trim(f.c_novo.value);
		if ((s_cnpj_cpf=="")||(!cnpj_cpf_ok(s_cnpj_cpf))) {
			alert("CNPJ/CPF inválido!!");
			f.c_novo.focus();
			return false;
			}
		}
		
	if (f.rb_op[1].checked) {
		s_dest="clientepesquisa.asp";
		s_op=OP_CONSULTA;
		s_cnpj_cpf=trim(f.c_cons.value);
		if ((s_cnpj_cpf=="")||(!cnpj_cpf_ok(s_cnpj_cpf))) {
			alert("CNPJ/CPF inválido!!");
			f.c_cons.focus();
			return false;
			}
		}
		
	if (s_dest=="") {
		alert("Escolha uma das opções!!");
		return false;
		}
	
	f.cnpj_cpf_selecionado.value=s_cnpj_cpf;
	f.operacao_selecionada.value=s_op;

	c = document.getElementById("bCadCliEXECUTAR");
	c.style.visibility = "hidden";
	setTimeout("restauraVisibility('bCadCliEXECUTAR')", 30000);
	
	window.status = "Aguarde ...";
	f.action=s_dest;
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



<body onload="focus()">

<!--  MENU SUPERIOR -->
<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="RIGHT" vAlign="BOTTOM"><p class="PEDIDO"><% = loja_nome & " (" & loja & ")" %><br>
	<%	s = usuario_nome
		if s = "" then s = usuario
		s = x_saudacao & ", " & s
		s = "<span class='Cd'>" & s & "</span><br>"
	%>
	<%=s%>
	<span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="senha.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="altera a senha atual do usuário" class="LAlteraSenha">altera senha</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></p></td>
	</tr>

</table>


<center>
<!--  ***********************************************************************************************  -->
<!--  F O R M U L Á R I O                         												       -->
<!--  ***********************************************************************************************  -->
<form method="post" id="fOP" name="fOP" onsubmit="if (!fOPConcluir(fOP)) return false">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<INPUT type="hidden" name='cnpj_cpf_selecionado' id="cnpj_cpf_selecionado" value=''>
<INPUT type="hidden" name='operacao_selecionada' id="operacao_selecionada" value=''>

<span class="T">CADASTRO DE CLIENTES</span>
<div class="QFn" align="CENTER">
<table class="TFn">
	<tr>
		<td NOWRAP>
			<input type="radio" name="rb_op" id="rb_op" value="1" class="CBOX" onclick="fOP.c_novo.focus();"><span style="cursor:default" onclick="fOP.rb_op[0].click();fOP.c_novo.focus();">Cadastrar Novo CNPJ/CPF</span>&nbsp;
				<input name="c_novo" id="c_novo" type="text" maxlength="18" size="20" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value) && cnpj_cpf_ok(this.value)) {this.value=cnpj_cpf_formata(this.value); if (fOPConcluir(fOP)) submit();} filtra_cnpj_cpf();" onclick="fOP.rb_op[0].click();"><br>
			<input type="radio" name="rb_op" id="rb_op" value="2" class="CBOX" onclick="fOP.c_cons.focus();"><span style="cursor:default" onclick="fOP.rb_op[1].click();fOP.c_cons.focus();">Consultar CNPJ/CPF</span>&nbsp;
				<input name="c_cons" id="c_cons" type="text" maxlength="18" size="20" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onclick="fOP.rb_op[1].click();" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value) && cnpj_cpf_ok(this.value)) {this.value=cnpj_cpf_formata(this.value); if (fOPConcluir(fOP)) submit();} filtra_cnpj_cpf();" onclick="fOP.rb_op[0].click();"><br>
			</td>
		</tr>
	</table>

	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bCadCliEXECUTAR" id="bCadCliEXECUTAR" type="submit" class="Botao" value="EXECUTAR" title="executa">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>

<!--  ***********************************************************************************************  -->
<!--  F O R M U L Á R I O                         												       -->
<!--  ***********************************************************************************************  -->
<br>
<form action="ClientePesquisa.asp" method="post" id="fPESQ" name="fPESQ" OnSubmit="if (!fPESQConcluir(fPESQ)) return false">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span class="T">PESQUISAR CLIENTE POR</span>
<div class="QFn" align="CENTER">
<table class="TFn">
	<tr>
		<td NOWRAP>
			<table cellPadding="0" CellSpacing="0">
			<tr><td NOWRAP class="C" align="right">CNPJ/CPF&nbsp;</td><td><input name="cnpj_cpf_selecionado" id="cnpj_cpf_selecionado" type="text" maxlength="18" size="20" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="if (digitou_enter(true) && tem_info(this.value) && cnpj_cpf_ok(this.value)) {this.value=cnpj_cpf_formata(this.value); if (fPESQConcluir(fPESQ)) submit();} filtra_cnpj_cpf();"></td></tr>
			<tr><td NOWRAP class="C" align="right">NOME&nbsp;</td><td><input name="nome_selecionado" id="nome_selecionado" type="text" maxlength="60" size="45" onkeypress="if (digitou_enter(true) && tem_info(this.value)) if (fPESQConcluir(fPESQ)) submit(); filtra_nome_identificador();"></td></tr>
			</table>
			</td>
		</tr>
	</table>

	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bPesqCliEXECUTAR" id="bPesqCliEXECUTAR" type="submit" class="Botao" value="EXECUTAR" title="executa">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>

<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
	<td align="CENTER"><a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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
