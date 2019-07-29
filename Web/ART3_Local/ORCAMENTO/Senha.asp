<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<%
'     =================
'	  S E N H A . A S P
'     =================
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
	
'	VERIFICA ID
	dim s, loja, loja_nome, usuario, usuario_nome, senha
	
'	OBTEM O ID
	loja = Session("loja_atual")
	usuario = Session("usuario_atual")
	senha = Session("senha_atual")
	usuario_nome = Trim(Session("usuario_nome_atual"))
	loja_nome = Trim(Session("loja_nome_atual"))

	if (loja="") or (usuario="") or (senha="") then Response.Redirect("Aviso.asp?id=" & ERR_SESSAO)

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("Aviso.asp?id=" & ERR_CONEXAO)

	dim primeira_vez, dt_ult_alteracao_senha
	dt_ult_alteracao_senha = Null
	set rs = cn.execute("SELECT dt_ult_alteracao_senha FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido = '" & usuario & "'")
	if not rs.eof then dt_ult_alteracao_senha = rs("dt_ult_alteracao_senha")

	primeira_vez = IsNull(dt_ult_alteracao_senha)
	
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
	<title>SENHA</title>
	</head>

<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
configura_painel();

function posiciona_foco( f ){
	if (trim(f.senha.value)==""){ 
		f.senha.focus();
		return true;
		}
	if (trim(f.novasenha.value)==""){ 
		f.novasenha.focus();
		return true;
		}
	if (trim(f.novasenha2.value)==""){ 
		f.novasenha2.focus();
		return true;
		}
}

function confere( f ){
var s1, s2, s3;
	s1 = ucase(trim(f.senha.value));
	s2 = ucase(trim(f.novasenha.value));
	s3 = ucase(trim(f.novasenha2.value));
	if (s1.length < 5){ 
		f.senha.focus();
		return false;
		}
	if (s2.length < 5){ 
		alert("A nova senha deve possuir no mínimo 5 caracteres.");
		f.novasenha.focus();
		return false;
		}
	if (s3.length < 5){ 
		alert("A confirmação da nova senha deve possuir no mínimo 5 caracteres.");
		f.novasenha2.focus();
		return false;
		}
	if (s2!=s3){
		alert("A confirmação da nova senha está incorreta.");
		f.novasenha.value="";
		f.novasenha2.value="";
		f.novasenha.focus();
		return false;
		}
	if (s1==s2){
		alert("A nova senha deve ser diferente da senha atual.");
		f.novasenha.value="";
		f.novasenha2.value="";
		f.novasenha.focus();
		return false;
		}
	if (s2 == ucase(trim(f.c_usuario.value))) {
		alert("A nova senha não pode ser igual ao identificador do usuário!");
		f.novasenha.value = "";
		f.novasenha2.value = "";
		f.novasenha.focus();
		return false;
		}
	
	return true;
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">



<body onload="posiciona_foco(fPWD);">
<!--  L O G O T I P O    E    L O J A  -->

<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">
	<%	s = usuario_nome
		if s = "" then s = usuario
		s = x_saudacao & ", " & s
		s = "<span class='Cd'>" & s & "</span><br>"
	%>
	<%=s%>
	<span class="Rc">
		<a href="sessaoencerra.asp" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></p></td>
	</tr>

</table>

<br>
<center>


<!--  A L T E R A Ç Ã O    D E    S E N H A  -->

<p class="T">ALTERAÇÃO DE SENHA
<% 
	s = ""
	if primeira_vez then s = "<br>(senha expirada)"
%>
<%=s%>
</p>


<form action="SenhaAtualiza.asp" method="post" id="fPWD" name="fPWD">
<input type="hidden" name="c_usuario" id="c_usuario" value="<%=usuario%>" />

<div class="QFn" style="width:300px" align="center">

	<p class="R" style="margin: 10 10 2 10">SENHA ATUAL</p>
	<input name="senha" id="senha" type="password" maxlength="15" style="text-align:center;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPWD.novasenha.focus();">
	
	<p class="R" style="margin: 10 10 2 10">SENHA NOVA</p>
	<input name="novasenha" id="novasenha" type="password" maxlength="15" style="text-align:center;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPWD.novasenha2.focus();">
	
	<p class="R" style="margin: 10 10 2 10">SENHA NOVA (CONFIRMA)</p>
	<input name="novasenha2" id="novasenha2" type="password" maxlength="15" style="text-align:center;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) if (confere(fPWD)) submit();">
	
	<p class="R" style="margin: 4 10 0 10">&nbsp;</p>
	<input name="CONSULTAR" id="CONSULTAR" type="button" class="Botao" 
		   value="CONFIRMAR" title="confirma a alteração da senha do usuário" onclick="if (confere(fPWD)) submit();">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
	
</div>
</form>

<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
<% if primeira_vez then %>
	<td align="center"><a href="SessaoEncerra.asp" title="cancela a alteração de senha e encerra a sessão">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
<% else %>
	<td align="center"><a href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
<% end if %>

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
