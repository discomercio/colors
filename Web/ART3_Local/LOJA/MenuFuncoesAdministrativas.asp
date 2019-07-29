<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%

'     ===========================================================
'	  M E N U F U N C O E S A D M I N I S T R A T I V A S . A S P
'     ===========================================================
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
	usuario_nome = Trim(Session("usuario_nome_atual"))
	loja = Trim(Session("loja_atual"))
	loja_nome = Session("loja_nome_atual")
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim idx
	
'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	if loja_nome = "" then loja_nome = trim(x_loja(loja))

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

<script language="JavaScript" type="text/javascript">
window.focus();
</script>

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>


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


<body>
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

<br>
<center>


<!--  ***********************************************************************************************  -->
<!--  F U N Ç Õ E S   A D M I N I S T R A T I V A S												       -->
<!--  ***********************************************************************************************  -->
<span class="T">FUNÇÕES ADMINISTRATIVAS</span>
<div class="QFn" align="CENTER" style="width:600px;">
<table class="TFn">
	<%if operacao_permitida(OP_LJA_CADASTRA_SENHA_DESCONTO, s_lista_operacoes_permitidas) then%>
	<tr>
		<td NOWRAP>
			<form action="SenhaDescSupPesqCliente.asp" id="fSD" name="fSD" method="post" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<input name="bEXECUTAR" id="bEXECUTAR" style='width:340px;' type="submit" class="Botao" value="Senha Desconto  >>" title="senha para desconto superior">
			</form>
			</td>
		</tr>
	<%end if%>

	<%if operacao_permitida(OP_LJA_CONS_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) Or _
		 operacao_permitida(OP_LJA_EDITA_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) Or _
		 operacao_permitida(OP_LJA_CADASTRAMENTO_INDICADOR, s_lista_operacoes_permitidas) then%>
	<tr>
		<td NOWRAP>
			<form action="MenuOrcamentistaEIndicador.asp" id="fSD" name="fSD" method="post" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<input name="bEXECUTAR" id="bEXECUTAR" style='width:340px;' type="submit" class="Botao" value="Orçamentistas / Indicadores  >>" title="cadastro de orçamentistas / indicadores">
			</form>
			</td>
		</tr>
	<%end if%>

	<%if operacao_permitida(OP_LJA_PREENCHER_INDICADOR_EM_PEDIDO_CADASTRADO, s_lista_operacoes_permitidas) then%>
	<tr>
		<td NOWRAP>
			<form action="PedidoPreencheIndicador.asp" id="fPPI" name="fPPI" method="post" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<input name="bEXECUTAR" id="bEXECUTAR" style='width:340px;' type="submit" class="Botao" value="Preencher Indicador em Pedido Cadastrado  >>" title="preencher indicador em pedido já cadastrado">
			</form>
			</td>
		</tr>
	<%end if%>

	</table>
</div>


<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
	<td align="center"><a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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
