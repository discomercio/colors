<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%

'     =========================
'	  ECProdutoCompostoMenu.asp
'     =========================
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



' _____________________________________________________________________________________________
'
'						I N I C I A L I Z A     P Á G I N A     A S P
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
'	OBTEM USUÁRIO
	dim s, usuario, usuario_nome
	usuario = trim(Session("usuario_atual"))
	usuario_nome = Trim(Session("usuario_nome_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	if Not operacao_permitida(OP_CEN_CAD_EC_PRODUTO_COMPOSTO, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

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

<script src="../GLOBAL/global.js" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript">
function fOPConcluir( f ){
var s_dest,s_op,s_fabricante,s_produto;
	
	s_dest="";
	s_op="";
	s_fabricante="";
	s_produto="";
	
	if (f.rb_op[0].checked) {
		s_dest="ECProdutoCompostoNovo.asp";
		s_op=OP_INCLUI;
		s_fabricante=f.c_fabricante_inclui.value;
		s_produto=f.c_produto_inclui.value;
		if (trim(f.c_fabricante_inclui.value)=="") {
			alert("Forneça o código do fabricante para o produto composto !!");
			f.c_fabricante_inclui.focus();
			return false;
			}
		if (trim(f.c_produto_inclui.value)=="") {
			alert("Forneça o código do produto para o produto composto !!");
			f.c_produto_inclui.focus();
			return false;
			}
		}
		
	if (f.rb_op[1].checked) {
		s_dest="ECProdutoCompostoEdita.asp";
		s_op=OP_CONSULTA;
		s_fabricante=f.c_fabricante_consulta.value;
		s_produto=f.c_produto_consulta.value;
		if (trim(f.c_fabricante_consulta.value)=="") {
			alert("Informe o código do fabricante !!");
			f.c_fabricante_consulta.focus();
			return false;
			}
		if (trim(f.c_produto_consulta.value)=="") {
			alert("Informe o código do produto !!");
			f.c_produto_consulta.focus();
			return false;
			}
		}
		
	if (f.rb_op[2].checked) {
		s_dest="ECProdutoCompostoLista.asp";
		}

	if (s_dest=="") {
		alert("Escolha uma das opções !!");
		return false;
		}
	
	f.fabricante_selecionado.value=s_fabricante;
	f.produto_selecionado.value=s_produto;
	f.operacao_selecionada.value=s_op;
	
	f.action=s_dest;
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
<link href="../global/e.css" Rel="stylesheet" Type="text/css">



<body onload="focus()">

<!--  MENU SUPERIOR -->  
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">CENTRAL&nbsp;&nbsp;ADMINISTRATIVA</span><br>
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
		</span></td>
	</tr>

</table>

<br />

<center>
<!--  ***********************************************************************************************  -->
<!--  F O R M U L Á R I O                         												       -->
<!--  ***********************************************************************************************  -->
<form method="post" id="fOP" name="fOP">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='fabricante_selecionado' value=''>
<input type="hidden" name='produto_selecionado' value=''>
<input type="hidden" name='operacao_selecionada' value=''>

<span class="T">E-Commerce: Cadastro de Produto Composto</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td align="left" nowrap>
			<table cellspacing="0" cellpadding="0">
				<tr align="left">
					<td align="left" colspan="2">
						<input type="radio" id="s1" name="rb_op" value="1" class="CBOX">
						<span class="CBOX" style="cursor:default" onclick="fOP.rb_op[0].click();fOP.c_fabricante_inclui.focus();">Cadastrar Novo</span>
					</td>
				</tr>
				<tr>
					<td align="left" style="width:40px;">
						&nbsp;
					</td>
					<td align="left">
						<table cellspacing="0" cellpadding="0">
							<tr>
								<td align="right" style="padding-bottom:4px;">
									<span class="L">Fabricante&nbsp;</span>
								</td>
								<td align="left" style="padding-bottom:4px;">
									<input name="c_fabricante_inclui" type="text" maxlength="3" style="width:40px;" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onclick="fOP.rb_op[0].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) {this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE); fOP.c_produto_inclui.focus();} filtra_fabricante();">
								</td>
							</tr>
							<tr>
								<td align="right">
									<span class="L">Produto&nbsp;</span>
								</td>
								<td align="left">
									<input name="c_produto_inclui" type="text" maxlength="6" style="width:70px;" onblur="this.value=normaliza_produto(this.value);" onclick="fOP.rb_op[0].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) {this.value=normaliza_produto(this.value); fOPConcluir(fOP);} filtra_produto();">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align="left" nowrap>
			<table cellspacing="0" cellpadding="0">
				<tr align="left">
					<td align="left" colspan="2">
						<input type="radio" id="s1" name="rb_op" value="2" class="CBOX">
						<span class="CBOX" style="cursor:default" onclick="fOP.rb_op[1].click();fOP.c_fabricante_consulta.focus();">Consultar</span>
					</td>
				</tr>
				<tr>
					<td align="left" style="width:40px;">
						&nbsp;
					</td>
					<td align="left">
						<table cellspacing="0" cellpadding="0">
							<tr>
								<td align="right" style="padding-bottom:4px;">
									<span class="L">Fabricante&nbsp;</span>
								</td>
								<td align="left" style="padding-bottom:4px;">
									<input name="c_fabricante_consulta" type="text" maxlength="3" style="width:40px;" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onclick="fOP.rb_op[1].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) {this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE); fOP.c_produto_consulta.focus();} filtra_fabricante();">
								</td>
							</tr>
							<tr>
								<td align="right">
									<span class="L">Produto&nbsp;</span>
								</td>
								<td align="left">
									<input name="c_produto_consulta" type="text" maxlength="6" style="width:70px;" onblur="this.value=normaliza_produto(this.value);" onclick="fOP.rb_op[1].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) {this.value=normaliza_produto(this.value); fOPConcluir(fOP);} filtra_produto();">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="left">
		<td align="left" nowrap>
			<input type="radio" id="s1" name="rb_op" value="3" class="CBOX"><span class="rbLink" onclick="fOP.rb_op[2].click(); fOPConcluir(fOP);">Consultar Lista</span>
		</td>
	</tr>
</table>

<br />
<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
<input name="bEXECUTAR" type="button" class="Botao" value="EXECUTAR" title="executa" onclick="fOPConcluir(fOP);">
<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>

<br />

<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
	<td align="center">
		<a href="MenuCadastro.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
			<img src="../botao/voltar.gif" width="176" height="55" border="0">
		</a>
	</td>
</tr>
</table>

</center>

</body>
</html>
