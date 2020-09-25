<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%

'     ===========================
'	  Produto.asp
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
	dim s, usuario, usuario_nome
	usuario = Trim(Session("usuario_atual"))
	usuario_nome = Trim(Session("usuario_nome_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	if Not operacao_permitida(OP_CEN_CADASTRO_PRODUTOS, s_lista_operacoes_permitidas) then 
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

<html>


<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fOPConcluir( f ){
var s_dest, s_op, s_fabricante, s_produto;
	
	s_dest="";
	s_op="";
    s_fabricante = "";
    s_produto = "";
	
	if (f.rb_op[0].checked) {
		s_dest="ProdutoEdita.asp";
		s_op=OP_INCLUI;
        s_fabricante = f.c_fabricante_novo.value;
        s_produto = f.c_produto_novo.value;
        if (trim(f.c_fabricante_novo.value)=="") {
			alert("Informe o código do fabricante do novo produto!");
            f.c_fabricante_novo.focus();
			return false;
			}
        if (trim(f.c_produto_novo.value) == "") {
            alert("Informe o código do novo produto!");
            f.c_produto_novo.focus();
            return false;
        }
		}
	
	if (f.rb_op[1].checked) {
		s_dest="ProdutoEdita.asp";
		s_op=OP_CONSULTA;
        s_fabricante = f.c_fabricante_cons.value;
        s_produto = f.c_produto_cons.value;
        if (trim(f.c_fabricante_cons.value)=="") {
			alert("Informe o código do fabricante do produto a ser consultado!");
            f.c_fabricante_cons.focus();
			return false;
			}
        if (trim(f.c_produto_cons.value) == "") {
            alert("Informe o código do produto a ser consultado!!");
            f.c_produto_cons.focus();
            return false;
        }
		}
		
	if (f.rb_op[2].checked) {
		s_dest="ProdutoLista.asp";
		}

	if (s_dest=="") {
		alert("Escolha uma das opções!!");
		return false;
		}
	
	f.fabricante_selecionado.value=s_fabricante;
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">



<body onload="focus()">

<!--  MENU SUPERIOR -->
<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">CENTRAL&nbsp;&nbsp;ADMINISTRATIVA<br>
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

<br />

<center>
<!--  ***********************************************************************************************  -->
<!--  F O R M U L Á R I O                         												       -->
<!--  ***********************************************************************************************  -->
<form method="post" id="fOP" name="fOP" onsubmit="if (!fOPConcluir(fOP)) return false">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='fabricante_selecionado' id="fabricante_selecionado" value='' />
<input type="hidden" name='produto_selecionado' id="produto_selecionado" value='' />
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='' />

<span class="T">CADASTRO DE PRODUTOS</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td align="left" nowrap>
			<table cellspacing="0" cellpadding="0">
				<tr align="left">
					<td align="left" colspan="2">
						<input type="radio" id="rb_op" name="rb_op" value="1" class="CBOX" onclick="fOP.c_fabricante_novo.focus()">
							<span class="CBOX" style="cursor:default" onclick="fOP.rb_op[0].click(); fOP.c_fabricante_novo.focus();">Cadastrar Novo</span>
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
									<input name="c_fabricante_novo" id="c_fabricante_novo" type="text" maxlength="3" style="width:40px;" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"
										onclick="fOP.rb_op[0].click()"
										onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit(); filtra_numerico();">
								</td>
							</tr>
							<tr>
								<td align="right">
									<span class="L">Produto&nbsp;</span>
								</td>
								<td align="left">
									<input name="c_produto_novo" type="text" maxlength="8" style="width:70px;" onblur="this.value=normaliza_produto(this.value);" onclick="fOP.rb_op[0].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) {this.value=normaliza_produto(this.value); fOPConcluir(fOP);} filtra_produto();">
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
						<input type="radio" id="rb_op" name="rb_op" value="2" class="CBOX" onclick="fOP.c_fabricante_cons.focus()">
							<span class="CBOX" style="cursor:default" onclick="fOP.rb_op[1].click(); fOP.c_fabricante_cons.focus();">Consultar</span>
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
									<input name="c_fabricante_cons" id="c_fabricante_cons" type="text" maxlength="3" style="width:40px;" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"
										onclick="fOP.rb_op[1].click()"
										onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit(); filtra_numerico();">
								</td>
							</tr>
							<tr>
								<td align="right">
									<span class="L">Produto&nbsp;</span>
								</td>
								<td align="left">
									<input name="c_produto_cons" type="text" maxlength="8" style="width:70px;" onblur="this.value=normaliza_produto(this.value);" onclick="fOP.rb_op[1].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) {this.value=normaliza_produto(this.value); fOPConcluir(fOP);} filtra_produto();">
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
			<input type="radio" id="rb_op" name="rb_op" value="3" class="CBOX">
				<span class="rbLink" onclick="fOP.rb_op[2].click(); fOPConcluir(fOP);">Consultar Lista</span>
			</td>
		</tr>
	</table>

	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bEXECUTAR" type="button" class="Botao" value="EXECUTAR" title="executa" onclick="fOPConcluir(fOP);">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
	<td align="center"><a href="MenuCadastro.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>

</body>
</html>
