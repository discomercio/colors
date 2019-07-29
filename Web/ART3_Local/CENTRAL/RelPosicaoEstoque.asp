<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  R E L P O S I C A O E S T O Q U E . A S P
'     ===============================================
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


	On Error GoTo 0
	Err.Clear

	dim usuario, s
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if (Not operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_BASICA, s_lista_operacoes_permitidas)) And (Not operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas)) then 
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
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fESTOQConsulta( f ) {
var i, b;
	b=false;
	for (i=0; i<f.rb_estoque.length; i++) {
		if (f.rb_estoque[i].checked) {
			b=true;
			break;
			}
		}
	if (!b) {
		alert("Selecione o estoque a ser consultado!!");
		return;
		}

	b=false;
	for (i=0; i<f.rb_detalhe.length; i++) {
		if (f.rb_detalhe[i].checked) {
			b=true;
			break;
			}
		}
	if (!b) {
		alert("Selecione o tipo de detalhamento da consulta!!");
		return;
		}

	if (trim(f.c_produto.value)!="") {
		if (!isEAN(trim(f.c_produto.value))) {
			if (trim(f.c_fabricante.value)=="") {
				alert("Informe o fabricante do produto!!");
				f.c_fabricante.focus();
				return;
				}
			}
		}
	
	dCONFIRMA.style.visibility="hidden";
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

<style type="text/css">
#rb_estoque {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
#rb_detalhe {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
.rbOpt
{
	vertical-align:bottom;
}
.lblOpt
{
	vertical-align:bottom;
}
</style>


<body onload="if (trim(fESTOQ.c_fabricante.value)=='') fESTOQ.c_fabricante.focus();">
<center>

<form id="fESTOQ" name="fESTOQ" method="post" action="RelPosicaoEstoqueExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque (Antigo)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS DA CONSULTA  -->
<table class="Qx" cellSpacing="0">
<!--  ESTOQUE  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MT" NOWRAP><span class="PLTe">Estoque de Interesse</span>
		<%	if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
		<br><input type="radio" class="rbOpt" <%=s%> tabindex="-1" id="rb_estoque" name="rb_estoque" value="<%=ID_ESTOQUE_VENDA%>"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[0].click();"
			>Venda</span>
			
		<%	if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
		<br><input type="radio" class="rbOpt" <%=s%> tabindex="-1" id="rb_estoque" name="rb_estoque" value="<%=ID_ESTOQUE_VENDIDO%>"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[1].click();"
			>Vendido</span>

		<%	if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
		<br><input type="radio" class="rbOpt" <%=s%> tabindex="-1" id="rb_estoque" name="rb_estoque" value="<%=ID_ESTOQUE_SHOW_ROOM%>"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[2].click();"
			>Show-Room</span>

		<%	if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_BASICA, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
		<br><input type="radio" class="rbOpt" <%=s%> tabindex="-1" id="rb_estoque" name="rb_estoque" value="<%=ID_ESTOQUE_DANIFICADOS%>"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[3].click();"
			>Produtos Danificados</span>

		<%	if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_BASICA, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
		<br><input type="radio" class="rbOpt" <%=s%> tabindex="-1" id="rb_estoque" name="rb_estoque" value="<%=ID_ESTOQUE_DEVOLUCAO%>"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[4].click();"
			>Devolução</span>
	</td>
	</tr>

<!--  TIPO DE DETALHAMENTO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MDBE" NOWRAP><span class="PLTe">Tipo de Detalhamento</span>
		<%
			s=" disabled" 
			if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas) then s=""
			if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_BASICA, s_lista_operacoes_permitidas) And (Not operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas)) then s=" checked"
		%>
		<br><input type="radio" class="rbOpt" <%=s%> tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="SINTETICO"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_detalhe[0].click();"
			>Sintético (sem custos)</span>

		<%	if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
		<br><input type="radio" class="rbOpt" <%=s%> tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="INTERMEDIARIO"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_detalhe[1].click();"
			>Intermediário (custos médios)</span>

		<%	if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
		<br><input type="radio" class="rbOpt" <%=s%> tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="COMPLETO"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_detalhe[2].click();"
			>Completo (custos diferenciados)</span>
	</td>
	</tr>

<!--  FABRICANTE/PRODUTO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">Fabricante</span>
		<br><input name="c_fabricante" id="c_fabricante" class="PLLe" maxlength="4" style="margin-left:2pt;width:50px;" onkeypress="if (digitou_enter(true)) fESTOQ.c_produto.focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"></td>
	<td class="MDB" style="border-left:0pt;"><span class="PLTe">Produto</span>
		<br><input name="c_produto" id="c_produto" class="PLLe" maxlength="13" style="margin-left:2pt;width:100px;" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_produto();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_PRODUTO); this.value=ucase(trim(this.value));"></td>
	</tr>

</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fESTOQConsulta(fESTOQ)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>
