<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ====================================================
'	  E S T O Q U E T R A N S F E R E P E D I D O . A S P
'     ====================================================
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_TRANSF_ENTRE_PED_PROD_ESTOQUE_VENDIDO, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim intCounter
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



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fOPConclui( f ) {
var i, b, ha_item;

	if (trim(f.c_pedido_origem.value)=="") {
		alert("Informe o nº do pedido que irá ceder as mercadorias disponíveis!!");
		f.c_pedido_origem.focus();
		return;
		}

	if (trim(f.c_pedido_destino.value)=="") {
		alert("Informe o nº do pedido que irá receber as mercadorias!!");
		f.c_pedido_destino.focus();
		return;
		}

	if (f.c_pedido_origem.value==f.c_pedido_destino.value) {
		alert("Pedido de origem e de destino devem ser diferentes!!");
		f.c_pedido_origem.focus();
		return;
		}
		
	ha_item=false;
	for (i=0; i < f.c_fabricante.length; i++) {
		b=false;
		
		if (converte_numero(f.c_qtde[i].value)!=0) b=true;
		if (trim(f.c_fabricante[i].value)!="") b=true;
		if (trim(f.c_produto[i].value)!="") b=true;
		
		if (b) {
			ha_item=true;
			if (converte_numero(f.c_qtde[i].value)<=0) {
				alert("Quantidade inválida!!");
				f.c_qtde[i].focus();
				return;
				}
	
			if (trim(f.c_produto[i].value)!="") {
				if (!isEAN(trim(f.c_produto[i].value))) {
					if (trim(f.c_fabricante[i].value)=="") {
						alert("Informe o fabricante do produto!!");
						f.c_fabricante[i].focus();
						return;
						}
					}
				}

			if (trim(f.c_produto[i].value)=="") {
				alert("Informe o código do produto a ser transferido!!");
				f.c_produto[i].focus();
				return;
				}
			}
		}

	if (!ha_item) {
		alert("Nenhum produto foi informado!!");
		f.c_qtde[0].focus();
		return;
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">


<body onload="if (trim(fOP.c_pedido_origem.value)=='') fOP.c_pedido_origem.focus();">
<center>

<form id="fOP" name="fOP" method="post" action="EstoqueTransferePedidoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Transferência Entre Pedidos de Produtos do Estoque Vendido</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS -->
<table class="Qx" cellspacing="0">
<!--  Nº PEDIDO  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="3" class="MT" nowrap><span class="PLTe">Pedido Origem</span>
		<br><input maxlength="10" class="PLLe" name="c_pedido_origem" id="c_pedido_origem" style="width:150px; margin-right:15px;" 
				onkeypress="if (digitou_enter(true) && tem_info(this.value)) {if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value); fOP.c_pedido_destino.focus();} filtra_pedido();"
				onblur="if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value);">
	</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="3" class="MDBE" nowrap><span class="PLTe">Pedido Destino</span>
		<br><input maxlength="10" class="PLLe" name="c_pedido_destino" id="c_pedido_destino" style="width:150px; margin-right:15px;"
				onkeypress="if (digitou_enter(true) && tem_info(this.value)) {if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value); fOP.c_qtde[0].focus();} filtra_pedido();"
				onblur="if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value);">
	</td>
	</tr>

<!--  QTDE/FABRICANTE/PRODUTO  -->
<%
   for intCounter=1 to MAX_ITENS_TRANSF_PRODUTOS_ENTRE_PEDIDOS
%>
	<tr bgcolor="#FFFFFF">
	<td class="PLTd" align="right" style="vertical-align:bottom;"><%=CStr(intCounter) & ".&nbsp;"%></td>
	<td class="MDBE" nowrap>
		<% if intCounter=1 then %>
		<span class="PLTe">Qtde</span>
		<br>
		<%end if%>
		<input name="c_qtde" id="c_qtde" class="PLLc" maxlength="4" 
			style="width:35px;" 
			onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(<%=Cstr(intCounter)%>!=1))) if (trim(this.value)=='') bCONFIRMA.focus(); else fOP.c_fabricante[<%=Cstr(intCounter-1)%>].focus(); filtra_numerico();">
	</td>
	<td class="MDB" nowrap>
		<% if intCounter=1 then %>
		<span class="PLTe" style="margin-right:2pt;">Fabricante</span>
		<br>
		<%end if%>
		<input name="c_fabricante" id="c_fabricante" class="PLLe" maxlength="4" 
			style="margin-left:2pt;width:30px;" 
			onkeypress="if (digitou_enter(true)&&tem_info(this.value)) c_produto[<%=Cstr(intCounter-1)%>].focus(); filtra_fabricante();" 
			onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);">
	</td>
	<td class="MDB" style="border-left:0pt;">
		<% if intCounter=1 then %>
		<span class="PLTe">Produto</span>
		<br>
		<%end if%>
		<input name="c_produto" id="c_produto" class="PLLe" maxlength="13" 
			style="margin-left:2pt;width:100px;" 
			onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {if (<%=Cstr(intCounter)%>==fOP.c_produto.length) bCONFIRMA.focus(); else fOP.c_qtde[<%=Cstr(intCounter)%>].focus();} filtra_produto();" 
			onblur="this.value=normaliza_codigo(this.value,TAM_MIN_PRODUTO); this.value=ucase(trim(this.value));">
	</td>
	</tr>

<% next %>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fOPConclui(fOP)" title="executa a transferência">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>
