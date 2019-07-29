<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  R O M A N E I O F I L T R O . A S P
'     ========================================================
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
	
	dim c_transportadora, c_dt_entrega, c_qtde_pedidos, c_lista_pedidos_selecionados, c_nfe_emitente
	c_transportadora = Trim(Request("c_transportadora"))
	c_dt_entrega = Trim(Request("c_dt_entrega"))
	c_lista_pedidos_selecionados = Trim(Request("c_lista_pedidos_selecionados"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
	
	dim alerta
	alerta=""
	
	dim i, vPedido, s_campo_pedidos
	s_campo_pedidos = ""
	vPedido = Split(c_lista_pedidos_selecionados, "|")
	for i=LBound(vPedido) to UBound(vPedido)
		if Trim("" & vPedido(i)) <> "" then
			if s_campo_pedidos <> "" then s_campo_pedidos = s_campo_pedidos & vbCrLf
			s_campo_pedidos = s_campo_pedidos & Trim("" & vPedido(i))
			end if
		next
	
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
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {

	if (trim(f.c_transportadora.value)=="") {
		alert("Informe a transportadora!!");
		f.c_transportadora.focus();
		return;
		}

	if (trim(f.c_dt_entrega.value)=="") {
		alert("Informe a data de coleta!!");
		f.c_dt_entrega.focus();
		return;
		}
		
	if (!isDate(f.c_dt_entrega)) {
		alert("Data inválida!!");
		f.c_dt_entrega.focus();
		return;
		}

	if (trim(f.c_conferente.value) == "") {
		alert("Informe o nome do conferente!!");
		f.c_conferente.focus();
		return;
		}

	if (trim(f.c_motorista.value) == "") {
		alert("Informe o nome do motorista!!");
		f.c_motorista.focus();
		return;
		}

	if (trim(f.c_placa_veiculo.value) == "") {
		alert("Informe a placa do veículo!!");
		f.c_placa_veiculo.focus();
		return;
		}

	if (!isPlacaVeiculoOk(f.c_placa_veiculo.value)) {
		alert("Placa de veículo inválida!!");
		f.c_placa_veiculo.focus();
		return;
		}
	
	if (trim(f.c_pedidos.value)=="") {
		alert("Informe o(s) pedido(s)!!");
		f.c_pedidos.focus();
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="A1" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>
<% else %>

<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="fFILTRO.c_num_coleta.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RomaneioConsiste.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_entrega" id="c_dt_entrega" value="<%=c_dt_entrega%>" />
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>" />
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Romaneio de Entrega</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellspacing="0" style="width:190px;">
<!--  TRANSPORTADORA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MT" align="left" nowrap><span class="PLTe">TRANSPORTADORA</span>
	<br>
		<span class="C" style="color:#000080;"><%=c_transportadora%></span>
		</td></tr>

<!--  DATA DE COLETA (RÓTULO ANTIGO: DATA DA ENTREGA)  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">DATA DE COLETA</span>
	<br>
		<span class="C" style="color:#000080;"><%=c_dt_entrega%></span>
	</td></tr>

<!--  Nº COLETA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">Nº COLETA</span>
	<br>
		<input maxlength="8" class="PLLe" style="width:120px;" name="c_num_coleta" id="c_num_coleta" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_transportadora_contato.focus();">
	</td></tr>

<!--  CONTATO NA TRANSPORTADORA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">CONTATO NA TRANSPORTADORA</span>
	<br>
		<input class="PLLe" maxlength="10" style="width:120px;" name="c_transportadora_contato" id="c_transportadora_contato" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_conferente.focus();">
	</td></tr>

<!--  CONFERENTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">CONFERENTE</span>
	<br>
		<input class="PLLe" maxlength="30" style="width:220px;" name="c_conferente" id="c_conferente" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_motorista.focus();">
	</td></tr>

<!--  MOTORISTA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">MOTORISTA</span>
	<br>
		<input class="PLLe" maxlength="30" style="width:220px;" name="c_motorista" id="c_motorista" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_placa_veiculo.focus();">
	</td></tr>

<!--  PLACA DO VEÍCULO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">PLACA DO VEÍCULO</span>
	<br>
		<input class="PLLe" maxlength="8" style="width:120px;" name="c_placa_veiculo" id="c_placa_veiculo" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_pedidos.focus();">
	</td></tr>

<!--  PEDIDOS  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">PEDIDO(S)</span>
	<br><center>
		<textarea rows="8" name="c_pedidos" id="c_pedidos" class="PLBe" style="font-size:9pt;width:110px;margin-bottom:4px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"><%=s_campo_pedidos%></textarea>
	</center>
	</td></tr>

</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="confirma a operação">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>
