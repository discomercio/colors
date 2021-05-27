<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================================
'	  PedidoEntregaUsandoRomaneioConsiste.asp
'     =====================================================
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

	class clPedidoEntregaUsandoRomaneioConsiste
		dim pedido
		dim blnErroConsistencia
		end class
	
	dim s, usuario, c_nsu_romaneio, s_filtro, strDisabled, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	c_nsu_romaneio = retorna_so_digitos(Trim(Request.Form("c_nsu_romaneio")))
	
	dim c_nfe_emitente
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
	
	dim v_pedido, i, j, intIndex, intQtdeTotalPedidos, intQtdePedidosOk
	intQtdeTotalPedidos = 0
	intQtdePedidosOk = 0
	redim v_pedido(0)
	set v_pedido(UBound(v_pedido)) = new clPedidoEntregaUsandoRomaneioConsiste
	v_pedido(Ubound(v_pedido)).pedido = ""

	dim blnErroFatal, strMsgAlertaPedido
	blnErroFatal = False
	strMsgAlertaPedido = ""
	
	dim alerta
	alerta = ""

	dim observacoes
	observacoes = ""
	
	if alerta = "" then
		if c_nsu_romaneio = "" then
			alerta = "Informe o NSU do romaneio"
			end if
		end if
	
	if alerta = "" then
		if c_nfe_emitente = "" then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Não foi informado o CD"
		elseif converte_numero(c_nfe_emitente) = 0 then
			alerta=texto_add_br(alerta)
			alerta = alerta & "É necessário definir um CD válido"
			end if
		end if


	dim rNfeEmitente
	set rNfeEmitente = le_nfe_emitente(c_nfe_emitente)
	
	if alerta = "" then
		s = "SELECT " & _
				" tN2.pedido," & _
				" tP.numero_loja," & _
				" tP.st_entrega," & _
				" tP.st_etg_imediata," & _
				" tP.id_nfe_emitente," & _
				" t_PEDIDO__BASE.analise_credito," & _
				" t_PEDIDO__BASE.PagtoAntecipadoStatus," & _
				" tP.PagtoAntecipadoQuitadoStatus" & _
			" FROM t_WMS_ROMANEIO_N1 tN1" & _
				" INNER JOIN t_WMS_ROMANEIO_N2 tN2 ON (tN2.id_wms_romaneio_n1 = tN1.id)" & _
				" INNER JOIN t_PEDIDO tP ON (tN2.pedido = tP.pedido)" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (tP.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" WHERE" & _
				" (tN1.id = " & c_nsu_romaneio & ")" & _
			" ORDER BY" & _
				" tP.data_hora," & _
				" tP.pedido"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Romaneio com NSU=" & c_nsu_romaneio & " não está cadastrado."
		else
			do while Not rs.Eof
				if Trim(v_pedido(Ubound(v_pedido)).pedido) <> "" then
					redim preserve v_pedido(Ubound(v_pedido)+1)
					set v_pedido(UBound(v_pedido)) = new clPedidoEntregaUsandoRomaneioConsiste
					end if
				
				intQtdeTotalPedidos = intQtdeTotalPedidos + 1
				v_pedido(Ubound(v_pedido)).pedido = Trim("" & rs("pedido"))
				v_pedido(UBound(v_pedido)).blnErroConsistencia = False
				
				if Trim("" & rs("st_entrega")) <> ST_ENTREGA_A_ENTREGAR then
					blnErroFatal = True
					v_pedido(UBound(v_pedido)).blnErroConsistencia = True
					strMsgAlertaPedido=texto_add_br(strMsgAlertaPedido)
					strMsgAlertaPedido=strMsgAlertaPedido & "Pedido " & Trim("" & rs("pedido")) & " possui status inválido para esta operação: " & Ucase(x_status_entrega(Trim("" & rs("st_entrega"))))
					end if
				if CInt(rs("analise_credito")) <> CInt(COD_AN_CREDITO_OK) then
					blnErroFatal = True
					v_pedido(UBound(v_pedido)).blnErroConsistencia = True
					strMsgAlertaPedido=texto_add_br(strMsgAlertaPedido)
					strMsgAlertaPedido=strMsgAlertaPedido & "Pedido " & Trim("" & rs("pedido")) & " possui análise de crédito em situação inválida: " & Ucase(descricao_analise_credito(rs("analise_credito")))
					end if
				if CInt(rs("st_etg_imediata")) <> CInt(COD_ETG_IMEDIATA_SIM) then
					blnErroFatal = True
					v_pedido(UBound(v_pedido)).blnErroConsistencia = True
					strMsgAlertaPedido=texto_add_br(strMsgAlertaPedido)
					strMsgAlertaPedido=strMsgAlertaPedido & "Pedido " & Trim("" & rs("pedido")) & " está com 'Entrega Imediata' assinalada com 'Não'"
					end if
				if (CInt(rs("PagtoAntecipadoStatus")) = CInt(COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO)) And (CInt(rs("PagtoAntecipadoQuitadoStatus")) <> CInt(COD_PAGTO_ANTECIPADO_QUITADO_STATUS_QUITADO)) then
					blnErroFatal = True
					v_pedido(UBound(v_pedido)).blnErroConsistencia = True
					strMsgAlertaPedido=texto_add_br(strMsgAlertaPedido)
					strMsgAlertaPedido=strMsgAlertaPedido & "Pedido " & Trim("" & rs("pedido")) & " está com status de pagamento antecipado '" & pagto_antecipado_quitado_descricao(rs("PagtoAntecipadoStatus"), rs("PagtoAntecipadoQuitadoStatus")) & "'"
					end if

			'	VERIFICA SE O CD DO USUÁRIO ESTÁ COERENTE COM O PEDIDO
				if CLng(rNfeEmitente.id) <> CLng(rs("id_nfe_emitente")) then
				'	ERRO: PEDIDO PERTENCE A OUTRO CD
					blnErroFatal = True
					v_pedido(UBound(v_pedido)).blnErroConsistencia = True
					strMsgAlertaPedido=texto_add_br(strMsgAlertaPedido)
					strMsgAlertaPedido=strMsgAlertaPedido & "Pedido " & Trim("" & rs("pedido")) & " pertence a outro CD"
					end if

				if Not v_pedido(UBound(v_pedido)).blnErroConsistencia then intQtdePedidosOk = intQtdePedidosOk + 1
				rs.MoveNext
				loop
			end if
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
	<title>CENTRAL</title>
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var qtde_pedidos=<%=Cstr(intQtdeTotalPedidos)%>;

function marcar_todos() {
	$(".ckbPedido:enabled").prop("checked", true);
}

function desmarcar_todos() {
	$(".ckbPedido:enabled").prop("checked", false);
}

function fETGConfirma(f) {
	var qtdePedidosSelecionados = 0;

	for (i = 0; i < qtde_pedidos; i++) {
		if ($("#ckb_pedido_" + (i + 1).toString()).is(":checked")) qtdePedidosSelecionados++;
	}

	if (qtdePedidosSelecionados == 0) {
		alert("Nenhum pedido foi selecionado!!");
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style type="text/css">
.spnPedido
{
	cursor:default;
	font-size:9pt;
}
.tdPedido
{
	width: 130px;
}
</style>

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:649px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ***************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR DADOS PARA CONFIRMAÇÃO  ********** -->
<!-- ***************************************************************** -->
<body onload="focus();">
<center>

<form id="fETG" name="fETG" method="post" action="PedidoEntregaUsandoRomaneioConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_nsu_romaneio" id="c_nsu_romaneio" value="<%=c_nsu_romaneio%>">
<input type="hidden" name="c_qtde_pedidos" id="c_qtde_pedidos" value="<%=Cstr(intQtdeTotalPedidos)%>">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Entrega de Mercadorias</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellpadding='1' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	s = c_nsu_romaneio
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><span class='N'>NSU do Romaneio:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>
<br>


<!-- ************   HÁ OBSERVAÇÕES?  ************ -->
<% if observacoes <> "" then %>
		<br>
		<span class="Lbl">OBSERVAÇÕES</span>
		<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'><%=observacoes%></p></div>
		<br><br>
<% end if %>

<!-- ************   HÁ MENSAGEM DE ALERTA?  ************ -->
<% if strMsgAlertaPedido <> "" then %>
		<br>
		<span class="Lbl">ALERTAS</span>
		<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'><%=strMsgAlertaPedido%></p></div>
		<br><br>
<% end if %>

<!--  PEDIDOS  -->
<table class="Qx" cellspacing="0">
	<tr bgColor="#FFFFFF">
		<td class="MT tdPedido" align="left" nowrap style='background:azure;'><span class="PLTe">Nº Pedido(s)&nbsp;</span></td>
	</tr>
<%	intIndex = 0
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		if v_pedido(i).pedido <> "" then
			intIndex = intIndex + 1
			if v_pedido(i).blnErroConsistencia then
				strDisabled = " disabled"
			else
				strDisabled = ""
				end if %>
			<tr>
				<td class="MDBE" align="left" valign="baseline" nowrap>
					<input type="checkbox" class="CBOX ckbPedido" name="ckb_pedido_<%=Cstr(intIndex)%>" id="ckb_pedido_<%=Cstr(intIndex)%>" value="<%=v_pedido(i).pedido%>" <%=strDisabled%> />
					<span class="C spnPedido" name="spn_pedido_<%=Cstr(intIndex)%>" id="spn_pedido_<%=Cstr(intIndex)%>" onclick="fETG.ckb_pedido_<%=Cstr(intIndex)%>.click();" <%=strDisabled%>><%=v_pedido(i).pedido%></span>
				</td>
			</tr>
<%			end if
		next		%>
</table>
<br />

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">
	<input name="bMarcarTodos" id="bMarcarTodos" type="button" class="Button" onclick="marcar_todos();" value="Marcar todos" title="assinala todas as operações" style="margin-left:6px;margin-bottom:2px">
	<input name="bDesmarcarTodos" id="bDesmarcarTodos" type="button" class="Button" onclick="desmarcar_todos();" value="Desmarcar todos" title="desmarca todas as operações" style="margin-left:6px;margin-right:6px;margin-bottom:2px">
	</td>
</tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<% if intQtdePedidosOk = 0 then %>
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
</tr>
<% else %>
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
	<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fETGConfirma(fETG)" title="confirma a entrega de mercadorias">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% end if %>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>