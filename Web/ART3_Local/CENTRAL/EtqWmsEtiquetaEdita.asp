<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  EtqWmsEtiquetaEdita.asp
'     =====================================
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
	dim s, usuario, c_id_wms_etq_n3
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
	
'	ETIQUETA A EDITAR
	c_id_wms_etq_n3 = retorna_so_digitos(Trim(request("c_id_wms_etq_n3")))

	if (c_id_wms_etq_n3 = "") Or (converte_numero(c_id_wms_etq_n3) = 0) then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICADOR_NAO_FORNECIDO)
	
'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	s = "SELECT " & _
			" tN1.id AS id_wms_etq_n1," & _
			" tN2.id AS id_wms_etq_n2," & _
			" tN3.id AS id_wms_etq_n3," & _
			" tN2.obs_2," & _
			" tN2.obs_3," & _
			" tN2.transportadora_id," & _
			" tCli.cnpj_cpf AS cnpj_cpf_cliente," & _
			" tCli.nome_iniciais_em_maiusculas AS nome_cliente," & _
			" tProd.descricao," & _
			" tProd.descricao_html" & _
		" FROM t_WMS_ETQ_N3_SEPARACAO_ZONA_PRODUTO tN3" & _
			" INNER JOIN t_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO tN2 ON (tN3.id_wms_etq_n2=tN2.id)" & _
			" INNER JOIN t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO tN1 ON (tN2.id_wms_etq_n1=tN1.id)" & _
			" INNER JOIN t_CLIENTE tCli ON (tN2.id_cliente = tCli.id)" & _
			" INNER JOIN t_PRODUTO tProd ON ((tN3.fabricante = tProd.fabricante) AND (tN3.produto = tProd.produto))" & _
		" WHERE" & _
			" (tN3.id = " & c_id_wms_etq_n3 & ")"
	set rs = cn.Execute(s)
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if rs.Eof then
		Response.Redirect("aviso.asp?id=" & ERR_REGISTRO_NAO_CADASTRADO)
		end if
%>


<%=DOCTYPE_LEGADO%>

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

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(document).ready(function() {
		$('#c_transportadora option[value="<%=Trim("" & rs("transportadora_id"))%>"]').attr({ selected: "selected" });
	});
</script>

<script language="JavaScript" type="text/javascript">
function AtualizaDados(f) {
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
.LblInfo
{
	color:#696969;
	font-size:9pt;
}
</style>


<body onload="fCAD.c_obs2.focus();">
<center>



<!--  DADOS DA ETIQUETA  -->

<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="center" valign="bottom"><p class="PEDIDO">Editar Dados de Etiqueta (WMS)<br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="EtqWmsEtiquetaAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_id_wms_etq_n1" id="c_id_wms_etq_n1" value="<%=Trim("" & rs("id_wms_etq_n1"))%>">
<input type="hidden" name="c_id_wms_etq_n2" id="c_id_wms_etq_n2" value="<%=Trim("" & rs("id_wms_etq_n2"))%>">
<input type="hidden" name="c_id_wms_etq_n3" id="c_id_wms_etq_n3" value="<%=c_id_wms_etq_n3%>">
<input type="hidden" name="c_obs2_original" id="c_obs2_original" value="<%=Trim("" & rs("obs_2"))%>">
<input type="hidden" name="c_obs3_original" id="c_obs3_original" value="<%=Trim("" & rs("obs_3"))%>">
<input type="hidden" name="c_transportadora_original" id="c_transportadora_original" value="<%=Trim("" & rs("transportadora_id"))%>">


<table width="400" class="Q" cellspacing="0">
<!-- ************   EXIBE DADOS P/ CONFERÊNCIA: CLIENTE / PRODUTO   ************ -->
	<tr>
		<td class="MB" align="left">
		<p class="R">Nº IDENTIFICAÇÃO ETIQUETA</p>
		<p class="C LblInfo"><%=c_id_wms_etq_n3%></p>
		</td>
	</tr>
	<tr>
		<td class="MB" align="left">
		<p class="R">CLIENTE</p>
		<p class="C LblInfo"><%=Trim("" & rs("nome_cliente"))%></p>
		</td>
	</tr>
	<tr>
		<td class="MB" align="left">
		<p class="R">PRODUTO</p>
		<p class="C LblInfo"><%=Trim("" & rs("descricao"))%></p>
		</td>
	</tr>
<!-- ************   NF (CAMPO OBS_2)   ************ -->
	<tr>
		<td class="MB" align="left">
			<p class="R">Obs II</p>
			<input id="c_obs2" name="c_obs2" class="TA" value="<%=Trim("" & rs("obs_2"))%>" maxlength="10" style="width:120px;margin-left:4px;"
				onkeypress="if (digitou_enter(true)) fCAD.c_obs3.focus(); filtra_numerico();">
		</td>
	</tr>
<!-- ************   NF (CAMPO OBS_3)   ************ -->
	<tr>
		<td align="left">
			<p class="R">Obs III</p>
			<input id="c_obs3" name="c_obs3" class="TA" value="<%=Trim("" & rs("obs_3"))%>" maxlength="10" style="width:120px;margin-left:4px;"
				onkeypress="if (digitou_enter(true)) fCAD.c_transportadora.focus(); filtra_numerico();">
		</td>
	</tr>
<!-- ************   TRANSPORTADORA   ************ -->
	<tr>
		<td class="MC" align="left">
			<p class="R">TRANSPORTADORA</p>
			<select id="c_transportadora" name="c_transportadora" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =transportadora_monta_itens_select(Null)%>
			</select>
		</td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="left"><a href="javascript:history.back()" title="voltar para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaDados(fCAD)" title="grava no banco de dados">
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