<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RomaneioPreFiltro.asp
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

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

'	CD
	dim i, qtde_nfe_emitente
	dim v_usuario_x_nfe_emitente
	dim id_nfe_emitente_selecionado
	v_usuario_x_nfe_emitente = obtem_lista_usuario_x_nfe_emitente(usuario)

	qtde_nfe_emitente = 0
	for i=Lbound(v_usuario_x_nfe_emitente) to UBound(v_usuario_x_nfe_emitente)
		if Not Isnull(v_usuario_x_nfe_emitente(i)) then
			qtde_nfe_emitente = qtde_nfe_emitente + 1
			id_nfe_emitente_selecionado = v_usuario_x_nfe_emitente(i)
			end if
		next
	
	if qtde_nfe_emitente > 1 then
	'	HÁ MAIS DO QUE 1 CD, ENTÃO SERÁ EXIBIDA A LISTA P/ O USUÁRIO SELECIONAR UM CD
		id_nfe_emitente_selecionado = 0
		end if
	
	if qtde_nfe_emitente = 0 then
	'	NÃO HÁ NENHUM CD CADASTRADO P/ ESTE USUÁRIO!!
		Response.Redirect("aviso.asp?id=" & ERR_NENHUM_CD_HABILITADO_PARA_USUARIO)
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
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
		$("#c_dt_entrega").hUtilUI('datepicker_padrao');
	});
</script>

<script language="JavaScript" type="text/javascript">
function limpaCampoTransp(f) {
	f.c_transportadora.options[0].selected = true;
}

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

	if (trim(f.c_nfe_emitente.value) == "") {
		alert("É necessário selecionar um CD!!");
		return;
	}

	if (converte_numero(f.c_nfe_emitente.value) == 0) {
		alert("CD selecionado é inválido!!");
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">


<body onload="fFILTRO.c_transportadora.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RomaneioPreFiltroConsiste.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<% if qtde_nfe_emitente = 1 then %>
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=Cstr(id_nfe_emitente_selecionado)%>" />
<% end if %>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Romaneio de Entrega</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellspacing="0" cellpadding="2">
<!--  TRANSPORTADORA  -->
	<tr>
		<td class="ME MD MC PLTe" nowrap align="left" valign="bottom">&nbsp;TRANSPORTADORA</td>
	</tr>
	<tr bgcolor="#FFFFFF" nowrap>
		<td class="ME MB MD" align="left">
			<table cellspacing="0" cellpadding="0" style="margin:1px 10px 6px 10px;">
			<tr>
				<td align="left" valign="middle">
					<select id="c_transportadora" name="c_transportadora" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entrega.focus();">
					<% =transportadora_monta_itens_select(Null) %>
					</select>
				</td>
				<td style="width:10px;"></td>
				<td align="left" valign="middle">
					<a name="bLimparTransp" id="bLimparTransp" href="javascript:limpaCampoTransp(fFILTRO)" title="limpa o filtro 'Transportadora'">
								<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  DATA DE COLETA (RÓTULO ANTIGO: DATA DA ENTREGA)  -->
	<tr>
		<td class="ME MD PLTe" nowrap align="left" valign="bottom">&nbsp;DATA DE COLETA</td>
	</tr>
	<tr bgcolor="#FFFFFF" nowrap>
	<td class="MDBE" align="left" valign="bottom" nowrap>
		<input class="Cc" maxlength="10" style="width:100px;margin:1px 3px 3px 10px;" name="c_dt_entrega" id="c_dt_entrega" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_data();">
	</td>
	</tr>

<% if qtde_nfe_emitente > 1 then %>
<tr>
	<td class="MB ME MD" align="left">
	<table class="Qx" cellspacing="0" cellpadding="0">
	<tr bgcolor="#FFFFFF">
		<td align="left" nowrap>
			<span class="PLTe">CD</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="left">
			<table style="margin: 4px 8px 4px 8px;" cellspacing="0" cellpadding="0">
				<tr bgcolor="#FFFFFF">
				<td align="left">
					<select id="c_nfe_emitente" name="c_nfe_emitente" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" style="margin-left:5px;margin-top:4pt; margin-bottom:4pt;">
						<%=wms_usuario_x_nfe_emitente_monta_itens_select(usuario, "")%>
					</select>
				</td>
				</tr>
			</table>
		</td>
	</tr>
	</table>
	</td>
</tr>
<% end if %>

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
</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
