<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===================================
'	  OrcamentoVirarPedidoSelManualCD.asp
'     ===================================
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

	dim usuario, loja, orcamento_selecionado, msg_erro
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	orcamento_selecionado = Trim(request("orcamento_selecionado"))
	if (orcamento_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_NAO_ESPECIFICADO)
	
	dim alerta
	alerta = ""

	dim c_ExibirCamposModoSelecaoCD, intIdx
	c_ExibirCamposModoSelecaoCD = Trim(Request("c_ExibirCamposModoSelecaoCD"))

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

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
	<title>LOJA</title>
	</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var loja="<%=loja%>";

function fORCConfirma(f) {
	if (f.c_ExibirCamposModoSelecaoCD.value == "S") {
		if ((!f.rb_selecao_cd[0].checked) && (!f.rb_selecao_cd[1].checked)) {
			strMsgErro = "É necessário informar o modo de seleção do CD (auto-split)!";
			alert(strMsgErro);
			return;
		}

		if (f.rb_selecao_cd[1].checked) {
			if (trim(f.c_id_nfe_emitente_selecao_manual.value) == "") {
				strMsgErro = "É necessário selecionar o CD que irá atender o pedido (sem auto-split)!";
				alert(strMsgErro);
				f.c_id_nfe_emitente_selecao_manual.focus();
				return;
			}
		}
	}

	dCONFIRMA.style.visibility = "hidden";
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
#rb_indicacao {
	margin: 0pt 2pt 1pt 10pt;
	}
</style>

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body>
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>


<% else %>
<body>
<center>

<form id="fORC" name="fORC" method="post" action="OrcamentoVirarPedido.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="orcamento_selecionado" id="orcamento_selecionado" value='<%=orcamento_selecionado%>'>
<input type="hidden" name="c_ExibirCamposModoSelecaoCD" id="c_ExibirCamposModoSelecaoCD" value='<%=c_ExibirCamposModoSelecaoCD%>'>


<!--  I D E N T I F I C A Ç Ã O   D O   O R Ç A M E N T O -->
<br />
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Transformar Pré-Pedido em Pedido<br />Nº <%=orcamento_selecionado%></span></td>
</tr>
</table>
<br>


<!-- ************   SELEÇÃO MANUAL DO CD   ************ -->
<% if operacao_permitida(OP_LJA_CADASTRA_NOVO_PEDIDO_SELECAO_MANUAL_CD, s_lista_operacoes_permitidas) And (c_ExibirCamposModoSelecaoCD = "S") then %>
<br />
<table class="Q" style="width:375px;" cellspacing="0">
  <tr>
	<td align="left">
	  <p class="Rf">Modo de Seleção do CD (Auto-Split)</p>
	</td>
  </tr>
  <tr>
	<td align="left">
	  <table width="100%" cellspacing="0" cellpadding="4" border="0">
		<!--  AUTOMÁTICO  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left" valign="baseline">
				  <% intIdx = 0 %>
				  <input type="radio" name="rb_selecao_cd" id="rb_selecao_cd_auto" value="<%=MODO_SELECAO_CD__AUTOMATICO%>"><span class="C" style="cursor:default" onclick="fORC.rb_selecao_cd[<%=Cstr(intIdx)%>].click();">Automático</span>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<tr>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td colspan="2" align="left" valign="baseline">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" name="rb_selecao_cd" id="rb_selecao_cd_manual" value="<%=MODO_SELECAO_CD__MANUAL%>"><span class="C" style="cursor:default" onclick="fORC.rb_selecao_cd[<%=Cstr(intIdx)%>].click();">Manual</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:40px;" align="left">&nbsp;</td>
				<td align="left">
				  <select id="c_id_nfe_emitente_selecao_manual" name="c_id_nfe_emitente_selecao_manual" onclick="fORC.rb_selecao_cd[<%=Cstr(intIdx)%>].click();">
					<% =wms_apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
				  </select>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
</table>
<% else %>
<input type="hidden" name="rb_selecao_cd" value="<%=MODO_SELECAO_CD__AUTOMATICO%>" />
<% end if %>


<!-- ************   SEPARADOR   ************ -->
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="749" cellspacing="0">
<tr>
	<td align="left"><a name="bCANCELA" id="bCANCELA" href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fORCConfirma(fORC)" title="vai para a página seguinte">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
<% end if %>

</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>