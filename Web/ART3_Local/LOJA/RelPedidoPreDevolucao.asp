<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================================
'	  RelPedidoPreDevolucao.asp
'     =======================================================
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

	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
    If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if
	
	dim s, s_sessionToken
	s_sessionToken = ""
	s = "SELECT Convert(varchar(36), SessionTokenModuloLoja) AS SessionTokenModuloLoja FROM t_USUARIO WHERE (usuario = '" & usuario & "')"
    if rs.State <> 0 then rs.Close
    rs.Open s, cn
	if Not rs.Eof then s_sessionToken = Trim("" & rs("SessionTokenModuloLoja"))
	if rs.State <> 0 then rs.Close

	dim intIdx
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
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	$(function () {
		$("#c_dt_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_termino").hUtilUI('datepicker_peq_filtro_final');

		$("#divAjaxRunning").hide();

		$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8

		$(document).ajaxStart(function () {
			$("#divAjaxRunning").show();
		})
		.ajaxStop(function () {
			$("#divAjaxRunning").hide();
		});

		//Every resize of window
		$(window).resize(function () {
			sizeDivAjaxRunning();
		});

		//Every scroll of window
		$(window).scroll(function () {
			sizeDivAjaxRunning();
		});
	});

//Dynamically assign height
	function sizeDivAjaxRunning() {
		var newTop = $(window).scrollTop() + "px";
		$("#divAjaxRunning").css("top", newTop);
	}

	function geraArquivoXLS(f) {
		var status, data_inicio, data_termino, loja;
		var usuario, sessionToken;
		var serverVariableUrl, strUrl, strUrlDownload;

		if (!consisteCamposFiltro(f)) return;

		status = $("input[name='rb_status']:checked").val();
		data_inicio = trim($("#c_dt_inicio").val());
		data_termino = trim($("#c_dt_termino").val());
		loja = trim($("#c_loja").val());

		usuario = "<%=usuario%>";
		sessionToken = $("#sessionToken").val();

		serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
		serverVariableUrl = serverVariableUrl.toUpperCase();
		serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("LOJA"));

		strUrl = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/RelPedidoPreDevolucaoXLS/';
		strUrlDownload = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/DownloadRelPedidoPreDevolucaoXLS/';

		$("#divAjaxRunning").show();
		var jqxhr = $.ajax({
			url: strUrl,
			type: "GET",
			async: true,
			dataType: "json",
			data: {
				usuario: usuario,
				loja: loja,
				sessionToken: sessionToken,
				filtro_status: status,
				filtro_data_inicio: data_inicio,
				filtro_data_termino: data_termino,
				filtro_lojas: loja
			}
		})
		.success(function (response) {
			$("#divAjaxRunning").hide();
			if (response.Status == "OK") {
				fDOWNLOAD.action = strUrlDownload + "?fileName=" + response.fileName;
				fDOWNLOAD.submit();
			}
			else if (response.Status == "Falha") {
				alert("Falha ao gerar o arquivo XLS!\n" + response.Exception);
			}
			else if (response.Status == "Vazio") {
				alert("Nenhum registro encontrado!");
			}
		})
		.fail(function (jqXHR, textStatus) {
			$("#divAjaxRunning").hide();
			var msgErro = "";
			if (textStatus.toString().length > 0) msgErro = "Mensagem de Status: " + textStatus.toString();
			try {
				if (jqXHR.status.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Status: " + jqXHR.status.toString(); }
			} catch (e) { }

			try {
				if (jqXHR.statusText.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Descrição do Status: " + jqXHR.statusText.toString(); }
			} catch (e) { }

			try {
				if (jqXHR.responseText.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Mensagem de Resposta: " + jqXHR.responseText.toString(); }
			} catch (e) { }

			alert("Falha ao tentar processar a consulta!!\n\n" + msgErro);
		});
	}

	function fFILTROConfirma(f) {
		if (!consisteCamposFiltro(f)) return;

		if (f.rb_saida[1].checked) {
			geraArquivoXLS(f);
		}
		else {
			dCONFIRMA.style.visibility = "hidden";
			window.status = "Aguarde ...";
			f.submit();
		}
	}

	function consisteCamposFiltro(f) {
		var i, blnFlag, s_status_selecionado;
		var s_de, s_ate;

		if (trim(f.c_dt_inicio.value) != "") {
			if (!isDate(f.c_dt_inicio)) {
				alert("Data inválida!!");
				f.c_dt_inicio.focus();
				return false;
			}
		}

		if (trim(f.c_dt_termino.value) != "") {
			if (!isDate(f.c_dt_termino)) {
				alert("Data inválida!!");
				f.c_dt_termino.focus();
				return false;
			}
		}

		s_de = trim(f.c_dt_inicio.value);
		s_ate = trim(f.c_dt_termino.value);
		if ((s_de != "") && (s_ate != "")) {
			s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
			s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
			if (s_de > s_ate) {
				alert("Data de término é menor que a data de início!!");
				f.c_dt_termino.focus();
				return false;
			}
		}

		s_status_selecionado = "";
		blnFlag = false;
		for (i = 0; i < f.rb_status.length; i++) {
			if (f.rb_status[i].checked) {
				blnFlag = true;
				s_status_selecionado = f.rb_status[i].value;
				break;
			}
		}
		if (!blnFlag) {
			alert("Selecione o status da pré-devolução!!");
			return false;
		}

		if ((s_status_selecionado == "FINALIZADA") || (s_status_selecionado == "REPROVADA")) {
			if ((trim(f.c_dt_inicio.value) == "") || (trim(f.c_dt_termino.value) == "")) {
				alert("É necessário informar o período da consulta para o status selecionado!!");
				return false;
			}
		}

		return true;
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
#divAjaxRunning
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	z-index:1001;
	background-color:grey;
	opacity: .6;
}
.AjaxImgLoader
{
	position: absolute;
	left: 50%;
	top: 50%;
	margin-left: -128px; /* -1 * image width / 2 */
	margin-top: -128px;  /* -1 * image height / 2 */
	display: block;
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


<body onload="focus();">

<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>

<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelPedidoPreDevolucaoExec.asp">
<input type="hidden" name="sessionToken" id="sessionToken" value="<%=s_sessionToken%>" />
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_loja" id="c_loja" value="<%=loja%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pré-Devolução</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0" style="width:300px;">
<!--  PERÍODO  -->
	<tr>
		<td class="MT PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;PERÍODO</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD" style="padding:5px;" nowrap>
			<table cellspacing="2" cellPadding="0"><tr bgColor="#FFFFFF"><td>
				<input class="PLLc" maxlength="10" style="width:80px;" name="c_dt_inicio" id="c_dt_inicio" onfocus="this.select();" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_termino.focus(); filtra_data();"
						/>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;&nbsp;&nbsp;até&nbsp;</span>&nbsp;<input class="PLLc" maxlength="10" style="width:80px; " name="c_dt_termino" id="c_dt_termino" onfocus="this.select();" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.focus();}" onkeypress="filtra_data();"
						/>
				</td></tr>
			</table>
		</td>
	</tr>
<!--  STATUS DA PRÉ-DEVOLUÇÃO  -->
	<tr>
		<td class="ME MB MD PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;STATUS DA PRÉ-DEVOLUÇÃO</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD" style="padding:5px;">
			<% intIdx=-1 %>
			<input type="radio" id="rb_status" name="rb_status" value="CADASTRADA" class="CBOX" style="margin-left:20px;" checked>
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Cadastrada</span>
			<br>
			<input type="radio" id="rb_status" name="rb_status" value="EM_ANDAMENTO" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Em Andamento</span>
			<br>
			<input type="radio" id="rb_status" name="rb_status" value="MERCADORIA_RECEBIDA" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Mercadoria Recebida</span>
            <br>
            <input type="radio" id="rb_status" name="rb_status" value="FINALIZADA" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Finalizada</span>
            <br>
			<input type="radio" id="rb_status" name="rb_status" value="REPROVADA" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Reprovada</span>
		</td>
	</tr>
<!--  SAÍDA DO RELATÓRIO  -->
	<tr>
		<td class="ME MB MD PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;SAÍDA</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD" style="padding:5px;">
		<input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="Html" checked><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_saida[0].click();"
			>Html</span>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="XLS"><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_saida[1].click();"
			>Excel</span>
		</td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>

<form method="POST" name="fDOWNLOAD" id="fDOWNLOAD">
<input type="hidden" name="usuario" value="<%=usuario%>" />
<input type="hidden" name="loja" value="<%=loja%>" />
<input type="hidden" name="sessionToken" value="<%=s_sessionToken%>" />
<input type="hidden" name="fileName" />
</form>

</body>
</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
