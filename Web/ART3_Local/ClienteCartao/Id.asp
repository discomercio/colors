<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->

<%
'     =========================
'	  Id.asp
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

'	LEMBRANDO QUE NAS PÁGINAS ACESSADAS DIRETAMENTE PELOS CLIENTES PARA FAZER O PAGAMENTO NÃO SE DEVE USAR 'SESSION',
'	JÁ QUE OCORRERAM VÁRIOS CASOS DE CLIENTES QUE NÃO CONSEGUIRAM INICIAR A SESSÃO (PROVAVELMENTE POR PROBLEMAS NA CONFIGURAÇÃO DE COOKIES OU MESMO DEVIDO A ANTIVIRUS)
	Session.Contents.RemoveAll
	Session.Abandon

	dim strScript
	dim strProtocolo, strServerName, strLocalAddr, strURL
	strServerName = Ucase(Trim(Request.ServerVariables("server_name")))
	strLocalAddr = Trim(Request.ServerVariables("local_addr"))
	strURL = request.ServerVariables("URL")
	strProtocolo = "https"
	if Not SITE_CLIENTE_USAR_PROTOCOLO_HTTPS then strProtocolo = "http"

'	LEMBRANDO QUE O ENDEREÇO 'PAGAMENTO.BONSHOP.COM.BR' É O SITE ALTERNATIVO LOCALIZADO NO SERVIDOR DO SISTEMA E QUE NÃO POSSUI CERTIFICADO SSL INSTALADO
	if (strServerName = "XXXDNN2581") Or (strServerName = "LOCALHOST") Or (strServerName = "PAGAMENTO.BONSHOP.COM.BR") then strProtocolo = "http"
	if (strServerName = "HOMOLOGACAO.CENTRAL85.COM.BR") then strProtocolo = "http"
	if (strServerName = "APPSERVER") Or (strServerName = "WIN2008R2") Or (strServerName = "WIN2008R2BS") then strProtocolo = "http"

	if Instr(Ucase(strURL), "HOMOLOG") > 0 then strProtocolo = "http"
	if Instr(Ucase(strURL), "TEST") > 0 then strProtocolo = "http"

	if strLocalAddr <> "" then
		if Instr(strLocalAddr, "182.168.0.") > 0 then strProtocolo = "http"
		if Instr(strLocalAddr, "182.168.2.") > 0 then strProtocolo = "http"
		if Instr(strLocalAddr, "192.168.0.") > 0 then strProtocolo = "http"
		if Instr(strLocalAddr, "192.168.2.") > 0 then strProtocolo = "http"
		if Instr(strLocalAddr, "127.0.0.") > 0 then strProtocolo = "http"
		if strLocalAddr = "::1" then strProtocolo = "http"
		end if
	
	dim vURL
	dim strNomePagina, strFolderL1, strFolderL2
	strNomePagina = ""
	strFolderL1 = ""
	strFolderL2 = ""
	if Trim(strURL) = "" then
		redim vURL(0)
	else
		vURL = Split(strURL, "/")
		if UBound(vURL) >= 2 then
			strNomePagina = Trim(vURL(UBound(vURL)))
			strFolderL2 = Trim(vURL(UBound(vURL)-1))
			strFolderL1 = Trim(vURL(UBound(vURL)-2))
			end if
		end if
	
	strScript = "<script language='JavaScript'>" & chr(13) & _
				"	var protocolo_selecionado = '" & strProtocolo & "';" & chr(13) & _
				"</script>" & chr(13)
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
	<title><%=SITE_CLIENTE_TITULO_JANELA%></title>
	</head>

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_GLOBAL%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__SSL_JS%>" language="JavaScript" type="text/javascript"></script>

<% =strScript %>

<script language="JavaScript" type="text/javascript">
var blnVersaoNavegadorOk = true;
var cnpj_cpf_digitado_anterior = "";

function posiciona_foco( f ){
	if (trim(f.cnpj_cpf_selecionado.value)==""){ 
		f.cnpj_cpf_selecionado.focus();
		return true;
		}
	if (trim(f.pedido_selecionado.value)==""){ 
		f.pedido_selecionado.focus();
		return true;
		}
	return false;
}

function confere(f) {

	if (!blnVersaoNavegadorOk) {
		alert("Esta versão do navegador não é suportada!!\nPor favor, utilize o Internet Explorer versão 7 ou superior!!");
		return false;
	}

	f.cnpj_cpf_selecionado.value=trim(f.cnpj_cpf_selecionado.value);
	f.pedido_selecionado.value=trim(f.pedido_selecionado.value);
	
	if (f.cnpj_cpf_selecionado.value=="") {
		alert("Preencha o CNPJ/CPF!!");
		f.cnpj_cpf_selecionado.focus();
		return false;
		}

	if (!cnpj_cpf_ok(f.cnpj_cpf_selecionado.value)) {
		alert("CNPJ/CPF digitado é inválido!!");
		f.cnpj_cpf_selecionado.focus();
		return false;
		}

	if (f.pedido_selecionado.value=="") {
		alert("Informe o número do pedido!!");
		f.pedido_selecionado.focus();
		return false;
		}
	
	window.status = "Aguarde ...";
	return true;
}

function trataOnKeypressCnpjCpf(evento) {
var blnCancelarKeystroke = false;
	try {
		if (HHO.digitouEnter(evento)) {
			blnCancelarKeystroke = true;
			if (HHO.temInfo(this.value) && cnpj_cpf_ok(this.value)) {
				this.value = cnpj_cpf_formata(this.value);
				$(this).hUtil('focusNext');
			}
		}
		if (!HHO.digitacaoCnpjCpfOk(evento)) blnCancelarKeystroke = true;
	}
	catch (e) {
		alert(e.message);
	}
	finally {
		if (blnCancelarKeystroke) {
			evento.preventDefault();
			return false;
		}
		else {
			return true;
		}
	}
}

function trataOnKeypressPedido(evento) {
var blnCancelarKeystroke = false;
	try {
		if (HHO.digitouEnter(evento)) {
			blnCancelarKeystroke = true;
			if (HHO.temInfo(this.value)) {
				if (normaliza_num_pedido(this.value) != '') this.value = normaliza_num_pedido(this.value);
				$("#bCONSULTAR").click();
			}
		}
		if (!HHO.digitacaoNumPedidoOk(evento)) blnCancelarKeystroke = true;
	}
	catch (e) {
		alert(e.message);
	}
	finally {
		if (blnCancelarKeystroke) {
			evento.preventDefault();
			return false;
		}
		else {
			return true;
		}
	}
}

function trataOnBlurCnpjCpf() {
	if (!cnpj_cpf_ok(this.value)) {
		if (retorna_so_digitos(this.value) != retorna_so_digitos(cnpj_cpf_digitado_anterior)) {
			cnpj_cpf_digitado_anterior = this.value;
			alert('CNPJ/CPF inválido!!');
			this.focus();
		}
	} else {
		this.value = cnpj_cpf_formata(this.value);
	}
}

function trataOnBlurPedido() {
	if (normaliza_num_pedido(this.value) != '') this.value = normaliza_num_pedido(this.value);
}

function isPedidoOLD01(pedido) {
	var c;

	// Evita que se acesse a página destinada ao ambiente antigo e ocorra exception devido ao BD inválido
	return false;

	if (pedido == null) return false;
	if (pedido.length == 0) return false;
	for (var i = 0; i < pedido.length; i++) {
		c = pedido[i];
		if (!isDigit(c)) {
			if (c === "M") {
				return true;
			} else {
				return false;
			}
		}
	}
}

function endsWith(texto, sufixo) {
	return texto.indexOf(sufixo, texto.length - sufixo.length) !== -1;
}

function setFormAction(f) {
	var folderL1 = "<%=strFolderL1%>";
	var folderL1Suffix = "OLD01";
	if (isPedidoOLD01(f.pedido_selecionado.value)) {
		if ((folderL1.toUpperCase().indexOf("ART3") == -1) && (!endsWith(folderL1, folderL1Suffix))) folderL1 += folderL1Suffix;
		f.action = "../../" + folderL1 + "/<%=strFolderL2%>/PedidoConsulta.asp";
	} else {
		if (endsWith(folderL1, folderL1Suffix)) folderL1 = folderL1.substr(0, folderL1.length - folderL1Suffix.length);
		f.action = "../../" + folderL1 + "/<%=strFolderL2%>/PedidoConsulta.asp";
	}
}
</script>

<script language="JavaScript" type="text/javascript">
	var urlAtual, urlDestino;
	$(document).ready(function() {
		if ((window.location.protocol == "http:") && (protocolo_selecionado == "https")) {
			urlAtual = window.location.href;
			urlDestino = urlAtual.replace("http:", "https:");
			$("#divMsg").html("O navegador será automaticamente redirecionado para um ambiente seguro em instantes...<br />Ou clique <a href='" + urlDestino + "'>aqui</a> para ser redirecionado agora.");
			$("#divMsg").show();
			window.location = urlDestino;
			return;
		}
		$('#cnpj_cpf_selecionado').keypress(trataOnKeypressCnpjCpf);
		$('#cnpj_cpf_selecionado').blur(trataOnBlurCnpjCpf);
		$('#pedido_selecionado').keypress(trataOnKeypressPedido);
		$('#pedido_selecionado').blur(trataOnBlurPedido);
	});
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
<link href="<%=URL_FILE__E_LOGO_TOP_BS_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
* {
	margin:0;
}
html, body
{
	height:100%;
	overflow-y:hidden;
}
body::before
{
	content: '';
	border: none;
	margin-top: 0px;
	margin-bottom: 0px;
	padding: 0px;
}
.wrapper
{
	min-height:100%;
	height:auto !important;
	height:100%;
	margin: 0 auto -11.5em;
}
.push
{
	height:1em;
}
.footer
{
	height:7em;
	background:#999;
}
.footer p
{
	color:#EEE;
	margin-left:6px;
	margin-right:6px;
}
.divTextoFooter
{
	background:#999;
}
.DivMsg
{
	border:2px double;
	padding:10px;
	font-family:Arial;
	font-weight:bold;
	font-size:10pt;
	text-align:center;
	color:Black;
	background-color:#F5F5F5;
	width:75%;
	margin-top:40px;
}
.divNavegadores
{
	text-align:right;
	vertical-align:bottom;
	height:4.5em;
}
.imgClearsale
{
	vertical-align:bottom;
}
.imgNavegadores
{
	margin-right: 0.5em;
	vertical-align:bottom;
}
</style>


<% if isHorarioManutencaoSistema then %>
<body>
<center>
<br />
<h1>Sistema em manutenção no período das <%=HORARIO_INICIO_MANUTENCAO_SISTEMA%> até <%=HORARIO_TERMINO_MANUTENCAO_SISTEMA%><br /><br />Por favor, tente mais tarde.</h1>
</center>
</body>
<% elseif isHorarioRebootServidor then %>
<body>
<center>
<br />
<h1>Sistema indisponível no período das <%=HORARIO_INICIO_REBOOT_SERVIDOR%> até <%=HORARIO_TERMINO_REBOOT_SERVIDOR%><br /><br />Por favor, tente mais tarde.</h1>
</center>
</body>
<% else %>
<body onload="window.status=''; if (!posiciona_foco(fPED)) fPED.bCONSULTAR.focus();">
<div class="wrapper">
<center>
<!--  LOGOTIPO  -->
<table class="notPrint" id="tbl_logotipo_bonshop" width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td align="center"><img alt="<%=SITE_CLIENTE_HEADER__ALT_IMG_TEXT%>" src="../imagem/<%=SITE_CLIENTE_HEADER__LOGOTIPO%>" /></td>
	</tr>
</table>
<table class="notPrint" id="pagina_tbl_cabecalho" cellspacing="0px" cellpadding="0px">
	<tbody>
		<tr style="height:78px;">
			<td id="topo_verde" colspan="3" align="left">
				<div id="moldura_do_letreiro">
					<div id="letreiro_div" style="display:block;"></div>
				</div>
				<div id="telefone"></div>
			</td>
		</tr>
		<tr>
			<td id="topo_azul" colspan="3" align="left">&nbsp;</td>
		</tr>
	</tbody>
</table>

<br />

<!--  L O G O N  -->
<form action="" method="post" id="fPED" name="fPED" onsubmit="if (!confere(fPED)) return false; setFormAction(fPED);">
<span class="T" style="color:green;">Consulta de Pedido</span>
<div class="QFn" style="width:300px; margin-top:6px;background-color:#F5F5F5;" align="center">
	<table style="margin: 15px 20px 0px 20px;" cellspacing="0" cellpadding="0">
	<tr>
		<td align="right" valign="bottom"><span class="R">CNPJ/CPF&nbsp;</span></td>
		<td align="left"><input name="cnpj_cpf_selecionado" id="cnpj_cpf_selecionado" type="text" maxlength="18" style="width:170px;padding-left:3px;letter-spacing:1px;"></td>
		</tr>
	<tr><td colspan="2" align="left"><span style="font-size:6pt;">&nbsp;</span></td></tr>
	<tr>
		<td align="right" valign="bottom"><span class="R">Nº&nbsp;PEDIDO&nbsp;</span></td>
		<td align="left"><input name="pedido_selecionado" id="pedido_selecionado" type="text" maxlength="10" style="width:170px;padding-left:3px;letter-spacing:1px;"></td>
		</tr>
	<tr><td colspan="2" align="left"><span style="font-size:8pt;">&nbsp;</span></td></tr>
	<tr><td colspan="2" align="center">
		<input name="bCONSULTAR" id="bCONSULTAR" type="submit" class="Botao"  style="background-color:#90EE90;border:double 3px #006400;"
			value="Consultar" title="consulta o pedido"></td>
		</tr>
	<tr><td colspan="2" align="left"><span style="font-size:8pt;">&nbsp;</span></td></tr>
	</table>
</div>
</form>

<div id="divMsg" class="C DivMsg" style="display:none;"></div>

<div class="push"></div>
</center>
</div>

<div class="divNavegadores"><a href="http://www.clearsale.com.br" target="_blank"><img class="imgClearsale" alt="Selo Clearsale" src="../imagem/clearsale.gif" /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img class="imgNavegadores" alt="Navegadores suportados" src="../imagem/Braspag/NavegadoresSuportados.png" /></div>

<div class="footer">
<center>
<div class="divTextoFooter" style="padding-top:2px;">
<p style="font-weight:bold;margin-top:6px;">Política de Segurança</p>
<p>Consulte o seu pedido e realize o pagamento usando cartão de crédito.</p>
<p>O cadeado de segurança será exibido no seu navegador, garantindo a autenticidade e a segurança do ambiente.</p>
<p>Os dados digitados são transmitidos protegidos por criptografia com alto nível de segurança.</p>
</div>
</center>
</div>

<% if SITE_CLIENTE_EXIBIR_LOGO_SSL then %>
<script language="JavaScript" type="text/javascript">
	logo_ssl_corner("../imagem/ssl/ssl_corner.gif");
</script>
<% end if %>

</body>
<% end if %>

</html>
