<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->

<%
'     =====================================
'	  C O N E C T A . A S P
'     =====================================
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

	dim strTarget, strTargetOutroAmb, urlOutroAmb
	dim strTituloAmbiente, strTituloOutroAmbiente
	dim idOutroAmbiente
	strTituloAmbiente = ""
	strTituloOutroAmbiente = ""
	idOutroAmbiente = 0
	if ID_PARAM_SITE = COD_SITE_ARTVEN_FABRICANTE then
		strTarget = IdFormTargetArtvenFabricante
	elseif ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then
		strTarget = IdFormTargetAssistenciaTecnica
	else
		strTarget = IdFormTargetArtvenBonshop
		if ID_AMBIENTE = ID_AMBIENTE__OLD01 then
			strTituloAmbiente = "Sistema"
			strTituloOutroAmbiente = "DIS"
		'	SE ESTÁ NA OLD01, O OUTRO AMBIENTE É DIS
			idOutroAmbiente = ID_AMBIENTE__DIS
			strTargetOutroAmb = "fArtDIS"
		'	VERIFICA SE ESTÁ EM AMBIENTE DE HOMOLOGAÇÃO
			if BRASPAG_AMBIENTE_HOMOLOGACAO then
				urlOutroAmb = "http://discomercio.com.br/homologacao"
			else
				urlOutroAmb = "http://discomercio.com.br/sistema"
				end if
		elseif ID_AMBIENTE = ID_AMBIENTE__DIS then
			strTituloAmbiente = "DIS"
			strTituloOutroAmbiente = "Sistema"
		'	SE ESTÁ NA DIS, O OUTRO AMBIENTE É OLD01
			idOutroAmbiente = ID_AMBIENTE__OLD01
			strTargetOutroAmb = "fArtOLD01"
		'	VERIFICA SE ESTÁ EM AMBIENTE DE HOMOLOGAÇÃO
			if BRASPAG_AMBIENTE_HOMOLOGACAO then
				urlOutroAmb = "http://central85.com.br/homologacao"
			else
				urlOutroAmb = "http://central85.com.br/sistema"
				end if
			end if
		end if

	strTarget = strTarget & "ORC"

'	24/02/2017 - o Carlos e o Rogério Donisete pediram p/ não exibir os nomes das empresas
	strTituloAmbiente = ""
	strTituloOutroAmbiente = ""
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
	<title>Área Restrita</title>
	</head>

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_GLOBAL%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var blnVersaoNavegadorOk = false;
if (isVersaoNavegadorOk()) blnVersaoNavegadorOk = true;

configura_painel_logon();

function limpa_campos ( f ) {
	f.usuario.value="";
	f.senha.value="";
	f.usuario.focus();
	window.status = "Concluído";
}

function posiciona_foco( f ){
	if (trim(f.usuario.value)==""){ 
		f.usuario.focus();
		return true;
		}
	if (trim(f.senha.value)==""){ 
		f.senha.focus();
		return true;
		}
}

function confere( f ){
var u, s;

	if (!blnVersaoNavegadorOk) {
		alert("Este site suporta apenas o Microsoft Internet Explorer!!\nCaso esteja usando esse navegador, por favor, tente ativar o 'Modo de Exibição de Compatibilidade' para este site!!");
		return false;
		}

	f.usuario.value=trim(f.usuario.value);
	f.senha.value=trim(f.senha.value);
	
	u = f.usuario.value;
	s = f.senha.value;
	
	if (u==""){ 
		f.usuario.focus();
		return false;
		}
	if (s==""){ 
		f.senha.focus();
		return false;
		}

	window.status = "Aguarde ...";
	return true;
}

</script>


<script language="JavaScript" type="text/javascript">
	$(document).ready(function() {
		$("div.Cielo_QFn").height($("div.BS_QFn").height());
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
<link href="<%=URL_FILE__ESCREEN_CSS%>" rel="stylesheet" type="text/css" media="screen">
<link href="<%=URL_FILE__E_LOGO_TOP_BS_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
* {
	margin:0;
}
html, body
{
	height:100%;
}
body::before {
	content: '';
	border:none;
	padding: 0px;
	margin-top:0px;
	margin-bottom:0px;
}
.wrapper
{
	min-height:100%;
	height:auto !important;
	height:100%;
	margin: 0 auto -4.0em;
}
.push
{
	height:1em;
}
.header
{
	height:3.5em;
	background:#FFFFFF;
}
.header p
{
	font-family:Arial;
	font-size:12pt;
	color:#005AAB;
	margin-left:6px;
	margin-right:6px;
}
.divTextoHeader
{
	background:#FFFFFF;
}
.footer
{
	height:4.0em;
	background:#FFFFFF;
}
.footer p
{
	font-family:Arial;
	font-size:12pt;
	color:#005AAB;
	margin-left:6px;
	margin-right:6px;
}
.divTextoFooter
{
	background:#FFFFFF;
}
.BS_QFn{
	border: 1pt solid #C0C0C0;
	margin: 0 50 0 50;
	width:500px;
}
.AT_QFn{
	background: #FFFFFF;
	border: 1pt solid #C0C0C0;
	margin: 0 50 0 50;
	background:#F1E4E4;
	width:500px;
}
.Cielo_QFn{
	background: #FFFFFF;
	border: 1pt solid #C0C0C0;
	margin: 0 50 0 50;
	background:#EAFBFB;
	width:500px;
}
.BS_Botao
{
	background: #FF7F00;
	border: double 3px #696969;
	cursor: pointer;
	text-align: center;
}
.AT_Botao
{
	background: #804040;
	border: double 3px #A9A9A9;
	cursor: pointer;
	text-align: center;
	COLOR: white;
}
.Cielo_Botao
{
	background: #94C9C9;
	border: double 3px #A9A9A9;
	cursor: pointer;
	text-align: center;
	COLOR: white;
}
.textoCielo
{
	font-family:Arial;
	font-size:10pt;
	color:#005AAB;
	margin-left:6px;
	margin-right:6px;
}
.NTit {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10pt;
	font-weight: bold;
	color: gray;
	}
<% if ID_AMBIENTE = ID_AMBIENTE__OLD01 then %>
.OA_Botao {
	background: #ABE0A6;
	border: 3px double #696969;
	cursor: pointer;
	text-align: center;
}
.OA_QFn{
	background: #FFFFFF;
	border: 1pt solid #C0C0C0;
	margin: 0 50 0 50;
	background:#EEFAED;
	width:500px;
	}
<% elseif ID_AMBIENTE = ID_AMBIENTE__DIS then %>
.OA_Botao {
	background: #FF7F00;
	border: 3px double #696969;
	cursor: pointer;
	text-align: center;
}
.OA_QFn{
	background: #FFFFFF;
	border: 1pt solid #C0C0C0;
	margin: 0 50 0 50;
	background:#FFF0E0;
	width:500px;
	}
<% end if %>
</style>


<% if isHorarioManutencaoSistema then %>
<body>
<center>
<br />
<h1>Sistema em manutenção no período das <%=HORARIO_INICIO_MANUTENCAO_SISTEMA%> até <%=HORARIO_TERMINO_MANUTENCAO_SISTEMA%><br /><br />Por favor, tente mais tarde.</h1>
</center>
</body>
<% else %>
<body onload="focus();">
<div class="wrapper">
<center>
<!--  L O G O N  -->
<table id="tbl_logotipo_bonshop" width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td align="center"><img alt="<%=SITE_PARCEIRO_HEADER__ALT_IMG_TEXT%>" src="../imagem/<%=SITE_PARCEIRO_HEADER__LOGOTIPO%>" /></td>
	</tr>
</table>
<table id="pagina_tbl_cabecalho" cellspacing="0px" cellpadding="0px">
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
<br />
<br />
<br />

<table cellpadding="0" cellspacing="0">
<tr>
	<% if ID_AMBIENTE <> ID_AMBIENTE__OLD01 then %>
	<td align="center" valign="bottom">
		<table cellpadding=0 cellspacing=0>
			<tr>
				<td align="center"><span class="NTit"><%=strTituloAmbiente%></span></td>
			</tr>
			<tr>
				<td align="center">
				<div>
				<form target="<%=strTarget%>" action="<%=URL_BASE_RELATIVA_SITE_ARTVEN3%>/orcamento/ConectaVerifica.asp" method="post" id="fID_BS" name="fID_BS" onsubmit="if (!blnVersaoNavegadorOk){alert('Esta versão do navegador não é suportada pelo site!!');return false;}">
				<div class="BS_QFn DefaultBkg" style="margin: 0 10 0 10; width:220px" align="center">
					<br />
					<p class="R" style="margin: 10 10 2 10">USUÁRIO</p>
					<input name="usuario" id="usuario" type="text" maxlength="<%=MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR%>" style="text-align:center;" onkeypress="filtra_nome_identificador(); if (digitou_enter(true) && tem_info(this.value)) fID_BS.senha.focus();">
					
					<br />
					<br />
					<p class="R" style="margin: 10 10 2 10">SENHA</p>
					<input name="senha" id="senha" type="password" maxlength="15" style="text-align:center;" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) bCONSULTAR.click();">
					
					<br />
					<p class="R" style="margin: 4 10 0 10">&nbsp;</p>
					<input name="bCONSULTAR" id="bCONSULTAR" type="button" class="Botao" 
							value="ENTRAR" title="inicia a sessão do usuário" onclick="if (confere(fID_BS)) {submit(); limpa_campos(fID_BS);}">
					<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
				</div>
				</form>
				</div>
				<div class="header">
				<div class="divTextoHeader">
				<p style="font-weight:bold;margin-top:6px;margin-bottom:6px;">Cadastre ou acompanhe<br />seu pedido</p>
				</div>
				</div>
				</td>
			</tr>
		</table>
	</td>
	<td style="width:40px" align="left">
		&nbsp;
	</td>
	<% end if %>
	<td align="center" valign="bottom" style="width:240px;">
		<table cellpadding="0" cellspacing="0" style="width:240px;">
			<tr>
				<td align="center">&nbsp;</td>
			</tr>
			<tr>
				<td align="center">
				<div>
				<form target="_blank" action="<%=URL_SITE_CLIENTE_PAGTO_CIELO%>" method="post" id="fCIELO" name="fCIELO">
				<div class="Cielo_QFn" style="margin: 0 10 0 10; width:220px;" align="center">
				<table width="100%" height="100%" cellpadding="0" cellspacing="0">
				<tr>
				<td align="center" valign="bottom">
					<br />
					<br />
					<p class="textoCielo" style="margin: 10 10 2 10;font-weight:bold;">Para pagamento de pedidos<br />no cartão de crédito</p>
					<br />
					<br />
					<p class="R" style="margin: 4 10 0 10">&nbsp;</p>
					<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Cielo_Botao" 
							value="CLIQUE AQUI" title="clique aqui para pagamento de pedidos no cartão de crédito">
					<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
				</td>
				</tr>
				</table>
				</div>
				</form>
				</div>
				<div class="header">
				<div class="divTextoHeader">
				<p style="font-weight:bold;margin-top:6px;margin-bottom:6px;">&nbsp;<br />&nbsp;</p>
				</div>
				</div>
				</td>
			</tr>
		</table>
	</td>
<!--
	<td style="width:40px;" align="left">
		&nbsp;
	</td>
	<td align="left" valign="bottom">
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td align="center">
				<div>
				<form target="<%=strTarget%>" action="<%=URL_BASE_RELATIVA_SITE_ASSISTENCIA_TECNICA%>/orcamento/ConectaVerifica.asp" method="post" id="fID_AT" name="fID_AT" onsubmit="if (!blnVersaoNavegadorOk){alert('Esta versão do navegador não é suportada pelo site!!');return false;}">
				<div class="AT_QFn" style="margin: 0 10 0 10; width:220px" align="center">
					<br />
					<p class="R" style="margin: 10 10 2 10">USUÁRIO</p>
					<input name="usuario" id="usuario" type="text" maxlength="10" style="text-align:center;" onkeypress="filtra_nome_identificador(); if (digitou_enter(true) && tem_info(this.value)) fID_AT.senha.focus();">
					
					<br />
					<br />
					<p class="R" style="margin: 10 10 2 10">SENHA</p>
					<input name="senha" id="senha" type="password" maxlength="15" style="text-align:center;" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) bCONSULTAR.click();">
					
					<br />
					<p class="R" style="margin: 4 10 0 10">&nbsp;</p>
					<input name="bCONSULTAR" id="bCONSULTAR" type="button" class="AT_Botao" 
							value="ENTRAR" title="inicia a sessão do usuário" onclick="if (confere(fID_AT)) {submit(); limpa_campos(fID_AT);}">
					<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
				</div>
				</form>
				</div>
				<div class="header">
				<div class="divTextoHeader">
				<p style="font-weight:bold;margin-top:6px;margin-bottom:6px;">Acompanhe sua solicitação<br />de peça em garantia</p>
				</div>
				</div>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="5">&nbsp;</td>
</tr>
-->

<% if (ID_AMBIENTE = ID_AMBIENTE__OLD01) OR (ID_AMBIENTE = ID_AMBIENTE__DIS) then %>
	<% if idOutroAmbiente <> ID_AMBIENTE__OLD01 then %>
	<td style="width:40px;" align="left">
		&nbsp;
	</td>
	<td align="left" valign="bottom">
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td align="center"><span class="NTit"><%=strTituloOutroAmbiente%></span></td>
			</tr>
			<tr>
				<td align="center">
				<div>
				<form target="<%=strTargetOutroAmb%>" action="<%=urlOutroAmb%>/orcamento/ConectaVerifica.asp" method="post" id="fID_OutroAmb" name="fID_OutroAmb" onsubmit="if (!blnVersaoNavegadorOk){alert('Esta versão do navegador não é suportada pelo site!!');return false;}">
				<div class="OA_QFn" style="margin: 0 10 0 10; width:220px" align="center">
					<br />
					<p class="R" style="margin: 10 10 2 10">USUÁRIO</p>
					<input name="usuario" id="usuario" type="text" maxlength="<%=MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR%>" style="text-align:center;" onkeypress="filtra_nome_identificador(); if (digitou_enter(true) && tem_info(this.value)) fID_OutroAmb.senha.focus();">
					
					<br />
					<br />
					<p class="R" style="margin: 10 10 2 10">SENHA</p>
					<input name="senha" id="senha" type="password" maxlength="15" style="text-align:center;" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) bCONSULTAR.click();">
					
					<br />
					<p class="R" style="margin: 4 10 0 10">&nbsp;</p>
					<input name="bCONSULTAR" id="bCONSULTAR" type="button" class="OA_Botao" 
							value="ENTRAR" title="inicia a sessão do usuário" onclick="if (confere(fID_OutroAmb)) { submit(); limpa_campos(fID_OutroAmb); }">
					<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
				</div>
				</form>
				</div>
				<div class="header">
				<div class="divTextoHeader">
				<p style="font-weight:bold;margin-top:6px;margin-bottom:6px;">Cadastre ou acompanhe<br />seu pedido</p>
				</div>
				</div>
				</td>
			</tr>
		</table>
	</td>
	<% end if %>
</tr>
<tr>
	<td colspan="5">&nbsp;</td>
</tr>
<% end if %>

<tr>
	<td colspan="3">&nbsp;</td>
</tr>
</table>

<noscript>
<p class="C" style="color:Red;">O navegador está com o JavaScript desabilitado!!</p>
<p class="C" style="color:Red;">É necessário habilitar o JavaScript para poder acessar o site!!</p>
<br />
</noscript>

</center>

<!-- ************   SEPARADOR   ************ -->
<table class="notPrint notVisible" width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint notVisible" width="100%" cellpadding="0" cellspacing="0">
<tr>
	<td align="center">
		<table class="notPrint" cellpadding="0" cellspacing="0">
			<tr>
			<td align="left">
				<a name="bEsqueciSenha" id="bEsqueciSenha" href="javascript:executaEsqueciSenha();" title="Esqueci a minha senha"><span class="C" style="cursor:pointer">Esqueci a minha senha</span></a>
			</td>
			</tr>
			<tr>
			<td align="left" valign="middle">
				<a name="bNaoSouCadastrado" id="bNaoSouCadastrado" href="javascript:executaNaoSouCadastrado();" title="Não sou cadastrado"><span class="C" style="cursor:pointer">Não sou cadastrado</span></a>
			</td>
			</tr>
		</table>
	</td>
</tr>
</table>

<div class="push"></div>
</div>

<div class="footer">
<center>
<div class="divTextoFooter">
<p style="font-weight:bold;margin-top:6px;margin-bottom:6px;">Se você não possui Usuário e Senha para acompanhar seus pedidos, solicite ao seu vendedor.<span style="display:none;"><br style="display:none;" />Usuário e Senha para acompanhar solicitações em garantia, solicite à Assistência Técnica.</span></p>
</div>
</center>
</div>

</body>
<% end if %>

</html>
