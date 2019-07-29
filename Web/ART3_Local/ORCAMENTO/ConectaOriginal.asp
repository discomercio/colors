<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->

<%
'     =====================================
'	  C O N E C T A O R I G I N A L . A S P
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

	dim strTarget
	if ID_PARAM_SITE = COD_SITE_ARTVEN_FABRICANTE then
		strTarget = IdFormTargetArtvenFabricante
	elseif ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then
		strTarget = IdFormTargetAssistenciaTecnica
	else
		strTarget = IdFormTargetArtvenBonshop
		end if

	strTarget = strTarget & "ORC"
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
	<title><%=TITULO_JANELA_MODULO_ORCAMENTO%></title>
	</head>

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
		alert("Esta versão do navegador não é suportada!!\nPor favor, utilize o Internet Explorer versão 7 ou superior!!");
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
.QFn
{
	background:#deecb8;
}
.Botao
{
	color:White;
	background: #064a93;
	border: double 3px #FFFAFA;
}
.wrapper
{
	min-height:100%;
	height:auto !important;
	height:100%;
	margin: 0 auto -2.5em;
}
.push
{
	height:1em;
}
.header
{
	height:2.5em;
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
	height:2.5em;
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
</style>


<% if isHorarioManutencaoSistema then %>
<body>
<center>
<br />
<h1>Sistema em manutenção no período das <%=HORARIO_INICIO_MANUTENCAO_SISTEMA%> até <%=HORARIO_TERMINO_MANUTENCAO_SISTEMA%><br /><br />Por favor, tente mais tarde.</h1>
</center>
</body>
<% else %>
<body onload="posiciona_foco(fID);">
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
			<td id="topo_verde" colspan="3">
				<div id="moldura_do_letreiro">
					<div id="letreiro_div" style="display:block;"></div>
				</div>
				<div id="telefone"></div>
			</td>
		</tr>
		<tr>
			<td id="topo_azul" colspan="3">&nbsp;</td>
		</tr>
	</tbody>
</table>

<br />
<br />
<div class="header">
<center>
<div class="divTextoHeader">
<p style="font-weight:bold;margin-top:6px;margin-bottom:6px;">Na Área Restrita você cadastra e acompanha todos os pedidos de seus clientes, dentro do sistema da Bonshop. Aproveite!!</p>
</div>
</center>
</div>

<br />
<br />
<form target="<%=strTarget%>" action="ConectaVerifica.asp" method="post" id="fID" name="fID" onsubmit="if (!blnVersaoNavegadorOk){alert('Esta versão do navegador não é suportada pelo site!!');return false;}">
<div class="QFn" style="margin: 0 10 0 10; width:220px" align="center">
	<p class="R" style="margin: 10 10 2 10">USUÁRIO</p>
	<input name="usuario" id="usuario" type="text" maxlength="<%=MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR%>" style="text-align:center;" onkeypress="filtra_nome_identificador(); if (digitou_enter(true) && tem_info(this.value)) fID.senha.focus();">
	
	<p class="R" style="margin: 10 10 2 10">SENHA</p>
	<input name="senha" id="senha" type="password" maxlength="15" style="text-align:center;" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) bCONSULTAR.click();">
	
	<p class="R" style="margin: 4 10 0 10">&nbsp;</p>
	<input name="bCONSULTAR" id="bCONSULTAR" type="button" class="Botao" 
		   value="ENTRAR" title="inicia a sessão do usuário" onclick="if (confere(fID)) {submit(); limpa_campos(fID);}">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
</div>
</form>

<noscript>
<p class="C" style="color:Red;">O navegador está com o JavaScript desabilitado!!</p>
<p class="C" style="color:Red;">É necessário habilitar o JavaScript para poder acessar o site!!</p>
<br />
</noscript>

</center>

<!-- ************   SEPARADOR   ************ -->
<table class="notPrint notVisible" width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
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
<p style="font-weight:bold;margin-top:6px;margin-bottom:6px;">Se você ainda não possui Usuário e Senha, solicite ao seu vendedor.</p>
</div>
</center>
</div>

</body>
<% end if %>

</html>
