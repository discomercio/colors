<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<%
'     =====================
'	  C O N E C T A . A S P
'     =====================
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

'	EXIBIÇÃO DE BOTÕES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True

	dim strTarget
	if ID_PARAM_SITE = COD_SITE_ARTVEN_FABRICANTE then
		strTarget = IdFormTargetArtvenFabricante
	elseif ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then
		strTarget = IdFormTargetAssistenciaTecnica
	else
		strTarget = IdFormTargetArtvenBonshop
		end if
	
	strTarget = strTarget & "CEN"
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
	<title>CENTRAL ADMINISTRATIVA - LOGON</title>
	</head>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var blnVersaoNavegadorOk = false;
if (isVersaoNavegadorOk()) blnVersaoNavegadorOk = true;

configura_painel_logon(); 

var fCepPopup;

function ProcessaSelecaoCEP(){};

function AbrePesquisaCep(){
var strUrl;
	try
		{
	//  SE JÁ HOUVER UMA JANELA DE PESQUISA DE CEP ABERTA, GARANTE QUE ELA SERÁ FECHADA 
	// E UMA NOVA SERÁ CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')
		fCepPopup=window.open("", "AjaxCepPesqPopup","status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=5,height=5,left=0,top=0");
		fCepPopup.close();
		}
	catch (e) {
	 // NOP
		}
	ProcessaSelecaoCEP=null;
	strUrl="../Global/AjaxCepPesqPopup.asp?ModoApenasConsulta=S";
	fCepPopup=window.open(strUrl, "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=980,height=650,left=0,top=0");
	fCepPopup.focus();
}

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

<script type="text/javascript">
	function exibeJanelaCEP_Consulta() {
		$.mostraJanelaCEP(null);
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__E_JANELABUSCACEP_CSS%>" rel="stylesheet" type="text/css">

<% if isHorarioManutencaoSistema then %>
<body>
<center>
<br />
<h1>Sistema em manutenção no período das <%=HORARIO_INICIO_MANUTENCAO_SISTEMA%> até <%=HORARIO_TERMINO_MANUTENCAO_SISTEMA%><br /><br />Por favor, tente mais tarde.</h1>
</center>
</body>
<% else %>
<body id="corpoPagina" onload="posiciona_foco(fID);">

<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->

<br />
<br />

<!--  L O G O N  -->
<span class="T">IDENTIFICAÇÃO</span>

<form TARGET="<%=strTarget%>" ACTION="conectaverifica.asp" METHOD="POST" id="fID" name="fID" onsubmit="if (!blnVersaoNavegadorOk){alert('Esta versão do navegador não é suportada pelo site!!');return false;}">
<div class="QFn" style="margin: 0 10 0 10; width:220px" align="CENTER">
	<br />
	<p class="R" style="margin: 10 10 2 10">USUÁRIO</p>
	<input name="usuario" id="usuario" type="text" maxlength="10" style="text-align:center;" onkeypress="filtra_nome_identificador(); if (digitou_enter(true) && tem_info(this.value)) fID.senha.focus()">
	
	<br />
	<br />
	<p class="R" style="margin: 10 10 2 10">SENHA DE ACESSO</p>
	<input name="senha" id="senha" type="password" maxlength="15" style="text-align:center;" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) bCONSULTAR.click();">
	
	<br />
	<p class="R" style="margin: 4 10 0 10">&nbsp;</p>
	<input name="bCONSULTAR" type="button" class="Botao" 
		   value="ENTRAR" title="inicia a sessão do usuário" onclick="if (confere(fID)) {submit(); limpa_campos(fID);}">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
</div>
</form>
</center>

<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width="100%" cellPadding="0" CellSpacing="0">
<tr><td align="right">
	<% if blnPesquisaCEPAntiga then %>
	<span name="bPesqCep" id="bPesqCep" class="LPesqCep" onclick="AbrePesquisaCep();">Pesquisar CEP</span>
	<% end if %>
	<% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;nbsp;nbsp;" %>
	<% if blnPesquisaCEPNova then %>
	<span name="bPesqCep" id="bPesqCep" class="LPesqCep" onclick="exibeJanelaCEP_Consulta();">Pesquisar CEP</span>
	<% end if %>
</td></tr>
</table>

</body>
<% end if %>

</html>
