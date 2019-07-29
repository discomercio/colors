<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  S E N H A D E S C S U P P E S Q C L I E N T E . A S P
'     ======================================================
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

	dim s, idx

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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
var iop, achou;
	achou = false;
	iop = -1;
	
	iop++;
	if (f.rb_op[iop].checked) {
		achou=true;
		if (trim(f.c_cnpj_cpf.value)=="") {
			alert("Preencha o CNPJ/CPF!!");
			f.c_cnpj_cpf.focus();
			return;
			}
		if (!cnpj_cpf_ok(f.c_cnpj_cpf.value)) {
			alert("CNPJ/CPF inválido!!");
			f.c_cnpj_cpf.focus();
			return;
			}
		}

	iop++;
	if (f.rb_op[iop].checked) {
		achou=true;
		if (trim(f.c_nome_completo.value)=="") {
			alert("Preencha o nome completo do cliente!!");
			f.c_nome_completo.focus();
			return;
			}
		}

	iop++;
	if (f.rb_op[iop].checked) {
		achou=true;
		}

	if (!achou) {
		alert("Selecione uma opção de pesquisa!!");
		return;
		}

	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}

function posiciona_foco() {
var f, i;
	f=fFILTRO;
	for (i=0; i<f.rb_op.length; i++) {
		if (f.rb_op[i].checked) {
			f.rb_op[i].click();
			return;
			}
		}
		
	focus();
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

<style type="text/css">
#rb_op {
	vertical-align: middle;
	}
#srb_op {
	cursor: default;
	margin: 0px 0px 0px 0px;
	vertical-align: middle;
	}
</style>

<body onload="posiciona_foco();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="SenhaDescSupPesqClienteExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Senha para Autorização de Desconto Superior</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  FILTROS  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MT" NOWRAP>
	<%	idx = 0	%>

<!--  CPF/CNPJ  -->
	<%  idx=idx+1 %>
	<input type="radio" name="rb_op" id="rb_op" value="<%=Cstr(idx)%>" class="CBOX" onclick="fFILTRO.c_cnpj_cpf.focus();">
	<span class="PLTe" name="srb_op" id="srb_op" onclick="fFILTRO.rb_op[<%=Cstr(idx-1)%>].click(); fFILTRO.c_cnpj_cpf.focus();">CNPJ/CPF</span>
		<br><input name="c_cnpj_cpf" id="c_cnpj_cpf" type="text" class="PLLe" maxlength="18" style="width:220px;" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="fFILTRO.rb_op[<%=Cstr(idx-1)%>].click(); if (digitou_enter(true) && tem_info(this.value) && cnpj_cpf_ok(this.value)) {this.blur();bCONFIRMA.click();} filtra_cnpj_cpf();" onclick="fFILTRO.rb_op[<%=Cstr(idx-1)%>].click();">
	</td></tr>

<!--  NOME DO CLIENTE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP>
	<%  idx=idx+1 %>
	<input type="radio" name="rb_op" id="rb_op" value="<%=Cstr(idx)%>" class="CBOX" onclick="fFILTRO.c_nome_completo.focus();">
	<span class="PLTe" name="srb_op" id="srb_op" onclick="fFILTRO.rb_op[<%=Cstr(idx-1)%>].click(); fFILTRO.c_nome_completo.focus();">NOME COMPLETO DO CLIENTE</span>
		<br><input name="c_nome_completo" id="c_nome_completo" class="PLLe" type="text" maxlength="60" style="width:300px;" onblur="this.value=trim(this.value);" onkeypress="fFILTRO.rb_op[<%=Cstr(idx-1)%>].click(); if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.click();" onclick="fFILTRO.rb_op[<%=Cstr(idx-1)%>].click();">
	</td></tr>

<!--  PARTE DO NOME DO CLIENTE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP>
	<%  idx=idx+1 %>
	<input type="radio" name="rb_op" id="rb_op" value="<%=Cstr(idx)%>" class="CBOX" onclick="fFILTRO.c_nome_parcial.focus();">
	<span class="PLTe" name="srb_op" id="srb_op" onclick="fFILTRO.rb_op[<%=Cstr(idx-1)%>].click(); fFILTRO.c_nome_parcial.focus();">PARTE DO NOME DO CLIENTE</span>
		<br><input name="c_nome_parcial" id="c_nome_parcial" class="PLLe" type="text" maxlength="60" style="width:300px;" onblur="this.value=trim(this.value);" onkeypress="fFILTRO.rb_op[<%=Cstr(idx-1)%>].click(); if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.click();" onclick="fFILTRO.rb_op[<%=Cstr(idx-1)%>].click();">
	</td></tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="MenuFuncoesAdministrativas.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página inicial">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>
