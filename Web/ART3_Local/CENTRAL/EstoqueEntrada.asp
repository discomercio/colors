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
'	  E S T O Q U E E N T R A D A . A S P
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim i
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_ENTRADA_MERCADORIAS_ESTOQUE, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

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
	<title>CENTRAL</title>
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<%  dim strScript
	strScript = _
		"<script language='JavaScript' type='text/javascript'>" & chr(13) & _
		"var MAX_TAM_T_ESTOQUE_CAMPO_OBS = " & Cstr(MAX_TAM_T_ESTOQUE_CAMPO_OBS) & ";" & chr(13) & _
		"</script>" & chr(13)
	Response.Write strScript
%>

<script language="JavaScript" type="text/javascript">
	$(function () {
	//  CADASTRAMENTO DE EMPRESA+CD ESTÁ HABILITADO?
		<% if Not CADASTRAR_WMS_CD_ENTRADA_ESTOQUE then %>
		$(".trWmsCd").hide();
		<% end if %>
	});

function fESTOQConfirma( f ) {
var i, b, ha_item;
	if (trim(f.c_fabricante.value)=="") {
		alert("Informe o código do fabricante!!");
		f.c_fabricante.focus();
		return;
		}
	if (trim(f.c_documento.value)=="") {
		alert("Preencha o número do documento!!");
		f.c_documento.focus();
		return;
		}
	<% if CADASTRAR_WMS_CD_ENTRADA_ESTOQUE then %>
	if (trim(f.c_id_nfe_emitente.value) == "") {
		alert("Selecione uma empresa!!");
		f.c_id_nfe_emitente.focus();
		return;
	}
	<% end if %>
	
	ha_item=false;
	for (i=0; i < f.c_codigo.length; i++) {
		b=false;
		if (trim(f.c_codigo[i].value)!="") b=true;
		if (trim(f.c_qtde[i].value)!="") b=true;
		
		if (b) {
			ha_item=true;
			if (trim(f.c_codigo[i].value)=="") {
				alert("Informe o código do produto!!");
				f.c_codigo[i].focus();
				return;
				}
			if (trim(f.c_qtde[i].value)=="") {
				alert("Informe a quantidade!!");
				f.c_qtde[i].focus();
				return;
				}
			if (parseInt(f.c_qtde[i].value)<=0) {
				alert("Quantidade inválida!!");
				f.c_qtde[i].focus();
				return;
				}
			}
		}

	if (!ha_item) {
		alert("Não há produtos na lista!!");
		f.c_codigo[0].focus();
		return;
		}
		
	s = "" + f.c_obs.value;
	if (s.length > MAX_TAM_T_ESTOQUE_CAMPO_OBS) {
		alert('Conteúdo de "Observações" excede em ' + (s.length-MAX_TAM_T_ESTOQUE_CAMPO_OBS) + ' caracteres o tamanho máximo de ' + MAX_TAM_T_ESTOQUE_CAMPO_OBS + '!!');
		f.c_obs.focus();
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

<style type="text/css">
select
{
	margin-left:8px;
}
#ckb_especial {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
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


<body onload="if (trim(fESTOQ.c_fabricante.value)=='') fESTOQ.c_fabricante.focus();">
<center>

<form id="fESTOQ" name="fESTOQ" method="post" action="EstoqueEntradaConsiste.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Entrada de Mercadorias no Estoque</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  CADASTRO DA ENTRADA DE MERCADORIAS NO ESTOQUE  -->
<table class="Qx" cellspacing="0" cellpadding="0">
<!--  EMPRESA COMPRADORA / CENTRO DE DISTRIBUIÇÃO  -->
	<tr bgcolor="#FFFFFF" class="trWmsCd">
		<td colspan="2">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td class="MT" align="left" width="50%"><span class="PLTe">Empresa</span>
			<br />
			<select id="c_id_nfe_emitente" name="c_id_nfe_emitente" style="margin-top:4pt;margin-bottom:4pt;min-width:100px;">
			<%=wms_apelido_empresa_nfe_emitente_monta_itens_select(Null)%>
			</select>
			</td>
		</tr>
		</table>
		</td>
	</tr>

<!--  FABRICANTE/DOCUMENTO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">Fabricante</span>
		<br><input name="c_fabricante" id="c_fabricante" class="PLLe" maxlength="4" style="width:100px;text-align:center;" onkeypress="if ((digitou_enter(true))&&tem_info(this.value)) fESTOQ.c_documento.focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"></td>
	<td class="MDB" style="border-left:0pt;" align="left"><span class="PLTe">Documento</span>
		<br><input name="c_documento" id="c_documento" class="PLLe" maxlength="30" style="margin-left:2pt;width:270px;" onkeypress="if (digitou_enter(true)) fESTOQ.c_codigo[0].focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"></td>
	</tr>

<!--  ENTRADA ESPECIAL  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Tipo de Cadastramento</span>
		<br><input type="checkbox" class="rbOpt" tabindex="-1" id="ckb_especial" name="ckb_especial" value="ESPECIAL_ON"
		<%if Not operacao_permitida(OP_CEN_ENTRADA_ESPECIAL_ESTOQUE, s_lista_operacoes_permitidas) then Response.Write " disabled" %>
		><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.ckb_especial.click();">Entrada Especial</span>
	</td>
	</tr>

<!--  OBS  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Observações</span>
		<br><textarea name="c_obs" id="c_obs" class="PLLe" rows="<%=Cstr(MAX_LINHAS_ESTOQUE_OBS)%>"
				style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_T_ESTOQUE_CAMPO_OBS);" onblur="this.value=trim(this.value);"
				></textarea>
	</td>
	</tr>
</table>
<br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0">
	<tr bgColor="#FFFFFF">
	<td align="left">&nbsp;</td>
	<td class="MB" align="left"><span class="PLTe">Produto</span></td>
	<td class="MB" align="left"><span class="PLTd">Qtde</span></td>
	</tr>
<% for i=1 to MAX_PRODUTOS_ENTRADA_ESTOQUE %>
	<tr>
	<td align="left">
		<input name="c_linha" id="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:30px;text-align:right;color:#808080;" 
			value="<%=Cstr(i) & ". " %>"></td>
	<td class="MDBE" align="left">
		<input name="c_codigo" id="c_codigo" class="PLLe" maxlength="13" style="width:100px;" 
			onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(<%=Cstr(i)%>!=1))) if (trim(this.value)=='') bCONFIRMA.focus(); else fESTOQ.c_qtde[<%=Cstr(i-1)%>].focus(); filtra_produto();" 
			onblur="this.value=normaliza_produto(this.value);"></td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde" class="PLLd" maxlength="4" style="width:35px;" 
			onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==fESTOQ.c_qtde.length) bCONFIRMA.focus(); else fESTOQ.c_codigo[<%=Cstr(i)%>].focus();} filtra_numerico();"></td>
	</tr>
<% next %>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fESTOQConfirma(fESTOQ)" title="vai para a página de confirmação">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
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