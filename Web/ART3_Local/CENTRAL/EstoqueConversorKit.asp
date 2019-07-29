<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================
'	  E S T O Q U E C O N V E R S O R K I T . A S P
'     =============================================
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
	if Not operacao_permitida(OP_CEN_CONVERSOR_KITS, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
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



<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function cancela_onpaste() {
	return false;
}

function fKITConfirma( f ) {
	var i, b, ha_item;

	if (trim(f.c_nfe_emitente.value) == "")
	{
		alert("Selecione a empresa cujo estoque será processado!!");
		f.c_nfe_emitente.focus();
		return;
	}
	if (!isEAN(trim(f.c_kit.value))) {
		if (trim(f.c_kit_fabricante.value)=="") {
			alert("Informe o código do fabricante do kit!!");
			f.c_kit_fabricante.focus();
			return;
			}
		}
	if (trim(f.c_kit.value)=="") {
		alert("Informe o código de produto do kit a ser gerado!!");
		f.c_kit.focus();
		return;
		}
	if (converte_numero(trim(f.c_kit_qtde.value))<=0) {
		alert("Informe a quantidade de kits a ser convertida!!");
		f.c_kit_qtde.focus();
		return;
		}
	
	ha_item=false;
	for (i=0; i < f.c_codigo.length; i++) {
		b=false;
		if (trim(f.c_fabricante[i].value)!="") b=true;
		if (trim(f.c_codigo[i].value)!="") b=true;
		if (trim(f.c_qtde[i].value)!="") b=true;
		
		if (b) {
			ha_item=true;
			if (!isEAN(trim(f.c_codigo[i].value))) {
				if (trim(f.c_fabricante[i].value)=="") {
					alert("Informe o fabricante do produto usado na composição do kit!!");
					f.c_fabricante[i].focus();
					return;
					}
				}
			if (trim(f.c_codigo[i].value)=="") {
				alert("Informe o código do produto usado na composição do kit!!");
				f.c_codigo[i].focus();
				return;
				}
			if (trim(f.c_qtde[i].value)=="") {
				alert("Informe a quantidade de produtos usada para compor 1 unidade do kit!!");
				f.c_qtde[i].focus();
				return;
				}
			if (parseInt(f.c_qtde[i].value)<=0) {
				alert("Quantidade inválida!!");
				f.c_qtde[i].focus();
				return;
				}
			if (isEAN(f.c_kit.value)) {
				if (f.c_kit.value==f.c_codigo[i].value) {
					alert("O código de produto do kit a ser gerado não pode constar na relação de produtos usados em sua composição!!");
					f.c_codigo[i].focus();
					return;
					}
				}
			else {
				if ((f.c_kit_fabricante.value==f.c_fabricante[i].value)&&(f.c_kit.value==f.c_codigo[i].value)) {
					alert("O código de produto do kit a ser gerado não pode constar na relação de produtos usados em sua composição!!");
					f.c_codigo[i].focus();
					return;
					}
				}
			}
		}

	if (!ha_item) {
		alert("Não há produtos na lista de composição do kit!!");
		f.c_fabricante[0].focus();
		return;
		}

	if (trim(f.c_ncm.value) == "") {
		alert("Informe o NCM do kit!!");
		f.c_ncm.focus();
		return;
	}

	if ((f.c_ncm.value.length != 2) && (f.c_ncm.value.length != 8)) {
		alert("NCM possui tamanho inválido!!");
		f.c_ncm.focus();
		return;
	}

	if (trim(f.c_ncm.value) != trim(f.c_ncm_redigite.value)) {
		alert("Falha na conferência do NCM redigitado!!");
		f.c_ncm_redigite.focus();
		return;
	}
	
	if (trim(f.c_cst.value) == "") {
		alert("Informe o CST (entrada) do kit!!");
		f.c_cst.focus();
		return;
	}

	if (f.c_cst.value.length != 3) {
		alert("CST possui tamanho inválido!!");
		f.c_cst.focus();
		return;
	}

	if (trim(f.c_cst.value) != trim(f.c_cst_redigite.value)) {
		alert("Falha na conferência do CST redigitado!!");
		f.c_cst_redigite.focus();
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


<body onload="if (trim(fKIT.c_nfe_emitente.value)=='') fKIT.c_nfe_emitente.focus();">
<center>

<form id="fKIT" name="fKIT" method="post" action="EstoqueConversorKitConsiste.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Conversor para Cadastramento de Kits</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  KIT A SER GERADO  -->
<table class="Qx" cellspacing="0" cellpadding="0">
	<!--  TÍTULO  -->
	<tr bgcolor="#FFFFFF">
	<td><span style="width:30px;">&nbsp;</span></td>
	<td colspan="3" class="MT" valign="middle" align="center" nowrap style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>KIT A SER GERADO</span></td>
	</tr>
	<!--  EMPRESA  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" colspan="3"><span class="PLTe">Empresa</span>
		<br />
		<select name="c_nfe_emitente" id="c_nfe_emitente" class="C" style="margin:4px 8px 10px 8px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
		<% =wms_apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
		</select>
	</td>
	</tr>
	<!--  FABRICANTE/PRODUTO/QTDE  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" nowrap><span class="PLTe">Fabricante&nbsp;</span>
		<br><input name="c_kit_fabricante" id="c_kit_fabricante" class="PLLe" maxlength="4" style="width:60px;" onkeypress="if (digitou_enter(true)) fKIT.c_kit.focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"></td>
	<td class="MDB" style="border-left:0pt;"><span class="PLTe">Kit</span>
		<br><input name="c_kit" id="c_kit" class="PLLe" maxlength="13" style="margin-left:2pt;width:100px;" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fKIT.c_kit_qtde.focus(); filtra_produto();" 
			onblur="this.value=normaliza_produto(this.value);"></td>
	<td class="MDB" align="right" nowrap><span class="PLTe" style="margin-right:2pt;">Qtde</span>
		<br><input name="c_kit_qtde" id="c_kit_qtde" class="PLLd" maxlength="4" style="width:35px;" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fKIT.c_ncm.focus(); filtra_numerico();"></td>
	</tr>
	<!--  NCM / NCM (redigite)  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="3" class="MDBE">
		<table width="100%" cellpadding=0 cellspacing=0>
			<tr>
				<td width="50%" class="MD">
					<span class="PLTe">NCM</span>
					<br><input name="c_ncm" id="c_ncm" class="PLLe" maxlength="8" style="margin-left:2pt;width:80px;"
						onkeypress="if (digitou_enter(true)) fKIT.c_ncm_redigite.focus(); filtra_numerico();"
						onblur="this.value=trim(this.value);">
				</td>
				<td width="50%">
					<span class="PLTe">NCM <span style="font-size:7pt;">[redigite]</span></span>
					<br><input name="c_ncm_redigite" id="c_ncm_redigite" class="PLLe" maxlength="8" style="margin-left:2pt;width:80px;"
						onkeypress="if (digitou_enter(true)) fKIT.c_cst.focus(); filtra_numerico();"
						onpaste="return cancela_onpaste()"
						onblur="this.value=trim(this.value);">
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<!--  CST / CST (redigite)  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="3" class="MDBE">
		<table width="100%" cellpadding=0 cellspacing=0>
			<tr>
				<td width="50%" class="MD">
					<span class="PLTe">CST (entrada)</span>
					<br><input name="c_cst" id="c_cst" class="PLLe" maxlength="3" style="margin-left:2pt;width:80px;"
						onkeypress="if (digitou_enter(true)) fKIT.c_cst_redigite.focus(); filtra_numerico();" 
						onblur="this.value=trim(this.value);">
				</td>
				<td width="50%">
					<span class="PLTe">CST (entrada) <span style="font-size:7pt;">[redigite]</span></span>
					<br><input name="c_cst_redigite" id="c_cst_redigite" class="PLLe" maxlength="3" style="margin-left:2pt;width:80px;"
						onkeypress="if (digitou_enter(true)) fKIT.c_documento.focus(); filtra_numerico();"
						onpaste="return cancela_onpaste()"
						onblur="this.value=trim(this.value);">
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<!--  DOCUMENTO  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="3" class="MDBE"><span class="PLTe">Documento</span>
		<br><input name="c_documento" id="c_documento" class="PLLe" maxlength="30" style="margin-left:2pt;width:270px;" onkeypress="if (digitou_enter(true)) fKIT.c_fabricante[0].focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"></td>
	</tr>
</table>
<br><br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellSpacing="0">
	<!--  TÍTULO DA TABELA  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="3" class="MT" valign="middle" align="center" nowrap style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>COMPOSIÇÃO DE 1 UNIDADE DO KIT</span></td>
	</tr>
	<!--  TÍTULO DAS COLUNAS  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE"><p class="PLTe">Fabricante&nbsp;</p></td>
	<td class="MDB"><p class="PLTe">Produto</p></td>
	<td class="MDB"><p class="PLTd">Qtde</p></td>
	</tr>
<% for i=1 to MAX_PRODUTOS_CONVERSOR_KIT %>
	<tr>
	<td>
		<input name="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:30px;text-align:right;color:#808080;" 
			value="<%=Cstr(i) & ". " %>"></td>
	<td class="MDBE">
		<input name="c_fabricante" class="PLLe" maxlength="4" style="width:60px;" 
			onkeypress="if (digitou_enter(true)) fKIT.c_codigo[<%=Cstr(i-1)%>].focus(); filtra_fabricante();" 
			onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"></td>
	<td class="MDB">
		<input name="c_codigo" class="PLLe" maxlength="13" style="width:100px;" 
			onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(<%=Cstr(i)%>!=1))) if (trim(this.value)=='') bCONFIRMA.focus(); else fKIT.c_qtde[<%=Cstr(i-1)%>].focus(); filtra_produto();" 
			onblur="this.value=normaliza_produto(this.value);"></td>
	<td class="MDB" align="right">
		<input name="c_qtde" class="PLLd" maxlength="4" style="width:35px;" 
			onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==fKIT.c_qtde.length) bCONFIRMA.focus(); else fKIT.c_fabricante[<%=Cstr(i)%>].focus();} filtra_numerico();"></td>
	</tr>
<% next %>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fKITConfirma(fKIT)" title="vai para a página de confirmação">
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