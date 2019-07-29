<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"    -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  F I L T R O R E L P R O D U T O S E M P R E S E N C A . A S P
'     =============================================================
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

    dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim pagina_destino
	pagina_destino=trim(request("pagina_destino"))
	if pagina_destino = "" then Response.Redirect("aviso.asp?id=" & ERR_PAG_DEST_INDEFINIDA)

	dim titulo_relatorio
	titulo_relatorio=trim(request("titulo_relatorio"))
	if titulo_relatorio = "" then Response.Redirect("aviso.asp?id=" & ERR_TIT_REL_INDEFINIDO)

	dim filtro_fabricante_obrigatorio
	filtro_fabricante_obrigatorio=Ucase(trim(request("filtro_fabricante_obrigatorio")))

	dim filtro_produto_obrigatorio
	filtro_produto_obrigatorio=Ucase(trim(request("filtro_produto_obrigatorio")))

	dim filtro_apenas_produto_permitido
	filtro_apenas_produto_permitido=Ucase(trim(request("filtro_apenas_produto_permitido")))

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



<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
	if (f.filtro_fabricante_obrigatorio.value=="S") {
		if (trim(f.c_fabricante.value)=="") {
			alert("Informe o código do fabricante!!");
			f.c_fabricante.focus();
			return;
			}
		}

	if (f.filtro_produto_obrigatorio.value=="S") {
		if (trim(f.c_produto.value)=="") {
			alert("Informe o código do produto!!");
			f.c_produto.focus();
			return;
			}
		}

	if ((trim(f.c_produto.value)!="") && (trim(f.c_fabricante.value)=="")){
		if (f.filtro_apenas_produto_permitido.value!="S") {
			if (trim(f.c_fabricante.value)=="") {
				alert("Informe o código do fabricante!!");
				f.c_fabricante.focus();
				return;
				}
			}
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


<body onload="fFILTRO.c_fabricante.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="<%=pagina_destino%>">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='pagina_destino' id="pagina_destino" value='<%=pagina_destino%>'>
<input type="hidden" name='titulo_relatorio' id="titulo_relatorio" value='<%=titulo_relatorio%>'>
<input type="hidden" name='filtro_fabricante_obrigatorio' id="filtro_fabricante_obrigatorio" value='<%=filtro_fabricante_obrigatorio%>'>
<input type="hidden" name='filtro_produto_obrigatorio' id="filtro_produto_obrigatorio" value='<%=filtro_produto_obrigatorio%>'>
<input type="hidden" name='filtro_apenas_produto_permitido' id="filtro_apenas_produto_permitido" value='<%=filtro_apenas_produto_permitido%>'>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO"><%=titulo_relatorio%></span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  FILTRO  -->
<table class="Qx" cellspacing="0">
<!--  EMPRESA  -->
	<tr bgcolor="#FFFFFF">
		<td class="MT" nowrap align="CENTER" style="background:azure;">
			<span class="PLTe">Empresa&nbsp;</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="center">
			<select id="c_empresa" name="c_empresa" style="margin:6px 10px 6px 10px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>    
<table class="Qx" cellspacing="0">
<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF">
		<td class="MT" nowrap align="CENTER" style="background:azure;">
			<span class="PLTe">Fabricante&nbsp;</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="center">
			<input name="c_fabricante" id="c_fabricante" class="PLLc" maxlength="4" style="margin-left:2pt;width:100px;"
				onkeypress="if (digitou_enter(true)) {this.blur(); fFILTRO.c_produto.focus();} filtra_fabricante();"
				onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);">
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
<!--  PRODUTO  -->
	<tr bgcolor="#FFFFFF">
		<td class="MT" nowrap align="center" style="background:azure;">
			<span class="PLTe">Produto&nbsp;</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="center">
			<input name="c_produto" id="c_produto" class="PLLc" maxlength="13" style="margin-left:2pt;width:100px;"
				onkeypress="if (digitou_enter(true)) {this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO)); bCONFIRMA.focus();} filtra_produto();"
				onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));">
		</td>
	</tr>
</table>
<br>
<table class="Qx" cellspacing="0">
<!--  ANÁLISE DE CRÉDITO  -->
	<tr bgcolor="#FFFFFF">
		<td class="MT" nowrap align="center" style="background:azure;">
			<span class="PLTe">Status da Análise de Crédito&nbsp;</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td class="MDE" align="left">
			<% intIdx = 0 %>
			<input type="radio" id="rb_analise_credito" name="rb_analise_credito" value="<%="TODOS"%>" checked><span class="C" style="cursor:default" onclick="fFILTRO.rb_analise_credito[<%=Cstr(intIdx)%>].click();">Todos</span>
		</td>
	</tr>
	<tr>
		<td class="MDE" align="left">
			<% intIdx = intIdx + 1 %>
			<input type="radio" id="rb_analise_credito" name="rb_analise_credito" value="<%=COD_AN_CREDITO_PENDENTE_VENDAS%>"><span class="C" style="cursor:default" onclick="fFILTRO.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS)%></span>
		</td>
	</tr>
	<tr>
		<td class="MDE" align="left">
			<% intIdx = intIdx + 1 %>
			<input type="radio" id="rb_analise_credito" name="rb_analise_credito" value="<%=COD_AN_CREDITO_PENDENTE_ENDERECO%>"><span class="C" style="cursor:default" onclick="fFILTRO.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE_ENDERECO)%></span>
		</td>
	</tr>
	<tr>
		<td class="MDE" align="left">
			<% intIdx = intIdx + 1 %>
			<input type="radio" id="rb_analise_credito" name="rb_analise_credito" value="<%=COD_AN_CREDITO_PENDENTE%>"><span class="C" style="cursor:default" onclick="fFILTRO.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE)%></span>
		</td>
	</tr>
	<tr>
		<td class="MDE" align="left">
			<% intIdx = intIdx + 1 %>
			<input type="radio" id="rb_analise_credito" name="rb_analise_credito" value="<%=COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO%>"><span class="C" style="cursor:default" onclick="fFILTRO.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)%></span>
		</td>
	</tr>
	<tr>
		<td class="MDE" align="left">
			<% intIdx = intIdx + 1 %>
			<input type="radio" id="rb_analise_credito" name="rb_analise_credito" value="<%=COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO%>"><span class="C" style="cursor:default" onclick="fFILTRO.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)%></span>
		</td>
	</tr>
	<tr>
		<td class="MDBE" align="left">
			<% intIdx = intIdx + 1 %>
			<input type="radio" id="rb_analise_credito" name="rb_analise_credito" value="<%=COD_AN_CREDITO_OK%>"><span class="C" style="cursor:default" onclick="fFILTRO.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_OK)%></span>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td class="MT" nowrap align="center" style="background:azure;">
			<span class="PLTe">Entrega Imediata&nbsp;</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td class="MDE" align="left">
			<% intIdx = 0 %>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" value="<%="TODOS"%>" checked><span class="C" style="cursor:default" onclick="fFILTRO.rb_etg_imediata[<%=Cstr(intIdx)%>].click();">Todos</span>
		</td>
	</tr>
	<tr>
		<td class="MDE" align="left">
			<% intIdx = intIdx + 1 %>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" value="<%=COD_ETG_IMEDIATA_NAO%>"><span class="C" style="cursor:default" onclick="fFILTRO.rb_etg_imediata[<%=Cstr(intIdx)%>].click();">Não</span>
		</td>
	</tr>
	<tr>
		<td class="MDBE" align="left">
			<% intIdx = intIdx + 1 %>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" value="<%=COD_ETG_IMEDIATA_SIM%>"><span class="C" style="cursor:default" onclick="fFILTRO.rb_etg_imediata[<%=Cstr(intIdx)%>].click();">Sim</span>
		</td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
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
</body>
</html>
