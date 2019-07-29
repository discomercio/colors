<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'	REVISADO P/ IE10

	On Error GoTo 0
	Err.Clear

	dim strModoApenasConsulta, strCepDefault
	strModoApenasConsulta=UCase(Trim(Request("ModoApenasConsulta")))
	strCepDefault=retorna_so_digitos(Trim(Request("CepDefault")))




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S
' _____________________________________________________________________________________________


function UF_monta_itens_select(byval id_default)
dim strResp, ha_default, strListaUF, strUF, vUF, intContador
	id_default = UCase(Trim("" & id_default))
	ha_default=False
	strListaUF="AC|AL|AM|AP|BA|CE|DF|ES|GO|MA|MG|MS|MT|PA|PB|PE|PI|PR|RJ|RN|RO|RR|RS|SC|SE|SP|TO"
	vUF=Split(strListaUF,"|")
	for intContador=LBound(vUF) to UBound(vUF)
		strUF = vUF(intContador)
		if (id_default<>"") And (id_default=strUF) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & strUF & "'>"
		strResp = strResp & strUF
		strResp = strResp & "</OPTION>" & chr(13)
		next

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
	
	UF_monta_itens_select = strResp
end function
%>

<html>
<head>
	<title>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Pesquisa de CEP
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</title>
</head>

<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var objAjaxPesqCep;
var objAjaxPesqLocalidades;
var OPCAO_PESQUISA_POR_CEP = "POR_CEP";
var OPCAO_PESQUISA_POR_ENDERECO = "POR_END";
var COL_CHECK = 0;
var COL_CEP = 1;
var COL_UF = 2;
var COL_LOCALIDADE = 3;
var COL_BAIRRO = 4;
var COL_LOGRADOURO = 5;
var COL_LOGRADOURO_COMPLEMENTO = 6;

function CancelarOperacao() {
	window.close();
}

function ConfirmarOperacao() {
var i, oRow, idxSelecionado, strCep, strUF, strLocalidade, strBairro, strLogradouro;
var strEnderecoNumero, strEnderecoComplemento;
	idxSelecionado=-1;
	for (i=0; i<rb_check.length;i++) {
		if (rb_check[i].checked) {
			idxSelecionado=i;
			break;
			}
		}
	if (idxSelecionado==-1) {
		alert("Nenhum CEP foi selecionado!!");
		return;
		}	

	if (trim(c_endereco_numero.value)=="") {
		alert("Informe o número do endereço!!");
		c_endereco_numero.focus();
		return;
		}
	
//	LEMBRE-SE: O ARRAY DE CAMPOS 'RB_CHECK' TEM O 1º CAMPO GERADO POR UM 'INPUT HIDDEN'
//  =========  C/ A FUNÇÃO DE SEMPRE GERAR UM ARRAY, MESMO NO CASO DA TABELA TER APENAS 1 LINHA.
//             PORTANTO, SEMPRE HAVERÁ 1 RB_CHECK A MAIS QUE O TOTAL DE LINHAS DE RESPOSTA E A 1ª
//             LINHA CORRESPONDE AO RB_CHECK[1] E NÃO AO RB_CHECK[0]
	idxSelecionado--;
	oRow=oTBodyDados.rows[idxSelecionado];
	strCep=trim(oRow.cells[COL_CEP].innerHTML);
	if (strCep=="&nbsp;") strCep="";
	strUF=trim(oRow.cells[COL_UF].innerHTML);
	if (strUF=="&nbsp;") strUF="";
	strLocalidade=trim(oRow.cells[COL_LOCALIDADE].innerHTML);
	if (strLocalidade=="&nbsp;") strLocalidade="";
	strBairro=trim(oRow.cells[COL_BAIRRO].innerHTML);
	if (strBairro=="&nbsp;") strBairro="";
	strLogradouro=trim(oRow.cells[COL_LOGRADOURO].innerHTML);
	if (strLogradouro=="&nbsp;") strLogradouro="";
	strEnderecoNumero=trim(c_endereco_numero.value);
	strEnderecoComplemento=trim(c_endereco_complemento.value);
	
	try {
	//  A JANELA 'OPENER' JÁ PODE TER SIDO FECHADA
		window.opener.ProcessaSelecaoCEP(strCep, strUF, strLocalidade, strBairro, strLogradouro, strEnderecoNumero, strEnderecoComplemento);
		}
	catch (e) {
	 // NOP
		}
	window.close();
}

function IniciaPainel() {
	if (trim(c_cep_pesq.value)!="") {
		ExecutaPesquisaCEP(OPCAO_PESQUISA_POR_CEP);
		}
	c_cep_pesq.select();
	c_cep_pesq.focus();
}

function LimpaListaLocalidades() {
var i, oOption;
	for (i=c_localidade_pesq.length-1; i >= 0; i--) {
		c_localidade_pesq.remove(i);
		}

//  Cria um item vazio
	oOption=document.createElement("OPTION");
	c_localidade_pesq.options.add(oOption);
	oOption.innerText="";
	oOption.value="";
}

function LimpaTabelaResultado() {
var i;
	for (i=oTBodyDados.rows.length-1; i >= 0; i--) {
		oTBodyDados.deleteRow(i);
		}
}

function TrataRespostaAjaxPesquisaLocalidades() {
var i, strAux, strResp, xmlDoc, oOption, oNodes;
	if (objAjaxPesqLocalidades.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxPesqLocalidades.responseText;
		if (strResp=="") {
			window.status="Concluído";
			divMsgAguarde.style.visibility="hidden";
			alert("Nenhuma localidade encontrada!!");
			return;
			}
		
		if (strResp!="") {
			try 
				{
				xmlDoc=objAjaxPesqLocalidades.responseXML.documentElement;
				for (i=0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
					oOption=document.createElement("OPTION");
					c_localidade_pesq.options.add(oOption);
					
					oNodes=xmlDoc.getElementsByTagName("localidade")[i];
					if (oNodes.childNodes.length > 0) strAux=oNodes.childNodes[0].nodeValue; else strAux="";
					if (strAux==null) strAux="";
					oOption.innerText=strAux;
					oOption.value=strAux;
					}
				}
			catch (e)
				{
				alert("Falha na consulta!!");
				}
			}
		window.status="Concluído";
		divMsgAguarde.style.visibility="hidden";
		c_localidade_pesq.focus();
		}
}

function CarregaLocalidades() {
var strUrl, strUF;
	objAjaxPesqLocalidades=GetXmlHttpObject();
	if (objAjaxPesqLocalidades==null) {
		alert("O browser NÃO possui suporte ao AJAX!!");
		return;
		}

//  Limpa lista de localidades
	LimpaListaLocalidades();
	LimpaTabelaResultado();
		
	strUF=trim(c_uf_pesq.value);
	if (strUF=="") {
		return;
		}
		
	window.status="Aguarde, pesquisando as localidades de " + c_uf_pesq.value + " ...";
	divMsgAguarde.style.visibility="";
		
	strUrl="AjaxCepLocalidadesPesqBD.asp";
	strUrl=strUrl+"?uf="+c_uf_pesq.value;
//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxPesqLocalidades.onreadystatechange=TrataRespostaAjaxPesquisaLocalidades;
	objAjaxPesqLocalidades.open("GET",strUrl,true);
	objAjaxPesqLocalidades.send(null);
}

function TrataRespostaAjaxPesquisaCEP() { 
var i, intQtdeLinhas, strAux, strResp, xmlDoc, oRow, oCell, oNodes;
	if (objAjaxPesqCep.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxPesqCep.responseText;
		if (strResp=="") {
			oRow=document.createElement("TR");
			oRow.style.backgroundColor="whitesmoke";
			oTBodyDados.appendChild(oRow);
			oCell=document.createElement("TD");
			strAux="<span class='N' style='font-size:14pt;font-weight:bold;color:red;'>Nenhum CEP encontrado</span>";
			oCell.colSpan=7;
			oCell.align="center";
			oCell.innerHTML=strAux;
			oRow.appendChild(oCell);
			window.status="Concluído";
			divMsgAguarde.style.visibility="hidden";
			return;
			}
		
		intQtdeLinhas=0;
		if (strResp!="") {
			try
				{
				xmlDoc=objAjaxPesqCep.responseXML.documentElement;
				for (i=0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
					intQtdeLinhas++;
					oRow=document.createElement("TR");
					oTBodyDados.appendChild(oRow);
					
					oCell=document.createElement("TD");
					strAux="<input type='RADIO' id='rb_check' name='rb_check'>";
					oCell.align="center";
					oCell.vAlign="top";
					oCell.innerHTML=strAux;
					oRow.appendChild(oCell);
					
					oCell=document.createElement("TD");
					oNodes=xmlDoc.getElementsByTagName("cep")[i];
					if (oNodes.childNodes.length > 0) strAux=oNodes.childNodes[0].nodeValue; else strAux="";
					if (strAux==null) strAux="";
					if (strAux=="") strAux="&nbsp;";
					oCell.noWrap=true;
					oCell.vAlign="top";
					oCell.innerHTML=strAux;
					oRow.appendChild(oCell);
					
					oCell=document.createElement("TD");
					oNodes=xmlDoc.getElementsByTagName("uf")[i];
					if (oNodes.childNodes.length > 0) strAux=oNodes.childNodes[0].nodeValue; else strAux="";
					if (strAux==null) strAux="";
					if (strAux=="") strAux="&nbsp;";
					oCell.noWrap=true;
					oCell.align="center";
					oCell.vAlign="top";
					oCell.innerHTML=strAux;
					oRow.appendChild(oCell);

					oCell=document.createElement("TD");
					oNodes=xmlDoc.getElementsByTagName("localidade")[i];
					if (oNodes.childNodes.length > 0) strAux=oNodes.childNodes[0].nodeValue; else strAux="";
					if (strAux==null) strAux="";
					if (strAux=="") strAux="&nbsp;";
					oCell.vAlign="top";
					oCell.innerHTML=strAux;
					oRow.appendChild(oCell);

					oCell=document.createElement("TD");
					oNodes=xmlDoc.getElementsByTagName("bairro")[i];
					if (oNodes.childNodes.length > 0) strAux=oNodes.childNodes[0].nodeValue; else strAux="";
					if (strAux==null) strAux="";
					if (strAux=="") strAux="&nbsp;";
					oCell.vAlign="top";
					oCell.innerHTML=strAux;
					oRow.appendChild(oCell);

					oCell=document.createElement("TD");
					oNodes=xmlDoc.getElementsByTagName("logradouro_nome")[i];
					if (oNodes.childNodes.length > 0) strAux=oNodes.childNodes[0].nodeValue; else strAux="";
					if (strAux==null) strAux="";
					if (strAux=="") strAux="&nbsp;";
					oCell.vAlign="top";
					oCell.innerHTML=strAux;
					oRow.appendChild(oCell);

					oCell=document.createElement("TD");
					oNodes=xmlDoc.getElementsByTagName("logradouro_complemento")[i];
					if (oNodes.childNodes.length > 0) strAux=oNodes.childNodes[0].nodeValue; else strAux="";
					if (strAux==null) strAux="";
					if (strAux=="") strAux="&nbsp;";
					oCell.vAlign="top";
					oCell.innerHTML=strAux;
					oRow.appendChild(oCell);
					}
				}
			catch (e)
				{
				alert("Falha na consulta!!");
				}
			}
		
		window.status = "Concluído";
		divMsgAguarde.style.visibility = "hidden";
		
	//  RETORNOU APENAS 1 REGISTRO?
		if (intQtdeLinhas == 1) {
			rb_check[1].checked = true;
			try {
				c_endereco_numero.focus();
				}
			catch (e) {
				// NOP
				}
			}
		}
}

function ExecutaPesquisaCEP(OpcaoPesquisaPor) {
var strUrl, strCep;
	objAjaxPesqCep=GetXmlHttpObject();
	if (objAjaxPesqCep==null) {
		alert("O browser NÃO possui suporte ao AJAX!!");
		return;
		}

	if (OpcaoPesquisaPor==OPCAO_PESQUISA_POR_CEP) {
		strCep=retorna_so_digitos(trim(c_cep_pesq.value));
		if ((strCep.length!=5)&&(strCep.length!=8)) {
			alert("CEP com tamanho inválido!!");
			c_cep_pesq.focus();
			return;
			}
		}
	else if (OpcaoPesquisaPor==OPCAO_PESQUISA_POR_ENDERECO) {
		if (trim(c_uf_pesq.value)=="") {
			alert("Informe a UF do endereço a ser pesquisado!!");
			c_uf_pesq.focus();
			return;
			}
		else if (!uf_ok(trim(c_uf_pesq.value))) {
			alert("UF inválida!!");
			c_uf_pesq.focus();
			return;
			}
		else if (trim(c_localidade_pesq.value)=="") {
			alert("Informe a localidade do endereço a ser pesquisado!!");
			c_localidade_pesq.focus();
			return;
			}
		}
	else {
		alert("Opção de pesquisa inválida!!");
		return;
		}
		
//  Limpa tabela com resultados
	LimpaTabelaResultado();

	window.status="Aguarde, pesquisando o CEP " + c_cep_pesq.value + " ...";
	divMsgAguarde.style.visibility="";
		
	strUrl="AjaxCepPesqBD.asp";
	if (OpcaoPesquisaPor==OPCAO_PESQUISA_POR_ENDERECO) {
		strUrl=strUrl+"?endereco="+c_endereco_pesq.value+"&uf="+c_uf_pesq.value+"&localidade="+c_localidade_pesq.value;
		}
	else {
		strUrl=strUrl+"?cep="+c_cep_pesq.value;
		}
	strUrl=strUrl+"&opcao="+OpcaoPesquisaPor;
//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxPesqCep.onreadystatechange=TrataRespostaAjaxPesquisaCEP;
	objAjaxPesqCep.open("GET",strUrl,true);
	objAjaxPesqCep.send(null);
}
</script>

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<body onload="IniciaPainel();">
<input type="hidden" name="ModoApenasConsulta" id="ModoApenasConsulta" value="<%=strModoApenasConsulta%>">
<input type="hidden" name="CepDefault" id="CepDefault" value="<%=strCepDefault%>">

<!-- MENSAGEM: "AGUARDE, PESQUISANDO NO BANCO DE DADOS" -->
<div id="divMsgAguarde" name="divMsgAguarde" align="center" style="position:absolute;left:10px;top:250px;width:935px;height:60px;z-index:9;border: 1pt solid #C0C0C0;background-color:lightyellow;visibility:hidden;">
	<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:15px;">
		<tr><td align="center">
			<table cellpadding="0" cellspacing="0">
				<tr>
				<td valign="bottom" align="right"><span style="color:orangered;font-weight:bold;font-style:italic;font-size:20pt;">Aguarde, processando requisição</span></td>
				<td style="width:20px;">&nbsp;</td>
				<td align="left"><img src="../imagem/aguarde.gif"border="0"></td>
				</tr>
			</table>
		</td></tr>
	</table>
</div>

<center>
<table>
	<tr><td><p class="PEDIDO">Pesquisa de CEP</p></td></tr>
</table>
</center>

<!-- ************   SEPARADOR   ************ -->
<table width="99%" cellPadding="0" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>

<br>

<!--  CAMPOS DE PESQUISA  -->
<span class="N">Pesquisa por CEP</span>
<div style="background:lightcyan;width:600px;align:center;border: 1pt solid gray;padding: 6 6 6 6;">
	<table border="0" cellspacing="0" cellpadding="0" style="margin-top:6px;">
	<tr>
		<td align="right">
			<span class="Cd">CEP</span>
		</td>
		<td nowrap>
			<input id="c_cep_pesq" name="c_cep_pesq" maxlength="9" size="11" value="<%=cep_formata(strCepDefault)%>" onkeypress="if (digitou_enter(true)) {bPesquisaCEP.click();} filtra_cep();" onblur="if (cep_ok(this.value)) this.value=cep_formata(this.value);">
			&nbsp;&nbsp;<span name="bPesquisaCEP" id="bPesquisaCEP" style='width:130px;font-size:10pt;' class="Botao" onclick="ExecutaPesquisaCEP(OPCAO_PESQUISA_POR_CEP);">Pesquisar</span>
		</td>
	</tr>
	</table>
</div>
<br>
<span class="N">Pesquisa por Endereço</span>
<div style="background:lightcyan;width:600px;align:center;border: 1pt solid gray;padding: 6 6 6 6;">
	<table border="0" cellspacing="0" cellpadding="4" style="margin-top:6px;">
	<tr>
		<td align="right" nowrap>
			<span class="Cd">UF</span>
		</td>
		<td nowrap>
			<select id="c_uf_pesq" name="c_uf_pesq" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;CarregaLocalidades();}" onchange="CarregaLocalidades();" onkeypress="if (digitou_enter(true)) c_localidade_pesq.focus();">
			<% =UF_monta_itens_select(Null) %>
			</select>
			&nbsp;
			<span class="Cd">Localidade</span>
			<select id="c_localidade_pesq" name="c_localidade_pesq" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" onkeypress="if (digitou_enter(true)) c_endereco_pesq.focus();">
			</select>
		</td>
	</tr>
	<tr>
		<td align="right" nowrap>
			<span class="Cd">Endereço</span>
		</td>
		<td nowrap>
			<input id="c_endereco_pesq" name="c_endereco_pesq" size="30" onkeypress="if (digitou_enter(true)) {bPesquisaEndereco.click();} filtra_nome_identificador();" onblur="this.value=trim(this.value);">
			&nbsp;&nbsp;<span name="bPesquisaEndereco" id="bPesquisaEndereco" style='width:130px;font-size:10pt;' class="Botao" onclick="ExecutaPesquisaCEP(OPCAO_PESQUISA_POR_ENDERECO);">Pesquisar</span>
		</td>
	</tr>
	</table>
</div>

<!--  DADOS DA RESPOSTA  -->
<!-- FORÇA A CRIAÇÃO DE UM ARRAY DE RADIO BUTTONS MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" id="rb_check" name="rb_check" value="">

<br>
<span class="N">Resultado</span>
<table id="tabResposta" border="1" cellspacing="0" cellpadding="4">
<thead bgcolor="lightyellow">
	<tr>
		<th align="left" style="font-size:9pt;width:30px;">&nbsp;</th>
		<th align="left" style="font-size:9pt;width:85px;">CEP</th>
		<th align="center" style="font-size:9pt;width:25px;">UF</th>
		<th align="left" style="font-size:9pt;width:130px;">Cidade</th>
		<th align="left" style="font-size:9pt;width:150px;">Bairro</th>
		<th align="left" style="font-size:9pt;width:220px;">Logradouro</th>
		<th align="left" style="font-size:9pt;width:220px;">Complemento</th>
	</tr>
</thead>
<tbody id="oTBodyDados">
</tbody>
</table>

<br>

<!-- ************   SEPARADOR   ************ -->
<table width="99%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<center>
<% if strModoApenasConsulta = "S" then %>
<input type="hidden" name="c_endereco_numero" id="c_endereco_numero" value="" />
<table>
	<tr>
		<td align="center">
			<span name="bFechar" id="bFechar" style='width:130px;font-size:12pt;' class="Botao" onclick="CancelarOperacao();">Fechar</span>
		</td>
	</tr>
</table>
<% else %>
<!-- ************   Nº / COMPLEMENTO   ************ -->
<table border="0" cellspacing="0" cellpadding="0" style="margin-top:6px;">
<tr>
	<td align="right">
		<span class="Cd">Nº</span>
	</td>
	<td nowrap>
		<input id="c_endereco_numero" name="c_endereco_numero" maxlength="20" size="25" value="" onkeypress="if (digitou_enter(true)) c_endereco_complemento.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);">
	</td>
	<td style="width:20px;">
		&nbsp;
	</td>
	<td align="right">
		<span class="Cd">Complemento</span>
	</td>
	<td nowrap>
		<input id="c_endereco_complemento" name="c_endereco_complemento" maxlength="60" size="45" value="" onkeypress="if (digitou_enter(true)) bConfirmar.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);">
	</td>
</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="99%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<table>
	<tr>
		<td align="center">
			<span name="bCancelar" id="bCancelar" style='width:130px;font-size:12pt;' class="Botao" onclick="CancelarOperacao();">Cancelar</span>
		</td>
		<td style="width:20px">&nbsp;</td>
		<td align="center">
			<span name="bConfirmar" id="bConfirmar" style='width:130px;font-size:12pt;' class="Botao" onclick="ConfirmarOperacao();">Confirmar</span>
		</td>
	</tr>
</table>
<% end if %>
</center>

</body>
</html>
