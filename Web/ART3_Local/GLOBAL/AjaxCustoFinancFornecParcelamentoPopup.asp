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

	dim intIndiceOpcao, intIndiceRow
	dim strSql, strQtdeParcelas, strPrecoLista, vl_lista
	dim c_fabricante, c_produto, c_loja, c_custoFinancFornecTipoParcelamento, c_custoFinancFornecQtdeParcelas
	c_fabricante=Trim(Request("fabricante"))
	c_produto=Trim(Request("produto"))
	c_loja=Trim(Request("loja"))
	c_custoFinancFornecTipoParcelamento=Trim(Request("tipoParcelamento"))
	c_custoFinancFornecQtdeParcelas=Trim(Request("qtdeParcelas"))

	dim alerta
	alerta = ""

	if c_fabricante = "" then
		alerta = "Código do fabricante do produto não foi informado."
	elseif c_produto = "" then
		alerta = "Código do produto não foi informado."
	elseif c_loja = "" then
		alerta = "Número da loja que está efetuando a venda não foi informado."
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	if alerta = "" then
		strSql = _
			"SELECT " & _
				"*" & _
			" FROM t_PRODUTO" & _
				" INNER JOIN t_PRODUTO_LOJA" & _
					" ON (t_PRODUTO.fabricante=t_PRODUTO_LOJA.fabricante) AND (t_PRODUTO.produto=t_PRODUTO_LOJA.produto)" & _
			" WHERE" & _
				" (t_PRODUTO.fabricante = '" & c_fabricante & "')" & _
				" AND (t_PRODUTO.produto = '" & c_produto & "')" & _
				" AND (CONVERT(smallint,loja) = " & c_loja & ")"
		rs.open strSql, cn
		if rs.Eof then
			alerta = "Falha ao consultar o preço do produto " & c_produto & " para a loja " & c_loja & "."
		else
			vl_lista = rs("preco_lista")
			end if
		end if
	




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S
' _____________________________________________________________________________________________


%>

<%=DOCTYPE_LEGADO%>

<html>
<head>
	<title>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Tabela de Preços
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</title>
</head>

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function cancelarOperacao() {
	window.close();
}

function realcaCorLinha(linhaTabela) {
	// Realça a cor da linha selecionada
	linhaTabela.style.background='yellow';
}

function normalizaCorLinha(linhaTabela) {
	linhaTabela.style.background='white';
}

function normalizaCorLinhaTodos() {
	$(".TR_OPCAO").css("background-color", "white");
}

function realcaOpcaoSelecionadaDefault() {
var f, i;
	f=fParcelamento;
//	LEMBRE-SE: O ARRAY DE CAMPOS 'rbOpcaoParcelamento' TEM O 1º CAMPO GERADO POR UM 'INPUT HIDDEN'
//  =========  C/ A FUNÇÃO DE SEMPRE GERAR UM ARRAY, MESMO NO CASO DA TABELA TER APENAS 1 LINHA.
//			   PORTANTO, SEMPRE HAVERÁ 1 rbOpcaoParcelamento A MAIS QUE O TOTAL DE OPÇÕES VÁLIDAS E
//			   A 1ª LINHA VÁLIDA CORRESPONDE AO rbOpcaoParcelamento[1] E NÃO AO rbOpcaoParcelamento[0]
	for (i=1; i<f.rbOpcaoParcelamento.length; i++) {
		if (f.rbOpcaoParcelamento[i].checked) {
			f.rbOpcaoParcelamento[i].click();
			break;
			}
		}
}

function confirmarOperacao() {
var f, i, strOpcaoSelecionada, v, strFabricante, strProduto;
var strTipoParcelamento, strPrecoLista;
var intQtdeParcelas=0;
	f=fParcelamento;
//	LEMBRE-SE: O ARRAY DE CAMPOS 'rbOpcaoParcelamento' TEM O 1º CAMPO GERADO POR UM 'INPUT HIDDEN'
//  =========  C/ A FUNÇÃO DE SEMPRE GERAR UM ARRAY, MESMO NO CASO DA TABELA TER APENAS 1 LINHA.
//			   PORTANTO, SEMPRE HAVERÁ 1 rbOpcaoParcelamento A MAIS QUE O TOTAL DE OPÇÕES VÁLIDAS E
//			   A 1ª LINHA VÁLIDA CORRESPONDE AO rbOpcaoParcelamento[1] E NÃO AO rbOpcaoParcelamento[0]
	strOpcaoSelecionada="";
	for (i=1; i<f.rbOpcaoParcelamento.length; i++) {
		if (f.rbOpcaoParcelamento[i].checked) {
			strOpcaoSelecionada = f.rbOpcaoParcelamento[i].value;
			break;
			}
		}

	if (strOpcaoSelecionada == "") {
		alert("Selecione uma das opções!!");
		return;
		}
		
	v=strOpcaoSelecionada.split("|");
	strTipoParcelamento = v[0];
	strPrecoLista = v[1];
	if ((strTipoParcelamento==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA)||(strTipoParcelamento==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA)) {
		intQtdeParcelas=converte_numero(v[2]);
		}

	strFabricante = f.c_fabricante.value;
	strProduto = f.c_produto.value;
	try {
	//  A JANELA 'OPENER' JÁ PODE TER SIDO FECHADA
		window.opener.processaSelecaoCustoFinancFornecParcelamento(strTipoParcelamento, strPrecoLista, intQtdeParcelas, strFabricante, strProduto);
		}
	catch (e) {
		alert(e.message);
		}
	window.close();
}
</script>


<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center">
		<span name="bCancelar" id="bCancelar" style='width:130px;font-size:12pt;' class="Botao" onclick="cancelarOperacao();">Fechar</span>
	</td>
</tr>
</table>
</center>
</body>



<% else %>
<body onload="realcaOpcaoSelecionadaDefault();">
<center>
<table>
	<tr><td align="center"><span class="PEDIDO">Tabela de Preços</span></td></tr>
	<tr><td align="center"><span class="N" style="font-size:12pt;">Produto: <%=c_produto & " - " & produto_formata_descricao_em_html(produto_descricao_html(c_fabricante, c_produto))%></span></td></tr>
</table>
</center>

<!-- ************   SEPARADOR   ************ -->
<table width="99%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<br>

<center>

<form id="fParcelamento" name="fParcelamento" method="post">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">
<input type="hidden" name="c_custoFinancFornecTipoParcelamento" id="c_custoFinancFornecTipoParcelamento" value="<%=c_custoFinancFornecTipoParcelamento%>">
<input type="hidden" name="c_custoFinancFornecQtdeParcelas" id="c_custoFinancFornecQtdeParcelas" value="<%=c_custoFinancFornecQtdeParcelas%>">

<!-- FORÇA A CRIAÇÃO DE UM ARRAY DE RADIO BUTTONS MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<%intIndiceOpcao=0%>
<input type="hidden" id="rbOpcaoParcelamento" name="rbOpcaoParcelamento" value="">

<table>
<tr>
<!-- *************************************************************************** -->
<!-- ******************************    À VISTA    ****************************** -->
<!-- *************************************************************************** -->
<td valign="top" align="center">
	<span class="N" style="font-size:12pt;color:#008000">À Vista</span>
	<table id="tabAVista" border="1" cellspacing="0" cellpadding="0">
	<thead bgcolor="azure">
		<tr>
			<th align="center" style="font-size:9pt;width:25px;">&nbsp;</th>
			<th align="right" style="font-size:9pt;width:80px;">Preço (<%=SIMBOLO_MONETARIO%>)</th>
		</tr>
	</thead>
	<tbody id="oTBodyAVista">
		<tr class="TR_OPCAO">
			<% strPrecoLista=formata_moeda(vl_lista) %>
			<td align="center">
				<%intIndiceOpcao=intIndiceOpcao+1%>
				<input type="radio" id='rbOpcaoParcelamento' name='rbOpcaoParcelamento' 
					value='<%=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA & "|" & strPrecoLista%>'
			<%if c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA then Response.Write " checked"%>
					onclick="normalizaCorLinhaTodos();realcaCorLinha(oTBodyAVista.getElementsByTagName('tr')[0]);"
					/>
			</td>
			<td align="left">
				<span class="Cd" style="cursor:default;" onclick="fParcelamento.rbOpcaoParcelamento[<%=intIndiceOpcao%>].click();"><%=strPrecoLista%></span>
			</td>
		</tr>
	</tbody>
	</table>
</td>
<td align="left"><span style="width:30px;"></span></td>
<!-- *************************************************************************** -->
<!-- ******************************  COM ENTRADA  ****************************** -->
<!-- *************************************************************************** -->
<td valign="top" align="center">
	<span class="N" style="font-size:12pt;color:#008000">Com Entrada</span><br>
	<table id="tabComEntrada" border="1" cellspacing="0" cellpadding="0">
	<thead bgcolor="azure">
		<tr>
			<th align="center" style="font-size:9pt;width:25px;">&nbsp;</th>
			<th align="center" style="font-size:9pt;width:60px;">Parcelas</th>
			<th align="right" style="font-size:9pt;width:80px;">Preço (<%=SIMBOLO_MONETARIO%>)</th>
		</tr>
	</thead>
	<tbody id="oTBodyComEntrada">
	<%
		strSql = _
			"SELECT " & _
				"*" & _
			" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
			" WHERE" & _
				" (fabricante = '" & c_fabricante & "')" & _
				" AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA & "')" & _
			" ORDER BY" & _
				" qtde_parcelas"
		if rs.State <> 0 then rs.Close
		rs.open strSql, cn
	%>
	
	<%if rs.Eof then%>
		<tr>
		<td colspan="3" align="left">
		<span class="Cc" style="color:red;font-weight:bold;margin-top:3px;">NÃO DISPONÍVEL</span>
		</td>
		</tr>
		<%end if%>

	<%
		intIndiceRow = -1
		do while Not rs.Eof
			intIndiceRow = intIndiceRow + 1
	%>
		<tr class="TR_OPCAO">
			<% 
				strQtdeParcelas=Trim("" & rs("qtde_parcelas"))
				strPrecoLista=formata_moeda(rs("coeficiente")*vl_lista) 
			%>
			<td align="center">
				<%intIndiceOpcao=intIndiceOpcao+1%>
				<input type="radio" id='rbOpcaoParcelamento' name='rbOpcaoParcelamento'
					value='<%=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA & "|" & strPrecoLista & "|" & strQtdeParcelas%>'
			<%if (c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) And (converte_numero(c_custoFinancFornecQtdeParcelas)=converte_numero(strQtdeParcelas)) then Response.Write " checked"%>
					onclick="normalizaCorLinhaTodos();realcaCorLinha(oTBodyComEntrada.getElementsByTagName('tr')[<%=intIndiceRow%>]);"
					/>
			</td>
			<td align="left">
				<span class="Cc" style="cursor:default;" onclick="fParcelamento.rbOpcaoParcelamento[<%=intIndiceOpcao%>].click();">1 + <%=strQtdeParcelas%></span>
			</td>
			<td align="left">
				<span class="Cd" style="cursor:default;" onclick="fParcelamento.rbOpcaoParcelamento[<%=intIndiceOpcao%>].click();"><%=strPrecoLista%></span>
			</td>
		</tr>
	<%	
		rs.MoveNext
		loop
	%>
	</tbody>
	</table>
</td>
<td align="left"><span style="width:30px;"></span></td>
<!-- *************************************************************************** -->
<!-- ******************************  SEM ENTRADA  ****************************** -->
<!-- *************************************************************************** -->
<td valign="top" align="center">
	<span class="N" style="font-size:12pt;color:#008000">Sem Entrada</span><br>
	<table id="tabSemEntrada" border="1" cellspacing="0" cellpadding="0">
	<thead bgcolor="azure">
		<tr>
			<th align="center" style="font-size:9pt;width:25px;">&nbsp;</th>
			<th align="center" style="font-size:9pt;width:60px;">Parcelas</th>
			<th align="right" style="font-size:9pt;width:80px;">Preço (<%=SIMBOLO_MONETARIO%>)</th>
		</tr>
	</thead>
	<tbody id="oTBodySemEntrada">
	<%
		strSql = _
			"SELECT " & _
				"*" & _
			" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
			" WHERE" & _
				" (fabricante = '" & c_fabricante & "')" & _
				" AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA & "')" & _
			" ORDER BY" & _
				" qtde_parcelas"
		if rs.State <> 0 then rs.Close
		rs.open strSql, cn
	%>

	<%if rs.Eof then%>
		<tr>
		<td colspan="3" align="left">
		<span class="Cc" style="color:red;font-weight:bold;margin-top:3px;">NÃO DISPONÍVEL</span>
		</td>
		</tr>
		<%end if%>

	<%
		intIndiceRow = -1
		do while Not rs.Eof
			intIndiceRow = intIndiceRow + 1
	%>
		<tr class="TR_OPCAO">
			<% 
				strQtdeParcelas=Trim("" & rs("qtde_parcelas"))
				strPrecoLista=formata_moeda(rs("coeficiente")*vl_lista) 
			%>
			<td align="center">
				<%intIndiceOpcao=intIndiceOpcao+1%>
				<input type="radio" id='rbOpcaoParcelamento' name='rbOpcaoParcelamento'
					value='<%=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA & "|" & strPrecoLista & "|" & strQtdeParcelas%>'
			<%if (c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) And (converte_numero(c_custoFinancFornecQtdeParcelas)=converte_numero(strQtdeParcelas)) then Response.Write " checked"%>
					onclick="normalizaCorLinhaTodos();realcaCorLinha(oTBodySemEntrada.getElementsByTagName('tr')[<%=intIndiceRow%>]);"
					/>
			</td>
			<td align="left">
				<span class="Cc" style="cursor:default;" onclick="fParcelamento.rbOpcaoParcelamento[<%=intIndiceOpcao%>].click();">0 + <%=strQtdeParcelas%></span>
			</td>
			<td align="left">
				<span class="Cd" style="cursor:default;" onclick="fParcelamento.rbOpcaoParcelamento[<%=intIndiceOpcao%>].click();"><%=strPrecoLista%></span>
			</td>
		</tr>
	<%	
		rs.MoveNext
		loop
	%>
	</tbody>
	</table>
</td>
</tr>
</table>
</form>

</center>

<!-- ************   SEPARADOR   ************ -->
<table width="99%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>

<center>
<table>
	<tr>
		<td align="center">
			<span name="bCancelar" id="bCancelar" style='width:130px;font-size:12pt;' class="Botao" onclick="cancelarOperacao();">Cancelar</span>
		</td>
		<td style="width:20px" align="left">&nbsp;</td>
		<td align="center">
			<span name="bConfirmar" id="bConfirmar" style='width:130px;font-size:12pt;' class="Botao" onclick="confirmarOperacao();">Confirmar</span>
		</td>
	</tr>
</table>

</center>

</body>
<% end if %>

</html>

<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
