<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelFarolProdutoCompradoExec.asp
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

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_FAROL_CADASTRO_PRODUTO_COMPRADO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	alerta = ""

	dim s, s_filtro, c_filtro_fabricante
	dim s_nome_fabricante

	c_filtro_fabricante = retorna_so_digitos(Request.Form("c_filtro_fabricante"))
	if c_filtro_fabricante <> "" then c_filtro_fabricante = normaliza_codigo(c_filtro_fabricante, TAM_MIN_FABRICANTE)
	
	if c_filtro_fabricante <> "" then
		s_nome_fabricante = fabricante_descricao(c_filtro_fabricante)
	else
		s_nome_fabricante = ""
		end if

	dim intQtdeProdutos
	intQtdeProdutos = 0




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' EXECUTA CONSULTA
'
sub executa_consulta
dim r
dim s, s_aux, s_class_aux, s_sql, s_sql_lista_base, x, cab_table, cab, fabricante_a, msg_erro
dim n, n_reg, qtde_fabricantes
dim n_sub_total_qtde_comprada, n_total_qtde_comprada
dim strJsScriptTotal
dim strJsScriptQtde
	
	strJsScriptTotal = ""
	strJsScriptQtde = ""

'	MONTA O SQL QUE SELECIONA A RELAÇÃO DE PRODUTOS
'	A LÓGICA CONSISTE EM SELECIONAR:
'		1) PRODUTOS QUE TENHAM SALDO NO ESTOQUE DE VENDA
'		2) PRODUTOS QUE CONSTEM COMO 'VENDÁVEIS'
'		3) PRODUTOS QUE TENHAM O CAMPO "farol_qtde_comprada" C/ VALOR MAIOR QUE ZERO
'	OBS: O USO DE 'UNION' SIMPLES ELIMINA AS LINHAS DUPLICADAS DOS RESULTADOS
'		 O USO DE 'UNION ALL' RETORNARIA TODAS AS LINHAS, INCLUSIVE AS DUPLICADAS
	s_sql_lista_base = _
		"SELECT DISTINCT" & _
			" fabricante," & _
			" produto" & _
		" FROM t_ESTOQUE_ITEM" & _
		" WHERE" & _
			" ((qtde - qtde_utilizada) > 0)"
			
	if c_filtro_fabricante <> "" then
		s_sql_lista_base = s_sql_lista_base & " AND (fabricante='" & c_filtro_fabricante & "')" 
		end if
	
	s_sql_lista_base = s_sql_lista_base & _
		" UNION " & _
		"SELECT DISTINCT" & _
			" t_PRODUTO.fabricante," & _
			" t_PRODUTO.produto" & _
		" FROM t_PRODUTO" & _
			" INNER JOIN" & _
				"(" & _
					"SELECT DISTINCT" & _
						" fabricante," & _
						" produto" & _
					" FROM t_PRODUTO_LOJA" & _
					" WHERE" & _
						" (vendavel = 'S')" & _
				") tPL_AUX ON (t_PRODUTO.fabricante=tPL_AUX.fabricante) AND (t_PRODUTO.produto=tPL_AUX.produto)" & _
		" WHERE" & _
			" (excluido_status = 0)"

	if c_filtro_fabricante <> "" then
		s_sql_lista_base = s_sql_lista_base & " AND (t_PRODUTO.fabricante='" & c_filtro_fabricante & "')" 
		end if
	
	s_sql_lista_base = s_sql_lista_base & _
		" UNION " & _
		" SELECT DISTINCT" & _
			" fabricante," & _
			" produto" & _
		" FROM t_PRODUTO" & _
		" WHERE" & _
			" (farol_qtde_comprada > 0)"

	if c_filtro_fabricante <> "" then
		s_sql_lista_base = s_sql_lista_base & " AND (fabricante='" & c_filtro_fabricante & "')" 
		end if


'	SELECT COMPLETO
	s_sql = _
		"SELECT" & _
			" t_PROD_LISTA_BASE.fabricante," & _
			" t_PROD_LISTA_BASE.produto," & _
			" t_PRODUTO.descricao," & _
			" t_PRODUTO.descricao_html," & _
			" t_PRODUTO.farol_qtde_comprada" & _
		" FROM (" & s_sql_lista_base & ") t_PROD_LISTA_BASE" & _
			" LEFT JOIN t_PRODUTO ON (t_PROD_LISTA_BASE.fabricante = t_PRODUTO.fabricante) AND (t_PROD_LISTA_BASE.produto = t_PRODUTO.produto)" & _
		" WHERE" & _
			" (descricao <> '.')" & _
			" AND (descricao <> '*')" & _
		" ORDER BY" & _
			" fabricante," & _
			" produto"


  ' CABEÇALHO
	cab_table = "<table cellspacing='0'>" & chr(13)
	cab = "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td valign='bottom' align='left' style='background:white;' nowrap>&nbsp;</td>" & chr(13) & _
		  "		<td valign='bottom' align='left' nowrap class='MT wCod'><span class='R'>Código</span></td>" & chr(13) & _
		  "		<td valign='bottom' align='left' nowrap class='MTBD wDescr'><span class='R'>Produto</span></td>" & chr(13) & _
		  "		<td valign='bottom' align='right' nowrap class='MTBD wQtde'><span class='Rd' style='font-weight:bold;'>Qtde</span><br /><span class='Rd' style='font-weight:bold;'>Comprada</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	n_reg = 0
	qtde_fabricantes = 0
	fabricante_a = "XXXXX"

	n_sub_total_qtde_comprada = 0
	n_total_qtde_comprada = 0

	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then
				strJsScriptTotal = strJsScriptTotal & _
					"vFabricante[""F" & fabricante_a & """]=" & Cstr(n_sub_total_qtde_comprada) & ";" & chr(13)
					
				x = x & "	<tr style='background:#FFFFDD;' nowrap>" & chr(13) & _
						"		<td align='left' style='background:#FFFFFF;'>&nbsp;</td>" & chr(13) & _
						"		<td align='right' colspan='2' class='MEB'><span class='Cd'>Total:</span></td>" & chr(13) & _
						"		<td align='left' class='MDB'>" & chr(13) & _
						"			<input type='text' id='c_sub_total_" & fabricante_a & "' name='c_sub_total_" & fabricante_a & "' class='PLLd cQtde' readonly tabindex=-1 value='" & formata_inteiro(n_sub_total_qtde_comprada) & "' />" & chr(13) & _
						"		</td>" & chr(13) & _
						"	</tr>" & chr(13) & _
						"</table>" & chr(13)
				Response.Write x
				x = "<br>" & chr(13) & _
					"<br>" & chr(13)
				end if

			x = x & cab_table
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<tr nowrap>" & chr(13) & _
					"		<td align='left' style='background:#FFFFFF;'>&nbsp;</td>" & chr(13) & _
					"		<td class='MDTE' align='center' colspan='3' style='background:azure;'><span class='C'>&nbsp;" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					cab
			n_sub_total_qtde_comprada = 0
			end if
		
	  ' CONTAGEM
		n_reg = n_reg + 1
		intQtdeProdutos = intQtdeProdutos + 1

		x = x & "	<tr id='TR_" & Cstr(n_reg) & "'>" & chr(13)

	 '> Nº LINHA
		x = x & "		<td class='NW dir' align='left'><span class='Rd pIdx'>" & Cstr(n_reg) & ".</span></td>" & chr(13)

	 '> CÓDIGO DO PRODUTO
		x = x & "		<td class='MDBE NW wCod' align='left'><span class='C pProd'>" & Trim("" & r("produto")) & "</span></td>" & chr(13)

	 '> DESCRIÇÃO DO PRODUTO
		x = x & "		<td class='MDB NW wDescr' align='left'><span class='C pDescr'>" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</span></td>" & chr(13)

	 '> QTDE COMPRADA
		n = r("farol_qtde_comprada")
		if n = 0 then s_class_aux = " CorQtdeZero" else s_class_aux = " CorQtdePositiva"
		x = x & "		<td class='MDB NW wQtde' align='left'>" & chr(13) & _
				"			<input type='hidden' id='c_fabr' name='c_fabr' value='" & Trim("" & r("fabricante")) & "' />" & chr(13) & _
				"			<input type='hidden' id='c_prod' name='c_prod' value='" & Trim("" & r("produto")) & "' />" & chr(13) & _
				"			<input type='hidden' id='c_qtde_original' name='c_qtde_original' value='" & Cstr(n) & "' />" & chr(13) & _
				"			<input type='text' id='c_qtde' name='c_qtde' class='PLLd cQtde" & s_class_aux & "' maxlength='6'" & _
								" value='" & formata_inteiro(n) & "'" & _
								" onfocus='trata_onfocus(this," & Cstr(n_reg) & ");'" & _
								" onkeydown='trata_onkeydown(this," & Cstr(n_reg) & ");'" & _
								" onkeypress='trata_onkeypress(this," & Cstr(n_reg) & ");'" & _
								" onblur='trata_onblur(this," & Cstr(n_reg) & ");'" & _
								" />" & chr(13) & _
				"		</td>" & chr(13)
	
		strJsScriptQtde = strJsScriptQtde & _
			"	vQtde[" & Cstr(n_reg) & "]=" & Cstr(n) & ";" & chr(13)
		
	'	TOTALIZAÇÃO
		n_sub_total_qtde_comprada = n_sub_total_qtde_comprada + n
		n_total_qtde_comprada = n_total_qtde_comprada + n
		
		x = x & "	</tr>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
	
  ' MOSTRA TOTAL DO ÚLTIMO FABRICANTE
	if n_reg <> 0 then 
		strJsScriptTotal = strJsScriptTotal & _
			"vFabricante[""F" & fabricante_a & """]=" & Cstr(n_sub_total_qtde_comprada) & ";" & chr(13)
			
		x = x & "	<tr style='background: #FFFFDD' nowrap>"  & chr(13) & _
				"		<td align='left' style='background:white;'>&nbsp;</td>" & chr(13) & _
				"		<td align='right' class='MEB' colspan='2' nowrap><span class='Cd'>" & "Total:" & "</span></td>" & chr(13) & _
				"		<td align='left' class='MDB' nowrap>" & chr(13) & _
				"			<input type='text' id='c_sub_total_" & fabricante_a & "' name='c_sub_total_" & fabricante_a & "' class='PLLd cQtde' readonly tabindex=-1 value='" & formata_inteiro(n_sub_total_qtde_comprada) & "' />" & chr(13) & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13)
	'>	TOTAL GERAL
		if qtde_fabricantes > 1 then
			x = x & _
				"	<tr><td colspan='4'>&nbsp;</td></tr>" & chr(13) & _
				"	<tr><td colspan='4'>&nbsp;</td></tr>" & chr(13) & _
				"	<tr style='background:honeydew'>" & chr(13) & _
				"		<td align='left' style='background:white;'>&nbsp;</td>" & chr(13) & _
				"		<td class='MTB ME' colspan='2' align='left' valign='bottom' nowrap><p class='Cd'>TOTAL GERAL:</p></td>" & chr(13) & _
				"		<td class='MTBD' align='left' valign='bottom' nowrap>" & chr(13) & _
				"			<input type='text' id='c_total' name='c_total' class='PLLd cQtde' readonly tabindex=-1 value='" & formata_inteiro(n_total_qtde_comprada) & "' />" & chr(13) & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab & _
			"	<tr nowrap>" & chr(13) & _
			"		<td align='left' style='background:white;'>&nbsp;</td>" & chr(13) & _
			"		<TD align='center' colspan='3' class='MDBE ALERTA'><span class='ALERTA'>&nbsp;Nenhum produto satisfaz as condições especificadas&nbsp;</span></td>" & chr(13) & _
			"	</tr>" & chr(13)
		end if

  ' FECHA TABELA DA ÚLTIMA LOJA
	x = x & "</table>" & chr(13)
	
	Response.write x

	if strJsScriptTotal <> "" then
		strJsScriptTotal = _
			"<script language='JavaScript'>" & chr(13) & _
			strJsScriptTotal & _
			"</script>" & chr(13)
		Response.write strJsScriptTotal
		end if
	
	if strJsScriptQtde <> "" then
		strJsScriptQtde = _
			"<script language='JavaScript'>" & chr(13) & _
			strJsScriptQtde & _
			"</script>" & chr(13)
		Response.write strJsScriptQtde
		end if
	
	if r.State <> 0 then r.Close
	set r=nothing
	
end sub
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
<script src="../Global/jquery.js" Language="JavaScript" Type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

var vFabricante = new Array();
var vQtde = new Array();
var qtde_anterior;

function zerarTudo() {
	$(".cQtde").val("0");
	$(".cQtde").removeClass("CorQtdePositiva");
	$(".cQtde").addClass("corQtdeZero");
	for (i = 0; i < vQtde.length; i++) {
		vQtde[i] = 0;
	}
	for (x in vFabricante) {
		vFabricante[x] = 0;
	}
	atualiza_total_geral();
}

function trata_onfocus(c, indice) {
	qtde_anterior = vQtde[indice];
	c.select();
	realca(c, indice);
}

function trata_onkeydown(c, indice) {
var f;
	f = fREL;
	if (window.event.keyCode == 38) {
		// KEY UP
		if ((indice - 1) > 0) f.c_qtde[indice - 1].focus();
	}
	else if (window.event.keyCode == 40) {
		// KEY DOWN
		// LEMBRANDO QUE O 1º CAMPO DO VETOR É APENAS P/ ASSEGURAR A CRIAÇÃO DE UM ARRAY NO CASO DE HAVER UM ÚNICO PRODUTO
		if ((indice + 1) < f.c_qtde.length) f.c_qtde[indice + 1].focus();
	}
}

function trata_onkeypress(c, indice) {
var f;
	f=fREL;
	if (digitou_enter(true)) { if ((indice + 1) < f.c_qtde.length) f.c_qtde[indice + 1].focus(); }
	filtra_numerico();
}

function trata_onblur(c, indice) {
var n;
	normaliza(c, indice);
	n = converte_numero(retorna_so_digitos(c.value));
	if (n == qtde_anterior) return;
	vQtde[indice] = n;
	c.value = formata_inteiro(n);
	totaliza(indice);
	if (n == 0) {
		$(c).removeClass("CorQtdePositiva");
		$(c).addClass("corQtdeZero");
	}
	else {
		$(c).removeClass("corQtdeZero");
		$(c).addClass("CorQtdePositiva");
	}
}

function realca(c, indice_row) {
	$("#TR_" + indice_row.toString()).addClass("Realcado");
}

function normaliza(c, indice_row) {
	$("#TR_" + indice_row.toString()).removeClass("Realcado");
}

function atualiza_total_geral() {
var t = 0;
	for (x in vFabricante) {
		t += vFabricante[x];
	}
	$("#c_total").val(formata_inteiro(t));
}

function totaliza(indice) {
var f;
var fabricante;
var sub_total = 0;
	f = fREL;
	fabricante = f.c_fabr[indice].value;
	// A LISTAGEM ESTÁ ORDENADA: COMEÇA SOMANDO A PARTIR DA PRÓPRIA LINHA ATÉ CHEGAR AO FINAL OU MUDAR DE FABRICANTE
	// LEMBRANDO QUE O VETOR vQtde ESTÁ ALINHADO COM O ARRAY DE CAMPOS c_qtde
	for (i = indice; i < f.c_fabr.length; i++) {
		if (f.c_fabr[i].value != fabricante) break;
		sub_total += vQtde[i];
	}
	// A LISTAGEM ESTÁ ORDENADA: COMEÇA SOMANDO A PARTIR DA LINHA ANTERIOR ATÉ CHEGAR AO COMEÇO OU MUDAR DE FABRICANTE
	// LEMBRANDO QUE O 1º CAMPO DO VETOR É APENAS P/ ASSEGURAR A CRIAÇÃO DE UM ARRAY NO CASO DE HAVER UM ÚNICO PRODUTO
	for (i = indice - 1; i > 0; i--) {
		if (f.c_fabr[i].value != fabricante) break;
		sub_total += vQtde[i];
	}
	$("#c_sub_total_" + fabricante).val(formata_inteiro(sub_total));
	vFabricante["F" + fabricante] = sub_total;
	atualiza_total_geral();
}

function fRELGravaDados() {
	fREL.action = "RelFarolProdutoCompradoGravaDados.asp";
	dCONFIRMA.style.visibility = "hidden";
	window.status = "Aguarde ...";
	fREL.submit();
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style type="text/css">
P.C { font-size:10pt; }
P.Cc { font-size:10pt; }
P.Cd { font-size:10pt; }
P.F { font-size:11pt; }
.NW
{
	white-space:nowrap;
}
.dir
{
	text-align:right;
}
.wCod
{
	width:70px;
}
.wDescr
{
	width:315px;
}
.wQtde
{
	width:80px;
}
.cQtde
{
	width:75px;
	font-size:11pt;
	font-weight:bold;
	background-color:transparent;
}
.pIdx
{
	margin-right:2px;
}
.pProd
{
	margin-left:2px;
}
.pDescr
{
	margin-left:2px;
}
.CorQtdeZero
{
	color:#696969;
	font-weight:normal;
}
.CorQtdePositiva
{
	color:Blue;
	font-weight:bold;
}
.Realcado
{
	background-color:#98FB98;
}
</style>


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
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>



<% else %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="c_filtro_fabricante" id="c_filtro_fabricante" value="<%=c_filtro_fabricante%>">

<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" id="c_fabr" name="c_fabr" value="" />
<input type="hidden" id="c_prod" name="c_prod" value="" />
<input type="hidden" id="c_qtde" name="c_qtde" value="" />
<input type="hidden" id="c_qtde_original" name="c_qtde_original" value="" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Produtos Comprados (Farol)</span>
	<br>
	<%	s = "<span class='N'>Emissão:&nbsp;" & formata_data_hora(Now) & "</span>"
		Response.Write s
	%>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellpadding='0' cellspacing='4' style='border-bottom:1px solid black;' border='0'>" & chr(13)
	s = Trim(c_filtro_fabricante)
	if s = "" then
		s = "TODOS"
	else
		if s_nome_fabricante <> "" then s = s & " - " & s_nome_fabricante 
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Fabricante:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>


<!--  RELATÓRIO  -->
<br>
<%	
	executa_consulta
%>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width='649' cellpadding='0' cellspacing='0' border='0' style="margin-top:0px;">
<tr>
	<td width="75%" align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkZerarTudo" href="javascript:zerarTudo();"><p class="Button BtnAll" style="margin-bottom:0px;">Zerar Tudo</p></a></td>
</tr>
</table>

<br />
<table class="notPrint" width="649" cellspacing="0">
<tr>
	<% if intQtdeProdutos > 0 then %>
	<td align="left">
		<a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
			<img src="../botao/voltar.gif" width="176" height="55" border="0">
		</a>
	</td>
	<td align="right">
		<div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0" /></a></div>
	</td>
	<% else %>
	<td align="center">
		<a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
			<img src="../botao/voltar.gif" width="176" height="55" border="0">
		</a>
	</td>
	<% end if %>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
