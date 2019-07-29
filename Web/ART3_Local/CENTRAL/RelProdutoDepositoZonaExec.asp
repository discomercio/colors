<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelProdutoDepositoZonaExec.asp
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
	if Not operacao_permitida(OP_CEN_REL_PRODUTO_DEPOSITO_ZONA, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	alerta = ""

	dim s, s_filtro, c_fabricante
	dim s_nome_fabricante

	c_fabricante = retorna_so_digitos(Request.Form("c_fabricante"))
	if c_fabricante <> "" then c_fabricante = normaliza_codigo(c_fabricante, TAM_MIN_FABRICANTE)
	
	if c_fabricante <> "" then
		s_nome_fabricante = fabricante_descricao(c_fabricante)
	else
		s_nome_fabricante = ""
		end if

	dim intQtdeProdutos
	intQtdeProdutos = 0
	
	dim strJsScriptCodZona
	strJsScriptCodZona = obtem_wms_deposito_zona_codigos
	strJsScriptCodZona = "<script language='JavaScript'>" & chr(13) & _
						 "var WMS_DEPOSITO_ZONA_CODIGOS = '" & strJsScriptCodZona & "';" & chr(13) & _
						 "</script>" & chr(13)




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' EXECUTA CONSULTA
'
sub executa_consulta
const w_codigo = 70
const w_descricao = 315
const w_qtd_todos = 60
const w_zona = 60
dim r
dim s, s_aux, s_sql, x, cab_table, cab, fabricante_a, msg_erro
dim n, n_total_linha, n_reg, qtde_fabricantes
dim n_sub_total_estoque_venda, n_sub_total_split_possivel, n_sub_total_a_separar, n_sub_total_todos
dim n_total_estoque_venda, n_total_split_possivel, n_total_a_separar, n_total_todos

'	SELECIONA TODOS OS PRODUTOS QUE POSSUEM ALGUM ITEM NA SITUAÇÃO DESEJADA
'	OBS: O USO DE 'UNION' SIMPLES ELIMINA AS LINHAS DUPLICADAS DOS RESULTADOS
'		 O USO DE 'UNION ALL' RETORNARIA TODAS AS LINHAS, INCLUSIVE AS DUPLICADAS
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.

'	PRODUTOS NO ESTOQUE DE VENDA
	s_sql = "SELECT DISTINCT" & _
				" fabricante," & _
				" produto" & _
			" FROM t_ESTOQUE_ITEM" & _
			" WHERE" & _
				" ((qtde-qtde_utilizada) > 0)"

	if c_fabricante <> "" then
		s_sql = s_sql & " AND (fabricante='" & c_fabricante & "')" 
		end if

'	PRODUTOS DO ESTOQUE DE PRODUTOS VENDIDOS (PEDIDOS EM STATUS SPLIT POSSÍVEL)
	s_sql = s_sql & _
			" UNION " & _
			"SELECT DISTINCT" & _
				" fabricante," & _
				" produto" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
			" WHERE" & _
				" (anulado_status=0)" & _
				" AND (estoque='" & ID_ESTOQUE_VENDIDO & "')" & _
				" AND (qtde > 0)" & _
				" AND (t_PEDIDO.st_entrega = '" & ST_ENTREGA_SPLIT_POSSIVEL & "')"

	if c_fabricante <> "" then
		s_sql = s_sql & " AND (fabricante='" & c_fabricante & "')" 
		end if

'	PRODUTOS DO ESTOQUE DE PRODUTOS VENDIDOS (PEDIDOS EM STATUS 'A SEPARAR')
	s_sql = s_sql & _
			" UNION " & _
			"SELECT DISTINCT" & _
				" fabricante," & _
				" produto" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
			" WHERE" & _
				" (anulado_status=0)" & _
				" AND (estoque='" & ID_ESTOQUE_VENDIDO & "')" & _
				" AND (qtde > 0)" & _
				" AND (t_PEDIDO.st_entrega = '" & ST_ENTREGA_SEPARAR & "')"

	if c_fabricante <> "" then
		s_sql = s_sql & " AND (fabricante='" & c_fabricante & "')" 
		end if

'	PRODUTOS DISPONÍVEIS PARA VENDA, INDEPENDENTEMENTE SE HÁ DISPONIBILIDADE NO ESTOQUE OU NÃO
	s_sql = s_sql & _
			" UNION " & _
			"SELECT DISTINCT" & _
				" t_PRODUTO.fabricante," & _
				" t_PRODUTO.produto" & _
			" FROM t_PRODUTO" & _
				" INNER JOIN t_PRODUTO_LOJA ON (t_PRODUTO.fabricante=t_PRODUTO_LOJA.fabricante) AND (t_PRODUTO.produto=t_PRODUTO_LOJA.produto)" & _
			" WHERE" & _
				" (vendavel = 'S')" & _
				" AND (descricao <> '.')" & _
				" AND (descricao <> '*')"

	if c_fabricante <> "" then
		s_sql = s_sql & " AND (t_PRODUTO.fabricante='" & c_fabricante & "')" 
		end if

'	A PARTIR DA CONSULTA QUE OBTÉM TODA A RELAÇÃO DE PRODUTOS A SER LISTADA,
'	REALIZA A CONSULTA P/ CALCULAR AS QUANTIDADES
	s_sql = "SELECT" & _
				" tAuxBase.fabricante," & _
				" tAuxBase.produto," & _
				" tProd.descricao," & _
				" tProd.descricao_html," & _
				" tMapZona.zona_codigo," & _
				"(" & _
					"SELECT" & _
						" Sum(qtde-qtde_utilizada)" & _
					" FROM t_ESTOQUE_ITEM tEI" & _
					" WHERE" & _
						" ((qtde-qtde_utilizada) > 0)" & _
						" AND (tEI.fabricante=tAuxBase.fabricante)" & _
						" AND (tEI.produto=tAuxBase.produto)" & _
				") AS qtde_estoque_venda," & _
				"(" & _
					"SELECT" & _
						" Sum(qtde)" & _
					" FROM t_ESTOQUE_MOVIMENTO tEM" & _
						" INNER JOIN t_PEDIDO tP ON (tEM.pedido=tP.pedido)" & _
					" WHERE" & _
						" (anulado_status=0)" & _
						" AND (estoque='" & ID_ESTOQUE_VENDIDO & "')" & _
						" AND (qtde > 0)" & _
						" AND (tP.st_entrega = '" & ST_ENTREGA_SPLIT_POSSIVEL & "')" & _
						" AND (tEM.fabricante=tAuxBase.fabricante)" & _
						" AND (tEM.produto=tAuxBase.produto)" & _
				") AS qtde_split_possivel," & _
				"(" & _
					"SELECT" & _
						" Sum(qtde)" & _
					" FROM t_ESTOQUE_MOVIMENTO tEM" & _
						" INNER JOIN t_PEDIDO tP ON (tEM.pedido=tP.pedido)" & _
					" WHERE" & _
						" (anulado_status=0)" & _
						" AND (estoque='" & ID_ESTOQUE_VENDIDO & "')" & _
						" AND (qtde > 0)" & _
						" AND (tP.st_entrega = '" & ST_ENTREGA_SEPARAR & "')" & _
						" AND (tEM.fabricante=tAuxBase.fabricante)" & _
						" AND (tEM.produto=tAuxBase.produto)" & _
				") AS qtde_a_separar" & _
			" FROM (" & s_sql & ") tAuxBase" & _
				" LEFT JOIN t_PRODUTO tProd ON (tAuxBase.fabricante=tProd.fabricante) AND (tAuxBase.produto=tProd.produto)" & _
				" LEFT JOIN t_WMS_DEPOSITO_MAP_ZONA tMapZona ON (tProd.deposito_zona_id=tMapZona.id)" & _
			" ORDER BY" & _
				" tAuxBase.fabricante," & _
				" tAuxBase.produto"


  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD valign='bottom' style='background:white;' NOWRAP>&nbsp;</TD>" & chr(13) & _
		  "		<TD valign='bottom' NOWRAP class='MT' style='width:" & Cstr(w_codigo) & "px'><P class='R'>Código</P></TD>" & chr(13) & _
		  "		<TD valign='bottom' NOWRAP class='MTBD' style='width:" & Cstr(w_descricao) & "px'><P class='R'>Produto</P></TD>" & chr(13) & _
		  "		<TD valign='bottom' NOWRAP class='MTBD' style='width:" & Cstr(w_qtd_todos) & "px'><P class='Rd' style='font-weight:bold;'>Qtde Total</P></TD>" & chr(13) & _
		  "		<TD valign='bottom' align='center' NOWRAP class='MTBD' style='width:" & Cstr(w_zona) & "px'><P class='Rc' style='font-weight:bold;'>Zona</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	qtde_fabricantes = 0
	fabricante_a = "XXXXX"

	n_sub_total_estoque_venda = 0
	n_sub_total_split_possivel = 0
	n_sub_total_a_separar = 0
	n_sub_total_todos = 0
	n_total_estoque_venda = 0
	n_total_split_possivel = 0
	n_total_a_separar = 0
	n_total_todos = 0

	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
		'	SUB-TOTAL POR FORNECEDOR
			if n_reg > 0 then
				x = x & "	<TR style='background: #FFFFDD' NOWRAP>" & chr(13) & _
						"		<TD style='background:white;'>&nbsp;</td>" & chr(13) & _
						"		<TD colspan='2' class='MEB'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						"		<TD class='MB'><P class='Cd'>" & formata_inteiro(n_sub_total_todos) & "</P></TD>" & chr(13) & _
						"		<TD class='MDB'>&nbsp;</TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x = "<BR>" & chr(13) & _
					"<BR>" & chr(13)
				end if

			x = x & cab_table
			qtde_fabricantes = qtde_fabricantes + 1
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<TD class='MDTE' align='center' colspan='4' style='background:azure;'><P class='C'>&nbsp;" & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13) & _
					cab
			n_sub_total_estoque_venda = 0
			n_sub_total_split_possivel = 0
			n_sub_total_a_separar = 0
			n_sub_total_todos = 0
			end if
		
	  ' CONTAGEM
		n_reg = n_reg + 1
		intQtdeProdutos = intQtdeProdutos + 1

		x = x & "	<TR NOWRAP id='TR_" & Cstr(n_reg) & "'>" & chr(13)

	 '> Nº LINHA
		x = x & "		<TD align='right' NOWRAP><P class='Rd' style='margin-right:2px;'>" & Cstr(n_reg) & ".</P></TD>" & chr(13)

	 '> CÓDIGO DO PRODUTO
		x = x & "		<TD class='MDBE' valign='bottom' style='width:" & Cstr(w_codigo) & "px;' NOWRAP><P class='C'>&nbsp;" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO DO PRODUTO
		x = x & "		<TD class='MDB' valign='bottom' style='width:" & Cstr(w_descricao) & "px'><P class='C' NOWRAP>&nbsp;" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

		n_total_linha = 0
		
	 '> ESTOQUE DE VENDA
		n = 0
		if Not IsNull(r("qtde_estoque_venda")) then n = r("qtde_estoque_venda")
		n_sub_total_estoque_venda = n_sub_total_estoque_venda + n
		n_total_estoque_venda = n_total_estoque_venda + n
		n_total_linha = n_total_linha + n

	 '> SPLIT POSSÍVEL
		n = 0
		if Not IsNull(r("qtde_split_possivel")) then n = r("qtde_split_possivel")
		n_sub_total_split_possivel = n_sub_total_split_possivel + n
		n_total_split_possivel = n_total_split_possivel + n
		n_total_linha = n_total_linha + n

	 '> A SEPARAR
		n = 0
		if Not IsNull(r("qtde_a_separar")) then n = r("qtde_a_separar")
		n_sub_total_a_separar = n_sub_total_a_separar + n
		n_total_a_separar = n_total_a_separar + n
		n_total_linha = n_total_linha + n

		n_sub_total_todos = n_sub_total_todos + n_total_linha
		n_total_todos = n_total_todos + n_total_linha

	 '> TOTAL
		x = x & "		<TD class='MDB' valign='bottom' style='width:" & Cstr(w_qtd_todos) & "px;' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(n_total_linha) & "</P></TD>" & chr(13)
		
	 '> ZONA
		x = x & "		<TD class='MDB' valign='bottom' style='width:" & Cstr(w_zona) & "px;' NOWRAP>" & chr(13) & _
							"<input type='hidden' id='c_wms_fabricante' name='c_wms_fabricante' value='" & Trim("" & r("fabricante")) & "' />" & chr(13) & _
							"<input type='hidden' id='c_wms_produto' name='c_wms_produto' value='" & Trim("" & r("produto")) & "' />" & chr(13) & _
							"<input type='hidden' id='c_wms_zona_original' name='c_wms_zona_original' value='" & Trim("" & r("zona_codigo")) & "' />" & chr(13) & _
							"<input type='text' id='c_wms_zona' name='c_wms_zona' class='PLLe' maxlength='1' style='text-align:center;width:60px;'" & _
								" value='" & Trim("" & r("zona_codigo")) & "'" & _
								" onfocus='this.select();realca_cor_linha(this," & Cstr(n_reg) & ");'" & _
								" onkeypress='if (digitou_enter(true)){if (" & Cstr(n_reg+1) & " < fREL.c_wms_zona.length) fREL.c_wms_zona[" & Cstr(n_reg+1) & "].focus(); else bCONFIRMA.focus();} filtra_digitacao_wms_deposito_zona_codigos(WMS_DEPOSITO_ZONA_CODIGOS, this.value);'" & _
								" onblur='this.value=ucase(trim(this.value));normaliza_cor_linha(this," & Cstr(n_reg) & "); if ((this.value!=fREL.c_wms_zona_original[" & Cstr(n_reg) & "].value)&&(!wms_deposito_codigo_ok(WMS_DEPOSITO_ZONA_CODIGOS,this.value))){alert(" & chr(34) & "Código inválido!" & chr(34) & "); this.focus();}'" & _
								" />" & chr(13) & _
						"</TD>" & chr(13)
		
		x = x & "	</TR>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA TOTAL DO ÚLTIMO FABRICANTE
	if n_reg <> 0 then 
		x = x & "	<TR style='background: #FFFFDD' NOWRAP>"  & chr(13) & _
				"		<TD style='background:white;'>&nbsp;</td>" & chr(13) & _
				"		<TD class='MEB' COLSPAN='2' NOWRAP><P class='Cd'>" & "Total:" & "</P></TD>" & chr(13) & _
				"		<TD class='MB' NOWRAP><P class='Cd'>" & formata_inteiro(n_sub_total_todos) & "</P></TD>" & chr(13) & _
				"		<TD class='MDB' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13)
	'>	TOTAL GERAL
		if qtde_fabricantes > 1 then
			x = x & _
				"<TR><TD COLSPAN='5'>&nbsp;</TD></TR>" & chr(13) & _
				"<TR><TD COLSPAN='5'>&nbsp;</TD></TR>" & chr(13) & _
				"	<TR style='background:honeydew'>" & chr(13) & _
				"		<TD style='background:white;'>&nbsp;</td>" & chr(13) & _
				"		<TD class='MTBE' style='width:" & Cstr(w_codigo) & "px;' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"		<TD class='MTB' style='width:" & Cstr(w_descricao) & "px' valign='bottom' NOWRAP><p class='Cd'>TOTAL GERAL:</p></TD>" & chr(13) & _
				"		<TD class='MTB' style='width:" & Cstr(w_qtd_todos) & "px;' valign='bottom' NOWRAP><p class='Cd'>" & formata_inteiro(n_total_todos) & "</p></TD>" & chr(13) & _
				"		<TD class='MTBD' style='width:" & Cstr(w_zona) & "px;' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD style='background:white;'>&nbsp;</td>" & chr(13) & _
			"		<TD colspan='4' class='MDBE'><P class='ALERTA'>&nbsp;Nenhum produto do estoque satisfaz as condições especificadas&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DA ÚLTIMA LOJA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

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



<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<%=strJsScriptCodZona%>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function realca_cor_linha(c, indice_row) {
var row;
	row = document.getElementById("TR_" + indice_row);
	row.style.backgroundColor = 'palegreen';
	c.style.backgroundColor = 'palegreen';
}

function normaliza_cor_linha(c, indice_row) {
var row;
	row = document.getElementById("TR_" + indice_row);
	row.style.backgroundColor = '';
	c.style.backgroundColor = '';
}

function fRELGravaDados() {
	window.status = "Aguarde ...";
	fREL.action = "RelProdutoDepositoZonaGravaDados.asp";
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
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">

<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" id="c_wms_fabricante" name="c_wms_fabricante" value="" />
<input type="hidden" id="c_wms_produto" name="c_wms_produto" value="" />
<input type="hidden" id="c_wms_zona" name="c_wms_zona" value="" />
<input type="hidden" id="c_wms_zona_original" name="c_wms_zona_original" value="" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Zona do Produto (Depósito)</span>
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
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='4' style='border-bottom:1px solid black' border='0'>" & chr(13)
	s = Trim(c_fabricante)
	if s = "" then
		s = "TODOS"
	else
		if s_nome_fabricante <> "" then s = s & " - " & s_nome_fabricante 
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Fabricante:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
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
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<% if intQtdeProdutos > 0 then %>
	<td align="left"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0" /></a></td>
	<% else %>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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
