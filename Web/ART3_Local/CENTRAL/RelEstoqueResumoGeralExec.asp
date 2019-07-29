<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelEstoqueResumoGeralExec.asp
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if (Not operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas)) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim s, s_filtro

	dim alerta
	alerta = ""



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r, s_sql_template, s_sql, s_estoque_placeholder
dim cab_table, cab, x
dim n_saldo_total, vl_total_geral
dim n_qtde, vl_valor

	If Not cria_recordset_otimista(r, msg_erro) then
		 Response.Write "Falha ao tentar criar recordset para acesso ao banco de dados"
		 Response.End
		 end if

	s_estoque_placeholder = "_XX_XX_"

	n_saldo_total = 0
	vl_total_geral = 0

  ' CABEÇALHO
	cab_table = "<table class='MC' cellspacing='0'>" & chr(13)
	cab = "	<tr style='background:azure' nowrap>" & chr(13) & _
			"		<td class='MD ME MB TdTitEstoque'><span class='R'>ESTOQUE</span></td>" & chr(13) & _
			"		<td class='MD MB TdTitQtde'><span class='Rd' style='font-weight:bold;'>QTDE</span></td>" & chr(13) & _
			"		<td class='MDB TdTitValor'><span class='Rd' style='font-weight:bold;'>VL TOTAL</span></td>" & chr(13) & _
			"	</tr>" & chr(13)
	x = cab_table & cab

	'CONSULTA ESTOQUE VENDA + SHOW ROOM
	s_sql = " SELECT" & _
				" Coalesce(SUM(preco_total),0) AS preco_total," & _
				" Coalesce(SUM(saldo),0) AS saldo" & _
			" FROM (" & _
				"SELECT" & _
					" t_ESTOQUE_ITEM.fabricante," & _
					" t_ESTOQUE_ITEM.produto," & _
					" SUM((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total," & _
					" SUM(t_ESTOQUE_ITEM.qtde-t_ESTOQUE_ITEM.qtde_utilizada) AS saldo" & _
				" FROM t_ESTOQUE_ITEM" & _
					" INNER JOIN t_ESTOQUE ON (t_ESTOQUE.id_estoque = t_ESTOQUE_ITEM.id_estoque)" & _
					" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
				" WHERE" & _
					" ((t_ESTOQUE_ITEM.qtde-t_ESTOQUE_ITEM.qtde_utilizada) > 0)" & _
				" GROUP BY" & _
					" t_ESTOQUE_ITEM.fabricante," & _
					" t_ESTOQUE_ITEM.produto" & _
				" UNION ALL " & _
				"SELECT" & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
					" t_ESTOQUE_MOVIMENTO.produto," & _
					" SUM(t_ESTOQUE_MOVIMENTO.qtde*t_ESTOQUE_ITEM.vl_custo2) AS preco_total," & _
					" SUM(t_ESTOQUE_MOVIMENTO.qtde) AS saldo" & _
				" FROM t_ESTOQUE_MOVIMENTO" & _
					" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
					" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
				" WHERE" & _
					" (anulado_status=0)" & _
					" AND (estoque='SHR')" & _
				" GROUP BY" & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
					" t_ESTOQUE_MOVIMENTO.produto" & _
				") tbl"

	n_qtde = 0
	vl_valor = 0
	if r.State <> 0 then r.Close
	r.Open s_sql, cn
	if Not r.Eof then
		n_qtde = r("saldo")
		vl_valor = r("preco_total")
		n_saldo_total = n_saldo_total + n_qtde
		vl_total_geral = vl_total_geral + vl_valor
		end if

	x = x & "	<tr>" & chr(13) & _
			"		<td class='MD ME MB TdCelEstoque'>" & _
					"<span class='C'>VENDA + SHOW ROOM</span>" & _
					"</td>" & chr(13) & _
			"		<td class='MD MB TdCelQtde'>" & _
					"<span class='Cd'>" & formata_inteiro(n_qtde) & "</span>" & _
					"</td>" & chr(13) & _
			"		<td class='MD MB TdCelValor'>" & _
					"<span class='Cd'>" & formata_moeda(vl_valor) & "</span>" & _
					"</td>" & chr(13) & _
			"	</tr>" & chr(13)

	'MONTA SQL GENÉRICO PARA CONSULTAS DE ESTOQUE ATRAVÉS DA TABELA t_ESTOQUE_MOVIMENTO
	s_sql_template = _
			"SELECT" & _
				" Coalesce(SUM(preco_total),0) AS preco_total," & _
				" Coalesce(SUM(saldo),0) AS saldo" & _
			" FROM (" & _
				"SELECT" & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
					" t_ESTOQUE_MOVIMENTO.produto," & _
					" SUM(t_ESTOQUE_MOVIMENTO.qtde*t_ESTOQUE_ITEM.vl_custo2) AS preco_total," & _
					" SUM(t_ESTOQUE_MOVIMENTO.qtde) AS saldo" & _
				" FROM t_ESTOQUE_MOVIMENTO" & _
					" INNER JOIN t_ESTOQUE ON (t_ESTOQUE.id_estoque = t_ESTOQUE_MOVIMENTO.id_estoque)" & _
					" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
					" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
				" WHERE" & _
					" (anulado_status=0)" & _
					" AND (estoque='" & s_estoque_placeholder & "')" & _
				" GROUP BY" & _
					" t_ESTOQUE_MOVIMENTO.fabricante," & _
					" t_ESTOQUE_MOVIMENTO.produto" & _
				") t"

	'ESTOQUE: VENDIDO
	s_sql = Replace(s_sql_template, s_estoque_placeholder, ID_ESTOQUE_VENDIDO)

	n_qtde = 0
	vl_valor = 0
	if r.State <> 0 then r.Close
	r.Open s_sql, cn
	if Not r.Eof then
		n_qtde = r("saldo")
		vl_valor = r("preco_total")
		n_saldo_total = n_saldo_total + n_qtde
		vl_total_geral = vl_total_geral + vl_valor
		end if

	x = x & "	<tr>" & chr(13) & _
			"		<td class='MD ME MB TdCelEstoque'>" & _
					"<span class='C'>VENDIDO</span>" & _
					"</td>" & chr(13) & _
			"		<td class='MD MB TdCelQtde'>" & _
					"<span class='Cd'>" & formata_inteiro(n_qtde) & "</span>" & _
					"</td>" & chr(13) & _
			"		<td class='MD MB TdCelValor'>" & _
					"<span class='Cd'>" & formata_moeda(vl_valor) & "</span>" & _
					"</td>" & chr(13) & _
			"	</tr>" & chr(13)

	'ESTOQUE: DANIFICADO
	s_sql = Replace(s_sql_template, s_estoque_placeholder, ID_ESTOQUE_DANIFICADOS)

	n_qtde = 0
	vl_valor = 0
	if r.State <> 0 then r.Close
	r.Open s_sql, cn
	if Not r.Eof then
		n_qtde = r("saldo")
		vl_valor = r("preco_total")
		n_saldo_total = n_saldo_total + n_qtde
		vl_total_geral = vl_total_geral + vl_valor
		end if

	x = x & "	<tr>" & chr(13) & _
			"		<td class='MD ME MB TdCelEstoque'>" & _
					"<span class='C'>DANIFICADOS</span>" & _
					"</td>" & chr(13) & _
			"		<td class='MD MB TdCelQtde'>" & _
					"<span class='Cd'>" & formata_inteiro(n_qtde) & "</span>" & _
					"</td>" & chr(13) & _
			"		<td class='MD MB TdCelValor'>" & _
					"<span class='Cd'>" & formata_moeda(vl_valor) & "</span>" & _
					"</td>" & chr(13) & _
			"	</tr>" & chr(13)

	'ESTOQUE: DEVOLUÇÃO
	s_sql = Replace(s_sql_template, s_estoque_placeholder, ID_ESTOQUE_DEVOLUCAO)

	n_qtde = 0
	vl_valor = 0
	if r.State <> 0 then r.Close
	r.Open s_sql, cn
	if Not r.Eof then
		n_qtde = r("saldo")
		vl_valor = r("preco_total")
		n_saldo_total = n_saldo_total + n_qtde
		vl_total_geral = vl_total_geral + vl_valor
		end if

	x = x & "	<tr>" & chr(13) & _
			"		<td class='MD ME MB TdCelEstoque'>" & _
					"<span class='C'>DEVOLUÇÃO</span>" & _
					"</td>" & chr(13) & _
			"		<td class='MD MB TdCelQtde'>" & _
					"<span class='Cd'>" & formata_inteiro(n_qtde) & "</span>" & _
					"</td>" & chr(13) & _
			"		<td class='MD MB TdCelValor'>" & _
					"<span class='Cd'>" & formata_moeda(vl_valor) & "</span>" & _
					"</td>" & chr(13) & _
			"	</tr>" & chr(13)

	'TOTAL GERAL
	x = x & "	<tr style='background:palegreen;'>" & chr(13) & _
			"		<td class='MD ME MB TdCelEstoque'>" & _
					"<span class='C'>TOTAL GERAL</span>" & _
					"</td>" & chr(13) & _
			"		<td class='MD MB TdCelQtde'>" & _
					"<span class='Cd'>" & formata_inteiro(n_saldo_total) & "</span>" & _
					"</td>" & chr(13) & _
			"		<td class='MD MB TdCelValor'>" & _
					"<span class='Cd'>" & formata_moeda(vl_total_geral) & "</span>" & _
					"</td>" & chr(13) & _
			"	</tr>" & chr(13)

	x = x & "</table>" & chr(13)

	Response.Write x

	if r.State <> 0 then r.Close
	set r = nothing

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



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
    window.status = 'Aguarde, executando a consulta ...';
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

<style type="text/css">
.TdLblFiltro{
	padding-left:15px;
	text-align:right;
	vertical-align:middle;
}
.TdTitEstoque
{
	text-align:left;
	vertical-align:bottom;
	width:180px;
}
.TdCelEstoque{
	text-align:left;
}
.TdTitQtde{
	text-align: right;
	vertical-align:bottom;
	width:100px;
}
.TdCelQtde{
	text-align: right;
}
.TdTitValor{
	text-align: right;
	vertical-align:bottom;
	width:150px;
}
.TdCelValor{
	text-align: right;
}
span.C, span.R {
	font-size: 10pt;
}

span.Cc, span.Rc {
	font-size: 10pt;
}

span.Cd, span.Rd {
	font-size: 10pt;
}

span.N {
	font-size: 11pt;
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
<br /><br />
<p class="TracoBottom"></p>
<table cellspacing="0">
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
<body onload="window.status='Concluído';" link="#000000" alink="#000000" vlink="#000000">

<center>

<form id="fESTOQ" name="fESTOQ" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque: Resumo Posição Geral</span>
		<br />
		<span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!--  CABEÇALHO  -->
<%
	s_filtro = "<table width='649' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black;padding-top:6px;padding-bottom:6px;' border='0'>" & chr(13) & _
				"	<tr>" & chr(13) & _
				"		<td class='TdLblFiltro' nowrap>" & _
							"<span class='N'>Emitido em:&nbsp;</span>" & _
						"</td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'>" & _
							"<span class='N'>" & formata_data_hora_sem_seg(Now) & "</span>" & _
						"</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr>" & chr(13) & _
				"		<td class='TdLblFiltro' nowrap>" & _
							"<span class='N'>Emitido por:&nbsp;</span>" & _
						"</td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'>" & _
							"<span class='N'>" & usuario & "</span>" & _
						"</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"</table>" & chr(13)
	Response.Write s_filtro
%>

<br />

<%
	consulta_executa
%>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br />


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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
