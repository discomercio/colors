<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L P R O D B L O Q U E A D O . A S P
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

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	alerta = ""

	dim s, c_fabricante, c_produto
	c_fabricante = retorna_so_digitos(Request.Form("c_fabricante"))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	
	if c_fabricante <> "" then
		s = normaliza_codigo(c_fabricante, TAM_MIN_FABRICANTE)
		if s <> "" then c_fabricante = s
		end if

	if c_produto <> "" then
		s = normaliza_produto(c_produto)
		if s <> "" then c_produto = s
		end if

	if c_fabricante = "" then
		alerta = "Especifique o código do fabricante."
	elseif c_produto = "" then
		alerta = "Especifique o código do produto."
		end if
				
	if alerta = "" then
		s = "SELECT fabricante, nome, razao_social FROM t_FABRICANTE WHERE (fabricante='" & c_fabricante & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			alerta = "Fabricante " & c_fabricante & " não está cadastrado."
			end if
		end if
	
	if alerta = "" then
	'	O FLAG "EXCLUIDO_STATUS" INDICA SE O PRODUTO ESTÁ EXCLUÍDO LOGICAMENTE DO SISTEMA!!
	'	A TABELA BÁSICA DE PRODUTOS MANTÉM INFORMAÇÕES DE PRODUTOS EXCLUÍDOS LOGICAMENTE 
	'	P/ MANTER A REFERÊNCIA COM OUTRAS TABELAS QUE NECESSITEM DE DADOS COMO DESCRIÇÃO, ETC.
		s = "SELECT fabricante, produto FROM t_PRODUTO WHERE" & _
			" (excluido_status = 0)" & _
			" AND (fabricante = '" & c_fabricante & "')" & _
			" AND (produto = '" & c_produto & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			alerta = "O produto " & c_produto & " do fabricante " & c_fabricante & " não está cadastrado."
			end if
		end if
		
	if alerta = "" then
		s = "SELECT loja, fabricante, produto FROM t_PRODUTO_LOJA WHERE" & _
			" (vendavel = 'S')" & _
			" AND (loja = '" & loja & "')" & _
			" AND (fabricante = '" & c_fabricante & "')" & _
			" AND (produto = '" & c_produto & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			alerta = "O produto " & c_produto & " do fabricante " & c_fabricante & " não está cadastrado."
			end if
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim s, s_aux, s_sql, x, cab_table, cab, msg_erro
dim n_reg

'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT t_PEDIDO.data, t_PEDIDO.pedido, t_PEDIDO_ITEM.fabricante, t_PEDIDO_ITEM.produto,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde) AS qtde" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			" INNER JOIN t_ESTOQUE_MOVIMENTO ON (t_PEDIDO_ITEM.pedido=t_ESTOQUE_MOVIMENTO.pedido) AND (t_PEDIDO_ITEM.fabricante=t_ESTOQUE_MOVIMENTO.fabricante) AND (t_PEDIDO_ITEM.produto=t_ESTOQUE_MOVIMENTO.produto)" & _
			" LEFT JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
			" WHERE (st_entrega='" & ST_ENTREGA_SPLIT_POSSIVEL & "')" & _
			" AND (t_PEDIDO.loja='" & loja & "')" & _
			" AND (t_ESTOQUE_MOVIMENTO.anulado_status=0)" & _
			" AND (t_ESTOQUE_MOVIMENTO.estoque='" & ID_ESTOQUE_VENDIDO & "')" & _
			" AND (t_ESTOQUE_MOVIMENTO.fabricante='" & c_fabricante & "')" & _
			" AND (t_ESTOQUE_MOVIMENTO.produto='" & c_produto & "')"
			
	s_sql = s_sql & " GROUP BY t_PEDIDO.data, t_PEDIDO.pedido, t_PEDIDO_ITEM.fabricante, t_PEDIDO_ITEM.produto,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome_iniciais_em_maiusculas"
		end if

	s_sql = s_sql & _
			" ORDER BY t_PEDIDO.data, t_PEDIDO.pedido"
	
  ' CABEÇALHO
	cab_table = "<TABLE class='Q' style='border-bottom:0px;' cellSpacing=0>" & chr(13)
	cab = "<TR style='background: #FFF0E0' nowrap>" & _
		  "<TD width='80' valign='bottom' nowrap class='MD MB'><P class='R'>Nº PEDIDO</P></TD>" & _
		  "<TD width='400' valign='bottom' nowrap class='MD MB'><P class='R'>CLIENTE</P></TD>" & _
		  "<TD width='45' valign='bottom' nowrap class='MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & _
		  "</TR>" & chr(13)
	
	s = c_fabricante
	s_aux = iniciais_em_maiusculas(x_fabricante(s))
	if (s<>"") And (s_aux<>"") then s = s & " - "
	s = s & s_aux
	x = cab_table & _
		"<TR><TD class='MB' align='right' style='background:azure;'>" & _
		"<P class='F'>Fabricante:</P></TD>" & _
		"<TD COLSPAN='2' class='MB' style='background:azure;'>" & _
		"<P class='F'>" & s & "</P></TD></TR>"

	s = c_produto
	s_aux = produto_formata_descricao_em_html(produto_descricao_html(c_fabricante, c_produto))
	if (s<>"") And (s_aux<>"") then s = s & " - "
	s = s & s_aux
	x = x & "<TR><TD class='MB' align='right' valign='baseline' style='background:azure;'>" & _
			"<P class='F'>Produto:</P></TD>" & _
			"<TD COLSPAN='2' class='MB' valign='baseline' style='background:azure;'>" & _
			"<P class='F'>" & s & "</P></TD></TR>"
	
	x = x & chr(13) & cab

	n_reg = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	 ' CONTAGEM
		n_reg = n_reg + 1

		x = x & "<TR NOWRAP>"

	'> Nº PEDIDO
		x = x & "	<TD class='MDB'><P class='C'>&nbsp;<a href='javascript:fLPRECOSConcluir(" & _
			chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & _
			Trim("" & r("pedido")) & "</a></P></TD>"

	'> CLIENTE
		x = x & "	<TD class='MDB'><P class='C'>&nbsp;" & Trim("" & r("nome_iniciais_em_maiusculas")) & "</P></TD>"

 	'> QTDE
		x = x & "	<TD class='MB' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(r("qtde")) & "</P></TD>"

		x = x & "</TR>"

		r.movenext
		loop
		
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = x & _
			"<TR NOWRAP>" & _
			"	<TD colspan='3' class='MB'><P class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</P></TD>" & _
			"</TR>"
		end if

  ' FECHA TABELA
	x = x & "</TABLE>"
	
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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fLPRECOSConcluir( id_pedido ) {
	fLPRECOS.action = "pedido.asp";
	fLPRECOS.pedido_selecionado.value = id_pedido;
	fLPRECOS.submit();
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
<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="fLPRECOS" name="fLPRECOS" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Produtos "Bloqueados" no Estoque de Entrega</span>
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
<br>


<!--  LISTAGEM  -->
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
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
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
