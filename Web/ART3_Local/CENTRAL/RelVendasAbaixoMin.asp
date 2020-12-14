<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L V E N D A S A B A I X O M I N . A S P
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
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_VENDAS_COM_DESCONTO_SUPERIOR, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	alerta = ""

	dim s, s_aux, s_filtro, c_dt_inicio, c_dt_termino
	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))

'	Período de consulta está restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_termino
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			if c_dt_inicio = "" then c_dt_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
			end if
		
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
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
dim s_sql, strAbaixoMinSupervAutorizador
dim x, cab_table, cab, cab_detalhe, pedido_a, vl
dim qtde_pedidos, n_reg, n_reg_total
	
	s_sql = "SELECT t_PEDIDO.data, t_PEDIDO.pedido, t_PEDIDO.vendedor," & _
			" abaixo_min_autorizador, abaixo_min_superv_autorizador, qtde, fabricante, produto," &_
			" descricao, descricao_html, preco_lista, desc_max, preco_venda,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome_iniciais_em_maiusculas"
		end if

	s_sql = s_sql & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			" LEFT JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
			" WHERE (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')" & _
			" AND (abaixo_min_status<>0)"

	if IsDate(c_dt_inicio) then
		s_sql = s_sql & " AND (data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		s_sql = s_sql & " AND (data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if

	s_sql = s_sql & " ORDER BY t_PEDIDO.data, t_PEDIDO.pedido, t_PEDIDO_ITEM.sequencia"

  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0 cellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & _
		  "		<TD style='width:70px' class='MT' valign='bottom' NOWRAP><P class='R'>Nº Pedido</P></TD>" & _
		  "		<TD style='width:303px' class='MTBD' valign='bottom' NOWRAP><P class='R'>Cliente</P></TD>" & _
		  "		<TD style='width:90px' class='MTBD' valign='bottom'><P class='R'>Vendedor</P></TD>" & _
		  "		<TD style='width:90px' class='MTBD' valign='bottom'><P class='R'>Desc Cadastrado Por</P></TD>" & _
		  "		<TD style='width:90px' class='MTBD' valign='bottom'><P class='R'>Desc Autorizado Por</P></TD>" & _
		  "	</TR>" & chr(13)

	cab_detalhe =	"	<TR style='background:#FFF0E0' NOWRAP>" & _
					"		<TD class='MDBE' style='width:40px;' align='right' valign='bottom' NOWRAP><P style='font-weight:bold;' class='Rd'>Qtde</P></TD>" & _
					"		<TD class='MDB' style='width:50px' valign='bottom' NOWRAP><P class='R'>Fabr</P></TD>" & _
					"		<TD class='MDB' style='width:70px' valign='bottom'><P class='R'>Produto</P></TD>" & _
					"		<TD class='MDB' valign='bottom'><P class='R'>Descrição</P></TD>" & _
					"		<TD class='MDB' style='width:90px;' align='right' valign='bottom'><P style='font-weight:bold;' class='Rd'>VL Unit Mín</P></TD>" & _
					"		<TD class='MDB' style='width:90px;' align='right' valign='bottom'><P style='font-weight:bold;' class='Rd'>VL Unit Venda</P></TD>" & _
					"	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_pedidos = 0
	
	pedido_a = "XXXXX"
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE PEDIDO?
		if Trim("" & r("pedido"))<>pedido_a then
			pedido_a = Trim("" & r("pedido"))
			qtde_pedidos = qtde_pedidos + 1
		  ' FECHA TABELA DO PEDIDO ANTERIOR
			if n_reg_total > 0 then
				x = x & "</TABLE></TD></TR>"
				x = x & "</TABLE>" & chr(13)
				Response.Write x
				end if
			
			x=""
			n_reg = 0

			strAbaixoMinSupervAutorizador = Trim("" & r("abaixo_min_superv_autorizador"))
			if strAbaixoMinSupervAutorizador = "" then strAbaixoMinSupervAutorizador = "&nbsp;"
			if n_reg_total > 0 then x = x & "<BR>"
			x = x & cab_table & cab
			x = x & "	<TR NOWRAP>" & _
				"		<TD class='MDBE'><P class='C'><a href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & _
				Trim("" & r("pedido")) & "</a></P></TD>" & _
				"		<TD class='MDB' style='width:303px'><P class='C'>" & Trim("" & r("nome_iniciais_em_maiusculas")) & "</P></TD>" & _
				"		<TD class='MDB'><P class='C'>" & Trim("" & r("vendedor")) & "</P></TD>" & _
				"		<TD class='MDB'><P class='C'>" & Trim("" & r("abaixo_min_autorizador")) & "</P></TD>" & _
				"		<TD class='MDB'><P class='C'>" & strAbaixoMinSupervAutorizador & "</P></TD>" & _
				"	</TR>" & chr(13) & _
				"	<TR><TD COLSPAN='5'>" & "<TABLE width='100%' cellSpacing=0 cellPadding=0>" & cab_detalhe
			end if
		

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>"

	 '> QTDE
		x = x & "		<TD align='right' valign='bottom' class='MDBE'><P class='Cd'>" & formata_inteiro(r("qtde")) & "</P></TD>"
	
	 '> FABRICANTE
		x = x & "		<TD class='MDB' valign='bottom'><P class='C'>" & Trim("" & r("fabricante")) & "</P></TD>"
		
	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='bottom'><P class='C'>" & Trim("" & r("produto")) & "</P></TD>"

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='bottom'><P class='C'>" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>"
		
	 '> VALOR UNITÁRIO MÍNIMO
		vl = r("preco_lista") - (r("preco_lista")*(r("desc_max")/100))
		x = x & "		<TD align='right' valign='bottom' class='MDB'><P class='Cd'>" & formata_moeda(vl) & "</P></TD>"

	 '> VALOR UNITÁRIO VENDA
		x = x & "		<TD align='right' valign='bottom' class='MDB'><P class='Cd'>" & formata_moeda(r("preco_venda")) & "</P></TD>"

		x = x & "</TR>" & chr(13)
			
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
		
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & _
				"		<TD class='MT' colspan='5'><P class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</P></TD>" & _
				"	</TR>"
	else
		x = x & "</TABLE></TD></TR>"
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
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir( id_pedido ) {
	fREL.action = "pedido.asp";
	fREL.pedido_selecionado.value = id_pedido;
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

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Vendas com Desconto Superior</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>"

	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux

	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Período:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & formata_data_hora(Now) & "</p></td></tr>"
	
	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
