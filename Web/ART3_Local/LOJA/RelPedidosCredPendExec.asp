<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L P E D I D O S C R E D P E N D E X E C . A S P
'     =================================================================
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
	
	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_REL_PEDIDOS_CREDITO_PENDENTE, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	alerta = ""
	
	dim vendedor_selecionado
	vendedor_selecionado = Trim(Request.Form("vendedor_selecionado"))
	if vendedor_selecionado = "" then
		if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then 
			alerta = "VENDEDOR NÃO FOI ESPECIFICADO."
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
dim x, s, s_aux, s_sql, cab_table, cab, n_reg, n_reg_total
dim s_where, s_from
dim vendedor_a, qtde_vendedores
dim vl_sub_total, vl_total_geral
dim w_cliente, w_valor

'	MONTA CLÁUSULA WHERE
'	DESCONSIDERA OS PEDIDOS-FILHOTE, POIS A ANÁLISE DE CRÉDITO É ANOTADA NO PEDIDO-BASE
	s_where = " (t_PEDIDO.tamanho_num_pedido = " & Cstr(TAM_MIN_ID_PEDIDO) & ")"

'	CRITÉRIO: ANÁLISE DE CRÉDITO PENDENTE
	s = " analise_credito = " & COD_AN_CREDITO_PENDENTE
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	CRITÉRIO: EXCLUI PEDIDOS CANCELADOS
	s = "(st_entrega <> '" & ST_ENTREGA_CANCELADO & "')"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	CRITÉRIO: LOJA (CADA LOJA SÓ PODE CONSULTAR SEUS PRÓPRIOS PEDIDOS)
	s = "(t_PEDIDO.numero_loja = " & loja & ")"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	CRITÉRIO: VENDEDOR
	if vendedor_selecionado <> "" then
		s = "(vendedor = '" & vendedor_selecionado & "')"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CLÁUSULA WHERE
	if s_where <> "" then s_where = " WHERE" & s_where
	
'	MONTA CLÁUSULA FROM
	s_from = " FROM t_PEDIDO INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)"

	s_sql = "SELECT t_PEDIDO.loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO.data, t_PEDIDO.pedido, t_PEDIDO.vl_total_familia,"
			
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.vendedor" & _
			s_from & _
			s_where

	s_sql = s_sql & " ORDER BY t_PEDIDO.vendedor, t_PEDIDO.data, t_PEDIDO.pedido"

  ' CABEÇALHO
	w_cliente = 250
	w_valor = 80
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure'>" & chr(13) & _
		  "		<TD class='MDTE' style='width:70px' valign='bottom' NOWRAP><P class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:70px' valign='bottom'><P class='R'>Data</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:" & Cstr(w_cliente) & "px' valign='bottom'><P class='R'>Cliente</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>VL Total</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_vendedores = 0
	vl_sub_total = 0
	vl_total_geral = 0

	vendedor_a = "XXXXXXXXXXXXXXX"

	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE VENDEDOR?
		if Trim("" & r("vendedor"))<>vendedor_a then
			vendedor_a = Trim("" & r("vendedor"))
			qtde_vendedores = qtde_vendedores + 1
		  ' FECHA TABELA DA LOJA ANTERIOR
			if n_reg > 0 then 
				x = x & "	<TR style='background: #FFFFDD'>" & chr(13) & _
						"		<TD class='MTBE' colspan='3' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
						"		<TD class='MTBD'><p class='Cd'>" & formata_moeda(vl_sub_total) & "</p></td>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			n_reg = 0
			vl_sub_total = 0

			if n_reg_total > 0 then x = x & "<BR>" & chr(13)
			s = UCase(Trim("" & r("vendedor")))
			s_aux = x_usuario(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			if s = "" then s = "&nbsp;"
			x = x & cab_table
			if s <> "" then 
				x = x & _
					"	<TR>" & chr(13) & _
					"		<TD class='MDTE' colspan='4' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
				end if
			x = x & cab
			end if
	
	 ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR>" & chr(13)

	'> Nº PEDIDO
		x = x & "		<TD valign='top' class='MDTE'><P class='C'>&nbsp;<a href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & _
				")' title='clique para consultar o pedido'>" & Trim("" & r("pedido")) & "</a></P></TD>" & chr(13)

	'> DATA
		s = formata_data(r("data"))
		x = x & "		<TD valign='top' class='MTD'><P class='Cc'>" & s & "</P></TD>" & chr(13)

	'> CLIENTE
		s = Trim("" & r("nome_iniciais_em_maiusculas"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='top' style='width:" & Cstr(w_cliente) & "px' class='MTD'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> VALOR DO PEDIDO
		s = formata_moeda(r("vl_total_familia"))
		x = x & "		<TD valign='top' align='right' style='width:" & Cstr(w_valor) & "px' class='MTD'><P class='Cnd'>" & s & "</P></TD>" & chr(13)

	'> TOTALIZAÇÃO DE VALORES
		vl_sub_total = vl_sub_total + r("vl_total_familia")
		vl_total_geral = vl_total_geral + r("vl_total_familia")
			
		x = x & "	</TR>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop

	
  ' MOSTRA TOTAL DO ÚLTIMO VENDEDOR
	if n_reg <> 0 then 
		x = x & "	<TR style='background: #FFFFDD'>" & chr(13) & _
				"		<TD colspan='3' class='MTBE' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
				"		<TD class='MTBD'><p class='Cd'>" & formata_moeda(vl_sub_total) & "</p></td>" & chr(13) & _
				"	</TR>" & chr(13)
		
	'>	TOTAL GERAL
		if qtde_vendedores > 1 then
			x = x & _
				"	<TR>" & chr(13) & _
				"		<TD colspan='4' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR>" & chr(13) & _
				"		<TD colspan='4' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:honeydew'>" & chr(13) & _
				"		<TD class='MTBE' colspan='3' NOWRAP><p class='Cd'>TOTAL GERAL:</p></td>" & chr(13) & _
				"		<TD class='MTBD'><p class='Cd'>" & formata_moeda(vl_total_geral) & "</p></td>" & chr(13) & _
				"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR>" & chr(13) & _
				"		<TD class='MT' colspan='4'><P class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir( id_pedido ){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=id_pedido;
	fREL.action = "Pedido.asp"
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
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos com Crédito Pendente</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

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
