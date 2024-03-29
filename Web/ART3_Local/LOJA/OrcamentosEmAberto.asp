<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  O R C A M E N T O S E M A B E R T O . A S P
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

'	VERIFICA PERMISS�O DE ACESSO DO USU�RIO
	if (Not operacao_permitida(OP_LJA_CONSULTA_ORCAMENTO, s_lista_operacoes_permitidas)) And _
	   (Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas)) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	alerta = ""





' _____________________________________________________________________________________________
'
'									F  U  N  �  �  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim x, s, s_aux, s_sql, cab_table, cab, n_reg, n_reg_total
dim s_where, s_from
dim orcamentista_a, qtde_orcamentistas
dim vl_sub_total, vl_total_geral
dim w_cliente, w_valor

'	MONTA CL�USULA WHERE
	s_where = ""

'	CRIT�RIO: STATUS DE FECHAMENTO DO OR�AMENTOS
	s = "(t_ORCAMENTO.st_fechamento='') OR (t_ORCAMENTO.st_fechamento IS NULL)"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	CRIT�RIO: STATUS DO OR�AMENTO
	s = "(t_ORCAMENTO.st_orcamento='') OR (t_ORCAMENTO.st_orcamento IS NULL)"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	CRIT�RIO: OR�AMENTOS QUE N�O VIRARAM PEDIDOS (NOVO CAMPO DE CONTROLE)
	s = "(t_ORCAMENTO.st_orc_virou_pedido = 0)"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	CRIT�RIO: LOJA (CADA LOJA S� PODE CONSULTAR SEUS PR�PRIOS OR�AMENTOS)
	s = "(CONVERT(smallint,t_ORCAMENTO.loja) = " & loja & ")"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	VERIFICA PERMISS�O DE ACESSO DO USU�RIO
	if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then 
	'	CRIT�RIO: VENDEDOR (PODE ACESSAR OR�AMENTOS DE TODOS OS VENDEDORES DA LOJA?)
		s = "(vendedor = '" & usuario & "')"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CL�USULA WHERE
	if s_where <> "" then s_where = " WHERE" & s_where
	
'	MONTA CL�USULA FROM
	s_from = " FROM t_ORCAMENTO INNER JOIN t_CLIENTE ON (t_ORCAMENTO.id_cliente=t_CLIENTE.id)"

	s_sql = "SELECT t_ORCAMENTO.loja, CONVERT(smallint,t_ORCAMENTO.loja) AS numero_loja," & _
			" t_ORCAMENTO.data, t_ORCAMENTO.orcamento, t_ORCAMENTO.vl_total,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_ORCAMENTO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_ORCAMENTO.orcamentista" & _
			s_from & _
			s_where

	s_sql = s_sql & " ORDER BY t_ORCAMENTO.orcamentista, t_ORCAMENTO.data, t_ORCAMENTO.nsu, t_ORCAMENTO.orcamento"

  ' CABE�ALHO
	w_cliente = 250
	w_valor = 80
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure'>" & chr(13) & _
		  "		<TD class='MDTE' style='width:70px' valign='bottom' NOWRAP><P class='R'> Pr�-Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:70px' valign='bottom'><P class='R'>Data</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:" & Cstr(w_cliente) & "px' valign='bottom'><P class='R'>Cliente</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>VL Total</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_orcamentistas = 0
	vl_sub_total = 0
	vl_total_geral = 0

	orcamentista_a = "XXXXXXXXXXXXXXX"

	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE OR�AMENTISTA?
		if Trim("" & r("orcamentista"))<>orcamentista_a then
			orcamentista_a = Trim("" & r("orcamentista"))
			qtde_orcamentistas = qtde_orcamentistas + 1
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
			s = UCase(Trim("" & r("orcamentista")))
			s_aux = x_orcamentista_e_indicador(s)
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

	'> N� OR�AMENTO
		x = x & "		<TD valign='top' class='MDTE'><P class='C'>&nbsp;<a href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("orcamento")) & chr(34) & _
				")' title='clique para consultar o pr�-pedido'>" & Trim("" & r("orcamento")) & "</a></P></TD>" & chr(13)

	'> DATA
		s = formata_data(r("data"))
		x = x & "		<TD valign='top' class='MTD'><P class='Cc'>" & s & "</P></TD>" & chr(13)

	'> CLIENTE
		s = Trim("" & r("nome_iniciais_em_maiusculas"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='top' style='width:" & Cstr(w_cliente) & "px' class='MTD'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> VALOR DO OR�AMENTO
		s = formata_moeda(r("vl_total"))
		x = x & "		<TD valign='top' align='right' style='width:" & Cstr(w_valor) & "px' class='MTD'><P class='Cnd'>" & s & "</P></TD>" & chr(13)

	'> TOTALIZA��O DE VALORES
		vl_sub_total = vl_sub_total + r("vl_total")
		vl_total_geral = vl_total_geral + r("vl_total")
			
		x = x & "	</TR>" & chr(13)
			
		r.MoveNext
		loop

	
  ' MOSTRA TOTAL DO �LTIMO OR�AMENTISTA
	if n_reg <> 0 then 
		x = x & "	<TR style='background: #FFFFDD'>" & chr(13) & _
				"		<TD colspan='3' class='MTBE' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
				"		<TD class='MTBD'><p class='Cd'>" & formata_moeda(vl_sub_total) & "</p></td>" & chr(13) & _
				"	</TR>" & chr(13)
				
	'>	TOTAL GERAL
		if qtde_orcamentistas > 1 then
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

  ' MOSTRA AVISO DE QUE N�O H� DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR>" & chr(13) & _
				"		<TD class='MT' colspan='4'><P class='ALERTA'>&nbsp;NENHUM PR�-PEDIDO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
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

function fRELConcluir( id_orcamento ){
	window.status = "Aguarde ...";
	fREL.orcamento_selecionado.value=id_orcamento;
	fREL.action = "Orcamento.asp"
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
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>



<% else %>
<!-- ***************************************************** -->
<!-- **********  P�GINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Conclu�do';" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="orcamento_selecionado" id="orcamento_selecionado" value="">

<!--  I D E N T I F I C A � � O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pr�-Pedidos Em Aberto</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!--  RELAT�RIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a p�gina anterior">
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
