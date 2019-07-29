<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L P R O D A L O C P E D S P L I T A V E L . A S P
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
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_PRODUTOS_SPLIT_POSSIVEL, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim s, alerta
	alerta = ""




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' EXECUTA RELATORIO
'
sub executa_relatorio
dim r
dim s, s_aux, s_sql, x, cab_table, cab, fabricante_a
dim n_reg, n_reg_total, n_saldo

'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto" & _
			", descricao, descricao_html, Sum(qtde) AS saldo" & _
			" FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque='" & ID_ESTOQUE_VENDIDO & "')" & _
			" AND (t_PEDIDO.st_entrega = '" & ST_ENTREGA_SPLIT_POSSIVEL & "')"
	
	s_sql = s_sql & " GROUP BY t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, descricao, descricao_html" & _
					" ORDER BY t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, descricao, descricao_html"

  ' CABEÇALHO
	cab_table = "<TABLE class='Q' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD width='75' valign='bottom' NOWRAP class='MD MB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
		  "		<TD width='480' valign='bottom' NOWRAP class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
		  "		<TD width='60' valign='bottom' NOWRAP class='MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table & cab
	n_reg = 0
	n_reg_total = 0
	n_saldo = 0
	fabricante_a = "XXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante")) <> fabricante_a then
			fabricante_a = Trim("" & r("fabricante"))
			s = Trim("" & r("fabricante"))
			s_aux = ucase(x_fabricante(s))
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR NOWRAP>" & chr(13) & _
					"		<TD class='MB' align='center' colspan='3'><P class='C'>&nbsp;" & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
			
		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='C'>&nbsp;" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "		<TD class='MDB' valign='bottom'><P class='C' NOWRAP>&nbsp;" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

	 '> SALDO
		x = x & "		<TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(r("saldo")) & "</P></TD>" & chr(13)

		n_saldo = n_saldo + r("saldo")
		
		x = x & "	</TR>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP style='background: #FFFFDD'>"  & chr(13) & _
				"		<TD COLSPAN='3' NOWRAP><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & formata_inteiro(n_saldo) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='3'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
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

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';
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

<form id="fESTOQ" name="fESTOQ" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Produtos Split Possível</span>
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

<!--  RELATÓRIO  -->
<br>
<%	
	executa_relatorio
%>

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
