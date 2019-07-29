<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  R E L P E D I D O S A N T E R I O R E S E X E C . A S P
'     ========================================================
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
	if Not operacao_permitida(OP_CEN_CONSULTA_PEDIDOS_ANTERIORES_CLIENTE, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim s, s_filtro
	dim rb_op, cliente_selecionado, c_pedido
	dim s_cnpj_cpf, s_nome
	
	s_cnpj_cpf = ""
	s_nome = ""
		
	dim alerta
	alerta = ""

	rb_op = Trim(Request("rb_op"))
	cliente_selecionado = Trim(Request("cliente_selecionado"))
	c_pedido = Trim(Request.Form("c_pedido"))
	s = normaliza_num_pedido(c_pedido)
	if s <> "" then c_pedido=s
	
	select case rb_op
		case "-1"
			if cliente_selecionado = "" then alerta = "Cliente não especificado."

		case "4"
			if c_pedido = "" then
				alerta = "Especifique o número do pedido."
			elseif normaliza_num_pedido(c_pedido) = "" then
				alerta = "Número de pedido inválido."
			else
				s = "SELECT pedido, id_cliente FROM t_PEDIDO WHERE (pedido='" & c_pedido & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta = "Pedido " & c_pedido & " não está cadastrado."
				else
					cliente_selecionado = Trim("" & rs("id_cliente"))
					end if
				end if

		case else: alerta = "Opção de pesquisa inválida."
		end select

	if alerta = "" then
		s = "SELECT cnpj_cpf, nome_iniciais_em_maiusculas FROM t_CLIENTE WHERE (id='" & cliente_selecionado & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if Not rs.Eof then
			s_cnpj_cpf = Trim("" & rs("cnpj_cpf"))
			s_nome = Trim("" & rs("nome_iniciais_em_maiusculas"))
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
dim x, s_sql, s_where, cab, n_reg

	s_sql = "SELECT" & _
				" tP.pedido, tP.data, tP.loja, tP.vendedor, tP.indicador," & _
				" tOI.cnpj_cpf, tOI.razao_social_nome_iniciais_em_maiusculas" & _
			" FROM t_PEDIDO tP" & _
				" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR tOI ON (tp.indicador=tOI.apelido)"
	
	select case rb_op
		case "-1"
			if cliente_selecionado <> "" then s_where = " (tP.id_cliente = '" & cliente_selecionado & "')"
			
		case "4"
			if c_pedido <> "" then s_where = " (tP.pedido='" & c_pedido & "')"

		case else
			s_where = ""
		end select

	if s_where <> "" then s_sql = s_sql & " WHERE" & s_where
		
	s_sql = s_sql & " ORDER BY tP.data, tP.pedido"

  ' CABEÇALHO
	cab = "<TABLE cellSpacing=0>" & chr(13) & _
		  "	<TR style='background:azure' NOWRAP>" & _
		  "		<TD class='MDTE TdPedido' valign='bottom' NOWRAP><P class='R'>&nbsp;Nº PEDIDO</P></TD>" & _
		  "		<TD class='MTD TdData' align='center' valign='bottom' NOWRAP><P class='Rc'>DATA</P></TD>" & _
		  "		<TD class='MTD TdLoja' align='center' valign='bottom' NOWRAP><P class='Rc'>LOJA</P></TD>" & _
		  "		<TD class='MTD TdVendedor' valign='bottom' NOWRAP><P class='R'>&nbsp;VENDEDOR</P></TD>" & _
		  "		<TD class='MTD TdIndicador' valign='bottom'><P class='R'>&nbsp;INDICADOR</P></TD>" & _
		  "	</TR>" & chr(13)

	x = cab
	n_reg = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	  ' CONTAGEM
		n_reg = n_reg + 1

		x = x & "	<TR NOWRAP>"

	 '> Nº PEDIDO
		x = x & "		<TD class='MDTE TdPedido'><P class='C'>&nbsp;<a href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & _
				")' title='clique para consultar o pedido'>" & Trim("" & r("pedido")) & "</a></P></TD>"

	 '> DATA
		x = x & "		<TD align='center' class='MTD TdData'><P class='Cc'>" & formata_data(r("data")) & "</P></TD>"
	
	 '> LOJA
		x = x & "		<TD align='center' class='MTD TdLoja'><P class='Cc'>" & Trim("" & r("loja")) & "</P></TD>"

	 '> VENDEDOR
		x = x & "		<TD align='left' class='MTD TdVendedor'><P class='C'>" & Trim("" & r("vendedor")) & "</P></TD>"
	
	 '> INDICADOR
		x = x & "		<TD align='left' class='MTD TdIndicador'><P class='C'>" & Trim("" & r("indicador")) & "</P>"
		if Trim("" & r("cnpj_cpf")) <> "" then
			x = x & "<p class='Cn'>" & cnpj_cpf_formata(Trim("" & r("cnpj_cpf"))) & " - " & Trim("" & r("razao_social_nome_iniciais_em_maiusculas")) & "</p>"
			end if
		x = x & "</TD>"

		x = x & "</TR>" & chr(13)

		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
	
  ' MOSTRA TOTAL DE PEDIDOS
	if n_reg <> 0 then 
		x = x & "<TR NOWRAP style='background: #FFFFDD'>" & _
				"<TD COLSPAN='5' class='MT' NOWRAP><p class='C'>" & _
				"TOTAL:&nbsp;&nbsp;" & formata_inteiro(n_reg) & "&nbsp;pedidos</p></td>" & _
				"</TR>"
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab & _
			"	<TR NOWRAP>" & _
				"		<TD class='MT' colspan='5'><P class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</P></TD>" & _
				"	</TR>"
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

function fRELConcluir( id_pedido ){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=id_pedido;
	fREL.action = "pedido.asp"
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
.TdPedido {
	width:80px;
}
.TdData {
	width:90px;
}
.TdLoja {
	width:50px;
}
.TdVendedor {
	width:100px;
}
.TdIndicador {
 width:200px;
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
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
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
<input type="hidden" name="rb_op" id="rb_op" value="<%=rb_op%>">
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value="<%=cliente_selecionado%>">
<input type="hidden" name="c_pedido" id="c_pedido" value="<%=c_pedido%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pedidos Anteriormente Efetuados por um Cliente</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>"
	
	if len(s_cnpj_cpf) = 11 then s = "CPF" else s = "CNPJ"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>" & s & ":&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & cnpj_cpf_formata(s_cnpj_cpf) & "</p></td></tr>"
	
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Cliente:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & s_nome & "</p></td></tr>"

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
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
