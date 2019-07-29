<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelPesquisaPedidoNF.asp
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

	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_PESQUISA_PEDIDO_POR_OBS2, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim s, s_filtro
	dim c_nf
	dim i, n, s_filtro_nf
	
	dim alerta
	alerta = ""

	c_nf = Trim(Request.Form("c_nf"))
	
	if c_nf = "" then alerta = "Não foi informado o nº da NF para pesquisar"

	if alerta = "" then
		'Remove caracteres não numéricos à esquerda
		do while ( (Not IsDigit(left(c_nf, 1))) And (Len(c_nf) > 1) )
			c_nf = right(c_nf, len(c_nf)-1)
			loop

		'Remove zeros à esquerda
		do while ( (left(c_nf, 1) = "0") And (Len(c_nf) > 1) )
			c_nf = right(c_nf, len(c_nf)-1)
			loop

		'Como os campos 'obs_2' e 'obs_3' armazenam números em formato texto, tenta
		'maximizar a capacidade de pesquisa p/ os casos em que foram
		'cadastrados zeros à esquerda
		s_filtro_nf = "'" & c_nf & "'"
		for i = (Len(c_nf)+1) to MAX_OBS_2 'Tamanho do campo no BD
			s_filtro_nf = s_filtro_nf & ","
			n = MAX_OBS_2 - i + 1
			s_filtro_nf = s_filtro_nf & "'" & String(n,"0") & c_nf & "'"
			next

		s = "SELECT" & _
				" pedido," & _
				" data," & _
				" obs_2," & _
				" obs_3," & _
				" nome_iniciais_em_maiusculas," & _
                " id_nfe_emitente," & _
				" loja" & _
			" FROM t_PEDIDO INNER JOIN t_CLIENTE" & _
				" ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
			" WHERE" & _
				"(" & _
					" (obs_2 IN (" & s_filtro_nf & "))" & _
					" OR " & _
					"(obs_3 IN (" & s_filtro_nf & "))" & _
				")" & _
			" ORDER BY" & _
				" data," & _
				" pedido"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			alerta = "Nenhum pedido encontrado com nº de NF " & c_nf
		else
			if rs.RecordCount = 1 then
				if Trim("" & rs("loja")) = loja then
					Response.Redirect("pedido.asp?pedido_selecionado=" & rs("pedido") & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
				else
					if PossuiAcessoLoja(usuario, Trim("" & rs("loja"))) then Response.Redirect("PedidoConsulta.asp?pedido_selecionado=" & rs("pedido") & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
					end if
				end if
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
dim x, cab, n_reg, s_cor

  ' CABEÇALHO
	cab = "<TABLE cellSpacing=0>" & chr(13) & _
		  "	<TR style='background:azure' NOWRAP>" & _
		  "		<TD class='MDTE' align='center' valign='bottom' NOWRAP><P style='width:75px' class='Rc'>NF</P></TD>" & _
		  "		<TD class='MTD' align='center' valign='bottom' NOWRAP><P style='width:75px' class='Rc'>NF REMESSA</P></TD>" & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:75px' class='R'>&nbsp;Nº PEDIDO</P></TD>" & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:70px' class='Rc'>EMPRESA</P></TD>" & _
		  "		<TD class='MTD' align='center' valign='bottom' NOWRAP><P style='width:80px' class='Rc'>DATA</P></TD>" & _
		  "		<TD class='MTD' align='center' valign='bottom' NOWRAP><P style='width:45px' class='Rc'>LOJA</P></TD>" & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:200px' class='Rc'>CLIENTE</P></TD>" & _
		  "	</TR>" & chr(13)

	x = cab
	n_reg = 0
	
	do while Not rs.Eof
	  ' CONTAGEM
		n_reg = n_reg + 1

		x = x & "	<TR NOWRAP>"

	 '> NF
		s = Trim("" & rs("obs_2"))
		s_cor = "dimgray"
		if converte_numero(c_nf) = converte_numero(s) then s_cor = "blue"
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MDTE'><p class='Cc' style='color:" & s_cor & ";'>" & s & "</p></td>" & chr(13)

	 '> NF REMESSA
		s = Trim("" & rs("obs_3"))
		s_cor = "dimgray"
		if converte_numero(c_nf) = converte_numero(s) then s_cor = "blue"
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD'><p class='Cc' style='color:" & s_cor & ";'>" & s & "</p></td>" & chr(13)

	 '> Nº PEDIDO
		if Trim("" & rs("loja")) = loja then
		'	PÁGINA NORMAL DE CONSULTA, C/ BOTÕES P/ OPERAÇÕES DE EDITAR/REMOVER/ETC
			x = x & "		<TD class='MTD'><P class='C'>&nbsp;<a href='javascript:fRELConcluir(" & _
					chr(34) & Trim("" & rs("pedido")) & chr(34) & _
					")' title='clique para consultar o pedido'>" & Trim("" & rs("pedido")) & "</a></P></TD>"
		else
		'	PÁGINA EXCLUSIVAMENTE P/ VISUALIZAR OS DADOS DO PEDIDO
			x = x & "		<TD class='MTD'><P class='C'>&nbsp;<a href='javascript:fRELApenasConsultaConcluir(" & _
					chr(34) & Trim("" & rs("pedido")) & chr(34) & _
					")' title='clique para consultar o pedido'>" & Trim("" & rs("pedido")) & "</a></P></TD>"
			end if

    '> EMITENTE
		x = x & "		<TD align='center' valign='top' class='MTD'><P class='Cc'>" & obtem_apelido_empresa_NFe_emitente(Trim("" & rs("id_nfe_emitente"))) & "</P></TD>" & chr(13)

	 '> DATA
		x = x & "		<TD align='center' class='MTD'><P class='Cc'>" & formata_data(rs("data")) & "</P></TD>"
	
	 '> LOJA
		x = x & "		<TD align='center' class='MTD'><P class='Cc'>" & Trim("" & rs("loja")) & "</P></TD>"
	
	 '> CLIENTE
		x = x & "		<TD class='MTD'><P class='Cn'>" & Trim("" & rs("nome_iniciais_em_maiusculas")) & "</P></TD>"
	
		x = x & "</TR>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		rs.MoveNext
		loop
		
  ' MOSTRA TOTAL DE PEDIDOS
	if n_reg <> 0 then 
		x = x & "<TR NOWRAP style='background: #FFFFDD'>" & _
				"<TD COLSPAN='7' class='MT' NOWRAP><p class='C'>" & _
				"TOTAL:&nbsp;&nbsp;" & formata_inteiro(n_reg) & "&nbsp;pedidos</p></td>" & _
				"</TR>"
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab & _
			"	<TR NOWRAP>" & _
				"		<TD class='MT' colspan='7'><P class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</P></TD>" & _
				"	</TR>"
		end if

  ' FECHA TABELA
	x = x & "</TABLE>"
	
	Response.write x

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
	fREL.action = "pedido.asp"
	fREL.submit(); 
}

function fRELApenasConsultaConcluir( id_pedido ){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=id_pedido;
	fREL.action = "PedidoConsulta.asp"
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
<input type="hidden" name="c_nf" id="c_nf" value="<%=c_nf%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="749" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pesquisa Pedido pelo Nº NF</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='749' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>"
	
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Nº NF:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & c_nf & "</p></td></tr>"

	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Data da Consulta:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & formata_data_hora(Now) & "</p></td></tr>"
	
	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="749" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="749" cellSpacing="0">
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
