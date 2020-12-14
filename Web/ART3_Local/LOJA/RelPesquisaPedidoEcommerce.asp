<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelPesquisaPedidoEcommerce.asp
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
	
	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim s, sql_interno, s_filtro
	dim c_num_pedido_aux, c_tipo_num_pedido, pedido_selecionado, loja_selecionada
	
	dim alerta
	alerta = ""

	c_num_pedido_aux = Trim(Request.Form("c_num_pedido_aux"))
    c_tipo_num_pedido = Trim(Request.Form("c_tipo_num_pedido"))
	
	if c_num_pedido_aux = "" then alerta = "Parâmetro de pesquisa não foi fornecido."

	if alerta = "" then
		if Len(c_num_pedido_aux) = TAMANHO_PEDIDO_MAGENTO then
			'VERIFICA SE O PREFIXO DO Nº MAGENTO CORRESPONDE À LOJA
			if Left(c_num_pedido_aux,1) = PREFIXO_PEDIDO_MAGENTO_AR_CLUBE then
				if c_tipo_num_pedido = OP_PESQ_PEDIDO_MAGENTO_BONSHOP then c_tipo_num_pedido = OP_PESQ_PEDIDO_MAGENTO_AR_CLUBE
			elseif Left(c_num_pedido_aux,1) = PREFIXO_PEDIDO_MAGENTO_BONSHOP then
				if c_tipo_num_pedido = OP_PESQ_PEDIDO_MAGENTO_AR_CLUBE then c_tipo_num_pedido = OP_PESQ_PEDIDO_MAGENTO_BONSHOP
				end if
			end if
		end if

	if alerta = "" then
		if c_tipo_num_pedido = OP_PESQ_PEDIDO_MAGENTO_BONSHOP then
			s = "SELECT TOP 1" & _
					" pedido_bs_x_ac" & _
				" FROM t_PEDIDO" & _
					" INNER JOIN t_LOJA ON (t_PEDIDO.loja = t_LOJA.loja)" & _
				" WHERE" & _
					" (unidade_negocio = '" & COD_UNIDADE_NEGOCIO_LOJA__BS & "')"

			if Len(c_num_pedido_aux) = TAMANHO_PEDIDO_MAGENTO then
				s = s & _
					" AND (pedido_bs_x_ac = '" & Trim("" & c_num_pedido_aux) & "')"
			else
				s = s & _
					" AND (pedido_bs_x_ac_reverso LIKE '" & StrReverse(Cstr(c_num_pedido_aux)) & "%')"
				end if

			s = s & _
				" ORDER BY" & _
					" data_hora DESC"

			sql_interno = s

			s = "SELECT " & _
					" pedido," & _
					" st_entrega," & _
					" pedido_bs_x_ac AS numEC," & _
					" data,"

			if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
				s = s & _
					" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
			else
				s = s & _
					" nome_iniciais_em_maiusculas,"
				end if

			s = s & _
					" t_PEDIDO.loja AS loja," & _
					" (" & _
						"SELECT" & _
							" Coalesce(Sum(qtde*preco_NF),0)" & _
						" FROM t_PEDIDO_ITEM" & _
						" WHERE" & _
							" t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido" & _
					") AS vl_pedido" & _
				" FROM t_PEDIDO INNER JOIN t_CLIENTE" & _
					" ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
					" INNER JOIN t_LOJA ON (t_PEDIDO.loja = t_LOJA.loja)" & _
				" WHERE" & _
					" (unidade_negocio = '" & COD_UNIDADE_NEGOCIO_LOJA__BS & "')" & _
					" AND (pedido_bs_x_ac IN (" & sql_interno & "))" & _
				" ORDER BY" & _
					" data_hora DESC"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta = "Nenhum pedido encontrado (Nº Pedido Magento: " & c_num_pedido_aux & ")"
			else
				pedido_selecionado = Trim("" & rs("pedido"))
				loja_selecionada = Trim("" & rs("loja"))
				rs.MoveNext
				'Se for o único registro encontrado, segue para a página do pedido
				if rs.Eof then
					if loja_selecionada = loja then
						Response.Redirect("pedido.asp?pedido_selecionado=" & pedido_selecionado & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
					else
						if PossuiAcessoLoja(usuario, loja_selecionada) then 
							Response.Redirect("PedidoConsulta.asp?pedido_selecionado=" & pedido_selecionado & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
						else
							cn.Close
							Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
							end if
						end if
				else
					'Posiciona novamente no início do recordset para exibir a lista de pedidos
					rs.MoveFirst
					end if
				end if
        elseif c_tipo_num_pedido = OP_PESQ_PEDIDO_MAGENTO_AR_CLUBE then
		    s = "SELECT TOP 1" & _
					" pedido_bs_x_ac" & _
			    " FROM t_PEDIDO" & _
			    " WHERE" & _
					" (loja = '" & NUMERO_LOJA_ECOMMERCE_AR_CLUBE & "')"

			if Len(c_num_pedido_aux) = TAMANHO_PEDIDO_MAGENTO then
				s = s & _
					" AND (pedido_bs_x_ac = '" & Trim("" & c_num_pedido_aux) & "')"
			else
				s = s & _
				    " AND (pedido_bs_x_ac_reverso LIKE '" & StrReverse(Cstr(c_num_pedido_aux)) & "%')"
				end if

			s = s & _
			    " ORDER BY" & _
				    " data_hora DESC"

			sql_interno = s

			s = "SELECT " & _
				    " pedido," & _
					" st_entrega," & _
					" pedido_bs_x_ac AS numEC," & _
				    " data,"

			if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
				s = s & _
					" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
			else
				s = s & _
					" nome_iniciais_em_maiusculas,"
				end if

			s = s & _
				    " loja" & _
			    " FROM t_PEDIDO INNER JOIN t_CLIENTE" & _
				    " ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
			    " WHERE" & _
					" (loja = '" & NUMERO_LOJA_ECOMMERCE_AR_CLUBE & "')" & _
					" AND (pedido_bs_x_ac IN (" & sql_interno & "))" & _
			    " ORDER BY" & _
				    " data_hora DESC"
		    if rs.State <> 0 then rs.Close
		    rs.open s, cn
		    if rs.Eof then
			    alerta = "Nenhum pedido encontrado (Nº Pedido Magento: " & c_num_pedido_aux & ")"
		    else
				pedido_selecionado = Trim("" & rs("pedido"))
				loja_selecionada = Trim("" & rs("loja"))
				rs.MoveNext
				'Se for o único registro encontrado, segue para a página do pedido
				if rs.Eof then
					if loja_selecionada = loja then
						Response.Redirect("pedido.asp?pedido_selecionado=" & pedido_selecionado & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
					else
						if PossuiAcessoLoja(usuario, loja_selecionada) then 
							Response.Redirect("PedidoConsulta.asp?pedido_selecionado=" & pedido_selecionado & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
						else
							cn.Close
							Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
							end if
						end if
				else
					'Posiciona novamente no início do recordset para exibir a lista de pedidos
					rs.MoveFirst
					end if
				end if
		elseif c_tipo_num_pedido = OP_PESQ_PEDIDO_MARKETPLACE_AR_CLUBE then
			 s = "SELECT" & _
					" pedido," & _
					" st_entrega," & _
					" pedido_bs_x_marketplace AS numEC," & _
					" data,"

			if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
				s = s & _
					" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
			else
				s = s & _
					" nome_iniciais_em_maiusculas,"
				end if

			s = s & _
					" loja" & _
				" FROM t_PEDIDO INNER JOIN t_CLIENTE" & _
					" ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
				" WHERE" & _
					" (loja = '" & NUMERO_LOJA_ECOMMERCE_AR_CLUBE & "')" & _
					" AND (pedido_bs_x_marketplace = '" & Cstr(c_num_pedido_aux) & "')" & _
				" ORDER BY" & _
					" data_hora DESC"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta = "Nenhum pedido encontrado (Nº Pedido Marketplace: " & c_num_pedido_aux & ")"
			else
				pedido_selecionado = Trim("" & rs("pedido"))
				loja_selecionada = Trim("" & rs("loja"))
				rs.MoveNext
				'Se for o único registro encontrado, segue para a página do pedido
				if rs.Eof then
					if loja_selecionada = loja then
						Response.Redirect("pedido.asp?pedido_selecionado=" & pedido_selecionado & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
					else
						if PossuiAcessoLoja(usuario, loja_selecionada) then 
							Response.Redirect("PedidoConsulta.asp?pedido_selecionado=" & pedido_selecionado & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
						else
							cn.Close
							Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
							end if
						end if
				else
					'Posiciona novamente no início do recordset para exibir a lista de pedidos
					rs.MoveFirst
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
		  "		<TD class='MDTE' valign='bottom' NOWRAP><P style='width:80px' class='R'>&nbsp;Nº PEDIDO</P></TD>" & _
		  "		<TD class='MTD' align='left' valign='bottom' NOWRAP><P style='width:100px' class='R'>Nº E-COMMERCE</P></TD>" & _
		  "		<TD class='MTD' align='left' valign='bottom' NOWRAP><P style='width:100px' class='R'>STATUS</P></TD>" & _
		  "		<TD class='MTD' align='center' valign='bottom' NOWRAP><P style='width:90px' class='Rc'>DATA</P></TD>" & _
		  "		<TD class='MTD' align='center' valign='bottom' NOWRAP><P style='width:50px' class='Rc'>LOJA</P></TD>" & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:200px' class='Rc'>CLIENTE</P></TD>" & _
		  "	</TR>" & chr(13)

	x = cab
	n_reg = 0
	
	do while Not rs.Eof
	  ' CONTAGEM
		n_reg = n_reg + 1

		x = x & "	<TR NOWRAP>"

	 '> Nº PEDIDO
		if Trim("" & rs("loja")) = loja then
		'	PÁGINA NORMAL DE CONSULTA, C/ BOTÕES P/ OPERAÇÕES DE EDITAR/REMOVER/ETC
			x = x & "		<TD class='MDTE'><P class='C'>&nbsp;<a href='javascript:fRELConcluir(" & _
					chr(34) & Trim("" & rs("pedido")) & chr(34) & _
					")' title='clique para consultar o pedido'>" & Trim("" & rs("pedido")) & "</a></P></TD>"
		else
		'	PÁGINA EXCLUSIVAMENTE P/ VISUALIZAR OS DADOS DO PEDIDO
			x = x & "		<TD class='MDTE'><P class='C'>&nbsp;<a href='javascript:fRELApenasConsultaConcluir(" & _
					chr(34) & Trim("" & rs("pedido")) & chr(34) & _
					")' title='clique para consultar o pedido'>" & Trim("" & rs("pedido")) & "</a></P></TD>"
			end if

	 '> Nº PEDIDO E-COMMERCE
		x = x & "		<TD align='left' class='MTD'><P class='C'>" & Trim("" & rs("numEC")) & "</P></TD>"

	 '> STATUS
		s_cor = x_status_entrega_cor(Trim("" & rs("st_entrega")), Trim("" & rs("pedido")))
		x = x & "		<TD align='left' class='MTD'><P class='C' style='color:" & s_cor & ";'>" & x_status_entrega(Trim("" & rs("st_entrega"))) & "</P></TD>"

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
				"<TD COLSPAN='6' class='MT' NOWRAP><p class='C'>" & _
				"TOTAL:&nbsp;&nbsp;" & formata_inteiro(n_reg) & "&nbsp;pedidos</p></td>" & _
				"</TR>"
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab & _
			"	<TR NOWRAP>" & _
				"		<TD class='MT' colspan='6'><P class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</P></TD>" & _
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
<input type="hidden" name="c_num_pedido_aux" id="c_num_pedido_aux" value="<%=c_num_pedido_aux%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="749" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pesquisa por Nº Pedido E-Commerce</span>
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
			   "<p class='N'>Nº E-Commerce Pesquisado:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & c_num_pedido_aux & "</p></td></tr>"

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
