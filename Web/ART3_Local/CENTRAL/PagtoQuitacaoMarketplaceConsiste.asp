<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ====================================================
'	  PagtoQuitacaoMarketplaceConsiste.asp
'     ====================================================
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


	class cl_PAGTO_QUITACAO_MARKETPLACE_CONSISTE
		dim pedido_marketplace
		dim pedido_ERP
		dim qtde_pedido_ERP_associados
		dim vl_total_familia_preco_NF
		dim vl_total_familia_pago
		dim vl_total_familia_devolucao_preco_NF
		dim msg_erro
		end class

	On Error GoTo 0
	Err.Clear

	dim s, usuario, lista_pedidos_marketplace, lista_pedidos_ERP, lista_pedido_marketplace_x_ERP, v_pedido, v_aux, i, j, achou
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
	dim alerta
	alerta = ""

	dim erro_consistencia, msg_erro_consistencia
	erro_consistencia = False
	msg_erro_consistencia = ""

	dim observacoes
	observacoes = ""
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim s_sql, s_where
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	lista_pedidos_marketplace = ucase(Trim(request("c_pedidos_quitacao_marketplace")))
	if (lista_pedidos_marketplace = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
	lista_pedidos_marketplace=substitui_caracteres(lista_pedidos_marketplace,chr(10),"")
	v_aux = split(lista_pedidos_marketplace,chr(13),-1)
	achou=False
	for i=Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			achou = True
			s = Trim(v_aux(i))
			if s <> "" then v_aux(i) = s
			end if
		next

	if Not achou then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	redim v_pedido(0)
	set v_pedido(Ubound(v_pedido))=New cl_PAGTO_QUITACAO_MARKETPLACE_CONSISTE
	with v_pedido(Ubound(v_pedido))
		.pedido_marketplace = ""
		.pedido_ERP = ""
		.qtde_pedido_ERP_associados = 0
		.vl_total_familia_preco_NF = 0
		.vl_total_familia_pago = 0
		.vl_total_familia_devolucao_preco_NF = 0
		.msg_erro = ""
		end with
	
	for i = Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			if Trim(v_pedido(Ubound(v_pedido)).pedido_marketplace)<>"" then
				redim preserve v_pedido(Ubound(v_pedido)+1)
				set v_pedido(Ubound(v_pedido))=New cl_PAGTO_QUITACAO_MARKETPLACE_CONSISTE
				end if
			with v_pedido(Ubound(v_pedido))
				.pedido_marketplace = Trim(v_aux(i))
				.pedido_ERP = ""
				.qtde_pedido_ERP_associados = 0
				.vl_total_familia_preco_NF = 0
				.vl_total_familia_pago = 0
				.vl_total_familia_devolucao_preco_NF = 0
				.msg_erro = ""
				end with
			end if
		next
	
'	VERIFICA SE HÁ Nº PEDIDOS DUPLICADOS
	for i=LBound(v_pedido) to UBound(v_pedido)
		for j=LBound(v_pedido) to (i-1)
			if Trim("" & v_pedido(i).pedido_marketplace) <> "" then
				if Trim("" & v_pedido(i).pedido_marketplace) = Trim("" & v_pedido(j).pedido_marketplace) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Nº pedido marketplace repetido: " & Trim("" & v_pedido(i).pedido_marketplace) & " nas linhas " & renumera_com_base1(Lbound(v_pedido),j) & " e " & renumera_com_base1(Lbound(v_pedido),i)
					end if
				end if
			next
		next

'	A PARTIR DO Nº DO PEDIDO MARKETPLACE, PESQUISA PELO Nº PEDIDO USADO NO SISTEMA
	if alerta = "" then
		s_where = ""
		for i=LBound(v_pedido) to UBound(v_pedido)
			if Trim("" & v_pedido(i).pedido_marketplace) <> "" then
				if s_where <> "" then s_where = s_where & ","
				s_where = s_where & "'" & Trim("" & v_pedido(i).pedido_marketplace) & "'"
				end if
			next

		s_sql = "SELECT DISTINCT" & _
					" pedido_base," & _
					" pedido_bs_x_marketplace," & _
					" (SELECT Count(*) FROM (SELECT DISTINCT pedido_base FROM t_PEDIDO t_PEDIDO__AUX WHERE (t_PEDIDO__AUX.pedido_bs_x_marketplace = t_PEDIDO.pedido_bs_x_marketplace) AND (t_PEDIDO__AUX.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')) t) AS qtde_pedidos_erp_associados," & _
					" SUBSTRING((SELECT DISTINCT ', ' + t_PEDIDO__AUX.pedido_base FROM t_PEDIDO t_PEDIDO__AUX WHERE (t_PEDIDO__AUX.pedido_bs_x_marketplace = t_PEDIDO.pedido_bs_x_marketplace) AND (t_PEDIDO__AUX.st_entrega <> '" & ST_ENTREGA_CANCELADO & "') ORDER BY ', ' + t_PEDIDO__AUX.pedido_base FOR XML PATH ('')), 3, 1024) AS pedidos_erp_associados" & _
				" FROM t_PEDIDO" & _
				" WHERE" & _
					" (pedido_bs_x_marketplace IN (" & s_where & "))"
		if rs.State <> 0 then rs.Close
		rs.open s_sql, cn
		do while Not rs.Eof
			for i=LBound(v_pedido) to UBound(v_pedido)
				if Trim("" & v_pedido(i).pedido_marketplace) <> "" then
					if Trim("" & v_pedido(i).pedido_marketplace) = Trim("" & rs("pedido_bs_x_marketplace")) then
						v_pedido(i).qtde_pedido_ERP_associados = CLng(rs("qtde_pedidos_erp_associados"))
						if v_pedido(i).qtde_pedido_ERP_associados = 1 then
							v_pedido(i).pedido_ERP = Trim("" & rs("pedido_base"))
						else
							if Trim("" & v_pedido(i).pedido_ERP) = "" then
								v_pedido(i).pedido_ERP = Trim("" & rs("pedidos_erp_associados"))
							else
								erro_consistencia = True
								msg_erro_consistencia = texto_add_br(msg_erro_consistencia)
								msg_erro_consistencia = msg_erro_consistencia & "O pedido marketplace " & Trim("" & v_pedido(i).pedido_marketplace) & " está associado com mais de um pedido ERP (" & Trim("" & rs("pedidos_erp_associados")) & ")"
								end if
							end if
						exit for
						end if
					end if
				next
			rs.MoveNext
			loop
		end if

'	HÁ PEDIDOS EM QUE NÃO FOI POSSÍVEL DETERMINAR O Nº PEDIDO NO SISTEMA?
	if alerta = "" then
		for i=LBound(v_pedido) to UBound(v_pedido)
			if Trim("" & v_pedido(i).pedido_marketplace) <> "" then
				if Trim("" & v_pedido(i).pedido_ERP) = "" then
					erro_consistencia = True
					msg_erro_consistencia = texto_add_br(msg_erro_consistencia)
					msg_erro_consistencia = msg_erro_consistencia & "Não foi localizado o nº pedido ERP para o pedido marketplace " & Trim("" & v_pedido(i).pedido_marketplace)
					end if
				end if
			next
		end if

'	HÁ DUPLICIDADE NO Nº PEDIDO ERP?
	if (alerta = "") And (Not erro_consistencia) then
		for i=LBound(v_pedido) to UBound(v_pedido)
			for j=LBound(v_pedido) to (i-1)
				if Trim("" & v_pedido(i).pedido_ERP) <> "" then
					if Trim("" & v_pedido(i).pedido_ERP) = Trim("" & v_pedido(j).pedido_ERP) then
						erro_consistencia = True
						msg_erro_consistencia = texto_add_br(msg_erro_consistencia)
						msg_erro_consistencia = msg_erro_consistencia & "Nº pedido ERP repetido: " & Trim("" & v_pedido(i).pedido_ERP) & " nas linhas " & renumera_com_base1(Lbound(v_pedido),j) & " e " & renumera_com_base1(Lbound(v_pedido),i)
						end if
					end if
				next
			next
		end if

	redim v_aux(Ubound(v_pedido))
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		v_aux(i)=v_pedido(i).pedido_marketplace
		next
	lista_pedidos_marketplace = join(v_aux,",")
	
	redim v_aux(Ubound(v_pedido))
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		v_aux(i)=v_pedido(i).pedido_ERP
		next
	lista_pedidos_ERP = join(v_aux,",")

	lista_pedido_marketplace_x_ERP = ""
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		if Trim("" & v_pedido(i).pedido_marketplace) <> "" then
			if lista_pedido_marketplace_x_ERP <> "" then lista_pedido_marketplace_x_ERP = lista_pedido_marketplace_x_ERP & "|"
			lista_pedido_marketplace_x_ERP = lista_pedido_marketplace_x_ERP & Trim("" & v_pedido(i).pedido_marketplace) & "=" & Trim("" & v_pedido(i).pedido_ERP)
			end if
		next

	dim vl_total_pedido, vl_total_pago, vl_total_devolucao
	dim vl_TotalFamiliaPrecoVenda_aux, vl_TotalFamiliaPrecoNF_aux, vl_TotalFamiliaPago_aux, vl_TotalFamiliaDevolucaoPrecoVenda_aux, vl_TotalFamiliaDevolucaoPrecoNF_aux
	dim vl_total_saldo, st_pagto, msg_erro
	
	vl_total_pedido = 0
	vl_total_pago = 0
	vl_total_devolucao = 0
	vl_total_saldo = 0
	if alerta = "" then
		for i = Lbound(v_pedido) to Ubound(v_pedido)
			with v_pedido(i)
				if (.pedido_ERP <> "") And (.qtde_pedido_ERP_associados = 1) then
					s = "SELECT pedido FROM t_PEDIDO WHERE (pedido='" & .pedido_ERP & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & .pedido_ERP & " não está cadastrado."
					else
					'	OBTÉM OS VALORES A PAGAR, JÁ PAGO E O STATUS DE PAGAMENTO (PARA TODA A FAMÍLIA DE PEDIDOS)
						if Not calcula_pagamentos(.pedido_ERP, vl_TotalFamiliaPrecoVenda_aux, vl_TotalFamiliaPrecoNF_aux, vl_TotalFamiliaPago_aux, vl_TotalFamiliaDevolucaoPrecoVenda_aux, vl_TotalFamiliaDevolucaoPrecoNF_aux, st_pagto, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						.vl_total_familia_preco_NF = vl_TotalFamiliaPrecoNF_aux
						.vl_total_familia_pago = vl_TotalFamiliaPago_aux
						.vl_total_familia_devolucao_preco_NF = vl_TotalFamiliaDevolucaoPrecoNF_aux
						vl_total_pedido = vl_total_pedido + vl_TotalFamiliaPrecoNF_aux
						vl_total_pago = vl_total_pago + vl_TotalFamiliaPago_aux
						vl_total_devolucao = vl_total_devolucao + vl_TotalFamiliaDevolucaoPrecoNF_aux
						if vl_TotalFamiliaPrecoNF_aux > (vl_TotalFamiliaPago_aux+vl_TotalFamiliaDevolucaoPrecoNF_aux) then vl_total_saldo = vl_total_saldo + (vl_TotalFamiliaPrecoNF_aux-vl_TotalFamiliaPago_aux-vl_TotalFamiliaDevolucaoPrecoNF_aux)
						if st_pagto = ST_PAGTO_PAGO then
							observacoes = texto_add_br(observacoes)
							observacoes = observacoes & "A família de pedidos " & retorna_num_pedido_base(.pedido_ERP) & " já consta como quitada."
							end if
						end if
					end if
				end with
			next
		end if

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
function fPAGTOConfirma( f ) {
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
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
<!-- ***************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR DADOS PARA CONFIRMAÇÃO  ********** -->
<!-- ***************************************************************** -->
<body onload="focus();">
<center>

<form id="fPAGTO" name="fPAGTO" method="post" action="PagtoQuitacaoMarketplaceConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedidos_selecionados_marketplace" id="pedidos_selecionados_marketplace" value="<%=lista_pedidos_marketplace%>">
<input type="hidden" name="pedidos_selecionados_ERP" id="pedidos_selecionados_ERP" value="<%=lista_pedidos_ERP%>">
<input type="hidden" name="lista_pedido_marketplace_x_ERP" id="lista_pedido_marketplace_x_ERP" value="<%=lista_pedido_marketplace_x_ERP%>">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Quitação de Pedidos (Marketplace)<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!-- ************   HÁ MENSAGEM DE ERRO DE CONSISTÊNCIA?  ************ -->
<% if msg_erro_consistencia <> "" then %>
		<span class="Lbl" style="color:red;">ERROS DE INCONSISTÊNCIA</span>
		<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'><%=msg_erro_consistencia%></p></div>
		<br><br>
<% end if %>

<!-- ************   HÁ OBSERVAÇÕES?  ************ -->
<% if observacoes <> "" then %>
		<span class="Lbl">OBSERVAÇÕES</span>
		<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'><%=observacoes%></p></div>
		<br><br>
<% end if %>


<!--  PEDIDOS  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
		<td class="MB"><span class="PLTe">&nbsp;</span></td>
		<td class="MB"><span class="PLTe">&nbsp;</span></td>
		<td colspan="4" class="MT" NOWRAP align="center"><span class="PLTd">Valores Totalizados por Família de Pedidos</span></td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td class="MEB" style="border-right:0;" NOWRAP><span class="PLTe">Nº Pedido</span><br /><span class="PLTe">Marketplace</span></td>
		<td class="MEB" style="border-right:0;" NOWRAP><span class="PLTe">Nº Pedido</span><br /><span class="PLTe">ERP</span></td>
		<td class="MDBE" NOWRAP align="right"><span class="PLTd">Valor do Pedido</span></td>
		<td class="MDB" NOWRAP style="border-left:0;" align="right"><span class="PLTd">Valor Já Pago</span></td>
		<td class="MDB" NOWRAP style="border-left:0;" align="right"><span class="PLTd">&nbsp;Valor das Devoluções</span></td>
		<td class="MDB" NOWRAP style="border-left:0;" align="right"><span class="PLTd">Saldo a Pagar</span></td>
	</tr>
<% for i=Lbound(v_pedido) to Ubound(v_pedido)
		with v_pedido(i)
			if .pedido_marketplace <> "" then %>
			<tr bgColor="#FFFFFF">
				<input type="hidden" name="c_pedido" id="c_pedido" readonly tabindex=-1 class="PLLe" style="width:60px;margin-left:2pt;" value="<%=.pedido_ERP%>">
				<td class="MEB"><input name="c_pedido_marketplace" id="c_pedido_marketplace" readonly tabindex=-1 class="PLLe" style="width:100px;margin-left:2pt;" 
					value="<%=.pedido_marketplace%>"></td>
				<td class="MEB C" style="width:70px;word-break:break-word;"><%=.pedido_ERP%></td>
				<td class="MDBE" NOWRAP align="right"><input name="c_vl_pedido" id="c_vl_pedido" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;" 
					value="<%=formata_moeda(.vl_total_familia_preco_NF)%>"></td>
				<td class="MDB" NOWRAP align="right"><input name="c_vl_pago" id="c_vl_pago" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;<% if .vl_total_familia_pago < 0 then Response.Write "color:red;"%>" 
					value="<%=formata_moeda(.vl_total_familia_pago)%>"></td>
				<td class="MDB" NOWRAP align="right"><input name="c_vl_devolucao" id="c_vl_devolucao" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;" 
					value="<%=formata_moeda(.vl_total_familia_devolucao_preco_NF)%>"></td>
				<td class="MDB" NOWRAP align="right"><input name="c_vl_saldo" id="c_vl_saldo" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;<%if .vl_total_familia_preco_NF < (.vl_total_familia_pago+.vl_total_familia_devolucao_preco_NF) then Response.Write "color:red;"%>" 
					value="<%=formata_moeda(.vl_total_familia_preco_NF - .vl_total_familia_pago - .vl_total_familia_devolucao_preco_NF)%>"></td>
			</tr>
<%				end if
			end with
		next		%>
	<tr bgColor="#FFFFFF">
		<td><span class="PLTe">&nbsp;</span></td>
		<td><span class="PLTe">&nbsp;</span></td>
		<td class="MDBE" NOWRAP align="right"><input name="c_vl_total_pedido" id="c_vl_total_pedido" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;color:navy;" 
			value="<%=formata_moeda(vl_total_pedido)%>"></td>
		<td class="MDB" NOWRAP align="right"><input name="c_vl_total_pago" id="c_vl_total_pago" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;color:<%if vl_total_pago >= 0 then Response.Write "navy" else Response.Write "red"%>;" 
			value="<%=formata_moeda(vl_total_pago)%>"></td>
		<td class="MDB" NOWRAP align="right"><input name="c_vl_total_devolucao" id="c_vl_total_devolucao" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;color:navy;" 
			value="<%=formata_moeda(vl_total_devolucao)%>"></td>
		<td class="MDB" NOWRAP align="right"><input name="c_vl_total_saldo" id="c_vl_total_saldo" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;color:<%if vl_total_saldo >= 0 then Response.Write "navy" else Response.Write "red"%>;" 
			value="<%=formata_moeda(vl_total_saldo)%>"></td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<% if erro_consistencia then %>
<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
<% else %>
<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
	<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPAGTOConfirma(fPAGTO)" title="confirma o registro do pagamento">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
<% end if %>
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