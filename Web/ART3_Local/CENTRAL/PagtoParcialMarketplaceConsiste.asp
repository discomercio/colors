<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ====================================================
'	  PagtoParcialMarketplaceConsiste.asp
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


	class cl_PAGTO_PARCIAL_MARKETPLACE_CONSISTE
		dim pedido_marketplace
		dim vl_parcela
		dim pedido_ERP
		dim st_pagto
		dim qtde_pedido_ERP_associados
		dim vl_total_familia_preco_NF
		dim vl_total_familia_pago
		dim vl_total_familia_devolucao_preco_NF
		dim msg_erro
		end class

	On Error GoTo 0
	Err.Clear

	dim s, s_aux
	dim usuario, lista_pedidos_marketplace, lista_vl_parcela, vl_aux, s_vl_aux
	dim v_pedido, v_aux, i, j
	dim qtde_pedidos, qtde_vl_parcela
	dim c_dados_pagto

	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
	dim alerta
	alerta = ""

	dim erro_consistencia, msg_erro_consistencia
	erro_consistencia = False
	msg_erro_consistencia = ""

	dim observacoes, observacoes_aux
	observacoes = ""
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim s_sql, s_where
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	lista_pedidos_marketplace = ucase(Trim(request("c_pedidos_pagto_parcial_marketplace")))
	if (lista_pedidos_marketplace = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	lista_vl_parcela = ucase(Trim(request("c_valor_pagto_parcial_marketplace")))
	if lista_vl_parcela = "" then Response.Redirect("aviso.asp?id=" & ERR_PARAMETRO_OBRIGATORIO_NAO_ESPECIFICADO)

'	CARREGA RELAÇÃO DE PEDIDOS EM UM VETOR
	lista_pedidos_marketplace=substitui_caracteres(lista_pedidos_marketplace,chr(10),"")
	v_aux = split(lista_pedidos_marketplace,chr(13),-1)
	qtde_pedidos = 0
	for i=Lbound(v_aux) to Ubound(v_aux)
		v_aux(i) = Trim(v_aux(i))
		if v_aux(i) <> "" then
			qtde_pedidos = qtde_pedidos + 1
			end if
		next

	if qtde_pedidos = 0 then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

'	TRANSFERE A LISTA DE PEDIDOS DO VETOR AUXILIAR PARA O VETOR COM DADOS COMPLETOS
	redim v_pedido(0)
	set v_pedido(Ubound(v_pedido))=New cl_PAGTO_PARCIAL_MARKETPLACE_CONSISTE
	with v_pedido(Ubound(v_pedido))
		.pedido_marketplace = ""
		.vl_parcela = 0
		.pedido_ERP = ""
		.st_pagto = ""
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
				set v_pedido(Ubound(v_pedido))=New cl_PAGTO_PARCIAL_MARKETPLACE_CONSISTE
				end if
			with v_pedido(Ubound(v_pedido))
				.pedido_marketplace = Trim(v_aux(i))
				.vl_parcela = 0
				.pedido_ERP = ""
				.st_pagto = ""
				.qtde_pedido_ERP_associados = 0
				.vl_total_familia_preco_NF = 0
				.vl_total_familia_pago = 0
				.vl_total_familia_devolucao_preco_NF = 0
				.msg_erro = ""
				end with
			end if
		next

'	CARREGA RELAÇÃO DE PARCELAS EM UM VETOR
	lista_vl_parcela=substitui_caracteres(lista_vl_parcela,chr(10),"")
	v_aux = split(lista_vl_parcela,chr(13),-1)
	qtde_vl_parcela = 0
	for i=Lbound(v_aux) to Ubound(v_aux)
		v_aux(i) = Trim(v_aux(i))
		if v_aux(i) <> "" then
			qtde_vl_parcela = qtde_vl_parcela + 1
			end if
		next

	if qtde_vl_parcela = 0 then Response.Redirect("aviso.asp?id=" & ERR_PARAMETRO_OBRIGATORIO_NAO_ESPECIFICADO)
	if qtde_pedidos <> qtde_vl_parcela then
		alerta=texto_add_br(alerta)
		alerta=alerta & "A quantidade de pedidos diverge da quantidade de valores: " & qtde_pedidos & " pedidos e " & qtde_vl_parcela & " valores"
		end if

'	TRANSFERE O VALOR DA PARCELA PARA O VETOR COM DADOS COMPLETOS
	dim idx_pedido
	idx_pedido = LBound(v_pedido) - 1
	if alerta = "" then
		for i = Lbound(v_aux) to Ubound(v_aux)
			if Trim(v_aux(i))<>"" then
				idx_pedido = idx_pedido + 1
				if idx_pedido > Ubound(v_pedido) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Falha ao tentar alinhar os valores com os pedidos: o índice do valor (" & Cstr(idx_pedido) & ") ultrapassa o índice superior da lista de pedidos (" & Cstr(Ubound(v_pedido)) & ")"
					exit for
				else
					vl_aux = converte_numero(v_aux(i))
					if vl_aux = 0 then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Valor inválido para o pedido " & v_pedido(idx_pedido).pedido_marketplace
					else
						v_pedido(idx_pedido).vl_parcela = vl_aux
						end if
					end if
				end if
			next
		end if

'	VERIFICA SE HÁ Nº PEDIDOS DUPLICADOS
	if alerta = "" then
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
		end if

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

'	ANALISA OS VALORES
	dim vl_total_pedido, vl_total_pago, vl_total_devolucao
	dim vl_TotalFamiliaPrecoVenda_aux, vl_TotalFamiliaPrecoNF_aux, vl_TotalFamiliaPago_aux, vl_TotalFamiliaDevolucaoPrecoVenda_aux, vl_TotalFamiliaDevolucaoPrecoNF_aux
	dim vl_total_saldo, st_pagto, msg_erro
	dim vl_total_parcela, vl_total_saldo_pagar_restante
	dim vl_saldo_a_pagar

	vl_total_pedido = 0
	vl_total_pago = 0
	vl_total_devolucao = 0
	vl_total_saldo = 0
	vl_total_parcela = 0
	vl_total_saldo_pagar_restante = 0

	if alerta = "" then
		for i = Lbound(v_pedido) to Ubound(v_pedido)
			observacoes_aux = ""
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
						.st_pagto = st_pagto
						.vl_total_familia_preco_NF = vl_TotalFamiliaPrecoNF_aux
						.vl_total_familia_pago = vl_TotalFamiliaPago_aux
						.vl_total_familia_devolucao_preco_NF = vl_TotalFamiliaDevolucaoPrecoNF_aux
						vl_total_pedido = vl_total_pedido + vl_TotalFamiliaPrecoNF_aux
						vl_total_pago = vl_total_pago + vl_TotalFamiliaPago_aux
						vl_total_devolucao = vl_total_devolucao + vl_TotalFamiliaDevolucaoPrecoNF_aux
						if vl_TotalFamiliaPrecoNF_aux > (vl_TotalFamiliaPago_aux+vl_TotalFamiliaDevolucaoPrecoNF_aux) then vl_total_saldo = vl_total_saldo + (vl_TotalFamiliaPrecoNF_aux-vl_TotalFamiliaPago_aux-vl_TotalFamiliaDevolucaoPrecoNF_aux)
						vl_saldo_a_pagar = vl_TotalFamiliaPrecoNF_aux - vl_TotalFamiliaPago_aux - vl_TotalFamiliaDevolucaoPrecoNF_aux
						vl_total_parcela = vl_total_parcela + .vl_parcela
						vl_total_saldo_pagar_restante = vl_total_saldo_pagar_restante + (.vl_total_familia_preco_NF - .vl_total_familia_pago - .vl_total_familia_devolucao_preco_NF - .vl_parcela)
						if st_pagto = ST_PAGTO_PAGO then
							observacoes_aux = texto_add_br(observacoes_aux)
							observacoes_aux = observacoes_aux & "A família de pedidos " & retorna_num_pedido_base(.pedido_ERP) & " já consta como quitada."
							end if
						if .vl_parcela > vl_saldo_a_pagar then
							observacoes_aux = texto_add_br(observacoes_aux)
							observacoes_aux = observacoes_aux & "O valor pago irá exceder o valor do saldo a pagar."
							end if
						end if
					end if

				if observacoes_aux <> "" then
					observacoes=texto_add_br(observacoes)
					observacoes=texto_add_br(observacoes)
					s = "Pedido " & .pedido_ERP & " (" & .pedido_marketplace & "):"
					observacoes=observacoes & texto_add_br(s) & observacoes_aux
					end if
				end with
			next
		end if

'	MONTA ESTRUTURA DE DADOS PARA GRAVAÇÃO NO BD
	c_dados_pagto = ""
	if alerta = "" then
		for i = Lbound(v_pedido) to Ubound(v_pedido)
			with v_pedido(i)
				if (.pedido_ERP <> "") And (.qtde_pedido_ERP_associados = 1) then
					s = .pedido_marketplace & "=" & .pedido_ERP & "=" & formata_moeda(.vl_parcela)
					if c_dados_pagto <> "" then c_dados_pagto = c_dados_pagto & "|"
					c_dados_pagto = c_dados_pagto & s
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

<form id="fPAGTO" name="fPAGTO" method="post" action="PagtoParcialMarketplaceConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dados_pagto" id="c_dados_pagto" value="<%=c_dados_pagto%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="870" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">PAGAMENTO PARCIAL (MARKETPLACE)<span class="C">&nbsp;</span></p></td>
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
		<td class="MB"><span class="PLTc">&nbsp;</span></td>
		<td colspan="6" class="MT" NOWRAP align="center"><span class="PLTd">Valores Totalizados por Família de Pedidos</span></td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td class="MEB" style="border-right:0;" NOWRAP valign="bottom"><span class="PLTe">Nº Pedido</span><br /><span class="PLTe">Marketplace</span></td>
		<td class="MEB" style="border-right:0;" NOWRAP valign="bottom"><span class="PLTe">Nº Pedido</span><br /><span class="PLTe">ERP</span></td>
		<td class="MEB" style="border-right:0;" NOWRAP align="center" valign="bottom"><span class="PLTc">Status</span><br /><span class="PLTc">Pagto</span></td>
		<td class="MDBE" NOWRAP align="right" valign="bottom"><span class="PLTd">Valor do</span><br /><span class="PLTd">Pedido</span></td>
		<td class="MDB" NOWRAP style="border-left:0;" align="right" valign="bottom"><span class="PLTd">Valor</span><br /><span class="PLTd">Já Pago</span></td>
		<td class="MDB" NOWRAP style="border-left:0;" align="right" valign="bottom"><span class="PLTd">Valor das</span><br /><span class="PLTd">Devoluções</span></td>
		<td class="MDB" NOWRAP style="border-left:0;" align="right" valign="bottom"><span class="PLTd">Saldo</span><br /><span class="PLTd">a Pagar</span></td>
		<td class="MDB" NOWRAP style="border-left:0;" align="right" valign="bottom"><span class="PLTd">Valor da</span><br /><span class="PLTd">Parcela Atual</span></td>
		<td class="MDB" NOWRAP style="border-left:0;" align="right" valign="bottom"><span class="PLTd">Saldo a</span><br /><span class="PLTd">Pagar Restante</span></td>
	</tr>
<% for i=Lbound(v_pedido) to Ubound(v_pedido)
		with v_pedido(i)
			if .pedido_marketplace <> "" then %>
			<tr bgColor="#FFFFFF">
				<input type="hidden" name="c_pedido" id="c_pedido" readonly tabindex=-1 class="PLLe" style="width:60px;margin-left:2pt;" value="<%=.pedido_ERP%>">
				<td class="MEB"><input name="c_pedido_marketplace" id="c_pedido_marketplace" readonly tabindex=-1 class="PLLe" style="width:100px;margin-left:2pt;" 
					value="<%=.pedido_marketplace%>"></td>
				<td class="MEB C" style="width:70px;word-break:break-word;"><%=.pedido_ERP%></td>
				<%s_aux = x_status_pagto_cor(.st_pagto)%>
				<td class="MEB" NOWRAP align="center"><span class="PLLc" style="color:<%=s_aux%>;"><%=Ucase(x_status_pagto(.st_pagto))%></span></td>
				<td class="MDBE" NOWRAP align="right"><input name="c_vl_pedido" id="c_vl_pedido" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;" 
					value="<%=formata_moeda(.vl_total_familia_preco_NF)%>"></td>
				<td class="MDB" NOWRAP align="right"><input name="c_vl_pago" id="c_vl_pago" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;<% if .vl_total_familia_pago < 0 then Response.Write "color:red;"%>" 
					value="<%=formata_moeda(.vl_total_familia_pago)%>"></td>
				<td class="MDB" NOWRAP align="right"><input name="c_vl_devolucao" id="c_vl_devolucao" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;" 
					value="<%=formata_moeda(.vl_total_familia_devolucao_preco_NF)%>"></td>
				<td class="MDB" NOWRAP align="right"><input name="c_vl_saldo" id="c_vl_saldo" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;<%if .vl_total_familia_preco_NF < (.vl_total_familia_pago+.vl_total_familia_devolucao_preco_NF) then Response.Write "color:red;"%>" 
					value="<%=formata_moeda(.vl_total_familia_preco_NF - .vl_total_familia_pago - .vl_total_familia_devolucao_preco_NF)%>"></td>
				<td class="MDB" NOWRAP align="right"><input name="c_vl_parcela" id="c_vl_parcela" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;<% if .vl_parcela < 0 then Response.Write "color:red;"%>"
					value="<%=formata_moeda(.vl_parcela)%>"></td>
				<%	vl_aux = .vl_total_familia_preco_NF - .vl_total_familia_pago - .vl_total_familia_devolucao_preco_NF - .vl_parcela
					if .pedido_ERP <> "" then
						s_vl_aux = formata_moeda(vl_aux)
					else
						s_vl_aux = ""
						end if
				%>
				<td class="MDB" NOWRAP align="right"><input name="c_vl_saldo_restante" id="c_vl_saldo_restante" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;<% if vl_aux < 0 then Response.Write "color:red;"%>" 
					value="<%=s_vl_aux%>"></td>
			</tr>
<%				end if
			end with
		next		%>
	<tr bgColor="#FFFFFF">
		<td><span class="PLTe">&nbsp;</span></td>
		<td><span class="PLTe">&nbsp;</span></td>
		<td><span class="PLTc">&nbsp;</span></td>
		<td class="MDBE" NOWRAP align="right"><input name="c_vl_total_pedido" id="c_vl_total_pedido" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;color:green;" 
			value="<%=formata_moeda(vl_total_pedido)%>"></td>
		<td class="MDB" NOWRAP align="right"><input name="c_vl_total_pago" id="c_vl_total_pago" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;color:<%if vl_total_pago >= 0 then Response.Write "green" else Response.Write "red"%>;" 
			value="<%=formata_moeda(vl_total_pago)%>"></td>
		<td class="MDB" NOWRAP align="right"><input name="c_vl_total_devolucao" id="c_vl_total_devolucao" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;color:green;" 
			value="<%=formata_moeda(vl_total_devolucao)%>"></td>
		<td class="MDB" NOWRAP align="right"><input name="c_vl_total_saldo" id="c_vl_total_saldo" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;color:<%if vl_total_saldo >= 0 then Response.Write "green" else Response.Write "red"%>;" 
			value="<%=formata_moeda(vl_total_saldo)%>"></td>
		<td class="MDB" NOWRAP align="right"><input name="c_vl_total_parcela" id="c_vl_total_parcela" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;color:<%if vl_total_parcela >= 0 then Response.Write "green" else Response.Write "red"%>;" 
			value="<%=formata_moeda(vl_total_parcela)%>"></td>
		<td class="MDB" NOWRAP align="right"><input name="c_vl_total_saldo_pagar_restante" id="c_vl_total_saldo_pagar_restante" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;color:<%if vl_total_saldo_pagar_restante >= 0 then Response.Write "green" else Response.Write "red"%>;" 
			value="<%=formata_moeda(vl_total_saldo_pagar_restante)%>"></td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="870" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<% if erro_consistencia then %>
<table class="notPrint" width="870" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
<% else %>
<table class="notPrint" width="870" cellSpacing="0">
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