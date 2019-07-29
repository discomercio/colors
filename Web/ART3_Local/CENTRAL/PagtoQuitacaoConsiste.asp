<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ====================================================
'	  P A G T O Q U I T A C A O C O N S I S T E . A S P
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


	On Error GoTo 0
	Err.Clear

	dim s, usuario, lista_pedidos, v_pedido, v_aux, i, j, achou
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	lista_pedidos = ucase(Trim(request("c_pedidos")))
	if (lista_pedidos = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
	lista_pedidos=substitui_caracteres(lista_pedidos,chr(10),"")
	v_aux = split(lista_pedidos,chr(13),-1)
	achou=False
	for i=Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			achou = True
			s = normaliza_num_pedido(v_aux(i))
			if s <> "" then v_aux(i) = s
			end if
		next

	if Not achou then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	redim v_pedido(0)
	set v_pedido(Ubound(v_pedido))=New cl_QUATRO_COLUNAS
	with v_pedido(Ubound(v_pedido))
		.c1=""
		.c2=0
		.c3=0
		.c4=0
		end with
	
	for i = Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			if Trim(v_pedido(Ubound(v_pedido)).c1)<>"" then
				redim preserve v_pedido(Ubound(v_pedido)+1)
				set v_pedido(Ubound(v_pedido))=New cl_QUATRO_COLUNAS
				end if
			with v_pedido(Ubound(v_pedido))
				.c1 = Trim(v_aux(i))
				.c2 = 0
				.c3 = 0
				.c4 = 0
				end with
			end if
		next
	
	redim v_aux(Ubound(v_pedido))
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		v_aux(i)=v_pedido(i).c1
		next

	lista_pedidos = join(v_aux,chr(13))
	
	dim alerta
	alerta = ""

	dim observacoes
	observacoes = ""
	
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		if v_pedido(i).c1<>"" then
			for j=Lbound(v_pedido) to (i-1)
				if (retorna_num_pedido_base(v_pedido(i).c1)<>"") And (retorna_num_pedido_base(v_pedido(i).c1) = retorna_num_pedido_base(v_pedido(j).c1)) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedidos " & v_pedido(i).c1 & " e " & v_pedido(j).c1 & " pertencem à mesma família de pedidos."
					exit for
					end if
				next
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
				if .c1 <> "" then
					s = "SELECT pedido FROM t_PEDIDO WHERE (pedido='" & .c1 & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & .c1 & " não está cadastrado."
					else
					'	OBTÉM OS VALORES A PAGAR, JÁ PAGO E O STATUS DE PAGAMENTO (PARA TODA A FAMÍLIA DE PEDIDOS)
						if Not calcula_pagamentos(.c1, vl_TotalFamiliaPrecoVenda_aux, vl_TotalFamiliaPrecoNF_aux, vl_TotalFamiliaPago_aux, vl_TotalFamiliaDevolucaoPrecoVenda_aux, vl_TotalFamiliaDevolucaoPrecoNF_aux, st_pagto, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						.c2 = vl_TotalFamiliaPrecoNF_aux
						.c3 = vl_TotalFamiliaPago_aux
						.c4 = vl_TotalFamiliaDevolucaoPrecoNF_aux
						vl_total_pedido = vl_total_pedido + vl_TotalFamiliaPrecoNF_aux
						vl_total_pago = vl_total_pago + vl_TotalFamiliaPago_aux
						vl_total_devolucao = vl_total_devolucao + vl_TotalFamiliaDevolucaoPrecoNF_aux
						if vl_TotalFamiliaPrecoNF_aux > (vl_TotalFamiliaPago_aux+vl_TotalFamiliaDevolucaoPrecoNF_aux) then vl_total_saldo = vl_total_saldo + (vl_TotalFamiliaPrecoNF_aux-vl_TotalFamiliaPago_aux-vl_TotalFamiliaDevolucaoPrecoNF_aux)
						if st_pagto = ST_PAGTO_PAGO then
							observacoes = texto_add_br(observacoes)
							observacoes = observacoes & "A família de pedidos " & retorna_num_pedido_base(.c1) & " já consta como quitada."
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

<form id="fPAGTO" name="fPAGTO" method="post" action="PagtoQuitacaoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedidos_selecionados" id="pedidos_selecionados" value="<%=lista_pedidos%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Quitação de Pedidos<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!-- ************   HÁ OBSERVAÇÕES?  ************ -->
<% if observacoes <> "" then %>
		<span class="Lbl">OBSERVAÇÕES</span>
		<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'><%=observacoes%></p></div>
		<br><br>
<% end if %>


<!--  PEDIDOS  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
		<td class="MB" NOWRAP><span class="PLTe">&nbsp;</span></td>
		<td colspan="4" class="MT" NOWRAP align="center"><span class="PLTd">Valores Totalizados por Família de Pedidos</span></td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td class="MEB" style="border-right:0;" NOWRAP><span class="PLTe">Nº Pedido</span></td>
		<td class="MDBE" NOWRAP align="right"><span class="PLTd">Valor do Pedido</span></td>
		<td class="MDB" NOWRAP style="border-left:0;" align="right"><span class="PLTd">Valor Já Pago</span></td>
		<td class="MDB" NOWRAP style="border-left:0;" align="right"><span class="PLTd">&nbsp;Valor das Devoluções</span></td>
		<td class="MDB" NOWRAP style="border-left:0;" align="right"><span class="PLTd">Saldo a Pagar</span></td>
	</tr>
<% for i=Lbound(v_pedido) to Ubound(v_pedido) 
		with v_pedido(i)
			if .c1 <> "" then	%>
			<tr bgColor="#FFFFFF">
				<td class="MEB" NOWRAP><input name="c_pedido" id="c_pedido" readonly tabindex=-1 class="PLLe" style="width:60px;margin-left:2pt;" 
					value="<%=.c1%>"></td>
				<td class="MDBE" NOWRAP align="right"><input name="c_vl_pedido" id="c_vl_pedido" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;" 
					value="<%=formata_moeda(.c2)%>"></td>
				<td class="MDB" NOWRAP align="right"><input name="c_vl_pago" id="c_vl_pago" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;<% if .c3 < 0 then Response.Write "color:red;"%>" 
					value="<%=formata_moeda(.c3)%>"></td>
				<td class="MDB" NOWRAP align="right"><input name="c_vl_devolucao" id="c_vl_devolucao" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;" 
					value="<%=formata_moeda(.c4)%>"></td>
				<td class="MDB" NOWRAP align="right"><input name="c_vl_saldo" id="c_vl_saldo" readonly tabindex=-1 class="PLLd" style="width:100px;margin-right:2pt;<%if .c2 < (.c3+.c4) then Response.Write "color:red;"%>" 
					value="<%=formata_moeda(.c2 - .c3 - .c4)%>"></td>
			</tr>
<%				end if
			end with
		next		%>
	<tr bgColor="#FFFFFF">
		<td NOWRAP><span class="PLTe">&nbsp;</span></td>
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
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


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