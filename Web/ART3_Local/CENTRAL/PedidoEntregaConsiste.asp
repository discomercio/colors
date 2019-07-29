<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================================
'	  P E D I D O E N T R E G A C O N S I S T E . A S P
'     =====================================================
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

	dim s, usuario, lista_pedidos, v_pedido, v_aux, i, j, achou, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	lista_pedidos = ucase(Trim(request("c_pedidos")))
	if (lista_pedidos = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
	dim c_nfe_emitente
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
	
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
	v_pedido(Ubound(v_pedido))=""
	for i = Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			if Trim(v_pedido(Ubound(v_pedido)))<>"" then
				redim preserve v_pedido(Ubound(v_pedido)+1)
				end if
			v_pedido(Ubound(v_pedido)) = Trim(v_aux(i))
			end if
		next
	
	lista_pedidos = join(v_pedido,chr(13))
	
	dim alerta
	alerta = ""

	dim observacoes
	observacoes = ""
	
	if alerta = "" then
		if c_nfe_emitente = "" then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Não foi informado o CD"
		elseif converte_numero(c_nfe_emitente) = 0 then
			alerta=texto_add_br(alerta)
			alerta = alerta & "É necessário definir um CD válido"
			end if
		end if

	dim rNfeEmitente
	set rNfeEmitente = le_nfe_emitente(c_nfe_emitente)
	
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		if v_pedido(i)<>"" then
			for j=Lbound(v_pedido) to (i-1)
				if v_pedido(i) = v_pedido(j) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & ": linha " & renumera_com_base1(Lbound(v_pedido),i) & " repete o mesmo pedido da linha " & renumera_com_base1(Lbound(v_pedido),j) & "."
					exit for
					end if
				next
			end if
		next

	if alerta = "" then
		for i = Lbound(v_pedido) to Ubound(v_pedido)
			if v_pedido(i) <> "" then
				s = "SELECT" & _
						" t_PEDIDO.pedido," & _
						" t_PEDIDO.numero_loja," & _
						" t_PEDIDO.st_entrega," & _
						" t_PEDIDO.st_etg_imediata," & _
						" t_PEDIDO.id_nfe_emitente," & _
						" t_PEDIDO__BASE.analise_credito" & _
					" FROM t_PEDIDO" & _
						" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
					" WHERE" & _
						" (t_PEDIDO.pedido='" & v_pedido(i) & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & " não está cadastrado."
				else
					if Trim("" & rs("st_entrega")) <> ST_ENTREGA_A_ENTREGAR then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido(i) & " possui status inválido para esta operação: " & Ucase(x_status_entrega(Trim("" & rs("st_entrega"))))
						end if
					if CInt(rs("analise_credito")) <> CInt(COD_AN_CREDITO_OK) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido(i) & " possui análise de crédito em situação inválida: " & Ucase(descricao_analise_credito(rs("analise_credito")))
						end if
					if CInt(rs("st_etg_imediata")) <> CInt(COD_ETG_IMEDIATA_SIM) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido(i) & " está com 'Entrega Imediata' assinalada com 'Não'"
						end if
					
				'	VERIFICA SE O CD DO USUÁRIO ESTÁ COERENTE COM O PEDIDO
					if CLng(rNfeEmitente.id) <> CLng(rs("id_nfe_emitente")) then
					'	ERRO: PEDIDO PERTENCE A OUTRO CD
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & Trim("" & rs("pedido")) & " pertence a outro CD"
						end if
					end if
				end if
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fETGConfirma( f ) {
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
<div class="MtAlerta" style="width:649px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
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

<form id="fETG" name="fETG" method="post" action="pedidoentregaconfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedidos_selecionados" id="pedidos_selecionados" value="<%=lista_pedidos%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Entrega de Mercadorias<span class="C">&nbsp;</span></span></td>
</tr>
</table>
<br>


<!-- ************   HÁ OBSERVAÇÕES?  ************ -->
<% if observacoes <> "" then %>
		<span class="Lbl">OBSERVAÇÕES</span>
		<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><span style='margin:5px 2px 5px 2px;'><%=observacoes%></span></div>
		<br><br>
<% end if %>

<!--  PEDIDOS  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
		<td class="MT" align="left" nowrap style='background:azure;'><span class="PLTe">Nº Pedido(s)&nbsp;</span></td>
	</tr>
<% for i=Lbound(v_pedido) to Ubound(v_pedido) 
		if v_pedido(i) <> "" then %>
			<tr bgColor="#FFFFFF">
				<td class="MDBE" align="left" nowrap><input name="c_pedido" id="c_pedido" readonly tabindex=-1 class="PLLe" style="width:60px;margin-left:2pt;" 
					value="<%=v_pedido(i)%>"></td>
			</tr>
<%			end if
		next		%>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
	<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fETGConfirma(fETG)" title="confirma a entrega de mercadorias">
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