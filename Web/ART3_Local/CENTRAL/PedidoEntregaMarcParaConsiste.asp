<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  P E D I D O E N T R E G A M A R C P A R A C O N S I S T E . A S P
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

	dim s, usuario, s_dt_entrega, lista_pedidos, v_pedido, v_aux, i, j, achou, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	s_dt_entrega = Trim(request("c_dt_entrega"))
	
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

	dim c_nfe_emitente
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))

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
		if s_dt_entrega = "" then
			alerta = "Informe a data para entrega"
		elseif Not IsDate(s_dt_entrega) then
			alerta = "Data de coleta é inválida"
		else
			if StrToDate(s_dt_entrega) < Date then alerta = "Data de coleta não pode ser uma data passada"
			end if
		end if
	
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
	
	if alerta = "" then
		for i = Lbound(v_pedido) to Ubound(v_pedido)
			if v_pedido(i) <> "" then
				s = "SELECT pedido, numero_loja, st_entrega, id_nfe_emitente FROM t_PEDIDO WHERE (pedido='" & v_pedido(i) & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & " não está cadastrado."
				else
					if Not IsEntregaAgendavel(Trim("" & rs("st_entrega"))) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido(i) & " possui status inválido para esta operação: " & Ucase(x_status_entrega(Trim("" & rs("st_entrega"))))
						end if

				'	VERIFICA SE O CD DO USUÁRIO ESTÁ COERENTE COM O PEDIDO
					if CLng(rNfeEmitente.id) <> CLng(rs("id_nfe_emitente")) then
					'	ERRO: PEDIDO PERTENCE A OUTRO CD
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido(i) & " pertence a outro CD"
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



<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fDTETGConfirma( f ) {
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

<form id="fDTETG" name="fDTETG" method="post" action="PedidoEntregaMarcParaConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedidos_selecionados" id="pedidos_selecionados" value="<%=lista_pedidos%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Data de Coleta<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!-- ************   HÁ OBSERVAÇÕES?  ************ -->
<% if observacoes <> "" then %>
		<span class="Lbl">OBSERVAÇÕES</span>
		<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><P style='margin:5px 2px 5px 2px;'><%=observacoes%></p></div>
		<br><br>
<% end if %>

<!--  PEDIDOS  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
		<td class="MT" NOWRAP style='background:azure;'><span class="PLTe">Data de Coleta&nbsp;</span></td>
	</tr>
			<tr bgColor="#FFFFFF">
				<td class="MDBE" NOWRAP><input name="c_dt_entrega" id="c_dt_entrega" readonly tabindex=-1 class="PLLe" style="width:60px;margin-left:2pt;" 
					value="<%=s_dt_entrega%>"></td>
					</tr>
	<tr bgColor="#FFFFFF"><td>&nbsp;</td></tr>
	<tr bgColor="#FFFFFF">
		<td class="MT" NOWRAP style='background:azure;'><span class="PLTe">Nº Pedido(s)&nbsp;</span></td>
	</tr>
<% for i=Lbound(v_pedido) to Ubound(v_pedido) 
		if v_pedido(i) <> "" then %>
			<tr bgColor="#FFFFFF">
				<td class="MDBE" NOWRAP><input name="c_pedido" id="c_pedido" readonly tabindex=-1 class="PLLe" style="width:60px;margin-left:2pt;" 
					value="<%=v_pedido(i)%>"></td>
			</tr>
<%			end if
		next		%>
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
	<td align="RIGHT"><div name="dCONFIRMA" id="dCONFIRMA">
	<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fDTETGConfirma(fDTETG)" title="confirma o agendamento da data de coleta">
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