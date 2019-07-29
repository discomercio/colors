<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ==============================================================================
'	  C R E D I T O O K A G U A R D A N D O D E P O S I T O C O N S I S T E . A S P
'     ==============================================================================
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
				s = "SELECT pedido, analise_credito FROM t_PEDIDO WHERE (pedido='" & v_pedido(i) & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & " n�o est� cadastrado."
				else
					if Cstr(rs("analise_credito")) <> Cstr(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido(i) & " possui status de an�lise de cr�dito inv�lido para esta opera��o: " & UCase(x_analise_credito(Trim("" & rs("analise_credito"))))
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
function fCrOkConfirma( f ) {
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
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;FONT-WEIGHT:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<BR><BR>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ***************************************************************** -->
<!-- **********  P�GINA PARA EXIBIR DADOS PARA CONFIRMA��O  ********** -->
<!-- ***************************************************************** -->
<body onload="focus();">
<center>

<form id="fCrOk" name="fCrOk" METHOD="POST" ACTION="CreditoOkAguardandoDepositoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type=HIDDEN name="pedidos_selecionados" id="pedidos_selecionados" value="<%=lista_pedidos%>">

<!--  I D E N T I F I C A � � O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="RIGHT" vAlign="BOTTOM"><p class="PEDIDO">Cr�dito OK (aguardando dep�sito)<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!-- ************   H� OBSERVA��ES?  ************ -->
<% if observacoes <> "" then %>
		<span class="Lbl">OBSERVA��ES</span>
		<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><P style='margin:5px 2px 5px 2px;'><%=observacoes%></p></div>
		<br><br>
<% end if %>

<!--  PEDIDOS  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
		<td class="MT" style='background:azure;' NOWRAP><span class="PLTe">N� Pedido(s)&nbsp;</span></td>
	</tr>
<% for i=Lbound(v_pedido) to Ubound(v_pedido) 
		if v_pedido(i) <> "" then %>
			<tr bgColor="#FFFFFF">
				<td class="MDBE" NOWRAP><input name="c_pedido" id="c_pedido" readonly tabindex=-1 class="PLLe" style="width:60px;margin-left:2pt;" 
					value="<%=v_pedido(i)%>"
					onkeypress="if (digitou_enter(true)) this.focus();"></td>
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
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para p�gina anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="RIGHT"><div name="dCONFIRMA" id="dCONFIRMA">
	<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fCrOkConfirma(fCrOk)" title="confirma o dep�sito">
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