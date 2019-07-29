<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  PedidoPreencheIndicadorConsiste.asp
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

	dim s, usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
	
	dim s_pedido
	dim c_indicador, lista_pedidos, v_pedido, v_aux, i, j, achou, msg_erro
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	c_indicador = Trim(Request.Form("c_indicador"))
	
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
		if c_indicador = "" then
			alerta = "Selecione um indicador"
			end if
		end if
	
	if alerta = "" then
		for i = Lbound(v_pedido) to Ubound(v_pedido)
			if v_pedido(i) <> "" then
				if IsPedidoFilhote(v_pedido(i)) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & " é um número de pedido-filhote!! Informe apenas números de pedidos-base!!"
					end if
				end if
			next
		end if
	
	dim v_indicador
	redim v_indicador(Ubound(v_pedido))
	
	dim id_pedido_base
	if alerta = "" then
		for i = Lbound(v_pedido) to Ubound(v_pedido)
			if v_pedido(i) <> "" then
				id_pedido_base = retorna_num_pedido_base(v_pedido(i))
				s = "SELECT" & _
						" pedido," & _
						" indicador," & _
						" loja," & _
						" indicador_editado_manual_status" & _
					" FROM t_PEDIDO" & _
					" WHERE" & _
						" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')" & _
					" ORDER BY" & _
						" pedido," & _
						" data"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & " não está cadastrado."
				else
					do while Not rs.Eof
						v_indicador(i) = Trim("" & rs("indicador"))
						
						if (Trim("" & rs("indicador")) <> "") And (rs("indicador_editado_manual_status") = 0) And (Ucase(Trim("" & rs("indicador"))) <> "NILTON SP") then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & Trim("" & rs("pedido")) & " já está preenchido com um indicador (" & Trim("" & rs("indicador")) & ")"
							end if
						if converte_numero(rs("loja")) <> converte_numero(loja) then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & Trim("" & rs("pedido")) & " não pertence a esta loja (loja do pedido: " & Trim("" & rs("loja")) & ")"
							end if
						
						rs.MoveNext
						loop
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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fOPConfirma(f) {
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

<form id="fOP" name="fOP" method="post" action="PedidoPreencheIndicadorConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">
<input type="hidden" name="c_pedidos" id="c_pedidos" value="<%=lista_pedidos%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Preencher Indicador em Pedido Cadastrado<span class="C">&nbsp;</span></p></td>
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
		<td class="MT" NOWRAP style='background:azure;'><span class="PLTe">Indicador&nbsp;</span></td>
	</tr>
			<tr bgColor="#FFFFFF">
				<td class="MDBE" NOWRAP>
				<p class="C"><%=c_indicador & " - " & Ucase(x_orcamentista_e_indicador(c_indicador))%></p>
				</td>
			</tr>
	<tr bgColor="#FFFFFF"><td>&nbsp;</td></tr>
	<tr bgColor="#FFFFFF">
		<td class="MT" NOWRAP style='background:azure;'><span class="PLTe">Nº Pedido(s)&nbsp;</span></td>
	</tr>
<% for i=Lbound(v_pedido) to Ubound(v_pedido)
		if v_pedido(i) <> "" then
			s_pedido = Trim(v_pedido(i))
			if Trim("" & v_indicador(i)) <> "" then s_pedido = s_pedido & " (" & v_indicador(i) & ")"
%>
			<tr bgColor="#FFFFFF">
				<td class="MDBE" NOWRAP>
				<p class="C"><%=s_pedido%></p>
				</td>
			</tr>
<%			end if
		next%>
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
	<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fOPConfirma(fOP)" title="confirma a gravação dos dados">
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