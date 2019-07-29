<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  P010PreRequisitos.asp
'     ===============================================
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


	On Error GoTo 0
	Err.Clear

	dim s, usuario, loja, pedido_selecionado, id_pedido_base
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
	id_pedido_base = retorna_num_pedido_base(pedido_selecionado)
	pedido_selecionado = id_pedido_base

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim alerta
	alerta = ""
	
	dim r_pedido, v_item
	if Not le_pedido(id_pedido_base, r_pedido, msg_erro) then
		alerta = msg_erro
	else
		if Trim(r_pedido.loja) <> loja then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_INVALIDO)
		if Not le_pedido_item_consolidado_familia(id_pedido_base, v_item, msg_erro) then alerta = msg_erro
		end if
	
	dim r_cliente
	set r_cliente = New cl_CLIENTE
	if alerta = "" then
		if Not x_cliente_bd(r_pedido.id_cliente, r_cliente) then
			if alerta <> "" then alerta = alerta & "<BR>"
			alerta = alerta & "Falha ao tentar obter do banco de dados os dados cadastrais do cliente"
			end if
		end if
	
	if alerta = "" then
		if CLng(r_pedido.st_end_entrega) <> 0 then
			s = substitui_caracteres(Ucase(r_pedido.EndEtg_endereco), "&", " E ")
			if (r_pedido.EndEtg_endereco <> "") And (r_pedido.EndEtg_endereco_numero <> "") then
				s = s & ", " & Ucase(r_pedido.EndEtg_endereco_numero)
				if Len(s) > 60 then
					if alerta <> "" then alerta = alerta & "<BR>"
					alerta = alerta & "O endereço de entrega excede o tamanho máximo de 60 caracteres!"
					end if
				end if

			if Trim("" & r_pedido.EndEtg_bairro) = "" then
				if alerta <> "" then alerta = alerta & "<BR>"
				alerta = alerta & "É necessário preencher o campo 'Bairro' nos dados do endereço de entrega do pedido!"
				end if
			end if
		end if
	
	if alerta = "" then
		if CLng(r_pedido.st_end_entrega) = 0 then
			if Trim("" & r_cliente.bairro) = "" then
				if alerta <> "" then alerta = alerta & "<BR>"
				alerta = alerta & "É necessário preencher o campo 'Bairro' no endereço do cadastro do cliente!"
				end if
			end if
		end if

	dim intNumParcelasFormaPagto
	intNumParcelasFormaPagto = 0
	if Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_A_VISTA) then
		if Cstr(r_pedido.av_forma_pagto) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = 1
	elseif Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_PARCELA_UNICA) then
		if Cstr(r_pedido.pu_forma_pagto) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = 1
	elseif Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_PARCELADO_CARTAO) then
		intNumParcelasFormaPagto = r_pedido.pc_qtde_parcelas
	elseif Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) then
		'NOP
	elseif Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) then
	'	ENTRADA + PRESTAÇÕES
		if Cstr(r_pedido.pce_forma_pagto_entrada) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = intNumParcelasFormaPagto + 1
		if Cstr(r_pedido.pce_forma_pagto_prestacao) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = intNumParcelasFormaPagto + r_pedido.pce_prestacao_qtde
	elseif Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) then
	'	1ª PRESTAÇÃO + DEMAIS PRESTAÇÕES
		if Cstr(r_pedido.pse_forma_pagto_prim_prest) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = intNumParcelasFormaPagto + 1
		if Cstr(r_pedido.pse_forma_pagto_demais_prest) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = intNumParcelasFormaPagto + r_pedido.pse_demais_prest_qtde
		end if
	
	if intNumParcelasFormaPagto = 0 then
		if alerta <> "" then alerta = alerta & "<BR>"
		alerta = alerta & "A forma de pagamento do pedido não especifica nenhuma parcela no cartão!"
		end if
	
	dim blnForcarAlterarEmail
	blnForcarAlterarEmail = False
	dim s_cnpj_cpf, s_email, s_erro_email
	if alerta = "" then
		s_cnpj_cpf = r_cliente.cnpj_cpf
		s_email = Trim(r_cliente.email)
		if s_email <> "" then
			if Not email_AF_ok(s_email, s_cnpj_cpf, s_erro_email) then
				blnForcarAlterarEmail = True
				end if
			end if
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
	<title>LOJA</title>
	</head>


<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fPAGTOConclui(f) {

	if (f.c_email_atual.value.toString().length == 0) {
		if (f.c_email_novo.value.toString().length == 0) {
			alert("É necessário cadastrar um endereço de e-mail!");
			return;
		}
	}

	if (f.c_email_novo.value != f.c_email_novo_redigite.value) {
		alert("O endereço de e-mail redigitado não confere!");
		return;
	}

	if (f.c_email_novo.value.toString().length > 0) {
		if (!email_ok(f.c_email_novo.value)) {
			alert("O novo endereço de e-mail possui formato inválido!");
			return;
		}
	}

	if (f.blnForcarAlterarEmail.value == "1") {
		if ((trim(f.c_email_novo.value) == "") || (f.c_email_atual.value.toUpperCase() == f.c_email_novo.value.toUpperCase())) {
			alert("É obrigatório informar um novo endereço de email!");
			f.c_email_novo.focus();
			return;
		}
	}
	
	f.action = "P020PreRequisitosConfirma.asp";
	window.status = "Aguarde ...";
	dCONFIRMA.style.visibility = 'hidden';
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
.C1{
	font-family: Arial, Helvetica, sans-serif;
	color: #000000;
	font-size: 10pt;
	font-weight: bold;
	font-style: normal;
	margin: 0pt 2pt 1pt 2pt;
	}
.Cd1{
	font-family: Arial, Helvetica, sans-serif;
	color: #000000;
	font-size: 10pt;
	font-weight: bold;
	font-style: normal;
	margin: 0pt 2pt 1pt 2pt;
	text-align: right;
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
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ****************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESUMO DO PEDIDO  ***************** -->
<!-- ****************************************************************** -->
<body onload="window.status='';fPAGTO.c_email_novo.focus();">
<center>

<form id="fPAGTO" name="fPAGTO" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" value="<%=pedido_selecionado%>">
<% if blnForcarAlterarEmail then %>
<input type="hidden" name="blnForcarAlterarEmail" id="blnForcarAlterarEmail" value="1">
<% else %>
<input type="hidden" name="blnForcarAlterarEmail" id="blnForcarAlterarEmail" value="0">
<% end if %>
<table cellspacing="0" width="649" style="border-bottom:1px solid black">
<tr>
	<td align="center" valign="bottom"><img src="../imagem/<%=BRASPAG_LOGOTIPO_LOJA%>"></td>
</tr>
</table>

<br>
<br>
<table class="Qx" cellspacing="0" border="0">
	<tr><td colspan="2" align="center">
	<% if blnForcarAlterarEmail then %>
	<span class="STP" style="font-size:14pt;">É necessário atualizar o endereço de e-mail</span><br /><span class="Cd1">Motivo: <%=s_erro_email%></span>
	<% else %>
	<span class="STP" style="font-size:14pt;">Confira se é necessário atualizar o endereço de e-mail</span>
	<% end if %>
	</td></tr>
	<tr><td colspan="2" style="height:6px;" align="left"></td></tr>
	<tr>
	<td align="right"><span class="Cd1">E-mail cadastrado</span></td>
	<td align="left"><input name="c_email_atual" id="c_email_atual" class="C1" style="width:250px;color:#696969;" value="<%=r_cliente.email%>" readonly tabindex=-1></td>
	</tr>
	<tr><td colspan="2" style="height:30px;" align="left"></td></tr>
	<tr><td colspan="2" align="center"><span class="STP" style="font-size:14pt;">Atualizar o endereço de e-mail</span></td></tr>
	<tr><td colspan="2" style="height:6px;" align="left"></td></tr>
	<tr>
	<td align="right"><span class="Cd1">Novo e-mail</span></td>
	<td align="left"><input name="c_email_novo" id="c_email_novo" class="C1" style="width:250px;" maxlength="60" value="" onkeypress="if (digitou_enter(true)) {fPAGTO.c_email_novo_redigite.focus();}" onblur="this.value=trim(this.value);"></td>
	</tr>
	<tr>
	<td align="right"><span class="Cd1">Redigite o novo e-mail</span></td>
	<td align="left"><input name="c_email_novo_redigite" id="c_email_novo_redigite" class="C1" style="width:250px;" maxlength="60" value="" onkeypress="if (digitou_enter(true)) {bCONFIRMA.focus();}" onblur="this.value=trim(this.value);"></td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>


<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" href="../Loja/pedido.asp?pedido_selecionado=<%=pedido_selecionado%>&url_back=X<%="&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right"><div name='dCONFIRMA' id='dCONFIRMA'><a name="bCONFIRMA" href="javascript:fPAGTOConclui(fPAGTO)" title="Continua com o processo de pagamento">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>

</form>

</center>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>