<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     =============================================
'	  P020PreRequisitosConfirma.asp
'     =============================================
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

	usuario = BRASPAG_USUARIO_CLIENTE

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	id_pedido_base = retorna_num_pedido_base(pedido_selecionado)

	dim cnpj_cpf_selecionado
	cnpj_cpf_selecionado = retorna_so_digitos(Request("cnpj_cpf_selecionado"))


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, t_CLIENTE, t_PEDIDO, msg_erro, msg_erro_aux
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(t_CLIENTE, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(t_PEDIDO, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_log
	s_log = ""
	
	dim alerta
	alerta = ""

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim c_email_novo, c_email_novo_redigite
	c_email_novo = LCase(Trim(Request("c_email_novo")))
	c_email_novo_redigite = LCase(Trim(Request("c_email_novo_redigite")))
	
	dim r_pedido, v_item
	if Not le_pedido(id_pedido_base, r_pedido, msg_erro) then
		alerta = msg_erro
	else
		loja = r_pedido.loja
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
		if c_email_novo <> "" then
			if c_email_novo <> c_email_novo_redigite then
				if alerta <> "" then alerta = alerta & "<BR>"
				alerta = alerta & "O endereço de e-mail redigitado não confere!"
				end if
			end if
		end if
	
	if alerta = "" then
		if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
			if (r_pedido.endereco_email = "") And (c_email_novo = "") then
				if alerta <> "" then alerta = alerta & "<BR>"
				alerta = alerta & "É necessário cadastrar um endereço de e-mail!"
				end if
		else
			if (r_cliente.email = "") And (c_email_novo = "") then
				if alerta <> "" then alerta = alerta & "<BR>"
				alerta = alerta & "É necessário cadastrar um endereço de e-mail!"
				end if
			end if
		end if
	
	dim s_cnpj_cpf, s_email
	if alerta = "" then
		if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
			s_cnpj_cpf = r_pedido.endereco_cnpj_cpf
		else
			s_cnpj_cpf = r_cliente.cnpj_cpf
			end if

		if c_email_novo <> "" then
			s_email = c_email_novo
			if Not email_AF_ok(s_email, s_cnpj_cpf, msg_erro_aux) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Endereço de email (" & s_email & ") não é válido!!<br />" & msg_erro_aux
				end if
		else
			if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
				s_email = r_pedido.endereco_email
			else
				s_email = r_cliente.email
				end if
			if Not email_AF_ok(s_email, s_cnpj_cpf, msg_erro_aux) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Endereço de email (" & s_email & ") não é válido!!<br />" & msg_erro_aux
				end if
			end if
		end if
	
	if alerta = "" then
		if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
			if (c_email_novo <> "") And (LCase(c_email_novo) <> LCase(r_pedido.endereco_email)) then
				s = "SELECT " & _
						"*" & _
					" FROM t_PEDIDO" & _
					" WHERE" & _
						"(pedido_base = '" & id_pedido_base & "')"
				if t_PEDIDO.State <> 0 then t_PEDIDO.Close
				t_PEDIDO.open s, cn
				do while Not t_PEDIDO.Eof
					if s_log <> "" then s_log = s_log & "; "
					s_log = s_log & "Pedido " & Trim("" & t_PEDIDO("pedido")) & ": " & formata_texto_log(Trim("" & t_PEDIDO("endereco_email"))) & " => " & formata_texto_log(c_email_novo)
					t_PEDIDO("endereco_email") = c_email_novo
					t_PEDIDO.Update
				
					t_PEDIDO.MoveNext
					loop
				
				if s_log <> "" then
					s_log = "Alteração do e-mail (pedido=" & id_pedido_base & "; CNPJ/CPF: " & cnpj_cpf_formata(r_pedido.endereco_cnpj_cpf) & "): " & s_log
					grava_log usuario, loja, id_pedido_base, "", OP_LOG_PEDIDO_ALTERACAO, s_log
					end if
				end if
		else
			if (c_email_novo <> "") And (LCase(c_email_novo) <> LCase(r_cliente.email)) then
				s = "SELECT " & _
						"*" & _
					" FROM t_CLIENTE" & _
					" WHERE" & _
						"(id = '" & r_pedido.id_cliente & "')"
				if t_CLIENTE.State <> 0 then t_CLIENTE.Close
				t_CLIENTE.open s, cn
				if Not t_CLIENTE.Eof then
					s_log = formata_texto_log(Trim("" & t_CLIENTE("email"))) & " => " & formata_texto_log(c_email_novo)
					t_CLIENTE("email_anterior") = Trim("" & t_CLIENTE("email"))
					t_CLIENTE("email") = c_email_novo
					t_CLIENTE("email_atualizacao_data") = Date
					t_CLIENTE("email_atualizacao_data_hora") = Now
					t_CLIENTE("email_atualizacao_usuario") = usuario
					t_CLIENTE.Update
				
					s_log = "Alteração do e-mail (id=" & r_cliente.id & "; CNPJ/CPF: " & cnpj_cpf_formata(r_cliente.cnpj_cpf) & "): " & s_log
					grava_log usuario, loja, "", r_pedido.id_cliente, OP_LOG_CLIENTE_ALTERACAO, s_log
					end if
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
	<title><%=SITE_CLIENTE_TITULO_JANELA%></title>
	</head>


<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__SSL_JS%>" language="JavaScript" type="text/javascript"></script>

<% if alerta = "" then %>
<script language="JavaScript" type="text/javascript">
	setTimeout('fBraspag.submit()', 100);
</script>
<% end if %>


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
<link href="<%=URL_FILE__E_LOGO_TOP_BS_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
body::before
{
	content: '';
	border: none;
	margin-top: 0px;
	margin-bottom: 0px;
	padding: 0px;
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
<body>
<center>

<form id="fBraspag" name="fBraspag" method="post" action="P030PagtoVerificaStatus.asp">
<input type="hidden" name='pedido_selecionado' value='<%=pedido_selecionado%>'>
<input type="hidden" name='cnpj_cpf_selecionado' value='<%=cnpj_cpf_selecionado%>'>
</form>

<b>Aguarde, redirecionando para exibir mensagem...</b>
<br />
<a name="bREDIRECIONA" id="bREDIRECIONA" href="javascript:fBraspag.submit();"><b>Se o redirecionamento não ocorrer automaticamente, clique aqui.</b></a>

</center>

<% if SITE_CLIENTE_EXIBIR_LOGO_SSL then %>
<script language="JavaScript" type="text/javascript">
	logo_ssl_corner("../imagem/ssl/ssl_corner.gif");
</script>
<% end if %>

</body>
<% end if %>

</html>

<%
	if t_CLIENTE.State <> 0 then t_CLIENTE.Close
	set t_CLIENTE = nothing

	if t_PEDIDO.State <> 0 then t_PEDIDO.Close
	set t_PEDIDO = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>