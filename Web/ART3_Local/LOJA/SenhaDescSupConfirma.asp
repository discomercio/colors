<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  S E N H A D E S C S U P C O N F I R M A . A S P
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

	dim s, s_log, usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim cliente_selecionado
	cliente_selecionado=Trim(request("cliente_selecionado"))
	if cliente_selecionado = "" then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)

	dim alerta
	alerta=""

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

'	OBTÉM DADOS DO FORMULÁRIO
	dim c_loja, id_desconto, c_supervisor_autorizador
	c_supervisor_autorizador = Trim(Request.Form("c_supervisor_autorizador"))
	c_loja = retorna_so_digitos(Trim(Request.Form("c_loja")))
	s = normaliza_codigo(c_loja, TAM_MIN_LOJA)
	if s <> "" then c_loja = s

	dim perc_max_desconto_cadastrado_na_loja
	perc_max_desconto_cadastrado_na_loja = obtem_perc_max_desconto_cadastrado_na_loja(c_loja)

	dim r_loja
	set r_loja = New cl_LOJA
	if Not x_loja_bd(c_loja, r_loja) then Response.Redirect("aviso.asp?id=" & ERR_LOJA_NAO_CADASTRADA)
	
	dim intCounter, intQtdeItens
	dim v_item
	redim v_item(0)
	set v_item(0) = New cl_ITEM_SENHA_DESCONTO
	intQtdeItens = Request.Form("c_produto").Count
	for intCounter = 1 to intQtdeItens
		s = Trim(Request.Form("c_produto")(intCounter))
		if s <> "" then
			if Trim(v_item(Ubound(v_item)).produto) <> "" then
				redim preserve v_item(Ubound(v_item)+1)
				set v_item(Ubound(v_item)) = New cl_ITEM_SENHA_DESCONTO
				end if
			with v_item(Ubound(v_item))
				.produto = UCase(Trim(Request.Form("c_produto")(intCounter)))
				s = retorna_so_digitos(Request.Form("c_fabricante")(intCounter))
				.fabricante = normaliza_codigo(s, TAM_MIN_FABRICANTE)
				s = Trim(Request.Form("c_desc_max_senha")(intCounter))
				.perc_desconto = converte_numero(s)
				end with
			end if
		next

	if c_loja = "" then
		alerta = "Não foi especificada a loja."
	elseif c_supervisor_autorizador = "" then
		alerta = "Não foi informado quem está autorizando o desconto."
		end if
		
	dim blnTemItem
	for intCounter = Lbound(v_item) to Ubound(v_item)
		with v_item(intCounter)
			blnTemItem = False
			if Trim("" & .fabricante) <> "" then blnTemItem = True
			if Trim("" & .produto) <> "" then blnTemItem = True
			if .perc_desconto > 0 then blnTemItem = True
		
			if blnTemItem then
				if Trim(.fabricante) = "" then
					alerta = "Não foi especificado o código do fabricante."
				elseif Trim(.produto) = "" then
					alerta = "Não foi especificado o código do produto."
				elseif (.perc_desconto <= 0) Or (.perc_desconto > 100) then
					alerta = "Percentual de desconto inválido."
				elseif (.perc_desconto > perc_max_desconto_cadastrado_na_loja) then
					alerta = "Percentual de desconto de " & formata_perc_desc(.perc_desconto) & "% para o produto '" & .produto & "' excede o máximo permitido para cadastramento na loja."
					end if
				end if
			end with
		next

	if alerta = "" then
		dim r_cliente
		set r_cliente = New cl_CLIENTE
		if Not x_cliente_bd(cliente_selecionado, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_CADASTRADO)

	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		If Not cria_recordset_pessimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	'	INFORMAÇÕES PARA O LOG
		s_log = ""
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)

				blnTemItem = False
				if Trim("" & .fabricante) <> "" then blnTemItem = True
				if Trim("" & .produto) <> "" then blnTemItem = True
				if .perc_desconto > 0 then blnTemItem = True
		
				if blnTemItem then
					if Not gera_nsu(NSU_DESC_SUP_AUTORIZACAO, id_desconto, msg_erro) then 
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
						end if

					s = "SELECT * FROM t_DESCONTO WHERE (id = 'XXX')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if Err <> 0 then
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						end if

					rs.AddNew
					rs("id") = id_desconto
					rs("data") = Now
					rs("autorizador") = usuario
					rs("supervisor_autorizador") = c_supervisor_autorizador
					rs("id_cliente") = cliente_selecionado
					rs("cnpj_cpf") = r_cliente.cnpj_cpf
					rs("fabricante") = .fabricante
					rs("produto") = .produto
					rs("desc_max") = .perc_desconto
					rs("loja") = c_loja
					rs.Update
					if Err <> 0 then
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						end if
					
				'	INFORMAÇÕES PARA O LOG
					if s_log <> "" then s_log = s_log & "; "
					s_log = s_log & _
							formata_perc_desc(.perc_desconto) & _
							"% p/ produto " & .produto & " (" & .fabricante & ")"
					end if  'if (blnTemItem)
				end with
			next

	'	INFORMAÇÕES PARA O LOG
		s_log = "Senha de desconto superior cadastrada para o cliente " & cnpj_cpf_formata(r_cliente.cnpj_cpf) & _
				" na loja " & c_loja & " autorizada por " & c_supervisor_autorizador & _
				": " & s_log
		grava_log usuario, c_loja, "", cliente_selecionado, OP_LOG_DESC_SUP_AUTORIZACAO_NA_LOJA, s_log
		
	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
		if Err=0 then 
			s = "Senha para autorização de desconto superior cadastrada com sucesso"
			Session(SESSION_CLIPBOARD) = s
			Response.Redirect("mensagem.asp?url_back=" & server.URLEncode("SenhaDescSupPesqCliente.asp") & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		else
			alerta=Cstr(Err) & ": " & Err.Description
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



<html>


<head>
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>



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
<% end if %>

</html>


<%
	on error resume next
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>