<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  PedidoRecebidoConfirma.asp
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

	dim s, usuario, msg_erro, s_log, s_log_aux, lista_pedidos, v_pedido, i, achou
	dim ckb_entrega, ckb_recebido, c_dt_recebido
	
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	OBTÉM DADOS DO FORMULÁRIO
	lista_pedidos = ucase(Trim(request("pedidos_selecionados")))
	if (lista_pedidos = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	ckb_entrega = Trim(Request.Form("ckb_entrega"))
	ckb_recebido = Trim(Request.Form("ckb_recebido"))
	c_dt_recebido = Trim(Request.Form("c_dt_recebido"))
	if ckb_recebido = "" then 
		c_dt_recebido = ""
	else
		if c_dt_recebido = "" then c_dt_recebido = formata_data(Date)
		end if
	
	lista_pedidos=substitui_caracteres(lista_pedidos,chr(10),"")
	v_pedido = split(lista_pedidos,chr(13),-1)
	achou=False
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		if Trim(v_pedido(i))<>"" then
			achou = True
			s = normaliza_num_pedido(v_pedido(i))
			if s <> "" then v_pedido(i) = s
			end if
		next

	if Not achou then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim alerta
	alerta=""

	if alerta = "" then
		 if (ckb_recebido <> "") And (c_dt_recebido <> "") then
			if StrToDate(c_dt_recebido) > Date then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Data de recebimento informada é inválida (" & c_dt_recebido & ")"
				end if
			end if
		end if
	
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	s_log = ""
	
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		If Not cria_recordset_pessimista(rs, msg_erro) then 
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

		if (alerta = "") And (ckb_entrega <> "") then
			for i=Lbound(v_pedido) to Ubound(v_pedido)
				if v_pedido(i) <> "" then
					s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & v_pedido(i) & "')"
					if rs.State <> 0 then rs.Close
					rs.Open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido(i) & " não foi encontrado."
					else
						if Trim("" & rs("st_entrega"))<>ST_ENTREGA_A_ENTREGAR then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & v_pedido(i) & " possui status inválido para a operação 'Entrega': " & Ucase(x_status_entrega(Trim("" & rs("st_entrega"))))
						else
							rs("st_entrega") = ST_ENTREGA_ENTREGUE
							rs("entregue_data") = Date
							rs("entregue_usuario") = usuario
							rs.Update
							if Err <> 0 then 
								alerta=texto_add_br(alerta)
								alerta=alerta & Cstr(Err) & ": " & Err.Description
								end if
							if rs.State <> 0 then rs.Close
						'	RETIRA AS MERCADORIAS DO ESTOQUE DE PRODUTOS 'VENDIDOS'
							if Not estoque_processa_entrega_mercadoria(usuario, v_pedido(i), msg_erro) then
							'	~~~~~~~~~~~~~~~~
								cn.RollbackTrans
							'	~~~~~~~~~~~~~~~~
								Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
								end if
							end if
						end if
					end if
				
			'	SE HOUVE ERRO, CANCELA O LAÇO
				if alerta <> "" then exit for
				next
			end if

		if (alerta = "") And (ckb_recebido <> "") then
			for i=Lbound(v_pedido) to Ubound(v_pedido)
				if v_pedido(i) <> "" then
					s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & v_pedido(i) & "')"
					if rs.State <> 0 then rs.Close
					rs.Open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido(i) & " não foi encontrado."
					else
						if Trim("" & rs("st_entrega"))<>ST_ENTREGA_ENTREGUE then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & v_pedido(i) & " possui status inválido para a operação 'Recebido': " & Ucase(x_status_entrega(Trim("" & rs("st_entrega"))))
						else
							rs("PedidoRecebidoStatus") = CLng(COD_ST_PEDIDO_RECEBIDO_SIM)
							rs("PedidoRecebidoData") = StrToDate(c_dt_recebido)
							rs("PedidoRecebidoDtHrUltAtualiz") = Now
							rs("PedidoRecebidoUsuarioUltAtualiz") = usuario
							if (Trim("" & rs("marketplace_codigo_origem")) <> "") And (CLng(rs("MarketplacePedidoRecebidoRegistrarStatus")) = CLng(0)) then
								rs("MarketplacePedidoRecebidoRegistrarStatus") = 1
								rs("MarketplacePedidoRecebidoRegistrarDataRecebido") = StrToDate(c_dt_recebido)
								rs("MarketplacePedidoRecebidoRegistrarDataHora") = Now
								rs("MarketplacePedidoRecebidoRegistrarUsuario") = usuario
								end if
							rs.Update
							if Err <> 0 then 
								alerta=texto_add_br(alerta)
								alerta=alerta & Cstr(Err) & ": " & Err.Description
								end if
							if rs.State <> 0 then rs.Close
							end if
						end if
					end if
				
			'	SE HOUVE ERRO, CANCELA O LAÇO
				if alerta <> "" then exit for
				next
			end if

		if alerta = "" then
			if s_log <> "" then s_log = s_log & "; "
			s_log = s_log & "Operação="
			s_log_aux = ""
			if ckb_entrega <> "" then s_log_aux = "Entrega"
			if ckb_recebido <> "" then
				if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
				s_log_aux = s_log_aux & "Recebido (" & c_dt_recebido & ")"
				end if
			s_log = s_log & s_log_aux
			
			if s_log <> "" then s_log = s_log & "; "
			s_log = s_log & "Pedidos="
			s_log_aux = ""
			for i=Lbound(v_pedido) to Ubound(v_pedido)
				if v_pedido(i) <> "" then
				'	INFORMAÇÕES PARA O LOG
					if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
					s_log_aux = s_log_aux & v_pedido(i)
					end if
				next
			s_log = s_log & s_log_aux
			
			grava_log usuario, "", "", "", OP_LOG_PEDIDO_RECEBIDO, s_log
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
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
	<title>CENTRAL</title>
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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>