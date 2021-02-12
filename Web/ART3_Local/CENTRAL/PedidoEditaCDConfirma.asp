<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/Global.asp" -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================================
'	  PedidoEditaCDConfirma.asp
'     ===============================================================
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

	dim s, msg_erro
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim alerta
	alerta=""
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    dim url_origem
    url_origem = Trim(Request("url_origem"))

	if Not operacao_permitida(OP_CEN_EDITA_PEDIDO_CD, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim pedido_selecionado
	pedido_selecionado = Trim(Request.Form("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim c_novo_CD
	c_novo_CD = Trim(Request.Form("c_novo_CD"))

	if c_novo_CD = "" then
		alerta = "Não foi informado o novo CD do pedido."
	elseif converte_numero(c_novo_CD) <= 0 then
		alerta = "A identificação do novo CD do pedido é inválida."
		end if

	dim r_pedido
	if alerta = "" then
		if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then
			alerta = msg_erro
			end if
		
		if r_pedido.st_entrega <> ST_ENTREGA_ESPERAR then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O status do pedido é inválido para realizar esta operação (" & x_status_entrega(r_pedido.st_entrega) & ")"
			end if
		
		if CLng(r_pedido.id_nfe_emitente) = CLng(c_novo_CD) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O novo CD selecionado para o pedido é igual ao CD já cadastrado (" & obtem_apelido_empresa_NFe_emitente(r_pedido.id_nfe_emitente) & ")!"
			end if
		end if


	dim id_nfe_emitente_anterior
	dim s_log
	s_log = ""

'	GRAVA A ALTERAÇÃO DO CD NO PEDIDO
	if alerta = "" then
	'	INICIA A TRANSAÇÃO
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if rs.State <> 0 then rs.Close
		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if
		
		if alerta = "" then
			s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & pedido_selecionado & "')"
			rs.Open s, cn
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Falha ao tentar localizar o registro do pedido " & pedido_selecionado
				end if
			end if

		if alerta = "" then
			id_nfe_emitente_anterior = rs("id_nfe_emitente")
			rs("id_nfe_emitente")=CLng(c_novo_CD)
			rs.Update
			if Err <> 0 then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if
			
			if rs.State <> 0 then rs.Close
			
			s_log = "Alteração do CD do pedido " & pedido_selecionado & " de " & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_anterior) & " (" & CStr(id_nfe_emitente_anterior) & ") para " & obtem_apelido_empresa_NFe_emitente(c_novo_CD) & " (" & c_novo_CD & ")"
			if s_log <> "" then grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_ALTERACAO, s_log
			
		'	PROCESSA OS PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE
			if Not estoque_processa_produtos_vendidos_sem_presenca_v2(CLng(c_novo_CD), usuario, msg_erro) then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
				end if
			end if 'if alerta = ""

		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then
				Response.Redirect("pedido.asp?pedido_selecionado=" & pedido_selecionado & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))) & "&url_origem=" & url_origem
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