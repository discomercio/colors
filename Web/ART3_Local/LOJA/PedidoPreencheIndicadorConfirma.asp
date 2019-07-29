<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  PedidoPreencheIndicadorConfirma.asp
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

	dim msg_erro, s_log, s_log_pedido_filhote, s_log_indicador, c_indicador, lista_pedidos, v_pedido, i, achou

'	OBTÉM DADOS DO FORMULÁRIO
	c_indicador = Trim(request("c_indicador"))
	lista_pedidos = ucase(Trim(request("c_pedidos")))
	if (lista_pedidos = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
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

	dim id_pedido_base
	dim dt_hr_gravacao
	dt_hr_gravacao = Now
	
	dim alerta
	alerta=""

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

		for i=Lbound(v_pedido) to Ubound(v_pedido)
			if v_pedido(i) <> "" then
				s_log_pedido_filhote = ""
				s_log_indicador = ""
				id_pedido_base = retorna_num_pedido_base(v_pedido(i))
				s = "SELECT * FROM t_PEDIDO WHERE (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "') ORDER BY pedido, data"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & " não foi encontrado."
				else
					do while Not rs.Eof
					'	INFORMAÇÕES PARA O LOG
						if IsPedidoFilhote(Trim("" & rs("pedido"))) then
							if s_log_pedido_filhote <> "" then s_log_pedido_filhote = s_log_pedido_filhote & ", "
							s_log_pedido_filhote = s_log_pedido_filhote & Trim("" & rs("pedido"))
							end if
						
						s_log_indicador = Trim("" & rs("indicador"))
						
						if (Trim("" & rs("indicador")) <> "") And (rs("indicador_editado_manual_status") = 0) And (Ucase(Trim("" & rs("indicador"))) <> "NILTON SP") then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & Trim("" & rs("pedido")) & " já está preenchido com um indicador: " & Trim("" & rs("indicador"))
						elseif converte_numero(rs("loja")) <> converte_numero(loja) then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & Trim("" & rs("pedido")) & " não pertence a esta loja (loja do pedido: " & Trim("" & rs("loja")) & ")"
						else
							if (Trim("" & rs("indicador")) <> "") And (rs("indicador_editado_manual_status") = 0) then rs("indicador_editado_manual_indicador_original") = Trim("" & rs("indicador"))
							rs("indicador")=c_indicador
							rs("indicador_editado_manual_status")=1
							rs("indicador_editado_manual_data_hora")=dt_hr_gravacao
							rs("indicador_editado_manual_usuario")=usuario
							rs.Update
							if Err <> 0 then 
								alerta=texto_add_br(alerta)
								alerta=alerta & Cstr(Err) & ": " & Err.Description
								end if
							end if
						rs.MoveNext
						loop
						
					if rs.State <> 0 then rs.Close
				
				'	INFORMAÇÕES PARA O LOG
					if s_log <> "" then s_log = s_log & ", "
					s_log = s_log & v_pedido(i)
					if s_log_pedido_filhote <> "" then
						s_log = s_log & " (" & s_log_pedido_filhote & ")"
						end if
					if s_log_indicador <> "" then
						s_log = s_log & " [" & s_log_indicador & "]"
						end if
					end if
				end if
			
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next

		if alerta = "" then
			s_log = "Preencher indicador em pedido cadastrado: Indicador = " & c_indicador & "; Pedido(s) = " & s_log
			grava_log usuario, loja, "", "", OP_LOG_PEDIDO_ALTERACAO, s_log
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
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
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
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>