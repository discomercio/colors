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
'	  PedidoEditaIndicadorConfirma.asp
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

	if Not operacao_permitida(OP_CEN_EDITA_PEDIDO_INDICADOR, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim pedido_selecionado, s_pedidos
	pedido_selecionado = Trim(Request.Form("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim c_indicador_novo
	c_indicador_novo = Trim(Request.Form("c_indicador_novo"))

	if c_indicador_novo = "" then
		alerta = "Não foi informado o novo indicador do pedido."
		end if

	dim r_pedido
	if alerta = "" then
		if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then
			alerta = msg_erro
			end if
		
		if r_pedido.st_entrega = ST_ENTREGA_CANCELADO then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O status do pedido é inválido para realizar esta operação (" & x_status_entrega(r_pedido.st_entrega) & ")"
			end if
		
		s = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido='" & c_indicador_novo & "')"
		set rs = cn.Execute(s)
		if rs.Eof then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Não foi localizado o cadastro do novo indicador selecionado ('" & c_indicador_novo & "')!"
			end if
		if rs.State <> 0 then rs.Close

		if UCase(r_pedido.indicador) = UCase(c_indicador_novo) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O novo indicador selecionado para o pedido é igual ao indicador já cadastrado (" & r_pedido.indicador & ")!"
			end if

		s = "SELECT * FROM t_COMISSAO_INDICADOR_N4 WHERE (pedido='" & r_pedido.pedido & "')"
		set rs = cn.Execute(s)
		if Not rs.Eof then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O indicador não pode ser alterado porque este pedido já foi processado no relatório de comissões!"
			end if
		if rs.State <> 0 then rs.Close
		end if

	if alerta = "" then
		s = "SELECT *" & _
			" FROM t_COMISSAO_INDICADOR_NFSe_N1 n1" & _
				" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 n2 ON (n1.id = n2.id_comissao_indicador_nfse_n1)" & _
				" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO n3Ped ON (n2.id = n3Ped.id_comissao_indicador_nfse_n2)" & _
			" WHERE" & _
				" (n1.status <> 0)" & _
				" AND (n3Ped.pedido = '" & r_pedido.pedido & "')"
		set rs = cn.Execute(s)
		if Not rs.Eof then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O indicador não pode ser alterado porque este pedido já foi processado no relatório de comissões (via NFSe)!"
			end if
		if rs.State <> 0 then rs.Close
		end if 'if alerta = "" then

	if alerta = "" then
		'O indicador é válido p/ toda a família de pedidos, portanto, verifica se há algum pedido da família que já teve a comissão processada
		s = "SELECT" & _
				" pedido" & _
			" FROM t_PEDIDO" & _
			" WHERE" & _
				" (pedido_base = '" & retorna_num_pedido_base(r_pedido.pedido) & "')" & _
				" AND (comissao_paga = " & COD_COMISSAO_PAGA & ")" & _
			" ORDER BY" & _
				" pedido"
		set rs = cn.Execute(s)
		if Not rs.Eof then
			s_pedidos = ""
			do while Not rs.Eof
				if s_pedidos <> "" then s_pedidos = s_pedidos & ", "
				s_pedidos = s_pedidos & Trim("" & rs("pedido"))
				rs.MoveNext
				loop
			alerta=texto_add_br(alerta)
			if s_pedidos = r_pedido.pedido then
				alerta=alerta & "O indicador não pode ser alterado porque este pedido já consta com status de comissão paga!"
			else
				alerta=alerta & "O indicador não pode ser alterado porque há pedidos desta família de pedidos que constam com status de comissão paga (" & s_pedidos & ")!"
				end if
			end if
		if rs.State <> 0 then rs.Close
		end if 'if alerta = "" then

	if alerta = "" then
		'O indicador é válido p/ toda a família de pedidos, portanto, verifica se há alguma devolução da família que já teve a comissão descontada
		s = "SELECT DISTINCT" & _
				" pedido" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
			" WHERE" & _
				" (pedido LIKE '" & retorna_num_pedido_base(r_pedido.pedido) & BD_CURINGA_TODOS & "')" & _
				" AND (comissao_descontada = " & COD_COMISSAO_DESCONTADA & ")" & _
			" ORDER BY" & _
				" pedido"
		set rs = cn.Execute(s)
		if Not rs.Eof then
			s_pedidos = ""
			do while Not rs.Eof
				if s_pedidos <> "" then s_pedidos = s_pedidos & ", "
				s_pedidos = s_pedidos & Trim("" & rs("pedido"))
				rs.MoveNext
				loop
			alerta=texto_add_br(alerta)
			if s_pedidos = r_pedido.pedido then
				alerta=alerta & "O indicador não pode ser alterado porque este pedido possui devolução(ões) descontada(s) da comissão!"
			else
				alerta=alerta & "O indicador não pode ser alterado porque esta família de pedidos possui devolução(ões) descontada(s) da comissão (" & s_pedidos & ")!"
				end if
			end if
		if rs.State <> 0 then rs.Close
		end if 'if alerta = "" then

	if alerta = "" then
		'O indicador é válido p/ toda a família de pedidos, portanto, verifica se há alguma perda da família que já teve a comissão descontada
		s = "SELECT DISTINCT" & _
				" pedido" & _
			" FROM t_PEDIDO_PERDA" & _
			" WHERE" & _
				" (pedido LIKE '" & retorna_num_pedido_base(r_pedido.pedido) & BD_CURINGA_TODOS & "')" & _
				" AND (comissao_descontada = " & COD_COMISSAO_DESCONTADA & ")" & _
			" ORDER BY" & _
				" pedido"
		set rs = cn.Execute(s)
		if Not rs.Eof then
			s_pedidos = ""
			do while Not rs.Eof
				if s_pedidos <> "" then s_pedidos = s_pedidos & ", "
				s_pedidos = s_pedidos & Trim("" & rs("pedido"))
				rs.MoveNext
				loop
			alerta=texto_add_br(alerta)
			if s_pedidos = r_pedido.pedido then
				alerta=alerta & "O indicador não pode ser alterado porque este pedido possui perda(s) descontada(s) da comissão!"
			else
				alerta=alerta & "O indicador não pode ser alterado porque esta família de pedidos possui perda(s) descontada(s) da comissão (" & s_pedidos & ")!"
				end if
			end if
		if rs.State <> 0 then rs.Close
		end if 'if alerta = "" then

	dim blnPedidoPassouPossuirIndicador, blnTratarRALiq, perc_desagio_RA, perc_desagio_RA_liquida, perc_limite_RA_sem_desagio
	blnPedidoPassouPossuirIndicador = False
	blnTratarRALiq = False
	perc_desagio_RA = 0
	perc_desagio_RA_liquida = 0
	perc_limite_RA_sem_desagio = 0
	if alerta = "" then
		'Se está sendo cadastrado agora um indicador em um pedido que não tinha indicador, realiza o tratamento p/ os campos de RA e RA líquida
		if (r_pedido.indicador = "") And (c_indicador_novo <> "") then
			blnPedidoPassouPossuirIndicador = True
			perc_desagio_RA = obtem_perc_desagio_RA_do_indicador(c_indicador_novo)
			perc_limite_RA_sem_desagio = obtem_perc_limite_RA_sem_desagio()

			if (Cstr(r_pedido.loja) <> Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE)) then
				blnTratarRALiq = True
				perc_desagio_RA_liquida = getParametroPercDesagioRALiquida
				end if
			end if
		end if 'if alerta = "" then


	dim sBlocoNotasMsg
	sBlocoNotasMsg = ""

	dim rEmailDestinatario
	dim id_email, corpo_mensagem, msg_erro_grava_email

	dim indicador_anterior
	dim s_log
	s_log = ""

'	GRAVA A ALTERAÇÃO DO INDICADOR NO PEDIDO
	if alerta = "" then
	'	INICIA A TRANSAÇÃO
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO then
		'	BLOQUEIA REGISTRO PARA EVITAR ACESSO CONCORRENTE (REALIZA O FLIP EM UM CAMPO BIT APENAS P/ ADQUIRIR O LOCK EXCLUSIVO)
		'	OBS: TODOS OS MÓDULOS DO SISTEMA QUE REALIZEM ESTA OPERAÇÃO DE CADASTRAMENTO DEVEM SINCRONIZAR O ACESSO OBTENDO O LOCK EXCLUSIVO DO REGISTRO DE CONTROLE DESIGNADO
			s = "UPDATE t_CONTROLE SET" & _
					" dummy = ~dummy" & _
				" WHERE" & _
					" id_nsu = '" & ID_XLOCK_SYNC_PEDIDO & "'"
			cn.Execute(s)
			end if

		if rs.State <> 0 then rs.Close
		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if
		
		if alerta = "" then
			s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & retorna_num_pedido_base(pedido_selecionado) & "')"
			rs.Open s, cn
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Falha ao tentar localizar o registro do pedido " & retorna_num_pedido_base(pedido_selecionado)
				end if
			end if

		if alerta = "" then
			indicador_anterior = rs("indicador")
			rs("indicador")=c_indicador_novo
			'Assegura que não haverá inconsistência c/ relação a vendedor do parceiro (se futuramente houver edição do vendedor do parceiro, ajustar este ponto)
			rs("IdIndicadorVendedor") = Null

			if blnPedidoPassouPossuirIndicador then
				rs("perc_desagio_RA") = perc_desagio_RA
				rs("perc_limite_RA_sem_desagio") = perc_limite_RA_sem_desagio
				
				if blnTratarRALiq then
					rs("perc_desagio_RA_liquida") = perc_desagio_RA_liquida
					end if
				end if

			if Trim("" & rs("analise_credito")) = COD_AN_CREDITO_OK then
				'Envia mensagem de alerta sobre alteração do indicador em pedido com status "crédito ok"
				set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaAlteracaoIndicadorEmPedidoCreditoOk)
				if Trim("" & rEmailDestinatario.campo_texto) <> "" then
					corpo_mensagem = "O usuário '" & usuario & "' alterou em " & formata_data_hora_sem_seg(Now) & " na Central o indicador do pedido " & pedido_selecionado & vbCrLf & _
										vbCrLf & _
										"Indicador anterior:" & vbCrLf & _
										Ucase(Trim(indicador_anterior)) & vbCrLf & _
										vbCrLf & _
										"Indicador atual:" & vbCrLf & _
										Ucase(Trim(c_indicador_novo))

					EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
													"", _
													rEmailDestinatario.campo_texto, _
													"", _
													"", _
													"Alteração do indicador em pedido com status 'Crédito OK' (pedido " & pedido_selecionado & ")", _
													corpo_mensagem, _
													Now, _
													id_email, _
													msg_erro_grava_email
					end if
				end if

			'Registra edição no bloco de notas
			sBlocoNotasMsg = "Edição do indicador realizada por '" & usuario & "' (status da análise de crédito: " & descricao_analise_credito(Trim("" & rs("analise_credito"))) & ")" & vbCrLf & _
							"Anterior: " & indicador_anterior & vbCrLf & _
							"Novo: " & c_indicador_novo
			if Not grava_bloco_notas_pedido(pedido_selecionado, ID_USUARIO_SISTEMA, "", COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_INDICADOR, msg_erro) then
				alerta = "Falha ao gravar bloco de notas com mensagem automática no pedido (" & pedido_selecionado & ")"
				end if
			'Assegura de gravar também no pedido-base pois trata-se de informação controlada através do pedido-base
			if IsPedidoFilhote(pedido_selecionado) then
				if Not grava_bloco_notas_pedido(retorna_num_pedido_base(pedido_selecionado), ID_USUARIO_SISTEMA, "", COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_EDICAO_INDICADOR, msg_erro) then
					alerta = "Falha ao gravar bloco de notas com mensagem automática no pedido (" & retorna_num_pedido_base(pedido_selecionado) & ")"
					end if
				end if

			rs.Update
			if Err <> 0 then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if
			
			if rs.State <> 0 then rs.Close
			
			'Ajusta o indicador de todos os pedidos da família
			if alerta = "" then
				s = "UPDATE t_PED__FILHOTE" & _
					" SET" & _
						" t_PED__FILHOTE.indicador = t_PED__BASE.indicador"
				
				if blnPedidoPassouPossuirIndicador then
					s = s & _
						", t_PED__FILHOTE.perc_desagio_RA = t_PED__BASE.perc_desagio_RA" & _
						", t_PED__FILHOTE.perc_limite_RA_sem_desagio = t_PED__BASE.perc_limite_RA_sem_desagio"
					
					if blnTratarRALiq then
						s = s & _
							", t_PED__FILHOTE.perc_desagio_RA_liquida = t_PED__BASE.perc_desagio_RA_liquida"
						end if
					end if
				
				s = s & _
					" FROM t_PEDIDO AS t_PED__FILHOTE" & _
						" INNER JOIN t_PEDIDO AS t_PED__BASE ON (t_PED__FILHOTE.pedido_base = t_PED__BASE.pedido)" & _
					" WHERE" & _
						" (t_PED__FILHOTE.pedido_base = '" & retorna_num_pedido_base(pedido_selecionado) & "')" & _
						" AND (t_PED__FILHOTE.pedido <> t_PED__FILHOTE.pedido_base)"
				cn.Execute(s)
				If Err <> 0 then
					alerta = "FALHA AO SINCRONIZAR O CAMPO 'indicador' (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if

			s_log = "Alteração do indicador do pedido " & pedido_selecionado & " de '" & indicador_anterior & "' para '" & c_indicador_novo & "'"
			if s_log <> "" then grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_ALTERACAO, s_log
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