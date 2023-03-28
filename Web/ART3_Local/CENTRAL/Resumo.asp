<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<!-- #include file = "../global/Global.asp"        -->

<% if Trim(Session("usuario_a_checar")) = "" then %>
	<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->
<% end if %>

<%
    
'     ===================
'	  R E S U M O . A S P
'     ===================
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




' _____________________________________________________________________________________________
'
'						I N I C I A L I Z A     P Á G I N A     A S P
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
'	EXIBIÇÃO DE BOTÕES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True
	
	
'	VERIFICA ID
	dim s, idx, usuario, usuario_nome, senha, senha_real, blnPossuiPermissaoAcesso, chave
	dim dt_ult_alteracao_senha, usuario_bloqueado, usuario_bloqueado_automatico, confere_login_no_bd, eh_primeira_execucao, strFlagPrimeiraExecucao
	dim idUsuario, qtdeConsecutivaFalhaLogin, max_tentativas_login, blnUsuarioCadastrado, blnSenhaConfereOk, dtHrBloqueioAutomatico
	dim id_email, assunto_mensagem, corpo_mensagem, remetente_mensagem, msg_erro_grava_email, rEmailDestinatario, ambiente_execucao
	
	confere_login_no_bd = (Trim(Session("usuario_a_checar")) <> "")
	usuario = Trim(Session("usuario_a_checar")): Session("usuario_a_checar") = " "
	senha = Trim(Session("senha_a_checar")): Session("senha_a_checar") = " "
	
	
'	OBTEM O ID
	if usuario = "" then usuario = Session("usuario_atual")
	if senha = "" then senha = Session("senha_atual")
	usuario_nome = Session("usuario_nome_atual")
	
	if (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
	if (senha = "") then Response.Redirect("aviso.asp?id=" & ERR_SENHA_NAO_INFORMADA)
	
	if isHorarioManutencaoSistema then Response.Redirect("aviso.asp?id=" & ERR_HORARIO_MANUTENCAO_SISTEMA)
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim nivel_acesso_bloco_notas, nivel_acesso_chamado
	dim s_lista_operacoes_permitidas, s_separacao, qtde_rel_glb, qtde_rel_com, qtde_rel_adm, qtde_rel_compras_logist, qtde_total_rel, intIdx
	dim strSessionCtrlTicket
	dim strMensagemAviso, strMensagemAvisoPopUp
	strMensagemAviso = ""
	strMensagemAvisoPopUp = ""
	
	strFlagPrimeiraExecucao = Request("FlagPrimeiraExecucao")
	if strFlagPrimeiraExecucao = "1" then eh_primeira_execucao = True
	
'	VERIFICA USUARIO E SENHA NO BD
	if confere_login_no_bd then
		eh_primeira_execucao = true
		blnPossuiPermissaoAcesso = false
		dt_ult_alteracao_senha = null
		usuario_bloqueado=false
		usuario_bloqueado_automatico=False
		blnUsuarioCadastrado=False
		blnSenhaConfereOk=False
		s = "SELECT" & _
				" Id" & _
				", nome" & _
				", senha" & _
				", datastamp" & _
				", dt_ult_alteracao_senha" & _
				", bloqueado" & _
				", SessionCtrlTicket" & _
				", SessionCtrlLoja" & _
				", SessionCtrlModulo" & _
				", SessionCtrlDtHrLogon" & _
				", StLoginBloqueadoAutomatico" & _
				", QtdeConsecutivaFalhaLogin" & _
			" FROM t_USUARIO" & _
			" WHERE" & _
				" (usuario='" & QuotedStr(usuario) & "')"
		set rs = cn.Execute(s)
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
		if Not rs.eof then
			blnUsuarioCadastrado=True
			if Trim("" & rs("SessionCtrlTicket")) <> "" then
				strMensagemAviso = "A sessão anterior não foi encerrada corretamente.<br>Para segurança da sua identidade, <i>sempre</i> encerre a sessão clicando no link <i>'encerra'</i>.<br>Esta ocorrência será gravada no histórico de auditoria."
				strMensagemAvisoPopUp = "**********   A T E N Ç Ã O ! !   **********\nA sessão anterior não foi encerrada corretamente.\nPara segurança da sua identidade, SEMPRE encerre a sessão clicando no link ENCERRA.\nEsta ocorrência será gravada no histórico de auditoria!!"
				s = "INSERT INTO t_SESSAO_ABANDONADA (" & _
						"usuario," & _
						"SessaoAbandonadaDtHrInicio," & _
						"SessaoAbandonadaLoja," & _
						"SessaoAbandonadaModulo," & _
						"SessaoSeguinteDtHrInicio," & _
						"SessaoSeguinteLoja," & _
						"SessaoSeguinteModulo" & _
					") VALUES (" & _
						"'" & QuotedStr(usuario) & "'," & _
						bd_formata_data_hora(rs("SessionCtrlDtHrLogon")) & "," & _
						"'" & Trim("" & rs("SessionCtrlLoja")) & "'," & _
						"'" & Trim("" & rs("SessionCtrlModulo")) & "'," & _
						bd_formata_data_hora(Session("DataHoraLogon")) & "," & _
						"''," & _
						"'" & SESSION_CTRL_MODULO_CENTRAL & "'" & _
					")"
				cn.Execute(s)
				end if
			
		'	TEM SENHA?
			if Trim("" & rs("datastamp")) = "" then usuario_bloqueado=true
		'	ACESSO BLOQUEADO?
			if rs("bloqueado")<>0 then usuario_bloqueado=true
		'	ACESSO BLOQUEADO AUTOMATICAMENTE POR EXCESSO DE TENTATIVAS C/ SENHA ERRADA?
			if rs("StLoginBloqueadoAutomatico")<>0 then usuario_bloqueado_automatico=true
			qtdeConsecutivaFalhaLogin = rs("QtdeConsecutivaFalhaLogin")
			max_tentativas_login = obtem_parametro_max_tentativas_login

			idUsuario = rs("Id")
			dt_ult_alteracao_senha = rs("dt_ult_alteracao_senha")
			usuario_nome = Trim("" & rs("nome"))
			
		'	VERIFICA SE POSSUI ALGUM ACESSO ÀS OPERAÇÕES DA CENTRAL
			s = "SELECT" & _
					" Count(*) AS qtde" & _
				" FROM t_PERFIL_X_USUARIO" & _
					" INNER JOIN t_PERFIL ON t_PERFIL_X_USUARIO.id_perfil=t_PERFIL.id" & _
					" INNER JOIN t_PERFIL_ITEM ON t_PERFIL.id=t_PERFIL_ITEM.id_perfil" & _
					" INNER JOIN t_OPERACAO ON t_PERFIL_ITEM.id_operacao=t_OPERACAO.id" & _
				" WHERE" & _
					" (t_PERFIL_X_USUARIO.usuario='" & usuario & "')" & _
					" AND (t_OPERACAO.modulo='" & COD_OP_MODULO_CENTRAL & "')"
			set r = cn.Execute(s)
			if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			if Not rs.Eof then
				if Not IsNull(r("qtde")) then
					if CLng(r("qtde")) > 0 then blnPossuiPermissaoAcesso = true
					end if
				end if
			
			senha_real= ""
			s = Trim("" & rs("datastamp"))
			chave = gera_chave(FATOR_BD)
			decodifica_dado s, senha_real, chave
			if UCase(trim(senha_real)) = UCase(trim(senha)) then
				'SENHA CONFERE OK
				blnSenhaConfereOk = True
			else
				if senha_real <> "" then senha = ""
				end if
			end if
		
		rs.close
		set rs = nothing

		'REGISTRA HISTÓRICO DE LOGIN (NA SEQUÊNCIA DE PRIORIDADE)
		'MOTIVO: USUÁRIO NÃO CADASTRADO
		if Not blnUsuarioCadastrado then
			'USUÁRIO NÃO CADASTRADO
			s = "INSERT INTO t_LOGIN_HISTORICO (" & _
					"StSucesso" & _
					", IP" & _
					", sistema_responsavel" & _
					", Login" & _
					", Motivo" & _
					", IdCfgModulo" & _
				") VALUES (" & _
					"0" & _
					", '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
					", " & CStr(COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP) & _
					", '" & QuotedStr(usuario) & "'" & _
					", '" & COD_CONTROLE_LOGIN_FALHA__USUARIO_NAO_CADASTRADO & "'" & _
					", " & CStr(ID_MODULO__CENTRAL) & _
				")"
			cn.Execute(s)

		'MOTIVO: SENHA NÃO CONFERE
		elseif Not blnSenhaConfereOk then
			qtdeConsecutivaFalhaLogin = qtdeConsecutivaFalhaLogin + 1

			'SENHA NÃO CONFERE
			s = "INSERT INTO t_LOGIN_HISTORICO (" & _
					"IdTipoUsuarioContexto" & _
					", IdUsuario" & _
					", StSucesso" & _
					", IP" & _
					", sistema_responsavel" & _
					", Login" & _
					", Motivo" & _
					", IdCfgModulo" & _
				") VALUES (" & _
					COD_USUARIO_CONTEXTO__USUARIO_INTERNO & _
					", " & CStr(idUsuario) & _
					", 0" & _
					", '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
					", " & CStr(COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP) & _
					", '" & QuotedStr(usuario) & "'" & _
					", '" & COD_CONTROLE_LOGIN_FALHA__SENHA_INVALIDA & "'" & _
					", " & CStr(ID_MODULO__CENTRAL) & _
				")"
			cn.Execute(s)

			'Incrementa quantidade de tentativas consecutivas com falha
			s = "UPDATE t_USUARIO SET" & _
					" QtdeConsecutivaFalhaLogin = QtdeConsecutivaFalhaLogin + 1" & _
				" WHERE" & _
					" (usuario = '" & QuotedStr(usuario) & "')"
			cn.Execute(s)

			if (Not usuario_bloqueado_automatico) And (qtdeConsecutivaFalhaLogin >= max_tentativas_login) then
				'Usuário será bloqueado automaticamente no próximo login
				dtHrBloqueioAutomatico = Now
				s = "UPDATE t_USUARIO SET" & _
						" StLoginBloqueadoAutomatico = 1" & _
						", DataHoraBloqueadoAutomatico = getdate()" & _
						", EnderecoIpBloqueadoAutomatico = '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
					" WHERE" & _
						" (Id = " & CStr(idUsuario) & ")"
				cn.Execute(s)

				'Envia e-mail de alerta sobre o bloqueio automático
				set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaLoginBloqueadoAutomatico)
				if Trim("" & rEmailDestinatario.campo_texto) <> "" then
					ambiente_execucao = getParametroFromCampoTexto(ID_PARAMETRO_AMBIENTE_EXECUCAO_OWNER) & "/" & getParametroFromCampoTexto(ID_PARAMETRO_AMBIENTE_EXECUCAO_CONTEXTO)
					assunto_mensagem = getParametroFromCampoTexto(ID_PARAMETRO_SubjectEmailAlertaLoginBloqueadoAutomatico)
					corpo_mensagem = getParametroFromCampoTexto(ID_PARAMETRO_BodyEmailAlertaLoginBloqueadoAutomatico)
					remetente_mensagem = getParametroFromCampoTexto(ID_PARAMETRO_EmailRemetenteAlertaLoginBloqueadoAutomatico)
					
					assunto_mensagem = Replace(assunto_mensagem, "[AMBIENTE]", ambiente_execucao)
					assunto_mensagem = Replace(assunto_mensagem, "[LOGIN_USUARIO]", usuario)
					assunto_mensagem = Replace(assunto_mensagem, "[DATA_HORA_BLOQUEIO]", formata_data_hora_sem_seg(dtHrBloqueioAutomatico))

					corpo_mensagem = Replace(corpo_mensagem, "[AMBIENTE]", ambiente_execucao)
					corpo_mensagem = Replace(corpo_mensagem, "[LOGIN_USUARIO]", usuario)
					corpo_mensagem = Replace(corpo_mensagem, "[IdTipoUsuarioContexto]", CStr(COD_USUARIO_CONTEXTO__USUARIO_INTERNO))
					corpo_mensagem = Replace(corpo_mensagem, "[IdUsuario]", CStr(idUsuario))
					corpo_mensagem = Replace(corpo_mensagem, "[IP]", QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))))
					corpo_mensagem = Replace(corpo_mensagem, "[DATA_HORA_BLOQUEIO]", formata_data_hora_sem_seg(dtHrBloqueioAutomatico))
					corpo_mensagem = Replace(corpo_mensagem, "[MAX_TENTATIVAS_LOGIN]", CStr(max_tentativas_login))

					EmailSndSvcGravaMensagemParaEnvio remetente_mensagem, _
													"", _
													rEmailDestinatario.campo_texto, _
													"", _
													"", _
													assunto_mensagem, _
													corpo_mensagem, _
													Now, _
													id_email, _
													msg_erro_grava_email
					end if 'if Trim("" & rEmailDestinatario.campo_texto) <> ""
				end if 'if (Not usuario_bloqueado_automatico) And (qtdeConsecutivaFalhaLogin >= max_tentativas_login)

		'MOTIVO: USUÁRIO ESTÁ BLOQUEADO AUTOMATICAMENTE
		elseif usuario_bloqueado_automatico then
			'USUÁRIO ENCONTRA-SE BLOQUEADO AUTOMATICAMENTE
			s = "INSERT INTO t_LOGIN_HISTORICO (" & _
					"IdTipoUsuarioContexto" & _
					", IdUsuario" & _
					", StSucesso" & _
					", IP" & _
					", sistema_responsavel" & _
					", Login" & _
					", Motivo" & _
					", IdCfgModulo" & _
				") VALUES (" & _
					COD_USUARIO_CONTEXTO__USUARIO_INTERNO & _
					", " & CStr(idUsuario) & _
					", 0" & _
					", '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
					", " & CStr(COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP) & _
					", '" & QuotedStr(usuario) & "'" & _
					", '" & COD_CONTROLE_LOGIN_FALHA__BLOQUEADO_AUTOMATICO & "'" & _
					", " & CStr(ID_MODULO__CENTRAL) & _
				")"
			cn.Execute(s)

		'MOTIVO: USUÁRIO ESTÁ BLOQUEADO MANUALMENTE
		elseif usuario_bloqueado then
			'USUÁRIO BLOQUEADO MANUALMENTE
			s = "INSERT INTO t_LOGIN_HISTORICO (" & _
					"IdTipoUsuarioContexto" & _
					", IdUsuario" & _
					", StSucesso" & _
					", IP" & _
					", sistema_responsavel" & _
					", Login" & _
					", Motivo" & _
					", IdCfgModulo" & _
				") VALUES (" & _
					COD_USUARIO_CONTEXTO__USUARIO_INTERNO & _
					", " & CStr(idUsuario) & _
					", 0" & _
					", '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
					", " & CStr(COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP) & _
					", '" & QuotedStr(usuario) & "'" & _
					", '" & COD_CONTROLE_LOGIN_FALHA__BLOQUEADO_MANUAL & "'" & _
					", " & CStr(ID_MODULO__CENTRAL) & _
				")"
			cn.Execute(s)

		'MOTIVO: USUÁRIO NÃO POSSUI PERMISSÃO DE ACESSO SUFICIENTE
		elseif Not blnPossuiPermissaoAcesso then
			'USUÁRIO SEM PERMISSÃO DE ACESSO
			s = "INSERT INTO t_LOGIN_HISTORICO (" & _
					"IdTipoUsuarioContexto" & _
					", IdUsuario" & _
					", StSucesso" & _
					", IP" & _
					", sistema_responsavel" & _
					", Login" & _
					", Motivo" & _
					", IdCfgModulo" & _
				") VALUES (" & _
					COD_USUARIO_CONTEXTO__USUARIO_INTERNO & _
					", " & CStr(idUsuario) & _
					", 0" & _
					", '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
					", " & CStr(COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP) & _
					", '" & QuotedStr(usuario) & "'" & _
					", '" & COD_CONTROLE_LOGIN_FALHA__PERMISSAO_INSUFICIENTE & "'" & _
					", " & CStr(ID_MODULO__CENTRAL) & _
				")"
			cn.Execute(s)
			end if 'if - elseif (falhas de login)


		if Not blnUsuarioCadastrado then
			cn.Close
			Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICACAO)
			end if

		if Not blnSenhaConfereOk then
			cn.Close
			Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICACAO)
			end if

		if usuario_bloqueado_automatico then
			cn.Close
			Response.Redirect("aviso.asp?id=" & ERR_USUARIO_BLOQUEADO_AUTOMATICO)
			end if
		
		if usuario_bloqueado then
			cn.Close
			Response.Redirect("aviso.asp?id=" & ERR_USUARIO_BLOQUEADO)
			end if

		if Not blnPossuiPermissaoAcesso then
			cn.Close
			Response.Redirect("aviso.asp?id=" & ERR_USUARIO_BLOQUEADO)
			end if


		s_lista_operacoes_permitidas = obtem_operacoes_permitidas_usuario(cn, usuario)
		nivel_acesso_bloco_notas = obtem_nivel_acesso_bloco_notas_pedido(cn, usuario)
        nivel_acesso_chamado = obtem_nivel_acesso_chamado_pedido(cn, usuario)
		
		Session("usuario_atual") = usuario
		Session("senha_atual") = senha
		Session("lista_operacoes_permitidas") = s_lista_operacoes_permitidas
		Session("usuario_nome_atual") = usuario_nome
		Session("nivel_acesso_bloco_notas") = Cstr(nivel_acesso_bloco_notas)
        Session("nivel_acesso_chamado") = Cstr(nivel_acesso_chamado)
		
		strSessionCtrlTicket = GeraTicketSessionCtrl(usuario)
		Session("SessionCtrlTicket") = strSessionCtrlTicket
		
		Session("SessionCtrlInfo") = MontaSessionCtrlInfo(usuario, SESSION_CTRL_MODULO_CENTRAL, "", strSessionCtrlTicket, Session("DataHoraLogon"), Now)
		
		s = "UPDATE t_USUARIO SET" & _
				" dt_ult_acesso = " & bd_formata_data_hora(Now) & _
				", QtdeConsecutivaFalhaLogin = 0" & _
				", SessionCtrlDtHrLogon = " & bd_formata_data_hora(Session("DataHoraLogon")) & _
				", SessionCtrlModulo = '" & SESSION_CTRL_MODULO_CENTRAL & "'" & _
				", SessionCtrlLoja = NULL" & _
				", SessionCtrlTicket = '" & strSessionCtrlTicket & "'" & _
				", SessionTokenModuloCentral = newid()" & _
				", DtHrSessionTokenModuloCentral = getdate()" & _
			" WHERE" & _
				" (usuario = '" & QuotedStr(usuario) & "')"
		cn.Execute(s)
		
		s = "INSERT INTO t_SESSAO_HISTORICO (" & _
				"Usuario, " & _
				"SessionCtrlTicket, " & _
				"DtHrInicio, " & _
				"Loja, " & _
				"Modulo, " & _
				"IP, " & _
				"UserAgent" & _
			") VALUES (" & _
				"'" & QuotedStr(usuario) & "'," & _
				"'" & strSessionCtrlTicket & "'," & _
				bd_formata_data_hora(Session("DataHoraLogon")) & "," & _
				"''," & _
				"'" & SESSION_CTRL_MODULO_CENTRAL & "'," & _
				"'" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'," & _
				"'" & QuotedStr(Trim("" & Request.ServerVariables("HTTP_USER_AGENT"))) & "'" & _
			")"
		cn.Execute(s)

		'Login bem sucedido
		s = "INSERT INTO t_LOGIN_HISTORICO (" & _
				"IdTipoUsuarioContexto" & _
				", IdUsuario" & _
				", StSucesso" & _
				", IP" & _
				", sistema_responsavel" & _
				", Login" & _
				", IdCfgModulo" & _
			") VALUES (" & _
				COD_USUARIO_CONTEXTO__USUARIO_INTERNO & _
				", " & CStr(idUsuario) & _
				", 1" & _
				", '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
				", " & CStr(COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP) & _
				", '" & QuotedStr(usuario) & "'" & _
				", " & CStr(ID_MODULO__CENTRAL) & _
			")"
		cn.Execute(s)


		if IsNull(dt_ult_alteracao_senha) then Response.Redirect("senha.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		
'		COM ESTE REDIRECT, A PÁGINA INICIAL PASSA A TER NA QUERY STRING OS DADOS NECESSÁRIOS P/ RECRIAR A
'		SESSÃO EXPIRADA.
'		QUANDO O USUÁRIO FAZIA O LOGON E NÃO NAVEGAVA P/ NENHUMA OUTRA TELA, AO CLICAR EM F5 NÃO ERA
'		POSSÍVEL RECRIAR A SESSÃO.
		Response.Redirect("resumo.asp?FlagPrimeiraExecucao=1&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		end if  'if (confere_login_no_bd)
	
	Dim strScript
	Dim vMsg()
	if Trim(Session("verificar_quadro_avisos")) <> "" then
		Session("verificar_quadro_avisos") = " "
		if recupera_avisos_nao_lidos("", usuario, vMsg) then Response.Redirect("quadroavisomostra.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		end if
	
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	CD
	dim i, qtde_nfe_emitente
	dim v_usuario_x_nfe_emitente
	dim id_nfe_emitente_selecionado
	v_usuario_x_nfe_emitente = obtem_lista_usuario_x_nfe_emitente(usuario)
	
	qtde_nfe_emitente = 0
	for i=Lbound(v_usuario_x_nfe_emitente) to UBound(v_usuario_x_nfe_emitente)
		if Not Isnull(v_usuario_x_nfe_emitente(i)) then
			qtde_nfe_emitente = qtde_nfe_emitente + 1
			id_nfe_emitente_selecionado = v_usuario_x_nfe_emitente(i)
			end if
		next
	
	if qtde_nfe_emitente > 1 then
	'	HÁ MAIS DO QUE 1 CD, ENTÃO SERÁ EXIBIDA A LISTA P/ O USUÁRIO SELECIONAR UM CD
		id_nfe_emitente_selecionado = 0
		end if

'   LIMPA EVENTUAIS LOCKS REMANESCENTES NOS RELATÓRIOS
    s = "UPDATE tCRUP SET" & _
            " locked = 0," & _
            " cod_motivo_lock_released = " & CTRL_RELATORIO_CodMotivoLockReleased_AcessadaTelaInicialCentral & "," & _
            " dt_hr_lock_released = getdate()" & _
        " FROM t_CTRL_RELATORIO_USUARIO_X_PEDIDO tCRUP INNER JOIN t_CTRL_RELATORIO tCR ON (tCRUP.id_relatorio = tCR.id)" & _
        " WHERE" & _
            " (tCR.modulo = 'CENTRAL')" & _
            " AND (tCRUP.usuario = '" & QuotedStr(Trim(Session("usuario_atual"))) & "')" & _
            " AND (locked = 1)"
    cn.Execute(s)
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
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
    $(document).ready(function() {
        $("#c_dt_entrega").hUtilUI('datepicker_padrao');
        $("#c_dt_recebido").hUtilUI('datepicker_padrao');

        $(document).tooltip();
        
        $("#divTelaCheia").css('filter', 'alpha(opacity=30)');
        <% if SWITCH_QUADRO_AVISO_POPUP = 1 then %>

        CarregaAvisoNovo();
        <%end if%>

    });
    
</script>
    <script>
    $(window).load(function() {
        $("#divQuadroAviso").hide();
        $("#divQuadroAvisoPai").hide();
        $("#divQuadroAvisoConteudo").hide();
    });
</script>
<% if SWITCH_QUADRO_AVISO_POPUP = 1 then %>
<script type="text/javascript">
    function TrataDadosAvisoNovo() {
        var f, i, strResp, textarea, label, usuario, xmlResp;
        
        if (objAjaxAvisoNovo.readyState == AJAX_REQUEST_IS_COMPLETE) {
            strResp = objAjaxAvisoNovo.responseText;
            if (strResp == "") {
                window.status = "Concluído";
                setTimeout(CarregaAvisoNovo, <%=TIMER_CARREGA_AVISO_NOVO_MILISSEGUNDOS%>);
                return;
            }

            $("#divQuadroAvisoConteudo").children().remove();

            if (strResp != "") { 
                try {
                    var div = document.getElementById("divQuadroAvisoConteudo");
                    var divPai = document.getElementById("divQuadroAviso");
                    xmlResp = objAjaxAvisoNovo.responseXML.documentElement;

                    for (i = 0; i < xmlResp.getElementsByTagName('registro').length; i++) {
                        divPai.style.textAlign = 'center';
                        textarea = document.createElement('TEXTAREA');
                        span = document.createElement('SPAN');
                        checkbox = document.createElement('INPUT');
                        label = document.createElement('LABEL');

                        span.className = 'Lbl';
                        span.style.position = 'relative';
                        span.style.left = '2px';
                        span.innerText = "Divulgado em: " + xmlResp.getElementsByTagName('datahora')[i].childNodes[0].nodeValue;
                        textarea.name = 'mensagem';
                        textarea.className = 'QuadroAviso';
                        textarea.readOnly = true;
                        textarea.style.display = 'block';
                        textarea.style.marginBottom = '0';
                        div.style.textAlign = 'left';
                        div.style.marginBottom = '20px';
                        textarea.innerText = xmlResp.getElementsByTagName('mensagem')[i].childNodes[0].nodeValue;
                        checkbox.type = 'checkbox';
                        checkbox.className = 'CBOX';
                        checkbox.name = 'xMsg';
                        checkbox.id = 'xMsg';
                        checkbox.value = xmlResp.getElementsByTagName('id')[i].childNodes[0].nodeValue;
                        label.className = 'CBOX';
                        label.innerText = 'Não exibir mais este aviso';

                        div.appendChild(span);
                        div.appendChild(document.createElement('BR'));
                        div.appendChild(textarea);
                        div.appendChild(document.createElement('BR'));
                        div.appendChild(checkbox);
                        div.appendChild(label);
                        div.appendChild(document.createElement('BR'));
                        div.appendChild(document.createElement('BR'));

                    }
                    $("#divQuadroAvisoPai").fadeIn();
                    $("#divQuadroAviso").fadeIn();
                    $("#divQuadroAvisoConteudo").fadeIn();
                    $("#divTelaCheia").fadeIn();
                    
                }
                catch (e) {
                    alert("Falha na consulta de novos avisos!!");
                }
            }
            window.status = "Concluído";
        }
    }

    function CarregaAvisoNovo() {
        var f, strUrl, usuario;
            usuario = "<%=usuario%>";
            objAjaxAvisoNovo = GetXmlHttpObject();
            if (objAjaxAvisoNovo == null) {
                alert("O browser NÃO possui suporte ao AJAX!!");
                return;
            }

            window.status = "Pesquisando por novos avisos ...";

            strUrl = "../Global/AjaxCarregaAvisosNovos.asp";
            strUrl = strUrl + "?loja=";
            strUrl = strUrl + "&usuario=" + usuario;
            //  Prevents server from using a cached file
            strUrl = strUrl + "&sid=" + Math.random() + Math.random();
            objAjaxAvisoNovo.onreadystatechange = TrataDadosAvisoNovo;
            objAjaxAvisoNovo.open("GET", strUrl, true);
            objAjaxAvisoNovo.send(null);
    }

    function fechaQuadroAviso(f, optLido) {

        $("#divQuadroAvisoPai").hide();
        $("#divTelaCheia").hide();
        $("#divQuadroAvisoConteudo").children().remove();
        setTimeout(CarregaAvisoNovo, <%=TIMER_CARREGA_AVISO_NOVO_MILISSEGUNDOS%>);
        
    }

    function RemoveAviso(f, optLido) {
        var i, max, aviso_selecionado;
        max = f["xMsg"].length;
        aviso_selecionado = "";
        for (i = 0; i < max; i++) {
            if (f["xMsg"][i].checked) {
                if (f["xMsg"][i].value != "") {
                    if (aviso_selecionado != "") aviso_selecionado = aviso_selecionado + "|";
                    aviso_selecionado = aviso_selecionado + f["xMsg"][i].value;
                }
            }
        }

        if (aviso_selecionado == "") {
            alert("Nenhum aviso selecionado!!");
            return;
        }

        var f, strUrl, usuario;
        usuario = "<%=usuario%>";
        objAjaxAvisoNovo = GetXmlHttpObject();
        if (objAjaxAvisoNovo == null) {
            alert("O browser NÃO possui suporte ao AJAX!!");
            return;
        }

        window.status = "Aguarde ...";

        strUrl = "../Global/AjaxGravaAvisoExibidoLido.asp";
        strUrl = strUrl + "?aviso_selecionado=" + aviso_selecionado;
        strUrl = strUrl + "&usuario=" + usuario;
        strUrl = strUrl + "&optLido=" + optLido;
        //  Prevents server from using a cached file
        strUrl = strUrl + "&sid=" + Math.random() + Math.random();
        objAjaxAvisoNovo.onreadystatechange = function () {
            if (objAjaxAvisoNovo.readyState == AJAX_REQUEST_IS_COMPLETE) {
                $("#divQuadroAvisoPai").hide();
                $("#divTelaCheia").hide();
                $("#divQuadroAvisoConteudo").children().remove();
                window.status = "Concluído";
                setTimeout(CarregaAvisoNovo, <%=TIMER_CARREGA_AVISO_NOVO_MILISSEGUNDOS%>);
                
            }
        }
        objAjaxAvisoNovo.open("GET", strUrl, true);
        objAjaxAvisoNovo.send(null); 
    }
</script>
<% end if %>
<script language="JavaScript" type="text/javascript">
window.focus();
</script>

<% if eh_primeira_execucao then %>
<script language="JavaScript" type="text/javascript">
configura_painel();
</script>
<% end if %>

<script language="JavaScript" type="text/javascript">
var fCepPopup;

<%=monta_funcao_js_normaliza_numero_pedido_e_sufixo%>

function ProcessaSelecaoCEP(){};

function AbrePesquisaCep(){
var strUrl;
	try
		{
	//  SE JÁ HOUVER UMA JANELA DE PESQUISA DE CEP ABERTA, GARANTE QUE ELA SERÁ FECHADA 
	// E UMA NOVA SERÁ CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')	
		fCepPopup=window.open("", "AjaxCepPesqPopup","status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=5,height=5,left=0,top=0");
		fCepPopup.close();
		}
	catch (e) {
	 // NOP
		}
	ProcessaSelecaoCEP=null;
	strUrl="../Global/AjaxCepPesqPopup.asp?ModoApenasConsulta=S";
	fCepPopup=window.open(strUrl, "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=980,height=650,left=0,top=0");
	fCepPopup.focus();
}

function fPAGTOPARCIALConcluir( f ){
	if (trim(f.c_pedido.value)=="") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedido.focus();
		return;
		}
	if (converte_numero(f.c_valor.value)==0) {
		alert("Valor inválido!!");
		f.c_valor.focus();
		return;
		}
	f.action="pagtoparcialconsiste.asp";
	f.submit();
}

function fPagtoParcialMarketplaceConcluir( f ){
	if (trim(f.c_pedidos_pagto_parcial_marketplace.value)=="") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos_pagto_parcial_marketplace.focus();
		return;
	}
	f.action="PagtoParcialMarketplaceConsiste.asp";
	f.submit();
}

function fQUITACAOConcluir( f ){
	if (trim(f.c_pedidos.value)=="") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos.focus();
		return;
		}
	f.action="pagtoquitacaoconsiste.asp";
	f.submit();
}

function fQuitacaoMarketplaceConcluir( f ){
	if (trim(f.c_pedidos_quitacao_marketplace.value)=="") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos_quitacao_marketplace.focus();
		return;
	}
	f.action="PagtoQuitacaoMarketplaceConsiste.asp";
	f.submit();
}

function fSEPConcluir( f ){
	if (trim(f.c_pedidos.value)=="") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos.focus();
		return;
		}

	if (converte_numero(f.c_qtde_nfe_emitente.value) == 0) {
		alert("Nenhum CD habilitado para o usuário!!");
		return;
	}

	if (trim(f.c_nfe_emitente.value) == "") {
		alert("É necessário selecionar um CD!!");
		return;
	}

	if (converte_numero(f.c_nfe_emitente.value) == 0) {
		alert("CD selecionado é inválido!!");
		return;
	}

	f.action = "pedidoseparacaoconsiste.asp";
	f.submit();
}

function fSEPUsandoRelConcluir(f) {
	if (trim(f.c_nsu_rel_separacao_zona.value) == "") {
		alert("Informe o NSU do Relatório!!");
		f.c_nsu_rel_separacao_zona.focus();
		return;
	}
	
	if (converte_numero(f.c_qtde_nfe_emitente.value) == 0)
	{
		alert("Nenhum CD habilitado para o usuário!!");
		return;
	}

	if (trim(f.c_nfe_emitente.value) == "") {
		alert("É necessário selecionar um CD!!");
		return;
	}

	if (converte_numero(f.c_nfe_emitente.value) == 0) {
		alert("CD selecionado é inválido!!");
		return;
	}

	f.action = "PedidoSeparacaoUsandoRelConsiste.asp";
	f.submit();
}

function fCrOkConcluir( f ){
	if (trim(f.c_pedidos.value)=="") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos.focus();
		return;
		}
	f.action="CreditoOkAguardandoDepositoConsiste.asp";
	f.submit();
}

function fDepDesbloqCrOkConcluir(f) {
	if (trim(f.c_pedidos.value) == "") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos.focus();
		return;
	}
	f.action = "CreditoOkDepositoAguardandoDesbloqueioConsiste.asp";
	f.submit();
}

function fPendVendasCrOkConcluir(f) {
	if (trim(f.c_pedidos.value) == "") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos.focus();
		return;
	}
	f.action = "CreditoOkPendenteVendasConsiste.asp";
	f.submit();
}

function fETGConcluir( f ){
	if (trim(f.c_pedidos.value)=="") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos.focus();
		return;
		}

	if (converte_numero(f.c_qtde_nfe_emitente.value) == 0) {
		alert("Nenhum CD habilitado para o usuário!!");
		return;
	}

	if (trim(f.c_nfe_emitente.value) == "") {
		alert("É necessário selecionar um CD!!");
		return;
	}

	if (converte_numero(f.c_nfe_emitente.value) == 0) {
		alert("CD selecionado é inválido!!");
		return;
	}

	f.action = "pedidoentregaconsiste.asp";
	f.submit();
}

function fETGRomaneioConcluir(f) {
	if (trim(f.c_nsu_romaneio.value) == "") {
		alert("Informe o NSU do Romaneio!!");
		f.c_nsu_romaneio.focus();
		return;
	}

	if (converte_numero(f.c_qtde_nfe_emitente.value) == 0) {
		alert("Nenhum CD habilitado para o usuário!!");
		return;
	}

	if (trim(f.c_nfe_emitente.value) == "") {
		alert("É necessário selecionar um CD!!");
		return;
	}

	if (converte_numero(f.c_nfe_emitente.value) == 0) {
		alert("CD selecionado é inválido!!");
		return;
	}

	f.action = "PedidoEntregaUsandoRomaneioConsiste.asp";
	f.submit();
}

function fPedRecConcluir( f ){
	if ( (trim(f.c_pedidos.value)=="")&&(trim(f.c_obs2.value)=="") ) {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos.focus();
		return;
		}
	if ((!f.ckb_entrega.checked)&&(!f.ckb_recebido.checked)) {
		alert("É necessário selecionar pelo menos uma das operações:\n        a) Entrega\n        b) Recebido");
		return;
		}

	if (converte_numero(f.c_qtde_nfe_emitente.value) == 0) {
		alert("Nenhum CD habilitado para o usuário!!");
		return;
	}

	if (trim(f.c_nfe_emitente.value) == "") {
		alert("É necessário selecionar um CD!!");
		return;
	}

	if (converte_numero(f.c_nfe_emitente.value) == 0) {
		alert("CD selecionado é inválido!!");
		return;
	}

	f.action = "PedidoRecebidoConsiste.asp";
	f.submit();
}

function fDTETGConcluir( f ){
	if (trim(f.c_dt_entrega.value)=="") {
		alert("Informe a data de coleta!!");
		f.c_dt_entrega.focus();
		return;
		}
	if (!isDate(f.c_dt_entrega)) {
		alert("Data inválida!!");
		f.c_dt_entrega.focus();
		return;
		}
	if (trim(f.c_pedidos.value)=="") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos.focus();
		return;
		}

	if (converte_numero(f.c_qtde_nfe_emitente.value) == 0) {
		alert("Nenhum CD habilitado para o usuário!!");
		return;
	}

	if (trim(f.c_nfe_emitente.value) == "") {
		alert("É necessário selecionar um CD!!");
		return;
	}

	if (converte_numero(f.c_nfe_emitente.value) == 0) {
		alert("CD selecionado é inválido!!");
		return;
	}

	f.action = "PedidoEntregaMarcParaConsiste.asp";
	f.submit();
}

function fAnotaTranspConcluir( f ){
	if (trim(f.c_transportadora.value)=="") {
		alert("Informe a transportadora!!");
		f.c_transportadora.focus();
		return;
		}
	if (trim(f.c_pedidos.value)=="") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos.focus();
		return;
	}

	if (converte_numero(f.c_qtde_nfe_emitente.value) == 0) {
		alert("Nenhum CD habilitado para o usuário!!");
		return;
	}

	if (trim(f.c_nfe_emitente.value) == "") {
		alert("É necessário selecionar um CD!!");
		return;
	}

	if (converte_numero(f.c_nfe_emitente.value) == 0) {
		alert("CD selecionado é inválido!!");
		return;
	}

	f.action="PedidoAnotaTranspConsiste.asp";
	f.submit();
}

function fCOMISSAOConcluir( f ){
var blnFlagOk,idx;
	if (trim(f.c_pedidos.value)=="") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos.focus();
		return;
		}
	
	blnFlagOk=false;
	idx=-1;
//  Marcar como comissão paga?
	idx++;
	if (f.rb_comissao_paga[idx].checked) blnFlagOk=true;
//  Marcar como comissão não-paga?
	idx++;
	if (f.rb_comissao_paga[idx].checked) blnFlagOk=true;
	
	if (!blnFlagOk) {
		alert("Informe se a comissão deve ser assinalada como paga ou não-paga!!");
		return;
		}
	
	f.action="ComissaoPagaConsiste.asp";
	f.submit();
}

function fCOMISSAODESCConcluir( f ){
var blnFlagOk,idx;
	if (trim(f.c_pedidos.value)=="") {
		alert("Nenhum pedido foi especificado!!");
		f.c_pedidos.focus();
		return;
		}
	
	blnFlagOk=false;
	idx=-1;
//  Marcar como comissão descontada?
	idx++;
	if (f.rb_comissao_descontada[idx].checked) blnFlagOk=true;
//  Marcar como comissão não-descontada?
	idx++;
	if (f.rb_comissao_descontada[idx].checked) blnFlagOk=true;
	
	if (!blnFlagOk) {
		alert("Informe se a devolução/perda deve ser assinalada como descontada ou não-descontada das comissões!!");
		return;
		}
	
	f.action="ComissaoDescConsiste.asp";
	f.submit();
}
</script>

<%
	strScript = _
		"<script language='JavaScript' type='text/javascript'>" & chr(13) & _
		"function fESTOQConcluir( f ){" & chr(13) & _
		"var s, iop;" & chr(13) & _
		"	iop=0;" & chr(13) & _
		"	s='';" & chr(13) & _
		"" & chr(13)
	
	if operacao_permitida(OP_CEN_ENTRADA_MERCADORIAS_ESTOQUE, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // ENTRADA DE MERCADORIAS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_op[iop].checked) {" & chr(13) & _
			"		s='EstoqueEntrada.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_ENTRADA_MERCADORIAS_ESTOQUE, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // ENTRADA DE MERCADORIAS VIA XML" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_op[iop].checked) {" & chr(13) & _
			"		s='EstoqueEntradaViaXmlUpload.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_CONVERSOR_KITS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // CONVERSOR DE KITS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_op[iop].checked) {" & chr(13) & _
			"		s='estoqueconversorkit.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_BASICO, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // TRANSFERÊNCIA ENTRE ESTOQUES" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_op[iop].checked) {" & chr(13) & _
			"		s='estoquetransfere.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_TRANSF_ENTRE_PED_PROD_ESTOQUE_VENDIDO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // TRANSFERÊNCIA ENTRE PEDIDOS DE PRODUTOS DO ESTOQUE VENDIDO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_op[iop].checked) {" & chr(13) & _
			"		s='estoquetransferepedido.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // TRANSFERÊNCIA DE PRODUTOS ENTRE CD'S" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_op[iop].checked) {" & chr(13) & _
			"		s='EstoqueTransfereEntreCDsFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	strScript = strScript & _
		"	if (s=='') {" & chr(13) & _
		"		alert('Escolha uma das funções!!');" & chr(13) & _
		"		return false;" & chr(13) & _
		"		}" & chr(13) & _
		"" & chr(13) & _
		"	window.status = 'Aguarde ...';" & chr(13) & _
		"	f.action=s;" & chr(13) & _
		"	f.submit();" & chr(13) & _
		"}" & chr(13) & _
		"" & chr(13) & _
		"</script>" & chr(13)
	
	Response.Write strScript
%>

<%
	strScript = _
		"<script language='JavaScript' type='text/javascript'>" & chr(13) & _
		"function fRELConcluir( f ){" & chr(13) & _
		"var s_dest, iop;" & chr(13) & _
		"	iop=0;" & chr(13) & _
		"	s_dest='';" & chr(13) & _
		"" & chr(13) & _
		" // **********  GERAL  **********" & chr(13) & _
		"" & chr(13)
	
	if operacao_permitida(OP_CEN_REL_MULTICRITERIO_PEDIDOS_ANALITICO, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_CEN_REL_MULTICRITERIO_PEDIDOS_SINTETICO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO MULTICRITÉRIO DE PEDIDOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidosMCrit.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_MULTICRITERIO_ORCAMENTOS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO MULTICRITÉRIO DE ORÇAMENTOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelOrcamentosMCrit.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	strScript = strScript & _
		"" & chr(13) & _
		" // **********  COMERCIAL  **********" & chr(13) & _
		"" & chr(13)
	
	if operacao_permitida(OP_CEN_REL_PRODUTOS_VENDIDOS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE PRODUTOS VENDIDOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelProdVendidos.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_PEDIDOS_COLOCADOS_NO_MES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // PEDIDOS COLOCADOS NO MÊS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidosColocados.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_VENDAS_COM_DESCONTO_SUPERIOR, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // VENDAS COM DESCONTO SUPERIOR" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		f.pagina_destino.value='RelVendasAbaixoMin.asp';" & chr(13) & _
			"		f.titulo_relatorio.value='Vendas com Desconto Superior';" & chr(13) & _
			"		f.filtro_obrigatorio_data_inicio.value = 'S';" & chr(13) & _
			"		f.filtro_obrigatorio_data_termino.value = 'S';" & chr(13) & _
			"		s_dest='FiltroPeriodo.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_COMISSAO_VENDEDORES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // COMISSÃO AOS VENDEDORES" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissao.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_SINTETICO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // COMISSÃO AOS VENDEDORES SINTÉTICO (TABELA PROGRESSIVA)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoTabelaProgressivaSintetico.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_ANALITICO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // COMISSÃO AOS VENDEDORES ANALÍTICO (TABELA PROGRESSIVA)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoTabelaProgressivaAnalitico.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // COMISSÃO AOS INDICADORES (ALTERADO P/ RELATÓRIO DE PEDIDOS INDICADORES)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadores.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE PEDIDOS INDICADORES (VIA NFS-e)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadoresNFSeP01Filtro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE PEDIDOS INDICADORES (VIA NFS-e) (CONSULTA)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadoresNFSeConsP01Filtro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

   ' if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO , s_lista_operacoes_permitidas) then
	'	strScript = strScript & _
	'		" // Comissão de indicadores: Cadastrar NF " & chr(13) & _
	'		"	iop++;" & chr(13) & _
	'		"	if (f.rb_rel[iop].checked) {" & chr(13) & _
	'		"		s_dest='ComissaoIndicadoresCadastraNf.asp';" & chr(13) & _
	'		"		}" & chr(13) & _
	'		"" & chr(13)
	'	end if

    if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO , s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // Relatório pedidos indicadores (processamento) " & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadoresPag.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // Relatório pedidos indicadores (consulta)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadoresConsultaPedido.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

     if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELAÇÃO DE DEPÓSITOS " & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadoresConsulta.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO , s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // Relatório pedidos indicadores com desconto (processamento) " & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadoresPagDesc.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // Relatório pedidos indicadores com desconto (consulta)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadoresDescConsultaPedido.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELAÇÃO DE DEPÓSITOS COM DESCONTO " & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadoresDescConsulta.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // COMISSÃO INDICADORES: PESQUISA INDICADOR " & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadoresPesquisa.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_CEN_REL_FATURAMENTO_VENDEDORES_EXT, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // FATURAMENTO VENDEDORES EXTERNOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelFatVendExt.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_COMISSAO_LOJA_POR_INDICACAO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // COMISSÃO DE LOJA POR INDICAÇÃO AOS VENDEDORES EXTERNOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoLojaIndicou.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_ANALISE_PEDIDOS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE ANÁLISE DE PEDIDOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelAnalisePedidos.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_DIVERGENCIA_CLIENTE_INDICADOR, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE DIVERGÊNCIA CLIENTE/INDICADOR" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelDivergenciaClienteIndicadorFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_METAS_INDICADOR, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE METAS DO INDICADOR" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelMetasIndicadorFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_PERFORMANCE_INDICADOR, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE PERFORMANCE POR INDICADOR" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPerformanceIndicadorFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_VENDAS_POR_BOLETO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE VENDAS POR BOLETO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelVendasPorBoletoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_PERFIL_PAGAMENTO_BOLETOS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // PERFIL DE PAGAMENTO DOS BOLETOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPagamentosBoletos.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_ECOMMERCE_EXPORTACAO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // E-COMMERCE: EXPORTAÇÃO DA TABELA DE PRODUTOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelEcommerceExportacao.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_DADOS_TABELA_DINAMICA, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // DADOS PARA TABELA DINÂMICA" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelTabelaDinamicaFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
    if operacao_permitida(OP_CEN_REL_PEDIDOS_CANCELADOS, s_lista_operacoes_permitidas) then
	    strScript = strScript & _
		    " // Relatório de Pedidos Cancelados " & chr(13) & _
		    "	iop++;" & chr(13) & _
		    "	if (f.rb_rel[iop].checked) {" & chr(13) & _
		    "		s_dest='RelPedidoCancelado.asp';" & chr(13) & _
		    "		}" & chr(13) & _
		    "" & chr(13)
	 end if

	
	strScript = strScript & _
		"" & chr(13) & _
		" // **********  ADMINISTRATIVO  **********" & chr(13) & _
		"" & chr(13)
	
    if operacao_permitida(OP_CEN_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE PRÉ-DEVOLUÇÃO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidoPreDevolucao.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_PRE_DEVOLUCAO_RECEBIMENTO_MERCADORIA, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE PRÉ-DEVOLUÇÃO REGISTRA MERCADORIA RECEBIDA" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidoPreDevolucaoMercadoriaRecebe.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_CEN_REL_SEPARACAO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO PARA SEPARAÇÃO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		f.pagina_destino.value='RelSeparacao.asp';" & chr(13) & _
			"		f.titulo_relatorio.value='Separação';" & chr(13) & _
			"		f.filtro_obrigatorio_data_inicio.value = 'N';" & chr(13) & _
			"		f.filtro_obrigatorio_data_termino.value = 'N';" & chr(13) & _
			"		s_dest='RelSeparacaoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_SEPARACAO_ZONA, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO PARA SEPARAÇÃO (ZONA)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelSeparacaoZonaFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		strScript = strScript & _
			" // RELATÓRIO PARA SEPARAÇÃO (ZONA) - CONSULTA" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelSeparacaoZonaConsultaFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_PRODUTOS_ESTOQUE_DEVOLUCAO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO: PRODUTOS NO ESTOQUE DE DEVOLUÇÃO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelEstoqueDevolucaoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_DEVOLUCAO_PRODUTOS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // DEVOLUÇÃO DE PRODUTOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelDevolucao.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_DEVOLUCAO_PRODUTOS2, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // DEVOLUÇÃO DE PRODUTOS II" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelDevolucao2Filtro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_PRODUTOS_SPLIT_POSSIVEL, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // PRODUTOS ALOCADOS PARA PEDIDOS COM STATUS 'SPLIT POSSÍVEL'" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelProdAlocPedSplitavel.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_ESTOQUE_VENDA_CRITICO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // ESTOQUE DE VENDA CRÍTICO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelEstoqueVendaCritico.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_MEIO_DIVULGACAO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // PEDIDOS COLOCADOS CLASSIFICADOS PELO MEIO DE DIVULGAÇÃO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidosColocadosMidia.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_LOG_ESTOQUE, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE LOG DE MOVIMENTAÇÕES DE ESTOQUE" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelLogEstoque.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_FATURAMENTO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // FATURAMENTO (ANTIGO)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelVendas.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_FATURAMENTO_CMVPV, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // FATURAMENTO (CMV PV)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelVendasCmvPv.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_FATURAMENTO2, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // FATURAMENTO II" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelFaturamento2Filtro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_FATURAMENTO3, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // FATURAMENTO III" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelFaturamento3Filtro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_VENDAS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // VENDAS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelVendasVariante.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_GERENCIAL_DE_VENDAS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO GERENCIAL DE VENDAS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelGerencialVendasFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_ROMANEIO_ENTREGA, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // ROMANEIO DE ENTREGA" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RomaneioPreFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_ANALISE_CREDITO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // ANÁLISE DE CRÉDITO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelAnaliseCreditoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_PESQUISA_INDICADORES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // PESQUISA DE INDICADORES" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='PesquisaDeIndicadoresFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_INDICADOR_SEM_AVALIACAO_DESEMPENHO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // INDICADOR SEM AVALIAÇÃO DE DESEMPENHO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelIndicadorSemAvaliacaoDesempenhoExec.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_CHECAGEM_NOVOS_PARCEIROS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE CHECAGEM DE NOVOS PARCEIROS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelChecagemNovosParceirosFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_FRETE_ANALITICO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE FRETE (ANALÍTICO)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelFreteAnaliticoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_FRETE_SINTETICO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE FRETE (SINTÉTICO)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelFreteSinteticoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_PEDIDO_NAO_RECEBIDO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE PEDIDOS NÃO RECEBIDOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidosNaoRecebidosFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_PEDIDO_MARKETPLACE_NAO_RECEBIDO_CLIENTE, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE PEDIDOS DE MARKETPLACE NÃO RECEBIDOS PELO CLIENTE" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidosMktplaceNaoRecebidos.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_REGISTRO_PEDIDO_MARKETPLACE_RECEBIDO_CLIENTE, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // REGISTRO DE PEDIDOS DE MARKETPLACE RECEBIDOS PELO CLIENTE" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidosMktplaceRecebidos.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE OCORRÊNCIAS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidoOcorrenciaFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_ESTATISTICAS_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE ESTATÍSTICAS DE OCORRÊNCIAS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidoOcorrenciaEstatisticasFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // ACOMPANHAMENTO DE CHAMADOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelAcompanhamentoChamadosFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE CHAMADOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidoChamadoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_ESTATISTICAS_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE ESTATÍSTICAS DE CHAMADOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidoChamadoEstatisticasFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_PESQUISA_ORDEM_SERVICO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // PESQUISA DE ORDEM DE SERVIÇO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPesquisaOrdemServicoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_CLIENTE_SPC, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE CLIENTES NEGATIVADOS (SPC)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelClientesNegativadosFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_TRANSACOES_CIELO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE TRANSAÇÕES CIELO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelTransacoesCieloFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_TRANSACOES_CIELO_ANDAMENTO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE TRANSAÇÕES CIELO EM ANDAMENTO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelTransacoesCieloAndamentoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_BRASPAG_TRANSACOES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE TRANSAÇÕES BRASPAG" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelBraspagTransacoesFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_BRASPAG_AF_REVIEW, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // REVISÃO MANUAL ANTIFRAUDE BRASPAG" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelBraspagAfReviewFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_CEN_REL_BRASPAG_TRANSACOES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE TRANSAÇÕES BRASPAG/CLEARSALE" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelBraspagClearsaleTransacoesFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	strScript = strScript & _
		"" & chr(13) & _
		" // **********  COMPRAS E LOGÍSTICA  **********" & chr(13) & _
		"" & chr(13)
	
'	RELATÓRIOS DE COMPRAS E LOGÍSTICA
	if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_BASICA, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // POSIÇÃO NOS ESTOQUES (ANTIGO)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPosicaoEstoque.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_BASICA_CMVPV, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA_CMVPV, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // POSIÇÃO NOS ESTOQUES (CMV PV)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPosicaoEstoqueCmvPv.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO: ESTOQUE II" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelEstoque2Filtro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO: ESTOQUE DE VENDA" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelEstoqueVendaCmvPv.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO: RESUMO POSIÇÃO GERAL" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelEstoqueResumoGeralFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO: Estoque (E-Commerce)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelEstoqueEcommerceFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_PRODUTOS_PENDENTES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // PRODUTOS PENDENTES" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		f.pagina_destino.value='RelProdutosSemPresenca.asp';" & chr(13) & _
			"		f.titulo_relatorio.value='Produtos Pendentes';" & chr(13) & _
			"		f.filtro_fabricante_obrigatorio.value = '';" & chr(13) & _
			"		f.filtro_produto_obrigatorio.value = '';" & chr(13) & _
			"		f.filtro_apenas_produto_permitido.value = '';" & chr(13) & _
			"		s_dest='FiltroRelProdutoSemPresenca.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_COMPRAS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // COMPRAS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComprasFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_COMPRAS2, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // COMPRAS II" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelCompras2Filtro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_RESUMO_OPERACOES_ENTRE_ESTOQUES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE RESUMO DE OPERAÇÕES ENTRE ESTOQUES" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelResumoOperacoesEntreEstoques.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_AUDITORIA_ESTOQUE, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE AUDITORIA DO ESTOQUE" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelAuditoriaEstoque.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_REGISTROS_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_CEN_EDITA_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // CONSULTA REGISTROS DE ENTRADA" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='estoqueconsultamcrit.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_CONTAGEM_ESTOQUE, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // CONTAGEM DE ESTOQUE" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelContagemEstoque.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_PRODUTO_DEPOSITO_ZONA, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // ZONA DO PRODUTO (DEPÓSITO)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelProdutoDepositoZonaFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_SOLICITACAO_COLETAS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // SOLICITAÇÃO DE COLETAS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelSolicitacaoColetasFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_FAROL_CADASTRO_PRODUTO_COMPRADO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO PRODUTOS COMPRADOS (FAROL)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelFarolProdutoCompradoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_FAROL_RESUMIDO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO FAROL RESUMIDO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelFarolResumidoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_SINTETICO_CUBAGEM_VOLUME_PESO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO SINTÉTICO DE CUBAGEM, VOLUME E PESO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelCubagemVolumePesoSinteticoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_SINTETICO_CUBAGEM_VOLUME_PESO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO HISTÓRICO SINTÉTICO DE CUBAGEM, VOLUME E PESO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelCubagemVolumePesoSinteticoHistFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_IMPOSTOS_PAGOS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE IMPOSTOS PAGOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelImpostosPagosFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	
	if operacao_permitida(OP_CEN_REL_CONTROLE_IMPOSTOS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE CONTROLE DE IMPOSTOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelControleImpostosFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
	

	strScript = strScript & _
		"	if (s_dest=='') {" & chr(13) & _
		"		alert('Escolha um dos relatórios!!');" & chr(13) & _
		"		return false;" & chr(13) & _
		"		}" & chr(13) & _
		"" & chr(13) & _
		"	window.status = 'Aguarde ...';" & chr(13) & _
		"	f.action = s_dest;" & chr(13) & _
		"	f.submit();" & chr(13) & _
		"}" & chr(13) & _
		"" & chr(13) & _
		"</script>" & chr(13)
	
	Response.Write strScript
%>

<script type="text/javascript">
	function exibeJanelaCEP_Consulta() {
		$.mostraJanelaCEP(null);
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__E_JANELABUSCACEP_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
    .Largura {
        width: 280px;
    }
</style>

<body id="corpoPagina"
<% if strMensagemAvisoPopUp <> "" then %>
onload="alert('<%=strMensagemAvisoPopUp%>');"
<%end if%>
>

 <!-- PopUp Quadro de Avisos -->
       <form name="fAVISO" id="fAVISO" method="post" action="QuadroAvisoLido.asp">
        <input type="hidden" name="aviso_selecionado" id="aviso_selecionado" value=''>
        <input type="hidden" class="CBOX" name="xMsg" id="xMsg" value="">
        <div id="divTelaCheia" style="width:100%;height:100%;position:fixed;left:0;top:0;display:none;background-color:#000;opacity:0.3"></div>
    <div id="divQuadroAvisoPai" style="width:1000px;height:65%;overflow:visible;position:fixed;top:50%;left:50%;right:0;margin-top:-330px;margin-left:-500px; border:4px solid #000">
        <a href="javascript:fechaQuadroAviso(fAVISO, 0);" title="Fechar" style="font-size:40pt;font-weight:bolder;color:#555;position:relative;left:970px;top:-50px;margin:0;z-index:100">
            <img src="../IMAGEM/close_button_32.png" title="Fechar" style="border:0" />
        </a>
        <div id="divQuadroAviso" style="background-color:#fff;width:1000px;height:100%;overflow:scroll;position:absolute;top:0;left:0;right:0;bottom:0;margin:auto;border:1px solid #000;">
            <div id="divQuadroAvisoConteudo" style="position:relative;height:auto;width:650px;top:10px;left:0;right:0;margin:auto;z-index:200;padding:0;"></div>
            <div name='dREMOVE' id='dREMOVE'><a href="javascript:RemoveAviso(fAVISO, 1);">
		    <img src="../botao/remover.gif" width="176" height="55" border="0" style="position:relative;bottom:0px;right:0;left:0;margin:auto"/></a></div>
        </div>
    </div>
        </form>

<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">CENTRAL&nbsp;&nbsp;ADMINISTRATIVA<br>
	<span class='Cd'>Conectado desde: <%=formata_hora(Session("DataHoraLogon"))%>&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;duração: <%=formata_hora(Date+(Now-Session("DataHoraLogon")))%>
	<% if Trim(Session("SessionCtrlRecuperadoAuto")) <> "" then Response.Write " &nbsp;(*)"%>
	</span><br>
	<%	s = usuario_nome
		if s = "" then s = usuario
		s = x_saudacao & ", " & s
		s = "<span class='Cd'>" & s & "</span><br>"
	%>
	<%=s%>
	<span class="Rc">
	<% if blnPesquisaCEPAntiga then %>
		<span name="bPesqCep" id="bPesqCep" class="LPesqCep" onclick="AbrePesquisaCep();">Pesquisar CEP</span>&nbsp;&nbsp;&nbsp;
	<% end if %>
	<% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;nbsp;nbsp;" %>
	<% if blnPesquisaCEPNova then %>
		<span name="bPesqCep" id="bPesqCep" class="LPesqCep" onclick="exibeJanelaCEP_Consulta();">Pesquisar CEP</span>&nbsp;&nbsp;&nbsp;
	<% end if %>
		<a href="senha.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="altera a senha atual do usuário" class="LAlteraSenha">altera senha</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></span></td>
	</tr>
</table>

<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->

<% if strMensagemAviso <> "" then %>
	<br><br>
	<span class="Lbl">AVISO</span>
	<div class='MtAlerta' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><span style='margin:5px 2px 5px 2px;'><%=strMensagemAviso%></span></div>
	<br>
<% end if %>

<br>


<!--  ***********************************************************************************************  -->
<!--  C O N S U L T A S																				   -->
<!--  ***********************************************************************************************  -->
<% 
if (operacao_permitida(OP_CEN_CONSULTA_PEDIDO, s_lista_operacoes_permitidas) Or _
	operacao_permitida(OP_CEN_CONSULTA_ORCAMENTO, s_lista_operacoes_permitidas) Or _
	operacao_permitida(OP_CEN_CONSULTA_PEDIDOS_ANTERIORES_CLIENTE, s_lista_operacoes_permitidas)) then
%>
<span class="T">CONSULTAS</span>
<div class="QFn" align="CENTER">
<table>
	<% if operacao_permitida(OP_CEN_CONSULTA_PEDIDO, s_lista_operacoes_permitidas) then %>
	<!--  C O N S U L T A   P E D I D O  -->
	<tr class="DefaultBkg">
		<td>
			<p class="Cd">Nº Pedido</p>
		</td>
		<td>
			<form action="pedido.asp" method="post" id="fPED" name="fPED" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<table cellSpacing="0" width="100%">
				<tr>
					<td align="left" style='width:140px'><input maxlength="10" name="pedido_selecionado" id="pedido_selecionado" style="width:115px;margin-right:15px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {if (normaliza_numero_pedido_e_sufixo(this.value)!='') this.value=normaliza_numero_pedido_e_sufixo(this.value); fPED.submit();} filtra_pedido();" onblur="if (normaliza_numero_pedido_e_sufixo(this.value)!='') {this.value=normaliza_numero_pedido_e_sufixo(this.value);}"></td>
					<td align="center"><input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao" 
											value="CONSULTAR" title="consulta um pedido específico"></td>
					</tr>
			</table>
			</form>
		</td>
	</tr>
	<% end if %>
	<% if operacao_permitida(OP_CEN_CONSULTA_ORCAMENTO, s_lista_operacoes_permitidas) then %>
	<!--  C O N S U L T A   O R Ç A M E N T O  -->
	<tr class="DefaultBkg">
		<td>
			<p class="Cd">Nº Orçamento</p>
		</td>
		<td>
			<form action="orcamento.asp" method="post" id="fORC" name="fORC" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<table cellSpacing="0" width="100%">
				<tr>
					<td align="left" style='width:140px'><input maxlength="10" name="orcamento_selecionado" id="orcamento_selecionado" style="width:115px;margin-right:15px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {if (normaliza_num_orcamento(this.value)!='') this.value=normaliza_num_orcamento(this.value); fORC.submit();} filtra_orcamento();" onblur="if (normaliza_num_orcamento(this.value)!='') this.value=normaliza_num_orcamento(this.value);"></td>
					<td align="center"><input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao" 
											value="CONSULTAR" title="consulta um orçamento específico"></td>
					</tr>
			</table>
			</form>
		</td>
	</tr>
	<% end if %>
	<% if operacao_permitida(OP_CEN_CONSULTA_PEDIDO, s_lista_operacoes_permitidas) then %>
	<!--  C O N S U L T A   P E D I D O  P E L O   C A M P O   Nº Nota Fiscal  -->
	<tr class="DefaultBkg">
		<td>
			<p class="Cd">Nº Nota Fiscal</p>
		</td>
		<td>
			<form action="RelPesquisaPedidoNF.asp" method="post" id="fRNF" name="fRNF" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<table cellSpacing="0" width="100%">
				<tr>
					<td align="left" style='width:140px'><input maxlength="10" name="c_nf" id="c_nf" style="width:115px;margin-right:15px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fRNF.submit();"></td>
					<td align="center"><input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao" 
											value="CONSULTAR" title="consulta pedido pelo campo 'Nº Nota Fiscal'"></td>
					</tr>
			</table>
			</form>
		</td>
	</tr>
    <!--  C O N S U L T A   P E D I D O   P E L O   N Ú M E R O   M A G E N T O (BONSHOP) -->
	<tr class="DefaultBkg">
		<td width="40%" align="left">
			<p class="Cd">Nº Magento (Bonshop)</p>
		</td>
		<td align="left">
			<form action="RelPesquisaPedidoEcommerce.asp" method="post" id="fNumMagento" name="fNumMagento" style="margin:4px 0px 4px 0px;">
            <input type="hidden" name="c_tipo_num_pedido" id="c_tipo_num_pedido" value="<%=OP_PESQ_PEDIDO_MAGENTO_BONSHOP%>" />
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<table cellspacing="0" width="100%">
				<tr>
					<td align="left" style='width:140px'>
						<input maxlength="9" name="c_num_pedido_aux" id="c_num_pedido_aux" style="width:115px;margin-right:15px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fNumMagento.submit();">
					</td>
					<td align="left">
						<input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao" value="CONSULTAR" title="consulta pedido pelo campo 'Número Magento'">
					</td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
    <!--  C O N S U L T A   P E D I D O   P E L O   N Ú M E R O   M A G E N T O (ARCLUBE) -->
	<tr class="DefaultBkg">
		<td width="40%" align="left">
			<p class="Cd">Nº Magento (Arclube)</p>
		</td>
		<td align="left">
			<form action="RelPesquisaPedidoEcommerce.asp" method="post" id="fNumMagento" name="fNumMagento" style="margin:4px 0px 4px 0px;">
            <input type="hidden" name="c_tipo_num_pedido" id="c_tipo_num_pedido" value="<%=OP_PESQ_PEDIDO_MAGENTO_AR_CLUBE%>" />
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<table cellspacing="0" width="100%">
				<tr>
					<td align="left" style='width:140px'>
						<input maxlength="9" name="c_num_pedido_aux" id="c_num_pedido_aux" style="width:115px;margin-right:15px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fNumMagento.submit();">
					</td>
					<td align="left">
						<input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao" value="CONSULTAR" title="consulta pedido pelo campo 'Número Magento'">
					</td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
    <!--  C O N S U L T A   P E D I D O   P E L O   N Ú M E R O   M A R K E T P L A C E -->
	<tr class="DefaultBkg">
		<td width="40%" align="left">
			<p class="Cd">Nº Marketplace</p>
		</td>
		<td align="left">
			<form action="RelPesquisaPedidoEcommerce.asp" method="post" id="fNumMarketplace" name="fNumMarketplace" style="margin:4px 0px 4px 0px;">
            <input type="hidden" name="c_tipo_num_pedido" id="c_tipo_num_pedido" value="<%=OP_PESQ_PEDIDO_MARKETPLACE_AR_CLUBE%>" />
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<table cellspacing="0" width="100%">
				<tr>
					<td align="left" style='width:140px'>
						<input maxlength="20" name="c_num_pedido_aux" id="c_num_pedido_aux" style="width:115px;margin-right:15px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fNumMarketplace.submit();">
					</td>
					<td align="left">
						<input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao" value="CONSULTAR" title="consulta pedido pelo campo 'Número Marketplace'">
					</td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
	<% end if %>
    
	<% if operacao_permitida(OP_CEN_CONSULTA_ORDEM_SERVICO, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_EDITA_ORDEM_SERVICO, s_lista_operacoes_permitidas) then %>
	<!--  CONSULTA ORDEM DE SERVIÇO PELO Nº DE SÉRIE DO PRODUTO  -->
	<tr class="DefaultBkg">
		<td>
			<p class="Cd">Ordem de Serviço (nº série)</p>
		</td>
		<td>
			<form action="OrdemServicoPesqNumSerie.asp" method="post" id="fOSNumSerie" name="fOSNumSerie" onsubmit="if (!tem_info(fOSNumSerie.c_num_serie.value)) {alert('Informe o nº de série!!'); fOSNumSerie.c_num_serie.focus(); return false;}" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<table cellSpacing="0" width="100%">
				<tr>
					<td align="left" style='width:140px'><input maxlength="20" name="c_num_serie" id="c_num_serie" style="width:115px;margin-right:15px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fOSNumSerie.submit();"></td>
					<td align="center"><input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao"
											value="CONSULTAR" title="consulta ordem de serviço pelo número de série"></td>
					</tr>
			</table>
			</form>
		</td>
	</tr>
	<% end if %>
	<% if operacao_permitida(OP_CEN_CONSULTA_ORDEM_SERVICO, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_EDITA_ORDEM_SERVICO, s_lista_operacoes_permitidas) then %>
	<!--  CONSULTA ORDEM DE SERVIÇO PELO Nº ORDEM DE SERVIÇO  -->
	<tr class="DefaultBkg">
		<td>
			<p class="Cd">Ordem de Serviço (nº OS)</p>
		</td>
		<td>
			<form action="OrdemServico.asp" method="post" id="fOS" name="fOS" onsubmit="if (!tem_info(fOS.num_OS.value)) {alert('Informe o nº da O.S.!!'); fOS.num_OS.focus(); return false;}" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<table cellSpacing="0" width="100%">
				<tr>
					<td align="left" style='width:140px'><input maxlength="12" name="num_OS" id="num_OS" style="width:115px;margin-right:15px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fOS.submit(); filtra_numerico();"></td>
					<td align="center"><input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao" 
											value="CONSULTAR" title="consulta ordem de serviço pelo número da OS"></td>
					</tr>
			</table>
			</form>
		</td>
	</tr>
	<% end if %>
	<% if operacao_permitida(OP_CEN_CONSULTA_PEDIDOS_ANTERIORES_CLIENTE, s_lista_operacoes_permitidas) then %>
	<!--  C O N S U L T A   C L I E N T E  -->
	<tr class="DefaultBkg">
		<td>
			<p class="Cd" style='cursor: pointer;' onclick="fRPA.bCONSULTAR.click();">Cliente</p>
		</td>
		<td>
			<form action="RelPedidosAnteriores.asp" method="post" id="fRPA" name="fRPA" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<table cellSpacing="0" width="100%">
				<tr>
					<td style='width:140px' align='center'><p class="Cn" style='margin-right:15px;'>. . . . . . . . . . . . . . . . . . . .</p></td>
					<td align="center"><input name="bCONSULTAR" id="bCONSULTAR" type="submit" class="Botao" 
											value="CONSULTAR" title="consulta pedidos anteriormente efetuados por um cliente"></td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
	<% end if %>
</table>
</div>
<br />
<% end if %>



<!--  ***********************************************************************************************  -->
<!--  P A G A M E N T O S																			   -->
<!--  ***********************************************************************************************  -->
<% 
if (operacao_permitida(OP_CEN_PAGTO_PARCIAL, s_lista_operacoes_permitidas) Or _
	operacao_permitida(OP_CEN_PAGTO_QUITACAO, s_lista_operacoes_permitidas)) then
%>
<form method="post" id="fPAGTO" name="fPAGTO">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span class="T">PAGAMENTOS</span>
<div class="QFn" align="center">
<% if operacao_permitida(OP_CEN_PAGTO_PARCIAL, s_lista_operacoes_permitidas) then %>
<table class="TFn" style="margin-bottom:5px;">
	<tr>
		<td nowrap>
			<table>
				<tr><td colspan="5" class="MT" style="border:0pt;" valign="middle" align="center" NOWRAP>
					<span class="PLTe" style="font-size:10pt;color:black;">PAGAMENTO PARCIAL</span></td></tr>
				<tr>
					<td nowrap>
					<span class="PLTe" style="vertical-align:middle;">Nº Pedido</span>
					<input maxlength="10" name="c_pedido" id="c_pedido" style="width:90px;" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPAGTO.c_valor.focus(); filtra_pedido();" onblur="if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value);"></td>
					<td><span style="width:10px;">&nbsp;</span></td>
					<td nowrap align="right"><span class="PLTe" style="vertical-align:middle;">Valor</span>
					<input maxlength="12" name="c_valor" id="c_valor" style="width:120px;text-align:right;" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) bPAGTOPARCIAL.click(); filtra_moeda();" onblur="this.value=formata_moeda(this.value);"></td>
					<td><span style="width:10px;">&nbsp;</span></td>
					<td nowrap>
					<input name="bPAGTOPARCIAL" id="bPAGTOPARCIAL" type="button" class="Botao" onclick="if (fPAGTOPARCIALConcluir(fPAGTO)) fPAGTO.submit();" value="EXECUTAR" title="executa"></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<!-- ************   SEPARADOR   ************ -->
<table width="100%" cellPadding="0" CellSpacing="0" style="border-bottom:1px solid #C0C0C0;"><tr><td><span></span></td></tr></table>
<table class="TFn" style="margin-bottom:5px;">
	<tr>
		<td nowrap>
			<table>
				<tr><td colspan="3" class="MT" style="border:0pt;" valign="middle" align="center" NOWRAP>
					<span class="PLTe" style="font-size:10pt;color:black;">PAGAMENTO PARCIAL (MARKETPLACE)</span></td></tr>
				<tr>
					<td nowrap>
						<span class="PLTe" style="vertical-align:middle;">Nº Pedido(s)</span>
						<br />
						<textarea rows="6" name="c_pedidos_pagto_parcial_marketplace" id="c_pedidos_pagto_parcial_marketplace" style="width:200px;"></textarea>
					</td>
					<td><span style="width:10px;">&nbsp;</span></td>
					<td nowrap align="right">
						<span class="PLTe" style="vertical-align:middle;">Valor</span>
						<br />
						<textarea rows="6" name="c_valor_pagto_parcial_marketplace" id="c_valor_pagto_parcial_marketplace" style="width:120px;text-align:right;"></textarea>
					</td>
				</tr>
				<tr>
					<td colspan="3" align="center" nowrap>
					<input name="bPAGTOPARCIALMKTP" id="bPAGTOPARCIALMKTP" type="button" class="Botao" onclick="if (fPagtoParcialMarketplaceConcluir(fPAGTO)) fPAGTO.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<% end if %>

<% if operacao_permitida(OP_CEN_PAGTO_PARCIAL, s_lista_operacoes_permitidas) And _
	  operacao_permitida(OP_CEN_PAGTO_QUITACAO, s_lista_operacoes_permitidas) then %>
<!-- ************   SEPARADOR   ************ -->
<table width="100%" cellPadding="0" CellSpacing="0" style="border-bottom:1px solid #C0C0C0;"><tr><td><span></span></td></tr></table>
<% end if %>

<% if operacao_permitida(OP_CEN_PAGTO_QUITACAO, s_lista_operacoes_permitidas) then %>
<table class="TFn" style="margin-top:5px;">
	<tr>
		<td align="center" nowrap>
			<table>
				<tr><td class="MT" style="border:0pt;" valign="middle" align="center" nowrap>
					<span class="PLTc" style="font-size:10pt;color:black;">QUITAÇÃO</span></td></tr>
				<tr>
					<td align="center" nowrap>
					<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
					<br>
					<textarea rows="6" name="c_pedidos" id="c_pedidos" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
					</td></tr>
				<tr>
					<td align="center">
					<input name="bPAGTOQUITACAO" id="bPAGTOQUITACAO" type="button" class="Botao" onclick="if (fQUITACAOConcluir(fPAGTO)) fPAGTO.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<!-- ************   SEPARADOR   ************ -->
<table width="100%" cellPadding="0" CellSpacing="0" style="border-bottom:1px solid #C0C0C0;"><tr><td><span></span></td></tr></table>
<table class="TFn" style="margin-top:5px;">
	<tr>
		<td align="center" nowrap>
			<table>
				<tr><td class="MT" style="border:0pt;" valign="middle" align="center" nowrap>
					<span class="PLTc" style="font-size:10pt;color:black;">QUITAÇÃO (MARKETPLACE)</span></td></tr>
				<tr>
					<td align="center" nowrap>
					<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
					<br>
					<textarea rows="6" name="c_pedidos_quitacao_marketplace" id="c_pedidos_quitacao_marketplace" style="width:200px;"></textarea>
					</td></tr>
				<tr>
					<td align="center">
					<input name="bPAGTOQUITACAOMKTP" id="bPAGTOQUITACAOMKTP" type="button" class="Botao" onclick="if (fQuitacaoMarketplaceConcluir(fPAGTO)) fPAGTO.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<% end if %>

</div>
</form>
<br />
<% end if %>


<!--  ***********************************************************************************************  -->
<!--  S E P A R A Ç Ã O   D E   M E R C A D O R I A S              									   -->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_SEPARACAO_MERCADORIAS, s_lista_operacoes_permitidas) then %>
<form method="post" id="fSEP" name="fSEP">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_qtde_nfe_emitente" id="c_qtde_nfe_emitente" value="<%=Cstr(qtde_nfe_emitente)%>" />
<% if qtde_nfe_emitente = 1 then %>
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=Cstr(id_nfe_emitente_selecionado)%>" />
<% end if %>
<span class="T">SEPARAÇÃO DE MERCADORIAS</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td nowrap>
			<table>
				<% if qtde_nfe_emitente > 1 then %>
				<tr>
					<td align="left">
					<table cellspacing="0" cellpadding="0">
					<tr>
						<td align="left" nowrap>
							<span class="PLTe">CD</span>
						</td>
					</tr>
					<tr>
						<td align="left">
							<table style="margin: 0px 0px 0px 0px;" cellspacing="0" cellpadding="0">
								<tr>
								<td align="left">
									<select id="c_nfe_emitente" name="c_nfe_emitente" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" style="margin-left:5px;margin-top:4pt; margin-bottom:4pt;">
										<%=wms_usuario_x_nfe_emitente_monta_itens_select(usuario, "")%>
									</select>
								</td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
					</td>
				</tr>
				<tr style="height:4px;"><td></td></tr>
				<% end if %>
				<tr>
					<td nowrap>
					<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
					<br>
					<textarea rows="6" name="c_pedidos" id="c_pedidos" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
					</td>
				</tr>
				<tr>
					<td align="center">
					<input name="bEXECUTA" id="bEXECUTA" type="button" class="Botao" onclick="if (fSEPConcluir(fSEP)) fSEP.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td>
					<span class="PLTe" style="vertical-align:top;">NSU do Relatório</span>
					<br>
					<input maxlength="12" name="c_nsu_rel_separacao_zona" id="c_nsu_rel_separacao_zona" style="width:115px;text-align:center;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {if (fSEPUsandoRelConcluir(fSEP)) fSEP.submit();} filtra_numerico();">
					</td>
				</tr>
				<tr>
					<td align="center">
					<input name="bExecutaSepUsandoRel" id="bExecutaSepUsandoRel" type="button" class="Botao" onclick="if (fSEPUsandoRelConcluir(fSEP)) fSEP.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
		</tr>
	</table>
</div>
</form>
<br />
<% end if %>


<!--  ***********************************************************************************************  -->
<!--  E N T R E G A   D E   M E R C A D O R I A S              										   -->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_ENTREGA_MERCADORIAS, s_lista_operacoes_permitidas) then %>
<form method="post" id="fETG" name="fETG">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_qtde_nfe_emitente" id="c_qtde_nfe_emitente" value="<%=Cstr(qtde_nfe_emitente)%>" />
<% if qtde_nfe_emitente = 1 then %>
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=Cstr(id_nfe_emitente_selecionado)%>" />
<% end if %>
<span class="T">ENTREGA DE MERCADORIAS</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td nowrap>
			<table>
				<% if qtde_nfe_emitente > 1 then %>
				<tr>
					<td align="left">
					<table cellspacing="0" cellpadding="0">
					<tr>
						<td align="left" nowrap>
							<span class="PLTe">CD</span>
						</td>
					</tr>
					<tr>
						<td align="left">
							<table style="margin: 0px 0px 0px 0px;" cellspacing="0" cellpadding="0">
								<tr>
								<td align="left">
									<select id="c_nfe_emitente" name="c_nfe_emitente" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" style="margin-left:5px;margin-top:4pt; margin-bottom:4pt;">
										<%=wms_usuario_x_nfe_emitente_monta_itens_select(usuario, "")%>
									</select>
								</td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
					</td>
				</tr>
				<tr style="height:4px;"><td></td></tr>
				<% end if %>
				<tr>
					<td nowrap>
					<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
					<br>
					<textarea rows="6" name="c_pedidos" id="c_pedidos" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
					</td></tr>
				<tr>
					<td align="center">
					<input name="bEXECUTA" id="bEXECUTA" type="button" class="Botao" onclick="if (fETGConcluir(fETG)) fETG.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td>
					<span class="PLTe" style="vertical-align:top;">NSU do Romaneio</span>
					<br>
					<input maxlength="12" name="c_nsu_romaneio" id="c_nsu_romaneio" style="width:115px;text-align:center;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {if (fETGRomaneioConcluir(fETG)) fETG.submit();} filtra_numerico();">
					</td>
				</tr>
				<tr>
					<td align="center">
					<input name="bExecutaEtgRomaneio" id="bExecutaEtgRomaneio" type="button" class="Botao" onclick="if (fETGRomaneioConcluir(fETG)) fETG.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
		</tr>
	</table>
</div>
</form>
<br />
<%end if%>


<!--  ***********************************************************************************************  -->
<!--  A N O T A Ç Ã O   D A   D A T A   "E N T R E G A   M A R C A D A   P A R A"					   -->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_AGENDAMENTO_ENTREGA, s_lista_operacoes_permitidas) then %>
<form method="post" id="fDTETG" name="fDTETG">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_qtde_nfe_emitente" id="c_qtde_nfe_emitente" value="<%=Cstr(qtde_nfe_emitente)%>" />
<% if qtde_nfe_emitente = 1 then %>
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=Cstr(id_nfe_emitente_selecionado)%>" />
<% end if %>
<span class="T">DATA DE COLETA</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td nowrap>
			<table>
				<% if qtde_nfe_emitente > 1 then %>
				<tr>
					<td align="left">
					<table cellspacing="0" cellpadding="0">
					<tr>
						<td align="left" nowrap>
							<span class="PLTe">CD</span>
						</td>
					</tr>
					<tr>
						<td align="left">
							<table style="margin: 0px 0px 0px 0px;" cellspacing="0" cellpadding="0">
								<tr>
								<td align="left">
									<select id="c_nfe_emitente" name="c_nfe_emitente" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" style="margin-left:5px;margin-top:4pt; margin-bottom:4pt;">
										<%=wms_usuario_x_nfe_emitente_monta_itens_select(usuario, "")%>
									</select>
								</td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
					</td>
				</tr>
				<tr style="height:4px;"><td></td></tr>
				<% end if %>
				<tr>
					<td align="left" nowrap>
						<span class="PLTe" style="vertical-align:top;">Data de Coleta</span>
						<br>
						<input class="Cc" maxlength="10" style="width:90px;" name="c_dt_entrega" id="c_dt_entrega" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fDTETG.c_pedidos.focus(); filtra_data();">
					</td>
				</tr>
				<tr>
					<td align="left" nowrap>
						<table border="0">
						<tr>
							<td align="left">
								<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
								<br>
								<textarea rows="6" name="c_pedidos" id="c_pedidos" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
							</td>
						</tr>
						<tr>
							<td align="center">
								<input name="bEXECUTA" id="bEXECUTA" type="button" class="Botao" onclick="if (fDTETGConcluir(fDTETG)) fDTETG.submit();" value="EXECUTAR" title="executa">
							</td>
						</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
		</tr>
	</table>
</div>
</form>
<br />
<%end if%>


<!--  ***********************************************************************************************  -->
<!--  P E D I D O S   R E C E B I D O S   P E L O   C L I E N T E									   -->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_ANOTA_PEDIDO_RECEBIDO, s_lista_operacoes_permitidas) then %>
<form method="post" id="fPedRec" name="fPedRec">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_qtde_nfe_emitente" id="c_qtde_nfe_emitente" value="<%=Cstr(qtde_nfe_emitente)%>" />
<% if qtde_nfe_emitente = 1 then %>
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=Cstr(id_nfe_emitente_selecionado)%>" />
<% end if %>
<span class="T">PEDIDOS RECEBIDOS PELO CLIENTE</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td align="center" nowrap>
			<table>
				<% if qtde_nfe_emitente > 1 then %>
				<tr>
					<td align="left">
					<table cellspacing="0" cellpadding="0">
					<tr>
						<td align="left" nowrap>
							<span class="PLTe">CD</span>
						</td>
					</tr>
					<tr>
						<td align="left">
							<table style="margin: 0px 0px 0px 0px;" cellspacing="0" cellpadding="0">
								<tr>
								<td align="left">
									<select id="c_nfe_emitente" name="c_nfe_emitente" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" style="margin-left:5px;margin-top:4pt; margin-bottom:4pt;">
										<%=wms_usuario_x_nfe_emitente_monta_itens_select(usuario, "")%>
									</select>
								</td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
					</td>
				</tr>
				<tr style="height:4px;"><td></td></tr>
				<% end if %>
				<tr>
					<td align="top" nowrap>
					<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
					<br>
					<textarea rows="6" name="c_pedidos" id="c_pedidos" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
					</td>
					<td style="width:15px;"></td>
					<td align="top" nowrap>
					<span class="PLTe" style="vertical-align:top;">Obs II</span>
					<br>
					<textarea rows="6" name="c_obs2" id="c_obs2" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_numerico();" onblur="this.value=trim(this.value);"></textarea>
					</td>
				</tr>
			</table>
			<table>
				<tr>
					<td colspan="2" align="left">
						<input type="checkbox" tabindex="-1" id="ckb_entrega" name="ckb_entrega"
							value="ON"><span class="C" style="cursor:default" 
							onclick="fPedRec.ckb_entrega.click();">Entrega</span>
					</td>
				</tr>
				<tr>
					<td align="left">
						<input type="checkbox" tabindex="-1" id="ckb_recebido" name="ckb_recebido"
							value="ON"><span class="C" style="cursor:default" 
							onclick="fPedRec.ckb_recebido.click();">Recebido em</span>
					</td>
					<td align="left">
						<input class="Cc" maxlength="10" style="width:90px;" name="c_dt_recebido" id="c_dt_recebido" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();} else {if (tem_info(this.value)) if (retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(this.value)) > retorna_so_digitos(formata_ddmmyyyy_yyyymmdd('<%=formata_data(Date)%>'))) {alert('Data inválida: deve ser menor ou igual a hoje!'); this.focus();}}" onkeypress="if (digitou_enter(true)) fPedRec.bEXECUTA.focus(); filtra_data();" onkeyup="if (trim(this.value)!='') fPedRec.ckb_recebido.checked=true;" onchange="if (trim(this.value)!='') fPedRec.ckb_recebido.checked=true;">
					</td>
				</tr>
			</table>
			<table>
				<tr>
					<td>
						<select id="c_transportadora" name="c_transportadora" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
						<% =transportadora_monta_itens_select(Null) %>
						</select>
					</td>
				</tr>
			</table>
			<table>
				<tr>
					<td align="center">
					<input name="bEXECUTA" id="bEXECUTA" type="button" class="Botao" onclick="if (fPedRecConcluir(fPedRec)) fPedRec.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
		</tr>
	</table>
</div>
</form>
<br />
<%end if%>


<!--  ***********************************************************************************************  -->
<!--  A N O T A Ç Ã O   D A   T R A N S P O R T A D O R A   N O   P E D I D O						   -->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_ANOTA_TRANSPORTADORA_NO_PEDIDO, s_lista_operacoes_permitidas) then %>
<form method="post" id="fAnotaTransp" name="fAnotaTransp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_qtde_nfe_emitente" id="c_qtde_nfe_emitente" value="<%=Cstr(qtde_nfe_emitente)%>" />
<% if qtde_nfe_emitente = 1 then %>
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=Cstr(id_nfe_emitente_selecionado)%>" />
<% end if %>
<span class="T">ANOTAR TRANSPORTADORA NO PEDIDO</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td nowrap>
			<table>
				<% if qtde_nfe_emitente > 1 then %>
				<tr>
					<td align="left">
					<table cellspacing="0" cellpadding="0">
					<tr>
						<td align="left" nowrap>
							<span class="PLTe">CD</span>
						</td>
					</tr>
					<tr>
						<td align="left">
							<table style="margin: 0px 0px 0px 0px;" cellspacing="0" cellpadding="0">
								<tr>
								<td align="left">
									<select id="c_nfe_emitente" name="c_nfe_emitente" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" style="margin-left:5px;margin-top:4pt; margin-bottom:4pt;">
										<%=wms_usuario_x_nfe_emitente_monta_itens_select(usuario, "")%>
									</select>
								</td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
					</td>
				</tr>
				<tr style="height:4px;"><td></td></tr>
				<% end if %>
				<tr>
					<td nowrap align="left">
					<span class="PLTe" style="vertical-align:top;">Transportadora</span>
					<br>
					<input class="C" maxlength="10" style="width:120px;" name="c_transportadora" id="c_transportadora" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) fAnotaTransp.c_pedidos.focus();">
					</td></tr>
				<tr>
					<td nowrap>
					<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
					<br>
					<textarea rows="6" name="c_pedidos" id="c_pedidos" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
					</td></tr>
				<tr>
					<td align="center">
					<input name="bEXECUTA" id="bEXECUTA" type="button" class="Botao" onclick="if (fAnotaTranspConcluir(fAnotaTransp)) fAnotaTransp.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
		</tr>
	</table>
</div>
</form>
<br />
<%end if%>


<!--  ***********************************************************************************************  -->
<!--  C O M I S S Ã O   P A G A / N Ã O - P A G A 													   -->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_FLAG_COMISSAO_PAGA, s_lista_operacoes_permitidas) then %>
<form method="post" id="fCOMISSAO" name="fCOMISSAO">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span class="T">COMISSÃO</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td nowrap>
			<table>
				<tr><td nowrap align="center">
					<table>
						<tr><td>
							<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
							<br>
							<textarea rows="6" name="c_pedidos" id="c_pedidos" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
						</td></tr>
					</table>
				</td></tr>
				<tr><td nowrap align="center">
					<table cellspacing='0' cellpadding='0'>
						<tr><td align="left">
							<% intIdx = 0 %>
							<input type="radio" id="rb_comissao_paga" name="rb_comissao_paga" value="S"><span class="C" style="cursor:default" onclick="fCOMISSAO.rb_comissao_paga[<%=Cstr(intIdx)%>].click();">Paga</span>
						</td></tr>
						<tr><td align="left">
							<% intIdx = intIdx+1 %>
							<input type="radio" id="rb_comissao_paga" name="rb_comissao_paga" value="N"><span class="C" style="cursor:default" onclick="fCOMISSAO.rb_comissao_paga[<%=Cstr(intIdx)%>].click();">Não-Paga</span>
						</td></tr>
					</table>
				
				</td></tr>
				<tr>
					<td align="center">
					<input name="bEXECUTA" id="bEXECUTA" type="button" class="Botao" onclick="if (fCOMISSAOConcluir(fCOMISSAO)) fCOMISSAO.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
		</tr>
	</table>
</div>
</form>
<br />
<%end if%>


<!--  ***********************************************************************************************  -->
<!--  C O M I S S Ã O   D E S C O N T A D A / N Ã O - D E S C O N T A D A								-->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_FLAG_COMISSAO_PAGA, s_lista_operacoes_permitidas) then %>
<form method="post" id="fCOMISSAODESC" name="fCOMISSAODESC">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span class="T">COMISSÃO (DESCONTOS)</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td nowrap>
			<table>
				<tr><td nowrap align="center">
					<table>
						<tr><td>
							<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
							<br>
							<textarea rows="6" name="c_pedidos" id="c_pedidos" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
						</td></tr>
					</table>
				</td></tr>
				<tr><td nowrap align="center">
					<table cellspacing='0' cellpadding='0'>
						<tr><td align="left">
							<% intIdx = 0 %>
							<input type="radio" id="rb_comissao_descontada" name="rb_comissao_descontada" value="S"><span class="C" style="cursor:default" onclick="fCOMISSAODESC.rb_comissao_descontada[<%=Cstr(intIdx)%>].click();">Descontada</span>
						</td></tr>
						<tr><td align="left">
							<% intIdx = intIdx+1 %>
							<input type="radio" id="rb_comissao_descontada" name="rb_comissao_descontada" value="N"><span class="C" style="cursor:default" onclick="fCOMISSAODESC.rb_comissao_descontada[<%=Cstr(intIdx)%>].click();">Não-Descontada</span>
						</td></tr>
					</table>
					
				</td></tr>
				<tr>
					<td align="center">
					<input name="bEXECUTA" id="bEXECUTA" type="button" class="Botao" onclick="if (fCOMISSAODESCConcluir(fCOMISSAODESC)) fCOMISSAODESC.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
		</tr>
	</table>
</div>
</form>
<br />
<%end if%>


<!--  ***********************************************************************************************  -->
<!--  "CRÉDITO OK (AGUARDANDO DEPÓSITO)" => "CRÉDITO OK"											   -->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_EDITA_ANALISE_CREDITO, s_lista_operacoes_permitidas) then %>
<form method="post" id="fCrOk" name="fCrOk">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span class="T">CRÉDITO OK (AGUARDANDO DEPÓSITO)</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td nowrap>
			<table>
				<tr>
					<td nowrap>
					<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
					<br>
					<textarea rows="6" name="c_pedidos" id="c_pedidos" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
					</td></tr>
				<tr>
					<td align="center">
					<input name="bEXECUTA" id="bEXECUTA" type="button" class="Botao" onclick="if (fCrOkConcluir(fCrOk)) fCrOk.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
		</tr>
	</table>
</div>
</form>
<br />
<% end if %>


<!--  ***********************************************************************************************  -->
<!--  "CRÉDITO OK (DEPÓSITO AGUARDANDO DESBLOQUEIO)" => "CRÉDITO OK"								   -->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_EDITA_ANALISE_CREDITO, s_lista_operacoes_permitidas) then %>
<form method="post" id="fDepDesbloqCrOk" name="fDepDesbloqCrOk">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span class="T">CRÉDITO OK<br />(DEPÓSITO AGUARDANDO DESBLOQUEIO)</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td nowrap>
			<table>
				<tr>
					<td nowrap>
					<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
					<br>
					<textarea rows="6" name="c_pedidos" id="c_pedidos" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
					</td>
				</tr>
				<tr>
					<td align="center">
					<input name="bDepDesbloqExecuta" id="bDepDesbloqExecuta" type="button" class="Botao" onclick="if (fDepDesbloqCrOkConcluir(fDepDesbloqCrOk)) fDepDesbloqCrOk.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
		</tr>
	</table>
</div>
</form>
<br />
<% end if %>


<!--  ***********************************************************************************************  -->
<!--  "PENDENTE VENDAS" => "CRÉDITO OK"								   -->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_EDITA_ANALISE_CREDITO, s_lista_operacoes_permitidas) then %>
<form method="post" id="fPendVendasCrOk" name="fPendVendasCrOk">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span class="T">CRÉDITO OK<br />(PENDENTE VENDAS)</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td nowrap>
			<table>
				<tr>
					<td nowrap>
					<span class="PLTe" style="vertical-align:top;">Nº Pedido(s)</span>
					<br>
					<textarea rows="6" name="c_pedidos" id="c_pedidos" style="width:110px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
					</td>
				</tr>
				<tr>
					<td align="center">
					<input name="bPendVendasExecuta" id="bPendVendasExecuta" type="button" class="Botao" onclick="if (fPendVendasCrOkConcluir(fPendVendasCrOk)) fPendVendasCrOk.submit();" value="EXECUTAR" title="executa">
					</td>
				</tr>
			</table>
		</td>
		</tr>
	</table>
</div>
</form>
<br />
<% end if %>


<!--  ***********************************************************************************************  -->
<!--  R E L A T Ó R I O S																		       -->
<!--  ***********************************************************************************************  -->
<%
	qtde_rel_glb = 0
	qtde_rel_com = 0
	qtde_rel_adm = 0
	qtde_rel_compras_logist = 0
	qtde_total_rel = 0
'	RELATÓRIOS GLOBAIS
	if operacao_permitida(OP_CEN_REL_MULTICRITERIO_PEDIDOS_ANALITICO, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_CEN_REL_MULTICRITERIO_PEDIDOS_SINTETICO, s_lista_operacoes_permitidas) then
		qtde_rel_glb=qtde_rel_glb+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_MULTICRITERIO_ORCAMENTOS, s_lista_operacoes_permitidas) then
		qtde_rel_glb=qtde_rel_glb+1
		qtde_total_rel=qtde_total_rel+1
		end if
'	RELATÓRIOS COMERCIAIS
	if operacao_permitida(OP_CEN_REL_PRODUTOS_VENDIDOS, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_PEDIDOS_COLOCADOS_NO_MES, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_VENDAS_COM_DESCONTO_SUPERIOR, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_COMISSAO_VENDEDORES, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_SINTETICO, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_ANALITICO, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
	'	RELATÓRIO PEDIDOS INDICADORES (VIA NFS-e)
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
	'	RELATÓRIO PEDIDOS INDICADORES (VIA NFS-e) (CONSULTA)
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
   ' if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
	'	comissão indicadores cadastra nf
	'	qtde_rel_com=qtde_rel_com+1
	'	qtde_total_rel=qtde_total_rel+1
	'	end if
    if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
	'	RELATÓRIO INDICADORES PAG
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if

    if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
	'	RELATÓRIO INDICADORES CONSULTA POR PEDIDO E INDICADOR
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if

        if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
	'	RELATÓRIO INDICADORES CONSULTA POR BANCO
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if

	if operacao_permitida(OP_CEN_REL_FATURAMENTO_VENDEDORES_EXT, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_COMISSAO_LOJA_POR_INDICACAO, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_ANALISE_PEDIDOS, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_DIVERGENCIA_CLIENTE_INDICADOR, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_METAS_INDICADOR, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_PERFORMANCE_INDICADOR, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_VENDAS_POR_BOLETO, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_ECOMMERCE_EXPORTACAO, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_DADOS_TABELA_DINAMICA, s_lista_operacoes_permitidas) then
		qtde_rel_com=qtde_rel_com+1
		qtde_total_rel=qtde_total_rel+1
		end if
'	RELATÓRIOS ADMINISTRATIVOS
    if operacao_permitida(OP_CEN_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) then
        qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
        end if
    if operacao_permitida(OP_CEN_PRE_DEVOLUCAO_RECEBIMENTO_MERCADORIA, s_lista_operacoes_permitidas) then
        qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
        end if
	if operacao_permitida(OP_CEN_REL_SEPARACAO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_SEPARACAO_ZONA, s_lista_operacoes_permitidas) then
	'	Separação (Zona)
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
	'	Separação (Zona) - Consulta
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_PRODUTOS_ESTOQUE_DEVOLUCAO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_DEVOLUCAO_PRODUTOS, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_DEVOLUCAO_PRODUTOS2, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_PRODUTOS_SPLIT_POSSIVEL, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_ESTOQUE_VENDA_CRITICO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_MEIO_DIVULGACAO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_LOG_ESTOQUE, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_FATURAMENTO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_FATURAMENTO_CMVPV, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_FATURAMENTO2, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_FATURAMENTO3, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_VENDAS, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_GERENCIAL_DE_VENDAS, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_ROMANEIO_ENTREGA, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_ANALISE_CREDITO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_PESQUISA_INDICADORES, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_INDICADOR_SEM_AVALIACAO_DESEMPENHO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_CHECAGEM_NOVOS_PARCEIROS, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_FRETE_ANALITICO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_FRETE_SINTETICO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_PEDIDO_NAO_RECEBIDO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_ESTATISTICAS_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
    if operacao_permitida(OP_CEN_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
    if operacao_permitida(OP_CEN_REL_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
    if operacao_permitida(OP_CEN_REL_ESTATISTICAS_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_PESQUISA_ORDEM_SERVICO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_CLIENTE_SPC, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_TRANSACOES_CIELO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_TRANSACOES_CIELO_ANDAMENTO, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_BRASPAG_TRANSACOES, s_lista_operacoes_permitidas) then
	'	Relatório de Transações Braspag
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_BRASPAG_AF_REVIEW, s_lista_operacoes_permitidas) then
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_BRASPAG_TRANSACOES, s_lista_operacoes_permitidas) then
	'	Relatório de Transações Braspag/Clearsale
		qtde_rel_adm=qtde_rel_adm+1
		qtde_total_rel=qtde_total_rel+1
		end if
'	RELATÓRIOS DE COMPRAS E LOGÍSTICA
	if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_BASICA, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_BASICA_CMVPV, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA_CMVPV, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
    if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
    if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
    if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_PRODUTOS_PENDENTES, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_COMPRAS, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_COMPRAS2, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_RESUMO_OPERACOES_ENTRE_ESTOQUES, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_AUDITORIA_ESTOQUE, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_REGISTROS_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_CEN_EDITA_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_CONTAGEM_ESTOQUE, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_PRODUTO_DEPOSITO_ZONA, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_SOLICITACAO_COLETAS, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_FAROL_CADASTRO_PRODUTO_COMPRADO, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_FAROL_RESUMIDO, s_lista_operacoes_permitidas) then
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_SINTETICO_CUBAGEM_VOLUME_PESO, s_lista_operacoes_permitidas) then
	'	RELATÓRIO DE CUBAGEM, VOLUME E PESO (SINTÉTICO)
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if
	if operacao_permitida(OP_CEN_REL_SINTETICO_CUBAGEM_VOLUME_PESO, s_lista_operacoes_permitidas) then
	'	RELATÓRIO HISTÓRICO DE CUBAGEM, VOLUME E PESO (SINTÉTICO)
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if

	if operacao_permitida(OP_CEN_REL_IMPOSTOS_PAGOS, s_lista_operacoes_permitidas) then
	'	RELATÓRIO DE IMPOSTOS PAGOS
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if

	if operacao_permitida(OP_CEN_REL_CONTROLE_IMPOSTOS, s_lista_operacoes_permitidas) then
	'	RELATÓRIO DE CONTROLE DE IMPOSTOS
		qtde_rel_compras_logist=qtde_rel_compras_logist+1
		qtde_total_rel=qtde_total_rel+1
		end if

%>

<% if qtde_total_rel > 0 then %>
<form method="post" id="fREL" name="fREL" onsubmit="if (!fRELConcluir(fREL)) return false;">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pagina_destino" id="pagina_destino" value=''>
<input type="hidden" name="titulo_relatorio" id="titulo_relatorio" value=''>
<input type="hidden" name="filtro_obrigatorio" id="filtro_obrigatorio" value=''>
<input type="hidden" name="filtro_obrigatorio_data_inicio" id="filtro_obrigatorio_data_inicio" value=''>
<input type="hidden" name="filtro_obrigatorio_data_termino" id="filtro_obrigatorio_data_termino" value=''>
<input type="hidden" name="filtro_fabricante_obrigatorio" id="filtro_fabricante_obrigatorio" value=''>
<input type="hidden" name="filtro_produto_obrigatorio" id="filtro_produto_obrigatorio" value=''>
<input type="hidden" name="filtro_apenas_produto_permitido" id="filtro_apenas_produto_permitido" value=''>
<!-- FORÇA A CRIAÇÃO DE UM ARRAY DE RADIO BUTTONS MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" name="rb_rel" id="rb_rel" value="">

<span id="sREL" class="T">RELATÓRIOS</span>
<div id="dREL" class="QFn" align="center" style="width:560px;">
<table width='100%' cellpadding="0" cellspacing="0" style='margin:6px 0px 10px 0px;'>
	<tr>
		<td align="left" nowrap>
		
			<div style='margin-left:60px;margin-right:30px;'>
	<%  idx = 0
		s_separacao = "" %>
	
	<%	' RELATÓRIO: MULTICRITÉRIO DE PEDIDOS
		if operacao_permitida(OP_CEN_REL_MULTICRITERIO_PEDIDOS_ANALITICO, s_lista_operacoes_permitidas) Or _
		   operacao_permitida(OP_CEN_REL_MULTICRITERIO_PEDIDOS_SINTETICO, s_lista_operacoes_permitidas) then
			idx=idx+1 
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
	
	<% dim s_saida_default, s_checked
	s_saida_default = get_default_valor_texto_bd(usuario, "RelPedidosMCrit|c_carrega_indicadores_estatico") %>
	
	<%	s_checked = ""
		if (InStr(s_saida_default, "ON") <> 0) then s_checked = " checked" %>
	
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório Multicritério de Pedidos</span>
				<input type="checkbox" name="ckb_carrega_indicadores"  id="ckb_carrega_indicadores" value="ON" <%=s_checked %> />
				<img src="../IMAGEM/exclamacao_14x14.png" id="exclamacao" style="cursor:pointer" title="Marque esta opção para que as listas de seleção no filtro sejam exibidas no modo estático" />
	<% end if %>
	
	<%	' RELATÓRIO: MULTICRITÉRIO DE ORÇAMENTOS
		if operacao_permitida(OP_CEN_REL_MULTICRITERIO_ORCAMENTOS, s_lista_operacoes_permitidas) then
			idx=idx+1 
			Response.Write s_separacao 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
	<%	s_saida_default = get_default_valor_texto_bd(usuario, "RelOrcamentosMCrit|c_carrega_indicadores_estatico") %>
	
	<%	s_checked = ""
		if (InStr(s_saida_default, "ON") <> 0) then s_checked = " checked" %>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório Multicritério de Orçamentos</span>
				<input type="checkbox" name="ckb_rel_mcrit_orc_carrega_indicadores" id="ckb_rel_mcrit_orc_carrega_indicadores" value="ON" <%=s_checked %> />
				<img src="../IMAGEM/exclamacao_14x14.png" id="rel_mcrit_orc_exclamacao" style="cursor:pointer" title="Marque esta opção para que as listas de seleção no filtro sejam exibidas no modo estático" />
	<% end if %>
			</div>
	
	<% if (qtde_rel_glb > 0) And (qtde_rel_com > 0) then %>
			<!-- ************   SEPARADOR   ************ -->
			<table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid #C0C0C0; margin: 6px 0px 6px 0px;"><tr><td><span></span></td></tr></table>
	<% end if %>
	
			<div style='margin-left:60px;margin-right:30px;'>
	
	<%	s_separacao = "" %>
	
	<%	' RELATÓRIO: PRODUTOS VENDIDOS
		if operacao_permitida(OP_CEN_REL_PRODUTOS_VENDIDOS, s_lista_operacoes_permitidas) then
			idx=idx+1
			s_separacao = "<br>"
			if (qtde_total_rel = 1) then s=" checked" else s="" 
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Produtos Vendidos</span>
	<% end if %>
	
	<%	' RELATÓRIO: PEDIDOS COLOCADOS NO MÊS
		if operacao_permitida(OP_CEN_REL_PEDIDOS_COLOCADOS_NO_MES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>"
			if (qtde_total_rel = 1) then s=" checked" else s="" 
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Pedidos Colocados no Mês</span>
	<% end if %>
	
	<%	' RELATÓRIO: VENDAS COM DESCONTO SUPERIOR
		if operacao_permitida(OP_CEN_REL_VENDAS_COM_DESCONTO_SUPERIOR, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Vendas com Desconto Superior</span>
	<% end if %>
	
	<%	' RELATÓRIO: COMISSÃO AOS VENDEDORES
		if operacao_permitida(OP_CEN_REL_COMISSAO_VENDEDORES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Comissão aos Vendedores</span>
	<% end if %>
	
	<%	' RELATÓRIO: COMISSÃO AOS VENDEDORES SINTÉTICO (TABELA PROGRESSIVA)
		if operacao_permitida(OP_CEN_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_SINTETICO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Comissão aos Vendedores Sintético (Tabela Progressiva)</span>
	<% end if %>
	
	<%	' RELATÓRIO: COMISSÃO AOS VENDEDORES ANALÍTICO (TABELA PROGRESSIVA)
		if operacao_permitida(OP_CEN_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_ANALITICO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Comissão aos Vendedores Analítico (Tabela Progressiva)</span>
	<% end if %>
	
	<%	' RELATÓRIO: COMISSÃO AOS INDICADORES (ALTERADO P/ RELATÓRIO DE PEDIDOS INDICADORES)
		if operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Indicadores</span>
	<% end if %>

	<%	' RELATÓRIO: RELATÓRIO DE PEDIDOS INDICADORES (VIA NFSe)
		if operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Indicadores (via NFS-e)</span>
	<% end if %>

	<%	' RELATÓRIO: RELATÓRIO DE PEDIDOS INDICADORES (VIA NFSe) (CONSULTA)
		if operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Indicadores (via NFS-e) (Consulta)</span>
	<% end if %>

    <%	' COMISSÃO INDICADORES: CADASTRA NF
		'if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
		'	idx=idx+1
		'	Response.Write s_separacao
		'	s_separacao = "<br>" 
		'	if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
		<!--	<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Comissão de Indicadores: Cadastrar NF</span> //-->
	    <%' end if %>
    <%	' RELATÓRIO: DE PEDIDOS INDICADORES PAGAMENTO
		if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Indicadores (Processamento)</span>
	<% end if %>
   
    <%	'  RELATÓRIO DE PEDIDOS INDICADORES(CONSULTA)
		if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Indicadores (Consulta)</span>
	<% end if %>

     <%	' CONSULTA RELAÇÃO DE DEPÓSITOS
		if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Consultar Relação de Depósitos</span>
	<% end if %>

    <%	' RELATÓRIO: DE PEDIDOS INDICADORES PAGAMENTO
		if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Indicadores com Desconto (Processamento)</span>
	<% end if %>

    <%	'  RELATÓRIO DE PEDIDOS INDICADORES COM DESCONTO (CONSULTA)
		if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Indicadores com Desconto (Consulta)</span>
	<% end if %>

     <%	' CONSULTA RELAÇÃO DE DEPÓSITOS
		if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Consultar Relação de Depósitos com Desconto</span>
	<% end if %>

    <%	' COMISSÃO INDICADORES: PESQUISA INDICADOR
		if operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Comissão Indicadores: Pesquisa Indicador</span>
	<% end if %>
	
	<%	' RELATÓRIO: FATURAMENTO VENDEDORES EXTERNOS
		if operacao_permitida(OP_CEN_REL_FATURAMENTO_VENDEDORES_EXT, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>"
			if (qtde_total_rel = 1) then s=" checked" else s="" 
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Faturamento Vendedores Externos</span>
	<% end if %>
	
	<%	' RELATÓRIO: COMISSÃO DE LOJA POR INDICAÇÃO
		if operacao_permitida(OP_CEN_REL_COMISSAO_LOJA_POR_INDICACAO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Comissão de Loja por Indicação</span>
	<% end if %>
	
	<%	' RELATÓRIO: ANÁLISE DE PEDIDOS
		if operacao_permitida(OP_CEN_REL_ANALISE_PEDIDOS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Análise de Pedidos</span>
	<% end if %>
	
	<%	' RELATÓRIO: DIVERGÊNCIA CLIENTE/INDICADOR
		if operacao_permitida(OP_CEN_REL_DIVERGENCIA_CLIENTE_INDICADOR, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Divergência Cliente/Indicador</span>
	<% end if %>
	
	<%	' RELATÓRIO: METAS DO INDICADOR
		if operacao_permitida(OP_CEN_REL_METAS_INDICADOR, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Metas do Indicador</span>
	<% end if %>
	
	<%	' RELATÓRIO: PERFORMANCE POR INDICADOR
		if operacao_permitida(OP_CEN_REL_PERFORMANCE_INDICADOR, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Performance por Indicador</span>
	<% end if %>
	
	<%	' RELATÓRIO: VENDAS POR BOLETO
		if operacao_permitida(OP_CEN_VENDAS_POR_BOLETO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Vendas por Boleto</span>
	<% end if %>

     <%	' PERFIL DE PAGAMENTO DOS BOLETOS
		if operacao_permitida(OP_CEN_REL_PERFIL_PAGAMENTO_BOLETOS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Perfil de Pagamento dos Boletos</span>
	<% end if %>
	
	<%	' E-COMMERCE: EXPORTAÇÃO DA TABELA DE PRODUTOS
		if operacao_permitida(OP_CEN_REL_ECOMMERCE_EXPORTACAO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>"
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>E-Commerce: Exportação da Tabela de Produtos</span>
	<% end if %>

	
	<%	' DADOS PARA TABELA DINÂMICA
		if operacao_permitida(OP_CEN_REL_DADOS_TABELA_DINAMICA, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>"
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Dados para Tabela Dinâmica</span>
	<% end if %>

    <%	' RELATÓRIO: PEDIDOS CANCELADOS
		if operacao_permitida(OP_CEN_REL_PEDIDOS_CANCELADOS , s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Cancelados </span>
	<% end if %>
	
			</div>
			
	<% if ((qtde_rel_glb + qtde_rel_com) > 0) And (qtde_rel_adm > 0) then %>
			<!-- ************   SEPARADOR   ************ -->
			<table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid #C0C0C0; margin: 6px 0px 6px 0px;"><tr><td><span></span></td></tr></table>
	<% end if %>
	
			<div style='margin-left:60px;margin-right:30px;'>

    <%	s_separacao = "" %>

    <%	' RELATÓRIO DE PRÉ-DEVOLUÇÃO
		if operacao_permitida(OP_CEN_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) then
			idx=idx+1
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pré-Devoluções</span>
	<% end if %>

    <%	' RELATÓRIO: PRÉ-DEVOLUÇÃO: REGISTRAR MERCADORIA RECEBIDA
		if operacao_permitida(OP_CEN_PRE_DEVOLUCAO_RECEBIMENTO_MERCADORIA, s_lista_operacoes_permitidas) then
			idx=idx+1
            Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Registrar Mercadoria Recebida</span>
	<% end if %>
	
	<%	' RELATÓRIO: SEPARAÇÃO
		if operacao_permitida(OP_CEN_REL_SEPARACAO, s_lista_operacoes_permitidas) then
			idx=idx+1
            Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Separação</span>
	<% end if %>
	
	<%	' RELATÓRIO: SEPARAÇÃO (ZONA)
		if operacao_permitida(OP_CEN_REL_SEPARACAO_ZONA, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Separação (Zona)</span>
	<% end if %>
	
	<%	' RELATÓRIO: SEPARAÇÃO (ZONA) - CONSULTA
		if operacao_permitida(OP_CEN_REL_SEPARACAO_ZONA, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Separação (Zona) - Consulta</span>
	<% end if %>
	
	<%	' RELATÓRIO: PRODUTOS NO ESTOQUE DE DEVOLUÇÃO
		if operacao_permitida(OP_CEN_REL_PRODUTOS_ESTOQUE_DEVOLUCAO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Produtos no Estoque de Devolução</span>
	<% end if %>
	
	<%	' RELATÓRIO: DEVOLUÇÃO DE PRODUTOS
		if operacao_permitida(OP_CEN_REL_DEVOLUCAO_PRODUTOS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Devolução de Produtos</span>
	<% end if %>
	
	<%	' RELATÓRIO: DEVOLUÇÃO DE PRODUTOS II
		if operacao_permitida(OP_CEN_REL_DEVOLUCAO_PRODUTOS2, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Devolução de Produtos II</span>
	<% end if %>
	
	<%	' RELATÓRIO: PRODUTOS SPLIT POSSÍVEL
		if operacao_permitida(OP_CEN_REL_PRODUTOS_SPLIT_POSSIVEL, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Produtos Split Possível</span>
	<% end if %>
	
	<%	' RELATÓRIO: ESTOQUE VENDA CRÍTICO
		if operacao_permitida(OP_CEN_REL_ESTOQUE_VENDA_CRITICO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Estoque de Venda Crítico</span>
	<% end if %>
	
	<%	' RELATÓRIO: MEIO DE DIVULGAÇÃO
		if operacao_permitida(OP_CEN_REL_MEIO_DIVULGACAO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Meio de Divulgação</span>
	<% end if %>
	
	<%	' RELATÓRIO: LOG ESTOQUE
		if operacao_permitida(OP_CEN_REL_LOG_ESTOQUE, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Log Estoque</span>
	<% end if %>
	
	<%	' RELATÓRIO: FATURAMENTO (ANTIGO)
		if operacao_permitida(OP_CEN_REL_FATURAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Faturamento (Antigo)</span>
	<% end if %>
	
	<%	' RELATÓRIO: FATURAMENTO (CMV PV)
		if operacao_permitida(OP_CEN_REL_FATURAMENTO_CMVPV, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Faturamento</span>
	<% end if %>
	
	<%	' RELATÓRIO: FATURAMENTO II
		if operacao_permitida(OP_CEN_REL_FATURAMENTO2, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Faturamento II</span>
	<% end if %>
	
	<%	' RELATÓRIO: FATURAMENTO III
		if operacao_permitida(OP_CEN_REL_FATURAMENTO3, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Faturamento III</span>
	<% end if %>
	
	<%	' RELATÓRIO: VENDAS
		if operacao_permitida(OP_CEN_REL_VENDAS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Vendas</span>
	<% end if %>
	
	<%	' RELATÓRIO: GERENCIAL VENDAS
		if operacao_permitida(OP_CEN_REL_GERENCIAL_DE_VENDAS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório Gerencial de Vendas</span>
	<% end if %>
	
	<%	' RELATÓRIO: ROMANEIO DE ENTREGA
		if operacao_permitida(OP_CEN_REL_ROMANEIO_ENTREGA, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Romaneio de Entrega</span>
	<% end if %>
	
	<%	' RELATÓRIO: ANÁLISE DE CRÉDITO
		if operacao_permitida(OP_CEN_REL_ANALISE_CREDITO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Análise de Crédito</span>
	<% end if %>
	
	<%	' PESQUISA DE INDICADORES
		if operacao_permitida(OP_CEN_PESQUISA_INDICADORES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Pesquisa de Indicadores</span>
	<% end if %>
	
	<%	' RELATÓRIO DE INDICADOR SEM AVALIAÇÃO DE DESEMPENHO
		if operacao_permitida(OP_CEN_REL_INDICADOR_SEM_AVALIACAO_DESEMPENHO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Indicador Sem Avaliação de Desempenho</span>
	<% end if %>
	
	<%	' RELATÓRIO DE CHECAGEM DE NOVOS PARCEIROS
		if operacao_permitida(OP_CEN_REL_CHECAGEM_NOVOS_PARCEIROS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Checagem de Novos Parceiros</span>
	<% end if %>
	
	<%	' RELATÓRIO DE FRETE (ANALÍTICO)
		if operacao_permitida(OP_CEN_REL_FRETE_ANALITICO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
	<%	s_saida_default = get_default_valor_texto_bd(usuario, "RelFreteAnalitico|c_carrega_indicadores_estatico") %>
	<%	s_checked = ""
		if (InStr(s_saida_default, "ON") <> 0) then s_checked = " checked" %>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Frete (Analítico)</span>
				<input type="checkbox" name="ckb_rel_frete_analit_carrega_indicadores" id="ckb_rel_frete_analit_carrega_indicadores" value="ON" <%=s_checked %> />
				<img src="../IMAGEM/exclamacao_14x14.png" id="rel_frete_analit_exclamacao" style="cursor:pointer" title="Marque esta opção para que as listas de seleção no filtro sejam exibidas no modo estático" />
	<% end if %>
	
	<%	' RELATÓRIO DE FRETE (SINTÉTICO)
		if operacao_permitida(OP_CEN_REL_FRETE_SINTETICO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
	<%	s_saida_default = get_default_valor_texto_bd(usuario, "RelFreteSintetico|c_carrega_indicadores_estatico") %>
	<%	s_checked = ""
		if (InStr(s_saida_default, "ON") <> 0) then s_checked = " checked" %>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Frete (Sintético)</span>
				<input type="checkbox" name="ckb_rel_frete_sint_carrega_indicadores" id="ckb_rel_frete_sint_carrega_indicadores" value="ON" <%=s_checked %> />
				<img src="../IMAGEM/exclamacao_14x14.png" id="rel_frete_sint_exclamacao" style="cursor:pointer" title="Marque esta opção para que as listas de seleção no filtro sejam exibidas no modo estático" />
	<% end if %>
	
	<%	' RELATÓRIO DE PEDIDOS NÃO RECEBIDOS
		if operacao_permitida(OP_CEN_REL_PEDIDO_NAO_RECEBIDO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Não Recebidos Pelo Cliente</span>
	<% end if %>

    <%	' RELATÓRIO DE PEDIDOS MARKETPLACE NÃO RECEBIDOS
		if operacao_permitida(OP_CEN_REL_PEDIDO_MARKETPLACE_NAO_RECEBIDO_CLIENTE, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Marketplace Não Recebidos Pelo Cliente</span>
	<% end if %>

    <%	' RELATÓRIO: REGISTRO DE PEDIDOS DE MARKETPLACE RECEBIDOS PELO CLIENTE
		if operacao_permitida(OP_CEN_REL_REGISTRO_PEDIDO_MARKETPLACE_RECEBIDO_CLIENTE, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Registro de Pedidos de Marketplace Recebidos Pelo Cliente</span>
	<% end if %>
	
	<%	' RELATÓRIO DE OCORRÊNCIAS
		if operacao_permitida(OP_CEN_REL_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Ocorrências</span>
	<% end if %>
	
	<%	' RELATÓRIO DE ESTATÍSTICAS DE OCORRÊNCIAS
		if operacao_permitida(OP_CEN_REL_ESTATISTICAS_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Estatísticas de Ocorrências</span>
	<% end if %>

    <%	' ACOMPANHAMENTO DE CHAMADOS
		if operacao_permitida(OP_CEN_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Acompanhamento de Chamados</span>
	<% end if %>

    <%	' RELATÓRIO DE CHAMADOS
		if operacao_permitida(OP_CEN_REL_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Chamados</span>
	<% end if %>

    <%	' RELATÓRIO DE ESTATÍSTICAS DE CHAMADOS
		if operacao_permitida(OP_CEN_REL_ESTATISTICAS_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Estatísticas de Chamados</span>
	<% end if %>
	
	<%	' PESQUISA DE ORDEM DE SERVIÇO
		if operacao_permitida(OP_CEN_REL_PESQUISA_ORDEM_SERVICO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Pesquisa de Ordem de Serviço</span>
	<% end if %>
	
	<%	' RELATÓRIO DE CLIENTES NEGATIVADOS
		if operacao_permitida(OP_CEN_REL_CLIENTE_SPC, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Clientes Negativados (SPC)</span>
	<% end if %>
	
	<%	' RELATÓRIO DE TRANSAÇÕES CIELO
		if operacao_permitida(OP_CEN_REL_TRANSACOES_CIELO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Transações Cielo</span>
	<% end if %>
	
	<%	' RELATÓRIO DE TRANSAÇÕES CIELO EM ANDAMENTO
		if operacao_permitida(OP_CEN_REL_TRANSACOES_CIELO_ANDAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>"
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Transações Cielo em Andamento</span>
	<% end if %>
	
	<%	' RELATÓRIO DE TRANSAÇÕES BRASPAG
		if operacao_permitida(OP_CEN_REL_BRASPAG_TRANSACOES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>"
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Transações Braspag</span>
	<% end if %>
	
	<%	' REVISÃO MANUAL ANTIFRAUDE BRASPAG
		if operacao_permitida(OP_CEN_REL_BRASPAG_AF_REVIEW, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Revisão Manual Antifraude Braspag</span>
	<% end if %>

	<%	' RELATÓRIO DE TRANSAÇÕES BRASPAG/CLEARSALE
		if operacao_permitida(OP_CEN_REL_BRASPAG_TRANSACOES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>"
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Transações Braspag/Clearsale</span>
	<% end if %>
	
			</div>
	
	<% if ((qtde_rel_glb + qtde_rel_com + qtde_rel_adm) > 0) And (qtde_rel_compras_logist > 0) then %>
			<!-- ************   SEPARADOR   ************ -->
			<table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid #C0C0C0; margin: 6px 0px 6px 0px;"><tr><td><span></span></td></tr></table>
	<% end if %>
	
			<div style='margin-left:60px;margin-right:30px;'>
			
	<%	s_separacao = "" %>
	
	<%	' RELATÓRIO: ESTOQUE (ANTIGO)
		if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_BASICA, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Estoque (Antigo)</span>
	<% end if %>
	
	<%	' RELATÓRIO: ESTOQUE (CMV PV)
		if operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_BASICA_CMVPV, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_REL_ESTOQUE_VISAO_COMPLETA_CMVPV, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Estoque</span>
	<% end if %>
	
	<%	' RELATÓRIO: ESTOQUE II
		if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Estoque II</span>
	<% end if %>

    <%	' RELATÓRIO: ESTOQUE DE VENDA
		if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Estoque de Venda</span>
	<% end if %>

	<%	' RELATÓRIO: ESTOQUE: RESUMO POSIÇÃO GERAL
		if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Estoque: Resumo Posição Geral</span>
	<% end if %>

    <%	' RELATÓRIO: Estoque (E-Commerce)
		if operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Estoque (E-Commerce)</span>
	<% end if %>
	
	<%	' RELATÓRIO: PRODUTOS PENDENTES
		if operacao_permitida(OP_CEN_REL_PRODUTOS_PENDENTES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Produtos Pendentes</span>
	<% end if %>
	
	<%	' RELATÓRIO: COMPRAS
		if operacao_permitida(OP_CEN_REL_COMPRAS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Compras</span>
	<% end if %>
	
	<%	' RELATÓRIO: COMPRAS II
		if operacao_permitida(OP_CEN_REL_COMPRAS2, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Compras II</span>
	<% end if %>
	
	<%	' RELATÓRIO: RESUMO DE OPERAÇÕES ENTRE ESTOQUES
		if operacao_permitida(OP_CEN_REL_RESUMO_OPERACOES_ENTRE_ESTOQUES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Resumo de Operações Entre Estoques</span>
	<% end if %>
	
	<%	' RELATÓRIO: AUDITORIA DO ESTOQUE
		if operacao_permitida(OP_CEN_REL_AUDITORIA_ESTOQUE, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Auditoria do Estoque</span>
	<% end if %>
	
	<%	' RELATÓRIO: REGISTROS ENTRADA ESTOQUE
		if operacao_permitida(OP_CEN_REL_REGISTROS_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) Or _
		   operacao_permitida(OP_CEN_EDITA_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Registros Entrada Estoque</span>
	<% end if %>
	
	<%	' RELATÓRIO: CONTAGEM DE ESTOQUE
		if operacao_permitida(OP_CEN_REL_CONTAGEM_ESTOQUE, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Contagem de Estoque</span>
	<% end if %>
	
	<%	' RELATÓRIO: ZONA DO PRODUTO (DEPÓSITO)
		if operacao_permitida(OP_CEN_REL_PRODUTO_DEPOSITO_ZONA, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Zona do Produto (Depósito)</span>
	<% end if %>
	
	<%	' RELATÓRIO: SOLICITAÇÃO DE COLETAS
		if operacao_permitida(OP_CEN_REL_SOLICITACAO_COLETAS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Solicitação de Coletas</span>
	<% end if %>
	
	<%	' RELATÓRIO DE PRODUTOS COMPRADOS (FAROL)
		if operacao_permitida(OP_CEN_REL_FAROL_CADASTRO_PRODUTO_COMPRADO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Produtos Comprados (Farol)</span>
	<% end if %>
	
	<%	' RELATÓRIO FAROL RESUMIDO
		if operacao_permitida(OP_CEN_REL_FAROL_RESUMIDO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório Farol Resumido</span>
	<% end if %>
	
	<%	' RELATÓRIO SINTÉTICO DE CUBAGEM, VOLUME E PESO
		if operacao_permitida(OP_CEN_REL_SINTETICO_CUBAGEM_VOLUME_PESO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Cubagem, Volume e Peso (Sintético)</span>
	<% end if %>
	
	<%	' RELATÓRIO HISTÓRICO SINTÉTICO DE CUBAGEM, VOLUME E PESO
		if operacao_permitida(OP_CEN_REL_SINTETICO_CUBAGEM_VOLUME_PESO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório Histórico de Cubagem, Volume e Peso (Sintético)</span>
	<% end if %>
	
		<%	' RELATÓRIO DE IMPOSTOS PAGOS
		if operacao_permitida(OP_CEN_REL_IMPOSTOS_PAGOS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Impostos Pagos</span>
	<% end if %>

		<%	' RELATÓRIO DE CONTROLE DE IMPOSTOS
		if operacao_permitida(OP_CEN_REL_CONTROLE_IMPOSTOS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_total_rel = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Controle de Impostos</span>
	<% end if %>

			</div>
			
			</td>
		</tr>
	</table>

	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" value="CONSULTAR O RELATÓRIO" title="consulta o relatório selecionado">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>
<br />
<% end if %>


<!--  ***********************************************************************************************  -->
<!--  E S T O Q U E																					   -->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_ENTRADA_MERCADORIAS_ESTOQUE, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CONVERSOR_KITS, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_BASICO, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_TRANSF_ENTRE_PED_PROD_ESTOQUE_VENDIDO, s_lista_operacoes_permitidas) then %>
<form method="post" id="fESTOQ" name="fESTOQ" onsubmit="if (!fESTOQConcluir(fESTOQ)) return false;">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FORÇA A CRIAÇÃO DE UM ARRAY DE RADIO BUTTONS MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" name="rb_op" id="rb_op" value="">
<span class="T">ESTOQUE</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td align="left" nowrap>
	<%	idx = 0
		s_separacao = "" %>
		
	<%	' ENTRADA DE MERCADORIAS NO ESTOQUE
		if operacao_permitida(OP_CEN_ENTRADA_MERCADORIAS_ESTOQUE, s_lista_operacoes_permitidas) then
			idx=idx+1 
			s_separacao = "<br>"
	%>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX"><span class="rbLink" onclick="fESTOQ.rb_op[<%=Cstr(idx)%>].click(); fESTOQ.bEXECUTAR.click();"
				>Entrada de Mercadorias no Estoque</span>
	<% end if %>
	
	<%	' ENTRADA DE MERCADORIAS NO ESTOQUE VIA XML
		if operacao_permitida(OP_CEN_ENTRADA_MERCADORIAS_ESTOQUE, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>"
	%>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX"><span class="rbLink" onclick="fESTOQ.rb_op[<%=Cstr(idx)%>].click(); fESTOQ.bEXECUTAR.click();"
				>Entrada de Mercadorias no Estoque (via XML)</span>
	<% end if %>
	
	<%	' CONVERSOR DE KITS
		if operacao_permitida(OP_CEN_CONVERSOR_KITS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
	%>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX"><span class="rbLink" onclick="fESTOQ.rb_op[<%=Cstr(idx)%>].click(); fESTOQ.bEXECUTAR.click();"
				>Conversor de Kits</span>
	<% end if %>
	
	<%	' TRANSFERÊNCIA/MOVIMENTAÇÃO DO ESTOQUE
		if operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_BASICO, s_lista_operacoes_permitidas) Or _
		   operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
	%>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX"><span class="rbLink" onclick="fESTOQ.rb_op[<%=Cstr(idx)%>].click(); fESTOQ.bEXECUTAR.click();"
				>Transferência/Movimentação do Estoque</span>
	<% end if %>
	
	<%	' TRANSFERÊNCIA ENTRE PEDIDOS DE PRODUTOS DO ESTOQUE VENDIDO
		if operacao_permitida(OP_CEN_TRANSF_ENTRE_PED_PROD_ESTOQUE_VENDIDO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
	%>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX"><span class="rbLink" onclick="fESTOQ.rb_op[<%=Cstr(idx)%>].click(); fESTOQ.bEXECUTAR.click();"
				>Transferência Entre Pedidos de Produtos do Estoque Vendido</span>
	<% end if %>

	<%	' TRANSFERÊNCIA DE PRODUTOS ENTRE CD'S
		if operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
	%>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX"><span class="rbLink" onclick="fESTOQ.rb_op[<%=Cstr(idx)%>].click(); fESTOQ.bEXECUTAR.click();"
				>Transferência de Produtos Entre CD's</span>
	<% end if %>

		</td>
		</tr>
	</table>
	
	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" value="EXECUTAR" title="executa">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
	
</div>
</form>
<br />
<%end if%>


<!--  ***********************************************************************************************  -->
<!--  B O T Õ E S  D O S  I T E N S  E M B U T I D O S    											   -->
<!--  ***********************************************************************************************  -->
<% if operacao_permitida(OP_CEN_ANOTA_VALOR_FRETE_NO_PEDIDO, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CADASTRA_SENHA_DESCONTO, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CADASTRO_PERFIL, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CADASTRO_USUARIOS, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CADASTRO_LOJAS, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CADASTRO_GRUPO_LOJAS, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CADASTRO_FABRICANTES, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CADASTRO_TRANSPORTADORAS, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CADASTRO_VEICULOS_MIDIA, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CAD_EC_PRODUTO_COMPOSTO, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CADASTRO_AVISOS, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_CAD_CEP, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_OPCOES_PAGTO_VISANET, s_lista_operacoes_permitidas) Or _ 
	  operacao_permitida(OP_CEN_LER_AVISOS_NAO_LIDOS, s_lista_operacoes_permitidas) Or _ 
	  operacao_permitida(OP_CEN_LER_AVISOS_TODOS, s_lista_operacoes_permitidas) Or _ 
	  operacao_permitida(OP_CEN_CADASTRA_PERDA, s_lista_operacoes_permitidas) then %>
<span class="T">&nbsp;</span>
<div class="QFn" align="center">
<table class="TFn">
	<% if operacao_permitida(OP_CEN_ANOTA_VALOR_FRETE_NO_PEDIDO, s_lista_operacoes_permitidas) then %>
	<tr>
		<td nowrap>
			<form action="PedidoAnotaFrete.asp" method="post" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao Largura" value="Anotar Frete no Pedido  >>" title="anota o valor do frete no(s) pedido(s)">
			</form>
			</td>
		</tr>
        <tr>
        <td nowrap>
			<form action="PedidoAnotaFreteDevolucao.asp" method="post" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao Largura"  value="Anota Frete de Devolução no Pedido  >>" title="anota o valor do frete de devolução ou reentrega no(s) pedido(s)">
			</form>
			</td>
		</tr>
	<%end if%>
	
	<% if operacao_permitida(OP_CEN_CADASTRA_SENHA_DESCONTO, s_lista_operacoes_permitidas) then %>
	<tr>
		<td nowrap>
			<form action="SenhaDescSupPesqCliente.asp" method="post" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao Largura" value="Senha Desconto  >>" title="senha para desconto superior">
			</form>
			</td>
		</tr>
	<%end if%>
	
	<% if operacao_permitida(OP_CEN_CADASTRO_PERFIL, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CADASTRO_USUARIOS, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CADASTRO_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CADASTRO_LOJAS, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CADASTRO_GRUPO_LOJAS, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CADASTRO_FABRICANTES, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CADASTRO_TRANSPORTADORAS, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CADASTRO_VEICULOS_MIDIA, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CAD_MENSAGEM_ALERTA_PRODUTOS, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CAD_EC_PRODUTO_COMPOSTO, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CADASTRO_AVISOS, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CAD_CEP, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_OPCOES_PAGTO_VISANET, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CAD_VL_APROV_AUTO_ANALISE_CREDITO, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CAD_PERC_LIMITE_RA_SEM_DESAGIO, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CAD_PERC_MAX_RT, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CAD_PERC_MAX_DESC_SEM_ZERAR_RT, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CAD_PARAMETROS_GLOBAIS, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CAD_TABELA_COMISSAO_VENDEDOR, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_CAD_TABELA_CUSTO_FINANCEIRO_FORNECEDOR, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_MULTI_CD_CADASTRO_REGRAS_CONSUMO_ESTOQUE, s_lista_operacoes_permitidas) Or _
		  operacao_permitida(OP_CEN_MULTI_CD_ASSOCIACAO_PRODUTO_REGRA, s_lista_operacoes_permitidas) then %>
	<tr>
		<td nowrap>
			<form action="MenuCadastro.asp" method="post" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao Largura" value="Cadastros  >>" title="menu: cadastros">
			</form>
			</td>
		</tr>
	<%end if%>
	
	<% if operacao_permitida(OP_CEN_LER_AVISOS_NAO_LIDOS, s_lista_operacoes_permitidas) Or _ 
		  operacao_permitida(OP_CEN_LER_AVISOS_TODOS, s_lista_operacoes_permitidas) Or _ 
		  operacao_permitida(OP_CEN_CADASTRA_PERDA, s_lista_operacoes_permitidas) Or _ 
		  operacao_permitida(OP_CEN_ETQWMS_EDITA_DADOS_ETIQUETA, s_lista_operacoes_permitidas) then %>
	<tr>
		<td nowrap>
			<form action="MenuOutrasFuncoes.asp" method="post" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao Largura" value="Outras Funções  >>" title="menu: outras funções">
			</form>
			</td>
		</tr>
	<%end if%>
	</table>
</div>
<br />
<%end if%>

</center>

</body>
</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
