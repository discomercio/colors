<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<% Response.Expires = -1 %>
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
	
    dim PRAZO_EXIBICAO_CANCEL_AUTO_PEDIDO
    PRAZO_EXIBICAO_CANCEL_AUTO_PEDIDO = 4

    Const TIPO_CONSULTA_INDICADORES_POR_CPFCNPJ = "CPFCNPJ"
    Const TIPO_CONSULTA_INDICADORES_POR_APELIDO = "APELIDO"

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
'	EXIBIÇÃO DE BOTÕES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True
	
	
'	VERIFICA ID
	dim s, loja, loja_nome, usuario, usuario_nome, senha, senha_real, cadastrado, chave, vendedor_loja, vendedor_externo, strFlagPrimeiraExecucao
	dim dt_ult_alteracao_senha, usuario_bloqueado, confere_login_no_bd, eh_primeira_execucao
	dim strScript, qtde_relatorios, idx, s_separacao, s_operationControlTicket, s_sessionToken
	
	confere_login_no_bd = (Trim(Session("usuario_a_checar")) <> "")	
	loja = Trim(Session("loja_a_checar")): Session("loja_a_checar") = " "
	usuario = Trim(Session("usuario_a_checar")): Session("usuario_a_checar") = " "
	senha = Trim(Session("senha_a_checar")): Session("senha_a_checar") = " "
	
'	OBTEM O ID
	if (loja <> "") then  
		if Len(loja) > 3 then 
			loja = "000"
		else
			loja = retorna_so_digitos(loja)
			if Not IsNumeric(loja) then
				loja = "000"
			elseif CLng(loja) < 1 then 
				loja = "000"
			elseif CLng(loja) > 999 then 
				loja = "000"
				end if
			end if
		end if
	
	if loja = "" then loja = Session("loja_atual")
	if usuario = "" then usuario = Session("usuario_atual")
	if senha = "" then senha = Session("senha_atual")
	usuario_nome = Session("usuario_nome_atual")
	loja_nome = Session("loja_nome_atual")

	dim nivel_acesso_bloco_notas, nivel_acesso_chamado
	dim s_lista_operacoes_permitidas
	dim strSessionCtrlTicket
	dim strMensagemAviso, strMensagemAvisoPopUp
	strMensagemAviso = ""
	strMensagemAvisoPopUp = ""
	
	if (loja = "") or (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
	if (loja = "000") then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICACAO_LOJA)
	if (senha = "") then Response.Redirect("aviso.asp?id=" & ERR_SENHA_NAO_INFORMADA)

	if isHorarioManutencaoSistema then Response.Redirect("aviso.asp?id=" & ERR_HORARIO_MANUTENCAO_SISTEMA)

	loja=normaliza_codigo(loja, TAM_MIN_LOJA)
	
'	CONECTA COM O BANCO DE DADOS
	dim cn,rs,rs2,msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
    If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
    If Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	strFlagPrimeiraExecucao = Request("FlagPrimeiraExecucao")
	if strFlagPrimeiraExecucao = "1" then eh_primeira_execucao = True

'	VERIFICA LOJA NO BD
	if confere_login_no_bd then
		eh_primeira_execucao = true
		loja_nome = trim(x_loja(loja))
		If loja_nome = "" then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICACAO_LOJA)

		s_lista_operacoes_permitidas = obtem_operacoes_permitidas_usuario(cn, usuario)
		nivel_acesso_bloco_notas = obtem_nivel_acesso_bloco_notas_pedido(cn, usuario)
        nivel_acesso_chamado = obtem_nivel_acesso_chamado_pedido(cn, usuario)

	'	VERIFICA USUARIO E SENHA NO BD
		cadastrado = false
		dt_ult_alteracao_senha = null
		usuario_bloqueado=false
		set rs = cn.Execute("select nome, senha, datastamp, dt_ult_alteracao_senha, bloqueado, vendedor_loja, vendedor_externo, SessionCtrlTicket, SessionCtrlLoja, SessionCtrlModulo, SessionCtrlDtHrLogon from t_USUARIO where usuario='" & usuario & "'")
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		
		if Not rs.eof then 
			if Trim("" & rs("SessionCtrlTicket")) <> "" then
				if Trim(Session("TrocaRapidaLoja")) <> "S" then
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
							"'" & SESSION_CTRL_MODULO_LOJA & "'" & _
						")"
					cn.Execute(s)
					end if
				end if
		
		'	TEM SENHA?
			if Trim("" & rs("datastamp")) = "" then usuario_bloqueado=true
		'	ACESSO BLOQUEADO?
			if rs("bloqueado")<>0 then usuario_bloqueado=true
			dt_ult_alteracao_senha = rs("dt_ult_alteracao_senha")
			usuario_nome = Trim("" & rs("nome"))
			vendedor_loja = (rs("vendedor_loja") <> 0)
			vendedor_externo = (rs("vendedor_externo") <> 0)
			if operacao_permitida(OP_CEN_ACESSO_TODAS_LOJAS, s_lista_operacoes_permitidas) then cadastrado = true
			if vendedor_loja then
				s="SELECT loja FROM t_USUARIO_X_LOJA WHERE (usuario='" & usuario & "') AND (CONVERT(smallint,loja)=" & loja & ")"
				set rs2 = cn.Execute(s)
				if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				if Not rs2.Eof then cadastrado = true
				end if

			senha_real = ""
			if cadastrado then
				s = Trim("" & rs("datastamp"))
				chave = gera_chave(FATOR_BD)
				decodifica_dado s, senha_real, chave
				if UCase(trim(senha_real)) <> UCase(trim(senha)) then 
					if senha_real <> "" then senha = ""
					end if
				end if
			end if

		rs.close
		set rs = nothing
			
		if (not cadastrado) or (senha="") then 
			cn.Close
			Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICACAO)
			end if

		if usuario_bloqueado then Response.Redirect("aviso.asp?id=" & ERR_USUARIO_BLOQUEADO)
		
		Session("loja_atual") = loja
		Session("usuario_atual") = usuario
		Session("senha_atual") = senha
		Session("lista_operacoes_permitidas") = s_lista_operacoes_permitidas
		Session("usuario_nome_atual") = usuario_nome
		Session("loja_nome_atual") = loja_nome
		Session("vendedor_loja") = vendedor_loja
		Session("vendedor_externo") = vendedor_externo
		Session("TrocaRapidaLoja") = " "
		Session("nivel_acesso_bloco_notas") = Cstr(nivel_acesso_bloco_notas)
        Session("nivel_acesso_chamado") = Cstr(nivel_acesso_chamado)
		
		strSessionCtrlTicket = GeraTicketSessionCtrl(usuario)
		Session("SessionCtrlTicket") = strSessionCtrlTicket

		Session("SessionCtrlInfo") = MontaSessionCtrlInfo(usuario, SESSION_CTRL_MODULO_LOJA, loja, strSessionCtrlTicket, Session("DataHoraLogon"), Now) 
		
		s = "UPDATE t_USUARIO SET" & _
				" dt_ult_acesso = " & bd_formata_data_hora(Now) & _
				", SessionCtrlDtHrLogon = " & bd_formata_data_hora(Session("DataHoraLogon")) & _
				", SessionCtrlModulo = '" & SESSION_CTRL_MODULO_LOJA & "'" & _
				", SessionCtrlLoja = '" & loja & "'" & _
				", SessionCtrlTicket = '" & strSessionCtrlTicket & "'" & _
				", SessionTokenModuloLoja = newid()" & _
				", DtHrSessionTokenModuloLoja = getdate()" & _
			" WHERE" & _
				" (usuario = '" & usuario & "')"
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
				"'" & loja & "'," & _
				"'" & SESSION_CTRL_MODULO_LOJA & "'," & _
				"'" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'," & _
				"'" & QuotedStr(Trim("" & Request.ServerVariables("HTTP_USER_AGENT"))) & "'" & _
			")"
		cn.Execute(s)

		if IsNull(dt_ult_alteracao_senha) then Response.Redirect("senha.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		
'		COM ESTE REDIRECT, A PÁGINA INICIAL PASSA A TER NA QUERY STRING OS DADOS NECESSÁRIOS P/ RECRIAR A
'		SESSÃO EXPIRADA.
'		QUANDO O USUÁRIO FAZIA O LOGON E NÃO NAVEGAVA P/ NENHUMA OUTRA TELA, AO CLICAR EM F5 NÃO ERA
'		POSSÍVEL RECRIAR A SESSÃO.
		Response.Redirect("resumo.asp?FlagPrimeiraExecucao=1&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		end if  'if (confere_login_no_bd)
		
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	Dim vMsg()
	if Trim(Session("verificar_quadro_avisos")) <> "" then
		Session("verificar_quadro_avisos") = " "
		if recupera_avisos_nao_lidos(loja, usuario, vMsg) then Response.Redirect("quadroavisomostra.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		end if

	Dim opcao_lista_pedidos,opcao_lista_indicadores
	opcao_lista_pedidos = Trim(request("opcao_lista_pedidos"))
    opcao_lista_indicadores = Trim(request("opcao_lista_indicadores"))
	dim blnExibeTrocaRapidaLoja
	blnExibeTrocaRapidaLoja = False
	if operacao_permitida(OP_LJA_LOGIN_TROCA_RAPIDA_LOJA, s_lista_operacoes_permitidas) then 
		if ContaNumLojasAcessoLiberado(usuario, loja) > 1 then blnExibeTrocaRapidaLoja = True
		end if

    if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
        PRAZO_EXIBICAO_CANCEL_AUTO_PEDIDO = 2
    end if

	s_sessionToken = ""
	s = "SELECT Convert(varchar(36), SessionTokenModuloLoja) AS SessionTokenModuloLoja FROM t_USUARIO WHERE (usuario = '" & usuario & "')"
    if rs.State <> 0 then rs.Close
    rs.Open s,cn
	if Not rs.Eof then s_sessionToken = Trim("" & rs("SessionTokenModuloLoja"))

'   LIMPA EVENTUAIS LOCKS REMANESCENTES NOS RELATÓRIOS
    s = "UPDATE tCRUP SET" & _
            " locked = 0," & _
            " cod_motivo_lock_released = " & CTRL_RELATORIO_CodMotivoLockReleased_AcessadaTelaInicialLoja & "," & _
            " dt_hr_lock_released = getdate()" & _
        " FROM t_CTRL_RELATORIO_USUARIO_X_PEDIDO tCRUP INNER JOIN t_CTRL_RELATORIO tCR ON (tCRUP.id_relatorio = tCR.id)" & _
        " WHERE" & _
            " (tCR.modulo = 'LOJA')" & _
            " AND (tCRUP.usuario = '" & QuotedStr(Trim(Session("usuario_atual"))) & "')" & _
            " AND (locked = 1)"
    cn.Execute(s)





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' __________________________
' L I S T A    P E D I D O S
'
function lista_pedidos
dim i,r,s,sql, s_aux


	s = "SELECT data, pedido, st_entrega, vendedor, cnpj_cpf, nome_iniciais_em_maiusculas, analise_credito, analise_credito_pendente_vendas_motivo FROM t_PEDIDO INNER JOIN t_CLIENTE ON t_PEDIDO.id_cliente=t_CLIENTE.id" & _
		" WHERE (loja='" & loja & "')" & _
		" AND (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')" & _
		" AND (st_entrega<>'" & ST_ENTREGA_ENTREGUE & "')"
		
	if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
		s = s & " AND (vendedor = '" & usuario & "')"
		end if
		
	s = s & _
		" ORDER BY data DESC, hora DESC, pedido DESC"

	set r = cn.Execute(s)

	s = "<table width='600' class='QS' cellSpacing='0'>" & chr(13) & _
		"<tr class='DefaultBkg'>" & chr(13) & _
		"<td class='MD MB' align='left'><p class='R'>DATA</p></td>" & chr(13) & _
		"<td class='MD MB' align='left'><p class='R'>PEDIDO</p></td>" & chr(13) & _
		"<td class='MD MB' align='left'><p class='R'>VENDEDOR</p></td>" & chr(13) & _
		"<td class='MD MB' align='left'><p class='R'>CLIENTE</p></td>" & chr(13) & _
		"<td class='MD MB' align='left'><p class='R'>SITUAÇÃO</p></td>" & chr(13) & _
        "<td class='MB' align='left'><p class='R'>STATUS</p></td>" & chr(13) & _
		"</tr>"

	i = 0
	do while Not r.Eof 

		i = i + 1
		if (i AND 1)=0 then
			s = s & "<tr nowrap class='DefaultBkg'>"
		else
			s = s & "<tr nowrap>"
			end if
		s = s & "	<td class='MD' style='width:40px' align='left' valign='top' nowrap><p class='C'>" & formata_data(r("data")) & "</p></td>" & chr(13)
		s = s & "	<td class='MD' style='width:40px' align='left' valign='top' nowrap><p class='C'>&nbsp;<a href='pedido.asp?pedido_selecionado=" & Trim("" & r("pedido")) & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "'>" & r("pedido") & "</a></p></td>" & chr(13)
		s = s & "	<td class='MD' style='width:60px' align='left' valign='top' nowrap><p class='C'>" & iniciais_em_maiusculas(Trim("" & r("vendedor"))) & "</p></td>" & chr(13)
		s = s & "	<td class='MD' style='width:150px' align='left'><p class='C'>" & Trim("" & r("nome_iniciais_em_maiusculas")) & "</p></td>" & chr(13)
		s = s & "	<td class='MD' style='width:50px' align='left' valign='top'><p class='C' style='color:" & x_status_entrega_cor(r("st_entrega"), r("pedido")) & "'>" & x_status_entrega(r("st_entrega")) & "</p></td>" & chr(13)
		s_aux = x_analise_credito(r("analise_credito"))
        if s_aux <> "" And Trim("" & r("analise_credito_pendente_vendas_motivo")) <> "" then
            if Cstr(r("analise_credito"))=Cstr(COD_AN_CREDITO_PENDENTE_VENDAS) then s_aux = s_aux & " (" & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__AC_PENDENTE_VENDAS_MOTIVO, r("analise_credito_pendente_vendas_motivo")) & ")"            
        end if
        s = s & "	<td style='width:80px' align='left' valign='top'><p class='C'>" & s_aux & "</p></td>" & chr(13)
        s = s & "</tr>" & chr(13)
		
		r.MoveNext
		loop

	if i = 0 then
		s = s & "<tr nowrap class='DefaultBkg'>" & _
				"<td align='center' colspan='5'>" & _
				"<p class='C' style='color:red;letter-spacing:1px;'>NENHUM PEDIDO ENCONTRADO.</p>" & _
				"</td></tr>"
		end if
		
	s = s & "</table>" & chr(13)

	lista_pedidos = s

	r.close
	set r=nothing
	
end function
'---------------------------------------
' _________________________________
' PRAZO DE CANCELAMENTO DO PEDIDO
'
'---------------------------------------
function lista_Cancelamento
dim strSql,strWhereBase,s,r,i,n_reg,data_final,cont,vRelat,v_data_final(),v_nome(),v_pedido(),v_descricao(),v_vendedor(),strSqlVlPagoCartao

strWhereBase = " (t1.st_entrega <> '" & ST_ENTREGA_ENTREGUE & "')" & _
								" AND (t1.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
								" AND (t1.st_entrega <> '" & ST_ENTREGA_A_ENTREGAR & "')" & _
								" AND (Coalesce(tPedBase.st_pagto, '') <> '" & ST_PAGTO_PAGO & "')" & _
								" AND (Coalesce(tPedBase.st_pagto, '') <> '" & ST_PAGTO_PARCIAL & "')" & _
                                " AND (tPedBase.loja = '" & loja & "')"

if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
    strWhereBase = strWhereBase & " AND (tPedBase.vendedor = '" & usuario & "')"
    end if

strSqlVlPagoCartao = " Coalesce(" & _
                    "(" & _
                    "SELECT" & _
                        " SUM(payment.valor_transacao)" & _
                    " FROM t_PAGTO_GW_PAG pag INNER JOIN t_PAGTO_GW_PAG_PAYMENT payment ON (pag.id = payment.id_pagto_gw_pag)" & _
                    " WHERE" & _
                        " (pag.pedido = t1.pedido_base)" & _
                        " AND" & _
                        "(" & _
                            " (ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "')" & _
                            " OR" & _
                            " (ult_GlobalStatus = '" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA & "')" & _
                        ")" & _
                    "), 0) AS vl_pago_cartao"

				strSql = "SELECT " & _
							"*" & _
						" FROM (" & _
							"SELECT" & _
								" t1.pedido," & _
                                " t1.pedido_base," & _
								" Coalesce(t1.obs_2, '') AS obs_2," & _
								" t1.transportadora_selecao_auto_status," & _
								" Coalesce(t1.transportadora_id, '') AS transportadora_id," & _
								" t1.st_entrega," & _
								" 'Pendente Cartão de Crédito' AS analise_credito_descricao," & _
								" " & PRAZO_CANCEL_AUTO_PEDIDO_PENDENTE_CARTAO_CREDITO & " AS prazo_cancelamento," & _
								" tPedBase.analise_credito," & _
								" tPedBase.data_hora AS analise_credito_data," & _
								" tPedBase.data AS analise_credito_data_sem_hora," & _
                                " tPedBase.vendedor," & _
                                " nome," &_
								" Coalesce(Datediff(day, tPedBase.data, Convert(datetime, Convert(varchar(10), getdate(), 121), 121)), 0) AS dias_decorridos," & _
								" (" & _
									"SELECT Count(*) FROM t_PEDIDO t2 WHERE (t2.pedido_base = t1.pedido_base) AND (t2.st_auto_split = 0) AND (t2.tamanho_num_pedido > " & TAM_MIN_ID_PEDIDO & ")" & _
								") AS qtde_pedido_filhote," & _
                                strSqlVlPagoCartao & _
							" FROM t_PEDIDO t1" & _
                            " INNER JOIN t_PEDIDO AS tPedBase ON (t1.pedido_base=tPedBase.pedido)" & _
                            " INNER JOIN t_CLIENTE on (t1.id_cliente = t_CLIENTE.id)" & _
							" WHERE" & _
								strWhereBase & _
								" AND (" & _
									"(tPedBase.analise_credito = " & COD_AN_CREDITO_ST_INICIAL & ") AND (tPedBase.st_forma_pagto_somente_cartao = 1)" & _
									" AND (Coalesce(Datediff(day, tPedBase.data, getdate()), 0) > (" & PRAZO_CANCEL_AUTO_PEDIDO_PENDENTE_CARTAO_CREDITO & " - " & PRAZO_EXIBICAO_CANCEL_AUTO_PEDIDO & "))" & _
								")" & _
							" UNION " & _
							"SELECT" & _
								" t1.pedido," & _
                                " t1.pedido_base," & _
								" Coalesce(t1.obs_2, '') AS obs_2," & _
								" t1.transportadora_selecao_auto_status," & _
								" Coalesce(t1.transportadora_id, '') AS transportadora_id," & _
								" t1.st_entrega," & _
								" 'Crédito OK (aguardando depósito)' AS analise_credito_descricao," & _
								" " & PRAZO_CANCEL_AUTO_PEDIDO_CREDITO_OK_AGUARDANDO_DEPOSITO & " AS prazo_cancelamento," & _
								" tPedBase.analise_credito," & _
								" tPedBase.analise_credito_data," & _
								" tPedBase.analise_credito_data_sem_hora," & _
                                " tPedBase.vendedor," & _
                                " nome," &_
								" Coalesce(Datediff(day, tPedBase.analise_credito_data_sem_hora, Convert(datetime, Convert(varchar(10), getdate(), 121), 121)), 0) AS dias_decorridos," & _
								" (" & _
									"SELECT Count(*) FROM t_PEDIDO t2 WHERE (t2.pedido_base = t1.pedido_base) AND (t2.st_auto_split = 0) AND (t2.tamanho_num_pedido > " & TAM_MIN_ID_PEDIDO & ")" & _
								") AS qtde_pedido_filhote," & _
                                strSqlVlPagoCartao & _
							" FROM t_PEDIDO t1" & _
                            " INNER JOIN t_PEDIDO AS tPedBase ON (t1.pedido_base=tPedBase.pedido)" & _
                            " INNER JOIN t_CLIENTE on (t1.id_cliente = t_CLIENTE.id)" & _
							" WHERE" & _
								strWhereBase & _
								" AND (" & _
									"(tPedBase.analise_credito = " & COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO & ")" & _
									" AND (tPedBase.analise_credito_data_sem_hora IS NOT NULL)" & _
									" AND (Coalesce(Datediff(day, tPedBase.analise_credito_data_sem_hora, getdate()), 0) > (" & PRAZO_CANCEL_AUTO_PEDIDO_CREDITO_OK_AGUARDANDO_DEPOSITO & " -  " & PRAZO_EXIBICAO_CANCEL_AUTO_PEDIDO & "))" & _
								")" & _
							" UNION " & _
							"SELECT" & _
								" t1.pedido," & _
								" t1.pedido_base," & _
								" Coalesce(t1.obs_2, '') AS obs_2," & _
								" t1.transportadora_selecao_auto_status," & _
								" Coalesce(t1.transportadora_id, '') AS transportadora_id," & _
								" t1.st_entrega," & _
								" 'Pendente Vendas' AS analise_credito_descricao," & _
								" " & PRAZO_CANCEL_AUTO_PEDIDO_PENDENTE_VENDAS & " AS prazo_cancelamento," & _
								" tPedBase.analise_credito," & _
								" tPedBase.analise_credito_data," & _
								" tPedBase.analise_credito_data_sem_hora," & _
                                " tPedBase.vendedor," & _
                                " nome," &_
								" Coalesce(Datediff(day, tPedBase.analise_credito_data_sem_hora, Convert(datetime, Convert(varchar(10), getdate(), 121), 121)), 0) AS dias_decorridos," & _
								" (" & _
									"SELECT Count(*) FROM t_PEDIDO t2 WHERE (t2.pedido_base = t1.pedido_base) AND (t2.st_auto_split = 0) AND (t2.tamanho_num_pedido > " & TAM_MIN_ID_PEDIDO & ")" & _
								") AS qtde_pedido_filhote," & _
                                strSqlVlPagoCartao & _
							" FROM t_PEDIDO t1" & _
                            " INNER JOIN t_PEDIDO AS tPedBase ON (t1.pedido_base=tPedBase.pedido)" & _
                            " INNER JOIN t_CLIENTE on (t1.id_cliente = t_CLIENTE.id)" & _
							" WHERE" & _
								strWhereBase & _
								" AND (" & _
									"(tPedBase.analise_credito = " & COD_AN_CREDITO_PENDENTE_VENDAS & ")" & _
									" AND (tPedBase.analise_credito_data_sem_hora IS NOT NULL)" & _
									" AND (Coalesce(Datediff(day, tPedBase.analise_credito_data_sem_hora, getdate()), 0) >  (" & PRAZO_CANCEL_AUTO_PEDIDO_PENDENTE_VENDAS & " -  " & PRAZO_EXIBICAO_CANCEL_AUTO_PEDIDO & "))" & _
								")" & _
							") t" & _
						" WHERE" & _
							" (qtde_pedido_filhote = 0)" & _
							" AND (LEN(obs_2) = 0)" & _
                            " AND (vl_pago_cartao = 0)" & _
							" AND ((transportadora_selecao_auto_status = 1) OR (LEN(Coalesce(transportadora_id,'')) = 0))" & _
						" ORDER BY" & _
                            " analise_credito_data_sem_hora," & _
							" analise_credito," & _							
							" pedido"

    set r = cn.Execute(strSql)


     s = "<table width='600' class='QS' cellSpacing='0' style='border-left:0px;'>" & chr(13) & _
		        "<tr class='DefaultBkg'>" & chr(13) & _
                "<td style='background:#FFF' class='MD MB'> &nbsp; </td>" & chr(13) & _
		        "<td class='MD MTB' align='left'><p class='R'>DATA FINAL</p></td>" & chr(13) & _
		        "<td class='MD MTB' align='left'><p class='R'>PEDIDO</p></td>" & chr(13)
    if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
	    s = s & "<td class='MD MTB' align='left'><p class='R'>VENDEDOR</p></td>" & chr(13)
    end if
	s = s & "<td class='MD MTB' align='left'><p class='R'>NOME DO CLIENTE</p></td>" & chr(13) & _
    "<td class='MTB' align='left'><p class='R'>ANÁLISE DE CRÉDITO</p></td>" & chr(13) & _                 
	"</tr>"
    i = 0
    cont = 0
    if not r.Eof then
	    do while Not r.Eof 
            data_final = DateAdd("d",r("prazo_cancelamento"),r("analise_credito_data_sem_hora"))           
                n_reg = n_reg + 1
                redim preserve v_data_final(cont)
                redim preserve v_nome(cont) 
                redim preserve v_pedido(cont)
                redim preserve v_descricao(cont)
                redim preserve v_vendedor(cont)
                v_data_final(cont) = data_final
                v_nome(cont) = r("nome")
                v_pedido(cont) = r("pedido")
                v_descricao(cont) = r("analise_credito_descricao")
                v_vendedor(cont) = r("vendedor")
                cont = cont + 1                 
		    r.MoveNext
	    loop
  ' ORDENAÇÃO       
        if n_reg <> 0 then
            redim vRelat(0)
	        set vRelat(0) = New cl_CINCO_COLUNAS
	        with vRelat(0)
		        .c1 = ""
		        .c2 = ""
		        .c3 = ""
		        .c4 = ""
		        .c5 = ""
		    end with
            if v_data_final(Ubound(v_data_final)) <> "" then
                for cont = 0 to Ubound(v_data_final)
                    if Trim(vRelat(ubound(vRelat)).c1) <> "" then
				        redim preserve vRelat(ubound(vRelat)+1)
				        set vRelat(ubound(vRelat)) = New cl_CINCO_COLUNAS
			        end if
			        with vRelat(ubound(vRelat))
				        .c1 =  v_data_final(cont)
                        .c2 =  v_pedido(cont)
                        .c3 =  v_nome(cont)
                        .c4 =  v_descricao(cont)
                        .c5 =  v_vendedor(cont)
			        end with
                next
            end if
            ordena_cl_cinco_colunas vRelat, 0, Ubound(vRelat)
            n_reg = 0 
  ' PREENCHE A TABELA         
            for cont = 0 to Ubound(v_data_final)    
                    n_reg = n_reg + 1	        
                    i = i + 1
		            if (i AND 1)=0 then
			            s = s & "<tr nowrap class='DefaultBkg'>" & chr(13)
		            else
			            s = s & "<tr nowrap>" & chr(13)
			        end if
		            s = s & "	<td class='tdReg' style='width:20px' nowrap><p class='Rd'>" & n_reg & ".</p></td>" & chr(13)
		            s = s & "	<td class='tdDataF' nowrap><p class='C'>" & formata_data(vRelat(cont).c1) & "</p></td>" & chr(13)
		            s = s & "   <td class='tdPedido' nowrap><p class='C'>&nbsp;<a href='javascript:fRELConsul(" & _
			        chr(34) & vRelat(cont).c2 & chr(34) & _
			        ")' title='clique para consultar o pedido'>" & vRelat(cont).c2 & "</a></p></td>" & chr(13)
                    if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
		                s = s & "	<td class='tdVen' nowrap><p class='C'>" & vRelat(cont).c5 & "</p></td>" & chr(13)
                    end if
		            s = s & "	<td class='tdCliente'><p class='C'>" & vRelat(cont).c3 & "</p></td>" & chr(13)
                    s = s & "	<td class='tdAnalise'><p class='C'>" & vRelat(cont).c4 & "</p></td>" & chr(13)      
		            s = s & "</tr>" & chr(13)
            next
        end if    
    end if
    if i = 0 then
	s = s & "<tr nowrap class='DefaultBkg'>" & _
            "<td style='width:20px;background: #FFF;' class='MD ME'><p class='Rd'>0.</p></td>" & _
			"<td align='center' colspan='5'>" & _
			"<p class='C' style='color:red;letter-spacing:1px;'>NENHUM PEDIDO.</p>" & _
			"</td></tr>"
	end if
		
    s = s & "</table>" & chr(13)

	lista_Cancelamento = s

	r.close
	set r=nothing

end function
'---------------------------------------
' _________________________________
' INDICADORES SEM ATIVIDADE RECENTE
'
'---------------------------------------

function lista_indicadores
dim i,r,s,s_sql,cont,vRelat()
dim v_apelido(),v_indicador(),v_pedido(),v_qtde(),n_reg,s_where
n_reg = 0
    
    s_where =  " AND ( vendedor = '" & Trim(replace(usuario, "'", "''")) & "')"
      
	s = " SELECT *" & _
        " FROM (" & _
	        " SELECT DATEDIFF(day, Coalesce(entregue_data, vendedor_dt_ult_atualizacao, dt_cadastro), getdate()) AS qtde_dias," & _
		        " pedido," & _
		        " apelido," & _
		        " razao_social_nome," & _
		        " vendedor" & _
	        " FROM (" & _
		        " SELECT (" & _
				        " SELECT entregue_data" & _
				        " FROM t_PEDIDO" & _
				        " WHERE t_PEDIDO.pedido = t.pedido" & _
				        " ) AS entregue_data," & _
			        " t.pedido," & _
			        " t.apelido," & _
			        " t.vendedor_dt_ult_atualizacao," & _
			        " t.dt_cadastro," & _
			        " t.razao_social_nome," & _
			        " t.vendedor" & _
		        " FROM (" & _
			        " SELECT (" & _
					        " SELECT TOP 1 pedido" & _
					        " FROM t_PEDIDO" & _
					        " WHERE st_entrega = '" & ST_ENTREGA_ENTREGUE & "'" & _
						        " AND indicador = t_ORCAMENTISTA_E_INDICADOR.apelido" & _
					        " ORDER BY entregue_data DESC," & _
						        " data_hora DESC" & _
					        " ) AS pedido," & _
				        " t_ORCAMENTISTA_E_INDICADOR.apelido," & _
				        " t_ORCAMENTISTA_E_INDICADOR.vendedor_dt_ult_atualizacao," & _
				        " t_ORCAMENTISTA_E_INDICADOR.dt_cadastro," & _
				        " t_ORCAMENTISTA_E_INDICADOR.razao_social_nome," & _
				        " t_ORCAMENTISTA_E_INDICADOR.vendedor" & _
			        " FROM t_ORCAMENTISTA_E_INDICADOR" & _
			        " WHERE  STATUS = 'A'" & _
				         s_where & _
			        " ) t" & _
		        " ) t2" & _
	        " ) t3" & _
        " WHERE (" & _
		        " (qtde_dias >= 45)" & _
		        " OR (qtde_dias IS NULL)" & _
		        " )" & _
        " ORDER BY CASE " & _
		        " WHEN qtde_dias IS NULL" & _
			        " THEN 999999" & _
		        " ELSE qtde_dias" & _
		        " END DESC"

	set r = cn.Execute(s)
  
     s = "<table width='600' class='QS' cellSpacing='0' style='border-left:0px;'>" & chr(13) & _
		        "<tr class='DefaultBkg'>" & chr(13) & _
                "<td style='background:#FFF' class='MD MB'> &nbsp; </td>" & chr(13) & _
		        "<td class='MD MTB' align='right'><p class='R'>QTDE <br> DE DIAS</p></td>" & chr(13) & _
		        "<td class='MD MTB' align='left'><p class='R'>PEDIDO</p></td>" & chr(13) & _
		        "<td class='MD MTB' align='left'><p class='R'>APELIDO</p></td>" & chr(13) & _
                "<td class='MTB' align='left'><p class='R'>NOME INDICADOR</p></td>" & chr(13) & _               
		        "</tr>"
	i = 0
    cont = 0
    if not r.Eof then
	    do while Not r.Eof 

            n_reg = n_reg + 1  
            redim preserve v_apelido(cont)
            redim preserve v_indicador(cont)
            redim preserve v_pedido(cont)
            redim preserve v_qtde(cont)
            redim preserve v_vendedores(cont)
            v_apelido(cont) = r("apelido")
            v_indicador(cont) = r("razao_social_nome")
            v_pedido(cont) = r("pedido")
            v_qtde(cont) = r("qtde_dias")           
            
  	 
            i = i + 1
		    if (i AND 1)=0 then
			    s = s & "<tr nowrap class='DefaultBkg'>" & chr(13)
		    else
			    s = s & "<tr nowrap>" & chr(13)
			end if
		    s = s & "	<td class='tdn_reg' nowrap><p class='Rd'>" & n_reg & ".</p></td>" & chr(13)
		    s = s & "	<td class='tdQtde' nowrap><p class='C'>" & v_qtde(cont) & "</p></td>" & chr(13)
		    s = s & "   <td class='tdPed' nowrap><p class='C'>&nbsp;<a href='javascript:fRELCon(" & _
			chr(34) & v_pedido(cont) & chr(34) & _
			")' title='clique para consultar o pedido'>" & v_pedido(cont) & "</a></p></td>" & chr(13)
		    s = s & "	<td class='tdAp' nowrap><p class='C'>" & iniciais_em_maiusculas( v_apelido(cont)) & "</p></td>" & chr(13)
            s = s & "	<td class='tdInd' nowrap><p class='C'>" &   v_indicador(cont) & "</p></td>" & chr(13)           
		    s = s & "</tr>" & chr(13)
                        
            cont = cont + 1
		    r.MoveNext
	    loop
    end if
    if i = 0 then
	s = s & "<tr nowrap class='DefaultBkg'>" & _
            "<td style='width:20px;background: #FFF;' class='MD ME'><p class='Rd'>0.</p></td>" & _
			"<td align='center' colspan='4'>" & _
			"<p class='C' style='color:red;letter-spacing:1px;'>NENHUM INDICADOR ENCONTRADO.</p>" & _
			"</td></tr>"
	end if

     s_sql = "SELECT COUNT('apelido') as n_ind  FROM t_ORCAMENTISTA_E_INDICADOR WHERE vendedor = '" & usuario & "' AND status = 'A'"
    if rs.State <> 0 then rs.Close
    rs.Open s_sql,cn
    s = s & "<tr nowrap style='background: #FFF;'>" & _                           
				        "<td align='center' colspan='7' class='MTE'>" & _
				        "<p class='C' style='letter-spacing:1px;'>TOTAL DE INDICADORES NA CARTEIRA: "& rs("n_ind") &"</p>" & _
				        "</td>" & _
            "</tr>"
		
    s = s & "</table>" & chr(13)

	lista_indicadores = s

	r.close
	set r=nothing
	
end function

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

<script language="JavaScript" type="text/javascript">
window.focus();
</script>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<% if eh_primeira_execucao then %>
<script language="JavaScript" type="text/javascript">
configura_painel();
</script>
<% end if %>

<script type="text/javascript">
//Dynamically assign height
function sizeDivAjaxRunning() {
	var newTop = $(window).scrollTop() + "px";
	$("#divAjaxRunning").css("top", newTop);
}
</script>

<script type="text/javascript">
    $(function() {

		$("#divAjaxRunning").hide();

		$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8

		$(document).ajaxStart(function () {
			$("#divAjaxRunning").show();
		})
		.ajaxStop(function () {
			$("#divAjaxRunning").hide();
		});

		//Every resize of window
		$(window).resize(function() {
			sizeDivAjaxRunning();
		});

		//Every scroll of window
		$(window).scroll(function() {
			sizeDivAjaxRunning();
		});
       
	   // Evita o submit do form ao pressionar a tecla Enter no campo do nº pedido
    	$(document).on("keypress", "#fPNEC", function(event) { 
    		return event.keyCode != 13;
    	});

        $(document).tooltip();
        
        $("#divTelaCheia").css('filter', 'alpha(opacity=30)');

        <% if SWITCH_QUADRO_AVISO_POPUP = 1 then %>

        CarregaAvisoNovo();
        <%end if%>

        var topo = $('#divQuadroOrcamento').offset().top - parseFloat($('#divQuadroOrcamento').css('margin-top').replace(/auto/, 0)) - parseFloat($('#divQuadroOrcamento').css('padding-top').replace(/auto/, 0));
        $('#divQuadroOrcamento').addClass('divFixo');

        // Carrega Lista de pré-pedidos
        CarregaOrcamentoNovo();
        setInterval(CarregaOrcamentoNovo, <%=TIMER_CARREGA_ORCAMENTO_NOVO_MILISSEGUNDOS%>);

    });
</script>
<script type="text/javascript">
    $(function() {
        $('#cpf_cnpj_selecionado').autocomplete({
            source: function( request, response ) {
                $.ajax({
                    url:  "../Global/JsonPesquisaIndicadores.asp",
                    dataType: "json",
                    data: {
                        q: request.term,
                        loja: "<%=loja%>",
                        tipo_consulta: "<%=TIPO_CONSULTA_INDICADORES_POR_CPFCNPJ%>",
                        modulo: "<%=COD_OP_MODULO_LOJA%>",
                        usuario: "<%=usuario%>"
                    },
                    success: function( data ) {
                        response( data );
                    }
                });
            },
            minLength: 3
        });
        $('#id_selecionado').autocomplete({
            source: function( request, response ) {
                $.ajax({
                    url:  "../Global/JsonPesquisaIndicadores.asp",
                    dataType: "json",
                    data: {
                        q: request.term,
                        loja: "<%=loja%>",
                        tipo_consulta: "<%=TIPO_CONSULTA_INDICADORES_POR_APELIDO%>",
                        modulo: "<%=COD_OP_MODULO_LOJA%>",
                        usuario: "<%=usuario%>"
                    },
                    success: function( data ) {
                        response( data );
                    }
                });
            },
            minLength: 3
        });
        
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
            strUrl = strUrl + "?loja=" + "<%=loja%>";
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
        strUrl = strUrl + "&loja=" + "<%=loja%>";
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

<script type="text/javascript">
    function TrataDadosOrcamentoNovo() {
        var f, i, strResp, usuario, xmlResp, qtde, divLista, t, r, c;
        
        if (objAjaxOrcamentoNovo.readyState == AJAX_REQUEST_IS_COMPLETE) {
            strResp = objAjaxOrcamentoNovo.responseText;
            $("#divListaPrePedidos").children().remove();
            if (strResp == "") {
                qtde = 0;
                $("#qtdePrePedidos").css('color', '#888');
                $("#qtdePrePedidos").text(qtde);
                window.status = "Concluído";
                return;
            }

            if (strResp != "") {
                try {
                    xmlResp = objAjaxOrcamentoNovo.responseXML.documentElement;
                    qtde = xmlResp.getElementsByTagName('registro').length;
                    divLista = document.getElementById('divListaPrePedidos');
                    t = document.createElement('TABLE');
                    t.border=0;
                    t.width='100%';
                    t.cellPadding = 0;
                    t.cellSpacing=0;
                    if (qtde > 0) {
                        $("#qtdePrePedidos").css('color', 'darkgreen');
                    }
                    else {
                        $("#qtdePrePedidos").css('color', '#888');
                    }
                    $("#qtdePrePedidos").text(qtde);
                    for (i = 0; i < xmlResp.getElementsByTagName('registro').length; i++) {
                        orc = document.createElement('A');
                        ind = document.createElement('A');
                        
                        orc.href = 'javascript:fPesqPrePedido("' + xmlResp.getElementsByTagName('orcamento')[i].childNodes[0].nodeValue + '")';
                        orc.innerText = xmlResp.getElementsByTagName('orcamento')[i].childNodes[0].nodeValue;
                        orc.className = 'C';

                        r = document.createElement('TR');
                        r.setAttribute('onmouseover', 'realca_cor_mouse_over(this)');
                        r.setAttribute('onmouseout', 'realca_cor_mouse_out(this)');
                        t.appendChild(r);

                        c = document.createElement('TD');
                        c.appendChild(orc);
                        r.appendChild(c);

                        ind.href = 'javascript:fPesqPrePedido("' + xmlResp.getElementsByTagName('orcamento')[i].childNodes[0].nodeValue + '")';
                        ind.innerText = xmlResp.getElementsByTagName('orcamentista')[i].childNodes[0].nodeValue;
                        ind.className = 'C';

                        c = document.createElement('TD');
                        c.appendChild(ind);
                        c.align='left';
                        r.appendChild(c);

                        divLista.appendChild(t);

                    }
                    
                }
                catch (e) {
                    alert("Falha na consulta de novos orçamentos!!");
                }
            }
            window.status = "Concluído";
        }
    }
    
    function CarregaOrcamentoNovo() {
        var f, strUrl, usuario;

        usuario = "<%=usuario%>";
        objAjaxOrcamentoNovo = GetXmlHttpObject();
        if (objAjaxOrcamentoNovo == null) {
            alert("O browser NÃO possui suporte ao AJAX!!");
            return;
        }

        window.status = "Pesquisando por novos orçamentos ...";

        strUrl = "../Global/AjaxCarregaOrcamentosNovos.asp";
        strUrl = strUrl + "?loja=" + "<%=loja%>";

        <%if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then%>
        strUrl = strUrl + "&vendedor=" + usuario;
        <%end if%>

        //  Prevents server from using a cached file
        strUrl = strUrl + "&sid=" + Math.random() + Math.random();
        objAjaxOrcamentoNovo.onreadystatechange = TrataDadosOrcamentoNovo;
        objAjaxOrcamentoNovo.open("GET", strUrl, true);
        objAjaxOrcamentoNovo.send(null);
    }

    function exibeOcultaListaPrePedidos() {
        if($("#divListaPrePedidos").is(':visible')) {
            $("#divListaPrePedidos").slideUp();
            $("#spnExpandirPrePedido").html("&or;");            
        }
        else {
            $("#divListaPrePedidos").slideDown();
            $("#spnExpandirPrePedido").html("&and;");          
        }
    }
    function realca_cor_mouse_over(c) {
        c.style.backgroundColor = '#ddd';
    }

    function realca_cor_mouse_out(c) {
        c.style.backgroundColor = '';
    }

</script>

<script language="JavaScript" type="text/javascript">
var fCepPopup;
var pedidoMagento;
var serverVariableUrl;

	serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
	serverVariableUrl = serverVariableUrl.toUpperCase();
	serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("LOJA"));

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

function fRELCon(id_pedido) {
    window.status = "Aguarde ...";
    fRELConsulta.pedido_selecionado.value = id_pedido;
    fRELConsulta.action = "pedido.asp"
    fRELConsulta.submit();
}
function fRELConsul(id_pedido) {
    window.status = "Aguarde ...";
    fRELCons.pedido_selecionado.value = id_pedido;
    fRELCons.action = "pedido.asp"
    fRELCons.submit();
}

function fOFConcluir( f ){
var s, iop;

	iop=-1;
	s="";

    // LEITURA DO QUADRO DE AVISOS (SOMENTE NÃO LIDOS)
	iop++;
	if (f.rb_op[iop].checked) {
		s="quadroavisomostra.asp";
		f.opcao_selecionada.value="";
		}

 // LEITURA DO QUADRO DE AVISOS (TODOS OS AVISOS)
	iop++;
	if (f.rb_op[iop].checked) {
		s="quadroavisomostra.asp";
		f.opcao_selecionada.value="S";
		}

 // FUNÇÕES ADMINISTRATIVAS (SENHA DE DESCONTO SUPERIOR)
	iop++;
	if (f.rb_op[iop].checked) {
		s="MenuFuncoesAdministrativas.asp";
		}

	if (s=="") {
		alert("Escolha uma das funções!!");
		return false;
		}

	window.status = "Aguarde ...";
	f.action=s;
	f.submit();
}

function restauraVisibility(nome_controle) {
var c;
	c = document.getElementById(nome_controle);
	if (c) c.style.visibility = "";
}

function fPNECConcluir(f)
{
	var msg;
	if (pedidoMagento.magentoSalesOrderInfo.increment_id == null)
	{
		msg="Pedido Magento nº " + pedidoMagento.numeroPedidoMagento + " não foi encontrado!!\nNão é possível prosseguir!!";
		alert(msg);
		return false;
	}

	if (pedidoMagento.cpfCnpjIdentificado.length == 0)
	{
		msg="CPF/CNPJ do cliente não foi localizado nos dados do pedido!!\nNão é possível prosseguir!!";
		alert(msg);
		return false;
	}

	if (pedidoMagento.erpSalesOrderJaCadastrado.pedido.length > 0)
	{
		msg="O pedido Magento nº " + pedidoMagento.numeroPedidoMagento + " já está cadastrado no sistema!!"+
				"\nPedido: " + pedidoMagento.erpSalesOrderJaCadastrado.pedido +
				"\nCadastrado em: " + pedidoMagento.erpSalesOrderJaCadastrado.dt_hr_cadastro_formatado +
				"\nCadastrado por: " + pedidoMagento.erpSalesOrderJaCadastrado.usuario_cadastro;
		alert(msg);
		return false;
	}

	c=document.getElementById("bPNECConcluir");
	c.style.visibility = "hidden";
	setTimeout("restauraVisibility('bPNECConcluir')", 30000);
	window.status = "Aguarde ...";
	f.submit();
}

function ConsultaPedidoMagentoAjax(f)
{
	var c, numero_magento, operationControlTicket, loja, usuario, sessionToken;
	numero_magento=f.c_numero_magento.value;
	numero_magento=retorna_so_digitos(numero_magento);
	if (numero_magento.length!=9)
	{
		alert("Nº pedido Magento com formato inválido!");
		f.c_numero_magento.focus();
		return false;
	}
	
	operationControlTicket = $("#operationControlTicket").val();
	loja = "<%=loja%>";
	usuario = "<%=usuario%>";
	sessionToken = $("#sessionToken").val();

	$("#divAjaxRunning").show();
	var jqxhr = $.ajax({
		url: 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/MagentoApi/GetPedido',
		type: "GET",
		dataType: 'json',
		data: {
			numeroPedidoMagento: numero_magento,
			operationControlTicket: operationControlTicket,
			loja: loja,
			usuario: usuario,
			sessionToken: sessionToken
		}
	})
	.done(function (response) {
		$("#divAjaxRunning").hide();
		pedidoMagento = response;
		fPNECConcluir(f);
	})
	.fail(function (jqXHR, textStatus) {
		$("#divAjaxRunning").hide();
		var msgErro = "";
		if (textStatus.toString().length > 0) msgErro = "Mensagem de Status: " + textStatus.toString();
		try {
			if (jqXHR.status.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Status: " + jqXHR.status.toString();}
		} catch (e) { }

		try {
			if (jqXHR.statusText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Descrição do Status: " + jqXHR.statusText.toString();}
		} catch (e) { }
		
		try {
			if (jqXHR.responseText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Mensagem de Resposta: " + jqXHR.responseText.toString();}
		} catch (e) { }
		
		alert("Falha ao tentar processar a requisição!!\n\n" + msgErro);
	});
}

function fCPConcluir( f ){
var s, c;
	s=f.cnpj_cpf_selecionado.value;
	s=retorna_so_digitos(s);
	if (!cnpj_cpf_ok(s)) {
		alert("CNPJ/CPF inválido!!");
		f.cnpj_cpf_selecionado.focus();
		return false;
		}
	c = document.getElementById("bPesqCliEXECUTAR");
	c.style.visibility = "hidden";
	setTimeout("restauraVisibility('bPesqCliEXECUTAR')", 30000);
	window.status = "Aguarde ...";
	f.submit();
}

function fESTConcluir( f ){
var s;
	s = retorna_so_digitos(f.c_fabricante.value);
	if (s=="") {
		alert("Preencha o código do fabricante!!");
		f.c_fabricante.focus();
		return;
		}
		
	s = trim(f.c_produto.value);
	if (s=="") {
		alert("Preencha o código do produto!!");
		f.c_produto.focus();
		return;
		}

	window.status = "Aguarde ...";
	f.submit();
}

function fLPRECOSConcluir( f ){
var s;
	s = retorna_so_digitos(f.c_fabricante.value);
	if (s=="") {
		alert("Preencha o código do fabricante!!");
		f.c_fabricante.focus();
		return;
		}

	window.status = "Aguarde ...";
	f.submit();
}

function fCustoFinancFornecConcluir( f ){
var s;
	window.status = "Aguarde ...";
	f.submit();
}

function fIndConsultarPorApelido( f ){
    window.status = "Aguarde ...";
    f.action = "OrcamentistaEIndicadorConsulta.asp";
    f.submit();
}

function fIndEditarPorApelido( f ){
    window.status = "Aguarde ...";
    f.action = "OrcamentistaEIndicadorEdita.asp";
    f.submit();
}

function fIndConsultarPorCpfCnpj( f ) {
    window.status = "Aguarde ...";
    f.action = "OrcamentistaEIndicadorConsulta.asp";
    f.submit();
}

function fIndEditarPorCpfCnpj( f ) {
    window.status = "Aguarde ...";
    f.action = "OrcamentistaEIndicadorEdita.asp";
    f.submit();
}

function fPRODBLOQConcluir( f ){
var s;
	s = retorna_so_digitos(f.c_fabricante.value);
	if (s=="") {
		alert("Preencha o código do fabricante!!");
		f.c_fabricante.focus();
		return;
		}
		
	s = trim(f.c_produto.value);
	if (s=="") {
		alert("Preencha o código do produto!!");
		f.c_produto.focus();
		return;
		}

	window.status = "Aguarde ...";
	f.submit();
}

function fPesqPrePedido(orcamento) {
    window.status = "Aguarde ...";
    fPEDPESQ.orcamento_selecionado.value=orcamento;
    fPEDPESQ.action = "Orcamento.asp"
    fPEDPESQ.submit(); 
}
</script>

<script type="text/javascript">
	function exibeJanelaCEP_Consulta() {
		$.mostraJanelaCEP(null);
	}
</script>

<% 
	strScript = _
		"<script language='JavaScript'>" & chr(13) & _
		"function fRELConcluir( f ){" & chr(13) & _
		"var s_dest, iop;" & chr(13) & _
		"	iop=0;" & chr(13) & _
		"	s_dest='';" & chr(13) & _
		"" & chr(13)
	
	if operacao_permitida(OP_LJA_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO: COMISSÃO AOS INDICADORES (ALTERADO P/ RELATÓRIO DE PEDIDOS INDICADORES)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadores.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
    if operacao_permitida(OP_LJA_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
        strScript = strScript & _
			" // Relatorio pedidos indicadores (processado)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadoresConsultaPagos.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
    end if
    if Day(Date) <= PRAZO_ACESSO_REL_PEDIDOS_INDICADORES_LOJA then
        if operacao_permitida(OP_LJA_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
    strScript = strScript & _
			" // Relatorio pedidos indicadores (Preview)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoIndicadoresConsultaExec.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
        end if
    end if
	if operacao_permitida(OP_LJA_REL_FATURAMENTO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // FATURAMENTO (ANTIGO)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelVendas.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_FATURAMENTO_CMVPV, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // FATURAMENTO (CMV PV)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelVendasCmvPv.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_VENDAS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // VENDAS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelVendasVariante.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_VENDAS_COM_DESCONTO_SUPERIOR, s_lista_operacoes_permitidas) then
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

	if operacao_permitida(OP_LJA_REL_MULTICRITERIO_PEDIDOS_ANALITICO, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_LJA_REL_MULTICRITERIO_PEDIDOS_SINTETICO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO MULTICRITÉRIO DE PEDIDOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidosMCrit.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_PEDIDOS_COLOCADOS_NO_MES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // PEDIDOS COLOCADOS NO MÊS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidosColocados.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_COMISSAO_VENDEDORES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // COMISSÃO AOS VENDEDORES" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissao.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_SINTETICO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // COMISSÃO AOS VENDEDORES SINTÉTICO (TABELA PROGRESSIVA)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoTabelaProgressivaSintetico.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_ANALITICO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // COMISSÃO AOS VENDEDORES ANALÍTICO (TABELA PROGRESSIVA)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelComissaoTabelaProgressivaAnalitico.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_SINTETICO, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // POSIÇÃO DO ESTOQUE DE VENDA (ANTIGO)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPosicaoEstoque.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_SINTETICO_CMVPV, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO_CMVPV, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // POSIÇÃO DO ESTOQUE DE VENDA (CMV PV)" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelEstoqueVendaCmvPv.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE PRÉ-DEVOLUÇÃO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidoPreDevolucao.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_PRODUTOS_ESTOQUE_DEVOLUCAO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO: PRODUTOS NO ESTOQUE DE DEVOLUÇÃO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelEstoqueDevolucaoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_DEVOLUCAO_PRODUTOS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // DEVOLUÇÃO DE PRODUTOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelDevolucao.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_MEIO_DIVULGACAO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // PEDIDOS COLOCADOS CLASSIFICADOS PELO MEIO DE DIVULGAÇÃO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidosColocadosMidia.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_GERENCIAL_DE_VENDAS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO GERENCIAL DE VENDAS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelGerencialVendasFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_PESQUISA_INDICADORES, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // PESQUISA DE INDICADORES" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='PesquisaDeIndicadoresFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_CHECAGEM_NOVOS_PARCEIROS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE CHECAGEM DE NOVOS PARCEIROS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelChecagemNovosParceirosFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_DIVERGENCIA_CLIENTE_INDICADOR, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE DIVERGÊNCIA CLIENTE/INDICADOR" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelDivergenciaClienteIndicadorFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_METAS_INDICADOR, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE METAS DO INDICADOR" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelMetasIndicadorFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_REL_PERFORMANCE_INDICADOR, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE PERFORMANCE POR INDICADOR" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPerformanceIndicadorFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_VENDAS_POR_BOLETO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE VENDAS POR BOLETO" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelVendasPorBoletoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

	if operacao_permitida(OP_LJA_OCORRENCIAS_EM_PEDIDOS_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE OCORRÊNCIAS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidoOcorrenciaExec.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_LJA_REL_ESTATISTICAS_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // Relatório de Estatísticas de Ocorrências " & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RElPedidoOcorrenciaEstatisticasFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_LJA_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // ACOMPANHAMENTO DE CHAMADOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelAcompanhamentoChamadosFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_LJA_REL_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE CHAMADOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidoChamadoFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_LJA_REL_ESTATISTICAS_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // RELATÓRIO DE ESTATÍSTICAS DE CHAMADOS" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidoChamadoEstatisticasFiltro.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

     if operacao_permitida(OP_LJA_REL_INDICADORES_SEM_ATIVIDADE_RECENTE, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // Relatório de Indicadores Sem Atividade Recente " & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelIndicadoresSemAtivRec.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if
     if operacao_permitida(OP_LJA_REL_PEDIDOS_CANCELADOS, s_lista_operacoes_permitidas) then
	    strScript = strScript & _
		    " // Relatório de Pedidos Cancelados " & chr(13) & _
		    "	iop++;" & chr(13) & _
		    "	if (f.rb_rel[iop].checked) {" & chr(13) & _
		    "		s_dest='RelPedidoCancelado.asp';" & chr(13) & _
		    "		}" & chr(13) & _
		    "" & chr(13)
	 end if

    if operacao_permitida(OP_LJA_REL_PEDIDO_MARKETPLACE_NAO_RECEBIDO_CLIENTE, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // REGISTRO DE PEDIDOS DE MARKETPLACE NÃO RECEBIDOS PELO CLIENTE" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidosMktplaceNaoRecebidos.asp';" & chr(13) & _
			"		}" & chr(13) & _
			"" & chr(13)
		end if

    if operacao_permitida(OP_LJA_REL_REGISTRO_PEDIDO_MARKETPLACE_RECEBIDO_CLIENTE, s_lista_operacoes_permitidas) then
		strScript = strScript & _
			" // REGISTRO DE PEDIDOS DE MARKETPLACE RECEBIDOS PELO CLIENTE" & chr(13) & _
			"	iop++;" & chr(13) & _
			"	if (f.rb_rel[iop].checked) {" & chr(13) & _
			"		s_dest='RelPedidosMktplaceRecebidos.asp';" & chr(13) & _
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__E_JANELABUSCACEP_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS %>" rel="Stylesheet" type="text/css" />


<style type="text/css">
.tdn_reg{
    background:#FFF;
    text-align:right;
    border-right: 1pt solid #C0C0C0;
    border-left: 1pt solid #C0C0C0;
}
.tdQtde{
    width:60px;
    text-align:right;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdPed{
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdVen{
    width:80px;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdAp{
    width:80px;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}	  
.tdInd{
    width:80px;
    text-align:left;
    vertical-align:top;
}
.tdData{
    border-right:0px;
    text-align:center;   
}
.tdReg{
    width:5%;
    background:#FFF;
    text-align:right;
    border-right: 1pt solid #C0C0C0;
    border-left: 1pt solid #C0C0C0;
}	 
.tdDataF{
    width:10%;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdAnalise{
    width:35%;
    text-align:left;
    vertical-align:top;
} 	     
.tdPedido{
    width:10%;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdCliente{
    width:40%;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}      
#divQuadroOrcamentoWrapper
{
	left:1px;
	position:absolute;
	margin-left:1px;
	width:110px;
	z-index:0;
}
#divQuadroOrcamento
{
<%if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then%>
	margin-top:125px;
<%else%>
    margin-top:152px;
<%end if%>
	border: 1px solid #A9A9A9;
	padding-top: 4px;
	padding-bottom: 4px;
	padding-left: 6px;
	padding-right: 6px;
	position: absolute;
	background-color: #F5F5F5;
	top:0;
	z-index:0;
    width: 170px;
    text-align: center;
}
#divQuadroOrcamento.divFixo
{
	position:fixed;
	top:0;
}
#divAjaxRunning
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	z-index:1001;
	background-color:grey;
	opacity: .6;
}
.AjaxImgLoader
{
	position: absolute;
	left: 50%;
	top: 50%;
	margin-left: -128px; /* -1 * image width / 2 */
	margin-top: -128px;  /* -1 * image height / 2 */
	display: block;
}
</style>
<style>
	.ui-autocomplete {
		max-height: 200px;
        max-width: 496px;
		overflow-y: auto;
		overflow-x: hidden;
	}

</style>

<body id="corpoPagina" link="navy" alink="navy" vlink="navy"
<% if strMensagemAvisoPopUp <> "" then %>
onload="alert('<%=strMensagemAvisoPopUp%>');"
<%end if%>
>
<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>

    <!-- PopUp Quadro de Avisos -->
       <form name="fAVISO" id="fAVISO" method="post" action="QuadroAvisoLido.asp">
        <input type="hidden" name="aviso_selecionado" id="aviso_selecionado" value=''>
        <input type="hidden" class="CBOX" name="xMsg" id="xMsg" value="">
        <div id="divTelaCheia" style="width:100%;height:100%;position:fixed;left:0;top:0;display:none;background-color:#000;opacity:0.3"></div>
        <div id="divQuadroAvisoPai" style="width:1000px;height:65%;overflow:visible;position:fixed;top:50%;left:50%;right:0;margin-top:-330px;margin-left:-500px; border:4px solid #000">
        <a href="javascript:fechaQuadroAviso(fAVISO, 0);" title="Fechar" style="font-size:40pt;font-weight:bolder;color:#555;position:relative;left:970px;top:-50px;margin:0;z-index:100;">
            <img src="../IMAGEM/close_button_32.png" title="Fechar" style="border:0" />
        </a>
        <div id="divQuadroAviso" style="background-color:#fff;width:1000px;height:100%;overflow:scroll;position:absolute;top:0;left:0;right:0;bottom:0;margin:auto;border:1px solid #000;display:none">
            <div id="divQuadroAvisoConteudo" style="position:relative;height:auto;width:650px;top:10px;left:0;right:0;margin:auto;z-index:200;padding:0;"></div>
            <div name='dREMOVE' id='dREMOVE'><a href="javascript:RemoveAviso(fAVISO, 1);">
		    <img src="../botao/remover.gif" width="176" height="55" border="0" style="position:relative;bottom:0px;right:0;left:0;margin:auto"/></a></div>
        </div>
    </div>
        </form>

    <!-- Quadro Orçamentos em Aberto -->


	<div id="divQuadroOrcamento" style="box-shadow: 1px 1px 1px #888888;width: 200px">
	<form action="pedido.asp" id="fPEDPESQ" name="fPEDPESQ" method="post" onsubmit="if (trim(fPEDPESQ.pedido_selecionado.value)=='')return false;">
	<INPUT type=HIDDEN name='SessionCtrlInfo' value='0x8b35d7bd6b31cb451b07c58351772f67a7875b07c3ab4d8d6b31df1fff8b6da545fb93dd0341276b056db55f976d014f278dd705f18dc90d49b9e99bdba9eb35fbd71f5197c5334b89f53165b3eb2569b3875393d50b41bfcf3371a5fd2b6fa3e75b'>
	<span class="Rf">PRÉ-PEDIDOS:</span><br />
    <div id="divQtdePrePedidos" style="width:100%; text-align:center;margin-top:10px">
        <a href="javascript:exibeOcultaListaPrePedidos();" title="clique para mostrar/esconder a lista de pré-pedidos">
        <span id="qtdePrePedidos" style="font-size: 24pt; font-weight:bold;color:#888"></span></a>
    </div>
	<input type="hidden" name="orcamento_selecionado" value="" />
	<br />
    <div id="divListaPrePedidos" style="width:100%;display:none">
	    
    </div>
        <a id="linkPrePedido" href="javascript:exibeOcultaListaPrePedidos();"><div id="seta" style="width:100%;text-align:center;"><span id="spnExpandirPrePedido" style="font-weight:bolder;font-size:x-small">&or;</span></div></a>
	</form>
	</div>


<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom">
	<% if Not blnExibeTrocaRapidaLoja then %>
		<p class="PEDIDO" style="margin-bottom:2px;"><% = loja_nome & " (" & loja & ")" %></p>
	<% else %>
		<form action="ConectaVerifica.asp" method="post" id="fTrocaRapidaLoja" name="fTrocaRapidaLoja" style="margin-bottom:6px;">
		<input type="hidden" name="TrocaRapidaLoja" id="TrocaRapidaLoja" value='S'>
		<input type="hidden" name="usuario" id="usuario" value='<%=usuario%>'>
		<input type="hidden" name="senha" id="senha" value='<%=senha%>'>
		<select class="CBOX" id="loja" name="loja" style="font-size:12pt;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" onchange="fTrocaRapidaLoja.submit();">
		<%=loja_troca_rapida_monta_itens_select(usuario, loja)%>
		</select>
		</form>
	<% end if %>
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
		</span>
	</td>
</tr>
</table>

<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->


<% if strMensagemAviso <> "" then %>
	<br><br>
	<span class="Lbl">AVISO</span>
	<div class='MtAlerta' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'><%=strMensagemAviso%></p></div>
	<br>
<% end if %>

<br>

<%
if operacao_permitida(OP_LJA_CADASTRA_NOVO_PEDIDO_EC_SEMI_AUTOMATICO, s_lista_operacoes_permitidas) Or _
	 operacao_permitida(OP_LJA_CADASTRA_NOVO_PEDIDO_EC_INDICADOR_SEMI_AUTOMATICO, s_lista_operacoes_permitidas) then
%>
<!--  NOVO PEDIDO DE E-COMMERCE (CADASTRAMENTO SEMI-AUTOMÁTICO)  -->
<%	s_operationControlTicket = ""
	s = "SELECT Convert(varchar(36), newid()) AS guid"
	if rs.State <> 0 then rs.Close
	rs.Open s,cn
	if Not rs.Eof then s_operationControlTicket = Trim("" & rs("guid"))
%>

<form action="PedidoNovoECConsiste.asp" method="post" id="fPNEC" name="fPNEC">
<input type="hidden" name="operationControlTicket" id="operationControlTicket" value="<%=s_operationControlTicket%>" />
<input type="hidden" name="sessionToken" id="sessionToken" value="<%=s_sessionToken%>" />
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<span id="sNOVOPEDEC" class="T">NOVO PEDIDO DO E-COMMERCE (SEMI-AUTOMÁTICO)</span>
<div class="QFn" align="center" style="width:600px">
	<p class="C" style="margin: 10 10 10 10">&nbsp;</p>
	<table cellpadding="0" cellspacing="0">
		<tr>
			<td class="R" align="right" nowrap>
				<p class="C" style="margin-top:5px;">Nº MAGENTO&nbsp;</p>
			</td>
			<td align="left">
				<input name="c_numero_magento" id="c_numero_magento" type="text" maxlength="9" size="20" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {this.value=trim(this.value); fPNEC.bPNECConcluir.click();} filtra_numerico();">
			</td>
		</tr>
	</table>
	<p class="R" style="margin: 4 10 0 10">&nbsp;</p>
	<table>
		<tr>
			<td align="center">
				<input name="bPNECConcluir" id="bPNECConcluir" type="button" class="Botao" onclick="ConsultaPedidoMagentoAjax(fPNEC);"
						value="PROSSEGUIR" title="cadastramento de pedido de e-commerce (semi-automático)">
			</td>
		</tr>
	</table>
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
</div>
</form>
<br />
<%
end if
%>


<%
if operacao_permitida(OP_LJA_CADASTRA_NOVO_PEDIDO, s_lista_operacoes_permitidas) then
%>
<!--  NOVO PEDIDO  -->
<form action="ClientePesquisa.asp" method="post" id="fCP" name="fCP" onsubmit="if (!fCPConcluir(fCP)) return false;">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span id="sNOVOPED" class="T">NOVO PEDIDO</span>
<div class="QFn" align="center" style="width:600px">

	<p class="C" style="margin: 10 10 10 10">PESQUISAR CLIENTE POR</p>
	<table cellpadding="0" cellspacing="0">
		<tr>
			<td class="R" align="right" nowrap>
				<p class="C" style="margin-top:5px;">CNPJ/CPF&nbsp;</p>
			</td>
			<td align="left">
				<input name="cnpj_cpf_selecionado" id="cnpj_cpf_selecionado" type="text" maxlength="18" size="20" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="if (digitou_enter(true) && tem_info(this.value) && cnpj_cpf_ok(this.value)) {this.value=cnpj_cpf_formata(this.value); fCPConcluir(fCP);} filtra_cnpj_cpf();">
			</td>
		</tr>
		<tr>
			<td class="R" align="right" nowrap>
				<p class="C" style="margin-top:5px;">NOME&nbsp;</p>
			</td>
			<td align="left">
				<input name="nome_selecionado" id="nome_selecionado" type="text" maxlength="60" size="45" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCPConcluir(fCP); filtra_nome_identificador();">
			</td>
		</tr>
		</table>
	<p class="R" style="margin: 4 10 0 10">&nbsp;</p>
	<table>
		<tr>
			<td align="left">
				<input name="bCADCLI" id="bCADCLI" type="button" class="Botao" 
					value="CADASTRO DE CLIENTES" title="operações no cadastro de clientes"
					onclick="window.location='cliente.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>'">
			</td>
			<td align="left">
				<span style="width:30px;"></span>
			</td>
			<td align="left">
				<input name="bPesqCliEXECUTAR" id="bPesqCliEXECUTAR" type="submit" class="Botao" 
						value="EXECUTAR CONSULTA" title="executa a pesquisa no cadastro de clientes">
			</td>
		</tr>
	</table>
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>
<% end if %>
<br />

<!--  CANCELAMENTO AUTOMÁTICO DE PEDIDO  -->
	<span class="T">CANCELAMENTO AUTOMÁTICO DE PEDIDO</span>
	<table width="600" class="QS" style="border:0px" cellspacing="0">
	<tr><td align="left"></td></tr>
	</table>

		
<form id="fRELCons" name="fRELCons" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
        <% Response.Write lista_cancelamento() %>			
</form>		

<!-- -------------->
<br />

<!--  INDICADORES SEM ATIVIDADE RECENTE   -->
	<span class="T">INDICADORES SEM ATIVIDADE RECENTE</span>
	<table width="600" class="QS" style="border:0px" cellspacing="0">
	<tr><td align="left"></td></tr>
	</table>

		
<form id="fRELConsulta" name="fRELConsulta" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
        <% Response.Write lista_indicadores() %>			
</form>		

<!-- -------------->

<%
if operacao_permitida(OP_LJA_CONSULTA_PEDIDO, s_lista_operacoes_permitidas) Or _
   operacao_permitida(OP_LJA_REL_PEDIDOS_CREDITO_PENDENTE, s_lista_operacoes_permitidas) Or _
   operacao_permitida(OP_LJA_REL_PEDIDOS_CREDITO_PENDENTE_VENDAS, s_lista_operacoes_permitidas) Or _
   operacao_permitida(OP_LJA_REL_PEDIDOS_PENDENTES_CARTAO, s_lista_operacoes_permitidas) Or _
   operacao_permitida(OP_LJA_PESQUISA_PEDIDO_POR_OBS2, s_lista_operacoes_permitidas) then
%>
	<br />
	<span class="T">CONSULTA PEDIDOS</span>
    <a name="ConsultaPedidos"></a>
	<table width="600" class="QS" style="border-left:0px;border-right:0px;" cellspacing="0">
	<tr><td align="left"></td></tr>
	</table>
	
	<%if operacao_permitida(OP_LJA_CONSULTA_PEDIDO, s_lista_operacoes_permitidas) then%>
		<!--  C O N S U L T A   P E D I D O S  -->
		<% if opcao_lista_pedidos <> "" then
				Response.Write lista_pedidos() %>
			<table width="600" class="QS" cellspacing="0">
				<tr class="DefaultBkg">
					<td align="center">
						<p class="R"><a href='resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>' onclick="javascript:window.status='Aguarde ...';">Ocultar lista dos últimos pedidos</a></p>
					</td>
				</tr>
			</table>
		<% else %>
			<table width="600" class="QS" cellspacing="0">
				<tr class="DefaultBkg">
					<td align="center">
						<p class="R"><a href='resumo.asp?opcao_lista_pedidos=S<%= "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>#ConsultaPedidos' onclick="javascript:window.status='Aguarde ...';">Exibir lista dos últimos pedidos</a></p>
					</td>
				</tr>
			</table>
		<% end if %>
	<%end if%>

	<% 
	if operacao_permitida(OP_LJA_REL_PEDIDOS_CREDITO_PENDENTE, s_lista_operacoes_permitidas) then
	%>
		<table width="600" class="QS" cellspacing="0">
			<% 
			if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
			%>
				<form action="RelPedidosCredPendFiltro.asp" method="post" id="fRelPedCredPend" name="fRelPedCredPend">
				<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
				<input type="hidden" name="vendedor_selecionado" id="vendedor_selecionado" value=''>
			<% else %>
				<form action="RelPedidosCredPendExec.asp" method="post" id="fRelPedCredPend" name="fRelPedCredPend">
				<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
				<input type="hidden" name="vendedor_selecionado" id="vendedor_selecionado" value='<%=usuario%>'>
			<% end if %>
			<tr class="DefaultBkg">
				<td align="center">
					<p class="R"><span class="R" style="color:navy;cursor:pointer;" onclick="javascript:window.status='Aguarde ...'; fRelPedCredPend.submit();">Pedidos com Crédito Pendente</span></p>
				</td>
			</tr>
			</form>
		</table>
	<% end if %>

	<% 
	if operacao_permitida(OP_LJA_REL_PEDIDOS_CREDITO_PENDENTE_VENDAS, s_lista_operacoes_permitidas) then
	%>
		<table width="600" class="QS" cellspacing="0">
			<% 
			if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
			%>
				<form action="RelPedidosCredPendVendasFiltro.asp" method="post" id="fRelPedCredPendVendas" name="fRelPedCredPendVendas">
				<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
				<input type="hidden" name="vendedor_selecionado" id="vendedor_selecionado" value=''>
			<% else %>
				<form action="RelPedidosCredPendVendasExec.asp" method="post" id="fRelPedCredPendVendas" name="fRelPedCredPendVendas">
				<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
				<input type="hidden" name="vendedor_selecionado" id="vendedor_selecionado" value='<%=usuario%>'>
			<% end if %>
			<tr class="DefaultBkg">
				<td align="center">
					<p class="R"><span class="R" style="color:navy;cursor:pointer;" onclick="javascript:window.status='Aguarde ...'; fRelPedCredPendVendas.submit();">Pedidos com Crédito Pendente Vendas</span></p>
				</td>
			</tr>
			</form>
		</table>
	<% end if %>

	<% 
	if operacao_permitida(OP_LJA_REL_PEDIDOS_PENDENTES_CARTAO, s_lista_operacoes_permitidas) then
	%>
		<table width="600" class="QS" cellspacing="0">
			<% 
			if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
			%>
				<form action="RelPedidosPendentesCartaoFiltro.asp" method="post" id="fRelPedPendCartao" name="fRelPedPendCartao">
				<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
				<input type="hidden" name="vendedor_selecionado" id="vendedor_selecionado" value=''>
			<% else %>
				<form action="RelPedidosPendentesCartaoExec.asp" method="post" id="fRelPedPendCartao" name="fRelPedPendCartao">
				<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
				<input type="hidden" name="vendedor_selecionado" id="vendedor_selecionado" value='<%=usuario%>'>
			<% end if %>
			<tr class="DefaultBkg">
				<td align="center">
					<p class="R"><span class="R" style="color:navy;cursor:pointer;" onclick="javascript:window.status='Aguarde ...'; fRelPedPendCartao.submit();">Pedidos Pendentes Cartão de Crédito</span></p>
				</td>
			</tr>
			</form>
		</table>
	<% end if %>

    <% 
	if operacao_permitida(OP_LJA_REL_PEDIDOS_CREDITO_PENDENTE_VENDAS, s_lista_operacoes_permitidas) then
	%>
		<table width="600" class="QS" cellspacing="0">
			<% 
			if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
			%>
				<form action="RelPedidosEnderecoPendenteFiltro.asp" method="post" id="fRelPedEndPendente" name="fRelPedEndPendente">
				<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
				<input type="hidden" name="vendedor_selecionado" id="vendedor_selecionado" value=''>
			<% else %>
				<form action="RelPedidosEnderecoPendenteExec.asp" method="post" id="fRelPedEndPendente" name="fRelPedEndPendente">
				<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
				<input type="hidden" name="vendedor_selecionado" id="vendedor_selecionado" value='<%=usuario%>'>
			<% end if %>
			<tr class="DefaultBkg">
				<td align="center">
					<p class="R"><span class="R" style="color:navy;cursor:pointer;" onclick="javascript:window.status='Aguarde ...'; fRelPedEndPendente.submit();">Pedidos com Endereço Pendente</span></p>
				</td>
			</tr>
			</form>
		</table>
	<% end if %>

	<%if operacao_permitida(OP_LJA_CONSULTA_PEDIDO, s_lista_operacoes_permitidas) then%>
		<table width="600" class="QS" cellspacing="0">
			<tr class="DefaultBkg">
				<td align="center"><p class="R"><a href='RelPedidosAnteriores.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>' onclick="javascript:window.status='Aguarde ...';">Pesquisa pedidos anteriormente efetuados por um cliente nesta loja</a></p></td>
			</tr>
		</table>
	<%end if%>

<table width="600" class="QS" cellspacing="0">
	<%if operacao_permitida(OP_LJA_CONSULTA_PEDIDO, s_lista_operacoes_permitidas) then%>
		<tr class="DefaultBkg">
			<td width="40%" align="left">
				<p class="Cd">Nº Pedido</p>
			</td>
			<td align="left">
				<form action="pedido.asp" method="post" id="fPED" name="fPED" style="margin:4px 0px 4px 0px;">
				<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
				<table cellspacing="0" width="100%">
					<tr>
						<td align="left" style="width:140px">
							<input maxlength="10" name="pedido_selecionado" id="pedido_selecionado" style="width:115px;margin-right:15px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {if (normaliza_numero_pedido_e_sufixo(this.value)!='') this.value=normaliza_numero_pedido_e_sufixo(this.value); fPED.submit();} filtra_pedido();" onblur="if (normaliza_numero_pedido_e_sufixo(this.value)!='') this.value=normaliza_numero_pedido_e_sufixo(this.value);">
						</td>
						<td align="left">
							<input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao" value="CONSULTAR" title="consulta o pedido">
						</td>
					</tr>
				</table>
				</form>
			</td>
		</tr>
	  <%end if%>
	
	<% if operacao_permitida(OP_LJA_PESQUISA_PEDIDO_POR_OBS2, s_lista_operacoes_permitidas) then %>
		<!--  C O N S U L T A   P E D I D O  P E L O   C A M P O   Nº Nota Fiscal -->
	<tr class="DefaultBkg">
		<td width="40%" align="left">
			<p class="Cd">Nº Nota Fiscal</p>
		</td>
		<td align="left">
			<form action="RelPesquisaPedidoNF.asp" method="post" id="fRNF" name="fRNF" style="margin:4px 0px 4px 0px;">
			<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
			<table cellspacing="0" width="100%">
				<tr>
					<td align="left" style='width:140px'>
						<input maxlength="10" name="c_nf" id="c_nf" style="width:115px;margin-right:15px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fRNF.submit();">
					</td>
					<td align="left">
						<input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao" value="CONSULTAR" title="consulta pedido pelo campo 'Nº Nota Fiscal'">
					</td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
	<% end if %>

	<% if isLojaBonshop(loja) And operacao_permitida(OP_LJA_CADASTRA_NOVO_PEDIDO_EC_INDICADOR_SEMI_AUTOMATICO, s_lista_operacoes_permitidas) then %>
    <!--  C O N S U L T A   P E D I D O   P E L O   N Ú M E R O   M A G E N T O -->
	<tr class="DefaultBkg">
		<td width="40%" align="left">
			<p class="Cd">Número Magento</p>
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
    <% elseif loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then %>
    <!--  C O N S U L T A   P E D I D O   P E L O   N Ú M E R O   M A G E N T O -->
	<tr class="DefaultBkg">
		<td width="40%" align="left">
			<p class="Cd">Número Magento</p>
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
			<p class="Cd">Número Marketplace</p>
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
	
</table>
<% end if %>



<% 
if operacao_permitida(OP_LJA_CONSULTA_ORCAMENTO, s_lista_operacoes_permitidas) then
%>
<!--  C O N S U L T A   O R Ç A M E N T O S  -->
<br />
<form action="orcamento.asp" method="post" id="fORC" name="fORC">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span class="T">CONSULTA PRÉ-PEDIDO</span>
<table width="600" class="Q" cellspacing="0">
	<tr class="DefaultBkg">
		<td align="center" class="MB"><p class="R"><a href='OrcamentosEmAberto.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>' onclick="javascript:window.status='Aguarde ...';">Pré-Pedido em aberto</a></p></td>
	</tr>
	<tr class="DefaultBkg">
		<td align="center"><p class="R"><a href='RelOrcamentosMCrit.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>' onclick="javascript:window.status='Aguarde ...';">Relatório multicritério de Pré-Pedido</a></p></td>
	</tr>
</table>

<table width="600" class="QS" cellspacing="0">
	<tr class="DefaultBkg">
		<td valign="middle" align="center">
			<p class="C" style="margin: 12px 0px 12px 0px;">Nº PRÉ-PEDIDO&nbsp;&nbsp;
				<input size="10" maxlength="10" name="orcamento_selecionado" id="orcamento_selecionado" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {if (normaliza_num_orcamento(this.value)!='') this.value=normaliza_num_orcamento(this.value); fORC.submit();} filtra_orcamento();" onblur="if (normaliza_num_orcamento(this.value)!='') this.value=normaliza_num_orcamento(this.value);">
				&nbsp;&nbsp;&nbsp;
				<input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao" 
							value="CONSULTAR" title="consulta um pré-pedido específico desta loja">
			</p>
		</td>
	</tr>
</table>
</form>
<% end if %>




<% 
if operacao_permitida(OP_LJA_CONSULTA_DISPONIBILIDADE_ESTOQUE, s_lista_operacoes_permitidas) then
%>
<!--  DISPONIBILIDADE NO ESTOQUE  -->
<br />
<form action="RelProdutoDisponivel.asp" method="post" id="fEST" name="fEST" onsubmit="if (!fESTConcluir(fEST)) return false;">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span id="sDISPEST" class="T">CONSULTA DISPONIBILIDADE NO ESTOQUE</span>
<div class="QFn" align="center" style="width:600px;">
	<table cellpadding="0" cellspacing="0" style="margin-top:10px;">
		<tr>
			<td class="R" align="left" nowrap>
				<p class="C" style="margin: 12px 0px 12px 0px;">FABRICANTE&nbsp;
				<input name="c_fabricante" id="c_fabricante" type="text" maxlength="4" style="width:40px;" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fEST.c_produto.focus(); filtra_fabricante();">
				&nbsp;&nbsp;&nbsp;&nbsp;PRODUTO&nbsp;
				<input name="c_produto" id="c_produto" type="text" maxlength="8" style="width:80px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {this.value=normaliza_produto(this.value); fESTConcluir(fEST);} filtra_produto();" onblur="this.value=normaliza_produto(this.value);">
				&nbsp;&nbsp;&nbsp;
				<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" value="CONSULTAR" title="consulta a disponibilidade do produto no estoque de venda">
				</p>
			</td>
		</tr>
	</table>
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
</div>
</form>
<% end if %>




<% 
if operacao_permitida(OP_LJA_CONSULTA_PROD_BLOQ_ESTOQUE_ENTREGA, s_lista_operacoes_permitidas) then
%>
<!--  PRODUTOS BLOQUEADOS NO ESTOQUE DE ENTREGA  -->
<br />
<form action="RelProdBloqueado.asp" method="post" id="fPRODBLOQ" name="fPRODBLOQ" onsubmit="if (!fPRODBLOQConcluir(fPRODBLOQ)) return false;">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span id="sPRODBLOQ" class="T">PRODUTOS "BLOQUEADOS" NO ESTOQUE DE ENTREGA</span>
<div class="QFn" align="center" style="width:600px;">
	<table cellpadding="0" cellspacing="0" style="margin-top:10px;">
		<tr>
			<td class="R" align="left" nowrap>
				<p class="C" style="margin: 12px 0px 12px 0px;">FABRICANTE&nbsp;
				<input name="c_fabricante" id="c_fabricante" type="text" maxlength="4" style="width:40px;" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPRODBLOQ.c_produto.focus(); filtra_fabricante();">
				&nbsp;&nbsp;&nbsp;&nbsp;PRODUTO&nbsp;
				<input name="c_produto" id="c_produto" type="text" maxlength="8" style="width:80px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {this.value=normaliza_produto(this.value); fPRODBLOQConcluir(fPRODBLOQ);} filtra_produto();" onblur="this.value=normaliza_produto(this.value);">
				&nbsp;&nbsp;&nbsp;
				<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" value="CONSULTAR" title="executa a consulta">
				</p>
			</td>
		</tr>
	</table>
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
</div>
</form>
<% end if %>




<% 
if operacao_permitida(OP_LJA_CONSULTA_LISTA_PRECOS, s_lista_operacoes_permitidas) then
%>
<!--  LISTA DE PREÇOS  -->
<br />
<form action="RelListaPrecos.asp" method="post" id="fLPRECOS" name="fLPRECOS" onsubmit="if (!fLPRECOSConcluir(fLPRECOS)) return false;">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span id="sLPRECOS" class="T">LISTA DE PREÇOS</span>
<div class="QFn" align="center" style="width:600px;">
	<table cellpadding="0" cellspacing="0" style="margin-top:10px;">
		<tr>
			<td class="R" align="left" nowrap>
				<p class="C" style="margin: 12px 0px 12px 0px;">FABRICANTE&nbsp;
				<input name="c_fabricante" id="c_fabricante" type="text" maxlength="4" style="width:40px;" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {this.value=normaliza_codigo(this.value, TAM_MIN_FABRICANTE); fLPRECOSConcluir(fLPRECOS);} filtra_fabricante();">
				&nbsp;&nbsp;&nbsp;
				<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" value="CONSULTAR" title="consulta a lista de preços do fabricante">
				</p>
			</td>
		</tr>
	</table>
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
</div>
</form>
<% end if %>


<% 
if operacao_permitida(OP_LJA_CONSULTA_TABELA_CUSTO_FINANCEIRO_FORNECEDOR, s_lista_operacoes_permitidas) then
%>
<!--  TABELA DE CUSTO FINANCEIRO  -->
<br />
<form action="RelTabelaCustoFinanceiro.asp" method="post" id="fCustoFinancFornec" name="fCustoFinancFornec" onsubmit="if (!fCustoFinancFornecConcluir(fCustoFinancFornec)) return false;">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span id="sCustoFinancFornec" class="T">TABELA DE CUSTO FINANCEIRO</span>
<div class="QFn" align="center" style="width:600px;">
	<table cellpadding="0" cellspacing="0" style="margin-top:10px;">
		<tr>
			<td class="R" align="left" nowrap>
				<p class="C" style="margin: 12px 0px 12px 0px;">FABRICANTE&nbsp;
				<input name="c_fabricante" id="c_fabricante" type="text" maxlength="4" style="width:40px;" 
					onblur="this.value=normaliza_codigo(this.value, TAM_MIN_FABRICANTE);" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) {this.value=normaliza_codigo(this.value, TAM_MIN_FABRICANTE); fCustoFinancFornecConcluir(fCustoFinancFornec);} filtra_fabricante();">
				&nbsp;&nbsp;&nbsp;
				<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" 
					value="CONSULTAR" title="consulta a tabela de custo financeiro do fabricante">
				</p>
			</td>
		</tr>
	</table>
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
</div>
</form>
<% end if %>

<% 
if operacao_permitida(OP_LJA_EDITA_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) Or _
    operacao_permitida(OP_LJA_CONS_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) then
%>
<!--  INDICADORES  -->
<br />
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<span id="sIndicadores" class="T">INDICADORES</span>
<div class="QFn" align="center" style="width:600px;">
	<table cellpadding="0" cellspacing="0" style="margin-top:10px;">
		<tr class="DefaultBkg">
			<td width="25%" align="left">
				<p class="Cd">APELIDO</p>
			</td>
			<td align="left">
				<form method="post" id="fIndApelido" name="fIndApelido" style="margin:4px 0px 4px 0px;">
				<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
                <input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
                <input type="hidden" name="url_origem" id="url_origem" value="Resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" />
				<table cellspacing="0" width="100%">
					<tr>
						<td align="left" style="width:140px">
							<input id="id_selecionado" name="id_selecionado" maxlength="20" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fIndConsultarPorApelido(fIndApelido);" />
						</td>
						<td align="left">
                            <a href='javascript:fIndConsultarPorApelido(fIndApelido)'>
							<img src='../imagem/lupa_20x20.png' style='border:0;width:18px;height:18px;margin-right:5px' title='Consultar cadastro'></a>
                            <%if operacao_permitida(OP_LJA_EDITA_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) then%>
                            <a href='javascript:fIndEditarPorApelido(fIndApelido)'>
                            <img src='../imagem/edita_20x20.gif' style='border:0;width:20px;height:20px' title='Editar cadastro'></a>
                            <%end if%>
						</td>
					</tr>
				</table>
				</form>
			</td>
		</tr>
        <tr class="DefaultBkg">
			<td width="25%" align="left">
				<p class="Cd">CPF/CNPJ</p>
			</td>
			<td align="left">
				<form method="post" id="fIndCpfCnpj" name="fIndCpfCnpj" style="margin:4px 0px 4px 0px;">
				<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
                <input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
                <input type="hidden" name="url_origem" id="url_origem" value="Resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" />
				<table cellspacing="0" width="100%">
					<tr>
						<td align="left" style="width:140px">
							<input maxlength="18" name="cpf_cnpj_selecionado" id="cpf_cnpj_selecionado" style="width:180px;margin-right:10px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fIndConsultarPorCpfCnpj(fIndCpfCnpj);" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" />
						</td>
						<td align="left">
							<a href='javascript:fIndConsultarPorCpfCnpj(fIndCpfCnpj)'>
							<img src='../imagem/lupa_20x20.png' style='border:0;width:18px;height:18px;margin-right:5px' title='Consultar cadastro'></a>
                            <%if operacao_permitida(OP_LJA_EDITA_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) then%>
                            <a href='javascript:fIndEditarPorCpfCnpj(fIndCpfCnpj)'>
                            <img src='../imagem/edita_20x20.gif' style='border:0;width:20px;height:20px' title='Editar cadastro'></a>
                            <%end if%>
						</td>
					</tr>
				</table>
				</form>
			</td>
		</tr>
	</table>
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
</div>
<% end if %>



<!--  ***********************************************************************************************  -->
<!--  R E L A T Ó R I O S                         												       -->
<!--  ***********************************************************************************************  -->
<%
	qtde_relatorios = 0
	if operacao_permitida(OP_LJA_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
    if operacao_permitida(OP_LJA_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
    if operacao_permitida(OP_LJA_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_FATURAMENTO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_FATURAMENTO_CMVPV, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_VENDAS, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_VENDAS_COM_DESCONTO_SUPERIOR, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_MULTICRITERIO_PEDIDOS_ANALITICO, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_LJA_REL_MULTICRITERIO_PEDIDOS_SINTETICO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_PEDIDOS_COLOCADOS_NO_MES, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_COMISSAO_VENDEDORES, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_SINTETICO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_ANALITICO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_SINTETICO, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_SINTETICO_CMVPV, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO_CMVPV, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
    if operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) then
        qtde_relatorios=qtde_relatorios+1
        end if
	if operacao_permitida(OP_LJA_REL_PRODUTOS_ESTOQUE_DEVOLUCAO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_DEVOLUCAO_PRODUTOS, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_MEIO_DIVULGACAO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_GERENCIAL_DE_VENDAS, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_PESQUISA_INDICADORES, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_CHECAGEM_NOVOS_PARCEIROS, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_DIVERGENCIA_CLIENTE_INDICADOR, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_METAS_INDICADOR, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_REL_PERFORMANCE_INDICADOR, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_VENDAS_POR_BOLETO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
	if operacao_permitida(OP_LJA_OCORRENCIAS_EM_PEDIDOS_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
    if operacao_permitida(OP_LJA_REL_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
    if operacao_permitida(OP_LJA_REL_ESTATISTICAS_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
    if operacao_permitida(OP_LJA_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		qtde_relatorios=qtde_relatorios+1
		end if
%>
<% if qtde_relatorios > 0 then %>
<form method="post" id="fREL" name="fREL" onsubmit="if (!fRELConcluir(fREL)) return false">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FORÇA A CRIAÇÃO DE UM ARRAY DE RADIO BUTTONS MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" name="rb_rel" id="rb_rel" class="CBOX" value="">
<input type="hidden" name="pagina_destino" id="pagina_destino" value="">
<input type="hidden" name="titulo_relatorio" id="titulo_relatorio" value="">
<input type="hidden" name="filtro_obrigatorio_data_inicio" id="filtro_obrigatorio_data_inicio" value="">
<input type="hidden" name="filtro_obrigatorio_data_termino" id="filtro_obrigatorio_data_termino" value="">

<br />
<span id="sREL" class="T">RELATÓRIOS</span>
<div id="dREL" class="QFn" align="center" style="width:600px;">
<table cellpadding="0" cellspacing="0" style='margin:6px 0px 10px 0px;'>
	<tr>
		<td align="left" nowrap>
			<div style='margin-left:4px;'>
	<%  idx = 0 
		s_separacao = "" %>
		
	<%	' RELATÓRIO: COMISSÃO AOS INDICADORES (ALTERADO P/ RELATÓRIO DE PEDIDOS INDICADORES)
		if operacao_permitida(OP_LJA_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
			idx=idx+1 
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
	
	<% dim s_saida_default_rel, s_checked_rel
	s_saida_default_rel = get_default_valor_texto_bd(usuario, "RelComissaoIndicadores|c_carrega_indicadores_estatico")
	s_checked_rel = ""
		if (InStr(s_saida_default_rel, "ON") <> 0) then s_checked_rel = " checked" %>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Indicadores</span>
				<input type="checkbox" name="ckb_carrega_indicadores_rel"  id="ckb_carrega_indicadores_rel" value="ON" <%=s_checked_rel %> />
				<img src="../IMAGEM/exclamacao_14x14.png" id="Img1" style="cursor:pointer" title="Marque esta opção para que as listas de seleção no filtro sejam exibidas no modo estático" />
	
	<% end if %>
	
       <%	' RELATÓRIO: DE PEDIDOS INDICADORES (PROCESSADO)

		 if operacao_permitida(OP_LJA_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%> 
    
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Indicadores (Processado)</span>
   
	<% end if %>

    	<%	'RELATÓRIO DE PEDIDOS INDICADORES (Preview)

		 if operacao_permitida(OP_LJA_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
            if Day(Date) <= PRAZO_ACESSO_REL_PEDIDOS_INDICADORES_LOJA then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%> 
    
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Indicadores (Preview)</span>
   
	<% end if %>
    <% end if %>

	<%	' RELATÓRIO: FATURAMENTO (ANTIGO)
		if operacao_permitida(OP_LJA_REL_FATURAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Faturamento (Antigo)</span>
	<% end if %>

	<%	' RELATÓRIO: FATURAMENTO (CMV PV)
		if operacao_permitida(OP_LJA_REL_FATURAMENTO_CMVPV, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Faturamento</span>
	<% end if %>

	<%	' RELATÓRIO: Vendas
		if operacao_permitida(OP_LJA_REL_VENDAS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Vendas</span>
	<% end if %>

	<%	' RELATÓRIO: VENDAS COM DESCONTO SUPERIOR
		if operacao_permitida(OP_LJA_REL_VENDAS_COM_DESCONTO_SUPERIOR, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Vendas com Desconto Superior</span>
	<% end if %>

	<%	' RELATÓRIO: MULTICRITÉRIO DE PEDIDOS
	dim s_saida_default, s_checked
	s_saida_default = get_default_valor_texto_bd(usuario, "RelPedidosMCrit|c_carrega_indicadores_estatico")
		if operacao_permitida(OP_LJA_REL_MULTICRITERIO_PEDIDOS_ANALITICO, s_lista_operacoes_permitidas) Or _
		   operacao_permitida(OP_LJA_REL_MULTICRITERIO_PEDIDOS_SINTETICO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
			s_checked = ""
		if (InStr(s_saida_default, "ON") <> 0) then s_checked = " checked"
	%>
			<input type="radio" id="Radio1" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório Multicritério de Pedidos</span>
				<input type="checkbox" name="ckb_carrega_indicadores"  id="ckb_carrega_indicadores" value="ON" <%=s_checked %> />
				<img src="../IMAGEM/exclamacao_14x14.png" id="exclamacao" style="cursor:pointer" title="Marque esta opção para que as listas de seleção no filtro sejam exibidas no modo estático" />
	
	<% end if %>

	<%	' RELATÓRIO: Pedidos Colocados no Mês
		if operacao_permitida(OP_LJA_REL_PEDIDOS_COLOCADOS_NO_MES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Pedidos Colocados no Mês</span>
	<% end if %>

	<%	' RELATÓRIO: Comissão aos Vendedores
		if operacao_permitida(OP_LJA_REL_COMISSAO_VENDEDORES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Comissão aos Vendedores</span>
	<% end if %>

	<%	' RELATÓRIO: COMISSÃO AOS VENDEDORES SINTÉTICO (TABELA PROGRESSIVA)
		if operacao_permitida(OP_LJA_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_SINTETICO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Comissão aos Vendedores Sintético (Tabela Progressiva)</span>
	<% end if %>

	<%	' RELATÓRIO: COMISSÃO AOS VENDEDORES ANALÍTICO (TABELA PROGRESSIVA)
		if operacao_permitida(OP_LJA_REL_COMISSAO_VENDEDORES_TABELA_PROGRESSIVA_ANALITICO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Comissão aos Vendedores Analítico (Tabela Progressiva)</span>
	<% end if %>

	<%	' RELATÓRIO: Estoque (ANTIGO)
		if operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_SINTETICO, s_lista_operacoes_permitidas) Or _
		   operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Estoque de Venda (Antigo)</span>
	<% end if %>

	<%	' RELATÓRIO: Estoque (CMV PV)
		if operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_SINTETICO_CMVPV, s_lista_operacoes_permitidas) Or _
		   operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO_CMVPV, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Estoque de Venda</span>
	<% end if %>

    <%	' RELATÓRIO DE PRÉ-DEVOLUÇÃO
		if operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) then
			idx=idx+1
            Response.Write s_separacao
			s_separacao = "<br>"
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pré-Devoluções</span>
	<% end if %>

	<%	' RELATÓRIO: Produtos no Estoque de Devolução
		if operacao_permitida(OP_LJA_REL_PRODUTOS_ESTOQUE_DEVOLUCAO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Produtos no Estoque de Devolução</span>
	<% end if %>

	<%	' RELATÓRIO: Devolução de Produtos
		if operacao_permitida(OP_LJA_REL_DEVOLUCAO_PRODUTOS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Devolução de Produtos</span>
	<% end if %>

	<%	' RELATÓRIO: Meio de Divulgação
		if operacao_permitida(OP_LJA_REL_MEIO_DIVULGACAO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Meio de Divulgação</span>
	<% end if %>

	<%	' RELATÓRIO: Gerencial de Vendas
		if operacao_permitida(OP_LJA_REL_GERENCIAL_DE_VENDAS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório Gerencial de Vendas</span>
	<% end if %>

	<%	' PESQUISA DE INDICADORES
		if operacao_permitida(OP_LJA_PESQUISA_INDICADORES, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Pesquisa de Indicadores</span>
	<% end if %>

	<%	' RELATÓRIO DE CHECAGEM DE NOVOS PARCEIROS
		if operacao_permitida(OP_LJA_REL_CHECAGEM_NOVOS_PARCEIROS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Checagem de Novos Parceiros</span>
	<% end if %>

	<%	' RELATÓRIO DE DIVERGÊNCIA CLIENTE/INDICADOR
		if operacao_permitida(OP_LJA_REL_DIVERGENCIA_CLIENTE_INDICADOR, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Divergência Cliente/Indicador</span>
	<% end if %>

	<%	' RELATÓRIO DE METAS DO INDICADOR
		if operacao_permitida(OP_LJA_REL_METAS_INDICADOR, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Metas do Indicador</span>
	<% end if %>

	<%	' RELATÓRIO: PERFORMANCE POR INDICADOR
		if operacao_permitida(OP_LJA_REL_PERFORMANCE_INDICADOR, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Performance por Indicador</span>
	<% end if %>

	<%	' RELATÓRIO: VENDAS POR BOLETO
		if operacao_permitida(OP_LJA_VENDAS_POR_BOLETO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Vendas por Boleto</span>
	<% end if %>

	<%	' RELATÓRIO: OCORRÊNCIAS
		if operacao_permitida(OP_LJA_OCORRENCIAS_EM_PEDIDOS_CADASTRAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Ocorrências</span>
	<% end if %>

    <%	' RELATÓRIO: Relatório de Estatísticas de Ocorrências 
		if operacao_permitida(OP_LJA_REL_ESTATISTICAS_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Estatísticas de Ocorrências </span>
	<% end if %>

    <%	' ACOMPANHAMENTO DE CHAMADOS
		if operacao_permitida(OP_LJA_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Acompanhamento de Chamados</span>
	<% end if %>

    <%	' RELATÓRIO DE CHAMADOS
		if operacao_permitida(OP_LJA_REL_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Chamados</span>
	<% end if %>

    <%	' RELATÓRIO DE ESTATÍSTICAS DE CHAMADOS
		if operacao_permitida(OP_LJA_REL_ESTATISTICAS_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Estatísticas de Chamados</span>
	<% end if %>

     <%	' RELATÓRIO: INDICADORES SEM ATIVIDADES RECENTES
		if operacao_permitida(OP_LJA_REL_INDICADORES_SEM_ATIVIDADE_RECENTE, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Indicadores sem Atividade Recente </span>
	<% end if %>

    <%	' RELATÓRIO: PEDIDOS CANCELADOS
		if operacao_permitida(OP_LJA_REL_PEDIDOS_CANCELADOS , s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Relatório de Pedidos Cancelados </span>
	<% end if %>

    <%	' RELATÓRIO: REGISTRO DE PEDIDOS DE MARKETPLACE NÃO RECEBIDOS PELO CLIENTE
		if operacao_permitida(OP_LJA_REL_PEDIDO_MARKETPLACE_NAO_RECEBIDO_CLIENTE, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Registro de Pedidos de Marketplace Não Recebidos Pelo Cliente</span>
	<% end if %>

    <%	' RELATÓRIO: REGISTRO DE PEDIDOS DE MARKETPLACE RECEBIDOS PELO CLIENTE
		if operacao_permitida(OP_LJA_REL_REGISTRO_PEDIDO_MARKETPLACE_RECEBIDO_CLIENTE, s_lista_operacoes_permitidas) then
			idx=idx+1
			Response.Write s_separacao
			s_separacao = "<br>" 
			if (qtde_relatorios = 1) then s=" checked" else s=""
	%>
			<input type="radio" id="rb_rel" name="rb_rel" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fREL.rb_rel[<%=Cstr(idx)%>].click(); if (fREL.rb_rel[<%=Cstr(idx)%>].checked) fREL.bEXECUTAR.click();"
				>Registro de Pedidos de Marketplace Recebidos Pelo Cliente</span>
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
<% end if %>



<% 
if operacao_permitida(OP_LJA_LER_AVISOS_NAO_LIDOS, s_lista_operacoes_permitidas) Or _
   operacao_permitida(OP_LJA_LER_AVISOS_TODOS, s_lista_operacoes_permitidas) Or _
   operacao_permitida(OP_LJA_CADASTRA_SENHA_DESCONTO, s_lista_operacoes_permitidas) Or _
   operacao_permitida(OP_LJA_CADASTRAMENTO_INDICADOR, s_lista_operacoes_permitidas) Or _
   operacao_permitida(OP_LJA_CONS_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) Or _
   operacao_permitida(OP_LJA_EDITA_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) Or _
   operacao_permitida(OP_LJA_PREENCHER_INDICADOR_EM_PEDIDO_CADASTRADO, s_lista_operacoes_permitidas) then
%>
<!--  ***********************************************************************************************  -->
<!--  O U T R A S   F U N Ç Õ E S                         										       -->
<!--  ***********************************************************************************************  -->
<form method="post" id="fOF" name="fOF" onsubmit="if (!fOFConcluir(fOF)) return false">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="opcao_selecionada" id="opcao_selecionada" value="">
<input type="hidden" name="opcao_alerta_se_nao_ha_aviso" id="opcao_alerta_se_nao_ha_aviso" value="S">

<br />
<span class="T">OUTRAS FUNÇÕES</span>
<div class="QFn" align="center" style="width:600px;">
<table class="TFn">
	<tr>
		<td align="left" nowrap>
			<% idx = 0 %>
			
			<% idx=idx+1 %>
			<% if operacao_permitida(OP_LJA_LER_AVISOS_NAO_LIDOS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOF.rb_op[<%=Cstr(idx-1)%>].click(); fOF.bEXECUTAR.click();"
				>Ler Quadro de Avisos (somente não lidos)</span><br>
			
			<% idx=idx+1 %>
			<% if operacao_permitida(OP_LJA_LER_AVISOS_TODOS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOF.rb_op[<%=Cstr(idx-1)%>].click(); fOF.bEXECUTAR.click();"
				>Ler Quadro de Avisos (todos os avisos)</span><br>
			
			<% idx=idx+1 %>
			<% if operacao_permitida(OP_LJA_CADASTRA_SENHA_DESCONTO, s_lista_operacoes_permitidas) Or _
				  operacao_permitida(OP_LJA_CADASTRAMENTO_INDICADOR, s_lista_operacoes_permitidas) Or _
				  operacao_permitida(OP_LJA_CONS_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) Or _
				  operacao_permitida(OP_LJA_EDITA_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) Or _
				  operacao_permitida(OP_LJA_PREENCHER_INDICADOR_EM_PEDIDO_CADASTRADO, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOF.rb_op[<%=Cstr(idx-1)%>].click(); fOF.bEXECUTAR.click();"
				>Funções Administrativas</span>
		</td>
	</tr>
</table>

	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" value="EXECUTAR" title="executa">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>
<% end if %>

<br />

</center>

</body>
</html>

<%

'	FECHA CONEXAO COM O BANCO DE DADOS
'	Obs.: Para que o fechamento seja imediato é necessário acertar
'		  o registro do IIS 4.0, desabilitando o "connection pooling".
'		  Ver artigo no MSDN (ID: Q189410)
	cn.Close
	set cn = nothing
%>
