<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  U S U A R I O A T U A L I Z A . A S P
'     =====================================
'
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
'			I N I C I A L I Z A     P Á G I N A     A S P     N O     S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
	dim s, s_aux, usuario, senha_cripto, alerta, chave
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r, rs, rsi
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim criou_novo_reg
	Dim s_log, s_log_perfil, s_log_perfil_anterior, s_log_loja_vendedor, s_log_loja_vendedor_anterior, s_log_usuario_x_cd, s_log_usuario_x_cd_anterior
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	s_log_perfil = ""
	s_log_perfil_anterior = ""
	s_log_loja_vendedor = ""
	s_log_loja_vendedor_anterior = ""
	s_log_usuario_x_cd = ""
	s_log_usuario_x_cd_anterior = ""
	campos_a_omitir = "|dt_ult_atualizacao|usuario_ult_atualizacao|timestamp|"
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim i, n
	dim s_usuario, s_senha, s_senha2, s_senha_original, s_nome, s_email, s_ddd, s_telefone, s_bloqueado, s_vendedor, operacao_selecionada, s_vendedor_ext, ckb_desbloquear_bloqueio_automatico
	operacao_selecionada=request("operacao_selecionada")
	s_usuario=UCase(trim(request("usuario_selecionado")))
	s_senha=trim(request("senha"))
	s_senha2=trim(request("senha2"))
	s_nome=trim(request("nome"))
    s_email=trim(request("email"))
	s_ddd = trim(request("ddd"))
	s_telefone = trim(request("telefone"))
	s_bloqueado=trim(request("rb_estado"))
	s_vendedor=trim(request("ckb_vendedor"))
	s_vendedor_ext=trim(request("ckb_vendedor_ext"))
	ckb_desbloquear_bloqueio_automatico=trim(request("ckb_desbloquear_bloqueio_automatico"))

'	SE FOR VENDEDOR DA LOJA, ARMAZENA RELAÇÃO DE LOJAS LIBERADAS
	dim qtde_loja_vendedor, v_loja_vendedor
	qtde_loja_vendedor = 0
	redim v_loja_vendedor(0)
	v_loja_vendedor(0) = ""
	n = Request.Form("ckb_loja_vendedor").Count 
	for i = 1 to n
		s = Trim(Request.Form("ckb_loja_vendedor")(i))
		if s <> "" then
			if Trim(v_loja_vendedor(ubound(v_loja_vendedor))) <> "" then
				redim preserve v_loja_vendedor(ubound(v_loja_vendedor)+1)
				v_loja_vendedor(ubound(v_loja_vendedor)) = ""
				end if
			v_loja_vendedor(ubound(v_loja_vendedor)) = s
			qtde_loja_vendedor = qtde_loja_vendedor + 1
			end if
		next
	
	
'	CONSISTÊNCIA
	if s_usuario = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)

	dim v_perfil, qtde_perfil
	qtde_perfil = 0
	redim v_perfil(0)
	v_perfil(0) = ""
	n = Request.Form("ckb_perfil").Count
	for i = 1 to n
		s = Trim(Request.Form("ckb_perfil")(i))
		if s <> "" then
			if Trim(v_perfil(ubound(v_perfil))) <> "" then
				redim preserve v_perfil(ubound(v_perfil)+1)
				v_perfil(ubound(v_perfil)) = ""
				end if
			v_perfil(ubound(v_perfil)) = s
			qtde_perfil = qtde_perfil + 1
			end if
		next
	
	dim v_usuario_x_cd
	redim v_usuario_x_cd(0)
	v_usuario_x_cd(0) = ""
	n = Request.Form("ckb_usuario_x_cd").Count
	for i = 1 to n
		s = Trim(Request.Form("ckb_usuario_x_cd")(i))
		if s <> "" then
			if Trim(v_usuario_x_cd(ubound(v_usuario_x_cd))) <> "" then
				redim preserve v_usuario_x_cd(ubound(v_usuario_x_cd)+1)
				v_usuario_x_cd(ubound(v_usuario_x_cd)) = ""
				end if
			v_usuario_x_cd(ubound(v_usuario_x_cd)) = s
			end if
		next
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(rsi, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	alerta = ""
	
	dim blnSenhaAlterada
	blnSenhaAlterada = False

	if operacao_selecionada <> OP_INCLUI then
		s_senha_original = ""
		s = "SELECT * FROM t_USUARIO WHERE (usuario = '" & s_usuario & "')"
		if rs.State <> 0 then rs.Close
		rs.Open s, cn
		if rs.Eof then
			alerta="USUÁRIO NÃO ENCONTRADO (" & s_usuario & ")"
		else
			s = Trim("" & rs("datastamp"))
			chave = gera_chave(FATOR_BD)
			decodifica_dado s, s_senha_original, chave
			if s_senha_original <> "" then
				if s_senha_original <> s_senha then blnSenhaAlterada = True
				end if
			end if
		end if

	if alerta = "" then
		if s_usuario = "" then
			alerta="IDENTIFICADOR DE USUÁRIO INVÁLIDO."	
		elseif s_nome = "" then
			alerta="PREENCHA O NOME DO USUÁRIO."
		elseif s_bloqueado = "" then
			alerta="INFORME SE O USUÁRIO TEM ACESSO PERMITIDO OU BLOQUEADO."
		elseif (s_vendedor=ID_VENDEDOR) And (qtde_loja_vendedor=0) then
			alerta="INFORME A(S) LOJA(S) DO VENDEDOR."
		elseif (operacao_selecionada = OP_INCLUI) And (len(s_senha) < TAM_MIN_SENHA) then
			alerta="A SENHA DEVE POSSUIR NO MÍNIMO " & TAM_MIN_SENHA & " CARACTERES."
		elseif (operacao_selecionada = OP_INCLUI) And (Not (tem_digito(s_senha) And tem_letra(s_senha))) then
			alerta = "A SENHA DEVE CONTER NO MÍNIMO 1 LETRA E 1 DÍGITO NUMÉRICO"
		elseif blnSenhaAlterada And (len(s_senha) < TAM_MIN_SENHA) then
			alerta="A SENHA DEVE POSSUIR NO MÍNIMO " & TAM_MIN_SENHA & " CARACTERES."
		elseif blnSenhaAlterada And (Not (tem_digito(s_senha) And tem_letra(s_senha))) then
			alerta = "A SENHA DEVE CONTER NO MÍNIMO 1 LETRA E 1 DÍGITO NUMÉRICO"
		elseif Ucase(s_senha) <> Ucase(s_senha2) then
			alerta="A CONFIRMAÇÃO DA SENHA NÃO ESTÁ CORRETA."
		elseif qtde_perfil = 0 then
			alerta="NENHUM PERFIL DE ACESSO FOI SELECIONADO."
			end if
		end if 'if alerta = ""

	'VALIDAÇÃO SE O IDENTIFICADOR JÁ ESTÁ EM USO NO CADASTRO DE INDICADORES (ASSEGURA QUE NÃO EXISTA USUÁRIO E INDICADOR COM MESMO IDENTIFICADOR)
	if alerta = "" then
		if operacao_selecionada = OP_INCLUI then
			s = "SELECT apelido, cnpj_cpf, razao_social_nome FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido = '" & s_usuario & "'"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Not rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Não é possível usar o identificador '" & s_usuario & "' porque já está em uso no cadastro de orçamentista/indicador"
				end if
			rs.Close
			end if
		end if

	if alerta <> "" then erro_consistencia=True
		
		
	chave = gera_chave(FATOR_BD)
	codifica_dado s_senha, senha_cripto, chave
	
	Err.Clear
	
'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			s="SELECT COUNT(*) AS qtde FROM t_PEDIDO WHERE (vendedor = '" & s_usuario & "')"
			r.Open s, cn
		'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
			if Cstr(r("qtde")) > Cstr(0) then
				erro_fatal=True
				alerta = "USUÁRIO NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE PEDIDOS."
				end if
			r.Close 
			
			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_ORCAMENTO WHERE (vendedor = '" & s_usuario & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "USUÁRIO NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE ORÇAMENTOS."
					end if
				r.Close 
				end if

			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_DESCONTO WHERE (autorizador = '" & s_usuario & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "USUÁRIO NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE AUTORIZAÇÃO PARA DESCONTO SUPERIOR."
					end if
				r.Close
				end if
				
			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_ESTOQUE WHERE (usuario = '" & s_usuario & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "USUÁRIO NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO EM OPERAÇÃO DE ENTRADA NO ESTOQUE."
					end if
				r.Close
				end if

			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_ESTOQUE_MOVIMENTO WHERE (usuario = '" & s_usuario & "') OR (anulado_usuario = '" & s_usuario & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "USUÁRIO NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO EM OPERAÇÃO DE MOVIMENTAÇÃO DO ESTOQUE."
					end if
				r.Close
				end if
			
			if Not erro_fatal then
			'	INFO P/ LOG
				s="SELECT apelido FROM t_PERFIL INNER JOIN t_PERFIL_X_USUARIO ON t_PERFIL.id = t_PERFIL_X_USUARIO.id_perfil WHERE t_PERFIL_X_USUARIO.usuario='" & s_usuario & "' ORDER BY apelido"
				if r.State <> 0 then r.Close
				r.Open s, cn
				do while Not r.Eof
					if s_log_perfil <> "" then s_log_perfil = s_log_perfil & ", "
					s_log_perfil = s_log_perfil & Cstr(r("apelido"))
					r.MoveNext
					loop
				
				if s_log_perfil = "" then s_log_perfil = "(nenhum)"
				if s_log_perfil <> "" then s_log_perfil = "perfil=" & s_log_perfil
				
			'	INFO P/ LOG
				s="SELECT loja FROM t_USUARIO_X_LOJA WHERE usuario='" & s_usuario & "' ORDER BY CONVERT(smallint,loja)"
				if r.State <> 0 then r.Close
				r.Open s, cn
				do while Not r.Eof
					if s_log_loja_vendedor <> "" then s_log_loja_vendedor = s_log_loja_vendedor & ", "
					s_log_loja_vendedor = s_log_loja_vendedor & Trim("" & r("loja"))
					r.MoveNext
					loop
				
				if s_log_loja_vendedor = "" then s_log_loja_vendedor = "(nenhuma)"
				if s_log_loja_vendedor <> "" then s_log_loja_vendedor = "loja(s)=" & s_log_loja_vendedor
				
			'	INFO P/ LOG
				s="SELECT apelido FROM t_NFe_EMITENTE INNER JOIN t_USUARIO_X_NFe_EMITENTE ON t_NFe_EMITENTE.id = t_USUARIO_X_NFe_EMITENTE.id_nfe_emitente WHERE t_USUARIO_X_NFe_EMITENTE.usuario = '" & s_usuario & "' ORDER BY t_USUARIO_X_NFe_EMITENTE.id_nfe_emitente"
				if r.State <> 0 then r.Close
				r.Open s, cn
				do while Not r.Eof
					if s_log_usuario_x_cd <> "" then s_log_usuario_x_cd = s_log_usuario_x_cd & ", "
					s_log_usuario_x_cd = s_log_usuario_x_cd & Ucase(Trim("" & r("apelido")))
					r.MoveNext
					loop
				
				if s_log_usuario_x_cd = "" then s_log_usuario_x_cd = "(nenhum)"
				if s_log_usuario_x_cd <> "" then s_log_usuario_x_cd = "CD=" & s_log_usuario_x_cd
				
			'	INFO P/ LOG
				s="SELECT * FROM t_USUARIO WHERE usuario = '" & s_usuario & "'"
				if r.State <> 0 then r.Close
				r.Open s, cn
				if Not r.EOF then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					s_log = log_via_vetor_monta_exclusao(vLog1)
					end if
				r.Close
				
			'	APAGA!!
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				s="DELETE FROM t_PERFIL_X_USUARIO WHERE usuario = '" & s_usuario & "'"
				cn.Execute(s)
				If Err <> 0 then 
					erro_fatal=True
					alerta = "FALHA AO REMOVER PERFIL DE ACESSO DO USUÁRIO (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				
				if Not erro_fatal then
					s="DELETE FROM t_USUARIO_X_LOJA WHERE usuario = '" & s_usuario & "'"
					cn.Execute(s)
					If Err <> 0 then 
						erro_fatal=True
						alerta = "FALHA AO REMOVER RELAÇÃO DE LOJAS LIBERADAS P/ ACESSO DESTE USUÁRIO (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if Not erro_fatal then
					s="DELETE FROM t_USUARIO_X_NFe_EMITENTE WHERE usuario = '" & s_usuario & "'"
					cn.Execute(s)
					If Err <> 0 then
						erro_fatal=True
						alerta = "FALHA AO REMOVER RELAÇÃO DE CD'S LIBERADOS P/ ACESSO DESTE USUÁRIO (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if Not erro_fatal then
					s="DELETE FROM t_USUARIO WHERE usuario = '" & s_usuario & "'"
					cn.Execute(s)
					If Err = 0 then 
						if (s_log <> "") And (s_log_perfil <> "") then s_log = s_log & "; "
						s_log = s_log & s_log_perfil
						if (s_log <> "") And (s_log_loja_vendedor <> "") then s_log = s_log & "; "
						s_log = s_log & s_log_loja_vendedor
						if (s_log <> "") And (s_log_usuario_x_cd <> "") then s_log = s_log & "; "
						s_log = s_log & s_log_usuario_x_cd
						if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_USUARIO_EXCLUSAO, s_log
					else
						erro_fatal=True
						alerta = "FALHA AO REMOVER O USUÁRIO (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if alerta = "" then
				'	~~~~~~~~~~~~~~
					cn.CommitTrans
				'	~~~~~~~~~~~~~~
					if Err <> 0 then 
						alerta=Cstr(Err) & ": " & Err.Description
						erro_fatal = True
						end if
				else
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Err.Clear
					end if
				end if
				

		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			if alerta = "" then 
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				s = "SELECT * FROM t_USUARIO WHERE usuario = '" & s_usuario & "'"
				r.Open s, cn
				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("usuario")=s_usuario
					r("dt_cadastro") = Date
					r("nivel") = " "
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir

					s="SELECT apelido FROM t_PERFIL INNER JOIN t_PERFIL_X_USUARIO ON t_PERFIL.id = t_PERFIL_X_USUARIO.id_perfil WHERE t_PERFIL_X_USUARIO.usuario='" & s_usuario & "' ORDER BY apelido"
					if rsi.State <> 0 then rsi.Close
					rsi.Open s, cn
					do while Not rsi.Eof
						if s_log_perfil_anterior <> "" then s_log_perfil_anterior = s_log_perfil_anterior & ", "
						s_log_perfil_anterior = s_log_perfil_anterior & Cstr(rsi("apelido"))
						rsi.MoveNext
						loop
				
					if s_log_perfil_anterior = "" then s_log_perfil_anterior = "(nenhum)"
					if s_log_perfil_anterior <> "" then s_log_perfil_anterior = "perfil (anterior): " & s_log_perfil_anterior
					
					s="SELECT loja FROM t_USUARIO_X_LOJA WHERE usuario='" & s_usuario & "' ORDER BY CONVERT(smallint,loja)"
					if rsi.State <> 0 then rsi.Close
					rsi.Open s, cn
					do while Not rsi.Eof
						if s_log_loja_vendedor_anterior <> "" then s_log_loja_vendedor_anterior = s_log_loja_vendedor_anterior & ", "
						s_log_loja_vendedor_anterior = s_log_loja_vendedor_anterior & Trim("" & rsi("loja"))
						rsi.MoveNext
						loop
				
					if s_log_loja_vendedor_anterior = "" then s_log_loja_vendedor_anterior = "(nenhuma)"
					if s_log_loja_vendedor_anterior <> "" then s_log_loja_vendedor_anterior = "lojas (anterior): " & s_log_loja_vendedor_anterior

					s="SELECT apelido FROM t_NFe_EMITENTE INNER JOIN t_USUARIO_X_NFe_EMITENTE ON t_NFe_EMITENTE.id = t_USUARIO_X_NFe_EMITENTE.id_nfe_emitente WHERE t_USUARIO_X_NFe_EMITENTE.usuario = '" & s_usuario & "' ORDER BY t_USUARIO_X_NFe_EMITENTE.id_nfe_emitente"
					if rsi.State <> 0 then rsi.Close
					rsi.Open s, cn
					do while Not rsi.Eof
						if s_log_usuario_x_cd_anterior <> "" then s_log_usuario_x_cd_anterior = s_log_usuario_x_cd_anterior & ", "
						s_log_usuario_x_cd_anterior = s_log_usuario_x_cd_anterior & Ucase(Trim("" & rsi("apelido")))
						rsi.MoveNext
						loop
				
					if s_log_usuario_x_cd_anterior = "" then s_log_usuario_x_cd_anterior = "(nenhum)"
					if s_log_usuario_x_cd_anterior <> "" then s_log_usuario_x_cd_anterior = "CD (anterior): " & s_log_usuario_x_cd_anterior
					
					if rsi.State <> 0 then rsi.Close
					end if
					
				if s_vendedor <> "" then
					r("vendedor_loja")=1
					r("loja")=""
				else
					r("vendedor_loja")=0
					r("loja")=""
					end if
					
				r("nome")=s_nome
                r("email")=s_email
				r("ddd")=retorna_so_digitos(s_ddd)
				r("telefone")=retorna_so_digitos(s_telefone)
				r("bloqueado")=CLng(s_bloqueado)
				r("dt_ult_atualizacao") = Now
				
				if trim("" & r("datastamp"))<>senha_cripto then
					r("datastamp")=senha_cripto
					r("senha") = gera_senha_aleatoria
					r("dt_ult_alteracao_senha") = Null
					'Ao alterar a senha, assegura que um eventual bloqueio automático de login também seja desbloqueado
					r("StLoginBloqueadoAutomatico") = 0
					r("QtdeConsecutivaFalhaLogin") = 0
					end if

				if s_vendedor_ext <> "" then
					r("vendedor_externo")=1
				else
					r("vendedor_externo")=0
					end if
				
				if ckb_desbloquear_bloqueio_automatico <> "" then
					r("StLoginBloqueadoAutomatico") = 0
					r("QtdeConsecutivaFalhaLogin") = 0
					end if

				r.Update

			'	PERFIL
				If Err = 0 then 
					s = "UPDATE t_PERFIL_X_USUARIO SET excluido_status = 1 WHERE usuario = '" & s_usuario & "'"
					cn.Execute(s)
					end if

				if Err = 0 then
				'	PERFIL
					for i = Lbound(v_perfil) to Ubound(v_perfil)
						if Trim(v_perfil(i)) <> "" then
							s = "SELECT * FROM t_PERFIL_X_USUARIO WHERE (usuario = '" & s_usuario & "') AND (id_perfil = '" & Trim(v_perfil(i)) & "')"
							if rs.State <> 0 then rs.Close
							rs.Open s, cn
							if Not rs.Eof then
								rs("excluido_status") = 0
							else
								rs.AddNew
								rs("usuario") = s_usuario
								rs("id_perfil") = Trim(v_perfil(i))
								rs("dt_cadastro") = Date
								rs("usuario_cadastro") = usuario
								end if
							rs.Update
							end if
							if Err <> 0 then exit for
						next
					end if

				if Err = 0 then
					s = "DELETE FROM t_PERFIL_X_USUARIO WHERE (usuario = '" & s_usuario & "') AND (excluido_status <> 0)"
					cn.Execute(s)
					end if
			
			'	LOJAS
				If Err = 0 then 
					s = "UPDATE t_USUARIO_X_LOJA SET excluido_status = 1 WHERE usuario = '" & s_usuario & "'"
					cn.Execute(s)
					end if

				if Err = 0 then
				'	SE FOR VENDEDOR DA LOJA, INDICA AS LOJAS LIBERADAS P/ ACESSO
					for i = Lbound(v_loja_vendedor) to Ubound(v_loja_vendedor)
						if Trim(v_loja_vendedor(i)) <> "" then
							s = "SELECT * FROM t_USUARIO_X_LOJA WHERE (usuario = '" & s_usuario & "') AND (loja = '" & Trim(v_loja_vendedor(i)) & "')"
							if rs.State <> 0 then rs.Close
							rs.Open s, cn
							if Not rs.Eof then
								rs("excluido_status") = 0
							else
								rs.AddNew
								rs("usuario") = s_usuario
								rs("loja") = Trim(v_loja_vendedor(i))
								rs("dt_cadastro") = Date
								rs("usuario_cadastro") = usuario
								end if
							rs.Update
							end if
							if Err <> 0 then exit for
						next
					end if

				if Err = 0 then
					s = "DELETE FROM t_USUARIO_X_LOJA WHERE (usuario = '" & s_usuario & "') AND (excluido_status <> 0)"
					cn.Execute(s)
					end if
				
			'	CD'S
				If Err = 0 then 
					s = "UPDATE t_USUARIO_X_NFe_EMITENTE SET excluido_status = 1 WHERE usuario = '" & s_usuario & "'"
					cn.Execute(s)
					end if

				if Err = 0 then
				'	CD'S
					for i = Lbound(v_usuario_x_cd) to Ubound(v_usuario_x_cd)
						if Trim(v_usuario_x_cd(i)) <> "" then
							s = "SELECT * FROM t_USUARIO_X_NFe_EMITENTE WHERE (usuario = '" & s_usuario & "') AND (id_nfe_emitente = " & Trim(v_usuario_x_cd(i)) & ")"
							if rs.State <> 0 then rs.Close
							rs.Open s, cn
							if Not rs.Eof then
								rs("excluido_status") = 0
							else
								rs.AddNew
								rs("usuario") = s_usuario
								rs("id_nfe_emitente") = Trim(v_usuario_x_cd(i))
								rs("usuario_cadastro") = usuario
								end if
							rs("dt_ult_atualizacao") = Date
							rs("dt_hr_ult_atualizacao") = Now
							rs("usuario_ult_atualizacao") = usuario
							rs.Update
							end if
						if Err <> 0 then exit for
						next
					end if

				if Err = 0 then
					s = "DELETE FROM t_USUARIO_X_NFe_EMITENTE WHERE (usuario = '" & s_usuario & "') AND (excluido_status <> 0)"
					cn.Execute(s)
					end if
				
			'	LOG
				If Err = 0 then 
					for i=Lbound(v_perfil) to Ubound(v_perfil)
						if Trim(v_perfil(i)) <> "" then
							if s_log_perfil <> "" then s_log_perfil = s_log_perfil & ", "
							s_log_perfil = s_log_perfil & x_perfil_apelido(v_perfil(i))
							end if
						next
					
					if s_log_perfil = "" then s_log_perfil = "(nenhum)"
					if s_log_perfil <> "" then s_log_perfil = "perfil (atual): " & s_log_perfil

					for i=Lbound(v_loja_vendedor) to Ubound(v_loja_vendedor)
						if Trim(v_loja_vendedor(i)) <> "" then
							if s_log_loja_vendedor <> "" then s_log_loja_vendedor = s_log_loja_vendedor & ", "
							s_log_loja_vendedor = s_log_loja_vendedor & v_loja_vendedor(i)
							end if
						next
					
					if s_log_loja_vendedor = "" then s_log_loja_vendedor = "(nenhuma)"
					if s_log_loja_vendedor <> "" then s_log_loja_vendedor = "lojas (atual): " & s_log_loja_vendedor

					for i=Lbound(v_usuario_x_cd) to Ubound(v_usuario_x_cd)
						if Trim(v_usuario_x_cd(i)) <> "" then
							if s_log_usuario_x_cd <> "" then s_log_usuario_x_cd = s_log_usuario_x_cd & ", "
							s_log_usuario_x_cd = s_log_usuario_x_cd & obtem_apelido_empresa_NFe_emitente(v_usuario_x_cd(i))
							end if
						next
					
					if s_log_usuario_x_cd = "" then s_log_usuario_x_cd = "(nenhum)"
					if s_log_usuario_x_cd <> "" then s_log_usuario_x_cd = "CD (atual): " & s_log_usuario_x_cd

					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then 
							if s_log_perfil <> "" then s_log = s_log & "; " & s_log_perfil
							if s_log_loja_vendedor <> "" then s_log = s_log & "; " & s_log_loja_vendedor
							if s_log_usuario_x_cd <> "" then s_log = s_log & "; " & s_log_usuario_x_cd
							grava_log usuario, "", "", "", OP_LOG_USUARIO_INCLUSAO, s_log
							end if
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if (s_log <> "") Or (s_log_perfil <> "") Or (s_log_loja_vendedor <> "") Or (s_log_usuario_x_cd <> "") Or (s_log_perfil_anterior <> "") Or (s_log_loja_vendedor_anterior <> "") Or (s_log_usuario_x_cd_anterior <> "") then
							if s_log <> "" then s_log = "; " & s_log
							s_log="usuario=" & Trim("" & r("usuario")) & s_log
							if s_log_perfil_anterior <> "" then s_log = s_log & "; " & s_log_perfil_anterior
							if s_log_perfil <> "" then s_log = s_log & "; " & s_log_perfil
							if s_log_loja_vendedor_anterior <> "" then s_log = s_log & "; " & s_log_loja_vendedor_anterior
							if s_log_loja_vendedor <> "" then s_log = s_log & "; " & s_log_loja_vendedor
							if s_log_usuario_x_cd_anterior <> "" then s_log = s_log & "; " & s_log_usuario_x_cd_anterior
							if s_log_usuario_x_cd <> "" then s_log = s_log & "; " & s_log_usuario_x_cd
							grava_log usuario, "", "", "", OP_LOG_USUARIO_ALTERACAO, s_log
							end if
						end if
				else
					erro_fatal=True
					alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
					end if

				if alerta = "" then
				'	~~~~~~~~~~~~~~
					cn.CommitTrans
				'	~~~~~~~~~~~~~~
					if Err <> 0 then 
						alerta=Cstr(Err) & ": " & Err.Description
						erro_fatal = True
						end if
				else
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Err.Clear
					end if

				if r.State <> 0 then r.Close
				set r = nothing
				end if
		
		
		case else
		'	 ====
			alerta="OPERAÇÃO INVÁLIDA."
			
		end select

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


<body onload="bVOLTAR.focus();">
<center>
<br>

<!--  T E L A  -->

<p class="T">A V I S O</p>

<% 
	s = ""
	s_aux="'MtAviso'"
	if alerta <> "" then
		s = "<p style='margin:5px 2px 5px 2px;'>" & alerta & "</p>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "USUÁRIO " & chr(34) & s_usuario & chr(34) & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "USUÁRIO " & chr(34) & s_usuario & chr(34) & " ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "USUÁRIO " & chr(34) & s_usuario & chr(34) & " EXCLUÍDO COM SUCESSO."
			end select			
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<br><br>

<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
<%
	s="usuario.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>
</body>
</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>