<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===================================================================
'	  O R C A M E N T I S T A E I N D I C A D O R A T U A L I Z A . A S P
'     ===================================================================
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
	
	dim s, s_aux, usuario, loja, senha_cripto, alerta, chave, s2
	
	usuario = trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r, rs, rs2, t
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim criou_novo_reg
	Dim s_log
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = "|dt_ult_atualizacao|usuario_ult_atualizacao|timestamp|"
	
'	FOI UM RELATÓRIO QUE ORIGINOU A EDIÇÃO DO INDICADOR?
	dim pagina_relatorio_originou_edicao
	pagina_relatorio_originou_edicao = Trim(Request.Form("pagina_relatorio_originou_edicao"))
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim ChecadoStatusBloqueado, blnChecadoStatusBloqueado
	ChecadoStatusBloqueado = Trim(Request.Form("ChecadoStatusBloqueado"))
	blnChecadoStatusBloqueado = CBool(ChecadoStatusBloqueado)
	
	dim operacao_selecionada, s_id_selecionado, s_tipo_PJ_PF, s_razao_social_nome, s_nome_fantasia, s_cnpj_cpf, s_ie_rg, s_responsavel_principal
	dim s_endereco, s_endereco_numero, s_endereco_complemento, s_bairro, s_cidade, s_uf, s_cep, s_ddd, s_telefone, s_fax, url_origem
	dim s_ddd_cel, s_tel_cel, s_contato
	dim s_banco, s_agencia, s_conta, s_favorecido
	dim s_senha, s_senha2
	dim s_vendedor, s_acesso, s_status
	dim strEmail, strEmail2, strEmail3, strCaptador
	dim strChecadoStatus, strObs
	dim s_forma_como_conheceu_codigo
	dim s_nextel, rb_estabelecimento
    dim s_etq_endereco, s_etq_endereco_numero, s_etq_endereco_complemento, s_etq_bairro, s_etq_cidade, s_etq_uf, s_etq_cep, s_etq_ddd_1, s_etq_ddd_2, s_etq_tel_1, s_etq_tel_2, s_etq_email
    dim msg,s_favorecido_cnpjcpf
    dim n, cont, s_contato_nome, s_contato_id, s_contato_data, s_contato_log_inclusao, s_contato_log_exclusao, s_contato_log_edicao
	
	operacao_selecionada = request("operacao_selecionada")
	s_id_selecionado = UCase(trim(Request.Form("id_selecionado")))
	s_tipo_PJ_PF = trim(Request.Form("tipo_PJ_PF"))
	s_razao_social_nome = trim(Request.Form("razao_social_nome"))
	s_responsavel_principal = trim(Request.Form("c_responsavel_principal"))
	s_nome_fantasia = trim(Request.Form("c_nome_fantasia"))
	s_cnpj_cpf = retorna_so_digitos(trim(Request.Form("cnpj_cpf")))
	s_ie_rg = trim(Request.Form("ie_rg"))
	s_endereco = trim(Request.Form("endereco"))
	s_endereco_numero = trim(Request.Form("endereco_numero"))
	s_endereco_complemento = trim(Request.Form("endereco_complemento"))
	s_bairro = trim(Request.Form("bairro"))
	s_cidade = trim(Request.Form("cidade"))
	s_uf = trim(Request.Form("uf"))
	s_cep = retorna_so_digitos(trim(Request.Form("cep")))
	s_ddd = retorna_so_digitos(trim(Request.Form("ddd")))
	s_telefone = retorna_so_digitos(trim(Request.Form("telefone")))
	s_fax = retorna_so_digitos(trim(Request.Form("fax")))
	s_ddd_cel = retorna_so_digitos(trim(Request.Form("ddd_cel")))
	s_tel_cel = retorna_so_digitos(trim(Request.Form("tel_cel")))
	s_senha=UCase(trim(Request.Form("senha")))
	s_senha2=UCase(trim(Request.Form("senha2")))
	s_contato = trim(Request.Form("contato"))
	s_acesso = trim(Request.Form("rb_acesso_hidden"))
	strEmail = trim(Request.Form("c_email"))
	strEmail2 = trim(Request.Form("c_email2"))
	strEmail3 = trim(Request.Form("c_email3"))
	s_nextel = trim(Request.Form("c_nextel"))
	rb_estabelecimento = trim(Request.Form("rb_estabelecimento"))
	s_acesso = Request.Form("rb_acesso")
	s_etq_endereco = trim(Request.Form("etq_endereco"))
    s_etq_endereco_numero = trim(Request.Form("etq_endereco_numero"))
    s_etq_endereco_complemento = trim(Request.Form("etq_endereco_complemento"))
    s_etq_bairro = trim(Request.Form("etq_bairro"))
    s_etq_cidade = trim(Request.Form("etq_cidade"))
    s_etq_uf = trim(Request.Form("etq_uf"))
    s_etq_cep = retorna_so_digitos(trim(Request.Form("etq_cep")))
    s_etq_ddd_1 = retorna_so_digitos(trim(Request.Form("etq_ddd_1")))
    s_etq_ddd_2 = retorna_so_digitos(trim(Request.Form("etq_ddd_2")))
    s_etq_tel_1 = retorna_so_digitos(trim(Request.Form("etq_tel_1")))
    s_etq_tel_2 = retorna_so_digitos(trim(Request.Form("etq_tel_2")))
    s_etq_email = trim(Request.Form("etq_email"))

    url_origem = Request("url_origem")

	if blnChecadoStatusBloqueado then
		strChecadoStatus = ""
	else
		strChecadoStatus = trim(Request.Form("rb_checado"))
		end if
	
	
	if s_id_selecionado = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)

	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if s_id_selecionado = "" then
		alerta="FORNEÇA UM IDENTIFICADOR (APELIDO) PARA O ORÇAMENTISTA / INDICADOR."
	elseif s_razao_social_nome = "" then
		if s_tipo_PJ_PF = ID_PJ then
			alerta="PREENCHA A RAZÃO SOCIAL DO ORÇAMENTISTA / INDICADOR."
		else
			alerta="PREENCHA O NOME DO ORÇAMENTISTA / INDICADOR."
			end if
	end if
	
	if operacao_selecionada=OP_INCLUI then
	if s_status = "" then
		alerta="INFORME SE O STATUS ESTÁ ATIVO OU INATIVO."
	elseif s_vendedor = "" then
		alerta="INFORME POR QUAL VENDEDOR O INDICADOR SERÁ ATENDIDO."
		end if
	
	if alerta = "" then
		if operacao_selecionada = OP_INCLUI then
			if strCaptador = "" then alerta="INFORME QUEM É O CAPTADOR."
			end if
		end if
	end if
	
	if alerta = "" then
		if (operacao_selecionada <> OP_EXCLUI) And (CLng(s_acesso) <> 0) then
			if len(s_senha) < 5 then
				alerta="A SENHA DEVE POSSUIR NO MÍNIMO 5 CARACTERES."
			elseif s_senha <> s_senha2 then
				alerta="A CONFIRMAÇÃO DA SENHA NÃO ESTÁ CORRETA."
				end if
			end if
		end if
	
	if alerta = "" then
		if (s_endereco<>"") Or (s_bairro<>"") Or (s_cidade<>"") Or (s_uf<>"") Or (s_cep<>"") then
			if s_endereco="" then
				alerta="PREENCHA O ENDEREÇO."
			elseif Len(s_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
				alerta="ENDEREÇO EXCEDE O TAMANHO MÁXIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(s_endereco)) & " CARACTERES<br>TAMANHO MÁXIMO: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " CARACTERES"
			elseif s_endereco_numero="" then
				alerta="PREENCHA O NÚMERO DO ENDEREÇO."
			elseif s_cidade="" then
				alerta="PREENCHA A CIDADE DO ENDEREÇO."
			elseif s_uf="" then
				alerta="PREENCHA A UF DO ENDEREÇO."
			elseif s_cep="" then
				alerta="PREENCHA O CEP DO ENDEREÇO."
				end if
			end if
		end if
	
	if alerta <> "" then erro_consistencia=True	
	
	if s_senha <> "" then
		chave = gera_chave(FATOR_BD)
		codifica_dado s_senha, senha_cripto, chave
		end if
		
	if s_senha <> "" then
		chave = gera_chave(FATOR_BD)
		codifica_dado s_senha, senha_cripto, chave
		end if
	
'	VALIDAÇÃO P/ PERMITIR SOMENTE UM CADASTRO POR LOJA P/ CADA CPF/CNPJ
	dim s_label
	dim r_orcamentista_e_indicador
	dim blnErroDuplicidadeCadastro
	blnErroDuplicidadeCadastro=False
	if alerta = "" then
		if operacao_selecionada  = OP_INCLUI then
			s = "SELECT" & _
					" apelido," & _
					" cnpj_cpf," & _
					" razao_social_nome," & _
					" loja" & _
				" FROM t_ORCAMENTISTA_E_INDICADOR" & _
				" WHERE" & _
					" (cnpj_cpf = '" & s_cnpj_cpf & "')" & _
					" AND (Convert(smallint, loja) = " & loja & ")" & _
				" ORDER BY" & _
					" apelido"
			set rs = cn.Execute(s)
			if Not rs.Eof then
				blnErroDuplicidadeCadastro=True
				if Len(s_cnpj_cpf) = 11 then
					s_label="CPF"
				else
					s_label="CNPJ"
					end if
				alerta="O " & s_label & " " & cnpj_cpf_formata(s_cnpj_cpf) & " já está cadastrado na loja " & loja

				do while Not rs.Eof
					alerta=texto_add_br(alerta)
					alerta=alerta & rs("apelido") & " - " & Trim("" & rs("razao_social_nome"))
					rs.MoveNext
					loop
				end if
		
			if rs.State <> 0 then rs.Close
		elseif operacao_selecionada <> OP_EXCLUI then
			' CONSISTE SOMENTE SE ESTIVER ALTERANDO O CPF/CNPJ (OBS: NESTA PÁGINA NÃO É POSSÍVEL ALTERAR A LOJA)
			call le_orcamentista_e_indicador(s_id_selecionado, r_orcamentista_e_indicador, msg_erro)
			if retorna_so_digitos(s_cnpj_cpf) <> retorna_so_digitos(r_orcamentista_e_indicador.cnpj_cpf) then
				s = "SELECT" & _
						" apelido," & _
						" cnpj_cpf," & _
						" razao_social_nome," & _
						" loja" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (cnpj_cpf = '" & s_cnpj_cpf & "')" & _
						" AND (Convert(smallint, loja) = " & loja & ")" & _
						" AND (apelido <> '" & s_id_selecionado & "')" & _
					" ORDER BY" & _
						" apelido"
				set rs = cn.Execute(s)
				if Not rs.Eof then
					blnErroDuplicidadeCadastro=True
					if Len(s_cnpj_cpf) = 11 then
						s_label="CPF"
					else
						s_label="CNPJ"
						end if
					alerta="Não é possível fazer a alteração porque o " & s_label & " "
					alerta=alerta & cnpj_cpf_formata(s_cnpj_cpf) & " já está cadastrado na loja " & loja

					do while Not rs.Eof
						alerta=texto_add_br(alerta)
						alerta=alerta & rs("apelido") & " - " & Trim("" & rs("razao_social_nome"))
						rs.MoveNext
						loop
					end if
		
				if rs.State <> 0 then rs.Close
				end if
			end if
		end if


	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
    if Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(t, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			s="SELECT COUNT(*) AS qtde FROM t_ORCAMENTO WHERE (orcamentista = '" & s_id_selecionado & "')"
			r.Open s, cn
		'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
			if Cstr(r("qtde")) > Cstr(0) then
				erro_fatal=True
				alerta = "ORÇAMENTISTA / INDICADOR NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE ORÇAMENTOS."
				end if
			r.Close 
			
			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_PEDIDO WHERE (indicador = '" & s_id_selecionado & "') OR (orcamentista = '" & s_id_selecionado & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "ORÇAMENTISTA / INDICADOR NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE PEDIDOS."
					end if
				r.Close 
				end if

			if Not erro_fatal then
			'	INFO P/ LOG
				s="SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido = '" & s_id_selecionado & "'"
                
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
				if TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO then
				'	BLOQUEIA REGISTRO PARA EVITAR ACESSO CONCORRENTE (REALIZA O FLIP EM UM CAMPO BIT APENAS P/ ADQUIRIR O LOCK EXCLUSIVO)
				'	OBS: TODOS OS MÓDULOS DO SISTEMA QUE REALIZEM ESTA OPERAÇÃO DE CADASTRAMENTO DEVEM SINCRONIZAR O ACESSO OBTENDO O LOCK EXCLUSIVO DO REGISTRO DE CONTROLE DESIGNADO
					s = "UPDATE t_CONTROLE SET" & _
							" dummy = ~dummy" & _
						" WHERE" & _
							" id_nsu = '" & ID_XLOCK_SYNC_ORCAMENTISTA_E_INDICADOR & "'"
					cn.Execute(s)
					end if

				s ="DELETE FROM t_ORCAMENTISTA_E_INDICADOR_LOG WHERE (apelido = '" & s_id_selecionado & "')"
				cn.Execute(s)
				If Err <> 0 then
					erro_fatal=True
					alerta = "FALHA AO EXCLUIR OS DADOS DE LOG DE EDIÇÃO DO ORÇAMENTISTA / INDICADOR (" & Cstr(Err) & ": " & Err.Description & ")."
					end if

				s="DELETE FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido = '" & s_id_selecionado & "'"
				cn.Execute(s)
				If Err = 0 then 
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_EXCLUSAO, s_log
				else
					erro_fatal=True
					alerta = "FALHA AO REMOVER O ORÇAMENTISTA / INDICADOR (" & Cstr(Err) & ": " & Err.Description & ")."
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
				if TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO then
				'	BLOQUEIA REGISTRO PARA EVITAR ACESSO CONCORRENTE (REALIZA O FLIP EM UM CAMPO BIT APENAS P/ ADQUIRIR O LOCK EXCLUSIVO)
				'	OBS: TODOS OS MÓDULOS DO SISTEMA QUE REALIZEM ESTA OPERAÇÃO DE CADASTRAMENTO DEVEM SINCRONIZAR O ACESSO OBTENDO O LOCK EXCLUSIVO DO REGISTRO DE CONTROLE DESIGNADO
					s = "UPDATE t_CONTROLE SET" & _
							" dummy = ~dummy" & _
						" WHERE" & _
							" id_nsu = '" & ID_XLOCK_SYNC_ORCAMENTISTA_E_INDICADOR & "'"
					cn.Execute(s)
					end if

                ' log cadastrou novo
                    if operacao_selecionada = OP_INCLUI then
                    s = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido = '" & s_id_selecionado & "'"
                    s2 = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_LOG WHERE (id = -1)"
				    r.Open s, cn
                    rs2.Open s2, cn
                    teste = ""
                    if rs2.EOF then 
                        if Not fin_gera_nsu(T_ORCAMENTISTA_E_INDICADOR_LOG, intNsuNovoLog, msg_erro) then 
			                alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		                else
			                if intNsuNovoLog <= 0 then
				                alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovoLog & ")"
				                end if
			                end if
                        rs2.AddNew
                        rs2("id") = intNsuNovoLog
                        rs2("apelido") = s_id_selecionado
                        rs2("loja") = loja
                        rs2("mensagem") = "Indicador cadastrado"
                        rs2("usuario") = usuario
                    rs2.Update
                    end if
                    if rs2.State <> 0 then rs2.Close
                  end if
                
                if operacao_selecionada = OP_CONSULTA then
                msg = ""
                dim x1, x2
                dim intNsuNovoLog
				s = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido = '" & s_id_selecionado & "'"
                s2 = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_LOG WHERE (id = -1)"
				r.Open s, cn
                rs2.Open s2, cn
                if rs2.EOF then 
                    
                    ' log razão social
                    if (s_razao_social_nome <> r("razao_social_nome")) then 
                        msg = msg & "Razão Social alterada <br>|de: " & r("razao_social_nome") & "<br>|para: " & s_razao_social_nome & "<br>"
                    end if
                    ' log responsável principal
                    if (s_responsavel_principal <> Trim("" & r("responsavel_principal"))) then 
                        if s_responsavel_principal = "" then x2 = "VAZIO" else x2 = s_responsavel_principal
                        if Trim("" & r("responsavel_principal")) = "" then x1 = "VAZIO" else x1 = r("responsavel_principal")
                        msg = msg & "Responsável principal alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if
                    ' log nome fantasia
                    if isNull(r("nome_fantasia")) then r("nome_fantasia") = ""
                    if (s_nome_fantasia <> r("nome_fantasia")) then
                        if s_nome_fantasia = "" then x2 = "VAZIO" else x2 = s_nome_fantasia
                        if r("nome_fantasia") = "" then x1 = "VAZIO" else x1 = r("nome_fantasia")
                        msg = msg & "Nome Fantasia alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log cnpj cpf
                    if (s_cnpj_cpf <> r("cnpj_cpf")) then
                        msg = msg & "CPF/CNPJ alterado <br>|de: " & r("cnpj_cpf") & "<br>|para: " & s_cnpj_cpf & "<br>"
                    end if                
                    ' log ie rg
                    if (s_ie_rg <> r("ie_rg")) then
                        if s_ie_rg = "" then x2 = "VAZIO" else x2 = s_ie_rg
                        if r("ie_rg") = "" then x1 = "VAZIO" else x1 = r("ie_rg")
                        msg = msg & "IE/RG alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log endereco
                    if isNull(r("endereco")) then r("endereco") = ""
                    if (s_endereco <> r("endereco")) then
                      msg = msg & "Endereço alterado <br>|de: " & r("endereco") & "<br>|para: " & s_endereco & "<br>"
                    end if

                    ' log endereco numero
                    if isNull(r("endereco_numero")) then r("endereco_numero") = ""
                    if (s_endereco_numero <> r("endereco_numero")) then
                        msg = msg & "Número do endereço alterado <br>|de: " & r("endereco_numero") & "<br>|para: " & s_endereco_numero & "<br>"
                    end if

                    ' log endereco complemento
                    if isNull(r("endereco_complemento")) then r("endereco_complemento") = ""
                    if (s_endereco_complemento <> r("endereco_complemento")) then
                        if s_endereco_complemento = "" then x2 = "VAZIO" else x2 = s_endereco_complemento
                        if r("endereco_complemento") = "" then x1 = "VAZIO" else x1 = r("endereco_complemento")
                        msg = msg & "Complemento do endereço alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log bairro
                    if isNull(r("bairro")) then r("bairro") = ""
                    if (s_bairro <> r("bairro")) then
                        if s_bairro = "" then x2 = "VAZIO" else x2 = s_bairro
                        if r("bairro") = "" then x1 = "VAZIO" else x1 = r("bairro")
                        msg = msg & "Bairro alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log cidade
                    if (s_cidade <> r("cidade")) then
                        msg = msg & "Cidade alterada <br>|de: " & r("cidade") & "<br>|para: " & s_cidade & "<br>"
                    end if

                    ' log uf
                    if (s_uf <> r("uf")) then
                        msg = msg & "UF alterado <br>|de: " & r("uf") & "<br>|para: " & s_uf & "<br>"
                    end if

                    ' log cep
                    if (s_cep <> r("cep")) then
                        msg = msg & "CEP alterado <br>|de: " & r("cep") & "<br>|para: " & s_cep & "<br>"
                    end if
                        
                    ' log ddd tel
                    if isNull(r("ddd")) then r("ddd") = ""
                    if isNull(r("telefone")) then r("telefone") = ""
                    if (s_ddd <> r("ddd") Or s_telefone <> r("telefone")) then
                        msg = msg & "Telefone alterado <br>|de: " & r("ddd") & "&nbsp;" & r("telefone") & "<br>|para: " & s_ddd & "&nbsp;" & s_telefone & "<br>"
                    end if

                    ' log fax
                    if (s_fax <> r("fax")) then
                        if s_fax = "" then x2 = "VAZIO" else x2 = s_fax
                        if r("fax") = "" then x1 = "VAZIO" else x1 = r("fax")
                        msg = msg & "Número do fax alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log ddd cel
                    if isNull(r("ddd_cel")) then r("ddd_cel") = ""
                    if isNull(r("tel_cel")) then r("tel_cel") = ""
                    if (s_ddd_cel <> r("ddd_cel") Or s_tel_cel <> r("tel_cel")) then
                        
                        if (r("ddd_cel") = "") then
                            msg = msg & "Telefone (Cel) alterado <br>|de: VAZIO<br>"
                        else 
                            msg = msg & "Telefone (Cel) alterado<br>|de: " & r("ddd_cel") & "&nbsp;" & r("tel_cel") & "<br>"
                        end if
                        if (s_ddd_cel = "") then
                            msg = msg & "|para: VAZIO<br>"
                        else 
                            msg = msg & "|para: " & s_ddd_cel & "&nbsp;" & s_tel_cel & "<br>"
                        end if

                    end if

                    ' log email 1
                    if isNull(r("email")) then r("email") = ""
                    if (strEmail <> r("email")) then
                        if strEmail = "" then x2 = "VAZIO" else x2 = strEmail
                        if r("email") = "" then x1 = "VAZIO" else x1 = r("email")
                        msg = msg & "Email (1) alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log email 2
                    if isNull(r("email2")) then r("email2") = ""
                    if (strEmail2 <> r("email2")) then
                        if strEmail2 = "" then x2 = "VAZIO" else x2 = strEmail2
                        if r("email2") = "" then x1 = "VAZIO" else x1 = r("email2")
                        msg = msg & "Email (2) alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log email 3
                    if isNull(r("email3")) then r("email3") = ""
                    if (strEmail3 <> r("email3")) then
                        if strEmail3 = "" then x2 = "VAZIO" else x2 = strEmail3
                        if r("email3") = "" then x1 = "VAZIO" else x1 = r("email3")
                        msg = msg & "Email (3) alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log nextel
                    if isNull(r("nextel")) then r("nextel") = ""
                    if (s_nextel <> r("nextel")) then
                        if s_nextel = "" then x2 = "VAZIO" else x2 = s_nextel
                        if r("nextel") = "" then x1 = "VAZIO" else x1 = r("nextel")
                        msg = msg & "Nextel alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log contato
                    if isNull(r("contato")) then r("contato") = ""
                    if (s_contato <> r("contato")) then
                        if s_contato = "" then x2 = "VAZIO" else x2 = s_contato
                        if r("contato") = "" then x1 = "VAZIO" else x1 = r("contato")
                        msg = msg & "Contato alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log estabelecimento
                    if isNull(r("tipo_estabelecimento")) then r("tipo_estabelecimento") = ""
                    if (rb_estabelecimento = "") then rb_estabelecimento = 0
                    if (CLng(rb_estabelecimento) <> r("tipo_estabelecimento")) then
                            function tipoEstabelecimento(num) 
                                dim s
                                if num = "" then num = 0
                                select case num 
                                    case 0
                                        s = "VAZIO"                               
                                    case 1  
                                        s = "Casa"
                                    case 2
                                        s = "Escritório"
                                    case 3
                                        s = "Loja"
                                    case 4  
                                        s = "Oficina"
                                    case else
                                        s = ""
                                    end select
                                tipoEstabelecimento = s
                            end function
                        msg = msg & "Tipo de estabelecimento alterado <br>|de: " & tipoEstabelecimento(r("tipo_estabelecimento")) & "<br>|para: " & tipoEstabelecimento(rb_estabelecimento) & "<br>"
                    end if

                    ' log acesso sistema
                    if isNull(r("hab_acesso_sistema")) then r("hab_acesso_sistema") = ""
                    if (s_acesso = "") then s_acesso = 0
                    if (CLng(s_acesso) <> r("hab_acesso_sistema")) then
                        function tipoAcessoSistema(num) 
                            dim s
                            select case num
                                case 1  
                                    s = "Liberado"
                                case 0
                                    s = "Bloqueado"
                                case else
                                    s = "VAZIO"
                                end select
                            tipoAcessoSistema = s
                        end function
                        msg = msg & "Acesso ao sistema alterado <br>|de: " & tipoAcessoSistema(r("hab_acesso_sistema")) & "<br>|para: " & tipoAcessoSistema(s_acesso) & "<br>"
                    end if

                    ' log senha
                    if isNull(r("datastamp")) then r("datastamp") = ""
                    if (senha_cripto <> r("datastamp")) then
                        msg = msg & "Senha alterada<br>|de: ******<br>|para: ******<br>"
                    end if

                    ' log endereco etiqueta
                    if isNull(r("etq_endereco")) then r("etq_endereco") = ""
                    if (s_etq_endereco <> r("etq_endereco")) then
                      if s_etq_endereco = "" then x2 = "VAZIO" else x2 = s_etq_endereco
                      if r("etq_endereco") = "" then x1 = "VAZIO" else x1 = r("etq_endereco")
                      msg = msg & "Endereço da etiqueta alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log endereco numero etiqueta
                    if isNull(r("etq_endereco_numero")) then r("etq_endereco_numero") = ""
                    if (s_etq_endereco_numero <> r("etq_endereco_numero")) then
                        if s_etq_endereco_numero = "" then x2 = "VAZIO" else x2 = s_etq_endereco_numero
                        if r("etq_endereco_numero") = "" then x1 = "VAZIO" else x1 = r("etq_endereco_numero")
                        msg = msg & "Número do endereço da etiqueta alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log endereco complemento etiqueta
                    if isNull(r("etq_endereco_complemento")) then r("etq_endereco_complemento") = ""
                    if (s_etq_endereco_complemento <> r("etq_endereco_complemento")) then
                        if s_etq_endereco_complemento = "" then x2 = "VAZIO" else x2 = s_etq_endereco_complemento
                        if r("etq_endereco_complemento") = "" then x1 = "VAZIO" else x1 = r("etq_endereco_complemento")
                        msg = msg & "Complemento do endereço da etiqueta alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log bairro etiqueta
                    if isNull(r("etq_bairro")) then r("etq_bairro") = ""
                    if (s_etq_bairro <> r("etq_bairro")) then
                        if s_etq_bairro = "" then x2 = "VAZIO" else x2 = s_etq_bairro
                        if r("etq_bairro") = "" then x1 = "VAZIO" else x1 = r("etq_bairro")
                        msg = msg & "Bairro da etiqueta alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log cidade etiqueta
                    if isNull(r("etq_cidade")) then r("etq_cidade") = ""
                    if (s_etq_cidade <> r("etq_cidade")) then
                        if s_etq_cidade = "" then x2 = "VAZIO" else x2 = s_etq_cidade
                        if r("etq_cidade") = "" then x1 = "VAZIO" else x1 = r("etq_cidade")
                        msg = msg & "Cidade da etiqueta alterada <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log uf etiqueta
                    if isNull(r("etq_uf")) then r("etq_uf") = ""
                    if (s_etq_uf <> r("etq_uf")) then
                        if s_etq_uf = "" then x2 = "VAZIO" else x2 = s_etq_uf
                        if r("etq_uf") = "" then x1 = "VAZIO" else x1 = r("etq_uf")
                        msg = msg & "UF da etiqueta alterada <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log cep etiqueta
                    if isNull(r("etq_cep")) then r("etq_cep") = ""
                    if (s_etq_cep <> r("etq_cep")) then
                        if s_etq_cep = "" then x2 = "VAZIO" else x2 = s_etq_cep
                        if r("etq_cep") = "" then x1 = "VAZIO" else x1 = r("etq_cep")
                        msg = msg & "CEP da etiqueta alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log ddd tel 1 etiqueta
                    if isNull(r("etq_ddd_1")) then r("etq_ddd_1") = ""
                    if isNull(r("etq_tel_1")) then r("etq_tel_1") = ""
                    if (s_etq_ddd_1 <> r("etq_ddd_1") Or s_etq_tel_1 <> r("etq_tel_1")) then
                        
                        if (r("etq_ddd_1") = "") then
                            msg = msg & "Telefone (1) da etiqueta alterado <br>|de: VAZIO<br>"
                        else 
                            msg = msg & "Telefone (1) da etiqueta alterado<br>|de: " & r("etq_ddd_1") & "&nbsp;" & r("etq_tel_1") & "<br>"
                        end if
                        if (s_etq_ddd_1 = "") then
                            msg = msg & "|para: VAZIO<br>"
                        else 
                            msg = msg & "|para: " & s_etq_ddd_1 & "&nbsp;" & s_etq_tel_1 & "<br>"
                        end if

                    end if

                    ' log ddd tel 2 etiqueta
                    if isNull(r("etq_ddd_2")) then r("etq_ddd_2") = ""
                    if isNull(r("etq_tel_2")) then r("etq_tel_2") = ""
                    if (s_etq_ddd_2 <> r("etq_ddd_2") Or s_etq_tel_2 <> r("etq_tel_2")) then
                        
                        if (r("etq_ddd_2") = "") then
                            msg = msg & "Telefone (2) da etiqueta alterado <br>|de: VAZIO<br>"
                        else 
                            msg = msg & "Telefone (2) da etiqueta alterado<br>|de: " & r("etq_ddd_2") & "&nbsp;" & r("etq_tel_2") & "<br>"
                        end if
                        if (s_etq_ddd_2 = "") then
                            msg = msg & "|para: VAZIO<br>"
                        else 
                            msg = msg & "|para: " & s_etq_ddd_2 & "&nbsp;" & s_etq_tel_2 & "<br>"
                        end if

                    end if

                    ' log email etiqueta 
                    if isNull(r("etq_email")) then r("etq_email") = ""
                    if (s_etq_email <> r("etq_email")) then
                        if s_etq_email = "" then x2 = "VAZIO" else x2 = s_etq_email
                        if r("etq_email") = "" then x1 = "VAZIO" else x1 = r("etq_email")
                        msg = msg & "Email da etiqueta alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    
                    ' Vendedores/contatos
                    n = Request.Form("c_indicador_contato").Count
                    s_contato_log_inclusao = ""
                    s_contato_log_exclusao = ""
                    s_contato_log_edicao = ""
                    for cont = 1 to n
                        s_contato_nome = Trim(Request.Form("c_indicador_contato")(cont))
                        s_contato_id = Trim(Request.Form("contato_id")(cont))
                        s_contato_data = Trim(Request.Form("c_indicador_contato_data")(cont))
                        s2 = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_CONTATOS WHERE (indicador = '" & s_id_selecionado & "' AND id='" & s_contato_id & "')"
                        t.Open s2, cn
                        if Not t.Eof then
                            if Trim("" & t("nome")) <> "" And s_contato_nome = "" then
                                if s_contato_log_exclusao = "" then s_contato_log_exclusao = "indicador=" & s_id_selecionado
                                s_contato_log_exclusao = s_contato_log_exclusao & "; nome: " & t("nome")
                                msg = msg & "Vendedor excluído da lista<br>|nome: " & t("nome") & "<br>"
                                s = "DELETE FROM t_ORCAMENTISTA_E_INDICADOR_CONTATOS WHERE (id = " & s_contato_id & ")"
                                cn.Execute(s)
                                end if
                            end if
                        if s_contato_nome <> "" then
                            if t.Eof then
                                if s_contato_log_inclusao = "" then s_contato_log_inclusao = "indicador=" & s_id_selecionado
                                s_contato_log_inclusao = s_contato_log_inclusao & "; nome: " & s_contato_nome
                                msg = msg & "Novo vendedor incluso na lista<br>|nome: " & s_contato_nome & "<br>"
                                t.AddNew
                                t("indicador") = s_id_selecionado
                                t("dt_cadastro") = Date
                                t("usuario_cadastro") = usuario
                            else
                                if s_contato_nome <> Trim("" & t("nome")) then
                                    if s_contato_log_edicao = "" then s_contato_log_edicao = "indicador=" & s_id_selecionado
                                    s_contato_log_edicao = s_contato_log_edicao & "; nome anterior: " & Trim("" & t("nome")) & "; nome atual: " & s_contato_nome
                                    msg = msg & "Vendedor alterado na lista<br>|de: " & Trim("" & t("nome")) & "<br>|para: " & s_contato_nome & "<br>"
                                    t("dt_cadastro") = Date
                                    t("usuario_ult_atualizacao") = usuario
                                    t("dt_ult_atualizacao") = Now
                                    end if
                                end if
                            t("nome") = s_contato_nome                    
                            t.Update

                            end if
                        if t.State <> 0 then t.Close
                        next
                    set t = nothing
                    
                    ' grava log contatos
                    if s_contato_log_exclusao <> "" then grava_log usuario, "", "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_CONTATOS__EXCLUSAO, s_contato_log_exclusao
                    if s_contato_log_inclusao <> "" then grava_log usuario, "", "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_CONTATOS__INCLUSAO, s_contato_log_inclusao
                    if s_contato_log_edicao <> "" then grava_log usuario, "", "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_CONTATOS__EDICAO, s_contato_log_edicao
                    
                    
                    ' Houve alteração?
                    if (msg <> "") then
                        if Not fin_gera_nsu(T_ORCAMENTISTA_E_INDICADOR_LOG, intNsuNovoLog, msg_erro) then 
			                alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		                else
			                if intNsuNovoLog <= 0 then
				                alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovoLog & ")"
				                end if
			                end if
                        rs2.AddNew
                        rs2("id") = intNsuNovoLog
                        rs2("apelido") = s_id_selecionado
                        rs2("loja") = loja
                        rs2("mensagem") = msg
                        rs2("usuario") = usuario
                    rs2.Update
                    end if

                end if

                if rs2.State <> 0 then rs2.Close
                end if

				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("apelido")=s_id_selecionado
					r("dt_cadastro") = Date
					r("usuario_cadastro") = usuario
					r("senha") = gera_senha_aleatoria
					r("sistema_responsavel_cadastro") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					end if
					
				r("dt_ult_atualizacao") = Now
				r("usuario_ult_atualizacao") = usuario
				
				r("tipo") = s_tipo_PJ_PF
				r("razao_social_nome")=s_razao_social_nome
				r("nome_fantasia") = s_nome_fantasia
				r("responsavel_principal") = s_responsavel_principal
				r("cnpj_cpf") = s_cnpj_cpf
				r("ie_rg") = s_ie_rg
				r("endereco") = s_endereco
				r("endereco_numero") = s_endereco_numero
				r("endereco_complemento") = s_endereco_complemento
				r("bairro") = s_bairro
				r("cidade") = s_cidade
				r("uf") = s_uf
				r("cep") = s_cep
				r("ddd") = s_ddd
				r("telefone") = s_telefone
				r("fax") = s_fax
				r("ddd_cel") = s_ddd_cel
				r("tel_cel") = s_tel_cel
				r("contato") = s_contato
				r("hab_acesso_sistema") = s_acesso
                r("etq_endereco") = s_etq_endereco
				r("etq_endereco_numero") = s_etq_endereco_numero
				r("etq_endereco_complemento") = s_etq_endereco_complemento
				r("etq_bairro") = s_etq_bairro
				r("etq_cidade") = s_etq_cidade
				r("etq_uf") = s_etq_uf
				r("etq_cep") = s_etq_cep
                r("etq_ddd_1") = s_etq_ddd_1
                r("etq_ddd_2") = s_etq_ddd_2
                r("etq_tel_1") = s_etq_tel_1
                r("etq_tel_2") = s_etq_tel_2
                r("etq_email") = s_etq_email				
				if trim("" & r("datastamp"))<>senha_cripto then
					r("datastamp")=senha_cripto
					r("senha") = gera_senha_aleatoria
					r("dt_ult_alteracao_senha") = Null
					end if
					
				r("email") = strEmail
				r("email2") = strEmail2
				r("email3") = strEmail3
				
				if (Not blnChecadoStatusBloqueado) then 
					if strChecadoStatus <> "" then
						if CLng(strChecadoStatus) <> r("checado_status") then
							r("checado_status") = CLng(strChecadoStatus)
							r("checado_data") = Now
							r("checado_usuario") = usuario
							end if
						end if
					end if
				
				r("nextel") = s_nextel
				
				if rb_estabelecimento <> "" then r("tipo_estabelecimento") = CLng(rb_estabelecimento)
				
				r("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP

				r.Update

				If Err = 0 then 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then 
							grava_log usuario, "", "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_INCLUSAO, s_log
							end if
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if (s_log <> "") then 
							if s_log <> "" then s_log = "; " & s_log
							s_log="apelido=" & Trim("" & r("apelido")) & s_log
							grava_log usuario, "", "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_ALTERACAO, s_log
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


	if alerta = "" then
		if pagina_relatorio_originou_edicao <> "" then
			s = pagina_relatorio_originou_edicao
			if InStr(s, "?") <> 0 then s = s & "&" else s = s & "?"
			s = s & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
			Response.Redirect(s)
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


<body onload="bVOLTAR.focus();">
<center>
<br>

<!--  T E L A  -->

<p class="T">A V I S O</p>

<% 
	s = ""
	s_aux="'MtAviso'"
	if alerta <> "" then
		s = "<P style='margin:5px 2px 5px 2px;'>" & alerta & "</P>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "ORÇAMENTISTA / INDICADOR " & chr(34) & s_id_selecionado & chr(34) & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "ORÇAMENTISTA / INDICADOR " & chr(34) & s_id_selecionado & chr(34) & " ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "ORÇAMENTISTA / INDICADOR " & chr(34) & s_id_selecionado & chr(34) & " EXCLUÍDO COM SUCESSO."
			end select
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:600px;font-weight:bold;" align="center"><%=s%></div>
<br><br>

<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
<%
	s="MenuOrcamentistaEIndicador.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<% if blnErroDuplicidadeCadastro then %>
    <td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back();"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<% else %>
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="<%=url_origem%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<% end if %>
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