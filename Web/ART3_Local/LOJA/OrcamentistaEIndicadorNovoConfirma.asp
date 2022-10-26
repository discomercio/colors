<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================================
'	  O R C A M E N T I S T A E I N D I C A D O R N O V O C O N F I R M A . A S P
'     ===========================================================================
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
	
	dim s, s_aux, usuario, loja, alerta
	
	usuario = trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	if Not operacao_permitida(OP_LJA_CADASTRAMENTO_INDICADOR, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim msg_erro
	dim cn, r, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim r_loja_user_session
	set r_loja_user_session = New cl_LOJA
	if Not x_loja_bd(loja, r_loja_user_session) then Response.Redirect("aviso.asp?id=" & ERR_LOJA_NAO_CADASTRADA)

	Dim criou_novo_reg, s_senha, s_senha2, chave
	Dim s_log
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = "|dt_ult_atualizacao|usuario_ult_atualizacao|timestamp|"
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, s_id_selecionado, s_tipo_PJ_PF, s_razao_social_nome, s_nome_fantasia, s_cnpj_cpf, s_ie_rg, s_responsavel_principal
	dim s_endereco, s_endereco_numero, s_endereco_complemento, s_bairro, s_cidade, s_uf, s_cep, s_ddd, s_telefone, s_fax
	dim s_ddd_cel, s_tel_cel, s_contato, senha_cripto
	dim s_banco, s_agencia, s_conta, s_favorecido, s_favorecido_cnpjcpf, conta_dv, agencia_dv, tipo_conta, tipo_operacao
	dim s_loja, s_vendedor, s_acesso, s_status, s_permite_RA_status
	dim s_perc_desagio_RA, strValorLimiteMensal, strEmail, strEmail2, strEmail3, s_forma_como_conheceu_codigo
	dim strObs
	dim s_nextel, rb_estabelecimento
    dim rs2, s2, intNsuNovoLog
    dim n, cont, s_contato_nome, s_contato_id, s_contato_data, s_contato_log_inclusao
	dim s_id_magento_b2b, id_magento_b2b

	operacao_selecionada = request("operacao_selecionada")
	s_id_selecionado = UCase(trim(Request.Form("id_selecionado")))
	s_id_magento_b2b = Trim(Request.Form("c_id_magento_b2b"))
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
	s_contato = trim(Request.Form("contato"))
	s_banco = retorna_so_digitos(trim(Request.Form("banco")))
	s_agencia = trim(Request.Form("agencia"))
	s_conta = trim(Request.Form("conta"))
	s_favorecido = trim(Request.Form("favorecido"))
	strEmail = trim(Request.Form("c_email"))
	strEmail2 = trim(Request.Form("c_email2"))
	strEmail3 = trim(Request.Form("c_email3"))
	s_forma_como_conheceu_codigo = trim(Request.Form("c_forma_como_conheceu_codigo"))
	strObs = trim(Request.Form("c_obs"))
	s_nextel = trim(Request.Form("c_nextel"))
	rb_estabelecimento = trim(Request.Form("rb_estabelecimento"))
    s_acesso = trim(Request.Form("rb_acesso"))
    s_senha=UCase(trim(Request.Form("senha")))
	s_senha2=UCase(trim(Request.Form("senha2")))
    conta_dv = trim(Request.Form("conta_dv"))
    agencia_dv = trim(Request.Form("agencia_dv"))
    tipo_conta = trim(Request.Form("tipo_conta"))
    tipo_operacao = trim(Request.Form("tipo_operacao"))
    s_favorecido_cnpjcpf = retorna_so_digitos(trim(Request.Form("favorecido_cnpjcpf")))

    if Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
'	VALOR PADRÃO
	s_perc_desagio_RA = le_parametro_bd(ID_PARAM_PercDesagioRAIndicadorParaCadastroFeitoNaLoja, msg_erro)
	strValorLimiteMensal = le_parametro_bd(ID_PARAM_VlLimiteMensalIndicadorParaCadastroFeitoNaLoja, msg_erro)
'	STATUS ATIVO
	s_status = "A"

'	PERMITE RA
	s_permite_RA_status = "0"
'	AUTOMATICAMENTE PREENCHER LOJA E 'ATENDIDO POR' COM O USUÁRIO QUE ESTÁ CADASTRANDO
	s_loja = loja
	s_loja=normaliza_codigo(s_loja, TAM_MIN_LOJA)
	s_vendedor = usuario


	if s_id_selecionado = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)


	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if s_id_selecionado = "" then
		alerta="FORNEÇA UM IDENTIFICADOR (APELIDO) PARA O INDICADOR."
	elseif s_razao_social_nome = "" then
		if s_tipo_PJ_PF = ID_PJ then
			alerta="PREENCHA A RAZÃO SOCIAL DO INDICADOR."
		else
			alerta="PREENCHA O NOME DO INDICADOR."
			end if
	elseif s_cnpj_cpf = "" then
		if s_tipo_PJ_PF = ID_PJ then
			alerta="PREENCHA O CNPJ DO INDICADOR."
		else
			alerta="PREENCHA O CPF DO INDICADOR."
			end if
	elseif s_ie_rg = "" then
		if s_tipo_PJ_PF = ID_PJ then
			alerta="PREENCHA A IE DO INDICADOR."
		else
			alerta="PREENCHA O RG DO INDICADOR."
			end if
	elseif s_endereco = "" then
		alerta="PREENCHA O ENDEREÇO."
	elseif Len(s_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
		alerta="ENDEREÇO EXCEDE O TAMANHO MÁXIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(s_endereco)) & " CARACTERES<br>TAMANHO MÁXIMO: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " CARACTERES"
	elseif s_endereco_numero="" then
		alerta="PREENCHA O NÚMERO DO ENDEREÇO."
	elseif s_bairro = "" then
		alerta="PREENCHA O BAIRRO."
	elseif s_cidade = "" then
		alerta="PREENCHA A CIDADE."
	elseif s_uf = "" then
		alerta="PREENCHA A UF."
	elseif s_cep = "" then
		alerta="PREENCHA O CEP."
	elseif s_ddd = "" then
		alerta="PREENCHA O DDD."
	elseif len(s_ddd) > 2 then
		alerta="O DDD EXCEDE O TAMANHO MÁXIMO."
	elseif s_telefone = "" then
		alerta="PREENCHA O TELEFONE."
	elseif len(s_telefone) > 9 then
		alerta="O TELEFONE EXCEDE O TAMANHO MÁXIMO."
	elseif s_fax = "" then
		alerta="PREENCHA O FAX."
	elseif len(s_fax) > 9 then
		alerta="O Nº DO FAX EXCEDE O TAMANHO MÁXIMO."
	elseif s_ddd_cel = "" then
		alerta="PREENCHA O DDD DO CELULAR."
	elseif len(s_ddd_cel) > 2 then
		alerta="O DDD DO CELULAR EXCEDE O TAMANHO MÁXIMO."
	elseif s_tel_cel = "" then
		alerta="PREENCHA O TELEFONE CELULAR."
	elseif len(s_tel_cel) > 9 then
		alerta="O TELEFONE CELULAR EXCEDE O TAMANHO MÁXIMO."
	elseif s_contato = "" then
		alerta="PREENCHA O NOME DO CONTATO."
	elseif s_banco = "" then
		alerta="PREENCHA O Nº DO BANCO."
	elseif s_agencia = "" then
		alerta="PREENCHA O Nº DA AGÊNCIA."
	elseif s_conta = "" then
		alerta="PREENCHA O Nº DA CONTA"
	elseif s_favorecido = "" then
		alerta="PREENCHA O NOME DO FAVORECIDO."
	elseif s_acesso = "" then
		alerta="INFORME O ACESSO AO SISTEMA ESTÁ LIBERADO OU BLOQUEADO."
	elseif s_status = "" then
		alerta="INFORME SE O STATUS ESTÁ ATIVO OU INATIVO."
	elseif s_permite_RA_status = "" then
		alerta="INFORME SE O RA É PERMITIDO OU NÃO."
	elseif Trim(x_loja(s_loja)) = "" then
		alerta="LOJA " & s_loja & " NÃO ESTÁ CADASTRADA."
	elseif s_vendedor = "" then
		alerta="INFORME POR QUAL VENDEDOR O INDICADOR SERÁ ATENDIDO."
	elseif (converte_numero(s_perc_desagio_RA)<0) Or (converte_numero(s_perc_desagio_RA)>100) then
		alerta="PERCENTUAL DE DESÁGIO DO RA É INVÁLIDO."
	elseif converte_numero(strValorLimiteMensal) < 0 then
		alerta="VALOR DO LIMITE MENSAL DE COMPRAS É INVÁLIDO (" & strValorLimiteMensal & ")."
	elseif (strEmail = "") And (strEmail2 = "") And (strEmail3 = "") then
		alerta="INFORME NO MÍNIMO UM ENDEREÇO DE E-MAIL."
	elseif s_forma_como_conheceu_codigo = "" then
		alerta="INDIQUE A FORMA PELA QUAL CONHECEU A DIS."
		end if
	
	if alerta = "" then
		if (r_loja_user_session.unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__AC) And (s_id_magento_b2b <> "") then
			if retorna_so_digitos(s_id_magento_b2b) <> substitui_caracteres(s_id_magento_b2b, ".", "") then
				alerta=texto_add_br(alerta)
				alerta=alerta & "ID Magento B2B informado está em formato inválido"
			elseif converte_numero(s_id_magento_b2b) <= 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "ID Magento B2B informado é inválido"
				end if
			
			if alerta = "" then
				id_magento_b2b = CLng(s_id_magento_b2b)
				if id_magento_b2b <= 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "ID Magento B2B possui valor inválido"
					end if
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
	dim blnErroDuplicidadeCadastro
	blnErroDuplicidadeCadastro=False
	if alerta = "" then
		s = "SELECT" & _
				" apelido," & _
				" cnpj_cpf," & _
				" razao_social_nome," & _
				" loja" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR" & _
			" WHERE" & _
				" (cnpj_cpf = '" & s_cnpj_cpf & "')" & _
				" AND (Convert(smallint, loja) = " & s_loja & ")" & _
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
			alerta="O " & s_label & " " & cnpj_cpf_formata(s_cnpj_cpf) & " já está cadastrado na loja " & s_loja

			do while Not rs.Eof
				alerta=texto_add_br(alerta)
				alerta=alerta & rs("apelido") & " - " & Trim("" & rs("razao_social_nome"))
				rs.MoveNext
				loop
			end if
		
		if rs.State <> 0 then rs.Close
		end if

	if alerta = "" then
		if (r_loja_user_session.unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__AC) And (id_magento_b2b > 0) then
			'VERIFICA SE O ID MAGENTO B2B JÁ ESTÁ EM USO
			s = "SELECT" & _
					" apelido," & _
					" cnpj_cpf," & _
					" razao_social_nome," & _
					" loja" & _
				" FROM t_ORCAMENTISTA_E_INDICADOR" & _
				" WHERE" & _
					" (id_magento_b2b = " & Cstr(id_magento_b2b) & ")"
			set rs = cn.Execute(s)
			if Not rs.Eof then
				blnErroDuplicidadeCadastro=True
				alerta=texto_add_br(alerta)
				alerta=alerta & "O ID Magento B2B " & Cstr(id_magento_b2b) & " já está cadastrado no parceiro " & Trim("" & rs("apelido")) & _
						 " (loja: " & Trim("" & rs("loja")) & _
						 ", CPF/CNPJ: " & cnpj_cpf_formata(Trim("" & rs("cnpj_cpf"))) & _
						 ", nome: " & Trim("" & rs("razao_social_nome")) & ")"
				end if
			if rs.State <> 0 then rs.Close
			end if
		end if

	Err.Clear
	
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_INCLUI
			'=========
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

				s = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido = '" & s_id_selecionado & "'"
				r.Open s, cn
				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("apelido")=s_id_selecionado
					r("dt_cadastro") = Date
					r("usuario_cadastro") = usuario
					r("sistema_responsavel_cadastro") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					end if
					
				if r_loja_user_session.unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__AC then
					if id_magento_b2b > 0 then
						r("id_magento_b2b") = id_magento_b2b
					else
						r("id_magento_b2b") = Null
						end if
					end if

				r("dt_ult_atualizacao") = Now
                r("vendedor_dt_ult_atualizacao") = Date
                r("vendedor_dt_hr_ult_atualizacao") = Now
                r("vendedor_usuario_ult_atualizacao") = usuario
				r("usuario_ult_atualizacao") = usuario
				
				r("tipo") = s_tipo_PJ_PF
				r("razao_social_nome")=s_razao_social_nome
				r("responsavel_principal") = s_responsavel_principal
				r("nome_fantasia") = s_nome_fantasia
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
				r("banco") = s_banco
				r("agencia") = s_agencia
				r("conta") = s_conta
                r("favorecido_cnpj_cpf") = s_favorecido_cnpjcpf
                r("conta_dv") = conta_dv
                r("agencia_dv") = agencia_dv
                r("tipo_conta") = tipo_conta
                r("conta_operacao") = tipo_operacao
				r("favorecido") = s_favorecido
				r("loja") = s_loja
				r("vendedor") = s_vendedor
				r("hab_acesso_sistema")=CLng(s_acesso)
				r("status") = s_status
				
				r("permite_RA_status") = CLng(s_permite_RA_status)
				r("permite_RA_usuario") = usuario
				r("permite_RA_data_hora") = Now
				
				r("datastamp")=""
				r("senha") = ""
					
				r("perc_desagio_RA") = converte_numero(s_perc_desagio_RA)
				r("vl_limite_mensal") = converte_numero(strValorLimiteMensal)
				r("email") = strEmail
				r("email2") = strEmail2
				r("email3") = strEmail3
				r("captador") = ""
				
				r("forma_como_conheceu_codigo") = s_forma_como_conheceu_codigo
				r("forma_como_conheceu_usuario") = usuario
				r("forma_como_conheceu_data") = Date
				r("forma_como_conheceu_data_hora") = Now

                if trim("" & r("datastamp"))<>senha_cripto then
					r("datastamp")=senha_cripto
					r("senha") = gera_senha_aleatoria
					r("dt_ult_alteracao_senha") = Null
					end if
				
				r("obs") = strObs
				
				r("nextel") = s_nextel
				
				if rb_estabelecimento <> "" then r("tipo_estabelecimento") = CLng(rb_estabelecimento)
				
				r("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP

				r.Update

                                ' Vendedores (contatos)
                n = Request.Form("c_indicador_contato").Count
                s_contato_log_inclusao = ""
                for cont = 1 to n
                    s_contato_nome = Trim(Request.Form("c_indicador_contato")(cont))
                    s_contato_id = Trim(Request.Form("contato_id")(cont))
                    s2 = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_CONTATOS WHERE (indicador = '" & s_id_selecionado & "' AND id='" & s_contato_id & "')"
                    rs2.Open s2, cn
                    if s_contato_nome <> "" then
                        if rs2.Eof then
                            if s_contato_log_inclusao = "" then s_contato_log_inclusao = "indicador=" & s_id_selecionado
                            s_contato_log_inclusao = s_contato_log_inclusao & "; nome: " & s_contato_nome
                            rs2.AddNew
                            rs2("indicador") = s_id_selecionado
                            rs2("dt_cadastro") = Date
                            rs2("usuario_cadastro") = usuario
                            end if
                        rs2("nome") = s_contato_nome                    
                        rs2.Update

                        end if
                    if rs2.State <> 0 then rs2.Close
                    next

                ' grava log contatos
                if s_contato_log_inclusao <> "" then grava_log usuario, loja, "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_CONTATOS__INCLUSAO, s_contato_log_inclusao

				If Err = 0 then 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then 
							grava_log usuario, loja, "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_INCLUSAO, s_log
							end if
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if (s_log <> "") then 
							if s_log <> "" then s_log = "; " & s_log
							s_log="apelido=" & Trim("" & r("apelido")) & s_log
							grava_log usuario, loja, "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_ALTERACAO, s_log
							end if
						end if
				else
					erro_fatal=True
					alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
					end if

                ' log cadastrou novo
                    
                    s2 = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_LOG WHERE (id = -1)"
                    rs2.Open s2, cn
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
			'====
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
				s = "INDICADOR " & chr(34) & s_id_selecionado & chr(34) & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "INDICADOR " & chr(34) & s_id_selecionado & chr(34) & " ALTERADO COM SUCESSO."
			end select
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:600px;font-weight:bold;" align="center"><%=s%></div>
<br><br>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
<%
	s="MenuOrcamentistaEIndicador.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<% if blnErroDuplicidadeCadastro then %>
    <td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back();"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<% else %>
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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