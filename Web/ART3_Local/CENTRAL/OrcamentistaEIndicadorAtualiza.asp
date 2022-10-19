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
	
	class cl_RESTRICAO_FORMA_PAGTO
		dim strIdFormaPagto
		dim blnRestricaoAtiva
		end class
	
	dim i
	dim s, s_aux, usuario, senha_cripto, alerta, chave
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	dim blnPossuiPermissaoFinanceiro
	blnPossuiPermissaoFinanceiro = False
	if operacao_permitida(OP_CEN_EDITA_ANALISE_CREDITO, s_lista_operacoes_permitidas) _
		OR _
		operacao_permitida(OP_CEN_PAGTO_PARCIAL, s_lista_operacoes_permitidas) _
		OR _
		operacao_permitida(OP_CEN_PAGTO_QUITACAO, s_lista_operacoes_permitidas) then
		blnPossuiPermissaoFinanceiro = True
		end if

	if (Not operacao_permitida(OP_CEN_CADASTRO_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) And _
	   (Not operacao_permitida(OP_CEN_REL_CHECAGEM_NOVOS_PARCEIROS, s_lista_operacoes_permitidas)) And _
	   (Not blnPossuiPermissaoFinanceiro) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r, rs, t, rs2, s2
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim criou_novo_reg
	Dim s_log, s_log_restricao_FP, s_log_desconto_edicao, s_log_desconto_incl, s_log_desconto_excl
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
    Dim vLog3()
	s_log = ""
	s_log_restricao_FP = ""
    s_log_desconto_excl = ""
    s_log_desconto_incl = ""
    s_log_desconto_edicao = ""
	campos_a_omitir = "|dt_ult_atualizacao|usuario_ult_atualizacao|timestamp|"
	
'	FOI UM RELATÓRIO QUE ORIGINOU A EDIÇÃO DO INDICADOR?
	dim pagina_relatorio_originou_edicao
	pagina_relatorio_originou_edicao = Trim(Request.Form("pagina_relatorio_originou_edicao"))
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim ChecadoStatusBloqueado, blnChecadoStatusBloqueado
	ChecadoStatusBloqueado = Trim(Request.Form("ChecadoStatusBloqueado"))
	blnChecadoStatusBloqueado = CBool(ChecadoStatusBloqueado)
	
	dim valorAux
	dim operacao_selecionada, s_id_selecionado, s_tipo_PJ_PF, s_razao_social_nome, s_nome_fantasia, s_cnpj_cpf, s_ie_rg, s_responsavel_principal
	dim s_endereco, s_endereco_numero, s_endereco_complemento, s_bairro, s_cidade, s_uf, s_cep, s_ddd, s_telefone, s_fax, url_origem
	dim s_ddd_cel, s_tel_cel, s_contato
	dim s_banco, s_agencia, s_conta, s_favorecido
	dim s_senha, s_senha2
	dim s_loja, s_vendedor, s_acesso, s_status, s_permite_RA_status
	dim s_desempenho_nota
	dim s_perc_desagio_RA, strValorLimiteMensal, strValorMeta
	dim strEmail, strEmail2, strEmail3, strCaptador
	dim strChecadoStatus, strObs
	dim s_forma_como_conheceu_codigo
	dim s_nextel, rb_estabelecimento
	dim c_lista_id_forma_pagto, v_lista_id_forma_pagto, v_FP_PF, v_FP_PJ, ckb_value
	dim lngNsuRestricaoFormaPagto, blnRestricaoFPNovo, blnRestricaoFPAlteracao
    dim s_etq_endereco, s_etq_endereco_numero, s_etq_endereco_complemento, s_etq_bairro, s_etq_cidade, s_etq_uf, s_etq_cep, s_etq_ddd_1, s_etq_ddd_2, s_etq_tel_1, s_etq_tel_2, s_etq_email
    dim msg,s_favorecido_cnpjcpf, conta_dv, agencia_dv, tipo_conta, tipo_operacao
    dim n, cont, s_id_desc, s_desc_desc, s_val_desc
    dim s_contato_nome, s_contato_id, s_contato_data, s_contato_log_inclusao, s_contato_log_exclusao, s_contato_log_edicao
	dim ckb_comissao_cartao_status, c_comissao_cartao_cpf, c_comissao_cartao_titular, c_comissao_NFSe_cnpj, c_comissao_NFSe_razao_social
	dim s_id_magento_b2b, s_id_magento_b2b_original, id_magento_b2b

	operacao_selecionada = request("operacao_selecionada")
	s_id_selecionado = UCase(trim(Request.Form("id_selecionado")))
	s_id_magento_b2b = Trim(Request.Form("c_id_magento_b2b"))
	s_id_magento_b2b_original = Trim(Request.Form("c_id_magento_b2b_original"))
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
    s_favorecido_cnpjcpf = retorna_so_digitos(trim(Request.Form("favorecido_cnpjcpf")))
	s_senha=UCase(trim(Request.Form("senha")))
	s_senha2=UCase(trim(Request.Form("senha2")))
	s_loja = trim(Request.Form("loja"))
	s_vendedor = trim(Request.Form("vendedor"))
	s_acesso = trim(Request.Form("rb_acesso"))
	s_status = trim(Request.Form("rb_status"))
	s_permite_RA_status = trim(Request.Form("rb_permite_RA_status"))
	s_desempenho_nota = trim(Request.Form("c_desempenho_nota"))
	s_perc_desagio_RA = trim(Request.Form("c_perc_desagio_RA"))
	strValorLimiteMensal = trim(Request.Form("c_vl_limite_mensal"))
	strValorMeta = trim(Request.Form("c_vl_meta"))
	strEmail = trim(Request.Form("c_email"))
	strEmail2 = trim(Request.Form("c_email2"))
	strEmail3 = trim(Request.Form("c_email3"))
	strCaptador = trim(Request.Form("c_captador"))
	strObs = trim(Request.Form("c_obs"))
	s_forma_como_conheceu_codigo = trim(Request.Form("c_forma_como_conheceu_codigo"))
	s_nextel = trim(Request.Form("c_nextel"))
	rb_estabelecimento = trim(Request.Form("rb_estabelecimento")) 
	c_lista_id_forma_pagto = Trim(Request.Form("c_lista_id_forma_pagto"))
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
    conta_dv = trim(Request.Form("conta_dv"))
    agencia_dv = trim(Request.Form("agencia_dv"))
    tipo_conta = trim(Request.Form("tipo_conta"))
    tipo_operacao = trim(Request.Form("tipo_operacao"))
	ckb_comissao_cartao_status = Trim(Request.Form("ckb_comissao_cartao_status"))
	c_comissao_cartao_cpf = retorna_so_digitos(Trim(Request.Form("c_comissao_cartao_cpf")))
	c_comissao_cartao_titular = Trim(Request.Form("c_comissao_cartao_titular"))
	c_comissao_NFSe_cnpj = retorna_so_digitos(Trim(Request.Form("c_comissao_NFSe_cnpj")))
	c_comissao_NFSe_razao_social = Trim(Request.Form("c_comissao_NFSe_razao_social"))

    url_origem = Request("url_origem")

	v_lista_id_forma_pagto = Split(c_lista_id_forma_pagto, "|")
'	PF
	redim v_FP_PF(0)
	set v_FP_PF(UBound(v_FP_PF)) = New cl_RESTRICAO_FORMA_PAGTO
	v_FP_PF(UBound(v_FP_PF)).strIdFormaPagto = ""
'	PJ
	redim v_FP_PJ(0)
	set v_FP_PJ(UBound(v_FP_PJ)) = New cl_RESTRICAO_FORMA_PAGTO
	v_FP_PJ(UBound(v_FP_PJ)).strIdFormaPagto = ""
'	LAÇO P/ LEITURA DOS CAMPOS
	for i=LBound(v_lista_id_forma_pagto) to UBound(v_lista_id_forma_pagto)
		if Trim(v_lista_id_forma_pagto(i)) <> "" then
		'	PF
			if v_FP_PF(UBound(v_FP_PF)).strIdFormaPagto <> "" then
				redim preserve v_FP_PF(UBound(v_FP_PF)+1)
				set v_FP_PF(UBound(v_FP_PF)) = New cl_RESTRICAO_FORMA_PAGTO
				v_FP_PF(UBound(v_FP_PF)).strIdFormaPagto = ""
				end if
			s = "ckb_" & ID_PF & "_" & Trim(v_lista_id_forma_pagto(i))
			ckb_value = Trim(Request.Form(s))
			v_FP_PF(UBound(v_FP_PF)).strIdFormaPagto = Trim(v_lista_id_forma_pagto(i))
			if ckb_value <> "" then
			'	CHECKBOX MARCADO: RESTRIÇÃO DA FORMA DE PAGTO ESTÁ ATIVA
				v_FP_PF(UBound(v_FP_PF)).blnRestricaoAtiva = True
			else
				v_FP_PF(UBound(v_FP_PF)).blnRestricaoAtiva = False
				end if
		'	PJ
			if v_FP_PJ(UBound(v_FP_PJ)).strIdFormaPagto <> "" then
				redim preserve v_FP_PJ(UBound(v_FP_PJ)+1)
				set v_FP_PJ(UBound(v_FP_PJ)) = New cl_RESTRICAO_FORMA_PAGTO
				v_FP_PJ(UBound(v_FP_PJ)).strIdFormaPagto = ""
				end if
			s = "ckb_" & ID_PJ & "_" & Trim(v_lista_id_forma_pagto(i))
			ckb_value = Trim(Request.Form(s))
			v_FP_PJ(UBound(v_FP_PJ)).strIdFormaPagto = Trim(v_lista_id_forma_pagto(i))
			if ckb_value <> "" then
			'	CHECKBOX MARCADO: RESTRIÇÃO DA FORMA DE PAGTO ESTÁ ATIVA
				v_FP_PJ(UBound(v_FP_PJ)).blnRestricaoAtiva = True
			else
				v_FP_PJ(UBound(v_FP_PJ)).blnRestricaoAtiva = False
				end if
			end if
		next
	
	if blnChecadoStatusBloqueado then
		strChecadoStatus = ""
	else
		strChecadoStatus = trim(Request.Form("rb_checado"))
		end if
	
	
	if s_id_selecionado = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)

	s_loja=normaliza_codigo(s_loja, TAM_MIN_LOJA)

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
	elseif converte_numero(strValorMeta) < 0 then
		alerta="VALOR DA META É INVÁLIDO (" & strValorMeta & ")."
		end if
	
	if alerta = "" then
		if operacao_selecionada = OP_INCLUI then
			if strCaptador = "" then alerta="INFORME QUEM É O CAPTADOR."
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
	
	if (alerta = "") And blnPossuiPermissaoFinanceiro then
		'Dados de pagamento da comissão: Cartão
		if ckb_comissao_cartao_status <> "" then
			'Se o checkbox "Pagamento Via Cartão" estiver assinalado, o preenchimento dos campos é obrigatório
			if (c_comissao_cartao_cpf = "") Or (Not cpf_ok(c_comissao_cartao_cpf)) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Informe o CPF do titular do cartão nos dados para pagamento da comissão"
				end if
			if c_comissao_cartao_titular = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Informe o nome do titular do cartão nos dados para pagamento da comissão"
				end if
		else
			'Se o checkbox "Pagamento Via Cartão" estiver desmarcado, o preenchimento dos campos é opcional de modo a não obrigar a apagar as informações,
			'mas ou todos os campos devem estar preenchidos ou todos devem estar vazios
			if (c_comissao_cartao_cpf <> "") Or (c_comissao_cartao_titular <> "") then
				if (c_comissao_cartao_cpf = "") Or (Not cpf_ok(c_comissao_cartao_cpf)) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Informe o CPF do titular do cartão nos dados para pagamento da comissão"
					end if
				if c_comissao_cartao_titular = "" then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Informe o nome do titular do cartão nos dados para pagamento da comissão"
					end if
				end if
			end if
		
		'Dados de pagamento da comissão: Emitente da NFSe
		if (c_comissao_NFSe_cnpj <> "") Or (c_comissao_NFSe_razao_social <> "") then
			if (c_comissao_NFSe_cnpj = "") Or (Not cnpj_cpf_ok(c_comissao_NFSe_cnpj)) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Informe o CNPJ do emitente da NFSe nos dados para pagamento da comissão"
				end if
			if c_comissao_NFSe_razao_social = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Informe a razão social do emitente da NFSe nos dados para pagamento da comissão"
				end if
			end if
		end if 'if (alerta = "") And blnPossuiPermissaoFinanceiro

	if alerta <> "" then erro_consistencia=True
	
	
	if s_senha <> "" then
		chave = gera_chave(FATOR_BD)
		codifica_dado s_senha, senha_cripto, chave
		end if
	

'	VALIDAÇÃO P/ PERMITIR SOMENTE UM CADASTRO POR LOJA P/ CADA CPF/CNPJ
	dim s_label
	dim r_orcamentista_e_indicador
	if operacao_selecionada <> OP_INCLUI then
		call le_orcamentista_e_indicador(s_id_selecionado, r_orcamentista_e_indicador, msg_erro)
		end if

	dim blnErroDuplicidadeCadastro
	blnErroDuplicidadeCadastro=False
	if alerta = "" then
		if operacao_selecionada = OP_INCLUI then
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
		elseif operacao_selecionada <> OP_EXCLUI then
			' CONSISTE SOMENTE SE ESTIVER ALTERANDO A LOJA OU O CPF/CNPJ
			if converte_numero(s_loja) <> converte_numero(r_orcamentista_e_indicador.loja) Or _
				retorna_so_digitos(s_cnpj_cpf) <> retorna_so_digitos(r_orcamentista_e_indicador.cnpj_cpf) then
				s = "SELECT" & _
						" apelido," & _
						" cnpj_cpf," & _
						" razao_social_nome," & _
						" loja" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (cnpj_cpf = '" & s_cnpj_cpf & "')" & _
						" AND (Convert(smallint, loja) = " & s_loja & ")" & _
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
					alerta=alerta & cnpj_cpf_formata(s_cnpj_cpf) & " já está cadastrado na loja " & s_loja

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

	dim r_loja_indicador_original, r_loja_indicador_form
	set r_loja_indicador_original = New cl_LOJA
	set r_loja_indicador_form = New cl_LOJA
	if alerta = "" then
		if operacao_selecionada <> OP_INCLUI then
			if Not x_loja_bd(r_orcamentista_e_indicador.loja, r_loja_indicador_original) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "A loja cadastrada para o indicador (" & r_orcamentista_e_indicador.loja & ") não foi encontrada"
				end if
			end if
		if s_loja <> "" then
			if Not x_loja_bd(s_loja, r_loja_indicador_form) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "A loja selecionada para o indicador (" & s_loja & ") não foi encontrada"
				end if
			end if
		end if 'if alerta = "" then

	dim blnVisivelIdMagentoB2B
	blnVisivelIdMagentoB2B = False
	if alerta = "" then
		if operacao_selecionada = OP_INCLUI then
			if Trim("" & r_loja_indicador_form.unidade_negocio) = COD_UNIDADE_NEGOCIO_LOJA__AC then
				blnVisivelIdMagentoB2B = True
				end if
		else
			if (Trim("" & r_loja_indicador_form.unidade_negocio) = COD_UNIDADE_NEGOCIO_LOJA__AC) _
				Or (r_loja_indicador_original.unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__AC) _
				Or ( (Trim("" & r_orcamentista_e_indicador.id_magento_b2b) <> "") And (Trim("" & r_orcamentista_e_indicador.id_magento_b2b) <> "0") ) then
				blnVisivelIdMagentoB2B = True
				end if
			end if
		end if 'if alerta = ""

	if alerta = "" then
		if blnVisivelIdMagentoB2B then
			if s_id_magento_b2b <> "" then
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
		end if 'if alerta = "" then

	if alerta = "" then
		if blnVisivelIdMagentoB2B then
			if id_magento_b2b > 0 then
				'VERIFICA SE O ID MAGENTO B2B JÁ ESTÁ EM USO
				s = "SELECT" & _
						" apelido," & _
						" cnpj_cpf," & _
						" razao_social_nome," & _
						" loja" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (id_magento_b2b = " & Cstr(id_magento_b2b) & ")" & _
						" AND (apelido <> '" & s_id_selecionado & "')"
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
				end if 'if id_magento_b2b > 0
			end if 'if blnVisivelIdMagentoB2B
		end if 'if alerta = ""

	'VALIDAÇÃO SE O IDENTIFICADOR JÁ ESTÁ EM USO NO CADASTRO DE USUÁRIOS (ASSEGURA QUE NÃO EXISTA USUÁRIO E INDICADOR COM MESMO IDENTIFICADOR)
	if alerta = "" then
		if operacao_selecionada = OP_INCLUI then
			s = "SELECT usuario, nome FROM t_USUARIO WHERE usuario = '" & s_id_selecionado & "'"
			set rs = cn.Execute(s)
			if Not rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Não é possível usar o identificador '" & s_id_selecionado & "' porque já está em uso no cadastro de usuários"
				end if
			end if
		end if

	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(t, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
    if Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

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
			'	~~~~~~~
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

				s ="DELETE FROM t_ORCAMENTISTA_E_INDICADOR_BLOCO_NOTAS WHERE (apelido = '" & s_id_selecionado & "')"
				cn.Execute(s)
				If Err <> 0 then
					erro_fatal=True
					alerta = "FALHA AO EXCLUIR OS DADOS DE BLOCO DE NOTAS DO ORÇAMENTISTA / INDICADOR (" & Cstr(Err) & ": " & Err.Description & ")."
					end if

				s ="DELETE FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (apelido = '" & s_id_selecionado & "')"
				cn.Execute(s)
				If Err <> 0 then
					erro_fatal=True
					alerta = "FALHA AO EXCLUIR OS DADOS DA PLANILHA DE DESCONTOS DO ORÇAMENTISTA / INDICADOR (" & Cstr(Err) & ": " & Err.Description & ")."
					end if

				s ="DELETE FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO WHERE (id_orcamentista_e_indicador = '" & s_id_selecionado & "')"
				cn.Execute(s)
				If Err <> 0 then
					erro_fatal=True
					alerta = "FALHA AO EXCLUIR OS DADOS DE RESTRIÇÃO NA FORMA DE PAGAMENTO DO ORÇAMENTISTA / INDICADOR (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				
				if alerta = "" then
					s="DELETE FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido = '" & s_id_selecionado & "'"
					cn.Execute(s)
					If Err = 0 then 
						if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_EXCLUSAO, s_log
					else
						erro_fatal=True
						alerta = "FALHA AO REMOVER O ORÇAMENTISTA / INDICADOR (" & Cstr(Err) & ": " & Err.Description & ")."
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
                if operacao_selecionada = OP_CONSULTA then
	            msg = ""
                dim x1, x2
                dim intNsuNovoLog, intNsuNovoDesconto
				
                s2 = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_LOG WHERE (id = -1)"
				
                rs2.Open s2, cn
                if rs2.EOF then 
                    ' ID Magento B2B
					if blnVisivelIdMagentoB2B then
						x1 = Trim("" & r("id_magento_b2b"))
						if x1 = "0" then x1 = ""
						x2 = Cstr(id_magento_b2b)
						if x2 = "0" then x2 = ""
						if (x1 <> x2) then
							if x1 = "" then x1 = "VAZIO"
							if x2 = "" then x2 = "VAZIO"
							msg = msg & "ID Magento B2B alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
							end if
						end if

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
                        msg = msg & "CPF/CNPJ alterado <br>|de: " & cnpj_cpf_formata(r("cnpj_cpf")) & "<br>|para: " & cnpj_cpf_formata(s_cnpj_cpf) & "<br>"
                    end if                   
                    ' log ie rg
                    if (s_ie_rg <> r("ie_rg")) then
                        if s_ie_rg = "" then x2 = "VAZIO" else x2 = s_ie_rg
                        if r("ie_rg") = "" then x1 = "VAZIO" else x1 = r("ie_rg")
                        msg = msg & "IE/RG alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log endereco
                    if (s_endereco <> r("endereco")) then
                      msg = msg & "Endereço alterado <br>|de: " & r("endereco") & "<br>|para: " & s_endereco & "<br>"
                    end if

                    ' log endereco numero
                    if (s_endereco_numero <> Trim("" & r("endereco_numero"))) then
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
                    if isNull(r("cidade")) then r("cidade") = ""
                    if (s_cidade <> r("cidade")) then
                        msg = msg & "Cidade alterada <br>|de: " & r("cidade") & "<br>|para: " & s_cidade & "<br>"
                    end if

                    ' log uf
                    if (s_uf <> r("uf")) then
                        msg = msg & "UF alterado <br>|de: " & r("uf") & "<br>|para: " & s_uf & "<br>"
                    end if

                    ' log cep
                    if isNull(r("cep")) then r("cep") = ""
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

                    ' log banco
                    if isNull(r("banco")) then r("banco") = ""
                    if (s_banco <> r("banco")) then
                        if s_banco = "" then x2 = "VAZIO" else x2 = s_banco
                        if r("banco") = "" then x1 = "VAZIO" else x1 = r("banco")
                        msg = msg & "Banco alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log agencia
                    if isNull(r("agencia")) then r("agencia") = ""
                    if isNull(r("agencia_dv")) then r("agencia_dv") = ""
                    if (s_agencia <> r("agencia") Or agencia_dv <> r("agencia_dv")) then
                        if s_agencia = "" then x2 = "VAZIO" else x2 = s_agencia
                        if agencia_dv <> "" then x2 = x2 & "-" & agencia_dv
                        if r("agencia") = "" then x1 = "VAZIO" else x1 = r("agencia")
                        if r("agencia_dv") <> "" then x1 = x1 & "-" & r("agencia_dv")
                        msg = msg & "Agência alterada <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log conta
                    if isNull(r("conta")) then r("conta") = ""
                    if isNull(r("conta_dv")) then r("conta_dv") = ""
                    if (s_conta <> r("conta") Or conta_dv <> r("conta_dv")) then
                        if s_conta = "" then x2 = "VAZIO" else x2 = s_conta
                        if conta_dv <> "" then x2 = x2 & "-" & conta_dv
                        if r("conta") = "" then x1 = "VAZIO" else x1 = r("conta") 
                        if r("conta_dv") <> "" then x1 = x1 & "-" & r("conta_dv")
                        msg = msg & "Conta alterada <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log tipo conta
                    if Trim("" & r("tipo_conta") <> tipo_conta) then
                        function tipoConta(x) 
                            dim s
                            select case x
                                case ""
                                    s = "VAZIO"
                                case "P"  
                                    s = "Poupança"
                                case "C"
                                    s = "Corrente"
                                case else
                                    s = "VAZIO"
                                end select
                            tipoConta = s
                        end function
                        if tipo_conta = "" then x2 = "VAZIO" else x2 = tipoConta(tipo_conta)
                        if Trim("" & r("tipo_conta")) = "" then x1 = "VAZIO" else x1 = tipoConta(r("tipo_conta"))
                        msg = msg & "Tipo de conta alterada <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log tipo operação
                    if Trim("" & r("conta_operacao") <> tipo_operacao) then
                        if tipo_operacao = "" then x2 = "VAZIO" else x2 = tipo_operacao
                        if Trim("" & r("conta_operacao")) = "" then x1 = "VAZIO" else x1 = r("conta_operacao")
                        msg = msg & "Tipo de operação alterada <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log favorecido
                    if isNull(r("favorecido")) then r("favorecido") = ""
                    if (s_favorecido <> r("favorecido")) then
                        if s_favorecido = "" then x2 = "VAZIO" else x2 = s_favorecido
                        if r("favorecido") = "" then x1 = "VAZIO" else x1 = r("favorecido")
                        msg = msg & "Favorecido alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log cnpj cpf favorecido
                    if (s_favorecido_cnpjcpf <> Trim("" & r("favorecido_cnpj_cpf"))) then
                        if s_favorecido_cnpjcpf = "" then x2 = "VAZIO" else x2 = cnpj_cpf_formata(s_favorecido_cnpjcpf)
                        if Trim("" & r("favorecido_cnpj_cpf")) = "" then x1 = "VAZIO" else x1 = cnpj_cpf_formata(r("favorecido_cnpj_cpf"))
                        msg = msg & "CPF/CNPJ do favorecido alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if    
                    
					'Dados para pagamento da comissão
					if blnPossuiPermissaoFinanceiro then
						'Checkbox Pagamento Via Cartão
						if ckb_comissao_cartao_status = "" then valorAux = 0 else valorAux = 1
						if r("comissao_cartao_status") <> valorAux then
							if r("comissao_cartao_status") = 0 then x1 = "Desmarcado" else x1 = "Marcado"
							if ckb_comissao_cartao_status = "" then x2 = "Desmarcado" else x2 = "Marcado"
							msg = msg & "[Pagamento da Comissão] Pagamento Via Cartão alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
							end if
						'CPF
						if c_comissao_cartao_cpf <> Trim("" & r("comissao_cartao_cpf")) then
							if Trim("" & r("comissao_cartao_cpf")) = "" then x1 = "VAZIO" else x1 = cnpj_cpf_formata(Trim("" & r("comissao_cartao_cpf")))
							if c_comissao_cartao_cpf = "" then x2 = "VAZIO" else x2 = cnpj_cpf_formata(c_comissao_cartao_cpf)
							msg = msg & "[Pagamento da Comissão] CPF do Titular do Cartão alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
							end if
						'Nome do Titular do Cartão
						if UCase(c_comissao_cartao_titular) <> UCase(Trim("" & r("comissao_cartao_titular"))) then
							if Trim("" & r("comissao_cartao_titular")) = "" then x1 = "VAZIO" else x1 = Trim("" & r("comissao_cartao_titular"))
							if c_comissao_cartao_titular = "" then x2 = "VAZIO" else x2 = c_comissao_cartao_titular
							msg = msg & "[Pagamento da Comissão] Nome do Titular do Cartão alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
							end if
						'CNPJ do Emitente da NFSe
						if c_comissao_NFSe_cnpj <> Trim("" & r("comissao_NFSe_cnpj")) then
							if Trim("" & r("comissao_NFSe_cnpj")) = "" then x1 = "VAZIO" else x1 = cnpj_cpf_formata(Trim("" & r("comissao_NFSe_cnpj")))
							if c_comissao_NFSe_cnpj = "" then x2 = "VAZIO" else x2 = cnpj_cpf_formata(c_comissao_NFSe_cnpj)
							msg = msg & "[Pagamento da Comissão] CNPJ do Emitente da NFSe alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
							end if
						'Razão Social do Emitente da NFSe
						if UCase(c_comissao_NFSe_razao_social) <> UCase(Trim("" & r("comissao_NFSe_razao_social"))) then
							if Trim("" & r("comissao_NFSe_razao_social")) = "" then x1 = "VAZIO" else x1 = Trim("" & r("comissao_NFSe_razao_social"))
							if c_comissao_NFSe_razao_social = "" then x2 = "VAZIO" else x2 = c_comissao_NFSe_razao_social
							msg = msg & "[Pagamento da Comissão] Razão Social do Emitente da NFSe alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
							end if
						end if

                    ' log status
                    if (s_status <> r("status")) then
                        function tipoStatus(x) 
                            dim s
                            select case x
                                case ""
                                    s = "VAZIO"
                                case "A"  
                                    s = "Ativo"
                                case "I"
                                    s = "Inativo"
                                case else
                                    s = "VAZIO"
                                end select
                            tipoStatus = s
                        end function
                        msg = msg & "Status alterado <br>|de: " & tipoStatus(r("status")) & "<br>|para: " & tipoStatus(s_status) & "<br>"
                    end if

                    ' log permite RA status
                    if (s_permite_RA_status = "") then s_permite_RA_status = -1
                    if (CLng(s_permite_RA_status) <> r("permite_RA_status")) then
                        function tipoRAStatus(num) 
                            dim s
                            select case num
                                case 1  
                                    s = "Sim"
                                case 0
                                    s = "Não"
                                case else
                                    s = "VAZIO"
                                end select
                            tipoRAStatus = s
                        end function
                        msg = msg & "Status 'Permite RA' alterado <br>|de: " & tipoRAStatus(r("permite_RA_status")) & "<br>|para: " & tipoRAStatus(s_permite_RA_status) & "<br>"
                    end if

                    ' log desempenho nota
                    if isNull(r("desempenho_nota")) then r("desempenho_nota") = ""
                    if (s_desempenho_nota <> r("desempenho_nota")) then
                        if s_desempenho_nota = "" then x2 = "VAZIO" else x2 = s_desempenho_nota
                        if r("desempenho_nota") = "" then x1 = "VAZIO" else x1 = r("desempenho_nota")
                        msg = msg & "Avaliação Desempenho alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log loja
                    if isNull(r("loja")) then r("loja") = ""
                    if (s_loja <> r("loja")) then
                        if s_loja = "" then x2 = "VAZIO" else x2 = s_loja
                        if r("loja") = "" then x1 = "VAZIO" else x1 = r("loja")
                        msg = msg & "Loja alterada <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log vendedor
                    if isNull(r("vendedor")) then r("vendedor") = ""
                    if (s_vendedor <> r("vendedor")) then
                        if s_vendedor = "" then x2 = "VAZIO" else x2 = s_vendedor
                        if r("vendedor") = "" then x1 = "VAZIO" else x1 = r("vendedor")
                        msg = msg & "Vendedor alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log perc desagio RA
                    if isNull(r("perc_desagio_RA")) then r("perc_desagio_RA") = ""
                    if (converte_numero(s_perc_desagio_RA) <> r("perc_desagio_RA")) then
                        if s_perc_desagio_RA = "" then x2 = "VAZIO" else x2 = s_perc_desagio_RA
                        if r("perc_desagio_RA") = "" then x1 = "VAZIO" else x1 = r("perc_desagio_RA")
                        msg = msg & "Percentual Deságio do RA alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log vl meta
                    if isNull(r("vl_meta")) then r("vl_meta") = ""
                    if (converte_numero(strValorMeta) <> r("vl_meta")) then
                        if strValorMeta = "" then x2 = "VAZIO" else x2 = strValorMeta
                        if r("vl_meta") = "" then x1 = "VAZIO" else x1 = r("vl_meta")
                        msg = msg & "Valor da meta alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log captador
                    if isNull(r("captador")) then r("captador") = ""
                    if (strCaptador <> r("captador")) then
                        if strCaptador = "" then x2 = "VAZIO" else x2 = strCaptador
                        if r("captador") = "" then x1 = "VAZIO" else x1 = r("captador")
                        msg = msg & "Captador alterado <br>|de: " & x1 & "<br>|para: " & x2 & "<br>"
                    end if

                    ' log forma como conheceu a bonshop
                    if isNull(r("forma_como_conheceu_codigo")) then r("forma_como_conheceu_codigo") = ""
                    if (s_forma_como_conheceu_codigo <> r("forma_como_conheceu_codigo")) then
                        if (obtem_descricao_tabela_t_codigo_descricao("CadOrcamentistaEIndicador_FormaComoConheceu", r("forma_como_conheceu_codigo"))) = "Código não cadastrado ()" then x1 = "VAZIO" else x1 = obtem_descricao_tabela_t_codigo_descricao("CadOrcamentistaEIndicador_FormaComoConheceu", r("forma_como_conheceu_codigo"))                 
                        msg = msg & "Forma como conheceu a Bonshop alterada <br>|de: " & x1 & "<br>|para: " & obtem_descricao_tabela_t_codigo_descricao("CadOrcamentistaEIndicador_FormaComoConheceu", s_forma_como_conheceu_codigo) & "<br>"
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

                    ' log observações
                    if isNull(r("obs")) then r("obs") = ""
                    if (strObs <> r("obs")) then
                        if strObs = "" then x2 = "VAZIO" else x2 = strObs
                        if r("obs") = "" then x1 = "VAZIO" else x1 = r("obs")
                        msg = msg & "Campo 'observações' alterado <br>|de: " & substitui_caracteres(x1, chr(13), "<br>") & "<br>|para: " & substitui_caracteres(x2, chr(13), "<br>") & "<br>"
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
                        rs2("loja") = ""
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
				
				if blnVisivelIdMagentoB2B then
					if id_magento_b2b > 0 then
						r("id_magento_b2b") = id_magento_b2b
					else
						r("id_magento_b2b") = Null
						end if
					end if

				r("dt_ult_atualizacao") = Now
				r("usuario_ult_atualizacao") = usuario
				if s_vendedor <> r("vendedor") then
                    r("vendedor_dt_ult_atualizacao") = Date
                    r("vendedor_dt_hr_ult_atualizacao") = Now
                    r("vendedor_usuario_ult_atualizacao") = usuario
                end if
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
				r("banco") = s_banco
				r("agencia") = s_agencia
				r("conta") = s_conta
				r("favorecido") = s_favorecido
				r("loja") = s_loja                
				r("vendedor") = s_vendedor
				r("hab_acesso_sistema")=CLng(s_acesso)
				r("status") = s_status
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
				r("favorecido_cnpj_cpf") = s_favorecido_cnpjcpf
                r("conta_dv") = conta_dv
                r("agencia_dv") = agencia_dv
                r("tipo_conta") = tipo_conta
                r("conta_operacao") = tipo_operacao

				if CLng(r("permite_RA_status")) <> CLng(s_permite_RA_status) then
					r("permite_RA_status") = CLng(s_permite_RA_status)
					r("permite_RA_usuario") = usuario
					r("permite_RA_data_hora") = Now
					end if
				
				if trim("" & r("datastamp"))<>senha_cripto then
					r("datastamp")=senha_cripto
					r("senha") = gera_senha_aleatoria
					r("dt_ult_alteracao_senha") = Null
					end if
				
				if s_desempenho_nota <> Trim("" & r("desempenho_nota")) then
					r("desempenho_nota") = s_desempenho_nota
					r("desempenho_nota_data") = Now
					r("desempenho_nota_usuario") = usuario
					end if
				
				r("perc_desagio_RA") = converte_numero(s_perc_desagio_RA)
				r("vl_limite_mensal") = converte_numero(strValorLimiteMensal)
				r("email") = strEmail
				r("email2") = strEmail2
				r("email3") = strEmail3
				r("captador") = strCaptador
				
				if converte_numero(Trim("" & r("vl_meta"))) <> converte_numero(strValorMeta) then
					r("vl_meta") = converte_numero(strValorMeta)
					r("UsuarioUltAtualizVlMeta") = usuario
					r("DtHrUltAtualizVlMeta") = Now
					end if
				
				if (Not blnChecadoStatusBloqueado) then 
					if strChecadoStatus <> "" then
						if CLng(strChecadoStatus) <> r("checado_status") then
							r("checado_status") = CLng(strChecadoStatus)
							r("checado_data") = Now
							r("checado_usuario") = usuario
							end if
						end if
					end if
				
				if Trim("" & r("forma_como_conheceu_codigo")) <> s_forma_como_conheceu_codigo then
					r("forma_como_conheceu_codigo_anterior") = r("forma_como_conheceu_codigo")
					r("forma_como_conheceu_codigo") = s_forma_como_conheceu_codigo
					r("forma_como_conheceu_usuario") = usuario
					r("forma_como_conheceu_data") = Date
					r("forma_como_conheceu_data_hora") = Now
					end if
				
				r("obs") = strObs
				
				r("nextel") = s_nextel
				
				if rb_estabelecimento <> "" then r("tipo_estabelecimento") = CLng(rb_estabelecimento)

				if blnPossuiPermissaoFinanceiro then
					if ckb_comissao_cartao_status = "" then valorAux = 0 else valorAux = 1
					r("comissao_cartao_status") = valorAux
					r("comissao_cartao_cpf") = c_comissao_cartao_cpf
					r("comissao_cartao_titular") = c_comissao_cartao_titular
					r("comissao_NFSe_cnpj") = c_comissao_NFSe_cnpj
					r("comissao_NFSe_razao_social") = c_comissao_NFSe_razao_social
					end if

				r("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP

				r.Update
                
            ' log cadastrou novo
                    if operacao_selecionada = OP_INCLUI then
                    
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
                        rs2("loja") = ""
                        rs2("mensagem") = "Indicador cadastrado"
                        rs2("usuario") = usuario
                    rs2.Update
                    end if
                    if rs2.State <> 0 then rs2.Close

				end if
				
            ' Vendedores/contatos (inclusão de novo cadastro)
            if operacao_selecionada = OP_INCLUI then
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
                            s = "DELETE FROM t_ORCAMENTISTA_E_INDICADOR_CONTATOS WHERE (id = " & s_contato_id & ")"
                            cn.Execute(s)
                            end if
                        end if
                    if s_contato_nome <> "" then
                        if t.Eof then
                            if s_contato_log_inclusao = "" then s_contato_log_inclusao = "indicador=" & s_id_selecionado
                            s_contato_log_inclusao = s_contato_log_inclusao & "; nome: " & s_contato_nome
                            t.AddNew
                            t("indicador") = s_id_selecionado
                            t("dt_cadastro") = Date
                            t("usuario_cadastro") = usuario
                        else
                            if s_contato_nome <> Trim("" & t("nome")) then
                                if s_contato_log_edicao = "" then s_contato_log_edicao = "indicador=" & s_id_selecionado
                                s_contato_log_edicao = s_contato_log_edicao & "; nome anterior: " & Trim("" & t("nome")) & "; nome atual: " & s_contato_nome
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
                end if

            ' grava tabela de desconto
            
            n = Request.Form("desc_descricao").Count
            for cont = 1 to n
                s_id_desc = Trim(Request.Form("id_desc")(cont))
                s_desc_desc = Trim(Request.Form("desc_descricao")(cont))
                s_val_desc = Trim(Request.Form("desc_valor")(cont))

            s2 = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (apelido = '" & s_id_selecionado & "' AND id='" & s_id_desc & "')"
            rs2.Open s2, cn
           
            if Not rs2.EOF then
                if Trim("" & rs2("descricao")) <> "" And s_desc_desc = "" then 
                        if s_log_desconto_excl = "" then s_log_desconto_excl = "apelido=" & s_id_selecionado
                        s_log_desconto_excl = s_log_desconto_excl & "; descrição: " & rs2("descricao") & "; valor: " & formata_moeda(rs2("valor"))
                        
                        s = "DELETE t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (id = '" & s_id_desc & "')"
                        cn.Execute(s)
                end if
            end if

           if s_desc_desc <> "" then
               if rs2.EOF then
                
                    if Not fin_gera_nsu(T_ORCAMENTISTA_E_INDICADOR_DESCONTO, intNsuNovoDesconto, msg_erro) then 
			                    alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		                    else
			                    if intNsuNovoDesconto <= 0 then
				                    alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovoDesconto & ")"
				                end if
			        end if
                    if s_log_desconto_incl = "" then s_log_desconto_incl = "apelido=" & s_id_selecionado
                    s_log_desconto_incl = s_log_desconto_incl & "; descrição: " & s_desc_desc & "; valor: " & formata_moeda(s_val_desc)
                       
                    rs2.AddNew
                    rs2("id") = intNsuNovoDesconto
                    rs2("apelido") = s_id_selecionado
                else

                    if (rs2("valor")<>converte_numero(s_val_desc) And rs2("descricao") <> s_desc_desc) then
                        if s_log_desconto_edicao = "" then s_log_desconto_edicao = "apelido=" & s_id_selecionado
                        s_log_desconto_edicao = s_log_desconto_edicao & "; descrição: " & rs2("descricao") & " => " & s_desc_desc & "; valor: " & formata_moeda(rs2("valor")) & " => " & formata_moeda(s_val_desc)
                    elseif (rs2("descricao") <> s_desc_desc) then
                        if s_log_desconto_edicao = "" then s_log_desconto_edicao = "apelido=" & s_id_selecionado
                        s_log_desconto_edicao = s_log_desconto_edicao & "; descrição: " & rs2("descricao") & " => " & s_desc_desc & "; valor: " & formata_moeda(rs2("valor"))

                    elseif (rs2("valor")<>converte_numero(s_val_desc)) then
                        if s_log_desconto_edicao = "" then s_log_desconto_edicao = "apelido=" & s_id_selecionado
                        s_log_desconto_edicao = s_log_desconto_edicao & "; descrição: " & rs2("descricao") & "; valor: " & formata_moeda(rs2("valor")) & " => " & formata_moeda(s_val_desc)

                   
                        
                    end if
                end if
                    
                    rs2("usuario") = usuario
                    rs2("ordenacao") = cont
                    rs2("descricao") = s_desc_desc
                    rs2("valor") = converte_numero(s_val_desc)
                    rs2.Update
               
            
            end if
            if rs2.State <> 0 then rs2.Close
            next
            
            ' grava log tabela de descontos
            if s_log_desconto_edicao <> "" then grava_log usuario, "", "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_TABELA_DESCONTO_EDICAO, s_log_desconto_edicao
            if s_log_desconto_incl <> "" then grava_log usuario, "", "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_TABELA_DESCONTO_INCLUSAO, s_log_desconto_incl
            if s_log_desconto_excl <> "" then  grava_log usuario, "", "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_TABELA_DESCONTO_EXCL, s_log_desconto_excl

			'	RESTRIÇÕES NA FORMA DE PAGTO P/ CLIENTE PF
				for i=LBound(v_FP_PF) to UBound(v_FP_PF)
					if Trim(v_FP_PF(i).strIdFormaPagto) <> "" then
						blnRestricaoFPNovo = False
						blnRestricaoFPAlteracao = False
						s = "SELECT " & _
								"*" & _
							" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO" & _
							" WHERE" & _
								" (id_orcamentista_e_indicador = '" & s_id_selecionado & "')" & _
								" AND (tipo_cliente = '" & ID_PF & "')" & _
								" AND (id_forma_pagto = " & v_FP_PF(i).strIdFormaPagto & ")"
						if t.State <> 0 then t.Close
						t.Open s, cn
						if t.Eof then
						'	SE NÃO EXISTE NENHUM REGISTRO, CRIA UM NOVO APENAS NO CASO DE CADASTRAR UMA RESTRIÇÃO ATIVA
							if v_FP_PF(i).blnRestricaoAtiva then
								blnRestricaoFPNovo = True
								t.AddNew
								if Not fin_gera_nsu(T_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO, lngNsuRestricaoFormaPagto, msg_erro) then
									alerta = "FALHA AO GERAR NSU PARA GRAVAR O REGISTRO DA RESTRIÇÃO DA FORMA DE PAGAMENTO (" & msg_erro & ")"
									exit for
									end if
								t("id") = lngNsuRestricaoFormaPagto
								t("id_orcamentista_e_indicador") = s_id_selecionado
								t("id_forma_pagto") = CInt(v_FP_PF(i).strIdFormaPagto)
								t("tipo_cliente") = ID_PF
								t("usuario_cadastro") = usuario
								end if
							end if
						
					'	HOUVE ALTERAÇÃO?
						if (Not blnRestricaoFPNovo) And (Not t.Eof) then
							if v_FP_PF(i).blnRestricaoAtiva then
								if t("st_restricao_ativa") = 0 then blnRestricaoFPAlteracao = true
							else
								if t("st_restricao_ativa") <> 0 then blnRestricaoFPAlteracao = true
								end if
							end if
						
					'	LOG
						if blnRestricaoFPNovo then
							if v_FP_PF(i).blnRestricaoAtiva then
								if s_log_restricao_FP <> "" then s_log_restricao_FP = s_log_restricao_FP & "; "
								s_log_restricao_FP = s_log_restricao_FP & x_opcao_forma_pagamento(v_FP_PF(i).strIdFormaPagto) & "(" & ID_PF & ")[novo]: Bloqueado"
								end if
						elseif blnRestricaoFPAlteracao then
							if s_log_restricao_FP <> "" then s_log_restricao_FP = s_log_restricao_FP & "; "
							s_log_restricao_FP = s_log_restricao_FP & x_opcao_forma_pagamento(v_FP_PF(i).strIdFormaPagto) & "(" & ID_PF & ")[alteração]: "
							if v_FP_PF(i).blnRestricaoAtiva then
								s_log_restricao_FP = s_log_restricao_FP & "Bloqueado"
							else
								s_log_restricao_FP = s_log_restricao_FP & "Liberado"
								end if
							end if
						
						if blnRestricaoFPNovo Or blnRestricaoFPAlteracao then
							if v_FP_PF(i).blnRestricaoAtiva then
								t("st_restricao_ativa") = 1
							else
								t("st_restricao_ativa") = 0
								end if
							t("dt_ult_atualizacao") = Date
							t("dt_hr_ult_atualizacao") = Now
							t("usuario_ult_atualizacao") = usuario
							t.Update
							end if
						end if
					next
				
			'	RESTRIÇÕES NA FORMA DE PAGTO P/ CLIENTE PJ
				for i=LBound(v_FP_PJ) to UBound(v_FP_PJ)
					if Trim(v_FP_PJ(i).strIdFormaPagto) <> "" then
						blnRestricaoFPNovo = False
						blnRestricaoFPAlteracao = False
						s = "SELECT " & _
								"*" & _
							" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO" & _
							" WHERE" & _
								" (id_orcamentista_e_indicador = '" & s_id_selecionado & "')" & _
								" AND (tipo_cliente = '" & ID_PJ & "')" & _
								" AND (id_forma_pagto = " & v_FP_PJ(i).strIdFormaPagto & ")"
						if t.State <> 0 then t.Close
						t.Open s, cn
						if t.Eof then
						'	SE NÃO EXISTE NENHUM REGISTRO, CRIA UM NOVO APENAS NO CASO DE CADASTRAR UMA RESTRIÇÃO ATIVA
							if v_FP_PJ(i).blnRestricaoAtiva then
								blnRestricaoFPNovo = True
								t.AddNew
								if Not fin_gera_nsu(T_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO, lngNsuRestricaoFormaPagto, msg_erro) then
									alerta = "FALHA AO GERAR NSU PARA GRAVAR O REGISTRO DA RESTRIÇÃO DA FORMA DE PAGAMENTO (" & msg_erro & ")"
									exit for
									end if
								t("id") = lngNsuRestricaoFormaPagto
								t("id_orcamentista_e_indicador") = s_id_selecionado
								t("id_forma_pagto") = CInt(v_FP_PJ(i).strIdFormaPagto)
								t("tipo_cliente") = ID_PJ
								t("usuario_cadastro") = usuario
								end if
							end if
						
					'	HOUVE ALTERAÇÃO?
						if (Not blnRestricaoFPNovo) And (Not t.Eof) then
							if v_FP_PJ(i).blnRestricaoAtiva then
								if t("st_restricao_ativa") = 0 then blnRestricaoFPAlteracao = true
							else
								if t("st_restricao_ativa") <> 0 then blnRestricaoFPAlteracao = true
								end if
							end if
						
					'	LOG
						if blnRestricaoFPNovo then
							if v_FP_PJ(i).blnRestricaoAtiva then
								if s_log_restricao_FP <> "" then s_log_restricao_FP = s_log_restricao_FP & "; "
								s_log_restricao_FP = s_log_restricao_FP & x_opcao_forma_pagamento(v_FP_PJ(i).strIdFormaPagto) & "(" & ID_PJ & ")[novo]: Bloqueado"
								end if
						elseif blnRestricaoFPAlteracao then
							if s_log_restricao_FP <> "" then s_log_restricao_FP = s_log_restricao_FP & "; "
							s_log_restricao_FP = s_log_restricao_FP & x_opcao_forma_pagamento(v_FP_PJ(i).strIdFormaPagto) & "(" & ID_PJ & ")[alteração]: "
							if v_FP_PJ(i).blnRestricaoAtiva then
								s_log_restricao_FP = s_log_restricao_FP & "Bloqueado"
							else
								s_log_restricao_FP = s_log_restricao_FP & "Liberado"
								end if
							end if
						
						if blnRestricaoFPNovo Or blnRestricaoFPAlteracao then
							if v_FP_PJ(i).blnRestricaoAtiva then
								t("st_restricao_ativa") = 1
							else
								t("st_restricao_ativa") = 0
								end if
							t("dt_ult_atualizacao") = Date
							t("dt_hr_ult_atualizacao") = Now
							t("usuario_ult_atualizacao") = usuario
							t.Update
							end if
						end if
					next
			
				if t.State <> 0 then t.Close
				set t = nothing
				
				If Err = 0 then 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if (s_log <> "") And (s_log_restricao_FP <> "") then s_log = s_log & "; "
						if s_log_restricao_FP <> "" then s_log = s_log & "Restrições na forma de pagamento: " & s_log_restricao_FP
						if s_log <> "" then 
							grava_log usuario, "", "", "", OP_LOG_ORCAMENTISTA_E_INDICADOR_INCLUSAO, s_log
							end if
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if (s_log <> "") And (s_log_restricao_FP <> "") then s_log = s_log & "; "
						if s_log_restricao_FP <> "" then s_log = s_log & "Restrições na forma de pagamento: " & s_log_restricao_FP
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

<table cellSpacing="0">
<tr>
<%
	s="MenuOrcamentistaEIndicador.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<% if blnErroDuplicidadeCadastro then %>
    <td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back();"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
    <% elseif url_origem <> "" then %>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=url_origem%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
    <% else %>
    <td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back();"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
    <%end if %>
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