<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  C L I E N T E A T U A L I Z A . A S P
'     ===========================================
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

    Const TEL_BONSHOP_1 = "1139344400"
    Const TEL_BONSHOP_2 = "1139344420"
    Const TEL_BONSHOP_3 = "1139344411"

	On Error GoTo 0
	Err.Clear
	
'	EXIBIÇÃO DE BOTÕES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True
	
	
	dim intIdx, intCounter
	dim s, s_aux, usuario, loja, alerta, exibir_botao_novo_item
	exibir_botao_novo_item = False
	
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim msg_erro, msg_erro_aux
	dim cn, r, tMAP_XML, tMAP_END_ETG, tMAP_END_COB
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim blnLojaHabilitadaProdCompostoECommerce
	blnLojaHabilitadaProdCompostoECommerce = isLojaHabilitadaProdCompostoECommerce(loja)

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	Dim criou_novo_reg_cliente, criou_novo_reg_aux
	Dim s_log, s_log_aux
	Dim campos_a_omitir, campos_a_omitir_ref_bancaria
	Dim campos_a_omitir_ref_comercial, campos_a_omitir_ref_profissional
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = "|dt_cadastro|usuario_cadastro|dt_ult_atualizacao|usuario_ult_atualizacao|"
	campos_a_omitir_ref_bancaria = "|id_cliente|ordem|excluido_status|dt_cadastro|usuario_cadastro|"
	campos_a_omitir_ref_comercial = "|id_cliente|ordem|excluido_status|dt_cadastro|usuario_cadastro|"
	campos_a_omitir_ref_profissional = "|id_cliente|ordem|excluido_status|dt_cadastro|usuario_cadastro|"
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
'	CADASTRAMENTO AUTOMÁTICO DE TRANSPORTADORA BASEADO NO CEP
	dim s_cep_original, s_cep_novo, s_transp_id_auto_novo, s_log_transp_auto
	
	dim s_nome_original
	s_nome_original = ""

'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, cliente_selecionado, cnpj_cpf_selecionado, s_nome, s_ie, s_rg, s_sexo
	dim s_contribuinte_icms, s_produtor_rural, s_contribuinte_icms_cadastrado, s_produtor_rural_cadastrado
	dim s_endereco, s_endereco_numero, s_endereco_complemento, s_bairro, s_cidade, s_uf, s_cep
	dim s_ddd_res, s_tel_res, s_ddd_com, s_tel_com, s_ramal_com, s_contato, s_dt_nasc, s_filiacao, s_obs_crediticias, s_midia, s_email, s_email_xml, s_email_boleto
	dim s_indicador, strCampoIndicadorEditavel
	dim eh_cpf
	dim pagina_retorno
	dim strScript
	dim s_tel_com_2, s_ddd_com_2, s_tel_cel, s_ddd_cel, s_ramal_com_2
	dim s_cliente_tipo

	operacao_selecionada=request("operacao_selecionada")
	cliente_selecionado=retorna_so_digitos(trim(request("cliente_selecionado")))
	cnpj_cpf_selecionado=retorna_so_digitos(trim(request("cnpj_cpf_selecionado")))
	s_nome=elimina_html_entities(Trim(request("nome")))
	s_ie=elimina_html_entities(Trim(request("ie")))
	s_rg=elimina_html_entities(Trim(request("rg")))
	s_sexo=Trim(request("sexo"))
	s_produtor_rural_cadastrado=Trim(request("produtor_rural_cadastrado"))
	s_produtor_rural=Trim(request("rb_produtor_rural"))
	s_contribuinte_icms_cadastrado=Trim(request("contribuinte_icms_cadastrado"))
	s_contribuinte_icms=Trim(request("rb_contribuinte_icms"))
	if (s_produtor_rural = COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) Then
		s_contribuinte_icms=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_INICIAL
		s_ie = ""
		end if
	s_endereco=elimina_html_entities(Trim(request("endereco")))
	s_endereco_numero=Trim(request("endereco_numero"))
	s_endereco_complemento=elimina_html_entities(Trim(request("endereco_complemento")))
	s_bairro=elimina_html_entities(Trim(request("bairro")))
	s_cidade=elimina_html_entities(Trim(request("cidade")))
	s_uf=Ucase(Trim(request("uf")))
	s_cep=retorna_so_digitos(Trim(request("cep")))
	s_ddd_res=retorna_so_digitos(Trim(request("ddd_res")))
	s_tel_res=retorna_so_digitos(Trim(request("tel_res")))
	s_ddd_com=retorna_so_digitos(Trim(request("ddd_com")))
	s_tel_com=retorna_so_digitos(Trim(request("tel_com")))
	s_ramal_com=retorna_so_digitos(Trim(request("ramal_com")))
	s_contato=Trim(request("contato"))
	s_dt_nasc=Trim(request("dt_nasc"))
	s_filiacao=Trim(request("filiacao"))
	s_obs_crediticias=Trim(request("obs_crediticias"))
	s_midia=retorna_so_digitos(Trim(request("midia")))
	s_email=LCase(Trim(request("email")))
	s_email_xml=LCase(Trim(request("email_xml")))
	s_email_boleto=LCase(Trim(request("email_boleto")))
	s_indicador=Trim(request("indicador"))
	strCampoIndicadorEditavel=Trim(request("CampoIndicadorEditavel"))
	s_tel_com_2=retorna_so_digitos(Trim(request("tel_com_2")))
	s_ddd_com_2=retorna_so_digitos(Trim(request("ddd_com_2")))
	s_tel_cel=retorna_so_digitos(Trim(request("tel_cel")))
	s_ddd_cel=retorna_so_digitos(Trim(request("ddd_cel")))
	s_ramal_com_2=retorna_so_digitos(Trim(request("ramal_com_2")))
	eh_cpf=(len(cnpj_cpf_selecionado)=11)
	if eh_cpf then s_cliente_tipo=ID_PF else s_cliente_tipo=ID_PJ

	pagina_retorno = Trim(request("pagina_retorno"))
	if pagina_retorno <> "" then
		if Instr(pagina_retorno, "?") > 0 then
			'A QueryString já tem pelo menos um parâmetro
			pagina_retorno = pagina_retorno & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
		else
			'A QueryString ainda não tem nenhum parâmetro
			pagina_retorno = pagina_retorno & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
			end if
		end if
	
	if cliente_selecionado = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	
	dim s_ddd, s_tel, s_nome_cliente, s_mag_end_etg_completo, s_mag_end_cob_completo
	s_nome_cliente = ""
	s_mag_end_etg_completo = ""
	s_mag_end_cob_completo = ""

	dim operacao_origem, c_numero_magento, operationControlTicket, sessionToken, id_magento_api_pedido_xml
	dim c_FlagCadSemiAutoPedMagento_FluxoOtimizado, rb_indicacao, rb_RA, c_indicador
	operacao_origem = Trim(Request("operacao_origem"))
	c_numero_magento = ""
	operationControlTicket = ""
	sessionToken = ""
	id_magento_api_pedido_xml = ""
	c_FlagCadSemiAutoPedMagento_FluxoOtimizado = ""
	rb_indicacao = ""
	rb_RA = ""
	c_indicador = ""
	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		c_numero_magento = Trim(Request("c_numero_magento"))
		operationControlTicket = Trim(Request("operationControlTicket"))
		sessionToken = Trim(Request("sessionToken"))
		id_magento_api_pedido_xml = Trim(Request("id_magento_api_pedido_xml"))
		c_FlagCadSemiAutoPedMagento_FluxoOtimizado = Trim(Request.Form("c_FlagCadSemiAutoPedMagento_FluxoOtimizado"))
		rb_indicacao = Trim(Request.Form("rb_indicacao"))
		rb_RA = Trim(Request.Form("rb_RA"))
		c_indicador = Trim(Request.Form("c_indicador"))

		If Not cria_recordset_otimista(tMAP_XML, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
		If Not cria_recordset_otimista(tMAP_END_ETG, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
		If Not cria_recordset_otimista(tMAP_END_COB, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

		s = "SELECT " & _
				"*" & _
			" FROM t_MAGENTO_API_PEDIDO_XML" & _
			" WHERE" & _
				" (id = " & id_magento_api_pedido_xml & ")"
		if tMAP_XML.State <> 0 then tMAP_XML.Close
		tMAP_XML.open s, cn
		if tMAP_XML.Eof then
			alerta = "Falha ao tentar localizar no banco de dados o registro com os dados do pedido Magento consultados via API (id = " & id_magento_api_pedido_xml & ")"
		else
			s_nome_cliente = UCase(ec_dados_formata_nome(tMAP_XML("customer_firstname"), tMAP_XML("customer_middlename"), tMAP_XML("customer_lastname"), 60))
			c_mag_cpf_cnpj_identificado = retorna_so_digitos(Trim("" & tMAP_XML("cpfCnpjIdentificado")))

			s = "SELECT " & _
					"*" & _
				" FROM t_MAGENTO_API_PEDIDO_XML_DECODE_ENDERECO" & _
				" WHERE" & _
					" (id_magento_api_pedido_xml = " & tMAP_XML("id") & ")" & _
					" AND (tipo_endereco = 'COB')"
			if tMAP_END_COB.State <> 0 then tMAP_END_COB.Close
			tMAP_END_COB.open s, cn

			s = "SELECT " & _
					"*" & _
				" FROM t_MAGENTO_API_PEDIDO_XML_DECODE_ENDERECO" & _
				" WHERE" & _
					" (id_magento_api_pedido_xml = " & tMAP_XML("id") & ")" & _
					" AND (tipo_endereco = 'ETG')"
			if tMAP_END_ETG.State <> 0 then tMAP_END_ETG.Close
			tMAP_END_ETG.open s, cn
			end if
		end if

	dim c_mag_customer_full_name, c_mag_customer_dob, c_mag_customer_email, c_mag_email_identificado, c_mag_cpf_cnpj_identificado
	dim c_mag_end_cob_email, c_mag_end_cob_telephone_ddd, c_mag_end_cob_telephone_numero, c_mag_end_cob_celular_ddd, c_mag_end_cob_celular_numero, c_mag_end_cob_fax_ddd, c_mag_end_cob_fax_numero, c_mag_end_cob_endereco, c_mag_end_cob_endereco_numero, c_mag_end_cob_complemento, c_mag_end_cob_bairro, c_mag_end_cob_cidade, c_mag_end_cob_uf, c_mag_end_cob_cep
	dim c_mag_end_etg_email, c_mag_end_etg_telephone_ddd, c_mag_end_etg_telephone_numero, c_mag_end_etg_celular_ddd, c_mag_end_etg_celular_numero, c_mag_end_etg_fax_ddd, c_mag_end_etg_fax_numero, c_mag_end_etg_endereco, c_mag_end_etg_endereco_numero, c_mag_end_etg_complemento, c_mag_end_etg_bairro, c_mag_end_etg_cidade, c_mag_end_etg_uf, c_mag_end_etg_cep
	c_mag_customer_full_name = ""
	c_mag_customer_dob = ""
	c_mag_customer_email = ""
	c_mag_email_identificado = ""
	c_mag_cpf_cnpj_identificado = ""
	c_mag_end_cob_email = ""
	c_mag_end_cob_telephone_ddd = ""
	c_mag_end_cob_telephone_numero = ""
	c_mag_end_cob_celular_ddd = ""
	c_mag_end_cob_celular_numero = ""
	c_mag_end_cob_fax_ddd = ""
	c_mag_end_cob_fax_numero = ""
	c_mag_end_cob_endereco = ""
	c_mag_end_cob_endereco_numero = ""
	c_mag_end_cob_complemento = ""
	c_mag_end_cob_bairro = ""
	c_mag_end_cob_cidade = ""
	c_mag_end_cob_uf = ""
	c_mag_end_cob_cep = ""
	c_mag_end_etg_email = ""
	c_mag_end_etg_telephone_ddd = ""
	c_mag_end_etg_telephone_numero = ""
	c_mag_end_etg_celular_ddd = ""
	c_mag_end_etg_celular_numero = ""
	c_mag_end_etg_fax_ddd = ""
	c_mag_end_etg_fax_numero = ""
	c_mag_end_etg_endereco = ""
	c_mag_end_etg_endereco_numero = ""
	c_mag_end_etg_complemento = ""
	c_mag_end_etg_bairro = ""
	c_mag_end_etg_cidade = ""
	c_mag_end_etg_uf = ""
	c_mag_end_etg_cep = ""

	if alerta = "" then
		if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
			c_mag_customer_full_name = Ucase(ec_dados_formata_nome(tMAP_XML("customer_firstname"), tMAP_XML("customer_middlename"), tMAP_XML("customer_lastname"), 60))
			if Trim("" & tMAP_XML("customer_dob")) <> "" then
				c_mag_customer_dob = formata_data(converte_para_datetime_from_yyyymmdd_hhmmss(Trim("" & tMAP_XML("customer_dob"))))
				end if
			c_mag_customer_email = ec_dados_filtra_email(Trim("" & tMAP_XML("customer_email")))
			if Not tMAP_END_COB.Eof then
				c_mag_end_cob_email = ec_dados_filtra_email(Trim("" & tMAP_END_COB("email")))
				call ec_dados_decodifica_telefone_formatado(tMAP_END_COB("telephone"), s_ddd, s_tel)
				c_mag_end_cob_telephone_ddd = s_ddd
				c_mag_end_cob_telephone_numero = s_tel
				call ec_dados_decodifica_telefone_formatado(tMAP_END_COB("celular"), s_ddd, s_tel)
				c_mag_end_cob_celular_ddd = s_ddd
				c_mag_end_cob_celular_numero = s_tel
				call ec_dados_decodifica_telefone_formatado(tMAP_END_COB("fax"), s_ddd, s_tel)
				c_mag_end_cob_fax_ddd = s_ddd
				c_mag_end_cob_fax_numero = s_tel
				'NORMALIZA TELEFONES, VERIFICANDO INCLUSIVE REPETIÇÕES
				call ec_dados_normaliza_telefones(c_mag_end_cob_telephone_ddd, c_mag_end_cob_telephone_numero, c_mag_end_cob_celular_ddd, c_mag_end_cob_celular_numero, c_mag_end_cob_fax_ddd, c_mag_end_cob_fax_numero)
				c_mag_end_cob_endereco = Trim("" & tMAP_END_COB("endereco"))
				c_mag_end_cob_endereco_numero = Trim("" & tMAP_END_COB("endereco_numero"))
				c_mag_end_cob_complemento = Trim("" & tMAP_END_COB("endereco_complemento"))
				c_mag_end_cob_bairro = Trim("" & tMAP_END_COB("bairro"))
				c_mag_end_cob_cidade = Trim("" & tMAP_END_COB("cidade"))
				c_mag_end_cob_uf = Trim("" & tMAP_END_COB("uf"))
				c_mag_end_cob_cep = Trim("" & tMAP_END_COB("cep"))
				s_mag_end_cob_completo = formata_endereco(c_mag_end_cob_endereco, c_mag_end_cob_endereco_numero, c_mag_end_cob_complemento, c_mag_end_cob_bairro, c_mag_end_cob_cidade, c_mag_end_cob_uf, c_mag_end_cob_cep)
				end if
			if Not tMAP_END_ETG.Eof then
				c_mag_end_etg_email = ec_dados_filtra_email(Trim("" & tMAP_END_ETG("email")))
				call ec_dados_decodifica_telefone_formatado(tMAP_END_ETG("telephone"), s_ddd, s_tel)
				c_mag_end_etg_telephone_ddd = s_ddd
				c_mag_end_etg_telephone_numero = s_tel
				call ec_dados_decodifica_telefone_formatado(tMAP_END_ETG("celular"), s_ddd, s_tel)
				c_mag_end_etg_celular_ddd = s_ddd
				c_mag_end_etg_celular_numero = s_tel
				call ec_dados_decodifica_telefone_formatado(tMAP_END_ETG("fax"), s_ddd, s_tel)
				c_mag_end_etg_fax_ddd = s_ddd
				c_mag_end_etg_fax_numero = s_tel
				'NORMALIZA TELEFONES, VERIFICANDO INCLUSIVE REPETIÇÕES
				call ec_dados_normaliza_telefones(c_mag_end_etg_telephone_ddd, c_mag_end_etg_telephone_numero, c_mag_end_etg_celular_ddd, c_mag_end_etg_celular_numero, c_mag_end_etg_fax_ddd, c_mag_end_etg_fax_numero)
				c_mag_end_etg_endereco = Trim("" & tMAP_END_ETG("endereco"))
				c_mag_end_etg_endereco_numero = Trim("" & tMAP_END_ETG("endereco_numero"))
				c_mag_end_etg_complemento = Trim("" & tMAP_END_ETG("endereco_complemento"))
				c_mag_end_etg_bairro = Trim("" & tMAP_END_ETG("bairro"))
				c_mag_end_etg_cidade = Trim("" & tMAP_END_ETG("cidade"))
				c_mag_end_etg_uf = Trim("" & tMAP_END_ETG("uf"))
				c_mag_end_etg_cep = Trim("" & tMAP_END_ETG("cep"))
				s_mag_end_etg_completo = formata_endereco(c_mag_end_etg_endereco, c_mag_end_etg_endereco_numero, c_mag_end_etg_complemento, c_mag_end_etg_bairro, c_mag_end_etg_cidade, c_mag_end_etg_uf, c_mag_end_etg_cep)
				end if
			
			c_mag_email_identificado = c_mag_customer_email
			if c_mag_email_identificado = "" then c_mag_email_identificado = c_mag_end_etg_email
			if c_mag_email_identificado = "" then c_mag_email_identificado = c_mag_end_cob_email
			end if
		end if

	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false

'	DADOS DO SÓCIO MAJORITÁRIO
	dim blnCadSocioMaj, blnConsistir, blnConsistirDadosBancarios
	dim strSocMajNome, strSocMajCpf, strSocMajBanco, strSocMajAgencia, strSocMajConta
	dim strSocMajDdd, strSocMajTelefone, strSocMajContato
	if (Not eh_cpf) then blnCadSocioMaj = True else blnCadSocioMaj = False
	if blnCadSocioMaj then
		strSocMajNome = Trim(Request.Form("c_SocioMajNome"))
		strSocMajCpf = Trim(Request.Form("c_SocioMajCpf"))
		strSocMajBanco = Trim(Request.Form("c_SocioMajBanco"))
		strSocMajAgencia = Trim(Request.Form("c_SocioMajAgencia"))
		strSocMajConta = Trim(Request.Form("c_SocioMajConta"))
		strSocMajDdd = Trim(Request.Form("c_SocioMajDdd"))
		strSocMajTelefone = Trim(Request.Form("c_SocioMajTelefone"))
		strSocMajContato = Trim(Request.Form("c_SocioMajContato"))
		end if

'	REF PROFISSIONAL
	dim blnCadRefProfissional
	dim vRefProfissional, intQtdeRefProfissional, intOrdemRefProfissional
	redim vRefProfissional(0)
	set vRefProfissional(Ubound(vRefProfissional)) = New cl_CLIENTE_REF_PROFISSIONAL
	vRefProfissional(Ubound(vRefProfissional)).id_cliente = ""
	if eh_cpf then blnCadRefProfissional=True else blnCadRefProfissional=False
	if blnCadRefProfissional then
		intQtdeRefProfissional = Request.Form("c_RefProfNomeEmpresa").Count
		intOrdemRefProfissional = 0
		for intCounter=1 to intQtdeRefProfissional
			s=Trim(Request.Form("c_RefProfNomeEmpresa")(intCounter))
			if s <> "" then
				if Trim(vRefProfissional(Ubound(vRefProfissional)).nome_empresa) <> "" then
					redim preserve vRefProfissional(Ubound(vRefProfissional)+1)
					set vRefProfissional(Ubound(vRefProfissional)) = New cl_CLIENTE_REF_PROFISSIONAL
					end if
				with vRefProfissional(Ubound(vRefProfissional))
					intOrdemRefProfissional = intOrdemRefProfissional + 1
					.id_cliente				= cliente_selecionado
					.ordem					= intOrdemRefProfissional
					.nome_empresa			= Trim(Request.Form("c_RefProfNomeEmpresa")(intCounter))
					.cargo					= Trim(Request.Form("c_RefProfCargo")(intCounter))
					.ddd					= Trim(Request.Form("c_RefProfDdd")(intCounter))
					.telefone				= retorna_so_digitos(Trim(Request.Form("c_RefProfTelefone")(intCounter)))
					s = Trim(Request.Form("c_RefProfPeriodoRegistro")(intCounter))
					if s = "" then 
						.periodo_registro	= Null 
					else 
						.periodo_registro	= StrToDate(s)
						end if
					.rendimentos			= converte_numero(Trim(Request.Form("c_RefProfRendimentos")(intCounter)))
					.cnpj					= retorna_so_digitos(Trim(Request.Form("c_RefProfCnpj")(intCounter)))
					end with
				end if
			next
		end if

'	REF COMERCIAL
	dim blnCadRefComercial
	dim vRefComercial, intQtdeRefComercial, intOrdemRefComercial
	redim vRefComercial(0)
	set vRefComercial(Ubound(vRefComercial)) = New cl_CLIENTE_REF_COMERCIAL
	vRefComercial(Ubound(vRefComercial)).id_cliente = ""
	if (Not eh_cpf) then blnCadRefComercial=True else blnCadRefComercial=False
	if blnCadRefComercial then
		intQtdeRefComercial = Request.Form("c_RefComercialNomeEmpresa").Count
		intOrdemRefComercial = 0
		for intCounter=1 to intQtdeRefComercial
			s=Trim(Request.Form("c_RefComercialNomeEmpresa")(intCounter))
			if s <> "" then
				if Trim(vRefComercial(Ubound(vRefComercial)).nome_empresa) <> "" then
					redim preserve vRefComercial(Ubound(vRefComercial)+1)
					set vRefComercial(Ubound(vRefComercial)) = New cl_CLIENTE_REF_COMERCIAL
					end if
				with vRefComercial(Ubound(vRefComercial))
					intOrdemRefComercial    = intOrdemRefComercial + 1
					.id_cliente				= cliente_selecionado
					.ordem					= intOrdemRefComercial
					.nome_empresa			= Trim(Request.Form("c_RefComercialNomeEmpresa")(intCounter))
					.contato				= Trim(Request.Form("c_RefComercialContato")(intCounter))
					.ddd					= Trim(Request.Form("c_RefComercialDdd")(intCounter))
					.telefone				= retorna_so_digitos(Trim(Request.Form("c_RefComercialTelefone")(intCounter)))
					end with
				end if
			next
		end if
	
'	REF BANCÁRIA
	dim blnCadRefBancaria
	dim vRefBancaria, intQtdeRefBancaria, intOrdemRefBancaria
	redim vRefBancaria(0)
	set vRefBancaria(Ubound(vRefBancaria)) = New cl_CLIENTE_REF_BANCARIA
	vRefBancaria(Ubound(vRefBancaria)).id_cliente = ""
	blnCadRefBancaria=True
	if blnCadRefBancaria then
		intQtdeRefBancaria = Request.Form("c_RefBancariaBanco").Count
		intOrdemRefBancaria = 0
		for intCounter=1 to intQtdeRefBancaria
			s=Trim(Request.Form("c_RefBancariaBanco")(intCounter))
			if s <> "" then
				if Trim(vRefBancaria(Ubound(vRefBancaria)).id_cliente) <> "" then
					redim preserve vRefBancaria(Ubound(vRefBancaria)+1)
					set vRefBancaria(Ubound(vRefBancaria)) = New cl_CLIENTE_REF_BANCARIA
					end if
				with vRefBancaria(Ubound(vRefBancaria))
					intOrdemRefBancaria     = intOrdemRefBancaria + 1
					.id_cliente				= cliente_selecionado
					.ordem					= intOrdemRefBancaria
					.banco					= Trim(Request.Form("c_RefBancariaBanco")(intCounter))
					.agencia				= Trim(Request.Form("c_RefBancariaAgencia")(intCounter))
					.conta					= Trim(Request.Form("c_RefBancariaConta")(intCounter))
					.ddd					= Trim(Request.Form("c_RefBancariaDdd")(intCounter))
					.telefone				= retorna_so_digitos(Trim(Request.Form("c_RefBancariaTelefone")(intCounter)))
					.contato				= Trim(Request.Form("c_RefBancariaContato")(intCounter))
					end with
				end if
			next
		end if
	
	alerta = ""
	if cnpj_cpf_selecionado = "" then
		alerta="CNPJ/CPF NÃO FORNECIDO."
	elseif Not cnpj_cpf_ok(cnpj_cpf_selecionado) then
		alerta="CNPJ/CPF INVÁLIDO."
	elseif s_nome = "" then
		if eh_cpf then
			alerta="PREENCHA O NOME DO CLIENTE."
		else
			alerta="PREENCHA A RAZÃO SOCIAL DO CLIENTE."
			end if
	elseif s_endereco = "" then
		alerta="PREENCHA O ENDEREÇO."
	elseif Len(s_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
		alerta="ENDEREÇO EXCEDE O TAMANHO MÁXIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(s_endereco)) & " CARACTERES<br>TAMANHO MÁXIMO: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " CARACTERES"
	elseif s_endereco_numero = "" then
		alerta="PREENCHA O NÚMERO DO ENDEREÇO."
	elseif s_bairro = "" then
		alerta="PREENCHA O BAIRRO."
	elseif s_cidade = "" then
		alerta="PREENCHA A CIDADE."
	elseif (s_uf="") Or (Not uf_ok(s_uf)) then
		alerta="UF INVÁLIDA."
	elseif s_cep = "" then
		alerta="INFORME O CEP."
	elseif Not cep_ok(s_cep) then
		alerta="CEP INVÁLIDO."
	elseif Not ddd_ok(s_ddd_res) then
		alerta="DDD INVÁLIDO."
	elseif Not telefone_ok(s_tel_res) then
		alerta="TELEFONE RESIDENCIAL INVÁLIDO."
	elseif (s_ddd_res <> "") And ((s_tel_res = "")) then
		alerta="PREENCHA O TELEFONE RESIDENCIAL."
	elseif (s_ddd_res = "") And ((s_tel_res <> "")) then
		alerta="PREENCHA O DDD."
	elseif Not ddd_ok(s_ddd_com) then
		alerta="DDD INVÁLIDO."
	elseif Not telefone_ok(s_tel_com) then
		alerta="TELEFONE COMERCIAL INVÁLIDO."
	elseif (s_ddd_com <> "") And ((s_tel_com = "")) then
		alerta="PREENCHA O TELEFONE COMERCIAL."
	elseif (s_ddd_com = "") And ((s_tel_com <> "")) then
		alerta="PREENCHA O DDD."
    elseif (s_ddd_cel = "") And ((s_tel_cel <> "")) then
		alerta="PREENCHA O DDD."
    elseif Not ddd_ok(s_ddd_cel) then
        alerta="DDD DO CELULAR INVÁLIDO."
    elseif Len(retorna_so_digitos(s_tel_cel)) > 9 then
        alerta="NÚMERO DO CELULAR INVÁLIDO."
	elseif eh_cpf And (s_tel_res="") And (s_tel_com="") And (s_tel_cel="") then
		alerta="PREENCHA PELO MENOS UM TELEFONE."
	elseif (Not eh_cpf) And (s_tel_com="") And (s_tel_com_2="") then
		alerta="PREENCHA O TELEFONE."
	elseif (s_ie="") And (s_contribuinte_icms = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
		alerta="PREENCHA A INSCRIÇÃO ESTADUAL."
'	elseif s_midia="" then
'		alerta="INDIQUE A FORMA PELA QUAL CONHECEU A DIS." 
	elseif (s_contribuinte_icms = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) And (s_ie="") then
		alerta="SE CLIENTE É CONTRIBUINTE DO ICMS A INSCRIÇÃO ESTADUAL DEVE SER PREENCHIDA."
		end if

	if alerta = "" then
		if False then
			if eh_cpf And (Not sexo_ok(s_sexo)) then
				alerta="INDIQUE QUAL O SEXO."
				end if
			end if
		end if

	if alerta = "" then
		if (s_produtor_rural = COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) Then
			if (s_contribuinte_icms <> COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) Or (s_ie = "") then
				alerta = "Para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE"
				end if
			end if
		end if
	
'	CONSISTÊNCIAS P/ EMISSÃO DE NFe
	dim s_tabela_municipios_IBGE
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
	'	I.E. É VÁLIDA?
		if ( (s_cliente_tipo = ID_PF) And (Cstr(s_produtor_rural) = Cstr(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM)) ) _
			Or _
			( (s_cliente_tipo = ID_PJ) And (Cstr(s_contribuinte_icms) = Cstr(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM)) ) _
			Or _
			( (s_cliente_tipo = ID_PJ) And (Cstr(s_contribuinte_icms) = Cstr(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO)) And (s_ie <> "") ) then
			if Not isInscricaoEstadualValida(s_ie, s_uf) then
				alerta="Preencha a IE (Inscrição Estadual) com um número válido!!" & _
						"<br>" & "Certifique-se de que a UF informada corresponde à UF responsável pelo registro da IE."
				end if
			end if
	
	'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
		dim s_lista_sugerida_municipios
		dim v_lista_sugerida_municipios
		dim iCounterLista, iNumeracaoLista
		if Not consiste_municipio_IBGE_ok(s_cidade, s_uf, s_lista_sugerida_municipios, msg_erro) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			if msg_erro <> "" then
				alerta = alerta & msg_erro
			else
				alerta = alerta & "Município '" & s_cidade & "' não consta na relação de municípios do IBGE para a UF de '" & s_uf & "'!!"
				if s_lista_sugerida_municipios <> "" then
					alerta = alerta & "<br>" & _
									  "Localize o município na lista abaixo e verifique se a grafia está correta!!"
					v_lista_sugerida_municipios = Split(s_lista_sugerida_municipios, chr(13))
					iNumeracaoLista=0
					for iCounterLista=LBound(v_lista_sugerida_municipios) to UBound(v_lista_sugerida_municipios)
						if Trim("" & v_lista_sugerida_municipios(iCounterLista)) <> "" then
							iNumeracaoLista=iNumeracaoLista+1
							s_tabela_municipios_IBGE = s_tabela_municipios_IBGE & _
												"	<tr>" & chr(13) & _
												"		<td align='right'>" & chr(13) & _
												"			<span class='N'>&nbsp;" & Cstr(iNumeracaoLista) & "." & "</span>" & chr(13) & _
												"		</td>" & chr(13) & _
												"		<td align='left'>" & chr(13) & _
												"			<span class='N'>" & Trim("" & v_lista_sugerida_municipios(iCounterLista)) & "</span>" & chr(13) & _
												"		</td>" & chr(13) & _
												"	</tr>" & chr(13)
							end if
						next

					if s_tabela_municipios_IBGE <> "" then
						s_tabela_municipios_IBGE = _
								"<table cellspacing='0' cellpadding='1'>" & chr(13) & _
								"	<tr>" & chr(13) & _
								"		<td align='center'>" & chr(13) & _
								"			<p class='N'>" & "Relação de municípios de '" & s_uf & "' que se iniciam com a letra '" & Ucase(left(s_cidade,1)) & "'" & "</p>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13) & _
								"	<tr>" & chr(13) & _
								"		<td align='center'>" & chr(13) &_
								"			<table cellspacing='0' border='1'>" & chr(13) & _
												s_tabela_municipios_IBGE & _
								"			</table>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13) & _
								"</table>" & chr(13)
						end if
					end if
				end if
			end if
		end if
	
	dim s_caracteres_invalidos
	if alerta = "" then
		if Not isTextoValido(s_nome, s_caracteres_invalidos) then
			alerta="O CAMPO 'NOME' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(s_endereco, s_caracteres_invalidos) then
			alerta="O CAMPO 'ENDEREÇO' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(s_endereco_numero, s_caracteres_invalidos) then
			alerta="O CAMPO NÚMERO DO ENDEREÇO POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(s_endereco_complemento, s_caracteres_invalidos) then
			alerta="O CAMPO 'COMPLEMENTO' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(s_bairro, s_caracteres_invalidos) then
			alerta="O CAMPO 'BAIRRO' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(s_cidade, s_caracteres_invalidos) then
			alerta="O CAMPO 'CIDADE' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(s_contato, s_caracteres_invalidos) then
			alerta="O CAMPO 'CONTATO' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(s_filiacao, s_caracteres_invalidos) then
			alerta="O CAMPO 'FILIAÇÃO' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(s_obs_crediticias, s_caracteres_invalidos) then
			alerta="O CAMPO 'OBSERVAÇÕES CREDITÍCIAS' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
			end if
		end if
	
	if (alerta="") And (s_dt_nasc<>"") then
		if (DateDiff("m", StrToDate(s_dt_nasc), Date)/12) < 10 then alerta = "DATA DE NASCIMENTO É INVÁLIDA."
		end if
	
	if alerta = "" then
	'	REF BANCÁRIA
		for intCounter=Lbound(vRefBancaria) to Ubound(vRefBancaria)
			if vRefBancaria(intCounter).id_cliente <> "" then
				with vRefBancaria(intCounter)
					if Trim(.banco) = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Ref Bancária (" & CStr(.ordem) & "): informe o banco."
						end if
					if Trim(.agencia) = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Ref Bancária (" & CStr(.ordem) & "): informe a agência."
						end if
					if Trim(.conta) = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Ref Bancária (" & CStr(.ordem) & "): informe o número da conta."
						end if
					end with
				end if
			next
		end if

	if alerta = "" then
	'	REF PROFISSIONAL
		for intCounter=Lbound(vRefProfissional) to Ubound(vRefProfissional)
			if vRefProfissional(intCounter).id_cliente <> "" then
				with vRefProfissional(intCounter)
					if Trim(.nome_empresa) = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Ref Profissional (" & CStr(.ordem) & "): informe o nome da empresa."
						end if
					if Trim(.cargo) = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Ref Profissional (" & CStr(.ordem) & "): informe o cargo."
						end if
					end with
				end if
			next
		end if

	if alerta = "" then
	'	REF COMERCIAL
		for intCounter=Lbound(vRefComercial) to Ubound(vRefComercial)
			if vRefComercial(intCounter).id_cliente <> "" then
				with vRefComercial(intCounter)
					if Trim(.nome_empresa) = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Ref Comercial (" & CStr(.ordem) & "): informe o nome da empresa."
						end if
					end with
				end if
			next
		end if

	if alerta = "" then
	'	DADOS DO SÓCIO MAJORITÁRIO
		if blnCadSocioMaj then
			blnConsistir = False
			blnConsistirDadosBancarios = False
			if strSocMajNome <> "" then blnConsistir = True
			if strSocMajCpf <> "" then blnConsistir = True
			if strSocMajBanco <> "" then 
				blnConsistir = True
				blnConsistirDadosBancarios = True
				end if
			if strSocMajAgencia <> "" then 
				blnConsistir = True
				blnConsistirDadosBancarios = True
				end if
			if strSocMajConta <> "" then 
				blnConsistir = True
				blnConsistirDadosBancarios = True
				end if
			if strSocMajDdd <> "" then blnConsistir = True
			if strSocMajTelefone <> "" then blnConsistir = True
			if strSocMajContato <> "" then blnConsistir = True
			
			if blnConsistir then
				if strSocMajNome = "" then 
					alerta=texto_add_br(alerta)
					alerta=alerta & "Informe o nome do sócio majoritário."
					end if
				end if
			if blnConsistirDadosBancarios then
				if strSocMajBanco = "" then 
					alerta=texto_add_br(alerta)
					alerta=alerta & "Informe o banco nos dados bancários do sócio majoritário."
					end if
				if strSocMajAgencia = "" then 
					alerta=texto_add_br(alerta)
					alerta=alerta & "Informe a agência nos dados bancários do sócio majoritário."
					end if
				if strSocMajConta = "" then 
					alerta=texto_add_br(alerta)
					alerta=alerta & "Informe o número da conta nos dados bancários do sócio majoritário."
					end if
				end if
			end if
		end if
	
	if alerta = "" then
		if operacao_selecionada <> OP_INCLUI then
			if Not operacao_permitida(OP_LJA_EDITA_CLIENTE_DADOS_CADASTRAIS, s_lista_operacoes_permitidas) then
				alerta = "Nível de acesso insuficiente para realizar esta operação."
				end if
			end if
		end if
	
	dim s_cnpj_cpf
	dim r_cliente
    dim blnVerificarTel
	if alerta = "" then
		if operacao_selecionada = OP_INCLUI then
			s_cnpj_cpf = cnpj_cpf_selecionado
		else
			set r_cliente = New cl_CLIENTE
			call x_cliente_bd(cliente_selecionado, r_cliente)
			s_cnpj_cpf = r_cliente.cnpj_cpf
			end if

    ' VERIFICA A DISPONIBILIDADE DO USO DO TELEFONE NO CADASTRO
        blnVerificarTel = False
        if operacao_selecionada = OP_INCLUI then
			blnVerificarTel = True
		else
			if s_tel_res <> "" And (s_ddd_res<>r_cliente.ddd_res Or s_tel_res<>r_cliente.tel_res) then blnVerificarTel = true
			end if
		
		if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then blnVerificarTel = False

        if blnVerificarTel then
            if s_tel_res <> "" then
                if (Cstr(s_ddd_res & s_tel_res) = TEL_BONSHOP_1) Or (Cstr(s_ddd_res & s_tel_res) = TEL_BONSHOP_2) Or (Cstr(s_ddd_res & s_tel_res) = TEL_BONSHOP_3) then
                    alerta="NÃO É PERMITIDO UTILIZAR TELEFONES DA BONSHOP NO CADASTRO DE CLIENTES."
                elseif verifica_telefones_repetidos(s_ddd_res, s_tel_res, s_cnpj_cpf) > NUM_MAXIMO_TELEFONES_REPETIDOS_CAD_CLIENTES then
                    alerta="TELEFONE RESIDENCIAL (" & s_ddd_res & ") " & s_tel_res & " JÁ ESTÁ SENDO UTILIZADO NO CADASTRO DE OUTROS CLIENTES. <br>Não foi possível concluir o cadastro."
                end if
            end if
        end if

        blnVerificarTel = False
		if operacao_selecionada = OP_INCLUI then
			blnVerificarTel = True
		else
			if s_tel_com <> "" And (s_ddd_com<>r_cliente.ddd_com Or s_tel_com<>r_cliente.tel_com) then blnVerificarTel = true
			end if

		if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then blnVerificarTel = False

        if blnVerificarTel then
            if s_tel_com <> "" then
                if (Cstr(s_ddd_com & s_tel_com) = TEL_BONSHOP_1) Or (Cstr(s_ddd_com & s_tel_com) = TEL_BONSHOP_2) Or (Cstr(s_ddd_com & s_tel_com) = TEL_BONSHOP_3) then
                    alerta="NÃO É PERMITIDO UTILIZAR TELEFONES DA BONSHOP NO CADASTRO DE CLIENTES."
                elseif verifica_telefones_repetidos(s_ddd_com, s_tel_com, s_cnpj_cpf) > NUM_MAXIMO_TELEFONES_REPETIDOS_CAD_CLIENTES then
                    alerta="TELEFONE COMERCIAL (" & s_ddd_com & ") " & s_tel_com & " JÁ ESTÁ SENDO UTILIZADO NO CADASTRO DE OUTROS CLIENTES. <br>Não foi possível concluir o cadastro."
                end if
            end if
        end if
        
        blnVerificarTel = False
		if operacao_selecionada = OP_INCLUI then
			blnVerificarTel = True
		else
			if s_tel_com_2 <> "" And (s_ddd_com_2<>r_cliente.ddd_com_2 Or s_tel_com_2<>r_cliente.tel_com_2) then blnVerificarTel = true
			end if

		if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then blnVerificarTel = False

        if blnVerificarTel then
            if s_tel_com_2 <> "" then
                if (Cstr(s_ddd_com_2 & s_tel_com_2) = TEL_BONSHOP_1) Or (Cstr(s_ddd_com_2 & s_tel_com_2) = TEL_BONSHOP_2) Or (Cstr(s_ddd_com_2 & s_tel_com_2) = TEL_BONSHOP_3) then
                    alerta="NÃO É PERMITIDO UTILIZAR TELEFONES DA BONSHOP NO CADASTRO DE CLIENTES."
                elseif verifica_telefones_repetidos(s_ddd_com_2, s_tel_com_2, s_cnpj_cpf) > NUM_MAXIMO_TELEFONES_REPETIDOS_CAD_CLIENTES then
                    alerta="TELEFONE COMERCIAL (" & s_ddd_com_2 & ") " & s_tel_com_2 & " JÁ ESTÁ SENDO UTILIZADO NO CADASTRO DE OUTROS CLIENTES. <br>Não foi possível concluir o cadastro."
                end if
            end if
        end if

        blnVerificarTel = False
		if operacao_selecionada = OP_INCLUI then
			blnVerificarTel = True
		else
			if s_tel_cel <> "" And (s_ddd_cel<>r_cliente.ddd_cel Or s_tel_cel<>r_cliente.tel_cel) then blnVerificarTel = true
			end if

		if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then blnVerificarTel = False

        if blnVerificarTel then
            if s_tel_cel <> "" then
                if (Cstr(s_ddd_cel & s_tel_cel) = TEL_BONSHOP_1) Or (Cstr(s_ddd_cel & s_tel_cel) = TEL_BONSHOP_2) Or (Cstr(s_ddd_cel & s_tel_cel) = TEL_BONSHOP_3) then
                    alerta="NÃO É PERMITIDO UTILIZAR TELEFONES DA BONSHOP NO CADASTRO DE CLIENTES."
                elseif verifica_telefones_repetidos(s_ddd_cel, s_tel_cel, s_cnpj_cpf) > NUM_MAXIMO_TELEFONES_REPETIDOS_CAD_CLIENTES then
                    alerta="TELEFONE CELULAR (" & s_ddd_cel & ") " & s_tel_cel & " JÁ ESTÁ SENDO UTILIZADO NO CADASTRO DE OUTROS CLIENTES. <br>Não foi possível concluir o cadastro."
                end if
            end if
        end if
        
		
		if s_email <> "" then
		'	CONSISTÊNCIA DESATIVADA TEMPORARIAMENTE
'			if Not email_AF_ok(s_email, s_cnpj_cpf, msg_erro_aux) then
'				alerta=texto_add_br(alerta)
'				alerta=alerta & "Endereço de email (" & s_email & ") não é válido!!<br />" & msg_erro_aux
'				end if
			end if
		end if
	
	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	if alerta = "" then
		if operacao_selecionada = OP_INCLUI then
		'	VERIFICA SE O CNPJ/CPF JÁ ESTÁ CADASTRADO (ESTE É UM TRATAMENTO P/ O ARCLUBE, EM QUE ATENDENTES DIFERENTES PODEM ESTAR CADASTRANDO SIMULTANEAMENTE O MESMO CLIENTE QUANDO
		'	ESTE REALIZA VÁRIOS PEDIDOS NO SITE).
			s = "SELECT id, cnpj_cpf, tipo, nome FROM t_CLIENTE WHERE (cnpj_cpf = '" & cnpj_cpf_selecionado & "')"
			if r.State <> 0 then r.Close
			r.Open s, cn
			if Not r.Eof then
				if Trim("" & r("tipo")) = ID_PF then
					alerta = "CPF " & cnpj_cpf_formata(Trim("" & r("cnpj_cpf"))) & " já está cadastrado!"
				else
					alerta = "CNPJ " & cnpj_cpf_formata(Trim("" & r("cnpj_cpf"))) & " já está cadastrado!"
					end if
				alerta=texto_add_br(alerta)
				alerta=alerta & "Clique <u>aqui</u> para consultar o cadastro do cliente"
				alerta = "<a href='ClienteEdita.asp?cliente_selecionado=" & Trim("" & r("id")) & "&cnpj_cpf_selecionado=" & Trim("" & r("cnpj_cpf")) & "&operacao_selecionada=" & OP_CONSULTA & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "'>" & _
						 "<span style='color:#ffff00;'>" & _
						 alerta & _
						 "</span>" & _
						 "</a>"
				end if
			if r.State <> 0 then r.Close
			end if
		end if

	'Verifica se está havendo edição no cadastro de cliente que possui pedido com status de análise de crédito 'crédito ok' e com entrega pendente
    'somente se st_memorizacao_completa_enderecos = 0; se != 0, o endereço é controlado em cada pedido separadamente
	dim blnHaPedidoAprovadoComEntregaPendente, listaPedidoAprovadoComEntregaPendente
	blnHaPedidoAprovadoComEntregaPendente = False
	listaPedidoAprovadoComEntregaPendente = ""
	if alerta = "" then
		if operacao_selecionada <> OP_EXCLUI then
			s = "SELECT" & _
					" tP.pedido" & _
				" FROM t_PEDIDO tP" & _
					" INNER JOIN t_PEDIDO tP__BASE ON (tP.pedido_base = tP__BASE.pedido)" & _
				" WHERE" & _
					" (tP.id_cliente = '" & cliente_selecionado & "')" & _
					" AND (tP.loja NOT IN ('" & NUMERO_LOJA_ECOMMERCE_AR_CLUBE & "', '" & NUMERO_LOJA_TRANSFERENCIA & "', '" & NUMERO_LOJA_KITS & "'))" & _
					" AND (tP__BASE.analise_credito = " & CStr(COD_AN_CREDITO_OK) & ")" & _
					" AND (tP.st_entrega NOT IN ('" & ST_ENTREGA_ENTREGUE & "', '" & ST_ENTREGA_CANCELADO & "'))" & _
					" AND (tP.st_memorizacao_completa_enderecos = 0)" & _
				" ORDER BY" & _
					" tP.data_hora"
			if r.State <> 0 then r.Close
			r.Open s, cn
			do while Not r.Eof
				blnHaPedidoAprovadoComEntregaPendente = True
				if listaPedidoAprovadoComEntregaPendente <> "" then listaPedidoAprovadoComEntregaPendente = listaPedidoAprovadoComEntregaPendente & ", "
				listaPedidoAprovadoComEntregaPendente = listaPedidoAprovadoComEntregaPendente & Trim("" & r("pedido"))
				r.MoveNext
				loop
			if r.State <> 0 then r.Close
			end if
		end if

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			s="SELECT COUNT(*) AS qtde FROM t_PEDIDO WHERE (id_cliente = '" & cliente_selecionado & "')"
			r.Open s, cn
		'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
			if Cstr(r("qtde")) > Cstr(0) then
				erro_fatal=True
				alerta = "CLIENTE NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE PEDIDOS."
				end if
			r.Close 

			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_ORCAMENTO WHERE (id_cliente = '" & cliente_selecionado & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "CLIENTE NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE ORÇAMENTOS."
					end if
				r.Close 
				end if
			
			if Not erro_fatal then
			'	INFO P/ LOG
				s="SELECT * FROM t_CLIENTE WHERE id = '" & cliente_selecionado & "'"
				r.Open s, cn
				if Not r.EOF then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					s_log = log_via_vetor_monta_exclusao(vLog1)
					end if
				r.Close
				end if
			
			if Not erro_fatal then
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
							" id_nsu = '" & ID_XLOCK_SYNC_CLIENTE & "'"
					cn.Execute(s)
					end if

				if Not erro_fatal then
					s="DELETE FROM t_CLIENTE_REF_BANCARIA WHERE id_cliente = '" & cliente_selecionado & "'"
					cn.Execute(s)
					If Err <> 0 then 
						erro_fatal=True
						alerta = "FALHA AO REMOVER A REF BANCÁRIA DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if

				if Not erro_fatal then
					s="DELETE FROM t_CLIENTE_REF_PROFISSIONAL WHERE id_cliente = '" & cliente_selecionado & "'"
					cn.Execute(s)
					If Err <> 0 then 
						erro_fatal=True
						alerta = "FALHA AO REMOVER A REF PROFISSIONAL DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if

				if Not erro_fatal then
					s="DELETE FROM t_CLIENTE_REF_COMERCIAL WHERE id_cliente = '" & cliente_selecionado & "'"
					cn.Execute(s)
					If Err <> 0 then 
						erro_fatal=True
						alerta = "FALHA AO REMOVER A REF COMERCIAL DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if Not erro_fatal then
					s="DELETE FROM t_CLIENTE WHERE id = '" & cliente_selecionado & "'"
					cn.Execute(s)
					If Err = 0 then 
						if s_log <> "" then grava_log usuario, loja, "", cliente_selecionado, OP_LOG_CLIENTE_EXCLUSAO, s_log
					else
						erro_fatal=True
						alerta = "FALHA AO REMOVER O CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if Not erro_fatal then
				'	~~~~~~~~~~~~~~
					cn.CommitTrans
				'	~~~~~~~~~~~~~~
				else
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
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
							" id_nsu = '" & ID_XLOCK_SYNC_CLIENTE & "'"
					cn.Execute(s)
					end if

				s = "SELECT * FROM t_CLIENTE WHERE id = '" & cliente_selecionado & "'"
				r.Open s, cn
				if r.EOF then 
					r.AddNew 
					criou_novo_reg_cliente = True
					r("id")=cliente_selecionado
					r("dt_cadastro") = Date
					r("usuario_cadastro") = usuario
					r("sistema_responsavel_cadastro") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
				else
					criou_novo_reg_cliente = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					end if
				
				s_cep_original = Trim("" & r("cep"))
				s_cep_novo = s_cep
				
				r("cnpj_cpf")=cnpj_cpf_selecionado
				r("tipo")=s_cliente_tipo
				r("ie")=s_ie
				r("rg")=s_rg
				s_nome_original = Trim("" & r("nome"))
				r("nome")=s_nome
				r("sexo")=s_sexo
				If (Trim(s_contribuinte_icms) <> "") And (s_contribuinte_icms <> s_contribuinte_icms_cadastrado) Then
					r("contribuinte_icms_status")=CInt(s_contribuinte_icms)
					r("contribuinte_icms_data")=Now
					r("contribuinte_icms_data_hora")=Now
					r("contribuinte_icms_usuario")=usuario
					End If
				If (eh_cpf) And (Trim(s_produtor_rural) <> "") And (s_produtor_rural <> s_produtor_rural_cadastrado) Then
					r("produtor_rural_status")=CInt(s_produtor_rural)
					r("produtor_rural_data")=Now
					r("produtor_rural_data_hora")=Now
					r("produtor_rural_usuario")=usuario
					End If
				r("endereco")=s_endereco
				r("endereco_numero")=s_endereco_numero
				r("endereco_complemento")=s_endereco_complemento
				r("bairro")=s_bairro
				r("cidade")=s_cidade
				r("cep")=s_cep
				r("uf")=s_uf
				r("ddd_res")=s_ddd_res
				r("tel_res")=s_tel_res
				r("ddd_com")=s_ddd_com
				r("tel_com")=s_tel_com
				r("ramal_com")=s_ramal_com
				r("contato")=s_contato
				r("ddd_cel")=s_ddd_cel
				r("tel_cel")=s_tel_cel
				r("ddd_com_2")=s_ddd_com_2
				r("tel_com_2")=s_tel_com_2
				r("ramal_com_2")=s_ramal_com_2
				if s_dt_nasc<>"" then
					r("dt_nasc")=StrToDate(s_dt_nasc)
				else
					r("dt_nasc")=Null
					end if
				r("filiacao")=s_filiacao
				r("obs_crediticias")=s_obs_crediticias
				r("midia")=s_midia
				r("email")=s_email
				r("email_xml")=s_email_xml
				r("email_boleto")=s_email_boleto
				r("dt_ult_atualizacao")=Now
				r("usuario_ult_atualizacao")=usuario

				if strCampoIndicadorEditavel = "S" then
					r("indicador")=s_indicador
					end if
				
				if blnCadSocioMaj then
					r("SocMaj_Nome")=strSocMajNome
					r("SocMaj_CPF")=retorna_so_digitos(strSocMajCpf)
					r("SocMaj_banco")=strSocMajBanco
					r("SocMaj_agencia")=strSocMajAgencia
					r("SocMaj_conta")=strSocMajConta
					r("SocMaj_ddd")=strSocMajDdd
					r("SocMaj_telefone")=retorna_so_digitos(strSocMajTelefone)
					r("SocMaj_contato")=strSocMajContato
					end if
				
				r("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP

				r.Update

				If Err = 0 then 
				'	PREPARA O LOG
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg_cliente then
						s_log = log_via_vetor_monta_inclusao(vLog2)
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if s_log <> "" then 
							s_log="id=" & Trim("" & r("id")) & "; " & s_log
							end if
						end if
				else
					erro_fatal=True
					alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				
				r.Close
				set r = nothing
				if Not cria_recordset_otimista(r, msg_erro) then 
					erro_fatal=True
					alerta = "FALHA AO CRIAR RECORDSET"
					end if
				
			'	REF BANCÁRIA
				if blnCadRefBancaria then
					if Not erro_fatal then
						s="UPDATE t_CLIENTE_REF_BANCARIA SET excluido_status=1 WHERE (id_cliente = '" & cliente_selecionado & "')"
						cn.Execute(s)
						If Err <> 0 then 
							erro_fatal=True
							alerta = "FALHA AO PREPARAR ALTERAÇÃO DOS DADOS DE REF BANCÁRIA DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						
						if Not erro_fatal then
							for intCounter=Lbound(vRefBancaria) to Ubound(vRefBancaria)
								with vRefBancaria(intCounter)
									if Trim(.id_cliente) <> "" then
										s = "SELECT " & _
												"*" & _
											" FROM t_CLIENTE_REF_BANCARIA" & _
											" WHERE" & _
												" (id_cliente = '" & Trim(.id_cliente) & "')" & _
												" AND (banco = '" & Trim("" & .banco) & "')" & _
												" AND (agencia = '" & Trim("" & .agencia) & "')" & _
												" AND (conta = '" & Trim("" & .conta) & "')"
										if r.State <> 0 then r.Close
										r.Open s, cn
										if r.EOF then 
											r.AddNew 
											criou_novo_reg_aux = True
										'	CAMPOS DA CHAVE PRIMÁRIA
											r("id_cliente") = Trim(.id_cliente)
											r("banco") = Trim(.banco)
											r("agencia") = Trim(.agencia)
											r("conta") = Trim(.conta)
											r("dt_cadastro") = Date
											r("usuario_cadastro") = usuario
										else
											criou_novo_reg_aux = False
											log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir_ref_bancaria
											end if
							
										r("ordem") = .ordem
										r("ddd") = Trim(.ddd)
										r("telefone") = Trim(.telefone)
										r("contato") = Trim(.contato)
										r("excluido_status") = 0
										
										r.Update
										
										If Err = 0 then 
										'	PREPARA O LOG
											log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir_ref_bancaria
											if criou_novo_reg_aux then
												s_log_aux = "Ref Bancária incluída: " & log_via_vetor_monta_inclusao(vLog2)
											else
												s_log_aux = log_via_vetor_monta_alteracao(vLog1, vLog2)
												if s_log_aux <> "" then 
													s_log_aux="Ref Bancária alterada (banco: " & Trim(.banco) & ", ag: " & Trim(.agencia) & ", conta: " & Trim(.conta) & "): " & s_log_aux
													end if
												end if
												
											if s_log_aux <> "" then
												if s_log <> "" then s_log = s_log & "; "
												s_log = s_log & s_log_aux
												end if
										else
											erro_fatal=True
											alerta = "FALHA AO GRAVAR OS DADOS DA REF BANCÁRIA (" & Cstr(Err) & ": " & Err.Description & ")."
											end if
				
										r.Close
										set r = nothing
										if Not cria_recordset_otimista(r, msg_erro) then 
											erro_fatal=True
											alerta = "FALHA AO CRIAR RECORDSET"
											end if
										end if
									end with
								next
							
							if Not erro_fatal then
							'	DADOS P/ O LOG
								s_log_aux=""
								s="SELECT * FROM t_CLIENTE_REF_BANCARIA WHERE (id_cliente = '" & cliente_selecionado & "') AND (excluido_status=1) ORDER BY ordem"
								set r = cn.Execute(s)
								do while Not r.EOF
									log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir_ref_bancaria
									if s_log_aux <> "" then s_log_aux = s_log_aux & "; "
									s_log_aux = s_log_aux & "Ref Bancária excluída: " & log_via_vetor_monta_exclusao(vLog1)
									r.MoveNext
									loop
								
								r.Close
								set r = nothing
								if Not cria_recordset_otimista(r, msg_erro) then 
									erro_fatal=True
									alerta = "FALHA AO CRIAR RECORDSET"
									end if
								
								if s_log_aux <> "" then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & s_log_aux
									end if
								
								s="DELETE FROM t_CLIENTE_REF_BANCARIA WHERE (id_cliente = '" & cliente_selecionado & "') AND (excluido_status=1)"
								cn.Execute(s)
								If Err <> 0 then 
									erro_fatal=True
									alerta = "FALHA AO ALTERAR DADOS DE REF BANCÁRIA DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
									end if
								end if
							end if
						end if
					end if '(if blnCadRefBancaria)

			'	REF PROFISSIONAL
				if blnCadRefProfissional then
					if Not erro_fatal then
						s="UPDATE t_CLIENTE_REF_PROFISSIONAL SET excluido_status=1 WHERE (id_cliente = '" & cliente_selecionado & "')"
						cn.Execute(s)
						If Err <> 0 then 
							erro_fatal=True
							alerta = "FALHA AO PREPARAR ALTERAÇÃO DOS DADOS DE REF PROFISSIONAL DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						
						if Not erro_fatal then
							for intCounter=Lbound(vRefProfissional) to Ubound(vRefProfissional)
								with vRefProfissional(intCounter)
									if Trim(.id_cliente) <> "" then
										s = "SELECT " & _
												"*" & _
											" FROM t_CLIENTE_REF_PROFISSIONAL" & _
											" WHERE" & _
												" (id_cliente = '" & Trim(.id_cliente) & "')" & _
												" AND (nome_empresa = '" & Trim("" & .nome_empresa) & "')" & _
												" AND (cargo = '" & Trim("" & .cargo) & "')"
										if r.State <> 0 then r.Close
										r.Open s, cn
										if r.EOF then 
											r.AddNew 
											criou_novo_reg_aux = True
										'	CAMPOS DA CHAVE PRIMÁRIA
											r("id_cliente") = Trim(.id_cliente)
											r("nome_empresa") = Trim(.nome_empresa)
											r("cargo") = Trim(.cargo)
											r("dt_cadastro") = Date
											r("usuario_cadastro") = usuario
										else
											criou_novo_reg_aux = False
											log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir_ref_profissional
											end if
											
										r("ordem") = .ordem
										r("ddd") = Trim(.ddd)
										r("telefone") = Trim(.telefone)
										r("periodo_registro") = .periodo_registro
										r("rendimentos") = .rendimentos
										r("cnpj") = .cnpj
										r("excluido_status") = 0
										
										r.Update
										
										If Err = 0 then 
										'	PREPARA O LOG
											log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir_ref_profissional
											if criou_novo_reg_aux then
												s_log_aux = "Ref Profissional incluída: " & log_via_vetor_monta_inclusao(vLog2)
											else
												s_log_aux = log_via_vetor_monta_alteracao(vLog1, vLog2)
												if s_log_aux <> "" then 
													s_log_aux="Ref Profissional alterada (empresa: " & Trim(.nome_empresa) & ", cargo: " & Trim(.cargo) & "): " & s_log_aux
													end if
												end if
												
											if s_log_aux <> "" then
												if s_log <> "" then s_log = s_log & "; "
												s_log = s_log & s_log_aux
												end if
										else
											erro_fatal=True
											alerta = "FALHA AO GRAVAR OS DADOS DA REF PROFISSIONAL (" & Cstr(Err) & ": " & Err.Description & ")."
											end if
				
										r.Close
										set r = nothing
										if Not cria_recordset_otimista(r, msg_erro) then 
											erro_fatal=True
											alerta = "FALHA AO CRIAR RECORDSET"
											end if
										end if
									end with
								next
							
							if Not erro_fatal then
							'	DADOS P/ O LOG
								s_log_aux=""
								s="SELECT * FROM t_CLIENTE_REF_PROFISSIONAL WHERE (id_cliente = '" & cliente_selecionado & "') AND (excluido_status=1) ORDER BY ordem"
								set r = cn.Execute(s)
								do while Not r.EOF
									log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir_ref_profissional
									if s_log_aux <> "" then s_log_aux = s_log_aux & "; "
									s_log_aux = s_log_aux & "Ref Profissional excluída: " & log_via_vetor_monta_exclusao(vLog1)
									r.MoveNext
									loop
								
								r.Close
								set r = nothing
								if Not cria_recordset_otimista(r, msg_erro) then 
									erro_fatal=True
									alerta = "FALHA AO CRIAR RECORDSET"
									end if
								
								if s_log_aux <> "" then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & s_log_aux
									end if
								
								s="DELETE FROM t_CLIENTE_REF_PROFISSIONAL WHERE (id_cliente = '" & cliente_selecionado & "') AND (excluido_status=1)"
								cn.Execute(s)
								If Err <> 0 then 
									erro_fatal=True
									alerta = "FALHA AO ALTERAR DADOS DE REF PROFISSIONAL DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
									end if
								end if
							end if
						end if
					end if '(if blnCadRefProfissional)

			'	REF COMERCIAL
				if blnCadRefComercial then
					if Not erro_fatal then
						s="UPDATE t_CLIENTE_REF_COMERCIAL SET excluido_status=1 WHERE (id_cliente = '" & cliente_selecionado & "')"
						cn.Execute(s)
						If Err <> 0 then 
							erro_fatal=True
							alerta = "FALHA AO PREPARAR ALTERAÇÃO DOS DADOS DE REF COMERCIAL DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						
						if Not erro_fatal then
							for intCounter=Lbound(vRefComercial) to Ubound(vRefComercial)
								with vRefComercial(intCounter)
									if Trim(.id_cliente) <> "" then
										s = "SELECT " & _
												"*" & _
											" FROM t_CLIENTE_REF_COMERCIAL" & _
											" WHERE" & _
												" (id_cliente = '" & Trim(.id_cliente) & "')" & _
												" AND (nome_empresa = '" & Trim("" & .nome_empresa) & "')"
										if r.State <> 0 then r.Close
										r.Open s, cn
										if r.EOF then 
											r.AddNew 
											criou_novo_reg_aux = True
										'	CAMPOS DA CHAVE PRIMÁRIA
											r("id_cliente") = Trim(.id_cliente)
											r("nome_empresa") = Trim(.nome_empresa)
											r("dt_cadastro") = Date
											r("usuario_cadastro") = usuario
										else
											criou_novo_reg_aux = False
											log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir_ref_comercial
											end if
										
										r("ordem") = .ordem
										r("contato") = Trim(.contato)
										r("ddd") = Trim(.ddd)
										r("telefone") = Trim(.telefone)
										r("excluido_status") = 0
										
										r.Update
										
										If Err = 0 then 
										'	PREPARA O LOG
											log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir_ref_comercial
											if criou_novo_reg_aux then
												s_log_aux = "Ref Comercial incluída: " & log_via_vetor_monta_inclusao(vLog2)
											else
												s_log_aux = log_via_vetor_monta_alteracao(vLog1, vLog2)
												if s_log_aux <> "" then 
													s_log_aux="Ref Comercial alterada (empresa: " & Trim(.nome_empresa) & "): " & s_log_aux
													end if
												end if
												
											if s_log_aux <> "" then
												if s_log <> "" then s_log = s_log & "; "
												s_log = s_log & s_log_aux
												end if
										else
											erro_fatal=True
											alerta = "FALHA AO GRAVAR OS DADOS DA REF COMERCIAL (" & Cstr(Err) & ": " & Err.Description & ")."
											end if
				
										r.Close
										set r = nothing
										if Not cria_recordset_otimista(r, msg_erro) then 
											erro_fatal=True
											alerta = "FALHA AO CRIAR RECORDSET"
											end if
										end if
									end with
								next
							
							if Not erro_fatal then
							'	DADOS P/ O LOG
								s_log_aux=""
								s="SELECT * FROM t_CLIENTE_REF_COMERCIAL WHERE (id_cliente = '" & cliente_selecionado & "') AND (excluido_status=1) ORDER BY ordem"
								set r = cn.Execute(s)
								do while Not r.EOF
									log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir_ref_comercial
									if s_log_aux <> "" then s_log_aux = s_log_aux & "; "
									s_log_aux = s_log_aux & "Ref Comercial excluída: " & log_via_vetor_monta_exclusao(vLog1)
									r.MoveNext
									loop
								
								r.Close
								set r = nothing
								if Not cria_recordset_otimista(r, msg_erro) then 
									erro_fatal=True
									alerta = "FALHA AO CRIAR RECORDSET"
									end if
								
								if s_log_aux <> "" then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & s_log_aux
									end if
								
								s="DELETE FROM t_CLIENTE_REF_COMERCIAL WHERE (id_cliente = '" & cliente_selecionado & "') AND (excluido_status=1)"
								cn.Execute(s)
								If Err <> 0 then 
									erro_fatal=True
									alerta = "FALHA AO ALTERAR DADOS DE REF COMERCIAL DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
									end if
								end if
							end if
						end if
					end if '(if blnCadRefComercial)
					
				if Not erro_fatal then
				'	GRAVA O LOG
					if criou_novo_reg_cliente then
						if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
							if s_log <> "" then s_log = s_log & ";"
							s_log = s_log & " Operação de origem: cadastramento semi-automático de pedido do e-commerce (nº Magento=" & c_numero_magento & ", t_MAGENTO_API_PEDIDO_XML.id=" & id_magento_api_pedido_xml & ")"
							end if
						if s_log <> "" then grava_log usuario, loja, "", cliente_selecionado, OP_LOG_CLIENTE_INCLUSAO, s_log
					else
						if s_log <> "" then grava_log usuario, loja, "", cliente_selecionado, OP_LOG_CLIENTE_ALTERACAO, s_log

						if blnHaPedidoAprovadoComEntregaPendente And (s_log <> "") then
							if (Instr(s_log, "endereco") <> 0) Or (Instr(s_log, "bairro") <> 0) Or (Instr(s_log, "cidade") <> 0) Or (Instr(s_log, "uf") <> 0) Or (Instr(s_log, "cep") <> 0) Or (Instr(s_log, "endereco_numero") <> 0) Or (Instr(s_log, "endereco_complemento") <> 0) then
								'Envia alerta de que houve edição no cadastro de cliente que possui pedido com status de análise de crédito 'crédito ok' e com entrega pendente
								dim rEmailDestinatario
								dim corpo_mensagem, id_email, msg_erro_grava_email
								set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoCadastroClienteComPedidoCreditoOkEntregaPendente)
								if Trim("" & rEmailDestinatario.campo_texto) <> "" then
									s_log_aux = substitui_caracteres(s_log, ";", vbCrLf)
									corpo_mensagem = "O usuário '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na loja " & loja & " o cadastro do cliente:" & vbCrLf & _
													 cnpj_cpf_formata(s_cnpj_cpf) & " - " & s_nome_original & vbCrLf & _
													 "Esse cliente possui pedido com status de análise de crédito 'Crédito OK' e com entrega pendente:" & vbCrLf & _
													 listaPedidoAprovadoComEntregaPendente & _
													 vbCrLf & vbCrLf & _
													 "Informações detalhadas sobre as alterações:" & vbCrLf & _
													 s_log_aux

									EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__SENTINELA_SISTEMA), _
																	"", _
																	rEmailDestinatario.campo_texto, _
																	"", _
																	"", _
																	"Edição no cadastro de cliente que possui pedido com status 'Crédito OK' e entrega pendente", _
																	corpo_mensagem, _
																	Now, _
																	id_email, _
																	msg_erro_grava_email
									end if
								end if
							end if
						end if
					end if
				
				' ANTES DA MEMORIZAÇÃO DE ENDEREÇOS FAZÍAMOS A SELEÇÃO AUTOMÁTICA DE TRANSPORTADORA BASEADO NO CEP DE PEDIDOS SEM NOTA FISCAL EMITIDA
				' AGORA FAZEMOS ISSO NA EDIÇÃO DO PEDIDO (PEDIDOATUALIZA.ASP)
				
				if Not erro_fatal then
				'	~~~~~~~~~~~~~~
					cn.CommitTrans
				'	~~~~~~~~~~~~~~
				else
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					end if
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
	<title>LOJA</title>
	</head>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var fCepPopup;

$(function () {
	var f;
	if ((typeof (fNEW) !== "undefined") && (fNEW !== null)) {
		f = fNEW;

		if (!f.rb_end_entrega[1].checked) {
            Disabled_change(f, true);
		}

		if (trim(fNEW.c_FormFieldValues.value) != "") {
			stringToForm(fNEW.c_FormFieldValues.value, $('#fNEW'));
		}
        trataProdutorRuralEndEtg_PF(null);
        trocarEndEtgTipoPessoa(null);
    }
});

function copyMagentoShipAddrToShipAddr() {
	var fFROM, fTO;
	var s, eh_cpf;
	fFROM = fMAG;
	fTO = fNEW;
	fTO.EndEtg_endereco.value = fFROM.c_mag_end_etg_endereco.value;
	fTO.EndEtg_endereco_numero.value = fFROM.c_mag_end_etg_endereco_numero.value;
	fTO.EndEtg_endereco_complemento.value = fFROM.c_mag_end_etg_complemento.value;
	fTO.EndEtg_bairro.value = fFROM.c_mag_end_etg_bairro.value;
	fTO.EndEtg_cidade.value = fFROM.c_mag_end_etg_cidade.value;
	fTO.EndEtg_uf.value = fFROM.c_mag_end_etg_uf.value;
	fTO.EndEtg_cep.value = cep_formata(fFROM.c_mag_end_etg_cep.value);
}

function Disabled_True(f) {
    Disabled_change(f, true);
}
function Disabled_False(f) {
    Disabled_change(f, false);
}

function Disabled_change(f, value) {

    if(f.EndEtg_nome) f.EndEtg_nome.disabled = value;
    f.EndEtg_endereco.disabled = value;
    f.EndEtg_endereco_numero.disabled = value;
    f.EndEtg_bairro.disabled = value;
    f.EndEtg_cidade.disabled = value;
    f.EndEtg_obs.disabled = value;
    f.EndEtg_uf.disabled = value;
    f.EndEtg_cep.disabled = value;
    f.bPesqCepEndEtgNovo.disabled = value;
    f.EndEtg_endereco_complemento.disabled = value;

    var lista = $(".Habilitar_EndEtg_outroendereco input");
    for (var i = 0; i < lista.length; i++) {
        lista[i].disabled = value;
    }
    trocarEndEtgTipoPessoa(null);
}

function ProcessaSelecaoCEP(){};

function AbrePesquisaCepEndEtg(){
var f, strUrl;
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
	f=fNEW;
	ProcessaSelecaoCEP=TrataCepEnderecoEntrega;
	strUrl="../Global/AjaxCepPesqPopup.asp";
	if (trim(f.EndEtg_cep.value)!="") strUrl=strUrl+"?CepDefault="+trim(f.EndEtg_cep.value);
	fCepPopup=window.open(strUrl, "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=980,height=650,left=0,top=0");
	fCepPopup.focus();
}

function TrataCepEnderecoEntrega(strCep, strUF, strLocalidade, strBairro, strLogradouro, strEnderecoNumero, strEnderecoComplemento) {
var f;
	f=fNEW;
	f.EndEtg_cep.value=cep_formata(strCep);
	f.EndEtg_uf.value=strUF;
	f.EndEtg_cidade.value=strLocalidade;
	f.EndEtg_bairro.value=strBairro;
	f.EndEtg_endereco.value=strLogradouro;
	f.EndEtg_endereco_numero.value=strEnderecoNumero;
	f.EndEtg_endereco_complemento.value=strEnderecoComplemento;
	f.EndEtg_endereco.focus();
	window.status="Concluído";
}


function fNEWConcluir( f ){
	if ((!f.rb_end_entrega[0].checked)&&(!f.rb_end_entrega[1].checked)) {
		alert('Informe se o endereço de entrega será o mesmo endereço do cadastro ou não!!');
		return;
		}

	if (f.rb_end_entrega[1].checked) {
		if (trim(f.EndEtg_endereco.value)=="") {
			alert('Preencha o endereço de entrega!!');
			f.EndEtg_endereco.focus();
			return;
			}

		if (trim(f.EndEtg_endereco_numero.value)=="") {
			alert('Preencha o número do endereço de entrega!!');
			f.EndEtg_endereco_numero.focus();
			return;
			}
			
		if (trim(f.EndEtg_bairro.value)=="") {
			alert('Preencha o bairro do endereço de entrega!!');
			f.EndEtg_bairro.focus();
			return;
			}

		if (trim(f.EndEtg_cidade.value)=="") {
			alert('Preencha a cidade do endereço de entrega!!');
			f.EndEtg_cidade.focus();
			return;
			}
		if (trim(f.EndEtg_obs.value) == "") {
		    alert('Selecione a justificativa do endereço de entrega!!');
		    f.EndEtg_obs.focus();
		    return;
		    }
		s=trim(f.EndEtg_uf.value);
		if ((s=="")||(!uf_ok(s))) {
			alert('UF inválida no endereço de entrega!!');
			f.EndEtg_uf.focus();
			return;
			}
			
		if (!cep_ok(f.EndEtg_cep.value)) {
			alert('CEP inválido no endereço de entrega!!');
			f.EndEtg_cep.focus();
			return;
			}


<%if blnUsarMemorizacaoCompletaEnderecos then%>
<%if Not eh_cpf then%>
            var EndEtg_tipo_pessoa = $('input[name="EndEtg_tipo_pessoa"]:checked').val();
            if (!EndEtg_tipo_pessoa)
                EndEtg_tipo_pessoa = "";
            if (EndEtg_tipo_pessoa != "PJ" && EndEtg_tipo_pessoa != "PF") {
                alert('Necessário escolher Pessoa Jurídica ou Pessoa Física no Endereço de entrega!!');
                f.EndEtg_tipo_pessoa.focus();
                return;
            }

            if (EndEtg_tipo_pessoa == "PJ") {
                //Campos PJ: 

                if (f.EndEtg_cnpj_cpf_PJ.value == "" || !cnpj_ok(f.EndEtg_cnpj_cpf_PJ.value)) {
                    alert('Endereço de entrega: CNPJ inválido!!');
                    f.EndEtg_cnpj_cpf_PJ.focus();
                    return;
                }

                if ($('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').length == 0) {
                    alert('Endereço de entrega: informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
                    f.EndEtg_contribuinte_icms_status_PJ.focus();
                    return;
                }

                if ((f.EndEtg_contribuinte_icms_status_PJ[1].checked) && (trim(f.EndEtg_ie_PJ.value) == "")) {
                    alert('Endereço de entrega: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
                    f.EndEtg_ie_PJ.focus();
                    return;
                }
                if ((f.EndEtg_contribuinte_icms_status_PJ[0].checked) && (f.EndEtg_ie_PJ.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Endereço de entrega: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                    f.EndEtg_ie_PJ.focus();
                    return;
                }
                if ((f.EndEtg_contribuinte_icms_status_PJ[1].checked) && (f.EndEtg_ie_PJ.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Endereço de entrega: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                    f.EndEtg_ie_PJ.focus();
                    return;
                }
                if (f.EndEtg_contribuinte_icms_status_PJ[2].checked) {
                    if (f.EndEtg_ie_PJ.value != "") {
                        alert("Endereço de entrega: se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
                        f.EndEtg_ie_PF.focus();
                        return;
                    }
                }

                if (trim(f.EndEtg_nome.value) == "") {
                    alert('Preencha a razão social no endereço de entrega!!');
                    f.EndEtg_nome.focus();
                    return;
                }

                /*
                telefones PJ:
                EndEtg_ddd_com
                EndEtg_tel_com
                EndEtg_ramal_com
                EndEtg_ddd_com_2
                EndEtg_tel_com_2
                EndEtg_ramal_com_2
*/

                if (!ddd_ok(f.EndEtg_ddd_com.value)) {
                    alert('Endereço de entrega: DDD inválido!!');
                    f.EndEtg_ddd_com.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_com.value)) {
                    alert('Endereço de entrega: telefone inválido!!');
                    f.EndEtg_tel_com.focus();
                    return;
                }
                if ((f.EndEtg_ddd_com.value == "") && (f.EndEtg_tel_com.value != "")) {
                    alert('Endereço de entrega: preencha o DDD do telefone.');
                    f.EndEtg_ddd_com.focus();
                    return;
                }
                if ((f.EndEtg_tel_com.value == "") && (f.EndEtg_ddd_com.value != "")) {
                    alert('Endereço de entrega: preencha o telefone.');
                    f.EndEtg_tel_com.focus();
                    return;
                }
                if (trim(f.EndEtg_ddd_com.value) == "" && trim(f.EndEtg_ramal_com.value) != "") {
                    alert('Endereço de entrega: DDD comercial inválido!!');
                    f.EndEtg_ddd_com.focus();
                    return;
                }



                if (!ddd_ok(f.EndEtg_ddd_com_2.value)) {
                    alert('Endereço de entrega: DDD inválido!!');
                    f.EndEtg_ddd_com_2.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_com_2.value)) {
                    alert('Endereço de entrega: telefone inválido!!');
                    f.EndEtg_tel_com_2.focus();
                    return;
                }
                if ((f.EndEtg_ddd_com_2.value == "") && (f.EndEtg_tel_com_2.value != "")) {
                    alert('Endereço de entrega: preencha o DDD do telefone.');
                    f.EndEtg_ddd_com_2.focus();
                    return;
                }
                if ((f.EndEtg_tel_com_2.value == "") && (f.EndEtg_ddd_com_2.value != "")) {
                    alert('Endereço de entrega: preencha o telefone.');
                    f.EndEtg_tel_com_2.focus();
                    return;
                }
                if (trim(f.EndEtg_ddd_com_2.value) == "" && trim(f.EndEtg_ramal_com_2.value) != "") {
                    alert('Endereço de entrega: DDD comercial 2 inválido!!');
                    f.EndEtg_ddd_com_2.focus();
                    return;
                }

            }
            else {
                //campos PF

                if (f.EndEtg_cnpj_cpf_PF.value == "" || !cpf_ok(f.EndEtg_cnpj_cpf_PF.value)) {
                    alert('Endereço de entrega: CPF inválido!!');
                    f.EndEtg_cnpj_cpf_PF.focus();
                    return;
                }

                if ((!f.EndEtg_produtor_rural_status_PF[0].checked) && (!f.EndEtg_produtor_rural_status_PF[1].checked)) {
                    alert('Endereço de entrega: informe se o cliente é produtor rural ou não!!');
                    return;
                }
                if (!f.EndEtg_produtor_rural_status_PF[0].checked) {
                    if (!f.EndEtg_contribuinte_icms_status_PF[1].checked) {
                        alert('Endereço de entrega: para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!!');
                        return;
                    }
                    if ((!f.EndEtg_contribuinte_icms_status_PF[0].checked) && (!f.EndEtg_contribuinte_icms_status_PF[1].checked) && (!f.EndEtg_contribuinte_icms_status_PF[2].checked)) {
                        alert('Endereço de entrega: informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
                        return;
                    }
                    if ((f.EndEtg_contribuinte_icms_status_PF[1].checked) && (trim(f.EndEtg_ie_PF.value) == "")) {
                        alert('Endereço de entrega: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
                        f.EndEtg_ie_PF.focus();
                        return;
                    }
                    if ((f.EndEtg_contribuinte_icms_status_PF[0].checked) && (f.EndEtg_ie_PF.value.toUpperCase().indexOf('ISEN') >= 0)) {
                        alert('Endereço de entrega: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                        f.EndEtg_ie_PF.focus();
                        return;
                    }
                    if ((f.EndEtg_contribuinte_icms_status_PF[1].checked) && (f.EndEtg_ie_PF.value.toUpperCase().indexOf('ISEN') >= 0)) {
                        alert('Endereço de entrega: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                        f.EndEtg_ie_PF.focus();
                        return;
                    }

                    if (f.EndEtg_contribuinte_icms_status_PF[2].checked) {
                        if (f.EndEtg_ie_PF.value != "") {
                            alert("Endereço de entrega: se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
                            f.EndEtg_ie_PF.focus();
                            return;
                        }
                    }
                }
            
                if (trim(f.EndEtg_nome.value) == "") {
                    alert('Preencha o nome no endereço de entrega!!');
                    f.EndEtg_nome.focus();
                    return;
                }

                /*
                telefones PF:
                EndEtg_ddd_res
                EndEtg_tel_res
                EndEtg_ddd_cel
                EndEtg_tel_cel
                */
                if (!ddd_ok(f.EndEtg_ddd_res.value)) {
                    alert('Endereço de entrega: DDD inválido!!');
                    f.EndEtg_ddd_res.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_res.value)) {
                    alert('Endereço de entrega: telefone inválido!!');
                    f.EndEtg_tel_res.focus();
                    return;
                }
                if ((trim(f.EndEtg_ddd_res.value) != "") || (trim(f.EndEtg_tel_res.value) != "")) {
                    if (trim(f.EndEtg_ddd_res.value) == "") {
                        alert('Endereço de entrega: preencha o DDD!!');
                        f.EndEtg_ddd_res.focus();
                        return;
                    }
                    if (trim(f.EndEtg_tel_res.value) == "") {
                        alert('Endereço de entrega: preencha o telefone!!');
                        f.EndEtg_tel_res.focus();
                        return;
                    }
                }

                if (!ddd_ok(f.EndEtg_ddd_cel.value)) {
                    alert('Endereço de entrega: DDD inválido!!');
                    f.EndEtg_ddd_cel.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_cel.value)) {
                    alert('Endereço de entrega: telefone inválido!!');
                    f.EndEtg_tel_cel.focus();
                    return;
                }
                if ((f.EndEtg_ddd_cel.value == "") && (f.EndEtg_tel_cel.value != "")) {
                    alert('Endereço de entrega: preencha o DDD do celular.');
                    f.EndEtg_tel_cel.focus();
                    return;
                }
                if ((f.EndEtg_tel_cel.value == "") && (f.EndEtg_ddd_cel.value != "")) {
                    alert('Endereço de entrega: preencha o número do celular.');
                    f.EndEtg_tel_cel.focus();
                    return;
                }


            }


<%end if%>
<%end if%>

		}

	fNEW.c_FormFieldValues.value = formToString($("#fNEW"));

    //campos do endereço de entrega que precisam de transformacao
    transferirCamposEndEtg(fNEW);

	dPEDIDO.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit(); 
}



    function transferirCamposEndEtg(fNEW) {
<%if blnUsarMemorizacaoCompletaEnderecos then %>
    <%if Not eh_cpf then %>
        //Transferimos os dados do endereço de entrega dos campos certos. 
        //Temos dois conjuntos de campos (para PF e PJ) porque o layout é muito diferente.
        var pj = $('input[name="EndEtg_tipo_pessoa"]:checked').val() == "PJ";
        if (pj) {
            fNEW.EndEtg_cnpj_cpf.value = fNEW.EndEtg_cnpj_cpf_PJ.value;
            fNEW.EndEtg_ie.value = fNEW.EndEtg_ie_PJ.value;
            fNEW.EndEtg_contribuinte_icms_status.value = $('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').val();
            if (!$('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').val())
                fNEW.EndEtg_contribuinte_icms_status.value = "";
        }
        else {
            fNEW.EndEtg_cnpj_cpf.value = fNEW.EndEtg_cnpj_cpf_PF.value;
            fNEW.EndEtg_ie.value = fNEW.EndEtg_ie_PF.value;
            fNEW.EndEtg_contribuinte_icms_status.value = $('input[name="EndEtg_contribuinte_icms_status_PF"]:checked').val();
            if (!$('input[name="EndEtg_contribuinte_icms_status_PF"]:checked').val())
                fNEW.EndEtg_contribuinte_icms_status.value = "";
            fNEW.EndEtg_produtor_rural_status.value = $('input[name="EndEtg_produtor_rural_status_PF"]:checked').val();
            if (!$('input[name="EndEtg_produtor_rural_status_PF"]:checked').val())
                fNEW.EndEtg_produtor_rural_status.value = "";
        }

        //os campos a mais são enviados junto. Deixamos enviar...
    <%end if%>
<%end if%>
    }

    //para mudar o tipo do endereço de entrega
    function trocarEndEtgTipoPessoa(novoTipo) {
<%if blnUsarMemorizacaoCompletaEnderecos then%>
        if (novoTipo && $('input[name="EndEtg_tipo_pessoa"]:disabled').length == 0)
            setarValorRadio($('input[name="EndEtg_tipo_pessoa"]'), novoTipo);

        var pj = $('input[name="EndEtg_tipo_pessoa"]:checked').val() == "PJ";

        if (pj) {
            $(".Mostrar_EndEtg_pf").css("display", "none");
            $(".Mostrar_EndEtg_pj").css("display", "");
            $("#Label_EndEtg_nome").text("RAZÃO SOCIAL");
        }
        else {
            //display block prejudica as tabelas
            $(".Mostrar_EndEtg_pf").css("display", "");
            $(".Mostrar_EndEtg_pj").css("display", "none");
            $("#Label_EndEtg_nome").text("NOME");
        }
<%else%>
        //oculta todos
        $(".Mostrar_EndEtg_pf").css("display", "none");
        $(".Mostrar_EndEtg_pj").css("display", "none");
        $(".Habilitar_EndEtg_outroendereco").css("display", "none");
<%end if%>
    }

    function trataContribuinteIcmsEndEtg_PJ(novoTipo)
    {
        if (novoTipo && $('input[name="EndEtg_contribuinte_icms_status_PJ"]:disabled').length == 0)
            setarValorRadio($('input[name="EndEtg_contribuinte_icms_status_PJ"]'),novoTipo);
    }
    function trataContribuinteIcmsEndEtg_PF(novoTipo)
    {
        if (novoTipo && $('input[name="EndEtg_contribuinte_icms_status_PF"]:disabled').length == 0)
            setarValorRadio($('input[name="EndEtg_contribuinte_icms_status_PF"]'),novoTipo);
    }

    function trataProdutorRuralEndEtg_PF(novoTipo) {
        //ao clicar na opção Produtor Rural, exibir/ocultar os campos apropriados (endereço de entrega)
        if (novoTipo && $('input[name="EndEtg_produtor_rural_status_PF"]:disabled').length == 0)
            setarValorRadio($('input[name="EndEtg_produtor_rural_status_PF"]'), novoTipo);

        var sim = $('input[name="EndEtg_produtor_rural_status_PF"]:checked').val() == "<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>";

        //contribuinte ICMS sempre aparece para PJ
        if(sim) {
            $(".Mostrar_EndEtg_contribuinte_icms_PF").css("display", "");
        }
        else {
            $(".Mostrar_EndEtg_contribuinte_icms_PF").css("display", "none");
        }
    }

    function trataProdutorRuralEndEtg_PJ(novoTipo) {
        if (novoTipo && $('input[name="EndEtg_produtor_rural_status_PJ"]:disabled').length == 0)
            setarValorRadio($('input[name="EndEtg_produtor_rural_status_PJ"]'), novoTipo);
    }

    //definir um valor como ativo em um radio 
    function setarValorRadio(array, valor)
    {
        for (var i = 0; i < array.length; i++)
        {
            var este = array[i];
            if (este.value == valor)
                este.checked = true;
        }
    }


</script>

<script type="text/javascript">
	function exibeJanelaCEP_Etg() {
		$.mostraJanelaCEP("EndEtg_cep", "EndEtg_uf", "EndEtg_cidade", "EndEtg_bairro", "EndEtg_endereco", "EndEtg_endereco_numero", "EndEtg_endereco_complemento");
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__E_JANELABUSCACEP_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
.TdCliLbl
{
	width:130px;
	text-align:right;
}
.TdCliCel
{
	width:520px;
	text-align:left;
}
.TdCliBtn
{
	width:30px;
	text-align:center;
}
</style>

<body id="corpoPagina">

<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->

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
				s = "CLIENTE " & cnpj_cpf_formata(cnpj_cpf_selecionado) & " CADASTRADO COM SUCESSO."
				exibir_botao_novo_item = True
			case OP_CONSULTA, OP_ALTERA
				s = "CLIENTE " & cnpj_cpf_formata(cnpj_cpf_selecionado) & " ALTERADO COM SUCESSO."
				exibir_botao_novo_item = True
			case OP_EXCLUI
				s = "CLIENTE " & cnpj_cpf_formata(cnpj_cpf_selecionado) & " EXCLUÍDO COM SUCESSO."
			end select			
		if s <> "" then s="<p style='margin:5px 2px 5px 2px;'>" & s & "</p>"
		end if
%>
<% if alerta = "" then %>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<% else %>
<div class=<%=s_aux%> style="width:649px;font-weight:bold;" align="center"><%=s%></div>
	<% if s_tabela_municipios_IBGE <> "" then %>
		<br /><br />
		<%=s_tabela_municipios_IBGE%>
	<% end if %>
<% end if %>
<br /><br />


<!-- ************   FORM PARA OPÇÃO DE CADASTRAR NOVO PEDIDO?  ************ -->
<% if exibir_botao_novo_item then %>

	<% if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then %>
	<!--  DADOS DO MAGENTO  -->
	<table class="Qx" cellspacing="0">
		<tr style="background-color:azure;">
			<td colspan="3" class="MC MB ME MD" align="center"><span class="N">Dados do Magento (pedido nº <%=c_numero_magento%>)</span></td>
		</tr>
		<tr>
			<td class="MB ME MD TdCliLbl"><span class="PLTd">Cliente</span></td>
			<td class="MB MD TdCliCel">
				<span class="C"><%=s_nome_cliente%></span>
				<% if c_mag_customer_dob <> "" then %>
				<span class="C">&nbsp;(<%=c_mag_customer_dob%>)</span>
				<% end if %>
				<% if c_mag_cpf_cnpj_identificado <> "" then %>
				<br /><span class="C"><%=cnpj_cpf_formata(c_mag_cpf_cnpj_identificado)%></span>
				<% end if %>
				<% if c_mag_email_identificado <> "" then %>
				<br /><span class="C"><%=c_mag_email_identificado%></span>
				<% end if %>
			</td>
		</tr>
		<tr>
			<td class="MB ME MD TdCliLbl"><span class="PLTd">Endereço de Cobrança</span></td>
			<td class="MB MD TdCliCel">
				<span class="C"><%=s_mag_end_cob_completo%></span>
				<% if c_mag_end_cob_telephone_numero <> "" then %>
				<br /><span class="C"><%=formata_ddd_telefone_ramal(c_mag_end_cob_telephone_ddd, c_mag_end_cob_telephone_numero, Null)%></span>
				<% end if %>
				<% if c_mag_end_cob_celular_numero <> "" then %>
				<br /><span class="C"><%=formata_ddd_telefone_ramal(c_mag_end_cob_celular_ddd, c_mag_end_cob_celular_numero, Null)%></span>
				<% end if %>
				<% if c_mag_end_cob_fax_numero <> "" then %>
				<br /><span class="C"><%=formata_ddd_telefone_ramal(c_mag_end_cob_fax_ddd, c_mag_end_cob_fax_numero, Null)%></span>
				<% end if %>
			</td>
		</tr>
		<tr>
			<td class="MB ME MD TdCliLbl"><span class="PLTd">Endereço de Entrega</span></td>
			<td class="MB MD TdCliCel">
				<span class="C"><%=s_mag_end_etg_completo%></span>
				<% if c_mag_end_etg_telephone_numero <> "" then %>
				<br /><span class="C"><%=formata_ddd_telefone_ramal(c_mag_end_etg_telephone_ddd, c_mag_end_etg_telephone_numero, Null)%></span>
				<% end if %>
				<% if c_mag_end_etg_celular_numero <> "" then %>
				<br /><span class="C"><%=formata_ddd_telefone_ramal(c_mag_end_etg_celular_ddd, c_mag_end_etg_celular_numero, Null)%></span>
				<% end if %>
				<% if c_mag_end_etg_fax_numero <> "" then %>
				<br /><span class="C"><%=formata_ddd_telefone_ramal(c_mag_end_etg_fax_ddd, c_mag_end_etg_fax_numero, Null)%></span>
				<% end if %>
			</td>
		</tr>
	</table>

	<table cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;width:698px;">
	<tr><td class="Rc" align="left">&nbsp;</td></tr>
	</table>
	<br />
	<% end if %>

	<% if blnLojaHabilitadaProdCompostoECommerce then
		s = "PedidoNovoProdCompostoMask.asp"
	else
		s = "pedidonovo.asp"
		end if %>
	<form action="<%=s%>" method="post" id="fNEW" name="fNEW" onsubmit="if (!fNEWConcluir(fNEW)) return false">
	<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
	<INPUT type="hidden" name='cliente_selecionado' id="cliente_selecionado" value='<%=cliente_selecionado%>'>
	<INPUT type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=OP_INCLUI%>'>
	<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />
	<input type="hidden" name="operacao_origem" id="operacao_origem" value="<%=operacao_origem%>" />
	<input type="hidden" name="id_magento_api_pedido_xml" id="id_magento_api_pedido_xml" value="<%=id_magento_api_pedido_xml%>" />
	<input type="hidden" name="c_numero_magento" id="c_numero_magento" value="<%=c_numero_magento%>" />
	<input type="hidden" name="operationControlTicket" id="operationControlTicket" value="<%=operationControlTicket%>" />
	<input type="hidden" name="sessionToken" id="sessionToken" value="<%=sessionToken%>" />
	<% if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then %>
	<input type="hidden" name="c_FlagCadSemiAutoPedMagento_FluxoOtimizado" id="c_FlagCadSemiAutoPedMagento_FluxoOtimizado" value="<%=c_FlagCadSemiAutoPedMagento_FluxoOtimizado%>" />
	<input type="hidden" name="rb_indicacao" id="rb_indicacao" value="<%=rb_indicacao%>" />
	<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>" />
	<input type="hidden" name="rb_RA" id="rb_RA" value="<%=rb_RA%>" />
	<% end if %>


<!-- ************   ENDEREÇO DE ENTREGA: S/N   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td align="left">
		<p class="R">ENDEREÇO DE ENTREGA</p><p class="C">
			<% intIdx = 0 %>
			<input type="radio" id="rb_end_entrega" name="rb_end_entrega" value="N"onclick="Disabled_True(fNEW);"><span class="C" style="cursor:default" onclick="fNEW.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_True(fNEW);">O mesmo endereço do cadastro</span>
			<% intIdx = intIdx + 1 %>
			<br><input type="radio" id="rb_end_entrega" name="rb_end_entrega" value="S"onclick="Disabled_False(fNEW);"><span class="C" style="cursor:default" onclick="fNEW.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_False(fNEW);">Outro endereço</span>
		</p>
		</td>
		<td style="width:40px;text-align:right;vertical-align:top;">
            <% if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then %>
    			<a href="javascript:copyMagentoShipAddrToShipAddr();"><img src="../IMAGEM/copia_20x20.png" name="btnMagentoCopyShipAddrToShipAddr" id="btnMagentoCopyShipAddrToShipAddr" title="Altera o endereço usando os dados do endereço de entrega obtidos do Magento" /></a>
            <% end if %>
		</td>
	</tr>
</table>


<!--  ************  TIPO DO ENDEREÇO DE ENTREGA: PF/PJ (SOMENTE SE O CLIENTE FOR PJ)   ************ -->

<%if blnUsarMemorizacaoCompletaEnderecos then%>
    <%if eh_cpf then%>
        <!-- ************   ENDEREÇO DE ENTREGA PARA CLIENTE PF   ************ -->
        <!-- Pegamos todos os atuais. Sem campos editáveis. -->
    <input type="hidden" id="EndEtg_tipo_pessoa" name="EndEtg_tipo_pessoa" value="PF"/>
    <input type="hidden" id="EndEtg_cnpj_cpf" name="EndEtg_cnpj_cpf" value="<%=s_cnpj_cpf%>"/>
    <input type="hidden" id="EndEtg_ie" name="EndEtg_ie" value="<%=s_ie%>"/>
    <input type="hidden" id="EndEtg_contribuinte_icms_status" name="EndEtg_contribuinte_icms_status" value="<%=s_contribuinte_icms%>"/>
    <input type="hidden" id="EndEtg_rg" name="EndEtg_rg" value="<%=s_rg%>"/>
    <input type="hidden" id="EndEtg_produtor_rural_status" name="EndEtg_produtor_rural_status" value="<%=s_produtor_rural%>"/>
    <input type="hidden" id="EndEtg_email" name="EndEtg_email" value="<%=s_email%>"/>
    <input type="hidden" id="EndEtg_email_xml" name="EndEtg_email_xml" value="<%=s_email_xml%>"/>
    <input type="hidden" id="EndEtg_nome" name="EndEtg_nome" value="<%=s_nome%>"/>


    <%else%>

    <table width="649" class="QS Habilitar_EndEtg_outroendereco" cellspacing="0">
	    <tr>
		    <td align="left">
		    <p class="R">TIPO</p><p class="C">
			    <input type="radio" id="EndEtg_tipo_pessoa_PJ" name="EndEtg_tipo_pessoa" value="PJ" onclick="trocarEndEtgTipoPessoa(null);" checked>
			    <span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PJ');">Pessoa Jurídica</span>
			    &nbsp;
			    <input type="radio" id="EndEtg_tipo_pessoa_PF" name="EndEtg_tipo_pessoa" value="PF" onclick="trocarEndEtgTipoPessoa(null);">
			    <span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PF');">Pessoa Física</span>
		    </p>
		    </td>
	    </tr>
    </table>

            <!-- ************   PJ: CNPJ/CONTRIBUINTE ICMS/IE - DO ENDEREÇO DE ENTREGA DE PJ ************ -->
            <!-- ************   PF: CPF/PRODUTOR RURAL/CONTRIBUINTE ICMS/IE - DO ENDEREÇO DE ENTREGA DE PJ  ************ -->
            <!-- fizemos dois conjuntos diferentes de campos porque a ordem é muito diferente -->
            <!-- EndEtg_rg EndEtg_email e EndEtg_email_xml vem diretamente do t_CLIENTE -->
    <input type="hidden" id="EndEtg_cnpj_cpf" name="EndEtg_cnpj_cpf" />
    <input type="hidden" id="EndEtg_ie" name="EndEtg_ie" />
    <input type="hidden" id="EndEtg_contribuinte_icms_status" name="EndEtg_contribuinte_icms_status" />
    <input type="hidden" id="EndEtg_rg" name="EndEtg_rg" value="<%=s_rg%>"/>
    <input type="hidden" id="EndEtg_produtor_rural_status" name="EndEtg_produtor_rural_status" />
    <input type="hidden" id="EndEtg_email" name="EndEtg_email" value="<%=s_email%>"/>
    <input type="hidden" id="EndEtg_email_xml" name="EndEtg_email_xml" value="<%=s_email_xml%>"/>



    <table width="649" class="QS Habilitar_EndEtg_outroendereco Mostrar_EndEtg_pj" cellspacing="0">
	    <tr>
		    <td width="210" align="left">
	    <p class="R">CNPJ</p><p class="C">
	    <input id="EndEtg_cnpj_cpf_PJ" name="EndEtg_cnpj_cpf_PJ" class="TA" value="" size="22" style="text-align:center; color:#0000ff"></p></td>

	    <td class="MDE" width="215" align="left"><p class="R">IE</p><p class="C">
		    <input id="EndEtg_ie_PJ" name="EndEtg_ie_PJ" class="TA" type="text" maxlength="20" size="25" value="" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_Nome.focus(); filtra_nome_identificador();"></p></td>

	    <td align="left" class="Mostrar_EndEtg_contribuinte_icms_PJ"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PJ_nao" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">Não</span>
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PJ_sim" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>');">Sim</span>
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PJ_isento" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>');">Isento</span></p></td>
	    </tr>
    </table>

    <table width="649" class="QS Habilitar_EndEtg_outroendereco Mostrar_EndEtg_pf" cellspacing="0">
	    <tr>
		    <td width="210" align="left">
	    <p class="R">CPF</p><p class="C">
	    <input id="EndEtg_cnpj_cpf_PF" name="EndEtg_cnpj_cpf_PF" class="TA" value="" size="22" style="text-align:center; color:#0000ff"></p></td>

	    <td align="left" class="ME" style="min-width: 110px;" ><p class="R">PRODUTOR RURAL</p><p class="C">
		    <input type="radio" id="EndEtg_produtor_rural_status_PF_nao" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>');">Não</span>
		    <input type="radio" id="EndEtg_produtor_rural_status_PF_sim" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>')">Sim</span></p></td>

	    <td align="left" class="MDE Mostrar_EndEtg_contribuinte_icms_PF"><p class="R">IE</p><p class="C">
		    <input id="EndEtg_ie_PF" name="EndEtg_ie_PF" class="TA" type="text" maxlength="20" size="13" value="" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_nome.focus(); filtra_nome_identificador();"></p>
	    </td>

	    <td align="left" class="Mostrar_EndEtg_contribuinte_icms_PF" ><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PF_nao" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">Não</span>
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PF_sim" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>');">Sim</span>
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PF_isento" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>');">Isento</span></p>
	    </td>
	    </tr>
    </table>


    <!-- ************   ENDEREÇO DE ENTREGA: NOME  ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td width="100%" align="left"><p class="R" id="Label_EndEtg_nome">RAZÃO SOCIAL</p><p class="C">
		    <input id="EndEtg_nome" name="EndEtg_nome" class="TA" value="" maxlength="60" size="85" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fNEW.EndEtg_endereco.focus(); filtra_nome_identificador();"></p></td>
	    </tr>
    </table>


    <%end if%>
<%end if%>


<!-- ************   ENDEREÇO DE ENTREGA: ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0" type="hidden">
	<tr>
	<td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C">
		<input id="EndEtg_endereco" name="EndEtg_endereco" class="TA" value="" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fNEW.EndEtg_endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO DE ENTREGA: Nº/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">Nº</p><p class="C">
		<input id="EndEtg_endereco_numero" name="EndEtg_endereco_numero" class="TA" value="" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fNEW.EndEtg_endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td width="50%" align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<input id="EndEtg_endereco_complemento" name="EndEtg_endereco_complemento" class="TA" value="" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fNEW.EndEtg_bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO DE ENTREGA: BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">BAIRRO</p><p class="C">
		<input id="EndEtg_bairro" name="EndEtg_bairro" class="TA" value="" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fNEW.EndEtg_cidade.focus(); filtra_nome_identificador();"></p></td>
	<td width="50%" align="left"><p class="R">CIDADE</p><p class="C">
		<input id="EndEtg_cidade" name="EndEtg_cidade" class="TA" value="" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fNEW.EndEtg_uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO DE ENTREGA: UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="50%" class="MD" align="left"><p class="R">UF</p><p class="C">
		<input id="EndEtg_uf" name="EndEtg_uf" class="TA" value="" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fNEW.EndEtg_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
	<td width="50%" align="left">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="left"><p class="R">CEP</p><p class="C">
				<input id="EndEtg_cep" name="EndEtg_cep" readonly tabindex=-1 class="TA" value="" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
			<td align="center" width="50%">
				<% if blnPesquisaCEPAntiga then %>
				<button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCepEndEtg();">Pesquisar CEP</button>
				<% end if %>
				<% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
				<% if blnPesquisaCEPNova then %>
				<button type="button" name="bPesqCepEndEtgNovo" id="bPesqCepEndEtgNovo" style='width:130px;font-size:10pt;' class="Botao" onclick="exibeJanelaCEP_Etg();">&nbsp;Busca de CEP&nbsp;</button>
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
	</tr>
</table>


<%if blnUsarMemorizacaoCompletaEnderecos then%>
    <%if eh_cpf then%>

        <!-- ************   ENDEREÇO DE ENTREGA PARA PF: TELEFONES   ************ -->
        <!-- pegamos todos em branco -->
        <input type="hidden" id="EndEtg_ddd_res" name="EndEtg_ddd_res" value=""/>
        <input type="hidden" id="EndEtg_tel_res" name="EndEtg_tel_res" value=""/>
        <input type="hidden" id="EndEtg_ddd_cel" name="EndEtg_ddd_cel" value=""/>
        <input type="hidden" id="EndEtg_tel_cel" name="EndEtg_tel_cel" value=""/>
        <input type="hidden" id="EndEtg_ddd_com" name="EndEtg_ddd_com" value=""/>
        <input type="hidden" id="EndEtg_tel_com" name="EndEtg_tel_com" value=""/>
        <input type="hidden" id="EndEtg_ramal_com" name="EndEtg_ramal_com" value=""/>
        <input type="hidden" id="EndEtg_ddd_com_2" name="EndEtg_ddd_com_2" value=""/>
        <input type="hidden" id="EndEtg_tel_com_2" name="EndEtg_tel_com_2" value=""/>
        <input type="hidden" id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" value=""/>

    <%else%>
        
        
        <!-- ************   ENDEREÇO DE ENTREGA: TELEFONE RESIDENCIAL   ************ -->
        <table width="649" class="QS Mostrar_EndEtg_pf Habilitar_EndEtg_outroendereco" cellspacing="0">
	        <tr>
	        <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="EndEtg_ddd_res" name="EndEtg_ddd_res" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fNEW.EndEtg_tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	        <td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
		        <input id="EndEtg_tel_res" name="EndEtg_tel_res" class="TA" value="" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fNEW.EndEtg_ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        </tr>
	        <tr>
	        <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="EndEtg_ddd_cel" name="EndEtg_ddd_cel" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fNEW.EndEtg_tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	        <td align="left" class="MC"><p class="R">CELULAR</p><p class="C">
		        <input id="EndEtg_tel_cel" name="EndEtg_tel_cel" class="TA" value="" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fNEW.EndEtg_obs.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Número de celular inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        </tr>
        </table>
	
        
        <!-- ************   ENDEREÇO DE ENTREGA: TELEFONE COMERCIAL   ************ -->
        <table width="649" class="QS Mostrar_EndEtg_pj Habilitar_EndEtg_outroendereco" cellspacing="0">
	        <tr>
	        <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="EndEtg_ddd_com" name="EndEtg_ddd_com" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fNEW.EndEtg_tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	        <td class="MD" align="left"><p class="R">TELEFONE </p><p class="C">
		        <input id="EndEtg_tel_com" name="EndEtg_tel_com" class="TA" value="" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fNEW.EndEtg_ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        <td align="left"><p class="R">RAMAL</p><p class="C">
		        <input id="EndEtg_ramal_com" name="EndEtg_ramal_com" class="TA" value="" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_ddd_com_2.focus(); filtra_numerico();"></p></td>
	        </tr>
	        <tr>
	            <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
	            <input id="EndEtg_ddd_com_2" name="EndEtg_ddd_com_2" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fNEW.EndEtg_tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!!');this.focus();}" /></p>  
	            </td>
	            <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
	            <input id="EndEtg_tel_com_2" name="EndEtg_tel_com_2" class="TA" value="" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fNEW.EndEtg_ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	            </td>
	            <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
	            <input id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" class="TA" value="" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_obs.focus(); filtra_numerico();" /></p>
	            </td>
	        </tr>
        </table>

    <% end if %>
<% end if %>


<!-- ************   JUSTIFIQUE O ENDEREÇO   ************ -->
<table  id="obs_endereco" width="649" class="QS" cellspacing="0">
	<tr>
	<td class="M" width="50%" align="left"><p class="R">JUSTIFIQUE O ENDEREÇO</p><p class="C">
		<select id="EndEtg_obs" name="EndEtg_obs" style="margin-right:225px;">			
			 <%=codigo_descricao_monta_itens_select_por_loja(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA, "", loja)%>
		</select></p></td>
	</tr>
</table>

	</form>
<% end if %>

<% if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then %>
<form id="fMAG" name="fMAG">
<input type="hidden" name="c_mag_customer_full_name" id="c_mag_customer_full_name" value="<%=c_mag_customer_full_name%>" />
<input type="hidden" name="c_mag_customer_dob" id="c_mag_customer_dob" value="<%=c_mag_customer_dob%>" />
<input type="hidden" name="c_mag_customer_email" id="c_mag_customer_email" value="<%=c_mag_customer_email%>" />
<input type="hidden" name="c_mag_email_identificado" id="c_mag_email_identificado" value="<%=c_mag_email_identificado%>" />
<input type="hidden" name="c_mag_end_cob_email" id="c_mag_end_cob_email" value="<%=c_mag_end_cob_email%>" />
<input type="hidden" name="c_mag_end_cob_telephone_ddd" id="c_mag_end_cob_telephone_ddd" value="<%=c_mag_end_cob_telephone_ddd%>" />
<input type="hidden" name="c_mag_end_cob_telephone_numero" id="c_mag_end_cob_telephone_numero" value="<%=c_mag_end_cob_telephone_numero%>" />
<input type="hidden" name="c_mag_end_cob_celular_ddd" id="c_mag_end_cob_celular_ddd" value="<%=c_mag_end_cob_celular_ddd%>" />
<input type="hidden" name="c_mag_end_cob_celular_numero" id="c_mag_end_cob_celular_numero" value="<%=c_mag_end_cob_celular_numero%>" />
<input type="hidden" name="c_mag_end_cob_endereco" id="c_mag_end_cob_endereco" value="<%=c_mag_end_cob_endereco%>" />
<input type="hidden" name="c_mag_end_cob_endereco_numero" id="c_mag_end_cob_endereco_numero" value="<%=c_mag_end_cob_endereco_numero%>" />
<input type="hidden" name="c_mag_end_cob_complemento" id="c_mag_end_cob_complemento" value="<%=c_mag_end_cob_complemento%>" />
<input type="hidden" name="c_mag_end_cob_bairro" id="c_mag_end_cob_bairro" value="<%=c_mag_end_cob_bairro%>" />
<input type="hidden" name="c_mag_end_cob_cidade" id="c_mag_end_cob_cidade" value="<%=c_mag_end_cob_cidade%>" />
<input type="hidden" name="c_mag_end_cob_uf" id="c_mag_end_cob_uf" value="<%=c_mag_end_cob_uf%>" />
<input type="hidden" name="c_mag_end_cob_cep" id="c_mag_end_cob_cep" value="<%=c_mag_end_cob_cep%>" />
<input type="hidden" name="c_mag_end_etg_email" id="c_mag_end_etg_email" value="<%=c_mag_end_etg_email%>" />
<input type="hidden" name="c_mag_end_etg_telephone_ddd" id="c_mag_end_etg_telephone_ddd" value="<%=c_mag_end_etg_telephone_ddd%>" />
<input type="hidden" name="c_mag_end_etg_telephone_numero" id="c_mag_end_etg_telephone_numero" value="<%=c_mag_end_etg_telephone_numero%>" />
<input type="hidden" name="c_mag_end_etg_celular_ddd" id="c_mag_end_etg_celular_ddd" value="<%=c_mag_end_etg_celular_ddd%>" />
<input type="hidden" name="c_mag_end_etg_celular_numero" id="c_mag_end_etg_celular_numero" value="<%=c_mag_end_etg_celular_numero%>" />
<input type="hidden" name="c_mag_end_etg_endereco" id="c_mag_end_etg_endereco" value="<%=c_mag_end_etg_endereco%>" />
<input type="hidden" name="c_mag_end_etg_endereco_numero" id="c_mag_end_etg_endereco_numero" value="<%=c_mag_end_etg_endereco_numero%>" />
<input type="hidden" name="c_mag_end_etg_complemento" id="c_mag_end_etg_complemento" value="<%=c_mag_end_etg_complemento%>" />
<input type="hidden" name="c_mag_end_etg_bairro" id="c_mag_end_etg_bairro" value="<%=c_mag_end_etg_bairro%>" />
<input type="hidden" name="c_mag_end_etg_cidade" id="c_mag_end_etg_cidade" value="<%=c_mag_end_etg_cidade%>" />
<input type="hidden" name="c_mag_end_etg_uf" id="c_mag_end_etg_uf" value="<%=c_mag_end_etg_uf%>" />
<input type="hidden" name="c_mag_end_etg_cep" id="c_mag_end_etg_cep" value="<%=c_mag_end_etg_cep%>" />
</form>
<% end if %>

<p class="TracoBottom"></p>

<table width="649" cellspacing="0">
<tr>
	<% if exibir_botao_novo_item then s="'left'" else s="'center'" %>
	<td align=<%=s%>>
		<%
			s="cliente.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
			if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
		%>
		<div name="dVOLTAR" id="dVOLTAR">
			<a href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
		</div>
	</td>

<% if exibir_botao_novo_item then %>
	<td align="right"><div name="dPEDIDO" id="dPEDIDO">
		<a name="bPEDIDO" id="bPEDIDO" href="javascript:fNEWConcluir(fNEW);" title="cadastra um novo pedido para este cliente">
		<img src="../botao/pedido.gif" width="176" height="55" border="0"></a></div>
	</td>
<% end if %>

</tr>
</table>

</center>
</body>


<% if (pagina_retorno <> "") And exibir_botao_novo_item then %>
	<%
		strScript = _
			"<script language='JavaScript'>" & chr(13) & _
			"dVOLTAR.style.visibility='hidden';" & chr(13) & _
			"dPEDIDO.style.visibility='hidden';" & chr(13) & _
			"window.status = 'Aguarde, carregando página ...';" & chr(13) & _
			"setTimeout(" & chr(34) & "window.location='" & pagina_retorno & "'" & chr(34) & ", 100);" & chr(13) & _
			"</script>" & chr(13)
		Response.Write strScript
	%>
<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		if tMAP_END_ETG.State <> 0 then tMAP_END_ETG.Close
		set tMAP_END_ETG = nothing

		if tMAP_END_COB.State <> 0 then tMAP_END_COB.Close
		set tMAP_END_COB = nothing

		if tMAP_XML.State <> 0 then tMAP_XML.Close
		set tMAP_XML = nothing
		end if

	cn.Close
	set cn = nothing
%>