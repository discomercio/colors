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
	
	dim intCounter
	dim s, s_aux, usuario, alerta
	
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim criou_novo_reg_aux
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
	dim operacao_selecionada
	dim cliente_selecionado, cnpj_cpf_selecionado, s_nome, s_ie, s_rg, s_sexo
	dim s_contribuinte_icms, s_produtor_rural, s_contribuinte_icms_cadastrado, s_produtor_rural_cadastrado
	dim s_endereco, s_endereco_numero, s_endereco_complemento, s_bairro, s_cidade, s_uf, s_cep
	dim s_ddd_res, s_tel_res, s_ddd_com, s_tel_com, s_ramal_com, s_contato, s_dt_nasc, s_filiacao, s_obs_crediticias, s_midia, s_email, s_email_xml
	dim s_indicador, strCampoIndicadorEditavel
	dim eh_cpf
	dim pagina_retorno
	dim s_tel_com_2, s_ddd_com_2, s_tel_cel, s_ddd_cel, s_ramal_com_2
	
	operacao_selecionada=Trim(request("operacao_selecionada"))
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
	s_indicador=Trim(request("indicador"))
	strCampoIndicadorEditavel=Trim(request("CampoIndicadorEditavel"))
	s_tel_com_2=retorna_so_digitos(Trim(request("tel_com_2")))
	s_ddd_com_2=retorna_so_digitos(Trim(request("ddd_com_2")))
	s_tel_cel=retorna_so_digitos(Trim(request("tel_cel")))
	s_ddd_cel=retorna_so_digitos(Trim(request("ddd_cel")))
	s_ramal_com_2=retorna_so_digitos(Trim(request("ramal_com_2")))

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
	
	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	eh_cpf=(len(cnpj_cpf_selecionado)=11)

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
	elseif eh_cpf And (Not sexo_ok(s_sexo)) then
		alerta="INDIQUE QUAL O SEXO."
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
	elseif s_cep="" then
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
    elseif eh_cpf And (s_ddd_cel = "") And ((s_tel_cel <> "")) then
		alerta="PREENCHA O DDD."
    elseif eh_cpf And Not ddd_ok(s_ddd_cel) then
        alerta="DDD DO CELULAR INVÁLIDO."
    elseif eh_cpf And Len(retorna_so_digitos(s_tel_cel)) > 9 then
        alerta="NÚMERO DO CELULAR INVÁLIDO."
	elseif eh_cpf And (s_tel_res="") And (s_tel_com="") And (s_tel_cel="") then
		alerta="PREENCHA PELO MENOS UM TELEFONE."
	elseif (Not eh_cpf) And (s_tel_com="") And (s_tel_com_2="") then
		alerta="PREENCHA O TELEFONE."
	elseif (s_ie="") And (s_contribuinte_icms = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
		alerta="PREENCHA A INSCRIÇÃO ESTADUAL."
'	elseif s_midia="" then
'		alerta="INDIQUE A FORMA PELA QUAL CONHECEU A BONSHOP."
	elseif (s_contribuinte_icms = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) And (s_ie="") then
		alerta="SE CLIENTE É CONTRIBUINTE DO ICMS A INSCRIÇÃO ESTADUAL DEVE SER PREENCHIDA."
		end if

'	CONSISTÊNCIAS P/ EMISSÃO DE NFe
	dim s_tabela_municipios_IBGE
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
	'	I.E. É VÁLIDA?
		if (Not eh_cpf) And (s_ie<>"") then
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
												"		<td>" & chr(13) & _
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
			if Not operacao_permitida(OP_CEN_EDITA_CLIENTE_DADOS_CADASTRAIS, s_lista_operacoes_permitidas) then
				alerta = "Nível de acesso insuficiente para realizar esta operação."
				end if
			end if
		end if
	
	dim s_cnpj_cpf
	dim r_cliente
	dim blnConsistirEmailAF
    dim blnVerificarTel
	blnConsistirEmailAF = False
	if alerta = "" then
		if operacao_selecionada = OP_INCLUI then
			s_cnpj_cpf = cnpj_cpf_selecionado
			blnConsistirEmailAF = True
		else
			set r_cliente = New cl_CLIENTE
			call x_cliente_bd(cliente_selecionado, r_cliente)
			s_cnpj_cpf = r_cliente.cnpj_cpf 
			if ucase(s_email) <> ucase(Trim(r_cliente.email)) then blnConsistirEmailAF = True
			end if
		
		if blnConsistirEmailAF And (s_email <> "") then
		'	CONSISTÊNCIA DESATIVADA TEMPORARIAMENTE
'			if Not email_AF_ok(s_email, s_cnpj_cpf, msg_erro_aux) then
'				alerta=texto_add_br(alerta)
'				alerta=alerta & "Endereço de email (" & s_email & ") não é válido!!<br />" & msg_erro_aux
'				end if
			end if
		end if

	if alerta = "" then
        ' VERIFICA A DISPONIBILIDADE DO USO DO TELEFONE NO CADASTRO
        blnVerificarTel = False
        if operacao_selecionada = OP_INCLUI then
			blnVerificarTel = True
		else
			if s_tel_res <> "" And (s_ddd_res<>r_cliente.ddd_res Or s_tel_res<>r_cliente.tel_res) then blnVerificarTel = true
			end if
        if blnVerificarTel then
            if s_tel_res <> "" then
                if (Cstr(s_ddd_res & s_tel_res) = TEL_BONSHOP_1) Or (Cstr(s_ddd_res & s_tel_res) = TEL_BONSHOP_2) Or (Cstr(s_ddd_res & s_tel_res) = TEL_BONSHOP_3) then
                    alerta="NÃO É PERMITIDO UTILIZAR TELEFONES DA BONSHOP NO CADASTRO DE CLIENTES."
                elseif verifica_telefones_repetidos(s_ddd_res, s_tel_res, s_cnpj_cpf) > NUM_MAXIMO_TELEFONES_REPETIDOS_CAD_CLIENTES then
                    alerta="TELEFONE RESIDENCIAL (" & s_ddd_res & ") " & s_tel_res & " JÁ ESTÁ SENDO UTILIZADO NO CADASTRO DE OUTROS CLIENTES. <br>Não foi possível concluir o cadastro."
                end if
            end if
        end if
	end if

	if alerta = "" then
        blnVerificarTel = False
		if operacao_selecionada = OP_INCLUI then
			blnVerificarTel = True
		else
			if s_tel_com <> "" And (s_ddd_com<>r_cliente.ddd_com Or s_tel_com<>r_cliente.tel_com) then blnVerificarTel = true
			end if
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
        if blnVerificarTel then
            if s_tel_cel <> "" then
                if (Cstr(s_ddd_cel & s_tel_cel) = TEL_BONSHOP_1) Or (Cstr(s_ddd_cel & s_tel_cel) = TEL_BONSHOP_2) Or (Cstr(s_ddd_cel & s_tel_cel) = TEL_BONSHOP_3) then
                    alerta="NÃO É PERMITIDO UTILIZAR TELEFONES DA BONSHOP NO CADASTRO DE CLIENTES."
                elseif verifica_telefones_repetidos(s_ddd_cel, s_tel_cel, s_cnpj_cpf) > NUM_MAXIMO_TELEFONES_REPETIDOS_CAD_CLIENTES then
                    alerta="TELEFONE CELULAR (" & s_ddd_cel & ") " & s_tel_cel & " JÁ ESTÁ SENDO UTILIZADO NO CADASTRO DE OUTROS CLIENTES. <br>Não foi possível concluir o cadastro."
                end if
            end if
        end if
	end if

	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	dim msg_erro, msg_erro_aux
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

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

'	ATUALIZA CADASTRO DO CLIENTE NO BD
	if alerta = "" then 
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		s = "SELECT * FROM t_CLIENTE WHERE id = '" & cliente_selecionado & "'"
		r.Open s, cn
		if r.EOF then 
			erro_fatal=True
			alerta = "CLIENTE (ID=" & cliente_selecionado & ") NÃO FOI ENCONTRADO NO BANCO DE DADOS."
			end if
	
		if Not erro_fatal then
			s_cep_original = Trim("" & r("cep"))
			s_cep_novo = s_cep
			
			log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
			r("cnpj_cpf")=cnpj_cpf_selecionado
			if eh_cpf then s=ID_PF else s=ID_PJ
			r("tipo")=s
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
			If (Trim(s_produtor_rural) <> "") And (s_produtor_rural <> s_produtor_rural_cadastrado) Then
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
				s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
				if s_log <> "" then 
					s_log="id=" & Trim("" & r("id")) & "; " & s_log
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
				if s_log <> "" then grava_log usuario, "", "", cliente_selecionado, OP_LOG_CLIENTE_ALTERACAO, s_log

				if blnHaPedidoAprovadoComEntregaPendente And (s_log <> "") then
					if (Instr(s_log, "endereco") <> 0) Or (Instr(s_log, "bairro") <> 0) Or (Instr(s_log, "cidade") <> 0) Or (Instr(s_log, "uf") <> 0) Or (Instr(s_log, "cep") <> 0) Or (Instr(s_log, "endereco_numero") <> 0) Or (Instr(s_log, "endereco_complemento") <> 0) then
						'Envia alerta de que houve edição no cadastro de cliente que possui pedido com status de análise de crédito 'crédito ok' e com entrega pendente
						dim rEmailDestinatario
						dim corpo_mensagem, id_email, msg_erro_grava_email
						set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaEdicaoCadastroClienteComPedidoCreditoOkEntregaPendente)
						if Trim("" & rEmailDestinatario.campo_texto) <> "" then
							s_log_aux = substitui_caracteres(s_log, ";", vbCrLf)
							corpo_mensagem = "O usuário '" & usuario & "' editou em " & formata_data_hora_sem_seg(Now) & " na Central o cadastro do cliente:" & vbCrLf & _
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
			
			if Not erro_fatal then
			'	PROCESSA SELEÇÃO AUTOMÁTICA DE TRANSPORTADORA BASEADO NO CEP
				s_cep_original = retorna_so_digitos(s_cep_original)
				s_cep_novo = retorna_so_digitos(s_cep_novo)
				if s_cep_original <> s_cep_novo then
					s_log_transp_auto = ""
					s_transp_id_auto_novo = ""
					if s_cep_novo <> "" then s_transp_id_auto_novo = obtem_transportadora_pelo_cep(s_cep_novo)
					
				'	SE O CEP MUDOU, VERIFICA PEDIDOS CUJA ENTREGA SERÁ NO ENDEREÇO DE CADASTRO DO CLIENTE
				'	RESTRIÇÕES:
				'		st_end_entrega = 0  ->  NÃO TEM ENDEREÇO DE ENTREGA
				'		Len(LTrim(RTrim(Coalesce(obs_2,'')))) = 0  ->  NF AINDA NÃO EMITIDA
					s = "SELECT " & _
							"*" & _
						" FROM t_PEDIDO" & _
						" WHERE" & _
							" (id_cliente = '" & cliente_selecionado & "')" & _
							" AND (st_entrega <> '" & ST_ENTREGA_ENTREGUE & "')" & _
							" AND (st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
							" AND (analise_credito <> " & COD_AN_CREDITO_OK & ")" & _
							" AND (st_end_entrega = 0)" & _
							" AND (Len(LTrim(RTrim(Coalesce(obs_2,'')))) = 0)" & _
							" AND (transportadora_selecao_auto_status = " & TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_S & ")" & _
							" AND (transportadora_selecao_auto_tipo_endereco = " & TRANSPORTADORA_SELECAO_AUTO_TIPO_ENDERECO_CLIENTE & ")" & _
							" AND (LTrim(RTrim(Coalesce(transportadora_id,''))) <> '" & s_transp_id_auto_novo & "')" & _
						" ORDER BY" & _
							" data_hora"
					if r.State <> 0 then r.Close
					r.Open s, cn
					do while Not r.EOF
						if Ucase(Trim("" & r("transportadora_id"))) <> Ucase(s_transp_id_auto_novo) then
							if s_log_transp_auto <> "" then s_log_transp_auto = s_log_transp_auto & "; "
							s_log_transp_auto = s_log_transp_auto & "Pedido " & Trim("" & r("pedido")) & ": '" & Trim("" & r("transportadora_id")) & "' => '" & s_transp_id_auto_novo & "'"
							r("transportadora_id") = s_transp_id_auto_novo
							r("transportadora_data") = Now
							r("transportadora_usuario") = usuario
							r("transportadora_selecao_auto_status") = TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_S
							r("transportadora_selecao_auto_cep") = s_cep_novo
							r("transportadora_selecao_auto_transportadora") = s_transp_id_auto_novo
							r("transportadora_selecao_auto_tipo_endereco") = TRANSPORTADORA_SELECAO_AUTO_TIPO_ENDERECO_CLIENTE
							r("transportadora_selecao_auto_data_hora") = Now
							r.Update
							end if
						r.MoveNext
						loop
					
					r.Close
					set r = nothing
					if Not cria_recordset_otimista(r, msg_erro) then 
						erro_fatal=True
						alerta = "FALHA AO CRIAR RECORDSET"
						end if
					
					if s_log_transp_auto <> "" then
						s_log_transp_auto = "Alteração da transportadora cadastrada de modo automático no pedido devido à alteração do CEP (" & cep_formata(s_cep_original) & " => " & cep_formata(s_cep_novo) & ") no cadastro do cliente (id: " & cliente_selecionado & ", CNPJ/CPF: " & cnpj_cpf_formata(cnpj_cpf_selecionado) & "): " & " " & s_log_transp_auto
						grava_log usuario, "", "", cliente_selecionado, OP_LOG_CLIENTE_ALTERACAO, s_log_transp_auto
						end if
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


<body>
<center>
<br>

<!--  T E L A  -->

<p class="T">A V I S O</p>

<% 
	s = ""
	s_aux = "'MtAviso'"
	if alerta <> "" then
		s = "<P style='margin:5px 2px 5px 2px;'>" & alerta & "</P>"
		s_aux = "'MtAlerta'"
	else
		s = "CLIENTE " & cnpj_cpf_formata(cnpj_cpf_selecionado) & " ALTERADO COM SUCESSO."
		s = "<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<% if alerta = "" then %>
<div class=<%=s_aux%> style="width:400px;FONT-WEIGHT:bold;" align="CENTER"><%=s%></div>
<% else %>
<div class=<%=s_aux%> style="width:649px;FONT-WEIGHT:bold;" align="CENTER"><%=s%></div>
	<% if s_tabela_municipios_IBGE <> "" then %>
		<br /><br />
		<%=s_tabela_municipios_IBGE%>
	<% end if %>
<% end if %>
<BR><BR>


<p class="TracoBottom"></p>

<table width="649" cellSpacing="0">
<tr>
	<td align='CENTER'>
		<div name="dVOLTAR" id="dVOLTAR">
		<%
			s="resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
			if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
		%>
			<a href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
		</div>
	</td>
</tr>
</table>

</center>
</body>


<% if (pagina_retorno <> "") And (Not erro_fatal) And (Not erro_consistencia) then %>
	<script language="JavaScript" type="text/javascript">
		dVOLTAR.style.visibility="hidden";
		window.status = "Aguarde, carregando página ...";
		setTimeout("window.location='<%=pagina_retorno%>'", 1000);
	</script>
<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>