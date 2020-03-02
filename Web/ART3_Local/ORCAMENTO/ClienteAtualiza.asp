<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->
<%
'     =====================================
'	  C L I E N T E A T U A L I Z A . A S P
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
	dim s, s_aux, usuario, loja, alerta, exibir_botao_novo_item, s_dest
	exibir_botao_novo_item = False
	
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim blnLojaHabilitadaProdCompostoECommerce
	blnLojaHabilitadaProdCompostoECommerce = isLojaHabilitadaProdCompostoECommerce(loja)

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
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, cliente_selecionado, cnpj_cpf_selecionado, s_nome, s_ie, s_rg, s_sexo
	dim s_contribuinte_icms, s_produtor_rural, s_contribuinte_icms_cadastrado, s_produtor_rural_cadastrado
	dim s_endereco, s_endereco_numero, s_endereco_complemento, s_bairro, s_cidade, s_uf, s_cep
	dim s_ddd_res, s_tel_res, s_ddd_com, s_tel_com, s_ramal_com, s_contato, s_dt_nasc, s_filiacao, s_obs_crediticias, s_midia, s_email, s_email_xml
	dim eh_cpf
	dim pagina_retorno
	dim s_tel_com_2, s_ddd_com_2, s_tel_cel, s_ddd_cel, s_ramal_com_2
	
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
	pagina_retorno = Trim(request("pagina_retorno"))
	s_tel_com_2=retorna_so_digitos(Trim(request("tel_com_2")))
	s_ddd_com_2=retorna_so_digitos(Trim(request("ddd_com_2")))
	s_tel_cel=retorna_so_digitos(Trim(request("tel_cel")))
	s_ddd_cel=retorna_so_digitos(Trim(request("ddd_cel")))
	s_ramal_com_2=retorna_so_digitos(Trim(request("ramal_com_2")))
	
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

	if operacao_selecionada = OP_INCLUI then 'operacao_selecionada
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
		elseif eh_cpf And (s_tel_res="") And (s_tel_com="") And (s_tel_cel="") then
			alerta="PREENCHA PELO MENOS UM TELEFONE."
		elseif (Not eh_cpf) And (s_tel_com="") And (s_tel_com_2="") then
			alerta="PREENCHA O TELEFONE."
		elseif (s_ie="") And (s_contribuinte_icms = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
			alerta="PREENCHA A INSCRIÇÃO ESTADUAL."
'		elseif s_midia="" then
'			alerta="INDIQUE A FORMA PELA QUAL CONHECEU A BONSHOP."
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
			if s_ie <> "" then
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
		end if 'operacao_selecionada

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
		
		if s_email <> "" then
		'	CONSISTÊNCIA DESATIVADA TEMPORARIAMENTE
'			if Not email_AF_ok(s_email, s_cnpj_cpf, msg_erro_aux) then
'				alerta=texto_add_br(alerta)
'				alerta=alerta & "Endereço de email (" & s_email & ") não é válido!!<br />" & msg_erro_aux
'				end if
			end if
		end if

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
	
	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	dim msg_erro, msg_erro_aux
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_INCLUI
		'	 =========
			if alerta = "" then 
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				s = "SELECT * FROM t_CLIENTE WHERE id = '" & cliente_selecionado & "'"
				r.Open s, cn
				if r.EOF then 
					r.AddNew 
					criou_novo_reg_cliente = True
					r("id")=cliente_selecionado
					r("dt_cadastro") = Date
					r("usuario_cadastro") = usuario
				'	O ORÇAMENTISTA É O INDICADOR
					r("indicador") = usuario
					r("sistema_responsavel_cadastro") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
					r("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
				else
					alerta = "REGISTRO COM ID=" & cliente_selecionado & " JÁ EXISTE."
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					end if
				end if 'if alerta = ""
			
			if alerta = "" then
				r("cnpj_cpf")=cnpj_cpf_selecionado
				if eh_cpf then s=ID_PF else s=ID_PJ
				r("tipo")=s
				r("ie")=s_ie
				r("rg")=s_rg
				r("nome")=s_nome
				r("sexo")=s_sexo
				If s_contribuinte_icms <> s_contribuinte_icms_cadastrado Then
					r("contribuinte_icms_status")=CInt(s_contribuinte_icms)
					r("contribuinte_icms_data")=Now
					r("contribuinte_icms_data_hora")=Now
					r("contribuinte_icms_usuario")=usuario
					End If
				If s_produtor_rural <> s_produtor_rural_cadastrado Then
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

				r.Update

				If Err = 0 then
				'	PREPARA O LOG 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg_cliente then
						s_log = log_via_vetor_monta_inclusao(vLog2)
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
						if s_log <> "" then grava_log usuario, loja, "", cliente_selecionado, OP_LOG_CLIENTE_INCLUSAO, s_log
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
				end if  'if alerta = "" then
		
		
		case OP_CONSULTA
		'	 ===========
			if alerta = "" then 
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				s = "SELECT * FROM t_CLIENTE WHERE id = '" & cliente_selecionado & "'"
				r.Open s, cn
				if r.EOF then 
					alerta = "REGISTRO COM ID=" & cliente_selecionado & " NÃO EXISTE."
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
						end if
				log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
				end if 'if alerta = ""
			
			if alerta = "" then
				r("ie")=s_ie
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
				r("dt_ult_atualizacao")=Now
				r("usuario_ult_atualizacao")=usuario

				r("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP

				r.Update

				If Err = 0 then
				'	PREPARA O LOG 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
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
					
				if Not erro_fatal then
				'	GRAVA O LOG
					if s_log <> "" then grava_log usuario, loja, "", cliente_selecionado, OP_LOG_CLIENTE_ALTERACAO, s_log
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
				end if  'if alerta = "" then
		
		
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
	<title><%=TITULO_JANELA_MODULO_ORCAMENTO%></title>
	</head>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var fCepPopup;

$(function () {
	var f;
	if ((typeof (fORC) !== "undefined") && (fORC !== null)) {
		f = fORC;

		if (!f.rb_end_entrega[1].checked) {
			f.EndEtg_endereco.disabled = true;
			f.EndEtg_endereco_numero.disabled = true;
			f.EndEtg_bairro.disabled = true;
			f.EndEtg_cidade.disabled = true;
			f.EndEtg_obs.disabled = true;
			f.EndEtg_uf.disabled = true;
			f.EndEtg_cep.disabled = true;
			f.bPesqCepEndEtgNovo.disabled = true;
			f.EndEtg_endereco_complemento.disabled = true;
		}

		if (trim(fORC.c_FormFieldValues.value) != "") {
			stringToForm(fORC.c_FormFieldValues.value, $('#fORC'));
		}
	}
});
function Disabled_True(f) {

    f.EndEtg_endereco.disabled = true;
    f.EndEtg_endereco_numero.disabled = true;
    f.EndEtg_bairro.disabled = true;
    f.EndEtg_cidade.disabled = true;
    f.EndEtg_obs.disabled = true;
    f.EndEtg_uf.disabled = true;
    f.EndEtg_cep.disabled = true;
    f.bPesqCepEndEtgNovo.disabled = true;
    f.EndEtg_endereco_complemento.disabled = true;
}
function Disabled_False(f) {

    f.EndEtg_endereco.disabled = false;
    f.EndEtg_endereco_numero.disabled = false;
    f.EndEtg_bairro.disabled = false;
    f.EndEtg_cidade.disabled = false;
    f.EndEtg_obs.disabled = false;
    f.EndEtg_uf.disabled = false;
    f.EndEtg_cep.disabled = false;
    f.bPesqCepEndEtgNovo.disabled = false;
    f.EndEtg_endereco_complemento.disabled = false;
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
	f=fORC;
	ProcessaSelecaoCEP=TrataCepEnderecoEntrega;
	strUrl="../Global/AjaxCepPesqPopup.asp";
	if (trim(f.EndEtg_cep.value)!="") strUrl=strUrl+"?CepDefault="+trim(f.EndEtg_cep.value);
	fCepPopup=window.open(strUrl, "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=980,height=650,left=0,top=0");
	fCepPopup.focus();
}

function TrataCepEnderecoEntrega(strCep, strUF, strLocalidade, strBairro, strLogradouro, strEnderecoNumero, strEnderecoComplemento) {
var f;
	f=fORC;
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

function fORCConcluir( f ){
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
		}

	fORC.c_FormFieldValues.value = formToString($("#fORC"));

	dORCAMENTO.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit(); 
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
		s = "<p style='margin:5px 2px 5px 2px;'>" & alerta & "</p>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "CLIENTE " & cnpj_cpf_formata(cnpj_cpf_selecionado) & " CADASTRADO COM SUCESSO."
				exibir_botao_novo_item = True
			case OP_CONSULTA
				s = "CLIENTE " & cnpj_cpf_formata(cnpj_cpf_selecionado) & " ATUALIZADO COM SUCESSO."
				exibir_botao_novo_item = True
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
<br><br>


<!-- ************   FORM PARA OPÇÃO DE CADASTRAR NOVO ORÇAMENTO?  ************ -->
<% if exibir_botao_novo_item then %>
	<% if blnLojaHabilitadaProdCompostoECommerce then
			s_dest = "OrcamentoNovoProdCompostoMask.asp"
		else
			s_dest = "OrcamentoNovo.asp"
		end if %>
	<form action="<%=s_dest%>" method="post" id="fORC" name="fORC" onsubmit="if (!fORCConcluir(fORC)) return false">
	<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=cliente_selecionado%>'>
	<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_INCLUI%>'>
	<input type="hidden" name="midia_selecionada" id="midia_selecionada" value='<%=s_midia%>'>
	<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />


<!-- ************   ENDEREÇO DE ENTREGA: S/N   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td align="left">
		<p class="R">ENDEREÇO DE ENTREGA</p><p class="C">
			<% intIdx = 0 %>
			<input type="radio" id="rb_end_entrega" name="rb_end_entrega" value="N" onclick="Disabled_True(fORC);"><span class="C" style="cursor:default" onclick="fORC.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_True(fORC);">O mesmo endereço do cadastro</span>
			<% intIdx = intIdx + 1 %>
			<br><input type="radio" id="rb_end_entrega" name="rb_end_entrega" value="S" onclick="Disabled_False(fORC);"><span class="C" style="cursor:default" onclick="fORC.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_False(fORC);">Outro endereço</span>
		</p>
		</td>
	</tr>
</table>

<!-- ************   ENDEREÇO DE ENTREGA: ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C">
		<input id="EndEtg_endereco" name="EndEtg_endereco" class="TA" value="" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO DE ENTREGA: Nº/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">Nº</p><p class="C">
		<input id="EndEtg_endereco_numero" name="EndEtg_endereco_numero" class="TA" value="" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td width="50%" align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<input id="EndEtg_endereco_complemento" name="EndEtg_endereco_complemento" class="TA" value="" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO DE ENTREGA: BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">BAIRRO</p><p class="C">
		<input id="EndEtg_bairro" name="EndEtg_bairro" class="TA" value="" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_cidade.focus(); filtra_nome_identificador();"></p></td>
	<td width="50%" align="left"><p class="R">CIDADE</p><p class="C">
		<input id="EndEtg_cidade" name="EndEtg_cidade" class="TA" value="" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO DE ENTREGA: UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="50%" class="MD" align="left"><p class="R">UF</p><p class="C">
		<input id="EndEtg_uf" name="EndEtg_uf" class="TA" value="" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fORC.EndEtg_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
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
<!-- ************   JUSTIFIQUE O ENDEREÇO   ************ -->
<table  id="obs_endereco" width="649" class="QS" cellspacing="0">
	<tr>
	<td class="M" width="50%" align="left"><p class="R">JUSTIFIQUE O ENDEREÇO</p><p class="C">
		<select id="EndEtg_obs" name="EndEtg_obs" style="margin-right:225px;">			
			 <%=codigo_descricao_monta_itens_select_por_loja(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA, "", loja)%>
		</select></td>
	</tr>
</table>
	</form>
<% end if %>

<p class="TracoBottom"></p>

<table width="649" cellspacing="0">
<tr>
	<% if exibir_botao_novo_item then s="'left'" else s="'center'" %>
	<td align=<%=s%>>
		<%
			s="resumo.asp"
			if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
		%>
		<div name="dVOLTAR" id="dVOLTAR">
			<a href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
		</div>
	</td>

<% if exibir_botao_novo_item then %>
	<td align="center"><div name="dORCAMENTO" id="dORCAMENTO">
		<a name="bORCAMENTO" id="bORCAMENTO" href="javascript:fORCConcluir(fORC);" title="cadastra um novo orçamento para este cliente">
		<img src="../botao/orcamento.gif" width="176" height="55" border="0"></a></div>
	</td>
<% end if %>

</tr>
</table>

</center>
</body>


<% if (pagina_retorno <> "") And exibir_botao_novo_item then %>
	<script language="JavaScript" type="text/javascript">
		dVOLTAR.style.visibility="hidden";
		dORCAMENTO.style.visibility="hidden";
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