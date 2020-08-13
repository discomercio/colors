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
'			I N I C I A L I Z A     P � G I N A     A S P     N O     S E R V I D O R
' _____________________________________________________________________________________________


    Const TEL_BONSHOP_1 = "1139344400"
    Const TEL_BONSHOP_2 = "1139344420"
    Const TEL_BONSHOP_3 = "1139344411"

	On Error GoTo 0
	Err.Clear
	
'	EXIBI��O DE BOT�ES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True
	
	
	dim intIdx, intCounter
	dim s, s_aux, usuario, loja, alerta, exibir_botao_novo_item, s_dest
	dim s_tabela_municipios_IBGE
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
	
'	OBT�M DADOS DO FORMUL�RIO ANTERIOR
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

'	DADOS DO S�CIO MAJORIT�RIO
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
		
'	REF BANC�RIA
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
			alerta="CNPJ/CPF N�O FORNECIDO."
		elseif Not cnpj_cpf_ok(cnpj_cpf_selecionado) then
			alerta="CNPJ/CPF INV�LIDO."
		elseif eh_cpf And (Not sexo_ok(s_sexo)) then
			alerta="INDIQUE QUAL O SEXO."
		elseif s_nome = "" then
			if eh_cpf then
				alerta="PREENCHA O NOME DO CLIENTE."
			else
				alerta="PREENCHA A RAZ�O SOCIAL DO CLIENTE."
				end if
		elseif s_endereco = "" then
			alerta="PREENCHA O ENDERE�O."
		elseif Len(s_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
			alerta="ENDERE�O EXCEDE O TAMANHO M�XIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(s_endereco)) & " CARACTERES<br>TAMANHO M�XIMO: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " CARACTERES"
		elseif s_endereco_numero = "" then
			alerta="PREENCHA O N�MERO DO ENDERE�O."
		elseif s_bairro = "" then
			alerta="PREENCHA O BAIRRO."
		elseif s_cidade = "" then
			alerta="PREENCHA A CIDADE."
		elseif (s_uf="") Or (Not uf_ok(s_uf)) then
			alerta="UF INV�LIDA."
		elseif s_cep = "" then
			alerta="INFORME O CEP."
		elseif Not cep_ok(s_cep) then
			alerta="CEP INV�LIDO."
		elseif Not ddd_ok(s_ddd_res) then
			alerta="DDD INV�LIDO."
		elseif Not telefone_ok(s_tel_res) then
			alerta="TELEFONE RESIDENCIAL INV�LIDO."
		elseif (s_ddd_res <> "") And ((s_tel_res = "")) then
			alerta="PREENCHA O TELEFONE RESIDENCIAL."
		elseif (s_ddd_res = "") And ((s_tel_res <> "")) then
			alerta="PREENCHA O DDD."
		elseif Not ddd_ok(s_ddd_com) then
			alerta="DDD INV�LIDO."
		elseif Not telefone_ok(s_tel_com) then
			alerta="TELEFONE COMERCIAL INV�LIDO."
		elseif (s_ddd_com <> "") And ((s_tel_com = "")) then
			alerta="PREENCHA O TELEFONE COMERCIAL."
		elseif (s_ddd_com = "") And ((s_tel_com <> "")) then
			alerta="PREENCHA O DDD."
		elseif eh_cpf And (s_tel_res="") And (s_tel_com="") And (s_tel_cel="") then
			alerta="PREENCHA PELO MENOS UM TELEFONE."
		elseif (Not eh_cpf) And (s_tel_com="") And (s_tel_com_2="") then
			alerta="PREENCHA O TELEFONE."
		elseif (s_ie="") And (s_contribuinte_icms = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
			alerta="PREENCHA A INSCRI��O ESTADUAL."
'		elseif s_midia="" then
'			alerta="INDIQUE A FORMA PELA QUAL CONHECEU A BONSHOP."
			end if


	    if alerta = "" then
		    if (s_produtor_rural = COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) Then
			    if (s_contribuinte_icms <> COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) Or (s_ie = "") then
				    alerta = "Para ser cadastrado como Produtor Rural, � necess�rio ser contribuinte do ICMS e possuir n� de IE"
				    end if
			    end if
		    end if
	
	'	CONSIST�NCIAS P/ EMISS�O DE NFe
		s_tabela_municipios_IBGE = ""
		if alerta = "" then
		'	I.E. � V�LIDA?
			if s_ie <> "" then
				if Not isInscricaoEstadualValida(s_ie, s_uf) then
					alerta="Preencha a IE (Inscri��o Estadual) com um n�mero v�lido!!" & _
							"<br>" & "Certifique-se de que a UF informada corresponde � UF respons�vel pelo registro da IE."
					end if
				end if
		
		'	MUNIC�PIO DE ACORDO C/ TABELA DO IBGE?
			dim s_lista_sugerida_municipios
			dim v_lista_sugerida_municipios
			dim iCounterLista, iNumeracaoLista
			if Not consiste_municipio_IBGE_ok(s_cidade, s_uf, s_lista_sugerida_municipios, msg_erro) then
				if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				if msg_erro <> "" then
					alerta = alerta & msg_erro
				else
					alerta = alerta & "Munic�pio '" & s_cidade & "' n�o consta na rela��o de munic�pios do IBGE para a UF de '" & s_uf & "'!!"
					if s_lista_sugerida_municipios <> "" then
						alerta = alerta & "<br>" & _
										  "Localize o munic�pio na lista abaixo e verifique se a grafia est� correta!!"
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
									"			<p class='N'>" & "Rela��o de munic�pios de '" & s_uf & "' que se iniciam com a letra '" & Ucase(left(s_cidade,1)) & "'" & "</p>" & chr(13) & _
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
				alerta="O CAMPO 'NOME' POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
			elseif Not isTextoValido(s_endereco, s_caracteres_invalidos) then
				alerta="O CAMPO 'ENDERE�O' POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
			elseif Not isTextoValido(s_endereco_numero, s_caracteres_invalidos) then
				alerta="O CAMPO N�MERO DO ENDERE�O POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
			elseif Not isTextoValido(s_endereco_complemento, s_caracteres_invalidos) then
				alerta="O CAMPO 'COMPLEMENTO' POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
			elseif Not isTextoValido(s_bairro, s_caracteres_invalidos) then
				alerta="O CAMPO 'BAIRRO' POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
			elseif Not isTextoValido(s_cidade, s_caracteres_invalidos) then
				alerta="O CAMPO 'CIDADE' POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
			elseif Not isTextoValido(s_contato, s_caracteres_invalidos) then
				alerta="O CAMPO 'CONTATO' POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
			elseif Not isTextoValido(s_filiacao, s_caracteres_invalidos) then
				alerta="O CAMPO 'FILIA��O' POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
			elseif Not isTextoValido(s_obs_crediticias, s_caracteres_invalidos) then
				alerta="O CAMPO 'OBSERVA��ES CREDIT�CIAS' POSSUI UM OU MAIS CARACTERES INV�LIDOS: " & s_caracteres_invalidos
				end if
			end if

		if (alerta="") And (s_dt_nasc<>"") then
			if (DateDiff("m", StrToDate(s_dt_nasc), Date)/12) < 10 then alerta = "DATA DE NASCIMENTO � INV�LIDA."
			end if

		if alerta = "" then
		'	REF BANC�RIA
			for intCounter=Lbound(vRefBancaria) to Ubound(vRefBancaria)
				if vRefBancaria(intCounter).id_cliente <> "" then
					with vRefBancaria(intCounter)
						if Trim(.banco) = "" then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Ref Banc�ria (" & CStr(.ordem) & "): informe o banco."
							end if
						if Trim(.agencia) = "" then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Ref Banc�ria (" & CStr(.ordem) & "): informe a ag�ncia."
							end if
						if Trim(.conta) = "" then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Ref Banc�ria (" & CStr(.ordem) & "): informe o n�mero da conta."
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
		'	DADOS DO S�CIO MAJORIT�RIO
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
						alerta=alerta & "Informe o nome do s�cio majorit�rio."
						end if
					end if
				if blnConsistirDadosBancarios then
					if strSocMajBanco = "" then 
						alerta=texto_add_br(alerta)
						alerta=alerta & "Informe o banco nos dados banc�rios do s�cio majorit�rio."
						end if
					if strSocMajAgencia = "" then 
						alerta=texto_add_br(alerta)
						alerta=alerta & "Informe a ag�ncia nos dados banc�rios do s�cio majorit�rio."
						end if
					if strSocMajConta = "" then 
						alerta=texto_add_br(alerta)
						alerta=alerta & "Informe o n�mero da conta nos dados banc�rios do s�cio majorit�rio."
						end if
					end if
				end if
			end if
		end if 'operacao_selecionada = OP_INCLUI

	if operacao_selecionada = OP_CONSULTA then

	    if alerta = "" then
		    if (s_produtor_rural = COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) Then
			    if (s_contribuinte_icms <> COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) Or (s_ie = "") then
				    alerta = "Para ser cadastrado como Produtor Rural, � necess�rio ser contribuinte do ICMS e possuir n� de IE"
				    end if
			    end if
		    end if
	
	'	CONSIST�NCIAS P/ EMISS�O DE NFe
		if alerta = "" then
		'	I.E. � V�LIDA?
			if s_ie <> "" then
				if Not isInscricaoEstadualValida(s_ie, s_uf) then
					alerta="Preencha a IE (Inscri��o Estadual) com um n�mero v�lido!!" & _
							"<br>" & "Certifique-se de que a UF informada corresponde � UF respons�vel pelo registro da IE."
					end if
				end if
			end if
		end if 'operacao_selecionada = OP_CONSULTA

	dim s_cnpj_cpf
	dim r_cliente
    dim blnVerificarTel
	if operacao_selecionada = OP_INCLUI then
		s_cnpj_cpf = cnpj_cpf_selecionado
	else
		set r_cliente = New cl_CLIENTE
		call x_cliente_bd(cliente_selecionado, r_cliente)
		s_cnpj_cpf = r_cliente.cnpj_cpf
		end if
		
	if alerta = "" then
		if s_email <> "" then
		'	CONSIST�NCIA DESATIVADA TEMPORARIAMENTE
'			if Not email_AF_ok(s_email, s_cnpj_cpf, msg_erro_aux) then
'				alerta=texto_add_br(alerta)
'				alerta=alerta & "Endere�o de email (" & s_email & ") n�o � v�lido!!<br />" & msg_erro_aux
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
                    alerta="N�O � PERMITIDO UTILIZAR TELEFONES DA BONSHOP NO CADASTRO DE CLIENTES."
                elseif verifica_telefones_repetidos(s_ddd_res, s_tel_res, s_cnpj_cpf) > NUM_MAXIMO_TELEFONES_REPETIDOS_CAD_CLIENTES then
                    alerta="TELEFONE RESIDENCIAL (" & s_ddd_res & ") " & s_tel_res & " J� EST� SENDO UTILIZADO NO CADASTRO DE OUTROS CLIENTES. <br>N�o foi poss�vel concluir o cadastro."
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
                    alerta="N�O � PERMITIDO UTILIZAR TELEFONES DA BONSHOP NO CADASTRO DE CLIENTES."
                elseif verifica_telefones_repetidos(s_ddd_com, s_tel_com, s_cnpj_cpf) > NUM_MAXIMO_TELEFONES_REPETIDOS_CAD_CLIENTES then
                    alerta="TELEFONE COMERCIAL (" & s_ddd_com & ") " & s_tel_com & " J� EST� SENDO UTILIZADO NO CADASTRO DE OUTROS CLIENTES. <br>N�o foi poss�vel concluir o cadastro."
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
                    alerta="N�O � PERMITIDO UTILIZAR TELEFONES DA BONSHOP NO CADASTRO DE CLIENTES."
                elseif verifica_telefones_repetidos(s_ddd_com_2, s_tel_com_2, s_cnpj_cpf) > NUM_MAXIMO_TELEFONES_REPETIDOS_CAD_CLIENTES then
                    alerta="TELEFONE COMERCIAL (" & s_ddd_com_2 & ") " & s_tel_com_2 & " J� EST� SENDO UTILIZADO NO CADASTRO DE OUTROS CLIENTES. <br>N�o foi poss�vel concluir o cadastro."
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
                    alerta="N�O � PERMITIDO UTILIZAR TELEFONES DA BONSHOP NO CADASTRO DE CLIENTES."
                elseif verifica_telefones_repetidos(s_ddd_cel, s_tel_cel, s_cnpj_cpf) > NUM_MAXIMO_TELEFONES_REPETIDOS_CAD_CLIENTES then
                    alerta="TELEFONE CELULAR (" & s_ddd_cel & ") " & s_tel_cel & " J� EST� SENDO UTILIZADO NO CADASTRO DE OUTROS CLIENTES. <br>N�o foi poss�vel concluir o cadastro."
                end if
            end if
        end if
	
	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	dim msg_erro, msg_erro_aux
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERA��O NO BD
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
				'	O OR�AMENTISTA � O INDICADOR
					r("indicador") = usuario
					r("sistema_responsavel_cadastro") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
					r("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
				else
					alerta = "REGISTRO COM ID=" & cliente_selecionado & " J� EXISTE."
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
				
			'	REF BANC�RIA
				if blnCadRefBancaria then
					if Not erro_fatal then
						s="UPDATE t_CLIENTE_REF_BANCARIA SET excluido_status=1 WHERE (id_cliente = '" & cliente_selecionado & "')"
						cn.Execute(s)
						If Err <> 0 then 
							erro_fatal=True
							alerta = "FALHA AO PREPARAR ALTERA��O DOS DADOS DE REF BANC�RIA DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
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
										'	CAMPOS DA CHAVE PRIM�RIA
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
												s_log_aux = "Ref Banc�ria inclu�da: " & log_via_vetor_monta_inclusao(vLog2)
											else
												s_log_aux = log_via_vetor_monta_alteracao(vLog1, vLog2)
												if s_log_aux <> "" then 
													s_log_aux="Ref Banc�ria alterada (banco: " & Trim(.banco) & ", ag: " & Trim(.agencia) & ", conta: " & Trim(.conta) & "): " & s_log_aux
													end if
												end if
											
											if s_log_aux <> "" then
												if s_log <> "" then s_log = s_log & "; "
												s_log = s_log & s_log_aux
												end if
										else
											erro_fatal=True
											alerta = "FALHA AO GRAVAR OS DADOS DA REF BANC�RIA (" & Cstr(Err) & ": " & Err.Description & ")."
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
									s_log_aux = s_log_aux & "Ref Banc�ria exclu�da: " & log_via_vetor_monta_exclusao(vLog1)
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
									alerta = "FALHA AO ALTERAR DADOS DE REF BANC�RIA DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
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
							alerta = "FALHA AO PREPARAR ALTERA��O DOS DADOS DE REF PROFISSIONAL DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
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
										'	CAMPOS DA CHAVE PRIM�RIA
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
												s_log_aux = "Ref Profissional inclu�da: " & log_via_vetor_monta_inclusao(vLog2)
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
									s_log_aux = s_log_aux & "Ref Profissional exclu�da: " & log_via_vetor_monta_exclusao(vLog1)
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
							alerta = "FALHA AO PREPARAR ALTERA��O DOS DADOS DE REF COMERCIAL DO CLIENTE (" & Cstr(Err) & ": " & Err.Description & ")."
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
										'	CAMPOS DA CHAVE PRIM�RIA
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
												s_log_aux = "Ref Comercial inclu�da: " & log_via_vetor_monta_inclusao(vLog2)
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
									s_log_aux = s_log_aux & "Ref Comercial exclu�da: " & log_via_vetor_monta_exclusao(vLog1)
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
					alerta = "REGISTRO COM ID=" & cliente_selecionado & " N�O EXISTE."
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
			alerta="OPERA��O INV�LIDA."
			
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
            Disabled_change(f, true);
        }

		if (trim(fORC.c_FormFieldValues.value) != "") {
			stringToForm(fORC.c_FormFieldValues.value, $('#fORC'));
		}
	}

    trataProdutorRuralEndEtg_PF(null);
    trocarEndEtgTipoPessoa(null);
});
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
	//  SE J� HOUVER UMA JANELA DE PESQUISA DE CEP ABERTA, GARANTE QUE ELA SER� FECHADA 
	// E UMA NOVA SER� CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')	
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
	window.status="Conclu�do";
}

function fORCConcluir( f ){
	if ((!f.rb_end_entrega[0].checked)&&(!f.rb_end_entrega[1].checked)) {
		alert('Informe se o endere�o de entrega ser� o mesmo endere�o do cadastro ou n�o!!');
		return;
		}

	if (f.rb_end_entrega[1].checked) {
		if (trim(f.EndEtg_endereco.value)=="") {
			alert('Preencha o endere�o de entrega!!');
			f.EndEtg_endereco.focus();
			return;
			}

		if (trim(f.EndEtg_endereco_numero.value)=="") {
			alert('Preencha o n�mero do endere�o de entrega!!');
			f.EndEtg_endereco_numero.focus();
			return;
			}

		if (trim(f.EndEtg_bairro.value)=="") {
			alert('Preencha o bairro do endere�o de entrega!!');
			f.EndEtg_bairro.focus();
			return;
			}

		if (trim(f.EndEtg_cidade.value)=="") {
			alert('Preencha a cidade do endere�o de entrega!!');
			f.EndEtg_cidade.focus();
			return;
			}
		if (trim(f.EndEtg_obs.value) == "") {
		    alert('Selecione a justificativa do endere�o de entrega!!');
		    f.EndEtg_obs.focus();
		    return;
		    }
		s=trim(f.EndEtg_uf.value);
		if ((s=="")||(!uf_ok(s))) {
			alert('UF inv�lida no endere�o de entrega!!');
			f.EndEtg_uf.focus();
			return;
			}
			
		if (!cep_ok(f.EndEtg_cep.value)) {
			alert('CEP inv�lido no endere�o de entrega!!');
			f.EndEtg_cep.focus();
			return;
			}
<%if blnUsarMemorizacaoCompletaEnderecos then%>
<%if Not eh_cpf then%>
            var EndEtg_tipo_pessoa = $('input[name="EndEtg_tipo_pessoa"]:checked').val();
            if (!EndEtg_tipo_pessoa)
                EndEtg_tipo_pessoa = "";
            if (EndEtg_tipo_pessoa != "PJ" && EndEtg_tipo_pessoa != "PF") {
                alert('Necess�rio escolher Pessoa Jur�dica ou Pessoa F�sica no Endere�o de entrega!!');
                f.EndEtg_tipo_pessoa.focus();
                return;
            }

            if (EndEtg_tipo_pessoa == "PJ") {
                //Campos PJ: 

                if (f.EndEtg_cnpj_cpf_PJ.value == "" || !cnpj_ok(f.EndEtg_cnpj_cpf_PJ.value)) {
                    alert('Endere�o de entrega: CNPJ inv�lido!!');
                    f.EndEtg_cnpj_cpf_PJ.focus();
                    return;
                }

                if ($('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').length == 0) {
                    alert('Endere�o de entrega: informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
                    f.EndEtg_contribuinte_icms_status_PJ.focus();
                    return;
                }

                if ((f.EndEtg_contribuinte_icms_status_PJ[1].checked) && (trim(f.EndEtg_ie_PJ.value) == "")) {
                    alert('Endere�o de entrega: se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
                    f.EndEtg_ie_PJ.focus();
                    return;
                }
                if ((f.EndEtg_contribuinte_icms_status_PJ[0].checked) && (f.EndEtg_ie_PJ.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Endere�o de entrega: se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                    f.EndEtg_ie_PJ.focus();
                    return;
                }
                if ((f.EndEtg_contribuinte_icms_status_PJ[1].checked) && (f.EndEtg_ie_PJ.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Endere�o de entrega: se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                    f.EndEtg_ie_PJ.focus();
                    return;
                }
                if (f.EndEtg_contribuinte_icms_status_PJ[2].checked) {
                    if (f.EndEtg_ie_PJ.value != "") {
                        alert("Endere�o de entrega: se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
                        f.EndEtg_ie_PF.focus();
                        return;
                    }
                }

                if (trim(f.EndEtg_nome.value) == "") {
                    alert('Preencha a raz�o social no endere�o de entrega!!');
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
                    alert('Endere�o de entrega: DDD inv�lido!!');
                    f.EndEtg_ddd_com.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_com.value)) {
                    alert('Endere�o de entrega: telefone inv�lido!!');
                    f.EndEtg_tel_com.focus();
                    return;
                }
                if ((f.EndEtg_ddd_com.value == "") && (f.EndEtg_tel_com.value != "")) {
                    alert('Endere�o de entrega: preencha o DDD do telefone.');
                    f.EndEtg_ddd_com.focus();
                    return;
                }
                if ((f.EndEtg_tel_com.value == "") && (f.EndEtg_ddd_com.value != "")) {
                    alert('Endere�o de entrega: preencha o telefone.');
                    f.EndEtg_tel_com.focus();
                    return;
                }
                if (trim(f.EndEtg_ddd_com.value) == "" && trim(f.EndEtg_ramal_com.value) != "") {
                    alert('Endere�o de entrega: DDD comercial inv�lido!!');
                    f.EndEtg_ddd_com.focus();
                    return;
                }


                if (!ddd_ok(f.EndEtg_ddd_com_2.value)) {
                    alert('Endere�o de entrega: DDD inv�lido!!');
                    f.EndEtg_ddd_com_2.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_com_2.value)) {
                    alert('Endere�o de entrega: telefone inv�lido!!');
                    f.EndEtg_tel_com_2.focus();
                    return;
                }
                if ((f.EndEtg_ddd_com_2.value == "") && (f.EndEtg_tel_com_2.value != "")) {
                    alert('Endere�o de entrega: preencha o DDD do telefone.');
                    f.EndEtg_ddd_com_2.focus();
                    return;
                }
                if ((f.EndEtg_tel_com_2.value == "") && (f.EndEtg_ddd_com_2.value != "")) {
                    alert('Endere�o de entrega: preencha o telefone.');
                    f.EndEtg_tel_com_2.focus();
                    return;
                }
                if (trim(f.EndEtg_ddd_com_2.value) == "" && trim(f.EndEtg_ramal_com_2.value) != "") {
                    alert('Endere�o de entrega: DDD comercial 2 inv�lido!!');
                    f.EndEtg_ddd_com_2.focus();
                    return;
                }

            }
            else {
                //campos PF

                if (f.EndEtg_cnpj_cpf_PF.value == "" || !cpf_ok(f.EndEtg_cnpj_cpf_PF.value)) {
                    alert('Endere�o de entrega: CPF inv�lido!!');
                    f.EndEtg_cnpj_cpf_PF.focus();
                    return;
                }

                if ((!f.EndEtg_produtor_rural_status_PF[0].checked) && (!f.EndEtg_produtor_rural_status_PF[1].checked)) {
                    alert('Endere�o de entrega: informe se o cliente � produtor rural ou n�o!!');
                    return;
                }
                if (!f.EndEtg_produtor_rural_status_PF[0].checked) {
                    if (!f.EndEtg_contribuinte_icms_status_PF[1].checked) {
                        alert('Endere�o de entrega: para ser cadastrado como Produtor Rural, � necess�rio ser contribuinte do ICMS e possuir n� de IE!!');
                        return;
                    }
                    if ((!f.EndEtg_contribuinte_icms_status_PF[0].checked) && (!f.EndEtg_contribuinte_icms_status_PF[1].checked) && (!f.EndEtg_contribuinte_icms_status_PF[2].checked)) {
                        alert('Endere�o de entrega: informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
                        return;
                    }
                    if ((f.EndEtg_contribuinte_icms_status_PF[1].checked) && (trim(f.EndEtg_ie_PF.value) == "")) {
                        alert('Endere�o de entrega: se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
                        f.EndEtg_ie_PF.focus();
                        return;
                    }
                    if ((f.EndEtg_contribuinte_icms_status_PF[0].checked) && (f.EndEtg_ie_PF.value.toUpperCase().indexOf('ISEN') >= 0)) {
                        alert('Endere�o de entrega: se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                        f.EndEtg_ie_PF.focus();
                        return;
                    }
                    if ((f.EndEtg_contribuinte_icms_status_PF[1].checked) && (f.EndEtg_ie_PF.value.toUpperCase().indexOf('ISEN') >= 0)) {
                        alert('Endere�o de entrega: se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                        f.EndEtg_ie_PF.focus();
                        return;
                    }

                    if (f.EndEtg_contribuinte_icms_status_PF[2].checked) {
                        if (f.EndEtg_ie_PF.value != "") {
                            alert("Endere�o de entrega: se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
                            f.EndEtg_ie_PF.focus();
                            return;
                        }
                    }
                }
            

                if (trim(f.EndEtg_nome.value) == "") {
                    alert('Preencha o nome no endere�o de entrega!!');
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
                    alert('Endere�o de entrega: DDD inv�lido!!');
                    f.EndEtg_ddd_res.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_res.value)) {
                    alert('Endere�o de entrega: telefone inv�lido!!');
                    f.EndEtg_tel_res.focus();
                    return;
                }
                if ((trim(f.EndEtg_ddd_res.value) != "") || (trim(f.EndEtg_tel_res.value) != "")) {
                    if (trim(f.EndEtg_ddd_res.value) == "") {
                        alert('Endere�o de entrega: preencha o DDD!!');
                        f.EndEtg_ddd_res.focus();
                        return;
                    }
                    if (trim(f.EndEtg_tel_res.value) == "") {
                        alert('Endere�o de entrega: preencha o telefone!!');
                        f.EndEtg_tel_res.focus();
                        return;
                    }
                }

                if (!ddd_ok(f.EndEtg_ddd_cel.value)) {
                    alert('Endere�o de entrega: DDD inv�lido!!');
                    f.EndEtg_ddd_cel.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_cel.value)) {
                    alert('Endere�o de entrega: telefone inv�lido!!');
                    f.EndEtg_tel_cel.focus();
                    return;
                }
                if ((f.EndEtg_ddd_cel.value == "") && (f.EndEtg_tel_cel.value != "")) {
                    alert('Endere�o de entrega: preencha o DDD do celular.');
                    f.EndEtg_tel_cel.focus();
                    return;
                }
                if ((f.EndEtg_tel_cel.value == "") && (f.EndEtg_ddd_cel.value != "")) {
                    alert('Endere�o de entrega: preencha o n�mero do celular.');
                    f.EndEtg_tel_cel.focus();
                    return;
                }


            }


<%end if%>
<%end if%>
		}


    //campos do endere�o de entrega que precisam de transformacao
    transferirCamposEndEtg(fORC);

	fORC.c_FormFieldValues.value = formToString($("#fORC"));

	dORCAMENTO.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit(); 
}


function transferirCamposEndEtg(formulario) {
<%if blnUsarMemorizacaoCompletaEnderecos then %>
<%if Not eh_cpf then %>
    //Transferimos os dados do endere�o de entrega dos campos certos. 
    //Temos dois conjuntos de campos (para PF e PJ) porque o layout � muito diferente.
    var pj = $('input[name="EndEtg_tipo_pessoa"]:checked').val() == "PJ";
    if (pj) {
        formulario.EndEtg_cnpj_cpf.value = formulario.EndEtg_cnpj_cpf_PJ.value;
        formulario.EndEtg_ie.value = formulario.EndEtg_ie_PJ.value;
        formulario.EndEtg_contribuinte_icms_status.value = $('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').val();
        if (!$('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').val())
            formulario.EndEtg_contribuinte_icms_status.value = "";
    }
    else {
        formulario.EndEtg_cnpj_cpf.value = formulario.EndEtg_cnpj_cpf_PF.value;
        formulario.EndEtg_ie.value = formulario.EndEtg_ie_PF.value;
        formulario.EndEtg_contribuinte_icms_status.value = $('input[name="EndEtg_contribuinte_icms_status_PF"]:checked').val();
        if (!$('input[name="EndEtg_contribuinte_icms_status_PF"]:checked').val())
            formulario.EndEtg_contribuinte_icms_status.value = "";
        formulario.EndEtg_produtor_rural_status.value = $('input[name="EndEtg_produtor_rural_status_PF"]:checked').val();
        if (!$('input[name="EndEtg_produtor_rural_status_PF"]:checked').val())
            formulario.EndEtg_produtor_rural_status.value = "";
    }

    //os campos a mais s�o enviados junto. Deixamos enviar...
<%end if%>
<%end if%>
}

//para mudar o tipo do endere�o de entrega
function trocarEndEtgTipoPessoa(novoTipo) {
<%if blnUsarMemorizacaoCompletaEnderecos then%>
    if (novoTipo && $('input[name="EndEtg_tipo_pessoa"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_tipo_pessoa"]'), novoTipo);

    var pj = $('input[name="EndEtg_tipo_pessoa"]:checked').val() == "PJ";

    if (pj) {
        $(".Mostrar_EndEtg_pf").css("display", "none");
        $(".Mostrar_EndEtg_pj").css("display", "");
        $("#Label_EndEtg_nome").text("RAZ�O SOCIAL");
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
    //ao clicar na op��o Produtor Rural, exibir/ocultar os campos apropriados (endere�o de entrega)
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


<!-- ************   FORM PARA OP��O DE CADASTRAR NOVO OR�AMENTO?  ************ -->
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


<!-- ************   DAODS CADASTRAIS   ************ -->
<%if blnUsarMemorizacaoCompletaEnderecos then%>
    <%
	set r_cliente = New cl_CLIENTE
	call x_cliente_bd(cliente_selecionado, r_cliente)
    %>
    <input type="hidden" name="orcamento_endereco_logradouro" id="orcamento_endereco_logradouro" value="<%=Trim("" & r_cliente.endereco) %>" />
    <input type="hidden" name="orcamento_endereco_bairro" id="orcamento_endereco_bairro" value="<%=Trim("" & r_cliente.bairro) %>" />
    <input type="hidden" name="orcamento_endereco_cidade" id="orcamento_endereco_cidade" value="<%=Trim("" & r_cliente.cidade) %>" />
    <input type="hidden" name="orcamento_endereco_uf" id="orcamento_endereco_uf" value="<%=Trim("" & r_cliente.uf) %>" />
    <input type="hidden" name="orcamento_endereco_cep" id="orcamento_endereco_cep" value="<%=Trim("" & r_cliente.cep) %>" />
    <input type="hidden" name="orcamento_endereco_numero" id="orcamento_endereco_numero" value="<%=Trim("" & r_cliente.endereco_numero) %>" />
    <input type="hidden" name="orcamento_endereco_complemento" id="orcamento_endereco_complemento" value="<%=Trim("" & r_cliente.endereco_complemento) %>" />
    <input type="hidden" name="orcamento_endereco_email" id="orcamento_endereco_email" value="<%=Trim("" & r_cliente.email) %>" />
    <input type="hidden" name="orcamento_endereco_email_xml" id="orcamento_endereco_email_xml" value="<%=Trim("" & r_cliente.email_xml) %>" />
    <input type="hidden" name="orcamento_endereco_nome" id="orcamento_endereco_nome" value="<%=Trim("" & r_cliente.nome) %>" />
    <input type="hidden" name="orcamento_endereco_ddd_res" id="orcamento_endereco_ddd_res" value="<%=Trim("" & r_cliente.ddd_res) %>" />
    <input type="hidden" name="orcamento_endereco_tel_res" id="orcamento_endereco_tel_res" value="<%=Trim("" & r_cliente.tel_res) %>" />
    <input type="hidden" name="orcamento_endereco_ddd_com" id="orcamento_endereco_ddd_com" value="<%=Trim("" & r_cliente.ddd_com) %>" />
    <input type="hidden" name="orcamento_endereco_tel_com" id="orcamento_endereco_tel_com" value="<%=Trim("" & r_cliente.tel_com) %>" />
    <input type="hidden" name="orcamento_endereco_ramal_com" id="orcamento_endereco_ramal_com" value="<%=Trim("" & r_cliente.ramal_com) %>" />
    <input type="hidden" name="orcamento_endereco_ddd_cel" id="orcamento_endereco_ddd_cel" value="<%=Trim("" & r_cliente.ddd_cel) %>" />
    <input type="hidden" name="orcamento_endereco_tel_cel" id="orcamento_endereco_tel_cel" value="<%=Trim("" & r_cliente.tel_cel) %>" />
    <input type="hidden" name="orcamento_endereco_ddd_com_2" id="orcamento_endereco_ddd_com_2" value="<%=Trim("" & r_cliente.ddd_com_2) %>" />
    <input type="hidden" name="orcamento_endereco_tel_com_2" id="orcamento_endereco_tel_com_2" value="<%=Trim("" & r_cliente.tel_com_2) %>" />
    <input type="hidden" name="orcamento_endereco_ramal_com_2" id="orcamento_endereco_ramal_com_2" value="<%=Trim("" & r_cliente.ramal_com_2) %>" />
    <input type="hidden" name="orcamento_endereco_tipo_pessoa" id="orcamento_endereco_tipo_pessoa" value="<%=Trim("" & r_cliente.tipo) %>" />
    <input type="hidden" name="orcamento_endereco_cnpj_cpf" id="orcamento_endereco_cnpj_cpf" value="<%=Trim("" & r_cliente.cnpj_cpf) %>" />
    <input type="hidden" name="orcamento_endereco_contribuinte_icms_status" id="orcamento_endereco_contribuinte_icms_status" value="<%=Trim("" & r_cliente.contribuinte_icms_status) %>" />
    <input type="hidden" name="orcamento_endereco_produtor_rural_status" id="orcamento_endereco_produtor_rural_status" value="<%=Trim("" & r_cliente.produtor_rural_status) %>" />
    <input type="hidden" name="orcamento_endereco_ie" id="orcamento_endereco_ie" value="<%=Trim("" & r_cliente.ie) %>" />
    <input type="hidden" name="orcamento_endereco_rg" id="orcamento_endereco_rg" value="<%=Trim("" & r_cliente.rg) %>" />
    <input type="hidden" name="orcamento_endereco_contato" id="orcamento_endereco_contato" value="<%=Trim("" & r_cliente.contato) %>" />

<%end if%>
        

<!-- ************   ENDERE�O DE ENTREGA: S/N   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td align="left">
		<p class="R">ENDERE�O DE ENTREGA</p><p class="C">
			<% intIdx = 0 %>
			<input type="radio" id="rb_end_entrega_nao" name="rb_end_entrega" value="N" onclick="Disabled_True(fORC);"><span class="C" style="cursor:default" onclick="fORC.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_True(fORC);">O mesmo endere�o do cadastro</span>
			<% intIdx = intIdx + 1 %>
			<br><input type="radio" id="rb_end_entrega_sim" name="rb_end_entrega" value="S" onclick="Disabled_False(fORC);"><span class="C" style="cursor:default" onclick="fORC.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_False(fORC);">Outro endere�o</span>
		</p>
		</td>
	</tr>
</table>



<!--  ************  TIPO DO ENDERE�O DE ENTREGA: PF/PJ (SOMENTE SE O CLIENTE FOR PJ)   ************ -->

<%if blnUsarMemorizacaoCompletaEnderecos then%>
    <%if eh_cpf then%>
        <!-- ************   ENDERE�O DE ENTREGA PARA CLIENTE PF   ************ -->
        <!-- Pegamos todos os atuais. Sem campos edit�veis. -->
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
			    <span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PJ');">Pessoa Jur�dica</span>
			    &nbsp;
			    <input type="radio" id="EndEtg_tipo_pessoa_PF" name="EndEtg_tipo_pessoa" value="PF" onclick="trocarEndEtgTipoPessoa(null);">
			    <span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PF');">Pessoa F�sica</span>
		    </p>
		    </td>
	    </tr>
    </table>

            <!-- ************   PJ: CNPJ/CONTRIBUINTE ICMS/IE - DO ENDERE�O DE ENTREGA DE PJ ************ -->
            <!-- ************   PF: CPF/PRODUTOR RURAL/CONTRIBUINTE ICMS/IE - DO ENDERE�O DE ENTREGA DE PJ  ************ -->
            <!-- fizemos dois conjuntos diferentes de campos porque a ordem � muito diferente -->
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
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PJ_nao" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">N�o</span>
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
		    <input type="radio" id="EndEtg_produtor_rural_status_PF_nao" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>');">N�o</span>
		    <input type="radio" id="EndEtg_produtor_rural_status_PF_sim" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>')">Sim</span></p></td>

	    <td align="left" class="MDE Mostrar_EndEtg_contribuinte_icms_PF"><p class="R">IE</p><p class="C">
		    <input id="EndEtg_ie_PF" name="EndEtg_ie_PF" class="TA" type="text" maxlength="20" size="13" value="" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_nome.focus(); filtra_nome_identificador();"></p>
	    </td>

	    <td align="left" class="Mostrar_EndEtg_contribuinte_icms_PF" ><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PF_nao" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">N�o</span>
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PF_sim" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>');">Sim</span>
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PF_isento" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>');">Isento</span></p>
	    </td>
	    </tr>
    </table>


    <!-- ************   ENDERE�O DE ENTREGA: NOME  ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td width="100%" align="left"><p class="R" id="Label_EndEtg_nome">RAZ�O SOCIAL</p><p class="C">
		    <input id="EndEtg_nome" name="EndEtg_nome" class="TA" value="" maxlength="60" size="85" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fNEW.EndEtg_endereco.focus(); filtra_nome_identificador();"></p></td>
	    </tr>
    </table>


    <%end if%>
<%end if%>


<!-- ************   ENDERE�O DE ENTREGA: ENDERE�O   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">ENDERE�O</p><p class="C">
		<input id="EndEtg_endereco" name="EndEtg_endereco" class="TA" value="" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDERE�O DE ENTREGA: N�/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">N�</p><p class="C">
		<input id="EndEtg_endereco_numero" name="EndEtg_endereco_numero" class="TA" value="" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td width="50%" align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<input id="EndEtg_endereco_complemento" name="EndEtg_endereco_complemento" class="TA" value="" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDERE�O DE ENTREGA: BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">BAIRRO</p><p class="C">
		<input id="EndEtg_bairro" name="EndEtg_bairro" class="TA" value="" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_cidade.focus(); filtra_nome_identificador();"></p></td>
	<td width="50%" align="left"><p class="R">CIDADE</p><p class="C">
		<input id="EndEtg_cidade" name="EndEtg_cidade" class="TA" value="" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDERE�O DE ENTREGA: UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="50%" class="MD" align="left"><p class="R">UF</p><p class="C">
		<input id="EndEtg_uf" name="EndEtg_uf" class="TA" value="" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fORC.EndEtg_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inv�lida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
	<td width="50%" align="left">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="left"><p class="R">CEP</p><p class="C">
				<input id="EndEtg_cep" name="EndEtg_cep" readonly tabindex=-1 class="TA" value="" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inv�lido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
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

        <!-- ************   ENDERE�O DE ENTREGA PARA PF: TELEFONES   ************ -->
        <!-- pegamos todos em branco (o usu�rio n�o poder� preencher eles) -->
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
        
        
        <!-- ************   ENDERE�O DE ENTREGA: TELEFONE RESIDENCIAL   ************ -->
        <table width="649" class="QS Mostrar_EndEtg_pf Habilitar_EndEtg_outroendereco" cellspacing="0">
	        <tr>
	        <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="EndEtg_ddd_res" name="EndEtg_ddd_res" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fNEW.EndEtg_tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	        <td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
		        <input id="EndEtg_tel_res" name="EndEtg_tel_res" class="TA" value="" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fNEW.EndEtg_ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        </tr>
	        <tr>
	        <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="EndEtg_ddd_cel" name="EndEtg_ddd_cel" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fNEW.EndEtg_tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	        <td align="left" class="MC"><p class="R">CELULAR</p><p class="C">
		        <input id="EndEtg_tel_cel" name="EndEtg_tel_cel" class="TA" value="" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fNEW.EndEtg_obs.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('N�mero de celular inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        </tr>
        </table>
	
        
        <!-- ************   ENDERE�O DE ENTREGA: TELEFONE COMERCIAL   ************ -->
        <table width="649" class="QS Mostrar_EndEtg_pj Habilitar_EndEtg_outroendereco" cellspacing="0">
	        <tr>
	        <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="EndEtg_ddd_com" name="EndEtg_ddd_com" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fNEW.EndEtg_tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	        <td class="MD" align="left"><p class="R">TELEFONE </p><p class="C">
		        <input id="EndEtg_tel_com" name="EndEtg_tel_com" class="TA" value="" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fNEW.EndEtg_ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        <td align="left"><p class="R">RAMAL</p><p class="C">
		        <input id="EndEtg_ramal_com" name="EndEtg_ramal_com" class="TA" value="" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_ddd_com_2.focus(); filtra_numerico();"></p></td>
	        </tr>
	        <tr>
	            <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
	            <input id="EndEtg_ddd_com_2" name="EndEtg_ddd_com_2" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fNEW.EndEtg_tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!!');this.focus();}" /></p>  
	            </td>
	            <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
	            <input id="EndEtg_tel_com_2" name="EndEtg_tel_com_2" class="TA" value="" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fNEW.EndEtg_ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	            </td>
	            <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
	            <input id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" class="TA" value="" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_obs.focus(); filtra_numerico();" /></p>
	            </td>
	        </tr>
        </table>

    <% end if %>
<% end if %>

<!-- ************   JUSTIFIQUE O ENDERE�O   ************ -->
<table  id="obs_endereco" width="649" class="QS" cellspacing="0">
	<tr>
	<td class="M" width="50%" align="left"><p class="R">JUSTIFIQUE O ENDERE�O</p><p class="C">
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
		<a name="bORCAMENTO" id="bORCAMENTO" href="javascript:fORCConcluir(fORC);" title="cadastra um novo or�amento para este cliente">
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
		window.status = "Aguarde, carregando p�gina ...";
		setTimeout("window.location='<%=pagina_retorno%>'", 1000);
	</script>
<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>