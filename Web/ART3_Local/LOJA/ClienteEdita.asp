<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  C L I E N T E E D I T A . A S P
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
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
'	EXIBIÇÃO DE BOTÕES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True
	
	
'	OBTEM O ID
	dim intCounter
	dim s, s_aux, usuario, loja, cnpj_cpf_selecionado, operacao_selecionada, pagina_retorno, s_codigo, s_descricao, s_codigo_e_descricao
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	loja = Trim(Session("loja_atual"))
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	CLIENTE A EDITAR
	operacao_selecionada = trim(request("operacao_selecionada"))
	cnpj_cpf_selecionado = retorna_so_digitos(trim(request("cnpj_cpf_selecionado")))
	
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	if (operacao_selecionada=OP_INCLUI) And (cnpj_cpf_selecionado="") then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO) 

	dim alerta
	alerta = ""

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
'	EDIÇÃO BLOQUEADA?
	dim edicao_bloqueada, blnEdicaoBloqueada
	edicao_bloqueada = ucase(trim(request("edicao_bloqueada")))
	blnEdicaoBloqueada = False
	if edicao_bloqueada = "S" then blnEdicaoBloqueada = True
	if Not operacao_permitida(OP_LJA_EDITA_CLIENTE_DADOS_CADASTRAIS, s_lista_operacoes_permitidas) then blnEdicaoBloqueada = True
	
'	ESTÁ DEFINIDA A PÁGINA QUE DEVE SER EXIBIDA APÓS A ATUALIZAÇÃO NO CADASTRO?
	pagina_retorno = trim(request("pagina_retorno"))


'	CONECTA COM O BANCO DE DADOS
	dim cn,rs,tRefBancaria,tRefComercial,tRefProfissional, tMAP_XML, tMAP_END_ETG, tMAP_END_COB
	dim cnBsp, blnHaRegistroBsp
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	if ID_AMBIENTE = ID_AMBIENTE__AT And operacao_selecionada = OP_INCLUI then
        if Not bdd_BS_conecta(cnBsp) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
        end if

	dim blnLojaHabilitadaProdCompostoECommerce
	blnLojaHabilitadaProdCompostoECommerce = isLojaHabilitadaProdCompostoECommerce(loja)

	dim intIdx
	Dim id_cliente, msg_erro
	if operacao_selecionada=OP_INCLUI then
		if Not gera_nsu(NSU_CADASTRO_CLIENTES, id_cliente, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
	else
		id_cliente = trim(request("cliente_selecionado"))
		if id_cliente = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
		end if

	s="select * from t_CLIENTE where (cnpj_cpf='" & cnpj_cpf_selecionado & "') Or (id='" & id_cliente & "')"
    set rs = cn.Execute(s)
    
    if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_JA_CADASTRADO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("clientepesquisa.asp?cnpj_cpf_selecionado=" & cnpj_cpf_selecionado & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		end if

    if ID_AMBIENTE = ID_AMBIENTE__AT And operacao_selecionada = OP_INCLUI then
        set rs = cnBsp.Execute(s)
        end if
	
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)

    blnHaRegistroBsp = False
    if ID_AMBIENTE = ID_AMBIENTE__AT And operacao_selecionada = OP_INCLUI And Not rs.EOF then blnHaRegistroBsp = True

	dim eh_cpf
	if (operacao_selecionada=OP_INCLUI) then
		s=cnpj_cpf_selecionado
	else
		s=Trim("" & rs("cnpj_cpf"))
		end if
		
	if len(s)=11 then eh_cpf=True else eh_cpf=False
	
'	REF BANCÁRIA
	dim blnCadRefBancaria
	dim int_MAX_REF_BANCARIA_CLIENTE
	dim strRefBancariaBanco, strRefBancariaAgencia, strRefBancariaConta
	dim strRefBancariaDdd, strRefBancariaTelefone, strRefBancariaContato
'	O cadastro de Referência Bancária será exibido p/ PF e PJ
	blnCadRefBancaria = True
	if eh_cpf then 
		int_MAX_REF_BANCARIA_CLIENTE = MAX_REF_BANCARIA_CLIENTE_PF
	else
		int_MAX_REF_BANCARIA_CLIENTE = MAX_REF_BANCARIA_CLIENTE_PJ
		end if

'	PJ: REF COMERCIAL
	dim blnCadRefComercial
	dim int_MAX_REF_COMERCIAL_CLIENTE
	dim strRefComercialNomeEmpresa, strRefComercialContato, strRefComercialDdd, strRefComercialTelefone
	if (Not eh_cpf) then blnCadRefComercial = True else blnCadRefComercial = False
	int_MAX_REF_COMERCIAL_CLIENTE = MAX_REF_COMERCIAL_CLIENTE_PJ

'	PF: REF PROFISSIONAL
	dim blnCadRefProfissional
	dim int_MAX_REF_PROFISSIONAL_CLIENTE
	dim strRefProfNomeEmpresa, strRefProfCargo, strRefProfDdd, strRefProfTelefone
	dim strRefProfPeriodoRegistro, strRefProfRendimentos, strRefProfCnpj
	if (eh_cpf) then blnCadRefProfissional = True else blnCadRefProfissional = False
	int_MAX_REF_PROFISSIONAL_CLIENTE = MAX_REF_PROFISSIONAL_CLIENTE_PF
	
'	PJ: DADOS DO SÓCIO MAJORITÁRIO
	dim blnCadSocioMaj
	if (Not eh_cpf) then blnCadSocioMaj = True else blnCadSocioMaj = False
	
'	INDICADOR
	dim blnCampoIndicadorEditavel
	blnCampoIndicadorEditavel = False
	if operacao_permitida(OP_LJA_EDITA_CLIENTE_CAMPO_INDICADOR, s_lista_operacoes_permitidas) then blnCampoIndicadorEditavel = True
	if Not operacao_permitida(OP_LJA_EDITA_CLIENTE_DADOS_CADASTRAIS, s_lista_operacoes_permitidas) then blnCampoIndicadorEditavel = False
	if operacao_selecionada=OP_INCLUI then blnCampoIndicadorEditavel = True
	
	dim intQtdeIndicadores, strCampoSelectIndicadores, strJsScriptArrayIndicadores
	intQtdeIndicadores = 0
	strCampoSelectIndicadores = ""
	strJsScriptArrayIndicadores = ""

	if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
		if blnCampoIndicadorEditavel then
			call indicadores_monta_itens_select(Trim("" & rs("indicador")), strCampoSelectIndicadores, strJsScriptArrayIndicadores)
			end if
	else
		call indicadores_monta_itens_select(Null, strCampoSelectIndicadores, strJsScriptArrayIndicadores)
		end if

	dim s_ddd, s_tel, s_nome_cliente, s_mag_end_etg_completo, s_mag_end_cob_completo
	s_nome_cliente = ""
	s_mag_end_etg_completo = ""
	s_mag_end_cob_completo = ""

	dim operacao_origem, c_numero_magento, operationControlTicket, sessionToken, id_magento_api_pedido_xml
	operacao_origem = Trim(Request("operacao_origem"))
	c_numero_magento = ""
	operationControlTicket = ""
	sessionToken = ""
	id_magento_api_pedido_xml = ""
	if alerta = "" then
		if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
			c_numero_magento = Trim(Request("c_numero_magento"))
			operationControlTicket = Trim(Request("operationControlTicket"))
			sessionToken = Trim(Request("sessionToken"))
			id_magento_api_pedido_xml = Trim(Request("id_magento_api_pedido_xml"))

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
		end if

	dim c_mag_customer_full_name, c_mag_customer_dob, c_mag_customer_email, c_mag_email_identificado
	dim c_mag_end_cob_email, c_mag_end_cob_telephone_ddd, c_mag_end_cob_telephone_numero, c_mag_end_cob_celular_ddd, c_mag_end_cob_celular_numero, c_mag_end_cob_fax_ddd, c_mag_end_cob_fax_numero, c_mag_end_cob_endereco, c_mag_end_cob_endereco_numero, c_mag_end_cob_complemento, c_mag_end_cob_bairro, c_mag_end_cob_cidade, c_mag_end_cob_uf, c_mag_end_cob_cep
	dim c_mag_end_etg_email, c_mag_end_etg_telephone_ddd, c_mag_end_etg_telephone_numero, c_mag_end_etg_celular_ddd, c_mag_end_etg_celular_numero, c_mag_end_etg_fax_ddd, c_mag_end_etg_fax_numero, c_mag_end_etg_endereco, c_mag_end_etg_endereco_numero, c_mag_end_etg_complemento, c_mag_end_etg_bairro, c_mag_end_etg_cidade, c_mag_end_etg_uf, c_mag_end_etg_cep
	c_mag_customer_full_name = ""
	c_mag_customer_dob = ""
	c_mag_customer_email = ""
	c_mag_email_identificado = ""
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





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ___________________________________
' L I S T A _ M I D I A
'
function lista_midia(byval id_default)
dim x,r,s,ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_MIDIA WHERE indisponivel=0 ORDER BY apelido")
	s= ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			s = s & "<option selected"
			ha_default=True
		else
			s = s & "<option"
			end if
		s = s & " value='" & x & "'>"
		s = s & Trim("" & r("apelido"))
		s = s & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		s = "<option selected value=''>&nbsp;</option>" & chr(13) & s
		end if
		
	lista_midia = s
	r.close
	set r=nothing
	
end function



' ___________________________________________________________________________
' INDICADORES MONTA ITENS SELECT
' LEMBRE-SE: O ORÇAMENTISTA É CONSIDERADO AUTOMATICAMENTE UM INDICADOR!!
function indicadores_monta_itens_select(byval id_default, byref strResp, byref strJsScript)
dim x, r, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False

	strJsScript = "<script language='JavaScript'>" & chr(13) & _
					"var vIndicador = new Array();" & chr(13) & _
					"vIndicador[0] = new oIndicador('', 0);" & chr(13)

	if ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then
		strSql = "SELECT " & _
					"*" & _
				" FROM t_ORCAMENTISTA_E_INDICADOR" & _
				" WHERE" & _
					" (apelido = '" & Trim("" & id_default) & "')" & _
					" OR " & _
					" (status = 'A')" & _
				" ORDER BY" & _
					" apelido"
	else
		'10/01/2020 - Unis - Desativação do acesso dos vendedores a todos os parceiros da Unis
		if (False And isLojaVrf(loja)) then
		'	TODOS OS VENDEDORES COMPARTILHAM OS MESMOS INDICADORES
			strSql = "SELECT " & _
						"*" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (apelido = '" & Trim("" & id_default) & "')" & _
						" OR " & _
						" ((status = 'A') AND (loja = '" & loja & "'))" & _
					" ORDER BY" & _
						" apelido"
		elseif (loja = NUMERO_LOJA_OLD03) Or (loja = NUMERO_LOJA_OLD03_BONIFICACAO) then
		'	OLD03: LISTA COMPLETA DOS INDICADORES LIBERADA
			strSql = "SELECT " & _
						"*" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (apelido = '" & Trim("" & id_default) & "')" & _
						" OR " & _
						" (status = 'A')" & _
					" ORDER BY" & _
						" apelido"
		else
			strSql = "SELECT " & _
						"*" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (apelido = '" & Trim("" & id_default) & "')" & _
						" OR " & _
						"(" & _
							" (status = 'A')" & _
							" AND (vendedor = '" & usuario & "')" & _
						")" & _
					" ORDER BY" & _
						" apelido"
			end if
		end if
	
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof
		intQtdeIndicadores = intQtdeIndicadores + 1
		x = Trim("" & r("apelido"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("razao_social_nome"))
		strResp = strResp & "</option>" & chr(13)
		
		strJsScript = strJsScript & _
						"vIndicador[vIndicador.length] = new oIndicador('" & QuotedStr(Trim("" & r("apelido"))) & "', " & Trim("" & r("permite_RA_status")) & ");" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
	
	strJsScript = strJsScript & "</script>" & chr(13)
	
	r.close
	set r=nothing
end function
%>


<%=DOCTYPE_LEGADO%>

<html>


<head>
    <title>LOJA</title>
</head>


<%
'		C L I E N T   S I D E   S C R I P T I N G
'
'      CCCCCCC   LLL        IIIII  EEEEEEEEE  NNN   NNN  TTTTTTTTT EEEEEEEEE
'     CCC   CCC  LLL         III   EEE        NNNN  NNN     TTT    EEE
'     CCC        LLL         III   EEE        NNNNN NNN     TTT    EEE
'     CCC        LLL         III   EEEEEE     NNN NNNNN     TTT    EEEEEE
'     CCC        LLL         III   EEE        NNN  NNNN     TTT    EEE
'     CCC   CCC  LLL   LLL   III   EEE        NNN   NNN     TTT    EEE
'      CCCCCCC   LLLLLLLLL  IIIII  EEEEEEEEE  NNN   NNN     TTT    EEEEEEEEE
'
%>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
    var ja_carregou = false;
    var conteudo_original;
    var fCepPopup;

    $(function () {
        var f;
        if ((typeof (fNEW) !== "undefined") && (fNEW !== null)) {
            f = fNEW;

            if (!f.rb_end_entrega[1].checked) {
                Disabled_change(f, true);
            }
        }

        // Trata o problema em que os campos do formulário são limpos após retornar à esta página c/ o history.back() pela 2ª vez quando ocorre erro de consistência
    if (trim(fCAD.c_FormFieldValues.value) != "")
    {
            stringToForm(fCAD.c_FormFieldValues.value, $('#fCAD'));
        }

        if ((typeof (fNEW) !== "undefined") && (fNEW !== null)) {
            if (trim(fNEW.c_FormFieldValues.value) != "") {
                stringToForm(fNEW.c_FormFieldValues.value, $('#fNEW'));
            }
        }
    });

    function copyMagentoCli() {
        var fFROM, fTO;
        fFROM = fMAG;
        fTO = fCAD;
        fTO.nome.value = fFROM.c_mag_customer_full_name.value;
        fTO.dt_nasc.value = fFROM.c_mag_customer_dob.value;
        fTO.email.value = fFROM.c_mag_email_identificado.value;
    }

    function copyMagentoBillAddrToBillAddr() {
        var fFROM, fTO;
        var s, eh_cpf;
        fFROM = fMAG;
        fTO = fCAD;
        fTO.endereco.value = fFROM.c_mag_end_cob_endereco.value;
        fTO.endereco_numero.value = fFROM.c_mag_end_cob_endereco_numero.value;
        fTO.endereco_complemento.value = fFROM.c_mag_end_cob_complemento.value;
        fTO.bairro.value = fFROM.c_mag_end_cob_bairro.value;
        fTO.cidade.value = fFROM.c_mag_end_cob_cidade.value;
        fTO.uf.value = fFROM.c_mag_end_cob_uf.value;
        fTO.cep.value = cep_formata(fFROM.c_mag_end_cob_cep.value);

        eh_cpf = false;
        s = retorna_so_digitos(fCAD.cnpj_cpf_selecionado.value);
        if (s.length == 11) eh_cpf = true;
        if (eh_cpf) {
            fTO.ddd_res.value = fFROM.c_mag_end_cob_telephone_ddd.value;
            fTO.tel_res.value = telefone_formata(fFROM.c_mag_end_cob_telephone_numero.value);
            fTO.ddd_cel.value = fFROM.c_mag_end_cob_celular_ddd.value;
            fTO.tel_cel.value = telefone_formata(fFROM.c_mag_end_cob_celular_numero.value);
        }
        else {
            fTO.ddd_com.value = fFROM.c_mag_end_cob_telephone_ddd.value;
            fTO.tel_com.value = telefone_formata(fFROM.c_mag_end_cob_telephone_numero.value);
            fTO.ddd_com_2.value = fFROM.c_mag_end_cob_celular_ddd.value;
            fTO.tel_com_2.value = telefone_formata(fFROM.c_mag_end_cob_celular_numero.value);
        }
    }

    function copyMagentoShipAddrToBillAddr() {
        var fFROM, fTO;
        var s, eh_cpf;
        fFROM = fMAG;
        fTO = fCAD;
        fTO.endereco.value = fFROM.c_mag_end_etg_endereco.value;
        fTO.endereco_numero.value = fFROM.c_mag_end_etg_endereco_numero.value;
        fTO.endereco_complemento.value = fFROM.c_mag_end_etg_complemento.value;
        fTO.bairro.value = fFROM.c_mag_end_etg_bairro.value;
        fTO.cidade.value = fFROM.c_mag_end_etg_cidade.value;
        fTO.uf.value = fFROM.c_mag_end_etg_uf.value;
        fTO.cep.value = cep_formata(fFROM.c_mag_end_etg_cep.value);

        eh_cpf = false;
        s = retorna_so_digitos(fCAD.cnpj_cpf_selecionado.value);
        if (s.length == 11) eh_cpf = true;
        if (eh_cpf) {
            fTO.ddd_res.value = fFROM.c_mag_end_etg_telephone_ddd.value;
            fTO.tel_res.value = telefone_formata(fFROM.c_mag_end_etg_telephone_numero.value);
            fTO.ddd_cel.value = fFROM.c_mag_end_etg_celular_ddd.value;
            fTO.tel_cel.value = telefone_formata(fFROM.c_mag_end_etg_celular_numero.value);
        }
        else {
            fTO.ddd_com.value = fFROM.c_mag_end_etg_telephone_ddd.value;
            fTO.tel_com.value = telefone_formata(fFROM.c_mag_end_etg_telephone_numero.value);
            fTO.ddd_com_2.value = fFROM.c_mag_end_etg_celular_ddd.value;
            fTO.tel_com_2.value = telefone_formata(fFROM.c_mag_end_etg_celular_numero.value);
        }
    }

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


    function ProcessaSelecaoCEP() { };

    function AbrePesquisaCep() {
        var f, strUrl;
	try
		{
            //  SE JÁ HOUVER UMA JANELA DE PESQUISA DE CEP ABERTA, GARANTE QUE ELA SERÁ FECHADA 
            // E UMA NOVA SERÁ CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')	
            fCepPopup = window.open("", "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=5,height=5,left=0,top=0");
            fCepPopup.close();
        }
        catch (e) {
            // NOP
        }
        f = fCAD;
        ProcessaSelecaoCEP = TrataCepEnderecoCadastro;
        strUrl = "../Global/AjaxCepPesqPopup.asp";
        if (trim(f.cep.value) != "") strUrl = strUrl + "?CepDefault=" + trim(f.cep.value);
        fCepPopup = window.open(strUrl, "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=980,height=650,left=0,top=0");
        fCepPopup.focus();
    }

    function TrataCepEnderecoCadastro(strCep, strUF, strLocalidade, strBairro, strLogradouro, strEnderecoNumero, strEnderecoComplemento) {
        var f;
        f = fCAD;
        f.cep.value = cep_formata(strCep);
        f.uf.value = strUF;
        f.cidade.value = strLocalidade;
        f.bairro.value = strBairro;
        f.endereco.value = strLogradouro;
        f.endereco_numero.value = strEnderecoNumero;
        f.endereco_complemento.value = strEnderecoComplemento;
        f.endereco.focus();
        window.status = "Concluído";
    }

    function AbrePesquisaCepEndEtg() {
        var f, strUrl;
	try
		{
            //  SE JÁ HOUVER UMA JANELA DE PESQUISA DE CEP ABERTA, GARANTE QUE ELA SERÁ FECHADA 
            // E UMA NOVA SERÁ CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')	
            fCepPopup = window.open("", "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=5,height=5,left=0,top=0");
            fCepPopup.close();
        }
        catch (e) {
            // NOP
        }
        f = fNEW;
        ProcessaSelecaoCEP = TrataCepEnderecoEntrega;
        strUrl = "../Global/AjaxCepPesqPopup.asp";
        if (trim(f.EndEtg_cep.value) != "") strUrl = strUrl + "?CepDefault=" + trim(f.EndEtg_cep.value);
        fCepPopup = window.open(strUrl, "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=980,height=650,left=0,top=0");
        fCepPopup.focus();
    }

    function TrataCepEnderecoEntrega(strCep, strUF, strLocalidade, strBairro, strLogradouro, strEnderecoNumero, strEnderecoComplemento) {
        var f;
        f = fNEW;
        f.EndEtg_cep.value = cep_formata(strCep);
        f.EndEtg_uf.value = strUF;
        f.EndEtg_cidade.value = strLocalidade;
        f.EndEtg_bairro.value = strBairro;
        f.EndEtg_endereco.value = strLogradouro;
        f.EndEtg_endereco_numero.value = strEnderecoNumero;
        f.EndEtg_endereco_complemento.value = strEnderecoComplemento;
        f.EndEtg_endereco.focus();
        window.status = "Concluído";
    }

    function consiste_endereco_cadastro(f) {

        if (trim(f.endereco.value) == "") {
            alert('Preencha o endereço!!');
            f.endereco.focus();
            return false;
        }

        if (trim(f.endereco_numero.value) == "") {
            alert('Preencha o número do endereço!!');
            f.endereco_numero.focus();
            return false;
        }

        if (trim(f.bairro.value) == "") {
            alert('Preencha o bairro!!');
            f.bairro.focus();
            return false;
        }

        if (trim(f.cidade.value) == "") {
            alert('Preencha a cidade!!');
            f.cidade.focus();
            return false;
        }

        s = trim(f.uf.value);
        if ((s == "") || (!uf_ok(s))) {
            alert('UF inválida!!');
            f.uf.focus();
            return false;
        }

        if (trim(f.cep.value) == "") {
            alert('Informe o CEP!!');
            return false;
        }

        if (!cep_ok(f.cep.value)) {
            alert('CEP inválido!!');
            f.cep.focus();
            return false;
        }

        return true;
    }

    function fNEWConcluir(f) {
        var s;
        var eh_cpf;
        if (!ja_carregou) return;

        s = retorna_so_digitos(fCAD.cnpj_cpf_selecionado.value);
        eh_cpf = false;
        if (s.length == 11) eh_cpf = true;

        s = retorna_dados_formulario(fCAD);
        if (s != conteudo_original) {
            if (!confirm("As alterações feitas serão perdidas!!\nContinua mesmo assim?")) return;
        }

        if ((!f.rb_end_entrega[0].checked) && (!f.rb_end_entrega[1].checked)) {
            alert('Informe se o endereço de entrega será o mesmo endereço do cadastro ou não!!');
            return;
        }


        if (f.rb_end_entrega[1].checked) {
            if (trim(f.EndEtg_endereco.value) == "") {
                alert('Preencha o endereço de entrega!!');
                f.EndEtg_endereco.focus();
                return;
            }

            if (trim(f.EndEtg_endereco_numero.value) == "") {
                alert('Preencha o número do endereço de entrega!!');
                f.EndEtg_endereco_numero.focus();
                return;
            }

            if (trim(f.EndEtg_bairro.value) == "") {
                alert('Preencha o bairro do endereço de entrega!!');
                f.EndEtg_bairro.focus();
                return;
            }

            if (trim(f.EndEtg_cidade.value) == "") {
                alert('Preencha a cidade do endereço de entrega!!');
                f.EndEtg_cidade.focus();
                return;
            }
            if (trim(f.EndEtg_obs.value) == "") {
                alert('Selecione a justificativa do endereço de entrega!!');
                f.EndEtg_obs.focus();
                return;
            }
            s = trim(f.EndEtg_uf.value);
            if ((s == "") || (!uf_ok(s))) {
                alert('UF inválida no endereço de entrega!!');
                f.EndEtg_uf.focus();
                return;
            }

            if (trim(f.EndEtg_cep.value) == "") {
                alert('Informe o CEP do endereço de entrega!!');
                f.EndEtg_cep.focus();
                return;
            }

            if (!cep_ok(f.EndEtg_cep.value)) {
                alert('CEP inválido no endereço de entrega!!');
                f.EndEtg_cep.focus();
                return;
            }


<%if Not eh_cpf then%>
            var EndEtg_tipo_pessoa = $('input[name="EndEtg_tipo_pessoa"]:checked').val();
            if (!EndEtg_tipo_pessoa)
                EndEtg_tipo_pessoa = "";
            if (EndEtg_tipo_pessoa != "PJ" && EndEtg_tipo_pessoa != "PF") {
                alert('Necessário escolher Pessoa Jurídica ou Pessoa Física no Endereço de entrega!!');
                f.EndEtg_tipo_pessoa.focus();
                return;
            }

            if (trim(f.EndEtg_nome.value) == "") {
                alert('Preencha o nome/razão social no endereço de entrega!!');
                f.EndEtg_nome.focus();
                return;
            }



            if (EndEtg_tipo_pessoa == "PJ") {
                //Campos PJ: 

                if (f.EndEtg_cnpj_cpf_PJ.value == "" || !cnpj_ok(f.EndEtg_cnpj_cpf_PJ.value)) {
                    alert('Endereço de entrega: CNPJ inválido!!');
                    f.EndEtg_cnpj_cpf_PF.focus();
                    return;
                }

                if ($('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').length == 0) {
                    alert('Endereço de entrega: selecione o tipo de contribuinte de ICMS!!');
                    f.EndEtg_contribuinte_icms_status_PJ.focus();
                    return;
                }

                /*
                sem validação: EndEtg_ie_PJ e  EndEtg_contribuinte_icms_status_PJ

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


                if ((trim(f.EndEtg_email.value) != "") && (!email_ok(f.EndEtg_email.value))) {
                    alert('Endereço de entrega: e-mail inválido!!');
                    f.EndEtg_email.focus();
                    return;
                }

                if ((trim(f.EndEtg_email_xml.value) != "") && (!email_ok(f.EndEtg_email_xml.value))) {
                    alert('Endereço de entrega: e-mail (XML) inválido!!');
                    f.EndEtg_email_xml.focus();
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

                //sem validação: EndEtg_rg_PF

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

        }

        if (trim(fCAD.cep.value) == "") {
            alert('É necessário preencher o CEP no cadastro do cliente!!');
            return;
        }

        if (eh_cpf) {
            if ((trim(fCAD.produtor_rural_cadastrado.value) == "0") ||
                ((fCAD.rb_produtor_rural[0].checked) && (fCAD.produtor_rural_cadastrado.value != fCAD.rb_produtor_rural[0].value)) ||
                ((fCAD.rb_produtor_rural[1].checked) && (fCAD.produtor_rural_cadastrado.value != fCAD.rb_produtor_rural[1].value))) {
                alert('É necessário gravar os dados do cadastro do cliente para que a opção Produtor Rural seja gravada!');
                return;
            }
            if ((!fCAD.rb_produtor_rural[0].checked) && (!fCAD.rb_produtor_rural[1].checked)) {
                alert('Informe se o cliente é produtor rural ou não!!');
                return;
            }
            if (fCAD.rb_produtor_rural[1].checked) {
                if (!fCAD.rb_contribuinte_icms[1].checked) {
                    alert('Para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!!');
                    return;
                }
                if ((!fCAD.rb_contribuinte_icms[0].checked) && (!fCAD.rb_contribuinte_icms[1].checked) && (!fCAD.rb_contribuinte_icms[2].checked)) {
                    alert('Informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
                    return;
                }
                if ((trim(fCAD.contribuinte_icms_cadastrado.value) == "0") ||
                    ((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[0].value)) ||
                    ((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[1].value)) ||
                    ((fCAD.rb_contribuinte_icms[2].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[2].value))) {
                    alert('É necessário gravar os dados do cadastro do cliente para que a opção Contribuinte ICMS seja gravada!');
                    return;
                }
            }
        }
        else {
            if ((trim(fCAD.contribuinte_icms_cadastrado.value) == "0") ||
                ((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[0].value)) ||
                ((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[1].value)) ||
                ((fCAD.rb_contribuinte_icms[2].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[2].value))) {
                alert('É necessário gravar os dados do cadastro do cliente para que a opção Contribuinte ICMS seja gravada!');
                return;
            }
        }

        //  Verifica se o endereço do cadastro está devidamente preenchido
        if (!consiste_endereco_cadastro(fCAD)) return;
        if (trim(fCAD.endereco_numero_cadastrado.value) == "") {
            if (trim(fCAD.endereco_numero.value) == "") {
                alert('É necessário preencher o número do endereço e, em seguida, gravar os dados do cadastro!');
            }
            else {
                alert('É necessário gravar os dados do cadastro do cliente para que o número do endereço seja gravado!');
            }
            return;
        }
        // Verifica se o campo IE está vazio quando contribuinte ICMS = isento
        if (eh_cpf) {
            if (!fCAD.rb_produtor_rural[0].checked) {
                if ((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                    fCAD.ie.focus();
                    return;
                }
                if ((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                    fCAD.ie.focus();
                    return;
                }
                if (fCAD.rb_contribuinte_icms[2].checked) {
                    if (fCAD.ie.value != "") {
                        alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
                        fCAD.ie.focus();
                        return;
                    }
                }
            }
        }
        else {
            if ((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                fCAD.ie.focus();
                return;
            }
            if ((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                fCAD.ie.focus();
                return;
            }
            if (fCAD.rb_contribuinte_icms[2].checked) {
                if (fCAD.ie.value != "") {
                    alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
                    fCAD.ie.focus();
                    return;
                }
            }
        }

    <% if CStr(loja) <> CStr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then %>
    // PARA CLIENTE PJ, É OBRIGATÓRIO O PREENCHIMENTO DO E-MAIL
    if (!eh_cpf) {
            if ((trim(fCAD.email_original.value) == "") && (trim(fCAD.email_xml_original.value) == "")) {
                alert("É obrigatório que o cliente tenha um endereço de e-mail cadastrado!");
                fCAD.email.focus();
                return;
            }
        }
    <% end if %>

            fNEW.c_FormFieldValues.value = formToString($("#fNEW"));

        //campos do endereço de entrega que precisam de transformacao
        transferirCamposEndEtg(fNEW);

        dPEDIDO.style.visibility = "hidden";
        window.status = "Aguarde ...";
        f.submit();
    }

    function RemoveCliente(f) {
        var b;
        if (!ja_carregou) return;

        b = window.confirm('Confirma a exclusão deste cliente?');
        if (b) {
            f.operacao_selecionada.value = OP_EXCLUI;
            dREMOVE.style.visibility = "hidden";
            window.status = "Aguarde ...";
            f.submit();
        }
    }

    function AtualizaCliente(f) {
        var s, eh_cpf, i, blnConsistir, blnConsistirDadosBancarios, blnOk;
        var blnCadRefBancaria, blnCadSocioMaj, blnCadRefComercial, blnCadRefProfissional;

        if (!ja_carregou) return;

        s = retorna_so_digitos(f.cnpj_cpf_selecionado.value);
        eh_cpf = false;
        if (s.length == 11) eh_cpf = true;

        if ((s == "") || (!cnpj_cpf_ok(s))) {
            alert('CNPJ/CPF inválido!!');
            return;
        }

        if (eh_cpf) {
            s = trim(f.sexo.value);
            if ((s == "") || (!sexo_ok(s))) {
                alert('Indique qual o sexo!!');
                f.sexo.focus();
                return;
            }
            if (!isDate(f.dt_nasc)) {
                alert('Data inválida!!');
                f.dt_nasc.focus();
                return;
            }
            if ((!f.rb_produtor_rural[0].checked) && (!f.rb_produtor_rural[1].checked)) {
                alert('Informe se o cliente é produtor rural ou não!!');
                return;
            }
            if (!f.rb_produtor_rural[0].checked) {
                if (!fCAD.rb_contribuinte_icms[1].checked) {
                    alert('Para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!!');
                    return;
                }
                if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
                    alert('Informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
                    return;
                }
                if ((f.rb_contribuinte_icms[1].checked) && (trim(f.ie.value) == "")) {
                    alert('Se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
                    f.ie.focus();
                    return;
                }
                if ((f.rb_contribuinte_icms[0].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                    f.ie.focus();
                    return;
                }
                if ((f.rb_contribuinte_icms[1].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                    f.ie.focus();
                    return;
                }
            }
        }
        else {
            //deixar de exigir preenchimento se cliente não é contribuinte?
            //s=trim(f.ie.value);
            //if (s=="") {
            //	alert('Preencha a Inscrição Estadual!!');
            //	f.ie.focus();
            //	return;
            //	}
            s = trim(f.contato.value);
            if (s == "") {
                alert('Informe o nome da pessoa para contato!!');
                f.contato.focus();
                return;
            }
            if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
                alert('Informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
                return;
            }
            if ((f.rb_contribuinte_icms[1].checked) && (trim(f.ie.value) == "")) {
                alert('Se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
                f.ie.focus();
                return;
            }
            if ((f.rb_contribuinte_icms[0].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                f.ie.focus();
                return;
            }
            if ((f.rb_contribuinte_icms[1].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                f.ie.focus();
                return;
            }
        }

        // Verifica se o campo IE está vazio quando contribuinte ICMS = isento
        if (eh_cpf) {
            if (!fCAD.rb_produtor_rural[0].checked) {
                if (fCAD.rb_contribuinte_icms[2].checked) {
                    if (fCAD.ie.value != "") {
                        alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
                        fCAD.ie.focus();
                        return;
                    }
                }
            }
        }
        else {
            if (fCAD.rb_contribuinte_icms[2].checked) {
                if (fCAD.ie.value != "") {
                    alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
                    fCAD.ie.focus();
                    return;
                }
            }
        }

        if (trim(f.nome.value) == "") {
            alert('Preencha o nome!!');
            f.nome.focus();
            return;
        }

        if (!consiste_endereco_cadastro(f)) return;

        if (eh_cpf) {
            if (!ddd_ok(f.ddd_res.value)) {
                alert('DDD inválido!!');
                f._res.focus();
                return;
            }
            if (!telefone_ok(f.tel_res.value)) {
                alert('Telefone inválido!!');
                f.tel_res.focus();
                return;
            }
            if ((trim(f.ddd_res.value) != "") || (trim(f.tel_res.value) != "")) {
                if (trim(f.ddd_res.value) == "") {
                    alert('Preencha o DDD!!');
                    f.ddd_res.focus();
                    return;
                }
                if (trim(f.tel_res.value) == "") {
                    alert('Preencha o telefone!!');
                    f.tel_res.focus();
                    return;
                }
            }

        }
        if (eh_cpf) {
            if (!ddd_ok(f.ddd_cel.value)) {
                alert('DDD inválido!!');
                f.ddd_cel.focus();
                return;
            }
            if (!telefone_ok(f.tel_cel.value)) {
                alert('Telefone inválido!!');
                f.tel_res.focus();
                return;
            }
            if ((f.ddd_cel.value == "") && (f.tel_cel.value != "")) {
                alert('Preencha o DDD do celular.');
                f.ddd_cel.focus();
                return;
            }
            if ((f.tel_cel.value == "") && (f.ddd_cel.value != "")) {
                alert('Preencha o número do celular.');
                f.tel_cel.focus();
                return;
            }
        }
        if (!eh_cpf) {
            if (!ddd_ok(f.ddd_com_2.value)) {
                alert('DDD inválido!!');
                f.ddd_com_2.focus();
                return;
            }
            if (!telefone_ok(f.tel_com_2.value)) {
                alert('Telefone inválido!!');
                f.tel_com_2.focus();
                return;
            }
            if ((f.ddd_com_2.value == "") && (f.tel_com_2.value != "")) {
                alert('Preencha o DDD do telefone.');
                f.ddd_com_2.focus();
                return;
            }
            if ((f.tel_com_2.value == "") && (f.ddd_com_2.value != "")) {
                alert('Preencha o telefone.');
                f.tel_com_2.focus();
                return;
            }

        }


        if (!ddd_ok(f.ddd_com.value)) {
            alert('DDD inválido!!');
            f.ddd_com.focus();
            return;
        }

        if (!telefone_ok(f.tel_com.value)) {
            alert('Telefone comercial inválido!!');
            f.tel_com.focus();
            return;
        }

        if ((trim(f.ddd_com.value) != "") || (trim(f.tel_com.value) != "")) {
            if (trim(f.ddd_com.value) == "") {
                alert('Preencha o DDD!!');
                f.ddd_com.focus();
                return;
            }
            if (trim(f.tel_com.value) == "") {
                alert('Preencha o telefone!!');
                f.tel_com.focus();
                return;
            }
        }

        if (eh_cpf) {
            if ((trim(f.tel_res.value) == "") && (trim(f.tel_com.value) == "") && (trim(f.tel_cel.value) == "")) {
                alert('Preencha pelo menos um telefone!!');
                return;
            }
        }
        else {
            if (trim(f.tel_com_2.value) == "") {
                if (trim(f.ddd_com.value) == "") {
                    alert('Preencha o DDD!!');
                    f.ddd_com.focus();
                    return;
                }
                if (trim(f.tel_com.value) == "") {
                    alert('Preencha o telefone!!');
                    f.tel_com.focus();
                    return;
                }
            }
        }

        if ((trim(f.email.value) != "") && (!email_ok(f.email.value))) {
            alert('E-mail inválido!!');
            f.email.focus();
            return;
        }

        if ((trim(f.email_xml.value) != "") && (!email_ok(f.email_xml.value))) {
            alert('E-mail (XML) inválido!!');
            f.email_xml.focus();
            return;
        }

    <% if CStr(loja) <> CStr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then %>
    // PARA CLIENTE PJ, É OBRIGATÓRIO O PREENCHIMENTO DO E-MAIL
    if (!eh_cpf) {
            if ((trim(fCAD.email.value) == "") && (trim(fCAD.email_xml.value) == "")) {
                alert("É obrigatório informar um endereço de e-mail");
                fCAD.email.focus();
                return;
            }
        }
    <% end if %>

/*
	if (trim(f.midia.options[f.midia.selectedIndex].value)=="") {
		alert('Indique a forma pela qual conheceu a Bonshop!!');
		return;
		}
*/

//  Ref Bancaria
		//  O cadastro de Referência Bancária será feito p/  PJ

		if (!eh_cpf) {
                blnCadRefBancaria = true;
                if (blnCadRefBancaria) {
                    for (i = 1; i < f.c_RefBancariaBanco.length; i++) {
                        blnConsistir = false;
                        if (trim(f.c_RefBancariaBanco[i].value) != "") blnConsistir = true;
                        if (trim(f.c_RefBancariaAgencia[i].value) != "") blnConsistir = true;
                        if (trim(f.c_RefBancariaConta[i].value) != "") blnConsistir = true;
                        if (trim(f.c_RefBancariaDdd[i].value) != "") blnConsistir = true;
                        if (trim(f.c_RefBancariaTelefone[i].value) != "") blnConsistir = true;
                        if (trim(f.c_RefBancariaContato[i].value) != "") blnConsistir = true;
                        if (blnConsistir) {
                            if (trim(f.c_RefBancariaBanco[i].value) == "") {
                                alert('Informe o banco no cadastro de Referência Bancária!!');
                                f.c_RefBancariaBanco[i].focus();
                                return;
                            }
                            if (trim(f.c_RefBancariaAgencia[i].value) == "") {
                                alert('Informe a agência no cadastro de Referência Bancária!!');
                                f.c_RefBancariaAgencia[i].focus();
                                return;
                            }
                            if (trim(f.c_RefBancariaConta[i].value) == "") {
                                alert('Informe o número da conta no cadastro de Referência Bancária!!');
                                f.c_RefBancariaConta[i].focus();
                                return;
                            }
                        }
                    }
                }
            }

        //  Ref Profissional
        //  O cadastro de Referência Profissional será feito apenas p/ PF
        /*
            if (eh_cpf) blnCadRefProfissional=true; else blnCadRefProfissional=false;
            if (blnCadRefProfissional) {
                for (i=1; i<f.c_RefProfNomeEmpresa.length; i++) {
                    blnConsistir=false;
                    if (trim(f.c_RefProfNomeEmpresa[i].value)!="") blnConsistir=true;
                    if (trim(f.c_RefProfCargo[i].value)!="") blnConsistir=true;
                    if (trim(f.c_RefProfDdd[i].value)!="") blnConsistir=true;
                    if (trim(f.c_RefProfTelefone[i].value)!="") blnConsistir=true;
                    if (trim(f.c_RefProfPeriodoRegistro[i].value)!="") blnConsistir=true;
                    if (trim(f.c_RefProfRendimentos[i].value)!="") blnConsistir=true;
                    if (trim(f.c_RefProfCnpj[i].value)!="") blnConsistir=true;
                    if (blnConsistir) {
                        if (trim(f.c_RefProfNomeEmpresa[i].value)=="") {
                            alert('Informe o nome da empresa no cadastro de Referência Profissional!!');
                            f.c_RefProfNomeEmpresa[i].focus();
                            return;
                            }
                        if (trim(f.c_RefProfCargo[i].value)=="") {
                            alert('Informe o cargo no cadastro de Referência Profissional!!');
                            f.c_RefProfCargo[i].focus();
                            return;
                            }
                        if (trim(f.c_RefProfCnpj[i].value)!="") {
                            if (!cnpj_ok(f.c_RefProfCnpj[i].value)) {
                                alert('CNPJ inválido!!');
                                f.c_RefProfCnpj[i].focus();
                                return;
                                }
                            }
                        }
                    }
                }
        */

        //  Ref Comercial
        //  O cadastro de Referência Comercial será feito apenas p/ PJ
        if (!eh_cpf) blnCadRefComercial = true; else blnCadRefComercial = false;
        if (blnCadRefComercial) {
            for (i = 1; i < f.c_RefComercialNomeEmpresa.length; i++) {
                blnConsistir = false;
                if (trim(f.c_RefComercialNomeEmpresa[i].value) != "") blnConsistir = true;
                if (trim(f.c_RefComercialContato[i].value) != "") blnConsistir = true;
                if (trim(f.c_RefComercialDdd[i].value) != "") blnConsistir = true;
                if (trim(f.c_RefComercialTelefone[i].value) != "") blnConsistir = true;
                if (blnConsistir) {
                    if (trim(f.c_RefComercialNomeEmpresa[i].value) == "") {
                        alert('Informe o nome da empresa no cadastro de Referência Comercial!!');
                        f.c_RefComercialNomeEmpresa[i].focus();
                        return;
                    }
                }
            }
        }

        //  Dados do Sócio Majoritário
        /*	if (!eh_cpf) blnCadSocioMaj=true; else blnCadSocioMaj=false;
            if (blnCadSocioMaj) {
                blnConsistir=false;
                blnConsistirDadosBancarios=false;
                if (trim(f.c_SocioMajNome.value)!="") blnConsistir=true;
                if (trim(f.c_SocioMajCpf.value)!="") blnConsistir=true;
                if (trim(f.c_SocioMajBanco.value)!="") {
                    blnConsistir=true;
                    blnConsistirDadosBancarios=true;
                    }
                if (trim(f.c_SocioMajAgencia.value)!="") {
                    blnConsistir=true;
                    blnConsistirDadosBancarios=true;
                    }
                if (trim(f.c_SocioMajConta.value)!="") {
                    blnConsistir=true;
                    blnConsistirDadosBancarios=true;
                    }
                if (trim(f.c_SocioMajDdd.value)!="") blnConsistir=true;
                if (trim(f.c_SocioMajTelefone.value)!="") blnConsistir=true;
                if (trim(f.c_SocioMajContato.value)!="") blnConsistir=true;
                if (blnConsistir) {
                    if (trim(f.c_SocioMajNome.value)=="") {
                        alert('Informe o nome do sócio majoritário!!');
                        f.c_SocioMajNome.focus();
                        return;
                        }
                    }
                if (blnConsistirDadosBancarios) {
                    if (trim(f.c_SocioMajBanco.value)=="") {
                        alert('Informe o banco nos dados bancários do sócio majoritário!!');
                        f.c_SocioMajBanco.focus();
                        return;
                        }
                    if (trim(f.c_SocioMajAgencia.value)=="") {
                        alert('Informe a agência nos dados bancários do sócio majoritário!!');
                        f.c_SocioMajAgencia.focus();
                        return;
                        }
                    if (trim(f.c_SocioMajConta.value)=="") {
                        alert('Informe o número da conta nos dados bancários do sócio majoritário!!');
                        f.c_SocioMajConta.focus();
                        return;
                        }
                    }
                }
        */

        fCAD.c_FormFieldValues.value = formToString($("#fCAD"));

        dATUALIZA.style.visibility = "hidden";
        window.status = "Aguarde ...";
        f.submit();
    }

</script>

<script type="text/javascript">

    function exibeJanelaCEP_Cli() {
        $.mostraJanelaCEP("cep", "uf", "cidade", "bairro", "endereco", "endereco_numero", "endereco_complemento");
    }

    function exibeJanelaCEP_Etg() {
        $.mostraJanelaCEP("EndEtg_cep", "EndEtg_uf", "EndEtg_cidade", "EndEtg_bairro", "EndEtg_endereco", "EndEtg_endereco_numero", "EndEtg_endereco_complemento");
    }

    function trataProdutorRural() {
        //ao clicar na opção Produtor Rural, exibir/ocultar os campos apropriados
        if ((typeof (fCAD.rb_produtor_rural) !== "undefined") && (fCAD.rb_produtor_rural !== null)) {
            if (!fCAD.rb_produtor_rural[1].checked) {
                $("#t_contribuinte_icms").css("display", "none");
            }
            else {
                $("#t_contribuinte_icms").css("display", "block");
            }
        }
    }



    function transferirCamposEndEtg(fNEW) {
        //Transferimos os dados do endereço de entrega dos campos certos. 
        //Temos dois conjuntos de campos (para PF e PJ) porque o layout é muito diferente.
        var pj = $('input[name="EndEtg_tipo_pessoa"]:checked').val() == "PJ";
        if (pj) {
            fNEW.EndEtg_cnpj_cpf = fNEW.EndEtg_cnpj_cpf_PJ;
            fNEW.EndEtg_ie = fNEW.EndEtg_ie_PJ;
            fNEW.EndEtg_contribuinte_icms_status = fNEW.EndEtg_contribuinte_icms_status_PJ;
        }
        else {
            fNEW.EndEtg_cnpj_cpf = fNEW.EndEtg_cnpj_cpf_PJ;
            fNEW.EndEtg_ie = fNEW.EndEtg_ie_PJ;
            fNEW.EndEtg_contribuinte_icms_status = fNEW.EndEtg_contribuinte_icms_status_PJ;
            fNEW.EndEtg_rg = fNEW.EndEtg_rg_PJ;
            fNEW.EndEtg_produtor_rural_status = fNEW.EndEtg_produtor_rural_status_PJ;
        }

        //Tip: Disabled <input> elements in a form will not be submitted!
        //entao deixamos como disabled todos os que usamos para montar estes dados!
        if(fNEW.EndEtg_cnpj_cpf_PJ) fNEW.EndEtg_cnpj_cpf_PJ.disabled = true;
        if(fNEW.EndEtg_ie_PJ) fNEW.EndEtg_ie_PJ.disabled = true;
        if(fNEW.EndEtg_contribuinte_icms_status_PJ) fNEW.EndEtg_contribuinte_icms_status_PJ.disabled = true;
        if(fNEW.EndEtg_cnpj_cpf_PJ) fNEW.EndEtg_cnpj_cpf_PJ.disabled = true;
        if(fNEW.EndEtg_ie_PJ) fNEW.EndEtg_ie_PJ.disabled = true;
        if(fNEW.EndEtg_contribuinte_icms_status_PJ) fNEW.EndEtg_contribuinte_icms_status_PJ.disabled = true;
        if(fNEW.EndEtg_rg_PJ) fNEW.EndEtg_rg_PJ.disabled = true;
        if(fNEW.EndEtg_produtor_rural_status_PJ) fNEW.EndEtg_produtor_rural_status_PJ.disabled = true;
    }

    //para mudar o tipo do endereço de entrega
    function trocarEndEtgTipoPessoa(novoTipo) {
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
            $(".Mostrar_EndEtg_contribuinte_icms_PF").css("display", "block");
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
        width: 130px;
        text-align: right;
    }
.TdCliCel
{
        width: 520px;
        text-align: left;
    }
.TdCliBtn
{
        width: 30px;
        text-align: center;
    }
</style>



<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body>
    <center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<%	if operacao_selecionada=OP_INCLUI then
		if eh_cpf then
			s = "fCAD.rg.focus();"
		else
			s = "fCAD.ie.focus();"
			end if
	else
		s = "focus();"
		end if
%>
<body id="corpoPagina" onload="<%=s%>conteudo_original=retorna_dados_formulario(fCAD);ja_carregou=true;trataProdutorRural();">

    <center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->


<!--  CADASTRO DO CLIENTE -->

<table cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;width:698px;" border="0">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Novo Cliente"
	else
		s = "Consulta/Edição de Cliente Cadastrado"
		end if
%>
	<td align="center" valign="bottom"><span class="PEDIDO"><%=s%></span></td>
</tr>
</table>
<br />

<!-- ************   EXIBE OBSERVAÇÕES CREDITÍCIAS?  ************ -->
<%	if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("obs_crediticias")) else s=""
	if s <> "" then %>
		<span class="Lbl" style="display:none">OBSERVAÇÕES CREDITÍCIAS</span>
		<div class='MtAviso' style="width:649px;FONT-WEIGHT:bold;border:1pt solid black;display:none;" align="CENTER"><P style='margin:5px 2px 5px 2px;'><%=s%></p></div>
		<br>
	<% end if %>


<!-- ************  CAMPOS DO CADASTRO  ************ -->
<form id="fCAD" name="fCAD" method="post" onload="trataProdutorRural();" action="ClienteAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<INPUT type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>
<INPUT type="hidden" name='cliente_selecionado' id="cliente_selecionado" value='<%=id_cliente%>'>
<INPUT type="hidden" name='pagina_retorno' id="pagina_retorno" value='<%=pagina_retorno%>'>
<%if operacao_selecionada=OP_CONSULTA then%>
<INPUT type="hidden" name='endereco_numero_cadastrado' id="endereco_numero_cadastrado" value='<%=Trim("" & rs("endereco_numero"))%>'>
<INPUT type="hidden" name='contribuinte_icms_cadastrado' id="contribuinte_icms_cadastrado" value='<%=Trim("" & rs("contribuinte_icms_status"))%>'>
<INPUT type="hidden" name='produtor_rural_cadastrado' id="produtor_rural_cadastrado" value='<%=Trim("" & rs("produtor_rural_status"))%>'>
<%else%>
<INPUT type="hidden" name='endereco_numero_cadastrado' id="endereco_numero_cadastrado" value=''>
<INPUT type="hidden" name='contribuinte_icms_cadastrado' id="contribuinte_icms_cadastrado" value=''>
<INPUT type="hidden" name='produtor_rural_cadastrado' id="produtor_rural_cadastrado" value=''>
<%end if%>

<%if blnCampoIndicadorEditavel then%>
<INPUT type="hidden" name='CampoIndicadorEditavel' id="CampoIndicadorEditavel" value='S'>
<%else%>
<INPUT type="hidden" name='CampoIndicadorEditavel' id="CampoIndicadorEditavel" value='N'>
<%end if%>

<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />
<input type="hidden" name="operacao_origem" id="operacao_origem" value="<%=operacao_origem%>" />
<input type="hidden" name="id_magento_api_pedido_xml" id="id_magento_api_pedido_xml" value="<%=id_magento_api_pedido_xml%>" />
<input type="hidden" name="c_numero_magento" id="c_numero_magento" value="<%=c_numero_magento%>" />
<input type="hidden" name="operationControlTicket" id="operationControlTicket" value="<%=operationControlTicket%>" />
<input type="hidden" name="sessionToken" id="sessionToken" value="<%=sessionToken%>" />


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
			<% if c_mag_email_identificado <> "" then %>
			<br /><span class="C"><%=c_mag_email_identificado%></span>
			<% end if %>
		</td>
		<td class="MB MD TdCliBtn"><a href="javascript:copyMagentoCli();"><img src="../IMAGEM/copia_20x20.png" name="btnMagentoCopyCli" id="btnMagentoCopyCli" title="Altera dados do cliente usando informações do Magento" /></a></td>
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
		<td class="MB MD TdCliBtn"><a href="javascript:copyMagentoBillAddrToBillAddr();"><img src="../IMAGEM/copia_20x20.png" name="btnMagentoCopyBillAddrToBillAddr" id="btnMagentoCopyBillAddrToBillAddr" title="Altera o endereço usando os dados do endereço de cobrança obtidos do Magento" /></a></td>
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
		<td class="MB MD TdCliBtn"><a href="javascript:copyMagentoShipAddrToBillAddr();"><img src="../IMAGEM/copia_20x20.png" name="btnMagentoCopyShipAddrToBillAddr" id="btnMagentoCopyShipAddrToBillAddr" title="Altera o endereço usando os dados do endereço de entrega obtidos do Magento" /></a></td>
	</tr>
</table>

<table cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;width:698px;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br />
<% end if %>

<% if blnHaRegistroBsp then %>
<table width="649" class="Q" cellspacing="0">
    <tr>
        <td width="100%" align="center" style="padding: 8px;">
            <span class="C" style="color: orange;">O formulário foi pré-preenchido conforme os dados do cadastro do cliente no sistema principal</span>
        </td>
    </tr>
</table>
<br />
<% end if %>

<!-- ************   CNPJ/IE OU CPF/RG/NASCIMENTO/SEXO  ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td width="210" align="left">
	<%if eh_cpf then s="CPF" else s="CNPJ"%>
	<p class="R"><%=s%></p><p class="C">
	<%	if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
			s=Trim("" & rs("cnpj_cpf"))
			s=cnpj_cpf_formata(s)
		else
			s=cnpj_cpf_formata(cnpj_cpf_selecionado)
			end if
	%>
	<input id="cnpj_cpf_selecionado" name="cnpj_cpf_selecionado" class="TA" value="<%=s%>" readonly size="22" style="text-align:center; color:#0000ff"></p></td>

<%if eh_cpf then%>
	<td class="MDE" width="210" align="left"><p class="R">RG</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("rg")) else s=""%>
		<input id="rg" name="rg" class="TA" type="text" maxlength="20" size="22" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.dt_nasc.focus(); filtra_nome_identificador();"></p></td>
	<td class="MD" align="left"><p class="R">NASCIMENTO</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=formata_data(rs("dt_nasc"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_customer_dob
				end if
		%>
		<input id="dt_nasc" name="dt_nasc" class="TA" type="text" maxlength="10" size="14" value="<%=s%>" onkeypress="if (digitou_enter(true) && isDate(this)) fCAD.sexo.focus(); filtra_data();" onblur="if (tem_info(this.value)) if (!isDate(this)) {alert('Data inválida!!');this.focus();}"></p></td>
	<td align="left"><p class="R">SEXO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("sexo")) else s=""%>
		<input id="sexo" name="sexo" class="TA" type="text" maxlength="1" size="2" value="<%=s%>" onkeypress="if (digitou_enter(true)) if (!tem_info(this.value)) fCAD.nome.focus(); else if (sexo_ok(this.value)) fCAD.nome.focus(); filtra_sexo();" onkeyup="this.value=ucase(this.value);"></p></td>

<%else%>
	<td class="MDE" width="215" align="left"><p class="R">IE</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("ie")) else s=""%>
		<input id="ie" name="ie" class="TA" type="text" maxlength="20" size="25" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.nome.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("contribuinte_icms_status")) else s=""%>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO then s_aux="checked" else s_aux=""%>
		<% intIdx = 0 %>
		<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Não</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Sim</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Isento</span></p></td>
<%end if%>
	</tr>
</table>

<!-- ************   PRODUTOR RURAL / CONTRIBUINTE ICMS / IE ************ -->
<%if eh_cpf then%>
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td align="left"><p class="R">PRODUTOR RURAL</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("produtor_rural_status")) else s=""%>
		<%if s=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO then s_aux="checked" else s_aux=""%>
		<% intIdx = 0 %>
		<input type="radio" id="rb_produtor_rural_nao" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fNEW.rb_produtor_rural[<%=Cstr(intIdx)%>].click();">Não</span>
		<%if s=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_produtor_rural_sim" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fNEW.rb_produtor_rural[<%=Cstr(intIdx)%>].click();">Sim</span></p></td>
	</tr>
</table>

<table width="649" class="QS" cellspacing="0" id="t_contribuinte_icms" onload="trataProdutorRural();">
	<tr>
	<td width="210" class="MD" align="left"><p class="R">IE</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("ie")) else s=""%>
		<input id="ie" name="ie" class="TA" type="text" maxlength="20" size="25" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.nome.focus(); filtra_nome_identificador();"></p></td>

	<td align="left"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("contribuinte_icms_status")) else s=""%>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO then s_aux="checked" else s_aux=""%>
		<% intIdx = 0 %>
		<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Não</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Sim</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Isento</span></p></td>
	</tr>
</table>
<% end if %>

<!-- ************   NOME  ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<%if eh_cpf then s="NOME" else s="RAZÃO SOCIAL"%>
	<td width="100%" align="left"><p class="R"><%=s%></p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then 
				s=Trim("" & rs("nome"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_customer_full_name
				end if
			%>
		<input id="nome" name="nome" class="TA" value="<%=s%>" maxlength="60" size="85" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("endereco"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_end_etg_endereco
				end if
			%>
		<input id="endereco" name="endereco" class="TA" value="<%=s%>" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   Nº/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">Nº</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("endereco_numero"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_end_etg_endereco_numero
				end if
			%>
		<input id="endereco_numero" name="endereco_numero" class="TA" value="<%=s%>" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("endereco_complemento"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_end_etg_complemento
				end if
			%>
		<input id="endereco_complemento" name="endereco_complemento" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">BAIRRO</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("bairro"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_end_etg_bairro
				end if
			%>
		<input id="bairro" name="bairro" class="TA" value="<%=s%>" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.cidade.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">CIDADE</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("cidade"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_end_etg_cidade
				end if
			%>
		<input id="cidade" name="cidade" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">UF</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("uf"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_end_etg_uf
				end if
			%>
		<input id="uf" name="uf" class="TA" value="<%=s%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) 
			<%if eh_cpf then Response.write "fCAD.ddd_res.focus();" else Response.Write "fCAD.ddd_com.focus();"%>" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
	<td width="50%" align="left">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="left"><p class="R">CEP</p><p class="C">
				<%	s=""
					if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
						s=Trim("" & rs("cep"))
					elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
						s = c_mag_end_etg_cep
						end if
					%>
				<input id="cep" name="cep" readonly tabindex=-1 class="TA" value="<%=cep_formata(s)%>" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) 
					<%if eh_cpf then Response.write "fCAD.ddd_res.focus();" else Response.Write "fCAD.ddd_com.focus();"%> 
					filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
			<td align="center" width="50%">
				<% if blnPesquisaCEPAntiga then %>
				<button type="button" name="bPesqCep" id="bPesqCep" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCep();">&nbsp;Pesquisar CEP&nbsp;</button>
				<% end if %>
				<% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
				<% if blnPesquisaCEPNova then %>
				<button type="button" name="bPesqCepNovo" id="bPesqCepNovo" style='width:130px;font-size:10pt;' class="Botao" onclick="exibeJanelaCEP_Cli();">&nbsp;Busca de CEP&nbsp;</button>
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
	</tr>
</table>

<!-- ************   TELEFONE RESIDENCIAL   ************ -->
<% if eh_cpf then %>
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("ddd_res"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_end_etg_telephone_ddd
				end if
			%>
		<input id="ddd_res" name="ddd_res" class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	<td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("tel_res"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_end_etg_telephone_numero
				end if
			%>
		<input id="tel_res" name="tel_res" class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	</tr>
	<tr>
	<td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("ddd_cel"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_end_etg_celular_ddd
				end if
			%>
		<input id="ddd_cel" name="ddd_cel" class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	<td align="left" class="MC"><p class="R">CELULAR</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("tel_cel"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = c_mag_end_etg_celular_numero
				end if
			%>
		<input id="tel_cel" name="tel_cel" class="TA" value="<%=telefone_formata(s)%>" maxlength="10" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ddd_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Número de celular inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	</tr>
</table>
<% end if %>
	
<!-- ************   TELEFONE COMERCIAL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<%
		s_ddd = ""
		s_tel = ""
		if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
			if Not eh_cpf then
				s_ddd = c_mag_end_etg_telephone_ddd
				s_tel = c_mag_end_etg_telephone_numero
			else
				s_ddd = c_mag_end_etg_fax_ddd
				s_tel = c_mag_end_etg_fax_numero
				end if
			end if
	%>
	<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("ddd_com"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = s_ddd
				end if
			%>
		<input id="ddd_com" name="ddd_com" class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	<%if eh_cpf then s=" COMERCIAL" else s=""%>
	<td class="MD" align="left"><p class="R">TELEFONE<%=s%></p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("tel_com"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = s_tel
				end if
			%>
		<input id="tel_com" name="tel_com" class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	<td align="left"><p class="R">RAMAL</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("ramal_com"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = ""
				end if
			%>
		<input id="ramal_com" name="ramal_com" class="TA" value="<%=s%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true))
			<%if Not eh_cpf then Response.Write "fCAD.ddd_com_2.focus();" else Response.Write "filiacao.focus();" %> filtra_numerico();"></p></td>
	</tr>
	<% if Not eh_cpf then %>
	<tr>
	    <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		<%
		s_ddd = ""
		s_tel = ""
		if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
			if c_mag_end_etg_celular_numero <> "" then
				s_ddd = c_mag_end_etg_celular_ddd
				s_tel = c_mag_end_etg_celular_numero
			else
				s_ddd = c_mag_end_etg_fax_ddd
				s_tel = c_mag_end_etg_fax_numero
				end if
			end if
		%>
	    <%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("ddd_com_2"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = s_ddd
				end if
			%>
	    <input id="ddd_com_2" name="ddd_com_2" class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!!');this.focus();}" /></p>  
	    </td>
	    <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
	    <%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("tel_com_2"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = s_tel
				end if
			%>
	    <input id="tel_com_2" name="tel_com_2" class="TA" value="<%=telefone_formata(s)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	    </td>
	    <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
	    <%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("ramal_com_2"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s = ""
				end if
			%>
	    <input id="ramal_com_2" name="ramal_com_2" class="TA" value="<%=s%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) <%if eh_cpf then Response.Write "fCAD.filiacao.focus();" else Response.Write "fCAD.contato.focus();"%> filtra_numerico();" /></p>
	    </td>
	</tr>
	<% end if %>
</table>

<% if eh_cpf then %>
<!-- ************   OBSERVAÇÃO (ANTIGO CAMPO FILIAÇÃO)   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">OBSERVAÇÃO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("filiacao")) else s=""%>
		<input id="filiacao" name="filiacao" class="TA" value="<%=s%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fCAD.email.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>
<% else %>
<!-- ************   CONTATO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">NOME DA PESSOA PARA CONTATO NA EMPRESA</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("contato")) else s=""%>
		<input id="contato" name="contato" class="TA" value="<%=s%>" maxlength="30" size="45" onkeypress="if (digitou_enter(true)) fCAD.email.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>
<% end if %>

<!-- ************   E-MAIL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">E-MAIL</p><p class="C">
		<%	s=""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s=Trim("" & rs("email"))
			elseif operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
				s=c_mag_email_identificado
				end if
			%>
		<input id="email" name="email" class="TA" value="<%=s%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fCAD.email_xml.focus(); filtra_email();"></p></td>
	    <input type="hidden" name="email_original" id="email_original" value="<%=s%>" />
    </tr>
</table>

<!-- ************   E-MAIL (XML)  ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">E-MAIL (XML)</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("email_xml")) else s=""%>
		<input id="email_xml" name="email_xml" class="TA" value="<%=s%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fCAD.obs_crediticias.focus(); filtra_email();"></p></td>
        <input type="hidden" name="email_xml_original" id="email_xml_original" value="<%=s%>" />
	</tr>
</table>

<!-- ************   OBS CREDITÍCIAS (INATIVO)   ************ -->

<table width="649" class="QS" cellspacing="0" style="display:none">
	<tr>
	<td width="100%" align="left"><p class="R">OBSERVAÇÕES CREDITÍCIAS</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("obs_crediticias")) else s=""%>
		<input id="obs_crediticias" name="obs_crediticias" class="TA" value="<%=s%>" maxlength="50" size="65" onkeypress="if (digitou_enter(true)) fCAD.midia.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   MÍDIA  (INATIVO) ************ -->
<table width="649" class="QS" cellspacing="0" style="display:none">
	<tr>
	<td width="100%" align="left"><p class="R">FORMA PELA QUAL CONHECEU A BONSHOP</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
			s=Trim("" & rs("midia"))
		else
			if CStr(loja)=CStr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then
			'	PRÉ-SELECIONA "INTERNET"
				s = "017"
			else
				s = ""
				end if
			end if
		%>
		<select id="midia" name="midia" style="margin-top:4pt; margin-bottom:4pt;">
			<%=lista_midia(s)%>
		</select>
	</tr>
</table>

<!-- ************   INDICADOR   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">INDICADOR</p>
	<% if blnCampoIndicadorEditavel then %>
		<% if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s_codigo=Trim("" & rs("indicador")) else s_codigo=""%>
		<p class="C">
		<select id="indicador" name="indicador" style="margin-top:4pt; margin-bottom:4pt; max-width:630px;">
			<% =strCampoSelectIndicadores %>
		</select>
		</p>
	<% else %>
		<%	s_codigo = ""
			s_descricao = ""
			s_codigo_e_descricao = ""
			if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
				s_codigo = Trim("" & rs("indicador"))
				if s_codigo <> "" then s_descricao = x_orcamentista_e_indicador(s_codigo)
				if (s_codigo <> "") And (s_descricao <> "") then s_codigo_e_descricao = s_codigo & " - " & s_descricao
				end if
		%>
		<p class="C">
		<input id="indicador" name="indicador" class="TA" value="<%=s_codigo_e_descricao%>" style="width:620px;" readonly />
		</p>
	<% end if %>
	</tr>
</table>


<!-- ************   REF BANCÁRIA   ************ -->
<%if blnCadRefBancaria then%>
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="c_RefBancariaBanco" id="c_RefBancariaBanco" value="">
<input type="hidden" name="c_RefBancariaAgencia" id="c_RefBancariaAgencia" value="">
<input type="hidden" name="c_RefBancariaConta" id="c_RefBancariaConta" value="">
<input type="hidden" name="c_RefBancariaDdd" id="c_RefBancariaDdd" value="">
<input type="hidden" name="c_RefBancariaTelefone" id="c_RefBancariaTelefone" value="">
<input type="hidden" name="c_RefBancariaContato" id="c_RefBancariaContato" value="">
	<% 
	if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
		s="SELECT * FROM t_CLIENTE_REF_BANCARIA WHERE (id_cliente='" & Trim("" & rs("id")) & "') ORDER BY ordem"
		set tRefBancaria = cn.Execute(s)
		end if
	%>
	
	<% for intCounter=1 to int_MAX_REF_BANCARIA_CLIENTE %>
		<%
		strRefBancariaBanco=""
		strRefBancariaAgencia=""
		strRefBancariaConta=""
		strRefBancariaDdd=""
		strRefBancariaTelefone=""
		strRefBancariaContato=""
		if (operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp) then
			if Not tRefBancaria.Eof then 
				strRefBancariaBanco=Trim("" & tRefBancaria("banco"))
				strRefBancariaAgencia=Trim("" & tRefBancaria("agencia"))
				strRefBancariaConta=Trim("" & tRefBancaria("conta"))
				strRefBancariaDdd=Trim("" & tRefBancaria("ddd"))
				strRefBancariaTelefone=Trim("" & tRefBancaria("telefone"))
				strRefBancariaContato=Trim("" & tRefBancaria("contato"))
				end if
			end if 
		%>
<% if Not eh_cpf then %>		
<br>

<table width="649" cellpadding="0" cellspacing="0">
	<tr>
		<td width="100%" align="left">
			<table width="100%" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">REFERÊNCIA BANCÁRIA<%if int_MAX_REF_BANCARIA_CLIENTE > 1 then Response.Write " (" & CStr(intCounter) & ")"%></p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td width="100%" class="MC" align="left">
						<p class="R">BANCO</p>
						<p class="C">
							<select name="c_RefBancariaBanco" id="c_RefBancariaBanco" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
							<%=banco_monta_itens_select(strRefBancariaBanco) %>
							</select>
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MD" align="left">
						<p class="R">AGÊNCIA</p>
						<p class="C">
							<input name="c_RefBancariaAgencia" id="c_RefBancariaAgencia" class="TA" maxlength="8" size="12" value="<%=strRefBancariaAgencia%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaConta[<%=CStr(intCounter)%>].focus(); filtra_agencia_bancaria();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">CONTA</p>
						<p class="C">
							<input name="c_RefBancariaConta" id="c_RefBancariaConta" class="TA" maxlength="12" value="<%=strRefBancariaConta%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaDdd[<%=CStr(intCounter)%>].focus(); filtra_conta_bancaria();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_RefBancariaDdd" id="c_RefBancariaDdd" class="TA" maxlength="2" size="4" value="<%=strRefBancariaDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefBancariaTelefone" id="c_RefBancariaTelefone" class="TA" maxlength="9" value="<%=telefone_formata(strRefBancariaTelefone)%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaContato[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">CONTATO</p>
						<p class="C">
							<input name="c_RefBancariaContato" id="c_RefBancariaContato" class="TA" maxlength="40"  style="width:600px;" value="<%=strRefBancariaContato%>" onkeypress="if (digitou_enter(true)) {if (<%=CStr(intCounter+1)%>==fCAD.c_RefBancariaAgencia.length) this.focus(); else fCAD.c_RefBancariaAgencia[<%=CStr(intCounter+1)%>].focus();} filtra_nome_identificador();">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<% end if %>
		<% 
		if (operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp) then
			if Not tRefBancaria.Eof then tRefBancaria.MoveNext
			end if
		%>
		
	<% next %>
	
	<% 
	if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
		tRefBancaria.Close
		end if
	%>
<%end if%>


<!-- ************   REF PROFISSIONAL   ************ -->
<%if blnCadRefProfissional then%>
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type=hidden name="c_RefProfNomeEmpresa" id="c_RefProfNomeEmpresa" value="">
<input type=hidden name="c_RefProfCargo" id="c_RefProfCargo" value="">
<input type=hidden name="c_RefProfDdd" id="c_RefProfDdd" value="">
<input type=hidden name="c_RefProfTelefone" id="c_RefProfTelefone" value="">
<input type=hidden name="c_RefProfPeriodoRegistro" id="c_RefProfPeriodoRegistro" value="">
<input type=hidden name="c_RefProfRendimentos" id="c_RefProfRendimentos" value="">
<input type=hidden name="c_RefProfCnpj" id="c_RefProfCnpj" value="">
	<% 
	if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
		s="SELECT * FROM t_CLIENTE_REF_PROFISSIONAL WHERE (id_cliente='" & Trim("" & rs("id")) & "') ORDER BY ordem"
		set tRefProfissional = cn.Execute(s)
		end if
	%>
	
	<% for intCounter=1 to int_MAX_REF_PROFISSIONAL_CLIENTE %>
		<%
		strRefProfNomeEmpresa=""
		strRefProfCargo=""
		strRefProfDdd=""
		strRefProfTelefone=""
		strRefProfPeriodoRegistro=""
		strRefProfRendimentos=""
		strRefProfCnpj=""
		if (operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp) then
			if Not tRefProfissional.Eof then 
				strRefProfNomeEmpresa=Trim("" & tRefProfissional("nome_empresa"))
				strRefProfCargo=Trim("" & tRefProfissional("cargo"))
				strRefProfDdd=Trim("" & tRefProfissional("ddd"))
				strRefProfTelefone=Trim("" & tRefProfissional("telefone"))
				strRefProfPeriodoRegistro=formata_data(tRefProfissional("periodo_registro"))
				strRefProfRendimentos=formata_moeda(tRefProfissional("rendimentos"))
				strRefProfCnpj=cnpj_cpf_formata(Trim("" & tRefProfissional("cnpj")))
				end if
			end if 
		%>

<table width="649" cellpadding="0" cellspacing="0" style="display:none">
	<tr>
		<td width="100%" align="left">
			<table width="100%" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">REFERÊNCIA PROFISSIONAL<%if int_MAX_REF_PROFISSIONAL_CLIENTE > 1 then Response.Write " (" & CStr(intCounter) & ")"%></p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MC MD" align="left">
						<p class="R">NOME DA EMPRESA</p>
						<p class="C">
							<input name="c_RefProfNomeEmpresa" id="c_RefProfNomeEmpresa" class="TA" maxlength="60"  style="width:450px;" value="<%=strRefProfNomeEmpresa%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfCnpj[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
						</p>
					</td>
					<td class="MC" align="left">
						<p class="R">CNPJ</p>
						<p class="C">
							<input name="c_RefProfCnpj" id="c_RefProfCnpj" class="TA" maxlength="18"  size="24" value="<%=strRefProfCnpj%>" onkeypress="if (digitou_enter(true) && cnpj_ok(this.value)) fCAD.c_RefProfCargo[<%=CStr(intCounter)%>].focus(); filtra_cnpj();" onblur="if (!cnpj_ok(this.value)) {alert('CNPJ inválido!!'); this.focus();} else this.value=cnpj_formata(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MD" align="left">
						<p class="R">CARGO</p>
						<p class="C">
							<input name="c_RefProfCargo" id="c_RefProfCargo" class="TA" maxlength="40" style="width:350px;" value="<%=strRefProfCargo%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfDdd[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_RefProfDdd" id="c_RefProfDdd" class="TA" maxlength="2" size=4 value="<%=strRefProfDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefProfTelefone" id="c_RefProfTelefone" class="TA" maxlength="9" value="<%=telefone_formata(strRefProfTelefone)%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfPeriodoRegistro[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MD" width="50%" align="left">
						<p class="R">REGISTRADO DESDE (DD/MM/AAAA)</p>
						<p class="C">
							<input name="c_RefProfPeriodoRegistro" id="c_RefProfPeriodoRegistro" class="TA" maxlength="10" value="<%=strRefProfPeriodoRegistro%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfRendimentos[<%=CStr(intCounter)%>].focus(); filtra_data();" onblur="if (!isDate(this)) {alert('Data inválida!!');this.focus();}">
						</p>
					</td>
					<td width="50%" align="left">
						<p class="R">RENDIMENTOS (<%=SIMBOLO_MONETARIO%>)</p>
						<p class="C">
							<input name="c_RefProfRendimentos" id="c_RefProfRendimentos" class="TA" maxlength="18" value="<%=strRefProfRendimentos%>" onkeypress="if (digitou_enter(true)) {if (<%=CStr(intCounter+1)%>==fCAD.c_RefProfNomeEmpresa.length) this.focus(); else fCAD.c_RefProfNomeEmpresa[<%=CStr(intCounter+1)%>].focus();} filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
		<% 
		if (operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp) then
			if Not tRefProfissional.Eof then tRefProfissional.MoveNext
			end if
		%>
		
	<% next %>
	
	<% 
	if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
		tRefProfissional.Close
		end if
	%>
<%end if%>


<!-- ************   REF COMERCIAL   ************ -->
<%if blnCadRefComercial then%>
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="c_RefComercialNomeEmpresa" id="c_RefComercialNomeEmpresa" value="">
<input type="hidden" name="c_RefComercialContato" id="c_RefComercialContato" value="">
<input type="hidden" name="c_RefComercialDdd" id="c_RefComercialDdd" value="">
<input type="hidden" name="c_RefComercialTelefone" id="c_RefComercialTelefone" value="">
	<% 
	if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
		s="SELECT * FROM t_CLIENTE_REF_COMERCIAL WHERE (id_cliente='" & Trim("" & rs("id")) & "') ORDER BY ordem"
		set tRefComercial = cn.Execute(s)
		end if
	%>
	
	<% for intCounter=1 to int_MAX_REF_COMERCIAL_CLIENTE %>
		<%
		strRefComercialNomeEmpresa=""
		strRefComercialContato=""
		strRefComercialDdd=""
		strRefComercialTelefone=""
		if (operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp) then
			if Not tRefComercial.Eof then 
				strRefComercialNomeEmpresa=Trim("" & tRefComercial("nome_empresa"))
				strRefComercialContato=Trim("" & tRefComercial("contato"))
				strRefComercialDdd=Trim("" & tRefComercial("ddd"))
				strRefComercialTelefone=Trim("" & tRefComercial("telefone"))
				end if
			end if 
		%>
<br>
<table width="649" cellpadding="0" cellspacing="0">
	<tr>
		<td width="100%" align="left">
			<table width="100%" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">REFERÊNCIA COMERCIAL<%if int_MAX_REF_COMERCIAL_CLIENTE > 1 then Response.Write " (" & CStr(intCounter) & ")"%></p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td colspan="4" class="MC" align="left">
						<p class="R">NOME DA EMPRESA</p>
						<p class="C">
							<input name="c_RefComercialNomeEmpresa" id="c_RefComercialNomeEmpresa" class="TA" maxlength="60"  style="width:600px;" value="<%=strRefComercialNomeEmpresa%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefComercialContato[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MD" align="left">
						<p class="R">CONTATO</p>
						<p class="C">
							<input name="c_RefComercialContato" id="c_RefComercialContato" class="TA" maxlength="40" value="<%=strRefComercialContato%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefComercialDdd[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_RefComercialDdd" id="c_RefComercialDdd" class="TA" maxlength="2" size="4" value="<%=strRefComercialDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefComercialTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefComercialTelefone" id="c_RefComercialTelefone" class="TA" maxlength="9" value="<%=telefone_formata(strRefComercialTelefone)%>" onkeypress="if (digitou_enter(true)) {if (<%=CStr(intCounter+1)%>==fCAD.c_RefComercialNomeEmpresa.length) this.focus(); else fCAD.c_RefComercialNomeEmpresa[<%=CStr(intCounter+1)%>].focus();} filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
		<% 
		if (operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp) then
			if Not tRefComercial.Eof then tRefComercial.MoveNext
			end if
		%>
		
	<% next %>
	
	<% 
	if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then
		tRefComercial.Close
		end if
	%>
<%end if%>


<!-- ************   PJ: DADOS DO SÓCIO MAJORITÁRIO (INATIVO)  ************ -->
<%if blnCadSocioMaj then%>
<br>
<table width="649" cellpadding="0" cellspacing="0" style="display:none">
	<tr>
		<td width="100%" align="left">
			<table width="100%" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">DADOS DO SÓCIO MAJORITÁRIO</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
				<td class="MC MD" width="85%" align="left"><p class="R">NOME</p><p class="C">
					<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("SocMaj_Nome")) else s=""%>
					<input id="c_SocioMajNome" name="c_SocioMajNome" class="TA" value="<%=s%>" maxlength="60" size="61" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajCpf.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"></p></td>
				<td class="MC" align="left"><p class="R">CPF</p><p class="C">
					<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("SocMaj_CPF")) else s=""%>
					<input id="c_SocioMajCpf" name="c_SocioMajCpf" class="TA" value="<%=cnpj_cpf_formata(s)%>" maxlength="14" size="15" onkeypress="if (digitou_enter(true) && cpf_ok(this.value)) fCAD.c_SocioMajBanco.focus(); filtra_numerico();" onblur="if (!cpf_ok(this.value)) {alert('CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);"></p></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">BANCO</p>
						<p class="C">
							<select name="c_SocioMajBanco" id="c_SocioMajBanco" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
							<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("SocMaj_banco")) else s=""%>
							<%=banco_monta_itens_select(s) %>
							</select>
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MD" align="left">
						<p class="R">AGÊNCIA</p>
						<p class="C">
							<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("SocMaj_agencia")) else s=""%>
							<input name="c_SocioMajAgencia" id="c_SocioMajAgencia" class="TA" maxlength="8" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajConta.focus(); filtra_agencia_bancaria();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">CONTA</p>
						<p class="C">
							<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("SocMaj_conta")) else s=""%>
							<input name="c_SocioMajConta" id="c_SocioMajConta" class="TA" maxlength="12" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajDdd.focus(); filtra_conta_bancaria();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("SocMaj_ddd")) else s=""%>
							<input name="c_SocioMajDdd" id="c_SocioMajDdd" class="TA" maxlength="2" size="4" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajTelefone.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("SocMaj_telefone")) else s=""%>
							<input name="c_SocioMajTelefone" id="c_SocioMajTelefone" class="TA" maxlength="9" value="<%=telefone_formata(s)%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajContato.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">CONTATO</p>
						<p class="C">
							<%if operacao_selecionada=OP_CONSULTA Or blnHaRegistroBsp then s=Trim("" & rs("SocMaj_contato")) else s=""%>
							<input name="c_SocioMajContato" id="c_SocioMajContato" class="TA" maxlength="40"  style="width:600px;" value="<%=s%>" onkeypress="filtra_nome_identificador();">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%end if%>

</form>


<!-- ************   FORM PARA OPÇÃO DE CADASTRAR NOVO PEDIDO?  ************ -->
<% if (operacao_selecionada = OP_CONSULTA) And operacao_permitida(OP_LJA_CADASTRA_NOVO_PEDIDO, s_lista_operacoes_permitidas) then %>
	<% if blnLojaHabilitadaProdCompostoECommerce then
		s = "PedidoNovoProdCompostoMask.asp"
	else
		s = "pedidonovo.asp"
		end if %>
	<form action="<%=s%>" method="post" id="fNEW" name="fNEW" onsubmit="if (!fNEWConcluir(fNEW)) return false">
	<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
	<INPUT type="hidden" name='cliente_selecionado' id="cliente_selecionado" value='<%=id_cliente%>'>
	<INPUT type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=OP_INCLUI%>'>
	<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />
	<input type="hidden" name="operacao_origem" id="operacao_origem" value="<%=operacao_origem%>" />
	<input type="hidden" name="id_magento_api_pedido_xml" id="id_magento_api_pedido_xml" value="<%=id_magento_api_pedido_xml%>" />
	<input type="hidden" name="c_numero_magento" id="c_numero_magento" value="<%=c_numero_magento%>" />
	<input type="hidden" name="operationControlTicket" id="operationControlTicket" value="<%=operationControlTicket%>" />
	<input type="hidden" name="sessionToken" id="sessionToken" value="<%=sessionToken%>" />

<!-- ************   ENDEREÇO DE ENTREGA: S/N   ************ -->
<br>
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td align="left">
		<p class="R">ENDEREÇO DE ENTREGA</p><p class="C">
			<% intIdx = 0 %>
			<input type="radio" id="rb_end_entrega_nao" name="rb_end_entrega" value="N" onclick="Disabled_True(fNEW);"><span class="C" style="cursor:default" onclick="fNEW.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_True(fNEW);">O mesmo endereço do cadastro</span>
			<% intIdx = intIdx + 1 %>
			<br><input type="radio" id="rb_end_entrega_sim" name="rb_end_entrega" value="S" onclick="Disabled_False(fNEW);"><span class="C" style="cursor:default" onclick="fNEW.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_False(fNEW);">Outro endereço</span>
		</p>
		</td>
		<td style="width:40px;text-align:right;vertical-align:top;">
			<a href="javascript:copyMagentoShipAddrToShipAddr();"><img src="../IMAGEM/copia_20x20.png" name="btnMagentoCopyShipAddrToShipAddr" id="btnMagentoCopyShipAddrToShipAddr" title="Altera o endereço usando os dados do endereço de entrega obtidos do Magento" /></a>
		</td>
	</tr>
</table>


<!--  ************  TIPO DO ENDEREÇO DE ENTREGA: PF/PJ (SOMENTE SE O CLIENTE FOR PJ)   ************ -->
<input type="hidden" name="st_memorizacao_completa_enderecos" id="st_memorizacao_completa_enderecos" value="1" />

<%if Not eh_cpf then%>
<table width="649" class="QS Habilitar_EndEtg_outroendereco" cellspacing="0">
	<tr>
		<td align="left">
		<p class="R">TIPO</p><p class="C">
			<input type="radio" name="EndEtg_tipo_pessoa" value="PJ" onclick="trocarEndEtgTipoPessoa(null);" checked>
			<span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PJ');">Pessoa Jurídica</span>
			&nbsp;
			<input type="radio" name="EndEtg_tipo_pessoa" value="PF" onclick="trocarEndEtgTipoPessoa(null);">
			<span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PF');">Pessoa Física</span>
		</p>
		</td>
	</tr>
</table>

        <!-- ************   PJ: CNPJ/CONTRIBUINTE ICMS/IE - DO ENDEREÇO DE ENTREGA DE PJ ************ -->
        <!-- ************   PF: CPF/RG/PRODUTOR RURAL/CONTRIBUINTE ICMS/IE - DO ENDEREÇO DE ENTREGA DE PJ  ************ -->
        <!-- fizemos dois conjuntos diferentes de campos porque a ordem é muito diferente -->

<input type="hidden" name="EndEtg_cnpj_cpf" />
<input type="hidden" name="EndEtg_ie" />
<input type="hidden" name="EndEtg_contribuinte_icms_status" />
<input type="hidden" name="EndEtg_rg" />
<input type="hidden" name="EndEtg_produtor_rural_status" />


<table width="649" class="QS Habilitar_EndEtg_outroendereco Mostrar_EndEtg_pj" cellspacing="0">
	<tr>
		<td width="210" align="left">
	<p class="R">CNPJ</p><p class="C">
	<input name="EndEtg_cnpj_cpf_PJ" class="TA" value="" size="22" style="text-align:center; color:#0000ff"></p></td>

	<td class="MDE" width="215" align="left"><p class="R">IE</p><p class="C">
		<input name="EndEtg_ie_PJ" class="TA" type="text" maxlength="20" size="25" value="" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_Nome.focus(); filtra_nome_identificador();"></p></td>

	<td align="left" class="Mostrar_EndEtg_contribuinte_icms_PJ"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		<input type="radio" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">Não</span>
		<input type="radio" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>');">Sim</span>
		<input type="radio" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>');">Isento</span></p></td>
	</tr>
</table>

<table width="649" class="QS Habilitar_EndEtg_outroendereco Mostrar_EndEtg_pf" cellspacing="0">
	<tr>
		<td width="210" align="left">
	<p class="R">CPF</p><p class="C">
	<input name="EndEtg_cnpj_cpf_PF" class="TA" value="" size="22" style="text-align:center; color:#0000ff"></p></td>

	<td class="MDE" width="210" align="left"><p class="R">RG</p><p class="C">
		<input name="EndEtg_rg_PF" class="TA" type="text" maxlength="20" size="22" value="" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_produtor_rural_status_PF.focus(); filtra_nome_identificador();"></p></td>


	<td align="left" ><p class="R">PRODUTOR RURAL</p><p class="C">
		<input type="radio" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>');">Não</span>
		<input type="radio" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>')">Sim</span></p></td>
	</tr>
</table>

<table width="649" class="QS Habilitar_EndEtg_outroendereco Mostrar_EndEtg_pf Mostrar_EndEtg_contribuinte_icms_PF" cellspacing="0">
	<tr>
	<td width="210" align="left"><p class="R">IE</p><p class="C">
		<input name="EndEtg_ie_PF" class="TA" type="text" maxlength="20" size="25" value="" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_Nome.focus(); filtra_nome_identificador();"></p></td>

	<td align="left" class="ME" ><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		<input type="radio" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">Não</span>
		<input type="radio" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>');">Sim</span>
		<input type="radio" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>');">Isento</span></p></td>
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



<!-- ************   ENDEREÇO DE ENTREGA: ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0">
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
				<button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCepEndEtg();">&nbsp;Pesquisar CEP&nbsp;</button>
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

<%if Not eh_cpf then%>

        
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
		<input id="EndEtg_tel_cel" name="EndEtg_tel_cel" class="TA" value="" maxlength="10" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fNEW.EndEtg_obs.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Número de celular inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
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
	    <input id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" class="TA" value="" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_email.focus(); filtra_numerico();" /></p>
	    </td>
	</tr>
</table>


<!-- ************   ENDEREÇO DE ENTREGA: E-MAIL   ************ -->
<table width="649" class="QS Mostrar_EndEtg_pj Habilitar_EndEtg_outroendereco" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">E-MAIL</p><p class="C">
		<input id="EndEtg_email" name="EndEtg_email" class="TA" value="" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_email_xml.focus(); filtra_email();"></p></td>
    </tr>
</table>

<!-- ************   ENDEREÇO DE ENTREGA: E-MAIL (XML)  ************ -->
<table width="649" class="QS Mostrar_EndEtg_pj Habilitar_EndEtg_outroendereco" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">E-MAIL (XML)</p><p class="C">
		<input id="EndEtg_email_xml" name="EndEtg_email_xml" class="TA" value="" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fNEW.EndEtg_obs.focus(); filtra_email();"></p></td>
	</tr>
</table>



<% end if %>


<!-- ************   JUSTIFIQUE O ENDEREÇO   ************ -->
<table id="obs_endereco" width="649" class="QS" cellspacing="0">
	<tr >
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


<!-- ************   SEPARADOR   ************ -->
<table width="698" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>

<% if operacao_selecionada = OP_INCLUI then %>
<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="left"><a href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="left">&nbsp;</td>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaCliente(fCAD)" title="atualiza o cadastro deste cliente">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
<% else %>
	<% if blnEdicaoBloqueada then %>
		<% if operacao_permitida(OP_LJA_CADASTRA_NOVO_PEDIDO, s_lista_operacoes_permitidas) then %>
		<table class="notPrint" width="649" cellspacing="0">
		<tr>
			<td align="left"><a href="javascript:history.back();" title="volta para a página anterior">
				<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
			</td>
			<td align="left">&nbsp;</td>
			<td align="right"><div name="dPEDIDO" id="dPEDIDO">
				<a name="bPEDIDO" id="bPEDIDO" href="javascript:fNEWConcluir(fNEW);" title="cadastra um novo pedido para este cliente">
				<img src="../botao/pedido.gif" width="176" height="55" border="0"></a></div>
			</td>
		</tr>
		</table>
		<% else %>
		<table class="notPrint" width="649" cellspacing="0">
		<tr>
			<td align="center"><a href="javascript:history.back();" title="volta para a página anterior">
				<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
			</td>
		</tr>
		</table>
		<% end if %>
	<% else %>
		<table class="notPrint" width="649" cellspacing="0">
		<tr>
			<td align="left"><a href="javascript:history.back();" title="volta para a página anterior">
				<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
			</td>
			<td align="center"><div name="dREMOVE" id="dREMOVE">
				<a href="javascript:RemoveCliente(fCAD);" title="remove o cliente cadastrado">
				<img src="../botao/remover.gif" width="176" height="55" border="0"></a></div>
			</td>
			<td align="right"><div name="dATUALIZA" id="dATUALIZA">
				<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaCliente(fCAD)" title="atualiza o cadastro deste cliente">
				<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
			</td>
		</tr>
		<% if operacao_permitida(OP_LJA_CADASTRA_NOVO_PEDIDO, s_lista_operacoes_permitidas) then %>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td align="right">
				<div name="dPEDIDO" id="dPEDIDO">
				<a name="bPEDIDO" id="bPEDIDO" href="javascript:fNEWConcluir(fNEW);" title="cadastra um novo pedido para este cliente">
				<img src="../botao/pedido.gif" width="176" height="55" border="0"></a></div>
			</td>
		</tr>
		<% end if %>
		</table>
	<% end if %>
<% end if %>

</center>
</body>

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

	rs.Close
	set rs = nothing
	
	cn.Close
	set cn = nothing

    if ID_AMBIENTE = ID_AMBIENTE__AT And operacao_selecionada = OP_INCLUI then
        cnBsp.Close
	    set cnBsp = nothing
        end if
%>