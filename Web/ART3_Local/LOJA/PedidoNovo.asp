<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================
'	  P E D I D O N O V O . A S P
'     ===========================
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


	On Error GoTo 0
	Err.Clear

	dim i, j, s, usuario, loja, cliente_selecionado, r_cliente, strAux, msg_erro
	dim idxSelecionado, blnAchou

	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	cliente_selecionado = Trim(request("cliente_selecionado"))
	if (cliente_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	dim intIdx, intIdxRA

	dim cn, tMAP_XML, tOI
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	set r_cliente = New cl_CLIENTE
	if Not x_cliente_bd(cliente_selecionado, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if Trim(r_cliente.endereco_numero) = "" then
		Response.Redirect("aviso.asp?id=" & ERR_CAD_CLIENTE_ENDERECO_NUMERO_NAO_PREENCHIDO)
	elseif Len(Trim(r_cliente.endereco)) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
		Response.Redirect("aviso.asp?id=" & ERR_CAD_CLIENTE_ENDERECO_EXCEDE_TAMANHO_MAXIMO)
		end if
	
	dim blnLojaHabilitadaProdCompostoECommerce
	blnLojaHabilitadaProdCompostoECommerce = isLojaHabilitadaProdCompostoECommerce(loja)

	dim rb_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento
	dim EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs
	rb_end_entrega = Trim(Request.Form("rb_end_entrega"))
	EndEtg_endereco = Trim(Request.Form("EndEtg_endereco"))
	EndEtg_endereco_numero = Trim(Request.Form("EndEtg_endereco_numero"))
	EndEtg_endereco_complemento = Trim(Request.Form("EndEtg_endereco_complemento"))
	EndEtg_bairro = Trim(Request.Form("EndEtg_bairro"))
	EndEtg_cidade = Trim(Request.Form("EndEtg_cidade"))
	EndEtg_uf = Trim(Request.Form("EndEtg_uf"))
	EndEtg_cep = Trim(Request.Form("EndEtg_cep"))
	EndEtg_obs = Trim(Request.Form("EndEtg_obs"))

	dim alerta
	alerta = ""
	
	dim s_nome_cliente, c_mag_cpf_cnpj_identificado, c_mag_installer_document
	dim operacao_origem, c_numero_magento, operationControlTicket, sessionToken, id_magento_api_pedido_xml
	operacao_origem = Trim(Request("operacao_origem"))
	c_numero_magento = ""
	operationControlTicket = ""
	sessionToken = ""
	id_magento_api_pedido_xml = ""
	s_nome_cliente = ""
	c_mag_cpf_cnpj_identificado = ""
	c_mag_installer_document = ""
	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		c_numero_magento = Trim(Request("c_numero_magento"))
		operationControlTicket = Trim(Request("operationControlTicket"))
		sessionToken = Trim(Request("sessionToken"))
		id_magento_api_pedido_xml = Trim(Request("id_magento_api_pedido_xml"))
		end if

	dim blnMagentoPedidoComIndicador, sListaLojaMagentoPedidoComIndicador, vLoja, rParametro
	dim sIdInstalador, sNomeInstalador, sIdVendedor, sNomeVendedor
	dim percCommissionValue, percCommissionDiscount
	blnMagentoPedidoComIndicador = False
	sListaLojaMagentoPedidoComIndicador = ""
	sIdInstalador = ""
	sNomeInstalador = ""
	sIdVendedor = ""
	sNomeVendedor = ""
	percCommissionValue = 0
	percCommissionDiscount = 0

	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		If Not cria_recordset_otimista(tMAP_XML, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
		
		set rParametro = get_registro_t_parametro(ID_PARAMETRO_MagentoPedidoComIndicadorListaLojaErp)
		sListaLojaMagentoPedidoComIndicador = Trim("" & rParametro.campo_texto)
		if sListaLojaMagentoPedidoComIndicador <> "" then
			vLoja = Split(sListaLojaMagentoPedidoComIndicador, ",")
			for i=LBound(vLoja) to UBound(vLoja)
				if Trim("" & vLoja(i)) = loja then
					blnMagentoPedidoComIndicador = True
					exit for
					end if
				next
			end if
		end if
	
	if alerta = "" then
		if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
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
				c_mag_installer_document = retorna_so_digitos(Trim("" & tMAP_XML("installer_document")))
				percCommissionValue = tMAP_XML("commission_value")
				percCommissionDiscount = tMAP_XML("commission_discount")

				if blnMagentoPedidoComIndicador then
					if c_mag_installer_document = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "O pedido Magento nº " & c_numero_magento & " não informa o CPF/CNPJ do instalador!"
					else
						If Not cria_recordset_otimista(tOI, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
						s = "SELECT " & _
								"*" & _
							" FROM t_ORCAMENTISTA_E_INDICADOR" & _
							" WHERE" & _
								" (cnpj_cpf = '" & retorna_so_digitos(c_mag_installer_document) & "')" & _
								" AND (Convert(smallint, loja) = " & loja & ")" & _
								" AND (status = 'A')"
						if tOI.State <> 0 then tOI.Close
						tOI.open s, cn
						if tOI.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "O pedido Magento nº " & c_numero_magento & " especifica o instalador com CPF/CNPJ " & cnpj_cpf_formata(c_mag_installer_document) & " que não foi localizado no banco de dados (loja: " & loja & ")!"
						else
							sIdInstalador = Trim("" & tOI("apelido"))
							sNomeInstalador = Trim("" & tOI("razao_social_nome"))
							sIdVendedor = Trim("" & tOI("vendedor"))
							sNomeVendedor = Trim("" & x_usuario (sIdVendedor))

							'VERIFICA SE HÁ MAIS DE UM INDICADOR CADASTRADO
							tOI.MoveNext
							if Not tOI.Eof then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Há mais de um indicador cadastrado com o CPF/CNPJ " & cnpj_cpf_formata(c_mag_installer_document) & " para a loja " & loja
								end if
							end if
						if tOI.State <> 0 then tOI.Close
						set tOI = nothing
						end if
					end if
				end if 'if tMAP_XML.Eof
			end if 'if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO
		end if 'if alerta = ""

	if Trim("" & r_cliente.cep) <> "" then
		if Len(retorna_so_digitos(Trim("" & r_cliente.cep))) < 8 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O CEP do cadastro do cliente está incompleto (CEP: " & Trim("" & r_cliente.cep) & ")"
			end if
		end if

	if rb_end_entrega = "S" then
		if EndEtg_cep <> "" then
			if Len(retorna_so_digitos(EndEtg_cep)) < 8 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O CEP do endereço de entrega está incompleto (CEP: " & EndEtg_cep & ")"
				end if
			end if
		end if

	dim intQtdeIndicadores, strCampoSelectIndicadores, strJsScriptArrayIndicadores
	intQtdeIndicadores = 0
	strCampoSelectIndicadores = ""
	strJsScriptArrayIndicadores = ""
	call indicadores_monta_itens_select(Null, strCampoSelectIndicadores, strJsScriptArrayIndicadores)
	
'	CONSISTÊNCIAS P/ EMISSÃO DE NFe
	dim s_lista_sugerida_municipios
	dim v_lista_sugerida_municipios
	dim iCounterLista, iNumeracaoLista
	dim s_tabela_municipios_IBGE
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
	'	DDD VÁLIDO?
		if Not ddd_ok(r_cliente.ddd_res) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			alerta = alerta & "DDD do telefone residencial é inválido!!"
			end if
			
		if Not ddd_ok(r_cliente.ddd_com) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			alerta = alerta & "DDD do telefone comercial é inválido!!"
			end if
			
	'	I.E. É VÁLIDA?
		if (r_cliente.contribuinte_icms_status = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
			if Not isInscricaoEstadualValida(r_cliente.ie, r_cliente.uf) then
				if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				alerta=alerta & "Corrija a IE (Inscrição Estadual) com um número válido!!" & _
						"<br>" & "Certifique-se de que a UF informada corresponde à UF responsável pelo registro da IE."
				end if
			end if

	'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
		if Not consiste_municipio_IBGE_ok(r_cliente.cidade, r_cliente.uf, s_lista_sugerida_municipios, msg_erro) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			if msg_erro <> "" then
				alerta = alerta & msg_erro
			else
				alerta = alerta & "Município '" & r_cliente.cidade & "' não consta na relação de municípios do IBGE para a UF de '" & r_cliente.uf & "'!!"
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
								"			<p class='N'>" & "Relação de municípios de '" & r_cliente.uf & "' que se iniciam com a letra '" & Ucase(left(r_cliente.cidade,1)) & "'" & "</p>" & chr(13) & _
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

	if alerta = "" then
		if rb_end_entrega = "S" then
		'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
			if Not consiste_municipio_IBGE_ok(EndEtg_cidade, EndEtg_uf, s_lista_sugerida_municipios, msg_erro) then
				if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				if msg_erro <> "" then
					alerta = alerta & msg_erro
				else
					alerta = alerta & "Município '" & EndEtg_cidade & "' não consta na relação de municípios do IBGE para a UF de '" & EndEtg_uf & "'!!"
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
									"			<p class='N'>" & "Relação de municípios de '" & EndEtg_uf & "' que se iniciam com a letra '" & Ucase(left(EndEtg_cidade,1)) & "'" & "</p>" & chr(13) & _
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
				end if 'if Not consiste_municipio_IBGE_ok()
			end if 'if rb_end_entrega = "S"
		end if 'if alerta = ""
	
	dim rs, rs2
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim strSql
	dim s_fabricante, s_produto, s_qtde, n_qtde, s_preco_lista, s_descricao
	dim n, intIdxProduto
	dim vProduto
	redim vProduto(0)
	set vProduto(0) = New cl_ITEM_PEDIDO
	vProduto(0).qtde = 0

	if alerta = "" then
		'VERIFICA SE O MESMO CÓDIGO FOI DIGITADO REPETIDO EM VÁRIAS LINHAS
		if blnLojaHabilitadaProdCompostoECommerce then
			dim vDuplic
			redim vDuplic(0)
			set vDuplic(0) = New cl_ITEM_PEDIDO

			n = Request.Form("c_produto").Count
			for i = 1 to n
				s_fabricante = Trim(Request.Form("c_fabricante")(i))
				s_fabricante = normaliza_codigo(s_fabricante, TAM_MIN_FABRICANTE)
				s_produto = Trim(Request.Form("c_produto")(i))
				s_produto = normaliza_codigo(s_produto, TAM_MIN_PRODUTO)
				s_qtde = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s_qtde) then n_qtde = CLng(s_qtde) else n_qtde = 0
				if Trim("" & vDuplic(UBound(vDuplic)).produto) <> "" then
					redim preserve vDuplic(UBound(vDuplic)+1)
					set vDuplic(UBound(vDuplic)) = New cl_ITEM_PEDIDO
					end if
				vDuplic(UBound(vDuplic)).fabricante = s_fabricante
				vDuplic(UBound(vDuplic)).produto = s_produto
				vDuplic(UBound(vDuplic)).qtde = n_qtde
				next

			for i = LBound(vDuplic) to UBound(vDuplic)
				if Trim("" & vDuplic(i).produto) <> "" then
					for j = LBound(vDuplic) to (i-1)
						if (vDuplic(i).fabricante = vDuplic(j).fabricante) And (vDuplic(i).produto = vDuplic(j).produto) then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto " & vDuplic(i).produto & " do fabricante " & vDuplic(i).fabricante & ": linha " & renumera_com_base1(LBound(vDuplic),i) & " repete o mesmo produto da linha " & renumera_com_base1(LBound(vDuplic),j)
							exit for
							end if
						next
					end if
				next
			end if
		end if

	if alerta = "" then
		if blnLojaHabilitadaProdCompostoECommerce then
			n = Request.Form("c_produto").Count
			for i = 1 to n
				s_fabricante = Trim(Request.Form("c_fabricante")(i))
				s_fabricante = normaliza_codigo(s_fabricante, TAM_MIN_FABRICANTE)
				s_produto = Trim(Request.Form("c_produto")(i))
				s_produto = normaliza_codigo(s_produto, TAM_MIN_PRODUTO)
				s_qtde = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s_qtde) then n_qtde = CLng(s_qtde) else n_qtde = 0
			'	INFORMOU APENAS O CÓDIGO DO PRODUTO NA TELA ANTERIOR
			'	TENTA RECUPERAR O CÓDIGO DO FABRICANTE (VERIFICANDO SE HÁ AMBIGUIDADE)
				if (s_fabricante = "") And (s_produto <> "") then
				'	VERIFICA SE É PRODUTO COMPOSTO
					strSql = "SELECT " & _
								"*" & _
							" FROM t_EC_PRODUTO_COMPOSTO t_EC_PC" & _
							" WHERE" & _
								" (produto_composto = '" & s_produto & "')"
					if rs.State <> 0 then rs.Close
					rs.Open strSql, cn
				'	É PRODUTO COMPOSTO
					if Not rs.Eof then
						s_fabricante = Trim("" & rs("fabricante_composto"))
						rs.MoveNext
						if Not rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Há mais de um produto composto com o código '" & s_produto & "'!!<br />Informe o código do fabricante para resolver a ambiguidade!!"
							end if
						if alerta = "" then
							strSql = "SELECT " & _
										"*" & _
									" FROM t_EC_PRODUTO_COMPOSTO_ITEM t_EC_PCI" & _
									" WHERE" & _
										" (fabricante_composto = '" & s_fabricante & "')" & _
										" AND (produto_composto = '" & s_produto & "')" & _
										" AND (excluido_status = 0)" & _
									" ORDER BY" & _
										" sequencia"
							if rs.State <> 0 then rs.Close
							rs.Open strSql, cn
							do while Not rs.Eof
								strSql = "SELECT " & _
											"*" & _
										" FROM t_PRODUTO tP" & _
											" INNER JOIN t_PRODUTO_LOJA tPL ON (tP.fabricante = tPL.fabricante) AND (tP.produto = tPL.produto)" & _
										" WHERE" & _
											" (tP.fabricante = '" & Trim("" & rs("fabricante_item")) & "')" & _
											" AND (tP.produto = '" & Trim("" & rs("produto_item")) & "')" & _
											" AND (loja = '" & loja & "')"
								if rs2.State <> 0 then rs2.Close
								rs2.Open strSql, cn
								if rs2.Eof then
									alerta=texto_add_br(alerta)
									alerta=alerta & "O produto (" & Trim("" & rs("fabricante_item")) & ")" & Trim("" & rs("produto_item")) & " não está disponível para a loja " & loja & "!!"
								else
									blnAchou = False
									idxSelecionado = -1
									for j=LBound(vProduto) to UBound(vProduto)
										if (Trim("" & vProduto(j).fabricante) = Trim("" & rs("fabricante_item"))) And (Trim("" & vProduto(j).produto) = Trim("" & rs("produto_item"))) then
											blnAchou = True
											idxSelecionado = j
											exit for
											end if
										next

									if Not blnAchou then
										if Trim(vProduto(ubound(vProduto)).produto) <> "" then
											redim preserve vProduto(ubound(vProduto)+1)
											set vProduto(ubound(vProduto)) = New cl_ITEM_PEDIDO
											vProduto(ubound(vProduto)).qtde = 0
											end if
										idxSelecionado = ubound(vProduto)
										end if

									with vProduto(idxSelecionado)
										.fabricante = Trim("" & rs("fabricante_item"))
										.produto = Trim("" & rs("produto_item"))
										.qtde = .qtde + (n_qtde * rs("qtde"))
										.preco_lista = rs2("preco_lista")
										.descricao = Trim("" & rs2("descricao"))
										.descricao_html = Trim("" & rs2("descricao_html"))
										end with
									end if
								rs.MoveNext
								loop
							end if
				'	É PRODUTO NORMAL
					else
						strSql = "SELECT " & _
									"*" & _
								" FROM t_PRODUTO tP" & _
									" INNER JOIN t_PRODUTO_LOJA tPL ON (tP.fabricante = tPL.fabricante) AND (tP.produto = tPL.produto)" & _
								" WHERE" & _
									" (tP.produto = '" & s_produto & "')" & _
									" AND (loja = '" & loja & "')"
						if rs.State <> 0 then rs.Close
						rs.Open strSql, cn
						if rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto '" & s_produto & "' não foi encontrado para a loja " & loja & "!!"
						else
							blnAchou = False
							idxSelecionado = -1
							for j=LBound(vProduto) to UBound(vProduto)
								if (Trim("" & vProduto(j).fabricante) = Trim("" & rs("fabricante"))) And (Trim("" & vProduto(j).produto) = Trim("" & rs("produto"))) then
									blnAchou = True
									idxSelecionado = j
									exit for
									end if
								next

							if Not blnAchou then
								if Trim(vProduto(ubound(vProduto)).produto) <> "" then
									redim preserve vProduto(ubound(vProduto)+1)
									set vProduto(ubound(vProduto)) = New cl_ITEM_PEDIDO
									vProduto(ubound(vProduto)).qtde = 0
									end if
								idxSelecionado = ubound(vProduto)
								end if

							with vProduto(idxSelecionado)
								.fabricante = Trim("" & rs("fabricante"))
								.produto = Trim("" & rs("produto"))
								.qtde = .qtde + n_qtde
								.preco_lista = rs("preco_lista")
								.descricao = Trim("" & rs("descricao"))
								.descricao_html = Trim("" & rs("descricao_html"))
								end with
							rs.MoveNext
							if Not rs.Eof then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Há mais de um produto com o código '" & s_produto & "'!!<br />Informe o código do fabricante para resolver a ambiguidade!!"
								end if
							end if
						end if ' Produto composto ou normal?
			'	INFORMOU O CÓDIGO DO FABRICANTE E DO PRODUTO NA TELA ANTERIOR
				elseif (s_fabricante <> "") And (s_produto <> "") then
				'	VERIFICA SE É PRODUTO COMPOSTO
					strSql = "SELECT " & _
								"*" & _
							" FROM t_EC_PRODUTO_COMPOSTO t_EC_PC" & _
							" WHERE" & _
								" (fabricante_composto = '" & s_fabricante & "')" & _
								" AND (produto_composto = '" & s_produto & "')"
					if rs.State <> 0 then rs.Close
					rs.Open strSql, cn
				'	É PRODUTO COMPOSTO
					if Not rs.Eof then
						strSql = "SELECT " & _
									"*" & _
								" FROM t_EC_PRODUTO_COMPOSTO_ITEM t_EC_PCI" & _
								" WHERE" & _
									" (fabricante_composto = '" & s_fabricante & "')" & _
									" AND (produto_composto = '" & s_produto & "')" & _
									" AND (excluido_status = 0)" & _
								" ORDER BY" & _
									" sequencia"
						if rs.State <> 0 then rs.Close
						rs.Open strSql, cn
						do while Not rs.Eof
							strSql = "SELECT " & _
										"*" & _
									" FROM t_PRODUTO tP" & _
										" INNER JOIN t_PRODUTO_LOJA tPL ON (tP.fabricante = tPL.fabricante) AND (tP.produto = tPL.produto)" & _
									" WHERE" & _
										" (tP.fabricante = '" & Trim("" & rs("fabricante_item")) & "')" & _
										" AND (tP.produto = '" & Trim("" & rs("produto_item")) & "')" & _
										" AND (loja = '" & loja & "')"
							if rs2.State <> 0 then rs2.Close
							rs2.Open strSql, cn
							if rs2.Eof then
								alerta=texto_add_br(alerta)
								alerta=alerta & "O produto (" & Trim("" & rs("fabricante_item")) & ")" & Trim("" & rs("produto_item")) & " não está disponível para a loja " & loja & "!!"
							else
								blnAchou = False
								idxSelecionado = -1
								for j=LBound(vProduto) to UBound(vProduto)
									if (Trim("" & vProduto(j).fabricante) = Trim("" & rs("fabricante_item"))) And (Trim("" & vProduto(j).produto) = Trim("" & rs("produto_item"))) then
										blnAchou = True
										idxSelecionado = j
										exit for
										end if
									next

								if Not blnAchou then
									if Trim(vProduto(ubound(vProduto)).produto) <> "" then
										redim preserve vProduto(ubound(vProduto)+1)
										set vProduto(ubound(vProduto)) = New cl_ITEM_PEDIDO
										vProduto(ubound(vProduto)).qtde = 0
										end if
									idxSelecionado = ubound(vProduto)
									end if

								with vProduto(idxSelecionado)
									.fabricante = Trim("" & rs("fabricante_item"))
									.produto = Trim("" & rs("produto_item"))
									.qtde = .qtde + (n_qtde * rs("qtde"))
									.preco_lista = rs2("preco_lista")
									.descricao = Trim("" & rs2("descricao"))
									.descricao_html = Trim("" & rs2("descricao_html"))
									end with
								end if
							rs.MoveNext
							loop
				'	É PRODUTO NORMAL
					else
						strSql = "SELECT " & _
									"*" & _
								" FROM t_PRODUTO tP" & _
									" INNER JOIN t_PRODUTO_LOJA tPL ON (tP.fabricante = tPL.fabricante) AND (tP.produto = tPL.produto)" & _
								" WHERE" & _
									" (tP.fabricante = '" & s_fabricante & "')" & _
									" AND (tP.produto = '" & s_produto & "')" & _
									" AND (loja = '" & loja & "')"
						if rs.State <> 0 then rs.Close
						rs.Open strSql, cn
						if rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto (" & s_fabricante & ")" & s_produto & " não foi encontrado para a loja " & loja & "!!"
						else
							blnAchou = False
							idxSelecionado = -1
							for j=LBound(vProduto) to UBound(vProduto)
								if (Trim("" & vProduto(j).fabricante) = Trim("" & rs("fabricante"))) And (Trim("" & vProduto(j).produto) = Trim("" & rs("produto"))) then
									blnAchou = True
									idxSelecionado = j
									exit for
									end if
								next

							if Not blnAchou then
								if Trim(vProduto(ubound(vProduto)).produto) <> "" then
									redim preserve vProduto(ubound(vProduto)+1)
									set vProduto(ubound(vProduto)) = New cl_ITEM_PEDIDO
									vProduto(ubound(vProduto)).qtde = 0
									end if
								idxSelecionado = ubound(vProduto)
								end if

							with vProduto(idxSelecionado)
								.fabricante = Trim("" & rs("fabricante"))
								.produto = Trim("" & rs("produto"))
								.qtde = .qtde + n_qtde
								.preco_lista = rs("preco_lista")
								.descricao = Trim("" & rs("descricao"))
								.descricao_html = Trim("" & rs("descricao_html"))
								end with
							end if
						end if ' Produto composto ou normal?
					end if
				next
			
			if alerta = "" then
				n = 0
				for i=LBound(vProduto) to UBound(vProduto)
					if Trim(vProduto(i).produto) <> "" then n = n + 1
					next
				if n > MAX_ITENS then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O número de itens que está sendo cadastrado (" & CStr(n) & ") excede o máximo permitido por pedido (" & CStr(MAX_ITENS) & ")!!"
					end if
				end if
			end if 'if blnLojaHabilitadaProdCompostoECommerce
		end if 'if alerta = ""




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _______________________________________
' L I S T A _ L O J A _ I N D I C O U
'
function lista_loja_indicou
dim r,s,s_aux
	set r = cn.Execute("SELECT * FROM t_LOJA WHERE (comissao_indicacao > 0) ORDER BY CONVERT(smallint,loja)")
	s= ""
	do while Not r.eof 
		s_aux = normaliza_codigo(Trim("" & r("loja")), TAM_MIN_LOJA)
		s = s & "<OPTION" & _
				" VALUE='" & s_aux & "'>" & s_aux
		
		s_aux = Trim("" & r("nome"))
		if s_aux = "" then s_aux = Trim("" & r("razao_social"))
		
		if s_aux <> "" then s = s & " - "
		s = s & s_aux
		s = s & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	s = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & s
		
	lista_loja_indicou = s

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
					" (status = 'A')" & _
				" ORDER BY" & _
					" apelido"
	else
		'10/01/2020 - Unis - Desativação do acesso dos vendedores a todos os parceiros da Unis
		if (False And isLojaVrf(loja)) Or (loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then
		'	TODOS OS VENDEDORES COMPARTILHAM OS MESMOS INDICADORES
			strSql = "SELECT " & _
						"*" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (status = 'A')" & _
						" AND (loja = '" & loja & "')" & _
					" ORDER BY" & _
						" apelido"
		elseif (loja = NUMERO_LOJA_OLD03) Or (loja = NUMERO_LOJA_OLD03_BONIFICACAO) Or (operacao_permitida(OP_LJA_SELECIONAR_QUALQUER_INDICADOR_EM_PEDIDO_NOVO, s_lista_operacoes_permitidas)) then
		'	OLD03: LISTA COMPLETA DOS INDICADORES LIBERADA
			strSql = "SELECT " & _
						"*" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (status = 'A')" & _
					" ORDER BY" & _
						" apelido"
		else
			strSql = "SELECT " & _
						"*" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (status = 'A')" & _
						" AND (vendedor = '" & usuario & "')" & _
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
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("razao_social_nome"))
		strResp = strResp & "</OPTION>" & chr(13)
		
		strJsScript = strJsScript & _
						"vIndicador[vIndicador.length] = new oIndicador('" & QuotedStr(Trim("" & r("apelido"))) & "', " & Trim("" & r("permite_RA_status")) & ");" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
	
	strJsScript = strJsScript & "</script>" & chr(13)
	
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


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
		$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8
		if (loja == "<%=NUMERO_LOJA_ECOMMERCE_AR_CLUBE%>") {
			$(".trRT").hide();
		}
	});

	//Every resize of window
	$(window).resize(function() {
		sizeDivAjaxRunning();
	});

	//Every scroll of window
	$(window).scroll(function() {
		sizeDivAjaxRunning();
	});

	//Dynamically assign height
	function sizeDivAjaxRunning() {
		var newTop = $(window).scrollTop() + "px";
		$("#divAjaxRunning").css("top", newTop);
	}
</script>

<script language="JavaScript" type="text/javascript">
function oIndicador(apelido, permite_RA_status) {
	this.apelido = apelido;
	this.permite_RA_status = permite_RA_status;
}
</script>

<% =strJsScriptArrayIndicadores %>

<script language="JavaScript" type="text/javascript">
var fCustoFinancFornecParcelamentoPopup;
var objAjaxCustoFinancFornecConsultaPreco;
var loja="<%=loja%>";

function processaSelecaoCustoFinancFornecParcelamento(){};

function abreTabelaCustoFinancFornecParcelamento(intIndex){
var f, strUrl;
	f=fPED;
	if (trim(f.c_fabricante[intIndex].value)=="") {
		alert("Informe o código do fabricante do produto!");
		f.c_fabricante[intIndex].focus();
		return;
		}
	if (trim(f.c_produto[intIndex].value)=="") {
		alert("Informe o código do produto!");
		f.c_produto[intIndex].focus();
		return;
		}
		
	try
		{
	//  SE JÁ HOUVER UMA JANELA DE TABELA DE PARCELAMENTO ABERTA, GARANTE QUE ELA SERÁ FECHADA
	//  E UMA NOVA SERÁ CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')
		fCustoFinancFornecParcelamentoPopup.close();
		}
	catch (e) {
	 // NOP
		}
	processaSelecaoCustoFinancFornecParcelamento=trataSelecaoCustoFinancFornecParcelamento;
	strUrl="../Global/AjaxCustoFinancFornecParcelamentoPopup.asp";
	strUrl=strUrl+"?fabricante="+trim(f.c_fabricante[intIndex].value)+"&produto="+trim(f.c_produto[intIndex].value)+"&loja="+trim(f.c_loja.value)+"&tipoParcelamento="+f.c_custoFinancFornecTipoParcelamento.value+"&qtdeParcelas="+f.c_custoFinancFornecQtdeParcelas.value;
	try
	{
		fCustoFinancFornecParcelamentoPopup=window.open(strUrl, "AjaxCustoFinancFornecParcelamentoPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=800,height=675,left=0,top=0");
	}
	catch (e) {
		alert("Falha ao ativar o painel com a tabela de preços!!\n"+e.message);
		}
	
	try
	{
		fCustoFinancFornecParcelamentoPopup.focus();
	}
	catch (e) {
	 // NOP
		}
}

function trataSelecaoCustoFinancFornecParcelamento(strTipoParcelamento, strPrecoLista, intQtdeParcelas, strFabricante, strProduto) {
var f,i,blnAlterou;
	f=fPED;
//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
//	(apesar de que isso não será aceito pelas consistências que serão feitas).
	for (i=0; i<f.c_fabricante.length; i++) {
		if ((f.c_fabricante[i].value==strFabricante)&&(f.c_produto[i].value==strProduto)) {
			f.c_preco_lista[i].value=strPrecoLista;
			f.c_preco_lista[i].style.color="black";
			}
		}
	blnAlterou=false;
	if (f.c_custoFinancFornecTipoParcelamento.value!=strTipoParcelamento) blnAlterou=true;
	if (!blnAlterou) {
		if ((strTipoParcelamento==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA)||(strTipoParcelamento==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA)) {
			if (converte_numero(intQtdeParcelas)!=converte_numero(f.c_custoFinancFornecQtdeParcelas.value)) blnAlterou=true;
			}
		}
//  Memoriza seleção atual
	f.c_custoFinancFornecTipoParcelamento.value=strTipoParcelamento;
	f.c_custoFinancFornecQtdeParcelas.value=intQtdeParcelas;
	
	if (blnAlterou) {
		f.c_custoFinancFornecParcelamentoDescricao.value=descricaoCustoFinancFornecTipoParcelamento(strTipoParcelamento);
		if (strTipoParcelamento==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) {
			f.c_custoFinancFornecParcelamentoDescricao.value += " (1+" + intQtdeParcelas + ")";
			}
		else if (strTipoParcelamento==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) {
			f.c_custoFinancFornecParcelamentoDescricao.value += " (0+" + intQtdeParcelas + ")";
			}
	
		// Houve alteração no tipo de parcelamento, portanto, é necessário atualizar os 
		// preços de lista de todos os produtos
		atualizaPrecos(strFabricante, strProduto);
		}
	window.status="Concluído";
}

function trataRespostaAjaxCustoFinancFornecSincronizaPrecos() {
var f, strResp, i, j, xmlDoc, oNodes;
var strFabricante,strProduto, strStatus, strPrecoLista, strDescricao, strMsgErro;
	f=fPED;
	if (objAjaxCustoFinancFornecConsultaPreco.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxCustoFinancFornecConsultaPreco.responseText;
		if (strResp=="") {
			alert("Falha ao consultar o preço!!");
			window.status="Concluído";
			$("#divAjaxRunning").hide();
			return;
			}
		
		if (strResp!="") {
			try
				{
				xmlDoc=objAjaxCustoFinancFornecConsultaPreco.responseXML.documentElement;
				for (i=0; i < xmlDoc.getElementsByTagName("ItemConsulta").length; i++) {
				//  Fabricante
					oNodes=xmlDoc.getElementsByTagName("fabricante")[i];
					if (oNodes.childNodes.length > 0) strFabricante=oNodes.childNodes[0].nodeValue; else strFabricante="";
					if (strFabricante==null) strFabricante="";
				//  Produto
					oNodes=xmlDoc.getElementsByTagName("produto")[i];
					if (oNodes.childNodes.length > 0) strProduto=oNodes.childNodes[0].nodeValue; else strProduto="";
					if (strProduto==null) strProduto="";
				//  Status
					oNodes=xmlDoc.getElementsByTagName("status")[i];
					if (oNodes.childNodes.length > 0) strStatus=oNodes.childNodes[0].nodeValue; else strStatus="";
					if (strStatus==null) strStatus="";
					if (strStatus=="OK") {
					//  Descrição
						oNodes=xmlDoc.getElementsByTagName("descricao")[i];
						if (oNodes.childNodes.length > 0) strDescricao=oNodes.childNodes[0].nodeValue; else strDescricao="";
						if (strDescricao==null) strDescricao="";
					//  Preço
						oNodes=xmlDoc.getElementsByTagName("precoLista")[i];
						if (oNodes.childNodes.length > 0) strPrecoLista=oNodes.childNodes[0].nodeValue; else strPrecoLista="";
						if (strPrecoLista==null) strPrecoLista="";
					//  Atualiza o preço
						if (strPrecoLista=="") {
							alert("Falha na consulta do preço do produto " + strProduto + "!!\n" + strMsgErro);
							}
						else {
							for (j=0; j<f.c_fabricante.length; j++) {
								if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
								//	(apesar de que isso não será aceito pelas consistências que serão feitas).
									f.c_preco_lista[j].value=strPrecoLista;
									f.c_descricao[j].value=strDescricao;
									f.c_preco_lista[j].style.color="black";
									}
								}
							}
						}
					else {
					//  Mensagem de Erro
						oNodes=xmlDoc.getElementsByTagName("msg_erro")[i];
						if (oNodes.childNodes.length > 0) strMsgErro=oNodes.childNodes[0].nodeValue; else strMsgErro="";
						if (strMsgErro==null) strMsgErro="";
						for (j=0; j<f.c_fabricante.length; j++) {
						//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
						//	(apesar de que isso não será aceito pelas consistências que serão feitas).
							if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								f.c_preco_lista[j].style.color=COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE;
								}
							}
						alert("Falha ao consultar o preço do produto " + strProduto + "!!\n" + strMsgErro);
						}
					}
				}
			catch (e)
				{
				alert("Falha na consulta do preço!!\n"+e.message);
				}
			}
		window.status="Concluído";
		$("#divAjaxRunning").hide();
		}
}

function atualizaPrecos(strFabricanteSelecionado, strProdutoSelecionado) {
var f, i, strListaProdutos, strUrl;
	f=fPED;
	objAjaxCustoFinancFornecConsultaPreco=GetXmlHttpObject();
	if (objAjaxCustoFinancFornecConsultaPreco==null) {
		alert("O browser NÃO possui suporte ao AJAX!!");
		return;
		}
		
	strListaProdutos="";
	for (i=0; i<f.c_fabricante.length; i++) {
		if ((trim(f.c_fabricante[i].value)!="")&&(trim(f.c_produto[i].value)!="")) {
		//  Não atualiza o preço do produto que acabou de ser consultado através da tabela de preços.
		//  Atualiza somente os demais produtos, se houver.
			if ((strFabricanteSelecionado!=trim(f.c_fabricante[i].value))||(strProdutoSelecionado!=trim(f.c_produto[i].value))) {
				if (strListaProdutos!="") strListaProdutos+=";";
				strListaProdutos += f.c_fabricante[i].value + "|" + f.c_produto[i].value;
				}
			}
		}
	if (strListaProdutos=="") return;
	
	window.status="Aguarde, consultando preços ...";
	$("#divAjaxRunning").show();
		
	strUrl = "../Global/AjaxCustoFinancFornecConsultaPrecoBD.asp";
	strUrl+="?tipoParcelamento="+f.c_custoFinancFornecTipoParcelamento.value;
	strUrl+="&qtdeParcelas="+f.c_custoFinancFornecQtdeParcelas.value;
	strUrl+="&loja="+f.c_loja.value;
	strUrl+="&listaProdutos="+strListaProdutos;
//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxCustoFinancFornecConsultaPreco.onreadystatechange=trataRespostaAjaxCustoFinancFornecSincronizaPrecos;
	objAjaxCustoFinancFornecConsultaPreco.open("GET",strUrl,true);
	objAjaxCustoFinancFornecConsultaPreco.send(null);
}

function trataRespostaAjaxCustoFinancFornecConsultaPreco() {
var f, strResp, i, j, xmlDoc, oNodes;
var strFabricante,strProduto, strStatus, strPrecoLista, strDescricao, strMsgErro;
	f=fPED;
	if (objAjaxCustoFinancFornecConsultaPreco.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxCustoFinancFornecConsultaPreco.responseText;
		if (strResp=="") {
			alert("Falha ao consultar o preço!!");
			window.status="Concluído";
			$("#divAjaxRunning").hide();
			return;
			}
		
		if (strResp!="") {
			try
				{
				xmlDoc=objAjaxCustoFinancFornecConsultaPreco.responseXML.documentElement;
				for (i=0; i < xmlDoc.getElementsByTagName("ItemConsulta").length; i++) {
				//  Fabricante
					oNodes=xmlDoc.getElementsByTagName("fabricante")[i];
					if (oNodes.childNodes.length > 0) strFabricante=oNodes.childNodes[0].nodeValue; else strFabricante="";
					if (strFabricante==null) strFabricante="";
				//  Produto
					oNodes=xmlDoc.getElementsByTagName("produto")[i];
					if (oNodes.childNodes.length > 0) strProduto=oNodes.childNodes[0].nodeValue; else strProduto="";
					if (strProduto==null) strProduto="";
				//  Descrição
					oNodes=xmlDoc.getElementsByTagName("descricao")[i];
					if (oNodes.childNodes.length > 0) strDescricao=oNodes.childNodes[0].nodeValue; else strDescricao="";
					if (strDescricao==null) strDescricao="";
					if (strDescricao!="") {
						for (j=0; j<f.c_fabricante.length; j++) {
							if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
							//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
							//	(apesar de que isso não será aceito pelas consistências que serão feitas).
								f.c_descricao[j].value=strDescricao;
								}
							}
						}
				//  Status
					oNodes=xmlDoc.getElementsByTagName("status")[i];
					if (oNodes.childNodes.length > 0) strStatus=oNodes.childNodes[0].nodeValue; else strStatus="";
					if (strStatus==null) strStatus="";
					if (strStatus=="OK") {
					//  Preço
						oNodes=xmlDoc.getElementsByTagName("precoLista")[i];
						if (oNodes.childNodes.length > 0) strPrecoLista=oNodes.childNodes[0].nodeValue; else strPrecoLista="";
						if (strPrecoLista==null) strPrecoLista="";
					//  Atualiza o preço
						if (strPrecoLista=="") {
							alert("Falha na consulta do preço do produto " + strProduto + "\n" + strMsgErro);
							}
						else {
							for (j=0; j<f.c_fabricante.length; j++) {
								if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
								//	(apesar de que isso não será aceito pelas consistências que serão feitas).
									f.c_preco_lista[j].value=strPrecoLista;
									f.c_preco_lista[j].style.color="black";
									}
								}
							}
						}
					else {
					//  Mensagem de Erro
						oNodes=xmlDoc.getElementsByTagName("msg_erro")[i];
						if (oNodes.childNodes.length > 0) strMsgErro=oNodes.childNodes[0].nodeValue; else strMsgErro="";
						if (strMsgErro==null) strMsgErro="";
						for (j=0; j<f.c_fabricante.length; j++) {
						//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
						//	(apesar de que isso não será aceito pelas consistências que serão feitas).
							if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								f.c_preco_lista[j].style.color=COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE;
								}
							}
						alert("Falha ao consultar o preço do produto " + strProduto + "\n" + strMsgErro);
						}
					}
				}
			catch (e)
				{
				alert("Falha na consulta do preço!!\n"+e.message);
				}
			}
		window.status="Concluído";
		$("#divAjaxRunning").hide();
		}
}

function consultaPreco(intIndice) {
var f, i, strProdutoSelecionado, strUrl;
	f=fPED;
	if (trim(f.c_fabricante[intIndice].value)=="") return;
	if (trim(f.c_produto[intIndice].value)=="") return;
	
	objAjaxCustoFinancFornecConsultaPreco=GetXmlHttpObject();
	if (objAjaxCustoFinancFornecConsultaPreco==null) {
		alert("O browser NÃO possui suporte ao AJAX!!");
		return;
		}
		
	strProdutoSelecionado=f.c_fabricante[intIndice].value + "|" + f.c_produto[intIndice].value;
	
	window.status="Aguarde, consultando preço ...";
	$("#divAjaxRunning").show();
	
	strUrl = "../Global/AjaxCustoFinancFornecConsultaPrecoBD.asp";
	strUrl+="?tipoParcelamento="+f.c_custoFinancFornecTipoParcelamento.value;
	strUrl+="&qtdeParcelas="+f.c_custoFinancFornecQtdeParcelas.value;
	strUrl+="&loja="+f.c_loja.value;
	strUrl+="&listaProdutos="+strProdutoSelecionado;
//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxCustoFinancFornecConsultaPreco.onreadystatechange=trataRespostaAjaxCustoFinancFornecConsultaPreco;
	objAjaxCustoFinancFornecConsultaPreco.open("GET",strUrl,true);
	objAjaxCustoFinancFornecConsultaPreco.send(null);
}

function trataLimpaLinha(intIndice) {
var f;
	f=fPED;
	if ((trim(f.c_fabricante[intIndice].value)=="")&&(trim(f.c_produto[intIndice].value)=="")) {
		f.c_qtde[intIndice].value="";
		f.c_descricao[intIndice].value="";
		f.c_preco_lista[intIndice].value="";
		}
}

function trata_indicador_onchange() {
var f, i;
	f = fPED;

	if (trim(f.c_indicador.value) == '') {
		if (loja == "<%=NUMERO_LOJA_ECOMMERCE_AR_CLUBE%>") {
			f.rb_RA[0].checked = false; // SEM RA
			f.rb_RA[0].disabled = true;
			f.rb_RA[1].checked = true;
			f.rb_RA[1].disabled = true;
		}
		else {
			f.rb_RA[0].checked = true; // SEM RA
			f.rb_RA[0].disabled = true;
			f.rb_RA[1].checked = false;
			f.rb_RA[1].disabled = true;
		}
		return;
	}

	for (i = 0; i < vIndicador.length; i++) {
		if (vIndicador[i].apelido == trim(f.c_indicador.value)) {
			if (vIndicador[i].permite_RA_status.toString() == '1') {
				if (loja == "<%=NUMERO_LOJA_ECOMMERCE_AR_CLUBE%>") {
					f.rb_RA[0].checked = false;
					f.rb_RA[0].disabled = false;
					f.rb_RA[1].checked = true;
					f.rb_RA[1].disabled = false;
				}
				else {
					f.rb_RA[0].checked = false;
					f.rb_RA[0].disabled = false;
					f.rb_RA[1].checked = false;
					f.rb_RA[1].disabled = false;
				}
				return;
			}
			break;
		}
	}

	if (loja == "<%=NUMERO_LOJA_ECOMMERCE_AR_CLUBE%>")
	{
		f.rb_RA[0].checked = false; // SEM RA
		f.rb_RA[0].disabled = true;
		f.rb_RA[1].checked = true;
		f.rb_RA[1].disabled = true;
	}
	else {
		f.rb_RA[0].checked = true; // SEM RA
		f.rb_RA[0].disabled = true;
		f.rb_RA[1].checked = false;
		f.rb_RA[1].disabled = true;
	}
}

function fPEDConfirma( f ) {
var s, i, b, ha_item, idx, blnIndicacaoOk, strMsgErro;
	ha_item=false;
	for (i=0; i < f.c_produto.length; i++) {
		b=false;
		if (trim(f.c_fabricante[i].value)!="") b=true;
		if (trim(f.c_produto[i].value)!="") b=true;
		if (trim(f.c_qtde[i].value)!="") b=true;
		
		if (b) {
			ha_item=true;
			if (trim(f.c_fabricante[i].value)=="") {
				alert("Informe o código do fabricante!!");
				f.c_fabricante[i].focus();
				return;
				}
			if (trim(f.c_produto[i].value)=="") {
				alert("Informe o código do produto!!");
				f.c_produto[i].focus();
				return;
				}
			if (trim(f.c_qtde[i].value)=="") {
				alert("Informe a quantidade!!");
				f.c_qtde[i].focus();
				return;
				}
			if (parseInt(f.c_qtde[i].value)<=0) {
				alert("Quantidade inválida!!");
				f.c_qtde[i].focus();
				return;
				}
			}
		}
		
	if (!ha_item) {
		alert("Não há produtos na lista!!");
		f.c_fabricante[0].focus();
		return;
		}
		
	if (f.vendedor_externo.value=="S") {
		if (trim(f.loja_indicou.options[f.loja_indicou.selectedIndex].value)=="") {
			alert('Especifique a loja que fez a indicação!!');
			return;
			}
		}
	
	if (f.c_ExibirCamposComSemIndicacao.value!="S") {
		blnIndicacaoOk=true;
	}
	else if (f.c_MagentoPedidoComIndicador.value=="S") {
		blnIndicacaoOk=true;
	}
	else {
		idx=-1;
		blnIndicacaoOk=false;
		//  Sem Indicação
		idx++;
		if (f.rb_indicacao[idx].checked) {
			blnIndicacaoOk=true;
		}
		
		//	Com Indicação
		idx++;
		if (f.rb_indicacao[idx].checked) {
			if (trim(f.c_indicador.value)=='') {
				alert('Selecione o "indicador"!!');
				f.c_indicador.focus();
				return;
			}
			if ((!f.rb_RA[0].checked)&&(!f.rb_RA[1].checked)) {
				alert('Informe se o pedido possui RA ou não!!');
				return;
			}
			blnIndicacaoOk=true;
			//  O indicador informado agora é diferente do indicador original no cadastro do cliente?
			if (loja != "<%=NUMERO_LOJA_ECOMMERCE_AR_CLUBE%>") {
				if (trim(f.c_indicador_original.value)!="") {
					if (trim(f.c_indicador.value)!=trim(f.c_indicador_original.value)) {
						s="O indicador selecionado é diferente do indicador que consta no cadastro deste cliente.\n\n##################################################\nFAVOR COMUNICAR AO GERENTE!!\n##################################################\n\nContinua mesmo assim?";
						if (!confirm(s)) return;
					}
				}
			}
		}

		if (!blnIndicacaoOk) {
			alert('Informe se o pedido é com indicação ou não!!');
			return;
		}
	}
	
	if (trim(f.c_custoFinancFornecTipoParcelamento.value)=="") {
		alert('Não foi informada a forma de pagamento!');
		return;
		}
		
	if ((f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA)||
		(f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA)) {
		if (converte_numero(f.c_custoFinancFornecQtdeParcelas.value)==0) {
			alert('Não foi informada a quantidade de parcelas da forma de pagamento!');
			return;
			}
		}
	
	strMsgErro="";
	for (i=0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value)!="") {
			if (f.c_preco_lista[i].style.color.toLowerCase()==COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE.toLowerCase()) {
				strMsgErro+="\n" + f.c_produto[i].value + " - " + f.c_descricao[i].value;
				}
			}
		}
	if (strMsgErro!="") {
		strMsgErro="A forma de pagamento " + KEY_ASPAS + f.c_custoFinancFornecParcelamentoDescricao.value.toLowerCase() + KEY_ASPAS + " não está disponível para o(s) produto(s):"+strMsgErro;
		alert(strMsgErro);
		return;
		}
	
	if (f.c_ExibirCamposModoSelecaoCD.value=="S")
	{
		if ((!f.rb_selecao_cd[0].checked)&&(!f.rb_selecao_cd[1].checked))
		{
			strMsgErro="É necessário informar o modo de seleção do CD (auto-split)!";
			alert(strMsgErro);
			return;
		}

		if (f.rb_selecao_cd[1].checked)
		{
			if (trim(f.c_id_nfe_emitente_selecao_manual.value)=="")
			{
				strMsgErro="É necessário selecionar o CD que irá atender o pedido (sem auto-split)!";
				alert(strMsgErro);
				f.c_id_nfe_emitente_selecao_manual.focus();
				return;
			}
		}
	}

	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style type="text/css">
#rb_indicacao {
	margin: 0pt 2pt 1pt 10pt;
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
<% if s_tabela_municipios_IBGE <> "" then %>
	<br /><br />
	<%=s_tabela_municipios_IBGE%>
<% end if %>
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
<body onload="trata_indicador_onchange(); if (trim(fPED.c_fabricante[0].value)=='') fPED.c_fabricante[0].focus();">
<center>

<form id="fPED" name="fPED" method="post" action="PedidoNovoConsiste.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_loja" id="c_loja" value='<%=loja%>'>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=cliente_selecionado%>'>
<input type="hidden" name="vendedor_externo" id="vendedor_externo" value='<%if Session("vendedor_externo") then Response.Write "S"%>'>
<input type="hidden" name="rb_end_entrega" id="rb_end_entrega" value='<%=rb_end_entrega%>'>
<input type="hidden" name="EndEtg_endereco" id="EndEtg_endereco" value="<%=EndEtg_endereco%>">
<input type="hidden" name="EndEtg_endereco_numero" id="EndEtg_endereco_numero" value="<%=EndEtg_endereco_numero%>">
<input type="hidden" name="EndEtg_endereco_complemento" id="EndEtg_endereco_complemento" value="<%=EndEtg_endereco_complemento%>">
<input type="hidden" name="EndEtg_bairro" id="EndEtg_bairro" value="<%=EndEtg_bairro%>">
<input type="hidden" name="EndEtg_cidade" id="EndEtg_cidade" value="<%=EndEtg_cidade%>">
<input type="hidden" name="EndEtg_uf" id="EndEtg_uf" value="<%=EndEtg_uf%>">
<input type="hidden" name="EndEtg_cep" id="EndEtg_cep" value="<%=EndEtg_cep%>">
<input type="hidden" name="c_indicador_original" id="c_indicador_original" value='<%=r_cliente.indicador%>'>
<input type="hidden" name="EndEtg_obs" id="EndEtg_obs" value='<%=EndEtg_obs%>'>
<%	if operacao_permitida(OP_LJA_EXIBIR_CAMPOS_COM_SEM_INDICACAO_AO_CADASTRAR_NOVO_PEDIDO, s_lista_operacoes_permitidas) then
		strAux="S"
	else
		strAux="N"
	end if
%>
<input type="hidden" name="c_ExibirCamposComSemIndicacao" id="c_ExibirCamposComSemIndicacao" value='<%=strAux%>'>
<% if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) And blnMagentoPedidoComIndicador then
		strAux="S"
	else
		strAux="N"
	end if
%>
<input type="hidden" name="c_MagentoPedidoComIndicador" id="c_MagentoPedidoComIndicador" value='<%=strAux%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamento" id="c_custoFinancFornecTipoParcelamento" value='<%=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelas" id="c_custoFinancFornecQtdeParcelas" value='0'>
<%	if operacao_permitida(OP_LJA_CADASTRA_NOVO_PEDIDO_SELECAO_MANUAL_CD, s_lista_operacoes_permitidas) then
		strAux="S"
	else
		strAux="N"
	end if
%>
<input type="hidden" name="c_ExibirCamposModoSelecaoCD" id="c_ExibirCamposModoSelecaoCD" value='<%=strAux%>'>
<input type="hidden" name="operacao_origem" id="operacao_origem" value="<%=operacao_origem%>" />
<input type="hidden" name="id_magento_api_pedido_xml" id="id_magento_api_pedido_xml" value="<%=id_magento_api_pedido_xml%>" />
<input type="hidden" name="c_numero_magento" id="c_numero_magento" value="<%=c_numero_magento%>" />
<input type="hidden" name="operationControlTicket" id="operationControlTicket" value="<%=operationControlTicket%>" />
<input type="hidden" name="sessionToken" id="sessionToken" value="<%=sessionToken%>" />

<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>


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
			<% if c_mag_cpf_cnpj_identificado <> "" then %>
			<br /><span class="C"><%=cnpj_cpf_formata(c_mag_cpf_cnpj_identificado)%></span>
			<% end if %>
		</td>
	</tr>
	<% if blnMagentoPedidoComIndicador then %>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">Indicador</span></td>
		<td class="MB MD TdCliCel">
			<span class="C"><%=cnpj_cpf_formata(c_mag_installer_document)%></span>
			<br /><span class="C"><%=sIdInstalador & " - " & sNomeInstalador%></span>
		</td>
	</tr>
	<% end if %>
</table>
<% end if %>

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<br />
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pedido Novo</span></td>
</tr>
</table>
<br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0" <%if blnLojaHabilitadaProdCompostoECommerce then Response.Write "style='display:none;'"%> >
	<tr bgColor="#FFFFFF">
	<td class="MB" align="left"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left"><span class="PLTe">Produto</span></td>
	<td class="MB" align="right"><span class="PLTd">Qtde</span></td>
	<td class="MB" align="left"><span class="PLTe">Descrição</span></td>
	<td class="MB" align="right"><span class="PLTd">VL Unit</span></td>
	<td align="left">&nbsp;</td>
	</tr>
<% intIdxProduto = LBound(vProduto)-1 %>
<% for i=1 to MAX_ITENS 
		intIdxProduto = intIdxProduto + 1
		s_fabricante = ""
		s_produto = ""
		s_qtde = ""
		s_preco_lista = ""
		s_descricao = ""
		if blnLojaHabilitadaProdCompostoECommerce then
			if intIdxProduto <= Ubound(vProduto) then
				if Trim("" & vProduto(intIdxProduto).produto) <> "" then
					with vProduto(intIdxProduto)
						s_fabricante = .fabricante
						s_produto = .produto
						s_qtde = CStr(.qtde)
						s_preco_lista = formata_moeda(.preco_lista)
						s_descricao = .descricao
						end with
					end if
				end if
			end if
%>
	<tr>
	<td class="MDBE" align="left">
		<input name="c_fabricante" id="c_fabricante" class="PLLe" maxlength="4" style="width:30px;" onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(<%=Cstr(i)%>!=1))) if (trim(this.value)=='') bCONFIRMA.focus(); else fPED.c_produto[<%=Cstr(i-1)%>].focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);trataLimpaLinha(<%=Cstr(i-1)%>);"
			<% 'A DECLARAÇÃO DA PROPRIEDADE VALUE APENAS SE HOUVER VALOR EVITA QUE O CAMPO SEJA LIMPO APÓS A SEGUNDA CHAMADA DO HISTORY.BACK() NA TELA SEGUINTE QUANDO OCORRE ERRO DE CONSISTÊNCIA E DESEJA-SE RETORNAR À ESTA TELA %>
			<% if s_fabricante <> "" then %>
			value="<%=s_fabricante%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="left">
		<input name="c_produto" id="c_produto" class="PLLe" maxlength="8" style="width:60px;" onkeypress="if (digitou_enter(true)) fPED.c_qtde[<%=Cstr(i-1)%>].focus(); filtra_produto();" onblur="this.value=normaliza_produto(this.value);consultaPreco(<%=Cstr(i-1)%>);trataLimpaLinha(<%=Cstr(i-1)%>);"
			<% if s_produto <> "" then %>
			value="<%=s_produto%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde" class="PLLd" maxlength="4" style="width:30px;" onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==fPED.c_qtde.length) bCONFIRMA.focus(); else fPED.c_fabricante[<%=Cstr(i)%>].focus();} filtra_numerico();"
			<% if s_qtde <> "" then %>
			value="<%=s_qtde%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="left">
		<input name="c_descricao" id="c_descricao" class="PLLe" style="width:377px;" readonly tabindex=-1
			<% if s_descricao <> "" then %>
			value="<%=s_descricao%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="right">
		<input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px;" readonly tabindex=-1 
			<% if s_preco_lista <> "" then %>
			value="<%=s_preco_lista%>"
			<% end if %>
			/>
	</td>
	<td align="left">&nbsp;<button type="button" name="bCustoFinancFornecParcelamento" id="bCustoFinancFornecParcelamento" style='width:50px;font-size:8pt;font-weight:bold;color:black;margin-bottom:1px;' class="Botao" onclick="abreTabelaCustoFinancFornecParcelamento(<%=i-1%>);"><%=SIMBOLO_MONETARIO%></button></td>
	</tr>
<% next %>
</table>



<div  <%if blnLojaHabilitadaProdCompostoECommerce then Response.Write "style='display:none;'"%>>
    <br />
    <span class="PLLe">Forma de pagamento: </span>
    <input name="c_custoFinancFornecParcelamentoDescricao" id="c_custoFinancFornecParcelamentoDescricao" class="PLLe" style="width:115px;color:#0000CD;font-weight:bold;"
	    value="<%=descricaoCustoFinancFornecTipoParcelamento(COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA)%>">
    <br />
</div>

<%	IF Session("vendedor_externo") THEN %>
	<br>
	<!-- ************   LOJA QUE INDICOU   ************ -->
	<table cellspacing="0">
		<tr>
		<td width="100%" align="left"><p class="R">LOJA QUE FEZ A INDICAÇÃO</p><p class="C">
			<select id="loja_indicou" name="loja_indicou" style="margin-top:4pt; margin-bottom:4pt;">
				<%=lista_loja_indicou%>
			</select>
		</tr>
	</table>
    <br /><br />
<%	END IF %>


<!-- ************   INDICADOR   ************ -->
<% if (operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO) And blnMagentoPedidoComIndicador then %>
	<input type="hidden" id="rb_indicacao" name="rb_indicacao" value="S" />
	<input type="hidden" id="rb_RA" name="rb_RA" value="S" />
	<input type="hidden" id="c_perc_RT" name="c_perc_RT" value="<%=formata_perc(CDbl(percCommissionValue) - CDbl(percCommissionDiscount))%>" />
	<input type="hidden" id="c_indicador" name="c_indicador" value="<%=sIdInstalador%>" />

	<table style="width:300px;" cellpadding="2" cellspacing="0" border="0">
	<tr>
		<td align="right"><span class="C">Comissão (cadastrada):</span></td>
		<td align="left"><span class="C"><%=formata_perc(CDbl(percCommissionValue))%>%</span></td>
	</tr>
	<tr>
		<td align="right"><span class="C">Comissão (desconto concedido):</span></td>
		<td align="left"><span class="C"><%=formata_perc(CDbl(percCommissionDiscount))%>%</span></td>
	</tr>
	<tr>
		<td align="right" class="MC"><span class="C">COM(%)</span></td>
		<td align="left" class="MC"><span class="C"><%=formata_perc(CDbl(percCommissionValue) - CDbl(percCommissionDiscount))%>%</span></td>
	</tr>
	</table>
<% elseif operacao_permitida(OP_LJA_EXIBIR_CAMPOS_COM_SEM_INDICACAO_AO_CADASTRAR_NOVO_PEDIDO, s_lista_operacoes_permitidas) then %>
<table class="Q" style="width:375px;" cellspacing="0">
  <tr>
	<td align="left">
	  <p class="Rf">Indicação</p>
	</td>
  </tr>  
  <tr>
	<td align="left">
	  <table width="100%" cellspacing="0" cellpadding="4" border="0">
		<!--  SEM INDICAÇÃO  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left" valign="baseline">
				  <% intIdx = 0 %>
				  <input type="radio" id="rb_indicacao" name="rb_indicacao" value="N" <%if intQtdeIndicadores=0 then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_indicacao[<%=Cstr(intIdx)%>].click();">Sem Indicação</span>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  COM INDICAÇÃO  -->
		<tr>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td colspan="2" align="left" valign="baseline">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_indicacao" name="rb_indicacao" value="S" <%if intQtdeIndicadores=0 then Response.Write " disabled"%>><span class="C" style="cursor:default" onclick="fPED.rb_indicacao[<%=Cstr(intIdx)%>].click();">Com Indicação</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:40px;" align="left">&nbsp;</td>
				<td align="left">
				  <select id="c_indicador" name="c_indicador" onclick="fPED.rb_indicacao[<%=Cstr(intIdx)%>].click();" onchange="trata_indicador_onchange();">
					<% =strCampoSelectIndicadores %>
				  </select>
				</td>
			  </tr>
			  <tr>
				<td align="left">&nbsp;</td>
				<td align="left">
					<% intIdxRA = 0 %>
					<input type="radio" id="rb_RA" name="rb_RA" value="N" onclick="fPED.rb_indicacao[<%=Cstr(intIdx)%>].click();"><span class="C" style="cursor:default" onclick="fPED.rb_RA[<%=Cstr(intIdxRA)%>].click();">Sem RA</span>
				</td>
			  </tr>
			  <tr>
				<td align="left">&nbsp;</td>
				<td align="left">
					<% intIdxRA = intIdxRA+1 %>
					<input type="radio" id="rb_RA" name="rb_RA" value="S" onclick="fPED.rb_indicacao[<%=Cstr(intIdx)%>].click();"><span class="C" style="cursor:default" onclick="fPED.rb_RA[<%=Cstr(intIdxRA)%>].click();">Com RA</span>
				</td>
			  </tr>
			  <% if operacao_permitida(OP_LJA_EXIBIR_CAMPO_RT_AO_CADASTRAR_NOVO_PEDIDO, s_lista_operacoes_permitidas) then %>
			  <tr class="trRT">
				<td align="left">&nbsp;</td>
				<td align="left" style="padding-bottom: 5px">
					<span class="C">COM(%)</span>
					<input id="c_perc_RT" name="c_perc_RT" value="" maxlength="5" size="5" 
							onclick="fPED.rb_indicacao[<%=Cstr(intIdx)%>].click();"
							onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_percentual();"
							onblur="this.value=formata_perc_RT(this.value); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();}">
				</td>
			  </tr>
			  <% end if %>
			  <!-------------- PEDIDO BONSHOP ---------------->
			  <% if ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then %>
			  <tr>
			    <td align="left" class="MC">&nbsp;</td>
			    <td align="left" class="MC" style="padding: 5px 0px">
			        <span class="C">Ref. Pedido Bonshop:</span>
			        <select id="pedBonshop" name="pedBonshop" style="width: 80px">
			            <option value="">&nbsp;</option>
			            <%
			            dim cn2
    If Not bdd_BS_conecta(cn2) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
    dim r, sqlString, strResp
     
     sqlString = "SELECT pedido FROM t_PEDIDO" & _
              " INNER JOIN t_CLIENTE ON (t_CLIENTE.id = t_PEDIDO.id_cliente)" & _
              " WHERE t_CLIENTE.cnpj_cpf = '" & r_cliente.cnpj_cpf & "'" & _
              " AND (st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
              " ORDER BY data DESC, pedido"
     set r = cn2.Execute(sqlString)
     strResp = ""
     do while Not r.eof
        strResp = strResp & "<option value='" & r("pedido") & "'>" 
        strResp = strResp & r("pedido")
        strResp = strResp & "</option>" & chr(13)
        r.MoveNext
     loop
     Response.Write strResp
     r.close
     set r=nothing
     cn2.Close
     set cn2 = nothing%>
			        </select>
			    </td>
			  </tr>
			  <% end if %>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
</table>
<%else%>
<input type="hidden" id="rb_indicacao" name="rb_indicacao" value="N">
<%end if%>


<!-- ************   SELEÇÃO MANUAL DO CD   ************ -->
<% if operacao_permitida(OP_LJA_CADASTRA_NOVO_PEDIDO_SELECAO_MANUAL_CD, s_lista_operacoes_permitidas) then %>
<br />
<table class="Q" style="width:375px;" cellspacing="0">
  <tr>
	<td align="left">
	  <p class="Rf">Modo de Seleção do CD (Auto-Split)</p>
	</td>
  </tr>
  <tr>
	<td align="left">
	  <table width="100%" cellspacing="0" cellpadding="4" border="0">
		<!--  AUTOMÁTICO  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left" valign="baseline">
				  <% intIdx = 0 %>
				  <input type="radio" name="rb_selecao_cd" id="rb_selecao_cd_auto" value="<%=MODO_SELECAO_CD__AUTOMATICO%>"><span class="C" style="cursor:default" onclick="fPED.rb_selecao_cd[<%=Cstr(intIdx)%>].click();">Automático</span>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<tr>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td colspan="2" align="left" valign="baseline">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" name="rb_selecao_cd" id="rb_selecao_cd_manual" value="<%=MODO_SELECAO_CD__MANUAL%>"><span class="C" style="cursor:default" onclick="fPED.rb_selecao_cd[<%=Cstr(intIdx)%>].click();">Manual</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:40px;" align="left">&nbsp;</td>
				<td align="left">
				  <select id="c_id_nfe_emitente_selecao_manual" name="c_id_nfe_emitente_selecao_manual" onclick="fPED.rb_selecao_cd[<%=Cstr(intIdx)%>].click();">
					<% =wms_apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
				  </select>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
</table>
<% else %>
<input type="hidden" name="rb_selecao_cd" value="<%=MODO_SELECAO_CD__AUTOMATICO%>" />
<% end if %>


<!-- ************   SEPARADOR   ************ -->
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="749" cellspacing="0">
<tr>
	<% if blnLojaHabilitadaProdCompostoECommerce then %>
	<td align="left"><a name="bCANCELA" id="bCANCELA" href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<% else %>
	<td align="left"><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela o novo pedido">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<% end if %>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDConfirma(fPED)" title="vai para a página de confirmação">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
<% end if %>

</html>

<%
	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		if tMAP_XML.State <> 0 then tMAP_XML.Close
		set tMAP_XML = nothing
		end if

	if rs.State <> 0 then rs.Close
	set rs = nothing

	if rs2.State <> 0 then rs2.Close
	set rs2 = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>