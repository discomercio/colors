<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->
<%
'     =================================
'	  O R C A M E N T O N O V O . A S P
'     =================================
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

	dim i, j, usuario, loja, cliente_selecionado, msg_erro
	dim idxSelecionado, blnAchou

	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("Aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("Aviso.asp?id=" & ERR_SESSAO) 
	
	
	cliente_selecionado = Trim(request("cliente_selecionado"))
	if (cliente_selecionado = "") then Response.Redirect("Aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)

	dim cn, r, strSql
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim max_qtde_itens
	max_qtde_itens = obtem_parametro_PedidoItem_MaxQtdeItens

	dim blnLojaHabilitadaProdCompostoECommerce
	blnLojaHabilitadaProdCompostoECommerce = isLojaHabilitadaProdCompostoECommerce(loja)

	dim r_cliente
	set r_cliente = New cl_CLIENTE
	if Not x_cliente_bd(cliente_selecionado, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_FALHA_RECUPERAR_DADOS)

	dim rb_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs
	dim EndEtg_email, EndEtg_email_xml, EndEtg_nome, EndEtg_ddd_res, EndEtg_tel_res, EndEtg_ddd_com, EndEtg_tel_com, EndEtg_ramal_com
	dim EndEtg_ddd_cel, EndEtg_tel_cel, EndEtg_ddd_com_2, EndEtg_tel_com_2, EndEtg_ramal_com_2
	dim EndEtg_tipo_pessoa, EndEtg_cnpj_cpf, EndEtg_contribuinte_icms_status, EndEtg_produtor_rural_status
	dim EndEtg_ie, EndEtg_rg
	rb_end_entrega = Trim(Request.Form("rb_end_entrega"))
	EndEtg_endereco = Trim(Request.Form("EndEtg_endereco"))
	EndEtg_endereco_numero = Trim(Request.Form("EndEtg_endereco_numero"))
	EndEtg_endereco_complemento = Trim(Request.Form("EndEtg_endereco_complemento"))
	EndEtg_bairro = Trim(Request.Form("EndEtg_bairro"))
	EndEtg_cidade = Trim(Request.Form("EndEtg_cidade"))
	EndEtg_uf = Trim(Request.Form("EndEtg_uf"))
	EndEtg_cep = Trim(Request.Form("EndEtg_cep"))
    EndEtg_obs = Trim(Request.Form("EndEtg_obs"))
	EndEtg_email = Trim(Request.Form("EndEtg_email"))
	EndEtg_email_xml = Trim(Request.Form("EndEtg_email_xml"))
	EndEtg_nome = Trim(Request.Form("EndEtg_nome"))
	EndEtg_tipo_pessoa = Trim(Request.Form("EndEtg_tipo_pessoa"))
	EndEtg_cnpj_cpf = Trim(Request.Form("EndEtg_cnpj_cpf"))
	EndEtg_contribuinte_icms_status = Trim(Request.Form("EndEtg_contribuinte_icms_status"))
	EndEtg_produtor_rural_status = Trim(Request.Form("EndEtg_produtor_rural_status"))
	EndEtg_ie = Trim(Request.Form("EndEtg_ie"))
	EndEtg_rg = Trim(Request.Form("EndEtg_rg"))

	'Tratamento para obter os dados de telefone conforme os campos exibidos no formulário
	'O objetivo é evitar que dados sejam gravados de forma inconsistente na seguinte situação:
	'	O usuário seleciona o tipo PJ e inicia o preenchimento dos dados do telefone, mas não informa o DDD
	'	Em seguida, altera o tipo para PF e realiza o preenchimento corretamente
	'	Os dados de telefone exibidos p/ o tipo PJ estão inconsistentes e não devem ser gravados no BD, até porque, mesmo que corretos, seriam informações que não pertencem ao contexto selecionado
	EndEtg_ddd_res = ""
	EndEtg_tel_res = ""
	EndEtg_ddd_com = ""
	EndEtg_tel_com = ""
	EndEtg_ramal_com = ""
	EndEtg_ddd_cel = ""
	EndEtg_tel_cel = ""
	EndEtg_ddd_com_2 = ""
	EndEtg_tel_com_2 = ""
	EndEtg_ramal_com_2 = ""

	if EndEtg_tipo_pessoa = ID_PF then
		EndEtg_ddd_res = Trim(Request.Form("EndEtg_ddd_res"))
		EndEtg_tel_res = Trim(Request.Form("EndEtg_tel_res"))
		EndEtg_ddd_cel = Trim(Request.Form("EndEtg_ddd_cel"))
		EndEtg_tel_cel = Trim(Request.Form("EndEtg_tel_cel"))
	else
		EndEtg_ddd_com = Trim(Request.Form("EndEtg_ddd_com"))
		EndEtg_tel_com = Trim(Request.Form("EndEtg_tel_com"))
		EndEtg_ramal_com = Trim(Request.Form("EndEtg_ramal_com"))
		EndEtg_ddd_com_2 = Trim(Request.Form("EndEtg_ddd_com_2"))
		EndEtg_tel_com_2 = Trim(Request.Form("EndEtg_tel_com_2"))
		EndEtg_ramal_com_2 = Trim(Request.Form("EndEtg_ramal_com_2"))
		end if

    dim orcamento_endereco_logradouro, orcamento_endereco_bairro, orcamento_endereco_cidade, orcamento_endereco_uf, orcamento_endereco_cep, orcamento_endereco_numero
    dim orcamento_endereco_complemento, orcamento_endereco_email, orcamento_endereco_email_xml, orcamento_endereco_nome, orcamento_endereco_ddd_res
    dim orcamento_endereco_tel_res, orcamento_endereco_ddd_com, orcamento_endereco_tel_com, orcamento_endereco_ramal_com, orcamento_endereco_ddd_cel
    dim orcamento_endereco_tel_cel, orcamento_endereco_ddd_com_2, orcamento_endereco_tel_com_2, orcamento_endereco_ramal_com_2, orcamento_endereco_tipo_pessoa
    dim orcamento_endereco_cnpj_cpf, orcamento_endereco_contribuinte_icms_status, orcamento_endereco_produtor_rural_status, orcamento_endereco_ie
    dim orcamento_endereco_rg, orcamento_endereco_contato
    orcamento_endereco_logradouro = Trim(Request.Form("orcamento_endereco_logradouro"))
    orcamento_endereco_bairro = Trim(Request.Form("orcamento_endereco_bairro"))
    orcamento_endereco_cidade = Trim(Request.Form("orcamento_endereco_cidade"))
    orcamento_endereco_uf = Trim(Request.Form("orcamento_endereco_uf"))
    orcamento_endereco_cep = Trim(Request.Form("orcamento_endereco_cep"))
    orcamento_endereco_numero = Trim(Request.Form("orcamento_endereco_numero"))
    orcamento_endereco_complemento = Trim(Request.Form("orcamento_endereco_complemento"))
    orcamento_endereco_email = Trim(Request.Form("orcamento_endereco_email"))
    orcamento_endereco_email_xml = Trim(Request.Form("orcamento_endereco_email_xml"))
    orcamento_endereco_nome = Trim(Request.Form("orcamento_endereco_nome"))
    orcamento_endereco_ddd_res = Trim(Request.Form("orcamento_endereco_ddd_res"))
    orcamento_endereco_tel_res = Trim(Request.Form("orcamento_endereco_tel_res"))
    orcamento_endereco_ddd_com = Trim(Request.Form("orcamento_endereco_ddd_com"))
    orcamento_endereco_tel_com = Trim(Request.Form("orcamento_endereco_tel_com"))
    orcamento_endereco_ramal_com = Trim(Request.Form("orcamento_endereco_ramal_com"))
    orcamento_endereco_ddd_cel = Trim(Request.Form("orcamento_endereco_ddd_cel"))
    orcamento_endereco_tel_cel = Trim(Request.Form("orcamento_endereco_tel_cel"))
    orcamento_endereco_ddd_com_2 = Trim(Request.Form("orcamento_endereco_ddd_com_2"))
    orcamento_endereco_tel_com_2 = Trim(Request.Form("orcamento_endereco_tel_com_2"))
    orcamento_endereco_ramal_com_2 = Trim(Request.Form("orcamento_endereco_ramal_com_2"))
    orcamento_endereco_tipo_pessoa = Trim(Request.Form("orcamento_endereco_tipo_pessoa"))
    orcamento_endereco_cnpj_cpf = Trim(Request.Form("orcamento_endereco_cnpj_cpf"))
    orcamento_endereco_contribuinte_icms_status = Trim(Request.Form("orcamento_endereco_contribuinte_icms_status"))
    orcamento_endereco_produtor_rural_status = Trim(Request.Form("orcamento_endereco_produtor_rural_status"))
    orcamento_endereco_ie = Trim(Request.Form("orcamento_endereco_ie"))
    orcamento_endereco_rg = Trim(Request.Form("orcamento_endereco_rg"))
    orcamento_endereco_contato = Trim(Request.Form("orcamento_endereco_contato"))

	dim vendedor
	vendedor = ""
	
	dim alerta
	alerta=""
	
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

	strSql = "SELECT vendedor FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido = '" & usuario & "')"
	set r = cn.execute(strSql)
	if r.Eof then
		alerta = "FALHA AO LOCALIZAR O REGISTRO NO BANCO DE DADOS"
	else
		vendedor = Trim("" & r("vendedor"))
		end if

	if alerta = "" then
		if vendedor = "" then alerta = "NÃO HÁ NENHUM VENDEDOR DEFINIDO PARA ATENDÊ-LO"
		end if
	
'	CONSISTÊNCIAS P/ EMISSÃO DE NFe
	dim s_tabela_municipios_IBGE
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
		if rb_end_entrega = "S" then
		'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
			dim s_lista_sugerida_municipios
			dim v_lista_sugerida_municipios
			dim iCounterLista, iNumeracaoLista
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

	dim s_fabricante, s_produto, s_qtde, n_qtde, s_preco_lista, s_descricao
	dim n, intIdxProduto
	dim vProduto
	redim vProduto(0)
	set vProduto(0) = New cl_ITEM_PEDIDO
	vProduto(0).qtde = 0

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
				if n > max_qtde_itens then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O número de itens que está sendo cadastrado (" & CStr(n) & ") excede o máximo permitido por pedido (" & CStr(max_qtde_itens) & ")!!"
					end if
				end if
			end if 'if blnLojaHabilitadaProdCompostoECommerce
		end if 'if alerta = ""

	if alerta = "" then
	'	Validação do DDD dos telefones
		if orcamento_endereco_tipo_pessoa = ID_PF then
			if (orcamento_endereco_tel_res <> "") And (Len(orcamento_endereco_ddd_res) < 2) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "DDD inválido para o telefone (endereço de cobrança): " & orcamento_endereco_tel_res
				end if

			if (orcamento_endereco_tel_cel <> "") And (Len(orcamento_endereco_ddd_cel) < 2) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "DDD inválido para o telefone (endereço de cobrança): " & orcamento_endereco_tel_cel
				end if

			if (orcamento_endereco_tel_com <> "") And (Len(orcamento_endereco_ddd_com) < 2) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "DDD inválido para o telefone (endereço de cobrança): " & orcamento_endereco_tel_com
				end if
		else
			if (orcamento_endereco_tel_com <> "") And (Len(orcamento_endereco_ddd_com) < 2) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "DDD inválido para o telefone (endereço de cobrança): " & orcamento_endereco_tel_com
				end if
			
			if (orcamento_endereco_tel_com_2 <> "") And (Len(orcamento_endereco_ddd_com_2) < 2) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "DDD inválido para o telefone (endereço de cobrança): " & orcamento_endereco_tel_com_2
				end if
			end if 'if orcamento_endereco_tipo_pessoa = ID_PF

		if rb_end_entrega = "S" then
			if EndEtg_tipo_pessoa = ID_PF then
				if (EndEtg_tel_res <> "") And (Len(EndEtg_ddd_res) < 2) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "DDD inválido para o telefone (endereço de entrega): " & EndEtg_tel_res
					end if

				if (EndEtg_tel_cel <> "") And (Len(EndEtg_ddd_cel) < 2) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "DDD inválido para o telefone (endereço de entrega): " & EndEtg_tel_cel
					end if
			else
				if (EndEtg_tel_com <> "") And (Len(EndEtg_ddd_com) < 2) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "DDD inválido para o telefone (endereço de entrega): " & EndEtg_tel_com
					end if

				if (EndEtg_tel_com_2 <> "") And (Len(EndEtg_ddd_com_2) < 2) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "DDD inválido para o telefone (endereço de entrega): " & EndEtg_tel_com_2
					end if
				end if 'if EndEtg_tipo_pessoa = ID_PF
			end if 'if rb_end_entrega = "S"
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title><%=TITULO_JANELA_MODULO_ORCAMENTO%></title>
	</head>



<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
		<% if alerta <> "" then %>
		return;
		<% end if %>

		// Trata o problema em que os campos do formulário são limpos após retornar à esta página c/ o history.back() pela 2ª vez quando ocorre erro de consistência
		if (trim(fORC.c_FormFieldValues.value) != "") {
			stringToForm(fORC.c_FormFieldValues.value, $('#fORC'));
		}

	    $("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8
	    <%if blnLojaHabilitadaProdCompostoECommerce then%>
		// Testa se o form existe (pode não existir se foi exibida uma mensagem de aviso/erro)
		if ($("#fORC").length) {
			fORC.submit();
		}
        <%end if%>

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
var fCustoFinancFornecParcelamentoPopup;
var objAjaxCustoFinancFornecConsultaPreco;

function processaSelecaoCustoFinancFornecParcelamento(){};

function abreTabelaCustoFinancFornecParcelamento(intIndex){
var f, strUrl;
	f=fORC;
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
	f=fORC;
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
	f=fORC;
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
	f=fORC;
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
	f=fORC;
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
	f=fORC;
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
	f=fORC;
	if ((trim(f.c_fabricante[intIndice].value)=="")&&(trim(f.c_produto[intIndice].value)=="")) {
		f.c_qtde[intIndice].value="";
		f.c_descricao[intIndice].value="";
		f.c_preco_lista[intIndice].value="";
		}
}

function fORCConfirma( f ) {
var i, b, ha_item, strMsgErro;
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

/*	if (trim(f.midia.value)=='') {
		alert('Indique a forma pela qual o cliente conheceu a DIS!!');
		f.midia.focus();
		return;
		}
*/
		
	
	if (trim(f.vendedor.value)=='') {
		alert("Indique um vendedor!!");
		f.vendedor.focus();
		return;
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

	fORC.c_FormFieldValues.value = formToString($("#fORC"));

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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
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
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
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
<body onload="if (trim(fORC.c_fabricante[0].value)=='') fORC.c_fabricante[0].focus();">
<center>

<form id="fORC" name="fORC" method="post" action="OrcamentoNovoConsiste.asp">
<input type="hidden" name="c_loja" id="c_loja" value='<%=loja%>'>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=cliente_selecionado%>'>
<input type="hidden" name="vendedor" id="vendedor" value='<%=vendedor%>'>
<input type="hidden" name="rb_end_entrega" id="rb_end_entrega" value='<%=rb_end_entrega%>'>
<input type="hidden" name="EndEtg_endereco" id="EndEtg_endereco" value="<%=EndEtg_endereco%>">
<input type="hidden" name="EndEtg_endereco_numero" id="EndEtg_endereco_numero" value="<%=EndEtg_endereco_numero%>">
<input type="hidden" name="EndEtg_endereco_complemento" id="EndEtg_endereco_complemento" value="<%=EndEtg_endereco_complemento%>">
<input type="hidden" name="EndEtg_bairro" id="EndEtg_bairro" value="<%=EndEtg_bairro%>">
<input type="hidden" name="EndEtg_cidade" id="EndEtg_cidade" value="<%=EndEtg_cidade%>">
<input type="hidden" name="EndEtg_uf" id="EndEtg_uf" value="<%=EndEtg_uf%>">
<input type="hidden" name="EndEtg_cep" id="EndEtg_cep" value="<%=EndEtg_cep%>">
<input type="hidden" name="EndEtg_obs" id="EndEtg_obs" value='<%=EndEtg_obs%>'>
<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />
<input type="hidden" name="c_custoFinancFornecTipoParcelamento" id="c_custoFinancFornecTipoParcelamento" value='<%=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelas" id="c_custoFinancFornecQtdeParcelas" value='0'>

<!--  CAMPOS ADICIONAIS DO ENDERECO DE ENTREGA  -->
<input type="hidden" name="EndEtg_email" id="EndEtg_email" value="<%=EndEtg_email%>" />
<input type="hidden" name="EndEtg_email_xml" id="EndEtg_email_xml" value="<%=EndEtg_email_xml%>" />
<input type="hidden" name="EndEtg_nome" id="EndEtg_nome" value="<%=EndEtg_nome%>" />
<input type="hidden" name="EndEtg_ddd_res" id="EndEtg_ddd_res" value="<%=EndEtg_ddd_res%>" />
<input type="hidden" name="EndEtg_tel_res" id="EndEtg_tel_res" value="<%=EndEtg_tel_res%>" />
<input type="hidden" name="EndEtg_ddd_com" id="EndEtg_ddd_com" value="<%=EndEtg_ddd_com%>" />
<input type="hidden" name="EndEtg_tel_com" id="EndEtg_tel_com" value="<%=EndEtg_tel_com%>" />
<input type="hidden" name="EndEtg_ramal_com" id="EndEtg_ramal_com" value="<%=EndEtg_ramal_com%>" />
<input type="hidden" name="EndEtg_ddd_cel" id="EndEtg_ddd_cel" value="<%=EndEtg_ddd_cel%>" />
<input type="hidden" name="EndEtg_tel_cel" id="EndEtg_tel_cel" value="<%=EndEtg_tel_cel%>" />
<input type="hidden" name="EndEtg_ddd_com_2" id="EndEtg_ddd_com_2" value="<%=EndEtg_ddd_com_2%>" />
<input type="hidden" name="EndEtg_tel_com_2" id="EndEtg_tel_com_2" value="<%=EndEtg_tel_com_2%>" />
<input type="hidden" name="EndEtg_ramal_com_2" id="EndEtg_ramal_com_2" value="<%=EndEtg_ramal_com_2%>" />
<input type="hidden" name="EndEtg_tipo_pessoa" id="EndEtg_tipo_pessoa" value="<%=EndEtg_tipo_pessoa%>" />
<input type="hidden" name="EndEtg_cnpj_cpf" id="EndEtg_cnpj_cpf" value="<%=EndEtg_cnpj_cpf%>" />
<input type="hidden" name="EndEtg_contribuinte_icms_status" id="EndEtg_contribuinte_icms_status" value="<%=EndEtg_contribuinte_icms_status%>" />
<input type="hidden" name="EndEtg_produtor_rural_status" id="EndEtg_produtor_rural_status" value="<%=EndEtg_produtor_rural_status%>" />
<input type="hidden" name="EndEtg_ie" id="EndEtg_ie" value="<%=EndEtg_ie%>" />
<input type="hidden" name="EndEtg_rg" id="EndEtg_rg" value="<%=EndEtg_rg%>" />

<input type="hidden" name="orcamento_endereco_logradouro" id="orcamento_endereco_logradouro" value="<%=orcamento_endereco_logradouro%>" />
<input type="hidden" name="orcamento_endereco_bairro" id="orcamento_endereco_bairro" value="<%=orcamento_endereco_bairro%>" />
<input type="hidden" name="orcamento_endereco_cidade" id="orcamento_endereco_cidade" value="<%=orcamento_endereco_cidade%>" />
<input type="hidden" name="orcamento_endereco_uf" id="orcamento_endereco_uf" value="<%=orcamento_endereco_uf%>" />
<input type="hidden" name="orcamento_endereco_cep" id="orcamento_endereco_cep" value="<%=orcamento_endereco_cep%>" />
<input type="hidden" name="orcamento_endereco_numero" id="orcamento_endereco_numero" value="<%=orcamento_endereco_numero%>" />
<input type="hidden" name="orcamento_endereco_complemento" id="orcamento_endereco_complemento" value="<%=orcamento_endereco_complemento%>" />
<input type="hidden" name="orcamento_endereco_email" id="orcamento_endereco_email" value="<%=orcamento_endereco_email%>" />
<input type="hidden" name="orcamento_endereco_email_xml" id="orcamento_endereco_email_xml" value="<%=orcamento_endereco_email_xml%>" />
<input type="hidden" name="orcamento_endereco_nome" id="orcamento_endereco_nome" value="<%=orcamento_endereco_nome%>" />
<input type="hidden" name="orcamento_endereco_ddd_res" id="orcamento_endereco_ddd_res" value="<%=orcamento_endereco_ddd_res%>" />
<input type="hidden" name="orcamento_endereco_tel_res" id="orcamento_endereco_tel_res" value="<%=orcamento_endereco_tel_res%>" />
<input type="hidden" name="orcamento_endereco_ddd_com" id="orcamento_endereco_ddd_com" value="<%=orcamento_endereco_ddd_com%>" />
<input type="hidden" name="orcamento_endereco_tel_com" id="orcamento_endereco_tel_com" value="<%=orcamento_endereco_tel_com%>" />
<input type="hidden" name="orcamento_endereco_ramal_com" id="orcamento_endereco_ramal_com" value="<%=orcamento_endereco_ramal_com%>" />
<input type="hidden" name="orcamento_endereco_ddd_cel" id="orcamento_endereco_ddd_cel" value="<%=orcamento_endereco_ddd_cel%>" />
<input type="hidden" name="orcamento_endereco_tel_cel" id="orcamento_endereco_tel_cel" value="<%=orcamento_endereco_tel_cel%>" />
<input type="hidden" name="orcamento_endereco_ddd_com_2" id="orcamento_endereco_ddd_com_2" value="<%=orcamento_endereco_ddd_com_2%>" />
<input type="hidden" name="orcamento_endereco_tel_com_2" id="orcamento_endereco_tel_com_2" value="<%=orcamento_endereco_tel_com_2%>" />
<input type="hidden" name="orcamento_endereco_ramal_com_2" id="orcamento_endereco_ramal_com_2" value="<%=orcamento_endereco_ramal_com_2%>" />
<input type="hidden" name="orcamento_endereco_tipo_pessoa" id="orcamento_endereco_tipo_pessoa" value="<%=orcamento_endereco_tipo_pessoa%>" />
<input type="hidden" name="orcamento_endereco_cnpj_cpf" id="orcamento_endereco_cnpj_cpf" value="<%=orcamento_endereco_cnpj_cpf%>" />
<input type="hidden" name="orcamento_endereco_contribuinte_icms_status" id="orcamento_endereco_contribuinte_icms_status" value="<%=orcamento_endereco_contribuinte_icms_status%>" />
<input type="hidden" name="orcamento_endereco_produtor_rural_status" id="orcamento_endereco_produtor_rural_status" value="<%=orcamento_endereco_produtor_rural_status%>" />
<input type="hidden" name="orcamento_endereco_ie" id="orcamento_endereco_ie" value="<%=orcamento_endereco_ie%>" />
<input type="hidden" name="orcamento_endereco_rg" id="orcamento_endereco_rg" value="<%=orcamento_endereco_rg%>" />
<input type="hidden" name="orcamento_endereco_contato" id="orcamento_endereco_contato" value="<%=orcamento_endereco_contato%>" />

<input type="hidden" name="insert_request_guid" id="insert_request_guid" value="<%=gera_uid%>" />


<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>


<!--  I D E N T I F I C A Ç Ã O   D O   O R Ç A M E N T O -->
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pré-Pedido Novo</span></td>
</tr>
</table>
<br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0" <%if blnLojaHabilitadaProdCompostoECommerce then Response.Write "style='display:none;'"%>>
	<tr bgcolor="#FFFFFF">
	<td class="MB" align="left"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left"><span class="PLTe">Produto</span></td>
	<td class="MB" align="right"><span class="PLTd">Qtde</span></td>
	<td class="MB" align="left"><span class="PLTe">Descrição</span></td>
	<td class="MB" align="right"><span class="PLTd">VL Unit</span></td>
	</tr>
<% intIdxProduto = LBound(vProduto)-1 %>
<% for i=1 to max_qtde_itens
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
		<input name="c_fabricante" id="c_fabricante_<%=Cstr(i)%>" class="PLLe" maxlength="4" style="width:30px;" onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(<%=Cstr(i)%>!=1))) if (trim(this.value)=='') fORC.midia.focus(); else fORC.c_produto[<%=Cstr(i-1)%>].focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);trataLimpaLinha(<%=Cstr(i-1)%>);"
			<% 'A DECLARAÇÃO DA PROPRIEDADE VALUE APENAS SE HOUVER VALOR EVITA QUE O CAMPO SEJA LIMPO APÓS A SEGUNDA CHAMADA DO HISTORY.BACK() NA TELA SEGUINTE QUANDO OCORRE ERRO DE CONSISTÊNCIA E DESEJA-SE RETORNAR À ESTA TELA %>
			<% if s_fabricante <> "" then %>
			value="<%=s_fabricante%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="left">
		<input name="c_produto" id="c_produto_<%=Cstr(i)%>" class="PLLe" maxlength="8" style="width:60px;" onkeypress="if (digitou_enter(true)) fORC.c_qtde[<%=Cstr(i-1)%>].focus(); filtra_produto();" onblur="this.value=normaliza_produto(this.value);consultaPreco(<%=Cstr(i-1)%>);trataLimpaLinha(<%=Cstr(i-1)%>);"
			<% if s_produto <> "" then %>
			value="<%=s_produto%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde_<%=Cstr(i)%>" class="PLLd" maxlength="4" style="width:30px;" onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==fORC.c_qtde.length) fORC.midia.focus(); else fORC.c_fabricante[<%=Cstr(i)%>].focus();} filtra_numerico();"
			<% if s_qtde <> "" then %>
			value="<%=s_qtde%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="left">
		<input name="c_descricao" id="c_descricao_<%=Cstr(i)%>" class="PLLe" style="width:377px;" readonly tabindex=-1
			<% if s_descricao <> "" then %>
			value="<%=s_descricao%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="right">
		<input name="c_preco_lista" id="c_preco_lista_<%=Cstr(i)%>" class="PLLd" style="width:62px;" readonly tabindex=-1
			<% if s_preco_lista <> "" then %>
			value="<%=s_preco_lista%>"
			<% end if %>
			/>
	</td>
	<td align="left">
		&nbsp;<button type="button" name="bCustoFinancFornecParcelamento" id="bCustoFinancFornecParcelamento" style='width:50px;font-size:8pt;font-weight:bold;color:black;margin-bottom:1px;' class="Botao" onclick="abreTabelaCustoFinancFornecParcelamento(<%=i-1%>);"><%=SIMBOLO_MONETARIO%></button>
	</td>
	</tr>
<% next %>
</table>

<br>
<div <%if blnLojaHabilitadaProdCompostoECommerce then Response.Write "style='display:none;'"%>>
    <span class="PLLe">Forma de pagamento: </span>
    <input name="c_custoFinancFornecParcelamentoDescricao" id="c_custoFinancFornecParcelamentoDescricao" class="PLLe" style="width:115px;color:#0000CD;font-weight:bold;"
	    value="<%=descricaoCustoFinancFornecTipoParcelamento(COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA)%>">
</div>
<br>

<!-- ************   MÍDIA (INATIVO)  ************ -->
<!-- <table cellspacing="0" style="width:375px" style="display:none">
	<tr>
	<td width="100%" align="left"><p class="R">FORMA PELA QUAL CONHECEU A DIS</p><p class="C">
		<select id="midia" name="midia" style="margin-top:4pt; margin-bottom:4pt;width:370px;">
			<%'=midia_monta_itens_select(r_cliente.midia)%>
		</select>
		</p></td>
	</tr>
</table>
//-->


<!-- ************   SEPARADOR   ************ -->
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;display:none">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="749" cellspacing="0" <%if blnLojaHabilitadaProdCompostoECommerce then Response.Write "style='display:none;'"%>>
<tr>
	<% if blnLojaHabilitadaProdCompostoECommerce then %>
	<td align="left"><a name="bCANCELA" id="bCANCELA" href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<% else %>
	<td align="left"><a name="bCANCELA" id="bCANCELA" href="Resumo.asp" title="cancela o novo pré-pedido">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<% end if %>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fORCConfirma(fORC)" title="vai para a página de confirmação">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>
    <%if blnLojaHabilitadaProdCompostoECommerce then%>
<!-- Aguarde //-->
<table width="749" class="notPrint">
    <tr>
        <td style="text-align:right;vertical-align:middle;width:50%;">
            <img src="../IMAGEM/aguarde.gif" />
        </td>
        <td style="text-align:left;vertical-align:middle;width:50%">
            <span class="C">Redirecionando...</span>
        </td>
    </tr>
</table>
    <% end if %>
</center>
</body>

<% end if %>

</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

	if rs2.State <> 0 then rs2.Close
	set rs2 = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>