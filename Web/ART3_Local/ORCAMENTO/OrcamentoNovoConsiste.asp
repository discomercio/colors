<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/Global.asp"    -->
<%
'     =================================================
'	  O R C A M E N T O N O V O C O N S I S T E . A S P
'     =================================================
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

	dim s, i, j, n, usuario, loja, cliente_selecionado
	dim qtde_estoque_total_disponivel, qtde_estoque_total_global_disponivel, blnAchou, blnDesativado
	dim midia, vendedor, s_perc_RT
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	cliente_selecionado = Trim(request("cliente_selecionado"))
	if (cliente_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim blnLojaHabilitadaProdCompostoECommerce
	blnLojaHabilitadaProdCompostoECommerce = isLojaHabilitadaProdCompostoECommerce(loja)

	midia = Trim(Request.Form("midia"))
	vendedor = Trim(Request.Form("vendedor"))
	s_perc_RT = Trim(Request.Form("c_perc_RT"))

	dim strPercVlPedidoLimiteRA, percPercVlPedidoLimiteRA
	percPercVlPedidoLimiteRA = obtem_PercVlPedidoLimiteRA()
	strPercVlPedidoLimiteRA = formata_perc(percPercVlPedidoLimiteRA)

	dim r_orcamentista_e_indicador
	if alerta = "" then
		if Not le_orcamentista_e_indicador(usuario, r_orcamentista_e_indicador, msg_erro) then
			alerta = "Falha ao recuperar os dados cadastrais!!"
			end if
		end if
	
	dim r_cliente
	set r_cliente = New cl_CLIENTE
	if alerta = "" then
		if Not x_cliente_bd(cliente_selecionado, r_cliente) then
			alerta = "Falha ao recuperar dados do cliente"
			end if
		end if
	
	dim eh_cpf
	eh_cpf=(len(r_cliente.cnpj_cpf)=11)

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim c_custoFinancFornecTipoParcelamento, c_custoFinancFornecQtdeParcelas, coeficiente
	c_custoFinancFornecTipoParcelamento = Trim(Request.Form("c_custoFinancFornecTipoParcelamento"))
	c_custoFinancFornecQtdeParcelas = Trim(Request.Form("c_custoFinancFornecQtdeParcelas"))
	
	dim rb_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep, EndEtg_obs
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
	EndEtg_ddd_res = Trim(Request.Form("EndEtg_ddd_res"))
	EndEtg_tel_res = Trim(Request.Form("EndEtg_tel_res"))
	EndEtg_ddd_com = Trim(Request.Form("EndEtg_ddd_com"))
	EndEtg_tel_com = Trim(Request.Form("EndEtg_tel_com"))
	EndEtg_ramal_com = Trim(Request.Form("EndEtg_ramal_com"))
	EndEtg_ddd_cel = Trim(Request.Form("EndEtg_ddd_cel"))
	EndEtg_tel_cel = Trim(Request.Form("EndEtg_tel_cel"))
	EndEtg_ddd_com_2 = Trim(Request.Form("EndEtg_ddd_com_2"))
	EndEtg_tel_com_2 = Trim(Request.Form("EndEtg_tel_com_2"))
	EndEtg_ramal_com_2 = Trim(Request.Form("EndEtg_ramal_com_2"))
	EndEtg_tipo_pessoa = Trim(Request.Form("EndEtg_tipo_pessoa"))
	EndEtg_cnpj_cpf = Trim(Request.Form("EndEtg_cnpj_cpf"))
	EndEtg_contribuinte_icms_status = Trim(Request.Form("EndEtg_contribuinte_icms_status"))
	EndEtg_produtor_rural_status = Trim(Request.Form("EndEtg_produtor_rural_status"))
	EndEtg_ie = Trim(Request.Form("EndEtg_ie"))
	EndEtg_rg = Trim(Request.Form("EndEtg_rg"))

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

	dim s_fabricante, s_produto, s_descricao, s_descricao_html, s_qtde, s_readonly
	dim s_preco_lista, s_vl_TotalItem, m_TotalItem, m_TotalItemComRA, m_TotalDestePedido, m_TotalDestePedidoComRA
	dim s_campo_focus
	dim s_TotalDestePedidoComRA
	dim intIdx
	
	dim v_item
	redim v_item(0)
	set v_item(0) = New cl_ITEM_ORCAMENTO_NOVO
	n = Request.Form("c_produto").Count
	for i = 1 to n
		s=Trim(Request.Form("c_produto")(i))
		if s <> "" then
			if Trim(v_item(ubound(v_item)).produto) <> "" then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ITEM_ORCAMENTO_NOVO
				end if
			with v_item(ubound(v_item))
				.produto=Ucase(Trim(Request.Form("c_produto")(i)))
				s=retorna_so_digitos(Request.Form("c_fabricante")(i))
				.fabricante=normaliza_codigo(s, TAM_MIN_FABRICANTE)
				s = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
				end with
			end if
		next
	
'	VERIFICA CADA UM DOS PRODUTOS SELECIONADOS
	dim alerta, alerta_aux
	alerta=""

'	if midia = "" then
'		alerta = "Indique a forma pela qual o cliente conheceu a Bonshop."
	if vendedor = "" then
		alerta = "Selecione um vendedor."
	elseif (converte_numero(s_perc_RT)<0) Or (converte_numero(s_perc_RT)>100) then
		alerte = "Percentual de comissão inválido."
		end if

	if alerta = "" then
		if orcamento_endereco_nome = "" then
			if eh_cpf then
				alerta="DADOS CADASTRAIS: PREENCHA O NOME DO CLIENTE."
			else
				alerta="DADOS CADASTRAIS: PREENCHA A RAZÃO SOCIAL DO CLIENTE."
				end if
		elseif orcamento_endereco_logradouro = "" then
			alerta="DADOS CADASTRAIS: PREENCHA O ENDEREÇO."
		elseif Len(orcamento_endereco_logradouro) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
			alerta="DADOS CADASTRAIS: ENDEREÇO EXCEDE O TAMANHO MÁXIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(orcamento_endereco_logradouro)) & " CARACTERES<br>TAMANHO MÁXIMO: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " CARACTERES"
		elseif orcamento_endereco_numero = "" then
			alerta="DADOS CADASTRAIS: PREENCHA O NÚMERO DO ENDEREÇO."
		elseif orcamento_endereco_bairro = "" then
			alerta="DADOS CADASTRAIS: PREENCHA O BAIRRO."
		elseif orcamento_endereco_cidade = "" then
			alerta="DADOS CADASTRAIS: PREENCHA A CIDADE."
		elseif (orcamento_endereco_uf="") Or (Not uf_ok(orcamento_endereco_uf)) then
			alerta="DADOS CADASTRAIS: UF INVÁLIDA."
		elseif orcamento_endereco_cep = "" then
			alerta="DADOS CADASTRAIS: INFORME O CEP."
		elseif Not cep_ok(orcamento_endereco_cep) then
			alerta="DADOS CADASTRAIS: CEP INVÁLIDO."
		elseif Not ddd_ok(orcamento_endereco_ddd_res) then
			alerta="DADOS CADASTRAIS: DDD INVÁLIDO."
		elseif Not telefone_ok(orcamento_endereco_tel_res) then
			alerta="DADOS CADASTRAIS: TELEFONE RESIDENCIAL INVÁLIDO."
		elseif (orcamento_endereco_ddd_res <> "") And ((orcamento_endereco_tel_res = "")) then
			alerta="DADOS CADASTRAIS: PREENCHA O TELEFONE RESIDENCIAL."
		elseif (orcamento_endereco_ddd_res = "") And ((orcamento_endereco_tel_res <> "")) then
			alerta="DADOS CADASTRAIS: PREENCHA O DDD."
		elseif Not ddd_ok(orcamento_endereco_ddd_com) then
			alerta="DADOS CADASTRAIS: DDD INVÁLIDO."
		elseif Not telefone_ok(orcamento_endereco_tel_com) then
			alerta="DADOS CADASTRAIS: TELEFONE COMERCIAL INVÁLIDO."
		elseif (orcamento_endereco_ddd_com <> "") And ((orcamento_endereco_tel_com = "")) then
			alerta="DADOS CADASTRAIS: PREENCHA O TELEFONE COMERCIAL."
		elseif (orcamento_endereco_ddd_com = "") And ((orcamento_endereco_tel_com <> "")) then
			alerta="DADOS CADASTRAIS: PREENCHA O DDD."
		elseif eh_cpf And (orcamento_endereco_tel_res="") And (orcamento_endereco_tel_com="") And (orcamento_endereco_tel_cel="") then
			alerta="DADOS CADASTRAIS: PREENCHA PELO MENOS UM TELEFONE."
		elseif (Not eh_cpf) And (orcamento_endereco_tel_com="") And (orcamento_endereco_tel_com_2="") then
			alerta="DADOS CADASTRAIS: PREENCHA O TELEFONE."
		elseif (orcamento_endereco_ie="") And (orcamento_endereco_contribuinte_icms_status = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
			alerta="DADOS CADASTRAIS: PREENCHA A INSCRIÇÃO ESTADUAL."
			end if
		end if

'	CONSISTÊNCIAS P/ EMISSÃO DE NFe
    dim s_tabela_municipios_IBGE 
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
	'	I.E. É VÁLIDA? (somente para pj; pf pode ter outro estado de cadastro)
		if not eh_cpf and orcamento_endereco_ie <> "" then
			if Not isInscricaoEstadualValida(orcamento_endereco_ie, orcamento_endereco_uf) then
				alerta="Preencha a IE (Inscrição Estadual) com um número válido!!" & _
						"<br>" & "Certifique-se de que a UF informada corresponde à UF responsável pelo registro da IE."
				end if
			end if
		
	'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
		dim s_lista_sugerida_municipios
		dim v_lista_sugerida_municipios
		dim iCounterLista, iNumeracaoLista
		if Not consiste_municipio_IBGE_ok(orcamento_endereco_cidade, orcamento_endereco_uf, s_lista_sugerida_municipios, msg_erro) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			if msg_erro <> "" then
				alerta = alerta & msg_erro
			else
				alerta = alerta & "Município '" & orcamento_endereco_cidade & "' não consta na relação de municípios do IBGE para a UF de '" & orcamento_endereco_uf & "'!!"
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
								"			<p class='N'>" & "Relação de municípios de '" & orcamento_endereco_uf & "' que se iniciam com a letra '" & Ucase(left(orcamento_endereco_cidade,1)) & "'" & "</p>" & chr(13) & _
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
		if Not isTextoValido(orcamento_endereco_nome, s_caracteres_invalidos) then
			alerta="DADOS CADASTRAIS: O CAMPO 'NOME' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(orcamento_endereco_logradouro, s_caracteres_invalidos) then
			alerta="DADOS CADASTRAIS: O CAMPO 'ENDEREÇO' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(orcamento_endereco_numero, s_caracteres_invalidos) then
			alerta="DADOS CADASTRAIS: O CAMPO NÚMERO DO ENDEREÇO POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(orcamento_endereco_complemento, s_caracteres_invalidos) then
			alerta="DADOS CADASTRAIS: O CAMPO 'COMPLEMENTO' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(orcamento_endereco_bairro, s_caracteres_invalidos) then
			alerta="DADOS CADASTRAIS: O CAMPO 'BAIRRO' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(orcamento_endereco_cidade, s_caracteres_invalidos) then
			alerta="DADOS CADASTRAIS: O CAMPO 'CIDADE' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(orcamento_endereco_contato, s_caracteres_invalidos) then
			alerta="DADOS CADASTRAIS: O CAMPO 'CONTATO' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
			end if
		end if


	if alerta = "" then
		if rb_end_entrega = "" then
			alerta = "Não foi informado se o endereço de entrega é o mesmo do cadastro ou não."
		elseif rb_end_entrega = "S" then
			if EndEtg_endereco = "" then
				alerta="PREENCHA O ENDEREÇO DE ENTREGA."
			elseif Len(EndEtg_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
				alerta="ENDEREÇO DE ENTREGA EXCEDE O TAMANHO MÁXIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(EndEtg_endereco)) & " CARACTERES<br>TAMANHO MÁXIMO: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " CARACTERES"
			elseif EndEtg_endereco_numero = "" then
				alerta="PREENCHA O NÚMERO DO ENDEREÇO DE ENTREGA."
			elseif EndEtg_bairro = "" then
				alerta="PREENCHA O BAIRRO DO ENDEREÇO DE ENTREGA."
			elseif EndEtg_cidade = "" then
				alerta="PREENCHA A CIDADE DO ENDEREÇO DE ENTREGA."
			elseif (EndEtg_uf="") Or (Not uf_ok(EndEtg_uf)) then
				alerta="UF INVÁLIDA NO ENDEREÇO DE ENTREGA."
			elseif Not cep_ok(EndEtg_cep) then
				alerta="CEP INVÁLIDO NO ENDEREÇO DE ENTREGA."
				end if



            if alerta = "" and not eh_cpf and blnUsarMemorizacaoCompletaEnderecos then
                if EndEtg_tipo_pessoa <> "PJ" and EndEtg_tipo_pessoa <> "PF" then
                    alerta = "Necessário escolher Pessoa Jurídica ou Pessoa Física no Endereço de entrega!!"
    			elseif EndEtg_nome = "" then
                    alerta = "Preencha o nome/razão social no endereço de entrega!!"
                    end if 
	
                if alerta = "" and EndEtg_tipo_pessoa = "PJ" then
                    '//Campos PJ: 
                    if EndEtg_cnpj_cpf = "" or not cnpj_ok(EndEtg_cnpj_cpf) then
                        alerta = "Endereço de entrega: CNPJ inválido!!"
                    elseif EndEtg_contribuinte_icms_status = "" then
                        alerta = "Endereço de entrega: selecione o tipo de contribuinte de ICMS!!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and EndEtg_ie = "" then
                        alerta = "Endereço de entrega: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) and InStr(EndEtg_ie, "ISEN") > 0 then 
                        alerta = "Endereço de entrega: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!"
                    elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and InStr(EndEtg_ie, "ISEN") > 0 then 
                        alerta = "Endereço de entrega: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!"
                    'telefones PJ:
                    'EndEtg_ddd_com
                    'EndEtg_tel_com
                    'EndEtg_ramal_com
                    'EndEtg_ddd_com_2
                    'EndEtg_tel_com_2
                    'EndEtg_ramal_com_2
                    elseif not ddd_ok(EndEtg_ddd_com) then
                        alerta = "Endereço de entrega: DDD inválido!!"
                    elseif not telefone_ok(EndEtg_tel_com) then
                        alerta = "Endereço de entrega: telefone inválido!!"
                    elseif EndEtg_ddd_com = "" and EndEtg_tel_com <> "" then
                        alerta = "Endereço de entrega: preencha o DDD do telefone."
                    elseif EndEtg_tel_com = "" and EndEtg_ddd_com <> "" then
                        alerta = "Endereço de entrega: preencha o telefone."

                    elseif not ddd_ok(EndEtg_ddd_com_2) then
                        alerta = "Endereço de entrega: DDD inválido!!"
                    elseif not telefone_ok(EndEtg_tel_com_2) then
                        alerta = "Endereço de entrega: telefone inválido!!"
                    elseif EndEtg_ddd_com_2 = "" and EndEtg_tel_com_2 <> "" then
                        alerta = "Endereço de entrega: preencha o DDD do telefone."
                    elseif EndEtg_tel_com_2 = "" and EndEtg_ddd_com_2 <> "" then
                        alerta = "Endereço de entrega: preencha o telefone."
                        end if 
                    end if 

                if alerta = "" and EndEtg_tipo_pessoa <> "PJ" then
                    '//campos PF
                    if EndEtg_cnpj_cpf = "" or not cpf_ok(EndEtg_cnpj_cpf) then
                        alerta = "Endereço de entrega: CPF inválido!!"
                    elseif EndEtg_produtor_rural_status = "" then
                        alerta = "Endereço de entrega: informe se o cliente é produtor rural ou não!!"
                    elseif converte_numero(EndEtg_produtor_rural_status) <> converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
                        if converte_numero(EndEtg_contribuinte_icms_status) <> converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                            alerta = "Endereço de entrega: para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!!"
                        elseif EndEtg_contribuinte_icms_status = "" then
                            alerta = "Endereço de entrega: informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!"
                        elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and EndEtg_ie = "" then
                            alerta = "Endereço de entrega: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!"
                        elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) and InStr(EndEtg_ie, "ISEN") > 0 then 
                            alerta = "Endereço de entrega: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!"
                        elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) and InStr(EndEtg_ie, "ISEN") > 0 then 
                            alerta = "Endereço de entrega: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!"
                        elseif converte_numero(EndEtg_contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) and EndEtg_ie <> "" then 
                            alerta = "Endereço de entrega: se o Contribuinte ICMS é isento, o campo IE deve ser vazio!"
                            end if
                        end if

                    if alerta = "" then
                        'telefones PF:
                        'EndEtg_ddd_res
                        'EndEtg_tel_res
                        'EndEtg_ddd_cel
                        'EndEtg_tel_cel
                        if not ddd_ok(retorna_so_digitos(EndEtg_ddd_res)) then
                            alerta = "Endereço de entrega: DDD inválido!!"
                        elseif not telefone_ok(retorna_so_digitos(EndEtg_tel_res)) then
                            alerta = "Endereço de entrega: telefone inválido!!"
                        elseif EndEtg_ddd_res <> "" or EndEtg_tel_res <> "" then
                            if EndEtg_ddd_res = "" then
                                alerta = "Endereço de entrega: preencha o DDD!!"
                            elseif EndEtg_tel_res = "" then
                                alerta = "Endereço de entrega: preencha o telefone!!"
                                end if
                            end if
                        end if

                    if alerta = "" then
                        if not ddd_ok(retorna_so_digitos(EndEtg_ddd_cel)) then
                            alerta = "Endereço de entrega: DDD inválido!!"
                        elseif not telefone_ok(retorna_so_digitos(EndEtg_tel_cel)) then
                            alerta = "Endereço de entrega: telefone inválido!!"
                        elseif EndEtg_ddd_cel = "" and EndEtg_tel_cel <> "" then
                            alerta = "Endereço de entrega: preencha o DDD do celular."
                        elseif EndEtg_tel_cel = "" and EndEtg_ddd_cel <> "" then
                            alerta = "Endereço de entrega: preencha o número do celular."
                            end if
                        end if

                    end if

		        if alerta = "" and EndEtg_ie <> "" then
			        if Not isInscricaoEstadualValida(EndEtg_ie, EndEtg_uf) then
				        alerta="Endereço de entrega: preencha a IE (Inscrição Estadual) com um número válido!!" & _
						        "<br>" & "Certifique-se de que a UF do endereço de entrega corresponde à UF responsável pelo registro da IE."
				        end if
			        end if

                end if

			end if
		end if
	
	if alerta="" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .qtde <= 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": quantidade " & cstr(.qtde) & " é inválida."
					end if

				for j=Lbound(v_item) to (i-1)
					if (.produto = v_item(j).produto) And (.fabricante = v_item(j).fabricante) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": linha " & renumera_com_base1(Lbound(v_item),i) & " repete o mesmo produto da linha " & renumera_com_base1(Lbound(v_item),j) & "."
						exit for
						end if
					next

				s = "SELECT " & _
						"*" & _
					" FROM t_PRODUTO" & _
						" INNER JOIN t_PRODUTO_LOJA" & _
							" ON (t_PRODUTO.fabricante=t_PRODUTO_LOJA.fabricante) AND (t_PRODUTO.produto=t_PRODUTO_LOJA.produto)" & _
					" WHERE" & _
						" (t_PRODUTO.fabricante='" & .fabricante & "')" & _
						" AND (t_PRODUTO.produto='" & .produto & "')" & _
						" AND (loja='" & loja & "')"
				set rs = cn.execute(s)
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " NÃO está cadastrado."
				else
					if Ucase(Trim("" & rs("vendavel"))) <> "S" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " NÃO está disponível para venda."
					elseif .qtde > rs("qtde_max_venda") then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": quantidade " & cstr(.qtde) & " excede o máximo permitido."
					else
						.preco_lista = rs("preco_lista")
						.margem = rs("margem")
						.desc_max = rs("desc_max")
						.comissao = rs("comissao")
						.preco_fabricante = rs("preco_fabricante")
						.vl_custo2 = rs("vl_custo2")
						.descricao = Trim("" & rs("descricao"))
						.descricao_html = Trim("" & rs("descricao_html"))
						.ean = Trim("" & rs("ean"))
						.grupo = Trim("" & rs("grupo"))
                        .subgrupo = Trim("" & rs("subgrupo"))
						.peso = rs("peso")
						.qtde_volumes = Trim("" & rs("qtde_volumes"))
						.cubagem = rs("cubagem")
						.ncm = Trim("" & rs("ncm"))
						.cst = Trim("" & rs("cst"))
						.descontinuado = Trim("" & rs("descontinuado"))
						end if
					end if
				rs.Close
				end with
			next
		end if

	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			s = "SELECT " & _
					"*" & _
				" FROM t_EC_PRODUTO_COMPOSTO_ITEM" & _
				" WHERE" & _
					" (fabricante_composto = '" & v_item(i).fabricante & "')" & _
					" AND (produto_composto = '" & v_item(i).produto & "')" & _
				" ORDER BY" & _
					" fabricante_item," & _
					" produto_item"
			set rs = cn.execute(s)
			if Not rs.Eof then
				s = ""
				do while Not rs.Eof
					if s <> "" then s = s & ", "
					s = s & Trim("" & rs("produto_item"))
					rs.MoveNext
					loop
				alerta=texto_add_br(alerta)
				alerta=alerta & "O código de produto " & v_item(i).produto & " do fabricante " & v_item(i).fabricante & " é somente um código auxiliar para agrupar os produtos " & s & " e não pode ser usado diretamente no pré-pedido!!"
				end if
			next
		end if

	if alerta = "" then
		if (c_custoFinancFornecTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA) And _
		   (c_custoFinancFornecTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) And _
		   (c_custoFinancFornecTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) then
			alerta = "A forma de pagamento não foi informada (à vista, com entrada, sem entrada)."
			end if
		end if
		
	if alerta = "" then
		if (c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) Or _
		   (c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) then
			if converte_numero(c_custoFinancFornecQtdeParcelas) <= 0 then
				alerta = "Não foi informada a quantidade de parcelas para a forma de pagamento selecionada (" & descricaoCustoFinancFornecTipoParcelamento(c_custoFinancFornecTipoParcelamento) &  ")"
				end if
			end if
		end if
	
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if c_custoFinancFornecTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA then
					coeficiente = 1
				else
					s = "SELECT " & _
							"*" & _
						" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
						" WHERE" & _
							" (fabricante = '" & .fabricante & "')" & _
							" AND (tipo_parcelamento = '" & c_custoFinancFornecTipoParcelamento & "')" & _
							" AND (qtde_parcelas = " & c_custoFinancFornecQtdeParcelas & ")"
					set rs = cn.execute(s)
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Opção de parcelamento não disponível para fornecedor " & .fabricante & ": " & decodificaCustoFinancFornecQtdeParcelas(c_custoFinancFornecTipoParcelamento, c_custoFinancFornecQtdeParcelas) & " parcela(s)"
					else
						coeficiente = converte_numero(rs("coeficiente"))
						.preco_lista=converte_numero(formata_moeda(coeficiente*.preco_lista))
						end if
					end if
				end with
			next
		end if

	if alerta = "" then
		if Not isTextoValido(EndEtg_endereco, s_caracteres_invalidos) then
			alerta="O CAMPO 'ENDEREÇO DE ENTREGA' POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_endereco_numero, s_caracteres_invalidos) then
			alerta="O CAMPO NÚMERO DO ENDEREÇO DE ENTREGA POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_endereco_complemento, s_caracteres_invalidos) then
			alerta="O CAMPO COMPLEMENTO DO ENDEREÇO DE ENTREGA POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_bairro, s_caracteres_invalidos) then
			alerta="O CAMPO BAIRRO DO ENDEREÇO DE ENTREGA POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_cidade, s_caracteres_invalidos) then
			alerta="O CAMPO CIDADE DO ENDEREÇO DE ENTREGA POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
		elseif Not isTextoValido(EndEtg_nome, s_caracteres_invalidos) then
			alerta="O CAMPO NOME DO ENDEREÇO DE ENTREGA POSSUI UM OU MAIS CARACTERES INVÁLIDOS: " & s_caracteres_invalidos
			end if
		end if


'	PARÂMETRO SOBRE COMO DEVE SER ANALISADA A DISPONIBILIDADE DO ESTOQUE DE VENDA (SEGUINDO AS REGRAS DE CONSUMO DO ESTOQUE POR CD OU PELO ESTOQUE GLOBAL)
'	O MOTIVO QUE JUSTIFICA A ANÁLISE PELO ESTOQUE GLOBAL É P/ O INSTALADOR NÃO PENSAR QUE A MENSAGEM DE PRODUTO SEM PRESENÇA NO ESTOQUE SIGNIFIQUE QUE NÃO HÁ
'	PRODUTO DISPONÍVEL P/ ENTREGA, JÁ QUE A LOGÍSTICA SE ENCARREGA DE FAZER A TRANSFERÊNCIA ENTRE CD'S P/ ATENDER O PEDIDO
	dim rCDEG
	set rCDEG = get_registro_t_parametro(ID_PARAMETRO_Flag_Orcamento_ConsisteDisponibilidadeEstoqueGlobal)

'	LÓGICA P/ CONSUMO DO ESTOQUE (REGRA DEFINIDA POR PRODUTO)
	dim tipo_pessoa
	dim descricao_tipo_pessoa
	tipo_pessoa = multi_cd_regra_determina_tipo_pessoa(r_cliente.tipo, r_cliente.contribuinte_icms_status, r_cliente.produtor_rural_status)
	descricao_tipo_pessoa = descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa)

	dim id_nfe_emitente_selecao_manual
	dim vProdRegra, iRegra, iCD, iItem, idxItem, qtde_CD_ativo
	id_nfe_emitente_selecao_manual = 0
	
	if alerta="" then
		'PREPARA O VETOR PARA RECUPERAR AS REGRAS DE CONSUMO DO ESTOQUE ASSOCIADAS AOS PRODUTOS
		redim vProdRegra(0)
		inicializa_cl_PEDIDO_SELECAO_PRODUTO_REGRA vProdRegra(UBound(vProdRegra))
		for i=LBound(v_item) to UBound(v_item)
			if vProdRegra(UBound(vProdRegra)).produto <> "" then
				redim preserve vProdRegra(UBound(vProdRegra)+1)
				inicializa_cl_PEDIDO_SELECAO_PRODUTO_REGRA vProdRegra(UBound(vProdRegra))
				end if
			vProdRegra(UBound(vProdRegra)).fabricante = v_item(i).fabricante
			vProdRegra(UBound(vProdRegra)).produto =v_item(i).produto
			next
		
		'RECUPERA AS REGRAS DE CONSUMO DO ESTOQUE ASSOCIADAS AOS PRODUTOS
		if Not obtemCtrlEstoqueProdutoRegra(r_cliente.uf, r_cliente.tipo, r_cliente.contribuinte_icms_status, r_cliente.produtor_rural_status, vProdRegra, msg_erro) then
			alerta = "Falha ao tentar obter a(s) regra(s) de consumo do estoque"
			if msg_erro <> "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & msg_erro
				end if
			end if
		end if 'if alerta=""

	if alerta="" then
		'VERIFICA SE HOUVE ERRO NA LEITURA DAS REGRAS DE CONSUMO DO ESTOQUE ASSOCIADAS AOS PRODUTOS
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			if Trim(vProdRegra(iRegra).produto) <> "" then
				if Not vProdRegra(iRegra).st_regra_ok then
					if Trim(vProdRegra(iRegra).msg_erro) <> "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & vProdRegra(iRegra).msg_erro
					else
						alerta=texto_add_br(alerta)
						alerta=alerta & "Falha desconhecida na leitura da regra de consumo do estoque para o produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " (UF: '" & r_cliente.uf & "', tipo de pessoa: '" & descricao_tipo_pessoa & "')"
						end if
					end if
				end if
			next
		end if 'if alerta=""

	if alerta="" then
		'VERIFICA SE AS REGRAS ASSOCIADAS AOS PRODUTOS ESTÃO OK
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			if Trim(vProdRegra(iRegra).produto) <> "" then
				if converte_numero(vProdRegra(iRegra).regra.id) = 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " não possui regra de consumo do estoque associada"
				elseif vProdRegra(iRegra).regra.st_inativo = 1 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " está desativada"
				elseif vProdRegra(iRegra).regra.regraUF.st_inativo = 1 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " está bloqueada para a UF '" & r_cliente.uf & "'"
				elseif vProdRegra(iRegra).regra.regraUF.regraPessoa.st_inativo = 1 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " está bloqueada para clientes '" & descricao_tipo_pessoa & "' da UF '" & r_cliente.uf & "'"
				elseif converte_numero(vProdRegra(iRegra).regra.regraUF.regraPessoa.spe_id_nfe_emitente) = 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " não especifica nenhum CD para aguardar produtos sem presença no estoque para clientes '" & descricao_tipo_pessoa & "' da UF '" & r_cliente.uf & "'"
				else
					qtde_CD_ativo = 0
					for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
						if converte_numero(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente) > 0 then
							if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0 then
								qtde_CD_ativo = qtde_CD_ativo + 1
								end if
							end if
						next
					if qtde_CD_ativo = 0 then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Regra de consumo do estoque '" & vProdRegra(iRegra).regra.apelido & "' associada ao produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " não especifica nenhum CD ativo para clientes '" & descricao_tipo_pessoa & "' da UF '" & r_cliente.uf & "'"
						end if
					end if
				end if
			next
		end if 'if alerta=""
	
	'NO CASO DE SELEÇÃO MANUAL DO CD, VERIFICA SE O CD SELECIONADO ESTÁ HABILITADO EM TODAS AS REGRAS
	if alerta="" then
		if id_nfe_emitente_selecao_manual <> 0 then
			alerta_aux = ""
			for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
				blnAchou = False
				blnDesativado = False
				if Trim(vProdRegra(iRegra).produto) <> "" then
					for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
						if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual then
							blnAchou = True
							if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 1 then blnDesativado = True
							exit for
							end if
						next
					end if

				if Not blnAchou then
					alerta_aux=texto_add_br(alerta_aux)
					alerta_aux=alerta_aux & "Produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & ": regra '" & vProdRegra(iRegra).regra.apelido & "' (Id=" & vProdRegra(iRegra).regra.id & ") não permite o CD '" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_selecao_manual) & "'"
				elseif blnDesativado then
					alerta_aux=texto_add_br(alerta_aux)
					alerta_aux=alerta_aux & "Produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & ": regra '" & vProdRegra(iRegra).regra.apelido & "' (Id=" & vProdRegra(iRegra).regra.id & ") define o CD '" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_selecao_manual) & "' como 'desativado'"
					end if
				next
			
			if alerta_aux <> "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O CD selecionado manualmente não pode ser usado devido aos seguintes motivos:"
				alerta=texto_add_br(alerta)
				alerta=alerta & alerta_aux
				end if
			end if
		end if

	dim erro_produto_indisponivel
	if alerta="" then
		'OBTÉM DISPONIBILIDADE DO PRODUTO NO ESTOQUE
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			if Trim(vProdRegra(iRegra).produto) <> "" then
				for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
					if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
						( (id_nfe_emitente_selecao_manual = 0) Or (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) ) then
						'VERIFICA SE O CD ESTÁ HABILITADO
						if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0 then
							idxItem = Lbound(v_item) - 1
							for iItem=Lbound(v_item) to Ubound(v_item)
								if (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) then
									idxItem = iItem
									exit for
									end if
								next
							if idxItem < Lbound(v_item) then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Falha ao localizar o produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto & " na lista de produtos a ser processada"
							else
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.fabricante = v_item(idxItem).fabricante
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.produto = v_item(idxItem).produto
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.descricao = v_item(idxItem).descricao
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.descricao_html = v_item(idxItem).descricao_html
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = v_item(idxItem).qtde
								vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque = 0
								if Not estoque_verifica_disponibilidade_integral_v2(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente, vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque) then
									alerta=texto_add_br(alerta)
									alerta=alerta & "Falha ao tentar consultar disponibilidade no estoque do produto (" & vProdRegra(iRegra).fabricante & ")" & vProdRegra(iRegra).produto
									end if
								end if
							end if
						end if

					if alerta <> "" then exit for
					next
				end if

			if alerta <> "" then exit for
			next
		end if 'if alerta=""

'	HÁ PRODUTO C/ ESTOQUE INSUFICIENTE (SOMANDO-SE O ESTOQUE DE TODAS AS EMPRESAS CANDIDATAS)
	erro_produto_indisponivel = False
	if alerta="" then
		for iItem=Lbound(v_item) to Ubound(v_item)
			if Trim(v_item(iItem).produto) <> "" then
				qtde_estoque_total_disponivel = 0
				qtde_estoque_total_global_disponivel = -1
				for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
					if Trim(vProdRegra(iRegra).produto) <> "" then
						for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
							if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
								( (id_nfe_emitente_selecao_manual = 0) Or (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) ) then
								'VERIFICA SE O CD ESTÁ HABILITADO
								if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0 then
									if (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) then
										qtde_estoque_total_disponivel = qtde_estoque_total_disponivel + vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque
										if qtde_estoque_total_global_disponivel = -1 then qtde_estoque_total_global_disponivel = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque_global
										end if
									end if
								end if
							next
						end if
					next

				if Trim("" & rCDEG.campo_inteiro) = "1" then
					if qtde_estoque_total_global_disponivel = -1 then
						v_item(iItem).qtde_estoque_total_disponivel = 0
					else
						v_item(iItem).qtde_estoque_total_disponivel = qtde_estoque_total_global_disponivel
						end if
				else
					v_item(iItem).qtde_estoque_total_disponivel = qtde_estoque_total_disponivel
					end if

				if v_item(iItem).qtde > v_item(iItem).qtde_estoque_total_disponivel then
					erro_produto_indisponivel = True
					end if
				end if
			next
		end if 'if alerta=""

'	ANALISA A QUANTIDADE DE PEDIDOS QUE SERÃO CADASTRADOS (AUTO-SPLIT)
'	INICIALIZA O CAMPO 'qtde_solicitada', POIS ELE IRÁ CONTROLAR A QUANTIDADE A SER ALOCADA NO ESTOQUE DE CADA EMPRESA
	if alerta = "" then
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
				vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = 0
				next
			next
		end if 'if alerta=""

'	REALIZA A ANÁLISE DA QUANTIDADE DE PEDIDOS NECESSÁRIA (AUTO-SPLIT)
	dim qtde_a_alocar
	if alerta = "" then
		for iItem=Lbound(v_item) to Ubound(v_item)
			if Trim(v_item(iItem).produto) <> "" then
			'	OS CD'S ESTÃO ORDENADOS DE ACORDO C/ A PRIORIZAÇÃO DEFINIDA PELA REGRA DE CONSUMO DO ESTOQUE
			'	SE O PRIMEIRO CD HABILITADO NÃO PUDER ATENDER INTEGRALMENTE A QUANTIDADE SOLICITADA DO PRODUTO,
			'	A QUANTIDADE RESTANTE SERÁ CONSUMIDA DOS DEMAIS CD'S.
			'	SE HOUVER ALGUMA QUANTIDADE RESIDUAL P/ FICAR NA LISTA DE PRODUTOS SEM PRESENÇA NO ESTOQUE:
			'		1) SELEÇÃO AUTOMÁTICA DE CD: A QUANTIDADE PENDENTE FICARÁ ALOCADA NO CD DEFINIDO P/ TAL
			'		2) SELEÇÃO MANUAL DE CD: A QUANTIDADE PENDENTE FICARÁ ALOCADA NO CD SELECIONADO MANUALMENTE
				qtde_a_alocar = v_item(iItem).qtde
				for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
					if qtde_a_alocar = 0 then exit for

					if Trim(vProdRegra(iRegra).produto) <> "" then
						for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
							if qtde_a_alocar = 0 then exit for

							if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
								( (id_nfe_emitente_selecao_manual = 0) Or (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) ) then
								'VERIFICA SE O CD ESTÁ HABILITADO
								if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 0 then
									if (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) then
										if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque >= qtde_a_alocar then
										'	HÁ QUANTIDADE DISPONÍVEL SUFICIENTE PARA INTEGRALMENTE
											vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = qtde_a_alocar
											qtde_a_alocar = 0
										elseif vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque > 0 then
										'	A QUANTIDADE DISPONÍVEL NO ESTOQUE É INSUFICIENTE P/ ATENDER INTEGRALMENTE À QUANTIDADE SOLICITADA,
										'	PORTANTO, A QUANTIDADE DISPONÍVEL NESTE CD SERÁ CONSUMIDA P/ ATENDER PARCIALMENTE À REQUISIÇÃO E A
										'	QUANTIDADE REMANESCENTE SERÁ ATENDIDA PELO PRÓXIMO CD DA LISTA OU ENTÃO SERÁ COLOCADA NA LISTA DE
										'	PRODUTOS SEM PRESENÇA NO ESTOQUE DO CD SELECIONADO P/ TAL NA REGRA.
											vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque
											qtde_a_alocar = qtde_a_alocar - vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_estoque
											end if
										end if
									end if
								end if
							next
						end if
					next

			'	RESTOU SALDO A ALOCAR NA LISTA DE PRODUTOS SEM PRESENÇA NO ESTOQUE?
				if qtde_a_alocar > 0 then
				'	LOCALIZA E ALOCA A QUANTIDADE PENDENTE:
				'		1) SELEÇÃO AUTOMÁTICA DE CD: A QUANTIDADE PENDENTE FICARÁ ALOCADA NO CD DEFINIDO P/ TAL
				'		2) SELEÇÃO MANUAL DE CD: A QUANTIDADE PENDENTE FICARÁ ALOCADA NO CD SELECIONADO MANUALMENTE
					for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
						if qtde_a_alocar = 0 then exit for

						if Trim(vProdRegra(iRegra).produto) <> "" then
							for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
								if qtde_a_alocar = 0 then exit for

								if id_nfe_emitente_selecao_manual = 0 then
									'MODO DE SELEÇÃO AUTOMÁTICO
									if ( (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) ) And _
										(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
										(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = vProdRegra(iRegra).regra.regraUF.regraPessoa.spe_id_nfe_emitente) then
										vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada + qtde_a_alocar
										qtde_a_alocar = 0
										exit for
										end if
								else
									'MODO DE SELEÇÃO MANUAL
									if ( (vProdRegra(iRegra).fabricante = v_item(iItem).fabricante) And (vProdRegra(iRegra).produto = v_item(iItem).produto) ) And _
										(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
										(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) then
										vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada = vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada + qtde_a_alocar
										qtde_a_alocar = 0
										exit for
										end if
									end if
								next
							end if
						next
					end if

			'	HOUVE FALHA EM ALOCAR A QUANTIDADE REMANESCENTE?
				if qtde_a_alocar > 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Falha ao processar a alocação de produtos no estoque: restaram " & qtde_a_alocar & " unidades do produto (" & v_item(iItem).fabricante & ")" & v_item(iItem).produto & " que não puderam ser alocados na lista de produtos sem presença no estoque de nenhum CD"
					end if
				end if
			next
		end if 'if alerta=""

'	CONTAGEM DE EMPRESAS QUE SERÃO USADAS NO AUTO-SPLIT, OU SEJA, A QUANTIDADE DE PEDIDOS QUE SERÁ CADASTRADA, JÁ QUE CADA PEDIDO SE REFERE AO ESTOQUE DE UMA EMPRESA
	dim qtde_empresa_selecionada, lista_empresa_selecionada
	qtde_empresa_selecionada = 0
	lista_empresa_selecionada = ""
	if alerta = "" then
		for iRegra=LBound(vProdRegra) to UBound(vProdRegra)
			if Trim(vProdRegra(iRegra).produto) <> "" then
				for iCD=LBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD)
					if (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente > 0) And _
						( (id_nfe_emitente_selecao_manual = 0) Or (vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = id_nfe_emitente_selecao_manual) ) then
						if vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).estoque.qtde_solicitada > 0 then
							s = "|" & vProdRegra(iRegra).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente & "|"
							if Instr(lista_empresa_selecionada, s) = 0 then
								qtde_empresa_selecionada = qtde_empresa_selecionada + 1
								lista_empresa_selecionada = lista_empresa_selecionada & s
								end if
							end if
						end if
					next
				end if
			next
		end if 'if alerta=""


'	HÁ ALGUM PRODUTO DESCONTINUADO?
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			if Trim(v_item(i).produto) <> "" then
				if Ucase(Trim(v_item(i).descontinuado)) = "S" then
					if v_item(i).qtde > v_item(i).qtde_estoque_total_disponivel then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto (" & v_item(i).fabricante & ")" & v_item(i).produto & " consta como 'descontinuado' e não há mais saldo suficiente no estoque para atender à quantidade solicitada."
						end if
					end if
				end if
			next
		end if

'	HÁ MENSAGENS DE ALERTA SOBRE OS PRODUTOS P/ SEREM EXIBIDAS?
	dim strScriptMsgAlerta
	dim strMensagem
	strScriptMsgAlerta = _
		"<script language='JavaScript'>" & chr(13) & _
		"var Pd = new Array();" & chr(13) & _
		"Pd[0] = new oPd('','','','');" & chr(13)

	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			s = "SELECT" & _
					" tpa.fabricante," & _
					" tpa.produto," & _
					" mensagem," & _
					" descricao" & _
				" FROM t_PRODUTO_X_ALERTA tpa" & _
					" INNER JOIN t_ALERTA_PRODUTO tap ON (tpa.id_alerta=tap.apelido)" & _
					" INNER JOIN t_PRODUTO tp ON (tpa.fabricante = tp.fabricante) AND (tpa.produto = tp.produto)" & _
				" WHERE" & _
					" (tpa.fabricante = '" & v_item(i).fabricante & "')" & _
					" AND (tpa.produto = '" & v_item(i).produto & "')" & _
					" AND (tap.ativo = 'S')" & _
				" ORDER BY" & _
					" tpa.dt_cadastro," & _
					" tpa.id_alerta"
			set rs = cn.execute(s)
			do while Not rs.Eof
				strMensagem=Trim("" & rs("mensagem"))
				strMensagem=Replace(strMensagem, chr(10), "")
				strMensagem=Replace(strMensagem, chr(13), "\n")
				strScriptMsgAlerta = strScriptMsgAlerta & _
					"Pd[Pd.length]=new oPd('" & Trim("" & rs("fabricante")) & "'" & _
					",'" & Trim("" & rs("produto")) & "'" & _
					",'" & filtra_nome_identificador(Trim("" & rs("descricao"))) & "'" & _
					",'" & filtra_nome_identificador(strMensagem) & "'" & _
					");" & chr(13)
				rs.MoveNext
				loop
			next
		end if
		
	strScriptMsgAlerta = strScriptMsgAlerta & _
		"</script>" & chr(13)
	
	dim bloquear_cadastramento_quando_produto_indiponivel
	bloquear_cadastramento_quando_produto_indiponivel = False
	if ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then bloquear_cadastramento_quando_produto_indiponivel = False
	
	dim strScriptJS
	strScriptJS = "<script language='JavaScript'>" & chr(13)
	if erro_produto_indisponivel then
		strScriptJS = strScriptJS & "var erro_produto_indisponivel = true;" & chr(13)
	else
		strScriptJS = strScriptJS & "var erro_produto_indisponivel = false;" & chr(13)
		end if
	if bloquear_cadastramento_quando_produto_indiponivel then
		strScriptJS = strScriptJS & "var bloquear_cadastramento_quando_produto_indiponivel = true;" & chr(13)
	else
		strScriptJS = strScriptJS & "var bloquear_cadastramento_quando_produto_indiponivel = false;" & chr(13)
		end if

	if blnLojaHabilitadaProdCompostoECommerce then
		strScriptJS = strScriptJS & "var formata_perc_desconto = formata_perc_2dec;" & chr(13)
	else
		strScriptJS = strScriptJS & "var formata_perc_desconto = formata_perc_desc;" & chr(13)
		end if

	strScriptJS = strScriptJS & _
				  "</script>" & chr(13)

	dim qtdeHiddenColumnTabProd
	qtdeHiddenColumnTabProd = 0
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


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<%=strScriptJS%>

<script type="text/javascript">
	$(function() {
		$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8
		<% if r_cliente.tipo = ID_PF then %>
		$(".TR_FP_PU").hide();
		$(".TR_FP_PSE").hide();
		<% end if %>
		<% if r_cliente.tipo = ID_PJ then %>
		$(".TR_FP_PSE").hide();
		<% end if %>
		$(".tdProdObs").hide(); <% qtdeHiddenColumnTabProd = qtdeHiddenColumnTabProd+1 %>
		$(".tdGarInd").hide();
		$(".rbGarIndNao").attr('checked', 'checked');
		$("#c_data_previsao_entrega").hUtilUI('datepicker_padrao');

        $("input[name = 'rb_etg_imediata']").change(function () {
            if ($("input[name='rb_etg_imediata']:checked").val() == '<%=COD_ETG_IMEDIATA_NAO%>') {
                $("#c_data_previsao_entrega").prop("readonly", false);
                $("#c_data_previsao_entrega").prop("disabled", false);
                $("#c_data_previsao_entrega").datepicker("enable");
            }
            else {
                $("#c_data_previsao_entrega").val("");
                $("#c_data_previsao_entrega").prop("readonly", true);
                $("#c_data_previsao_entrega").prop("disabled", true);
                $("#c_data_previsao_entrega").datepicker("disable");
            }
        });
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
var objAjaxCustoFinancFornecConsultaPreco;

function processaFormaPagtoDefault() {
var f, i;
	f=fORC;
	for (i=0; i<fORC.rb_forma_pagto.length; i++) {
		if (fORC.rb_forma_pagto[i].checked) {
			fORC.rb_forma_pagto[i].click();
			break;
			}
		}

	f.c_custoFinancFornecParcelamentoDescricao.value=descricaoCustoFinancFornecTipoParcelamento(f.c_custoFinancFornecTipoParcelamento.value);
	if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (1+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}
	else if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (0+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}
}

function trataRespostaAjaxCustoFinancFornecSincronizaPrecos() {
var f, strResp, i, j, xmlDoc, oNodes;
var strFabricante,strProduto, strStatus, strPrecoLista, strMsgErro, strCodigoErro;
var percDesc,vlLista,vlVenda,strMsgErroAlert;
	f=fORC;
	strMsgErroAlert="";
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
								    if (f.c_preco_lista[j].value == f.c_vl_unitario[j].value) f.c_vl_unitario[j].value=strPrecoLista;
								    if (f.c_permite_RA_status.value == "1")
								    {
								    	if (f.c_preco_lista[j].value == f.c_vl_NF[j].value) f.c_vl_NF[j].value=strPrecoLista;
								    }
									f.c_preco_lista[j].value = strPrecoLista;
									f.c_preco_lista[j].style.color = "black";
									$(f.c_preco_lista[j]).removeClass('CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE');
									}
								}
							}
						}
					else {
					//  Código do Erro
						oNodes=xmlDoc.getElementsByTagName("codigo_erro")[i];
						if (oNodes.childNodes.length > 0) strCodigoErro=oNodes.childNodes[0].nodeValue; else strCodigoErro="";
						if (strCodigoErro==null) strCodigoErro="";
					//  Mensagem de Erro
						oNodes=xmlDoc.getElementsByTagName("msg_erro")[i];
						if (oNodes.childNodes.length > 0) strMsgErro=oNodes.childNodes[0].nodeValue; else strMsgErro="";
						if (strMsgErro==null) strMsgErro="";
						for (j=0; j<f.c_fabricante.length; j++) {
						//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
						//	(apesar de que isso não será aceito pelas consistências que serão feitas).
							if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								f.c_preco_lista[j].style.color = COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE;
								$(f.c_preco_lista[j]).addClass('CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE');
								}
							}
						if (strMsgErroAlert!="") strMsgErroAlert+="\n\n";
						strMsgErroAlert+="Falha ao consultar o preço do produto " + strProduto + "!!\n" + strMsgErro;
						}
					}
				}
			catch (e)
				{
				alert("Falha na consulta do preço!!\n"+e.message);
				}
			}
			
		if (strMsgErroAlert!="") alert(strMsgErroAlert);
		
		recalcula_total_todas_linhas(); 
		recalcula_RA();
		recalcula_parcelas();
		
		window.status="Concluído";
		$("#divAjaxRunning").hide();
		}
}

function recalculaCustoFinanceiroPrecoLista() {
var f, i, strListaProdutos, strUrl, strOpcaoFormaPagto;
	f=fORC;
	objAjaxCustoFinancFornecConsultaPreco=GetXmlHttpObject();
	if (objAjaxCustoFinancFornecConsultaPreco==null) {
		alert("O browser NÃO possui suporte ao AJAX!!");
		return;
		}
		
	strListaProdutos="";
	for (i=0; i<f.c_fabricante.length; i++) {
		if ((trim(f.c_fabricante[i].value)!="")&&(trim(f.c_produto[i].value)!="")) {
			if (strListaProdutos!="") strListaProdutos+=";";
			strListaProdutos += f.c_fabricante[i].value + "|" + f.c_produto[i].value;
			}
		}
	if (strListaProdutos=="") return;
	
//  Converte as opções de forma de pagamento do pedido em uma opção que possa tratada pela tabela de custo financeiro
	strOpcaoFormaPagto="";
	for (i=0; i<fORC.rb_forma_pagto.length; i++) {
		if (fORC.rb_forma_pagto[i].checked) {
			strOpcaoFormaPagto=f.rb_forma_pagto[i].value;
			break;
			}
		}
	if (strOpcaoFormaPagto=="") return;
	
	if (strOpcaoFormaPagto==COD_FORMA_PAGTO_A_VISTA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA;
		f.c_custoFinancFornecQtdeParcelas.value='0';
		}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELA_UNICA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value='1';
		}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_CARTAO) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value=f.c_pc_qtde.value;
		}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value=f.c_pc_maquineta_qtde.value;
	}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value=f.c_pce_prestacao_qtde.value;
		}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value=(converte_numero(f.c_pse_demais_prest_qtde.value)+1).toString();
		}
	else {
		f.c_custoFinancFornecTipoParcelamento.value="";
		f.c_custoFinancFornecQtdeParcelas.value="";
		}
		
	if (trim(f.c_custoFinancFornecQtdeParcelas.value)=="") return;

//  Não consulta novamente se for a mesma consulta anterior
	if ((f.c_custoFinancFornecTipoParcelamento.value==f.c_custoFinancFornecTipoParcelamentoUltConsulta.value)&&
	    (f.c_custoFinancFornecQtdeParcelas.value==f.c_custoFinancFornecQtdeParcelasUltConsulta.value)) return;
	    
	f.c_custoFinancFornecTipoParcelamentoUltConsulta.value=f.c_custoFinancFornecTipoParcelamento.value;
	f.c_custoFinancFornecQtdeParcelasUltConsulta.value=f.c_custoFinancFornecQtdeParcelas.value;

	f.c_custoFinancFornecParcelamentoDescricao.value=descricaoCustoFinancFornecTipoParcelamento(f.c_custoFinancFornecTipoParcelamento.value);
	if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (1+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}
	else if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (0+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}

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

function oPd(fabricante, produto, descricao, mensagem) {
	this.fabricante = fabricante;
	this.produto = produto;
	this.descricao = descricao;
	this.mensagem = mensagem;
}

function recalcula_parcelas_forma_pagto_selecionada() {
	var s_selecionado, vt, s_vt;
	s_selecionado = $('input[name=rb_forma_pagto]:checked').val();
	if (s_selecionado == COD_FORMA_PAGTO_PARCELA_UNICA) {
		vt = fp_vl_total_pedido();
		s_vt = formata_moeda(vt);
		$('#c_pu_valor').val(s_vt);
	}
	else if (s_selecionado == COD_FORMA_PAGTO_PARCELADO_CARTAO) {
		pc_calcula_valor_parcela();
	}
	else if (s_selecionado == COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) {
		pc_maquineta_calcula_valor_parcela();
	}
}

// RETORNA O VALOR TOTAL DO PEDIDO A SER USADO P/ CALCULAR A FORMA DE PAGAMENTO
function fp_vl_total_pedido( ) {
var f,i,mTotVenda,mTotNF;
	f=fORC;
	mTotVenda=0;
	for (i=0; i<f.c_qtde.length; i++) mTotVenda=mTotVenda+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_unitario[i].value);
	mTotNF = 0;
	if (f.c_permite_RA_status.value == '1') {
		for (i = 0; i < f.c_qtde.length; i++) mTotNF = mTotNF + converte_numero(f.c_qtde[i].value) * converte_numero(f.c_vl_NF[i].value);
	}
//  Retorna total de preço NF (tem valor de NF, ou seja, pedido c/ RA)?
	if (mTotNF > 0) {
		return mTotNF;
		}
//  Retorna total de preço de venda
	else {
		return mTotVenda;
		}
}

// PARCELA ÚNICA
function pu_atualiza_valor( ){
var f,vt;
	f=fORC;
	vt=fp_vl_total_pedido();
	f.c_pu_valor.value=formata_moeda(vt);
}

// PARCELADO NO CARTÃO (INTERNET)
function pc_calcula_valor_parcela( ){
var f,n,t;
	f=fORC;
	if (trim(f.c_pc_qtde.value)=='') return;
	n=converte_numero(f.c_pc_qtde.value);
	if (n<=0) return;
	t=fp_vl_total_pedido();
	p=t/n;
	f.c_pc_valor.value=formata_moeda(p);
}

// PARCELADO NO CARTÃO (MAQUINETA)
function pc_maquineta_calcula_valor_parcela( ){
	var f,n,t;
	f=fORC;
	if (trim(f.c_pc_maquineta_qtde.value)=='') return;
	n=converte_numero(f.c_pc_maquineta_qtde.value);
	if (n<=0) return;
	t=fp_vl_total_pedido();
	p=t/n;
	f.c_pc_maquineta_valor.value=formata_moeda(p);
}

// PARCELADO COM ENTRADA
function pce_preenche_sugestao_intervalo() {
var f;
	f=fORC;
	if (converte_numero(trim(f.c_pce_prestacao_periodo.value))>0) return;
	f.c_pce_prestacao_periodo.value='30';
}

function pce_calcula_valor_parcela( ){
var f,n,e,t;
	f=fORC;
	t=fp_vl_total_pedido();
	if (trim(f.c_pce_entrada_valor.value)=='') return;
	e=converte_numero(f.c_pce_entrada_valor.value);
	if (e<=0) return;
	if (trim(f.c_pce_prestacao_qtde.value)=='') return;
	n=converte_numero(f.c_pce_prestacao_qtde.value);
	if (n<=0) return;
	p=(t-e)/n;
	f.c_pce_prestacao_valor.value=formata_moeda(p);
}

// PARCELADO SEM ENTRADA
function pse_preenche_sugestao_intervalo() {
var f;
	f=fORC;
	if (converte_numero(trim(f.c_pse_demais_prest_periodo.value))>0) return;
	f.c_pse_demais_prest_periodo.value='30';
}

function pse_calcula_valor_parcela( ){
var f,n,e,t;
	f=fORC;
	t=fp_vl_total_pedido();
	if (trim(f.c_pse_prim_prest_valor.value)=='') return;
	e=converte_numero(f.c_pse_prim_prest_valor.value);
	if (e<=0) return;
	if (trim(f.c_pse_demais_prest_qtde.value)=='') return;
	n=converte_numero(f.c_pse_demais_prest_qtde.value);
	if (n<=0) return;
	p=(t-e)/n;
	f.c_pse_demais_prest_valor.value=formata_moeda(p);
}

function pce_sugestao_forma_pagto( ) {
var f, p, s, i, n;
	f=fORC;
	f.c_forma_pagto.value="";
	p=converte_numero(f.c_pce_prestacao_periodo.value);
	if (p<=0) return;
	n=converte_numero(f.c_pce_prestacao_qtde.value);
	if (n<=0) return;
	s='0';
	for (i=1; i<=n; i++) {
		s=s+'/';
		s=s+formata_inteiro(i*p);
		}
	f.c_forma_pagto.value=s;
}

function pse_sugestao_forma_pagto( ) {
var f, p1, p2, s, i, n;
	f=fORC;
	f.c_forma_pagto.value="";
	p1=converte_numero(f.c_pse_prim_prest_apos.value);
	if (p1<=0) return;
	p2=converte_numero(f.c_pse_demais_prest_periodo.value);
	if (p2<=0) return;
	n=converte_numero(f.c_pse_demais_prest_qtde.value);
	if (n<=0) return;
	s=formata_inteiro(p1);
	for (i=1; i<=n; i++) {
		s=s+'/';
		s=s+formata_inteiro(i*p2);
		}
	f.c_forma_pagto.value=s;
}

function restaura_cor_desconto( ) {
var f,i;
	f=fORC;
	for (i=0; i < f.c_desc.length; i++) {
		if (converte_numero(f.c_desc[i].value)<0) f.c_desc[i].style.color="red"; else f.c_desc[i].style.color="green";
		}
}

function trata_edicao_RA(index) {
	var f;
	f = fPED;
	if (f.c_permite_RA_status.value != '1') f.c_vl_NF[index].value = f.c_vl_unitario[index].value;
}

function calcula_desconto(idx) {
	var f, s, i, m, d, m_lista, m_unit;
	f = fORC;
	if (f.c_produto[idx].value == "") return;
	d = converte_numero(f.c_desc[idx].value);
	m_lista = converte_numero(f.c_preco_lista[idx].value);
	m_unit = m_lista - (m_lista * d / 100);
	f.c_vl_unitario[idx].value = formata_moeda(m_unit);
	s = formata_moeda(parseInt(f.c_qtde[idx].value) * m_unit);
	if (f.c_vl_total[idx].value != s) f.c_vl_total[idx].value = s;
	m = 0;
	for (i = 0; i < f.c_vl_total.length; i++) m = m + converte_numero(f.c_vl_total[i].value);
	s = formata_moeda(m);
	if (f.c_total_geral.value != s) f.c_total_geral.value = s;
}

function recalcula_total_linha( id ) {
	var idx, m, m_lista, m_unit, d, f, i, s;
	f=fORC;
	idx=parseInt(id)-1;
	if (f.c_produto[idx].value=="") return;
	m_lista=converte_numero(f.c_preco_lista[idx].value);
	m_unit=converte_numero(f.c_vl_unitario[idx].value);
	if (m_lista==0) d=0; else d=100*(m_lista-m_unit)/m_lista;
	if (d<0) f.c_desc[idx].style.color="red"; else f.c_desc[idx].style.color="green";
	if (d == 0) s = ""; else s = formata_perc_desconto(d);
	if (f.c_desc[idx].value!=s) f.c_desc[idx].value=s;
	s=formata_moeda(parseInt(f.c_qtde[idx].value)*m_unit);
	if (f.c_vl_total[idx].value!=s) f.c_vl_total[idx].value=s;
	m=0;
	for (i=0; i<f.c_vl_total.length; i++) m=m+converte_numero(f.c_vl_total[i].value);
	s=formata_moeda(m);
	if (f.c_total_geral.value!=s) f.c_total_geral.value=s;
}

function recalcula_total( linha ) {
var idx, m, m_lista, m_unit, d, f, i, s;
	f=fORC;
	idx=parseInt(linha)-1;
	if (f.c_produto[idx].value=="") return;
	m_lista=converte_numero(f.c_preco_lista[idx].value);
	m_unit=converte_numero(f.c_vl_unitario[idx].value);
	if (m_lista==0) d=0; else d=100*(m_lista-m_unit)/m_lista;
	if (d<0) f.c_desc[idx].style.color="red"; else f.c_desc[idx].style.color="green";
	if (d==0) s=""; else s=formata_perc_desconto(d);
	if (f.c_desc[idx].value!=s) f.c_desc[idx].value=s;
	s=formata_moeda(parseInt(f.c_qtde[idx].value)*m_unit);
	if (f.c_vl_total[idx].value!=s) f.c_vl_total[idx].value=s;
	m=0;
	for (i=0; i<f.c_vl_total.length; i++) m=m+converte_numero(f.c_vl_total[i].value);
	s=formata_moeda(m);
	if (f.c_total_geral.value!=s) f.c_total_geral.value=s;
}

function recalcula_total_todas_linhas() {
var f,i,vt,m_lista,m_unit,d,m,s;
	f = fORC;
	vt=0;
	for (i=0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value)!="") {
			m_lista=converte_numero(f.c_preco_lista[i].value);
			m_unit=converte_numero(f.c_vl_unitario[i].value);
			if (m_lista==0) d=0; else d=100*(m_lista-m_unit)/m_lista;
			if (d<0) f.c_desc[i].style.color="red"; else f.c_desc[i].style.color="green";
			if (d==0) s=""; else s=formata_perc_desconto(d);
			if (f.c_desc[i].value!=s) f.c_desc[i].value=s;
			m=parseInt(f.c_qtde[i].value)*m_unit;
			f.c_vl_total[i].value=formata_moeda(m);
			vt=vt+m;
			}
		}
	f.c_total_geral.value=formata_moeda(vt);
}

function recalcula_RA( ) {
var f,i,mTotVenda,mTotNF;
	f = fORC;
	if (f.c_permite_RA_status.value != '1') return;
	mTotVenda=0;
	for (i=0; i<f.c_vl_total.length; i++) mTotVenda=mTotVenda+converte_numero(f.c_vl_total[i].value);
	mTotNF=0;
	for (i=0; i<f.c_qtde.length; i++) mTotNF=mTotNF+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_NF[i].value);
	f.c_total_NF.value = formata_moeda(mTotNF);
	f.c_total_RA.value = formata_moeda(mTotNF-mTotVenda);
	if (mTotNF >=mTotVenda) f.c_total_RA.style.color="green"; else f.c_total_RA.style.color="red";
}

function consiste_forma_pagto( blnComAvisos ) {
var f,idx,vtNF,vtFP,ve,ni,nip,n,vp;
var MAX_ERRO_ARREDONDAMENTO = 0.1;
	f=fORC;
	vtNF=fp_vl_total_pedido();
	vtFP=0;
	idx=-1;
	
//	À Vista
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.op_av_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento!!');
				f.op_av_forma_pagto.focus();
				}
			return false;
			}
		return true;
		}

//	Parcela Única
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.op_pu_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento da parcela única!!');
				f.op_pu_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pu_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da parcela única!!');
				f.c_pu_valor.focus();
				}
			return false;
			}
		ve=converte_numero(f.c_pu_valor.value);
		vtFP=ve;
		if (ve<=0) {
			if (blnComAvisos) {
				alert('Valor da parcela única é inválido!!');
				f.c_pu_valor.focus();
				}
			return false;
			}
		if (trim(f.c_pu_vencto_apos.value)=='') {
			if (blnComAvisos) {
				alert('Indique o intervalo de vencimento da parcela única!!');
				f.c_pu_vencto_apos.focus();
				}
			return false;
			}
		nip=converte_numero(f.c_pu_vencto_apos.value);
		if (nip<=0) {
			if (blnComAvisos) {
				alert('Intervalo de vencimento da parcela única é inválido!!');
				f.c_pu_vencto_apos.focus();
				}
			return false;
			}
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
			    alert('Há divergência entre o valor total do pré-pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
				f.c_pu_valor.focus();
				}
			return false;
			}
		return true;
		}

//	Parcelado no cartão (internet)
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.c_pc_qtde.value)=='') {
			if (blnComAvisos) {
				alert('Indique a quantidade de parcelas!!');
				f.c_pc_qtde.focus();
				}
			return false;
			}
		n=converte_numero(f.c_pc_qtde.value);
		if (n < 1) {
			if (blnComAvisos) {
				alert('Quantidade de parcelas inválida!!');
				f.c_pc_qtde.focus();
				}
			return false;
			}
		if (trim(f.c_pc_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da parcela!!');
				f.c_pc_valor.focus();
				}
			return false;
			}
		vp=converte_numero(f.c_pc_valor.value);
		if (vp<=0) {
			if (blnComAvisos) {
				alert('Valor de parcela inválido!!');
				f.c_pc_valor.focus();
				}
			return false;
			}
		vtFP=n*vp;
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
			    alert('Há divergência entre o valor total do pré-pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
				f.c_pc_valor.focus();
				}
			return false;
			}
		return true;
		}

	//	Parcelado no cartão (maquineta)
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.c_pc_maquineta_qtde.value)=='') {
			if (blnComAvisos) {
				alert('Indique a quantidade de parcelas!!');
				f.c_pc_maquineta_qtde.focus();
			}
			return false;
		}
		n=converte_numero(f.c_pc_maquineta_qtde.value);
		if (n < 1) {
			if (blnComAvisos) {
				alert('Quantidade de parcelas inválida!!');
				f.c_pc_maquineta_qtde.focus();
			}
			return false;
		}
		if (trim(f.c_pc_maquineta_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da parcela!!');
				f.c_pc_maquineta_valor.focus();
			}
			return false;
		}
		vp=converte_numero(f.c_pc_maquineta_valor.value);
		if (vp<=0) {
			if (blnComAvisos) {
				alert('Valor de parcela inválido!!');
				f.c_pc_maquineta_valor.focus();
			}
			return false;
		}
		vtFP=n*vp;
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
				alert('Há divergência entre o valor total do pré-pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
				f.c_pc_maquineta_valor.focus();
			}
			return false;
		}
		return true;
	}

//	Parcelado com entrada
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.op_pce_entrada_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento da entrada!!');
				f.op_pce_entrada_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pce_entrada_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da entrada!!');
				f.c_pce_entrada_valor.focus();
				}
			return false;
			}
		ve=converte_numero(f.c_pce_entrada_valor.value);
		if (ve<=0) {
			if (blnComAvisos) {
				alert('Valor da entrada inválido!!');
				f.c_pce_entrada_valor.focus();
				}
			return false;
			}
		if (trim(f.op_pce_prestacao_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento das prestações!!');
				f.op_pce_prestacao_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pce_prestacao_qtde.value)=='') {
			if (blnComAvisos) {
				alert('Indique a quantidade de prestações!!');
				f.c_pce_prestacao_qtde.focus();
				}
			return false;
			}
		n=converte_numero(f.c_pce_prestacao_qtde.value);
		if (n<=0) {
			if (blnComAvisos) {
				alert('Quantidade de prestações inválida!!');
				f.c_pce_prestacao_qtde.focus();
				}
			return false;
			}
		if (trim(f.c_pce_prestacao_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da prestação!!');
				f.c_pce_prestacao_valor.focus();
				}
			return false;
			}
		vp=converte_numero(f.c_pce_prestacao_valor.value);
		if (vp<=0) {
			if (blnComAvisos) {
				alert('Valor de prestação inválido!!');
				f.c_pce_prestacao_valor.focus();
				}
			return false;
			}
		vtFP=ve+(n*vp);
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
			    alert('Há divergência entre o valor total do pré-pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
				f.c_pce_prestacao_valor.focus();
				}
			return false;
			}
		if (trim(f.c_pce_prestacao_periodo.value)=='') {
			if (blnComAvisos) {
				alert('Indique o intervalo de vencimento entre as parcelas!!');
				f.c_pce_prestacao_periodo.focus();
				}
			return false;
			}
		ni=converte_numero(f.c_pce_prestacao_periodo.value);
		if (ni<=0) {
			if (blnComAvisos) {
				alert('Intervalo de vencimento inválido!!');
				f.c_pce_prestacao_periodo.focus();
				}
			return false;
			}
		return true;
		}

//	Parcelado sem entrada
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.op_pse_prim_prest_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento da 1ª prestação!!');
				f.op_pse_prim_prest_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pse_prim_prest_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da 1ª prestação!!');
				f.c_pse_prim_prest_valor.focus();
				}
			return false;
			}
		ve=converte_numero(f.c_pse_prim_prest_valor.value);
		if (ve<=0) {
			if (blnComAvisos) {
				alert('Valor da 1ª prestação inválido!!');
				f.c_pse_prim_prest_valor.focus();
				}
			return false;
			}
		if (trim(f.c_pse_prim_prest_apos.value)=='') {
			if (blnComAvisos) {
				alert('Indique o intervalo de vencimento da 1ª parcela!!');
				f.c_pse_prim_prest_apos.focus();
				}
			return false;
			}
		nip=converte_numero(f.c_pse_prim_prest_apos.value);
		if (nip<=0) {
			if (blnComAvisos) {
				alert('Intervalo de vencimento da 1ª parcela é inválido!!');
				f.c_pse_prim_prest_apos.focus();
				}
			return false;
			}
		if (trim(f.op_pse_demais_prest_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento das demais prestações!!');
				f.op_pse_demais_prest_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pse_demais_prest_qtde.value)=='') {
			if (blnComAvisos) {
				alert('Indique a quantidade das demais prestações!!');
				f.c_pse_demais_prest_qtde.focus();
				}
			return false;
			}
		n=converte_numero(f.c_pse_demais_prest_qtde.value);
		if (n<=0) {
			if (blnComAvisos) {
				alert('Quantidade de prestações inválida!!');
				f.c_pse_demais_prest_qtde.focus();
				}
			return false;
			}
		if (trim(f.c_pse_demais_prest_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor das demais prestações!!');
				f.c_pse_demais_prest_valor.focus();
				}
			return false;
			}
		vp=converte_numero(f.c_pse_demais_prest_valor.value);
		if (vp<=0) {
			if (blnComAvisos) {
				alert('Valor de prestação inválido!!');
				f.c_pse_demais_prest_valor.focus();
				}
			return false;
			}
		vtFP=ve+(n*vp);
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
			    alert('Há divergência entre o valor total do pré-pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
				f.c_pse_demais_prest_valor.focus();
				}
			return false;
			}
		if (trim(f.c_pse_demais_prest_periodo.value)=='') {
			if (blnComAvisos) {
				alert('Indique o intervalo de vencimento entre as parcelas!!');
				f.c_pse_demais_prest_periodo.focus();
				}
			return false;
			}
		ni=converte_numero(f.c_pse_demais_prest_periodo.value);
		if (ni<=0) {
			if (blnComAvisos) {
				alert('Intervalo de vencimento inválido!!');
				f.c_pse_demais_prest_periodo.focus();
				}
			return false;
			}
		return true;
		}
		
	if (blnComAvisos) {
		// Nenhuma forma de pagamento foi escolhida
		alert('Indique a forma de pagamento!!');
		}
		
	return false;
}

function recalcula_parcelas() {
    var f, idx;
    f = fORC;
    idx=-1;

    idx++;
    idx++;
    if (f.rb_forma_pagto[idx].checked) {
        pu_atualiza_valor();
        return;
    }

    idx++;
    if (f.rb_forma_pagto[idx].checked) {
        pc_calcula_valor_parcela();
        return;
    }

    idx++;
    if (f.rb_forma_pagto[idx].checked) {
        pce_calcula_valor_parcela();
        return;
    }

    idx++;
    if (f.rb_forma_pagto[idx].checked) {
        pse_calcula_valor_parcela();
        return;
    }
  
}

function fORCConfirma( f ) {
var i,s,blnFlag,vlAux,strMsgErro;
var blnConfirmaDifRAeValores=false;

	recalcula_total_todas_linhas();

	recalcula_RA();

	s = "" + f.c_obs1.value;
	if (s.length > MAX_TAM_OBS1) {
		alert('Conteúdo de "Observações " excede em ' + (s.length-MAX_TAM_OBS1) + ' caracteres o tamanho máximo de ' + MAX_TAM_OBS1 + '!!');
		f.c_obs1.focus();
		return;
		}
	
	s = "" + f.c_forma_pagto.value;
	if (s.length > MAX_TAM_FORMA_PAGTO) {
		alert('Conteúdo de "Forma de Pagamento" excede em ' + (s.length-MAX_TAM_FORMA_PAGTO) + ' caracteres o tamanho máximo de ' + MAX_TAM_FORMA_PAGTO + '!!');
		f.c_forma_pagto.focus();
		return;
		}

//  Consiste a nova versão da forma de pagamento
	if (!consiste_forma_pagto(true)) return;

//	Limita o RA a um percentual do valor do pedido
	if (f.c_permite_RA_status.value == '1') {
		if (converte_numero(f.c_PercVlPedidoLimiteRA.value) != 0) {
			vlAux = (converte_numero(f.c_PercVlPedidoLimiteRA.value) / 100) * converte_numero(f.c_total_geral.value);
			if (converte_numero(f.c_total_RA.value) > vlAux) {
				alert('O valor total de RA excede o limite permitido!!');
				return;
			}
		}

		if (blnConfirmaDifRAeValores) {
			if (converte_numero(f.c_total_RA.value) != 0) {
				if (!confirm("O valor do RA é de " + SIMBOLO_MONETARIO + " " + formata_moeda(converte_numero(f.c_total_RA.value)) + "\nContinua?")) return;
			}
		}
	}
	
	blnFlag=false;
	for (i=0; i < f.rb_etg_imediata.length; i++) {
		if (f.rb_etg_imediata[i].checked) blnFlag=true;
		}
	if (!blnFlag) {
		alert('Selecione uma opção para o campo "Entrega Imediata"');
		return;
		}

    if (f.rb_etg_imediata[0].checked) {
        if (trim(f.c_data_previsao_entrega.value) == "") {
            alert("Informe a data de previsão de entrega!");
            f.c_data_previsao_entrega.focus();
            return;
        }

        if (!isDate(f.c_data_previsao_entrega)) {
            alert("Data de previsão de entrega é inválida!");
            f.c_data_previsao_entrega.focus();
            return;
        }

        if (retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(f.c_data_previsao_entrega.value)) <= retorna_so_digitos(formata_ddmmyyyy_yyyymmdd('<%=formata_data(Date)%>'))) {
            alert("Data de previsão de entrega deve ser uma data futura!");
            f.c_data_previsao_entrega.focus();
            return;
        }
    }

	blnFlag=false;
	for (i=0; i < f.rb_bem_uso_consumo.length; i++) {
		if (f.rb_bem_uso_consumo[i].checked) blnFlag=true;
		}
	if (!blnFlag) {
		alert('Informe se é "Bem de Uso/Consumo"');
		return;
		}

	blnFlag=false;
	for (i=0; i < f.rb_instalador_instala.length; i++) {
		if (f.rb_instalador_instala[i].checked) blnFlag=true;
		}
	if (!blnFlag) {
		alert('Preencha o campo "Instalador Instala"');
		return;
		}

	blnFlag=false;
	for (i=0; i < f.rb_garantia_indicador.length; i++) {
		if (f.rb_garantia_indicador[i].checked) blnFlag=true;
		}
	if (!blnFlag) {
		alert('Preencha o campo "Garantia Indicador"');
		return;
		}
		
//  Há mensagens de alerta para os produtos?
//  A primeira posição do vetor é vazia, apenas p/ garantir que o vetor existe mesmo quando não há mensagens
	for (i=1; i < Pd.length; i++) {
		if (trim(Pd[i].mensagem)!="") {
			strProduto="Produto: " + trim(Pd[i].produto) + " - " + trim(Pd[i].descricao);
			strLinha=new Array(strProduto.length).join("=");
			strMsgAlerta=strLinha + "\n" + strProduto + "\n" + strLinha + "\n\n" + trim(Pd[i].mensagem) + "\n";
			if (!confirm(strMsgAlerta)) return;
			}
		}

	strMsgErro="";
	for (i=0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value)!="") {
			if ($(f.c_preco_lista[i]).hasClass('CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE')) {
				strMsgErro+="\n" + f.c_produto[i].value + " - " + f.c_descricao[i].value;
				}
			}
		}
	if (strMsgErro!="") {
		strMsgErro="A forma de pagamento " + KEY_ASPAS + f.c_custoFinancFornecParcelamentoDescricao.value.toLowerCase() + KEY_ASPAS + " não está disponível para o(s) produto(s):"+strMsgErro;
		alert(strMsgErro);
		return;
		}

	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}
</script>

<%
	Response.Write strScriptMsgAlerta
%>




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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
#rb_etg_imediata, #rb_bem_uso_consumo, #rb_instalador_instala {
	margin: 0pt 2pt 1pt 3pt;
	vertical-align: top;
	}
#rb_forma_pagto {
	margin: 0pt 2pt 1pt 10pt;
	vertical-align: top;
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
</style>

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
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
<!-- ************************************************************* -->
<!-- **********  PÁGINA PARA EDITAR ITENS DO ORÇAMENTO  ********** -->
<!-- ************************************************************* -->
<body onload="if (!(erro_produto_indisponivel&&bloquear_cadastramento_quando_produto_indiponivel)) {processaFormaPagtoDefault();restaura_cor_desconto();fORC.c_obs1.focus();}">
<center>

<form id="fORC" name="fORC" method="post" action="OrcamentoNovoConfirma.asp">
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=cliente_selecionado%>'>
<% if erro_produto_indisponivel then s="S" else s="" %>
<input type="hidden" name="opcao_venda_sem_estoque" id="opcao_venda_sem_estoque" value='<%=s%>'>
<input type="hidden" name="midia" id="midia" value='<%=midia%>'>
<input type="hidden" name="vendedor" id="vendedor" value='<%=vendedor%>'>
<input type="hidden" name="rb_end_entrega" id="rb_end_entrega" value='<%=rb_end_entrega%>'>
<input type="hidden" name="EndEtg_endereco" id="EndEtg_endereco" value="<%=EndEtg_endereco%>">
<input type="hidden" name="EndEtg_endereco_numero" id="EndEtg_endereco_numero" value="<%=EndEtg_endereco_numero%>">
<input type="hidden" name="EndEtg_endereco_complemento" id="EndEtg_endereco_complemento" value="<%=EndEtg_endereco_complemento%>">
<input type="hidden" name="EndEtg_bairro" id="EndEtg_bairro" value="<%=EndEtg_bairro%>">
<input type="hidden" name="EndEtg_cidade" id="EndEtg_cidade" value="<%=EndEtg_cidade%>">
<input type="hidden" name="EndEtg_uf" id="EndEtg_uf" value="<%=EndEtg_uf%>">
<input type="hidden" name="EndEtg_cep" id="EndEtg_cep" value="<%=EndEtg_cep%>">
<input type="hidden" name="c_PercVlPedidoLimiteRA" id="c_PercVlPedidoLimiteRA" value='<%=strPercVlPedidoLimiteRA%>'>
<input type="hidden" name="c_permite_RA_status" id="c_permite_RA_status" value='<%=r_orcamentista_e_indicador.permite_RA_status%>' />
<input type="hidden" name="c_loja" id="c_loja" value='<%=loja%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamento" id="c_custoFinancFornecTipoParcelamento" value='<%=c_custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelas" id="c_custoFinancFornecQtdeParcelas" value='<%=c_custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamentoUltConsulta" id="c_custoFinancFornecTipoParcelamentoUltConsulta" value='<%=c_custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelasUltConsulta" id="c_custoFinancFornecQtdeParcelasUltConsulta" value='<%=c_custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecParcelamentoDescricao" id="c_custoFinancFornecParcelamentoDescricao" value=''>
<input type="hidden" name="EndEtg_obs" id="EndEtg_obs" value='<%=EndEtg_obs%>'>


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


<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>


<!--  I D E N T I F I C A Ç Ã O   D O   O R Ç A M E N T O -->
<br />
<table width="741" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pré-Pedido Novo</span></td>
</tr>
</table>
<br>


<% if erro_produto_indisponivel then %>
<!--  RELAÇÃO DE PRODUTOS SEM PRESENÇA NO ESTOQUE -->
<table class="Qx" cellspacing="0">
	<tr><td class="MB ALERTA" colspan="6" align="center"><span class="ALERTA" style="font-size:9pt;">PRODUTOS SEM PRESENÇA NO ESTOQUE</span></td></tr>
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left"><span class="PLTe">Fabr</span></td>
	<td class="MDB" align="left"><span class="PLTe">Produto</span></td>
	<td class="MDB" align="left"><span class="PLTe">Descrição</span></td>
	<td class="MDB" align="right"><span class="PLTd">Solicitado</span></td>
	<td class="MDB" align="right"><span class="PLTd">Disponível</span></td>
	<td class="MDB" align="right"><span class="PLTd">Faltam</span></td>
	</tr>

<%	for i=Lbound(v_item) to Ubound(v_item)
		if Trim(v_item(i).produto) <> "" then
			with v_item(i)
				if .qtde > .qtde_estoque_total_disponivel then
%>
			<tr>
			<td class="MDBE" align="left"><input name="c_spe_fabricante" id="c_spe_fabricante" class="PLLe" style="width:26px;"
				value='<%=.fabricante%>' readonly tabindex=-1></td>
			<td class="MDB" align="left"><input name="c_spe_produto" id="c_spe_produto" class="PLLe" style="width:55px;"
				value='<%=.produto%>' readonly tabindex=-1></td>
			<td class="MDB" align="left">
				<span class="PLLe" style="width:333px;"><%=produto_formata_descricao_em_html(.descricao_html)%></span>
				<input type="hidden" name="c_spe_descricao" id="c_spe_descricao" value='<%=.descricao%>'>
			</td>
			<td class="MDB" align="right"><input name="c_spe_qtde_solicitada" id="c_spe_qtde_solicitada" class="PLLd" style="width:70px;"
				value='<%=Cstr(.qtde)%>' readonly tabindex=-1></td>
			<td class="MDB" align="right"><input name="c_spe_qtde_estoque" id="c_spe_qtde_estoque" class="PLLd" style="width:70px;"
				value='<%=Cstr(.qtde_estoque_total_disponivel)%>' readonly tabindex=-1></td>
			<td class="MDB" align="right"><input name="c_spe_saldo" id="c_spe_saldo" class="PLLd" style="width:70px;color:red;"
				value='<%=Cstr(Abs(.qtde_estoque_total_disponivel - .qtde))%>' readonly tabindex=-1></td>
			</tr>
		<%			end if
				end with
			end if
		next		%>
</table>
<% end if %>

<% if Not (erro_produto_indisponivel And bloquear_cadastramento_quando_produto_indiponivel) then %>
<br>
<br>
<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0">
	<tr bgcolor="#FFFFFF">
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Produto</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Descrição</span></td>
	<td class="MB tdProdObs" align="left" valign="bottom"><span class="PLTe">Observações</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Qtde</span></td>
	<% if r_orcamentista_e_indicador.permite_RA_status = 1 then %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Preço</span></td>
	<% end if %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Lista</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Desc%</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Venda</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Total</span></td>
	</tr>

<% m_TotalDestePedido=0
   m_TotalDestePedidoComRA=0
   n = Lbound(v_item)-1
   for i=1 to MAX_ITENS 
	 s_readonly = "readonly tabindex=-1"
	 n = n+1
	 if n <= Ubound(v_item) then
		with v_item(n)
			s_fabricante=.fabricante
			s_produto=.produto
			s_descricao=.descricao
			s_descricao_html=produto_formata_descricao_em_html(.descricao_html)
			s_qtde=.qtde
			s_preco_lista=formata_moeda(.preco_lista)
			m_TotalItem=.qtde * .preco_lista
		'	INICIALMENTE, O PRECO_NF É O MESMO VALOR DO PRECO_LISTA, FICANDO DIFERENTE APENAS SE FOR EDITADO
			m_TotalItemComRA=.qtde * .preco_lista
			s_vl_TotalItem=formata_moeda(m_TotalItem)
			m_TotalDestePedido=m_TotalDestePedido + m_TotalItem
			m_TotalDestePedidoComRA=m_TotalDestePedidoComRA + m_TotalItemComRA
			s_readonly = ""
			end with
	 else
		s_fabricante=""
		s_produto=""
		s_descricao=""
		s_descricao_html=""
		s_qtde=""
		s_preco_lista=""
		s_vl_TotalItem=""
		end if
%>
	<tr>
	<td class="MDBE" align="left">
		<input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:26px;"
			value='<%=s_fabricante%>' readonly tabindex=-1 />
	</td>
	<td class="MDB" align="left">
		<input name="c_produto" id="c_produto" class="PLLe" style="width:55px;"
			value='<%=s_produto%>' readonly tabindex=-1 />
	</td>
	<td class="MDB" align="left" style="width:277px;">
		<span class="PLLe"><%=s_descricao_html%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value='<%=s_descricao%>' />
	</td>
	<td class="MDB tdProdObs" align="left">
		<% if blnLojaHabilitadaProdCompostoECommerce then s_campo_focus="c_desc" else s_campo_focus="c_vl_unitario"%>
		<input name="c_obs" id="c_obs" maxlength="10" class="PLLe" style="width:80px;"
			onkeypress="if (digitou_enter(true)) <%if r_orcamentista_e_indicador.permite_RA_status = 1 then Response.Write "fORC.c_vl_NF" else Response.Write "fORC." & s_campo_focus%>[<%=Cstr(i-1)%>].focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
			value='' <%=s_readonly%>
			/>
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde" class="PLLd" style="width:27px;"
			value='<%=s_qtde%>' readonly tabindex=-1
			/>
	</td>
	<% if r_orcamentista_e_indicador.permite_RA_status = 1 then %>
	<td class="MDB" align="right">
		<% if blnLojaHabilitadaProdCompostoECommerce then s_campo_focus="c_desc" else s_campo_focus="c_vl_unitario"%>
		<input name="c_vl_NF" id="c_vl_NF" class="PLLd" style="width:62px;"
			onkeypress="if (digitou_enter(true)) fORC.<%=s_campo_focus%>[<%=Cstr(i-1)%>].focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); recalcula_RA(); recalcula_parcelas();"
			value='<%=s_preco_lista%>' <%=s_readonly%>
			/>
	</td>
	<% end if %>
	<td class="MDB" align="right">
		<input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px;"
			value='<%=s_preco_lista%>' readonly tabindex=-1
			/>
	</td>
	<td class="MDB" align="right">
		<input name="c_desc" id="c_desc" class="PLLd" style="width:36px;" value=""
		<% if blnLojaHabilitadaProdCompostoECommerce then %>
			<%=s_readonly%>
			onkeypress="if (digitou_enter(true)){fORC.c_vl_unitario[<%=Cstr(i-1)%>].focus();} filtra_percentual();"
			onblur="this.value=formata_perc_desconto(this.value); calcula_desconto(<%=Cstr(i-1)%>); trata_edicao_RA(<%=Cstr(i-1)%>); recalcula_total_linha(<%=Cstr(i)%>); recalcula_RA();"
		<% else %>
			readonly tabindex=-1
		<% end if %>
			/>
	</td>
	<td class="MDB" align="right">
		<input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px;"
			onkeypress="if (digitou_enter(true)) {if ((<%=Cstr(i)%>==fORC.c_vl_unitario.length)||(trim(fORC.c_produto[<%=Cstr(i)%>].value)=='')) fORC.c_obs1.focus(); else fORC.c_obs[<%=Cstr(i)%>].focus();} filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); recalcula_total(<%=Cstr(i)%>); recalcula_RA(); recalcula_parcelas();"
			value='<%=s_preco_lista%>' <%=s_readonly%>
			/>
	</td>
	<td class="MDB" align="right">
		<input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px;" 
			value='<%=s_vl_TotalItem%>' readonly tabindex=-1
			/>
	</td>
	</tr>
<% next %>
	<tr>
	<% if r_orcamentista_e_indicador.permite_RA_status = 1 then %>
	<td colspan="<%=Cstr(4-qtdeHiddenColumnTabProd)%>" align="left">
		<table cellspacing=0 cellpadding=0 width='100%' style="margin-top:4px;">
			<tr>
			<td width="60%" align="left">&nbsp;</td>
			<td align="right">
			<table cellspacing="0" cellpadding="0">
				<tr>
					<td class="MTBE" align="left"><span class="PLTe">&nbsp;RA</span></td>
					<td class="MTBD" align="right"><input name="c_total_RA" id="c_total_RA" class="PLLd" style="width:70px;color:blue;" 
						value='' readonly tabindex=-1></td>
				</tr>
			</table>
			</td>
			</tr>
		</table>
	</td>
	<td class="MD" align="left">&nbsp;</td>
	<td class="MDB" align="right">
		<%s_TotalDestePedidoComRA=formata_moeda(m_TotalDestePedidoComRA)%>
		<input name="c_total_NF" id="c_total_NF" class="PLLd" style="width:70px;color:blue;" 
				value='<%=s_TotalDestePedidoComRA%>' readonly tabindex=-1>
	</td>
	<% else %>
	<td colspan="<%=Cstr(5-qtdeHiddenColumnTabProd)%>" align="left">&nbsp;</td>
	<% end if %>
	<td colspan="3" class="MD" align="left">&nbsp;</td>
	<td class="MDB" align="right"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:blue;" 
		value='<%=formata_moeda(m_TotalDestePedido)%>' readonly tabindex=-1></td>
	</tr>
</table>

<br>
<table class="Q" cellspacing="0">
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observações</p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:641px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS1);" onblur="this.value=trim(this.value);"
				></textarea>
		</td>
	</tr>
	<tr>
		<td class="MD" align="left" nowrap><p class="Rf">Nº Nota Fiscal</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" maxlength="10" style="width:100px;margin-left:2pt;" onkeypress="filtra_nome_identificador();" onblur="this.value=trim(this.value);"
				value='' readonly tabindex=-1>
		</td>
		<td class="MD" align="left" nowrap><p class="Rf">Entrega Imediata</p>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
				value="<%=COD_ETG_IMEDIATA_NAO%>"><span class="C" style="cursor:default" onclick="fORC.rb_etg_imediata[0].click();">Não</span>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
				value="<%=COD_ETG_IMEDIATA_SIM%>"><span class="C" style="cursor:default" onclick="fORC.rb_etg_imediata[1].click();">Sim</span>
		</td>
		<td align="left" nowrap><p class="Rf">Bem de Uso/Consumo</p>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				value="<%=COD_ST_BEM_USO_CONSUMO_NAO%>"><span class="C" style="cursor:default" onclick="fORC.rb_bem_uso_consumo[0].click();">Não</span>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				value="<%=COD_ST_BEM_USO_CONSUMO_SIM%>"><span class="C" style="cursor:default" onclick="fORC.rb_bem_uso_consumo[1].click();">Sim</span>
		</td>
		<td class="ME" align="left" nowrap><p class="Rf">Instalador Instala</p>
			<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
				value="<%=COD_INSTALADOR_INSTALA_NAO%>"><span class="C" style="cursor:default" onclick="fORC.rb_instalador_instala[0].click();">Não</span>
			<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
				value="<%=COD_INSTALADOR_INSTALA_SIM%>"><span class="C" style="cursor:default" onclick="fORC.rb_instalador_instala[1].click();">Sim</span>
		</td>
		<td class="ME tdGarInd" align="left" nowrap><p class="Rf">Garantia Indicador</p>
			<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" class="rbGarIndNao"
				value="<%=COD_GARANTIA_INDICADOR_STATUS__NAO%>"><span class="C" style="cursor:default" onclick="fORC.rb_garantia_indicador[0].click();">Não</span>
			<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador"
				value="<%=COD_GARANTIA_INDICADOR_STATUS__SIM%>"><span class="C" style="cursor:default" onclick="fORC.rb_garantia_indicador[1].click();">Sim</span>
		</td>
	</tr>
    <tr>
		<td class="MC" align="left" colspan="5">
			<p class="Rf">Previsão de Entrega</p>
			<input type="text" class="PLLc" name="c_data_previsao_entrega" id="c_data_previsao_entrega" maxlength="10" style="width:90px;" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="filtra_data();" />
		</td>
    </tr>
</table>
	
<!--  NOVA VERSÃO DA FORMA DE PAGAMENTO  -->
<br>
<table class="Q" style="width:649px;" cellspacing="0">
  <tr>
	<td align="left">
	  <p class="Rf">Forma de Pagamento</p>
	</td>
  </tr>  
  <tr>
	<td align="left">
	  <table width="100%" cellspacing="0" cellpadding="4" border="0">
		<!--  À VISTA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left">
				  <% intIdx = 0 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_A_VISTA%>"
						<%if c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA then Response.Write " checked"%>
						onclick="recalculaCustoFinanceiroPrecoLista();"
						><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">À Vista</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <select id="op_av_forma_pagto" name="op_av_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<% =forma_pagto_liberada_av_monta_itens_select(Null, r_orcamentista_e_indicador.apelido, r_cliente.tipo) %>
				  </select>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELA ÚNICA  -->
		<tr class="TR_FP_PU">
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELA_UNICA%>" 
						<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) And _
							 (converte_numero(c_custoFinancFornecQtdeParcelas)=1) then Response.Write " checked"%>
						onclick="recalculaCustoFinanceiroPrecoLista();pu_atualiza_valor();"
						><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcela Única</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <select id="op_pu_forma_pagto" name="op_pu_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<% =forma_pagto_liberada_da_parcela_unica_monta_itens_select(Null, r_orcamentista_e_indicador.apelido, r_cliente.tipo) %>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pu_valor" id="c_pu_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pu_vencto_apos.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);" value=''
				  ><span style="width:10px;">&nbsp;</span
				  ><span class="C">vencendo após</span
				  ><input name="c_pu_vencto_apos" id="c_pu_vencto_apos" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_forma_pagto.focus(); filtra_numerico();" value=''
				  ><span class="C">dias</span>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO NO CARTÃO (INTERNET)  -->
		<% if is_restricao_ativa_forma_pagto(r_orcamentista_e_indicador.apelido, ID_FORMA_PAGTO_CARTAO, r_cliente.tipo) then%>
		<tr style="display:none;">
		<% else %>
		<tr>
		<% end if %>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELADO_CARTAO%>"
						onclick="recalculaCustoFinanceiroPrecoLista();"
						><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado no Cartão (internet)</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <input name="c_pc_qtde" id="c_pc_qtde" class="Cc" maxlength="2" style="width:30px;"  onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pc_valor.focus(); filtra_numerico();" onblur="pc_calcula_valor_parcela();recalculaCustoFinanceiroPrecoLista();" value=''>
				</td>
				<td align="left"><span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span></td>
				<td align="left">
				  <input name="c_pc_valor" id="c_pc_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);" value=''>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO NO CARTÃO (MAQUINETA)  -->
		<% if is_restricao_ativa_forma_pagto(r_orcamentista_e_indicador.apelido, ID_FORMA_PAGTO_CARTAO_MAQUINETA, r_cliente.tipo) then%>
		<tr style="display:none;">
		<% else %>
		<tr>
		<% end if %>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA%>"
						onclick="recalculaCustoFinanceiroPrecoLista();"
						><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado no Cartão (maquineta)</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <input name="c_pc_maquineta_qtde" id="c_pc_maquineta_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pc_maquineta_valor.focus(); filtra_numerico();" onblur="pc_maquineta_calcula_valor_parcela();recalculaCustoFinanceiroPrecoLista();" value=''>
				</td>
				<td align="left"><span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span></td>
				<td align="left">
				  <input name="c_pc_maquineta_valor" id="c_pc_maquineta_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);" value=''>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO COM ENTRADA  -->
		<tr>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td colspan="3" align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA%>" 
						<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) then Response.Write " checked"%>
						onclick="recalculaCustoFinanceiroPrecoLista();pce_preenche_sugestao_intervalo();"
						><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado com Entrada</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">Entrada&nbsp;</span></td>
				<td align="left">
				  <select id="op_pce_entrada_forma_pagto" name="op_pce_entrada_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<% =forma_pagto_liberada_da_entrada_monta_itens_select(Null, r_orcamentista_e_indicador.apelido, r_cliente.tipo) %>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pce_entrada_valor" id="c_pce_entrada_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.op_pce_prestacao_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);" value=''>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">Prestações&nbsp;</span></td>
				<td align="left">
				  <select id="op_pce_prestacao_forma_pagto" name="op_pce_prestacao_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<% =forma_pagto_liberada_da_prestacao_monta_itens_select(Null, r_orcamentista_e_indicador.apelido, r_cliente.tipo) %>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <input name="c_pce_prestacao_qtde" id="c_pce_prestacao_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onblur="recalculaCustoFinanceiroPrecoLista();pce_calcula_valor_parcela();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pce_prestacao_valor.focus(); filtra_numerico();" 
						value='<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) then Response.Write c_custoFinancFornecQtdeParcelas%>'
						>
				  <span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pce_prestacao_valor" id="c_pce_prestacao_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pce_prestacao_periodo.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);" value=''>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td colspan="2" align="left"><span class="C">Parcelas vencendo a cada</span
				><input name="c_pce_prestacao_periodo" id="c_pce_prestacao_periodo" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_forma_pagto.focus(); filtra_numerico();" 
						value='<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) then Response.Write "30"%>'
				><span class="C">dias</span
				><span style="width:10px;">&nbsp;</span
				><span class="notPrint"><input name="b_pce_SugereFormaPagto" id="b_pce_SugereFormaPagto" type="button" class="Button" onclick="pce_sugestao_forma_pagto();" value="sugestão automática" title="preenche o campo 'Forma de Pagamento' com uma sugestão de texto"></span
				></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO SEM ENTRADA  -->
		<tr class="TR_FP_PSE">
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td colspan="3" align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA%>" 
						<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) And _
							 (converte_numero(c_custoFinancFornecQtdeParcelas)>1) then Response.Write " checked"%>
						onclick="pse_preenche_sugestao_intervalo();recalculaCustoFinanceiroPrecoLista();"
						><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado sem Entrada</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">1ª Prestação&nbsp;</span></td>
				<td align="left">
				  <select id="op_pse_prim_prest_forma_pagto" name="op_pse_prim_prest_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<% =forma_pagto_liberada_da_prestacao_monta_itens_select(Null, r_orcamentista_e_indicador.apelido, r_cliente.tipo) %>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pse_prim_prest_valor" id="c_pse_prim_prest_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pse_prim_prest_apos.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); pse_calcula_valor_parcela();" value=''
				  ><span style="width:10px;">&nbsp;</span
				  ><span class="C">vencendo após</span
				  ><input name="c_pse_prim_prest_apos" id="c_pse_prim_prest_apos" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.op_pse_demais_prest_forma_pagto.focus(); filtra_numerico();" value=''
				  ><span class="C">dias</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">Demais Prestações&nbsp;</span></td>
				<td align="left">
				  <select id="op_pse_demais_prest_forma_pagto" name="op_pse_demais_prest_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<% =forma_pagto_liberada_da_prestacao_monta_itens_select(Null, r_orcamentista_e_indicador.apelido, r_cliente.tipo) %>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <input name="c_pse_demais_prest_qtde" id="c_pse_demais_prest_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onblur="pse_calcula_valor_parcela();recalculaCustoFinanceiroPrecoLista();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pse_demais_prest_valor.focus(); filtra_numerico();" 
						value='<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) And (converte_numero(c_custoFinancFornecQtdeParcelas)>1) then Response.Write Cstr(converte_numero(c_custoFinancFornecQtdeParcelas)-1)%>'
						>
				  <span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pse_demais_prest_valor" id="c_pse_demais_prest_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pse_demais_prest_periodo.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); " value=''>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td colspan="2" align="left"><span class="C">Parcelas vencendo a cada</span
				><input name="c_pse_demais_prest_periodo" id="c_pse_demais_prest_periodo" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_forma_pagto.focus(); filtra_numerico();" 
						value='<%if (c_custoFinancFornecTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) And (converte_numero(c_custoFinancFornecQtdeParcelas)>1) then Response.Write "30"%>'
				><span class="C">dias</span
				><span style="width:10px;">&nbsp;</span
				><span class="notPrint"><input name="b_pse_SugereFormaPagto" id="b_pse_SugereFormaPagto" type="BUTTON" class="Button" onclick="pse_sugestao_forma_pagto();" value="sugestão automática" title="preenche o campo 'Forma de Pagamento' com uma sugestão de texto"></span
				></td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
  <tr>
	<td class="MC" align="left">
	  <p class="Rf">Descrição da Forma de Pagamento</p>
		<textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
			style="width:641px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_FORMA_PAGTO);" onblur="this.value=trim(this.value);"
			></textarea>
	</td>
  </tr>
</table>
<% end if 'if (Not (erro_produto_indisponivel And bloquear_cadastramento_quando_produto_indiponivel)) %>


<!-- ************   SEPARADOR   ************ -->
<table width="741" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<% if erro_produto_indisponivel And bloquear_cadastramento_quando_produto_indiponivel then %>
	<tr>
		<td align="center"><a name="bVOLTAR" id="A1" href="javascript:history.back()" title="volta para página anterior">
			<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	</tr>
<% else %>
	<tr>
		<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
			<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
		<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
			<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fORCConfirma(fORC)" title="confirma o novo pré-pedido">
			<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
		</td>
	</tr>
<% end if %>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>