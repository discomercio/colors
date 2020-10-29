<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  PedidoNovoECConsiste
'     =============================================================
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

	class cl_MAP_ITEM
		dim sku
		dim qty_ordered
		dim price
		dim name
		end class

	dim s, usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, tMAP_XML, tMAP_END_COB, tMAP_END_ETG, tMAP_ITEM, tPROD, tPCI, tPED, tPEDITM, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tMAP_XML, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tMAP_END_COB, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tMAP_END_ETG, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tMAP_ITEM, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tPROD, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tPCI, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tPED, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tPEDITM, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim alerta
	alerta = ""
	
	dim blnFlagCadSemiAutoPedMagentoCadAutoClienteNovo
	blnFlagCadSemiAutoPedMagentoCadAutoClienteNovo = isActivatedFlagCadSemiAutoPedMagentoCadAutoClienteNovo

	dim rowspan_placeholder
	rowspan_placeholder = " rowspan=" & chr(34) & "XXX" & chr(34)

	dim s_id_cliente, s_nome_cliente, s_end_cobranca, s_end_entrega, s_value, s_cor, iv, s_row, qtde_rowspan, operacao_selecionada, id_magento_api_pedido_xml
	s_id_cliente = ""
	s_nome_cliente = ""
	s_end_entrega = ""
	operacao_selecionada = OP_INCLUI
	id_magento_api_pedido_xml = ""

	dim v_map_item
	redim v_map_item(0)
	set v_map_item(UBound(v_map_item)) = new cl_MAP_ITEM
	v_map_item(UBound(v_map_item)).sku = ""

'	OBTÉM DADOS DO FORMULÁRIO
	dim c_numero_magento, operationControlTicket, sessionToken
	c_numero_magento = Trim(Request("c_numero_magento"))
	operationControlTicket = Trim(Request("operationControlTicket"))
	sessionToken = Trim(Request("sessionToken"))

	s = "SELECT " & _
			"*" & _
		" FROM t_MAGENTO_API_PEDIDO_XML" & _
		" WHERE" & _
			" (pedido_magento = '" & c_numero_magento & "')" & _
			" AND (operationControlTicket = '" & operationControlTicket & "')"
	if tMAP_XML.State <> 0 then tMAP_XML.Close
	tMAP_XML.open s, cn
	if tMAP_XML.Eof then
		alerta = "Falha ao tentar localizar no banco de dados o pedido Magento nº " & c_numero_magento & " (operationControlTicket = " & operationControlTicket & ")"
	else
		id_magento_api_pedido_xml = Trim("" & tMAP_XML("id"))
		s_nome_cliente = UCase(ec_dados_formata_nome(tMAP_XML("customer_firstname"), tMAP_XML("customer_middlename"), tMAP_XML("customer_lastname"), Null))
		end if

	if alerta = "" then
		s = "SELECT " & _
				"*" & _
			" FROM t_MAGENTO_API_PEDIDO_XML_DECODE_ENDERECO" & _
			" WHERE" & _
				" (id_magento_api_pedido_xml = " & tMAP_XML("id") & ")" & _
				" AND (tipo_endereco = 'COB')"
		if tMAP_END_COB.State <> 0 then tMAP_END_COB.Close
		tMAP_END_COB.open s, cn
		if tMAP_END_COB.Eof then
			alerta = "Falha ao tentar localizar no banco de dados o registro do endereço de cobrança do pedido Magento nº " & c_numero_magento & " (operationControlTicket = " & operationControlTicket & ")"
		else
			s_end_cobranca = formata_endereco(Trim("" & tMAP_END_COB("endereco")), Trim("" & tMAP_END_COB("endereco_numero")), Trim("" & tMAP_END_COB("endereco_complemento")), Trim("" & tMAP_END_COB("bairro")), Trim("" & tMAP_END_COB("cidade")), Trim("" & tMAP_END_COB("uf")), Trim("" & tMAP_END_COB("cep")))
			end if
		end if

	if alerta = "" then
		s = "SELECT " & _
				"*" & _
			" FROM t_MAGENTO_API_PEDIDO_XML_DECODE_ENDERECO" & _
			" WHERE" & _
				" (id_magento_api_pedido_xml = " & tMAP_XML("id") & ")" & _
				" AND (tipo_endereco = 'ETG')"
		if tMAP_END_ETG.State <> 0 then tMAP_END_ETG.Close
		tMAP_END_ETG.open s, cn
		if tMAP_END_ETG.Eof then
			alerta = "Falha ao tentar localizar no banco de dados o registro do endereço de entrega do pedido Magento nº " & c_numero_magento & " (operationControlTicket = " & operationControlTicket & ")"
		else
			s_end_entrega = formata_endereco(Trim("" & tMAP_END_ETG("endereco")), Trim("" & tMAP_END_ETG("endereco_numero")), Trim("" & tMAP_END_ETG("endereco_complemento")), Trim("" & tMAP_END_ETG("bairro")), Trim("" & tMAP_END_ETG("cidade")), Trim("" & tMAP_END_ETG("uf")), Trim("" & tMAP_END_ETG("cep")))
			end if
		end if

	if alerta = "" then
		s = "SELECT " & _
				"*" & _
			" FROM t_MAGENTO_API_PEDIDO_XML_DECODE_ITEM" & _
			" WHERE" & _
				" (id_magento_api_pedido_xml = " & tMAP_XML("id") & ")" & _
				" AND (product_type <> 'configurable')" & _
			" ORDER BY" & _
				" id"
		if tMAP_ITEM.State <> 0 then tMAP_ITEM.Close
		tMAP_ITEM.open s, cn
		if tMAP_ITEM.Eof then
			alerta = "Falha ao tentar localizar no banco de dados os itens do pedido Magento nº " & c_numero_magento & " (operationControlTicket = " & operationControlTicket & ")"
		else
			do while Not tMAP_ITEM.Eof
				if Trim("" & v_map_item(UBound(v_map_item)).sku) <> "" then
					redim preserve v_map_item(UBound(v_map_item)+1)
					set v_map_item(UBound(v_map_item)) = new cl_MAP_ITEM
					end if

				v_map_item(UBound(v_map_item)).sku = Trim("" & tMAP_ITEM("sku"))
				v_map_item(UBound(v_map_item)).qty_ordered = CLng(tMAP_ITEM("qty_ordered"))
				v_map_item(UBound(v_map_item)).price = converte_numero(formata_moeda(tMAP_ITEM("price")))
				v_map_item(UBound(v_map_item)).name = Trim("" & tMAP_ITEM("name"))

				tMAP_ITEM.MoveNext
				loop
			end if
		end if

	if alerta = "" then
		s = "SELECT " & _
				"*" & _
			" FROM t_CLIENTE" & _
			" WHERE" & _
				" (cnpj_cpf = '" & retorna_so_digitos(Trim("" & tMAP_XML("cpfCnpjIdentificado"))) & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if Not rs.Eof then
			s_id_cliente = Trim("" & rs("id"))
			if s_id_cliente <> "" then operacao_selecionada = OP_CONSULTA
			end if
		end if
	
	if alerta = "" then
		' Verifica se todos os SKU's existem no sistema
		for iv=LBound(v_map_item) to UBound(v_map_item)
			if Trim("" & v_map_item(iv).sku) <> "" then
				s = "SELECT" & _
						" fabricante," & _
						" produto" & _
					" FROM t_PRODUTO" & _
					" WHERE" & _
						" (produto = '" & normaliza_codigo(Trim("" & v_map_item(iv).sku), TAM_MIN_PRODUTO) & "')" & _
						" AND (excluido_status = 0)"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					s = "SELECT" & _
							" produto_item" & _
						" FROM t_EC_PRODUTO_COMPOSTO_ITEM" & _
						" WHERE" & _
							" (produto_composto = '" & normaliza_codigo(Trim("" & v_map_item(iv).sku), TAM_MIN_PRODUTO) & "')" & _
							" AND (excluido_status = 0)"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "SKU " & Trim("" & v_map_item(iv).sku) & " não consta no sistema!!"
						end if
					end if
				end if
			next
		end if

	dim s_lista_sugerida_municipios_end_cob, s_lista_sugerida_municipios_end_etg
	dim blnErroIbgeEndCob, blnErroIbgeEndEtg, blnEnderecosIguais, blnFlagCadSemiAutoPedMagento_FluxoOtimizado, tipoCliente
	dim rb_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep
	dim EndEtg_obs, EndEtg_email, EndEtg_email_xml, EndEtg_nome, EndEtg_tipo_pessoa
	dim EndEtg_ddd_res, EndEtg_tel_res, EndEtg_ddd_com, EndEtg_tel_com, EndEtg_ramal_com, EndEtg_ddd_com_2, EndEtg_tel_com_2, EndEtg_ramal_com_2, EndEtg_ddd_cel, EndEtg_tel_cel
	dim EndEtg_cnpj_cpf, EndEtg_contribuinte_icms_status, EndEtg_produtor_rural_status, EndEtg_ie, EndEtg_rg
	dim rb_indicacao, c_indicador, rb_RA
	dim s_ddd, s_tel
	blnErroIbgeEndCob = False
	blnErroIbgeEndEtg = False
	blnEnderecosIguais = False
	blnFlagCadSemiAutoPedMagento_FluxoOtimizado = False
	tipoCliente = ""
	rb_end_entrega = "N"
	EndEtg_endereco = ""
	EndEtg_endereco_numero = ""
	EndEtg_endereco_complemento = ""
	EndEtg_bairro = ""
	EndEtg_cidade = ""
	EndEtg_uf = ""
	EndEtg_cep = ""
	EndEtg_obs = ""
	EndEtg_email = ""
	EndEtg_email_xml = ""
	EndEtg_nome = ""
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
	EndEtg_tipo_pessoa = ""
	EndEtg_cnpj_cpf = ""
	EndEtg_contribuinte_icms_status = ""
	EndEtg_produtor_rural_status = ""
	EndEtg_ie = ""
	EndEtg_rg = ""
	rb_indicacao = "N"
	c_indicador = ""
	rb_RA = "N"

	if alerta = "" then
		if Len(retorna_so_digitos(Trim("" & tMAP_XML("cpfCnpjIdentificado")))) = 11 then
			tipoCliente = ID_PF
		else
			tipoCliente = ID_PJ
			end if

		blnEnderecosIguais = isEnderecoMagentoIgual(Trim("" & tMAP_END_COB("endereco")), _
													Trim("" & tMAP_END_COB("endereco_numero")), _
													Trim("" & tMAP_END_COB("endereco_complemento")), _
													Trim("" & tMAP_END_COB("bairro")), _
													Trim("" & tMAP_END_COB("cidade")), _
													Trim("" & tMAP_END_COB("uf")), _
													Trim("" & tMAP_END_COB("cep")), _
													Trim("" & tMAP_END_ETG("endereco")), _
													Trim("" & tMAP_END_ETG("endereco_numero")), _
													Trim("" & tMAP_END_ETG("endereco_complemento")), _
													Trim("" & tMAP_END_ETG("bairro")), _
													Trim("" & tMAP_END_ETG("cidade")), _
													Trim("" & tMAP_END_ETG("uf")), _
													Trim("" & tMAP_END_ETG("cep")))

		'VERIFICA SE ESTÁ ATIVADO O FLAG DO PARÂMETRO DE CADASTRAMENTO AUTOMÁTICO DE CLIENTE NOVO NA WEB API PARA CADASTRAMENTO SEMI-AUTOMÁTICO DE PEDIDOS DO MAGENTO
		'CASO ESTEJA, O FLUXO NORMAL É DIRECIONAR AUTOMATICAMENTE ATÉ CHEGAR NA PÁGINA PedidoNovoConsiste.asp
		'ENTRETANTO, É FEITA A VERIFICAÇÃO DO CADASTRO DO MUNICÍPIO NO IBGE, POIS CASO EXISTA INCONSISTÊNCIA, O FLUXO É DIRECIONADO PARA O CADASTRO DE EDIÇÃO DO CLIENTE
		'LEMBRANDO QUE NOS PEDIDOS DE E-COMMERCE, NO CASO DE CLIENTES PF, SOMENTE O ENDEREÇO DE ENTREGA É UTILIZADO
		if blnFlagCadSemiAutoPedMagentoCadAutoClienteNovo then
			'É CLIENTE PJ?
			if tipoCliente = ID_PJ then
				if Not consiste_municipio_IBGE_ok(Trim("" & tMAP_END_COB("cidade")), Trim("" & tMAP_END_COB("uf")), s_lista_sugerida_municipios_end_cob, msg_erro) then
					blnErroIbgeEndCob = True
					end if
				end if

			if Not consiste_municipio_IBGE_ok(Trim("" & tMAP_END_ETG("cidade")), Trim("" & tMAP_END_ETG("uf")), s_lista_sugerida_municipios_end_etg, msg_erro) then
				 blnErroIbgeEndEtg = True
				 end if
			
			if tipoCliente = ID_PJ then
				'DEVIDO À AUSÊNCIA DE INFORMAÇÃO NO MAGENTO SOBRE O CLIENTE PJ SER CONTRIBUINTE DE ICMS E NÚMERO DE IE, O FLUXO É SEMPRE DIRECIONADO PARA O CADASTRO DE CLIENTES PARA REVISÃO MANUAL
				'*DESATIVADO* if (Not blnErroIbgeEndCob) And (Not blnErroIbgeEndEtg) then blnFlagCadSemiAutoPedMagento_FluxoOtimizado = True
			else
				'CLIENTES PF CUJO CADASTRO É CRIADO COM DADOS DO MAGENTO: ASSUME-SE QUE NUNCA SÃO PRODUTORES RURAIS DEVIDO AO NÚMERO EXTREMAMENTE BAIXO DE CASOS
				if (Not blnErroIbgeEndEtg) then blnFlagCadSemiAutoPedMagento_FluxoOtimizado = True
				end if

			if blnFlagCadSemiAutoPedMagento_FluxoOtimizado then
				if tipoCliente = ID_PF then
					rb_end_entrega = "N"
				else
					if Not blnEnderecosIguais then rb_end_entrega = "S"
					end if
				
				if rb_end_entrega = "S" then
					EndEtg_tipo_pessoa = tipoCliente
					EndEtg_endereco = Trim("" & tMAP_END_ETG("endereco"))
					EndEtg_endereco_numero = Trim("" & tMAP_END_ETG("endereco_numero"))
					EndEtg_endereco_complemento = Trim("" & tMAP_END_ETG("endereco_complemento"))
					EndEtg_bairro = Trim("" & tMAP_END_ETG("bairro"))
					EndEtg_cidade = Trim("" & tMAP_END_ETG("cidade"))
					EndEtg_uf = Trim("" & tMAP_END_ETG("uf"))
					EndEtg_cep = retorna_so_digitos(Trim("" & tMAP_END_ETG("cep")))
				'	CÓDIGO DA JUSTIFICATIVA DO ENDEREÇO DE ENTREGA (GRAVADO EM t_PEDIDO.EndEtg_cod_justificativa): "007" = Pedido Arclube
					EndEtg_obs = "007"
					EndEtg_email = Trim("" & tMAP_XML("customer_email"))
					EndEtg_email_xml = ""
					EndEtg_nome = UCase(ec_dados_formata_nome(tMAP_XML("firstname"), tMAP_XML("middlename"), tMAP_XML("lastname"), 60))
					if tipoCliente = ID_PF then
						call ec_dados_decodifica_telefone_formatado(tMAP_END_ETG("telephone"), s_ddd, s_tel)
						EndEtg_ddd_res = s_ddd
						EndEtg_tel_res = s_tel

						call ec_dados_decodifica_telefone_formatado(tMAP_END_ETG("celular"), s_ddd, s_tel)
						EndEtg_ddd_cel = s_ddd
						EndEtg_tel_cel = s_tel

						EndEtg_cnpj_cpf = retorna_so_digitos(Trim("" & tMAP_XML("cpfCnpjIdentificado")))
						EndEtg_contribuinte_icms_status = ""
						EndEtg_ie = ""
						EndEtg_produtor_rural_status = ""
						EndEtg_rg = Trim("" & tMAP_END_ETG("rg"))
					else
						call ec_dados_decodifica_telefone_formatado(tMAP_END_ETG("telephone"), s_ddd, s_tel)
						EndEtg_ddd_com = s_ddd
						EndEtg_tel_com = s_tel

						call ec_dados_decodifica_telefone_formatado(tMAP_END_ETG("celular"), s_ddd, s_tel)
						EndEtg_ddd_com_2 = s_ddd
						EndEtg_tel_com_2 = s_tel
						EndEtg_ramal_com_2 = ""

						EndEtg_cnpj_cpf = retorna_so_digitos(Trim("" & tMAP_END_ETG("cpfcnpj")))
						if (EndEtg_cnpj_cpf = "") Or (Not cnpj_ok(EndEtg_cnpj_cpf)) then EndEtg_cnpj_cpf = retorna_so_digitos(Trim("" & tMAP_XML("cpfCnpjIdentificado")))
						EndEtg_contribuinte_icms_status = ""
						EndEtg_ie = Trim("" & tMAP_END_ETG("ie"))
						end if
					end if 'if rb_end_entrega = "S"

				'Tem valor de frete?
				'Se houver valor de frete, o pedido será cadastrado automaticamente como tendo RA e o indicador será o 'FRETE'
				if converte_numero(tMAP_XML("shipping_amount")) > 0 then
					rb_indicacao = "S"
					c_indicador = "FRETE"
					rb_RA = "S"
					end if
				end if 'if blnFlagCadSemiAutoPedMagento_FluxoOtimizado
			end if 'if blnFlagCadSemiAutoPedMagentoCadAutoClienteNovo
		end if 'if alerta = ""
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



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fPNECConfirma(f) {
    f.action = "ClienteEdita.asp";
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}
function fPNEC2Confirma(f) {
	f.action = "PedidoNovo.asp";
	dCONFIRMA2.style.visibility = "hidden";
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
.TdCliLbl
{
	width:150px;
	text-align:right;
}
.TdCliCel
{
	width:540px;
	text-align:left;
}
.TdPedSku
{
	width:60px;
	text-align:center;
	vertical-align:middle;
}
.TdPedQty
{
	width:60px;
	text-align:right;
	vertical-align:middle;
}
.TdPedDescrMag
{
	width:200px;
	text-align:left;
	vertical-align:middle;
}
.TdPedCodSist
{
	width:60px;
	text-align:center;
	vertical-align:middle;
}
.SpnPedCodSist
{
	color:blue;
}
.TdPedQtde
{
	width:60px;
	text-align:right;
	vertical-align:middle;
}
.SpnPedQtde
{
	color:blue;
}
.TdPedDescrSist
{
	width:240px;
	text-align:left;
	vertical-align:middle;
}
.SpnPedDescrSist
{
	color:blue;
}
.TdPedAntPedLbl
{
	width:70px;
	text-align:left;
}
.TdPedAntPed
{
	width:70px;
	text-align:left;
	vertical-align:middle;
}
.SpnPedAntPed
{
	text-align:left;
}
.TdPedAntDataLbl
{
	width:70px;
	text-align:center;
}
.TdPedAntData
{
	width:70px;
	text-align:center;
	vertical-align:middle;
}
.SpnPedAntData
{
	text-align:center;
}
.TdPedAntStEntLbl
{
	width:120px;
	text-align:center;
}
.TdPedAntStEnt
{
	width:120px;
	text-align:center;
	vertical-align:middle;
}
.SpnPedAntStEnt
{
	text-align:center;
}
.TdPedAntProdLbl
{
	width:70px;
	text-align:center;
}
.TdPedAntProd
{
	width:70px;
	text-align:center;
	vertical-align:middle;
}
.SpnPedAntProd
{
	text-align:center;
}
.TdPedAntQtdeLbl
{
	width:50px;
	text-align:right;
}
.TdPedAntQtde
{
	width:50px;
	text-align:right;
	vertical-align:middle;
}
.SpnPedAntQtde
{
	text-align:right;
}
.TdPedAntDescLbl
{
	width:300px;
	text-align:left;
}
.TdPedAntDesc
{
	width:300px;
	text-align:left;
	vertical-align:middle;
}
.SpnPedAntDesc
{
	text-align:left;
}
.TdPedDadosLbl
{
	width:100px;
	text-align:right;
}
.TdPedDadosCel
{
	text-align:left;
}
.TdEndLbl
{
	width:150px;
	text-align:right;
}
.TdEndMsg
{
	width:540px;
	text-align:left;
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
<!-- *************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR DADOS DE CONFIRMAÇÃO  ********** -->
<!-- *************************************************************** -->
<body onload="focus();">
<center>

<form id="fPNEC" name="fPNEC" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value="<%=operacao_selecionada%>" />
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value="<%=s_id_cliente%>" />
<input type="hidden" name="cnpj_cpf_selecionado" id="cnpj_cpf_selecionado" value="<%=retorna_so_digitos(Trim("" & tMAP_XML("cpfCnpjIdentificado")))%>" />
<input type="hidden" name="c_numero_magento" id="c_numero_magento" value="<%=c_numero_magento%>" />
<input type="hidden" name="operationControlTicket" id="operationControlTicket" value="<%=operationControlTicket%>" />
<input type="hidden" name="sessionToken" id="sessionToken" value="<%=sessionToken%>" />
<input type="hidden" name="operacao_origem" id="operacao_origem" value="<%=OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO%>" />
<input type="hidden" name="id_magento_api_pedido_xml" id="id_magento_api_pedido_xml" value="<%=id_magento_api_pedido_xml%>" />
<% if blnFlagCadSemiAutoPedMagento_FluxoOtimizado then %>
<input type="hidden" name="c_FlagCadSemiAutoPedMagento_FluxoOtimizado" value="1" />
<input type="hidden" name="rb_end_entrega" id="rb_end_entrega" value="<%=rb_end_entrega%>" />
<input type="hidden" name="EndEtg_endereco" id="EndEtg_endereco" value="<%=EndEtg_endereco%>" />
<input type="hidden" name="EndEtg_endereco_numero" id="EndEtg_endereco_numero" value="<%=EndEtg_endereco_numero%>" />
<input type="hidden" name="EndEtg_endereco_complemento" id="EndEtg_endereco_complemento" value="<%=EndEtg_endereco_complemento%>" />
<input type="hidden" name="EndEtg_bairro" id="EndEtg_bairro" value="<%=EndEtg_bairro%>" />
<input type="hidden" name="EndEtg_cidade" id="EndEtg_cidade" value="<%=EndEtg_cidade%>" />
<input type="hidden" name="EndEtg_uf" id="EndEtg_uf" value="<%=EndEtg_uf%>" />
<input type="hidden" name="EndEtg_cep" id="EndEtg_cep" value="<%=EndEtg_cep%>" />
<input type="hidden" name="EndEtg_obs" id="EndEtg_obs" value="<%=EndEtg_obs%>" />
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
<input type="hidden" name="rb_indicacao" id="rb_indicacao" value="<%=rb_indicacao%>" />
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>" />
<input type="hidden" name="rb_RA" id="rb_RA" value="<%=rb_RA%>" />
<!-- LISTA DE PRODUTOS -->
<% for iv=LBound(v_map_item) to Ubound(v_map_item) 
		if Trim(v_map_item(iv).sku) <> "" then %>
<input type="hidden" name="c_fabricante" value="" />
<input type="hidden" name="c_produto" value="<%=v_map_item(iv).sku%>" />
<input type="hidden" name="c_qtde" value="<%=Cstr(v_map_item(iv).qty_ordered)%>" />
<%			end if
		next
%>
<!-- FORÇA A CRIAÇÃO DO ARRAY MESMO QUANDO HÁ SOMENTE 1 PRODUTO -->
<input type="hidden" name="c_fabricante" value="" />
<input type="hidden" name="c_produto" value="" />
<input type="hidden" name="c_qtde" value="" />
<% end if %>



<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="849" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Novo pedido do e-commerce (semi-automático)<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>

<!--  CLIENTE  -->
<table class="Qx" cellspacing="0">
	<tr style="background-color:azure;">
		<td colspan="2" class="MC MB ME MD" align="center"><span class="N">CLIENTE</span></td>
	</tr>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">CPF/CNPJ</span></td>
		<td class="MB MD TdCliCel"><span class="C"><%=cnpj_cpf_formata(Trim("" & tMAP_XML("cpfCnpjIdentificado")))%></span></td>
	</tr>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">Nome</span></td>
		<td class="MB MD TdCliCel"><span class="C"><%=s_nome_cliente%></span></td>
	</tr>
	<tr>
		<% if s_id_cliente <> "" then
				s_value = "SIM"
				s_cor = "green"
			else
				s_value = "NÃO"
				s_cor = "red"
				end if
		%>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">Cadastrado no Sistema</span></td>
		<td class="MB MD TdCliCel"><span class="C" style="color:<%=s_cor%>;"><%=s_value%></span></td>
	</tr>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">Endereço de Entrega</span></td>
		<td class="MB MD TdCliCel"><span class="C"><%=s_end_entrega%></span></td>
	</tr>
</table>

<!--  PULA LINHA  -->
<br /><br />

<!--  DADOS DO PEDIDO  -->
<table class="Qx" cellspacing="0">
	<tr style="background-color:azure;">
		<td colspan="6" class="MC MB ME MD" align="center"><span class="N">DADOS DO PEDIDO Nº <%=c_numero_magento%></span></td>
	</tr>
	<tr>
		<td colspan="6" class="MB ME MD">
			<table cellspacing="0" width="100%">
				<tr>
					<%
						s_value = UCase(Trim("" & tMAP_XML("status")))
						if (s_value = "CANCELED") Or (s_value = "CANCELLED") then s_cor="red" else s_cor="black"
					%>
					<td class="MB MD TdPedDadosLbl" align="right"><span class="PLTd">Status (Magento)</span></td>
					<td class="MB TdPedDadosCel"><span class="C" style="color:<%=s_cor%>;"><%=Trim("" & tMAP_XML("status_descricao"))%></span></td>
				</tr>
				<tr>
					<%
						s_value = UCase(Trim("" & tMAP_XML("state")))
						if (s_value = "CANCELED") Or (s_value = "CANCELLED") then s_cor="red" else s_cor="black"
					%>
					<td class="MD TdPedDadosLbl" align="right"><span class="PLTd">State (Magento)</span></td>
					<td class="TdPedDadosCel"><span class="C" style="color:<%=s_cor%>;"><%=Trim("" & tMAP_XML("state_descricao"))%></span></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="MB ME MD TdPedSku"><span class="PLTc">SKU</span></td>
		<td class="MB MD TdPedQty"><span class="PLTd">Qtde</span></td>
		<td class="MB MD TdPedDescrMag"><span class="PLTe">Descrição (Magento)</span></td>
		<td class="MB MD TdPedCodSist"><span class="PLTc">Cód Sist</span></td>
		<td class="MB MD TdPedQtde"><span class="PLTd">Qtde</span></td>
		<td class="MB MD TdPedDescrSist"><span class="PLTe">Descrição (Sistema)</span></td>
	</tr>
<%
	for iv=LBound(v_map_item) to UBound(v_map_item)
		if Trim("" & v_map_item(iv).sku) <> "" then
			qtde_rowspan = 0
			s_row = "	<tr>" & chr(13) & _
					"		<td" & rowspan_placeholder & " class=""MB ME MD TdPedSku""><span class=""C SpnPedSku"">" & v_map_item(iv).sku & "</span></td>" & chr(13) & _
					"		<td" & rowspan_placeholder & " class=""MB MD TdPedQty""><span class=""C SpnPedQty"">" & v_map_item(iv).qty_ordered & "</span></td>" & chr(13) & _
					"		<td" & rowspan_placeholder & " class=""MB MD TdPedDescrMag""><span class=""C SpnPedDescrMag"">" & v_map_item(iv).name & "</span></td>" & chr(13)

			s = "SELECT" & _
					" tPCI.qtde," & _
					" tP.fabricante," & _
					" tP.produto," & _
					" tP.descricao," & _
					" tP.descricao_html" & _
				" FROM t_EC_PRODUTO_COMPOSTO_ITEM tPCI" & _
					" INNER JOIN t_PRODUTO tP ON (tPCI.fabricante_item = tP.fabricante) AND (tPCI.produto_item = tP.produto)" & _
				" WHERE" & _
					" (tPCI.produto_composto = '" & normaliza_codigo(v_map_item(iv).sku, TAM_MIN_PRODUTO) & "')" & _
					" AND (tPCI.excluido_status = 0)" & _
				" ORDER BY" & _
					" sequencia"
			if tPCI.State <> 0 then tPCI.Close
			tPCI.open s, cn
			if Not tPCI.Eof then
				do while Not tPCI.Eof
					qtde_rowspan = qtde_rowspan + 1
					if qtde_rowspan > 1 then s_row = s_row & _
												"	</tr>" & chr(13) & _
												"	<tr>" & chr(13)
					s_row = s_row & _
							"		<td class=""MB MD TdPedCodSist""><span class=""C SpnPedCodSist"">" & Trim("" & tPCI("produto")) & "</span></td>" & chr(13) & _
							"		<td class=""MB MD TdPedQtde""><span class=""C SpnPedQtde"">" & CStr(v_map_item(iv).qty_ordered * tPCI("qtde")) & "</span></td>" & chr(13) & _
							"		<td class=""MB MD TdPedDescrSist""><span class=""C SpnPedDescrSist"">" & Trim("" & tPCI("descricao")) & "</span></td>" & chr(13)
					tPCI.MoveNext
					loop
			else
				s = "SELECT " & _
						"*" & _
					" FROM t_PRODUTO" & _
					" WHERE" & _
						" (produto = '" & normaliza_codigo(v_map_item(iv).sku, TAM_MIN_PRODUTO) & "')"
				if tPROD.State <> 0 then tPROD.Close
				tPROD.open s, cn
				if Not tPROD.Eof then
					qtde_rowspan = qtde_rowspan + 1
					s_row = s_row & _
							"		<td class=""MB MD TdPedCodSist""><span class=""C SpnPedCodSist"">" & Trim("" & tPROD("produto")) & "</span></td>" & chr(13) & _
							"		<td class=""MB MD TdPedQtde""><span class=""C SpnPedQtde"">" & CStr(v_map_item(iv).qty_ordered) & "</span></td>" & chr(13) & _
							"		<td class=""MB MD TdPedDescrSist""><span class=""C SpnPedDescrSist"">" & Trim("" & tPROD("descricao")) & "</span></td>" & chr(13)
					end if
				end if

			s_row = s_row & _
					"	</tr>" & chr(13)

			if qtde_rowspan <= 1 then
				s_row = Replace(s_row, rowspan_placeholder, "")
			else
				s_row = Replace(s_row, rowspan_placeholder, " rowspan=" & chr(34) & Cstr(qtde_rowspan) & chr(34))
				end if

			Response.Write s_row
			end if
		next
%>
</table>

<!--  PULA LINHA  -->
<br /><br />

<!--  PEDIDOS ANTERIORES  -->
<%
	s = "SELECT" & _
			" tP.pedido," & _
			" tP.data," & _
			" tP.st_entrega" & _
		" FROM t_PEDIDO tP" & _
			" INNER JOIN t_CLIENTE tC ON (tP.id_cliente = tC.id)" & _
		" WHERE" & _
			" (tC.cnpj_cpf = '" & retorna_so_digitos(Trim("" & tMAP_XML("cpfCnpjIdentificado"))) & "')" & _
		" ORDER BY" & _
			" tP.data_hora DESC," & _
			" tP.pedido"
	if tPED.State <> 0 then tPED.Close
	tPED.open s, cn
%>
<table class="Qx" cellspacing="0">
	<tr style="background-color:azure;">
		<td colspan="6" class="MC MB ME MD" align="center"><span class="N">PEDIDOS ANTERIORES</span></td>
	</tr>
<%
	if tPED.Eof then
%>
	<tr>
		<td colspan="6" class="MB ME MD" align="center"><span class="N">&nbsp;(nenhum pedido encontrado)&nbsp;</span></td>
	</tr>
<%
	else
%>
	<tr>
		<td class="MB ME MD TdPedAntPedLbl"><span class="PLTe">Pedido</span></td>
		<td class="MB MD TdPedAntDataLbl"><span class="PLTc">Data</span></td>
		<td class="MB MD TdPedAntStEntLbl"><span class="PLTc">Status</span></td>
		<td class="MB MD TdPedAntProdLbl"><span class="PLTc">Prod</span></td>
		<td class="MB MD TdPedAntQtdeLbl"><span class="PLTd">Qtde</span></td>
		<td class="MB MD TdPedAntDescLbl"><span class="PLTe">Descrição</span></td>
	</tr>
<%
		do while Not tPED.Eof
			qtde_rowspan = 0
			s_row = "	<tr>" & chr(13) & _
					"		<td" & rowspan_placeholder & " class=""MB ME MD TdPedAntPed""><span class=""C SpnPedAntPed"">" & Trim("" & tPED("pedido")) & "</span></td>" & chr(13) & _
					"		<td" & rowspan_placeholder & " class=""MB MD TdPedAntData""><span class=""C SpnPedAntData"">" & formata_data(tPED("data")) & "</span></td>" & chr(13) & _
					"		<td" & rowspan_placeholder & " class=""MB MD TdPedAntStEnt""><span class=""C SpnPedAntStEnt"" style=""color:" & x_status_entrega_cor(Trim("" & tPED("st_entrega")), Trim("" & tPED("pedido"))) & ";"">" & x_status_entrega(Trim("" & tPED("st_entrega"))) & "</span></td>" & chr(13)

			s = "SELECT" & _
					" tPI.fabricante," & _
					" tPI.produto," & _
					" tPI.qtde," & _
					" tP.descricao," & _
					" tP.descricao_html" & _
				" FROM t_PEDIDO_ITEM tPI" & _
					" INNER JOIN t_PRODUTO tP ON (tPI.fabricante = tP.fabricante) AND (tPI.produto = tP.produto)" & _
				" WHERE" & _
					" (pedido = '" & Trim("" & tPED("pedido")) & "')" & _
				" ORDER BY" & _
					" tPI.sequencia"
			if tPEDITM.State <> 0 then tPEDITM.Close
			tPEDITM.open s, cn
			do while Not tPEDITM.Eof
				qtde_rowspan = qtde_rowspan + 1
				if qtde_rowspan > 1 then s_row = s_row & _
											"	</tr>" & chr(13) & _
											"	<tr>" & chr(13)
				s_row = s_row & _
						"		<td class=""MB MD TdPedAntProd""><span class=""C SpnPedAntProd"">" & Trim("" & tPEDITM("produto")) & "</span></td>" & chr(13) & _
						"		<td class=""MB MD TdPedAntQtde""><span class=""C SpnPedAntQtde"">" & Trim("" & tPEDITM("qtde")) & "</span></td>" & chr(13) & _
						"		<td class=""MB MD TdPedAntDesc""><span class=""C SpnPedAntDesc"">" & Trim("" & tPEDITM("descricao")) & "</span></td>" & chr(13)
				tPEDITM.MoveNext
				loop

			s_row = s_row & _
					"	</tr>" & chr(13)

			if qtde_rowspan <= 1 then
				s_row = Replace(s_row, rowspan_placeholder, "")
			else
				s_row = Replace(s_row, rowspan_placeholder, " rowspan=" & chr(34) & Cstr(qtde_rowspan) & chr(34))
				end if

			Response.Write s_row
%>
<%
			tPED.MoveNext
			loop
		end if
%>
</table>

<% if blnFlagCadSemiAutoPedMagentoCadAutoClienteNovo And _
	  ( _
		( (tipoCliente = ID_PF) And (blnErroIbgeEndEtg) ) _
		Or _
		( (tipoCliente = ID_PJ) And (blnErroIbgeEndCob Or blnErroIbgeEndEtg) ) _
	  ) then %>
<!--  PULA LINHA  -->
<br /><br />
<table class="Qx" cellspacing="0">
	<tr style="background-color:lightyellow;">
		<td colspan="2" class="MC MB ME MD" align="center"><span class="N" style="color:red;">ATENÇÃO</span></td>
	</tr>
	<% if blnErroIbgeEndCob And (tipoCliente = ID_PJ) then %>
	<tr>
		<td class="MB ME MD TdEndLbl"><span class="PLTd">Endereço de Cobrança</span></td>
		<td class="MB MD TdEndMsg"><span class="C">O município de <%=Trim("" & tMAP_END_COB("cidade"))%>/<%=Trim("" & tMAP_END_COB("uf"))%> não consta no cadastro do IBGE</span></td>
	</tr>
	<% end if %>
	<% if blnErroIbgeEndEtg And _
		  ( _
			(tipoCliente = ID_PF) _
			Or _
			( (tipoCliente = ID_PJ) And (Not blnEnderecosIguais) ) _
		  ) then
	%>
	<tr>
		<td class="MB ME MD TdEndLbl"><span class="PLTd">Endereço de Entrega</span></td>
		<td class="MB MD TdEndMsg"><span class="C">O município de <%=Trim("" & tMAP_END_ETG("cidade"))%>/<%=Trim("" & tMAP_END_ETG("uf"))%> não consta no cadastro do IBGE</span></td>
	</tr>
	<% end if %>
	<tr>
		<td class="MB ME MD TdEndLbl"><span class="PLTd">Aviso</span></td>
		<td class="MB MD TdEndMsg"><span class="C">Devido à necessidade de corrigir o nome do município, o fluxo de cadastramento será alterado para seguir para a página de edição do cliente</span></td>
	</tr>
</table>
<% end if %>

<!-- ************   SEPARADOR   ************ -->
<table width="849" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="849" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPNECConfirma(fPNEC)" title="prossegue com o cadastramento do pedido de e-commerce seguindo para o cadastro do cliente">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% if blnFlagCadSemiAutoPedMagento_FluxoOtimizado then %>
<tr>
	<td>&nbsp;</td>
	<td align="right"><div name="dCONFIRMA2" id="dCONFIRMA2"><a name="bCONFIRMA2" id="bCONFIRMA2" href="javascript:fPNEC2Confirma(fPNEC)" title="prossegue com o cadastramento do pedido de e-commerce seguindo o fluxo otimizado">
		<img src="../botao/pedido.gif" width="176" height="55" border="0"></a></div>
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
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
	if tMAP_ITEM.State <> 0 then tMAP_ITEM.Close
	set tMAP_ITEM = nothing
	
	if tMAP_END_COB.State <> 0 then tMAP_END_COB.Close
	set tMAP_END_COB = nothing

	if tMAP_END_ETG.State <> 0 then tMAP_END_ETG.Close
	set tMAP_END_ETG = nothing

	if tMAP_XML.State <> 0 then tMAP_XML.Close
	set tMAP_XML = nothing

	if tPROD.State <> 0 then tPROD.Close
	set tPROD = nothing

	if tPCI.State <> 0 then tPCI.Close
	set tPCI = nothing

	if tPED.State <> 0 then tPED.Close
	set tPED = nothing

	if tPEDITM.State <> 0 then tPEDITM.Close
	set tPEDITM = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>