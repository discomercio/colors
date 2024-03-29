<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  P E D I D O E D I T A . A S P
'     ===========================================
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

'	EXIBI��O DE BOT�ES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True


	dim s, usuario, loja, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    dim url_origem
    url_origem = Trim(Request("url_origem"))

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
	dim i, n, nColSpan, x, s_fabricante, s_produto, s_descricao, s_descricao_html, s_qtde, s_preco_lista, s_desc_dado
	dim s_vl_unitario, s_vl_TotalItem, m_TotalItem, m_TotalDestePedido, m_TotalItemComRA, m_TotalDestePedidoComRA
	dim s_preco_NF, m_TotalFamiliaParcelaRA
	dim m_total_RA_deste_pedido, m_total_venda_deste_pedido, m_total_RA_outros, m_total_venda_outros
	dim m_total_NF_deste_pedido, m_total_NF_outros
	dim s_readonly, s_readonly_RT, s_readonly_RA, rs, sql
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim max_qtde_itens
	max_qtde_itens = obtem_parametro_PedidoItem_MaxQtdeItens




' _____________________________________________________________________________________________
'
'									F  U  N  �  �  E  S 
' _____________________________________________________________________________________________
' ____________________________________________________________________________
' INDICADORES MONTA ITENS SELECT
function indicadores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT apelido, razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR WHERE (Id NOT IN (" & Cstr(ID_NSU_ORCAMENTISTA_E_INDICADOR__RESTRICAO_FP_TODOS) & "," & Cstr(ID_NSU_ORCAMENTISTA_E_INDICADOR__SEM_INDICADOR) & ")) ORDER BY apelido")
	strResp = "<option value=''>&nbsp;</option>"
	do while Not r.eof 
		x = UCase(Trim("" & r("apelido")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("razao_social_nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
		
	indicadores_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' ____________________________________________
' JUSTIFICATIVA ENDERE�O MONTA ITENS SELECT
'
function justificativa_endereco_etg_monta_itens(byval grupo, byval id_default)
dim s, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_CODIGO_DESCRICAO" & _
		" WHERE" & _
			" (grupo='" & grupo & "')" & _						
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.Eof
		x = UCase(Trim("" & r("codigo")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & iniciais_em_maiusculas(Trim("" & r("descricao")))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
		
	justificativa_endereco_etg_monta_itens = strResp
	r.close
	set r=nothing
end function

	dim r_pedido, v_item, alerta, msg_erro
	alerta=""
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
	else
		if Trim(r_pedido.loja) <> loja then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_INVALIDO)
		if Not le_pedido_item(pedido_selecionado, v_item, msg_erro) then alerta = msg_erro
		'Assegura que dados cadastrados anteriormente sejam exibidos corretamente, mesmo se o par�metro da quantidade m�xima de itens tiver sido reduzido
		if VectorLength(v_item) > max_qtde_itens then max_qtde_itens = VectorLength(v_item)
		end if

	dim r_cliente, tipo_cliente
	set r_cliente = New cl_CLIENTE
	dim xcliente_bd_resultado
	xcliente_bd_resultado = x_cliente_bd(r_pedido.id_cliente, r_cliente)
	tipo_cliente = r_cliente.tipo
	
	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
    'Definido em 20/03/2020: para os pedidos criado antes da memoriza��o completa, vamos usar a tela anterior.
    'N�o queremos exigir que quem editar o pedido seja obrigado a preenhcer o CNPJ do endere�o sde entrega. Ent�o, para
    'um pedido criado sem a memoriza��o, ele continua sempre sem a memoriza��o.
    if r_pedido.st_memorizacao_completa_enderecos = 0 then
        blnUsarMemorizacaoCompletaEnderecos  = false
        end if

	dim eh_cpf
	if len(r_cliente.cnpj_cpf)=11 then eh_cpf=True else eh_cpf=False

    'le as vari�veis da origem certa: ou do pedido ou do cliente, todas comecam com cliente__
    dim cliente__tipo, cliente__cnpj_cpf, cliente__rg, cliente__ie, cliente__nome
    dim cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep
    dim cliente__tel_res, cliente__ddd_res, cliente__tel_com, cliente__ddd_com, cliente__ramal_com, cliente__tel_cel, cliente__ddd_cel
    dim cliente__tel_com_2, cliente__ddd_com_2, cliente__ramal_com_2, cliente__email, cliente__email_xml, cliente__icms, cliente__produtor_rural_status

    cliente__tipo = r_cliente.tipo
    cliente__cnpj_cpf = r_cliente.cnpj_cpf
	cliente__rg = r_cliente.rg
    cliente__ie = r_cliente.ie
    cliente__nome = r_cliente.nome
    cliente__endereco = r_cliente.endereco
    cliente__endereco_numero = r_cliente.endereco_numero
    cliente__endereco_complemento = r_cliente.endereco_complemento
    cliente__bairro = r_cliente.bairro
    cliente__cidade = r_cliente.cidade
    cliente__uf = r_cliente.uf
    cliente__cep = r_cliente.cep
    cliente__tel_res = r_cliente.tel_res
    cliente__ddd_res = r_cliente.ddd_res
    cliente__tel_com = r_cliente.tel_com
    cliente__ddd_com = r_cliente.ddd_com
    cliente__ramal_com = r_cliente.ramal_com
    cliente__tel_cel = r_cliente.tel_cel
    cliente__ddd_cel = r_cliente.ddd_cel
    cliente__tel_com_2 = r_cliente.tel_com_2
    cliente__ddd_com_2 = r_cliente.ddd_com_2
    cliente__ramal_com_2 = r_cliente.ramal_com_2
    cliente__email = r_cliente.email
	cliente__email_xml = r_cliente.email_xml
	cliente__icms = r_cliente.contribuinte_icms_status
	cliente__produtor_rural_status = r_cliente.produtor_rural_status
   
    if blnUsarMemorizacaoCompletaEnderecos and r_pedido.st_memorizacao_completa_enderecos <> 0 then 
        cliente__tipo = r_pedido.endereco_tipo_pessoa
        cliente__cnpj_cpf = r_pedido.endereco_cnpj_cpf
	    cliente__rg = r_pedido.endereco_rg
        cliente__ie = r_pedido.endereco_ie
        cliente__nome = r_pedido.endereco_nome
        cliente__endereco = r_pedido.endereco_logradouro
        cliente__endereco_numero = r_pedido.endereco_numero
        cliente__endereco_complemento = r_pedido.endereco_complemento
        cliente__bairro = r_pedido.endereco_bairro
        cliente__cidade = r_pedido.endereco_cidade
        cliente__uf = r_pedido.endereco_uf
        cliente__cep = r_pedido.endereco_cep
        cliente__tel_res = r_pedido.endereco_tel_res
        cliente__ddd_res = r_pedido.endereco_ddd_res
        cliente__tel_com = r_pedido.endereco_tel_com
        cliente__ddd_com = r_pedido.endereco_ddd_com
        cliente__ramal_com = r_pedido.endereco_ramal_com
        cliente__tel_cel = r_pedido.endereco_tel_cel
        cliente__ddd_cel = r_pedido.endereco_ddd_cel
        cliente__tel_com_2 = r_pedido.endereco_tel_com_2
        cliente__ddd_com_2 = r_pedido.endereco_ddd_com_2
        cliente__ramal_com_2 = r_pedido.endereco_ramal_com_2
        cliente__email = r_pedido.endereco_email
		cliente__email_xml = r_pedido.endereco_email_xml
		cliente__icms = r_pedido.endereco_contribuinte_icms_status
		cliente__produtor_rural_status = r_pedido.endereco_produtor_rural_status
    end if

	dim rCD
	set rCD = obtem_perc_max_comissao_e_desconto_por_loja(loja)

'	OBT�M A RELA��O DE MEIOS DE PAGAMENTO PREFERENCIAIS (QUE FAZEM USO O PERCENTUAL DE COMISS�O+DESCONTO N�VEL 2)
	dim rP, vMPN2, strScriptJS_MPN2
	set rP = get_registro_t_parametro(ID_PARAMETRO_PercMaxComissaoEDesconto_Nivel2_MeiosPagto)
	
	strScriptJS_MPN2 = "<script type='text/javascript'>" & chr(13) & _
						"var vMPN2 = new Array();" & chr(13) & _
						"vMPN2[0] = 0;" & chr(13)
	if Trim("" & rP.id) <> "" then
		vMPN2 = Split(rP.campo_texto, ",")
		for i=Lbound(vMPN2) to Ubound(vMPN2)
			vMPN2(i) = Trim("" & vMPN2(i))
			if vMPN2(i) <> "" then
				strScriptJS_MPN2 = strScriptJS_MPN2 & _
									"vMPN2[vMPN2.length] = " & vMPN2(i) & ";" & chr(13)
				end if
			next
		end if
	strScriptJS_MPN2 = strScriptJS_MPN2 & _
						"</script>" & chr(13)
	
	dim strPercMaxRT, strPercMaxRTAlcada1, strPercMaxRTAlcada2, strPercMaxRTAlcada3
	dim strPercMaxComissaoEDesconto, strPercMaxComissaoEDescontoPj, strPercMaxComissaoEDescontoNivel2, strPercMaxComissaoEDescontoNivel2Pj
	dim strPercMaxDescAlcada1Pf, strPercMaxDescAlcada1Pj, strPercMaxDescAlcada2Pf, strPercMaxDescAlcada2Pj, strPercMaxDescAlcada3Pf, strPercMaxDescAlcada3Pj
	strPercMaxRT = formata_perc(rCD.perc_max_comissao)
	strPercMaxComissaoEDesconto = formata_perc(rCD.perc_max_comissao_e_desconto)
	strPercMaxComissaoEDescontoPj = formata_perc(rCD.perc_max_comissao_e_desconto_pj)
	strPercMaxComissaoEDescontoNivel2 = formata_perc(rCD.perc_max_comissao_e_desconto_nivel2)
	strPercMaxComissaoEDescontoNivel2Pj = formata_perc(rCD.perc_max_comissao_e_desconto_nivel2_pj)
	strPercMaxRTAlcada1 = "0"
	strPercMaxDescAlcada1Pf = "0"
	strPercMaxDescAlcada1Pj = "0"
	strPercMaxRTAlcada2 = "0"
	strPercMaxDescAlcada2Pf = "0"
	strPercMaxDescAlcada2Pj = "0"
	strPercMaxRTAlcada3 = "0"
	strPercMaxDescAlcada3Pf = "0"
	strPercMaxDescAlcada3Pj = "0"
	
	if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_1, s_lista_operacoes_permitidas) then
		strPercMaxRTAlcada1 = formata_perc(rCD.perc_max_comissao_alcada1)
		strPercMaxDescAlcada1Pf = formata_perc(rCD.perc_max_comissao_e_desconto_alcada1_pf)
		strPercMaxDescAlcada1Pj = formata_perc(rCD.perc_max_comissao_e_desconto_alcada1_pj)
		end if

	if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_2, s_lista_operacoes_permitidas) then
		strPercMaxRTAlcada2 = formata_perc(rCD.perc_max_comissao_alcada2)
		strPercMaxDescAlcada2Pf = formata_perc(rCD.perc_max_comissao_e_desconto_alcada2_pf)
		strPercMaxDescAlcada2Pj = formata_perc(rCD.perc_max_comissao_e_desconto_alcada2_pj)
		end if

	if operacao_permitida(OP_LJA_DESC_SUP_ALCADA_3, s_lista_operacoes_permitidas) then
		strPercMaxRTAlcada3 = formata_perc(rCD.perc_max_comissao_alcada3)
		strPercMaxDescAlcada3Pf = formata_perc(rCD.perc_max_comissao_e_desconto_alcada3_pf)
		strPercMaxDescAlcada3Pj = formata_perc(rCD.perc_max_comissao_e_desconto_alcada3_pj)
		end if

	dim blnUsuarioDeptoFinanceiro, vDeptoSetorUsuario
	blnUsuarioDeptoFinanceiro = False
	
	if alerta = "" then
		if Not obtem_Usuario_x_DeptoSetor(usuario, vDeptoSetorUsuario, msg_erro) then
			alerta=texto_add_br(alerta)
			alerta = alerta & msg_erro
		else
			for i=LBound(vDeptoSetorUsuario) to UBound(vDeptoSetorUsuario)
				if (vDeptoSetorUsuario(i).StInativo = 0) then
					if (vDeptoSetorUsuario(i).Id = ID_DEPTO_SETOR__FIN_FINANCEIRO) Or (vDeptoSetorUsuario(i).Id = ID_DEPTO_SETOR__FIN_CREDITO) then
						blnUsuarioDeptoFinanceiro = True
						exit for
						end if
					end if
				next
			end if
		end if

	dim blnTemRA
	blnTemRA = False
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			if Trim("" & v_item(i).produto) <> "" then
				if v_item(i).preco_NF <> v_item(i).preco_venda then
					blnTemRA = True
					exit for
					end if
				end if
			next
		end if

	dim s_aux, s2, s3, s4, r_loja, s_cor, s_falta
	dim v_disp
	dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF
	dim vl_saldo_a_pagar, s_vl_saldo_a_pagar, st_pagto
	dim v_item_devolvido, s_devolucoes
	dim v_pedido_perda, s_perdas, vl_total_perdas, vl_total_frete, frete_transportadora_id, frete_numero_NF, intQtdeFrete, frete_serie_NF
	dim intIdx
	dim strDisabled
	s_devolucoes = ""
	s_perdas = ""
	vl_total_perdas = 0
	
	if alerta = "" then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then
			redim v_disp(Ubound(v_item))
			for i=Lbound(v_disp) to Ubound(v_disp)
				set v_disp(i) = New cl_ITEM_STATUS_ESTOQUE
				v_disp(i).pedido		= v_item(i).pedido
				v_disp(i).fabricante	= v_item(i).fabricante
				v_disp(i).produto		= v_item(i).produto
				v_disp(i).qtde			= v_item(i).qtde
				next
			
			if Not estoque_verifica_status_item(v_disp, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			end if

	'	OBT�M OS VALORES A PAGAR, J� PAGO E O STATUS DE PAGAMENTO (PARA TODA A FAM�LIA DE PEDIDOS)
		if Not calcula_pagamentos(pedido_selecionado, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		m_TotalFamiliaParcelaRA = vl_TotalFamiliaPrecoNF - vl_TotalFamiliaPrecoVenda
		vl_saldo_a_pagar = vl_TotalFamiliaPrecoNF - vl_TotalFamiliaPago - vl_TotalFamiliaDevolucaoPrecoNF
		s_vl_saldo_a_pagar = formata_moeda(vl_saldo_a_pagar)
	'	VALORES NEGATIVOS REPRESENTAM O 'CR�DITO' QUE O CLIENTE POSSUI EM CASO DE PEDIDOS CANCELADOS QUE HAVIAM SIDO PAGOS
		if (st_pagto = ST_PAGTO_PAGO) And (vl_saldo_a_pagar > 0) then s_vl_saldo_a_pagar = ""
		
	'	H� DEVOLU��ES?
		if Not le_pedido_item_devolvido(pedido_selecionado, v_item_devolvido, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		for i=Lbound(v_item_devolvido) to Ubound(v_item_devolvido)
			with v_item_devolvido(i)
				if .produto <> "" then
					if .qtde = 1 then s = "" else s = "s"
					if s_devolucoes <> "" then s_devolucoes = s_devolucoes & chr(13) & "<br>" & chr(13)
					s_devolucoes = s_devolucoes & formata_data(.devolucao_data) & " " & _
								   formata_hhnnss_para_hh_nn(.devolucao_hora) & " - " & _
								   formata_inteiro(.qtde) & " unidade" & s & " do " & .produto & " - " & produto_formata_descricao_em_html(.descricao_html)
					if Trim(.motivo) <> "" then s_devolucoes = s_devolucoes & " (" & .motivo & ")"
					if .NFe_numero_NF > 0 then s_devolucoes = s_devolucoes & " [NF: " & .NFe_numero_NF & "]"
					end if
				end with
			next

	'	H� PERDAS?
		if Not le_pedido_perda(pedido_selecionado, v_pedido_perda, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		for i=Lbound(v_pedido_perda) to Ubound(v_pedido_perda)
			with v_pedido_perda(i)
				if .id <> "" then
					vl_total_perdas = vl_total_perdas + .valor
					if s_perdas <> "" then s_perdas = s_perdas & chr(13) & "<br>" & chr(13)
					s_perdas = s_perdas & formata_data(.data) & " " & _
							   formata_hhnnss_para_hh_nn_ss(.hora) & ": " & SIMBOLO_MONETARIO & " " & formata_moeda(.valor)
					if Trim(.obs) <> "" then s_perdas = s_perdas & " (" & .obs & ")"
					end if
				end with
			next
		end if

	dim blnPedidoEntregue
	blnPedidoEntregue = False
	if Trim("" & r_pedido.st_entrega) = ST_ENTREGA_ENTREGUE then blnPedidoEntregue = True

	dim blnNFEmitida
	blnNFEmitida = False
	if Trim("" & r_pedido.obs_2) <> "" then blnNFEmitida = True
	
	dim blnAnaliseCreditoProcessado
	blnAnaliseCreditoProcessado = False
	if Trim("" & r_pedido.analise_credito) <> Cstr(COD_AN_CREDITO_ST_INICIAL) And _
	   Trim("" & r_pedido.analise_credito) <> Cstr(COD_AN_CREDITO_PENDENTE) And _
	   Trim("" & r_pedido.analise_credito) <> Cstr(COD_AN_CREDITO_NAO_ANALISADO) then 
	   blnAnaliseCreditoProcessado = True
	   end if

	'Edi��o do indicador est� liberada?
    sql = "SELECT * FROM t_COMISSAO_INDICADOR_N4 WHERE (pedido='" & r_pedido.pedido & "')"
    set rs = cn.Execute(sql)
    dim blnIndicadorEdicaoLiberada
    blnIndicadorEdicaoLiberada = False
    if operacao_permitida(OP_LJA_EDITA_PEDIDO_INDICADOR, s_lista_operacoes_permitidas) then
        if r_pedido.st_entrega<>ST_ENTREGA_CANCELADO And rs.Eof then
            blnIndicadorEdicaoLiberada = True
        end if 
    end if
    if rs.State <> 0 then rs.Close
	'04/Jun/2021: a edi��o do indicador foi desmembrada em uma opera��o espec�fica para melhorar o tempo de carregamento da p�gina de edi��o do pedido (PedidoEditaIndicador.asp)
	blnIndicadorEdicaoLiberada = False
	
	dim blnObs1EdicaoLiberada
	blnObs1EdicaoLiberada = False
	if operacao_permitida(OP_LJA_EDITA_PEDIDO, s_lista_operacoes_permitidas) Or _
	   operacao_permitida(OP_LJA_EDITA_PEDIDO_OBS1, s_lista_operacoes_permitidas) then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then
			if (Not blnAnaliseCreditoProcessado) And (Not blnNFEmitida) then blnObs1EdicaoLiberada = True
			end if
		end if
	
	dim nivelEdicaoFormaPagto
	nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_BLOQUEADA
	if operacao_permitida(OP_LJA_EDITA_PEDIDO_FORMA_PAGTO, s_lista_operacoes_permitidas) Or operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
		nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL

		' Analisa situa��es que liberam apenas parcialmente a edi��o da forma de pagamento, ou seja,
		' pode-se alterar os valores da forma de pagamento atualmente selecionada, mas n�o se pode
		' alterar a forma de pagamento e nem os meios de pagamento (ex: de '� Vista' para 
		' 'Parcelado com Entrada' ou de 'Dep�sito' para 'Boleto').
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'Se o status da an�lise de cr�dito est� em uma situa��o que demanda uma confirma��o manual do depto de an�lise de cr�dito, bloqueia
		'a edi��o da forma de pagamento para n�o haver o risco de uma altera��o ser feita sem o conhecimento do depto de an�lise de cr�dito.
		'Qualquer altera��o necess�ria na forma de pagamento deve ser solicitada ao depto de an�lise de cr�dito.
		if (nivelEdicaoFormaPagto > COD_NIVEL_EDICAO_LIBERADA_PARCIAL) _
			AND _
			(Cstr(r_pedido.loja) <> Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE)) _
			AND _
			( _
				(Trim("" & r_pedido.analise_credito) = Cstr(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)) _
				OR (Trim("" & r_pedido.analise_credito) = Cstr(COD_AN_CREDITO_OK)) _
			) then
			nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_PARCIAL
			end if
		
		' Analisa situa��es em que a edi��o da forma de pagamento deve ser bloqueada totalmente
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if Trim("" & r_pedido.st_entrega) = ST_ENTREGA_ENTREGUE then
			nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_BLOQUEADA
			end if

		if Trim("" & r_pedido.st_entrega) = ST_ENTREGA_CANCELADO then
			nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_BLOQUEADA
			end if
		end if 'if operacao_permitida(OP_LJA_EDITA_PEDIDO_FORMA_PAGTO, s_lista_operacoes_permitidas) Or operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas)

	dim perc_comissao_e_desconto_n1_n2_a_utilizar
	if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_BLOQUEADA then
		'Determina os percentuais de comiss�o de desconto n�vel 1 e 2
		perc_comissao_e_desconto_n1_n2_a_utilizar = obtem_perc_comissao_e_desconto_n1_n2_a_utilizar(tipo_cliente, r_pedido, v_item)
		end if

	dim strPercLimiteRASemDesagio, strPercDesagio
	if alerta = "" then
		strPercLimiteRASemDesagio = formata_perc(r_pedido.perc_limite_RA_sem_desagio)
		strPercDesagio = formata_perc(r_pedido.perc_desagio_RA)
		end if

'	SE FOI APLICADO DES�GIO E HOUVE ALGUM PEDIDO ENTREGUE NESTA FAM�LIA
'	DE PEDIDOS, ENT�O � OBRIGAT�RIO QUE O DES�GIO SEJA MANTIDO DAQUI P/ FRENTE.
	dim strOpcaoForcaDesagio, qtde_pedidos_entregues
	strOpcaoForcaDesagio = "N"
	qtde_pedidos_entregues = familia_pedidos_qtde_pedidos_entregues(pedido_selecionado)
	if (CStr(r_pedido.st_tem_desagio_RA)<>CStr(0)) And (qtde_pedidos_entregues > 0) then strOpcaoForcaDesagio = "S"

'	CRIT�RIO PARA EDITAR O ENDERE�O DE ENTREGA E OS DADOS CADASTRAIS
	dim blnEndEntregaEdicaoLiberada
	blnEndEntregaEdicaoLiberada = False
	if operacao_permitida(OP_LJA_EDITA_PEDIDO, s_lista_operacoes_permitidas) then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then
            
			if r_pedido.obs_2 = "" then blnEndEntregaEdicaoLiberada = True
			end if
		end if

	dim blnDadosCadastraisEdicaoLiberada
	blnDadosCadastraisEdicaoLiberada = blnEndEntregaEdicaoLiberada

	dim strAtributosDadosCadastrais
	strAtributosDadosCadastrais = ""
	if not blnDadosCadastraisEdicaoLiberada then
		strAtributosDadosCadastrais = " readonly tabindex=-1 "
		end if
	dim strAtributosRadioboxDadosCadastrais
	strAtributosRadioboxDadosCadastrais = ""
	if not blnDadosCadastraisEdicaoLiberada then
		strAtributosRadioboxDadosCadastrais = " disabled "
		end if

'	Para assegurar a consist�ncia entre o valor total de NF e o total da forma de pagamento,
'	a edi��o fica permitida somente se o usu�rio puder editar a forma de pagamento!
	dim bln_RA_EdicaoLiberada
	bln_RA_EdicaoLiberada = False
	if operacao_permitida(OP_LJA_EDITA_RA, s_lista_operacoes_permitidas) then
		if (Cstr(r_pedido.comissao_paga) = Cstr(COD_COMISSAO_NAO_PAGA)) And (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then bln_RA_EdicaoLiberada = True
		end if
	
	'A regra de edi��o do percentual de RT leva em considera��o que o percentual � �nico p/ toda a fam�lia de pedidos
	dim blnFamiliaPedidosPossuiPedidoEntregueMesAnterior, blnFamiliaPedidosPossuiPedidoComissaoPaga, blnFamiliaPedidosPossuiPedidoComissaoDescontada
	blnFamiliaPedidosPossuiPedidoEntregueMesAnterior = False
	blnFamiliaPedidosPossuiPedidoComissaoPaga = False
	blnFamiliaPedidosPossuiPedidoComissaoDescontada = False

	sql = "SELECT" & _
				" pedido" & _
				", comissao_descontada" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
			" WHERE" & _
				" (pedido LIKE '" & retorna_num_pedido_base(pedido_selecionado) & BD_CURINGA_TODOS & "')" & _
				" AND (comissao_descontada = " & COD_COMISSAO_DESCONTADA & ")"
	set rs = cn.Execute(sql)
	if Not rs.Eof then blnFamiliaPedidosPossuiPedidoComissaoDescontada = True
	if rs.State <> 0 then rs.Close

	sql = "SELECT" & _
				" pedido" & _
				", comissao_descontada" & _
			" FROM t_PEDIDO_PERDA" & _
			" WHERE" & _
				" (pedido LIKE '" & retorna_num_pedido_base(pedido_selecionado) & BD_CURINGA_TODOS & "')" & _
				" AND (comissao_descontada = " & COD_COMISSAO_DESCONTADA & ")"
	set rs = cn.Execute(sql)
	if Not rs.Eof then blnFamiliaPedidosPossuiPedidoComissaoDescontada = True
	if rs.State <> 0 then rs.Close

	sql = "SELECT" & _
				" pedido" & _
				", st_entrega" & _
				", entregue_data" & _
				", comissao_paga" & _
			" FROM t_PEDIDO" & _
			" WHERE" & _
				" (pedido LIKE '" & retorna_num_pedido_base(pedido_selecionado) & BD_CURINGA_TODOS & "')"
	set rs = cn.Execute(sql)
	do while Not rs.Eof
		if (Trim("" & rs("st_entrega")) = ST_ENTREGA_ENTREGUE) And (Not IsMesmoAnoEMes(rs("entregue_data"), Date)) then blnFamiliaPedidosPossuiPedidoEntregueMesAnterior = True
		if CLng(rs("comissao_paga")) = CLng(COD_COMISSAO_PAGA) then blnFamiliaPedidosPossuiPedidoComissaoPaga = True
		rs.MoveNext
		loop
	if rs.State <> 0 then rs.Close
	
	dim bln_RT_EdicaoLiberada, rEdicaoRTMaxPrazo
	bln_RT_EdicaoLiberada = False
	set rEdicaoRTMaxPrazo = get_registro_t_parametro(ID_PARAMETRO_Pedido_RT_Edicao_MaxPrazo)
	if operacao_permitida(OP_LJA_EDITA_RT, s_lista_operacoes_permitidas) _
		And ( (rEdicaoRTMaxPrazo.campo_inteiro = 0) Or (Abs(DateDiff("d", r_pedido.data, Date)) <= rEdicaoRTMaxPrazo.campo_inteiro) ) then
		if (Not blnFamiliaPedidosPossuiPedidoComissaoPaga) _
			And (Not blnFamiliaPedidosPossuiPedidoComissaoDescontada) _
			And (Not blnFamiliaPedidosPossuiPedidoEntregueMesAnterior) then
			bln_RT_EdicaoLiberada = True
			end if
		end if
	
	dim blnItemPedidoEdicaoLiberada
	blnItemPedidoEdicaoLiberada = False
	if operacao_permitida(OP_LJA_EDITA_ITEM_DO_PEDIDO, s_lista_operacoes_permitidas) then
		if (Not IsPedidoEncerrado(r_pedido.st_entrega)) then
			if (Not blnAnaliseCreditoProcessado) And (Not blnNFEmitida) And (Cstr(r_pedido.comissao_paga)=Cstr(COD_COMISSAO_NAO_PAGA)) then blnItemPedidoEdicaoLiberada = True
			end if
		end if

    sql = "SELECT * FROM t_PEDIDO_NFe_EMISSAO_SOLICITADA WHERE (pedido='" & r_pedido.pedido & "')"
    set rs = cn.Execute(sql)
	dim blnEtgImediataEdicaoLiberada
	blnEtgImediataEdicaoLiberada = False
    if rs.Eof then
	    if operacao_permitida(OP_LJA_EDITA_CAMPO_ENTREGA_IMEDIATA, s_lista_operacoes_permitidas) then
		    if Not IsPedidoEncerrado(r_pedido.st_entrega) then
				if r_pedido.st_entrega = ST_ENTREGA_A_ENTREGAR then
					'NOP
					'A edi��o do campo 'Entrega Imediata' foi bloqueada em pedidos com status 'A Entregar' e a edi��o s� pode ser realizada na Central mediante permiss�o de acesso espec�fica
				else
					if (Cstr(r_pedido.obs_2)="") And (Cstr(r_pedido.obs_3)="") then
						blnEtgImediataEdicaoLiberada = True
						end if
					end if
			    end if
		    end if
        end if
    if rs.State <> 0 then rs.Close

    dim blnNumPedidoECommerceEdicaoLiberada
    blnNumPedidoECommerceEdicaoLiberada=False
    if operacao_permitida(OP_LJA_EDITA_PEDIDO_NUM_PEDIDO_ECOMMERCE, s_lista_operacoes_permitidas) And _
		( (loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE) Or (r_pedido.plataforma_origem_pedido = COD_PLATAFORMA_ORIGEM_PEDIDO__MAGENTO) ) then
        blnNumPedidoECommerceEdicaoLiberada=True
    end if
	
	dim blnAnaliseCreditoEdicaoLiberada
	blnAnaliseCreditoEdicaoLiberada = False
	if operacao_permitida(OP_LJA_EDITA_ANALISE_CREDITO_PENDENTE_VENDAS, s_lista_operacoes_permitidas) then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then
			if ( (r_pedido.transportadora_id = "") Or (r_pedido.transportadora_selecao_auto_status <> 0) ) _
				And (Trim("" & r_pedido.a_entregar_data_marcada) = "") then
				if (Cstr(r_pedido.analise_credito) = Cstr(COD_AN_CREDITO_PENDENTE_VENDAS)) then blnAnaliseCreditoEdicaoLiberada = True
				end if
			end if
		end if
		
	dim blnBemUsoConsumoEdicaoLiberada
	blnBemUsoConsumoEdicaoLiberada = False
	if operacao_permitida(OP_LJA_EDITA_PEDIDO, s_lista_operacoes_permitidas) then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then blnBemUsoConsumoEdicaoLiberada = True
		end if
	
	dim blnGarantiaIndicadorEdicaoLiberada
	blnGarantiaIndicadorEdicaoLiberada = False
	if operacao_permitida(OP_LJA_EDITA_PEDIDO_GARANTIA_INDICADOR, s_lista_operacoes_permitidas) then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then blnGarantiaIndicadorEdicaoLiberada = True
		end if
		
	dim strScriptJS
	strScriptJS = "<script language='JavaScript'>" & chr(13) & _
				  "var PERC_DESAGIO_RA_LIQUIDA_PEDIDO = " & js_formata_numero(r_pedido.perc_desagio_RA_liquida) & ";" & chr(13) & _
				  "var nivelEdicaoFormaPagto = " & CStr(nivelEdicaoFormaPagto) & ";" & chr(13)

	if blnTemRA then s = "true" else s = "false"
	strScriptJS = strScriptJS & _
				  "var formata_perc_desconto = formata_perc_2dec;" & chr(13) & _
				  "var formata_perc_desc_linear = formata_perc_2dec;" & chr(13) & _
				  "var blnTemRA = " & s & ";" & chr(13)

	if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_BLOQUEADA then
		strScriptJS = strScriptJS & _
				  "var perc_comissao_e_desconto_n1_n2_a_utilizar = " & js_formata_numero(perc_comissao_e_desconto_n1_n2_a_utilizar) & ";" & chr(13)
		end if

	if blnUsuarioDeptoFinanceiro then s = "true" else s = "false"
	strScriptJS = strScriptJS & _
				  "var blnUsuarioDeptoFinanceiro = " & s & ";" & chr(13)

	strScriptJS = strScriptJS & _
				  "</script>" & chr(13)
	
	dim strScriptJS_FPO
	strScriptJS_FPO = r_pedido.tipo_parcelamento
	if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then
		strScriptJS_FPO = strScriptJS_FPO & _
							"|" & r_pedido.av_forma_pagto
	elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then
		strScriptJS_FPO = strScriptJS_FPO & _
							"|" & r_pedido.pu_forma_pagto & _
							"|" & formata_moeda(r_pedido.pu_valor) & _
							"|" & Cstr(r_pedido.pu_vencto_apos)
	elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then
		strScriptJS_FPO = strScriptJS_FPO & _
							"|" & Cstr(r_pedido.pc_qtde_parcelas) & _
							"|" & formata_moeda(r_pedido.pc_valor_parcela)
	elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
		strScriptJS_FPO = strScriptJS_FPO & _
							"|" & Cstr(r_pedido.pc_maquineta_qtde_parcelas) & _
							"|" & formata_moeda(r_pedido.pc_maquineta_valor_parcela)
	elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
		strScriptJS_FPO = strScriptJS_FPO & _
							"|" & r_pedido.pce_forma_pagto_entrada & _
							"|" & formata_moeda(r_pedido.pce_entrada_valor) & _
							"|" & r_pedido.pce_forma_pagto_prestacao & _
							"|" & Cstr(r_pedido.pce_prestacao_qtde) & _
							"|" & formata_moeda(r_pedido.pce_prestacao_valor) & _
							"|" & Cstr(r_pedido.pce_prestacao_periodo)
	elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
		strScriptJS_FPO = strScriptJS_FPO & _
							"|" & r_pedido.pse_forma_pagto_prim_prest & _
							"|" & formata_moeda(r_pedido.pse_prim_prest_valor) & _
							"|" & Cstr(r_pedido.pse_prim_prest_apos) & _
							"|" & r_pedido.pse_forma_pagto_demais_prest & _
							"|" & Cstr(r_pedido.pse_demais_prest_qtde) & _
							"|" & formata_moeda(r_pedido.pse_demais_prest_valor) & _
							"|" & Cstr(r_pedido.pse_demais_prest_periodo)
		end if
	
	strScriptJS_FPO = "<script type='text/javascript'>" & chr(13) & _
					  "var formaPagamentoOriginal='" & strScriptJS_FPO & "';" & chr(13) & _
					  "</script>" & chr(13)
	
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
	<title>LOJA<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>



<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<%=strScriptJS%>
<%=strScriptJS_MPN2%>
<%=strScriptJS_FPO%>

<script type="text/javascript">
	$(function() {
		$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPAR�NCIA NO IE8
		$(".tdGarInd").hide();
		// Para a nova vers�o da forma de pagamento
		if ($(".tdGarInd").prev("td").hasClass("MD")) {$(".tdGarInd").prev("td").removeClass("MD")};
		// Para a vers�o antiga da forma de pagamento
		if ($(".tdGarInd").prev("td").hasClass("MDB")) {$(".tdGarInd").prev("td").removeClass("MDB").addClass("MB")}

        $("#c_data_previsao_entrega").hUtilUI('datepicker_padrao');

        $("input[name = 'rb_etg_imediata']").change(function () {
            configuraCampoDataPrevisaoEntrega();
        });

        configuraCampoDataPrevisaoEntrega();
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

    function configuraCampoDataPrevisaoEntrega() {
		if (($("input[name='rb_etg_imediata']:checked").val() == '<%=COD_ETG_IMEDIATA_NAO%>') && ($("#blnEtgImediataEdicaoLiberada").val() == '<%=CStr(True)%>')) {
            $("#c_data_previsao_entrega").prop("readonly", false);
            $("#c_data_previsao_entrega").prop("disabled", false);
            $("#c_data_previsao_entrega").datepicker("enable");
        }
        else {
			if ($("#blnEtgImediataEdicaoLiberada").val() == '<%=CStr(True)%>') $("#c_data_previsao_entrega").val("");
            $("#c_data_previsao_entrega").prop("readonly", true);
            $("#c_data_previsao_entrega").prop("disabled", true);
            $("#c_data_previsao_entrega").datepicker("disable");
        }
    }
</script>

<script language="JavaScript" type="text/javascript">
var objAjaxCustoFinancFornecConsultaPreco;
var blnConfirmaDifRAeValores=false;
var fCepPopup;
var objSenhaDesconto;
var COD_NIVEL_EDICAO_LIBERADA_TOTAL = <%=COD_NIVEL_EDICAO_LIBERADA_TOTAL%>;
var COD_NIVEL_EDICAO_LIBERADA_PARCIAL = <%=COD_NIVEL_EDICAO_LIBERADA_PARCIAL%>;
var COD_NIVEL_EDICAO_BLOQUEADA = <%=COD_NIVEL_EDICAO_BLOQUEADA%>;
var MAX_VALOR_MARGEM_ERRO_PAGAMENTO = <%=js_formata_numero(MAX_VALOR_MARGEM_ERRO_PAGAMENTO)%>;

$(function() {
    var f;
    f = fPED;
    if (f.blnEndEntregaEdicaoLiberada.value == "<%=Cstr(True)%>") {
    	$("#EndEtg_obs option[value='<%=r_pedido.EndEtg_cod_justificativa%>']").attr("selected", true);
    	// VERIFICAR MUDAN�A NOS CAMPOS
    	f.Verifica_End_Entrega.value = f.EndEtg_endereco.value;
    	f.Verifica_num.value = f.EndEtg_endereco_numero.value;
    	f.Verifica_Cidade.value = f.EndEtg_cidade.value;
    	f.Verifica_UF.value = f.EndEtg_uf.value;
    	f.Verifica_CEP.value = f.EndEtg_cep.value;
    	f.Verifica_Justificativa.value = f.EndEtg_obs.value;
    }

    trataProdutorRuralEndEtg_PF(null);
    trocarEndEtgTipoPessoa(null);
});

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
	f=fPED;
	ProcessaSelecaoCEP=TrataCepEnderecoEntrega;
	strUrl="../Global/AjaxCepPesqPopup.asp";
	if (trim(f.EndEtg_cep.value)!="") strUrl=strUrl+"?CepDefault="+trim(f.EndEtg_cep.value);
	fCepPopup=window.open(strUrl, "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=980,height=650,left=0,top=0");
	fCepPopup.focus();
}

function TrataCepEnderecoEntrega(strCep, strUF, strLocalidade, strBairro, strLogradouro, strEnderecoNumero, strEnderecoComplemento) {
var f;
	f=fPED;
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

function processaFormaPagtoDefault() {
var f, i;
	f=fPED;

	if (nivelEdicaoFormaPagto == COD_NIVEL_EDICAO_BLOQUEADA) return;

	// Vers�o antiga da forma de pagamento?
	if (f.tipo_parcelamento.value=="0") return;
		
//  O pedido foi cadastrado j� com a nova pol�tica de custo financeiro por fornecedor?
	if (f.c_custoFinancFornecTipoParcelamento.value=="") return;
	
	for (i=0; i<fPED.rb_forma_pagto.length; i++) {
		if (fPED.rb_forma_pagto[i].checked) {
			fPED.rb_forma_pagto[i].click();
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
	f=fPED;

	if (nivelEdicaoFormaPagto == COD_NIVEL_EDICAO_BLOQUEADA) return;

//  O pedido foi cadastrado j� com a nova pol�tica de custo financeiro por fornecedor?
	if (f.c_custoFinancFornecTipoParcelamento.value=="") return;

	strMsgErroAlert="";
	if (objAjaxCustoFinancFornecConsultaPreco.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxCustoFinancFornecConsultaPreco.responseText;
		if (strResp=="") {
			alert("Falha ao consultar o pre�o!!");
			window.status="Conclu�do";
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
					//  Pre�o
						oNodes=xmlDoc.getElementsByTagName("precoLista")[i];
						if (oNodes.childNodes.length > 0) strPrecoLista=oNodes.childNodes[0].nodeValue; else strPrecoLista="";
						if (strPrecoLista==null) strPrecoLista="";
					//  Atualiza o pre�o
						if (strPrecoLista=="") {
							alert("Falha na consulta do pre�o do produto " + strProduto + "!!\n" + strMsgErro);
							}
						else {
							for (j=0; j<f.c_fabricante.length; j++) {
								if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								//  Percorre o la�o at� o final para o caso do usu�rio ter digitado o mesmo produto em v�rias linhas
								//	(apesar de que isso n�o ser� aceito pelas consist�ncias que ser�o feitas).
									f.c_preco_lista[j].value=strPrecoLista;
									f.c_preco_lista[j].style.color="black";
									}
								}
							}
						}
					else {
					//  C�digo do Erro
						oNodes=xmlDoc.getElementsByTagName("codigo_erro")[i];
						if (oNodes.childNodes.length > 0) strCodigoErro=oNodes.childNodes[0].nodeValue; else strCodigoErro="";
						if (strCodigoErro==null) strCodigoErro="";
					//  Mensagem de Erro
						oNodes=xmlDoc.getElementsByTagName("msg_erro")[i];
						if (oNodes.childNodes.length > 0) strMsgErro=oNodes.childNodes[0].nodeValue; else strMsgErro="";
						if (strMsgErro==null) strMsgErro="";
						for (j=0; j<f.c_fabricante.length; j++) {
						//  Percorre o la�o at� o final para o caso do usu�rio ter digitado o mesmo produto em v�rias linhas
						//	(apesar de que isso n�o ser� aceito pelas consist�ncias que ser�o feitas).
							if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								f.c_preco_lista[j].style.color=COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE;
								}
							}
						if (strMsgErroAlert!="") strMsgErroAlert+="\n\n";
						strMsgErroAlert+="Falha ao consultar o pre�o do produto " + strProduto + "!!\n" + strMsgErro;
						}
					}
				}
			catch (e)
				{
				alert("Falha na consulta do pre�o!!\n"+e.message);
				}
			}
		
		if (strMsgErroAlert!="") alert(strMsgErroAlert);
		
		recalcula_total_todas_linhas(); 
		recalcula_RA();
		recalcula_RA_Liquido();
			
		window.status="Conclu�do";
		$("#divAjaxRunning").hide();
		}
}

function recalculaCustoFinanceiroPrecoLista() {
var f, i, strListaProdutos, strUrl, strOpcaoFormaPagto;
	f=fPED;

	if (nivelEdicaoFormaPagto == COD_NIVEL_EDICAO_BLOQUEADA) return;

//  O pedido foi cadastrado j� com a nova pol�tica de custo financeiro por fornecedor?
	if (f.c_custoFinancFornecTipoParcelamento.value=="") return;

	objAjaxCustoFinancFornecConsultaPreco=GetXmlHttpObject();
	if (objAjaxCustoFinancFornecConsultaPreco==null) {
		alert("O browser N�O possui suporte ao AJAX!!");
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
	
//  Converte as op��es de forma de pagamento do pedido em uma op��o que possa tratada pela tabela de custo financeiro
	strOpcaoFormaPagto="";
	for (i=0; i<fPED.rb_forma_pagto.length; i++) {
		if (fPED.rb_forma_pagto[i].checked) {
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

//  N�o consulta novamente se for a mesma consulta anterior
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

	window.status="Aguarde, consultando pre�os ...";
	$("#divAjaxRunning").show();
	
	strUrl = "../Global/AjaxCustoFinancFornecConsultaPrecoBD.asp";
	strUrl+="?tipoParcelamento="+f.c_custoFinancFornecTipoParcelamento.value;
	strUrl+="&qtdeParcelas="+f.c_custoFinancFornecQtdeParcelas.value;
	strUrl += "&loja=" + f.c_loja.value;
	strUrl += "&pedido=" + f.pedido_selecionado.value;
	strUrl+="&listaProdutos="+strListaProdutos;
//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxCustoFinancFornecConsultaPreco.onreadystatechange=trataRespostaAjaxCustoFinancFornecSincronizaPrecos;
	objAjaxCustoFinancFornecConsultaPreco.open("GET",strUrl,true);
	objAjaxCustoFinancFornecConsultaPreco.send(null);
}

function executa_consulta_senha_desconto(id_cliente, loja) {
	var postData = "id_cliente=" + id_cliente + "&loja=" + loja;
	// Prevents server from using a cached file
	var url = "../Global/JsonConsultaSenhaDescontoBD.asp" + "?anticache=" + Math.random() + Math.random();
	window.status = "Consultando banco de dados...";
	var responseText = synchronous_ajax(url, postData);
	objSenhaDesconto = eval("(" + responseText + ")");
	window.status = "Conclu�do";
}

function isFormaPagtoEditada(f) {
var idx;
var s_forma_pagto = "";

	if (nivelEdicaoFormaPagto == COD_NIVEL_EDICAO_BLOQUEADA) return false;
	
	idx = -1;
	
	//	� Vista
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_forma_pagto += trim(f.rb_forma_pagto[idx].value) +
						"|" + trim(f.op_av_forma_pagto.value);
	}
	
	//	Parcela �nica
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_forma_pagto += trim(f.rb_forma_pagto[idx].value) +
						"|" + trim(f.op_pu_forma_pagto.value) +
						"|" + trim(f.c_pu_valor.value) +
						"|" + trim(f.c_pu_vencto_apos.value);
	}

	//	Parcelado no Cart�o (internet)
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_forma_pagto += trim(f.rb_forma_pagto[idx].value) +
						"|" + trim(f.c_pc_qtde.value) +
						"|" + trim(f.c_pc_valor.value);
	}
	
	//	Parcelado no Cart�o (maquineta)
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_forma_pagto += trim(f.rb_forma_pagto[idx].value) +
						"|" + trim(f.c_pc_maquineta_qtde.value) +
						"|" + trim(f.c_pc_maquineta_valor.value);
	}

	//	Parcelado Com Entrada
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_forma_pagto += trim(f.rb_forma_pagto[idx].value) +
						"|" + trim(f.op_pce_entrada_forma_pagto.value) +
						"|" + trim(f.c_pce_entrada_valor.value) +
						"|" + trim(f.op_pce_prestacao_forma_pagto.value) +
						"|" + trim(f.c_pce_prestacao_qtde.value) +
						"|" + trim(f.c_pce_prestacao_valor.value) +
						"|" + trim(f.c_pce_prestacao_periodo.value);
	}

	//	Parcelado Sem Entrada
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_forma_pagto += trim(f.rb_forma_pagto[idx].value) +
						"|" + trim(f.op_pse_prim_prest_forma_pagto.value) +
						"|" + trim(f.c_pse_prim_prest_valor.value) +
						"|" + trim(f.c_pse_prim_prest_apos.value) +
						"|" + trim(f.op_pse_demais_prest_forma_pagto.value) +
						"|" + trim(f.c_pse_demais_prest_qtde.value) +
						"|" + trim(f.c_pse_demais_prest_valor.value) +
						"|" + trim(f.c_pse_demais_prest_periodo.value);
	}

	if (formaPagamentoOriginal != s_forma_pagto) return true;

	return false;
}

function obtem_perc_comissao_e_desconto_a_utilizar(f, vl_total_pedido, perc_comissao_e_desconto_nivel1, perc_comissao_e_desconto_nivel1_pj, perc_comissao_e_desconto_nivel2, perc_comissao_e_desconto_nivel2_pj) {
var i, idx, s_pg, blnPreferencial;
var vlNivel1 = 0;
var vlNivel2 = 0;

	if (nivelEdicaoFormaPagto == COD_NIVEL_EDICAO_BLOQUEADA) {
		return perc_comissao_e_desconto_n1_n2_a_utilizar;
	}

	// ANALISA QUAL � O MEIO DE PAGAMENTO PREDOMINANTE
	idx = -1;
	//	� Vista
	//	=======
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_pg = trim(f.op_av_forma_pagto.value);
		if (s_pg == '') return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado � um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);
		}
		//	O meio de pagamento n�o � preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}

	//	Parcela �nica
	//	=============
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_pg = trim(f.op_pu_forma_pagto.value);
		if (s_pg == '') return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado � um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);
		}
		//	O meio de pagamento n�o � preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}

	//	Parcelado no Cart�o (internet)
	//	==============================
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_pg = ID_FORMA_PAGTO_CARTAO;
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado � um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);
		}
		//	O meio de pagamento n�o � preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}

	//	Parcelado no Cart�o (maquineta)
	//	===============================
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		s_pg = ID_FORMA_PAGTO_CARTAO_MAQUINETA;
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado � um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);
		}
		//	O meio de pagamento n�o � preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}

	//	Parcelado Com Entrada
	//	=====================
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		//	Identifica e contabiliza o valor da entrada
		blnPreferencial = false;
		s_pg = trim(f.op_pce_entrada_forma_pagto.value);
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado � um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) {
				blnPreferencial = true;
				break;
			}
		}

		if (blnPreferencial) {
			vlNivel2 = converte_numero(trim(f.c_pce_entrada_valor.value));
		}
		else {
			vlNivel1 = converte_numero(trim(f.c_pce_entrada_valor.value));
		}

		//	Identifica e contabiliza o valor das parcelas
		blnPreferencial = false;
		s_pg = trim(f.op_pce_prestacao_forma_pagto.value);
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado � um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) {
				blnPreferencial = true;
				break;
			}
		}

		if (blnPreferencial) {
			vlNivel2 += converte_numero(f.c_pce_prestacao_qtde.value) * converte_numero(f.c_pce_prestacao_valor.value);
		}
		else {
			vlNivel1 += converte_numero(f.c_pce_prestacao_qtde.value) * converte_numero(f.c_pce_prestacao_valor.value);
		}

		//	O montante a pagar por meio de pagamento preferencial � maior que 50% do total?
		if (vlNivel2 > (vl_total_pedido / 2)) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);

		//	O meio de pagamento n�o � preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}

	//	Parcelado Sem Entrada
	//	=====================
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		//	Identifica e contabiliza o valor da 1� parcela
		blnPreferencial = false;
		s_pg = trim(f.op_pse_prim_prest_forma_pagto.value);
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado � um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) {
				blnPreferencial = true;
				break;
			}
		}

		if (blnPreferencial) {
			vlNivel2 = converte_numero(trim(f.c_pse_prim_prest_valor.value));
		}
		else {
			vlNivel1 = converte_numero(trim(f.c_pse_prim_prest_valor.value));
		}

		//	Identifica e contabiliza o valor das parcelas
		blnPreferencial = false;
		s_pg = trim(f.op_pse_demais_prest_forma_pagto.value);
		for (i = 0; i < vMPN2.length; i++) {
			//	O meio de pagamento selecionado � um dos preferenciais
			if (parseInt(s_pg) == parseInt(vMPN2[i])) {
				blnPreferencial = true;
				break;
			}
		}

		if (blnPreferencial) {
			vlNivel2 += converte_numero(trim(f.c_pse_demais_prest_qtde.value)) * converte_numero(trim(f.c_pse_demais_prest_valor.value));
		}
		else {
			vlNivel1 += converte_numero(trim(f.c_pse_demais_prest_qtde.value)) * converte_numero(trim(f.c_pse_demais_prest_valor.value));
		}

		//	O montante a pagar por meio de pagamento preferencial � maior que 50% do total?
		if (vlNivel2 > (vl_total_pedido / 2)) return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel2_pj : perc_comissao_e_desconto_nivel2);

		//	O meio de pagamento n�o � preferencial
		return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
	}

	//	O meio de pagamento n�o � preferencial
	return (f.c_tipo_cliente.value == ID_PJ ? perc_comissao_e_desconto_nivel1_pj : perc_comissao_e_desconto_nivel1);
}

function calcula_vl_total_preco_venda(f) {
var mTotVenda;
	mTotVenda = 0;
	for (i = 0; i < f.c_qtde.length; i++) mTotVenda = mTotVenda + converte_numero(f.c_qtde[i].value) * converte_numero(f.c_vl_unitario[i].value);
	return mTotVenda;
}

// RETORNA O VALOR TOTAL DO PEDIDO A SER USADO P/ CALCULAR A FORMA DE PAGAMENTO
function fp_vl_total_pedido( ) {
var f,i,mTotVenda,mTotNFBase,mTotNF;
	f=fPED;
	mTotNFBase = converte_numero(f.c_total_NF_base.value);
	mTotNFBase = mTotNFBase - converte_numero(f.c_total_devolucoes_NF.value);
	mTotVenda=0;
	for (i=0; i<f.c_qtde.length; i++) mTotVenda=mTotVenda+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_unitario[i].value);
	mTotNF=0;
	for (i=0; i<f.c_qtde.length; i++) mTotNF=mTotNF+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_NF[i].value);
//  Retorna total de pre�o NF (tem valor de NF, ou seja, pedido c/ RA)?
	if (mTotNF > 0) {
		return mTotNFBase+mTotNF;
		}
//  Retorna total de pre�o de venda
	else {
		return mTotNFBase+mTotVenda;
		}
}

// PARCELA �NICA
function pu_atualiza_valor( ){
var f,vt;
	f=fPED;
	if (converte_numero(trim(f.c_pu_valor.value))>0) return;
	vt=fp_vl_total_pedido();
	f.c_pu_valor.value=formata_moeda(vt);
}

// PARCELADO NO CART�O (INTERNET)
function pc_calcula_valor_parcela( ){
var f,n,t;
	f=fPED;
	if (trim(f.c_pc_qtde.value)=='') return;
	n=converte_numero(f.c_pc_qtde.value);
	if (n<=0) return;
	t=fp_vl_total_pedido();
	p=t/n;
	f.c_pc_valor.value=formata_moeda(p);
}

// PARCELADO NO CART�O (MAQUINETA)
function pc_maquineta_calcula_valor_parcela( ){
	var f,n,t;
	f=fPED;
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
	f=fPED;
	if (converte_numero(trim(f.c_pce_prestacao_periodo.value))>0) return;
	f.c_pce_prestacao_periodo.value='30';
}

function pce_calcula_valor_parcela( ){
var f,n,e,t;
	f=fPED;
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
	f=fPED;
	if (converte_numero(trim(f.c_pse_demais_prest_periodo.value))>0) return;
	f.c_pse_demais_prest_periodo.value='30';
}

function pse_calcula_valor_parcela( ){
var f,n,e,t;
	f=fPED;
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
	f=fPED;
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
	f=fPED;
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

function recalcula_RA( ) {
var f,i,mTotVenda,mTotNF,mTotRABase,vl_RA;
	f=fPED;
	mTotVenda=0;
	mTotRABase=converte_numero(f.c_total_RA_base.value);
	for (i=0; i<f.c_vl_total.length; i++) mTotVenda=mTotVenda+converte_numero(f.c_vl_total[i].value);
	mTotNF=0;
	for (i=0; i<f.c_qtde.length; i++) mTotNF=mTotNF+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_NF[i].value);
	f.c_total_NF.value = formata_moeda(mTotNF);
	vl_RA=mTotRABase+(mTotNF-mTotVenda);
	f.c_total_RA.value = formata_moeda(vl_RA);
	if (vl_RA>=0) f.c_total_RA.style.color="green"; else f.c_total_RA.style.color="red";
}

function recalcula_RA_Liquido( ) {
var f,i,mTotVenda,mTotNF,mTotRABase,vl_RA,vl_RA_liquido;
var r_RA_liquido;
	f=fPED;

	recalcula_total_todas_linhas();
	
	mTotVenda=0;
	mTotRABase=converte_numero(f.c_total_RA_base.value);
	for (i=0; i<f.c_vl_total.length; i++) mTotVenda=mTotVenda+converte_numero(f.c_vl_total[i].value);
	mTotNF=0;
	for (i=0; i<f.c_qtde.length; i++) mTotNF=mTotNF+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_NF[i].value);
	vl_RA=mTotRABase+(mTotNF-mTotVenda);

	r_RA_liquido = new calcula_total_RA_liquido(PERC_DESAGIO_RA_LIQUIDA_PEDIDO, vl_RA);
	vl_RA_liquido = r_RA_liquido.vl_total_RA_liquido;
	f.c_total_RA_Liquido.value = formata_moeda(vl_RA_liquido);
	if (vl_RA_liquido>=0) f.c_total_RA_Liquido.style.color="green"; else f.c_total_RA_Liquido.style.color="red";
}

function calcula_desconto(idx) {
	var f, s, i, m, d, m_lista, m_unit;
	f = fPED;
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

function atualiza_itens_com_desc_linear() {
	var f;
	f = fPED;
	if (trim(f.c_desc_linear.value) == "") return;
	f.c_desc_linear.value = formata_perc_desc_linear(f.c_desc_linear.value);
	if (trim(f.c_desc_linear.value) == "") return;
	for (i = 0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value) != "") {
			f.c_desc[i].value = f.c_desc_linear.value;
			calcula_desconto(i);
			if (!blnTemRA) f.c_vl_NF[i].value = f.c_vl_unitario[i].value;
		}
	}
	recalcula_total_todas_linhas();
	recalcula_RA();
	recalcula_RA_Liquido();
}

function recalcula_total_linha( id ) {
var idx, m, m_lista, m_unit, d, f, i, s;
	f=fPED;
	idx=parseInt(id)-1;
	if (f.c_produto[idx].value=="") return;
	m_lista=converte_numero(f.c_preco_lista[idx].value);
	m_unit=converte_numero(f.c_vl_unitario[idx].value);
	if (m_lista==0) d=0; else d=100*(m_lista-m_unit)/m_lista;
	if (d == 0) s = ""; else s = formata_perc_desconto(d);
	if (f.c_desc[idx].value!=s) f.c_desc[idx].value=s;
	s=formata_moeda(parseInt(f.c_qtde[idx].value)*m_unit);
	if (f.c_vl_total[idx].value!=s) f.c_vl_total[idx].value=s;
	m=0;
	for (i=0; i<f.c_vl_total.length; i++) m=m+converte_numero(f.c_vl_total[i].value);
	s=formata_moeda(m);
	if (f.c_total_geral.value!=s) f.c_total_geral.value=s;
	f.c_desc_medio_total.value = formata_perc_desc_linear(calcula_desconto_medio());
}

function recalcula_total_todas_linhas() {
var f,i,t,m_lista,m_unit,d,m,s;
	f = fPED;
	t=0;
	for (i=0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value)!="") {
			m_lista=converte_numero(f.c_preco_lista[i].value);
			m_unit=converte_numero(f.c_vl_unitario[i].value);
			if (m_lista==0) d=0; else d=100*(m_lista-m_unit)/m_lista;
			if (d == 0) s = ""; else s = formata_perc_desconto(d);
			if (f.c_desc[i].value!=s) f.c_desc[i].value=s;
			m=parseInt(f.c_qtde[i].value)*m_unit;
			f.c_vl_total[i].value=formata_moeda(m);
			t=t+m;
			}
		}
	f.c_total_geral.value=formata_moeda(t);
	f.c_desc_medio_total.value = formata_perc_desc_linear(calcula_desconto_medio());
}

function preenche_sugestao_forma_pagto( ) {
var f, n, t, p, s;
	f=fPED;
	n=converte_numero(f.c_qtde_parcelas.value);
	t=converte_numero(f.c_total_geral.value);
	if (n > 0) {
		p=t/n;
		s = "Pagamento em " + n;
		if (n==1) s = s + " parcela de "; else s = s + " parcelas de ";
		s = s + SIMBOLO_MONETARIO + " " + formata_moeda(p);
		f.c_forma_pagto.value=s;
		}
	else f.c_forma_pagto.value="";
}

function consiste_forma_pagto( blnComAvisos ) {
var f,idx,vtNF,vtFP,ve,ni,nip,n,vp;
var MAX_ERRO_ARREDONDAMENTO = 0.1;
	f = fPED;
	vtNF=fp_vl_total_pedido();
	vtFP=0;
	idx=-1;
	
//	� Vista
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

//	Parcela �nica
	idx++;
	if (f.rb_forma_pagto[idx].checked) {
		if (trim(f.op_pu_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento da parcela �nica!!');
				f.op_pu_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pu_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da parcela �nica!!');
				f.c_pu_valor.focus();
				}
			return false;
			}
		ve=converte_numero(f.c_pu_valor.value);
		vtFP=ve;
		if (ve<=0) {
			if (blnComAvisos) {
				alert('Valor da parcela �nica � inv�lido!!');
				f.c_pu_valor.focus();
				}
			return false;
			}
		if (trim(f.c_pu_vencto_apos.value)=='') {
			if (blnComAvisos) {
				alert('Indique o intervalo de vencimento da parcela �nica!!');
				f.c_pu_vencto_apos.focus();
				}
			return false;
			}
		nip=converte_numero(f.c_pu_vencto_apos.value);
		if (nip<=0) {
			if (blnComAvisos) {
				alert('Intervalo de vencimento da parcela �nica � inv�lido!!');
				f.c_pu_vencto_apos.focus();
				}
			return false;
			}
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
				alert('H� diverg�ncia entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito atrav�s da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
				f.c_pu_valor.focus();
				}
			return false;
			}
		return true;
		}

//	Parcelado no cart�o (internet)
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
				alert('Quantidade de parcelas inv�lida!!');
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
				alert('Valor de parcela inv�lido!!');
				f.c_pc_valor.focus();
				}
			return false;
			}
		vtFP=n*vp;
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
				alert('H� diverg�ncia entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito atrav�s da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
				f.c_pc_valor.focus();
				}
			return false;
			}
		return true;
		}

	//	Parcelado no cart�o (maquineta)
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
				alert('Quantidade de parcelas inv�lida!!');
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
				alert('Valor de parcela inv�lido!!');
				f.c_pc_maquineta_valor.focus();
			}
			return false;
		}
		vtFP=n*vp;
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
				alert('H� diverg�ncia entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito atrav�s da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
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
				alert('Valor da entrada inv�lido!!');
				f.c_pce_entrada_valor.focus();
				}
			return false;
			}
		if (trim(f.op_pce_prestacao_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento das presta��es!!');
				f.op_pce_prestacao_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pce_prestacao_qtde.value)=='') {
			if (blnComAvisos) {
				alert('Indique a quantidade de presta��es!!');
				f.c_pce_prestacao_qtde.focus();
				}
			return false;
			}
		n=converte_numero(f.c_pce_prestacao_qtde.value);
		if (n<=0) {
			if (blnComAvisos) {
				alert('Quantidade de presta��es inv�lida!!');
				f.c_pce_prestacao_qtde.focus();
				}
			return false;
			}
		if (trim(f.c_pce_prestacao_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da presta��o!!');
				f.c_pce_prestacao_valor.focus();
				}
			return false;
			}
		vp=converte_numero(f.c_pce_prestacao_valor.value);
		if (vp<=0) {
			if (blnComAvisos) {
				alert('Valor de presta��o inv�lido!!');
				f.c_pce_prestacao_valor.focus();
				}
			return false;
			}
		vtFP=ve+(n*vp);
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
				alert('H� diverg�ncia entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito atrav�s da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
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
				alert('Intervalo de vencimento inv�lido!!');
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
				alert('Indique a forma de pagamento da 1� presta��o!!');
				f.op_pse_prim_prest_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pse_prim_prest_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor da 1� presta��o!!');
				f.c_pse_prim_prest_valor.focus();
				}
			return false;
			}
		ve=converte_numero(f.c_pse_prim_prest_valor.value);
		if (ve<=0) {
			if (blnComAvisos) {
				alert('Valor da 1� presta��o inv�lido!!');
				f.c_pse_prim_prest_valor.focus();
				}
			return false;
			}
		if (trim(f.c_pse_prim_prest_apos.value)=='') {
			if (blnComAvisos) {
				alert('Indique o intervalo de vencimento da 1� parcela!!');
				f.c_pse_prim_prest_apos.focus();
				}
			return false;
			}
		nip=converte_numero(f.c_pse_prim_prest_apos.value);
		if (nip<=0) {
			if (blnComAvisos) {
				alert('Intervalo de vencimento da 1� parcela � inv�lido!!');
				f.c_pse_prim_prest_apos.focus();
				}
			return false;
			}
		if (trim(f.op_pse_demais_prest_forma_pagto.value)=='') {
			if (blnComAvisos) {
				alert('Indique a forma de pagamento das demais presta��es!!');
				f.op_pse_demais_prest_forma_pagto.focus();
				}
			return false;
			}
		if (trim(f.c_pse_demais_prest_qtde.value)=='') {
			if (blnComAvisos) {
				alert('Indique a quantidade das demais presta��es!!');
				f.c_pse_demais_prest_qtde.focus();
				}
			return false;
			}
		n=converte_numero(f.c_pse_demais_prest_qtde.value);
		if (n<=0) {
			if (blnComAvisos) {
				alert('Quantidade de presta��es inv�lida!!');
				f.c_pse_demais_prest_qtde.focus();
				}
			return false;
			}
		if (trim(f.c_pse_demais_prest_valor.value)=='') {
			if (blnComAvisos) {
				alert('Indique o valor das demais presta��es!!');
				f.c_pse_demais_prest_valor.focus();
				}
			return false;
			}
		vp=converte_numero(f.c_pse_demais_prest_valor.value);
		if (vp<=0) {
			if (blnComAvisos) {
				alert('Valor de presta��o inv�lido!!');
				f.c_pse_demais_prest_valor.focus();
				}
			return false;
			}
		vtFP=ve+(n*vp);
		if (Math.abs(vtFP-vtNF)>MAX_ERRO_ARREDONDAMENTO) {
			if (blnComAvisos) {
				alert('H� diverg�ncia entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito atrav�s da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
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
				alert('Intervalo de vencimento inv�lido!!');
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

function LimparCamposEndEtg( f ) {
	f.EndEtg_endereco.value="";
	f.EndEtg_endereco_numero.value="";
	f.EndEtg_endereco_complemento.value="";
	f.EndEtg_bairro.value="";
	f.EndEtg_cidade.value="";
	f.EndEtg_uf.value="";
	f.EndEtg_cep.value="";
    f.EndEtg_obs.selectedIndex = 0;

    <%if blnUsarMemorizacaoCompletaEnderecos then %>
        f.EndEtg_email.value = "";
	    f.EndEtg_email_xml.value = "";
    <% end if%>

    <%if blnUsarMemorizacaoCompletaEnderecos and not eh_cpf then %>
        f.EndEtg_tipo_pessoa[0].checked = false;
        f.EndEtg_tipo_pessoa[1].checked = false;
	    f.EndEtg_cnpj_cpf_PJ.value="";
        f.EndEtg_ie_PJ.value="";
        f.EndEtg_contribuinte_icms_status_PJ[0].checked = false;
        f.EndEtg_contribuinte_icms_status_PJ[1].checked = false;
        f.EndEtg_contribuinte_icms_status_PJ[2].checked = false;

        f.EndEtg_cnpj_cpf_PF.value="";
        f.EndEtg_produtor_rural_status_PF[0].checked = false;
        f.EndEtg_produtor_rural_status_PF[1].checked = false;
        f.EndEtg_ie_PF.value="";
        f.EndEtg_contribuinte_icms_status_PF[0].checked = false;
        f.EndEtg_contribuinte_icms_status_PF[1].checked = false;
        f.EndEtg_contribuinte_icms_status_PF[2].checked = false;

        f.EndEtg_nome.value="";
        f.EndEtg_ddd_res.value="";
        f.EndEtg_tel_res.value="";
        f.EndEtg_ddd_cel.value="";
        f.EndEtg_tel_cel.value="";
        f.EndEtg_ddd_com.value="";
        f.EndEtg_tel_com.value="";
        f.EndEtg_ramal_com.value="";
        f.EndEtg_ddd_com_2.value="";
        f.EndEtg_tel_com_2.value="";
        f.EndEtg_ramal_com_2.value = "";

        trataProdutorRuralEndEtg_PF(null);
        trocarEndEtgTipoPessoa(null);
    <%end if%>
}

function calcula_desconto_medio() {
	var f, i, vl_total_preco_lista, vl_total_preco_venda, perc_desc_medio;

	f = fPED;
	vl_total_preco_lista = 0;
	vl_total_preco_venda = 0;

	// La�o p/ produtos
	for (i = 0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value) != "") {
			vl_total_preco_lista += converte_numero(f.c_qtde[i].value) * converte_numero(f.c_preco_lista[i].value);
			vl_total_preco_venda += converte_numero(f.c_qtde[i].value) * converte_numero(f.c_vl_unitario[i].value);
		}
	}

	if (vl_total_preco_lista == 0) {
		perc_desc_medio = 0;
	}
	else {
		perc_desc_medio = 100 * (vl_total_preco_lista - vl_total_preco_venda) / vl_total_preco_lista;
	}
	return perc_desc_medio;
}

function trata_edicao_RA(index) {
var f;
	f = fPED;
	if ((f.c_permite_RA_status.value != '1') && (f.c_st_violado_permite_RA_status.value == '0')) f.c_vl_NF[index].value = f.c_vl_unitario[index].value;
}

function fPEDConfirma( f ) {
var s, blnTemEndEntrega, blnHouveEdicaoVlUnitario, blnHouveEdicaoVlUnitarioComToleranciaArred, strMsgErro;
var i, j, vl_preco_lista, vl_preco_venda, perc_desc;
var perc_RT, perc_RT_novo, perc_max_RT_padrao, perc_max_comissao_e_desconto, perc_max_comissao_e_desconto_pj, perc_max_comissao_e_desconto_nivel2, perc_max_comissao_e_desconto_nivel2_pj, perc_senha_desconto, perc_desc_medio;
var perc_max_RT_a_utilizar, perc_max_comissao_e_desconto_a_utilizar;
var perc_max_desc_alcada_1_pf, perc_max_desc_alcada_1_pj, perc_max_desc_alcada_2_pf, perc_max_desc_alcada_2_pj, perc_max_desc_alcada_3_pf, perc_max_desc_alcada_3_pj;
var perc_max_comissao_alcada1, perc_max_comissao_alcada2, perc_max_comissao_alcada3;
var blnFormaPagtoEditada;
var NUMERO_LOJA_ECOMMERCE_AR_CLUBE = "<%=NUMERO_LOJA_ECOMMERCE_AR_CLUBE%>";

	recalcula_total_todas_linhas();

	if (f.c_loja.value != NUMERO_LOJA_ECOMMERCE_AR_CLUBE) {
	    if (f.c_indicador.value == "") {
	        if(f.c_perc_RT.value != "") {
	            if (parseFloat(f.c_perc_RT.value.replace(',','.')) > 0) {
	                alert('N�o � poss�vel gravar o pedido com o campo "Indicador" vazio e "COM(%)" maior do que zero!!');
	                f.c_perc_RT.focus();
	                return;
	            }
	        }	        
	    }
	}

	s = "" + f.c_obs1.value;
	if (s.length > MAX_TAM_OBS1) {
		alert('Conte�do de "Observa��es " excede em ' + (s.length-MAX_TAM_OBS1) + ' caracteres o tamanho m�ximo de ' + MAX_TAM_OBS1 + '!!');
		f.c_obs1.focus();
		return;
	}

	s = "" + f.c_nf_texto.value;
	if (s.length > MAX_TAM_NF_TEXTO) {
	    alert('Conte�do de "Constar na NF" excede em ' + (s.length-MAX_TAM_NF_TEXTO) + ' caracteres o tamanho m�ximo de ' + MAX_TAM_NF_TEXTO + '!!');
	    f.c_nf_texto.focus();
	    return;
	}

	s = "" + f.c_forma_pagto.value;
	if (s.length > MAX_TAM_FORMA_PAGTO) {
		alert('Conte�do de "Forma de Pagamento" excede em ' + (s.length-MAX_TAM_FORMA_PAGTO) + ' caracteres o tamanho m�ximo de ' + MAX_TAM_FORMA_PAGTO + '!!');
		f.c_forma_pagto.focus();
		return;
		}

//  Consiste a nova vers�o da forma de pagamento
	if (f.versao_forma_pagamento.value == '2') {
		if (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL) {
			if (!consiste_forma_pagto(true)) return;
			}
		}

	recalcula_RA();
	recalcula_RA_Liquido();

	if (blnConfirmaDifRAeValores) {
		if (f.c_total_RA.value != f.c_total_RA_original.value) {
			if (!confirm("O valor do RA � de " + SIMBOLO_MONETARIO + " " + formata_moeda(converte_numero(f.c_total_RA.value))+"\nContinua?")) return;
			}
		}

	if (f.blnEndEntregaEdicaoLiberada.value == "<%=Cstr(True)%>") {
		    blnTemEndEntrega=false;
		if (trim(f.EndEtg_endereco.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_endereco_numero.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_endereco_complemento.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_bairro.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_cidade.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_uf.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_cep.value)!="") blnTemEndEntrega=true;
        if (trim(f.EndEtg_obs.value) != "") blnTemEndEntrega = true;

<%if blnUsarMemorizacaoCompletaEnderecos then %>
        if (trim(f.EndEtg_email.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_email_xml.value) != "") blnTemEndEntrega = true;
<% end if%>

<%if blnUsarMemorizacaoCompletaEnderecos and not eh_cpf then %>

        if( $('input[name="EndEtg_tipo_pessoa"]:checked').val()) blnTemEndEntrega = true;

        //simplesmente testamos todos os campos, qualquer valor em qq campo significa preenchimento
        //n�o deve estar em campo oculto porque o usu�rio deve clicar no X para limpar, e o X limpa todos os campos, inclusive os n�o visiveis no momento

        //pj
        if (trim(f.EndEtg_cnpj_cpf_PJ.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_ie_PJ.value) != "") blnTemEndEntrega = true;
        if( $('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').val()) blnTemEndEntrega = true;

        //pf
        if (trim(f.EndEtg_cnpj_cpf_PF.value) != "") blnTemEndEntrega = true;
        if( $('input[name="EndEtg_produtor_rural_status_PF"]:checked').val()) blnTemEndEntrega = true;
        if (trim(f.EndEtg_ie_PF.value) != "") blnTemEndEntrega = true;
        if( $('input[name="EndEtg_contribuinte_icms_status_PF"]:checked').val()) blnTemEndEntrega = true;

        //ambos
        if (trim(f.EndEtg_nome.value) != "") blnTemEndEntrega = true;

        //pj
        if (trim(f.EndEtg_ddd_com.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_tel_com.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_ramal_com.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_ddd_com_2.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_tel_com_2.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_ramal_com_2.value) != "") blnTemEndEntrega = true;

        //pf
        if (trim(f.EndEtg_ddd_res.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_tel_res.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_ddd_cel.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_tel_cel.value) != "") blnTemEndEntrega = true;
<%end if%>



	<%if r_pedido.st_memorizacao_completa_enderecos <> 0 then %>

		if (trim(f.endereco__endereco.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__endereco.focus();
            return;
        }
        if (trim(f.endereco__bairro.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__bairro.focus();
            return;
        }

        if (trim(f.endereco__numero.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__numero.focus();
            return;
        }
        if (trim(f.endereco__cidade.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__cidade.focus();
            return;
        }

        if (trim(f.endereco__uf.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__uf.focus();
            return;
        }

        if (trim(f.endereco__cep.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__cep.focus();
            return;
        }

        if ((trim(f.cliente__email.value) != "") && (!email_ok(f.cliente__email.value))) {
            alert('E-mail inv�lido!!');
            f.cliente__email.focus();
            return;
        }

        if ((trim(f.cliente__email_xml.value) != "") && (!email_ok(f.cliente__email_xml.value))) {
            alert('E-mail xml inv�lido!!');
            f.cliente__email_xml.focus();
            return;
        }


       <% if cliente__tipo = ID_PF then %>

		if ( (trim(f.cliente__ddd_res.value) != "" && !ddd_ok(f.cliente__ddd_res.value)) || (trim(f.cliente__ddd_res.value) == "" && trim(f.cliente__tel_res.value) != "") ) {
            alert('DDD inv�lido!!');
            f.cliente__ddd_res.focus();
            return;
        }

		if ( (trim(f.cliente__tel_res.value) != "" && !telefone_ok(f.cliente__tel_res.value)) || (trim(f.cliente__ddd_res.value) != "" && trim(f.cliente__tel_res.value) == "") ) {
            alert('Telefone residencial inv�lido!!');
            f.cliente__tel_res.focus();
            return;
        }

		if ( (trim(f.cliente__ddd_cel.value) != "" && !ddd_ok(f.cliente__ddd_cel.value)) || (trim(f.cliente__ddd_cel.value) == "" && trim(f.cliente__tel_cel.value) != "") ) {
            alert('Celular com DDD inv�lido!!');
            f.cliente__ddd_cel.focus();
            return;
        }

		if ( (trim(f.cliente__tel_cel.value) != "" && !telefone_ok(f.cliente__tel_cel.value)) || (trim(f.cliente__ddd_cel.value) != "" && trim(f.cliente__tel_cel.value) == "") ) {
            alert('Telefone celular inv�lido!!');
            f.cliente__tel_cel.focus();
            return;
        }


		if ( (trim(f.cliente__ddd_com.value) != "" && !ddd_ok(f.cliente__ddd_com.value)) || (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__tel_com.value) != "") ) {
            alert('DDD comercial inv�lido!!');
            f.cliente__ddd_com.focus();
            return;
        }

		if ( (trim(f.cliente__tel_com.value) != "" && !telefone_ok(f.cliente__tel_com.value)) || (trim(f.cliente__ddd_com.value) != "" && trim(f.cliente__tel_com.value) == "") ) {
            alert('Telefone comercial inv�lido!!');
            f.cliente__tel_com.focus();
            return;
        }

		if (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
            alert('DDD comercial inv�lido!!');
            f.cliente__ddd_com.focus();
            return;
        }

		if (trim(f.cliente__tel_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
            alert('Telefone comercial inv�lido!!');
            f.cliente__tel_com.focus();
            return;
        }

        if (trim(f.cliente__tel_res.value) == "" && trim(f.cliente__tel_cel.value) == "" && trim(f.cliente__tel_com.value) == "") {
            alert('Necess�rio preencher ao menos um telefone!!');
            f.cliente__ddd_cel.focus();
            return;
        }



        if (f.rb_produtor_rural[1].checked) {
            if (!f.rb_contribuinte_icms[1].checked) {
                alert('Para ser cadastrado como Produtor Rural, � necess�rio ser contribuinte do ICMS e possuir n� de IE!!');
                return;
            }
            if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
                alert('Informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
                return;
            }
            if ((f.rb_contribuinte_icms[1].checked) && (trim(f.cliente__ie.value) == "")) {
                alert('Se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
                f.cliente__ie.focus();
                return;
            }
            if ((f.rb_contribuinte_icms[0].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                alert('Se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                f.cliente__ie.focus();
                return;
            }
            if ((f.rb_contribuinte_icms[1].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                alert('Se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                f.cliente__ie.focus();
                return;
            }
            if (f.rb_contribuinte_icms[2].checked) {
                if (f.cliente__ie.value != "") {
                    alert("Se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
                    f.cliente__ie.focus();
                    return;
                }
            }
        }


		<% else %>

        if ((trim(f.cliente__email.value) != "") && (!email_ok(f.cliente__email.value))) {
            alert('E-mail inv�lido!!');
            f.cliente__email.focus();
            return;
        }

        if ((trim(f.cliente__email_xml.value) != "") && (!email_ok(f.cliente__email_xml.value))) {
            alert('E-mail (XML) inv�lido!!');
            f.cliente__email_xml.focus();
            return;
        }

           <% if CStr(r_pedido.loja) <> CStr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then %>
            // PARA CLIENTE PJ, � OBRIGAT�RIO O PREENCHIMENTO DO E-MAIL
            if ((trim(f.cliente__email.value) == "") && (trim(f.cliente__email_xml.value) == "")) {
                alert("� obrigat�rio informar um endere�o de e-mail");
                f.cliente__email.focus();
                return;
            }
            <% end if %>

        if ((f.rb_contribuinte_icms[1].checked) && (trim(f.cliente__ie.value) == "")) {
            alert('Se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
            f.cliente__ie.focus();
            return;
        }
        if ((f.rb_contribuinte_icms[0].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
            alert('Se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
            f.cliente__ie.focus();
            return;
        }
        if ((f.rb_contribuinte_icms[1].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
            alert('Se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
            f.cliente__ie.focus();
            return;
        }
        if (f.rb_contribuinte_icms[2].checked) {
            if (f.cliente__ie.value != "") {
                alert("Se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
                f.cliente__ie.focus();
                return;
            }
        }
		if ( (trim(f.cliente__ddd_com.value) != "" && !ddd_ok(f.cliente__ddd_com.value)) || (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__tel_com.value) != "") ) {
            alert('DDD comercial inv�lido!!');
            f.cliente__ddd_com.focus();
            return;
        }

		if (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
            alert('DDD comercial inv�lido!!');
            f.cliente__ddd_com.focus();
            return;
        }

		if ( (trim(f.cliente__tel_com.value) != "" && !telefone_ok(f.cliente__tel_com.value)) || (trim(f.cliente__ddd_com.value) != "" && trim(f.cliente__tel_com.value) == "") ) {
            alert('Telefone comercial inv�lido!!');
            f.cliente__tel_com.focus();
            return;
        }

		if (trim(f.cliente__tel_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
            alert('Telefone comercial inv�lido!!');
            f.cliente__tel_com.focus();
            return;
        }

		if ( (trim(f.cliente__ddd_com_2.value) != "" && !ddd_ok(f.cliente__ddd_com_2.value)) || (trim(f.cliente__ddd_com_2.value) == "" && trim(f.cliente__tel_com_2.value) != "") ) {
            alert('DDD comercial 2 inv�lido!!');
            f.cliente__ddd_com_2.focus();
            return;
        }

		if (trim(f.cliente__ddd_com_2.value) == "" && trim(f.cliente__ramal_com_2.value) != "") {
            alert('DDD comercial 2 inv�lido!!');
            f.cliente__ddd_com_2.focus();
            return;
        }

		if ( (trim(f.cliente__tel_com_2.value) != "" && !telefone_ok(f.cliente__tel_com_2.value)) || (trim(f.cliente__ddd_com_2.value) != "" && trim(f.cliente__tel_com_2.value) == "") ) {
            alert('Telefone comercial 2 inv�lido!!');
            f.cliente__tel_com_2.focus();
            return;
        }

		if (trim(f.cliente__tel_com_2.value) == "" && trim(f.cliente__ramal_com_2.value) != "") {
            alert('Telefone comercial 2 inv�lido!!');
            f.cliente__tel_com_2.focus();
            return;
        }

        if (trim(f.cliente__tel_com.value) == "" && trim(f.cliente__tel_com_2.value) == "") {
            alert('Necess�rio preencher ao menos um telefone!!');
            f.cliente__ddd_com.focus();
            return;
        }

		<% end if%>

		
<% end if%>


		if (blnTemEndEntrega) {
		    var blnEndEtg_obs
		    blnEndEtg_obs = false;
		    if ((f.EndEtg_endereco.value != f.Verifica_End_Entrega.value) || (f.EndEtg_endereco_numero.value != f.Verifica_num.value) || (f.EndEtg_cidade.value != f.Verifica_Cidade.value) || (f.EndEtg_uf.value != f.Verifica_UF.value) || (f.EndEtg_cep.value != f.Verifica_CEP.value) || (f.EndEtg_obs.value != f.Verifica_Justificativa.value)){
		        blnEndEtg_obs = true;
		    }
			if (trim(f.EndEtg_endereco.value)=="") {
				alert('Endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_endereco.focus();
				return;
				}

			if (trim(f.EndEtg_endereco_numero.value)=="") {
				alert('O n�mero do endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_endereco_numero.focus();
				return;
				}

			if (trim(f.EndEtg_bairro.value)=="") {
				alert('Bairro do endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_bairro.focus();
				return;
				}

			if (trim(f.EndEtg_cidade.value)=="") {
				alert('Cidade do endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_cidade.focus();
				return;
			    }
			if ((trim(f.EndEtg_obs.value)=="")  && blnEndEtg_obs == true) {
			    alert('Justificativa do endere�o de entrega n�o foi preenchido corretamente!!');
			    f.EndEtg_obs.focus();
			    return;
			    }
			s=trim(f.EndEtg_uf.value);
			if ((s=="")||(!uf_ok(s))) {
				alert('UF do endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_uf.focus();
				return;
				}
				
			if (!cep_ok(f.EndEtg_cep.value)) {
				alert('CEP do endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_cep.focus();
				return;
				}



<%if blnUsarMemorizacaoCompletaEnderecos and Not eh_cpf then%>
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

<% end if%>

<%if blnUsarMemorizacaoCompletaEnderecos then %>
			//validar enderecos de email
			if ((trim(f.EndEtg_email.value) != "") && (!email_ok(f.EndEtg_email.value))) {
                alert('Endere�o de entrega: e-mail inv�lido!!');
                f.EndEtg_email.focus();
                return;
            }

            if ((trim(f.EndEtg_email_xml.value) != "") && (!email_ok(f.EndEtg_email_xml.value))) {
                alert('Endere�o de entrega: e-mail (XML) inv�lido!!');
                f.EndEtg_email_xml.focus();
                return;
            }
<% end if%>

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
		strMsgErro="A forma de pagamento " + KEY_ASPAS + f.c_custoFinancFornecParcelamentoDescricao.value.toLowerCase() + KEY_ASPAS + " n�o est� dispon�vel para o(s) produto(s):"+strMsgErro;
		alert(strMsgErro);
		return;
		}

	blnHouveEdicaoVlUnitario = false;
	blnHouveEdicaoVlUnitarioComToleranciaArred = false;
	for (i = 0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value) != "") {
			if (f.c_vl_unitario[i].value != f.c_vl_unitario_original[i].value) {
				blnHouveEdicaoVlUnitario = true;
				if (Math.abs(converte_numero(f.c_vl_unitario[i].value) - converte_numero(f.c_vl_unitario_original[i].value)) > MAX_VALOR_MARGEM_ERRO_PAGAMENTO) blnHouveEdicaoVlUnitarioComToleranciaArred = true;
				break;
			}
		}
	}

    if (f.blnEtgImediataEdicaoLiberada.value == "<%=Cstr(True)%>") {
        if (f.rb_etg_imediata[0].checked) {
            if (trim(f.c_data_previsao_entrega.value) == "") {
                alert("Informe a data de previs�o de entrega!");
                f.c_data_previsao_entrega.focus();
                return;
            }

            if (!isDate(f.c_data_previsao_entrega)) {
                alert("Data de previs�o de entrega � inv�lida!");
                f.c_data_previsao_entrega.focus();
                return;
            }

            if (retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(f.c_data_previsao_entrega.value)) <= retorna_so_digitos(formata_ddmmyyyy_yyyymmdd('<%=formata_data(Date)%>'))) {
                alert("Data de previs�o de entrega deve ser uma data futura!");
                f.c_data_previsao_entrega.focus();
                return;
            }
        }
    }

	// Percentual m�ximo de comiss�o e desconto
	// ========================================
	// Lembretes:
	//	Regra anterior:
	//		1) Na 'Central', se o usu�rio possuir as permiss�es de acesso 'Editar Item do Pedido' e 'Editar RT e RA' poder� editar
	//			livremente o percentual de RT e o pre�o de venda (desconto) do produto, j� que nenhuma consist�ncia ser� realizada.
	//		2) Na 'Loja', o percentual de RT pode ser alterado se o usu�rio possuir a permiss�o de acesso 'Editar RT', sendo que
	//			somente a RT deve ser editada para que um percentual qualquer seja aceito sem que a consist�ncia seja realizada.
	//			Caso o pre�o de venda (desconto) seja alterado, ser�o aplicadas as verifica��es do 'percentual m�ximo de comiss�o e desconto'.
	//	Regra atual:
	//		Tanto na Central quanto na Loja, o pre�o de venda (desconto) do produto e a RT ser�o validados com base nos percentuais definidos em
	//		"Percentual M�ximo de Comiss�o e Desconto por Loja", o que inclui os descontos por al�ada.
	blnFormaPagtoEditada = isFormaPagtoEditada(f);
	if (blnFormaPagtoEditada)
		f.blnFormaPagtoEditada.value = '<%=Cstr(True)%>';
	else
		f.blnFormaPagtoEditada.value = '<%=Cstr(False)%>';
	
	if (blnHouveEdicaoVlUnitario || blnFormaPagtoEditada || (f.c_perc_RT.value != f.c_perc_RT_original.value)) {
		f.c_consiste_perc_max_comissao_e_desconto.value = 'S';
		// Consiste percentual m�ximo de comiss�o e desconto
		objSenhaDesconto = null;
		perc_RT = converte_numero(f.c_perc_RT.value);
		perc_max_RT_padrao = converte_numero(f.c_PercMaxRT.value);
		perc_max_RT_a_utilizar = perc_max_RT_padrao;

		perc_max_comissao_e_desconto = converte_numero(f.c_PercMaxComissaoEDesconto.value);
		perc_max_comissao_e_desconto_pj = converte_numero(f.c_PercMaxComissaoEDescontoPj.value);
		perc_max_comissao_e_desconto_nivel2 = converte_numero(f.c_PercMaxComissaoEDescontoNivel2.value);
		perc_max_comissao_e_desconto_nivel2_pj = converte_numero(f.c_PercMaxComissaoEDescontoNivel2Pj.value);
		perc_max_comissao_e_desconto_a_utilizar = obtem_perc_comissao_e_desconto_a_utilizar(f, calcula_vl_total_preco_venda(f), perc_max_comissao_e_desconto, perc_max_comissao_e_desconto_pj, perc_max_comissao_e_desconto_nivel2, perc_max_comissao_e_desconto_nivel2_pj);
		perc_desc_medio = calcula_desconto_medio();

		// Verifica se o usu�rio tem permiss�o de desconto por al�ada
		perc_max_comissao_alcada1 = converte_numero(f.c_PercMaxRTAlcada1.value);
		perc_max_desc_alcada_1_pf = converte_numero(f.c_PercMaxDescAlcada1Pf.value);
		perc_max_desc_alcada_1_pj = converte_numero(f.c_PercMaxDescAlcada1Pj.value);
		perc_max_comissao_alcada2 = converte_numero(f.c_PercMaxRTAlcada2.value);
		perc_max_desc_alcada_2_pf = converte_numero(f.c_PercMaxDescAlcada2Pf.value);
		perc_max_desc_alcada_2_pj = converte_numero(f.c_PercMaxDescAlcada2Pj.value);
		perc_max_comissao_alcada3 = converte_numero(f.c_PercMaxRTAlcada3.value);
		perc_max_desc_alcada_3_pf = converte_numero(f.c_PercMaxDescAlcada3Pf.value);
		perc_max_desc_alcada_3_pj = converte_numero(f.c_PercMaxDescAlcada3Pj.value);

		if (perc_max_comissao_alcada1 > perc_max_RT_a_utilizar) perc_max_RT_a_utilizar = perc_max_comissao_alcada1;
		if (perc_max_comissao_alcada2 > perc_max_RT_a_utilizar) perc_max_RT_a_utilizar = perc_max_comissao_alcada2;
		if (perc_max_comissao_alcada3 > perc_max_RT_a_utilizar) perc_max_RT_a_utilizar = perc_max_comissao_alcada3;

		if (f.c_tipo_cliente.value == ID_PF) {
			if (perc_max_desc_alcada_1_pf > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_1_pf;
			if (perc_max_desc_alcada_2_pf > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_2_pf;
			if (perc_max_desc_alcada_3_pf > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_3_pf;
		}
		else {
			if (perc_max_desc_alcada_1_pj > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_1_pj;
			if (perc_max_desc_alcada_2_pj > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_2_pj;
			if (perc_max_desc_alcada_3_pj > perc_max_comissao_e_desconto_a_utilizar) perc_max_comissao_e_desconto_a_utilizar = perc_max_desc_alcada_3_pj;
		}

		// Verifica se todos os produtos cujo desconto excedem o m�ximo permitido possuem senha de desconto dispon�vel
		// La�o p/ produtos
		strMsgErro = "";
		for (i = 0; i < f.c_produto.length; i++) {
			if ((trim(f.c_produto[i].value) != "") && ((blnFormaPagtoEditada && (!blnUsuarioDeptoFinanceiro)) || (f.c_vl_unitario[i].value != f.c_vl_unitario_original[i].value))) {
				perc_senha_desconto = 0;
				vl_preco_lista = converte_numero(f.c_preco_lista[i].value);
				vl_preco_venda = converte_numero(f.c_vl_unitario[i].value);
				if (vl_preco_lista == 0) {
					perc_desc = 0;
				}
				else {
					perc_desc = 100 * (vl_preco_lista - vl_preco_venda) / vl_preco_lista;
				}

				// Tem desconto: sim
				if (perc_desc != 0) {
					// Desconto excede limite m�ximo: sim
					if (perc_desc > perc_max_comissao_e_desconto_a_utilizar) {
						// Tem senha de desconto?
						if (objSenhaDesconto == null) {
							executa_consulta_senha_desconto(f.cliente_selecionado.value, f.c_loja.value);
						}
						for (j = 0; j < objSenhaDesconto.item.length; j++) {
							if ((objSenhaDesconto.item[j].fabricante == f.c_fabricante[i].value) && (objSenhaDesconto.item[j].produto == f.c_produto[i].value)) {
								perc_senha_desconto = converte_numero(objSenhaDesconto.item[j].desc_max);
								break;
							}
						}
						// Tem senha de desconto: sim
						if (perc_senha_desconto != 0) {
							// Senha de desconto N�O cobre desconto
							if (perc_senha_desconto < perc_desc) {
								if (strMsgErro != "") strMsgErro += "\n";
								strMsgErro += "O desconto do produto '" + f.c_descricao[i].value + "' (" + formata_numero(perc_desc, 2) + "%) excede o m�ximo autorizado!!";
							}
						}
						// N�o tem senha de desconto
						else {
							if (strMsgErro != "") strMsgErro += "\n";
							strMsgErro += "O desconto do produto '" + f.c_descricao[i].value + "' (" + formata_numero(perc_desc, 2) + "%) excede o m�ximo permitido!!";
						}
					} // if (perc_desc > perc_max_comissao_e_desconto_a_utilizar)
				} // if (perc_desc != 0)
			} // if (trim(f.c_produto[i].value) != "")
		} // for (la�o produtos)

		if (strMsgErro != "") {
			strMsgErro += "\n\nN�o � poss�vel continuar!!";
			alert(strMsgErro);
			return;
		}

		// Redu��o autom�tica da RT caso a soma da comiss�o com o desconto m�dio exceda o m�ximo
		var blnIgnoraProcReducaoAutomaticaRT;
		blnIgnoraProcReducaoAutomaticaRT = false;
		if (f.c_loja.value == NUMERO_LOJA_ECOMMERCE_AR_CLUBE) blnIgnoraProcReducaoAutomaticaRT = true;
		// Tem RT?
		if (perc_RT == 0) blnIgnoraProcReducaoAutomaticaRT = true;
		// Edi��o sendo feita pelo depto financeiro? (Obs: desde que n�o tenha editado a RT e nem do pre�o de venda)
		if (blnUsuarioDeptoFinanceiro && (!blnHouveEdicaoVlUnitarioComToleranciaArred) && (f.c_perc_RT.value == f.c_perc_RT_original.value)) blnIgnoraProcReducaoAutomaticaRT = true;

		if (!blnIgnoraProcReducaoAutomaticaRT) {
			// RT excede limite m�ximo?
			if (f.c_perc_RT.value != f.c_perc_RT_original.value) {
				if (perc_RT > perc_max_RT_a_utilizar) {
					alert("Percentual de comiss�o excede o m�ximo permitido!!");
					return;
				}
			}

			// Neste ponto, � certo que todos os produtos que possuem desconto est�o dentro do m�ximo permitido
			// ou possuem senha de desconto autorizando.
			// Verifica-se agora se � necess�rio reduzir automaticamente o percentual da RT usando p/ o c�lculo
			// o percentual de desconto m�dio.
			perc_RT_novo = Math.min(perc_RT, (perc_max_comissao_e_desconto_a_utilizar - perc_desc_medio));
			if (perc_RT_novo < 0) perc_RT_novo = 0;

			// O percentual de RT ser� alterado automaticamente, solicita confirma��o
			if (perc_RT_novo != perc_RT) {
				s = "A soma dos percentuais de comiss�o (" + formata_numero(perc_RT, 2) + "%) e de desconto m�dio do(s) produto(s) (" + formata_numero(perc_desc_medio, 2) + "%) totaliza " + formata_numero(perc_desc_medio + perc_RT, 2) + "% e excede o m�ximo permitido!!" +
					"\nA comiss�o ser� reduzida automaticamente para " + formata_numero(perc_RT_novo, 2) + "%!!" +
					"\nContinua?";
				if (!confirm(s)) {
					s = "Opera��o cancelada!!";
					alert(s);
					return;
				}
				else {
					// Verifica se o novo percentual de RT est� dentro do limite definido p/ o perfil do usu�rio que est� editando o pedido
					if (perc_RT_novo > perc_max_RT_a_utilizar) {
						s = "O percentual de comiss�o (" + formata_numero(perc_RT_novo, 2) + "%) excede o m�ximo permitido!!" +
							"\nA comiss�o ser� reduzida automaticamente para " + formata_numero(perc_max_RT_a_utilizar, 2) + "%!!" +
							"\nContinua?";
						if (!confirm(s)) {
							s = "Opera��o cancelada!!";
							alert(s);
							return;
						}
						else {
							// Novo percentual de RT
							perc_RT_novo = perc_max_RT_a_utilizar;
						}
					}

					// Novo percentual de RT
					f.c_perc_RT.value = formata_perc_RT(perc_RT_novo);
					f.c_gravar_perc_RT_novo.value = "S";
					perc_RT = perc_RT_novo;
				}
			}
		} // if (!blnIgnoraProcReducaoAutomaticaRT)
	} // if (blnHouveEdicaoVlUnitario || blnFormaPagtoEditada || (f.c_perc_RT.value != f.c_perc_RT_original.value))

    //campos do endere�o de entrega que precisam de transformacao
    transferirCamposEndEtg(f);


	f.action="pedidoatualiza.asp";
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
    }



    function transferirCamposEndEtg(fNEW) {
<%if blnUsarMemorizacaoCompletaEnderecos and Not eh_cpf and blnEndEntregaEdicaoLiberada then%>
        //Transferimos os dados do endere�o de entrega dos campos certos. 
        //Temos dois conjuntos de campos (para PF e PJ) porque o layout � muito diferente.
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

        //os campos a mais s�o enviados junto. Deixamos enviar...
<%end if%>
    }

    //para mudar o tipo do endere�o de entrega
    function trocarEndEtgTipoPessoa(novoTipo) {
<%if blnUsarMemorizacaoCompletaEnderecos then%>
        if (novoTipo && $('input[name="EndEtg_tipo_pessoa"]:disabled').length == 0)
            setarValorRadio($('input[name="EndEtg_tipo_pessoa"]'), novoTipo);

        var pf = $('input[name="EndEtg_tipo_pessoa"]:checked').val() == "PF";

        //se nao tiver nada selecionado queremos tratar cono pj
        if (!pf) {
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

    function trataProdutorRural() {
        //ao clicar na op��o Produtor Rural, exibir/ocultar os campos apropriados
        if (!fPED.rb_produtor_rural[1].checked) {
            $("#t_contribuinte_icms").css("display", "none");
        }
        else {
            $("#t_contribuinte_icms").css("display", "block");
        }
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

    function exibeJanelaCEP() {
        $.mostraJanelaCEP("endereco__cep", "endereco__uf", "endereco__cidade", "endereco__bairro", "endereco__endereco", "endereco__numero", "endereco__complemento");
    }
</script>
<script language='JavaScript'>
    function SomenteNumero(e){
        var tecla=(window.event)?event.keyCode:e.which;   
        if((tecla>47 && tecla<58)) return true;
        else{
            if (tecla==8 || tecla==0) return true;
            else  return false;
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
#rb_etg_imediata, #rb_bem_uso_consumo {
	margin: 0pt 2pt 1pt 15pt;
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
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
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
<!-- ********************************************************** -->
<!-- **********  P�GINA PARA EDITAR ITENS DO PEDIDO  ********** -->
<!-- ********************************************************** -->
<body id="corpoPagina" onload="processaFormaPagtoDefault();">
<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=r_pedido.id_cliente%>'>
<input type="hidden" name="c_cnpj_cpf" id="c_cnpj_cpf" value='<%=r_cliente.cnpj_cpf%>'>
<input type="hidden" name="c_tipo_cliente" id="c_tipo_cliente" value='<%=r_cliente.tipo%>'>
<input type="hidden" name="c_consiste_perc_max_comissao_e_desconto" id="c_consiste_perc_max_comissao_e_desconto" value=''>
<input type="hidden" name="c_PercMaxRT" id="c_PercMaxRT" value='<%=strPercMaxRT%>'>
<input type="hidden" name="c_PercMaxComissaoEDesconto" id="c_PercMaxComissaoEDesconto" value='<%=strPercMaxComissaoEDesconto%>'>
<input type="hidden" name="c_PercMaxComissaoEDescontoPj" id="c_PercMaxComissaoEDescontoPj" value='<%=strPercMaxComissaoEDescontoPj%>'>
<input type="hidden" name="c_PercMaxComissaoEDescontoNivel2" id="c_PercMaxComissaoEDescontoNivel2" value='<%=strPercMaxComissaoEDescontoNivel2%>'>
<input type="hidden" name="c_PercMaxComissaoEDescontoNivel2Pj" id="c_PercMaxComissaoEDescontoNivel2Pj" value='<%=strPercMaxComissaoEDescontoNivel2Pj%>'>
<input type="hidden" name="c_PercMaxRTAlcada1" id="c_PercMaxRTAlcada1" value="<%=strPercMaxRTAlcada1%>" />
<input type="hidden" name="c_PercMaxDescAlcada1Pf" id="c_PercMaxDescAlcada1Pf" value="<%=strPercMaxDescAlcada1Pf%>" />
<input type="hidden" name="c_PercMaxDescAlcada1Pj" id="c_PercMaxDescAlcada1Pj" value="<%=strPercMaxDescAlcada1Pj%>" />
<input type="hidden" name="c_PercMaxRTAlcada2" id="c_PercMaxRTAlcada2" value="<%=strPercMaxRTAlcada2%>" />
<input type="hidden" name="c_PercMaxDescAlcada2Pf" id="c_PercMaxDescAlcada2Pf" value="<%=strPercMaxDescAlcada2Pf%>" />
<input type="hidden" name="c_PercMaxDescAlcada2Pj" id="c_PercMaxDescAlcada2Pj" value="<%=strPercMaxDescAlcada2Pj%>" />
<input type="hidden" name="c_PercMaxRTAlcada3" id="c_PercMaxRTAlcada3" value="<%=strPercMaxRTAlcada3%>" />
<input type="hidden" name="c_PercMaxDescAlcada3Pf" id="c_PercMaxDescAlcada3Pf" value="<%=strPercMaxDescAlcada3Pf%>" />
<input type="hidden" name="c_PercMaxDescAlcada3Pj" id="c_PercMaxDescAlcada3Pj" value="<%=strPercMaxDescAlcada3Pj%>" />
<input type="hidden" name="c_permite_RA_status" id="c_permite_RA_status" value='<%=r_pedido.permite_RA_status%>' />
<input type="hidden" name="c_st_violado_permite_RA_status" id="c_st_violado_permite_RA_status" value='<%=r_pedido.st_violado_permite_RA_status%>' />
<input type="hidden" name="c_gravar_perc_RT_novo" id="c_gravar_perc_RT_novo" value='N' />
<input type="hidden" name="c_PercLimiteRASemDesagio" id="c_PercLimiteRASemDesagio" value='<%=strPercLimiteRASemDesagio%>'>
<input type="hidden" name="c_PercDesagio" id="c_PercDesagio" value='<%=strPercDesagio%>'>
<input type="hidden" name="c_ped_bonshop" id="c_ped_bonshop" value='<%=r_pedido.pedido_bs_x_at %>' />
<input type="hidden" name="c_opcao_forca_desagio" id="c_opcao_forca_desagio" value='<%=strOpcaoForcaDesagio %>'>
<input type="hidden" name="c_qtde_pedidos_entregues" id="c_qtde_pedidos_entregues" value='<%=CStr(qtde_pedidos_entregues)%>'>
<input type="hidden" name="c_qtde_parcelas_desagio_RA" id="c_qtde_parcelas_desagio_RA" value='<%=CStr(r_pedido.qtde_parcelas_desagio_RA)%>'>
<input type="hidden" name="tipo_parcelamento" id="tipo_parcelamento" value='<%=r_pedido.tipo_parcelamento%>'>
<input type="hidden" name="c_loja" id="c_loja" value='<%=r_pedido.loja%>'>
<input type="hidden" name="GarantiaIndicadorStatusOriginal" id="GarantiaIndicadorStatusOriginal" value='<%=r_pedido.GarantiaIndicadorStatus%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamentoOriginal" id="c_custoFinancFornecTipoParcelamentoOriginal" value='<%=r_pedido.custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelasOriginal" id="c_custoFinancFornecQtdeParcelasOriginal" value='<%=r_pedido.custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamento" id="c_custoFinancFornecTipoParcelamento" value='<%=r_pedido.custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelas" id="c_custoFinancFornecQtdeParcelas" value='<%=r_pedido.custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamentoUltConsulta" id="c_custoFinancFornecTipoParcelamentoUltConsulta" value='<%=r_pedido.custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelasUltConsulta" id="c_custoFinancFornecQtdeParcelasUltConsulta" value='<%=r_pedido.custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecParcelamentoDescricao" id="c_custoFinancFornecParcelamentoDescricao" value=''>
<input type="hidden" name="blnIndicadorEdicaoLiberada" id="blnIndicadorEdicaoLiberada" value='<%=Cstr(blnIndicadorEdicaoLiberada)%>'>
<input type="hidden" name="blnNumPedidoECommerceEdicaoLiberada" id="blnNumPedidoECommerceEdicaoLiberada" value="<%=Cstr(blnNumPedidoECommerceEdicaoLiberada)%>" />
<input type="hidden" name="blnObs1EdicaoLiberada" id="blnObs1EdicaoLiberada" value='<%=Cstr(blnObs1EdicaoLiberada)%>'>
<input type="hidden" name="nivelEdicaoFormaPagto" id="nivelEdicaoFormaPagto" value='<%=Cstr(nivelEdicaoFormaPagto)%>'>
<input type="hidden" name="blnFormaPagtoEditada" id="blnFormaPagtoEditada" />
<input type="hidden" name="bln_RA_EdicaoLiberada" id="bln_RA_EdicaoLiberada" value='<%=Cstr(bln_RA_EdicaoLiberada)%>'>
<input type="hidden" name="bln_RT_EdicaoLiberada" id="bln_RT_EdicaoLiberada" value='<%=Cstr(bln_RT_EdicaoLiberada)%>'>
<input type="hidden" name="blnItemPedidoEdicaoLiberada" id="blnItemPedidoEdicaoLiberada" value='<%=Cstr(blnItemPedidoEdicaoLiberada)%>'>
<input type="hidden" name="blnEtgImediataEdicaoLiberada" id="blnEtgImediataEdicaoLiberada" value='<%=Cstr(blnEtgImediataEdicaoLiberada)%>'>
<input type="hidden" name="blnEndEntregaEdicaoLiberada" id="blnEndEntregaEdicaoLiberada" value='<%=Cstr(blnEndEntregaEdicaoLiberada)%>'>
<input type="hidden" name="blnAnaliseCreditoEdicaoLiberada" id="blnAnaliseCreditoEdicaoLiberada" value='<%=Cstr(blnAnaliseCreditoEdicaoLiberada)%>'>
<input type="hidden" name="blnBemUsoConsumoEdicaoLiberada" id="blnBemUsoConsumoEdicaoLiberada" value='<%=Cstr(blnBemUsoConsumoEdicaoLiberada)%>'>
<input type="hidden" name="blnGarantiaIndicadorEdicaoLiberada" id="blnGarantiaIndicadorEdicaoLiberada" value='<%=Cstr(blnGarantiaIndicadorEdicaoLiberada)%>'>
<input type="hidden" name="url_origem" id="url_origem" value="<%=url_origem%>" />
<% if Not blnIndicadorEdicaoLiberada then %>
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=r_pedido.indicador%>" />
<% end if %>

<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>


<!--  I D E N T I F I C A � � O   D O   P E D I D O  -->
<%=MontaHeaderIdentificacaoPedido(pedido_selecionado, r_pedido, 649)%>
<br>

<!--  L O J A   -->
<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	set r_loja = New cl_LOJA
	if x_loja_bd(r_pedido.loja, r_loja) then
		with r_loja
			if Trim(.razao_social) <> "" then
				s = Trim(.razao_social)
			else
				s = Trim(.nome)
				end if
			end with
		end if
%>
	<td class="MD" align="left"><p class="Rf">LOJA</p><p class="C"><%=s%>&nbsp;</p></td>
	<td width="145" class="MD" align="left"><p class="Rf">INDICADOR</p>
        <% if Not blnIndicadorEdicaoLiberada then %>
        <p class="C"><%=r_pedido.indicador%></p>
        <%else%>
        <select name="c_indicador" id="c_indicador" style="margin-top:0; margin-bottom:0;width:120px" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
            <%=indicadores_monta_itens_select(r_pedido.indicador)%>
        </select>
        <%end if %>
	</td>
	<td width="145" align="left"><p class="Rf">VENDEDOR</p><p class="C"><%=r_pedido.vendedor%>&nbsp;</p></td>
	</tr>
	</table>

<br>



<!--  ENDERE�O DO CLIENTE  -->
<% if r_pedido.st_memorizacao_completa_enderecos = 0 then %>
	<!--  CLIENTE   -->
	<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	if xcliente_bd_resultado then
%>
<%	if cliente__tipo = ID_PF then s_aux="CPF" else s_aux="CNPJ"
	s = cnpj_cpf_formata(cliente__cnpj_cpf) 
%>
		<td align="left" width="50%" class="MD"><p class="Rf"><%=s_aux%></p>
		
			<p class="C"><%=s%>&nbsp;</p>
		
		</td>
		<%
		if cliente__tipo = ID_PF then s = Trim(cliente__rg) else s = Trim(cliente__ie)
			if cliente__tipo = ID_PF then 
%>
	<td align="left"><p class="Rf">RG</p><p class="C"><%=s%>&nbsp;</p></td>
<% else %>
	<td align="left"><p class="Rf">IE</p><p class="C"><%=s%>&nbsp;</p></td>
<% end if %>
		</tr>
	<%
		if Trim(cliente__nome) <> "" then
			s = Trim(cliente__nome)
			end if
		end if
	
	if cliente__tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZ�O SOCIAL DO CLIENTE"
%>
    <tr>
	<td class="MC" align="left" colspan="2"><p class="Rf"><%=s_aux%></p>
	
		<p class="C"><%=s%>&nbsp;</p>
	
		</td>
	</tr>
	</table>
	
	<!--  ENDERE�O DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	s = formata_endereco(cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep)
%>		
		<td align="left"><p class="Rf">ENDERE�O</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
</table>
	
	<!--  TELEFONE DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
	<tr>
<%	s = ""
	if Trim(cliente__tel_res) <> "" then
		s = telefone_formata(Trim(cliente__tel_res))
		s_aux=Trim(cliente__ddd_res)
		if s_aux<>"" then s = "(" & s_aux & ") " & s
		end if
	
	s2 = ""
	if Trim(cliente__tel_com) <> "" then
		s2 = telefone_formata(Trim(cliente__tel_com))
		s_aux = Trim(cliente__ddd_com)
		if s_aux<>"" then s2 = "(" & s_aux & ") " & s2
		s_aux = Trim(cliente__ramal_com)
		if s_aux<>"" then s2 = s2 & "  (R. " & s_aux & ")"
		end if
	if Trim(cliente__tel_cel) <> "" then
		s3 = telefone_formata(Trim(cliente__tel_cel))
		s_aux = Trim(cliente__ddd_cel)
		if s_aux<>"" then s3 = "(" & s_aux & ") " & s3
		end if
	if Trim(cliente__tel_com_2) <> "" then
		s4 = telefone_formata(Trim(cliente__tel_com_2))
		s_aux = Trim(cliente__ddd_com_2)
		if s_aux<>"" then s4 = "(" & s_aux & ") " & s4
		s_aux = Trim(cliente__ramal_com_2)
		if s_aux<>"" then s4 = s4 & "  (R. " & s_aux & ")"
		end if
	
%>

<% if cliente__tipo = ID_PF then %>
	<td class="MD" width="33%" align="left"><p class="Rf">TELEFONE RESIDENCIAL</p><p class="C"><%=s%>&nbsp;</p></td>
	<td class="MD" width="33%" align="left"><p class="Rf">TELEFONE COMERCIAL</p><p class="C"><%=s2%>&nbsp;</p></td>
		<td align="left"><p class="Rf">CELULAR</p><p class="C"><%=s3%>&nbsp;</p></td>

<% else %>
	<td class="MD" width="50%" align="left"><p class="Rf">TELEFONE</p><p class="C"><%=s2%>&nbsp;</p></td>
	<td width="50%" align="left"><p class="Rf">TELEFONE</p><p class="C"><%=s4%>&nbsp;</p></td>

<% end if %>

	</tr>
</table>
	
	<!--  E-MAIL DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left"><p class="Rf">E-MAIL</p><p class="C"><%=Trim(cliente__email)%>&nbsp;</p></td>
	</tr>
</table>

<% else %>

			<!--  CLIENTE   -->
	<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	if xcliente_bd_resultado then
%>
<%	if cliente__tipo = ID_PF then s_aux="CPF" else s_aux="CNPJ"
	s = cnpj_cpf_formata(cliente__cnpj_cpf) 
%>
		<td align="left" class="MD" width="210"><p class="Rf"><%=s_aux%></p>
		
			<!--<p class="C"><%=s%>&nbsp;</p>-->
			<input id="endereco__cpf_cnpj" name="endereco__cpf_cnpj" readonly="readonly" class="TA" maxlength="72" style="width:310px;" value="<%=s%>">
		
		</td>
<%
if cliente__tipo = ID_PF then s = Trim(cliente__rg) else s = Trim(cliente__ie)
if cliente__tipo = ID_PF then 
%>
	<td align="left"><p class="Rf">RG</p><input id="cliente__rg" name="cliente__rg" class="TA" maxlength="72" style="width:310px;" value="<%=s%>" <%=strAtributosDadosCadastrais%> ></td>
	</tr>
	</table>


	<table width="649" class="QS" cellspacing="0">
		<tr>
			<td align="left"><p class="R">PRODUTOR RURAL</p><p class="C">
				<%s=cliente__produtor_rural_status%>
				<%if s = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then s_aux="checked" else s_aux=""%>
				
				<input type="radio" id="rb_produtor_rural_nao" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" <%=s_aux%> onclick="trataProdutorRural();" <%=strAtributosRadioboxDadosCadastrais%> ><span class="C" style="cursor:default" onclick="fPED.rb_produtor_rural[0].click();">N�o</span>
				<%if s = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) then s_aux="checked" else s_aux=""%>
				
				<input type="radio" id="rb_produtor_rural_sim" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" <%=s_aux%> onclick="trataProdutorRural();" <%=strAtributosRadioboxDadosCadastrais%> ><span class="C" style="cursor:default" onclick="fPED.rb_produtor_rural[1].click();">Sim</span></p>
				<% if not blnDadosCadastraisEdicaoLiberada then %>
					<input type="hidden" name="rb_produtor_rural" value="<%=cliente__produtor_rural_status%>" />
				<% end if %>
			</td>
		</tr>
	</table>
	<script type="text/javascript">
        $(function () { trataProdutorRural(); });
	</script>

	

	<table width="649" class="QS" cellspacing="0" id="t_contribuinte_icms">
		<tr>
			<%s=cliente__ie%>
			<td width="210" class="MD" align="left"><p class="R">IE</p><p class="C">
				<input id="cliente__ie" name="cliente__ie" class="TA" maxlength="72" style="width:310px;" value="<%=s%>"  <%=strAtributosDadosCadastrais%> /></p>
			</td>
			<td align="left"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
				<%s=cliente__icms%>
				<%if s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then s_aux="checked" else s_aux=""%>
				<% intIdx = 0 %>
				<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s_aux%>  <%=strAtributosRadioboxDadosCadastrais%> ><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">N�o</span>
				<%if s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then s_aux="checked" else s_aux=""%>
				<% intIdx = intIdx + 1 %>
				<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s_aux%>  <%=strAtributosRadioboxDadosCadastrais%> ><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Sim</span>
				<%if s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then s_aux="checked" else s_aux=""%>
				<% intIdx = intIdx + 1 %>
				<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s_aux%>  <%=strAtributosRadioboxDadosCadastrais%> ><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Isento</span></p>
				<% if not blnDadosCadastraisEdicaoLiberada then %>
					<input type="hidden" name="rb_contribuinte_icms" value="<%=cliente__icms%>" />
				<% end if %>
			</td>
		</tr>
	</table>
	

<% else %>
	<td width="215" align="left"><p class="Rf">IE</p><input id="cliente__ie" name="cliente__ie" class="TA" maxlength="72" style="width:310px;" value="<%=s%>"  <%=strAtributosDadosCadastrais%> ></td>
	</tr>
	<tr>
		<td class="MC" align="left" colspan="2"><p class="R">CONTRIBUINTE ICMS</p><p class="C">

				<%
                    s = " "
                    if r_pedido.endereco_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
                        s = " checked "
                    end if
                %>
			
			<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s%>  <%=strAtributosRadioboxDadosCadastrais%> ><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[1].click();">N�o</span>
				<%
                    s = " "
                    if r_pedido.endereco_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                        s = " checked "
                    end if
                %>
			<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s%>  <%=strAtributosRadioboxDadosCadastrais%> ><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[2].click();">Sim</span>
				<%
                    s = " "
                    if r_pedido.endereco_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
                        s = " checked "
                    end if
                %>
			<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s%>  <%=strAtributosRadioboxDadosCadastrais%> ><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[3].click();">Isento</span></p>
			
			<% if not blnDadosCadastraisEdicaoLiberada then %>
				<input type="hidden" name="rb_contribuinte_icms" value="<%=r_pedido.endereco_contribuinte_icms_status%>" />
			<% end if %>
		</td>
	</tr>
	</table>
<% end if %>
		
		
	<table width="649" class="QS" cellspacing="0">
		
	<%
		if Trim(cliente__nome) <> "" then
			s = Trim(cliente__nome)
			end if
		end if
	
	if cliente__tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZ�O SOCIAL DO CLIENTE"
%>
    <tr>
	<td align="left" colspan="2"><p class="Rf"><%=s_aux%></p>
	
		
		<input id="cliente__nome" name="cliente__nome" class="TA" value="<%=s%>" maxlength="60" style="width:635px;"  <%=strAtributosDadosCadastrais%> />
				
	
		</td>
	</tr>
	</table>
	
	<!--  ENDERE�O DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
	    <tr>           
		    <td colspan="2" class="MB" align="left"><p class="Rf">ENDERE�O</p><input id="endereco__endereco" name="endereco__endereco" class="TA" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.endereco__numero.focus(); filtra_nome_identificador();" value="<%=cliente__endereco%>" <%=strAtributosDadosCadastrais%> ></td>
	    </tr>
	    <tr>
		    <td class="MDB" align="left"><p class="Rf">N�</p><input id="endereco__numero" name="endereco__numero" class="TA" maxlength="<%=MAX_TAMANHO_CAMPO_ENDERECO_NUMERO%>" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.endereco__complemento.focus(); filtra_nome_identificador();" value="<%=cliente__endereco_numero%>" <%=strAtributosDadosCadastrais%> ></td>
		    <td class="MB" align="left"><p class="Rf">COMPLEMENTO</p><input id="endereco__complemento" name="endereco__complemento" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.endereco__bairro.focus(); filtra_nome_identificador();" value="<%=cliente__endereco_complemento%>" <%=strAtributosDadosCadastrais%> ></td>
	    </tr>
	    <tr>
		    <td class="MDB" align="left"><p class="Rf">BAIRRO</p><input id="endereco__bairro" name="endereco__bairro" class="TA" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.endereco__cidade.focus(); filtra_nome_identificador();" value="<%=cliente__bairro%>" <%=strAtributosDadosCadastrais%> ></td>
		    <td class="MB" align="left"><p class="Rf">CIDADE</p><input id="endereco__cidade" name="endereco__cidade" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.endereco__uf.focus(); filtra_nome_identificador();" value="<%=cliente__cidade%>" <%=strAtributosDadosCadastrais%> ></td>
	    </tr>
	    <tr>
		    <td width="50%" class="MD" align="left"><p class="Rf">UF</p><input id="endereco__uf" name="endereco__uf" class="TA" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fPED.endereco__cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inv�lida!!');this.focus();} else this.value=ucase(this.value);" value="<%=cliente__uf%>" <%=strAtributosDadosCadastrais%> ></td>
		    <td>
			    <table width="100%" cellspacing="0" cellpadding="0">
			    <tr>
			    <td width="50%" align="left"><p class="Rf">CEP</p><input id="endereco__cep" name="endereco__cep" readonly tabindex=-1 class="TA" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inv�lido!!');this.focus();} else this.value=cep_formata(this.value);" value='<%=cep_formata(cliente__cep)%>' <%=strAtributosDadosCadastrais%> ></td>
			    <td align="center">
				    <% if blnDadosCadastraisEdicaoLiberada then %>
						<% if blnPesquisaCEPAntiga then %>
						<button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCepEndEtg();">Pesquisar CEP</button>
						<% end if %>
						<% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
						<% if blnPesquisaCEPNova then %>
						<button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="exibeJanelaCEP();">Pesquisar CEP</button>
						<% end if %>
					<% end if %>
			    </td>
			    </tr>
			    </table>
		    </td>
	    </tr>
    </table>
	
		<% if cliente__tipo = ID_PF then %>
			<!--  TELEFONE DO CLIENTE  -->
			<table width="649" class="QS" cellspacing="0">
	            <tr>
					<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_res" name="cliente__ddd_res" class="TA" value="<%=cliente__ddd_res%>" maxlength="4" size="5" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.cliente__tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p>
					</td>
					<td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
						<input id="cliente__tel_res" name="cliente__tel_res" class="TA" value="<%=telefone_formata(cliente__tel_res)%>" maxlength="11" size="12" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.cliente__ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
					</td>
	            </tr>
			</table>
			<table width="649" class="QS" cellspacing="0">
				<tr>
					<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_cel" name="cliente__ddd_cel" class="TA" value="<%=cliente__ddd_cel%>" maxlength="4" size="5" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.cliente__tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p>
					</td>
					<td align="left"><p class="R">CELULAR</p><p class="C">
						<input id="cliente__tel_cel" name="cliente__tel_cel" class="TA" value="<%=telefone_formata(cliente__tel_cel)%>" maxlength="9" size="12" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.cliente__ddd_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('N�mero de celular inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
					</td>
	            </tr>
			</table>
			<table width="649" class="QS" cellspacing="0">
	            <tr>
					<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_com" name="cliente__ddd_com" class="TA" value="<%=cliente__ddd_com%>" maxlength="4" size="5" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.cliente__tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p>
					</td>
					<td class="MD" align="left"><p class="R">COMERCIAL</p><p class="C">
						<input id="cliente__tel_com" name="cliente__tel_com" class="TA" value="<%=telefone_formata(cliente__tel_com)%>" maxlength="9" size="12" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.cliente__ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('N�mero de telefone comercial inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
					</td>
					<td align="left"><p class="R">RAMAL</p><p class="C">
						<input id="cliente__ramal_com" name="cliente__ramal_com" class="TA" value="<%=cliente__ramal_com%>" maxlength="4" size="6" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true)) fPED.cliente__email.focus(); filtra_numerico();"></p>
					</td>
				</tr>
				
			</table>	

		<% else %>
			<!--  TELEFONE DO CLIENTE  -->
			<table width="649" class="QS" cellspacing="0">
	            <tr>
					<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_com" name="cliente__ddd_com" class="TA" value="<%=cliente__ddd_com%>" maxlength="4" size="5" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.cliente__tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
					<td class="MD" align="left"><p class="R">TELEFONE </p><p class="C">
						<input id="cliente__tel_com" name="cliente__tel_com" class="TA" value="<%=telefone_formata(cliente__tel_com)%>" maxlength="11" size="12" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.cliente__ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
					<td align="left"><p class="R">RAMAL</p><p class="C">
						<input id="cliente__ramal_com" name="cliente__ramal_com" class="TA" value="<%=cliente__ramal_com%>" maxlength="4" size="6" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true)) fPED.cliente__ddd_com_2.focus(); filtra_numerico();"></p>
					</td>
	            </tr>
	            <tr>
	                <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_com_2" name="cliente__ddd_com_2" class="TA" value="<%=cliente__ddd_com_2%>" maxlength="4" size="5" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.cliente__tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!!');this.focus();}" /></p>  
	                </td>
	                <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
						<input id="cliente__tel_com_2" name="cliente__tel_com_2" class="TA" value="<%=telefone_formata(cliente__tel_com_2)%>" maxlength="9" size="12" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.cliente__ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	                </td>
	                <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
						<input id="cliente__ramal_com_2" name="cliente__ramal_com_2" class="TA" value="<%=cliente__ramal_com_2%>" maxlength="4" size="6" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true)) fPED.cliente__email.focus(); filtra_numerico();" /></p>
	                </td>
	            </tr>
            </table>
		<% end if %>

	<!--  E-MAIL DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
		 <tr>           
		    <td colspan="2" class="Rf" align="left"><p class="Rf">E-MAIL</p>
				<input id="cliente__email" name="cliente__email" class="TA" maxlength="60" style="width:635px;" value="<%=cliente__email%>" <%=strAtributosDadosCadastrais%> onkeypress="if (digitou_enter(true)) fPED.cliente__email_xml.focus(); filtra_email();" />

		    </td>
	    </tr>
	</table>

	 <!-- ************   E-MAIL (XML)  ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td width="100%" align="left"><p class="R">E-MAIL (XML)</p><p class="C">
		    <input id="cliente__email_xml" name="cliente__email_xml" value="<%=cliente__email_xml%>" class="TA" maxlength="60" size="74" <%=strAtributosDadosCadastrais%> onkeypress="filtra_email();"></p></td>
	    </tr>
    </table>

<%end if%>





<br>
<%
	dim estilo_superior_entrega 
	estilo_superior_entrega = "Q"
%>

<% if Not blnEndEntregaEdicaoLiberada then %>
<!--  ENDERE�O DE ENTREGA  -->
<%	
	s = pedido_formata_endereco_entrega(r_pedido, r_cliente)
%>		
<table width="649" class="<%=estilo_superior_entrega%>" cellspacing="0" style="table-layout:fixed">
	<%
		estilo_superior_entrega = "QS"
	%>
	<tr>
		<td align="left"><p class="Rf">ENDERE�O DE ENTREGA</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
    <%	if r_pedido.EndEtg_cod_justificativa <> "" then %>		
	<tr>
		<td align="left" style="word-wrap:break-word"><p class="C"><%=obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA,r_pedido.EndEtg_cod_justificativa)%>&nbsp;</p></td>
	</tr>
    <%end if %>
</table>
<% else %>


    <% if blnUsarMemorizacaoCompletaEnderecos then %>
        <!--  ************  TIPO DO ENDERE�O DE ENTREGA: PF/PJ (SOMENTE SE O CLIENTE FOR PJ)   ************ -->

        <%if eh_cpf then%>
            <!-- ************   ENDERE�O DE ENTREGA PARA CLIENTE PF   ************ -->
			<!-- Pegamos todos os atuais. Sem campos edit�veis. Pegamos os atuais dos dados cadastrais do cliente, n�o do campo em si. -->
			<!-- Como n�o s�o edit�veis, sempre v�o ser iguais aos cadastrais. E se removermos o endere�o de entrega e informarmos novamente, eles devem ser preenchidos. -->
			<input type="hidden" id="EndEtg_tipo_pessoa" name="EndEtg_tipo_pessoa" value="PF"/>
			<input type="hidden" id="EndEtg_cnpj_cpf" name="EndEtg_cnpj_cpf" value="<%=cliente__cnpj_cpf%>"/>
			<input type="hidden" id="EndEtg_ie" name="EndEtg_ie" value="<%=cliente__ie%>"/>
			<input type="hidden" id="EndEtg_contribuinte_icms_status" name="EndEtg_contribuinte_icms_status" value="<%=cliente__icms%>"/>
			<input type="hidden" id="EndEtg_rg" name="EndEtg_rg" value="<%=cliente__rg%>"/>
			<input type="hidden" id="EndEtg_produtor_rural_status" name="EndEtg_produtor_rural_status" value="<%=cliente__produtor_rural_status%>"/>
			<input type="hidden" id="EndEtg_nome" name="EndEtg_nome" value="<%=cliente__nome%>"/>

        <%else%>
            <table width="649" class="<%=estilo_superior_entrega%> Habilitar_EndEtg_outroendereco" cellspacing="0">
				<%
					estilo_superior_entrega = "QS"
				%>
	            <tr>
		            <td align="left">
		            <p class="R">ENDERE�O DE ENTREGA</p><p class="C">
                        <%
                            s = " "
                            if r_pedido.EndEtg_tipo_pessoa = ID_PJ then
                                s = " checked "
                            end if
                        %>
			            <input type="radio" id="EndEtg_tipo_pessoa_PJ" name="EndEtg_tipo_pessoa" value="PJ" onclick="trocarEndEtgTipoPessoa(null);" <%=s%> >
			            <span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PJ');">Pessoa Jur�dica</span>
			            &nbsp;
                        <%
                            s = " "
                            if r_pedido.EndEtg_tipo_pessoa = ID_PF then
                                s = " checked "
                            end if
                        %>
			            <input type="radio" id="EndEtg_tipo_pessoa_PF" name="EndEtg_tipo_pessoa" value="PF" onclick="trocarEndEtgTipoPessoa(null);" <%=s%> >
			            <span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PF');">Pessoa F�sica</span>
		            </p>
		            </td>
	            </tr>
            </table>

			<!-- ************   PJ: CNPJ/CONTRIBUINTE ICMS/IE - DO ENDERE�O DE ENTREGA DE PJ ************ -->
			<!-- ************   PF: CPF/PRODUTOR RURAL/CONTRIBUINTE ICMS/IE - DO ENDERE�O DE ENTREGA DE PJ  ************ -->
			<!-- fizemos dois conjuntos diferentes de campos porque a ordem � muito diferente -->
			<input type="hidden" id="EndEtg_cnpj_cpf" name="EndEtg_cnpj_cpf" />
			<input type="hidden" id="EndEtg_ie" name="EndEtg_ie" />
			<input type="hidden" id="EndEtg_contribuinte_icms_status" name="EndEtg_contribuinte_icms_status" />
			<input type="hidden" id="EndEtg_rg" name="EndEtg_rg" value="<%=cliente__rg%>"/>
			<input type="hidden" id="EndEtg_produtor_rural_status" name="EndEtg_produtor_rural_status" />

            <table width="649" class="QS Habilitar_EndEtg_outroendereco Mostrar_EndEtg_pj" cellspacing="0">
	            <tr>
		            <td width="210" align="left">
	            <p class="R">CNPJ</p><p class="C">

	            <input id="EndEtg_cnpj_cpf_PJ" name="EndEtg_cnpj_cpf_PJ" class="TA" value="<%=r_pedido.EndEtg_cnpj_cpf%>" size="22" style="text-align:center; color:#0000ff"></p></td>

	            <td class="MDE" width="215" align="left"><p class="R">IE</p><p class="C">
		            <input id="EndEtg_ie_PJ" name="EndEtg_ie_PJ" class="TA" type="text" maxlength="20" size="25" value="<%=r_pedido.EndEtg_ie%>" onkeypress="if (digitou_enter(true)) fPED.EndEtg_nome.focus(); filtra_nome_identificador();"></p></td>

	            <td align="left" class="Mostrar_EndEtg_contribuinte_icms_PJ"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
                    <%
                        s = " "
                        if r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
                            s = " checked "
                        end if
                    %>
		            <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PJ_nao" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">N�o</span>
                    <%
                        s = " "
                        if r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                            s = " checked "
                        end if
                    %>
		            <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PJ_sim" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>');">Sim</span>
                    <%
                        s = " "
                        if r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
                            s = " checked "
                        end if
                    %>
		            <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PJ_isento" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>');">Isento</span></p></td>
	            </tr>
            </table>

            <table width="649" class="QS Habilitar_EndEtg_outroendereco Mostrar_EndEtg_pf" cellspacing="0">
	            <tr>
		            <td width="210" align="left">
	            <p class="R">CPF</p><p class="C">
	            <input id="EndEtg_cnpj_cpf_PF" name="EndEtg_cnpj_cpf_PF" class="TA" value="<%=r_pedido.EndEtg_cnpj_cpf%>" size="22" style="text-align:center; color:#0000ff"></p></td>

	            <td align="left" class="ME" style="min-width: 110px;" ><p class="R">PRODUTOR RURAL</p><p class="C">
                    <%
                        s = " "
                        if r_pedido.EndEtg_produtor_rural_status = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
                            s = " checked "
                        end if
                    %>
		            <input type="radio"  <%=s%> id="EndEtg_produtor_rural_status_PF_nao" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>');">N�o</span>
                    <%
                        s = " "
                        if r_pedido.EndEtg_produtor_rural_status = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) then
                            s = " checked "
                        end if
                    %>
		            <input type="radio"  <%=s%> id="EndEtg_produtor_rural_status_PF_sim" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>')">Sim</span></p></td>

	            <td align="left" class="MDE Mostrar_EndEtg_contribuinte_icms_PF"><p class="R">IE</p><p class="C">
		            <input id="EndEtg_ie_PF" name="EndEtg_ie_PF" class="TA" type="text" maxlength="20" size="13" value="<%=r_pedido.EndEtg_ie%>" onkeypress="if (digitou_enter(true)) fPED.EndEtg_nome.focus(); filtra_nome_identificador();"></p>
	            </td>

	            <td align="left" class="Mostrar_EndEtg_contribuinte_icms_PF" ><p class="R">CONTRIBUINTE ICMS</p><p class="C">
                    <%
                        s = " "
                        if r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
                            s = " checked "
                        end if
                    %>
		            <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PF_nao" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">N�o</span>
                    <%
                        s = " "
                        if r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                            s = " checked "
                        end if
                    %>
		            <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PF_sim" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>');">Sim</span>
                    <%
                        s = " "
                        if r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
                            s = " checked "
                        end if
                    %>
		            <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PF_isento" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>');">Isento</span></p>
	            </td>
	            </tr>
            </table>



            <!-- ************   ENDERE�O DE ENTREGA: NOME  ************ -->
            <table width="649" class="QS" cellspacing="0">
	            <tr>
	            <td width="100%" align="left"><p class="R" id="Label_EndEtg_nome">RAZ�O SOCIAL</p><p class="C">
		            <input id="EndEtg_nome" name="EndEtg_nome" class="TA" value="<%=r_pedido.EndEtg_nome%>" maxlength="60" size="85" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_endereco.focus(); filtra_nome_identificador();"></p></td>
	            </tr>
            </table>

        <%end if%>
    <%end if%> <% 'blnUsarMemorizacaoCompletaEnderecos %>



    <table width="649" class="<%=estilo_superior_entrega%>" cellspacing="0">
		<%
			estilo_superior_entrega = "QS"
		%>
	    <tr>
            <%
                s = "ENDERE�O"
                if eh_cpf then
                    s = "ENDERE�O DE ENTREGA"
                    end if
                %>
		    <td colspan="2" class="MB" align="left"><p class="Rf"><%=s%></p><input id="EndEtg_endereco" name="EndEtg_endereco" class="TA" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_endereco_numero.focus(); filtra_nome_identificador();" value="<%=r_pedido.EndEtg_endereco%>"></td>
	    </tr>
	    <tr>
		    <td class="MDB" align="left"><p class="Rf">N�</p><input id="EndEtg_endereco_numero" name="EndEtg_endereco_numero" class="TA" maxlength="<%=MAX_TAMANHO_CAMPO_ENDERECO_NUMERO%>" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_endereco_complemento.focus(); filtra_nome_identificador();" value="<%=r_pedido.EndEtg_endereco_numero%>"></td>
		    <td class="MB" align="left"><p class="Rf">COMPLEMENTO</p><input id="EndEtg_endereco_complemento" name="EndEtg_endereco_complemento" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_bairro.focus(); filtra_nome_identificador();" value="<%=r_pedido.EndEtg_endereco_complemento%>"></td>
	    </tr>
	    <tr>
		    <td class="MDB" align="left"><p class="Rf">BAIRRO</p><input id="EndEtg_bairro" name="EndEtg_bairro" class="TA" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_cidade.focus(); filtra_nome_identificador();" value="<%=r_pedido.EndEtg_bairro%>"></td>
		    <td class="MB" align="left"><p class="Rf">CIDADE</p><input id="EndEtg_cidade" name="EndEtg_cidade" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_uf.focus(); filtra_nome_identificador();" value="<%=r_pedido.EndEtg_cidade%>"></td>
	    </tr>
	    <tr>
		    <td width="50%" class="MD" align="left"><p class="Rf">UF</p><input id="EndEtg_uf" name="EndEtg_uf"class="TA" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fPED.EndEtg_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inv�lida!!');this.focus();} else this.value=ucase(this.value);" value="<%=r_pedido.EndEtg_uf%>"></td>
		    <td align="left">
			    <table width="100%" cellspacing="0" cellpadding="0">
			    <tr>
			    <td width="50%" align="left"><p class="Rf">CEP</p><input id="EndEtg_cep" name="EndEtg_cep" readonly tabindex=-1 class="TA" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inv�lido!!');this.focus();} else this.value=cep_formata(this.value);" value='<%=cep_formata(r_pedido.EndEtg_cep)%>'></td>
			    <td align="center">
				    <% if blnPesquisaCEPAntiga then %>
				    <button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCepEndEtg();">Pesquisar CEP</button>
				    <% end if %>
				    <% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
				    <% if blnPesquisaCEPNova then %>
				    <button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="exibeJanelaCEP_Etg();">Pesquisar CEP</button>
				    <% end if %>
				    <a name="bLimparEndEtg" id="bLimparEndEtg" href="javascript:LimparCamposEndEtg(fPED)" title="limpa o endere�o de entrega">
					    <img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
			    </td>
			    </tr>
			    </table>
		    </td>
	    </tr>
    </table>


    <% if blnUsarMemorizacaoCompletaEnderecos then %>
        <%if eh_cpf then%>

            <!-- ************   ENDERE�O DE ENTREGA PARA PF: TELEFONES   ************ -->
			<!-- Pegamos todos os atuais. Sem campos edit�veis. -->
            <input type="hidden" id="EndEtg_ddd_res" name="EndEtg_ddd_res" value="<%=r_pedido.EndEtg_ddd_res%>"/>
            <input type="hidden" id="EndEtg_tel_res" name="EndEtg_tel_res" value="<%=r_pedido.EndEtg_tel_res%>"/>
            <input type="hidden" id="EndEtg_ddd_cel" name="EndEtg_ddd_cel" value="<%=r_pedido.EndEtg_ddd_cel%>"/>
            <input type="hidden" id="EndEtg_tel_cel" name="EndEtg_tel_cel" value="<%=r_pedido.EndEtg_tel_cel%>"/>
            <input type="hidden" id="EndEtg_ddd_com" name="EndEtg_ddd_com" value="<%=r_pedido.EndEtg_ddd_com%>"/>
            <input type="hidden" id="EndEtg_tel_com" name="EndEtg_tel_com" value="<%=r_pedido.EndEtg_tel_com%>"/>
            <input type="hidden" id="EndEtg_ramal_com" name="EndEtg_ramal_com" value="<%=r_pedido.EndEtg_ramal_com%>"/>
            <input type="hidden" id="EndEtg_ddd_com_2" name="EndEtg_ddd_com_2" value="<%=r_pedido.EndEtg_ddd_com_2%>"/>
            <input type="hidden" id="EndEtg_tel_com_2" name="EndEtg_tel_com_2" value="<%=r_pedido.EndEtg_tel_com_2%>"/>
            <input type="hidden" id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" value="<%=r_pedido.EndEtg_ramal_com_2%>"/>

        <%else%>
        
            <!-- ************   ENDERE�O DE ENTREGA: TELEFONE RESIDENCIAL   ************ -->
            <table width="649" class="QS Mostrar_EndEtg_pf Habilitar_EndEtg_outroendereco" cellspacing="0">
	            <tr>
	            <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		            <input id="EndEtg_ddd_res" name="EndEtg_ddd_res" class="TA" value="<%=r_pedido.EndEtg_ddd_res%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	            <td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
		            <input id="EndEtg_tel_res" name="EndEtg_tel_res" class="TA" value="<%=r_pedido.EndEtg_tel_res%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	            </tr>
	            <tr>
	            <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		            <input id="EndEtg_ddd_cel" name="EndEtg_ddd_cel" class="TA" value="<%=r_pedido.EndEtg_ddd_cel%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	            <td align="left" class="MC"><p class="R">CELULAR</p><p class="C">
		            <input id="EndEtg_tel_cel" name="EndEtg_tel_cel" class="TA" value="<%=r_pedido.EndEtg_tel_cel%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_email.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('N�mero de celular inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	            </tr>
            </table>
	
            <!-- ************   ENDERE�O DE ENTREGA: TELEFONE COMERCIAL   ************ -->
            <table width="649" class="QS Mostrar_EndEtg_pj Habilitar_EndEtg_outroendereco" cellspacing="0">
	            <tr>
	            <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		            <input id="EndEtg_ddd_com" name="EndEtg_ddd_com" class="TA" value="<%=r_pedido.EndEtg_ddd_com%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	            <td class="MD" align="left"><p class="R">TELEFONE </p><p class="C">
		            <input id="EndEtg_tel_com" name="EndEtg_tel_com" class="TA" value="<%=r_pedido.EndEtg_tel_com%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	            <td align="left"><p class="R">RAMAL</p><p class="C">
		            <input id="EndEtg_ramal_com" name="EndEtg_ramal_com" class="TA" value="<%=r_pedido.EndEtg_ramal_com%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fPED.EndEtg_ddd_com_2.focus(); filtra_numerico();"></p></td>
	            </tr>
	            <tr>
	                <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
	                <input id="EndEtg_ddd_com_2" name="EndEtg_ddd_com_2" class="TA" value="<%=r_pedido.EndEtg_ddd_com_2%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!!');this.focus();}" /></p>  
	                </td>
	                <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
	                <input id="EndEtg_tel_com_2" name="EndEtg_tel_com_2" class="TA" value="<%=r_pedido.EndEtg_tel_com_2%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	                </td>
	                <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
	                <input id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" class="TA" value="<%=r_pedido.EndEtg_ramal_com_2%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fPED.EndEtg_email.focus(); filtra_numerico();" /></p>
	                </td>
	            </tr>
            </table>

        <% end if %>

		<!-- ************   E-MAIL   ************ -->
		<table width="649" class="QS" cellspacing="0">
			<tr>
			<td width="100%" align="left"><p class="R">E-MAIL</p><p class="C">
				<input id="EndEtg_email" name="EndEtg_email" class="TA" value="<%=r_pedido.EndEtg_email%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fPED.EndEtg_email_xml.focus(); filtra_email();"></p></td>
			</tr>
		</table>

		<!-- ************   E-MAIL (XML)  ************ -->
		<table width="649" class="QS" cellspacing="0">
			<tr>
			<td width="100%" align="left"><p class="R">E-MAIL (XML)</p><p class="C">
				<input id="EndEtg_email_xml" name="EndEtg_email_xml" class="TA" value="<%=r_pedido.EndEtg_email_xml%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fPED.EndEtg_obs.focus(); filtra_email();"></p></td>
			</tr>
		</table>

    <%end if%> <% 'blnUsarMemorizacaoCompletaEnderecos %>


    <!-- ************   JUSTIFIQUE O ENDERE�O   ************ -->
    <table id="obs_endereco" width="649" class="QS" cellspacing="0">
	    <tr >
	    <td class="M" width="50%" align="left"><p class="R">JUSTIFIQUE O ENDERE�O</p><p class="C">
		    <select id="EndEtg_obs" name="EndEtg_obs" style="margin-right:225px;">			
                   <option value="">&nbsp;</option>	
			     <%=justificativa_endereco_etg_monta_itens(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA, r_pedido.EndEtg_cod_justificativa)%>
		    </select></p></td>
	    </tr>
    </table>


<% end if %>
<!--  R E L A � � O   D E   P R O D U T O S  -->
<br>
<br>
<table class="Qx" cellspacing="0">
	<%
	' Para assegurar a consist�ncia entre o valor total de NF e o total da forma de pagamento,
	' a edi��o fica permitida somente se o usu�rio puder editar os valores na forma de pagamento!
	if (nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL) And blnItemPedidoEdicaoLiberada then
	%>
	<tr bgColor="#FFFFFF">
	<% if blnTemRA Or (r_pedido.permite_RA_status = 1) then nColSpan=6 else nColSpan=5 %>
	<td colspan="<%=CStr(nColSpan)%>" align="left">&nbsp;</td>
	<td colspan="2" align="right"><span class="PLTe">Desc Linear (%)&nbsp;<input name="c_desc_linear" id="c_desc_linear" class="Cd" style="width:36px;" 
		onkeypress="if (digitou_enter(true)){this.value=formata_perc_desc_linear(this.value);fPED.btnDescLinear.focus();} filtra_percentual();"
		onblur="this.value=formata_perc_desc_linear(this.value);"
		/></span></td>
	<td colspan="2" align="left"><input type="button" name="btnDescLinear" id="btnDescLinear" class="Button" onclick="atualiza_itens_com_desc_linear();" value="Aplicar" title="aplicar o desconto em todos os itens" style="margin-left:1px;margin-bottom:2px;" /></td>
	</tr>
	<tr bgColor="#FFFFFF">
	<% if blnTemRA Or (r_pedido.permite_RA_status = 1) then nColSpan=10 else nColSpan=9 %>
	<td colspan="<%=CStr(nColSpan)%>" align="left" style="height:6px;"></td>
	</tr>
	<% end if %>
	<tr bgColor="#FFFFFF">
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Produto</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Descri��o</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Qtd</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Falt</span></td>
	<% if blnTemRA Or (r_pedido.permite_RA_status = 1) then %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Pre�o</span></td>
	<% end if %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Lista</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Desc</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Venda</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Total</span></td>
	</tr>

<% m_TotalDestePedido=0
   m_TotalDestePedidoComRA=0
   m_total_RA_deste_pedido=0
   m_total_venda_deste_pedido=0
   m_total_NF_deste_pedido=0
   n = Lbound(v_item)-1
   s_readonly_RT = "readonly tabindex=-1"
   if bln_RT_EdicaoLiberada then s_readonly_RT = ""

   for i=1 to max_qtde_itens
	 s_readonly = "readonly tabindex=-1"
	 s_readonly_RA = "readonly tabindex=-1"
	 n = n+1
	 s_cor = "black"
	 if n <= Ubound(v_item) then
		with v_item(n)
			s_fabricante=.fabricante
			s_produto=.produto
			s_descricao=.descricao
			s_descricao_html=produto_formata_descricao_em_html(.descricao_html)
			s_qtde=.qtde
			s_preco_lista=formata_moeda(.preco_lista)
			if .desc_dado=0 then s_desc_dado="" else s_desc_dado=formata_perc(.desc_dado)
			s_vl_unitario=formata_moeda(.preco_venda)
			s_preco_NF=formata_moeda(.preco_NF)
			m_TotalItem=.qtde * .preco_venda
			m_TotalItemComRA=.qtde * .preco_NF
			s_vl_TotalItem=formata_moeda(m_TotalItem)
			m_TotalDestePedido=m_TotalDestePedido + m_TotalItem
			m_TotalDestePedidoComRA=m_TotalDestePedidoComRA + m_TotalItemComRA
			
			m_total_RA_deste_pedido = m_total_RA_deste_pedido + (.qtde * (.preco_NF - .preco_venda))
			m_total_venda_deste_pedido = m_total_venda_deste_pedido + (.qtde * .preco_venda)
			m_total_NF_deste_pedido = m_total_NF_deste_pedido + (.qtde * .preco_NF)
			
			' Para assegurar a consist�ncia entre o valor total de NF e o total da forma de pagamento,
			' a edi��o fica permitida somente se o usu�rio puder editar os valores na forma de pagamento!
			if nivelEdicaoFormaPagto >= COD_NIVEL_EDICAO_LIBERADA_PARCIAL then
				if blnItemPedidoEdicaoLiberada then s_readonly = ""
				if bln_RA_EdicaoLiberada And (r_pedido.permite_RA_status = 1) then s_readonly_RA = ""
				end if
			end with
			
		s_falta=""
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then
			with v_disp(n)
				if .qtde_estoque_sem_presenca<>0 then s_falta=Cstr(.qtde_estoque_sem_presenca)
				s_cor = x_cor_item(.qtde, .qtde_estoque_vendido, .qtde_estoque_sem_presenca)
				end with
			end if
			
	 else
		s_fabricante=""
		s_produto=""
		s_descricao=""
		s_descricao_html=""
		s_qtde=""
		s_falta=""
		s_preco_lista=""
		s_desc_dado=""
		s_vl_unitario=""
		s_preco_NF=""
		s_vl_TotalItem=""
		end if
%>
	<tr>
	<td class="MDBE" align="left"><input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:25px; color:<%=s_cor%>"
		value='<%=s_fabricante%>' readonly tabindex=-1></td>
	<td class="MDB" align="left"><input name="c_produto" id="c_produto" class="PLLe" style="width:54px; color:<%=s_cor%>"
		value='<%=s_produto%>' readonly tabindex=-1></td>
	<td class="MDB" style="width:269px;" align="left">
		<span class="PLLe" style="color:<%=s_cor%>"><%=s_descricao_html%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value='<%=s_descricao%>'>
	</td>
	<td class="MDB" align="right"><input name="c_qtde" id="c_qtde" class="PLLd" style="width:21px; color:<%=s_cor%>"
		value='<%=s_qtde%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_qtde_falta" id="c_qtde_falta" class="PLLd" style="width:20px; color:<%=s_cor%>"
		value='<%=s_falta%>' readonly tabindex=-1></td>
	<% if blnTemRA Or (r_pedido.permite_RA_status = 1) then %>
	<td class="MDB" align="right"><input name="c_vl_NF" id="c_vl_NF" class="PLLd" style="width:62px; color:<%=s_cor%>"
		onkeypress="if (digitou_enter(true)) fPED.c_vl_unitario[<%=Cstr(i-1)%>].focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); trata_edicao_RA(<%=Cstr(i-1)%>); recalcula_RA();recalcula_RA_Liquido();"
		value='<%=s_preco_NF%>' <%=s_readonly_RA%>></td>
	<% else %>
	<input type="hidden" name='c_vl_NF' id="c_vl_NF" value='<%=s_preco_NF%>'>
	<% end if %>
	<input type="hidden" name='c_vl_NF_original' id="c_vl_NF_original" value='<%=s_preco_NF%>'>
	<td class="MDB" align="right"><input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_preco_lista%>' readonly tabindex=-1></td>
	<input type="hidden" name="c_preco_lista_original" id="c_preco_lista_original" value="<%=s_preco_lista%>" />
	<td class="MDB" align="right"><input name="c_desc" id="c_desc" class="PLLd" style="width:36px; color:<%=s_cor%>"
		value='<%=s_desc_dado%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px; color:<%=s_cor%>"
		onkeypress="if (digitou_enter(true)) {if ((<%=Cstr(i)%>==fPED.c_vl_unitario.length)||(trim(fPED.c_produto[<%=Cstr(i)%>].value)=='')) fPED.c_obs1.focus(); else <% if blnTemRA Or (r_pedido.permite_RA_status = 1) then Response.Write "fPED.c_vl_NF" else Response.Write "fPED.c_vl_unitario"%>[<%=Cstr(i)%>].focus();} filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); trata_edicao_RA(<%=Cstr(i-1)%>); recalcula_total_linha(<%=Cstr(i)%>); recalcula_RA();recalcula_RA_Liquido();"
		value='<%=s_vl_unitario%>' <%=s_readonly%>></td>
	<input type="hidden" name="c_vl_unitario_original" id="c_vl_unitario_original" value='<%=s_vl_unitario%>' />
	<td class="MDB" align="right"><input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px; color:<%=s_cor%>" 
		value='<%=s_vl_TotalItem%>' readonly tabindex=-1></td>
	</tr>
<% next %>

<%
'  O TOTAL DO RA (REPASSE AUTOM�TICO) REFERENTE AOS ITENS DOS OUTROS PEDIDOS DESTA FAM�LIA
   m_total_RA_outros = m_TotalFamiliaParcelaRA - m_total_RA_deste_pedido
   m_total_venda_outros = vl_TotalFamiliaPrecoVenda - m_total_venda_deste_pedido
   m_total_NF_outros = vl_TotalFamiliaPrecoNF - m_total_NF_deste_pedido
%>
	<tr>
	<td colspan="4" align="left">
		<table cellspacing="0" cellpadding="0" width='100%' style="margin-top:4px;">
		<tr>
			<td width="20%" align="left">&nbsp;</td>
			<% if blnTemRA Or (r_pedido.permite_RA_status = 1) then %>
			<td align="right">
			<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
				<tr>
				<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;RA L�quido</span></td>
				<td class="MTBD" align="right"><input name="c_total_RA_Liquido" id="c_total_RA_Liquido" class="PLLd" style="width:70px;color:<%if r_pedido.vl_total_RA_liquido >=0 then Response.Write " green" else Response.Write " red"%>;" 
					value='<%=formata_moeda(r_pedido.vl_total_RA_liquido)%>' readonly tabindex=-1></td>
				</tr>
			</table>
			</td>
			<td align="right">
			<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
				<tr>
				<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;RA Bruto</span></td>
				<td class="MTBD" align="right"><input name="c_total_RA" id="c_total_RA" class="PLLd" style="width:70px;color:<%if m_TotalFamiliaParcelaRA >=0 then Response.Write " green" else Response.Write " red"%>;" 
					value='<%=formata_moeda(m_TotalFamiliaParcelaRA)%>' readonly tabindex=-1></td>
				</tr>
			</table>
			</td>
			<% else %>
			<input type="hidden" name="c_total_RA_Liquido" id="c_total_RA_Liquido" value='<%=formata_moeda(r_pedido.vl_total_RA_liquido)%>'>
			<input type="hidden" name="c_total_RA" id="c_total_RA" value='<%=formata_moeda(m_TotalFamiliaParcelaRA)%>'>
			<% end if %>
			<td align="right">
				<table cellspacing="0" cellpadding="0">
				<tr>
				<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;COM(%)</span></td>
				<td class="MTBD" align="right"><input name="c_perc_RT" id="c_perc_RT" class="PLLd" style="width:30px;color:blue;" 
					value='<%=formata_perc_RT(r_pedido.perc_RT)%>' maxlength="5"
					onkeypress="if (digitou_enter(true)) fPED.c_obs1.focus(); filtra_percentual();"
					onblur="this.value=formata_perc_RT(this.value); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inv�lido!!');this.focus();}"
					<%=s_readonly_RT%>
					></td>
					<input type="hidden" name="c_perc_RT_original" id="c_perc_RT_original" value='<%=formata_perc_RT(r_pedido.perc_RT)%>' />
				</tr>
			</table>
			</td>
		</tr>
		</table>
	</td>
	<% if blnTemRA Or (r_pedido.permite_RA_status = 1) then %>
	<td class="MD" align="left">&nbsp;</td>
	<td class="MDB" align="right">
		<input name="c_total_NF" id="c_total_NF" class="PLLd" style="width:70px;color:blue;" 
				value='<%=formata_moeda(m_TotalDestePedidoComRA)%>' readonly tabindex=-1>
	</td>
	<% else %>
	<td align="left">&nbsp;</td>
	<input type="hidden" name="c_total_NF" id="c_total_NF" value='<%=formata_moeda(m_TotalDestePedidoComRA)%>'>
	<% end if %>

	<td class="MD" align="left">&nbsp;</td>
	<td class="MDB" align="right"><input name="c_desc_medio_total" id="c_desc_medio_total" class="PLLd" style="width:36px;color:blue;" readonly tabindex=-1 /></td>
	<td class="MD" align="left">&nbsp;</td>

	<td class="MDB" align="right"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:blue;" 
		value='<%=formata_moeda(m_TotalDestePedido)%>' readonly tabindex=-1></td>
	</tr>
</table>

<input type="hidden" name="c_total_RA_original" id="c_total_RA_original" value='<%=formata_moeda(m_TotalFamiliaParcelaRA)%>'>
<input type="hidden" name="c_total_RA_base" id="c_total_RA_base" value='<%=formata_moeda(m_total_RA_outros)%>'>
<input type="hidden" name="c_total_venda_original" id="c_total_venda_original" value='<%=formata_moeda(vl_TotalFamiliaPrecoVenda)%>'>
<input type="hidden" name="c_total_venda_base" id="c_total_venda_base" value='<%=formata_moeda(m_total_venda_outros)%>'>
<input type="hidden" name="c_total_NF_base" id="c_total_NF_base" value='<%=formata_moeda(m_total_NF_outros)%>'>
<input type="hidden" name="c_total_devolucoes_NF" id="c_total_devolucoes_NF" value='<%=formata_moeda(vl_TotalFamiliaDevolucaoPrecoNF)%>'>

<% if r_pedido.tipo_parcelamento = 0 then %>
<input type="hidden" name="versao_forma_pagamento" id="versao_forma_pagamento" value='1'>
<!--  TRATA VERS�O ANTIGA DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" style="width:649px;" cellspacing="0">
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observa��es </p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS1);" onblur="this.value=trim(this.value);"
				<% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
				><%=r_pedido.obs_1%></textarea>
		</td>
	</tr>
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">N� Nota Fiscal</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" maxlength="10" style="width:85px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_qtde_parcelas.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
				value='<%=r_pedido.obs_2%>' readonly tabindex=-1>
		</td>
	</tr>
	<tr>
		<td class="MDB" align="left" nowrap><p class="Rf">Parcelas</p>
			<table cellspacing="0" cellpadding="0" width="100%"><tr>
				<td align="left"><input name="c_qtde_parcelas" id="c_qtde_parcelas" class="PLLc" maxlength="2" style="width:60px;" onkeypress="if (digitou_enter(true)) fPED.c_forma_pagto.focus(); filtra_numerico();"
						<% if (nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_BLOQUEADA) then Response.Write " readonly tabindex=-1 " %>
						value='<%if (r_pedido.qtde_parcelas<>0) Or (r_pedido.forma_pagto<>"") then Response.write Cstr(r_pedido.qtde_parcelas)%>'></td>
			</tr></table>
		</td>
		<td class="MDB" align="left" nowrap><p class="Rf">Entrega Imediata</p>
			<% if blnEtgImediataEdicaoLiberada then strDisabled = "" else strDisabled = " disabled" %>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
				<%=strDisabled%>
				value="<%=COD_ETG_IMEDIATA_NAO%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[0].click();">N�o</span>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
				<%=strDisabled%>
				value="<%=COD_ETG_IMEDIATA_SIM%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[1].click();">Sim</span>
		</td>
		<td class="MDB" align="left" nowrap><p class="Rf">Bem de Uso/Consumo</p>
			<% if Not blnBemUsoConsumoEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				<%=strDisabled%>
				value="<%=COD_ST_BEM_USO_CONSUMO_NAO%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[0].click();">N�o</span>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				<%=strDisabled%>
				value="<%=COD_ST_BEM_USO_CONSUMO_SIM%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[1].click();">Sim</span>
		</td>
		<td class="MDB" align="left" valign="top" nowrap><p class="Rf">Instalador Instala&nbsp;</p>
		<% 	if Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_NAO) then
				s = "N�O"
			elseif Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MB tdGarInd" align="left" valign="top" nowrap><p class="Rf">Garantia Indicador</p>
			<% if Not blnGarantiaIndicadorEdicaoLiberada then strDisabled=" DISABLED" else strDisabled=""%>
			<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
				<%=strDisabled%>
				value="<%=COD_GARANTIA_INDICADOR_STATUS__NAO%>" <%if Cstr(r_pedido.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[0].click();">N�o</span>
			<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
				<%=strDisabled%>
				value="<%=COD_GARANTIA_INDICADOR_STATUS__SIM%>" <%if Cstr(r_pedido.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[1].click();">Sim</span>
		</td>
	</tr>
	<% if ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then %>
			<tr>
			    <td class="MC" colspan="2"><p class="Rf">Referente Pedido Bonshop: </p>
			    </td>
			    <td class="MC" colspan="4" align="left">
			        <select id="pedBonshop" name="pedBonshop" style="width: 120px">
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
        strResp = strResp & "<option value='" & r("pedido") & "'"
        if (r("pedido") = r_pedido.pedido_bs_x_at) then
            strResp = strResp & " selected"
        end if
        strResp = strResp & ">"
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
	<tr>
		<td colspan="5" align="left"><p class="Rf">Forma de Pagamento</p>
			<textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
				style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_FORMA_PAGTO);" onblur="this.value=trim(this.value);"
				<% if (nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_BLOQUEADA) then Response.Write " readonly tabindex=-1 " %>
				><%=r_pedido.forma_pagto%></textarea>
		</td>
	</tr>
</table>

<% else %>

<!--  TRATA NOVA VERS�O DA FORMA DE PAGAMENTO   -->
<input type="hidden" name="versao_forma_pagamento" id="versao_forma_pagamento" value='2'>
<br>

	<% if (nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_BLOQUEADA) then %>
		<table class="Q" style="width:649px;" cellspacing="0">
			<tr>
				<td class="MB" align="left"><p class="Rf">Observa��es </p>
					<textarea name="c_obs1" id="c_obs1" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
						style="width:99%;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS1);" onblur="this.value=trim(this.value);"
						<% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
						><%=r_pedido.obs_1%></textarea>
				</td>
			</tr>
            <tr>
		        <td class="MB" align="left"><p class="Rf">Constar na NF</p>
			        <textarea name="c_nf_texto" id="c_nf_texto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_NF_TEXTO_CONSTAR)%>" 
				        style="width:99%;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_NF_TEXTO);" onblur="this.value=trim(this.value);"
				        <% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
                        ><%=r_pedido.NFe_texto_constar%></textarea>
		        </td>
	        </tr>
            <tr>
				<td width="100%">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td class="MB MD" align="left" nowrap width="40%"><p class="Rf">xPed</p>
								<input name="c_num_pedido_compra" id="c_num_pedido_compra" class="PLLe" maxlength="15" style="width:100px;margin-left:2pt;" onkeypress="filtra_nome_identificador();" onblur="this.value=trim(this.value);"
								<% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
									value='<%=r_pedido.NFe_xPed%>'>
							</td>
							<td class="MB" align="left">
								<p class="Rf">Previs�o de Entrega</p>
								<input name="c_data_previsao_entrega" id="c_data_previsao_entrega" class="PLLe" maxlength="10" style="width:90px;margin-left:2pt"
								<% if Not blnEtgImediataEdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
									value="<%=formata_data(r_pedido.PrevisaoEntregaData)%>" />
							</td>
						</tr>
					</table>
				</td>
            </tr>
			<tr>
				<td width="100%">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<% if (loja=NUMERO_LOJA_ECOMMERCE_AR_CLUBE) Or (r_pedido.plataforma_origem_pedido = COD_PLATAFORMA_ORIGEM_PEDIDO__MAGENTO) then %>
							<td class="MD" align="left" valign="top" nowrap width="20%"><p class="Rf">N�mero Magento</p>
								<input name="c_pedido_ac" id="c_pedido_ac" class="PLLe" style="width:90px;margin-left:2pt;" maxlength="9" onkeypress="return SomenteNumero(event)"
								   <%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " readonly tabindex=-1" %>  value='<%=r_pedido.pedido_ac%>'>
							</td>
							<% end if %>
							<td class="MD" align="left" valign="top" nowrap width="20%"><p class="Rf">Entrega Imediata</p>
								<% if blnEtgImediataEdicaoLiberada then strDisabled = "" else strDisabled = " disabled" %>
								<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
									<%=strDisabled%>
									value="<%=COD_ETG_IMEDIATA_NAO%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[0].click();">N�o</span>
								<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
									<%=strDisabled%>
									value="<%=COD_ETG_IMEDIATA_SIM%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[1].click();">Sim</span>
							</td>
							<td class="MD" align="left" valign="top" nowrap width="20%"><p class="Rf">Bem de Uso/Consumo</p>
								<% if Not blnBemUsoConsumoEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
								<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
									<%=strDisabled%>
									value="<%=COD_ST_BEM_USO_CONSUMO_NAO%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[0].click();">N�o</span>
								<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
									<%=strDisabled%>
									value="<%=COD_ST_BEM_USO_CONSUMO_SIM%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[1].click();">Sim</span>
							</td>
							<td class="MD" align="left" valign="top" nowrap width="20%"><p class="Rf">Instalador Instala</p>
							<% 	if Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_NAO) then
									s = "N�O"
								elseif Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
									s = "SIM"
								else
									s = ""
									end if
					
								if s="" then s="&nbsp;"
							%>
							<span class="C" style="margin-top:3px;"><%=s%></span>
							</td>
							<td class="tdGarInd" align="left" valign="top" nowrap width="20%"><p class="Rf">Garantia Indicador</p>
								<% if Not blnGarantiaIndicadorEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
								<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
									<%=strDisabled%>
									value="<%=COD_GARANTIA_INDICADOR_STATUS__NAO%>" <%if Cstr(r_pedido.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[0].click();">N�o</span>
								<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
									<%=strDisabled%>
									value="<%=COD_GARANTIA_INDICADOR_STATUS__SIM%>" <%if Cstr(r_pedido.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[1].click();">Sim</span>
							</td>
						</tr>
					</table>
				</td>
			</tr>

			<% if ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then %>
			<tr>
				<td width="100%">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td class="MC" width="25%"><p class="Rf">Referente Pedido Bonshop: </p>
							</td>
							<td class="MC" align="left">
								<select id="pedBonshop" name="pedBonshop" style="width: 120px;margin:6px 4px 6px 4px;">
									<option value="">&nbsp;</option>
			            <%
			            
    If Not bdd_BS_conecta(cn2) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
    
     sqlString = "SELECT pedido FROM t_PEDIDO" & _
              " INNER JOIN t_CLIENTE ON (t_CLIENTE.id = t_PEDIDO.id_cliente)" & _
              " WHERE t_CLIENTE.cnpj_cpf = '" & r_cliente.cnpj_cpf & "'" & _
              " AND (st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
              " ORDER BY data DESC, pedido"
     set r = cn2.Execute(sqlString)
     strResp = ""
     do while Not r.eof
        strResp = strResp & "<option value='" & r("pedido") & "'"
        if (r("pedido") = r_pedido.pedido_bs_x_at) then
            strResp = strResp & " selected"
        end if
        strResp = strResp & ">"
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
					</table>
				</td>
			</tr>
			<% end if %>

            <% if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then %>
            <tr>
				<td width="100%">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td class="MC MD" align="left" nowrap valign="top" width="40%"><p class="Rf">N� Pedido Marketplace</p>
								<input name="c_numero_mktplace" id="c_numero_mktplace" class="PLLe" maxlength="20" style="width:135px;margin-left:2pt;margin-top:5px;" onkeypress="filtra_nome_identificador();return SomenteNumero(event)" onblur="this.value=trim(this.value);"
								<%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " readonly tabindex=-1" %> value="<%=r_pedido.pedido_bs_x_marketplace%>">
							</td>
							<td class="MC" align="left" nowrap valign="top"><p class="Rf">Origem do Pedido</p>
								<select name="c_origem_pedido" id="c_origem_pedido" style="margin: 3px; 3px; 3px"<%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " disabled tabindex=-1" %>>
									<%=codigo_descricao_monta_itens_select_all(GRUPO_T_CODIGO_DESCRICAO__PEDIDOECOMMERCE_ORIGEM, r_pedido.marketplace_codigo_origem) %>
								</select>
							</td>
						</tr>
					</table>
				</td>
            </tr>
            <% end if %>

			<tr>
				<td width="100%">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td class="MC MD" align="left" nowrap width="33.3%"><p class="Rf">N� Nota Fiscal</p>
								<input name="c_obs2" id="c_obs2" class="PLLe" maxlength="10" style="width:67px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_obs3.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
									value='<%=r_pedido.obs_2%>' readonly tabindex=-1>
							</td>
							<td class="MC MD" align="left" nowrap width="33.3%"><p class="Rf">NF Simples Remessa</p>
								<input name="c_obs3" id="c_obs3" class="PLLe" maxlength="10" style="width:67px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_obs4.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
									value='<%=r_pedido.obs_3%>' readonly tabindex=-1>
							</td>
							<td class="MC" nowrap align="left" width="33.3%"><p class="Rf">NF Entrega Futura</p>
								<input name="c_obs4" id="c_obs4" class="PLLe" maxlength="10" style="width:75px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_obs4.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
									value='<%=r_pedido.obs_4%>' readonly tabindex=-1>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<br>
		<table class="Q" style="width:649px;" cellspacing="0">
		  <tr>
			<td align="left"><p class="Rf">Forma de Pagamento</p></td>
		  </tr>
		  <tr>
			<td align="left">
			  <table width="100%" cellspacing="0" cellpadding="0" border="0">
				<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then %>
				<!--  � VISTA  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="0" border="0">
					  <tr>
						<td align="left"><span class="C">� Vista&nbsp&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.av_forma_pagto)%>)</span></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
				<!--  PARCELA �NICA  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="0" border="0">
					  <tr>
						<td align="left"><span class="C">Parcela �nica:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pu_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pu_forma_pagto)%>)&nbsp;&nbsp;vencendo ap�s&nbsp;<%=formata_inteiro(r_pedido.pu_vencto_apos)%>&nbsp;dias</span></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
				<!--  PARCELADO NO CART�O (INTERNET)  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="0" border="0">
					  <tr>
						<td align="left"><span class="C">Parcelado no Cart�o (internet) em&nbsp;&nbsp;<%=Cstr(r_pedido.pc_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_pedido.pc_valor_parcela)%></span></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
				<!--  PARCELADO NO CART�O (MAQUINETA)  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="0" border="0">
					  <tr>
						<td align="left"><span class="C">Parcelado no Cart�o (maquineta) em&nbsp;&nbsp;<%=Cstr(r_pedido.pc_maquineta_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_pedido.pc_maquineta_valor_parcela)%></span></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
				<!--  PARCELADO COM ENTRADA  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="0" border="0">
					  <tr>
						<td align="left"><span class="C">Entrada:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pce_entrada_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pce_forma_pagto_entrada)%>)</span></td>
					  </tr>
					  <tr>
						<td align="left"><span class="C">Presta��es:&nbsp;&nbsp;<%=formata_inteiro(r_pedido.pce_prestacao_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pce_prestacao_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pce_forma_pagto_prestacao)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=formata_inteiro(r_pedido.pce_prestacao_periodo)%>&nbsp;dias</span></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
				<!--  PARCELADO SEM ENTRADA  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="0" border="0">
					  <tr>
						<td align="left"><span class="C">1� Presta��o:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pse_prim_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pse_forma_pagto_prim_prest)%>)&nbsp;&nbsp;vencendo ap�s&nbsp;<%=formata_inteiro(r_pedido.pse_prim_prest_apos)%>&nbsp;dias</span></td>
					  </tr>
					  <tr>
						<td align="left"><span class="C">Demais Presta��es:&nbsp;&nbsp;<%=Cstr(r_pedido.pse_demais_prest_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pse_demais_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pse_forma_pagto_demais_prest)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=Cstr(r_pedido.pse_demais_prest_periodo)%>&nbsp;dias</span></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<% end if %>
			  </table>
			</td>
		  </tr>
		  <tr>
			<td class="MC" align="left"><p class="Rf">Informa��es Sobre An�lise de Cr�dito</p>
			  <textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
						style="width:642px;margin-left:2pt;"
						readonly tabindex=-1><%=r_pedido.forma_pagto%></textarea>
			</td>
		  </tr>
		</table>
	
	<% else %>
		<!--  EDI��O LIBERADA (TOTAL OU PARCIALMENTE)   -->
		<table class="Q" style="width:649px;" cellspacing="0">
			<tr>
				<td class="MB" align="left"><p class="Rf">Observa��es </p>
					<textarea name="c_obs1" id="c_obs1" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
						style="width:99%;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS1);" onblur="this.value=trim(this.value);"
						<% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
						><%=r_pedido.obs_1%></textarea>
				</td>
			</tr>
            <tr>
		        <td class="MB" align="left"><p class="Rf">Constar na NF</p>
			        <textarea name="c_nf_texto" id="c_nf_texto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_NF_TEXTO_CONSTAR)%>" 
				        style="width:99%;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_NF_TEXTO);" onblur="this.value=trim(this.value);"
				        <% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
                        ><%=r_pedido.NFe_texto_constar%></textarea>
		        </td>
	        </tr>
            <tr>
				<td width="100%">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td class="MB MD" align="left" nowrap width="40%"><p class="Rf">xPed</p>
								<input name="c_num_pedido_compra" id="c_num_pedido_compra" class="PLLe" maxlength="15" style="width:100px;margin-left:2pt;" onkeypress="filtra_nome_identificador();" onblur="this.value=trim(this.value);"
								<% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
									value='<%=r_pedido.NFe_xPed%>'>
							</td>
							<td class="MB" align="left">
								<p class="Rf">Previs�o de Entrega</p>
								<input name="c_data_previsao_entrega" id="c_data_previsao_entrega" class="PLLe" maxlength="10" style="width:90px;margin-left:2pt"
								<% if Not blnEtgImediataEdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
									value="<%=formata_data(r_pedido.PrevisaoEntregaData)%>" />
							</td>
						</tr>
					</table>
				</td>
            </tr>
			<tr>
				<td width="100%">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<%if blnNumPedidoECommerceEdicaoLiberada then %>
							<td class="MD" align="left" valign="top" nowrap width="20%"><p class="Rf">N�mero Magento</p>
								<input name="c_pedido_ac" id="c_pedido_ac" class="PLLe" style="width:90px;margin-left:2pt;" maxlength="9" onkeypress="return SomenteNumero(event)"
									<%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " readonly tabindex=-1" %> value='<%=r_pedido.pedido_ac%>'>
							</td>
							<% end if %>
							<td class="MD" align="left" valign="top" nowrap width="20%"><p class="Rf">Entrega Imediata</p>
								<% if blnEtgImediataEdicaoLiberada then strDisabled = "" else strDisabled = " disabled" %>
								<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
									<%=strDisabled%>
									value="<%=COD_ETG_IMEDIATA_NAO%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[0].click();">N�o</span>
								<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
									<%=strDisabled%>
									value="<%=COD_ETG_IMEDIATA_SIM%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[1].click();">Sim</span>
							</td>
							<td class="MD" align="left" valign="top" nowrap width="20%"><p class="Rf">Bem Uso/Consumo</p>
								<% if Not blnBemUsoConsumoEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
								<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
									<%=strDisabled%>
									value="<%=COD_ST_BEM_USO_CONSUMO_NAO%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[0].click();">N�o</span>
								<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
									<%=strDisabled%>
									value="<%=COD_ST_BEM_USO_CONSUMO_SIM%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[1].click();">Sim</span>
							</td>
							<td class="MD" align="left" valign="top" nowrap width="20%"><p class="Rf">Instalador Instala</p>
							<% 	if Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_NAO) then
									s = "N�O"
								elseif Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
									s = "SIM"
								else
									s = ""
									end if
					
								if s="" then s="&nbsp;"
							%>
							<span class="C" style="margin-top:3px;"><%=s%></span>
							</td>
							<td class="tdGarInd" align="left" valign="top" nowrap width="20%"><p class="Rf">Garantia Indicador</p>
								<% if Not blnGarantiaIndicadorEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
								<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
									<%=strDisabled%>
									value="<%=COD_GARANTIA_INDICADOR_STATUS__NAO%>" <%if Cstr(r_pedido.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[0].click();">N�o</span>
								<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
									<%=strDisabled%>
									value="<%=COD_GARANTIA_INDICADOR_STATUS__SIM%>" <%if Cstr(r_pedido.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[1].click();">Sim</span>
							</td>
						</tr>
					</table>
				</td>
			</tr>

			<% if ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then %>
			<tr>
				<td width="100%">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td class="MC" width="25%"><p class="Rf">Referente Pedido Bonshop: </p>
							</td>
							<td class="MC" colspan="4" align="left">
								<select id="pedBonshop" name="pedBonshop" style="width: 120px;margin:6px 4px 6px 4px;">
									<option value="">&nbsp;</option>
			            <%
			            
    If Not bdd_BS_conecta(cn2) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
    
     sqlString = "SELECT pedido FROM t_PEDIDO" & _
              " INNER JOIN t_CLIENTE ON (t_CLIENTE.id = t_PEDIDO.id_cliente)" & _
              " WHERE t_CLIENTE.cnpj_cpf = '" & r_cliente.cnpj_cpf & "'" & _
              " AND (st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
              " ORDER BY data DESC, pedido"
     set r = cn2.Execute(sqlString)
     strResp = ""
     do while Not r.eof
        strResp = strResp & "<option value='" & r("pedido") & "'"
        if (r("pedido") = r_pedido.pedido_bs_x_at) then
            strResp = strResp & " selected"
        end if
        strResp = strResp & ">"
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
					</table>
				</td>
			</tr>
			<% end if %>

            <% if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then %>
            <tr>
				<td width="100%">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td class="MC MD" align="left" nowrap valign="top" width="40%"><p class="Rf">N� Pedido Marketplace</p>
								<input name="c_numero_mktplace" id="c_numero_mktplace" class="PLLe" maxlength="20" style="width:135px;margin-left:2pt;margin-top:5px;" onkeypress="filtra_nome_identificador();return SomenteNumero(event)" onblur="this.value=trim(this.value);"
								 <%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " readonly tabindex=-1" %> value="<%=r_pedido.pedido_bs_x_marketplace%>">
							</td>
							<td class="MC" align="left" nowrap valign="top"><p class="Rf">Origem do Pedido</p>
								<select name="c_origem_pedido" id="c_origem_pedido" style="margin: 3px; 3px; 3px"<%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " disabled tabindex=-1" %>>
									<%=codigo_descricao_monta_itens_select_all(GRUPO_T_CODIGO_DESCRICAO__PEDIDOECOMMERCE_ORIGEM, r_pedido.marketplace_codigo_origem) %>
								</select>
							</td>
						</tr>
					</table>
				</td>
            </tr>
            <% end if %>

			<tr>
				<td width="100%">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td class="MC MD" align="left" nowrap width="33.3%"><p class="Rf">N� Nota Fiscal</p>
								<input name="c_obs2" id="c_obs2" class="PLLe" maxlength="10" style="width:67px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_obs3.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
									value='<%=r_pedido.obs_2%>' readonly tabindex=-1>
							</td>
							<td class="MC MD" align="left" nowrap width="33.3%"><p class="Rf">NF Simples Remessa</p>
								<input name="c_obs3" id="c_obs3" class="PLLe" maxlength="10" style="width:67px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_obs4.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
									value='<%=r_pedido.obs_3%>' readonly tabindex=-1>
							</td>
							<td class="MC" nowrap align="left" width="33.3%"><p class="Rf">NF Entrega Futura</p>
								<input name="c_obs4" id="c_obs4" class="PLLe" maxlength="10" style="width:75px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_obs4.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
									value='<%=r_pedido.obs_4%>' readonly tabindex=-1>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
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
				<!--  � VISTA  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="1" border="0">
					  <tr>
						<td align="left">
						  <% intIdx = 0 %>
						  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
								value="<%=COD_FORMA_PAGTO_A_VISTA%>"
								<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then Response.Write " checked"%>
								<% if (Cstr(r_pedido.tipo_parcelamento) <> COD_FORMA_PAGTO_A_VISTA) And (nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then Response.Write " disabled"%>
								onclick="recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">� Vista</span>
						</td>
						<td align="left">&nbsp;</td>
						<td align="left">
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then strDisabled="" else strDisabled=" disabled"
							%>
						  <select id="op_av_forma_pagto" name="op_av_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();" <%=strDisabled%>>
							<%	if operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then 
										Response.Write forma_pagto_av_monta_itens_select_incluindo_default(r_pedido.av_forma_pagto)
									else
										Response.Write forma_pagto_av_monta_itens_select(Null)
										end if
								else
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then 
										Response.Write forma_pagto_liberada_av_monta_itens_select_incluindo_default(r_pedido.av_forma_pagto, r_pedido.indicador, r_cliente.tipo)
									else
										Response.Write forma_pagto_liberada_av_monta_itens_select(Null, r_pedido.indicador, r_cliente.tipo)
										end if
									end if
							%>
						  </select>
						</td>
					  </tr>
					</table>
				  </td>
				</tr>
				<!--  PARCELA �NICA  -->
				<tr>
				  <td class="MC" align="left">
					<table cellspacing="0" cellpadding="1" border="0">
					  <tr>
						<td colspan="3" align="left">
						  <% intIdx = intIdx+1 %>
						  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
								value="<%=COD_FORMA_PAGTO_PARCELA_UNICA%>"
								<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then Response.Write " checked"%>
								<% if (Cstr(r_pedido.tipo_parcelamento) <> COD_FORMA_PAGTO_PARCELA_UNICA) And (nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then Response.Write " disabled"%>
								onclick="pu_atualiza_valor();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcela �nica</span>
						</td>
						<td align="left">&nbsp;</td>
						<td align="left">
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then strDisabled="" else strDisabled=" disabled"
							%>
						  <select id="op_pu_forma_pagto" name="op_pu_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();" <%=strDisabled%>>
							<%	if operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then
										Response.Write forma_pagto_da_parcela_unica_monta_itens_select_incluindo_default(r_pedido.pu_forma_pagto)
									else
										Response.Write forma_pagto_da_parcela_unica_monta_itens_select(Null)
										end if
								else
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then
										Response.Write forma_pagto_liberada_da_parcela_unica_monta_itens_select_incluindo_default(r_pedido.pu_forma_pagto, r_pedido.indicador, r_cliente.tipo)
									else
										Response.Write forma_pagto_liberada_da_parcela_unica_monta_itens_select(Null, r_pedido.indicador, r_cliente.tipo)
										end if
									end if
							%>
						  </select>
						  <span style="width:10px;">&nbsp;</span>
						  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
						  ><input name="c_pu_valor" id="c_pu_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pu_vencto_apos.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);recalcula_RA_Liquido();" 
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
								value="<%=formata_moeda(r_pedido.pu_valor)%>"
							<% else %>
								value=""
							<% end if %>
						  ><span style="width:10px;">&nbsp;</span
						  ><span class="C">vencendo ap�s</span
						  ><input name="c_pu_vencto_apos" id="c_pu_vencto_apos" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();"
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
								value="<%=Cstr(r_pedido.pu_vencto_apos)%>"
							<% else %>
								value=""
							<% end if %>
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then s_readonly="" else s_readonly=" readonly tabindex=-1"
							%>
							<%=s_readonly%>
						  ><span class="C">dias</span>
						</td>
					  </tr>
					</table>
				  </td>
				</tr>
				<!--  PARCELADO NO CART�O (INTERNET)  -->
				<% if operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) Or _
						(Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO) Or _
						(Not is_restricao_ativa_forma_pagto(r_pedido.indicador, ID_FORMA_PAGTO_CARTAO, r_cliente.tipo)) then %>
				<tr>
				<% else %>
				<tr style="display:none;">
				<% end if %>
				  <td class="MC" align="left">
					<table cellspacing="0" cellpadding="1" border="0">
					  <tr>
						<td align="left">
						  <% intIdx = intIdx+1 %>
						  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
								value="<%=COD_FORMA_PAGTO_PARCELADO_CARTAO%>"
								<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then Response.Write " checked"%>
								<% if (Cstr(r_pedido.tipo_parcelamento) <> COD_FORMA_PAGTO_PARCELADO_CARTAO) And (nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then Response.Write " disabled"%>
								onclick="recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado no Cart�o (internet)</span>
						</td>
						<td align="left">&nbsp;</td>
						<td align="left">
						  <input name="c_pc_qtde" id="c_pc_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pc_valor.focus(); filtra_numerico();" onblur="pc_calcula_valor_parcela();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();" 
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
								value="<%=Cstr(r_pedido.pc_qtde_parcelas)%>"
							<% else %>
								value=""
							<% end if %>
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then s_readonly="" else s_readonly=" readonly tabindex=-1"
							%>
							<%=s_readonly%>
						  >
						</td>
						<td align="left"><span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span></td>
						<td align="left">
						  <input name="c_pc_valor" id="c_pc_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);recalcula_RA_Liquido();" 
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
								value="<%=formata_moeda(r_pedido.pc_valor_parcela)%>"
							<% else %>
								value=""
							<% end if %>
						  >
						</td>
					  </tr>
					</table>
				  </td>
				</tr>
				<!--  PARCELADO NO CART�O (MAQUINETA)  -->
				<% if operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) Or _
						(Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) Or _
						(Not is_restricao_ativa_forma_pagto(r_pedido.indicador, ID_FORMA_PAGTO_CARTAO_MAQUINETA, r_cliente.tipo)) then %>
				<tr>
				<% else %>
				<tr style="display:none;">
				<% end if %>
				  <td class="MC" align="left">
					<table cellspacing="0" cellpadding="1" border="0">
					  <tr>
						<td align="left">
						  <% intIdx = intIdx+1 %>
						  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
								value="<%=COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA%>"
								<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then Response.Write " checked"%>
								<% if (Cstr(r_pedido.tipo_parcelamento) <> COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) And (nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then Response.Write " disabled"%>
								onclick="recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado no Cart�o (maquineta)</span>
						</td>
						<td align="left">&nbsp;</td>
						<td align="left">
						  <input name="c_pc_maquineta_qtde" id="c_pc_maquineta_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pc_maquineta_valor.focus(); filtra_numerico();" onblur="pc_maquineta_calcula_valor_parcela();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();" 
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
								value="<%=Cstr(r_pedido.pc_maquineta_qtde_parcelas)%>"
							<% else %>
								value=""
							<% end if %>
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then s_readonly="" else s_readonly=" readonly tabindex=-1"
							%>
							<%=s_readonly%>
						  >
						</td>
						<td align="left"><span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span></td>
						<td align="left">
						  <input name="c_pc_maquineta_valor" id="c_pc_maquineta_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);recalcula_RA_Liquido();" 
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
								value="<%=formata_moeda(r_pedido.pc_maquineta_valor_parcela)%>"
							<% else %>
								value=""
							<% end if %>
						  >
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
								<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then Response.Write " checked"%>
								<% if (Cstr(r_pedido.tipo_parcelamento) <> COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) And (nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then Response.Write " disabled"%>
								onclick="pce_preenche_sugestao_intervalo();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado com Entrada</span>
						</td>
					  </tr>
					  <tr>
						<td style="width:60px;" align="left">&nbsp;</td>
						<td align="right"><span class="C">Entrada&nbsp;</span></td>
						<td align="left">
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then strDisabled="" else strDisabled=" disabled"
							%>
						  <select id="op_pce_entrada_forma_pagto" name="op_pce_entrada_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();" <%=strDisabled%>>
							<%	if operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
										Response.Write forma_pagto_da_entrada_monta_itens_select_incluindo_default(r_pedido.pce_forma_pagto_entrada)
									else
										Response.Write forma_pagto_da_entrada_monta_itens_select(Null)
										end if
								else
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
										Response.Write forma_pagto_liberada_da_entrada_monta_itens_select_incluindo_default(r_pedido.pce_forma_pagto_entrada, r_pedido.indicador, r_cliente.tipo)
									else
										Response.Write forma_pagto_liberada_da_entrada_monta_itens_select(Null, r_pedido.indicador, r_cliente.tipo)
										end if
									end if
							%>
						  </select>
						  <span style="width:10px;">&nbsp;</span>
						  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
						  ><input name="c_pce_entrada_valor" id="c_pce_entrada_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.op_pce_prestacao_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);pce_calcula_valor_parcela();recalcula_RA_Liquido();" 
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
								value="<%=formata_moeda(r_pedido.pce_entrada_valor)%>"
							<% else %>
								value=""
							<% end if %>
						  >
						</td>
					  </tr>
					  <tr>
						<td style="width:60px;" align="left">&nbsp;</td>
						<td align="right"><span class="C">Presta��es&nbsp;</span></td>
						<td align="left">
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then strDisabled="" else strDisabled=" disabled"
							%>
						  <select id="op_pce_prestacao_forma_pagto" name="op_pce_prestacao_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();" <%=strDisabled%>>
							<%	if operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
										Response.Write forma_pagto_da_prestacao_monta_itens_select_incluindo_default(r_pedido.pce_forma_pagto_prestacao)
									else
										Response.Write forma_pagto_da_prestacao_monta_itens_select(Null)
										end if
								else
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
										Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select_incluindo_default(r_pedido.pce_forma_pagto_prestacao, r_pedido.indicador, r_cliente.tipo)
									else
										Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select(Null, r_pedido.indicador, r_cliente.tipo)
										end if
									end if
							%>
						  </select>
						  <span style="width:10px;">&nbsp;</span>
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then s_readonly="" else s_readonly=" readonly tabindex=-1"
							%>
						  <input name="c_pce_prestacao_qtde" id="c_pce_prestacao_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pce_prestacao_valor.focus(); filtra_numerico();" onblur="pce_calcula_valor_parcela();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
								value="<%=Cstr(r_pedido.pce_prestacao_qtde)%>"
							<% else %>
								value=""
							<% end if %>
							<%=s_readonly%>
						  ><span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span
						  ><input name="c_pce_prestacao_valor" id="c_pce_prestacao_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pce_prestacao_periodo.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);recalcula_RA_Liquido();" 
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
								value="<%=formata_moeda(r_pedido.pce_prestacao_valor)%>"
							<% else %>
								value=""
							<% end if %>
						  >
						</td>
					  </tr>
					  <tr>
						<td style="width:60px;" align="left">&nbsp;</td>
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then s_readonly="" else s_readonly=" readonly tabindex=-1"
							%>
						<td colspan="2" align="left"><span class="C">Parcelas vencendo a cada</span
						><input name="c_pce_prestacao_periodo" id="c_pce_prestacao_periodo" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();"
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
								value="<%=Cstr(r_pedido.pce_prestacao_periodo)%>"
							<% else %>
								value=""
							<% end if %>
							<%=s_readonly%>
						><span class="C">dias</span
						><span style="width:10px;">&nbsp;</span
						><span class="notPrint"><input name="b_pce_SugereFormaPagto" id="b_pce_SugereFormaPagto" type="button" class="Button" style="visibility:hidden;" onclick="pce_sugestao_forma_pagto();" value="sugest�o autom�tica" title="preenche o campo 'Forma de Pagamento' com uma sugest�o de texto"></span
						></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<!--  PARCELADO SEM ENTRADA  -->
				<tr>
				  <td class="MC" align="left">
					<table cellspacing="0" cellpadding="1" border="0">
					  <tr>
						<td colspan="3" align="left">
						  <% intIdx = intIdx+1 %>
						  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
								value="<%=COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA%>"
								<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then Response.Write " checked"%>
								<% if (Cstr(r_pedido.tipo_parcelamento) <> COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) And (nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_PARCIAL) then Response.Write " disabled"%>
								onclick="pse_preenche_sugestao_intervalo();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado sem Entrada</span>
						</td>
					  </tr>
					  <tr>
						<td style="width:60px;" align="left">&nbsp;</td>
						<td align="right"><span class="C">1� Presta��o&nbsp;</span></td>
						<td align="left">
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then strDisabled="" else strDisabled=" disabled"
							%>
						  <select id="op_pse_prim_prest_forma_pagto" name="op_pse_prim_prest_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();" <%=strDisabled%>>
							<%	if operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
										Response.Write forma_pagto_da_prestacao_monta_itens_select_incluindo_default(r_pedido.pse_forma_pagto_prim_prest)
									else
										Response.Write forma_pagto_da_prestacao_monta_itens_select(Null)
										end if
								else
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
										Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select_incluindo_default(r_pedido.pse_forma_pagto_prim_prest, r_pedido.indicador, r_cliente.tipo)
									else
										Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select(Null, r_pedido.indicador, r_cliente.tipo)
										end if
									end if
							%>
						  </select>
						  <span style="width:10px;">&nbsp;</span>
						  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
						  ><input name="c_pse_prim_prest_valor" id="c_pse_prim_prest_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pse_prim_prest_apos.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); pse_calcula_valor_parcela();recalcula_RA_Liquido();" 
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
								value="<%=formata_moeda(r_pedido.pse_prim_prest_valor)%>"
							<% else %>
								value=""
							<% end if %>
						  ><span style="width:10px;">&nbsp;</span
						  ><span class="C">vencendo ap�s</span
						  ><input name="c_pse_prim_prest_apos" id="c_pse_prim_prest_apos" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.op_pse_demais_prest_forma_pagto.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();"
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
								value="<%=Cstr(r_pedido.pse_prim_prest_apos)%>"
							<% else %>
								value=""
							<% end if %>
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then s_readonly="" else s_readonly=" readonly tabindex=-1"
							%>
							<%=s_readonly%>
						  ><span class="C">dias</span>
						</td>
					  </tr>
					  <tr>
						<td style="width:60px;" align="left">&nbsp;</td>
						<td align="right"><span class="C">Demais Presta��es&nbsp;</span></td>
						<td align="left">
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then strDisabled="" else strDisabled=" disabled"
							%>
						  <select id="op_pse_demais_prest_forma_pagto" name="op_pse_demais_prest_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();" <%=strDisabled%>>
							<%	if operacao_permitida(OP_LJA_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
										Response.Write forma_pagto_da_prestacao_monta_itens_select_incluindo_default(r_pedido.pse_forma_pagto_demais_prest)
									else
										Response.Write forma_pagto_da_prestacao_monta_itens_select(Null)
										end if
								else
									if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
										Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select_incluindo_default(r_pedido.pse_forma_pagto_demais_prest, r_pedido.indicador, r_cliente.tipo)
									else
										Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select(Null, r_pedido.indicador, r_cliente.tipo)
										end if
									end if
							%>
						  </select>
						  <span style="width:10px;">&nbsp;</span>
						  <input name="c_pse_demais_prest_qtde" id="c_pse_demais_prest_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pse_demais_prest_valor.focus(); filtra_numerico();" onblur="pse_calcula_valor_parcela();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
								value="<%=Cstr(r_pedido.pse_demais_prest_qtde)%>"
							<% else %>
								value=""
							<% end if %>
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then s_readonly="" else s_readonly=" readonly tabindex=-1"
							%>
							<%=s_readonly%>
						  >
						  <span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span
						  ><input name="c_pse_demais_prest_valor" id="c_pse_demais_prest_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pse_demais_prest_periodo.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);recalcula_RA_Liquido();" 
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
								value="<%=formata_moeda(r_pedido.pse_demais_prest_valor)%>"
							<% else %>
								value=""
							<% end if %>
						  >
						</td>
					  </tr>
					  <tr>
						<td style="width:60px;" align="left">&nbsp;</td>
						<td colspan="2" align="left"><span class="C">Parcelas vencendo a cada</span
						><input name="c_pse_demais_prest_periodo" id="c_pse_demais_prest_periodo" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();"
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
								value="<%=Cstr(r_pedido.pse_demais_prest_periodo)%>"
							<% else %>
								value=""
							<% end if %>
							<%	'No n�vel COD_NIVEL_EDICAO_LIBERADA_PARCIAL pode-se apenas editar valores, mas n�o a forma de pagamento ou o meio de pagamento
								if nivelEdicaoFormaPagto = COD_NIVEL_EDICAO_LIBERADA_TOTAL then s_readonly="" else s_readonly=" readonly tabindex=-1"
							%>
							<%=s_readonly%>
						><span class="C">dias</span
						><span style="width:10px;">&nbsp;</span
						><span class="notPrint"><input name="b_pse_SugereFormaPagto" id="b_pse_SugereFormaPagto" type="button" class="Button" style="visibility:hidden;" onclick="pse_sugestao_forma_pagto();" value="sugest�o autom�tica" title="preenche o campo 'Forma de Pagamento' com uma sugest�o de texto"></span
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
			  <p class="Rf">Informa��es Sobre An�lise de Cr�dito</p>
				<textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
					style="width:641px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_FORMA_PAGTO);" onblur="this.value=trim(this.value);"
					><%=r_pedido.forma_pagto%></textarea>
			</td>
		  </tr>  
		</table>
	<% end if %>
	
<% end if %>


<!--  STATUS DE PAGAMENTO   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td width="16.67%" class="MD" align="left" valign="bottom"><p class="Rf">Status de Pagto</p></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><p class="Rf">VL Total&nbsp;&nbsp;(Fam�lia)&nbsp;</p></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><p class="Rf">VL Pago&nbsp;</p></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><p class="Rf">VL Devolu��es&nbsp;</p></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><p class="Rf">VL Perdas&nbsp;</p></td>
	<td width="16.65%" align="right" valign="bottom"><p class="Rf">Saldo a Pagar&nbsp;</p></td>
</tr>
<tr>
	<% s_aux = x_status_pagto_cor(st_pagto) 
	   s = Ucase(x_status_pagto(st_pagto)) %>
	<td width="16.67%" class="MD" align="left"><p class="C" style="color:<%=s_aux%>;"><%=s%>&nbsp;</p></td>
	<% s = formata_moeda(vl_TotalFamiliaPrecoNF) %>
	<td width="16.67%" align="right" class="MD"><p class="Cd"><%=s%></p></td>
	<% s = formata_moeda(vl_TotalFamiliaPago) %>
	<td width="16.67%" align="right" class="MD"><p class="Cd" style="color:<%
		if vl_TotalFamiliaPago >= 0 then Response.Write "black" else Response.Write "red" 
		%>;"><%=s%></p></td>
	<% s = formata_moeda(vl_TotalFamiliaDevolucaoPrecoNF) %>
	<td width="16.67%" align="right" class="MD"><p class="Cd"><%=s%></p></td>
	<% s = formata_moeda(vl_total_perdas) %>
	<td width="16.67%" align="right" class="MD"><p class="Cd"><%=s%></p></td>
	<td width="16.65%" align="right"><p class="Cd" style="color:<% 
		if vl_saldo_a_pagar >= 0 then Response.Write "black" else Response.Write "red" 
		%>;"><%=s_vl_saldo_a_pagar%></p></td>
</tr>
</table>


<!--  AN�LISE DE CR�DITO   -->
<% if blnAnaliseCreditoEdicaoLiberada then %>
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<%	s=x_analise_credito(r_pedido.analise_credito)
		if s="" then s="&nbsp;"
	%>
	<td align="left"><p class="Rf">AN�LISE DE CR�DITO</p>
			<%intIdx=0%>
			<input type="radio" id="rb_analise_credito" name="rb_analise_credito" 
				value="<%=COD_AN_CREDITO_PENDENTE_VENDAS%>" <%if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE_VENDAS) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS)%></span>
			<%intIdx=intIdx+1%>
			<input type="radio" id="rb_analise_credito" name="rb_analise_credito" 
				value="<%=COD_AN_CREDITO_PENDENTE%>" <%if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE)%></span>
	</td>
</tr>
</table>
<% else %>
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<%	s=x_analise_credito(r_pedido.analise_credito)
		if s <> "" then
			s_aux=formata_data_e_talvez_hora(r_pedido.analise_credito_data)
			if s_aux <> "" then s = s & " &nbsp; (" & s_aux & ")"
			end if
		if s="" then s="&nbsp;"
	%>
	<td align="left"><p class="Rf">AN�LISE DE CR�DITO</p><p class="C" style="color:<%=x_analise_credito_cor(r_pedido.analise_credito)%>;"><%=s%></p></td>
</tr>
</table>
<% end if %>


<% if s_devolucoes <> "" then %>
<!--  DEVOLU��ES   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td align="left"><p class="Rf" style="color:red;">DEVOLU��O DE MERCADORIAS</p><p class="C"><%=s_devolucoes%></p></td>
</tr>
</table>
<% end if %>


<% if s_perdas <> "" then %>
<!--  PERDAS   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td align="left"><p class="Rf" style="color:red;">PERDAS</p><p class="C"><%=s_perdas%></p></td>
</tr>
</table>
<% end if %>


<% if IsEntregaAgendavel(r_pedido.st_entrega) then %>
<!--  DATA DE COLETA   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<%	s=formata_data(r_pedido.a_entregar_data_marcada)
		if s="" then s="&nbsp;"
	%>
	<td align="left"><p class="Rf">DATA DE COLETA</p><p class="C"><%=s%></p></td>
</tr>
</table>
<% end if %>


<% if r_pedido.transportadora_id <> "" then %>
<!--  TRANSPORTADORA   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<%	s = r_pedido.transportadora_id & " (" & x_transportadora(r_pedido.transportadora_id) & ")"
		if s="" then s="&nbsp;"
	%>
	<td align="left"><p class="Rf">TRANSPORTADORA</p><p class="C"><%=s%></p></td>
	
<!--   FRETES   -->

    <%  s = "SELECT * FROM t_PEDIDO_FRETE WHERE pedido='" & r_pedido.pedido & "' ORDER BY dt_cadastro" 
        x = ""
        intQtdeFrete = 0
        vl_total_frete = 0
        set rs = cn.execute(s)

        do while Not rs.Eof
            frete_transportadora_id = Trim("" & rs("transportadora_id"))
            frete_numero_NF = Trim("" & rs("numero_NF"))
            frete_serie_NF = Trim("" & rs("serie_NF"))
            if frete_numero_NF = "0" then frete_numero_NF = ""
            if frete_serie_NF = "0" then 
                frete_serie_NF = ""
            else
                frete_serie_NF = NFeFormataSerieNF(frete_serie_NF)
            end if
            if intQtdeFrete > 0 then x = x & "</tr><tr>" & chr(13)
            
            x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & UCase(rs("transportadora_id")) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:150px;'><span class='C'>" & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE, rs("codigo_tipo_frete")) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & obtem_apelido_empresa_NFe_emitente(rs("id_nfe_emitente")) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:80px;'><span class='C'>" & frete_numero_NF & "</td>" & chr(13)
                x = x & "<td class='MD MB' align='center' style='width:50px;'><span class='C'>" & frete_serie_NF & "</td>" & chr(13)
                x = x & "<td class='MB' align='right' style='width:97px;padding-right: 5px'><span class='C'>" & formata_moeda(rs("vl_frete")) & "</td>" & chr(13)
            
            
            intQtdeFrete = intQtdeFrete + 1
            vl_total_frete = vl_total_frete + rs("vl_frete")
        rs.MoveNext
        loop
        s = formata_moeda(vl_total_frete) 
    %>

	
	

</tr>
</table>
<br />
<table id="tFretes" width="649" class="Q" cellspacing="0" style="border-bottom:0">
    <tr>
        <td class="MB" align="left" style="width:130px;" colspan="6"><p class="Rf">FRETES</p></td>

    </tr>
    <tr>
        <td class="MD MB" align="center" style="width:130px;"><p class="Rf">TRANSPORTADORA</p></td>
        <td class="MD MB" align="center" style="width:150px;"><p class="Rf">TIPO DE FRETE</p></td>
        <td class="MD MB" align="center" style="width:130px;"><p class="Rf">EMITENTE</p></td>
        <td class="MD MB" align="center" style="width:80px;"><p class="Rf">N�MERO NF</p></td>
        <td class="MD MB" align="center" style="width:80px;"><p class="Rf">S�RIE NF</p></td>
        <td class="MB" align="right" style="width:50px;padding-right: 5px"><p class="Rf">VALOR</p></td>

    </tr>
    <tr>
        <%=x%>
    </tr>
    <tr>
        <td class="MB MD" colspan="5" align="right" valign="bottom"><p class="Cd">TOTAL</p></td>
        <td class="MB" align="right" style="width:65px;padding-right: 5px">
            <p class="Cd"><%=s%></p>
	</td>
    </tr>
</table>
<% end if %>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>
<input type="hidden" name="Verifica_End_Entrega" id="Verifica_End_Entrega" value=''>
<input type="hidden" name="Verifica_num" id="Verifica_num" value=''>
<input type="hidden" name="Verifica_Cidade" id="Verifica_Cidade" value=''>
<input type="hidden" name="Verifica_UF" id="Verifica_UF" value=''>
<input type="hidden" name="Verifica_CEP" id="Verifica_CEP" value=''>
<input type="hidden" name="Verifica_Justificativa" id="Verifica_Justificativa" value=''>

<!-- ************   BOT�ES   ************ -->
<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para p�gina anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDConfirma(fPED)" title="confirma as altera��es">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
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