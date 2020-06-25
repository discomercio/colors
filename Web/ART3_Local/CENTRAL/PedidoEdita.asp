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

'	EXIBIÇÃO DE BOTÕES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True


	dim s, usuario, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    dim url_origem
    url_origem = Trim(Request("url_origem"))

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
	dim i, n, x, s_fabricante, s_produto, s_descricao, s_descricao_html, s_qtde, s_preco_lista, s_desc_dado, s_vl_unitario
	dim s_preco_NF, m_total_NF
	dim m_total_RA_deste_pedido, m_total_venda_deste_pedido, m_total_RA_outros, m_total_venda_outros
	dim s_vl_TotalItem, m_TotalItem, m_TotalDestePedido, m_TotalItemComRA, m_TotalDestePedidoComRA
	dim m_TotalFamiliaParcelaRA
	dim m_total_NF_deste_pedido, m_total_NF_outros
	dim s_readonly, s_readonly_RT, s_readonly_RA, rs, sql, intQtdeFrete
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim r_pedido, v_item, alerta, msg_erro
	alerta=""
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
	else
		if Not le_pedido_item(pedido_selecionado, v_item, msg_erro) then alerta = msg_erro
		end if
	
	dim r_cliente
	set r_cliente = New cl_CLIENTE
	dim xcliente_bd_resultado
	xcliente_bd_resultado = x_cliente_bd(r_pedido.id_cliente, r_cliente)
	
	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
    'Definido em 20/03/2020: para os pedidos criado antes da memorização completa, vamos usar a tela anterior.
    'Não queremos exigir que quem editar o pedido seja obrigado a preenhcer o CNPJ do endereço sde entrega. Então, para
    'um pedido criado sem a memorização, ele continua sempre sem a memorização.
    if r_pedido.st_memorizacao_completa_enderecos = 0 then
        blnUsarMemorizacaoCompletaEnderecos  = false
        end if

	dim eh_cpf
	if len(r_cliente.cnpj_cpf)=11 then eh_cpf=True else eh_cpf=False

    'le as variáveis da origem certa: ou do pedido ou do cliente, todas comecam com cliente__
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


	dim s_aux, s2, s3, s4, r_loja, s_cor, s_falta
	dim v_disp
	dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF
	dim vl_saldo_a_pagar, s_vl_saldo_a_pagar, st_pagto
	dim v_item_devolvido, s_devolucoes, blnHaDevolucoes
	dim v_pedido_perda, s_perdas, vl_total_perdas, vl_total_frete
	dim intIdx
	dim strDisabled
	s_devolucoes = ""
	blnHaDevolucoes = False
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

	'	OBTÉM OS VALORES A PAGAR, JÁ PAGO E O STATUS DE PAGAMENTO (PARA TODA A FAMÍLIA DE PEDIDOS)
		if Not calcula_pagamentos(pedido_selecionado, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		m_TotalFamiliaParcelaRA = vl_TotalFamiliaPrecoNF - vl_TotalFamiliaPrecoVenda
		vl_saldo_a_pagar = vl_TotalFamiliaPrecoNF - vl_TotalFamiliaPago - vl_TotalFamiliaDevolucaoPrecoNF
		s_vl_saldo_a_pagar = formata_moeda(vl_saldo_a_pagar)
	'	VALORES NEGATIVOS REPRESENTAM O 'CRÉDITO' QUE O CLIENTE POSSUI EM CASO DE PEDIDOS CANCELADOS QUE HAVIAM SIDO PAGOS
		if (st_pagto = ST_PAGTO_PAGO) And (vl_saldo_a_pagar > 0) then s_vl_saldo_a_pagar = ""
		
	'	HÁ DEVOLUÇÕES?
		if Not le_pedido_item_devolvido(pedido_selecionado, v_item_devolvido, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		for i=Lbound(v_item_devolvido) to Ubound(v_item_devolvido)
			with v_item_devolvido(i)
				if .produto <> "" then
					blnHaDevolucoes = True
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

	'	HÁ PERDAS?
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
	
    sql = "SELECT * FROM t_COMISSAO_INDICADOR_N4 WHERE (pedido='" & r_pedido.pedido & "')"
    set rs = cn.Execute(sql)
    dim blnIndicadorEdicaoLiberada
    blnIndicadorEdicaoLiberada = False
    if operacao_permitida(OP_CEN_EDITA_PEDIDO_INDICADOR, s_lista_operacoes_permitidas) then
        if r_pedido.st_entrega<>ST_ENTREGA_CANCELADO And rs.Eof then
            blnIndicadorEdicaoLiberada = True
        end if 
    end if
    if rs.State <> 0 then rs.Close

    dim blnNumPedidoECommerceEdicaoLiberada
    blnNumPedidoECommerceEdicaoLiberada=False
    if operacao_permitida(OP_CEN_EDITA_PEDIDO_NUM_PEDIDO_ECOMMERCE, s_lista_operacoes_permitidas) then
        blnNumPedidoECommerceEdicaoLiberada=True
    end if

	dim blnObs1EdicaoLiberada
	blnObs1EdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO_OBS1, s_lista_operacoes_permitidas) then
		blnObs1EdicaoLiberada = True
	elseif operacao_permitida(OP_CEN_EDITA_PEDIDO, s_lista_operacoes_permitidas) then
		if Not blnNFEmitida then blnObs1EdicaoLiberada = True
		end if

	dim blnObs2EdicaoLiberada, blnObs3EdicaoLiberada
	blnObs2EdicaoLiberada = False
	blnObs3EdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO_OBS2, s_lista_operacoes_permitidas) then
		blnObs2EdicaoLiberada = True
		blnObs3EdicaoLiberada = True
	elseif operacao_permitida(OP_CEN_EDITA_PEDIDO, s_lista_operacoes_permitidas) then
		if Not blnPedidoEntregue then
			blnObs2EdicaoLiberada = True
			blnObs3EdicaoLiberada = True
			end if
		end if

	dim blnFormaPagtoEdicaoLiberada
	blnFormaPagtoEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO_FORMA_PAGTO, s_lista_operacoes_permitidas) then
		blnFormaPagtoEdicaoLiberada = True
		end if

	dim blnEntregaImediataEdicaoLiberada
	blnEntregaImediataEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO, s_lista_operacoes_permitidas) then
		if (Not IsPedidoEncerrado(r_pedido.st_entrega)) Or (Trim(r_pedido.obs_2) = "") then blnEntregaImediataEdicaoLiberada = True
		end if

	dim strPercLimiteRASemDesagio, strPercDesagio
	if alerta = "" then
		strPercLimiteRASemDesagio = formata_perc(r_pedido.perc_limite_RA_sem_desagio)
		strPercDesagio = formata_perc(r_pedido.perc_desagio_RA)
		end if

'	SE FOI APLICADO DESÁGIO E HOUVE ALGUM PEDIDO ENTREGUE NESTA FAMÍLIA
'	DE PEDIDOS, ENTÃO É OBRIGATÓRIO QUE O DESÁGIO SEJA MANTIDO DAQUI P/ FRENTE.
	dim strOpcaoForcaDesagio, qtde_pedidos_entregues
	strOpcaoForcaDesagio = "N"
	qtde_pedidos_entregues = familia_pedidos_qtde_pedidos_entregues(pedido_selecionado)
	if (CStr(r_pedido.st_tem_desagio_RA)<>CStr(0)) And (qtde_pedidos_entregues > 0) then strOpcaoForcaDesagio = "S"
	
	dim strTextoIndicador
	dim r_orcamentista_e_indicador
	if alerta = "" then
		call le_orcamentista_e_indicador(r_pedido.indicador, r_orcamentista_e_indicador, msg_erro)
		end if
	
	dim blnEndEntregaEdicaoLiberada
	blnEndEntregaEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO, s_lista_operacoes_permitidas) then
		if r_pedido.obs_2 = "" then blnEndEntregaEdicaoLiberada = True
		end if

	dim blnTransportadoraEdicaoLiberada
	blnTransportadoraEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO_TRANSPORTADORA, s_lista_operacoes_permitidas) then
		if (Not IsPedidoEncerrado(r_pedido.st_entrega)) And (r_pedido.transportadora_id <> "") then blnTransportadoraEdicaoLiberada = True
		end if

	dim blnValorFreteEdicaoLiberada
	blnValorFreteEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO_VALOR_FRETE, s_lista_operacoes_permitidas) then
		if r_pedido.transportadora_id <> "" then blnValorFreteEdicaoLiberada = True
		end if
		
	dim blnInstaladorInstalaEdicaoLiberada
	blnInstaladorInstalaEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO, s_lista_operacoes_permitidas) then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then blnInstaladorInstalaEdicaoLiberada = True
		end if
	
	dim blnBemUsoConsumoEdicaoLiberada
	blnBemUsoConsumoEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO, s_lista_operacoes_permitidas) then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then blnBemUsoConsumoEdicaoLiberada = True
		end if
	
	dim blnAEntregarStatusEdicaoLiberada
	blnAEntregarStatusEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO, s_lista_operacoes_permitidas) then
		if IsEntregaAgendavel(r_pedido.st_entrega) then blnAEntregarStatusEdicaoLiberada = True
		end if
	
	dim bln_RT_e_RA_EdicaoLiberada
	bln_RT_e_RA_EdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_RT_E_RA, s_lista_operacoes_permitidas) then
		if Cstr(r_pedido.comissao_paga) = Cstr(COD_COMISSAO_NAO_PAGA) then bln_RT_e_RA_EdicaoLiberada = True
		end if
	
	dim blnItemPedidoEdicaoLiberada
	blnItemPedidoEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_ITEM_DO_PEDIDO, s_lista_operacoes_permitidas) then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then blnItemPedidoEdicaoLiberada = True
		if (Trim("" & r_pedido.st_entrega) = ST_ENTREGA_ENTREGUE) And (IsMesmoAnoEMes(r_pedido.entregue_data, Date)) then blnItemPedidoEdicaoLiberada = True
		end if
		
	dim blnPedidoRecebidoStatusEdicaoLiberada
	blnPedidoRecebidoStatusEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO_STATUS_PEDIDO_RECEBIDO, s_lista_operacoes_permitidas) then
	  if Cstr(r_pedido.PedidoRecebidoStatus) = Cstr(COD_ST_PEDIDO_RECEBIDO_SIM) then blnPedidoRecebidoStatusEdicaoLiberada = True
	  end if

	dim blnGarantiaIndicadorEdicaoLiberada
	blnGarantiaIndicadorEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO_GARANTIA_INDICADOR, s_lista_operacoes_permitidas) then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then blnGarantiaIndicadorEdicaoLiberada = True
		end if

	dim blnDadosNFeMercadoriasDevolvidasEdicaoLiberada
	blnDadosNFeMercadoriasDevolvidasEdicaoLiberada = False
	if operacao_permitida(OP_CEN_EDITA_PEDIDO_DADOS_NFE_MERCADORIAS_DEVOLVIDAS, s_lista_operacoes_permitidas) then
		blnDadosNFeMercadoriasDevolvidasEdicaoLiberada = True
		end if
	
	dim blnAnaliseCreditoEdicaoLiberada
	blnAnaliseCreditoEdicaoLiberada = False

	dim strScriptJS
	strScriptJS = "<script language='JavaScript' type='text/javascript'>" & chr(13) & _
				  "var PERC_DESAGIO_RA_LIQUIDA_PEDIDO = " & js_formata_numero(r_pedido.perc_desagio_RA_liquida) & ";" & chr(13) & _
				  "</script>" & chr(13)




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________________
' TRANSPORTADORA_MONTA_SELECT
'
function transportadora_monta_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_TRANSPORTADORA ORDER BY id")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & Trim("" & r("id")) & " - " & Trim("" & r("nome"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
	else
		strResp = "<OPTION VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	transportadora_monta_select = strResp
	r.close
	set r=nothing
end function


' _____________________________________________
' TRANSPORTADORA_MONTA_SELECT_SOMENTE_APELIDO
'
function transportadora_monta_select_somente_apelido(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_TRANSPORTADORA ORDER BY id")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & Trim("" & r("id"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
	else
		strResp = "<OPTION VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	transportadora_monta_select_somente_apelido = strResp
	r.close
	set r=nothing
end function

' _____________________________________________
' TIPO_FRETE_MONTA_ITENS_SELECT
'
function tipo_frete_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)
    if id_default = "" Or id_default = null then id_default = COD_TIPO_FRETE__ENTREGA_NORMAL

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE grupo='Pedido_TipoFrete' AND st_inativo=0")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("codigo"))
		if (id_default=x) then
			strResp = strResp & "<option selected"
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
		
	tipo_frete_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' _____________________________________________
' INDICADORES MONTA ITENS SELECT
'
function indicadores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT apelido, razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido <> '" & ID_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FP_TODOS & "' ORDER BY apelido")
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
' JUSTIFICATIVA ENDEREÇO MONTA ITENS SELECT
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
	<title>CENTRAL<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_GLOBAL%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<%=strScriptJS%>

<script type="text/javascript">
	$(function() {
	    $("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8

	    <%if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE_VENDAS) then %>
            $('#trPendVendasMotivo').show();
	    <%else%>
            $('#trPendVendasMotivo').hide();
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

<script language="JavaScript" type="text/javascript">
var objAjaxCustoFinancFornecConsultaPreco;
var blnConfirmaDifRAeValores=false;
var fCepPopup;

$(function() {
    var f;
    f = fPED;
    if (f.blnEndEntregaEdicaoLiberada.value == "<%=Cstr(True)%>") {
    	$("#EndEtg_obs option[value='<%=r_pedido.EndEtg_cod_justificativa%>']").attr("selected", true);
    	// VERIFICAR MUDANÇA NOS CAMPOS
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
	//  SE JÁ HOUVER UMA JANELA DE PESQUISA DE CEP ABERTA, GARANTE QUE ELA SERÁ FECHADA 
	// E UMA NOVA SERÁ CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')	
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
	window.status="Concluído";
}

function processaFormaPagtoDefault() {
var f, i;
	f=fPED;

	// Versão antiga da forma de pagamento?
	if (f.tipo_parcelamento.value=="0") return;
	
//  O pedido foi cadastrado já com a nova política de custo financeiro por fornecedor?
	if (f.c_custoFinancFornecTipoParcelamento.value=="") return;
	if (f.blnFormaPagtoEdicaoLiberada.value != "<%=Cstr(True)%>") return;
	
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

//  O pedido foi cadastrado já com a nova política de custo financeiro por fornecedor?
	if (f.c_custoFinancFornecTipoParcelamento.value=="") return;
	if (f.blnFormaPagtoEdicaoLiberada.value != "<%=Cstr(True)%>") return;

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
									f.c_preco_lista[j].value=strPrecoLista;
									f.c_preco_lista[j].style.color="black";
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
								f.c_preco_lista[j].style.color=COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE;
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
		recalcula_RA_Liquido();
			
		window.status="Concluído";
		$("#divAjaxRunning").hide();
		}
}

function recalculaCustoFinanceiroPrecoLista() {
var f, i, strListaProdutos, strUrl, strOpcaoFormaPagto;
	f=fPED;

//  O pedido foi cadastrado já com a nova política de custo financeiro por fornecedor?
	if (f.c_custoFinancFornecTipoParcelamento.value=="") return;
	if (f.blnFormaPagtoEdicaoLiberada.value != "<%=Cstr(True)%>") return;

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
//  Retorna total de preço NF (tem valor de NF, ou seja, pedido c/ RA)?
	if (mTotNF > 0) {
		return mTotNFBase+mTotNF;
		}
//  Retorna total de preço de venda
	else {
		return mTotNFBase+mTotVenda;
		}
}

// PARCELA ÚNICA
function pu_atualiza_valor( ){
var f,vt;
	f=fPED;
	if (converte_numero(trim(f.c_pu_valor.value))>0) return;
	vt=fp_vl_total_pedido();
	f.c_pu_valor.value=formata_moeda(vt);
}

// PARCELADO NO CARTÃO (INTERNET)
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

// PARCELADO NO CARTÃO (MAQUINETA)
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

function recalcula_total_linha( idx ) {
var idx, m, m_lista, m_unit, d, f, i, s;
	f=fPED;
	idx=parseInt(idx)-1;
	if (f.c_produto[idx].value=="") return;
	m_lista=converte_numero(f.c_preco_lista[idx].value);
	m_unit=converte_numero(f.c_vl_unitario[idx].value);
	if (m_lista==0) d=0; else d=100*(m_lista-m_unit)/m_lista;
	if (d==0) s=""; else s=formata_perc_desc(d);
	if (f.c_desc[idx].value!=s) f.c_desc[idx].value=s;
	s=formata_moeda(parseInt(f.c_qtde[idx].value)*m_unit);
	if (f.c_vl_total[idx].value!=s) f.c_vl_total[idx].value=s;
	m=0;
	for (i=0; i<f.c_vl_total.length; i++) m=m+converte_numero(f.c_vl_total[i].value);
	s=formata_moeda(m);
	if (f.c_total_geral.value!=s) f.c_total_geral.value=s;
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
			if (d==0) s=""; else s=formata_perc_desc(d);
			if (f.c_desc[i].value!=s) f.c_desc[i].value=s;
			m=parseInt(f.c_qtde[i].value)*m_unit;
			f.c_vl_total[i].value=formata_moeda(m);
			t=t+m;
			}
		}
	f.c_total_geral.value=formata_moeda(t);
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
	f=fPED;
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
				alert('Há divergência entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
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
				alert('Há divergência entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
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
				alert('Há divergência entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
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
				alert('Há divergência entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
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
				alert('Há divergência entre o valor total do pedido (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtNF) + ') e o valor total descrito através da forma de pagamento (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vtFP) + ')!!');
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

function LimparCamposEndEtg( f ) {
	f.EndEtg_endereco.value="";
	f.EndEtg_endereco_numero.value="";
	f.EndEtg_endereco_complemento.value="";
	f.EndEtg_bairro.value="";
	f.EndEtg_cidade.value="";
	f.EndEtg_uf.value="";
	f.EndEtg_cep.value="";
	f.EndEtg_obs.selectedIndex = 0;

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

function exibeOcultaPendenteVendasMotivo() {
    if ($('#rb_analise_credito').is(':checked')) {
        $('#trPendVendasMotivo').show();
    }
    else {
        $('#trPendVendasMotivo').hide();
    }
}

function fPEDConfirma( f ) {
var s,i, blnTemEndEntrega,blnTemTransportadora,blnExcFreteMarcado, strMsgErro;
var NUMERO_LOJA_ECOMMERCE_AR_CLUBE = "<%=NUMERO_LOJA_ECOMMERCE_AR_CLUBE%>";

	recalcula_total_todas_linhas();

	if (f.c_loja.value != NUMERO_LOJA_ECOMMERCE_AR_CLUBE) {
	    if (f.c_indicador.value == "") {
	        if(f.c_perc_RT.value != "") {
	            if (parseFloat(f.c_perc_RT.value.replace(',','.')) > 0) {
	                alert('Não é possível gravar o pedido com o campo "Indicador" vazio e "COM(%)" maior do que zero!!');
	                f.c_perc_RT.focus();
	                return;
	            }
	        }	        
	    }
	}

	s = "" + f.c_obs1.value;
	if (s.length > MAX_TAM_OBS1) {
		alert('Conteúdo de "Observações " excede em ' + (s.length-MAX_TAM_OBS1) + ' caracteres o tamanho máximo de ' + MAX_TAM_OBS1 + '!!');
		f.c_obs1.focus();
		return;
	}

	s = "" + f.c_nf_texto.value;
	if (s.length > MAX_TAM_NF_TEXTO) {
	    alert('Conteúdo de "Constar na NF" excede em ' + (s.length-MAX_TAM_NF_TEXTO) + ' caracteres o tamanho máximo de ' + MAX_TAM_NF_TEXTO + '!!');
	    f.c_nf_texto.focus();
	    return;
	}

	s = "" + f.c_forma_pagto.value;
	if (s.length > MAX_TAM_FORMA_PAGTO) {
		alert('Conteúdo de "Forma de Pagamento" excede em ' + (s.length-MAX_TAM_FORMA_PAGTO) + ' caracteres o tamanho máximo de ' + MAX_TAM_FORMA_PAGTO + '!!');
		f.c_forma_pagto.focus();
		return;
		}

//  Consiste a nova versão da forma de pagamento
	if ((f.versao_forma_pagamento.value == '2') && (f.blnFormaPagtoEdicaoLiberada.value == '<%=Cstr(True)%>')) {
		if (!consiste_forma_pagto(true)) return;
		}

	recalcula_RA();
	recalcula_RA_Liquido();

	if (blnConfirmaDifRAeValores) {
		if (f.c_total_RA.value != f.c_total_RA_original.value) {
			if (!confirm("O valor do RA é de " + SIMBOLO_MONETARIO + " " + formata_moeda(converte_numero(f.c_total_RA.value))+"\nContinua?")) return;
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
		if (trim(f.EndEtg_obs.value)!="") blnTemEndEntrega=true;

<%if blnUsarMemorizacaoCompletaEnderecos and not eh_cpf then %>

        if( $('input[name="EndEtg_tipo_pessoa"]:checked').val()) blnTemEndEntrega = true;

        //simplesmente testamos todos os campos, qualquer valor em qq campo significa preenchimento
        //não deve estar em campo oculto porque o usuário deve clicar no X para limpar, e o X limpa todos os campos, inclusive os não visiveis no momento

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
            alert('Endereço não foi preenchido corretamente!!');
            f.endereco__endereco.focus();
            return;
        }
        if (trim(f.endereco__bairro.value) == "") {
            alert('Endereço não foi preenchido corretamente!!');
            f.endereco__bairro.focus();
            return;
        }

        if (trim(f.endereco__numero.value) == "") {
            alert('Endereço não foi preenchido corretamente!!');
            f.endereco__numero.focus();
            return;
        }
        if (trim(f.endereco__cidade.value) == "") {
            alert('Endereço não foi preenchido corretamente!!');
            f.endereco__cidade.focus();
            return;
        }

        if (trim(f.endereco__uf.value) == "") {
            alert('Endereço não foi preenchido corretamente!!');
            f.endereco__uf.focus();
            return;
        }

        if (trim(f.endereco__cep.value) == "") {
            alert('Endereço não foi preenchido corretamente!!');
            f.endereco__cep.focus();
            return;
        }

        if ((trim(f.cliente__email.value) != "") && (!email_ok(f.cliente__email.value))) {
            alert('E-mail inválido!!');
            f.cliente__email.focus();
            return;
        }

        if ((trim(f.cliente__email_xml.value) != "") && (!email_ok(f.cliente__email_xml.value))) {
            alert('E-mail xml inválido!!');
            f.cliente__email_xml.focus();
            return;
        }


       <% if cliente__tipo = ID_PF then %>

		if ( (trim(f.cliente__ddd_res.value) != "" && !ddd_ok(f.cliente__ddd_res.value)) || (trim(f.cliente__ddd_res.value) == "" && trim(f.cliente__tel_res.value) != "") ) {
            alert('DDD inválido!!');
            f.cliente__ddd_res.focus();
            return;
        }

		if ( (trim(f.cliente__tel_res.value) != "" && !telefone_ok(f.cliente__tel_res.value)) || (trim(f.cliente__ddd_res.value) != "" && trim(f.cliente__tel_res.value) == "") ) {
            alert('Telefone residencial inválido!!');
            f.cliente__tel_res.focus();
            return;
        }

		if ( (trim(f.cliente__ddd_cel.value) != "" && !ddd_ok(f.cliente__ddd_cel.value)) || (trim(f.cliente__ddd_cel.value) == "" && trim(f.cliente__tel_cel.value) != "") ) {
            alert('Celular com DDD inválido!!');
            f.cliente__ddd_cel.focus();
            return;
        }

		if ( (trim(f.cliente__tel_cel.value) != "" && !telefone_ok(f.cliente__tel_cel.value)) || (trim(f.cliente__ddd_cel.value) != "" && trim(f.cliente__tel_cel.value) == "") ) {
            alert('Telefone celular inválido!!');
            f.cliente__tel_cel.focus();
            return;
        }


		if ( (trim(f.cliente__ddd_com.value) != "" && !ddd_ok(f.cliente__ddd_com.value)) || (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__tel_com.value) != "") ) {
            alert('DDD comercial inválido!!');
            f.cliente__ddd_com.focus();
            return;
        }

		if ( (trim(f.cliente__tel_com.value) != "" && !telefone_ok(f.cliente__tel_com.value)) || (trim(f.cliente__ddd_com.value) != "" && trim(f.cliente__tel_com.value) == "") ) {
            alert('Telefone comercial inválido!!');
            f.cliente__tel_com.focus();
            return;
        }

		if (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
            alert('DDD comercial inválido!!');
            f.cliente__ddd_com.focus();
            return;
        }

		if (trim(f.cliente__tel_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
            alert('Telefone comercial inválido!!');
            f.cliente__tel_com.focus();
            return;
        }

        if (trim(f.cliente__tel_res.value) == "" && trim(f.cliente__tel_cel.value) == "" && trim(f.cliente__tel_com.value) == "") {
            alert('Necessário preencher ao menos um telefone!!');
            f.cliente__ddd_cel.focus();
            return;
        }



        if (f.rb_produtor_rural[1].checked) {
            if (!f.rb_contribuinte_icms[1].checked) {
                alert('Para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!!');
                return;
            }
            if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
                alert('Informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
                return;
            }
            if ((f.rb_contribuinte_icms[1].checked) && (trim(f.cliente__ie.value) == "")) {
                alert('Se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
                f.cliente__ie.focus();
                return;
            }
            if ((f.rb_contribuinte_icms[0].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                f.cliente__ie.focus();
                return;
            }
            if ((f.rb_contribuinte_icms[1].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                f.cliente__ie.focus();
                return;
            }
            if (f.rb_contribuinte_icms[2].checked) {
                if (f.cliente__ie.value != "") {
                    alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
                    f.cliente__ie.focus();
                    return;
                }
            }
        }


		<% else %>

        if ((trim(f.cliente__email.value) != "") && (!email_ok(f.cliente__email.value))) {
            alert('E-mail inválido!!');
            f.cliente__email.focus();
            return;
        }

        if ((trim(f.cliente__email_xml.value) != "") && (!email_ok(f.cliente__email_xml.value))) {
            alert('E-mail (XML) inválido!!');
            f.cliente__email_xml.focus();
            return;
        }

           <% if CStr(r_pedido.loja) <> CStr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then %>
            // PARA CLIENTE PJ, É OBRIGATÓRIO O PREENCHIMENTO DO E-MAIL
            if ((trim(f.cliente__email.value) == "") && (trim(f.cliente__email_xml.value) == "")) {
                alert("É obrigatório informar um endereço de e-mail");
                f.cliente__email.focus();
                return;
            }
            <% end if %>

        if ((f.rb_contribuinte_icms[1].checked) && (trim(f.cliente__ie.value) == "")) {
            alert('Se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
            f.cliente__ie.focus();
            return;
        }
        if ((f.rb_contribuinte_icms[0].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
            alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
            f.cliente__ie.focus();
            return;
        }
        if ((f.rb_contribuinte_icms[1].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
            alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
            f.cliente__ie.focus();
            return;
        }
        if (f.rb_contribuinte_icms[2].checked) {
            if (f.cliente__ie.value != "") {
                alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
                f.cliente__ie.focus();
                return;
            }
        }
		if ( (trim(f.cliente__ddd_com.value) != "" && !ddd_ok(f.cliente__ddd_com.value)) || (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__tel_com.value) != "") ) {
            alert('DDD comercial inválido!!');
            f.cliente__ddd_com.focus();
            return;
        }

		if (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
            alert('DDD comercial inválido!!');
            f.cliente__ddd_com.focus();
            return;
        }

		if ( (trim(f.cliente__tel_com.value) != "" && !telefone_ok(f.cliente__tel_com.value)) || (trim(f.cliente__ddd_com.value) != "" && trim(f.cliente__tel_com.value) == "") ) {
            alert('Telefone comercial inválido!!');
            f.cliente__tel_com.focus();
            return;
        }

		if (trim(f.cliente__tel_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
            alert('Telefone comercial inválido!!');
            f.cliente__tel_com.focus();
            return;
        }

		if ( (trim(f.cliente__ddd_com_2.value) != "" && !ddd_ok(f.cliente__ddd_com_2.value)) || (trim(f.cliente__ddd_com_2.value) == "" && trim(f.cliente__tel_com_2.value) != "") ) {
            alert('DDD comercial 2 inválido!!');
            f.cliente__ddd_com_2.focus();
            return;
        }

		if (trim(f.cliente__ddd_com_2.value) == "" && trim(f.cliente__ramal_com_2.value) != "") {
            alert('DDD comercial 2 inválido!!');
            f.cliente__ddd_com_2.focus();
            return;
        }

		if ( (trim(f.cliente__tel_com_2.value) != "" && !telefone_ok(f.cliente__tel_com_2.value)) || (trim(f.cliente__ddd_com_2.value) != "" && trim(f.cliente__tel_com_2.value) == "") ) {
            alert('Telefone comercial 2 inválido!!');
            f.cliente__tel_com_2.focus();
            return;
        }

		if (trim(f.cliente__tel_com_2.value) == "" && trim(f.cliente__ramal_com_2.value) != "") {
            alert('Telefone comercial 2 inválido!!');
            f.cliente__tel_com_2.focus();
            return;
        }

        if (trim(f.cliente__tel_com.value) == "" && trim(f.cliente__tel_com_2.value) == "") {
            alert('Necessário preencher ao menos um telefone!!');
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
				alert('Endereço de entrega não foi preenchido corretamente!!');
				f.EndEtg_endereco.focus();
				return;
				}

			if (trim(f.EndEtg_endereco_numero.value)=="") {
				alert('O número do endereço de entrega não foi preenchido corretamente!!');
				f.EndEtg_endereco_numero.focus();
				return;
				}

			if (trim(f.EndEtg_bairro.value)=="") {
				alert('Bairro do endereço de entrega não foi preenchido corretamente!!');
				f.EndEtg_bairro.focus();
				return;
				}
			
			if (trim(f.EndEtg_cidade.value)=="") {
				alert('Cidade do endereço de entrega não foi preenchido corretamente!!');
				f.EndEtg_cidade.focus();
				return;
			    }

			if ((trim(f.EndEtg_obs.value)=="")  && blnEndEtg_obs == true) {
			    alert('Justificativa do endereço de entrega não foi preenchido corretamente!!');
			    f.EndEtg_obs.focus();
			    return;
			    }

			s=trim(f.EndEtg_uf.value);
			if ((s=="")||(!uf_ok(s))) {
				alert('UF do endereço de entrega não foi preenchido corretamente!!');
				f.EndEtg_uf.focus();
				return;
				}
				
			if (!cep_ok(f.EndEtg_cep.value)) {
				alert('CEP do endereço de entrega não foi preenchido corretamente!!');
				f.EndEtg_cep.focus();
				return;
				}



<%if blnUsarMemorizacaoCompletaEnderecos and Not eh_cpf then%>
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

			}
		}

	if (f.blnTransportadoraEdicaoLiberada.value == "<%=Cstr(True)%>") {
		blnTemTransportadora=false;
		if (trim(f.c_transportadora_id.value)!="") blnTemTransportadora=true;

		if (blnTemTransportadora) {
			if (trim(f.c_transportadora_id.value)=="") {
				alert('Informe a transportadora!!');
				f.c_transportadora_id.focus();
				return;
				}
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

	if($('input[name=rb_analise_credito]:checked').val() != $('#c_analise_credito_inicial').val()) {
	    if ($('#rb_analise_credito').is(':checked')) {
	        if($('#c_pendente_vendas_motivo').val() == "") {
	            alert('Selecione o motivo para o status "Pendente Vendas"!!');
	            $('#c_pendente_vendas_motivo').focus();
	            return;
	        }
	    }
	}

	// Percentual máximo de comissão e desconto
	// ========================================
	// Lembretes:
	//	1) Na 'Central', se o usuário possuir as permissões de acesso 'Editar Item do Pedido' e 'Editar RT e RA' poderá editar
	//	   livremente o percentual de RT e o preço de venda (desconto) do produto, já que nenhuma consistência será realizada.
	//	2) Na 'Loja', o percentual de RT pode ser alterado se o usuário possuir a permissão de acesso 'Editar RT', sendo que
	//	   somente a RT deve ser editada para que um percentual qualquer seja aceito sem que a consistência seja realizada.
	//	   Caso o preço de venda (desconto) seja alterado, serão aplicadas as verificações do 'percentual máximo de comissão e desconto'.

    //campos do endereço de entrega que precisam de transformacao
    transferirCamposEndEtg(f);

	f.action="pedidoatualiza.asp";
	dCONFIRMA.style.visibility="hidden";
	

	blnExcFreteMarcado = false;
	for (i = 0; i < f.c_valor_frete.length; i++) {
	    if ($("#ckb_exclui_frete_" + i).is(":checked")) {
	        blnExcFreteMarcado = true;
	        break;
	    }
	}

	if (blnExcFreteMarcado) {
	    if (window.confirm("Tem certeza que deseja excluir o(s) frete(s) marcado(s)?")) {
	        window.status = "Aguarde ...";
	        f.submit();
	    }
	    else {
	        dCONFIRMA.style.visibility="";            
	        return;
        }
	}
	else {
	    window.status = "Aguarde ...";
	    f.submit();
	}
}

function transferirCamposEndEtg(fNEW) {
<%if blnUsarMemorizacaoCompletaEnderecos and Not eh_cpf then%>
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
}

//para mudar o tipo do endereço de entrega
function trocarEndEtgTipoPessoa(novoTipo) {
<%if blnUsarMemorizacaoCompletaEnderecos then%>
    if (novoTipo && $('input[name="EndEtg_tipo_pessoa"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_tipo_pessoa"]'), novoTipo);

    var pf = $('input[name="EndEtg_tipo_pessoa"]:checked').val() == "PF";

    //se nao tiver nada selecionado queremos tratar cono pj
    if (!pf) {
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

function trataProdutorRural() {
    //ao clicar na opção Produtor Rural, exibir/ocultar os campos apropriados
    if (!fPED.rb_produtor_rural[1].checked) {
        $("#t_contribuinte_icms").css("display", "none");
    }
    else {
        $("#t_contribuinte_icms").css("display", "block");
    }
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
    function ConfirmaExclusaoFrete() {

        $("#msgExcFrete").css('display', 'block');
        $("#caixa-confirmacao").dialog({
            resizable: false,
            height: 175,
            width: 500,
            scroll: false,

            modal: true,

            buttons: {
                "Sim": function () {
                    $(this).dialog("close");
                    window.status = "Aguarde ...";
                    f.submit();
                },
                "Não": function () {
                    $(this).dialog("close");
                    $("#msgExcFrete").css('display', 'none');
                    return;
                }
            }
        });
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
<link href="<%=URL_FILE__E_JANELABUSCACEP_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
#rb_etg_imediata, #rb_bem_uso_consumo, #rb_instalador_instala {
	margin: 0pt 2pt 1pt 6pt;
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
<!-- ********************************************************** -->
<!-- **********  PÁGINA PARA EDITAR ITENS DO PEDIDO  ********** -->
<!-- ********************************************************** -->
<body id="corpoPagina" onload="processaFormaPagtoDefault();trataProdutorRural();">
<center>


<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="c_PercLimiteRASemDesagio" id="c_PercLimiteRASemDesagio" value='<%=strPercLimiteRASemDesagio%>'>
<input type="hidden" name="c_PercDesagio" id="c_PercDesagio" value='<%=strPercDesagio%>'>
<input type="hidden" name="c_opcao_forca_desagio" id="c_opcao_forca_desagio" value='<%=strOpcaoForcaDesagio %>'>
<input type="hidden" name="c_qtde_pedidos_entregues" id="c_qtde_pedidos_entregues" value='<%=CStr(qtde_pedidos_entregues)%>'>
<input type="hidden" name="c_qtde_parcelas_desagio_RA" id="c_qtde_parcelas_desagio_RA" value='<%=CStr(r_pedido.qtde_parcelas_desagio_RA)%>'>
<input type="hidden" name="tipo_parcelamento" id="tipo_parcelamento" value='<%=r_pedido.tipo_parcelamento%>'>
<input type="hidden" name="GarantiaIndicadorStatusOriginal" id="GarantiaIndicadorStatusOriginal" value='<%=r_pedido.GarantiaIndicadorStatus%>'>
<input type="hidden" name="c_loja" id="c_loja" value='<%=r_pedido.loja%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamentoOriginal" id="c_custoFinancFornecTipoParcelamentoOriginal" value='<%=r_pedido.custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelasOriginal" id="c_custoFinancFornecQtdeParcelasOriginal" value='<%=r_pedido.custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamento" id="c_custoFinancFornecTipoParcelamento" value='<%=r_pedido.custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelas" id="c_custoFinancFornecQtdeParcelas" value='<%=r_pedido.custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamentoUltConsulta" id="c_custoFinancFornecTipoParcelamentoUltConsulta" value='<%=r_pedido.custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelasUltConsulta" id="c_custoFinancFornecQtdeParcelasUltConsulta" value='<%=r_pedido.custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecParcelamentoDescricao" id="c_custoFinancFornecParcelamentoDescricao" value=''>
<input type="hidden" name="c_ped_bonshop" id="c_ped_bonshop" value='<%=r_pedido.pedido_bs_x_at %>' />
<input type="hidden" name="blnIndicadorEdicaoLiberada" id="blnIndicadorEdicaoLiberada" value='<%=Cstr(blnIndicadorEdicaoLiberada)%>'>
<input type="hidden" name="blnNumPedidoECommerceEdicaoLiberada" id="blnNumPedidoECommerceEdicaoLiberada" value="<%=Cstr(blnNumPedidoECommerceEdicaoLiberada)%>" />
<input type="hidden" name="blnObs1EdicaoLiberada" id="blnObs1EdicaoLiberada" value='<%=Cstr(blnObs1EdicaoLiberada)%>'>
<input type="hidden" name="blnObs2EdicaoLiberada" id="blnObs2EdicaoLiberada" value='<%=Cstr(blnObs2EdicaoLiberada)%>'>
<input type="hidden" name="blnObs3EdicaoLiberada" id="blnObs3EdicaoLiberada" value='<%=Cstr(blnObs3EdicaoLiberada)%>'>
<input type="hidden" name="blnFormaPagtoEdicaoLiberada" id="blnFormaPagtoEdicaoLiberada" value='<%=Cstr(blnFormaPagtoEdicaoLiberada)%>'>
<input type="hidden" name="blnEndEntregaEdicaoLiberada" id="blnEndEntregaEdicaoLiberada" value='<%=Cstr(blnEndEntregaEdicaoLiberada)%>'>
<input type="hidden" name="blnTransportadoraEdicaoLiberada" id="blnTransportadoraEdicaoLiberada" value='<%=Cstr(blnTransportadoraEdicaoLiberada)%>'>
<input type="hidden" name="blnValorFreteEdicaoLiberada" id="blnValorFreteEdicaoLiberada" value='<%=Cstr(blnValorFreteEdicaoLiberada)%>'>
<input type="hidden" name="blnGarantiaIndicadorEdicaoLiberada" id="blnGarantiaIndicadorEdicaoLiberada" value='<%=Cstr(blnGarantiaIndicadorEdicaoLiberada)%>'>
<input type="hidden" name="blnEntregaImediataEdicaoLiberada" id="blnEntregaImediataEdicaoLiberada" value='<%=Cstr(blnEntregaImediataEdicaoLiberada)%>'>
<input type="hidden" name="blnInstaladorInstalaEdicaoLiberada" id="blnInstaladorInstalaEdicaoLiberada" value='<%=Cstr(blnInstaladorInstalaEdicaoLiberada)%>'>
<input type="hidden" name="blnBemUsoConsumoEdicaoLiberada" id="blnBemUsoConsumoEdicaoLiberada" value='<%=Cstr(blnBemUsoConsumoEdicaoLiberada)%>'>
<input type="hidden" name="bln_RT_e_RA_EdicaoLiberada" id="bln_RT_e_RA_EdicaoLiberada" value='<%=Cstr(bln_RT_e_RA_EdicaoLiberada)%>'>
<input type="hidden" name="blnItemPedidoEdicaoLiberada" id="blnItemPedidoEdicaoLiberada" value='<%=Cstr(blnItemPedidoEdicaoLiberada)%>'>
<input type="hidden" name="blnPedidoRecebidoStatusEdicaoLiberada" id="blnPedidoRecebidoStatusEdicaoLiberada" value='<%=Cstr(blnPedidoRecebidoStatusEdicaoLiberada)%>'>
<input type="hidden" name="blnAEntregarStatusEdicaoLiberada" id="blnAEntregarStatusEdicaoLiberada" value='<%=Cstr(blnAEntregarStatusEdicaoLiberada)%>'>
<input type="hidden" name="blnDadosNFeMercadoriasDevolvidasEdicaoLiberada" id="blnDadosNFeMercadoriasDevolvidasEdicaoLiberada" value='<%=Cstr(blnDadosNFeMercadoriasDevolvidasEdicaoLiberada)%>'>
<input type="hidden" name="c_valor_frete" id="c_valor_frete" value="" />
<input type="hidden" name="url_origem" id="url_origem" value="<%=url_origem%>" />
<input type="hidden" name="c_analise_credito_inicial" id="c_analise_credito_inicial" value='<%=r_pedido.analise_credito%>' />
<% if Not blnIndicadorEdicaoLiberada then %>
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=r_pedido.indicador%>" />
<% end if %>

<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>
<!-- CAIXA DE CONFIRMAÇÃO DE EXCLUSÃO DE FRETE -->
       <div id="caixa-confirmacao" title="Excluir frete?">
  <span id="msgExcFrete" style="display:none">Deseja realmente excluir o(s) frete(s) marcados?</span>
</div>
<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
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
	strTextoIndicador = ""
	if r_pedido.indicador <> "" then
		strTextoIndicador = r_pedido.indicador
		if r_orcamentista_e_indicador.desempenho_nota <> "" then
			strTextoIndicador = strTextoIndicador & " (" & r_orcamentista_e_indicador.desempenho_nota & ")"
			end if
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




<!--  ENDEREÇO DO CLIENTE  -->
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
	
	if cliente__tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZÃO SOCIAL DO CLIENTE"
%>
    <tr>
	<td class="MC" align="left" colspan="2"><p class="Rf"><%=s_aux%></p>
	
		<p class="C"><%=s%>&nbsp;</p>
	
		</td>
	</tr>
	</table>
	
	<!--  ENDEREÇO DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	s = formata_endereco(cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep)
%>		
		<td align="left"><p class="Rf">ENDEREÇO</p><p class="C"><%=s%>&nbsp;</p></td>
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
	<td align="left"><p class="Rf">RG</p><input id="cliente__rg" name="cliente__rg" class="TA" maxlength="72" style="width:310px;" value="<%=s%>"></td>
	</tr>
	</table>


	<table width="649" class="QS" cellspacing="0">
		<tr>
			<td align="left"><p class="R">PRODUTOR RURAL</p><p class="C">
				<%s=cliente__produtor_rural_status%>
				<%if s = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then s_aux="checked" else s_aux=""%>
				
				<input type="radio" id="rb_produtor_rural_nao" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fPED.rb_produtor_rural[0].click();">Não</span>
				<%if s = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) then s_aux="checked" else s_aux=""%>
				
				<input type="radio" id="rb_produtor_rural_sim" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fPED.rb_produtor_rural[1].click();">Sim</span></p>
			</td>
		</tr>
	</table>


	

	<table width="649" class="QS" cellspacing="0" id="t_contribuinte_icms" onload="trataProdutorRural();">
		<tr>
			<%s=cliente__ie%>
			<td width="210" class="MD" align="left"><p class="R">IE</p><p class="C">
				<input id="cliente__ie" name="cliente__ie" class="TA" maxlength="72" style="width:310px;" value="<%=s%>" /></p>
			</td>
			<td align="left"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
				<%s=cliente__icms%>
				<%if s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then s_aux="checked" else s_aux=""%>
				<% intIdx = 0 %>
				<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Não</span>
				<%if s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then s_aux="checked" else s_aux=""%>
				<% intIdx = intIdx + 1 %>
				<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Sim</span>
				<%if s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then s_aux="checked" else s_aux=""%>
				<% intIdx = intIdx + 1 %>
				<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Isento</span></p>
			</td>
		</tr>
	</table>
	

<% else %>
	<td width="215"  align="left"><p class="Rf">IE</p><input id="cliente__ie" name="cliente__ie" class="TA" maxlength="72" style="width:310px;" value="<%=s%>"></td>
	</tr>
	<tr>
		<td class="MC" align="left" colspan="2"><p class="R">CONTRIBUINTE ICMS</p><p class="C">

				<%
                    s = " "
                    if cliente__icms = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
                        s = " checked "
                    end if
                %>
			
			<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s%>><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[1].click();">Não</span>
				<%
                    s = " "
                    if cliente__icms = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                        s = " checked "
                    end if
                %>
			<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s%>><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[2].click();">Sim</span>
				<%
                    s = " "
                    if cliente__icms = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
                        s = " checked "
                    end if
                %>
			<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s%>><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[3].click();">Isento</span></p>
			
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
	
	if cliente__tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZÃO SOCIAL DO CLIENTE"
%>
    <tr>
	<td class="MD" align="left" colspan="2"><p class="Rf"><%=s_aux%></p>
	
		
		<input id="cliente__nome" name="cliente__nome" class="TA" value="<%=s%>" maxlength="60" style="width:635px;" />
				
	
		</td>
	</tr>
	</table>
	
	<!--  ENDEREÇO DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
	    <tr>           
		    <td colspan="2" class="MB" align="left"><p class="Rf">ENDEREÇO</p><input id="endereco__endereco" name="endereco__endereco" class="TA" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_endereco_numero.focus(); filtra_nome_identificador();" value="<%=cliente__endereco%>"></td>
	    </tr>
	    <tr>
		    <td class="MDB" align="left"><p class="Rf">Nº</p><input id="endereco__numero" name="endereco__numero" class="TA" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_endereco_complemento.focus(); filtra_nome_identificador();" value="<%=cliente__endereco_numero%>"></td>
		    <td class="MB" align="left"><p class="Rf">COMPLEMENTO</p><input id="endereco__complemento" name="endereco__complemento" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_bairro.focus(); filtra_nome_identificador();" value="<%=cliente__endereco_complemento%>"></td>
	    </tr>
	    <tr>
		    <td class="MDB" align="left"><p class="Rf">BAIRRO</p><input id="endereco__bairro" name="endereco__bairro" class="TA" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_cidade.focus(); filtra_nome_identificador();" value="<%=cliente__bairro%>"></td>
		    <td class="MB" align="left"><p class="Rf">CIDADE</p><input id="endereco__cidade" name="endereco__cidade" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_uf.focus(); filtra_nome_identificador();" value="<%=cliente__cidade%>"></td>
	    </tr>
	    <tr>
		    <td width="50%" class="MD" align="left"><p class="Rf">UF</p><input id="endereco__uf" name="endereco__uf" class="TA" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fPED.EndEtg_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);" value="<%=cliente__uf%>"></td>
		    <td>
			    <table width="100%" cellspacing="0" cellpadding="0">
			    <tr>
			    <td width="50%" align="left"><p class="Rf">CEP</p><input id="endereco__cep" name="endereco__cep" readonly tabindex=-1 class="TA" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);" value='<%=cep_formata(cliente__cep)%>'></td>
			    <td align="center">
				    <% if blnPesquisaCEPAntiga then %>
				    <button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCepEndEtg();">Pesquisar CEP</button>
				    <% end if %>
				    <% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
				    <% if blnPesquisaCEPNova then %>
				    <button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="exibeJanelaCEP();">Pesquisar CEP</button>
				    <% end if %>
				    <a name="bLimparEndEtg" id="bLimparEndEtg" href="javascript:LimparCamposEndEtg(fPED)" title="limpa o endereço de entrega">
					    <img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
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
						<input id="cliente__ddd_res" name="cliente__ddd_res" class="TA" value="<%=cliente__ddd_res%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p>
					</td>
					<td class="MD" align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
						<input id="cliente__tel_res" name="cliente__tel_res" class="TA" value="<%=telefone_formata(cliente__tel_res)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
					</td>
	            </tr>
			</table>
			<table width="649" class="QS" cellspacing="0">
				<tr>
					<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_cel" name="cliente__ddd_cel" class="TA" value="<%=cliente__ddd_cel%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p>
					</td>
					<td align="left" class="MD"><p class="R">CELULAR</p><p class="C">
						<input id="cliente__tel_cel" name="cliente__tel_cel" class="TA" value="<%=telefone_formata(cliente__tel_cel)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_obs.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Número de celular inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
					</td>
	            </tr>
			</table>
			<table width="649" class="QS" cellspacing="0">
	            <tr>
					<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_com" name="cliente__ddd_com" class="TA" value="<%=cliente__ddd_com%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p>
					</td>
					<td class="MD" align="left"><p class="R">COMERCIAL</p><p class="C">
						<input id="cliente__tel_com" name="cliente__tel_com" class="TA" value="<%=telefone_formata(cliente__tel_com)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_obs.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Número de telefone comercial inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
					</td>
					<td align="left"><p class="R">RAMAL</p><p class="C">
						<input id="cliente__ramal_com" name="cliente__ramal_com" class="TA" value="<%=cliente__ramal_com%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fPED.EndEtg_ddd_com_2.focus(); filtra_numerico();"></p>
					</td>
				</tr>
				
			</table>	

		<% else %>
			<!--  TELEFONE DO CLIENTE  -->
			<table width="649" class="QS" cellspacing="0">
	            <tr>
					<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_com" name="cliente__ddd_com" class="TA" value="<%=cliente__ddd_com%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
					<td class="MD" align="left"><p class="R">TELEFONE </p><p class="C">
						<input id="cliente__tel_com" name="cliente__tel_com" class="TA" value="<%=telefone_formata(cliente__tel_com)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
					<td align="left"><p class="R">RAMAL</p><p class="C">
						<input id="cliente__ramal_com" name="cliente__ramal_com" class="TA" value="<%=cliente__ramal_com%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fPED.EndEtg_ddd_com_2.focus(); filtra_numerico();"></p>
					</td>
	            </tr>
	            <tr>
	                <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_com_2" name="cliente__ddd_com_2" class="TA" value="<%=cliente__ddd_com_2%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!!');this.focus();}" /></p>  
	                </td>
	                <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
						<input id="cliente__tel_com_2" name="cliente__tel_com_2" class="TA" value="<%=telefone_formata(cliente__tel_com_2)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	                </td>
	                <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
						<input id="cliente__ramal_com_2" name="cliente__ramal_com_2" class="TA" value="<%=cliente__ramal_com_2%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fPED.EndEtg_obs.focus(); filtra_numerico();" /></p>
	                </td>
	            </tr>
            </table>
		<% end if %>

	<!--  E-MAIL DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
		 <tr>           
		    <td colspan="2" class="Rf" align="left"><p class="Rf">E-MAIL</p>
				<input id="cliente__email" name="cliente__email" class="TA" maxlength="60" style="width:635px;" value="<%=cliente__email%>" onkeypress="Sfiltra_email();" />

		    </td>
	    </tr>
	</table>

	 <!-- ************   E-MAIL (XML)  ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td width="100%" align="left"><p class="R">E-MAIL (XML)</p><p class="C">
		    <input id="cliente__email_xml" name="cliente__email_xml" value="<%=cliente__email_xml%>" class="TA" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fPED.rb_end_entrega_nao.focus(); filtra_email();"></p></td>
	    </tr>
    </table>

<%end if%>





<% if Not blnEndEntregaEdicaoLiberada then %>
<!--  ENDEREÇO DE ENTREGA  -->
<%	
	s = pedido_formata_endereco_entrega(r_pedido, r_cliente)
%>		
<table width="649" class="QS" cellspacing="0" style="table-layout:fixed">
	<tr>
		<td align="left"><p class="Rf">ENDEREÇO DE ENTREGA</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
    <%	if r_pedido.EndEtg_cod_justificativa <> "" then %>		
	<tr>
		<td align="left" style="word-wrap:break-word"><p class="C"><%=obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA,r_pedido.EndEtg_cod_justificativa)%>&nbsp;</p></td>
	</tr>
    <%end if %>
</table>
<% else %>

    <% if blnUsarMemorizacaoCompletaEnderecos then %>
        <!--  ************  TIPO DO ENDEREÇO DE ENTREGA: PF/PJ (SOMENTE SE O CLIENTE FOR PJ)   ************ -->

        <%if eh_cpf then%>
            <!-- ************   ENDEREÇO DE ENTREGA PARA CLIENTE PF   ************ -->
            <!-- Pegamos todos os atuais. Sem campos editáveis. Pegamos os atuais do cadastro do cliente, não do pedido em si. -->
            <input type="hidden" id="EndEtg_tipo_pessoa" name="EndEtg_tipo_pessoa" value="PF"/>
            <input type="hidden" id="EndEtg_cnpj_cpf" name="EndEtg_cnpj_cpf" value="<%=r_cliente.cnpj_cpf%>"/>
            <input type="hidden" id="EndEtg_ie" name="EndEtg_ie" value="<%=r_cliente.ie%>"/>
            <input type="hidden" id="EndEtg_contribuinte_icms_status" name="EndEtg_contribuinte_icms_status" value="<%=r_cliente.contribuinte_icms_status%>"/>
            <input type="hidden" id="EndEtg_rg" name="EndEtg_rg" value="<%=r_cliente.rg%>"/>
            <input type="hidden" id="EndEtg_produtor_rural_status" name="EndEtg_produtor_rural_status" value="<%=r_cliente.produtor_rural_status%>"/>
            <input type="hidden" id="EndEtg_email" name="EndEtg_email" value="<%=r_cliente.email%>"/>
            <input type="hidden" id="EndEtg_email_xml" name="EndEtg_email_xml" value="<%=r_cliente.email_xml%>"/>
            <input type="hidden" id="EndEtg_nome" name="EndEtg_nome" value="<%=r_cliente.nome%>"/>


        <%else%>
            <table width="649" class="QS Habilitar_EndEtg_outroendereco" cellspacing="0">
	            <tr>
		            <td align="left">
		            <p class="R">ENDEREÇO DE ENTREGA</p><p class="C">
                        <%
                            s = " "
                            if r_pedido.EndEtg_tipo_pessoa = ID_PJ then
                                s = " checked "
                            end if
                        %>
			            <input type="radio" id="EndEtg_tipo_pessoa_PJ" name="EndEtg_tipo_pessoa" value="PJ" onclick="trocarEndEtgTipoPessoa(null);" <%=s%> >
			            <span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PJ');">Pessoa Jurídica</span>
			            &nbsp;
                        <%
                            s = " "
                            if r_pedido.EndEtg_tipo_pessoa = ID_PF then
                                s = " checked "
                            end if
                        %>
			            <input type="radio" id="EndEtg_tipo_pessoa_PF" name="EndEtg_tipo_pessoa" value="PF" onclick="trocarEndEtgTipoPessoa(null);" <%=s%> >
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
            <input type="hidden" id="EndEtg_rg" name="EndEtg_rg" value="<%=r_cliente.rg%>"/>
            <input type="hidden" id="EndEtg_produtor_rural_status" name="EndEtg_produtor_rural_status" />
            <input type="hidden" id="EndEtg_email" name="EndEtg_email" value="<%=r_cliente.email%>"/>
            <input type="hidden" id="EndEtg_email_xml" name="EndEtg_email_xml" value="<%=r_cliente.email_xml%>"/>


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
		            <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PJ_nao" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">Não</span>
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
		            <input type="radio"  <%=s%> id="EndEtg_produtor_rural_status_PF_nao" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>');">Não</span>
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
		            <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PF_nao" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">Não</span>
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



            <!-- ************   ENDEREÇO DE ENTREGA: NOME  ************ -->
            <table width="649" class="QS" cellspacing="0">
	            <tr>
	            <td width="100%" align="left"><p class="R" id="Label_EndEtg_nome">RAZÃO SOCIAL</p><p class="C">
		            <input id="EndEtg_nome" name="EndEtg_nome" class="TA" value="<%=r_pedido.EndEtg_nome%>" maxlength="60" size="85" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_endereco.focus(); filtra_nome_identificador();"></p></td>
	            </tr>
            </table>

        <%end if%>
    <%end if%> <% 'blnUsarMemorizacaoCompletaEnderecos %>



    <table width="649" class="QS" cellspacing="0">
	    <tr>
            <%
                s = "ENDEREÇO"
                if eh_cpf then
                    s = "ENDEREÇO DE ENTREGA"
                    end if
                %>
		    <td colspan="2" class="MB" align="left"><p class="Rf"><%=s%></p><input id="EndEtg_endereco" name="EndEtg_endereco" class="TA" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_endereco_numero.focus(); filtra_nome_identificador();" value="<%=r_pedido.EndEtg_endereco%>"></td>
	    </tr>
	    <tr>
		    <td class="MDB" align="left"><p class="Rf">Nº</p><input id="EndEtg_endereco_numero" name="EndEtg_endereco_numero" class="TA" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_endereco_complemento.focus(); filtra_nome_identificador();" value="<%=r_pedido.EndEtg_endereco_numero%>"></td>
		    <td class="MB" align="left"><p class="Rf">COMPLEMENTO</p><input id="EndEtg_endereco_complemento" name="EndEtg_endereco_complemento" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_bairro.focus(); filtra_nome_identificador();" value="<%=r_pedido.EndEtg_endereco_complemento%>"></td>
	    </tr>
	    <tr>
		    <td class="MDB" align="left"><p class="Rf">BAIRRO</p><input id="EndEtg_bairro" name="EndEtg_bairro" class="TA" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_cidade.focus(); filtra_nome_identificador();" value="<%=r_pedido.EndEtg_bairro%>"></td>
		    <td class="MB" align="left"><p class="Rf">CIDADE</p><input id="EndEtg_cidade" name="EndEtg_cidade" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_uf.focus(); filtra_nome_identificador();" value="<%=r_pedido.EndEtg_cidade%>"></td>
	    </tr>
	    <tr>
		    <td width="50%" class="MD" align="left"><p class="Rf">UF</p><input id="EndEtg_uf" name="EndEtg_uf" class="TA" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fPED.EndEtg_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);" value="<%=r_pedido.EndEtg_uf%>"></td>
		    <td>
			    <table width="100%" cellspacing="0" cellpadding="0">
			    <tr>
			    <td width="50%" align="left"><p class="Rf">CEP</p><input id="EndEtg_cep" name="EndEtg_cep" readonly tabindex=-1 class="TA" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);" value='<%=cep_formata(r_pedido.EndEtg_cep)%>'></td>
			    <td align="center">
				    <% if blnPesquisaCEPAntiga then %>
				    <button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCepEndEtg();">Pesquisar CEP</button>
				    <% end if %>
				    <% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
				    <% if blnPesquisaCEPNova then %>
				    <button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="exibeJanelaCEP_Etg();">Pesquisar CEP</button>
				    <% end if %>
				    <a name="bLimparEndEtg" id="bLimparEndEtg" href="javascript:LimparCamposEndEtg(fPED)" title="limpa o endereço de entrega">
					    <img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
			    </td>
			    </tr>
			    </table>
		    </td>
	    </tr>
    </table>


    <% if blnUsarMemorizacaoCompletaEnderecos then %>
        <%if eh_cpf then%>

            <!-- ************   ENDEREÇO DE ENTREGA PARA PF: TELEFONES   ************ -->
            <!-- Pegamos todos os atuais. Sem campos editáveis. Pegamos os atuais do cadastro do cliente, não do pedido em si. -->
            <input type="hidden" id="EndEtg_ddd_res" name="EndEtg_ddd_res" value="<%=r_cliente.ddd_res%>"/>
            <input type="hidden" id="EndEtg_tel_res" name="EndEtg_tel_res" value="<%=r_cliente.tel_res%>"/>
            <input type="hidden" id="EndEtg_ddd_cel" name="EndEtg_ddd_cel" value="<%=r_cliente.ddd_cel%>"/>
            <input type="hidden" id="EndEtg_tel_cel" name="EndEtg_tel_cel" value="<%=r_cliente.tel_cel%>"/>
            <input type="hidden" id="EndEtg_ddd_com" name="EndEtg_ddd_com" value="<%=r_cliente.ddd_com%>"/>
            <input type="hidden" id="EndEtg_tel_com" name="EndEtg_tel_com" value="<%=r_cliente.tel_com%>"/>
            <input type="hidden" id="EndEtg_ramal_com" name="EndEtg_ramal_com" value="<%=r_cliente.ramal_com%>"/>
            <input type="hidden" id="EndEtg_ddd_com_2" name="EndEtg_ddd_com_2" value="<%=r_cliente.ddd_com_2%>"/>
            <input type="hidden" id="EndEtg_tel_com_2" name="EndEtg_tel_com_2" value="<%=r_cliente.tel_com_2%>"/>
            <input type="hidden" id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" value="<%=r_cliente.ramal_com_2%>"/>

        <%else%>
        
            <!-- ************   ENDEREÇO DE ENTREGA: TELEFONE RESIDENCIAL   ************ -->
            <table width="649" class="QS Mostrar_EndEtg_pf Habilitar_EndEtg_outroendereco" cellspacing="0">
	            <tr>
	            <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		            <input id="EndEtg_ddd_res" name="EndEtg_ddd_res" class="TA" value="<%=r_pedido.EndEtg_ddd_res%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	            <td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
		            <input id="EndEtg_tel_res" name="EndEtg_tel_res" class="TA" value="<%=r_pedido.EndEtg_tel_res%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	            </tr>
	            <tr>
	            <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		            <input id="EndEtg_ddd_cel" name="EndEtg_ddd_cel" class="TA" value="<%=r_pedido.EndEtg_ddd_cel%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	            <td align="left" class="MC"><p class="R">CELULAR</p><p class="C">
		            <input id="EndEtg_tel_cel" name="EndEtg_tel_cel" class="TA" value="<%=r_pedido.EndEtg_tel_cel%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_obs.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Número de celular inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	            </tr>
            </table>
	
            <!-- ************   ENDEREÇO DE ENTREGA: TELEFONE COMERCIAL   ************ -->
            <table width="649" class="QS Mostrar_EndEtg_pj Habilitar_EndEtg_outroendereco" cellspacing="0">
	            <tr>
	            <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		            <input id="EndEtg_ddd_com" name="EndEtg_ddd_com" class="TA" value="<%=r_pedido.EndEtg_ddd_com%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	            <td class="MD" align="left"><p class="R">TELEFONE </p><p class="C">
		            <input id="EndEtg_tel_com" name="EndEtg_tel_com" class="TA" value="<%=r_pedido.EndEtg_tel_com%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	            <td align="left"><p class="R">RAMAL</p><p class="C">
		            <input id="EndEtg_ramal_com" name="EndEtg_ramal_com" class="TA" value="<%=r_pedido.EndEtg_ramal_com%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fPED.EndEtg_ddd_com_2.focus(); filtra_numerico();"></p></td>
	            </tr>
	            <tr>
	                <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
	                <input id="EndEtg_ddd_com_2" name="EndEtg_ddd_com_2" class="TA" value="<%=r_pedido.EndEtg_ddd_com_2%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!!');this.focus();}" /></p>  
	                </td>
	                <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
	                <input id="EndEtg_tel_com_2" name="EndEtg_tel_com_2" class="TA" value="<%=r_pedido.EndEtg_tel_com_2%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	                </td>
	                <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
	                <input id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" class="TA" value="<%=r_pedido.EndEtg_ramal_com_2%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fPED.EndEtg_obs.focus(); filtra_numerico();" /></p>
	                </td>
	            </tr>
            </table>

        <% end if %>
    <%end if%> <% 'blnUsarMemorizacaoCompletaEnderecos %>


    <!-- ************   JUSTIFIQUE O ENDEREÇO   ************ -->
    <table id="obs_endereco" width="649" class="QS" cellspacing="0">
	    <tr >
	    <td class="M" width="50%" align="left"><p class="R">JUSTIFIQUE O ENDEREÇO</p><p class="C">
           <select id="EndEtg_obs" name="EndEtg_obs" style="margin-right:225px;">			
               <option value="">&nbsp;</option>	
			 <%=justificativa_endereco_etg_monta_itens(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA, r_pedido.EndEtg_cod_justificativa)%>
		   </select></p></td>
           
	</tr>
</table>	
<% end if %>
<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<br>
<br>
<table class="Qx" cellspacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB" valign="bottom" align="left"><span class="PLTe">Fabr</span></td>
	<td class="MB" valign="bottom" align="left"><span class="PLTe">Produto</span></td>
	<td class="MB" valign="bottom" align="left"><span class="PLTe">Descrição</span></td>
	<td class="MB" valign="bottom" align="right"><span class="PLTd">Qtd</span></td>
	<td class="MB" valign="bottom" align="right"><span class="PLTd">Falt</span></td>
	<td class="MB" valign="bottom" align="right"><span class="PLTd">Preço</span></td>
	<td class="MB" valign="bottom" align="right"><span class="PLTd">VL Lista</span></td>
	<td class="MB" valign="bottom" align="right"><span class="PLTd">Desc</span></td>
	<td class="MB" valign="bottom" align="right"><span class="PLTd">VL Venda</span></td>
	<td class="MB" valign="bottom" align="right"><span class="PLTd">VL Total</span></td>
	</tr>

<% m_TotalDestePedido=0
   m_TotalDestePedidoComRA=0
   m_total_RA_deste_pedido=0
   m_total_venda_deste_pedido=0
   m_total_NF_deste_pedido=0
   n = Lbound(v_item)-1
   s_readonly_RT = "readonly tabindex=-1"
   if bln_RT_e_RA_EdicaoLiberada then s_readonly_RT = ""
	
   for i=1 to MAX_ITENS 
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
			if .desc_dado=0 then s_desc_dado="" else s_desc_dado=formata_perc_desc(.desc_dado)
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
			
			if blnItemPedidoEdicaoLiberada then s_readonly = ""
			
			' Para assegurar a consistência entre o valor total de NF e o total da forma de pagamento,
			' a edição fica permitida somente se o usuário puder editar a forma de pagamento!
			s_readonly_RA = s_readonly_RT
			if Not blnFormaPagtoEdicaoLiberada then s_readonly_RA = "readonly tabindex=-1"
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
	<td class="MDB" align="right"><input name="c_vl_NF" id="c_vl_NF" class="PLLd" style="width:62px; color:<%=s_cor%>"
		onkeypress="if (digitou_enter(true)) fPED.c_vl_unitario[<%=Cstr(i-1)%>].focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); recalcula_RA();recalcula_RA_Liquido();"
		value='<%=s_preco_NF%>' <%=s_readonly_RA%>></td>
	<td class="MDB" align="right"><input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_preco_lista%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_desc" id="c_desc" class="PLLd" style="width:28px; color:<%=s_cor%>"
		value='<%=s_desc_dado%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px; color:<%=s_cor%>"
		onkeypress="if (digitou_enter(true)) {if ((<%=Cstr(i)%>==fPED.c_vl_unitario.length)||(trim(fPED.c_produto[<%=Cstr(i)%>].value)=='')) fPED.c_obs1.focus(); else fPED.c_vl_NF[<%=Cstr(i)%>].focus();} filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); recalcula_total_linha(<%=Cstr(i)%>); recalcula_RA();recalcula_RA_Liquido();"
		value='<%=s_vl_unitario%>' <%=s_readonly%>></td>
	<td class="MDB" align="right"><input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px; color:<%=s_cor%>" 
		value='<%=s_vl_TotalItem%>' readonly tabindex=-1></td>
	</tr>
<% next %>

<%
'  O TOTAL DO RA (REPASSE AUTOMÁTICO) REFERENTE AOS ITENS DOS OUTROS PEDIDOS DESTA FAMÍLIA
   m_total_RA_outros = m_TotalFamiliaParcelaRA - m_total_RA_deste_pedido
   m_total_venda_outros = vl_TotalFamiliaPrecoVenda - m_total_venda_deste_pedido
   m_total_NF_outros = vl_TotalFamiliaPrecoNF - m_total_NF_deste_pedido
%>
	<tr>
	<td colspan="4" align="left">
		<table cellspacing="0" cellpadding="0" width='100%' style="margin-top:4px;">
		<tr>
			<td align="left" width="20%">&nbsp;</td>
			<td align="right">
			<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
				<tr>
				<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;RA Líquido</span></td>
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
			<td align="right">
				<table cellspacing="0" cellpadding="0">
					<tr>
					<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;COM(%)</span></td>
					<td class="MTBD" align="right"><input name="c_perc_RT" id="c_perc_RT" class="PLLd" style="width:30px;color:blue;" 
						value='<%=formata_perc_RT(r_pedido.perc_RT)%>' maxlength="5" 
						onkeypress="if (digitou_enter(true)) fPED.c_obs1.focus(); filtra_percentual();"
						onblur="this.value=formata_perc_RT(this.value); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();}"
						<%=s_readonly_RT%>
						></td>
					</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
	<td class="MD" align="left">&nbsp;</td>
	<td class="MDB" align="right">
		<input name="c_total_NF" id="c_total_NF" class="PLLd" style="width:70px;color:blue;" 
				value='<%=formata_moeda(m_TotalDestePedidoComRA)%>' readonly tabindex=-1>
	</td>
	<td colspan="3" class="MD" align="left">&nbsp;</td>
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
<!--  TRATA VERSÃO ANTIGA DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" style="width:649px;" cellspacing="0">
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observações </p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS1);" onblur="this.value=trim(this.value);"
				<% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
				><%=r_pedido.obs_1%></textarea>
		</td>
	</tr>
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Nº Nota Fiscal</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" maxlength="10" style="width:85px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_qtde_parcelas.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
			<% if Not blnObs2EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
				value='<%=r_pedido.obs_2%>'>
		</td>
	</tr>
	<tr>
		<td class="MDB" nowrap width="10%" align="left"><p class="Rf">Parcelas</p>
			<table cellspacing="0" cellpadding="0" width="100%"><tr>
				<td align="left"><input name="c_qtde_parcelas" id="c_qtde_parcelas" class="PLLc" maxlength="2" style="width:60px;" onkeypress="if (digitou_enter(true)) fPED.c_forma_pagto.focus(); filtra_numerico();"
						<% if Not blnFormaPagtoEdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
						value='<%if (r_pedido.qtde_parcelas<>0) Or (r_pedido.forma_pagto<>"") then Response.write Cstr(r_pedido.qtde_parcelas)%>'></td>
			</tr></table>
		</td>
		<td class="MDB" nowrap align="left" valign="top"><p class="Rf">Entrega Imediata</p>
			<% if Not blnEntregaImediataEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
				<%=strDisabled%>
				value="<%=COD_ETG_IMEDIATA_NAO%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[0].click();">Não</span>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
				<%=strDisabled%>
				value="<%=COD_ETG_IMEDIATA_SIM%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[1].click();">Sim</span>
		</td>
		<td class="MDB" nowrap align="left" valign="top"><p class="Rf">Bem de Uso/Consumo</p>
			<% if Not blnBemUsoConsumoEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				<%=strDisabled%>
				value="<%=COD_ST_BEM_USO_CONSUMO_NAO%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[0].click();">Não</span>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				<%=strDisabled%>
				value="<%=COD_ST_BEM_USO_CONSUMO_SIM%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[1].click();">Sim</span>
		</td>
		<td  class="MDB" align="left" nowrap><p class="Rf">Instalador Instala</p>
			<% if Not blnInstaladorInstalaEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
			<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
				<%=strDisabled%>
				value="<%=COD_INSTALADOR_INSTALA_NAO%>" <%if Cstr(r_pedido.InstaladorInstalaStatus)=Cstr(COD_INSTALADOR_INSTALA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_instalador_instala[0].click();">Não</span>
			<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
				<%=strDisabled%>
				value="<%=COD_INSTALADOR_INSTALA_SIM%>" <%if Cstr(r_pedido.InstaladorInstalaStatus)=Cstr(COD_INSTALADOR_INSTALA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_instalador_instala[1].click();">Sim</span>
		</td>
		<td class="MB" nowrap align="left" valign="top"><p class="Rf">Garantia Indicador</p>
			<% if Not blnGarantiaIndicadorEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
			<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
				<%=strDisabled%>
				value="<%=COD_GARANTIA_INDICADOR_STATUS__NAO%>" <%if Cstr(r_pedido.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[0].click();">Não</span>
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
		<td align="left" colspan="5"><p class="Rf">Forma de Pagamento</p>
			<textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
				style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_FORMA_PAGTO);" onblur="this.value=trim(this.value);"
				<% if Not blnFormaPagtoEdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
				><%=r_pedido.forma_pagto%></textarea>
		</td>
	</tr>
</table>

<% else %>

<!--  TRATA NOVA VERSÃO DA FORMA DE PAGAMENTO   -->
<input type="hidden" name="versao_forma_pagamento" id="versao_forma_pagamento" value='2'>
<br>

	<% if Not blnFormaPagtoEdicaoLiberada then %>
		<!--  EDIÇÃO ESTÁ BLOQUEADA EM PEDIDO ENTREGUE   -->
		<table class="Q" style="width:649px;" cellspacing="0">
			<tr>
				<td class="MB" colspan="7" align="left"><p class="Rf">Observações </p>
					<textarea name="c_obs1" id="c_obs1" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
						style="width:99%;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS1);" onblur="this.value=trim(this.value);"
						<% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
						><%=r_pedido.obs_1%></textarea>
				</td>
			</tr>
            <tr>
		        <td class="MB" colspan="7" align="left"><p class="Rf">Constar na NF</p>
			        <textarea name="c_nf_texto" id="c_nf_texto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_NF_TEXTO_CONSTAR)%>" 
				        style="width:99%;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_NF_TEXTO);" onblur="this.value=trim(this.value);"
				        <% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
                        ><%=r_pedido.NFe_texto_constar%></textarea>
		        </td>
	        </tr>
            <tr>
                <td class="MB" align="left" colspan="7" nowrap><p class="Rf">xPed</p>
			        <input name="c_num_pedido_compra" id="c_num_pedido_compra" class="PLLe" maxlength="15" style="width:100px;margin-left:2pt;" onkeypress="filtra_nome_identificador();" onblur="this.value=trim(this.value);"
				    <% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
                        value='<%=r_pedido.NFe_xPed%>'>
		        </td>
            </tr>
			<tr>
				<td class="MD" nowrap align="left"><p class="Rf">Nº Nota Fiscal</p>
					<input name="c_obs2" id="c_obs2" class="PLLe" maxlength="10" style="width:75px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_obs3.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
					<% if Not blnObs2EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
						value='<%=r_pedido.obs_2%>'>
				</td>
				<td class="MD" nowrap align="left"><p class="Rf">NF Simples Remessa</p>
					<input name="c_obs3" id="c_obs3" class="PLLe" maxlength="10" style="width:75px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_qtde_parcelas.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
					<% if Not blnObs3EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
						value='<%=r_pedido.obs_3%>'>
				</td>
                <td class="MD" align="left" valign="top" nowrap><p class="Rf">Número Magento</p>
			        <input name="c_pedido_ac" id="c_pedido_ac" class="PLLe" style="width:90px;margin-left:2pt;" maxlength="9" onkeypress="if (digitou_enter(true)) fPED.c_qtde_parcelas.focus(); filtra_nome_identificador();return SomenteNumero(event)" onblur="this.value=trim(this.value);"
				        value='<%=r_pedido.pedido_ac%>'
						<%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " readonly tabindex=-1" %>
						/>
		        </td>
				<td class="MD" nowrap align="left" valign="top"><p class="Rf">Entrega Imediata</p>
					<% if Not blnEntregaImediataEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
					<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
						<%=strDisabled%>
						value="<%=COD_ETG_IMEDIATA_NAO%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[0].click();">Não</span>
					<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
						<%=strDisabled%>
						value="<%=COD_ETG_IMEDIATA_SIM%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[1].click();">Sim</span>
				</td>
				<td class="MD" nowrap align="left" valign="top"><p class="Rf">Bem Uso/Consumo</p>
					<% if Not blnBemUsoConsumoEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
					<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
						<%=strDisabled%>
						value="<%=COD_ST_BEM_USO_CONSUMO_NAO%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[0].click();">Não</span>
					<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
						<%=strDisabled%>
						value="<%=COD_ST_BEM_USO_CONSUMO_SIM%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[1].click();">Sim</span>
				</td>
				<td class="MD" nowrap align="left"><p class="Rf">Instalador Instala</p>
					<% if Not blnInstaladorInstalaEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
					<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
						<%=strDisabled%>
						value="<%=COD_INSTALADOR_INSTALA_NAO%>" <%if Cstr(r_pedido.InstaladorInstalaStatus)=Cstr(COD_INSTALADOR_INSTALA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_instalador_instala[0].click();">Não</span>
					<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
						<%=strDisabled%>
						value="<%=COD_INSTALADOR_INSTALA_SIM%>" <%if Cstr(r_pedido.InstaladorInstalaStatus)=Cstr(COD_INSTALADOR_INSTALA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_instalador_instala[1].click();">Sim</span>
				</td>
				<td nowrap align="left" valign="top"><p class="Rf">Garantia Indicador</p>
					<% if Not blnGarantiaIndicadorEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
					<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
						<%=strDisabled%>
						value="<%=COD_GARANTIA_INDICADOR_STATUS__NAO%>" <%if Cstr(r_pedido.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[0].click();">Não</span>
					<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
						<%=strDisabled%>
						value="<%=COD_GARANTIA_INDICADOR_STATUS__SIM%>" <%if Cstr(r_pedido.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[1].click();">Sim</span>
				</td>
			</tr>
			<% if ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then %>
			<tr>
			    <td class="MC" colspan="2"><p class="Rf">Referente Pedido Bonshop: </p>
			    </td>
			    <td class="MC" colspan="5" align="left">
			        <select id="Select1" name="pedBonshop" style="width: 120px">
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
			<% end if %>

            <% if Cstr(r_pedido.loja) = Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then %>
            <tr>
                <td class="MC MD" align="left" nowrap valign="top"><p class="Rf">Nº Pedido Marketplace</p>
			        <input name="c_numero_mktplace" id="c_numero_mktplace" class="PLLe" maxlength="20" style="width:135px;margin-left:2pt;margin-top:5px;" onkeypress="filtra_nome_identificador();return SomenteNumero(event)" onblur="this.value=trim(this.value);"
				    <%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " readonly tabindex=-1" %> value="<%=r_pedido.pedido_bs_x_marketplace%>">
		        </td>
                <td class="MC" colspan="6" align="left" nowrap valign="top"><p class="Rf">Origem do Pedido</p>
			        <select name="c_origem_pedido" id="c_origem_pedido" style="margin: 3px; 3px; 3px"<%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " disabled tabindex=-1" %>>
                        <%=codigo_descricao_monta_itens_select_all(GRUPO_T_CODIGO_DESCRICAO__PEDIDOECOMMERCE_ORIGEM, r_pedido.marketplace_codigo_origem) %>
			        </select>
		        </td>
            </tr>
            <% end if %>

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
				<!--  À VISTA  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="0" border="0">
					  <tr>
						<td align="left"><span class="C">À Vista&nbsp&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.av_forma_pagto)%>)</span></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
				<!--  PARCELA ÚNICA  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="0" border="0">
					  <tr>
						<td align="left"><span class="C">Parcela Única:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pu_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pu_forma_pagto)%>)&nbsp;&nbsp;vencendo após&nbsp;<%=formata_inteiro(r_pedido.pu_vencto_apos)%>&nbsp;dias</span></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
				<!--  PARCELADO NO CARTÃO (INTERNET)  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="0" border="0">
					  <tr>
						<td align="left"><span class="C">Parcelado no Cartão (internet) em&nbsp;&nbsp;<%=Cstr(r_pedido.pc_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_pedido.pc_valor_parcela)%></span></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
				<!--  PARCELADO NO CARTÃO (MAQUINETA)  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="0" border="0">
					  <tr>
						<td align="left"><span class="C">Parcelado no Cartão (maquineta) em&nbsp;&nbsp;<%=Cstr(r_pedido.pc_maquineta_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_pedido.pc_maquineta_valor_parcela)%></span></td>
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
						<td align="left"><span class="C">Prestações:&nbsp;&nbsp;<%=formata_inteiro(r_pedido.pce_prestacao_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pce_prestacao_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pce_forma_pagto_prestacao)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=formata_inteiro(r_pedido.pce_prestacao_periodo)%>&nbsp;dias</span></td>
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
						<td align="left"><span class="C">1ª Prestação:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pse_prim_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pse_forma_pagto_prim_prest)%>)&nbsp;&nbsp;vencendo após&nbsp;<%=formata_inteiro(r_pedido.pse_prim_prest_apos)%>&nbsp;dias</span></td>
					  </tr>
					  <tr>
						<td align="left"><span class="C">Demais Prestações:&nbsp;&nbsp;<%=Cstr(r_pedido.pse_demais_prest_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pse_demais_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pse_forma_pagto_demais_prest)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=Cstr(r_pedido.pse_demais_prest_periodo)%>&nbsp;dias</span></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<% end if %>
			  </table>
			</td>
		  </tr>
		  <tr>
			<td class="MC" align="left"><p class="Rf">Informações Sobre Análise de Crédito</p>
			  <textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
						style="width:642px;margin-left:2pt;"
						readonly tabindex=-1><%=r_pedido.forma_pagto%></textarea>
			</td>
		  </tr>
		</table>
	
	<% else %>
		<!--  EDIÇÃO LIBERADA   -->
		<table class="Q" style="width:649px;" cellspacing="0">
			<tr>
				<td class="MB" colspan="7" align="left"><p class="Rf">Observações</p>
					<textarea name="c_obs1" id="c_obs1" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
						style="width:99%;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS1);" onblur="this.value=trim(this.value);"
						<% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
						><%=r_pedido.obs_1%></textarea>
				</td>
			</tr>
            <tr>
		        <td class="MB" colspan="7" align="left"><p class="Rf">Constar na NF</p>
			        <textarea name="c_nf_texto" id="c_nf_texto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_NF_TEXTO_CONSTAR)%>" 
				        style="width:99%;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_NF_TEXTO);" onblur="this.value=trim(this.value);"
				        <% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
                        ><%=r_pedido.NFe_texto_constar%></textarea>
		        </td>
	        </tr>
            <tr>
                <td class="MB" align="left" colspan="7" nowrap><p class="Rf">xPed</p>
			        <input name="c_num_pedido_compra" id="c_num_pedido_compra" class="PLLe" maxlength="15" style="width:100px;margin-left:2pt;" onkeypress="filtra_nome_identificador();" onblur="this.value=trim(this.value);"
				    <% if Not blnObs1EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
                        value='<%=r_pedido.NFe_xPed%>'>
		        </td>
            </tr>
			<tr>
				<td class="MD" nowrap align="left"><p class="Rf">Nº Nota Fiscal</p>
					<input name="c_obs2" id="c_obs2" class="PLLe" maxlength="10" style="width:75px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_obs3.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
						<% if Not blnObs2EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
						value='<%=r_pedido.obs_2%>'>
				</td>
				<td class="MD" nowrap align="left"><p class="Rf">NF Simples Remessa</p>
					<input name="c_obs3" id="c_obs3" class="PLLe" maxlength="10" style="width:75px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_pedido_ac.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
						<% if Not blnObs3EdicaoLiberada then Response.Write " readonly tabindex=-1 " %>
						value='<%=r_pedido.obs_3%>'>
				</td>
                <td class="MD" align="left" valign="top" nowrap><p class="Rf">Número Magento</p>
			        <input name="c_pedido_ac" id="c_pedido_ac" maxlength="9" class="PLLe" style="width:90px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fPED.c_qtde_parcelas.focus(); filtra_nome_identificador();return SomenteNumero(event)" onblur="this.value=trim(this.value);"
				        value='<%=r_pedido.pedido_ac%>'
						<%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " readonly tabindex=-1" %>
						/>
		        </td>
				<td class="MD" nowrap align="left" valign="top"><p class="Rf">Entrega Imediata</p>
					<% if Not blnEntregaImediataEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
					<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
						<%=strDisabled%>
						value="<%=COD_ETG_IMEDIATA_NAO%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[0].click();">Não</span>
					<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
						<%=strDisabled%>
						value="<%=COD_ETG_IMEDIATA_SIM%>" <%if Cstr(r_pedido.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_etg_imediata[1].click();">Sim</span>
				</td>
				<td class="MD" nowrap align="left" valign="top"><p class="Rf">Bem Uso/Consumo</p>
					<% if Not blnBemUsoConsumoEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
					<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
						<%=strDisabled%>
						value="<%=COD_ST_BEM_USO_CONSUMO_NAO%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[0].click();">Não</span>
					<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
						<%=strDisabled%>
						value="<%=COD_ST_BEM_USO_CONSUMO_SIM%>" <%if Cstr(r_pedido.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_bem_uso_consumo[1].click();">Sim</span>
				</td>
				<td class="MD" nowrap align="left"><p class="Rf">Instalador Instala</p>
					<% if Not blnInstaladorInstalaEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
					<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
						<%=strDisabled%>
						value="<%=COD_INSTALADOR_INSTALA_NAO%>" <%if Cstr(r_pedido.InstaladorInstalaStatus)=Cstr(COD_INSTALADOR_INSTALA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_instalador_instala[0].click();">Não</span>
					<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
						<%=strDisabled%>
						value="<%=COD_INSTALADOR_INSTALA_SIM%>" <%if Cstr(r_pedido.InstaladorInstalaStatus)=Cstr(COD_INSTALADOR_INSTALA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_instalador_instala[1].click();">Sim</span>
				</td>
				<td nowrap align="left" valign="top"><p class="Rf">Garantia Indicador</p>
					<% if Not blnGarantiaIndicadorEdicaoLiberada then strDisabled=" disabled" else strDisabled=""%>
					<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
						<%=strDisabled%>
						value="<%=COD_GARANTIA_INDICADOR_STATUS__NAO%>" <%if Cstr(r_pedido.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[0].click();">Não</span>
					<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
						<%=strDisabled%>
						value="<%=COD_GARANTIA_INDICADOR_STATUS__SIM%>" <%if Cstr(r_pedido.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_garantia_indicador[1].click();">Sim</span>
				</td>
			</tr>
			<% if ID_PARAM_SITE = COD_SITE_ASSISTENCIA_TECNICA then %>
			<tr>
			    <td class="MC" colspan="2"><p class="Rf">Referente Pedido Bonshop: </p>
			    </td>
			    <td class="MC" colspan="5" align="left">
			        <select id="Select2" name="pedBonshop" style="width: 120px">
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
			<% end if %>

            <% if Cstr(r_pedido.loja) = Cstr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then %>
            <tr>
                <td class="MC MD" align="left" nowrap valign="top"><p class="Rf">Nº Pedido Marketplace</p>
			        <input name="c_numero_mktplace" id="c_numero_mktplace" class="PLLe" maxlength="20" style="width:135px;margin-left:2pt;margin-top:5px;" onkeypress="filtra_nome_identificador();return SomenteNumero(event)" onblur="this.value=trim(this.value);"
					<%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " readonly tabindex=-1" %> value="<%=r_pedido.pedido_bs_x_marketplace%>">
		        </td>
                <td class="MC" colspan="6" align="left" nowrap valign="top"><p class="Rf">Origem do Pedido</p>
			        <select name="c_origem_pedido" id="c_origem_pedido" style="margin: 3px; 3px; 3px"<%if Not blnNumPedidoECommerceEdicaoLiberada then Response.Write " disabled tabindex=-1" %>>
                        <%=codigo_descricao_monta_itens_select_all(GRUPO_T_CODIGO_DESCRICAO__PEDIDOECOMMERCE_ORIGEM, r_pedido.marketplace_codigo_origem) %>
			        </select>
		        </td>
            </tr>
            <% end if %>

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
				<!--  À VISTA  -->
				<tr>
				  <td align="left">
					<table cellspacing="0" cellpadding="1" border="0">
					  <tr>
						<td align="left">
						  <% intIdx = 0 %>
						  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
								value="<%=COD_FORMA_PAGTO_A_VISTA%>"
								<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then Response.Write " checked"%>
								onclick="recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">À Vista</span>
						</td>
						<td align="left">&nbsp;</td>
						<td align="left">
						  <select id="op_av_forma_pagto" name="op_av_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
							<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
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
				<!--  PARCELA ÚNICA  -->
				<tr>
				  <td class="MC" align="left">
					<table cellspacing="0" cellpadding="1" border="0">
					  <tr>
						<td colspan="3" align="left">
						  <% intIdx = intIdx+1 %>
						  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
								value="<%=COD_FORMA_PAGTO_PARCELA_UNICA%>"
								<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then Response.Write " checked"%>
								onclick="pu_atualiza_valor();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcela Única</span>
						</td>
						<td align="left">&nbsp;</td>
						<td align="left">
						  <select id="op_pu_forma_pagto" name="op_pu_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
							<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
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
						  ><span class="C">vencendo após</span
						  ><input name="c_pu_vencto_apos" id="c_pu_vencto_apos" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();"
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
								value="<%=Cstr(r_pedido.pu_vencto_apos)%>"
							<% else %>
								value=""
							<% end if %>
						  ><span class="C">dias</span>
						</td>
					  </tr>
					</table>
				  </td>
				</tr>
				<!--  PARCELADO NO CARTÃO (INTERNET)  -->
				<% if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) Or _
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
								onclick="recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado no Cartão (internet)</span>
						</td>
						<td align="left">&nbsp;</td>
						<td align="left">
						  <input name="c_pc_qtde" id="c_pc_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pc_valor.focus(); filtra_numerico();" onblur="pc_calcula_valor_parcela();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();" 
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
								value="<%=Cstr(r_pedido.pc_qtde_parcelas)%>"
							<% else %>
								value=""
							<% end if %>
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
				<!--  PARCELADO NO CARTÃO (MAQUINETA)  -->
				<% if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) Or _
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
								onclick="recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado no Cartão (maquineta)</span>
						</td>
						<td align="left">&nbsp;</td>
						<td align="left">
						  <input name="c_pc_maquineta_qtde" id="c_pc_maquineta_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pc_maquineta_valor.focus(); filtra_numerico();" onblur="pc_maquineta_calcula_valor_parcela();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();" 
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
								value="<%=Cstr(r_pedido.pc_maquineta_qtde_parcelas)%>"
							<% else %>
								value=""
							<% end if %>
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
								onclick="pce_preenche_sugestao_intervalo();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado com Entrada</span>
						</td>
					  </tr>
					  <tr>
						<td style="width:60px;" align="left">&nbsp;</td>
						<td align="right"><span class="C">Entrada&nbsp;</span></td>
						<td align="left">
						  <select id="op_pce_entrada_forma_pagto" name="op_pce_entrada_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
							<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
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
						<td align="right"><span class="C">Prestações&nbsp;</span></td>
						<td align="left">
						  <select id="op_pce_prestacao_forma_pagto" name="op_pce_prestacao_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
							<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
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
						  <input name="c_pce_prestacao_qtde" id="c_pce_prestacao_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_pce_prestacao_valor.focus(); filtra_numerico();" onblur="pce_calcula_valor_parcela();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
								value="<%=Cstr(r_pedido.pce_prestacao_qtde)%>"
							<% else %>
								value=""
							<% end if %>
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
						<td colspan="2" align="left"><span class="C">Parcelas vencendo a cada</span
						><input name="c_pce_prestacao_periodo" id="c_pce_prestacao_periodo" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.c_forma_pagto.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();"
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
								value="<%=Cstr(r_pedido.pce_prestacao_periodo)%>"
							<% else %>
								value=""
							<% end if %>
						><span class="C">dias</span
						><span style="width:10px;">&nbsp;</span
						><span class="notPrint"><input name="b_pce_SugereFormaPagto" id="b_pce_SugereFormaPagto" type="button" class="Button" style="visibility:hidden;" onclick="pce_sugestao_forma_pagto();" value="sugestão automática" title="preenche o campo 'Forma de Pagamento' com uma sugestão de texto"></span
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
								onclick="pse_preenche_sugestao_intervalo();recalcula_RA_Liquido();recalculaCustoFinanceiroPrecoLista();"
								><span class="C" style="cursor:default" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado sem Entrada</span>
						</td>
					  </tr>
					  <tr>
						<td style="width:60px;" align="left">&nbsp;</td>
						<td align="right"><span class="C">1ª Prestação&nbsp;</span></td>
						<td align="left">
						  <select id="op_pse_prim_prest_forma_pagto" name="op_pse_prim_prest_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
							<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
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
						  ><span class="C">vencendo após</span
						  ><input name="c_pse_prim_prest_apos" id="c_pse_prim_prest_apos" class="Cc" maxlength="2" style="width:30px;" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fPED.op_pse_demais_prest_forma_pagto.focus(); filtra_numerico();" onblur="recalcula_RA_Liquido();"
							<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
								value="<%=Cstr(r_pedido.pse_prim_prest_apos)%>"
							<% else %>
								value=""
							<% end if %>
						  ><span class="C">dias</span>
						</td>
					  </tr>
					  <tr>
						<td style="width:60px;" align="left">&nbsp;</td>
						<td align="right"><span class="C">Demais Prestações&nbsp;</span></td>
						<td align="left">
						  <select id="op_pse_demais_prest_forma_pagto" name="op_pse_demais_prest_forma_pagto" onclick="fPED.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onchange="recalcula_RA_Liquido();">
							<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
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
						><span class="C">dias</span
						><span style="width:10px;">&nbsp;</span
						><span class="notPrint"><input name="b_pse_SugereFormaPagto" id="b_pse_SugereFormaPagto" type="button" class="Button" style="visibility:hidden;" onclick="pse_sugestao_forma_pagto();" value="sugestão automática" title="preenche o campo 'Forma de Pagamento' com uma sugestão de texto"></span
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
			  <p class="Rf">Informações Sobre Análise de Crédito</p>
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
	<td width="16.67%" class="MD" align="right" valign="bottom"><p class="Rf">VL Total&nbsp;&nbsp;(Família)&nbsp;</p></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><p class="Rf">VL Pago&nbsp;</p></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><p class="Rf">VL Devoluções&nbsp;</p></td>
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


<!--  ANÁLISE DE CRÉDITO   -->
<%	if operacao_permitida(OP_CEN_EDITA_ANALISE_CREDITO, s_lista_operacoes_permitidas) _
	   And ( (r_pedido.transportadora_id = "") OR (r_pedido.transportadora_selecao_auto_status <> 0) ) _
	   And (Trim("" & r_pedido.a_entregar_data_marcada) = "") _
	   And _
	  ( _
		( Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE) ) _
		  Or _
		( Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE_VENDAS) ) _
		  Or _
		( Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE_ENDERECO) ) _
		  Or _
		( Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO) ) _
		  Or _
		( Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO) ) _
		  Or _
		( Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_OK) ) _
		  Or _
		( (Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_ST_INICIAL)) And (r_pedido.analise_endereco_tratar_status<>0) And (r_pedido.analise_endereco_tratado_status<>0) ) _
	   ) then
		
		blnAnaliseCreditoEdicaoLiberada = True
%>
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td align="left">
		<p class="Rf">ANÁLISE DE CRÉDITO</p>
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td>
				<%intIdx=0%>
				<input type="radio" id="rb_analise_credito" name="rb_analise_credito" 
					value="<%=COD_AN_CREDITO_PENDENTE_VENDAS%>" <%if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE_VENDAS) then Response.Write " checked"%> onchange="exibeOcultaPendenteVendasMotivo()"><span class="C" style="cursor:default;color:red;" onclick="fPED.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS)%></span>
			</td>
			<td>
				<%intIdx=intIdx+1%>
				<input type="radio" id="rb_analise_credito" name="rb_analise_credito" 
					value="<%=COD_AN_CREDITO_PENDENTE_ENDERECO%>" <%if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE_ENDERECO) then Response.Write " checked"%> onchange="exibeOcultaPendenteVendasMotivo()"><span class="C" style="cursor:default;color:red;" onclick="fPED.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE_ENDERECO)%></span>
			</td>
			<td>
				<%intIdx=intIdx+1%>
				<input type="radio" id="rb_analise_credito" name="rb_analise_credito" 
					value="<%=COD_AN_CREDITO_PENDENTE%>" <%if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE) then Response.Write " checked"%> onchange="exibeOcultaPendenteVendasMotivo()"><span class="C" style="cursor:default;color:red;" onclick="fPED.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE)%></span>
			</td>
			<td>
				<%intIdx=intIdx+1%>
				<input type="radio" id="rb_analise_credito" name="rb_analise_credito" 
					value="<%=COD_AN_CREDITO_OK%>" <%if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_OK) then Response.Write " checked"%> onchange="exibeOcultaPendenteVendasMotivo()"><span class="C" style="cursor:default;color:green;" onclick="fPED.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_OK)%></span>
			</td>
		</tr>
        <tr id="trPendVendasMotivo">
            <td class="" align="left" colspan="9">
                <span class="C">Motivo: </span>
                <select name="c_pendente_vendas_motivo" id="c_pendente_vendas_motivo">
                    <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__AC_PENDENTE_VENDAS_MOTIVO, r_pedido.analise_credito_pendente_vendas_motivo) %>
                </select>
            </td>
        </tr>
		<tr>
			<td>
					<%intIdx=intIdx+1%>
					<input type="radio" id="rb_analise_credito" name="rb_analise_credito" 
						value="<%=COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO%>" <%if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO) then Response.Write " checked"%> onchange="exibeOcultaPendenteVendasMotivo()"><span class="C" style="cursor:default;color:darkorange;" onclick="fPED.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)%></span>
			</td>
			<td colspan="3">
					<%intIdx=intIdx+1%>
					<input type="radio" id="rb_analise_credito" name="rb_analise_credito" 
						value="<%=COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO%>" <%if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO) then Response.Write " checked"%> onchange="exibeOcultaPendenteVendasMotivo()"><span class="C" style="cursor:default;color:darkorange;" onclick="fPED.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)%></span>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<%	elseif operacao_permitida(OP_CEN_EDITA_ANALISE_CREDITO_PENDENTE_VENDAS, s_lista_operacoes_permitidas) _
	   And (r_pedido.transportadora_id = "") _
	   And (Trim("" & r_pedido.a_entregar_data_marcada) = "") _
	   And ( Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE_VENDAS) ) then
		
		blnAnaliseCreditoEdicaoLiberada = True
%>
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td align="left"><p class="Rf">ANÁLISE DE CRÉDITO</p>
			<%intIdx=0%>
			<input type="radio" id="rb_analise_credito" name="rb_analise_credito" 
				value="<%=COD_AN_CREDITO_PENDENTE_VENDAS%>" <%if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE_VENDAS) then Response.Write " checked"%> onchange="exibeOcultaPendenteVendasMotivo()"><span class="C" style="cursor:default" onclick="fPED.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS)%></span>
			<%intIdx=intIdx+1%>
			<input type="radio" id="rb_analise_credito" name="rb_analise_credito" 
				value="<%=COD_AN_CREDITO_PENDENTE%>" <%if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE) then Response.Write " checked"%> onchange="exibeOcultaPendenteVendasMotivo()"><span class="C" style="cursor:default" onclick="fPED.rb_analise_credito[<%=Cstr(intIdx)%>].click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE)%></span>
	</td>
</tr>

<tr id="trPendVendasMotivo">
            <td class="" align="left" colspan="2">
                <span class="C">Motivo: </span>
                <select name="c_pendente_vendas_motivo" id="c_pendente_vendas_motivo">
                    <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__AC_PENDENTE_VENDAS_MOTIVO, r_pedido.analise_credito_pendente_vendas_motivo) %>
                </select>
            </td>
        </tr>
</table>
<% else %>
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<%	s=x_analise_credito(r_pedido.analise_credito)
		if s <> "" then
            if Cstr(r_pedido.analise_credito)=Cstr(COD_AN_CREDITO_PENDENTE_VENDAS) then 
                if r_pedido.analise_credito_pendente_vendas_motivo <> "" then s = s & " (" & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__AC_PENDENTE_VENDAS_MOTIVO, r_pedido.analise_credito_pendente_vendas_motivo) & ")"  
                end if
			s_aux=formata_data_e_talvez_hora(r_pedido.analise_credito_data)
			if s_aux <> "" then s = s & " &nbsp; (" & s_aux & ")"
			end if
		if s="" then s="&nbsp;"
	%>
	<td align="left"><p class="Rf">ANÁLISE DE CRÉDITO</p><p class="C" style="color:<%=x_analise_credito_cor(r_pedido.analise_credito)%>;"><%=s%></p></td>
</tr>
</table>
<% end if %>

<input type="hidden" name="blnAnaliseCreditoEdicaoLiberada" id="blnAnaliseCreditoEdicaoLiberada" value='<%=Cstr(blnAnaliseCreditoEdicaoLiberada)%>'>


<% if blnHaDevolucoes And blnDadosNFeMercadoriasDevolvidasEdicaoLiberada then %>
<!--  DEVOLUÇÕES (EDITÁVEL)  -->
<!-- FORÇA A CRIAÇÃO DE UM ARRAY MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" name="c_item_devolvido_id" id="c_item_devolvido_id" value="">
<input type="hidden" name="c_item_devolvido_fabricante" id="c_item_devolvido_fabricante" value="">
<input type="hidden" name="c_item_devolvido_produto" id="c_item_devolvido_produto" value="">
<input type="hidden" name="c_item_devolvido_nfe_emitente" id="c_item_devolvido_nfe_emitente" value="">
<input type="hidden" name="c_item_devolvido_nfe_serie" id="c_item_devolvido_nfe_serie" value="">
<input type="hidden" name="c_item_devolvido_nfe_numero" id="c_item_devolvido_nfe_numero" value="">

<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td colspan="3" align="left"><p class="Rf" style="color:red;">DEVOLUÇÃO DE MERCADORIAS</p></td>
</tr>
<% for i=Lbound(v_item_devolvido) to Ubound(v_item_devolvido) 
		with v_item_devolvido(i)%>
<tr>
<input type="hidden" name="c_item_devolvido_id" id="c_item_devolvido_id" value="<%=.id%>">
<input type="hidden" name="c_item_devolvido_fabricante" id="c_item_devolvido_fabricante" value="<%=.fabricante%>">
<input type="hidden" name="c_item_devolvido_produto" id="c_item_devolvido_produto" value="<%=.produto%>">
	<%	if .qtde = 1 then s = "" else s = "s"
		s = formata_data(.devolucao_data) & " " & _
			formata_hhnnss_para_hh_nn(.devolucao_hora) & " - " & _
			formata_inteiro(.qtde) & " unidade" & s & " do " & .produto & " - " & produto_formata_descricao_em_html(.descricao_html)
		if Trim(.motivo) <> "" then	s = s & " (" & .motivo & ")"
	%>
	<td class="MC" colspan="3" align="left">
		<p class="C"><%=s%></p>
	</td>
</tr>
<tr>
	<td align="left">
		<p class="Rf">Emitente</p>
		<select id="c_item_devolvido_nfe_emitente" name="c_item_devolvido_nfe_emitente" class="C" style="margin-right:8px;width:300px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
		<% =nfe_emitente_monta_itens_select(.id_nfe_emitente) %>
		</select>
	</td>
	<td width="17%" align="left">
		<p class="Rf">Nº Série</p>
		<% if .NFe_serie_NF > 0 then s=Cstr(.NFe_serie_NF) else s="" %>
		<input name="c_item_devolvido_nfe_serie" id="c_item_devolvido_nfe_serie" class="Cc" maxlength="3" style="width:50px;" onkeypress="filtra_numerico();" value="<%=s%>">
	</td>
	<td width="17%" align="left">
		<p class="Rf">Nº NFe</p>
		<% if .NFe_numero_NF > 0 then s=Cstr(.NFe_numero_NF) else s="" %>
		<input name="c_item_devolvido_nfe_numero" id="c_item_devolvido_nfe_numero" class="Cc" maxlength="9" style="width:80px;" onkeypress="filtra_numerico();" value="<%=s%>">
	</td>
</tr>
<%		end with
	next %>
</table>
<% elseif blnHaDevolucoes then %>
<!--  DEVOLUÇÕES (SOMENTE CONSULTA)  -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td align="left"><p class="Rf" style="color:red;">DEVOLUÇÃO DE MERCADORIAS</p><p class="C"><%=s_devolucoes%></p></td>
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


<% if blnAEntregarStatusEdicaoLiberada then %>
<!--  DATA DE COLETA   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td align="left"><p class="Rf">DATA DE COLETA</p>
	<input name="c_a_entregar_data_marcada" id="c_a_entregar_data_marcada" maxlength="10" class="PLLe" style="width:100px;" value="<%=formata_data(r_pedido.a_entregar_data_marcada)%>"
		<%	if r_pedido.a_entregar_status<>0 then
				s = "if (tem_info(this.value)&&(this.value!='" & formata_data(r_pedido.a_entregar_data_marcada) & "')) if (retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(this.value)) < retorna_so_digitos(formata_ddmmyyyy_yyyymmdd('" & formata_data(Date) & "'))) {alert('Data inválida!'); this.focus();}"
			else
				s = "if (tem_info(this.value)) if (retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(this.value)) < retorna_so_digitos(formata_ddmmyyyy_yyyymmdd('" & formata_data(Date) & "'))) {alert('Data inválida: deve ser maior ou igual a hoje!'); this.focus();}"
				end if
		%>
		onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();} else {<%=s%>}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_data();"></td>
</tr>
</table>
<% end if %>


<% if r_pedido.transportadora_id <> "" then %>
  <% if Not blnTransportadoraEdicaoLiberada then %>
<!--  TRANSPORTADORA (SOMENTE CONSULTA)  -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td valign="top" align="left"><p class="Rf">TRANSPORTADORA</p></td>
</tr>
<tr>

    <%s=formata_data_e_talvez_hora(r_pedido.transportadora_data)
		if s <> "" then s = s & " - "
		s = s & r_pedido.transportadora_id & " (" & x_transportadora(r_pedido.transportadora_id) & ")"
		if s="" then s="&nbsp;" %>

	<td align="left">
		<span class="C"><%=s%></span>
	</td>

    <%  s = "SELECT * FROM t_PEDIDO_FRETE pf WHERE pedido='" & r_pedido.pedido & "' ORDER BY dt_hr_cadastro" 
        x = ""
        intQtdeFrete = 1
        vl_total_frete = 0
        set rs = cn.execute(s)

        do while Not rs.Eof
            if intQtdeFrete > 0 then x = x & "</tr><tr>" & chr(13)
            if Not blnValorFreteEdicaoLiberada then
                x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & UCase(rs("transportadora_id")) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:150px;'><span class='C'>" & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE, rs("codigo_tipo_frete")) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & cnpj_cpf_formata(Trim("" & rs("emissor_cnpj"))) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:80px;'><span class='C'>" & Trim("" & rs("numero_NF")) & "</td>" & chr(13)
                x = x & "<td class='MD MB' align='center' style='width:50px;'><span class='C'>" & NFeFormataSerieNF(Trim("" & rs("serie_NF"))) & "</td>" & chr(13)
                x = x & "<td class='MD MB' align='right' style='width:97px;padding-right: 5px'><span class='C'>" & formata_moeda(rs("vl_frete")) & "</td>" & chr(13)
                x = x & "<td class='MB' align='center' style='width:30px;'><span class='C'>&nbsp;</span></td>" & chr(13)              
            else
                x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & chr(13) & _
                        "       <select name='c_frete_transportadora_id' id='c_transportadora_id' style='width:130px'>" & chr(13) & _
                            transportadora_monta_select_somente_apelido(rs("transportadora_id")) & _
                        "   </select>" & chr(13)
                x = x & "<td class='MD MB' align='center' style='width:150px;'><span class='C'>" & chr(13) & _
                        "       <select name='c_tipo_frete' id='c_tipo_frete' style='width: 150px'>" & chr(13) & _
                        codigo_descricao_monta_itens_select_all(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE, rs("codigo_tipo_frete")) & _
                        "   </select>" & chr(13) & _
                        "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & chr(13) & _
                        "       <select name='c_frete_emitente' id='c_frete_emitente' style='width: 130px'>" & chr(13) & _
                        nfe_emitente_monta_itens_select(rs("id_nfe_emitente")) & _
                        "   </select>" & chr(13) & _
                        "</td>" & chr(13)   
                x = x & "<td class='MD MB' align='center' style='width:80px;'><span class='C'>" & chr(13) & _
                        "   <input name='c_frete_numero_NF' id='c_frete_numero_NF' class='PLLd' maxlength='10' style='width:80px;margin-bottom:2px;' value='" & rs("numero_NF") & "'></td>" & chr(13) & _
                        "</td>" & chr(13)   
                x = x & "<td class='MD MB' align='center' style='width:80px;'><span class='C'>" & chr(13) & _
                        "   <input name='c_frete_serie_NF' id='c_frete_serie_NF' class='PLLd' maxlength='10' style='width:50px;margin-bottom:2px;' value='" & rs("serie_NF") & "'></td>" & chr(13) & _
                        "</td>" & chr(13)   
                x = x & "<td class='MD MB' align='right' style='width:50px;padding-right: 5px'><span class='C'>" & chr(13) & _
                        "   <input name='c_valor_frete' id='c_valor_frete' class='PLLd' maxlength='18' style='width:50px;margin-bottom:2px;' onkeypress='filtra_moeda_positivo();' onblur='this.value=formata_moeda(this.value);' value='" & formata_moeda(rs("vl_frete")) & "'></td>" & chr(13) & _
                        "</td>" & chr(13)  
                x = x & "<td class='MB' align='center' style='width:30px;'><span class='C'><input type='checkbox' name='ckb_exclui_frete_" + CStr(intQtdeFrete) & "' id='ckb_exclui_frete_" + CStr(intQtdeFrete) & "' title='selecione para excluir o frete' value='" & rs("id") & "'>" & chr(13)   
                x = x & "<input type='hidden' name='frete_id' id='frete_id' value='" & rs("id") & "'></td>" & chr(13)
            end if
            
            intQtdeFrete = intQtdeFrete + 1
            vl_total_frete = vl_total_frete + rs("vl_frete")
        rs.MoveNext
        loop
        s = formata_moeda(vl_total_frete) 
    %>
	
</tr>
</table>
<br />
<table width="649" class="Q" cellspacing="0" style="border-bottom:0">
    <tr>
        <td class="MD MB" align="center" style="width:130px;"><p class="Rf">TRANSPORTADORA</p></td>
        <td class="MD MB" align="center" style="width:150px;"><p class="Rf">TIPO DE FRETE</p></td>
        <td class="MD MB" align="center" style="width:130px;"><p class="Rf">EMITENTE NF</p></td>
        <td class="MD MB" align="center" style="width:80px;"><p class="Rf">NÚMERO NF</p></td>
        <td class="MD MB" align="center" style="width:50px;"><p class="Rf">SÉRIE NF</p></td>
        <td class="MD MB" align="right" style="width:97px;padding-right: 5px"><p class="Rf">VALOR</p></td>
        <td class="MB" align="center" style="width:30px"><p class="Rf" style="color:darkred;font-weight:900;font-size:12pt;" title="excluir frete">&times;</p></td>
    </tr>
    <tr>
        <%=x%>
    </tr>
</table>
  <% else %>


<!--  TRANSPORTADORA (EDITÁVEL)  -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td valign="top" align="left"><p class="Rf">TRANSPORTADORA</p></td>
</tr>
<tr>

	<td align="left">
		<select id="c_transportadora_id" name="c_transportadora_id" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" style="margin-left:5px;margin-top:4pt; margin-bottom:4pt;">
			<%=transportadora_monta_select(r_pedido.transportadora_id)%>
		</select>
	</td>

    <%  s = "SELECT * FROM t_PEDIDO_FRETE WHERE pedido='" & r_pedido.pedido & "' ORDER BY dt_hr_cadastro" 
        x = ""
        intQtdeFrete = 1
        vl_total_frete = 0
        set rs = cn.execute(s)

        do while Not rs.Eof
            if intQtdeFrete > 0 then x = x & "</tr><tr>" & chr(13)
            if Not blnValorFreteEdicaoLiberada then
                x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & UCase(rs("transportadora_id")) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:150px;'><span class='C'>" & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE, rs("codigo_tipo_frete")) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & cnpj_cpf_formata(Trim("" & rs("emissor_cnpj"))) & "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:80px;'><span class='C'>" & Trim("" & rs("numero_NF")) & "</td>" & chr(13)
                x = x & "<td class='MD MB' align='center' style='width:50px;'><span class='C'>" & NFeFormataSerieNF(Trim("" & rs("serie_NF"))) & "</td>" & chr(13)
                x = x & "<td class='MD MB' align='right' style='width:97px;padding-right: 5px'><span class='C'>" & formata_moeda(rs("vl_frete")) & "</td>" & chr(13)
                x = x & "<td class='MB' align='center' style='width:30px;'><span class='C'>&nbsp;</span></td>" & chr(13)              
            else
                x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & chr(13) & _
                        "       <select name='c_frete_transportadora_id' id='c_transportadora_id' style='width:130px'>" & chr(13) & _
                            transportadora_monta_select_somente_apelido(rs("transportadora_id")) & _
                        "   </select>" & chr(13)
                x = x & "<td class='MD MB' align='center' style='width:150px;'><span class='C'>" & chr(13) & _
                        "       <select name='c_tipo_frete' id='c_tipo_frete' style='width: 150px'>" & chr(13) & _
                        codigo_descricao_monta_itens_select_all(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE, rs("codigo_tipo_frete")) & _
                        "   </select>" & chr(13) & _
                        "</td>" & chr(13)    
                x = x & "<td class='MD MB' align='center' style='width:130px;'><span class='C'>" & chr(13) & _
                        "       <select name='c_frete_emitente' id='c_frete_emitente' style='width: 130px'>" & chr(13) & _
                        wms_nfe_emitente_monta_itens_select(rs("id_nfe_emitente")) & _
                        "   </select>" & chr(13) & _
                        "</td>" & chr(13)   
                x = x & "<td class='MD MB' align='center' style='width:80px;'><span class='C'>" & chr(13) & _
                        "   <input name='c_frete_numero_NF' id='c_frete_numero_NF' class='PLLd' maxlength='10' style='width:80px;margin-bottom:2px;' value='" & rs("numero_NF") & "'></td>" & chr(13) & _
                        "</td>" & chr(13)   
                x = x & "<td class='MD MB' align='center' style='width:80px;'><span class='C'>" & chr(13) & _
                        "   <input name='c_frete_serie_NF' id='c_frete_serie_NF' class='PLLd' maxlength='10' style='width:50px;margin-bottom:2px;' value='" & rs("serie_NF") & "'></td>" & chr(13) & _
                        "</td>" & chr(13)   
                x = x & "<td class='MD MB' align='right' style='width:50px;padding-right: 5px'><span class='C'>" & chr(13) & _
                        "   <input name='c_valor_frete' id='c_valor_frete' class='PLLd' maxlength='18' style='width:50px;margin-bottom:2px;' onkeypress='filtra_moeda_positivo();' onblur='this.value=formata_moeda(this.value);' value='" & formata_moeda(rs("vl_frete")) & "'></td>" & chr(13) & _
                        "</td>" & chr(13)  
                x = x & "<td class='MB' align='center' style='width:30px;'><span class='C'><input type='checkbox' name='ckb_exclui_frete_" + CStr(intQtdeFrete) & "' id='ckb_exclui_frete_" + CStr(intQtdeFrete) & "' title='selecione para excluir o frete' value='" & rs("id") & "'>" & chr(13)   
                x = x & "<input type='hidden' name='frete_id' id='frete_id' value='" & rs("id") & "'></td>" & chr(13)
            end if
            
            intQtdeFrete = intQtdeFrete + 1
            vl_total_frete = vl_total_frete + rs("vl_frete")
        rs.MoveNext
        loop
        s = formata_moeda(vl_total_frete) 
    %>
	
</tr>
</table>
<br />
<table width="649" class="Q" cellspacing="0" style="border-bottom:0">
    <tr>
        <td class="MD MB" align="center" style="width:130px;"><p class="Rf">TRANSPORTADORA</p></td>
        <td class="MD MB" align="center" style="width:150px;"><p class="Rf">TIPO DE FRETE</p></td>
        <td class="MD MB" align="center" style="width:130px;"><p class="Rf">EMITENTE</p></td>
        <td class="MD MB" align="center" style="width:80px;"><p class="Rf">NÚMERO NF</p></td>
        <td class="MD MB" align="center" style="width:50px;"><p class="Rf">SÉRIE NF</p></td>
        <td class="MD MB" align="right" style="width:97px;padding-right: 5px"><p class="Rf">VALOR</p></td>
        <td class="MB" align="center" style="width:30px"><p class="Rf" style="color:darkred;font-weight:900;font-size:12pt;" title="excluir frete">&times;</p></td>
    </tr>
    <tr>
        <%=x%>
    </tr>
</table>
  <% end if %>
<% end if %>


<% if blnPedidoRecebidoStatusEdicaoLiberada then %>
<!--  EDITA FLAG QUE INDICA SE O PEDIDO FOI RECEBIDO PELO CLIENTE   -->
<br>
<table width="649" class="Q" cellspacing="0">
<tr>
	<td align="left"><p class="Rf">PEDIDO RECEBIDO PELO CLIENTE</p>
		<input type="radio" id="rb_PedidoRecebidoStatus" name="rb_PedidoRecebidoStatus" 
			style="margin-left:15px;"
			value="<%=COD_ST_PEDIDO_RECEBIDO_NAO%>" <%if Cstr(r_pedido.PedidoRecebidoStatus)=Cstr(COD_ST_PEDIDO_RECEBIDO_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_PedidoRecebidoStatus[0].click();">Não</span>
		<input type="radio" id="rb_PedidoRecebidoStatus" name="rb_PedidoRecebidoStatus" 
			style="margin-left:30px;margin-right:0px;"
			value="<%=COD_ST_PEDIDO_RECEBIDO_SIM%>" <%if Cstr(r_pedido.PedidoRecebidoStatus)=Cstr(COD_ST_PEDIDO_RECEBIDO_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fPED.rb_PedidoRecebidoStatus[1].click();">Sim, no dia</span>
		<input name="c_PedidoRecebidoData" id="c_PedidoRecebidoData" maxlength="10" class="Cc" style="width:100px;" value="<%=formata_data(r_pedido.PedidoRecebidoData)%>"
		<%	if r_pedido.PedidoRecebidoStatus<>0 then
				s = "if (tem_info(this.value)&&(this.value!='" & formata_data(r_pedido.PedidoRecebidoData) & "')) if (retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(this.value)) > retorna_so_digitos(formata_ddmmyyyy_yyyymmdd('" & formata_data(Date) & "'))) {alert('Data inválida: deve ser menor ou igual a hoje!'); this.focus();}"
			else
				s = "if (tem_info(this.value)) if (retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(this.value)) > retorna_so_digitos(formata_ddmmyyyy_yyyymmdd('" & formata_data(Date) & "'))) {alert('Data inválida: deve ser menor ou igual a hoje!'); this.focus();}"
				end if
		%>
		onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();} else {<%=s%>}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_data();">
		</td>
</tr>
</table>
<% end if %>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td align="left" class="Rc">&nbsp;</td></tr>
</table>
<br>

<input type="hidden" name="Verifica_End_Entrega" id="Verifica_End_Entrega" value=''>
<input type="hidden" name="Verifica_num" id="Verifica_num" value=''>
<input type="hidden" name="Verifica_Cidade" id="Verifica_Cidade" value=''>
<input type="hidden" name="Verifica_UF" id="Verifica_UF" value=''>
<input type="hidden" name="Verifica_CEP" id="Verifica_CEP" value=''>
<input type="hidden" name="Verifica_Justificativa" id="Verifica_Justificativa" value=''>

<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDConfirma(fPED)" title="confirma as alterações">
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