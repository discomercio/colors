<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=true %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L P E D I D O S M C R I T E X E C . A S P
'     =================================================================
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

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
	const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
	const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
	const MSO_NUMBER_FORMAT_DECIMAL = "\#\#\#\,\#\#\#\,\#\#0\.00"
	const MSO_NUMBER_FORMAT_TEXTO = "\@"
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISS�O DE ACESSO DO USU�RIO
	if (Not operacao_permitida(OP_CEN_REL_MULTICRITERIO_PEDIDOS_ANALITICO, s_lista_operacoes_permitidas)) And _
	   (Not operacao_permitida(OP_CEN_REL_MULTICRITERIO_PEDIDOS_SINTETICO, s_lista_operacoes_permitidas)) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	dim i
	dim s, s_aux, s_aux_dti, s_aux_dtf, s_filtro, flag_ok, cadastrado
	dim ckb_st_entrega_esperar, ckb_st_entrega_split, ckb_st_entrega_exceto_cancelados, ckb_st_entrega_exceto_entregues
	dim ckb_st_entrega_separar_sem_marc, ckb_st_entrega_separar_com_marc
	dim ckb_st_entrega_a_entregar_sem_marc, ckb_st_entrega_a_entregar_com_marc, ckb_pedido_nao_recebido_pelo_cliente, ckb_pedido_recebido_pelo_cliente
	dim c_dt_coleta_a_separar_inicio, c_dt_coleta_a_separar_termino, c_dt_coleta_st_a_entregar_inicio, c_dt_coleta_st_a_entregar_termino
	dim ckb_st_entrega_entregue, c_dt_entregue_inicio, c_dt_entregue_termino
	dim ckb_st_entrega_cancelado, c_dt_cancelado_inicio, c_dt_cancelado_termino
	dim ckb_st_pagto_pago, ckb_st_pagto_nao_pago, ckb_st_pagto_pago_parcial
	dim ckb_periodo_cadastro, c_dt_cadastro_inicio, c_dt_cadastro_termino
	dim ckb_entrega_marcada_para, c_dt_entrega_inicio, c_dt_entrega_termino
	dim ckb_periodo_emissao_NF_venda, c_dt_NF_venda_inicio, c_dt_NF_venda_termino
	dim ckb_periodo_emissao_NF_remessa, c_dt_NF_remessa_inicio, c_dt_NF_remessa_termino
	dim ckb_produto, c_fabricante, c_produto, c_grupo, v_grupos, ckb_somente_pedidos_produto_alocado
	dim rb_loja, c_loja, c_loja_de, c_loja_ate, vLoja, vLojaAux
	dim c_cliente_cnpj_cpf, c_cliente_uf
	dim c_transportadora, c_transportadora_multiplo
	dim ckb_visanet
	dim ckb_analise_credito_st_inicial, ckb_analise_credito_pendente_vendas, ckb_analise_credito_pendente_endereco, ckb_analise_credito_pendente, ckb_analise_credito_pendente_cartao
	dim ckb_analise_credito_ok, ckb_analise_credito_ok_aguardando_deposito, ckb_analise_credito_ok_deposito_aguardando_desbloqueio
	dim ckb_analise_credito_pendente_pagto_antecipado_boleto, ckb_analise_credito_ok_aguardando_pagto_boleto_av
	dim ckb_entrega_imediata_sim, ckb_entrega_imediata_nao, c_dt_previsao_entrega_inicio, c_dt_previsao_entrega_termino
	dim op_forma_pagto, c_forma_pagto_qtde_parc
	dim c_vendedor, c_indicador
	dim ckb_obs2_preenchido, ckb_obs2_nao_preenchido, ckb_indicador_preenchido, ckb_indicador_nao_preenchido, ckb_nao_exibir_links
	dim rb_saida
	dim data_pedido
    dim c_pedido_origem, c_grupo_pedido_origem,c_empresa
	dim c_FormFieldValues
    dim blnMostraMotivoCancelado, c_cancelados_ordena
	dim ckb_exibir_vendedor, ckb_exibir_parceiro, ckb_exibir_uf, ckb_exibir_data_previsao_entrega
	dim ckb_pagto_antecipado_status_nao, ckb_pagto_antecipado_status_sim, ckb_pagto_antecipado_quitado_status_pendente, ckb_pagto_antecipado_quitado_status_quitado
	dim ckb_exibir_cidade_etg, ckb_exibir_uf_etg, ckb_exibir_data_entrega, ckb_exibir_data_previsao_entrega_transp, ckb_exibir_data_recebido_cliente, ckb_exibir_qtde_volumes, ckb_exibir_peso, ckb_exibir_cubagem

	alerta = ""

	ckb_st_entrega_exceto_cancelados = Trim(Request.Form("ckb_st_entrega_exceto_cancelados"))
	ckb_st_entrega_exceto_entregues = Trim(Request.Form("ckb_st_entrega_exceto_entregues"))
	ckb_st_entrega_esperar = Trim(Request.Form("ckb_st_entrega_esperar"))
	ckb_st_entrega_split = Trim(Request.Form("ckb_st_entrega_split"))
	ckb_st_entrega_separar_sem_marc = Trim(Request.Form("ckb_st_entrega_separar_sem_marc"))
	ckb_st_entrega_separar_com_marc = Trim(Request.Form("ckb_st_entrega_separar_com_marc"))
	c_dt_coleta_a_separar_inicio = Trim(Request.Form("c_dt_coleta_a_separar_inicio"))
	c_dt_coleta_a_separar_termino = Trim(Request.Form("c_dt_coleta_a_separar_termino"))
	ckb_st_entrega_a_entregar_sem_marc = Trim(Request.Form("ckb_st_entrega_a_entregar_sem_marc"))
	ckb_st_entrega_a_entregar_com_marc = Trim(Request.Form("ckb_st_entrega_a_entregar_com_marc"))
	c_dt_coleta_st_a_entregar_inicio = Trim(Request.Form("c_dt_coleta_st_a_entregar_inicio"))
	c_dt_coleta_st_a_entregar_termino = Trim(Request.Form("c_dt_coleta_st_a_entregar_termino"))
	ckb_pedido_nao_recebido_pelo_cliente = Trim(Request.Form("ckb_pedido_nao_recebido_pelo_cliente"))
	ckb_pedido_recebido_pelo_cliente = Trim(Request.Form("ckb_pedido_recebido_pelo_cliente"))
	ckb_st_entrega_entregue = Trim(Request.Form("ckb_st_entrega_entregue"))
	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))
	ckb_st_entrega_cancelado = Trim(Request.Form("ckb_st_entrega_cancelado"))
	c_dt_cancelado_inicio = Trim(Request.Form("c_dt_cancelado_inicio"))
	c_dt_cancelado_termino = Trim(Request.Form("c_dt_cancelado_termino"))
	ckb_st_pagto_pago = Trim(Request.Form("ckb_st_pagto_pago"))
	ckb_st_pagto_nao_pago = Trim(Request.Form("ckb_st_pagto_nao_pago"))
	ckb_st_pagto_pago_parcial = Trim(Request.Form("ckb_st_pagto_pago_parcial"))
	ckb_periodo_cadastro = Trim(Request.Form("ckb_periodo_cadastro"))
	c_dt_cadastro_inicio = Trim(Request.Form("c_dt_cadastro_inicio"))
	c_dt_cadastro_termino = Trim(Request.Form("c_dt_cadastro_termino"))
	ckb_entrega_marcada_para = Trim(Request.Form("ckb_entrega_marcada_para"))
	c_dt_entrega_inicio = Trim(Request.Form("c_dt_entrega_inicio"))
	c_dt_entrega_termino = Trim(Request.Form("c_dt_entrega_termino"))
	ckb_periodo_emissao_NF_venda = Trim(Request.Form("ckb_periodo_emissao_NF_venda"))
	c_dt_NF_venda_inicio = Trim(Request.Form("c_dt_NF_venda_inicio"))
	c_dt_NF_venda_termino = Trim(Request.Form("c_dt_NF_venda_termino"))
	ckb_periodo_emissao_NF_remessa = Trim(Request.Form("ckb_periodo_emissao_NF_remessa"))
	c_dt_NF_remessa_inicio = Trim(Request.Form("c_dt_NF_remessa_inicio"))
	c_dt_NF_remessa_termino = Trim(Request.Form("c_dt_NF_remessa_termino"))
	ckb_produto = Trim(Request.Form("ckb_produto"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	ckb_somente_pedidos_produto_alocado = Trim(Request.Form("ckb_somente_pedidos_produto_alocado"))
	rb_loja = Ucase(Trim(Request.Form("rb_loja")))
	c_loja = Trim(Request.Form("c_loja"))
	c_loja_de = Trim(Request.Form("c_loja_de"))
	c_loja_ate = Trim(Request.Form("c_loja_ate"))
	c_cliente_cnpj_cpf=retorna_so_digitos(trim(request("c_cliente_cnpj_cpf")))
    c_cliente_uf=trim(request("c_cliente_uf"))
	c_transportadora = filtra_nome_identificador(UCase(Trim(Request.Form("c_transportadora"))))
	c_transportadora_multiplo = Trim(Request.Form("c_transportadora_multiplo"))
	ckb_visanet = Trim(Request.Form("ckb_visanet"))
	ckb_analise_credito_st_inicial = Trim(Request.Form("ckb_analise_credito_st_inicial"))
	ckb_analise_credito_pendente_vendas = Trim(Request.Form("ckb_analise_credito_pendente_vendas"))
	ckb_analise_credito_pendente_endereco = Trim(Request.Form("ckb_analise_credito_pendente_endereco"))
	ckb_analise_credito_pendente = Trim(Request.Form("ckb_analise_credito_pendente"))
	ckb_analise_credito_pendente_cartao = Trim(Request.Form("ckb_analise_credito_pendente_cartao"))
	ckb_analise_credito_pendente_pagto_antecipado_boleto = Trim(Request.Form("ckb_analise_credito_pendente_pagto_antecipado_boleto"))
	ckb_analise_credito_ok = Trim(Request.Form("ckb_analise_credito_ok"))
	ckb_analise_credito_ok_aguardando_deposito = Trim(Request.Form("ckb_analise_credito_ok_aguardando_deposito"))
	ckb_analise_credito_ok_deposito_aguardando_desbloqueio = Trim(Request.Form("ckb_analise_credito_ok_deposito_aguardando_desbloqueio"))
	ckb_analise_credito_ok_aguardando_pagto_boleto_av = Trim(Request.Form("ckb_analise_credito_ok_aguardando_pagto_boleto_av"))
	ckb_entrega_imediata_sim = Trim(Request.Form("ckb_entrega_imediata_sim"))
	ckb_entrega_imediata_nao = Trim(Request.Form("ckb_entrega_imediata_nao"))
	c_dt_previsao_entrega_inicio = Trim(Request.Form("c_dt_previsao_entrega_inicio"))
	c_dt_previsao_entrega_termino = Trim(Request.Form("c_dt_previsao_entrega_termino"))
	op_forma_pagto = Trim(Request.Form("op_forma_pagto"))
	c_forma_pagto_qtde_parc = retorna_so_digitos(Trim(Request.Form("c_forma_pagto_qtde_parc")))
	c_vendedor = Trim(Request.Form("c_vendedor"))
	c_indicador = Trim(Request.Form("c_indicador"))
	ckb_obs2_preenchido = Trim(Request.Form("ckb_obs2_preenchido"))
	ckb_obs2_nao_preenchido = Trim(Request.Form("ckb_obs2_nao_preenchido"))
	ckb_nao_exibir_links = Trim(Request.Form("ckb_nao_exibir_links"))
	ckb_indicador_preenchido = Trim(Request.Form("ckb_indicador_preenchido"))
	ckb_indicador_nao_preenchido = Trim(Request.Form("ckb_indicador_nao_preenchido"))
	rb_saida = Ucase(Trim(Request.Form("rb_saida")))
    c_pedido_origem = Trim(Request.Form("c_pedido_origem"))
    c_empresa = Trim(Request.Form("c_empresa"))
    c_grupo_pedido_origem = Trim(Request.Form("c_grupo_pedido_origem"))
	c_FormFieldValues = Trim(Request.Form("c_FormFieldValues"))
    c_grupo = Trim(Request.Form("c_grupo"))
    c_cancelados_ordena = Trim(Request.Form("c_cancelados_ordena"))
	ckb_exibir_vendedor = Trim(Request.Form("ckb_exibir_vendedor"))
	ckb_exibir_parceiro = Trim(Request.Form("ckb_exibir_parceiro"))
	ckb_exibir_uf = Trim(Request.Form("ckb_exibir_uf"))
	ckb_exibir_data_previsao_entrega = Trim(Request.Form("ckb_exibir_data_previsao_entrega"))
	ckb_pagto_antecipado_status_nao = Trim(Request.Form("ckb_pagto_antecipado_status_nao"))
	ckb_pagto_antecipado_status_sim = Trim(Request.Form("ckb_pagto_antecipado_status_sim"))
	ckb_pagto_antecipado_quitado_status_pendente = Trim(Request.Form("ckb_pagto_antecipado_quitado_status_pendente"))
	ckb_pagto_antecipado_quitado_status_quitado = Trim(Request.Form("ckb_pagto_antecipado_quitado_status_quitado"))
	ckb_exibir_cidade_etg = Trim(Request.Form("ckb_exibir_cidade_etg"))
	ckb_exibir_uf_etg = Trim(Request.Form("ckb_exibir_uf_etg"))
	ckb_exibir_data_entrega = Trim(Request.Form("ckb_exibir_data_entrega"))
	ckb_exibir_data_previsao_entrega_transp = Trim(Request.Form("ckb_exibir_data_previsao_entrega_transp"))
	ckb_exibir_data_recebido_cliente = Trim(Request.Form("ckb_exibir_data_recebido_cliente"))
	ckb_exibir_qtde_volumes = Trim(Request.Form("ckb_exibir_qtde_volumes"))
	ckb_exibir_peso = Trim(Request.Form("ckb_exibir_peso"))
	ckb_exibir_cubagem = Trim(Request.Form("ckb_exibir_cubagem"))

	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|FormFields", c_FormFieldValues)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_nao_exibir_links", ckb_nao_exibir_links)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_vendedor", ckb_exibir_vendedor)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_parceiro", ckb_exibir_parceiro)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_uf", ckb_exibir_uf)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_data_previsao_entrega", ckb_exibir_data_previsao_entrega)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_cidade_etg", ckb_exibir_cidade_etg)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_uf_etg", ckb_exibir_uf_etg)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_data_entrega", ckb_exibir_data_entrega)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_data_previsao_entrega_transp", ckb_exibir_data_previsao_entrega_transp)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_data_recebido_cliente", ckb_exibir_data_recebido_cliente)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_qtde_volumes", ckb_exibir_qtde_volumes)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_peso", ckb_exibir_peso)
	call set_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_cubagem", ckb_exibir_cubagem)

	if alerta = "" then
		if c_fabricante <> "" then
			s = "SELECT fabricante FROM t_FABRICANTE WHERE (fabricante='" & c_fabricante & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "FABRICANTE " & c_fabricante & " N�O EST� CADASTRADO."
				end if
			end if
		end if
		
	if alerta = "" then
		if c_produto <> "" then
			if (Not IsEAN(c_produto)) And (c_fabricante="") then
				alerta=texto_add_br(alerta)
				alerta=alerta & "N�O FOI ESPECIFICADO O FABRICANTE DO PRODUTO A SER CONSULTADO."
			else
				s = "SELECT * FROM t_PRODUTO WHERE"
				if IsEAN(c_produto) then
					s = s & " (ean='" & c_produto & "')"
				else
					s = s & " (fabricante='" & c_fabricante & "') AND (produto='" & c_produto & "')"
					end if
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if Not rs.Eof then
					flag_ok = True
					if IsEAN(c_produto) And (c_fabricante<>"") then
						if (c_fabricante<>Trim("" & rs("fabricante"))) then
							flag_ok = False
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto a ser consultado " & c_produto & " N�O pertence ao fabricante " & c_fabricante & "."
							end if
						end if
					if flag_ok then
					'	CARREGA C�DIGO INTERNO DO PRODUTO
						c_fabricante = Trim("" & rs("fabricante"))
						c_produto = Trim("" & rs("produto"))
						end if
					end if
				end if
			end if
		end if
		
	redim vLoja(0)
	vLoja(0) = ""
	if alerta = "" then
		if rb_loja = "UMA" then
			if c_loja = "" then
				alerta = "Especifique o n�mero da loja."
			else
				c_loja = substitui_caracteres(c_loja, ",", " ")
				c_loja = substitui_caracteres(c_loja, ";", " ")
				vLojaAux = Split(c_loja, " ")
				for i = LBound(vLojaAux) to UBound(vLojaAux)
					if Trim("" & vLojaAux(i)) <> "" then
						if Trim("" & vLoja(UBound(vLoja))) <> "" then
							redim preserve vLoja(UBound(vLoja)+1)
							vLoja(UBound(vLoja)) = ""
							end if
						vLoja(UBound(vLoja)) = retorna_so_digitos(Trim("" & vLojaAux(i)))
						end if
					next

				for i=LBound(vLoja) to UBound(vLoja)
					if Trim("" & vLoja(i)) <> "" then
						s = "SELECT loja FROM t_LOJA WHERE (loja='" & Trim("" & vLoja(i)) & "')"
						if rs.State <> 0 then rs.Close
						rs.open s, cn
						if rs.Eof then
							alerta = "Loja " & Trim("" & vLoja(i)) & " n�o est� cadastrada."
							end if
						end if
					next
				end if
		elseif rb_loja = "FAIXA" then
			if (c_loja_de="") And (c_loja_ate="") then
				alerta = "Especifique o intervalo de lojas para consulta."
			else
				if c_loja_de <> "" then
					s = "SELECT loja FROM t_LOJA WHERE (loja='" & c_loja_de & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta = alerta & "Loja " & c_loja_de & " n�o est� cadastrada."
						end if
					end if
				
				if c_loja_ate <> "" then
					s = "SELECT loja FROM t_LOJA WHERE (loja='" & c_loja_ate & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta = alerta & "Loja " & c_loja_ate & " n�o est� cadastrada."
						end if
					end if
				end if
			end if
		end if
		
	if alerta = "" then
		if c_cliente_cnpj_cpf <> "" then
			if Not cnpj_cpf_ok(c_cliente_cnpj_cpf) then
				alerta=texto_add_br(alerta)
				alerta = alerta & "CNPJ/CPF do cliente � inv�lido."
				end if
			end if
		end if
	
	if alerta = "" then
		if c_transportadora <> "" then
			if Trim(x_transportadora(c_transportadora)) = "" then
				alerta=texto_add_br(alerta)
				alerta = alerta & "Transportadora '" & c_transportadora & "' N�O est� cadastrada."
				end if
			end if
		end if

	dim s_sessionToken
	s_sessionToken = ""
	s = "SELECT Convert(varchar(36), SessionTokenModuloCentral) AS SessionTokenModuloCentral FROM t_USUARIO WHERE (usuario = '" & usuario & "')"
	if rs.State <> 0 then rs.Close
	rs.open s, cn
	if Not rs.Eof then s_sessionToken = Trim("" & rs("SessionTokenModuloCentral"))
	if rs.State <> 0 then rs.Close


'	Per�odo de consulta est� restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)

		if alerta = "" then
			if ckb_st_entrega_separar_com_marc <> "" then
				if (c_dt_coleta_a_separar_inicio <> "") And (c_dt_coleta_a_separar_termino <> "") then
					if StrToDate(c_dt_coleta_a_separar_termino) < StrToDate(c_dt_coleta_a_separar_inicio) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "A data de t�rmino (" & c_dt_coleta_a_separar_termino & ") do per�odo de coleta � anterior � data de in�cio (" & c_dt_coleta_a_separar_inicio & ")!"
						end if
					end if
				end if
			end if

		if alerta = "" then
			if ckb_st_entrega_a_entregar_com_marc <> "" then
				if (c_dt_coleta_st_a_entregar_inicio <> "") And (c_dt_coleta_st_a_entregar_termino <> "") then
					if StrToDate(c_dt_coleta_st_a_entregar_termino) < StrToDate(c_dt_coleta_st_a_entregar_inicio) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "A data de t�rmino (" & c_dt_coleta_st_a_entregar_termino & ") do per�odo de coleta � anterior � data de in�cio (" & c_dt_coleta_st_a_entregar_inicio & ")!"
						end if
					end if
				end if
			end if

	'	COLOCADOS ENTRE
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_cadastro_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_cadastro_termino
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			if c_dt_cadastro_inicio = "" then c_dt_cadastro_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
			end if
		
	'	ENTREGUE ENTRE
		if ckb_st_entrega_entregue <> "" then
			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_entregue_inicio
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_entregue_termino
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				if c_dt_entregue_inicio = "" then c_dt_entregue_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
				end if
			end if

	'	CANCELADO ENTRE
		if ckb_st_entrega_cancelado <> "" then
			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_cancelado_inicio
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_cancelado_termino
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				if c_dt_cancelado_inicio = "" then c_dt_cancelado_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
				end if
			end if
		
	'	DATA DE COLETA (R�TULO ANTIGO: ENTREGA MARCADA ENTRE)
		if ckb_entrega_marcada_para <> "" then
			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_entrega_inicio
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_entrega_termino
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				if c_dt_entrega_inicio = "" then c_dt_entrega_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
				end if
			end if
		
	'	PER�ODO DE EMISS�O DA NF DE VENDA
		if ckb_periodo_emissao_NF_venda <> "" then
			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_NF_venda_inicio
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_NF_venda_termino
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				if c_dt_NF_venda_inicio = "" then c_dt_NF_venda_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
				end if
			end if

	'	PER�ODO DE EMISS�O DA NF DE REMESSA
		if ckb_periodo_emissao_NF_remessa <> "" then
			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_NF_remessa_inicio
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				strDtRefDDMMYYYY = c_dt_NF_remessa_termino
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if

			if alerta = "" then
				if c_dt_NF_remessa_inicio = "" then c_dt_NF_remessa_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
				end if
			end if

	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if

    '   MOSTRA COLUNAS DE MOTIVO CANCELAMENTO E VALOR ORIGINAL DO PEDIDO CANCELADO?
    blnMostraMotivoCancelado = False
    if ckb_st_entrega_cancelado <> "" then blnMostraMotivoCancelado = True
	
	dim blnSaidaExcel
	blnSaidaExcel = False
	if alerta = "" then
		if rb_saida = "XLS" then
			blnSaidaExcel = True
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=RelPedMultiCrit_" & formata_data_yyyymmdd(Now) & "_" & formata_hora_hhnnss(Now) & ".xls"
			Response.Write "<h2>Relat�rio Multicrit�rio de Pedidos</h2>"
			Response.Write excel_monta_texto_filtro
			Response.Write "<br><br>"
			consulta_executa
			Response.End
			end if
		end if



' _____________________________________________________________________________________________
'
'									F  U  N  �  �  E  S 
' _____________________________________________________________________________________________

function monta_link_view_pedido(byval id_pedido, byval usuario)
dim strLink
	monta_link_view_pedido = ""
	id_pedido = Trim("" & id_pedido)
	if id_pedido = "" then exit function
	strLink = "<a href='javascript:fPEDConsultaView(" & _
				chr(34) & id_pedido & chr(34) & _
				"," & _
				chr(34) & usuario & chr(34) & _
				")' title='clique para consultar o pedido " & id_pedido & "'>" & _
				"&nbsp;<img id='imgPedidoConsultaView' src='../imagem/doc_preview_12.png' class='notPrint' />" & _
				"</a>"
	monta_link_view_pedido=strLink
end function

' _____________________________________
' EXCEL MONTA TEXTO FILTRO
'
function excel_monta_texto_filtro
dim s, s_aux, s_resp

	s_resp = ""
	s = ""
	s_aux = Lcase(x_status_entrega(ckb_st_entrega_esperar))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_status_entrega(ckb_st_entrega_split))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_separar_sem_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s_aux = s_aux & " (sem data de coleta)"
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_separar_com_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		if c_dt_coleta_a_separar_inicio <> "" then s_aux_dti = c_dt_coleta_a_separar_inicio else s_aux_dti = "N.I."
		if c_dt_coleta_a_separar_termino <> "" then s_aux_dtf = c_dt_coleta_a_separar_termino else s_aux_dtf = "N.I."
		s_aux = s_aux & " (com data de coleta: " & s_aux_dti & " a " & s_aux_dtf & ")"
		s = s & s_aux
		end if
		
	s_aux = Lcase(x_status_entrega(ckb_st_entrega_a_entregar_sem_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s_aux = s_aux & " (sem data de coleta)"
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_a_entregar_com_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		if c_dt_coleta_st_a_entregar_inicio <> "" then s_aux_dti = c_dt_coleta_st_a_entregar_inicio else s_aux_dti = "N.I."
		if c_dt_coleta_st_a_entregar_termino <> "" then s_aux_dtf = c_dt_coleta_st_a_entregar_termino else s_aux_dtf = "N.I."
		s_aux = s_aux & " (com data de coleta: " & s_aux_dti & " a " & s_aux_dtf & ")"
		s = s & s_aux
		end if

    s_aux = Lcase(x_status_entrega(ckb_st_entrega_esperar))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_entregue))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		s_aux = c_dt_entregue_inicio
		if s_aux = "" then s_aux = "N.I."
		s_aux = " (" & s_aux & " a "
		s = s & s_aux
		s_aux = c_dt_entregue_termino
		if s_aux = "" then s_aux = "N.I."
		s_aux = s_aux & ")"
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_cancelado))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		s_aux = c_dt_cancelado_inicio
		if s_aux = "" then s_aux = "N.I."
		s_aux = " (" & s_aux & " a "
		s = s & s_aux
		s_aux = c_dt_cancelado_termino
		if s_aux = "" then s_aux = "N.I."
		s_aux = s_aux & ")"
		s = s & s_aux
		end if

	if ckb_st_entrega_exceto_cancelados <> "" then
		s_aux = "exceto cancelados"
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if    

	if ckb_st_entrega_exceto_entregues <> "" then
		s_aux = "exceto entregues"
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	if s <> "" then
		s_resp = s_resp & "Status de Entrega: " & s
		s_resp = s_resp & "<br>"
		end if

	s = ""
    if ckb_pedido_nao_recebido_pelo_cliente <> "" then
		s_aux = "n�o recebidos"
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	if ckb_pedido_recebido_pelo_cliente <> "" then
		s_aux = "recebidos"
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	if s <> "" then
		s_resp = s_resp & "Pedidos Recebidos pelo Cliente: " & s
		s_resp = s_resp & "<br>"
		end if

	s = ""
	s_aux = Lcase(x_status_pagto(ckb_st_pagto_pago))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_status_pagto(ckb_st_pagto_nao_pago))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_pagto(ckb_st_pagto_pago_parcial))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	if s <> "" then
		s_resp = s_resp & "Status de Pagamento: " & s
		s_resp = s_resp & "<br>"
		end if

	s = ""
	if ckb_pagto_antecipado_status_nao <> "" then
		if s <> "" then s = s & ", "
		s = s & "N�o"
		end if

	if ckb_pagto_antecipado_status_sim <> "" then
		if s <> "" then s = s & ", "
		s = s & "Sim"
		end if

	if s <> "" then
		s_resp = s_resp & "Pagamento Antecipado: " & s
		s_resp = s_resp & "<br>"
		end if

	s = ""
	if ckb_pagto_antecipado_quitado_status_pendente <> "" then
		if s <> "" then s = s & ", "
		s = s & pagto_antecipado_quitado_descricao(COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO, COD_PAGTO_ANTECIPADO_QUITADO_STATUS_PENDENTE)
		end if

	if ckb_pagto_antecipado_quitado_status_quitado <> "" then
		if s <> "" then s = s & ", "
		s = s & pagto_antecipado_quitado_descricao(COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO, COD_PAGTO_ANTECIPADO_QUITADO_STATUS_QUITADO)
		end if

	if s <> "" then
		s_resp = s_resp & "Status Pagamento Antecipado: " & s
		s_resp = s_resp & "<br>"
		end if

	s = ""
	if ckb_analise_credito_st_inicial <> "" then
		s_aux = "status inicial"
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_vendas))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_endereco))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_cartao))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_pagto_antecipado_boleto))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok_aguardando_deposito))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok_deposito_aguardando_desbloqueio))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok_aguardando_pagto_boleto_av))
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	if s <> "" then
		s_resp = s_resp & "An�lise de Cr�dito: " & s
		s_resp = s_resp & "<br>"
		end if

	s = ""
	s_aux = ""
	if CStr(ckb_entrega_imediata_sim) = CStr(COD_ETG_IMEDIATA_SIM) then s_aux = "sim"
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if
	
	s_aux = ""
	if CStr(ckb_entrega_imediata_nao) = CStr(COD_ETG_IMEDIATA_NAO) then s_aux = "n�o"
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		s = s & " (previs�o de entrega: "
		s_aux = c_dt_previsao_entrega_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s = s & " a "
		s_aux = c_dt_previsao_entrega_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s = s & ")"
		end if

	if s <> "" then
		s_resp = s_resp & "Entrega Imediata: " & s
		s_resp = s_resp & "<br>"
		end if
	
	'Geral: campo Obs II
	s = ""
	s_aux = ""
	if ckb_obs2_preenchido <> "" then s_aux = "N� NF preenchido"
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if
	
	s_aux = ""
	if ckb_obs2_nao_preenchido <> "" then s_aux = "N� NF n�o preenchido"
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	if s <> "" then
		s_resp = s_resp & "Geral: " & s
		s_resp = s_resp & "<br>"
		end if

	'Indicador preenchido
	s = ""
	s_aux = ""
	if ckb_indicador_preenchido <> "" then s_aux = "Indicador preenchido"
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if
	
	s_aux = ""
	if ckb_indicador_nao_preenchido <> "" then s_aux = "Indicador n�o preenchido"
	if s_aux<>"" then
		if s <> "" then s = s & ", "
		s = s & s_aux
		end if

	if s <> "" then
		s_resp = s_resp & "Indicador: " & s
		s_resp = s_resp & "<br>"
		end if

	if (c_dt_cadastro_inicio <> "") Or (c_dt_cadastro_termino <> "") then
		s = ""
		s_aux = c_dt_cadastro_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " e "
		s_aux = c_dt_cadastro_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_resp = s_resp & "Pedidos colocados entre: " & s
		s_resp = s_resp & "<br>"
		end if

	if ckb_entrega_marcada_para <> "" then
		s = ""
		s_aux = c_dt_entrega_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " a "
		s_aux = c_dt_entrega_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_resp = s_resp & "Data de coleta: " & s
		s_resp = s_resp & "<br>"
		end if
	
	if ckb_periodo_emissao_NF_venda <> "" then
		s = ""
		s_aux = c_dt_NF_venda_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " a "
		s_aux = c_dt_NF_venda_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_resp = s_resp & "Emiss�o NF Venda: " & s
		s_resp = s_resp & "<br>"
		end if

	if ckb_periodo_emissao_NF_remessa <> "" then
		s = ""
		s_aux = c_dt_NF_remessa_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " a "
		s_aux = c_dt_NF_remessa_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_resp = s_resp & "Emiss�o NF Remessa: " & s
		s_resp = s_resp & "<br>"
		end if

	if ckb_produto <> "" then 
		s_aux = c_fabricante
		if s_aux = "" then s_aux = "todos"
		s = "fabricante: " & s_aux
		s_aux = c_produto
		if s_aux = "" then s_aux = "todos"
		s = s & ", produto: " & s_aux
		s_resp = s_resp & "Somente pedidos que incluam: " & s
		if ckb_somente_pedidos_produto_alocado <> "" then s_resp = s_resp & " (somente pedidos que possuam o produto alocado)"
		s_resp = s_resp & "<br>"
		end if

	select case rb_loja
		case "TODAS": s = "todas"
		case "UMA"
			s = ""
			for i=LBound(vLoja) to UBound(vLoja)
				if s <> "" then s = s & ", "
				s = s & Trim("" & vLoja(i))
				next
		case "FAIXA"
			s = ""
			s_aux = c_loja_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " a "
			s_aux = c_loja_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
		case else: s = ""
		end select
	
	s_resp = s_resp & "Lojas: " & s
	s_resp = s_resp & "<br>"

	if op_forma_pagto <> "" then
		s = x_opcao_forma_pagamento(op_forma_pagto)
		if s = "" then s = " "
		s_resp = s_resp & "Forma Pagto: " & s
		s_resp = s_resp & "<br>"
		end if

	if c_forma_pagto_qtde_parc <> "" then
		s = c_forma_pagto_qtde_parc
		if s = "" then s = " "
		s_resp = s_resp & "N� Parcelas: " & s
		s_resp = s_resp & "<br>"
		end if

	if c_cliente_cnpj_cpf <> "" then
		s = cnpj_cpf_formata(c_cliente_cnpj_cpf)
		s_aux = x_cliente_por_cnpj_cpf(c_cliente_cnpj_cpf, cadastrado)
		if Not cadastrado then s_aux = "N�o Cadastrado"
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		if s = "" then s = " "
		s_resp = s_resp & "Cliente: " & s
		s_resp = s_resp & "<br>"
		end if

	if ckb_visanet <> "" then
		s_resp = s_resp & "Cart�o de Cr�dito: " & "somente pedidos pagos usando cart�o de cr�dito"
		s_resp = s_resp & "<br>"
		end if
	
	if c_transportadora <> "" then
		s = c_transportadora
		s_aux = iniciais_em_maiusculas(x_transportadora(c_transportadora))
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_resp = s_resp & "Transportadora: " & s
		s_resp = s_resp & "<br>"
		end if

	if c_transportadora_multiplo <> "" then
		s_resp = s_resp & "Transportadora(s): " & c_transportadora_multiplo
		s_resp = s_resp & "<br>"
		end if

	if c_vendedor <> "" then
		s = c_vendedor
		s_aux = x_usuario(c_vendedor)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_resp = s_resp & "Vendedor: " & s
		s_resp = s_resp & "<br>"
		end if

	if c_indicador <> "" then
		s = c_indicador
		s_aux = x_orcamentista_e_indicador(c_indicador)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_resp = s_resp & "Indicador: " & s
		s_resp = s_resp & "<br>"
		end if
	
	s_resp = s_resp & "Emiss�o: " & formata_data_hora(Now)
	s_resp = s_resp & "<br><br>"
	
	excel_monta_texto_filtro = s_resp
end function



' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim blnPorFornecedor
dim s, s_aux, s_periodo_aux, s_cor, s_bkg_color, s_nbsp, s_align, s_nowrap, s_sql, cab_table, cab, n_reg, n_reg_total, n_colspan, n_colspan_final, s_colspan_final, s_loja
dim s_where, s_where_aux, s_where_ext, s_from, cont
dim vl_total_faturamento, vl_sub_total_faturamento, vl_total_pago, vl_sub_total_pago
dim vl_total_faturamento_NF, vl_sub_total_faturamento_NF
dim vl_a_pagar, vl_sub_total_a_pagar, vl_total_a_pagar
dim vl_total_fornecedor, vl_sub_total_fornecedor
dim vl_total_fornecedor_NF, vl_sub_total_fornecedor_NF
dim vl_total_pedido_original, vl_sub_total_pedido_original, vl_pedido_original
dim total_qtde_vol, sub_total_qtde_vol, total_cubagem, sub_total_cubagem, total_peso, sub_total_peso
dim x, loja_a, qtde_lojas
dim w_pedido, w_pedido_magento, w_data, w_NF, w_cliente, w_st_entrega, w_valor, w_motivo_cancelamento
dim w_vendedor, w_indicador, w_uf_cad_cliente, w_cidade_etg, w_uf_etg, w_qtde_vol, w_cubagem, w_peso
dim blnRelAnalitico
dim intNumLinha
dim s_grupo_origem
dim s_link_rastreio, s_link_rastreio2, s_numero_NF
dim rPSSW
dim sLinkView
dim vTransportadora
	
	set rPSSW = get_registro_t_parametro(ID_PARAMETRO_SSW_Rastreamento_Lista_Transportadoras)

'	RELAT�RIO SINT�TICO OU ANAL�TICO?
	blnRelAnalitico=False
	if operacao_permitida(OP_CEN_REL_MULTICRITERIO_PEDIDOS_ANALITICO, s_lista_operacoes_permitidas) then blnRelAnalitico=True

	s_colspan_final = ""
	n_colspan_final = 1
	if ckb_exibir_data_previsao_entrega <> "" then n_colspan_final = n_colspan_final + 1
	if ckb_exibir_cidade_etg <> "" then n_colspan_final = n_colspan_final + 1
	if ckb_exibir_uf_etg <> "" then n_colspan_final = n_colspan_final + 1
	if ckb_exibir_data_entrega <> "" then n_colspan_final = n_colspan_final + 1
	if ckb_exibir_data_previsao_entrega_transp <> "" then n_colspan_final = n_colspan_final + 1
	if ckb_exibir_data_recebido_cliente <> "" then n_colspan_final = n_colspan_final + 1
	if ckb_exibir_qtde_volumes <> "" then n_colspan_final = n_colspan_final + 1
	if ckb_exibir_peso <> "" then n_colspan_final = n_colspan_final + 1
	if ckb_exibir_cubagem <> "" then n_colspan_final = n_colspan_final + 1
	
	if Not blnMostraMotivoCancelado then
		if ckb_exibir_vendedor <> "" then n_colspan_final = n_colspan_final + 1
		if ckb_exibir_parceiro <> "" then n_colspan_final = n_colspan_final + 1
		if ckb_exibir_uf <> "" then n_colspan_final = n_colspan_final + 1
		end if
	if n_colspan_final > 0 then s_colspan_final = " colspan=" & Cstr(n_colspan_final)

'	MONTA CL�USULA WHERE
	s_where = ""
	s_where_ext = ""

'	CRIT�RIO: STATUS DE ENTREGA
	s = ""
	s_aux = ckb_st_entrega_esperar
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.st_entrega = '" & s_aux & "')"
		end if

	s_aux = ckb_st_entrega_split
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.st_entrega = '" & s_aux & "')"
		end if

	s_aux = ckb_st_entrega_separar_sem_marc
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO.st_entrega = '" & s_aux & "')AND(t_PEDIDO.a_entregar_status=0))"
		end if

	s_aux = ckb_st_entrega_separar_com_marc
	if s_aux <> "" then
		s_where_aux = ""
		if c_dt_coleta_a_separar_inicio <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (t_PEDIDO.a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_coleta_a_separar_inicio)) & ")"
			end if
		if c_dt_coleta_a_separar_termino <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (t_PEDIDO.a_entregar_data_marcada < " & bd_formata_data(StrToDate(c_dt_coleta_a_separar_termino)+1) & ")"
			end if
		if s_where_aux <> "" then s_where_aux = " AND" & s_where_aux
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO.st_entrega = '" & s_aux & "')AND(t_PEDIDO.a_entregar_status<>0)" & s_where_aux & ")"
		end if

	s_aux = ckb_st_entrega_a_entregar_sem_marc
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO.st_entrega = '" & s_aux & "')AND(t_PEDIDO.a_entregar_status=0))"
		end if

	s_aux = ckb_st_entrega_a_entregar_com_marc
	if s_aux <> "" then
		s_where_aux = ""
		if c_dt_coleta_st_a_entregar_inicio <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (t_PEDIDO.a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_coleta_st_a_entregar_inicio)) & ")"
			end if
		if c_dt_coleta_st_a_entregar_termino <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (t_PEDIDO.a_entregar_data_marcada < " & bd_formata_data(StrToDate(c_dt_coleta_st_a_entregar_termino)+1) & ")"
			end if
		if s_where_aux <> "" then s_where_aux = " AND" & s_where_aux
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO.st_entrega = '" & s_aux & "')AND(t_PEDIDO.a_entregar_status<>0)" & s_where_aux & ")"
		end if

'	ENTREGUE ENTRE
	if ckb_st_entrega_entregue <> "" then
		s_aux = ""
		if c_dt_entregue_inicio <> "" then 
			if s_aux <> "" then s_aux = s_aux & " AND"
			s_aux = s_aux & " (t_PEDIDO.entregue_data >= " & bd_formata_data(StrToDate(c_dt_entregue_inicio)) & ")"
			end if
		if c_dt_entregue_termino <> "" then 
			if s_aux <> "" then s_aux = s_aux & " AND"
			s_aux = s_aux & " (t_PEDIDO.entregue_data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
			end if
		
		if s_aux <> "" then s_aux = s_aux & " AND"
		s_aux = s_aux & " (t_PEDIDO.st_entrega = '" & ckb_st_entrega_entregue & "')"
		
		if s_aux <> "" then s_aux = " (" & s_aux & ")"
		if s <> "" then s = s & " OR"
		s = s & s_aux
		end if

'	CANCELADO ENTRE
	if ckb_st_entrega_cancelado <> "" then
		s_aux = ""
		if c_dt_cancelado_inicio <> "" then 
			if s_aux <> "" then s_aux = s_aux & " AND"
			s_aux = s_aux & " (t_PEDIDO.cancelado_data >= " & bd_formata_data(StrToDate(c_dt_cancelado_inicio)) & ")"
			end if
		if c_dt_cancelado_termino <> "" then 
			if s_aux <> "" then s_aux = s_aux & " AND"
			s_aux = s_aux & " (t_PEDIDO.cancelado_data < " & bd_formata_data(StrToDate(c_dt_cancelado_termino)+1) & ")"
			end if
		
		if s_aux <> "" then s_aux = s_aux & " AND"
		s_aux = s_aux & " (t_PEDIDO.st_entrega = '" & ckb_st_entrega_cancelado & "')"
		
		if s_aux <> "" then s_aux = " (" & s_aux & ")"
		if s <> "" then s = s & " OR"
		s = s & s_aux
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	EXCETO CANCELADOS
	if ckb_st_entrega_exceto_cancelados <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')"
		end if

'	EXCETO ENTREGUES
	if ckb_st_entrega_exceto_entregues <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.st_entrega <> '" & ST_ENTREGA_ENTREGUE & "')"
		end if

'	CRIT�RIO: PEDIDOS RECEBIDOS PELO CLIENTE
	s = ""
    s_aux = ckb_pedido_nao_recebido_pelo_cliente
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.PedidoRecebidoStatus = 0)"
		end if

	s_aux = ckb_pedido_recebido_pelo_cliente
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.PedidoRecebidoStatus = 1)"
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRIT�RIO: STATUS DE PAGAMENTO
	s = ""
	s_aux = ckb_st_pagto_pago
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.st_pagto = '" & s_aux & "')"
		end if

	s_aux = ckb_st_pagto_nao_pago
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.st_pagto = '" & s_aux & "')"
		end if
	
	s_aux = ckb_st_pagto_pago_parcial
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.st_pagto = '" & s_aux & "')"
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRIT�RIO: PAGAMENTO ANTECIPADO
	s = ""
	s_aux = ckb_pagto_antecipado_status_nao
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.PagtoAntecipadoStatus = '" & s_aux & "')"
		end if

	s_aux = ckb_pagto_antecipado_status_sim
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.PagtoAntecipadoStatus = '" & s_aux & "')"
		end if

	if s <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRIT�RIO: STATUS PAGAMENTO ANTECIPADO
	s = ""
	s_aux = ckb_pagto_antecipado_quitado_status_pendente
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.PagtoAntecipadoQuitadoStatus = '" & s_aux & "')"
		end if

	s_aux = ckb_pagto_antecipado_quitado_status_quitado
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.PagtoAntecipadoQuitadoStatus = '" & s_aux & "')"
		end if

	if s <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " ((t_PEDIDO__BASE.PagtoAntecipadoStatus = " & COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO & ") AND (" & s & "))"
		end if

'	CRIT�RIO: AN�LISE DE CR�DITO
	s = ""

	s_aux = ckb_analise_credito_st_inicial
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if
	
	s_aux = ckb_analise_credito_pendente_vendas
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_pendente_endereco
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_pendente
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if
	
	s_aux = ckb_analise_credito_pendente_pagto_antecipado_boleto
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_ok
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_ok_aguardando_deposito
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_ok_deposito_aguardando_desbloqueio
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

	s_aux = ckb_analise_credito_ok_aguardando_pagto_boleto_av
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO__BASE.analise_credito = " & s_aux & ")"
		end if

'	O STATUS "PENDENTE CART�O DE CR�DITO" N�O EXISTE NO BD, � UMA SITUA��O DEFINIDA
'	PELA COMBINA��O DO STATUS COD_AN_CREDITO_ST_INICIAL + FORMA DE PAGTO USANDO SOMENTE PAGAMENTO POR CART�O
	s_aux = ckb_analise_credito_pendente_cartao
	if s_aux <> "" then
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO__BASE.analise_credito = " & COD_AN_CREDITO_ST_INICIAL & ") AND (t_PEDIDO__BASE.st_forma_pagto_somente_cartao = 1))"
		end if
	
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRIT�RIO: ENTREGA IMEDIATA
	s = ""
	if ckb_entrega_imediata_sim <> "" then
		if s <> "" then s = s & " OR"
		s = s & " (t_PEDIDO.st_etg_imediata = " & COD_ETG_IMEDIATA_SIM & ")"
		end if
	
	if ckb_entrega_imediata_nao <> "" then
		s_periodo_aux = ""
		if c_dt_previsao_entrega_inicio <> "" then
			s_periodo_aux = " (t_PEDIDO.PrevisaoEntregaData >= " & bd_formata_data(StrToDate(c_dt_previsao_entrega_inicio)) & ")"
			end if
		if c_dt_previsao_entrega_termino <> "" then
			if s_periodo_aux <> "" then s_periodo_aux = s_periodo_aux & " AND"
			s_periodo_aux = s_periodo_aux & " (t_PEDIDO.PrevisaoEntregaData < " & bd_formata_data(StrToDate(c_dt_previsao_entrega_termino)+1) & ")"
			end if
		if s_periodo_aux <> "" then s_periodo_aux = " AND" & s_periodo_aux
		if s <> "" then s = s & " OR"
		s = s & " ((t_PEDIDO.st_etg_imediata = " & COD_ETG_IMEDIATA_NAO & ")" & s_periodo_aux & ")"
		end if

	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CRIT�RIO: FORMA DE PAGAMENTO (NOVA VERS�O)
	s = ""
	if op_forma_pagto <> "" then
		s = " (t_PEDIDO__BASE.av_forma_pagto = " & op_forma_pagto & ")" & _
			" OR (t_PEDIDO__BASE.pu_forma_pagto = " & op_forma_pagto & ")" & _
			" OR (t_PEDIDO__BASE.pce_forma_pagto_entrada = " & op_forma_pagto & ")" & _
			" OR (t_PEDIDO__BASE.pce_forma_pagto_prestacao = " & op_forma_pagto & ")" & _
			" OR (t_PEDIDO__BASE.pse_forma_pagto_prim_prest = " & op_forma_pagto & ")" & _
			" OR (t_PEDIDO__BASE.pse_forma_pagto_demais_prest = " & op_forma_pagto & ")"
		if op_forma_pagto = ID_FORMA_PAGTO_CARTAO then
			s = s & " OR (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_CARTAO & ")"
		elseif op_forma_pagto = ID_FORMA_PAGTO_CARTAO_MAQUINETA then
			s = s & " OR (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA & ")"
			end if
		end if
	
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRIT�RIO: QUANTIDADE DE PARCELAS
	s = ""
	if c_forma_pagto_qtde_parc <> "" then
		s = " (t_PEDIDO__BASE.qtde_parcelas = " & c_forma_pagto_qtde_parc & ")"
		end if
	
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CRIT�RIO: PER�ODO DE CADASTRAMENTO DO PEDIDO
	s = ""
	if c_dt_cadastro_inicio <> "" then
		if s <> "" then s = s & " AND"
		s = s & " (t_PEDIDO.data >= " & bd_formata_data(StrToDate(c_dt_cadastro_inicio)) & ")"
		end if
		
	if c_dt_cadastro_termino <> "" then
		if s <> "" then s = s & " AND"
		s = s & " (t_PEDIDO.data < " & bd_formata_data(StrToDate(c_dt_cadastro_termino)+1) & ")"
		end if
	
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CRIT�RIO: DATA DE COLETA (R�TULO ANTIGO: ENTREGA MARCADA PARA)
	if ckb_entrega_marcada_para <> "" then
		s = ""
		if c_dt_entrega_inicio <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_entrega_inicio)) & ")"
			end if
		
		if c_dt_entrega_termino <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.a_entregar_data_marcada < " & bd_formata_data(StrToDate(c_dt_entrega_termino)+1) & ")"
			end if
		
		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
		end if
	
'	CRIT�RIO: PER�ODO DE EMISS�O DA NF DE VENDA
	if ckb_periodo_emissao_NF_venda <> "" then
		'PER�ODO NF
		'DEVIDO � FORMA COMO A DATA DE EMISS�O DA NF � OBTIDA, N�O � POSS�VEL APLICAR A RESTRI��O POR DATA NA CONSULTA BASE
		'PORTANTO, PARA MINIMIZAR A QUANTIDADE DE REGISTROS SELECIONADOS, FORAM ADOTADOS OS SEGUINTES CRIT�RIOS:
		'	1) AS RESTRI��ES S�O APLICADAS EM 2 MOMENTOS DISTINTOS: NA CONSULTA BASE (INTERNA) E NA CONSULTA GERAL (EXTERNA)
		'	2) CONSULTA BASE:
		'		2a) RESTRINGE POR PEDIDOS QUE TENHAM N�MERO DE NF
		'		2b) EXCLUI OS PEDIDOS CANCELADOS
		'	3) CONSULTA GERAL
		'		3a) AS DATAS DE IN�CIO E FIM DO PER�ODO S�O APLICADAS SOBRE A DATA DE EMISS�O RETORNADA PELA CONSULTA BASE
		s = ""
		if IsDate(c_dt_NF_venda_inicio) Or IsDate(c_dt_NF_venda_termino) then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.num_obs_2 > 0) AND (t_PEDIDO.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')"
			end if

		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if

		if IsDate(c_dt_NF_venda_inicio) then
			if s_where_ext <> "" then s_where_ext = s_where_ext & " AND"
			s_where_ext = s_where_ext & " (dt_emissao_venda >= " & bd_formata_data(StrToDate(c_dt_NF_venda_inicio)) & ")"
			end if

		if IsDate(c_dt_NF_venda_termino) then
			if s_where_ext <> "" then s_where_ext = s_where_ext & " AND"
			s_where_ext = s_where_ext & " (dt_emissao_venda < " & bd_formata_data(StrToDate(c_dt_NF_venda_termino)+1) & ")"
			end if
		end if

'	CRIT�RIO: PER�ODO DE EMISS�O DA NF DE REMESSA
	if ckb_periodo_emissao_NF_remessa <> "" then
		'PER�ODO NF
		'DEVIDO � FORMA COMO A DATA DE EMISS�O DA NF � OBTIDA, N�O � POSS�VEL APLICAR A RESTRI��O POR DATA NA CONSULTA BASE
		'PORTANTO, PARA MINIMIZAR A QUANTIDADE DE REGISTROS SELECIONADOS, FORAM ADOTADOS OS SEGUINTES CRIT�RIOS:
		'	1) AS RESTRI��ES S�O APLICADAS EM 2 MOMENTOS DISTINTOS: NA CONSULTA BASE (INTERNA) E NA CONSULTA GERAL (EXTERNA)
		'	2) CONSULTA BASE:
		'		2a) RESTRINGE POR PEDIDOS QUE TENHAM N�MERO DE NF
		'		2b) EXCLUI OS PEDIDOS CANCELADOS
		'	3) CONSULTA GERAL
		'		3a) AS DATAS DE IN�CIO E FIM DO PER�ODO S�O APLICADAS SOBRE A DATA DE EMISS�O RETORNADA PELA CONSULTA BASE
		s = ""
		if IsDate(c_dt_NF_remessa_inicio) Or IsDate(c_dt_NF_remessa_termino) then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.num_obs_3 > 0) AND (t_PEDIDO.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')"
			end if

		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if

		if IsDate(c_dt_NF_remessa_inicio) then
			if s_where_ext <> "" then s_where_ext = s_where_ext & " AND"
			s_where_ext = s_where_ext & " (dt_emissao_remessa >= " & bd_formata_data(StrToDate(c_dt_NF_remessa_inicio)) & ")"
			end if

		if IsDate(c_dt_NF_remessa_termino) then
			if s_where_ext <> "" then s_where_ext = s_where_ext & " AND"
			s_where_ext = s_where_ext & " (dt_emissao_remessa < " & bd_formata_data(StrToDate(c_dt_NF_remessa_termino)+1) & ")"
			end if
		end if

'	CRIT�RIO: PRODUTO
	blnPorFornecedor = False
	if ckb_produto <> "" then
		s = ""
		if c_fabricante <> "" then
			blnPorFornecedor = True
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO_ITEM.fabricante = '" & c_fabricante & "')"
			end if
		
		if c_produto <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO_ITEM.produto = '" & c_produto & "')"
			end if

		if ckb_somente_pedidos_produto_alocado <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (ISNULL(t_ESTOQUE_MOVIMENTO__AUX.qtde_produto_alocada,0) > 0)"
			end if

		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
		end if

        s = ""
	    if c_grupo <> "" then
	        v_grupos = split(c_grupo, ", ")
	        for cont = Lbound(v_grupos) to Ubound(v_grupos)
	            if s <> "" then s = s & " OR"
		        s = s & _
			        " (t_PRODUTO.grupo = '" & v_grupos(cont) & "')"
	        next
	        if s <> "" then 
			    if s_where <> "" then s_where = s_where & " AND"
			    s_where = s_where & " (" & s & ")"
			end if
        end if

' CRIT�RIO: ORIGEM DO PEDIDO (GRUPO)
    s = ""
    if c_grupo_pedido_origem <> "" then
        s_grupo_origem = "SELECT codigo FROM t_CODIGO_DESCRICAO WHERE (codigo_pai = '" & c_grupo_pedido_origem & "') AND grupo='PedidoECommerce_Origem'"
        if rs.State <> 0 then rs.Close
	    rs.open s_grupo_origem, cn
		if rs.Eof then
            alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_grupo_pedido_origem & " N�O EXISTE."
        else
            do while Not rs.Eof
                if s <> "" then s = s & ", "
                s = s & "'" & rs("codigo") & "'"      
                rs.MoveNext
            loop
            s = " t_PEDIDO.marketplace_codigo_origem IN (" & s & ")"
        end if
        if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
		end if
    end if

' CRIT�RIO: ORIGEM DO PEDIDO
    s = ""
    if c_pedido_origem <> "" then
        s = s & " t_PEDIDO.marketplace_codigo_origem = " & c_pedido_origem & ""

        if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
		end if
    end if
' CRIT�RIO: EMPRESA
    s = ""
    if c_empresa <> "" then
        s = s & " t_PEDIDO.id_nfe_emitente = '" & c_empresa & "'"

        if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
		end if
    end if

'	CRIT�RIO: LOJA
	if rb_loja = "UMA" then
		s = ""
		for i=LBound(vLoja) to UBound(vLoja)
			if Trim("" & vLoja(i)) <> "" then
				if s <> "" then s = s & ", "
				s = s & Trim("" & vLoja(i))
				end if
			next
		if s <> "" then
			s = " (t_PEDIDO.numero_loja IN (" & s & "))"
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
	elseif rb_loja = "FAIXA" then
		s = ""
		if c_loja_de <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.numero_loja >= " & c_loja_de & ")"
			end if

		if c_loja_ate <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.numero_loja <= " & c_loja_ate & ")"
			end if
		
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRIT�RIO: TRANSPORTADORA
	if c_transportadora <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.transportadora_id = '" & c_transportadora & "')"
		end if
	
'	CRIT�RIO: TRANSPORTADORAS
	if c_transportadora_multiplo <> "" then
		s_where_aux = ""
		vTransportadora = Split(c_transportadora_multiplo, ", ")
		for i=LBound(vTransportadora) to UBound(vTransportadora)
			if Trim(vTransportadora(i)) <> "" then
				if s_where_aux <> "" then s_where_aux = s_where_aux & ", "
				s_where_aux = s_where_aux & "'" & Trim(vTransportadora(i)) & "'"
				end if
			next
		if s_where_aux <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_PEDIDO.transportadora_id IN (" & s_where_aux & "))"
			end if
		end if

'	CRIT�RIO: CLIENTE
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		if c_cliente_cnpj_cpf <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_PEDIDO.endereco_cnpj_cpf = '" & retorna_so_digitos(c_cliente_cnpj_cpf) & "')"
			end if
		if c_cliente_uf <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_PEDIDO.endereco_uf = '" & c_cliente_uf & "')"
		end if
	else
		if c_cliente_cnpj_cpf <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE.cnpj_cpf = '" & retorna_so_digitos(c_cliente_cnpj_cpf) & "')"
			end if
		if c_cliente_uf <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_CLIENTE.uf = '" & c_cliente_uf & "')"
		end if
	end if

'	CRIT�RIO: CART�O DE CR�DITO (ANTIGAMENTE PELA VISANET, DEPOIS PELA CIELO E AGORA PELA BRASPAG)
	if ckb_visanet <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & _
					" (" & _
						"(" & _
							"(t_PEDIDO_PAGTO_VISANET__SUCESSO.operacao = '" & OP_VISANET_PAGAMENTO & "')" & _
							" AND (t_PEDIDO_PAGTO_VISANET__SUCESSO.concluido_status<>0)" & _
							" AND (t_PEDIDO_PAGTO_VISANET__SUCESSO.sucesso_status<>0)" & _
							" AND (t_PEDIDO_PAGTO_VISANET__SUCESSO.cancelado_status=0)" & _
						")" & _
						" OR " & _
						"(" & _
							"(t_PEDIDO_PAGTO_CIELO__SUCESSO.operacao = '" & OP_CIELO_OPERACAO__PAGAMENTO & "')" & _
							" AND (t_PEDIDO_PAGTO_CIELO__SUCESSO.sucesso_final_status<>0)" & _
							" AND (t_PEDIDO_PAGTO_CIELO__SUCESSO.cancelado_status=0)" & _
						")" & _
						" OR " & _
						"(" & _
							"(t_PEDIDO_PAGTO_BRASPAG__SUCESSO.operacao = '" & OP_BRASPAG_OPERACAO__AF_PAG & "')" & _
							" AND (t_PEDIDO_PAGTO_BRASPAG__SUCESSO.ult_PAG_GlobalStatus IN ('" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "','" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA & "'))" & _
						")" & _
					")"
		end if
	
'	CRIT�RIO: VENDEDOR
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.vendedor = '" & c_vendedor & "')"
		end if

'	CRIT�RIO: INDICADOR (LEMBRE-SE: O OR�AMENTISTA DE UM OR�AMENTO � USADO AUTOMATICAMENTE COMO O INDICADOR DO PEDIDO QUANDO O OR�AMENTO VIRA PEDIDO)
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if
		
'	CRIT�RIO: CAMPO OBS2 PREENCHIDO
	if ckb_obs2_preenchido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (RTrim(Coalesce(t_PEDIDO.obs_2,'')) <> '')"
		end if
		
'	CRIT�RIO: CAMPO OBS2 N�O PREENCHIDO
	if ckb_obs2_nao_preenchido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (RTrim(Coalesce(t_PEDIDO.obs_2,'')) = '')"
		end if
		
'	CRIT�RIO: CAMPO INDICADOR PREENCHIDO
	if ckb_indicador_preenchido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (RTrim(Coalesce(t_PEDIDO.indicador,'')) <> '')"
		end if
		
'	CRIT�RIO: CAMPO INDICADOR N�O PREENCHIDO
	if ckb_indicador_nao_preenchido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (RTrim(Coalesce(t_PEDIDO.indicador,'')) = '')"
		end if

'	CL�USULA WHERE
	if s_where <> "" then s_where = " WHERE" & s_where
	
	
'	MONTA CL�USULA FROM
	s_from = " FROM t_PEDIDO" & _
			 " INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			 " INNER JOIN t_NFe_EMITENTE ON (t_PEDIDO.id_nfe_emitente = t_NFe_EMITENTE.id)"
	
	if ckb_produto <> "" then
		s_from = s_from & " INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)"
		end if

    if c_grupo <> "" then
		if ckb_produto = "" then s_from = s_from & " INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)"
        s_from = s_from & " INNER JOIN t_PRODUTO ON (t_PEDIDO_ITEM.produto = t_PRODUTO.produto)"
    end if
	
	if c_cliente_cnpj_cpf <> "" Or c_cliente_uf <> "" then
		s_from = s_from & " INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)"
	else
		s_from = s_from & " LEFT JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)"
		end if
	
'	PAGAMENTO POR CART�O (ANTIGAMENTE PELA VISANET, DEPOIS PELA CIELO E AGORA PELA BRASPAG)
	if ckb_visanet <> "" then
		s_from = s_from & " LEFT JOIN (" & _
								"SELECT " & _
									"*" & _
								" FROM t_PEDIDO_PAGTO_VISANET" & _
								" WHERE" & _
									" (t_PEDIDO_PAGTO_VISANET.operacao = '" & OP_VISANET_PAGAMENTO & "')" & _
									" AND (t_PEDIDO_PAGTO_VISANET.concluido_status<>0)" & _
									" AND (t_PEDIDO_PAGTO_VISANET.sucesso_status<>0)" & _
									" AND (t_PEDIDO_PAGTO_VISANET.cancelado_status=0)" & _
								") AS t_PEDIDO_PAGTO_VISANET__SUCESSO ON (t_PEDIDO.pedido=t_PEDIDO_PAGTO_VISANET__SUCESSO.pedido)"
		
		s_from = s_from & " LEFT JOIN (" & _
								"SELECT " & _
									"*" & _
								" FROM t_PEDIDO_PAGTO_CIELO" & _
								" WHERE" & _
									" (t_PEDIDO_PAGTO_CIELO.operacao = '" & OP_CIELO_OPERACAO__PAGAMENTO & "')" & _
									" AND (t_PEDIDO_PAGTO_CIELO.sucesso_final_status<>0)" & _
									" AND (t_PEDIDO_PAGTO_CIELO.cancelado_status=0)" & _
								") AS t_PEDIDO_PAGTO_CIELO__SUCESSO ON (t_PEDIDO.pedido=t_PEDIDO_PAGTO_CIELO__SUCESSO.pedido)"
		
		s_from = s_from & " LEFT JOIN (" & _
								"SELECT " & _
									"*" & _
								" FROM t_PEDIDO_PAGTO_BRASPAG" & _
								" WHERE" & _
									" (t_PEDIDO_PAGTO_BRASPAG.operacao = '" & OP_BRASPAG_OPERACAO__AF_PAG & "')" & _
									" AND (t_PEDIDO_PAGTO_BRASPAG.ult_PAG_GlobalStatus IN ('" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "','" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__AUTORIZADA & "'))" & _
								") AS t_PEDIDO_PAGTO_BRASPAG__SUCESSO ON (t_PEDIDO.pedido=t_PEDIDO_PAGTO_BRASPAG__SUCESSO.pedido)"
		end if

'	CRIA UMA "DERIVED TABLE" PARA OBTER O TOTAL EM DEVOLU��ES DO PEDIDO
	s_from = s_from & _
			" LEFT JOIN (" & _
				"SELECT pedido," & _
				" Sum(qtde) AS qtde_produtos_devolvidos," & _
				" Sum(qtde*preco_venda) AS vl_devolucao_pedido," & _
				" Sum(qtde*preco_NF) AS vl_devolucao_pedido_NF" & _
				" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
				" GROUP BY pedido" & _
				") AS t_PEDIDO_ITEM_DEVOLVIDO__AUX" & _
			" ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO__AUX.pedido)"

'	CRIA UMA "DERIVED TABLE" PARA OBTER O VALOR TOTAL DO PEDIDO
	s_from = s_from & _
			" LEFT JOIN (" & _
				"SELECT t_PEDIDO_ITEM.pedido AS pedido," & _
				" Sum(qtde*preco_venda) AS vl_total_pedido," & _
				" Sum(qtde*preco_NF) AS vl_total_pedido_NF" & _
				" FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO" & _
				" ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
				" WHERE (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')" & _
				" GROUP BY t_PEDIDO_ITEM.pedido" & _
				") AS t_PEDIDO__VL_TOTAL" & _
			" ON (t_PEDIDO.pedido=t_PEDIDO__VL_TOTAL.pedido)"

'	CRIA UMA "DERIVED TABLE" PARA OBTER O TOTAL EM PAGAMENTOS DO PEDIDO
	s_from = s_from & _
			" LEFT JOIN (" & _
				"SELECT pedido," & _
				" Sum(valor) AS vl_pago_pedido" & _
				" FROM t_PEDIDO_PAGAMENTO" & _
				" GROUP BY pedido" & _
				") AS t_PEDIDO__VL_PAGO" & _
			" ON (t_PEDIDO.pedido=t_PEDIDO__VL_PAGO.pedido)"

'	CRIA UMA "DERIVED TABLE" PARA OBTER O VALOR TOTAL RELATIVO AO FORNECEDOR
	if blnPorFornecedor then
		s_from = s_from & _
				" LEFT JOIN (" & _
					"SELECT t_PEDIDO_ITEM.pedido AS pedido," & _
					" Sum(qtde*preco_venda) AS vl_total_fornecedor," & _
					" Sum(qtde*preco_NF) AS vl_total_fornecedor_NF" & _
					" FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO" & _
					" ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
					" WHERE (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')" & _
					" AND (fabricante = '" & c_fabricante & "')" & _
					" GROUP BY t_PEDIDO_ITEM.pedido" & _
					") AS t_PEDIDO__VL_FORNECEDOR" & _
				" ON (t_PEDIDO.pedido=t_PEDIDO__VL_FORNECEDOR.pedido)"
		end if

'	CRIA UMA "DERIVED TABLE" PARA OBTER A QUANTIDADE DE PRODUTO ALOCADO PARA O PEDIDO
	if (ckb_produto <> "") And (ckb_somente_pedidos_produto_alocado <> "") then
		s_from = s_from & _
				" LEFT JOIN (" & _
					"SELECT " & _
					" pedido, fabricante, produto," & _
					" Sum(qtde) AS qtde_produto_alocada" & _
					" FROM t_ESTOQUE_MOVIMENTO" & _
					" WHERE" & _
						" (anulado_status = 0)" & _
						" AND (estoque NOT IN ('" & ID_ESTOQUE_SEM_PRESENCA & "'))" & _
						" AND (fabricante = '" & c_fabricante & "')" & _
						" AND (produto = '" & c_produto & "')" & _
					" GROUP BY pedido, fabricante, produto" & _
					") AS t_ESTOQUE_MOVIMENTO__AUX" & _
				" ON (t_PEDIDO_ITEM.pedido=t_ESTOQUE_MOVIMENTO__AUX.pedido) AND (t_PEDIDO_ITEM.fabricante=t_ESTOQUE_MOVIMENTO__AUX.fabricante) AND (t_PEDIDO_ITEM.produto=t_ESTOQUE_MOVIMENTO__AUX.produto)"
		end if

'	OBS: SINTAXE DA FUN��O ISNULL():
'		 ISNULL ( check_expression , replacement_value )
'		 SE "check_expression" FOR NULL, RETORNA "replacement_value"
	s_sql = "SELECT DISTINCT t_PEDIDO.loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO.data, t_PEDIDO.nsu_pedido_base, t_PEDIDO.nsu, t_PEDIDO.pedido, t_PEDIDO.pedido_bs_x_ac, t_PEDIDO.obs_2, t_PEDIDO.obs_3," & _
			" t_PEDIDO.st_entrega, t_PEDIDO.PrevisaoEntregaData, t_PEDIDO.transportadora_id,"

	if ckb_periodo_emissao_NF_venda <> "" then
		s_sql = s_sql & _
				" (SELECT TOP 1 Convert(datetime, ide__dEmi, 121) FROM t_NFe_IMAGEM WHERE (t_NFe_IMAGEM.NFe_numero_NF = t_PEDIDO.num_obs_2) AND (t_NFe_IMAGEM.id_nfe_emitente = t_PEDIDO.id_nfe_emitente) AND (t_NFe_IMAGEM.ide__tpNF = '1') AND (t_NFe_IMAGEM.st_anulado = 0) AND (t_NFe_IMAGEM.codigo_retorno_NFe_T1 = 1) ORDER BY id DESC) AS dt_emissao_venda,"
		end if

	if ckb_periodo_emissao_NF_remessa <> "" then
		s_sql = s_sql & _
				" (SELECT TOP 1 Convert(datetime, ide__dEmi, 121) FROM t_NFe_IMAGEM WHERE (t_NFe_IMAGEM.NFe_numero_NF = t_PEDIDO.num_obs_3) AND (t_NFe_IMAGEM.id_nfe_emitente = t_PEDIDO.id_nfe_emitente) AND (t_NFe_IMAGEM.ide__tpNF = '1') AND (t_NFe_IMAGEM.st_anulado = 0) AND (t_NFe_IMAGEM.codigo_retorno_NFe_T1 = 1) ORDER BY id DESC) AS dt_emissao_remessa,"
		end if

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome AS nome," & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome, t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.st_end_entrega," & _
			" t_PEDIDO.endereco_cidade AS cidade_cliente," & _
			" t_PEDIDO.endereco_uf AS uf_cliente," & _
			" t_PEDIDO.EndEtg_cidade," & _
			" t_PEDIDO.EndEtg_uf," & _
			" t_NFe_EMITENTE.cnpj AS cnpj_emitente," & _
			" t_PEDIDO__BASE.st_pagto," & _
            " t_PEDIDO__BASE.vendedor," & _
            " t_PEDIDO__BASE.indicador," & _
            " t_PEDIDO.cancelado_codigo_motivo," & _
            " t_PEDIDO.cancelado_codigo_sub_motivo," & _
			" t_PEDIDO.entregue_data," & _
			" t_PEDIDO.PrevisaoEntregaTranspData," & _
			" t_PEDIDO.PedidoRecebidoStatus," & _
			" t_PEDIDO.PedidoRecebidoData," & _
			" ISNULL(t_PEDIDO__VL_TOTAL.vl_total_pedido,0) AS vl_total_pedido," & _
			" ISNULL(t_PEDIDO__VL_TOTAL.vl_total_pedido_NF,0) AS vl_total_pedido_NF," & _
			" ISNULL(t_PEDIDO__VL_PAGO.vl_pago_pedido,0) AS vl_pago_pedido," & _
			" ISNULL(t_PEDIDO_ITEM_DEVOLVIDO__AUX.vl_devolucao_pedido,0) AS vl_devolucao_pedido," & _
			" ISNULL(t_PEDIDO_ITEM_DEVOLVIDO__AUX.vl_devolucao_pedido_NF,0) AS vl_devolucao_pedido_NF," & _
			" ISNULL(t_PEDIDO_ITEM_DEVOLVIDO__AUX.qtde_produtos_devolvidos,0) AS qtde_produtos_devolvidos"

    if blnMostraMotivoCancelado then
        s_sql = s_sql & _            
            ", Coalesce((SELECT Sum(qtde * preco_venda) FROM t_PEDIDO_ITEM WHERE (pedido = t_PEDIDO.pedido)), 0) AS vl_total_original"
    end if
	
	if blnPorFornecedor then
		s_sql = s_sql & _
				", ISNULL(t_PEDIDO__VL_FORNECEDOR.vl_total_fornecedor,0) AS vl_total_fornecedor" & _
				", ISNULL(t_PEDIDO__VL_FORNECEDOR.vl_total_fornecedor_NF,0) AS vl_total_fornecedor_NF"
		end if
	
	if ckb_exibir_qtde_volumes <> "" then
		s_sql = s_sql & _
				", Coalesce((SELECT Sum(qtde*qtde_volumes) FROM t_PEDIDO_ITEM WHERE (pedido = t_PEDIDO.pedido)), 0) AS total_pedido_qtde_volumes"
		end if

	if ckb_exibir_peso <> "" then
		s_sql = s_sql & _
				", Coalesce((SELECT Sum(qtde*peso) FROM t_PEDIDO_ITEM WHERE (pedido = t_PEDIDO.pedido)), 0) AS total_pedido_peso"
		end if

	if ckb_exibir_cubagem <> "" then
		s_sql = s_sql & _
				", Coalesce((SELECT Sum(qtde*cubagem) FROM t_PEDIDO_ITEM WHERE (pedido = t_PEDIDO.pedido)), 0) AS total_pedido_cubagem"
		end if

	s_sql = s_sql & _
			s_from & _
			s_where

	s_sql = "SELECT " & _
				"*" & _
			" FROM (" & s_sql & ") t"

	if s_where_ext <> "" then
		s_sql = s_sql & _
				" WHERE" & _
				s_where_ext
		end if

    if ckb_st_entrega_cancelado <> "" then
        if c_cancelados_ordena = "VENDEDOR" then
            s_sql = s_sql & " ORDER BY numero_loja, vendedor, indicador, data, nsu_pedido_base, nsu, pedido"
        else
	        s_sql = s_sql & " ORDER BY numero_loja, data, nsu_pedido_base, nsu, pedido"
        end if
    else
	    s_sql = s_sql & " ORDER BY numero_loja, data, nsu_pedido_base, nsu, pedido"
    end if

  ' CABE�ALHO
	w_pedido = 70
	w_pedido_magento = 70
	w_data = 70
	w_NF = 50
	w_vendedor = 80
	w_indicador = 80
	w_uf_cad_cliente = 30
	w_cidade_etg = 80
	w_uf_etg = 30
	w_qtde_vol = 40
	w_cubagem = 40
	w_peso = 60

	if blnPorFornecedor then
		if blnRelAnalitico then
			w_cliente = 201
		else
			w_cliente = 400
			end if
		w_st_entrega = 74
		w_valor = 70
	else
		if blnRelAnalitico then
			w_cliente = 250
		else
			w_cliente = 400
			end if
		w_st_entrega = 70
		w_valor = 80
		end if
	
	if blnSaidaExcel then
		w_pedido = 80
		w_pedido_magento = 90
		w_data = 80
		w_NF = 70
		w_valor = 120
		w_st_entrega = 100
		end if
	
	w_motivo_cancelamento = 200

	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure'>" & chr(13) & _
		  "		<TD valign='bottom' style='background:white;' NOWRAP>&nbsp;</TD>" & chr(13) & _
		  "     <TD class='MT' style='width:" & Cstr(w_pedido) & "px' valign='bottom' NOWRAP><span class='R' style='font-weight:bold;'>N� Pedido</span></TD>" & chr(13) & _
          "<!--Magento-->" & chr(13) & _
		  "		<TD class='MTBD' align='center' style='width:" & Cstr(w_data) & "px' valign='bottom'><span class='R' style='font-weight:bold;'>Data</span></TD>" & chr(13) & _
		  "		<td class='MTBD' align='center' style='width:" & Cstr(w_NF) & "px' valign='bottom'><span class='R' style='font-weight:bold;'>NF</span></TD>" & chr(13) & _
		  "		<TD class='MTBD' style='width:" & Cstr(w_cliente) & "px' valign='bottom'><span class='R' style='font-weight:bold;'>Cliente</span></TD>" & chr(13)
	
	if blnPorFornecedor then
		if blnRelAnalitico then
			cab = cab & _ 
				  "		<TD class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL Fornec</span></TD>" & chr(13) & _
				  "		<TD class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL Fornec (RA)</span></TD>" & chr(13) & _
				  "		<TD class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL Pedido</span></TD>" & chr(13) & _
				  "		<TD class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL Pedido (RA)</span></TD>" & chr(13)
			end if
	else
		if blnRelAnalitico then
			cab = cab & _ 
				  "		<TD class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>Valor</span></TD>" & chr(13) & _
				  "		<TD class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>Valor (RA)</span></TD>" & chr(13)
			end if
		end if
	
	if blnRelAnalitico then
		cab = cab & _ 
			  "		<TD class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL Pago</span></TD>" & chr(13) & _
			  "		<TD class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL A Pagar</span></TD>" & chr(13)
		end if
	
	if blnSaidaExcel then
		cab = cab & _
			  "		<TD class='MTBD' style='width:" & Cstr(w_st_entrega) & "px' valign='bottom'><span class='R' style='font-weight:bold;'>Status de<br style='mso-data-placement:same-cell;' />Entrega</span></TD>" & chr(13)
	else
		cab = cab & _
			  "		<TD class='MTBD' style='width:" & Cstr(w_st_entrega) & "px' valign='bottom' NOWRAP><span class='R' style='display:block;font-weight:bold;'>Status de<br />Entrega</span></TD>" & chr(13)
		end if

	if ckb_exibir_data_entrega <> "" then
		if blnSaidaExcel then
			cab = cab & _
					"		<td class='MTBD' style='width:" & Cstr(w_data) & "px' align='center' valign='bottom'><span class='R' style='font-weight:bold;'>Data<br style='mso-data-placement:same-cell;' />Entrega</span></td>" & chr(13)
		else
			cab = cab & _
					"		<td class='MTBD' style='width:" & Cstr(w_data) & "px' align='center' valign='bottom'><span class='R' style='display:block;font-weight:bold;'>Data<br />Entrega</span></td>" & chr(13)
			end if
		end if

    if ckb_exibir_data_previsao_entrega <> "" then
		if blnSaidaExcel then
			cab = cab & _
					"		<TD class='MTBD' style='width:" & Cstr(w_data) & "px' align='center' valign='bottom'><span class='R' style='font-weight:bold;'>Previs�o de<br style='mso-data-placement:same-cell;' />Entrega</span></TD>" & chr(13)
		else
			cab = cab & _
				"		<td class='MTBD' style='width:" & Cstr(w_data) & "px' align='center' valign='bottom' NOWRAP><span class='R' style='display:block;font-weight:bold;'>Previs�o de<br />Entrega</span></td>" & chr(13)
			end if
		end if

	if ckb_exibir_data_previsao_entrega_transp <> "" then
		if blnSaidaExcel then
			cab = cab & _
					"		<TD class='MTBD' style='width:" & Cstr(w_data) & "px' align='center' valign='bottom'><span class='R' style='font-weight:bold;'>Previs�o de<br style='mso-data-placement:same-cell;' />Entrega (Transp)</span></TD>" & chr(13)
		else
			cab = cab & _
				"		<td class='MTBD' style='width:" & Cstr(w_data) & "px' align='center' valign='bottom' NOWRAP><span class='R' style='display:block;font-weight:bold;'>Previs�o de<br />Entrega<br />(Transp)</span></td>" & chr(13)
			end if
		end if

	if ckb_exibir_data_recebido_cliente <> "" then
		if blnSaidaExcel then
			cab = cab & _
					"		<TD class='MTBD' style='width:" & Cstr(w_data) & "px' align='center' valign='bottom'><span class='R' style='font-weight:bold;'>Receb<br style='mso-data-placement:same-cell;' />Cliente</span></TD>" & chr(13)
		else
			cab = cab & _
				"		<td class='MTBD' style='width:" & Cstr(w_data) & "px' align='center' valign='bottom' NOWRAP><span class='R' style='display:block;font-weight:bold;'>Receb<br />Cliente</span></td>" & chr(13)
			end if
		end if

	if ckb_exibir_cidade_etg <> "" then
		if blnSaidaExcel then
			cab = cab & _
					"		<td class='MTBD' style='width:" & Cstr(w_cidade_etg) & "px' valign='bottom'><span class='R' style='font-weight:bold;'>Cidade<br style='mso-data-placement:same-cell;' />(Etg)</span></td>" & chr(13)
		else
			cab = cab & _
					"		<td class='MTBD' style='width:" & Cstr(w_cidade_etg) & "px' valign='bottom' NOWRAP><span class='R' style='display:block;font-weight:bold;'>Cidade<br />(Etg)</span></td>" & chr(13)
			end if
		end if

	if ckb_exibir_uf_etg <> "" then
		if blnSaidaExcel then
			cab = cab & _
					"		<td class='MTBD' style='width:" & Cstr(w_uf_etg) & "px' valign='bottom'><span class='R' style='font-weight:bold;'>UF<br style='mso-data-placement:same-cell;' />(Etg)</span></td>" & chr(13)
		else
			cab = cab & _
					"		<td class='MTBD' style='width:" & Cstr(w_uf_etg) & "px' valign='bottom'><span class='R' style='display:block;font-weight:bold;'>UF<br />(Etg)</span></td>" & chr(13)
			end if
		end if

	if ckb_exibir_qtde_volumes <> "" then
		if blnSaidaExcel then
			cab = cab & _
					"		<TD class='MTBD' style='width:" & Cstr(w_qtde_vol) & "px' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>Qtde<br style='mso-data-placement:same-cell;' />Vol</span></TD>" & chr(13)
		else
			cab = cab & _
				"		<td class='MTBD' style='width:" & Cstr(w_qtde_vol) & "px' align='right' valign='bottom' NOWRAP><span class='Rd' style='display:block;font-weight:bold;'>Qtde<br />Vol</span></td>" & chr(13)
			end if
		end if

	if ckb_exibir_cubagem <> "" then
		cab = cab & _
			"		<td class='MTBD' style='width:" & Cstr(w_cubagem) & "px' align='right' valign='bottom' NOWRAP><span class='Rd' style='font-weight:bold;'>Cubagem</span></td>" & chr(13)
		end if

	if ckb_exibir_peso <> "" then
		cab = cab & _
			"		<td class='MTBD' style='width:" & Cstr(w_peso) & "px' align='right' valign='bottom' NOWRAP><span class='Rd' style='font-weight:bold;'>Peso</span></td>" & chr(13)
		end if

    if blnMostraMotivoCancelado then
        cab = cab & _
                "		<td class='MTBD' style='width:" & Cstr(w_vendedor) & "px' valign='bottom' NOWRAP><span class='R' style='font-weight:bold;'>Vendedor</span></td>" & chr(13) & _
                "		<td class='MTBD' style='width:" & Cstr(w_indicador) & "px' valign='bottom' NOWRAP><span class='R' style='font-weight:bold;'>Indicador</span></td>" & chr(13)
        if ckb_exibir_uf <> "" then
			cab = cab & _
				"		<td class='MTBD' style='width:" & Cstr(w_uf_cad_cliente) & "px' valign='bottom' NOWRAP><span class='R' style='display:block;font-weight:bold;'>UF<br />(Cad<br />Cliente)</span></td>" & chr(13)
			end if
		if blnRelAnalitico then
            cab = cab & _
                "		<TD class='MTBD' style='width:" & Cstr(w_valor) & "px' align='right' valign='bottom' NOWRAP><span class='Rd' style='font-weight:bold;'>VL Original</span></TD>" & chr(13)
        end if
        cab = cab & _
            "		<TD class='MTBD' style='width:" & Cstr(w_motivo_cancelamento) & "px' valign='bottom' NOWRAP><span class='R' style='font-weight:bold;'>Motivo Cancelamento</span></TD>" & chr(13)
    else
		if ckb_exibir_vendedor <> "" then
			cab = cab & _
					"		<td class='MTBD' style='width:" & Cstr(w_vendedor) & "px' valign='bottom' NOWRAP><span class='R' style='font-weight:bold;'>Vendedor</span></td>" & chr(13)
			end if

		if ckb_exibir_parceiro <> "" then
			cab = cab & _
                "		<td class='MTBD' style='width:" & Cstr(w_indicador) & "px' valign='bottom' NOWRAP><span class='R' style='font-weight:bold;'>Indicador</span></td>" & chr(13)
			end if

		if ckb_exibir_uf <> "" then
			if blnSaidaExcel then
				cab = cab & _
					"		<td class='MTBD' style='width:" & Cstr(w_uf_cad_cliente) & "px' valign='bottom'><span class='R' style='font-weight:bold;'>UF<br style='mso-data-placement:same-cell;' />(Cad<br style='mso-data-placement:same-cell;' />Cliente)</span></td>" & chr(13)
			else
				cab = cab & _
					"		<td class='MTBD' style='width:" & Cstr(w_uf_cad_cliente) & "px' valign='bottom'><span class='R' style='display:block;font-weight:bold;'>UF<br />(Cad<br />Cliente)</span></td>" & chr(13)
				end if
			end if
	end if
	
	cab = cab & _
		  "	</TR>" & chr(13)
	
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_lojas = 0
	vl_total_faturamento = 0
	vl_total_faturamento_NF = 0
	vl_total_pago = 0
	vl_total_a_pagar = 0
	vl_total_fornecedor = 0
	vl_total_fornecedor_NF = 0
    vl_total_pedido_original = 0
	total_qtde_vol = 0
	total_cubagem = 0
	total_peso = 0
	intNumLinha = 0
	
	loja_a = "XXXXX"
	if blnSaidaExcel then s_nbsp = "" else s_nbsp = "&nbsp;"

	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE LOJA?
		if Trim("" & r("loja"))<>loja_a then
			loja_a = Trim("" & r("loja"))
			qtde_lojas = qtde_lojas + 1
		  ' FECHA TABELA DA LOJA ANTERIOR
			if n_reg > 0 then 
				if blnRelAnalitico then
					s_cor = ""
					if vl_sub_total_a_pagar < 0 then s_cor = " style='color:red;'"
					x = x & "	<TR class='RowTotalizacao' style='background: #FFFFDD'>" & chr(13) & _
							"		<TD style='background:white;'>&nbsp;</td>" & chr(13)
                    if s_loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
							x = x & "		<TD class='MEB' align='right' COLSPAN='5' NOWRAP><span class='Cd' style='font-weight:bold;'>TOTAL:</span></td>" & chr(13)
					else
                            x = x & "		<TD class='MEB' align='right' COLSPAN='4' NOWRAP><span class='Cd' style='font-weight:bold;'>TOTAL:</span></td>" & chr(13)
                    end if
					if blnPorFornecedor then
						x = x & _
							"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_fornecedor) & "</span></td>" & chr(13) & _
							"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_fornecedor_NF) & "</span></td>" & chr(13)
						end if
					
					x = x & _
							"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_faturamento) & "</span></td>" & chr(13) & _
							"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_faturamento_NF) & "</span></td>" & chr(13) & _
							"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_pago) & "</span></td>" & chr(13) & _
							"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd'" & s_cor & " style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_a_pagar) & "</span></td>" & chr(13)

					x = x & _
							"		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)

					if ckb_exibir_data_entrega <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_data_previsao_entrega <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_data_previsao_entrega_transp <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_data_recebido_cliente <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_cidade_etg <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_uf_etg <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_qtde_volumes <> "" then
						x = x & _
							"		<TD align='right' style='width:" & Cstr(w_qtde_vol) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(sub_total_qtde_vol) & "</span></td>" & chr(13)
						end if
					if ckb_exibir_cubagem <> "" then
						x = x & _
							"		<TD align='right' style='width:" & Cstr(w_cubagem) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_DECIMAL & chr(34) & ";'>" & formata_numero(sub_total_cubagem, 2) & "</span></td>" & chr(13)
						end if
					if ckb_exibir_peso <> "" then
						x = x & _
							"		<TD align='right' style='width:" & Cstr(w_peso) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_DECIMAL & chr(34) & ";'>" & formata_numero(sub_total_peso, 2) & "</span></td>" & chr(13)
						end if

                    if blnMostraMotivoCancelado then
                        x = x & _
                            "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13) & _
                            "       <TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
						if ckb_exibir_uf <> "" then
							x = x & _
								"       <TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
							end if
						x = x & _
							"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_pedido_original) & "</span></td>" & chr(13) & _
							"		<TD class='MDB'><span class='C'>&nbsp;</span></td>" & chr(13)
                    else
						if ckb_exibir_vendedor <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
						if ckb_exibir_parceiro <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
						if ckb_exibir_uf <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
						end if

                    x = x & _
							"	</TR>" & chr(13)

					end if
				
				x = x & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			n_reg = 0
			vl_sub_total_faturamento = 0
			vl_sub_total_faturamento_NF = 0
			vl_sub_total_pago = 0
			vl_sub_total_a_pagar = 0
			vl_sub_total_fornecedor = 0
            vl_sub_total_pedido_original = 0
			vl_sub_total_fornecedor_NF = 0
			sub_total_qtde_vol = 0
			sub_total_cubagem = 0
			sub_total_peso = 0

            s_loja = Trim("" & r("loja"))
			if n_reg_total > 0 then x = x & "<BR>" & chr(13)
			s = Trim("" & r("loja"))
			s_aux = x_loja(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then
				if blnPorFornecedor then
					if blnRelAnalitico then 
                        if blnMostraMotivoCancelado then
                            n_colspan = 16
                        else
						    n_colspan = 12
                        end if
					else
						if blnMostraMotivoCancelado then
                            n_colspan = 9
                        else
						    n_colspan = 6
                            end if
						end if
				else 
					if blnRelAnalitico then
                        if s_loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
						    if blnMostraMotivoCancelado then
								n_colspan = 14
                            else
						        n_colspan = 10
                            end if 
                        else
                            if blnMostraMotivoCancelado then
                                n_colspan = 13
                            else
					            n_colspan = 9
                            end if
                        end if
					else
						if blnMostraMotivoCancelado then
                            n_colspan = 9
                        else
						    n_colspan = 6
                            end if
						end if
					end if
				
				if ckb_exibir_data_entrega <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_data_previsao_entrega <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_cidade_etg <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_uf_etg <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_data_previsao_entrega_transp <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_data_recebido_cliente <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_qtde_volumes <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_peso <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_cubagem <> "" then n_colspan = n_colspan + 1

				if Not blnMostraMotivoCancelado then
					if ckb_exibir_vendedor <> "" then n_colspan = n_colspan + 1
					if ckb_exibir_parceiro <> "" then n_colspan = n_colspan + 1
					if ckb_exibir_uf <> "" then n_colspan = n_colspan + 1
				else
					if ckb_exibir_uf <> "" then n_colspan = n_colspan + 1
					end if

				if blnSaidaExcel then 
					s_bkg_color = "tomato"
					s_align = " align='center'"
				else
					s_bkg_color = "azure"
					s_align = ""
					end if
				x = x & _
					"	<TR>" & chr(13) & _
					"		<TD style='background:white;'>" & s_nbsp & "</td>" & chr(13) & _
					"		<TD class='MDTE' COLSPAN='" & Cstr(n_colspan) & "'" & s_align & " valign='bottom' style='background:" & s_bkg_color & ";'><span class='N' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</span></td>" & chr(13) & _
					"	</TR>" & chr(13)
				end if

            
            if s_loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
                cab = Replace(cab, "<!--Magento-->", "		<td class='MTBD' style='width:" & Cstr(w_pedido_magento) & "px;font-weight:bold' align='left' valign='bottom'><span class='R'>N�mero Magento</span></td>")
            else
                cab = Replace(cab,"		<td class='MTBD' style='width:" & Cstr(w_pedido_magento) & "px;font-weight:bold' align='left' valign='bottom'><span class='R'>N�mero Magento</span></td>", "<!--Magento-->")
            end if
			x = x & cab
			end if

	 ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1
		intNumLinha = intNumLinha + 1

		if blnSaidaExcel OR (ckb_nao_exibir_links <> "") then
			x = x & "	<TR>" & chr(13)
		else
			x = x & "	<TR onmouseover='realca_cor_mouse_over(this);' onmouseout='realca_cor_mouse_out(this);'>" & chr(13)
			end if

	'> N� DA LINHA
		if blnSaidaExcel then
			x = x & "		<TD valign='middle' align='right' NOWRAP><span class='Rd' style='margin-right:2px;color:gray;font-style:italic;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & Cstr(intNumLinha) & ". </span></TD>" & chr(13)
		else
			x = x & "		<TD valign='middle' align='right' NOWRAP><span class='Rd' style='margin-right:2px;'>" & Cstr(intNumLinha) & ".</span></TD>" & chr(13)
			end if
			
			
	'> N� PEDIDO
		if blnSaidaExcel then
			x = x & "		<TD valign='middle' class='MDBE'><span class='C' style='font-weight:bold;'>" & Trim("" & r("pedido")) & "</span></TD>" & chr(13)
		else
			if ckb_nao_exibir_links <> "" then
				sLinkView = ""
			else
				sLinkView = monta_link_view_pedido(Trim("" & r("pedido")), usuario)
				end if
			if sLinkView = "" then
				x = x & "		<TD valign='middle' class='MDBE' nowrap>" & _
						"<span class='C'>&nbsp;<a href='javascript:fRELConcluir(" & chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & Trim("" & r("pedido")) & "</a></span>" & _
						"</TD>" & chr(13)
			else
				x = x & "		<TD valign='middle' class='MDBE' nowrap>" & chr(13) & _
						"			<table with='100%' cellpadding='0' cellspacing='0'>" & chr(13) & _
						"				<tr>" & chr(13) & _
						"					<td width='90%' align='left'>" & chr(13) & _
						"<span class='C'>&nbsp;<a href='javascript:fRELConcluir(" & chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & Trim("" & r("pedido")) & "</a></span>" & "</td>" & chr(13) & _
						"					<td align='right'>" & sLinkView & "</td>" & chr(13) & _
						"				</tr>" & chr(13) & _
						"			</table>" & chr(13) & _
						"</TD>" & chr(13)
				end if
			end if

    '> PEDIDO MAGENTO
        if Trim("" & r("loja")) = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
		    x = x & "		<td align='left' valign='middle' class='MDB'><span class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>&nbsp;" & Trim("" & r("pedido_bs_x_ac")) & "</span></td>" & chr(13)
        end if
	
	'> DATA DO PEDIDO
	    x = x & "		<TD align='center' valign='middle' class='MDB'><span class='Cn''>" & formata_data(r("data")) & "</span></TD>" & chr(13)
		
	'> NF
		if (ckb_nao_exibir_links <> "") Or blnSaidaExcel then
			s_numero_NF = Trim("" & r("obs_2"))
			if (s_numero_NF <> "") And (Trim("" & r("obs_3")) <> "") then s_numero_NF = s_numero_NF & ", "
			s_numero_NF = s_numero_NF & Trim("" & r("obs_3"))
		else
			s_link_rastreio = monta_link_rastreio_do_emitente(Trim("" & r("cnpj_emitente")), Trim("" & r("obs_2")), Trim("" & r("transportadora_id")), Trim("" & rPSSW.campo_texto), Trim("" & r("loja")))
			if s_link_rastreio <> "" then s_link_rastreio = "&nbsp;" & s_link_rastreio
			s_link_rastreio = Trim("" & r("obs_2")) & s_link_rastreio
			s_link_rastreio2 = monta_link_rastreio_do_emitente(Trim("" & r("cnpj_emitente")), Trim("" & r("obs_3")), Trim("" & r("transportadora_id")), Trim("" & rPSSW.campo_texto), Trim("" & r("loja")))
			if s_link_rastreio2 <> "" then s_link_rastreio2 = "&nbsp;" & s_link_rastreio2
			s_link_rastreio2 = Trim("" & r("obs_3")) & s_link_rastreio2
			if (s_link_rastreio <> "") And (s_link_rastreio2 <> "") then s_link_rastreio = s_link_rastreio & "<br />"
			s_numero_NF = s_link_rastreio & s_link_rastreio2
			end if
		x = x & "		<TD align='left' valign='middle' class='MDB'><span class='Cn' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_numero_NF & "</span></TD>" & chr(13)
		
	'> CLIENTE
		if blnSaidaExcel then s_nowrap = " NOWRAP" else s_nowrap = ""
		s = Trim("" & r("nome_iniciais_em_maiusculas"))
		if (s = "") And (Not blnSaidaExcel) then s = "&nbsp;"
		x = x & "		<TD valign='middle' style='width:" & Cstr(w_cliente) & "px' class='MDB'" & s_nowrap & "><span class='Cn' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s & "</span></TD>" & chr(13)

	'> VALOR DO FORNECEDOR
		if blnPorFornecedor then
			if blnRelAnalitico then
				s = formata_moeda(r("vl_total_fornecedor"))
				x = x & "		<TD valign='middle' align='right' style='width:" & Cstr(w_valor) & "px' class='MDB'><span class='Cnd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & s & "</span></TD>" & chr(13)
				end if
			end if
		
	'> VALOR DO FORNECEDOR COM RA
		if blnPorFornecedor then
			if blnRelAnalitico then
				s = formata_moeda(r("vl_total_fornecedor_NF"))
				x = x & "		<TD valign='middle' align='right' style='width:" & Cstr(w_valor) & "px' class='MDB'><span class='Cnd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & s & "</span></TD>" & chr(13)
				end if
			end if
		
	'> VALOR DO PEDIDO
		if blnRelAnalitico then
			s = formata_moeda(r("vl_total_pedido"))
			x = x & "		<TD valign='middle' align='right' style='width:" & Cstr(w_valor) & "px' class='MDB'><span class='Cnd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & s & "</span></TD>" & chr(13)
			end if
		
	'> VALOR DO PEDIDO COM RA
		if blnRelAnalitico then
			s = formata_moeda(r("vl_total_pedido_NF"))
			x = x & "		<TD valign='middle' align='right' style='width:" & Cstr(w_valor) & "px' class='MDB'><span class='Cnd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & s & "</span></TD>" & chr(13)
			end if
		
	'> VALOR J� PAGO
		if blnRelAnalitico then
			s = formata_moeda(r("vl_pago_pedido"))
			x = x & "		<TD valign='middle' align='right' style='width:" & Cstr(w_valor) & "px' class='MDB'><span class='Cnd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & s & "</span></TD>" & chr(13)
			end if
		
	'> VALOR A PAGAR
		vl_a_pagar = 0
		s_cor = ""
		vl_a_pagar = r("vl_total_pedido_NF")-r("vl_pago_pedido")-r("vl_devolucao_pedido_NF")
	'	VALORES NEGATIVOS REPRESENTAM O 'CR�DITO' QUE O CLIENTE POSSUI EM CASO DE PEDIDOS CANCELADOS QUE HAVIAM SIDO PAGOS
		if (Trim("" & r("st_pagto")) = ST_PAGTO_PAGO) And (vl_a_pagar > 0)  then vl_a_pagar = 0
		s = formata_moeda(vl_a_pagar)
		if blnRelAnalitico then
			if vl_a_pagar < 0 then s_cor = "color:red;"
			x = x & "		<TD valign='middle' align='right' style='width:" & Cstr(w_valor) & "px' class='MDB'><span class='Cnd' style='" & s_cor & "mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & s & "</span></TD>" & chr(13)
			end if

	'> STATUS DE ENTREGA
		s = Trim("" & r("st_entrega"))
		if s <> "" then 
			s = x_status_entrega(s)
			if (Trim("" & r("st_entrega"))=ST_ENTREGA_ENTREGUE) And (converte_numero(r("qtde_produtos_devolvidos"))>0) then s = s & " (*)"
			end if
		if (s = "") And (Not blnSaidaExcel) then s = "&nbsp;"
		x = x & "		<TD valign='middle' style='width:" & Cstr(w_st_entrega) & "px' class='MDB'><span class='Cn' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s & "</span></TD>" & chr(13)
	
	'> DATA DA ENTREGA (OPCIONAL)
		if ckb_exibir_data_entrega <> "" then
			if Trim("" & r("st_entrega")) = ST_ENTREGA_ENTREGUE then
				s = formata_data(r("entregue_data"))
			else
				s = "&nbsp;"
				end if
			x = x & "		<TD align='center' valign='middle' style='width:" & Cstr(w_data) & "px' class='MDB'><span class='Cn'>" & s & "</span></TD>" & chr(13)
			end if

	'> DATA PREVIS�O DE ENTREGA (OPCIONAL)
		if ckb_exibir_data_previsao_entrega <> "" then
			s = formata_data(r("PrevisaoEntregaData"))
			x = x & "		<TD align='center' valign='middle' style='width:" & Cstr(w_data) & "px' class='MDB'><span class='Cn'>" & s & "</span></TD>" & chr(13)
			end if

	'> DATA PREVIS�O ENTREGA TRANSPORTADORA (OPCIONAL)
		s_cor = ""
		if (ckb_exibir_data_previsao_entrega_transp <> "") And (ckb_exibir_data_recebido_cliente <> "") then
			if IsDate(r("PrevisaoEntregaTranspData")) And IsDate(r("PedidoRecebidoData")) then
				if r("PedidoRecebidoData") > r("PrevisaoEntregaTranspData") then s_cor = " style='color:red;'"
				end if
			end if

		if ckb_exibir_data_previsao_entrega_transp <> "" then
			s = formata_data(r("PrevisaoEntregaTranspData"))
			x = x & "		<TD align='center' valign='middle' style='width:" & Cstr(w_data) & "px' class='MDB'><span class='Cn' " & s_cor & ">" & s & "</span></TD>" & chr(13)
			end if

	'> DATA DE RECEBIMENTO DO PEDIDO PELO CLIENTE (OPCIONAL)
		if ckb_exibir_data_recebido_cliente <> "" then
			if r("PedidoRecebidoStatus") = 1 then
				s = formata_data(r("PedidoRecebidoData"))
			else
				s = "&nbsp;"
				end if
			x = x & "		<TD align='center' valign='middle' style='width:" & Cstr(w_data) & "px' class='MDB'><span class='Cn'>" & s & "</span></TD>" & chr(13)
			end if

	'> CIDADE (ENTREGA) (OPCIONAL)
		if ckb_exibir_cidade_etg <> "" then
			if CInt(r("st_end_entrega")) <> 0 then
				s = Trim("" & r("EndEtg_cidade"))
			else
				s = Trim("" & r("cidade_cliente"))
				end if
			x = x & "		<TD valign='middle' style='width:" & Cstr(w_cidade_etg) & "px' class='MDB'><span class='Cn'>" & s & "</span></TD>" & chr(13)
			end if

	'> UF (ENTREGA) (OPCIONAL)
		if ckb_exibir_uf_etg <> "" then
			if CInt(r("st_end_entrega")) <> 0 then
				s = Trim("" & r("EndEtg_uf"))
			else
				s = Trim("" & r("uf_cliente"))
				end if
			x = x & "		<TD valign='middle' style='width:" & Cstr(w_uf_etg) & "px' class='MDB'><span class='Cn'>" & s & "</span></TD>" & chr(13)
			end if

	'> QTDE VOLUMES (OPCIONAL)
		if ckb_exibir_qtde_volumes <> "" then
			s = formata_inteiro(r("total_pedido_qtde_volumes"))
			x = x & "		<TD align='right' valign='middle' style='width:" & Cstr(w_qtde_vol) & "px' class='MDB'><span class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & s & "</span></TD>" & chr(13)
			end if

	'> CUBAGEM (OPCIONAL)
		if ckb_exibir_cubagem <> "" then
			s = formata_numero(r("total_pedido_cubagem"), 2)
			x = x & "		<TD align='right' valign='middle' style='width:" & Cstr(w_cubagem) & "px' class='MDB'><span class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_DECIMAL & chr(34) & ";'>" & s & "</span></TD>" & chr(13)
			end if

	'> PESO (OPCIONAL)
		if ckb_exibir_peso <> "" then
			s = formata_numero(r("total_pedido_peso"), 2)
			x = x & "		<TD align='right' valign='middle' style='width:" & Cstr(w_peso) & "px' class='MDB'><span class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_DECIMAL & chr(34) & ";'>" & s & "</span></TD>" & chr(13)
			end if

    '> VENDEDOR
        if blnMostraMotivoCancelado then
            s = Trim("" & r("vendedor"))
            x = x & "		<TD valign='middle' style='width:" & Cstr(w_vendedor) & "px' class='MDB'><span class='Cn'>" & s & "</span></TD>" & chr(13)
        end if

    '> INDICADOR
        if blnMostraMotivoCancelado then
            s = Trim("" & r("indicador"))
            x = x & "		<TD valign='middle' style='width:" & Cstr(w_indicador) & "px' class='MDB'><span class='Cn'>" & s & "</span></TD>" & chr(13)
        end if

	'> UF (OPCIONAL)
		if blnMostraMotivoCancelado then
			if ckb_exibir_uf <> "" then
				s = Trim("" & r("uf_cliente"))
				x = x & "		<TD valign='middle' style='width:" & Cstr(w_uf_cad_cliente) & "px' class='MDB'><span class='Cn'>" & s & "</span></TD>" & chr(13)
				end if
			end if

    '> VALOR ORIGINAL DO PEDIDO
        if blnMostraMotivoCancelado then
            if blnRelAnalitico then
                s = formata_moeda(r("vl_total_original"))
                x = x & "		<TD valign='middle' align='right' style='width:" & Cstr(w_valor) & "px' class='MDB'><span class='Cnd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & s & "</span></TD>" & chr(13)
            end if
        end if

    '> MOTIVO CANCELAMENTO
        if blnMostraMotivoCancelado then
			s = ""
            if Trim("" & r("cancelado_codigo_motivo")) <> "" then
                s = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CANCELAMENTOPEDIDO_MOTIVO, Trim("" & r("cancelado_codigo_motivo")))
            end if
            if Trim("" & r("cancelado_codigo_sub_motivo")) <> "" then
                s = s & " (" & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CANCELAMENTOPEDIDO_MOTIVO_SUB, Trim("" & r("cancelado_codigo_sub_motivo"))) & ")"
            end if
			if s = "" then s = "&nbsp;"
		    x = x & "		<TD valign='middle' style='width:" & Cstr(w_motivo_cancelamento) & "px' class='MDB'><span class='Cn' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s & "</span></TD>" & chr(13)            
        end if

	'> CAMPOS OPCIONAIS
		if Not blnMostraMotivoCancelado then
			'> VENDEDOR
			if ckb_exibir_vendedor <> "" then
				s = Trim("" & r("vendedor"))
				x = x & "		<TD valign='middle' style='width:" & Cstr(w_vendedor) & "px' class='MDB'><span class='Cn'>" & s & "</span></TD>" & chr(13)
				end if

			'> PARCEIRO
			if ckb_exibir_parceiro <> "" then
				s = Trim("" & r("indicador"))
				x = x & "		<TD valign='middle' style='width:" & Cstr(w_indicador) & "px' class='MDB'><span class='Cn'>" & s & "</span></TD>" & chr(13)
				end if

			'> UF
			if ckb_exibir_uf <> "" then
				s = Trim("" & r("uf_cliente"))
				x = x & "		<TD valign='middle' style='width:" & Cstr(w_uf_cad_cliente) & "px' class='MDB'><span class='Cn'>" & s & "</span></TD>" & chr(13)
				end if
			end if

	'> TOTALIZA��O DE VALORES
		vl_sub_total_faturamento = vl_sub_total_faturamento + r("vl_total_pedido")
		vl_sub_total_faturamento_NF = vl_sub_total_faturamento_NF + r("vl_total_pedido_NF")
		vl_sub_total_pago = vl_sub_total_pago + r("vl_pago_pedido")
		vl_sub_total_a_pagar = vl_sub_total_a_pagar + vl_a_pagar
		if blnPorFornecedor then 
			vl_sub_total_fornecedor = vl_sub_total_fornecedor + r("vl_total_fornecedor")
			vl_sub_total_fornecedor_NF = vl_sub_total_fornecedor_NF + r("vl_total_fornecedor_NF")
			end if
        if blnMostraMotivoCancelado then
            vl_sub_total_pedido_original = vl_sub_total_pedido_original + r("vl_total_original")
        end if
		
		if ckb_exibir_qtde_volumes <> "" then sub_total_qtde_vol = sub_total_qtde_vol + r("total_pedido_qtde_volumes")
		if ckb_exibir_cubagem <> "" then sub_total_cubagem = sub_total_cubagem + r("total_pedido_cubagem")
		if ckb_exibir_peso <> "" then sub_total_peso = sub_total_peso + r("total_pedido_peso")

		vl_total_faturamento = vl_total_faturamento + r("vl_total_pedido")
		vl_total_faturamento_NF = vl_total_faturamento_NF + r("vl_total_pedido_NF")
		vl_total_pago = vl_total_pago + r("vl_pago_pedido")
		vl_total_a_pagar = vl_total_a_pagar + vl_a_pagar
		if blnPorFornecedor then 
			vl_total_fornecedor = vl_total_fornecedor + r("vl_total_fornecedor")
			vl_total_fornecedor_NF = vl_total_fornecedor_NF + r("vl_total_fornecedor_NF")
			end if
        if blnMostraMotivoCancelado then
            vl_total_pedido_original = vl_total_pedido_original + r("vl_total_original")
        end if
		
		if ckb_exibir_qtde_volumes <> "" then total_qtde_vol = total_qtde_vol + r("total_pedido_qtde_volumes")
		if ckb_exibir_cubagem <> "" then total_cubagem = total_cubagem + r("total_pedido_cubagem")
		if ckb_exibir_peso <> "" then total_peso = total_peso + r("total_pedido_peso")

		x = x & "	</TR>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
	
	
	
  ' MOSTRA TOTAL DA �LTIMA LOJA
	if blnRelAnalitico then
		if n_reg <> 0 then 
			s_cor = ""
			if vl_sub_total_a_pagar < 0 then s_cor = "color:red;"
			x = x & "	<TR class='RowTotalizacao' style='background: #FFFFDD'>" & chr(13) & _
					"		<TD style='background:white;'>&nbsp;</td>" & chr(13)
            if s_loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
					x = x & "		<TD COLSPAN='5' align='right' class='MEB' NOWRAP><span class='Cd' style='font-weight:bold;'>TOTAL:</span></td>" & chr(13)
			else    
                    x = x & "		<TD COLSPAN='4' align='right' class='MEB' NOWRAP><span class='Cd' style='font-weight:bold;'>TOTAL:</span></td>" & chr(13)
            end if
			if blnPorFornecedor then
				x = x & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_fornecedor) & "</span></td>" & chr(13) & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_fornecedor_NF) & "</span></td>" & chr(13)
				end if
			
			x = x & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_faturamento) & "</span></td>" & chr(13) & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_faturamento_NF) & "</span></td>" & chr(13) & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_pago) & "</span></td>" & chr(13) & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='" & s_cor & "font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_a_pagar) & "</span></td>" & chr(13) & _
					"		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)

					if ckb_exibir_data_entrega <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_data_previsao_entrega <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_data_previsao_entrega_transp <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_data_recebido_cliente <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_cidade_etg <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_uf_etg <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_qtde_volumes <> "" then
						x = x & _
							"		<TD align='right' style='width:" & Cstr(w_qtde_vol) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(sub_total_qtde_vol) & "</span></td>" & chr(13)
						end if
					if ckb_exibir_cubagem <> "" then
						x = x & _
							"		<TD align='right' style='width:" & Cstr(w_cubagem) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_DECIMAL & chr(34) & ";'>" & formata_numero(sub_total_cubagem, 2) & "</span></td>" & chr(13)
						end if
					if ckb_exibir_peso <> "" then
						x = x & _
							"		<TD align='right' style='width:" & Cstr(w_peso) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_DECIMAL & chr(34) & ";'>" & formata_numero(sub_total_peso, 2) & "</span></td>" & chr(13)
						end if

            if blnMostraMotivoCancelado then
                x = x & _
                    "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13) & _
                    "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
				if ckb_exibir_uf <> "" then
					x = x & _
						"		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
					end if
				x = x & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total_pedido_original) & "</span></td>" & chr(13) & _
					"		<TD class='MDB'><span class='C'>&nbsp;</span></td>" & chr(13)
            else
				if ckb_exibir_vendedor <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
				if ckb_exibir_parceiro <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
				if ckb_exibir_uf <> "" then x = x & "		<TD class='MB'><span class='C'>&nbsp;</span></td>" & chr(13)
				end if

            x = x & _
					"	</TR>" & chr(13)
			
		'>	TOTAL GERAL
			if qtde_lojas > 1 then
				s_cor = ""
				if vl_total_a_pagar < 0 then s_cor = "color:red;"
				if blnPorFornecedor then n_colspan = 11 else n_colspan = 9
				if ckb_exibir_data_entrega <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_data_previsao_entrega <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_cidade_etg <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_uf_etg <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_data_previsao_entrega_transp <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_data_recebido_cliente <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_qtde_volumes <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_peso <> "" then n_colspan = n_colspan + 1
				if ckb_exibir_cubagem <> "" then n_colspan = n_colspan + 1

				if Not blnMostraMotivoCancelado then
					if ckb_exibir_vendedor <> "" then n_colspan = n_colspan + 1
					if ckb_exibir_parceiro <> "" then n_colspan = n_colspan + 1
					if ckb_exibir_uf <> "" then n_colspan = n_colspan + 1
					end if
				x = x & _
					"	<TR>" & chr(13) & _
					"		<TD COLSPAN='" & Cstr(n_colspan) & "' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD COLSPAN='" & Cstr(n_colspan) & "' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"	<TR class='RowTotalizacao' style='background:honeydew'>" & chr(13) & _
					"		<TD style='background:white;'>&nbsp;</td>" & chr(13)
                if s_loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
					x = x & "		<TD class='MTBE' align='right' COLSPAN='5' NOWRAP><span class='Cd' style='font-weight:bold;'>TOTAL GERAL:</span></td>" & chr(13)
                else
                    x = x & "		<TD class='MTBE' align='right' COLSPAN='4' NOWRAP><span class='Cd' style='font-weight:bold;'>TOTAL GERAL:</span></td>" & chr(13)
                end if
					
				if blnPorFornecedor then
					x = x & _
						"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_fornecedor) & "</span></td>" & chr(13) & _
						"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_fornecedor_NF) & "</span></td>" & chr(13)
					end if
				
				x = x & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_faturamento) & "</span></td>" & chr(13) & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_faturamento_NF) & "</span></td>" & chr(13) & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_pago) & "</span></td>" & chr(13) & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><span class='Cd' style='" & s_cor & "font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_a_pagar) & "</span></td>" & chr(13) & _
					"		<TD class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)

				if ckb_exibir_data_entrega <> "" then x = x & "<TD class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)
				if ckb_exibir_data_previsao_entrega <> "" then x = x & "<TD class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)
				if ckb_exibir_data_previsao_entrega_transp <> "" then x = x & "<TD class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)
				if ckb_exibir_data_recebido_cliente <> "" then x = x & "<TD class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)
				if ckb_exibir_cidade_etg <> "" then x = x & "<TD class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)
				if ckb_exibir_uf_etg <> "" then x = x & "<TD class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)
				if ckb_exibir_qtde_volumes <> "" then
					x = x & _
						"		<TD align='right' style='width:" & Cstr(w_qtde_vol) & "px' class='MTB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(total_qtde_vol) & "</span></td>" & chr(13)
					end if
				if ckb_exibir_cubagem <> "" then
					x = x & _
						"		<TD align='right' style='width:" & Cstr(w_cubagem) & "px' class='MTB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_DECIMAL & chr(34) & ";'>" & formata_numero(total_cubagem, 2) & "</span></td>" & chr(13)
					end if
				if ckb_exibir_peso <> "" then
					x = x & _
						"		<TD align='right' style='width:" & Cstr(w_peso) & "px' class='MTB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_DECIMAL & chr(34) & ";'>" & formata_numero(total_peso, 2) & "</span></td>" & chr(13)
					end if

                if blnMostraMotivoCancelado then
                    x = x & _
                    "		<td class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13) & _
                    "		<td class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_uf <> "" then
						x = x & _
							"		<td class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)
						end if
					x = x & _
					"		<TD align='right' style='width:" & Cstr(w_valor) & "px' class='MTB'><span class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_pedido_original) & "</span></td>" & chr(13) & _
					"		<TD class='MTBD'><span class='C'>&nbsp;</span></td>" & chr(13)
                else
					if ckb_exibir_vendedor <> "" then x = x & "<TD class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_parceiro <> "" then x = x & "<TD class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)
					if ckb_exibir_uf <> "" then x = x & "<TD class='MTB'><span class='C'>&nbsp;</span></td>" & chr(13)
				end if

                x = x & _
                    "	</TR>" & chr(13)

				end if
			end if
		end if

  ' MOSTRA AVISO DE QUE N�O H� DADOS!!
	if n_reg_total = 0 then
		if blnPorFornecedor then
			if blnRelAnalitico then
				if blnMostraMotivoCancelado then
                    n_colspan = 15
                else
					n_colspan = 11
                end if 
			else
				if blnMostraMotivoCancelado then
                    n_colspan = 8
                else
				    n_colspan = 5
                    end if
				end if
		else 
			if blnRelAnalitico then
				if blnMostraMotivoCancelado then
                    n_colspan = 13
                else
					n_colspan = 9
                end if
			else
				if blnMostraMotivoCancelado then
                    n_colspan = 8
                else
					n_colspan = 5
                    end if
				end if
			end if
		
		if ckb_exibir_data_entrega <> "" then n_colspan = n_colspan + 1
		if ckb_exibir_data_previsao_entrega <> "" then n_colspan = n_colspan + 1
		if ckb_exibir_cidade_etg <> "" then n_colspan = n_colspan + 1
		if ckb_exibir_uf_etg <> "" then n_colspan = n_colspan + 1
		if ckb_exibir_data_previsao_entrega_transp <> "" then n_colspan = n_colspan + 1
		if ckb_exibir_data_recebido_cliente <> "" then n_colspan = n_colspan + 1
		if ckb_exibir_qtde_volumes <> "" then n_colspan = n_colspan + 1
		if ckb_exibir_peso <> "" then n_colspan = n_colspan + 1
		if ckb_exibir_cubagem <> "" then n_colspan = n_colspan + 1

		if Not blnMostraMotivoCancelado then
			if ckb_exibir_vendedor <> "" then n_colspan = n_colspan + 1
			if ckb_exibir_parceiro <> "" then n_colspan = n_colspan + 1
			if ckb_exibir_uf <> "" then n_colspan = n_colspan + 1
		else
			if ckb_exibir_uf <> "" then n_colspan = n_colspan + 1
			end if

		x = cab_table & cab
		x = x & "	<TR>" & chr(13) & _
				"		<TD style='background:white;'>&nbsp;</td>" & chr(13) & _
				"		<TD class='MDBE ALERTA' align='center' colspan='" & Cstr(n_colspan) & "'><span class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</span></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DA �LTIMA LOJA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	
end sub

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
	<title>CENTRAL</title>
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__RASTREIO_VIA_WEBAPI_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
    var historyBackCount = 1;
	var windowScrollTopAnterior;

	var urlBaseSsw = '<%=URL_SSW_BASE%>';
	var urlWebApiRastreio;
	var serverVariableUrl;
	serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
	serverVariableUrl = serverVariableUrl.toUpperCase();
	serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));
	urlWebApiRastreio = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/GetData/PageContentViaHttpGet';

	$(document).ready(function () {
		$(".RowTotalizacao td:last-child").addClass("MD");
		$("#divPedidoConsultaView").hide();
		$("#divRastreioConsultaView").hide();
		$('#divInternoRastreioConsultaView').addClass('divFixo');
		sizeDivPedidoConsultaView();
		sizeDivRastreioConsultaView();

		$('#divInternoPedidoConsultaView').addClass('divFixo');

		$(document).keyup(function (e) {
			if (e.keyCode == 27) {
				fechaDivRastreioConsultaView();
				fechaDivPedidoConsultaView();
			}
		});

		$("#divPedidoConsultaView").click(function () {
			fechaDivPedidoConsultaView();
		});

		$("#imgFechaDivPedidoConsultaView").click(function () {
			fechaDivPedidoConsultaView();
		});

		$("#divRastreioConsultaView").click(function () {
			fechaDivRastreioConsultaView();
		});

		$("#imgFechaDivRastreioConsultaView").click(function () {
			fechaDivRastreioConsultaView();
		});
	});

	//Every resize of window
	$(window).resize(function () {
		sizeDivRastreioConsultaView();
		sizeDivPedidoConsultaView();
	});

	function sizeDivPedidoConsultaView() {
		var newHeight = $(document).height() + "px";
		$("#divPedidoConsultaView").css("height", newHeight);
	}

    function sizeDivRastreioConsultaView() {
        var newHeight = $(document).height() + "px";
        $("#divRastreioConsultaView").css("height", newHeight);
    }

	function fechaDivPedidoConsultaView() {
		$(window).scrollTop(windowScrollTopAnterior);
		$("#divPedidoConsultaView").fadeOut();
		$("#iframePedidoConsultaView").attr("src", "");
	}

    function fechaDivRastreioConsultaView() {
        $("#divRastreioConsultaView").fadeOut();
        $("#iframeRastreioConsultaView").attr("src", "");
    }

    function fRastreioConsultaView(url) {
        historyBackCount++;
        sizeDivRastreioConsultaView();
        $("#iframeRastreioConsultaView").attr("src", url);
        $("#divRastreioConsultaView").fadeIn();
    }

	function fRastreioConsultaViaWebApiView(url) {
		executaRastreioConsultaViaWebApiView(url, urlBaseSsw, urlWebApiRastreio, "<%=usuario%>", "<%=s_sessionToken%>", "#iframeRastreioConsultaView", "#divRastreioConsultaView");
	}

	function fPEDConsultaView(id_pedido, usuario) {
		historyBackCount++;
		windowScrollTopAnterior = $(window).scrollTop();
		sizeDivPedidoConsultaView();
		$("#iframePedidoConsultaView").attr("src", "PedidoConsultaView.asp?pedido_selecionado=" + id_pedido + "&pedido_selecionado_inicial=" + id_pedido + "&usuario=" + usuario);
		$("#divPedidoConsultaView").fadeIn();
	}
</script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function realca_cor_mouse_over(c) {
	c.style.backgroundColor = 'palegreen';
}

function realca_cor_mouse_out(c) {
	c.style.backgroundColor = '';
}

function fRELConcluir( id_pedido ){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=id_pedido;
	fREL.action = "pedido.asp"
	fREL.submit(); 
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
<link href="<%=URL_FILE__ESCREEN_CSS%>" rel="stylesheet" type="text/css" media="screen">

<style type="text/css">
html
{
	overflow-y: scroll;
	height:100%;
	margin:0px;
}
body
{
	height:100%;
	margin:0px;
}
#divPedidoConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsultaView
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsultaView.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivPedidoConsultaView
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframePedidoConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
#divRastreioConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoRastreioConsultaView
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#fff;
	opacity: 1;
}
#divInternoRastreioConsultaView.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivRastreioConsultaView
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframeRastreioConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
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
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>



<% else %>
<!-- ***************************************************** -->
<!-- **********  P�GINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Conclu�do';" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<input type="hidden" name="url_origem" id="url_origem" value="RelPedidosMCrit.asp" />

<!--  I D E N T I F I C A � � O   D A   T E L A  -->
<table width="849" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relat�rio Multicrit�rio de Pedidos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='849' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
	s = ""
	s_aux = Lcase(x_status_entrega(ckb_st_entrega_esperar))
	if s_aux<>"" then
	'	DEVIDO AO WORD WRAP: S� FAZ WORD WRAP QUANDO ENCONTRA CHR(32), OU SEJA, MANT�M AGRUPADO TEXTO COM &nbsp;
		if s <> "" then s = s & ",&nbsp; "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_status_entrega(ckb_st_entrega_split))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_separar_sem_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		s_aux = s_aux & " (sem data de coleta)"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_separar_com_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		if c_dt_coleta_a_separar_inicio <> "" then s_aux_dti = c_dt_coleta_a_separar_inicio else s_aux_dti = "N.I."
		if c_dt_coleta_a_separar_termino <> "" then s_aux_dtf = c_dt_coleta_a_separar_termino else s_aux_dtf = "N.I."
		s_aux = s_aux & " (com data de coleta: " & s_aux_dti & " a " & s_aux_dtf & ")"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_status_entrega(ckb_st_entrega_a_entregar_sem_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		s_aux = s_aux & " (sem data de coleta)"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_a_entregar_com_marc))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		if c_dt_coleta_st_a_entregar_inicio <> "" then s_aux_dti = c_dt_coleta_st_a_entregar_inicio else s_aux_dti = "N.I."
		if c_dt_coleta_st_a_entregar_termino <> "" then s_aux_dtf = c_dt_coleta_st_a_entregar_termino else s_aux_dtf = "N.I."
		s_aux = s_aux & " (com data de coleta: " & s_aux_dti & " a " & s_aux_dtf & ")"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_entregue))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		s = s & s_aux
		s_aux = c_dt_entregue_inicio
		if s_aux = "" then s_aux = "N.I."
		s_aux = " (" & s_aux & " a "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		s_aux = c_dt_entregue_termino
		if s_aux = "" then s_aux = "N.I."
		s_aux = s_aux & ")"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_entrega(ckb_st_entrega_cancelado))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp; "
		s = s & s_aux
		s_aux = c_dt_cancelado_inicio
		if s_aux = "" then s_aux = "N.I."
		s_aux = " (" & s_aux & " a "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		s_aux = c_dt_cancelado_termino
		if s_aux = "" then s_aux = "N.I."
		s_aux = s_aux & ")"
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	if ckb_st_entrega_exceto_cancelados <> "" then
		s_aux = "exceto cancelados"
		if s <> "" then s = s & ",&nbsp; "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	if ckb_st_entrega_exceto_entregues <> "" then
		s_aux = "exceto entregues"
		if s <> "" then s = s & ",&nbsp; "
		s_aux = replace(s_aux, " ", "&nbsp;")
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Status de Entrega:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = ""
    if ckb_pedido_nao_recebido_pelo_cliente <> "" then
		s_aux = "n�o recebidos"
		if s <> "" then s = s & ",&nbsp; "
		s = s & s_aux
		end if

	if ckb_pedido_recebido_pelo_cliente <> "" then
		s_aux = "recebidos"
		if s <> "" then s = s & ",&nbsp; "
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Pedidos Recebidos pelo Cliente:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = ""
	s_aux = Lcase(x_status_pagto(ckb_st_pagto_pago))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_status_pagto(ckb_st_pagto_nao_pago))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_status_pagto(ckb_st_pagto_pago_parcial))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Status de Pagamento:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = ""
	if ckb_pagto_antecipado_status_nao <> "" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & "N�o"
		end if

	if ckb_pagto_antecipado_status_sim <> "" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & "Sim"
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Pagamento Antecipado:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = ""
	if ckb_pagto_antecipado_quitado_status_pendente <> "" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & pagto_antecipado_quitado_descricao(COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO, COD_PAGTO_ANTECIPADO_QUITADO_STATUS_PENDENTE)
		end if

	if ckb_pagto_antecipado_quitado_status_quitado <> "" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & pagto_antecipado_quitado_descricao(COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO, COD_PAGTO_ANTECIPADO_QUITADO_STATUS_QUITADO)
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Status Pagamento Antecipado:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = ""
	
	if ckb_analise_credito_st_inicial <> "" then
		s_aux = "status inicial"
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_vendas))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_endereco))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_cartao))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = Lcase(x_analise_credito(ckb_analise_credito_pendente_pagto_antecipado_boleto))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok_aguardando_deposito))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok_deposito_aguardando_desbloqueio))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	s_aux = Lcase(x_analise_credito(ckb_analise_credito_ok_aguardando_pagto_boleto_av))
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>An�lise de Cr�dito:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = ""
	s_aux = ""
	if CStr(ckb_entrega_imediata_sim) = CStr(COD_ETG_IMEDIATA_SIM) then s_aux = "sim"
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = ""
	if CStr(ckb_entrega_imediata_nao) = CStr(COD_ETG_IMEDIATA_NAO) then s_aux = "n�o"
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		s = s & " (previs�o de entrega: "
		s_aux = c_dt_previsao_entrega_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s = s & " a "
		s_aux = c_dt_previsao_entrega_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s = s & ")"
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Entrega Imediata:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if
		
	'Geral: campo Obs II
	s = ""
	s_aux = ""
	if ckb_obs2_preenchido <> "" then s_aux = "N� NF preenchido"
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = ""
	if ckb_obs2_nao_preenchido <> "" then s_aux = "N� NF n�o preenchido"
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Geral:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	'Indicador preenchido
	s = ""
	s_aux = ""
	if ckb_indicador_preenchido <> "" then s_aux = "Indicador preenchido"
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if
	
	s_aux = ""
	if ckb_indicador_nao_preenchido <> "" then s_aux = "Indicador n�o preenchido"
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & s_aux
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Indicador:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if (c_dt_cadastro_inicio <> "") Or (c_dt_cadastro_termino <> "") then
		s = ""
		s_aux = c_dt_cadastro_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " e "
		s_aux = c_dt_cadastro_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Pedidos colocados entre:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if ckb_entrega_marcada_para <> "" then
		s = ""
		s_aux = c_dt_entrega_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " a "
		s_aux = c_dt_entrega_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Data de coleta:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if ckb_periodo_emissao_NF_venda <> "" then
		s = ""
		s_aux = c_dt_NF_venda_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " a "
		s_aux = c_dt_NF_venda_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Emiss�o NF Venda:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if ckb_periodo_emissao_NF_remessa <> "" then
		s = ""
		s_aux = c_dt_NF_remessa_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " a "
		s_aux = c_dt_NF_remessa_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Emiss�o NF Remessa:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if ckb_produto <> "" then 
		s_aux = c_fabricante
		if s_aux = "" then s_aux = "todos"
		s = "fabricante: " & s_aux
		s_aux = c_produto
		if s_aux = "" then s_aux = "todos"
		s = s & ",&nbsp;&nbsp;produto: " & s_aux
		if ckb_somente_pedidos_produto_alocado <> "" then s = s & "&nbsp;&nbsp;(somente pedidos que possuam o produto alocado)"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Somente pedidos que incluam:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

    s = c_grupo
	if s = "" then 
		s = "todos"
	else
        s_filtro = s_filtro & _
			"	<tr>" & chr(13) & _
			"		<td align='right' valign='top' nowrap>" & _
			"<span class='N'>Grupo(s) de Produtos:&nbsp;</span></td>" & chr(13) & _
			"		<td align='left' valign='top'>" & _
			"<span class='N'>" & s & "</span></td>" & chr(13) & _
			"	</tr>" & chr(13)
	end if

    s = c_grupo_pedido_origem
	if s = "" then 
		s = "todos"
	else
		s = obtem_descricao_tabela_t_codigo_descricao("PedidoECommerce_Origem_Grupo", c_grupo_pedido_origem)
        s_filtro = s_filtro & _
			"	<tr>" & chr(13) & _
			"		<td align='right' valign='top' nowrap>" & _
			"<span class='N'>Origem Pedido (Grupo):&nbsp;</span></td>" & chr(13) & _
			"		<td align='left' valign='top'>" & _
			"<span class='N'>" & s & "</span></td>" & chr(13) & _
			"	</tr>" & chr(13)
	end if

    s = c_pedido_origem
	if s = "" then 
		s = "todos"
	else
		s = obtem_descricao_tabela_t_codigo_descricao("PedidoECommerce_Origem", c_pedido_origem)
        s_filtro = s_filtro & _
			"	<tr>" & chr(13) & _
			"		<td align='right' valign='top' nowrap>" & _
			"<span class='N'>Origem do Pedido:&nbsp;</span></td>" & chr(13) & _
			"		<td align='left' valign='top'>" & _
			"<span class='N'>" & s & "</span></td>" & chr(13) & _
			"	</tr>" & chr(13)
    end if
    
    s = c_empresa
	if s = "" then 
		s = "todas"
	else
		s =  obtem_apelido_empresa_NFe_emitente(c_empresa)
		end if
	s_filtro = s_filtro & _
			"	<tr>" & chr(13) & _
			"		<td align='right' valign='top' nowrap>" & _
			"<span class='N'>Empresa:&nbsp;</span></td>" & chr(13) & _
			"		<td align='left' valign='top'>" & _
			"<span class='N'>" & s & "</span></td>" & chr(13) & _
			"	</tr>" & chr(13)

	select case rb_loja
		case "TODAS": s = "todas"
		case "UMA"
			s = ""
			for i=LBound(vLoja) to UBound(vLoja)
				if s <> "" then s = s & ", "
				s = s & Trim("" & vLoja(i))
				next
		case "FAIXA"
			s = ""
			s_aux = c_loja_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " a "
			s_aux = c_loja_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
		case else: s = ""
		end select
	
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><span class='N'>Lojas:&nbsp;</span></td>" & chr(13) & _
				"		<td valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

	if op_forma_pagto <> "" then
		s = x_opcao_forma_pagamento(op_forma_pagto)
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Forma Pagto:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if c_forma_pagto_qtde_parc <> "" then
		s = c_forma_pagto_qtde_parc
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>N� Parcelas:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if c_cliente_cnpj_cpf <> "" then
		s = cnpj_cpf_formata(c_cliente_cnpj_cpf)
		s_aux = x_cliente_por_cnpj_cpf(c_cliente_cnpj_cpf, cadastrado)
		if Not cadastrado then s_aux = "N�o Cadastrado"
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Cliente:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

    if c_cliente_uf <> "" then
		s = c_cliente_uf
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>UF do Cliente:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if ckb_visanet <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Cart�o de Cr�dito:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>somente pedidos pagos usando cart�o de cr�dito</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if
	
	if c_transportadora <> "" then
		s = c_transportadora
		s_aux = iniciais_em_maiusculas(x_transportadora(c_transportadora))
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Transportadora:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if c_transportadora_multiplo <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Transportadora(s):&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & c_transportadora_multiplo & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if c_vendedor <> "" then
		s = c_vendedor
		s_aux = x_usuario(c_vendedor)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Vendedor:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	if c_indicador <> "" then
		s = c_indicador
		s_aux = x_orcamentista_e_indicador(c_indicador)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><span class='N'>Indicador:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if
	
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><span class='N'>Emiss�o:&nbsp;</span></td>" & chr(13) & _
				"		<td valign='top' width='99%'><span class='N'>" & formata_data_hora(Now) & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELAT�RIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="849" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="849" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="RelPedidosMCrit.asp<%= "?" & "url_back=X&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a p�gina anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

</center>

<div id="divRastreioConsultaView"><center><div id="divInternoRastreioConsultaView"><img id="imgFechaDivRastreioConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframeRastreioConsultaView"></iframe></div></center></div>
<div id="divPedidoConsultaView"><center><div id="divInternoPedidoConsultaView"><img id="imgFechaDivPedidoConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframePedidoConsultaView"></iframe></div></center></div>

</body>

<% end if %>

</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
