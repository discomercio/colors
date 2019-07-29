<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Braspag.asp"    -->

<%
'     =============================================================
'	  AjaxBraspagCSOpcoesParcelamento.asp
'     =============================================================
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
	
	dim bandeira, pedido, valor_pagamento_cartao, vl_pagamento_cartao, msg_erro
	bandeira = Trim(Request("bandeira"))
	pedido = Trim(Request("pedido"))
	valor_pagamento_cartao = Trim(Request("valor_pagamento"))
	vl_pagamento_cartao = converte_numero(valor_pagamento_cartao)

	if pedido = "" then
		Response.Write "Número do pedido não foi informado!!"
		Response.End
		end if
	
	if bandeira = "" then
		Response.Write "Bandeira do cartão não foi informada!!"
		Response.End
		end if
	
	if valor_pagamento_cartao = "" then
		Response.Write "Valor do pagamento não foi informado!!"
		Response.End
		end if
	
	if vl_pagamento_cartao <= 0 then
		Response.Write "Valor do pagamento informado é inválido!!"
		Response.End
		end if

	dim rs, cn
	
'	CONECTA COM O BANCO DE DADOS
	If Not bdd_conecta(cn) then
		msg_erro = erro_descricao(ERR_CONEXAO)
		Response.Write msg_erro
		Response.End
		end if
	
	If Not cria_recordset_otimista(rs, msg_erro) then
		msg_erro = erro_descricao(ERR_FALHA_OPERACAO_CRIAR_ADO)
		Response.Write msg_erro
		Response.End
		end if

'	OBTÉM A QUANTIDADE DE PARCELAS DEFINIDA NO PEDIDO
	dim r_pedido
	call le_pedido(pedido, r_pedido, msg_erro)

	dim intNumParcelasFormaPagto
	intNumParcelasFormaPagto = 0
	if Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_A_VISTA) then
		if Cstr(r_pedido.av_forma_pagto) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = 1
	elseif Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_PARCELA_UNICA) then
		if Cstr(r_pedido.pu_forma_pagto) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = 1
	elseif Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_PARCELADO_CARTAO) then
		intNumParcelasFormaPagto = r_pedido.pc_qtde_parcelas
	elseif Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) then
		'NOP
	elseif Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) then
	'	ENTRADA + PRESTAÇÕES
		if Cstr(r_pedido.pce_forma_pagto_entrada) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = intNumParcelasFormaPagto + 1
		if Cstr(r_pedido.pce_forma_pagto_prestacao) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = intNumParcelasFormaPagto + r_pedido.pce_prestacao_qtde
	elseif Cstr(r_pedido.tipo_parcelamento) = Cstr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) then
	'	1ª PRESTAÇÃO + DEMAIS PRESTAÇÕES
		if Cstr(r_pedido.pse_forma_pagto_prim_prest) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = intNumParcelasFormaPagto + 1
		if Cstr(r_pedido.pse_forma_pagto_demais_prest) = ID_FORMA_PAGTO_CARTAO then intNumParcelasFormaPagto = intNumParcelasFormaPagto + r_pedido.pse_demais_prest_qtde
	else
		intNumParcelasFormaPagto = 1
		end if
	
	if intNumParcelasFormaPagto = 0 then intNumParcelasFormaPagto = 1

'	CONSULTA PARÂMETROS ARMAZENADOS NO BD
	dim s, qtde_parc_loja, vl_min_loja, qtde_parc_cartao, vl_min_cartao
	qtde_parc_loja = 0
	vl_min_loja = 0
	qtde_parc_cartao = 0
	vl_min_cartao = 0
	
	s = "SELECT * FROM t_PRAZO_PAGTO_VISANET WHERE tipo = '" & BraspagObtemIdRegistroBdPrazoPagtoLoja(bandeira) & "'"
	rs.Open s, cn
	if Not rs.Eof then
		qtde_parc_loja = rs("qtde_parcelas")
		vl_min_loja = rs("vl_min_parcela")
		end if

	s = "SELECT * FROM t_PRAZO_PAGTO_VISANET WHERE tipo = '" & BraspagObtemIdRegistroBdPrazoPagtoEmissor(bandeira) & "'"
	if rs.State <> 0 then rs.Close
	rs.Open s, cn
	if Not rs.Eof then
		qtde_parc_cartao = rs("qtde_parcelas")
		vl_min_cartao = rs("vl_min_parcela")
		end if
	
	if qtde_parc_loja > intNumParcelasFormaPagto then qtde_parc_loja = intNumParcelasFormaPagto
	if qtde_parc_cartao > intNumParcelasFormaPagto then qtde_parc_cartao = intNumParcelasFormaPagto

	if rs.State <> 0 then rs.Close
	set rs=nothing

	cn.Close
	set cn = nothing
	
'	DETERMINA AS OPÇÕES QUE SERÃO DISPONIBILIZADAS
	dim i, n, qtde_parcelas_loja_ok
	dim credito_habilitado, debito_habilitado, texto_obs_juros

	if bandeira = BRASPAG_BANDEIRA__VISA then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = BRASPAG_BANDEIRA__MASTERCARD then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = BRASPAG_BANDEIRA__AMEX then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = BRASPAG_BANDEIRA__ELO then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = BRASPAG_BANDEIRA__DINERS then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = BRASPAG_BANDEIRA__DISCOVER then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = BRASPAG_BANDEIRA__AURA then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = BRASPAG_BANDEIRA__JCB then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = BRASPAG_BANDEIRA__CELULAR then
		credito_habilitado = True
		debito_habilitado = True
	else
		credito_habilitado = False
		debito_habilitado = False
		end if
	
	qtde_parcelas_loja_ok = 0

	dim strResp
	strResp = ""

	if debito_habilitado then
		if strResp <> "" then strResp = strResp & ","
		strResp = strResp & _
				"{" & _
					"""value"" : """ & "A" & """," & _
					"""description"" : """ & SIMBOLO_MONETARIO & " " & formata_moeda(vl_pagamento_cartao) & " à Vista (no Débito)" & """," & _
					"""obs_juros"" : """ & "" & """" & _
				"}"
		end if

	if credito_habilitado then
		if strResp <> "" then strResp = strResp & ","
		strResp = strResp & _
				"{" & _
					"""value"" : """ & "0" & """," & _
					"""description"" : """ & SIMBOLO_MONETARIO & " " & formata_moeda(vl_pagamento_cartao) & " à Vista (no Crédito)" & """," & _
					"""obs_juros"" : """ & "" & """" & _
				"}"
		end if

	if qtde_parc_loja > 1 then
		for i = 2 to qtde_parc_loja
			if (vl_pagamento_cartao / i) < vl_min_loja then exit for
			qtde_parcelas_loja_ok = i
			if strResp <> "" then strResp = strResp & ","
			strResp = strResp & _
					"{" & _
						"""value"" : """ & "PL|" & CStr(i) & """," & _
						"""description"" : """ & Cstr(i) & "x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_pagamento_cartao/i) & " iguais" & """," & _
						"""obs_juros"" : """ & "" & """" & _
					"}"
			next
		end if
	
	if qtde_parc_cartao > 1 then
		n = 0
		texto_obs_juros = ""
		for i = 2 to qtde_parc_cartao
			if i > qtde_parcelas_loja_ok then
				if (vl_pagamento_cartao / i) < vl_min_cartao then exit for
				n = n + 1
				if n = 1 then 
					if bandeira = BRASPAG_BANDEIRA__VISA Then
						texto_obs_juros = "Verificar a taxa de juros junto ao emissor do cartão Visa"
					elseif bandeira = BRASPAG_BANDEIRA__MASTERCARD Then
						texto_obs_juros = "Verificar a taxa de juros junto ao emissor do cartão Mastercard"
					elseif bandeira = BRASPAG_BANDEIRA__AMEX Then
						texto_obs_juros = "Verificar a taxa de juros junto ao emissor do cartão Amex"
					elseif bandeira = BRASPAG_BANDEIRA__ELO Then
						texto_obs_juros = "Verificar a taxa de juros junto ao emissor do cartão ELO"
					elseif bandeira = BRASPAG_BANDEIRA__DINERS Then
						texto_obs_juros = "Verificar a taxa de juros junto ao emissor do cartão Diners"
					elseif bandeira = BRASPAG_BANDEIRA__DISCOVER Then
						texto_obs_juros = "Verificar a taxa de juros junto ao emissor do cartão Discover"
					elseif bandeira = BRASPAG_BANDEIRA__AURA Then
						texto_obs_juros = "Verificar a taxa de juros junto ao emissor do cartão Aura"
					elseif bandeira = BRASPAG_BANDEIRA__JCB Then
						texto_obs_juros = "Verificar a taxa de juros junto ao emissor do cartão JCB"
						end if
					end if
				
				if strResp <> "" then strResp = strResp & ","
				strResp = strResp & _
						"{" & _
							"""value"" : """ & "PC|" & CStr(i) & """," & _
							"""description"" : """ & Cstr(i) & "x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_pagamento_cartao/i) & " mais juros" & """," & _
							"""obs_juros"" : """ & texto_obs_juros & """" & _
						"}"

				end if
			next
		end if

'	ENVIA RESPOSTA
	strResp = "[" & strResp & "]"
	Response.Write strResp
	Response.End
%>
