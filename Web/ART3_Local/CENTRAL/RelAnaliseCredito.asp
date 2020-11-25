<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  R E L A N A L I S E C R E D I T O . A S P
'     ========================================================
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
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, rs2, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_ANALISE_CREDITO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim s, lista_pedidos, v_pedido, v_aux, i, tem_filtro_pedido, origem
	origem = ucase(Trim(request("origem")))
	lista_pedidos = ucase(Trim(request("c_pedidos")))
	lista_pedidos=substitui_caracteres(lista_pedidos,chr(10),"")
	v_aux = split(lista_pedidos,chr(13),-1)
	for i=Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			s = normaliza_num_pedido(v_aux(i))
			if s <> "" then v_aux(i) = s
			end if
		next
	
	tem_filtro_pedido=False
	redim v_pedido(0)
	v_pedido(Ubound(v_pedido))=""
	for i = Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			if Trim(v_pedido(Ubound(v_pedido)))<>"" then
				redim preserve v_pedido(Ubound(v_pedido)+1)
				end if
			tem_filtro_pedido = True
			v_pedido(Ubound(v_pedido)) = Trim(v_aux(i))
			end if
		next
	
	lista_pedidos = join(v_pedido,chr(13))

	dim v
	dim c_valor_inferior, c_valor_superior, c_lista_loja, s_lista_loja, v_loja
	dim c_vendedor, c_indicador
	
	if origem="A" then
		c_lista_loja = Trim(Session("c_lista_loja"))
		c_valor_inferior = Trim(Session("c_valor_inferior"))
		c_valor_superior = Trim(Session("c_valor_superior"))
		c_vendedor = Trim(Session("c_vendedor"))
		c_indicador = Trim(Session("c_indicador"))
		Session("c_lista_loja") = ""
		Session("c_valor_inferior") = ""
		Session("c_valor_superior") = ""
		Session("c_vendedor") = ""
		Session("c_indicador") = ""
	else
		c_lista_loja = Trim(Request.Form("c_lista_loja"))
		c_valor_inferior = Trim(Request.Form("c_valor_inferior"))
		c_valor_superior = Trim(Request.Form("c_valor_superior"))
		c_vendedor = Trim(Request.Form("c_vendedor"))
		c_indicador = Trim(Request.Form("c_indicador"))
		end if
	
	s_lista_loja = substitui_caracteres(c_lista_loja,chr(10),"")
	v_loja = split(s_lista_loja,chr(13),-1)







' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const MAX_PEDIDOS = 5
const NUM_LINHAS_OBS = 4
const NUM_LINHAS_DESCR_FORMA_PAGTO = 4
const intEspacamentoOpcoesAnaliseCredito = 28
const MAX_PEDIDOS_POR_LINHA_ANALISE_ENDERECO = 8
dim r, tP, tPI, tPI2
dim msg_erro
dim strSocMaj, strRefBanc, strRefCom, strRefProf, strRefProfCnpj
dim vRefBancaria, vRefComercial, vRefProfissional
dim intCounter, intIndice, intQtdePedido, intQtdeLinhasPedido
dim s, s_aux, s_sql, cab_table, cab, n_reg, n_pedido, n_pedidos_anteriores
dim s_where, s_from, s_where_pedido, s_where_loja, s_where_aux
dim s_nome, s_cnpj_cpf, s_endereco, s_endereco_entrega, s_tel_res, s_tel_com, s_rg, s_email, s_email_xml
dim s_descricao_forma_pagto_a, s_forma_pagto_a, s_forma_pagto_ped_ant, s_obs1_a
dim s_indicador, s_desempenho_nota
dim strInfoAnEnd
dim x, pedido_a
dim w_fabricante, w_produto, w_descricao, w_qtde
dim w_preco_lista, w_desconto, w_vl_unitario, w_vl_total
dim st_pagto, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, s_cor, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, vl_saldo_a_pagar
dim vl_item, vl_total_item
dim intPedido
dim mudou_pedido
dim vl_filtro_valor_inferior, vl_filtro_valor_superior
dim s_desc, perc_desc, vl_unitario, vl_lista
dim blnPulaRegistro
dim intResto, intQtdeTotalPedidosAnEndereco
dim blnAnEnderecoUsaEndParceiro
dim iCountPedFamilia

	if Not cria_recordset_otimista(tP, msg_erro) then
		Response.Write "Falha ao tentar criar recordset: " & msg_erro
		exit sub
		end if

	if Not cria_recordset_otimista(tPI, msg_erro) then
		Response.Write "Falha ao tentar criar recordset: " & msg_erro
		exit sub
		end if

	if Not cria_recordset_otimista(tPI2, msg_erro) then
		Response.Write "Falha ao tentar criar recordset: " & msg_erro
		exit sub
		end if

	vl_filtro_valor_inferior = converte_numero(c_valor_inferior)
	vl_filtro_valor_superior = converte_numero(c_valor_superior)

'	MONTA CLÁUSULA WHERE
'	DESCONSIDERA OS PEDIDOS-FILHOTE, POIS A ANÁLISE DE CRÉDITO É ANOTADA NO PEDIDO-BASE
	s_where = " (t_PEDIDO.tamanho_num_pedido = " & Cstr(TAM_MIN_ID_PEDIDO) & ")"

'	CRITÉRIO: ANÁLISE DE CRÉDITO
	s = " t_PEDIDO__BASE.analise_credito = " & COD_AN_CREDITO_ST_INICIAL
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	DESCONSIDERA PEDIDOS CANCELADOS
'	NO CASO DE FAMÍLIA DE PEDIDOS DECORRENTE DE AUTO-SPLIT, É PERMITIDO QUE ALGUNS DOS PEDIDOS TENHAM SIDO CANCELADOS, MAS NÃO TODOS, ENTÃO VERIFICA SE HÁ ALGUM PEDIDO NA FAMÍLIA QUE NÃO ESTEJA CANCELADO
	s = " (" & _
			" EXISTS " & _
			"(SELECT TOP 1 tP.pedido FROM t_PEDIDO tP WHERE (tP.pedido LIKE t_PEDIDO__BASE.pedido + '" & BD_CURINGA_TODOS & "') AND (tP.st_entrega <> '" & ST_ENTREGA_CANCELADO & "'))" & _
		")"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & s
	
'	CRITÉRIO: FILTRA PEDIDO(S)
	s_where_pedido = ""
	if tem_filtro_pedido then
		for intPedido=Lbound(v_pedido) to Ubound(v_pedido)
			if Trim(v_pedido(intPedido)) <> "" then
				if s_where_pedido <> "" then s_where_pedido = s_where_pedido & " OR"
				s_where_pedido = s_where_pedido & " (t_PEDIDO.pedido = '" & retorna_num_pedido_base(Trim(v_pedido(intPedido))) & "')"
				end if
			next
		end if

	if s_where_pedido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s_where_pedido & ")"
		end if
		
'	CRITÉRIO: LOJAS
	s_where_loja = ""
	for i=Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
				s_where_loja = s_where_loja & " (t_PEDIDO__BASE.numero_loja = " & v_loja(i) & ")"
			else
				s = ""
				if v(Lbound(v))<>"" then 
					if s <> "" then s = s & " AND"
					s = s & " (t_PEDIDO__BASE.numero_loja >= " & v(Lbound(v)) & ")"
					end if
				if v(Ubound(v))<>"" then
					if s <> "" then s = s & " AND"
					s = s & " (t_PEDIDO__BASE.numero_loja <= " & v(Ubound(v)) & ")"
					end if
				if s <> "" then 
					if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
					s_where_loja = s_where_loja & " (" & s & ")"
					end if
				end if
			end if
		next

	if s_where_loja <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s_where_loja & ")"
		end if

'	CRITÉRIO: VENDEDOR
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.vendedor = '" & c_vendedor & "')"
		end if

'	CRITÉRIO: INDICADOR
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if


	if s_where <> "" then s_where = " WHERE" & s_where
	
'	MONTA CLÁUSULA FROM
	s_from = " FROM t_PEDIDO" & _
			 " INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			 " INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			 " INNER JOIN t_USUARIO ON (t_PEDIDO__BASE.vendedor=t_USUARIO.usuario)" & _
			 " LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)"

	s_sql = "SELECT t_PEDIDO__BASE.loja, t_PEDIDO__BASE.numero_loja," & _
			" t_PEDIDO.data, t_PEDIDO.hora, t_PEDIDO.pedido_base, t_PEDIDO.pedido, t_PEDIDO.st_entrega, t_PEDIDO.entregue_data, t_PEDIDO.cancelado_data," & _
			" t_PEDIDO__BASE.id_cliente, t_PEDIDO.obs_1, t_PEDIDO__BASE.analise_endereco_tratar_status, t_PEDIDO__BASE.analise_endereco_tratado_status," & _
			" t_PEDIDO__BASE.vendedor, t_USUARIO.nome AS nome_vendedor," & _
			" t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.forma_pagto, t_PEDIDO__BASE.tipo_parcelamento, t_PEDIDO__BASE.av_forma_pagto," & _
			" t_PEDIDO__BASE.pu_forma_pagto, t_PEDIDO__BASE.pu_valor, t_PEDIDO__BASE.pu_vencto_apos," & _
			" t_PEDIDO__BASE.pc_qtde_parcelas, t_PEDIDO__BASE.pc_valor_parcela, t_PEDIDO__BASE.pc_maquineta_qtde_parcelas, t_PEDIDO__BASE.pc_maquineta_valor_parcela," & _
			" t_PEDIDO__BASE.pce_forma_pagto_entrada, t_PEDIDO__BASE.pce_forma_pagto_prestacao, t_PEDIDO__BASE.pce_entrada_valor, t_PEDIDO__BASE.pce_prestacao_qtde," & _
			" t_PEDIDO__BASE.pce_prestacao_valor, t_PEDIDO__BASE.pce_prestacao_periodo, t_PEDIDO__BASE.pse_forma_pagto_prim_prest," & _
			" t_PEDIDO__BASE.pse_forma_pagto_demais_prest, t_PEDIDO__BASE.pse_prim_prest_valor, t_PEDIDO__BASE.pse_prim_prest_apos," & _
			" t_PEDIDO__BASE.pse_demais_prest_qtde, t_PEDIDO__BASE.pse_demais_prest_valor, t_PEDIDO__BASE.pse_demais_prest_periodo," & _
			" t_PEDIDO.st_end_entrega, t_PEDIDO.EndEtg_endereco, t_PEDIDO.EndEtg_endereco_numero, t_PEDIDO.EndEtg_endereco_complemento," & _
			" t_PEDIDO.EndEtg_bairro, t_PEDIDO.EndEtg_cidade, t_PEDIDO.EndEtg_uf, t_PEDIDO.EndEtg_cep," & _
			" t_CLIENTE.nome, t_CLIENTE.nome_iniciais_em_maiusculas, t_CLIENTE.tipo, t_CLIENTE.cnpj_cpf, t_CLIENTE.rg, t_CLIENTE.ie," & _
			" t_CLIENTE.ddd_res, t_CLIENTE.tel_res, t_CLIENTE.ddd_com, t_CLIENTE.tel_com, t_CLIENTE.ramal_com," & _
			" t_CLIENTE.endereco, t_CLIENTE.endereco_numero, t_CLIENTE.endereco_complemento, t_CLIENTE.bairro, t_CLIENTE.cidade, t_CLIENTE.uf, t_CLIENTE.cep, t_CLIENTE.email," & _
			" t_CLIENTE.SocMaj_Nome, t_CLIENTE.SocMaj_CPF, t_CLIENTE.SocMaj_banco, t_CLIENTE.SocMaj_agencia," & _
			" t_CLIENTE.SocMaj_conta, t_CLIENTE.SocMaj_ddd, t_CLIENTE.SocMaj_telefone, t_CLIENTE.SocMaj_contato," & _
			" t_ORCAMENTISTA_E_INDICADOR.apelido AS indicador, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota," & _
			" t_PEDIDO.st_memorizacao_completa_enderecos, t_CLIENTE.email_xml, t_CLIENTE.produtor_rural_status, t_CLIENTE.contribuinte_icms_status, " & _
			" t_PEDIDO.endereco_rg, t_PEDIDO.endereco_ie, t_PEDIDO.endereco_nome, t_PEDIDO.endereco_logradouro as pedido_endereco_logradouro, " & _
			" t_PEDIDO.endereco_numero as pedido_endereco_numero, t_PEDIDO.endereco_complemento as pedido_endereco_complemento, " & _
			" t_PEDIDO.endereco_bairro as pedido_endereco_bairro, t_PEDIDO.endereco_cidade as pedido_endereco_cidade, " & _
			" t_PEDIDO.endereco_uf as pedido_endereco_uf, t_PEDIDO.endereco_cep as pedido_endereco_cep, " & _
			" t_PEDIDO.endereco_tel_res, t_PEDIDO.endereco_ddd_res, t_PEDIDO.endereco_tel_com, t_PEDIDO.endereco_ddd_com, t_PEDIDO.endereco_ramal_com, " & _
			" t_PEDIDO.endereco_tel_cel, t_PEDIDO.endereco_ddd_cel, t_PEDIDO.endereco_tel_com_2, t_PEDIDO.endereco_ddd_com_2, t_PEDIDO.endereco_ramal_com_2, " & _
			" t_PEDIDO.endereco_email, t_PEDIDO.endereco_email_xml, t_PEDIDO.endereco_produtor_rural_status, t_PEDIDO.endereco_contribuinte_icms_status "
	
	if (vl_filtro_valor_inferior > 0) Or (vl_filtro_valor_superior > 0) then
		s_sql = s_sql & _
			", " & _
			"(" & _
				"SELECT" & _
					" Coalesce(SUM(preco_NF*qtde), 0)" & _
				" FROM t_PEDIDO_ITEM tPIAux" & _
					" INNER JOIN t_PEDIDO tPAux ON (tPIAux.pedido=tPAux.pedido)" & _
				" WHERE" & _
					" (tPAux.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
					" AND (tPAux.pedido_base = t_PEDIDO.pedido_base)" & _
			") AS vl_total_item"
		end if
	
	s_sql = s_sql & _
			s_from & _
			s_where
	
'	FILTRO: VALOR INFERIOR E VALOR SUPERIOR
	if (vl_filtro_valor_inferior > 0) Or (vl_filtro_valor_superior > 0) then
		s_where_aux = ""
		
		if vl_filtro_valor_inferior > 0 then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & _
				" (vl_total_item >= " & bd_formata_numero(vl_filtro_valor_inferior) & ")"
			end if
		
		if vl_filtro_valor_superior > 0 then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & _
				" (vl_total_item <= " & bd_formata_numero(vl_filtro_valor_superior) & ")"
			end if
		
		if s_where_aux <> "" then s_where_aux = " WHERE" & s_where_aux
		s_sql = "SELECT " & _
				"*" & _
				" FROM (" & s_sql & ") t" & _
				s_where_aux & _
				" ORDER BY" & _
					" data, pedido_base, pedido"
	else
		s_sql = s_sql & " ORDER BY t_PEDIDO__BASE.data, t_PEDIDO.pedido_base, t_PEDIDO.pedido"
		end if
	
  ' CABEÇALHO
	w_fabricante = 29
	w_produto = 54
	w_descricao = 279
	w_qtde = 26
	w_preco_lista = 69
	w_desconto = 40
	w_vl_unitario = 64
	w_vl_total = 79
	
	cab_table = "<table style='border:2px solid black;' cellspacing='0' cellpadding='0'>" & chr(13)
	cab = "	<tr style='background:mintcream;'>" & chr(13) & _
		  "		<td class='MDTE' style='width:" & cstr(w_fabricante) & "px' align='left' valign='bottom' nowrap><span class='PLTe'>Fabr</span></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_produto) & "px' align='left' valign='bottom'><span class='PLTe'>Produto</span></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_descricao) & "px' align='left' valign='bottom'><span class='PLTe'>Descrição</span></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_qtde) & "px' align='right' valign='bottom'><span class='PLTd'>Qtd</span></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_preco_lista) & "px' align='right' valign='bottom'><span class='PLTd'>Preço Lista</span></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_desconto) & "px' align='right' valign='bottom'><span class='PLTd'>Desc</span></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_vl_unitario) & "px' align='right' valign='bottom'><span class='PLTd'>Valor Unit</span></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_vl_total) & "px' align='right' valign='bottom'><span class='PLTd'>Valor Total</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	n_reg = 0
	n_pedido = 0
	pedido_a = "XXXXX"
	s_forma_pagto_a = ""
	s_descricao_forma_pagto_a = ""
	s_obs1_a = ""
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
		
		blnPulaRegistro = False
		
	'	PENDENTE CARTÃO DE CRÉDITO?
		if Not blnPulaRegistro then
			if is_forma_pagto_somente_cartao(r) And (Trim("" & r("st_pagto")) <> ST_PAGTO_PAGO) then blnPulaRegistro = True
			end if
		
	'	CARTÃO (MAQUINETA)?
		if Not blnPulaRegistro then
			if Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
			'	SE A CONSULTA FOI P/ TRATAR A FILA, OU SEJA, SEM ESPECIFICAR NENHUM NÚMERO DE PEDIDO NO FILTRO, IGNORA OS PEDIDOS QUE SEJAM PARCELADOS NO CARTÃO (MAQUINETA)
			'	MAS SE O USUÁRIO ESPECIFICOU PEDIDO(S) NO FILTRO, CONSIDERA QUE O USUÁRIO QUER REALIZAR A ALTERAÇÃO DO STATUS DA ANÁLISE DE CRÉDITO DELES
				if Not tem_filtro_pedido then blnPulaRegistro = True
				end if
			end if

		if Not blnPulaRegistro then
			mudou_pedido = (Trim("" & r("pedido")) <> pedido_a)
			
		'	IMPRIME TOTAL DO PEDIDO ANTERIOR?
			if mudou_pedido And (x <> "") then
				x = x & "	<tr>" & chr(13) & _
						"		<td colspan='8' align='left' style='border-top:1px solid #C0C0C0;'>" & chr(13)& _
						"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
						"				<tr>" & chr(13) & _
						"					<td width='20%' class='MDE' align='left' valign='bottom'><span class='Rf'>Status de Pagto</span></td>" & chr(13) & _
						"					<td width='20%' class='MD' align='right' valign='bottom'><span class='Rf'>VL Total&nbsp;&nbsp;(Família)&nbsp;</span></td>" & chr(13) & _
						"					<td width='20%' class='MD' align='right' valign='bottom'><span class='Rf'>VL Pago&nbsp;</span></td>" & chr(13) & _
						"					<td width='20%' class='MD' align='right' valign='bottom'><span class='Rf'>VL Devoluções&nbsp;</span></td>" & chr(13) & _
						"					<td width='20%' class='MD' align='right' valign='bottom'><span class='Rf'>Total dos Itens&nbsp;</span></td>" & chr(13) & _
						"				</tr>" & chr(13) & _
						"				<tr>" & chr(13) & _
						"					<td width='20%' class='MDE' align='left'><span class='C'>" & Ucase(x_status_pagto(st_pagto)) & "&nbsp;</span></td>" & chr(13) & _
						"					<td width='20%' class='MD' align='right'><span class='Cd'>" & formata_moeda(vl_TotalFamiliaPrecoNF) & "</span></td>" & chr(13)
					
				if vl_TotalFamiliaPago >= 0 then s_cor = "black" else s_cor = "red"
				x = x & "					<td width='20%' class='MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_TotalFamiliaPago) & "</span></td>" & chr(13) & _
						"					<td width='20%' class='MD' align='right'><span class='Cd'>" & formata_moeda(vl_TotalFamiliaDevolucaoPrecoNF) & "</span></td>" & chr(13) & _
						"					<td width='20%' class='MD' align='right'><span class='Cd'>" & formata_moeda(vl_total_item) & "</span></td>" & chr(13) & _
						"				</tr>" & chr(13) 
							
				x = x & "			</table>" & chr(13) & _
						"		</td>" & chr(13) & _
						"	</tr>" & chr(13)
				
			'	OBS
				x = x & "	<tr>" & chr(13) & _
										"		<td class='MDTE' colspan='8' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
										"			<span class='PLTc' style='font-weight:bold;'>Observações I</span>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)
				x = x & "	<tr>" & chr(13) & _
										"		<td class='MDTE' colspan='8' align='left' valign='bottom'>" & chr(13) & _
										"			<textarea name='c_obs1' id='c_obs1' readonly tabindex=-1 class='PLLe' rows='" & Cstr(NUM_LINHAS_OBS) & "' style='width:642px;margin-left:2pt;' onkeypress='limita_tamanho(this,MAX_TAM_OBS1);' onblur='this.value=trim(this.value);'>" & _
										s_obs1_a & "</textarea>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)
				
			'	FORMA DE PAGAMENTO
				s = s_forma_pagto_a
				if s = "" then s = "&nbsp;"
				x = x & "	<tr>" & chr(13) & _
										"		<td class='MDTE' colspan='8' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
										"			<span class='PLTc' style='font-weight:bold;'>Forma de Pagamento</span>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)
				x = x & "	<tr>" & chr(13) & _
										"		<td class='MDTE' colspan='8' align='left' valign='bottom' style='padding-left:2px;'>" & chr(13) & _
										"			<span class='Cn'>" & s & "</span>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)
			
			'	DESCRIÇÃO DA FORMA DE PAGAMENTO
				x = x & "	<tr>" & chr(13) & _
										"		<td class='MDTE' colspan='8' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
										"			<span class='PLTc' style='font-weight:bold;'>Descrição da Forma de Pagamento</span>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)
				x = x & "	<tr>" & chr(13) & _
										"		<td class='MDTE' colspan='8' align='left' valign='bottom' style='width:642px;'>" & chr(13) & _
										"			<textarea name='c_descr_forma_pagto' id='c_descr_forma_pagto' class='PLLe' rows='" & Cstr(NUM_LINHAS_DESCR_FORMA_PAGTO) & "' style='width:642px;margin-left:2pt;' onkeypress='limita_tamanho(this,MAX_TAM_FORMA_PAGTO);' onblur='this.value=trim(this.value);'>" & _
										s_descricao_forma_pagto_a & "</textarea>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)
			
			'	ANÁLISE DE CRÉDITO
				x = x & "	<tr>" & chr(13) & _
										"		<td class='MDTE' colspan='8' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
										"			<span class='PLTc' style='font-weight:bold;'>Análise de Crédito</span>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)
				x = x & "	<tr>" & chr(13) & _
										"		<td colspan='8' align='left' valign='bottom'>" & chr(13) & _
										"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
										"				<tr>" & chr(13) & _
										"					<td class='MC ME' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
										"					<td class='MC' align='left'>" & chr(13) & _
										"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_PENDENTE_VENDAS & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:red;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[0].click();'>" & x_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS) & "</span>" & chr(13) &_
										"					</td>" & chr(13) & _
										"					<td class='MC' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
										"					<td class='MC' align='left'>" & chr(13) & _
										"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_PENDENTE_ENDERECO & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:red;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[1].click();'>" & x_analise_credito(COD_AN_CREDITO_PENDENTE_ENDERECO) & "</span>" & chr(13) &_
										"					</td>" & chr(13) & _
										"					<td class='MC' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
										"					<td class='MC' align='left'>" & chr(13) & _
										"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_PENDENTE & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:red;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[2].click();'>" & x_analise_credito(COD_AN_CREDITO_PENDENTE) & "</span>" & chr(13) &_
										"					</td>" & chr(13) & _
										"					<td class='MC' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
										"					<td class='MC' align='right'>" & chr(13) & _
										"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_OK & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:green;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[3].click();'>" & x_analise_credito(COD_AN_CREDITO_OK) & "</span>" & chr(13) &_
										"					</td>" & chr(13) & _
										"					<td class='MC MD' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
										"				</tr>" & chr(13) & _
                                        "               <tr id='trPendVendasMotivo_" & Cstr(n_pedido) & "' style='display:none;'>" & chr(13) & _
                                        "                   <td class='MC ME' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
                                        "                   <td class='MC' align='left' colspan='8'>" & chr(13) & _
                                        "                       <span class='C'>Motivo: </span>" & chr(13) & _
                                        "                       <select name='c_pendente_vendas_motivo_" & Cstr(n_pedido) & "' id='c_pendente_vendas_motivo_" & Cstr(n_pedido) & "'>" & chr(13) & _
                                                                    codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__AC_PENDENTE_VENDAS_MOTIVO, Null) & chr(13) & _
                                        "                       </select>" & chr(13) & _
                                        "                   </td>" & chr(13) & _
                                        "               </tr>" & _
										"				<tr>" & chr(13) & _
										"					<td class='MC' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
										"					<td class='MC' align='left'>" & chr(13) & _
										"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:darkorange;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[4].click();'>" & x_analise_credito(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO) & "</span>" & chr(13) &_
										"					</td>" & chr(13) & _
										"					<td class='MC' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
										"					<td class='MC' colspan='5' align='left'>" & chr(13) & _
										"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:darkorange;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[5].click();'>" & x_analise_credito(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO) & "</span>" & chr(13) &_
										"					</td>" & chr(13) & _
										"					<td class='MC MD' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
										"				</tr>" & chr(13) & _
										"			</table>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)
										
				Response.Write x
				x = ""
				end if

		'	IMPRIME DADOS DO CLIENTE?
			if mudou_pedido then
				vl_total_item = 0
				
			'	PULA LINHA ENTRE PEDIDOS
				if n_reg > 0 then
					x = x & "</table>" & chr(13) & _
							"<br>" & chr(13) & _
							"<br>" & chr(13) & _
							cab_table
				else
					x = cab_table
					end if
					
			'	LOJA
				s = Trim("" & r("loja"))
				s_aux = x_loja(s)
				if (s<>"") And (s_aux<>"") then s = s & " - "
				s = s & s_aux
				x = x & "	<tr>" & chr(13) & _
										"		<td class='MDTE' colspan='8' align='left' valign='bottom' style='background:mintcream;'>" & chr(13) & _
										"			<span class='N' style='margin-left:12px;'>" & s & "</span>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)
										
			'	VENDEDOR / INDICADOR
				s_indicador = Trim("" & r("indicador"))
				if s_indicador = "" then s_indicador = "&nbsp;"
				s_desempenho_nota = Trim("" & r("desempenho_nota"))
				if s_desempenho_nota = "" then 
					s_desempenho_nota = "&nbsp;"
				else
					s_desempenho_nota = "(" & s_desempenho_nota & ") "
					end if
					
				s = Trim("" & r("vendedor"))
				s_aux = Trim("" & r("nome_vendedor"))
				if (s<>"") And (s_aux<>"") then s_aux = " (" & s_aux & ")"
				s = s & s_aux
				x = x & "	<tr>" & chr(13) & _
							"		<td class='MDTE' colspan='8' align='left'>" & chr(13) & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
										"		<td class='MD' align='left' valign='bottom'>" & chr(13) & _
										"			<span class='Cn' style='font-weight:bold;margin-left:12px;'>" & s & "</span>" & chr(13) & _
										"		</td>" & chr(13) & _
										"		<td width='25%' align='left' valign='bottom'>" & chr(13) & _
										"			<span class='Cn' style='font-weight:bold;margin-left:12px;'>" & s_desempenho_nota & s_indicador & "</span>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13) & _
							"			</table>" & _
							"		</td>" & chr(13) & _
							"	</tr>" & chr(13)
			
			'	OBTÉM OUTROS DADOS DO PEDIDO
				if Not calcula_pagamentos(Trim("" & r("pedido")), vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				vl_saldo_a_pagar = vl_TotalFamiliaPrecoNF-vl_TotalFamiliaPago-vl_TotalFamiliaDevolucaoPrecoNF
				
			'	Nº PEDIDO (EXIBE TODA A FAMÍLIA DE PEDIDOS, PRINCIPALMENTE PORQUE O PEDIDO-BASE PODE TER SIDO CANCELADO E OS PEDIDOS-FILHOTE NÃO)
				s_sql = "SELECT" & _
							" pedido_base," & _
							" pedido," & _
							" data," & _
							" hora," & _
							" st_entrega," & _
							" entregue_data," & _
							" cancelado_data" & _
						" FROM t_PEDIDO" & _
						" WHERE" & _
							" (pedido_base = '" & Trim("" & r("pedido_base")) & "')" & _
						" ORDER BY" & _
							" pedido"
				if tP.State <> 0 then tP.Close
				tP.Open s_sql, cn
				
				iCountPedFamilia = 0
				do while Not tP.Eof
					iCountPedFamilia = iCountPedFamilia + 1
					s = ""
					if Trim("" & tP("st_entrega")) = ST_ENTREGA_ENTREGUE then
						s = formata_data(tP("entregue_data"))
					elseif Trim("" & tP("st_entrega")) = ST_ENTREGA_CANCELADO then
						s = formata_data(tP("cancelado_data"))
						end if
					if s<>"" then s="  (" & s & ")"
				
					x = x & "	<tr>" & chr(13) & _
							"		<td class='MDTE' colspan='8' align='left'>" & chr(13) & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
							"				<tr>" & chr(13) & _
							"					<td style='width:80px;' align='left'>" & chr(13)
					
					if iCountPedFamilia = 1 then
					'	CRIA O CAMPO HIDDEN SOMENTE P/ O 1º PEDIDO DA LISTA DE PEDIDOS DA FAMÍLIA, CASO CONTRÁRIO, OCORRERIA A 'QUEBRA' DA SINCRONIA
					'	COM OS DEMAIS CAMPOS DO FORMULÁRIO QUE CONTROLAM A GRAVAÇÃO DOS DADOS, JÁ QUE OS CAMPOS ESTÃO ASSOCIADOS PELO SEU ÍNDICE NO
					'	ARRAY DE CAMPOS.
						x = x & _
							"						<input type='hidden' name='c_pedido' id='c_pedido' value='" & Trim("" & tP("pedido")) & "'>" & chr(13)
						end if

					x = x & "						<span class='Cn' style='font-weight:bold;margin-left:12px;'>" & _
													Trim("" & tP("pedido")) & "</span>" & chr(13) &_
							"					</td>" & chr(13) & _
							"					<td style='width:80px;' align='left'>" & chr(13) & _
							"						<span class='Cn'>" & formata_data(tP("data")) & "</span>" & chr(13) & _
							"					</td>" & chr(13) & _
							"					<td align='left' align='left'>" & chr(13) & _
							"						<span class='Cn'>" & x_status_entrega(tP("st_entrega")) & s & "</span>" & chr(13) & _
							"					</td>" & chr(13) & _
							"				</tr>" & chr(13) & _
							"			</table>" & chr(13) & _
							"		</td>" & chr(13) & _
							"	</tr>" & chr(13)
					tP.MoveNext
					loop
				
			'	PEDIDOS ANTERIORES (CONTAGEM SOMENTE POR PEDIDO-BASE)
				n_pedidos_anteriores = 0
				s_sql = "SELECT" & _
							" Count(*) AS qtde" & _
						" FROM t_PEDIDO" & _
						" WHERE" & _
							" (id_cliente = '" & Trim("" & r("id_cliente")) & "')" & _
							" AND (pedido = pedido_base)" & _
							" AND (pedido_base <> '" & Trim("" & r("pedido_base")) & "')" & _
							" AND ((CONVERT(varchar(8),data,112)+hora) < '" & formata_data_yyyymmdd(r("data")) & r("hora") & "')"
						
				if rs.State <> 0 then rs.Close
				rs.open s_sql, cn
				if Not rs.Eof then
					if Not IsNull(rs("qtde")) then n_pedidos_anteriores = rs("qtde")
					end if
						
				x = x & "	<tr>" & chr(13) & _
										"		<td class='MDTE' colspan='8' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
										"			<span class='PLTc' style='font-weight:bold;'>Pedidos Anteriores:&nbsp;&nbsp;" & formata_inteiro(n_pedidos_anteriores) & "</span>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)
										
				s_sql = "SELECT TOP 5 data, pedido_base, pedido, forma_pagto, tipo_parcelamento, av_forma_pagto," & _
						" pu_forma_pagto, pu_valor, pu_vencto_apos," & _
						" pc_qtde_parcelas, pc_valor_parcela," & _
						" pc_maquineta_qtde_parcelas, pc_maquineta_valor_parcela," & _
						" pce_forma_pagto_entrada, pce_forma_pagto_prestacao, pce_entrada_valor, pce_prestacao_qtde," & _
						" pce_prestacao_valor, pce_prestacao_periodo, pse_forma_pagto_prim_prest," & _
						" pse_forma_pagto_demais_prest, pse_prim_prest_valor, pse_prim_prest_apos," & _
						" pse_demais_prest_qtde, pse_demais_prest_valor, pse_demais_prest_periodo" & _
						" FROM t_PEDIDO" & _
						" WHERE (id_cliente = '" & Trim("" & r("id_cliente")) & "')" & _
						" AND (pedido = pedido_base)" & _
						" AND (pedido_base <> '" & Trim("" & r("pedido_base")) & "')" & _
						" AND (data_hora < CONVERT(datetime, '" & formata_data_com_separador_yyyymmdd(r("data"),"-") & " " & formata_hhnnss_para_hh_nn_ss(Trim("" & r("hora"))) & "', 120))" & _
						" ORDER BY data_hora DESC, pedido DESC"
				if rs.State <> 0 then rs.Close
				rs.open s_sql, cn
				if rs.Eof then
					x = x & "	<tr>" & chr(13) & _
											"		<td class='MDTE' colspan='8' align='center' valign='bottom'>" & chr(13) & _
											"			<span class='Cn' style='text-align:center;'>(Nenhum)</span>" & chr(13) & _
											"		</td>" & chr(13) & _
											"	</tr>" & chr(13)
				else
					x = x & "	<tr>" & chr(13) & _
							"		<td colspan='8' align='left'>" & chr(13) & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13)
					do while Not rs.Eof
						s_forma_pagto_ped_ant = ""
						if rs("tipo_parcelamento") = 0 then
						'	VERSÃO ANTIGA DA FORMA DE PAGAMENTO
							s_forma_pagto_ped_ant = substitui_caracteres(Trim("" & rs("forma_pagto")), chr(13), "<br>")
						else
						'	VERSÃO NOVA DA FORMA DE PAGAMENTO
							if Trim("" & rs("tipo_parcelamento")) = COD_FORMA_PAGTO_A_VISTA then
								s_forma_pagto_ped_ant = "À Vista  (" & x_opcao_forma_pagamento(rs("av_forma_pagto")) & ")"
							elseif Trim("" & rs("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELA_UNICA then
								s_forma_pagto_ped_ant = "Parcela Única:  " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("pu_valor")) & "  (" & x_opcao_forma_pagamento(rs("pu_forma_pagto")) & ")  vencendo após " & Cstr(rs("pu_vencto_apos")) & " dias"
							elseif Trim("" & rs("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_CARTAO then
								s_forma_pagto_ped_ant = "Parcelado no Cartão (internet) em " & Cstr(rs("pc_qtde_parcelas")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("pc_valor_parcela"))
							elseif Trim("" & rs("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
								s_forma_pagto_ped_ant = "Parcelado no Cartão (maquineta) em " & Cstr(rs("pc_maquineta_qtde_parcelas")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("pc_maquineta_valor_parcela"))
							elseif Trim("" & rs("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
								s_forma_pagto_ped_ant = "Entrada:  " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("pce_entrada_valor")) & "  (" & x_opcao_forma_pagamento(rs("pce_forma_pagto_entrada")) & ")" & _
										chr(13) & "<br>" & chr(13) & _
										"Prestações:  " & Cstr(rs("pce_prestacao_qtde")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("pce_prestacao_valor")) & _
										"  (" & x_opcao_forma_pagamento(rs("pce_forma_pagto_prestacao")) & ")  vencendo a cada " & _
										Cstr(rs("pce_prestacao_periodo")) & " dias"
							elseif Trim("" & rs("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
								s_forma_pagto_ped_ant = "1ª Prestação:  " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("pse_prim_prest_valor")) & "  (" & x_opcao_forma_pagamento(rs("pse_forma_pagto_prim_prest")) & ")  vencendo após " & Cstr(rs("pse_prim_prest_apos")) & " dias" & _
										chr(13) & "<br>" & chr(13) & _
										"Demais Prestações:  " & Cstr(rs("pse_demais_prest_qtde")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("pse_demais_prest_valor")) & _
										"  (" & x_opcao_forma_pagamento(rs("pse_forma_pagto_demais_prest")) & ")  vencendo a cada " & _
										Cstr(rs("pse_demais_prest_periodo")) & " dias"
								end if
							s = substitui_caracteres(Trim("" & rs("forma_pagto")), chr(13), "<br>")
							if (s_forma_pagto_ped_ant<>"") And (s<>"") then s_forma_pagto_ped_ant = s_forma_pagto_ped_ant & chr(13) & "<br>"
							s_forma_pagto_ped_ant = s_forma_pagto_ped_ant & s
							end if
						
						if s_forma_pagto_ped_ant = "" then s_forma_pagto_ped_ant = "&nbsp;"
						x = x & _
							"				<tr>" & chr(13) & _
							"					<td class='MDTE' style='width:50px;' align='left' nowrap>" & chr(13) & _
							"						<span class='Cn'>" & _
													Trim("" & rs("pedido")) & "</span>" & chr(13) &_
							"					</td>" & chr(13) & _
							"					<td class='MTD' align='left' style='width:591px;padding-left:2px;'>" & chr(13) & _
							"						<span class='Cn'>" & s_forma_pagto_ped_ant & "</span>" & chr(13) & _
							"					</td>" & chr(13) & _
							"				</tr>" & chr(13)
						rs.MoveNext
						loop
					
					x = x & _
							"			</table>" & chr(13) & _
							"		</td>" & chr(13) & _
							"	</tr>" & chr(13)
					end if		
				
				
			'	DADOS DO CLIENTE
				x = x & "	<tr>" & chr(13) & _
										"		<td class='MDTE' colspan='8' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
										"			<span class='PLTc' style='font-weight:bold;'>Cliente</span>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)
				
				dim cliente__nome_iniciais_em_maiusculas, cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento
				dim cliente__bairro, cliente__cidade, cliente__uf, cliente__cep, cliente__cnpj_cpf
				dim cliente__tel_res, cliente__ddd_res, cliente__tel_com, cliente__ddd_com, cliente__ramal_com
				dim cliente__ie, cliente__rg, cliente__email, cliente__email_xml
				cliente__nome_iniciais_em_maiusculas = Trim("" & r("nome_iniciais_em_maiusculas"))
				cliente__endereco = Trim("" & r("endereco"))
				cliente__endereco_numero = Trim("" & r("endereco_numero"))
				cliente__endereco_complemento = Trim("" & r("endereco_complemento"))
				cliente__bairro = Trim("" & r("bairro"))
				cliente__cidade = Trim("" & r("cidade"))
				cliente__uf = Trim("" & r("uf"))
				cliente__cep = Trim("" & r("cep"))
				cliente__cnpj_cpf = Trim("" & r("cnpj_cpf"))
				cliente__tel_res = Trim("" & r("tel_res"))
				cliente__ddd_res = Trim("" & r("ddd_res"))
				cliente__tel_com = Trim("" & r("tel_com"))
				cliente__ddd_com = Trim("" & r("ddd_com"))
				cliente__ramal_com = Trim("" & r("ramal_com"))
				cliente__ie = Trim("" & r("ie"))
				cliente__rg = Trim("" & r("rg"))
				cliente__email = Trim("" & r("email"))
				cliente__email_xml = Trim("" & r("email_xml"))
				if Trim("" & r("st_memorizacao_completa_enderecos")) <> 0 then
					cliente__nome_iniciais_em_maiusculas = iniciais_em_maiusculas(Trim("" & r("endereco_nome")))
					cliente__endereco = Trim("" & r("pedido_endereco_logradouro"))
					cliente__endereco_numero = Trim("" & r("pedido_endereco_numero"))
					cliente__endereco_complemento = Trim("" & r("pedido_endereco_complemento"))
					cliente__bairro = Trim("" & r("pedido_endereco_bairro"))
					cliente__cidade = Trim("" & r("pedido_endereco_cidade"))
					cliente__uf = Trim("" & r("pedido_endereco_uf"))
					cliente__cep = Trim("" & r("pedido_endereco_cep"))
					cliente__cnpj_cpf = Trim("" & r("cnpj_cpf")) ' este usamos o principal porque sempre é igual
					cliente__tel_res = Trim("" & r("endereco_tel_res"))
					cliente__ddd_res = Trim("" & r("endereco_ddd_res"))
					cliente__tel_com = Trim("" & r("endereco_tel_com"))
					cliente__ddd_com = Trim("" & r("endereco_ddd_com"))
					cliente__ramal_com = Trim("" & r("endereco_ramal_com"))
					cliente__ie = Trim("" & r("endereco_ie"))
					cliente__rg = Trim("" & r("endereco_rg"))
					cliente__email = Trim("" & r("endereco_email"))
					cliente__email_xml = Trim("" & r("endereco_email_xml"))
					end if

			'	nome
				s_nome = "&nbsp;"
				if cliente__nome_iniciais_em_maiusculas <> "" then s_nome = cliente__nome_iniciais_em_maiusculas
			'	endereço
				s_endereco = "&nbsp;"
				if cliente__endereco <> "" then
					s_endereco = iniciais_em_maiusculas(cliente__endereco)
					s = cliente__endereco_numero
					if s<>"" then s_endereco=s_endereco & ", " & s
					s = cliente__endereco_complemento
					if s<>"" then s_endereco=s_endereco & " " & s
					s = iniciais_em_maiusculas(cliente__bairro)
					if s<>"" then s_endereco=s_endereco & " - " & s
					s = iniciais_em_maiusculas(cliente__cidade)
					if s<>"" then s_endereco=s_endereco & " - " & s
					s=UCase(cliente__uf)
					if s<>"" then s_endereco=s_endereco & " - " & s
					s=cliente__cep
					if s<>"" then s_endereco=s_endereco & " - " & cep_formata(s)
					end if
			'	endereço de entrega
				s_endereco_entrega = "&nbsp;"
				if r("st_end_entrega") <> 0 then
					s_endereco_entrega = formata_endereco(iniciais_em_maiusculas(Trim("" & r("EndEtg_endereco"))), Trim("" & r("EndEtg_endereco_numero")), Trim("" & r("EndEtg_endereco_complemento")), iniciais_em_maiusculas(Trim("" & r("EndEtg_bairro"))), iniciais_em_maiusculas(Trim("" & r("EndEtg_cidade"))), Ucase(Trim("" & r("EndEtg_uf"))), retorna_so_digitos(Trim("" & r("EndEtg_cep"))))
					end if
				
			'	cnpj/cpf
				s_cnpj_cpf = "CPF: "
				if cliente__cnpj_cpf <> "" then
					s_cnpj_cpf = cnpj_cpf_formata(cliente__cnpj_cpf)
					if Len(cliente__cnpj_cpf) = 14 then
						s_cnpj_cpf = "CNPJ: " & s_cnpj_cpf
					else
						s_cnpj_cpf = "CPF: " & s_cnpj_cpf
						end if
					end if

			'	telefone residencial
				s_tel_res = ""
				if cliente__tel_res <> "" then
					s = cliente__tel_res
					s_tel_res = telefone_formata(s)
					s = cliente__ddd_res
					if s <> "" then s_tel_res = "(" & s & ") " & s_tel_res
					end if
				s_tel_res = "Tel Res: " & s_tel_res
				
			'	telefone comercial
				s_tel_com = ""
				if cliente__tel_com <> "" then
					s = cliente__tel_com
					s_tel_com = telefone_formata(s)
					s = cliente__ddd_com
					if s <> "" then s_tel_com = "(" & s & ") " & s_tel_com
					s = cliente__ramal_com
					if s<>"" then s_tel_com = s_tel_com & "  (R." & s & ")"
					end if
				s_tel_com = "Tel Com: " & s_tel_com
				
			'	rg
				s_rg = "&nbsp;"
				if Trim("" & r("tipo")) = ID_PJ then
					if cliente__ie <> "" then s_rg = "IE: " & cliente__ie
				else
					if cliente__rg <> "" then s_rg = "RG: " & cliente__rg
					end if
			'	e-mail
				s_email = ""
				if cliente__email <> "" then
					s_email = "E-mail: " & cliente__email
					end if
			'	e-mail-xml
				s_email_xml = ""
				if cliente__email_xml <> "" then
					s_email_xml = "E-mail (XML): " & cliente__email_xml
					end if
			'	concatena e-mail e e-mail-xml
				if s_email = "" then
					s_email = s_email_xml
				else
					if s_email_xml <> "" then
						s_email = s_email  & " - " & s_email_xml
						end if
					end if
				if s_email = "" then
					s_email = "E-mail: "
					end if

			'	Referência Bancária
				strRefBanc = ""
				if (Trim("" & r("tipo")) = ID_PF) Or (Trim("" & r("tipo")) = ID_PJ) then
					strRefBanc = strRefBanc & _
						"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
						"				<tr>" & chr(13) & _
						"					<td class='MC' colspan='2' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
						"						<span class='PLTc' style='font-weight:bold;'>Referência Bancária</span>" & chr(13) & _
						"					</td>" & chr(13) & _
						"				</tr>" & chr(13) & _
						"			</table>" & chr(13)
					Call le_cliente_ref_bancaria(Trim("" & r("id_cliente")), vRefBancaria, msg_erro)
					intIndice = 0
					for intCounter=Lbound(vRefBancaria) to Ubound(vRefBancaria)
						with vRefBancaria(intCounter)
							if Trim(.id_cliente) <> "" then
								intIndice = intIndice + 1
								strRefBanc = strRefBanc & _
									"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
									"				<tr>" & chr(13) & _
									"					<td class='MC MD' align='left' valign='bottom'>" & _
															"<span class='Cn'>" & _
															iniciais_em_maiusculas(x_banco(Trim("" & .banco))) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"					<td class='MC MD' align='left' valign='bottom' width='25%'>" & _
															"<span class='Cn'>" & _
															"Agência: " & Trim("" & .agencia) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"					<td class='MC' align='left' valign='bottom' width='25%'>" & _
															"<span class='Cn'>" & _
															"Conta: " & Trim("" & .conta) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"				</tr>" & chr(13) & _
									"			</table>" & chr(13) & _
									"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
									"				<tr>" & chr(13) & _
									"					<td class='MC MD' align='left' valign='bottom'>" & _
															"<span class='Cn'>Tel:" & _
															" (" & _
															Trim("" & .ddd) & _
															") " & _
															telefone_formata(Trim("" & .telefone)) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"					<td class='MC' align='left' valign='bottom' width='50%'>" & _
															"<span class='Cn'>" & _
															"Contato: " & Trim("" & .contato) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"				</tr>" & chr(13) & _
									"			</table>" & chr(13)
								end if
							end with
						next
					
					if intIndice = 0 then
					'	Linha em branco (p/ evidenciar ausência de dados)
						strRefBanc = strRefBanc & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
							"				<tr>" & chr(13) & _
							"					<td class='MC' align='left' valign='bottom'>" & _
													"<span class='Cn'>" & _
													"&nbsp;" & _
													"</span>" & _
												"</td>" & chr(13) & _
							"				</tr>" & chr(13) & _
							"			</table>" & chr(13)
						end if
					end if

			'	Referência Profissional
				strRefProf = ""
				if Trim("" & r("tipo")) = ID_PF then
					strRefProf = strRefProf & _
						"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
						"				<tr>" & chr(13) & _
						"					<td class='MC' colspan='2' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
						"						<span class='PLTc' style='font-weight:bold;'>Referência Profissional</span>" & chr(13) & _
						"					</td>" & chr(13) & _
						"				</tr>" & chr(13) & _
						"			</table>" & chr(13)
					Call le_cliente_ref_profissional(Trim("" & r("id_cliente")), vRefProfissional, msg_erro)
					intIndice = 0
					for intCounter=Lbound(vRefProfissional) to Ubound(vRefProfissional)
						with vRefProfissional(intCounter)
							if Trim(.id_cliente) <> "" then
								intIndice = intIndice + 1
								
								strRefProfCnpj=Trim(.cnpj)
								if strRefProfCnpj <> "" then 
									strRefProfCnpj=" (" & cnpj_cpf_formata(strRefProfCnpj) & ")"
									end if
								
								strRefProf = strRefProf & _
									"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
									"				<tr>" & chr(13) & _
									"					<td class='MC MD' align='left' valign='bottom'>" & _
															"<span class='Cn'>" & _
															iniciais_em_maiusculas(Trim("" & .nome_empresa)) & _
															strRefProfCnpj & _
															"</span>" & _
														"</td>" & chr(13) & _
									"					<td class='MC' align='left' valign='bottom' width='25%'>" & _
															"<span class='Cn'>Tel:" & _
															" (" & _
															Trim("" & .ddd) & _
															") " & _
															telefone_formata(Trim("" & .telefone)) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"				</tr>" & chr(13) & _
									"			</table>" & chr(13) & _
									"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
									"				<tr>" & chr(13) & _
									"					<td class='MC MD' align='left' valign='bottom' width='25%'>" & _
															"<span class='Cn'>" & _
															"Período Registro: " & formata_data(.periodo_registro) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"					<td class='MC MD' align='left' valign='bottom'>" & _
															"<span class='Cn'>" & _
															"Cargo: " & iniciais_em_maiusculas(Trim("" & .cargo)) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"					<td class='MC' align='left' valign='bottom' width='25%'>" & _
															"<span class='Cn'>" & _
															"Rendimentos: " & formata_moeda(.rendimentos) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"				</tr>" & chr(13) & _
									"			</table>" & chr(13)
								end if
							end with
						next
					
					if intIndice = 0 then
					'	Linha em branco (p/ evidenciar ausência de dados)
						strRefProf = strRefProf & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
							"				<tr>" & chr(13) & _
							"					<td class='MC' align='left' valign='bottom'>" & _
													"<span class='Cn'>" & _
													"&nbsp;" & _
													"</span>" & _
												"</td>" & chr(13) & _
							"				</tr>" & chr(13) & _
							"			</table>" & chr(13)
						end if
					end if

			'	Referência Comercial
				strRefCom = ""
				if Trim("" & r("tipo")) = ID_PJ then
					strRefCom = strRefCom & _
						"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
						"				<tr>" & chr(13) & _
						"					<td class='MC' colspan='2' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
						"						<span class='PLTc' style='font-weight:bold;'>Referência Comercial</span>" & chr(13) & _
						"					</td>" & chr(13) & _
						"				</tr>" & chr(13) & _
						"			</table>" & chr(13)
					Call le_cliente_ref_comercial(Trim("" & r("id_cliente")), vRefComercial, msg_erro)
					intIndice = 0
					for intCounter=Lbound(vRefComercial) to Ubound(vRefComercial)
						with vRefComercial(intCounter)
							if Trim(.id_cliente) <> "" then
								intIndice = intIndice + 1
								strRefCom = strRefCom & _
									"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
									"				<tr>" & chr(13) & _
									"					<td class='MC MD' align='left' valign='bottom' style='width:10px;'>" & _
															"<span class='Cc' style='font-weight:bold;color:#808080;'>" & _
																"(" & CStr(intIndice) & ")" & _
															"</span> " & _
									"					<td class='MC MD' align='left' valign='bottom'>" & _
															"<span class='Cn'>" & _
															iniciais_em_maiusculas(Trim("" & .nome_empresa)) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"					<td class='MC MD' align='left' valign='bottom' width='25%'>" & _
															"<span class='Cn'>Tel:" & _
															" (" & _
															Trim("" & .ddd) & _
															") " & _
															telefone_formata(Trim("" & .telefone)) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"					<td class='MC' align='left' valign='bottom' width='25%'>" & _
															"<span class='Cn'>" & _
															"Contato: " & Trim("" & .contato) & _
															"</span>" & _
														"</td>" & chr(13) & _
									"				</tr>" & chr(13) & _
									"			</table>" & chr(13)
								end if
							end with
						next
					
					if intIndice = 0 then
					'	Linha em branco (p/ evidenciar ausência de dados)
						strRefCom = strRefCom & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
							"				<tr>" & chr(13) & _
							"					<td class='MC' align='left' valign='bottom'>" & _
													"<span class='Cn'>" & _
													"&nbsp;" & _
													"</span>" & _
												"</td>" & chr(13) & _
							"				</tr>" & chr(13) & _
							"			</table>" & chr(13)
						end if
					end if

			'	Dados do sócio majoritário
				strSocMaj = ""
				if Trim("" & r("tipo")) = ID_PJ then
					strSocMaj = strSocMaj & _
						"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
						"				<tr>" & chr(13) & _
						"					<td class='MC' colspan='2' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
						"						<span class='PLTc' style='font-weight:bold;'>Dados do Sócio Majoritário</span>" & chr(13) & _
						"					</td>" & chr(13) & _
						"				</tr>" & chr(13) & _
						"			</table>" & chr(13)
					if Trim("" & r("SocMaj_Nome")) = "" then
					'	Linha em branco (p/ evidenciar ausência de dados)
						strSocMaj = strSocMaj & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
							"				<tr>" & chr(13) & _
							"					<td class='MC' align='left' valign='bottom'>" & _
													"<span class='Cn'>" & _
													"&nbsp;" & _
													"</span>" & _
												"</td>" & chr(13) & _
							"				</tr>" & chr(13) & _
							"			</table>" & chr(13)
					else
						strSocMaj = strSocMaj & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
							"				<tr>" & chr(13) & _
							"					<td class='MC MD' align='left' valign='bottom'>" & _
													"<span class='Cn'>" & _
													iniciais_em_maiusculas(Trim("" & r("SocMaj_Nome"))) & _
													"</span>" & _
												"</td>" & chr(13) & _
							"					<td class='MC' align='left' valign='bottom' width='25%'>" & _
													"<span class='Cn'>" & _
													"CPF: " & Trim("" & r("SocMaj_CPF")) & _
													"</span>" & _
												"</td>" & chr(13) & _
							"				</tr>" & chr(13) & _
							"			</table>" & chr(13) & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
							"				<tr>" & chr(13) & _
							"					<td class='MC MD' align='left' valign='bottom'>" & _
													"<span class='Cn'>" & _
													iniciais_em_maiusculas(x_banco(Trim("" & r("SocMaj_banco")))) & _
													"</span>" & _
												"</td>" & chr(13) & _
							"					<td class='MC MD' align='left' valign='bottom' width='25%'>" & _
													"<span class='Cn'>" & _
													"Agência: " & Trim("" & r("SocMaj_agencia")) & _
													"</span>" & _
												"</td>" & chr(13) & _
							"					<td class='MC' align='left' valign='bottom' width='25%'>" & _
													"<span class='Cn'>" & _
													"Conta: " & Trim("" & r("SocMaj_conta")) & _
													"</span>" & _
												"</td>" & chr(13) & _
							"				</tr>" & chr(13) & _
							"			</table>" & chr(13) & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
							"				<tr>" & chr(13) & _
							"					<td class='MC MD' align='left' valign='bottom'>" & _
													"<span class='Cn'>Tel:" & _
													" (" & _
													Trim("" & r("SocMaj_ddd")) & _
													") " & _
													telefone_formata(Trim("" & r("SocMaj_telefone"))) & _
													"</span>" & _
												"</td>" & chr(13) & _
							"					<td class='MC' align='left' valign='bottom' width='50%'>" & _
													"<span class='Cn'>" & _
													"Contato: " & Trim("" & r("SocMaj_contato")) & _
													"</span>" & _
												"</td>" & chr(13) & _
							"				</tr>" & chr(13) & _
							"			</table>" & chr(13)
						end if
					end if
				
				x = x & "	<tr>" & chr(13) & _
						"		<td class='MDTE' colspan='8' align='left'>" & chr(13) & _
						"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
						"				<tr>" & chr(13) & _
						"					<td colspan='2' class='MD' align='left' valign='bottom'><span class='Cn'>" & s_nome & "</span></td>" & chr(13) & _
						"					<td align='left' valign='bottom'><span class='Cn'>" & s_cnpj_cpf & "</span></td>" & chr(13) & _
						"				</tr>" & chr(13) & _
						"				<tr>" & chr(13) & _
						"					<td colspan='3' class='MC' align='left' valign='bottom'><span class='Cn'>" & s_endereco & "</span></td>" & chr(13) & _
						"				</tr>" & chr(13)
				if r("st_end_entrega") <> 0 then
					x = x & _
						"				<tr>" & chr(13) & _
						"					<td colspan='3' class='MC' align='left' valign='bottom'><span class='Cni'>End. Entrega: </span><span class='Cn'>" & s_endereco_entrega & "</span></td>" & chr(13) & _
						"				</tr>" & chr(13)
					end if
				x = x & _
						"				<tr>" & chr(13) & _
						"					<td width='34%' class='MTD' align='left' valign='bottom'><span class='Cn'>" & s_tel_res & "</span></td>" & chr(13) & _
						"					<td width='33%' class='MTD' align='left' valign='bottom'><span class='Cn'>" & s_tel_com & "</span></td>" & chr(13) & _
						"					<td width='33%' class='MC' align='left' valign='bottom'><span class='Cn'>" & s_rg & "</span></td>" & chr(13) & _
						"				</tr>" & chr(13) & _
						"				<tr>" & chr(13) & _
						"					<td colspan='3' class='MC' align='left' valign='bottom'><span class='Cn'>" & s_email & "</span></td>" & chr(13) & _
						"				</tr>" & chr(13) & _
						"			</table>" & chr(13) & _
									strRefBanc & _
									strRefProf & _
									strRefCom & _
									strSocMaj & _
						"		</td>" & chr(13) & _
						"	</tr>" & chr(13)
				
			'	ENDEREÇO PARECE SER IGUAL AO USADO POR OUTRO CLIENTE ANTERIORMENTE (POSSÍVEL FRAUDE)?
				if r("analise_endereco_tratar_status") <> 0 then
					strInfoAnEnd = ""
					x = x & _
						"	<tr>" & chr(13) & _
						"		<td class='MDTE' colspan='8' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
						"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
						"				<tr>" & chr(13) & _
						"					<td style='width:250px;' align='left'>&nbsp;</td>" & chr(13) & _
						"					<td align='center' valign='middle'>" & _
												"<span id='spanTitAnEnd_" & Cstr(n_pedido+1) & "' class='PLTc TIT_INFO_AN_END_BLOCO' style='font-weight:bold;color:red;vertical-align:middle;'>Análise do Endereço</span>" & chr(13) & _
												"<a id='hrefTitAnEnd_" & Cstr(n_pedido+1) & "' href='javascript:exibeOcultaTodosInfoAnEnd(" & chr(34) & Cstr(n_pedido+1) & chr(34) & ");' title='clique para exibir mais detalhes'>" & _
												"<img id='imgPlusMinusTitAnEnd_" & Cstr(n_pedido+1) & "' style='vertical-align:bottom;margin-bottom:2px;' src='../imagem/plus.gif' />" & _
												"</a>" & _
											"</td>" & chr(13) & _
						"					<td style='width:250px;' align='right' valign='middle'>"
					if r("analise_endereco_tratado_status") = 0 then
						x = x & _
												"<input type='checkbox' tabindex='-1' id='ckb_analise_endereco_" & Cstr(n_pedido+1) & "' name='ckb_analise_endereco_" & Cstr(n_pedido+1) & "' value='AN_END_MARCAR_JA_TRATADO_OK'><span class='PLTdi' onclick='fREL.ckb_analise_endereco_" & Cstr(n_pedido+1) & ".click();'>Marcar como já analisado (OK)</span>"
						end if
					x = x & _
												"&nbsp;</td>" & chr(13) & _
						"				</tr>" & chr(13) & _
						"			</table>" & chr(13) & _
						"		</td>" & chr(13) & _
						"	</tr>" & chr(13) & _
						"	<tr>" & chr(13) & _
						"		<td class='MDTE' colspan='8' align='left'>" & chr(13) & _
						"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13)
					
				'	EXIBE DADOS COLETADOS P/ ANÁLISE
					intQtdeTotalPedidosAnEndereco = 0
					intQtdePedido = 0
					intQtdeLinhasPedido = 0
					
				'	VERIFICA SE HÁ COINCIDÊNCIA C/ ENDEREÇO DO PARCEIRO
					blnAnEnderecoUsaEndParceiro = False
					s_sql = "SELECT" & _
								" tP.indicador," & _
								" tOI.razao_social_nome_iniciais_em_maiusculas AS nome_indicador," & _
								" tOI.cnpj_cpf AS cnpj_cpf_indicador," & _
								" tPAEC.*" & _
							" FROM t_PEDIDO_ANALISE_ENDERECO tPAE" & _
								" INNER JOIN t_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO tPAEC ON (tPAE.id = tPAEC.id_pedido_analise_endereco)" & _
								" LEFT JOIN t_PEDIDO tP ON (tPAE.pedido = tP.pedido)" & _
								" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR tOI ON (tP.indicador = tOI.apelido)" & _
							" WHERE" & _
								" (tPAE.pedido = '" & Trim("" & r("pedido")) & "')" & _
								" AND (tPAEC.tipo_endereco = '" & COD_PEDIDO_AN_ENDERECO__END_PARCEIRO & "')" & _
							" ORDER BY" & _
								" tPAE.id," & _
								" tPAEC.id"
					if rs.State <> 0 then rs.Close
					rs.open s_sql, cn
					do while Not rs.Eof
						blnAnEnderecoUsaEndParceiro = True
						intQtdeTotalPedidosAnEndereco = intQtdeTotalPedidosAnEndereco + 1
						if intQtdeTotalPedidosAnEndereco > MAX_AN_ENDERECO_QTDE_PEDIDOS_EXIBICAO then exit do
						intResto = intQtdePedido Mod MAX_PEDIDOS_POR_LINHA_ANALISE_ENDERECO
						if (intQtdePedido = 0) Or (intResto = 0) then
							intQtdePedido = 0
							if intQtdeLinhasPedido > 0 then
								x = x & "				</tr>" & chr(13)
								end if
							x = x & "				<tr>" & chr(13)
							intQtdeLinhasPedido = intQtdeLinhasPedido + 1
							end if
						
						x = x & _
							"					<td align='left' valign='bottom'>" & chr(13) & _
								"<span class='C' id='spanPedidoAnEnd_" & Trim("" & rs("id")) & "'>Indicador</span>" & _
							"<a id='hrefPedAnEnd_" & Trim("" & rs("id")) & "' class='hrefAnEndBloco_" & Cstr(n_pedido+1) & "' href='javascript:exibeOcultaInfoAnEnd(" & chr(34) & Trim("" & rs("id")) & chr(34) & ");' title='clique para exibir mais detalhes'>" & _
								"<img id='imgPlusMinusPedAnEnd_" & Trim("" & rs("id")) & "' class='imgPlusMinusAnEndBloco_" & Cstr(n_pedido+1) & "' style='vertical-align:bottom;margin-bottom:0px;' src='../imagem/plus.gif' />" & _
							"</a>" & _
							"					</td>" & chr(13)
						
						strInfoAnEnd = strInfoAnEnd & _
							"	<tr id='TR_INFO_AN_END_LN1_" & Trim("" & rs("id")) & "' class='TR_INFO_AN_END TR_INFO_AN_END_BLOCO_" & Cstr(n_pedido+1) & "'>" & chr(13) & _
							"		<td align='left' valign='bottom' class='MC tdAnEndPed'>" & chr(13) & _
									"<a href='javascript:ocultaInfoAnEnd(" & chr(34) & Trim("" & rs("id")) & chr(34) & ");' title='clique para ocultar os detalhes'>" & _
										"<img id='imgMinusPedAnEnd_" & Trim("" & rs("id")) & "' style='vertical-align:bottom;margin-left:2px;margin-bottom:1px;' src='../imagem/minus.gif' />" & chr(13) & _
									"</a>" & _
										"<span class='Cn'>Indicador</span>" & _
							"		</td>" & chr(13) & _
							"		<td align='left' class='MC'>" & chr(13) & _
										"<span class='Cn'>" & _
										Trim("" & rs("indicador")) & " - " & Trim("" & rs("nome_indicador")) & " ("
						
						s_aux = retorna_so_digitos(Trim("" & rs("cnpj_cpf_indicador")))
						if Len(s_aux) = 11 then
							strInfoAnEnd = strInfoAnEnd & "CPF: " & s_aux & ")"
						else
							strInfoAnEnd = strInfoAnEnd & "CNPJ: " & s_aux & ")"
							end if
						
						strInfoAnEnd = strInfoAnEnd & _
										"</span>" & _
							"		</td>" & chr(13) & _
							"	</tr>" & chr(13) & _
							"	<tr id='TR_INFO_AN_END_LN2_" & Trim("" & rs("id")) & "' class='TR_INFO_AN_END TR_INFO_AN_END_BLOCO_" & Cstr(n_pedido+1) & "'>" & chr(13) & _
							"		<td align='left'>&nbsp;</td>" & chr(13) & _
							"		<td align='left'>" & chr(13)
						
						s_aux = "End. do Indicador: "
						s = formata_endereco(iniciais_em_maiusculas(Trim("" & rs("endereco_logradouro"))), Trim("" & rs("endereco_numero")), Trim("" & rs("endereco_complemento")), iniciais_em_maiusculas(Trim("" & rs("endereco_bairro"))), iniciais_em_maiusculas(Trim("" & rs("endereco_cidade"))), Ucase(Trim("" & rs("endereco_uf"))), retorna_so_digitos(Trim("" & rs("endereco_cep"))))
						strInfoAnEnd = strInfoAnEnd & _
										"<span class='Cni'>" & _
										s_aux & _
										"</span>" & _
										"<span class='Cn'>" & _
										s & _
										"</span>"
										
						strInfoAnEnd = strInfoAnEnd & _
							"		</td>" & chr(13) & _
							"	</tr>" & chr(13)
						
						intQtdePedido = intQtdePedido + 1
						
						rs.MoveNext
						loop
					
				'	VERIFICA SE HÁ COINCIDÊNCIA C/ ENDEREÇO DE OUTROS CLIENTES
					s_sql = "SELECT " & _
								"*" & _
							" FROM t_PEDIDO_ANALISE_ENDERECO" & _
							" WHERE" & _
								" (pedido = '" & Trim("" & r("pedido")) & "')" & _
							" ORDER BY" & _
								" id"
					if rs.State <> 0 then rs.Close
					rs.open s_sql, cn
					if rs.Eof then
						if Not blnAnEnderecoUsaEndParceiro then
							x = x & _
								"				<tr>" & chr(13) & _
								"					<td align='left'>" & chr(13) & _
													"&nbsp;" & _
								"					</td>" & chr(13) & _
								"				</tr>" & chr(13)
							end if
					else
						do while Not rs.Eof
							if intQtdeTotalPedidosAnEndereco > MAX_AN_ENDERECO_QTDE_PEDIDOS_EXIBICAO then exit do
							s_sql = "SELECT" & _
										" tPAEC.*," & _
										" tC.nome_iniciais_em_maiusculas," & _
										" tC.cnpj_cpf" & _
									" FROM t_PEDIDO_ANALISE_ENDERECO_CONFRONTACAO tPAEC" & _
										" LEFT JOIN t_CLIENTE tC ON (tPAEC.id_cliente=tC.id)" & _
									" WHERE" & _
										" (tPAEC.id_pedido_analise_endereco = " & Trim("" & rs("id")) & ")" & _
										" AND (tPAEC.tipo_endereco <> '" & COD_PEDIDO_AN_ENDERECO__END_PARCEIRO & "')" & _
									" ORDER BY" & _
										" tPAEC.id"
							if rs2.State <> 0 then rs2.Close
							rs2.open s_sql, cn
							do while Not rs2.Eof
								intQtdeTotalPedidosAnEndereco = intQtdeTotalPedidosAnEndereco + 1
								if intQtdeTotalPedidosAnEndereco > MAX_AN_ENDERECO_QTDE_PEDIDOS_EXIBICAO then exit do
								intResto = intQtdePedido Mod MAX_PEDIDOS_POR_LINHA_ANALISE_ENDERECO
								if (intQtdePedido = 0) Or (intResto = 0) then
									intQtdePedido = 0
									if intQtdeLinhasPedido > 0 then
										x = x & "				</tr>" & chr(13)
										end if
									x = x & "				<tr>" & chr(13)
									intQtdeLinhasPedido = intQtdeLinhasPedido + 1
									end if
								
								x = x & _
									"					<td align='left' valign='bottom'>" & chr(13) & _
										"<span class='C' id='spanPedidoAnEnd_" & Trim("" & rs2("id")) & "'>" & Trim("" & rs2("pedido")) & "</span>" & _
									"<a id='hrefPedAnEnd_" & Trim("" & rs2("id")) & "' class='hrefAnEndBloco_" & Cstr(n_pedido+1) & "' href='javascript:exibeOcultaInfoAnEnd(" & chr(34) & Trim("" & rs2("id")) & chr(34) & ");' title='clique para exibir mais detalhes'>" & _
										"<img id='imgPlusMinusPedAnEnd_" & Trim("" & rs2("id")) & "' class='imgPlusMinusAnEndBloco_" & Cstr(n_pedido+1) & "' style='vertical-align:bottom;margin-bottom:0px;' src='../imagem/plus.gif' />" & _
									"</a>" & _
									"					</td>" & chr(13)
								
								strInfoAnEnd = strInfoAnEnd & _
									"	<tr id='TR_INFO_AN_END_LN1_" & Trim("" & rs2("id")) & "' class='TR_INFO_AN_END TR_INFO_AN_END_BLOCO_" & Cstr(n_pedido+1) & "'>" & chr(13) & _
									"		<td align='left' valign='bottom' class='MC tdAnEndPed'>" & chr(13) & _
											"<a href='javascript:ocultaInfoAnEnd(" & chr(34) & Trim("" & rs2("id")) & chr(34) & ");' title='clique para ocultar os detalhes'>" & _
												"<img id='imgMinusPedAnEnd_" & Trim("" & rs2("id")) & "' style='vertical-align:bottom;margin-left:2px;margin-bottom:1px;' src='../imagem/minus.gif' />" & chr(13) & _
											"</a>" & _
												"<span class='Cn'>" & Trim("" & rs2("pedido")) & "</span>" & _
									"		</td>" & chr(13) & _
									"		<td align='left' class='MC'>" & chr(13) & _
												"<span class='Cn'>" & _
												Trim("" & rs2("nome_iniciais_em_maiusculas")) & " ("
								
								s_aux = retorna_so_digitos(Trim("" & rs2("cnpj_cpf")))
								if Len(s_aux) = 11 then
									strInfoAnEnd = strInfoAnEnd & "CPF: " & s_aux & ")"
								else
									strInfoAnEnd = strInfoAnEnd & "CNPJ: " & s_aux & ")"
									end if
								
								strInfoAnEnd = strInfoAnEnd & _
												"</span>" & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"	<tr id='TR_INFO_AN_END_LN2_" & Trim("" & rs2("id")) & "' class='TR_INFO_AN_END TR_INFO_AN_END_BLOCO_" & Cstr(n_pedido+1) & "'>" & chr(13) & _
									"		<td align='left'>&nbsp;</td>" & chr(13) & _
									"		<td align='left'>" & chr(13)
								
								if Trim("" & rs2("tipo_endereco")) = COD_PEDIDO_AN_ENDERECO__END_ENTREGA then
									s_aux = "End. Entrega: "
								else
									s_aux = "End. Cadastro: "
									end if
								s = formata_endereco(iniciais_em_maiusculas(Trim("" & rs2("endereco_logradouro"))), Trim("" & rs2("endereco_numero")), Trim("" & rs2("endereco_complemento")), iniciais_em_maiusculas(Trim("" & rs2("endereco_bairro"))), iniciais_em_maiusculas(Trim("" & rs2("endereco_cidade"))), Ucase(Trim("" & rs2("endereco_uf"))), retorna_so_digitos(Trim("" & rs2("endereco_cep"))))
								strInfoAnEnd = strInfoAnEnd & _
												"<span class='Cni'>" & _
												s_aux & _
												"</span>" & _
												"<span class='Cn'>" & _
												s & _
												"</span>"
												
								strInfoAnEnd = strInfoAnEnd & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13)
								
								intQtdePedido = intQtdePedido + 1
								rs2.MoveNext
								loop
							
							if intQtdeTotalPedidosAnEndereco > MAX_AN_ENDERECO_QTDE_PEDIDOS_EXIBICAO then exit do
							rs.MoveNext
							loop
						end if
					
					x = x & _
						"				</tr>" & chr(13) & _
						"			</table>" & chr(13) & _
						"		</td>" & chr(13) & _
						"	</tr>" & chr(13)
					
					if strInfoAnEnd <> "" then
						x = x & _
							"	<tr>" & chr(13) & _
							"		<td class='ME MD' colspan='" & MAX_PEDIDOS_POR_LINHA_ANALISE_ENDERECO & "' align='left'>" & chr(13) & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
										strInfoAnEnd & _
							"			</table>" & chr(13) & _
							"		</td>" & chr(13) & _
							"	</tr>" & chr(13)
						end if
					end if
				
			'	CABEÇALHO C/ TÍTULOS DOS ITENS DO PEDIDO
				x = x & cab
				end if
			
		'	MEMORIZA PEDIDO ATUAL
			if mudou_pedido then
				n_pedido = n_pedido + 1
				mudou_pedido = False
				pedido_a = Trim("" & r("pedido"))
				end if
				
		 ' CONTAGEM
			n_reg = n_reg + 1
			
		'	MEMORIZA CAMPOS
			s_obs1_a = Trim("" & r("obs_1"))
			s_descricao_forma_pagto_a = Trim("" & r("forma_pagto"))
			s_forma_pagto_a = ""
			if Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_A_VISTA then
				s_forma_pagto_a = "À Vista  (" & x_opcao_forma_pagamento(r("av_forma_pagto")) & ")"
			elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELA_UNICA then
				s_forma_pagto_a = "Parcela Única:  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pu_valor")) & "  (" & x_opcao_forma_pagamento(r("pu_forma_pagto")) & ")  vencendo após " & Cstr(r("pu_vencto_apos")) & " dias"
			elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_CARTAO then
				s_forma_pagto_a = "Parcelado no Cartão (internet) em " & Cstr(r("pc_qtde_parcelas")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pc_valor_parcela"))
			elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
				s_forma_pagto_a = "Parcelado no Cartão (maquineta) em " & Cstr(r("pc_maquineta_qtde_parcelas")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pc_maquineta_valor_parcela"))
			elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
				s_forma_pagto_a = "Entrada:  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pce_entrada_valor")) & "  (" & x_opcao_forma_pagamento(r("pce_forma_pagto_entrada")) & ")" & _
								  chr(13) & "<br>" & chr(13) & _
								  "Prestações:  " & Cstr(r("pce_prestacao_qtde")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pce_prestacao_valor")) & _
								  "  (" & x_opcao_forma_pagamento(r("pce_forma_pagto_prestacao")) & ")  vencendo a cada " & _
								  Cstr(r("pce_prestacao_periodo")) & " dias"
			elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
				s_forma_pagto_a = "1ª Prestação:  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pse_prim_prest_valor")) & "  (" & x_opcao_forma_pagamento(r("pse_forma_pagto_prim_prest")) & ")  vencendo após " & Cstr(r("pse_prim_prest_apos")) & " dias" & _
								  chr(13) & "<br>" & chr(13) & _
								  "Demais Prestações:  " & Cstr(r("pse_demais_prest_qtde")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pse_demais_prest_valor")) & _
								  "  (" & x_opcao_forma_pagamento(r("pse_forma_pagto_demais_prest")) & ")  vencendo a cada " & _
								  Cstr(r("pse_demais_prest_periodo")) & " dias"
				end if

			
			s_sql = "SELECT" & _
						" fabricante," & _
						" produto," & _
						" SUM(qtde) AS qtde," & _
						" SUM(qtde*preco_venda) AS total_preco_venda," & _
						" SUM(qtde*preco_NF) AS total_preco_NF" & _
					" FROM t_PEDIDO_ITEM tPI" & _
						" INNER JOIN t_PEDIDO tP ON (tPI.pedido = tP.pedido)" & _
					" WHERE" & _
						" (tPI.pedido LIKE '" & Trim("" & r("pedido_base")) & BD_CURINGA_TODOS & "')" & _
						" AND (tP.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
					" GROUP BY" & _
						" fabricante," & _
						" produto" & _
					" ORDER BY" & _
						" fabricante," & _
						" produto"
			if tPI.State <> 0 then tPI.Close
			tPI.Open s_sql, cn

			do while Not tPI.Eof
				s_sql = "SELECT TOP 1" & _
							" preco_lista," & _
							" descricao_html" & _
						" FROM t_PEDIDO_ITEM" & _
						" WHERE" & _
							" (pedido LIKE '" & Trim("" & r("pedido_base")) & BD_CURINGA_TODOS & "')" & _
							" AND (fabricante = '" & Trim("" & tPI("fabricante")) & "')" & _
							" AND (produto = '" & Trim("" & tPI("produto")) & "')"
				if tPI2.State <> 0 then tPI2.Close
				tPI2.Open s_sql, cn

				x = x & "	<tr>" & chr(13)

			'> FABRICANTE
				if Not tPI.Eof then
					s = Trim("" & tPI("fabricante"))
				else
					s = ""
					end if
				if s = "" then s = "&nbsp;"
				x = x & "		<td align='left' valign='bottom' class='MDTE'>" & chr(13) & _
						"			<span class='Cn'>" & s & "</span>" & chr(13) & _
						"		</td>" & chr(13)

			'> PRODUTO
				if Not tPI.Eof then
					s = Trim("" & tPI("produto"))
				else
					s = ""
					end if
				if s = "" then s = "&nbsp;"
				x = x & "		<td align='left' valign='bottom' class='MTD'>" & chr(13) & _
						"			<span class='Cn'>" & s & "</span>" & chr(13) & _
						"		</td>" & chr(13)

			'> DESCRIÇÃO DO PRODUTO
				if Not tPI2.Eof then
					s = produto_formata_descricao_em_html(Trim("" & tPI2("descricao_html")))
				else
					s = ""
					end if
				if s = "" then s = "&nbsp;"
				x = x & "		<td align='left' valign='bottom' class='MTD'>" & chr(13) & _
						"			<span class='Cn'>" & s & "</span>" & chr(13) & _
						"		</td>" & chr(13)

			'> QUANTIDADE
				if Not tPI.Eof then
					s = formata_inteiro(tPI("qtde"))
				else
					s = ""
					end if
				if s = "" then s = "&nbsp;"
				x = x & "		<td align='right' valign='bottom' class='MTD'>" & chr(13) & _
						"			<span class='Cnd'>" & s & "</span>" & chr(13) & _
						"		</td>" & chr(13)

			'> PREÇO DE LISTA
				if Not tPI2.Eof then
					vl_lista=tPI2("preco_lista")
					s = formata_moeda(vl_lista)
				else
					vl_lista=0
					s = ""
					end if
				if s = "" then s = "&nbsp;"
				x = x & "		<td align='right' valign='bottom' class='MTD'>" & chr(13) & _
						"			<span class='Cnd'>" & s & "</span>" & chr(13) & _
						"		</td>" & chr(13)

			'> DESCONTO
				if Not tPI.Eof then
					vl_unitario=tPI("total_preco_NF") / tPI("qtde")
				else
					vl_unitario=0
					end if
			
				if vl_unitario < vl_lista then
					if vl_lista > 0 then
						perc_desc=100*((vl_lista-vl_unitario)/vl_lista)
					else
						perc_desc=0
					end if
					s_desc = formata_perc(perc_desc)
				else
					perc_desc=0
					s_desc=""
					end if
			
				s = s_desc
				if s = "" then s = "&nbsp;"
				x = x & "		<td align='right' valign='bottom' class='MTD'>" & chr(13) & _
						"			<span class='Cnd'>" & s & "</span>" & chr(13) & _
						"		</td>" & chr(13)

			'> VALOR UNITÁRIO
				s = formata_moeda(vl_unitario)
				if s = "" then s = "&nbsp;"
				x = x & "		<td align='right' valign='bottom' class='MTD'>" & chr(13) &_
						"			<span class='Cnd'>" & s & "</span>" & chr(13) &_
						"		</td>" & chr(13)

			'> VALOR TOTAL DO ITEM
				if Not tPI.Eof then
					vl_item = tPI("qtde") * vl_unitario
					s = formata_moeda(vl_item)
				else
					vl_item = 0
					s = ""
					end if
				if s = "" then s = "&nbsp;"
				x = x & "		<td align='right' valign='bottom' class='MTD'>" & chr(13) &_
						"			<span class='Cnd'>" & s & "</span>" & chr(13) & _
						"		</td>" & chr(13)

			'> TOTALIZAÇÃO DE VALORES
				vl_total_item = vl_total_item + vl_item

				x = x & "	</tr>" & chr(13)

				tPI.MoveNext
				loop
			
			if Not tem_filtro_pedido then
				if n_pedido >= MAX_PEDIDOS then exit do
				end if
			
			end if  ' if Not blnPulaRegistro
		
		r.MoveNext
		loop
		
		
  ' FECHA O ÚLTIMO PEDIDO
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO PEDIDO
		x = x & "	<tr>" & chr(13) & _
				"		<td colspan='8' style='border-top:1px solid #C0C0C0;' align='left'>" & chr(13)& _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<tr>" & chr(13) & _
				"					<td width='20%' class='MDE' align='left' valign='bottom'><span class='Rf'>Status de Pagto</span></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right' valign='bottom'><span class='Rf'>VL Total&nbsp;&nbsp;(Família)&nbsp;</span></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right' valign='bottom'><span class='Rf'>VL Pago&nbsp;</span></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right' valign='bottom'><span class='Rf'>VL Devoluções&nbsp;</span></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right' valign='bottom'><span class='Rf'>Total dos Itens&nbsp;</span></td>" & chr(13) & _
				"				</tr>" & chr(13) & _
				"				<tr>" & chr(13) & _
				"					<td width='20%' class='MDE' align='left'><span class='C'>" & Ucase(x_status_pagto(st_pagto)) & "&nbsp;</span></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right'><span class='Cd'>" & formata_moeda(vl_TotalFamiliaPrecoNF) & "</span></td>" & chr(13)
				
		if vl_TotalFamiliaPago >= 0 then s_cor = "black" else s_cor = "red"
		x = x & "					<td width='20%' class='MD' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_TotalFamiliaPago) & "</span></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right'><span class='Cd'>" & formata_moeda(vl_TotalFamiliaDevolucaoPrecoNF) & "</span></td>" & chr(13) & _
				"					<td width='20%' class='MD' align='right'><span class='Cd'>" & formata_moeda(vl_total_item) & "</span></td>" & chr(13) & _
				"				</tr>" & chr(13) 

		x = x & "			</table>" & chr(13) & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13)
		
	'	OBS
		x = x & "	<tr>" & chr(13) & _
								"		<td class='MDTE' colspan='8' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
								"			<span class='PLTc' style='font-weight:bold;'>Observações I</span>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13)
		x = x & "	<tr>" & chr(13) & _
								"		<td class='MDTE' colspan='8' align='left' valign='bottom'>" & chr(13) & _
								"			<textarea name='c_obs1' id='c_obs1' readonly tabindex=-1 class='PLLe' rows='" & Cstr(NUM_LINHAS_OBS) & "' style='width:642px;margin-left:2pt;' onkeypress='limita_tamanho(this,MAX_TAM_OBS1);' onblur='this.value=trim(this.value);'>" & _
								s_obs1_a & "</textarea>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13)
		
	'	FORMA DE PAGAMENTO
		s = s_forma_pagto_a
		if s = "" then s = "&nbsp;"
		x = x & "	<tr>" & chr(13) & _
								"		<td class='MDTE' colspan='8' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
								"			<span class='PLTc' style='font-weight:bold;'>Forma de Pagamento</span>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13)
		x = x & "	<tr>" & chr(13) & _
								"		<td class='MDTE' colspan='8' align='left' valign='bottom' style='padding-left:2px;'>" & chr(13) & _
								"			<span class='Cn'>" & s & "</span>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13)
			
	'	DESCRIÇÃO DA FORMA DE PAGAMENTO
		x = x & "	<tr>" & chr(13) & _
								"		<td class='MDTE' colspan='8' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
								"			<span class='PLTc' style='font-weight:bold;'>Descrição da Forma de Pagamento</span>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13)
		x = x & "	<tr>" & chr(13) & _
								"		<td class='MDTE' colspan='8' align='left' valign='bottom' style='width:642px;'>" & chr(13) & _
								"			<textarea name='c_descr_forma_pagto' id='c_descr_forma_pagto' class='PLLe' rows='" & Cstr(NUM_LINHAS_DESCR_FORMA_PAGTO) & "' style='width:642px;margin-left:2pt;' onkeypress='limita_tamanho(this,MAX_TAM_FORMA_PAGTO);' onblur='this.value=trim(this.value);'>" & _
								s_descricao_forma_pagto_a & "</textarea>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13)
		
	'	ANÁLISE DE CRÉDITO
		x = x & "	<tr>" & chr(13) & _
								"		<td class='MDTE' colspan='8' align='center' valign='bottom' style='background:mintcream;'>" & chr(13) & _
								"			<span class='PLTc' style='font-weight:bold;'>Análise de Crédito</span>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13)
		x = x & "	<tr>" & chr(13) & _
								"		<td colspan='8' align='left' valign='bottom'>" & chr(13) & _
								"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
								"				<tr>" & chr(13) & _
								"					<td class='MC ME' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
								"					<td class='MC' align='left'>" & chr(13) & _
								"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_PENDENTE_VENDAS & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:red;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[0].click();'>" & x_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS) & "</span>" & chr(13) &_
								"					</td>" & chr(13) & _
								"					<td class='MC' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
								"					<td class='MC' align='left'>" & chr(13) & _
								"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_PENDENTE_ENDERECO & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:red;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[1].click();'>" & x_analise_credito(COD_AN_CREDITO_PENDENTE_ENDERECO) & "</span>" & chr(13) &_
								"					</td>" & chr(13) & _
								"					<td class='MC' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
								"					<td class='MC' align='left'>" & chr(13) & _
								"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_PENDENTE & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:red;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[2].click();'>" & x_analise_credito(COD_AN_CREDITO_PENDENTE) & "</span>" & chr(13) &_
								"					</td>" & chr(13) & _
								"					<td class='MC' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
								"					<td class='MC' align='right'>" & chr(13) & _
								"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_OK & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:green;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[3].click();'>" & x_analise_credito(COD_AN_CREDITO_OK) & "</span>" & chr(13) &_
								"					</td>" & chr(13) & _
								"					<td class='MC' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
								"				</tr>" & chr(13) & _
                                "               <tr id='trPendVendasMotivo_" & Cstr(n_pedido) & "' style='display:none;'>" & chr(13) & _
                                "                   <td class='MC ME' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
                                "                   <td class='MC' align='left' colspan='8'>" & chr(13) & _
                                "                       <span class='C'>Motivo: </span>" & chr(13) & _
                                "                       <select name='c_pendente_vendas_motivo_" & Cstr(n_pedido) & "' id='c_pendente_vendas_motivo_" & Cstr(n_pedido) & "'>" & chr(13) & _
                                                            codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__AC_PENDENTE_VENDAS_MOTIVO, Null) & chr(13) & _
                                "                       </select>" & chr(13) & _
                                "                   </td>" & chr(13) & _
                                "               </tr>" & _
								"				<tr>" & chr(13) & _
								"					<td class='MC MB' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
								"					<td class='MC MB' align='left'>" & chr(13) & _
								"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:darkorange;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[4].click();'>" & x_analise_credito(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO) & "</span>" & chr(13) &_
								"					</td>" & chr(13) & _
								"					<td class='MC MB' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
								"					<td class='MC MB' colspan='5' align='left'>" & chr(13) & _
								"						<input type='radio' id='rb_credito_ped_" & Cstr(n_pedido) & "' name='rb_credito_ped_" & Cstr(n_pedido) & "' value='" & COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO & "' onchange='exibeOcultaPendenteVendasMotivo(" & Cstr(n_pedido) & ")'><span class='C' style='cursor:default;color:darkorange;' onclick='fREL.rb_credito_ped_" & Cstr(n_pedido) & "[5].click();'>" & x_analise_credito(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO) & "</span>" & chr(13) &_
								"					</td>" & chr(13) & _
								"					<td class='MC MB MD' style='width:" & Cstr(intEspacamentoOpcoesAnaliseCredito) & "px;' align='left'>&nbsp;</td>" & chr(13) & _
								"				</tr>" & chr(13) & _
								"			</table>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13)
		end if


  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table 
		x = x & "	<tr>" & chr(13) & _
				"		<td class='MT ALERTA' colspan='8' style='width:647px;' align='center'><span class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if


  ' FECHA TABELA
	x = x & "</table>" & chr(13)

    ' QTDE PEDIDOS
    x = x & "<input type='hidden' name='c_qtde_total_pedidos' id='c_qtde_total_pedidos' value='" & Cstr(n_pedido) & "'>"
	
	Response.write x

	if tP.State <> 0 then tP.Close
	set tP=nothing

	if tPI.State <> 0 then tPI.Close
	set tPI=nothing

	if tPI2.State <> 0 then tPI2.Close
	set tPI2=nothing

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

<script language="JavaScript" type="text/javascript">
    $(document).ready(function () {
		$(".TR_INFO_AN_END").hide().addClass("TR_INFO_AN_END_HIDDEN");
		$(".TIT_INFO_AN_END_BLOCO").addClass("TR_INFO_AN_END_HIDDEN");
	});
</script>

<script language="JavaScript" type="text/javascript">
function ocultaInfoAnEnd(id_row) {
var s_id_ln1, s_id_ln2, s_id_img, s_id_href;
	s_id_ln1 = "#TR_INFO_AN_END_LN1_" + id_row;
	s_id_ln2 = "#TR_INFO_AN_END_LN2_" + id_row;
	s_id_img = "#imgPlusMinusPedAnEnd_" + id_row;
	s_id_href = "#hrefPedAnEnd_" + id_row;
	$(s_id_ln1).hide();
	$(s_id_ln1).addClass("TR_INFO_AN_END_HIDDEN");
	$(s_id_ln2).hide();
	$(s_id_ln2).addClass("TR_INFO_AN_END_HIDDEN");
	$(s_id_img).attr({ src: '../imagem/plus.gif' });
	$(s_id_href).attr({ title: 'clique para exibir mais detalhes' });
}

function exibeOcultaInfoAnEnd(id_row) {
var s_id_ln1, s_id_ln2, s_id_img, s_id_href;
	s_id_ln1 = "#TR_INFO_AN_END_LN1_" + id_row;
	s_id_ln2 = "#TR_INFO_AN_END_LN2_" + id_row;
	s_id_img = "#imgPlusMinusPedAnEnd_" + id_row;
	s_id_href = "#hrefPedAnEnd_" + id_row;
	if ($(s_id_ln1).hasClass("TR_INFO_AN_END_HIDDEN")) {
		$(s_id_ln1).show();
		$(s_id_ln1).removeClass("TR_INFO_AN_END_HIDDEN");
		$(s_id_ln2).show();
		$(s_id_ln2).removeClass("TR_INFO_AN_END_HIDDEN");
		$(s_id_img).attr({ src: '../imagem/minus.gif' });
		$(s_id_href).attr({title: 'clique para ocultar os detalhes' });
	}
	else {
		$(s_id_ln1).hide();
		$(s_id_ln1).addClass("TR_INFO_AN_END_HIDDEN");
		$(s_id_ln2).hide();
		$(s_id_ln2).addClass("TR_INFO_AN_END_HIDDEN");
		$(s_id_img).attr({ src: '../imagem/plus.gif' });
		$(s_id_href).attr({ title: 'clique para exibir mais detalhes' });
	}
}

function exibeOcultaTodosInfoAnEnd(indice_bloco) {
var s_tit_id_img, s_tit_id_href, s_tit_id_span;
var s_item_img_classe, s_item_href_classe;
var s_classe;
	s_classe = ".TR_INFO_AN_END_BLOCO_" + indice_bloco;
	s_tit_id_img = "#imgPlusMinusTitAnEnd_" + indice_bloco;
	s_tit_id_href = "#hrefTitAnEnd_" + indice_bloco;
	s_tit_id_span = "#spanTitAnEnd_" + indice_bloco;
	s_item_img_classe = ".imgPlusMinusAnEndBloco_" + indice_bloco;
	s_item_href_classe = ".hrefAnEndBloco_" + indice_bloco;
	if ($(s_tit_id_span).hasClass("TR_INFO_AN_END_HIDDEN")) {
		$(s_classe).show();
		$(s_classe).removeClass("TR_INFO_AN_END_HIDDEN");
		$(s_tit_id_span).removeClass("TR_INFO_AN_END_HIDDEN");
		$(s_tit_id_img).attr({ src: '../imagem/minus.gif' });
		$(s_tit_id_href).attr({ title: 'clique para ocultar os detalhes' });
		$(s_item_img_classe).attr({ src: '../imagem/minus.gif' });
		$(s_item_href_classe).attr({ title: 'clique para ocultar os detalhes' });
	}
	else {
		$(s_classe).hide();
		$(s_classe).addClass("TR_INFO_AN_END_HIDDEN");
		$(s_tit_id_span).addClass("TR_INFO_AN_END_HIDDEN");
		$(s_tit_id_img).attr({ src: '../imagem/plus.gif' });
		$(s_tit_id_href).attr({ title: 'clique para exibir mais detalhes' });
		$(s_item_img_classe).attr({ src: '../imagem/plus.gif' });
		$(s_item_href_classe).attr({ title: 'clique para exibir mais detalhes' });
	}
}

function exibeOcultaPendenteVendasMotivo(idx_bloco) {
    if ($('#rb_credito_ped_' + idx_bloco).is(':checked')) {
        $('#trPendVendasMotivo_' + idx_bloco).show();
    }
    else {
        $('#trPendVendasMotivo_' + idx_bloco).hide();
    }
}

function fRELConfirma(f) {
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

<style type="text/css">
.Cni{
	font-family: Arial, Helvetica, sans-serif;
	color: #808080;
	font-size: 8pt;
	font-style: italic;
	font-weight: bold;
	margin: 0pt 2pt 1pt 2pt;
}
.tdAnEndPed
{
	width:80px;
}
.PLTdi{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 8pt;
	color: #696969;
	font-weight: normal;
	font-style:italic;
	letter-spacing:normal;
	text-align:right;
	vertical-align: middle;
	margin-right:2pt;
	cursor:default;
	}
</style>


<body onload="focus();" link="#ffffff" alink="#ffffff" vlink="#ffffff">
<center>

<form id="fREL" name="fREL" method="post" action="RelAnaliseCreditoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Análise de Crédito</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!--  RELATÓRIO  -->
<br>

<input type="hidden" name="opcao_filtro_pedido" id="opcao_filtro_pedido" value="<%if tem_filtro_pedido then Response.Write "S" else Response.Write "N"%>">
<input type="hidden" name="c_lista_loja" id="c_lista_loja" value="<%=c_lista_loja%>">
<input type="hidden" name="c_valor_inferior" id="c_valor_inferior" value="<%=c_valor_inferior%>">
<input type="hidden" name="c_valor_superior" id="c_valor_superior" value="<%=c_valor_superior%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">

<!-- FORÇA A CRIAÇÃO DE UM ARRAY MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" name="c_pedido" id="c_pedido" value="">
<input type="hidden" name="c_obs1" id="c_obs1" value="">
<input type="hidden" name="c_descr_forma_pagto" id="c_descr_forma_pagto" value="">

<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align='left'>&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" 
		<% if origem="A" then %>
			href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>"
		<% else %>
			href="javascript:history.back()"
		<% end if %>
	title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELConfirma(fREL)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
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
