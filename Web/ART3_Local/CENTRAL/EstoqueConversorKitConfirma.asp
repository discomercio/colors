<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  E S T O Q U E C O N V E R S O R K I T C O N F I R M A . A S P
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

	dim s, s_aux, s_log, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	if Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim v_item
	dim i, j, n, s_kit, s_kit_fabricante, s_documento, n_kit_qtde
	dim c_ncm, c_cst
	dim v_kit, s_id_estoque_temp, i_kit, i_prod, preco_fabricante_kit, vl_custo2_kit, preco_fabricante_aux, vl_custo2_aux
	dim vl_BC_ICMS_ST_kit, vl_ICMS_ST_kit, vl_BC_ICMS_ST_aux, vl_ICMS_ST_aux
	dim s_id_estoque_lote, s_id_estoque
	dim c_nfe_emitente, id_nfe_emitente

	dim alerta
	alerta = ""
	
'	OBTÉM DADOS DIGITADOS NO FORMULÁRIO
	s_kit_fabricante = normaliza_codigo(retorna_so_digitos(Request.Form("c_kit_fabricante")), TAM_MIN_FABRICANTE)
	s_documento = Trim(Request.Form("c_documento"))
	s_kit = Ucase(Trim(Request.Form("c_kit")))
	s = Trim(Request.Form("c_kit_qtde"))
	if IsNumeric(s) then n_kit_qtde = CLng(s) else n_kit_qtde = 0
	c_ncm = retorna_so_digitos(Trim(Request.Form("c_ncm")))
	c_cst = retorna_so_digitos(Trim(Request.Form("c_cst")))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
	id_nfe_emitente = converte_numero(c_nfe_emitente)
	
	redim v_item(0)
	set v_item(0) = New cl_ITEM_PEDIDO
	n = Request.Form("c_produto").Count
	for i = 1 to n
		s=Trim(Request.Form("c_produto")(i))
		if s <> "" then
			if Trim(v_item(ubound(v_item)).produto) <> "" then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ITEM_PEDIDO
				end if
			with v_item(ubound(v_item))
				s = retorna_so_digitos(Request.Form("c_fabricante")(i))
				s = normaliza_codigo(s, TAM_MIN_FABRICANTE)
				.fabricante = s
				.produto=Ucase(Trim(Request.Form("c_produto")(i)))
				s = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
				end with
			end if
		next
	
'	VERIFICA DISPONIBILIDADE NO ESTOQUE
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
			'	QUANTIDADE DE PRODUTOS A SAIR DO ESTOQUE = QUANTIDADE DE KITS x QTDE DE PRODUTOS POR KIT
				n = n_kit_qtde * .qtde
				s = "SELECT" & _
						" SUM(qtde-qtde_utilizada) AS total" & _
					" FROM t_ESTOQUE tE INNER JOIN t_ESTOQUE_ITEM tEI ON (tE.id_estoque = tEI.id_estoque)" & _
					" WHERE" & _
						" (tE.id_nfe_emitente = " & id_nfe_emitente & ")" & _
						" AND (tEI.fabricante = '" & Trim(.fabricante) & "')" & _
						" AND (tEI.produto = '" & Trim(.produto) & "')" & _
						" AND ((qtde-qtde_utilizada) > 0)"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				j=0
				if Not rs.Eof then 
					if Not IsNull(rs("total")) then j = CLng(rs("total"))
					end if
				if n > j then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Faltam " & CStr(n-j) & " unidades no estoque do produto " & .produto & " do fabricante " & .fabricante & " (estoque: " & obtem_apelido_empresa_NFe_emitente(c_nfe_emitente) & ")."
					end if
				end with
			next
		end if

'	VERIFICA SE A OPERAÇÃO ESTÁ SENDO FEITA EM DUPLICIDADE
	if alerta = "" then
		s = "SELECT" & _
				" t_ESTOQUE.id_estoque, data_entrada, hora_entrada, t_ESTOQUE_ITEM.fabricante, produto, qtde" & _
			" FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
			" WHERE" & _
				" (t_ESTOQUE.id_nfe_emitente = " & c_nfe_emitente & ")" & _
				" AND (data_entrada=" & bd_formata_data(Date) & ")" & _
				" AND (hora_entrada >='" & formata_hora_hhnnss(Now-converte_min_to_dec(20))& "')" & _
				" AND (t_ESTOQUE_ITEM.fabricante='" & s_kit_fabricante & "')" & _
				" AND (produto='" & s_kit & "')" & _
				" AND (qtde=" & Cstr(n_kit_qtde) & ")" & _
				" AND (usuario='" & usuario & "')" & _
				" AND (documento='" & s_documento & "')" & _
				" AND (kit<>0)"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if Not rs.Eof then
			alerta = "Esta operação de conversão do kit " & s_kit & " já foi realizada às " & formata_hhnnss_para_hh_nn_ss(Trim("" & rs("hora_entrada"))) & "."
			end if
		end if
	
	if alerta = "" then
	'	INICIA GERAÇÃO DOS KITS
	'	=======================
	'	SUBSÍDIOS: O PREÇO FABRICANTE DO KIT É A SOMA DO PREÇO FABRICANTE DE CADA PRODUTO
	'		USADO NA SUA COMPOSIÇÃO.  ISSO SIGNIFICA QUE CADA UNIDADE DO KIT PODE TER UM
	'		PREÇO FABRICANTE DIFERENTE SE OS PRODUTOS USADOS NA COMPOSIÇÃO PERTENCEREM A
	'		DIFERENTES LOTES DE ENTRADA DO ESTOQUE.
	'		PARA CALCULAR C/ EXATIDÃO O PREÇO FABRICANTE DE CADA UNIDADE DO KIT, A ROTINA
	'		PROCESSA CADA UNIDADE DO KIT SEPARADAMENTE.  OU SEJA, SE FOREM CADASTRADOS 50
	'		KITS, ENTÃO A ROTINA IRÁ EXECUTAR 50 VEZES A OPERAÇÃO DE SAÍDA DO ESTOQUE DOS
	'		PRODUTOS USADOS NA COMPOSIÇÃO DO KIT.  DURANTE O PROCESSAMENTO DA SAÍDA DO
	'		ESTOQUE, É FEITO O CÁLCULO DA FORMAÇÃO DO PREÇO FABRICANTE DE CADA UNIDADE DO
	'		KIT.  EM SEGUIDA, AS UNIDADES DO KIT SÃO AGRUPADAS SE ELAS POSSUÍREM O MESMO
	'		VALOR DE PREÇO FABRICANTE E SE ELAS ESTIVEREM CONSISTENTES C/ A POLÍTICA FIFO
	'		DO ESTOQUE.
	'		PORTANTO, COMO A ANÁLISE DO AGRUPAMENTO DAS UNIDADES DO KIT É FEITA SOMENTE
	'		DEPOIS DA SAÍDA DO ESTOQUE DOS PRODUTOS QUE OS COMPÕEM, ENTÃO É INFORMADO
	'		UM Nº ID_ESTOQUE TEMPORÁRIO P/ SER REGISTRADO NO MOVIMENTO DO ESTOQUE
	'		(T_ESTOQUE_MOVIMENTO) DURANTE A SAÍDA DOS PRODUTOS.
	'		APÓS A ANÁLISE DE AGRUPAMENTO, SE A UNIDADE DO KIT FOR AGRUPADA COM OUTRAS
	'		UNIDADES, ENTÃO O Nº TEMPORÁRIO É TROCADO PELO Nº USADO POR ESSAS OUTRAS
	'		UNIDADES.  CASO CONTRÁRIO, O Nº TEMPORÁRIO É TROCADO PELO Nº DO NOVO LOTE
	'		DE ENTRADA NO ESTOQUE.

		redim v_kit(0)
		set v_kit(Ubound(v_kit)) = New cl_AGRUPA_KIT_POR_PRECO
		with v_kit(Ubound(v_kit))
			.id_estoque = ""
			.qtde = 0
			.preco_fabricante = 0
			.vl_custo2 = 0
			.vl_BC_ICMS_ST = 0
			.vl_ICMS_ST = 0
			end with

	'	GERA CHAVE TEMPORÁRIA
		if Not gera_id_estoque_temp(s_id_estoque_temp, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)

		if rs.State <> 0 then rs.Close
		set rs = nothing

	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

	'	PROCESSA A GERAÇÃO DE KITS UNIDADE POR UNIDADE.
	'	ISSO É FEITO P/ CALCULAR O VALOR DO PREÇO FABRICANTE DO KIT,
	'	QUE É A SOMA DO PREÇO FABRICANTE DOS PRODUTOS QUE COMPÕEM O KIT.
		for i_kit = 1 To n_kit_qtde
			
			if alerta <> "" then exit for
			
			preco_fabricante_kit = 0
			vl_custo2_kit = 0
			vl_BC_ICMS_ST_kit = 0
			vl_ICMS_ST_kit = 0
		'	PARA CADA UNIDADE DO KIT, FAZ A SAÍDA DOS PRODUTOS QUE O COMPÕEM
			for i_prod = Lbound(v_item) to Ubound(v_item)
				with v_item(i_prod)
					if (Trim(.produto)<>"") And (.qtde > 0) then
						if Not estoque_produto_saida_para_kit_v2(usuario, s_id_estoque_temp, id_nfe_emitente, .fabricante, .produto, .qtde, preco_fabricante_aux, vl_custo2_aux, vl_BC_ICMS_ST_aux, vl_ICMS_ST_aux, msg_erro) then
						'	~~~~~~~~~~~~~~~~
							cn.RollbackTrans
						'	~~~~~~~~~~~~~~~~
							Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
							end if
						
						preco_fabricante_kit = preco_fabricante_kit + preco_fabricante_aux
						vl_custo2_kit = vl_custo2_kit + vl_custo2_aux
						vl_BC_ICMS_ST_kit = vl_BC_ICMS_ST_kit + vl_BC_ICMS_ST_aux
						vl_ICMS_ST_kit = vl_ICMS_ST_kit + vl_ICMS_ST_aux
						end if
					end with
				next
	
		'	SE A COMPOSIÇÃO DESTA UNIDADE GEROU O MESMO PREÇO FABRICANTE QUE A UNIDADE ANTERIOR,
		'	ENTÃO INCREMENTA O ÚLTIMO LOTE.  CASO CONTRÁRIO, INICIA UM NOVO LOTE.
		'	PARA MANTER CONSISTENTE A POLÍTICA FIFO DE ESTOQUE DOS PRODUTOS QUE COMPÕEM O KIT,
		'	NÃO SE DEVE AGRUPAR A LOTES ANTERIORES AO ÚLTIMO (LEMBRE-SE DE QUE OS PRODUTOS QUE
		'	COMPÕEM O KIT SOMENTE IRÃO GERAR FATURAMENTO QUANDO O PRÓPRIO KIT FOR VENDIDO).
			s_id_estoque_lote = ""
			if (Trim(v_kit(Ubound(v_kit)).id_estoque)<>"") And (v_kit(Ubound(v_kit)).qtde > 0) then
				if (v_kit(Ubound(v_kit)).preco_fabricante = preco_fabricante_kit) And _
					(v_kit(Ubound(v_kit)).vl_custo2 = vl_custo2_kit) And _
					(v_kit(Ubound(v_kit)).vl_BC_ICMS_ST = vl_BC_ICMS_ST_kit) And _
					(v_kit(Ubound(v_kit)).vl_ICMS_ST = vl_ICMS_ST_kit) then
					s_id_estoque_lote = Trim(v_kit(Ubound(v_kit)).id_estoque)
				'	AGRUPA ESTA UNIDADE DO KIT
					v_kit(Ubound(v_kit)).qtde = v_kit(Ubound(v_kit)).qtde + 1
					end if
				end if

		'	DEVE AGRUPAR C/ OUTRAS UNIDADES DO KIT!!
			If s_id_estoque_lote <> "" Then
				s = "SELECT " & _
						"*" & _
					" FROM t_ESTOQUE_ITEM" & _
					" WHERE" & _
						" (id_estoque='" & s_id_estoque_lote & "')" & _
						" AND (fabricante='" & s_kit_fabricante & "')" & _
						" AND (produto='" & Trim(s_kit) & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "A operação falhou porque o registro do kit no estoque não foi encontrado!!"
				else
				'	AGRUPA ESTA UNIDADE DO KIT!!
					rs("qtde") = rs("qtde") + 1
					rs.Update
					rs.Close
					end if

		'	CRIA NOVO LOTE NO ESTOQUE P/ ESTA UNIDADE DO KIT!!
			Else
			'	GERA A CHAVE P/ A NOVA ENTRADA NO ESTOQUE
				If Not gera_id_estoque(s_id_estoque, msg_erro) Then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
					end if
				
				s_id_estoque_lote = s_id_estoque
				
			'	PREPARA NOVA ENTRADA NO VETOR?
				If (Trim(v_kit(UBound(v_kit)).id_estoque) <> "") Then
					ReDim Preserve v_kit(UBound(v_kit)+1)
					set v_kit(Ubound(v_kit)) = New cl_AGRUPA_KIT_POR_PRECO
					With v_kit(UBound(v_kit))
						.id_estoque = ""
						.qtde = 0
						.preco_fabricante = 0
						.vl_custo2 = 0
						.vl_BC_ICMS_ST = 0
						.vl_ICMS_ST = 0
						End With
					End If
				
			'	INCLUI ESTA UNIDADE DO KIT NO VETOR
				With v_kit(UBound(v_kit))
					.id_estoque = s_id_estoque_lote
					.qtde = 1
					.preco_fabricante = preco_fabricante_kit
					.vl_custo2 = vl_custo2_kit
					.vl_BC_ICMS_ST = vl_BC_ICMS_ST_kit
					.vl_ICMS_ST = vl_ICMS_ST_kit
					End With

				s = "INSERT INTO t_ESTOQUE (" & _
						"id_estoque, data_entrada, hora_entrada, id_nfe_emitente, fabricante, documento," & _
						" usuario, data_ult_movimento, kit" & _
					") VALUES (" & _
						"'" & s_id_estoque_lote & "'" & _
						"," & bd_formata_data(Date) & _
						",'" & retorna_so_digitos(formata_hora(Now)) & "'" & _
						", " & Cstr(id_nfe_emitente) & _
						",'" & s_kit_fabricante & "'" & _
						",'" & s_documento & "'" & _
						",'" & usuario & "'" & _
						"," & bd_formata_data(Date) & _
						", " & "1" & _
					")"
				cn.Execute(s)
				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
					end if

			'	GRAVA INFORMAÇÕES DO KIT NO ESTOQUE
				s = "INSERT INTO T_ESTOQUE_ITEM (" & _
						"id_estoque, fabricante, produto, qtde, preco_fabricante, vl_custo2," & _
						" vl_BC_ICMS_ST, vl_ICMS_ST," & _
						" ncm, cst," & _
						" data_ult_movimento, sequencia" & _
					") VALUES (" & _
						"'" & s_id_estoque_lote & "'" & _
						",'" & s_kit_fabricante & "'" & _
						",'" & s_kit & "'" & _
						"," & "1" & _
						"," & bd_formata_numero(preco_fabricante_kit) & _
						"," & bd_formata_numero(vl_custo2_kit) & _
						"," & bd_formata_numero(vl_BC_ICMS_ST_kit) & _
						"," & bd_formata_numero(vl_ICMS_ST_kit) & _
						",'" & c_ncm & "'" & _
						",'" & c_cst & "'" & _
						"," & bd_formata_data(Date) & _
						"," & "1" & _
					")"
				cn.Execute(s)
				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
					end if
				End If


		'	SUBSTITUI O Nº ID_ESTOQUE TEMPORÁRIO PELO DEFINITIVO
			s = "UPDATE t_ESTOQUE_MOVIMENTO SET" & _
					" kit_id_estoque='" & Trim(s_id_estoque_lote) & "'" & _
				" WHERE" & _
					" (kit_id_estoque='" & Trim(s_id_estoque_temp) & "')"
			cn.Execute(s)
			if Err <> 0 then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
				end if
			
			Next  ' PRÓXIMA UNIDADE DO KIT


		'Log de movimentação do estoque
		if Not grava_log_estoque_v2(usuario, id_nfe_emitente, s_kit_fabricante, s_kit, n_kit_qtde, n_kit_qtde, OP_ESTOQUE_LOG_ENTRADA_VIA_KIT, "", ID_ESTOQUE_VENDA, "", "", "", "", s_documento, "Entrada no estoque de venda via conversão de kit", "") then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
			end if


	'	MONTA SUMÁRIO DOS KITS GERADOS
		Const MAX_CAMPO_VALOR_SUMARIO = 12
		Const MAX_CAMPO_QTDE_SUMARIO = 5
		Const MAX_MARGEM_SUMARIO = 8
		Const TAM_SEPARACAO_COL = 4
		dim s_sumario, s_linha, preco_fabricante_total, vl_custo2_total, qtde_total
		dim vl_BC_ICMS_ST_total, vl_ICMS_ST_total
		
		s_sumario = ""
		preco_fabricante_total = 0
		vl_custo2_total = 0
		vl_BC_ICMS_ST_total = 0
		vl_ICMS_ST_total = 0
		qtde_total = 0
		For i_prod = LBound(v_kit) To UBound(v_kit)
			With v_kit(i_prod)
				If (Trim(.id_estoque) <> "") And (.qtde > 0) Then
					s_linha = Space(MAX_MARGEM_SUMARIO)
					
					s = formata_inteiro(.qtde)
					Do While Len(s) < MAX_CAMPO_QTDE_SUMARIO: s = " " & s: Loop
					s_linha = s_linha & s
					
					s = formata_moeda(.preco_fabricante)
					Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
					s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s
					
					s = formata_moeda(.qtde * .preco_fabricante)
					Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
					s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s

					s = formata_moeda(.vl_custo2)
					Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
					s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s
					
					s = formata_moeda(.qtde * .vl_custo2)
					Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
					s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s
					
					s = formata_moeda(.vl_BC_ICMS_ST)
					Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
					s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s
					
					s = formata_moeda(.vl_ICMS_ST)
					Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
					s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s
					
					If s_sumario <> "" Then s_sumario = s_sumario & Chr(13)
					s_sumario = s_sumario & s_linha
					
					qtde_total = qtde_total + .qtde
					preco_fabricante_total = preco_fabricante_total + (.qtde * .preco_fabricante)
					vl_custo2_total = vl_custo2_total + (.qtde * .vl_custo2)
					vl_BC_ICMS_ST_total = vl_BC_ICMS_ST_total + (.qtde * .vl_BC_ICMS_ST)
					vl_ICMS_ST_total = vl_ICMS_ST_total + (.qtde * .vl_ICMS_ST)
					End If
				End With
			Next

		If s_sumario <> "" Then
		'	ACRESCENTA TÍTULOS
			s_linha = Space(MAX_MARGEM_SUMARIO)
			
			s = "QTDE"
			Do While Len(s) < MAX_CAMPO_QTDE_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & s
			
			s = "PREÇO FABRIC"
			Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s
			
			s = "TOT FABRIC"
			Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s

			s = "CUSTO II"
			Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s
			
			s = "TOT CUSTO II"
			Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s
			
			s = "BC ICMS ST"
			Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s
			
			s = "VL ICMS ST"
			Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & Space(TAM_SEPARACAO_COL) & s
			
			n = MAX_MARGEM_SUMARIO + MAX_CAMPO_QTDE_SUMARIO + TAM_SEPARACAO_COL + _
				MAX_CAMPO_VALOR_SUMARIO + TAM_SEPARACAO_COL + _
				MAX_CAMPO_VALOR_SUMARIO + TAM_SEPARACAO_COL + _
				MAX_CAMPO_VALOR_SUMARIO + TAM_SEPARACAO_COL + _
				MAX_CAMPO_VALOR_SUMARIO + TAM_SEPARACAO_COL + _
				MAX_CAMPO_VALOR_SUMARIO + TAM_SEPARACAO_COL + _
				MAX_CAMPO_VALOR_SUMARIO
			
			s_sumario = s_linha & Chr(13) & _
						String(n, "=") & _
						Chr(13) & s_sumario
			
		'	EXIBE TOTAL GERAL
			s_sumario = s_sumario & Chr(13) & _
						String(n, "=")
			
			s = "TOTAL:"
			Do While Len(s) < MAX_MARGEM_SUMARIO: s = s & " ": Loop
			s_linha = s
			
			s = formata_inteiro(qtde_total)
			Do While Len(s) < MAX_CAMPO_QTDE_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & s
			
			n = TAM_SEPARACAO_COL + MAX_CAMPO_VALOR_SUMARIO + TAM_SEPARACAO_COL
			s_linha = s_linha & Space(n)
			
			s = formata_moeda(preco_fabricante_total)
			Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & s

			n = TAM_SEPARACAO_COL + MAX_CAMPO_VALOR_SUMARIO + TAM_SEPARACAO_COL
			s_linha = s_linha & Space(n)
			
			s = formata_moeda(vl_custo2_total)
			Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & s
			
			n = TAM_SEPARACAO_COL
			s_linha = s_linha & Space(n)
			
			s = formata_moeda(vl_BC_ICMS_ST_total)
			Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & s
			
			n = TAM_SEPARACAO_COL
			s_linha = s_linha & Space(n)
			
			s = formata_moeda(vl_ICMS_ST_total)
			Do While Len(s) < MAX_CAMPO_VALOR_SUMARIO: s = " " & s: Loop
			s_linha = s_linha & s
			
			s_sumario = s_sumario & Chr(13) & s_linha
			End If


	'	INFORMAÇÕES P/ O LOG
		s_log = ""
		For i_prod = LBound(v_item) To UBound(v_item)
			with v_item(i_prod)
				If (Trim(.produto) <> "") And (.qtde > 0) Then
					If s_log <> "" Then s_log = s_log & ", "
					s_log = s_log & CStr(.qtde) & "x" & Trim(.produto)  & "(" & Trim(.fabricante) & ")"
					End If
				end with
			Next
		
		If s_log <> "" Then
			s_log = "Cadastramento de " & Cstr(n_kit_qtde) & _
					" unidades do kit " & Trim(s_kit) & _
					" do fabricante " & s_kit_fabricante & _
					" (NCM: " & c_ncm & ", CST (entrada): " & c_cst & ", Empresa: " & id_nfe_emitente & " - " & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente) & ")" & _
					" composto por: " & s_log
			s_log = s_log & chr(13) & "Sumário dos Kits Gerados:" & chr(13)& s_sumario
			End If

		if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_ESTOQUE_CONVERSAO_KIT, s_log
		
		
	'	PROCESSA OS PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE
		if Not estoque_processa_produtos_vendidos_sem_presenca_v2(id_nfe_emitente, usuario, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
			end if
		
	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
		if Err=0 then 
		'	AGRUPA INFORMAÇÕES PARA A PÁGINA DE SUMÁRIO EXIBIR DETALHES DA OPERAÇÃO
			s = Cstr(id_nfe_emitente) & chr(0) & s_kit_fabricante & chr(0) & s_kit & chr(0) & Cstr(n_kit_qtde) & chr(0) & c_ncm & chr(0) & c_cst & chr(0) & s_documento
			s_aux = ""
			for i = Lbound(v_item) to Ubound(v_item)
				with v_item(i)
					if Trim(.produto) <> "" then
						if s_aux <> "" then s_aux = s_aux & chr(1)
						s_aux = s_aux & .fabricante & chr(2) & .produto & chr(2) & Cstr(.qtde)
						end if
					end with
				next
			if s_aux <> "" then s = s & chr(0)
			s = s & s_aux
			
			s_aux = ""
			for i = Lbound(v_kit) to Ubound(v_kit)
				with v_kit(i)
					if Trim(.id_estoque) <> "" then 
						if s_aux <> "" then s_aux = s_aux & chr(1)
						s_aux = s_aux & .id_estoque & chr(2) & Cstr(.qtde) & chr(2) & formata_moeda(.preco_fabricante) & chr(2) & formata_moeda(.vl_custo2) & chr(2) & formata_moeda(.vl_BC_ICMS_ST) & chr(2) & formata_moeda(.vl_ICMS_ST)
						end if
					end with
				next
			if s_aux <> "" then s = s & chr(0)
			s = s & s_aux
			
			s = s & chr(0) & s_sumario
			Session(SESSION_CLIPBOARD) = s
			Response.Redirect("estoqueconversorkitsumario.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		else
			alerta=Cstr(Err) & ": " & Err.Description
			end if
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



<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>



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
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
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