<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelFaturamento2Exec.asp
'     ======================================================
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
	
	class cl_RelFaturamento2Exec_Totalizacao
		dim qtde
		dim vl_tabela
		dim vl_saida
		dim vl_entrada
		dim vl_RT
		dim vl_lucro_liquido
		end class
	
	class cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
		dim blnHaDados
		dim custoFinancFornecCoeficiente
		dim qtde
		dim vl_tabela
		dim vl_saida
		dim vl_entrada
		dim vl_RT
		dim vl_lucro_liquido
		end class
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_FATURAMENTO2, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, lista_loja, s_filtro_loja, v_loja, v, i, flag_ok
	dim c_dt_inicio, c_dt_termino, c_loja, c_fabricante, c_produto, c_grupo, c_vendedor, c_indicador, c_pedido
    dim v_grupo_pedido_origem, v_pedido_origem, c_grupo_pedido_origem, c_pedido_origem
	dim s_nome_vendedor
	dim op_forma_pagto, c_forma_pagto_qtde_parc
	dim rb_tipo_cliente, rb_saida
	dim c_uf_pesq
    dim c_empresa

	alerta = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	c_indicador = Ucase(Trim(Request.Form("c_indicador")))
	c_pedido = Ucase(Trim(Request.Form("c_pedido")))
    c_grupo_pedido_origem = Trim(Request.Form("c_grupo_pedido_origem"))
    c_pedido_origem = Trim(Request.Form("c_pedido_origem"))
    c_empresa =Trim(Request.Form("c_empresa"))

	if c_pedido <> "" then
		if normaliza_num_pedido(c_pedido) <> "" then c_pedido = normaliza_num_pedido(c_pedido)
		end if

	c_loja = Trim(Request.Form("c_loja"))
	lista_loja = substitui_caracteres(c_loja,chr(10),"")
	v_loja = split(lista_loja,chr(13),-1)
	
	op_forma_pagto = Trim(Request.Form("op_forma_pagto"))
	c_forma_pagto_qtde_parc = retorna_so_digitos(Trim(Request.Form("c_forma_pagto_qtde_parc")))
	rb_tipo_cliente = Trim(Request.Form("rb_tipo_cliente"))
    rb_saida = Trim(Request.Form("rb_saida"))
	c_uf_pesq = Ucase(Trim(Request.Form("c_uf_pesq")))
	
	if alerta = "" then
		if c_fabricante <> "" then
			s = "SELECT fabricante FROM t_FABRICANTE WHERE (fabricante='" & c_fabricante & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "FABRICANTE " & c_fabricante & " NÃO ESTÁ CADASTRADO."
				end if
			end if
		end if

	if alerta = "" then
		if c_pedido <> "" then
			s = "SELECT pedido FROM t_PEDIDO WHERE (pedido='" & c_pedido & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "PEDIDO " & c_pedido & " NÃO ESTÁ CADASTRADO."
				end if
			end if
		end if
		
	if alerta = "" then
		if c_produto <> "" then
			if (Not IsEAN(c_produto)) And (c_fabricante="") then
				alerta=texto_add_br(alerta)
				alerta=alerta & "NÃO FOI ESPECIFICADO O FABRICANTE DO PRODUTO A SER CONSULTADO."
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
							alerta=alerta & "Produto a ser consultado " & c_produto & " NÃO pertence ao fabricante " & c_fabricante & "."
							end if
						end if
					if flag_ok then
					'	CARREGA CÓDIGO INTERNO DO PRODUTO
						c_fabricante = Trim("" & rs("fabricante"))
						c_produto = Trim("" & rs("produto"))
						end if
					end if
				end if
			end if
		end if
	
	if alerta = "" then
		s_nome_vendedor = ""
		if c_vendedor <> "" then
			s = "SELECT nome FROM t_USUARIO WHERE (usuario='" & c_vendedor & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "VENDEDOR " & c_vendedor & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_vendedor = Ucase(Trim("" & rs("nome")))
				end if
			end if
		end if
		
		
'	Período de consulta está restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_termino
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			if c_dt_inicio = "" then c_dt_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
			end if
		
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if


    Const SAIDA_FABRICANTE = "FABRICANTE"
    Const SAIDA_LOJA = "LOJA" 


' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ___________________________________________________________________
' inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
'
sub inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(byref r)
	r.blnHaDados = False
	r.custoFinancFornecCoeficiente = 0
	r.qtde = 0
	r.vl_tabela = 0
	r.vl_saida = 0
	r.vl_entrada = 0
	r.vl_RT = 0
	r.vl_lucro_liquido = 0
end sub


' ___________________________________________________________________
' QuickSort_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
' OBS: ALGORITMO É RECURSIVO
'
sub QuickSort_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(ByRef vetor, ByVal inf, ByVal sup)
dim i, j
dim ref, temp

	set ref = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
	set temp = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc

'	LAÇO DE ORDENAÇÃO
	Do
		i = inf
		j = sup
		
		ref.blnHaDados = vetor((inf + sup) \ 2).blnHaDados
		ref.custoFinancFornecCoeficiente = vetor((inf + sup) \ 2).custoFinancFornecCoeficiente
		ref.qtde = vetor((inf + sup) \ 2).qtde
		ref.vl_tabela = vetor((inf + sup) \ 2).vl_tabela
		ref.vl_saida = vetor((inf + sup) \ 2).vl_saida
		ref.vl_entrada = vetor((inf + sup) \ 2).vl_entrada
		ref.vl_RT = vetor((inf + sup) \ 2).vl_RT
		ref.vl_lucro_liquido = vetor((inf + sup) \ 2).vl_lucro_liquido

		Do
			Do
				If ref.custoFinancFornecCoeficiente > vetor(i).custoFinancFornecCoeficiente Then i = i + 1 Else Exit Do
				Loop

			Do
				If ref.custoFinancFornecCoeficiente < vetor(j).custoFinancFornecCoeficiente Then j = j - 1 Else Exit Do
				Loop

			If i <= j Then
				temp.blnHaDados = vetor(i).blnHaDados
				temp.custoFinancFornecCoeficiente = vetor(i).custoFinancFornecCoeficiente
				temp.qtde= vetor(i).qtde
				temp.vl_tabela = vetor(i).vl_tabela
				temp.vl_saida = vetor(i).vl_saida
				temp.vl_entrada = vetor(i).vl_entrada
				temp.vl_RT = vetor(i).vl_RT
				temp.vl_lucro_liquido = vetor(i).vl_lucro_liquido

				vetor(i).blnHaDados = vetor(j).blnHaDados
				vetor(i).custoFinancFornecCoeficiente = vetor(j).custoFinancFornecCoeficiente
				vetor(i).qtde = vetor(j).qtde
				vetor(i).vl_tabela = vetor(j).vl_tabela
				vetor(i).vl_saida = vetor(j).vl_saida
				vetor(i).vl_entrada = vetor(j).vl_entrada
				vetor(i).vl_RT = vetor(j).vl_RT
				vetor(i).vl_lucro_liquido = vetor(j).vl_lucro_liquido

				vetor(j).blnHaDados = temp.blnHaDados
				vetor(j).custoFinancFornecCoeficiente = temp.custoFinancFornecCoeficiente
				vetor(j).qtde = temp.qtde
				vetor(j).vl_tabela = temp.vl_tabela
				vetor(j).vl_saida = temp.vl_saida
				vetor(j).vl_entrada = temp.vl_entrada
				vetor(j).vl_RT = temp.vl_RT
				vetor(j).vl_lucro_liquido = temp.vl_lucro_liquido

				i = i + 1
				j = j - 1
				End If

			Loop Until i > j

		If inf < j Then QuickSort_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc vetor, inf, j

		inf = i

		Loop Until i >= sup

end sub


' ___________________________________________________________________
' ordena_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
'
sub ordena_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(ByRef vetor, ByVal inf, ByVal sup)
	If inf > sup Then Exit Sub
	QuickSort_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc vetor, inf, sup
end sub


' _____________________________________
' CONSULTA EXECUTA FABRICANTE
'
sub consulta_executa_fabricante
dim r
dim s, s_where, s_where_venda, s_where_devolucao, s_where_loja, s_cor, s_where_temp
dim s_aux, s_sql, cab_table, cab, n_reg, n_reg_total, x, fabricante_a, s_class
dim perc_lucro_bruto, perc_comissao, perc_lucro_liq, perc_custo_financ, perc_desc_medio
dim vl_lucro_liquido
dim i, v, qtde_fabricantes, intIdx, intIteracao
dim blnAchou
dim r_total, r_sub_total
dim v_total_custo_financ, v_sub_total_custo_financ

	set r_total = New cl_RelFaturamento2Exec_Totalizacao
	set r_sub_total = New cl_RelFaturamento2Exec_Totalizacao
	
	redim v_total_custo_financ(0)
	set v_total_custo_financ(0) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
	inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_total_custo_financ(0))
	
	redim v_sub_total_custo_financ(0)
	set v_sub_total_custo_financ(0) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
	inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_sub_total_custo_financ(0))
	
'	CRITÉRIOS COMUNS
	s_where = ""
	if c_grupo <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.grupo = '" & c_grupo & "')"
		
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.grupo IS NOT NULL)"
		end if
	
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.vendedor = '" & c_vendedor & "')"
		end if

	if c_pedido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.pedido = '" & c_pedido & "')"
		end if
		
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if

    s_where_temp = ""
    if c_grupo_pedido_origem <> "" then
        v_grupo_pedido_origem = split(c_grupo_pedido_origem, ", ")
        for i = LBound(v_grupo_pedido_origem) to UBound(v_grupo_pedido_origem)
            s = "SELECT codigo FROM t_CODIGO_DESCRICAO WHERE (codigo_pai = '" & v_grupo_pedido_origem(i) & "') AND grupo='PedidoECommerce_Origem'"
            if rs.State <> 0 then rs.Close
	        rs.open s, cn
		    if rs.Eof then
                alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_grupo_pedido_origem & " NÃO EXISTE."
                exit for
            else
                do while Not rs.Eof
                    if s_where_temp <> "" then s_where_temp = s_where_temp & ", "
                    s_where_temp = s_where_temp & "'" & rs("codigo") & "'"      
                    rs.MoveNext
                loop
            end if
        next
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_PEDIDO.marketplace_codigo_origem IN (" & s_where_temp & "))"
    end if

    s_where_temp = ""
    v_pedido_origem = split(c_pedido_origem, ", ")
    if c_pedido_origem <> "" then
        for i = LBound(v_pedido_origem) to UBound(v_pedido_origem)
            if s_where_temp <> "" then s_where_temp = s_where_temp & ", "
            s_where_temp = s_where_temp & "'" & v_pedido_origem(i) & "'"
        next
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_PEDIDO.marketplace_codigo_origem IN (" & s_where_temp & "))"
    end if
		
'	CRITÉRIO: FORMA DE PAGAMENTO
	s = ""
	if op_forma_pagto <> "" then
		s = " ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_A_VISTA & ") AND (t_PEDIDO__BASE.av_forma_pagto = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELA_UNICA & ") AND (t_PEDIDO__BASE.pu_forma_pagto = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA & ") AND (t_PEDIDO__BASE.pce_forma_pagto_entrada = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA & ") AND (t_PEDIDO__BASE.pce_forma_pagto_prestacao = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA & ") AND (t_PEDIDO__BASE.pse_forma_pagto_prim_prest = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA & ") AND (t_PEDIDO__BASE.pse_forma_pagto_demais_prest = " & op_forma_pagto & ") )"
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

'	CRITÉRIO: QUANTIDADE DE PARCELAS
	s = ""
	if c_forma_pagto_qtde_parc <> "" then
		s = " (t_PEDIDO__BASE.qtde_parcelas = " & c_forma_pagto_qtde_parc & ")"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
		
'	CRITÉRIO: TIPO DE CLIENTE
	s = ""
	if rb_tipo_cliente <> "" then
		s = " (t_CLIENTE.tipo = '" & rb_tipo_cliente & "')"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRITÉRIO: UF DO CLIENTE
	s = ""
	if c_uf_pesq <> "" then
		s = " (t_CLIENTE.uf = '" & c_uf_pesq & "')"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

    if c_empresa <> "" then
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_ESTOQUE.id_nfe_emitente= " & c_empresa & ")"
    end if
		
	s_where_loja = ""
	for i=Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
				s_where_loja = s_where_loja & " (t_PEDIDO.numero_loja = " & v_loja(i) & ")"
			else
				s = ""
				if v(Lbound(v))<>"" then 
					if s <> "" then s = s & " AND"
					s = s & " (t_PEDIDO.numero_loja >= " & v(Lbound(v)) & ")"
					end if
				if v(Ubound(v))<>"" then
					if s <> "" then s = s & " AND"
					s = s & " (t_PEDIDO.numero_loja <= " & v(Ubound(v)) & ")"
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

'	CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	s_where_venda = ""
	if IsDate(c_dt_inicio) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if

	if c_fabricante <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_ESTOQUE_MOVIMENTO.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_ESTOQUE_MOVIMENTO.produto = '" & c_produto & "')"
		end if
	
'	CRITÉRIOS PARA DEVOLUÇÕES
	s_where_devolucao = ""
	if IsDate(c_dt_inicio) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
	if c_fabricante <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.produto = '" & c_produto & "')"
		end if

'	A) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
'	B) O CAMPO 'QTDE' A SER USADO DEVE SER DA TABELA T_ESTOQUE_MOVIMENTO, JÁ
'	QUE UM PEDIDO PODE TER USADO DIVERSOS LOTES DO ESTOQUE PARA ATENDER A
'	UM ÚNICO PRODUTO.  NESSE CASO, HAVERÁ MAIS DE UM REGISTRO EM 
'	T_ESTOQUE_MOVIMENTO SE RELACIONANDO COM O MESMO REGISTRO DE T_PEDIDO_ITEM.
'	A SOMA DE 'QTDE' DOS REGISTROS DE T_ESTOQUE_MOVIMENTO RESULTAM NO VALOR
'	DE 'QTDE' DE T_PEDIDO_ITEM.
	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then s = " AND" & s
	s_sql = "SELECT t_ESTOQUE_MOVIMENTO.fabricante AS fabricante, t_ESTOQUE_MOVIMENTO.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" t_PEDIDO_ITEM.custoFinancFornecCoeficiente," & _
			" Sum((t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda)*(t_PEDIDO__BASE.perc_RT/100)) AS valor_RT," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_lista) AS valor_tabela," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda) AS valor_saida," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde) AS qtde," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_ESTOQUE_ITEM.vl_custo2) AS valor_entrada" & _
			" FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
			" INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
			" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_PEDIDO_ITEM ON ((t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto))" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque='" & ID_ESTOQUE_ENTREGUE & "')" & _
			s & _
			" GROUP BY t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, t_PEDIDO_ITEM.custoFinancFornecCoeficiente"
	
'	LEMBRE-SE: NA DEVOLUÇÃO DO PRODUTO, É CRIADA UMA ENTRADA NO ESTOQUE DE VENDA P/
'	REPRESENTAR A ENTRADA DA MERCADORIA NO ESTOQUE. ENTRETANTO, A QUANTIDADE
'	DEVOLVIDA FICA INICIALMENTE TODA ALOCADA P/ O ESTOQUE DE DEVOLUÇÃO, DEVIDO
'	À NECESSIDADE DE TRATAR A MERCADORIA ANTES DE DISPONIBILIZA-LA P/ VENDA.
'	IMPORTANTE: NO CASO DE OCORRER A DEVOLUÇÃO DE VÁRIAS UNIDADES, PODEM SER
'	CRIADOS VÁRIOS REGISTROS DE ESTOQUE A DIFERENTES CUSTOS DE AQUISIÇÃO.
	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT t_PEDIDO_ITEM_DEVOLVIDO.fabricante AS fabricante, t_PEDIDO_ITEM_DEVOLVIDO.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.custoFinancFornecCoeficiente," & _
			" Sum(-(t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda)*(t_PEDIDO__BASE.perc_RT/100)) AS valor_RT," & _
			" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_lista) AS valor_tabela," & _
			" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS valor_saida," & _
			" Sum(-t_ESTOQUE_ITEM.qtde) AS qtde," & _
			" Sum(-t_ESTOQUE_ITEM.qtde*t_ESTOQUE_ITEM.vl_custo2) AS valor_entrada" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)" & _
			" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_ESTOQUE ON (t_PEDIDO_ITEM_DEVOLVIDO.id=t_ESTOQUE.devolucao_id_item_devolvido)" & _
			" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_ESTOQUE_ITEM.produto))" & _ 
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			" LEFT JOIN t_PRODUTO ON ((t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_PRODUTO.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_PRODUTO.produto))" & _
			s & _
			" GROUP BY t_PEDIDO_ITEM_DEVOLVIDO.fabricante, t_PEDIDO_ITEM_DEVOLVIDO.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, t_PEDIDO_ITEM_DEVOLVIDO.custoFinancFornecCoeficiente"
			
	s_sql = s_sql & " ORDER BY fabricante, produto, descricao, descricao_html, custoFinancFornecCoeficiente, qtde DESC"

  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE tdCodProd' style='vertical-align:bottom' NOWRAP><P class='R'>Código</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdDescrProd' style='vertical-align:bottom' NOWRAP><P class='R'>Descrição</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdCustoFinanc' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Custo Financ (%)</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdQtde' align='right' style='vertical-align:bottom' NOWRAP><P class='Rd' style='font-weight:bold;'>Qtde</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdFatTab' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Faturamento Tabela (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdDescMedio' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Desc Médio (%)</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdFatTot' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Faturamento Real (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdCmvTot' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>CMV Total (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdLucro' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Lucro (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdPercLucroTot' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>% Lucro</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdVlRt' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>COM (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdRtPercFat' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>COM (% do Fat)</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdLucroLiq' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Lucro Bruto (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdLucroLiqPercFat' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Lucro Bruto (% do Fat)</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
'	CALCULA ANTECIPADAMENTE OS TOTAIS P/ OS CÁLCULOS DE PERCENTUAIS SOBRE O FATURAMENTO TOTAL
	with r_total
		.qtde = 0
		.vl_tabela = 0
		.vl_saida = 0
		.vl_entrada = 0
		.vl_RT = 0
		.vl_lucro_liquido = 0
		end with
	
	set r = cn.execute(s_sql)
	n_reg = 0
	do while Not r.Eof
		n_reg = n_reg + 1
		
		with r_total
			.qtde = .qtde + r("qtde")
			.vl_tabela = .vl_tabela + r("valor_tabela")
			.vl_saida = .vl_saida + r("valor_saida")
			.vl_entrada = .vl_entrada + r("valor_entrada")
			.vl_RT = .vl_RT + r("valor_RT")
			.vl_lucro_liquido = .vl_lucro_liquido + (r("valor_saida")-r("valor_entrada")) - r("valor_RT")
			end with
		
		blnAchou = False
		intIdx = -1
		for i=Lbound(v_total_custo_financ) to Ubound(v_total_custo_financ)
			if (v_total_custo_financ(i).custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")) And v_total_custo_financ(i).blnHaDados then
				blnAchou = True
				intIdx = i
				exit for
				end if
			next
		
		if Not blnAchou then
			redim preserve v_total_custo_financ(Ubound(v_total_custo_financ)+1)
			set v_total_custo_financ(Ubound(v_total_custo_financ)) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
			inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_total_custo_financ(Ubound(v_total_custo_financ)))
			intIdx = Ubound(v_total_custo_financ)
			with v_total_custo_financ(intIdx)
				.blnHaDados = True
				.custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")
				end with
			end if
		
		with v_total_custo_financ(intIdx)
			.qtde = .qtde + r("qtde")
			.vl_tabela = .vl_tabela + r("valor_tabela")
			.vl_saida = .vl_saida + r("valor_saida")
			.vl_entrada = .vl_entrada + r("valor_entrada")
			.vl_RT = .vl_RT + r("valor_RT")
			.vl_lucro_liquido = .vl_lucro_liquido + (r("valor_saida")-r("valor_entrada")) - r("valor_RT")
			end with
		
		r.MoveNext
		loop

	if n_reg > 0 then r.MoveFirst
	
	ordena_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc v_total_custo_financ, Lbound(v_total_custo_financ), Ubound(v_total_custo_financ)
	
'	LAÇO PARA PROCESSAMENTO DO RELATÓRIO
	x = ""
	n_reg = 0
	n_reg_total = 0
	r_sub_total.qtde = 0
	qtde_fabricantes = 0
	
	fabricante_a = "XXXXX"
	do while Not r.Eof
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante"))<>fabricante_a then
			fabricante_a = Trim("" & r("fabricante"))
			qtde_fabricantes = qtde_fabricantes + 1
		'	FECHA TABELA DO FABRICANTE ANTERIOR
			if n_reg_total > 0 then
			'	EXIBE SUBTOTAL POR COEFICIENTE DE CUSTO FINANCEIRO
				intIteracao = 0
				for i=Lbound(v_sub_total_custo_financ) to Ubound(v_sub_total_custo_financ)
					with v_sub_total_custo_financ(i)
						if .blnHaDados then
							intIteracao = intIteracao + 1
							s_cor="black"
							if .qtde < 0 then s_cor="red"
							if .vl_tabela < 0 then s_cor="red"
							if .vl_saida < 0 then s_cor="red"
							if .vl_entrada < 0 then s_cor="red"
							if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
							perc_custo_financ = 100 * (.custoFinancFornecCoeficiente - 1)
							
							if .vl_saida = 0 then
								perc_lucro_bruto = 0
							else
								perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
								end if
							
							if .vl_saida = 0 then
								perc_comissao = 0
								perc_lucro_liq = 0
							else
								perc_comissao = (.vl_RT/.vl_saida)*100
								perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
								end if
							
							if .vl_tabela = 0 then
								perc_desc_medio = 0
							else
								perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
								end if
							
							if intIteracao > 1 then
								s_class = "MB"
							else
								s_class = "MTB"
								end if
							
							x = x & _
								"	<TR NOWRAP style='background:white;'>" & chr(13) & _
								"		<TD class='" & s_class & " ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
								"&nbsp;</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_custo_financ) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
								"	</TR>" & chr(13)
							end if
						end with
					next

			'	EXIBE SUBTOTAL GERAL
				with r_sub_total
					if .vl_saida = 0 then
						perc_lucro_bruto = 0
					else
						perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_comissao = 0
					else
						perc_comissao = (.vl_RT/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_lucro_liq = 0
					else
						perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
						end if
					
					if .vl_tabela = 0 then
						perc_desc_medio = 0
					else
						perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
						end if
					end with
				
				s_cor="black"
				with r_sub_total
					if .qtde < 0 then s_cor="red"
					if .vl_tabela < 0 then s_cor="red"
					if .vl_saida < 0 then s_cor="red"
					if .vl_entrada < 0 then s_cor="red"
					if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
					x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
							"		<TD class='MB ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
							"TOTAL:</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & "&nbsp;" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
							"		<TD class='MB MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
							"	</TR>" & chr(13) & _
							"</TABLE>" & chr(13)
					end with

				Response.Write x
				x="<BR>" & chr(13)
				end if

			redim v_sub_total_custo_financ(0)
			set v_sub_total_custo_financ(0) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
			inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_sub_total_custo_financ(0))
			
			n_reg = 0
			with r_sub_total
				.vl_tabela = 0
				.vl_saida = 0
				.vl_entrada = 0
				.qtde = 0
				.vl_RT = 0
				.vl_lucro_liquido = 0
				end with

			if n_reg_total > 0 then x = x & "<BR>" & chr(13)
			s = Trim("" & r("fabricante"))
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then x = x & "	<TR><TD class='MDTE' COLSPAN='14' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td></tr>" & chr(13)
			x = x & cab
			end if
		

	'	CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)

		s_cor="black"
		if IsNumeric(r("qtde")) then if Clng(r("qtde")) < 0 then s_cor="red"

	 '> CÓDIGO DO PRODUTO
		x = x & "		<TD class='MDTE tdCodProd' valign='bottom'><P class='Cn' style='color:" & s_cor & ";'>" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO DO PRODUTO
		s = Trim("" & r("descricao_html"))
		if s <> "" then s = produto_formata_descricao_em_html(s) else s = "&nbsp;"
		x = x & "		<TD class='MTD tdDescrProd' valign='bottom'><P class='Cn' style='color:" & s_cor & ";'>" & s & "</P></TD>" & chr(13)

	 '> CUSTO FINANCEIRO DO FABRICANTE (%)
		perc_custo_financ = 100 * (r("custoFinancFornecCoeficiente") - 1)
		x = x & "		<TD align='right' class='MTD tdCustoFinanc'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_custo_financ) & "%" & "</P></TD>" & chr(13)

	 '> QUANTIDADE
		x = x & "		<TD align='right' valign='bottom' class='MTD tdQtde'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(r("qtde")) & "</P></TD>" & chr(13)

	 '> VALOR TABELA
		x = x & "		<TD align='right' valign='bottom' class='MTD tdFatTab'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_tabela")) & "</P></TD>" & chr(13)

	 '> DESCONTO MÉDIO
		if r("valor_tabela") = 0 then
			perc_desc_medio = 0
		else
			perc_desc_medio = 100 * (r("valor_tabela") - r("valor_saida")) / r("valor_tabela")
			end if
		x = x & "		<TD align='right' class='MTD tdDescMedio'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</P></TD>" & chr(13)

	 '> VALOR SAÍDA
		x = x & "		<TD align='right' valign='bottom' class='MTD tdFatTot'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_saida")) & "</P></TD>" & chr(13)

	 '> VALOR ENTRADA
		x = x & "		<TD align='right' valign='bottom' class='MTD tdCmvTot'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_entrada")) & "</P></TD>" & chr(13)

	 '> LUCRO BRUTO
		x = x & "		<TD align='right' valign='bottom' class='MTD tdLucro'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_saida")-r("valor_entrada")) & "</P></TD>" & chr(13)

	 '> PERCENTUAL DO LUCRO BRUTO TOTAL
		if CCur(r("valor_saida")) = CCur(0) then
			perc_lucro_bruto = 0
		else
			perc_lucro_bruto = ((r("valor_saida")-r("valor_entrada"))/r("valor_saida"))*100
			end if
		x = x & "		<TD align='right' valign='bottom' class='MTD tdPercLucroTot'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</P></TD>" & chr(13)
		
	 '> VALOR COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
		x = x & "		<TD align='right' class='MTD tdVlRt'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_RT")) & "</P></TD>" & chr(13)
		
	 '> COMISSÃO (% DO FAT)
		if CCur(r("valor_saida")) = CCur(0) then
			perc_comissao = 0
		else
			perc_comissao = (r("valor_RT")/r("valor_saida"))*100
			end if
		x = x & "		<TD align='right' class='MTD tdRtPercFat'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</P></TD>" & chr(13)
		
	 '> LUCRO LÍQUIDO (DESCONTADA A COMISSÃO)
		vl_lucro_liquido = r("valor_saida")-r("valor_entrada")-r("valor_RT")
		x = x & "		<TD align='right' class='MTD tdLucroLiq'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lucro_liquido) & "</P></TD>" & chr(13)
	 
	 '> PERCENTUAL DO LUCRO LÍQUIDO SOBRE O FATURAMENTO
		if CCur(r("valor_saida")) = CCur(0) then
			perc_lucro_liq = 0
		else
			perc_lucro_liq = (vl_lucro_liquido/r("valor_saida"))*100
			end if
		x = x & "		<TD align='right' class='MTD tdLucroLiqPercFat'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</P></TD>" & chr(13)
		
	'	SUBTOTAL POR COEFICIENTE DE CUSTO FINANCEIRO
		blnAchou = False
		intIdx = -1
		for i=Lbound(v_sub_total_custo_financ) to UBound(v_sub_total_custo_financ)
			if (v_sub_total_custo_financ(i).custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")) And v_sub_total_custo_financ(i).blnHaDados then
				blnAchou = True
				intIdx = i
				exit for
				end if
			next
		
		if Not blnAchou then
			redim preserve v_sub_total_custo_financ(Ubound(v_sub_total_custo_financ)+1)
			set v_sub_total_custo_financ(Ubound(v_sub_total_custo_financ)) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
			inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_sub_total_custo_financ(Ubound(v_sub_total_custo_financ)))
			intIdx = Ubound(v_sub_total_custo_financ)
			with v_sub_total_custo_financ(intIdx)
				.blnHaDados = True
				.custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")
				end with
			end if
		
		with v_sub_total_custo_financ(intIdx)
			.qtde = .qtde + r("qtde")
			.vl_tabela = .vl_tabela + r("valor_tabela")
			.vl_saida = .vl_saida + r("valor_saida")
			.vl_entrada = .vl_entrada + r("valor_entrada")
			.vl_RT = .vl_RT + r("valor_RT")
			.vl_lucro_liquido = .vl_lucro_liquido + vl_lucro_liquido
			end with
		
		if Not blnAchou then
			ordena_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc v_sub_total_custo_financ, LBound(v_sub_total_custo_financ), UBound(v_sub_total_custo_financ)
			end if
		
	'	SUBTOTAL GERAL
		with r_sub_total
			.qtde = .qtde + r("qtde")
			.vl_tabela = .vl_tabela + r("valor_tabela")
			.vl_saida = .vl_saida + r("valor_saida")
			.vl_entrada = .vl_entrada + r("valor_entrada")
			.vl_RT = .vl_RT + r("valor_RT")
			.vl_lucro_liquido = .vl_lucro_liquido + vl_lucro_liquido
			end with
		
		x = x & "	</TR>" & chr(13)
			
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
'	MOSTRA SUBTOTAL DO ÚLTIMO FABRICANTE
	if n_reg <> 0 then 
	'	EXIBE SUBTOTAL POR COEFICIENTE DE CUSTO FINANCEIRO
		intIteracao = 0
		for i=Lbound(v_sub_total_custo_financ) to Ubound(v_sub_total_custo_financ)
			with v_sub_total_custo_financ(i)
				if .blnHaDados then
					intIteracao = intIteracao + 1
					s_cor="black"
					if .qtde < 0 then s_cor="red"
					if .vl_tabela < 0 then s_cor="red"
					if .vl_saida < 0 then s_cor="red"
					if .vl_entrada < 0 then s_cor="red"
					if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
				
					perc_custo_financ = 100 * (.custoFinancFornecCoeficiente - 1)
					
					if .vl_saida = 0 then
						perc_lucro_bruto = 0
					else
						perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_comissao = 0
						perc_lucro_liq = 0
					else
						perc_comissao = (.vl_RT/.vl_saida)*100
						perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
						end if
					
					if .vl_tabela = 0 then
						perc_desc_medio = 0
					else
						perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
						end if
					
					if intIteracao > 1 then
						s_class = "MB"
					else
						s_class = "MTB"
						end if
					
					x = x & _
						"	<TR NOWRAP style='background:white;'>" & chr(13) & _
						"		<TD class='" & s_class & " ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
						"&nbsp;</p></td>" & chr(13) & _
						"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_custo_financ) & "%" & "</p></td>" & chr(13) & _
						"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
						"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
						"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
						"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
						"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
						"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
						"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
						"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
						"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
						"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
						"		<TD class='" & s_class & " MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
						"	</TR>" & chr(13)
					end if
				end with
			next

	'	EXIBE SUBTOTAL GERAL
		with r_sub_total
			if .vl_saida = 0 then
				perc_lucro_bruto = 0
			else
				perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
				end if
			
			if .vl_saida = 0 then
				perc_comissao = 0
			else
				perc_comissao = (.vl_RT/.vl_saida)*100
				end if
			
			if .vl_saida = 0 then
				perc_lucro_liq = 0
			else
				perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
				end if
			
			if .vl_tabela = 0 then
				perc_desc_medio = 0
			else
				perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
				end if
			end with
		
		s_cor="black"
		with r_sub_total
			if .qtde < 0 then s_cor="red"
			if .vl_tabela < 0 then s_cor="red"
			if .vl_saida < 0 then s_cor="red"
			if .vl_entrada < 0 then s_cor="red"
			if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
			
			x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='2' class='MB ME' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
					"TOTAL:</p></td>" & chr(13) & _
					"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & "&nbsp;" & "</p></td>" & chr(13) & _
					"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
					"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
					"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
					"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
					"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
					"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
					"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
					"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
					"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
					"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
					"		<TD class='MB MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			end with
		
	'>	TOTAL GERAL
		if qtde_fabricantes > 1 then
		'	TOTALIZAÇÃO POR CUSTO FINANCEIRO
			x = x & "	<TR><TD COLSPAN='14' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
					"	<TR><TD COLSPAN='14' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13)
			
			x = x & _
				"	<TR><TD class='MDTE' COLSPAN='14' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & "TOTAL GERAL" & "</p></td></tr>" & chr(13)
			
			intIteracao = 0
			for i=Lbound(v_total_custo_financ) to Ubound(v_total_custo_financ)
				with v_total_custo_financ(i)
					if .blnHaDados then
						intIteracao = intIteracao + 1
						s_cor="black"
						if .qtde < 0 then s_cor="red"
						if .vl_tabela < 0 then s_cor="red"
						if .vl_saida < 0 then s_cor="red"
						if .vl_entrada < 0 then s_cor="red"
						if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
					
						perc_custo_financ = 100 * (.custoFinancFornecCoeficiente - 1)
						
						if .vl_saida = 0 then
							perc_lucro_bruto = 0
						else
							perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
							end if
						
						if .vl_saida = 0 then
							perc_comissao = 0
							perc_lucro_liq = 0
						else
							perc_comissao = (.vl_RT/.vl_saida)*100
							perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
							end if
						
						if .vl_tabela = 0 then
							perc_desc_medio = 0
						else
							perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
							end if
						
						if intIteracao > 1 then
							s_class = "MB"
						else
							s_class = "MTB"
							end if
						
						x = x & _
							"	<TR NOWRAP style='background:white;'>" & chr(13) & _
							"		<TD class='" & s_class & " ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
							"&nbsp;</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_custo_financ) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & " MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
							"	</TR>" & chr(13)
						end if
					end with
				next
		
		'	TOTALIZAÇÃO GERAL
			with r_total
				if .vl_saida = 0 then
					perc_lucro_bruto = 0
				else
					perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
					end if
				
				if .vl_saida = 0 then
					perc_comissao = 0
					perc_lucro_liq = 0
				else
					perc_comissao = (.vl_RT/.vl_saida)*100
					perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
					end if
				
				if .vl_tabela = 0 then
					perc_desc_medio = 0
				else
					perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
					end if
				end with
			
			s_cor="black"
			with r_total
				if .qtde < 0 then s_cor="red"
				if .vl_tabela < 0 then s_cor="red"
				if .vl_saida < 0 then s_cor="red"
				if .vl_entrada < 0 then s_cor="red"
				if (.vl_saida-.vl_entrada) < 0 then s_cor="red"

				x = x & "	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
						"		<TD class='MB ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
						"TOTAL GERAL:</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & "&nbsp;" & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
						"		<TD class='MB MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
						"	</TR>" & chr(13)
				end with
			end if
		end if

'	MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = ""
		if c_fabricante <> "" then
			s = c_fabricante
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			if s <> "" then x = x & cab_table & "	<TR><TD class='MDTE' COLSPAN='14' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td></tr>" & chr(13) & cab
		else
			x = x & cab_table & cab
			end if

		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='14'><P class='ALERTA'>&nbsp;NENHUM PRODUTO SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

'	FECHA TABELA DO ÚLTIMO FABRICANTE
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	
end sub

' _____________________________________
' CONSULTA EXECUTA LOJA
'
sub consulta_executa_loja
dim r
dim s, s_where, s_where_venda, s_where_devolucao, s_where_loja, s_cor, s_where_temp
dim s_aux, s_sql, cab_table, cab, n_reg, n_reg_total, x, fabricante_a, loja_a, s_class
dim perc_lucro_bruto, perc_comissao, perc_lucro_liq, perc_custo_financ, perc_desc_medio
dim vl_lucro_liquido
dim i, v, qtde_fabricantes, qtde_lojas, intIdx, intIteracao
dim blnAchou, blnPulaTotalFab
dim r_total, r_sub_total, r_sub_total_loja
dim v_total_custo_financ, v_sub_total_custo_financ
dim v_total_custo_financ_loja, v_sub_total_custo_financ_loja


	set r_total = New cl_RelFaturamento2Exec_Totalizacao
	set r_sub_total = New cl_RelFaturamento2Exec_Totalizacao

    set r_sub_total_loja = New cl_RelFaturamento2Exec_Totalizacao
	
	redim v_total_custo_financ(0)
	set v_total_custo_financ(0) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
	inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_total_custo_financ(0))
	
	redim v_sub_total_custo_financ(0)
	set v_sub_total_custo_financ(0) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
	inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_sub_total_custo_financ(0))

    redim v_total_custo_financ_loja(0)
    set v_total_custo_financ_loja(0) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
    inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_total_custo_financ_loja(0))

    redim v_sub_total_custo_financ_loja(0)
    set v_sub_total_custo_financ_loja(0) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
    inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_sub_total_custo_financ_loja(0))
	
'	CRITÉRIOS COMUNS
	s_where = ""
	if c_grupo <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.grupo = '" & c_grupo & "')"
		
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.grupo IS NOT NULL)"
		end if
	
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.vendedor = '" & c_vendedor & "')"
		end if

	if c_pedido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.pedido = '" & c_pedido & "')"
		end if
		
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if

    s_where_temp = ""
    if c_grupo_pedido_origem <> "" then
        v_grupo_pedido_origem = split(c_grupo_pedido_origem, ", ")
        for i = LBound(v_grupo_pedido_origem) to UBound(v_grupo_pedido_origem)
            s = "SELECT codigo FROM t_CODIGO_DESCRICAO WHERE (codigo_pai = '" & v_grupo_pedido_origem(i) & "') AND grupo='PedidoECommerce_Origem'"
            if rs.State <> 0 then rs.Close
	        rs.open s, cn
		    if rs.Eof then
                alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_grupo_pedido_origem & " NÃO EXISTE."
                exit for
            else
                do while Not rs.Eof
                    if s_where_temp <> "" then s_where_temp = s_where_temp & ", "
                    s_where_temp = s_where_temp & "'" & rs("codigo") & "'"      
                    rs.MoveNext
                loop
            end if
        next
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_PEDIDO.marketplace_codigo_origem IN (" & s_where_temp & "))"
    end if

    s_where_temp = ""
    v_pedido_origem = split(c_pedido_origem, ", ")
    if c_pedido_origem <> "" then
        for i = LBound(v_pedido_origem) to UBound(v_pedido_origem)
            if s_where_temp <> "" then s_where_temp = s_where_temp & ", "
            s_where_temp = s_where_temp & "'" & v_pedido_origem(i) & "'"
        next
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_PEDIDO.marketplace_codigo_origem IN (" & s_where_temp & "))"
    end if
		
'	CRITÉRIO: FORMA DE PAGAMENTO
	s = ""
	if op_forma_pagto <> "" then
		s = " ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_A_VISTA & ") AND (t_PEDIDO__BASE.av_forma_pagto = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELA_UNICA & ") AND (t_PEDIDO__BASE.pu_forma_pagto = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA & ") AND (t_PEDIDO__BASE.pce_forma_pagto_entrada = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA & ") AND (t_PEDIDO__BASE.pce_forma_pagto_prestacao = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA & ") AND (t_PEDIDO__BASE.pse_forma_pagto_prim_prest = " & op_forma_pagto & ") )" & _
			" OR ( (t_PEDIDO__BASE.tipo_parcelamento = " & COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA & ") AND (t_PEDIDO__BASE.pse_forma_pagto_demais_prest = " & op_forma_pagto & ") )"
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

'	CRITÉRIO: QUANTIDADE DE PARCELAS
	s = ""
	if c_forma_pagto_qtde_parc <> "" then
		s = " (t_PEDIDO__BASE.qtde_parcelas = " & c_forma_pagto_qtde_parc & ")"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
		
'	CRITÉRIO: TIPO DE CLIENTE
	s = ""
	if rb_tipo_cliente <> "" then
		s = " (t_CLIENTE.tipo = '" & rb_tipo_cliente & "')"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if

'	CRITÉRIO: UF DO CLIENTE
	s = ""
	if c_uf_pesq <> "" then
		s = " (t_CLIENTE.uf = '" & c_uf_pesq & "')"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if


    if c_empresa <> "" then
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_ESTOQUE.id_nfe_emitente= " & c_empresa & ")"
    end if
		
	s_where_loja = ""
	for i=Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
				s_where_loja = s_where_loja & " (t_PEDIDO.numero_loja = " & v_loja(i) & ")"
			else
				s = ""
				if v(Lbound(v))<>"" then 
					if s <> "" then s = s & " AND"
					s = s & " (t_PEDIDO.numero_loja >= " & v(Lbound(v)) & ")"
					end if
				if v(Ubound(v))<>"" then
					if s <> "" then s = s & " AND"
					s = s & " (t_PEDIDO.numero_loja <= " & v(Ubound(v)) & ")"
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

'	CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	s_where_venda = ""
	if IsDate(c_dt_inicio) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if

	if c_fabricante <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_ESTOQUE_MOVIMENTO.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_ESTOQUE_MOVIMENTO.produto = '" & c_produto & "')"
		end if
	
'	CRITÉRIOS PARA DEVOLUÇÕES
	s_where_devolucao = ""
	if IsDate(c_dt_inicio) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
	if c_fabricante <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.produto = '" & c_produto & "')"
		end if

'	A) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
'	B) O CAMPO 'QTDE' A SER USADO DEVE SER DA TABELA T_ESTOQUE_MOVIMENTO, JÁ
'	QUE UM PEDIDO PODE TER USADO DIVERSOS LOTES DO ESTOQUE PARA ATENDER A
'	UM ÚNICO PRODUTO.  NESSE CASO, HAVERÁ MAIS DE UM REGISTRO EM 
'	T_ESTOQUE_MOVIMENTO SE RELACIONANDO COM O MESMO REGISTRO DE T_PEDIDO_ITEM.
'	A SOMA DE 'QTDE' DOS REGISTROS DE T_ESTOQUE_MOVIMENTO RESULTAM NO VALOR
'	DE 'QTDE' DE T_PEDIDO_ITEM.
	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then s = " AND" & s
	s_sql = "SELECT t_PEDIDO.loja AS loja, t_ESTOQUE_MOVIMENTO.fabricante AS fabricante, t_ESTOQUE_MOVIMENTO.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" t_PEDIDO_ITEM.custoFinancFornecCoeficiente," & _
			" Sum((t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda)*(t_PEDIDO__BASE.perc_RT/100)) AS valor_RT," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_lista) AS valor_tabela," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda) AS valor_saida," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde) AS qtde," & _
			" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_ESTOQUE_ITEM.vl_custo2) AS valor_entrada" & _
			" FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
			" INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
			" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_PEDIDO_ITEM ON ((t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto))" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
            " INNER JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.id_estoque = t_ESTOQUE.id_estoque)" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque='" & ID_ESTOQUE_ENTREGUE & "')" & _
			s & _
			" GROUP BY t_PEDIDO.loja, t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, t_PEDIDO_ITEM.custoFinancFornecCoeficiente"
	
'	LEMBRE-SE: NA DEVOLUÇÃO DO PRODUTO, É CRIADA UMA ENTRADA NO ESTOQUE DE VENDA P/
'	REPRESENTAR A ENTRADA DA MERCADORIA NO ESTOQUE. ENTRETANTO, A QUANTIDADE
'	DEVOLVIDA FICA INICIALMENTE TODA ALOCADA P/ O ESTOQUE DE DEVOLUÇÃO, DEVIDO
'	À NECESSIDADE DE TRATAR A MERCADORIA ANTES DE DISPONIBILIZA-LA P/ VENDA.
'	IMPORTANTE: NO CASO DE OCORRER A DEVOLUÇÃO DE VÁRIAS UNIDADES, PODEM SER
'	CRIADOS VÁRIOS REGISTROS DE ESTOQUE A DIFERENTES CUSTOS DE AQUISIÇÃO.
	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT t_PEDIDO.loja AS loja, t_PEDIDO_ITEM_DEVOLVIDO.fabricante AS fabricante, t_PEDIDO_ITEM_DEVOLVIDO.produto AS produto," & _
			" t_PRODUTO.descricao AS descricao," & _
			" t_PRODUTO.descricao_html AS descricao_html," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.custoFinancFornecCoeficiente," & _
			" Sum(-(t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda)*(t_PEDIDO__BASE.perc_RT/100)) AS valor_RT," & _
			" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_lista) AS valor_tabela," & _
			" Sum(-t_ESTOQUE_ITEM.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS valor_saida," & _
			" Sum(-t_ESTOQUE_ITEM.qtde) AS qtde," & _
			" Sum(-t_ESTOQUE_ITEM.qtde*t_ESTOQUE_ITEM.vl_custo2) AS valor_entrada" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)" & _
			" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_ESTOQUE ON (t_PEDIDO_ITEM_DEVOLVIDO.id=t_ESTOQUE.devolucao_id_item_devolvido)" & _
			" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_ESTOQUE_ITEM.produto))" & _ 
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			" LEFT JOIN t_PRODUTO ON ((t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_PRODUTO.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_PRODUTO.produto))" & _
			s & _
			" GROUP BY t_PEDIDO.loja, t_PEDIDO_ITEM_DEVOLVIDO.fabricante, t_PEDIDO_ITEM_DEVOLVIDO.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, t_PEDIDO_ITEM_DEVOLVIDO.custoFinancFornecCoeficiente"
			
	s_sql = s_sql & " ORDER BY loja, fabricante, produto, descricao, descricao_html, custoFinancFornecCoeficiente, qtde DESC"

  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE tdCodProd' style='vertical-align:bottom' NOWRAP><P class='R'>Código</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdDescrProd' style='vertical-align:bottom' NOWRAP><P class='R'>Descrição</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdCustoFinanc' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Custo Financ (%)</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdQtde' align='right' style='vertical-align:bottom' NOWRAP><P class='Rd' style='font-weight:bold;'>Qtde</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdFatTab' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Faturamento Tabela (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdDescMedio' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Desc Médio (%)</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdFatTot' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Faturamento Real (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdCmvTot' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>CMV Total (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdLucro' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Lucro (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdPercLucroTot' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>% Lucro</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdVlRt' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>COM (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdRtPercFat' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>COM (% do Fat)</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdLucroLiq' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Lucro Bruto (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdLucroLiqPercFat' align='right' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Lucro Bruto (% do Fat)</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
'	CALCULA ANTECIPADAMENTE OS TOTAIS P/ OS CÁLCULOS DE PERCENTUAIS SOBRE O FATURAMENTO TOTAL
	with r_total
		.qtde = 0
		.vl_tabela = 0
		.vl_saida = 0
		.vl_entrada = 0
		.vl_RT = 0
		.vl_lucro_liquido = 0
		end with
	
	set r = cn.execute(s_sql)
	n_reg = 0
	do while Not r.Eof
		n_reg = n_reg + 1
		
		with r_total
			.qtde = .qtde + r("qtde")
			.vl_tabela = .vl_tabela + r("valor_tabela")
			.vl_saida = .vl_saida + r("valor_saida")
			.vl_entrada = .vl_entrada + r("valor_entrada")
			.vl_RT = .vl_RT + r("valor_RT")
			.vl_lucro_liquido = .vl_lucro_liquido + (r("valor_saida")-r("valor_entrada")) - r("valor_RT")
			end with
		
		blnAchou = False
		intIdx = -1
		for i=Lbound(v_total_custo_financ) to Ubound(v_total_custo_financ)
			if (v_total_custo_financ(i).custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")) And v_total_custo_financ(i).blnHaDados then
				blnAchou = True
				intIdx = i
				exit for
				end if
			next
		
		if Not blnAchou then
			redim preserve v_total_custo_financ(Ubound(v_total_custo_financ)+1)
			set v_total_custo_financ(Ubound(v_total_custo_financ)) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
			inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_total_custo_financ(Ubound(v_total_custo_financ)))
			intIdx = Ubound(v_total_custo_financ)
			with v_total_custo_financ(intIdx)
				.blnHaDados = True
				.custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")
				end with
			end if
		
		with v_total_custo_financ(intIdx)
			.qtde = .qtde + r("qtde")
			.vl_tabela = .vl_tabela + r("valor_tabela")
			.vl_saida = .vl_saida + r("valor_saida")
			.vl_entrada = .vl_entrada + r("valor_entrada")
			.vl_RT = .vl_RT + r("valor_RT")
			.vl_lucro_liquido = .vl_lucro_liquido + (r("valor_saida")-r("valor_entrada")) - r("valor_RT")
			end with

        blnAchou = False
		intIdx = -1
		for i=Lbound(v_total_custo_financ_loja) to Ubound(v_total_custo_financ_loja)
			if (v_total_custo_financ_loja(i).custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")) And v_total_custo_financ_loja(i).blnHaDados then
				blnAchou = True
				intIdx = i
				exit for
				end if
			next
		
		if Not blnAchou then
			redim preserve v_total_custo_financ_loja(Ubound(v_total_custo_financ_loja)+1)
			set v_total_custo_financ_loja(Ubound(v_total_custo_financ_loja)) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
			inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_total_custo_financ_loja(Ubound(v_total_custo_financ_loja)))
			intIdx = Ubound(v_total_custo_financ_loja)
			with v_total_custo_financ_loja(intIdx)
				.blnHaDados = True
				.custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")
				end with
			end if
		
		with v_total_custo_financ_loja(intIdx)
			.qtde = .qtde + r("qtde")
			.vl_tabela = .vl_tabela + r("valor_tabela")
			.vl_saida = .vl_saida + r("valor_saida")
			.vl_entrada = .vl_entrada + r("valor_entrada")
			.vl_RT = .vl_RT + r("valor_RT")
			.vl_lucro_liquido = .vl_lucro_liquido + (r("valor_saida")-r("valor_entrada")) - r("valor_RT")
			end with
		
		r.MoveNext
		loop

	if n_reg > 0 then r.MoveFirst
	
	ordena_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc v_total_custo_financ, Lbound(v_total_custo_financ), Ubound(v_total_custo_financ)
	ordena_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc v_total_custo_financ_loja, Lbound(v_total_custo_financ_loja), Ubound(v_total_custo_financ_loja)
	
'	LAÇO PARA PROCESSAMENTO DO RELATÓRIO
	x = ""
	n_reg = 0
	n_reg_total = 0
	r_sub_total.qtde = 0
    r_sub_total_loja.qtde = 0
	qtde_fabricantes = 0
    qtde_lojas = 0
	blnPulaTotalFab = False

	fabricante_a = "XXXXX"
	do while Not r.Eof

    '	MUDOU DE LOJA?
		if Trim("" & r("loja"))<>loja_a then
			loja_a = Trim("" & r("loja"))
			qtde_lojas = qtde_lojas + 1
            fabricante_a = "XXXXXXX"
		'	FECHA TABELA DA LOJA ANTERIOR
			if n_reg_total > 0 then
                blnPulaTotalFab = True

			'	EXIBE SUBTOTAL POR COEFICIENTE DE CUSTO FINANCEIRO DO FABRICANTE
        
				intIteracao = 0
				for i=Lbound(v_sub_total_custo_financ) to Ubound(v_sub_total_custo_financ)
					with v_sub_total_custo_financ(i)
						if .blnHaDados then
							intIteracao = intIteracao + 1
							s_cor="black"
							if .qtde < 0 then s_cor="red"
							if .vl_tabela < 0 then s_cor="red"
							if .vl_saida < 0 then s_cor="red"
							if .vl_entrada < 0 then s_cor="red"
							if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
							perc_custo_financ = 100 * (.custoFinancFornecCoeficiente - 1)
							
							if .vl_saida = 0 then
								perc_lucro_bruto = 0
							else
								perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
								end if
							
							if .vl_saida = 0 then
								perc_comissao = 0
								perc_lucro_liq = 0
							else
								perc_comissao = (.vl_RT/.vl_saida)*100
								perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
								end if
							
							if .vl_tabela = 0 then
								perc_desc_medio = 0
							else
								perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
								end if
							
							if intIteracao > 1 then
								s_class = "MB"
							else
								s_class = "MTB"
								end if
							
							x = x & _
								"	<TR NOWRAP style='background:white;'>" & chr(13) & _
								"		<TD class='" & s_class & " ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
								"&nbsp;</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_custo_financ) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
								"	</TR>" & chr(13)
							end if
						end with
					next

			'	EXIBE SUBTOTAL GERAL DO FABRICANTE
				with r_sub_total
					if .vl_saida = 0 then
						perc_lucro_bruto = 0
					else
						perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_comissao = 0
					else
						perc_comissao = (.vl_RT/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_lucro_liq = 0
					else
						perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
						end if
					
					if .vl_tabela = 0 then
						perc_desc_medio = 0
					else
						perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
						end if
					end with
				
				s_cor="black"
				with r_sub_total
					if .qtde < 0 then s_cor="red"
					if .vl_tabela < 0 then s_cor="red"
					if .vl_saida < 0 then s_cor="red"
					if .vl_entrada < 0 then s_cor="red"
					if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
					x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
							"		<TD class='MB ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
							"TOTAL:</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & "&nbsp;" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
							"		<TD class='MB MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
							"	</TR>" & chr(13) & _
							"</TABLE>" & chr(13)
					end with

				Response.Write x
				x = ""

			'	EXIBE SUBTOTAL POR COEFICIENTE DE CUSTO FINANCEIRO DA LOJA
				intIteracao = 0
				x = x & "</td></tr>"
				x = x & "<tr><td colspan='14' class='MD ME' style='background-color:white'>&nbsp;</td></tr>"
				
				for i=Lbound(v_sub_total_custo_financ_loja) to Ubound(v_sub_total_custo_financ_loja)
					with v_sub_total_custo_financ_loja(i)
						if .blnHaDados then
							intIteracao = intIteracao + 1
							s_cor="black"
							if .qtde < 0 then s_cor="red"
							if .vl_tabela < 0 then s_cor="red"
							if .vl_saida < 0 then s_cor="red"
							if .vl_entrada < 0 then s_cor="red"
							if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
							perc_custo_financ = 100 * (.custoFinancFornecCoeficiente - 1)
							
							if .vl_saida = 0 then
								perc_lucro_bruto = 0
							else
								perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
								end if
							
							if .vl_saida = 0 then
								perc_comissao = 0
								perc_lucro_liq = 0
							else
								perc_comissao = (.vl_RT/.vl_saida)*100
								perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
								end if
							
							if .vl_tabela = 0 then
								perc_desc_medio = 0
							else
								perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
								end if
							
							if intIteracao > 1 then
								s_class = "MB"
							else
								s_class = "MTB"
								end if
							x = x & _
								"	<TR NOWRAP style='background:#eee;'>" & chr(13) & _
								"		<TD class='" & s_class & " ME tdCodProd' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
								"&nbsp;</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdDescrProd'><p class='Cd' style='color:" & s_cor & ";'>&nbsp;</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdCustoFinanc'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_custo_financ) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdQtde'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
								"		<TD class='" & s_class & " tdFatTab'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdDescMedio'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdFatTot'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdCmvTot'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdLucro'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdPercLucroTot'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdVlRt'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdRtPercFat'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdLucroLiq'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " MD tdLucroLiqPercFat'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
								"	</TR>" & chr(13)
							end if
						end with
					next

			'	EXIBE SUBTOTAL GERAL
				with r_sub_total_loja
					if .vl_saida = 0 then
						perc_lucro_bruto = 0
					else
						perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_comissao = 0
					else
						perc_comissao = (.vl_RT/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_lucro_liq = 0
					else
						perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
						end if
					
					if .vl_tabela = 0 then
						perc_desc_medio = 0
					else
						perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
						end if
					end with
				
				s_cor="black"
				with r_sub_total_loja
					if .qtde < 0 then s_cor="red"
					if .vl_tabela < 0 then s_cor="red"
					if .vl_saida < 0 then s_cor="red"
					if .vl_entrada < 0 then s_cor="red"
					if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
					x = x & "	<TR NOWRAP style='background: #ccc'>" & chr(13) & _
							"		<TD class='MB ME' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
							"&nbsp;</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & _ 
                            "TOTAL DA LOJA:" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & "&nbsp;" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
							"		<TD class='MB MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
							"	</TR>" & chr(13) & _
							"</TABLE>" & chr(13)
					end with
                
				Response.Write x
				x="<BR>" & chr(13)
				end if

			redim v_sub_total_custo_financ_loja(0)
			set v_sub_total_custo_financ_loja(0) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
			inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_sub_total_custo_financ_loja(0))
			
			n_reg = 0
			with r_sub_total_loja
				.vl_tabela = 0
				.vl_saida = 0
				.vl_entrada = 0
				.qtde = 0
				.vl_RT = 0
				.vl_lucro_liquido = 0
				end with

			if n_reg_total > 0 then x = x & "<BR>" & chr(13)
			s = Trim("" & r("loja"))
			x = x & "<table cellspacing=0 cellpadding=0 id='tableLoja_" & qtde_lojas & "'>"
			if s <> "" then x = x & "	<TR><TD title='exibe ou oculta o detalhamento da loja' onclick='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(qtde_lojas) & chr(34) & ")' class='MDTE' valign='bottom' colspan='14' style='background:#fff6e5;cursor:hand'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13)
            x = x & "<td><a name='aExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(qtde_lojas) & chr(34) & ")' title='exibe ou oculta o detalhamento da loja'><img src='../BOTAO/view_bottom.png' class='notPrint' style='margin-left: 3px'></a></td></tr>" & chr(13)
            x = x & "<tr style='display:none'><td class='ME MD' colspan='14'>"
			end if
    
	'	MUDOU DE FABRICANTE?
		if Trim("" & r("fabricante"))<>fabricante_a then
			fabricante_a = Trim("" & r("fabricante"))
			qtde_fabricantes = qtde_fabricantes + 1
		'	FECHA TABELA DO FABRICANTE ANTERIOR
			if n_reg_total > 0 And Not blnPulaTotalFab then
			'	EXIBE SUBTOTAL POR COEFICIENTE DE CUSTO FINANCEIRO
				intIteracao = 0
				for i=Lbound(v_sub_total_custo_financ) to Ubound(v_sub_total_custo_financ)
					with v_sub_total_custo_financ(i)
						if .blnHaDados then
							intIteracao = intIteracao + 1
							s_cor="black"
							if .qtde < 0 then s_cor="red"
							if .vl_tabela < 0 then s_cor="red"
							if .vl_saida < 0 then s_cor="red"
							if .vl_entrada < 0 then s_cor="red"
							if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
							perc_custo_financ = 100 * (.custoFinancFornecCoeficiente - 1)
							
							if .vl_saida = 0 then
								perc_lucro_bruto = 0
							else
								perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
								end if
							
							if .vl_saida = 0 then
								perc_comissao = 0
								perc_lucro_liq = 0
							else
								perc_comissao = (.vl_RT/.vl_saida)*100
								perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
								end if
							
							if .vl_tabela = 0 then
								perc_desc_medio = 0
							else
								perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
								end if
							
							if intIteracao > 1 then
								s_class = "MB"
							else
								s_class = "MTB"
								end if
							
							x = x & _
								"	<TR NOWRAP style='background:white;'>" & chr(13) & _
								"		<TD class='" & s_class & " ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
								"&nbsp;</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_custo_financ) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
								"	</TR>" & chr(13)
							end if
						end with
					next

			'	EXIBE SUBTOTAL GERAL
				with r_sub_total
					if .vl_saida = 0 then
						perc_lucro_bruto = 0
					else
						perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_comissao = 0
					else
						perc_comissao = (.vl_RT/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_lucro_liq = 0
					else
						perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
						end if
					
					if .vl_tabela = 0 then
						perc_desc_medio = 0
					else
						perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
						end if
					end with
				
				s_cor="black"
				with r_sub_total
					if .qtde < 0 then s_cor="red"
					if .vl_tabela < 0 then s_cor="red"
					if .vl_saida < 0 then s_cor="red"
					if .vl_entrada < 0 then s_cor="red"
					if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
					x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
							"		<TD class='MB ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
							"TOTAL:</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & "&nbsp;" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
							"		<TD class='MB MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
							"	</TR>" & chr(13) & _
							"</TABLE>" & chr(13)
					end with

				Response.Write x
				x="<BR>" & chr(13)
				end if

			redim v_sub_total_custo_financ(0)
			set v_sub_total_custo_financ(0) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
			inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_sub_total_custo_financ(0))
			
			n_reg = 0
			with r_sub_total
				.vl_tabela = 0
				.vl_saida = 0
				.vl_entrada = 0
				.qtde = 0
				.vl_RT = 0
				.vl_lucro_liquido = 0
				end with

			if n_reg_total > 0 then x = x & "<BR>" & chr(13)
			s = Trim("" & r("fabricante"))
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then x = x & "	<TR><TD class='MDTE' COLSPAN='14' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td></tr>" & chr(13)
			x = x & cab
            blnPulaTotalFab = False
			end if		

	'	CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)

		s_cor="black"
		if IsNumeric(r("qtde")) then if Clng(r("qtde")) < 0 then s_cor="red"

	 '> CÓDIGO DO PRODUTO
		x = x & "		<TD class='MDTE tdCodProd' valign='bottom'><P class='Cn' style='color:" & s_cor & ";'>" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO DO PRODUTO
		s = Trim("" & r("descricao_html"))
		if s <> "" then s = produto_formata_descricao_em_html(s) else s = "&nbsp;"
		x = x & "		<TD class='MTD tdDescrProd' valign='bottom'><P class='Cn' style='color:" & s_cor & ";'>" & s & "</P></TD>" & chr(13)

	 '> CUSTO FINANCEIRO DO FABRICANTE (%)
		perc_custo_financ = 100 * (r("custoFinancFornecCoeficiente") - 1)
		x = x & "		<TD align='right' class='MTD tdCustoFinanc'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_custo_financ) & "%" & "</P></TD>" & chr(13)

	 '> QUANTIDADE
		x = x & "		<TD align='right' valign='bottom' class='MTD tdQtde'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_inteiro(r("qtde")) & "</P></TD>" & chr(13)

	 '> VALOR TABELA
		x = x & "		<TD align='right' valign='bottom' class='MTD tdFatTab'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_tabela")) & "</P></TD>" & chr(13)

	 '> DESCONTO MÉDIO
		if r("valor_tabela") = 0 then
			perc_desc_medio = 0
		else
			perc_desc_medio = 100 * (r("valor_tabela") - r("valor_saida")) / r("valor_tabela")
			end if
		x = x & "		<TD align='right' class='MTD tdDescMedio'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</P></TD>" & chr(13)

	 '> VALOR SAÍDA
		x = x & "		<TD align='right' valign='bottom' class='MTD tdFatTot'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_saida")) & "</P></TD>" & chr(13)

	 '> VALOR ENTRADA
		x = x & "		<TD align='right' valign='bottom' class='MTD tdCmvTot'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_entrada")) & "</P></TD>" & chr(13)

	 '> LUCRO BRUTO
		x = x & "		<TD align='right' valign='bottom' class='MTD tdLucro'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_saida")-r("valor_entrada")) & "</P></TD>" & chr(13)

	 '> PERCENTUAL DO LUCRO BRUTO TOTAL
		if CCur(r("valor_saida")) = CCur(0) then
			perc_lucro_bruto = 0
		else
			perc_lucro_bruto = ((r("valor_saida")-r("valor_entrada"))/r("valor_saida"))*100
			end if
		x = x & "		<TD align='right' valign='bottom' class='MTD tdPercLucroTot'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</P></TD>" & chr(13)
		
	 '> VALOR COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
		x = x & "		<TD align='right' class='MTD tdVlRt'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_RT")) & "</P></TD>" & chr(13)
		
	 '> COMISSÃO (% DO FAT)
		if CCur(r("valor_saida")) = CCur(0) then
			perc_comissao = 0
		else
			perc_comissao = (r("valor_RT")/r("valor_saida"))*100
			end if
		x = x & "		<TD align='right' class='MTD tdRtPercFat'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</P></TD>" & chr(13)
		
	 '> LUCRO LÍQUIDO (DESCONTADA A COMISSÃO)
		vl_lucro_liquido = r("valor_saida")-r("valor_entrada")-r("valor_RT")
		x = x & "		<TD align='right' class='MTD tdLucroLiq'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_lucro_liquido) & "</P></TD>" & chr(13)
	 
	 '> PERCENTUAL DO LUCRO LÍQUIDO SOBRE O FATURAMENTO
		if CCur(r("valor_saida")) = CCur(0) then
			perc_lucro_liq = 0
		else
			perc_lucro_liq = (vl_lucro_liquido/r("valor_saida"))*100
			end if
		x = x & "		<TD align='right' class='MTD tdLucroLiqPercFat'><P class='Cnd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</P></TD>" & chr(13)
		
	'	SUBTOTAL POR COEFICIENTE DE CUSTO FINANCEIRO
		blnAchou = False
		intIdx = -1
		for i=Lbound(v_sub_total_custo_financ) to UBound(v_sub_total_custo_financ)
			if (v_sub_total_custo_financ(i).custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")) And v_sub_total_custo_financ(i).blnHaDados then
				blnAchou = True
				intIdx = i
				exit for
				end if
			next
		
		if Not blnAchou then
			redim preserve v_sub_total_custo_financ(Ubound(v_sub_total_custo_financ)+1)
			set v_sub_total_custo_financ(Ubound(v_sub_total_custo_financ)) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
			inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_sub_total_custo_financ(Ubound(v_sub_total_custo_financ)))
			intIdx = Ubound(v_sub_total_custo_financ)
			with v_sub_total_custo_financ(intIdx)
				.blnHaDados = True
				.custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")
				end with
			end if
		
		with v_sub_total_custo_financ(intIdx)
			.qtde = .qtde + r("qtde")
			.vl_tabela = .vl_tabela + r("valor_tabela")
			.vl_saida = .vl_saida + r("valor_saida")
			.vl_entrada = .vl_entrada + r("valor_entrada")
			.vl_RT = .vl_RT + r("valor_RT")
			.vl_lucro_liquido = .vl_lucro_liquido + vl_lucro_liquido
			end with
		
		if Not blnAchou then
			ordena_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc v_sub_total_custo_financ, LBound(v_sub_total_custo_financ), UBound(v_sub_total_custo_financ)
			end if

    '	SUBTOTAL POR COEFICIENTE DE CUSTO FINANCEIRO DA LOJA
		blnAchou = False
		intIdx = -1
		for i=Lbound(v_sub_total_custo_financ_loja) to UBound(v_sub_total_custo_financ_loja)
			if (v_sub_total_custo_financ_loja(i).custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")) And v_sub_total_custo_financ_loja(i).blnHaDados then
				blnAchou = True
				intIdx = i
				exit for
				end if
			next
		
		if Not blnAchou then
			redim preserve v_sub_total_custo_financ_loja(Ubound(v_sub_total_custo_financ_loja)+1)
			set v_sub_total_custo_financ_loja(Ubound(v_sub_total_custo_financ_loja)) = New cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc
			inicializa_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc(v_sub_total_custo_financ_loja(Ubound(v_sub_total_custo_financ_loja)))
			intIdx = Ubound(v_sub_total_custo_financ_loja)
			with v_sub_total_custo_financ_loja(intIdx)
				.blnHaDados = True
				.custoFinancFornecCoeficiente = r("custoFinancFornecCoeficiente")
				end with
			end if
		
		with v_sub_total_custo_financ_loja(intIdx)
			.qtde = .qtde + r("qtde")
			.vl_tabela = .vl_tabela + r("valor_tabela")
			.vl_saida = .vl_saida + r("valor_saida")
			.vl_entrada = .vl_entrada + r("valor_entrada")
			.vl_RT = .vl_RT + r("valor_RT")
			.vl_lucro_liquido = .vl_lucro_liquido + vl_lucro_liquido
			end with
		
		if Not blnAchou then
			ordena_cl_RelFaturamento2Exec_TotalizacaoPorCustoFinanc v_sub_total_custo_financ_loja, LBound(v_sub_total_custo_financ_loja), UBound(v_sub_total_custo_financ_loja)
			end if
		
	'	SUBTOTAL GERAL FABRICANTE
		with r_sub_total
			.qtde = .qtde + r("qtde")
			.vl_tabela = .vl_tabela + r("valor_tabela")
			.vl_saida = .vl_saida + r("valor_saida")
			.vl_entrada = .vl_entrada + r("valor_entrada")
			.vl_RT = .vl_RT + r("valor_RT")
			.vl_lucro_liquido = .vl_lucro_liquido + vl_lucro_liquido
			end with

    '	SUBTOTAL GERAL LOJA
		with r_sub_total_loja
			.qtde = .qtde + r("qtde")
			.vl_tabela = .vl_tabela + r("valor_tabela")
			.vl_saida = .vl_saida + r("valor_saida")
			.vl_entrada = .vl_entrada + r("valor_entrada")
			.vl_RT = .vl_RT + r("valor_RT")
			.vl_lucro_liquido = .vl_lucro_liquido + vl_lucro_liquido
			end with
		
		x = x & "	</TR>" & chr(13)
			
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
    '   MOSTRA O TOTAL DA ÚLTIMA LOJA
        if n_reg <> 0 then
			'	EXIBE SUBTOTAL POR COEFICIENTE DE CUSTO FINANCEIRO DO FABRICANTE        
            intIteracao = 0
				for i=Lbound(v_sub_total_custo_financ) to Ubound(v_sub_total_custo_financ)
					with v_sub_total_custo_financ(i)
						if .blnHaDados then
							intIteracao = intIteracao + 1
							s_cor="black"
							if .qtde < 0 then s_cor="red"
							if .vl_tabela < 0 then s_cor="red"
							if .vl_saida < 0 then s_cor="red"
							if .vl_entrada < 0 then s_cor="red"
							if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
							perc_custo_financ = 100 * (.custoFinancFornecCoeficiente - 1)
							
							if .vl_saida = 0 then
								perc_lucro_bruto = 0
							else
								perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
								end if
							
							if .vl_saida = 0 then
								perc_comissao = 0
								perc_lucro_liq = 0
							else
								perc_comissao = (.vl_RT/.vl_saida)*100
								perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
								end if
							
							if .vl_tabela = 0 then
								perc_desc_medio = 0
							else
								perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
								end if
							
							if intIteracao > 1 then
								s_class = "MB"
							else
								s_class = "MTB"
								end if
							
							x = x & _
								"	<TR NOWRAP style='background:white;'>" & chr(13) & _
								"		<TD class='" & s_class & " ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
								"&nbsp;</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_custo_financ) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
								"	</TR>" & chr(13)
							end if
						end with
					next

			'	EXIBE SUBTOTAL GERAL DO FABRICANTE
				with r_sub_total
					if .vl_saida = 0 then
						perc_lucro_bruto = 0
					else
						perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_comissao = 0
					else
						perc_comissao = (.vl_RT/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_lucro_liq = 0
					else
						perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
						end if
					
					if .vl_tabela = 0 then
						perc_desc_medio = 0
					else
						perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
						end if
					end with
				
				s_cor="black"
				with r_sub_total
					if .qtde < 0 then s_cor="red"
					if .vl_tabela < 0 then s_cor="red"
					if .vl_saida < 0 then s_cor="red"
					if .vl_entrada < 0 then s_cor="red"
					if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
					x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
							"		<TD class='MB ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
							"TOTAL:</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & "&nbsp;" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
							"		<TD class='MB MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
							"	</TR>" & chr(13) & _
							"</TABLE>" & chr(13)
					end with

				Response.Write x
                x = ""

			'	EXIBE SUBTOTAL POR COEFICIENTE DE CUSTO FINANCEIRO DA LOJA
				intIteracao = 0
                x = x & "</td></tr>"
                x = x & "<tr><td colspan='14' class='MD ME' style='background-color:white'>&nbsp;</td></tr>"
                
				for i=Lbound(v_sub_total_custo_financ_loja) to Ubound(v_sub_total_custo_financ_loja)
					with v_sub_total_custo_financ_loja(i)
						if .blnHaDados then
							intIteracao = intIteracao + 1
							s_cor="black"
							if .qtde < 0 then s_cor="red"
							if .vl_tabela < 0 then s_cor="red"
							if .vl_saida < 0 then s_cor="red"
							if .vl_entrada < 0 then s_cor="red"
							if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
							perc_custo_financ = 100 * (.custoFinancFornecCoeficiente - 1)
							
							if .vl_saida = 0 then
								perc_lucro_bruto = 0
							else
								perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
								end if
							
							if .vl_saida = 0 then
								perc_comissao = 0
								perc_lucro_liq = 0
							else
								perc_comissao = (.vl_RT/.vl_saida)*100
								perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
								end if
							
							if .vl_tabela = 0 then
								perc_desc_medio = 0
							else
								perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
								end if
							
							if intIteracao > 1 then
								s_class = "MB"
							else
								s_class = "MTB"
								end if
							x = x & _
								"	<TR NOWRAP style='background:#eee;'>" & chr(13) & _
								"		<TD class='" & s_class & " ME tdCodProd' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
								"&nbsp;</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdDescrProd'><p class='Cd' style='color:" & s_cor & ";'>&nbsp;</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdCustoFinanc'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_custo_financ) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdQtde'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
								"		<TD class='" & s_class & " tdFatTab'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdDescMedio'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdFatTot'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdCmvTot'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdLucro'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdPercLucroTot'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdVlRt'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdRtPercFat'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " tdLucroLiq'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
								"		<TD class='" & s_class & " MD tdLucroLiqPercFat'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
								"	</TR>" & chr(13)
							end if
						end with
					next

			'	EXIBE SUBTOTAL GERAL
				with r_sub_total_loja
					if .vl_saida = 0 then
						perc_lucro_bruto = 0
					else
						perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_comissao = 0
					else
						perc_comissao = (.vl_RT/.vl_saida)*100
						end if
					
					if .vl_saida = 0 then
						perc_lucro_liq = 0
					else
						perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
						end if
					
					if .vl_tabela = 0 then
						perc_desc_medio = 0
					else
						perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
						end if
					end with
				
				s_cor="black"
				with r_sub_total_loja
					if .qtde < 0 then s_cor="red"
					if .vl_tabela < 0 then s_cor="red"
					if .vl_saida < 0 then s_cor="red"
					if .vl_entrada < 0 then s_cor="red"
					if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
						
					x = x & "	<TR NOWRAP style='background: #ccc'>" & chr(13) & _
							"		<TD class='MB ME' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
							"&nbsp;</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & _ 
                            "TOTAL DA LOJA:" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & "&nbsp;" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
							"		<TD class='MB MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
							"	</TR>" & chr(13)
					end with
		
	'>	TOTAL GERAL
		if qtde_lojas > 1 then
		'	TOTALIZAÇÃO POR CUSTO FINANCEIRO
			x = x & "	<TR><TD COLSPAN='14' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
					"	<TR><TD COLSPAN='14' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13)
			
			x = x & _
				"	<TR><TD class='MDTE' COLSPAN='14' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & "TOTAL GERAL" & "</p></td></tr>" & chr(13)
			
			intIteracao = 0
			for i=Lbound(v_total_custo_financ) to Ubound(v_total_custo_financ)
				with v_total_custo_financ(i)
					if .blnHaDados then
						intIteracao = intIteracao + 1
						s_cor="black"
						if .qtde < 0 then s_cor="red"
						if .vl_tabela < 0 then s_cor="red"
						if .vl_saida < 0 then s_cor="red"
						if .vl_entrada < 0 then s_cor="red"
						if (.vl_saida-.vl_entrada) < 0 then s_cor="red"
					
						perc_custo_financ = 100 * (.custoFinancFornecCoeficiente - 1)
						
						if .vl_saida = 0 then
							perc_lucro_bruto = 0
						else
							perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
							end if
						
						if .vl_saida = 0 then
							perc_comissao = 0
							perc_lucro_liq = 0
						else
							perc_comissao = (.vl_RT/.vl_saida)*100
							perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
							end if
						
						if .vl_tabela = 0 then
							perc_desc_medio = 0
						else
							perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
							end if
						
						if intIteracao > 1 then
							s_class = "MB"
						else
							s_class = "MTB"
							end if
						
						x = x & _
							"	<TR NOWRAP style='background:white;'>" & chr(13) & _
							"		<TD class='" & s_class & " ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
							"&nbsp;</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_custo_financ) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & "'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
							"		<TD class='" & s_class & " MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
							"	</TR>" & chr(13)
						end if
					end with
				next
		
		'	TOTALIZAÇÃO GERAL
			with r_total
				if .vl_saida = 0 then
					perc_lucro_bruto = 0
				else
					perc_lucro_bruto = ((.vl_saida-.vl_entrada)/.vl_saida)*100
					end if
				
				if .vl_saida = 0 then
					perc_comissao = 0
					perc_lucro_liq = 0
				else
					perc_comissao = (.vl_RT/.vl_saida)*100
					perc_lucro_liq = (.vl_lucro_liquido/.vl_saida)*100
					end if
				
				if .vl_tabela = 0 then
					perc_desc_medio = 0
				else
					perc_desc_medio = 100 * (.vl_tabela - .vl_saida) / .vl_tabela
					end if
				end with
			
			s_cor="black"
			with r_total
				if .qtde < 0 then s_cor="red"
				if .vl_tabela < 0 then s_cor="red"
				if .vl_saida < 0 then s_cor="red"
				if .vl_entrada < 0 then s_cor="red"
				if (.vl_saida-.vl_entrada) < 0 then s_cor="red"

				x = x & "	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
						"		<TD class='MB ME' COLSPAN='2' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
						"TOTAL GERAL:</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & "&nbsp;" & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_inteiro(.qtde) & "</p></TD>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_tabela) & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_desc_medio) & "%" & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida) & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_entrada) & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_saida-.vl_entrada) & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_bruto) & "%" & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_RT) & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_comissao) & "%" & "</p></td>" & chr(13) & _
						"		<TD class='MB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(.vl_lucro_liquido) & "</p></td>" & chr(13) & _
						"		<TD class='MB MD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_perc(perc_lucro_liq) & "%" & "</p></td>" & chr(13) & _
						"	</TR>" & chr(13)
				end with
			end if
		end if

'	MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = ""
		if c_fabricante <> "" then
			s = c_fabricante
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			if s <> "" then x = x & cab_table & "	<TR><TD class='MDTE' COLSPAN='14' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td></tr>" & chr(13) & cab
		else
			x = x & cab_table & cab
			end if

		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='14'><P class='ALERTA'>&nbsp;NENHUM PRODUTO SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

'	FECHA TABELA DO ÚLTIMO FABRICANTE
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

<html>
    
<head>
	<title>CENTRAL</title>
	</head>

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';
</script>
<script type="text/javascript">
    function fExibeOcultaCampos(indice) {
        $('#tableLoja_' + indice + ' tr:eq(1)').toggle();
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
.tdCodProd{
	vertical-align: top;
	width: 50px;
	}
.tdDescrProd{
	vertical-align: top;
	width: 190px;
	}
.tdQtde{
	vertical-align: top;
	width: 40px;
	}
.tdCustoFinanc{
	vertical-align: top;
	width: 45px;
	}
.tdFatTab{
	vertical-align: top;
	width: 80px;
	}
.tdDescMedio{
	vertical-align: top;
	width: 54px;
	}
.tdFatTot{
	vertical-align: top;
	width: 80px;
	}
.tdCmvTot{
	vertical-align: top;
	width: 80px;
	}
.tdLucro{
	vertical-align: top;
	width: 80px;
	}
.tdPercLucroTot{
	vertical-align: top;
	width: 54px;
	}
.tdVlRt{
	vertical-align: top;
	width: 80px;
	}
.tdRtPercFat{
	vertical-align: top;
	width: 54px;
	}
.tdLucroLiq{
	vertical-align: top;
	width: 80px;
	}
.tdLucroLiqPercFat{
	vertical-align: top;
	width: 54px;
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
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>
    
<% else %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">
<input type="hidden" name="c_grupo" id="c_grupo" value="<%=c_grupo%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="c_pedido" id="c_pedido" value="<%=c_pedido%>">
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="1064" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Faturamento II</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='1064' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>"
	
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Período:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s_aux = ""
	if c_fabricante <> "" then s_aux = x_fabricante(c_fabricante)
	s = c_fabricante
	if (s <> "") And (s_aux <> "") then s = s & " - " & s_aux
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Fabricante:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"
	
	s = c_produto
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Produto:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"
	
	s = c_grupo
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Grupo de Produtos:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s = c_vendedor
	if s = "" then 
		s = "todos"
	else
		if (s_nome_vendedor <> "") And (s_nome_vendedor <> c_vendedor) then s = s & " (" & s_nome_vendedor & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Vendedor:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s = c_indicador
	if s = "" then 
		s = "todos"
	else
		s = s & " (" & x_orcamentista_e_indicador(c_indicador) & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Indicador:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s = c_pedido
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Pedido:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

    s = c_grupo_pedido_origem
	if s = "" then 
		s = "todos"
	else
        v_grupo_pedido_origem = split(c_grupo_pedido_origem, ", ")
        s = ""
        for i = Lbound(v_grupo_pedido_origem) to Ubound(v_grupo_pedido_origem)
            if s <> "" then s = s & ", "
		    s = s & obtem_descricao_tabela_t_codigo_descricao("PedidoECommerce_Origem_Grupo", v_grupo_pedido_origem(i))
        next
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Origem Pedido (Grupo):&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = c_pedido_origem
	if s = "" then 
		s = "todos"
	else
		v_pedido_origem = split(c_pedido_origem, ", ")
        s = ""
        for i = Lbound(v_pedido_origem) to Ubound(v_pedido_origem)
            if s <> "" then s = s & ", "
		    s = s & obtem_descricao_tabela_t_codigo_descricao("PedidoECommerce_Origem", v_pedido_origem(i))
        next
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Origem do Pedido:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	s = x_opcao_forma_pagamento(op_forma_pagto)
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Forma de Pagamento:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	if c_forma_pagto_qtde_parc <> "" then
		s = c_forma_pagto_qtde_parc
		if s = "" then s = "&nbsp;"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP><p class='N'>Nº Parcelas:&nbsp;</p></td>" & chr(13) & _
					"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

	s = rb_tipo_cliente
	if s = "" then 
		s = "todos"
	elseif s = ID_PF then
		s = "Pessoa Física"
	elseif s = ID_PJ then
		s = "Pessoa Jurídica"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Tipo de Cliente:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s = c_uf_pesq
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>UF:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

    s = c_empresa
	if s = "" then 
		s = "todas"
	else
		s = obtem_apelido_empresa_NFe_emitente(c_empresa)
		end if
	 s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				       "<span class='N'>Empresa:&nbsp;</span></td><td align='left' valign='top'>" & _
				       "<span class='N'>" & s & "</span></td></tr>"

	s_filtro_loja = ""
	for i = Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
				s_filtro_loja = s_filtro_loja & v_loja(i)
			else
				if (v(Lbound(v))<>"") And (v(Ubound(v))<>"") then 
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Lbound(v)) & " a " & v(Ubound(v))
				elseif (v(Lbound(v))<>"") And (v(Ubound(v))="") then
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Lbound(v)) & " e acima"
				elseif (v(Lbound(v))="") And (v(Ubound(v))<>"") then
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Ubound(v)) & " e abaixo"
					end if
				end if
			end if
		next
	s = s_filtro_loja
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Loja(s):&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & formata_data_hora(Now) & "</p></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% 
    select case rb_saida
        case SAIDA_FABRICANTE
            consulta_executa_fabricante 
        case SAIDA_LOJA
            consulta_executa_loja
    end select
    
%>

<!-- ************   SEPARADOR   ************ -->
<table width="1064" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>
    
<table class="notPrint" width="1064" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

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
