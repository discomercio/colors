<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =========================================================================
'	  R E L C O M I S S A O I N D I C A D O R E S P A G D E S C E X E C . A S P
'     =========================================================================
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

    if Not operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
      Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
    end if
'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	dim s, s_aux, s_filtro
	dim ckb_st_entrega_entregue, c_dt_entregue_mes, c_dt_entregue_ano, str_data, mes, ano
	dim ckb_comissao_paga_sim, ckb_comissao_paga_nao
	dim ckb_st_pagto_pago, ckb_st_pagto_nao_pago, ckb_st_pagto_pago_parcial
	dim c_vendedor, c_indicador
	dim c_loja, lista_loja, s_filtro_loja, v_loja, v, i
	dim rb_visao, blnVisaoSintetica
    dim v_vendedor, vendedor_temp, j, aviso
    
    v_vendedor = ""
	alerta = ""

	ckb_st_entrega_entregue = Trim(Request.Form("ckb_st_entrega_entregue"))
	c_dt_entregue_mes = Trim(Request.Form("c_dt_entregue_mes"))
	c_dt_entregue_ano = Trim(Request.Form("c_dt_entregue_ano"))
    mes = c_dt_entregue_mes
    ano = c_dt_entregue_ano

    if len(Cstr(mes)) = 1 then mes =  "0" & Cstr(mes)
    str_data = "01/" & Cstr(mes) & "/" & Cstr(ano)
    str_data = strToDate(str_data)
    str_data = DateAdd("m",1,str_data)

	c_vendedor = Trim(Request.Form("c_vendedor"))
	c_indicador = Trim(Request.Form("c_indicador"))

	ckb_comissao_paga_sim = Trim(Request.Form("ckb_comissao_paga_sim"))
	ckb_comissao_paga_nao = Trim(Request.Form("ckb_comissao_paga_nao"))

	ckb_st_pagto_pago = Trim(Request.Form("ckb_st_pagto_pago"))
	ckb_st_pagto_nao_pago = Trim(Request.Form("ckb_st_pagto_nao_pago"))
	ckb_st_pagto_pago_parcial = Trim(Request.Form("ckb_st_pagto_pago_parcial"))
	rb_visao = Trim(Request.Form("rb_visao"))
	
	blnVisaoSintetica = False
	if rb_visao = "SINTETICA" then blnVisaoSintetica = True
	
	c_loja = Trim(Request.Form("c_loja"))
	lista_loja = substitui_caracteres(c_loja,chr(10),"")
	v_loja = split(lista_loja,chr(13),-1)
    v_vendedor = split(c_vendedor, ", ")

dim o
dim resultadoCalculo,resultadoDigito,QtdeCedulas,TotalCedula,limitador(5)
dim dadosCalculo
set o = createobject( "ComPlusCalcCedulas.ComPlusCalcCedulas" )
dim v_cedulas,aux(5),y,totalArredondado, cont
dim cedulas()
dim qtdeCedula()
        
' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'

sub consulta_executa
const VENDA_NORMAL = "VEN"
const DEVOLUCAO = "DEV"
const PERDA = "PER"
dim r
dim s, s_aux, s_sql, x, cab_table, cab, indicador_a, n_reg, n_reg_total, qtde_indicadores, ind_anterior
dim vl_preco_venda, vl_sub_total_preco_venda, vl_total_preco_venda, cont2
dim vl_preco_NF, vl_sub_total_preco_NF, vl_total_preco_NF
dim vl_RT, vl_sub_total_RT, vl_total_RT
dim vl_RA, vl_sub_total_RA, vl_total_RA
dim vl_RA_liquido, vl_sub_total_RA_liquido, vl_total_RA_liquido
dim vl_RA_diferenca, vl_sub_total_RA_diferenca, vl_total_RA_diferenca
dim perc_RT, operacao
dim s_where, s_where_venda, s_where_devolucao, s_where_perdas, s_where_loja
dim s_where_comissao_paga, s_where_comissao_descontada, s_where_st_pagto
dim s_cor, s_sinal, s_cor_sinal
dim s_banco, s_banco_nome, s_agencia, s_conta, s_favorecido
dim s_nome_cliente, s_desempenho_nota
dim s_checked, s_class, s_class_td, idx_bloco, s_new_cab
dim s_lista_completa_venda_normal, s_lista_completa_devolucao, s_lista_completa_perda, s_lista_completa_pedidos, qtde_reg_descontos
dim s_lista_comissao, s_lista_RA_bruto, s_lista_RA_liquido, s_lista_total_comissao, s_lista_total_comissao_arredondado, s_lista_meio_pagto, s_lista_vl_pedido, s_lista_total_RA, s_lista_total_RA_arredondado
dim s_lista_total_RA_arredondado_desc, s_lista_total_RT_arredondado_desc
dim sub_total_comissao, banco, atual,qtdeChq, s_disabled,msg_desconto
dim cod_motivo_desconto, cod_motivo_negativo, total_desconto_planilha
dim vl_sub_total_RT_arredondado,vl_sub_total_RA_arredondado, total_cedulas, cedulas_descricao
dim conta_vendedor, vendedor_processado,sub_total_com_RA,sub_total_com_RT, vl_total_pagto,sub_total_com,sub_total_com_desc,v_desconto_descricao(),v_desconto_valor(),contador,valor_desconto,qtde_registro_desc
dim vl_sub_total_RA_arredondado_desc, vl_sub_total_RT_arredondado_desc, vl_total_pagto_desc
dim vl_RT_desc_aux, vl_RA_desc_aux

    valor_desconto = 0
    cod_motivo_desconto = ""
    cod_motivo_negativo = ""
    s_disabled = ""
    atual = ""
    total_desconto_planilha = ""
    conta_vendedor=0
    vendedor_processado=""
    aviso=""
    operacao=""
    sub_total_com_RA = 0
    sub_total_com_RT = 0
    	if s_where_comissao_descontada <> "" then
		if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		s_where_perdas = s_where_perdas & " (" & s_where_comissao_descontada & ")"
		end if
		
	s = "SELECT vendedor FROM t_COMISSAO_INDICADOR_N2 WHERE (competencia_ano='" & c_dt_entregue_ano & "' AND competencia_mes = '" & c_dt_entregue_mes & "')"
        
    
    for i=Lbound(v_vendedor) to Ubound(v_vendedor)
        rs.Open s, cn
        do while not rs.Eof
                if v_vendedor(i) = rs("vendedor") then 
                    if vendedor_processado <> "" then vendedor_processado = vendedor_processado & ", "
                    vendedor_processado = vendedor_processado & rs("vendedor")
                    conta_vendedor = conta_vendedor+1
                end if
        
        rs.MoveNext
        loop
        if rs.State <> 0 then rs.Close

    next 
    
    if conta_vendedor >= 1 then
        if conta_vendedor = 1 then
            aviso=  "<div class='MtAlerta' style='width:600px;font-weight:bold;' align='center'><p style='margin:5px 2px 5px 2px;'>" & _
                    "O vendedor " & vendedor_processado & " já foi processado no mês de competência informado" & _
                    "<a href='RelComissaoIndicadoresDescConsultaPedidoExec.asp?vendedor=" & vendedor_processado & "&ano_competencia=" & c_dt_entregue_ano & "&mes_competencia=" & c_dt_entregue_mes & "' style='text-decoration:underline;color:#CCC'>" & _
                    "<br />Clique aqui para consultar o relatório gerado.</p></a></div><br />"
        else
            aviso =  "<div class='MtAlerta' style='width:600px;font-weight:bold;' align='center'><p style='margin:5px 2px 5px 2px;'>" & _
                        "Os vendedores " & vendedor_processado & " já foram processados no mês de competência informado</p></div><br />"
        end if
    end if
           
    if aviso <> "" then Response.Write aviso

    if aviso = "" then

        '	CRITÉRIOS COMUNS
		'	PROCESSA SOMENTE OS PARCEIROS NÃO HABILITADOS PARA RECEBEREM A COMISSÃO VIA CARTÃO
	        s_where = " (LEN(Coalesce(t_PEDIDO__BASE.indicador, '')) > 0)" & _
					" AND (t_ORCAMENTISTA_E_INDICADOR.comissao_cartao_status = 0)"

	        if c_vendedor <> "" then
		        if s_where <> "" then s_where = s_where & " AND"
                vendedor_temp = ""
                for j = LBound(v_vendedor) to UBound(v_vendedor)
                    if vendedor_temp <> "" then vendedor_temp = vendedor_temp & " OR"
                    vendedor_temp = vendedor_temp & " (t_PEDIDO__BASE.vendedor = '" & Trim(replace(v_vendedor(j), "'", "''")) & "')"
                next
                if vendedor_temp <> "" then
                    vendedor_temp = " (" & vendedor_temp & ") "
                    s_where = s_where & vendedor_temp
                end if
	        end if

        '	CRITÉRIO: COMISSÃO PAGA
        '	A) VENDAS
	        s_where_comissao_paga = ""
	        s = ""

		        if s <> "" then s = s & " OR"
		        s = s & " (t_PEDIDO.comissao_paga = '0')"

	        if s <> "" then 
		        if s_where_comissao_paga <> "" then s_where_comissao_paga = s_where_comissao_paga & " AND"
		        s_where_comissao_paga = s_where_comissao_paga & " (" & s & ")"
		        end if
		
        '	B) PERDAS/DEVOLUÇÕES
	        s_where_comissao_descontada = ""
	        s = ""

		        if s <> "" then s = s & " OR"
		        s = s & " (comissao_descontada = '0')"
	
	        if s <> "" then 
		        if s_where_comissao_descontada <> "" then s_where_comissao_descontada = s_where_comissao_descontada & " AND"
		        s_where_comissao_descontada = s_where_comissao_descontada & " (" & s & ")"
		        end if

        '	CRITÉRIO: STATUS DE PAGAMENTO
	        s_where_st_pagto = ""
	        s = ""

		        if s <> "" then s = s & " OR"
		        s = s & " ((t_PEDIDO__BASE.st_pagto = 'S') AND (t_PEDIDO__BASE.dt_st_pagto < " & bd_formata_data(str_data) & "))"

	        if s <> "" then 
		        if s_where_st_pagto <> "" then s_where_st_pagto = s_where_st_pagto & " AND"
		        s_where_st_pagto = s_where_st_pagto & " (" & s & ")"
		        end if
	
        '	CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	        s_where_venda = ""
            if c_dt_entregue_mes <> "" And c_dt_entregue_ano <> "" then
		        if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		        s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data < " & bd_formata_data(str_data) & ")"
            end if
		
		
	        if s_where_comissao_paga <> "" then
		        if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		        s_where_venda = s_where_venda & " (" & s_where_comissao_paga & ")"
		        end if

	        if s_where_st_pagto <> "" then
		        if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		        s_where_venda = s_where_venda & " (" & s_where_st_pagto & ")"
		        end if
		
        '	CRITÉRIOS PARA DEVOLUÇÕES
	        s_where_devolucao = ""
                if c_dt_entregue_mes <> "" And c_dt_entregue_ano <> "" then
		        if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		        s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(str_data) & ")"
            end if

	        if s_where_comissao_descontada <> "" then
		        if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		        s_where_devolucao = s_where_devolucao & " (" & s_where_comissao_descontada & ")"
		        end if

        '	CRITÉRIOS PARA PERDAS
	        s_where_perdas = ""
                if c_dt_entregue_mes <> "" And c_dt_entregue_ano <> "" then
		        if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		        s_where_perdas = s_where_perdas & " (t_PEDIDO_PERDA.data < " & bd_formata_data(str_data) & ")"
            end if
		
	        if s_where_comissao_descontada <> "" then
		        if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		        s_where_perdas = s_where_perdas & " (" & s_where_comissao_descontada & ")"
		        end if
			
		
        '	VENDAS NORMAIS
	        s = s_where
	        if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	        s = s & s_where_venda
	        if s <> "" then s = " AND" & s
	        s_sql = "SELECT" & _
			        " '" & VENDA_NORMAL & "' AS operacao," & _
			        " t_PEDIDO.pedido AS id_registro," & _
			        " t_PEDIDO.comissao_paga AS status_comissao," & _
			        " coalesce(t_ORCAMENTISTA_E_INDICADOR.desempenho_nota,'') as desempenho_nota," & _
                    " t_ORCAMENTISTA_E_INDICADOR.banco," & _
			        " t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"

			if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
				s_sql = s_sql & _
						" t_PEDIDO.endereco_nome AS nome," & _
						" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
			else
				s_sql = s_sql & _
						" t_CLIENTE.nome," & _
						" t_CLIENTE.nome_iniciais_em_maiusculas,"
				end if

			s_sql = s_sql & _
			        " t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			        " t_PEDIDO.entregue_data AS data," & _
			        " t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			        " t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			        " t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA," & _
			        " t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
			        " Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.preco_venda) AS total_preco_venda," & _
			        " Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.preco_NF) AS total_preco_NF" & _
			        " FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			        " INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			        " INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			        " LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			        " WHERE (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
			        s & _
			        " GROUP BY t_PEDIDO.pedido, t_PEDIDO.comissao_paga, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"

			if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
				s_sql = s_sql & _
						" t_PEDIDO.endereco_nome," & _
						" t_PEDIDO.endereco_nome_iniciais_em_maiusculas,"
			else
				s_sql = s_sql & _
						" t_CLIENTE.nome," & _
						" t_CLIENTE.nome_iniciais_em_maiusculas,"
				end if

			s_sql = s_sql & _
					" t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO.entregue_data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"

        '	ITENS DEVOLVIDOS
	        s = s_where
	        if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	        s = s & s_where_devolucao
	        if s <> "" then s = " WHERE " & s
	        s_sql = s_sql & " UNION ALL " & _
			        "SELECT" & _
			        " '" & DEVOLUCAO & "' AS operacao," & _
			        " t_PEDIDO_ITEM_DEVOLVIDO.id AS id_registro," & _
			        " t_PEDIDO_ITEM_DEVOLVIDO.comissao_descontada AS status_comissao," & _
			        " coalesce(t_ORCAMENTISTA_E_INDICADOR.desempenho_nota,'') AS desempenho_nota," & _
                    " t_ORCAMENTISTA_E_INDICADOR.banco," & _
			        " t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"

			if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
				s_sql = s_sql & _
					" t_PEDIDO.endereco_nome AS nome," & _
					" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
			else
				s_sql = s_sql & _
					" t_CLIENTE.nome," & _
					" t_CLIENTE.nome_iniciais_em_maiusculas,"
				end if

			s_sql = s_sql & _
			        " t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			        " t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data," & _
			        " t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			        " t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			        " t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA," & _
			        " t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
			        " Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS total_preco_venda," & _
			        " Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_NF) AS total_preco_NF" & _
			        " FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			        " INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido)" & _
			        " INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			        " LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			        s & _
			        " GROUP BY t_PEDIDO_ITEM_DEVOLVIDO.id, t_PEDIDO_ITEM_DEVOLVIDO.comissao_descontada, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"

			if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
				s_sql = s_sql & _
						" t_PEDIDO.endereco_nome," & _
						" t_PEDIDO.endereco_nome_iniciais_em_maiusculas,"
			else
				s_sql = s_sql & _
						" t_CLIENTE.nome," & _
						" t_CLIENTE.nome_iniciais_em_maiusculas,"
				end if

			s_sql = s_sql & _
					" t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"

        '	PERDAS
	        s = s_where
	        if (s <> "") And (s_where_perdas <> "") then s = s & " AND"
	        s = s & s_where_perdas
	        if s <> "" then s = " WHERE " & s
	        s_sql = s_sql & " UNION ALL " & _
			        "SELECT" & _
			        " '" & PERDA & "' AS operacao," & _
			        " t_PEDIDO_PERDA.id AS id_registro," & _
			        " t_PEDIDO_PERDA.comissao_descontada AS status_comissao," & _
			        " coalesce(t_ORCAMENTISTA_E_INDICADOR.desempenho_nota,'') as desempenho_nota," & _
                    " t_ORCAMENTISTA_E_INDICADOR.banco," & _
			        " t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"

			if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
				s_sql = s_sql & _
						" t_PEDIDO.endereco_nome AS nome," & _
						" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
			else
				s_sql = s_sql & _
						" t_CLIENTE.nome," & _
						" t_CLIENTE.nome_iniciais_em_maiusculas,"
				end if

			s_sql = s_sql & _
			        " t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			        " t_PEDIDO_PERDA.data AS data," & _
			        " t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			        " t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			        " t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA," & _
			        " t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
			        " Sum(-t_PEDIDO_PERDA.valor) AS total_preco_venda," & _
			        " Sum(-t_PEDIDO_PERDA.valor) AS total_preco_NF" & _
			        " FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			        " INNER JOIN t_PEDIDO_PERDA ON (t_PEDIDO.pedido=t_PEDIDO_PERDA.pedido)" & _
			        " INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			        " LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			        s & _
			        " GROUP BY t_PEDIDO_PERDA.id, t_PEDIDO_PERDA.comissao_descontada, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"

			if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
				s_sql = s_sql & _
					" t_PEDIDO.endereco_nome," & _
					" t_PEDIDO.endereco_nome_iniciais_em_maiusculas,"
			else
				s_sql = s_sql & _
					" t_CLIENTE.nome," & _
					" t_CLIENTE.nome_iniciais_em_maiusculas,"
				end if

			s_sql = s_sql & _
					" t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_PERDA.data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"

	        s_sql = "SELECT " & _
				        "*" & _
			        " FROM (" & _
				        s_sql & _
				        ") t" & _
			        " ORDER BY t.vendedor, t.desempenho_nota, t.indicador, t.numero_loja, t.data, t.pedido, t.total_preco_venda DESC"

            ' CABEÇALHO
	        cab_table = "<table border='0' cellspacing='0' id='tableDados'>" & chr(13)
	        cab = "	<tr style='background:azure' nowrap>" & chr(13) & _
		            "		<td class='MDTE tdLoja' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>Loja</span></td>" & chr(13) & _
		            "		<td class='MTD tdOrcamento' align='left' valign='bottom' nowrap><span class='R VISAO_ANALIT'>Nº Orçam</span></td>" & chr(13) & _
		            "		<td class='MTD tdPedido' align='left' valign='bottom' nowrap><span class='R VISAO_ANALIT'>Nº Pedido</span></td>" & chr(13) & _
		            "		<td class='MTD tdData' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>Data</span></td>" & chr(13) & _
		            "		<td class='MTD tdVlPedido' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL Pedido</span></td>" & chr(13) & _
		            "		<td class='MTD tdVlRT' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>COM (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		            "		<td class='MTD tdVlRABruto' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Bruto (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		            "		<td class='MTD tdVlRALiq' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Líq (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		            "		<td class='MTD tdVlRADif' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Dif (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		            "		<td class='MTD tdStPagto' align='left' valign='bottom'><span class='R VISAO_ANALIT' style='font-weight:bold;'>St Pagto</span></td>" & chr(13) & _
		            "		<td class='MTD tdSinal' align='center' valign='bottom'><span class='Rc VISAO_ANALIT' style='font-weight:bold;'>+/-</span></td>" & chr(13) & _
		            "		<td valign='bottom' class='notPrint BkgWhite' align='left'>&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & "_NNNNN_" & chr(34) & ");' title='exibe ou oculta os dados'><img src='../botao/view_bottom.png' border='0'></a></td>" & chr(13) & _
		            "	</tr>" & chr(13)
	
	        x="<BR>" & chr(13)
	        n_reg = 0
	        n_reg_total = 0
	        idx_bloco = 0
	        qtde_indicadores = 0
	        vl_sub_total_preco_venda = 0
	        vl_total_preco_venda = 0
	        vl_sub_total_preco_NF = 0
	        vl_total_preco_NF = 0
	        vl_sub_total_RT = 0
	        vl_total_RT = 0
	        vl_sub_total_RA = 0
	        vl_total_RA = 0
	        vl_sub_total_RA_liquido = 0
	        vl_total_RA_liquido = 0
	        vl_sub_total_RA_diferenca = 0
	        vl_total_RA_diferenca = 0
	        s_lista_completa_venda_normal = ""
	        s_lista_completa_devolucao = ""
	        s_lista_completa_perda = ""
            s_lista_completa_pedidos = ""
            s_lista_comissao = ""
            s_lista_RA_bruto = ""
            s_lista_RA_liquido = ""
            s_lista_vl_pedido = ""
            s_lista_total_comissao = ""
            s_lista_total_comissao_arredondado = ""
            s_lista_total_RA = ""
            s_lista_total_RA_arredondado = ""
            s_lista_total_RT_arredondado_desc = ""
            s_lista_total_RA_arredondado_desc = ""
            qtde_reg_descontos = ""
            total_cedulas = ""
            cedulas_descricao = ""
            sub_total_comissao = 0
            vl_sub_total_RT_arredondado = 0
            vl_sub_total_RT_arredondado = 0
	        indicador_a = "XXXXXXXXXXXX"
            sub_total_com = 0
            sub_total_com_desc = 0
	        set r = cn.execute(s_sql)

        dim totalComDin,totalComChqOutros,totalComchqBradesco
            totalComDin= 0
            totalComChqOutros=0
            totalComchqBradesco=0
     

            limitador(0)= "10"
            limitador(1)= "10"
            limitador(2)= "10"
            limitador(3)= "10"
            limitador(4)= "10"
            limitador(5)= "10"

	        do while Not r.Eof
    
	        '	MUDOU DE INDICADOR?
	        if Trim("" & r("indicador"))<>indicador_a then
			        indicador_a = Trim("" & r("indicador"))
			        idx_bloco = idx_bloco + 1
			        qtde_indicadores = qtde_indicadores + 1      

                '-   CALCULO CÉDULAS
            dim j, z,totalcomissao
            z = 0
            j = 0

                if sub_total_com_desc <= 0 Or sub_total_com_desc > 300 Or (banco = "237" Or banco = "077") then
                        if (sub_total_com_desc  >0 Or sub_total_com_desc > 300) And (banco <> "237" Or banco <> "077") then 
                            totalComChqOutros = totalComChqOutros + Cstr(sub_total_com_desc)                       
                        elseif sub_total_com_desc <0 then 
           
                        else
                        totalComchqBradesco = totalComchqBradesco + Cstr(sub_total_com_desc)
                        end if
                        
                 else                       
                        dadosCalculo = Clng(sub_total_com_desc)                                                                 
                        dadosCalculo = o.CalculaCedulas(dadosCalculo,"2#"& limitador(5) &"|5#"&limitador(4)&"|10#"&limitador(3)&"|20#"&limitador(2)&"|50#"&limitador(1)&"|100#"&limitador(0)&"",resultadoCalculo)                        
                        dadosCalculo = resultadocalculo
                        v_cedulas = Split(dadosCalculo,"|")   
                        for cont=0 to Ubound(v_cedulas)
                            if cont mod 2 = 0 then
                                redim  preserve cedulas(j)
                                ' cedulas(j)=  cint(v_cedulas(cont))
                                j = j +1
                            else
                                redim preserve qtdeCedula(z)
                                qtdeCedula(z) = Clng(v_cedulas(cont))
                                z = z + 1                   
                            end if
                        next 
                    
                 end if
                
		            ' FECHA TABELA DO INDICADOR ANTERIOR
			        if n_reg_total > 0 then 
                        s_lista_total_comissao = s_lista_total_comissao & vl_sub_total_RT & ";"
                        s_lista_total_RA = s_lista_total_RA & vl_sub_total_RA_liquido & ";"
                        if rs2.State <> 0 then rs2.Close
                        rs2.Open "SELECT COUNT(*) qtde_Desconto, descricao,valor,ordenacao  FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (apelido = '" & ind_anterior & "') GROUP BY  descricao,valor,ordenacao ORDER BY ordenacao", cn
                        
                        msg_desconto = ""
                        if sub_total_com_desc >= 0 then 
                            s_checked=" checked"      
                        else  
                            s_disabled = " disabled"
                            cod_motivo_negativo = cod_motivo_negativo & ind_anterior & ", "
                        end if

                        if Not rs2.Eof then

                                valor_desconto = 0
                                contador = 0
                                qtde_registro_desc = 0
                                Erase v_desconto_descricao
                                Erase v_desconto_valor
                                do while Not rs2.EoF         
                                    redim preserve v_desconto_descricao(contador) 
                                    redim preserve  v_desconto_valor(contador)                 
                                    v_desconto_descricao(contador) = rs2("descricao")
                                    v_desconto_valor(contador) = rs2("valor")
                                    valor_desconto = valor_desconto + v_desconto_valor(contador)
                                    qtde_registro_desc = qtde_registro_desc + 1
                                    contador = contador + 1
                                    rs2.MoveNext
		                        loop
                                s_checked=""
                                s_disabled = " disabled"
                                msg_desconto =  "Planilha de Desconto:&nbsp; " & Cstr(qtde_registro_desc) & " registro(s) no valor total de&nbsp;" & formata_moeda(valor_desconto)
                                cod_motivo_desconto = cod_motivo_desconto & ind_anterior & ", "
                                total_desconto_planilha = total_desconto_planilha & converte_numero(valor_desconto) & ";"
                                qtde_reg_descontos = qtde_reg_descontos & qtde_registro_desc & ";"
                            else
                                msg_desconto= ""
                        end if
                    ' TOTAL DO INDICADOR
				        s_cor="black"
				        if vl_sub_total_preco_venda < 0 then s_cor="red"
				        if vl_sub_total_RT < 0 then s_cor="red"
				        if vl_sub_total_RA < 0 then s_cor="red"
				        if vl_sub_total_RA_liquido < 0 then s_cor="red"
				        x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
						        "		<td class='MTBE' colspan='4' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
						        "TOTAL:</span></td>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _
                                "       <input type='hidden' id='sub_total_preco_venda_" & idx_bloco - 1 & "' value='" & vl_sub_total_preco_venda & "'>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _
                                "       <input type='hidden' id='sub_total_comissao_" & idx_bloco - 1 & "' value='" & vl_sub_total_RT & "'>" & chr(13) 
                                
                                if vl_sub_total_RA_arredondado >0 and vl_sub_total_RT_arredondado < 0then
                                    x = x & "       <input type='hidden' id='sub_total_comissaoRT_" & idx_bloco - 1 & "' value='" & Clng(0) & "'>" & chr(13)
                                else if vl_sub_total_RA_arredondado >=0 and vl_sub_total_RT_arredondado >= 0 then
                                    x = x & "       <input type='hidden' id='sub_total_comissaoRT_" & idx_bloco  - 1& "' value='" & Clng(vl_sub_total_RT_arredondado_desc) & "'>" & chr(13) 
                                else
                                    x = x & "       <input type='hidden' id='sub_total_comissaoRT_" & idx_bloco - 1 & "' value='" & Clng(sub_total_com_desc) & "'>" & chr(13) 
                                end if
                                end if                               
                                x = x &_
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</span></td>" & chr(13) & _
                                "       <input type='hidden' id='sub_total_RA_" & idx_bloco - 1 & "' value='" & vl_sub_total_RA & "'>" & chr(13) 
                                if vl_sub_total_RA_arredondado < 0 and vl_sub_total_RT_arredondado > 0then
                                    x = x & "       <input type='hidden' id='sub_total_comissaoRA_" & idx_bloco  - 1& "' value='" & Clng(0) & "'>" & chr(13)
                                else if vl_sub_total_RA_arredondado >=0 and vl_sub_total_RT_arredondado >= 0 then
                                    x = x & "       <input type='hidden' id='sub_total_comissaoRA_" & idx_bloco  - 1& "' value='" & Clng(vl_sub_total_RA_arredondado_desc) & "'>" & chr(13) 
                                else   
                                     x = x & "       <input type='hidden' id='sub_total_comissaoRA_" & idx_bloco  - 1& "' value='" & Clng(sub_total_com_desc) & "'>" & chr(13)
                                end if
                                end if
                                x = x &_                                
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _
                                "       <input type='hidden'  id='sub_total_RA_liquido_" & idx_bloco - 1 & "' value='" & vl_sub_total_RA_liquido & "'>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _
                                "       <input type='hidden'  id='sub_total_RA_diferenca_" & idx_bloco - 1 & "' value='" & vl_sub_total_RA_diferenca & "'>" & chr(13) & _
						        "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
						        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
						        "	</tr>" & chr(13) & _ 
		                        "   <tr>" & chr(13) 
                         ' MENSAGEM DE DESCONTO E ARREDONDAMENTO DO RT E RA
    
                                if sub_total_com_desc = 0 Or sub_total_com_desc > 300 Or (banco = "237" Or banco = "077") then
                                    x = x & "       <td align='left' colspan='5' nowrap>" & chr(13)
                                    if msg_desconto <> "" then x = x & "<span class='Cd'><a href='javascript:abreDesconto(" & idx_bloco -1 & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"& chr(13)
                                    if vl_sub_total_RT_arredondado >=0  then 
                                        s_cor = "black "
                                    end if   
                                    x = x &             "<td align='right' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RT_arredondado)&"</span>" & chr(13) 
                                    if vl_sub_total_RA_arredondado >=0  then 
                                        s_cor = "black "
                                    else
                                        s_cor = "red"
                                    end if       
                                    x = x &             "            <td align='right' colspan='2' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RA_arredondado)&"</span>" & chr(13)
                                       
                                elseif sub_total_com_desc < 301 and sub_total_com_desc >= 0 then
                                        x = x & "       <td align='left' colspan='5' nowrap>"& chr(13)
                                        if msg_desconto <> "" then 
                                        x = x & "<span class='Cd'><a href='javascript:abreDesconto(" & idx_bloco -1 & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"& chr(13)
                                        else   
                                x = x & "<span class='Rd' style='color: black;'>Cédulas: "& chr(13) 
                            ' QUANTIDADE DE CEDULAS                
                                for cont = 0 to UBound(qtdeCedula)
                                        if (cont = 0 And qtdeCedula(cont) <> 0) then                 
                                            if (qtdeCedula(cont) > 1) then
                                                x = x & qtdeCedula(cont) & "&times;"
                                                cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                            end if
                                            cedulas_descricao = cedulas_descricao & formata_moeda("100")
                                            x = x & formata_moeda("100") & chr(13) & _
                                        "       <input type='hidden'  id='cedulas100_" & idx_bloco -1 & "' value='100'>" & chr(13) & _
                                        "       <input type='hidden'  id='total_cedulas100_" & idx_bloco -1 & "' value='" & qtdeCedula(cont) & "'>" & chr(13)
                                            if (qtdeCedula(1) <> 0 Or qtdeCedula(2) <> 0 Or qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then
                                            x = x & " + "
                                            cedulas_descricao = cedulas_descricao & " + "
                                            end if
                          
                                    elseif (cont = 1 And qtdeCedula(cont) <> 0) then
                                        if (qtdeCedula(cont) > 1) then
                                            x = x & qtdeCedula(cont) & "&times;"
                                            cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                        end if
                                            cedulas_descricao = cedulas_descricao & formata_moeda("50")
                                            x = x & formata_moeda("50")  & _
                                            "       <input type='hidden'  id='cedulas50_" & idx_bloco -1 & "' value='50'>" & chr(13) & _
                                            "       <input type='hidden'  id='total_cedulas50_" & idx_bloco -1 & "' value='" & qtdeCedula(cont) & "'>" & chr(13) 
                                        if (qtdeCedula(2) <> 0 Or qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then
                                            x = x & " + "
                                            cedulas_descricao = cedulas_descricao & " + "
                                        end if
                           
                                    elseif (cont = 2 And qtdeCedula(cont) <> 0) then
                                        if (qtdeCedula(cont) > 1) then
                                            x = x & qtdeCedula(cont) & "&times;"  
                                            cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                        end if
                                        cedulas_descricao = cedulas_descricao & formata_moeda("20")
                                            x = x & formata_moeda("20")  & _
                                    "       <input type='hidden'  id='cedulas20_" & idx_bloco -1 & "' value='20'>" & chr(13) & _
                                    "       <input type='hidden'  id='total_cedulas20_" & idx_bloco -1 & "' value='" & qtdeCedula(2) & "'>" & chr(13) 
                                        if (qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then
                                            x = x & " + "
                                            cedulas_descricao = cedulas_descricao & " + "
                                        end if
                        
                                    elseif (cont = 3 And qtdeCedula(cont) <> 0) then
                                        if (qtdeCedula(cont) > 1) then
                                            x = x & qtdeCedula(cont) & "&times;"
                                            cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                        end if
                                        cedulas_descricao = cedulas_descricao & formata_moeda("10")
                                            x = x & formata_moeda("10") & _
                                    "       <input type='hidden'  id='cedulas10_" & idx_bloco -1 & "' value='10'>" & chr(13) & _
                                    "       <input type='hidden'  id='total_cedulas10_" & idx_bloco -1 & "' value='" & qtdeCedula(3) & "'>" & chr(13) 
                                        if (qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then
                                            x = x & " + "
                                            cedulas_descricao = cedulas_descricao & " + "
                                        end if

                                    elseif (cont = 4 And qtdeCedula(cont) <> 0) then
                                        if (qtdeCedula(cont) > 1) then
                                            x = x & qtdeCedula(cont) & "&times;"
                                            cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                        end if
                                            cedulas_descricao = cedulas_descricao & formata_moeda("5")
                                            x = x & formata_moeda("5") & _
                                    "       <input type='hidden'  id='cedulas5_" & idx_bloco -1 & "' value='5'>" & chr(13) & _
                                    "       <input type='hidden'  id='total_cedulas5_" & idx_bloco -1 & "' value='" & qtdeCedula(cont) & "'>" & chr(13) 
                                        if (qtdeCedula(5) <> 0) then 
                                            x = x & " + "
                                            cedulas_descricao = cedulas_descricao & " + "
                                        end if
                       
                                    elseif (cont = 5 And qtdeCedula(cont) <> 0) then
                                        if (qtdeCedula(cont) > 1) then
                                            x = x & qtdeCedula(cont) & "&times;"
                                            cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                        end if
                                        cedulas_descricao = cedulas_descricao & formata_moeda("2")
                                            x = x & formata_moeda("2") & _
                                    "       <input type='hidden'  id='cedulas2_" & idx_bloco -1 & "' value='2'>" & chr(13) & _
                                    "       <input type='hidden'  id='total_cedulas2_" & idx_bloco -1 & "' value='" & qtdeCedula(cont) & "'>" & chr(13) 
                                    end if
                                    total_cedulas = total_cedulas & qtdeCedula(cont)
                                    if cont < Ubound(qtdeCedula) then total_cedulas = total_cedulas & ","
                    
                                next
                                    cedulas_descricao = cedulas_descricao & ";"
                                    total_cedulas = total_cedulas & ";"
                   
                            end if
                        ' COMISSAO RT E RA
                                if vl_sub_total_RT_arredondado >=0 then 
                                    s_cor = "black "
                                end if                               
                                    x = x & " <td align='right' nowrap><span class='Rd' style='color:" & s_cor & ";margin-right:0.3;'>"& formata_moeda(vl_sub_total_RT_arredondado)&"</span>"& chr(13) 
                                if vl_sub_total_RA_arredondado >=0 then 
                                    s_cor = "black "
                                else
                                    s_cor = "red"
                                end if                               
                                x = x &  " <td align='right' colspan='2' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RA_arredondado)& "</span>"& chr(13) 
                                     
                            end if
                                if sub_total_com_desc >= 0 then 
                                s_cor = "black"
                                else 
                                s_cor = "red"
                                end if
                                x = x & "       <td align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>"

                        ' MEIO DE PAGAMENTO
                                             
                            if sub_total_com_desc <= 0 then
                                s_lista_meio_pagto = s_lista_meio_pagto & " " & ";"
                                x = x & "&nbsp;</span></td>" & chr(13)
                                if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                vl_total_pagto = vl_total_pagto & "0"
                            elseif sub_total_com_desc > 300 And (banco <> "237" And banco <> "077") then                           
                                        s_lista_meio_pagto = s_lista_meio_pagto & "CHQ" & ";"
                                        x = x & "CHQ:</span></td>" & chr(13) &_
                                        "       <input type='hidden'  id='forma_pag_" & idx_bloco - 1 & "' value='CHQ'>" & chr(13) 
                                    if vl_sub_total_RT_arredondado >0 and vl_sub_total_RA_arredondado < 0 then
                                        x = x &  "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco - 1 & "' value='"& sub_total_com_desc & "'>" & chr(13) 
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                        vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                    elseif vl_sub_total_RA_arredondado >0 and vl_sub_total_RT_arredondado < 0 then
                                        x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco - 1 & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                        vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                    else
                                        x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco - 1 & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                        vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                    end if
                            
                            elseif (banco = "237" Or banco = "077") then
                                    if banco = "237" Then
                                        s_lista_meio_pagto = s_lista_meio_pagto & "DEP" & ";"
                                    ElseIf banco = "077" Then
                                        s_lista_meio_pagto = s_lista_meio_pagto & "DEP1" & ";"
                                    End if
                                    x = x & "DEP:</span></td>" & chr(13) &_
                                    "       <input type='hidden'  id='forma_pag_" & idx_bloco - 1 & "' value='DEP'>" & chr(13) 
                                        if vl_sub_total_RT_arredondado >0 and vl_sub_total_RA_arredondado < 0 then
                                            x = x &  "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco - 1 & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                            if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                            vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                        elseif vl_sub_total_RA_arredondado >0 and vl_sub_total_RT_arredondado < 0 then
                                            x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco - 1 & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                            if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                            vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                        else
                                            x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco - 1 & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                            if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                            vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                        end if
                            else 
                            s_lista_meio_pagto = s_lista_meio_pagto & "DIN" & ";"
                            
                            x = x & "DIN:</span></td>" & chr(13) &_
                            "       <input type='hidden'  id='forma_pag_" & idx_bloco - 1 & "' value='DIN'>" & chr(13) 
                               if vl_sub_total_RT_arredondado >0 and vl_sub_total_RA_arredondado < 0 then
                                    x = x &  "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco - 1 & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                    if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                    vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                elseif vl_sub_total_RA_arredondado >0 and vl_sub_total_RT_arredondado < 0 then
                                    x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco - 1 & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                    if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                    vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                else
                                    x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco - 1 & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                    if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                    vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                    end if

                            end if
   
                        ' SUB TOTAL DA COMISSAO ARREDONDADO COM DESCONTO
                            if sub_total_com_desc >=0  then                               
                                x = x & "       <td align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(sub_total_com_desc) & "</span></td>" & chr(13)        
                                
                            end if
                            
                            x = x &"	</tr>" & chr(13) 
                    
                            ' SUB TOTAL DA COMISSAO ARREDONDADO
                            if sub_total_com >=0  then                               
                                x = x & "       <td align='right' colspan='11'><span class='Cd' style='color:gray;'>Sem desconto: " & formata_moeda(sub_total_com) & "</span></td>" & chr(13)        
                                
                            end if

                            if msg_desconto <> "" then
                            x = x &"   <tr>" & chr(13)& _
                                   "          <td  class='table_Desconto' id='table_Desconto_"& idx_bloco -1  &"'"" colspan='15' >" & chr(13)& _
                                   "          <table colspan='2' align='left' >"& chr(13)
                                   for contador = 0 to Ubound(v_desconto_descricao)                                 
                                       x = x & "   <tr>" & chr(13)& _
                                                "       <td width='15'>&nbsp;</td>" & chr(13)& _
                                                "       <td  align='left' width='400' ><span class='Cd'style='color: red;' >"& v_desconto_descricao(contador)& "</span></td>"& _
                                                "       <td align='left' ><span class='Cd'style='color: red;' > R$ "& formata_moeda(v_desconto_valor(contador))& "</span></td>"& _
                                                "   </tr>"                                     
                                   next          
                                   x = x & "   </table>"& chr(13)& _
                                           "   </td>"& chr(13)& _
                                           "</tr>"
                            end if
        
						        x = x &     "   </table>" & chr(13)
                    
                                x = "<table cellpadding='0' cellspacing='0'><tr><td valign='top'><br />" & chr(13) & _
                                    "   <input type='checkbox' name='ckb_com_pg' class='CKB_COM' id='ckb_comissao_paga_tit_bloco_" & idx_bloco -1 & "' onclick='trata_ckb_onclick();calculaTotalComissao();alternaCheck(" & idx_bloco -1 & ");' value='" & atual & "' " & s_checked & s_disabled & " />" & chr(13) & _ 
                                    "   <input type='checkbox' style='display:none' name='ckb_com_pg_i' id='ckb_comissao_paga_tit_bloco_indicador_" & idx_bloco -1 & "' value='" & ind_anterior & "' />" & chr(13) & _
                                    "</td><td valign='top'>" & x & "</td></tr></table>" & chr(13)
                            s_lista_total_comissao_arredondado = s_lista_total_comissao_arredondado & vl_sub_total_RT_arredondado & ";"
                            s_lista_total_RA_arredondado = s_lista_total_RA_arredondado & vl_sub_total_RA_arredondado & ";"
                            s_lista_total_RT_arredondado_desc = s_lista_total_RT_arredondado_desc & vl_sub_total_RT_arredondado_desc & ";"
                            s_lista_total_RA_arredondado_desc = s_lista_total_RA_arredondado_desc & vl_sub_total_RA_arredondado_desc & ";"
             
                        atual = ""
                        s_checked=""
                        s_disabled=""
				        Response.Write x
				        x="<BR>" & chr(13)
				        end if
    
                    s_sql = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido = '" & Trim("" & r("indicador")) & "')"
			        if rs.State <> 0 then rs.Close
			        rs.Open s_sql, cn
			        if Not rs.Eof then
				        s_banco = Trim("" & rs("banco"))
				        s_agencia = Trim("" & rs("agencia"))
				        s_conta = Trim("" & rs("conta"))
				        s_favorecido = Trim("" & rs("favorecido"))
				        s_banco_nome = x_banco(s_banco)
				        if (s_banco <> "") And (s_banco_nome <> "") then s_banco = s_banco & " - " & s_banco_nome
			        else
				        s_banco = ""
				        s_banco_nome = ""
				        s_agencia = ""
				        s_conta = ""
				        s_favorecido = ""
				        end if

                    x = x & Replace(cab_table, "tableDados", "tableDados_" & idx_bloco)
                    x = x & "	<tr>" & chr(13)
            
                    s = Trim("" & r("indicador"))
			        s_aux = x_orcamentista_e_indicador(s)
			        if (s<>"") And (s_aux<>"") then s = s & " - "
			        s = s & s_aux
			
			        if s <> "" then
						x = x & "		<td class='MDTE' colspan='11' align='left' valign='bottom' class='MB' style='background:azure;'><span class='N'>&nbsp;" & s_desempenho_nota & s & "</span></td>" & chr(13) & _
								"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
								"	</tr>" & chr(13)
						end if

					x = x & _
							"	<tr>" & chr(13) & _
							"		<td class='MDTE' colspan='11' align='left' valign='bottom' class='MB' style='background:white;'>" & chr(13) & _
							"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
							"				<tr>" & chr(13) & _
							"					<td align='right' valign='bottom' nowrap><span class='Cn'>Pagamento da Comissão via NFSe:</span></td>" & chr(13) & _
							"					<td width='90%' align='left' valign='bottom' nowrap><span class='Cn'>" & chr(13)

					if Trim("" & rs("comissao_NFSe_cnpj")) <> "" then
						x = x & cnpj_cpf_formata(Trim("" & rs("comissao_NFSe_cnpj"))) & " - " & Trim("" & rs("comissao_NFSe_razao_social"))
					else
						x = x & "N.I."
						end if

					x = x & _
										"</span></td>" & chr(13) & _
										"				</tr>" & chr(13) & _
										"			</table>" & chr(13) & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13)

					x = x & _
								"	<tr>" & chr(13) & _
								"		<td class='MDTE' colspan='11' align='left' valign='bottom' class='MB' style='background:whitesmoke;'>" & chr(13) & _
								"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
                                       
								"				<tr>" & chr(13) & _
								"					<td colspan='3' align='left' valign='bottom' style='vertical-align:middle'><div valign='bottom' style='height:14px;max-height:14px;overflow:hidden;vertical-align:middle'><span class='Cn'>Banco: " & rs("banco") & " - " & x_banco(rs("banco")) &  "</span></div></td>" & chr(13) & _
								"				</tr>" & chr(13) & _
								"				<tr>" & chr(13) & _
								"					<td class='MTD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>Agência: " & rs("agencia")
                    if Trim("" & rs("agencia_dv")) <> "" then
                        x = x & "-" & rs("agencia_dv") & chr(13)
                    end if
    
                    x = x & "</span></td>" & chr(13) & _
								"					<td class='MC MD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>"

                    if Trim("" & rs("tipo_conta")) <> "" then
                        if rs("tipo_conta") = "P" then
                            x = x & "C/P: "
                        elseif rs("tipo_conta") = "C" then
                            x = x & "C/C: "
                        end if
                    else
                        x = x & "Conta: "
                    end if

                    if Trim("" & rs("conta_operacao")) <> "" then
                        x = x & rs("conta_operacao") & "-"
                    end if               
    
                    x = x & rs("conta")
    
                    if Trim("" & rs("conta_dv")) <> "" then
                        x = x & "-" & rs("conta_dv") & chr(13)
                    end if

					x = x & "					<td class='MC' width='60%' align='left' valign='bottom'><span class='Cn'>Favorecido: " & s_favorecido & "</span></td>" & chr(13) & _
								"				</tr>" & chr(13) & _
								"			</table>" & chr(13) & _
								"		</td>" & chr(13) & _
								"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
								"	</tr>" & chr(13)
			        s_new_cab = Replace(cab, "ckb_comissao_paga_tit_bloco", "ckb_comissao_paga_tit_bloco_" & idx_bloco)
			        s_new_cab = Replace(s_new_cab, "trata_ckb_onclick();", "trata_ckb_onclick(" & chr(34) & idx_bloco & chr(34) & ");")
			        s_new_cab = Replace(s_new_cab, "_NNNNN_", CStr(idx_bloco))
			        x = x & s_new_cab

			        n_reg = 0
			        vl_sub_total_preco_venda = 0
			        vl_sub_total_preco_NF = 0
			        vl_sub_total_RT = 0
			        vl_sub_total_RA = 0
			        vl_sub_total_RA_liquido = 0
			        vl_sub_total_RA_diferenca = 0

	        end if
   
            if atual <> "" then atual = atual & ", "
                atual = atual & Trim("" & r("pedido"))

	            ' CONTAGEM
		        n_reg = n_reg + 1
		        n_reg_total = n_reg_total + 1
                banco = rs("banco")
            ' CÁLCULOS

            '	EVITA DIFERENÇAS DE ARREDONDAMENTO
		        vl_preco_venda = converte_numero(formata_moeda(r("total_preco_venda")))
		        vl_preco_NF = converte_numero(formata_moeda(r("total_preco_NF")))
		        perc_RT = r("perc_RT")
                
		        vl_RT = (perc_RT/100) * vl_preco_venda
                
		        vl_RA = vl_preco_NF - vl_preco_venda
                
		        if Not calcula_total_RA_liquido(r("perc_desagio_RA_liquida"), vl_RA, vl_RA_liquido) then
			        Response.Write "FALHA AO CALCULAR O RA LÍQUIDO."
			        Response.End 
			        end if
		
		        vl_RA_diferenca = vl_RA - vl_RA_liquido

            ' CÁLCULOS DE SUB TOTAL
                vl_sub_total_preco_venda = vl_sub_total_preco_venda + r("total_preco_venda")
		        vl_total_preco_venda = vl_total_preco_venda + r("total_preco_venda")
		        vl_sub_total_preco_NF = vl_sub_total_preco_NF + r("total_preco_NF")
		        vl_total_preco_NF = vl_total_preco_NF + r("total_preco_NF")
		        vl_sub_total_RT = vl_sub_total_RT +  vl_RT
		        vl_total_RT = vl_total_RT + vl_RT        
		        vl_sub_total_RA = vl_sub_total_RA +  vl_RA      
		        vl_total_RA = vl_total_RA + vl_RA
		        vl_sub_total_RA_liquido = vl_sub_total_RA_liquido + vl_RA_liquido
		        vl_total_RA_liquido = vl_total_RA_liquido + vl_RA_liquido
                vl_sub_total_RA_diferenca = vl_sub_total_RA_diferenca + vl_RA_diferenca
		        vl_total_RA_diferenca = vl_total_RA_diferenca + vl_RA_diferenca
                sub_total_comissao = vl_sub_total_RT + vl_sub_total_RA_liquido
                vl_RT_desc_aux = converte_numero(vl_sub_total_RT - (vl_sub_total_RT * (COMISSAO_INDICADOR_PERC_DESCONTO_SEM_NF / 100)))
                vl_RA_desc_aux = converte_numero(vl_sub_total_RA_liquido - (vl_sub_total_RA_liquido * (COMISSAO_INDICADOR_PERC_DESCONTO_SEM_NF / 100)))

                if ((vl_RT_desc_aux >= 0 and vl_RA_desc_aux >= 0) AND (((vl_RT_desc_aux + vl_RA_desc_aux) < 301) AND (banco <> "237" Or banco <> "077"))) then 
                    vl_sub_total_RT_arredondado = Clng(o.digitoFinal(formata_moeda(vl_sub_total_RT)))
                    vl_sub_total_RA_arredondado = Clng(o.digitoFinal(formata_moeda(vl_sub_total_RA_liquido)))
                    vl_sub_total_RT_arredondado_desc =Clng(o.digitoFinal(formata_moeda(vl_sub_total_RT - (vl_sub_total_RT * (COMISSAO_INDICADOR_PERC_DESCONTO_SEM_NF / 100)))))
                    vl_sub_total_RA_arredondado_desc = Clng(o.digitoFinal(formata_moeda(vl_sub_total_RA_liquido - (vl_sub_total_RA_liquido * (COMISSAO_INDICADOR_PERC_DESCONTO_SEM_NF / 100)))))
                else
                     vl_sub_total_RT_arredondado = floor(vl_sub_total_RT)
                     vl_sub_total_RA_arredondado = floor(vl_sub_total_RA_liquido)  
                    vl_sub_total_RT_arredondado_desc =floor(vl_sub_total_RT - (vl_sub_total_RT * (COMISSAO_INDICADOR_PERC_DESCONTO_SEM_NF / 100)))
                    vl_sub_total_RA_arredondado_desc = floor(vl_sub_total_RA_liquido - (vl_sub_total_RA_liquido * (COMISSAO_INDICADOR_PERC_DESCONTO_SEM_NF / 100)))
                end if
               
                if vl_sub_total_RT_arredondado >=0 or vl_sub_total_RA_arredondado >=0 then
                    if vl_sub_total_RT_arredondado < 0 or vl_sub_total_RA_arredondado <0 then
                        sub_total_com = o.digitoFinal(vl_sub_total_RT+vl_sub_total_RA_liquido)
                        sub_total_com_desc = vl_sub_total_RT_arredondado_desc+vl_sub_total_RA_arredondado_desc
                        'converte_numero(sub_total_com_desc - (sub_total_com_desc * (COMISSAO_INDICADOR_PERC_DESCONTO_SEM_NF / 100))))
                    else 
                        sub_total_com = vl_sub_total_RT_arredondado + vl_sub_total_RA_arredondado
                        sub_total_com_desc = vl_sub_total_RT_arredondado_desc + vl_sub_total_RA_arredondado_desc
                    end if
                end if
                        
                if vl_sub_total_RT <= 0 and vl_sub_total_RA_liquido <= 0 then 
                    sub_total_com = vl_sub_total_RT_arredondado + vl_sub_total_RA_arredondado
                    sub_total_com_desc = vl_sub_total_RT_arredondado_desc + vl_sub_total_RA_arredondado_desc
                    'sub_total_com_desc = o.digitoFinal(formata_moeda(sub_total_com_desc - (sub_total_com_desc * (COMISSAO_INDICADOR_PERC_DESCONTO_SEM_NF / 100))))
                end if

                if Trim("" & r("operacao")) = VENDA_NORMAL then
			        if s_lista_completa_venda_normal <> "" then s_lista_completa_venda_normal = s_lista_completa_venda_normal & ";"
			        s_lista_completa_venda_normal = s_lista_completa_venda_normal & Trim("" & r("id_registro"))
			    
		        elseif Trim("" & r("operacao")) = DEVOLUCAO then
			        if s_lista_completa_devolucao <> "" then s_lista_completa_devolucao = s_lista_completa_devolucao & ";"
			        s_lista_completa_devolucao = s_lista_completa_devolucao & Trim("" & r("id_registro"))
			        s_class = s_class & " CKB_COM_DEV"
			    
		        elseif Trim("" & r("operacao")) = PERDA then
			        if s_lista_completa_perda <> "" then s_lista_completa_perda = s_lista_completa_perda & ";"
			        s_lista_completa_perda = s_lista_completa_perda & Trim("" & r("id_registro"))
			        s_class = s_class & " CKB_COM_PERDA"
			    
			        end if

                if operacao <> "" then operacao = operacao & ", "
                operacao = operacao & r("operacao")
                if s_lista_completa_pedidos <> "" then s_lista_completa_pedidos = s_lista_completa_pedidos & ";"
			        s_lista_completa_pedidos = s_lista_completa_pedidos & Trim("" & r("pedido"))

                '> CHECK BOX
	            '	É USADO O CÓDIGO DA OPERAÇÃO (VENDA NORMAL, DEVOLUÇÃO, PERDA) P/ NÃO CORRER O RISCO DE HAVER CONFLITO DEVIDO A ID'S REPETIDOS ENTRE AS OPERAÇÕES
		        s_class = " CKB_COM_BL_" & idx_bloco
		        s_class_td = ""

		        x = x & "	<tr nowrap class='VISAO_ANALIT'>"  & chr(13)
		
		        if (vl_preco_venda < 0) Or (vl_RT < 0) Or (vl_RA < 0) Or (vl_RA_liquido < 0) then
			        s_cor = "red"
			        s_cor_sinal = "red"
			        s_sinal = "-"
		        else
			        s_cor = "black"
			        s_cor_sinal = "green"
			        s_sinal = "+"
			        end if
		
	            '> LOJA
		        x = x & "		<td class='MDTE tdLoja' align='center'><span class='Cnc' style='color:" & s_cor & ";'>" & Trim("" & r("loja")) & "</span></td>" & chr(13)

	            '> Nº ORÇAMENTO
		        s = Trim("" & r("orcamento"))
		        if s = "" then s = "&nbsp;"
		        x = x & "		<td class='MTD tdOrcamento' align='left'><span class='Cn'><a style='color:" & s_cor & ";' href='javascript:fORCConsulta(" & _
				        chr(34) & s & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o orçamento'>" & _
				        s & "</a></span></td>" & chr(13)

	            '> Nº PEDIDO
		        s_nome_cliente = Trim("" & r("nome_iniciais_em_maiusculas"))
		        s_nome_cliente = Left(s_nome_cliente, 15)
		
		        x = x & "		<td class='MTD tdPedido' align='left'><span class='Cn'><a style='color:" & s_cor & ";' href='javascript:fPEDConsulta(" & _
				        chr(34) & Trim("" & r("pedido")) & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o pedido'>" & _
				        Trim("" & r("pedido")) & "<br>" & s_nome_cliente & "</a></span></td>" & chr(13)

	            '> DATA
		        s = formata_data(r("data"))
		        x = x & "		<td align='center' class='MTD tdData'><span class='Cnc' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)

	            '> VALOR DO PEDIDO (PREÇO DE VENDA)
		        x = x & "		<td align='right' class='MTD tdVlPedido'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_preco_venda) & "</span></td>" & chr(13)
                s_lista_vl_pedido = s_lista_vl_pedido & vl_preco_venda & ";"

	            '> COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
		        x = x & "		<td align='right' class='MTD tdVlRT'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RT) & "</span></td>" & chr(13)
                s_lista_comissao = s_lista_comissao & vl_RT & ";"

	            '> RA BRUTO
		        x = x & "		<td align='right' class='MTD tdVlRABruto'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA) & "</span></td>" & chr(13)
                s_lista_RA_bruto = s_lista_RA_bruto & formata_moeda(vl_RA) & ";"

	            '> RA LÍQUIDO
		        x = x & "		<td align='right' class='MTD tdVlRALiq'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_liquido) & "</span></td>" & chr(13)
                s_lista_RA_liquido = s_lista_RA_liquido & formata_moeda(vl_RA_liquido) & ";"

	            '> RA DIFERENÇA
		        x = x & "		<td align='right' class='MTD tdVlRADif'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_diferenca) & "</span></td>" & chr(13)

	            '> STATUS DE PAGAMENTO
		        x = x & "		<td class='MTD tdStPagto' align='left'><span class='Cn' style='color:" & s_cor & ";'>" & x_status_pagto(Trim("" & r("st_pagto"))) & "</span></td>" & chr(13)

	            '> +/-
		        x = x & "		<td align='center' class='MTD tdSinal'><span class='C' style='font-family:Courier,Arial;color:" & s_cor_sinal & "'>" & s_sinal & "</span></td>" & chr(13)
		
	            '> COLUNA DA FIGURA (EXPANDE/RECOLHE)
		        x = x & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13)
		
		
		
		        x = x & "	</tr>" & chr(13)
		            
		        ind_anterior = r("indicador")
			    
		        r.MoveNext
		        loop
		    
            ' MOSTRA TOTAL DO ÚLTIMO INDICADOR
	        if n_reg <> 0 then 
		        s_cor="black"
		        if vl_sub_total_preco_venda < 0 then s_cor="red"
		        if vl_sub_total_RT < 0 then s_cor="red"
		        if vl_sub_total_RA < 0 then s_cor="red"
		        if vl_sub_total_RA_liquido < 0 then s_cor="red"
                sub_total_comissao = vl_sub_total_RT + vl_sub_total_RA_liquido
                msg_desconto = ""
            '-------Calculo Cedula

                z=0

                if sub_total_com_desc <= 0 Or sub_total_com_desc > 300 Or (banco = "237" Or banco = "077") then
                    if (sub_total_com_desc =0 Or sub_total_com_desc > 300) And (banco <> "237" Or banco <> "077") then 
                            totalComChqOutros = totalComChqOutros + Cstr(sub_total_com_desc)
                    elseif sub_total_com_desc <0 then 
                
                    else
                        totalComchqBradesco = totalComchqBradesco + Cstr(sub_total_com_desc)
                    end if
                else
                    dadosCalculo = Clng(sub_total_com_desc)              
                    totalComDin = totalComDin + dadosCalculo
                    dadosCalculo = o.CalculaCedulas(dadosCalculo,"2#"& limitador(5) &"|5#"&limitador(4)&"|10#"&limitador(3)&"|20#"&limitador(2)&"|50#"&limitador(1)&"|100#"&limitador(0)&"",resultadoCalculo)             
                    dadosCalculo = resultadocalculo
                    v_cedulas = Split(dadosCalculo,"|")
                    for cont=0 to Ubound(v_cedulas)
                        if cont mod 2 = 0 then
                            redim  preserve cedulas(j)
                            'cedulas(j)=  cint(v_cedulas(cont))
                            'j = j +1
                        else
                            redim preserve qtdeCedula(z)
                            qtdeCedula(z) = cint(v_cedulas(cont)) 
                            z = z + 1                   
                        end if
                    next 
                end if
                valor_desconto = 0
                qtde_registro_desc = 0  
                contador = 0
                msg_desconto = ""
                if rs2.State <> 0 then rs2.Close
                rs2.Open "SELECT COUNT(*) qtde_Desconto, descricao,valor,ordenacao  FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (apelido = '" & ind_anterior & "') GROUP BY  descricao,valor,ordenacao ORDER BY ordenacao", cn
                if Not rs2.Eof then

                            s_checked=""
                            s_disabled = " disabled"                          
                            do while Not rs2.EoF         
                                redim preserve v_desconto_descricao(contador) 
                                redim preserve  v_desconto_valor(contador)                 
                                v_desconto_descricao(contador) = rs2("descricao")
                                v_desconto_valor(contador) = rs2("valor")
                                qtde_registro_desc = qtde_registro_desc + 1
                                valor_desconto = valor_desconto + v_desconto_valor(contador)
                                contador = contador + 1
                                rs2.MoveNext
		                    loop
                            msg_desconto =  "Planilha de Desconto:&nbsp; " & Cstr(qtde_registro_desc) & " registro(s) no valor total de&nbsp;" & formata_moeda(valor_desconto)
                            cod_motivo_desconto = cod_motivo_desconto & ind_anterior & ", "
                            total_desconto_planilha = total_desconto_planilha & converte_numero(valor_desconto) & ";"
                            qtde_reg_descontos = qtde_reg_descontos & qtde_registro_desc & ";"

                end if 

                s_lista_total_comissao = s_lista_total_comissao & vl_sub_total_RT & ";"
                s_lista_total_RA = s_lista_total_RA & vl_sub_total_RA_liquido & ";"
		        x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
				        "		<td colspan='4' class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
										        "TOTAL:</span></td>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _
                                "       <input type='hidden'  id='sub_total_preco_venda_" & idx_bloco  & "' value='" & vl_sub_total_preco_venda & "'>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _
                                "       <input type='hidden'  id='sub_total_comissao_" & idx_bloco  & "' value='" & vl_sub_total_RT & "'>" & chr(13) 
                               if vl_sub_total_RA_arredondado >=0 and vl_sub_total_RT_arredondado >=0 then
                                    x = x & "       <input type='hidden' id='sub_total_comissaoRT_" & idx_bloco  & "' value='" & Clng(vl_sub_total_RT_arredondado_desc) & "'>" & chr(13)
                                else if vl_sub_total_RA_arredondado >=0 and  vl_sub_total_RT_arredondado < 0 then
                                    x = x & "       <input type='hidden' id='sub_total_comissaoRT_" & idx_bloco  & "' value='" & Clng(0) & "'>" & chr(13)
                                else
                                    x = x & "       <input type='hidden' id='sub_total_comissaoRT_" & idx_bloco  & "' value='" & Clng(sub_total_com_desc) & "'>" & chr(13)
                                end if
                                end if                               
                                x = x &_
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</span></td>" & chr(13) & _
                                "       <input type='hidden' id='sub_total_RA_" & idx_bloco  & "' value='" & vl_sub_total_RA & "'>" & chr(13) 
                                if vl_sub_total_RA_arredondado >=0 and vl_sub_total_RT_arredondado >=0 then
                                    x = x & "       <input type='hidden' id='sub_total_comissaoRA_" & idx_bloco   & "' value='" & Clng(vl_sub_total_RA_arredondado_desc) & "'>" & chr(13)
                                else if vl_sub_total_RA_arredondado < 0 and vl_sub_total_RT_arredondado > 0 then
                                    x = x & "       <input type='hidden' id='sub_total_comissaoRA_" & idx_bloco   & "' value='" & Clng(0) & "'>" & chr(13)
                                else
                                    x = x & "       <input type='hidden' id='sub_total_comissaoRA_" & idx_bloco   & "' value='" & Clng(sub_total_com_desc) & "'>" & chr(13)
                                end if
                                end if
                                x = x &_         
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _
                                "       <input type='hidden'  id='sub_total_RA_liquido_" & idx_bloco  & "' value='" & vl_sub_total_RA_liquido & "'>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _
                                "       <input type='hidden'  id='sub_total_RA_diferenca_" & idx_bloco  & "' value='" & vl_sub_total_RA_diferenca & "'>" & chr(13) & _
						        "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
						        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
						        "	</tr>" & chr(13) & _
				        "	</tr>" & chr(13) &_
	                    "   <tr>" & chr(13) 
                        s_lista_total_comissao_arredondado = s_lista_total_comissao_arredondado & vl_sub_total_RT_arredondado & ";"
                        s_lista_total_RA_arredondado = s_lista_total_RA_arredondado & vl_sub_total_RA_arredondado & ";"
                        s_lista_total_RT_arredondado_desc = s_lista_total_RT_arredondado_desc & vl_sub_total_RT_arredondado_desc & ";"
                        s_lista_total_RA_arredondado_desc = s_lista_total_RA_arredondado_desc & vl_sub_total_RA_arredondado_desc & ";"
                            if sub_total_com_desc = 0  or sub_total_com_desc > 300 Or (banco = "237" Or banco = "077")  then
                                        x = x & "<td align='left' colspan='5' nowrap><span class='Cd' style='color:" & s_cor & ";'></span>" 
                                    if msg_desconto <> "" then x = x & "<span class='Cd'><a href='javascript:abreDesconto(" & idx_bloco  & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"
                                        x = x & " <td align='right' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RT_arredondado)&"</span>"  &_
                                                " <td align='right' colspan='2' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RA_arredondado)&"</span>" 
                                           
                                    else if sub_total_com_desc >= 0 then
                                            x = x & "       <td align='left' colspan='5' nowrap>"
                                            if msg_desconto <> "" then 
                                                x = x & "<span class='Cd' ><a href='javascript:abreDesconto(" & idx_bloco  & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"
                                            else   
                                                x = x & "<span class='Rd' style='color: black;'>Cédulas: "                  
                                            for cont = 0 to UBound(qtdeCedula)
                                                    if (cont = 0 And qtdeCedula(cont) <> 0) then                 
                                                        if (qtdeCedula(cont) > 1) then
                                                            x = x & qtdeCedula(cont) & "&times;"
                                                            cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                                        end if
                                                        cedulas_descricao = cedulas_descricao & formata_moeda("100")
                                                        x = x & formata_moeda("100") & _
                                                        "       <input type='hidden'  id='cedulas100_" & idx_bloco  & "' value='100'>" & chr(13) & _
                                                        "       <input type='hidden'  id='total_cedulas100_" & idx_bloco  & "' value='" & qtdeCedula(cont) & "'>" & chr(13)
                                                        if (qtdeCedula(1) <> 0 Or qtdeCedula(2) <> 0 Or qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then
                                                            x = x & " + "
                                                            cedulas_descricao = cedulas_descricao & " + "
                                                        end if
                          
                                                elseif (cont = 1 And qtdeCedula(cont) <> 0) then
                                                    if (qtdeCedula(cont) > 1) then
                                                        x = x & qtdeCedula(cont) & "&times;"
                                                        cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                                    end if
                                                    cedulas_descricao = cedulas_descricao & formata_moeda("50")
                                                    x = x & formata_moeda("50") & _
                                                    "       <input type='hidden'  id='cedulas50_" & idx_bloco  & "' value='50'>" & chr(13) & _
                                                    "       <input type='hidden'  id='total_cedulas50_" & idx_bloco  & "' value='" & qtdeCedula(cont) & "'>" & chr(13) 
                                                    if (qtdeCedula(2) <> 0 Or qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then
                                                        x = x & " + "
                                                        cedulas_descricao = cedulas_descricao & " + "
                                                    end if
                           
                                                elseif (cont = 2 And qtdeCedula(cont) <> 0) then
                                                    if (qtdeCedula(cont) > 1) then
                                                        x = x & qtdeCedula(cont) & "&times;"
                                                        cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                                    end if
                                                    cedulas_descricao = cedulas_descricao & formata_moeda("20")
                                                    x = x & formata_moeda("20") & _
                                                    "       <input type='hidden'  id='cedulas20_" & idx_bloco  & "' value='20'>" & chr(13) & _
                                                    "       <input type='hidden'  id='total_cedulas20_" & idx_bloco  & "' value='" & qtdeCedula(2) & "'>" & chr(13) 
                                                    if (qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then
                                                        x = x & " + "
                                                        cedulas_descricao = cedulas_descricao & " + "
                                                    end if
                        
                                                elseif (cont = 3 And qtdeCedula(cont) <> 0) then
                                                    if (qtdeCedula(cont) > 1) then
                                                        x = x & qtdeCedula(cont) & "&times;"
                                                        cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                                    end if
                                                    cedulas_descricao = cedulas_descricao & formata_moeda("10")
                                                    x = x & formata_moeda("10") & _
                                                    "       <input type='hidden'  id='cedulas10_" & idx_bloco  & "' value='10'>" & chr(13) & _
                                                    "       <input type='hidden'  id='total_cedulas10_" & idx_bloco  & "' value='" & qtdeCedula(3) & "'>" & chr(13) 
                                                    if (qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then
                                                        x = x & " + "
                                                        cedulas_descricao = cedulas_descricao & " + "
                                                    end if

                                                elseif (cont = 4 And qtdeCedula(cont) <> 0) then
                                                    if (qtdeCedula(cont) > 1) then
                                                        x = x & qtdeCedula(cont) & "&times;"
                                                        cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                                    end if
                                                    cedulas_descricao = cedulas_descricao & formata_moeda("5")
                                                    x = x & formata_moeda("5") & _
                                                    "       <input type='hidden'  id='cedulas5_" & idx_bloco  & "' value='5'>" & chr(13) & _
                                                    "       <input type='hidden'  id='total_cedulas5_" & idx_bloco  & "' value='" & qtdeCedula(cont) & "'>" & chr(13) 
                                                    if (qtdeCedula(5) <> 0) then 
                                                        x = x & " + "
                                                        cedulas_descricao = cedulas_descricao & " + "
                                                    end if
                       
                                                elseif (cont = 5 And qtdeCedula(cont) <> 0) then
                                                    if (qtdeCedula(cont) > 1) then
                                                        x = x & qtdeCedula(cont) & "&times;"
                                                        cedulas_descricao = cedulas_descricao & qtdeCedula(cont) & "x"
                                                    end if
                                                    cedulas_descricao = cedulas_descricao & formata_moeda("2")
                                                    x = x & formata_moeda("2") & _
                                                    "       <input type='hidden'  id='cedulas2_" & idx_bloco  & "' value='2'>" & chr(13) & _
                                                    "       <input type='hidden'  id='total_cedulas2_" & idx_bloco  & "' value='" & qtdeCedula(cont) & "'>" & chr(13) 
                                                end if
                                                total_cedulas = total_cedulas & qtdeCedula(cont)
                                                if cont < Ubound(qtdeCedula) then total_cedulas = total_cedulas & ","                   
                    
                                            next

                                        end if
                                x = x & " <td align='right' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RT_arredondado)&"</span>"  &_
                                            " <td align='right' colspan='2' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RA_arredondado)&"</span>" 
                                end if                
                            end if
                  
                                x = x & "       <td align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>"


                            if sub_total_com_desc <= 0 then
                                s_lista_meio_pagto = s_lista_meio_pagto & " " & ";"
                                x = x & "&nbsp;</span></td>" & chr(13)
                                if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                vl_total_pagto = vl_total_pagto & "0"
                            elseif sub_total_com_desc > 300 And (banco <> "237" And banco <> "077") then
                                        s_lista_meio_pagto = s_lista_meio_pagto & "CHQ" & ";"
                                        x = x & "CHQ:</span></td>" & chr(13) &_
                                        "       <input type='hidden'  id='forma_pag_" & idx_bloco  & "' value='CHQ'>" & chr(13) 
                                    if vl_sub_total_RT_arredondado >0 and vl_sub_total_RA_arredondado < 0 then
                                        x = x &  "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco  & "' value='"& sub_total_com_desc & "'>" & chr(13)
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                            vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                    elseif vl_sub_total_RA_arredondado >0 and vl_sub_total_RT_arredondado < 0then
                                            x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco  & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                            vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                    else
                                        x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco  & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                            vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                        end if
                            elseif (banco = "237" Or banco = "077") then
                                        if banco = "237" Then
                                            s_lista_meio_pagto = s_lista_meio_pagto & "DEP" & ";"
                                        elseif banco = "077" Then
                                            s_lista_meio_pagto = s_lista_meio_pagto & "DEP1" & ";"
                                        end if
                                        x = x & "DEP:</span></td>" & chr(13) &_
                                        "       <input type='hidden'  id='forma_pag_" & idx_bloco  & "' value='DEP'>" & chr(13) 
                                    if vl_sub_total_RT_arredondado >0 and vl_sub_total_RA_arredondado < 0 then
                                        x = x &  "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco  & "' value='"& sub_total_com_desc & "'>" & chr(13) 
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                            vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                        elseif vl_sub_total_RA_arredondado >0 and vl_sub_total_RT_arredondado < 0then
                                            x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco  & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                            vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                    else
                                        x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco  & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                        vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                        end if
                            else 
                                        s_lista_meio_pagto = s_lista_meio_pagto & "DIN" & ";"
                                        x = x & "DIN:</span></td>" & chr(13) &_
                                        "       <input type='hidden'  id='forma_pag_" & idx_bloco  & "' value='DIN'>" & chr(13) 
                                    if vl_sub_total_RT_arredondado >0 and vl_sub_total_RA_arredondado < 0 then
                                            x = x &  "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco  & "' value='"& sub_total_com_desc & "'>" & chr(13) 
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                            vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                        elseif vl_sub_total_RA_arredondado >0 and vl_sub_total_RT_arredondado < 0then
                                            x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco  & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";" 
                                            vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                    else
                                        x = x & "       <input type='hidden'  id='forma_pag_valor_" & idx_bloco  & "' value='"& sub_total_com_desc &"'>" & chr(13) 
                                        if vl_total_pagto <> "" then vl_total_pagto = vl_total_pagto & ";"
                                            vl_total_pagto = vl_total_pagto & sub_total_com_desc
                                        end if

                            end if


                            if sub_total_com_desc >=0 then
                                x = x & "       <td align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(sub_total_com_desc) & "</span></td>" & chr(13) 
                            end if     
                            x = x & "	</tr>" & chr(13)

                            if sub_total_com >=0 then
                                x = x & "       <td align='right' colspan='11'><span class='Cd' style='color:gray;'>Sem desconto: " & formata_moeda(sub_total_com) & "</span></td>" & chr(13) 
                            end if     

                            if msg_desconto <> "" then
                            x = x &"   <tr>" & chr(13)& _
                                   "          <td  class='table_Desconto' id='table_Desconto_"& idx_bloco   &"'"" colspan='15' >" & chr(13)& _
                                   "          <table colspan='2' align='left' >"& chr(13)
                                   for contador = 0 to Ubound(v_desconto_descricao)                                 
                                       x = x & "   <tr>" & chr(13)& _
                                                "       <td width='15'>&nbsp;</td>" & chr(13)& _
                                                "       <td  align='left' width='400' ><span class='Cd'style='color: red;' >"& v_desconto_descricao(contador)& "</span></td>"& _
                                                "       <td align='left' ><span class='Cd'style='color: red;' > R$ "& formata_moeda(v_desconto_valor(contador))& "</span></td>"& _
                                                "   </tr>"                                     
                                   next          
                                   x = x & "   </table>"& chr(13)& _
                                           "   </td>"& chr(13)& _
                                           "</tr>"
                            end if
        
						        	        

	        '>	TOTAL GERAL
		        if qtde_indicadores >= 1 then
			        s_cor="black"

			        x = x & "	<tr>" & chr(13) & _
					        "		<td colspan='12' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
					        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					        "	</tr>" & chr(13) & _
					        "	<tr>" & chr(13) & _
					        "		<td colspan='12' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
					        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					        "	</tr>" & chr(13) & _
					        "	<tr nowrap style='background:honeydew'>" & chr(13) & _
					        "		<td class='MTBE' colspan='4' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					        "TOTAL GERAL:</span></td>" & chr(13) & _
					        "		<td class='MTB' align='right'><span id ='total_VlPedido' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_preco_venda) & "</span></td>" & chr(13) & _
					        "		<td class='MTB' align='right'><span id='totalComissao' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					        "		<td class='MTB' align='right'><span id='total_RA' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA) & "</span></td>" & chr(13) & _
					        "		<td class='MTB' align='right'><span id='total_RAliq' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) & _
					        "		<td class='MTB' align='right'><span id='total_RAdif' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_diferenca) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
					        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					        "	</tr>" & chr(13) &_
                            " </table>" & chr(13)
                   
                            if sub_total_com_desc >= 0 and valor_desconto = 0 then 
                                s_checked=" checked"    
                            else    
                                s_disabled = " disabled"
                                cod_motivo_negativo = cod_motivo_negativo & ind_anterior & ", "
                            end if  
   
                            x = "<table cellpadding='0' cellspacing='0'><tr><td valign='top'><br />" & chr(13) & _
                                "   <input type='checkbox' name='ckb_com_pg' class='CKB_COM' id='ckb_comissao_paga_tit_bloco_" & idx_bloco & "' onclick='trata_ckb_onclick();calculaTotalComissao();alternaCheck(" & idx_bloco & ");' value='" & atual & "' " & s_checked & s_disabled & " />" & chr(13) & _ 
                                "   <input type='checkbox'  style='display:none' name='ckb_com_pg_i' id='ckb_comissao_paga_tit_bloco_indicador_" & idx_bloco & "' value='" & ind_anterior & "' />" & chr(13) & _
                                "</td><td valign='top'>" & x & "</td></tr></table>" & chr(13)
                
                            x = x & "<br>"& chr(13)

                            x = x & " <table  cellspacing='0'  width='700px'> " & chr(13) & _
                            "   <tr  nowrap style='background:honeydew'>"& chr(13) & _
                            "       <td width='30%' class='MDTE' align='left'><span class='Cd' style='color:black;' > TOTAL COMISSÃO ARREDONDADO</td> "& chr(13) & _
                            "       <td class='MD MC'  style='background:honeydew'><span id='totalComissaoAd' class='Cd'>" & formata_moeda(totalComDin+totalComChqOutros+totalComChqBradesco) &" </td> "& chr(13) & _
                            "   </tr>"& chr(13) & _
                            "   <tr nowrap>"& chr(13) & _
                            "       <td  style='background:honeydew;' width='30%' class='MDTE' align='left'><span class='Cd' style='color:black;' valign='bottom'>Comissão em CHQ  </td>" & chr(13) & _
                            "       <td class='MD MC'  style='background:honeydew'><span id='totalCHQ' class='Cd'>" & formata_moeda(totalComChqOutros)& " </td> "& chr(13) & _
                            "   </tr>"& chr(13) & _
                            "    <tr nowrap >"& chr(13) & _
                            "       <td  style='background:honeydew' width='30%' class='MDTE' align='left'><span class='Cd' style='color:black;' >Comissão em DEP</td> "& chr(13) & _
                            "       <td class='MTD' style='background:honeydew'><span  id='totalComissaoDEP' class='Cd'>"& formata_moeda(totalComChqBradesco)&" </td> "& chr(13) & _
                            "   </tr>"& chr(13) & _
                            "   <tr nowrap >"& chr(13) & _
                            "       <td style='background:honeydew' width='30%' class='MDTE' align='left'><span  class='Cd' style='color:black;' >Comissão em DIN</td>" & chr(13) & _
                            "       <td class='MD MC' style='background:honeydew'><span id='totalComissaoDIN' class='Cd'>"& formata_moeda(totalComDin)&"</td>"& chr(13) & _
                            "   </tr>"& chr(13) & _
                            "   <tr nowrap >"& chr(13) & _
                            "       <td  style='background:honeydew' width='30%' class='MTBE MD' align='left'><span class='Cd' style='color:black;' >Qtde Cedulas para Comissão em DIN</td>"& chr(13) & _
                            "       <td class='MTB MD' align='left'  style='background:honeydew'><span id='totalCedulasDIN' class='Cd'>" & aux(0) & "&times;100,00 "&" + "& aux(1) & "&times;50,00 "&" + "& aux(2) & "&times;20,00 "&" + "& aux(3) & "&times;10,00 "&" + "& aux(4) & "&times;5,00 "&" + "& aux(5) & "&times; 2,00"&" </td>"& chr(13) & _
                            "   </tr>"& chr(13) & _
                            "</table>"& chr(13) & _
                    
                    
                            "<script type='text/javascript'>" & chr(13) & _
                            "function retornaTotalBloco() {" & chr(13) & _
                            "var totalBloco;"& chr(13) & _
                            "totalBloco = " & idx_bloco & ";" & chr(13) & _
                            "return totalBloco; }" & chr(13) & _
                            "</script>" & chr(13)

			        end if
		        end if

            ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	        if n_reg_total = 0 then
		        x = cab_table & cab
		        x = x & "	<tr nowrap>" & chr(13) & _
				        "		<td class='MT ALERTA' colspan='12' align='center'><span class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS NO PERÍODO ESPECIFICADO&nbsp;</span></td>" & chr(13) & _
				        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				        "	</tr>" & chr(13)
		        end if

            ' FECHA TABELA DO ÚLTIMO INDICADOR
	        x = x & "</table>" & chr(13)

	        Response.write x

	        x = "<input type='hidden' name='c_lst_vn' id='c_lst_vn' value='" & s_lista_completa_venda_normal & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_ped' id='c_lst_ped' value='" & s_lista_completa_pedidos & "' />" & chr(13) & _
		        "<input type='hidden' name='c_lst_d' id ='c_lst_d' value='" & s_lista_completa_devolucao & "' />" & chr(13) & _
		        "<input type='hidden' name='c_lst_perd' id='c_lst_perd' value='" & s_lista_completa_perda & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_vl' id='c_lst_vl' value='" & s_lista_vl_pedido & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_ttl_com' id='c_lst_ttl_com' value='" & s_lista_total_comissao & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_com_a' id='c_lst_ttl_com_a' value='" & s_lista_total_comissao_arredondado & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_com_a_desc' id='c_lst_ttl_com_a_desc' value='" & s_lista_total_RT_arredondado_desc & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_ra' id='c_lst_ttl_ra' value='" & s_lista_total_RA & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_ra_a' id='c_lst_ttl_ra_a' value='" & s_lista_total_RA_arredondado & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_ra_a_desc' id='c_lst_ttl_ra_a_desc' value='" & s_lista_total_RA_arredondado_desc & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_mp' id='c_lst_mp' value='" & s_lista_meio_pagto & "' />" & chr(13) & _
                "<input type='hidden' name='cod_m_d' id='cod_m_d' value='" & cod_motivo_desconto & "' />" & chr(13) & _
                "<input type='hidden' name='cod_m_n' id='cod_m_n' value='" & cod_motivo_negativo & "' />" & chr(13) & _
                "<input type='hidden' name='ttl_d_p' id='ttl_d_p' value='" & total_desconto_planilha & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_com' id='c_lst_com' value='" & s_lista_comissao & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_ra_b' id='c_lst_ra_b' value='" & s_lista_RA_bruto & "' />" & chr(13) & _
                "<input type='hidden' name='c_lst_ra_l' id='c_lst_ra_l' value='" & s_lista_RA_liquido & "' />" & chr(13) & _
                "<input type='hidden' name='c_qtde_r_d' id='c_qtde_r_d' value='" & qtde_reg_descontos & "' />" & chr(13) & _
                "<input type='hidden' name='c_cd' id='c_ttl_cd' value='" & total_cedulas & "' />" & chr(13) & _
                "<input type='hidden' name='c_cd_d' id='c_cd_d' value='" & cedulas_descricao & "' />" & chr(13) & _
                "<input type='hidden' name='c_op' id='c_op' value='" & operacao & "' />" & chr(13) & _
                "<input type='hidden' name='c_ttl_pagto' id='c_ttl_pagto' value='" & vl_total_pagto & "' />" & chr(13)

	        Response.write x
	
            if r.State <> 0 then r.Close
            if rs2.State <> 0 then rs2.Close
            set r=nothing

    end if
	
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
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var windowScrollTopAnterior;
window.status = 'Aguarde, executando a consulta ...';

$(function() {
	$("#divPedidoConsulta").hide();

	sizeDivPedidoConsulta();

	$('#divInternoPedidoConsulta').addClass('divFixo');

	$(document).keyup(function(e) {
		if (e.keyCode == 27) fechaDivPedidoConsulta();
	});

	$("#divPedidoConsulta").click(function() {
		fechaDivPedidoConsulta();
	});

	$("#imgFechaDivPedidoConsulta").click(function() {
		fechaDivPedidoConsulta();
	});

	// EXIBE O REALCE NOS CHECKBOXES QUE SÃO EXIBIDOS INICIALMENTE ASSINALADOS
	$(".CKB_COM:enabled").each(function() {
		if ($(this).is(":checked")) {
			$(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
		}
		else {
			$(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
		}
	})

	// EVENTO P/ REALÇAR OU NÃO CONFORME SE MARCA/DESMARCA O CHECKBOX
	$(".CKB_COM:enabled").click(function() {
		if ($(this).is(":checked")) {
			$(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
		}
		else {
			$(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
		}
	})

	var i, ttl_bloco;
	ttl_bloco = retornaTotalBloco();
	for (i = 1; i <= ttl_bloco; i++) {
	    if ($("#ckb_comissao_paga_tit_bloco_" + i).is(':checked')) {
	        $("#ckb_comissao_paga_tit_bloco_indicador_" + i).prop('checked', true);
	    }
	    else {
	        $("#ckb_comissao_paga_tit_bloco_indicador_" + i).attr('checked', false);
	    }
	}

	$(".table_Desconto").hide();
});

function abreDesconto(idx_bloco) {
    var s_seletor = "#table_Desconto_" + idx_bloco;

    $(s_seletor).toggle();
}
//Every resize of window
$(window).resize(function() {
	sizeDivPedidoConsulta();
});

function sizeDivPedidoConsulta() {
	var newHeight = $(document).height() + "px";
	$("#divPedidoConsulta").css("height", newHeight);
}

function fechaDivPedidoConsulta() {
	$(window).scrollTop(windowScrollTopAnterior);
	$("#divPedidoConsulta").fadeOut();
	$("#iframePedidoConsulta").attr("src", "");
}

function marcar_todos() {
	$(".CKB_COM:enabled")
		.prop("checked", true)
		.parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
}

function desmarcar_todos() {
	$(".CKB_COM:enabled")
		.prop("checked", false)
		.parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
}

function trata_ckb_onclick(idx_bloco) {
var s_id, s_class;
	s_id = "#ckb_comissao_paga_tit_bloco_" + idx_bloco;
	s_class = ".CKB_COM_BL_" + idx_bloco;
	if ($(s_id).is(":checked")) {
		$(s_class).prop("checked", true);
		$(s_class).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
	}
	else {
		$(s_class).prop("checked", false);
		$(s_class).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
	}
}

function fExibeOcultaCampos(indice_bloco) {
	var s_seletor = "#tableDados_" + indice_bloco + " .VISAO_ANALIT";
	$(s_seletor).toggle();
}

function expandir_todos() {
	$(".VISAO_ANALIT").show();
}

function recolher_todos() {
	$(".VISAO_ANALIT").hide();
}

function fPEDConsulta(id_pedido, usuario) {
	windowScrollTopAnterior = $(window).scrollTop();
	sizeDivPedidoConsulta();
	$("#iframePedidoConsulta").attr("src", "PedidoConsultaView.asp?pedido_selecionado=" + id_pedido + "&pedido_selecionado_inicial=" + id_pedido + "&usuario=" + usuario);
	$("#divPedidoConsulta").fadeIn();
}

function fORCConsulta(id_orcamento, usuario) {
	windowScrollTopAnterior = $(window).scrollTop();
	sizeDivPedidoConsulta();
	$("#iframePedidoConsulta").attr("src", "OrcamentoConsultaView.asp?orcamento_selecionado=" + id_orcamento + "&orcamento_selecionado_inicial=" + id_orcamento + "&usuario=" + usuario);
	$("#divPedidoConsulta").fadeIn();
}

function fRELGravaDados(f) {
	window.status = "Aguarde ...";
	bCONFIRMA.style.visibility = "hidden";
	f.action = "RelComissaoIndicadoresPagDescExecConfirma.asp";

	f.submit();
}
</script>
<script type="text/javascript">
    function calculaTotalComissao() {
        var total, totalCHQ, ttl_bloco, i, n, n1, n2, n3, totalComissaoAd, totalDEP, totalDIN, totalVLPedido, totalRA, totalRALiq, totalRADif, totalRTA, totalRAA;
        var total100, total50, total20, total10, total5, total2, cedulas100, cedulas50, cedulas20, cedulas10, cedulas5, cedulas2, textDin, textComissao;
        var VlPedido, RA, RAliq, RADif, RTA, RAA, TotalqtdeChq;
    ttl_bloco = retornaTotalBloco();
    TotalqtdeChq = 0;
    total100 = 0; total50 = 0; total20 = 0; total10 = 0; total5 = 0; total2 = 0; cedulas100 = 0; cedulas50 = 0; cedulas20 = 0; cedulas10 = 0; cedulas5 = 0; cedulas2 = 0;
    totalDIN = 0; totalDEP = 0; totalComissaoAd = 0; total = 0; totalCHQ = 0; n1 = 0; n3 = 0; textDin = ""; textComissao = "";
    totalVLPedido = 0; totalRA = 0; totalRALiq = 0; totalRADif = 0; totalRTA = 0; totalRAA = 0; RTA = 0; RAA = 0;
    VlPedido = 0;
    for (i=1;i<=ttl_bloco;i++) {        
        if ($("#ckb_comissao_paga_tit_bloco_"+i).is(':checked')) {
            n = converte_numero($("#sub_total_comissao_"+i).val());
            n2 = converte_numero($("#sub_total_comissaoAd_" + i).val());
            VlPedido = converte_numero($("#sub_total_preco_venda_" + i).val());
            RA = converte_numero($("#sub_total_RA_" + i).val());
            RAliq = converte_numero($("#sub_total_RA_liquido_" + i).val());
            RADif = converte_numero($("#sub_total_RA_diferenca_" + i).val());
            RADif = converte_numero($("#sub_total_RA_diferenca_" + i).val());
            RTA = converte_numero($("#sub_total_comissaoRT_" + i).val());
            RAA = converte_numero($("#sub_total_comissaoRA_" + i).val());
        }
        else {
        n = 0;
        n2 = 0;
        VlPedido = 0;
        RA = 0;
        RAliq = 0;
        RADif = 0;
        RTA = 0;
        RAA = 0;
        }

        
        //CHQ
       if(($("#forma_pag_"+i).val() == "CHQ") && $("#ckb_comissao_paga_tit_bloco_"+i).is(':checked') ){
           n1 = converte_numero($("#forma_pag_valor_" + i).val());
           TotalqtdeChq = TotalqtdeChq + 1;
            }
        else{
        n1 = 0;
       }
        //DEP
       if((($("#forma_pag_"+i).val() == "DEP") || (($("#forma_pag_"+i).val() == "DEP1"))) && $("#ckb_comissao_paga_tit_bloco_"+i).is(':checked') ){
                n3 = converte_numero($("#forma_pag_valor_"+i).val());
            }
        else{
        n3 = 0;
        }
        //DIN
         if(($("#forma_pag_"+i).val() == "DIN") && $("#ckb_comissao_paga_tit_bloco_"+i).is(':checked') ){
                n4 = converte_numero($("#forma_pag_valor_"+i).val());
            }
        else{
        n4 = 0;
        }
        //100
         if (($("#cedulas100_" + i).val() == "100") && $("#ckb_comissao_paga_tit_bloco_" + i).is(':checked')) {
             cedulas100 = converte_numero($("#total_cedulas100_" +i).val());
         }
         else {
             cedulas100 = 0;
         }
        //50
         if (($("#cedulas50_" + i).val() == "50") && $("#ckb_comissao_paga_tit_bloco_" + i).is(':checked')) {
             cedulas50 = converte_numero($("#total_cedulas50_" + i).val());
         }
         else {
             cedulas50 = 0;
         }
        //20
         if (($("#cedulas20_" + i).val() == "20") && $("#ckb_comissao_paga_tit_bloco_" + i).is(':checked')) {
             cedulas20 = converte_numero($("#total_cedulas20_" + i).val());
         }
         else {
             cedulas20 = 0;
         }
        //10
         if (($("#cedulas10_" + i).val() == "10") && $("#ckb_comissao_paga_tit_bloco_" + i).is(':checked')) {
             cedulas10 = converte_numero($("#total_cedulas10_" + i).val());
         }
         else {
             cedulas10 = 0;
         }
        //5
         if (($("#cedulas5_" + i).val() == "5") && $("#ckb_comissao_paga_tit_bloco_" + i).is(':checked')) {
             cedulas5 = converte_numero($("#total_cedulas5_" + i).val());
         }
         else {
             cedulas5 = 0;
         }
        //2
         if (($("#cedulas2_" + i).val() == "2") && $("#ckb_comissao_paga_tit_bloco_" + i).is(':checked')) {
             cedulas2 = converte_numero($("#total_cedulas2_" + i).val());
         }
         else {
             cedulas2 = 0;
         }

        // Total de cedulas 
         total100 += cedulas100; total50 += cedulas50; total20 += cedulas20; total10 += cedulas10; total5 += cedulas5; total2 += cedulas2;
        // Total de CHQ,DIN,DEP,COMISSAO ARREDONDADO E SEM ARREDONDAR.
        total += n; totalCHQ += n1; totalComissaoAd += n2; totalDEP += n3; totalDIN += n4;
        totalVLPedido += VlPedido;
        totalRA += RA; totalRALiq += RAliq; totalRADif += RADif;
        totalRAA += RAA; totalRTA += RTA;
    }

   // totalRTA = totalRTA - (totalRTA * (COMISSAO_INDICADOR_PERC_DESCONTO_SEM_NF / 100));
  //  totalRAA = totalRAA - (totalRAA * (COMISSAO_INDICADOR_PERC_DESCONTO_SEM_NF / 100));


    $("#totalComissao").text(formata_moeda(total));
    
    if (TotalqtdeChq > 0) {
        $("#totalCHQ").html(formata_moeda(totalCHQ) + "&nbsp;(Nº cheques&nbsp;" + TotalqtdeChq + ")");
    }
    else{
         $("#totalCHQ").html(formata_moeda(totalCHQ) + "&nbsp;");
    }
    
    $("#totalComissaoDEP").text(formata_moeda(totalDEP));
    $("#totalComissaoDIN").text(formata_moeda(totalDIN));
    $("#total_VlPedido").text(formata_moeda(totalVLPedido));
    $("#total_RA").text(formata_moeda(totalRA));
    $("#total_RAliq").text(formata_moeda(totalRALiq));
    $("#total_RAdif").text(formata_moeda(totalRADif));
   
    if (totalRTA != 0) {
        textComissao = textComissao + "COM: " + formata_moeda(String(totalRTA))
    if (totalRAA != 0) { textComissao = textComissao + "&nbsp;+&nbsp;" }
    }
    if (totalRAA != 0) {
        textComissao = textComissao + "RA: " + formata_moeda(String(totalRAA))
        if (totalRTA != 0) { textComissao = textComissao + "&nbsp;=&nbsp;" + formata_moeda((totalRTA + totalRAA)) }
    }
   
    $("#totalComissaoAd").html(textComissao);
    if(total100 != 0)
    {
        textDin = textDin + String(total100) + "&times;100,00";
        if ((total50 != 0) || (total20 != 0) || (total10 != 0) || (total5 != 0) || (total2 != 0)) { textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"; }
    }
    if (total50 != 0) {
        textDin = textDin + String(total50) + "&times;50,00";
        if ((total20 != 0) || (total10 != 0) || (total5 != 0) || (total2 != 0)) { textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"; }
    }
    if (total20 != 0) {
        textDin = textDin + String(total20) + "&times;20,00";
        if ((total10 != 0) || (total5 != 0) || (total2 != 0)) { textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"; }
    }
    if (total10 != 0) {
        textDin = textDin + String(total10) + "&times;10,00";
        if ((total5 != 0) || (total2 != 0)) { textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"; }
    }
    if (total5 != 0) {
        textDin = textDin + String(total5) + "&times;5,00";
        if ( (total2 != 0)) { textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"; }
    }
    if (total2 != 0) {
        textDin = textDin + String(total2) + "&times;2,00";
    }
    
        $("#totalCedulasDIN").html(textDin);

    }
    function alternaCheck(i) {
            if ($("#ckb_comissao_paga_tit_bloco_" + i).is(':checked')) {
                $("#ckb_comissao_paga_tit_bloco_indicador_" + i).prop("checked", true);
            }
            else {
                $("#ckb_comissao_paga_tit_bloco_indicador_" + i).prop('checked', false);
            }
            
    }

    function alternaCheckTodos() {
    var i, ttl_bloco;
	ttl_bloco = retornaTotalBloco();
	for (i = 1; i <= ttl_bloco; i++) {
	    if ($("#ckb_comissao_paga_tit_bloco_" + i).is(':checked')) {
	        $("#ckb_comissao_paga_tit_bloco_indicador_" + i).prop('checked', true);
	    }
	    else {
	        $("#ckb_comissao_paga_tit_bloco_indicador_" + i).attr('checked', false);
	    }
	}
}
</script>
<script type="text/javascript">
 $(function() {
    calculaTotalComissao();
});
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
.tdCkb
{
	width: 20px;
}
.tdLoja{
	width: 28px;
	}
.tdOrcamento{
	width: 58px;
	}
.tdPedido{
	width: 83px;
	}
.tdData{
	width: 60px;
	}
.tdVlPedido{
	width: 70px;
	}
.tdVlRT{
	width: 60px;
	}
.tdVlRABruto{
	width: 60px;
	}
.tdVlRALiq{
	width: 60px;
	}
.tdVlRADif{
	width: 60px;
	}
.tdStPagto{
	width: 60px;
	}
.tdSinal{
	width: 18px;
	}
.BTN_LNK
{
	min-width:140px;
}
.CKB_HIGHLIGHT
{
	background-color:#90EE90;
}
#divPedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsulta
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
#divInternoPedidoConsulta.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivPedidoConsulta
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframePedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
.BkgWhite
{
	background-color:#ffffff;
}
.VISAO_ANALIT
{
	<%if blnVisaoSintetica then Response.Write "display:none;"%>
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

<% else
     %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<input type="hidden" name="mes" id="mes" value="<%=c_dt_entregue_mes%>">
<input type="hidden" name="ano" id="ano" value="<%=c_dt_entregue_ano%>">
<input type="hidden" name="c_usuario_sessao" id="c_usuario_sessao" value="<%=usuario%>" />
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores Com Desconto (Processamento)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<%
	s_filtro = "<table width='709' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	PERÍODO: MÊS DE COMPETÊNCIA
	s = ""
	if (c_dt_entregue_mes <> "") Or (c_dt_entregue_ano <> "") then
	'	DEVIDO AO WORD WRAP: SÓ FAZ WORD WRAP QUANDO ENCONTRA CHR(32), OU SEJA, MANTÉM AGRUPADO TEXTO COM &nbsp;
		if s <> "" then s = s & ",&nbsp; "
		s_aux = c_dt_entregue_mes
		if s_aux = "" then 
            s_aux = "N.I."
        else
            if c_dt_entregue_ano = "" then 
                s_aux = "N.I."
            else
	    	    s_aux = " " & s_aux & "/"
		        s_aux = replace(s_aux, " ", "&nbsp;")
		        s = s & s_aux
		        s_aux = c_dt_entregue_ano
		        s_aux = replace(s_aux, " ", "&nbsp;")
            end if
        end if
		s = s & s_aux  
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Mês de competência:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	COMISSÃO PAGA
	s = ""
	s_aux = ckb_comissao_paga_sim
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & "paga"
		end if
	
	s_aux = ckb_comissao_paga_nao
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = s & "não-paga"
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Comissão:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	STATUS DE PAGAMENTO
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
					"		<td align='right' valign='top' nowrap><span class='N'>Status de Pagamento:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	INDICADOR
	if c_indicador <> "" then
		s = c_indicador
		s_aux = x_orcamentista_e_indicador(c_indicador)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Indicador:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	VENDEDOR
	if c_vendedor <> "" then
		s = c_vendedor
		s_aux = x_usuario(c_vendedor)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Vendedor(es):&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	LOJA(S)
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
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Loja(s):&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	EMISSÃO
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td class="Rc" align="left">&nbsp;</td>
</tr>
</table>


    <%if aviso="" then%>
<table class="notPrint" width="709" cellpadding="0" cellspacing="0" style="margin-top:5px;">
<tr>
	<td align="right">
		<button type="button" name="bExpandirTodos" id="bExpandirTodos" class="Button BTN_LNK" onclick="expandir_todos();" title="expandir todas as linhas de dados" style="margin-left:6px;margin-bottom:2px">Expandir Tudo</button>
		&nbsp;
		<button type="button" name="bRecolherTodos" id="bRecolherTodos" class="Button BTN_LNK" onclick="recolher_todos();" title="recolher todas as linhas de dados" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Recolher Tudo</button>
		&nbsp;
		<button type="button" name="bMarcarTodos" id="bMarcarTodos" class="Button BTN_LNK" onclick="marcar_todos();calculaTotalComissao();alternaCheckTodos();" title="assinala todos os pedidos para gravar o status da comissão como paga" style="margin-left:6px;margin-bottom:2px">Marcar todos</button>
		&nbsp;
		<button type="button" name="bDesmarcarTodos" id="bDesmarcarTodos" class="Button BTN_LNK" onclick="desmarcar_todos();calculaTotalComissao();alternaCheckTodos();" title="desmarca todos os pedidos para gravar o status da comissão como não-paga" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Desmarcar todos</button>
	</td>
</tr>
</table>
    <%end if%>

<br />
<table class="notPrint" width="709" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="left">&nbsp;</td>
	<td align="right">
        <%if aviso = "" then %>
		<div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	    <%end if %>
    </td>
</tr>
</table>


</form>

</center>

<div id="divPedidoConsulta"><center><div id="divInternoPedidoConsulta"><img id="imgFechaDivPedidoConsulta" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframePedidoConsulta"></iframe></div></center></div>

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
