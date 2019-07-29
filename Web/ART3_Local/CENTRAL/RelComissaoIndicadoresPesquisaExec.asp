<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================================
'	  RelComissaoIndicadoresPesquisaExec.asp
'     ===========================================================================
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

	dim alerta
	dim s, s_aux, s_filtro
	dim c_cnpj_cpf, str_data
    
	alerta = ""

	c_cnpj_cpf = retorna_so_digitos(Trim(Request.Form("c_cnpj_cpf")))

    str_data = "01/" & Month(Date()) & "/" & Year(Date())
    str_data = strToDate(str_data)
        
' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'

sub consulta_executa
const VENDA_NORMAL = "VENDA_NORMAL"
const DEVOLUCAO = "DEVOLUCAO"
const PERDA = "PERDA"
dim r
dim s, s_aux, s_sql, x, cab_table, cab, n_reg, n_reg_total, qtde_indicadores, ind_anterior
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
dim s_banco, s_banco_nome, s_agencia, s_conta, s_favorecido, s_favorecido_cnpj_cpf
dim s_nome_cliente, s_desempenho_nota
dim s_checked, s_class, s_class_td, idx_bloco, s_new_cab
dim s_lista_completa_venda_normal, s_lista_completa_devolucao, s_lista_completa_perda, s_lista_completa_pedidos, qtde_reg_descontos
dim s_lista_comissao, s_lista_RA_bruto, s_lista_RA_liquido, s_lista_total_comissao, s_lista_total_comissao_arredondado, s_lista_meio_pagto, s_lista_vl_pedido, s_lista_total_RA, s_lista_total_RA_arredondado
dim banco, atual,qtdeChq
dim cod_motivo_desconto, cod_motivo_negativo, total_desconto_planilha
dim vl_sub_total_RT_arredondado,vl_sub_total_RA_arredondado, total_cedulas, cedulas_descricao
dim conta_vendedor, vendedor_processado,sub_total_com_RA,sub_total_com_RT, vl_total_pagto,sub_total_com,v_desconto_descricao(),v_desconto_valor(),contador,valor_desconto,qtde_registro_desc
dim vendedor_a, s_id, msg_desconto
dim s_lista_ja_marcado_venda_normal, s_lista_ja_marcado_devolucao, s_lista_ja_marcado_perda

    sub_total_com_RA = 0
    sub_total_com_RT = 0
    	if s_where_comissao_descontada <> "" then
		if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		s_where_perdas = s_where_perdas & " (" & s_where_comissao_descontada & ")"
		end if
		
    if alerta = "" then

        '	CRITÉRIO: COMISSÃO PAGA
        '	A) VENDAS
	        s_where_comissao_paga = ""
	        s = ""

		        if s <> "" then s = s & " OR"
		        s = s & " (t_PEDIDO.comissao_paga = 0)"

	        if s <> "" then 
		        if s_where_comissao_paga <> "" then s_where_comissao_paga = s_where_comissao_paga & " AND"
		        s_where_comissao_paga = s_where_comissao_paga & " (" & s & ")"
		        end if
		
        '	B) PERDAS/DEVOLUÇÕES
	        s_where_comissao_descontada = ""
	        s = ""

		        if s <> "" then s = s & " OR"
		        s = s & " (comissao_descontada = 0)"
	
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
		        if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		        s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data < " & bd_formata_data(str_data) & ")"
		
		
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
		        if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		        s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(str_data) & ")"

	        if s_where_comissao_descontada <> "" then
		        if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		        s_where_devolucao = s_where_devolucao & " (" & s_where_comissao_descontada & ")"
		        end if

        '	CRITÉRIOS PARA PERDAS
	        s_where_perdas = ""
		        if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		        s_where_perdas = s_where_perdas & " (t_PEDIDO_PERDA.data < " & bd_formata_data(str_data) & ")"
		
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
			        " t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor, t_CLIENTE.nome, t_CLIENTE.nome_iniciais_em_maiusculas," & _
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
                    " AND (t_ORCAMENTISTA_E_INDICADOR.cnpj_cpf = '" & c_cnpj_cpf & "')" & _
			        s & _
			        " GROUP BY t_PEDIDO.pedido, t_PEDIDO.comissao_paga, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor, t_CLIENTE.nome, t_CLIENTE.nome_iniciais_em_maiusculas, t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO.entregue_data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"

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
			        " t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor, t_CLIENTE.nome, t_CLIENTE.nome_iniciais_em_maiusculas," & _
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
                    " AND (t_ORCAMENTISTA_E_INDICADOR.cnpj_cpf = '" & c_cnpj_cpf & "')" & _                    
			        " GROUP BY t_PEDIDO_ITEM_DEVOLVIDO.id, t_PEDIDO_ITEM_DEVOLVIDO.comissao_descontada, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor, t_CLIENTE.nome, t_CLIENTE.nome_iniciais_em_maiusculas, t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"

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
			        " t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor, t_CLIENTE.nome, t_CLIENTE.nome_iniciais_em_maiusculas," & _
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
                    " AND (t_ORCAMENTISTA_E_INDICADOR.cnpj_cpf = '" & c_cnpj_cpf & "')" & _
			        " GROUP BY t_PEDIDO_PERDA.id, t_PEDIDO_PERDA.comissao_descontada, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor, t_CLIENTE.nome, t_CLIENTE.nome_iniciais_em_maiusculas, t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_PERDA.data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"

	        s_sql = "SELECT " & _
				        "*" & _
			        " FROM (" & _
				        s_sql & _
				        ") t" & _
			        " ORDER BY t.vendedor, t.desempenho_nota, t.indicador, t.numero_loja, t.data, t.pedido, t.total_preco_venda DESC"

            ' CABEÇALHO
	        cab_table = "<table border='0' cellspacing='0' id='tableDados'>" & chr(13)
	        cab = "	<tr style='background:azure' nowrap>" & chr(13) & _
                    "		<td class='MDTE tdCkb' align='center' valign='bottom' nowrap><input type='checkbox' name='ckb_comissao_paga_tit_bloco' id='ckb_comissao_paga_tit_bloco' class='CKB_COM VISAO_ANALIT' onclick='trata_ckb_onclick();' /></td>" & chr(13) & _
		            "		<td class='MTD tdLoja' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>Loja</span></td>" & chr(13) & _
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
		            "	</tr>" & chr(13)
	
	        x = ""
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
            qtde_reg_descontos = ""
            total_cedulas = ""
            s_lista_ja_marcado_venda_normal = ""
	        s_lista_ja_marcado_devolucao = ""
	        s_lista_ja_marcado_perda = ""
            cedulas_descricao = ""
            vl_sub_total_RT_arredondado = 0
            vl_sub_total_RT_arredondado = 0
	        vendedor_a = "XXXXXXXXXXXX"
            sub_total_com = 0
	        set r = cn.execute(s_sql)

	        do while Not r.Eof
    
	        '	MUDOU DE VENDEDOR?
	        if Trim("" & r("vendedor"))<>vendedor_a then
			        vendedor_a = Trim("" & r("vendedor"))
			        idx_bloco = idx_bloco + 1
			        qtde_indicadores = qtde_indicadores + 1      
                
		            ' FECHA TABELA DO INDICADOR ANTERIOR
			        if n_reg_total > 0 then 
                      
                    ' TOTAL DO INDICADOR
				        s_cor="black"
				        if vl_sub_total_preco_venda < 0 then s_cor="red"
				        if vl_sub_total_RT < 0 then s_cor="red"
				        if vl_sub_total_RA < 0 then s_cor="red"
				        if vl_sub_total_RA_liquido < 0 then s_cor="red"

                        rs2.Open "SELECT COUNT(*) qtde_Desconto, descricao,valor,ordenacao  FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (apelido = '" & ind_anterior & "') GROUP BY  descricao,valor,ordenacao ORDER BY ordenacao", cn
                        msg_desconto = ""
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
                                msg_desconto =  "Planilha de Desconto:&nbsp; " & Cstr(qtde_registro_desc) & " registro(s) no valor total de&nbsp;" & formata_moeda(valor_desconto)                       
                        else
                                msg_desconto= ""
                    
                            end if

				        x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
						        "		<td class='MTBE' colspan='5' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
						        "TOTAL:</span></td>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _
                                "       <input type='hidden' id='sub_total_preco_venda_" & idx_bloco - 1 & "' value='" & vl_sub_total_preco_venda & "'>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13)
                                                       
                                x = x &_    
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</span></td>" & chr(13) & _
                                "       <input type='hidden' id='sub_total_RA_" & idx_bloco - 1 & "' value='" & vl_sub_total_RA & "'>" & chr(13) 

                                x = x &_                                
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _
                                "       <input type='hidden'  id='sub_total_RA_liquido_" & idx_bloco - 1 & "' value='" & vl_sub_total_RA_liquido & "'>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _
                                "       <input type='hidden'  id='sub_total_RA_diferenca_" & idx_bloco - 1 & "' value='" & vl_sub_total_RA_diferenca & "'>" & chr(13) & _
						        "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
						        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
                                "   <tr>" & chr(13) & _
                                "       <td align='left' colspan='9' nowrap>" & chr(13)

                                if msg_desconto <> "" then 
                                    x = x & "<span class='Cd' ><a href='javascript:abreDesconto(" & idx_bloco -1  & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"
                                    end if

                                x = x & "       <td align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & chr(13) & _
                                "       <td align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(sub_total_com) & "</span></td>" & chr(13) & _
						        "	</tr>" & chr(13)

                                if msg_desconto <> "" then
                                    x = x &"   <tr>" & chr(13) & _
                                   "          <td  class='VISAO_ANALIT' id='table_Desconto_"& idx_bloco -1   &"' style='display: none;' colspan='15' >" & chr(13)& _
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
                    
             
                        atual = ""
				        Response.Write x
				        x="<BR>" & chr(13)
				        end if
    
                    s_sql = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido = '" & Trim("" & r("indicador")) & "')"
			        if rs.State <> 0 then rs.Close
                    if rs2.State <> 0 then rs2.Close 

           
			        rs.Open s_sql, cn
			        if Not rs.Eof then
				        s_banco = Trim("" & rs("banco"))
				        s_agencia = Trim("" & rs("agencia"))
				        s_conta = Trim("" & rs("conta"))
				        s_favorecido = Trim("" & rs("favorecido"))
						s_favorecido_cnpj_cpf = Trim("" & rs("favorecido_cnpj_cpf"))
						if s_favorecido_cnpj_cpf <> "" then s_favorecido_cnpj_cpf = cnpj_cpf_formata(s_favorecido_cnpj_cpf)
				        s_banco_nome = x_banco(s_banco)
				        if (s_banco <> "") And (s_banco_nome <> "") then s_banco = s_banco & " - " & s_banco_nome
			        else
				        s_banco = ""
				        s_banco_nome = ""
				        s_agencia = ""
				        s_conta = ""
				        s_favorecido = ""
						s_favorecido_cnpj_cpf = ""
				        end if

                    x = x & Replace(cab_table, "tableDados", "tableDados_" & idx_bloco)
                    x = x & "	<tr>" & chr(13)
            
                    s = Trim("" & r("indicador"))
			        s_aux = x_orcamentista_e_indicador(s)
			        if (s<>"") And (s_aux<>"") then s = s & " - "
			        s = s & s_aux
			        x = x & "		<td colspan='12' align='left' valign='bottom' class='MB' style='background:white;'><span class='N'>Vendedor:&nbsp;" & Trim("" & r("vendedor")) & "</span></td>" & chr(13) & _
                            "   </tr>" & chr(13) & _
                            "   <tr>" & chr(13) & _
                            "   <td>&nbsp;</td>" & chr(13) & _
                            "   </tr>" & chr(13) & _
                            "   <tr>" & chr(13)
			        if s <> "" then x = x & "		<td class='MDTE' colspan='12' align='left' valign='bottom' class='MB' style='background:azure;'><span class='N'>&nbsp;" & s_desempenho_nota & s & "</span></td>" & chr(13) & _
									        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
									        "	</tr>" & chr(13) & _
									        "	<tr>" & chr(13) & _
									        "		<td class='MDTE' colspan='12' align='left' valign='bottom' class='MB' style='background:whitesmoke;'>" & chr(13) & _
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
									        "				</tr>" & chr(13)

								if Len(retorna_so_digitos(s_favorecido_cnpj_cpf)) = 11 then
									s_aux = "CPF"
								else
									s_aux = "CNPJ"
									end if

								x = x & _
									"				<tr>" & chr(13) & _
									"					<td colspan='3' class='MC' align='left' valign='bottom'><span class='Cn'>" & s_aux & ": " & s_favorecido_cnpj_cpf & "</span></td>" & chr(13) & _
									"				</tr>" & chr(13)

						x = x & _
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
                

                sub_total_com = vl_sub_total_RT + vl_sub_total_RA_liquido

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

     '> CHECK BOX
	 '	É USADO O CÓDIGO DA OPERAÇÃO (VENDA NORMAL, DEVOLUÇÃO, PERDA) P/ NÃO CORRER O RISCO DE HAVER CONFLITO DEVIDO A ID'S REPETIDOS ENTRE AS OPERAÇÕES
		s_id = "ckb_comissao_paga_" & Trim("" & r("operacao")) & "_" & Trim("" & r("id_registro"))
		s_checked = ""
		s_class = " CKB_COM_BL_" & idx_bloco
		s_class_td = ""
		if Trim("" & r("operacao")) = VENDA_NORMAL then
			if s_lista_completa_venda_normal <> "" then s_lista_completa_venda_normal = s_lista_completa_venda_normal & ";"
			s_lista_completa_venda_normal = s_lista_completa_venda_normal & Trim("" & r("id_registro"))
			s_class = s_class & " CKB_COM_VDNORM"
			if CLng(r("status_comissao")) = CLng(COD_COMISSAO_PAGA) then
				s_checked = " checked"
				s_class_td = s_class_td & " CKB_HIGHLIGHT"
				if s_lista_ja_marcado_venda_normal <> "" then s_lista_ja_marcado_venda_normal = s_lista_ja_marcado_venda_normal & ";"
				s_lista_ja_marcado_venda_normal = s_lista_ja_marcado_venda_normal & Trim("" & r("id_registro"))
				end if
		elseif Trim("" & r("operacao")) = DEVOLUCAO then
			if s_lista_completa_devolucao <> "" then s_lista_completa_devolucao = s_lista_completa_devolucao & ";"
			s_lista_completa_devolucao = s_lista_completa_devolucao & Trim("" & r("id_registro"))
			s_class = s_class & " CKB_COM_DEV"
			if CLng(r("status_comissao")) = CLng(COD_COMISSAO_DESCONTADA) then
				s_checked = " checked"
				s_class_td = s_class_td & " CKB_HIGHLIGHT"
				if s_lista_ja_marcado_devolucao <> "" then s_lista_ja_marcado_devolucao = s_lista_ja_marcado_devolucao & ";"
				s_lista_ja_marcado_devolucao = s_lista_ja_marcado_devolucao & Trim("" & r("id_registro"))
				end if
		elseif Trim("" & r("operacao")) = PERDA then
			if s_lista_completa_perda <> "" then s_lista_completa_perda = s_lista_completa_perda & ";"
			s_lista_completa_perda = s_lista_completa_perda & Trim("" & r("id_registro"))
			s_class = s_class & " CKB_COM_PERDA"
			if CLng(r("status_comissao")) = CLng(COD_COMISSAO_DESCONTADA) then
				s_checked = " checked"
				s_class_td = s_class_td & " CKB_HIGHLIGHT"
				if s_lista_ja_marcado_perda <> "" then s_lista_ja_marcado_perda = s_lista_ja_marcado_perda & ";"
				s_lista_ja_marcado_perda = s_lista_ja_marcado_perda & Trim("" & r("id_registro"))
				end if
			end if
		    x = x & "		<td class='MDTE tdCkb" & s_class_td & "' align='center'><input type='checkbox' class='CKB_COM " & s_class & "' name='" & s_id & "' id='" & s_id & "' value='" & Trim("" & r("id_registro")) & "|" & Trim("" & r("operacao")) & "'" & s_checked & " /></td>" & chr(13)
		
	            '> LOJA
		        x = x & "		<td class='MTD tdLoja' align='center'><span class='Cnc' style='color:" & s_cor & ";'>" & Trim("" & r("loja")) & "</span></td>" & chr(13)

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

                        rs2.Open "SELECT COUNT(*) qtde_Desconto, descricao,valor,ordenacao  FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (apelido = '" & ind_anterior & "') GROUP BY  descricao,valor,ordenacao ORDER BY ordenacao", cn
                        msg_desconto = ""
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
                                msg_desconto =  "Planilha de Desconto:&nbsp; " & Cstr(qtde_registro_desc) & " registro(s) no valor total de&nbsp;" & formata_moeda(valor_desconto)                       
                        else
                                msg_desconto= ""
                    
                            end if

		        x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
				        "		<td colspan='5' class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
										        "TOTAL:</span></td>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _
                                "       <input type='hidden'  id='sub_total_preco_venda_" & idx_bloco  & "' value='" & vl_sub_total_preco_venda & "'>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _
                                "       <input type='hidden'  id='sub_total_comissao_" & idx_bloco  & "' value='" & vl_sub_total_RT & "'>" & chr(13) 
                        
                                x = x &_
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</span></td>" & chr(13) & _
                                "       <input type='hidden' id='sub_total_RA_" & idx_bloco  & "' value='" & vl_sub_total_RA & "'>" & chr(13) 

                                x = x &_         
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _
						        "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _
						        "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
						        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
						        "	</tr>" & chr(13) & _
				        "	</tr>" & chr(13) &_
	                    "   <tr>" & chr(13) & _
                        "       <td align='left' colspan='9' nowrap>" & chr(13)

                        if msg_desconto <> "" then 
                            x = x & "<span class='Cd' ><a href='javascript:abreDesconto(" & idx_bloco & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"
                            end if

                        x = x & "       <td align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & chr(13)

                        x = x & "       <td align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(sub_total_com) & "</span></td>" & chr(13) 

                        x = x & "	</tr>" & chr(13)     
    
                        if msg_desconto <> "" then
                                x = x &"   <tr>" & chr(13) & _
                                "          <td  class='VISAO_ANALIT' id='table_Desconto_"& idx_bloco &"' style='display: none;' colspan='15' >" & chr(13)& _
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
					        "		<td colspan='13' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
					        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					        "	</tr>" & chr(13) & _
					        "	<tr>" & chr(13) & _
					        "		<td colspan='13' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
					        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					        "	</tr>" & chr(13) & _
					        "	<tr nowrap style='background:honeydew'>" & chr(13) & _
					        "		<td class='MTBE' colspan='5' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					        "TOTAL GERAL:</span></td>" & chr(13) & _
					        "		<td class='MTB' align='right'><span id ='total_VlPedido' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_preco_venda) & "</span></td>" & chr(13) & _
					        "		<td class='MTB' align='right'><span id='totalComissao' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					        "		<td class='MTB' align='right'><span id='total_RA' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA) & "</span></td>" & chr(13) & _
					        "		<td class='MTB' align='right'><span id='total_RAliq' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) & _
					        "		<td class='MTB' align='right'><span id='total_RAdif' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_diferenca) & "</span></td>" & chr(13) & _
					        "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT+vl_total_RA_liquido) & "</span></td>" & chr(13) & _
					        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					        "	</tr>" & chr(13) &_
                            " </table>" & chr(13)
                
			        end if
		        end if

            ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	        if n_reg_total = 0 then
		        x = cab_table & cab
		        x = x & "	<tr nowrap>" & chr(13) & _
				        "		<td class='MT ALERTA' colspan='13' align='center'><span class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS DO INDICADOR ESPECIFICADO&nbsp;</span></td>" & chr(13) & _
				        "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				        "	</tr>" & chr(13)
		        end if

            ' FECHA TABELA DO ÚLTIMO INDICADOR
	        x = x & "</table>" & chr(13)

	        Response.write x

            x = "<input type='hidden' name='c_lista_completa_venda_normal' id='c_lista_completa_venda_normal' value='" & s_lista_completa_venda_normal & "' />" & chr(13) & _
		    "<input type='hidden' name='c_lista_completa_devolucao' id ='c_lista_completa_devolucao' value='" & s_lista_completa_devolucao & "' />" & chr(13) & _
		    "<input type='hidden' name='c_lista_completa_perda' id='c_lista_completa_perda' value='" & s_lista_completa_perda & "' />" & chr(13) & _
		    "<input type='hidden' name='c_lista_ja_marcado_venda_normal' id='c_lista_ja_marcado_venda_normal' value='" & s_lista_ja_marcado_venda_normal & "' />" & chr(13) & _
		    "<input type='hidden' name='c_lista_ja_marcado_devolucao' id='c_lista_ja_marcado_devolucao' value='" & s_lista_ja_marcado_devolucao & "' />" & chr(13) & _
		    "<input type='hidden' name='c_lista_ja_marcado_perda' id='c_lista_ja_marcado_perda' value='" & s_lista_ja_marcado_perda & "' />" & chr(13)

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
	$(".CKB_COM:enabled").each(function () {
	    if ($(this).is(":checked")) {
	        $(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
	    }
	    else {
	        $(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
	    }
	})

    // EVENTO P/ REALÇAR OU NÃO CONFORME SE MARCA/DESMARCA O CHECKBOX
	$(".CKB_COM:enabled").click(function () {
	    if ($(this).is(":checked")) {
	        $(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
	    }
	    else {
	        $(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
	    }
	})
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
    f.action = "RelComissaoIndicadoresPesquisaGravaDados.asp";

    f.submit();
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


<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<input type="hidden" name="orcamento_selecionado" id="orcamento_selecionado" value="">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Comissão Indicadores: Pesquisa Indicador</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<%
	s_filtro = "<table width='709' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)


'	INDICADOR
    s = ""
	if c_cnpj_cpf <> "" then
		s = cnpj_cpf_formata(c_cnpj_cpf)
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>CNPJ/CPF:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

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

<%if alerta="" then%>
<table class="notPrint" width="709" cellpadding="0" cellspacing="0" style="margin-top:5px;">
<tr>
	<td align="left">
		<button type="button" name="bMarcarTodos" id="bMarcarTodos" class="Button BTN_LNK" onclick="marcar_todos();" title="assinala todos os pedidos para gravar o status da comissão como paga" style="margin-left:6px;margin-bottom:2px">Marcar todos</button>
		&nbsp;
		<button type="button" name="bDesmarcarTodos" id="bDesmarcarTodos" class="Button BTN_LNK" onclick="desmarcar_todos();" title="desmarca todos os pedidos para gravar o status da comissão como não-paga" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Desmarcar todos</button>
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
        <%if alerta = "" then %>
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
