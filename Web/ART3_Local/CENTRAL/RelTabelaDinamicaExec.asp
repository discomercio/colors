<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelTabelaDinamicaExec.asp
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
	
	const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
	const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
	const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
	const MSO_NUMBER_FORMAT_TEXTO = "\@"
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	if (usuario = "") then usuario = Trim(Request("c_usuario_sessao"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	if s_lista_operacoes_permitidas = "" then
		s_lista_operacoes_permitidas = obtem_operacoes_permitidas_usuario(cn, usuario)
		Session("lista_operacoes_permitidas") = s_lista_operacoes_permitidas
		end if

'	VERIFICA PERMISS�O DE ACESSO DO USU�RIO
	if Not operacao_permitida(OP_CEN_REL_DADOS_TABELA_DINAMICA, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim c_dt_faturamento_inicio, c_dt_faturamento_termino
	dim c_fabricante, c_grupo, c_potencia_BTU, c_ciclo, c_posicao_mercado
	dim c_loja, rb_tipo_cliente
	dim s, s_aux, s_filtro, s_filtro_loja, lista_loja, v_loja, v, i
    dim ckb_AGRUPAMENTO

	alerta = ""

	c_dt_faturamento_inicio = Trim(Request.Form("c_dt_faturamento_inicio"))
	c_dt_faturamento_termino = Trim(Request.Form("c_dt_faturamento_termino"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_potencia_BTU = Trim(Request.Form("c_potencia_BTU"))
	c_ciclo = Trim(Request.Form("c_ciclo"))
	c_posicao_mercado = Trim(Request.Form("c_posicao_mercado"))
	rb_tipo_cliente = Trim(Request.Form("rb_tipo_cliente"))
	
	c_loja = Trim(Request.Form("c_loja"))
	lista_loja = substitui_caracteres(c_loja,chr(10),"")
	v_loja = split(lista_loja,chr(13),-1)

    ckb_AGRUPAMENTO = Trim(Request.Form("ckb_AGRUPAMENTO"))

'	CAMPOS DE SA�DA SELECIONADOS
	dim ckb_COL_DATA, ckb_COL_NF, ckb_COL_DT_EMISSAO_NF, ckb_COL_PEDIDO, ckb_COL_VENDEDOR, ckb_COL_INDICADOR
	dim ckb_COL_CPF_CNPJ_CLIENTE, ckb_COL_NOME_CLIENTE, ckb_COL_RT
	dim ckb_COL_PRODUTO, ckb_COL_DESCRICAO_PRODUTO, ckb_COL_VL_UNITARIO, ckb_COL_QTDE
	dim ckb_COL_VL_CUSTO, ckb_COL_VL_LISTA, ckb_COL_GRUPO, ckb_COL_POTENCIA_BTU
	dim ckb_COL_CICLO, ckb_COL_POSICAO_MERCADO, ckb_COL_MARCA, ckb_COL_TRANSPORTADORA
	dim ckb_COL_CIDADE, ckb_COL_UF, ckb_COL_QTDE_PARCELAS, ckb_COL_MEIO_PAGAMENTO, ckb_COL_TEL, ckb_COL_EMAIL
    dim ckb_COL_PERC_DESC, ckb_COL_CUBAGEM, ckb_COL_PESO, ckb_COL_FRETE
    dim ckb_COL_INDICADOR_EMAILS, ckb_COL_INDICADOR_CPF_CNPJ, ckb_COL_INDICADOR_ENDERECO, ckb_COL_INDICADOR_CIDADE, ckb_COL_INDICADOR_UF
	
	ckb_COL_DATA = Trim(Request.Form("ckb_COL_DATA"))
	ckb_COL_NF = Trim(Request.Form("ckb_COL_NF"))
	ckb_COL_DT_EMISSAO_NF = Trim(Request.Form("ckb_COL_DT_EMISSAO_NF"))
	ckb_COL_PEDIDO = Trim(Request.Form("ckb_COL_PEDIDO"))
	ckb_COL_VENDEDOR = Trim(Request.Form("ckb_COL_VENDEDOR"))
	ckb_COL_INDICADOR = Trim(Request.Form("ckb_COL_INDICADOR"))
	ckb_COL_CPF_CNPJ_CLIENTE = Trim(Request.Form("ckb_COL_CPF_CNPJ_CLIENTE"))
	ckb_COL_NOME_CLIENTE = Trim(Request.Form("ckb_COL_NOME_CLIENTE"))
	ckb_COL_RT = Trim(Request.Form("ckb_COL_RT"))
	ckb_COL_PRODUTO = Trim(Request.Form("ckb_COL_PRODUTO"))
	ckb_COL_DESCRICAO_PRODUTO = Trim(Request.Form("ckb_COL_DESCRICAO_PRODUTO"))
	ckb_COL_VL_UNITARIO = Trim(Request.Form("ckb_COL_VL_UNITARIO"))
	ckb_COL_QTDE = Trim(Request.Form("ckb_COL_QTDE"))
	ckb_COL_VL_CUSTO = Trim(Request.Form("ckb_COL_VL_CUSTO"))
	ckb_COL_VL_LISTA = Trim(Request.Form("ckb_COL_VL_LISTA"))
	ckb_COL_GRUPO = Trim(Request.Form("ckb_COL_GRUPO"))
	ckb_COL_POTENCIA_BTU = Trim(Request.Form("ckb_COL_POTENCIA_BTU"))
	ckb_COL_CICLO = Trim(Request.Form("ckb_COL_CICLO"))
	ckb_COL_POSICAO_MERCADO = Trim(Request.Form("ckb_COL_POSICAO_MERCADO"))
	ckb_COL_MARCA = Trim(Request.Form("ckb_COL_MARCA"))
	ckb_COL_TRANSPORTADORA = Trim(Request.Form("ckb_COL_TRANSPORTADORA"))
	ckb_COL_CIDADE = Trim(Request.Form("ckb_COL_CIDADE"))
	ckb_COL_UF = Trim(Request.Form("ckb_COL_UF"))
	ckb_COL_QTDE_PARCELAS = Trim(Request.Form("ckb_COL_QTDE_PARCELAS"))
	ckb_COL_MEIO_PAGAMENTO = Trim(Request.Form("ckb_COL_MEIO_PAGAMENTO"))
    ckb_COL_TEL = Trim(Request.Form("ckb_COL_TEL"))
    ckb_COL_EMAIL = Trim(Request.Form("ckb_COL_EMAIL"))
    ckb_COL_PERC_DESC = Trim(Request.Form("ckb_COL_PERC_DESC"))
    ckb_COL_CUBAGEM = Trim(Request.Form("ckb_COL_CUBAGEM"))
    ckb_COL_PESO = Trim(Request.Form("ckb_COL_PESO"))
    ckb_COL_FRETE = Trim(Request.Form("ckb_COL_FRETE"))
    ckb_COL_INDICADOR_EMAILS = Trim(Request.Form("ckb_COL_INDICADOR_EMAILS"))
    ckb_COL_INDICADOR_ENDERECO = Trim(Request.Form("ckb_COL_INDICADOR_ENDERECO"))
    ckb_COL_INDICADOR_CIDADE = Trim(Request.Form("ckb_COL_INDICADOR_CIDADE"))
    ckb_COL_INDICADOR_UF = Trim(Request.Form("ckb_COL_INDICADOR_UF"))
    ckb_COL_INDICADOR_CPF_CNPJ = Trim(Request.Form("ckb_COL_INDICADOR_CPF_CNPJ"))
	
	dim s_campos_saida
	s_campos_saida = "|"
	if alerta = "" then
		if ckb_COL_DATA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_DATA" & "|"
		if ckb_COL_NF <> "" then s_campos_saida = s_campos_saida & "ckb_COL_NF" & "|"
		if ckb_COL_DT_EMISSAO_NF <> "" then s_campos_saida = s_campos_saida & "ckb_COL_DT_EMISSAO_NF" & "|"
		if ckb_COL_PEDIDO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_PEDIDO" & "|"
		if ckb_COL_CPF_CNPJ_CLIENTE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_CPF_CNPJ_CLIENTE" & "|"
		if ckb_COL_NOME_CLIENTE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_NOME_CLIENTE" & "|"
		if ckb_COL_CIDADE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_CIDADE" & "|"
		if ckb_COL_UF <> "" then s_campos_saida = s_campos_saida & "ckb_COL_UF" & "|"
        if ckb_COL_TEL <> "" then s_campos_saida = s_campos_saida & "ckb_COL_TEL" & "|"
        if ckb_COL_EMAIL <> "" then s_campos_saida = s_campos_saida & "ckb_COL_EMAIL" & "|"
		if ckb_COL_VENDEDOR <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VENDEDOR" & "|"
		if ckb_COL_INDICADOR <> "" then s_campos_saida = s_campos_saida & "ckb_COL_INDICADOR" & "|"
		if ckb_COL_TRANSPORTADORA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_TRANSPORTADORA" & "|"
		if ckb_COL_INDICADOR_CPF_CNPJ <> "" then s_campos_saida = s_campos_saida & "ckb_COL_INDICADOR_CPF_CNPJ" & "|"
		if ckb_COL_INDICADOR_ENDERECO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_INDICADOR_ENDERECO" & "|"
		if ckb_COL_INDICADOR_CIDADE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_INDICADOR_CIDADE" & "|"
		if ckb_COL_INDICADOR_UF <> "" then s_campos_saida = s_campos_saida & "ckb_COL_INDICADOR_UF" & "|"
		if ckb_COL_INDICADOR_EMAILS <> "" then s_campos_saida = s_campos_saida & "ckb_COL_INDICADOR_EMAILS" & "|"
		if ckb_COL_MARCA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_MARCA" & "|"
		if ckb_COL_GRUPO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_GRUPO" & "|"
		if ckb_COL_POTENCIA_BTU <> "" then s_campos_saida = s_campos_saida & "ckb_COL_POTENCIA_BTU" & "|"
		if ckb_COL_CICLO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_CICLO" & "|"
		if ckb_COL_POSICAO_MERCADO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_POSICAO_MERCADO" & "|"
		if ckb_COL_PRODUTO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_PRODUTO" & "|"
		if ckb_COL_DESCRICAO_PRODUTO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_DESCRICAO_PRODUTO" & "|"
		if ckb_COL_QTDE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_QTDE" & "|"
        if ckb_COL_PERC_DESC <> "" then s_campos_saida = s_campos_saida & "ckb_COL_PERC_DESC" & "|"
        if ckb_COL_CUBAGEM <> "" then s_campos_saida = s_campos_saida & "ckb_COL_CUBAGEM" & "|"
        if ckb_COL_PESO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_PESO" & "|"
        if ckb_COL_FRETE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_FRETE" & "|"
		if ckb_COL_VL_CUSTO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_CUSTO" & "|"
		if ckb_COL_VL_LISTA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_LISTA" & "|"
		if ckb_COL_VL_UNITARIO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_UNITARIO" & "|"
		if ckb_COL_RT <> "" then s_campos_saida = s_campos_saida & "ckb_COL_RT" & "|"	
		if ckb_COL_QTDE_PARCELAS <> "" then s_campos_saida = s_campos_saida & "ckb_COL_QTDE_PARCELAS" & "|"
		if ckb_COL_MEIO_PAGAMENTO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_MEIO_PAGAMENTO" & "|"
		
		
		
		
		if s_campos_saida = "|" then s_campos_saida = "NENHUM"
		call set_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|campos_saida_selecionados", s_campos_saida)
		
		call set_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_dt_faturamento_inicio", c_dt_faturamento_inicio)
		call set_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_dt_faturamento_termino", c_dt_faturamento_termino)
		call set_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_fabricante", c_fabricante)
		call set_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_grupo", c_grupo)
		call set_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_potencia_BTU", c_potencia_BTU)
		call set_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_ciclo", c_ciclo)
		call set_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_posicao_mercado", c_posicao_mercado)
		call set_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|rb_tipo_cliente", rb_tipo_cliente)
		call set_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_loja", c_loja)
		end if
	
	if alerta = "" then
		Response.ContentType = "application/csv"
		Response.AddHeader "Content-Disposition", "attachment; filename=TabDinamica_" & formata_data_yyyymmdd(Now) & "_" & substitui_caracteres(formata_hora(Now),":","") & ".csv"
		consulta_executa
		Response.End
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  �  �  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const SEPARADOR_DECIMAL = ","
dim s, s_sql, x, x_cab, s_where, s_where_venda, s_where_devolucao, s_where_loja
dim perc_RT, vl_RT, vl_preco_venda, n_reg, n_reg_total
dim r
dim tipo_parc
dim s_qtde, item_peso, item_cubagem, item_qtde

'	CRIT�RIOS COMUNS
'	================
	s_where = ""
	s_where_venda = ""
	s_where_devolucao = ""

    item_peso = 0
    item_cubagem = 0
    item_qtde = 0
	
	if IsDate(c_dt_faturamento_inicio) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data >= " & bd_formata_data(StrToDate(c_dt_faturamento_inicio)) & ")"
		
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(StrToDate(c_dt_faturamento_inicio)) & ")"
		end if
		
	if IsDate(c_dt_faturamento_termino) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data < " & bd_formata_data(StrToDate(c_dt_faturamento_termino)+1) & ")"
		
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(StrToDate(c_dt_faturamento_termino)+1) & ")"
		end if
	
	if c_fabricante <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO_ITEM.fabricante = '" & c_fabricante & "')"
		
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_grupo <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.grupo = '" & c_grupo & "')"
		end if
	
	if c_potencia_BTU <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.potencia_BTU = " & c_potencia_BTU & ")"
		end if
	
	if c_ciclo <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.ciclo = '" & c_ciclo & "')"
		end if
	
	if c_posicao_mercado <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PRODUTO.posicao_mercado = '" & c_posicao_mercado & "')"
		end if
	
	if rb_tipo_cliente <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_CLIENTE.tipo = '" & rb_tipo_cliente & "')"
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


'	MONTA CONSULTA
'	==============
'	VENDAS NORMAIS
	s_sql = "SELECT" & _
				" t_PEDIDO.data_hora," & _
				" t_PEDIDO.entregue_data AS faturamento_data," & _
				" (SELECT TOP 1 dt_emissao FROM t_NFe_EMISSAO WHERE (t_NFe_EMISSAO.pedido = t_PEDIDO.pedido) AND (t_NFe_EMISSAO.tipo_NF = '1') AND (t_NFe_EMISSAO.st_anulado = 0) AND (t_NFe_EMISSAO.codigo_retorno_NFe_T1 = 1) ORDER BY id) AS dt_emissao," & _
				" (SELECT TOP 1 NFe_numero_NF FROM t_NFe_EMISSAO WHERE (t_NFe_EMISSAO.pedido = t_PEDIDO.pedido) AND (t_NFe_EMISSAO.tipo_NF = '1') AND (t_NFe_EMISSAO.st_anulado = 0) AND (t_NFe_EMISSAO.codigo_retorno_NFe_T1 = 1) ORDER BY id) AS numero_NF," & _
				" t_PEDIDO.pedido," & _
				" t_PEDIDO.obs_2," & _
				" t_PEDIDO.transportadora_id," & _
				" t_PEDIDO__BASE.vendedor," & _
				" t_PEDIDO__BASE.indicador," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas AS nome_cliente," & _
				" t_CLIENTE.cnpj_cpf," & _
				" t_PEDIDO__BASE.perc_RT," & _
				" t_PEDIDO_ITEM.qtde," & _
				" t_PEDIDO_ITEM.fabricante," & _
				" t_PEDIDO_ITEM.produto," & _
				" t_PEDIDO_ITEM.descricao," & _
				" t_PEDIDO_ITEM.preco_venda," & _
				" t_PEDIDO_ITEM.preco_lista," & _
                " t_PEDIDO_ITEM.desc_dado," & _
                " t_PEDIDO_ITEM.cubagem," & _
                " t_PEDIDO_ITEM.peso," & _
				" t_PRODUTO.grupo," & _
				" t_PRODUTO.potencia_BTU," & _
				" t_PRODUTO.ciclo," & _
				" t_PRODUTO.posicao_mercado," & _
				" t_FABRICANTE.nome AS nome_fabricante," & _
				" t_CLIENTE.cidade AS cidade," & _
				" t_CLIENTE.uf AS uf," & _
                " t_CLIENTE.tipo AS tipo_cliente," & _
                " t_CLIENTE.ddd_res," & _
                " t_CLIENTE.tel_res," & _
                " t_CLIENTE.ddd_cel," & _
                " t_CLIENTE.tel_cel," & _
                " t_CLIENTE.ddd_com," & _
                " t_CLIENTE.tel_com," & _
                " t_CLIENTE.ddd_com_2," & _
                " t_CLIENTE.tel_com_2," & _
                " t_CLIENTE.ramal_com," & _
                " t_CLIENTE.ramal_com_2," & _
                " t_CLIENTE.email," & _
                " t_ORCAMENTISTA_E_INDICADOR.cnpj_cpf AS indicador_cnpj_cpf," & _
                " t_ORCAMENTISTA_E_INDICADOR.endereco AS indicador_endereco," & _
                " t_ORCAMENTISTA_E_INDICADOR.endereco_numero AS indicador_endereco_numero," & _
                " t_ORCAMENTISTA_E_INDICADOR.endereco_complemento AS indicador_endereco_complemento," & _
                " t_ORCAMENTISTA_E_INDICADOR.bairro AS indicador_bairro," & _
                " t_ORCAMENTISTA_E_INDICADOR.cidade AS indicador_cidade," & _
                " t_ORCAMENTISTA_E_INDICADOR.uf AS indicador_uf," & _
                " t_ORCAMENTISTA_E_INDICADOR.cep AS indicador_cep," & _
                " t_ORCAMENTISTA_E_INDICADOR.email AS indicador_email," & _
                " t_ORCAMENTISTA_E_INDICADOR.email2 AS indicador_email2," & _
                " t_ORCAMENTISTA_E_INDICADOR.email3 AS indicador_email3," & _
				" t_PEDIDO__BASE.qtde_parcelas AS qtde_parcelas," & _
				" t_PEDIDO__BASE.tipo_parcelamento AS tipo_parcelamento," & _
				" t_PEDIDO__BASE.av_forma_pagto AS forma_pagamento_av," & _
				" t_PEDIDO__BASE.pce_forma_pagto_prestacao AS parcelamento_c_entrada," & _
				" t_PEDIDO__BASE.pse_forma_pagto_demais_prest AS parcelamento_s_entrada," & _
				" t_PEDIDO__BASE.pu_forma_pagto AS parcela_unica," & _
                " (SELECT SUM(vl_frete) AS vl_frete FROM t_PEDIDO_FRETE WHERE (pedido=t_PEDIDO.pedido)) AS vl_frete," & _
				" (SELECT TOP 1 vl_custo2 FROM t_ESTOQUE_ITEM tEI WHERE (tEI.fabricante = t_PEDIDO_ITEM.fabricante) AND (tEI.produto = t_PEDIDO_ITEM.produto) ORDER BY id_estoque DESC) AS vl_custo2" & _
			" FROM t_PEDIDO" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base = t_PEDIDO__BASE.pedido)" & _
                " LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador = t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
				" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido = t_PEDIDO_ITEM.pedido)" & _
				" INNER JOIN t_PRODUTO ON (t_PRODUTO.fabricante = t_PEDIDO_ITEM.fabricante) AND (t_PRODUTO.produto = t_PEDIDO_ITEM.produto)" & _
				" INNER JOIN t_FABRICANTE ON (t_PRODUTO.fabricante = t_FABRICANTE.fabricante)" & _
				" INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente = t_CLIENTE.id)" & _
			" WHERE" & _
				" (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')"
	
	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then
		s_sql = s_sql & " AND" & s
		end if
	
'	DEVOLU��ES
'	OBS: O USO DE 'UNION' SIMPLES ELIMINA AS LINHAS DUPLICADAS DOS RESULTADOS
'		 O USO DE 'UNION ALL' RETORNARIA TODAS AS LINHAS, INCLUSIVE AS DUPLICADAS
	s_sql = s_sql & " UNION ALL " & _
			"SELECT" & _
				" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data_hora," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS faturamento_data," & _
				" NULL AS dt_emissao," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.NFe_numero_NF AS numero_NF," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.pedido," & _
				" t_PEDIDO.obs_2," & _
				" t_PEDIDO.transportadora_id," & _
				" t_PEDIDO__BASE.vendedor," & _
				" t_PEDIDO__BASE.indicador," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas AS nome_cliente," & _
				" t_CLIENTE.cnpj_cpf," & _
				" t_PEDIDO__BASE.perc_RT," & _
				" -t_PEDIDO_ITEM_DEVOLVIDO.qtde," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.fabricante," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.produto," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.descricao," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.preco_venda," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.preco_lista," & _
                " t_PEDIDO_ITEM_DEVOLVIDO.desc_dado," & _
                " t_PEDIDO_ITEM_DEVOLVIDO.cubagem," & _
                " t_PEDIDO_ITEM_DEVOLVIDO.peso," & _
				" t_PRODUTO.grupo," & _
				" t_PRODUTO.potencia_BTU," & _
				" t_PRODUTO.ciclo," & _
				" t_PRODUTO.posicao_mercado," & _
				" t_FABRICANTE.nome AS nome_fabricante," & _
				" t_CLIENTE.cidade AS cidade," & _
				" t_CLIENTE.uf AS uf," & _
                " t_CLIENTE.tipo AS tipo_cliente," & _
                " t_CLIENTE.ddd_res," & _
                " t_CLIENTE.tel_res," & _
                " t_CLIENTE.ddd_cel," & _
                " t_CLIENTE.tel_cel," & _
                " t_CLIENTE.ddd_com," & _
                " t_CLIENTE.tel_com," & _
                " t_CLIENTE.ddd_com_2," & _
                " t_CLIENTE.tel_com_2," & _
                " t_CLIENTE.ramal_com," & _
                " t_CLIENTE.ramal_com_2," & _
                " t_CLIENTE.email," & _
                " t_ORCAMENTISTA_E_INDICADOR.cnpj_cpf AS indicador_cnpj_cpf," & _
                " t_ORCAMENTISTA_E_INDICADOR.endereco AS indicador_endereco," & _
                " t_ORCAMENTISTA_E_INDICADOR.endereco_numero AS indicador_endereco_numero," & _
                " t_ORCAMENTISTA_E_INDICADOR.endereco_complemento AS indicador_endereco_complemento," & _
                " t_ORCAMENTISTA_E_INDICADOR.bairro AS indicador_bairro," & _
                " t_ORCAMENTISTA_E_INDICADOR.cidade AS indicador_cidade," & _
                " t_ORCAMENTISTA_E_INDICADOR.uf AS indicador_uf," & _
                " t_ORCAMENTISTA_E_INDICADOR.cep AS indicador_cep," & _
                " t_ORCAMENTISTA_E_INDICADOR.email AS indicador_email," & _
                " t_ORCAMENTISTA_E_INDICADOR.email2 AS indicador_email2," & _
                " t_ORCAMENTISTA_E_INDICADOR.email3 AS indicador_email3," & _
				" t_PEDIDO.qtde_parcelas AS qtde_parcelas," & _
				" t_PEDIDO__BASE.tipo_parcelamento AS tipo_parcelamento," & _
				" t_PEDIDO__BASE.av_forma_pagto AS forma_pagamento_av," & _
				" t_PEDIDO__BASE.pce_forma_pagto_prestacao AS parcelamento_c_entrada," & _
				" t_PEDIDO__BASE.pse_forma_pagto_demais_prest AS parcelamento_s_entrada," & _
				" t_PEDIDO__BASE.pu_forma_pagto AS parcela_unica," & _
                " (SELECT SUM(vl_frete) AS vl_frete FROM t_PEDIDO_FRETE WHERE (pedido=t_PEDIDO.pedido)) AS vl_frete," & _
				" (SELECT TOP 1 vl_custo2 FROM t_ESTOQUE_ITEM tEI WHERE (tEI.fabricante = t_PEDIDO_ITEM_DEVOLVIDO.fabricante) AND (tEI.produto = t_PEDIDO_ITEM_DEVOLVIDO.produto) ORDER BY id_estoque DESC) AS vl_custo2" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
				" INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido = t_PEDIDO.pedido)" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base = t_PEDIDO__BASE.pedido)" & _
                " LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador = t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
				" INNER JOIN t_PRODUTO ON (t_PRODUTO.fabricante = t_PEDIDO_ITEM_DEVOLVIDO.fabricante) AND (t_PRODUTO.produto = t_PEDIDO_ITEM_DEVOLVIDO.produto)" & _
				" INNER JOIN t_FABRICANTE ON (t_PRODUTO.fabricante = t_FABRICANTE.fabricante)" & _
				" INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente = t_CLIENTE.id)" & _
			" WHERE" & _
				" (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')"
	
	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then
		s_sql = s_sql & " AND" & s
		end if
	
	s_sql = "SELECT " & _
				"*" & _
			" FROM (" & s_sql & ") t"
	
	s_sql = s_sql & _
			" ORDER BY" & _
				" faturamento_data," & _
				" pedido," & _
				" fabricante," & _
				" produto"
	
	x_cab = ""
	if ckb_COL_DATA <> "" then x_cab = x_cab & "DATA;"
	if ckb_COL_NF <> "" then x_cab = x_cab & "NF;"
	if ckb_COL_DT_EMISSAO_NF <> "" then x_cab = x_cab & "EMISSAO NF;"
	if ckb_COL_PEDIDO <> "" then x_cab = x_cab & "PEDIDO;"
	if ckb_COL_CPF_CNPJ_CLIENTE <> "" then x_cab = x_cab & "CPF/CNPJ;"
	if ckb_COL_NOME_CLIENTE <> "" then x_cab = x_cab & "NOME CLIENTE;"
	if ckb_COL_CIDADE <> "" then x_cab = x_cab & "CIDADE;"
	if ckb_COL_UF <> "" then x_cab = x_cab & "UF;"
    if ckb_COL_TEL <> "" then x_cab = x_cab & "TELEFONE;TELEFONE;TELEFONE;"
    if ckb_COL_EMAIL <> "" then x_cab = x_cab & "EMAIL;"
	if ckb_COL_VENDEDOR <> "" then x_cab = x_cab & "VENDEDOR;"
	if ckb_COL_INDICADOR <> "" then x_cab = x_cab & "INDICADOR;"
	if ckb_COL_TRANSPORTADORA <> "" then x_cab = x_cab & "TRANSPORTADORA;"
	if ckb_COL_INDICADOR_CPF_CNPJ <> "" then x_cab = x_cab & "CPF/CNPJ IND;"
	if ckb_COL_INDICADOR_ENDERECO <> "" then x_cab = x_cab & "ENDERECO IND;"
	if ckb_COL_INDICADOR_CIDADE <> "" then x_cab = x_cab & "CIDADE IND;"
	if ckb_COL_INDICADOR_UF <> "" then x_cab = x_cab & "UF IND;"
	if ckb_COL_INDICADOR_EMAILS <> "" then x_cab = x_cab & "EMAIL IND;EMAIL IND;EMAIL IND;"
	if ckb_COL_MARCA <> "" then x_cab = x_cab & "MARCA;"
	if ckb_COL_GRUPO <> "" then x_cab = x_cab & "GRUPO;"
	if ckb_COL_POTENCIA_BTU <> "" then x_cab = x_cab & "BTU;"
	if ckb_COL_CICLO <> "" then x_cab = x_cab & "CICLO;"
	if ckb_COL_POSICAO_MERCADO <> "" then x_cab = x_cab & "POS MERC;"
	if ckb_COL_PRODUTO <> "" then x_cab = x_cab & "PRODUTO;"
	if ckb_COL_DESCRICAO_PRODUTO <> "" then x_cab = x_cab & "DESCRICAO;"
	if ckb_COL_QTDE <> "" then x_cab = x_cab & "QTDE;"
    if ckb_COL_PERC_DESC <> "" then x_cab = x_cab & "DESC %;"
    if ckb_COL_CUBAGEM <> "" then x_cab = x_cab & "CUBAGEM;"
    if ckb_COL_PESO <> "" then x_cab = x_cab & "PESO;"
    if ckb_COL_FRETE <> "" then x_cab = x_cab & "VL FRETE;"
	if ckb_COL_VL_CUSTO <> "" then x_cab = x_cab & "VL CUSTO;"
	if ckb_COL_VL_LISTA <> "" then x_cab = x_cab & "VL LISTA;"
	if ckb_COL_VL_UNITARIO <> "" then x_cab = x_cab & "VL UNITARIO;"
	if ckb_COL_RT <> "" then x_cab = x_cab & "RT;"
	if ckb_COL_QTDE_PARCELAS <> "" then x_cab = x_cab & "QTDE PARCELAS;"
	if ckb_COL_MEIO_PAGAMENTO <> "" then x_cab = x_cab & "MEIO DE PAGAMENTO;"
	
	
	
	
	x = ""
	n_reg = 0
	n_reg_total = 0
    item_qtde = 1

	set r = cn.execute(s_sql)
	
	if Not r.Eof then x = x_cab & vbCrLf
	
	do while Not r.Eof
		
	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

        if ckb_AGRUPAMENTO <> "" then
            item_qtde = CInt(Trim("" & r("qtde")))
        end if

        for i=1 to Abs(item_qtde)

	 '> DATA
		if ckb_COL_DATA <> "" then
			x = x & formata_data(r("faturamento_data")) & ";"
			end if
		
	 '> NF
		if ckb_COL_NF <> "" then
			s = Trim("" & r("numero_NF"))
			if s = "" then
				s = Trim("" & r("obs_2"))
				end if
			x = x & s & ";"
			end if
		
	'> DATA DA EMISS�O NF
		if ckb_COL_DT_EMISSAO_NF <> "" then
			if Trim("" & r("dt_emissao")) <> "" then 
				s = formata_data(r("dt_emissao"))
			else
				s = ""
				end if
			x = x & s & ";"
			end if

	 '> PEDIDO
		if ckb_COL_PEDIDO <> "" then
			x = x & Trim("" & r("pedido")) & ";"
			end if
			
	'> CLIENTE: CPF/CNPJ
		if ckb_COL_CPF_CNPJ_CLIENTE <> "" then
			x = x & cnpj_cpf_formata(Trim("" & r("cnpj_cpf"))) & ";"
			end if
			
	'> CLIENTE: NOME
		if ckb_COL_NOME_CLIENTE <> "" then
			s = Trim("" & r("nome_cliente"))
			s = substitui_caracteres(s, ";", ",")
			x = x & s & ";"
			end if
			
	 '> CIDADE
	    if ckb_COL_CIDADE <> "" then
	        x = x & Trim("" & r("cidade")) & ";"
	    end if
	 
	 '> UF
	    if ckb_COL_UF <> "" then
	        x = x & Trim("" & r("uf")) & ";"
	    end if

     '> TELEFONES
        if ckb_COL_TEL <> "" then
            if CStr(r("tipo_cliente")) = ID_PF then
                x = x & iif( (Trim("" & r("ddd_res")) <> ""),   "(" & Trim("" & r("ddd_res")) & ") " & Trim("" & r("tel_res")) & ";",   ";" )
                x = x & iif( (Trim("" & r("ddd_cel")) <> ""),   "(" & Trim("" & r("ddd_cel")) & ") " & Trim("" & r("tel_cel")) & ";",   ";" )
                x = x & iif( (Trim("" & r("ddd_com")) <> ""),   "(" & Trim("" & r("ddd_com")) & ") " & Trim("" & r("tel_com")),   "" )
                x = x & iif( (Trim("" & r("ramal_com")) <> ""),   " R:" & Trim("" & r("ramal_com")) & ";",  ";" )
            elseif CStr(r("tipo_cliente")) = ID_PJ then
                x = x & iif( (Trim("" & r("ddd_com")) <> ""),   "(" & Trim("" & r("ddd_com")) & ") " & Trim("" & r("tel_com")),   "" )
                x = x & iif( (Trim("" & r("ramal_com")) <> ""),   " R:" & Trim("" & r("ramal_com")) & ";",  ";" )   
                x = x & iif( (Trim("" & r("ddd_com_2")) <> ""),   "(" & Trim("" & r("ddd_com_2")) & ") " & Trim("" & r("tel_com_2")),   "" )
                x = x & iif( (Trim("" & r("ramal_com_2")) <> ""),   " R:" & Trim("" & r("ramal_com_2")) & ";",   ";" )  
                x = x & ";"             
            end if
        end if
     
    '> E-MAIL
		if ckb_COL_EMAIL <> "" then
			x = x & Trim("" & r("email")) & ";"
			end if        
		
	 '> VENDEDOR
		if ckb_COL_VENDEDOR <> "" then
			x = x & Trim("" & r("vendedor")) & ";"
			end if
		
	 '> INDICADOR
		if ckb_COL_INDICADOR <> "" then
			x = x & Trim("" & r("indicador")) & ";"
			end if
		
	 '> TRANSPORTADORA
		if ckb_COL_TRANSPORTADORA <> "" then
			x = x & UCase(Trim("" & r("transportadora_id"))) & ";"
			end if 

    '> INDICADOR: CPF/CNPJ
		if ckb_COL_INDICADOR_CPF_CNPJ <> "" then
			x = x & cnpj_cpf_formata(Trim("" & r("indicador_cnpj_cpf"))) & ";"
			end if 

    '> INDICADOR: ENDERE�O
        if ckb_COL_INDICADOR_ENDERECO <> "" then
            x = x & formata_endereco(Trim("" & r("indicador_endereco")), Trim("" & r("indicador_endereco_numero")), Trim("" & r("indicador_endereco_complemento")), Trim("" & r("indicador_bairro")), "", "", Trim("" & r("indicador_cep"))) & ";"
            end if
			
    '> INDICADOR: CIDADE
	    if ckb_COL_INDICADOR_CIDADE <> "" then
	        x = x & Trim("" & r("indicador_cidade")) & ";"
	    end if
	 
	 '> INDICADOR: UF
	    if ckb_COL_INDICADOR_UF <> "" then
	        x = x & Trim("" & r("indicador_uf")) & ";"
	    end if

    '> INDICADOR: E-MAIL 
		if ckb_COL_INDICADOR_EMAILS <> "" then
			x = x & Trim("" & r("indicador_email")) & ";"
			end if

    '> INDICADOR: E-MAIL 2
		if ckb_COL_INDICADOR_EMAILS <> "" then
			x = x & Trim("" & r("indicador_email2")) & ";"
			end if

    '> INDICADOR: E-MAIL 3
		if ckb_COL_INDICADOR_EMAILS <> "" then
			x = x & Trim("" & r("indicador_email3")) & ";"
			end if

	'> NOME FABRICANTE
		if ckb_COL_MARCA <> "" then
			x = x & UCase(Trim("" & r("nome_fabricante"))) & ";"
			end if
			
	'> GRUPO
		if ckb_COL_GRUPO <> "" then
			x = x & Trim("" & r("grupo")) & ";"
			end if
		
	 '> BTU
		if ckb_COL_POTENCIA_BTU <> "" then
			s = Trim("" & r("potencia_BTU"))
			if s = "0" then s = ""
			x = x & s & ";"
			end if
		
	 '> CICLO
		if ckb_COL_CICLO <> "" then
			x = x & Trim("" & r("ciclo")) & ";"
			end if
		
	 '> POSI��O MERCADO
		if ckb_COL_POSICAO_MERCADO <> "" then
			x = x & Trim("" & r("posicao_mercado")) & ";"
			end if
		
     '> C�DIGO DO PRODUTO
		if ckb_COL_PRODUTO <> "" then
		 '	FOR�A P/ SER TRATADO COMO TEXTO
			x = x & chr(34) & "=" & chr(34) & chr(34) & Trim("" & r("produto")) & chr(34) & chr(34) & chr(34) & ";"
			end if
		
	 '> DESCRI��O DO PRODUTO
		if ckb_COL_DESCRICAO_PRODUTO <> "" then
			s = Trim("" & r("descricao"))
			s = substitui_caracteres(s, ";", ",")
			x = x & s & ";"
			end if
	
    ' DESMEMBRAR ITENS ?
	    if ckb_AGRUPAMENTO <> "" then
            if CInt(Trim("" & r("qtde"))) < 0 then
                s_qtde = CInt(Trim("" & r("qtde"))) / CInt(Trim("" & r("qtde"))) * (-1)
            else 
                s_qtde = CInt(Trim("" & r("qtde"))) / CInt(Trim("" & r("qtde")))
            end if
        else
		    s_qtde = Trim("" & r("qtde"))
        end if

    '> QTDE
		if ckb_COL_QTDE <> "" then
            x = x & s_qtde & ";"                 
		end if

    '> PERCENTUAL DE DESCONTO
        if ckb_COL_PERC_DESC <> "" then
            x = x & formata_perc_desc(Trim("" & r("desc_dado"))) & ";"
        end if

    '> CUBAGEM
        if ckb_COL_CUBAGEM <> "" then
            item_cubagem = converte_numero(s_qtde) * converte_numero(r("cubagem"))
            x = x & formata_numero(item_cubagem, 2) & ";"
        end if

    '> PESO
        if ckb_COL_PESO <> "" then
            item_peso = s_qtde * r("peso")
            x = x & item_peso & ";"
        end if

    '> FRETE
        if ckb_COL_FRETE <> "" then
            s = formata_moeda(Trim("" & r("vl_frete")))
            if s = "" then s = 0
            x = x & s & ";"       
        end if
			
	'> VALOR CUSTO
		if ckb_COL_VL_CUSTO <> "" then
		'	EXPORTAR VALOR UTILIZANDO '.' COMO SEPARADOR DECIMAL
			s = substitui_caracteres(bd_formata_moeda(r("vl_custo2")), ".", SEPARADOR_DECIMAL)
			x = x & s & ";"
			end if
		
	 '> PRE�O DE LISTA
		if ckb_COL_VL_LISTA <> "" then
		'	EXPORTAR VALOR UTILIZANDO '.' COMO SEPARADOR DECIMAL
			s = substitui_caracteres(bd_formata_moeda(r("preco_lista")), ".", SEPARADOR_DECIMAL)
			x = x & s & ";"
			end if
			
	'> VALOR UNIT�RIO
		if ckb_COL_VL_UNITARIO <> "" then
		'	EXPORTAR VALOR UTILIZANDO '.' COMO SEPARADOR DECIMAL
			s = substitui_caracteres(bd_formata_moeda(r("preco_venda")), ".", SEPARADOR_DECIMAL)
			x = x & s & ";"
			end if
		
	 '> RT
		perc_RT = r("perc_RT")
	'	EVITA DIFEREN�AS DE ARREDONDAMENTO
		vl_preco_venda = converte_numero(formata_moeda(r("preco_venda")))
		vl_RT = (perc_RT/100) * vl_preco_venda
		if ckb_COL_RT <> "" then
			s = substitui_caracteres(bd_formata_moeda(vl_RT), ".", SEPARADOR_DECIMAL)
			x = x & s & ";"
			end if
			
	 '> QTDE DE PARCELAS

	    if ckb_COL_QTDE_PARCELAS <> "" then
	        x = x & Trim("" & r("qtde_parcelas")) & ";"
	    end if
	    
     '> MEIO DE PAGAMENTO
        if ckb_COL_MEIO_PAGAMENTO <> "" then
        	tipo_parc = Trim("" & r("tipo_parcelamento"))
            if tipo_parc = 1 then       
                 s = x_opcao_forma_pagamento(Trim("" & r("forma_pagamento_av"))) 
            elseif tipo_parc = 2 then    
                     s = x_opcao_forma_pagamento(Trim(ID_FORMA_PAGTO_CARTAO))
            elseif tipo_parc = 3 then    
                     s = x_opcao_forma_pagamento(Trim("" & r("parcelamento_c_entrada"))) 
            elseif tipo_parc = 4 then    
                     s = x_opcao_forma_pagamento(Trim("" & r("parcelamento_s_entrada"))) 
            elseif tipo_parc = 5 then     
                    s = x_opcao_forma_pagamento(Trim("" & r("parcela_unica"))) 
			elseif tipo_parc = 6 then
                     s = x_opcao_forma_pagamento(Trim(ID_FORMA_PAGTO_CARTAO_MAQUINETA))
            end if          
            x = x & s & ";"
        end if            
                
		
		x = x & vbCrLf
		
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if

        next
		
		r.MoveNext
		loop
		

	if r.State <> 0 then r.Close
	set r=nothing
	
'	MOSTRA AVISO DE QUE N�O H� DADOS!!
	if n_reg_total = 0 then
		x = "NENHUM PRODUTO ENCONTRADO"
		end if
	
	Response.write x
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



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';
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
<!-- *************************************************************************** -->
<!-- **********  P�GINA PARA EXIBIR RESULTADO   (apenas para testes)  ********** -->
<!-- *************************************************************************** -->
<body onload="window.status='Conclu�do';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_faturamento_inicio" id="c_dt_faturamento_inicio" value="<%=c_dt_faturamento_inicio%>">
<input type="hidden" name="c_dt_faturamento_termino" id="c_dt_faturamento_termino" value="<%=c_dt_faturamento_termino%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_grupo" id="c_grupo" value="<%=c_grupo%>">
<input type="hidden" name="c_potencia_BTU" id="c_potencia_BTU" value="<%=c_potencia_BTU%>">
<input type="hidden" name="c_ciclo" id="c_ciclo" value="<%=c_ciclo%>">
<input type="hidden" name="c_posicao_mercado" id="c_posicao_mercado" value="<%=c_posicao_mercado%>">
<input type="hidden" name="rb_tipo_cliente" id="rb_tipo_cliente" value="<%=rb_tipo_cliente%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">


<!--  I D E N T I F I C A � � O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Dados para Tabela Din�mica</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
'	PER�ODO
	s = ""
	s_aux = c_dt_faturamento_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_faturamento_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Per�odo:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
'	FABRICANTE
	s = c_fabricante
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Fabricante:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	GRUPO
	s = c_grupo
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Grupo:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	POT�NCIA (BTU/h)
	s = c_potencia_BTU
	if (s = "") Or (s = "0") then s = "N.I." else s = formata_inteiro(s)
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>BTU/h:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	CICLO
	s = c_ciclo
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Ciclo:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	POSI��O MERCADO
	s = c_posicao_mercado
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Posi��o Mercado:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	TIPO DE CLIENTE
	s = rb_tipo_cliente
	if s = "" then
		s = "todos"
	elseif s = ID_PF then
		s = "Pessoa F�sica"
	elseif s = ID_PJ then
		s = "Pessoa Jur�dica"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Tipo de Cliente:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

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
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Loja(s):&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
'	EMISS�O
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Emiss�o:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & formata_data_hora(Now) & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELAT�RIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align='left'>&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a p�gina anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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
