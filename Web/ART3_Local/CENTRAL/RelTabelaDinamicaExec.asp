<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/global.asp"    -->
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
	
	Server.ScriptTimeout = 2 * MAX_SERVER_SCRIPT_EXTENDED_TIMEOUT_EM_SEG
	
	const ID_RELATORIO = "RelTabelaDinamicaFiltro"

	const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
	const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
	const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
	const MSO_NUMBER_FORMAT_TEXTO = "\@"
	
	class cl_TIPO_FRETE
		dim codigo
		dim descricao
		dim colSQL
		dim blnFreteDevolucao
		end class

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	if (usuario = "") then usuario = Trim(Request("c_usuario_sessao"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	CONECTA COM O BANCO DE DADOS
	dim cn, r, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
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

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta, msg_erro_consistencia_parametros
	dim c_dt_faturamento_inicio, c_dt_faturamento_termino, c_dt_NF_venda_inicio, c_dt_NF_venda_termino, c_dt_NF_remessa_inicio, c_dt_NF_remessa_termino
	dim ckb_pedido_nao_recebido_pelo_cliente, ckb_pedido_recebido_pelo_cliente
	dim c_fabricante, c_grupo, c_potencia_BTU, c_ciclo, c_posicao_mercado, c_grupo_pedido_origem, c_entrega_imediata
	dim c_loja, rb_tipo_cliente
	dim s, s_aux, s_filtro, s_filtro_loja, lista_loja, v_loja, v, i
	dim v_grupo_pedido_origem
    dim ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET, ckb_AGRUPAMENTO, ckb_CONSOLIDAR_PEDIDO
	dim ckb_COMPATIBILIDADE

	alerta = ""
	msg_erro_consistencia_parametros = ""

	c_dt_faturamento_inicio = Trim(Request.Form("c_dt_faturamento_inicio"))
	c_dt_faturamento_termino = Trim(Request.Form("c_dt_faturamento_termino"))
	c_dt_NF_venda_inicio = Trim(Request.Form("c_dt_NF_venda_inicio"))
	c_dt_NF_venda_termino = Trim(Request.Form("c_dt_NF_venda_termino"))
	c_dt_NF_remessa_inicio = Trim(Request.Form("c_dt_NF_remessa_inicio"))
	c_dt_NF_remessa_termino = Trim(Request.Form("c_dt_NF_remessa_termino"))
	ckb_pedido_nao_recebido_pelo_cliente = Trim(Request.Form("ckb_pedido_nao_recebido_pelo_cliente"))
	ckb_pedido_recebido_pelo_cliente = Trim(Request.Form("ckb_pedido_recebido_pelo_cliente"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_potencia_BTU = Trim(Request.Form("c_potencia_BTU"))
	c_ciclo = Trim(Request.Form("c_ciclo"))
	c_posicao_mercado = Trim(Request.Form("c_posicao_mercado"))
	c_grupo_pedido_origem = Trim(Request.Form("c_grupo_pedido_origem"))
	c_entrega_imediata = Trim(Request.Form("c_entrega_imediata"))
	rb_tipo_cliente = Trim(Request.Form("rb_tipo_cliente"))
	
	c_loja = Trim(Request.Form("c_loja"))
	lista_loja = substitui_caracteres(c_loja,chr(10),"")
	v_loja = split(lista_loja,chr(13),-1)

	ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET = Trim(Request.Form("ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET"))
    ckb_AGRUPAMENTO = Trim(Request.Form("ckb_AGRUPAMENTO"))
	ckb_CONSOLIDAR_PEDIDO = Trim(Request.Form("ckb_CONSOLIDAR_PEDIDO"))
	ckb_COMPATIBILIDADE = Trim(Request.Form("ckb_COMPATIBILIDADE"))

'	CAMPOS DE SA�DA SELECIONADOS
	dim ckb_COL_DT_CADASTRO, ckb_COL_DT_ENTREGA, ckb_COL_NF, ckb_COL_DT_EMISSAO_NF, ckb_COL_NF_REMESSA, ckb_COL_DT_EMISSAO_NF_REMESSA, ckb_COL_LOJA, ckb_COL_PEDIDO, ckb_COL_PEDIDO_MARKETPLACE, ckb_COL_GRUPO_PEDIDO_ORIGEM
	dim ckb_COL_VENDEDOR, ckb_COL_INDICADOR
	dim ckb_COL_CPF_CNPJ_CLIENTE, ckb_COL_CONTRIBUINTE_ICMS, ckb_COL_NOME_CLIENTE, ckb_COL_VL_RA, ckb_COL_RT, ckb_COL_ICMS_UF_DEST
	dim ckb_COL_PRODUTO, ckb_COL_NAC_IMP, ckb_COL_DESCRICAO_PRODUTO, ckb_COL_VL_NF, ckb_COL_VL_UNITARIO, ckb_COL_VL_CUSTO_REAL_TOTAL, ckb_COL_VL_TOTAL_NF, ckb_COL_VL_TOTAL, ckb_COL_QTDE
	dim ckb_COL_VL_CUSTO_ULT_ENTRADA, ckb_COL_VL_CUSTO_REAL, ckb_COL_VL_LISTA, ckb_COL_GRUPO, ckb_COL_POTENCIA_BTU
	dim ckb_COL_CICLO, ckb_COL_POSICAO_MERCADO, ckb_COL_MARCA, ckb_COL_TRANSPORTADORA, ckb_COL_ENTREGA_IMEDIATA, ckb_COL_DT_PREVISAO_ETG_TRANSP, ckb_COL_DT_RECEB_CLIENTE
	dim ckb_COL_CIDADE, ckb_COL_UF, ckb_COL_CIDADE_ETG, ckb_COL_UF_ETG
	dim ckb_COL_QTDE_PARCELAS, ckb_COL_MEIO_PAGAMENTO, ckb_COL_VL_PAGO_CARTAO_INTERNET, ckb_COL_CHAVE_NFE, ckb_COL_TEL, ckb_COL_EMAIL
    dim ckb_COL_PERC_DESC, ckb_COL_CUBAGEM, ckb_COL_PESO, ckb_COL_QTDE_VOLUMES, ckb_COL_FRETE, ckb_COL_FRETE_DETALHADO
    dim ckb_COL_INDICADOR_EMAILS, ckb_COL_INDICADOR_CPF_CNPJ, ckb_COL_INDICADOR_ENDERECO, ckb_COL_INDICADOR_CIDADE, ckb_COL_INDICADOR_UF
	
	ckb_COL_DT_CADASTRO = Trim(Request.Form("ckb_COL_DT_CADASTRO"))
	ckb_COL_DT_ENTREGA = Trim(Request.Form("ckb_COL_DT_ENTREGA"))
	ckb_COL_NF = Trim(Request.Form("ckb_COL_NF"))
	ckb_COL_DT_EMISSAO_NF = Trim(Request.Form("ckb_COL_DT_EMISSAO_NF"))
	ckb_COL_NF_REMESSA = Trim(Request.Form("ckb_COL_NF_REMESSA"))
	ckb_COL_DT_EMISSAO_NF_REMESSA = Trim(Request.Form("ckb_COL_DT_EMISSAO_NF_REMESSA"))
	ckb_COL_LOJA = Trim(Request.Form("ckb_COL_LOJA"))
	ckb_COL_PEDIDO = Trim(Request.Form("ckb_COL_PEDIDO"))
	ckb_COL_PEDIDO_MARKETPLACE = Trim(Request.Form("ckb_COL_PEDIDO_MARKETPLACE"))
	ckb_COL_GRUPO_PEDIDO_ORIGEM = Trim(Request.Form("ckb_COL_GRUPO_PEDIDO_ORIGEM"))
	ckb_COL_VENDEDOR = Trim(Request.Form("ckb_COL_VENDEDOR"))
	ckb_COL_INDICADOR = Trim(Request.Form("ckb_COL_INDICADOR"))
	ckb_COL_CPF_CNPJ_CLIENTE = Trim(Request.Form("ckb_COL_CPF_CNPJ_CLIENTE"))
	ckb_COL_CONTRIBUINTE_ICMS = Trim(Request.Form("ckb_COL_CONTRIBUINTE_ICMS"))
	ckb_COL_NOME_CLIENTE = Trim(Request.Form("ckb_COL_NOME_CLIENTE"))
	ckb_COL_VL_RA = Trim(Request.Form("ckb_COL_VL_RA"))
	ckb_COL_RT = Trim(Request.Form("ckb_COL_RT"))
	ckb_COL_ICMS_UF_DEST = Trim(Request.Form("ckb_COL_ICMS_UF_DEST"))
	ckb_COL_PRODUTO = Trim(Request.Form("ckb_COL_PRODUTO"))
	ckb_COL_NAC_IMP = Trim(Request.Form("ckb_COL_NAC_IMP"))
	ckb_COL_DESCRICAO_PRODUTO = Trim(Request.Form("ckb_COL_DESCRICAO_PRODUTO"))
	ckb_COL_VL_NF = Trim(Request.Form("ckb_COL_VL_NF"))
	ckb_COL_VL_UNITARIO = Trim(Request.Form("ckb_COL_VL_UNITARIO"))
	ckb_COL_VL_CUSTO_REAL_TOTAL = Trim(Request.Form("ckb_COL_VL_CUSTO_REAL_TOTAL"))
	ckb_COL_VL_TOTAL_NF = Trim(Request.Form("ckb_COL_VL_TOTAL_NF"))
	ckb_COL_VL_TOTAL = Trim(Request.Form("ckb_COL_VL_TOTAL"))
	ckb_COL_QTDE = Trim(Request.Form("ckb_COL_QTDE"))
	ckb_COL_VL_CUSTO_ULT_ENTRADA = Trim(Request.Form("ckb_COL_VL_CUSTO_ULT_ENTRADA"))
	ckb_COL_VL_CUSTO_REAL = Trim(Request.Form("ckb_COL_VL_CUSTO_REAL"))
	ckb_COL_VL_LISTA = Trim(Request.Form("ckb_COL_VL_LISTA"))
	ckb_COL_GRUPO = Trim(Request.Form("ckb_COL_GRUPO"))
	ckb_COL_POTENCIA_BTU = Trim(Request.Form("ckb_COL_POTENCIA_BTU"))
	ckb_COL_CICLO = Trim(Request.Form("ckb_COL_CICLO"))
	ckb_COL_POSICAO_MERCADO = Trim(Request.Form("ckb_COL_POSICAO_MERCADO"))
	ckb_COL_MARCA = Trim(Request.Form("ckb_COL_MARCA"))
	ckb_COL_TRANSPORTADORA = Trim(Request.Form("ckb_COL_TRANSPORTADORA"))
	ckb_COL_ENTREGA_IMEDIATA = Trim(Request.Form("ckb_COL_ENTREGA_IMEDIATA"))
	ckb_COL_DT_PREVISAO_ETG_TRANSP = Trim(Request.Form("ckb_COL_DT_PREVISAO_ETG_TRANSP"))
	ckb_COL_DT_RECEB_CLIENTE = Trim(Request.Form("ckb_COL_DT_RECEB_CLIENTE"))
	ckb_COL_CIDADE = Trim(Request.Form("ckb_COL_CIDADE"))
	ckb_COL_UF = Trim(Request.Form("ckb_COL_UF"))
	ckb_COL_CIDADE_ETG = Trim(Request.Form("ckb_COL_CIDADE_ETG"))
	ckb_COL_UF_ETG = Trim(Request.Form("ckb_COL_UF_ETG"))
	ckb_COL_QTDE_PARCELAS = Trim(Request.Form("ckb_COL_QTDE_PARCELAS"))
	ckb_COL_MEIO_PAGAMENTO = Trim(Request.Form("ckb_COL_MEIO_PAGAMENTO"))
	ckb_COL_VL_PAGO_CARTAO_INTERNET = Trim(Request.Form("ckb_COL_VL_PAGO_CARTAO_INTERNET"))
	ckb_COL_CHAVE_NFE = Trim(Request.Form("ckb_COL_CHAVE_NFE"))
	ckb_COL_TEL = Trim(Request.Form("ckb_COL_TEL"))
	ckb_COL_EMAIL = Trim(Request.Form("ckb_COL_EMAIL"))
	ckb_COL_PERC_DESC = Trim(Request.Form("ckb_COL_PERC_DESC"))
	ckb_COL_CUBAGEM = Trim(Request.Form("ckb_COL_CUBAGEM"))
	ckb_COL_PESO = Trim(Request.Form("ckb_COL_PESO"))
	ckb_COL_QTDE_VOLUMES = Trim(Request.Form("ckb_COL_QTDE_VOLUMES"))
	ckb_COL_FRETE = Trim(Request.Form("ckb_COL_FRETE"))
	ckb_COL_FRETE_DETALHADO = Trim(Request.Form("ckb_COL_FRETE_DETALHADO"))
	ckb_COL_INDICADOR_EMAILS = Trim(Request.Form("ckb_COL_INDICADOR_EMAILS"))
	ckb_COL_INDICADOR_ENDERECO = Trim(Request.Form("ckb_COL_INDICADOR_ENDERECO"))
	ckb_COL_INDICADOR_CIDADE = Trim(Request.Form("ckb_COL_INDICADOR_CIDADE"))
	ckb_COL_INDICADOR_UF = Trim(Request.Form("ckb_COL_INDICADOR_UF"))
	ckb_COL_INDICADOR_CPF_CNPJ = Trim(Request.Form("ckb_COL_INDICADOR_CPF_CNPJ"))
	
	dim s_campos_saida
	s_campos_saida = "|"
	if alerta = "" then
		if ckb_COL_DT_CADASTRO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_DT_CADASTRO" & "|"
		if ckb_COL_DT_ENTREGA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_DT_ENTREGA" & "|"
		if ckb_COL_NF <> "" then s_campos_saida = s_campos_saida & "ckb_COL_NF" & "|"
		if ckb_COL_DT_EMISSAO_NF <> "" then s_campos_saida = s_campos_saida & "ckb_COL_DT_EMISSAO_NF" & "|"
		if ckb_COL_NF_REMESSA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_NF_REMESSA" & "|"
		if ckb_COL_DT_EMISSAO_NF_REMESSA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_DT_EMISSAO_NF_REMESSA" & "|"
		if ckb_COL_LOJA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_LOJA" & "|"
		if ckb_COL_PEDIDO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_PEDIDO" & "|"
		if ckb_COL_PEDIDO_MARKETPLACE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_PEDIDO_MARKETPLACE" & "|"
		if ckb_COL_GRUPO_PEDIDO_ORIGEM <> "" then s_campos_saida = s_campos_saida & "ckb_COL_GRUPO_PEDIDO_ORIGEM" & "|"
		if ckb_COL_CPF_CNPJ_CLIENTE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_CPF_CNPJ_CLIENTE" & "|"
		if ckb_COL_CONTRIBUINTE_ICMS <> "" then s_campos_saida = s_campos_saida & "ckb_COL_CONTRIBUINTE_ICMS" & "|"
		if ckb_COL_NOME_CLIENTE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_NOME_CLIENTE" & "|"
		if ckb_COL_CIDADE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_CIDADE" & "|"
		if ckb_COL_UF <> "" then s_campos_saida = s_campos_saida & "ckb_COL_UF" & "|"
		if ckb_COL_CIDADE_ETG <> "" then s_campos_saida = s_campos_saida & "ckb_COL_CIDADE_ETG" & "|"
		if ckb_COL_UF_ETG <> "" then s_campos_saida = s_campos_saida & "ckb_COL_UF_ETG" & "|"
        if ckb_COL_TEL <> "" then s_campos_saida = s_campos_saida & "ckb_COL_TEL" & "|"
        if ckb_COL_EMAIL <> "" then s_campos_saida = s_campos_saida & "ckb_COL_EMAIL" & "|"
		if ckb_COL_VENDEDOR <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VENDEDOR" & "|"
		if ckb_COL_INDICADOR <> "" then s_campos_saida = s_campos_saida & "ckb_COL_INDICADOR" & "|"
		if ckb_COL_TRANSPORTADORA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_TRANSPORTADORA" & "|"
		if ckb_COL_ENTREGA_IMEDIATA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_ENTREGA_IMEDIATA" & "|"
		if ckb_COL_DT_PREVISAO_ETG_TRANSP <> "" then s_campos_saida = s_campos_saida & "ckb_COL_DT_PREVISAO_ETG_TRANSP" & "|"
		if ckb_COL_DT_RECEB_CLIENTE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_DT_RECEB_CLIENTE" & "|"
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
		if ckb_COL_NAC_IMP <> "" then s_campos_saida = s_campos_saida & "ckb_COL_NAC_IMP" & "|"
		if ckb_COL_DESCRICAO_PRODUTO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_DESCRICAO_PRODUTO" & "|"
		if ckb_COL_QTDE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_QTDE" & "|"
		if ckb_COL_PERC_DESC <> "" then s_campos_saida = s_campos_saida & "ckb_COL_PERC_DESC" & "|"
		if ckb_COL_CUBAGEM <> "" then s_campos_saida = s_campos_saida & "ckb_COL_CUBAGEM" & "|"
		if ckb_COL_PESO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_PESO" & "|"
		if ckb_COL_QTDE_VOLUMES <> "" then s_campos_saida = s_campos_saida & "ckb_COL_QTDE_VOLUMES" & "|"
		if ckb_COL_FRETE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_FRETE" & "|"
		if ckb_COL_FRETE_DETALHADO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_FRETE_DETALHADO" & "|"
		if ckb_COL_VL_CUSTO_ULT_ENTRADA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_CUSTO_ULT_ENTRADA" & "|"
		if ckb_COL_VL_CUSTO_REAL <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_CUSTO_REAL" & "|"
		if ckb_COL_VL_LISTA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_LISTA" & "|"
		if ckb_COL_VL_NF <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_NF" & "|"
		if ckb_COL_VL_UNITARIO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_UNITARIO" & "|"
		if ckb_COL_VL_CUSTO_REAL_TOTAL <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_CUSTO_REAL_TOTAL" & "|"
		if ckb_COL_VL_TOTAL_NF <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_TOTAL_NF" & "|"
		if ckb_COL_VL_TOTAL <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_TOTAL" & "|"
		if ckb_COL_VL_RA <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_RA" & "|"
		if ckb_COL_RT <> "" then s_campos_saida = s_campos_saida & "ckb_COL_RT" & "|"
		if ckb_COL_ICMS_UF_DEST <> "" then s_campos_saida = s_campos_saida & "ckb_COL_ICMS_UF_DEST" & "|"
		if ckb_COL_QTDE_PARCELAS <> "" then s_campos_saida = s_campos_saida & "ckb_COL_QTDE_PARCELAS" & "|"
		if ckb_COL_MEIO_PAGAMENTO <> "" then s_campos_saida = s_campos_saida & "ckb_COL_MEIO_PAGAMENTO" & "|"
		if ckb_COL_VL_PAGO_CARTAO_INTERNET <> "" then s_campos_saida = s_campos_saida & "ckb_COL_VL_PAGO_CARTAO_INTERNET" & "|"
		if ckb_COL_CHAVE_NFE <> "" then s_campos_saida = s_campos_saida & "ckb_COL_CHAVE_NFE" & "|"
		
		if s_campos_saida = "|" then s_campos_saida = "NENHUM"
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|campos_saida_selecionados", s_campos_saida)
		
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_faturamento_inicio", c_dt_faturamento_inicio)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_faturamento_termino", c_dt_faturamento_termino)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_NF_venda_inicio", c_dt_NF_venda_inicio)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_NF_venda_termino", c_dt_NF_venda_termino)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_NF_remessa_inicio", c_dt_NF_remessa_inicio)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_NF_remessa_termino", c_dt_NF_remessa_termino)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|ckb_pedido_nao_recebido_pelo_cliente", ckb_pedido_nao_recebido_pelo_cliente)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|ckb_pedido_recebido_pelo_cliente", ckb_pedido_recebido_pelo_cliente)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_fabricante", c_fabricante)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_grupo", c_grupo)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_potencia_BTU", c_potencia_BTU)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_ciclo", c_ciclo)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_posicao_mercado", c_posicao_mercado)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_entrega_imediata", c_entrega_imediata)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|rb_tipo_cliente", rb_tipo_cliente)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_loja", c_loja)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET", ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|ckb_CONSOLIDAR_PEDIDO", ckb_CONSOLIDAR_PEDIDO)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|ckb_COMPATIBILIDADE", ckb_COMPATIBILIDADE)
		end if
	
	if alerta = "" then
		if ckb_CONSOLIDAR_PEDIDO <> "" then
			'SE A CONSOLIDA��O POR PEDIDO ESTIVER ATIVA, ASSEGURA QUE AS OP��ES INCOMPAT�VEIS ESTEJAM DESATIVADAS
			ckb_AGRUPAMENTO = ""
			ckb_COL_MARCA = ""
			ckb_COL_GRUPO = ""
			ckb_COL_POTENCIA_BTU = ""
			ckb_COL_CICLO = ""
			ckb_COL_POSICAO_MERCADO = ""
			ckb_COL_PRODUTO = ""
			ckb_COL_NAC_IMP = ""
			ckb_COL_DESCRICAO_PRODUTO = ""
			ckb_COL_QTDE = ""
			ckb_COL_PERC_DESC = ""
			ckb_COL_VL_CUSTO_ULT_ENTRADA = ""
			ckb_COL_VL_CUSTO_REAL = ""
			ckb_COL_VL_LISTA = ""
			ckb_COL_VL_NF = ""
			ckb_COL_VL_UNITARIO = ""
		else
			'CAMPOS DISPON�VEIS SOMENTE QUANDO A OP��O DE CONSOLIDAR POR PEDIDO ESTIVER ATIVA
			ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET = ""
			ckb_COL_VL_PAGO_CARTAO_INTERNET = ""
			end if
		end if

	if alerta = "" then
		if c_entrega_imediata <> "" then
			if (c_dt_faturamento_inicio <> "") Or (c_dt_faturamento_termino <> "") then
				if msg_erro_consistencia_parametros <> "" then msg_erro_consistencia_parametros = msg_erro_consistencia_parametros & vbCrLf
				msg_erro_consistencia_parametros = msg_erro_consistencia_parametros & "O filtro de status de Entrega Imediata n�o pode operar em conjunto com o filtro 'PER�ODO (ENTREGUE)'!"
				end if
			if ckb_COL_VL_CUSTO_REAL <> "" then
				if msg_erro_consistencia_parametros <> "" then msg_erro_consistencia_parametros = msg_erro_consistencia_parametros & vbCrLf
				msg_erro_consistencia_parametros = msg_erro_consistencia_parametros & "O filtro de status de Entrega Imediata n�o pode operar em conjunto com o campo de sa�da 'VL Custo (Real)'!"
				end if
			if ckb_COL_VL_CUSTO_REAL_TOTAL <> "" then
				if msg_erro_consistencia_parametros <> "" then msg_erro_consistencia_parametros = msg_erro_consistencia_parametros & vbCrLf
				msg_erro_consistencia_parametros = msg_erro_consistencia_parametros & "O filtro de status de Entrega Imediata n�o pode operar em conjunto com o campo de sa�da 'VL Custo Total (Real)'!"
				end if
			if ckb_COL_NAC_IMP <> "" then
				if msg_erro_consistencia_parametros <> "" then msg_erro_consistencia_parametros = msg_erro_consistencia_parametros & vbCrLf
				msg_erro_consistencia_parametros = msg_erro_consistencia_parametros & "O filtro de status de Entrega Imediata n�o pode operar em conjunto com o campo de sa�da 'Nacional/Importado'!"
				end if
			end if
		end if

	if alerta = "" then
		Response.ContentType = "application/csv"
		Response.AddHeader "Content-Disposition", "attachment; filename=TabDinamica_" & formata_data_yyyymmdd(Now) & "_" & substitui_caracteres(formata_hora(Now),":","") & ".csv"
		if msg_erro_consistencia_parametros = "" then
			consulta_executa
		else
			Response.Write msg_erro_consistencia_parametros
			end if
		Response.End
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  �  �  E  S 
' _____________________________________________________________________________________________

' ------------------------------------------------------------------------
'   descricao_icms_contribuinte_x_produtor_rural
'
function descricao_icms_contribuinte_x_produtor_rural(byval tipo_pessoa, byval contribuinte_icms_status, byval produtor_rural_status)
dim strResp

    tipo_pessoa = Trim(tipo_pessoa)
    contribuinte_icms_status = Trim(contribuinte_icms_status)
    produtor_rural_status = Trim(produtor_rural_status)

    if tipo_pessoa = ID_PJ then
        select case contribuinte_icms_status
            case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO
                strResp = "N�o Contribuinte"
            case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM
                strResp = "Contribuinte"
            case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO
                strResp = "Isento"
            case else
                strResp = ""
        end select
    elseif tipo_pessoa = ID_PF then
        select case produtor_rural_status
            case COD_ST_CLIENTE_PRODUTOR_RURAL_NAO
                strResp = ""
            case COD_ST_CLIENTE_PRODUTOR_RURAL_SIM
                strResp = "Produtor Rural"
            case else
                strResp = ""
        end select
    end if

    descricao_icms_contribuinte_x_produtor_rural = strResp
end function



' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const SEPARADOR_DECIMAL = ","
dim s, s_sql, s_sql_aux, s_cst, x, x_cab, s_where, s_where_aux, s_where_temp, s_where_venda, s_where_devolucao, s_where_loja, s_where_lista_codigo_frete_devolucao, s_where_periodo_NF
dim perc_RT, vl_RA, vl_RT, vl_preco_venda, vl_preco_NF, n_reg, n_reg_total, n_reg_total_passo1
dim tipo_parc
dim s_qtde, item_peso, item_qtde_volumes, item_cubagem, item_qtde, s_campo_qtde
dim s_vICMSUFDest, vl_vICMSUFDest, s_vICMSUFDest_unitario, vl_vICMSUFDest_unitario, s_det__qCom, n_det__qCom, vl_frete_proporcional
dim v
dim vNFeAConsultar, vNFeChave, iQI
dim strNfeT1ServidorBd, strNfeT1NomeBd, strNfeT1UsuarioBd, strNfeT1SenhaCriptografadaBd, senha_decodificada, chave, s_pesq_nf, idxNFeLocalizada
dim vTipoFrete

	n_reg_total_passo1 = -1

	if ckb_COL_CHAVE_NFE <> "" then
		redim vNFeChave(0)
		set vNFeChave(UBound(vNFeChave)) = New cl_DEZ_COLUNAS
		vNFeChave(UBound(vNFeChave)).CampoOrdenacao = ""
		vNFeChave(UBound(vNFeChave)).c1 = ""
		vNFeChave(UBound(vNFeChave)).c2 = ""
		vNFeChave(UBound(vNFeChave)).c3 = ""

		redim vNFeAConsultar(0)
		set vNFeAConsultar(Ubound(vNFeAConsultar)) = New cl_TRES_COLUNAS
		vNFeAConsultar(UBound(vNFeAConsultar)).c1 = ""
		set vNFeAConsultar(UBound(vNFeAConsultar)).c2 = nothing
		vNFeAConsultar(UBound(vNFeAConsultar)).c3 = ""

		'LOCALIZA TODOS OS EMITENTES DE NFE
		s = "SELECT * FROM t_NFe_EMITENTE ORDER BY id"
		if r.State <> 0 then r.Close
		r.open s, cn
		do while Not r.Eof
			if vNFeAConsultar(UBound(vNFeAConsultar)).c1 <> "" then
				redim preserve vNFeAConsultar(Ubound(vNFeAConsultar)+1)
				set vNFeAConsultar(Ubound(vNFeAConsultar)) = New cl_TRES_COLUNAS
				vNFeAConsultar(UBound(vNFeAConsultar)).c1 = ""
				vNFeAConsultar(UBound(vNFeAConsultar)).c3 = ""
				end if
			vNFeAConsultar(UBound(vNFeAConsultar)).c1 = Trim("" & r("id"))

			strNfeT1ServidorBd = Trim("" & r("NFe_T1_servidor_BD"))
			strNfeT1NomeBd = Trim("" & r("NFe_T1_nome_BD"))
			strNfeT1UsuarioBd = Trim("" & r("NFe_T1_usuario_BD"))
			strNfeT1SenhaCriptografadaBd = Trim("" & r("NFe_T1_senha_BD"))

			chave = gera_chave(FATOR_BD)
			decodifica_dado strNfeT1SenhaCriptografadaBd, senha_decodificada, chave

			s = "Provider=SQLOLEDB;" & _
				"Data Source=" & strNfeT1ServidorBd & ";" & _
				"Initial Catalog=" & strNfeT1NomeBd & ";" & _
				"User ID=" & strNfeT1UsuarioBd & ";" & _
				"Password=" & senha_decodificada & ";"
			set vNFeAConsultar(UBound(vNFeAConsultar)).c2 = server.CreateObject("ADODB.Connection")
			vNFeAConsultar(UBound(vNFeAConsultar)).c2.ConnectionTimeout = 45
			vNFeAConsultar(UBound(vNFeAConsultar)).c2.CommandTimeout = 900
			vNFeAConsultar(UBound(vNFeAConsultar)).c2.ConnectionString = s
			On Error Resume Next
			Err.Clear
			vNFeAConsultar(UBound(vNFeAConsultar)).c2.Open
			if Err <> 0 then set vNFeAConsultar(UBound(vNFeAConsultar)).c2 = nothing
			On Error GoTo 0
			Err.Clear

			r.MoveNext
			loop
		end if 'if ckb_COL_CHAVE_NFE <> ""

'	OBT�M O C�DIGO REFERENTE AO FRETE DE DEVOLU��O
	s_where_lista_codigo_frete_devolucao = ""
	s = "SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo = '" & GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE & "') AND (parametro_campo_texto = 'DEV')"
	if r.State <> 0 then r.Close
	r.open s, cn
	do while Not r.Eof
		if s_where_lista_codigo_frete_devolucao <> "" then s_where_lista_codigo_frete_devolucao = s_where_lista_codigo_frete_devolucao & ","
		s_where_lista_codigo_frete_devolucao = s_where_lista_codigo_frete_devolucao & "'" & Trim("" & r("codigo")) & "'"
		r.MoveNext
		loop

	if s_where_lista_codigo_frete_devolucao <> "" then s_where_lista_codigo_frete_devolucao = " (" & s_where_lista_codigo_frete_devolucao & ")"

	if (ckb_COL_FRETE <> "") And (ckb_COL_FRETE_DETALHADO <> "") then
	'	OBT�M AS INFORMA��ES SOBRE CADA TIPO DE FRETE
		redim vTipoFrete(0)
		set vTipoFrete(UBound(vTipoFrete)) = new cl_TIPO_FRETE
		vTipoFrete(UBound(vTipoFrete)).codigo = ""

	'	ORDENA POR parametro_campo_texto PARA QUE O FRETE DE DEVOLU��O FIQUE NO FINAL
		s = "SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo = '" & GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE & "') ORDER BY parametro_campo_texto, ordenacao"
		if r.State <> 0 then r.Close
		r.open s, cn
		do while Not r.Eof
			if vTipoFrete(UBound(vTipoFrete)).codigo <> "" then
				redim preserve vTipoFrete(UBound(vTipoFrete)+1)
				set vTipoFrete(UBound(vTipoFrete)) = new cl_TIPO_FRETE
				vTipoFrete(UBound(vTipoFrete)).codigo = ""
				end if
			with vTipoFrete(UBound(vTipoFrete))
				.codigo = Trim("" & r("codigo"))
				.descricao = Trim("" & r("descricao"))
				.colSQL = "vlFrete_" & Trim("" & r("codigo"))
				if Trim("" & r("parametro_campo_texto")) = "DEV" then
					.blnFreteDevolucao = True
				else
					.blnFreteDevolucao = False
					end if
				end with

			r.MoveNext
			loop
		end if


'	CRIT�RIOS COMUNS
'	================
	s_where = ""
	s_where_venda = ""
	s_where_devolucao = ""
	s_where_periodo_NF = ""

	item_peso = 0
	item_qtde_volumes = 0
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

	'PER�ODO NF
	'DEVIDO � FORMA COMO A DATA DE EMISS�O DA NF � OBTIDA, N�O � POSS�VEL APLICAR A RESTRI��O POR DATA NA CONSULTA BASE
	'PORTANTO, PARA MINIMIZAR A QUANTIDADE DE REGISTROS SELECIONADOS, FORAM ADOTADOS OS SEGUINTES CRIT�RIOS:
	'	1) AS RESTRI��ES S�O APLICADAS EM 2 MOMENTOS DISTINTOS: NA CONSULTA BASE (INTERNA) E NA CONSULTA GERAL (EXTERNA)
	'	2) CONSULTA BASE:
	'		2a) RESTRINGE POR PEDIDOS QUE TENHAM N�MERO DE NF
	'		2b) EXCLUI OS PEDIDOS CANCELADOS
	'	3) CONSULTA GERAL
	'		3a) AS DATAS DE IN�CIO E FIM DO PER�ODO S�O APLICADAS SOBRE A DATA DE EMISS�O RETORNADA PELA CONSULTA BASE
	if IsDate(c_dt_NF_venda_inicio) Or IsDate(c_dt_NF_venda_termino) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.num_obs_2 > 0) AND (t_PEDIDO.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')"
			
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.NFe_numero_NF > 0)"
		end if

	if IsDate(c_dt_NF_venda_inicio) then
		if s_where_periodo_NF <> "" then s_where_periodo_NF = s_where_periodo_NF & " AND"
		s_where_periodo_NF = s_where_periodo_NF & " (dt_emissao >= " & bd_formata_data(StrToDate(c_dt_NF_venda_inicio)) & ")"
		end if
		
	if IsDate(c_dt_NF_venda_termino) then
		if s_where_periodo_NF <> "" then s_where_periodo_NF = s_where_periodo_NF & " AND"
		s_where_periodo_NF = s_where_periodo_NF & " (dt_emissao < " & bd_formata_data(StrToDate(c_dt_NF_venda_termino)+1) & ")"
		end if

	'PER�ODO NF REMESSA
	if IsDate(c_dt_NF_remessa_inicio) Or IsDate(c_dt_NF_remessa_termino) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.num_obs_3 > 0) AND (t_PEDIDO.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')"
		
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.NFe_numero_NF > 0)"
		end if

	if IsDate(c_dt_NF_remessa_inicio) then
		if s_where_periodo_NF <> "" then s_where_periodo_NF = s_where_periodo_NF & " AND"
		s_where_periodo_NF = s_where_periodo_NF & " (dt_emissao_remessa >= " & bd_formata_data(StrToDate(c_dt_NF_remessa_inicio)) & ")"
		end if
		
	if IsDate(c_dt_NF_remessa_termino) then
		if s_where_periodo_NF <> "" then s_where_periodo_NF = s_where_periodo_NF & " AND"
		s_where_periodo_NF = s_where_periodo_NF & " (dt_emissao_remessa < " & bd_formata_data(StrToDate(c_dt_NF_remessa_termino)+1) & ")"
		end if

	'CRIT�RIO: PEDIDOS RECEBIDOS PELO CLIENTE
	s_where_aux = ""
	if ckb_pedido_nao_recebido_pelo_cliente <> "" then
		if s_where_aux <> "" then s_where_aux = s_where_aux & " OR"
		s_where_aux = s_where_aux & " (t_PEDIDO.PedidoRecebidoStatus = 0)"
		end if

	if ckb_pedido_recebido_pelo_cliente <> "" then
		if s_where_aux <> "" then s_where_aux = s_where_aux & " OR"
		s_where_aux = s_where_aux & " (t_PEDIDO.PedidoRecebidoStatus = 1)"
		end if

	if s_where_aux <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s_where_aux & ")"
		end if

	'CRIT�RIO: FABRICANTE
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
	
	if c_entrega_imediata <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.st_etg_imediata = " & c_entrega_imediata & ") AND (t_PEDIDO.st_entrega NOT IN ('" & ST_ENTREGA_CANCELADO & "', '" & ST_ENTREGA_ENTREGUE & "'))"
		end if

    s_where_temp = ""
    if c_grupo_pedido_origem <> "" then
        v_grupo_pedido_origem = split(c_grupo_pedido_origem, ", ")
        for i = LBound(v_grupo_pedido_origem) to UBound(v_grupo_pedido_origem)
            s = "SELECT codigo FROM t_CODIGO_DESCRICAO WHERE (codigo_pai = '" & v_grupo_pedido_origem(i) & "') AND grupo='PedidoECommerce_Origem'"
            if rs.State <> 0 then rs.Close
	        rs.open s, cn
		    if rs.Eof then
                alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_grupo_pedido_origem & " N�O EXISTE."
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
				" 'VENDA_NORMAL' AS operacao" & _
				", t_PEDIDO.data_hora"

	if ckb_COL_ICMS_UF_DEST <> "" then
		s_sql = s_sql & _
				", Convert(DATETIME, t_NFe_IMAGEM_NORMALIZADO.ide__dEmi, 121) AS dt_emissao" & _
				", (SELECT TOP 1 Convert(datetime, ide__dEmi, 121) FROM t_NFe_IMAGEM WHERE (t_NFe_IMAGEM.NFe_numero_NF = t_PEDIDO.num_obs_3) AND (t_NFe_IMAGEM.id_nfe_emitente = t_PEDIDO.id_nfe_emitente) AND (t_NFe_IMAGEM.ide__tpNF = '1') AND (t_NFe_IMAGEM.st_anulado = 0) AND (t_NFe_IMAGEM.codigo_retorno_NFe_T1 = 1) ORDER BY id DESC) AS dt_emissao_remessa"
	else
		s_sql = s_sql & _
				", (SELECT TOP 1 Convert(datetime, ide__dEmi, 121) FROM t_NFe_IMAGEM WHERE (t_NFe_IMAGEM.NFe_numero_NF = t_PEDIDO.num_obs_2) AND (t_NFe_IMAGEM.id_nfe_emitente = t_PEDIDO.id_nfe_emitente) AND (t_NFe_IMAGEM.ide__tpNF = '1') AND (t_NFe_IMAGEM.st_anulado = 0) AND (t_NFe_IMAGEM.codigo_retorno_NFe_T1 = 1) ORDER BY id DESC) AS dt_emissao" & _
				", (SELECT TOP 1 Convert(datetime, ide__dEmi, 121) FROM t_NFe_IMAGEM WHERE (t_NFe_IMAGEM.NFe_numero_NF = t_PEDIDO.num_obs_3) AND (t_NFe_IMAGEM.id_nfe_emitente = t_PEDIDO.id_nfe_emitente) AND (t_NFe_IMAGEM.ide__tpNF = '1') AND (t_NFe_IMAGEM.st_anulado = 0) AND (t_NFe_IMAGEM.codigo_retorno_NFe_T1 = 1) ORDER BY id DESC) AS dt_emissao_remessa"
		end if

	s_sql = s_sql & _
				", t_PEDIDO.id_nfe_emitente" & _
				", t_PEDIDO.num_obs_2 AS numero_NF" & _
				", t_PEDIDO.obs_2" & _
				", t_PEDIDO.num_obs_3 AS numero_NF_remessa" & _
				", t_PEDIDO.obs_3" & _
				", t_PEDIDO.loja" & _
				", t_PEDIDO.pedido" & _
				", t_PEDIDO.pedido_bs_x_marketplace" & _
				", t_PEDIDO.marketplace_codigo_origem" & _
				", t_PEDIDO.st_entrega" & _
				", t_PEDIDO.entregue_data" & _
				", t_PEDIDO.st_etg_imediata" & _
				", t_PEDIDO.PrevisaoEntregaData" & _
				", t_PEDIDO.PrevisaoEntregaTranspData" & _
				", t_PEDIDO.PedidoRecebidoStatus" & _
				", t_PEDIDO.PedidoRecebidoData" & _
				", tGrupoPedidoOrigemDescricao.descricao AS GrupoPedidoOrigemDescricao" & _
				", tPedidoOrigemDescricao.descricao AS PedidoOrigemDescricao" & _
				", t_PEDIDO.transportadora_id" & _
				", t_PEDIDO__BASE.vendedor" & _
				", t_PEDIDO__BASE.indicador" & _
				", t_PEDIDO__BASE.perc_RT"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				", t_PEDIDO.endereco_nome AS nome_cliente" & _
                ", t_PEDIDO.endereco_tipo_pessoa AS tipo_cliente" & _
				", t_PEDIDO.endereco_cnpj_cpf AS cnpj_cpf" & _
				", t_PEDIDO.endereco_contribuinte_icms_status AS contribuinte_icms_status" & _
				", t_PEDIDO.endereco_produtor_rural_status AS produtor_rural_status" & _
				", t_PEDIDO.endereco_cidade AS cidade" & _
				", t_PEDIDO.endereco_uf AS uf" & _
                ", t_PEDIDO.endereco_ddd_res AS ddd_res" & _
                ", t_PEDIDO.endereco_tel_res AS tel_res" & _
                ", t_PEDIDO.endereco_ddd_cel AS ddd_cel" & _
                ", t_PEDIDO.endereco_tel_cel AS tel_cel" & _
                ", t_PEDIDO.endereco_ddd_com AS ddd_com" & _
                ", t_PEDIDO.endereco_tel_com AS tel_com" & _
                ", t_PEDIDO.endereco_ddd_com_2 AS ddd_com_2" & _
                ", t_PEDIDO.endereco_tel_com_2 AS tel_com_2" & _
                ", t_PEDIDO.endereco_ramal_com AS ramal_com" & _
                ", t_PEDIDO.endereco_ramal_com_2 AS ramal_com_2" & _
                ", t_PEDIDO.endereco_email AS email"
	else
		s_sql = s_sql & _
				", t_CLIENTE.nome_iniciais_em_maiusculas AS nome_cliente" & _
                ", t_CLIENTE.tipo AS tipo_cliente" & _
				", t_CLIENTE.cnpj_cpf" & _
				", t_CLIENTE.contribuinte_icms_status" & _
				", t_CLIENTE.produtor_rural_status" & _
				", t_CLIENTE.cidade AS cidade" & _
				", t_CLIENTE.uf AS uf" & _
                ", t_CLIENTE.ddd_res" & _
                ", t_CLIENTE.tel_res" & _
                ", t_CLIENTE.ddd_cel" & _
                ", t_CLIENTE.tel_cel" & _
                ", t_CLIENTE.ddd_com" & _
                ", t_CLIENTE.tel_com" & _
                ", t_CLIENTE.ddd_com_2" & _
                ", t_CLIENTE.tel_com_2" & _
                ", t_CLIENTE.ramal_com" & _
                ", t_CLIENTE.ramal_com_2" & _
                ", t_CLIENTE.email"
		end if

	s_sql = s_sql & _
				", t_PEDIDO.st_end_entrega" & _
				", t_PEDIDO.EndEtg_cidade" & _
				", t_PEDIDO.EndEtg_uf" & _
				", t_ORCAMENTISTA_E_INDICADOR.cnpj_cpf AS indicador_cnpj_cpf" & _
				", t_ORCAMENTISTA_E_INDICADOR.endereco AS indicador_endereco" & _
				", t_ORCAMENTISTA_E_INDICADOR.endereco_numero AS indicador_endereco_numero" & _
				", t_ORCAMENTISTA_E_INDICADOR.endereco_complemento AS indicador_endereco_complemento" & _
				", t_ORCAMENTISTA_E_INDICADOR.bairro AS indicador_bairro" & _
				", t_ORCAMENTISTA_E_INDICADOR.cidade AS indicador_cidade" & _
				", t_ORCAMENTISTA_E_INDICADOR.uf AS indicador_uf" & _
				", t_ORCAMENTISTA_E_INDICADOR.cep AS indicador_cep" & _
				", t_ORCAMENTISTA_E_INDICADOR.email AS indicador_email" & _
				", t_ORCAMENTISTA_E_INDICADOR.email2 AS indicador_email2" & _
				", t_ORCAMENTISTA_E_INDICADOR.email3 AS indicador_email3" & _
				", t_PEDIDO__BASE.qtde_parcelas AS qtde_parcelas" & _
				", t_PEDIDO__BASE.tipo_parcelamento AS tipo_parcelamento" & _
				", t_PEDIDO__BASE.av_forma_pagto AS forma_pagamento_av" & _
				", t_PEDIDO__BASE.pce_forma_pagto_prestacao AS parcelamento_c_entrada" & _
				", t_PEDIDO__BASE.pse_forma_pagto_demais_prest AS parcelamento_s_entrada" & _
				", t_PEDIDO__BASE.pu_forma_pagto AS parcela_unica"
	
	s_where_aux = ""
	if s_where_lista_codigo_frete_devolucao <> "" then
	'	EXCLUI OS FRETES DE DEVOLU��O
		s_where_aux = " AND (codigo_tipo_frete NOT IN " & s_where_lista_codigo_frete_devolucao & ")"
		end if

	if ckb_COL_FRETE <> "" then
		if ckb_COL_FRETE_DETALHADO = "" then
			s_sql = s_sql & _
						", (SELECT Coalesce(SUM(vl_frete),0) AS vl_frete FROM t_PEDIDO_FRETE WHERE (t_PEDIDO_FRETE.pedido=t_PEDIDO.pedido)" & s_where_aux & ") AS vl_frete"
		else
			for i=LBound(vTipoFrete) to UBound(vTipoFrete)
				with vTipoFrete(i)
					if .codigo <> "" then
						'QUANDO A LISTAGEM � ITEM A ITEM, O VALOR DO FRETE DE DEVOLU��O SOMENTE � EXIBIDO NOS ITENS DEVOLVIDOS E VICE-VERSA
						if (ckb_CONSOLIDAR_PEDIDO = "") And .blnFreteDevolucao then
							s_sql = s_sql & _
										", 0 AS " & .colSQL
						else
							s_sql = s_sql & _
										", (SELECT Coalesce(SUM(vl_frete),0) AS vl_frete FROM t_PEDIDO_FRETE WHERE (t_PEDIDO_FRETE.pedido=t_PEDIDO.pedido) AND (codigo_tipo_frete = '" & .codigo & "')) AS " & .colSQL
							end if
						end if
					end with
				next
			end if

		s_sql = s_sql & _
					", (SELECT Coalesce(SUM(qtde * preco_venda),0) AS vl_total_produtos_calc_frete FROM t_PEDIDO_ITEM WHERE (t_PEDIDO_ITEM.pedido = t_PEDIDO.pedido)) AS vl_total_produtos_calc_frete"
		end if

	if ckb_CONSOLIDAR_PEDIDO = "" then
		'N�O CONSOLIDA POR PEDIDO, OU SEJA, O RESULTADO SER� POR ITEM DE PEDIDO
		if (ckb_COL_VL_CUSTO_REAL <> "") Or (ckb_COL_VL_CUSTO_REAL_TOTAL <> "") then
			s_sql = s_sql & _
					", t_ESTOQUE_MOVIMENTO.qtde"
		else
			s_sql = s_sql & _
					", t_PEDIDO_ITEM.qtde"
			end if

		if ckb_COL_NAC_IMP <> "" then
			s_sql = s_sql & _
					", t_ESTOQUE_ITEM.cst"
			end if

		s_sql = s_sql & _
					", t_PEDIDO_ITEM.fabricante" & _
					", t_PEDIDO_ITEM.produto" & _
					", t_PEDIDO_ITEM.descricao" & _
					", t_PEDIDO_ITEM.preco_venda" & _
					", t_PEDIDO_ITEM.preco_lista" & _
					", t_PEDIDO_ITEM.preco_NF" & _
					", t_PEDIDO_ITEM.desc_dado" & _
					", t_PEDIDO_ITEM.cubagem" & _
					", t_PEDIDO_ITEM.peso" & _
					", t_PEDIDO_ITEM.qtde_volumes" & _
					", t_PRODUTO.grupo" & _
					", t_PRODUTO.potencia_BTU" & _
					", t_PRODUTO.ciclo" & _
					", t_PRODUTO.posicao_mercado" & _
					", t_FABRICANTE.nome AS nome_fabricante"

		s_sql = s_sql & _
					", (SELECT TOP 1 vl_custo2 FROM t_ESTOQUE tE INNER JOIN t_ESTOQUE_ITEM tEI ON (tE.id_estoque = tEI.id_estoque) WHERE (tE.devolucao_status = 0) AND (tEI.fabricante = t_PEDIDO_ITEM.fabricante) AND (tEI.produto = t_PEDIDO_ITEM.produto) AND (tEI.qtde > 0) ORDER BY tEI.id_estoque DESC) AS vl_custo2_ult_entrada"

		if (ckb_COL_VL_CUSTO_REAL <> "") Or (ckb_COL_VL_CUSTO_REAL_TOTAL <> "") then
			s_sql = s_sql & _
					", (t_ESTOQUE_ITEM.vl_custo2) AS vl_custo2_real"
			end if

		if ckb_COL_ICMS_UF_DEST <> "" then
			s_sql = s_sql & _
					", t_NFe_IMAGEM_ITEM_NORMALIZADO.ICMSUFDest__vICMSUFDest AS vICMSUFDest" & _
					", t_NFe_IMAGEM_ITEM_NORMALIZADO.det__qCom AS det__qCom"
			end if

	else
		'CONSOLIDA��O POR PEDIDO
		if (ckb_COL_VL_CUSTO_REAL <> "") Or (ckb_COL_VL_CUSTO_REAL_TOTAL <> "") then
			s_campo_qtde = "t_ESTOQUE_MOVIMENTO.qtde"
		else
			s_campo_qtde = "t_PEDIDO_ITEM.qtde"
			end if

		s_sql = s_sql & _
					", (" & s_campo_qtde & " * t_PEDIDO_ITEM.cubagem) AS cubagem_total_item" & _
					", (" & s_campo_qtde & " * t_PEDIDO_ITEM.peso) AS peso_total_item" & _
					", (" & s_campo_qtde & " * t_PEDIDO_ITEM.qtde_volumes) AS qtde_volumes_total_item" & _
					", (" & s_campo_qtde & " * t_PEDIDO_ITEM.preco_NF) AS preco_NF_total_item" & _
					", (" & s_campo_qtde & " * t_PEDIDO_ITEM.preco_venda) AS preco_venda_total_item"

		if (ckb_COL_VL_CUSTO_REAL <> "") Or (ckb_COL_VL_CUSTO_REAL_TOTAL <> "") then
			s_sql = s_sql & _
					", (" & s_campo_qtde & " * t_ESTOQUE_ITEM.vl_custo2) AS vl_custo2_real_total_item"
			end if

		if ckb_COL_ICMS_UF_DEST <> "" then
			s_sql = s_sql & _
					", t_NFe_IMAGEM_ITEM_NORMALIZADO.ICMSUFDest__vICMSUFDest AS vICMSUFDest"
			end if

		if ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET <> "" then
			s_sql = s_sql & _
					", (SELECT SUM(valor_transacao) AS vl_pago_cartao_internet_total_pedido FROM t_PAGTO_GW_PAG tPAG INNER JOIN t_PAGTO_GW_PAG_PAYMENT tPAY ON (tPAG.id=tPAY.id_pagto_gw_pag) INNER JOIN t_PEDIDO tPedPagtoCartao ON (tPAG.pedido=tPedPagtoCartao.pedido) WHERE (tPedPagtoCartao.pedido_base=t_PEDIDO.pedido_base) AND (tPAY.ult_GlobalStatus='" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "')) AS vl_pago_cartao_internet_total_familia_pedido"
			end if

		if ckb_COL_VL_PAGO_CARTAO_INTERNET <> "" then
			s_sql = s_sql & _
					", (SELECT SUM(valor_transacao) AS vl_pago_cartao_internet_total_pedido FROM t_PAGTO_GW_PAG tPAG INNER JOIN t_PAGTO_GW_PAG_PAYMENT tPAY ON (tPAG.id=tPAY.id_pagto_gw_pag) WHERE (tPAG.pedido=t_PEDIDO.pedido) AND (tPAY.ult_GlobalStatus='" & BRASPAG_PAGADOR_CARTAO_GLOBAL_STATUS__CAPTURADA & "')) AS vl_pago_cartao_internet_total_pedido"
			end if
		end if 'if ckb_CONSOLIDAR_PEDIDO = ""


	s_sql = s_sql & _
			" FROM t_PEDIDO" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base = t_PEDIDO__BASE.pedido)" & _
                " LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador = t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
				" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido = t_PEDIDO_ITEM.pedido)" & _
				" INNER JOIN t_PRODUTO ON (t_PRODUTO.fabricante = t_PEDIDO_ITEM.fabricante) AND (t_PRODUTO.produto = t_PEDIDO_ITEM.produto)" & _
				" INNER JOIN t_FABRICANTE ON (t_PRODUTO.fabricante = t_FABRICANTE.fabricante)" & _
				" INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente = t_CLIENTE.id)"

	if (ckb_COL_VL_CUSTO_REAL <> "") Or (ckb_COL_VL_CUSTO_REAL_TOTAL <> "") Or (ckb_COL_NAC_IMP <> "") then
		s_sql = s_sql & _
				" INNER JOIN t_ESTOQUE_MOVIMENTO ON ((t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto))" & _
				" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto)AND(t_ESTOQUE_ITEM.qtde > 0))"
		end if

	' Monta derived table para acessar os dados de NFe
	' Inicialmente, esses dados estavam sendo obtidos atrav�s de selects introduzidos diretamente no select principal, ou seja, uma consulta interna p/ cada campo referente a esses dados da NFe,
	' mas isso se mostrou muito ineficiente em termos de performance.
	' Na tabela t_NFe_IMAGEM pode haver mais de um registro com o mesmo n�mero de nota do mesmo emitente, isso pode ocorrer devido � reutiliza��o do n�mero ap�s uma emiss�o que tenha sido rejeitada pela Sefaz.
	' Na tabela t_NFe_IMAGEM_ITEM tamb�m pode haver mais de um registro para o mesmo c�digo de produto, isso pode ocorrer quando o pedido consumiu produtos de estoques diferentes e caso esses produtos tenham
	' alguns dados diferentes entre si, como c�digos de CST, por exemplo.
	' Para que essas derived tables tenham apenas um registro por NFe ou produto, � usada a t�cnica em que se obt�m o ID mais recente de um grupo de registros similares.
	if ckb_COL_ICMS_UF_DEST <> "" then
		s_sql = s_sql & _
				" LEFT JOIN (" & _
					"SELECT" & _
						" t_NFE_IMAGEM.*" & _
					" FROM t_NFe_IMAGEM INNER JOIN (" & _
						"SELECT" & _
							" id_nfe_emitente, NFe_serie_NF, NFe_numero_NF, Max(id) AS id" & _
						" FROM t_NFe_IMAGEM" & _
						" WHERE" & _
							" (t_NFe_IMAGEM.ide__tpNF = '1')" & _
							" AND (t_NFe_IMAGEM.st_anulado = 0)" & _
							" AND (t_NFe_IMAGEM.codigo_retorno_NFe_T1 = 1)" & _
						" GROUP BY" & _
							" id_nfe_emitente, NFe_serie_NF, NFe_numero_NF" & _
						") t_NFe_IMAGEM_max_id" & _
						" ON (t_NFe_IMAGEM_max_id.id = t_NFe_IMAGEM.id) AND (t_NFe_IMAGEM_max_id.id_nfe_emitente = t_NFe_IMAGEM.id_nfe_emitente)" & _
					" WHERE" & _
						" (t_NFe_IMAGEM.ide__tpNF = '1') AND (t_NFe_IMAGEM.st_anulado = 0) AND (t_NFe_IMAGEM.codigo_retorno_NFe_T1 = 1)" & _
				") t_NFe_IMAGEM_NORMALIZADO ON (t_NFe_IMAGEM_NORMALIZADO.id_nfe_emitente = t_PEDIDO.id_nfe_emitente) AND (t_NFe_IMAGEM_NORMALIZADO.NFe_numero_NF = t_PEDIDO.num_obs_2)"

		s_sql = s_sql & _
				" LEFT JOIN (" & _
					"SELECT" & _
						" t_NFe_IMAGEM_ITEM.*" & _
					" FROM t_NFe_IMAGEM_ITEM INNER JOIN (" & _
						"SELECT" & _
							" id_nfe_imagem, fabricante, produto, Max(id) AS id" & _
						" FROM t_NFe_IMAGEM_ITEM" & _
						" GROUP BY id_nfe_imagem, fabricante, produto" & _
						") t_NFe_IMAGEM_ITEM_max_id ON (t_NFe_IMAGEM_ITEM_max_id.id = t_NFe_IMAGEM_ITEM.id)" & _
					") t_NFe_IMAGEM_ITEM_NORMALIZADO ON (t_NFe_IMAGEM_NORMALIZADO.id = t_NFe_IMAGEM_ITEM_NORMALIZADO.id_nfe_imagem) AND (t_NFe_IMAGEM_ITEM_NORMALIZADO.fabricante = t_PEDIDO_ITEM.fabricante) AND (t_NFe_IMAGEM_ITEM_NORMALIZADO.produto = t_PEDIDO_ITEM.produto)"
		end if 'if ckb_COL_ICMS_UF_DEST <> ""

	s_sql = s_sql & _
			" LEFT JOIN t_CODIGO_DESCRICAO tPedidoOrigemDescricao ON (tPedidoOrigemDescricao.grupo = 'PedidoECommerce_Origem') AND (tPedidoOrigemDescricao.codigo = t_PEDIDO.marketplace_codigo_origem)" & _
			" LEFT JOIN t_CODIGO_DESCRICAO tGrupoPedidoOrigemDescricao ON (tGrupoPedidoOrigemDescricao.grupo = 'PedidoECommerce_Origem_Grupo') AND (tGrupoPedidoOrigemDescricao.codigo = tPedidoOrigemDescricao.codigo_pai)"

	if c_entrega_imediata <> "" then
		'QUANDO O FILTRO STATUS DA ENTREGA IMEDIATA � USADO, OS PEDIDOS A SEREM PESQUISADOS AINDA N�O EST�O COMO ENTREGUES, PODENDO AT� ESTAR AGUARDANDO A CHEGADA DE PRODUTOS NO ESTOQUE
		s_sql = s_sql & _
				" WHERE" & _
					" (t_PEDIDO.st_entrega NOT IN ('" & ST_ENTREGA_CANCELADO & "', '" & ST_ENTREGA_ENTREGUE & "'))"
	else
		s_sql = s_sql & _
				" WHERE" & _
					" (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')"
		end if

	if (ckb_COL_VL_CUSTO_REAL <> "") Or (ckb_COL_VL_CUSTO_REAL_TOTAL <> "") Or (ckb_COL_NAC_IMP <> "") then
		s_sql = s_sql & _
				" AND (t_ESTOQUE_MOVIMENTO.anulado_status=0)" & _
				" AND (t_ESTOQUE_MOVIMENTO.estoque='" & ID_ESTOQUE_ENTREGUE & "')"
		end if

	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then
		s_sql = s_sql & " AND" & s
		end if
	
	if ckb_CONSOLIDAR_PEDIDO <> "" then
		s_sql_aux = " operacao" & _
					", data_hora" & _
					", dt_emissao" & _
					", dt_emissao_remessa" & _
					", id_nfe_emitente" & _
					", numero_NF" & _
					", obs_2" & _
					", numero_NF_remessa" & _
					", obs_3" & _
					", loja" & _
					", pedido" & _
					", pedido_bs_x_marketplace" & _
					", marketplace_codigo_origem" & _
					", st_entrega" & _
					", entregue_data" & _
					", st_etg_imediata" & _
					", PrevisaoEntregaData" & _
					", PrevisaoEntregaTranspData" & _
					", PedidoRecebidoStatus" & _
					", PedidoRecebidoData" & _
					", GrupoPedidoOrigemDescricao" & _
					", PedidoOrigemDescricao" & _
					", transportadora_id" & _
					", vendedor" & _
					", indicador" & _
					", perc_RT" & _
					", nome_cliente" & _
					", tipo_cliente" & _
					", cnpj_cpf" & _
					", contribuinte_icms_status" & _
					", produtor_rural_status" & _
					", cidade" & _
					", uf" & _
					", st_end_entrega" & _
					", EndEtg_cidade" & _
					", EndEtg_uf" & _
					", ddd_res" & _
					", tel_res" & _
					", ddd_cel" & _
					", tel_cel" & _
					", ddd_com" & _
					", tel_com" & _
					", ddd_com_2" & _
					", tel_com_2" & _
					", ramal_com" & _
					", ramal_com_2" & _
					", email" & _
					", indicador_cnpj_cpf" & _
					", indicador_endereco" & _
					", indicador_endereco_numero" & _
					", indicador_endereco_complemento" & _
					", indicador_bairro" & _
					", indicador_cidade" & _
					", indicador_uf" & _
					", indicador_cep" & _
					", indicador_email" & _
					", indicador_email2" & _
					", indicador_email3" & _
					", qtde_parcelas" & _
					", tipo_parcelamento" & _
					", forma_pagamento_av" & _
					", parcelamento_c_entrada" & _
					", parcelamento_s_entrada" & _
					", parcela_unica"

		if ckb_COL_FRETE <> "" then
			if ckb_COL_FRETE_DETALHADO = "" then
				s_sql_aux = s_sql_aux & _
						", vl_frete"
			else
				for i=LBound(vTipoFrete) to UBound(vTipoFrete)
					with vTipoFrete(i)
						if .codigo <> "" then
							s_sql_aux = s_sql_aux & _
									", " & .colSQL
							end if
						end with
					next
				end if

			s_sql_aux = s_sql_aux & _
					", vl_total_produtos_calc_frete"
			end if
		
		if ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET <> "" then
			s_sql_aux = s_sql_aux & _
					", vl_pago_cartao_internet_total_familia_pedido"
			end if

		if ckb_COL_VL_PAGO_CARTAO_INTERNET <> "" then
			s_sql_aux = s_sql_aux & _
					", vl_pago_cartao_internet_total_pedido"
			end if

		if ckb_COL_ICMS_UF_DEST <> "" then
			s_sql_aux = s_sql_aux & _
					", SUM(Convert(money, vICMSUFDest)) AS vICMSUFDestTotalPedido"
			end if

		s_sql_aux = s_sql_aux & _
					", SUM(cubagem_total_item) AS cubagem_total_pedido" & _
					", SUM(peso_total_item) AS peso_total_pedido" & _
					", SUM(qtde_volumes_total_item) AS qtde_volumes_total_pedido" & _
					", SUM(preco_NF_total_item) AS preco_NF_total_pedido" & _
					", SUM(preco_venda_total_item) AS preco_venda_total_pedido"

		if (ckb_COL_VL_CUSTO_REAL <> "") Or (ckb_COL_VL_CUSTO_REAL_TOTAL <> "") then
			s_sql_aux = s_sql_aux & _
					", SUM(vl_custo2_real_total_item) AS vl_custo2_real_total_pedido"
			end if

		s_where_aux = ""
		if ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (vl_pago_cartao_internet_total_familia_pedido > 0)"
			end if

		if s_where_aux <> "" then s_where_aux = " WHERE" & s_where_aux

		s_sql = "SELECT " & _
					s_sql_aux & _
				" FROM (" & _
					s_sql & _
					") tVendaNormal" & _
				s_where_aux & _
				" GROUP BY" & _
					" operacao" & _
					", data_hora" & _
					", dt_emissao" & _
					", dt_emissao_remessa" & _
					", id_nfe_emitente" & _
					", numero_NF" & _
					", obs_2" & _
					", numero_NF_remessa" & _
					", obs_3" & _
					", loja" & _
					", pedido" & _
					", pedido_bs_x_marketplace" & _
					", marketplace_codigo_origem" & _
					", st_entrega" & _
					", entregue_data" & _
					", st_etg_imediata" & _
					", PrevisaoEntregaData" & _
					", PrevisaoEntregaTranspData" & _
					", PedidoRecebidoStatus" & _
					", PedidoRecebidoData" & _
					", GrupoPedidoOrigemDescricao" & _
					", PedidoOrigemDescricao" & _
					", transportadora_id" & _
					", vendedor" & _
					", indicador" & _
					", perc_RT" & _
					", nome_cliente" & _
					", tipo_cliente" & _
					", cnpj_cpf" & _
					", contribuinte_icms_status" & _
					", produtor_rural_status" & _
					", cidade" & _
					", uf" & _
					", st_end_entrega" & _
					", EndEtg_cidade" & _
					", EndEtg_uf" & _
					", ddd_res" & _
					", tel_res" & _
					", ddd_cel" & _
					", tel_cel" & _
					", ddd_com" & _
					", tel_com" & _
					", ddd_com_2" & _
					", tel_com_2" & _
					", ramal_com" & _
					", ramal_com_2" & _
					", email" & _
					", indicador_cnpj_cpf" & _
					", indicador_endereco" & _
					", indicador_endereco_numero" & _
					", indicador_endereco_complemento" & _
					", indicador_bairro" & _
					", indicador_cidade" & _
					", indicador_uf" & _
					", indicador_cep" & _
					", indicador_email" & _
					", indicador_email2" & _
					", indicador_email3" & _
					", qtde_parcelas" & _
					", tipo_parcelamento" & _
					", forma_pagamento_av" & _
					", parcelamento_c_entrada" & _
					", parcelamento_s_entrada" & _
					", parcela_unica"

		if ckb_COL_FRETE <> "" then
			if ckb_COL_FRETE_DETALHADO = "" then
				s_sql = s_sql & _
						", vl_frete"
			else
				for i=LBound(vTipoFrete) to UBound(vTipoFrete)
					with vTipoFrete(i)
						if .codigo <> "" then
							s_sql = s_sql & _
									", " & .colSQL
							end if
						end with
					next
				end if

			s_sql = s_sql & _
					", vl_total_produtos_calc_frete"
			end if

		if ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET <> "" then
			s_sql = s_sql & _
					", vl_pago_cartao_internet_total_familia_pedido"
			end if

		if ckb_COL_VL_PAGO_CARTAO_INTERNET <> "" then
			s_sql = s_sql & _
					", vl_pago_cartao_internet_total_pedido"
			end if
		end if

	'NA CONSOLIDA��O POR PEDIDO, AS DEVOLU��ES S�O IGNORADAS (DEFINI��O DO NICHOLAS EM 23/08/2021)
	if ckb_CONSOLIDAR_PEDIDO = "" then
	'	DEVOLU��ES
	'	OBS: O USO DE 'UNION' SIMPLES ELIMINA AS LINHAS DUPLICADAS DOS RESULTADOS
	'		 O USO DE 'UNION ALL' RETORNARIA TODAS AS LINHAS, INCLUSIVE AS DUPLICADAS
		s_sql = s_sql & " UNION ALL " & _
				"SELECT" & _
					" 'DEVOLUCAO' AS operacao" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data_hora" & _
					", NULL AS dt_emissao" & _
					", NULL AS dt_emissao_remessa" & _
					", t_PEDIDO.id_nfe_emitente" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.NFe_numero_NF AS numero_NF" & _
					", t_PEDIDO.obs_2" & _
					", 0 AS numero_NF_remessa" & _
					", t_PEDIDO.obs_3" & _
					", t_PEDIDO__BASE.loja" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.pedido" & _
					", t_PEDIDO.pedido_bs_x_marketplace" & _
					", t_PEDIDO.marketplace_codigo_origem" & _
					", t_PEDIDO.st_entrega" & _
					", t_PEDIDO.entregue_data" & _
					", t_PEDIDO.st_etg_imediata" & _
					", t_PEDIDO.PrevisaoEntregaData" & _
					", t_PEDIDO.PrevisaoEntregaTranspData" & _
					", t_PEDIDO.PedidoRecebidoStatus" & _
					", t_PEDIDO.PedidoRecebidoData" & _
					", tGrupoPedidoOrigemDescricao.descricao AS GrupoPedidoOrigemDescricao" & _
					", tPedidoOrigemDescricao.descricao AS PedidoOrigemDescricao" & _
					", t_PEDIDO.transportadora_id" & _
					", t_PEDIDO__BASE.vendedor" & _
					", t_PEDIDO__BASE.indicador" & _
					", t_PEDIDO__BASE.perc_RT"

		if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
			s_sql = s_sql & _
					", t_PEDIDO.endereco_nome AS nome_cliente" & _
					", t_PEDIDO.endereco_tipo_pessoa AS tipo_cliente" & _
					", t_PEDIDO.endereco_cnpj_cpf AS cnpj_cpf" & _
					", t_PEDIDO.endereco_contribuinte_icms_status AS contribuinte_icms_status" & _
					", t_PEDIDO.endereco_produtor_rural_status AS produtor_rural_status" & _
					", t_PEDIDO.endereco_cidade AS cidade" & _
					", t_PEDIDO.endereco_uf AS uf" & _
					", t_PEDIDO.endereco_ddd_res AS ddd_res" & _
					", t_PEDIDO.endereco_tel_res AS tel_res" & _
					", t_PEDIDO.endereco_ddd_cel AS ddd_cel" & _
					", t_PEDIDO.endereco_tel_cel AS tel_cel" & _
					", t_PEDIDO.endereco_ddd_com AS ddd_com" & _
					", t_PEDIDO.endereco_tel_com AS tel_com" & _
					", t_PEDIDO.endereco_ddd_com_2 AS ddd_com_2" & _
					", t_PEDIDO.endereco_tel_com_2 AS tel_com_2" & _
					", t_PEDIDO.endereco_ramal_com AS ramal_com" & _
					", t_PEDIDO.endereco_ramal_com_2 AS ramal_com_2" & _
					", t_PEDIDO.endereco_email AS email"
		else
			s_sql = s_sql & _
					", t_CLIENTE.nome_iniciais_em_maiusculas AS nome_cliente" & _
					", t_CLIENTE.tipo AS tipo_cliente" & _
					", t_CLIENTE.cnpj_cpf" & _
					", t_CLIENTE.contribuinte_icms_status" & _
					", t_CLIENTE.produtor_rural_status" & _
					", t_CLIENTE.cidade AS cidade" & _
					", t_CLIENTE.uf AS uf" & _
					", t_CLIENTE.ddd_res" & _
					", t_CLIENTE.tel_res" & _
					", t_CLIENTE.ddd_cel" & _
					", t_CLIENTE.tel_cel" & _
					", t_CLIENTE.ddd_com" & _
					", t_CLIENTE.tel_com" & _
					", t_CLIENTE.ddd_com_2" & _
					", t_CLIENTE.tel_com_2" & _
					", t_CLIENTE.ramal_com" & _
					", t_CLIENTE.ramal_com_2" & _
					", t_CLIENTE.email"
			end if

		s_sql = s_sql & _
					", t_PEDIDO.st_end_entrega" & _
					", t_PEDIDO.EndEtg_cidade" & _
					", t_PEDIDO.EndEtg_uf" & _
					", t_ORCAMENTISTA_E_INDICADOR.cnpj_cpf AS indicador_cnpj_cpf" & _
					", t_ORCAMENTISTA_E_INDICADOR.endereco AS indicador_endereco" & _
					", t_ORCAMENTISTA_E_INDICADOR.endereco_numero AS indicador_endereco_numero" & _
					", t_ORCAMENTISTA_E_INDICADOR.endereco_complemento AS indicador_endereco_complemento" & _
					", t_ORCAMENTISTA_E_INDICADOR.bairro AS indicador_bairro" & _
					", t_ORCAMENTISTA_E_INDICADOR.cidade AS indicador_cidade" & _
					", t_ORCAMENTISTA_E_INDICADOR.uf AS indicador_uf" & _
					", t_ORCAMENTISTA_E_INDICADOR.cep AS indicador_cep" & _
					", t_ORCAMENTISTA_E_INDICADOR.email AS indicador_email" & _
					", t_ORCAMENTISTA_E_INDICADOR.email2 AS indicador_email2" & _
					", t_ORCAMENTISTA_E_INDICADOR.email3 AS indicador_email3" & _
					", t_PEDIDO.qtde_parcelas AS qtde_parcelas" & _
					", t_PEDIDO__BASE.tipo_parcelamento AS tipo_parcelamento" & _
					", t_PEDIDO__BASE.av_forma_pagto AS forma_pagamento_av" & _
					", t_PEDIDO__BASE.pce_forma_pagto_prestacao AS parcelamento_c_entrada" & _
					", t_PEDIDO__BASE.pse_forma_pagto_demais_prest AS parcelamento_s_entrada" & _
					", t_PEDIDO__BASE.pu_forma_pagto AS parcela_unica"

		s_where_aux = ""
		if s_where_lista_codigo_frete_devolucao <> "" then
		'	SOMENTE OS FRETES DE DEVOLU��O
			s_where_aux = " AND (codigo_tipo_frete IN " & s_where_lista_codigo_frete_devolucao & ")"
			end if

		if ckb_COL_FRETE <> "" then
			if ckb_COL_FRETE_DETALHADO = "" then
				s_sql = s_sql & _
							", (SELECT Coalesce(SUM(vl_frete),0) AS vl_frete FROM t_PEDIDO_FRETE WHERE (t_PEDIDO_FRETE.pedido=t_PEDIDO.pedido)" & s_where_aux & ") AS vl_frete"
			else
				for i=LBound(vTipoFrete) to UBound(vTipoFrete)
					with vTipoFrete(i)
						if .codigo <> "" then
							'QUANDO A LISTAGEM � ITEM A ITEM, O VALOR DO FRETE DE DEVOLU��O SOMENTE � EXIBIDO NOS ITENS DEVOLVIDOS E VICE-VERSA
							if (ckb_CONSOLIDAR_PEDIDO = "") And (Not .blnFreteDevolucao) then
								s_sql = s_sql & _
											", 0 AS " & .colSQL
							else
								s_sql = s_sql & _
											", (SELECT Coalesce(SUM(vl_frete),0) AS vl_frete FROM t_PEDIDO_FRETE WHERE (t_PEDIDO_FRETE.pedido=t_PEDIDO.pedido) AND (codigo_tipo_frete = '" & .codigo & "')) AS " & .colSQL
								end if
							end if
						end with
					next
				end if

			s_sql = s_sql & _
						", (SELECT Coalesce(SUM(qtde * preco_venda),0) AS vl_total_produtos_calc_frete FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (t_PEDIDO_ITEM_DEVOLVIDO.pedido = t_PEDIDO.pedido)) AS vl_total_produtos_calc_frete"
			end if

		if (ckb_COL_VL_CUSTO_REAL <> "") Or (ckb_COL_VL_CUSTO_REAL_TOTAL <> "") then
			s_sql = s_sql & _
					", -t_ESTOQUE_ITEM.qtde"
		else
			s_sql = s_sql & _
					", -t_PEDIDO_ITEM_DEVOLVIDO.qtde"
			end if

		if ckb_COL_NAC_IMP <> "" then
			s_sql = s_sql & _
					", t_ESTOQUE_ITEM.cst"
			end if

		s_sql = s_sql & _
					", t_PEDIDO_ITEM_DEVOLVIDO.fabricante" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.produto" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.descricao" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.preco_venda" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.preco_lista" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.preco_NF" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.desc_dado" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.cubagem" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.peso" & _
					", t_PEDIDO_ITEM_DEVOLVIDO.qtde_volumes" & _
					", t_PRODUTO.grupo" & _
					", t_PRODUTO.potencia_BTU" & _
					", t_PRODUTO.ciclo" & _
					", t_PRODUTO.posicao_mercado" & _
					", t_FABRICANTE.nome AS nome_fabricante"

		s_sql = s_sql & _
					", (SELECT TOP 1 vl_custo2 FROM t_ESTOQUE tE INNER JOIN t_ESTOQUE_ITEM tEI ON (tE.id_estoque = tEI.id_estoque) WHERE (tE.devolucao_status = 0) AND (tEI.fabricante = t_PEDIDO_ITEM_DEVOLVIDO.fabricante) AND (tEI.produto = t_PEDIDO_ITEM_DEVOLVIDO.produto) AND (tEI.qtde > 0) ORDER BY tEI.id_estoque DESC) AS vl_custo2_ult_entrada"

		if (ckb_COL_VL_CUSTO_REAL <> "") Or (ckb_COL_VL_CUSTO_REAL_TOTAL <> "") then
			s_sql = s_sql & _
					", (t_ESTOQUE_ITEM.vl_custo2) AS vl_custo2_real"
			end if

		if ckb_COL_ICMS_UF_DEST <> "" then
			s_sql = s_sql & _
					", NULL AS vICMSUFDest" & _
					", NULL AS det__qCom"
			end if

		s_sql = s_sql & _
				" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
					" INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido = t_PEDIDO.pedido)" & _
					" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base = t_PEDIDO__BASE.pedido)" & _
					" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador = t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
					" INNER JOIN t_PRODUTO ON (t_PRODUTO.fabricante = t_PEDIDO_ITEM_DEVOLVIDO.fabricante) AND (t_PRODUTO.produto = t_PEDIDO_ITEM_DEVOLVIDO.produto)" & _
					" INNER JOIN t_FABRICANTE ON (t_PRODUTO.fabricante = t_FABRICANTE.fabricante)" & _
					" INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente = t_CLIENTE.id)"

		if (ckb_COL_VL_CUSTO_REAL <> "") Or (ckb_COL_VL_CUSTO_REAL_TOTAL <> "") Or (ckb_COL_NAC_IMP <> "") then
			s_sql = s_sql & _
					" INNER JOIN t_ESTOQUE ON (t_PEDIDO_ITEM_DEVOLVIDO.id=t_ESTOQUE.devolucao_id_item_devolvido)" & _
					" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)AND(t_PEDIDO_ITEM_DEVOLVIDO.fabricante=t_ESTOQUE_ITEM.fabricante)AND(t_PEDIDO_ITEM_DEVOLVIDO.produto=t_ESTOQUE_ITEM.produto)AND(t_ESTOQUE_ITEM.qtde > 0))"
			end if

		s_sql = s_sql & _
				" LEFT JOIN t_CODIGO_DESCRICAO tPedidoOrigemDescricao ON (tPedidoOrigemDescricao.grupo = 'PedidoECommerce_Origem') AND (tPedidoOrigemDescricao.codigo = t_PEDIDO.marketplace_codigo_origem)" & _
				" LEFT JOIN t_CODIGO_DESCRICAO tGrupoPedidoOrigemDescricao ON (tGrupoPedidoOrigemDescricao.grupo = 'PedidoECommerce_Origem_Grupo') AND (tGrupoPedidoOrigemDescricao.codigo = tPedidoOrigemDescricao.codigo_pai)"

		if c_entrega_imediata <> "" then
			'QUANDO O FILTRO STATUS DA ENTREGA IMEDIATA � USADO, OS PEDIDOS A SEREM PESQUISADOS AINDA N�O EST�O COMO ENTREGUES, PODENDO AT� ESTAR AGUARDANDO A CHEGADA DE PRODUTOS NO ESTOQUE
			s_sql = s_sql & _
					" WHERE" & _
						" (t_PEDIDO.st_entrega NOT IN ('" & ST_ENTREGA_CANCELADO & "', '" & ST_ENTREGA_ENTREGUE & "'))"
		else
			s_sql = s_sql & _
					" WHERE" & _
						" (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')"
			end if
	
		s = s_where
		if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
		s = s & s_where_devolucao
		if s <> "" then
			s_sql = s_sql & " AND" & s
			end if
		
		end if 'if ckb_CONSOLIDAR_PEDIDO = ""


	s_sql = "SELECT " & _
				"*" & _
			" FROM (" & s_sql & ") t"
	
	s_where_aux = ""

	' Tratamento para evitar erro que ocorre quando h� registro do estoque com o campo 'qtde' com valor zerado.
	' Essa situa��o em que a 'qtde' � zero n�o deveria ocorrer, entretanto, devido a algumas corre��es de problemas ocorridos anteriormente em opera��es no estoque,
	' h� alguns registros de entrada no estoque em que a 'qtde' foi ajustada para zero atrav�s de interven��o manual no banco de dados.
	' Lembrando que as devolu��es s�o tratadas com valores negativos de qtde.
	if ckb_CONSOLIDAR_PEDIDO = "" then
		if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
		s_where_aux = s_where_aux & " (qtde <> 0)"
		end if

	if s_where_periodo_NF <> "" then
		if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
		s_where_aux = s_where_aux & s_where_periodo_NF
		end if

	if s_where_aux <> "" then s_where_aux = " WHERE" & s_where_aux

	s_sql = s_sql & _
			s_where_aux

	s_sql = s_sql & _
			" ORDER BY" & _
				" pedido"
	
	if ckb_CONSOLIDAR_PEDIDO = "" then
		s_sql = s_sql & _
				", fabricante" & _
				", produto" & _
				", qtde"
		end if
	
	if ckb_COL_CHAVE_NFE <> "" then
		'OBT�M OS N�MERO DE NF QUE DEVEM SER CONSULTADOS PARA OBTER A CHAVE DE ACESSO
		n_reg_total_passo1 = 0
		if r.State <> 0 then r.Close
		r.open s_sql, cn
		do while Not r.Eof
			n_reg_total_passo1 = n_reg_total_passo1 + 1

			if Trim("" & r("id_nfe_emitente")) <> "" then
				for i=LBound(vNFeAConsultar) to UBound(vNFeAConsultar)
					if Trim("" & r("id_nfe_emitente")) = vNFeAConsultar(i).c1 then
						if (Trim("" & r("numero_NF")) <> "") And (Trim("" & r("operacao")) = "VENDA_NORMAL") then
							s = "'" & NFeFormataNumeroNF(Trim("" & r("numero_NF"))) & "'"
							if Instr(vNFeAConsultar(i).c3, s) = 0 then
								if vNFeAConsultar(i).c3 <> "" then vNFeAConsultar(i).c3 = vNFeAConsultar(i).c3 & ","
								vNFeAConsultar(i).c3 = vNFeAConsultar(i).c3 & s
								end if
							end if
						exit for
						end if
					next
				end if

			r.MoveNext
			loop
		
		'REALIZA A CONSULTA EM CADA BD DE NFE
		for i=LBound(vNFeAConsultar) to UBound(vNFeAConsultar)
			if (vNFeAConsultar(i).c1 <> "") And (Not (vNFeAConsultar(i).c2 is nothing)) And (vNFeAConsultar(i).c3 <> "") then
				s = "SELECT DISTINCT" & _
						" Serie," & _
						" Nfe," & _
						" ChaveAcesso" & _
					" FROM NFE" & _
					" WHERE" & _
						" (CodProcAtual = 100)" & _
						" AND (Coalesce(CANCELADA,0) = 0)" & _
						" AND (LEN(RTRIM(Coalesce(ChaveAcesso,''))) > 0)" & _
						" AND (Nfe IN (" & vNFeAConsultar(i).c3 & "))"
				if r.State <> 0 then r.Close
				r.open s, vNFeAConsultar(i).c2
				do while Not r.Eof
					if vNFeChave(UBound(vNFeChave)).c1 <> "" then
						redim preserve vNFeChave(UBound(vNFeChave)+1)
						set vNFeChave(UBound(vNFeChave)) = New cl_DEZ_COLUNAS
						vNFeChave(UBound(vNFeChave)).CampoOrdenacao = ""
						vNFeChave(UBound(vNFeChave)).c1 = ""
						vNFeChave(UBound(vNFeChave)).c2 = ""
						vNFeChave(UBound(vNFeChave)).c3 = ""
						end if

					vNFeChave(UBound(vNFeChave)).c1 = vNFeAConsultar(i).c1
					vNFeChave(UBound(vNFeChave)).c2 = Trim("" & r("Nfe"))
					vNFeChave(UBound(vNFeChave)).c3 = Trim("" & r("ChaveAcesso"))
					vNFeChave(UBound(vNFeChave)).CampoOrdenacao = normaliza_codigo(vNFeAConsultar(i).c1, 6) & "|" & Trim("" & r("Nfe"))

					r.MoveNext
					loop
				end if
			next

		'FECHA AS CONEX�ES COM O BD DE NFE
		for i=LBound(vNFeAConsultar) to UBound(vNFeAConsultar)
			if Not (vNFeAConsultar(i).c2 is nothing) then
				vNFeAConsultar(i).c2.Close
				set vNFeAConsultar(i).c2 = nothing
				end if
			next
		
		ordena_cl_dez_colunas vNFeChave, 0, UBound(vNFeChave)
		end if 'if ckb_COL_CHAVE_NFE <> ""


	x_cab = ""
	if ckb_COL_DT_CADASTRO <> "" then x_cab = x_cab & "Data (Cadastro);"
	if ckb_COL_NF <> "" then x_cab = x_cab & "NF;"
	if ckb_COL_DT_EMISSAO_NF <> "" then x_cab = x_cab & "Emissao NF;"
	if ckb_COL_NF_REMESSA <> "" then x_cab = x_cab & "NF Remessa;"
	if ckb_COL_DT_EMISSAO_NF_REMESSA <> "" then x_cab = x_cab & "Emissao NF Remessa;"
	if ckb_COL_LOJA <> "" then x_cab = x_cab & "Loja;"
	if ckb_COL_PEDIDO <> "" then x_cab = x_cab & "Pedido;"
	if ckb_COL_PEDIDO_MARKETPLACE <> "" then x_cab = x_cab & "Pedido Marketplace;"
	if ckb_COL_GRUPO_PEDIDO_ORIGEM <> "" then x_cab = x_cab & "Origem Pedido (Grupo);"
	if ckb_COL_CPF_CNPJ_CLIENTE <> "" then x_cab = x_cab & "CPF/CNPJ;"
	if ckb_COL_CONTRIBUINTE_ICMS <> "" then x_cab = x_cab & "Contrib ICMS;"
	if ckb_COL_NOME_CLIENTE <> "" then x_cab = x_cab & "Nome Cliente;"
	if ckb_COL_CIDADE <> "" then x_cab = x_cab & "Cidade (cadastro);"
	if ckb_COL_UF <> "" then x_cab = x_cab & "UF (cadastro);"
	if ckb_COL_CIDADE_ETG <> "" then x_cab = x_cab & "Cidade (entrega);"
	if ckb_COL_UF_ETG <> "" then x_cab = x_cab & "UF (entrega);"
    if ckb_COL_TEL <> "" then x_cab = x_cab & "Telefone;Telefone;Telefone;"
    if ckb_COL_EMAIL <> "" then x_cab = x_cab & "Email;"
	if ckb_COL_VENDEDOR <> "" then x_cab = x_cab & "Vendedor;"
	if ckb_COL_INDICADOR <> "" then x_cab = x_cab & "Indicador;"
	if ckb_COL_TRANSPORTADORA <> "" then x_cab = x_cab & "Transportadora;"
	if ckb_COL_ENTREGA_IMEDIATA <> "" then x_cab = x_cab & "Entrega Imediata;Previsao Entrega;"
	if ckb_COL_DT_ENTREGA <> "" then x_cab = x_cab & "Data de Entrega;"
	if ckb_COL_DT_PREVISAO_ETG_TRANSP <> "" then x_cab = x_cab & "Previs�o de Entrega (Transp);"
	if ckb_COL_DT_RECEB_CLIENTE <> "" then x_cab = x_cab & "Receb Cliente;"
	if ckb_COL_INDICADOR_CPF_CNPJ <> "" then x_cab = x_cab & "CPF/CNPJ Ind;"
	if ckb_COL_INDICADOR_ENDERECO <> "" then x_cab = x_cab & "Endereco Ind;"
	if ckb_COL_INDICADOR_CIDADE <> "" then x_cab = x_cab & "Cidade Ind;"
	if ckb_COL_INDICADOR_UF <> "" then x_cab = x_cab & "UF Ind;"
	if ckb_COL_INDICADOR_EMAILS <> "" then x_cab = x_cab & "Email Ind;Email Ind;Email Ind;"
	if ckb_COL_MARCA <> "" then x_cab = x_cab & "Marca;"
	if ckb_COL_GRUPO <> "" then x_cab = x_cab & "Grupo;"
	if ckb_COL_POTENCIA_BTU <> "" then x_cab = x_cab & "BTU;"
	if ckb_COL_CICLO <> "" then x_cab = x_cab & "Ciclo;"
	if ckb_COL_POSICAO_MERCADO <> "" then x_cab = x_cab & "Pos Merc;"
	if ckb_COL_PRODUTO <> "" then x_cab = x_cab & "Produto;"
	if ckb_COL_NAC_IMP <> "" then x_cab = x_cab & "Nac/Imp;"
	if ckb_COL_DESCRICAO_PRODUTO <> "" then x_cab = x_cab & "Descricao;"
	if ckb_COL_QTDE <> "" then x_cab = x_cab & "Qtde;"
	if ckb_COL_PERC_DESC <> "" then x_cab = x_cab & "Desc %;"
	if ckb_COL_CUBAGEM <> "" then x_cab = x_cab & "Cubagem;"
	if ckb_COL_PESO <> "" then x_cab = x_cab & "Peso;"
	if ckb_COL_QTDE_VOLUMES <> "" then x_cab = x_cab & "Volumes;"
	
	if ckb_COL_FRETE <> "" then
		if ckb_COL_FRETE_DETALHADO = "" then
			x_cab = x_cab & "VL Frete;"
		else
			for i=LBound(vTipoFrete) to UBound(vTipoFrete)
				with vTipoFrete(i)
					if .codigo <> "" then
						x_cab = x_cab & "VL Frete (" & .descricao & ");"
						end if
					end with
				next
			end if
		end if

	if ckb_COL_VL_CUSTO_ULT_ENTRADA <> "" then x_cab = x_cab & "VL Custo (�lt Entrada);"
	if ckb_COL_VL_CUSTO_REAL <> "" then x_cab = x_cab & "VL Custo (Real);"
	if ckb_COL_VL_LISTA <> "" then x_cab = x_cab & "VL Lista;"
	if ckb_COL_VL_NF <> "" then x_cab = x_cab & "VL NF;"
	if ckb_COL_VL_UNITARIO <> "" then x_cab = x_cab & "VL Unitario;"
	if ckb_COL_VL_CUSTO_REAL_TOTAL <> "" then x_cab = x_cab & "VL Custo Total (Real);"
	if ckb_COL_VL_TOTAL_NF <> "" then x_cab = x_cab & "VL Total NF;"
	if ckb_COL_VL_TOTAL <> "" then x_cab = x_cab & "VL Total;"
	if ckb_COL_VL_RA <> "" then x_cab = x_cab & "VL RA;"
	if ckb_COL_RT <> "" then x_cab = x_cab & "RT;"
	if ckb_COL_ICMS_UF_DEST <> "" then
		if ckb_CONSOLIDAR_PEDIDO = "" then
			x_cab = x_cab & "ICMS UF Destino (Unit);"
		else
			x_cab = x_cab & "ICMS UF Destino;"
			end if
		end if
	if ckb_COL_QTDE_PARCELAS <> "" then x_cab = x_cab & "Qtde Parcelas;"
	if ckb_COL_MEIO_PAGAMENTO <> "" then x_cab = x_cab & "Meio de Pagamento;"
	if (ckb_COL_VL_PAGO_CARTAO_INTERNET <> "") And (ckb_CONSOLIDAR_PEDIDO <> "") then x_cab = x_cab & "VL Pago Cart�o (Internet);"
	if ckb_COL_CHAVE_NFE <> "" then x_cab = x_cab & "Chave NFE;"
	
	
	
	
	x = ""
	n_reg = 0
	n_reg_total = 0
    item_qtde = 1

	if r.State <> 0 then r.Close
	r.open s_sql, cn
	
	if Not r.Eof then x = x_cab & vbCrLf
	
	do while Not r.Eof
		
	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

        if ckb_AGRUPAMENTO <> "" then
            item_qtde = CInt(Trim("" & r("qtde")))
        end if

        for iQI=1 to Abs(item_qtde)

		 '> DATA (CADASTRO)
			if ckb_COL_DT_CADASTRO <> "" then
				x = x & formata_data(r("data_hora")) & ";"
				end if
		
		 '> NF
			if ckb_COL_NF <> "" then
				s = Trim("" & r("numero_NF"))
				if (Trim("" & r("operacao")) = "VENDA_NORMAL") And ((s = "") Or (s = "0")) then
					s = Trim("" & r("obs_2"))
					end if
				
				if s <> "" then
					if ckb_COMPATIBILIDADE <> "" then
						s = chr(34) & "=" & chr(34) & chr(34) & s & chr(34) & chr(34) & chr(34)
					else
						s = "=" & chr(34) & s & chr(34)
						end if
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

		 '> NF REMESSA
			if ckb_COL_NF_REMESSA <> "" then
				s = Trim("" & r("numero_NF_remessa"))
				if (Trim("" & r("operacao")) = "VENDA_NORMAL") And ((s = "") Or (s = "0")) then
					s = Trim("" & r("obs_3"))
					end if
				
				if s <> "" then
					if ckb_COMPATIBILIDADE <> "" then
						s = chr(34) & "=" & chr(34) & chr(34) & s & chr(34) & chr(34) & chr(34)
					else
						s = "=" & chr(34) & s & chr(34)
						end if
					end if

				x = x & s & ";"
				end if
		
		'> DATA DA EMISS�O NF REMESSA
			if ckb_COL_DT_EMISSAO_NF_REMESSA <> "" then
				if Trim("" & r("dt_emissao_remessa")) <> "" then 
					s = formata_data(r("dt_emissao_remessa"))
				else
					s = ""
					end if
				x = x & s & ";"
				end if

		 '> LOJA
			if ckb_COL_LOJA <> "" then
				x = x & Trim("" & r("loja")) & ";"
				end if

		 '> PEDIDO
			if ckb_COL_PEDIDO <> "" then
				x = x & Trim("" & r("pedido")) & ";"
				end if
			
		 '> PEDIDO MARKETPLACE
			if ckb_COL_PEDIDO_MARKETPLACE <> "" then
				'FOR�A PARA O EXCEL TRATAR COMO TEXTO E N�O N�MERO
				s = ""
				if ckb_COMPATIBILIDADE <> "" then
					s = chr(34) & "=" & chr(34) & chr(34) & Trim("" & r("pedido_bs_x_marketplace")) & chr(34) & chr(34) & chr(34)
				else
					s = "=" & chr(34) & Trim("" & r("pedido_bs_x_marketplace")) & chr(34)
					end if
				
				x = x & s & ";"
				end if

		 '> ORIGEM DO PEDIDO (GRUPO)
			if ckb_COL_GRUPO_PEDIDO_ORIGEM <> "" then
				x = x & Trim("" & r("GrupoPedidoOrigemDescricao")) & ";"
				end if

		'> CLIENTE: CPF/CNPJ
			if ckb_COL_CPF_CNPJ_CLIENTE <> "" then
				x = x & cnpj_cpf_formata(Trim("" & r("cnpj_cpf"))) & ";"
				end if
			
		'> CLIENTE: CONTRIBUINTE ICMS
			if ckb_COL_CONTRIBUINTE_ICMS <> "" then
				x = x & descricao_icms_contribuinte_x_produtor_rural(Trim("" & r("tipo_cliente")), Trim("" & r("contribuinte_icms_status")), Trim("" & r("produtor_rural_status"))) & ";"
				end if
			

		'> CLIENTE: NOME
			if ckb_COL_NOME_CLIENTE <> "" then
				s = Trim("" & r("nome_cliente"))
				s = Replace(s, ";", ",")
				x = x & s & ";"
				end if
			
		 '> CIDADE
			if ckb_COL_CIDADE <> "" then
				s = Trim("" & r("cidade"))
				s = Replace(s, ";", ",")
				x = x & s & ";"
			end if
	 
		 '> UF
			if ckb_COL_UF <> "" then
				x = x & Trim("" & r("uf")) & ";"
			end if

		 '> CIDADE (ETG)
			if ckb_COL_CIDADE_ETG <> "" then
				if r("st_end_entrega") <> 0 then
					s = Trim("" & r("EndEtg_cidade"))
				else
					s = Trim("" & r("cidade"))
					end if
				s = Replace(s, ";", ",")
				x = x & s & ";"
			end if
	 
		 '> UF (ETG)
			if ckb_COL_UF_ETG <> "" then
				if r("st_end_entrega") <> 0 then
					x = x & Trim("" & r("EndEtg_uf")) & ";"
				else
					x = x & Trim("" & r("uf")) & ";"
					end if
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
				s = Trim("" & r("email"))
				s = Replace(s, ";", ",")
				x = x & s & ";"
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

		 '> ENTREGA IMEDIATA + DATA DE PREVIS�O DE ENTREGA
			if ckb_COL_ENTREGA_IMEDIATA <> "" then
				s = ";"
				if Trim("" & r("operacao")) = "VENDA_NORMAL" then
					if Trim("" & r("st_etg_imediata")) = Cstr(COD_ETG_IMEDIATA_SIM) then
						s = "S;"
					elseif Trim("" & r("st_etg_imediata")) = Cstr(COD_ETG_IMEDIATA_NAO) then
						s = "N;" & formata_data(r("PrevisaoEntregaData"))
						end if
					end if
				x = x & s & ";"
				end if 
		
		 '> DATA (ENTREGA)
			if ckb_COL_DT_ENTREGA <> "" then
				s = ""
				if (Trim("" & r("operacao")) = "VENDA_NORMAL") And (Trim("" & r("st_entrega")) = ST_ENTREGA_ENTREGUE) And IsDate(r("entregue_data")) then
					s = formata_data(r("entregue_data"))
				elseif Trim("" & r("operacao")) = "DEVOLUCAO" then
					s = formata_data(r("data_hora"))
					end if
				x = x & s & ";"
				end if
		
		'> PREVIS�O DE ENTREGA DA TRANSPORTADORA
			if ckb_COL_DT_PREVISAO_ETG_TRANSP <> "" then
				s = ""
				if (Trim("" & r("operacao")) = "VENDA_NORMAL") And IsDate(r("PrevisaoEntregaTranspData")) then
					s = formata_data(r("PrevisaoEntregaTranspData"))
					end if
				x = x & s & ";"
				end if

		'> DATA DE RECEBIMENTO PELO CLIENTE
			if ckb_COL_DT_RECEB_CLIENTE <> "" then
				s = ""
				if (Trim("" & r("operacao")) = "VENDA_NORMAL") And (r("PedidoRecebidoStatus") = 1) And IsDate(r("PedidoRecebidoData")) then
					s = formata_data(r("PedidoRecebidoData"))
					end if
				x = x & s & ";"
				end if

		'> INDICADOR: CPF/CNPJ
			if ckb_COL_INDICADOR_CPF_CNPJ <> "" then
				x = x & cnpj_cpf_formata(Trim("" & r("indicador_cnpj_cpf"))) & ";"
				end if 

		'> INDICADOR: ENDERE�O
			if ckb_COL_INDICADOR_ENDERECO <> "" then
				s = formata_endereco(Trim("" & r("indicador_endereco")), Trim("" & r("indicador_endereco_numero")), Trim("" & r("indicador_endereco_complemento")), Trim("" & r("indicador_bairro")), "", "", Trim("" & r("indicador_cep")))
				s = Replace(s, ";", ",")
				x = x & s & ";"
				end if
			
		'> INDICADOR: CIDADE
			if ckb_COL_INDICADOR_CIDADE <> "" then
				s = Trim("" & r("indicador_cidade"))
				s = Replace(s, ";", ",")
				x = x & s & ";"
			end if
	 
		 '> INDICADOR: UF
			if ckb_COL_INDICADOR_UF <> "" then
				x = x & Trim("" & r("indicador_uf")) & ";"
			end if

		'> INDICADOR: E-MAIL 
			if ckb_COL_INDICADOR_EMAILS <> "" then
				s = Trim("" & r("indicador_email"))
				s = Replace(s, ";", ",")
				x = x & s & ";"
				end if

		'> INDICADOR: E-MAIL 2
			if ckb_COL_INDICADOR_EMAILS <> "" then
				s = Trim("" & r("indicador_email2"))
				s = Replace(s, ";", ",")
				x = x & s & ";"
				end if

		'> INDICADOR: E-MAIL 3
			if ckb_COL_INDICADOR_EMAILS <> "" then
				s = Trim("" & r("indicador_email3"))
				s = Replace(s, ";", ",")
				x = x & s & ";"
				end if

		'> NOME FABRICANTE
			if ckb_COL_MARCA <> "" then
				s = UCase(Trim("" & r("nome_fabricante")))
				s = Replace(s, ";", ",")
				x = x & s & ";"
				end if
			
		'> GRUPO
			if ckb_COL_GRUPO <> "" then
				s = Trim("" & r("grupo"))
				s = Replace(s, ";", ",")
				x = x & s & ";"
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
				if ckb_COMPATIBILIDADE <> "" then
					x = x & chr(34) & "=" & chr(34) & chr(34) & Trim("" & r("produto")) & chr(34) & chr(34) & chr(34) & ";"
				else
					x = x & "=" & chr(34) & Trim("" & r("produto")) & chr(34) & ";"
					end if
				end if

		 '> NACIONAL/IMPORTADO
			if ckb_COL_NAC_IMP <> "" then
				s_cst = Trim("" & r("cst"))
				if s_cst = "000" then
					s = "Nacional"
				elseif s_cst = "200" then
					s = "Importado"
				elseif (s_cst = "060") Or (s_cst = "241") Or (s_cst = "260") then
					s = "Importado"
				else
					s = s_cst
					end if
					
				x = x & s & ";"
				end if
		
		 '> DESCRI��O DO PRODUTO
			if ckb_COL_DESCRICAO_PRODUTO <> "" then
				s = Trim("" & r("descricao"))
				s = Replace(s, ";", ",")
				x = x & s & ";"
				end if
	
			if ckb_CONSOLIDAR_PEDIDO = "" then
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
				end if

		'> PERCENTUAL DE DESCONTO
			if ckb_COL_PERC_DESC <> "" then
				x = x & formata_perc_desc(Trim("" & r("desc_dado"))) & ";"
			end if

		'> CUBAGEM
			if ckb_COL_CUBAGEM <> "" then
				if ckb_CONSOLIDAR_PEDIDO = "" then
					item_cubagem = converte_numero(s_qtde) * converte_numero(r("cubagem"))
					x = x & formata_numero(item_cubagem, 2) & ";"
				else
					x = x & formata_numero(r("cubagem_total_pedido"), 2) & ";"
					end if
			end if

		'> PESO
			if ckb_COL_PESO <> "" then
				if ckb_CONSOLIDAR_PEDIDO = "" then
					item_peso = s_qtde * r("peso")
					x = x & item_peso & ";"
				else
					x = x & r("peso_total_pedido") & ";"
					end if
			end if

		'> QTDE VOLUMES
			if ckb_COL_QTDE_VOLUMES <> "" then
				if ckb_CONSOLIDAR_PEDIDO = "" then
					item_qtde_volumes = s_qtde * r("qtde_volumes")
					x = x & item_qtde_volumes & ";"
				else
					x = x & r("qtde_volumes_total_pedido") & ";"
					end if
			end if

		'> FRETE
			if ckb_COL_FRETE <> "" then
				if ckb_COL_FRETE_DETALHADO = "" then
					if ckb_CONSOLIDAR_PEDIDO = "" then
					'	CALCULA O VALOR PROPORCIONAL DO FRETE (LEMBRANDO QUE O VALOR DO FRETE OBTIDO � O TOTAL EM FRETES, MAS OS FRETES DE DEVOLU��O S�O COMPUTADOS APENAS P/ AS DEVOLU��ES)
						vl_frete_proporcional = 0
						if r("vl_total_produtos_calc_frete") <> 0 then
							vl_frete_proporcional = (Abs(CLng(s_qtde)) * r("preco_venda")) * (r("vl_frete") / r("vl_total_produtos_calc_frete"))
							end if
						s = formata_moeda(vl_frete_proporcional)
						if s = "" then s = 0
						x = x & s & ";"
					else 'if ckb_CONSOLIDAR_PEDIDO = ""
						s = formata_moeda(r("vl_frete"))
						if s = "" then s = 0
						x = x & s & ";"
						end if 'if ckb_CONSOLIDAR_PEDIDO = ""
				else 'if ckb_COL_FRETE_DETALHADO = ""
					if ckb_CONSOLIDAR_PEDIDO = "" then
					'	CALCULA O VALOR PROPORCIONAL DO FRETE (LEMBRANDO QUE O VALOR DO FRETE OBTIDO � O TOTAL EM FRETES, MAS OS FRETES DE DEVOLU��O S�O COMPUTADOS APENAS P/ AS DEVOLU��ES)
						for i=LBound(vTipoFrete) to UBound(vTipoFrete)
							with vTipoFrete(i)
								if .codigo <> "" then
									vl_frete_proporcional = 0
									if r("vl_total_produtos_calc_frete") <> 0 then
										vl_frete_proporcional = (Abs(CLng(s_qtde)) * r("preco_venda")) * (r(.colSQL) / r("vl_total_produtos_calc_frete"))
										end if
									s = formata_moeda(vl_frete_proporcional)
									if s = "" then s = 0
									x = x & s & ";"
									end if
								end with
							next
					else 'if ckb_CONSOLIDAR_PEDIDO = ""
						for i=LBound(vTipoFrete) to UBound(vTipoFrete)
							with vTipoFrete(i)
								if .codigo <> "" then
									s = formata_moeda(r(.colSQL))
									if s = "" then s = 0
									x = x & s & ";"
									end if
								end with
							next
						end if 'if ckb_CONSOLIDAR_PEDIDO = ""
					end if 'if ckb_COL_FRETE_DETALHADO = ""
			end if
			
		'> VALOR CUSTO (�LT ENTRADA)
			if ckb_COL_VL_CUSTO_ULT_ENTRADA <> "" then
			'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
				s = substitui_caracteres(bd_formata_moeda(r("vl_custo2_ult_entrada")), ".", SEPARADOR_DECIMAL)
				x = x & s & ";"
				end if

		'> VALOR CUSTO (REAL)
			if ckb_COL_VL_CUSTO_REAL <> "" then
			'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
				s = substitui_caracteres(bd_formata_moeda(r("vl_custo2_real")), ".", SEPARADOR_DECIMAL)
				x = x & s & ";"
				end if
		
		 '> PRE�O DE LISTA
			if ckb_COL_VL_LISTA <> "" then
			'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
				s = substitui_caracteres(bd_formata_moeda(r("preco_lista")), ".", SEPARADOR_DECIMAL)
				x = x & s & ";"
				end if
			
		'> VALOR NF
			if ckb_COL_VL_NF <> "" then
			'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
				s = substitui_caracteres(bd_formata_moeda(r("preco_NF")), ".", SEPARADOR_DECIMAL)
				x = x & s & ";"
				end if

		'> VALOR UNIT�RIO
			if ckb_COL_VL_UNITARIO <> "" then
			'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
				s = substitui_caracteres(bd_formata_moeda(r("preco_venda")), ".", SEPARADOR_DECIMAL)
				x = x & s & ";"
				end if
		
		'> VALOR CUSTO TOTAL (REAL)
			if ckb_COL_VL_CUSTO_REAL_TOTAL <> "" then
				if ckb_CONSOLIDAR_PEDIDO = "" then
				'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
					s = substitui_caracteres(bd_formata_moeda(CLng(s_qtde) * r("vl_custo2_real")), ".", SEPARADOR_DECIMAL)
					x = x & s & ";"
				else
				'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
					s = substitui_caracteres(bd_formata_moeda(r("vl_custo2_real_total_pedido")), ".", SEPARADOR_DECIMAL)
					x = x & s & ";"
					end if
				end if
		
		'> VALOR TOTAL NF
			if ckb_COL_VL_TOTAL_NF <> "" then
				if ckb_CONSOLIDAR_PEDIDO = "" then
				'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
					s = substitui_caracteres(bd_formata_moeda(CLng(s_qtde) * r("preco_NF")), ".", SEPARADOR_DECIMAL)
					x = x & s & ";"
				else
				'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
					s = substitui_caracteres(bd_formata_moeda(r("preco_NF_total_pedido")), ".", SEPARADOR_DECIMAL)
					x = x & s & ";"
					end if
				end if

		'> VALOR TOTAL
			if ckb_COL_VL_TOTAL <> "" then
				if ckb_CONSOLIDAR_PEDIDO = "" then
				'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
					s = substitui_caracteres(bd_formata_moeda(CLng(s_qtde) * r("preco_venda")), ".", SEPARADOR_DECIMAL)
					x = x & s & ";"
				else
				'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
					s = substitui_caracteres(bd_formata_moeda(r("preco_venda_total_pedido")), ".", SEPARADOR_DECIMAL)
					x = x & s & ";"
					end if
				end if

		 '> VL RA
			if ckb_COL_VL_RA <> "" then
				if ckb_CONSOLIDAR_PEDIDO = "" then
					vl_preco_venda = converte_numero(formata_moeda(r("preco_venda")))
					vl_preco_NF = converte_numero(formata_moeda(r("preco_NF")))
					'CALCULA VALOR DA RA, MANTENDO O SINAL (POSITIVO/NEGATIVO)
					vl_RA = CLng(s_qtde) * (vl_preco_NF - vl_preco_venda)
					s = substitui_caracteres(bd_formata_moeda(vl_RA), ".", SEPARADOR_DECIMAL)
					x = x & s & ";"
				else
					vl_RA = r("preco_NF_total_pedido") - r("preco_venda_total_pedido")
					s = substitui_caracteres(bd_formata_moeda(vl_RA), ".", SEPARADOR_DECIMAL)
					x = x & s & ";"
					end if
				end if

		 '> RT
			perc_RT = r("perc_RT")
			if ckb_COL_RT <> "" then
				if ckb_CONSOLIDAR_PEDIDO = "" then
				'	EVITA DIFEREN�AS DE ARREDONDAMENTO
					vl_preco_venda = converte_numero(formata_moeda(r("preco_venda")))
				'	CALCULA VL RT UNIT�RIO, MAS MANTENDO O SINAL (POSITIVO/NEGATIVO)
					vl_RT = (perc_RT/100) * (CLng(s_qtde)/Abs(CLng(s_qtde)) * vl_preco_venda)
					s = substitui_caracteres(bd_formata_moeda(vl_RT), ".", SEPARADOR_DECIMAL)
					x = x & s & ";"
				else
					vl_RT = (perc_RT/100) * r("preco_venda_total_pedido")
					s = substitui_caracteres(bd_formata_moeda(vl_RT), ".", SEPARADOR_DECIMAL)
					x = x & s & ";"
					end if
				end if
			
		'>  ICMS UF DESTINO (UNIT�RIO)
			if ckb_COL_ICMS_UF_DEST <> "" then
				if ckb_CONSOLIDAR_PEDIDO = "" then
					s_vICMSUFDest_unitario = ""
					s_vICMSUFDest = Trim("" & r("vICMSUFDest"))
					s_det__qCom = Trim("" & r("det__qCom"))
					if s_vICMSUFDest <> "" then
						vl_vICMSUFDest = converte_numero(s_vICMSUFDest)
						if s_det__qCom <> "" then
							'A quantidade est� formatada com 4 decimais: 1 unidade = 1.0000
							v = Split(s_det__qCom, ".")
							if Trim("" & v(Lbound(v))) <> "" then
								n_det__qCom = CLng(Trim("" & v(Lbound(v))))
								vl_vICMSUFDest_unitario = vl_vICMSUFDest / n_det__qCom
								s_vICMSUFDest_unitario = substitui_caracteres(bd_formata_moeda(vl_vICMSUFDest_unitario), ".", SEPARADOR_DECIMAL)
								end if
							end if
						end if
					x = x & s_vICMSUFDest_unitario & ";"
				else
					s_vICMSUFDest = Trim("" & r("vICMSUFDestTotalPedido"))
					if s_vICMSUFDest <> "" then
						s_vICMSUFDest = substitui_caracteres(bd_formata_moeda(r("vICMSUFDestTotalPedido")), ".", SEPARADOR_DECIMAL)
						end if
					x = x & s_vICMSUFDest & ";"
					end if
				end if

		 '> QTDE DE PARCELAS
			if ckb_COL_QTDE_PARCELAS <> "" then
				x = x & Trim("" & r("qtde_parcelas")) & ";"
			end if
	    
		 '> MEIO DE PAGAMENTO
			if ckb_COL_MEIO_PAGAMENTO <> "" then
				s = ""
        		tipo_parc = Trim("" & r("tipo_parcelamento"))
				if tipo_parc = COD_FORMA_PAGTO_A_VISTA then       
					 s = x_opcao_forma_pagamento(Trim("" & r("forma_pagamento_av"))) 
				elseif tipo_parc = COD_FORMA_PAGTO_PARCELADO_CARTAO then    
						 s = x_opcao_forma_pagamento(Trim(ID_FORMA_PAGTO_CARTAO))
				elseif tipo_parc = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then    
						 s = x_opcao_forma_pagamento(Trim("" & r("parcelamento_c_entrada"))) 
				elseif tipo_parc = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then    
						 s = x_opcao_forma_pagamento(Trim("" & r("parcelamento_s_entrada"))) 
				elseif tipo_parc = COD_FORMA_PAGTO_PARCELA_UNICA then     
						s = x_opcao_forma_pagamento(Trim("" & r("parcela_unica"))) 
				elseif tipo_parc = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
						 s = x_opcao_forma_pagamento(Trim(ID_FORMA_PAGTO_CARTAO_MAQUINETA))
				end if          
				x = x & s & ";"
			end if            

		'> VALOR PAGO NO CART�O (INTERNET)
			if ckb_COL_VL_PAGO_CARTAO_INTERNET <> "" then
				if ckb_CONSOLIDAR_PEDIDO <> "" then
				'	EXPORTAR VALOR UTILIZANDO SEPARADOR DECIMAL DEFINIDO
					s = ""
					if Trim("" & r("vl_pago_cartao_internet_total_pedido")) <> "" then
						s = substitui_caracteres(bd_formata_moeda(r("vl_pago_cartao_internet_total_pedido")), ".", SEPARADOR_DECIMAL)
						end if
					x = x & s & ";"
					end if
				end if

		'> CHAVE DE ACESSO NFE
			if ckb_COL_CHAVE_NFE <> "" then
				s = ""

				if Trim("" & r("operacao")) = "VENDA_NORMAL" then
					s_pesq_nf = normaliza_codigo(Trim("" & r("id_nfe_emitente")), 6) & "|" & NFeFormataNumeroNF(Trim("" & r("numero_NF")))
					if localiza_cl_dez_colunas(vNFeChave, s_pesq_nf, idxNFeLocalizada) then
						'FOR�A PARA O EXCEL TRATAR COMO TEXTO E N�O N�MERO
						if ckb_COMPATIBILIDADE <> "" then
							s = chr(34) & "=" & chr(34) & chr(34) & vNFeChave(idxNFeLocalizada).c3 & chr(34) & chr(34) & chr(34)
						else
							s = "=" & chr(34) & vNFeChave(idxNFeLocalizada).c3 & chr(34)
							end if
						end if

'					for i=LBound(vNFeChave) to UBound(vNFeChave)
'						if (vNFeChave(i).c1 = Trim("" & r("id_nfe_emitente"))) And (vNFeChave(i).c2 = NFeFormataNumeroNF(Trim("" & r("numero_NF")))) then
'							'FOR�A PARA O EXCEL TRATAR COMO TEXTO E N�O N�MERO
'							s = "=" & chr(34) & vNFeChave(i).c3 & chr(34)
'							exit for
'							end if
'						next
					end if

				x = x & s & ";"
				end if
		
			x = x & vbCrLf
		
			if (n_reg_total mod 100) = 0 then
				Response.Write x
				x = ""
				end if

        next 'for iQI=1 to Abs(item_qtde)
		
		r.MoveNext
		loop
		

	if r.State <> 0 then r.Close
	set r=nothing
	
'	MOSTRA AVISO DE QUE N�O H� DADOS!!
	if n_reg_total = 0 then
		x = "NENHUM PRODUTO ENCONTRADO"
	elseif (n_reg_total <> n_reg_total_passo1) And (n_reg_total_passo1 <> -1) then
		x = "OCORREU UMA INCONSIST�NCIA NA PROCESSAMENTO DO RELAT�RIO, FAVOR EXECUTAR NOVAMENTE!"
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
<input type="hidden" name="c_dt_NF_venda_inicio" id="c_dt_NF_venda_inicio" value="<%=c_dt_NF_venda_inicio%>">
<input type="hidden" name="c_dt_NF_venda_termino" id="c_dt_NF_venda_termino" value="<%=c_dt_NF_venda_termino%>">
<input type="hidden" name="c_dt_NF_remessa_inicio" id="c_dt_NF_remessa_inicio" value="<%=c_dt_NF_remessa_inicio%>">
<input type="hidden" name="c_dt_NF_remessa_termino" id="c_dt_NF_remessa_termino" value="<%=c_dt_NF_remessa_termino%>">
<input type="hidden" name="ckb_pedido_nao_recebido_pelo_cliente" id="ckb_pedido_nao_recebido_pelo_cliente" value="<%=ckb_pedido_nao_recebido_pelo_cliente%>" />
<input type="hidden" name="ckb_pedido_recebido_pelo_cliente" id="ckb_pedido_recebido_pelo_cliente" value="<%=ckb_pedido_recebido_pelo_cliente%>" />
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_grupo" id="c_grupo" value="<%=c_grupo%>">
<input type="hidden" name="c_potencia_BTU" id="c_potencia_BTU" value="<%=c_potencia_BTU%>">
<input type="hidden" name="c_ciclo" id="c_ciclo" value="<%=c_ciclo%>">
<input type="hidden" name="c_posicao_mercado" id="c_posicao_mercado" value="<%=c_posicao_mercado%>">
<input type="hidden" name="c_entrega_imediata" id="c_entrega_imediata" value="<%=c_entrega_imediata%>">
<input type="hidden" name="rb_tipo_cliente" id="rb_tipo_cliente" value="<%=rb_tipo_cliente%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">
<input type="hidden" name="ckb_CONSOLIDAR_PEDIDO" id="ckb_CONSOLIDAR_PEDIDO" value="<%=ckb_CONSOLIDAR_PEDIDO%>" />


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
	
'	PER�ODO (ENTREGUE)
	s = ""
	s_aux = c_dt_faturamento_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_faturamento_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Per�odo (Entregue):&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
'	PER�ODO NF VENDA
	s = ""
	s_aux = c_dt_NF_venda_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_NF_venda_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Per�odo NF Venda:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

'	PER�ODO NF REMESSA
	s = ""
	s_aux = c_dt_NF_remessa_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_NF_remessa_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Per�odo NF Remessa:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

'	PEDIDOS RECEBIDOS PELO CLIENTE
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
					"		<td align='right' valign='top' nowrap><span class='N'>Pedidos Recebidos pelo Cliente:&nbsp;</span></td>" & chr(13) & _
					"		<td valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

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

'	ENTREGA IMEDIATA
	s = c_entrega_imediata
	if s = "" then
		s = "N.I."
	else
		s = decodifica_etg_imediata(c_entrega_imediata)
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Entrega Imediata:&nbsp;</span></td>" & chr(13) & _
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

	if ckb_CONSOLIDAR_PEDIDO <> "" then
		s = "Sim"
	else
		s = "N�o"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Consolidado por pedido:&nbsp;</span></td><td align='left' valign='top'>" & _
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
	if r.State <> 0 then r.Close
	set r = nothing

	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
