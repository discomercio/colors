<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelSeparacaoZona.asp
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

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_EXTENDED_TIMEOUT_EM_SEG
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim dt_emissao, dt_hr_emissao
	dt_emissao = Date
	dt_hr_emissao = Now
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_SEPARACAO_ZONA, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta, s, s_aux, s_table_produtos_sem_zona, s_filtro, intSequencia
	dim c_dt_inicio, c_dt_termino, rb_nfe, c_transportadora, c_qtde_max_pedidos, c_lista_pedidos_selecionados, c_qtde_total_pedidos_disponiveis, c_nfe_emitente
	dim vPedidos
	alerta = ""
	s_table_produtos_sem_zona = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_transportadora = Trim(Request.Form("c_transportadora"))
	rb_nfe = Trim(Request.Form("rb_nfe"))
	c_qtde_max_pedidos = retorna_so_digitos(Request.Form("c_qtde_max_pedidos"))
	c_lista_pedidos_selecionados = Trim(Request.Form("c_lista_pedidos_selecionados"))
	c_qtde_total_pedidos_disponiveis = Trim(Request.Form("c_qtde_total_pedidos_disponiveis"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))

	vPedidos = Split(c_lista_pedidos_selecionados, "|", -1)
	
	if (Trim("" & vPedidos(Lbound(vPedidos))) <> "INICIO") Or (Trim("" & vPedidos(Ubound(vPedidos))) <> "TERMINO") then
		alerta="RELAÇÃO DOS PEDIDOS SELECIONADOS ESTÁ EM FORMATO INVÁLIDO!"
		end if
	
	dim i, s_sql_lista_pedidos, s_pedido, n_pedidos_selecionados
	s_sql_lista_pedidos = ""
	n_pedidos_selecionados = 0
	if alerta = "" then
		for i=LBound(vPedidos) to UBound(vPedidos)
			s_pedido = Trim(vPedidos(i))
			if (s_pedido <> "") And (s_pedido <> "INICIO") And (s_pedido <> "TERMINO") then
				n_pedidos_selecionados = n_pedidos_selecionados + 1
				if s_sql_lista_pedidos <> "" then s_sql_lista_pedidos = s_sql_lista_pedidos & ","
				s_sql_lista_pedidos = s_sql_lista_pedidos & "'" & s_pedido & "'"
				end if
			next
		
		if n_pedidos_selecionados = 0 then
			alerta = "A LISTA DE PEDIDOS SELECIONADOS ESTÁ VAZIA!"
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

	if alerta = "" then
	'	CONSISTÊNCIA: SE ALGUM PRODUTO APTO A APARECER NESTE RELATÓRIO NÃO POSSUIR
	'	============  A ZONA CADASTRADA, EXIBE UM AVISO E IMPEDE A EXIBIÇÃO DO RELATÓRIO.
	'	ESTA LÓGICA VISA PREVENIR O RISCO DE FALHAS!
	'	LEMBRANDO QUE UM PEDIDO É "MONTADO" A PARTIR DA CONSOLIDAÇÃO DE PRODUTOS
	'	SEPARADOS E TRAZIDOS DE VÁRIAS ZONAS DO DEPÓSITO.
		s = "SELECT DISTINCT" & _
				" t_PRODUTO.fabricante," & _
				" t_PRODUTO.produto," & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html," & _
				" zona_codigo" & _
			" FROM t_PEDIDO" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
				" LEFT JOIN t_PRODUTO ON ((t_PEDIDO_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_PEDIDO_ITEM.produto=t_PRODUTO.produto))" & _
				" LEFT JOIN t_WMS_DEPOSITO_MAP_ZONA ON (t_PRODUTO.deposito_zona_id=t_WMS_DEPOSITO_MAP_ZONA.id)" & _
			" WHERE" & _
				" (t_PEDIDO.st_entrega='" & ST_ENTREGA_SEPARAR & "')" & _
				" AND (t_PEDIDO.a_entregar_data_marcada IS NOT NULL)" & _
				" AND (t_PEDIDO.st_etg_imediata = " & COD_ETG_IMEDIATA_SIM & ")" & _
				" AND (t_PEDIDO__BASE.analise_credito = " & COD_AN_CREDITO_OK & ")"
		
		s = "SELECT " & _
				"*" & _
			" FROM (" & s & ") t" & _
			" WHERE" & _
				" (zona_codigo IS NULL)" & _
			" ORDER BY" & _
				" fabricante," & _
				" produto"
		
		intSequencia=0
		set rs=cn.Execute(s)
		do while Not rs.Eof
			intSequencia = intSequencia + 1
			s_table_produtos_sem_zona = s_table_produtos_sem_zona & _
										"	<tr>" & chr(13) & _
										"		<td align='left'><span class='Rd' style='margin-right:2px;'>" & Cstr(intSequencia) & ".</span></td>" & chr(13) & _
										"		<td class='MB ME MD' align='left'><span class='C'>" & Trim("" & rs("fabricante")) & "</span></td>" & chr(13) & _
										"		<td class='MB MD' align='left'><span class='C'>" & Trim("" & rs("produto")) & "</span></td>" & chr(13) & _
										"		<td class='MB MD' align='left'><span class='C'>" & produto_formata_descricao_em_html(Trim("" & rs("descricao_html"))) & "</span></td>" & chr(13) & _
										"	</tr>" & chr(13)
			rs.MoveNext
			loop
		
		if s_table_produtos_sem_zona <> "" then
			s_table_produtos_sem_zona = "<table cellspacing='0' cellpadding='0'>" & chr(13) & _
										"	<tr>" & chr(13) &_
										"		<td align='left'>&nbsp;</td>" & chr(13) & _
										"		<td class='MT' style='width:40px;' align='left'><span class='R'>FABR</span></td>" & chr(13) & _
										"		<td class='MTBD' style='width:60px;' align='left'><span class='R'>PRODUTO</span></td>" & chr(13) & _
										"		<td class='MTBD' align='left'><span class='R'>DESCRIÇÃO</span></td>" & chr(13) & _
										"	</tr>" & chr(13) & _
										s_table_produtos_sem_zona & _
										"</table>" & chr(13)
			alerta = "Os seguintes produtos NÃO possuem a zona cadastrada:<br>"
			end if
		end if
	
	if alerta = "" then
		if c_nfe_emitente = "" then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Não foi informado o CD"
		elseif converte_numero(c_nfe_emitente) = 0 then
			alerta=texto_add_br(alerta)
			alerta = alerta & "É necessário definir um CD válido"
			end if
		end if
	
	dim strScriptJS
	strScriptJS = ""
	
	dim lngNsuWmsEtqN1, lngNsuWmsEtqN2, lngNsuWmsEtqN3, msg_erro





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA RELATORIO
' Subsídios: o depósito está dividido em zonas, cada qual com uma equipe de separadores.
' ========== Cada equipe de separadores realiza o "picking" de produtos e os leva até uma
' área de consolidação. A equipe de consolidação irá reunir os itens que compõem o pedido
' e levá-los até o box da transportadora adequada.
' Inicialmente, o filtro deste relatório possuía um campo destinado a restringir por zona,
' para que o relatório fosse executado uma vez p/ cada uma das zonas e mais uma consulta
' abrangendo todas as zonas p/ o consolidador.
' O problema dessa abordagem é que os pedidos podem ser alterados de modo que passem a
' constar no relatório durante esse processo de impressão, causando inconsistência entre os
' relatórios de cada uma das zonas com o relatório do consolidador.
' A solução implementada é realizar uma única consulta ao banco de dados e, a partir dela,
' gerar todos os relatórios de uma vez.
sub consulta_relatorio
const COD_ZONA_ID__TODOS = 999999
const COD_ZONA_CODIGO__TODOS = "***TODOS***"
dim r, t
dim s, s_aux, s_erro_fatal, s_erro_html, s_zona, s_sql, s_where, pedido_a, fabricante_a, s_numero_NF, s_transportadora
dim i, iZona, iRel, n_qtde_zona, n_reg_total, idx
dim vProd(), vZona(), vRel()
dim v
dim x, xRel
dim s_linha_branco, s_tit_zona, s_lista_zonas_cadastradas
dim blnProcessaRegistro
dim strJS_AllTablesCollapse, strJS_AllTablesNotPrint
dim s_log, s_log_filtro
dim lngRecordsAffected, intSequenciaN2, intSequenciaN3
dim rNfeEmitente

'	VETOR QUE ARMAZENA TODOS OS REGISTROS (A ORDENAÇÃO DEVE SER FEITA NA CONSULTA SQL)
	redim vRel(0)
	set vRel(0) = New cl_REL_SEPARACAO_ZONA
	vRel(0).produto = ""

'	ARMAZENA A LISTA DE ZONAS CADASTRADAS
	redim vZona(0)
	set vZona(0) = New cl_ZONA_DEPOSITO
	vZona(0).zona_id = 0
	
	s_sql = "SELECT" & _
				" id," & _
				" zona_codigo" & _
			" FROM t_WMS_DEPOSITO_MAP_ZONA" & _
			" WHERE" & _
				" (st_ativo <> 0)" & _
			" ORDER BY" &_
				" zona_codigo"
	set r = cn.Execute(s_sql)
	do while Not r.Eof
		if vZona(Ubound(vZona)).zona_id <> 0 then
			redim preserve vZona(Ubound(vZona)+1)
			set vZona(Ubound(vZona)) = New cl_ZONA_DEPOSITO
			end if
		
		with vZona(Ubound(vZona))
			.zona_id = CLng(r("id"))
			.zona_codigo = Trim("" & r("zona_codigo"))
			end with
		
		r.MoveNext
		loop

	if r.State <> 0 then r.Close
	set r=nothing

'	ADICIONA O ITEM P/ GERAR O RELATÓRIO DE TODAS AS ZONAS
	redim preserve vZona(Ubound(vZona)+1)
	set vZona(Ubound(vZona)) = New cl_ZONA_DEPOSITO
	with vZona(Ubound(vZona))
		.zona_id = COD_ZONA_ID__TODOS
		.zona_codigo = COD_ZONA_CODIGO__TODOS
		end with

	s_sql = "SELECT" & _
				" t_PEDIDO.pedido," & _
				" t_PEDIDO.data," & _
				" t_PEDIDO.hora," & _
				" t_PEDIDO.obs_2," & _
				" t_PEDIDO.obs_3," & _
				" t_PEDIDO.loja," & _
				" t_PEDIDO.transportadora_id," & _
				" t_PEDIDO.id_cliente," & _
				" t_PEDIDO.st_end_entrega," & _
				" t_PEDIDO.EndEtg_endereco," & _
				" t_PEDIDO.EndEtg_endereco_numero," & _
				" t_PEDIDO.EndEtg_endereco_complemento," & _
				" t_PEDIDO.EndEtg_bairro," & _
				" t_PEDIDO.EndEtg_cidade," & _
				" t_PEDIDO.EndEtg_uf," & _
				" t_PEDIDO.EndEtg_cep," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas AS nome_cliente," & _
				" t_CLIENTE.endereco AS cliente_endereco," & _
				" t_CLIENTE.endereco_numero AS cliente_endereco_numero," & _
				" t_CLIENTE.endereco_complemento AS cliente_endereco_complemento," & _
				" t_CLIENTE.bairro AS cliente_bairro," & _
				" t_CLIENTE.cidade AS cliente_cidade," & _
				" t_CLIENTE.uf AS cliente_uf," & _
				" t_CLIENTE.cep AS cliente_cep," & _
				" t_FABRICANTE.nome AS nome_fabricante," & _
				" t_FABRICANTE.razao_social AS razao_social_fabricante," & _
				" t_PEDIDO_ITEM.fabricante," & _
				" t_PEDIDO_ITEM.produto," & _
				" t_PEDIDO_ITEM.qtde," & _
				" t_PEDIDO_ITEM.qtde_volumes," & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html," & _
				" t_PRODUTO.deposito_zona_id AS zona_id," & _
				" t_WMS_DEPOSITO_MAP_ZONA.zona_codigo," & _
				" t_PEDIDO.num_obs_2 AS numNFeFaturamento," & _
				" t_PEDIDO.num_obs_3 AS numNFeRemessa," & _
				" (" & _
					"SELECT " & _
						" Sum(qtde*qtde_volumes)" & _
					" FROM t_PEDIDO_ITEM" & _
					" WHERE" & _
						" (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
				") AS qtde_volumes_pedido" & _
			" FROM t_PEDIDO" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
				" LEFT JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
				" LEFT JOIN t_PRODUTO ON ((t_PEDIDO_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_PEDIDO_ITEM.produto=t_PRODUTO.produto))" & _
				" LEFT JOIN t_FABRICANTE ON (t_PRODUTO.fabricante=t_FABRICANTE.fabricante)" & _
				" LEFT JOIN t_WMS_DEPOSITO_MAP_ZONA ON (t_PRODUTO.deposito_zona_id=t_WMS_DEPOSITO_MAP_ZONA.id)" & _
			" WHERE" & _
				" (t_PEDIDO.st_entrega='" & ST_ENTREGA_SEPARAR & "')" & _
				" AND (t_PEDIDO.a_entregar_data_marcada IS NOT NULL)" & _
				" AND (t_PEDIDO.st_etg_imediata = " & COD_ETG_IMEDIATA_SIM & ")" & _
				" AND (t_PEDIDO__BASE.analise_credito = " & COD_AN_CREDITO_OK & ")"
	
	if IsDate(c_dt_inicio) then
		s_sql = s_sql & " AND (t_PEDIDO.a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		s_sql = s_sql & " AND (t_PEDIDO.a_entregar_data_marcada < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if

	if c_transportadora <> "" then
		s_sql = s_sql & " AND (t_PEDIDO.transportadora_id = '" & c_transportadora & "')"
		end if
	
'	OWNER DO PEDIDO
	set rNfeEmitente = le_nfe_emitente(c_nfe_emitente)
	s_sql = s_sql & " AND (t_PEDIDO.id_nfe_emitente = " & rNfeEmitente.id & ")"
	
	if s_sql_lista_pedidos <> "" then
		s_sql = s_sql & " AND (t_PEDIDO.pedido IN (" & s_sql_lista_pedidos & "))"
		end if
	
	s_sql = "SELECT " & _
				"*" & _
			" FROM " & _
				"(" & s_sql & ") t"
	
'	NFe EMITIDA?
	s_where = ""
	if rb_nfe = "EMITIDA" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " ((numNFeFaturamento <> 0) OR (numNFeRemessa <> 0))"
	elseif rb_nfe = "NAO_EMITIDA" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " ((numNFeFaturamento = 0) AND (numNFeRemessa = 0))"
		end if
	
	if s_where <> "" then s_where = " WHERE" & s_where
	s_sql = s_sql & s_where
	s_sql = s_sql & " ORDER BY transportadora_id, data, hora, pedido, fabricante, produto"

	set r = cn.execute(s_sql)

'	ARMAZENA TODOS OS REGISTROS NO VETOR
	do while Not r.Eof
		if Trim("" & vRel(Ubound(vRel)).produto) <> "" then
			redim preserve vRel(Ubound(vRel)+1)
			set vRel(Ubound(vRel)) = New cl_REL_SEPARACAO_ZONA
			end if
		
		with vRel(Ubound(vRel))
			.pedido = Trim("" & r("pedido"))
			.obs_2 = Trim("" & r("obs_2"))
			.obs_3 = Trim("" & r("obs_3"))
			.loja = Trim("" & r("loja"))
			.id_cliente = Trim("" & r("id_cliente"))
			.nome_cliente = Trim("" & r("nome_cliente"))
			.transportadora_id = Trim("" & r("transportadora_id"))
			.fabricante = Trim("" & r("fabricante"))
			.produto = Trim("" & r("produto"))
			.qtde = r("qtde")
			.qtde_volumes = r("qtde_volumes")
			.descricao = Trim("" & r("descricao"))
			.descricao_html = Trim("" & r("descricao_html"))
			.zona_id = r("zona_id")
			.zona_codigo = Trim("" & r("zona_codigo"))
			'Somente NF de Remessa, quando houver
			if r("numNFeRemessa") > 0 then
				.numeroNFe = Trim("" & r("numNFeRemessa"))
			elseif r("numNFeFaturamento") > 0 then
				.numeroNFe = Trim("" & r("numNFeFaturamento"))
			else
				.numeroNFe = ""
				end if
			.qtde_volumes_pedido = r("qtde_volumes_pedido")
			.nome_fabricante = Trim("" & r("nome_fabricante"))
			if .nome_fabricante = "" then .nome_fabricante = Trim("" & r("razao_social_fabricante"))
			if Trim("" & r("st_end_entrega")) <> "0" then
				.destino_tipo_endereco = COD_WMS_ENDERECO_DESTINO__END_ENTREGA
				.destino_endereco = Trim("" & r("EndEtg_endereco"))
				.destino_endereco_numero = Trim("" & r("EndEtg_endereco_numero"))
				.destino_endereco_complemento = Trim("" & r("EndEtg_endereco_complemento"))
				.destino_bairro = Trim("" & r("EndEtg_bairro"))
				.destino_cidade = Trim("" & r("EndEtg_cidade"))
				.destino_uf = Ucase(Trim("" & r("EndEtg_uf")))
				.destino_cep = Trim("" & r("EndEtg_cep"))
			else
				.destino_tipo_endereco = COD_WMS_ENDERECO_DESTINO__CAD_CLIENTE
				.destino_endereco = Trim("" & r("cliente_endereco"))
				.destino_endereco_numero = Trim("" & r("cliente_endereco_numero"))
				.destino_endereco_complemento = Trim("" & r("cliente_endereco_complemento"))
				.destino_bairro = Trim("" & r("cliente_bairro"))
				.destino_cidade = Trim("" & r("cliente_cidade"))
				.destino_uf = Ucase(Trim("" & r("cliente_uf")))
				.destino_cep = Trim("" & r("cliente_cep"))
				end if
			end with
		
		r.MoveNext
		loop

	if r.State <> 0 then r.Close
	set r=nothing

	s_lista_zonas_cadastradas = ""
	n_qtde_zona = 0
	xRel = ""
	x = ""
	strJS_AllTablesCollapse = ""
	strJS_AllTablesNotPrint = ""

'	GERA O RELATÓRIO PARA CADA UMA DAS ZONAS
	for iZona=Lbound(vZona) to Ubound(vZona)
		if converte_numero(vZona(iZona).zona_id) <> 0 then
		
		'	AS ROTINAS DE ORDENAÇÃO USAM VETORES QUE SE INICIAM NA POSIÇÃO 1
			redim vProd(1)
			for i = Lbound(vProd) to Ubound(vProd)
				set vProd(i) = New cl_CINCO_COLUNAS
				with vProd(i)
					.c1 = ""
					.c2 = 0
					.c3 = ""
					.c4 = ""
					.c5 = ""
					end with
				next
			
			n_reg_total = 0
			pedido_a = "XXXXXXXXXX"
			
			n_qtde_zona = n_qtde_zona + 1
			if s_lista_zonas_cadastradas <> "" then s_lista_zonas_cadastradas = s_lista_zonas_cadastradas & "|"
			s_lista_zonas_cadastradas = s_lista_zonas_cadastradas & vZona(iZona).zona_id & "§" & vZona(iZona).zona_codigo
			
			for iRel=Lbound(vRel) to Ubound(vRel)
				blnProcessaRegistro = False
				if Trim("" & vRel(iRel).produto) <> "" then
					if Trim("" & vRel(iRel).zona_id) = Trim("" & vZona(iZona).zona_id) then blnProcessaRegistro = True
					if vZona(iZona).zona_id = COD_ZONA_ID__TODOS then blnProcessaRegistro = True
					end if
				
				if blnProcessaRegistro then
					with vRel(iRel)
						if Trim("" & .pedido) <> pedido_a then
							pedido_a = Trim("" & .pedido)
						'	FECHA TABELA DO PEDIDO ANTERIOR
							if n_reg_total > 0 then
								x = x & "		</table>" & chr(13) & _
										"	</td></tr>" & chr(13) & _
										"</table>" & chr(13) & _
										"<br>" & chr(13)
								end if

						'	Nº DA NF
							s_numero_NF = Trim("" & .obs_2)
							if Trim("" & .obs_3) <> "" then s_numero_NF = Trim("" & .obs_3)
							if s_numero_NF = "" then s_numero_NF = "&nbsp;"
							
						'	TRANSPORTADORA
							s_transportadora = iniciais_em_maiusculas(Trim("" & .transportadora_id))
							if s_transportadora = "" then s_transportadora = "&nbsp;"
							
						'	TABELA P/ O PRÓXIMO PEDIDO
							x = x & chr(13) & _
								"<table cellspacing='0' cellpadding='0'>" & chr(13) & _
								"	<tr><td align='left'>" & chr(13) & _
								"		<table class='Q' cellspacing=0 cellpadding=0>" & chr(13) & _
								"			<tr style='background:#FFF0E0' nowrap>" & chr(13)& _
								"				<td class='tdPedido' align='left' nowrap><span class='C'><a href='javascript:fRELConcluir(" & _
												chr(34) & Trim("" & .pedido) & chr(34) & ")' title='clique para consultar o pedido'>" & _
													Trim("" & .pedido) & "</a></span></td>" & chr(13) & _
								"				<td class='tdQtVol' align='left' nowrap><span class='C'>" & _
													Trim("" & .qtde_volumes_pedido) & "</span></td>" & chr(13) & _
								"				<td class='tdNF' align='left' nowrap><span class='C'>" & _
													s_numero_NF & "</span></td>" & chr(13) & _
								"				<td class='tdLoja' align='left' nowrap><span class='C'>Lj&nbsp;" & _
													Trim("" & .loja) & "</span></td>" & chr(13) & _
								"				<td class='tdCliente' align='left'><span class='C'>" & _
													Trim("" & .nome_cliente) & "</span></td>" & chr(13) & _
								"				<td class='tdTransp' align='left' nowrap><span class='C'>" & _
													s_transportadora & "</span></td>" & chr(13) & _
								"			</tr>" & chr(13) & _
								"		</table>" & chr(13) & _
								"	</td></tr>" & chr(13) & _
								"	<tr><td align='left'>" & chr(13) & _
								"		<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13)
							end if

					  ' CONTAGEM
						n_reg_total = n_reg_total + 1
						
					'	LISTAGEM
						x = x & "			<tr nowrap>" & chr(13)

					 '> QTDE
						x = x & "				<td class='MDBE tdQtde' align='left'><span class='Cd'>&nbsp;" & _
							formata_inteiro(.qtde) & "</span></td>"  & chr(13)

 					 '> PRODUTO
						x = x & "				<td class='MDB tdProd' align='left' nowrap><span class='C' nowrap>&nbsp;" & _
							produto_formata_descricao_em_html(Trim("" & .descricao_html)) & _
							"&nbsp;&nbsp;(Cód:&nbsp;" & Trim("" & .produto) & _
							"&nbsp;&nbsp;Fabr:&nbsp;" & Trim("" & .fabricante) & ")" & _
							"</span></td>" & chr(13)

 					 '> ZONA
 						s = Trim("" & .zona_codigo)
 						if s = "" then s = "&nbsp;"
 						x = x & "				<td class='MDB tdZona' align='left' nowrap><span class='Cc' nowrap>" & s & "</span></td>" & chr(13)

						x = x & "			</tr>" & chr(13)

					'	TOTALIZAÇÃO
						s = Trim("" & .fabricante) & "|" & Trim("" & .produto)
						if localiza_cl_cinco_colunas(vProd, s, idx) then
							with vProd(idx)
								.c2 = .c2 + CLng(vRel(iRel).qtde)
								end with
						else
							if (vProd(Ubound(vProd)).c1<>"") then
								redim preserve vProd(Ubound(vProd)+1)
								set vProd(Ubound(vProd)) = New cl_CINCO_COLUNAS
								end if
							with vProd(Ubound(vProd))
								.c1 = Trim("" & vRel(iRel).fabricante) & "|" & Trim("" & vRel(iRel).produto)
								.c2 = CLng(vRel(iRel).qtde)
								.c3 = Trim("" & vRel(iRel).zona_codigo)
								.c4 = Trim("" & vRel(iRel).descricao_html)
								.c5 = Trim("" & vRel(iRel).nome_fabricante)
								end with
							ordena_cl_cinco_colunas vProd, 1, Ubound(vProd)
							end if

						end with
					end if ' if blnProcessaRegistro
				next
			
		'	FINALIZAÇÃO
			if n_reg_total <> 0 then 
			'	FECHA ÚLTIMA TABELA
				x = x & "		</table>" & chr(13) & _
						"	</td></tr>" & chr(13) & _
						"</table>" & chr(13)
			'	TOTAIS
				x = x & chr(13) & "<br><br>" & chr(13) & _
					"<table class='Q' style='border-bottom:0px;' cellspacing=0 cellpadding=0>" & chr(13) & _
					"	<tr style='background:azure' nowrap>" & chr(13) & _
					"		<td class='MB' align='center' colspan='4'><span class='C'>&nbsp;TOTAL</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
			'	LEMBRANDO QUE O VETOR ESTÁ ORDENADO
				fabricante_a = "XXXXX"
				for i = Lbound(vProd) to Ubound(vProd)
					with vProd(i)
						if Trim("" & .c1) <> "" then
							v = Split(.c1, "|", -1)
							if Trim("" & v(0)) <> fabricante_a then
								fabricante_a = Trim("" & v(0))
								s = Trim("" & v(0))
								s_aux = ucase(.c5)
								if (s<>"") And (s_aux<>"") then s = s & " - "
								s = s & s_aux
								x = x & "	<tr nowrap>" & chr(13) & _
									"		<td colspan='4' class='MB' align='left'><span class='Cc'>" & _
									s & _
									"</span>" & chr(13) & _
									"	</tr>" & chr(13)
								end if
							
							s_zona = Trim(.c3)
							if s_zona = "" then s_zona = "&nbsp;"
							
							x = x & "	<tr nowrap>" & chr(13) & _
								"		<td class='MDB tdTotCodProd' align='left'><span class='C'>&nbsp;" & _
								Trim("" & v(1)) & "</span></td>"  & chr(13) & _
								"		<td class='MDB tdTotDescrProd' align='left' nowrap><span class='C' nowrap>&nbsp;" & _
								produto_formata_descricao_em_html(.c4) & _
								"</span></td>" & chr(13) & _
								"		<td class='MDB tdTotQtde' align='left'><span class='Cd'>&nbsp;" & _
								formata_inteiro(.c2) & "</span></td>"  & chr(13) & _
								"		<td class='MB tdTotZona' align='left'><span class='Cc'>" & s_zona & "</span></td>" & chr(13) & _
								"	</tr>" & chr(13)
							end if
						end with
					next
				x = x & "</table>"
				end if

		  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
			if n_reg_total = 0 then
				x = x & _
					"<table class='Q' width='100%' cellspacing='0'>" & chr(13) & _
					"	<tr>" & chr(13) & _
					"		<td class='ALERTA' align='center'><span class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS PARA SEPARAR.&nbsp;</span></td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"</table>"
				end if
			
			s_linha_branco = ""
			if n_qtde_zona > 1 then
				s_linha_branco = "	<tr class='notPrint'><td colspan='2' align='left'>&nbsp;</td></tr>" & chr(13) & _
								 "	<tr class='notPrint'><td colspan='2' align='left'>&nbsp;</td></tr>" & chr(13)
				end if
			
			if vZona(iZona).zona_id = COD_ZONA_ID__TODOS then
				s_tit_zona = "Todas"
			else
				s_tit_zona = vZona(iZona).zona_codigo
				end if
			
			x = s_linha_branco & _
				"	<tr id='trTitZona_" & vZona(iZona).zona_id & "' class='notPrint'>" & chr(13) & _
				"		<td nowrap style='width:120px;' align='left'><span class='STP'>Zona: " & s_tit_zona & "</span></td>" & chr(13) & _
				"		<td align='left' class='notPrint' width='90%'><a name='bImprime' href='javascript:fImprime(" & vZona(iZona).zona_id & ")' title='Imprime o relatório da zona: " & s_tit_zona & "' style='margin:3px;'><img name='imgPrinter' src='../botao/Printer.png' border='0'></a>" & _
							"&nbsp;<a name='bExibe' href='javascript:fExibeRelZona(" & vZona(iZona).zona_id & ")' title='Exibe o relatório da zona: " & s_tit_zona & "' style='margin:3px;'><img name='imgView' src='../botao/view_bottom.png' border='0'></a></td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr id='trRelZona_" & vZona(iZona).zona_id & "' class='notPrint'>" & chr(13) & _
				"		<td colspan='2' align='center'>" & chr(13) & _
						x & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13)

			xRel = xRel & x
			x = ""
			
			strJS_AllTablesCollapse = strJS_AllTablesCollapse & _
									  "	oTrRel = document.getElementById('trRelZona_" & vZona(iZona).zona_id & "');" & chr(13) & _
									  "	oTrRel.style.display = 'none';" & chr(13)
			strJS_AllTablesNotPrint = strJS_AllTablesNotPrint & _
									  "	oTrRel = document.getElementById('trRelZona_" & vZona(iZona).zona_id & "');" & chr(13) & _
									  "	oTrRel.className = 'notPrint';" & chr(13) & _
									  "	oTrTitZona = document.getElementById('trTitZona_" & vZona(iZona).zona_id & "');" & chr(13) & _
									  "	oTrTitZona.className = 'notPrint';" & chr(13)
			end if ' if converte_numero(vZona(iZona).zona_id) <> 0
		next


'	LOG
	s_log = ""

'	LOG DOS FILTROS UTILIZADOS NA CONSULTA
	s_log_filtro = ""
	
'	PERÍODO
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	if s_log_filtro <> "" then s_log_filtro = s_log_filtro & "; "
	s_log_filtro = s_log_filtro & "Período: " & s

'	NFe EMITIDA
	s = ""
	s_aux = rb_nfe
	if s_aux = "EMITIDA" then
		s = "Somente com NFe já emitida"
	elseif s_aux = "NAO_EMITIDA" then
		s = "Somente com NFe não emitida"
	else
		s = "Ambos"
		end if
	if s_log_filtro <> "" then s_log_filtro = s_log_filtro & "; "
	s_log_filtro = s_log_filtro & "NFe emitida: " & s

'	TRANSPORTADORA
	s = c_transportadora
	if s = "" then s = "N.I."
	if s_log_filtro <> "" then s_log_filtro = s_log_filtro & "; "
	s_log_filtro = s_log_filtro & "Transportadora: " & s

'	QTDE MÁXIMA DE PEDIDOS
	s = c_qtde_max_pedidos
	if s = "" then s = "N.I."
	if s_log_filtro <> "" then s_log_filtro = s_log_filtro & "; "
	s_log_filtro = s_log_filtro & "Qtde máxima de pedidos: " & s & " (total de pedidos disponíveis de acordo com os critérios de seleção: " & c_qtde_total_pedidos_disponiveis & ")"
	
	if s_log <> "" then s_log = s_log & ";  "
	s_log = s_log & "Filtros = (" & s_log_filtro & ")"
	

'	ANOTA NOS ITENS DE PEDIDO AS INFORMAÇÕES REFERENTES A ESTE RELATÓRIO
	s_erro_fatal = ""
'	~~~~~~~~~~~~~
	cn.BeginTrans
'	~~~~~~~~~~~~~
	if Not fin_gera_nsu(T_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO, lngNsuWmsEtqN1, msg_erro) then
		s_erro_fatal = "FALHA AO GERAR NSU PARA IDENTIFICAÇÃO DESTA EXECUÇÃO DO RELATÓRIO (" & msg_erro & ")"
		end if

	if s_erro_fatal = "" then
		if Not cria_recordset_pessimista(t, msg_erro) then
			s_erro_fatal = "FALHA AO TENTAR CRIAR UM OBJETO ADO PARA ACESSO AO BANCO DE DADOS."
			end if
		end if
	
	if s_erro_fatal = "" then
		s = "SELECT * FROM t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO WHERE (id = -1)"
		t.open s, cn
		t.AddNew
		t("id") = lngNsuWmsEtqN1
		t("usuario") = usuario
		t("dt_emissao") = dt_emissao
		t("dt_hr_emissao") = dt_hr_emissao
		t("filtro_dt_inicio") = c_dt_inicio
		t("filtro_dt_termino") = c_dt_termino
		t("filtro_NFe_emitida") = rb_nfe
		t("filtro_transportadora") = c_transportadora
		t("filtro_qtde_max_pedidos") = c_qtde_max_pedidos
		t("filtro_qtde_disponivel_pedidos") = c_qtde_total_pedidos_disponiveis
		t("lista_zonas_cadastradas") = s_lista_zonas_cadastradas
		t.Update
		end if
	
	if s_erro_fatal = "" then
		intSequenciaN2 = 0
		intSequenciaN3 = 0
		pedido_a = "XXXXXXXXXX"
		for iRel=Lbound(vRel) to Ubound(vRel)
			if Trim(vRel(iRel).produto) <> "" then
			'	MUDOU PEDIDO?
				if Trim(vRel(iRel).pedido) = pedido_a then
					if s_log <> "" then s_log = s_log & ", "
				else
					pedido_a = Trim(vRel(iRel).pedido)
					if s_log <> "" then s_log = s_log & ";  "
					s_log = s_log & Trim(vRel(iRel).pedido) & " = "
					
					if Not fin_gera_nsu(T_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO, lngNsuWmsEtqN2, msg_erro) then
						s_erro_fatal = "FALHA AO GERAR NSU PARA IDENTIFICAÇÃO DESTA EXECUÇÃO DO RELATÓRIO (" & msg_erro & ")"
						exit for
						end if
					
					intSequenciaN2 = intSequenciaN2 + 1
					s = "SELECT * FROM t_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO WHERE (id = -1)"
					if t.State <> 0 then t.Close
					t.open s, cn
					t.AddNew
					t("id") = lngNsuWmsEtqN2
					t("id_wms_etq_n1") = lngNsuWmsEtqN1
					t("sequencia") = intSequenciaN2
					t("pedido") = Trim(vRel(iRel).pedido)
					t("obs_2") = Trim("" & vRel(iRel).obs_2)
					t("obs_3") = Trim("" & vRel(iRel).obs_3)
					t("numeroNFe") = Trim("" & vRel(iRel).numeroNFe)
					t("loja") = Trim("" & vRel(iRel).loja)
					t("id_cliente") = Trim("" & vRel(iRel).id_cliente)
					t("transportadora_id") = Trim("" & vRel(iRel).transportadora_id)
					t("qtde_volumes_pedido") = vRel(iRel).qtde_volumes_pedido
					t("destino_tipo_endereco") = vRel(iRel).destino_tipo_endereco
					t("destino_endereco") = vRel(iRel).destino_endereco
					t("destino_endereco_numero") = vRel(iRel).destino_endereco_numero
					t("destino_endereco_complemento") = vRel(iRel).destino_endereco_complemento
					t("destino_bairro") = vRel(iRel).destino_bairro
					t("destino_cidade") = vRel(iRel).destino_cidade
					t("destino_uf") = vRel(iRel).destino_uf
					t("destino_cep") = vRel(iRel).destino_cep
					t.Update
					end if 'if Trim(vRel(iRel).pedido) = pedido_a
				
				s_log = s_log & Cstr(vRel(iRel).qtde) & "x(" & vRel(iRel).fabricante & ")" & vRel(iRel).produto & "[" & vRel(iRel).zona_codigo & "]"
				
				s_sql = "UPDATE" & _
							" t_PEDIDO_ITEM" & _
						" SET" & _
							" separacao_rel_nsu = " & lngNsuWmsEtqN1 & "," & _
							" separacao_data = Convert(varchar(10),getdate(), 121)," & _
							" separacao_data_hora = getdate()," & _
							" separacao_deposito_zona_id = " & vRel(iRel).zona_id & _
						" WHERE" & _
							" (pedido = '" & Trim(vRel(iRel).pedido) & "')" & _
							" AND (fabricante = '" & Trim(vRel(iRel).fabricante) & "')" & _
							" AND (produto = '" & Trim(vRel(iRel).produto) & "')"
				cn.Execute s_sql, lngRecordsAffected
				if lngRecordsAffected <> 1 then
					s_erro_fatal = "Erro ao tentar atualizar o pedido " & Trim(vRel(iRel).pedido) & ": anotação de informações no item de pedido referente ao produto (" & Trim(vRel(iRel).fabricante) & ")" & Trim(vRel(iRel).produto) & " falhou!!"
					exit for
					end if
				
				if Not fin_gera_nsu(T_WMS_ETQ_N3_SEPARACAO_ZONA_PRODUTO, lngNsuWmsEtqN3, msg_erro) then
					s_erro_fatal = "FALHA AO GERAR NSU PARA IDENTIFICAÇÃO DESTA EXECUÇÃO DO RELATÓRIO (" & msg_erro & ")"
					exit for
					end if
				
				intSequenciaN3 = intSequenciaN3 + 1
				s = "SELECT * FROM t_WMS_ETQ_N3_SEPARACAO_ZONA_PRODUTO WHERE (id = -1)"
				if t.State <> 0 then t.Close
				t.open s, cn
				t.AddNew
				t("id") = lngNsuWmsEtqN3
				t("id_wms_etq_n2") = lngNsuWmsEtqN2
				t("sequencia") = intSequenciaN3
				t("zona_id") = vRel(iRel).zona_id
				t("zona_codigo") = Trim("" & vRel(iRel).zona_codigo)
				t("fabricante") = vRel(iRel).fabricante
				t("produto") = vRel(iRel).produto
				t("qtde") = vRel(iRel).qtde
				t("qtde_volumes") = vRel(iRel).qtde_volumes
				t.Update
				end if 'if Trim(vRel(iRel).produto) <> ""
			next 'for iRel=Lbound(vRel) to Ubound(vRel)
		end if
	
	if s_erro_fatal = "" then
		if s_log <> "" then
			s_log = "T_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO.id=" & Cstr(lngNsuWmsEtqN1) & ";  " & s_log
			grava_log usuario, "", "", "", OP_LOG_REL_SEPARACAO_ZONA, s_log
			end if
		end if
	
	if s_erro_fatal = "" then
	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
	else
	'	~~~~~~~~~~~~~~~~
		cn.RollbackTrans
	'	~~~~~~~~~~~~~~~~
		s_erro_html = "<br>" & chr(13) & _
				"<p class='T'>A V I S O</p>" & chr(13) & _
				"<div class='MtAlerta' style='width:600px;font-weight:bold;' align='center'><span style='margin:5px 2px 5px 2px;'>" & s_erro_fatal & "</span></div>" & chr(13) & _
				"<br>" & chr(13) & _
				"<br>" & chr(13) & _
				"<span class=" & chr(34) & "TracoBottom" & chr(34) & "></span>" & chr(13) & _
				"<table cellspacing='0'>" & chr(13) & _
				"	<tr>" & chr(13) & _
				"		<td align='center'><a name='bVOLTAR' id='bVOLTAR' href='javascript:history.back()'><img src='../botao/voltar.gif' width='176' height='55' border='0'></a></td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"</table>" & chr(13) & _
				"</form>" & chr(13) & _
				"</div>" & chr(13) & _
				"</center>" & chr(13) & _
				"</body>" & chr(13) & _
				"</html>" & chr(13)
		Response.Write s_erro_html
		Response.Flush
		Response.End
		end if
	
'	EXIBE O RELATÓRIO
	xRel = "<table width='649' cellspacing='0' cellpadding='0' border='0'>" & chr(13) & _
			xRel & _
			"</table>" & chr(13)
	Response.write xRel

	if strJS_AllTablesCollapse <> "" then
		strJS_AllTablesCollapse = "function AllTablesCollapse() {" & chr(13) & _
								  "var oTrRel;" & chr(13) & _
								  strJS_AllTablesCollapse & _
								  "}" & chr(13)
		end if
		
	if strJS_AllTablesNotPrint <> "" then
		strJS_AllTablesNotPrint = "function AllTablesNotPrint() {" & chr(13) & _
								  "var divAlertaImpressao, divBody, oTrRel, oTrTitZona;" & chr(13) & _
								  "	divAlertaImpressao = document.getElementById('divAlertaImpressao');" & chr(13) & _
								  "	divAlertaImpressao.className = 'notVisible';" & chr(13) & _
								  "	divBody = document.getElementById('divBody');" & chr(13) & _
								  "	divBody.className = 'notPrint';" & chr(13) & _
								  strJS_AllTablesNotPrint & _
								  "}" & chr(13)
		end if
		
	strScriptJS = "<script language='JavaScript' type='text/javascript'>" & chr(13) & _
				  "var nsuRelatorio=" & Cstr(lngNsuWmsEtqN1) & ";" & chr(13) & _
				  strJS_AllTablesCollapse & _
				  strJS_AllTablesNotPrint & _
				  "</script>" & chr(13)
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
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(document).ready(function() {
		$("#spanFiltroNsuRelatorio").text(nsuRelatorio);
	});
</script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir( id_pedido ) {
	fREL.action = "pedido.asp";
	fREL.pedido_selecionado.value = id_pedido;
	fREL.submit();
}

function fExibeRelZona(idx) {
var oTrRel, oTrTitZona, divAlertaImpressao, divBody, strStyleDisplay;
	oTrRel = document.getElementById("trRelZona_" + idx);
	strStyleDisplay = oTrRel.style.display;
	
	AllTablesNotPrint();
	AllTablesCollapse();

	if (strStyleDisplay == "none") {
		divAlertaImpressao = document.getElementById("divAlertaImpressao");
		divAlertaImpressao.className = "notVisible notPrint";
		divBody = document.getElementById("divBody");
		divBody.className = "";
		oTrTitZona = document.getElementById("trTitZona_" + idx);
		oTrTitZona.className = "";
		oTrRel.style.display = "";
		oTrRel.className = "";
	}
}

function fImprime(idx) {
var oTrRel, oTrTitZona, divAlertaImpressao, divBody;
	AllTablesNotPrint();
	AllTablesCollapse();
	divAlertaImpressao = document.getElementById("divAlertaImpressao");
	divAlertaImpressao.className = "notVisible notPrint";
	divBody = document.getElementById("divBody");
	divBody.className = "";
	oTrTitZona = document.getElementById("trTitZona_" + idx);
	oTrTitZona.className = "";
	oTrRel = document.getElementById("trRelZona_" + idx);
	oTrRel.style.display = "";
	oTrRel.className = "";
	print();
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
<link href="<%=URL_FILE__ESCREEN_CSS%>" Rel="stylesheet" Type="text/css" media="screen">

<style type="text/css">
P.C { font-size:10pt; }
P.Cc { font-size:10pt; }
P.Cd { font-size:10pt; }
.Nni
{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10pt;
	font-weight: normal;
	font-style: italic;
}
.tdPedido
{
	vertical-align:bottom;
	width:75px;
	border-right:1px solid;
}
.tdQtVol
{
	vertical-align:bottom;
	text-align:center;
	border-right:1px solid;
	width:35px;
}
.tdNF
{
	vertical-align:bottom;
	width:70px;
	border-right:1px solid;
}
.tdLoja
{
	vertical-align:bottom;
	width:45px;
	border-right:1px solid;
}
.tdCliente
{
	vertical-align:bottom;
	width:337px;
	border-right:1px solid;
}
.tdTransp
{
	vertical-align:bottom;
	width:80px;
}
.tdQtde
{
	vertical-align:bottom;
	width:40px;
}
.tdProd
{
	vertical-align:bottom;
}
.tdZona
{
	vertical-align:bottom;
	width:30px;
}
.tdTotCodProd
{
	vertical-align:bottom;
	width:65px;
}
.tdTotDescrProd
{
	vertical-align:bottom;
	width:400px;
}
.tdTotQtde
{
	vertical-align:bottom;
	width:50px;
}
.tdTotZona
{
	vertical-align:bottom;
	width:30px;
}
</style>


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body>
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<% if s_table_produtos_sem_zona <> "" then Response.Write "<br>" & s_table_produtos_sem_zona%>
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
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';AllTablesCollapse();AllTablesNotPrint();" link=#000000 alink=#000000 vlink=#000000>
<center>
<div id="divAlertaImpressao" class="notVisible"><p class="ALERTA">Para imprimir os dados do relatório,<br />clique no ícone de impressora ao lado do título de cada zona.</p></div>

<div id="divBody" class="notPrint">
<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="rb_nfe" id="rb_nfe" value="<%=rb_nfe%>">
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Separação (Zona)</span>
	<br class="notPrint"><span class="Rc notPrint">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<%
	s_filtro ="<table width='649' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black'>" & chr(13)
	
'	PERÍODO
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>Período:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	NFe EMITIDA
	s = ""
	s_aux = rb_nfe
	if s_aux = "EMITIDA" then
		s = "Somente com NFe já emitida"
	elseif s_aux = "NAO_EMITIDA" then
		s = "Somente com NFe não emitida"
	else
		s = "Ambos"
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>NFe Emitida:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	TRANSPORTADORA
	s = ""
	s_aux = c_transportadora
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>Transportadora:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	QTDE MÁXIMA DE PEDIDOS
	s = ""
	s_aux = c_qtde_max_pedidos
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>Qtde Máxima de Pedidos:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	NSU RELATÓRIO
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>NSU do Relatório:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span id='spanFiltroNsuRelatorio' class='N'></span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	EMISSÃO
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>Emissão:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & formata_data_hora_sem_seg(dt_hr_emissao) & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)

	Response.Write s_filtro
%>
<br>


<!--  RELATÓRIO  -->
<% consulta_relatorio %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</form>
</div>

</center>
</body>

<%=strScriptJS%>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
