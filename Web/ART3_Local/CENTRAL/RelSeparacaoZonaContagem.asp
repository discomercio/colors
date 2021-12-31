<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelSeparacaoZonaContagem.asp
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

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	Const ST_NFE_CANCELADA = "CAN"
	Const ST_NFE_AUTORIZADA = "AUT"

	class cl_NFe_CHECK
		dim pedido
		dim st_nfe_verificada
		dim st_pedido_emissao_nfe_ok
		dim nfe_fatura
		dim st_nfe_fatura
		dim nfe_remessa
		dim st_nfe_remessa
		end class

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_SEPARACAO_ZONA, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim n_total_pedidos_selecionados, n_total_pedidos_disponiveis
	n_total_pedidos_selecionados = 0
	n_total_pedidos_disponiveis = 0
	
	dim intQtdePedidosListadosRelAnterior, strTablePedidosListadosRelAnterior, strListaPedidosListadosRelAnterior
	intQtdePedidosListadosRelAnterior = 0
	strTablePedidosListadosRelAnterior = ""
	strListaPedidosListadosRelAnterior = "|"
	
	dim alerta, s, s_aux, s_table_produtos_sem_zona, s_filtro, intSequencia
	dim c_dt_inicio, c_dt_termino, rb_nfe, c_transportadora, c_qtde_max_pedidos, c_nfe_emitente
	alerta = ""
	s_table_produtos_sem_zona = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_transportadora = Trim(Request.Form("c_transportadora"))
	rb_nfe = Trim(Request.Form("rb_nfe"))
	c_qtde_max_pedidos = retorna_so_digitos(Request.Form("c_qtde_max_pedidos"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))

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

	dim rNfeEmitente
	
	if alerta = "" then
		set rNfeEmitente = le_nfe_emitente(c_nfe_emitente)
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
				" t_PEDIDO_ITEM.descricao," & _
				" t_PEDIDO_ITEM.descricao_html," & _
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
				" AND (t_PEDIDO__BASE.analise_credito = " & COD_AN_CREDITO_OK & ")" & _
				" AND (" & _
						"(t_PEDIDO__BASE.PagtoAntecipadoStatus = " & COD_PAGTO_ANTECIPADO_STATUS_NORMAL & ")" & _
						" OR " & _
						"((t_PEDIDO__BASE.PagtoAntecipadoStatus = " & COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO & ") AND (t_PEDIDO.PagtoAntecipadoQuitadoStatus = " & COD_PAGTO_ANTECIPADO_QUITADO_STATUS_QUITADO & "))" & _
					")" & _
				" AND (t_PEDIDO.id_nfe_emitente = " & rNfeEmitente.id & ")"
		
		if IsDate(c_dt_inicio) then
			s = s & " AND (t_PEDIDO.a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
			end if

		if IsDate(c_dt_termino) then
			s = s & " AND (t_PEDIDO.a_entregar_data_marcada < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
			end if

		if c_transportadora <> "" then
			s = s & " AND (t_PEDIDO.transportadora_id = '" & c_transportadora & "')"
			end if

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
		
		if rs.State <> 0 then rs.Close
		set rs=nothing

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

	'Localiza dados para a conexão com o BD de NFe
	dim dbcNFe
	dim strNfeT1ServidorBd, strNfeT1NomeBd, strNfeT1UsuarioBd, strNfeT1SenhaCriptografadaBd
	dim chave
	dim senha_decodificada

	if alerta = "" then
		s = "SELECT" & _
				" NFe_T1_servidor_BD," & _
				" NFe_T1_nome_BD," & _
				" NFe_T1_usuario_BD," & _
				" NFe_T1_senha_BD" & _
			" FROM t_NFe_EMITENTE" & _
			" WHERE" & _
				" (id = " & c_nfe_emitente & ")"
		set rs = cn.Execute(s)
		if rs.Eof then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Não foram localizados os parâmetros de conexão ao BD de NFe (ID=" & c_nfe_emitente & ")"
		else
			strNfeT1ServidorBd = Trim("" & rs("NFe_T1_servidor_BD"))
			strNfeT1NomeBd = Trim("" & rs("NFe_T1_nome_BD"))
			strNfeT1UsuarioBd = Trim("" & rs("NFe_T1_usuario_BD"))
			strNfeT1SenhaCriptografadaBd = Trim("" & rs("NFe_T1_senha_BD"))
			
			chave = gera_chave(FATOR_BD)
			decodifica_dado strNfeT1SenhaCriptografadaBd, senha_decodificada, chave
			end if
		
		if rs.State <> 0 then rs.Close
		set rs=nothing
		end if

	if alerta = "" then
		'Abre a conexão com o BD de NFe
		s = "Provider=SQLOLEDB;" & _
			"Data Source=" & strNfeT1ServidorBd & ";" & _
			"Initial Catalog=" & strNfeT1NomeBd & ";" & _
			"User ID=" & strNfeT1UsuarioBd & ";" & _
			"Password=" & senha_decodificada & ";"
		set dbcNFe = server.CreateObject("ADODB.Connection")
		dbcNFe.ConnectionTimeout = 45
		dbcNFe.CommandTimeout = 900
		dbcNFe.ConnectionString = s
		dbcNFe.Open
		end if




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

sub inicializa_cl_NFe_CHECK(byref o)
	o.pedido = ""
	o.st_nfe_verificada = False
	o.st_pedido_emissao_nfe_ok = False
	o.nfe_fatura = ""
	o.st_nfe_fatura = ""
	o.nfe_remessa = ""
	o.st_nfe_remessa = ""
end sub


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
dim r, tNFE
dim s, s_sql, s_where, s_sql_lista_pedidos
dim s_sql_nfe, s_where_nfe
dim i, iZona, iRel, iStep, iNFe, idxNFe, n_qtde_zona, n_reg
dim vZona(), vRel(), vNFe()
dim cab_table, cab
dim x
dim s_tit_zona
dim blnProcessaRegistro
dim s_pedido_aux, lista_pedidos, lista_pedidos_selecionados, strQtdePedidos
dim qtde_pedidos, qtde_produtos, qtde_volumes
dim n_qtde_max_pedidos
dim blnPedidoComNFeInvalida

'	VETOR QUE ARMAZENA TODOS OS REGISTROS (A ORDENAÇÃO DEVE SER FEITA NA CONSULTA SQL)
	redim vRel(0)
	set vRel(0) = New cl_REL_SEPARACAO_ZONA
	vRel(0).produto = ""

'	ARMAZENA A LISTA DE ZONAS CADASTRADAS
	redim vZona(0)
	set vZona(0) = New cl_ZONA_DEPOSITO
	vZona(0).zona_id = 0
	
'	VETOR USADO NA VERIFICAÇÃO DO STATUS DA NFe
	redim vNFe(0)
	set vNFe(0) = New cl_NFe_CHECK
	inicializa_cl_NFe_CHECK vNFe(0)


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

	s_sql_lista_pedidos = ""
	lista_pedidos = ""

	for iStep = 1 to 2
	'	OBSERVANDO QUE:
	'		obs_2 = Nº Nota Fiscal
	'		obs_3 = NF Simples Remessa, quando houver
	'		num_obs_2 = Computed column com o valor do campo 'obs_2' convertido para INT
	'		num_obs_3 = Computed column com o valor do campo 'obs_3' convertido para INT
		s_sql = "SELECT" & _
					" t_PEDIDO.pedido," & _
					" t_PEDIDO.data," & _
					" t_PEDIDO.hora," & _
					" t_PEDIDO.obs_2," & _
					" t_PEDIDO.num_obs_2," & _
					" t_PEDIDO.obs_3," & _
					" t_PEDIDO.num_obs_3," & _
					" t_PEDIDO.loja," & _
					" t_PEDIDO.transportadora_id," & _
					" t_PEDIDO.id_cliente,"

		if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
			s_sql = s_sql & _
						" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_cliente,"
		else
			s_sql = s_sql & _
						" t_CLIENTE.nome_iniciais_em_maiusculas AS nome_cliente,"
			end if

		s_sql = s_sql & _
					" t_FABRICANTE.nome AS nome_fabricante," & _
					" t_FABRICANTE.razao_social AS razao_social_fabricante," & _
					" t_PEDIDO_ITEM.fabricante," & _
					" t_PEDIDO_ITEM.produto," & _
					" t_PEDIDO_ITEM.qtde," & _
					" t_PEDIDO_ITEM.qtde_volumes," & _
					" t_PEDIDO_ITEM.descricao," & _
					" t_PEDIDO_ITEM.descricao_html," & _
					" t_PRODUTO.deposito_zona_id AS zona_id," & _
					" t_WMS_DEPOSITO_MAP_ZONA.zona_codigo," & _
					" (" & _
						"CASE" & _
							" WHEN t_PEDIDO.num_obs_3 > 0 THEN t_PEDIDO.num_obs_3" & _
							" WHEN t_PEDIDO.num_obs_2 > 0 THEN t_PEDIDO.num_obs_2" & _
							" ELSE NULL" & _
						" END) AS numeroNFe," & _
					" (" & _
						"SELECT" & _
							" TOP 1 id_wms_etq_n1" & _
						" FROM t_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO tN2" & _
						" WHERE" & _
							" (tN2.pedido=t_PEDIDO.pedido)" & _
						" ORDER BY" & _
							" id_wms_etq_n1 DESC" & _
					") AS nsuRelSeparacaoZonaAnterior," & _
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
					" AND (t_PEDIDO__BASE.analise_credito = " & COD_AN_CREDITO_OK & ")" & _
					" AND (" & _
							"(t_PEDIDO__BASE.PagtoAntecipadoStatus = " & COD_PAGTO_ANTECIPADO_STATUS_NORMAL & ")" & _
							" OR " & _
							"((t_PEDIDO__BASE.PagtoAntecipadoStatus = " & COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO & ") AND (t_PEDIDO.PagtoAntecipadoQuitadoStatus = " & COD_PAGTO_ANTECIPADO_QUITADO_STATUS_QUITADO & "))" & _
						")" & _
					" AND (t_PEDIDO.id_nfe_emitente = " & rNfeEmitente.id & ")"

		if IsDate(c_dt_inicio) then
			s_sql = s_sql & " AND (t_PEDIDO.a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
			end if
			
		if IsDate(c_dt_termino) then
			s_sql = s_sql & " AND (t_PEDIDO.a_entregar_data_marcada < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
			end if

		if c_transportadora <> "" then
			s_sql = s_sql & " AND (t_PEDIDO.transportadora_id = '" & c_transportadora & "')"
			end if
		
		if iStep = 2 then
			if s_sql_lista_pedidos <> "" then
				s_sql = s_sql & " AND (t_PEDIDO.pedido IN (" & s_sql_lista_pedidos & "))"
			else
				'Nenhum pedido em situação válida foi encontrado no passo 1
				'Isso pode ocorrer se nenhum pedido estiver com a NFe em situação válida
				s_sql = s_sql & " AND (t_PEDIDO.pedido IN ('XXXXXXXXXXXX'))"
				end if
			end if
		
		s_sql = "SELECT " & _
					"*" & _
				" FROM " & _
					"(" & s_sql & ") t"
		
	'	NFe EMITIDA?
		s_where = ""
		if rb_nfe = "EMITIDA" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (numeroNFe IS NOT NULL)"
		elseif rb_nfe = "NAO_EMITIDA" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (numeroNFe IS NULL)"
			end if
		
		if s_where <> "" then s_where = " WHERE" & s_where
		s_sql = s_sql & s_where
		s_sql = s_sql & " ORDER BY data, hora, pedido"
	
		set r = cn.execute(s_sql)

	'	NÃO HÁ LIMITAÇÃO DE QTDE MÁXIMA DE PEDIDOS
		if c_qtde_max_pedidos = "" then
			n_qtde_max_pedidos = 0
		else
			n_qtde_max_pedidos = converte_numero(c_qtde_max_pedidos)
			end if
		
	'	VERIFICA SE HÁ PEDIDOS COM NFe EM SITUAÇÃO DIFERENTE DE AUTORIZADA
		if iStep = 1 then
			do while Not r.Eof
				'Verifica se o pedido já existe no vetor
				idxNFe = -1
				for iNFe=Lbound(vNFe) to Ubound(vNFe)
					if vNFe(iNFe).pedido = Trim("" & r("pedido")) then
						idxNFe = iNFe
						exit for
						end if
					next

				'Se o pedido ainda não está no vetor, armazena os dados
				if idxNFe = -1 then
					if vNFe(Ubound(vNFe)).pedido <> "" then
						redim preserve vNFe(Ubound(vNFe)+1)
						set vNFe(Ubound(vNFe)) = New cl_NFe_CHECK
						inicializa_cl_NFe_CHECK vNFe(Ubound(vNFe))
						end if
					idxNFe = Ubound(vNFe)
					vNFe(idxNFe).pedido = Trim("" & r("pedido"))
					if r("num_obs_2") > 0 then vNFe(idxNFe).nfe_fatura = NFeFormataNumeroNF(r("num_obs_2"))
					if r("num_obs_3") > 0 then vNFe(idxNFe).nfe_remessa = NFeFormataNumeroNF(r("num_obs_3"))
					end if

				r.MoveNext
				loop
			
			if r.State <> 0 then r.Close
			set r=nothing

			'MONTA A CONSULTA PARA O BD DE NFe DA TARGET ONE
			s_where_nfe = ""
			for iNFe = LBound(vNFe) to UBound(vNFe)
				if vNFe(iNFe).pedido <> "" then
					if vNFe(iNFe).nfe_fatura <> "" then
						if s_where_nfe <> "" then s_where_nfe = s_where_nfe & ","
						s_where_nfe = s_where_nfe & "'" & vNFe(iNFe).nfe_fatura & "'"
						end if

					if vNFe(iNFe).nfe_remessa <> "" then
						if s_where_nfe <> "" then s_where_nfe = s_where_nfe & ","
						s_where_nfe = s_where_nfe & "'" & vNFe(iNFe).nfe_remessa & "'"
						end if
					end if
				next

			if s_where_nfe <> "" then
				s_where_nfe = " (Nfe IN (" & s_where_nfe & "))"
				s_sql_nfe = "SELECT" & _
								" Nfe," & _
								" Serie," & _
								" Convert(tinyint, Coalesce(CANCELADA,0)) AS CANCELADA," & _
								" CodProcAtual" & _
							" FROM NFE" & _
							" WHERE" & _
								s_where_nfe
				set tNFE = dbcNFe.Execute(s_sql_nfe)
				do while Not tNFE.Eof
					idxNFe = -1
					for iNFe=LBound(vNFe) to UBound(vNFe)
						if Trim("" & tNFE("Nfe")) = vNFe(iNFe).nfe_fatura then
							vNFe(iNFe).st_nfe_verificada = True
							if Trim("" & tNFE("CANCELADA")) = "1" then
								vNFe(iNFe).st_nfe_fatura = ST_NFE_CANCELADA
							elseif Trim("" & tNFE("CodProcAtual")) = "100" then
								vNFe(iNFe).st_nfe_fatura = ST_NFE_AUTORIZADA
								end if
							exit for
						elseif Trim("" & tNFE("Nfe")) = vNFe(iNFe).nfe_remessa then
							vNFe(iNFe).st_nfe_verificada = True
							if Trim("" & tNFE("CANCELADA")) = "1" then
								vNFe(iNFe).st_nfe_remessa = ST_NFE_CANCELADA
							elseif Trim("" & tNFE("CodProcAtual")) = "100" then
								vNFe(iNFe).st_nfe_remessa = ST_NFE_AUTORIZADA
								end if
							exit for
							end if
						next

					tNFE.MoveNext
					loop

				if tNFE.State <> 0 then tNFE.Close
				set tNFE=nothing
				
				'Consolida o status de emissão da NFe do pedido, lembrando que um pedido pode ter duas NFes (fatura e remessa)
				for iNFe=LBound(vNFe) to UBound(vNFe)
					if vNFe(iNFe).pedido <> "" then
						if Not vNFe(iNFe).st_nfe_verificada then
							vNFe(iNFe).st_pedido_emissao_nfe_ok = False
						else
							blnPedidoComNFeInvalida = False
							if vNFe(iNFe).nfe_fatura <> "" then
								if vNFe(iNFe).st_nfe_fatura <> ST_NFE_AUTORIZADA then blnPedidoComNFeInvalida = True
								end if
						
							if vNFe(iNFe).nfe_remessa <> "" then
								if vNFe(iNFe).st_nfe_remessa <> ST_NFE_AUTORIZADA then blnPedidoComNFeInvalida = True
								end if
						
							if blnPedidoComNFeInvalida then
								vNFe(iNFe).st_pedido_emissao_nfe_ok = False
							else
								vNFe(iNFe).st_pedido_emissao_nfe_ok = True
								end if
							end if
						end if
					next
				end if 'if s_where_nfe <> ""

			'Monta lista de pedidos que serão considerados para o relatório
			for iNFe=LBound(vNFe) to UBound(vNFe)
				if vNFe(iNFe).pedido <> "" then
					if vNFe(iNFe).st_pedido_emissao_nfe_ok Or _
						( (rb_nfe <> "EMITIDA") AND ((vNFe(iNFe).nfe_fatura="") AND (vNFe(iNFe).nfe_remessa="")) ) then
						s_pedido_aux = Trim("" & vNFe(iNFe).pedido)
						if Instr(lista_pedidos, "|" & s_pedido_aux & "|") = 0 then
							n_total_pedidos_disponiveis = n_total_pedidos_disponiveis + 1
							lista_pedidos = lista_pedidos & "|" & s_pedido_aux & "|"
							if (n_total_pedidos_disponiveis <= n_qtde_max_pedidos) Or (n_qtde_max_pedidos = 0) then
								if s_sql_lista_pedidos <> "" then s_sql_lista_pedidos = s_sql_lista_pedidos & ","
								s_sql_lista_pedidos = s_sql_lista_pedidos & "'" & s_pedido_aux & "'"
								end if
							end if
						end if
					end if
				next
			end if 'if iStep = 1
		next ' for iStep = 1 to 2
	
	
'	ARMAZENA TODOS OS REGISTROS NO VETOR
	do while Not r.Eof
		if Trim("" & vRel(Ubound(vRel)).produto) <> "" then
			redim preserve vRel(Ubound(vRel)+1)
			set vRel(Ubound(vRel)) = New cl_REL_SEPARACAO_ZONA
			end if
		
		with vRel(Ubound(vRel))
			.pedido = Trim("" & r("pedido"))
			'Número da NFe sem formatação com zeros à esquerda
			.obs_2 = Trim("" & r("obs_2"))
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
			.numeroNFe = Trim("" & r("numeroNFe"))
			.qtde_volumes_pedido = r("qtde_volumes_pedido")
			.nome_fabricante = Trim("" & r("nome_fabricante"))
			if .nome_fabricante = "" then .nome_fabricante = Trim("" & r("razao_social_fabricante"))
			end with
		
		if Trim("" & r("nsuRelSeparacaoZonaAnterior")) <> "" then
			if Instr(strListaPedidosListadosRelAnterior, Trim("" & r("pedido"))) = 0 then
				intQtdePedidosListadosRelAnterior = intQtdePedidosListadosRelAnterior + 1
				strListaPedidosListadosRelAnterior = strListaPedidosListadosRelAnterior & Trim("" & r("pedido")) & "|"
				strTablePedidosListadosRelAnterior = strTablePedidosListadosRelAnterior & _
					"	<tr>" & chr(13) & _
					"		<td class='MB ME MD' align='left'><span class='PLLe RowLstRelAnt'>" & Trim("" & r("pedido")) & "</span></td>" & chr(13) & _
					"		<td class='MB MD' align='right'><span class='PLLd RowLstRelAnt'>" & Trim("" & r("nsuRelSeparacaoZonaAnterior")) & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
				end if
			end if
		
		r.MoveNext
		loop

	if r.State <> 0 then r.Close
	set r=nothing

	if strTablePedidosListadosRelAnterior <> "" then
		strTablePedidosListadosRelAnterior = "<table cellspacing='0' cellpadding='1' border='0'>" & chr(13) & _
											"	<tr>" & chr(13) & _
											"		<td class='MC MB ME MD' align='left' style='background:azure;'>" & "<span class='PLTe TitLstRelAnt'>Pedido</span>" & "</td>" & chr(13) & _
											"		<td class='MC MB MD' align='right' style='background:azure;'>" & "<span class='PLTd TitLstRelAnt'>NSU</span>" & "</td>" & chr(13) & _
											"	</tr>" & chr(13) & _
												strTablePedidosListadosRelAnterior & _
											"</table>" & chr(13)
		end if
	
	cab_table = "<table id='tableRelSeparacao' class='Q' style='border-bottom:0px;' cellspacing='0' cellpadding='2'>" & chr(13)
	cab =	"	<tr style='background:azure' nowrap>" & chr(13) & _
			"		<th class='MB MD thZona' align='center' valign='bottom'><span class='C thCel'>Zona</span></th>" & chr(13) & _
			"		<th class='MB MD thQtdePed' align='right' valign='bottom'><span class='Cd thCel'>Qtde</span><br /><span class='Cd thCel'>Pedidos</span></th>" & chr(13) & _
			"		<th class='MB MD thQtdeProd' align='right' valign='bottom'><span class='Cd thCel'>Qtde</span><br /><span class='Cd thCel'>Produtos</span></th>" & chr(13) & _
			"		<th class='MB thQtdeVol' align='right' valign='bottom'><span class='Cd thCel'>Qtde</span><br /><span class='Cd thCel'>Volumes</span></th>" & chr(13) & _
			"	</tr>" & chr(13)
	
	n_qtde_zona = 0
	x = cab_table & cab

'	GERA O RELATÓRIO PARA CADA UMA DAS ZONAS
	for iZona=Lbound(vZona) to Ubound(vZona)
		if converte_numero(vZona(iZona).zona_id) <> 0 then
			n_reg = 0
			n_qtde_zona = n_qtde_zona + 1
			lista_pedidos = ""
			qtde_pedidos = 0
			qtde_produtos = 0
			qtde_volumes = 0
			
			for iRel=Lbound(vRel) to Ubound(vRel)
				blnProcessaRegistro = False
				if Trim("" & vRel(iRel).produto) <> "" then
					if Trim("" & vRel(iRel).zona_id) = Trim("" & vZona(iZona).zona_id) then blnProcessaRegistro = True
					if vZona(iZona).zona_id = COD_ZONA_ID__TODOS then blnProcessaRegistro = True
					end if
				
				if blnProcessaRegistro then
					with vRel(iRel)
						s_pedido_aux = "|" & Trim("" & .pedido) & "|"
						if Instr(lista_pedidos, s_pedido_aux) = 0 then
							lista_pedidos = lista_pedidos & s_pedido_aux
							qtde_pedidos = qtde_pedidos + 1
							end if
						
						qtde_produtos = qtde_produtos + .qtde
						qtde_volumes = qtde_volumes + (.qtde * .qtde_volumes)
						
					  ' CONTAGEM
						n_reg = n_reg + 1
						if vZona(iZona).zona_id = COD_ZONA_ID__TODOS then n_total_pedidos_selecionados = n_total_pedidos_selecionados + 1
						end with
					end if ' if blnProcessaRegistro
				next 'for iRel=Lbound(vRel) to Ubound(vRel)
			
			if vZona(iZona).zona_id = COD_ZONA_ID__TODOS then
				s_tit_zona = "Todas"
				strQtdePedidos = Cstr(qtde_pedidos)
			else
				s_tit_zona = vZona(iZona).zona_codigo
				strQtdePedidos = "&nbsp;"
				end if
			
		'	FINALIZAÇÃO
			if n_reg <> 0 then 
			'	TOTAIS
				x = x & _
					"	<tr>" & chr(13) & _
					"		<td class='MB MD tdZona' align='center'><span class='Rc tdCel'>" & Ucase(s_tit_zona) & "</span></td>" & chr(13) & _
					"		<td class='MB MD tdQtdePed' align='right'><span class='Rd tdCel'>" & strQtdePedidos & "</span></td>" & chr(13) & _
					"		<td class='MB MD tdQtdeProd' align='right'><span class='Rd tdCel'>" & qtde_produtos & "</span></td>" & chr(13) & _
					"		<td class='MB tdQtdeVol' align='right'><span class='Rd tdCel'>" & qtde_volumes & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
				end if

		  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
			if n_reg = 0 then
				x = x & _
					"	<tr>" & chr(13) & _
					"		<td class='MB MD tdZona' align='center'><span class='Rc tdCel'>" & Ucase(s_tit_zona) & "</span></td>" & chr(13) & _
					"		<td class='MB MD tdQtdePed' align='right'><span class='Rd tdCel'>&nbsp;</span></td>" & chr(13) & _
					"		<td class='MB MD tdQtdeProd' align='right'><span class='Rd tdCel'>&nbsp;</span></td>" & chr(13) & _
					"		<td class='MB tdQtdeVol' align='right'><span class='Rd tdCel'>&nbsp;</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
				end if
			end if ' if converte_numero(vZona(iZona).zona_id) <> 0
		next ' for iZona=Lbound(vZona) to Ubound(vZona)
	
	
'	EXIBE O RELATÓRIO
	x = x & _
		"</table>" & chr(13)
	Response.write x
	
	x = "<br>" & chr(13) & _
		"<table width='649' cellpadding='0' cellspacing='0'>" & chr(13) & _
		"	<tr>" & chr(13) & _
		"		<td colspan='4' align='left'>&nbsp;</td>" & chr(13) & _
		"	</tr>" & chr(13) & _
		"	<tr>" & chr(13) & _
		"		<td colspan='4' align='left'><span class='Lbl'>Total de pedidos disponíveis de acordo com os critérios de seleção: " & Cstr(n_total_pedidos_disponiveis) & "</span></td>" & chr(13) & _
		"	</tr>" & chr(13) & _
		"<table>" & chr(13)
	Response.write x
	
'	CAMPO QUE ARMAZENA A RELAÇÃO DE PEDIDOS SELECIONADOS
	lista_pedidos_selecionados = "INICIO|"
	for iRel=Lbound(vRel) to Ubound(vRel)
		if Trim("" & vRel(iRel).pedido) <> "" then
			s_pedido_aux = "|" & Trim("" & vRel(iRel).pedido) & "|"
			if Instr(lista_pedidos_selecionados, s_pedido_aux) = 0 then
				lista_pedidos_selecionados = lista_pedidos_selecionados & Trim("" & vRel(iRel).pedido) & "|"
				end if
			end if
		next
	lista_pedidos_selecionados = lista_pedidos_selecionados & "TERMINO"
	
	x = chr(13) & _
		"<input type='hidden' name='c_lista_pedidos_selecionados' id='c_lista_pedidos_selecionados' value='" & lista_pedidos_selecionados & "' />" & chr(13) & _
		"<input type='hidden' name='c_qtde_total_pedidos_disponiveis' id='c_qtde_total_pedidos_disponiveis' value='" & Cstr(n_total_pedidos_disponiveis) & "' />" & chr(13)
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


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(document).ready(function() {
		$("#tableRelSeparacao tr").not(':first').hover(
			function() {
				$(this).css("background", "#98FB98");
			},
			function() {
				$(this).css("background", "");
			}
		)
	});
</script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConfirma( f ) {
	dCONFIRMA.style.visibility = "hidden";
	window.status = "Aguarde ...";
	f.action = "RelSeparacaoZona.asp";
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__ESCREEN_CSS%>" rel="stylesheet" type="text/css" media="screen">

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
.thCel
{
	font-size:12pt;
	font-weight:bold;
	color:#000000;
}
.tdCel
{
	font-size:12pt;
	font-weight:bold;
	color:#000000;
}
.thZona
{
	vertical-align:bottom;
	width:100px;
}
.thQtdePed
{
	vertical-align:bottom;
	width:100px;
}
.thQtdeProd
{
	vertical-align:bottom;
	width:100px;
}
.thQtdeVol
{
	vertical-align:bottom;
	width:100px;
}
.tdZona
{
	vertical-align:bottom;
	width:100px;
}
.tdQtdePed
{
	vertical-align:bottom;
	width:100px;
}
.tdQtdeProd
{
	vertical-align:bottom;
	width:100px;
}
.tdQtdeVol
{
	vertical-align:bottom;
	width:100px;
}
.TitLstRelAnt 
{
	font-size:10pt;
	margin-left:5px;
	margin-right:5px;
}
.RowLstRelAnt
{
	font-size:10pt;
	margin-left:5px;
	margin-right:5px;
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
<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>
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
<input type="hidden" name="c_qtde_max_pedidos" id="c_qtde_max_pedidos" value="<%=c_qtde_max_pedidos%>">
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>" />


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
				"		<td align='right' valign='top' nowrap><span class='Nni'>NFe emitida:&nbsp;</span></td>" & chr(13) & _
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


'	CD
	s = obtem_apelido_empresa_NFe_emitente(c_nfe_emitente)
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>CD:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	EMISSÃO
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>Emissão:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & formata_data_hora_sem_seg(Now) & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)

	Response.Write s_filtro
%>
<br>
<br>


<!--  RELATÓRIO  -->
<% consulta_relatorio %>


<!--  HÁ PEDIDOS LISTADOS EM RELATÓRIO ANTERIOR?  -->
<% if strTablePedidosListadosRelAnterior <> "" then %>
<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-top:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
</br>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center">
<span>ATENÇÃO!<br />Os seguintes pedidos já foram listados em relatório anterior!!<br />Por favor, verifique se houve algum problema!!</span>
</div>
<%
		Response.Write "</br></br>"
		Response.Write strTablePedidosListadosRelAnterior
		Response.Write "</br>"
	end if%>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-top:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>


<table class="notPrint" width="649" cellspacing="0">
<% if (n_total_pedidos_selecionados = 0) Or (intQtdePedidosListadosRelAnterior > 0) then %>
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
<% else %>
<tr>
	<td align="left"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELConfirma(fREL)" title="confirma o processamento do relatório">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% end if %>
</table>

</form>
</div>

</center>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	if alerta = "" then
		dbcNFe.Close
		set dbcNFe = nothing
		end if

	cn.Close
	set cn = nothing
%>
