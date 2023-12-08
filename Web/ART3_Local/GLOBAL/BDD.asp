<%

'     =============
'	  B D D . A S P
'     =============
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



' ======================================================================================
' Registra o log de páginas acessadas
' A rotina é executada toda vez que é acessada uma página ASP que declara um include
' ao BDD.ASP
' ======================================================================================
	grava_log_acesso_pagina
' ======================================================================================


' _____________________________________________________________________________________________
'
'									D E C L A R A Ç Õ E S
' _____________________________________________________________________________________________


    class cl_ITEM_ESTOQUE_ENTRADA_XML
		dim id_estoque
		dim fabricante
		dim produto
		dim qtde
		dim qtde_utilizada
		dim preco_fabricante
		dim data_ult_movimento
		dim sequencia
		dim vl_custo2
		dim vl_BC_ICMS_ST
		dim vl_ICMS_ST
		dim ncm
		dim ncm_redigite
		dim cst
		dim cst_redigite
        dim ean
        dim ean_original
        dim aliq_ipi
        dim aliq_icms
        dim vl_ipi
        dim preco_origem
        dim produto_xml
		dim vl_frete
		end class


' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ______________________________
' GRAVA LOG ACESSO PAGINA
'
sub grava_log_acesso_pagina
dim cn__LogPagAcessada
dim rs__LogPagAcessada
dim str__LogPagAcessada_NomePagina
dim str__LogPagAcessada_IP
dim str__LogPagAcessada_usuario
dim str__LogPagAcessada_SQL
dim int__LogPagAcessada_IdLogPagina
dim int__LogPagAcessada_IdLogPaginaUsuario
dim int__LogPagAcessada_IdLogPaginaIpOrigem
	
	if bdd_conecta(cn__LogPagAcessada) then
		'	Verifica se a página já está cadastrada em t_LOG_PAGINA
		int__LogPagAcessada_IdLogPagina = 0
		str__LogPagAcessada_NomePagina = LCase(Trim("" & Request.ServerVariables("SCRIPT_NAME")))
		str__LogPagAcessada_SQL = "SELECT * FROM t_LOG_PAGINA WHERE (Pagina = '" & QuotedStr(str__LogPagAcessada_NomePagina) & "')"
		set rs__LogPagAcessada = cn__LogPagAcessada.Execute(str__LogPagAcessada_SQL)
		if Not rs__LogPagAcessada.Eof then
			int__LogPagAcessada_IdLogPagina = CLng(rs__LogPagAcessada("Id"))
		else
			str__LogPagAcessada_SQL = "SET NOCOUNT ON; INSERT INTO t_LOG_PAGINA (Pagina) VALUES ('" & QuotedStr(str__LogPagAcessada_NomePagina) & "'); SELECT SCOPE_IDENTITY() AS Id;"
			set rs__LogPagAcessada = cn__LogPagAcessada.Execute(str__LogPagAcessada_SQL)
			if Not rs__LogPagAcessada.Eof then int__LogPagAcessada_IdLogPagina = CLng(rs__LogPagAcessada("Id"))
			end if

		'	Verifica se o usuário já está cadastrado em t_LOG_PAGINA_USUARIO
		int__LogPagAcessada_IdLogPaginaUsuario = 0
		str__LogPagAcessada_usuario = Trim("" & Session("usuario_atual"))
		if str__LogPagAcessada_usuario <> "" then
			str__LogPagAcessada_SQL = "SELECT * FROM t_LOG_PAGINA_USUARIO WHERE (Usuario = '" & QuotedStr(str__LogPagAcessada_usuario) & "')"
			set rs__LogPagAcessada = cn__LogPagAcessada.Execute(str__LogPagAcessada_SQL)
			if Not rs__LogPagAcessada.Eof then
				int__LogPagAcessada_IdLogPaginaUsuario = CLng(rs__LogPagAcessada("Id"))
			else
				str__LogPagAcessada_SQL = "SET NOCOUNT ON; INSERT INTO t_LOG_PAGINA_USUARIO (Usuario) VALUES ('" & QuotedStr(str__LogPagAcessada_usuario) & "'); SELECT SCOPE_IDENTITY() AS Id;"
				set rs__LogPagAcessada = cn__LogPagAcessada.Execute(str__LogPagAcessada_SQL)
				if Not rs__LogPagAcessada.Eof then int__LogPagAcessada_IdLogPaginaUsuario = CLng(rs__LogPagAcessada("Id"))
				end if
			end if

		'	Verifica se o IP de origem já está cadastrado em t_LOG_PAGINA_IP_ORIGEM
		int__LogPagAcessada_IdLogPaginaIpOrigem = 0
		str__LogPagAcessada_IP = Trim("" & Request.ServerVariables("REMOTE_ADDR"))
		if str__LogPagAcessada_IP <> "" then
			str__LogPagAcessada_SQL = "SELECT * FROM t_LOG_PAGINA_IP_ORIGEM WHERE (Ip = '" & QuotedStr(str__LogPagAcessada_IP) & "')"
			set rs__LogPagAcessada = cn__LogPagAcessada.Execute(str__LogPagAcessada_SQL)
			if Not rs__LogPagAcessada.Eof then
				int__LogPagAcessada_IdLogPaginaIpOrigem = CLng(rs__LogPagAcessada("Id"))
			else
				str__LogPagAcessada_SQL = "SET NOCOUNT ON; INSERT INTO t_LOG_PAGINA_IP_ORIGEM (Ip) VALUES ('" & QuotedStr(str__LogPagAcessada_IP) & "'); SELECT SCOPE_IDENTITY() AS Id;"
				set rs__LogPagAcessada = cn__LogPagAcessada.Execute(str__LogPagAcessada_SQL)
				if Not rs__LogPagAcessada.Eof then int__LogPagAcessada_IdLogPaginaIpOrigem = CLng(rs__LogPagAcessada("Id"))
				end if
			end if

		'	Registra o acesso à página
		if int__LogPagAcessada_IdLogPagina > 0 then
			str__LogPagAcessada_SQL = "INSERT INTO t_LOG_PAGINA_ACESSO (IdLogPagina, IdLogPaginaUsuario, IdLogPaginaIpOrigem) VALUES (" & int__LogPagAcessada_IdLogPagina & ", " & int__LogPagAcessada_IdLogPaginaUsuario & ", " & int__LogPagAcessada_IdLogPaginaIpOrigem & ")"
			cn__LogPagAcessada.Execute(str__LogPagAcessada_SQL)
			end if

		cn__LogPagAcessada.Close
		set cn__LogPagAcessada = nothing
		end if ' if bdd_conecta(cn__LogPagAcessada) then
end sub


' ______________________________
' B D D _ C E P _ C O N E C T A
'
function bdd_cep_conecta ( cn )
dim s
dim chave
dim senha_decodificada

	bdd_cep_conecta = False
	
	if is_sgbd_access then
		s = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
			"Data Source=" & NOME_BD_CEP & ";" & _
			"User ID=" & USUARIO_BD_CEP & ";" & _
			"Password=" & SENHA_BD_CEP & ";"
	else	
	'   DECODIFICA SENHA DO BD
		chave = gera_chave(FATOR_BD)
		decodifica_dado SENHA_BD_CEP, senha_decodificada, chave
		s = "Provider=SQLOLEDB;" & _
			"Data Source=" & SERVIDOR_BD_CEP & ";" & _
			"Initial Catalog=" & NOME_BD_CEP & ";" & _
			"User ID=" & USUARIO_BD_CEP & ";" & _
			"Password=" & senha_decodificada & ";"
		end if

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionTimeout = 45
	cn.CommandTimeout = 900
	cn.ConnectionString = s
	cn.Open

	If Err <> 0 then 
		cn.Close
		set cn = nothing
		exit function
		end if

	bdd_cep_conecta = True
	
end function



' _____________________
' B D D _ C O N E C T A
'
function bdd_conecta ( cn )
dim s
dim chave
dim senha_decodificada

	bdd_conecta = False
	
	if is_sgbd_access then
		s = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
			"Data Source=" & NOME_BD & ";" & _
			"User ID=" & USUARIO_BD & ";" & _
			"Password=" & SENHA_BD & ";"
	else	
	'   DECODIFICA SENHA DO BD
		chave = gera_chave(FATOR_BD)
		decodifica_dado SENHA_BD, senha_decodificada, chave
		s = "Provider=SQLOLEDB;" & _
			"Data Source=" & SERVIDOR_BD & ";" & _
			"Initial Catalog=" & NOME_BD & ";" & _
			"User ID=" & USUARIO_BD & ";" & _
			"Password=" & senha_decodificada & ";"
		end if

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionTimeout = 45
	cn.CommandTimeout = 900
	cn.ConnectionString = s
	cn.Open

	If Err <> 0 then 
		cn.Close
		set cn = nothing
		exit function
		end if

	bdd_conecta = True
	
end function



' ____________________________________
' B D D _ C O N E C T A _ R P I F C
'
function bdd_conecta_RPIFC ( cn )
dim s
dim chave
dim senha_decodificada

	bdd_conecta_RPIFC = False
	
	if is_sgbd_access then
		s = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
			"Data Source=" & NOME_BD_RPIFC & ";" & _
			"User ID=" & USUARIO_BD_RPIFC & ";" & _
			"Password=" & SENHA_BD_RPIFC & ";"
	else	
	'   DECODIFICA SENHA DO BD
		chave = gera_chave(FATOR_BD)
		decodifica_dado SENHA_BD_RPIFC, senha_decodificada, chave
		s = "Provider=SQLOLEDB;" & _
			"Data Source=" & SERVIDOR_BD_RPIFC & ";" & _
			"Initial Catalog=" & NOME_BD_RPIFC & ";" & _
			"User ID=" & USUARIO_BD_RPIFC & ";" & _
			"Password=" & senha_decodificada & ";"
		end if

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionTimeout = 45
	cn.CommandTimeout = 900
	cn.ConnectionString = s
	cn.Open

	If Err <> 0 then 
		cn.Close
		set cn = nothing
		exit function
		end if

	bdd_conecta_RPIFC = True
	
end function



' ___________________________
' B D D _ A T _ C O N E C T A
'
function bdd_AT_conecta ( cn )
dim s
dim chave
dim senha_decodificada

	bdd_AT_conecta = False
	
	if is_sgbd_access then
		s = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
			"Data Source=" & NOME_BD_AT & ";" & _
			"User ID=" & USUARIO_BD_AT & ";" & _
			"Password=" & SENHA_BD_AT & ";"
	else	
	'   DECODIFICA SENHA DO BD
		chave = gera_chave(FATOR_BD)
		decodifica_dado SENHA_BD_AT, senha_decodificada, chave
		s = "Provider=SQLOLEDB;" & _
			"Data Source=" & SERVIDOR_BD_AT & ";" & _
			"Initial Catalog=" & NOME_BD_AT & ";" & _
			"User ID=" & USUARIO_BD_AT & ";" & _
			"Password=" & senha_decodificada & ";"
		end if

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionTimeout = 45
	cn.CommandTimeout = 900
	cn.ConnectionString = s
	cn.Open

	If Err <> 0 then 
		cn.Close
		set cn = nothing
		exit function
		end if

	bdd_AT_conecta = True
	
end function



' ___________________________
' B D D _ B S _ C O N E C T A
'
function bdd_BS_conecta ( cn )
dim s
dim chave
dim senha_decodificada

	bdd_BS_conecta = False
	
	if is_sgbd_access then
		s = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
			"Data Source=" & NOME_BD_BS & ";" & _
			"User ID=" & USUARIO_BD_BS & ";" & _
			"Password=" & SENHA_BD_BS & ";"
	else	
	'   DECODIFICA SENHA DO BD
		chave = gera_chave(FATOR_BD)
		decodifica_dado SENHA_BD_BS, senha_decodificada, chave
		s = "Provider=SQLOLEDB;" & _
			"Data Source=" & SERVIDOR_BD_BS & ";" & _
			"Initial Catalog=" & NOME_BD_BS & ";" & _
			"User ID=" & USUARIO_BD_BS & ";" & _
			"Password=" & senha_decodificada & ";"
		end if

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionTimeout = 45
	cn.CommandTimeout = 900
	cn.ConnectionString = s
	cn.Open

	If Err <> 0 then 
		cn.Close
		set cn = nothing
		exit function
		end if

	bdd_BS_conecta = True
	
end function



' ___________________________________________
' CRIA RECORDSET OTIMISTA
'
function cria_recordset_otimista(byref r, byref msg_erro)
	cria_recordset_otimista = False
	msg_erro = ""
	set r = Server.CreateObject("ADODB.Recordset")
	r.CursorLocation = 3	'adUseClient = 3 (IMPORTANTE para UPDATE com SQLOLEDB)
	r.LockType = 3			'adLockOtimistic = 3 
	r.CacheSize = 30
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if
	cria_recordset_otimista = True
end function



' ___________________________________________
' CRIA RECORDSET PESSIMISTA
'
function cria_recordset_pessimista(byref r, byref msg_erro)
	cria_recordset_pessimista = False
	msg_erro = ""
	set r = Server.CreateObject("ADODB.Recordset")
	r.CursorLocation = 3	'adUseClient = 3  (IMPORTANTE para UPDATE com SQLOLEDB)
	r.CursorType = 2		'adOpenDynamic = 2
	r.LockType = 2			'adLockPessimistic = 2 (LOCK DO REGISTRO)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if
	cria_recordset_pessimista = True
end function



' ___________________________________________
' CRIA RECORDSET SERVERSIDE
' CursorLocationEnum Values:
'	adUseNone = 1 - OBSOLETE (appears only for backward compatibility). Does not use cursor services
'	adUseServer = 2 - Default. Uses a server-side cursor
'	adUseClient = 3 - Uses a client-side cursor supplied by a local cursor library. For backward compatibility, the synonym adUseClientBatch is also supported
' CursorTypeEnum Values:
'	adOpenUnspecified = -1 - Does not specify the type of cursor.
'	adOpenForwardOnly = 0 - Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
'	adOpenKeyset = 1 - Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
'	adOpenDynamic = 2 - Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
'	adOpenStatic = 3 - Uses a static cursor. A static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
' LockTypeEnum Values:
'	adLockUnspecified = -1 - Unspecified type of lock. Clones inherits lock type from the original Recordset.
'	adLockReadOnly = 1 - Read-only records
'	adLockPessimistic = 2 - Pessimistic locking, record by record. The provider lock records immediately after editing
'	adLockOptimistic = 3 - Optimistic locking, record by record. The provider lock records only when calling update
'	adLockBatchOptimistic = 4 - Optimistic batch updates. Required for batch update mode
function cria_recordset_serverside(byref r, byref msg_erro)
	cria_recordset_serverside = False
	msg_erro = ""
	set r = Server.CreateObject("ADODB.Recordset")
	r.CursorLocation = 2	'adUseServer = 2  (IMPORTANTE: NÃO USAR para UPDATE com SQLOLEDB)
	r.CursorType = 2		'adOpenDynamic = 2
	r.LockType = 2			'adLockPessimistic = 2 (LOCK DO REGISTRO)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if
	cria_recordset_serverside = True
end function



' _______________________________
' IS RESTRICAO ATIVA FORMA PAGTO
'
function is_restricao_ativa_forma_pagto(byval id_orcamentista_e_indicador, byval id_forma_pagto, byval tipo_cliente)
dim r
dim s_sql
	s_sql = "SELECT " & _
				"*" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO" & _
			" WHERE" & _
				" (id_orcamentista_e_indicador = '" & Trim(id_orcamentista_e_indicador) & "')" & _
				" AND (id_forma_pagto = " & id_forma_pagto & ")" & _
				" AND (tipo_cliente = '" & Trim(tipo_cliente) & "')" & _
				" AND (st_restricao_ativa <> 0)"
	set r = cn.Execute(s_sql)
	if Not r.Eof then
		is_restricao_ativa_forma_pagto = True
	else
		is_restricao_ativa_forma_pagto = False
		end if
	
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________
' PRODUTO DESCRICAO
'
function produto_descricao(byval f, byval p)
dim r
dim s
dim s_sql
	f = Trim("" & f)
	p = Trim("" & p)
	s_sql = "SELECT fabricante, produto, ean, descricao FROM t_PRODUTO WHERE"
	if IsEAN(p) then
		s_sql = s_sql & " (ean='" & p & "')"
	else
		s_sql = s_sql & " (fabricante='" & f & "') AND (produto='" & p & "')"
		end if
	s = ""
	set r = cn.Execute(s_sql)
	if Not r.Eof then
		s = Trim("" & r("descricao"))
		if IsEAN(p) And (f<>"") then
			if f<>Trim("" & r("fabricante")) then s="PRODUTO " & p & " NÃO PERTENCE AO FABRICANTE " & f
			end if
	else
		s = "NÃO CADASTRADO"
		end if
	produto_descricao = s
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________
' PRODUTO DESCRICAO HTML
'
function produto_descricao_html(byval f, byval p)
dim r
dim s
dim s_sql
	f = Trim("" & f)
	p = Trim("" & p)
	s_sql = "SELECT fabricante, produto, ean, descricao, descricao_html FROM t_PRODUTO WHERE"
	if IsEAN(p) then
		s_sql = s_sql & " (ean='" & p & "')"
	else
		s_sql = s_sql & " (fabricante='" & f & "') AND (produto='" & p & "')"
		end if
	s = ""
	set r = cn.Execute(s_sql)
	if Not r.Eof then
		s = Trim("" & r("descricao_html"))
		if IsEAN(p) And (f<>"") then
			if f<>Trim("" & r("fabricante")) then s="PRODUTO " & p & " NÃO PERTENCE AO FABRICANTE " & f
			end if
	else
		s = "NÃO CADASTRADO"
		end if
	produto_descricao_html = s
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________
' FABRICANTE DESCRICAO
'
function fabricante_descricao(byval f)
dim r
dim s
	f = Trim("" & f)
	s = ""
	set r = cn.Execute("select nome, razao_social from t_FABRICANTE where fabricante = '" & f & "'")
	if Not r.Eof then
		s = Trim("" & r("razao_social"))
		if s = "" then s = Trim("" & r("nome"))
	else
		s = "NÃO CADASTRADO"
		end if
	fabricante_descricao = s
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________
' X _ F A B R I C A N T E
'
function x_fabricante (byval f)
dim r, s
	f = Trim("" & f)
	s = ""
	set r = cn.Execute("select nome, razao_social from t_FABRICANTE where fabricante = '" & f & "'")
	if not r.eof then 
		s = Trim("" & r("nome"))
		if s = "" then s = Trim("" & r("razao_social"))
		end if
	x_fabricante = s
	if r.State <> 0 then r.Close
	set r = nothing
end function



' ___________________________________
' X _ F A B R I C A N T E _ E _ C O D
'
function x_fabricante_e_cod (byval f)
dim s
	f = Trim("" & f)
	s = x_fabricante(f)
	if s <> "" then x_fabricante_e_cod = s & " (" & f & ")"
end function



' _____________________________________
' X _ L O J A _ B D
'
function x_loja_bd(byval l, byref rl)
dim r
	x_loja_bd = False
	l = Trim("" & l)
	set r = cn.Execute("select * from t_LOJA where loja = '" & l & "'")
	if not r.Eof then 
		x_loja_bd = True
		with rl
			.loja = Trim("" & r("loja"))
			.cnpj = Trim("" & r("cnpj"))
			.ie = Trim("" & r("ie"))
			.nome = Trim("" & r("nome"))
			.razao_social = Trim("" & r("razao_social"))
			.endereco = Trim("" & r("endereco"))
			.endereco_numero = Trim("" & r("endereco_numero"))
			.endereco_complemento = Trim("" & r("endereco_complemento"))
			.bairro = Trim("" & r("bairro"))
			.cidade = Trim("" & r("cidade"))
			.uf = Trim("" & r("uf"))
			.cep = Trim("" & r("cep"))
			.ddd = Trim("" & r("ddd"))
			.telefone = Trim("" & r("telefone"))
			.fax = Trim("" & r("fax"))
			.dt_cadastro = r("dt_cadastro")
			.dt_ult_atualizacao = r("dt_ult_atualizacao")
			.comissao_indicacao = r("comissao_indicacao")
			.PercMaxSenhaDesconto = r("PercMaxSenhaDesconto")
			.PercMaxDescSemZerarRT = r("PercMaxDescSemZerarRT")
			.unidade_negocio = Trim("" & r("unidade_negocio"))
			.id_plano_contas_empresa = r("id_plano_contas_empresa")
			.id_plano_contas_grupo = r("id_plano_contas_grupo")
			.id_plano_contas_conta = r("id_plano_contas_conta")
			.natureza = Trim("" & r("natureza"))
			.unidade_negocio = Trim("" & r("unidade_negocio"))
			.perc_max_comissao = r("perc_max_comissao")
			.perc_max_comissao_e_desconto = r("perc_max_comissao_e_desconto")
			.perc_max_comissao_e_desconto_nivel2 = r("perc_max_comissao_e_desconto_nivel2")
			.perc_max_comissao_e_desconto_nivel2_pj = r("perc_max_comissao_e_desconto_nivel2_pj")
			.perc_max_comissao_e_desconto_pj = r("perc_max_comissao_e_desconto_pj")
			.perc_max_comissao_e_desconto_alcada1_pf = r("perc_max_comissao_e_desconto_alcada1_pf")
			.perc_max_comissao_e_desconto_alcada1_pj = r("perc_max_comissao_e_desconto_alcada1_pj")
			.perc_max_comissao_e_desconto_alcada2_pf = r("perc_max_comissao_e_desconto_alcada2_pf")
			.perc_max_comissao_e_desconto_alcada2_pj = r("perc_max_comissao_e_desconto_alcada2_pj")
			.perc_max_comissao_e_desconto_alcada3_pf = r("perc_max_comissao_e_desconto_alcada3_pf")
			.perc_max_comissao_e_desconto_alcada3_pj = r("perc_max_comissao_e_desconto_alcada3_pj")
			.magento_api_urlWebService = Trim("" & r("magento_api_urlWebService"))
			.magento_api_username = Trim("" & r("magento_api_username"))
			.magento_api_password = Trim("" & r("magento_api_password"))
			.magento_api_versao = r("magento_api_versao")
			.magento_api_rest_endpoint = Trim("" & r("magento_api_rest_endpoint"))
			.magento_api_rest_access_token = Trim("" & r("magento_api_rest_access_token"))
			.magento_api_rest_force_get_sales_order_by_entity_id = r("magento_api_rest_force_get_sales_order_by_entity_id")
			end with
		end if

	if r.State <> 0 then r.Close
	set r = nothing
end function



' ___________
' X _ L O J A
'
function x_loja (byval l)
dim r
	l = Trim("" & l)
	set r = cn.Execute("select nome from t_LOJA where loja = '" & l & "'")
	if not r.eof then x_loja = Trim("" & r("nome"))
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________
' X _ L O J A _ E _ C O D
'
function x_loja_e_cod (byval l)
dim s
	l = Trim("" & l)
	s = x_loja(l)
	if s <> "" then x_loja_e_cod = s & " (" & l & ")"
end function



' _______________________________
' X _ T R A N S P O R T A D O R A
'
function x_transportadora (byval c)
dim r
	c = Trim("" & c)
	set r = cn.Execute("select nome from t_TRANSPORTADORA where id = '" & c & "'")
	if not r.eof then x_transportadora = Trim("" & r("nome"))
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _________________
' X _ U S U A R I O
'
function x_usuario (byval u)
dim r
	u = Trim("" & u)
	set r = cn.Execute("select nome_iniciais_em_maiusculas from t_USUARIO where usuario = '" & u & "'")
	if not r.eof then x_usuario = Trim("" & r("nome_iniciais_em_maiusculas"))
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _____________________________________________
' X _ C L I E N T E _ P E D I D O
'
function x_cliente_pedido (byval pedido, byref cadastrado)
dim r, s_id_cliente
	x_cliente_pedido = ""
	cadastrado = False
	pedido = normaliza_num_pedido(Trim("" & pedido))
	set r = cn.Execute("SELECT id_cliente FROM t_PEDIDO WHERE pedido = '" & pedido & "'")
	if r.Eof then exit function
	
	s_id_cliente = Trim("" & r("id_cliente"))
	if r.State <> 0 then r.Close
	set r = nothing
	set r = cn.Execute("select nome_iniciais_em_maiusculas from t_CLIENTE where (id = '" & s_id_cliente & "')")
	if not r.eof then 
		x_cliente_pedido = Trim("" & r("nome_iniciais_em_maiusculas"))
		cadastrado = True
		end if
	if r.State <> 0 then r.Close
	set r = nothing
end function




' _____________________________________________
' X _ C L I E N T E _ P O R _ C N P J _ C P F
'
function x_cliente_por_cnpj_cpf (byval cnpj_cpf, byref cadastrado)
dim r
	x_cliente_por_cnpj_cpf = ""
	cadastrado = False
	cnpj_cpf = retorna_so_digitos(Trim("" & cnpj_cpf))
	set r = cn.Execute("select nome_iniciais_em_maiusculas from t_CLIENTE where (cnpj_cpf = '" & cnpj_cpf & "')")
	if not r.eof then 
		x_cliente_por_cnpj_cpf = Trim("" & r("nome_iniciais_em_maiusculas"))
		cadastrado = True
		end if
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _____________________________________
' X _ C L I E N T E _ B D
'
function x_cliente_bd(byval c, byref rc)
dim r
	x_cliente_bd = False
	c = Trim("" & c)
	set r = cn.Execute("select * from t_CLIENTE where id = '" & c & "'")
	if not r.Eof then 
		x_cliente_bd = True
		with rc
			.id = Trim("" & r("id"))
			.cnpj_cpf = Trim("" & r("cnpj_cpf"))
			.tipo = Trim("" & r("tipo"))
			.ie = Trim("" & r("ie"))
			.produtor_rural_status = Trim("" & r("produtor_rural_status"))
			.contribuinte_icms_status = Trim("" & r("contribuinte_icms_status"))
			.rg = Trim("" & r("rg"))
			.nome = Trim("" & r("nome"))
			.nome_iniciais_em_maiusculas = Trim("" & r("nome_iniciais_em_maiusculas"))
			.sexo = Trim("" & r("sexo"))
			.endereco = Trim("" & r("endereco"))
			.endereco_numero = Trim("" & r("endereco_numero"))
			.endereco_complemento = Trim("" & r("endereco_complemento"))
			.bairro = Trim("" & r("bairro"))
			.cidade = Trim("" & r("cidade"))
			.uf = Trim("" & r("uf"))
			.cep = Trim("" & r("cep"))
			.ddd_res = Trim("" & r("ddd_res"))
			.tel_res = Trim("" & r("tel_res"))
			.ddd_com = Trim("" & r("ddd_com"))
			.tel_com = Trim("" & r("tel_com"))
			.ramal_com = Trim("" & r("ramal_com"))
			.contato = Trim("" & r("contato"))
			.dt_nasc = r("dt_nasc")
			.filiacao = Trim("" & r("filiacao"))
			.obs_crediticias = Trim("" & r("obs_crediticias"))
			.midia = Trim("" & r("midia"))
			.email = Trim("" & r("email"))
			.email_opcoes = Trim("" & r("email_opcoes"))
			.email_xml = Trim("" & r("email_xml"))
			.dt_cadastro = r("dt_cadastro")
			.dt_ult_atualizacao = r("dt_ult_atualizacao")
			.SocMaj_Nome = Trim("" & r("SocMaj_Nome"))
			.SocMaj_CPF = Trim("" & r("SocMaj_CPF"))
			.SocMaj_banco = Trim("" & r("SocMaj_banco"))
			.SocMaj_agencia = Trim("" & r("SocMaj_agencia"))
			.SocMaj_conta = Trim("" & r("SocMaj_conta"))
			.SocMaj_ddd = Trim("" & r("SocMaj_ddd"))
			.SocMaj_telefone = Trim("" & r("SocMaj_telefone"))
			.SocMaj_contato = Trim("" & r("SocMaj_contato"))
			.usuario_cadastro = Trim("" & r("usuario_cadastro"))
			.usuario_ult_atualizacao = Trim("" & r("usuario_ult_atualizacao"))
			.indicador = Trim("" & r("indicador"))
			.ddd_cel = Trim("" & r("ddd_cel"))
			.tel_cel = Trim("" & r("tel_cel"))
			.ddd_com_2 = Trim("" & r("ddd_com_2"))
			.tel_com_2 = Trim("" & r("tel_com_2"))
			.ramal_com_2 = Trim("" & r("ramal_com_2"))
			end with
		end if

	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________
' X _ M I D I A
'
function x_midia (byval c)
dim r, s
	c = Trim("" & c)
	s = ""
	set r = cn.Execute("SELECT id, apelido FROM t_MIDIA WHERE (id = '" & c & "')")
	if not r.Eof then 
		s = Trim("" & r("apelido"))
		end if
	x_midia = s
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _________________________________
' X _ P E R F I L _ A P E L I D O
'
function x_perfil_apelido (byval id)
dim r
	id = Trim("" & id)
	set r = cn.Execute("select apelido from t_PERFIL where id = '" & id & "'")
	if not r.eof then x_perfil_apelido = Trim("" & r("apelido"))
	if r.State <> 0 then r.Close
	set r = nothing
end function



' ___________________________________________
' OBTEM NIVEL ACESSO BLOCO NOTAS PEDIDO
'
function obtem_nivel_acesso_bloco_notas_pedido(ByRef cnBancoDados, byval usuario)
dim r
dim s, s_aux, cod_resp
	cod_resp = COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__NAO_DEFINIDO
	s="SELECT Coalesce(Max(nivel_acesso_bloco_notas_pedido), " & Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__NAO_DEFINIDO) & ") AS max_nivel_acesso_bloco_notas_pedido FROM t_PERFIL INNER JOIN t_PERFIL_X_USUARIO ON t_PERFIL.id=t_PERFIL_X_USUARIO.id_perfil WHERE (usuario = '" & usuario & "')"
	set r = cnBancoDados.Execute(s)
	if Not r.Eof then
		cod_resp = r("max_nivel_acesso_bloco_notas_pedido")
		end if
	obtem_nivel_acesso_bloco_notas_pedido=cod_resp
end function


' ___________________________________________
' OBTEM NIVEL ACESSO CHAMADO PEDIDO
'
function obtem_nivel_acesso_chamado_pedido(ByRef cnBancoDados, byval usuario)
dim r
dim s, s_aux, cod_resp
	cod_resp = COD_NIVEL_ACESSO_CHAMADO_PEDIDO__NAO_DEFINIDO
	s="SELECT Coalesce(Max(nivel_acesso_chamado), " & Cstr(COD_NIVEL_ACESSO_CHAMADO_PEDIDO__NAO_DEFINIDO) & ") AS max_nivel_acesso_chamado FROM t_PERFIL INNER JOIN t_PERFIL_X_USUARIO ON t_PERFIL.id=t_PERFIL_X_USUARIO.id_perfil WHERE (usuario = '" & usuario & "')"
	set r = cnBancoDados.Execute(s)
	if Not r.Eof then
		cod_resp = r("max_nivel_acesso_chamado")
		end if
	obtem_nivel_acesso_chamado_pedido=cod_resp
end function


' ___________________________________________
' OBTEM OPERACOES PERMITIDAS USUARIO
'
function obtem_operacoes_permitidas_usuario(ByRef cnBancoDados, byval usuario)
dim r
dim s, s_aux, s_resp
	s_resp=""
	s="SELECT DISTINCT " & _
			" id_operacao" & _
		" FROM t_PERFIL" & _
			" INNER JOIN t_PERFIL_ITEM ON t_PERFIL.id=t_PERFIL_ITEM.id_perfil" & _
			" INNER JOIN t_PERFIL_X_USUARIO ON t_PERFIL.id=t_PERFIL_X_USUARIO.id_perfil" & _
			" INNER JOIN t_OPERACAO ON (t_PERFIL_ITEM.id_operacao=t_OPERACAO.id)" & _
		" WHERE" & _
			" (usuario = '" & usuario & "')" & _
			" AND (t_PERFIL.st_inativo = 0)" & _
			" AND (t_OPERACAO.st_inativo = 0)" & _
		" ORDER BY" & _
			" id_operacao"
	set r = cnBancoDados.Execute(s)
	do while Not r.Eof
		s = Trim(Cstr(r("id_operacao")))
		if s <> "" then
			s_aux = "|" & s & "|"
			if Instr(s_resp, s_aux) = 0 then
				if Right(s_resp,1) <> "|" then s_resp = s_resp & "|"
				s_resp = s_resp & s
				end if
			end if
		r.MoveNext
		loop
	if (s_resp <> "") And (Right(s_resp,1) <> "|") then s_resp = s_resp & "|"
	obtem_operacoes_permitidas_usuario=s_resp
end function



' ___________________________________________
' OPERACAO PERMITIDA
'
function operacao_permitida(byval id_operacao, byval lista_operacoes_permitidas)
dim s
	operacao_permitida = False
	s = Trim(Cstr(id_operacao))
	if s = "" then exit function
	s = "|" & s & "|"
	if Instr(lista_operacoes_permitidas, s) > 0 then operacao_permitida = True
end function



' ___________________________________________
' GRAVA LOG
'
function grava_log(byval usuario, byval loja, byval pedido, byval id_cliente, byval operacao, byval complemento)
dim s, rs
dim msg_erro
	grava_log = False
	if Err <> 0 then Err.Clear
	
	if Not cria_recordset_otimista(rs, msg_erro) then exit function
	
	if is_sgbd_access then
		s = "select * from t_LOG where (data < CDate('30/12/1899'))"
	else
		s = "select * from t_LOG where (data < '" & BD_DATA_NULA & "')"
		end if
		
	rs.Open s, cn
	if Err = 0 then 
		rs.AddNew
		if Err = 0 then
		  ' LEMBRANDO QUE A DATA É INSERIDA, VIA DEFAULT DA COLUNA, COM O VALOR DE getdate()
			rs("usuario")=usuario
			rs("loja")=loja
			rs("pedido")=pedido
			rs("id_cliente")=id_cliente
			rs("operacao")=operacao
			rs("complemento")=complemento
			rs.update
			if Err = 0 then grava_log = True
			end if
		end if
	if rs.State <> 0 then rs.Close
	set rs=nothing
end function



' ___________________________________________
' GRAVA LOG ESTOQUE V2
'
function grava_log_estoque_v2(byval strUsuario, byval id_nfe_emitente, byval strFabricante, byval strProduto, byval intQtdeSolicitada, byval intQtdeAtendida, byval strOperacao, byval strCodEstoqueOrigem, byval strCodEstoqueDestino, byval strLojaEstoqueOrigem, byval strLojaEstoqueDestino, byval strPedidoEstoqueOrigem, byval strPedidoEstoqueDestino, byval strDocumento, byval strComplemento, byval strIdOrdemServico)
dim s, rs
dim msg_erro
	grava_log_estoque_v2 = False
	if Not cria_recordset_otimista(rs, msg_erro) then exit function
	
	s = "select * from t_ESTOQUE_LOG where (data < '" & BD_DATA_NULA & "')"
	rs.Open s, cn
	if Err = 0 then 
		rs.AddNew
		if Err = 0 then
			rs("data")=Date
			rs("data_hora")=Now
			rs("usuario")=strUsuario
			rs("id_nfe_emitente")=id_nfe_emitente
			rs("fabricante")=strFabricante
			rs("produto")=strProduto
			rs("qtde_solicitada")=intQtdeSolicitada
			rs("qtde_atendida")=intQtdeAtendida
			rs("operacao")=strOperacao
			rs("cod_estoque_origem")=strCodEstoqueOrigem
			rs("cod_estoque_destino")=strCodEstoqueDestino
			rs("loja_estoque_origem")=strLojaEstoqueOrigem
			rs("loja_estoque_destino")=strLojaEstoqueDestino
			rs("pedido_estoque_origem")=strPedidoEstoqueOrigem
			rs("pedido_estoque_destino")=strPedidoEstoqueDestino
			rs("documento")=strDocumento
			rs("complemento")=Left(strComplemento, 80)
			rs("id_ordem_servico") = strIdOrdemServico
			rs.update
			if Err = 0 then grava_log_estoque_v2 = True
			end if
		end if
	if rs.State <> 0 then rs.Close
	set rs=nothing
end function



' ___________________________________________
' LOG VIA VETOR CARREGA DO RECORDSET
'
function log_via_vetor_carrega_do_recordset(rs, v, byval campos_a_omitir)
const LISTA_CAMPOS_A_OMITIR = "|TIMESTAMP|SENHA|"
dim i
dim s
dim s_campos_a_omitir

	redim v(0)
	set v(0) = new cl_LOG_VIA_VETOR

	if rs.Eof Or rs.Bof then exit function
		
	s_campos_a_omitir=LISTA_CAMPOS_A_OMITIR
	campos_a_omitir = Ucase(Trim("" & campos_a_omitir))
	if Trim(campos_a_omitir)<>"" then
		if Left(campos_a_omitir,1) <> "|" then s_campos_a_omitir = s_campos_a_omitir & "|"
		s_campos_a_omitir = s_campos_a_omitir & campos_a_omitir
		if Right(s_campos_a_omitir,1) <> "|" then s_campos_a_omitir = s_campos_a_omitir & "|"
		end if

	for i=0 to rs.Fields.Count-1
	'	IGNORA CAMPOS TIMESTAMP (PRINCIPALMENTE PORQUE CAUSAM ERRO NA CONVERSÃO PARA STRING)	
		s = "|" & Ucase(rs.Fields(i).Name) & "|"
		if InStr(1, s_campos_a_omitir, s) = 0 then
			if Trim(v(ubound(v)).nome)<>"" then 
				redim preserve v(ubound(v)+1)
				set v(ubound(v)) = new cl_LOG_VIA_VETOR
				end if

			v(ubound(v)).nome=rs.Fields(i).Name
			if IsNull(rs.Fields(i).Value) then s = "" else s = rs.Fields(i).Value
			v(ubound(v)).valor = s
			end if
		next
end function



' ___________________________________________
' LOG VIA VETOR MONTA ALTERACAO
'
function log_via_vetor_monta_alteracao(v1, v2)
dim i
dim s
dim s_aux
	if Lbound(v1) <> Lbound(v2) then exit function
	if Ubound(v1) <> Ubound(v2) then exit function
	
	s=""
	for i=Lbound(v1) to Ubound(v1)
		if (Trim(v1(i).nome) <> "") And (Trim(v1(i).nome)=Trim(v2(i).nome)) then
			if Trim(v1(i).valor) <> Trim(v2(i).valor) then
				if s <> "" then s = s & "; "
				s_aux=Trim(v1(i).valor)
				if s_aux = "" then s_aux = chr(34) & chr(34)
				s = s & Trim(v1(i).nome) & ": " & s_aux & " => "
				s_aux = Trim(v2(i).valor)
				if s_aux = "" then s_aux = chr(34) & chr(34)
				s = s & s_aux
				end if
			end if
		next
	log_via_vetor_monta_alteracao = s
end function



' ___________________________________________
' LOG VIA VETOR MONTA INCLUSAO
'
function log_via_vetor_monta_inclusao(v)
dim i
dim s
dim s_aux
	s=""
	for i=Lbound(v) to Ubound(v)
		if (Trim(v(i).nome) <> "") then
			if s <> "" then s = s & "; "
			s_aux = Trim(v(i).valor)
			if s_aux = "" then s_aux = chr(34) & chr(34)
			s = s & Trim(v(i).nome) & "=" & s_aux
			end if
		next
	log_via_vetor_monta_inclusao = s
end function



' ___________________________________________
' LOG VIA VETOR MONTA EXCLUSAO
'
function log_via_vetor_monta_exclusao(v)
dim i
dim s
dim s_aux
	s=""
	for i=Lbound(v) to Ubound(v)
		if (Trim(v(i).nome) <> "") then
			if s <> "" then s = s & "; "
			s_aux = Trim(v(i).valor)
			if s_aux = "" then s_aux = chr(34) & chr(34)
			s = s & Trim(v(i).nome) & "=" & s_aux
			end if
		next
	log_via_vetor_monta_exclusao = s
end function



' ___________________________________________
' RECUPERA AVISOS NAO LIDOS
'
function recupera_avisos_nao_lidos(byval loja, byval usuario, v)
dim s, s_where, rs

	recupera_avisos_nao_lidos=False
	
	redim v(0)
	set v(0) = new cl_QUADRO_AVISOS

	loja = Trim("" & loja)
	if loja <> "" then loja = normaliza_codigo(loja, TAM_MIN_LOJA)
	
	s = ""
	
	if is_sgbd_access then
		if loja <> "" then s = " OR (destinatario='" & loja & "')"
		s_where = " ((destinatario='') OR (destinatario IS NULL)" & s & ")"

		s = "SELECT * FROM t_AVISO WHERE (id NOT IN (SELECT id FROM t_AVISO_LIDO WHERE usuario='" & usuario & "'))" & _
			" AND " & s_where & _
			" ORDER BY dt_ult_atualizacao DESC"

		set rs = cn.Execute(s)
		if Err <> 0 then exit function

		do while Not rs.EOF
		'   ESTA MENSAGEM AINDA NÃO FOI LIDA
			if Not IsNull(rs("mensagem")) then 
				if Trim(v(ubound(v)).id_aviso)<>"" then
					redim preserve v(ubound(v)+1)
					set v(ubound(v)) = new cl_QUADRO_AVISOS
					end if
				v(ubound(v)).id_aviso = Trim("" & rs("id"))
				v(ubound(v)).mensagem = rs("mensagem")
				v(ubound(v)).dt_ult_atualizacao = rs("dt_ult_atualizacao")
				v(ubound(v)).usuario = Trim("" & rs("usuario"))
				v(ubound(v)).lido = ""
				v(ubound(v)).dt_lido = Null
				recupera_avisos_nao_lidos = True
				end if
			
			rs.MoveNext 
			loop
			
	else
		if loja <> "" then s = " OR (destinatario='" & loja & "')"
		s_where = " WHERE ((destinatario='') OR (destinatario IS NULL)" & s & ")"

		s = "SELECT t_AVISO.*, t_AVISO_LIDO.id AS id_aviso_lido" & _
			" FROM t_AVISO LEFT JOIN t_AVISO_LIDO ON" & _
			" (t_AVISO.id = t_AVISO_LIDO.id) AND" & _
			" ('" & usuario & "' = t_AVISO_LIDO.usuario)" & _
			s_where & _
			" ORDER BY dt_ult_atualizacao DESC"
		
		set rs = cn.Execute(s)
		if Err <> 0 then exit function

		do while Not rs.EOF
		'   ESTA MENSAGEM AINDA NÃO FOI LIDA
			if Trim("" & rs("id_aviso_lido")) = "" then
				if Not IsNull(rs("mensagem")) then 
					if Trim(v(ubound(v)).id_aviso)<>"" then
						redim preserve v(ubound(v)+1)
						set v(ubound(v)) = new cl_QUADRO_AVISOS
						end if
					v(ubound(v)).id_aviso = Trim("" & rs("id"))
					v(ubound(v)).mensagem = rs("mensagem")
					v(ubound(v)).dt_ult_atualizacao = rs("dt_ult_atualizacao")
					v(ubound(v)).usuario = Trim("" & rs("usuario"))
					v(ubound(v)).lido = ""
					v(ubound(v)).dt_lido = Null
					recupera_avisos_nao_lidos = True
					end if
				end if
			
			rs.MoveNext 
			loop
		
		end if

	if rs.State <> 0 then rs.Close
	set rs = nothing
	
end function



' ___________________________________________
' RECUPERA AVISOS
'
function recupera_avisos(byval loja, byval usuario, v)
dim s, s_where, rs

	recupera_avisos=False
	
	redim v(0)
	set v(0) = new cl_QUADRO_AVISOS

	loja = Trim("" & loja)
	if loja <> "" then loja = normaliza_codigo(loja, TAM_MIN_LOJA)
	
	s = ""
	if loja <> "" then s = " OR (destinatario='" & loja & "')"
	s_where = " WHERE ((destinatario='') OR (destinatario IS NULL)" & s & ")"

	s = "SELECT t_AVISO.*, t_AVISO_LIDO.id AS id_aviso_lido, t_AVISO_LIDO.data AS data_leitura" & _
		" FROM t_AVISO LEFT JOIN t_AVISO_LIDO ON" & _
		" (t_AVISO.id = t_AVISO_LIDO.id) AND" & _
		" ('" & usuario & "' = t_AVISO_LIDO.usuario)" & _
		s_where & _
		" ORDER BY dt_ult_atualizacao DESC"

	set rs = cn.Execute(s)
	if Err <> 0 then exit function

	do while Not rs.EOF
		if Not IsNull(rs("mensagem")) then 
			recupera_avisos = True

			if Trim(v(ubound(v)).id_aviso)<>"" then
				redim preserve v(ubound(v)+1)
				set v(ubound(v)) = new cl_QUADRO_AVISOS
				end if
			
			v(ubound(v)).id_aviso = Trim("" & rs("id"))
			v(ubound(v)).mensagem = rs("mensagem")
			v(ubound(v)).dt_ult_atualizacao = rs("dt_ult_atualizacao")
			v(ubound(v)).usuario = Trim("" & rs("usuario"))

		'   JÁ FOI LIDO?
			if Trim("" & rs("id_aviso_lido")) = "" then
				v(ubound(v)).lido = ""
				v(ubound(v)).dt_lido = Null
			else
				v(ubound(v)).lido = "S"
				v(ubound(v)).dt_lido = rs("data_leitura")
				end if
			end if
		
		rs.MoveNext 
		loop

	if rs.State <> 0 then rs.Close
	set rs = nothing
end function



' ___________________________________________
' GERA NSU
'
function gera_nsu(byval id_nsu, byref nsu_novo, byref msg_erro)
dim r, s, n_nsu, s_nsu, n_tentativas, update_OK

	On Error Resume Next
	
	gera_nsu = False
	nsu_novo = ""
	msg_erro = ""
	
	id_nsu = Trim("" & id_nsu)
	if id_nsu = "" then 
		msg_erro = "Não foi especificado o NSU a ser gerado!!"
		exit function
		end if
	
'	JÁ ESTÁ EM SITUAÇÃO DE ERRO?
	if Err <> 0 then 
		msg_erro = "Não foi possível gerar o NSU, pois há erro pendente: " & Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
'   OBTEM PROXIMO NSU EM T_CONTROLE
	n_tentativas = 0
	update_OK = False
	do while Not update_OK
		n_tentativas = n_tentativas + 1
	 '	MAIS DE 100 TENTATIVAS?
		if n_tentativas > 100 then 
			msg_erro = "Não foi possível gerar o NSU. " & msg_erro
			exit function
			end if

		Err.Clear 
		msg_erro = ""
		
		if TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO then
		'	BLOQUEIA REGISTRO PARA EVITAR ACESSO CONCORRENTE (REALIZA O FLIP EM UM CAMPO BIT APENAS P/ ADQUIRIR O LOCK EXCLUSIVO)
			s = "UPDATE t_CONTROLE SET dummy = ~dummy WHERE id_nsu='" & id_nsu & "'"
			cn.Execute(s)
			end if

		n_nsu = -1
		s = "SELECT * FROM t_CONTROLE WHERE id_nsu='" & id_nsu & "'"
		if Not cria_recordset_pessimista(r, msg_erro) then exit function
		r.Open s, cn

	'	AINDA NÃO EXISTE REGISTRO PARA ESTE NSU
		if r.EOF then
			if r.State <> 0 then r.Close 
			set r = nothing
			msg_erro = "Não existe registro na tabela de controle para poder gerar este NSU!!" 
			exit function
			end if

		if Not IsNull(r("nsu")) then
			if IsNumeric(r("nsu")) then 
				if r("seq_anual") <> 0 then
				'	CASO O RELÓGIO DO SERVIDOR SEJA ALTERADO P/ DATAS FUTURAS E PASSADAS, EVITA QUE O CAMPO 'ano_letra_seq' SEJA INCREMENTADO VÁRIAS VEZES
					if Year(Date) > Year(r("dt_ult_atualizacao")) then
						s = "0"
						s=normaliza_codigo(s, TAM_MAX_NSU)
						r("nsu") = s
						r("dt_ult_atualizacao") = Date
						if Trim("" & r("ano_letra_seq")) <> "" then
							r("ano_letra_seq") = Chr(Asc(r("ano_letra_seq")) + r("ano_letra_step"))
							end if
						end if
					end if
					
				n_nsu = CLng(r("nsu"))
				end if
			end if
			
		if n_nsu < 0 then
			if r.State <> 0 then r.Close 
			set r = nothing
			msg_erro = "O NSU gerado é inválido!!"
			exit function
			end if
	
	'   GERA NOVO NÚMERO
		n_nsu = n_nsu + 1
		s_nsu = Cstr(n_nsu)
		s_nsu=normaliza_codigo(s_nsu, TAM_MAX_NSU)

		r("nsu") = s_nsu
	'	Caso o relógio do servidor seja alterado p/ datas futuras e passadas, evita que o campo 'ano_letra_seq' seja incrementado várias vezes através
	'	do controle que impede o campo 'dt_ult_atualizacao' de receber uma data menor do que aquela que ele já possui
		if Trim("" & r("dt_ult_atualizacao")) = "" then
			r("dt_ult_atualizacao") = Date
		else
			if Date > r("dt_ult_atualizacao") then r("dt_ult_atualizacao") = Date
			end if
		r.Update
		
		if Err = 0 then 
			update_OK = True 
		else 
			msg_erro = "Não foi possível gerar o NSU, pois ocorreu o seguinte erro: " & Cstr(Err) & ": " & Err.Description
			end if
		
		if r.State <> 0 then r.Close 
		set r = nothing
		loop
		
	if Err <> 0 then 
		msg_erro = "Não foi possível gerar o NSU, pois ocorreu o seguinte erro: " & Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if Not update_OK then 
		msg_erro = "Falha ao tentar gerar o NSU."
		exit function
		end if

	nsu_novo = s_nsu
	gera_nsu = True
	
end function



' ___________________________________________
' FIN GERA NSU
'
function fin_gera_nsu(byval idNsu, byref nsu, byref msg_erro)
dim t, strSql, intRetorno, intRecordsAffected
dim intQtdeTentativas, intNsuUltimo, intNsuNovo, blnSucesso
	fin_gera_nsu=False
	msg_erro=""
	nsu=0
	strSql = "SELECT" & _
				" Count(*) AS qtde" & _
			" FROM t_FIN_CONTROLE" & _
			" WHERE" & _
				" (id='" & idNsu & "')"
	set t=cn.Execute(strSql)
	if Not t.Eof then intRetorno=Clng(t("qtde")) else intRetorno=Clng(0)

'	NÃO ESTÁ CADASTRADO, ENTÃO CADASTRA AGORA
	if intRetorno=0 then
		strSql = "INSERT INTO t_FIN_CONTROLE (" & _
					"id, " & _
					"nsu, " & _
					"dt_hr_ult_atualizacao" & _
				") VALUES (" & _
					"'" & idNsu & "'," & _
					"0," & _
					"getdate()" & _
				")"
		cn.Execute strSql, intRecordsAffected
		if intRecordsAffected <> 1 then
			msg_erro = "Falha ao criar o registro para geração de NSU (" & idNsu & ")!!"
			exit function
			end if
		end if

'	LAÇO DE TENTATIVAS PARA GERAR O NSU (DEVIDO A ACESSO CONCORRENTE)
	intQtdeTentativas=0
	do 
		intQtdeTentativas = intQtdeTentativas + 1
		
		if TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO then
		'	BLOQUEIA REGISTRO PARA EVITAR ACESSO CONCORRENTE (REALIZA O FLIP EM UM CAMPO BIT APENAS P/ ADQUIRIR O LOCK EXCLUSIVO)
			strSql = "UPDATE t_FIN_CONTROLE SET" & _
						" dummy = ~dummy" & _
					" WHERE" & _
						" id = '" & idNsu & "'"
			cn.Execute(strSql)
			end if

	'	OBTÉM O ÚLTIMO NSU USADO
		strSql = "SELECT" & _
					" nsu" & _
				" FROM t_FIN_CONTROLE" & _
				" WHERE" & _
					" id = '" & idNsu & "'"
		set t=cn.Execute(strSql)
		if t.Eof then
			strMsgErro = "Falha ao localizar o registro para geração de NSU (" & idNsu & ")!!"
			Exit Function
		else
			intNsuUltimo = Clng(t("nsu"))
			end if

	'	INCREMENTA 1
		intNsuNovo = intNsuUltimo + 1
		
	'	TENTA ATUALIZAR O BANCO DE DADOS
		strSql = "UPDATE t_FIN_CONTROLE SET" & _
					" nsu = " & CStr(intNsuNovo) & "," & _
					" dt_hr_ult_atualizacao = getdate()" & _
				" WHERE" & _
					" (id = '" & idNsu & "')" & _
					" AND (nsu = " & CStr(intNsuUltimo) & ")"
		cn.Execute strSql, intRecordsAffected
		If intRecordsAffected = 1 Then
			blnSucesso = True
			nsu = intNsuNovo
			end if
		
		Loop While (Not blnSucesso) And (intQtdeTentativas < 10)

	If Not blnSucesso Then
		strMsgErro = "Falha ao tentar gerar o NSU!!"
		Exit Function
		End If
	
	fin_gera_nsu = True

end function



function le_ano_letra_seq_tabela_controle(byval id_nsu, byref ano_letra_seq, byref msg_erro)
dim r, s
	
	On Error Resume Next
	
	le_ano_letra_seq_tabela_controle = False
	ano_letra_seq = ""
	msg_erro = ""

	s = "SELECT ano_letra_seq FROM t_CONTROLE WHERE id_nsu='" & id_nsu & "'"
	if Not cria_recordset_pessimista(r, msg_erro) then exit function
	r.Open s, cn
	if r.EOF then
		if r.State <> 0 then r.Close
		set r = nothing
		msg_erro = "Não existe registro na tabela de controle com o id = '"& id_nsu & "'"
		exit function
		end if

	ano_letra_seq = Trim("" & r("ano_letra_seq"))
	if r.State <> 0 then r.Close
	set r = nothing

	le_ano_letra_seq_tabela_controle = True
end function



' ___________________________________________
' GERA NUM PEDIDO
'
function gera_num_pedido(byref num_pedido, byref msg_erro)
dim s_num
dim s_letra_ano
dim s_descarte
dim n_descarte
	gera_num_pedido=False
	num_pedido=""
	msg_erro=""
	if Not gera_nsu(NSU_PEDIDO, s_num, msg_erro) then exit function
	n_descarte = len(s_num)-TAM_MIN_NUM_PEDIDO
	s_descarte = Left(s_num, n_descarte)
	if s_descarte <> String(n_descarte, "0") then exit function
	s_num = Right(s_num, TAM_MIN_NUM_PEDIDO)
'	OBTÉM A LETRA PARA O SUFIXO DO PEDIDO DE ACORDO C/ O ANO DA GERAÇÃO DO NSU (IMPORTANTE: FAZER A LEITURA SOMENTE APÓS GERAR O NSU, POIS A LETRA PODE TER SIDO ALTERADA DEVIDO À MUDANÇA DE ANO!!)
	if Not le_ano_letra_seq_tabela_controle(NSU_PEDIDO, s_letra_ano, msg_erro) then exit function
	num_pedido = s_num & s_letra_ano
	gera_num_pedido=True
end function



' ___________________________________________
' GERA NUM PEDIDO TEMP
'
function gera_num_pedido_temp(byref num_pedido, byref msg_erro)
dim s_num
dim s_letra_ano
dim s_descarte
dim n_descarte
	gera_num_pedido_temp=False
	num_pedido=""
	msg_erro=""
	if Not gera_nsu(NSU_PEDIDO_TEMPORARIO, s_num, msg_erro) then exit function
	n_descarte = len(s_num)-TAM_MIN_NUM_PEDIDO
	s_descarte = Left(s_num, n_descarte)
	if s_descarte <> String(n_descarte, "0") then exit function
	s_num = Right(s_num, TAM_MIN_NUM_PEDIDO)
'	OBTÉM A LETRA PARA O SUFIXO DO PEDIDO DE ACORDO C/ O ANO DA GERAÇÃO DO NSU (IMPORTANTE: FAZER A LEITURA SOMENTE APÓS GERAR O NSU, POIS A LETRA PODE TER SIDO ALTERADA DEVIDO À MUDANÇA DE ANO!!)
	if Not le_ano_letra_seq_tabela_controle(NSU_PEDIDO_TEMPORARIO, s_letra_ano, msg_erro) then exit function
	num_pedido = "T" & s_num & s_letra_ano
	gera_num_pedido_temp=True
end function



' ___________________________________________
' GERA NUM ORCAMENTO
'
function gera_num_orcamento(byref num_orcamento, byref msg_erro)
dim s_num
dim s_letra_ano
dim s_descarte
dim n_descarte
	gera_num_orcamento=False
	num_orcamento=""
	msg_erro=""
	if Not gera_nsu(NSU_ORCAMENTO, s_num, msg_erro) then exit function
	n_descarte = len(s_num)-TAM_MAX_NUM_ORCAMENTO
	s_descarte = Left(s_num, n_descarte)
	if s_descarte <> String(n_descarte, "0") then exit function
	s_num = Right(s_num, TAM_MAX_NUM_ORCAMENTO)
	'MANTÉM OS ZEROS À ESQUERDA SOMENTE ATÉ O TAMANHO 'TAM_MIN_NUM_ORCAMENTO', SE A QUANTIDADE DE DÍGITOS EXCEDER 'TAM_MIN_NUM_ORCAMENTO', OS DÍGITOS À ESQUERDA DEVEM SER SIGNIFICATIVOS
	do while ((Left(s_num, 1) = "0") And (Len(s_num) > TAM_MIN_NUM_ORCAMENTO)): s_num = Right(s_num, (Len(s_num)-1)): loop
'	OBTÉM A LETRA PARA O SUFIXO DO PRÉ-PEDIDO DE ACORDO C/ O ANO DA GERAÇÃO DO NSU (IMPORTANTE: FAZER A LEITURA SOMENTE APÓS GERAR O NSU, POIS A LETRA PODE TER SIDO ALTERADA DEVIDO À MUDANÇA DE ANO!!)
	if Not le_ano_letra_seq_tabela_controle(NSU_ORCAMENTO, s_letra_ano, msg_erro) then exit function
	num_orcamento = s_num & s_letra_ano
	gera_num_orcamento=True
end function



' ___________________________________________
' GERA NUM ORCAMENTO TEMP
'
function gera_num_orcamento_temp(byref num_orcamento, byref msg_erro)
dim s_num
dim s_letra_ano
dim s_descarte
dim n_descarte
	gera_num_orcamento_temp=False
	num_orcamento=""
	msg_erro=""
	if Not gera_nsu(NSU_ORCAMENTO_TEMPORARIO, s_num, msg_erro) then exit function
	n_descarte = len(s_num)-TAM_MAX_NUM_ORCAMENTO
	s_descarte = Left(s_num, n_descarte)
	if s_descarte <> String(n_descarte, "0") then exit function
	s_num = Right(s_num, TAM_MAX_NUM_ORCAMENTO)
	'MANTÉM OS ZEROS À ESQUERDA SOMENTE ATÉ O TAMANHO 'TAM_MIN_NUM_ORCAMENTO', SE A QUANTIDADE DE DÍGITOS EXCEDER 'TAM_MIN_NUM_ORCAMENTO', OS DÍGITOS À ESQUERDA DEVEM SER SIGNIFICATIVOS
	do while ((Left(s_num, 1) = "0") And (Len(s_num) > TAM_MIN_NUM_ORCAMENTO)): s_num = Right(s_num, (Len(s_num)-1)): loop
'	OBTÉM A LETRA PARA O SUFIXO DO PRÉ-PEDIDO DE ACORDO C/ O ANO DA GERAÇÃO DO NSU (IMPORTANTE: FAZER A LEITURA SOMENTE APÓS GERAR O NSU, POIS A LETRA PODE TER SIDO ALTERADA DEVIDO À MUDANÇA DE ANO!!)
	if Not le_ano_letra_seq_tabela_controle(NSU_ORCAMENTO_TEMPORARIO, s_letra_ano, msg_erro) then exit function
	num_orcamento = "T" & s_num & s_letra_ano
	gera_num_orcamento_temp=True
end function



' ___________________________________________
' GERA ID ESTOQUE
'
function gera_id_estoque(byref id_estoque, byref msg_erro)
	gera_id_estoque=gera_nsu(NSU_ID_ESTOQUE, id_estoque, msg_erro)
end function



' ___________________________________________
' GERA ID ESTOQUE MOVTO
'
function gera_id_estoque_movto(byref id_estoque_movto, byref msg_erro)
	gera_id_estoque_movto=gera_nsu(NSU_ID_ESTOQUE_MOVTO, id_estoque_movto, msg_erro)
end function



' ___________________________________________
' GERA ID ESTOQUE TEMP
'
function gera_id_estoque_temp(byref id_estoque, byref msg_erro)
dim s
	gera_id_estoque_temp = False
	id_estoque = ""
	msg_erro = ""
	if Not gera_nsu(NSU_ID_ESTOQUE_TEMP, s, msg_erro) then exit function
	if len(s) >= TAM_MAX_NSU then s = Right(s, TAM_MAX_NSU-1)
	s = "T" & s
	id_estoque = s
	gera_id_estoque_temp = True
end function



' ___________________________________________
' LE PEDIDO
'
function le_pedido(byval id_pedido, byref r_pedido, byref msg_erro)
dim s
dim rs
dim id_pedido_base
dim blnUsarMemorizacaoCompletaEnderecos

	le_pedido = False
	msg_erro = ""
	id_pedido=Trim("" & id_pedido)
	set r_pedido = New cl_PEDIDO

	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	s = "SELECT" & _
			" *" & _
			"," & montaSubqueryGetUsuarioContexto("InstaladorInstalaUsuarioUltAtualiz", "DecodNome_InstaladorInstalaUsuarioUltAtualiz") & _
			"," & montaSubqueryGetUsuarioContexto("GarantiaIndicadorUsuarioUltAtualiz", "DecodNome_GarantiaIndicadorUsuarioUltAtualiz") & _
			"," & montaSubqueryGetUsuarioContexto("etg_imediata_usuario", "DecodNome_etg_imediata_usuario") & _
			"," & montaSubqueryGetUsuarioContexto("PrevisaoEntregaUsuarioUltAtualiz", "DecodNome_PrevisaoEntregaUsuarioUltAtualiz") & _
		" FROM t_PEDIDO" & _
		" WHERE" & _
			" (pedido='" & id_pedido & "')"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.EOF then
		msg_erro="Pedido nº " & id_pedido & " não está cadastrado."
	else
		with r_pedido
			.pedido						= Trim("" & rs("pedido"))
			.loja						= Trim("" & rs("loja"))
			.data						= rs("data")
			.hora						= Trim("" & rs("hora"))
			.id_cliente					= Trim("" & rs("id_cliente"))
			.midia						= Trim("" & rs("midia"))
			.servicos					= Trim("" & rs("servicos"))
			.vl_servicos				= rs("vl_servicos")
			.vendedor					= Trim("" & rs("vendedor"))
			.st_entrega					= Trim("" & rs("st_entrega"))
			.entregue_data				= rs("entregue_data")
			.entregue_usuario			= Trim("" & rs("entregue_usuario"))
			.cancelado_data				= rs("cancelado_data")
			.cancelado_usuario			= Trim("" & rs("cancelado_usuario"))
			.st_pagto					= Trim("" & rs("st_pagto"))
			.dt_st_pagto				= rs("dt_st_pagto")
			.dt_hr_st_pagto				= rs("dt_hr_st_pagto")
			.usuario_st_pagto			= Trim("" & rs("usuario_st_pagto"))
			.st_recebido				= Trim("" & rs("st_recebido"))
			.obs_1						= Trim("" & rs("obs_1"))
			.obs_2						= Trim("" & rs("obs_2"))
			.obs_3						= Trim("" & rs("obs_3"))
			.obs_4						= Trim("" & rs("obs_4"))
			.qtde_parcelas				= rs("qtde_parcelas")
			.forma_pagto				= Trim("" & rs("forma_pagto"))
			.vl_total_familia			= rs("vl_total_familia")
			.vl_pago_familia			= rs("vl_pago_familia")
			.split_status				= rs("split_status")
			.split_data					= rs("split_data")
			.split_hora					= Trim("" & rs("split_hora"))
			.split_usuario				= Trim("" & rs("split_usuario"))
			.a_entregar_status			= rs("a_entregar_status")
			.a_entregar_data_marcada	= rs("a_entregar_data_marcada")
			.a_entregar_data			= rs("a_entregar_data")
			.a_entregar_hora			= Trim("" & rs("a_entregar_hora"))
			.a_entregar_usuario			= Trim("" & rs("a_entregar_usuario"))
			.loja_indicou				= Trim("" & rs("loja_indicou"))
			.comissao_loja_indicou		= rs("comissao_loja_indicou")
			.vl_frete					= rs("vl_frete")
			.transportadora_id			= Trim("" & rs("transportadora_id"))
			.transportadora_data		= rs("transportadora_data")
			.transportadora_usuario		= Trim("" & rs("transportadora_usuario"))
			.analise_credito			= rs("analise_credito")
			.analise_credito_data		= rs("analise_credito_data")
			.analise_credito_usuario	= Trim("" & rs("analise_credito_usuario"))
			.analise_credito_pendente_vendas_motivo = rs("analise_credito_pendente_vendas_motivo")
			.tipo_parcelamento			= rs("tipo_parcelamento")
			.av_forma_pagto				= rs("av_forma_pagto")
			.pu_forma_pagto 			= rs("pu_forma_pagto")
			.pu_valor					= rs("pu_valor")
			.pu_vencto_apos				= rs("pu_vencto_apos")
			.pc_qtde_parcelas			= rs("pc_qtde_parcelas")
			.pc_valor_parcela			= rs("pc_valor_parcela")
			.pc_maquineta_qtde_parcelas = rs("pc_maquineta_qtde_parcelas")
			.pc_maquineta_valor_parcela = rs("pc_maquineta_valor_parcela")
			.pce_forma_pagto_entrada	= rs("pce_forma_pagto_entrada")
			.pce_forma_pagto_prestacao	= rs("pce_forma_pagto_prestacao")
			.pce_entrada_valor			= rs("pce_entrada_valor")
			.pce_prestacao_qtde			= rs("pce_prestacao_qtde")
			.pce_prestacao_valor		= rs("pce_prestacao_valor")
			.pce_prestacao_periodo		= rs("pce_prestacao_periodo")
			.pse_forma_pagto_prim_prest	= rs("pse_forma_pagto_prim_prest")
			.pse_forma_pagto_demais_prest = rs("pse_forma_pagto_demais_prest")
			.pse_prim_prest_valor		= rs("pse_prim_prest_valor")
			.pse_prim_prest_apos		= rs("pse_prim_prest_apos")
			.pse_demais_prest_qtde		= rs("pse_demais_prest_qtde")
			.pse_demais_prest_valor		= rs("pse_demais_prest_valor")
			.pse_demais_prest_periodo	= rs("pse_demais_prest_periodo")
			.custoFinancFornecTipoParcelamento	= Trim("" & rs("custoFinancFornecTipoParcelamento"))
			.custoFinancFornecQtdeParcelas		= rs("custoFinancFornecQtdeParcelas")
			.indicador					= Trim("" & rs("indicador"))
			.vl_total_NF 				= rs("vl_total_NF")
			.vl_total_RA				= rs("vl_total_RA")
			.perc_RT					= rs("perc_RT")
			.st_orc_virou_pedido		= rs("st_orc_virou_pedido")
			.orcamento					= Trim("" & rs("orcamento"))
			.orcamentista				= Trim("" & rs("orcamentista"))
			.comissao_paga				= rs("comissao_paga")
			.comissao_paga_ult_op		= Trim("" & rs("comissao_paga_ult_op"))
			.comissao_paga_data			= rs("comissao_paga_data")
			.comissao_paga_usuario		= Trim("" & rs("comissao_paga_usuario"))
			.perc_desagio_RA			= rs("perc_desagio_RA")
			.perc_limite_RA_sem_desagio	= rs("perc_limite_RA_sem_desagio")
			.vl_total_RA_liquido		= rs("vl_total_RA_liquido")
			.st_tem_desagio_RA			= rs("st_tem_desagio_RA")
			.qtde_parcelas_desagio_RA	= rs("qtde_parcelas_desagio_RA")
			.transportadora_num_coleta	= Trim("" & rs("transportadora_num_coleta"))
			.transportadora_contato		= Trim("" & rs("transportadora_contato"))
			.pedido_bs_x_at				= Trim("" & rs("pedido_bs_x_at"))
            .pedido_ac                  = Trim("" & rs("pedido_bs_x_ac"))
            .pedido_bs_x_marketplace    = Trim("" & rs("pedido_bs_x_marketplace"))
            .marketplace_codigo_origem  = Trim("" & rs("marketplace_codigo_origem"))
            .pedido_ac_reverso          = Trim("" & rs("pedido_bs_x_ac_reverso"))

			.st_memorizacao_completa_enderecos = 0
			.endereco_logradouro = ""
			.endereco_numero = ""
			.endereco_complemento = ""
			.endereco_bairro = ""
			.endereco_cidade = ""
			.endereco_uf = ""
			.endereco_cep = ""
			.endereco_email = ""
			.endereco_email_xml = ""
			.endereco_nome = ""
			.endereco_nome_iniciais_em_maiusculas = ""
			.endereco_ddd_res = ""
			.endereco_tel_res = ""
			.endereco_ddd_com = ""
			.endereco_tel_com = ""
			.endereco_ramal_com = ""
			.endereco_ddd_cel = ""
			.endereco_tel_cel = ""
			.endereco_ddd_com_2 = ""
			.endereco_tel_com_2 = ""
			.endereco_ramal_com_2 = ""
			.endereco_tipo_pessoa = ""
			.endereco_cnpj_cpf = ""
			.endereco_contribuinte_icms_status = 0
			.endereco_produtor_rural_status = 0
			.endereco_ie = ""
			.endereco_rg = ""
			.endereco_contato = ""

			.endereco_memorizado_status		= rs("endereco_memorizado_status")
			if CLng(.endereco_memorizado_status) <> 0 then
				.endereco_logradouro			= Trim("" & rs("endereco_logradouro"))
				.endereco_numero				= Trim("" & rs("endereco_numero"))
				.endereco_complemento			= Trim("" & rs("endereco_complemento"))
				.endereco_bairro				= Trim("" & rs("endereco_bairro"))
				.endereco_cidade				= Trim("" & rs("endereco_cidade"))
				.endereco_uf					= Trim("" & rs("endereco_uf"))
				.endereco_cep					= Trim("" & rs("endereco_cep"))
				if blnUsarMemorizacaoCompletaEnderecos then
					.st_memorizacao_completa_enderecos = rs("st_memorizacao_completa_enderecos")
					.endereco_email = Trim("" & rs("endereco_email"))
					.endereco_email_xml = Trim("" & rs("endereco_email_xml"))
					.endereco_nome = Trim("" & rs("endereco_nome"))
					.endereco_nome_iniciais_em_maiusculas = Trim("" & rs("endereco_nome_iniciais_em_maiusculas"))
					.endereco_ddd_res = Trim("" & rs("endereco_ddd_res"))
					.endereco_tel_res = Trim("" & rs("endereco_tel_res"))
					.endereco_ddd_com = Trim("" & rs("endereco_ddd_com"))
					.endereco_tel_com = Trim("" & rs("endereco_tel_com"))
					.endereco_ramal_com = Trim("" & rs("endereco_ramal_com"))
					.endereco_ddd_cel = Trim("" & rs("endereco_ddd_cel"))
					.endereco_tel_cel = Trim("" & rs("endereco_tel_cel"))
					.endereco_ddd_com_2 = Trim("" & rs("endereco_ddd_com_2"))
					.endereco_tel_com_2 = Trim("" & rs("endereco_tel_com_2"))
					.endereco_ramal_com_2 = Trim("" & rs("endereco_ramal_com_2"))
					.endereco_tipo_pessoa = Trim("" & rs("endereco_tipo_pessoa"))
					.endereco_cnpj_cpf = Trim("" & rs("endereco_cnpj_cpf"))
					.endereco_contribuinte_icms_status = rs("endereco_contribuinte_icms_status")
					.endereco_produtor_rural_status = rs("endereco_produtor_rural_status")
					.endereco_ie = Trim("" & rs("endereco_ie"))
					.endereco_rg = Trim("" & rs("endereco_rg"))
					.endereco_contato = Trim("" & rs("endereco_contato"))
					end if
				end if

			.EndEtg_endereco = ""
			.EndEtg_endereco_numero = ""
			.EndEtg_endereco_complemento = ""
			.EndEtg_bairro = ""
			.EndEtg_cidade = ""
			.EndEtg_uf = ""
			.EndEtg_cep = ""
			.EndEtg_email = ""
			.EndEtg_email_xml = ""
			.EndEtg_nome = ""
			.EndEtg_nome_iniciais_em_maiusculas = ""
			.EndEtg_ddd_res = ""
			.EndEtg_tel_res = ""
			.EndEtg_ddd_com = ""
			.EndEtg_tel_com = ""
			.EndEtg_ramal_com = ""
			.EndEtg_ddd_cel = ""
			.EndEtg_tel_cel = ""
			.EndEtg_ddd_com_2 = ""
			.EndEtg_tel_com_2 = ""
			.EndEtg_ramal_com_2 = ""
			.EndEtg_tipo_pessoa = ""
			.EndEtg_cnpj_cpf = ""
			.EndEtg_contribuinte_icms_status = 0
			.EndEtg_produtor_rural_status = 0
			.EndEtg_ie = ""
			.EndEtg_rg = ""
			.st_end_entrega				= rs("st_end_entrega")
			if CLng(.st_end_entrega) <> 0 then
				.EndEtg_endereco			= Trim("" & rs("EndEtg_endereco"))
				.EndEtg_endereco_numero		= Trim("" & rs("EndEtg_endereco_numero"))
				.EndEtg_endereco_complemento = Trim("" & rs("EndEtg_endereco_complemento"))
				.EndEtg_bairro				= Trim("" & rs("EndEtg_bairro"))
				.EndEtg_cidade				= Trim("" & rs("EndEtg_cidade"))
				.EndEtg_uf					= Trim("" & rs("EndEtg_uf"))
				.EndEtg_cep					= Trim("" & rs("EndEtg_cep"))
				if blnUsarMemorizacaoCompletaEnderecos then
					.EndEtg_email = Trim("" & rs("EndEtg_email"))
					.EndEtg_email_xml = Trim("" & rs("EndEtg_email_xml"))
					.EndEtg_nome = Trim("" & rs("EndEtg_nome"))
					.EndEtg_nome_iniciais_em_maiusculas = Trim("" & rs("EndEtg_nome_iniciais_em_maiusculas"))
					.EndEtg_ddd_res = Trim("" & rs("EndEtg_ddd_res"))
					.EndEtg_tel_res = Trim("" & rs("EndEtg_tel_res"))
					.EndEtg_ddd_com = Trim("" & rs("EndEtg_ddd_com"))
					.EndEtg_tel_com = Trim("" & rs("EndEtg_tel_com"))
					.EndEtg_ramal_com = Trim("" & rs("EndEtg_ramal_com"))
					.EndEtg_ddd_cel = Trim("" & rs("EndEtg_ddd_cel"))
					.EndEtg_tel_cel = Trim("" & rs("EndEtg_tel_cel"))
					.EndEtg_ddd_com_2 = Trim("" & rs("EndEtg_ddd_com_2"))
					.EndEtg_tel_com_2 = Trim("" & rs("EndEtg_tel_com_2"))
					.EndEtg_ramal_com_2 = Trim("" & rs("EndEtg_ramal_com_2"))
					.EndEtg_tipo_pessoa = Trim("" & rs("EndEtg_tipo_pessoa"))
					.EndEtg_cnpj_cpf = Trim("" & rs("EndEtg_cnpj_cpf"))
					.EndEtg_contribuinte_icms_status = rs("EndEtg_contribuinte_icms_status")
					.EndEtg_produtor_rural_status = rs("EndEtg_produtor_rural_status")
					.EndEtg_ie = Trim("" & rs("EndEtg_ie"))
					.EndEtg_rg = Trim("" & rs("EndEtg_rg"))
					end if
				end if

			.analise_endereco_tratar_status = rs("analise_endereco_tratar_status")
			.analise_endereco_tratado_status = rs("analise_endereco_tratado_status")
			.analise_endereco_tratado_data	= rs("analise_endereco_tratado_data")
			.analise_endereco_tratado_data_hora = rs("analise_endereco_tratado_data_hora")
			.analise_endereco_tratado_usuario = Trim("" & rs("analise_endereco_tratado_usuario"))

			.st_etg_imediata			= rs("st_etg_imediata")
			.etg_imediata_data			= rs("etg_imediata_data")
			if Trim("" & rs("DecodNome_etg_imediata_usuario")) = "" then
				.etg_imediata_usuario		= Trim("" & rs("etg_imediata_usuario"))
			else
				.etg_imediata_usuario = Left(Trim("" & rs("DecodNome_etg_imediata_usuario")), MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR)
				if Left(Trim("" & rs("etg_imediata_usuario")), 3) = "[" & Cstr(COD_USUARIO_CONTEXTO__VENDEDOR_DO_PARCEIRO) & "]" then
					.etg_imediata_usuario = "[VP] " & .etg_imediata_usuario
					end if
				end if
			.etg_imediata_usuario_RawData = Trim("" & rs("etg_imediata_usuario"))
			.PrevisaoEntregaData = rs("PrevisaoEntregaData")
			if Trim("" & rs("DecodNome_PrevisaoEntregaUsuarioUltAtualiz")) = "" then
				.PrevisaoEntregaUsuarioUltAtualiz = Trim("" & rs("PrevisaoEntregaUsuarioUltAtualiz"))
			else
				.PrevisaoEntregaUsuarioUltAtualiz = Left(Trim("" & rs("DecodNome_PrevisaoEntregaUsuarioUltAtualiz")), MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR)
				if Left(Trim("" & rs("PrevisaoEntregaUsuarioUltAtualiz")), 3) = "[" & Cstr(COD_USUARIO_CONTEXTO__VENDEDOR_DO_PARCEIRO) & "]" then
					.PrevisaoEntregaUsuarioUltAtualiz = "[VP] " & .PrevisaoEntregaUsuarioUltAtualiz
					end if
				end if
			.PrevisaoEntregaUsuarioUltAtualiz_RawData = Trim("" & rs("PrevisaoEntregaUsuarioUltAtualiz"))
			.PrevisaoEntregaDtHrUltAtualiz = rs("PrevisaoEntregaDtHrUltAtualiz")
			.frete_status				= rs("frete_status")
			.frete_valor				= rs("frete_valor")
			.frete_data					= rs("frete_data")
			.frete_usuario				= Trim("" & rs("frete_usuario"))
			.StBemUsoConsumo			= rs("StBemUsoConsumo")
			.PedidoRecebidoStatus		= rs("PedidoRecebidoStatus")
			.PedidoRecebidoData			= rs("PedidoRecebidoData")
			.PedidoRecebidoUsuarioUltAtualiz = Trim("" & rs("PedidoRecebidoUsuarioUltAtualiz"))
			.PedidoRecebidoDtHrUltAtualiz = rs("PedidoRecebidoDtHrUltAtualiz")
			.InstaladorInstalaStatus	= rs("InstaladorInstalaStatus")
			if Trim("" & rs("DecodNome_InstaladorInstalaUsuarioUltAtualiz")) = "" then
				.InstaladorInstalaUsuarioUltAtualiz = Trim("" & rs("InstaladorInstalaUsuarioUltAtualiz"))
			else
				.InstaladorInstalaUsuarioUltAtualiz = Left(Trim("" & rs("DecodNome_InstaladorInstalaUsuarioUltAtualiz")), MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR)
				if Left(Trim("" & rs("InstaladorInstalaUsuarioUltAtualiz")), 3) = "[" & Cstr(COD_USUARIO_CONTEXTO__VENDEDOR_DO_PARCEIRO) & "]" then
					.InstaladorInstalaUsuarioUltAtualiz = "[VP] " & .InstaladorInstalaUsuarioUltAtualiz
					end if
				end if
			.InstaladorInstalaUsuarioUltAtualiz_RawData = Trim("" & rs("InstaladorInstalaUsuarioUltAtualiz"))
			.InstaladorInstalaDtHrUltAtualiz = rs("InstaladorInstalaDtHrUltAtualiz")
			.GarantiaIndicadorStatus	= rs("GarantiaIndicadorStatus")
			if Trim("" & rs("DecodNome_GarantiaIndicadorUsuarioUltAtualiz")) = "" then
				.GarantiaIndicadorUsuarioUltAtualiz = rs("GarantiaIndicadorUsuarioUltAtualiz")
			else
				.GarantiaIndicadorUsuarioUltAtualiz = Left(Trim("" & rs("DecodNome_GarantiaIndicadorUsuarioUltAtualiz")), MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR)
				if Left(Trim("" & rs("GarantiaIndicadorUsuarioUltAtualiz")), 3) = "[" & Cstr(COD_USUARIO_CONTEXTO__VENDEDOR_DO_PARCEIRO) & "]" then
					.GarantiaIndicadorUsuarioUltAtualiz = "[VP] " & .GarantiaIndicadorUsuarioUltAtualiz
					end if
				end if
			.GarantiaIndicadorUsuarioUltAtualiz_RawData = rs("GarantiaIndicadorUsuarioUltAtualiz")
			.GarantiaIndicadorDtHrUltAtualiz = rs("GarantiaIndicadorDtHrUltAtualiz")
			.perc_desagio_RA_liquida	= rs("perc_desagio_RA_liquida")
			.permite_RA_status			= rs("permite_RA_status")
			.st_violado_permite_RA_status		= rs("st_violado_permite_RA_status")
			.dt_hr_violado_permite_RA_status	= rs("dt_hr_violado_permite_RA_status")
			.usuario_violado_permite_RA_status	= Trim("" & rs("usuario_violado_permite_RA_status"))
			.opcao_possui_RA			= Trim("" & rs("opcao_possui_RA"))
			.transportadora_selecao_auto_status	= rs("transportadora_selecao_auto_status")
            .cancelado_data_hora        = rs("cancelado_data_hora")
            .cancelado_motivo           = Trim("" & rs("cancelado_motivo"))
            .cancelado_codigo_motivo    = Trim("" & rs("cancelado_codigo_motivo"))
            .cancelado_codigo_sub_motivo = Trim("" & rs("cancelado_codigo_sub_motivo"))
            .EndEtg_obs = Trim("" & rs("EndEtg_obs"))
            .EndEtg_cod_justificativa = Trim("" & rs("EndEtg_cod_justificativa"))
			.id_nfe_emitente = rs("id_nfe_emitente")
            .NFe_texto_constar = Trim("" & rs("NFe_texto_constar"))
            .NFe_xPed = Trim("" & rs("NFe_xPed"))
			.st_auto_split = rs("st_auto_split")
            .st_forma_pagto_possui_parcela_cartao = rs("st_forma_pagto_possui_parcela_cartao")
			.st_forma_pagto_possui_parcela_cartao_maquineta = rs("st_forma_pagto_possui_parcela_cartao_maquineta")
			.usuario_cadastro = Trim("" & rs("usuario_cadastro"))
			.plataforma_origem_pedido = rs("plataforma_origem_pedido")
			.PagtoAntecipadoStatus = rs("PagtoAntecipadoStatus")
			.PagtoAntecipadoDataHora = rs("PagtoAntecipadoDataHora")
			.PagtoAntecipadoUsuario = Trim("" & rs("PagtoAntecipadoUsuario"))
			.PagtoAntecipadoQuitadoStatus = rs("PagtoAntecipadoQuitadoStatus")
			.PagtoAntecipadoQuitadoDataHora = rs("PagtoAntecipadoQuitadoDataHora")
			.PagtoAntecipadoQuitadoUsuario = Trim("" & rs("PagtoAntecipadoQuitadoUsuario"))
			.PrevisaoEntregaTranspData = rs("PrevisaoEntregaTranspData")
			.PrevisaoEntregaTranspUsuarioUltAtualiz = Trim("" & rs("PrevisaoEntregaTranspUsuarioUltAtualiz"))
			.PrevisaoEntregaTranspDtHrUltAtualiz = rs("PrevisaoEntregaTranspDtHrUltAtualiz")
			.IdOrcamentoCotacao = rs("IdOrcamentoCotacao")
			.IdIndicadorVendedor = rs("IdIndicadorVendedor")
			.vl_frete_total_cobrado_cliente = rs("vl_frete_total_cobrado_cliente")
			.vl_base_calculo_frete_total_cobrado_cliente = rs("vl_base_calculo_frete_total_cobrado_cliente")
			.perc_max_comissao_padrao = rs("perc_max_comissao_padrao")
			.perc_max_comissao_e_desconto_padrao = rs("perc_max_comissao_e_desconto_padrao")
			.InstaladorInstalaIdTipoUsuarioContexto = rs("InstaladorInstalaIdTipoUsuarioContexto")
			.InstaladorInstalaIdUsuarioUltAtualiz = rs("InstaladorInstalaIdUsuarioUltAtualiz")
			.GarantiaIndicadorIdTipoUsuarioContexto = rs("GarantiaIndicadorIdTipoUsuarioContexto")
			.GarantiaIndicadorIdUsuarioUltAtualiz = rs("GarantiaIndicadorIdUsuarioUltAtualiz")
			.EtgImediataIdTipoUsuarioContexto = rs("EtgImediataIdTipoUsuarioContexto")
			.EtgImediataIdUsuarioUltAtualiz = rs("EtgImediataIdUsuarioUltAtualiz")
			.PrevisaoEntregaIdTipoUsuarioContexto = rs("PrevisaoEntregaIdTipoUsuarioContexto")
			.PrevisaoEntregaIdUsuarioUltAtualiz = rs("PrevisaoEntregaIdUsuarioUltAtualiz")
			.UsuarioCadastroIdTipoUsuarioContexto = rs("UsuarioCadastroIdTipoUsuarioContexto")
			.UsuarioCadastroId = rs("UsuarioCadastroId")
			end with
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

'	SE O Nº DO PEDIDO É DE PEDIDO-FILHOTE, OBTÉM OS DADOS QUE FICAM SEMPRE NO PEDIDO-BASE!!!
	if IsPedidoFilhote(id_pedido) then
		id_pedido_base = retorna_num_pedido_base(id_pedido)
		s="SELECT * FROM t_PEDIDO WHERE (pedido='" & id_pedido_base & "')"
		if rs.State <> 0 then rs.Close
		set rs=cn.Execute(s)
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
			
		if rs.EOF then
			msg_erro="Pedido base nº " & id_pedido_base & " não está cadastrado."
		else
			with r_pedido
				.st_pagto					= Trim("" & rs("st_pagto"))
				.dt_st_pagto				= rs("dt_st_pagto")
				.dt_hr_st_pagto				= rs("dt_hr_st_pagto")
				.usuario_st_pagto			= Trim("" & rs("usuario_st_pagto"))
				.st_recebido				= Trim("" & rs("st_recebido"))
				.vl_total_familia			= rs("vl_total_familia")
				.vl_pago_familia			= rs("vl_pago_familia")
				.analise_credito			= rs("analise_credito")
				.analise_credito_data		= rs("analise_credito_data")
				.analise_credito_usuario	= rs("analise_credito_usuario")
				.analise_credito_pendente_vendas_motivo = rs("analise_credito_pendente_vendas_motivo")
				.tipo_parcelamento			= rs("tipo_parcelamento")
				.av_forma_pagto				= rs("av_forma_pagto")
				.pu_forma_pagto 			= rs("pu_forma_pagto")
				.pu_valor					= rs("pu_valor")
				.pu_vencto_apos				= rs("pu_vencto_apos")
				.pc_qtde_parcelas			= rs("pc_qtde_parcelas")
				.pc_valor_parcela			= rs("pc_valor_parcela")
				.pc_maquineta_qtde_parcelas = rs("pc_maquineta_qtde_parcelas")
				.pc_maquineta_valor_parcela = rs("pc_maquineta_valor_parcela")
				.pce_forma_pagto_entrada	= rs("pce_forma_pagto_entrada")
				.pce_forma_pagto_prestacao	= rs("pce_forma_pagto_prestacao")
				.pce_entrada_valor			= rs("pce_entrada_valor")
				.pce_prestacao_qtde			= rs("pce_prestacao_qtde")
				.pce_prestacao_valor		= rs("pce_prestacao_valor")
				.pce_prestacao_periodo		= rs("pce_prestacao_periodo")
				.pse_forma_pagto_prim_prest	= rs("pse_forma_pagto_prim_prest")
				.pse_forma_pagto_demais_prest = rs("pse_forma_pagto_demais_prest")
				.pse_prim_prest_valor		= rs("pse_prim_prest_valor")
				.pse_prim_prest_apos		= rs("pse_prim_prest_apos")
				.pse_demais_prest_qtde		= rs("pse_demais_prest_qtde")
				.pse_demais_prest_valor		= rs("pse_demais_prest_valor")
				.pse_demais_prest_periodo	= rs("pse_demais_prest_periodo")
	            .st_forma_pagto_possui_parcela_cartao = rs("st_forma_pagto_possui_parcela_cartao")
	            .st_forma_pagto_possui_parcela_cartao_maquineta = rs("st_forma_pagto_possui_parcela_cartao_maquineta")
				.custoFinancFornecTipoParcelamento	= Trim("" & rs("custoFinancFornecTipoParcelamento"))
				.custoFinancFornecQtdeParcelas 		= rs("custoFinancFornecQtdeParcelas")
				.indicador					= Trim("" & rs("indicador"))
				.vl_total_NF 				= rs("vl_total_NF")
				.vl_total_RA				= rs("vl_total_RA")
				.perc_RT					= rs("perc_RT")
				.st_orc_virou_pedido		= rs("st_orc_virou_pedido")
				.orcamento					= Trim("" & rs("orcamento"))
				.orcamentista				= Trim("" & rs("orcamentista"))
				.comissao_paga				= rs("comissao_paga")
				.comissao_paga_ult_op		= Trim("" & rs("comissao_paga_ult_op"))
				.comissao_paga_data			= rs("comissao_paga_data")
				.comissao_paga_usuario		= Trim("" & rs("comissao_paga_usuario"))
				.perc_desagio_RA			= rs("perc_desagio_RA")
				.perc_limite_RA_sem_desagio	= rs("perc_limite_RA_sem_desagio")
				.vl_total_RA_liquido		= rs("vl_total_RA_liquido")
				.st_tem_desagio_RA			= rs("st_tem_desagio_RA")
				.qtde_parcelas_desagio_RA	= rs("qtde_parcelas_desagio_RA")
				.perc_desagio_RA_liquida	= rs("perc_desagio_RA_liquida")
				
				'A implementação da memorização completa dos endereços assume que cada pedido, inclusive os pedidos-filhote,
				'armazena consigo os dados de endereço.
				if Not blnUsarMemorizacaoCompletaEnderecos then
					.endereco_memorizado_status		= rs("endereco_memorizado_status")
					.endereco_logradouro			= Trim("" & rs("endereco_logradouro"))
					.endereco_numero				= Trim("" & rs("endereco_numero"))
					.endereco_complemento			= Trim("" & rs("endereco_complemento"))
					.endereco_bairro				= Trim("" & rs("endereco_bairro"))
					.endereco_cidade				= Trim("" & rs("endereco_cidade"))
					.endereco_uf					= Trim("" & rs("endereco_uf"))
					.endereco_cep					= Trim("" & rs("endereco_cep"))
					end if

				.analise_endereco_tratar_status = rs("analise_endereco_tratar_status")
				.analise_endereco_tratado_status = rs("analise_endereco_tratado_status")
				.analise_endereco_tratado_data	= rs("analise_endereco_tratado_data")
				.analise_endereco_tratado_data_hora = rs("analise_endereco_tratado_data_hora")
				.analise_endereco_tratado_usuario = Trim("" & rs("analise_endereco_tratado_usuario"))
				.pedido_bs_x_at				= Trim("" & rs("pedido_bs_x_at"))
				.pedido_ac                  = Trim("" & rs("pedido_bs_x_ac"))
				.pedido_bs_x_marketplace    = Trim("" & rs("pedido_bs_x_marketplace"))
				.marketplace_codigo_origem  = Trim("" & rs("marketplace_codigo_origem"))
				.pedido_ac_reverso          = Trim("" & rs("pedido_bs_x_ac_reverso"))
				.usuario_cadastro = Trim("" & rs("usuario_cadastro"))
				.plataforma_origem_pedido = rs("plataforma_origem_pedido")
				'O status que indica se o pedido será pago através de pagamento antecipado é único para toda a família de pedidos
				'e está atrelado ao processo de análise de crédito, que também é único por família de pedidos. Por esse motivo,
				'a informação deve ser sempre armazenada e lida do pedido-base, mesmo que este esteja cancelado e o(s) pedido(s)-filhote não.
				'Já a informação que indica se o pagamento antecipado foi quitado é armazenado em cada pedido individualmente, pois pode
				'ocorrer de serem gerados conjuntos de boletos separados para o pedido-pai e o(s) pedido(s)-filhote.
				.PagtoAntecipadoStatus = rs("PagtoAntecipadoStatus")
				.PagtoAntecipadoDataHora = rs("PagtoAntecipadoDataHora")
				.PagtoAntecipadoUsuario = Trim("" & rs("PagtoAntecipadoUsuario"))
				.IdOrcamentoCotacao = rs("IdOrcamentoCotacao")
				.IdIndicadorVendedor = rs("IdIndicadorVendedor")
				.vl_frete_total_cobrado_cliente = rs("vl_frete_total_cobrado_cliente")
				.vl_base_calculo_frete_total_cobrado_cliente = rs("vl_base_calculo_frete_total_cobrado_cliente")
				.perc_max_comissao_padrao = rs("perc_max_comissao_padrao")
				.perc_max_comissao_e_desconto_padrao = rs("perc_max_comissao_e_desconto_padrao")
				.InstaladorInstalaIdTipoUsuarioContexto = rs("InstaladorInstalaIdTipoUsuarioContexto")
				.InstaladorInstalaIdUsuarioUltAtualiz = rs("InstaladorInstalaIdUsuarioUltAtualiz")
				.GarantiaIndicadorIdTipoUsuarioContexto = rs("GarantiaIndicadorIdTipoUsuarioContexto")
				.GarantiaIndicadorIdUsuarioUltAtualiz = rs("GarantiaIndicadorIdUsuarioUltAtualiz")
				.EtgImediataIdTipoUsuarioContexto = rs("EtgImediataIdTipoUsuarioContexto")
				.EtgImediataIdUsuarioUltAtualiz = rs("EtgImediataIdUsuarioUltAtualiz")
				.PrevisaoEntregaIdTipoUsuarioContexto = rs("PrevisaoEntregaIdTipoUsuarioContexto")
				.PrevisaoEntregaIdUsuarioUltAtualiz = rs("PrevisaoEntregaIdUsuarioUltAtualiz")
				.UsuarioCadastroIdTipoUsuarioContexto = rs("UsuarioCadastroIdTipoUsuarioContexto")
				.UsuarioCadastroId = rs("UsuarioCadastroId")
				end with
			end if
		end if
		
	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_pedido=True
end function



' ___________________________________________
' LE PEDIDO ITEM
'
function le_pedido_item(byval id_pedido, byref v_pedido_item, byref msg_erro)
dim s
dim rs
	le_pedido_item = False
	msg_erro = ""
	id_pedido=Trim("" & id_pedido)
	redim v_pedido_item(0)
	set v_pedido_item(0) = New cl_ITEM_PEDIDO
	
	s="SELECT * FROM t_PEDIDO_ITEM WHERE (pedido='" & id_pedido & "') ORDER BY sequencia"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Não há itens cadastrados para o pedido nº " & id_pedido & "."
	else
		do while Not rs.EOF 
			if Trim(v_pedido_item(Ubound(v_pedido_item)).produto)<>"" then
				redim preserve v_pedido_item(Ubound(v_pedido_item)+1)
				set v_pedido_item(ubound(v_pedido_item)) = New cl_ITEM_PEDIDO
				end if
			with v_pedido_item(Ubound(v_pedido_item))
				.pedido					= Trim("" & rs("pedido"))
				.fabricante				= Trim("" & rs("fabricante"))
				.produto				= Trim("" & rs("produto"))
				.qtde					= rs("qtde")
				.desc_dado				= rs("desc_dado")
				.preco_venda			= rs("preco_venda")
				.preco_NF				= rs("preco_NF")
				.preco_fabricante		= rs("preco_fabricante")
				.vl_custo2				= rs("vl_custo2")
				.preco_lista			= rs("preco_lista")
				.margem					= rs("margem")
				.desc_max				= rs("desc_max")
				.comissao				= rs("comissao")
				.descricao				= Trim("" & rs("descricao"))
				.descricao_html			= Trim("" & rs("descricao_html"))
				.ean					= Trim("" & rs("ean"))
				.grupo					= Trim("" & rs("grupo"))
                .subgrupo				= Trim("" & rs("subgrupo"))
				.peso					= rs("peso")
				.qtde_volumes			= rs("qtde_volumes")
				.abaixo_min_status		= rs("abaixo_min_status")
				.abaixo_min_autorizacao	= Trim("" & rs("abaixo_min_autorizacao"))
				.abaixo_min_autorizador	= Trim("" & rs("abaixo_min_autorizador"))
				.abaixo_min_superv_autorizador	= Trim("" & rs("abaixo_min_superv_autorizador"))
				.sequencia				= rs("sequencia")
				.markup_fabricante		= rs("markup_fabricante")
				.custoFinancFornecCoeficiente = rs("custoFinancFornecCoeficiente")
				.custoFinancFornecPrecoListaBase = rs("custoFinancFornecPrecoListaBase")
				.cubagem				= rs("cubagem")
				.ncm					= Trim("" & rs("ncm"))
				.cst					= Trim("" & rs("cst"))
				.descontinuado			= Trim("" & rs("descontinuado"))
				.cod_produto_xml_fabricante = Trim("" & rs("cod_produto_xml_fabricante"))
				.cod_produto_alfanum_fabricante = Trim("" & rs("cod_produto_alfanum_fabricante"))
				.potencia_valor = rs("potencia_valor")
				.id_unidade_potencia = rs("id_unidade_potencia")
				'O campo 'StatusDescontoSuperior' é do tipo 'bit' e o ASP converte como se fosse um boolean, portanto, é feito um tratamento manual
				if rs("StatusDescontoSuperior") = 0 then .StatusDescontoSuperior = 0 else .StatusDescontoSuperior = 1
				.IdUsuarioDescontoSuperior = rs("IdUsuarioDescontoSuperior")
				.DataHoraDescontoSuperior = rs("DataHoraDescontoSuperior")
				end with
			rs.MoveNext
			Loop
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_pedido_item=True
end function



' ___________________________________________
' LE PEDIDO ITEM SERVICO
'
function le_pedido_item_servico(byval id_pedido, byref v_pedido_item_servico, byref msg_erro)
dim s
dim rs
	le_pedido_item_servico = False
	msg_erro = ""
	id_pedido=Trim("" & id_pedido)
	redim v_pedido_item_servico(0)
	set v_pedido_item_servico(0) = New cl_ITEM_PEDIDO_SERVICO
	
	s="SELECT * FROM t_PEDIDO_ITEM_SERVICO WHERE (pedido='" & id_pedido & "') ORDER BY sequencia"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Não há itens de serviço cadastrados para o pedido nº " & id_pedido & "."
	else
		do while Not rs.EOF 
			if Trim(v_pedido_item_servico(Ubound(v_pedido_item_servico)).produto)<>"" then
				redim preserve v_pedido_item_servico(Ubound(v_pedido_item_servico)+1)
				set v_pedido_item_servico(ubound(v_pedido_item_servico)) = New cl_ITEM_PEDIDO_SERVICO
				end if
			with v_pedido_item_servico(Ubound(v_pedido_item_servico))
				.pedido					= Trim("" & rs("pedido"))
				.fabricante				= Trim("" & rs("fabricante"))
				.produto				= Trim("" & rs("produto"))
				.qtde					= rs("qtde")
				.desc_dado				= rs("desc_dado")
				.preco_venda			= rs("preco_venda")
				.preco_NF				= rs("preco_NF")
				.preco_fabricante		= rs("preco_fabricante")
				.vl_custo2				= rs("vl_custo2")
				.preco_lista			= rs("preco_lista")
				.margem					= rs("margem")
				.desc_max				= rs("desc_max")
				.comissao				= rs("comissao")
				.descricao				= Trim("" & rs("descricao"))
				.descricao_html			= Trim("" & rs("descricao_html"))
				.ean					= Trim("" & rs("ean"))
				.grupo					= Trim("" & rs("grupo"))
                .subgrupo				= Trim("" & rs("subgrupo"))
				.peso					= rs("peso")
				.qtde_volumes			= rs("qtde_volumes")
				.abaixo_min_status		= rs("abaixo_min_status")
				.abaixo_min_autorizacao	= Trim("" & rs("abaixo_min_autorizacao"))
				.abaixo_min_autorizador	= Trim("" & rs("abaixo_min_autorizador"))
				.abaixo_min_superv_autorizador	= Trim("" & rs("abaixo_min_superv_autorizador"))
				.sequencia				= rs("sequencia")
				.markup_fabricante		= rs("markup_fabricante")
				.custoFinancFornecCoeficiente = rs("custoFinancFornecCoeficiente")
				.custoFinancFornecPrecoListaBase = rs("custoFinancFornecPrecoListaBase")
				.cubagem				= rs("cubagem")
				.ncm					= Trim("" & rs("ncm"))
				.cst					= Trim("" & rs("cst"))
				.descontinuado			= Trim("" & rs("descontinuado"))
				.cod_produto_xml_fabricante = Trim("" & rs("cod_produto_xml_fabricante"))
				.cod_produto_alfanum_fabricante = Trim("" & rs("cod_produto_alfanum_fabricante"))
				.potencia_valor = rs("potencia_valor")
				.id_unidade_potencia = rs("id_unidade_potencia")
				'O campo 'StatusDescontoSuperior' é do tipo 'bit' e o ASP converte como se fosse um boolean, portanto, é feito um tratamento manual
				if rs("StatusDescontoSuperior") = 0 then .StatusDescontoSuperior = 0 else .StatusDescontoSuperior = 1
				.IdUsuarioDescontoSuperior = rs("IdUsuarioDescontoSuperior")
				.DataHoraDescontoSuperior = rs("DataHoraDescontoSuperior")
				end with
			rs.MoveNext
			Loop
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_pedido_item_servico=True
end function



' ___________________________________________
' LE PEDIDO ITEM CONSOLIDADO FAMILIA
'
function le_pedido_item_consolidado_familia(byval id_pedido, byref v_pedido_item, byref msg_erro)
dim id_pedido_base
dim s, i_seq
dim rs, rsi
	le_pedido_item_consolidado_familia = False
	msg_erro = ""
	id_pedido=Trim("" & id_pedido)
	id_pedido_base = retorna_num_pedido_base(id_pedido)
	redim v_pedido_item(0)
	set v_pedido_item(0) = New cl_ITEM_PEDIDO
	
	if Not cria_recordset_otimista(rsi, msg_erro) then exit function

	s = "SELECT" & _
			" fabricante," & _
			" produto," & _
			" SUM(qtde) AS total_qtde," & _
			" SUM(qtde*preco_venda) AS total_preco_venda," & _
			" SUM(qtde*preco_NF) AS total_preco_NF" & _
		" FROM t_PEDIDO_ITEM tPI" & _
			" INNER JOIN t_PEDIDO tP ON (tPI.pedido = tP.pedido)" & _
		" WHERE" & _
			" (tPI.pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')" & _
			" AND (tP.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
		" GROUP BY" & _
			" fabricante," & _
			" produto" & _
		" ORDER BY" & _
			" fabricante," & _
			" produto"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Não há itens cadastrados para o pedido nº " & id_pedido_base & "."
	else
		i_seq = 0
		do while Not rs.EOF 
			if Trim(v_pedido_item(Ubound(v_pedido_item)).produto)<>"" then
				redim preserve v_pedido_item(Ubound(v_pedido_item)+1)
				set v_pedido_item(ubound(v_pedido_item)) = New cl_ITEM_PEDIDO
				end if
			with v_pedido_item(Ubound(v_pedido_item))
				.pedido					= id_pedido_base
				.fabricante				= Trim("" & rs("fabricante"))
				.produto				= Trim("" & rs("produto"))
				.qtde					= rs("total_qtde")
				
			'	UTILIZA O PREÇO MÉDIO DE VENDA E NF, POIS OS VALORES PODEM TER SIDO EDITADOS E ESTAREM DIFERENTES NOS PEDIDOS PAI E FILHOTES
				if rs("total_qtde") = 0 then
					.preco_venda = 0
					.preco_NF = 0
				else
					.preco_venda = rs("total_preco_venda") / rs("total_qtde")
					.preco_NF = rs("total_preco_NF") / rs("total_qtde")
					end if

				i_seq = i_seq + 1

				s = "SELECT TOP 1 " & _
						"*" & _
					" FROM t_PEDIDO_ITEM" & _
					" WHERE" & _
						" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')" & _
						" AND (fabricante = '" & Trim("" & rs("fabricante")) & "')" & _
						" AND (produto = '" & Trim("" & rs("produto")) & "')" & _
					" ORDER BY" & _
						" pedido"
				if rsi.State <> 0 then rsi.Close
				rsi.Open s, cn
				if Not rsi.Eof then
					.desc_dado				= rsi("desc_dado")
					if .preco_venda = 0 then .preco_venda = rsi("preco_venda")
					if .preco_NF = 0 then .preco_NF = rsi("preco_NF")
					.preco_fabricante		= rsi("preco_fabricante")
					.vl_custo2				= rsi("vl_custo2")
					.preco_lista			= rsi("preco_lista")
					.margem					= rsi("margem")
					.desc_max				= rsi("desc_max")
					.comissao				= rsi("comissao")
					.descricao				= Trim("" & rsi("descricao"))
					.descricao_html			= Trim("" & rsi("descricao_html"))
					.ean					= Trim("" & rsi("ean"))
					.grupo					= Trim("" & rsi("grupo"))
                    .subgrupo				= Trim("" & rsi("subgrupo"))
					.peso					= rsi("peso")
					.qtde_volumes			= rsi("qtde_volumes")
					.abaixo_min_status		= rsi("abaixo_min_status")
					.abaixo_min_autorizacao	= Trim("" & rsi("abaixo_min_autorizacao"))
					.abaixo_min_autorizador	= Trim("" & rsi("abaixo_min_autorizador"))
					.abaixo_min_superv_autorizador	= Trim("" & rsi("abaixo_min_superv_autorizador"))
					.sequencia				= i_seq
					.markup_fabricante		= rsi("markup_fabricante")
					.custoFinancFornecCoeficiente = rsi("custoFinancFornecCoeficiente")
					.custoFinancFornecPrecoListaBase = rsi("custoFinancFornecPrecoListaBase")
					.cubagem				= rsi("cubagem")
					.ncm					= Trim("" & rsi("ncm"))
					.cst					= Trim("" & rsi("cst"))
					.descontinuado			= Trim("" & rsi("descontinuado"))
					.cod_produto_xml_fabricante = Trim("" & rsi("cod_produto_xml_fabricante"))
					.cod_produto_alfanum_fabricante = Trim("" & rsi("cod_produto_alfanum_fabricante"))
					.potencia_valor = rsi("potencia_valor")
					.id_unidade_potencia = rsi("id_unidade_potencia")
					end if
				end with
			rs.MoveNext
			Loop
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close
	if rsi.State <> 0 then rsi.Close

	if msg_erro = "" then le_pedido_item_consolidado_familia=True
end function



' ___________________________________________
' LE ORCAMENTISTA E INDICADOR
'
function le_orcamentista_e_indicador(byval apelido, byref r_orcamentista_e_indicador, byref msg_erro)
dim s
dim rs

	le_orcamentista_e_indicador = False
	msg_erro = ""
	apelido=Trim("" & apelido)
	set r_orcamentista_e_indicador = New cl_ORCAMENTISTA_E_INDICADOR
	if apelido = "" then
		r_orcamentista_e_indicador.apelido = ""
		r_orcamentista_e_indicador.Id = 0
		le_orcamentista_e_indicador = True
		exit function
		end if
		
	s="SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido='" & QuotedStr(apelido) & "')"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Orçamentista/Indicador '" & apelido & "' não está cadastrado."
	else
		with r_orcamentista_e_indicador
			.apelido					= Trim("" & rs("apelido"))
			.Id							= rs("Id")
			.id_magento_b2b				= rs("id_magento_b2b")
			.cnpj_cpf					= Trim("" & rs("cnpj_cpf"))
			.tipo						= Trim("" & rs("tipo"))
			.ie_rg						= Trim("" & rs("ie_rg"))
			.razao_social_nome			= Trim("" & rs("razao_social_nome"))
			.endereco					= Trim("" & rs("endereco"))
			.endereco_numero			= Trim("" & rs("endereco_numero"))
			.endereco_complemento		= Trim("" & rs("endereco_complemento"))
			.bairro						= Trim("" & rs("bairro"))
			.cidade						= Trim("" & rs("cidade"))
			.uf							= Trim("" & rs("uf"))
			.cep						= Trim("" & rs("cep"))
			.ddd						= Trim("" & rs("ddd"))
			.telefone					= Trim("" & rs("telefone"))
			.fax						= Trim("" & rs("fax"))
			.ddd_cel					= Trim("" & rs("ddd_cel"))
			.tel_cel					= Trim("" & rs("tel_cel"))
			.contato					= Trim("" & rs("contato"))
			.banco						= Trim("" & rs("banco"))
			.agencia					= Trim("" & rs("agencia"))
			.conta						= Trim("" & rs("conta"))
			.favorecido					= Trim("" & rs("favorecido"))
			.loja						= Trim("" & rs("loja"))
			.vendedor					= Trim("" & rs("vendedor"))
			.hab_acesso_sistema			= rs("hab_acesso_sistema")
			.status						= Trim("" & rs("status"))
			.senha						= Trim("" & rs("senha"))
			.datastamp					= Trim("" & rs("datastamp"))
			.dt_ult_alteracao_senha		= rs("dt_ult_alteracao_senha")
			.dt_cadastro				= rs("dt_cadastro")
			.usuario_cadastro			= Trim("" & rs("usuario_cadastro"))
			.dt_ult_atualizacao			= rs("dt_ult_atualizacao")
			.usuario_ult_atualizacao	= Trim("" & rs("usuario_ult_atualizacao"))
			.dt_ult_acesso				= rs("dt_ult_acesso")
			.desempenho_nota			= Trim("" & rs("desempenho_nota"))
			.desempenho_nota_data		= rs("desempenho_nota_data")
			.desempenho_nota_usuario	= Trim("" & rs("desempenho_nota_usuario"))
			.perc_desagio_RA			= rs("perc_desagio_RA")
			.vl_limite_mensal			= rs("vl_limite_mensal")
			.email						= Trim("" & rs("email"))
			.captador					= Trim("" & rs("captador"))
			.checado_status				= rs("checado_status")
			.checado_data				= rs("checado_data")
			.checado_usuario			= Trim("" & rs("checado_usuario"))
			.obs						= Trim("" & rs("obs"))
			.vl_meta					= rs("vl_meta")
			.UsuarioUltAtualizVlMeta	= Trim("" & rs("UsuarioUltAtualizVlMeta"))
			.DtHrUltAtualizVlMeta		= rs("DtHrUltAtualizVlMeta")
			.permite_RA_status			= rs("permite_RA_status")
			.QtdeConsecutivaFalhaLogin = rs("QtdeConsecutivaFalhaLogin")
			.StLoginBloqueadoAutomatico = rs("StLoginBloqueadoAutomatico")
			.DataHoraBloqueadoAutomatico = rs("DataHoraBloqueadoAutomatico")
			.EnderecoIpBloqueadoAutomatico = Trim("" & rs("EnderecoIpBloqueadoAutomatico"))
			end with
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_orcamentista_e_indicador=True
end function



' ___________________________________________
' X _ S T A T U S _ P A G T O
'
function x_status_pagto(byval status)
dim s
	status = Trim("" & status)
	select case status
		case ST_PAGTO_PAGO: s="Pago"
		case ST_PAGTO_NAO_PAGO: s="Não-Pago"
		case ST_PAGTO_PARCIAL: s="Pago Parcial"
		case else s=""
		end select
	x_status_pagto = s
end function



' ___________________________________________
' X _ S T A T U S _ P A G T O _ C O R
'
function x_status_pagto_cor(byval status)
dim s
	status = Trim("" & status)
	select case status
		case ST_PAGTO_PAGO: s="green"
		case ST_PAGTO_NAO_PAGO: s="red"
		case ST_PAGTO_PARCIAL: s="deeppink"
		case else s=""
		end select
	x_status_pagto_cor = s
end function



' ___________________________________________
' X _ S T A T U S _ E N T R E G A
'
function x_status_entrega(byval status)
dim s
	status = Trim("" & status)
	select case status
		case ST_ENTREGA_ESPERAR: s="Esperar Mercadoria"
		case ST_ENTREGA_SPLIT_POSSIVEL: s="Split Possível"
		case ST_ENTREGA_SEPARAR: s="Separar Mercadoria"
		case ST_ENTREGA_A_ENTREGAR: s="A Entregar"
		case ST_ENTREGA_ENTREGUE: s="Entregue"
		case ST_ENTREGA_CANCELADO: s="Cancelado"
		case else s=""
		end select
	x_status_entrega=s
end function



' ___________________________________________
' X _ S T A T U S _ E N T R E G A _ C O R
'
function x_status_entrega_cor(byval status, byval id_pedido)
dim s_cor
dim r, s_sql
	status = Trim("" & status)
	select case status
		case ST_ENTREGA_ESPERAR: s_cor="deeppink"
		case ST_ENTREGA_SPLIT_POSSIVEL: s_cor="darkorange"
		case ST_ENTREGA_SEPARAR: s_cor="maroon"
		case ST_ENTREGA_A_ENTREGAR: s_cor="blue"
		case ST_ENTREGA_ENTREGUE: s_cor="green"
		case ST_ENTREGA_CANCELADO: s_cor="red"
		case else s_cor="black"
		end select
	
'	HÁ PRODUTOS DEVOLVIDOS?
	if status = ST_ENTREGA_ENTREGUE then
		id_pedido = Trim("" & id_pedido)
		s_sql = "SELECT Count(*) AS qtde FROM t_PEDIDO_ITEM_DEVOLVIDO" &_ 
				" WHERE (pedido='" & id_pedido & "')"
		set r = cn.Execute(s_sql)
		if Not r.Eof then
		'   ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
			if Cstr(r("qtde")) > Cstr(0) then
				s_cor="red"
				end if
			end if
		end if
		
	x_status_entrega_cor=s_cor
end function



' ___________________________________________
' IS PEDIDO ENCERRADO
'
function IsPedidoEncerrado(byval status)
	status = Trim("" & status)
	IsPedidoEncerrado=((status=ST_ENTREGA_ENTREGUE) OR (status=ST_ENTREGA_CANCELADO))
end function



' ___________________________________________
' IS PEDIDO SPLITABLE
'
function IsPedidoSplitable(byval status)
	IsPedidoSplitable = False
	status = Trim("" & status)
	if IsPedidoEncerrado(status) then exit function
	if status=ST_ENTREGA_ESPERAR then exit function
	if status=ST_ENTREGA_A_ENTREGAR then exit function
	IsPedidoSplitable = True
end function



' ___________________________________________
' IS PEDIDO CANCELAVEL
'
function IsPedidoCancelavel(byval status)
	IsPedidoCancelavel = False
	status = Trim("" & status)
	if IsPedidoEncerrado(status) then exit function
	IsPedidoCancelavel = True
end function



' ___________________________________________
' IS PEDIDO CANCELAVEL NA LOJA
'
function IsPedidoCancelavelNaLoja(byval strStEntrega, byval strStAnaliseCredito)
	IsPedidoCancelavelNaLoja = False
	strStEntrega = Trim("" & strStEntrega)
	strStAnaliseCredito = Trim("" & strStAnaliseCredito)
	if IsPedidoEncerrado(strStEntrega) then exit function
	if strStEntrega=ST_ENTREGA_A_ENTREGAR then exit function
	if strStAnaliseCredito=COD_AN_CREDITO_OK then exit function
	IsPedidoCancelavelNaLoja = True
end function



' ___________________________________________
' IS ENTREGA AGENDAVEL
'
function IsEntregaAgendavel(byval status)
	IsEntregaAgendavel = False
	status = Trim("" & status)
	if (status=ST_ENTREGA_A_ENTREGAR) OR (status=ST_ENTREGA_SEPARAR) then
		IsEntregaAgendavel = True
		end if
end function



' ___________________________________________
' IS PEDIDO ENTREGAVEL
'
function IsPedidoEntregavel(byval status)
	IsPedidoEntregavel = False
	status = Trim("" & status)
	if (status=ST_ENTREGA_A_ENTREGAR) then
		IsPedidoEntregavel = True
		end if
end function



' ___________________________________________
' IS PEDIDO ROMANEIO POSSIVEL
' 12.Nov.2013: o status 'ST_ENTREGA_SEPARAR' deixou de ser aceito como um status válido p/ esta operação.
function IsPedidoRomaneioPossivel(byval status)
	IsPedidoRomaneioPossivel = False
	status = Trim("" & status)
	if (status=ST_ENTREGA_A_ENTREGAR) then
		IsPedidoRomaneioPossivel = True
		end if
end function



' ___________________________________________
' X _ C O R _ I T E M
'
function x_cor_item(byval qtde, byval qtde_estoque_vendido, byval qtde_estoque_sem_presenca)
dim s
	s="black"
	
	if (qtde<=0) Or (qtde<>(qtde_estoque_vendido+qtde_estoque_sem_presenca)) then
		x_cor_item=s
		exit function
		end if

	if (qtde_estoque_vendido<>0) And (qtde_estoque_sem_presenca<>0) then
		s="darkorange"
	elseif (qtde_estoque_sem_presenca=0) then
		s="black"
	elseif (qtde_estoque_vendido=0) then
		s="red"
		end if
		
	x_cor_item=s
end function



' ___________________________________________
' GERA NUM PEDIDO FILHOTE
'
function gera_num_pedido_filhote(byval id_pedido, byref id_pedido_filhote, byref msg_erro)
dim s
dim r
dim s_sufixo
dim id_pedido_base
	gera_num_pedido_filhote = False
	id_pedido_filhote = ""
	msg_erro = ""
	
	id_pedido = Trim("" & id_pedido)
	id_pedido_base = retorna_num_pedido_base(id_pedido)
	
	s = "SELECT pedido FROM t_PEDIDO WHERE" & _
		" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "') ORDER BY pedido DESC"
	set r = cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if r.EOF then 
		msg_erro = "Pedido " & id_pedido_base & " não foi encontrado."
		if r.State <> 0 then r.Close
		set r = nothing
		end if
	
	s = Trim("" & r("pedido"))
	s_sufixo = retorna_sufixo_pedido_filhote(s)
	
	if s_sufixo = "" then s_sufixo = Chr(Asc("A")-1)
	s_sufixo = Ucase(Chr(Asc(s_sufixo)+1))
	if s_sufixo > "Z" then 
		msg_erro="Pedido " & id_pedido_base & " está com numeração de filhotes de pedido esgotada."
		if r.State <> 0 then r.Close
		set r = nothing
		exit function
		end if

	id_pedido_filhote = id_pedido_base & COD_SEPARADOR_FILHOTE & s_sufixo
	
	s = "SELECT pedido FROM t_PEDIDO WHERE" & _
		" (pedido='" & id_pedido_filhote & "')"
	if r.State <> 0 then r.Close
	set r=cn.Execute(s)
	if Not r.EOF then msg_erro="Número para filhote de pedido " & id_pedido_filhote & " já está em uso."
	if r.State <> 0 then r.Close
	set r=nothing
	
	if msg_erro <> "" then exit function	
	gera_num_pedido_filhote = True
end function



' ___________________________________________
' RECUPERA FAMILIA PEDIDO
'
function recupera_familia_pedido(byval id_pedido, byref v_pedido, byref msg_erro)
dim s
dim rs
dim id_pedido_base
	recupera_familia_pedido = False
	msg_erro = ""
	id_pedido = Trim("" & id_pedido)
	redim v_pedido(0)
	v_pedido(Ubound(v_pedido))=""
	id_pedido_base = retorna_num_pedido_base(id_pedido)
	s = "SELECT pedido FROM t_PEDIDO WHERE" & _
		" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')" & _
		" ORDER BY pedido"
	set rs=cn.Execute(s)
	if Err <> 0 then 
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if
	do while Not rs.EOF 
		if v_pedido(Ubound(v_pedido))<>"" then
			redim preserve v_pedido(Ubound(v_pedido)+1)
			v_pedido(Ubound(v_pedido))=""
			end if
		v_pedido(Ubound(v_pedido))=Trim("" & rs("pedido"))
		rs.MoveNext 
		loop
	if rs.State <> 0 then rs.Close
	set rs=nothing	
	recupera_familia_pedido = True
end function



' ___________________________________________
' RECUPERA FAMILIA PEDIDO DETALHE SPLIT
'
function recupera_familia_pedido_detalhe_split(byval id_pedido, byref v_pedido, byref msg_erro)
dim s
dim rs
dim id_pedido_base
	recupera_familia_pedido_detalhe_split = False
	msg_erro = ""
	id_pedido = Trim("" & id_pedido)
	redim v_pedido(0)
	set v_pedido(Ubound(v_pedido)) = new cl_PEDIDO_DETALHE_TIPO_SPLIT
	v_pedido(Ubound(v_pedido)).pedido = ""
	id_pedido_base = retorna_num_pedido_base(id_pedido)
	s = "SELECT" & _
			" pedido," & _
			" pedido_base," & _
			" st_entrega," & _
			" split_status," & _
			" st_auto_split," & _
			" id_nfe_emitente" & _
		" FROM t_PEDIDO" & _
		" WHERE" & _
			" (pedido_base = '" & id_pedido_base & "')" & _
		" ORDER BY" & _
			" pedido"
	set rs=cn.Execute(s)
	if Err <> 0 then 
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if
	do while Not rs.EOF
		if v_pedido(Ubound(v_pedido)).pedido<>"" then
			redim preserve v_pedido(Ubound(v_pedido)+1)
			set v_pedido(UBound(v_pedido)) = new cl_PEDIDO_DETALHE_TIPO_SPLIT
			v_pedido(Ubound(v_pedido)).pedido = ""
			end if
		with v_pedido(Ubound(v_pedido))
			.pedido = Trim("" & rs("pedido"))
			.pedido_base = Trim("" & rs("pedido_base"))
			.id_nfe_emitente = rs("id_nfe_emitente")
			.st_entrega = Trim("" & rs("st_entrega"))
			if Trim(.pedido) = Trim(.pedido_base) then
			'	PEDIDO-BASE
				.tipo_split = ""
			else
			'	PEDIDO-FILHOTE
				if (rs("split_status") = 1) And (rs("st_auto_split") = 0) then
					.tipo_split = TIPO_SPLIT__MANUAL
				elseif (rs("split_status") = 1) And (rs("st_auto_split") = 1) then
					.tipo_split = TIPO_SPLIT__AUTOMATICO
				else
					.tipo_split = ""
					end if
				end if
			end with
		rs.MoveNext
		loop
	if rs.State <> 0 then rs.Close
	set rs=nothing
	recupera_familia_pedido_detalhe_split = True
end function



' ___________________________________________
' LE ESTOQUE
'
function le_estoque(byval id_estoque, byref r_estoque, byref msg_erro)
dim s
dim rs
	le_estoque=False
	id_estoque=Trim("" & id_estoque)
	set r_estoque = New cl_ESTOQUE
	s="SELECT * FROM t_ESTOQUE WHERE (id_estoque='" & id_estoque & "')"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Lote do estoque nº " & id_estoque & " não está cadastrado."
		if rs.State <> 0 then rs.Close
		exit function
	else
		with r_estoque
			.id_estoque				= Trim("" & rs("id_estoque"))
			.data_entrada			= rs("data_entrada")
			.hora_entrada			= Trim("" & rs("hora_entrada"))
			.fabricante				= Trim("" & rs("fabricante"))
			.documento				= Trim("" & rs("documento"))
			.usuario				= Trim("" & rs("usuario"))
			.data_ult_movimento		= rs("data_ult_movimento")
			.kit					= rs("kit")
			.entrada_especial		= rs("entrada_especial")
			.devolucao_status		= rs("devolucao_status")
			.devolucao_data			= rs("devolucao_data")
			.devolucao_hora			= Trim("" & rs("devolucao_hora"))
			.devolucao_usuario		= Trim("" & rs("devolucao_usuario"))
			.devolucao_loja			= Trim("" & rs("devolucao_loja"))
			.devolucao_pedido		= Trim("" & rs("devolucao_pedido"))
			.devolucao_id_estoque	= Trim("" & rs("devolucao_id_estoque"))
			.obs					= Trim("" & rs("obs"))
			.id_nfe_emitente		= rs("id_nfe_emitente")
			end with
		end if	

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	le_estoque=True
end function

' ___________________________________________
' LE ESTOQUE_AGIO
'
function le_estoque_agio(byval id_estoque, byref r_estoque, byref msg_erro)
dim s
dim rs
	le_estoque_agio=False
	id_estoque=Trim("" & id_estoque)
	set r_estoque = New cl_ESTOQUE_AGIO
	s="SELECT * FROM t_ESTOQUE WHERE (id_estoque='" & id_estoque & "')"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Lote do estoque nº " & id_estoque & " não está cadastrado."
		if rs.State <> 0 then rs.Close
		exit function
	else
		with r_estoque
			.id_estoque				 		= Trim("" & rs("id_estoque"))
			.data_entrada			 		= rs("data_entrada")
			.hora_entrada		 			= Trim("" & rs("hora_entrada"))
			.fabricante			 			= Trim("" & rs("fabricante"))
			.documento				 		= Trim("" & rs("documento"))
			.usuario				 		= Trim("" & rs("usuario"))
			.data_ult_movimento		 		= rs("data_ult_movimento")
			.kit					 		= rs("kit")
			.entrada_especial	 			= rs("entrada_especial")
			.devolucao_status 				= rs("devolucao_status")
			.devolucao_data		 			= rs("devolucao_data")
			.devolucao_hora		 			= Trim("" & rs("devolucao_hora"))
			.devolucao_usuario	 			= Trim("" & rs("devolucao_usuario"))
			.devolucao_loja 				= Trim("" & rs("devolucao_loja"))
			.devolucao_pedido           	= Trim("" & rs("devolucao_pedido"))
			.devolucao_id_estoque          	= Trim("" & rs("devolucao_id_estoque"))
			.obs				           	= Trim("" & rs("obs"))
			.id_nfe_emitente           		= rs("id_nfe_emitente")
            .entrada_tipo                   = Trim("" & rs("entrada_tipo"))
            .perc_agio      		        = rs("perc_agio")
            .data_emissao_NF_entrada   		= rs("data_emissao_NF_entrada")
			end with
		end if	

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	le_estoque_agio=True
end function



' ___________________________________________
' LE ESTOQUE ITEM
'
function le_estoque_item(byval id_estoque, byref v_item, byref msg_erro)
dim s
dim rs
	le_estoque_item=False
	id_estoque=Trim("" & id_estoque)
	redim v_item(0)
	set v_item(0) = New cl_ITEM_ESTOQUE
	
	s="SELECT * FROM t_ESTOQUE_ITEM WHERE (id_estoque='" & id_estoque & "') ORDER BY sequencia"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Não há mercadorias cadastradas para o lote nº " & id_estoque & " do estoque."
	else
		do while Not rs.EOF 
			if Trim(v_item(Ubound(v_item)).produto)<>"" then
				redim preserve v_item(Ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ITEM_ESTOQUE
				end if
			with v_item(Ubound(v_item))
				.id_estoque				= Trim("" & rs("id_estoque"))
				.fabricante				= Trim("" & rs("fabricante"))
				.produto				= Trim("" & rs("produto"))
				.qtde					= rs("qtde")
				.qtde_utilizada			= rs("qtde_utilizada")
				.preco_fabricante		= rs("preco_fabricante")
				.vl_custo2				= rs("vl_custo2")
				.vl_BC_ICMS_ST			= rs("vl_BC_ICMS_ST")
				.vl_ICMS_ST				= rs("vl_ICMS_ST")
				.data_ult_movimento		= rs("data_ult_movimento")
				.sequencia				= rs("sequencia")
				.ncm					= rs("ncm")
				.cst					= rs("cst")
				end with
			rs.MoveNext 
			Loop
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	le_estoque_item=True
end function



' ___________________________________________
' LE ESTOQUE ITEM_XML
'
function le_estoque_item_xml(byval id_estoque, byref v_item, byref msg_erro)
dim s
dim rs
	le_estoque_item_xml=False
	id_estoque=Trim("" & id_estoque)
	redim v_item(0)
	set v_item(0) = New cl_ITEM_ESTOQUE_ENTRADA_XML
	
	s="SELECT * FROM t_ESTOQUE_ITEM WHERE (id_estoque='" & id_estoque & "') ORDER BY sequencia"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Não há mercadorias cadastradas para o lote nº " & id_estoque & " do estoque."
	else
		do while Not rs.EOF 
			if Trim(v_item(Ubound(v_item)).produto)<>"" then
				redim preserve v_item(Ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ITEM_ESTOQUE_ENTRADA_XML
				end if
			with v_item(Ubound(v_item))
				.id_estoque				= Trim("" & rs("id_estoque"))
				.fabricante				= Trim("" & rs("fabricante"))
				.produto				= Trim("" & rs("produto"))
				.qtde					= rs("qtde")
				.qtde_utilizada			= rs("qtde_utilizada")
				.preco_fabricante		= rs("preco_fabricante")
				.vl_custo2				= rs("vl_custo2")
				.vl_BC_ICMS_ST			= rs("vl_BC_ICMS_ST")
				.vl_ICMS_ST				= rs("vl_ICMS_ST")
				.data_ult_movimento		= rs("data_ult_movimento")
				.sequencia				= rs("sequencia")
				.ncm					= rs("ncm")
				.cst					= rs("cst")
                .ean    				= rs("ean")
                .vl_ipi                 = rs("vl_ipi")
                .aliq_ipi				= rs("aliq_ipi")
                .aliq_icms				= rs("aliq_icms")
                .vl_frete               = rs("vl_frete")
				end with
			rs.MoveNext 
			Loop
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	le_estoque_item_xml=True
end function




' ___________________________________________
' ORCAMENTO CALCULA TOTAL NF E RA
'
function orcamento_calcula_total_NF_e_RA(byval id_orcamento, byref vl_total_NF, byref vl_total_RA, byref msg_erro)
dim s
dim rs
	orcamento_calcula_total_NF_e_RA = False
	
	id_orcamento = Trim("" & id_orcamento)
	vl_total_NF = 0
	vl_total_RA = 0
	msg_erro = ""
	
	s = "SELECT orcamento FROM t_ORCAMENTO" & _
		" WHERE (orcamento='" & id_orcamento & "')"
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.Eof then
		msg_erro = "Orçamento " & id_orcamento & " não foi encontrado."
		exit function
		end if
		
	s = "SELECT SUM(qtde*preco_NF) AS total_NF," & _
		" SUM(qtde*(preco_NF-preco_venda)) AS total_RA" & _
		" FROM t_ORCAMENTO_ITEM INNER JOIN t_ORCAMENTO" & _
		" ON (t_ORCAMENTO_ITEM.orcamento=t_ORCAMENTO.orcamento)" & _
		" WHERE (t_ORCAMENTO.orcamento = '" & id_orcamento & "')"
	if rs.State <> 0 then rs.Close
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then
		if Not IsNull(rs("total_NF")) then vl_total_NF = rs("total_NF")
		if Not IsNull(rs("total_RA")) then vl_total_RA = rs("total_RA")
		end if

	orcamento_calcula_total_NF_e_RA = True
end function



' ___________________________________________
' CALCULA PAGAMENTOS
'
function calcula_pagamentos(byval id_pedido, _
							byref vl_TotalFamiliaPrecoVenda, _
							byref vl_TotalFamiliaPrecoNF, _
							byref vl_TotalFamiliaPago, _
							byref vl_TotalFamiliaDevolucaoPrecoVenda, _
							byref vl_TotalFamiliaDevolucaoPrecoNF, _
							byref st_pagto, _
							byref msg_erro)
dim s
dim rs
dim id_pedido_base
	calcula_pagamentos = False
	
	id_pedido = Trim("" & id_pedido)
	vl_TotalFamiliaPrecoVenda = 0
	vl_TotalFamiliaPrecoNF = 0
	vl_TotalFamiliaPago = 0
	vl_TotalFamiliaDevolucaoPrecoVenda = 0
	vl_TotalFamiliaDevolucaoPrecoNF = 0
	st_pagto = ""
	msg_erro = ""
	
	id_pedido_base = retorna_num_pedido_base(id_pedido)

	s = "SELECT" & _
			" pedido," & _
			" st_pagto" & _
		" FROM t_PEDIDO" & _
		" WHERE" & _
			" (pedido='" & id_pedido_base & "')"
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.Eof then
		msg_erro = "Pedido-base " & id_pedido_base & " não foi encontrado."
		exit function
		end if
		
	st_pagto = Trim("" & rs("st_pagto"))	
			
	s = "SELECT" & _
			" Coalesce(SUM(valor), 0) AS total" & _
		" FROM t_PEDIDO_PAGAMENTO" & _
		" WHERE" & _
			" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')"
	if rs.State <> 0 then rs.Close
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then vl_TotalFamiliaPago = rs("total")

	s = "SELECT" & _
			" Coalesce(SUM(qtde*preco_venda), 0) AS total," & _
			" Coalesce(SUM(qtde*preco_NF), 0) AS total_NF" & _
		" FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO" & _
			" ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
		" WHERE" & _
			" (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')" & _
			" AND (t_PEDIDO.pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')"
	if rs.State <> 0 then rs.Close
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then
		vl_TotalFamiliaPrecoVenda = rs("total")
		vl_TotalFamiliaPrecoNF = rs("total_NF")
		end if
		
	s = "SELECT" & _
			" Coalesce(SUM(qtde*preco_venda), 0) AS total," & _
			" Coalesce(SUM(qtde*preco_NF), 0) AS total_NF" & _
		" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
		" WHERE" & _
			" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')"
	if rs.State <> 0 then rs.Close
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then
		vl_TotalFamiliaDevolucaoPrecoVenda = rs("total")
		vl_TotalFamiliaDevolucaoPrecoNF = rs("total_NF")
		end if
	
	calcula_pagamentos = True
end function



' ___________________________________________
' CALCULA PAGAMENTOS CARTAO
'
function calcula_pagamentos_cartao(byval id_pedido, _
									byref vlPedidoTotalPagoCartao, _
									byref blnPedidoHouveEstorno, _
									byref vlFamiliaTotalPagoCartao, _
									byref blnFamiliaHouveEstorno, _
									byref strMsgErro)
dim s
dim rs
dim id_pedido_base

	calcula_pagamentos_cartao = False
	
	id_pedido = Trim("" & id_pedido)
	vlPedidoTotalPagoCartao = 0
	blnPedidoHouveEstorno = False
	vlFamiliaTotalPagoCartao = 0
	blnFamiliaHouveEstorno = False
	strMsgErro = ""
	
	id_pedido_base = retorna_num_pedido_base(id_pedido)

'	VALOR PAGO EM CARTÃO (PEDIDO)
	s = "SELECT" & _
			" Coalesce(SUM(valor), 0) AS total" & _
		" FROM t_PEDIDO_PAGAMENTO" & _
		" WHERE" & _
			" (pedido = '" & id_pedido & "')" & _
			" AND " & _
			"(" & _
				"(tipo_pagto = '" & COD_PAGTO_GW_BRASPAG_CLEARSALE & "')" & _
				" OR " & _
				"(tipo_pagto = '" & COD_PAGTO_BRASPAG & "')" & _
				" OR " & _
				"(tipo_pagto = '" & COD_PAGTO_CIELO & "')" & _
				" OR " & _
				"(tipo_pagto = '" & COD_PAGTO_VISANET & "')" & _
			")"
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then vlPedidoTotalPagoCartao = rs("total")
	
'	HOUVE ESTORNO (PEDIDO)?
	s = "SELECT TOP 1 " & _
			"*" & _
		" FROM t_PEDIDO_PAGAMENTO" & _
		" WHERE" & _
			" (pedido = '" & id_pedido & "')" & _
			" AND (valor < 0)"
	if rs.State <> 0 then rs.Close
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then blnPedidoHouveEstorno = True

'	VALOR PAGO EM CARTÃO (FAMÍLIA DE PEDIDOS)
	s = "SELECT" & _
			" Coalesce(SUM(valor), 0) AS total" & _
		" FROM t_PEDIDO_PAGAMENTO" & _
		" WHERE" & _
			" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')" & _
			" AND " & _
			"(" & _
				"(tipo_pagto = '" & COD_PAGTO_GW_BRASPAG_CLEARSALE & "')" & _
				" OR " & _
				"(tipo_pagto = '" & COD_PAGTO_BRASPAG & "')" & _
				" OR " & _
				"(tipo_pagto = '" & COD_PAGTO_CIELO & "')" & _
				" OR " & _
				"(tipo_pagto = '" & COD_PAGTO_VISANET & "')" & _
			")"
	if rs.State <> 0 then rs.Close
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then vlFamiliaTotalPagoCartao = rs("total")
	
'	HOUVE ESTORNO (FAMÍLIA DE PEDIDOS)?
	s = "SELECT TOP 1 " & _
			"*" & _
		" FROM t_PEDIDO_PAGAMENTO" & _
		" WHERE" & _
			" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')" & _
			" AND (valor < 0)"
	if rs.State <> 0 then rs.Close
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then blnFamiliaHouveEstorno = True
	
	calcula_pagamentos_cartao = True
end function



' ___________________________________________
' CALCULA VALOR TOTAL PEDIDO
'
function calcula_valor_total_pedido(byval id_pedido, byref vl_pedido, byref msg_erro)
dim s
dim rs
dim id_pedido_base
	calcula_valor_total_pedido = False
	
	id_pedido = Trim("" & id_pedido)
	vl_pedido = 0
	msg_erro = ""
	
	id_pedido_base = retorna_num_pedido_base(id_pedido)

	s = "SELECT SUM(qtde*preco_venda) AS total FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO" & _
		" ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
		" WHERE (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')" & _
		" AND (t_PEDIDO.pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')"
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then
		if Not IsNull(rs("total")) then vl_pedido = rs("total")
		end if
		
	if rs.State <> 0 then rs.Close
	calcula_valor_total_pedido = True
end function


' ___________________________________________________________
' CALCULA TOTAIS FAMILIA PEDIDO
'
function calcula_totais_familia_pedido(byval id_pedido, _
										byref vl_NF_deste_pedido, byref vl_venda_deste_pedido, _
										byref vl_total_NF_com_cancelado, byref vl_total_NF_sem_cancelado, _
										byref vl_total_venda_com_cancelado, byref vl_total_venda_sem_cancelado, _
										byref qtde_pedidos_familia, byref qtde_pedidos_cancelados, byref msg_erro)
dim s
dim rs
dim id_pedido_base
	calcula_totais_familia_pedido = False
	
	id_pedido = Trim("" & id_pedido)
	vl_NF_deste_pedido = 0
	vl_venda_deste_pedido = 0
	vl_total_NF_com_cancelado = 0
	vl_total_NF_sem_cancelado = 0
	vl_total_venda_com_cancelado = 0
	vl_total_venda_sem_cancelado = 0
	qtde_pedidos_familia = 0
	qtde_pedidos_cancelados = 0
	msg_erro = ""
	
	id_pedido_base = retorna_num_pedido_base(id_pedido)

	s = "SELECT" & _
			" t_PEDIDO.pedido," & _
			" t_PEDIDO.st_entrega," & _
			" Coalesce(SUM(qtde*preco_NF),0) AS vl_NF," & _
			" Coalesce(SUM(qtde*preco_venda),0) AS vl_venda" & _
		" FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM" & _
			" ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
		" WHERE" & _
			" (t_PEDIDO.pedido_base = '" & id_pedido_base & "')" & _
		" GROUP BY" & _
			" t_PEDIDO.pedido," & _
			" t_PEDIDO.st_entrega" & _
		" ORDER BY" & _
			" t_PEDIDO.pedido"
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	do while Not rs.Eof
		qtde_pedidos_familia = qtde_pedidos_familia + 1
		if Trim("" & rs("st_entrega")) = ST_ENTREGA_CANCELADO then qtde_pedidos_cancelados = qtde_pedidos_cancelados + 1

		if Trim("" & rs("pedido")) = id_pedido then
			vl_NF_deste_pedido = vl_NF_deste_pedido + rs("vl_NF")
			vl_venda_deste_pedido = vl_venda_deste_pedido + rs("vl_venda")
			end if

		vl_total_NF_com_cancelado = vl_total_NF_com_cancelado + rs("vl_NF")
		vl_total_venda_com_cancelado = vl_total_venda_com_cancelado + rs("vl_venda")

		if Trim("" & rs("st_entrega")) <> ST_ENTREGA_CANCELADO then
			vl_total_NF_sem_cancelado = vl_total_NF_sem_cancelado + rs("vl_NF")
			vl_total_venda_sem_cancelado = vl_total_venda_sem_cancelado + rs("vl_venda")
			end if

		rs.MoveNext
		loop
		
	if rs.State <> 0 then rs.Close
	calcula_totais_familia_pedido = True
end function


' ___________________________________________
' CALCULA VALOR EM PERDAS
'
function calcula_valor_em_perdas(byval id_pedido, byref vl_perda, byref msg_erro)
dim s
dim rs
dim id_pedido_base

	calcula_valor_em_perdas = False
	id_pedido = Trim("" & id_pedido)
	vl_perda = 0
	msg_erro = ""
	
	id_pedido_base = retorna_num_pedido_base(id_pedido)

	s = "SELECT SUM(valor) AS total FROM t_PEDIDO_PERDA" & _
		" WHERE (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')"
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then
		if Not IsNull(rs("total")) then vl_perda = rs("total")
		end if
		
	calcula_valor_em_perdas = True
end function



' ___________________________________________
' LE PEDIDO ITEM DEVOLVIDO
'
function le_pedido_item_devolvido(byval id_pedido, byref v_item_devolvido, byref msg_erro)
dim s
dim rs
	le_pedido_item_devolvido = False
	msg_erro = ""
	id_pedido=Trim("" & id_pedido)
	redim v_item_devolvido(0)
	set v_item_devolvido(0) = New cl_PEDIDO_ITEM_DEVOLVIDO
	
	s="SELECT * FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (pedido='" & id_pedido & "') ORDER BY devolucao_data, devolucao_hora"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	do while Not rs.EOF 
		if Trim(v_item_devolvido(Ubound(v_item_devolvido)).produto)<>"" then
			redim preserve v_item_devolvido(Ubound(v_item_devolvido)+1)
			set v_item_devolvido(ubound(v_item_devolvido)) = New cl_PEDIDO_ITEM_DEVOLVIDO
			end if
		with v_item_devolvido(Ubound(v_item_devolvido))
			.id						= Trim("" & rs("id"))
			.devolucao_data			= rs("devolucao_data")
			.devolucao_hora			= Trim("" & rs("devolucao_hora"))
			.devolucao_usuario		= Trim("" & rs("devolucao_usuario"))
			.pedido					= Trim("" & rs("pedido"))
			.fabricante				= Trim("" & rs("fabricante"))
			.produto				= Trim("" & rs("produto"))
			.qtde					= rs("qtde")
			.desc_dado				= rs("desc_dado")
			.preco_venda			= rs("preco_venda")
			.preco_NF				= rs("preco_NF")
			.preco_fabricante		= rs("preco_fabricante")
			.vl_custo2				= rs("vl_custo2")
			.preco_lista			= rs("preco_lista")
			.margem					= rs("margem")
			.desc_max				= rs("desc_max")
			.comissao				= rs("comissao")
			.descricao				= Trim("" & rs("descricao"))
			.descricao_html			= Trim("" & rs("descricao_html"))
			.ean					= Trim("" & rs("ean"))
			.grupo					= Trim("" & rs("grupo"))
            .subgrupo				= Trim("" & rs("subgrupo"))
			.peso					= rs("peso")
			.qtde_volumes			= rs("qtde_volumes")
			.abaixo_min_status		= rs("abaixo_min_status")
			.abaixo_min_autorizacao	= Trim("" & rs("abaixo_min_autorizacao"))
			.abaixo_min_autorizador	= Trim("" & rs("abaixo_min_autorizador"))
			.abaixo_min_superv_autorizador	= Trim("" & rs("abaixo_min_superv_autorizador"))
			.markup_fabricante		= rs("markup_fabricante")
			.motivo					= Trim("" & rs("motivo"))
			.cubagem				= rs("cubagem")
			.ncm					= Trim("" & rs("ncm"))
			.cst					= Trim("" & rs("cst"))
			.id_nfe_emitente		= rs("id_nfe_emitente")
			.NFe_serie_NF			= rs("NFe_serie_NF")
			.NFe_numero_NF			= rs("NFe_numero_NF")
			.dt_hr_anotacao_numero_NF = rs("dt_hr_anotacao_numero_NF")
			.usuario_anotacao_numero_NF = Trim("" & rs("usuario_anotacao_numero_NF"))
			.descontinuado			= Trim("" & rs("descontinuado"))
			.cod_produto_xml_fabricante = Trim("" & rs("cod_produto_xml_fabricante"))
			.cod_produto_alfanum_fabricante = Trim("" & rs("cod_produto_alfanum_fabricante"))
			.potencia_valor = rs("potencia_valor")
			.id_unidade_potencia = rs("id_unidade_potencia")
				'O campo 'StatusDescontoSuperior' é do tipo 'bit' e o ASP converte como se fosse um boolean, portanto, é feito um tratamento manual
				if rs("StatusDescontoSuperior") = 0 then .StatusDescontoSuperior = 0 else .StatusDescontoSuperior = 1
			.IdUsuarioDescontoSuperior = rs("IdUsuarioDescontoSuperior")
			.DataHoraDescontoSuperior = rs("DataHoraDescontoSuperior")
			end with
		rs.MoveNext
		Loop

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_pedido_item_devolvido=True
end function



' ___________________________________________
' LE PEDIDO ITEM DEVOLVIDO FAMILIA
'
function le_pedido_item_devolvido_familia(byval id_pedido, byref v_item_devolvido, byref msg_erro)
dim id_pedido_base
dim s
dim rs
	le_pedido_item_devolvido_familia = False
	msg_erro = ""
	id_pedido=Trim("" & id_pedido)
	id_pedido_base = retorna_num_pedido_base(id_pedido)
	redim v_item_devolvido(0)
	set v_item_devolvido(0) = New cl_PEDIDO_ITEM_DEVOLVIDO
	
	s="SELECT " & _
			"*" & _
		" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
		" WHERE" & _
			" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')" & _
		" ORDER BY" & _
			" devolucao_data," & _
			" devolucao_hora"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	do while Not rs.EOF 
		if Trim(v_item_devolvido(Ubound(v_item_devolvido)).produto)<>"" then
			redim preserve v_item_devolvido(Ubound(v_item_devolvido)+1)
			set v_item_devolvido(ubound(v_item_devolvido)) = New cl_PEDIDO_ITEM_DEVOLVIDO
			end if
		with v_item_devolvido(Ubound(v_item_devolvido))
			.id						= Trim("" & rs("id"))
			.devolucao_data			= rs("devolucao_data")
			.devolucao_hora			= Trim("" & rs("devolucao_hora"))
			.devolucao_usuario		= Trim("" & rs("devolucao_usuario"))
			.pedido					= Trim("" & rs("pedido"))
			.fabricante				= Trim("" & rs("fabricante"))
			.produto				= Trim("" & rs("produto"))
			.qtde					= rs("qtde")
			.desc_dado				= rs("desc_dado")
			.preco_venda			= rs("preco_venda")
			.preco_NF				= rs("preco_NF")
			.preco_fabricante		= rs("preco_fabricante")
			.vl_custo2				= rs("vl_custo2")
			.preco_lista			= rs("preco_lista")
			.margem					= rs("margem")
			.desc_max				= rs("desc_max")
			.comissao				= rs("comissao")
			.descricao				= Trim("" & rs("descricao"))
			.descricao_html			= Trim("" & rs("descricao_html"))
			.ean					= Trim("" & rs("ean"))
			.grupo					= Trim("" & rs("grupo"))
            .subgrupo				= Trim("" & rs("subgrupo"))
			.peso					= rs("peso")
			.qtde_volumes			= rs("qtde_volumes")
			.abaixo_min_status		= rs("abaixo_min_status")
			.abaixo_min_autorizacao	= Trim("" & rs("abaixo_min_autorizacao"))
			.abaixo_min_autorizador	= Trim("" & rs("abaixo_min_autorizador"))
			.abaixo_min_superv_autorizador	= Trim("" & rs("abaixo_min_superv_autorizador"))
			.markup_fabricante		= rs("markup_fabricante")
			.motivo					= Trim("" & rs("motivo"))
			.cubagem				= rs("cubagem")
			.ncm					= Trim("" & rs("ncm"))
			.cst					= Trim("" & rs("cst"))
			.id_nfe_emitente		= rs("id_nfe_emitente")
			.NFe_serie_NF			= rs("NFe_serie_NF")
			.NFe_numero_NF			= rs("NFe_numero_NF")
			.dt_hr_anotacao_numero_NF = rs("dt_hr_anotacao_numero_NF")
			.usuario_anotacao_numero_NF = Trim("" & rs("usuario_anotacao_numero_NF"))
			.descontinuado			= Trim("" & rs("descontinuado"))
			.cod_produto_xml_fabricante = Trim("" & rs("cod_produto_xml_fabricante"))
			.cod_produto_alfanum_fabricante = Trim("" & rs("cod_produto_alfanum_fabricante"))
			.potencia_valor = rs("potencia_valor")
			.id_unidade_potencia = rs("id_unidade_potencia")
				'O campo 'StatusDescontoSuperior' é do tipo 'bit' e o ASP converte como se fosse um boolean, portanto, é feito um tratamento manual
				if rs("StatusDescontoSuperior") = 0 then .StatusDescontoSuperior = 0 else .StatusDescontoSuperior = 1
			.IdUsuarioDescontoSuperior = rs("IdUsuarioDescontoSuperior")
			.DataHoraDescontoSuperior = rs("DataHoraDescontoSuperior")
			end with
		rs.MoveNext
		Loop

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_pedido_item_devolvido_familia=True
end function



' ___________________________________________
' LE PEDIDO PERDA
'
function le_pedido_perda(byval id_pedido, byref v_pedido_perda, byref msg_erro)
dim s
dim rs
	le_pedido_perda = False
	msg_erro = ""
	id_pedido=Trim("" & id_pedido)
	redim v_pedido_perda(0)
	set v_pedido_perda(0) = New cl_PEDIDO_PERDA
	
	s="SELECT * FROM t_PEDIDO_PERDA WHERE (pedido='" & id_pedido & "') ORDER BY data, hora"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	do while Not rs.EOF 
		if Trim(v_pedido_perda(Ubound(v_pedido_perda)).id)<>"" then
			redim preserve v_pedido_perda(Ubound(v_pedido_perda)+1)
			set v_pedido_perda(ubound(v_pedido_perda)) = New cl_PEDIDO_PERDA
			end if
		with v_pedido_perda(Ubound(v_pedido_perda))
			.id						= Trim("" & rs("id"))
			.pedido					= Trim("" & rs("pedido"))
			.data					= rs("data")
			.hora					= Trim("" & rs("hora"))
			.valor					= rs("valor")
			.obs					= Trim("" & rs("obs"))
			.usuario				= Trim("" & rs("usuario"))
			end with
		rs.MoveNext 
		Loop

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_pedido_perda=True
end function



' ___________________________________________
' X _ E S T O Q U E
'
function x_estoque(byval tipo_estoque)
dim s
	tipo_estoque = Trim("" & tipo_estoque)
	select case tipo_estoque
		case ID_ESTOQUE_VENDA: s="Venda"
		case ID_ESTOQUE_VENDIDO: s="Vendido"
		case ID_ESTOQUE_SEM_PRESENCA: s="Sem Presença"
		case ID_ESTOQUE_KIT: s="Kits Convertidos"
		case ID_ESTOQUE_SHOW_ROOM: s="Show-Room"
		case ID_ESTOQUE_DANIFICADOS: s="Danificado"
		case ID_ESTOQUE_DEVOLUCAO: s="Devolvido"
		case ID_ESTOQUE_ROUBO: s="Roubo/Dano"
		case ID_ESTOQUE_ENTREGUE: s="Entregue"
		case else s=""
		end select
	x_estoque=s
end function



' ___________________________________________
' X _ O P E R A C A O _ E S T O Q U E
'
function x_operacao_estoque(byval tipo_operacao)
dim s
	tipo_operacao = Trim("" & tipo_operacao)
	select case tipo_operacao
		case OP_ESTOQUE_ENTRADA: s="Entrada no Estoque"
		case OP_ESTOQUE_VENDA: s="Venda"
		case OP_ESTOQUE_CONVERSAO_KIT: s="Conversão de Kits"
		case OP_ESTOQUE_TRANSFERENCIA: s="Transferência entre Estoques"
		case OP_ESTOQUE_ENTREGA: s="Entrega de Pedido"
		case OP_ESTOQUE_DEVOLUCAO: s="Devolução"
		case else s=""
		end select
	x_operacao_estoque=s
end function



' ___________________________________________
' X _ O P E R A C A O _ L O G _ E S T O Q U E
'
function x_operacao_log_estoque(byval tipo_operacao)
dim s
	tipo_operacao = Trim("" & tipo_operacao)
	select case tipo_operacao
		case OP_ESTOQUE_LOG_ENTRADA: s="Entrada no estoque"
		case OP_ESTOQUE_LOG_VENDA: s="Venda"
		case OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA: s="Venda (sem presença)"
		case OP_ESTOQUE_LOG_CONVERSAO_KIT: s="Conversão de kit"
		case OP_ESTOQUE_LOG_ENTRADA_VIA_KIT: s="Entrada no estoque (conversão de kit)"
		case OP_ESTOQUE_LOG_TRANSFERENCIA: s="Transferência entre estoques"
		case OP_ESTOQUE_LOG_ENTREGA: s="Entrega de mercadoria"
		case OP_ESTOQUE_LOG_DEVOLUCAO: s="Devolução de mercadoria"
		case OP_ESTOQUE_LOG_ESTORNO: s="Estorno p/ estoque de venda"
		case OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA: s="Cancela do estoque s/ presença"
		case OP_ESTOQUE_LOG_SPLIT: s="Split de pedido"
		case OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE: s="Entrada no estoque: exclui"
		case OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM: s="Entrada no estoque: inclui item"
		case OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA: s="Entrada no estoque: aumenta qtde"
		case OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA: s="Entrada no estoque: diminui qtde"
		case OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA: s="Atende pedido em espera"
		case OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS: s="Transferência entre pedidos"
		case else s=""
		end select
	x_operacao_log_estoque=s
end function



' ___________________________________________
' LE ORCAMENTO
'
function le_orcamento(byval id_orcamento, byref r_orcamento, byref msg_erro)
dim s
dim rs
dim blnUsarMemorizacaoCompletaEnderecos

	le_orcamento = False
	msg_erro = ""
	id_orcamento=Trim("" & id_orcamento)
	set r_orcamento = New cl_ORCAMENTO
	s = "SELECT" & _
			" *" & _
			"," & montaSubqueryGetUsuarioContexto("InstaladorInstalaUsuarioUltAtualiz", "DecodNome_InstaladorInstalaUsuarioUltAtualiz") & _
			"," & montaSubqueryGetUsuarioContexto("GarantiaIndicadorUsuarioUltAtualiz", "DecodNome_GarantiaIndicadorUsuarioUltAtualiz") & _
			"," & montaSubqueryGetUsuarioContexto("etg_imediata_usuario", "DecodNome_etg_imediata_usuario") & _
			"," & montaSubqueryGetUsuarioContexto("PrevisaoEntregaUsuarioUltAtualiz", "DecodNome_PrevisaoEntregaUsuarioUltAtualiz") & _
		" FROM t_ORCAMENTO" & _
		" WHERE" & _
			" (orcamento='" & id_orcamento & "')"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	if rs.EOF then
		msg_erro="Orçamento nº " & id_orcamento & " não está cadastrado."
	else
		with r_orcamento
			.orcamento					= Trim("" & rs("orcamento"))
			.loja						= Trim("" & rs("loja"))
			.data						= rs("data")
			.hora						= Trim("" & rs("hora"))
			.id_cliente					= Trim("" & rs("id_cliente"))
			.midia						= Trim("" & rs("midia"))
			.servicos					= Trim("" & rs("servicos"))
			.vl_servicos				= rs("vl_servicos")
			.vendedor					= Trim("" & rs("vendedor"))
			.obs_1						= Trim("" & rs("obs_1"))
			.obs_2						= Trim("" & rs("obs_2"))
			.qtde_parcelas				= rs("qtde_parcelas")
			.forma_pagto				= Trim("" & rs("forma_pagto"))
			.st_orcamento				= Trim("" & rs("st_orcamento"))
			.cancelado_data				= rs("cancelado_data")
			.cancelado_usuario			= Trim("" & rs("cancelado_usuario"))
			.st_fechamento				= Trim("" & rs("st_fechamento"))
			.fechamento_data			= rs("fechamento_data")
			.fechamento_usuario			= Trim("" & rs("fechamento_usuario"))
			.loja_indicou				= Trim("" & rs("loja_indicou"))
			.comissao_loja_indicou		= rs("comissao_loja_indicou")
			.venda_externa				= rs("venda_externa")
			.tipo_parcelamento			= rs("tipo_parcelamento")
			.av_forma_pagto				= rs("av_forma_pagto")
			.pu_forma_pagto 			= rs("pu_forma_pagto")
			.pu_valor					= rs("pu_valor")
			.pu_vencto_apos				= rs("pu_vencto_apos")
			.pc_qtde_parcelas			= rs("pc_qtde_parcelas")
			.pc_valor_parcela			= rs("pc_valor_parcela")
			.pc_maquineta_qtde_parcelas = rs("pc_maquineta_qtde_parcelas")
			.pc_maquineta_valor_parcela = rs("pc_maquineta_valor_parcela")
			.pce_forma_pagto_entrada	= rs("pce_forma_pagto_entrada")
			.pce_forma_pagto_prestacao	= rs("pce_forma_pagto_prestacao")
			.pce_entrada_valor			= rs("pce_entrada_valor")
			.pce_prestacao_qtde			= rs("pce_prestacao_qtde")
			.pce_prestacao_valor		= rs("pce_prestacao_valor")
			.pce_prestacao_periodo		= rs("pce_prestacao_periodo")
			.pse_forma_pagto_prim_prest	= rs("pse_forma_pagto_prim_prest")
			.pse_forma_pagto_demais_prest = rs("pse_forma_pagto_demais_prest")
			.pse_prim_prest_valor		= rs("pse_prim_prest_valor")
			.pse_prim_prest_apos		= rs("pse_prim_prest_apos")
			.pse_demais_prest_qtde		= rs("pse_demais_prest_qtde")
			.pse_demais_prest_valor		= rs("pse_demais_prest_valor")
			.pse_demais_prest_periodo	= rs("pse_demais_prest_periodo")
			.custoFinancFornecTipoParcelamento	= Trim("" & rs("custoFinancFornecTipoParcelamento"))
			.custoFinancFornecQtdeParcelas		= rs("custoFinancFornecQtdeParcelas")
			.vl_total					= rs("vl_total")
			.vl_total_NF 				= rs("vl_total_NF")
			.vl_total_RA				= rs("vl_total_RA")
			.perc_RT					= rs("perc_RT")
			.orcamentista				= Trim("" & rs("orcamentista"))
			.st_orc_virou_pedido		= rs("st_orc_virou_pedido")
			.pedido						= Trim("" & rs("pedido"))

			.st_memorizacao_completa_enderecos = 0
			.endereco_logradouro = ""
			.endereco_numero = ""
			.endereco_complemento = ""
			.endereco_bairro = ""
			.endereco_cidade = ""
			.endereco_uf = ""
			.endereco_cep = ""
			.endereco_email = ""
			.endereco_email_xml = ""
			.endereco_nome = ""
			.endereco_nome_iniciais_em_maiusculas = ""
			.endereco_ddd_res = ""
			.endereco_tel_res = ""
			.endereco_ddd_com = ""
			.endereco_tel_com = ""
			.endereco_ramal_com = ""
			.endereco_ddd_cel = ""
			.endereco_tel_cel = ""
			.endereco_ddd_com_2 = ""
			.endereco_tel_com_2 = ""
			.endereco_ramal_com_2 = ""
			.endereco_tipo_pessoa = ""
			.endereco_cnpj_cpf = ""
			.endereco_contribuinte_icms_status = 0
			.endereco_produtor_rural_status = 0
			.endereco_ie = ""
			.endereco_rg = ""
			.endereco_contato = ""

			'O orçamento não armazenava o endereço de cobrança anteriormente da forma como ocorria no pedido
			if blnUsarMemorizacaoCompletaEnderecos then
				.st_memorizacao_completa_enderecos = rs("st_memorizacao_completa_enderecos")
				if CLng(.st_memorizacao_completa_enderecos) <> 0 then
					.endereco_logradouro			= Trim("" & rs("endereco_logradouro"))
					.endereco_numero				= Trim("" & rs("endereco_numero"))
					.endereco_complemento			= Trim("" & rs("endereco_complemento"))
					.endereco_bairro				= Trim("" & rs("endereco_bairro"))
					.endereco_cidade				= Trim("" & rs("endereco_cidade"))
					.endereco_uf					= Trim("" & rs("endereco_uf"))
					.endereco_cep					= Trim("" & rs("endereco_cep"))
					.endereco_email = Trim("" & rs("endereco_email"))
					.endereco_email_xml = Trim("" & rs("endereco_email_xml"))
					.endereco_nome = Trim("" & rs("endereco_nome"))
					.endereco_nome_iniciais_em_maiusculas = Trim("" & rs("endereco_nome_iniciais_em_maiusculas"))
					.endereco_ddd_res = Trim("" & rs("endereco_ddd_res"))
					.endereco_tel_res = Trim("" & rs("endereco_tel_res"))
					.endereco_ddd_com = Trim("" & rs("endereco_ddd_com"))
					.endereco_tel_com = Trim("" & rs("endereco_tel_com"))
					.endereco_ramal_com = Trim("" & rs("endereco_ramal_com"))
					.endereco_ddd_cel = Trim("" & rs("endereco_ddd_cel"))
					.endereco_tel_cel = Trim("" & rs("endereco_tel_cel"))
					.endereco_ddd_com_2 = Trim("" & rs("endereco_ddd_com_2"))
					.endereco_tel_com_2 = Trim("" & rs("endereco_tel_com_2"))
					.endereco_ramal_com_2 = Trim("" & rs("endereco_ramal_com_2"))
					.endereco_tipo_pessoa = Trim("" & rs("endereco_tipo_pessoa"))
					.endereco_cnpj_cpf = Trim("" & rs("endereco_cnpj_cpf"))
					.endereco_contribuinte_icms_status = rs("endereco_contribuinte_icms_status")
					.endereco_produtor_rural_status = rs("endereco_produtor_rural_status")
					.endereco_ie = Trim("" & rs("endereco_ie"))
					.endereco_rg = Trim("" & rs("endereco_rg"))
					.endereco_contato = Trim("" & rs("endereco_contato"))
					end if
				end if

			.EndEtg_endereco = ""
			.EndEtg_endereco_numero = ""
			.EndEtg_endereco_complemento = ""
			.EndEtg_bairro = ""
			.EndEtg_cidade = ""
			.EndEtg_uf = ""
			.EndEtg_cep = ""
			.EndEtg_email = ""
			.EndEtg_email_xml = ""
			.EndEtg_nome = ""
			.EndEtg_nome_iniciais_em_maiusculas = ""
			.EndEtg_ddd_res = ""
			.EndEtg_tel_res = ""
			.EndEtg_ddd_com = ""
			.EndEtg_tel_com = ""
			.EndEtg_ramal_com = ""
			.EndEtg_ddd_cel = ""
			.EndEtg_tel_cel = ""
			.EndEtg_ddd_com_2 = ""
			.EndEtg_tel_com_2 = ""
			.EndEtg_ramal_com_2 = ""
			.EndEtg_tipo_pessoa = ""
			.EndEtg_cnpj_cpf = ""
			.EndEtg_contribuinte_icms_status = 0
			.EndEtg_produtor_rural_status = 0
			.EndEtg_ie = ""
			.EndEtg_rg = ""
			.st_end_entrega				= rs("st_end_entrega")
			if CLng(.st_end_entrega) <> 0 then
				.EndEtg_endereco			= Trim("" & rs("EndEtg_endereco"))
				.EndEtg_endereco_numero		= Trim("" & rs("EndEtg_endereco_numero"))
				.EndEtg_endereco_complemento = Trim("" & rs("EndEtg_endereco_complemento"))
				.EndEtg_bairro				= Trim("" & rs("EndEtg_bairro"))
				.EndEtg_cidade				= Trim("" & rs("EndEtg_cidade"))
				.EndEtg_uf					= Trim("" & rs("EndEtg_uf"))
				.EndEtg_cep					= Trim("" & rs("EndEtg_cep"))
				if blnUsarMemorizacaoCompletaEnderecos and CLng(.st_memorizacao_completa_enderecos) <> 0 then
					.EndEtg_email = Trim("" & rs("EndEtg_email"))
					.EndEtg_email_xml = Trim("" & rs("EndEtg_email_xml"))
					.EndEtg_nome = Trim("" & rs("EndEtg_nome"))
					.EndEtg_nome_iniciais_em_maiusculas = Trim("" & rs("EndEtg_nome_iniciais_em_maiusculas"))
					.EndEtg_ddd_res = Trim("" & rs("EndEtg_ddd_res"))
					.EndEtg_tel_res = Trim("" & rs("EndEtg_tel_res"))
					.EndEtg_ddd_com = Trim("" & rs("EndEtg_ddd_com"))
					.EndEtg_tel_com = Trim("" & rs("EndEtg_tel_com"))
					.EndEtg_ramal_com = Trim("" & rs("EndEtg_ramal_com"))
					.EndEtg_ddd_cel = Trim("" & rs("EndEtg_ddd_cel"))
					.EndEtg_tel_cel = Trim("" & rs("EndEtg_tel_cel"))
					.EndEtg_ddd_com_2 = Trim("" & rs("EndEtg_ddd_com_2"))
					.EndEtg_tel_com_2 = Trim("" & rs("EndEtg_tel_com_2"))
					.EndEtg_ramal_com_2 = Trim("" & rs("EndEtg_ramal_com_2"))
					.EndEtg_tipo_pessoa = Trim("" & rs("EndEtg_tipo_pessoa"))
					.EndEtg_cnpj_cpf = Trim("" & rs("EndEtg_cnpj_cpf"))
					.EndEtg_contribuinte_icms_status = rs("EndEtg_contribuinte_icms_status")
					.EndEtg_produtor_rural_status = rs("EndEtg_produtor_rural_status")
					.EndEtg_ie = Trim("" & rs("EndEtg_ie"))
					.EndEtg_rg = Trim("" & rs("EndEtg_rg"))
					end if
				end if

			.st_etg_imediata			= rs("st_etg_imediata")
			.etg_imediata_data			= rs("etg_imediata_data")
			if Trim("" & rs("DecodNome_etg_imediata_usuario")) = "" then
				.etg_imediata_usuario		= Trim("" & rs("etg_imediata_usuario"))
			else
				.etg_imediata_usuario = Left(Trim("" & rs("DecodNome_etg_imediata_usuario")), MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR)
				if Left(Trim("" & rs("etg_imediata_usuario")), 3) = "[" & Cstr(COD_USUARIO_CONTEXTO__VENDEDOR_DO_PARCEIRO) & "]" then
					.etg_imediata_usuario = "[VP] " & .etg_imediata_usuario
					end if
				end if
			.etg_imediata_usuario_RawData = Trim("" & rs("etg_imediata_usuario"))
			.PrevisaoEntregaData = rs("PrevisaoEntregaData")
			if Trim("" & rs("DecodNome_PrevisaoEntregaUsuarioUltAtualiz")) = "" then
				.PrevisaoEntregaUsuarioUltAtualiz = Trim("" & rs("PrevisaoEntregaUsuarioUltAtualiz"))
			else
				.PrevisaoEntregaUsuarioUltAtualiz = Left(Trim("" & rs("DecodNome_PrevisaoEntregaUsuarioUltAtualiz")), MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR)
				if Left(Trim("" & rs("PrevisaoEntregaUsuarioUltAtualiz")), 3) = "[" & Cstr(COD_USUARIO_CONTEXTO__VENDEDOR_DO_PARCEIRO) & "]" then
					.PrevisaoEntregaUsuarioUltAtualiz = "[VP] " & .PrevisaoEntregaUsuarioUltAtualiz
					end if
				end if
			.PrevisaoEntregaUsuarioUltAtualiz_RawData = Trim("" & rs("PrevisaoEntregaUsuarioUltAtualiz"))
			.PrevisaoEntregaDtHrUltAtualiz = rs("PrevisaoEntregaDtHrUltAtualiz")
			.StBemUsoConsumo			= rs("StBemUsoConsumo")
			.InstaladorInstalaStatus	= rs("InstaladorInstalaStatus")
			if Trim("" & rs("DecodNome_InstaladorInstalaUsuarioUltAtualiz")) = "" then
				.InstaladorInstalaUsuarioUltAtualiz = Trim("" & rs("InstaladorInstalaUsuarioUltAtualiz"))
			else
				.InstaladorInstalaUsuarioUltAtualiz = Left(Trim("" & rs("DecodNome_InstaladorInstalaUsuarioUltAtualiz")), MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR)
				if Left(Trim("" & rs("InstaladorInstalaUsuarioUltAtualiz")), 3) = "[" & Cstr(COD_USUARIO_CONTEXTO__VENDEDOR_DO_PARCEIRO) & "]" then
					.InstaladorInstalaUsuarioUltAtualiz = "[VP] " & .InstaladorInstalaUsuarioUltAtualiz
					end if
				end if
			.InstaladorInstalaUsuarioUltAtualiz_RawData = Trim("" & rs("InstaladorInstalaUsuarioUltAtualiz"))
			.InstaladorInstalaDtHrUltAtualiz = rs("InstaladorInstalaDtHrUltAtualiz")
			.GarantiaIndicadorStatus	= rs("GarantiaIndicadorStatus")
			if Trim("" & rs("DecodNome_GarantiaIndicadorUsuarioUltAtualiz")) = "" then
				.GarantiaIndicadorUsuarioUltAtualiz = rs("GarantiaIndicadorUsuarioUltAtualiz")
			else
				.GarantiaIndicadorUsuarioUltAtualiz = Left(Trim("" & rs("DecodNome_GarantiaIndicadorUsuarioUltAtualiz")), MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR)
				if Left(Trim("" & rs("GarantiaIndicadorUsuarioUltAtualiz")), 3) = "[" & Cstr(COD_USUARIO_CONTEXTO__VENDEDOR_DO_PARCEIRO) & "]" then
					.GarantiaIndicadorUsuarioUltAtualiz = "[VP] " & .GarantiaIndicadorUsuarioUltAtualiz
					end if
				end if
			.GarantiaIndicadorUsuarioUltAtualiz_RawData = rs("GarantiaIndicadorUsuarioUltAtualiz")
			.GarantiaIndicadorDtHrUltAtualiz = rs("GarantiaIndicadorDtHrUltAtualiz")
			.perc_desagio_RA_liquida	= rs("perc_desagio_RA_liquida")
			.permite_RA_status			= rs("permite_RA_status")
			.st_violado_permite_RA_status		= rs("st_violado_permite_RA_status")
			.dt_hr_violado_permite_RA_status	= rs("dt_hr_violado_permite_RA_status")
			.usuario_violado_permite_RA_status	= Trim("" & rs("usuario_violado_permite_RA_status"))
            .EndEtg_obs                 = Trim("" & rs("EndEtg_obs"))
            .EndEtg_cod_justificativa   = Trim("" & rs("EndEtg_cod_justificativa"))
			.sistema_responsavel_cadastro = rs("sistema_responsavel_cadastro")
			.sistema_responsavel_atualizacao = rs("sistema_responsavel_atualizacao")
			.IdOrcamentoCotacao = rs("IdOrcamentoCotacao")
			.IdIndicadorVendedor = rs("IdIndicadorVendedor")
			.perc_max_comissao_padrao = rs("perc_max_comissao_padrao")
			.perc_max_comissao_e_desconto_padrao = rs("perc_max_comissao_e_desconto_padrao")
			.InstaladorInstalaIdTipoUsuarioContexto = rs("InstaladorInstalaIdTipoUsuarioContexto")
			.InstaladorInstalaIdUsuarioUltAtualiz = rs("InstaladorInstalaIdUsuarioUltAtualiz")
			.GarantiaIndicadorIdTipoUsuarioContexto = rs("GarantiaIndicadorIdTipoUsuarioContexto")
			.GarantiaIndicadorIdUsuarioUltAtualiz = rs("GarantiaIndicadorIdUsuarioUltAtualiz")
			.EtgImediataIdTipoUsuarioContexto = rs("EtgImediataIdTipoUsuarioContexto")
			.EtgImediataIdUsuarioUltAtualiz = rs("EtgImediataIdUsuarioUltAtualiz")
			.PrevisaoEntregaIdTipoUsuarioContexto = rs("PrevisaoEntregaIdTipoUsuarioContexto")
			.PrevisaoEntregaIdUsuarioUltAtualiz = rs("PrevisaoEntregaIdUsuarioUltAtualiz")
			.UsuarioCadastroIdTipoUsuarioContexto = rs("UsuarioCadastroIdTipoUsuarioContexto")
			.UsuarioCadastroId = rs("UsuarioCadastroId")
			end with
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_orcamento=True
end function



' ___________________________________________
' LE ORCAMENTO ITEM
'
function le_orcamento_item(byval id_orcamento, byref v_orcamento_item, byref msg_erro)
dim s
dim rs
	le_orcamento_item = False
	msg_erro = ""
	id_orcamento=Trim("" & id_orcamento)
	redim v_orcamento_item(0)
	set v_orcamento_item(0) = New cl_ITEM_ORCAMENTO
	
	s="SELECT * FROM t_ORCAMENTO_ITEM WHERE (orcamento='" & id_orcamento & "') ORDER BY sequencia"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Não há itens cadastrados para o orçamento nº " & id_orcamento & "."
	else
		do while Not rs.EOF 
			if Trim(v_orcamento_item(Ubound(v_orcamento_item)).produto)<>"" then
				redim preserve v_orcamento_item(Ubound(v_orcamento_item)+1)
				set v_orcamento_item(ubound(v_orcamento_item)) = New cl_ITEM_ORCAMENTO
				end if
			with v_orcamento_item(Ubound(v_orcamento_item))
				.orcamento				= Trim("" & rs("orcamento"))
				.fabricante				= Trim("" & rs("fabricante"))
				.produto				= Trim("" & rs("produto"))
				.qtde					= rs("qtde")
				.qtde_spe				= rs("qtde_spe")
				.desc_dado				= rs("desc_dado")
				.preco_venda			= rs("preco_venda")
				.preco_NF				= rs("preco_NF")
				.preco_fabricante		= rs("preco_fabricante")
				.vl_custo2				= rs("vl_custo2")
				.preco_lista			= rs("preco_lista")
				.margem					= rs("margem")
				.desc_max				= rs("desc_max")
				.comissao				= rs("comissao")
				.descricao				= Trim("" & rs("descricao"))
				.descricao_html			= Trim("" & rs("descricao_html"))
				.obs					= Trim("" & rs("obs"))
				.ean					= Trim("" & rs("ean"))
				.grupo					= Trim("" & rs("grupo"))
                .subgrupo				= Trim("" & rs("subgrupo"))
				.peso					= rs("peso")
				.qtde_volumes			= rs("qtde_volumes")
				.abaixo_min_status		= rs("abaixo_min_status")
				.abaixo_min_autorizacao	= Trim("" & rs("abaixo_min_autorizacao"))
				.abaixo_min_autorizador	= Trim("" & rs("abaixo_min_autorizador"))
				.abaixo_min_superv_autorizador	= Trim("" & rs("abaixo_min_superv_autorizador"))
				.markup_fabricante		= rs("markup_fabricante")
				.sequencia				= rs("sequencia")
				.custoFinancFornecCoeficiente = rs("custoFinancFornecCoeficiente")
				.custoFinancFornecPrecoListaBase = rs("custoFinancFornecPrecoListaBase")
				.cubagem				= rs("cubagem")
				.ncm					= Trim("" & rs("ncm"))
				.cst					= Trim("" & rs("cst"))
				.descontinuado			= Trim("" & rs("descontinuado"))
				'O campo 'StatusDescontoSuperior' é do tipo 'bit' e o ASP converte como se fosse um boolean, portanto, é feito um tratamento manual
				if rs("StatusDescontoSuperior") = 0 then .StatusDescontoSuperior = 0 else .StatusDescontoSuperior = 1
				.IdUsuarioDescontoSuperior = rs("IdUsuarioDescontoSuperior")
				.DataHoraDescontoSuperior = rs("DataHoraDescontoSuperior")
				end with
			rs.MoveNext 
			Loop
		end if	

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_orcamento_item=True
end function



' ___________________________________________
' X _ S T _ O R C A M E N T O
'
function x_st_orcamento(byval status)
dim s
	status = Trim("" & status)
	select case status
		case ST_ORCAMENTO_CANCELADO: s="Cancelado"
		case else s=""
		end select
	x_st_orcamento=s
end function



' ___________________________________________
' X _ S T _ O R C A M E N T O _ C O R
'
function x_st_orcamento_cor(byval status)
dim s_cor
	status = Trim("" & status)
	select case status
		case ST_ORCAMENTO_CANCELADO: s_cor="red"
		case else s_cor="black"
		end select
	x_st_orcamento_cor=s_cor
end function



' ___________________________________________
' IS ORCAMENTO CANCELAVEL
'
function IsOrcamentoCancelavel(byval status)
	IsOrcamentoCancelavel = False
	status = Trim("" & status)
	if status = ST_ORCAMENTO_CANCELADO then exit function
	IsOrcamentoCancelavel = True
end function



' ___________________________________________
' IS ORCAMENTO APTO VIRAR PEDIDO
'
function IsOrcamentoAptoVirarPedido(byval status)
	IsOrcamentoAptoVirarPedido = False
	status = Trim("" & status)
	if status = ST_ORCAMENTO_CANCELADO then exit function
	IsOrcamentoAptoVirarPedido = True
end function



' ___________________________________________
' LE ORCAMENTO COTACAO
'
function le_orcamento_cotacao(byval IdOrcamentoCotacao, byref r_orcamento_cotacao, byref msg_erro)
dim s
dim rs

	le_orcamento_cotacao = False
	msg_erro = ""
	IdOrcamentoCotacao = Trim("" & IdOrcamentoCotacao)
	set r_orcamento_cotacao = New cl_ORCAMENTO_COTACAO

	if IdOrcamentoCotacao = "" then
		r_orcamento_cotacao.Id = 0
		exit function
		end if

	s = "SELECT" & _
			" *" & _
		" FROM t_ORCAMENTO_COTACAO" & _
		" WHERE" & _
			" (Id = " & IdOrcamentoCotacao & ")"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.EOF then
		msg_erro="Orçamento/cotação nº " & IdOrcamentoCotacao & " não está cadastrado."
	else
		with r_orcamento_cotacao
			.Id = rs("Id")
			.Loja = Trim("" & rs("Loja"))
			.NomeCliente = Trim("" & rs("NomeCliente"))
			.NomeObra = Trim("" & rs("NomeObra"))
			.IdVendedor = rs("IdVendedor")
			.IdIndicador = rs("IdIndicador")
			.IdIndicadorVendedor = rs("IdIndicadorVendedor")
			.Email = Trim("" & rs("Email"))
			.Telefone = Trim("" & rs("Telefone"))
			.AceiteWhatsApp = rs("AceiteWhatsApp")
			.UF = Trim("" & rs("UF"))
			.TipoCliente = Trim("" & rs("TipoCliente"))
			.ContribuinteIcms = rs("ContribuinteIcms")
			.Validade = rs("Validade")
			.ValidadeAnterior = rs("ValidadeAnterior")
			.QtdeRenovacao = rs("QtdeRenovacao")
			.IdUsuarioUltRenovacao = rs("IdUsuarioUltRenovacao")
			.DataHoraUltRenovacao = rs("DataHoraUltRenovacao")
			.Observacao = Trim("" & rs("Observacao"))
			.InstaladorInstalaStatus = rs("InstaladorInstalaStatus")
			.GarantiaIndicadorStatus = rs("GarantiaIndicadorStatus")
			.StEtgImediata = rs("StEtgImediata")
			.PrevisaoEntregaData = rs("PrevisaoEntregaData")
			.Status = rs("Status")
			.IdTipoUsuarioContextoUltStatus = rs("IdTipoUsuarioContextoUltStatus")
			.IdUsuarioUltStatus = rs("IdUsuarioUltStatus")
			.DataUltStatus = rs("DataUltStatus")
			.DataHoraUltStatus = rs("DataHoraUltStatus")
			.VersaoPoliticaCredito = Trim("" & rs("VersaoPoliticaCredito"))
			.VersaoPoliticaPrivacidade = Trim("" & rs("VersaoPoliticaPrivacidade"))
			.IdOrcamento = Trim("" & rs("IdOrcamento"))
			.IdPedido = Trim("" & rs("IdPedido"))
			.perc_max_comissao_padrao = rs("perc_max_comissao_padrao")
			.perc_max_comissao_e_desconto_padrao = rs("perc_max_comissao_e_desconto_padrao")
			.IdTipoUsuarioContextoCadastro = rs("IdTipoUsuarioContextoCadastro")
			.IdUsuarioCadastro = rs("IdUsuarioCadastro")
			.DataCadastro = rs("DataCadastro")
			.DataHoraCadastro = rs("DataHoraCadastro")
			.IdTipoUsuarioContextoUltAtualizacao = rs("IdTipoUsuarioContextoUltAtualizacao")
			.IdUsuarioUltAtualizacao = rs("IdUsuarioUltAtualizacao")
			.DataHoraUltAtualizacao = rs("DataHoraUltAtualizacao")
			.InstaladorInstalaIdTipoUsuarioContexto = rs("InstaladorInstalaIdTipoUsuarioContexto")
			.InstaladorInstalaIdUsuarioUltAtualiz = rs("InstaladorInstalaIdUsuarioUltAtualiz")
			.InstaladorInstalaDtHrUltAtualiz = rs("InstaladorInstalaDtHrUltAtualiz")
			.GarantiaIndicadorIdTipoUsuarioContexto = rs("GarantiaIndicadorIdTipoUsuarioContexto")
			.GarantiaIndicadorIdUsuarioUltAtualiz = rs("GarantiaIndicadorIdUsuarioUltAtualiz")
			.GarantiaIndicadorDtHrUltAtualiz = rs("GarantiaIndicadorDtHrUltAtualiz")
			.EtgImediataIdTipoUsuarioContexto = rs("EtgImediataIdTipoUsuarioContexto")
			.EtgImediataIdUsuarioUltAtualiz = rs("EtgImediataIdUsuarioUltAtualiz")
			.EtgImediataDtHrUltAtualiz = rs("EtgImediataDtHrUltAtualiz")
			.PrevisaoEntregaIdTipoUsuarioContexto = rs("PrevisaoEntregaIdTipoUsuarioContexto")
			.PrevisaoEntregaIdUsuarioUltAtualiz = rs("PrevisaoEntregaIdUsuarioUltAtualiz")
			.PrevisaoEntregaDtHrUltAtualiz = rs("PrevisaoEntregaDtHrUltAtualiz")
			.IdTipoUsuarioContextoUltRenovacao = rs("IdTipoUsuarioContextoUltRenovacao")
			end with
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_orcamento_cotacao=True
end function



' ___________________________________________________
' D E S C R I C A O _ A N A L I S E _ C R E D I T O
'
function descricao_analise_credito(byval codigo)
dim s
	codigo = Trim("" & codigo)
	select case codigo
		case COD_AN_CREDITO_ST_INICIAL: s="Aguardando Análise Inicial"
		case COD_AN_CREDITO_PENDENTE: s="Pendente"
		case COD_AN_CREDITO_PENDENTE_VENDAS: s="Pendente Vendas"
		case COD_AN_CREDITO_PENDENTE_ENDERECO: s="Pendente Endereço"
		case COD_AN_CREDITO_OK: s="Crédito OK"
		case COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO: s="Crédito OK (aguardando depósito)"
		case COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO: s="Crédito OK (depósito aguardando desbloqueio)"
		case COD_AN_CREDITO_NAO_ANALISADO: s="Pedido Sem Análise de Crédito"
		case COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV: s = "Crédito OK (aguardando pagto boleto AV)"
		case COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO: s = "Pendente - Pagto Antecipado Boleto"
		case else s=""
		end select
	descricao_analise_credito=s
end function



' ___________________________________________________
' X _ A N A L I S E _ C R E D I T O
'
function x_analise_credito(byval codigo)
dim s
	codigo = Trim("" & codigo)
	select case codigo
		case COD_AN_CREDITO_ST_INICIAL: s=""
		case COD_AN_CREDITO_PENDENTE: s="Pendente"
		case COD_AN_CREDITO_PENDENTE_VENDAS: s="Pendente Vendas"
		case COD_AN_CREDITO_PENDENTE_ENDERECO: s="Pendente Endereço"
		case COD_AN_CREDITO_OK: s="Crédito OK"
		case COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO: s="Crédito OK (aguardando depósito)"
		case COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO: s="Crédito OK (depósito aguardando desbloqueio)"
		case COD_AN_CREDITO_NAO_ANALISADO: s=""
		case COD_AN_CREDITO_PENDENTE_CARTAO: s="Pendente Cartão de Crédito"
		case COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV: s = "Crédito OK (aguardando pagto boleto AV)"
		case COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO: s = "Pendente - Pagto Antecipado Boleto"
		case else s=""
		end select
	x_analise_credito=s
end function



' ___________________________________________
' X _ A N A L I S E _ C R E D I T O _ C O R
'
function x_analise_credito_cor(byval codigo)
dim s_cor
	codigo = Trim("" & codigo)
	select case codigo
		case COD_AN_CREDITO_PENDENTE: s_cor="red"
		case COD_AN_CREDITO_PENDENTE_VENDAS: s_cor="red"
		case COD_AN_CREDITO_PENDENTE_ENDERECO: s_cor="red"
		case COD_AN_CREDITO_OK: s_cor="green"
		case COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO: s_cor="darkorange"
		case COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO: s_cor="darkorange"
		case COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV: s_cor="darkorange"
		case COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO: s_cor="blue"
		case else s_cor="black"
		end select
	x_analise_credito_cor=s_cor
end function



function decodifica_etg_imediata(byval codigo)
dim s
	codigo = Trim("" & codigo)
	select case codigo
		case COD_ETG_IMEDIATA_ST_INICIAL: s=""
		case COD_ETG_IMEDIATA_NAO: s="Não"
		case COD_ETG_IMEDIATA_SIM: s="Sim"
		case COD_ETG_IMEDIATA_NAO_DEFINIDO: s=""
		case else s=""
		end select
	decodifica_etg_imediata=s
end function



' ___________________________________________________
' X _ O P C A O _ F O R M A _ P A G A M E N T O
'
function x_opcao_forma_pagamento(byval codigo)
dim s
	codigo = Trim("" & codigo)
	select case codigo
		case ID_FORMA_PAGTO_DINHEIRO: s="Dinheiro"
		case ID_FORMA_PAGTO_DEPOSITO: s="Depósito"
		case ID_FORMA_PAGTO_CHEQUE: s="Cheque"
		case ID_FORMA_PAGTO_BOLETO: s="Boleto"
		case ID_FORMA_PAGTO_CARTAO: s="Cartão (internet)"
		case ID_FORMA_PAGTO_CARTAO_MAQUINETA: s="Cartão (maquineta)"
		case ID_FORMA_PAGTO_BOLETO_AV: s="Boleto AV"
		case else s=""
		end select
	x_opcao_forma_pagamento=s
end function



' ___________________________________________________
' X _ T I P O _ P A R C E L A M E N T O
'
function x_tipo_parcelamento(byval codigo)
dim s
	codigo = Trim("" & codigo)
	select case codigo
		case COD_FORMA_PAGTO_A_VISTA: s="À Vista"
		case COD_FORMA_PAGTO_PARCELA_UNICA: s="Parcela Única"
		case COD_FORMA_PAGTO_PARCELADO_CARTAO: s="Parcelado no Cartão (internet)"
		case COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA: s="Parcelado no Cartão (maquineta)"
		case COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA: s="Parcelado com Entrada"
		case COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA: s="Parcelado sem Entrada"
		case else s=""
		end select
	x_tipo_parcelamento=s
end function



' _______________________________________________________
' X _ O R C A M E N T I S T A _ E _ I N D I C A D O R
'
function x_orcamentista_e_indicador(byval l)
dim r
	l = Trim("" & l)
	set r = cn.Execute("SELECT razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido = '" & l & "'")
	if not r.eof then x_orcamentista_e_indicador = Trim("" & r("razao_social_nome_iniciais_em_maiusculas"))
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________________________________________
' X _ B A N C O
'
function x_banco(byval l)
dim r
	l = Trim("" & l)
	if l = "" then exit function
	set r = cn.Execute("SELECT descricao FROM t_BANCO WHERE (CONVERT(smallint, codigo) = " & l & ")")
	if not r.eof then x_banco = Trim("" & r("descricao"))
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________________________________________
' GET REGISTRO T_PARAMETRO
'
function get_registro_t_parametro(ByVal id_registro)
dim r, rx, s
	set rx = New cl_t_PARAMETRO
	limpa_cl_t_PARAMETRO rx
	
'	LEMBRANDO QUE SE TRATA DE UM PONTEIRO P/ O OBJETO
	set get_registro_t_parametro = rx
	
	id_registro = Trim("" & id_registro)
	s = "SELECT " & _
			"*" & _
		" FROM t_PARAMETRO" & _
		" WHERE" & _
			" (id = '" & id_registro & "')"
	set r = cn.Execute(s)
	if Not r.Eof then
		rx.id = Trim("" & r("id"))
		rx.campo_inteiro = r("campo_inteiro")
		rx.campo_monetario = r("campo_monetario")
		rx.campo_real = r("campo_real")
		rx.campo_data = r("campo_data")
		rx.campo_texto = "" & r("campo_texto")
		rx.campo_2_texto = "" & r("campo_2_texto")
		rx.dt_hr_ult_atualizacao = r("dt_hr_ult_atualizacao")
		rx.usuario_ult_atualizacao = Trim("" & r("usuario_ult_atualizacao"))
		end if
	
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________________________________________
' L E _ P A R A M E T R O _ B D
'
function le_parametro_bd(byval id_nsu, byref msg_erro)
dim s
dim rs
	le_parametro_bd = ""
	msg_erro = ""
	
	s = "SELECT nsu FROM t_CONTROLE WHERE (id_nsu = '" & id_nsu & "')"
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then
		if Not IsNull(rs("nsu")) then le_parametro_bd = "" & rs("nsu")
		end if
		
	if rs.State <> 0 then rs.Close
end function



' _______________________________________________________
' GET DEFAULT VALOR TEXTO BD
'
function get_default_valor_texto_bd(byval usuario, byval nome_chave)
dim s
dim rs
	get_default_valor_texto_bd = ""
	s = "SELECT valor_default_texto FROM t_DEFAULT WHERE (usuario = '" & usuario & "') AND (nome_chave = '" & nome_chave & "')"
	set rs = cn.Execute(s)
	if Not rs.Eof then
		get_default_valor_texto_bd = Trim("" & rs("valor_default_texto"))
		end if
	if rs.State <> 0 then rs.Close
end function



' _______________________________________________________
' SET DEFAULT VALOR TEXTO BD
'
function set_default_valor_texto_bd(byval usuario, byval nome_chave, byval valor_texto)
dim s, msg_erro
dim rs
	set_default_valor_texto_bd = False
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function
	s = "SELECT * FROM t_DEFAULT WHERE (usuario = '" & usuario & "') AND (nome_chave = '" & nome_chave & "')"
	rs.Open s, cn
	if rs.Eof then
		rs.AddNew
		rs("usuario") = usuario
		rs("nome_chave") = nome_chave
		rs("dt_hr_cadastro") = Now
		rs("dt_hr_ult_atualizacao") = Now
		end if
	
	if Ucase(Trim("" & rs("valor_default_texto"))) <> Ucase(Trim(valor_texto)) then
		rs("valor_default_texto") = valor_texto
		rs("dt_hr_ult_atualizacao") = Now
		end if
	
	rs.Update
	
	if rs.State <> 0 then rs.Close
end function



' _______________________________________________________
' GET DEFAULT VALOR INTEIRO BD
'
function get_default_valor_inteiro_bd(byval usuario, byval nome_chave)
dim s
dim rs
	get_default_valor_inteiro_bd = ""
	s = "SELECT valor_default_inteiro FROM t_DEFAULT WHERE (usuario = '" & usuario & "') AND (nome_chave = '" & nome_chave & "')"
	set rs = cn.Execute(s)
	if Not rs.Eof then
		get_default_valor_inteiro_bd = Trim("" & rs("valor_default_inteiro"))
		end if
	if rs.State <> 0 then rs.Close
end function



' _______________________________________________________
' SET DEFAULT VALOR INTEIRO BD
'
function set_default_valor_inteiro_bd(byval usuario, byval nome_chave, byval valor_inteiro)
dim s, msg_erro
dim rs
	set_default_valor_inteiro_bd = False
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function
	s = "SELECT * FROM t_DEFAULT WHERE (usuario = '" & usuario & "') AND (nome_chave = '" & nome_chave & "')"
	rs.Open s, cn
	if rs.Eof then
		rs.AddNew
		rs("usuario") = usuario
		rs("nome_chave") = nome_chave
		rs("dt_hr_cadastro") = Now
		rs("dt_hr_ult_atualizacao") = Now
		end if
	
	if Ucase(Trim("" & rs("valor_default_inteiro"))) <> Ucase(Trim(valor_inteiro)) then
		rs("valor_default_inteiro") = valor_inteiro
		rs("dt_hr_ult_atualizacao") = Now
		end if
	
	rs.Update
	
	if rs.State <> 0 then rs.Close
end function



' _________________________________________________________________
' OBTEM FLAG BloqueiaCadastramentoQuandoProdutoSemEstoque Pedido
'
function obtem_flag_BloqueiaCadastramentoQuandoProdutoSemEstoque_Pedido(byval numeroLojaPedido)
dim rParam, sLoja
	obtem_flag_BloqueiaCadastramentoQuandoProdutoSemEstoque_Pedido = 0
	set rParam = get_registro_t_parametro(ID_PARAMETRO_BloqueiaCadastramentoQuandoProdutoSemEstoque_Pedido_FlagHabilitacao)
	if Trim("" & rParam.id) <> "" then
		'Se não informou a loja, apenas retorna o flag, não há como verificar se a loja está na lista que deve ignorar a regra
		if (Trim("" & numeroLojaPedido) = "") Or (converte_numero(Trim("" & numeroLojaPedido)) = 0) then
			obtem_flag_BloqueiaCadastramentoQuandoProdutoSemEstoque_Pedido = rParam.campo_inteiro
		else
			'Se a loja não está na lista de lojas que deve ignorar a regra, retorna o flag conforme está configurado no BD
			'Se a loja estiver na lista, a função irá retornar o valor default do flag (zero), ou seja, irá ignorar a regra
			sLoja = "|" & Trim("" & numeroLojaPedido) & "|"
			if Instr(Trim(rParam.campo_texto), sLoja) = 0 then obtem_flag_BloqueiaCadastramentoQuandoProdutoSemEstoque_Pedido = rParam.campo_inteiro
			end if
		end if
	set rParam = Nothing
end function



' _________________________________________________________________
' OBTEM FLAG BloqueiaCadastramentoQuandoProdutoSemEstoque PrePedido
'
function obtem_flag_BloqueiaCadastramentoQuandoProdutoSemEstoque_PrePedido(byval numeroLojaPrePedido)
dim rParam, sLoja
	obtem_flag_BloqueiaCadastramentoQuandoProdutoSemEstoque_PrePedido = 0
	set rParam = get_registro_t_parametro(ID_PARAMETRO_BloqueiaCadastramentoQuandoProdutoSemEstoque_PrePedido_FlagHabilitacao)
	if Trim("" & rParam.id) <> "" then
		'Se não informou a loja, apenas retorna o flag, não há como verificar se a loja está na lista que deve ignorar a regra
		if (Trim("" & numeroLojaPrePedido) = "") Or (converte_numero(Trim("" & numeroLojaPrePedido)) = 0) then
			obtem_flag_BloqueiaCadastramentoQuandoProdutoSemEstoque_PrePedido = rParam.campo_inteiro
		else
			'Se a loja não está na lista de lojas que deve ignorar a regra, retorna o flag conforme está configurado no BD
			'Se a loja estiver na lista, a função irá retornar o valor default do flag (zero), ou seja, irá ignorar a regra
			sLoja = "|" & Trim("" & numeroLojaPrePedido) & "|"
			if Instr(Trim(rParam.campo_texto), sLoja) = 0 then obtem_flag_BloqueiaCadastramentoQuandoProdutoSemEstoque_PrePedido = rParam.campo_inteiro
			end if
		end if
	set rParam = Nothing
end function



' ___________________________________________________________
' OBTEM PARAMETRO PEDIDO ITEM MAX QTDE ITENS
'
function obtem_parametro_PedidoItem_MaxQtdeItens
dim rParam
	obtem_parametro_PedidoItem_MaxQtdeItens = 0
	set rParam = get_registro_t_parametro(ID_PARAMETRO_PedidoItem_MaxQtdeItens)
	if Trim("" & rParam.id) <> "" then obtem_parametro_PedidoItem_MaxQtdeItens = rParam.campo_inteiro
	set rParam = Nothing
end function



' ___________________________________________________________
' OBTEM PARAMETRO TRANSF PRODUTOS ENTRE CDs MAX QTDE ITENS
'
function obtem_parametro_TransfProdutosEntreCDs_MaxQtdeItens
dim rParam
	obtem_parametro_TransfProdutosEntreCDs_MaxQtdeItens = 0
	set rParam = get_registro_t_parametro(ID_PARAMETRO_TransfProdutosEntreCDs_MaxQtdeItens)
	if Trim("" & rParam.id) <> "" then obtem_parametro_TransfProdutosEntreCDs_MaxQtdeItens = rParam.campo_inteiro
	set rParam = Nothing
end function



' ___________________________________________________________
' OBTEM PARAMETRO TRANSF PRODUTOS ENTRE PEDIDOS MAX QTDE ITENS
'
function obtem_parametro_TransfProdutosEntrePedidos_MaxQtdeItens
dim rParam
	obtem_parametro_TransfProdutosEntrePedidos_MaxQtdeItens = 0
	set rParam = get_registro_t_parametro(ID_PARAMETRO_TransfProdutosEntrePedidos_MaxQtdeItens)
	if Trim("" & rParam.id) <> "" then obtem_parametro_TransfProdutosEntrePedidos_MaxQtdeItens = rParam.campo_inteiro
	set rParam = Nothing
end function



' ___________________________________________________________
' OBTEM PARAMETRO SENHA DESCONTO SUPERIOR MAX QTDE ITENS
'
function obtem_parametro_SenhaDescontoSuperior_MaxQtdeItens
dim rParam
	obtem_parametro_SenhaDescontoSuperior_MaxQtdeItens = 0
	set rParam = get_registro_t_parametro(ID_PARAMETRO_SenhaDescontoSuperior_MaxQtdeItens)
	if Trim("" & rParam.id) <> "" then obtem_parametro_SenhaDescontoSuperior_MaxQtdeItens = rParam.campo_inteiro
	set rParam = Nothing
end function



' ___________________________________________________________
' OBTEM PARAMETRO MAX TENTATIVAS LOGIN
'
function obtem_parametro_max_tentativas_login
dim rParam
	obtem_parametro_max_tentativas_login = 0
	set rParam = get_registro_t_parametro(ID_PARAM_MAX_TENTATIVAS_LOGIN)
	if Trim("" & rParam.id) <> "" then obtem_parametro_max_tentativas_login = rParam.campo_inteiro
	set rParam = Nothing
end function



' _______________________________________________________
' OBTEM PercVlPedidoLimiteRA
'
function obtem_PercVlPedidoLimiteRA
dim msg_erro
	obtem_PercVlPedidoLimiteRA = converte_numero(le_parametro_bd(ID_PARAM_PercVlPedidoLimiteRA, msg_erro))
end function



' _______________________________________________________
' OBTEM MAX DIAS DT INICIAL FILTRO PERIODO
'
function obtem_max_dias_dt_inicial_filtro_periodo
dim msg_erro
	obtem_max_dias_dt_inicial_filtro_periodo = converte_numero(le_parametro_bd(ID_PARAM_MAX_DIAS_DT_INICIAL_FILTRO_PERIODO, msg_erro))
end function



' _______________________________________________________
' OBTEM PERC MAX DESCONTO CADASTRADO NA LOJA
'
function obtem_perc_max_desconto_cadastrado_na_loja(byval l)
dim r
	obtem_perc_max_desconto_cadastrado_na_loja = 0
	l = Trim("" & l)
	if l = "" then exit function
	set r = cn.Execute("SELECT PercMaxSenhaDesconto FROM t_LOJA WHERE (loja = '" & l & "')")
	if Not r.Eof then obtem_perc_max_desconto_cadastrado_na_loja = r("PercMaxSenhaDesconto")
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________________________________________
' OBTEM PERC LIMITE RA SEM DESAGIO
'
function obtem_perc_limite_RA_sem_desagio
dim msg_erro
	obtem_perc_limite_RA_sem_desagio = converte_numero(le_parametro_bd(ID_PARAM_PERC_LIMITE_RA_SEM_DESAGIO, msg_erro))
end function



' ___________________________________________________________
' OBTEM MAX PERIODO LINK DANFE DISPONIVEL NO PEDIDO EM DIAS
'
function obtem_max_periodo_link_danfe_disponivel_no_pedido_em_dias
dim rParam
	obtem_max_periodo_link_danfe_disponivel_no_pedido_em_dias = 0
	set rParam = get_registro_t_parametro(ID_PARAM_MAX_PERIODO_LINK_DANFE_DISPONIVEL_NO_PEDIDO_EM_DIAS)
	if Trim("" & rParam.id) <> "" then obtem_max_periodo_link_danfe_disponivel_no_pedido_em_dias = rParam.campo_inteiro
	set rParam = Nothing
end function



' _______________________________________________________
' OBTEM PERC MAX RT
'
function obtem_perc_max_RT
dim msg_erro
	obtem_perc_max_RT = converte_numero(le_parametro_bd(ID_PARAM_PERC_MAX_RT, msg_erro))
end function



' _______________________________________________________
' OBTEM PERC MAX DESC SEM ZERAR RT
'
function obtem_perc_max_desc_sem_zerar_RT(byval l)
dim r
	obtem_perc_max_desc_sem_zerar_RT = 0
	l = Trim("" & l)
	if l = "" then exit function
	set r = cn.Execute("SELECT PercMaxDescSemZerarRT FROM t_LOJA WHERE (loja = '" & l & "')")
	if Not r.Eof then obtem_perc_max_desc_sem_zerar_RT = r("PercMaxDescSemZerarRT")
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________________________________________
' OBTEM PERC MAX COMISSAO E DESCONTO POR LOJA
'
function obtem_perc_max_comissao_e_desconto_por_loja(byval loja)
dim s
dim rx
dim rs
	set rx = New cl_PERC_MAX_COMISSAO_E_DESCONTO_POR_LOJA
	rx.perc_max_comissao = 0
	rx.perc_max_comissao_e_desconto = 0
	rx.perc_max_comissao_e_desconto_pj = 0
	rx.perc_max_comissao_e_desconto_nivel2 = 0
	rx.perc_max_comissao_e_desconto_nivel2_pj = 0
	rx.perc_max_comissao_alcada1 = 0
	rx.perc_max_comissao_e_desconto_alcada1_pf = 0
	rx.perc_max_comissao_e_desconto_alcada1_pj = 0
	rx.perc_max_comissao_alcada2 = 0
	rx.perc_max_comissao_e_desconto_alcada2_pf = 0
	rx.perc_max_comissao_e_desconto_alcada2_pj = 0
	rx.perc_max_comissao_alcada3 = 0
	rx.perc_max_comissao_e_desconto_alcada3_pf = 0
	rx.perc_max_comissao_e_desconto_alcada3_pj = 0
	rx.isCadastrado = False
	
'	LEMBRANDO QUE SE TRATA DE UM PONTEIRO P/ O OBJETO
	set obtem_perc_max_comissao_e_desconto_por_loja = rx
	
	loja = Trim("" & loja)
	s = "SELECT" & _
			" perc_max_comissao," & _
			" perc_max_comissao_e_desconto," & _
			" perc_max_comissao_e_desconto_pj," & _
			" perc_max_comissao_e_desconto_nivel2," & _
			" perc_max_comissao_e_desconto_nivel2_pj," & _
			" perc_max_comissao_alcada1," & _
			" perc_max_comissao_e_desconto_alcada1_pf," & _
			" perc_max_comissao_e_desconto_alcada1_pj," & _
			" perc_max_comissao_alcada2," & _
			" perc_max_comissao_e_desconto_alcada2_pf," & _
			" perc_max_comissao_e_desconto_alcada2_pj," & _
			" perc_max_comissao_alcada3," & _
			" perc_max_comissao_e_desconto_alcada3_pf," & _
			" perc_max_comissao_e_desconto_alcada3_pj" & _
		" FROM t_LOJA" & _
		" WHERE" & _
			" (CONVERT(smallint,loja) = " & loja & ")"
	set rs = cn.Execute(s)
	if Not rs.Eof then
		rx.perc_max_comissao = rs("perc_max_comissao")
		rx.perc_max_comissao_e_desconto = rs("perc_max_comissao_e_desconto")
		rx.perc_max_comissao_e_desconto_pj = rs("perc_max_comissao_e_desconto_pj")
		rx.perc_max_comissao_e_desconto_nivel2 = rs("perc_max_comissao_e_desconto_nivel2")
		rx.perc_max_comissao_e_desconto_nivel2_pj = rs("perc_max_comissao_e_desconto_nivel2_pj")
		rx.perc_max_comissao_alcada1 = rs("perc_max_comissao_alcada1")
		rx.perc_max_comissao_e_desconto_alcada1_pf = rs("perc_max_comissao_e_desconto_alcada1_pf")
		rx.perc_max_comissao_e_desconto_alcada1_pj = rs("perc_max_comissao_e_desconto_alcada1_pj")
		rx.perc_max_comissao_alcada2 = rs("perc_max_comissao_alcada2")
		rx.perc_max_comissao_e_desconto_alcada2_pf = rs("perc_max_comissao_e_desconto_alcada2_pf")
		rx.perc_max_comissao_e_desconto_alcada2_pj = rs("perc_max_comissao_e_desconto_alcada2_pj")
		rx.perc_max_comissao_alcada3 = rs("perc_max_comissao_alcada3")
		rx.perc_max_comissao_e_desconto_alcada3_pf = rs("perc_max_comissao_e_desconto_alcada3_pf")
		rx.perc_max_comissao_e_desconto_alcada3_pj = rs("perc_max_comissao_e_desconto_alcada3_pj")
		rx.isCadastrado = True
		end if
	
	if rs.State <> 0 then rs.Close
	set rs = nothing
end function



' _______________________________________________________
' OBTEM PERC DESAGIO RA DO INDICADOR
'
function obtem_perc_desagio_RA_do_indicador(byval l)
dim r
	obtem_perc_desagio_RA_do_indicador = 0
	l = Trim("" & l)
	if l = "" then exit function
	set r = cn.Execute("SELECT perc_desagio_RA FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido = '" & l & "'")
	if not r.eof then obtem_perc_desagio_RA_do_indicador = r("perc_desagio_RA")
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________________________________________
' OBTEM LIMITE MENSAL COMPRAS DO INDICADOR
'
function obtem_limite_mensal_compras_do_indicador(byval l)
dim r
	obtem_limite_mensal_compras_do_indicador = 0
	l = Trim("" & l)
	if l = "" then exit function
	set r = cn.Execute("SELECT vl_limite_mensal FROM t_ORCAMENTISTA_E_INDICADOR WHERE apelido = '" & l & "'")
	if not r.eof then obtem_limite_mensal_compras_do_indicador = r("vl_limite_mensal")
	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________________________________________
' CALCULA LIMITE MENSAL CONSUMIDO DO INDICADOR
'
function calcula_limite_mensal_consumido_do_indicador(byval strIndicador, byval dtReferencia)
dim strDia, strMes, strAno, strSql
dim r, dtReferenciaAux, dtInferior, dtSuperior
dim vl_consumido, vl_devolucao
	calcula_limite_mensal_consumido_do_indicador = 0
	strIndicador = Trim("" & strIndicador)
	if strIndicador = "" then exit function
	if Not IsDate(dtReferencia) then exit function
'	OBTÉM O 1º DIA DO MÊS DE REFERÊNCIA
	if Not decodifica_data(dtReferencia, strDia, strMes, strAno) then exit function
	dtInferior = StrToDate("01/" & strMes & "/" & strAno)
'	OBTÉM O 1º DIA DO MÊS SUBSEQUENTE AO MÊS DE REFERÊNCIA
	dtReferenciaAux = DateAdd("m", 1, dtReferencia)
	if Not decodifica_data(dtReferenciaAux, strDia, strMes, strAno) then exit function
	dtSuperior = StrToDate("01/" & strMes & "/" & strAno)
'	MONTA SQL (PEDIDOS CADASTRADOS)
	strSql = "SELECT" & _
				" ISNULL(SUM(qtde*preco_venda),0) AS vl_total" & _
			" FROM t_PEDIDO tP INNER JOIN t_PEDIDO_ITEM tPI" & _
				" ON (tP.pedido=tPI.pedido)" & _
			" WHERE" & _
				" (st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
				" AND (indicador = '" & strIndicador & "')" & _
				" AND (data >= " & bd_monta_data(dtInferior) & ")" & _
				" AND (data < " & bd_monta_data(dtSuperior) & ")"
	set r = cn.Execute(strSql)
	if Not r.eof then 
		vl_consumido = r("vl_total")
	else
		vl_consumido = 0
		end if

'	MONTA SQL (ITENS DEVOLVIDOS)
	strSql = "SELECT" & _
				" ISNULL(SUM(qtde*preco_venda),0) AS vl_total_devolucao" & _
			" FROM t_PEDIDO tP INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO tPID" & _
				" ON (tP.pedido=tPID.pedido)" & _
			" WHERE" & _
				" (indicador = '" & strIndicador & "')" & _
				" AND (data >= " & bd_monta_data(dtInferior) & ")" & _
				" AND (data < " & bd_monta_data(dtSuperior) & ")"
	set r = cn.Execute(strSql)
	if Not r.eof then 
		vl_devolucao = r("vl_total_devolucao")
	else
		vl_devolucao = 0
		end if

	vl_consumido = vl_consumido - vl_devolucao
	if vl_consumido < 0 then vl_consumido = 0
	
	calcula_limite_mensal_consumido_do_indicador = vl_consumido

	if r.State <> 0 then r.Close
	set r = nothing
end function



' _______________________________________________________________________________________
' CALCULA VALOR PRESENTE
' O parâmetro "perc_taxa" deve informar a taxa de juros.
'	Ex: se for 2%, passar 0,02
function CalculaValorPresente(byval vl_valor_futuro, byval perc_taxa, byval n_periodos)
dim vl_valor_presente
	'PV = FV / (1+i)^n
	vl_valor_presente = vl_valor_futuro / ((1 + perc_taxa) ^ n_periodos)
	CalculaValorPresente = vl_valor_presente
end function



' ___________________________________________________________________________
' CALCULA TOTAL RA LIQUIDO BD
function calcula_total_RA_liquido_BD(byval id_pedido, byref vl_total_RA_liquido, byref msg_erro)
dim s
dim rs
dim rspb
dim id_pedido_base
dim vl_total_RA
dim percentual_desagio_RA_liquido

	calcula_total_RA_liquido_BD = False
	
	id_pedido = Trim("" & id_pedido)
	vl_total_RA_liquido = 0
	msg_erro = ""
	
	id_pedido_base = retorna_num_pedido_base(id_pedido)

	s = "SELECT " & _
			"*" & _
		" FROM t_PEDIDO" & _
		" WHERE" & _
			" (pedido='" & id_pedido_base & "')"
	set rspb = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rspb.Eof then
		msg_erro = "Pedido-base " & id_pedido_base & " não foi encontrado."
		exit function
		end if

	percentual_desagio_RA_liquido = rspb("perc_desagio_RA_liquida")
	
'	OBTÉM OS VALORES TOTAIS DE NF, RA E VENDA
	vl_total_RA = 0
	s = "SELECT" & _
			" SUM(qtde*(preco_NF-preco_venda)) AS total_RA" & _
		" FROM t_PEDIDO_ITEM" & _
			" INNER JOIN t_PEDIDO" & _
				" ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" & _
		" WHERE" & _
			" (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')" & _
			" AND (t_PEDIDO.pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')"
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if Not rs.Eof then
		if Not IsNull(rs("total_RA")) then vl_total_RA = rs("total_RA")
		end if
	
	vl_total_RA_liquido = CCur(vl_total_RA - (percentual_desagio_RA_liquido/100)*vl_total_RA)
	vl_total_RA_liquido = converte_numero(formata_moeda(vl_total_RA_liquido))
	
	calcula_total_RA_liquido_BD = True
end function



' ___________________________________________
' CALCULA TOTAL RA LIQUIDO
'
function calcula_total_RA_liquido(byval percentual_desagio_RA_liquida, byval vl_total_RA, byref vl_total_RA_liquido)
	calcula_total_RA_liquido = False
	vl_total_RA_liquido = CCur(vl_total_RA - (percentual_desagio_RA_liquida/100)*vl_total_RA)
	vl_total_RA_liquido = converte_numero(formata_moeda(vl_total_RA_liquido))
	calcula_total_RA_liquido = True
end function



' ___________________________________________
' FAMILIA PEDIDOS QTDE PEDIDOS ENTREGUES
'
function familia_pedidos_qtde_pedidos_entregues(byval id_pedido)
dim r, s, id_pedido_base
	familia_pedidos_qtde_pedidos_entregues = 0
	id_pedido=Trim("" & id_pedido)
	id_pedido_base = retorna_num_pedido_base(id_pedido)
	s = "SELECT" & _
			" Coalesce(Count(*), 0) As qtde" & _
		" FROM t_PEDIDO" & _
		" WHERE" & _
			" (pedido LIKE '" & id_pedido_base & BD_CURINGA_TODOS & "')" & _
			" AND (st_entrega = '" & ST_ENTREGA_ENTREGUE & "')"
	set r = cn.Execute(s)
	if Not r.Eof then
		familia_pedidos_qtde_pedidos_entregues = CLng(r("qtde"))
		end if
end function



' ___________________________________________
' LE ORDEM SERVICO
'
function le_ordem_servico(byval id_ordem_servico, byref r_ordem_servico, byref msg_erro)
dim s
dim rs

	le_ordem_servico = False
	msg_erro = ""
	id_ordem_servico=retorna_so_digitos(Trim("" & id_ordem_servico))
	id_ordem_servico=normaliza_codigo(id_ordem_servico, TAM_MAX_NSU)
	set r_ordem_servico = New cl_ORDEM_SERVICO
	s="SELECT * FROM t_ORDEM_SERVICO WHERE (ordem_servico='" & id_ordem_servico & "')"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Ordem de serviço nº " & id_ordem_servico & " não está cadastrado."
	else
		with r_ordem_servico
			.ordem_servico				= Trim("" & rs("ordem_servico"))
			.usuario					= Trim("" & rs("usuario"))				
			.data						= rs("data")
			.hora						= Trim("" & rs("hora"))
			.situacao_status			= Trim("" & rs("situacao_status"))
			.situacao_data				= rs("situacao_data")
			.situacao_usuario			= Trim("" & rs("situacao_usuario"))
			.pedido						= Trim("" & rs("pedido"))
			.fabricante					= Trim("" & rs("fabricante"))
			.produto					= Trim("" & rs("produto"))
			.qtde						= rs("qtde")
			.ean						= Trim("" & rs("ean"))
			.descricao					= Trim("" & rs("descricao"))
			.descricao_html				= Trim("" & rs("descricao_html"))
			.obs_pecas_necessarias		= Trim("" & rs("obs_pecas_necessarias"))
			.nf							= Trim("" & rs("nf"))
			.indicador					= Trim("" & rs("indicador"))
			.id_cliente					= Trim("" & rs("id_cliente"))
			.tipo_cliente				= Trim("" & rs("tipo_cliente"))
			.nome_cliente				= Trim("" & rs("nome_cliente"))
			.endereco					= Trim("" & rs("endereco"))
			.endereco_numero			= Trim("" & rs("endereco_numero"))
			.endereco_complemento		= Trim("" & rs("endereco_complemento"))
			.bairro						= Trim("" & rs("bairro"))
			.cidade						= Trim("" & rs("cidade"))
			.uf							= Trim("" & rs("uf"))
			.cep						= Trim("" & rs("cep"))
			.ddd_res					= Trim("" & rs("ddd_res"))
			.tel_res					= Trim("" & rs("tel_res"))
			.ddd_com					= Trim("" & rs("ddd_com"))
			.tel_com					= Trim("" & rs("tel_com"))
			.ramal_com					= Trim("" & rs("ramal_com"))
			.contato					= Trim("" & rs("contato"))
			.cod_estoque_origem			= Trim("" & rs("cod_estoque_origem"))
			.loja_estoque_origem		= Trim("" & rs("loja_estoque_origem"))
			.cod_estoque_destino		= Trim("" & rs("cod_estoque_destino"))
			.loja_estoque_destino		= Trim("" & rs("loja_estoque_destino"))
			.pedido_destino				= Trim("" & rs("pedido_destino"))
			.id_nfe_emitente			= rs("id_nfe_emitente")
			end with
		end if	

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_ordem_servico=True
end function



' ___________________________________________
' LE ORDEM SERVICO ITEM
'
function le_ordem_servico_item(byval id_ordem_servico, byref v_item, byref msg_erro)
dim s
dim rs
	le_ordem_servico_item = False
	msg_erro = ""
	id_ordem_servico=retorna_so_digitos(Trim("" & id_ordem_servico))
	id_ordem_servico=normaliza_codigo(id_ordem_servico, TAM_MAX_NSU)
	redim v_item(0)
	set v_item(0) = New cl_ORDEM_SERVICO_ITEM
	
	s="SELECT * FROM t_ORDEM_SERVICO_ITEM WHERE (ordem_servico='" & id_ordem_servico & "') ORDER BY sequencia"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Não há itens cadastrados para a ordem de serviço nº " & id_ordem_servico & "."
	else
		do while Not rs.EOF 
			if Trim(v_item(Ubound(v_item)).ordem_servico)<>"" then
				redim preserve v_item(Ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ORDEM_SERVICO_ITEM
				end if
			with v_item(Ubound(v_item))
				.ordem_servico			= Trim("" & rs("ordem_servico"))
				.num_serie				= Trim("" & rs("num_serie"))
				.tipo					= Trim("" & rs("tipo"))
				.descricao_volume		= Trim("" & rs("descricao_volume"))
				.obs_problema			= Trim("" & rs("obs_problema"))
				.sequencia				= rs("sequencia")
				end with
			rs.MoveNext 
			Loop
		end if	

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close		

	if msg_erro = "" then le_ordem_servico_item=True
end function



' ___________________________________________
' X _ O S _ S T A T U S
'
function x_OS_status(byval status)
dim s
	status = Trim("" & status)
	select case status
		case ST_OS_EM_ANDAMENTO: s="Em Andamento"
		case ST_OS_ENCERRADA: s="Encerrada"
		case ST_OS_CANCELADA: s="Cancelada"
		case else s=""
		end select
	x_OS_status=s
end function



' ___________________________________________
' X _ O S _ S T A T U S _ C O R
'
function x_OS_status_cor(byval status)
dim s_cor
	status = Trim("" & status)
	select case status
		case ST_OS_EM_ANDAMENTO: s_cor="blue"
		case ST_OS_ENCERRADA: s_cor="green"
		case ST_OS_CANCELADA: s_cor="red"
		case else s_cor="black"
		end select
	x_OS_status_cor=s_cor
end function



' ___________________________________________
' POSSUI ACESSO LOJA
'
function PossuiAcessoLoja(byval usuario, byval loja)
dim r, strSql
	PossuiAcessoLoja = False
	usuario = Trim("" & usuario)
	loja = Trim("" & loja)
	if usuario = "" then exit function
	if loja = "" then exit function
	strSql = "SELECT" & _
				" loja" & _
			" FROM t_USUARIO_X_LOJA" & _
			" WHERE" & _
				" (usuario = '" & usuario & "')" & _
				" AND (CONVERT(smallint, loja) = " & loja & ")"
	set r = cn.Execute(strSql)
	if Not r.Eof then PossuiAcessoLoja = True
	if r.State <> 0 then r.Close
	set r = nothing
end function



' ___________________________________________
' CONTA NUM LOJAS ACESSO LIBERADO
'
function ContaNumLojasAcessoLiberado(byval strUsuario, byval loja_atual)
dim r, strSql
	ContaNumLojasAcessoLiberado = 0
	strUsuario = Trim("" & strUsuario)
	loja_atual = Trim("" & loja_atual)
	if strUsuario = "" then exit function
	if loja_atual = "" then
		strSql = "SELECT" & _
					" Coalesce(Count(*),0) AS qtde" & _
				" FROM t_USUARIO_X_LOJA" & _
				" WHERE" & _
					" (usuario = '" & strUsuario & "')"
	else 
	'	LEMBRE-SE: O USUÁRIO QUE TEM PERMISSÃO DE ACESSO A TODAS AS LOJAS PODE
	'	ACESSAR UMA LOJA QUE NÃO ESTÁ CADASTRADA EM t_USUARIO_X_LOJA
		strSql = "SELECT" & _
					" Coalesce(Count(*),0) AS qtde" & _
				" FROM " & _
					"(" & _
						"SELECT DISTINCT" & _
							" loja" & _
						" FROM " & _
							"(" & _
								"SELECT" & _
									" loja" & _
								" FROM t_USUARIO_X_LOJA" & _
								" WHERE" & _
									" (usuario = '" & strUsuario & "')" & _
								" UNION " & _
								"SELECT" & _
									" loja" & _
								" FROM t_LOJA" & _
								" WHERE" & _
									" (CONVERT(smallint, loja) = " & loja_atual & ")" & _
							") t__AUX_UNION" & _
					") t__AUX_COUNT"
		end if
	set r = cn.Execute(strSql)
	if Not r.Eof then ContaNumLojasAcessoLiberado = CLng(r("qtde"))
	if r.State <> 0 then r.Close
	set r = nothing
end function



' ___________________________________________
' LE CLIENTE REF BANCARIA
'
function le_cliente_ref_bancaria(byval id_cliente, byref vReg, byref msg_erro)
dim s
dim rs
	le_cliente_ref_bancaria = False
	msg_erro = ""
	id_cliente=Trim("" & id_cliente)
	redim vReg(0)
	set vReg(0) = New cl_CLIENTE_REF_BANCARIA
	
	s="SELECT * FROM t_CLIENTE_REF_BANCARIA WHERE (id_cliente='" & id_cliente & "') ORDER BY ordem"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	do while Not rs.EOF 
		if Trim(vReg(Ubound(vReg)).id_cliente)<>"" then
			redim preserve vReg(Ubound(vReg)+1)
			set vReg(ubound(vReg)) = New cl_CLIENTE_REF_BANCARIA
			end if
		with vReg(Ubound(vReg))
			.id_cliente				= Trim("" & rs("id_cliente"))
			.banco					= Trim("" & rs("banco"))
			.agencia				= Trim("" & rs("agencia"))
			.conta					= Trim("" & rs("conta"))
			.ddd					= Trim("" & rs("ddd"))
			.telefone				= Trim("" & rs("telefone"))
			.contato				= Trim("" & rs("contato"))
			.ordem					= rs("ordem")
			.dt_cadastro			= rs("dt_cadastro")
			.usuario_cadastro		= Trim("" & rs("usuario_cadastro"))
			end with
		rs.MoveNext 
		Loop

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close		

	if msg_erro = "" then le_cliente_ref_bancaria=True
end function



' ___________________________________________
' LE CLIENTE REF COMERCIAL
'
function le_cliente_ref_comercial(byval id_cliente, byref vReg, byref msg_erro)
dim s
dim rs
	le_cliente_ref_comercial = False
	msg_erro = ""
	id_cliente=Trim("" & id_cliente)
	redim vReg(0)
	set vReg(0) = New cl_CLIENTE_REF_COMERCIAL
	
	s="SELECT * FROM t_CLIENTE_REF_COMERCIAL WHERE (id_cliente='" & id_cliente & "') ORDER BY ordem"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	do while Not rs.EOF 
		if Trim(vReg(Ubound(vReg)).id_cliente)<>"" then
			redim preserve vReg(Ubound(vReg)+1)
			set vReg(ubound(vReg)) = New cl_CLIENTE_REF_COMERCIAL
			end if
		with vReg(Ubound(vReg))
			.id_cliente				= Trim("" & rs("id_cliente"))
			.nome_empresa			= Trim("" & rs("nome_empresa"))
			.contato				= Trim("" & rs("contato"))
			.ddd					= Trim("" & rs("ddd"))
			.telefone				= Trim("" & rs("telefone"))
			.ordem					= rs("ordem")
			.dt_cadastro			= rs("dt_cadastro")
			.usuario_cadastro		= Trim("" & rs("usuario_cadastro"))
			end with
		rs.MoveNext 
		Loop

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close		

	if msg_erro = "" then le_cliente_ref_comercial=True
end function



' ___________________________________________
' LE CLIENTE REF PROFISSIONAL
'
function le_cliente_ref_profissional(byval id_cliente, byref vReg, byref msg_erro)
dim s
dim rs
	le_cliente_ref_profissional = False
	msg_erro = ""
	id_cliente=Trim("" & id_cliente)
	redim vReg(0)
	set vReg(0) = New cl_CLIENTE_REF_PROFISSIONAL
	
	s="SELECT * FROM t_CLIENTE_REF_PROFISSIONAL WHERE (id_cliente='" & id_cliente & "') ORDER BY ordem"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	do while Not rs.EOF 
		if Trim(vReg(Ubound(vReg)).id_cliente)<>"" then
			redim preserve vReg(Ubound(vReg)+1)
			set vReg(ubound(vReg)) = New cl_CLIENTE_REF_PROFISSIONAL
			end if
		with vReg(Ubound(vReg))
			.id_cliente				= Trim("" & rs("id_cliente"))
			.nome_empresa			= Trim("" & rs("nome_empresa"))
			.cargo					= Trim("" & rs("cargo"))
			.ddd					= Trim("" & rs("ddd"))
			.telefone				= Trim("" & rs("telefone"))
			.periodo_registro		= rs("periodo_registro")
			.rendimentos			= rs("rendimentos")
			.ordem					= rs("ordem")
			.dt_cadastro			= rs("dt_cadastro")
			.usuario_cadastro		= Trim("" & rs("usuario_cadastro"))
			.cnpj					= Trim("" & rs("cnpj"))
			end with
		rs.MoveNext 
		Loop

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_cliente_ref_profissional=True
end function



' ___________________________________________
' CONSISTE MUNICIPIO IBGE OK
'
function consiste_municipio_IBGE_ok(byval municipio, byval uf, byref lista_sugerida_municipios, byref msg_erro)
dim s, strCodUF
dim strNfeT1ServidorBd, strNfeT1NomeBd, strNfeT1UsuarioBd, strNfeT1SenhaCriptografadaBd
dim chave
dim senha_decodificada
dim tNE, tUF, tMunicipio
dim dbcNFe
	
	consiste_municipio_IBGE_ok = False
	lista_sugerida_municipios = ""
	msg_erro = ""

'	CONSISTE PARÂMETROS
	if Trim("" & municipio) = "" then
		msg_erro = "Não é possível consistir o município através da relação de municípios do IBGE: nenhum município foi informado!!"
		exit function
		end if
		
	if Trim("" & uf) = "" then
		msg_erro = "Não é possível consistir o município através da relação de municípios do IBGE: a UF não foi informada!!"
		exit function
		end if
		
	if Not uf_ok(uf) then
		msg_erro = "Não é possível consistir o município através da relação de municípios do IBGE: a UF é inválida (" & uf & ")!!"
		exit function
		end if
			
'   OBTÉM O EMITENTE PADRÃO
    s = "SELECT" & _
            " NFe_T1_servidor_BD," & _
            " NFe_T1_nome_BD," & _
            " NFe_T1_usuario_BD," & _
            " NFe_T1_senha_BD" & _
        " FROM t_NFe_EMITENTE" & _
        " WHERE" & _
            " (NFe_st_emitente_padrao = 1)"
	set tNE = cn.Execute(s)
	if tNE.Eof then 
		msg_erro = "Não há um emitente de NFe padrão definido no sistema!!"
		exit function
		end if

	strNfeT1ServidorBd = Trim("" & tNE("NFe_T1_servidor_BD"))
	strNfeT1NomeBd = Trim("" & tNE("NFe_T1_nome_BD"))
	strNfeT1UsuarioBd = Trim("" & tNE("NFe_T1_usuario_BD"))
	strNfeT1SenhaCriptografadaBd = Trim("" & tNE("NFe_T1_senha_BD"))
	
	tNE.Close
	set tNE = nothing
	
	chave = gera_chave(FATOR_BD)
	decodifica_dado strNfeT1SenhaCriptografadaBd, senha_decodificada, chave
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

	s = "SELECT " & _
			"*" & _
		" FROM NFE_UF" & _
		" WHERE" & _
			" (SiglaUF = '" & Ucase(uf) & "')"
	set tUF = dbcNFe.Execute(s)
	if tUF.Eof then
		msg_erro = "Não é possível consistir o município através da relação de municípios do IBGE: a UF '" & uf & "' não foi localizada na relação do IBGE!!"
		dbcNFe.Close
		set dbcNFe = nothing
		exit function
		end if

	strCodUF = Trim("" & tUF("CodUF"))
	
	tUF.Close
	set tUF = nothing
	
	s = "SELECT " & _
			"*" & _
		" FROM NFE_MUNICIPIO" & _
		" WHERE" & _
			" (CodMunic LIKE '" & strCodUF & BD_CURINGA_TODOS & "')" & _
			" AND (Descricao = '" & QuotedStr(municipio) & "' COLLATE Latin1_General_CI_AI)"
	set tMunicipio = dbcNFe.Execute(s)
	if Not tMunicipio.Eof then
	'	ACHOU O MUNICÍPIO NA LISTA!!
		consiste_municipio_IBGE_ok = True
		dbcNFe.Close
		set dbcNFe = nothing
		exit function
		end if
	
'	NÃO ENCONTROU O MUNICÍPIO, ENTÃO MONTA UMA LISTA DE SUGESTÕES C/ OS POSSÍVEIS MUNICÍPIOS
'	SERÃO DADOS COMO SUGESTÃO TODOS OS MUNICÍPIOS DA UF QUE SE INICIEM C/ A MESMA LETRA DO MUNICÍPIO INFORMADO
	s = "SELECT " & _
			"*" & _
		" FROM NFE_MUNICIPIO" & _
		" WHERE" & _
			" (CodMunic LIKE '" & strCodUF & BD_CURINGA_TODOS & "')" & _
			" AND (Descricao LIKE '" & Left(municipio,1) & BD_CURINGA_TODOS & "' COLLATE Latin1_General_CI_AI)" & _
		" ORDER BY" & _
			" Descricao"
	set tMunicipio = dbcNFe.Execute(s)
	do while Not tMunicipio.Eof
		if lista_sugerida_municipios <> "" then lista_sugerida_municipios = lista_sugerida_municipios & chr(13)
		lista_sugerida_municipios = lista_sugerida_municipios & Trim("" & tMunicipio("Descricao"))
		tMunicipio.MoveNext
		loop
	
	dbcNFe.Close
	set dbcNFe = nothing
end function



' ___________________________________________
' OBTEM DESCRICAO TABELA T_CODIGO_DESCRICAO
'
function obtem_descricao_tabela_t_codigo_descricao(byval grupo, byval codigo)
dim r
dim s, s_resp
	s_resp=""
	s = "SELECT descricao FROM t_CODIGO_DESCRICAO WHERE (grupo='" & grupo & "') AND (codigo='" & codigo & "')"
	set r = cn.Execute(s)
	if Not r.Eof then
		 s_resp = Trim("" & r("descricao"))
	else
		s_resp = "Código não cadastrado (" & codigo & ")"
		end if
	obtem_descricao_tabela_t_codigo_descricao=s_resp
	if r.State <> 0 then r.Close
	set r =  nothing
end function



' ___________________________________________
' OBTEM WMS DEPOSITO ZONA CODIGOS
'
function obtem_wms_deposito_zona_codigos
dim rs
dim s, strResp

	obtem_wms_deposito_zona_codigos = ""
	strResp = ""

	s = "SELECT" & _
			" zona_codigo" & _
		" FROM t_WMS_DEPOSITO_MAP_ZONA" & _
		" WHERE" & _
			" (st_ativo <> 0)" & _
		" ORDER BY" & _
			" zona_codigo"
	set rs = cn.Execute(s)
	do while Not rs.Eof
		if strResp <> "" then strResp = strResp & "|"
		strResp = strResp & UCase(Trim("" & rs("zona_codigo")))
		rs.MoveNext
		loop
	
	if strResp <> "" then strResp = "|" & strResp & "|"
	obtem_wms_deposito_zona_codigos = strResp
	
	if rs.State <> 0 then rs.Close
	set rs =  nothing
end function



' ___________________________________________
' WMS DEPOSITO ZONA OBTEM DESCRICAO
'
function wms_deposito_zona_obtem_descricao(byval zona_id)
dim r
dim s, s_resp
	s_resp=""
	s = "SELECT zona_codigo FROM t_WMS_DEPOSITO_MAP_ZONA WHERE (id = " & zona_id & ")"
	set r = cn.Execute(s)
	if Not r.Eof then
		 s_resp = Trim("" & r("zona_codigo"))
	else
		s_resp = "Código não cadastrado (" & zona_id & ")"
		end if
	wms_deposito_zona_obtem_descricao=s_resp
	if r.State <> 0 then r.Close
	set r =  nothing
end function



' _________________
' gera_uid
'
function gera_uid
dim r
	set r = cn.Execute("SELECT Convert(varchar(36), NEWID()) AS uid")
	if Not r.Eof then gera_uid = Trim("" & r("uid"))
	if r.State <> 0 then r.Close
	set r = nothing
end function



' ___________________________________________
' LE EC PRODUTO COMPOSTO ITEM
'
function le_EC_produto_composto_item(byval fabricante, byval produto, byref v_item, byref msg_erro)
dim s
dim rs
	le_EC_produto_composto_item = False
	msg_erro = ""
	fabricante=Trim("" & fabricante)
	produto=Trim("" & produto)
	redim v_item(0)
	set v_item(0) = New cl_EC_ITEM_PRODUTO_COMPOSTO
	
	s="SELECT * FROM t_EC_PRODUTO_COMPOSTO_ITEM WHERE (fabricante_composto='" & fabricante & "') AND (produto_composto='" & produto & "') ORDER BY sequencia"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Não há itens cadastrados para o produto " & produto & "."
	else
		do while Not rs.EOF 
			if Trim(v_item(Ubound(v_item)).produto_composto)<>"" then
				redim preserve v_item(Ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_EC_ITEM_PRODUTO_COMPOSTO
				end if
			with v_item(Ubound(v_item))
				.fabricante_composto				= Trim("" & rs("fabricante_composto"))
				.produto_composto					= Trim("" & rs("produto_composto"))
				.fabricante_item					= Trim("" & rs("fabricante_item"))
				.produto_item						= Trim("" & rs("produto_item"))
				.qtde								= Trim("" & rs("qtde"))
				.sequencia							= Trim("" & rs("sequencia"))
				.dt_cadastro						= Trim("" & rs("dt_cadastro"))
				.usuario_cadastro					= Trim("" & rs("usuario_cadastro"))
				.dt_ult_atualizacao					= Trim("" & rs("dt_ult_atualizacao"))
				.usuario_ult_atualizacao			= Trim("" & rs("usuario_ult_atualizacao"))
				end with
			rs.MoveNext
			Loop
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_EC_produto_composto_item=True
end function



' ___________________________________________
' boletoArqRetornoObtemUltimaDtCredito
' Status de t_FIN_BOLETO_ARQ_RETORNO.st_processamento
'		EM_PROCESSAMENTO = 1
'		SUCESSO = 2
'		FALHA = 3
function boletoArqRetornoObtemUltimaDtCredito(byref dtCreditoArqRetorno)
dim s
dim rs
	boletoArqRetornoObtemUltimaDtCredito = False
	
	s = "SELECT TOP 1" & _
			" dt_credito," & _
			" nome_arq_retorno," & _
			" dt_hr_processamento," & _
			" usuario_processamento" & _
		" FROM t_FIN_BOLETO_ARQ_RETORNO" & _
		" WHERE" & _
			" (st_processamento = 2)" & _
			" AND (dt_credito IS NOT NULL)" & _
		" ORDER BY" & _
			" dt_credito DESC"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	if Not rs.EOF then
		dtCreditoArqRetorno = rs("dt_credito")
		boletoArqRetornoObtemUltimaDtCredito = True
		end if
	
	if rs.State <> 0 then rs.Close
	set rs =  nothing
end function



' ___________________________________________
' boletoArqRetornoObtemUltimaDtGravacaoArquivo
' Status de t_FIN_BOLETO_ARQ_RETORNO.st_processamento
'		EM_PROCESSAMENTO = 1
'		SUCESSO = 2
'		FALHA = 3
function boletoArqRetornoObtemUltimaDtGravacaoArquivo(byref dtGravacaoArquivoArqRetorno)
dim s
dim rs
	boletoArqRetornoObtemUltimaDtGravacaoArquivo = False
	
	s = "SELECT TOP 1" & _
			" dt_gravacao_arquivo," & _
			" nome_arq_retorno," & _
			" dt_hr_processamento," & _
			" usuario_processamento" & _
		" FROM t_FIN_BOLETO_ARQ_RETORNO" & _
		" WHERE" & _
			" (st_processamento = 2)" & _
			" AND (dt_gravacao_arquivo IS NOT NULL)" & _
		" ORDER BY" & _
			" dt_gravacao_arquivo DESC"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	if Not rs.EOF then
		dtGravacaoArquivoArqRetorno = rs("dt_gravacao_arquivo")
		boletoArqRetornoObtemUltimaDtGravacaoArquivo = True
		end if
	
	if rs.State <> 0 then rs.Close
	set rs =  nothing
end function



' ___________________________________________
' obtemDataReferenciaLimitePagamentoEmAtraso
'
function obtemDataReferenciaLimitePagamentoEmAtraso
dim dtGravacaoArquivoUltArqRetorno
	if Not boletoArqRetornoObtemUltimaDtGravacaoArquivo(dtGravacaoArquivoUltArqRetorno) then
		dtGravacaoArquivoUltArqRetorno = DateAdd("d", -1, Date)
		end if
	
	if Not IsDate(dtGravacaoArquivoUltArqRetorno) then dtGravacaoArquivoUltArqRetorno = DateAdd("d", -1, Date)
	obtemDataReferenciaLimitePagamentoEmAtraso = dtGravacaoArquivoUltArqRetorno
end function



' ___________________________________________
' email_AF_ok
' Consiste se o email é válido para ser usado
' na consulta antifraude nos pagamentos por cartão
function email_AF_ok(ByVal email, ByVal cnpj_cpf, ByRef msg_erro)
dim strDominio
dim s_sql, s_erro
dim tAux
	email_AF_ok = False
	email = Trim("" & email)
	cnpj_cpf = retorna_so_digitos(Trim("" & cnpj_cpf))
	msg_erro = ""
	
	strDominio = "bonshop.com.br"
	if InStr(email, "@" & strDominio) <> 0 then
		msg_erro = "Endereço do domínio " & strDominio & " não é permitido!"
		exit function
		end if
	
	strDominio = "shopvendas.com.br"
	if InStr(email, "@" & strDominio) <> 0 then
		msg_erro = "Endereço do domínio " & strDominio & " não é permitido!"
		exit function
		end if
	
	strDominio = "discomercio.com.br"
	if InStr(email, "@" & strDominio) <> 0 then
		msg_erro = "Endereço do domínio " & strDominio & " não é permitido!"
		exit function
		end if
	
	if Not cria_recordset_otimista(tAux, s_erro) then exit function

	s_sql = "SELECT " & _
				"apelido" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR" & _
			" WHERE " & _
				"(cnpj_cpf <> '" & cnpj_cpf & "')" & _
				" AND " & _
				"(" & _
					"(email = '" & email & "')" & _
					" OR " & _
					"(email2 = '" & email & "')" & _
					" OR " & _
					"(email3 = '" & email & "')" & _
				")"
	if tAux.State <> 0 then tAux.Close
	tAux.Open s_sql, cn
	if Not tAux.Eof then
		msg_erro = "O endereço de email já está sendo usado por um usuário do sistema!"
		exit function
		end if
	
'	FILIAIS DE UMA EMPRESA PODEM USAR O MESMO E-MAIL
	s_sql = "SELECT " & _
				"id" & _
			" FROM t_CLIENTE" & _
			" WHERE " & _
				"(" & _
					"(" & _
						"(tipo = '" & ID_PF & "')" & _
						" AND " & _
						"(cnpj_cpf <> '" & cnpj_cpf & "')" & _
					")" & _
					" OR " & _
					"(" & _
						"(tipo = '" & ID_PJ & "')" & _
						" AND " & _
						"(SUBSTRING(cnpj_cpf,1,8) <> '" & Left(cnpj_cpf,8) & "')" & _
					")" & _
				")" & _
				" AND " & _
				"(" & _
					"(email = '" & email & "')" & _
					" OR " & _
					"(email_xml = '" & email & "')" & _
				")"
	if tAux.State <> 0 then tAux.Close
	tAux.Open s_sql, cn
	if Not tAux.Eof then
		msg_erro = "O endereço de email já está sendo usado por um outro cliente!"
		exit function
		end if
	
	if tAux.State <> 0 then tAux.Close
	
	email_AF_ok = True
end function


'____________________________________________
'verifica_telefones_repetidos
'
function verifica_telefones_repetidos(ByVal ddd, ByVal tel, ByVal cnpj_cpf)
dim LISTA_BRANCA
dim rs, s, s_sql, ttl, s_erro
	
	verifica_telefones_repetidos = 0

	LISTA_BRANCA = "|(11)32683471|"
	s = "|("& Trim("" & ddd) & ")" & Trim("" & tel) & "|"
	if Instr(LISTA_BRANCA, s) <> 0 then exit function

    s_sql = "SELECT id" & _
	" FROM t_CLIENTE" & _
	" WHERE (" & _
			"(" & _
				"ddd_res IN ('" & ddd & "', '0" & ddd & "')" & _
				" AND (tel_res = '" & tel & "')" & _
			")" & _
			" OR (" & _
				"ddd_com IN ('" & ddd & "', '0" & ddd & "')" & _
                " AND (tel_com = '" & tel & "')" & _
			")" & _
			" OR (" & _
				"ddd_cel IN ('" & ddd & "', '0" & ddd & "')" & _
				" AND (tel_cel = '" & tel & "')" & _
			")" & _
			" OR (" & _
				"ddd_com_2 IN ('" & ddd & "', '0" & ddd & "')" & _
				"AND (tel_com_2 = '" & tel & "')" & _
			")" & _
		")" & _
		" AND " & _
		"("

	if Len(Trim("" & cnpj_cpf)) = 11 then
		s_sql = s_sql & _
				" (cnpj_cpf <> '" & cnpj_cpf & "')"
	else
		s_sql = s_sql & _
				" (LEN(cnpj_cpf) = 14)" & _
				" AND (LEFT(cnpj_cpf,8) <> '" & Left(Trim("" & cnpj_cpf),8) & "')"
		end if

	s_sql = s_sql & ")"

	s_sql = s_sql & _
	" UNION ALL" & _
	" SELECT apelido" & _
	" FROM t_ORCAMENTISTA_E_INDICADOR" & _
	" WHERE (" & _
		"(" & _        
			"(" & _
				"ddd IN ('" & ddd & "', '0" & ddd & "')" & _
				" AND (telefone = '" & tel & "')" & _
			")" & _
			" OR (" & _
				"ddd_cel IN ('" & ddd & "', '0" & ddd & "')" & _
                " AND (tel_cel = '" & tel & "')" & _
			")" & _
		")" & _
		" AND " & _
		"("

	if Len(Trim("" & cnpj_cpf)) = 11 then
		s_sql = s_sql & _
				" (cnpj_cpf <> '" & cnpj_cpf & "')"
	else
		s_sql = s_sql & _
				" (LEN(cnpj_cpf) = 14)" & _
				" AND (LEFT(cnpj_cpf,8) <> '" & Left(Trim("" & cnpj_cpf),8) & "')"
		end if

	s_sql = s_sql & _
			")" & _
		")"

    s_sql = "SELECT COUNT(*) AS qtde_telefone FROM (" & s_sql & ") tCAD"

    if Not cria_recordset_otimista(rs, s_erro) then exit function

    if rs.State <> 0 then rs.Close
	rs.Open s_sql, cn
    ttl=rs("qtde_telefone")
    if rs.State <> 0 then rs.Close

    verifica_telefones_repetidos = ttl

end function


' ____________________________________________________________________
' obtem_lista_usuario_x_nfe_emitente
' Obtém a lista de CD's liberados para o usuário especificado.
' Cada empresa emitente de NFe que esteja com os campos 'st_ativo'
' e 'st_habilitado_ctrl_estoque' habilitados (valor 1) deve ser
' considerado como um CD (Centro de Distribuição).
' Toda mercadoria que entra no estoque de venda pertence a uma
' determinada empresa cadastrada na tabela t_NFe_EMITENTE e a
' nota fiscal de saída deve ser emitida pela mesma.
function obtem_lista_usuario_x_nfe_emitente(ByVal usuario)
dim vResp, r, s_sql
	usuario = Trim("" & usuario)
	redim vResp(0)
	vResp(Ubound(vResp)) = Null
	s_sql = "SELECT " & _
				"id_nfe_emitente" & _
			" FROM t_USUARIO_X_NFe_EMITENTE tUXNE" & _
				" INNER JOIN t_NFe_EMITENTE tNE ON (tUXNE.id_nfe_emitente = tNE.id)" & _
			" WHERE" & _
				" (usuario = '" & usuario & "')" & _
				" AND (tUXNE.excluido_status = 0)" & _
				" AND (tNE.st_ativo = 1)" & _
				" AND (tNE.st_habilitado_ctrl_estoque = 1)" & _
			" ORDER BY" & _
				" tNE.ordem"
	set r = cn.Execute(s_sql)
	do while Not r.Eof
		if Not Isnull(vResp(Ubound(vResp))) then
			redim preserve vResp(Ubound(vResp)+1)
			end if
		vResp(Ubound(vResp)) = r("id_nfe_emitente")
		r.MoveNext
		loop
	obtem_lista_usuario_x_nfe_emitente = vResp
	if r.State <> 0 then r.Close
	set r =  nothing
end function



' ____________________________________________________________________________________
' calcula_vl_pagto_em_cartao
' Calcula o valor previsto na forma de pagamento usando cartão de crédito
function calcula_vl_pagto_em_cartao(Byval pedido, Byref msg_erro)
dim rPed, m_vl_cartao, blnCartaoPagtoIntegral
	calcula_vl_pagto_em_cartao = 0
	msg_erro = ""
	if Not le_pedido(pedido, rPed, msg_erro) then exit function

	m_vl_cartao = 0
	blnCartaoPagtoIntegral = false

	if CStr(rPed.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_A_VISTA) then
		if Cstr(rPed.av_forma_pagto) = ID_FORMA_PAGTO_CARTAO then
			m_vl_cartao = rPed.vl_total_NF
			blnCartaoPagtoIntegral = true
			end if
	elseif CStr(rPed.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_CARTAO) then
		m_vl_cartao = rPed.pc_qtde_parcelas * rPed.pc_valor_parcela
		blnCartaoPagtoIntegral = true
	elseif CStr(rPed.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) then
		'NOP
	elseif CStr(rPed.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) then
		if Cstr(rPed.pce_forma_pagto_entrada) = ID_FORMA_PAGTO_CARTAO then
			m_vl_cartao = rPed.pce_entrada_valor
			end if
		if Cstr(rPed.pce_forma_pagto_prestacao) = ID_FORMA_PAGTO_CARTAO then
			m_vl_cartao = m_vl_cartao + (rPed.pce_prestacao_qtde * rPed.pce_prestacao_valor)
			end if
		if (Cstr(rPed.pce_forma_pagto_entrada) = ID_FORMA_PAGTO_CARTAO) And _
			(Cstr(rPed.pce_forma_pagto_prestacao) = ID_FORMA_PAGTO_CARTAO) then
			blnCartaoPagtoIntegral = true
			end if
	elseif CStr(rPed.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) then
		if Cstr(rPed.pse_forma_pagto_prim_prest) = ID_FORMA_PAGTO_CARTAO then
			m_vl_cartao = rPed.pse_prim_prest_valor
			end if
		if Cstr(rPed.pse_forma_pagto_demais_prest) = ID_FORMA_PAGTO_CARTAO then
			m_vl_cartao = m_vl_cartao + (rPed.pse_demais_prest_qtde * rPed.pse_demais_prest_valor)
			end if
				
		if (Cstr(rPed.pse_forma_pagto_prim_prest) = ID_FORMA_PAGTO_CARTAO) And _
			(Cstr(rPed.pse_forma_pagto_demais_prest) = ID_FORMA_PAGTO_CARTAO) then
			blnCartaoPagtoIntegral = true
			end if
	elseif CStr(rPed.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELA_UNICA) then
		if Cstr(rPed.pu_forma_pagto) = ID_FORMA_PAGTO_CARTAO then
			m_vl_cartao = rPed.pu_valor
			blnCartaoPagtoIntegral = true
			end if
		end if
	
	if blnCartaoPagtoIntegral And (m_vl_cartao = 0) then m_vl_cartao = rPed.vl_total_NF

	calcula_vl_pagto_em_cartao = m_vl_cartao
end function


function obtem_empresa_NFe_emitente_padrao(byref id_nfe_emitente, byref msg_erro)
dim s
dim tNE
	obtem_empresa_NFe_emitente_padrao = False
	id_nfe_emitente = 0
	msg_erro = ""

	s = "SELECT" & _
			" id" & _
		" FROM t_NFe_EMITENTE" & _
		" WHERE" & _
			" (NFe_st_emitente_padrao = 1)"
	set tNE = cn.Execute(s)
	if tNE.Eof then 
		msg_erro = "Não há um emitente de NFe padrão definido no sistema!!"
		exit function
		end if
	id_nfe_emitente = tNE("id")
	obtem_empresa_NFe_emitente_padrao = True
	tNE.Close
	set tNE = nothing
end function


' _______________________________________________________
' LE NFE EMITENTE
'
function le_nfe_emitente(ByVal id)
dim r, rx, ry, s
	set rx = New cl_NFE_EMITENTE
	limpa_cl_NFE_EMITENTE rx
	
'	LEMBRANDO QUE SE TRATA DE UM PONTEIRO P/ O OBJETO
	set le_nfe_emitente = rx
	
	s = "SELECT " & _
			"e.*, " & _
            "n.NFe_Numero_NF, " & _
            "n.NFe_Serie_NF " & _
		" FROM t_NFe_EMITENTE e" & _
		" INNER JOIN t_NFe_EMITENTE_NUMERACAO n ON e.cnpj = n.cnpj" & _
		" WHERE" & _
			" (e.id = " & Cstr(id) & ")"
	set r = cn.Execute(s)
	if Not r.Eof then
		rx.id = r("id")
		rx.id_boleto_cedente = r("id_boleto_cedente")
		rx.braspag_id_boleto_cedente = r("braspag_id_boleto_cedente")
		rx.st_ativo = r("st_ativo")
		rx.apelido = Trim("" & r("apelido"))
		rx.cnpj = Trim("" & r("cnpj"))
		rx.razao_social = Trim("" & r("razao_social"))
		rx.endereco = Trim("" & r("endereco"))
		rx.endereco_numero = Trim("" & r("endereco_numero"))
		rx.endereco_complemento = Trim("" & r("endereco_complemento"))
		rx.bairro = Trim("" & r("bairro"))
		rx.cidade = Trim("" & r("cidade"))
		rx.uf = Trim("" & r("uf"))
		rx.cep = Trim("" & r("cep"))
		rx.NFe_st_emitente_padrao = r("NFe_st_emitente_padrao")
		rx.NFe_serie_NF = r("NFe_serie_NF")
		rx.NFe_numero_NF = r("NFe_numero_NF")
		rx.NFe_T1_servidor_BD = Trim("" & r("NFe_T1_servidor_BD"))
		rx.NFe_T1_nome_BD = Trim("" & r("NFe_T1_nome_BD"))
		rx.NFe_T1_usuario_BD = Trim("" & r("NFe_T1_usuario_BD"))
		rx.NFe_T1_senha_BD = Trim("" & r("NFe_T1_senha_BD"))
		rx.dt_cadastro = r("dt_cadastro")
		rx.dt_hr_cadastro = r("dt_hr_cadastro")
		rx.usuario_cadastro = Trim("" & r("usuario_cadastro"))
		rx.dt_ult_atualizacao = r("dt_ult_atualizacao")
		rx.dt_hr_ult_atualizacao = r("dt_hr_ult_atualizacao")
		rx.usuario_ult_atualizacao = Trim("" & r("usuario_ult_atualizacao"))
		rx.st_habilitado_ctrl_estoque = r("st_habilitado_ctrl_estoque")
		rx.ordem = r("ordem")
		rx.texto_fixo_especifico = Trim("" & r("texto_fixo_especifico"))
		end if
	
	if r.State <> 0 then r.Close
	set r = nothing

	if rx.id <> 0 then
		s = "SELECT * FROM t_NFe_EMITENTE_CFG_DANFE WHERE (id_nfe_emitente = " & rx.id & ") ORDER BY ordenacao"
		set r = cn.Execute(s)
		do while Not r.Eof
			set ry = new cl_NFE_EMITENTE_CFG_DANFE
			ry.id = r("id")
			ry.id_nfe_emitente = r("id_nfe_emitente")
			ry.min_tamanho_serie_NFe = r("min_tamanho_serie_NFe")
			ry.min_tamanho_numero_NFe = r("min_tamanho_numero_NFe")
			ry.convencao_nome_arq_pdf_danfe = Trim("" & r("convencao_nome_arq_pdf_danfe"))
			ry.diretorio_pdf_danfe = Trim("" & r("diretorio_pdf_danfe"))
			ry.convencao_nome_arq_xml_nfe = Trim("" & r("convencao_nome_arq_xml_nfe"))
			ry.diretorio_xml_nfe = Trim("" & r("diretorio_xml_nfe"))
			ry.dt_hr_cadastro = r("dt_hr_cadastro")
			ry.ordenacao = r("ordenacao")
			rx.AddItemCfgDanfe(ry)
			r.MoveNext
			loop
		if r.State <> 0 then r.Close
		set r = nothing
		end if

end function


' _______________________________________________________
' obtem_apelido_empresa_NFe_emitente
'
function obtem_apelido_empresa_NFe_emitente(byVal id_nfe_emitente)
dim s, empresa
dim tNE
	obtem_apelido_empresa_NFe_emitente = ""

    if id_nfe_emitente = "-1" then
        obtem_apelido_empresa_NFe_emitente = "Cliente"
        exit function
    end if

	if converte_numero(id_nfe_emitente) = 0 then exit function

	s = "SELECT" & _
			" apelido" & _
		" FROM t_NFe_EMITENTE" & _
		" WHERE" & _
			" (id = " & id_nfe_emitente & ")"
	set tNE = cn.Execute(s)
    if Not tNE.Eof then empresa = tNE("apelido")
	obtem_apelido_empresa_NFe_emitente = empresa
	tNE.Close
	set tNE = nothing
end function


' _______________________________________________________
' monta_campo_log
'
function monta_campo_log(byval campo)
dim s_resp
	monta_campo_log = ""
	if IsNull(campo) then
		s_resp = "(null)"
	elseif Trim("" & campo) = "" then
		s_resp = "(vazio)"
	else
		s_resp = campo
		end if
	monta_campo_log = s_resp
end function


' _______________________________________________________
' gravaFinSvcLog
'
function gravaFinSvcLog(byval operacao, _
						byval tabela, _
						byval descricao, _
						byval complemento_1, _
						byval complemento_2, _
						byval complemento_3, _
						byval complemento_4, _
						byval complemento_5, _
						byval complemento_6, _
						byref msg_erro_operacao)
dim id_finsvc_log
dim msg_erro, msg_erro_aux
dim s_sql
dim tFL

	gravaFinSvcLog = False
	msg_erro_operacao = ""
	id_finsvc_log = 0

	if Not fin_gera_nsu(T_FINSVC_LOG, id_finsvc_log, msg_erro_aux) then
		msg_erro_operacao = "Falha ao tentar gerar o NSU para o registro da tabela " & T_FINSVC_LOG
		exit function
		end if

	if Not cria_recordset_otimista(tFL, msg_erro) then
		msg_erro_operacao = msg_erro
		exit function
		end if

	s_sql = "SELECT " & _
				"*" & _
			" FROM t_FINSVC_LOG" & _
			" WHERE" & _
				" (id = -1)"
	tFL.Open s_sql, cn
	if Err <> 0 then
		msg_erro_operacao = Err.Description
		exit function
		end if

	tFL.AddNew
	tFL("id") = id_finsvc_log
	tFL("operacao") = operacao
	tFL("tabela") = tabela
	tFL("descricao") = descricao
	tFL("complemento_1") = complemento_1
	tFL("complemento_2") = complemento_2
	tFL("complemento_3") = complemento_3
	tFL("complemento_4") = complemento_4
	tFL("complemento_5") = complemento_5
	tFL("complemento_6") = complemento_6
	tFL.Update

	if Err <> 0 then
		msg_erro_operacao = Err.Description
		exit function
		end if

	if tFL.State <> 0 then tFL.Close
	set tFL = nothing

	gravaFinSvcLog = True
end function


' ____________________________________________________________________________________________________________
' EmailSndSvcGravaMensagemParaEnvio
' Grava as informações de uma mensagem de email a ser enviada.
' Parâmetros:
'		email_remetente_sistema = Endereço de e-mail do remetente que enviará a mensagem (refere-se à conta de email usada pelo serviço de envio de mensagens do sistema).
'		remetente_from = Endereço de email do usuário do sistema responsável pela geração da mensagem (informação opcional, mas se informada, será usada no campo ReplyTo da mensagem caso o parâmetro 'FLAG_PEDIDO_CHAMADO_EMAIL_USAR_REPLY_TO' esteja ativado).
'		destinatario_To = Um ou mais e-mails de destinatários da mensagem. Os e-mails podem ser separados por espaço em branco, vírgula ou ponto e vírgula.
'		destinatario_Cc = Um ou mais e-mails que receberão cópia da mensagem. Os e-mails podem ser separados por espaço em branco, vírgula ou ponto e vírgula.
'		destinatario_Cco = Um ou mais e-mails que receberão cópia oculta da mensagem. Os e-mails podem ser separados por espaço em branco, vírgula ou ponto e vírgula.
'		assunto = Texto que aparecerá no Subject da mensagem.
'		corpo_mensagem = Texto com o conteúdo da mensagem.
'		dt_hr_agendamento_envio = Data e horário quando a mensagem será enviada.
'		id_mensagem = Retorna o id da mensagem na tabela, caso a gravação tenha ocorrido.
'		msg_erro_grava_msg = Retorna uma mensagem de erro, caso a gravação não tenha ocorrido.
'		Retorno da função = true: a gravação foi realizada; false: a gravação não foi realizada
function EmailSndSvcGravaMensagemParaEnvio(byval email_remetente_sistema, _
											byval remetente_from, _
											byval destinatario_To, _
											byval destinatario_Cc, _
											byval destinatario_Cco, _
											byval assunto, _
											byval corpo_mensagem, _
											byval dt_hr_agendamento_envio, _
											byref id_mensagem, _
											byref msg_erro_grava_msg)
const NOME_DESTA_ROTINA = "ASP-EmailSndSvcGravaMensagemParaEnvio()"
const ESS_ENVIO_NAO_HABILITADO = 0
const ESS_ENVIO_HABILITADO = 1
dim s_sql, msg_erro
dim svclog_descricao, svclog_complemento_1, svclog_complemento_2, svclog_complemento_3, svclog_horario_agendamento
dim tER, tEM
dim id_remetente
dim existeEmailDeDestino
dim v, i
dim email_aux

	' Inicialização
	EmailSndSvcGravaMensagemParaEnvio = False
	id_mensagem = 0
	msg_erro_grava_msg = ""
	
	email_remetente_sistema = Trim("" & email_remetente_sistema)
	remetente_from = Trim("" & remetente_from)
	destinatario_To = Trim("" & destinatario_To)
	destinatario_Cc = Trim("" & destinatario_Cc)
	destinatario_Cco = Trim("" & destinatario_Cco)
	assunto = Trim("" & assunto)
	corpo_mensagem = Trim("" & corpo_mensagem)

	id_remetente = 0
	existeEmailDeDestino = False

	' Consistências
	if Len(email_remetente_sistema) = 0 then
		msg_erro_grava_msg = "E-mail do remetente não preenchido"
		exit function
		end if

	if Not isEmailOk(email_remetente_sistema) then
		msg_erro_grava_msg = "E-mail do remetente preenchido com caracteres inválidos"
		exit function
		end if

	s_sql = "SELECT " & _
				"*" & _
			" FROM T_EMAILSNDSVC_REMETENTE" & _
			" WHERE" & _
				" (email_remetente = '" & email_remetente_sistema & "')"
	set tER = cn.Execute(s_sql)
	do while Not tER.Eof
		if Trim("" & tER("st_envio_mensagem_habilitado")) = Cstr(ESS_ENVIO_HABILITADO) then
			id_remetente = CLng(tER("id"))
			exit do
			end if
		tER.MoveNext
		loop

	if tER.State <> 0 then tER.Close
	set tER = nothing

	if id_remetente = 0 then
		msg_erro_grava_msg = "O remetente informado não está cadastrado ou está com o envio de mensagens desabilitado"
		exit function
		end if

	' Verifica se foi informado o email do usuário do sistema responsável pelo envio da mensagem
	if remetente_from <> "" then
		if Not isEmailOk(remetente_from) then
			msg_erro_grava_msg = "E-mail do usuário remetente preenchido com caracteres inválidos"
			exit function
			end if
		end if

	' Verifica se o e-mail do destinatário está preenchido com caracteres válidos
	if destinatario_To <> "" then
		email_aux = Replace(destinatario_To, " ", ",")
		email_aux = Replace(email_aux, ";", ",")
		v = Split(email_aux,",")
		for i=LBound(v) to UBound(v)
			if Trim("" & v(i)) <> "" then
				if Not isEmailOk(Trim("" & v(i))) then
					msg_erro_grava_msg = "E-mail do campo <<Para>> preenchido com caracteres inválidos"
					exit function
				else
					existeEmailDeDestino = True
					end if
				end if
			next
		end if

	' Verifica se o e-mail de cópia está preenchido com caracteres válidos
	if destinatario_Cc <> "" then
		email_aux = Replace(destinatario_Cc, " ", ",")
		email_aux = Replace(email_aux, ";", ",")
		v = Split(email_aux,",")
		for i=LBound(v) to UBound(v)
			if Trim("" & v(i)) <> "" then
				if Not isEmailOk(Trim("" & v(i))) then
					msg_erro_grava_msg = "E-mail do campo <<Com cópia>> preenchido com caracteres inválidos"
					exit function
				else
					existeEmailDeDestino = True
					end if
				end if
			next
		end if

	' Verifica se o e-mail de cópia oculta está preenchido com caracteres válidos
	if destinatario_Cco <> "" then
		email_aux = Replace(destinatario_Cco, " ", ",")
		email_aux = Replace(email_aux, ";", ",")
		v = Split(email_aux,",")
		for i=LBound(v) to UBound(v)
			if Trim("" & v(i)) <> "" then
				if Not isEmailOk(Trim("" & v(i))) then
					msg_erro_grava_msg = "E-mail do campo <<Com cópia oculta>> preenchido com caracteres inválidos"
					exit function
				else
					existeEmailDeDestino = True
					end if
				end if
			next
		end if

	' Há algum campo de destinatário preenchido?
	if Not existeEmailDeDestino then
		msg_erro_grava_msg = "Não foi fornecido e-mail de nenhum destinatário para a mensagem"
		exit function
		end if

	' Verificar se o campo de assunto está preenchido
	if assunto = "" then
		msg_erro_grava_msg = "O campo Assunto não foi preenchido"
		exit function
		end if

	' Verificar se o corpo da mensagem está preenchido
	if corpo_mensagem = "" then
		msg_erro_grava_msg = "O corpo da mensagem não foi preenchido"
		exit function
		end if

	' Gravação
	if Not fin_gera_nsu(T_EMAILSNDSVC_MENSAGEM, id_mensagem, msg_erro_grava_msg) then
		msg_erro_grava_msg = "Problema na geração do NSU da mensagem"
		exit function
		end if

	if Not cria_recordset_otimista(tEM, msg_erro) then
		msg_erro_grava_msg = msg_erro
		exit function
		end if

	s_sql = "SELECT " & _
				"*" & _
			" FROM T_EMAILSNDSVC_MENSAGEM" & _
			" WHERE" & _
				" (id = -1)"
	tEM.Open s_sql, cn
	if Err <> 0 then
		msg_erro_grava_msg = Err.Description
		exit function
		end if

	tEM.AddNew
	tEM("id") = id_mensagem
	tEM("id_remetente") = id_remetente
	if FLAG_PEDIDO_CHAMADO_EMAIL_USAR_REPLY_TO And (remetente_from <> "") then
		tEM("replyToMsg") = remetente_from
		tEM("st_replyToMsg") = 1
		end if
	if destinatario_To <> "" then tEM("destinatario_To") = destinatario_To
	if destinatario_Cc <> "" then tEM("destinatario_Cc") = destinatario_Cc
	if destinatario_Cco <> "" then tEM("destinatario_Cco") = destinatario_Cco
	tEM("assunto") = assunto
	tEM("corpo_mensagem") = corpo_mensagem
	if Not IsNull(dt_hr_agendamento_envio) then
		if IsDate(dt_hr_agendamento_envio) then
			tEM("dt_hr_agendamento_envio") = dt_hr_agendamento_envio
			end if
		end if
	tEM.Update

	if Err <> 0 then
		msg_erro_grava_msg = Err.Description
		exit function
		end if

	if tEM.State <> 0 then tEM.Close
	set tEM = nothing

	svclog_horario_agendamento = "Imediato"
	if Not IsNull(dt_hr_agendamento_envio) then
		if IsDate(dt_hr_agendamento_envio) then
			svclog_horario_agendamento = formata_data_hora(dt_hr_agendamento_envio)
			end if
		end if

	svclog_descricao = "Gravação de email na fila de mensagens: sucesso" & " (id=" & Cstr(id_mensagem) & ", dt_hr_agendamento_envio=" & svclog_horario_agendamento & ")"
	svclog_complemento_1 = "To: " & monta_campo_log(destinatario_To) & vbLf & _
							"Cc: " & monta_campo_log(destinatario_Cc) & vbLf & _
							"Cco: " & monta_campo_log(destinatario_Cco)
	svclog_complemento_2 = "Assunto:" & vbLf & assunto
	svclog_complemento_3 = "Corpo:" & vbLf & corpo_mensagem
	call gravaFinSvcLog(NOME_DESTA_ROTINA, _
						T_EMAILSNDSVC_MENSAGEM, _
						svclog_descricao, _
						svclog_complemento_1, _
						svclog_complemento_2, _
						svclog_complemento_3, _
						"", _
						"", _
						"", _
						msg_erro)

	EmailSndSvcGravaMensagemParaEnvio = True
end function


' _______________________________________________________
' obtemCtrlEstoqueProdutoRegra
'
function obtemCtrlEstoqueProdutoRegra(byval uf, byval tipo_cliente, byval contribuinte_icms_status, byval produtor_rural_status, byref vProdRegra, byref msg_erro)
dim s_sql, r, tNE, iProd, iCD, tipo_pessoa, id_wms_regra_cd, idxCd
	obtemCtrlEstoqueProdutoRegra = False
	msg_erro = ""

	If Not cria_recordset_otimista(r, msg_erro) then
		msg_erro = "Falha ao tentar criar recordset"
		exit function
		end if

	If Not cria_recordset_otimista(tNE, msg_erro) then
		msg_erro = "Falha ao tentar criar recordset para consultar dados da t_NFe_EMITENTE"
		exit function
		end if

	tipo_pessoa = multi_cd_regra_determina_tipo_pessoa(tipo_cliente, contribuinte_icms_status, produtor_rural_status)
	if tipo_pessoa = "" then
		msg_erro = "Falha na leitura da regra de consumo do estoque para a UF '" & uf & "': não foi possível determinar o tipo de pessoa (tipo_cliente=" & tipo_cliente & ", contribuinte_icms_status=" & contribuinte_icms_status & ", produtor_rural_status=" & produtor_rural_status & ")"
		exit function
		end if

	for iProd=LBound(vProdRegra) to UBound(vProdRegra)
		if Trim("" & vProdRegra(iProd).produto) <> "" then
			s = "SELECT * FROM t_PRODUTO_X_WMS_REGRA_CD WHERE (fabricante = '" & vProdRegra(iProd).fabricante & "') AND (produto = '" & vProdRegra(iProd).produto & "')"
			if r.State <> 0 then r.Close
			r.open s, cn
			if r.Eof then
				vProdRegra(iProd).msg_erro = "Falha na leitura da regra de consumo do estoque para a UF '" & uf & "' e '" & descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa) & "': produto (" & vProdRegra(iProd).fabricante & ")" & vProdRegra(iProd).produto & " não possui regra associada"
			else
				id_wms_regra_cd = CLng(r("id_wms_regra_cd"))
				if id_wms_regra_cd = 0 then
					vProdRegra(iProd).msg_erro = "Falha na leitura da regra de consumo do estoque para a UF '" & uf & "' e '" & descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa) & "': produto (" & vProdRegra(iProd).fabricante & ")" & vProdRegra(iProd).produto & " não está associado a nenhuma regra"
				else
					s = "SELECT * FROM t_WMS_REGRA_CD WHERE (id = " & id_wms_regra_cd & ")"
					if r.State <> 0 then r.Close
					r.open s, cn
					if r.Eof then
						vProdRegra(iProd).msg_erro = "Falha na leitura da regra de consumo do estoque para a UF '" & uf & "' e '" & descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa) & "': regra associada ao produto (" & vProdRegra(iProd).fabricante & ")" & vProdRegra(iProd).produto & " não foi localizada no banco de dados (Id=" & id_wms_regra_cd & ")"
					else
						vProdRegra(iProd).st_regra_ok = True
						vProdRegra(iProd).regra.id = r("id")
						vProdRegra(iProd).regra.st_inativo = r("st_inativo")
						vProdRegra(iProd).regra.apelido = Trim("" & r("apelido"))
						vProdRegra(iProd).regra.descricao = Trim("" & r("descricao"))
						s = "SELECT * FROM t_WMS_REGRA_CD_X_UF WHERE (id_wms_regra_cd = " & id_wms_regra_cd & ") AND (uf = '" & uf & "')"
						if r.State <> 0 then r.Close
						r.open s, cn
						if r.Eof then
							vProdRegra(iProd).st_regra_ok = False
							vProdRegra(iProd).msg_erro = "Falha na leitura da regra de consumo do estoque para a UF '" & uf & "' e '" & descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa) & "': regra associada ao produto (" & vProdRegra(iProd).fabricante & ")" & vProdRegra(iProd).produto & " não está cadastrada para a UF '" & uf & "' (Id=" & id_wms_regra_cd & ")"
						else
							vProdRegra(iProd).regra.regraUF.id = r("id")
							vProdRegra(iProd).regra.regraUF.id_wms_regra_cd = r("id_wms_regra_cd")
							vProdRegra(iProd).regra.regraUF.uf = Trim("" & r("uf"))
							vProdRegra(iProd).regra.regraUF.st_inativo = r("st_inativo")
							s = "SELECT * FROM t_WMS_REGRA_CD_X_UF_X_PESSOA WHERE (id_wms_regra_cd_x_uf = " & vProdRegra(iProd).regra.regraUF.id & ") AND (tipo_pessoa = '" & tipo_pessoa & "')"
							if r.State <> 0 then r.Close
							r.open s, cn
							if r.Eof then
								vProdRegra(iProd).st_regra_ok = False
								vProdRegra(iProd).msg_erro = "Falha na leitura da regra de consumo do estoque para a UF '" & uf & "' e '" & descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa) & "': regra associada ao produto (" & vProdRegra(iProd).fabricante & ")" & vProdRegra(iProd).produto & " não está cadastrada para '" & descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa) & "' (Id=" & id_wms_regra_cd & ")"
							else
								vProdRegra(iProd).regra.regraUF.regraPessoa.id = r("id")
								vProdRegra(iProd).regra.regraUF.regraPessoa.id_wms_regra_cd_x_uf = r("id_wms_regra_cd_x_uf")
								vProdRegra(iProd).regra.regraUF.regraPessoa.tipo_pessoa = Trim("" & r("tipo_pessoa"))
								vProdRegra(iProd).regra.regraUF.regraPessoa.st_inativo = r("st_inativo")
								vProdRegra(iProd).regra.regraUF.regraPessoa.spe_id_nfe_emitente = r("spe_id_nfe_emitente")
								if converte_numero(vProdRegra(iProd).regra.regraUF.regraPessoa.spe_id_nfe_emitente) = 0 then
									vProdRegra(iProd).st_regra_ok = False
									vProdRegra(iProd).msg_erro = "Falha na leitura da regra de consumo do estoque para a UF '" & uf & "' e '" & descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa) & "': regra associada ao produto (" & vProdRegra(iProd).fabricante & ")" & vProdRegra(iProd).produto & " não especifica nenhum CD para aguardar produtos sem presença no estoque (Id=" & id_wms_regra_cd & ")"
								else
									'VERIFICA O CADASTRO PRINCIPAL DO NFe EMITENTE PARA VERIFICAR SE ESTÁ HABILITADO
									s = "SELECT * FROM t_NFe_EMITENTE WHERE (id = " & r("spe_id_nfe_emitente") & ")"
									if tNE.State <> 0 then tNE.Close
									tNE.open s, cn
									if Not tNE.Eof then
										if tNE("st_ativo") <> 1 then
											vProdRegra(iProd).st_regra_ok = False
											vProdRegra(iProd).msg_erro = "Falha na regra de consumo do estoque para a UF '" & uf & "' e '" & descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa) & "': regra associada ao produto (" & vProdRegra(iProd).fabricante & ")" & vProdRegra(iProd).produto & " especifica um CD para aguardar produtos sem presença no estoque que não está habilitado (Id=" & id_wms_regra_cd & ")"
											end if
										end if
									end if
								s = "SELECT * FROM t_WMS_REGRA_CD_X_UF_X_PESSOA_X_CD WHERE (id_wms_regra_cd_x_uf_x_pessoa = " & vProdRegra(iProd).regra.regraUF.regraPessoa.id & ") ORDER BY ordem_prioridade"
								if r.State <> 0 then r.Close
								r.open s, cn
								if r.Eof then
									vProdRegra(iProd).st_regra_ok = False
									vProdRegra(iProd).msg_erro = "Falha na leitura da regra de consumo do estoque para a UF '" & uf & "' e '" & descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa) & "': regra associada ao produto (" & vProdRegra(iProd).fabricante & ")" & vProdRegra(iProd).produto & " não especifica nenhum CD para consumo do estoque (Id=" & id_wms_regra_cd & ")"
								else
									idxCd = LBound(vProdRegra(iProd).regra.regraUF.regraPessoa.vCD) - 1
									do while Not r.Eof
										idxCd = idxCd + 1
										vProdRegra(iProd).regra.regraUF.regraPessoa.vCD(idxCd).id = r("id")
										vProdRegra(iProd).regra.regraUF.regraPessoa.vCD(idxCd).id_wms_regra_cd_x_uf_x_pessoa = r("id_wms_regra_cd_x_uf_x_pessoa")
										vProdRegra(iProd).regra.regraUF.regraPessoa.vCD(idxCd).id_nfe_emitente = r("id_nfe_emitente")
										vProdRegra(iProd).regra.regraUF.regraPessoa.vCD(idxCd).ordem_prioridade = r("ordem_prioridade")
										vProdRegra(iProd).regra.regraUF.regraPessoa.vCD(idxCd).st_inativo = r("st_inativo")
										
										'VERIFICA O CADASTRO PRINCIPAL DO NFe EMITENTE PARA VERIFICAR SE ESTÁ HABILITADO
										s = "SELECT * FROM t_NFe_EMITENTE WHERE (id = " & r("id_nfe_emitente") & ")"
										if tNE.State <> 0 then tNE.Close
										tNE.open s, cn
										if Not tNE.Eof then
											if tNE("st_ativo") <> 1 then vProdRegra(iProd).regra.regraUF.regraPessoa.vCD(idxCd).st_inativo = 1
											end if

										r.MoveNext
										loop

									'VERIFICA SE O CD SELECIONADO PARA ALOCAR OS PRODUTOS PENDENTES ESTÁ HABILITADO NA LISTA DOS CD'S P/ PRODUTOS DISPONÍVEIS
									'OBSERVAÇÃO: ESTA CHECAGEM É FEITA POR UMA QUESTÃO DE COERÊNCIA, POIS SE OCORRER A SITUAÇÃO EM QUE UM CD SELECIONADO P/
									'PRODUTOS PENDENTES NÃO ESTEJA HABILITADO P/ OS PRODUTOS DISPONÍVEIS, O SISTEMA IRIA PROCESSAR DA SEGUINTE FORMA:
									'	1) AO ANALISAR A DISPONIBILIDADE DOS PRODUTOS NOS CD'S, SE O CD 'X' ESTIVER DESATIVADO ELE SERÁ IGNORADO.
									'	2) SE AO FINAL DA ALOCAÇÃO DOS PRODUTOS DISPONÍVEIS RESTAR UMA QUANTIDADE PENDENTE P/ SER ALOCADA COMO
									'		PRODUTO SEM PRESENÇA NO ESTOQUE E, CASO O CD SELECIONADO P/ TAL SEJA O CD 'X', O SISTEMA IRÁ ALOCAR A
									'		QUANTIDADE REMANESCENTE P/ O CD 'X', SENDO QUE SE ESSE CD POSSUIR DISPONIBILIDADE NO ESTOQUE, OS PRODUTOS SERÃO
									'		CONSUMIDOS NORMALMENTE AO INVÉS DE FICAREM COMO PENDENTES.
									for iCD=LBound(vProdRegra(iProd).regra.regraUF.regraPessoa.vCD) to UBound(vProdRegra(iProd).regra.regraUF.regraPessoa.vCD)
										if vProdRegra(iProd).regra.regraUF.regraPessoa.vCD(iCD).id_nfe_emitente = vProdRegra(iProd).regra.regraUF.regraPessoa.spe_id_nfe_emitente then
											if vProdRegra(iProd).regra.regraUF.regraPessoa.vCD(iCD).st_inativo = 1 then
												vProdRegra(iProd).st_regra_ok = False
												vProdRegra(iProd).msg_erro = "Falha na leitura da regra de consumo do estoque para a UF '" & uf & "' e '" & descricao_multi_CD_regra_tipo_pessoa(tipo_pessoa) & "': regra associada ao produto (" & vProdRegra(iProd).fabricante & ")" & vProdRegra(iProd).produto & " especifica o CD '" & obtem_apelido_empresa_NFe_emitente(vProdRegra(iProd).regra.regraUF.regraPessoa.spe_id_nfe_emitente) & "' para alocação de produtos sem presença no estoque, sendo que este CD está desativado para processar produtos disponíveis (Id=" & id_wms_regra_cd & ")"
												end if
											exit for
											end if
										next
									end if
								end if
							end if
						end if
					end if
				end if
			end if
		next 'for iProd

	if r.State <> 0 then r.Close
	set r = nothing

	if tNE.State <> 0 then tNE.Close
	set tNE = nothing

	obtemCtrlEstoqueProdutoRegra = True
end function


' ___________________________________________
' LE INDICADOR
'
function le_indicador(byval apelido, byref r_indicador, byref msg_erro)
dim s
dim rs

	le_indicador = False
	msg_erro = ""
	apelido=Trim("" & apelido)
	set r_indicador = New cl_INDICADOR
	s="SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido='" & apelido & "')"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
	if rs.EOF then
		msg_erro="Indicador " & apelido & " não está cadastrado."
	else
		with r_indicador
			.apelido					= rs("apelido")
			.Id							= rs("Id")
			.id_magento_b2b				= rs("id_magento_b2b")
			.cnpj_cpf					= rs("cnpj_cpf")
			.tipo						= rs("tipo")
			.ie_rg						= Trim("" & rs("ie_rg"))
			.razao_social_nome			= Trim("" & rs("razao_social_nome"))
			.endereco					= Trim("" & rs("endereco"))
			.endereco_numero			= Trim("" & rs("endereco_numero"))
			.endereco_complemento		= Trim("" & rs("endereco_complemento"))
			.bairro					    = Trim("" & rs("bairro"))
			.cidade						= Trim("" & rs("cidade"))
			.uf 						= Trim("" & rs("uf"))
			.cep				        = Trim("" & rs("cep"))
			.ddd				        = Trim("" & rs("ddd"))
			.telefone				    = Trim("" & rs("telefone"))
			.fax				        = Trim("" & rs("fax"))
			.ddd_cel			        = Trim("" & rs("ddd_cel"))
			.tel_cel				    = Trim("" & rs("tel_cel"))
			.contato			        = Trim("" & rs("contato"))
			.banco			            = Trim("" & rs("banco"))
			.agencia				    = Trim("" & rs("agencia"))
			.conta		                = Trim("" & rs("conta"))
			.favorecido				    = Trim("" & rs("favorecido"))
			.favorecido_cnpj_cpf		= Trim("" & rs("favorecido_cnpj_cpf"))
			.agencia_dv				    = Trim("" & rs("agencia_dv"))
			.conta_operacao 			= Trim("" & rs("conta_operacao"))
			.conta_dv					= Trim("" & rs("conta_dv"))
			.tipo_conta				    = Trim("" & rs("tipo_conta"))
			.loja			            = Trim("" & rs("loja"))
			.vendedor			        = Trim("" & rs("vendedor"))
			.email	                    = Trim("" & rs("email"))
			.email2	                    = Trim("" & rs("email2"))
			.email3	                    = Trim("" & rs("email3"))
			.captador	                = Trim("" & rs("captador"))
			.nome_fantasia			    = Trim("" & rs("nome_fantasia"))
			.razao_social_nome_iniciais_em_maiusculas			= Trim("" & rs("razao_social_nome_iniciais_em_maiusculas"))
			end with
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_indicador=True
end function


' ________________________________________________________
' isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
'
function isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
dim rFPUECM
	isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = False
	set rFPUECM = get_registro_t_parametro(ID_PARAMETRO_Flag_Pedido_MemorizacaoCompletaEnderecos)
	if Trim("" & rFPUECM.campo_inteiro) = "1" then isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = True
	set rFPUECM = Nothing
end function


' ________________________________________________________
' isActivatedFlagCadParceiroDadosBancariosEdicaoBloqueada
'
function isActivatedFlagCadParceiroDadosBancariosEdicaoBloqueada
dim rFCPDBEB
	isActivatedFlagCadParceiroDadosBancariosEdicaoBloqueada = False
	set rFCPDBEB = get_registro_t_parametro(ID_PARAMETRO_Flag_CadastroParceiro_DadosBancarios_EdicaoBloqueada)
	if Trim("" & rFCPDBEB.campo_inteiro) = "1" then isActivatedFlagCadParceiroDadosBancariosEdicaoBloqueada = True
	set rFCPDBEB = Nothing
end function

' ________________________________________________________
' isActivatedFlagCadSemiAutoPedMagentoCadAutoClienteNovo
'
function isActivatedFlagCadSemiAutoPedMagentoCadAutoClienteNovo
dim rFCSAPMCACN
	isActivatedFlagCadSemiAutoPedMagentoCadAutoClienteNovo = False
	set rFCSAPMCACN = get_registro_t_parametro(ID_PARAMETRO_FLAG_CAD_SEMI_AUTO_PED_MAGENTO_CADASTRAR_AUTOMATICAMENTE_CLIENTE_NOVO)
	if Trim("" & rFCSAPMCACN.campo_inteiro) = "1" then isActivatedFlagCadSemiAutoPedMagentoCadAutoClienteNovo = True
	set rFCSAPMCACN = Nothing
end function


' __________________________________________________________________
' isActivatedFlagCadSemiAutoPedMagentoUsarCamposValorMktpDataSource
'
function isActivatedFlagCadSemiAutoPedMagentoUsarCamposValorMktpDataSource
dim rFCSAPMUCVMDS
	isActivatedFlagCadSemiAutoPedMagentoUsarCamposValorMktpDataSource = False
	set rFCSAPMUCVMDS = get_registro_t_parametro(ID_PARAMETRO_FLAG_CAD_SEMI_AUTO_PED_MAGENTO_USAR_CAMPOS_VALOR_MKTP_DATASOURCE)
	if Trim("" & rFCSAPMUCVMDS.campo_inteiro) = "1" then isActivatedFlagCadSemiAutoPedMagentoUsarCamposValorMktpDataSource = True
	set rFCSAPMUCVMDS = Nothing
end function


' __________________________________________________________________
' isActivatedFlagCadSemiAutoPedMagentoRateioFreteAutomatico
'
function isActivatedFlagCadSemiAutoPedMagentoRateioFreteAutomatico
dim rFCSAPMRFA
	isActivatedFlagCadSemiAutoPedMagentoRateioFreteAutomatico = False
	set rFCSAPMRFA = get_registro_t_parametro(ID_PARAMETRO_FLAG_CAD_SEMI_AUTO_PED_MAGENTO_RATEIO_FRETE_AUTOMATICO)
	if Trim("" & rFCSAPMRFA.campo_inteiro) = "1" then isActivatedFlagCadSemiAutoPedMagentoRateioFreteAutomatico = True
	set rFCSAPMRFA = Nothing
end function


' ________________________________________________________
' getParametroPercDesagioRALiquida
'
function getParametroPercDesagioRALiquida
dim rP
	getParametroPercDesagioRALiquida = 0
	set rP = get_registro_t_parametro(ID_PARAMETRO_PERC_DESAGIO_RA_LIQUIDA)
	if Trim("" & rP.campo_real) <> "" then getParametroPercDesagioRALiquida = rP.campo_real
	set rP = Nothing
end function


' ________________________________________________________
' getParametroPrazoAcessoRelPedidosIndicadoresLoja
'
function getParametroPrazoAcessoRelPedidosIndicadoresLoja
dim rP
	getParametroPrazoAcessoRelPedidosIndicadoresLoja = 0
	set rP = get_registro_t_parametro(ID_PARAMETRO_PRAZO_ACESSO_REL_PEDIDOS_INDICADORES_LOJA)
	if Trim("" & rP.campo_inteiro) <> "" then getParametroPrazoAcessoRelPedidosIndicadoresLoja = rP.campo_inteiro
	set rP = Nothing
end function


' ___________________________________________________________
' getParametroPrazoAcessoRelPedidosIndicadoresLojaRtPendente
'
function getParametroPrazoAcessoRelPedidosIndicadoresLojaRtPendente(byval unidade_negocio)
dim rP
	getParametroPrazoAcessoRelPedidosIndicadoresLojaRtPendente = 0
	
	if Trim("" & unidade_negocio) = "" then exit function

	set rP = get_registro_t_parametro(ID_PARAMETRO_PRAZO_ACESSO_REL_PEDIDOS_INDICADORES_LOJA_RT_PENDENTE)
	if Trim("" & rP.campo_inteiro) <> "" then
		if Instr(Ucase(Trim("" & rP.campo_texto)), "|" & Ucase(unidade_negocio) & "|") <> 0 then
			getParametroPrazoAcessoRelPedidosIndicadoresLojaRtPendente = rP.campo_inteiro
			end if
		end if
	set rP = Nothing
end function


' ___________________________________
' isLojaVrf
'
function isLojaVrf(byval loja)
dim s, tLAux
	isLojaVrf = False

	loja  = Trim("" & loja)
	
	s = "SELECT * FROM t_LOJA WHERE (loja = '" & loja & "') AND (unidade_negocio = '" & COD_UNIDADE_NEGOCIO_LOJA__VRF & "')"
	set tLAux = cn.Execute(s)
	if Not tLAux.Eof then
		isLojaVrf = True
		tLAux.Close
		set tLAux = nothing
		end if
end function


' ___________________________________
' isLojaBonshop
'
function isLojaBonshop(byval loja)
dim s, tLAux
	isLojaBonshop = False

	loja  = Trim("" & loja)
	
	s = "SELECT * FROM t_LOJA WHERE (loja = '" & loja & "') AND (unidade_negocio = '" & COD_UNIDADE_NEGOCIO_LOJA__BS & "')"
	set tLAux = cn.Execute(s)
	if Not tLAux.Eof then
		isLojaBonshop = True
		tLAux.Close
		set tLAux = nothing
		end if
end function


' ___________________________________
' isLojaGarantia
'
function isLojaGarantia(byval loja)
dim s, tLAux
	isLojaGarantia = False

	loja  = Trim("" & loja)
	
	s = "SELECT * FROM t_LOJA WHERE (loja = '" & loja & "') AND (unidade_negocio = '" & COD_UNIDADE_NEGOCIO_LOJA__GARANTIA & "')"
	set tLAux = cn.Execute(s)
	if Not tLAux.Eof then
		isLojaGarantia = True
		tLAux.Close
		set tLAux = nothing
		end if
end function


' ------------------------------------------------------------------------
'   isLojaHabilitadaProdCompostoECommerce
'
function isLojaHabilitadaProdCompostoECommerce(byval loja)
dim blnLojaHabilitada
	isLojaHabilitadaProdCompostoECommerce = False
	loja = Trim("" & loja)
	blnLojaHabilitada = False
	if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
		blnLojaHabilitada = True
	elseif isLojaBonshop(loja) then
		blnLojaHabilitada = True
	elseif isLojaVrf(loja) then
		blnLojaHabilitada = True
	elseif loja = NUMERO_LOJA_MARCELO_ARTVEN then
		blnLojaHabilitada = True
		end if

	if blnLojaHabilitada = True then isLojaHabilitadaProdCompostoECommerce = True
end function


' ------------------------------------------------------------------------
'   le_usuario
'
function le_usuario(byval id_usuario, byref r_usuario, byref msg_erro)
dim s
dim rs

	le_usuario = False
	msg_erro = ""
	id_usuario = Trim("" & id_usuario)
	set r_usuario = New cl_USUARIO
	s="SELECT * FROM t_USUARIO WHERE (usuario = '" & id_usuario & "')"
	set rs=cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.EOF then
		msg_erro="Usuário " & id_usuario & " não está cadastrado."
	else
		with r_usuario
			.usuario = Trim("" & rs("usuario"))
			.nivel = Trim("" & rs("nivel"))
			.loja = Trim("" & rs("loja"))
			.senha = Trim("" & rs("senha"))
			.nome = Trim("" & rs("nome"))
			.datastamp = Trim("" & rs("datastamp"))
			.bloqueado = rs("bloqueado")
			.dt_cadastro = rs("dt_cadastro")
			.dt_ult_atualizacao = rs("dt_ult_atualizacao")
			.dt_ult_alteracao_senha = rs("dt_ult_alteracao_senha")
			.dt_ult_acesso = rs("dt_ult_acesso")
			.vendedor_externo = rs("vendedor_externo")
			.vendedor_loja = rs("vendedor_loja")
			.SessionCtrlTicket = Trim("" & rs("SessionCtrlTicket"))
			.SessionCtrlLoja = Trim("" & rs("SessionCtrlLoja"))
			.SessionCtrlModulo = Trim("" & rs("SessionCtrlModulo"))
			.SessionCtrlDtHrLogon = rs("SessionCtrlDtHrLogon")
			.fin_email_remetente = Trim("" & rs("fin_email_remetente"))
			.fin_servidor_smtp = Trim("" & rs("fin_servidor_smtp"))
			.fin_usuario_smtp = Trim("" & rs("fin_usuario_smtp"))
			.fin_senha_smtp = Trim("" & rs("fin_senha_smtp"))
			.fin_display_name_remetente = Trim("" & rs("fin_display_name_remetente"))
			.nome_iniciais_em_maiusculas = Trim("" & rs("nome_iniciais_em_maiusculas"))
			.fin_servidor_smtp_porta = rs("fin_servidor_smtp_porta")
			.email = Trim("" & rs("email"))
			.SessionTokenModuloCentral = Trim("" & rs("SessionTokenModuloCentral"))
			.DtHrSessionTokenModuloCentral = rs("DtHrSessionTokenModuloCentral")
			.SessionTokenModuloLoja = Trim("" & rs("SessionTokenModuloLoja"))
			.DtHrSessionTokenModuloLoja = rs("DtHrSessionTokenModuloLoja")
			.fin_smtp_enable_ssl = rs("fin_smtp_enable_ssl")
			.Id = rs("Id")
			.QtdeConsecutivaFalhaLogin = rs("QtdeConsecutivaFalhaLogin")
			.StLoginBloqueadoAutomatico = rs("StLoginBloqueadoAutomatico")
			.DataHoraBloqueadoAutomatico = rs("DataHoraBloqueadoAutomatico")
			.EnderecoIpBloqueadoAutomatico = Trim("" & rs("EnderecoIpBloqueadoAutomatico"))
			end with
		end if

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	with r_usuario
		.nivel_acesso_bloco_notas_pedido = COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__NAO_DEFINIDO
		.nivel_acesso_chamado = COD_NIVEL_ACESSO_CHAMADO_PEDIDO__NAO_DEFINIDO
		
		s="SELECT Coalesce(Max(nivel_acesso_bloco_notas_pedido), " & Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__NAO_DEFINIDO) & ") AS max_nivel_acesso_bloco_notas_pedido FROM t_PERFIL INNER JOIN t_PERFIL_X_USUARIO ON t_PERFIL.id=t_PERFIL_X_USUARIO.id_perfil WHERE (usuario = '" & id_usuario & "')"
		if rs.State <> 0 then rs.Close
		set rs=cn.Execute(s)
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if

		if Not rs.Eof then
			.nivel_acesso_bloco_notas_pedido = rs("max_nivel_acesso_bloco_notas_pedido")
			end if
		
		s="SELECT Coalesce(Max(nivel_acesso_chamado), " & Cstr(COD_NIVEL_ACESSO_CHAMADO_PEDIDO__NAO_DEFINIDO) & ") AS max_nivel_acesso_chamado FROM t_PERFIL INNER JOIN t_PERFIL_X_USUARIO ON t_PERFIL.id=t_PERFIL_X_USUARIO.id_perfil WHERE (usuario = '" & id_usuario & "')"
		if rs.State <> 0 then rs.Close
		set rs=cn.Execute(s)
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if

		if Not rs.Eof then
			.nivel_acesso_chamado = rs("max_nivel_acesso_chamado")
			end if
		end with

	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.State <> 0 then rs.Close

	if msg_erro = "" then le_usuario = True
end function


' ________________________________________________________
' getParametroFromCampoTexto
'
function getParametroFromCampoTexto(ByVal id_registro)
dim rP
	set rP = get_registro_t_parametro(id_registro)
	getParametroFromCampoTexto = rP.campo_texto
	set rP = Nothing
end function


' ________________________________________________________
' getParametroFromCampoInteiro
'
function getParametroFromCampoInteiro(ByVal id_registro)
dim rP
	set rP = get_registro_t_parametro(id_registro)
	getParametroFromCampoInteiro = rP.campo_inteiro
	set rP = Nothing
end function


' ________________________________________________________
' getParametroFromCampoReal
'
function getParametroFromCampoReal(ByVal id_registro)
dim rP
	set rP = get_registro_t_parametro(id_registro)
	getParametroFromCampoReal = rP.campo_real
	set rP = Nothing
end function


' ________________________________________________________
' getParametroFromCampoMonetario
'
function getParametroFromCampoMonetario(ByVal id_registro)
dim rP
	set rP = get_registro_t_parametro(id_registro)
	getParametroFromCampoMonetario = rP.campo_monetario
	set rP = Nothing
end function


' ________________________________________________________
' getParametroFromCampoData
'
function getParametroFromCampoData(ByVal id_registro)
dim rP
	set rP = get_registro_t_parametro(id_registro)
	getParametroFromCampoData = rP.campo_data
	set rP = Nothing
end function


' ________________________________________________________
' getParametroFromCampoData
'
function grava_bloco_notas_pedido(ByVal numeroPedido, ByVal idUsuario, ByVal numeroLoja, ByVal nivelAcesso, ByVal mensagem, ByVal tipo_mensagem, ByRef msg_erro)
dim s, intNsuNovoBlocoNotas, msg_erro_aux
dim tBN
dim campos_a_omitir
dim vLog()
dim s_log

	grava_bloco_notas_pedido = False
	msg_erro = ""

	if Not fin_gera_nsu(T_PEDIDO_BLOCO_NOTAS, intNsuNovoBlocoNotas, msg_erro_aux) then
		msg_erro = "Falha ao gerar NSU para o novo registro de bloco de notas (" & msg_erro_aux & ")!"
		exit function
	else
		if intNsuNovoBlocoNotas <= 0 then
			msg_erro = "NSU gerado para o novo registro de bloco de notas é inválido (" & intNsuNovoBlocoNotas & ")!"
			exit function
			end if
		end if

	if Not cria_recordset_otimista(tBN, msg_erro_aux) then
		msg_erro = "Falha ao tentar criar recordset durante a gravação do bloco de notas (" & msg_erro_aux & ")!"
		exit function
		end if

	s = "SELECT * FROM t_PEDIDO_BLOCO_NOTAS WHERE (id = -1)"
	tBN.Open s, cn
	tBN.AddNew
	tBN("id") = intNsuNovoBlocoNotas
	tBN("pedido") = numeroPedido
	tBN("usuario") = idUsuario
	tBN("loja") = Trim("" & numeroLoja)
	tBN("nivel_acesso") = CLng(nivelAcesso)
	tBN("mensagem") = mensagem
	tBN("tipo_mensagem") = tipo_mensagem
	tBN.Update

	if Err <> 0 then
		msg_erro_grava_msg = Err.Description
		exit function
		end if

	s_log = ""
	campos_a_omitir = "|dt_cadastro|dt_hr_cadastro|anulado_status|anulado_usuario|anulado_data|anulado_data_hora|"

	log_via_vetor_carrega_do_recordset tBN, vLog, campos_a_omitir
	s_log = log_via_vetor_monta_inclusao(vLog)

	if tBN.State <> 0 then tBN.Close
	set tBN = nothing

	if s_log <> "" then grava_log idUsuario, numeroLoja, numeroPedido, "", OP_LOG_PEDIDO_BLOCO_NOTAS_INCLUSAO, s_log

	grava_bloco_notas_pedido = True
end function



' ________________________________________________________
' montaSubqueryGetUsuarioContexto
' Monta uma subquery para obter a identificação do usuário
' quando os dados armazenados de forma codificada:
'    [N] 999999
'        N = 1 -> Usuário interno (t_USUARIO.Id)
'        N = 2 -> Parceiro (t_ORCAMENTISTA_E_INDICADOR.Id)
'        N = 3 -> Vendedor do Parceiro (t_ORCAMENTISTA_E_INDICADOR_VENDEDOR.Id)
'        N = 4 -> Cliente
'    999999 = Id do registro
function montaSubqueryGetUsuarioContexto(byval nomeCampo, byval alias)
dim s_sql, s_sql_convert_int
	montaSubqueryGetUsuarioContexto = ""

	if Trim("" & nomeCampo) = "" then exit function

	s_sql_convert_int = " CONVERT(int, LTRIM(SUBSTRING(" & nomeCampo & ", CHARINDEX(']', " & nomeCampo & ", 1) + 1, LEN(" & nomeCampo & "))))"

	s_sql = " (CASE" & _
				" WHEN SUBSTRING(" & nomeCampo & ", 1, 3) = '[1]' THEN" & _
					" (SELECT usuario FROM t_USUARIO tU_SQAux WHERE tU_SQAux.Id = " & s_sql_convert_int & ")" & _
				" WHEN SUBSTRING(" & nomeCampo & ", 1, 3) = '[2]' THEN" & _
					" (SELECT apelido FROM t_ORCAMENTISTA_E_INDICADOR tOI_SQAux WHERE tOI_SQAux.Id = " & s_sql_convert_int & ")" & _
				" WHEN SUBSTRING(" & nomeCampo & ", 1, 3) = '[3]' THEN" & _
					" (SELECT Nome FROM t_ORCAMENTISTA_E_INDICADOR_VENDEDOR tOIV_SQAux WHERE tOIV_SQAux.Id = " & s_sql_convert_int & ")" & _
				" ELSE" & _
					" NULL" & _
			" END)"

	if Trim("" & alias) <> "" then s_sql = s_sql & " AS " & alias

	montaSubqueryGetUsuarioContexto = s_sql
end function


' _____________________________________________________________________
' obtem_perc_comissao_e_desconto_n1_n2_a_utilizar
' Obtém o percentual máximo de comissão e desconto (somente
' analisando os níveis 1 e 2, ou seja, não considerando as alçadas)
function obtem_perc_comissao_e_desconto_n1_n2_a_utilizar(byval tipoCliente, byref rPed, byref vItem)
dim i, vl_total_preco_venda, s_pg, perc_comissao_e_desconto_a_utilizar
dim rCD, rP, vMPN2
dim blnPreferencial, vlNivel1, vlNivel2

	obtem_perc_comissao_e_desconto_n1_n2_a_utilizar = 0

	set rCD = obtem_perc_max_comissao_e_desconto_por_loja(rPed.loja)
	
	set rP = get_registro_t_parametro(ID_PARAMETRO_PercMaxComissaoEDesconto_Nivel2_MeiosPagto)
	if Trim("" & rP.id) <> "" then
		vMPN2 = Split(rP.campo_texto, ",")
		for i=Lbound(vMPN2) to Ubound(vMPN2)
			vMPN2(i) = Trim("" & vMPN2(i))
			next
	else
		redim vMPN2(0)
		vMPN2(0) = ""
		end if
	
	vl_total_preco_venda = 0
	for i=Lbound(vItem) to Ubound(vItem)
		with vItem(i)
			if .produto <> "" then
				vl_total_preco_venda = vl_total_preco_venda + (.qtde * .preco_venda)
				end if
			end with
		next
	
	'Percentuais padrão (nível 1)
	if tipoCliente = ID_PJ then
		perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_pj
	else
		perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto
		end if
	
	if Cstr(rPed.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then
		s_pg = CStr(rPed.av_forma_pagto)
		if s_pg <> "" then
			for i=Lbound(vMPN2) to Ubound(vMPN2)
			'	O meio de pagamento selecionado é um dos preferenciais
				if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
					if tipoCliente = ID_PJ then
						perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
					else
						perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
						end if
					exit for
					end if
				next
			end if
	elseif Cstr(rPed.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then
		s_pg = CStr(rPed.pu_forma_pagto)
		if s_pg <> "" then
			for i=Lbound(vMPN2) to Ubound(vMPN2)
			'	O meio de pagamento selecionado é um dos preferenciais
				if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
					if tipoCliente = ID_PJ then
						perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
					else
						perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
						end if
					exit for
					end if
				next
			end if
	elseif Cstr(rPed.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then
		s_pg = Trim(ID_FORMA_PAGTO_CARTAO)
		if s_pg <> "" then
			for i=Lbound(vMPN2) to Ubound(vMPN2)
			'	O meio de pagamento selecionado é um dos preferenciais
				if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
					if tipoCliente = ID_PJ then
						perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
					else
						perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
						end if
					exit for
					end if
				next
			end if
	elseif Cstr(rPed.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
		s_pg = Trim(ID_FORMA_PAGTO_CARTAO_MAQUINETA)
		if s_pg <> "" then
			for i=Lbound(vMPN2) to Ubound(vMPN2)
			'	O meio de pagamento selecionado é um dos preferenciais
				if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
					if tipoCliente = ID_PJ then
						perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
					else
						perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
						end if
					exit for
					end if
				next
			end if
	elseif Cstr(rPed.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
	'	Identifica e contabiliza o valor da entrada
		blnPreferencial = False
		s_pg = CStr(rPed.pce_forma_pagto_entrada)
		if s_pg <> "" then
			for i=Lbound(vMPN2) to Ubound(vMPN2)
			'	O meio de pagamento selecionado é um dos preferenciais
				if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
					blnPreferencial = True
					exit for
					end if
				next
			end if
		
		if blnPreferencial then
			vlNivel2 = converte_numero(rPed.pce_entrada_valor)
		else
			vlNivel1 = converte_numero(rPed.pce_entrada_valor)
			end if
		
	'	Identifica e contabiliza o valor das parcelas
		blnPreferencial = False
		s_pg = CStr(rPed.pce_forma_pagto_prestacao)
		if s_pg <> "" then
			for i=Lbound(vMPN2) to Ubound(vMPN2)
			'	O meio de pagamento selecionado é um dos preferenciais
				if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
					blnPreferencial = True
					exit for
					end if
				next
			end if
		
		if blnPreferencial then
			vlNivel2 = vlNivel2 + converte_numero(rPed.pce_prestacao_qtde) * converte_numero(rPed.pce_prestacao_valor)
		else
			vlNivel1 = vlNivel1 + converte_numero(rPed.pce_prestacao_qtde) * converte_numero(rPed.pce_prestacao_valor)
			end if
	
	'	O montante a pagar por meio de pagamento preferencial é maior que 50% do total?
		if vlNivel2 > (vl_total_preco_venda/2) then
			if tipoCliente = ID_PJ then
				perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
			else
				perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
				end if
			end if
		
	elseif Cstr(rPed.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
	'	Identifica e contabiliza o valor da 1ª parcela
		blnPreferencial = False
		s_pg = CStr(rPed.pse_forma_pagto_prim_prest)
		if s_pg <> "" then
			for i=Lbound(vMPN2) to Ubound(vMPN2)
			'	O meio de pagamento selecionado é um dos preferenciais
				if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
					blnPreferencial = True
					exit for
					end if
				next
			end if
		
		if blnPreferencial then
			vlNivel2 = converte_numero(rPed.pse_prim_prest_valor)
		else
			vlNivel1 = converte_numero(rPed.pse_prim_prest_valor)
			end if
		
	'	Identifica e contabiliza o valor das parcelas
		blnPreferencial = False
		s_pg = CStr(rPed.pse_forma_pagto_demais_prest)
		if s_pg <> "" then
			for i=Lbound(vMPN2) to Ubound(vMPN2)
			'	O meio de pagamento selecionado é um dos preferenciais
				if Trim("" & s_pg) = Trim("" & vMPN2(i)) then
					blnPreferencial = True
					exit for
					end if
				next
			end if
		
		if blnPreferencial then
			vlNivel2 = vlNivel2 + converte_numero(rPed.pse_demais_prest_qtde) * converte_numero(rPed.pse_demais_prest_valor)
		else
			vlNivel1 = vlNivel1 + converte_numero(rPed.pse_demais_prest_qtde) * converte_numero(rPed.pse_demais_prest_valor)
			end if
		
	'	O montante a pagar por meio de pagamento preferencial é maior que 50% do total?
		if vlNivel2 > (vl_total_preco_venda/2) then
			if tipoCliente = ID_PJ then
				perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2_pj
			else
				perc_comissao_e_desconto_a_utilizar = rCD.perc_max_comissao_e_desconto_nivel2
				end if
			end if
		end if

	obtem_perc_comissao_e_desconto_n1_n2_a_utilizar = perc_comissao_e_desconto_a_utilizar
end function


' _______________________________________________________
' inicializa_cl_DEPTO_SETOR
'
sub inicializa_cl_DEPTO_SETOR(byref rDS)
	rDS.Id = 0
	rDS.Sigla = ""
	rDS.Nome = ""
	rDS.StInativo = 0
	rDS.Observacoes = ""
	rDS.UsuarioResponsavelN1 = ""
	rDS.UsuarioResponsavelN2 = ""
end sub

' _______________________________________________________
' obtem_Usuario_x_DeptoSetor
'
function obtem_Usuario_x_DeptoSetor(byval usuario, byref vDeptoSetor, byref msg_erro)
dim s_sql, r
	
	obtem_Usuario_x_DeptoSetor = False
	msg_erro = ""
	
	redim vDeptoSetor(0)
	set vDeptoSetor(UBound(vDeptoSetor)) = new cl_DEPTO_SETOR
	inicializa_cl_DEPTO_SETOR vDeptoSetor(UBound(vDeptoSetor))
	usuario = Trim("" & usuario)

	s_sql = "SELECT " & _
				"*" & _
			" FROM t_DEPTO_SETOR tDS" & _
				" INNER JOIN t_USUARIO_X_DEPTO_SETOR tUDS ON (tUDS.IdDeptoSetor = tDS.Id)" & _
			" WHERE" & _
				" (tDS.StInativo = 0)" & _
				" AND (tUDS.excluido_status = 0)" & _
				" AND (tUDS.usuario = '" & QuotedStr(usuario) & "')" & _
			" ORDER BY" & _
				" Id"
	set r=cn.Execute(s_sql)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if r.EOF then
		'O usuário pode não estar associado a nenhum depto/setor específico, isso não seria um erro, basta retornar a lista vazia
		obtem_Usuario_x_DeptoSetor = True
		if r.State <> 0 then r.Close
		exit function
	else
		do while Not r.EOF
			if vDeptoSetor(UBound(vDeptoSetor)).Id <> 0 then
				redim preserve vDeptoSetor(UBound(vDeptoSetor)+1)
				set vDeptoSetor(UBound(vDeptoSetor)) = new cl_DEPTO_SETOR
				inicializa_cl_DEPTO_SETOR vDeptoSetor(UBound(vDeptoSetor))
				end if
			with vDeptoSetor(UBound(vDeptoSetor))
				.Id = r("Id")
				.Sigla = Trim("" & r("Sigla"))
				.Nome = Trim("" & r("Nome"))
				.StInativo = r("StInativo")
				.Observacoes = Trim("" & r("Observacoes"))
				.UsuarioResponsavelN1 = Trim("" & r("UsuarioResponsavelN1"))
				.UsuarioResponsavelN2 = Trim("" & r("UsuarioResponsavelN2"))
				end with
			r.MoveNext
			loop
		if r.State <> 0 then r.Close
		end if

	obtem_Usuario_x_DeptoSetor = True
end function
%>
