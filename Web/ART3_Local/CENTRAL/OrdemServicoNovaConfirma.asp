<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  O R D E M S E R V I C O N O V A C O N F I R M A . A S P
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

	dim s, s_log, s_chave_OS, i, j, n, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim url_back, url_back_aux
	dim s_id_nfe_emitente
	dim s_tipo, s_loja_origem, s_pedido, s_op_descricao, c_obs_pecas_necessarias, s_aux
	dim v_aux, s_cod_estoque_origem, s_cod_estoque_destino, s_fluxo, qtde_transferida
	dim s_chave_OS_origem
	dim s_pedido_origem
	dim s_loja_destino
	dim s_pedido_destino
	dim alerta
	alerta=""

'	OBTÉM DADOS DO FORMULÁRIO
	url_back = Trim(Request("url_back"))
	s_id_nfe_emitente = Trim(Request.Form("c_id_nfe_emitente"))
	s_tipo = Ucase(Trim(Request.Form("rb_tipo")))
	s_loja_origem = Trim(Request.Form("c_loja"))
	s_loja_origem = normaliza_codigo(s_loja_origem, TAM_MIN_LOJA)
	s_op_descricao = Trim(Request.Form("op_selecionada_descricao"))
	s_pedido = Trim(Request.Form("c_pedido"))
	s_pedido_destino = normaliza_num_pedido(Trim(Request.Form("c_pedido")))
	c_obs_pecas_necessarias = Trim(Request.Form("c_obs_pecas_necessarias"))
	s_loja_destino = ""

	if InStr(s_tipo, "TRANSF_") > 0 then	
	'	TRANSFERÊNCIA ENTRE ESTOQUES
		v_aux = Split(s_tipo, "_")
		s_cod_estoque_origem = v_aux(Ubound(v_aux)-1)
		s_cod_estoque_destino = v_aux(Ubound(v_aux))
		s_fluxo = "TRANSF"
		if (s_cod_estoque_origem<>ID_ESTOQUE_SHOW_ROOM)And(s_cod_estoque_origem<>ID_ESTOQUE_DEVOLUCAO) then s_loja_origem=""
	else
	'	SAÍDA DO ESTOQUE DE VENDA
		s_cod_estoque_origem = ID_ESTOQUE_VENDA
		s_cod_estoque_destino = ID_ESTOQUE_DANIFICADOS
		s_fluxo = Left(s_tipo, 3)
		end if

	if alerta = "" then
		if s_cod_estoque_destino <> ID_ESTOQUE_DANIFICADOS then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Estoque de destino é inválido para esta operação."
			end if
		end if
		
	dim r_item
	set r_item = New cl_ITEM_PEDIDO

	s=retorna_so_digitos(Request.Form("c_fabricante"))
	if s = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Não foi informado o fabricante."
	else
		r_item.fabricante = normaliza_codigo(s, TAM_MIN_FABRICANTE)
		end if
		
	s=Trim(Request.Form("c_produto"))
	if s = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Não foi informado o produto."
	else
		r_item.produto=Ucase(Trim(s))
		end if
		
	s = Trim(Request.Form("c_qtde"))
	if IsNumeric(s) then r_item.qtde = CLng(s) else r_item.qtde = 0
	if r_item.qtde <> 1 then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Quantidade inválida: uma única unidade deve ser transferida por vez."
		end if
	
	dim v_OS_item
	redim v_OS_item(0)
	set v_OS_item(0) = New cl_ORDEM_SERVICO_ITEM
	n = Request.Form("c_descricao_volume").Count
	for i = 1 to n
		s=Trim(Request.Form("c_descricao_volume")(i))
		if s <> "" then
			if Trim(v_OS_item(ubound(v_OS_item)).descricao_volume) <> "" then
				redim preserve v_OS_item(ubound(v_OS_item)+1)
				set v_OS_item(ubound(v_OS_item)) = New cl_ORDEM_SERVICO_ITEM
				end if
			with v_OS_item(ubound(v_OS_item))
				.num_serie = Trim(Request.Form("c_num_serie")(i))
				.tipo = Trim(Request.Form("c_tipo")(i))
				.descricao_volume = Trim(Request.Form("c_descricao_volume")(i))
				.obs_problema = Trim(Request.Form("c_obs_problema")(i))
				end with
			end if
		next
	
	if converte_numero(s_id_nfe_emitente) = 0 then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Não foi informada a empresa."
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_pessimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim r_pedido, r_cliente
	
	if s_pedido_destino <> "" then
		if Not le_pedido(s_pedido_destino, r_pedido, msg_erro) then 
			alerta=texto_add_br(alerta)
			alerta = alerta & msg_erro
			end if

		set r_cliente = New cl_CLIENTE
		if Not x_cliente_bd(r_pedido.id_cliente, r_cliente) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Falha ao obter os dados do cliente (id=" & r_pedido.id_cliente & ")."
			end if

		if alerta = "" then
			if converte_numero(s_id_nfe_emitente) <> converte_numero(r_pedido.id_nfe_emitente) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O pedido " & r_pedido.pedido & " não está vinculado ao CD '" & obtem_apelido_empresa_NFe_emitente(s_id_nfe_emitente) & "'"
				end if
			end if
	else
		set r_pedido = new cl_PEDIDO
		set r_cliente = New cl_CLIENTE
		end if

	if alerta = "" then
		for i=Lbound(v_OS_item) to Ubound(v_OS_item)
			with v_OS_item(i)
				if Len(.obs_problema) > MAX_TAM_OS_OBS_PROBLEMA then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O campo 'Problema' do volume '" & .descricao_volume & "' excede o tamanho máximo de " & CStr(MAX_TAM_OS_OBS_PROBLEMA) & " caracteres."
					end if
				end with
			next
		
		if Len(c_obs_pecas_necessarias) > MAX_TAM_OS_OBS_PECAS_NECESSARIAS then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O campo 'Peças Necessárias' excede o tamanho máximo de " & Cstr(MAX_TAM_OS_OBS_PECAS_NECESSARIAS) & " caracteres."
			end if
		end if

	if alerta = "" then
		s = "SELECT * FROM t_PRODUTO WHERE (fabricante='" & r_item.fabricante & "') AND (produto='" & r_item.produto & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Produto " & r_item.produto & " não está cadastrado."
		else
			r_item.ean = Trim("" & rs("ean"))
			r_item.descricao = Trim("" & rs("descricao"))
			r_item.descricao_html = Trim("" & rs("descricao_html"))
			end if
		end if


'	ESTOQUE DE DEVOLUÇÃO
	Dim blnEstoqueDevolucao
	blnEstoqueDevolucao=False
	if alerta = "" then
		if (s_tipo="ENT_DEV") then blnEstoqueDevolucao=True
		if (s_tipo="TRANSF_DEV_DAN") then blnEstoqueDevolucao=True
		if (s_tipo="TRANSF_DEV_ROU") then blnEstoqueDevolucao=True
		if blnEstoqueDevolucao then
			if s_pedido = "" then
				alerta = "Número do pedido não informado."
			else
				s = "SELECT " & _
						"*" & _
					" FROM t_PEDIDO" & _
					" WHERE" & _
						" (pedido='" & s_pedido & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & s_pedido & " NÃO está cadastrado."
					end if
				end if
			
			if alerta = "" then
				if converte_numero(rs("loja")) <> converte_numero(s_loja_origem) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & s_pedido & " NÃO pertence à loja " & s_loja_origem & "."
					end if
				end if
			
			if alerta = "" then
				if converte_numero(s_id_nfe_emitente) <> converte_numero(rs("id_nfe_emitente")) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O pedido " & Trim("" & rs("pedido")) & " não está vinculado ao CD '" & obtem_apelido_empresa_NFe_emitente(s_id_nfe_emitente) & "'"
					end if
				end if

			if alerta = "" then
				with r_item
				'	QUANTIDADE DE PRODUTOS A SER TRANSFERIDA
				'	IMPORTANTE: NA TABELA T_ESTOQUE_MOVIMENTO, SOMENTE O ESTOQUE LÓGICO 'SPE' (SEM PRESENÇA NO ESTOQUE) NÃO POSSUI CONTEÚDO NO CAMPO 'id_estoque'.
					s = "SELECT" & _
							" SUM(qtde) AS total" & _
						" FROM t_ESTOQUE_MOVIMENTO tEM" & _
							" INNER JOIN t_ESTOQUE tE ON (tEM.id_estoque = tE.id_estoque)" & _
						" WHERE" & _
							" (id_nfe_emitente = " & s_id_nfe_emitente & ")" & _
							" AND (tEM.anulado_status=0)" & _
							" AND (tEM.fabricante='" & Trim(.fabricante) & "')" & _
							" AND (tEM.produto='" & Trim(.produto) & "')" & _
							" AND (tEM.estoque='" & ID_ESTOQUE_DEVOLUCAO & "')" & _
							" AND (tEM.loja='" & s_loja_origem & "')" & _
							" AND (tEM.pedido='" & s_pedido & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					j=0
					if Not rs.Eof then 
						if Not IsNull(rs("total")) then j = CLng(rs("total"))
						end if
					if .qtde > j then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & s_pedido & ": faltam " & CStr(.qtde-j) & " unidades do produto " & .produto & " do fabricante " & .fabricante & "."
						end if
					end with
				end if
			end if
		end if


'	VERIFICA DISPONIBILIDADE NO ESTOQUE
	if alerta = "" then
		with r_item
		'	IMPORTANTE: NA TABELA T_ESTOQUE_MOVIMENTO, SOMENTE O ESTOQUE LÓGICO 'SPE' (SEM PRESENÇA NO ESTOQUE) NÃO POSSUI CONTEÚDO NO CAMPO 'id_estoque'.
			if s_fluxo = "TRANSF" then
				s = "SELECT" & _
						" SUM(tEM.qtde) AS total" & _
					" FROM t_ESTOQUE_MOVIMENTO tEM" & _
						" INNER JOIN t_ESTOQUE tE ON (tEM.id_estoque = tE.id_estoque)" & _
					" WHERE" & _
						" (tEM.anulado_status=0)" & _
						" AND (tE.id_nfe_emitente = " & s_id_nfe_emitente & ")" & _
						" AND (tEM.fabricante='" & Trim(.fabricante) & "')" & _
						" AND (tEM.produto='" & Trim(.produto) & "')" & _
						" AND (tEM.estoque='" & s_cod_estoque_origem & "')"
				if (s_cod_estoque_origem=ID_ESTOQUE_SHOW_ROOM)Or(s_cod_estoque_origem=ID_ESTOQUE_DEVOLUCAO) then
					s = s & " AND (tEM.loja='" & s_loja_origem & "')"
					end if
			else
			'	QUANTIDADE DE PRODUTOS A SER TRANSFERIDA
				s = "SELECT" & _
						" SUM(qtde-qtde_utilizada) AS total" & _
					" FROM t_ESTOQUE tE" & _
						" INNER JOIN t_ESTOQUE_ITEM tEI ON (tE.id_estoque = tEI.id_estoque)" & _
					" WHERE" & _
						" (tE.id_nfe_emitente = " & s_id_nfe_emitente & ")" & _
						" AND (tEI.fabricante = '" & Trim(.fabricante) & "')" & _
						" AND (tEI.produto = '" & Trim(.produto) & "')" & _
						" AND ((qtde-qtde_utilizada) > 0)"
				end if
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			j=0
			if Not rs.Eof then 
				if Not IsNull(rs("total")) then j = CLng(rs("total"))
				end if
			if .qtde > j then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Faltam " & CStr(.qtde-j) & " unidades do produto " & .produto & " do fabricante " & .fabricante & " no CD '" & obtem_apelido_empresa_NFe_emitente(s_id_nfe_emitente) & "'"
				end if
			end with
		end if

	if alerta = "" then
		if (s_fluxo <> "SAI") And (s_fluxo <> "TRANSF") then
			alerta = "Operação de movimentação de estoque inválida."
			end if
		end if
		
	if alerta = "" then
	'	INFORMAÇÕES PARA O LOG (MOVIMENTAÇÃO DO ESTOQUE)
		s_log = ""
		with r_item
			if s_fluxo="TRANSF" then
				s_log = s_log & log_estoque_monta_transferencia(.qtde, .fabricante, .produto)
			else
				s_log = s_log & log_estoque_monta_decremento(.qtde, .fabricante, .produto)
				end if
			end with

		s = s_op_descricao & " (id_nfe_emitente = " & s_id_nfe_emitente & ")"
		s_log = s & ":" & s_log

		if Not gera_nsu(NSU_ORDEM_SERVICO, s_chave_OS, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
		
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		with r_item
			if s_fluxo="TRANSF" then
				s_chave_OS_origem = ""
				s_pedido_origem = ""
				if blnEstoqueDevolucao then s_pedido_origem = s_pedido
				if Not estoque_produto_transfere_entre_estoques_v2( _
									usuario, _
									s_id_nfe_emitente, _
									.fabricante, _
									.produto, _
									.qtde, _
									qtde_transferida, _
									s_cod_estoque_origem, _
									s_loja_origem, _
									s_chave_OS_origem, _
									s_pedido_origem, _
									ID_ESTOQUE_DANIFICADOS, _
									s_loja_destino, _
									s_chave_OS, _
									s_pedido_destino, _
									msg_erro _
									) then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
					end if
			else
				if Not estoque_produto_saida_por_transferencia_v2( _
									usuario, _
									ID_ESTOQUE_DANIFICADOS, _
									s_loja_destino, _
									s_id_nfe_emitente, _
									.fabricante, _
									.produto, _
									.qtde, _
									s_chave_OS, _
									s_pedido_destino, _
									msg_erro _
									) then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
					end if
				end if
			end with
		
	'	GRAVA LOG DA MOVIMENTAÇÃO NO ESTOQUE
		grava_log usuario, s_loja_origem, s_pedido, "", OP_LOG_ESTOQUE_TRANSFERENCIA, s_log

		' GRAVA A ORDEM DE SERVIÇO
		s = "SELECT * FROM t_ORDEM_SERVICO WHERE ordem_servico='X'"
		if rs.State <> 0 then rs.Close
		rs.Open s, cn
		rs.AddNew 
		rs("ordem_servico") = s_chave_OS
		rs("usuario") = usuario
		rs("data") = Date
		rs("hora") = retorna_so_digitos(formata_hora(Now))
		rs("situacao_status") = ST_OS_EM_ANDAMENTO
		rs("situacao_data") = Now
		rs("situacao_usuario") = usuario
		rs("fabricante") = r_item.fabricante
		rs("produto") = r_item.produto
		rs("qtde") = r_item.qtde
		rs("ean") = r_item.ean
		rs("descricao") = r_item.descricao
		rs("descricao_html") = r_item.descricao_html
		rs("obs_pecas_necessarias") = c_obs_pecas_necessarias
		if s_pedido_destino <> "" then 
			rs("pedido") = s_pedido_destino
			rs("nf") = r_pedido.obs_2
			rs("indicador") = r_pedido.indicador
			rs("id_cliente") = r_pedido.id_cliente
			rs("tipo_cliente") = r_cliente.tipo
			rs("nome_cliente") = r_cliente.nome
			rs("endereco") = r_cliente.endereco
			rs("endereco_numero") = r_cliente.endereco_numero
			rs("endereco_complemento") = r_cliente.endereco_complemento
			rs("bairro") = r_cliente.bairro
			rs("cidade") = r_cliente.cidade
			rs("uf") = r_cliente.uf
			rs("cep") = r_cliente.cep
			rs("ddd_res") = r_cliente.ddd_res
			rs("tel_res") = r_cliente.tel_res
			rs("ddd_com") = r_cliente.ddd_com
			rs("tel_com") = r_cliente.tel_com
			rs("ramal_com") = r_cliente.ramal_com
			rs("contato") = r_cliente.contato
			end if
		rs("cod_estoque_origem") = s_cod_estoque_origem
		rs("loja_estoque_origem") = s_loja_origem
		rs("id_nfe_emitente") = converte_numero(s_id_nfe_emitente)
		rs.Update
		
		if Err <> 0 then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			end if

		for i=Lbound(v_OS_item) to Ubound(v_OS_item)
			with v_OS_item(i)
				s="SELECT * FROM t_ORDEM_SERVICO_ITEM WHERE ordem_servico='X'"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				rs.AddNew 
				rs("ordem_servico") = s_chave_OS
				rs("num_serie") = .num_serie
				rs("tipo") = .tipo
				rs("descricao_volume") = .descricao_volume
				rs("obs_problema") = .obs_problema
				rs("sequencia") = renumera_com_base1(Lbound(v_OS_item), i)
				rs.Update
				
				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
				end with
			next

	'	MONTA DADOS P/ O LOG DA ORDEM DE SERVIÇO
		s_log = "Nova ordem de serviço: nº " & s_chave_OS & "; pedido=" & s_pedido_destino & "; id_nfe_emitente=" & s_id_nfe_emitente & "; loja=" & s_loja_origem & "; estoque_origem=" & s_cod_estoque_origem & "; fabricante=" & r_item.fabricante & "; produto=" & r_item.produto & "; qtde=" & Cstr(r_item.qtde) & "; nf=" & r_pedido.obs_2 & "; id_cliente=" & r_pedido.id_cliente & "; nome_cliente=" & r_cliente.nome & "; obs_pecas_necessarias=" & c_obs_pecas_necessarias
		for i=Lbound(v_OS_item) to Ubound(v_OS_item)
			with v_OS_item(i)
				if s_log <> "" then s_log = s_log & chr(13)
				s_aux = .tipo
				if s_aux = "" then s_aux = chr(34) & chr(34)
				s_log = s_log & "Volume " & Cstr(renumera_com_base1(Lbound(v_OS_item), i)) & ": num_serie=" & .num_serie & "; tipo=" & s_aux & "; descricao_volume=" & .descricao_volume & "; obs_problema=" & .obs_problema
				end with
			next
		
	'	GRAVA LOG DA NOVA ORDEM DE SERVIÇO
		grava_log usuario, "", "", "", OP_LOG_ORDEM_SERVICO_INCLUSAO, s_log

	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
		if Err=0 then 
			if url_back = "" then
				url_back_aux = "X"
			else
				url_back_aux = url_back
				end if
			s = "Ordem de serviço cadastrada:  nº <a href='OrdemServico.asp?num_OS=" & s_chave_OS & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "' title='Clique para consultar a Ordem de Serviço'><u>" & formata_num_OS_tela(s_chave_OS) & "</u></a>" & _
				"<br>" & _
				"Transferência do estoque concluída com sucesso: " & _
				Lcase(s_op_descricao)
			Session(SESSION_CLIPBOARD) = s
			Response.Redirect("mensagem.asp" & "?url_back=" & url_back_aux & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		else
			alerta=texto_add_br(alerta)
			alerta=alerta & Cstr(Err) & ": " & Err.Description
			end if
		end if


	if alerta <> "" then 
		alerta = texto_add_br(s_op_descricao) & alerta
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



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>



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