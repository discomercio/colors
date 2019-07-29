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
'	  E S T O Q U E T R A N S F E R E C O N F I R M A . A S P
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

	dim s, s_log, i, j, n, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_id_nfe_emitente
	dim s_tipo, s_loja, s_pedido, s_op_descricao, s_cod_estoque, s_fluxo, s_ckb_spe, s_ckb_spe_descricao, qtde_estornada
	dim v_aux, s_cod_estoque_origem, s_cod_estoque_destino, s_loja_origem, s_loja_destino, qtde_transferida
	dim s_chave_OS_origem
	dim s_pedido_origem
	dim s_chave_OS_destino
	dim s_pedido_destino

	dim v_item
	dim alerta
	alerta=""

'	OBTÉM DADOS DO FORMULÁRIO
	s_id_nfe_emitente = Trim(Request.Form("c_id_nfe_emitente"))
	s_tipo = Ucase(Trim(Request.Form("rb_tipo")))
	s_loja = Trim(Request.Form("c_loja"))
	s_loja = normaliza_codigo(s_loja, TAM_MIN_LOJA)
	s_pedido = Trim(Request.Form("c_pedido"))
	s_op_descricao = Trim(Request.Form("op_selecionada_descricao"))
	s_ckb_spe = Ucase(Trim(Request.Form("ckb_spe")))
	s_ckb_spe_descricao = Trim(Request.Form("ckb_spe_descricao"))

	s_cod_estoque = ""
	s_cod_estoque_origem = ""
	s_cod_estoque_destino = ""

	if InStr(s_tipo, "TRANSF_") > 0 then
		v_aux = Split(s_tipo, "_")
		s_cod_estoque_origem = v_aux(Ubound(v_aux)-1)
		s_cod_estoque_destino = v_aux(Ubound(v_aux))
		s_fluxo = "TRANSF"
	else
		s_cod_estoque = Right(s_tipo, 3)
		s_fluxo = Left(s_tipo, 3)
		end if
	
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
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	if s_fluxo = "" then
		alerta = "Operação selecionada é inválida."
		end if

	if s_fluxo = "TRANSF" then
		if (s_cod_estoque_origem="") Or (s_cod_estoque_destino="") then
			alerta = "Operação selecionada é inválida."
			end if
	else
		if (s_fluxo<>"SAI") And (s_fluxo<>"ENT") then
			alerta = "Operação selecionada é inválida."
		elseif (s_cod_estoque="") then
			alerta = "Operação selecionada é inválida."
			end if
		end if
	
	if alerta = "" then
		if converte_numero(s_id_nfe_emitente) = 0 then
			alerta = "É necessário selecionar a empresa (CD)."
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
				if converte_numero(rs("loja")) <> converte_numero(s_loja) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & s_pedido & " NÃO pertence à loja " & s_loja & "."
					end if
				if converte_numero(rs("id_nfe_emitente")) <> converte_numero(s_id_nfe_emitente) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & s_pedido & " NÃO está vinculado ao CD '" & obtem_apelido_empresa_NFe_emitente(s_id_nfe_emitente) & "'"
					end if
				end if
			
			if alerta = "" then
				for i=Lbound(v_item) to Ubound(v_item)
					with v_item(i)
					'	QUANTIDADE DE PRODUTOS A SER TRANSFERIDA
						s = "SELECT" & _
								" SUM(qtde) AS total" & _
							" FROM t_ESTOQUE_MOVIMENTO" & _
							" WHERE" & _
								" (anulado_status=0)" & _
								" AND (fabricante='" & Trim(.fabricante) & "')" & _
								" AND (produto='" & Trim(.produto) & "')" & _
								" AND (estoque='" & ID_ESTOQUE_DEVOLUCAO & "')" & _
								" AND (loja='" & s_loja & "')" & _
								" AND (pedido='" & s_pedido & "')"
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
					next
				end if
			end if
		end if

	
'	VERIFICA DISPONIBILIDADE NO ESTOQUE
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
			'	QUANTIDADE DE PRODUTOS A SER TRANSFERIDA
			'	OBS: RESSALTANDO QUE, NA TABELA T_ESTOQUE_MOVIMENTO, SOMENTE O ESTOQUE LÓGICO 'SPE' (SEM PRESENÇA NO ESTOQUE) NÃO POSSUI CONTEÚDO NO CAMPO 'id_estoque'.
				s = ""
				if s_fluxo="SAI" then
					s = "SELECT" & _
							" SUM(qtde-qtde_utilizada) AS total" & _
						" FROM t_ESTOQUE tE" & _
							" INNER JOIN t_ESTOQUE_ITEM tEI ON (tE.id_estoque = tEI.id_estoque)" & _
						" WHERE" & _
							" (tE.id_nfe_emitente = " & s_id_nfe_emitente & ")" & _
							" AND (tEI.fabricante = '" & Trim(.fabricante) & "')" & _
							" AND (tEI.produto = '" & Trim(.produto) & "')" & _
							" AND ((qtde-qtde_utilizada) > 0)"
				elseif s_fluxo="ENT" then
					s = "SELECT" & _
							" SUM(qtde) AS total" & _
						" FROM t_ESTOQUE_MOVIMENTO tEM" & _
							" INNER JOIN t_ESTOQUE tE ON (tEM.id_estoque = tE.id_estoque)" & _
						" WHERE" & _
							" (tEM.anulado_status=0)" & _
							" AND (tE.id_nfe_emitente = " & s_id_nfe_emitente & ")" & _
							" AND (tEM.fabricante='" & Trim(.fabricante) & "')" & _
							" AND (tEM.produto='" & Trim(.produto) & "')" & _
							" AND (tEM.estoque='" & s_cod_estoque & "')"
					if (s_cod_estoque=ID_ESTOQUE_SHOW_ROOM)Or(s_cod_estoque=ID_ESTOQUE_DEVOLUCAO) then
						s = s & " AND (tEM.loja='" & s_loja & "')"
						end if
				elseif s_fluxo="TRANSF" then
					s = "SELECT" & _
							" SUM(qtde) AS total" & _
						" FROM t_ESTOQUE_MOVIMENTO tEM" & _
							" INNER JOIN t_ESTOQUE tE ON (tEM.id_estoque = tE.id_estoque)" & _
						" WHERE" & _
							" (tEM.anulado_status=0)" & _
							" AND (tE.id_nfe_emitente = " & s_id_nfe_emitente & ")" & _
							" AND (tEM.fabricante='" & Trim(.fabricante) & "')" & _
							" AND (tEM.produto='" & Trim(.produto) & "')" & _
							" AND (tEM.estoque='" & s_cod_estoque_origem & "')"
					if (s_cod_estoque_origem=ID_ESTOQUE_SHOW_ROOM)Or(s_cod_estoque_origem=ID_ESTOQUE_DEVOLUCAO) then
						s = s & " AND (tEM.loja='" & s_loja & "')"
						end if
					end if
				
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				j=0
				if Not rs.Eof then 
					if Not IsNull(rs("total")) then j = CLng(rs("total"))
					end if
				if .qtde > j then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Faltam " & CStr(.qtde-j) & " unidades do produto " & .produto & " do fabricante " & .fabricante & " (CD '" & obtem_apelido_empresa_NFe_emitente(s_id_nfe_emitente) & "')"
					end if
				end with
			next
		end if

	
	if alerta = "" then
	'	INFORMAÇÕES PARA O LOG
		s_log = ""
		for i = Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .produto <> "" then
					if s_fluxo="SAI" then
						s_log = s_log & log_estoque_monta_decremento(.qtde, .fabricante, .produto)
					elseif s_fluxo="ENT" then
						s_log = s_log & log_estoque_monta_incremento(.qtde, .fabricante, .produto)
					elseif s_fluxo="TRANSF" then
						s_log = s_log & log_estoque_monta_transferencia(.qtde, .fabricante, .produto)
						end if
					end if
				end with
			next

		s = s_op_descricao
		if s_ckb_spe<>"" then s = s & " (" & Lcase(s_ckb_spe_descricao) & ")"
		s = s & " (id_nfe_emitente=" & s_id_nfe_emitente & ")"
		s_log = s & ":" & s_log
		
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if s_fluxo="SAI" then
					if Not estoque_produto_saida_por_transferencia_v2(usuario, s_cod_estoque, s_loja, s_id_nfe_emitente, .fabricante, .produto, .qtde, "", "", msg_erro) then
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
						end if
				elseif s_fluxo="ENT" then
					s_pedido_origem = ""
					if blnEstoqueDevolucao then s_pedido_origem = s_pedido
					if Not estoque_produto_estorna_por_transferencia_v2(usuario, s_cod_estoque, s_loja, s_pedido_origem, s_id_nfe_emitente, .fabricante, .produto, .qtde, qtde_estornada, "", msg_erro) then
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
						end if
				elseif s_fluxo="TRANSF" then
					'Esta lógica de obter os valores para a loja do estoque de origem e do estoque de destino
					'funciona apenas porque entre as opções de "transferência entre estoques" não há nenhuma
					'combinação em que ambos os estoques (origem e destino) exijam a loja.
					s_loja_origem = ""
					s_loja_destino = ""
					if (s_cod_estoque_origem = ID_ESTOQUE_SHOW_ROOM) Or (s_cod_estoque_origem = ID_ESTOQUE_DEVOLUCAO) then s_loja_origem = s_loja
					if (s_cod_estoque_destino = ID_ESTOQUE_SHOW_ROOM) Or (s_cod_estoque_destino = ID_ESTOQUE_DEVOLUCAO) then s_loja_destino = s_loja
					s_chave_OS_origem = ""
					s_pedido_origem = ""
					if blnEstoqueDevolucao then s_pedido_origem = s_pedido
					s_chave_OS_destino = ""
					s_pedido_destino = ""
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
									s_cod_estoque_destino, _
									s_loja_destino, _
									s_chave_OS_destino, _
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
			next
		
		grava_log usuario, s_loja, s_pedido, "", OP_LOG_ESTOQUE_TRANSFERENCIA, s_log

	'	PROCESSA OS PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE?
		if (s_fluxo="ENT") And (s_ckb_spe="") then
			if Not estoque_processa_produtos_vendidos_sem_presenca_v2(s_id_nfe_emitente, usuario, msg_erro) then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
				end if
			end if

	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
		if Err=0 then 
			s = "Transferência/movimentação do estoque concluída com sucesso: " & _
				Lcase(s_op_descricao)
			Session(SESSION_CLIPBOARD) = s
			Response.Redirect("mensagem.asp" & "?url_back=EstoqueTransfere.asp&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		else
			alerta=Cstr(Err) & ": " & Err.Description
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