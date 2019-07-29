<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ====================================================================
'	  EstoqueTransferePedidoConfirma.asp
'     ====================================================================
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

	dim s, s_log, strMsgResultado, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim c_pedido_origem, c_pedido_destino
	dim flag_ok, n
	dim alerta
	alerta=""

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	dim intRecordsAffected
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	OBTÉM DADOS DO FORMULÁRIO
	c_pedido_origem = Ucase(Trim(Request.Form("c_pedido_origem")))
	s = normaliza_num_pedido(c_pedido_origem)
	if s <> "" then c_pedido_origem = s

	c_pedido_destino = Ucase(Trim(Request.Form("c_pedido_destino")))
	s = normaliza_num_pedido(c_pedido_destino)
	if s <> "" then c_pedido_destino = s
	
	dim intCounter, intCounterAux, intQtdeItens
	dim v_item
	redim v_item(0)
	set v_item(0) = New cl_ESTOQUE_TRANSFERE_PEDIDO
	intQtdeItens = Request.Form("c_produto").Count
	for intCounter = 1 to intQtdeItens
		s = Trim(Request.Form("c_produto")(intCounter))
		if s <> "" then
			if Trim(v_item(Ubound(v_item)).produto) <> "" then
				redim preserve v_item(Ubound(v_item)+1)
				set v_item(Ubound(v_item)) = New cl_ESTOQUE_TRANSFERE_PEDIDO
				end if
			with v_item(Ubound(v_item))
				.produto = UCase(Trim(Request.Form("c_produto")(intCounter)))
				s = retorna_so_digitos(Request.Form("c_fabricante")(intCounter))
				.fabricante = normaliza_codigo(s, TAM_MIN_FABRICANTE)
				s = Trim(Request.Form("c_qtde")(intCounter))
				.qtde = converte_numero(s)
				end with
			end if
		next
	
'	CONSISTE CAMPOS
	if c_pedido_origem = "" then
		alerta = "Especifique o número do pedido que irá ceder as mercadorias disponíveis."
	elseif c_pedido_destino = "" then
		alerta = "Especifique o número do pedido que irá receber as mercadorias."
	elseif c_pedido_origem = c_pedido_destino then
		alerta = "Pedido de origem e de destino devem ser diferentes."
		end if

'	VERIFICA SE HÁ PRODUTOS REPETIDOS
	if alerta = "" then
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
				for intCounterAux=Lbound(v_item) to (intCounter-1)
					if (.produto = v_item(intCounterAux).produto) And (.fabricante = v_item(intCounterAux).fabricante) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": linha " & renumera_com_base1(Lbound(v_item),intCounter) & " repete o mesmo produto da linha " & renumera_com_base1(Lbound(v_item),intCounterAux) & "."
						exit for
						end if
					next
				end with
			next
		end if

'	VERIFICA SE HÁ ALGUM ITEM E SE OS DADOS ESTÃO PREENCHIDOS CORRETAMENTE
	dim blnTemItem
	dim blnConsistirLinha
	blnTemItem = False
	for intCounter = Lbound(v_item) to Ubound(v_item)
		if alerta = "" then
			with v_item(intCounter)
				blnConsistirLinha = False
				if Trim(.fabricante) <> "" then 
					blnTemItem = True
					blnConsistirLinha = True
				elseif Trim(.produto) <> "" then
					blnTemItem = True
					blnConsistirLinha = True
				elseif CLng(.qtde) > 0 then
					blnTemItem = True
					blnConsistirLinha = True
					end if

				if blnConsistirLinha then
					if .fabricante = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Linha " & renumera_com_base1(Lbound(v_item),intCounter) & ": não foi especificado o código do fabricante."
					elseif .produto = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Linha " & renumera_com_base1(Lbound(v_item),intCounter) & ": não foi especificado o código do produto."
					elseif CLng(.qtde) <= 0 then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Linha " & renumera_com_base1(Lbound(v_item),intCounter) & ": quantidade a transferir é inválida."
						end if
					end if
				end with
			end if
		next

	if alerta = "" then
		if Not blnTemItem then
			alerta = "Nenhum produto foi especificado para transferência."
			end if
		end if
		
'	CONSISTE PRODUTO
	if alerta = "" then
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
				if Trim(.produto) <> "" then
					if (Not IsEAN(.produto)) And (.fabricante="") then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Linha " & renumera_com_base1(Lbound(v_item),intCounter) & ": não foi especificado o fabricante do produto."
					else
						s = "SELECT " & _
								"*" & _
							" FROM t_PRODUTO" & _
							" WHERE"
						if IsEAN(.produto) then
							s = s & " (ean='" & .produto & "')"
						else
							s = s & " (fabricante='" & .fabricante & "') AND (produto='" & .produto & "')"
							end if
		
						if rs.State <> 0 then rs.Close
						rs.open s, cn
						if rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto " & .produto & " NÃO está cadastrado."
						else
							flag_ok = True
							if IsEAN(.produto) And (.fabricante<>"") then
								if (.fabricante<>Trim("" & rs("fabricante"))) then
									flag_ok = False
									alerta=texto_add_br(alerta)
									alerta=alerta & "Produto " & .produto & " NÃO pertence ao fabricante " & .fabricante & "."
									end if
								end if
							if flag_ok then
							'   CARREGA CÓDIGO INTERNO DO PRODUTO
								.fabricante = Trim("" & rs("fabricante"))
								.produto = Trim("" & rs("produto"))
								.produto_descricao = Trim("" & rs("descricao"))
								.produto_descricao_html = Trim("" & rs("descricao_html"))
								end if
							end if
						end if
					end if
				end with
			next
		end if
	
'	CONSISTE FABRICANTE
	if alerta = "" then
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
				if Trim(.produto) <> "" then
					s = "SELECT" & _
							" nome," & _
							" razao_social" & _
						" FROM t_FABRICANTE" & _
						" WHERE" & _
							" (fabricante='" & .fabricante & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					.nome_fabricante = ""
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Fabricante " & .fabricante & " não está cadastrado."
					else
						.nome_fabricante = Trim("" & rs("razao_social"))
						if .nome_fabricante = "" then .nome_fabricante = Trim("" & rs("nome"))
						.nome_fabricante = iniciais_em_maiusculas(.nome_fabricante)
						end if
					end if
				end with
			next
		end if

'	CONSISTE PEDIDOS
	dim id_nfe_emitente_pedido_origem, id_nfe_emitente_pedido_destino
	id_nfe_emitente_pedido_origem = 0
	id_nfe_emitente_pedido_destino = 0
	if alerta = "" then
		s = "SELECT" & _
				" pedido, id_nfe_emitente" & _
			" FROM t_PEDIDO" & _
			" WHERE" & _
				" (pedido='" & c_pedido_origem & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Pedido " & c_pedido_origem & " não está cadastrado."
		else
			id_nfe_emitente_pedido_origem = CLng(rs("id_nfe_emitente"))
			end if

		s = "SELECT" & _
				" pedido, id_nfe_emitente" & _
			" FROM t_PEDIDO" & _
			" WHERE" & _
				" (pedido='" & c_pedido_destino & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Pedido " & c_pedido_destino & " não está cadastrado."
		else
			id_nfe_emitente_pedido_destino = CLng(rs("id_nfe_emitente"))
			end if
		end if

'	SE O PEDIDO DE DESTINO JÁ ESTÁ CONSUMINDO ESTOQUE DE UM DETERMINADO CD, CONSISTE SE O PEDIDO DE ORIGEM É COMPATÍVEL
	if id_nfe_emitente_pedido_destino <> 0 then
		if id_nfe_emitente_pedido_destino <> id_nfe_emitente_pedido_origem then
			alerta=texto_add_br(alerta)
			alerta=alerta & "A operação não pode ser realizada porque os pedidos estão associados a estoques de empresas diferentes:" & _
							"<br />Pedido de origem (" & c_pedido_origem & "): " & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_pedido_origem) & _
							"<br />Pedido de destino (" & c_pedido_destino & "): " & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_pedido_destino)
			end if
		end if

'	CONSISTE ITEM DO PEDIDO (ORIGEM)
	if alerta = "" then
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
				if Trim(.produto) <> "" then
					s = "SELECT" & _
							" pedido," & _
							" fabricante," & _
							" produto," & _
							" qtde" & _
						" FROM t_PEDIDO_ITEM" & _
						" WHERE" & _
							" (pedido='" & c_pedido_origem & "')" & _
							" AND (fabricante='" & .fabricante & "')" & _
							" AND (produto='" & .produto & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido de origem " & c_pedido_origem & " não possui o produto " & .produto & " do fabricante " & .fabricante & "."
					else
						if CLng(rs("qtde")) < CLng(.qtde) then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Não é possível transferir " & formata_inteiro(.qtde) & _
											" unidades do produto " & .produto & " porque o pedido de origem " & _
											c_pedido_origem & " possui apenas " & formata_inteiro(rs("qtde")) & " unidades."
							end if
						end if
					end if
				end with
			next
		end if

'	CONSISTE ITEM DO PEDIDO (DESTINO)
	if alerta = "" then
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
				if Trim(.produto) <> "" then
					s = "SELECT" & _
							" pedido," & _
							" fabricante," & _
							" produto," & _
							" qtde" & _
						" FROM t_PEDIDO_ITEM" & _
						" WHERE" & _
							" (pedido='" & c_pedido_destino & "')" & _
							" AND (fabricante='" & .fabricante & "')" & _
							" AND (produto='" & .produto & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido de destino " & c_pedido_destino & " não possui o produto " & .produto & " do fabricante " & .fabricante & "."
					else
						if CLng(rs("qtde")) < CLng(.qtde) then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Não é possível transferir " & formata_inteiro(.qtde) & _
											" unidades do produto " & .produto & " porque o pedido de destino " & _
											c_pedido_destino & " indica apenas " & formata_inteiro(rs("qtde")) & " unidades."
							end if
						end if
					end if
				end with
			next
		end if

'	CONSISTE ESTOQUE (PEDIDO ORIGEM)
	if alerta = "" then
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
				if Trim(.produto) <> "" then
				' 	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
				'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
				'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
					s = "SELECT" & _
							" Sum(qtde) AS total" & _
						" FROM t_ESTOQUE_MOVIMENTO" & _
						" WHERE" & _
							" (anulado_status=0)" & _
							" AND (pedido='" & c_pedido_origem & "')" & _
							" AND (fabricante='" & .fabricante & "')" & _
							" AND (produto='" & .produto & "')" & _
							" AND (estoque='" & ID_ESTOQUE_VENDIDO & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					n = 0
					if Not rs.Eof then
						if Not IsNull(rs("total")) then n = CLng(rs("total"))
						end if
				'	NÃO DISPÕE DE QUANTIDADE SUFICIENTE
					if CLng(n) < CLng(.qtde) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Não é possível transferir " & formata_inteiro(.qtde) & _
										" unidades porque o pedido " & c_pedido_origem & _
										" dispõe de apenas " & formata_inteiro(n) & " unidades."
						end if
					end if
				end with
			next
		end if

'	CONSISTE ESTOQUE (PEDIDO DESTINO)
	if alerta = "" then
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
				if Trim(.produto) <> "" then
				' 	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
				'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
				'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
					s = "SELECT" & _
							" Sum(qtde) AS total" & _
						" FROM t_ESTOQUE_MOVIMENTO" & _
						" WHERE" & _
							" (anulado_status=0)" & _
							" AND (pedido='" & c_pedido_destino & "')" & _
							" AND (fabricante='" & .fabricante & "')" & _
							" AND (produto='" & .produto & "')" & _
							" AND (estoque='" & ID_ESTOQUE_SEM_PRESENCA & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					n = 0
					if Not rs.Eof then
						if Not IsNull(rs("total")) then n = CLng(rs("total"))
						end if
				'	QUANTIDADE A TRANSFERIR EXCEDE O NECESSÁRIO
					if CLng(n) < CLng(.qtde) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Não é possível transferir " & formata_inteiro(.qtde) & _
										" unidades porque o pedido " & c_pedido_destino & _
										" aguarda apenas " & formata_inteiro(n) & " unidades."
						end if
					end if
				end with
			next
		end if

	if alerta = "" then
	'	INFORMAÇÕES PARA O LOG
		s_log = "Transferência no estoque de Produtos Vendidos do pedido " & c_pedido_origem & _
				" para o pedido " & c_pedido_destino & ":"
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
				if Trim(.produto) <> "" then
					s_log = s_log & " " & formata_inteiro(.qtde) & "x" & Trim(.produto) & "(" & Trim(.fabricante) & ")"
					end if
				end with
			next
			
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
	'	SE O PEDIDO DE DESTINO NÃO ESTÁ ASSOCIADO A NENHUM CD, REGISTRA ESSA ASSOCIAÇÃO P/ O MESMO CD EM QUE ESTÁ O PRODUTO QUE ELE ESTÁ RECEBENDO
		if id_nfe_emitente_pedido_destino = 0 then
			s = "UPDATE t_PEDIDO SET id_nfe_emitente = " & Cstr(id_nfe_emitente_pedido_origem) & " WHERE (pedido = '" & c_pedido_destino & "') AND (id_nfe_emitente = 0)"
			cn.Execute s, intRecordsAffected
			if intRecordsAffected <> 1 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Falha ao tentar associar o pedido de destino (" & c_pedido_destino & ") ao estoque da empresa '" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_pedido_origem) & "'"
			else
				if s_log <> "" then s_log = s_log & "; "
				s_log = s_log & "Pedido de destino (" & c_pedido_destino & ") associado automaticamente ao estoque da empresa " & Cstr(id_nfe_emitente_pedido_origem) & " - " & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_pedido_origem)
				end if
			end if

		if alerta = "" then
			for intCounter = Lbound(v_item) to Ubound(v_item)
				with v_item(intCounter)
					if Trim(.produto) <> "" then
					'	EXECUTA A TRANSFERÊNCIA
						if Not estoque_transfere_produto_vendido_entre_pedidos_v2(usuario, c_pedido_origem, c_pedido_destino, .fabricante, .produto, .qtde, msg_erro) then
						'	~~~~~~~~~~~~~~~~
							cn.RollbackTrans
						'	~~~~~~~~~~~~~~~~
							if s_log <> "" then s_log = "FALHA: " & s_log & ";" & chr(13)
							s_log = s_log & msg_erro
							grava_log usuario, "", "", "", OP_LOG_ESTOQUE_TRANSF_PEDIDO, s_log
							Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
							end if
						end if
					end with
				next
			end if

		if alerta = "" then
			grava_log usuario, "", "", "", OP_LOG_ESTOQUE_TRANSF_PEDIDO, s_log
			end if

		if alerta <> "" then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
		else
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				strMsgResultado = _
					"Transferência concluída com sucesso!!" & chr(13) & _
					"<br>" & chr(13) & _
					"<br>" & chr(13) & _
					"<table CellPadding=0 CellSpacing=0>" & chr(13) & _
					"	<tr>" & chr(13) & _
					"		<td align='right' NOWRAP valign='top' class='Np'>Pedido origem:&nbsp;</td>" & chr(13) & _
					"		<td valign='top' class='Np'>" & c_pedido_origem & "</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"<tr>" & chr(13) & _
					"		<td align='right' NOWRAP valign='top' class='Np'>Pedido destino:&nbsp;</td>" & chr(13) & _
					"		<td valign='top' class='Np'>" & c_pedido_destino & "</td>" & chr(13) & _
					"</tr>" & chr(13)
				
				for intCounter = Lbound(v_item) to Ubound(v_item)
					with v_item(intCounter)
						if Trim(.produto) <> "" then
							strMsgResultado = strMsgResultado & _
								"	<tr>" & chr(13) & _
								"		<td align='right' NOWRAP valign='bottom' class='Np'>Quantidade:&nbsp;</td>" & chr(13) & _
								"		<td valign='bottom' class='Np'>" & formata_inteiro(.qtde) & " unidades do produto " & .produto & " - " & produto_formata_descricao_em_html(.produto_descricao_html) & "</td>" & chr(13) & _
								"	</tr>" & chr(13)
							end if
						end with
					next
			
				strMsgResultado = strMsgResultado & _
					"</table>" & chr(13) & _
					"<br>" & chr(13)
			
				Session(SESSION_CLIPBOARD) = strMsgResultado
				Response.Redirect("mensagem.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
			end if
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