<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  O R D E M S E R V I C O E N C E R R A C O N F I R M A . A S P
'     =============================================================
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

	dim s, s_log, s_num_OS, s_chave_OS, j, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim url_back, url_back_aux
	dim s_ckb_spe, s_ckb_spe_descricao
	dim s_tipo, s_op_descricao, qtde_transferida, qtde_estornada
	dim s_loja_destino
	dim v_aux, s_cod_estoque_origem, s_cod_estoque_destino, s_fluxo
	dim s_loja_origem
	dim s_pedido_origem
	dim s_chave_OS_destino
	dim s_pedido_destino

'	OBTÉM DADOS DO FORMULÁRIO
	url_back = Trim(Request("url_back"))
	s_op_descricao = Trim(Request.Form("op_selecionada_descricao"))
	s_num_OS = Trim(Request.Form("c_num_OS"))
	s_chave_OS = normaliza_codigo(retorna_so_digitos(s_num_OS), TAM_MAX_NSU)
	s_ckb_spe = Ucase(Trim(Request.Form("ckb_spe")))
	s_ckb_spe_descricao = Trim(Request.Form("ckb_spe_descricao"))
	s_loja_destino = Trim(Request.Form("c_loja"))
	s_loja_destino = normaliza_codigo(s_loja_destino, TAM_MIN_LOJA)
	s_tipo = Ucase(Trim(Request.Form("rb_tipo")))
	s_pedido_destino = Ucase(Trim(Request.Form("c_pedido")))
	if s_pedido_destino <> "" then s_pedido_destino = normaliza_num_pedido(s_pedido_destino)

	dim alerta
	alerta=""

	if InStr(s_tipo, "TRANSF_") > 0 then
		v_aux = Split(s_tipo, "_")
		s_cod_estoque_origem = v_aux(Ubound(v_aux)-1)
		s_cod_estoque_destino = v_aux(Ubound(v_aux))
		s_fluxo = "TRANSF"
		if (s_cod_estoque_destino<>ID_ESTOQUE_SHOW_ROOM)And(s_cod_estoque_destino<>ID_ESTOQUE_DEVOLUCAO) then s_loja_destino=""
		if s_cod_estoque_destino<>ID_ESTOQUE_DEVOLUCAO then s_pedido_destino = ""
	elseif InStr(s_tipo, "ENT_") > 0 then
		s_cod_estoque_origem = ID_ESTOQUE_DANIFICADOS
		s_cod_estoque_destino = ID_ESTOQUE_VENDA
		s_fluxo = Left(s_tipo, 3)
		s_loja_destino=""
		s_pedido_destino = ""
	else
		alerta = "Operação desconhecida."
		end if

	if alerta = "" then
		if s_chave_OS = "" then alerta = "Nº da ordem de serviço não foi informado."
		end if
	
	if alerta = "" then
		if (s_cod_estoque_destino=ID_ESTOQUE_SHOW_ROOM)Or(s_cod_estoque_destino=ID_ESTOQUE_DEVOLUCAO) then
			if s_loja_destino = "" then
				alerta = "Loja não foi informada."
				end if
			end if
		
		if s_cod_estoque_destino=ID_ESTOQUE_DEVOLUCAO then
			if s_pedido_destino = "" then
				alerta = "Pedido não foi informado."
				end if
			end if
		end if
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_pessimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim r_OS, r_OS_item
	
	if alerta = "" then
		if Not le_ordem_servico(s_chave_OS, r_OS, msg_erro) then 
			alerta = msg_erro
		else
			if Not le_ordem_servico_item(s_chave_OS, r_OS_item, msg_erro) then alerta = msg_erro
			end if
		end if

	if alerta = "" then
		if s_pedido_destino <> "" then
			s = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" & s_pedido_destino & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if Not rs.Eof then
				if converte_numero(rs("id_nfe_emitente")) <> converte_numero(r_OS.id_nfe_emitente) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "A ordem de serviço " & s_chave_OS & " e o pedido " & s_pedido_destino & " estão vinculadas a CD's diferentes: '" & obtem_apelido_empresa_NFe_emitente(r_OS.id_nfe_emitente) & "' e '" & obtem_apelido_empresa_NFe_emitente(rs("id_nfe_emitente")) & "', respectivamente."
					end if
				end if
			end if
		end if

'	VERIFICA DISPONIBILIDADE NO ESTOQUE
'	IMPORTANTE: NA TABELA T_ESTOQUE_MOVIMENTO, SOMENTE O ESTOQUE LÓGICO 'SPE' (SEM PRESENÇA NO ESTOQUE) NÃO POSSUI CONTEÚDO NO CAMPO 'id_estoque'.
	if alerta = "" then
		with r_OS
		'	QUANTIDADE DE PRODUTOS A SER TRANSFERIDA
			s = "SELECT" & _
					" SUM(qtde) AS total" & _
				" FROM t_ESTOQUE_MOVIMENTO tEM" & _
					" INNER JOIN t_ESTOQUE tE ON (tEM.id_estoque = tE.id_estoque)" & _
				" WHERE" & _
					" (tEM.anulado_status=0)" & _
					" AND (tE.id_nfe_emitente = " & .id_nfe_emitente & ")" & _
					" AND (tEM.fabricante='" & Trim(.fabricante) & "')" & _
					" AND (tEM.produto='" & Trim(.produto) & "')" & _
					" AND (tEM.estoque='" & ID_ESTOQUE_DANIFICADOS & "')" & _
					" AND (tEM.id_ordem_servico='" & s_chave_OS & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			j=0
			if Not rs.Eof then 
				if Not IsNull(rs("total")) then j = CLng(rs("total"))
				end if
			if .qtde > j then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Faltam " & CStr(.qtde-j) & " unidades do produto " & .produto & " do fabricante " & .fabricante & " (CD: '" & obtem_apelido_empresa_NFe_emitente(.id_nfe_emitente) & "')"
				end if
			end with
		end if

	if alerta = "" then
		if (s_fluxo <> "ENT") And (s_fluxo <> "TRANSF") then
			alerta = "Operação de movimentação de estoque inválida."
			end if
		end if
		
	if alerta = "" then
	'	INFORMAÇÕES PARA O LOG (MOVIMENTAÇÃO DO ESTOQUE)
		s_log = ""
		with r_OS
			if s_fluxo="ENT" then
				s_log = s_log & log_estoque_monta_incremento(.qtde, .fabricante, .produto)
			elseif s_fluxo="TRANSF" then
				s_log = s_log & log_estoque_monta_transferencia(.qtde, .fabricante, .produto)
				end if
			end with

		s = s_op_descricao
		if s_ckb_spe<>"" then s = s & " (" & Lcase(s_ckb_spe_descricao) & ")"
		s = s & " (id_nfe_emitente=" & r_OS.id_nfe_emitente & ", ordem de serviço nº " & formata_num_OS_tela(s_chave_OS) & ")"
		s_log = s & ":" & s_log

	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		with r_OS
			s_loja_origem = ""
			s_pedido_origem = ""
			s_chave_OS_destino = ""
			
			if s_fluxo="ENT" then
				if Not estoque_produto_estorna_por_transferencia_v2( _
									usuario, _
									ID_ESTOQUE_DANIFICADOS, _
									s_loja_origem, _
									s_pedido_origem, _
									.id_nfe_emitente, _
									.fabricante, _
									.produto, _
									.qtde, _
									qtde_estornada, _
									s_chave_OS, _
									msg_erro _
									) then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
					end if
			elseif s_fluxo="TRANSF" then
				if Not estoque_produto_transfere_entre_estoques_v2( _
									usuario, _
									.id_nfe_emitente, _
									.fabricante, _
									.produto, _
									.qtde, _
									qtde_transferida, _
									ID_ESTOQUE_DANIFICADOS, _
									s_loja_origem, _
									s_chave_OS, _
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
		
	'	GRAVA LOG DA MOVIMENTAÇÃO NO ESTOQUE
		grava_log usuario, "", "", "", OP_LOG_ESTOQUE_TRANSFERENCIA, s_log

	'	PROCESSA OS PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE?
		if (s_fluxo="ENT") And (s_ckb_spe="") then
			if Not estoque_processa_produtos_vendidos_sem_presenca_v2(r_OS.id_nfe_emitente, usuario, msg_erro) then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
				end if
			end if

		' ENCERRA A ORDEM DE SERVIÇO
		s = "SELECT * FROM t_ORDEM_SERVICO WHERE ordem_servico='" & s_chave_OS & "'"
		if rs.State <> 0 then rs.Close
		rs.Open s, cn
		if rs.Eof then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_ORDEM_SERVICO_NAO_CADASTRADA)
		else
			rs("situacao_status") = ST_OS_ENCERRADA
			rs("situacao_data") = Now
			rs("situacao_usuario") = usuario
			rs("cod_estoque_destino") = s_cod_estoque_destino
			rs("loja_estoque_destino") = s_loja_destino
			rs("pedido_destino") = s_pedido_destino
			rs.Update
			end if
			
		if Err <> 0 then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			end if

	'	MONTA DADOS P/ O LOG DA ORDEM DE SERVIÇO
		s_log = "Encerra ordem de serviço: nº " & s_chave_OS
		
	'	GRAVA LOG DA ORDEM DE SERVIÇO
		grava_log usuario, "", "", "", OP_LOG_ORDEM_SERVICO_ENCERRA, s_log

	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
		if Err=0 then 
			if url_back = "" then
				url_back_aux = "X"
			else
				url_back_aux = url_back
				end if
			s = "Ordem de serviço encerrada:  nº <a href='OrdemServico.asp?num_OS=" & s_chave_OS & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "' title='Clique para consultar a Ordem de Serviço'><u>" & formata_num_OS_tela(s_chave_OS) & "</u></a>" & _
				"<br>" & _
				"Transferência do estoque concluída com sucesso: " & _
				Lcase(s_op_descricao)
			Session(SESSION_CLIPBOARD) = s
			Response.Redirect("mensagem.asp" & "?url_back=" & url_back_aux & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
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