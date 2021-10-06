<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/global.asp" -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================================
'	  PedidoBlocoNotasItemDevolvidoNovoConfirma.asp
'     ===============================================================
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

	dim s, msg_erro
	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim alerta
	alerta=""
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_LJA_BLOCO_NOTAS_ITEM_DEVOLVIDO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if
	
	dim pedido_selecionado
	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim id_item_devolvido
	id_item_devolvido = Trim(Request("id_item_devolvido"))
	if (id_item_devolvido = "") then Response.Redirect("aviso.asp?id=" & ERR_PARAMETRO_OBRIGATORIO_NAO_ESPECIFICADO)

	dim c_mensagem
	c_mensagem = Trim(Request("c_mensagem"))
	
	if c_mensagem = "" then
		alerta = "Não foi escrita nenhuma mensagem para gravar no bloco de notas de itens devolvidos."
	elseif len(c_mensagem) > MAX_TAM_MENSAGEM_BLOCO_NOTAS_EM_ITEM_DEVOLVIDO then
		alerta = "O tamanho da mensagem (" & Cstr(len(c_mensagem)) & ") excede o tamanho máximo permitido de " & Cstr(MAX_TAM_MENSAGEM_BLOCO_NOTAS_EM_ITEM_DEVOLVIDO) & " caracteres."
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rsMail
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(rsMail, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim emailAdministradorDevolucoes
	dim dtHrMensagem
	dim corpo_mensagem, id_email, msg_erro_grava_email
	dim r_pedido
	if alerta = "" then
		if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then
			alerta = msg_erro
			end if
		end if

	dim s_email_remetente
	dim s_dados_cliente
	dim s_unidade_negocio
	s_dados_cliente = ""
	s_unidade_negocio = ""
	s_email_remetente = getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__PEDIDO_DEVOLUCAO)

	dim campos_a_omitir
	dim vLog()
	dim s_log
	s_log = ""
	campos_a_omitir = "|dt_cadastro|dt_hr_cadastro|anulado_status|anulado_usuario|anulado_data|anulado_data_hora|"


'	GRAVA A MENSAGEM NO BLOCO DE NOTAS
	if alerta = "" then
	'	INICIA A TRANSAÇÃO
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if
			
	'	GERA O NSU PARA GRAVAR A MENSAGEM
		dim intNsuNovoBlocoNotas
		if Not fin_gera_nsu(T_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS, intNsuNovoBlocoNotas, msg_erro) then
			alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		else
			if intNsuNovoBlocoNotas <= 0 then
				alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovoBlocoNotas & ")"
				end if
			end if
		
		if alerta = "" then
			s = "SELECT * FROM t_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS WHERE (id = -1)"
			rs.Open s, cn
			rs.AddNew 
			rs("id")=intNsuNovoBlocoNotas
			rs("id_item_devolvido")=id_item_devolvido
			rs("usuario")=usuario
			rs("loja")=loja
			rs("mensagem")=c_mensagem
			rs.Update 
			if Err <> 0 then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if
			
			log_via_vetor_carrega_do_recordset rs, vLog, campos_a_omitir
			s_log = log_via_vetor_monta_inclusao(vLog)
			
			if rs.State <> 0 then rs.Close
				
			if s_log <> "" then grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS_INCLUSAO, s_log
			end if
		
		if alerta = "" then
			dtHrMensagem = Now
			'Foi encontrado o email para ser usado como remetente da mensagem?
			if s_email_remetente <> "" then
				'Obtém dados do cliente
				s = "SELECT" & _
						" p.st_memorizacao_completa_enderecos," & _
						" p.endereco_nome_iniciais_em_maiusculas AS endereco_nome," & _
						" p.endereco_cnpj_cpf," & _
						" c.nome_iniciais_em_maiusculas AS cliente_nome," & _
						" c.cnpj_cpf AS cliente_cnpj_cpf," & _
						" p.loja," & _
						" lj.unidade_negocio" & _
					" FROM t_PEDIDO p" & _
						" INNER JOIN t_CLIENTE c ON (p.id_cliente = c.id)" & _
						" INNER JOIN t_LOJA lj ON (p.loja = lj.loja)" & _
					" WHERE" & _
						" (p.pedido = '" & pedido_selecionado & "')"
				if rsMail.State <> 0 then rsMail.Close
				rsMail.Open s, cn
				if Not rsMail.Eof then
					s_unidade_negocio = Trim("" & rsMail("unidade_negocio"))
					if rsMail("st_memorizacao_completa_enderecos") <> 0 then
						s_dados_cliente = "Cliente: " & Trim("" & rsMail("endereco_nome")) & " (" & cnpj_cpf_formata(Trim("" & rsMail("endereco_cnpj_cpf"))) & ")"
					else
						s_dados_cliente = "Cliente: " & Trim("" & rsMail("cliente_nome")) & " (" & cnpj_cpf_formata(Trim("" & rsMail("cliente_cnpj_cpf"))) & ")"
						end if
					end if
				
				'Envia mensagem de aviso para o administrador de devoluções
				set emailAdministradorDevolucoes = get_registro_t_parametro(ID_PARAMETRO_PEDIDO_DEVOLUCAO_EMAIL_ADMINISTRADOR)
				if Trim("" & emailAdministradorDevolucoes.campo_texto) <> "" then
					if (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__BS) Or (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__VRF) then
						corpo_mensagem = "Nova mensagem registrada no 'Bloco de Notas (Devolução de Mercadorias)' do pedido " & pedido_selecionado & " por " & usuario & " em " & formata_data_hora_sem_seg(dtHrMensagem) & _
										vbCrLf & _
										"Pedido: " & pedido_selecionado & _
										vbCrLf & _
										s_dados_cliente & _
										vbCrLf & vbCrLf & _
										String(30, "-") & "( Início )" & String(30, "-") & _
										vbCrLf & _
										c_mensagem & _
										vbCrLf & _
										String(31, "-") & "( Fim )" & String(32, "-") & _
										vbCrLf & vbCrLf & _
										"Atenção: esta é uma mensagem automática, NÃO responda a este e-mail!"
						EmailSndSvcGravaMensagemParaEnvio s_email_remetente, _
														"", _
														emailAdministradorDevolucoes.campo_texto, _
														"", _
														"", _
														"Nova mensagem registrada no 'Bloco de Notas (Devolução de Mercadorias)' do pedido " & pedido_selecionado, _
														corpo_mensagem, _
														Now, _
														id_email, _
														msg_erro_grava_email
						end if 'if (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__BS) Or (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__VRF)
					end if 'if Trim("" & emailAdministradorDevolucoes.campo_texto) <> ""
				end if 'if s_email_remetente <> ""
			end if 'if alerta = ""

		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("pedido.asp?pedido_selecionado=" & pedido_selecionado & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
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
	<title>LOJA</title>
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
<div class="MtAlerta" style="width:600px;FONT-WEIGHT:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<BR><BR>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>
<% end if %>

</html>


<%
	if rsMail.State <> 0 then rsMail.Close
	set rsMail = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>