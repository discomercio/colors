<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================================
'	  P E D I D O O C O R R E N C I A N O V A C O N F I R M A . A S P
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
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim alerta
	alerta=""
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    dim url_origem
    url_origem = Trim(Request("url_origem"))

	if Not operacao_permitida(OP_CEN_OCORRENCIAS_EM_PEDIDOS_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim pedido_selecionado
	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim c_texto, c_contato, c_ddd_1, c_tel_1, c_ddd_2, c_tel_2, motivo_ocorrencia
	c_texto = Trim(Request("c_texto"))
	c_contato = Trim(Request("c_contato"))
	c_ddd_1 = retorna_so_digitos(Trim(Request("c_ddd_1")))
	c_tel_1 = retorna_so_digitos(Trim(Request("c_tel_1")))
	c_ddd_2 = retorna_so_digitos(Trim(Request("c_ddd_2")))
	c_tel_2 = retorna_so_digitos(Trim(Request("c_tel_2")))
    motivo_ocorrencia = Trim(Request("motivo_ocorrencia"))
	
	if len(c_texto) > MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS then
		alerta = "O tamanho do texto (" & Cstr(len(c_texto)) & ") excede o tamanho máximo permitido de " & Cstr(MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS) & " caracteres."
    elseif motivo_ocorrencia = "" then
        alerta = "Não foi informado o motivo da abertura da ocorrência."
	elseif c_contato = "" then
		alerta = "Informe o nome da pessoa para contato."
	elseif Not ddd_ok(c_ddd_1) then
		alerta="DDD do 1º telefone é inválido."
	elseif Not telefone_ok(c_tel_1) then
		alerta="Número do 1º telefone é inválido."
	elseif (c_ddd_1 <> "") And ((c_tel_1 = "")) then
		alerta="Preencha o número do 1º telefone."
	elseif (c_ddd_1 = "") And ((c_tel_1 <> "")) then
		alerta="Preencha o DDD do 1º telefone."
	elseif Not ddd_ok(c_ddd_2) then
		alerta="DDD do 2º telefone é inválido."
	elseif Not telefone_ok(c_tel_2) then
		alerta="Número do 2º telefone é inválido."
	elseif (c_ddd_2 <> "") And ((c_tel_2 = "")) then
		alerta="Preencha o número do 2º telefone."
	elseif (c_ddd_2 = "") And ((c_tel_2 <> "")) then
		alerta="Preencha o DDD do 2º telefone."
	elseif (c_tel_1 = "") And (c_tel_2 = "") then
		alerta="Informe pelo menos um número de telefone para contato."
		end if

	dim campos_a_omitir
	dim vLog()
	dim s_log
	s_log = ""
	campos_a_omitir = "|dt_cadastro|dt_hr_cadastro|finalizado_status|finalizado_usuario|finalizado_data|finalizado_data_hora|"


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2, rsMail
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(rsMail, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	VERIFICA SE JÁ EXISTE OCORRÊNCIA ABERTA
'	TAMBÉM PREVINE CONTRA DUPLICAÇÃO CAUSADA POR NOVO ACIONAMENTO APÓS USAR O BOTÃO 'BACK' DO BROWSER
	if alerta = "" then
		s = "SELECT " & _
				"*" & _
			" FROM t_PEDIDO_OCORRENCIA" & _
			" WHERE" & _
				" (pedido = '" & pedido_selecionado & "')" & _
				" AND (finalizado_status = 0)" & _
			" ORDER BY" & _
				" dt_hr_cadastro DESC," & _
				" id DESC"
		set rs2 = cn.Execute(s)
		if Not rs2.Eof then
			alerta = "Cadastramento cancelado, pois já existe uma ocorrência ainda não finalizada cadastrada em " & formata_data_hora_sem_seg(rs2("dt_hr_cadastro"))
			end if
		set rs2 = Nothing
		end if

	dim id_email, corpo_mensagem, msg_erro_grava_email, dtHrMensagem
	dim s_email_vendedor
	s_email_vendedor = ""
	if alerta = "" then
		s = "SELECT" & _
				" tP.vendedor," & _
				" tU.email" & _
			" FROM t_PEDIDO tP" & _
				" INNER JOIN t_USUARIO tU ON (tP.vendedor = tU.usuario)" & _
			" WHERE" & _
				" (tP.pedido = '" & pedido_selecionado & "')"
		if rsMail.State <> 0 then rsMail.Close
		rsMail.Open s, cn
		if Not rsMail.Eof then
			s_email_vendedor = LCase(Trim("" & rsMail("email")))
			end if
		end if

	dim s_email_remetente
	dim r_usuario
	dim s_dados_cliente
	dim s_unidade_negocio
	s_dados_cliente = ""
	s_unidade_negocio = ""
	s_email_remetente = getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__MENSAGEM_SISTEMA)
	call le_usuario(usuario, r_usuario, msg_erro)

	if alerta = "" then
		'Se encontrou e-mail do vendedor para enviar mensagem de aviso, obtém demais informações para a montagem da mensagem
		if s_email_vendedor <> "" then
			call le_usuario(usuario, r_usuario, msg_erro)

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
			end if
		end if

	dim r_pedido
	if alerta = "" then
		if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then
			alerta = msg_erro
			end if
		end if

	dim s_descricao_motivo_abertura
	s_descricao_motivo_abertura = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__MOTIVO_ABERTURA, motivo_ocorrencia)


'	GRAVA A OCORRÊNCIA
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
			
	'	GERA O NSU PARA GRAVAR A OCORRÊNCIA
		dim intNsuNovaOcorrencia
		if Not fin_gera_nsu(T_PEDIDO_OCORRENCIA, intNsuNovaOcorrencia, msg_erro) then
			alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		else
			if intNsuNovaOcorrencia <= 0 then
				alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovaOcorrencia & ")"
				end if
			end if
		
		if alerta = "" then
			s = "SELECT * FROM t_PEDIDO_OCORRENCIA WHERE (id = -1)"
			rs.Open s, cn
			rs.AddNew 
			rs("id")=intNsuNovaOcorrencia
			rs("pedido")=pedido_selecionado
			rs("usuario_cadastro")=usuario
            rs("cod_motivo_abertura")=motivo_ocorrencia
			rs("texto_ocorrencia")=c_texto
			rs("contato") = c_contato
			rs("ddd_1") = c_ddd_1
			rs("tel_1") = c_tel_1
			rs("ddd_2") = c_ddd_2
			rs("tel_2") = c_tel_2
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
				
			if s_log <> "" then grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_OCORRENCIA_INCLUSAO, s_log
			end if
			
		if alerta = "" then
			if s_email_vendedor <> "" then
				if s_email_remetente <> "" then
					if UCase(usuario) <> UCase(r_pedido.vendedor) then
						if (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__BS) Or (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__VRF) then
							dtHrMensagem = Now

							corpo_mensagem = "Usuário '" & usuario & "' (" & r_usuario.nome_iniciais_em_maiusculas & ") cadastrou uma nova ocorrência no pedido " & pedido_selecionado & " em " & formata_data_hora_sem_seg(dtHrMensagem) & _
											vbCrLf & _
											"Pedido: " & pedido_selecionado & _
											vbCrLf & _
											s_dados_cliente & _
											vbCrLf & _
											"Ocorrência: " & s_descricao_motivo_abertura

							if Trim(c_texto) <> "" then
								corpo_mensagem = corpo_mensagem & _
												vbCrLf & vbCrLf & _
												String(30, "-") & "( Início )" & String(30, "-") & _
												vbCrLf & _
												c_texto & _
												vbCrLf & _
												String(31, "-") & "( Fim )" & String(32, "-")
								end if

							corpo_mensagem = corpo_mensagem & _
											vbCrLf & vbCrLf & _
											"Atenção: esta é uma mensagem automática, NÃO responda a este e-mail!"

							'Envia e-mail para o vendedor
							EmailSndSvcGravaMensagemParaEnvio s_email_remetente, _
															"", _
															s_email_vendedor, _
															"", _
															"", _
															"Nova ocorrência cadastrada no pedido " & pedido_selecionado, _
															corpo_mensagem, _
															Now, _
															id_email, _
															msg_erro_grava_email
							end if 'if (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__BS) Or (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__VRF)
						end if 'if UCase(usuario) <> UCase(r_pedido.vendedor)
					end if 'if s_email_remetente <> ""
				end if 'if s_email_vendedor <> ""
			end if 'if alerta = "" then

		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("pedido.asp?pedido_selecionado=" & pedido_selecionado & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))) & "&url_origem=" & url_origem
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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>