<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/Global.asp" -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================================
'	  P E D I D O B L O C O N O T A S N O V O C O N F I R M A . A S P
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

	if Not operacao_permitida(OP_CEN_BLOCO_NOTAS_PEDIDO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim pedido_selecionado
	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim c_mensagem, c_nivel_acesso_bloco_notas, ckb_notificar_vendedor, ckb_notificar_demais_particip, s_vendedor, s_email_vendedor
	c_mensagem = Trim(Request("c_mensagem"))
	c_nivel_acesso_bloco_notas = Trim(Request("c_nivel_acesso_bloco_notas"))
	ckb_notificar_vendedor = Trim(Request("ckb_notificar_vendedor"))
	ckb_notificar_demais_particip = Trim(Request("ckb_notificar_demais_particip"))

	if c_mensagem = "" then
		alerta = "Não foi escrita nenhuma mensagem para gravar no bloco de notas."
	elseif len(c_mensagem) > MAX_TAM_MENSAGEM_BLOCO_NOTAS then
		alerta = "O tamanho da mensagem (" & Cstr(len(c_mensagem)) & ") excede o tamanho máximo permitido de " & Cstr(MAX_TAM_MENSAGEM_BLOCO_NOTAS) & " caracteres."
	elseif c_nivel_acesso_bloco_notas = "" then
		alerta = "Não foi definido o nível de acesso para a leitura da mensagem."
	elseif converte_numero(c_nivel_acesso_bloco_notas) = 0 then
		alerta = "Nível de acesso definido para a leitura da mensagem é inválido: " & c_nivel_acesso_bloco_notas
	elseif converte_numero(c_nivel_acesso_bloco_notas) < converte_numero(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO) then
		alerta = "Nível de acesso definido para a leitura da mensagem é inválido: " & c_nivel_acesso_bloco_notas
	elseif converte_numero(c_nivel_acesso_bloco_notas) > converte_numero(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__SIGILOSO) then
		alerta = "Nível de acesso definido para a leitura da mensagem é inválido: " & c_nivel_acesso_bloco_notas
		end if

	dim id_email, corpo_mensagem, msg_erro_grava_email, dtHrMensagem
	dim s_dados_cliente
	dim rParametro, r_usuario, r_pedido, r_cliente
	dim i, v_demais_particip
	redim v_demais_particip(0)
	set v_demais_particip(ubound(v_demais_particip)) = new cl_TRES_COLUNAS
	v_demais_particip(ubound(v_demais_particip)).c1 = ""

	s_vendedor = ""
	s_email_vendedor = ""
	if alerta = "" then
		if ckb_notificar_vendedor <> "" then
			s = "SELECT" & _
					" tP.vendedor," & _
					" tU.email," & _
					" (SELECT Coalesce(Max(nivel_acesso_bloco_notas_pedido), " & Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__NAO_DEFINIDO) & ") AS max_nivel_acesso_bloco_notas_pedido FROM t_PERFIL INNER JOIN t_PERFIL_X_USUARIO ON t_PERFIL.id=t_PERFIL_X_USUARIO.id_perfil WHERE (t_PERFIL_X_USUARIO.usuario = tU.usuario)) AS nivel_acesso_bloco_notas_pedido" & _
				" FROM t_PEDIDO tP" & _
					" INNER JOIN t_USUARIO tU ON (tP.vendedor = tU.usuario)" & _
				" WHERE" & _
					" (tP.pedido = '" & pedido_selecionado & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Not rs.Eof then
				if Trim("" & rs("email")) = "" then
					alerta = "O vendedor não possui e-mail cadastrado e, portanto, não é possível enviar a mensagem de notificação!"
				elseif converte_numero(rs("nivel_acesso_bloco_notas_pedido")) < converte_numero(c_nivel_acesso_bloco_notas) then
					alerta = "O vendedor não possui nível de acesso suficiente para ler a mensagem!"
					end if
				end if

			s_vendedor = Trim("" & rs("vendedor"))
			s_email_vendedor = LCase(Trim("" & rs("email")))
			end if 'if ckb_notificar_vendedor <> ""

		if ckb_notificar_demais_particip <> "" then
			s = "SELECT DISTINCT" & _
					" tPBN.usuario," & _
					" tU.email," & _
					" (SELECT Coalesce(Max(nivel_acesso_bloco_notas_pedido), " & Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__NAO_DEFINIDO) & ") AS max_nivel_acesso_bloco_notas_pedido FROM t_PERFIL INNER JOIN t_PERFIL_X_USUARIO ON t_PERFIL.id=t_PERFIL_X_USUARIO.id_perfil WHERE (t_PERFIL_X_USUARIO.usuario = tU.usuario)) AS nivel_acesso_bloco_notas_pedido" & _
				" FROM t_PEDIDO_BLOCO_NOTAS tPBN" & _
					" LEFT JOIN t_USUARIO tU ON (tU.usuario = tPBN.usuario)" & _
				" WHERE" & _
					" (pedido = '" & pedido_selecionado & "')" & _
					" AND (tPBN.usuario NOT IN ('" & usuario & "','" & s_vendedor & "'))"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			do while Not rs.Eof
				'Possui acesso suficiente p/ ler a mensagem?
				if converte_numero(rs("nivel_acesso_bloco_notas_pedido")) >= converte_numero(c_nivel_acesso_bloco_notas) then
					if v_demais_particip(ubound(v_demais_particip)).c1 <> "" then
						redim preserve v_demais_particip(ubound(v_demais_particip)+1)
						set v_demais_particip(ubound(v_demais_particip)) = new cl_TRES_COLUNAS
						end if
					v_demais_particip(ubound(v_demais_particip)).c1 = Trim("" & rs("usuario"))
					v_demais_particip(ubound(v_demais_particip)).c2 = rs("nivel_acesso_bloco_notas_pedido")
					v_demais_particip(ubound(v_demais_particip)).c3 = Trim("" & rs("email"))
					end if

				rs.MoveNext
				loop
			end if 'if ckb_notificar_demais_particip <> ""
		end if 'if alerta = ""

	if alerta = "" then
		if (ckb_notificar_vendedor <> "") Or (ckb_notificar_demais_particip <> "") then
			set rParametro = get_registro_t_parametro(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__MENSAGEM_SISTEMA)

			call le_usuario(usuario, r_usuario, msg_erro)
			call le_pedido(pedido_selecionado, r_pedido, msg_erro)
			
			set r_cliente = New cl_CLIENTE
			call x_cliente_bd(r_pedido.id_cliente, r_cliente)
			end if
		end if

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
		if rs.State <> 0 then rs.Close
		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if
		
	'	GERA O NSU PARA GRAVAR A MENSAGEM
		dim intNsuNovoBlocoNotas
		if Not fin_gera_nsu(T_PEDIDO_BLOCO_NOTAS, intNsuNovoBlocoNotas, msg_erro) then 
			alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		else
			if intNsuNovoBlocoNotas <= 0 then
				alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovoBlocoNotas & ")"
				end if
			end if
		
		if alerta = "" then
			s = "SELECT * FROM t_PEDIDO_BLOCO_NOTAS WHERE (id = -1)"
			rs.Open s, cn
			rs.AddNew 
			rs("id")=intNsuNovoBlocoNotas
			rs("pedido")=pedido_selecionado
			rs("usuario")=usuario
			rs("nivel_acesso")=CLng(c_nivel_acesso_bloco_notas)
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
				
			if s_log <> "" then grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_BLOCO_NOTAS_INCLUSAO, s_log
			end if 'if alerta = ""
		
		if alerta = "" then
			if (ckb_notificar_vendedor <> "") Or (ckb_notificar_demais_particip <> "") then
				dtHrMensagem = Now
		
				if r_pedido.st_memorizacao_completa_enderecos <> 0 then
					s_dados_cliente = "Cliente: " & r_pedido.endereco_nome_iniciais_em_maiusculas & " (" & cnpj_cpf_formata(r_pedido.endereco_cnpj_cpf) & ")"
				else
					s_dados_cliente = "Cliente: " & r_cliente.nome_iniciais_em_maiusculas & " (" & cnpj_cpf_formata(r_cliente.cnpj_cpf) & ")"
					end if
			
				corpo_mensagem = "Usuário '" & usuario & "' (" & r_usuario.nome_iniciais_em_maiusculas & ") registrou uma mensagem no bloco de notas do pedido " & pedido_selecionado & " em " & formata_data_hora_sem_seg(dtHrMensagem) & _
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

				if Trim("" & rParametro.campo_texto) <> "" then
					if (ckb_notificar_vendedor <> "") And (s_email_vendedor <> "") then
						'Envia e-mail para o vendedor
						EmailSndSvcGravaMensagemParaEnvio Trim("" & rParametro.campo_texto), _
														"", _
														s_email_vendedor, _
														"", _
														"", _
														"Nova mensagem registrada no bloco de notas do pedido " & pedido_selecionado, _
														corpo_mensagem, _
														Now, _
														id_email, _
														msg_erro_grava_email
						end if 'if (ckb_notificar_vendedor <> "") And (s_email_vendedor <> "")
				
					if ckb_notificar_demais_particip <> "" then
						'Envia e-mail para os demais participantes que tenham escrito mensagens anteriormente no bloco de notas
						for i=LBound(v_demais_particip) to UBound(v_demais_particip)
							if (Trim("" & v_demais_particip(i).c1) <> "") And (Trim("" & v_demais_particip(i).c3) <> "") then
								EmailSndSvcGravaMensagemParaEnvio Trim("" & rParametro.campo_texto), _
																"", _
																Trim("" & v_demais_particip(i).c3), _
																"", _
																"", _
																"Nova mensagem registrada no bloco de notas do pedido " & pedido_selecionado, _
																corpo_mensagem, _
																Now, _
																id_email, _
																msg_erro_grava_email
								end if 'if (Trim("" & v_demais_particip(i).c1) <> "") And (Trim("" & v_demais_particip(i).c3) <> "")
							next
						end if 'if ckb_notificar_demais_particip <> ""
					end if 'if Trim("" & rParametro.campo_texto) <> ""
				end if 'if (ckb_notificar_vendedor <> "") Or (ckb_notificar_demais_particip <> "")
			end if 'if alerta = ""

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