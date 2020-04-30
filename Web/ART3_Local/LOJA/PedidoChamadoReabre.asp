<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ==============================================
'	  P E D I D O C H A M A D O R E A B R E . A S P
'     ==============================================
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
	dim usuario, usuario_email
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim alerta
	alerta=""
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim blnIsUsuarioResponsavelDepto, blnIsUsuarioCadastroChamado
    blnIsUsuarioCadastroChamado = CBool(Request.Form("blnIsUsuarioCadastroChamado"))
    blnIsUsuarioResponsavelDepto = CBool(Request.Form("blnIsUsuarioResponsavelDepto"))

	if Not operacao_permitida(OP_LJA_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) And _
    Not blnIsUsuarioResponsavelDepto And _
    Not blnIsUsuarioCadastroChamado then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim pedido_selecionado
    dim usuario_tx, usuario_rx

    dim motivo_abertura
    dim cod_motivo_finalizacao, finalizado_usuario, finalizado_data_hora, texto_finalizacao, nivel_acesso_chamado, texto_msg
    
	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim id_chamado
	id_chamado = Trim(request("id_chamado"))
	if (id_chamado = "") then Response.Redirect("aviso.asp?id=" & ERR_PARAMETRO_OBRIGATORIO_NAO_ESPECIFICADO)

	dim campos_a_omitir
	dim vLog()
	dim s_log
	s_log = ""
	campos_a_omitir = "|dt_cadastro|dt_hr_cadastro|"


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

'	GRAVA A MENSAGEM P/ ESTE CHAMADO
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

        ' recupera email do usuário logado
        usuario_email = obtem_email_usuario(usuario)

		if rs.State <> 0 then rs.Close
		s = "SELECT tPC.usuario_cadastro AS chamado_usuario_cadastro, * FROM t_PEDIDO_CHAMADO tPC INNER JOIN t_PEDIDO_CHAMADO_DEPTO tPCD ON (tPC.id_depto=tPCD.id) WHERE (tPC.id = '" & id_chamado & "')"
		rs.Open s, cn
        if CInt(rs("finalizado_status")) = 0 then
            alerta = "Registro do chamado (id=" & id_chamado & ") do pedido " & pedido_selecionado & " já se encontra aberto/em andamento!!"
        end if
		if alerta = "" then
            cod_motivo_finalizacao=rs("cod_motivo_finalizacao")
            texto_finalizacao=rs("texto_finalizacao")
            finalizado_usuario=rs("finalizado_usuario")
			finalizado_data_hora=rs("finalizado_data_hora")
            motivo_abertura=rs("cod_motivo_abertura")
            nivel_acesso_chamado=rs("nivel_acesso")
            usuario_tx = Trim("" & rs("chamado_usuario_cadastro"))
            usuario_rx = Trim("" & rs("usuario_responsavel"))

			rs("finalizado_status")=0      
            rs("cod_motivo_finalizacao")=""
            rs("texto_finalizacao")=""
            rs("chamado_reaberto_contagem")=CInt(rs("chamado_reaberto_contagem"))+1
            rs("chamado_reaberto_dt_ult")=Now
            rs("chamado_reaberto_usuario_ult")=usuario
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
			
			if s_log <> "" then grava_log usuario, "", pedido_selecionado, "", "PED CHAMD REABRE", s_log
			end if

        '	GERA O NSU PARA GRAVAR A MENSAGEM P/ ESTE CHAMADO
		dim intNsuNovaChamadoMensagem
		if Not fin_gera_nsu(T_PEDIDO_CHAMADO_MENSAGEM, intNsuNovaChamadoMensagem, msg_erro) then
			alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		else
			if intNsuNovaChamadoMensagem <= 0 then
				alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovaChamadoMensagem & ")"
				end if
			end if

        if alerta = "" then
            texto_msg = "O CHAMADO FOI REABERTO POR " & usuario & "." & chr(13) & _
                        "INFORMAÇÕES DA ÚLTIMA VEZ EM QUE FOI FINALIZADO:" & chr(13) & _
                        "FINALIZADO POR: " & finalizado_usuario & " em " & formata_data_hora(finalizado_data_hora) & "." & chr(13) & _
                        "MOTIVO DA FINALIZAÇÃO: " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_FINALIZACAO, cod_motivo_finalizacao) & chr(13) & _
                        "SOLUÇÃO: " & texto_finalizacao 

			s = "SELECT * FROM t_PEDIDO_CHAMADO_MENSAGEM WHERE (id = -1)"
			rs.Open s, cn
			rs.AddNew 
			rs("id")=intNsuNovaChamadoMensagem
			rs("id_chamado")=CLng(id_chamado)
			rs("usuario_cadastro")="SISTEMA"
			rs("fluxo_mensagem") = ""
			rs("texto_mensagem")=texto_msg
            rs("nivel_acesso")=nivel_acesso_chamado
            rs("tipo_usuario_cadastro")="S"
			rs.Update 
			if Err <> 0 then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if
        end if
		
        if rs.State <> 0 then rs.Close

		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~

        ' recupera os emails de todos os envolvidos no chamado
            dim corpo_mensagem, msg_erro_grava_email, id_email, destinatario_cc, destinatario_to
            destinatario_cc = ""
            s = "SELECT DISTINCT email, usuario" & _
                   " FROM t_USUARIO tU" & _
                   " INNER JOIN t_PEDIDO_CHAMADO_MENSAGEM tPCM ON" & _ 
		                   " (tU.usuario = tPCM.usuario_cadastro)" & _
                   " WHERE tPCM.id_chamado = " & id_chamado & _
                   " UNION" & _ 
                   " SELECT DISTINCT email, usuario" & _
                   " FROM t_USUARIO tU" & _
                   " INNER JOIN t_PEDIDO_CHAMADO tPC ON" & _ 
		                   " (tU.usuario = tPC.usuario_cadastro)" & _
                   " WHERE tPC.id = " & id_chamado & _
                   " UNION" & _ 
                   " SELECT DISTINCT email, usuario" & _
                   " FROM t_USUARIO tU" & _
                   " INNER JOIN t_PEDIDO_CHAMADO tPC ON" & _ 
		                   " (tU.usuario = tPC.usuario_cadastro)" & _
                   " INNER JOIN t_PEDIDO_CHAMADO_DEPTO tPCD ON" & _ 
		                   " (tPC.id_depto = tPCD.id)" & _
                   " WHERE tPC.id = " & id_chamado
      
                corpo_mensagem = "O usuário " & usuario & " - " & x_usuario(usuario) & " reabriu o chamado referente ao pedido " & pedido_selecionado & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Reaberto por: " & usuario & " em " & Now & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Motivo da abertura: " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA, motivo_abertura) & "." & chr(13) & chr(10)
    
                if usuario = usuario_rx then
                    destinatario_to = obtem_email_usuario(usuario_tx)
                else
                    destinatario_to = obtem_email_usuario(usuario_rx)
                end if

                rs.Open s, cn
                do while Not rs.Eof

                    if Trim("" & rs("email")) <> usuario_email And Trim("" & rs("email")) <> destinatario_to And Trim("" & rs("email")) <> "" then
                        if destinatario_cc <> "" then destinatario_cc = destinatario_cc & ";"
                        destinatario_cc = destinatario_cc & Trim("" & rs("email"))
                    end if
                    rs.MoveNext
                loop
			    if rs.State <> 0 then rs.Close

                if usuario_email = "" then
                    corpo_mensagem = corpo_mensagem & "---------------------------------------------------------------------------------------------------------"   & chr(13) & chr(10)
                    corpo_mensagem = corpo_mensagem & "E-MAIL ENVIADO AUTOMATICAMENTE PELO SISTEMA. NÃO RESPONDA ESTE E-MAIL, POIS ESTA CONTA NÃO É MONITORADA!!"                 
                    end if
                
                if destinatario_to = "" and destinatario_cc <> "" then
                    destinatario_to = getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__CHAMADOS_EM_PEDIDOS)
                end if

                if destinatario_to <> "" then
                    EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__CHAMADOS_EM_PEDIDOS), _
                                                        usuario_email, _
                                                        destinatario_to, _
                                                        destinatario_cc, _
                                                        "", _
                                                        "Chamado reaberto referente ao pedido " & pedido_selecionado, _
                                                        corpo_mensagem, _
                                                        Now, _
                                                        id_email, _
                                                        msg_erro_grava_email
                end if

                if msg_erro_grava_email <> "" then
                    Err=1
                end if

			    if Err=0 then 
				    Response.Redirect("pedido.asp?pedido_selecionado=" & pedido_selecionado & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))) & "#aChamados"
			    else
			    'alerta=Cstr(Err) & ": " & Err.Description
                    alerta = "Erro ao gravar email. ID_EMAIL: " & CStr(id_email) & " MSG ERRO: " & msg_erro_grava_email	
                
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