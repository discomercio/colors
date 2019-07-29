<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================================================
'	  RelPedidoChamadoGravaDados.asp
'     ===============================================================================
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

	class cl_TIPO_GRAVA_REL_CHAMADO
		dim id_chamado
		dim pedido
		dim mensagem
		dim motivo_finalizacao
		dim texto_finalizacao
        dim msg_nivel_acesso
        dim motivo_abertura
		end class
		
	dim s, msg_erro
	dim usuario, usuario_email, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim alerta
	alerta=""
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    dim c_dt_cad_chamado_inicio, c_dt_cad_chamado_termino
	dim rb_status, rb_posicao, c_qtde_chamados, intQtdeChamados, vChamado
    dim c_depto, blnHasDepto, c_motivo_abertura_filtro
	rb_status=Trim(Request("rb_status"))
    rb_posicao=Trim(Request("rb_posicao"))
	c_qtde_chamados=Trim(Request("c_qtde_chamados"))
    c_depto = Trim(Request("c_depto"))
    blnHasDepto = Trim(Request("blnHasDepto"))
    c_motivo_abertura_filtro = Trim(Request("c_motivo_abertura"))
	intQtdeChamados=CInt(c_qtde_chamados)
	
	redim vChamado(0)
	set vChamado(Ubound(vChamado)) = new cl_TIPO_GRAVA_REL_CHAMADO
	vChamado(Ubound(vChamado)).id_chamado = ""
	
	dim i
	dim c_id_chamado, c_pedido, c_nova_msg, c_motivo_finalizacao, c_solucao, c_nivel_acesso_chamado, c_motivo_abertura
    dim c_rel_origem
    c_rel_origem = Request.Form("c_rel_origem")
    c_dt_cad_chamado_inicio=Trim(Request("c_dt_cad_chamado_inicio"))
	c_dt_cad_chamado_termino=Trim(Request("c_dt_cad_chamado_termino"))
	for i = 1 to intQtdeChamados
		c_id_chamado = Trim(Request.Form("c_id_chamado_" & Cstr(i)))
		c_pedido = Trim(Request.Form("c_pedido_" & Cstr(i)))
		c_nova_msg = Trim(Request.Form("c_nova_msg_" & Cstr(i)))
		c_motivo_finalizacao = Trim(Request.Form("c_motivo_finalizacao_" & Cstr(i)))
		c_motivo_abertura = Trim(Request.Form("c_motivo_abertura_" & Cstr(i)))
        c_nivel_acesso_chamado = Trim(Request.Form("c_nivel_acesso_chamado_" & Cstr(i)))
		c_solucao = Trim(Request.Form("c_solucao_" & Cstr(i)))
		if (c_id_chamado<>"") And ( (c_nova_msg<>"") Or (c_motivo_finalizacao<>"") Or (c_solucao<>"") ) then
			if vChamado(Ubound(vChamado)).id_chamado <> "" then
				redim preserve vChamado(Ubound(vChamado)+1)
				set vChamado(Ubound(vChamado)) = new cl_TIPO_GRAVA_REL_CHAMADO
				end if
			vChamado(Ubound(vChamado)).id_chamado = c_id_chamado
			vChamado(Ubound(vChamado)).pedido = c_pedido
			vChamado(Ubound(vChamado)).mensagem = c_nova_msg
			vChamado(Ubound(vChamado)).motivo_finalizacao = c_motivo_finalizacao
			vChamado(Ubound(vChamado)).motivo_abertura = c_motivo_abertura
			vChamado(Ubound(vChamado)).texto_finalizacao = c_solucao
            vChamado(Ubound(vChamado)).msg_nivel_acesso = c_nivel_acesso_chamado
			end if
		next

	for i=Lbound(vChamado) to Ubound(vChamado)
		if Trim(vChamado(i).id_chamado)<>"" then
			if len(vChamado(i).mensagem) > MAX_TAM_MENSAGEM_CHAMADOS_EM_PEDIDOS then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O tamanho do texto da mensagem (" & Cstr(len(vChamado(i).mensagem)) & ")  do chamado do pedido " & vChamado(i).pedido & " excede o tamanho máximo permitido de " & Cstr(MAX_TAM_MENSAGEM_CHAMADOS_EM_PEDIDOS) & " caracteres."
			elseif len(vChamado(i).texto_finalizacao) > MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O tamanho do texto descrevendo a solução (" & Cstr(len(vChamado(i).texto_finalizacao)) & ") do chamado do pedido " & vChamado(i).pedido & " excede o tamanho máximo permitido de " & Cstr(MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS) & " caracteres."
				end if

			if (Trim(vChamado(i).motivo_finalizacao)<>"") Or (Trim(vChamado(i).texto_finalizacao)<>"") then
				if Trim(vChamado(i).motivo_finalizacao)="" then
					'alerta=texto_add_br(alerta)
				'	alerta=alerta & "Não foi selecionado o motivo da finalização para o pedido " & vChamado(i).pedido & "!!<br>Ao finalizar um chamado, é necessário informar o motivo da finalização e o texto descrevendo a solução."
				elseif Trim(vChamado(i).texto_finalizacao)="" then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Não foi informado o texto descrevendo a solução do chamado do pedido " & vChamado(i).pedido & "!!<br>Ao finalizar um chamado, é necessário informar o motivo da finalização e o texto descrevendo a solução."
					end if
				end if
			end if
		next
	
	
	dim intNsuNovaChamadoMensagem, fluxo_mensagem
	dim campos_a_omitir
	dim vLog(), vLog1(), vLog2()
	dim s_log
    dim corpo_mensagem, msg_erro_grava_email, id_email, destinatario_cc, destinatario_to
    dim usuario_tx, usuario_rx

	s_log = ""
	campos_a_omitir = "|dt_cadastro|dt_hr_cadastro|finalizado_data|finalizado_data_hora|"
    fluxo_mensagem = ""


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2
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

		if Not cria_recordset_pessimista(rs2, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if
			
		for i=Lbound(vChamado) to Ubound(vChamado)
			if Trim(vChamado(i).id_chamado)<>"" then

			'	TEM MENSAGEM NOVA P/ GRAVAR?
				if Trim(vChamado(i).mensagem)<>"" then

                    if vChamado(i).msg_nivel_acesso = "" then
                        alerta=texto_add_br(alerta)
					    alerta=alerta & "Não foi informado o nível de acesso para a mensagem do pedido " & vChamado(i).pedido & "."
                    end if

                    '   VERIFICA SE O AUTOR DA MENSAGEM É O MESMO QUE ABRIU O CHAMADO OU SE É O RESPONSÁVEL PELO DEPARTAMENTO
                    s = "SELECT tPC.usuario_cadastro AS chamado_usuario, tPCD.usuario_responsavel FROM t_PEDIDO_CHAMADO tPC INNER JOIN t_PEDIDO_CHAMADO_DEPTO tPCD ON (tPC.id_depto=tPCD.id)" & _
                            " WHERE tPC.id = '" & vChamado(i).id_chamado & "'"
                    rs.Open s, cn
                    if Not rs.Eof then
                        usuario_tx = Trim("" & rs("chamado_usuario"))
                        usuario_rx = Trim("" & rs("usuario_responsavel"))
                        
                        if UCase(usuario) = UCase(usuario_rx) then 
                            fluxo_mensagem = COD_FLUXO_MENSAGEM_CHAMADOS_EM_PEDIDOS__RX
                        elseif UCase(usuario) = UCase(usuario_tx) then
                            fluxo_mensagem = COD_FLUXO_MENSAGEM_CHAMADOS_EM_PEDIDOS__TX
                        else
                            fluxo_mensagem = ""
                        end if
                    end if        

		            if rs.State <> 0 then rs.Close            

				'	GERA O NSU PARA GRAVAR A MENSAGEM P/ ESTE CHAMADO
					if Not fin_gera_nsu(T_PEDIDO_CHAMADO_MENSAGEM, intNsuNovaChamadoMensagem, msg_erro) then
						alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
					else
						if intNsuNovaChamadoMensagem <= 0 then
							alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovaChamadoMensagem & ")"
							end if
						end if
					
					if alerta = "" then
						s = "SELECT * FROM t_PEDIDO_CHAMADO_MENSAGEM WHERE (id = -1)"
						rs.Open s, cn
						rs.AddNew 
						rs("id")=intNsuNovaChamadoMensagem
						rs("id_chamado")=CLng(vChamado(i).id_chamado)
						rs("usuario_cadastro")=usuario
						rs("fluxo_mensagem") = fluxo_mensagem
						rs("texto_mensagem")=Trim(vChamado(i).mensagem)
                        rs("nivel_acesso")=vChamado(i).msg_nivel_acesso
                        rs("loja")=loja
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

                    ' recupera os emails de todos os envolvidos no chamado
                    destinatario_cc = ""
                    s = "SELECT DISTINCT email, usuario" & _
                           " FROM t_USUARIO tU" & _
                           " INNER JOIN t_PEDIDO_CHAMADO_MENSAGEM tPCM ON" & _ 
		                           " (tU.usuario = tPCM.usuario_cadastro)" & _
                           " WHERE tPCM.id_chamado = " & vChamado(i).id_chamado
                           

                corpo_mensagem = "O usuário " & usuario & " - " & x_usuario(usuario) & " incluiu uma nova mensagem no chamado do pedido " & vChamado(i).pedido & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Data da inclusão: " & Now & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Loja: " & CStr(loja) & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Nível de acesso: " & nivel_acesso_chamado_pedido_descricao(vChamado(i).msg_nivel_acesso) & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Mensagem: " & vChamado(i).mensagem & "." & chr(13) & chr(10)
    
                if usuario = usuario_rx then
                    if CInt(obtem_nivel_acesso_chamado_pedido(cn, usuario_tx)) >= CInt(vChamado(i).msg_nivel_acesso) then
                        destinatario_to = obtem_email_usuario(usuario_tx)
                    end if
                else
                    if CInt(obtem_nivel_acesso_chamado_pedido(cn, usuario_rx)) >= CInt(vChamado(i).msg_nivel_acesso) then
                        destinatario_to = obtem_email_usuario(usuario_rx)
                    end if
                end if

                rs.Open s, cn
                do while Not rs.Eof

                    if Trim("" & rs("email")) <> usuario_email And Trim("" & rs("email")) <> destinatario_to And Trim("" & rs("email")) <> "" then
                        if CInt(obtem_nivel_acesso_chamado_pedido(cn, Trim("" & rs("usuario")))) >= CInt(vChamado(i).msg_nivel_acesso) then                        
                            if destinatario_cc <> "" then destinatario_cc = destinatario_cc & ";"
                            destinatario_cc = destinatario_cc & Trim("" & rs("email"))
                        end if
                    end if
                    rs.MoveNext
                loop
			    if rs.State <> 0 then rs.Close

                if usuario_email = "" then
                    corpo_mensagem = corpo_mensagem & "---------------------------------------------------------------------------------------------------------"   & chr(13) & chr(10)
                    corpo_mensagem = corpo_mensagem & "E-MAIL ENVIADO AUTOMATICAMENTE PELO SISTEMA. NÃO RESPONDA ESTE E-MAIL, POIS ESTA CONTA NÃO É MONITORADA!!"                   
                    end if

                if destinatario_to = "" and destinatario_cc <> "" then
                    destinatario_to = EMAILSNDSVC_REMETENTE__CHAMADOS_EM_PEDIDOS
                end if
                
                if destinatario_to <> "" then
                    EmailSndSvcGravaMensagemParaEnvio EMAILSNDSVC_REMETENTE__CHAMADOS_EM_PEDIDOS, _
                                                        usuario_email, _
                                                        destinatario_to, _
                                                        destinatario_cc, _
                                                        "", _
                                                        "Nova mensagem incluída no chamado do pedido " & vChamado(i).pedido, _
                                                        corpo_mensagem, _
                                                        Now, _
                                                        id_email, _
                                                        msg_erro_grava_email
                end if                        
							
						if s_log <> "" then grava_log usuario, "", vChamado(i).pedido, "", OP_LOG_PEDIDO_CHAMADO_MENSAGEM_INCLUSAO, s_log
						end if
					end if  'if Trim(vChamado(i).mensagem)<>""
					
			'	FINALIZA O CHAMADO?
				if Trim(vChamado(i).motivo_finalizacao)<>"" then

                    '   VERIFICA SE O USUÁRIO QUE FINALIZOU É O MESMO QUE ABRIU O CHAMADO OU SE É O RESPONSÁVEL PELO DEPARTAMENTO
                    s = "SELECT tPC.usuario_cadastro AS chamado_usuario, tPCD.usuario_responsavel FROM t_PEDIDO_CHAMADO tPC INNER JOIN t_PEDIDO_CHAMADO_DEPTO tPCD ON (tPC.id_depto=tPCD.id)" & _
                            " WHERE tPC.id = '" & vChamado(i).id_chamado & "'"
                    rs.Open s, cn
                    if Not rs.Eof then
                        usuario_tx = Trim("" & rs("chamado_usuario"))
                        usuario_rx = Trim("" & rs("usuario_responsavel"))
                        
                        if usuario = usuario_tx then 
                            fluxo_mensagem = COD_FLUXO_MENSAGEM_CHAMADOS_EM_PEDIDOS__TX
                        elseif usuario = usuario_rx then
                            fluxo_mensagem = COD_FLUXO_MENSAGEM_CHAMADOS_EM_PEDIDOS__RX
                        else
                            fluxo_mensagem = ""
                        end if
                    end if        

		            if rs.State <> 0 then rs.Close 

					s = "SELECT * FROM t_PEDIDO_CHAMADO WHERE (id = " & vChamado(i).id_chamado & ")"
					rs2.Open s, cn
					if rs2.Eof then
						alerta = "Registro do chamado (id=" & vChamado(i).id_chamado & ") do pedido " & vChamado(i).pedido & " não foi localizado no banco de dados!!"
						exit for
						end if
					
					if CInt(rs2("finalizado_status"))<>0 then
						alerta = "Registro do chamado (id=" & vChamado(i).id_chamado & ") do pedido " & vChamado(i).pedido & " já se encontra finalizado!!"
						exit for
						end if
						
					if alerta = "" then
						log_via_vetor_carrega_do_recordset rs2, vLog1, campos_a_omitir
						rs2("finalizado_status")=1
						rs2("finalizado_usuario")=usuario
						rs2("finalizado_data")=Date
						rs2("finalizado_data_hora")=Now
						rs2("cod_motivo_finalizacao")=vChamado(i).motivo_finalizacao
						rs2("texto_finalizacao")=vChamado(i).texto_finalizacao
						rs2.Update
						
						if Err <> 0 then
							alerta = Cstr(Err) & ": " & Err.Description
						else

                        ' recupera os emails de todos os envolvidos no chamado
                    destinatario_cc = ""
                    s = "SELECT DISTINCT email" & _
                           " FROM t_USUARIO tU" & _
                           " INNER JOIN t_PEDIDO_CHAMADO_MENSAGEM tPCM ON" & _ 
		                           " (tU.usuario = tPCM.usuario_cadastro)" & _
                           " WHERE tPCM.id_chamado = " & vChamado(i).id_chamado

                corpo_mensagem = "O usuário " & usuario & " - " & x_usuario(usuario) & " finalizou o chamado referente ao pedido " & vChamado(i).pedido & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Finalizado por: " & usuario & " em " & Now & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Motivo da abertura: " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA, vChamado(i).motivo_abertura) & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Motivo da finalização: " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_FINALIZACAO, vChamado(i).motivo_finalizacao) & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Solução: " & vChamado(i).texto_finalizacao & chr(13) & chr(10)        
        
                if usuario = usuario_rx then
                    destinatario_to = obtem_email_usuario(usuario_tx)
                else
                    destinatario_to = obtem_email_usuario(usuario_rx)
                end if

                rs.Open s, cn
                do while Not rs.Eof

                    if Trim("" & rs("email")) <> usuario_email And Trim("" & rs("email")) <> destinatario_to then
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
                
                if destinatario_to <> "" then
                    EmailSndSvcGravaMensagemParaEnvio EMAILSNDSVC_REMETENTE__CHAMADOS_EM_PEDIDOS, _
                                                        usuario_email, _
                                                        destinatario_to, _
                                                        destinatario_cc, _
                                                        "", _
                                                        "Chamado finalizado referente ao pedido " & vChamado(i).pedido, _
                                                        corpo_mensagem, _
                                                        Now, _
                                                        id_email, _
                                                        msg_erro_grava_email
                end if                
                        
							log_via_vetor_carrega_do_recordset rs2, vLog2, campos_a_omitir
							s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
							grava_log usuario, "", vChamado(i).pedido, "", OP_LOG_PEDIDO_CHAMADO_FINALIZACAO, s_log
							end if
						end if  'if alerta = ""
					
					if rs2.State <> 0 then rs2.Close
					end if  'if Trim(vChamado(i).motivo_finalizacao)<>""
				end if  'if Trim(vChamado(i).id_chamado)<>""
			next
			
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
                if c_rel_origem = "ACOMPANHAMENTO_CHAMADOS" then
				    Response.Redirect("RelAcompanhamentoChamadosExec.asp?origem=A&rb_status=" & rb_status & "&rb_posicao=" & rb_posicao & "&c_motivo_abertura=" & c_motivo_abertura_filtro & "&c_dt_cad_chamado_inicio=" & c_dt_cad_chamado_inicio & "&c_dt_cad_chamado_termino=" & c_dt_cad_chamado_termino & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
                elseif c_rel_origem = "RELATORIO_CHAMADOS" then
				    Response.Redirect("RelPedidoChamado.asp?origem=A&rb_status=" & rb_status & "&blnHasDepto=" & blnHasDepto & "&c_depto=" & c_depto & "&c_motivo_abertura=" & c_motivo_abertura_filtro & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
                end if
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