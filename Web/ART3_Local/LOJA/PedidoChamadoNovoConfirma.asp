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
'	  P E D I D O C H A M A D O N O V O C O N F I R M A . A S P
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
	dim usuario, loja, usuario_email
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim alerta
	alerta=""
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    dim url_origem
    url_origem = Trim(Request("url_origem"))

	if Not operacao_permitida(OP_LJA_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim pedido_selecionado
	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim c_texto, c_contato, c_ddd_1, c_tel_1, c_ddd_2, c_tel_2, cod_motivo_abertura, c_depto, c_nivel_acesso_chamado
	c_texto = Trim(Request("c_texto"))
	c_contato = Trim(Request("c_contato"))
	c_ddd_1 = retorna_so_digitos(Trim(Request("c_ddd_1")))
	c_tel_1 = retorna_so_digitos(Trim(Request("c_tel_1")))
	c_ddd_2 = retorna_so_digitos(Trim(Request("c_ddd_2")))
	c_tel_2 = retorna_so_digitos(Trim(Request("c_tel_2")))
    cod_motivo_abertura = Trim(Request("motivo_chamado"))
    c_depto = Trim(Request("c_depto"))
    c_nivel_acesso_chamado = Trim(Request("c_nivel_acesso_chamado"))
	
	if len(c_texto) > MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS then
		alerta = "O tamanho do texto (" & Cstr(len(c_texto)) & ") excede o tamanho máximo permitido de " & Cstr(MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS) & " caracteres."
    elseif cod_motivo_abertura = "" then
        alerta = "Não foi informado o motivo da abertura do chamado."
    elseif c_depto = "" then
        alerta = "Não foi informado o departamento responsável pelo chamado."
    elseif c_nivel_acesso_chamado = "" then
        alerta = "Não foi informado o nível de acesso ao chamado."
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
	dim cn, rs, rs2
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

'	GRAVA O CHAMADO
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

	'	GERA O NSU PARA GRAVAR O CHAMADO
		dim intNsuNovoChamado
		if Not fin_gera_nsu(T_PEDIDO_CHAMADO, intNsuNovoChamado, msg_erro) then
			alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		else
			if intNsuNovoChamado <= 0 then
				alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovoChamado & ")"
				end if
			end if
		
		if alerta = "" then
			s = "SELECT * FROM t_PEDIDO_CHAMADO WHERE (id = -1)"
			rs.Open s, cn
			rs.AddNew 
			rs("id")=intNsuNovoChamado
			rs("pedido")=pedido_selecionado
			rs("usuario_cadastro")=usuario
            rs("cod_motivo_abertura")=cod_motivo_abertura
			rs("texto_chamado")=c_texto
            rs("nivel_acesso")=c_nivel_acesso_chamado
            rs("id_depto")=c_depto
            't_PEDIDO_CHAMADO.contato varchar(40) mas pode vir maior se o usuário usar o botão para preencher automaticamente. O browser nao respeita o maxlength quando o campo é alterado por javascript
			rs("contato") = Left(c_contato, 40)
			rs("ddd_1") = c_ddd_1
			rs("tel_1") = c_tel_1
			rs("ddd_2") = c_ddd_2
			rs("tel_2") = c_tel_2
            rs("loja") = loja
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
				
			if s_log <> "" then grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_CHAMADO_INCLUSAO, s_log
			end if
			
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~

            ' envia email para o usuário responsável pelo departamento
            dim corpo_mensagem, msg_erro_grava_email, id_email, destinatario
            destinatario = ""

            s = "SELECT usuario_responsavel," & _
                    " descricao AS descricao_depto," & _
                    " Coalesce((SELECT email FROM t_USUARIO WHERE usuario=tPCD.usuario_responsavel), '') AS email" & _ 
                " FROM t_PEDIDO_CHAMADO_DEPTO tPCD" & _
                " WHERE id = " & c_depto

            rs.Open s, cn
            if Not rs.Eof then
                corpo_mensagem = "Foi aberto um novo chamado para o pedido " & pedido_selecionado & " destinado ao seguinte departamento: " & UCase(rs("descricao_depto")) & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Aberto por: " & usuario & " - " & x_usuario(usuario) & " em " & Now & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Loja: " & CStr(loja) & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Nível de acesso: " & nivel_acesso_chamado_pedido_descricao(c_nivel_acesso_chamado) & "." & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Motivo da abertura: " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA, cod_motivo_abertura) & "." & chr(13) & chr(10) & chr(13) & chr(10)
                corpo_mensagem = corpo_mensagem & "Descrição: " & chr(13) & chr(10) & c_texto & "." & chr(13) & chr(10) & chr(13) & chr(10)

                if usuario_email = "" then
                    corpo_mensagem = corpo_mensagem & "---------------------------------------------------------------------------------------------------------"   & chr(13) & chr(10)
                    corpo_mensagem = corpo_mensagem & "E-MAIL ENVIADO AUTOMATICAMENTE PELO SISTEMA. NÃO RESPONDA ESTE E-MAIL, POIS ESTA CONTA NÃO É MONITORADA!!"                    
                    end if

                if UCase(usuario) <> UCase(Trim("" & rs("usuario_responsavel"))) then
                    destinatario = Trim("" & rs("email"))
                end if
                                
                if destinatario <> "" then
                EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__CHAMADOS_EM_PEDIDOS), _
                                                    usuario_email, _
                                                    destinatario, _
                                                    "", _
                                                    "", _
                                                    "Nova abertura de chamado para o pedido " & pedido_selecionado, _
                                                    corpo_mensagem, _
                                                    Now, _
                                                    id_email, _
                                                    msg_erro_grava_email
                end if
            end if

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