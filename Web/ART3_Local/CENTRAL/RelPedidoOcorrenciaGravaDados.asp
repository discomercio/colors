<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================================================
'	  RelPedidoOcorrenciaGravaDados.asp
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

	class cl_TIPO_GRAVA_REL_OCORRENCIA
		dim id_ocorrencia
		dim pedido
		dim mensagem
		dim tipo_ocorrencia
		dim texto_finalizacao
		end class
		
	dim s, msg_erro
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim alerta
	alerta=""
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_REL_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim rb_status, c_qtde_ocorrencias, intQtdeOcorrencias, vOcorrencia, c_loja
	rb_status=Trim(Request("rb_status"))
	c_loja=Trim(Request("c_loja"))
	c_qtde_ocorrencias=Trim(Request("c_qtde_ocorrencias"))
	intQtdeOcorrencias=CInt(c_qtde_ocorrencias)
	
	redim vOcorrencia(0)
	set vOcorrencia(Ubound(vOcorrencia)) = new cl_TIPO_GRAVA_REL_OCORRENCIA
	vOcorrencia(Ubound(vOcorrencia)).id_ocorrencia = ""
	
	dim i
	dim c_id_ocorrencia, c_pedido, c_nova_msg, c_tipo_ocorrencia, c_solucao
	for i = 1 to intQtdeOcorrencias
		c_id_ocorrencia = Trim(Request.Form("c_id_ocorrencia_" & Cstr(i)))
		c_pedido = Trim(Request.Form("c_pedido_" & Cstr(i)))
		c_nova_msg = Trim(Request.Form("c_nova_msg_" & Cstr(i)))
		c_tipo_ocorrencia = Trim(Request.Form("c_tipo_ocorrencia_" & Cstr(i)))
		c_solucao = Trim(Request.Form("c_solucao_" & Cstr(i)))
		if (c_id_ocorrencia<>"") And ( (c_nova_msg<>"") Or (c_tipo_ocorrencia<>"") Or (c_solucao<>"") ) then
			if vOcorrencia(Ubound(vOcorrencia)).id_ocorrencia <> "" then
				redim preserve vOcorrencia(Ubound(vOcorrencia)+1)
				set vOcorrencia(Ubound(vOcorrencia)) = new cl_TIPO_GRAVA_REL_OCORRENCIA
				end if
			vOcorrencia(Ubound(vOcorrencia)).id_ocorrencia = c_id_ocorrencia
			vOcorrencia(Ubound(vOcorrencia)).pedido = c_pedido
			vOcorrencia(Ubound(vOcorrencia)).mensagem = c_nova_msg
			vOcorrencia(Ubound(vOcorrencia)).tipo_ocorrencia = c_tipo_ocorrencia
			vOcorrencia(Ubound(vOcorrencia)).texto_finalizacao = c_solucao
			end if
		next

	for i=Lbound(vOcorrencia) to Ubound(vOcorrencia)
		if Trim(vOcorrencia(i).id_ocorrencia)<>"" then
			if len(vOcorrencia(i).mensagem) > MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O tamanho do texto da mensagem (" & Cstr(len(vOcorrencia(i).mensagem)) & ")  da ocorrência do pedido " & vOcorrencia(i).pedido & " excede o tamanho máximo permitido de " & Cstr(MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS) & " caracteres."
			elseif len(vOcorrencia(i).texto_finalizacao) > MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O tamanho do texto descrevendo a solução (" & Cstr(len(vOcorrencia(i).texto_finalizacao)) & ")  da ocorrência do pedido " & vOcorrencia(i).pedido & " excede o tamanho máximo permitido de " & Cstr(MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS) & " caracteres."
				end if

			if (Trim(vOcorrencia(i).tipo_ocorrencia)<>"") Or (Trim(vOcorrencia(i).texto_finalizacao)<>"") then
				if Trim(vOcorrencia(i).tipo_ocorrencia)="" then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Não foi selecionado o tipo de ocorrência para o pedido " & vOcorrencia(i).pedido & "!!<br>Ao finalizar uma ocorrência, é necessário informar o tipo de ocorrência e o texto descrevendo a solução."
				elseif Trim(vOcorrencia(i).texto_finalizacao)="" then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Não foi informado o texto descrevendo a solução da ocorrência do pedido " & vOcorrencia(i).pedido & "!!<br>Ao finalizar uma ocorrência, é necessário informar o tipo de ocorrência e o texto descrevendo a solução."
					end if
				end if
			end if
		next
	
	
	dim intNsuNovaOcorrenciaMensagem
	dim campos_a_omitir
	dim vLog(), vLog1(), vLog2()
	dim s_log
	s_log = ""
	campos_a_omitir = "|dt_cadastro|dt_hr_cadastro|finalizado_data|finalizado_data_hora|"


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

'	GRAVA A MENSAGEM P/ ESTA OCORRÊNCIA
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

		if Not cria_recordset_pessimista(rs2, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if
			
		for i=Lbound(vOcorrencia) to Ubound(vOcorrencia)
			if Trim(vOcorrencia(i).id_ocorrencia)<>"" then
			'	TEM MENSAGEM NOVA P/ GRAVAR?
				if Trim(vOcorrencia(i).mensagem)<>"" then
				'	GERA O NSU PARA GRAVAR A MENSAGEM P/ ESTA OCORRÊNCIA
					if Not fin_gera_nsu(T_PEDIDO_OCORRENCIA_MENSAGEM, intNsuNovaOcorrenciaMensagem, msg_erro) then
						alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
					else
						if intNsuNovaOcorrenciaMensagem <= 0 then
							alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovaOcorrenciaMensagem & ")"
							end if
						end if
					
					if alerta = "" then
						s = "SELECT * FROM t_PEDIDO_OCORRENCIA_MENSAGEM WHERE (id = -1)"
						rs.Open s, cn
						rs.AddNew 
						rs("id")=intNsuNovaOcorrenciaMensagem
						rs("id_ocorrencia")=CLng(vOcorrencia(i).id_ocorrencia)
						rs("usuario_cadastro")=usuario
						rs("fluxo_mensagem") = COD_FLUXO_MENSAGEM_OCORRENCIAS_EM_PEDIDOS__CENTRAL_PARA_LOJA
						rs("texto_mensagem")=Trim(vOcorrencia(i).mensagem)
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
							
						if s_log <> "" then grava_log usuario, "", vOcorrencia(i).pedido, "", OP_LOG_PEDIDO_OCORRENCIA_MENSAGEM_INCLUSAO, s_log
						end if
					end if  'if Trim(vOcorrencia(i).mensagem)<>""
					
			'	FINALIZA A OCORRÊNCIA?
				if Trim(vOcorrencia(i).tipo_ocorrencia)<>"" then
					s = "SELECT * FROM t_PEDIDO_OCORRENCIA WHERE (id = " & vOcorrencia(i).id_ocorrencia & ")"
					rs2.Open s, cn
					if rs2.Eof then
						alerta = "Registro da ocorrência (id=" & vOcorrencia(i).id_ocorrencia & ") do pedido " & vOcorrencia(i).pedido & " não foi localizado no banco de dados!!"
						exit for
						end if
					
					if CInt(rs2("finalizado_status"))<>0 then
						alerta = "Registro da ocorrência (id=" & vOcorrencia(i).id_ocorrencia & ") do pedido " & vOcorrencia(i).pedido & " já se encontra finalizado!!"
						exit for
						end if
						
					if alerta = "" then
						log_via_vetor_carrega_do_recordset rs2, vLog1, campos_a_omitir
						rs2("finalizado_status")=1
						rs2("finalizado_usuario")=usuario
						rs2("finalizado_data")=Date
						rs2("finalizado_data_hora")=Now
						rs2("tipo_ocorrencia")=vOcorrencia(i).tipo_ocorrencia
						rs2("texto_finalizacao")=vOcorrencia(i).texto_finalizacao
						rs2.Update
						
						if Err <> 0 then
							alerta = Cstr(Err) & ": " & Err.Description
						else
							log_via_vetor_carrega_do_recordset rs2, vLog2, campos_a_omitir
							s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
							grava_log usuario, "", vOcorrencia(i).pedido, "", OP_LOG_PEDIDO_OCORRENCIA_FINALIZACAO, s_log
							end if
						end if  'if alerta = ""
					
					if rs2.State <> 0 then rs2.Close
					end if  'if Trim(vOcorrencia(i).tipo_ocorrencia)<>""
				end if  'if Trim(vOcorrencia(i).id_ocorrencia)<>""
			next
			
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("RelPedidoOcorrencia.asp?origem=A&rb_status=" & rb_status & "&c_loja=" & c_loja & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
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