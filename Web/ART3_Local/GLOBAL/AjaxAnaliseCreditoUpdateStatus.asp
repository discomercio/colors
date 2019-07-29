<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     =============================================================
'	  AjaxAnaliseCreditoUpdateStatus.asp
'     =============================================================
'
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


' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
	dim s, usuario, msg_erro
	dim pedido, novo_status_analise_credito, id_operacao_origem, pagina_origem
	dim pedido_base, s_analise_credito_a, s_log, strRespSeparador

	usuario = Trim(Request("usuario"))
	pedido = Trim(Request("pedido"))
	novo_status_analise_credito = Trim(Request("novo_status_analise_credito"))
	id_operacao_origem = Trim(Request("id_operacao_origem"))
	pagina_origem = Trim(Request("pagina_origem"))

	if usuario = "" then
		Response.Write "Identificador do usuário não foi informado!!"
		Response.End
		end if
	
	if pedido = "" then
		Response.Write "Número do pedido não foi informado!!"
		Response.End
		end if
	
	if novo_status_analise_credito = "" then
		Response.Write "Novo status da análise de crédito não foi informado!!"
		Response.End
		end if
	
	if retorna_so_digitos(novo_status_analise_credito) = "" then
		Response.Write "Novo status da análise de crédito é inválido!!"
		Response.End
		end if
	
	dim cn, tPed
'	CONECTA COM O BANCO DE DADOS
	If Not bdd_conecta(cn) then
		msg_erro = erro_descricao(ERR_CONEXAO)
		Response.Write msg_erro
		Response.End
		end if
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = obtem_operacoes_permitidas_usuario(cn, usuario)
	
	if Not (operacao_permitida(OP_CEN_EDITA_ANALISE_CREDITO, s_lista_operacoes_permitidas) Or _
			operacao_permitida(OP_CEN_REL_ANALISE_CREDITO, s_lista_operacoes_permitidas)) then
		Response.Write "Usuário não possui permissão de acesso para alterar o status da análise de crédito!!"
		Response.End
		end if
	
	if Not cria_recordset_pessimista(tPed, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim strResp, blnHouveErro, blnSucesso
	strResp = ""
	blnHouveErro = False
	blnSucesso = False
	s_analise_credito_a = ""
	
	pedido_base = retorna_num_pedido_base(pedido)
	s = "SELECT * FROM t_PEDIDO WHERE pedido='" & pedido_base & "'"
	tPed.Open s, cn
	if tPed.Eof then
		blnHouveErro = True
		strResp = """msg_erro"" : """ & "Pedido " & pedido_base & " não foi localizado!" & """"
	else
		s_analise_credito_a = Trim("" & tPed("analise_credito"))
		end if
	
	if s_analise_credito_a = Trim("" & novo_status_analise_credito) then
		blnHouveErro = True
		strResp = """msg_erro"" : """ & _
					"Pedido " & pedido_base & " já está com o status de análise de crédito '" & _
					descricao_analise_credito(novo_status_analise_credito) & _
					"' desde " & formata_data_hora(tPed("analise_credito_data")) & _
					" cadastrado por " & Trim ("" & tPed("analise_credito_usuario")) & _
					""""
		end if
	
	s_log = ""
	if Not blnHouveErro then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
	'	ATUALIZA O STATUS DA ANÁLISE DE CRÉDITO
		s_log = s_log & " (status análise crédito anterior: " & _
				tPed("analise_credito") & " - " & descricao_analise_credito(tPed("analise_credito")) & _
				" em " & formata_data_hora(tPed("analise_credito_data")) & _
				" por " & Trim("" & tPed("analise_credito_usuario")) & _
				")"
		tPed("analise_credito")=CLng(novo_status_analise_credito)
		tPed("analise_credito_data")=Now
		tPed("analise_credito_usuario")=usuario
		
		s_log = "Alteração do status da análise de crédito do pedido " & _
				pedido_base & _
				" em " & _
				formata_data_hora(tPed("analise_credito_data")) & _
				" por " & _
				Trim("" & tPed("analise_credito_usuario")) & _
				" através de operação acionada na página" & _
				" '" & pagina_origem & "'" & _
				", id_operacao_origem=" & id_operacao_origem & _
				s_log
		tPed.Update
		if s_log <> "" then grava_log usuario, "", pedido, "", OP_LOG_PEDIDO_ALTERACAO, s_log
		blnSucesso = True
	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
		end if
	
	if blnSucesso then
		strResp = """resultado_operacao"" : """ & "OK" & """," & _
				  """id_operacao_origem"" : """ & id_operacao_origem & """," & _
				  """msg_erro"" : """ & "" & """"
	elseif blnHouveErro then
		strRespSeparador = ""
		if strResp <> "" then strRespSeparador = ","
		strResp = """resultado_operacao"" : """ & "ERRO" & """," & _
				  """id_operacao_origem"" : """ & id_operacao_origem & """" & _
				  strRespSeparador & _
				  strResp
		end if
	
	if tPed.State <> 0 then tPed.Close
	
	cn.Close
	set cn = nothing
	
	strResp = "{ " & _
			  strResp & _
			  " }"
	
'	ENVIA RESPOSTA
	Response.Write strResp
	Response.End
%>
