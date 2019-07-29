<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =========================================================
'	  O R D E M S E R V I C O A T U A L I Z A . A S P
'     =========================================================
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

	dim s, s_aux, i, intIndice, n, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_num_OS, s_chave_OS, c_obs_pecas_necessarias

'	OBTÉM DADOS DO FORMULÁRIO
	s_num_OS = Trim(Request.Form("c_num_OS"))
	s_chave_OS = normaliza_codigo(retorna_so_digitos(s_num_OS), TAM_MAX_NSU)
	
	c_obs_pecas_necessarias = Trim(Request.Form("c_obs_pecas_necessarias"))

	dim alerta
	alerta=""

	dim s_log_OS, s_log_OS_item
	dim vLog1()
	dim vLog2()
	dim campos_a_omitir

	campos_a_omitir = "|timestamp|excluido_status|"
	
	s_log_OS_item = ""

	dim v_OS_item
	redim v_OS_item(0)
	set v_OS_item(0) = New cl_ORDEM_SERVICO_ITEM
	n = Request.Form("c_descricao_volume").Count
	for i = 1 to n
		s=Trim(Request.Form("c_descricao_volume")(i))
		if s <> "" then
			if Trim(v_OS_item(ubound(v_OS_item)).descricao_volume) <> "" then
				redim preserve v_OS_item(ubound(v_OS_item)+1)
				set v_OS_item(ubound(v_OS_item)) = New cl_ORDEM_SERVICO_ITEM
				end if
			with v_OS_item(ubound(v_OS_item))
				.num_serie = Trim(Request.Form("c_num_serie")(i))
				.tipo = Trim(Request.Form("c_tipo")(i))
				.descricao_volume = Trim(Request.Form("c_descricao_volume")(i))
				.obs_problema = Trim(Request.Form("c_obs_problema")(i))
				end with
			end if
		next
	
	if alerta = "" then
		for i=Lbound(v_OS_item) to Ubound(v_OS_item)
			with v_OS_item(i)
				if Len(.obs_problema) > MAX_TAM_OS_OBS_PROBLEMA then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O campo 'Problema' do volume '" & .descricao_volume & "' excede o tamanho máximo de " & CStr(MAX_TAM_OS_OBS_PROBLEMA) & " caracteres."
					end if
				end with
			next
		
		if Len(c_obs_pecas_necessarias) > MAX_TAM_OS_OBS_PECAS_NECESSARIAS then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O campo 'Peças Necessárias' excede o tamanho máximo de " & Cstr(MAX_TAM_OS_OBS_PECAS_NECESSARIAS) & " caracteres."
			end if
		end if
	
	dim s_log_aux
	dim s_nome_campo
	dim s_editado, s_original
	dim r_editado, r_original
	set r_editado = New cl_ORDEM_SERVICO_ITEM
	set r_original = New cl_ORDEM_SERVICO_ITEM
	n = Request.Form("c_descricao_volume").Count
	for i = 1 to n
		with r_editado
			.num_serie = Trim(Request.Form("c_num_serie")(i))
			.tipo = Trim(Request.Form("c_tipo")(i))
			.descricao_volume = Trim(Request.Form("c_descricao_volume")(i))
			.obs_problema = Trim(Request.Form("c_obs_problema")(i))
			end with

		with r_original
			.num_serie = Trim(Request.Form("c_num_serie_original")(i))
			.tipo = Trim(Request.Form("c_tipo_original")(i))
			.descricao_volume = Trim(Request.Form("c_descricao_volume_original")(i))
			.obs_problema = Trim(Request.Form("c_obs_problema_original")(i))
			end with
		
		s_log_aux = ""
		
		s_nome_campo = "num_serie"
		s_editado = r_editado.num_serie
		s_original = r_original.num_serie
		if s_editado <> s_original then 
			if s_log_aux <> "" then s_log_aux = s_log_aux & "; "
			if s_editado = "" then s_editado = chr(34) & chr(34)
			if s_original = "" then s_original = chr(34) & chr(34)
			s_log_aux = s_log_aux & s_nome_campo & ": " & s_original & " => " & s_editado
			end if

		s_nome_campo = "tipo"
		s_editado = r_editado.tipo
		s_original = r_original.tipo
		if s_editado <> s_original then 
			if s_log_aux <> "" then s_log_aux = s_log_aux & "; "
			if s_editado = "" then s_editado = chr(34) & chr(34)
			if s_original = "" then s_original = chr(34) & chr(34)
			s_log_aux = s_log_aux & s_nome_campo & ": " & s_original & " => " & s_editado
			end if
		
		s_nome_campo = "descricao_volume"
		s_editado = r_editado.descricao_volume
		s_original = r_original.descricao_volume
		if s_editado <> s_original then 
			if s_log_aux <> "" then s_log_aux = s_log_aux & "; "
			if s_editado = "" then s_editado = chr(34) & chr(34)
			if s_original = "" then s_original = chr(34) & chr(34)
			s_log_aux = s_log_aux & s_nome_campo & ": " & s_original & " => " & s_editado
			end if

		s_nome_campo = "obs_problema"
		s_editado = r_editado.obs_problema
		s_original = r_original.obs_problema
		if s_editado <> s_original then 
			if s_log_aux <> "" then s_log_aux = s_log_aux & "; "
			if s_editado = "" then s_editado = chr(34) & chr(34)
			if s_original = "" then s_original = chr(34) & chr(34)
			s_log_aux = s_log_aux & s_nome_campo & ": " & s_original & " => " & s_editado
			end if

		if s_log_aux <> "" then
			if s_log_OS_item <> "" then s_log_OS_item = s_log_OS_item & chr(13)
			s_log_OS_item = s_log_OS_item & "Volume " & CStr(i) & ": " & s_log_aux
			end if
		next
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_pessimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
		
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		' ATUALIZA A ORDEM DE SERVIÇO
		s = "SELECT * FROM t_ORDEM_SERVICO WHERE ordem_servico='" & s_chave_OS & "'"
		if rs.State <> 0 then rs.Close
		rs.Open s, cn
		if Err <> 0 then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		elseif rs.EOF then
			alerta = "Ordem de serviço " & s_chave_OS & " não foi encontrado."
		else
			log_via_vetor_carrega_do_recordset rs, vLog1, campos_a_omitir
			rs("obs_pecas_necessarias") = c_obs_pecas_necessarias
			rs.Update
			log_via_vetor_carrega_do_recordset rs, vLog2, campos_a_omitir
			s_log_OS = log_via_vetor_monta_alteracao(vLog1, vLog2)
			end if

		if Err <> 0 then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			end if

		s = "UPDATE t_ORDEM_SERVICO_ITEM SET excluido_status = 1 WHERE ordem_servico = '" & s_chave_OS & "'"
		cn.Execute(s)
		if Err <> 0 then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			end if

		for i=Lbound(v_OS_item) to Ubound(v_OS_item)
			with v_OS_item(i)
				intIndice = renumera_com_base1(Lbound(v_OS_item), i)
				s="SELECT * FROM t_ORDEM_SERVICO_ITEM WHERE (ordem_servico='" & s_chave_OS & "') AND (sequencia = " & Cstr(intIndice) & ")"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				If Not rs.Eof then
					rs("excluido_status") = 0
				else
					rs.AddNew
					rs("ordem_servico") = s_chave_OS
					rs("sequencia") = intIndice
					end if

				rs("num_serie") = .num_serie
				rs("tipo") = .tipo
				rs("descricao_volume") = .descricao_volume
				rs("obs_problema") = .obs_problema
				rs.Update

				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
				end with
			next

		if Err = 0 then
			s = "DELETE FROM t_ORDEM_SERVICO_ITEM WHERE (ordem_servico = '" & s_chave_OS & "') AND (excluido_status <> 0)"
			cn.Execute(s)
			end if
		
	'	GRAVA LOG DA NOVA ORDEM DE SERVIÇO
		if (s_log_OS <> "") And (s_log_OS_item <> "") then s_log_OS = s_log_OS & chr(13)
		s_log_OS = s_log_OS & s_log_OS_item
		if s_log_OS <> "" then 
			s_log_OS = "O.S = " & s_chave_OS & "; " & s_log_OS
			grava_log usuario, "", "", "", OP_LOG_ORDEM_SERVICO_ALTERACAO, s_log_OS
			end if

	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~

		if Err=0 then 
			Response.Redirect("OrdemServico.asp?num_OS=" & s_num_OS & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		else
			alerta=Cstr(Err) & ": " & Err.Description
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
<table cellspacing="0">
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