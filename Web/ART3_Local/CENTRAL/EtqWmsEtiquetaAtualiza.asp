<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  EtqWmsEtiquetaAtualiza.asp
'     ===========================================
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
'			I N I C I A L I Z A     P Á G I N A     A S P     N O     S E R V I D O R
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
	dim s, s_aux, usuario, alerta
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim s_log
	s_log = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim blnExecutarUpdate
	dim c_id_wms_etq_n1, c_id_wms_etq_n2, c_id_wms_etq_n3, c_obs2, c_obs2_original, c_obs3, c_obs3_original, c_transportadora, c_transportadora_original
	c_id_wms_etq_n1=retorna_so_digitos(Trim(Request.Form("c_id_wms_etq_n1")))
	c_id_wms_etq_n2=retorna_so_digitos(Trim(Request.Form("c_id_wms_etq_n2")))
	c_id_wms_etq_n3=retorna_so_digitos(Trim(Request.Form("c_id_wms_etq_n3")))
	c_obs2=Trim(Request.Form("c_obs2"))
	c_obs2_original=Trim(Request.Form("c_obs2_original"))
	c_obs3=Trim(Request.Form("c_obs3"))
	c_obs3_original=Trim(Request.Form("c_obs3_original"))
	c_transportadora=Trim(Request.Form("c_transportadora"))
	c_transportadora_original=Trim(Request.Form("c_transportadora_original"))

	if (c_id_wms_etq_n1 = "") Or (converte_numero(c_id_wms_etq_n1) = 0) then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICADOR_NAO_FORNECIDO)
	if (c_id_wms_etq_n2 = "") Or (converte_numero(c_id_wms_etq_n2) = 0) then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICADOR_NAO_FORNECIDO)
	if (c_id_wms_etq_n3 = "") Or (converte_numero(c_id_wms_etq_n3) = 0) then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICADOR_NAO_FORNECIDO)
	
	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if (c_obs2=c_obs2_original) And (c_obs3=c_obs3_original) And (c_transportadora=c_transportadora_original) then
		alerta="NENHUMA ALTERAÇÃO FOI REALIZADA"
		end if
	
	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)


'	EXECUTA OPERAÇÃO NO BD
	if alerta = "" then
		s = "SELECT " & _
				"*" & _
			" FROM t_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO" & _
			" WHERE" & _
				" (id = " & c_id_wms_etq_n2 & ")"
		r.Open s, cn
		if r.EOF then
			alerta="Não foi encontrado o registro no banco de dados (id=" & c_id_wms_etq_n2 & ")"
			end if
		
		if alerta = "" then
			blnExecutarUpdate = False
			
			if Trim("" & r("obs_2")) <> c_obs2 then
				if s_log <> "" then s_log = s_log & "; "
				s_log = s_log & "obs_2: " & formata_texto_log(Trim("" & r("obs_2"))) & " => " & formata_texto_log(c_obs2)
				r("obs_2") = c_obs2
				blnExecutarUpdate = True
				end if
			
			if Trim("" & r("obs_3")) <> c_obs3 then
				if s_log <> "" then s_log = s_log & "; "
				s_log = s_log & "obs_3: " & formata_texto_log(Trim("" & r("obs_3"))) & " => " & formata_texto_log(c_obs3)
				r("obs_3") = c_obs3
				blnExecutarUpdate = True
				end if
			
			if Trim("" & r("transportadora_id")) <> c_transportadora then
				if s_log <> "" then s_log = s_log & "; "
				s_log = s_log & "transportadora_id: " & formata_texto_log(Trim("" & r("transportadora_id"))) & " => " & formata_texto_log(c_transportadora)
				r("transportadora_id") = c_transportadora
				blnExecutarUpdate = True
				end if
			
			if blnExecutarUpdate then r.Update

			If Err = 0 then 
				if s_log <> "" then 
					s_log="id_wms_etq_n1=" & c_id_wms_etq_n1 & ", id_wms_etq_n2=" & c_id_wms_etq_n2 & ", id_wms_etq_n3=" & c_id_wms_etq_n3 & "; " & s_log
					grava_log usuario, "", "", "", OP_LOG_ETQWMS_ETIQUETA_ALTERACAO, s_log
					end if
			else
				erro_fatal=True
				alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
				end if
			end if
		
		r.Close
		set r = nothing
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
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>



<!-- C A S C A D I N G   S T Y L E   S H E E T

	 CCCCCCC    SSSSSSS    SSSSSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	CCC        SSS        SSS
	CCC         SSSS       SSSS
	CCC            SSSS       SSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	 CCCCCCC    SSSSSSS    SSSSSSS
-->

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">


<body onload="bVOLTAR.focus();">
<center>
<br>

<!--  T E L A  -->

<p class="T">A V I S O</p>

<% 
	s = ""
	s_aux="'MtAviso'"
	if alerta <> "" then
		s = "<p style='margin:5px 2px 5px 2px;'>" & alerta & "</p>"
		s_aux="'MtAlerta'"
	else
		s = "DADOS ALTERADOS COM SUCESSO"
		if s <> "" then s="<p style='margin:5px 2px 5px 2px;'>" & s & "</p>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<br><br>

<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
<%
	s="EtqWmsEtiquetaObtemId.asp" & "?url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>
</body>
</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>