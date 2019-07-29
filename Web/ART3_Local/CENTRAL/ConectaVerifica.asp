<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<%
'     =====================
'	  CONECTAVERIFICA.ASP
'     =====================
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


' _____________________________________________________________________________________________
'
'						I N I C I A L I Z A     P Á G I N A     A S P
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear

	if Trim(Session("usuario_atual")) <> "" then Response.Redirect("sessaoviolada.asp")

	Session("usuario_a_checar") = filtra_nome_identificador(UCase( Left(trim(request("usuario")),MAX_TAMANHO_ID_USUARIO) ))
	Session("senha_a_checar") = filtra_nome_identificador(UCase( Left(trim(request("senha")),MAX_TAMANHO_SENHA) ))
	Session("verificar_quadro_avisos") = "S"
	Session("DataHoraLogon") = Now
	Session("DataHoraUltRefreshSession") = Now
	Response.Redirect("resumo.asp")

%>
