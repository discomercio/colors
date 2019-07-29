<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<%
'     =====================
'	  SESSAOVIOLADA.ASP
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

'	ATUALIZA BANCO DE DADOS
	if Trim(Session("usuario_atual")) <> "" then
		dim cn
		dim strSQL
		if bdd_conecta(cn) then
			strSQL = "UPDATE t_USUARIO SET" & _
						" SessionCtrlTicket = NULL," & _
						" SessionCtrlLoja = NULL," & _
						" SessionCtrlModulo = NULL," & _
						" SessionCtrlDtHrLogon = NULL" & _
					" WHERE" & _
						" usuario = '" & Trim(Session("usuario_atual")) & "'"
			cn.Execute(strSQL)
			
			strSQL = "UPDATE t_SESSAO_HISTORICO SET" & _
						" DtHrTermino = " & bd_formata_data_hora(Now) & _
					 " WHERE" & _
						" usuario = '" & QuotedStr(Trim("" & Session("usuario_atual"))) & "'" & _
						" AND DtHrInicio >= " & bd_formata_data_hora(Now-1) & _
						" AND SessionCtrlTicket = '" & Trim(Session("SessionCtrlTicket")) & "'"
			cn.Execute(strSQL)

			cn.Close
			end if
		set cn = nothing
		end if

'	ENCERRA A SESSÃO
	Session("usuario_atual") = " "
	Session("senha_atual") = " "
	Session.Abandon

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

<script language="JavaScript" type="text/javascript">
window.focus();
</script>

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

<body>
<center>
<br>
<!--  A V I S O  -->
<p class="T">A V I S O</p>
<div class="ALERTA" style="width:300px" align="center">
	<p style="margin:5px 2px 5px 2px;"
>ERRO AO INICIAR A SESSÃO!!<br
><br
>JÁ EXISTE OUTRO USUÁRIO CONECTADO OU A SESSÃO ANTERIOR NÃO FOI ENCERRADA CORRETAMENTE.</p>
	</div>
<br><br>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
	<td align="center">
		<input name="bFECHAR" id="bFECHAR" type="button" class="Botao" 
			   value="ENCERRA SESSÃO" title="encerra a sessão" onclick="window.location='sessaoencerra.asp';">
		</td>
</tr>
</table>

</center>
</body>

</html>
