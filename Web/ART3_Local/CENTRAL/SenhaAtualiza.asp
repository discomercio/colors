<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================
'	  S E N H A A T U A L I Z A . A S P
'     =================================
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
	
	dim s, msg_erro, usuario, senha, novasenha, senha_cripto, alerta, chave
	alerta = ""
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
'	OBTEM DADOS DO FORM ANTERIOR
	senha = UCase(trim(request("senha")))
	novasenha = UCase(trim(request("novasenha")))
	
	usuario = trim(Session("usuario_atual"))
	if (usuario = "") then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
		end if

	if senha <> UCase(trim(Session("senha_atual"))) then alerta = "SENHA ATUAL INVÁLIDA."
	
'	ALTERA A SENHA NO BD
	if alerta = "" then 
		chave = gera_chave(FATOR_BD)
		codifica_dado novasenha, senha_cripto, chave
		Err.Clear
		if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
		s = "SELECT datastamp, dt_ult_alteracao_senha, dt_ult_atualizacao, senha FROM t_USUARIO WHERE usuario = '" & UCase(usuario) & "'"
		r.Open s, cn
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		
		if Not r.eof then 
			r("datastamp") = senha_cripto
			r("dt_ult_alteracao_senha") = date
			r("dt_ult_atualizacao") = Now
			r("senha") = gera_senha_aleatoria
			r.Update
			If Err = 0 then 
				grava_log usuario, "", "", "", OP_LOG_SENHA_ALTERACAO, "SENHA ALTERADA PELO USUÁRIO"
			else
				alerta = "NÃO FOI POSSÍVEL ALTERAR A SENHA."
				end if
			end if
		r.Close
		set r = nothing
		end if
	
	if alerta = "" then 
		Session("senha_atual") = novasenha
		Response.Redirect("resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
	else 
		Response.Redirect("aviso.asp?id=" & ERR_SENHA_INVALIDA)
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
	<title>SENHA</title>
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">



<body>

<!--  T E L A  -->

<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO"><span class="C">&nbsp;</span></p></td>
</tr>
</table>

<br><br>
<center>

<%= "<p class='ALERTA'>" & alerta & "</p>"%>

<br>
<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
	<td align="center"><a href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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