<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================
'	  Q U A D R O A V I S O L I D O . A S P
'     =============================================
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
	
	dim s, s_aux, usuario, loja, alerta, i
	
	usuario = trim(Session("usuario_atual"))
	loja = Session("loja_atual")
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim s_log
	s_log = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim aviso_selecionado
	Dim vAviso
	aviso_selecionado=trim(request("aviso_selecionado"))
	if aviso_selecionado<>"" then vAviso=split(aviso_selecionado,"|", -1)

	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	alerta = ""
	
	if Trim(aviso_selecionado) = "" then alerta = "NENHUM AVISO SELECIONADO."

	if alerta <> "" then erro_consistencia=True	
	
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	if alerta = "" then 
		cn.BeginTrans
		for i=Lbound(vAviso) to Ubound(vAviso)
			if Err = 0 then
			'	VERIFICA SE O AVISO AINDA EXISTE E OBTÉM INFO P/ LOG
				s = "SELECT * FROM t_AVISO WHERE (id='" & Trim(vAviso(i)) & "')"
				r.Open s, cn
				if Err = 0 then
					if Not r.EOF then
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & r("dt_ult_atualizacao") & " (id=" & Trim("" & r("id")) & ")"
						r.Close 
					'	VERIFICA SE JÁ NÃO ESTÁ MARCADO COMO LIDO
						s = "SELECT * FROM t_AVISO_LIDO WHERE (id='" & Trim(vAviso(i)) & "') AND (usuario='" & usuario & "')"
						r.Open s, cn
						if Err = 0 then
							if r.EOF then
							'	MARCA COMO LIDO
								r.AddNew
								if Err = 0 then
									r("id") = Trim(vAviso(i))
									r("usuario") = usuario
									r("data") = Now
									r.Update 
									end if
								end if
							end if
						r.Close 
						end if
					end if
				end if
			next
		
		if (Err=0) And (s_log<>"") then 
			s_log = "Leitura do aviso divulgado em: " & s_log
			grava_log usuario, loja, "", "", OP_LOG_AVISO_LIDO, s_log
			end if

		if Err = 0 then 
			cn.CommitTrans 
		else 
			cn.RollbackTrans
			end if

		if Err = 0 then 
			Response.Redirect("resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		else
			erro_fatal=True
			alerta = "FALHA AO ATUALIZAR BANCO DE DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
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


<body onload="bVOLTAR.focus();">
<center>
<br>

<!--  T E L A  -->

<p class="T">A V I S O</p>

<% 
	s = ""
	s_aux="'MtAviso'"
	if alerta <> "" then
		s = "<P style='margin:5px 2px 5px 2px;'>" & alerta & "</P>"
		s_aux="'MtAlerta'"
	else
		s = "AVISO(S) MARCADO(S) COMO LIDO(S) COM SUCESSO."
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<br><br>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
<%
	s="resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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