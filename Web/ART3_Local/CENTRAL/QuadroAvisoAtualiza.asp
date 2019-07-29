<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================
'	  Q U A D R O A V I S O A T U A L I Z A . A S P
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
	
	dim s, s_aux, usuario, alerta
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	Dim criou_novo_reg, editou_texto
	Dim s_log
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, aviso_selecionado, s_aviso, c_destinatario
	operacao_selecionada=request("operacao_selecionada")
	aviso_selecionado=trim(request("aviso_selecionado"))
	s_aviso=Trim(request("mensagem"))
	c_destinatario = Trim(request("c_destinatario"))
	
	if aviso_selecionado = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	
	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	alerta = ""
	
	if s_aviso = "" then
		alerta="NÃO HÁ TEXTO NA MENSAGEM DE AVISO."
		end if
	
	if alerta = "" then
		if (operacao_selecionada=OP_INCLUI) OR (operacao_selecionada=OP_CONSULTA) then
			if c_destinatario <> "" then
				s = "SELECT * FROM t_LOJA WHERE (loja='" & c_destinatario & "')"
				r.Open s, cn
				if r.Eof then
					alerta = "LOJA " & c_destinatario & " NÃO ESTÁ CADASTRADA."
					end if
				end if
			end if
		end if
	
	if alerta <> "" then erro_consistencia=True
	
	if r.State <> 0 then r.Close
	Err.Clear
	
'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
		'	INFO P/ LOG
			s="SELECT * FROM t_AVISO WHERE id = '" & aviso_selecionado & "'"
			r.Open s, cn
			if Not r.EOF then
				log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
				s_log = log_via_vetor_monta_exclusao(vLog1)
				end if
			r.Close
			
		'	APAGA!!
			cn.BeginTrans
			s="DELETE FROM t_AVISO_LIDO WHERE id = '" & aviso_selecionado & "'"
			cn.Execute(s)
			if Err = 0 then
				s="DELETE FROM t_AVISO WHERE id = '" & aviso_selecionado & "'"
				cn.Execute(s)
				If Err = 0 then 
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_AVISO_EXCLUSAO, s_log
				else
					erro_fatal=True
					alerta = "FALHA AO REMOVER O AVISO (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
			else
				erro_fatal=True
				alerta = "FALHA AO REMOVER O AVISO (" & Cstr(Err) & ": " & Err.Description & ")."
				end if
					
			if Err = 0 then
				cn.CommitTrans
			else
				cn.RollbackTrans
				if alerta = "" then alerta = "FALHA AO REMOVER O AVISO (" & Cstr(Err) & ": " & Err.Description & ")."
				end if


		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			if alerta = "" then 
				editou_texto=False
				s = "SELECT * FROM t_AVISO WHERE id = '" & aviso_selecionado & "'"
				r.Open s, cn
				if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				cn.BeginTrans
				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("id")=aviso_selecionado
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					end if

				if s_aviso <> "" then
					if IsNull(r("mensagem")) then 
						editou_texto=True
					else
						if r("mensagem") <> s_aviso then editou_texto=True
						end if
					r("mensagem")=s_aviso
				else
					if Not IsNull(r("mensagem")) then editou_texto = True
					r("mensagem")=Null
					end if
					
				if Trim("" & r("destinatario")) <> c_destinatario then
					editou_texto = True
					r("destinatario") = c_destinatario
					end if
					
				if editou_texto then
					r("usuario")=usuario
					r("dt_ult_atualizacao")=Now
					end if
				
				r.Update

			'	SE ALTEROU O TEXTO, FORÇA QUE OS USUÁRIOS LEIAM NOVAMENTE O AVISO
				if (Err = 0) And editou_texto And (operacao_selecionada <> OP_INCLUI) then
					s="DELETE FROM t_AVISO_LIDO WHERE id='" & aviso_selecionado & "'"
					cn.Execute(s)
					end if
					
				If Err = 0 then 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_AVISO_INCLUSAO, s_log
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if s_log <> "" then 
							s_log="aviso=" & Trim("" & r("id")) & "; " & s_log
							grava_log usuario, "", "", "", OP_LOG_AVISO_ALTERACAO, s_log
							end if
						end if
				else
					erro_fatal=True
					alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				
				if Err = 0 then
					cn.CommitTrans
				else
					cn.RollbackTrans
					if alerta = "" then alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				
				r.Close
				set r = nothing
				end if

		case else
		'	 ====
			alerta="OPERAÇÃO INVÁLIDA."
			
		end select
		
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
		s = "<p style='margin:5px 2px 5px 2px;'>" & alerta & "</p>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "AVISO CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "AVISO ATUALIZADO COM SUCESSO."
			case OP_EXCLUI
				s = "AVISO EXCLUÍDO COM SUCESSO."
			end select
		
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<br><br>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
<%
	s="quadroaviso.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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