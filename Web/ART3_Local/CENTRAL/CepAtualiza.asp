<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===================================================================
'	  C E P A T U A L I Z A . A S P
'     ===================================================================
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
	dim cn, cn_cep, r, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not bdd_cep_conecta(cn_cep) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim criou_novo_reg
	Dim s_log
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim tipo_operacao, c_cep, c_uf, c_cidade, c_bairro
	dim c_tipo_logradouro, c_logradouro, c_complemento_logradouro
	tipo_operacao = request("tipo_operacao")
	c_cep = retorna_so_digitos(trim(Request.Form("c_cep")))
	c_uf = trim(Request.Form("c_uf"))
	c_cidade = trim(Request.Form("c_cidade"))
	c_bairro = trim(Request.Form("c_bairro"))
	c_tipo_logradouro = trim(Request.Form("c_tipo_logradouro"))
	c_logradouro = trim(Request.Form("c_logradouro"))
	c_complemento_logradouro = trim(Request.Form("c_complemento_logradouro"))
	
	dim erro_consistencia, erro_fatal
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if c_cep = "" then
		alerta="CEP NÃO FOI INFORMADO."
	elseif Len(c_cep) <> 8 then
		alerta="CEP COM TAMANHO INVÁLIDO."
	elseif c_uf = "" then
		alerta="UF NÃO FOI INFORMADA."
	elseif Not uf_ok(c_uf) then
		alerta="UF INVÁLIDA."
	elseif c_cidade = "" then
		alerta="CIDADE NÃO FOI INFORMADA."
	elseif c_bairro = "" then
		alerta="BAIRRO NÃO FOI INFORMADO."
	elseif c_logradouro = "" then
		alerta="LOGRADOURO NÃO FOI INFORMADO."
		end if
	
	if alerta <> "" then erro_consistencia=True	
		
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case tipo_operacao
		case OP_EXCLUI
		'	 =========
		'	INFO P/ LOG
			s="SELECT * FROM t_CEP_LOGRADOURO WHERE cep8_log = '" & c_cep & "'"
			if r.State <> 0 then r.Close
			r.Open s, cn_cep
			if Not r.EOF then
				log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
				s_log = log_via_vetor_monta_exclusao(vLog1)
				end if
			r.Close
				
		'	APAGA!!
		'	~~~~~~~~~~~~~~~~~
			cn_cep.BeginTrans
		'	~~~~~~~~~~~~~~~~~
			s="DELETE FROM t_CEP_LOGRADOURO WHERE cep8_log = '" & c_cep & "'"
			cn_cep.Execute(s)
			If Err = 0 then 
				if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_CEP_EXCLUSAO, s_log
			else
				erro_fatal=True
				alerta = "FALHA AO REMOVER O CEP (" & Cstr(Err) & ": " & Err.Description & ")."
				end if
			
			if alerta = "" then
			'	~~~~~~~~~~~~~~~~~~
				cn_cep.CommitTrans
			'	~~~~~~~~~~~~~~~~~~
				if Err <> 0 then 
					alerta=Cstr(Err) & ": " & Err.Description
					erro_fatal = True
					end if
			else
			'	~~~~~~~~~~~~~~~~~~~~
				cn_cep.RollbackTrans
			'	~~~~~~~~~~~~~~~~~~~~
				Err.Clear
				end if
				

		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			if alerta = "" then 
			'	~~~~~~~~~~~~~~~~~
				cn_cep.BeginTrans
			'	~~~~~~~~~~~~~~~~~
				s = "SELECT * FROM t_CEP_LOGRADOURO WHERE cep8_log = '" & c_cep & "'"
				r.Open s, cn_cep
				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("cep8_log")=c_cep
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					end if
					
				r("uf_log") = c_uf
				r("nome_local") = c_cidade
				r("extenso_bai") = c_bairro
				r("abrev_tipo") = c_tipo_logradouro
				r("nome_log") = c_logradouro
				r("comple_log") = c_complemento_logradouro
				r.Update

				If Err = 0 then 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then 
							grava_log usuario, "", "", "", OP_LOG_CEP_INCLUSAO, s_log
							end if
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if (s_log <> "") then 
							if s_log <> "" then s_log = "; " & s_log
							s_log="CEP=" & Trim("" & r("cep8_log")) & s_log
							grava_log usuario, "", "", "", OP_LOG_CEP_ALTERACAO, s_log
							end if
						end if
				else
					erro_fatal=True
					alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
					end if

				if alerta = "" then
				'	~~~~~~~~~~~~~~~~~~
					cn_cep.CommitTrans
				'	~~~~~~~~~~~~~~~~~~
					if Err <> 0 then 
						alerta=Cstr(Err) & ": " & Err.Description
						erro_fatal = True
						end if
				else
				'	~~~~~~~~~~~~~~~~~~~~
					cn_cep.RollbackTrans
				'	~~~~~~~~~~~~~~~~~~~~
					Err.Clear
					end if

				if r.State <> 0 then r.Close
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
		s = "<P style='margin:5px 2px 5px 2px;'>" & alerta & "</P>"
		s_aux="'MtAlerta'"
	else
		select case tipo_operacao
			case OP_INCLUI
				s = "CEP " & chr(34) & c_cep & chr(34) & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "CEP " & chr(34) & c_cep & chr(34) & " ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "CEP " & chr(34) & c_cep & chr(34) & " EXCLUÍDO COM SUCESSO."
			end select
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:600px;FONT-WEIGHT:bold;" align="CENTER"><%=s%></div>
<BR><BR>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
<%
	s="MenuCadastro.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>
</body>
</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing

	cn_cep.Close
	set cn_cep = nothing
%>