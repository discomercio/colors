<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  FinCadUnidadeNegocioAtualiza.asp
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
	
	dim intNsuNovo
	dim s, s_aux, usuario, alerta
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim criou_novo_reg
	Dim s_log
	Dim campos_a_omitir, campos_a_omitir_exclusao
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = "|dt_cadastro|dt_hr_cadastro|"
	campos_a_omitir_exclusao = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, s_id, s_apelido, s_descricao, s_st_ativo
	operacao_selecionada=Request.Form("operacao_selecionada")
	s_id=retorna_so_digitos(Trim(Request.Form("id_selecionado")))
	s_apelido=Trim(Request.Form("c_apelido"))
	s_descricao=Trim(Request.Form("c_descricao"))
	s_st_ativo=Trim(Request.Form("rb_st_ativo"))

	if operacao_selecionada <> OP_INCLUI then
		if converte_numero(s_id) <= 0 then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
		end if
	
	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if (operacao_selecionada <> OP_INCLUI) And (s_id = "") then
		alerta="NÚMERO DE IDENTIFICAÇÃO INVÁLIDO."
	elseif s_apelido = "" then
		alerta="INFORME O APELIDO DA UNIDADE DE NEGÓCIOS."
	elseif s_descricao = "" then
		alerta="PREENCHA A DESCRIÇÃO DA UNIDADE DE NEGÓCIOS."
	elseif s_st_ativo = "" then
		alerta="INFORME O STATUS DA UNIDADE DE NEGÓCIOS."
		end if
	
	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			s = "SELECT" & _
					" TOP 1 *" & _
				" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM" & _
				" WHERE" & _
					" (id_unidade_negocio = " & s_id & ")"
			r.Open s, cn
			if Not r.Eof then
				erro_fatal=True
				alerta = "REGISTRO NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE RATEIO."
				end if
			r.Close 

			if Not erro_fatal then
				s = "SELECT" & _
						" TOP 1 *" & _
					" FROM t_FIN_UNIDADE_NEGOCIO_HISTORICO_RATEIO_ITEM" & _
					" WHERE" & _
						" (id_unidade_negocio = " & s_id & ")"
				r.Open s, cn
				if Not r.Eof then
					erro_fatal=True
					alerta = "REGISTRO NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE HISTÓRICO DO RATEIO."
					end if
				r.Close 
				end if

			if Not erro_fatal then
			'	INFO P/ LOG
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_UNIDADE_NEGOCIO" & _
					" WHERE" & _
						" (id = " & s_id & ")"
				r.Open s, cn
				if Not r.EOF then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir_exclusao
					s_log = log_via_vetor_monta_exclusao(vLog1)
					end if
				r.Close
			
			'	APAGA!!
				s = "DELETE" & _
					" FROM t_FIN_UNIDADE_NEGOCIO" & _
					" WHERE" &  _
						" (id = " & s_id & ")"
				cn.Execute(s)
				If Err = 0 then
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_UNIDADE_NEGOCIO_EXCLUSAO, s_log
				else
					erro_fatal=True
					alerta = "FALHA AO EXCLUIR O REGISTRO (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if


		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			if alerta = "" then
				if operacao_selecionada = OP_INCLUI then
				'	GERA O NSU PARA O NOVO REGISTRO
					if Not fin_gera_nsu(T_FIN_UNIDADE_NEGOCIO, intNsuNovo, msg_erro) then
						alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
					else
						if intNsuNovo <= 0 then
							alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovo & ")"
							end if
						end if
					
					if alerta = "" then
						s_id = Cstr(intNsuNovo)
						end if
					end if
				end if
			
			if alerta = "" then
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_UNIDADE_NEGOCIO" & _
					" WHERE" & _
						 " (id = " & s_id & ")"
				r.Open s, cn
				if r.EOF then
					r.AddNew 
					criou_novo_reg = True
					r("id") = CLng(s_id)
					r("usuario_cadastro") = usuario
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					end if
				
				r("apelido")=s_apelido
				r("descricao")=s_descricao
				r("st_ativo")=CLng(s_st_ativo)
				r("dt_ult_atualizacao")=Date
				r("dt_hr_ult_atualizacao")=Now
				r("usuario_ult_atualizacao")=usuario
				
				r.Update

				If Err = 0 then 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_UNIDADE_NEGOCIO_INCLUSAO, s_log
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if s_log <> "" then 
							s_log="Id=" & Trim("" & r("id")) & "; " & s_log
							grava_log usuario, "", "", "", OP_LOG_UNIDADE_NEGOCIO_ALTERACAO, s_log
							end if
						end if
				else
					erro_fatal=True
					alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
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
		s = "<P style='margin:5px 2px 5px 2px;'>" & alerta & "</P>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "REGISTRO (ID=" & chr(34) & s_id & chr(34) & ") CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "REGISTRO (ID=" & chr(34) & s_id & chr(34) & ") ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "REGISTRO (ID=" & chr(34) & s_id & chr(34) & ") EXCLUÍDO COM SUCESSO."
			end select			
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;FONT-WEIGHT:bold;" align="CENTER"><%=s%></div>
<BR><BR>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
<%
	s="FinCadUnidadeNegocioMenu.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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
%>