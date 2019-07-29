<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  FinCadPlanoContasContaAtualiza.asp
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

	Dim criou_novo_reg
	Dim s_log
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, s_id, s_descricao, s_st_ativo, s_grupo, s_natureza
	operacao_selecionada=Request.Form("operacao_selecionada")
	s_id=retorna_so_digitos(Trim(Request.Form("id_selecionado")))
	s_descricao=Trim(Request.Form("c_descricao"))
	s_st_ativo=Trim(Request.Form("rb_st_ativo"))
	s_grupo=Trim(Request.Form("c_grupo"))
	s_natureza=Trim(Request.Form("rb_natureza"))

	if converte_numero(s_id) <= 0 then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	
	s_id=normaliza_codigo(s_id, TAM_PLANO_CONTAS__CONTA)

	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if s_id = "" then
		alerta="NÚMERO DE IDENTIFICAÇÃO INVÁLIDO."
	elseif s_descricao = "" then
		alerta="PREENCHA A DESCRIÇÃO DA CONTA."
	elseif s_st_ativo = "" then
		alerta="INFORME O STATUS DA CONTA."
	elseif s_grupo = "" then
		alerta="SELECIONE O GRUPO DE CONTAS AO QUAL ESTA CONTA DEVE SER VINCULADA."
	elseif converte_numero(s_grupo)=0 then
		alerta="INFORME O GRUPO DE CONTAS AO QUAL ESTA CONTA DEVE SER VINCULADA."
	elseif s_natureza="" then
		alerta="INFORME A NATUREZA DA CONTA."
	elseif (Cstr(s_natureza)<>Cstr(COD_FIN_NATUREZA__CREDITO)) And (Cstr(s_natureza)<>Cstr(COD_FIN_NATUREZA__DEBITO)) then
		alerta="NATUREZA DA CONTA É INVÁLIDA."
		end if
	
	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	if operacao_selecionada = OP_INCLUI then
		if alerta = "" then
			s = "SELECT " & _
					"*" & _
				" FROM t_FIN_PLANO_CONTAS_CONTA" & _
				" WHERE" & _
					" (id = " & s_id & ")" & _
					" AND (natureza = '" & s_natureza & "')"
			r.Open s, cn
			if Not r.Eof then
				alerta = "Conta " & chr(34) & normaliza_codigo(s_id, TAM_PLANO_CONTAS__CONTA) & chr(34) & _
						 " já está cadastrada para a natureza " & _
						 chr(34) & finNaturezaDescricao(s_natureza) & chr(34)
				end if
			r.Close
			end if
		end if
	
	
'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			s = "SELECT" & _
					" TOP 1 *" & _
				" FROM t_FIN_FLUXO_CAIXA" & _
				" WHERE" & _
					" (id_plano_contas_conta = " & s_id & ")" & _
					" AND (natureza = '" & s_natureza & "')"
			r.Open s, cn
			if Not r.Eof then
				erro_fatal=True
				alerta = "REGISTRO NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE FLUXO DE CAIXA."
				end if
			r.Close 

			if Not erro_fatal then
			'	INFO P/ LOG
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_PLANO_CONTAS_CONTA" & _
					" WHERE" & _
						" (id = " & s_id & ")" & _
						" AND (natureza = '" & s_natureza & "')"
				r.Open s, cn
				if Not r.EOF then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					s_log = log_via_vetor_monta_exclusao(vLog1)
					end if
				r.Close
			
			'	APAGA!!
				s = "DELETE" & _
					" FROM t_FIN_PLANO_CONTAS_CONTA" & _
					" WHERE" &  _
						" (id = " & s_id & ")" & _
						" AND (natureza = '" & s_natureza & "')"
				cn.Execute(s)
				If Err = 0 then 
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_PLANO_CONTAS_CONTA_EXCLUSAO, s_log
				else
					erro_fatal=True
					alerta = "FALHA AO EXCLUIR O REGISTRO (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if


		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			if alerta = "" then
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_PLANO_CONTAS_GRUPO" & _
					" WHERE" & _
						" (id = " & s_grupo & ")"
				r.Open s, cn
				if r.Eof then
					alerta = "Grupo de contas " & chr(34) & normaliza_codigo(s_grupo,TAM_PLANO_CONTAS__GRUPO) & chr(34) & " não está cadastrado."
				else
				'	SOMENTE ACEITA UM GRUPO DE CONTAS INATIVADO SE FOR UM GRUPO QUE JÁ ESTAVA CADASTRADO ANTERIORMENTE
					if Cstr(r("st_ativo")) <> Cstr(COD_FIN_ST_ATIVO__ATIVO) then
						r.Close
						s = "SELECT " & _
								"*" & _
							" FROM t_FIN_PLANO_CONTAS_CONTA" & _
							" WHERE" & _
								 " (id = " & s_id & ")" & _
								 " AND (natureza = '" & s_natureza & "')"
						r.Open s, cn
						if Not r.Eof then
							if converte_numero(s_grupo) <> converte_numero(r("id_plano_contas_grupo")) then
								alerta = "Grupo de contas " & chr(34) & normaliza_codigo(s_grupo,TAM_PLANO_CONTAS__GRUPO) & " - " & Trim("" & r("descricao")) & chr(34) & " não está com status ativo."
								end if
							end if
						end if
					end if
				r.Close
				end if
			
			if alerta = "" then 
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_PLANO_CONTAS_CONTA" & _
					" WHERE" & _
						 " (id = " & s_id & ")" & _
						 " AND (natureza = '" & s_natureza & "')"
				r.Open s, cn
				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("id") = CLng(s_id)
					r("natureza")=s_natureza
					r("dt_cadastro") = Now
					r("usuario_cadastro") = usuario
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					end if
				
				r("id_plano_contas_grupo")=CLng(s_grupo)
				r("st_ativo")=CLng(s_st_ativo)
				r("descricao")=s_descricao
				r("dt_ult_atualizacao")=Now
				r("usuario_ult_atualizacao")=usuario
				
				r.Update

				If Err = 0 then 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_PLANO_CONTAS_CONTA_INCLUSAO, s_log
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if s_log <> "" then 
							s_log="Id=" & Trim("" & r("id")) & "; " & s_log
							grava_log usuario, "", "", "", OP_LOG_PLANO_CONTAS_CONTA_ALTERACAO, s_log
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
		s = "<P style='margin:5px 2px 5px 2px;'>" & alerta & "</P>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "REGISTRO ID=" & chr(34) & s_id & chr(34) & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "REGISTRO ID=" & chr(34) & s_id & chr(34) & " ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "REGISTRO ID=" & chr(34) & s_id & chr(34) & " EXCLUÍDO COM SUCESSO."
			end select			
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<BR><BR>

<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
<%
	s="FinCadPlanoContasContaMenu.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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