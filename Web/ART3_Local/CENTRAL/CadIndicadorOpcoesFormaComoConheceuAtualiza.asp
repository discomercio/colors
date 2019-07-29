<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  CadIndicadorOpcoesFormaComoConheceuAtualiza.asp
'     ===============================================
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
	alerta = ""
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	if Not operacao_permitida(OP_CEN_CAD_INDICADOR_OPCOES_FORMA_COMO_CONHECEU, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim s_log
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, s_id, s_descricao, s_st_inativo
	dim intOrdenacao
	operacao_selecionada=request("operacao_selecionada")
	s_id=retorna_so_digitos(trim(request("id_selecionado")))
	s_descricao=Trim(request("c_descricao"))
	s_st_inativo=trim(request("rb_st_inativo"))

	dim intNsuNovo
	if operacao_selecionada = OP_INCLUI then
	'	GERA O NSU P/ O NOVO REGISTRO
		if Not fin_gera_nsu(T_ORCAMENTISTA_E_INDICADOR__forma_como_conheceu_codigo, intNsuNovo,msg_erro) then
			alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		else
			if intNsuNovo <= 0 then
				alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovo & ")"
			else
				s_id = Cstr(intNsuNovo)
				end if
			end if
		end if
	
	if (s_id = "") Or (converte_numero(s_id) = 0) then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	
	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	if s_id = "" then
		alerta="Nº IDENTIFICAÇÃO DO REGISTRO É INVÁLIDO."
	elseif s_descricao = "" then
		alerta="INFORME UMA DESCRIÇÃO."
	elseif s_st_inativo = "" then
		alerta="INFORME SE O STATUS DEVE SER DISPONÍVEL OU INDISPONÍVEL."
		end if
	
	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
			'=========
			s="SELECT COUNT(*) AS qtde FROM t_ORCAMENTISTA_E_INDICADOR WHERE (forma_como_conheceu_codigo = '" & s_id & "')"
			r.Open s, cn
		'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
			if CLng(r("qtde")) > CLng(0) then
				erro_fatal=True
				alerta = "REGISTRO NÃO PODE SER EXCLUÍDO PORQUE ESTÁ SENDO REFERENCIADO NO CADASTRO DE INDICADORES."
				end if
			r.Close
			
			if Not erro_fatal then
			'	INFO P/ LOG
				s = "SELECT " & _
						"*" & _
					" FROM t_CODIGO_DESCRICAO" & _
					" WHERE" & _
						" (grupo = '" & GRUPO_T_CODIGO_DESCRICAO__CAD_ORCAMENTISTA_E_INDICADOR__FORMA_COMO_CONHECEU & "')" & _
						" AND (codigo = '" & s_id & "')"
				r.Open s, cn
				if Not r.EOF then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					s_log = log_via_vetor_monta_exclusao(vLog1)
					end if
				r.Close
			
			'	APAGA!!
				s = "DELETE" & _
					" FROM t_CODIGO_DESCRICAO" & _
					" WHERE" & _
						" (grupo = '" & GRUPO_T_CODIGO_DESCRICAO__CAD_ORCAMENTISTA_E_INDICADOR__FORMA_COMO_CONHECEU & "')" & _
						" AND (codigo = '" & s_id & "')"
				cn.Execute(s)
				If Err = 0 then 
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_CAD_INDICADOR_OPCAO_FORMA_COMO_CONHECEU_EXCLUSAO, s_log
				else
					erro_fatal=True
					alerta = "FALHA AO EXCLUIR O REGISTRO (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if
		
		
		case OP_INCLUI
			'=========
			if alerta = "" then
				s = "SELECT " & _
						"*" & _
					" FROM t_CODIGO_DESCRICAO" & _
					" WHERE" & _
						" (grupo = '" & GRUPO_T_CODIGO_DESCRICAO__CAD_ORCAMENTISTA_E_INDICADOR__FORMA_COMO_CONHECEU & "')" & _
						" AND (codigo = '" & s_id & "')"
				r.Open s, cn
				if Not r.Eof then
					erro_fatal = True
					alerta = "O ID '" & s_id & "' JÁ ESTÁ EM USO!!"
				else
					r.AddNew
					r("grupo") = GRUPO_T_CODIGO_DESCRICAO__CAD_ORCAMENTISTA_E_INDICADOR__FORMA_COMO_CONHECEU
					r("codigo") = s_id
					r("dt_hr_cadastro") = Now
					r("usuario_cadastro") = usuario
					r("descricao") = s_descricao
					r("st_inativo") = CLng(s_st_inativo)
					r("dt_hr_ult_atualizacao") = Now
					r("usuario_ult_atualizacao") = usuario
					r.Update
				
					If Err = 0 then 
						log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_CAD_INDICADOR_OPCAO_FORMA_COMO_CONHECEU_INCLUSAO, s_log
					else
						erro_fatal=True
						alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				r.Close
				end if
		
		
		case OP_CONSULTA
			'===========
			if alerta = "" then
				s = "SELECT " & _
						"*" & _
					" FROM t_CODIGO_DESCRICAO" & _
					" WHERE" & _
						" (grupo = '" & GRUPO_T_CODIGO_DESCRICAO__CAD_ORCAMENTISTA_E_INDICADOR__FORMA_COMO_CONHECEU & "')" & _
						" AND (codigo = '" & s_id & "')"
				r.Open s, cn
				if r.EOF then
					erro_fatal = True
					alerta = "FALHA AO LOCALIZAR O REGISTRO COM ID '" & s_id & "'!!"
				else
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					r("descricao") = s_descricao
					r("st_inativo") = CLng(s_st_inativo)
					r("dt_hr_ult_atualizacao") = Now
					r("usuario_ult_atualizacao") = usuario
					r.Update

					If Err = 0 then 
						log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if s_log <> "" then 
							s_log="Codigo=" & Trim("" & r("codigo")) & "; " & s_log
							grava_log usuario, "", "", "", OP_LOG_CAD_INDICADOR_OPCAO_FORMA_COMO_CONHECEU_ALTERACAO, s_log
							end if
					else
						erro_fatal=True
						alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				r.Close
				end if
		
		
		case else
			'====
			alerta="OPERAÇÃO INVÁLIDA."
		
		end select
	
	
'	ACERTA O CAMPO ORDENAÇÃO
'	========================
	if (alerta = "") And (Not erro_fatal) then
		if (operacao_selecionada = OP_INCLUI) OR (operacao_selecionada = OP_CONSULTA) then
			intOrdenacao = 0
			s = "SELECT " & _
					"*" & _
				" FROM t_CODIGO_DESCRICAO" & _
				" WHERE" & _
					" (grupo = '" & GRUPO_T_CODIGO_DESCRICAO__CAD_ORCAMENTISTA_E_INDICADOR__FORMA_COMO_CONHECEU & "')" & _
				" ORDER BY" & _
					" descricao" & SQL_COLLATE_CASE_ACCENT
			r.Open s, cn
			do while Not r.Eof
				intOrdenacao = intOrdenacao + 1
				r("ordenacao") = intOrdenacao
				r.Update
				r.MoveNext
				loop
			
			r.Close
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
				s = "REGISTRO ID=" & s_id & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "REGISTRO ID=" & s_id & " ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "REGISTRO ID=" & s_id & " EXCLUÍDO COM SUCESSO."
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
	s="CadIndicadorOpcoesFormaComoConheceu.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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
	if r.State <> 0 then r.Close
	set r = nothing

	cn.Close
	set cn = nothing
%>