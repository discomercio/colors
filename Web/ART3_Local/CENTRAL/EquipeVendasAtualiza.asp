<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  EquipeVendasAtualiza.asp
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
	dim cn, r, r_aux
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim criou_novo_reg
	Dim s_log, s_log_lista_original, s_log_lista_atual
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, s_apelido, s_descricao, s_supervisor, s_id_equipe_vendas, s_apelido_equipe_vendas
	operacao_selecionada=Request.Form("operacao_selecionada")
	s_apelido=Trim(Request.Form("id_selecionado"))
	s_descricao=Trim(Request.Form("c_descricao"))
	s_supervisor=Trim(Request.Form("c_supervisor"))

	if s_apelido = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	
	dim i, n, vMembros
	redim vMembros(0)
	vMembros(0) = ""
	n = Request.Form("chk_membros").Count
	for i = 1 to n
		s = Trim(Request.Form("chk_membros")(i))
		if s <> "" then
			if Trim(vMembros(ubound(vMembros))) <> "" then
				redim preserve vMembros(ubound(vMembros)+1)
				end if
			vMembros(ubound(vMembros)) = s
			end if
		next

	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if s_descricao = "" then
		alerta="PREENCHA A DESCRIÇÃO."
	elseif s_supervisor = "" then
		alerta="SELECIONE UM SUPERVISOR."
		end if
	
	if alerta <> "" then erro_consistencia=True	
		
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			s = "SELECT " & _
					"*" & _
				" FROM t_EQUIPE_VENDAS" & _
				" WHERE" & _
					" (apelido = '" & s_apelido & "')"
			r.Open s, cn
			if Not r.EOF then
				s_id_equipe_vendas = Trim("" & r("id"))
				s_apelido_equipe_vendas = Trim("" & r("apelido"))
			'	INFO P/ LOG
				log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
				s_log = log_via_vetor_monta_exclusao(vLog1)
			else
				erro_fatal=True
				alerta="FALHA AO LOCALIZAR O REGISTRO DA EQUIPE DE VENDAS: " & s_apelido
				end if
			r.Close
			
			if alerta = "" then
				s = "SELECT" & _
						" usuario" & _
					" FROM t_EQUIPE_VENDAS_X_USUARIO" & _
					" WHERE" & _
						" (id_equipe_vendas = " & s_id_equipe_vendas & ")" & _
					" ORDER BY" & _
						" usuario"
				r.Open s, cn
				s_aux = ""
				do while Not r.EOF
					if s_aux <> "" then s_aux = s_aux & ", "
					s_aux = s_aux & Trim("" & r("usuario"))
					r.MoveNext
					loop
				
				if s_aux = "" then s_aux = "(vazio)"
			'	INFO P/ LOG
				if s_log <> "" then s_log = s_log & "; listagem da equipe = " & s_aux
				end if
			
			if alerta = "" then
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
			'	APAGA REGISTROS DA TABELA-FILHA!!
				s = "DELETE" & _
					" FROM t_EQUIPE_VENDAS_X_USUARIO" & _
					" WHERE" &  _
						" (id_equipe_vendas = " & s_id_equipe_vendas & ")"
				cn.Execute(s)
				If Err <> 0 then 
					erro_fatal=True
					alerta = "FALHA AO EXCLUIR O(S) REGISTRO(S) NA TABELA-FILHA (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if
			
			if alerta = "" then
			'	APAGA O REGISTRO PRINCIPAL
				s = "DELETE" & _
					" FROM t_EQUIPE_VENDAS" & _
					" WHERE" & _
						" (id = " & s_id_equipe_vendas & ")"
				cn.Execute(s)
				If Err <> 0 then 
					erro_fatal=True
					alerta = "FALHA AO EXCLUIR O REGISTRO PRINCIPAL (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if
			
			if alerta = "" then
			'	~~~~~~~~~~~~~~
				cn.CommitTrans
			'	~~~~~~~~~~~~~~
				if Err <> 0 then 
					alerta=Cstr(Err) & ": " & Err.Description
					erro_fatal = True
					end if
			else
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Err.Clear
				end if
			
			if alerta = "" then
				if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_EQUIPE_VENDAS_EXCLUSAO, s_log
				end if


		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			s_log_lista_original = ""
			s_log_lista_atual = ""
			
			if operacao_selecionada = OP_INCLUI then
				if alerta = "" then
					if Not fin_gera_nsu(T_EQUIPE_VENDAS, intNsuNovo, msg_erro) then 
						alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
					else
						if intNsuNovo <= 0 then
							alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovo & ")"
							end if
						end if
					end if
				end if
			
			if alerta = "" then 
				s = "SELECT " & _
						"*" & _
					" FROM t_EQUIPE_VENDAS" & _
					" WHERE" & _
						 " (apelido = '" & s_apelido & "')"
				r.Open s, cn
				
				if r.EOF And (operacao_selecionada = OP_CONSULTA) then
					alerta = "FALHA AO LOCALIZAR O REGISTRO DA EQUIPE DE VENDAS: " & s_apelido
					end if
				
				if alerta = "" then
				'	~~~~~~~~~~~~~
					cn.BeginTrans
				'	~~~~~~~~~~~~~
					if r.EOF then
						r.AddNew 
						criou_novo_reg = True
						r("id") = intNsuNovo
						r("apelido") = s_apelido
						r("dt_cadastro") = Now
						r("usuario_cadastro") = usuario
					else
						criou_novo_reg = False
						log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
						end if
					
					r("descricao")=s_descricao
					r("supervisor")=s_supervisor
					r("dt_ult_atualizacao")=Now
					r("usuario_ult_atualizacao")=usuario
					
					s_id_equipe_vendas = Trim("" & r("id"))
					s_apelido_equipe_vendas = Trim("" & r("apelido"))
					
					r.Update

					If Err = 0 then 
						log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
						if criou_novo_reg then
							s_log = log_via_vetor_monta_inclusao(vLog2)
						else
							s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
							end if
					else
						erro_fatal=True
						alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					
				'	INFO P/ O LOG
					if alerta = "" then
						s = "SELECT " & SCHEMA_BD & ".ConcatenaMembrosEquipeVendas(" & s_id_equipe_vendas & ", ', ') AS membros"
						set r_aux = cn.Execute(s)
						if Not r_aux.EOF then s_log_lista_original = Trim("" & r_aux("membros"))
						if s_log_lista_original = "" then s_log_lista_original = "(lista vazia)"
						end if
					
					if alerta = "" then
						s = "UPDATE t_EQUIPE_VENDAS_X_USUARIO SET" & _
								" excluido_status = 1" & _
							" WHERE" & _
								" (id_equipe_vendas = " & s_id_equipe_vendas & ")"
						cn.Execute(s)
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO ALTERAR REGISTROS DURANTE ATUALIZAÇÃO DA LISTA DE MEMBROS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if
					
					if alerta = "" then
						for i=Lbound(vMembros) to ubound(vMembros)
							if alerta = "" then
								if Trim(vMembros(i)) <> "" then
									s = "SELECT " & _
											"*" & _
										" FROM t_EQUIPE_VENDAS_X_USUARIO" & _
										" WHERE" & _
											" (id_equipe_vendas = " & s_id_equipe_vendas & ")" & _
											" AND (usuario = '" & Trim(vMembros(i)) & "')"
									if r.State <> 0 then r.Close
									r.Open s, cn
									if Not r.EOF then
										r("excluido_status") = 0
									else
										r.AddNew
										r("id_equipe_vendas") = CLng(s_id_equipe_vendas)
										r("usuario") = Trim(vMembros(i))
										r("dt_cadastro") = Now
										r("usuario_cadastro") = usuario
										r("excluido_status") = 0
										end if
									
									r.Update
									if Err <> 0 then
										erro_fatal=True
										alerta = "FALHA AO INSERIR REGISTROS DURANTE ATUALIZAÇÃO DA LISTA DE MEMBROS (" & Cstr(Err) & ": " & Err.Description & ")."
										end if
									end if
								end if
							next
						end if

					if alerta = "" then
						s = "DELETE FROM t_EQUIPE_VENDAS_X_USUARIO" & _
							" WHERE" & _
								" (id_equipe_vendas = " & s_id_equipe_vendas & ")" & _
								" AND (excluido_status = 1)"
						cn.Execute(s)
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO EXCLUIR REGISTROS DURANTE ATUALIZAÇÃO DA LISTA DE MEMBROS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if
						
				'	INFO P/ O LOG
					if alerta = "" then
						s = "SELECT " & SCHEMA_BD & ".ConcatenaMembrosEquipeVendas(" & s_id_equipe_vendas & ", ', ') AS membros"
						set r_aux = cn.Execute(s)
						if Not r_aux.EOF then s_log_lista_atual = Trim("" & r_aux("membros"))
						if s_log_lista_atual = "" then s_log_lista_atual = "(lista vazia)"
						end if
					
					if alerta = "" then
					'	~~~~~~~~~~~~~~
						cn.CommitTrans
					'	~~~~~~~~~~~~~~
						if Err <> 0 then 
							alerta=Cstr(Err) & ": " & Err.Description
							erro_fatal = True
							end if
					else
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Err.Clear
						end if

					if alerta = "" then
						if criou_novo_reg then
							if s_log <> "" then 
								s_log = s_log & "; membros: " & s_log_lista_atual
								grava_log usuario, "", "", "", OP_LOG_EQUIPE_VENDAS_INCLUSAO, s_log
								end if
						else
							if (s_log <> "") Or (s_log_lista_original <> s_log_lista_atual) then 
								if s_log <> "" then s_log="Id=" & s_id_equipe_vendas & " (" & s_apelido_equipe_vendas & "); " & s_log
								if s_log_lista_original <> s_log_lista_atual then s_log = s_log & "; membros originais = " & s_log_lista_original & "; membros atuais = " & s_log_lista_atual
								grava_log usuario, "", "", "", OP_LOG_EQUIPE_VENDAS_ALTERACAO, s_log
								end if
							end if
						end if
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
				s = "REGISTRO DA EQUIPE " & chr(34) & s_apelido & chr(34) & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "REGISTRO DA EQUIPE " & chr(34) & s_apelido & chr(34) & " ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "REGISTRO DA EQUIPE " & chr(34) & s_apelido & chr(34) & " EXCLUÍDO COM SUCESSO."
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
	s="MenuEquipeVendas.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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