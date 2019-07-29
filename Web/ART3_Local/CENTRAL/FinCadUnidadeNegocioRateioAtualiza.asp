<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  FinCadUnidadeNegocioRateioAtualiza.asp
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
	
	dim intNsuNovo, intNsuHistoricoNovo
	dim s, s_aux, usuario, alerta
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim s_log, s_log_original, s_log_novo
	Dim campos_a_omitir, campos_a_omitir_exclusao
	Dim vLog1()
	s_log = ""
	s_log_original = ""
	s_log_novo = ""
	campos_a_omitir = "|dt_cadastro|dt_hr_cadastro|excluido_status|"
	campos_a_omitir_exclusao = "|excluido_status|"
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, s_id, s_id_historico, c_plano_contas_conta, c_natureza, s_st_ativo
	operacao_selecionada=Request.Form("operacao_selecionada")
	s_id=retorna_so_digitos(Trim(Request.Form("id_selecionado")))
	s_st_ativo=Trim(Request.Form("rb_st_ativo"))
	
	dim i, n, vRateio, perc_total
	redim vRateio(0)
	set vRateio(Ubound(vRateio)) = New cl_DUAS_COLUNAS
	vRateio(Ubound(vRateio)).c1 = ""
	
	perc_total = 0
	n = Request.Form("c_unidade_negocio_id").Count
	for i = 1 to n
		s = Trim(Request.Form("c_unidade_negocio_id")(i))
		if s <> "" then
			if Trim(vRateio(Ubound(vRateio)).c1) <> "" then
				redim preserve vRateio(Ubound(vRateio)+1)
				set vRateio(Ubound(vRateio)) = New cl_DUAS_COLUNAS
				end if
			with vRateio(Ubound(vRateio))
				.c1 = Trim(Request.Form("c_unidade_negocio_id")(i))
				.c2 = Trim(Request.Form("c_perc_rateio")(i))
				perc_total = perc_total + converte_numero(.c2)
				end with
			end if
		next
	
'	PLANO DE CONTAS: DECODIFICA A INFORMAÇÃO
	dim v
	c_plano_contas_conta = ""
	c_natureza = ""
	s=Trim(Request.Form("c_plano_contas_conta"))
	if s <> "" then
		v = Split(s, "|")
		c_natureza = Trim(v(Lbound(v)))
		c_plano_contas_conta = Trim(v(Ubound(v)))
		end if
	
	if operacao_selecionada <> OP_INCLUI then
		if converte_numero(s_id) <= 0 then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
		end if
	
	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if (operacao_selecionada <> OP_INCLUI) And (s_id = "") then
		alerta="NÚMERO DE IDENTIFICAÇÃO INVÁLIDO."
	elseif c_plano_contas_conta = "" then
		alerta="INFORME O PLANO DE CONTA."
	elseif converte_numero(c_plano_contas_conta) = 0 then
		alerta="PLANO DE CONTA INVÁLIDO."
	elseif (c_natureza <> COD_FIN_NATUREZA__CREDITO) And (c_natureza <> COD_FIN_NATUREZA__DEBITO) then
		alerta="NATUREZA DO PLANO DE CONTA É INVÁLIDA."
	elseif s_st_ativo = "" then
		alerta="INFORME O STATUS DO RATEIO."
	elseif perc_total <> 100 then
		alerta = "O RATEIO NÃO TOTALIZA 100% (" & formata_perc(perc_total) & "%)"
		end if
	
	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			if Not erro_fatal then
			'	INFO P/ LOG
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO" & _
					" WHERE" & _
						" (id = " & s_id & ")"
				r.Open s, cn
				if Not r.EOF then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir_exclusao
					s_log = log_via_vetor_monta_exclusao(vLog1)
					end if
				r.Close
				
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM" & _
					" WHERE" & _
						" (id_rateio = " & s_id & ")" & _
					" ORDER BY" & _
						" id_unidade_negocio"
				r.Open s, cn
				do while Not r.Eof
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir_exclusao
					s = log_via_vetor_monta_exclusao(vLog1)
					if s <> "" then
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & "(" & s & ")"
						end if
					r.MoveNext
					loop
				
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
			'	GRAVA NO HISTÓRICO
				if Not erro_fatal then
				'	GERA O NSU PARA O NOVO REGISTRO DO HISTÓRICO
					if Not fin_gera_nsu(T_FIN_UNIDADE_NEGOCIO_HISTORICO_RATEIO, intNsuHistoricoNovo, msg_erro) then
						alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO DO HISTÓRICO (" & msg_erro & ")"
					else
						if intNsuHistoricoNovo <= 0 then
							erro_fatal = True
							alerta = "NSU GERADO É INVÁLIDO (" & intNsuHistoricoNovo & ")"
						else
							s_id_historico = Cstr(intNsuHistoricoNovo)
							end if
						end if
					end if
				
				if Not erro_fatal then
					s = "INSERT INTO t_FIN_UNIDADE_NEGOCIO_HISTORICO_RATEIO " & _
							"(" & _
							"id, " & _
							"id_rateio, " & _
							"id_plano_contas_conta, " & _
							"natureza, " & _
							"st_ativo, " & _
							"st_rateio_excluido, " & _
							"descricao_plano_contas_conta, " & _
							"usuario_cadastro" & _
							")" & _
						" SELECT " & _
							s_id_historico & ", " & _
							"tFUNR.id, " & _
							"tFUNR.id_plano_contas_conta, " & _
							"tFUNR.natureza, " & _
							"tFUNR.st_ativo, " & _
							"1, " & _
							"tFPCC.descricao, " & _
							"'" & QuotedStr(usuario) & "'" & _
						" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO tFUNR" & _
							" INNER JOIN t_FIN_PLANO_CONTAS_CONTA tFPCC ON (tFUNR.id_plano_contas_conta=tFPCC.id) AND (tFUNR.natureza=tFPCC.natureza)" & _
						" WHERE" & _
							" (tFUNR.id = " & s_id & ")"
					cn.Execute(s)
					If Err <> 0 then
						erro_fatal=True
						alerta = "FALHA AO TENTAR GRAVAR O HISTÓRICO (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if Not erro_fatal then
					s = "INSERT INTO t_FIN_UNIDADE_NEGOCIO_HISTORICO_RATEIO_ITEM " & _
							"(" & _
							"id_historico_rateio, " & _
							"id_unidade_negocio, " & _
							"apelido_unidade_negocio, " & _
							"descricao_unidade_negocio, " & _
							"perc_rateio" & _
							")" & _
						" SELECT " & _
							s_id_historico & ", " & _
							"id_unidade_negocio, " & _
							"apelido, " & _
							"descricao, " & _
							"perc_rateio" & _
						" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM tFUNRI" & _
							" INNER JOIN t_FIN_UNIDADE_NEGOCIO tFUN ON (tFUNRI.id_unidade_negocio=tFUN.id)" & _
						" WHERE" & _
							" (id_rateio = " & s_id & ")"
					cn.Execute(s)
					If Err <> 0 then
						erro_fatal=True
						alerta = "FALHA AO TENTAR GRAVAR O HISTÓRICO (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if Not erro_fatal then
				'	APAGA!!
					s = "DELETE" & _
						" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM" & _
						" WHERE" &  _
							" (id_rateio = " & s_id & ")"
					cn.Execute(s)
					If Err <> 0 then
						erro_fatal=True
						alerta = "FALHA AO EXCLUIR O REGISTRO PRINCIPAL (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if Not erro_fatal then
					s = "DELETE" & _
						" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO" & _
						" WHERE" &  _
							" (id = " & s_id & ")"
					cn.Execute(s)
					If Err <> 0 then
						erro_fatal=True
						alerta = "FALHA AO EXCLUIR O REGISTRO (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if Not erro_fatal then
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_UNIDADE_NEGOCIO_RATEIO_EXCLUSAO, s_log
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
				end if


		case OP_INCLUI
		'	 =========
			if alerta = "" then
			'	GERA O NSU PARA O NOVO REGISTRO
				if Not fin_gera_nsu(T_FIN_UNIDADE_NEGOCIO_RATEIO, intNsuNovo, msg_erro) then
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
			
			if alerta = "" then
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO" & _
					" WHERE" & _
						 " (id = -1)"
				if r.State <> 0 then r.Close
				r.Open s, cn
				r.AddNew
				r("id") = CLng(s_id)
				r("usuario_cadastro") = usuario
				r("id_plano_contas_conta") = CLng(c_plano_contas_conta)
				r("natureza") = c_natureza
				r("st_ativo") = CLng(s_st_ativo)
				r("dt_ult_atualizacao") = Date
				r("dt_hr_ult_atualizacao") = Now
				r("usuario_ult_atualizacao") = usuario
				r.Update

				for i=Lbound(vRateio) to Ubound(vRateio)
					if (Trim(vRateio(i).c1) <> "") And _
						(converte_numero(vRateio(i).c2) <> 0) then
						s = "SELECT " & _
								"*" & _
							" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM" & _
							" WHERE" & _
								" (id_rateio = -1)" & _
								" AND (id_unidade_negocio = -1)"
						if r.State <> 0 then r.Close
						r.Open s, cn
						r.AddNew
						r("id_rateio") = CLng(s_id)
						r("id_unidade_negocio") = CLng(vRateio(i).c1)
						r("perc_rateio") = converte_numero(vRateio(i).c2)
						r.Update
						end if
					next
				
				if Err <> 0 then
					erro_fatal=True
					alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				
			'	OBTÉM DADOS P/ O LOG
				if alerta = "" then
					s = "SELECT " & _
							"*" & _
						" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO" & _
						" WHERE" & _
							" (id = " & s_id & ")"
					if r.State <> 0 then r.Close
					r.Open s, cn
					if Not r.Eof then
						log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & log_via_vetor_monta_inclusao(vLog1)
						end if
					
					s = "SELECT " & _
							"*" & _
						" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM" & _
						" WHERE" & _
							" (id_rateio = " & s_id & ")" & _
						" ORDER BY" & _
							" id_unidade_negocio"
					if r.State <> 0 then r.Close
					r.Open s, cn
					do while Not r.Eof
						log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & "(" & log_via_vetor_monta_inclusao(vLog1) & ")"
						r.MoveNext
						loop
					end if
				
			'	GRAVA LOG
				if alerta = "" then
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_UNIDADE_NEGOCIO_RATEIO_INCLUSAO, s_log
					end if
				
			'	GRAVA NO HISTÓRICO
				if alerta = "" then
				'	GERA O NSU PARA O NOVO REGISTRO DO HISTÓRICO
					if Not fin_gera_nsu(T_FIN_UNIDADE_NEGOCIO_HISTORICO_RATEIO, intNsuHistoricoNovo, msg_erro) then
						alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO DO HISTÓRICO (" & msg_erro & ")"
					else
						if intNsuHistoricoNovo <= 0 then
							alerta = "NSU GERADO É INVÁLIDO (" & intNsuHistoricoNovo & ")"
							end if
						end if
					end if
					
				if alerta = "" then
					s_id_historico = Cstr(intNsuHistoricoNovo)
					end if
				
				if alerta = "" then
					s = "INSERT INTO t_FIN_UNIDADE_NEGOCIO_HISTORICO_RATEIO " & _
							"(" & _
							"id, " & _
							"id_rateio, " & _
							"id_plano_contas_conta, " & _
							"natureza, " & _
							"st_ativo, " & _
							"descricao_plano_contas_conta, " & _
							"usuario_cadastro" & _
							")" & _
						" SELECT " & _
							s_id_historico & ", " & _
							"tFUNR.id, " & _
							"tFUNR.id_plano_contas_conta, " & _
							"tFUNR.natureza, " & _
							"tFUNR.st_ativo, " & _
							"tFPCC.descricao, " & _
							"'" & QuotedStr(usuario) & "'" & _
						" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO tFUNR" & _
							" INNER JOIN t_FIN_PLANO_CONTAS_CONTA tFPCC ON (tFUNR.id_plano_contas_conta=tFPCC.id) AND (tFUNR.natureza=tFPCC.natureza)" & _
						" WHERE" & _
							" (tFUNR.id = " & s_id & ")"
					cn.Execute(s)
					If Err <> 0 then
						erro_fatal=True
						alerta = "FALHA AO TENTAR GRAVAR O HISTÓRICO (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if alerta = "" then
					s = "INSERT INTO t_FIN_UNIDADE_NEGOCIO_HISTORICO_RATEIO_ITEM " & _
							"(" & _
							"id_historico_rateio, " & _
							"id_unidade_negocio, " & _
							"apelido_unidade_negocio, " & _
							"descricao_unidade_negocio, " & _
							"perc_rateio" & _
							")" & _
						" SELECT " & _
							s_id_historico & ", " & _
							"id_unidade_negocio, " & _
							"apelido, " & _
							"descricao, " & _
							"perc_rateio" & _
						" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM tFUNRI" & _
							" INNER JOIN t_FIN_UNIDADE_NEGOCIO tFUN ON (tFUNRI.id_unidade_negocio=tFUN.id)" & _
						" WHERE" & _
							" (id_rateio = " & s_id & ")"
					cn.Execute(s)
					If Err <> 0 then
						erro_fatal=True
						alerta = "FALHA AO TENTAR GRAVAR O HISTÓRICO (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if alerta = "" then
				'	~~~~~~~~~~~~~~
					cn.CommitTrans
				'	~~~~~~~~~~~~~~
					if Err <> 0 then 
						alerta = Cstr(Err) & ": " & Err.Description
						end if
				else
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					end if
				
				if r.State <> 0 then r.Close
				set r = nothing
				end if
		
		
		case OP_CONSULTA
		'	 ===========
			if alerta = "" then
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
			'	OBTÉM DADOS P/ O LOG (DADOS ORIGINAIS)
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO" & _
					" WHERE" & _
						" (id = " & s_id & ")"
				if r.State <> 0 then r.Close
				r.Open s, cn
				if Not r.Eof then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					if s_log_original <> "" then s_log_original = s_log_original & "; "
					s_log_original = s_log_original & log_via_vetor_monta_inclusao(vLog1)
					end if
				
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM" & _
					" WHERE" & _
						" (id_rateio = " & s_id & ")" & _
					" ORDER BY" & _
						" id_unidade_negocio"
				if r.State <> 0 then r.Close
				r.Open s, cn
				do while Not r.Eof
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					if s_log_original <> "" then s_log_original = s_log_original & "; "
					s_log_original = s_log_original & "(" & log_via_vetor_monta_inclusao(vLog1) & ")"
					r.MoveNext
					loop
				
			'	GRAVA DADOS
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO" & _
					" WHERE" & _
						 " (id = " & s_id & ")"
				if r.State <> 0 then r.Close
				r.Open s, cn
				if r.EOF then
					alerta = "NÃO FOI LOCALIZADO O REGISTRO DO RATEIO (ID=" & s_id & ")"
				else
					r("st_ativo") = CLng(s_st_ativo)
					r("dt_ult_atualizacao") = Date
					r("dt_hr_ult_atualizacao") = Now
					r("usuario_ult_atualizacao") = usuario
					r.Update
					end if
				
				s = "UPDATE" & _
						" t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM" & _
					" SET" & _
						" excluido_status = 1" & _
					" WHERE" & _
						" (id_rateio = " & s_id & ")"
				cn.Execute(s)
				
				for i=Lbound(vRateio) to Ubound(vRateio)
					if (Trim(vRateio(i).c1) <> "") And _
						(converte_numero(Trim(vRateio(i).c2)) <> 0) then
						s = "SELECT " & _
								"*" & _
							" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM" & _
							" WHERE" & _
								" (id_rateio = " & s_id & ")" & _
								" AND (id_unidade_negocio = " & vRateio(i).c1 & ")"
						if r.State <> 0 then r.Close
						r.Open s, cn
						if r.Eof then
							r.AddNew
							r("id_rateio") = CLng(s_id)
							r("id_unidade_negocio") = CLng(vRateio(i).c1)
							end if
						r("perc_rateio") = converte_numero(vRateio(i).c2)
						r("excluido_status") = 0
						r.Update
						end if
					next
				
			'	APAGA AS UNIDADES DE NEGÓCIO QUE NÃO PARTICIPAM DO RATEIO
				s = "DELETE" & _
					" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM" & _
					" WHERE" & _
						" (id_rateio = " & s_id & ")" & _
						" AND (excluido_status <> 0)"
				cn.Execute(s)
				
			'	OBTÉM DADOS P/ O LOG (DADOS ATUALIZADOS)
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO" & _
					" WHERE" & _
						" (id = " & s_id & ")"
				if r.State <> 0 then r.Close
				r.Open s, cn
				if Not r.Eof then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					if s_log_novo <> "" then s_log_novo = s_log_novo & "; "
					s_log_novo = s_log_novo & log_via_vetor_monta_inclusao(vLog1)
					end if
				
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM" & _
					" WHERE" & _
						" (id_rateio = " & s_id & ")" & _
					" ORDER BY" & _
						" id_unidade_negocio"
				if r.State <> 0 then r.Close
				r.Open s, cn
				do while Not r.Eof
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					if s_log_novo <> "" then s_log_novo = s_log_novo & "; "
					s_log_novo = s_log_novo & "(" & log_via_vetor_monta_inclusao(vLog1) & ")"
					r.MoveNext
					loop
				
			'	GRAVA LOG
				if alerta = "" then
					s_log = s_log_original & " => " & s_log_novo
					grava_log usuario, "", "", "", OP_LOG_UNIDADE_NEGOCIO_RATEIO_ALTERACAO, s_log
					end if
				
			'	GRAVA NO HISTÓRICO
				if alerta = "" then
				'	GERA O NSU PARA O NOVO REGISTRO DO HISTÓRICO
					if Not fin_gera_nsu(T_FIN_UNIDADE_NEGOCIO_HISTORICO_RATEIO, intNsuHistoricoNovo, msg_erro) then
						alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO DO HISTÓRICO (" & msg_erro & ")"
					else
						if intNsuHistoricoNovo <= 0 then
							alerta = "NSU GERADO É INVÁLIDO (" & intNsuHistoricoNovo & ")"
							end if
						end if
					end if
					
				if alerta = "" then
					s_id_historico = Cstr(intNsuHistoricoNovo)
					end if
				
				if alerta = "" then
					s = "INSERT INTO t_FIN_UNIDADE_NEGOCIO_HISTORICO_RATEIO " & _
							"(" & _
							"id, " & _
							"id_rateio, " & _
							"id_plano_contas_conta, " & _
							"natureza, " & _
							"st_ativo, " & _
							"descricao_plano_contas_conta, " & _
							"usuario_cadastro" & _
							")" & _
						" SELECT " & _
							s_id_historico & ", " & _
							"tFUNR.id, " & _
							"tFUNR.id_plano_contas_conta, " & _
							"tFUNR.natureza, " & _
							"tFUNR.st_ativo, " & _
							"tFPCC.descricao, " & _
							"'" & QuotedStr(usuario) & "'" & _
						" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO tFUNR" & _
							" INNER JOIN t_FIN_PLANO_CONTAS_CONTA tFPCC ON (tFUNR.id_plano_contas_conta=tFPCC.id) AND (tFUNR.natureza=tFPCC.natureza)" & _
						" WHERE" & _
							" (tFUNR.id = " & s_id & ")"
					cn.Execute(s)
					If Err <> 0 then
						erro_fatal=True
						alerta = "FALHA AO TENTAR GRAVAR O HISTÓRICO (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if alerta = "" then
					s = "INSERT INTO t_FIN_UNIDADE_NEGOCIO_HISTORICO_RATEIO_ITEM " & _
							"(" & _
							"id_historico_rateio, " & _
							"id_unidade_negocio, " & _
							"apelido_unidade_negocio, " & _
							"descricao_unidade_negocio, " & _
							"perc_rateio" & _
							")" & _
						" SELECT " & _
							s_id_historico & ", " & _
							"id_unidade_negocio, " & _
							"apelido, " & _
							"descricao, " & _
							"perc_rateio" & _
						" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM tFUNRI" & _
							" INNER JOIN t_FIN_UNIDADE_NEGOCIO tFUN ON (tFUNRI.id_unidade_negocio=tFUN.id)" & _
						" WHERE" & _
							" (id_rateio = " & s_id & ")"
					cn.Execute(s)
					If Err <> 0 then
						erro_fatal=True
						alerta = "FALHA AO TENTAR GRAVAR O HISTÓRICO (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
				if alerta = "" then
				'	~~~~~~~~~~~~~~
					cn.CommitTrans
				'	~~~~~~~~~~~~~~
					if Err <> 0 then 
						alerta = Cstr(Err) & ": " & Err.Description
						end if
				else
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
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
		select case operacao_selecionada
			case OP_INCLUI
				s = "RATEIO (PLANO DE CONTA=" & normaliza_codigo(c_plano_contas_conta,TAM_PLANO_CONTAS__CONTA) & ") CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "RATEIO (PLANO DE CONTA=" & normaliza_codigo(c_plano_contas_conta,TAM_PLANO_CONTAS__CONTA) & ") ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "RATEIO (PLANO DE CONTA=" & normaliza_codigo(c_plano_contas_conta,TAM_PLANO_CONTAS__CONTA) & ") EXCLUÍDO COM SUCESSO."
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
	s="FinCadUnidadeNegocioRateioMenu.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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