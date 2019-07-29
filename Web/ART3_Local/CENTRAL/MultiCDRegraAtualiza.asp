<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  MultiCDRegraAtualiza.asp
'     =====================================
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
	
	dim s, s_aux, s_campo, s_value, usuario, alerta
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	if Not operacao_permitida(OP_CEN_MULTI_CD_CADASTRO_REGRAS_CONSUMO_ESTOQUE, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r
	dim tRegra, tRegraUf, tRegraUfPessoa, tRegraUfPessoaCd
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim MARGEM_N1
	dim MARGEM_N2
	dim MARGEM_N3
	MARGEM_N1 = String(1, vbTab)
	MARGEM_N2 = String(2, vbTab)
	MARGEM_N3 = String(3, vbTab)

	Dim s_log, s_log_aux, s_log_produtos_associados
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	s_log_produtos_associados = ""
	campos_a_omitir = "|dt_ult_atualizacao|dt_hr_ult_atualizacao|usuario_ult_atualizacao|timestamp|"
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim id_selecionado, s_apelido_regra, s_descricao, operacao_selecionada, ckb_regra_st_inativo
	operacao_selecionada = request("operacao_selecionada")
	id_selecionado = Trim(Request("id_selecionado"))
	s_apelido_regra = Trim(request("c_apelido"))
	s_descricao = Trim(request("c_descricao"))
	ckb_regra_st_inativo = Trim(Request("ckb_regra_st_inativo"))

	if s_apelido_regra = "" then Response.Redirect("aviso.asp?id=" & ERR_MULTI_CD_REGRA_APELIDO_NAO_INFORMADO)
	
	dim vCD, iCD, c_lista_cd, idxCD
	c_lista_cd = Trim(Request("c_lista_cd"))
	if c_lista_cd = "" then
		redim vCD(0)
		vCD(UBound(vCD)) = 0
	else
		vCD = Split(c_lista_cd, "|")
		for iCD=LBound(vCD) to UBound(vCD)
			vCD(iCD) = converte_numero(vCD(iCD))
			next
		end if

	dim vUF, iUF
	vUF = UF_get_array

	dim vPessoa, iPessoa
	vPessoa = TipoPessoa_get_array

	' INICIALIZAÇÃO
	dim oRegra
	inicializa_cl_CADASTRO_MULTI_CD_REGRA oRegra
	
	' PREENCHIMENTO
	if operacao_selecionada = OP_INCLUI then
		oRegra.id = 0
	else
		oRegra.id = converte_numero(id_selecionado)
		end if

	oRegra.apelido = s_apelido_regra
	oRegra.descricao = s_descricao
	if ckb_regra_st_inativo <> "" then
		oRegra.st_inativo = 1
	else
		oRegra.st_inativo = 0
		end if
	
	' LÊ AS CONFIGURAÇÕES DE CADA UF
	for iUF=LBound(vUF) to UBound(vUF)
		' UF
		oRegra.vUF(iUF).uf = vUF(iUF)
		' CHECK BOX "UF DESATIVADA"
		s_campo = "ckb_UF_" & vUF(iUF)
		s_value = Trim(Request(s_campo))
		if s_value <> "" then
			oRegra.vUF(iUF).st_inativo = 1
		else
			oRegra.vUF(iUF).st_inativo = 0
			end if

		' PARA CADA UF, LÊ AS CONFIGURAÇÕES P/ CADA TIPO DE PESSOA
		for iPessoa=LBound(vPessoa) to UBound(vPessoa)
			' TIPO DE PESSOA
			oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa = vPessoa(iPessoa)
			' CHECK BOX "PESSOA DESATIVADA"
			s_campo = "ckb_pessoa_" & vUF(iUF) & "_" & vPessoa(iPessoa)
			s_value = Trim(Request(s_campo))
			if s_value <> "" then
				oRegra.vUF(iUF).vPessoa(iPessoa).st_inativo = 1
			else
				oRegra.vUF(iUF).vPessoa(iPessoa).st_inativo = 0
				end if
			' CD SPE
			s_campo = "c_cd_spe_" & vUF(iUF) & "_" & vPessoa(iPessoa)
			s_value = Trim(Request(s_campo))
			oRegra.vUF(iUF).vPessoa(iPessoa).spe_id_nfe_emitente = converte_numero(s_value)

			' PARA CADA TIPO DE PESSOA, LÊ OS CD'S CONFIGURADOS PARA SEREM USADOS NO CASO DE PRODUTOS DISPONÍVEIS
			idxCD = LBound(oRegra.vUF(LBound(oRegra.vUF)).vPessoa(LBound(oRegra.vUF(LBound(oRegra.vUF)).vPessoa)).vCD) - 1
			for iCD=LBound(vCD) to UBound(vCD)
				' ID_NFE_EMITENTE
				s_campo = "c_id_nfe_emitente_" & vUF(iUF) & "_" & vPessoa(iPessoa) & "_" & vCD(iCD)
				s_value = Trim(Request(s_campo))
				if converte_numero(s_value) <> 0 then
					idxCD = idxCD + 1
					oRegra.vUF(iUF).vPessoa(iPessoa).vCD(idxCD).id_nfe_emitente = converte_numero(s_value)
					' CHECK BOX "DESATIVADO"
					s_campo = "ckb_cd_" & vUF(iUF) & "_" & vPessoa(iPessoa) & "_" & vCD(iCD)
					s_value = Trim(Request(s_campo))
					if s_value <> "" then
						oRegra.vUF(iUF).vPessoa(iPessoa).vCD(idxCD).st_inativo = 1
					else
						oRegra.vUF(iUF).vPessoa(iPessoa).vCD(idxCD).st_inativo = 0
						end if
					' ORDEM DE PRIORIDADE
					s_campo = "c_ordem_" & vUF(iUF) & "_" & vPessoa(iPessoa) & "_" & vCD(iCD)
					s_value = Trim(Request(s_campo))
					oRegra.vUF(iUF).vPessoa(iPessoa).vCD(idxCD).ordem_prioridade = converte_numero(s_value)
					end if
				next
			next
		next
	
	dim erro_consistencia, erro_fatal
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if Trim("" & oRegra.apelido) = "" then
		alerta="APELIDO DA REGRA NÃO FOI PREENCHIDO."
		end if
	
	if alerta = "" then
		if operacao_selecionada <> OP_EXCLUI then
			for iUF=LBound(oRegra.vUF) to UBound(oRegra.vUF)
				if Trim("" & oRegra.vUF(iUF).uf) <> "" then
					if (oRegra.vUF(iUF).st_inativo = 0) then
						for iPessoa=LBound(oRegra.vUF(iUF).vPessoa) to UBound(oRegra.vUF(iUF).vPessoa)
							if (oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa <> "") And (oRegra.vUF(iUF).vPessoa(iPessoa).st_inativo = 0) then
								if converte_numero(oRegra.vUF(iUF).vPessoa(iPessoa).spe_id_nfe_emitente) = 0 then
									alerta=texto_add_br(alerta)
									alerta=alerta & "É necessário informar o CD a ser usado para os produtos 'Sem Presença no Estoque' para a UF '" & oRegra.vUF(iUF).uf & "' no caso de '" & descricao_multi_CD_regra_tipo_pessoa(oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa) & "'"
								else
									for iCD=LBound(oRegra.vUF(iUF).vPessoa(iPessoa).vCD) to UBound(oRegra.vUF(iUF).vPessoa(iPessoa).vCD)
										if oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).id_nfe_emitente = oRegra.vUF(iUF).vPessoa(iPessoa).spe_id_nfe_emitente then
											if oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).st_inativo = 1 then
												alerta=texto_add_br(alerta)
												alerta=alerta & "O CD selecionado para os produtos 'Sem Presença no Estoque' para a UF '" & oRegra.vUF(iUF).uf & "' no caso de '" & descricao_multi_CD_regra_tipo_pessoa(oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa) & "' não pode estar desativado na relação de CD's a ser usada para os produtos disponíveis"
												exit for
												end if
											end if
										next
									end if
								end if
							next
						end if
					end if
				next
			end if
		end if

	if alerta <> "" then erro_consistencia=True

	Err.Clear

	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(tRegra, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(tRegraUf, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(tRegraUfPessoa, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(tRegraUfPessoaCd, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim id_regra_cd, id_regra_cd_uf, id_regra_cd_uf_pessoa, id_regra_cd_uf_pessoa_cd
	
'	GERA O ID P/ A NOVA REGRA?
	if alerta = "" then
		if operacao_selecionada = OP_INCLUI then
			s = "SELECT * FROM t_WMS_REGRA_CD WHERE (apelido = '" & QuotedStr(s_apelido_regra) & "')"
			if tRegra.State <> 0 then tRegra.Close
			tRegra.Open s, cn
			if Not tRegra.Eof then
				alerta = "JÁ EXISTE UMA REGRA CADASTRADA COM O APELIDO '" & Trim("" & tRegra("apelido")) & "'"
				end if

			if alerta = "" then
				if Not fin_gera_nsu(NSU_WMS_REGRA_CD, id_regra_cd, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
				end if
		else
			s = "SELECT * FROM t_WMS_REGRA_CD WHERE (id = " & id_selecionado & ")"
			if tRegra.State <> 0 then tRegra.Close
			tRegra.Open s, cn
			if tRegra.Eof then
				alerta = "FALHA AO TENTAR LOCALIZAR O REGISTRO DA REGRA NO BANCO DE DADOS (ID=" & id_selecionado & ")"
				end if
			end if
		end if
		

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			if alerta = "" then
				s="SELECT * FROM t_PRODUTO_X_WMS_REGRA_CD WHERE (id_wms_regra_cd = " & id_selecionado & ") ORDER BY fabricante, produto"
				if r.State <> 0 then r.Close
				r.Open s, cn
				do while Not r.Eof
					if s_log_produtos_associados <> "" then s_log_produtos_associados = s_log_produtos_associados & ", "
					s_log_produtos_associados = s_log_produtos_associados & "(" & Trim("" & r("fabricante")) & ")" & Trim("" & r("produto"))
					r.MoveNext
					loop
				if r.State <> 0 then r.Close
				
				if Not erro_fatal then
				'	INFO P/ LOG
					s="SELECT * FROM t_WMS_REGRA_CD WHERE (id = " & id_selecionado & ")"
					if tRegra.State <> 0 then tRegra.Close
					tRegra.Open s, cn
					if Not tRegra.EOF then
						log_via_vetor_carrega_do_recordset tRegra, vLog1, campos_a_omitir
						s_log_aux = log_via_vetor_monta_exclusao(vLog1)
						if s_log <> "" then s_log = s_log & chr(13)
						s_log = s_log & s_log_aux

						s="SELECT * FROM t_WMS_REGRA_CD_X_UF WHERE (id_wms_regra_cd = " & tRegra("id") & ") ORDER BY uf"
						if tRegraUf.State <> 0 then tRegraUf.Close
						tRegraUf.Open s, cn
						do while Not tRegraUf.Eof
							log_via_vetor_carrega_do_recordset tRegraUf, vLog1, campos_a_omitir
							s_log_aux = log_via_vetor_monta_exclusao(vLog1)
							if s_log <> "" then s_log = s_log & chr(13) & chr(13)
							s_log = s_log & MARGEM_N1 & s_log_aux

							s="SELECT * FROM t_WMS_REGRA_CD_X_UF_X_PESSOA WHERE (id_wms_regra_cd_x_uf = " & tRegraUf("id") & ")"
							if tRegraUfPessoa.State <> 0 then tRegraUfPessoa.Close
							tRegraUfPessoa.Open s, cn
							do while Not tRegraUfPessoa.Eof
								log_via_vetor_carrega_do_recordset tRegraUfPessoa, vLog1, campos_a_omitir
								s_log_aux = log_via_vetor_monta_exclusao(vLog1)
								if s_log <> "" then s_log = s_log & chr(13)
								s_log = s_log & MARGEM_N2 & s_log_aux
								
								s="SELECT * FROM t_WMS_REGRA_CD_X_UF_X_PESSOA_X_CD WHERE (id_wms_regra_cd_x_uf_x_pessoa = " & tRegraUfPessoa("id") & ")"
								if tRegraUfPessoaCd.State <> 0 then tRegraUfPessoaCd.Close
								tRegraUfPessoaCd.Open s, cn
								do while Not tRegraUfPessoaCd.Eof
									log_via_vetor_carrega_do_recordset tRegraUfPessoaCd, vLog1, campos_a_omitir
									s_log_aux = log_via_vetor_monta_exclusao(vLog1)
									if s_log <> "" then s_log = s_log & chr(13)
									s_log = s_log & MARGEM_N3 & s_log_aux

									tRegraUfPessoaCd.MoveNext
									loop

								tRegraUfPessoa.MoveNext
								loop

							tRegraUf.MoveNext
							loop
						
						end if
					
				'	APAGA!!
				'	LEMBRANDO QUE AS FOREIGN KEYS ESTÃO CRIADAS COM 'ON DELETE CASCADE'
				'	~~~~~~~~~~~~~
					cn.BeginTrans
				'	~~~~~~~~~~~~~
					if Not erro_fatal then
						s="DELETE FROM t_WMS_REGRA_CD WHERE (id = " & id_selecionado & ")"
						cn.Execute(s)

						If Err = 0 then
							if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_MULTI_CD_REGRA_EXCLUSAO, s_log
						else
							erro_fatal=True
							alerta = "FALHA AO REMOVER A REGRA '" & s_apelido_regra & "' (ID=" & id_selecionado & "): " & Cstr(Err) & ": " & Err.Description
							end if
						
						If Not erro_fatal then
							s = "DELETE FROM t_PRODUTO_X_WMS_REGRA_CD WHERE (id_wms_regra_cd = " & id_selecionado & ")"
							cn.Execute(s)
							
							If Err = 0 then
								s_aux = "Cancelamento do vínculo com os produtos associados devido à exclusão da regra '" & s_apelido_regra & "' (Id=" & id_selecionado & ")"
								if s_log_produtos_associados = "" then
									s_log_produtos_associados = s_aux & ": nenhum produto associado à regra"
								else
									s_log_produtos_associados = s_aux & ": " & s_log_produtos_associados
									end if
								grava_log usuario, "", "", "", OP_LOG_MULTI_CD_REGRA_EXCLUSAO_PRODUTOS_ASSOCIADOS, s_log_produtos_associados
							else
								erro_fatal=True
								alerta = "FALHA AO REMOVER OS VÍNCULOS COM OS PRODUTOS ASSOCIADOS À REGRA '" & s_apelido_regra & "' (ID=" & id_selecionado & "): " & Cstr(Err) & ": " & Err.Description
								end if
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
					end if
				end if


		case OP_INCLUI
		'	 =========
			if alerta = "" then
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				s = "SELECT * FROM t_WMS_REGRA_CD WHERE (id = -1)"
				if tRegra.State <> 0 then tRegra.Close
				tRegra.Open s, cn
				tRegra.AddNew
				tRegra("id") = id_regra_cd
				tRegra("st_inativo") = oRegra.st_inativo
				tRegra("apelido") = oRegra.apelido
				tRegra("descricao") = oRegra.descricao
				tRegra("usuario_cadastro") = usuario
				tRegra("dt_cadastro") = Date
				tRegra("dt_hr_cadastro") = Now
				tRegra("usuario_ult_atualizacao") = usuario
				tRegra("dt_ult_atualizacao") = Date
				tRegra("dt_hr_ult_atualizacao") = Now
				
				log_via_vetor_carrega_do_recordset tRegra, vLog1, campos_a_omitir
				s_log_aux = log_via_vetor_monta_inclusao(vLog1)
				if s_log <> "" then s_log = s_log & chr(13)
				s_log = s_log & s_log_aux
				
				tRegra.Update

				if Err <> 0 then
					alerta=Cstr(Err) & ": " & Err.Description
					erro_fatal = True
					end if

				if alerta = "" then
					for iUF=LBound(oRegra.vUF) to UBound(oRegra.vUF)
						if Trim("" & oRegra.vUF(iUF).uf) <> "" then
							if Not fin_gera_nsu(NSU_WMS_REGRA_CD_X_UF, id_regra_cd_uf, msg_erro) then
								alerta = "Falha ao tentar gerar o NSU para a tabela " & NSU_WMS_REGRA_CD_X_UF
								erro_fatal = True
								end if

							if alerta = "" then
								s = "SELECT * FROM t_WMS_REGRA_CD_X_UF WHERE (uf = 'XX')"
								if tRegraUf.State <> 0 then tRegraUf.Close
								tRegraUf.Open s, cn
								tRegraUf.AddNew
								tRegraUf("id") = id_regra_cd_uf
								tRegraUf("id_wms_regra_cd") = id_regra_cd
								tRegraUf("uf") = oRegra.vUF(iUF).uf
								tRegraUf("st_inativo") = oRegra.vUF(iUF).st_inativo
								
								log_via_vetor_carrega_do_recordset tRegraUf, vLog1, campos_a_omitir
								s_log_aux = log_via_vetor_monta_inclusao(vLog1)
								if s_log <> "" then s_log = s_log & chr(13) & chr(13)
								s_log = s_log & MARGEM_N1 & s_log_aux
								
								tRegraUf.Update

								if Err <> 0 then
									alerta=Cstr(Err) & ": " & Err.Description
									erro_fatal = True
									exit for
									end if
								end if

							if alerta = "" then
								for iPessoa=LBound(oRegra.vUF(iUF).vPessoa) to UBound(oRegra.vUF(iUF).vPessoa)
									if oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa <> "" then
										if Not fin_gera_nsu(NSU_WMS_REGRA_CD_X_UF_X_PESSOA, id_regra_cd_uf_pessoa, msg_erro) then
											alerta = "Falha ao tentar gerar o NSU para a tabela " & NSU_WMS_REGRA_CD_X_UF_X_PESSOA
											erro_fatal = True
											end if

										if alerta = "" then
											s = "SELECT * FROM t_WMS_REGRA_CD_X_UF_X_PESSOA WHERE (id = -1)"
											if tRegraUfPessoa.State <> 0 then tRegraUfPessoa.Close
											tRegraUfPessoa.Open s, cn
											tRegraUfPessoa.AddNew
											tRegraUfPessoa("id") = id_regra_cd_uf_pessoa
											tRegraUfPessoa("id_wms_regra_cd_x_uf") = id_regra_cd_uf
											tRegraUfPessoa("tipo_pessoa") = oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa
											tRegraUfPessoa("st_inativo") = oRegra.vUF(iUF).vPessoa(iPessoa).st_inativo
											tRegraUfPessoa("spe_id_nfe_emitente") = converte_numero(oRegra.vUF(iUF).vPessoa(iPessoa).spe_id_nfe_emitente)

											log_via_vetor_carrega_do_recordset tRegraUfPessoa, vLog1, campos_a_omitir
											s_log_aux = log_via_vetor_monta_inclusao(vLog1)
											if s_log <> "" then s_log = s_log & chr(13)
											s_log = s_log & MARGEM_N2 & s_log_aux

											tRegraUfPessoa.Update

											if Err <> 0 then
												alerta=Cstr(Err) & ": " & Err.Description
												erro_fatal = True
												exit for
												end if
											end if

										if alerta = "" then
											for iCD=LBound(oRegra.vUF(iUF).vPessoa(iPessoa).vCD) to UBound(oRegra.vUF(iUF).vPessoa(iPessoa).vCD)
												if oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).id_nfe_emitente <> 0 then
													if Not fin_gera_nsu(NSU_WMS_REGRA_CD_X_UF_X_PESSOA_X_CD, id_regra_cd_uf_pessoa_cd, msg_erro) then
														alerta = "Falha ao tentar gerar o NSU para a tabela " & NSU_WMS_REGRA_CD_X_UF_X_PESSOA_X_CD
														erro_fatal = True
														end if
													
													if alerta = "" then
														s = "SELECT * FROM t_WMS_REGRA_CD_X_UF_X_PESSOA_X_CD WHERE (id = -1)"
														if tRegraUfPessoaCd.State <> 0 then tRegraUfPessoaCd.Close
														tRegraUfPessoaCd.Open s, cn
														tRegraUfPessoaCd.AddNew

														tRegraUfPessoaCd("id") = id_regra_cd_uf_pessoa_cd
														tRegraUfPessoaCd("id_wms_regra_cd_x_uf_x_pessoa") = id_regra_cd_uf_pessoa
														tRegraUfPessoaCd("id_nfe_emitente") = oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).id_nfe_emitente
														tRegraUfPessoaCd("ordem_prioridade") = oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).ordem_prioridade
														tRegraUfPessoaCd("st_inativo") = oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).st_inativo

														log_via_vetor_carrega_do_recordset tRegraUfPessoaCd, vLog1, campos_a_omitir
														s_log_aux = log_via_vetor_monta_inclusao(vLog1)
														if s_log <> "" then s_log = s_log & chr(13)
														s_log = s_log & MARGEM_N3 & s_log_aux

														tRegraUfPessoaCd.Update

														if Err <> 0 then
															alerta=Cstr(Err) & ": " & Err.Description
															erro_fatal = True
															exit for
															end if
														end if
													end if
												
												if alerta <> "" then exit for
												next 'for iCD
											end if
										end if

									if alerta <> "" then exit for
									next 'for iPessoa
								end if
							end if

						if alerta <> "" then exit for
						next 'for iUF
					end if

				if alerta = "" then
					grava_log usuario, "", "", "", OP_LOG_MULTI_CD_REGRA_INCLUSAO, s_log
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
				end if

		case OP_CONSULTA
		'	 ===========
			if alerta = "" then
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				s_log = "Edição da regra de consumo do estoque (multi-CD): Apelido = '" & oRegra.apelido & "' (Id=" & id_selecionado & ")"

				s = "SELECT * FROM t_WMS_REGRA_CD WHERE (id = " & id_selecionado & ")"
				if tRegra.State <> 0 then tRegra.Close
				tRegra.Open s, cn
				if tRegra.EOF then
					alerta = "Falha ao tentar localizar o registro da regra (Id=" & id_selecionado & ")"
					end if

				if alerta = "" then
					log_via_vetor_carrega_do_recordset tRegra, vLog1, campos_a_omitir
				
					tRegra("st_inativo") = converte_numero(oRegra.st_inativo)
					tRegra("apelido") = oRegra.apelido
					tRegra("descricao") = oRegra.descricao
					tRegra("dt_ult_atualizacao") = Date
					tRegra("dt_hr_ult_atualizacao") = Now
					tRegra("usuario_ult_atualizacao") = usuario
					tRegra.Update

					log_via_vetor_carrega_do_recordset tRegra, vLog2, campos_a_omitir

					s_log_aux = log_via_vetor_monta_alteracao(vLog1, vLog2)
					if s_log_aux <> "" then
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & s_log_aux
						end if
					end if

				if alerta = "" then
					for iUF=LBound(oRegra.vUF) to UBound(oRegra.vUF)
						if Trim("" & oRegra.vUF(iUF).uf) <> "" then
							s = "SELECT * FROM t_WMS_REGRA_CD_X_UF WHERE (id_wms_regra_cd = " & id_selecionado & ") AND (uf = '" & Trim("" & oRegra.vUF(iUF).uf) & "')"
							if tRegraUf.State <> 0 then tRegraUf.Close
							tRegraUf.Open s, cn
							if Not tRegraUf.Eof then
								id_regra_cd_uf = tRegraUf("id")
								log_via_vetor_carrega_do_recordset tRegraUf, vLog1, campos_a_omitir
								tRegraUf("st_inativo") = converte_numero(oRegra.vUF(iUF).st_inativo)
								tRegraUf.Update
								log_via_vetor_carrega_do_recordset tRegraUf, vLog2, campos_a_omitir
								s_log_aux = log_via_vetor_monta_alteracao(vLog1, vLog2)
								if s_log_aux <> "" then
									if s_log <> "" then s_log = s_log & chr(13) & chr(13)
									s_log = s_log & MARGEM_N1 & "Alteração UF=" & oRegra.vUF(iUF).uf & ": " & s_log_aux
									end if
							else
								if Not fin_gera_nsu(NSU_WMS_REGRA_CD_X_UF, id_regra_cd_uf, msg_erro) then
									alerta = "Falha ao tentar gerar o NSU para a tabela " & NSU_WMS_REGRA_CD_X_UF
									erro_fatal = True
									end if

								if alerta = "" then
									tRegraUf.AddNew
									tRegraUf("id") = id_regra_cd_uf
									tRegraUf("id_wms_regra_cd") = id_selecionado
									tRegraUf("uf") = oRegra.vUF(iUF).uf
									tRegraUf("st_inativo") = oRegra.vUF(iUF).st_inativo
									tRegraUf.Update

									log_via_vetor_carrega_do_recordset tRegraUf, vLog1, campos_a_omitir
									s_log_aux = log_via_vetor_monta_inclusao(vLog1)
									if s_log <> "" then s_log = s_log & chr(13) & chr(13)
									s_log = s_log & MARGEM_N1 & "Inclusão UF=" & oRegra.vUF(iUF).uf & ": " & s_log_aux
									end if
								end if
							
							if alerta = "" then
								for iPessoa=LBound(oRegra.vUF(iUF).vPessoa) to UBound(oRegra.vUF(iUF).vPessoa)
									if oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa <> "" then
										s = "SELECT * FROM t_WMS_REGRA_CD_X_UF_X_PESSOA WHERE (id_wms_regra_cd_x_uf = " & id_regra_cd_uf & ") AND (tipo_pessoa = '" & oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa & "')"
										if tRegraUfPessoa.State <> 0 then tRegraUfPessoa.Close
										tRegraUfPessoa.Open s, cn
										if Not tRegraUfPessoa.Eof then
											id_regra_cd_uf_pessoa = tRegraUfPessoa("id")
											log_via_vetor_carrega_do_recordset tRegraUfPessoa, vLog1, campos_a_omitir
											tRegraUfPessoa("st_inativo") = converte_numero(oRegra.vUF(iUF).vPessoa(iPessoa).st_inativo)
											tRegraUfPessoa("spe_id_nfe_emitente") = converte_numero(oRegra.vUF(iUF).vPessoa(iPessoa).spe_id_nfe_emitente)
											tRegraUfPessoa.Update
											log_via_vetor_carrega_do_recordset tRegraUfPessoa, vLog2, campos_a_omitir
											s_log_aux = log_via_vetor_monta_alteracao(vLog1, vLog2)
											if s_log_aux <> "" then
												if s_log <> "" then s_log = s_log & chr(13)
												s_log = s_log & MARGEM_N2 & "Alteração: UF=" & oRegra.vUF(iUF).uf & ", tipo_pessoa=" & oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa & ": " & s_log_aux
												end if
										else
											if Not fin_gera_nsu(NSU_WMS_REGRA_CD_X_UF_X_PESSOA, id_regra_cd_uf_pessoa, msg_erro) then
												alerta = "Falha ao tentar gerar o NSU para a tabela " & NSU_WMS_REGRA_CD_X_UF_X_PESSOA
												erro_fatal = True
												end if
											
											if alerta = "" then
												tRegraUfPessoa.AddNew
												tRegraUfPessoa("id") = id_regra_cd_uf_pessoa
												tRegraUfPessoa("id_wms_regra_cd_x_uf") = id_regra_cd_uf
												tRegraUfPessoa("tipo_pessoa") = oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa
												tRegraUfPessoa("st_inativo") = converte_numero(oRegra.vUF(iUF).vPessoa(iPessoa).st_inativo)
												tRegraUfPessoa("spe_id_nfe_emitente") = converte_numero(oRegra.vUF(iUF).vPessoa(iPessoa).spe_id_nfe_emitente)
												tRegraUfPessoa.Update

												log_via_vetor_carrega_do_recordset tRegraUfPessoa, vLog1, campos_a_omitir
												s_log_aux = log_via_vetor_monta_inclusao(vLog1)
												if s_log_aux <> "" then
													if s_log <> "" then s_log = s_log & chr(13)
													s_log = s_log & MARGEM_N2 & "Inclusão: UF=" & oRegra.vUF(iUF).uf & ", tipo_pessoa=" & oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa & ": " & s_log_aux
													end if
												end if
											end if
										end if
									
									if alerta = "" then
										for iCD=LBound(oRegra.vUF(iUF).vPessoa(iPessoa).vCD) to UBound(oRegra.vUF(iUF).vPessoa(iPessoa).vCD)
											if oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).id_nfe_emitente <> 0 then
												s = "SELECT * FROM t_WMS_REGRA_CD_X_UF_X_PESSOA_X_CD WHERE (id_wms_regra_cd_x_uf_x_pessoa = " & id_regra_cd_uf_pessoa & ") AND (id_nfe_emitente = " & oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).id_nfe_emitente & ")"
												if tRegraUfPessoaCd.State <> 0 then tRegraUfPessoaCd.Close
												tRegraUfPessoaCd.Open s, cn
												if Not tRegraUfPessoaCd.Eof then
													id_regra_cd_uf_pessoa_cd = tRegraUfPessoaCd("id")
													log_via_vetor_carrega_do_recordset tRegraUfPessoaCd, vLog1, campos_a_omitir
													tRegraUfPessoaCd("ordem_prioridade") = converte_numero(oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).ordem_prioridade)
													tRegraUfPessoaCd("st_inativo") = converte_numero(oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).st_inativo)
													tRegraUfPessoaCd.Update
													log_via_vetor_carrega_do_recordset tRegraUfPessoaCd, vLog2, campos_a_omitir
													s_log_aux = log_via_vetor_monta_alteracao(vLog1, vLog2)
													if s_log_aux <> "" then
														if s_log <> "" then s_log = s_log & chr(13)
														s_log = s_log & MARGEM_N3 & "Alteração: UF=" & oRegra.vUF(iUF).uf & ", tipo_pessoa=" & oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa & ", id_nfe_emitente=" & oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).id_nfe_emitente & ": " & s_log_aux
														end if
												else
													if Not fin_gera_nsu(NSU_WMS_REGRA_CD_X_UF_X_PESSOA_X_CD, id_regra_cd_uf_pessoa_cd, msg_erro) then
														alerta = "Falha ao tentar gerar o NSU para a tabela " & NSU_WMS_REGRA_CD_X_UF_X_PESSOA_X_CD
														erro_fatal = True
														end if

													if alerta = "" then
														tRegraUfPessoaCd.AddNew
														tRegraUfPessoaCd("id") = id_regra_cd_uf_pessoa_cd
														tRegraUfPessoaCd("id_wms_regra_cd_x_uf_x_pessoa") = id_regra_cd_uf_pessoa
														tRegraUfPessoaCd("id_nfe_emitente") = converte_numero(oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).id_nfe_emitente)
														tRegraUfPessoaCd("ordem_prioridade") = converte_numero(oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).ordem_prioridade)
														tRegraUfPessoaCd("st_inativo") = converte_numero(oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).st_inativo)
														tRegraUfPessoaCd.Update

														log_via_vetor_carrega_do_recordset tRegraUfPessoaCd, vLog1, campos_a_omitir
														s_log_aux = log_via_vetor_monta_inclusao(vLog1)
														if s_log_aux <> "" then
															if s_log <> "" then s_log = s_log & chr(13)
															s_log = s_log & MARGEM_N3 & "Inclusão: UF=" & oRegra.vUF(iUF).uf & ", tipo_pessoa=" & oRegra.vUF(iUF).vPessoa(iPessoa).tipo_pessoa & ", id_nfe_emitente=" & oRegra.vUF(iUF).vPessoa(iPessoa).vCD(iCD).id_nfe_emitente & ": " & s_log_aux
															end if
														end if
													end if
												end if
											next 'for iCD
										end if
									
									if alerta <> "" then exit for
									next 'for iPessoa
								end if
							end if

						if alerta <> "" then exit for
						next 'for iUF
					end if

				if Err <> 0 then
					erro_fatal=True
					alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
					end if

				if alerta = "" then
					if s_log <> "" then
						grava_log usuario, "", "", "", OP_LOG_MULTI_CD_REGRA_ALTERACAO, s_log
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
		s = "<p style='margin:5px 2px 5px 2px;'>" & alerta & "</P>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "REGRA " & chr(34) & s_apelido_regra & chr(34) & " CADASTRADA COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "REGRA " & chr(34) & s_apelido_regra & chr(34) & " ALTERADA COM SUCESSO."
			case OP_EXCLUI
				s = "REGRA " & chr(34) & s_apelido_regra & chr(34) & " EXCLUÍDA COM SUCESSO."
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
	s="MultiCDRegraMenu.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>
</body>
</html>


<%
	if tRegra.State <> 0 then tRegra.Close
	set tRegra = nothing

	if tRegraUf.State <> 0 then tRegraUf.Close
	set tRegraUf = nothing

	if tRegraUfPessoa.State <> 0 then tRegraUfPessoa.Close
	set tRegraUfPessoa = nothing

	if tRegraUfPessoaCd.State <> 0 then tRegraUfPessoaCd.Close
	set tRegraUfPessoaCd = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>