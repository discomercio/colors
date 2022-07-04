<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelComissaoIndicadoresNFSeP07ProcessarBotaoMagico.asp
'     =================================================================
'
'
'	  S E R V E R   S I D E   S C R I P T I N G
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


	On Error GoTo 0
	Err.Clear
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	const ID_RELATORIO = "CENTRAL/RelComissaoIndicadoresNFSe"

'	COMO O TRATAMENTO DO RELAT�RIO PODE SER DEMORADO, CASO A SESS�O EXPIRE E O TRATAMENTO
'	DE SESS�O EXPIRADA N�O CONSIGA RESTAUR�-LA, OBT�M A IDENTIFICA��O DO USU�RIO A PARTIR DE
'	UM CAMPO HIDDEN CRIADO NA P�GINA CHAMADORA EXCLUSIVAMENTE P/ ISSO.
	dim s, s_aux, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	if (usuario = "") then usuario = Trim(Request("c_usuario_sessao"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISS�O DE ACESSO DO USU�RIO
	if Not operacao_permitida(OP_CEN_FLAG_COMISSAO_PAGA, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
'	FILTROS
	dim c_cnpj_nfse
	dim ckb_id_indicador
	dim rb_visao

	c_cnpj_nfse = retorna_so_digitos(Request.Form("c_cnpj_nfse"))
	ckb_id_indicador = Trim(Request.Form("ckb_id_indicador"))
	rb_visao = Trim(Request.Form("rb_visao"))

	dim id_nsu_N1, intRecordsAffected, id_cfg_tabela_origem, proc_comissao_request_guid, proc_fluxo_caixa_request_guid
	id_nsu_N1 = Trim(Request.Form("id_nsu_N1"))
	proc_comissao_request_guid = Trim(Request.Form("proc_comissao_request_guid"))
	proc_fluxo_caixa_request_guid = Trim(Request.Form("proc_fluxo_caixa_request_guid"))

	dim alerta
	alerta=""

	dim mensagem, s_rel_comissao_paga, s_rel_devolucao_descontada, s_rel_perda_descontada
	mensagem = ""
	s_rel_comissao_paga = ""
	s_rel_devolucao_descontada = ""
	s_rel_perda_descontada = ""

	dim blnErroFatal
	blnErroFatal = False

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, cn2, rs, tN1, tN2, tIndicador, tN3Ped, tFC
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not bdd_conecta_RPIFC(cn2) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tN1, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tN2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tN3Ped, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tIndicador, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim vl_total_RT_RALiq
	dim c_numero_nfse, dt_competencia, dt_mes_competencia, fluxo_caixa_plano_contas_grupo_RT
	dim c_fluxo_caixa_dt_competencia, c_fluxo_caixa_conta_corrente, c_fluxo_caixa_empresa, c_fluxo_caixa_plano_contas_RT
	c_numero_nfse = retorna_so_digitos(Request.Form("c_numero_nfse"))
	c_fluxo_caixa_dt_competencia = Trim(Request.Form("c_fluxo_caixa_dt_competencia"))
	c_fluxo_caixa_conta_corrente = Trim(Request.Form("c_fluxo_caixa_conta_corrente"))
	c_fluxo_caixa_empresa = Trim(Request.Form("c_fluxo_caixa_empresa"))
	c_fluxo_caixa_plano_contas_RT = Trim(Request.Form("c_fluxo_caixa_plano_contas_RT"))

	if alerta = "" then
		if c_numero_nfse = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O n�mero da NFS-e n�o foi informado."
		else
			if converte_numero(c_numero_nfse) = 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O n�mero da NFS-e informado � inv�lido."
				end if
			end if
		end if 'if alerta = ""

	if alerta = "" then
		if c_fluxo_caixa_dt_competencia = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "A data de compet�ncia n�o foi informada."
		else
			dt_competencia = StrToDate(c_fluxo_caixa_dt_competencia)
			if Not IsDate(dt_competencia) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Data de compet�ncia informada � inv�lida."
				end if
			end if
		end if 'if alerta = ""

	if alerta = "" then
		if c_fluxo_caixa_conta_corrente = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "A conta corrente n�o foi informada."
		else
			s = "SELECT * FROM t_FIN_CONTA_CORRENTE WHERE (id = " & c_fluxo_caixa_conta_corrente & ")"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn2
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "A conta corrente ID=" & c_fluxo_caixa_conta_corrente & " n�o foi encontrada no banco de dados."
				end if
			end if
		end if 'if alerta = ""

	if alerta = "" then
		if c_fluxo_caixa_empresa = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "A empresa para o lan�amento no fluxo de caixa n�o foi informada."
		else
			s = "SELECT * FROM t_FIN_PLANO_CONTAS_EMPRESA WHERE (id = " & c_fluxo_caixa_empresa & ")"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn2
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "A empresa ID=" & c_fluxo_caixa_empresa & " n�o foi encontrada no banco de dados."
				end if
			end if
		end if 'if alerta = ""

	'Emily em 01/07/2022: os valores de comiss�o (RT) e RA L�quido est�o sendo somados e o valor total est� sendo lan�ado na conta 1400
	if alerta = "" then
		if c_fluxo_caixa_plano_contas_RT = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O plano de contas para comiss�o (RT) n�o foi informado."
		else
			s = "SELECT * FROM t_FIN_PLANO_CONTAS_CONTA WHERE (id = " & c_fluxo_caixa_plano_contas_RT & ") AND (natureza = '" & COD_FIN_NATUREZA__DEBITO & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn2
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O plano de contas ID=" & c_fluxo_caixa_plano_contas_RT & " n�o foi encontrado no banco de dados."
			else
				fluxo_caixa_plano_contas_grupo_RT = rs("id_plano_contas_grupo")
				end if
			end if
		end if 'if alerta = ""

	dim max_fc_descricao
	max_fc_descricao = 0
	if alerta = "" then
		'OBT�M O TAMANHO DO CAMPO t_FIN_FLUXO_CAIXA.descricao
		s = "SELECT" & _
				" sc.length" & _
			" FROM syscolumns sc" & _
				" INNER JOIN sysobjects so ON (sc.id = so.id)" & _
			" WHERE" & _
				" (so.type = 'U')" & _
				" AND (so.name = 't_FIN_FLUXO_CAIXA')" & _
				" AND (sc.name = 'descricao')"
		if rs.State <> 0 then rs.Close
		rs.Open s, cn2
		if rs.Eof then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Falha ao tentar determinar o tamanho do campo t_FIN_FLUXO_CAIXA.descricao"
		else
			max_fc_descricao = rs("length")
			end if
		end if 'if alerta = ""

	if alerta = "" then
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_fluxo_caixa_dt_competencia", c_fluxo_caixa_dt_competencia)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_fluxo_caixa_conta_corrente", c_fluxo_caixa_conta_corrente)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_fluxo_caixa_empresa", c_fluxo_caixa_empresa)
		call set_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_fluxo_caixa_plano_contas_RT", c_fluxo_caixa_plano_contas_RT)
		end if 'if alerta = ""

	if alerta = "" then
	'	TRATAMENTO P/ OS CASOS EM QUE: USU�RIO EST� TENTANDO USAR O BOT�O VOLTAR, OCORREU DUPLO CLIQUE OU USU�RIO ATUALIZOU A P�GINA ENQUANTO AINDA ESTAVA PROCESSANDO (DUPLO ACIONAMENTO)
	'	Esse tratamento � feito atrav�s do campo proc_fluxo_caixa_request_guid (t_COMISSAO_INDICADOR_NFSe_N1.proc_fluxo_caixa_request_guid)
		if proc_fluxo_caixa_request_guid <> "" then
			s = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N1 WHERE (proc_fluxo_caixa_request_guid = '" & proc_fluxo_caixa_request_guid & "')"
			tN1.Open s, cn
			if Not tN1.Eof then
				blnErroFatal = True
				alerta=texto_add_br(alerta)
				alerta=alerta & "Este relat�rio j� processou os lan�amentos do fluxo de caixa em " & formata_data_hora_sem_seg(tN1("proc_fluxo_caixa_data_hora")) & " por " & Trim("" & tN1("proc_fluxo_caixa_usuario")) & " (NSU = " & Trim("" & tN1("id")) & ")" & _
								"<br /><br />" & _
								"<a style='color:black;' href='javascript:fRelSumario(fSumario," & Trim("" & tN1("id")) & ")'><button type='button' class='Button C'>Consultar Detalhes</button></a>"
				end if
			end if 'if proc_fluxo_caixa_request_guid <> ""
		end if 'if alerta = ""

	if alerta = "" then
		s = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N1 WHERE (id = " & id_nsu_N1 & ")"
		if tN1.State <> 0 then tN1.Close
		tN1.Open s, cn
		if tN1.Eof then
			blnErroFatal = True
			alerta=texto_add_br(alerta)
			alerta=alerta & "Falha ao tentar localizar dados do relat�rio (NSU = " & id_nsu_N1 & ")"
		else
			if c_cnpj_nfse = "" then c_cnpj_nfse = Trim("" & tN1("NFSe_cnpj"))
			if tN1("proc_comissao_status") = 0 then
				blnErroFatal = True
				alerta=texto_add_br(alerta)
				alerta=alerta & "O relat�rio ainda n�o processou o pagamento das comiss�es nos pedidos (NSU = " & id_nsu_N1 & ")" & _
								"<br /><br />" & _
								"<a style='color:black;' href='javascript:fRelSumario(fSumario," & Trim("" & tN1("id")) & ")'><button type='button' class='Button C'>Consultar Detalhes</button></a>"
			elseif tN1("proc_fluxo_caixa_status") <> 0 then
				blnErroFatal = True
				alerta=texto_add_br(alerta)
				alerta=alerta & "Este relat�rio j� processou os lan�amentos do fluxo de caixa em " & formata_data_hora_sem_seg(tN1("proc_fluxo_caixa_data_hora")) & " por " & Trim("" & tN1("proc_fluxo_caixa_usuario")) & " (NSU = " & Trim("" & tN1("id")) & ")" & _
								"<br /><br />" & _
								"<a style='color:black;' href='javascript:fRelSumario(fSumario," & Trim("" & tN1("id")) & ")'><button type='button' class='Button C'>Consultar Detalhes</button></a>"
				end if
			end if

		if alerta = "" then
			'Data de Compet�ncia 2: registra o m�s de compet�ncia do relat�rio
			s = "01/" & normaliza_codigo(Trim("" & tN1("competencia_mes")), 2) & "/" & Trim("" & tN1("competencia_ano"))
			dt_mes_competencia = StrToDate(s)
			end if 'if alerta = ""
		
		if alerta = "" then
			vl_total_RT_RALiq = tN1("vl_total_geral_selecionado_RT") + tN1("vl_total_geral_selecionado_RA_liquido")
			if vl_total_RT_RALiq = 0 then
				blnErroFatal = True
				alerta=texto_add_br(alerta)
				alerta=alerta & "O valor total da comiss�o (RT + RA l�quido) � zero!" & _
								"<br /><br />" & _
								"<a style='color:black;' href='javascript:fRelSumario(fSumario," & Trim("" & tN1("id")) & ")'><button type='button' class='Button C'>Consultar Detalhes</button></a>"
			elseif vl_total_RT_RALiq < 0 then
				blnErroFatal = True
				alerta=texto_add_br(alerta)
				alerta=alerta & "O valor total da comiss�o (RT + RA l�quido) � negativo!" & _
								"<br /><br />" & _
								"<a style='color:black;' href='javascript:fRelSumario(fSumario," & Trim("" & tN1("id")) & ")'><button type='button' class='Button C'>Consultar Detalhes</button></a>"
				end if
			end if 'if alerta = ""
		end if 'if alerta = ""

	dim s_comissao_NFSe_razao_social, s_fc_descricao
	s_comissao_NFSe_razao_social = ""

	if alerta = "" then
		s = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N2 WHERE (id_comissao_indicador_nfse_n1 = " & id_nsu_N1 & ")"
		tN2.Open s, cn
		if tN2.Eof then
			blnErroFatal = True
			alerta=texto_add_br(alerta)
			alerta=alerta & "Falha ao tentar localizar dados do relat�rio com a identifica��o do indicador (NSU = " & id_nsu_N1 & ")"
		else
			s = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (Id = " & Trim ("" & tN2("id_indicador")) & ")"
			tIndicador.Open s, cn
			if tIndicador.Eof then
				blnErroFatal = True
				alerta=texto_add_br(alerta)
				alerta=alerta & "Falha ao tentar localizar dados cadastrais do indicador (ID = " & Trim ("" & tN2("id_indicador")) & ")"
			else
				s_comissao_NFSe_razao_social = Trim("" & tIndicador("comissao_NFSe_razao_social"))
				end if
			end if
		end if 'if alerta = ""


	dim qtde_insert_fc, msg_sucesso
	qtde_insert_fc = 0
	msg_sucesso = ""

	dim s_log
	s_log = ""

	dim intNsuNovoFluxoCaixa

	'Registra no relat�rio os dados definidos pelo usu�rio para o(s) lan�amento(s) no fluxo de caixa
	if alerta = "" then
		tN1("fluxo_caixa_dt_competencia") = dt_competencia
		tN1("fluxo_caixa_id_conta_corrente") = CInt(c_fluxo_caixa_conta_corrente)
		tN1("fluxo_caixa_id_plano_contas_empresa") = CInt(c_fluxo_caixa_empresa)
		tN1("fluxo_caixa_comissao_id_plano_contas_conta") = CLng(c_fluxo_caixa_plano_contas_RT)
		'Emily em 01/07/2022: os valores de comiss�o (RT) e RA L�quido est�o sendo somados e o valor total est� sendo lan�ado na conta 1400
		'tN1("fluxo_caixa_RA_id_plano_contas_conta") = CLng(c_fluxo_caixa_plano_contas_RA)
		tN1("NFSe_numero") = converte_numero(c_numero_nfse)
		tN1.Update
		end if 'if alerta = ""

	'GRAVA LAN�AMENTOS NO FLUXO DE CAIXA
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if alerta = "" then
		'REGISTRA A QUANTIDADE DE TENTATIVAS FORA DA TRANSA��O
		s = "UPDATE t_COMISSAO_INDICADOR_NFSe_N1 SET" & _
				" proc_fluxo_caixa_qtde_tentativas = proc_fluxo_caixa_qtde_tentativas + 1" & _
			" WHERE" & _
				" (id = " & id_nsu_N1 & ")"
		cn.Execute s, intRecordsAffected

	'	~~~~~~~~~~~~~~~~~~~~~~~~~
		cn.BeginTrans
		cn2.Execute("BEGIN TRAN")
	'	~~~~~~~~~~~~~~~~~~~~~~~~~

		If Not cria_recordset_pessimista(tFC, msg_erro) then
		'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			cn.RollbackTrans
			cn2.Execute("ROLLBACK TRAN")
		'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

	'	TRATAMENTO P/ OS CASOS EM QUE: USU�RIO EST� TENTANDO USAR O BOT�O VOLTAR, OCORREU DUPLO CLIQUE OU USU�RIO ATUALIZOU A P�GINA ENQUANTO AINDA ESTAVA PROCESSANDO (DUPLO ACIONAMENTO)
	'	Esse tratamento � feito atrav�s do campo proc_fluxo_caixa_request_guid (t_COMISSAO_INDICADOR_NFSe_N1.proc_fluxo_caixa_request_guid)
		if proc_fluxo_caixa_request_guid <> "" then
			s = "UPDATE t_COMISSAO_INDICADOR_NFSe_N1 SET" & _
					" proc_fluxo_caixa_request_guid = '" & proc_fluxo_caixa_request_guid & "'" & _
				" WHERE" & _
					" (id = " & id_nsu_N1 & ")"
			cn.Execute s, intRecordsAffected
			if intRecordsAffected <> 1 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Falha ao tentar atualizar o registro do relat�rio no banco de dados (NSU = " & id_nsu_N1 & ")!!<br />Processamento cancelado!!"
				end if
			end if

		if alerta = "" then
			'Emily em 01/07/2022: os valores de comiss�o (RT) e RA L�quido est�o sendo somados e o valor total est� sendo lan�ado na conta 1400
			if vl_total_RT_RALiq > 0 then
				if Not fin_gera_nsu_fluxo_caixa(T_FIN_FLUXO_CAIXA, cn2, intNsuNovoFluxoCaixa, msg_erro) then
					alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
				else
					if intNsuNovoFluxoCaixa <= 0 then
						alerta = "NSU GERADO � INV�LIDO (" & intNsuNovoFluxoCaixa & ")"
						end if
					end if

				if alerta = "" then
					s = "SELECT * FROM t_FIN_FLUXO_CAIXA WHERE (id= -1)"
					if tFC.State <> 0 then tFC.Close
					tFC.Open s, cn2
					tFC.AddNew
					tFC("id") = intNsuNovoFluxoCaixa
					tFC("id_conta_corrente") = CInt(c_fluxo_caixa_conta_corrente)
					tFC("id_plano_contas_empresa") = CInt(c_fluxo_caixa_empresa)
					tFC("natureza") = COD_FIN_NATUREZA__DEBITO
					tFC("st_sem_efeito") = 0
					tFC("id_plano_contas_grupo") = CLng(fluxo_caixa_plano_contas_grupo_RT)
					tFC("id_plano_contas_conta") = CLng(c_fluxo_caixa_plano_contas_RT)
					tFC("valor") = vl_total_RT_RALiq
					tFC("dt_competencia") = dt_competencia
					tFC("tipo_cadastro") = "S"
					tFC("editado_manual") = "N"
					tFC("dt_cadastro") = Date
					tFC("dt_hr_cadastro") = Now
					tFC("usuario_cadastro") = usuario
					tFC("dt_ult_atualizacao") = Date
					tFC("dt_hr_ult_atualizacao") = Now
					tFC("usuario_ult_atualizacao") = usuario
					tFC("ctrl_pagto_id_parcela") = tN1("id")
					tFC("ctrl_pagto_modulo") = COD_FIN_FLUXO_CAIXA_CTRL_PAGTO_MODULO__COMISSAO_INDICADOR_VIA_NFSe
					tFC("ctrl_pagto_status") = 1
					tFC("ctrl_pagto_id_ambiente_origem") = ID_AMBIENTE
					tFC("dt_mes_competencia") = dt_mes_competencia
					'Formato: [Raz�o social do emitente da NFS-e] REF MM
					s_fc_descricao = " REF " & normaliza_codigo(Trim("" & tN1("competencia_mes")), 2)
					s_fc_descricao = Left(s_comissao_NFSe_razao_social, (max_fc_descricao - Len(s_fc_descricao))) & s_fc_descricao
					tFC("descricao") = s_fc_descricao
					tFC("numero_NF") = converte_numero(c_numero_nfse)
					tFC("cnpj_cpf") = Trim("" & tN1("NFSe_cnpj"))
					tFC.Update
					qtde_insert_fc = qtde_insert_fc + 1
					if s_log <> "" then s_log = s_log & "; "
					s_log = s_log & "Grava��o no fluxo de caixa do lan�amento de comiss�o:" & _
									" t_FIN_FLUXO_CAIXA.id=" & Trim("" & tFC("id")) & _
									", id_conta_corrente=" & Trim("" & tFC("id_conta_corrente")) & _
									", id_plano_contas_empresa=" & Trim("" & tFC("id_plano_contas_empresa")) & _
									", id_plano_contas_grupo=" & Trim("" & tFC("id_plano_contas_grupo")) & _
									", id_plano_contas_conta=" & Trim("" & tFC("id_plano_contas_conta")) & _
									", valor=" & formata_moeda(tFC("valor")) & _
									", dt_competencia=" & formata_data(tFC("dt_competencia")) & _
									", numero_NF=" & tFC("numero_NF") & _
									", cnpj_cpf=" & tFC("cnpj_cpf") & _
									", descricao=" & Trim("" & tFC("descricao"))
					end if 'if alerta = ""
				end if 'if tN1("vl_total_geral_selecionado_RT") > 0
			end if 'if alerta = ""

		if alerta = "" then
			s = "UPDATE t_COMISSAO_INDICADOR_NFSe_N1 SET" & _
					" status = 2" & _
					", dt_hr_ult_atualizacao = getdate()" & _
					", usuario_ult_atualizacao = '" & QuotedStr(usuario) & "'" & _
					", proc_fluxo_caixa_status = 1" & _
					", proc_fluxo_caixa_data = Convert(datetime, Convert(varchar(10),getdate(), 121), 121)" & _
					", proc_fluxo_caixa_data_hora = getdate()" & _
					", proc_fluxo_caixa_usuario = '" & QuotedStr(usuario) & "'" & _
				" WHERE" & _
					" (id = " & id_nsu_N1 & ")" & _
					" AND (status = 1)" & _
					" AND (proc_fluxo_caixa_status = 0)"
			cn.Execute s, intRecordsAffected
			if intRecordsAffected <> 1 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Falha ao tentar atualizar o status do relat�rio no banco de dados (NSU = " & id_nsu_N1 & ")!!<br />Processamento cancelado!!"
				end if
			end if 'if alerta = ""

		if alerta = "" then
			'GRAVA LOG
			if s_log <> "" then
				s_log = "Relat�rio de Pedidos Indicadores via NFS-e (t_COMISSAO_INDICADOR_NFSe_N1.id=" & id_nsu_N1 & "): " & s_log
				grava_log usuario, "", "", "", OP_LOG_REL_COMISSAO_INDICADORES_NFSe_LANCTO_FC, s_log
				end if

			msg_sucesso = "Processamento autom�tico no fluxo de caixa realizado com sucesso: " & Cstr(qtde_insert_fc)
			if qtde_insert_fc = 1 then
				msg_sucesso = msg_sucesso & " lan�amento gravado"
			else
				msg_sucesso = msg_sucesso & " lan�amentos gravados"
				end if

		'	~~~~~~~~~~~~~~~~~~~~~~~~~~
			cn.CommitTrans
			cn2.Execute("COMMIT TRAN")
		'	~~~~~~~~~~~~~~~~~~~~~~~~~~
			if Err=0 then
				' NOP: Response.Redirect("resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			cn.RollbackTrans
			cn2.Execute("ROLLBACK TRAN")
		'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			end if 'if alerta = ""
		end if 'if alerta = ""





' _____________________________________________________________________________________________
'
'									F  U  N  �  �  E  S 
' _____________________________________________________________________________________________

' ___________________________________________
' FIN GERA NSU FLUXO CAIXA
'
function fin_gera_nsu_fluxo_caixa(byval idNsu, byref cnRef, byref nsu, byref msg_erro)
dim t, strSql, intRetorno, intRecordsAffected
dim intQtdeTentativas, intNsuUltimo, intNsuNovo, blnSucesso
	fin_gera_nsu_fluxo_caixa = False
	msg_erro=""
	nsu=0
	strSql = "SELECT" & _
				" Count(*) AS qtde" & _
			" FROM t_FIN_CONTROLE" & _
			" WHERE" & _
				" (id='" & idNsu & "')"
	set t=cnRef.Execute(strSql)
	if Not t.Eof then intRetorno=Clng(t("qtde")) else intRetorno=Clng(0)

'	N�O EST� CADASTRADO, ENT�O CADASTRA AGORA
	if intRetorno=0 then
		strSql = "INSERT INTO t_FIN_CONTROLE (" & _
					"id, " & _
					"nsu, " & _
					"dt_hr_ult_atualizacao" & _
				") VALUES (" & _
					"'" & idNsu & "'," & _
					"0," & _
					"getdate()" & _
				")"
		cnRef.Execute strSql, intRecordsAffected
		if intRecordsAffected <> 1 then
			msg_erro = "Falha ao criar o registro para gera��o de NSU (" & idNsu & ")!!"
			exit function
			end if
		end if

'	LA�O DE TENTATIVAS PARA GERAR O NSU (DEVIDO A ACESSO CONCORRENTE)
	intQtdeTentativas=0
	do 
		intQtdeTentativas = intQtdeTentativas + 1
		
	'	OBT�M O �LTIMO NSU USADO
		strSql = "SELECT" & _
					" nsu" & _
				" FROM t_FIN_CONTROLE" & _
				" WHERE" & _
					" id = '" & idNsu & "'"
		set t=cnRef.Execute(strSql)
		if t.Eof then
			strMsgErro = "Falha ao localizar o registro para gera��o de NSU (" & idNsu & ")!!"
			Exit Function
		else
			intNsuUltimo = Clng(t("nsu"))
			end if

	'	INCREMENTA 1
		intNsuNovo = intNsuUltimo + 1
		
	'	TENTA ATUALIZAR O BANCO DE DADOS
		strSql = "UPDATE t_FIN_CONTROLE SET" & _
					" nsu = " & CStr(intNsuNovo) & "," & _
					" dt_hr_ult_atualizacao = getdate()" & _
				" WHERE" & _
					" (id = '" & idNsu & "')" & _
					" AND (nsu = " & CStr(intNsuUltimo) & ")"
		cnRef.Execute strSql, intRecordsAffected
		If intRecordsAffected = 1 Then
			blnSucesso = True
			nsu = intNsuNovo
			end if
		
		Loop While (Not blnSucesso) And (intQtdeTentativas < 10)

	If Not blnSucesso Then
		strMsgErro = "Falha ao tentar gerar o NSU!!"
		Exit Function
		End If
	
	fin_gera_nsu_fluxo_caixa = True

end function


function fluxo_caixa_conta_corrente_monta_descricao(byval id_conta_corrente)
dim r, s, strResp

	fluxo_caixa_conta_corrente_monta_descricao = ""

	id_conta_corrente = Trim("" & id_conta_corrente)
	if id_conta_corrente = "" then exit function

	s = "SELECT * FROM t_FIN_CONTA_CORRENTE WHERE (id = " & id_conta_corrente & ")"
	set r = cn2.Execute(s)

	strResp = ""
	if Not r.Eof then
		strResp = Trim("" & r("banco")) & " &nbsp; " & Trim("" & r("agencia")) & " &nbsp; " & Trim("" & r("conta")) & " &nbsp; " & Trim("" & r("descricao"))
		end if

	fluxo_caixa_conta_corrente_monta_descricao = strResp
	r.close
	set r = Nothing
end function


function fluxo_caixa_empresa_monta_descricao(byval id_empresa)
dim r, s, strResp

	fluxo_caixa_empresa_monta_descricao = ""

	id_empresa = Trim("" & id_empresa)
	if id_empresa = "" then exit function

	s = "SELECT * FROM t_FIN_PLANO_CONTAS_EMPRESA WHERE (id = " & id_empresa & ")"
	set r = cn2.Execute(s)

	strResp = ""
	if Not r.Eof then
		strResp = Trim("" & r("id")) & " - " & Trim("" & r("descricao"))
		end if

	fluxo_caixa_empresa_monta_descricao = strResp
	r.close
	set r = Nothing
end function

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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL</title>
	</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	$(function () {
	});

	function fAvisoVoltar(f) {
		f.action = "RelComissaoIndicadoresNFSeP01Filtro.asp?url_back=X";
		f.submit();
	}

	function fRetornar(f) {
		f.action = "RelComissaoIndicadoresNFSeP01Filtro.asp?url_back=X";
		dVOLTAR.style.visibility = "hidden";
		f.submit();
	}

	function fRelSumario(f, id_nsu_N1) {
		f.id_nsu_N1.value = id_nsu_N1;
		f.action = "RelComissaoIndicadoresNFSeP06BotaoMagico.asp";
		f.submit();
	}
</script>



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
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
.TdLabel{
	width:200px;
}

.TdInfo{
	width:600px;
}
</style>

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br />
<!--  T E L A  -->
<form id="fAviso" name="fAviso" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="ckb_id_indicador" id="ckb_id_indicador" value="<%=ckb_id_indicador%>" />
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>">


<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br /><br />
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center">
        <% if blnErroFatal then %>
        <a name="bVOLTAR" id="bVOLTAR" href="javascript:fAvisoVoltar(fAviso)">
        <% else %>
        <a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()">
        <% end if %>
        <img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

<form name="fSumario" id="fSumario" method="post">
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="ckb_id_indicador" id="ckb_id_indicador" value="<%=ckb_id_indicador%>" />
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>">
<input type="hidden" name="id_nsu_N1" id="id_nsu_N1" />
<input type="hidden" name="proc_comissao_request_guid" id="proc_comissao_request_guid" value="<%=proc_comissao_request_guid%>" />
<input type="hidden" name="proc_fluxo_caixa_request_guid" id="proc_fluxo_caixa_request_guid" value="<%=proc_fluxo_caixa_request_guid%>" />
</form>

</center>
</body>

<% else %>

<!-- ***************************************************** -->
<!-- **********  P�GINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Conclu�do';" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="f" name="f" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="ckb_id_indicador" id="ckb_id_indicador" value="<%=ckb_id_indicador%>" />
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>">
<input type="hidden" name="id_nsu_N1" id="id_nsu_N1" value="<%=id_nsu_N1%>" />
<input type="hidden" name="proc_comissao_request_guid" id="proc_comissao_request_guid" value="<%=proc_comissao_request_guid%>" />
<input type="hidden" name="proc_fluxo_caixa_request_guid" id="proc_fluxo_caixa_request_guid" value="<%=gera_uid%>" />


<!--  I D E N T I F I C A � � O   D A   T E L A  -->
<table width="820" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
	<tr>
		<td align="right" valign="bottom"><span class="PEDIDO">Relat�rio de Pedidos Indicadores (via NFS-e)</span></td>
	</tr>
</table>
<br />
<br />

<%
	s = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N1 WHERE (id = " & id_nsu_N1 & ")"
	if tN1.State <> 0 then tN1.Close
	tN1.Open s, cn
	if tN1.Eof then
		mensagem = "Falha ao tentar localizar dados do relat�rio (NSU = " & id_nsu_N1 & ")"
	else
		if tN1("proc_fluxo_caixa_status") = 0 then
			if tN1("proc_fluxo_caixa_qtde_tentativas") > 0 then
				mensagem = "O processamento autom�tico no fluxo de caixa N�O foi realizado devido a falhas"
				end if
			end if
		end if
%>

<% if msg_sucesso <> "" then %>
<!-- ************   MENSAGEM DE SUCESSO ************ -->
<div class='MtAviso' style='width:800px;font-weight:bold;' align='center'>
<p style='margin:5px 2px 5px 2px;'><%=msg_sucesso%></p></div>
<br />
<br />
<% end if %>


<!-- ************   MENSAGEM  ************ -->
<table class="Qx" style="width:800px;" cellpadding="1" cellspacing="0">
	<tr>
		<td class="MT TdLabel" align="right"><span class="Cd">NSU do Relat�rio:</span></td>
		<td class="MTBD TdInfo" align="left"><span class="C"><%=CStr(id_nsu_N1)%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" align="right"><span class="Cd">CNPJ NFS-e:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=cnpj_cpf_formata(Trim("" & tN1("NFSe_cnpj")))%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Comiss�o Processada em:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=formata_data_e_talvez_hora_hhmm(tN1("proc_comissao_data_hora"))%> por <%=Trim("" & tN1("proc_comissao_usuario"))%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Lan�amentos Processados em:</span></td>
		<% if tN1("proc_fluxo_caixa_status") = 1 then %>
		<td class="MDB TdInfo" align="left"><span class="C"><%=formata_data_e_talvez_hora_hhmm(tN1("proc_fluxo_caixa_data_hora"))%> por <%=Trim("" & tN1("proc_fluxo_caixa_usuario"))%></span></td>
		<% else %>
		<td class="MDB TdInfo" align="left"><span class="C" style="color:red;">N�o processado</span></td>
		<% end if %>
	</tr>
	<tr>
		<%
			s = "SELECT" & _
						" tN3Ped.*" & _
					" FROM t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN3Ped.id_comissao_indicador_nfse_n2 = tN2.id)" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N1 tN1 ON (tN2.id_comissao_indicador_nfse_n1 = tN1.id)" & _
					" WHERE" & _
						" (tN1.id = " & id_nsu_N1 & ")" & _
						" AND (tN3Ped.st_selecionado = 1)" & _
						" AND (tN3Ped.id_cfg_tabela_origem = " & ID_CFG_TABELA_ORIGEM_T_PEDIDO & ")" & _
					" ORDER BY" & _
						" tN3Ped.id"
			if tN3Ped.State <> 0 then tN3Ped.Close
			tN3Ped.Open s, cn
			do while Not tN3Ped.Eof
				if s_rel_comissao_paga <> "" then s_rel_comissao_paga = s_rel_comissao_paga & ", "
				s_rel_comissao_paga = s_rel_comissao_paga & Trim("" & tN3Ped("pedido"))
				tN3Ped.MoveNext
				loop

		s_aux = s_rel_comissao_paga
		if s_aux = "" then s_aux = "(nenhum pedido)"
		%>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Comiss�o Paga:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<%
			s = "SELECT" & _
						" tN3Ped.*" & _
					" FROM t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN3Ped.id_comissao_indicador_nfse_n2 = tN2.id)" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N1 tN1 ON (tN2.id_comissao_indicador_nfse_n1 = tN1.id)" & _
					" WHERE" & _
						" (tN1.id = " & id_nsu_N1 & ")" & _
						" AND (tN3Ped.st_selecionado = 1)" & _
						" AND (tN3Ped.id_cfg_tabela_origem = " & ID_CFG_TABELA_ORIGEM_T_PEDIDO_ITEM_DEVOLVIDO & ")" & _
					" ORDER BY" & _
						" tN3Ped.id"
			if tN3Ped.State <> 0 then tN3Ped.Close
			tN3Ped.Open s, cn
			do while Not tN3Ped.Eof
				if s_rel_devolucao_descontada <> "" then s_rel_devolucao_descontada = s_rel_devolucao_descontada & ", "
				s_rel_devolucao_descontada = s_rel_devolucao_descontada & Trim("" & tN3Ped("pedido"))
				tN3Ped.MoveNext
				loop

		%>
		<% s_aux = s_rel_devolucao_descontada
			if s_aux = "" then s_aux = "(nenhum pedido)"
		%>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Devolu��o Descontada:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<%
			s = "SELECT" & _
						" tN3Ped.*" & _
					" FROM t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN3Ped.id_comissao_indicador_nfse_n2 = tN2.id)" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N1 tN1 ON (tN2.id_comissao_indicador_nfse_n1 = tN1.id)" & _
					" WHERE" & _
						" (tN1.id = " & id_nsu_N1 & ")" & _
						" AND (tN3Ped.st_selecionado = 1)" & _
						" AND (tN3Ped.id_cfg_tabela_origem = " & ID_CFG_TABELA_ORIGEM_T_PEDIDO_PERDA & ")" & _
					" ORDER BY" & _
						" tN3Ped.id"
			if tN3Ped.State <> 0 then tN3Ped.Close
			tN3Ped.Open s, cn
			do while Not tN3Ped.Eof
				if s_rel_perda_descontada <> "" then s_rel_perda_descontada = s_rel_perda_descontada & ", "
				s_rel_perda_descontada = s_rel_perda_descontada & Trim("" & tN3Ped("pedido"))
				tN3Ped.MoveNext
				loop
		%>
		<% s_aux = s_rel_perda_descontada
			if s_aux = "" then s_aux = "(nenhum pedido)"
		%>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Perda Descontada:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<% s_aux = formata_moeda(tN1("vl_total_geral_selecionado_RT"))
			if s_aux = "" then s_aux = "?"
		%>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Comiss�o (RT):</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<% s_aux = formata_moeda(tN1("vl_total_geral_selecionado_RA_liquido"))
			if s_aux = "" then s_aux = "?"
		%>
		<td class="MDBE TdLabel" align="right"><span class="Cd">RA L�quido:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<% s_aux = formata_moeda(tN1("vl_total_geral_selecionado_RT") + tN1("vl_total_geral_selecionado_RA_liquido"))
			if s_aux = "" then s_aux = "?"
		%>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Total Comiss�o (RT + RA L�quido):</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
</table>
<br />
<br />
<table class="Qx" style="width:800px;" cellpadding="1" cellspacing="0">
	<tr>
		<td class="MT TdLabel" valign="middle" align="right"><span class="Cd">N�mero NFS-e:</span></td>
		<td class="MTBD TdInfo" align="left"><span class="C"><%=Trim("" & tN1("NFSe_numero"))%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Data Compet�ncia:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=formata_data(tN1("fluxo_caixa_dt_competencia"))%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Conta Corrente:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=fluxo_caixa_conta_corrente_monta_descricao(tN1("fluxo_caixa_id_conta_corrente"))%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Empresa:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=fluxo_caixa_empresa_monta_descricao(tN1("fluxo_caixa_id_plano_contas_empresa"))%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Plano Contas:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=Trim("" & tN1("fluxo_caixa_comissao_id_plano_contas_conta"))%></span></td>
	</tr>
</table>
<br />



<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="820" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<!-- ************   LINKS: P�GINA INICIAL / ENCERRA SESS�O   ************ -->
<table class="notPrint" width="820" cellpadding="0" cellspacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOT�ES   ************ -->
<table class="notPrint" width="820" cellspacing="0">
<% if tN1("proc_fluxo_caixa_status") = 0 then %>
<tr>
	<td align="left"><div name="dVOLTAR" id="dVOLTAR"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fRetornar(f)" title="Retornar para o in�cio do relat�rio">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="right"><div name="dPROCESSAR" id="dPROCESSAR"><a name="bPROCESSAR" id="bPROCESSAR" href="javascript:fProcessar(f)" title="Gravar lan�amentos no fluxo de caixa">
		<img src="../botao/processar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% else %>
<tr>
	<td align="center"><div name="dVOLTAR" id="dVOLTAR"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fRetornar(f)" title="Retornar para o in�cio do relat�rio">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% end if %>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

	if tN1.State <> 0 then tN1.Close
	set tN1 = nothing
	
	if tN2.State <> 0 then tN2.Close
	set tN2 = nothing

	if tN3Ped.State <> 0 then tN3Ped.Close
	set tN3Ped = nothing

	if tIndicador.State <> 0 then tIndicador.Close
	set tIndicador = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn2.Close
	set cn2 = nothing

	cn.Close
	set cn = nothing
%>