<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelComissaoIndicadoresNFSeP05GravaDados.asp
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
	
	class cl_REL_COMISSAO
		dim ckb_comissao_name
		dim id_registro
		dim ckb_comissao_status
		dim comissao_status_original
		end class

	const VENDA_NORMAL = "VEN"
	const DEVOLUCAO = "DEV"
	const PERDA = "PER"

'	COMO O TRATAMENTO DO RELATÓRIO PODE SER DEMORADO, CASO A SESSÃO EXPIRE E O TRATAMENTO
'	DE SESSÃO EXPIRADA NÃO CONSIGA RESTAURÁ-LA, OBTÉM A IDENTIFICAÇÃO DO USUÁRIO A PARTIR DE
'	UM CAMPO HIDDEN CRIADO NA PÁGINA CHAMADORA EXCLUSIVAMENTE P/ ISSO.
	dim s, s_aux, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	if (usuario = "") then usuario = Trim(Request("c_usuario_sessao"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_FLAG_COMISSAO_PAGA, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim s_log, s_log_venda_normal, s_log_devolucao, s_log_perda, s_novo_status, s_novo_op
	
'	FILTROS
	dim c_cnpj_nfse
	dim ckb_id_indicador
	dim rb_visao, proc_comissao_request_guid

	c_cnpj_nfse = retorna_so_digitos(Request.Form("c_cnpj_nfse"))
	ckb_id_indicador = Trim(Request.Form("ckb_id_indicador"))
	rb_visao = Trim(Request.Form("rb_visao"))
	proc_comissao_request_guid = Trim(Request.Form("proc_comissao_request_guid"))

'	OBTÉM DADOS DO FORMULÁRIO
	dim c_lista_completa_venda_normal, c_lista_completa_devolucao, c_lista_completa_perda
	c_lista_completa_venda_normal = Trim(Request.Form("c_lista_completa_venda_normal"))
	c_lista_completa_devolucao = Trim(Request.Form("c_lista_completa_devolucao"))
	c_lista_completa_perda = Trim(Request.Form("c_lista_completa_perda"))

	'Totais dos registros selecionados para realizar conferência
	dim vl_total_geral_selecionado_preco_venda_bd, vl_total_geral_selecionado_preco_NF_bd, vl_total_geral_selecionado_RT_bd, vl_total_geral_selecionado_RA_bruto_bd, vl_total_geral_selecionado_RA_liq_bd, vl_total_geral_selecionado_RA_dif_bd
	dim c_total_geral_selecionado_preco_venda, c_total_geral_selecionado_preco_NF, c_total_geral_selecionado_RT, c_total_geral_selecionado_RA_bruto, c_total_geral_selecionado_RA_liq, c_total_geral_selecionado_RA_dif
	c_total_geral_selecionado_preco_venda = Trim(Request.Form("c_total_geral_selecionado_preco_venda"))
	c_total_geral_selecionado_preco_NF = Trim(Request.Form("c_total_geral_selecionado_preco_NF"))
	c_total_geral_selecionado_RT = Trim(Request.Form("c_total_geral_selecionado_RT"))
	c_total_geral_selecionado_RA_bruto = Trim(Request.Form("c_total_geral_selecionado_RA_bruto"))
	c_total_geral_selecionado_RA_liq = Trim(Request.Form("c_total_geral_selecionado_RA_liq"))
	c_total_geral_selecionado_RA_dif = Trim(Request.Form("c_total_geral_selecionado_RA_dif"))

	dim i, s_name, s_name_valor_oginal, s_id_registro, qtde_total_reg_update, qtde_venda_normal_update, qtde_devolucao_update, qtde_perda_update
	dim v_lista_completa_venda_normal, v_lista_completa_devolucao, v_lista_completa_perda
	v_lista_completa_venda_normal = Split(c_lista_completa_venda_normal, ";", -1)
	v_lista_completa_devolucao = Split(c_lista_completa_devolucao, ";", -1)
	v_lista_completa_perda = Split(c_lista_completa_perda, ";", -1)
	
	dim id_nsu_N1, intRecordsAffected, id_cfg_tabela_origem
	id_nsu_N1 = Trim(Request.Form("id_nsu_N1"))

	dim v_venda_normal, v_devolucao, v_perda
	
	redim v_venda_normal(0)
	set v_venda_normal(UBound(v_venda_normal)) = New cl_REL_COMISSAO
	v_venda_normal(UBound(v_venda_normal)).ckb_comissao_name = ""
	for i=LBound(v_lista_completa_venda_normal) to Ubound(v_lista_completa_venda_normal)
		if Trim(v_lista_completa_venda_normal(i)) <> "" then
			s_id_registro = Trim(v_lista_completa_venda_normal(i))
			s_name = "ckb_comissao_paga_" & VENDA_NORMAL & "_" & s_id_registro
			s_name_valor_oginal = s_name & "_original"
			if v_venda_normal(Ubound(v_venda_normal)).ckb_comissao_name <> "" then
				redim preserve v_venda_normal(Ubound(v_venda_normal)+1)
				set v_venda_normal(UBound(v_venda_normal)) = New cl_REL_COMISSAO
				end if
			with v_venda_normal(Ubound(v_venda_normal))
				.ckb_comissao_name = s_name
				.id_registro = s_id_registro
				.ckb_comissao_status = Trim(Request.Form(s_name))
				'Recupera o valor original do status
				.comissao_status_original = Trim(Request.Form(s_name_valor_oginal))
				end with
			end if
		next
	
	redim v_devolucao(0)
	set v_devolucao(UBound(v_devolucao)) = New cl_REL_COMISSAO
	v_devolucao(UBound(v_devolucao)).ckb_comissao_name = ""
	for i=LBound(v_lista_completa_devolucao) to Ubound(v_lista_completa_devolucao)
		if Trim(v_lista_completa_devolucao(i)) <> "" then
			s_id_registro = Trim(v_lista_completa_devolucao(i))
			s_name = "ckb_comissao_paga_" & DEVOLUCAO & "_" & s_id_registro
			s_name_valor_oginal = s_name & "_original"
			if v_devolucao(Ubound(v_devolucao)).ckb_comissao_name <> "" then
				redim preserve v_devolucao(Ubound(v_devolucao)+1)
				set v_devolucao(UBound(v_devolucao)) = New cl_REL_COMISSAO
				end if
			with v_devolucao(Ubound(v_devolucao))
				.ckb_comissao_name = s_name
				.id_registro = s_id_registro
				.ckb_comissao_status = Trim(Request.Form(s_name))
				'Recupera o valor original do status
				.comissao_status_original = Trim(Request.Form(s_name_valor_oginal))
				end with
			end if
		next
	
	redim v_perda(0)
	set v_perda(UBound(v_perda)) = New cl_REL_COMISSAO
	v_perda(UBound(v_perda)).ckb_comissao_name = ""
	for i=LBound(v_lista_completa_perda) to Ubound(v_lista_completa_perda)
		if Trim(v_lista_completa_perda(i)) <> "" then
			s_id_registro = Trim(v_lista_completa_perda(i))
			s_name = "ckb_comissao_paga_" & PERDA & "_" & s_id_registro
			s_name_valor_oginal = s_name & "_original"
			if v_perda(Ubound(v_perda)).ckb_comissao_name <> "" then
				redim preserve v_perda(Ubound(v_perda)+1)
				set v_perda(UBound(v_perda)) = New cl_REL_COMISSAO
				end if
			with v_perda(Ubound(v_perda))
				.ckb_comissao_name = s_name
				.id_registro = s_id_registro
				.ckb_comissao_status = Trim(Request.Form(s_name))
				'Recupera o valor original do status
				.comissao_status_original = Trim(Request.Form(s_name_valor_oginal))
				end with
			end if
		next
	
	dim s_rel_comissao_paga, s_rel_devolucao_descontada, s_rel_perda_descontada
	s_rel_comissao_paga = ""
	s_rel_devolucao_descontada = ""
	s_rel_perda_descontada = ""
	
	dim alerta
	alerta=""
	
	dim vl_total_RT_RALiq
	vl_total_RT_RALiq = converte_numero(c_total_geral_selecionado_RT) + converte_numero(c_total_geral_selecionado_RA_liq)
	
	if alerta = "" then
		if vl_total_RT_RALiq = 0 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O valor total da comissão (RT + RA líquido) é zero"
			end if

		if vl_total_RT_RALiq < 0 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O valor total da comissão (RT + RA líquido) é negativo"
			end if
		end if 'if alerta = ""

	dim blnErroFatal, blnRelJaProcessado
	blnErroFatal = False
	blnRelJaProcessado = False

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, tN1, tN3Ped
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	s = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N1 WHERE (id = " & id_nsu_N1 & ")"
	rs.Open s, cn
	if rs.Eof then
		blnErroFatal = True
		alerta=texto_add_br(alerta)
		alerta=alerta & "Falha ao tentar localizar dados do relatório (NSU = " & id_nsu_N1 & ")"
	else
		if rs("proc_comissao_status") <> 0 then
			blnRelJaProcessado = True
			end if
		end if
	
	if rs.State <> 0 then rs.Close

	if (Not blnErroFatal) And (Not blnRelJaProcessado) then
	'	TRATAMENTO P/ OS CASOS EM QUE: USUÁRIO ESTÁ TENTANDO USAR O BOTÃO VOLTAR, OCORREU DUPLO CLIQUE OU USUÁRIO ATUALIZOU A PÁGINA ENQUANTO AINDA ESTAVA PROCESSANDO (DUPLO ACIONAMENTO)
	'	Esse tratamento é feito através do campo proc_comissao_request_guid (t_COMISSAO_INDICADOR_NFSe_N1.proc_comissao_request_guid)
		if proc_comissao_request_guid <> "" then
			s = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N1 WHERE (proc_comissao_request_guid = '" & proc_comissao_request_guid & "')"
			rs.Open s, cn
			if Not rs.Eof then
				blnErroFatal = True
				alerta=texto_add_br(alerta)
				alerta=alerta & "Este relatório já foi processado em " & formata_data_hora_sem_seg(rs("proc_comissao_data_hora")) & " por " & Trim("" & rs("proc_comissao_usuario")) & " (NSU = " & Trim("" & rs("id")) & ")" & _
								"<br /><br />" & _
								"<a style='color:black;' href='javascript:fRelSumario(fSumario," & Trim("" & rs("id")) & ")'><button type='button' class='Button C'>Consultar Detalhes</button></a>"
				end if

			if rs.State <> 0 then rs.Close
			end if 'if proc_comissao_request_guid <> ""
		end if 'if (Not blnErroFatal) And (Not blnRelJaProcessado)

	s_log = ""
	qtde_total_reg_update = 0
	qtde_venda_normal_update = 0
	qtde_devolucao_update = 0
	qtde_perda_update = 0
	
	if (alerta = "") And (Not blnErroFatal) And (Not blnRelJaProcessado) then
		if rs.State <> 0 then rs.Close
		set rs = nothing
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		If Not cria_recordset_pessimista(rs, msg_erro) then 
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

	'	TRATAMENTO P/ OS CASOS EM QUE: USUÁRIO ESTÁ TENTANDO USAR O BOTÃO VOLTAR, OCORREU DUPLO CLIQUE OU USUÁRIO ATUALIZOU A PÁGINA ENQUANTO AINDA ESTAVA PROCESSANDO (DUPLO ACIONAMENTO)
	'	Esse tratamento é feito através do campo proc_comissao_request_guid (t_COMISSAO_INDICADOR_NFSe_N1.proc_comissao_request_guid)
		if proc_comissao_request_guid <> "" then
			s = "UPDATE t_COMISSAO_INDICADOR_NFSe_N1 SET" & _
					" proc_comissao_request_guid = '" & proc_comissao_request_guid & "'" & _
				" WHERE" & _
					" (id = " & id_nsu_N1 & ")"
			cn.Execute s, intRecordsAffected
			if intRecordsAffected <> 1 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Falha ao tentar atualizar o registro do relatório no banco de dados (NSU = " & id_nsu_N1 & ")!!<br />Processamento cancelado!!"
				end if
			end if

		'Tratamento para pedidos (vendas normais)
		s_log_venda_normal = ""
		for i=Lbound(v_venda_normal) to Ubound(v_venda_normal)
			if v_venda_normal(i).ckb_comissao_name <> "" then
				s_id_registro = Trim(v_venda_normal(i).id_registro)
				s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & s_id_registro & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & s_id_registro & " não foi encontrado."
				else
				'   VERIFICA SE USUÁRIO FEZ EDIÇÃO COM RELAÇÃO AO CONTEÚDO ORIGINAL
				'	CHECKBOX ESTAVA MARCADO
					if Trim(v_venda_normal(i).ckb_comissao_status) <> "" then
						'Verifica se o conteúdo original foi alterado por outro usuário
						if CLng(v_venda_normal(i).comissao_status_original) <> rs("comissao_paga") then
							alerta=texto_add_br(alerta)
							alerta=alerta & "O status da comissão do pedido " & Trim("" & rs("pedido")) & " foi alterado por outro usuário (" & Trim("" & rs("comissao_paga_usuario")) & ") durante o processamento deste relatório!!<br />Será necessário refazer a consulta para obter dados atualizados!!"
							blnErroFatal = True
							end if
						
						if Not blnErroFatal then
							s_novo_status = Cstr(COD_COMISSAO_PAGA)
							s_novo_op = "S"
						
							if s_rel_comissao_paga <> "" then s_rel_comissao_paga = s_rel_comissao_paga & ", "
							s_rel_comissao_paga = s_rel_comissao_paga & Trim("" & rs("pedido"))

							if s_log_venda_normal <> "" then s_log_venda_normal = s_log_venda_normal & ", "
							s_log_venda_normal = s_log_venda_normal & s_id_registro & ": " & rs("comissao_paga") & " => " & s_novo_status
							
							rs("comissao_paga")=CLng(s_novo_status)
							rs("comissao_paga_ult_op")=s_novo_op
							rs("comissao_paga_data")=Date
							rs("comissao_paga_usuario")=usuario
							
							qtde_total_reg_update = qtde_total_reg_update + 1
							qtde_venda_normal_update = qtde_venda_normal_update + 1
							rs.Update
							if Err <> 0 then
								alerta=texto_add_br(alerta)
								alerta=alerta & Cstr(Err) & ": " & Err.Description
								end if
								
							if alerta = "" then
								id_cfg_tabela_origem = ID_CFG_TABELA_ORIGEM_T_PEDIDO
								s = "UPDATE tN3Ped SET" & _
										" st_comissao_novo = " & s_novo_status & _
										", st_selecionado = 1" & _
									" FROM t_COMISSAO_INDICADOR_NFSe_N1 tN1" & _
										" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN1.id = tN2.id_comissao_indicador_nfse_n1)" & _
										" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped ON (tN2.id = tN3Ped.id_comissao_indicador_nfse_n2)" & _
									" WHERE" & _
										" (tN1.id = " & id_nsu_N1 & ")" & _
										" AND (tN3Ped.id_cfg_tabela_origem = " & CStr(id_cfg_tabela_origem) & ")" & _
										" AND (tN3Ped.id_registro_tabela_origem = '" & s_id_registro & "')"
								cn.Execute s, intRecordsAffected
								if intRecordsAffected <> 1 then
									alerta=texto_add_br(alerta)
									alerta=alerta & "Falha ao tentar processar o registro em t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO (id_registro_tabela_origem = " & s_id_registro & ", id_cfg_tabela_origem = " & CStr(id_cfg_tabela_origem) & ", t_COMISSAO_INDICADOR_NFSe_N1.id = " & id_nsu_N1 & ")!!<br />Processamento cancelado!!"
									blnErroFatal = True
									end if
								end if 'if alerta = ""
							end if 'if Not blnErroFatal
						end if 'if Trim(v_venda_normal(i).ckb_comissao_status) <> ""
					end if 'if rs.Eof
				if rs.State <> 0 then rs.Close
				end if 'if v_venda_normal(i).ckb_comissao_name <> ""
			
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next
		
		'Tratamento para devoluções
		s_log_devolucao = ""
		for i=Lbound(v_devolucao) to Ubound(v_devolucao)
			if v_devolucao(i).ckb_comissao_name <> "" then
				s_id_registro = Trim(v_devolucao(i).id_registro)
				s = "SELECT * FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (id = '" & s_id_registro & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Registro de item devolvido " & s_id_registro & " não foi encontrado."
				else
				'   VERIFICA SE USUÁRIO FEZ EDIÇÃO COM RELAÇÃO AO CONTEÚDO ORIGINAL
				'	CHECKBOX ESTAVA MARCADO
					if Trim(v_devolucao(i).ckb_comissao_status) <> "" then
						'Verifica se o conteúdo original foi alterado por outro usuário
						if CLng(v_devolucao(i).comissao_status_original) <> rs("comissao_descontada") then
							alerta=texto_add_br(alerta)
							alerta=alerta & "O status da devolução do pedido " & Trim("" & rs("pedido")) & " (produto: " & Trim("" & rs("produto")) & ") foi alterado por outro usuário (" & Trim("" & rs("comissao_descontada_usuario")) & ") durante o processamento deste relatório!!<br />Será necessário refazer a consulta para obter dados atualizados!!"
							blnErroFatal = True
							end if

						if Not blnErroFatal then
							s_novo_status = Cstr(COD_COMISSAO_DESCONTADA)
							s_novo_op = "S"

							if s_rel_devolucao_descontada <> "" then s_rel_devolucao_descontada = s_rel_devolucao_descontada & ", "
							s_rel_devolucao_descontada = s_rel_devolucao_descontada & Trim("" & rs("pedido"))
							
							if s_log_devolucao <> "" then s_log_devolucao = s_log_devolucao & ", "
							s_log_devolucao = s_log_devolucao & s_id_registro & "(" & rs("pedido") & ")" & ": " & rs("comissao_descontada") & " => " & s_novo_status
							
							rs("comissao_descontada")=CLng(s_novo_status)
							rs("comissao_descontada_ult_op")=s_novo_op
							rs("comissao_descontada_data")=Date
							rs("comissao_descontada_usuario")=usuario
							
							qtde_total_reg_update = qtde_total_reg_update + 1
							qtde_devolucao_update = qtde_devolucao_update + 1
							rs.Update
							if Err <> 0 then
								alerta=texto_add_br(alerta)
								alerta=alerta & Cstr(Err) & ": " & Err.Description
								end if

							if alerta = "" then
								id_cfg_tabela_origem = ID_CFG_TABELA_ORIGEM_T_PEDIDO_ITEM_DEVOLVIDO
								s = "UPDATE tN3Ped SET" & _
										" st_comissao_novo = " & s_novo_status & _
										", st_selecionado = 1" & _
									" FROM t_COMISSAO_INDICADOR_NFSe_N1 tN1" & _
										" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN1.id = tN2.id_comissao_indicador_nfse_n1)" & _
										" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped ON (tN2.id = tN3Ped.id_comissao_indicador_nfse_n2)" & _
									" WHERE" & _
										" (tN1.id = " & id_nsu_N1 & ")" & _
										" AND (tN3Ped.id_cfg_tabela_origem = " & CStr(id_cfg_tabela_origem) & ")" & _
										" AND (tN3Ped.id_registro_tabela_origem = '" & s_id_registro & "')"
								cn.Execute s, intRecordsAffected
								if intRecordsAffected <> 1 then
									alerta=texto_add_br(alerta)
									alerta=alerta & "Falha ao tentar processar o registro em t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO (id_registro_tabela_origem = " & s_id_registro & ", id_cfg_tabela_origem = " & CStr(id_cfg_tabela_origem) & ", t_COMISSAO_INDICADOR_NFSe_N1.id = " & id_nsu_N1 & ")!!<br />Processamento cancelado!!"
									blnErroFatal = True
									end if
								end if 'if alerta = ""
							end if 'if Not blnErroFatal
						end if 'if Trim(v_devolucao(i).ckb_comissao_status) <> ""
					end if 'if rs.Eof
				if rs.State <> 0 then rs.Close
				end if 'if v_devolucao(i).ckb_comissao_name <> ""
			
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next
		
		'Tratamento para perdas
		s_log_perda = ""
		for i=Lbound(v_perda) to Ubound(v_perda)
			if v_perda(i).ckb_comissao_name <> "" then
				s_id_registro = Trim(v_perda(i).id_registro)
				s = "SELECT * FROM t_PEDIDO_PERDA WHERE (id = '" & s_id_registro & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Registro de perda " & s_id_registro & " não foi encontrado."
				else
				'   VERIFICA SE USUÁRIO FEZ EDIÇÃO COM RELAÇÃO AO CONTEÚDO ORIGINAL
				'	CHECKBOX ESTAVA MARCADO
					if Trim(v_perda(i).ckb_comissao_status) <> "" then
						'Verifica se o conteúdo original foi alterado por outro usuário
						if CLng(v_perda(i).comissao_status_original) <> rs("comissao_descontada") then
							alerta=texto_add_br(alerta)
							alerta=alerta & "O status do valor de perda do pedido " & Trim("" & rs("pedido")) & " foi alterado por outro usuário (" & Trim("" & rs("comissao_descontada_usuario")) & ") durante o processamento deste relatório!!<br />Será necessário refazer a consulta para obter dados atualizados!!"
							blnErroFatal = True
							end if

						if Not blnErroFatal then
							s_novo_status = Cstr(COD_COMISSAO_DESCONTADA)
							s_novo_op = "S"

							if s_rel_perda_descontada <> "" then s_rel_perda_descontada = s_rel_perda_descontada & ", "
							s_rel_perda_descontada = s_rel_perda_descontada & Trim("" & rs("pedido"))
					
							if s_log_perda <> "" then s_log_perda = s_log_perda & ", "
							s_log_perda = s_log_perda & s_id_registro & "(" & rs("pedido") & ")" & ": " & rs("comissao_descontada") & " => " & s_novo_status
							
							rs("comissao_descontada")=CLng(s_novo_status)
							rs("comissao_descontada_ult_op")=s_novo_op
							rs("comissao_descontada_data")=Date
							rs("comissao_descontada_usuario")=usuario
							
							qtde_total_reg_update = qtde_total_reg_update + 1
							qtde_perda_update = qtde_perda_update + 1
							rs.Update
							if Err <> 0 then
								alerta=texto_add_br(alerta)
								alerta=alerta & Cstr(Err) & ": " & Err.Description
								end if

							if alerta = "" then
								id_cfg_tabela_origem = ID_CFG_TABELA_ORIGEM_T_PEDIDO_PERDA
								s = "UPDATE tN3Ped SET" & _
										" st_comissao_novo = " & s_novo_status & _
										", st_selecionado = 1" & _
									" FROM t_COMISSAO_INDICADOR_NFSe_N1 tN1" & _
										" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN1.id = tN2.id_comissao_indicador_nfse_n1)" & _
										" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped ON (tN2.id = tN3Ped.id_comissao_indicador_nfse_n2)" & _
									" WHERE" & _
										" (tN1.id = " & id_nsu_N1 & ")" & _
										" AND (tN3Ped.id_cfg_tabela_origem = " & CStr(id_cfg_tabela_origem) & ")" & _
										" AND (tN3Ped.id_registro_tabela_origem = '" & s_id_registro & "')"
								cn.Execute s, intRecordsAffected
								if intRecordsAffected <> 1 then
									alerta=texto_add_br(alerta)
									alerta=alerta & "Falha ao tentar processar o registro em t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO (id_registro_tabela_origem = " & s_id_registro & ", id_cfg_tabela_origem = " & CStr(id_cfg_tabela_origem) & ", t_COMISSAO_INDICADOR_NFSe_N1.id = " & id_nsu_N1 & ")!!<br />Processamento cancelado!!"
									blnErroFatal = True
									end if
								end if 'if alerta = ""
							end if 'if Not blnErroFatal
						end if 'if Trim(v_perda(i).ckb_comissao_status) <> ""
					end if 'if rs.Eof
				if rs.State <> 0 then rs.Close
				end if 'if v_perda(i).ckb_comissao_name <> ""
			
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next
		
		'Conferência dos valores
		if alerta = "" then
			s = "SELECT" & _
					" SUM(vl_preco_venda) AS vl_total_geral_selecionado_preco_venda" & _
					", SUM(vl_preco_NF) AS vl_total_geral_selecionado_preco_NF" & _
					", SUM(vl_RT) AS vl_total_geral_selecionado_RT" & _
					", SUM(vl_RA_bruto) AS vl_total_geral_selecionado_RA_bruto" & _
					", SUM(vl_RA_liquido) AS vl_total_geral_selecionado_RA_liq" & _
					", SUM(vl_RA_dif) AS vl_total_geral_selecionado_RA_dif" & _
				" FROM t_COMISSAO_INDICADOR_NFSe_N1 tN1" & _
					" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN1.id = tN2.id_comissao_indicador_nfse_n1)" & _
					" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped ON (tN2.id = tN3Ped.id_comissao_indicador_nfse_n2)" & _
				" WHERE" & _
					" (tN1.id = " & id_nsu_N1 & ")" & _
					" AND (tN3Ped.st_selecionado = 1)"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Falha ao consultar valores processados no banco de dados (NSU = " & id_nsu_N1 & ")"
			else
				vl_total_geral_selecionado_preco_venda_bd = rs("vl_total_geral_selecionado_preco_venda")
				vl_total_geral_selecionado_preco_NF_bd = rs("vl_total_geral_selecionado_preco_NF")
				vl_total_geral_selecionado_RT_bd = rs("vl_total_geral_selecionado_RT")
				vl_total_geral_selecionado_RA_bruto_bd = rs("vl_total_geral_selecionado_RA_bruto")
				vl_total_geral_selecionado_RA_liq_bd = rs("vl_total_geral_selecionado_RA_liq")
				vl_total_geral_selecionado_RA_dif_bd = rs("vl_total_geral_selecionado_RA_dif")
				end if
			end if 'if alerta = ""

		if alerta = "" then
			if converte_numero(c_total_geral_selecionado_preco_venda) <> vl_total_geral_selecionado_preco_venda_bd then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Divergência na conferência do valor total (preço de venda): " & formata_moeda(converte_numero(c_total_geral_selecionado_preco_venda)) & " e " & formata_moeda(vl_total_geral_selecionado_preco_venda_bd) & " (NSU = " & id_nsu_N1 & ")"
				end if
			end if 'if alerta = ""

		if alerta = "" then
			if converte_numero(c_total_geral_selecionado_preco_NF) <> vl_total_geral_selecionado_preco_NF_bd then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Divergência na conferência do valor total (preço NF): " & formata_moeda(converte_numero(c_total_geral_selecionado_preco_NF)) & " e " & formata_moeda(vl_total_geral_selecionado_preco_NF_bd) & " (NSU = " & id_nsu_N1 & ")"
				end if
			end if 'if alerta = ""

		if alerta = "" then
			if converte_numero(c_total_geral_selecionado_RT) <> vl_total_geral_selecionado_RT_bd then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Divergência na conferência do valor de RT: " & formata_moeda(converte_numero(c_total_geral_selecionado_RT)) & " e " & formata_moeda(vl_total_geral_selecionado_RT_bd) & " (NSU = " & id_nsu_N1 & ")"
				end if
			end if 'if alerta = ""

		if alerta = "" then
			if converte_numero(c_total_geral_selecionado_RA_bruto) <> vl_total_geral_selecionado_RA_bruto_bd then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Divergência na conferência do valor de RA bruto: " & formata_moeda(converte_numero(c_total_geral_selecionado_RA_bruto)) & " e " & formata_moeda(vl_total_geral_selecionado_RA_bruto_bd) & " (NSU = " & id_nsu_N1 & ")"
				end if
			end if 'if alerta = ""

		if alerta = "" then
			if converte_numero(c_total_geral_selecionado_RA_liq) <> vl_total_geral_selecionado_RA_liq_bd then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Divergência na conferência do valor de RA líquido: " & formata_moeda(converte_numero(c_total_geral_selecionado_RA_liq)) & " e " & formata_moeda(vl_total_geral_selecionado_RA_liq_bd) & " (NSU = " & id_nsu_N1 & ")"
				end if
			end if 'if alerta = ""

		if alerta = "" then
			s = "UPDATE t_COMISSAO_INDICADOR_NFSe_N1 SET" & _
					" status = 1" & _
					", dt_hr_ult_atualizacao = getdate()" & _
					", usuario_ult_atualizacao = '" & QuotedStr(usuario) & "'" & _
					", proc_comissao_status = 1" & _
					", proc_comissao_data = Convert(datetime, Convert(varchar(10),getdate(), 121), 121)" & _
					", proc_comissao_data_hora = getdate()" & _
					", proc_comissao_usuario = '" & QuotedStr(usuario) & "'" & _
					", vl_total_geral_selecionado_preco_venda = " & bd_formata_numero(vl_total_geral_selecionado_preco_venda_bd) & _
					", vl_total_geral_selecionado_preco_NF = " & bd_formata_numero(vl_total_geral_selecionado_preco_NF_bd) & _
					", vl_total_geral_selecionado_RT = " & bd_formata_numero(vl_total_geral_selecionado_RT_bd) & _
					", vl_total_geral_selecionado_RA_bruto = " & bd_formata_numero(vl_total_geral_selecionado_RA_bruto_bd) & _
					", vl_total_geral_selecionado_RA_liquido = " & bd_formata_numero(vl_total_geral_selecionado_RA_liq_bd) & _
					", vl_total_geral_selecionado_RA_dif = " & bd_formata_numero(vl_total_geral_selecionado_RA_dif_bd) & _
				" WHERE" & _
					" (id = " & id_nsu_N1 & ")" & _
					" AND (status = 0)" & _
					" AND (proc_comissao_status = 0)"
			cn.Execute s, intRecordsAffected
			if intRecordsAffected <> 1 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Falha ao tentar atualizar o status do relatório no banco de dados (NSU = " & id_nsu_N1 & ")!!<br />Processamento cancelado!!"
				end if
			end if 'if alerta = ""

		if alerta = "" then
			if s_log_venda_normal <> "" then
				s_log_venda_normal = "Venda Normal (t_PEDIDO.comissao_paga): " & s_log_venda_normal
				if s_log <> "" then s_log = s_log & "; " & chr(13)
				s_log = s_log & s_log_venda_normal
				end if
			if s_log_devolucao <> "" then
				s_log_devolucao = "Devolução (t_PEDIDO_ITEM_DEVOLVIDO.comissao_descontada): " & s_log_devolucao
				if s_log <> "" then s_log = s_log & "; " & chr(13)
				s_log = s_log & s_log_devolucao
				end if
			if s_log_perda <> "" then
				s_log_perda = "Perda (t_PEDIDO_PERDA.comissao_descontada): " & s_log_perda
				if s_log <> "" then s_log = s_log & "; " & chr(13)
				s_log = s_log & s_log_perda
				end if
			if s_log <> "" then
				s_log = "Edição do status de comissão paga/não-paga através do 'Relatório de Pedidos Indicadores (via NFS-e) [t_COMISSAO_INDICADOR_NFSe_N1.id = " & id_nsu_N1 & "]': CNPJ NFS-e: " & cnpj_cpf_formata(c_cnpj_nfse) & "; Total registros atualizados = " & Cstr(qtde_total_reg_update) & " (Venda Normal: " & Cstr(qtde_venda_normal_update) & ", Devolução: " & Cstr(qtde_devolucao_update) & ", Perda: " & Cstr(qtde_perda_update) & ")" & chr(13) & s_log
				grava_log usuario, "", "", "", OP_LOG_PEDIDO_ALTERACAO, s_log
				end if
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then
				' NOP: Response.Redirect("resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
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

	function fBotaoMagico(f) {
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
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
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
</form>

</center>
</body>

<% else %>

<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';bVOLTAR.focus();" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="f" name="f" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="ckb_id_indicador" id="ckb_id_indicador" value="<%=ckb_id_indicador%>" />
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>">
<input type="hidden" name="id_nsu_N1" id="id_nsu_N1" value="<%=id_nsu_N1%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="820" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
	<tr>
		<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (via NFS-e)</span></td>
	</tr>
</table>
<br />
<br />

<%
	If Not cria_recordset_otimista(tN1, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tN3Ped, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	s = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N1 WHERE (id = " & id_nsu_N1 & ")"
	tN1.Open s, cn
	if tN1.Eof then
		blnErroFatal = True
%>
<div class='MtAlerta' style='width:800px;font-weight:bold;' align='center'>
<p style='margin:5px 2px 5px 2px;'>Falha ao tentar localizar dados do relatório (NSU = <%=id_nsu_N1%>)</p></div>
<br />
<br />
<%
		end if
%>

<% if Not blnErroFatal then %>

<% if blnRelJaProcessado then %>
<!-- ************   MENSAGEM DE ALERTA INFORMANDO QUE O RELATÓRIO JÁ FOI PROCESSADO ANTERIORMENTE ************ -->
<div class='MtAlerta' style='width:800px;font-weight:bold;' align='center'>
<p style='margin:5px 2px 5px 2px;'>Este relatório já foi processado em <%=formata_data_hora_sem_seg(tN1("proc_comissao_data_hora"))%> por <%=Trim("" & tN1("proc_comissao_usuario"))%></p></div>
<br />
<br />
<% end if 'if blnRelJaProcessado %>


<!-- ************   MENSAGEM  ************ -->
<table class="Qx" style="width:800px;" cellpadding="1" cellspacing="0">
	<tr>
		<td class="MT TdLabel" align="right"><span class="Cd">NSU do Relatório:</span></td>
		<td class="MTBD TdInfo" align="left"><span class="C"><%=CStr(id_nsu_N1)%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" align="right"><span class="Cd">CNPJ NFS-e:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=cnpj_cpf_formata(c_cnpj_nfse)%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Comissão Processada em:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=formata_data_e_talvez_hora_hhmm(tN1("proc_comissao_data_hora"))%> por <%=Trim("" & tN1("proc_comissao_usuario"))%></span></td>
	</tr>
	<% if tN1("proc_fluxo_caixa_status") = 1 then %>
	<tr>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Lançamentos Processados em:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=formata_data_e_talvez_hora_hhmm(tN1("proc_fluxo_caixa_data_hora"))%> por <%=Trim("" & tN1("proc_fluxo_caixa_usuario"))%></span></td>
	</tr>
	<% end if %>
	<%
		if qtde_total_reg_update = 0 then
			s = "SELECT" & _
						" Count(*) AS QtdeSelecionada" & _
					" FROM t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN3Ped.id_comissao_indicador_nfse_n2 = tN2.id)" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N1 tN1 ON (tN2.id_comissao_indicador_nfse_n1 = tN1.id)" & _
					" WHERE" & _
						" (tN1.id = " & id_nsu_N1 & ")" & _
						" AND (tN3Ped.st_selecionado = 1)"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Not rs.Eof then
				qtde_total_reg_update = rs("QtdeSelecionada")
				end if
			end if 'if qtde_total_reg_update = 0
	%>
	<% if qtde_total_reg_update = 0 then %>
	<tr>
		<td colspan="2" align="center"><span class="C">Nenhuma alteração foi realizada para gravar</span></td>
	</tr>
	<% else %>
	<tr>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Dados gravados com sucesso:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=Cstr(qtde_total_reg_update)%> registro(s) atualizado(s)</span></td>
	</tr>
	<tr>
		<%
			if s_rel_comissao_paga = "" then
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
				end if 'if s_rel_comissao_paga = ""

		s_aux = s_rel_comissao_paga
		if s_aux = "" then s_aux = "(nenhum pedido)"
		%>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Comissão Paga:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<%
			if s_rel_devolucao_descontada = "" then
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
				end if 'if s_rel_devolucao_descontada = ""
		%>
		<% s_aux = s_rel_devolucao_descontada
			if s_aux = "" then s_aux = "(nenhum pedido)"
		%>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Devolução Descontada:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<%
			if s_rel_perda_descontada = "" then
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
				end if 'if s_rel_perda_descontada = ""

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
		<td class="MDBE TdLabel" align="right"><span class="Cd">Comissão (RT):</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<% s_aux = formata_moeda(tN1("vl_total_geral_selecionado_RA_liquido"))
			if s_aux = "" then s_aux = "?"
		%>
		<td class="MDBE TdLabel" align="right"><span class="Cd">RA Líquido:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<% s_aux = formata_moeda(tN1("vl_total_geral_selecionado_RT") + tN1("vl_total_geral_selecionado_RA_liquido"))
			if s_aux = "" then s_aux = "?"
		%>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Total Comissão (RT + RA Líquido):</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<% end if %>
</table>
<br />



<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="820" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<!-- ************   LINKS: PÁGINA INICIAL / ENCERRA SESSÃO   ************ -->
<table class="notPrint" width="820" cellpadding="0" cellspacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="820" cellspacing="0">
<% if tN1("proc_fluxo_caixa_status") = 0 then %>
<tr>
	<td align="left"><div name="dVOLTAR" id="dVOLTAR"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fRetornar(f)" title="Retornar para o início do relatório">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="right"><div name="dBOTAOMAGICO" id="dBOTAOMAGICO"><a name="bBOTAOMAGICO" id="bBOTAOMAGICO" href="javascript:fBotaoMagico(f)" title="Botão mágico">
		<img src="../botao/magico.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% else %>
<tr>
	<td align="center"><div name="dVOLTAR" id="dVOLTAR"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fRetornar(f)" title="Retornar para o início do relatório">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% end if 'if tN1("proc_fluxo_caixa_status") = 0 %>
<% end if 'if Not blnErroFatal %>
</table>
</form>

</center>
</body>

<%
	if tN1.State <> 0 then tN1.Close
	set tN1 = nothing
	
	if tN3Ped.State <> 0 then tN3Ped.Close
	set tN3Ped = nothing
%>

<% end if 'if alerta <> "" then - else %>

</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>