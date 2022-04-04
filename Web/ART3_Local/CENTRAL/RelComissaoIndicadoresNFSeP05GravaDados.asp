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
	
	const VENDA_NORMAL = "VENDA_NORMAL"
	const DEVOLUCAO = "DEVOLUCAO"
	const PERDA = "PERDA"

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
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim s_log, s_log_venda_normal, s_log_devolucao, s_log_perda, s_novo_status, s_novo_op
	
'	FILTROS
	dim c_cnpj_nfse
	dim ckb_id_indicador
	dim ckb_st_entrega_entregue, c_dt_entregue_inicio, c_dt_entregue_termino
	dim ckb_comissao_paga_sim, ckb_comissao_paga_nao
	dim ckb_st_pagto_pago, ckb_st_pagto_nao_pago, ckb_st_pagto_pago_parcial
	dim rb_visao

	c_cnpj_nfse = retorna_so_digitos(Request.Form("c_cnpj_nfse"))
	ckb_id_indicador = Trim(Request.Form("ckb_id_indicador"))
	ckb_st_entrega_entregue = Trim(Request.Form("ckb_st_entrega_entregue"))
	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))

	ckb_comissao_paga_sim = Trim(Request.Form("ckb_comissao_paga_sim"))
	ckb_comissao_paga_nao = Trim(Request.Form("ckb_comissao_paga_nao"))

	ckb_st_pagto_pago = Trim(Request.Form("ckb_st_pagto_pago"))
	ckb_st_pagto_nao_pago = Trim(Request.Form("ckb_st_pagto_nao_pago"))
	ckb_st_pagto_pago_parcial = Trim(Request.Form("ckb_st_pagto_pago_parcial"))
	rb_visao = Trim(Request.Form("rb_visao"))

'	OBTÉM DADOS DO FORMULÁRIO
	dim c_lista_completa_venda_normal, c_lista_completa_devolucao, c_lista_completa_perda
	dim c_lista_ja_marcado_venda_normal, c_lista_ja_marcado_devolucao, c_lista_ja_marcado_perda
	c_lista_completa_venda_normal = Trim(Request.Form("c_lista_completa_venda_normal"))
	c_lista_completa_devolucao = Trim(Request.Form("c_lista_completa_devolucao"))
	c_lista_completa_perda = Trim(Request.Form("c_lista_completa_perda"))
	c_lista_ja_marcado_venda_normal = Trim(Request.Form("c_lista_ja_marcado_venda_normal"))
	c_lista_ja_marcado_devolucao = Trim(Request.Form("c_lista_ja_marcado_devolucao"))
	c_lista_ja_marcado_perda = Trim(Request.Form("c_lista_ja_marcado_perda"))

	dim i, s_name, s_name_valor_oginal, s_id_registro, qtde_total_reg_update, qtde_venda_normal_update, qtde_devolucao_update, qtde_perda_update, blnEditou, blnErroFatal
	dim v_lista_completa_venda_normal, v_lista_completa_devolucao, v_lista_completa_perda
	v_lista_completa_venda_normal = Split(c_lista_completa_venda_normal, ";", -1)
	v_lista_completa_devolucao = Split(c_lista_completa_devolucao, ";", -1)
	v_lista_completa_perda = Split(c_lista_completa_perda, ";", -1)
	
	dim v_venda_normal, v_devolucao, v_perda
	
	redim v_venda_normal(0)
	set v_venda_normal(UBound(v_venda_normal)) = New cl_QUATRO_COLUNAS
	v_venda_normal(UBound(v_venda_normal)).c1 = ""
	for i=LBound(v_lista_completa_venda_normal) to Ubound(v_lista_completa_venda_normal)
		if Trim(v_lista_completa_venda_normal(i)) <> "" then
			s_id_registro = Trim(v_lista_completa_venda_normal(i))
			s_name = "ckb_comissao_paga_" & VENDA_NORMAL & "_" & s_id_registro
			s_name_valor_oginal = s_name & "_original"
			if v_venda_normal(Ubound(v_venda_normal)).c1 <> "" then
				redim preserve v_venda_normal(Ubound(v_venda_normal)+1)
				set v_venda_normal(UBound(v_venda_normal)) = New cl_QUATRO_COLUNAS
				end if
			with v_venda_normal(Ubound(v_venda_normal))
				.c1 = s_name
				.c2 = s_id_registro
				.c3 = Trim(Request.Form(s_name))
				'Recupera o valor original do status
				.c4 = Trim(Request.Form(s_name_valor_oginal))
				end with
			end if
		next
	
	redim v_devolucao(0)
	set v_devolucao(UBound(v_devolucao)) = New cl_QUATRO_COLUNAS
	v_devolucao(UBound(v_devolucao)).c1 = ""
	for i=LBound(v_lista_completa_devolucao) to Ubound(v_lista_completa_devolucao)
		if Trim(v_lista_completa_devolucao(i)) <> "" then
			s_id_registro = Trim(v_lista_completa_devolucao(i))
			s_name = "ckb_comissao_paga_" & DEVOLUCAO & "_" & s_id_registro
			s_name_valor_oginal = s_name & "_original"
			if v_devolucao(Ubound(v_devolucao)).c1 <> "" then
				redim preserve v_devolucao(Ubound(v_devolucao)+1)
				set v_devolucao(UBound(v_devolucao)) = New cl_QUATRO_COLUNAS
				end if
			with v_devolucao(Ubound(v_devolucao))
				.c1 = s_name
				.c2 = s_id_registro
				.c3 = Trim(Request.Form(s_name))
				'Recupera o valor original do status
				.c4 = Trim(Request.Form(s_name_valor_oginal))
				end with
			end if
		next
	
	redim v_perda(0)
	set v_perda(UBound(v_perda)) = New cl_QUATRO_COLUNAS
	v_perda(UBound(v_perda)).c1 = ""
	for i=LBound(v_lista_completa_perda) to Ubound(v_lista_completa_perda)
		if Trim(v_lista_completa_perda(i)) <> "" then
			s_id_registro = Trim(v_lista_completa_perda(i))
			s_name = "ckb_comissao_paga_" & PERDA & "_" & s_id_registro
			s_name_valor_oginal = s_name & "_original"
			if v_perda(Ubound(v_perda)).c1 <> "" then
				redim preserve v_perda(Ubound(v_perda)+1)
				set v_perda(UBound(v_perda)) = New cl_QUATRO_COLUNAS
				end if
			with v_perda(Ubound(v_perda))
				.c1 = s_name
				.c2 = s_id_registro
				.c3 = Trim(Request.Form(s_name))
				'Recupera o valor original do status
				.c4 = Trim(Request.Form(s_name_valor_oginal))
				end with
			end if
		next
	
	dim s_rel_comissao_paga, s_rel_comissao_nao_paga, s_rel_devolucao_descontada, s_rel_devolucao_nao_descontada, s_rel_perda_descontada, s_rel_perda_nao_descontada
	s_rel_comissao_paga = ""
	s_rel_comissao_nao_paga = ""
	s_rel_devolucao_descontada = ""
	s_rel_devolucao_nao_descontada = ""
	s_rel_perda_descontada = ""
	s_rel_perda_nao_descontada = ""
	
	dim alerta
	alerta=""

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	blnErroFatal = False
	s_log = ""
	qtde_total_reg_update = 0
	qtde_venda_normal_update = 0
	qtde_devolucao_update = 0
	qtde_perda_update = 0
	
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		If Not cria_recordset_pessimista(rs, msg_erro) then 
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

		'Tratamento para pedidos (vendas normais)
		s_log_venda_normal = ""
		for i=Lbound(v_venda_normal) to Ubound(v_venda_normal)
			blnEditou = False
			if v_venda_normal(i).c1 <> "" then
				s_id_registro = Trim(v_venda_normal(i).c2)
				s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & s_id_registro & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & s_id_registro & " não foi encontrado."
				else
				'   VERIFICA SE USUÁRIO FEZ EDIÇÃO COM RELAÇÃO AO CONTEÚDO ORIGINAL
				'	CHECKBOX ESTAVA MARCADO
					if Trim(v_venda_normal(i).c3) <> "" then
						s_novo_status = Cstr(COD_COMISSAO_PAGA)
				'	CHECKBOX ESTAVA DESMARCADO
					else
						s_novo_status = Cstr(COD_COMISSAO_NAO_PAGA)
						end if
					
					if Trim("" & s_novo_status) <> Trim("" & v_venda_normal(i).c4) then blnEditou = True

					if blnEditou then
						'Verifica se o conteúdo original foi alterado por outro usuário
						if CLng(v_venda_normal(i).c4) <> rs("comissao_paga") then
							alerta=texto_add_br(alerta)
							alerta=alerta & "O status da comissão do pedido " & Trim("" & rs("pedido")) & " foi alterado por outro usuário (" & Trim("" & rs("comissao_paga_usuario")) & ") durante o processamento deste relatório!!<br />Será necessário refazer a consulta para obter dados atualizados!!"
							blnErroFatal = True
							end if
						
						if Not blnErroFatal then
						'	CHECKBOX ESTAVA MARCADO
							if Trim(v_venda_normal(i).c3) <> "" then
								s_novo_op = "S"
								if s_rel_comissao_paga <> "" then s_rel_comissao_paga = s_rel_comissao_paga & ", "
								s_rel_comissao_paga = s_rel_comissao_paga & Trim("" & rs("pedido"))
						'	CHECKBOX ESTAVA DESMARCADO
							else
								s_novo_op = "N"
								if s_rel_comissao_nao_paga <> "" then s_rel_comissao_nao_paga = s_rel_comissao_nao_paga & ", "
								s_rel_comissao_nao_paga = s_rel_comissao_nao_paga & Trim("" & rs("pedido"))
								end if
					
							if rs("comissao_paga") <> CLng(s_novo_status) then
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
								end if 'if rs("comissao_paga") <> CLng(s_novo_status)
							end if 'if Not blnErroFatal
						end if 'if blnEditou
					end if 'if rs.Eof
				if rs.State <> 0 then rs.Close
				end if 'if v_venda_normal(i).c1 <> ""
			
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next
		
		'Tratamento para devoluções
		s_log_devolucao = ""
		for i=Lbound(v_devolucao) to Ubound(v_devolucao)
			blnEditou = False
			if v_devolucao(i).c1 <> "" then
				s_id_registro = Trim(v_devolucao(i).c2)
				s = "SELECT * FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (id = '" & s_id_registro & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Registro de item devolvido " & s_id_registro & " não foi encontrado."
				else
				'   VERIFICA SE USUÁRIO FEZ EDIÇÃO COM RELAÇÃO AO CONTEÚDO ORIGINAL
				'	CHECKBOX ESTAVA MARCADO
					if Trim(v_devolucao(i).c3) <> "" then
						s_novo_status = Cstr(COD_COMISSAO_DESCONTADA)
				'	CHECKBOX ESTAVA DESMARCADO
					else
						s_novo_status = Cstr(COD_COMISSAO_NAO_DESCONTADA)
						end if

					if Trim("" & s_novo_status) <> Trim("" & v_devolucao(i).c4) then blnEditou = True

					if blnEditou then
						'Verifica se o conteúdo original foi alterado por outro usuário
						if CLng(v_devolucao(i).c4) <> rs("comissao_descontada") then
							alerta=texto_add_br(alerta)
							alerta=alerta & "O status da devolução do pedido " & Trim("" & rs("pedido")) & " (produto: " & Trim("" & rs("produto")) & ") foi alterado por outro usuário (" & Trim("" & rs("comissao_descontada_usuario")) & ") durante o processamento deste relatório!!<br />Será necessário refazer a consulta para obter dados atualizados!!"
							blnErroFatal = True
							end if

						if Not blnErroFatal then
						'	CHECKBOX ESTAVA MARCADO
							if Trim(v_devolucao(i).c3) <> "" then
								s_novo_op = "S"
								if s_rel_devolucao_descontada <> "" then s_rel_devolucao_descontada = s_rel_devolucao_descontada & ", "
								s_rel_devolucao_descontada = s_rel_devolucao_descontada & Trim("" & rs("pedido"))
						'	CHECKBOX ESTAVA DESMARCADO
							else
								s_novo_op = "N"
								if s_rel_devolucao_nao_descontada <> "" then s_rel_devolucao_nao_descontada = s_rel_devolucao_nao_descontada & ", "
								s_rel_devolucao_nao_descontada = s_rel_devolucao_nao_descontada & Trim("" & rs("pedido"))
								end if
					
							if rs("comissao_descontada") <> CLng(s_novo_status) then
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
								end if 'if rs("comissao_descontada") <> CLng(s_novo_status)
							end if 'if Not blnErroFatal
						end if 'if blnEditou
					end if 'if rs.Eof
				if rs.State <> 0 then rs.Close
				end if 'if v_devolucao(i).c1 <> ""
			
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next
		
		'Tratamento para perdas
		s_log_perda = ""
		for i=Lbound(v_perda) to Ubound(v_perda)
			blnEditou = False
			if v_perda(i).c1 <> "" then
				s_id_registro = Trim(v_perda(i).c2)
				s = "SELECT * FROM t_PEDIDO_PERDA WHERE (id = '" & s_id_registro & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Registro de perda " & s_id_registro & " não foi encontrado."
				else
				'   VERIFICA SE USUÁRIO FEZ EDIÇÃO COM RELAÇÃO AO CONTEÚDO ORIGINAL
				'	CHECKBOX ESTAVA MARCADO
					if Trim(v_perda(i).c3) <> "" then
						s_novo_status = Cstr(COD_COMISSAO_DESCONTADA)
				'	CHECKBOX ESTAVA DESMARCADO
					else
						s_novo_status = Cstr(COD_COMISSAO_NAO_DESCONTADA)
						end if

					if Trim("" & s_novo_status) <> Trim("" & v_perda(i).c4) then blnEditou = True

					if blnEditou then
						'Verifica se o conteúdo original foi alterado por outro usuário
						if CLng(v_perda(i).c4) <> rs("comissao_descontada") then
							alerta=texto_add_br(alerta)
							alerta=alerta & "O status do valor de perda do pedido " & Trim("" & rs("pedido")) & " foi alterado por outro usuário (" & Trim("" & rs("comissao_descontada_usuario")) & ") durante o processamento deste relatório!!<br />Será necessário refazer a consulta para obter dados atualizados!!"
							blnErroFatal = True
							end if

						if Not blnErroFatal then
						'	CHECKBOX ESTAVA MARCADO
							if Trim(v_perda(i).c3) <> "" then
								s_novo_op = "S"
								if s_rel_perda_descontada <> "" then s_rel_perda_descontada = s_rel_perda_descontada & ", "
								s_rel_perda_descontada = s_rel_perda_descontada & Trim("" & rs("pedido"))
						'	CHECKBOX ESTAVA DESMARCADO
							else
								s_novo_op = "N"
								if s_rel_perda_nao_descontada <> "" then s_rel_perda_nao_descontada = s_rel_perda_nao_descontada & ", "
								s_rel_perda_nao_descontada = s_rel_perda_nao_descontada & Trim("" & rs("pedido"))
								end if
					
							if rs("comissao_descontada") <> CLng(s_novo_status) then
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
								end if 'if rs("comissao_descontada") <> CLng(s_novo_status)
							end if 'if Not blnErroFatal
						end if 'if blnEditou
					end if 'if rs.Eof
				if rs.State <> 0 then rs.Close
				end if 'if v_perda(i).c1 <> ""
			
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next
		
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
				s_log = "Edição do status de comissão paga/não-paga através do 'Relatório de Pedidos Indicadores (via NFS-e)': CNPJ NFS-e: " & cnpj_cpf_formata(c_cnpj_nfse) & "; Total registros atualizados = " & Cstr(qtde_total_reg_update) & " (Venda Normal: " & Cstr(qtde_venda_normal_update) & ", Devolução: " & Cstr(qtde_devolucao_update) & ", Perda: " & Cstr(qtde_perda_update) & ")" & chr(13) & s_log
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
<input type="hidden" name="ckb_st_entrega_entregue" id="ckb_st_entrega_entregue" value="<%=ckb_st_entrega_entregue%>">
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="<%=c_dt_entregue_inicio%>">
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>">
<input type="hidden" name="ckb_comissao_paga_sim" id="ckb_comissao_paga_sim" value="<%=ckb_comissao_paga_sim%>">
<input type="hidden" name="ckb_comissao_paga_nao" id="ckb_comissao_paga_nao" value="<%=ckb_comissao_paga_nao%>">
<input type="hidden" name="ckb_st_pagto_pago" id="ckb_st_pagto_pago" value="<%=ckb_st_pagto_pago%>">
<input type="hidden" name="ckb_st_pagto_nao_pago" id="ckb_st_pagto_nao_pago" value="<%=ckb_st_pagto_nao_pago%>">
<input type="hidden" name="ckb_st_pagto_pago_parcial" id="ckb_st_pagto_pago_parcial" value="<%=ckb_st_pagto_pago_parcial%>">

<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
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
<input type="hidden" name="ckb_st_entrega_entregue" id="ckb_st_entrega_entregue" value="<%=ckb_st_entrega_entregue%>">
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="<%=c_dt_entregue_inicio%>">
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>">
<input type="hidden" name="ckb_comissao_paga_sim" id="ckb_comissao_paga_sim" value="<%=ckb_comissao_paga_sim%>">
<input type="hidden" name="ckb_comissao_paga_nao" id="ckb_comissao_paga_nao" value="<%=ckb_comissao_paga_nao%>">
<input type="hidden" name="ckb_st_pagto_pago" id="ckb_st_pagto_pago" value="<%=ckb_st_pagto_pago%>">
<input type="hidden" name="ckb_st_pagto_nao_pago" id="ckb_st_pagto_nao_pago" value="<%=ckb_st_pagto_nao_pago%>">
<input type="hidden" name="ckb_st_pagto_pago_parcial" id="ckb_st_pagto_pago_parcial" value="<%=ckb_st_pagto_pago_parcial%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (via NFS-e)</span></td>
</tr>
</table>
<br>
<br>

<!-- ************   MENSAGEM  ************ -->
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center">
	<% if qtde_total_reg_update = 0 then %>
	<span style='margin:5px 2px 5px 2px;'>Nenhuma alteração foi realizada para gravar</span>
	<% else %>
	<span style='margin:5px 2px 5px 2px;'>CNPJ NFS-e: <%=cnpj_cpf_formata(c_cnpj_nfse)%></span>
	<br /><br />
	<span style='margin:5px 2px 5px 2px;'>Processado em <%=formata_data_e_talvez_hora_hhmm(Now)%> por <%=usuario%></span>
	<br /><br />
	<span style='margin:5px 2px 5px 2px;'>Dados gravados com sucesso: <%=Cstr(qtde_total_reg_update)%> registro(s) atualizado(s)</span>
	<br /><br />
	<span style='margin:5px 2px 5px 2px;'>Comissão Paga:</span>
	<br />
		<% s_aux = s_rel_comissao_paga
			if s_aux = "" then s_aux = "(nenhum pedido)"
		%>
	<span style='margin:5px 2px 5px 2px;'><%=s_aux%></span>
    <% if s_rel_comissao_nao_paga <> "" then %>
    <br /><br />
    <span style='margin:5px 2px 5px 2px;'>Comissão <span style='color:red;'>NÃO</span> Paga:</span>
    <br />
    <span style='margin:5px 2px 5px 2px;'><%=s_rel_comissao_nao_paga%></span>
    <% end if %>
	<br /><br />
	<span style='margin:5px 2px 5px 2px;'>Devolução Descontada:</span>
	<br />
		<% s_aux = s_rel_devolucao_descontada
			if s_aux = "" then s_aux = "(nenhum pedido)"
		%>
	<span style='margin:5px 2px 5px 2px;'><%=s_aux%></span>
	<% if s_rel_devolucao_nao_descontada <> "" then %>
    <br /><br />
    <span style='margin:5px 2px 5px 2px;'>Devolução <span style='color:red;'>NÃO</span> Descontada:</span>
    <br />
    <span style='margin:5px 2px 5px 2px;'><%=s_rel_devolucao_nao_descontada%></span>
    <% end if %>
    <br /><br />
	<span style='margin:5px 2px 5px 2px;'>Perda Descontada:</span>
	<br />
		<% s_aux = s_rel_perda_descontada
			if s_aux = "" then s_aux = "(nenhum pedido)"
		%>
	<span style='margin:5px 2px 5px 2px;'><%=s_aux%></span>
    <% if s_rel_perda_nao_descontada <> "" then %>
    <br /><br />
    <span style='margin:5px 2px 5px 2px;'>Perda <span style='color:red;'>NÃO</span> Descontada:</span>
    <br />
    <span style='margin:5px 2px 5px 2px;'><%=s_rel_perda_nao_descontada%></span>
    <% end if %>
	<% end if %>
</div>
<br>


<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<!-- ************   LINKS: PÁGINA INICIAL / ENCERRA SESSÃO   ************ -->
<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center"><div name="dVOLTAR" id="dVOLTAR"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fRetornar(f)" title="Retornar para o início do relatório">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>