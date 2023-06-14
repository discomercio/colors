<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================================
'	  CadPercMaxComissaoEDescPorLojaConfirma.asp
'     ===========================================================================
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

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_CAD_PERC_MAX_COMISSAO_E_DESC_POR_LOJA, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	Dim s_log, s_log_aux
	s_log = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	class cl_CadPercMaxComissaoEDescPorLoja
		dim num_loja
		dim nome_loja
		dim perc_max_comissao
		dim perc_max_comissao_e_desconto
		dim perc_max_comissao_e_desconto_pj
		dim perc_max_comissao_e_desconto_nivel2
		dim perc_max_comissao_e_desconto_nivel2_pj
		dim perc_max_comissao_alcada1
		dim perc_alcada1_pf
		dim perc_alcada1_pj
		dim perc_max_comissao_alcada2
		dim perc_alcada2_pf
		dim perc_alcada2_pj
		dim perc_max_comissao_alcada3
		dim perc_alcada3_pf
		dim perc_alcada3_pj
		end class

	dim vTabela
	redim vTabela(0)
	set vTabela(Ubound(vTabela)) = new cl_CadPercMaxComissaoEDescPorLoja
	
	dim i, n
	n = Request.Form("c_loja").Count
	for i = 1 to n
		s = Trim(Request.Form("c_loja")(i))
		if s <> "" then
			if Trim(vTabela(Ubound(vTabela)).num_loja) <> "" then
				redim preserve vTabela(Ubound(vTabela)+1)
				set vTabela(Ubound(vTabela)) = new cl_CadPercMaxComissaoEDescPorLoja
				end if
			vTabela(Ubound(vTabela)).num_loja = Trim(Request.Form("c_loja")(i))
			vTabela(Ubound(vTabela)).nome_loja = Trim(Request.Form("c_nome_loja")(i))
			vTabela(Ubound(vTabela)).perc_max_comissao = converte_numero(Trim(Request.Form("c_perc_comissao")(i)))
			vTabela(Ubound(vTabela)).perc_max_comissao_e_desconto = converte_numero(Trim(Request.Form("c_perc_comissao_e_desconto")(i)))
			vTabela(Ubound(vTabela)).perc_max_comissao_e_desconto_pj = converte_numero(Trim(Request.Form("c_perc_comissao_e_desconto_pj")(i)))
			vTabela(Ubound(vTabela)).perc_max_comissao_e_desconto_nivel2 = converte_numero(Trim(Request.Form("c_perc_comissao_e_desconto_nivel2")(i)))
			vTabela(Ubound(vTabela)).perc_max_comissao_e_desconto_nivel2_pj = converte_numero(Trim(Request.Form("c_perc_comissao_e_desconto_nivel2_pj")(i)))
			vTabela(Ubound(vTabela)).perc_max_comissao_alcada1 = converte_numero(Trim(Request.Form("c_perc_comissao_alcada1")(i)))
			vTabela(Ubound(vTabela)).perc_alcada1_pf = converte_numero(Trim(Request.Form("c_perc_alcada1_pf")(i)))
			vTabela(Ubound(vTabela)).perc_alcada1_pj = converte_numero(Trim(Request.Form("c_perc_alcada1_pj")(i)))
			vTabela(Ubound(vTabela)).perc_max_comissao_alcada2 = converte_numero(Trim(Request.Form("c_perc_comissao_alcada2")(i)))
			vTabela(Ubound(vTabela)).perc_alcada2_pf = converte_numero(Trim(Request.Form("c_perc_alcada2_pf")(i)))
			vTabela(Ubound(vTabela)).perc_alcada2_pj = converte_numero(Trim(Request.Form("c_perc_alcada2_pj")(i)))
			vTabela(Ubound(vTabela)).perc_max_comissao_alcada3 = converte_numero(Trim(Request.Form("c_perc_comissao_alcada3")(i)))
			vTabela(Ubound(vTabela)).perc_alcada3_pf = converte_numero(Trim(Request.Form("c_perc_alcada3_pf")(i)))
			vTabela(Ubound(vTabela)).perc_alcada3_pj = converte_numero(Trim(Request.Form("c_perc_alcada3_pj")(i)))
			end if
		next
	
	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	for i = Lbound(vTabela) to Ubound(vTabela)
		if Trim("" & vTabela(i).num_loja) <> "" then
			if vTabela(i).perc_max_comissao < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao_e_desconto < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (NÍVEL 1 - PF) NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao_e_desconto > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (NÍVEL 1 - PF) NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao_e_desconto_pj < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (NÍVEL 1 - PJ) NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao_e_desconto_pj > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (NÍVEL 1 - PJ) NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao_e_desconto_nivel2 < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (NÍVEL 2 - PF) NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao_e_desconto_nivel2 > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (NÍVEL 2 - PF) NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao_e_desconto_nivel2_pj < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (NÍVEL 2 - PJ) NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao_e_desconto_nivel2_pj > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (NÍVEL 2 - PJ) NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao_alcada1 < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ", ALÇADA 1)"
			elseif vTabela(i).perc_max_comissao_alcada1 > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ", ALÇADA 1)"
			elseif vTabela(i).perc_alcada1_pf < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 1 - PF) NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_alcada1_pf > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 1 - PF) NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_alcada1_pj < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 1 - PJ) NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_alcada1_pj > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 1 - PJ) NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao_alcada2 < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ", ALÇADA 2)"
			elseif vTabela(i).perc_max_comissao_alcada2 > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ", ALÇADA 2)"
			elseif vTabela(i).perc_alcada2_pf < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 2 - PF) NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_alcada2_pf > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 2 - PF) NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_alcada2_pj < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 2 - PJ) NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_alcada2_pj > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 2 - PJ) NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_max_comissao_alcada3 < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ", ALÇADA 3)"
			elseif vTabela(i).perc_max_comissao_alcada3 > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ", ALÇADA 3)"
			elseif vTabela(i).perc_alcada3_pf < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 3 - PF) NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_alcada3_pf > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 3 - PF) NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_alcada3_pj < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 3 - PJ) NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).perc_alcada3_pj > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL DE COMISSÃO+DESCONTO (ALÇADA 3 - PJ) NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
				end if
			end if
		next
	
	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	GRAVA OS DADOS!!
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		for i=Lbound(vTabela) to Ubound(vTabela)
			if Trim("" & vTabela(i).num_loja) <> "" then
				s = "SELECT * FROM t_LOJA WHERE (loja = '" & vTabela(i).num_loja & "')"
				if r.State <> 0 then r.Close
				r.Open s, cn
				if r.Eof then
					alerta = texto_add_br(alerta)
					alerta = alerta & "LOJA '" & vTabela(i).num_loja & "' NÃO FOI ENCONTRADA NO BANCO DE DADOS."
					end if
				
				if alerta = "" then
					s_log_aux = ""
					
					if converte_numero(r("perc_max_comissao")) <> converte_numero(vTabela(i).perc_max_comissao) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao: " & formata_perc(r("perc_max_comissao")) & " => " & formata_perc(vTabela(i).perc_max_comissao) & "]"
						r("perc_max_comissao") = vTabela(i).perc_max_comissao
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if
					
					if converte_numero(r("perc_max_comissao_e_desconto")) <> converte_numero(vTabela(i).perc_max_comissao_e_desconto) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_e_desconto: " & formata_perc(r("perc_max_comissao_e_desconto")) & " => " & formata_perc(vTabela(i).perc_max_comissao_e_desconto) & "]"
						r("perc_max_comissao_e_desconto") = vTabela(i).perc_max_comissao_e_desconto
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if
					
					if converte_numero(r("perc_max_comissao_e_desconto_pj")) <> converte_numero(vTabela(i).perc_max_comissao_e_desconto_pj) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_e_desconto_pj: " & formata_perc(r("perc_max_comissao_e_desconto_pj")) & " => " & formata_perc(vTabela(i).perc_max_comissao_e_desconto_pj) & "]"
						r("perc_max_comissao_e_desconto_pj") = vTabela(i).perc_max_comissao_e_desconto_pj
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if
					
					if converte_numero(r("perc_max_comissao_e_desconto_nivel2")) <> converte_numero(vTabela(i).perc_max_comissao_e_desconto_nivel2) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_e_desconto_nivel2: " & formata_perc(r("perc_max_comissao_e_desconto_nivel2")) & " => " & formata_perc(vTabela(i).perc_max_comissao_e_desconto_nivel2) & "]"
						r("perc_max_comissao_e_desconto_nivel2") = vTabela(i).perc_max_comissao_e_desconto_nivel2
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if
					
					if converte_numero(r("perc_max_comissao_e_desconto_nivel2_pj")) <> converte_numero(vTabela(i).perc_max_comissao_e_desconto_nivel2_pj) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_e_desconto_nivel2_pj: " & formata_perc(r("perc_max_comissao_e_desconto_nivel2_pj")) & " => " & formata_perc(vTabela(i).perc_max_comissao_e_desconto_nivel2_pj) & "]"
						r("perc_max_comissao_e_desconto_nivel2_pj") = vTabela(i).perc_max_comissao_e_desconto_nivel2_pj
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if
					
					'Alçada 1 - Máx RT
					if converte_numero(r("perc_max_comissao_alcada1")) <> converte_numero(vTabela(i).perc_max_comissao_alcada1) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_alcada1: " & formata_perc(r("perc_max_comissao_alcada1")) & " => " & formata_perc(vTabela(i).perc_max_comissao_alcada1) & "]"
						r("perc_max_comissao_alcada1") = vTabela(i).perc_max_comissao_alcada1
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if
					
					'Alçada 1 - PF
					if converte_numero(r("perc_max_comissao_e_desconto_alcada1_pf")) <> converte_numero(vTabela(i).perc_alcada1_pf) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_e_desconto_alcada1_pf: " & formata_perc(r("perc_max_comissao_e_desconto_alcada1_pf")) & " => " & formata_perc(vTabela(i).perc_alcada1_pf) & "]"
						r("perc_max_comissao_e_desconto_alcada1_pf") = vTabela(i).perc_alcada1_pf
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if

					'Alçada 1 - PJ
					if converte_numero(r("perc_max_comissao_e_desconto_alcada1_pj")) <> converte_numero(vTabela(i).perc_alcada1_pj) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_e_desconto_alcada1_pj: " & formata_perc(r("perc_max_comissao_e_desconto_alcada1_pj")) & " => " & formata_perc(vTabela(i).perc_alcada1_pj) & "]"
						r("perc_max_comissao_e_desconto_alcada1_pj") = vTabela(i).perc_alcada1_pj
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if

					'Alçada 2 - Máx RT
					if converte_numero(r("perc_max_comissao_alcada2")) <> converte_numero(vTabela(i).perc_max_comissao_alcada2) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_alcada2: " & formata_perc(r("perc_max_comissao_alcada2")) & " => " & formata_perc(vTabela(i).perc_max_comissao_alcada2) & "]"
						r("perc_max_comissao_alcada2") = vTabela(i).perc_max_comissao_alcada2
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if

					'Alçada 2 - PF
					if converte_numero(r("perc_max_comissao_e_desconto_alcada2_pf")) <> converte_numero(vTabela(i).perc_alcada2_pf) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_e_desconto_alcada2_pf: " & formata_perc(r("perc_max_comissao_e_desconto_alcada2_pf")) & " => " & formata_perc(vTabela(i).perc_alcada2_pf) & "]"
						r("perc_max_comissao_e_desconto_alcada2_pf") = vTabela(i).perc_alcada2_pf
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if

					'Alçada 2 - PJ
					if converte_numero(r("perc_max_comissao_e_desconto_alcada2_pj")) <> converte_numero(vTabela(i).perc_alcada2_pj) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_e_desconto_alcada2_pj: " & formata_perc(r("perc_max_comissao_e_desconto_alcada2_pj")) & " => " & formata_perc(vTabela(i).perc_alcada2_pj) & "]"
						r("perc_max_comissao_e_desconto_alcada2_pj") = vTabela(i).perc_alcada2_pj
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if

					'Alçada 3 - Máx RT
					if converte_numero(r("perc_max_comissao_alcada3")) <> converte_numero(vTabela(i).perc_max_comissao_alcada3) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_alcada3: " & formata_perc(r("perc_max_comissao_alcada3")) & " => " & formata_perc(vTabela(i).perc_max_comissao_alcada3) & "]"
						r("perc_max_comissao_alcada3") = vTabela(i).perc_max_comissao_alcada3
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if

					'Alçada 3 - PF
					if converte_numero(r("perc_max_comissao_e_desconto_alcada3_pf")) <> converte_numero(vTabela(i).perc_alcada3_pf) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_e_desconto_alcada3_pf: " & formata_perc(r("perc_max_comissao_e_desconto_alcada3_pf")) & " => " & formata_perc(vTabela(i).perc_alcada3_pf) & "]"
						r("perc_max_comissao_e_desconto_alcada3_pf") = vTabela(i).perc_alcada3_pf
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if

					'Alçada 3 - PJ
					if converte_numero(r("perc_max_comissao_e_desconto_alcada3_pj")) <> converte_numero(vTabela(i).perc_alcada3_pj) then
						if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
						s_log_aux = s_log_aux & "[perc_max_comissao_e_desconto_alcada3_pj: " & formata_perc(r("perc_max_comissao_e_desconto_alcada3_pj")) & " => " & formata_perc(vTabela(i).perc_alcada3_pj) & "]"
						r("perc_max_comissao_e_desconto_alcada3_pj") = vTabela(i).perc_alcada3_pj
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if

					if s_log_aux <> "" then
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & "(Loja " & vTabela(i).num_loja & ": " & s_log_aux & ")"
						end if
					end if
				end if
			next

		if r.State <> 0 then r.Close
		set r = nothing
		
		if alerta = "" then
			if s_log <> "" then
				s_log = "Alteração do percentual máximo de comissão e desconto por loja da(s) seguinte(s) loja(s): " & s_log
				grava_log usuario, "", "", "", OP_LOG_PERC_MAX_COMISSAO_E_DESC_POR_LOJA, s_log
				end if
			end if

		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err <> 0 then alerta=Cstr(Err) & ": " & Err.Description
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
		s = "<p style='margin:5px 2px 5px 2px;'>" & alerta & "</p>"
		s_aux="'MtAlerta'"
	else
		s = "DADOS ALTERADOS COM SUCESSO."
		if s <> "" then s="<p style='margin:5px 2px 5px 2px;'>" & s & "</p>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<br /><br />

<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
<%
	s="resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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
