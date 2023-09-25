<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  P E R F I L A T U A L I Z A . A S P
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
	
	dim s, s_aux, usuario, alerta
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r, rs, rsi
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim criou_novo_reg
	Dim s_log, s_log_itens, s_log_itens_anterior
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	s_log_itens_anterior = ""
	campos_a_omitir = "|dt_ult_atualizacao|usuario_ult_atualizacao|timestamp|"
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim s_apelido_perfil, s_descricao, operacao_selecionada, s_nivel_acesso_bloco_notas_pedido, s_nivel_acesso_chamado_pedido, rb_st_inativo
	operacao_selecionada=request("operacao_selecionada")
	s_apelido_perfil=UCase(trim(request("perfil_selecionado")))
	s_descricao=trim(request("c_descricao"))
	s_nivel_acesso_bloco_notas_pedido=trim(request("c_nivel_acesso_bloco_notas"))
    s_nivel_acesso_chamado_pedido=trim(request("c_nivel_acesso_chamado"))
	rb_st_inativo = Trim(Request("rb_st_inativo"))

	if s_apelido_perfil = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
		
	dim i, n, v_op_central, v_op_loja, v_op_orcto_cotacao, qtde_op_central, qtde_op_loja, qtde_op_orcto_cotacao, qtdeOpNivelAcessoBlocoNotas, qtdeOpNivelAcessoChamados
	qtdeOpNivelAcessoBlocoNotas = 0
    qtdeOpNivelAcessoChamados = 0
	
'	OPERAÇÕES DA CENTRAL
	qtde_op_central = 0
	redim v_op_central(0)
	v_op_central(0) = ""
	n = Request.Form("ckb_op_central").Count
	for i = 1 to n
		s = Trim(Request.Form("ckb_op_central")(i))
		if s <> "" then
			if Trim(v_op_central(ubound(v_op_central))) <> "" then
				redim preserve v_op_central(ubound(v_op_central)+1)
				v_op_central(ubound(v_op_central)) = ""
				end if
			v_op_central(ubound(v_op_central)) = s
			if (s = Cstr(OP_CEN_BLOCO_NOTAS_PEDIDO_LEITURA)) Or (s = Cstr(OP_CEN_BLOCO_NOTAS_PEDIDO_CADASTRAMENTO)) then qtdeOpNivelAcessoBlocoNotas=qtdeOpNivelAcessoBlocoNotas+1
			if (s = Cstr(OP_CEN_PEDIDO_CHAMADO_LEITURA_QUALQUER_CHAMADO)) Or (s = Cstr(OP_CEN_PEDIDO_CHAMADO_CADASTRAMENTO)) then qtdeOpNivelAcessoChamados=qtdeOpNivelAcessoChamados+1
			qtde_op_central = qtde_op_central + 1
			end if
		next
	
'	OPERAÇÕES DA LOJA
	qtde_op_loja = 0
	redim v_op_loja(0)
	v_op_loja(0) = ""
	n = Request.Form("ckb_op_loja").Count
	for i = 1 to n
		s = Trim(Request.Form("ckb_op_loja")(i))
		if s <> "" then
			if Trim(v_op_loja(ubound(v_op_loja))) <> "" then
				redim preserve v_op_loja(ubound(v_op_loja)+1)
				v_op_loja(ubound(v_op_loja)) = ""
				end if
			v_op_loja(ubound(v_op_loja)) = s
			if (s = Cstr(OP_LJA_BLOCO_NOTAS_PEDIDO_LEITURA)) Or (s = Cstr(OP_LJA_BLOCO_NOTAS_PEDIDO_CADASTRAMENTO)) then qtdeOpNivelAcessoBlocoNotas=qtdeOpNivelAcessoBlocoNotas+1
			if (s = Cstr(OP_LJA_PEDIDO_CHAMADO_LEITURA_QUALQUER_CHAMADO)) Or (s = Cstr(OP_LJA_PEDIDO_CHAMADO_CADASTRAMENTO)) then qtdeOpNivelAcessoChamados=qtdeOpNivelAcessoChamados+1
			qtde_op_loja = qtde_op_loja + 1
			end if
		next
	
'	OPERAÇÕES DO MÓDULO ORÇAMENTO/COTAÇÃO
	qtde_op_orcto_cotacao = 0
	redim v_op_orcto_cotacao(0)
	v_op_orcto_cotacao(0) = ""
	n = Request.Form("ckb_op_orcto_cotacao").Count
	for i = 1 to n
		s = Trim(Request.Form("ckb_op_orcto_cotacao")(i))
		if s <> "" then
			if Trim(v_op_orcto_cotacao(ubound(v_op_orcto_cotacao))) <> "" then
				redim preserve v_op_orcto_cotacao(ubound(v_op_orcto_cotacao)+1)
				v_op_orcto_cotacao(ubound(v_op_orcto_cotacao)) = ""
				end if
			v_op_orcto_cotacao(ubound(v_op_orcto_cotacao)) = s
			qtde_op_orcto_cotacao = qtde_op_orcto_cotacao + 1
			end if
		next
	
	dim erro_consistencia, erro_fatal
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if s_apelido_perfil = "" then
		alerta="IDENTIFICADOR DO PERFIL É INVÁLIDO."	
	elseif (operacao_selecionada = OP_INCLUI) And (s_apelido_perfil <> filtra_nome_identificador(s_apelido_perfil)) then
		alerta="IDENTIFICADOR CONTÉM CARACTERE(S) INVÁLIDO(S)!"
	elseif s_descricao = "" then
		alerta="PREENCHA A DESCRIÇÃO DO PERFIL."
	elseif rb_st_inativo = "" then
		alerta="STATUS INVÁLIDO DO PERFIL."
	elseif (qtdeOpNivelAcessoBlocoNotas>0) And (s_nivel_acesso_bloco_notas_pedido="") then
		alerta="O NÍVEL DE ACESSO PARA O BLOCO DE NOTAS DO PEDIDO NÃO FOI DEFINIDO."
	elseif (qtdeOpNivelAcessoBlocoNotas=0) And (s_nivel_acesso_bloco_notas_pedido<>"") then
		alerta="O NÍVEL DE ACESSO PARA O BLOCO DE NOTAS DO PEDIDO FOI DEFINIDO, MAS NENHUMA OPERAÇÃO REFERENTE AO BLOCO DE NOTAS FOI HABILITADA."
    elseif (qtdeOpNivelAcessoChamados>0) And (s_nivel_acesso_chamado_pedido="") then
		alerta="O NÍVEL DE ACESSO PARA OS CHAMADOS DO PEDIDO NÃO FOI DEFINIDO."
	elseif (qtdeOpNivelAcessoChamados=0) And (s_nivel_acesso_chamado_pedido<>"") then
		alerta="O NÍVEL DE ACESSO PARA OS CHAMADOS DO PEDIDO FOI DEFINIDO, MAS NENHUMA OPERAÇÃO DE LEITURA OU CADASTRAMENTO DE CHAMADOS FOI HABILITADA."
	elseif (qtde_op_central = 0) And (qtde_op_loja = 0) And (qtde_op_orcto_cotacao = 0) then
		alerta="NENHUMA OPERAÇÃO DA LISTA FOI SELECIONADA."
		end if
	
	if alerta <> "" then erro_consistencia=True	

	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(rsi, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim id_perfil, id_perfil_item
	id_perfil = ""
	
'	GERA O ID P/ O NOVO PERFIL?
	if alerta = "" then
		if operacao_selecionada = OP_INCLUI then
			if Not gera_nsu(NSU_CADASTRO_PERFIL, id_perfil, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
		else
			s = "SELECT * FROM t_PERFIL WHERE apelido = '" & QuotedStr(s_apelido_perfil) & "'"
			if r.State <> 0 then r.Close
			r.Open s, cn
			if Not r.Eof then 
				id_perfil = Cstr(r("id"))
			else
				alerta = "PERFIL " & chr(34) & s_apelido_perfil & chr(34) & " NÃO FOI LOCALIZADO NO BANCO DE DADOS."
				end if
			end if
		end if
		

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			if alerta = "" then
				s="SELECT COUNT(*) AS qtde FROM t_PERFIL_X_USUARIO INNER JOIN t_PERFIL ON t_PERFIL_X_USUARIO.id_perfil=t_PERFIL.id WHERE (t_PERFIL.apelido = '" & QuotedStr(s_apelido_perfil) & "')"
				if r.State <> 0 then r.Close
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "PERFIL NÃO PODE SER REMOVIDO PORQUE AINDA ESTÁ ASSOCIADO A " & formata_inteiro(r("qtde")) & " USUÁRIOS."
					end if
				if r.State <> 0 then r.Close
				
				if Not erro_fatal then
					s="SELECT Coalesce(COUNT(*),0) AS qtde FROM t_PERCENTUAL_COMISSAO_VENDEDOR WHERE (id_perfil = '" & id_perfil & "')"
					if r.State <> 0 then r.Close
					r.Open s, cn
				'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
					if CLng(r("qtde")) > CLng(0) then
						erro_fatal=True
						alerta = "PERFIL NÃO PODE SER REMOVIDO PORQUE AINDA ESTÁ ASSOCIADO A UMA TABELA DE COMISSÃO DOS VENDEDORES."
						end if
					if r.State <> 0 then r.Close
					end if
					
				if Not erro_fatal then
				'	INFO P/ LOG
					s="SELECT * FROM t_PERFIL WHERE apelido = '" & QuotedStr(s_apelido_perfil) & "'"
					if r.State <> 0 then r.Close
					r.Open s, cn
					if Not r.EOF then
						log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
						s_log = log_via_vetor_monta_exclusao(vLog1)
						end if
					if r.State <> 0 then r.Close
					
				'	APAGA!!
				'	~~~~~~~~~~~~~
					cn.BeginTrans
				'	~~~~~~~~~~~~~
					if id_perfil <> "" then
						s = "SELECT" & _
								" id_operacao" & _
							" FROM t_PERFIL_ITEM" & _
								" INNER JOIN t_OPERACAO ON (t_PERFIL_ITEM.id_operacao = t_OPERACAO.id)" & _
							" WHERE" & _
								" (id_perfil = '" & id_perfil & "')" & _
							" ORDER BY" & _
								" modulo," & _
								" ordenacao"
						if r.State <> 0 then r.Close
						r.Open s, cn
						s_log_itens = ""
						do while Not r.Eof
							if s_log_itens <> "" then s_log_itens = s_log_itens & ","
							s_log_itens = s_log_itens & Cstr(r("id_operacao"))
							r.MoveNext
							loop
						
						if s_log_itens = "" then s_log_itens = "(nenhuma)"
						if s_log_itens <> "" then s_log_itens = "operações=" & s_log_itens
						if (s_log <> "") And (s_log_itens <> "") then s_log = s_log & "; " & s_log_itens
						
						s="DELETE FROM t_PERFIL_ITEM WHERE id_perfil = '" & id_perfil & "'"
						cn.Execute(s)
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO REMOVER ITENS DO PERFIL (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if
						
					if Not erro_fatal then
						s="DELETE FROM t_PERFIL WHERE apelido = '" & QuotedStr(s_apelido_perfil) & "'"
						cn.Execute(s)
						If Err = 0 then 
							if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_PERFIL_EXCLUSAO, s_log
						else
							erro_fatal=True
							alerta = "FALHA AO REMOVER O PERFIL (" & Cstr(Err) & ": " & Err.Description & ")."
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
				

		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			if alerta = "" then 
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				s = "SELECT * FROM t_PERFIL WHERE apelido = '" & QuotedStr(s_apelido_perfil) & "'"
				if r.State <> 0 then r.Close
				r.Open s, cn
				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("id")=id_perfil
					r("apelido")=s_apelido_perfil
					r("dt_cadastro") = Date
					r("usuario_cadastro") = usuario
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir

					s = "SELECT" & _
							" id_operacao" & _
						" FROM t_PERFIL_ITEM" & _
							" INNER JOIN t_OPERACAO ON (t_PERFIL_ITEM.id_operacao = t_OPERACAO.id)" & _
						" WHERE" & _
							" (id_perfil = '" & id_perfil & "')" & _
						" ORDER BY" & _
							" modulo," & _
							" ordenacao"
					if rsi.State <> 0 then rsi.Close
					rsi.Open s, cn
					s_log_itens_anterior = ""
					do while Not rsi.Eof
						if s_log_itens_anterior <> "" then s_log_itens_anterior = s_log_itens_anterior & ","
						s_log_itens_anterior = s_log_itens_anterior & Cstr(rsi("id_operacao"))
						rsi.MoveNext
						loop

					if s_log_itens_anterior = "" then s_log_itens_anterior = "(nenhuma)"
					if s_log_itens_anterior <> "" then s_log_itens_anterior = "operações (anterior): " & s_log_itens_anterior
					
					if rsi.State <> 0 then rsi.Close
					end if
					
				r("descricao")=s_descricao
				r("nivel_acesso_bloco_notas_pedido")=converte_numero(s_nivel_acesso_bloco_notas_pedido)
                r("nivel_acesso_chamado")=converte_numero(s_nivel_acesso_chamado_pedido)
				r("dt_ult_atualizacao") = Now
				r("usuario_ult_atualizacao") = usuario
				r("st_inativo") = CInt(rb_st_inativo)
				
				r.Update

				If Err = 0 then 
					s = "UPDATE t_PERFIL_ITEM SET excluido_status = 1 WHERE id_perfil = '" & id_perfil & "'"
					cn.Execute(s)
					end if

				if Err = 0 then
				'	CENTRAL
					for i = Lbound(v_op_central) to Ubound(v_op_central)
						if Trim(v_op_central(i)) <> "" then
							s = "SELECT * FROM t_PERFIL_ITEM WHERE (id_perfil = '" & id_perfil & "') AND (id_operacao = " & Trim(v_op_central(i)) & ")"
							if rs.State <> 0 then rs.Close
							rs.Open s, cn
							if Not rs.Eof then
								rs("excluido_status") = 0
							else
								if Not gera_nsu(NSU_CADASTRO_PERFIL_ITEM, id_perfil_item, msg_erro) then 
									alerta = "FALHA AO GERAR O NSU DO ITEM DE PERFIL (" & Cstr(Err) & ": " & Err.Description & ")."
									erro_fatal = True
									exit for
									end if
								rs.AddNew
								rs("id") = id_perfil_item
								rs("id_perfil") = id_perfil
								rs("id_operacao") = Trim(v_op_central(i))
								rs("dt_cadastro") = Date
								rs("usuario_cadastro") = usuario
								end if
							rs.Update
							if Err <> 0 then exit for
							end if
						next
					end if

				if Err = 0 then
				'	LOJA
					for i = Lbound(v_op_loja) to Ubound(v_op_loja)
						if Trim(v_op_loja(i)) <> "" then
							s = "SELECT * FROM t_PERFIL_ITEM WHERE (id_perfil = '" & id_perfil & "') AND (id_operacao = " & Trim(v_op_loja(i)) & ")"
							if rs.State <> 0 then rs.Close
							rs.Open s, cn
							if Not rs.Eof then
								rs("excluido_status") = 0
							else
								if Not gera_nsu(NSU_CADASTRO_PERFIL_ITEM, id_perfil_item, msg_erro) then 
									alerta = "FALHA AO GERAR O NSU DO ITEM DE PERFIL (" & Cstr(Err) & ": " & Err.Description & ")."
									erro_fatal = True
									exit for
									end if
								rs.AddNew
								rs("id") = id_perfil_item
								rs("id_perfil") = id_perfil
								rs("id_operacao") = Trim(v_op_loja(i))
								rs("dt_cadastro") = Date
								rs("usuario_cadastro") = usuario
								end if
							rs.Update
							if Err <> 0 then exit for
							end if
						next
					end if
				
				if Err = 0 then
				'	ORÇAMENTO/COTAÇÃO
					for i = Lbound(v_op_orcto_cotacao) to Ubound(v_op_orcto_cotacao)
						if Trim(v_op_orcto_cotacao(i)) <> "" then
							s = "SELECT * FROM t_PERFIL_ITEM WHERE (id_perfil = '" & id_perfil & "') AND (id_operacao = " & Trim(v_op_orcto_cotacao(i)) & ")"
							if rs.State <> 0 then rs.Close
							rs.Open s, cn
							if Not rs.Eof then
								rs("excluido_status") = 0
							else
								if Not gera_nsu(NSU_CADASTRO_PERFIL_ITEM, id_perfil_item, msg_erro) then 
									alerta = "FALHA AO GERAR O NSU DO ITEM DE PERFIL (" & Cstr(Err) & ": " & Err.Description & ")."
									erro_fatal = True
									exit for
									end if
								rs.AddNew
								rs("id") = id_perfil_item
								rs("id_perfil") = id_perfil
								rs("id_operacao") = Trim(v_op_orcto_cotacao(i))
								rs("dt_cadastro") = Date
								rs("usuario_cadastro") = usuario
								end if
							rs.Update
							if Err <> 0 then exit for
							end if
						next
					end if
				
				if Err = 0 then
					s = "DELETE FROM t_PERFIL_ITEM WHERE (id_perfil = '" & id_perfil & "') AND (excluido_status <> 0)"
					cn.Execute(s)
					end if
					
				If Err = 0 then 
					s_log_itens = ""
					for i=Lbound(v_op_central) to Ubound(v_op_central)
						if Trim(v_op_central(i)) <> "" then
							if s_log_itens <> "" then s_log_itens = s_log_itens & ","
							s_log_itens = s_log_itens & v_op_central(i)
							end if
						next
					
					for i=Lbound(v_op_loja) to Ubound(v_op_loja)
						if Trim(v_op_loja(i)) <> "" then
							if s_log_itens <> "" then s_log_itens = s_log_itens & ","
							s_log_itens = s_log_itens & v_op_loja(i)
							end if
						next
					
					for i=Lbound(v_op_orcto_cotacao) to Ubound(v_op_orcto_cotacao)
						if Trim(v_op_orcto_cotacao(i)) <> "" then
							if s_log_itens <> "" then s_log_itens = s_log_itens & ","
							s_log_itens = s_log_itens & v_op_orcto_cotacao(i)
							end if
						next
					
					if s_log_itens = "" then s_log_itens = "(nenhuma)"
					if s_log_itens <> "" then s_log_itens = "operações (atual): " & s_log_itens
					
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then 
							if s_log_itens <> "" then s_log = s_log & "; " & s_log_itens
							grava_log usuario, "", "", "", OP_LOG_PERFIL_INCLUSAO, s_log
							end if
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if (s_log <> "") Or (s_log_itens <> "") Or (s_log_itens_anterior <> "") then
							if s_log <> "" then s_log = "; " & s_log
							s_log="perfil=" & Trim("" & r("apelido")) & s_log
							if s_log_itens_anterior <> "" then s_log = s_log & "; " & s_log_itens_anterior
							if s_log_itens <> "" then s_log = s_log & "; " & s_log_itens
							grava_log usuario, "", "", "", OP_LOG_PERFIL_ALTERACAO, s_log
							end if
						end if
				else
					erro_fatal=True
					alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
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
		s = "<p style='margin:5px 2px 5px 2px;'>" & alerta & "</P>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "PERFIL " & chr(34) & s_apelido_perfil & chr(34) & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "PERFIL " & chr(34) & s_apelido_perfil & chr(34) & " ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "PERFIL " & chr(34) & s_apelido_perfil & chr(34) & " EXCLUÍDO COM SUCESSO."
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
	s="Perfil.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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