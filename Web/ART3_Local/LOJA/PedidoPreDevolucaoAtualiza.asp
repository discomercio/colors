<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/global.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================
'	  P E D I D O P R E D E V O L U C A O A T U A L I Z A . A S P
'     ===========================================================
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

	dim s, usuario, pedido_selecionado, id_pedido_base
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, sx, r, rsMail
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim vLog1()
	dim vLog2()
	dim campos_a_omitir
	campos_a_omitir = "|dt_hr_ult_atualizacao|"

	dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, vl_saldo_a_pagar, st_pagto, st_pagto_a
	dim iv, j, k, n, v_devol, alerta, deve_devolver, s_log, s_log_aux, s_log_indicador_desconto, msg_erro, id_item_devolvido, id_orcamentista_indicador_desconto
    dim id_chamado, cod_motivo_abertura_chamado, texto_chamado
    dim c_procedimento, c_local_coleta, c_coleta_endereco, c_coleta_endereco_numero, c_coleta_bairro, c_coleta_cep, c_coleta_cidade, c_coleta_uf, c_coleta_complemento
    dim c_motivo_devolucao, c_motivo_descricao, c_taxa, c_taxa_forma_pagto, c_taxa_percentual, c_taxa_valor, c_taxa_responsavel, c_credito_transacao, c_pedido_novo, c_observacoes
    dim c_cliente_banco, c_cliente_agencia, c_cliente_conta, c_cliente_favorecido, c_total_devolucao, c_pedido_cliente, c_pedido_indicador, taxa_valor_a_cobrar, c_taxa_observacoes, c_credito_observacoes
    dim rb_status, id_devolucao, c_devolucao_usuario, c_vendedor, st_devolucao_atual
    dim corpo_mensagem, msg_erro_grava_email, id_email, s_destinatario, s_remetente
    dim upload_file_guid, v_upload_file_guid, guid_item, id_upload_file, delete_file_scheduled_date
    dim c_pedido_possui_parcela_cartao
    dim url_back, rb_status_back, c_loja
	
    c_procedimento = Request.Form("c_procedimento")
    c_local_coleta = Request.Form("c_local_coleta")
    c_coleta_endereco = Trim(Request.Form("c_coleta_endereco"))
    c_coleta_endereco_numero = Trim(Request.Form("c_coleta_endereco_numero"))
    c_coleta_bairro = Trim(Request.Form("c_coleta_bairro"))
    c_coleta_cep = retorna_so_digitos(Trim(Request.Form("c_coleta_cep")))
    c_coleta_cidade = Trim(Request.Form("c_coleta_cidade"))
    c_coleta_uf = Trim(Request.Form("c_coleta_uf"))
    c_coleta_complemento = Trim(Request.Form("c_coleta_complemento"))
    c_motivo_devolucao = Request.Form("c_motivo_devolucao")
    c_motivo_descricao = Trim(Request.Form("c_motivo_descricao"))
    c_taxa = Request.Form("c_taxa")
    c_taxa_forma_pagto = Request.Form("c_taxa_forma_pagto")
    c_taxa_percentual = Request.Form("c_taxa_percentual")
    c_taxa_valor = Request.Form("c_taxa_valor")
    c_taxa_responsavel = Request.Form("c_taxa_responsavel")
    c_credito_transacao = Request.Form("c_credito_transacao")
    c_pedido_novo = Trim(Request.Form("c_pedido_novo"))
    c_observacoes = Trim(Request.Form("c_observacoes"))
    c_cliente_banco = Trim(Request.Form("c_cliente_banco"))
    c_cliente_agencia = Trim(Request.Form("c_cliente_agencia"))
    c_cliente_conta = Trim(Request.Form("c_cliente_conta"))
    c_cliente_favorecido = Trim(Request.Form("c_cliente_favorecido"))
    c_total_devolucao = Request.Form("c_total_geral")
    c_pedido_cliente = Request.Form("c_pedido_cliente")
    rb_status = Request.Form("rb_status")
    id_devolucao = Request.Form("id_devolucao")
    c_devolucao_usuario = Request.Form("c_devolucao_usuario")
    c_pedido_indicador = Trim(Request.Form("c_pedido_indicador"))
    c_vendedor = Trim(Request.Form("c_vendedor"))
    upload_file_guid = Trim(Request.Form("upload_file_guid_returned"))
    st_devolucao_atual = Request.Form("st_devolucao")
    c_taxa_observacoes = Trim(Request.Form("c_taxa_observacoes"))
    c_credito_observacoes = Trim(Request.Form("c_credito_observacoes"))
    c_pedido_possui_parcela_cartao = Request.Form("c_pedido_possui_parcela_cartao")

    redim v_devol(0)
	set v_devol(Ubound(v_devol)) = New cl_ITEM_DEVOLUCAO_MERCADORIAS
	v_devol(Ubound(v_devol)).produto = ""
	v_devol(Ubound(v_devol)).qtde_a_devolver = 0
	v_devol(Ubound(v_devol)).motivo = ""
    v_devol(Ubound(v_devol)).vl_unitario = 0

	s_log = ""
	alerta = ""
	deve_devolver = False

    dim id_pedido_chamado_depto
    dim r_pedido, r_vendedor
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then
		alerta = msg_erro
		end if

    if alerta = "" then
        call le_usuario(r_pedido.vendedor, r_vendedor, msg_erro)
        end if

    if (r_pedido.st_forma_pagto_possui_parcela_cartao = 1) Or (r_pedido.st_forma_pagto_possui_parcela_cartao_maquineta = 1) then
        '2=Financeiro/Devolução (Pagamento em Cartão) -> para pedidos que tenham pagamento em cartão
        id_pedido_chamado_depto = 3
    else
        '2=Financeiro/Devolução -> para pedidos que não tenham pagamento em cartão
        id_pedido_chamado_depto = 2
        end if

    if c_procedimento = "" then
        alerta = texto_add_br(alerta)
        alerta = alerta & "Procedimento não foi informado."
        end if
    if c_local_coleta = "" then
        alerta = texto_add_br(alerta)
		alerta = alerta & "Local de coleta não foi informado."
        end if
    if c_coleta_cep <> "" then
        if Not cep_ok(c_coleta_cep) then
            alerta = texto_add_br(alerta)
		    alerta = alerta & "CEP do endereço de coleta inválido."
            end if
        end if
    if c_motivo_devolucao = "" then
        alerta = texto_add_br(alerta)
		alerta = alerta & "Motivo da devolução não foi informado."
        end if

    if Len(c_motivo_descricao) > 600 then
        alerta = texto_add_br(alerta)
		alerta = alerta & "Descrição do motivo da devolução (" & Cstr(len(c_motivo_descricao)) & ") excede o tamanho máximo permitido de " & Cstr(600) & " caracteres."
        end if
    if Len(c_observacoes) > 600 then
        alerta = texto_add_br(alerta)
		alerta = alerta & "Observações gerais (" & Cstr(len(c_observacoes)) & ") excede o tamanho máximo permitido de " & Cstr(600) & " caracteres."
        end if
    if Len(c_taxa_observacoes) > 600 then
        alerta = texto_add_br(alerta)
		alerta = alerta & "Observações da taxa administrativa (" & Cstr(len(c_taxa_observacoes)) & ") excede o tamanho máximo permitido de " & Cstr(600) & " caracteres."
        end if
    if Len(c_credito_observacoes) > 600 then
        alerta = texto_add_br(alerta)
		alerta = alerta & "Observações do crédito (" & Cstr(len(c_credito_observacoes)) & ") excede o tamanho máximo permitido de " & Cstr(600) & " caracteres."
        end if

    if st_devolucao_atual = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA then
        if c_credito_transacao = "" then
            alerta = texto_add_br(alerta)
		    alerta = alerta & "Transação não foi informada."
            end if
        if c_taxa = TAXA_ADMINISTRATIVA__SIM then
            if c_taxa_forma_pagto = "" then
                alerta = texto_add_br(alerta)
		        alerta = alerta & "Forma de pagamento da taxa administrativa não foi informada."
                end if
            if c_taxa_forma_pagto = COD_PEDIDO_DEVOLUCAO_TAXA_FORMA_PAGTO__DESCONTO_COMISSAO then
                if c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__CLIENTE then
                    alerta = texto_add_br(alerta)
		            alerta = alerta & "Responsável pelo pagamento da taxa administrativa não pode ser o cliente quando a forma de pagamento é 'Desconto de Comissão'."
                    end if
                end if
            if c_taxa_forma_pagto = COD_PEDIDO_DEVOLUCAO_TAXA_FORMA_PAGTO__ABATIMENTO_CREDITO then
                if c_taxa_responsavel <> COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__CLIENTE then
                    alerta = texto_add_br(alerta)
		            alerta = alerta & "Responsável pelo pagamento da taxa administrativa deve ser o cliente quando a forma de pagamento é 'Abatimento no crédito'."
                    end if
                end if
            if c_taxa_percentual = "" then
                alerta = texto_add_br(alerta)
		        alerta = alerta & "Percentual da taxa administrativa não foi informado."
                end if
            if c_taxa_responsavel = "" then
                alerta = texto_add_br(alerta)
		        alerta = alerta & "Responsável pelo pagamento da taxa administrativa não foi informado."
                end if
            if c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__PARCEIRO Or _
                    c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__VENDEDOR_PARCEIRO then
                if c_pedido_indicador = "" then
                    alerta = texto_add_br(alerta)
		            alerta = alerta & "Não existe parceiro vinculado ao pedido. Selecione outro responsável pela taxa administrativa."
                    end if
                end if
            end if
        if c_credito_transacao = CREDITO_TRANSACAO__REEMBOLSO then
            if c_cliente_banco = "" then
                alerta = texto_add_br(alerta)
		        alerta = alerta & "Código do banco não foi preenchido."
                end if
            if c_cliente_agencia = "" then
                alerta = texto_add_br(alerta)
		        alerta = alerta & "Número da agência não foi preenchida."
                end if
            if c_cliente_conta = "" then
                alerta = texto_add_br(alerta)
		        alerta = alerta & "Número da conta não foi preenchida."
                end if
            if c_cliente_favorecido = "" then
                alerta = texto_add_br(alerta)
		        alerta = alerta & "Favorecido da conta não foi preenchido."
                end if
        elseif c_credito_transacao = CREDITO_TRANSACAO__TRANSFERENCIA then
            if c_pedido_novo = "" then
                alerta = texto_add_br(alerta)
		        alerta = alerta & "Pedido novo para transferência do crédito não foi informado."
                end if
        elseif c_credito_transacao = CREDITO_TRANSACAO__ESTORNO then
            if c_pedido_possui_parcela_cartao <> "1" then
                alerta = texto_add_br(alerta)
		        alerta = alerta & "A transação do crédito não pode ser 'Estorno' porque o pedido não possui pagamento via cartão de crédito."
                end if
            end if
        end if

    if alerta = "" then
	    n = Request.Form("c_qtde_devolucao").Count
	    for iv = 1 to n
		    s=Trim(Request.Form("c_produto")(iv))
		    if s <> "" then
			    if Trim(v_devol(Ubound(v_devol)).produto) <> "" then
				    redim preserve v_devol(ubound(v_devol)+1)
				    set v_devol(ubound(v_devol)) = New cl_ITEM_DEVOLUCAO_MERCADORIAS
				    end if
			    with v_devol(ubound(v_devol))
				    .pedido=pedido_selecionado
				    .produto=Ucase(Trim(Request.Form("c_produto")(iv)))
				
				    s=retorna_so_digitos(Request.Form("c_fabricante")(iv))
				    .fabricante=normaliza_codigo(s, TAM_MIN_FABRICANTE)
				
				    s = Trim(Request.Form("c_qtde_devolucao")(iv))
				    if IsNumeric(s) then .qtde_a_devolver = CLng(s) else .qtde_a_devolver = 0

                    s = Trim(Request.Form("c_vl_unitario")(iv))
                    .vl_unitario = converte_numero(s)
				
				    if .qtde_a_devolver > 0 then deve_devolver = True
				    end with
			    end if
		    next

        if Not deve_devolver then
		    alerta = "Não foi especificado nenhum produto para a operação de devolução."
	        end if

        end if ' alerta

	dim s_dados_cliente
	dim s_unidade_negocio
	dim s_descricao_status_devolucao, s_cor_status_devolucao
	dim strEmailDestinatarioAlerta
	dim strEmailAdministrador
	set strEmailAdministrador = get_registro_t_parametro("PEDIDO_DEVOLUCAO_EMAIL_ADMINISTRADOR_2")
	'16/09/2020: a Gabriela Hernandes solicitou para enviar o email de alerta para o vendedor apenas
	'strEmailDestinatarioAlerta = Trim("" & strEmailAdministrador.campo_texto)
	strEmailDestinatarioAlerta = Trim("" & r_vendedor.email)

	if Not cria_recordset_otimista(rsMail, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	s_dados_cliente = ""
	s_unidade_negocio = ""
	if strEmailDestinatarioAlerta <> "" then
		s = "SELECT" & _
				" p.st_memorizacao_completa_enderecos," & _
				" p.endereco_nome_iniciais_em_maiusculas AS endereco_nome," & _
				" p.endereco_cnpj_cpf," & _
				" c.nome_iniciais_em_maiusculas AS cliente_nome," & _
				" c.cnpj_cpf AS cliente_cnpj_cpf," & _
				" p.loja," & _
				" lj.unidade_negocio" & _
			" FROM t_PEDIDO p" & _
				" INNER JOIN t_CLIENTE c ON (p.id_cliente = c.id)" & _
				" INNER JOIN t_LOJA lj ON (p.loja = lj.loja)" & _
			" WHERE" & _
				" (p.pedido = '" & pedido_selecionado & "')"
		if rsMail.State <> 0 then rsMail.Close
		rsMail.Open s, cn
		if Not rsMail.Eof then
			s_unidade_negocio = Trim("" & rsMail("unidade_negocio"))
			if rsMail("st_memorizacao_completa_enderecos") <> 0 then
				s_dados_cliente = "Cliente: " & Trim("" & rsMail("endereco_nome")) & " (" & cnpj_cpf_formata(Trim("" & rsMail("endereco_cnpj_cpf"))) & ")"
			else
				s_dados_cliente = "Cliente: " & Trim("" & rsMail("cliente_nome")) & " (" & cnpj_cpf_formata(Trim("" & rsMail("cliente_cnpj_cpf"))) & ")"
				end if
			end if
		end if


	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
        if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

		if Not cria_recordset_otimista(sx, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if        

        s = "SELECT * FROM t_PEDIDO_DEVOLUCAO WHERE (id='" & id_devolucao & "')"
        if rs.State <> 0 then rs.Close
        rs.Open s, cn

        if Not rs.Eof then
			log_via_vetor_carrega_do_recordset rs, vLog1, campos_a_omitir

            if rb_status <> "" then
				obtem_descricao_status_devolucao rb_status, s_descricao_status_devolucao, s_cor_status_devolucao

                rs("status") = rb_status
                rs("status_usuario") = usuario
                rs("status_data") = Date
                rs("status_data_hora") = Now
                if rb_status = COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO then
                    rs("st_aprovado") = 1
                    rs("usuario_aprovado") = usuario
                    rs("dt_aprovado") = Date
                    rs("dt_hr_aprovado") = Now
                    delete_file_scheduled_date = DateAdd("m", 2, Date)
        
                    if strEmailDestinatarioAlerta <> "" then
                        corpo_mensagem = "O status da devolução nº " & id_devolucao & " do pedido " & pedido_selecionado & " foi alterado para '" & s_descricao_status_devolucao & "' por " & usuario & " em " & formata_data_hora_sem_seg(Now) & _
										vbCrLf & _
										"Pedido: " & pedido_selecionado & _
										vbCrLf & _
										"Devolução nº " & id_devolucao & _
										vbCrLf & _
										s_dados_cliente & _
										vbCrLf & vbCrLf & _
										"Atenção: esta é uma mensagem automática, NÃO responda a este e-mail!"
                        EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__PEDIDO_DEVOLUCAO), _
                                                        "", _
                                                        strEmailDestinatarioAlerta, _
                                                        "", _
                                                        "", _
                                                        "Status da devolução nº " & id_devolucao & " do pedido " & pedido_selecionado & " alterado para '" & s_descricao_status_devolucao & "'", _
                                                        corpo_mensagem, _
                                                        Now, _
                                                        id_email, _
                                                        msg_erro_grava_email
                        end if

                elseif rb_status = COD_ST_PEDIDO_DEVOLUCAO__REPROVADA then
                    rs("st_reprovado") = 1
                    rs("usuario_reprovado") = usuario
                    rs("dt_reprovado") = Date
                    rs("dt_hr_reprovado") = Now
                    delete_file_scheduled_date = DateAdd("m", 1, Date)

					if strEmailDestinatarioAlerta <> "" then
						corpo_mensagem = "O status da devolução nº " & id_devolucao & " do pedido " & pedido_selecionado & " foi alterado para '" & s_descricao_status_devolucao & "' por " & usuario & " em " & formata_data_hora_sem_seg(Now) & _
										vbCrLf & _
										"Pedido: " & pedido_selecionado & _
										vbCrLf & _
										"Devolução nº " & id_devolucao & _
										vbCrLf & _
										s_dados_cliente & _
										vbCrLf & vbCrLf & _
										"Atenção: esta é uma mensagem automática, NÃO responda a este e-mail!"
						EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__PEDIDO_DEVOLUCAO), _
														"", _
														strEmailDestinatarioAlerta, _
														"", _
														"", _
														"Status da devolução nº " & id_devolucao & " do pedido " & pedido_selecionado & " alterado para '" & s_descricao_status_devolucao & "'", _
														corpo_mensagem, _
														Now, _
														id_email, _
														msg_erro_grava_email
						end if
					end if

                if Not cria_recordset_otimista(r, msg_erro) then
		            '	~~~~~~~~~~~~~~~~
			            cn.RollbackTrans
		            '	~~~~~~~~~~~~~~~~
			            Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			            end if    
                    
                    s = "SELECT * FROM t_UPLOAD_FILE INNER JOIN t_PEDIDO_DEVOLUCAO_IMAGEM ON (t_UPLOAD_FILE.id=t_PEDIDO_DEVOLUCAO_IMAGEM.id_upload_file)" & _
                            " WHERE (id_pedido_devolucao = " & id_devolucao & ")"
			        r.Open s, cn
                    if Not r.Eof then
                        do while Not r.Eof
                            r("st_delete_file")=1
                            r("usuario_delete_file")=usuario
                            r("dt_delete_file")=Date
                            r("dt_hr_delete_file")=Now
                            r("dt_delete_file_scheduled_date")=delete_file_scheduled_date
                            r.Update
                            r.MoveNext
                            loop
                        end if
                        if Err <> 0 then
			        '	~~~~~~~~~~~~~~~~
				        cn.RollbackTrans
			        '	~~~~~~~~~~~~~~~~
				        Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				        end if

                url_back = Request.Form("url_back")
                rb_status_back = Request.Form("rb_status_back")
                c_loja = Request.Form("c_loja")
        
                end if
            
            rs("cod_procedimento") = c_procedimento
            rs("cod_local_coleta") = c_local_coleta
            rs("endereco_coleta_logradouro") = c_coleta_endereco
            rs("endereco_coleta_numero") = c_coleta_endereco_numero
            rs("endereco_coleta_bairro") = c_coleta_bairro
            rs("endereco_coleta_cidade") = c_coleta_cidade
            rs("endereco_coleta_uf") = c_coleta_uf
            rs("endereco_coleta_cep") = c_coleta_cep
            rs("endereco_coleta_complemento") = c_coleta_complemento
            rs("cod_devolucao_motivo") = c_motivo_devolucao
            rs("motivo_observacao") = c_motivo_descricao
            
			if ( Ucase(Trim("" & rs("taxa_observacoes"))) <> Ucase(Trim("" & c_taxa_observacoes)) ) Or _
				( (st_devolucao_atual = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA) And (Trim("" & rs("taxa_flag")) <> Trim("" & c_taxa)) ) Or _
				
				( (st_devolucao_atual = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA) And (c_taxa = TAXA_ADMINISTRATIVA__SIM) And _
					( _
					(converte_numero(rs("taxa_percentual")) <> converte_numero(c_taxa_percentual)) Or _
					(Trim("" & rs("cod_taxa_forma_pagto")) <> Trim("" & c_taxa_forma_pagto)) Or _
					(Trim("" & rs("cod_taxa_responsavel")) <> Trim("" & c_taxa_responsavel)) _
					) _
				) then
				rs("taxa_usuario_ult_atualizacao") = usuario
				rs("taxa_dt_hr_ult_atualizacao") = Now
				end if

			if ( (c_credito_observacoes <> "") And (Ucase(Trim("" & rs("credito_observacoes"))) <> Ucase(Trim("" & c_credito_observacoes))) ) Or _
				( (c_pedido_novo <> "") And (Ucase(Trim("" & rs("pedido_novo"))) <> Ucase(Trim("" & c_pedido_novo))) ) Or _
				( (st_devolucao_atual = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA) And (Trim("" & rs("cod_credito_transacao")) <> Trim("" & c_credito_transacao)) ) Or _
				( (st_devolucao_atual = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA) And (c_credito_transacao = CREDITO_TRANSACAO__REEMBOLSO) And _
					( _
						( Ucase(Trim("" & rs("banco"))) <> Ucase(Trim("" & c_cliente_banco)) ) Or _
						( Ucase(Trim("" & rs("agencia"))) <> Ucase(Trim("" & c_cliente_agencia)) ) Or _
						( Ucase(Trim("" & rs("conta"))) <> Ucase(Trim("" & c_cliente_conta)) ) Or _
						( Ucase(Trim("" & rs("favorecido"))) <> Ucase(Trim("" & c_cliente_favorecido)) ) _
					) _
				) then
				rs("credito_usuario_ult_atualizacao") = usuario
				rs("credito_dt_hr_ult_atualizacao") = Now
				end if

			if st_devolucao_atual = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA then
                if c_taxa = TAXA_ADMINISTRATIVA__SIM Or rs("taxa_flag") <> c_taxa then
                    rs("taxa_percentual") = converte_numero(c_taxa_percentual)
                    rs("cod_taxa_forma_pagto") = c_taxa_forma_pagto
                    rs("cod_taxa_responsavel") = c_taxa_responsavel
                    end if
                rs("taxa_flag") = c_taxa
                rs("vl_devolucao") = converte_numero(c_total_devolucao)
                rs("cod_credito_transacao") = c_credito_transacao
                if c_credito_transacao = CREDITO_TRANSACAO__REEMBOLSO then
                    rs("banco") = c_cliente_banco
                    rs("agencia") = c_cliente_agencia
                    rs("conta") = c_cliente_conta
                    rs("favorecido") = c_cliente_favorecido
                    end if
                end if

            if c_pedido_novo <> "" then rs("pedido_novo") = c_pedido_novo
            if c_taxa_observacoes <> "" then rs("taxa_observacoes") = c_taxa_observacoes
            if c_credito_observacoes <> "" then rs("credito_observacoes") = c_credito_observacoes
            rs("observacao") = c_observacoes
            rs("usuario_ult_atualizacao") = usuario
            rs("dt_ult_atualizacao") = Date
            rs("dt_hr_ult_atualizacao") = Now
            rs.Update
            if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			else
				log_via_vetor_carrega_do_recordset rs, vLog2, campos_a_omitir
				s_log_aux = log_via_vetor_monta_alteracao(vLog1, vLog2)
				if s_log_aux <> "" then
					if s_log <> "" then s_log = s_log & "; "
					s_log = s_log & s_log_aux
					end if
				end if
            end if

        if rb_status <> "" then

            if rb_status = COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO then
                
                ' PROCESSA DEVOLUÇÃO DOS ITENS
                ' ============================
                for iv = Lbound(v_devol) to Ubound(v_devol)
			    with v_devol(iv)
				    if (Trim(.produto)<>"") And (.qtde_a_devolver > 0) then
				    '	REGISTRA O PRODUTO DEVOLVIDO
					    s = "SELECT * FROM t_PEDIDO_ITEM" & _
						    " WHERE (pedido='" & .pedido & "')" & _
						    " AND (fabricante='" & .fabricante & "')" & _
						    " AND (produto='" & .produto & "')"
					    if sx.State <> 0 then sx.Close
					    sx.open s, cn
					    if Err <> 0 then
					    '	~~~~~~~~~~~~~~~~
						    cn.RollbackTrans
					    '	~~~~~~~~~~~~~~~~
						    Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						    end if
		
					    if sx.Eof then
						    alerta = texto_add_br(alerta)
						    alerta = alerta & "Item de pedido referente ao produto " & .produto & " do fabricante " & _
										      .fabricante & " não foi encontrado."
					    else
						    if Not gera_nsu(NSU_PEDIDO_ITEM_DEVOLVIDO, id_item_devolvido, msg_erro) then 
						    '	~~~~~~~~~~~~~~~~
							    cn.RollbackTrans
						    '	~~~~~~~~~~~~~~~~
							    Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
							    end if

						    s = "SELECT * FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (pedido='X')"
						    if rs.State <> 0 then rs.Close
						    rs.Open s, cn
						    rs.AddNew
						    for j = 0 to rs.Fields.Count-1
							    for k = 0 to sx.Fields.Count-1
								    if Ucase(rs.Fields(j).Name)=Ucase(sx.Fields(k).Name) then
									    rs.Fields(j) = sx.Fields(k)
									    exit for
									    end if
								    next
							    next
						
						    rs("id") = id_item_devolvido
						    rs("qtde") = .qtde_a_devolver
						    rs("motivo") = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__MOTIVO, c_motivo_devolucao)
						    rs("devolucao_data") = Date
						    rs("devolucao_hora") = retorna_so_digitos(formata_hora(Now))
						    rs("devolucao_usuario") = c_devolucao_usuario
						    rs.Update 
						    s_log = s_log & log_produto_monta(.qtde_a_devolver, .fabricante, .produto)
						    if Err <> 0 then
						    '	~~~~~~~~~~~~~~~~
							    cn.RollbackTrans
						    '	~~~~~~~~~~~~~~~~
							    Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
							    end if
						
						    if Not estoque_processa_devolucao_mercadorias_v2(c_devolucao_usuario, .pedido, .fabricante, .produto, id_item_devolvido, .qtde_a_devolver, msg_erro) then
						    '	~~~~~~~~~~~~~~~~
							    cn.RollbackTrans
						    '	~~~~~~~~~~~~~~~~
							    Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
							    end if
						    end if
					    end if
				    end with
			    next

                '	ATUALIZA O STATUS DE PAGAMENTO (SE NECESSÁRIO)
	            '	==============================================
		        if alerta = "" then
		        '	OBTÉM OS VALORES A PAGAR, JÁ PAGO E O STATUS DE PAGAMENTO (PARA TODA A FAMÍLIA DE PEDIDOS)
			        if Not calcula_pagamentos(pedido_selecionado, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then 
			        '	~~~~~~~~~~~~~~~~
				        cn.RollbackTrans
			        '	~~~~~~~~~~~~~~~~
				        Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				        end if
			
			        vl_saldo_a_pagar = vl_TotalFamiliaPrecoNF-vl_TotalFamiliaPago-vl_TotalFamiliaDevolucaoPrecoNF
			
			        id_pedido_base = retorna_num_pedido_base(pedido_selecionado)
			        s = "SELECT * FROM t_PEDIDO WHERE (pedido='" & id_pedido_base & "')"
			        if rs.State <> 0 then rs.Close
			        rs.Open s, cn
			        if Err <> 0 then
			        '	~~~~~~~~~~~~~~~~
				        cn.RollbackTrans
			        '	~~~~~~~~~~~~~~~~
				        Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				        end if
			
			        if rs.Eof then
				        alerta = texto_add_br(alerta)
				        alerta = alerta & "Pedido-base " & id_pedido_base & " não foi encontrado."
			        else
				        st_pagto_a = Trim("" & rs("st_pagto"))
				        if (vl_TotalFamiliaDevolucaoPrecoNF + vl_TotalFamiliaPago) >= (vl_TotalFamiliaPrecoNF - MAX_VALOR_MARGEM_ERRO_PAGAMENTO) then
					        if Trim("" & rs("st_pagto")) <> ST_PAGTO_PAGO then
								rs("dt_st_pagto") = Date
								rs("dt_hr_st_pagto") = Now
								rs("usuario_st_pagto") = usuario
								end if
							rs("st_pagto") = ST_PAGTO_PAGO
				        elseif vl_TotalFamiliaPago > 0 then
					        if Trim("" & rs("st_pagto")) <> ST_PAGTO_PARCIAL then
								rs("dt_st_pagto") = Date
								rs("dt_hr_st_pagto") = Now
								rs("usuario_st_pagto") = usuario
								end if
							rs("st_pagto") = ST_PAGTO_PARCIAL
				        else
					        if Trim("" & rs("st_pagto")) <> ST_PAGTO_NAO_PAGO then
								rs("dt_st_pagto") = Date
								rs("dt_hr_st_pagto") = Now
								rs("usuario_st_pagto") = usuario
								end if
							rs("st_pagto") = ST_PAGTO_NAO_PAGO
					        end if

				        if st_pagto_a <> Trim("" & rs("st_pagto")) then
					        s = formata_texto_log(Lcase(x_status_pagto(st_pagto_a))) & " => " & _
						        formata_texto_log(Lcase(x_status_pagto(Trim("" & rs("st_pagto")))))
				        else
					        s = formata_texto_log(Lcase(x_status_pagto(Trim("" & rs("st_pagto")))))
					        end if
				
				        if s_log <> "" then s_log = s_log & ", "
				        s_log = s_log & "st_pagto: " & s & ", " & _
						        "valor do pedido: " & SIMBOLO_MONETARIO & " " & _
						        formata_moeda(vl_TotalFamiliaPrecoNF) & ", " & _
						        "valor pago: " & SIMBOLO_MONETARIO & " " & _
						        formata_moeda(vl_TotalFamiliaPago) & ", " & _
						        "valor das devoluções: " & SIMBOLO_MONETARIO & " " & _
						        formata_moeda(vl_TotalFamiliaDevolucaoPrecoNF)

				        rs.Update
				        if Err <> 0 then
				        '	~~~~~~~~~~~~~~~~
					        cn.RollbackTrans
				        '	~~~~~~~~~~~~~~~~
					        Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					        end if
				        end if
			        end if

                ' recupera email do usuário logado
                s_remetente = obtem_email_usuario(usuario)

                ' SE HOUVER COBRANÇA DE TAXA ADMINISTRATIVA
                if alerta = "" then
                    if c_taxa = TAXA_ADMINISTRATIVA__SIM then
                        texto_chamado = ""
                        taxa_valor_a_cobrar = converte_numero(c_taxa_valor)
                        if c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__VENDEDOR_PARCEIRO then 
                            taxa_valor_a_cobrar = converte_numero(c_taxa_valor)/2
                            end if
                        if c_taxa_forma_pagto = COD_PEDIDO_DEVOLUCAO_TAXA_FORMA_PAGTO__DESCONTO_COMISSAO then
                            if c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__PARCEIRO Or _
                                    c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__VENDEDOR_PARCEIRO then
                                if Not fin_gera_nsu(T_ORCAMENTISTA_E_INDICADOR_DESCONTO, id_orcamentista_indicador_desconto, msg_erro) then 
			                    '	~~~~~~~~~~~~~~~~
							        cn.RollbackTrans
						        '	~~~~~~~~~~~~~~~~
							        Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
				                    end if
                                s = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (id = -1)"
                                if rs.State <> 0 then rs.Close
						        rs.Open s, cn
						        rs.AddNew
                                rs("id")=id_orcamentista_indicador_desconto
                                rs("apelido")=c_pedido_indicador
                                rs("usuario")="SISTEMA"
                                rs("valor")=taxa_valor_a_cobrar
                                rs("descricao")=pedido_selecionado & " - Cobrança de taxa administrativa ref. devolução de mercadoria."
                                rs.Update
                                s_log_indicador_desconto = _
                                    "Indicador=" & c_pedido_indicador & "; " & _
                                    "Valor=" & taxa_valor_a_cobrar & "; " & _
                                    "Descrição=Cobrança de taxa administrativa ref. devolução de mercadoria"
                                if Err <> 0 then
			                    '	~~~~~~~~~~~~~~~~
				                    cn.RollbackTrans
			                    '	~~~~~~~~~~~~~~~~
				                    Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				                    end if
                                    ' grava log do registro de desconto no cadastro do parceiro
                                    if s_log_indicador_desconto <> "" then
                                        s_log_indicador_desconto = "Registro na tabela de desconto do indicador: " & s_log_indicador_desconto
			                            grava_log c_devolucao_usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_DEVOLUCAO, s_log_indicador_desconto
                                        end if
                                end if
                            if c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__VENDEDOR_PARCEIRO Or _
                                    c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__VENDEDOR then
                                cod_motivo_abertura_chamado = "1020"
                                texto_chamado = "Desconto de comissão do vendedor " & c_vendedor & " no valor de " & SIMBOLO_MONETARIO & formata_moeda(taxa_valor_a_cobrar) & " referente cobrança de taxa administrativa em devolução de mercadoria."
                            elseif c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__PARCEIRO then
								cod_motivo_abertura_chamado = "1026"
								texto_chamado = "Desconto de comissão do parceiro " & c_pedido_indicador & " no valor de " & SIMBOLO_MONETARIO & formata_moeda(taxa_valor_a_cobrar) & " referente cobrança de taxa administrativa em devolução de mercadoria."
								end if
                        elseif c_taxa_forma_pagto = COD_PEDIDO_DEVOLUCAO_TAXA_FORMA_PAGTO__DEPOSITO_BANCARIO then
                                cod_motivo_abertura_chamado = "1021"
                                if c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__VENDEDOR then
                                    texto_chamado = "Devolução de mercadoria - Receber depósito bancário no valor de " & SIMBOLO_MONETARIO & formata_moeda(taxa_valor_a_cobrar) & " referente cobrança de taxa administrativa sob responsabilidade do vendedor " & c_vendedor & "."
                                elseif c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__VENDEDOR_PARCEIRO then
                                    texto_chamado = "Devolução de mercadoria - Receber depósito bancário no valor de " & SIMBOLO_MONETARIO & formata_moeda(converte_numero(c_taxa_valor)) & " divido (" & SIMBOLO_MONETARIO & formata_moeda(taxa_valor_a_cobrar) & " para cada) entre o parceiro (" & c_pedido_indicador & ") e vendedor (" & c_vendedor & ") referente cobrança de taxa administrativa."
                                elseif c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__CLIENTE then
                                    texto_chamado = "Devolução de mercadoria - Receber depósito bancário no valor de " & SIMBOLO_MONETARIO & formata_moeda(taxa_valor_a_cobrar) & " referente cobrança de taxa administrativa sob responsabilidade do cliente vinculado ao pedido."
                                elseif c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__PARCEIRO then
                                    texto_chamado = "Devolução de mercadoria - Receber depósito bancário no valor de " & SIMBOLO_MONETARIO & formata_moeda(taxa_valor_a_cobrar) & " referente cobrança de taxa administrativa sob responsabilidade do parceiro (" & c_pedido_indicador & ")."                                    
                                    end if
                        elseif c_taxa_forma_pagto = COD_PEDIDO_DEVOLUCAO_TAXA_FORMA_PAGTO__ABATIMENTO_CREDITO then
                                cod_motivo_abertura_chamado = "1025"
                                if c_taxa_responsavel = COD_PEDIDO_DEVOLUCAO_TAXA_RESPONSAVEL__CLIENTE then
                                    texto_chamado = "Devolução de mercadoria - Abater o valor da taxa administrativa no crédito obtido pelo cliente."
                                    end if
                            end if
                        ' ABRE CHAMADO PARA O DEPTO FINANCEIRO INFORMANDO A COBRANÇA DE TAXA ADMINISTRATIVA
                        if texto_chamado <> "" then
                            if c_taxa_observacoes <> "" then
                                texto_chamado = texto_chamado & chr(13) & chr(13) & "Observações:" & chr(13) & c_taxa_observacoes
                                end if
                            if Not fin_gera_nsu(T_PEDIDO_CHAMADO, id_chamado, msg_erro) then 
			                '	~~~~~~~~~~~~~~~~
							    cn.RollbackTrans
						    '	~~~~~~~~~~~~~~~~
							    Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
				                end if
                            s = "SELECT * FROM t_PEDIDO_CHAMADO WHERE (id = -1)"
                            if rs.State <> 0 then rs.Close
                            rs.Open s, cn
			                rs.AddNew
                            rs("id")=id_chamado
                            rs("pedido")=pedido_selecionado
                            rs("usuario_cadastro")=usuario
                            rs("cod_motivo_abertura")=cod_motivo_abertura_chamado
                            rs("texto_chamado")=texto_chamado
                            rs("nivel_acesso")=COD_NIVEL_ACESSO_CHAMADO_PEDIDO__RESTRITO
                            rs("id_depto")=id_pedido_chamado_depto
                            rs("contato")=usuario
                            rs.Update
                            if Err <> 0 then
			                '	~~~~~~~~~~~~~~~~
				                cn.RollbackTrans
			                '	~~~~~~~~~~~~~~~~
				                Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				                end if
                            ' envia email para o usuário responsável pelo departamento
                            s_destinatario = ""
                            s = "SELECT usuario_responsavel," & _
                                    " descricao AS descricao_depto," & _
                                    " Coalesce((SELECT email FROM t_USUARIO WHERE usuario=tPCD.usuario_responsavel), '') AS email" & _ 
                                " FROM t_PEDIDO_CHAMADO_DEPTO tPCD" & _
                                " WHERE (id = " & Cstr(id_pedido_chamado_depto) & ")"
                            if rs.State <> 0 then rs.Close
                            rs.Open s, cn
                            if Not rs.Eof then
                                corpo_mensagem = "Foi aberto um novo chamado para o pedido " & pedido_selecionado & " destinado ao seguinte departamento: " & UCase(rs("descricao_depto")) & "." & chr(13) & chr(10)
                                corpo_mensagem = corpo_mensagem & "Aberto por: " & usuario & " - " & x_usuario(usuario) & " em " & Now & "." & chr(13) & chr(10)
                                corpo_mensagem = corpo_mensagem & "Nível de acesso: " & nivel_acesso_chamado_pedido_descricao(COD_NIVEL_ACESSO_CHAMADO_PEDIDO__RESTRITO) & "." & chr(13) & chr(10)
                                corpo_mensagem = corpo_mensagem & "Motivo da abertura: " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA, cod_motivo_abertura_chamado) & "." & chr(13) & chr(10) & chr(13) & chr(10)
                                corpo_mensagem = corpo_mensagem & "Descrição: " & chr(13) & chr(10) & texto_chamado & chr(13) & chr(10) & chr(13) & chr(10)

                                if s_remetente = "" then
                                    corpo_mensagem = corpo_mensagem & "---------------------------------------------------------------------------------------------------------"   & chr(13) & chr(10)
                                    corpo_mensagem = corpo_mensagem & "E-MAIL ENVIADO AUTOMATICAMENTE PELO SISTEMA. NÃO RESPONDA ESTE E-MAIL, POIS ESTA CONTA NÃO É MONITORADA!!"                    
                                    end if

                                s_destinatario = Trim("" & rs("email"))

                                if s_destinatario <> "" then
                                EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__CHAMADOS_EM_PEDIDOS), _
                                                                    s_remetente, _
                                                                    s_destinatario, _
                                                                    "", _
                                                                    "", _
                                                                    "Nova abertura de chamado para o pedido " & pedido_selecionado, _
                                                                    corpo_mensagem, _
                                                                    Now, _
                                                                    id_email, _
                                                                    msg_erro_grava_email
                                    end if
                                end if

                            end if

                        if Err <> 0 then
                        '	~~~~~~~~~~~~~~~~
				            cn.RollbackTrans
			            '	~~~~~~~~~~~~~~~~
                            alerta=Cstr(Err) & ": " & Err.Description
                            end if

                        end if ' c_taxa = TAXA_ADMINISTRATIVA__SIM
                    end if ' alerta = ""
                
                ' CRÉDITO
                if alerta = "" then
                    if c_credito_transacao <> "" then
                        texto_chamado = ""
                        if c_credito_transacao = CREDITO_TRANSACAO__TRANSFERENCIA then
                            cod_motivo_abertura_chamado = "1022"
                            texto_chamado = "Devolução de mercadoria - Providenciar a transferência do crédito para o novo pedido " & c_pedido_novo & "." 
                        elseif c_credito_transacao = CREDITO_TRANSACAO__ESTORNO then
                            cod_motivo_abertura_chamado = "1024"
                            texto_chamado = "Devolução de mercadoria - Providenciar o estorno no cartão de crédito do cliente."
                        elseif c_credito_transacao = CREDITO_TRANSACAO__REEMBOLSO then
                            cod_motivo_abertura_chamado = "1023"
                            texto_chamado = "Devolução de mercadoria - Providenciar o reembolso através de depósito para a seguinte conta bancária:" & chr(13) & _
                                        "Banco: " & c_cliente_banco & " - " & x_banco(c_cliente_banco) & chr(13) & _
                                        "Agência: " & c_cliente_agencia & chr(13) & _
                                        "Conta: " & c_cliente_conta & chr(13) & _
                                        "Favorecido: " & c_cliente_favorecido
                            end if

                        ' ABRE O CHAMADO DESTINADO AO DEPTO FINANCEIRO INFORMANDO SOBRE O CRÉDITO
                        if c_credito_observacoes <> "" then
                                texto_chamado = texto_chamado & chr(13) & chr(13) & "Observações:" & chr(13) & c_credito_observacoes
                                end if
                        if Not fin_gera_nsu(T_PEDIDO_CHAMADO, id_chamado, msg_erro) then 
			                '	~~~~~~~~~~~~~~~~
							    cn.RollbackTrans
						    '	~~~~~~~~~~~~~~~~
							    Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
				                end if
                            s = "SELECT * FROM t_PEDIDO_CHAMADO WHERE (id = -1)"
                            if rs.State <> 0 then rs.Close
                            rs.Open s, cn
			                rs.AddNew
                            rs("id")=id_chamado
                            rs("pedido")=pedido_selecionado
                            rs("usuario_cadastro")=usuario
                            rs("cod_motivo_abertura")=cod_motivo_abertura_chamado
                            rs("texto_chamado")=texto_chamado 
                            rs("nivel_acesso")=COD_NIVEL_ACESSO_CHAMADO_PEDIDO__RESTRITO
                            rs("id_depto")=id_pedido_chamado_depto
                            rs("contato")=usuario
                            rs.Update
                            if Err <> 0 then
			                '	~~~~~~~~~~~~~~~~
				                cn.RollbackTrans
			                '	~~~~~~~~~~~~~~~~
				                Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				                end if
                            ' envia email para o usuário responsável pelo departamento
                            s_destinatario = ""
                            s = "SELECT usuario_responsavel," & _
                                    " descricao AS descricao_depto," & _
                                    " Coalesce((SELECT email FROM t_USUARIO WHERE usuario=tPCD.usuario_responsavel), '') AS email" & _ 
                                " FROM t_PEDIDO_CHAMADO_DEPTO tPCD" & _
                                " WHERE (id = " & Cstr(id_pedido_chamado_depto) & ")"
                            if rs.State <> 0 then rs.Close
                            rs.Open s, cn
                            if Not rs.Eof then
                                corpo_mensagem = "Foi aberto um novo chamado para o pedido " & pedido_selecionado & " destinado ao seguinte departamento: " & UCase(rs("descricao_depto")) & "." & chr(13) & chr(10)
                                corpo_mensagem = corpo_mensagem & "Aberto por: " & usuario & " - " & x_usuario(usuario) & " em " & Now & "." & chr(13) & chr(10)
                                corpo_mensagem = corpo_mensagem & "Nível de acesso: " & nivel_acesso_chamado_pedido_descricao(COD_NIVEL_ACESSO_CHAMADO_PEDIDO__RESTRITO) & "." & chr(13) & chr(10)
                                corpo_mensagem = corpo_mensagem & "Motivo da abertura: " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA, cod_motivo_abertura_chamado) & "." & chr(13) & chr(10) & chr(13) & chr(10)
                                corpo_mensagem = corpo_mensagem & "Descrição: " & chr(13) & chr(10) & texto_chamado & chr(13) & chr(10) & chr(13) & chr(10)

                                if s_remetente = "" then
                                    corpo_mensagem = corpo_mensagem & "---------------------------------------------------------------------------------------------------------"   & chr(13) & chr(10)
                                    corpo_mensagem = corpo_mensagem & "E-MAIL ENVIADO AUTOMATICAMENTE PELO SISTEMA. NÃO RESPONDA ESTE E-MAIL, POIS ESTA CONTA NÃO É MONITORADA!!"                    
                                    end if

                                s_destinatario = Trim("" & rs("email"))

                                if s_destinatario <> "" then
                                EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__CHAMADOS_EM_PEDIDOS), _
                                                                    s_remetente, _
                                                                    s_destinatario, _
                                                                    "", _
                                                                    "", _
                                                                    "Nova abertura de chamado para o pedido " & pedido_selecionado, _
                                                                    corpo_mensagem, _
                                                                    Now, _
                                                                    id_email, _
                                                                    msg_erro_grava_email
                                    end if
                                end if

                        end if
                    end if ' alerta = ""

                end if ' rb_status = COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO
            end if 'rb_status <> ""

    '   GRAVA OS ARQUIVOS DE FOTO
    if alerta = "" then
        if upload_file_guid <> "" then
            v_upload_file_guid = Split(upload_file_guid, "|")
            for each guid_item in v_upload_file_guid
                ' recupera ID através do Guid
                if guid_item <> "" then
                    s = "SELECT id FROM t_UPLOAD_FILE WHERE (Convert(varchar(36), guid) = '" & guid_item & "')"
                    if rs.State <> 0 then rs.Close
			        rs.Open s, cn
                    if Not rs.Eof then
                        id_upload_file = rs("id")
                        s = "SELECT * FROM t_PEDIDO_DEVOLUCAO_IMAGEM WHERE (id_pedido_devolucao=-1)"
                        if rs.State <> 0 then rs.Close
			            rs.Open s, cn
			            rs.AddNew
                        rs("id_pedido_devolucao") = id_devolucao
                        rs("id_upload_file") = id_upload_file
                        rs.Update
                        if Err <> 0 then
				        '	~~~~~~~~~~~~~~~~
					        cn.RollbackTrans
				        '	~~~~~~~~~~~~~~~~
					        alerta = "Falha ao registrar o arquivo de foto."
                            exit for
					        end if
                        end if
                    end if
                next
            if Err=0 then
                for each guid_item in v_upload_file_guid
                    if guid_item <> "" then
                        s = "SELECT * FROM t_UPLOAD_FILE WHERE (Convert(varchar(36), guid) = '" & guid_item & "')"
                        if rs.State <> 0 then rs.Close
			            rs.Open s, cn
                        if Not rs.Eof then
                            rs("st_confirmation_ok")=1
                            rs.Update
                            if Err <> 0 then
				            '	~~~~~~~~~~~~~~~~
					            cn.RollbackTrans
				            '	~~~~~~~~~~~~~~~~
					            alerta = "Falha ao registrar o arquivo de foto."
                                exit for
					            end if    
                            end if 
                        end if
                    next
                end if
            end if
        end if
			
			
	'	GRAVA O LOG E CONCLUI A TRANSAÇÃO
	'	=================================
		if alerta = "" then
			s_log = "Atualizar pré-devolução:" & s_log
			grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_DEVOLUCAO, s_log
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				s = "Pré-devolução atualizada com sucesso!!"
				Session(SESSION_CLIPBOARD) = s
				Response.Redirect("PedidoPreDevolucaoAtualizaMensagem.asp" & "?url_back=" & url_back & "&rb_status=" & rb_status_back & "&c_loja=" & c_loja & "&origem=A&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			end if
		
		if sx.State <> 0 then sx.Close
		set sx = nothing
		if rs.State <> 0 then rs.Close
		set rs = nothing
        if r.State <> 0 then r.Close
		set r = nothing
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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>



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

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>
<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
