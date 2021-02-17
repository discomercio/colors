<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===================================================================
'	  P E D I D O P R E D E V O L U C A O N O V A C O N F I R M A . A S P
'     ===================================================================
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
	if Not operacao_permitida(OP_CEN_PRE_DEVOLUCAO_CADASTRAMENTO, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, sx
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, vl_saldo_a_pagar, st_pagto, st_pagto_a
	dim iv, j, k, n, v_devol, alerta, deve_devolver, s_log, msg_erro, id_item_devolvido, id_pedido_devolucao
    dim c_procedimento, c_local_coleta, c_coleta_endereco, c_coleta_endereco_numero, c_coleta_bairro, c_coleta_cep, c_coleta_cidade, c_coleta_uf, c_coleta_complemento
    dim c_motivo_devolucao, c_motivo_descricao, c_taxa, c_taxa_forma_pagto, c_taxa_percentual, c_taxa_valor, c_taxa_responsavel, c_credito_transacao, c_pedido_novo, c_observacoes
    dim c_cliente_banco, c_cliente_agencia, c_cliente_conta, c_cliente_favorecido, c_total_devolucao, c_pedido_cliente, c_pedido_indicador, upload_file_guid, v_upload_file_guid, guid_item, id_upload_file
    dim corpo_mensagem, id_email, msg_erro_grava_email
    dim c_pedido_possui_parcela_cartao
    dim c_taxa_observacoes, c_credito_observacoes
    dim r_pedido
	
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
    c_pedido_indicador = Trim(Request.Form("c_pedido_indicador"))
    upload_file_guid = Trim(Request.Form("upload_file_guid_returned"))
    c_pedido_possui_parcela_cartao = Request.Form("c_pedido_possui_parcela_cartao")
    c_taxa_observacoes = Trim(Request.Form("c_taxa_observacoes"))
    c_credito_observacoes = Trim(Request.Form("c_credito_observacoes"))
    
    redim v_devol(0)
	set v_devol(Ubound(v_devol)) = New cl_ITEM_DEVOLUCAO_MERCADORIAS
	v_devol(Ubound(v_devol)).produto = ""
	v_devol(Ubound(v_devol)).qtde_a_devolver = 0
	v_devol(Ubound(v_devol)).motivo = ""
    v_devol(Ubound(v_devol)).vl_unitario = 0

	s_log = ""
	alerta = ""
	deve_devolver = False

    if c_procedimento = "" then
        alerta = "Procedimento não foi informado."
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
    if c_motivo_descricao = "" then
        alerta = texto_add_br(alerta)
		alerta = alerta & "Não foi inserida uma descrição para o motivo da devolução."
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
				
				    s = Trim(Request.Form("c_qtde")(iv))
				    if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
				
				    s = Trim(Request.Form("c_devolucao_anterior")(iv))
				    .qtde_devolvida_anteriormente = converte_numero(s) 

                    s = Trim(Request.Form("c_devolucao_pendente")(iv))
                    .qtde_devolucao_pendente = converte_numero(s)
				
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
	    else
            if Not estoque_verifica_mercadorias_para_devolucao(v_devol, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)            
		    for iv = Lbound(v_devol) to Ubound(v_devol)
			    with v_devol(iv)
				    if .produto <> "" then
					    if .qtde_a_devolver > (.qtde - (.qtde_devolvida_anteriormente + .qtde_devolucao_pendente)) then
						    alerta = texto_add_br(alerta)
						    alerta = alerta & "Produto " & .produto & " do fabricante " & .fabricante & " especifica quantidade inválida para devolução."
						    end if
					
					    end if
				    end with
			    next
		    end if
        end if ' alerta

    if alerta = "" then
        if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then
            alerta = msg_erro
            end if
        end if

	dim r_usuario
	if alerta = "" then
		'Obtém maiores informações sobre o usuário que está realizando o cadastro da pré-devolução para montar a mensagem do e-mail de aviso
		call le_usuario(usuario, r_usuario, msg_erro)
		end if

	dim r_vendedor
	if alerta = "" then
		'Se o usuário que estiver cadastrando a pré-devolução não for o vendedor do pedido, o vendedor irá receber um email de aviso sobre a pré-devolução
		if UCase(Trim("" & r_pedido.vendedor)) <> UCase(usuario) then
			call le_usuario(Trim("" & r_pedido.vendedor), r_vendedor, msg_erro)
			end if
		end if

	dim dtHrMensagem
	dim s_dados_produtos_devolucao
	s_dados_produtos_devolucao = ""

	dim s_dados_cliente
	dim r_cliente
	if alerta = "" then
		set r_cliente = New cl_CLIENTE
		call x_cliente_bd(r_pedido.id_cliente, r_cliente)

		if r_pedido.st_memorizacao_completa_enderecos <> 0 then
			s_dados_cliente = "Cliente: " & r_pedido.endereco_nome_iniciais_em_maiusculas & " (" & cnpj_cpf_formata(r_pedido.endereco_cnpj_cpf) & ")"
		else
			s_dados_cliente = "Cliente: " & r_cliente.nome_iniciais_em_maiusculas & " (" & cnpj_cpf_formata(r_cliente.cnpj_cpf) & ")"
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

        if Not gera_nsu(T_PEDIDO_DEVOLUCAO, id_pedido_devolucao, msg_erro) then 
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
			end if

        s = "SELECT * FROM t_PEDIDO_DEVOLUCAO WHERE (id=-1)"
        if rs.State <> 0 then rs.Close
        rs.Open s, cn

        if rs.Eof then
            rs.AddNew
            rs("id") = id_pedido_devolucao
            rs("pedido") = pedido_selecionado
            rs("usuario_cadastro") = usuario
            rs("status") = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA
            rs("status_usuario") = usuario
            rs("status_data") = Date
            rs("status_data_hora") = Now
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
            rs("taxa_flag") = c_taxa
            if c_taxa = TAXA_ADMINISTRATIVA__SIM then
                rs("taxa_percentual") = converte_numero(c_taxa_percentual)
                rs("cod_taxa_forma_pagto") = c_taxa_forma_pagto
                rs("cod_taxa_responsavel") = c_taxa_responsavel
                end if
            rs("vl_devolucao") = converte_numero(c_total_devolucao)
            rs("cod_credito_transacao") = c_credito_transacao
            if c_credito_transacao = CREDITO_TRANSACAO__REEMBOLSO then
                rs("banco") = c_cliente_banco
                rs("agencia") = c_cliente_agencia
                rs("conta") = c_cliente_conta
                rs("favorecido") = c_cliente_favorecido
                end if
            if c_pedido_novo <> "" then rs("pedido_novo") = c_pedido_novo
            if c_taxa_observacoes <> "" then rs("taxa_observacoes") = c_taxa_observacoes
            if c_credito_observacoes <> "" then rs("credito_observacoes") = c_credito_observacoes
            rs("observacao") = c_observacoes
            rs.Update
            if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
            end if

		for iv = Lbound(v_devol) to Ubound(v_devol)
			with v_devol(iv)
				if (Trim(.produto)<>"") And (.qtde_a_devolver > 0) then
				'	REGISTRA O PRODUTO A SER DEVOLVIDO
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
						s = "SELECT * FROM t_PEDIDO_DEVOLUCAO_ITEM WHERE (produto='XX')"
						if rs.State <> 0 then rs.Close
						rs.Open s, cn
						rs.AddNew
                        rs("id_pedido_devolucao") = id_pedido_devolucao
						rs("qtde") = .qtde_a_devolver
						rs("fabricante") = .fabricante
                        rs("produto") = .produto
                        rs("vl_unitario") = .vl_unitario
						rs.Update
						s_log = s_log & log_produto_monta(.qtde_a_devolver, .fabricante, .produto)
						if s_dados_produtos_devolucao <> "" then s_dados_produtos_devolucao = s_dados_produtos_devolucao & vbCrLf
						s_dados_produtos_devolucao = s_dados_produtos_devolucao & .qtde_a_devolver & " x " & .produto & "(" & .fabricante & ")"
						if Err <> 0 then
						'	~~~~~~~~~~~~~~~~
							cn.RollbackTrans
						'	~~~~~~~~~~~~~~~~
							Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
							end if
						end if
					end if
				end with
			next
			
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
                        rs("id_pedido_devolucao") = id_pedido_devolucao
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
			dtHrMensagem = Now
            ' envia e-mail para o operador das devoluções
            if isLojaBonshop(r_pedido.loja) Or isLojaVrf(r_pedido.loja) then
                dim strEmailAdministrador
                set strEmailAdministrador = get_registro_t_parametro("PEDIDO_DEVOLUCAO_EMAIL_ADMINISTRADOR")
                if Trim("" & strEmailAdministrador.campo_texto) <> "" then
					corpo_mensagem = "Usuário '" & usuario & "' (" & r_usuario.nome_iniciais_em_maiusculas & ") cadastrou uma nova pré-devolução no pedido " & pedido_selecionado & " em " & formata_data_hora_sem_seg(dtHrMensagem) & _
									vbCrLf & _
									"Pedido: " & pedido_selecionado & _
									vbCrLf & _
									"Devolução nº " & RetiraZerosAEsquerda(Cstr(id_pedido_devolucao)) & _
									vbCrLf & _
									s_dados_cliente & _
									vbCrLf & vbCrLf & _
									String(30, "-") & "( Início )" & String(30, "-") & _
									vbCrLf & _
									s_dados_produtos_devolucao & _
									vbCrLf & _
									String(31, "-") & "( Fim )" & String(32, "-") & _
									vbCrLf & vbCrLf & _
									"Atenção: esta é uma mensagem automática, NÃO responda a este e-mail!"

                    EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__PEDIDO_DEVOLUCAO), _
                                                    "", _
                                                    strEmailAdministrador.campo_texto, _
                                                    "", _
                                                    "", _
                                                    "Nova pré-devolução cadastrada no pedido " & pedido_selecionado, _
                                                    corpo_mensagem, _
                                                    Now, _
                                                    id_email, _
                                                    msg_erro_grava_email
                    end if 'if Trim("" & strEmailAdministrador.campo_texto) <> ""
                
				if UCase(usuario) <> UCase(r_pedido.vendedor) then
					if Trim("" & r_vendedor.email) <> "" then
						'Envia email para o vendedor
						corpo_mensagem = "Usuário '" & usuario & "' (" & r_usuario.nome_iniciais_em_maiusculas & ") cadastrou uma nova pré-devolução no pedido " & pedido_selecionado & " em " & formata_data_hora_sem_seg(dtHrMensagem) & _
										vbCrLf & _
										"Pedido: " & pedido_selecionado & _
										vbCrLf & _
										"Devolução nº " & RetiraZerosAEsquerda(Cstr(id_pedido_devolucao)) & _
										vbCrLf & _
										s_dados_cliente & _
										vbCrLf & vbCrLf & _
										String(30, "-") & "( Início )" & String(30, "-") & _
										vbCrLf & _
										s_dados_produtos_devolucao & _
										vbCrLf & _
										String(31, "-") & "( Fim )" & String(32, "-") & _
										vbCrLf & vbCrLf & _
										"Atenção: esta é uma mensagem automática, NÃO responda a este e-mail!"

						EmailSndSvcGravaMensagemParaEnvio getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__PEDIDO_DEVOLUCAO), _
														"", _
														Trim("" & r_vendedor.email), _
														"", _
														"", _
														"Nova pré-devolução cadastrada no pedido " & pedido_selecionado, _
														corpo_mensagem, _
														Now, _
														id_email, _
														msg_erro_grava_email
						end if 'if Trim("" & r_vendedor.email) <> ""
					end if 'if UCase(usuario) <> UCase(r_pedido.vendedor)
				end if 'if isLojaBonshop(r_pedido.loja) Or isLojaVrf(r_pedido.loja)

			s_log = "Cadastro de pré-devolução:" & s_log
			grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_DEVOLUCAO, s_log
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				s = "Pré-devolução cadastrada com sucesso!!"
				Session(SESSION_CLIPBOARD) = s
				Response.Redirect("mensagem.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else 'if alerta = ""
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			end if 'if alerta = ""
		
		if sx.State <> 0 then sx.Close
		set sx = nothing
		if rs.State <> 0 then rs.Close
		set rs = nothing
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
	<title>CENTRAL</title>
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
