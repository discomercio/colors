<%
' =========================================
'          E S T O Q U E
' =========================================


' ---------------------------------------------------------------
'   LOG_ESTOQUE_MONTA_TRANSFERENCIA
function log_estoque_monta_transferencia(byval quantidade, byval id_fabricante, byval id_produto)
dim s    
    s = " " & CStr(quantidade) & "x" & Trim(id_produto)
    if Trim(id_fabricante) <> "" then s = s & "(" & Trim(id_fabricante) & ")"
    log_estoque_monta_transferencia = s
end function



' ---------------------------------------------------------------
'   LOG_ESTOQUE_MONTA_DECREMENTO
function log_estoque_monta_decremento(byval quantidade, byval id_fabricante, byval id_produto)
dim s    
    s = " -" & CStr(quantidade) & "x" & Trim(id_produto)
    if Trim(id_fabricante) <> "" then s = s & "(" & Trim(id_fabricante) & ")"
    log_estoque_monta_decremento = s
end function



' ---------------------------------------------------------------
'   LOG_ESTOQUE_MONTA_INCREMENTO
function log_estoque_monta_incremento(byval quantidade, byval id_fabricante, byval id_produto)
dim s    
    s = " +" & CStr(quantidade) & "x" & Trim(id_produto)
    if Trim(id_fabricante) <> "" then s = s & "(" & Trim(id_fabricante) & ")"
    log_estoque_monta_incremento = s
end function



' ---------------------------------------------------------------
'   ESTOQUE_VERIFICA_DISPONIBILIDADE_INTEGRAL
'   Retorno da função:
'      False - Ocorreu falha ao tentar verificar o estoque.
'      True - Conseguiu fazer a verificação do estoque.
'   Esta função consulta o banco de dados para verificar se
'   existem produtos suficientes no "estoque de venda" para
'   atender ao pedido.
'   Note que os produtos a serem analisados são informados
'   através do vetor do parâmetro v_item.
'   Esta rotina é a original (antes da implementação do auto-split) e foi
'   mantida p/ ser usada na operação de cadastramento de orçamento.
function estoque_verifica_disponibilidade_integral(byref v_item, byref erro_produto_indisponivel)
dim i
dim s
dim rs
	estoque_verifica_disponibilidade_integral=False
	erro_produto_indisponivel=False
	for i=LBound(v_item) to Ubound(v_item)
		with v_item(i)
            .qtde_estoque=0
			if (.qtde_solicitada > 0) And (Trim(.produto)<>"") then
                s = "SELECT Sum(qtde - qtde_utilizada) AS saldo" & _
                    " FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
                    " WHERE (t_ESTOQUE.fabricante='" & .fabricante & "') AND (produto='" & .produto & "') AND ((qtde-qtde_utilizada)>0)"
				set rs=cn.Execute(s)
				if Err<>0 then exit function
				if Not rs.Eof then
					if Not IsNull(rs("saldo")) then if IsNumeric(rs("saldo")) then .qtde_estoque=CLng(rs("saldo"))
					end if
				if .qtde_solicitada > .qtde_estoque then
					erro_produto_indisponivel=True
					end if
				if rs.State <> 0 then rs.Close
				if Err<>0 then exit function
				end if
			end with
		next	
	estoque_verifica_disponibilidade_integral=True
end function



' ---------------------------------------------------------------
'   ESTOQUE_VERIFICA_DISPONIBILIDADE_INTEGRAL_V2
'   Retorno da função:
'      False - Ocorreu falha ao tentar verificar o estoque.
'      True - Conseguiu fazer a verificação do estoque.
'   Esta função consulta o banco de dados para verificar se
'   existem produtos suficientes no "estoque de venda" para
'   atender ao pedido.
'   Note que os produtos a serem analisados são informados
'   através do parâmetro 'item', que é um objeto da
'   classe cl_CTRL_ESTOQUE_PEDIDO_ITEM_NOVO
function estoque_verifica_disponibilidade_integral_v2(ByVal id_nfe_emitente, byref item)
dim s
dim rs
	estoque_verifica_disponibilidade_integral_v2=False
	with item
		.qtde_estoque=0
		if (.qtde_solicitada > 0) And (Trim(.produto)<>"") then
			'Calcula quantidade em estoque no CD especificado
			s = "SELECT" & _
					" Sum(qtde - qtde_utilizada) AS saldo" & _
				" FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
				" WHERE" & _
					" (t_ESTOQUE.id_nfe_emitente = " & Trim("" & id_nfe_emitente) & ")" & _
					" AND (t_ESTOQUE.fabricante='" & .fabricante & "')" & _
					" AND (produto='" & .produto & "')" & _
					" AND ((qtde-qtde_utilizada) > 0)"
			set rs=cn.Execute(s)
			if Err<>0 then exit function
			if Not rs.Eof then
				if Not IsNull(rs("saldo")) then if IsNumeric(rs("saldo")) then .qtde_estoque=CLng(rs("saldo"))
				end if
			if rs.State <> 0 then rs.Close
			if Err<>0 then exit function
			
			'Calcula quantidade em estoque global (quantidade total disponível em todos os CD's)
			s = "SELECT" & _
					" Sum(qtde - qtde_utilizada) AS saldo" & _
				" FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
				" WHERE" & _
					" (t_ESTOQUE.fabricante='" & .fabricante & "')" & _
					" AND (produto='" & .produto & "')" & _
					" AND ((qtde-qtde_utilizada) > 0)" & _
					" AND (" & _
						"(t_ESTOQUE.id_nfe_emitente = " & Trim("" & id_nfe_emitente) & ")" & _
						" OR " & _
						"(" & _
							"t_ESTOQUE.id_nfe_emitente IN " & _
							"(SELECT id FROM t_NFe_EMITENTE WHERE (st_habilitado_ctrl_estoque = 1) AND (st_ativo = 1))" & _
						")" & _
					")"
			set rs=cn.Execute(s)
			if Err<>0 then exit function
			if Not rs.Eof then
				if Not IsNull(rs("saldo")) then if IsNumeric(rs("saldo")) then .qtde_estoque_global=CLng(rs("saldo"))
				end if
			if rs.State <> 0 then rs.Close
			if Err<>0 then exit function
			end if
		end with
	estoque_verifica_disponibilidade_integral_v2=True
end function



' --------------------------------------------------------------------
'   ESTOQUE_PRODUTO_SAIDA_V2
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros entre as várias tabelas.
'   Esta função processa a saída dos produtos do "estoque de venda"
'   para o "estoque vendido".  No caso de não haver produtos sufi-
'   cientes no "estoque de venda" e desde que esteja autorizado
'   através do parâmetro "qtde_autorizada_sem_presenca", os produtos
'   que faltam são colocados automaticamente na lista de produtos
'   vendidos sem presença no estoque.
function estoque_produto_saida_v2(byval id_usuario, byval id_pedido, _
								byval id_nfe_emitente, byval id_fabricante, byval id_produto, _
								byval qtde_a_sair, byval qtde_autorizada_sem_presenca, _
								byref qtde_estoque_vendido, byref qtde_estoque_sem_presenca, _
								byref msg_erro)
dim s_sql
dim s_chave
dim qtde_disponivel
Dim v_estoque()
Dim iv
dim rs
Dim qtde_aux
Dim qtde_utilizada_aux
Dim qtde_movto
Dim qtde_movimentada

	estoque_produto_saida_v2=False

	msg_erro=""
	qtde_estoque_vendido=0
	qtde_estoque_sem_presenca=0
	
    If (qtde_a_sair <= 0) Or (Trim(id_produto) = "") Or (Trim(id_pedido) = "") Then
        estoque_produto_saida_v2 = True
        exit function
        end if

'	OBTÉM OS "LOTES" DO PRODUTO DISPONÍVEIS NO ESTOQUE (POLÍTICA FIFO)
	s_sql = "SELECT" & _
				" t_ESTOQUE.id_estoque," & _
				" (qtde - qtde_utilizada) AS saldo" & _
			" FROM t_ESTOQUE" & _
				" INNER JOIN t_ESTOQUE_ITEM ON" & " (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
			" WHERE" & _
				" (t_ESTOQUE.id_nfe_emitente = " & Trim("" & id_nfe_emitente) & ")" & _
				" AND (t_ESTOQUE_ITEM.fabricante='" & id_fabricante & "')" & _
				" AND (produto='" & id_produto & "')" & _
				" AND ((qtde - qtde_utilizada) > 0)" & _
			" ORDER BY" & _
				" data_entrada," & _
				" t_ESTOQUE.id_estoque"

    ReDim v_estoque(0)
    v_estoque(UBound(v_estoque)) = ""

    set rs=cn.Execute(s_sql)
    if Err<>0 then 
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
    qtde_disponivel = 0
    do while Not rs.Eof
	'	ARMAZENA AS ENTRADAS NO ESTOQUE CANDIDATAS À SAÍDA DE PRODUTOS
		If v_estoque(UBound(v_estoque)) <> "" Then
          ReDim Preserve v_estoque(UBound(v_estoque) + 1)
          v_estoque(UBound(v_estoque)) = ""
          End If
      v_estoque(UBound(v_estoque)) = Trim("" & rs("id_estoque"))
      qtde_disponivel = qtde_disponivel + CLng(rs("saldo"))
      rs.MoveNext 
      Loop

'	NÃO HÁ PRODUTOS SUFICIENTES NO ESTOQUE!!
    if (qtde_a_sair-qtde_autorizada_sem_presenca) > qtde_disponivel then 
		msg_erro="Produto " & id_produto & " do fabricante " & id_fabricante & ": faltam " & _
                 Cstr((qtde_a_sair-qtde_autorizada_sem_presenca)-qtde_disponivel) & " unidades no estoque (" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente) & ") para poder atender ao pedido."
		exit function
		end if

	if rs.State <> 0 then rs.Close
	set rs=nothing
	
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

'	REALIZA A SAÍDA DO ESTOQUE!!
    qtde_movimentada = 0
    For iv = LBound(v_estoque) To UBound(v_estoque)
        If Trim(v_estoque(iv)) <> "" Then
          ' A QUANTIDADE NECESSÁRIA JÁ FOI RETIRADA DO ESTOQUE!!
            If qtde_movimentada >= qtde_a_sair Then Exit For

          ' T_ESTOQUE_ITEM: SAÍDA DE PRODUTOS
            s_sql = "SELECT qtde, qtde_utilizada, data_ult_movimento FROM t_ESTOQUE_ITEM WHERE" & _
                    " (id_estoque = '" & Trim(v_estoque(iv)) & "')" & _
                    " AND (fabricante = '" & id_fabricante & "')" & _
                    " AND (produto = '" & id_produto & "')"
			qtde_movto=0
			qtde_aux = 0
			qtde_utilizada_aux = 0
			rs.Open s_sql, cn
			if Not rs.EOF then
                qtde_aux = CLng(rs("qtde"))
                qtde_utilizada_aux = CLng(rs("qtde_utilizada"))
				If (qtde_a_sair - qtde_movimentada) > (qtde_aux - qtde_utilizada_aux) Then
				  ' QUANTIDADE DE PRODUTOS DESTE ITEM DE ESTOQUE É INSUFICIENTE P/ ATENDER O PEDIDO
				    qtde_movto = qtde_aux - qtde_utilizada_aux
				Else
				  ' QUANTIDADE DE PRODUTOS DESTE ITEM SOZINHO É SUFICIENTE P/ ATENDER O PEDIDO
				    qtde_movto = qtde_a_sair - qtde_movimentada
				    End If
				rs("qtde_utilizada")=qtde_utilizada_aux + qtde_movto
				rs("data_ult_movimento")=Date
				rs.Update 
				if Err<>0 then 
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
                end if
			if rs.State <> 0 then rs.Close
			
          ' CONTABILIZA QUANTIDADE MOVIMENTADA
            qtde_movimentada = qtde_movimentada + qtde_movto

          ' REGISTRA O MOVIMENTO DE SAÍDA NO ESTOQUE
            if Not gera_id_estoque_movto(s_chave, msg_erro) then 
				msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
				exit function
				end if

            s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
                    " (id_movimento, data, hora, usuario, id_estoque, fabricante, produto," & _
                    " qtde, operacao, estoque, pedido, loja, kit, kit_id_estoque) VALUES (" & _
                    "'" & s_chave & "'," & _
                    bd_formata_data(Date) & "," & _
                    "'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
                    "'" & id_usuario & "'," & _
                    "'" & Trim(v_estoque(iv)) & "'," & _
                    "'" & id_fabricante & "'," & _
                    "'" & id_produto & "'," & _
                    CStr(qtde_movto) & "," & _
                    "'" & OP_ESTOQUE_VENDA & "'," & _
                    "'" & ID_ESTOQUE_VENDIDO & "'," & _
                    "'" & id_pedido & "'," & _
                    "'', 0, '')"
			cn.Execute(s_sql)
			if Err<>0 then 
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

          ' T_ESTOQUE: ATUALIZA DATA DO ÚLTIMO MOVIMENTO
            s_sql = "SELECT id_estoque, data_ult_movimento FROM t_ESTOQUE WHERE" & _
                    " (id_estoque = '" & v_estoque(iv) & "')"

			rs.Open s_sql, cn
			if Not rs.EOF then
				rs("data_ult_movimento")=Date
				rs.Update 
				if Err<>0 then 
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				end if
			if rs.State <> 0 then rs.Close

          ' JÁ CONSEGUIU ALOCAR TUDO?
            If qtde_movimentada >= qtde_a_sair Then Exit For
			end if
		next

	
'   NÃO CONSEGUIU MOVIMENTAR A QUANTIDADE SUFICIENTE	
	if qtde_movimentada < (qtde_a_sair-qtde_autorizada_sem_presenca) then 
		msg_erro="Produto " & id_produto & " do fabricante " & id_fabricante & ": faltam " & _
                 Cstr((qtde_a_sair-qtde_autorizada_sem_presenca)-qtde_movimentada) & " unidades no estoque para poder atender ao pedido."
		exit function
		end if
	
'   REGISTRA A VENDA SEM PRESENÇA NO ESTOQUE
	if (qtde_movimentada < qtde_a_sair) then
      ' REGISTRA O MOVIMENTO DE SAÍDA NO ESTOQUE
        if Not gera_id_estoque_movto(s_chave, msg_erro) then 
			msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
			exit function
			end if
        qtde_estoque_sem_presenca=qtde_a_sair - qtde_movimentada
        s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
                " (id_movimento, data, hora, usuario, id_estoque, fabricante, produto," & _
                " qtde, operacao, estoque, pedido, loja, kit, kit_id_estoque) VALUES (" & _
                "'" & s_chave & "'," & _
                bd_formata_data(Date) & "," & _
                "'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
                "'" & id_usuario & "'," & _
                "''," & _
                "'" & id_fabricante & "'," & _
                "'" & id_produto & "'," & _
                CStr(qtde_estoque_sem_presenca) & "," & _
                "'" & OP_ESTOQUE_VENDA & "'," & _
                "'" & ID_ESTOQUE_SEM_PRESENCA & "'," & _
                "'" & id_pedido & "'," & _
                "'', 0, '')"
		cn.Execute(s_sql)
		if Err<>0 then 
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		end if
		
	qtde_estoque_vendido=qtde_movimentada

	'Log de movimentação do estoque
	if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, id_fabricante, id_produto, qtde_a_sair, qtde_estoque_vendido, OP_ESTOQUE_LOG_VENDA, ID_ESTOQUE_VENDA, ID_ESTOQUE_VENDIDO, "", "", "", id_pedido, "", "", "") then
		msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
		exit function
		end if

	if qtde_estoque_sem_presenca > 0 then
		if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, id_fabricante, id_produto, qtde_estoque_sem_presenca, qtde_estoque_sem_presenca, OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA, "", ID_ESTOQUE_SEM_PRESENCA, "", "", "", id_pedido, "", "", "") then
			msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
			exit function
			end if
		end if
		
	estoque_produto_saida_v2=True
	
end function



' --------------------------------------------------------------------
'   ESTOQUE_VERIFICA_STATUS_ITEM
'   Retorno da função:
'      False - Ocorreu falha ao tentar consultar o banco de dados.
'      True - Conseguiu consultar o banco de dados.
'   Esta função consulta o banco de dados para contabilizar a
'   quantidade de produtos que estão no "estoque vendido" e na
'   lista de produtos vendidos "sem presença no estoque".
'   Note que os itens de pedido a serem analisados são passados
'   pelo vetor do parâmetro v_item.
function estoque_verifica_status_item(byref v_item, byref msg_erro)
dim s
dim s_sql
dim i
dim rs
	estoque_verifica_status_item = False
	msg_erro = ""
	
	for i=Lbound(v_item) to Ubound(v_item)
		with v_item(i)
			.qtde_estoque_vendido = 0
			.qtde_estoque_sem_presenca = 0
			
		'  LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
		'  OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
		'  FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
			s_sql = "SELECT Sum(qtde) AS total FROM t_ESTOQUE_MOVIMENTO WHERE (anulado_status=0)" & _
					" AND (pedido='" & .pedido & "')" & _
					" AND (fabricante='" & .fabricante & "')" & _
					" AND (produto='" & .produto & "')"

			s = s_sql & " AND (estoque='" & ID_ESTOQUE_VENDIDO & "')"
			set rs=cn.execute(s)
			if Err<>0 then 
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			if Not rs.EOF then if IsNumeric(rs("total")) then .qtde_estoque_vendido = CLng(rs("total"))
			if rs.State <> 0 then rs.Close
		
			s = s_sql & " AND (estoque='" & ID_ESTOQUE_SEM_PRESENCA & "')"
			set rs=cn.execute(s)
			if Err<>0 then 
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			if Not rs.EOF then if IsNumeric(rs("total")) then .qtde_estoque_sem_presenca = CLng(rs("total"))
			if rs.State <> 0 then rs.Close
			end with
		next
		
	estoque_verifica_status_item = True
end function



' --------------------------------------------------------------------
'   ESTOQUE_PRODUTO_ESTORNA
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros entre as várias tabelas.
'   Esta função estorna a quantidade de produtos indicada no parâmetro
'   "qtde_a_estornar" do "estoque vendido" para o "estoque de venda".
'   Se o parâmetro "qtde_a_estornar" for especificado com o valor
'   "COD_NEGATIVO_UM", então o estorno será integral.
'   27/01/2017: revisado para estar em conformidade c/ o controle de estoque por empresa.
function estoque_produto_estorna(ByVal id_usuario, ByVal id_pedido, _
								 ByVal id_fabricante, ByVal id_produto, ByVal qtde_a_estornar, _
								 ByRef qtde_estornada, ByRef msg_erro)
dim n
dim iv
dim rs
dim s_chave
dim s_sql
dim v_estoque
dim id_estoque_aux
dim qtde_aux
dim qtde_utilizada_aux
dim qtde_movto
dim operacao_aux
dim blnGravarLog
dim id_nfe_emitente

	estoque_produto_estorna = False
    msg_erro = ""
    qtde_estornada = 0
	id_nfe_emitente = 0

	id_usuario = Trim("" & id_usuario)
	id_pedido = Trim("" & id_pedido)
	id_fabricante = Trim("" & id_fabricante)
	id_produto = Trim("" & id_produto)

  ' 1) LEMBRE-SE DE QUE PODE HAVER MAIS DE UM REGISTRO EM T_ESTOQUE_MOVIMENTO 
  '    P/ CADA PRODUTO, POIS PODEM TER SIDO USADOS DIFERENTES LOTES DO ESTOQUE 
  '    P/ ATENDER A UM ÚNICO PEDIDO!!
  ' 2) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
  '    OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
  '    FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
    ReDim v_estoque(0)
    v_estoque(UBound(v_estoque)) = ""
	
    s_sql = "SELECT id_movimento FROM t_ESTOQUE INNER JOIN t_ESTOQUE_MOVIMENTO ON (t_ESTOQUE.id_estoque = t_ESTOQUE_MOVIMENTO.id_estoque)" & _
			" WHERE (anulado_status = 0)" & _
            " AND (estoque = '" & ID_ESTOQUE_VENDIDO & "')" & _
            " AND (pedido = '" & id_pedido & "')" & _
            " AND (t_ESTOQUE_MOVIMENTO.fabricante = '" & id_fabricante & "')" & _
            " AND (produto = '" & id_produto & "')" & _
            " ORDER BY t_ESTOQUE.data_entrada DESC, t_ESTOQUE.id_estoque DESC"
	set rs=cn.execute(s_sql)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	do while Not rs.EOF 
        If v_estoque(UBound(v_estoque)) <> "" Then
            ReDim Preserve v_estoque(UBound(v_estoque) + 1)
            v_estoque(UBound(v_estoque)) = ""
            End If
		v_estoque(UBound(v_estoque)) = Trim("" & rs("id_movimento"))
		rs.MoveNext 
		loop
		
	if rs.State <> 0 then rs.Close
	set rs=nothing
	
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function
			
	for iv=LBound(v_estoque) To UBound(v_estoque)
		If Trim(v_estoque(iv)) <> "" Then
          
          ' JÁ ESTORNOU TUDO?
            If qtde_a_estornar <> COD_NEGATIVO_UM Then
                If qtde_estornada >= qtde_a_estornar Then Exit For
                End If
			
		  ' T_ESTOQUE_MOVIMENTO: ANULA O MOVIMENTO	
		  ' ======================================
            s_sql = "SELECT *" & _
					" FROM t_ESTOQUE_MOVIMENTO" & _
					" WHERE (anulado_status = 0)" & _
                    " AND (id_movimento = '" & Trim(v_estoque(iv)) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro="Falha ao acessar o registro de movimento no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				exit function
				end if

			id_estoque_aux = Trim("" & rs("id_estoque"))
			qtde_aux = CLng(rs("qtde"))
			operacao_aux = Trim("" & rs("operacao"))
			
			qtde_movto=qtde_aux
			
          ' É PARA ESTORNAR TUDO OU UMA QUANTIDADE ESPECIFICADA?
            If qtde_a_estornar <> COD_NEGATIVO_UM Then
              ' A QUANTIDADE QUE FALTA SER ESTORNADA É MENOR QUE A QUANTIDADE DO MOVIMENTO
                If (qtde_a_estornar - qtde_estornada) < qtde_aux Then
                    qtde_movto = qtde_a_estornar - qtde_estornada
                    End If
                End If
			
          ' ANULA O MOVIMENTO
			rs("anulado_status") = 1
			rs("anulado_data") = Date
			rs("anulado_hora") = retorna_so_digitos(formata_hora(Now))
			rs("anulado_usuario") = id_usuario
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

          ' ESTORNO PARCIAL: O MOVIMENTO ORIGINAL FOI ANULADO E UM NOVO MOVIMENTO 
          ' C/ A QUANTIDADE RESTANTE DEVE SER GRAVADO!!
            If qtde_movto < qtde_aux Then
              ' REGISTRA O MOVIMENTO DE SAÍDA NO ESTOQUE
				if Not gera_id_estoque_movto(s_chave, msg_erro) then 
					msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
					exit function
					end if
				
                s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
                        " (id_movimento, data, hora, usuario, pedido, fabricante, produto, id_estoque," & _
                        " qtde, operacao, estoque, loja, kit, kit_id_estoque) VALUES (" & _
                        "'" & s_chave & "'," & _
                        bd_formata_data(Date) & "," & _
                        "'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
                        "'" & id_usuario & "'," & _
                        "'" & id_pedido & "'," & _
                        "'" & id_fabricante & "'," & _
                        "'" & id_produto & "'," & _
                        "'" & id_estoque_aux & "'," & _
                        CStr(qtde_aux - qtde_movto) & "," & _
                        "'" & operacao_aux & "'," & _
                        "'" & ID_ESTOQUE_VENDIDO & "'," & _
                        "'', 0, '')"
				cn.Execute(s_sql)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
                End If
			
		  
		  ' T_ESTOQUE_ITEM: ESTORNA PRODUTOS AO SALDO
		  ' =========================================
            s_sql = "SELECT data_ult_movimento, qtde_utilizada FROM t_ESTOQUE_ITEM WHERE" & _
                    " (id_estoque = '" & id_estoque_aux & "') AND" & _
                    " (fabricante = '" & id_fabricante & "') AND" & _
                    " (produto = '" & id_produto & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro = "Falha ao acessar o registro no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				exit function
				end if
			
			qtde_utilizada_aux = CLng(rs("qtde_utilizada"))

          ' PRECAUÇÃO (P/ GARANTIR QUE "QTDE_UTILIZADA" NUNCA FICARÁ C/ VALOR NEGATIVO)!!
            n = qtde_movto
            If qtde_utilizada_aux < qtde_movto Then n = qtde_utilizada_aux
			
			rs("qtde_utilizada") = rs("qtde_utilizada") - n
			rs("data_ult_movimento") = Date
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
          
          ' CONTABILIZA QUANTIDADE ESTORNADA
            qtde_estornada = qtde_estornada + qtde_movto
                                                                
          
          ' T_ESTOQUE: ATUALIZA DATA DO ÚLTIMO MOVIMENTO
          ' ============================================
            s_sql = "SELECT data_ult_movimento, id_nfe_emitente FROM t_ESTOQUE WHERE" & _
                    " (id_estoque = '" & id_estoque_aux & "')"
            
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro = "Falha ao acessar o registro principal no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				exit function
				end if
			
			id_nfe_emitente = rs("id_nfe_emitente")

			rs("data_ult_movimento") = Date
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			end if
		next

	blnGravarLog=True
	if (qtde_a_estornar = COD_NEGATIVO_UM) And (qtde_estornada = 0) then blnGravarLog=False

	if blnGravarLog then
		'Log de movimentação do estoque
		if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, id_fabricante, id_produto, qtde_a_estornar, qtde_estornada, OP_ESTOQUE_LOG_ESTORNO, ID_ESTOQUE_VENDIDO, ID_ESTOQUE_VENDA, "", "", id_pedido, "", "", "", "") then
			msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
			exit function
			end if
		end if
				
	estoque_produto_estorna = True
end function



' --------------------------------------------------------------------
'   ESTOQUE_PRODUTO_CANCELA_LISTA_SEM_PRESENCA
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros entre as várias tabelas.
'   Esta função cancela a quantidade de produtos indicada no parâmetro
'   "qtde_a_cancelar" da lista de produtos vendidos sem presença no 
'	estoque.
'   Se o parâmetro "qtde_a_cancelar" for especificado com o valor
'   "COD_NEGATIVO_UM", então o cancelamento será integral.
'   27/01/2017: revisado p/ estar em conformidade c/ o controle de estoque por empresa.
function estoque_produto_cancela_lista_sem_presenca(ByVal id_usuario, ByVal id_pedido, _
							ByVal id_fabricante, ByVal id_produto, ByVal qtde_a_cancelar, _
							ByRef qtde_cancelada, ByRef msg_erro)
dim iv
dim rs
dim s_chave
dim s_sql
dim v_estoque
dim qtde_aux
dim qtde_movto
dim operacao_aux
dim blnGravarLog
dim id_nfe_emitente

	estoque_produto_cancela_lista_sem_presenca = False
    msg_erro = ""
    qtde_cancelada = 0
	
	id_usuario = Trim("" & id_usuario)
	id_pedido = Trim("" & id_pedido)
	id_fabricante = Trim("" & id_fabricante)
	id_produto = Trim("" & id_produto)

	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

	s_sql = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" & id_pedido & "')"
	if rs.State <> 0 then rs.Close
	rs.Open s_sql, cn
	if rs.Eof then
		msg_erro="Falha ao tentar localizar o registro do pedido " & id_pedido & "!!"
		exit function
		end if

	id_nfe_emitente = rs("id_nfe_emitente")

'   LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'   OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'   FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
    ReDim v_estoque(0)
    v_estoque(UBound(v_estoque)) = ""
	
    s_sql = "SELECT id_movimento FROM t_ESTOQUE_MOVIMENTO" & _
			" WHERE (anulado_status = 0)" & _
            " AND (estoque = '" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
            " AND (pedido = '" & id_pedido & "')" & _
            " AND (fabricante = '" & id_fabricante & "')" & _
            " AND (produto = '" & id_produto & "')"
	if rs.State <> 0 then rs.Close
	rs.Open s_sql, cn
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	do while Not rs.EOF 
        If v_estoque(UBound(v_estoque)) <> "" Then
            ReDim Preserve v_estoque(UBound(v_estoque) + 1)
            v_estoque(UBound(v_estoque)) = ""
            End If
		v_estoque(UBound(v_estoque)) = Trim("" & rs("id_movimento"))
		rs.MoveNext 
		loop
		
		
	for iv=LBound(v_estoque) To UBound(v_estoque)
		If Trim(v_estoque(iv)) <> "" Then
          
          ' JÁ CANCELOU TUDO?
            If qtde_a_cancelar <> COD_NEGATIVO_UM Then
                If qtde_cancelada >= qtde_a_cancelar Then Exit For
                End If
			
		  ' T_ESTOQUE_MOVIMENTO: ANULA O MOVIMENTO	
		  ' ======================================
            s_sql = "SELECT *" & _
					" FROM t_ESTOQUE_MOVIMENTO" & _
					" WHERE (anulado_status = 0)" & _
                    " AND (id_movimento = '" & Trim(v_estoque(iv)) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro="Falha ao acessar o registro de movimento no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				exit function
				end if

			qtde_aux = CLng(rs("qtde"))
			operacao_aux = Trim("" & rs("operacao"))
			
			qtde_movto=qtde_aux
			
          ' É PARA CANCELAR TUDO OU UMA QUANTIDADE ESPECIFICADA?
            If qtde_a_cancelar <> COD_NEGATIVO_UM Then
              ' A QUANTIDADE QUE FALTA SER CANCELADA É MENOR QUE A QUANTIDADE DO MOVIMENTO
                If (qtde_a_cancelar - qtde_cancelada) < qtde_aux Then
                    qtde_movto = qtde_a_cancelar - qtde_cancelada
                    End If
                End If
			
          ' ANULA O MOVIMENTO
			rs("anulado_status") = 1
			rs("anulado_data") = Date
			rs("anulado_hora") = retorna_so_digitos(formata_hora(Now))
			rs("anulado_usuario") = id_usuario
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

          ' CANCELAMENTO PARCIAL: O MOVIMENTO ORIGINAL FOI ANULADO E UM NOVO MOVIMENTO 
          ' C/ A QUANTIDADE RESTANTE DEVE SER GRAVADO!!
            If qtde_movto < qtde_aux Then
              ' REGISTRA O MOVIMENTO QUE CONTABILIZA OS PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE
				if Not gera_id_estoque_movto(s_chave, msg_erro) then 
					msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
					exit function
					end if
				
                s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
                        " (id_movimento, data, hora, usuario, pedido, fabricante, produto, id_estoque," & _
                        " qtde, operacao, estoque, loja, kit, kit_id_estoque) VALUES (" & _
                        "'" & s_chave & "'," & _
                        bd_formata_data(Date) & "," & _
                        "'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
                        "'" & id_usuario & "'," & _
                        "'" & id_pedido & "'," & _
                        "'" & id_fabricante & "'," & _
                        "'" & id_produto & "'," & _
                        "''," & _
                        CStr(qtde_aux - qtde_movto) & "," & _
                        "'" & operacao_aux & "'," & _
                        "'" & ID_ESTOQUE_SEM_PRESENCA & "'," & _
                        "'', 0, '')"
				cn.Execute(s_sql)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
                End If
			
          
          ' CONTABILIZA QUANTIDADE CANCELADA
            qtde_cancelada = qtde_cancelada + qtde_movto
			end if
		next

	blnGravarLog=True
	if (qtde_a_cancelar = COD_NEGATIVO_UM) And (qtde_cancelada = 0) then blnGravarLog=False
	
	if blnGravarLog then
		'Log de movimentação do estoque
		if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, id_fabricante, id_produto, qtde_a_cancelar, qtde_cancelada, OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA, ID_ESTOQUE_SEM_PRESENCA, "", "", "", id_pedido, "", "", "", "") then
			msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
			exit function
			end if
		end if
				
	estoque_produto_cancela_lista_sem_presenca = True
end function



' --------------------------------------------------------------------
'   ESTOQUE_PEDIDO_CANCELA
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros entre as várias tabelas.
'   Esta função processa o cancelamento do pedido com relação
'   aos produtos no estoque. 
'   Portanto, os produtos que estiverem no "estoque vendido" serão
'   estornados ao "estoque de venda".
'   Os produtos que estiverem na lista de produtos vendidos
'   "sem presença no estoque" serão cancelados.
'	O log da movimentação no estoque (T_ESTOQUE_LOG) é gravado
'	dentro das rotinas chamadas por esta:
'		1) estoque_produto_estorna()
'		2) estoque_produto_cancela_lista_sem_presenca()
'   27/01/2017: revisado p/ estar em conformidade c/ o controle de estoque por empresa.
function estoque_pedido_cancela(byval id_usuario, byval id_pedido, byref info_log, byref msg_erro)
dim i
dim rs
dim s_sql
dim qtde_estornada
dim qtde_cancelada
dim s_log_estorno
dim s_log_cancela
dim v_produto

	estoque_pedido_cancela = False
    msg_erro = ""
    info_log = ""
	
	s_log_estorno=""
	s_log_cancela=""

	redim v_produto(0)
	set v_produto(Ubound(v_produto)) = New cl_DUAS_COLUNAS
	v_produto(Ubound(v_produto)).c1 = ""
	v_produto(Ubound(v_produto)).c2 = ""
		
    s_sql = "SELECT fabricante, produto FROM t_PEDIDO_ITEM" & _
			" WHERE (pedido = '" & id_pedido & "')"
	set rs=cn.execute(s_sql)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

	do while Not rs.EOF 
		if Trim(v_produto(Ubound(v_produto)).c2)<>"" then
			redim preserve v_produto(Ubound(v_produto)+1)
			set v_produto(Ubound(v_produto)) = New cl_DUAS_COLUNAS
			end if
		with v_produto(Ubound(v_produto))
			.c1 = Trim("" & rs("fabricante"))
			.c2 = Trim("" & rs("produto"))
			end with
		rs.MoveNext 
		loop

	if rs.State <> 0 then rs.Close
	set rs=nothing

	for i = Lbound(v_produto) to Ubound(v_produto)			
		with v_produto(i)
			if .c2 <> "" then
				If Not estoque_produto_estorna(id_usuario, id_pedido, .c1, .c2, COD_NEGATIVO_UM, qtde_estornada, msg_erro) then	exit function
				If Not estoque_produto_cancela_lista_sem_presenca(id_usuario, id_pedido, .c1, .c2, COD_NEGATIVO_UM, qtde_cancelada, msg_erro) then exit function

				if qtde_estornada > 0 then s_log_estorno=s_log_estorno & log_estoque_monta_incremento(qtde_estornada, .c1, .c2)
				if qtde_cancelada > 0 then s_log_cancela=s_log_cancela & log_estoque_monta_incremento(qtde_cancelada, .c1, .c2)
				end if
			end with
		next

	if s_log_estorno <> "" then s_log_estorno = "Produtos estornados do estoque vendido para o estoque de venda:" & s_log_estorno
	if s_log_cancela <> "" then s_log_cancela = "Produtos cancelados da lista de produtos vendidos sem presença no estoque:" & s_log_cancela
	
	info_log=s_log_estorno
	if info_log <> "" then info_log=info_log & chr(13)
	info_log=info_log & s_log_cancela
	
	estoque_pedido_cancela = True
end function



' --------------------------------------------------------------------
'   ESTOQUE_PRODUTO_SPLIT_V2
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função processa o split do pedido, ou seja, transfere os
'   produtos disponíveis de um pedido para seu pedido filhote.
function estoque_produto_split_v2(byval id_usuario, byval id_pedido, byval id_pedido_filhote, _
							   byval id_fabricante, byval id_produto, _
							   byval qtde_a_transferir, byref msg_erro)
dim rs
dim iv
dim s_sql
dim s_chave
dim v_estoque
dim id_estoque_aux
dim qtde_aux
dim operacao_aux
dim qtde_transferida
dim qtde_movto
dim id_nfe_emitente_pedido, id_nfe_emitente_pedido_filhote

	estoque_produto_split_v2 = False
	msg_erro = ""
	
	id_usuario = Trim("" & id_usuario)
	id_pedido = Trim("" & id_pedido)
	id_pedido_filhote = Trim("" & id_pedido_filhote)
	id_fabricante = Trim("" & id_fabricante)
	id_produto = Trim("" & id_produto)
	
	if Not IsNumeric(qtde_a_transferir) then exit function
	qtde_a_transferir = CLng(qtde_a_transferir)
	
'	VERIFICA SE O PEDIDO E O PEDIDO-FILHOTE ESTÃO VINCULADOS AO ESTOQUE DA MESMA EMPRESA
	id_nfe_emitente_pedido = 0
	id_nfe_emitente_pedido_filhote = 0
	s_sql = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" & id_pedido & "')"
	set rs=cn.Execute(s_sql)
	if Not rs.Eof then
		id_nfe_emitente_pedido = CLng(rs("id_nfe_emitente"))
		end if

	if rs.State <> 0 then rs.Close
	set rs=nothing

	s_sql = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" & id_pedido_filhote & "')"
	set rs=cn.Execute(s_sql)
	if Not rs.Eof then
		id_nfe_emitente_pedido_filhote = CLng(rs("id_nfe_emitente"))
		end if

	if rs.State <> 0 then rs.Close
	set rs=nothing

	if id_nfe_emitente_pedido <> id_nfe_emitente_pedido_filhote then
		msg_erro="A operação não pode ser realizada porque os pedidos estão associados a estoques de empresas diferentes:" & _
				"<br />Pedido (" & id_pedido & "): " & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_pedido) & _
				"<br />Pedido-filhote (" & id_pedido_filhote & "): " & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_pedido_filhote)
		exit function
		end if
	
  ' 1) LEMBRE-SE DE QUE PODE HAVER MAIS DE UM REGISTRO EM T_ESTOQUE_MOVIMENTO 
  '    P/ CADA PRODUTO, POIS PODEM TER SIDO USADOS DIFERENTES LOTES DO ESTOQUE 
  '    P/ ATENDER A UM ÚNICO PEDIDO!!
  ' 2) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
  '    OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
  '    FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.

    ReDim v_estoque(0)
    v_estoque(UBound(v_estoque)) = ""

'   SERÃO TRANSFERIDOS PARA O PEDIDO FILHOTE OS PRODUTOS QUE ENTRARAM HÁ MAIS 
'   TEMPO NO ESTOQUE (FIFO), JÁ QUE O PEDIDO FILHOTE SERÁ ENTREGUE ANTES QUE
'   O PEDIDO ORIGINAL.
    s_sql = "SELECT id_movimento FROM t_ESTOQUE INNER JOIN t_ESTOQUE_MOVIMENTO ON (t_ESTOQUE.id_estoque = t_ESTOQUE_MOVIMENTO.id_estoque)" & _
			" WHERE (anulado_status = 0)" & _
            " AND (estoque = '" & ID_ESTOQUE_VENDIDO & "')" & _
            " AND (pedido = '" & id_pedido & "')" & _
            " AND (t_ESTOQUE_MOVIMENTO.fabricante = '" & id_fabricante & "')" & _
            " AND (produto = '" & id_produto & "')" & _
            " ORDER BY t_ESTOQUE.data_entrada, t_ESTOQUE.id_estoque"
	set rs=cn.execute(s_sql)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	do while Not rs.EOF 
        If v_estoque(UBound(v_estoque)) <> "" Then
            ReDim Preserve v_estoque(UBound(v_estoque) + 1)
            v_estoque(UBound(v_estoque)) = ""
            End If
		v_estoque(UBound(v_estoque)) = Trim("" & rs("id_movimento"))
		rs.MoveNext 
		loop
		
	if rs.State <> 0 then rs.Close
	set rs=nothing
	
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function
			
	qtde_transferida = 0
	for iv=LBound(v_estoque) To UBound(v_estoque)
		If Trim(v_estoque(iv)) <> "" Then

          ' JÁ TRANSFERIU TUDO?
			If qtde_transferida >= qtde_a_transferir Then Exit For

		  ' T_ESTOQUE_MOVIMENTO: ANULA O MOVIMENTO
		  ' ======================================
            s_sql = "SELECT *" & _
					" FROM t_ESTOQUE_MOVIMENTO" & _
					" WHERE (anulado_status = 0)" & _
                    " AND (id_movimento = '" & Trim(v_estoque(iv)) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro="Falha ao acessar o registro de movimento no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				exit function
				end if

			id_estoque_aux = Trim("" & rs("id_estoque"))
			qtde_aux = CLng(rs("qtde"))
			operacao_aux = Trim("" & rs("operacao"))
			
			qtde_movto=qtde_aux
			
          ' A QUANTIDADE QUE FALTA SER TRANSFERIDA É MENOR QUE A QUANTIDADE DO MOVIMENTO
            If (qtde_a_transferir - qtde_transferida) < qtde_aux Then
                qtde_movto = qtde_a_transferir - qtde_transferida
                End If
			
          ' ANULA O MOVIMENTO
			rs("anulado_status") = 1
			rs("anulado_data") = Date
			rs("anulado_hora") = retorna_so_digitos(formata_hora(Now))
			rs("anulado_usuario") = id_usuario
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

          ' REGISTRA O MOVIMENTO DE SAÍDA NO ESTOQUE PARA O PEDIDO FILHOTE
			if Not gera_id_estoque_movto(s_chave, msg_erro) then 
				msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
				exit function
				end if
				
            s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
                    " (id_movimento, data, hora, usuario, pedido, fabricante, produto, id_estoque," & _
                    " qtde, operacao, estoque, loja, kit, kit_id_estoque) VALUES (" & _
                    "'" & s_chave & "'," & _
                    bd_formata_data(Date) & "," & _
                    "'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
                    "'" & id_usuario & "'," & _
                    "'" & id_pedido_filhote & "'," & _
                    "'" & id_fabricante & "'," & _
                    "'" & id_produto & "'," & _
                    "'" & id_estoque_aux & "'," & _
                    CStr(qtde_movto) & "," & _
                    "'" & operacao_aux & "'," & _
                    "'" & ID_ESTOQUE_VENDIDO & "'," & _
                    "'', 0, '')"
			cn.Execute(s_sql)
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if


          ' TRANSFERÊNCIA PARCIAL: O MOVIMENTO ORIGINAL FOI ANULADO E UM NOVO MOVIMENTO 
          ' C/ A QUANTIDADE RESTANTE DEVE SER GRAVADO!!
            If qtde_movto < qtde_aux Then
              ' REGISTRA O MOVIMENTO DE SAÍDA NO ESTOQUE PARA A QUANTIDADE RESTANTE NO PEDIDO ORIGINAL
				if Not gera_id_estoque_movto(s_chave, msg_erro) then 
					msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
					exit function
					end if
				
                s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
                        " (id_movimento, data, hora, usuario, pedido, fabricante, produto, id_estoque," & _
                        " qtde, operacao, estoque, loja, kit, kit_id_estoque) VALUES (" & _
                        "'" & s_chave & "'," & _
                        bd_formata_data(Date) & "," & _
                        "'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
                        "'" & id_usuario & "'," & _
                        "'" & id_pedido & "'," & _
                        "'" & id_fabricante & "'," & _
                        "'" & id_produto & "'," & _
                        "'" & id_estoque_aux & "'," & _
                        CStr(qtde_aux - qtde_movto) & "," & _
                        "'" & operacao_aux & "'," & _
                        "'" & ID_ESTOQUE_VENDIDO & "'," & _
                        "'', 0, '')"
				cn.Execute(s_sql)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
                End If


          ' CONTABILIZA QUANTIDADE TRANSFERIDA
            qtde_transferida = qtde_transferida + qtde_movto
			
			
		  ' T_ESTOQUE_ITEM: ATUALIZA DATA DO ÚLTIMO MOVIMENTO
		  ' =================================================
            s_sql = "SELECT data_ult_movimento FROM t_ESTOQUE_ITEM WHERE" & _
                    " (id_estoque = '" & id_estoque_aux & "') AND" & _
                    " (fabricante = '" & id_fabricante & "') AND" & _
                    " (produto = '" & id_produto & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro = "Falha ao acessar o registro no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				exit function
				end if

			rs("data_ult_movimento") = Date
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
		  
		  
          ' T_ESTOQUE: ATUALIZA DATA DO ÚLTIMO MOVIMENTO
          ' ============================================
            s_sql = "SELECT data_ult_movimento FROM t_ESTOQUE WHERE" & _
                    " (id_estoque = '" & id_estoque_aux & "')"
            
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro = "Falha ao acessar o registro principal no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				exit function
				end if
			
			rs("data_ult_movimento") = Date
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			end if
		next
	
'   CONSEGUIU TRANSFERIR?
	if qtde_transferida < qtde_a_transferir then 
		msg_erro="Produto " & id_produto & " do fabricante " & id_fabricante & ": " & Cstr(qtde_a_transferir - qtde_transferida) & " unidades não foram transferidas."
		exit function
		end if
	
	'Log de movimentação do estoque
	if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente_pedido, id_fabricante, id_produto, qtde_a_transferir, qtde_transferida, OP_ESTOQUE_LOG_SPLIT, ID_ESTOQUE_VENDIDO, ID_ESTOQUE_VENDIDO, "", "", id_pedido, id_pedido_filhote, "", "", "") then
		msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
		exit function
		end if
	
	estoque_produto_split_v2 = True
end function



' --------------------------------------------------------------------
'   ESTOQUE NOVA ENTRADA MERCADORIAS
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função grava a entrada de mercadorias no estoque, sendo que
'   o identificador gerado para o lote no estoque é retornado no 
'	próprio parâmetro.
'	A opção "entrada especial" é usado p/ não computar no relatório de compras.
function estoque_nova_entrada_mercadorias(byref r_estoque, byval v_item, byref msg_erro)
dim id_estoque
dim s_sql
dim i
dim i_seq
dim strComplemento

	estoque_nova_entrada_mercadorias = False
	msg_erro = ""

	If Not gera_id_estoque(id_estoque, msg_erro) Then Exit Function

'	INFORMAÇÃO ADICIONAL PARA O LOG DA MOVIMENTAÇÃO NO ESTOQUE
	strComplemento = ""
	if Cstr(r_estoque.entrada_especial) <> Cstr(0) then strComplemento = "ENTRADA_ESPECIAL"
	
' 	GRAVA INFORMAÇÕES BÁSICAS DA ENTRADA NO ESTOQUE
	With r_estoque
		s_sql = "INSERT INTO t_ESTOQUE" & _
				" (id_estoque, data_entrada, hora_entrada, fabricante, documento," & _
				" usuario, data_ult_movimento, kit, entrada_especial, obs, id_nfe_emitente" & _
			") VALUES (" & _
				"'" & id_estoque & "'" & _
				"," & bd_formata_data(.data_entrada) & _
				",'" & Trim(.hora_entrada) & "'" & _
				",'" & Trim(.fabricante) & "'" & _
				",'" & Trim(.documento) & "'" & _
				",'" & Trim(.usuario) & "'" & _
				"," & bd_formata_data(.data_ult_movimento) & _
				"," & Cstr(.kit) & _
				"," & Cstr(.entrada_especial) & _
				",'" & QuotedStr(Trim(.obs)) & "'" & _
				"," & Cstr(.id_nfe_emitente) & _
				")"
		cn.Execute(s_sql)
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		end with
	
'	GRAVA LISTA DE PRODUTOS
	i_seq = 0
	For i = LBound(v_item) To UBound(v_item)
		With v_item(i)
			If Trim(.produto) <> "" Then
				i_seq = i_seq + 1
				If Not IsDate(.data_ult_movimento) Then .data_ult_movimento = Date
				s_sql = "INSERT INTO T_ESTOQUE_ITEM" & _
						" (id_estoque, fabricante, produto, qtde, preco_fabricante, vl_custo2," & _
						" vl_BC_ICMS_ST, vl_ICMS_ST," & _
						" ncm, cst," & _
						" data_ult_movimento, sequencia)" & _
						" VALUES (" & _
						"'" & id_estoque & "'" & _
						",'" & Trim(.fabricante) & "'" & _
						",'" & Trim(.produto) & "'" & _
						"," & CStr(.qtde) & _
						"," & bd_formata_numero(.preco_fabricante) & _
						"," & bd_formata_numero(.vl_custo2) & _
						"," & bd_formata_numero(.vl_BC_ICMS_ST) & _
						"," & bd_formata_numero(.vl_ICMS_ST) & _
						",'" & Trim(.ncm) & "'" & _
						",'" & Trim(.cst) & "'" & _
						"," & bd_formata_data(.data_ult_movimento) & _
						"," & CStr(i_seq) & _
						")"
				cn.Execute(s_sql)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				
				'Log de movimentação do estoque
				if Not grava_log_estoque_v2(r_estoque.usuario, r_estoque.id_nfe_emitente, .fabricante, .produto, .qtde, .qtde, OP_ESTOQUE_LOG_ENTRADA, "", ID_ESTOQUE_VENDA, "", "", "", "", r_estoque.documento, strComplemento, "") then
					msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
					exit function
					end if
				End If
			End With
		Next
	
	r_estoque.id_estoque = id_estoque
	
	estoque_nova_entrada_mercadorias = True
end function

' --------------------------------------------------------------------
'   ESTOQUE NOVA ENTRADA MERCADORIAS AGIO
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função grava a entrada de mercadorias no estoque, sendo que
'   o identificador gerado para o lote no estoque é retornado no 
'	próprio parâmetro.
'	A opção "entrada especial" é usado p/ não computar no relatório de compras.
function estoque_nova_entrada_mercadorias_agio(byref r_estoque, byval v_item, byref msg_erro)
dim id_estoque
dim s_sql
dim i
dim i_seq
dim strComplemento

	estoque_nova_entrada_mercadorias_agio = False
	msg_erro = ""

	If Not gera_id_estoque(id_estoque, msg_erro) Then Exit Function

'	INFORMAÇÃO ADICIONAL PARA O LOG DA MOVIMENTAÇÃO NO ESTOQUE
	strComplemento = ""
	if Cstr(r_estoque.entrada_especial) <> Cstr(0) then strComplemento = "ENTRADA_ESPECIAL"
	
' 	GRAVA INFORMAÇÕES BÁSICAS DA ENTRADA NO ESTOQUE
	With r_estoque
		s_sql = "INSERT INTO t_ESTOQUE" & _
				" (id_estoque, data_entrada, hora_entrada, fabricante, documento," & _
				" usuario, data_ult_movimento, kit, entrada_especial, obs, id_nfe_emitente, perc_agio, entrada_tipo" & _
			") VALUES (" & _
				"'" & id_estoque & "'" & _
				"," & bd_formata_data(.data_entrada) & _
				",'" & Trim(.hora_entrada) & "'" & _
				",'" & Trim(.fabricante) & "'" & _
				",'" & Trim(.documento) & "'" & _
				",'" & Trim(.usuario) & "'" & _
				"," & bd_formata_data(.data_ult_movimento) & _
				"," & Cstr(.kit) & _
				"," & Cstr(.entrada_especial) & _
				",'" & QuotedStr(Trim(.obs)) & "'" & _
				"," & Cstr(.id_nfe_emitente) & _
                "," & bd_formata_numero(.perc_agio) & _
                "," & "0" & _
				")"
		cn.Execute(s_sql)
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		end with
	
'	GRAVA LISTA DE PRODUTOS
	i_seq = 0
	For i = LBound(v_item) To UBound(v_item)
		With v_item(i)
			If Trim(.produto) <> "" Then
				i_seq = i_seq + 1
				If Not IsDate(.data_ult_movimento) Then .data_ult_movimento = Date
				s_sql = "INSERT INTO T_ESTOQUE_ITEM" & _
						" (id_estoque, fabricante, produto, qtde, preco_fabricante, vl_custo2," & _
						" vl_BC_ICMS_ST, vl_ICMS_ST," & _
						" ncm, cst," & _
						" data_ult_movimento, sequencia, aliq_ipi, aliq_icms, vl_ipi)" & _
						" VALUES (" & _
						"'" & id_estoque & "'" & _
						",'" & Trim(.fabricante) & "'" & _
						",'" & Trim(.produto) & "'" & _
						"," & CStr(.qtde) & _
						"," & bd_formata_numero(.preco_fabricante) & _
						"," & bd_formata_numero(.vl_custo2) & _
						"," & bd_formata_numero(.vl_BC_ICMS_ST) & _
						"," & bd_formata_numero(.vl_ICMS_ST) & _
						",'" & Trim(.ncm) & "'" & _
						",'" & Trim(.cst) & "'" & _
						"," & bd_formata_data(.data_ult_movimento) & _
						"," & CStr(i_seq) & _
                        "," & bd_formata_numero(.aliq_ipi) & _
                        "," & bd_formata_numero(.aliq_icms) & _
                        "," & bd_formata_numero(.vl_ipi) & _
						")"
				cn.Execute(s_sql)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				
				'Log de movimentação do estoque
				if Not grava_log_estoque_v2(r_estoque.usuario, r_estoque.id_nfe_emitente, .fabricante, .produto, .qtde, .qtde, OP_ESTOQUE_LOG_ENTRADA, "", ID_ESTOQUE_VENDA, "", "", "", "", r_estoque.documento, strComplemento, "") then
					msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
					exit function
					end if
				End If
			End With
		Next
	
	r_estoque.id_estoque = id_estoque
	
	estoque_nova_entrada_mercadorias_agio = True
end function

' --------------------------------------------------------------------
'   ESTOQUE NOVA ENTRADA MERCADORIAS XML
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função grava a entrada de mercadorias no estoque, sendo que
'   o identificador gerado para o lote no estoque é retornado no 
'	próprio parâmetro.
'	A opção "entrada especial" é usado p/ não computar no relatório de compras.
function estoque_nova_entrada_mercadorias_xml(byref r_estoque, byval v_item, byref msg_erro)
dim id_estoque
dim s_sql
dim i
dim i_seq
dim strComplemento

	estoque_nova_entrada_mercadorias_xml = False
	msg_erro = ""

	If Not gera_id_estoque(id_estoque, msg_erro) Then Exit Function

'	INFORMAÇÃO ADICIONAL PARA O LOG DA MOVIMENTAÇÃO NO ESTOQUE
	strComplemento = ""
	if Cstr(r_estoque.entrada_especial) <> Cstr(0) then strComplemento = "ENTRADA_ESPECIAL"
	
' 	GRAVA INFORMAÇÕES BÁSICAS DA ENTRADA NO ESTOQUE
	With r_estoque
		s_sql = "INSERT INTO t_ESTOQUE" & _
				" (id_estoque, data_entrada, hora_entrada, fabricante, documento," & _
				" usuario, data_ult_movimento, kit, entrada_especial, obs, id_nfe_emitente, perc_agio, entrada_tipo " & _
			") VALUES (" & _
				"'" & id_estoque & "'" & _
				"," & bd_formata_data(.data_entrada) & _
				",'" & Trim(.hora_entrada) & "'" & _
				",'" & Trim(.fabricante) & "'" & _
				",'" & Trim(.documento) & "'" & _
				",'" & Trim(.usuario) & "'" & _
				"," & bd_formata_data(.data_ult_movimento) & _
				"," & Cstr(.kit) & _
				"," & Cstr(.entrada_especial) & _
				",'" & QuotedStr(Trim(.obs)) & "'" & _
				"," & Cstr(.id_nfe_emitente) & _
                "," & bd_formata_numero(.perc_agio) & _
                "," & "1" & _
				")"
		cn.Execute(s_sql)
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		end with
	
'	GRAVA LISTA DE PRODUTOS
	i_seq = 0
	For i = LBound(v_item) To UBound(v_item)
		With v_item(i)
			If Trim(.produto) <> "" Then
				i_seq = i_seq + 1
				If Not IsDate(.data_ult_movimento) Then .data_ult_movimento = Date
				s_sql = "INSERT INTO T_ESTOQUE_ITEM" & _
						" (id_estoque, fabricante, produto, qtde, preco_fabricante, vl_custo2," & _
						" vl_BC_ICMS_ST, vl_ICMS_ST," & _
						" ncm, cst," & _
						" data_ult_movimento, sequencia, ean, produto_xml, " & _
                        " vl_ipi, aliq_ipi, aliq_icms " & _
						") VALUES (" & _
						"'" & id_estoque & "'" & _
						",'" & Trim(.fabricante) & "'" & _
						",'" & Trim(.produto) & "'" & _
						"," & CStr(.qtde) & _
						"," & bd_formata_numero(.preco_fabricante) & _
						"," & bd_formata_numero(.vl_custo2) & _
						"," & bd_formata_numero(.vl_BC_ICMS_ST) & _
						"," & bd_formata_numero(.vl_ICMS_ST) & _
						",'" & Trim(.ncm) & "'" & _
						",'" & Trim(.cst) & "'" & _
						"," & bd_formata_data(.data_ult_movimento) & _
						"," & CStr(i_seq) & _
						",'" & Trim(.ean) & "'" & _
						",'" & Trim(.produto_xml) & "'" & _
                        "," & bd_formata_numero(.vl_ipi) & _
                        "," & bd_formata_numero(.aliq_ipi) & _
                        "," & bd_formata_numero(.aliq_icms) & _
						")"
				cn.Execute(s_sql)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				
				'Log de movimentação do estoque
				if Not grava_log_estoque_v2(r_estoque.usuario, r_estoque.id_nfe_emitente, .fabricante, .produto, .qtde, .qtde, OP_ESTOQUE_LOG_ENTRADA, "", ID_ESTOQUE_VENDA, "", "", "", "", r_estoque.documento, strComplemento, "") then
					msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
					exit function
					end if
				End If
			End With
		Next
	
	r_estoque.id_estoque = id_estoque
	
	estoque_nova_entrada_mercadorias_xml = True
end function


' --------------------------------------------------------------------
'   ESTOQUE REMOVE
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função remove o "lote" de mercadorias do estoque, desde
'   que isso seja possível.
function estoque_remove(byval id_usuario, byval id_estoque, byref info_log, byref msg_erro)
dim s
dim rs
dim n_item
dim s_log_base
dim blnEntradaEspecial
dim strComplemento
dim v_item
dim i
dim id_nfe_emitente

	estoque_remove = False
	msg_erro = ""
	info_log = ""
	s_log_base = ""

	blnEntradaEspecial=False
	strComplemento=""	
	id_estoque = Trim("" & id_estoque)
	
	s = "SELECT * FROM t_ESTOQUE WHERE (id_estoque='" & id_estoque & "')"
	set rs = cn.execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.Eof then
		msg_erro = "Registro de entrada de mercadorias no estoque nº " & id_estoque & " não está cadastrado."
	else
		id_nfe_emitente = rs("id_nfe_emitente")
		if rs("kit") <> 0 then msg_erro = "Não é possível reverter o cadastramento de kits após a gravação!!"
		s_log_base = "Exclusão do registro de estoque (" & id_estoque & ")" & _
					 " do fabricante " & Trim("" & rs("fabricante")) & _
					 ", entrada em " & formata_data(rs("data_entrada"))
		
		s = formata_hhnnss_para_hh_nn_ss(Trim("" & rs("hora_entrada")))
		if s <> "" then s = " - " & s					 
		s_log_base = s_log_base & s
		
		s_log_base = s_log_base & ", documento " & Trim("" & rs("documento")) & _
					 ", cadastrado por " & Trim("" & rs("usuario"))
		
		if rs("entrada_especial") <> 0 then	
			blnEntradaEspecial=True
			s_log_base = s_log_base & ", registrado como entrada especial"
			end if
		end if
	if rs.State <> 0 then rs.Close	
		
'	ERRO!!
	if msg_erro <> "" then exit function

	if blnEntradaEspecial then strComplemento="ENTRADA_ESPECIAL"
	
    s = "SELECT fabricante, produto, qtde_utilizada FROM t_ESTOQUE_ITEM WHERE" & _
        " (id_estoque='" & id_estoque & "') AND (qtde_utilizada > 0)"
	set rs = cn.execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	n_item = 0
	do while Not rs.Eof
		n_item = n_item + 1
		msg_erro = texto_add_br(msg_erro)
		msg_erro = msg_erro & CStr(n_item) & ") o produto " & Trim("" & rs("produto")) & " já teve " & Trim("" & rs("qtde_utilizada")) & " unidades utilizadas."
		rs.movenext
		Loop
	if rs.State <> 0 then rs.Close

'	ERRO!!
	if msg_erro <> "" then exit function
	
'	INFORMAÇÕES PARA O LOG
	redim v_item(0)
	set v_item(0)  = new cl_ITEM_ENTRADA_ESTOQUE

	s = "SELECT fabricante, produto, qtde FROM t_ESTOQUE_ITEM WHERE" & _
		" (id_estoque='" & id_estoque & "') ORDER BY sequencia"
	set rs = cn.execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	do while Not rs.Eof
		if Trim(v_item(Ubound(v_item)).produto)<>"" then
			redim preserve v_item(Ubound(v_item)+1)
			set v_item(ubound(v_item)) = New cl_ITEM_ENTRADA_ESTOQUE
			end if
		with v_item(Ubound(v_item))
			.fabricante   = rs("fabricante")
			.produto      = rs("produto")
			.qtde         = rs("qtde")
			end with
	
		info_log = info_log & log_estoque_monta_decremento(rs("qtde"), "", Trim("" & rs("produto")))
	
		rs.movenext
		loop
	if rs.State <> 0 then rs.Close

'	GRAVA LOG NO BD
'	Grava após fechar o recordset da consulta para evitar o erro:
'		Microsoft OLE DB Provider for SQL Server error '80004005' 
'		Cannot create new connection because in manual or distributed transaction mode.
	for i = Lbound(v_item) to Ubound(v_item)
		'Log de movimentação do estoque
		with v_item(i)
			if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, .fabricante, .produto, .qtde, .qtde, OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE, ID_ESTOQUE_VENDA, "", "", "", "", "", "", strComplemento, "") then
				msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
				exit function
				end if
			end with
		next

'	TENTA ELIMINAR OS REGISTROS DA IMPORTAÇÃO VIA XML, SE HOUVER

    s = "DELETE FROM t_ESTOQUE_XML_ITEM WHERE"  & _
		" (id_estoque_xml in (SELECT id FROM t_ESTOQUE_XML WHERE "& _
		" (id_estoque = '" & id_estoque & "'))) "
	cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

    s = "DELETE FROM t_ESTOQUE_XML WHERE" & _
		" (id_estoque = '" & id_estoque & "') "
	cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if


'	TENTA ELIMINAR A LISTA DE PRODUTOS
    s = "DELETE FROM t_ESTOQUE_ITEM WHERE" & _
		" (id_estoque = '" & id_estoque & "') AND (qtde_utilizada = 0)"
	cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
'	VERIFICA SE A LISTA COMPLETA FOI REMOVIDA
    s = "SELECT fabricante, produto, qtde_utilizada FROM t_ESTOQUE_ITEM WHERE" & _
		" (id_estoque = '" & id_estoque & "') ORDER BY sequencia"
	set rs = cn.execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

    n_item = 0
	do while Not rs.Eof
		n_item = n_item + 1
		msg_erro = texto_add_br(msg_erro)
		msg_erro = msg_erro & CStr(n_item) & ") o produto " & Trim("" & rs("produto")) & " já teve " & Trim("" & rs("qtde_utilizada")) & " unidades utilizadas."
		rs.movenext
		Loop
	if rs.State <> 0 then rs.Close	
	
'	ERRO!!
	if msg_erro <> "" then exit function
			
'	CONSEGUIU REMOVER A LISTA DE PRODUTOS, AGORA REMOVE O REGISTRO DE ENTRADA NO ESTOQUE
    s = "DELETE FROM t_ESTOQUE WHERE" & _
		" (id_estoque = '" & id_estoque & "')"
	cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	info_log = s_log_base & ":" & info_log
	
	estoque_remove = True
end function



' --------------------------------------------------------------------
'   ESTOQUE ATUALIZA
'   Retorno da função:
'      False - Ocorreu falha ao tentar alterar os dados do estoque.
'      True - Conseguiu alterar os dados do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função altera os dados cadastrais do "lote" de mercadorias 
'	do estoque e/ou a quantidade dos produtos cadastrados.
function estoque_atualiza(byval id_usuario, byval r_estoque, byval v_item, byref info_log, byref msg_erro)
dim s
dim i
dim j
dim i_ref
dim achou
dim i_seq
dim n_item
dim n_movto
dim rs
dim s_log
dim gravou_item
dim r_estoque_bd
dim v_item_bd
dim qtde_aux
dim qtde_utilizada_aux
dim vl_BC_ICMS_ST_aux
dim vl_ICMS_ST_aux
dim preco_fabricante_aux
dim vl_custo2_aux
dim ncm_aux
dim cst_aux
dim aliq_ipi_aux
dim aliq_icms_aux
dim vl_ipi_aux
dim qtde_delta
dim vLog1()
dim vLog2()
dim campos_a_omitir
dim strComplemento

	estoque_atualiza = False
	msg_erro = ""
	info_log = ""

	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

	if Not le_estoque(r_estoque.id_estoque, r_estoque_bd, msg_erro) then exit function
	if Not le_estoque_item(r_estoque.id_estoque, v_item_bd, msg_erro) then exit function

	strComplemento=""
	if Cstr(r_estoque.entrada_especial) <> Cstr(0) then strComplemento="ENTRADA_ESPECIAL"
			
	gravou_item = False
	s_log = ""
	campos_a_omitir = ""
	
'	PRODUTOS NOVOS NA LISTA
'	=======================
'	COLOCA OS PRODUTOS NOVOS NO FINAL, SENDO QUE OS ÍNDICES DE SEQUENCIAÇÃO SERÃO 
'	COMPACTADOS MAIS ADIANTE.
	i_seq = UBound(v_item)
	for i = Lbound(v_item) to Ubound(v_item)
		if Trim(v_item(i).produto) <> "" then
			achou = False
			for j = Lbound(v_item_bd) to Ubound(v_item_bd)
				if Trim(v_item(i).produto) = Trim(v_item_bd(j).produto) then
					achou = True
					exit for
					end if
				next
			
		'	É UM PRODUTO NOVO NA LISTA
			If Not achou Then
				i_seq = i_seq + 1
				with v_item(i)
					s = "SELECT * FROM t_PRODUTO WHERE" & _
						" (fabricante='" & r_estoque.fabricante & "')" 
					if IsEAN(.produto) then
						s = s & " AND (ean='" & .produto & "')"
					else
						s = s & " AND (produto='" & .produto & "')"
						end if

					if rs.State <> 0 then rs.Close
					rs.Open s, cn
					if Err <> 0 then
						msg_erro = Cstr(Err) & ": " & Err.Description
						exit function
						end if
					
					if rs.Eof then
						msg_erro = "Produto " & .produto & " do fabricante " & r_estoque.fabricante & " não está cadastrado."
						if rs.State <> 0 then rs.Close
						exit function
					else
						.fabricante = Trim(r_estoque.fabricante)
					'	CARREGA CÓDIGO INTERNO DO PRODUTO
						.produto = Trim("" & rs("produto"))
						if (.preco_fabricante = 0) And (rs("preco_fabricante") <> 0) then .preco_fabricante = rs("preco_fabricante")
						if (.vl_custo2 = 0) And (rs("vl_custo2") <> 0) then .vl_custo2 = rs("vl_custo2")
						end if
					
					s = "INSERT INTO T_ESTOQUE_ITEM" & _
						" (id_estoque, fabricante, produto, qtde, preco_fabricante, vl_custo2," & _
						" vl_BC_ICMS_ST, vl_ICMS_ST," & _
						" ncm, cst," & _
						" data_ult_movimento, sequencia)" & _
						" VALUES (" & _
						"'" & r_estoque.id_estoque & "'" & _
						",'" & .fabricante & "'" & _
						",'" & .produto & "'" & _
						"," & CStr(.qtde) & _
						"," & bd_formata_numero(.preco_fabricante) & _
						"," & bd_formata_numero(.vl_custo2) & _
						"," & bd_formata_numero(.vl_BC_ICMS_ST) & _
						"," & bd_formata_numero(.vl_ICMS_ST) & _
						",'" & Trim(.ncm) & "'" & _
						",'" & Trim(.cst) & "'" & _
						"," & bd_formata_data(Date) & _
						"," & CStr(i_seq) & _
						")"
					cn.Execute(s)
					if Err <> 0 then
						msg_erro=Cstr(Err) & ": " & Err.Description
						exit function
						end if
					
					gravou_item = True
				'	INFORMAÇÕES P/ O LOG
					s_log = s_log & log_estoque_monta_incremento(.qtde, "", .produto)
					
					'Log de movimentação do estoque
					if Not grava_log_estoque_v2(id_usuario, r_estoque.id_nfe_emitente, .fabricante, .produto, .qtde, .qtde, OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM, "", ID_ESTOQUE_VENDA, "", "", "", "", Trim(r_estoque.documento), strComplemento, "") then
						msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
						exit function
						end if
					end with
				end if
			end if
		next
		
		
'	PRODUTOS ALTERADOS/EXCLUÍDOS
'	============================
'	PRODUTOS ALTERADOS: VERIFICA SE A NOVA QUANTIDADE ESTÁ CONSISTENTE C/ RELAÇÃO À QTDE UTILIZADA.
'	PRODUTOS EXCLUÍDOS: NÃO EXISTE EXCLUSÃO DIRETA, DEVE-SE CADASTRAR COMO UMA ALTERAÇÃO P/ 
'						"QUANTIDADE = 0", E MAIS ADIANTE HAVERÁ UMA ROTINA QUE REMOVE OS REGISTROS 
'						C/ QUANTIDADE ZERADA.
	for i = Lbound(v_item) to Ubound(v_item)
		if Trim(v_item(i).produto) <> "" then
			achou = False
			for j = Lbound(v_item_bd) to Ubound(v_item_bd)
				if Trim(v_item(i).produto) = Trim(v_item_bd(j).produto) then
					achou = True
					i_ref = j
					exit for
					end if
				next
		
		'	É UMA ALTERAÇÃO NO PRODUTO?
			if achou then
				if (v_item(i).qtde <> v_item_bd(i_ref).qtde) Or _
				   (v_item(i).vl_BC_ICMS_ST <> v_item_bd(i_ref).vl_BC_ICMS_ST) Or _
				   (v_item(i).vl_ICMS_ST <> v_item_bd(i_ref).vl_ICMS_ST) Or _
				   (v_item(i).preco_fabricante <> v_item_bd(i_ref).preco_fabricante) Or _
				   (v_item(i).vl_custo2 <> v_item_bd(i_ref).vl_custo2) Or _
				   (v_item(i).ncm <> v_item_bd(i_ref).ncm) Or _
				   (v_item(i).cst <> v_item_bd(i_ref).cst) then
					with v_item(i)
						s = "SELECT * FROM t_ESTOQUE_ITEM WHERE" & _
							" (id_estoque='" & r_estoque.id_estoque & "')" & _
							" AND (fabricante='" & .fabricante & "')" & _
							" AND (produto='" & .produto & "')"
						if rs.State <> 0 then rs.Close
						rs.Open s, cn
						if Err <> 0 then
							msg_erro=Cstr(Err) & ": " & Err.Description
							exit function
							end if
						
						if rs.Eof then
							msg_erro = "Não foi encontrado o registro do produto " & .produto & " do lote do estoque nº " & r_estoque.id_estoque
							if rs.State <> 0 then rs.Close
							exit function
						else
							qtde_aux = rs("qtde")
							qtde_utilizada_aux = rs("qtde_utilizada")
							vl_BC_ICMS_ST_aux = rs("vl_BC_ICMS_ST")
							vl_ICMS_ST_aux = rs("vl_ICMS_ST")
							preco_fabricante_aux = rs("preco_fabricante")
							vl_custo2_aux = rs("vl_custo2")
							ncm_aux = Trim("" & rs("ncm"))
							cst_aux = Trim("" & rs("cst"))
                            aliq_ipi_aux = rs("aliq_ipi")
                            aliq_icms_aux = rs("aliq_icms")
                            vl_ipi_aux = rs("vl_ipi")
						'	QUANTIDADE ESTÁ CONSISTENTE, ENTÃO PODE ATUALIZAR O REGISTRO
							if .qtde < rs("qtde_utilizada") then
								msg_erro = texto_add_br(msg_erro)
								msg_erro = msg_erro & "A quantidade do produto " & .produto & " NÃO foi alterada de " & CStr(qtde_aux) & " para " & CStr(.qtde) & ", pois " & CStr(qtde_utilizada_aux) & " unidades já foram utilizadas!!"
							else
								rs("qtde") = .qtde
								rs("vl_BC_ICMS_ST") = converte_numero(.vl_BC_ICMS_ST)
								rs("vl_ICMS_ST") = converte_numero(.vl_ICMS_ST)
								rs("preco_fabricante") = converte_numero(.preco_fabricante)
								rs("vl_custo2") = converte_numero(.vl_custo2)
								rs("ncm") = Trim(.ncm)
								rs("cst") = Trim(.cst)
                                rs("aliq_ipi") = Trim(.aliq_ipi)
                                rs("aliq_icms") = Trim(.aliq_icms)
                                rs("vl_ipi") = Trim(.vl_ipi)
								rs("data_ult_movimento") = Date
								rs.Update
								if Err <> 0 then
									msg_erro=Cstr(Err) & ": " & Err.Description
									exit function
									end if
								
								gravou_item = True
							'	INFORMAÇÕES P/ O LOG
								If qtde_aux > .qtde Then
									qtde_delta = qtde_aux - .qtde
									if qtde_delta <> 0 then
										if s_log <> "" then s_log = s_log & ";"
										s_log = s_log & log_estoque_monta_decremento((qtde_aux - .qtde), "", .produto)
										'Log de movimentação do estoque
										if Not grava_log_estoque_v2(id_usuario, r_estoque.id_nfe_emitente, rs("fabricante"), rs("produto"), qtde_delta, qtde_delta, OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA, ID_ESTOQUE_VENDA, "", "", "", "", "", Trim(r_estoque.documento), strComplemento, "") then
											msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
											exit function
											end if
										end if
								Else
									qtde_delta = .qtde - qtde_aux
									if qtde_delta <> 0 then
										if s_log <> "" then s_log = s_log & ";"
										s_log = s_log & log_estoque_monta_incremento((.qtde - qtde_aux), "", .produto)
										'Log de movimentação do estoque
										if Not grava_log_estoque_v2(id_usuario, r_estoque.id_nfe_emitente, rs("fabricante"), rs("produto"), qtde_delta, qtde_delta, OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA, "", ID_ESTOQUE_VENDA, "", "", "", "", Trim(r_estoque.documento), strComplemento, "") then
											msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
											exit function
											end if
										end if
									End If
								
								if converte_numero(preco_fabricante_aux) <> converte_numero(.preco_fabricante) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": preco_fabricante: " & formata_moeda(preco_fabricante_aux) & " => " & formata_moeda(.preco_fabricante)
									end if

								if converte_numero(vl_custo2_aux) <> converte_numero(.vl_custo2) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_custo2: " & formata_moeda(vl_custo2_aux) & " => " & formata_moeda(.vl_custo2)
									end if

								if converte_numero(vl_BC_ICMS_ST_aux) <> converte_numero(.vl_BC_ICMS_ST) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_BC_ICMS_ST: " & formata_moeda(vl_BC_ICMS_ST_aux) & " => " & formata_moeda(.vl_BC_ICMS_ST)
									end if
								
								if converte_numero(vl_ICMS_ST_aux) <> converte_numero(.vl_ICMS_ST) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_ICMS_ST: " & formata_moeda(vl_ICMS_ST_aux) & " => " & formata_moeda(.vl_ICMS_ST)
									end if
								
								if Trim(ncm_aux) <> Trim(.ncm) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": NCM: " & Trim(ncm_aux) & " => " & Trim(.ncm)
									end if
								
								if Trim(cst_aux) <> Trim(.cst) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": CST: " & Trim(cst_aux) & " => " & Trim(.cst)
									end if

								if converte_numero(aliq_ipi_aux) <> converte_numero(.aliq_ipi) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": aliq_ipi: " & formata_numero(aliq_ipi_aux, 0) & " => " & formata_numero(.aliq_ipi)
									end if

                                if converte_numero(aliq_ipi_aux) <> converte_numero(.aliq_ipi) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": aliq_ipi: " & formata_numero(aliq_ipi_aux, 0) & " => " & formata_numero(.aliq_ipi, 0)
									end if
            
                                if converte_numero(aliq_icms_aux) <> converte_numero(.aliq_icms) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": aliq_icms: " & formata_numero(aliq_icms_aux, 0) & " => " & formata_numero(.aliq_icms, 0)
									end if

                                if converte_numero(vl_ipi_aux) <> converte_numero(.vl_ipi) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_ipi: " & formata_moeda(vl_ipi_aux) & " => " & formata_moeda(.vl_ipi)
									end if

								end if
							end if
						end with
					end if
				end if
			end if
		next


'	INFORMAÇÕES P/ O LOG
	if s_log <> "" then s_log = "alteração nos dados de produtos:" & s_log


'	PRODUTOS QUE ALTERARAM A QUANTIDADE P/ ZERO: A QTDE IGUAL A ZERO ASSEGURA QUE 
'	NENHUMA OPERAÇÃO DE SAÍDA DE ESTOQUE SERÁ FEITA NESSES REGISTROS, ENTÃO PODEM 
'	SER ELIMINADOS DIRETAMENTE.
	s = "DELETE FROM t_ESTOQUE_ITEM WHERE" & _
		" (id_estoque='" & Trim(r_estoque.id_estoque) & "') AND (qtde=0) AND (qtde_utilizada=0)"
	cn.execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
  '	SE NÃO RESTOU NENHUM PRODUTO, PODE REMOVER O REGISTRO DE ENTRADA NO ESTOQUE
	s = "SELECT COUNT(*) AS total FROM t_ESTOQUE_ITEM WHERE" & _
		" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
	n_item = -1
	if rs.State <> 0 then rs.Close
	rs.Open s, cn 
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	if Not rs.Eof then
		if IsNumeric(rs("total")) then n_item = CLng(rs("total"))
		end if
	
'	VERIFICA SE HÁ VÍNCULOS NA TABELA DE MOVIMENTOS DO ESTOQUE
'	CASO SIM, MANTÉM O REGISTRO DE ENTRADA NO ESTOQUE P/ FINS DE HISTÓRICO
	s = "SELECT COUNT(*) AS total FROM t_ESTOQUE_MOVIMENTO WHERE" & _
		" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
	n_movto = -1
	if rs.State <> 0 then rs.Close
	rs.Open s, cn
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	if Not rs.Eof then
		if IsNumeric(rs("total")) then n_movto = CLng(rs("total"))
		end if
	
	If (n_item = 0) And (n_movto = 0) Then
	'	EXCLUI O REGISTRO DE ESTOQUE!!
		s = "DELETE FROM t_ESTOQUE WHERE" & _
			" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
		cn.execute(s)
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		
	'	VERIFICA SE CONSEGUIU EXCLUIR 
		s = "SELECT id_estoque FROM t_ESTOQUE WHERE" & _
			" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
		if rs.State <> 0 then rs.Close
		rs.Open s, cn
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		
		if Not rs.Eof then
			msg_erro="Falha ao tentar remover o registro do lote do estoque nº " & r_estoque.id_estoque
			exit function
			end if
	
	else
	'	ATUALIZA INFORMAÇÕES DO REGISTRO DE ENTRADA NO ESTOQUE
		if gravou_item Or _
		   (Trim(r_estoque.documento)<>Trim(r_estoque_bd.documento)) Or _
		   (Trim(r_estoque.obs)<>Trim(r_estoque_bd.obs)) Or _
		   (r_estoque.entrada_especial<>r_estoque_bd.entrada_especial) then
			s = "SELECT * FROM t_ESTOQUE WHERE" & _
				" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
				
			if rs.Eof then
				msg_erro="O registro do lote do estoque nº " & r_estoque.id_estoque & " não foi encontrado."
				if rs.State <> 0 then rs.Close
				exit function
				end if

			log_via_vetor_carrega_do_recordset rs, vLog1, campos_a_omitir
			rs("documento") = Trim(r_estoque.documento)
			rs("obs") = Trim(r_estoque.obs)
			rs("entrada_especial") = r_estoque.entrada_especial
			if gravou_item then rs("data_ult_movimento") = Date
			rs.Update

			log_via_vetor_carrega_do_recordset rs, vLog2, campos_a_omitir
			s = log_via_vetor_monta_alteracao(vLog1, vLog2)
			if s <> "" then 
				if s_log <> "" then s_log = "; " & s_log
				s_log = s & s_log
				end if
			end if
		
	'	COMPACTA A SEQUÊNCIA
		i_seq = 0
		s = "SELECT * FROM t_ESTOQUE_ITEM WHERE (id_estoque='" & Trim(r_estoque.id_estoque) & "') ORDER BY sequencia"
		if rs.State <> 0 then rs.Close
		rs.Open s, cn
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		
		do while Not rs.Eof
			i_seq = i_seq + 1
			rs("sequencia") = i_seq
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			rs.movenext
			loop
		end if
	
	
	if msg_erro <> "" then exit function
	
	if s_log <> "" then
		s_log = "Alteração no registro de estoque (" & Trim(r_estoque.id_estoque) & ")" & _
				" do fabricante " & Trim(r_estoque.fabricante) & "," & _
				" entrada em " & formata_data(r_estoque.data_entrada) & "," & _
				" documento " & Trim(r_estoque.documento) & ": " & _
				s_log
		info_log = s_log
		end if
		
	estoque_atualiza = True
end function


' --------------------------------------------------------------------
'   ESTOQUE ATUALIZA AGIO
'   Retorno da função:
'      False - Ocorreu falha ao tentar alterar os dados do estoque.
'      True - Conseguiu alterar os dados do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função altera os dados cadastrais do "lote" de mercadorias 
'	do estoque e/ou a quantidade dos produtos cadastrados.
function estoque_atualiza_agio(byval id_usuario, byval r_estoque, byval v_item, byref info_log, byref msg_erro)
dim s
dim i
dim j
dim i_ref
dim achou
dim i_seq
dim n_item
dim n_movto
dim rs
dim s_log
dim gravou_item
dim r_estoque_bd
dim v_item_bd
dim qtde_aux
dim qtde_utilizada_aux
dim vl_BC_ICMS_ST_aux
dim vl_ICMS_ST_aux
dim preco_fabricante_aux
dim vl_custo2_aux
dim ncm_aux
dim cst_aux
dim aliq_ipi_aux
dim vl_ipi_aux
dim aliq_icms_aux
dim qtde_delta
dim vLog1()
dim vLog2()
dim campos_a_omitir
dim strComplemento

	estoque_atualiza_agio = False
	msg_erro = ""
	info_log = ""

	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

	if Not le_estoque_agio(r_estoque.id_estoque, r_estoque_bd, msg_erro) then exit function
	if Not le_estoque_item_xml(r_estoque.id_estoque, v_item_bd, msg_erro) then exit function

	strComplemento=""
	if Cstr(r_estoque.entrada_especial) <> Cstr(0) then strComplemento="ENTRADA_ESPECIAL"
			
	gravou_item = False
	s_log = ""
	campos_a_omitir = ""
	
'	PRODUTOS NOVOS NA LISTA
'	=======================
'	COLOCA OS PRODUTOS NOVOS NO FINAL, SENDO QUE OS ÍNDICES DE SEQUENCIAÇÃO SERÃO 
'	COMPACTADOS MAIS ADIANTE.
	i_seq = UBound(v_item)
	for i = Lbound(v_item) to Ubound(v_item)
		if Trim(v_item(i).produto) <> "" then
			achou = False
			for j = Lbound(v_item_bd) to Ubound(v_item_bd)
				if Trim(v_item(i).produto) = Trim(v_item_bd(j).produto) then
					achou = True
					exit for
					end if
				next
			
		'	É UM PRODUTO NOVO NA LISTA
			If Not achou Then
				i_seq = i_seq + 1
				with v_item(i)
					s = "SELECT * FROM t_PRODUTO WHERE" & _
						" (fabricante='" & r_estoque.fabricante & "')" 
					if IsEAN(.produto) then
						s = s & " AND (ean='" & .produto & "')"
					else
						s = s & " AND (produto='" & .produto & "')"
						end if

					if rs.State <> 0 then rs.Close
					rs.Open s, cn
					if Err <> 0 then
						msg_erro = Cstr(Err) & ": " & Err.Description
						exit function
						end if
					
					if rs.Eof then
						msg_erro = "Produto " & .produto & " do fabricante " & r_estoque.fabricante & " não está cadastrado."
						if rs.State <> 0 then rs.Close
						exit function
					else
						.fabricante = Trim(r_estoque.fabricante)
					'	CARREGA CÓDIGO INTERNO DO PRODUTO
						.produto = Trim("" & rs("produto"))
						if (.preco_fabricante = 0) And (rs("preco_fabricante") <> 0) then .preco_fabricante = rs("preco_fabricante")
						if (.vl_custo2 = 0) And (rs("vl_custo2") <> 0) then .vl_custo2 = rs("vl_custo2")
						end if
					
					s = "INSERT INTO T_ESTOQUE_ITEM" & _
						" (id_estoque, fabricante, produto, qtde, preco_fabricante, vl_custo2," & _
						" vl_BC_ICMS_ST, vl_ICMS_ST," & _
						" ncm, cst," & _
						" data_ult_movimento, sequencia,"  & _
                        " aliq_ipi, aliq_icms, vl_ipi)" & _
						" VALUES (" & _
						"'" & r_estoque.id_estoque & "'" & _
						",'" & .fabricante & "'" & _
						",'" & .produto & "'" & _
						"," & CStr(.qtde) & _
						"," & bd_formata_numero(.preco_fabricante) & _
						"," & bd_formata_numero(.vl_custo2) & _
						"," & bd_formata_numero(.vl_BC_ICMS_ST) & _
						"," & bd_formata_numero(.vl_ICMS_ST) & _
						",'" & Trim(.ncm) & "'" & _
						",'" & Trim(.cst) & "'" & _
						"," & bd_formata_data(Date) & _
						"," & CStr(i_seq) & _
                        "," & bd_formata_numero(.aliq_ipi) & _
                        "," & bd_formata_numero(.aliq_icms) & _
                        "," & bd_formata_numero(.vl_ipi) & _
						")"
					cn.Execute(s)
					if Err <> 0 then
						msg_erro=Cstr(Err) & ": " & Err.Description
						exit function
						end if
					
					gravou_item = True
				'	INFORMAÇÕES P/ O LOG
					s_log = s_log & log_estoque_monta_incremento(.qtde, "", .produto)
					
					'Log de movimentação do estoque
					if Not grava_log_estoque_v2(id_usuario, r_estoque.id_nfe_emitente, .fabricante, .produto, .qtde, .qtde, OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM, "", ID_ESTOQUE_VENDA, "", "", "", "", Trim(r_estoque.documento), strComplemento, "") then
						msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
						exit function
						end if
					end with
				end if
			end if
		next
		
		
'	PRODUTOS ALTERADOS/EXCLUÍDOS
'	============================
'	PRODUTOS ALTERADOS: VERIFICA SE A NOVA QUANTIDADE ESTÁ CONSISTENTE C/ RELAÇÃO À QTDE UTILIZADA.
'	PRODUTOS EXCLUÍDOS: NÃO EXISTE EXCLUSÃO DIRETA, DEVE-SE CADASTRAR COMO UMA ALTERAÇÃO P/ 
'						"QUANTIDADE = 0", E MAIS ADIANTE HAVERÁ UMA ROTINA QUE REMOVE OS REGISTROS 
'						C/ QUANTIDADE ZERADA.
	for i = Lbound(v_item) to Ubound(v_item)
		if Trim(v_item(i).produto) <> "" then
			achou = False
			for j = Lbound(v_item_bd) to Ubound(v_item_bd)
				if Trim(v_item(i).produto) = Trim(v_item_bd(j).produto) then
					achou = True
					i_ref = j
					exit for
					end if
				next
		
		'	É UMA ALTERAÇÃO NO PRODUTO?
			if achou then
				if (v_item(i).qtde <> v_item_bd(i_ref).qtde) Or _
				   (v_item(i).vl_BC_ICMS_ST <> v_item_bd(i_ref).vl_BC_ICMS_ST) Or _
				   (v_item(i).vl_ICMS_ST <> v_item_bd(i_ref).vl_ICMS_ST) Or _
				   (v_item(i).preco_fabricante <> v_item_bd(i_ref).preco_fabricante) Or _
				   (v_item(i).vl_custo2 <> v_item_bd(i_ref).vl_custo2) Or _
				   (v_item(i).ncm <> v_item_bd(i_ref).ncm) Or _
				   (v_item(i).cst <> v_item_bd(i_ref).cst) Or _
				   (v_item(i).aliq_ipi <> v_item_bd(i_ref).aliq_ipi) Or _
				   (v_item(i).aliq_icms <> v_item_bd(i_ref).aliq_icms) Or _
				   (v_item(i).vl_ipi <> v_item_bd(i_ref).vl_ipi) then
					with v_item(i)
						s = "SELECT * FROM t_ESTOQUE_ITEM WHERE" & _
							" (id_estoque='" & r_estoque.id_estoque & "')" & _
							" AND (fabricante='" & .fabricante & "')" & _
							" AND (produto='" & .produto & "')"
						if rs.State <> 0 then rs.Close
						rs.Open s, cn
						if Err <> 0 then
							msg_erro=Cstr(Err) & ": " & Err.Description
							exit function
							end if
						
						if rs.Eof then
							msg_erro = "Não foi encontrado o registro do produto " & .produto & " do lote do estoque nº " & r_estoque.id_estoque
							if rs.State <> 0 then rs.Close
							exit function
						else
							qtde_aux = rs("qtde")
							qtde_utilizada_aux = rs("qtde_utilizada")
							vl_BC_ICMS_ST_aux = rs("vl_BC_ICMS_ST")
							vl_ICMS_ST_aux = rs("vl_ICMS_ST")
							preco_fabricante_aux = rs("preco_fabricante")
							vl_custo2_aux = rs("vl_custo2")
							ncm_aux = Trim("" & rs("ncm"))
							cst_aux = Trim("" & rs("cst"))
                            aliq_ipi_aux = rs("aliq_ipi")
							vl_ipi_aux = rs("vl_ipi")
							aliq_icms_aux = rs("aliq_icms")
						'	QUANTIDADE ESTÁ CONSISTENTE, ENTÃO PODE ATUALIZAR O REGISTRO
							if .qtde < rs("qtde_utilizada") then
								msg_erro = texto_add_br(msg_erro)
								msg_erro = msg_erro & "A quantidade do produto " & .produto & " NÃO foi alterada de " & CStr(qtde_aux) & " para " & CStr(.qtde) & ", pois " & CStr(qtde_utilizada_aux) & " unidades já foram utilizadas!!"
							else
								rs("qtde") = .qtde
								rs("vl_BC_ICMS_ST") = converte_numero(.vl_BC_ICMS_ST)
								rs("vl_ICMS_ST") = converte_numero(.vl_ICMS_ST)
								rs("preco_fabricante") = converte_numero(.preco_fabricante)
								rs("vl_custo2") = converte_numero(.vl_custo2)
								rs("ncm") = Trim(.ncm)
								rs("cst") = Trim(.cst)
								rs("aliq_ipi") = converte_numero(.aliq_ipi)
								rs("vl_ipi") = converte_numero(.vl_ipi)
								rs("aliq_icms") = converte_numero(.aliq_icms)
								rs("data_ult_movimento") = Date
								rs.Update
								if Err <> 0 then
									msg_erro=Cstr(Err) & ": " & Err.Description
									exit function
									end if
								
								gravou_item = True
							'	INFORMAÇÕES P/ O LOG
								If qtde_aux > .qtde Then
									qtde_delta = qtde_aux - .qtde
									if qtde_delta <> 0 then
										if s_log <> "" then s_log = s_log & ";"
										s_log = s_log & log_estoque_monta_decremento((qtde_aux - .qtde), "", .produto)
										'Log de movimentação do estoque
										if Not grava_log_estoque_v2(id_usuario, r_estoque.id_nfe_emitente, rs("fabricante"), rs("produto"), qtde_delta, qtde_delta, OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA, ID_ESTOQUE_VENDA, "", "", "", "", "", Trim(r_estoque.documento), strComplemento, "") then
											msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
											exit function
											end if
										end if
								Else
									qtde_delta = .qtde - qtde_aux
									if qtde_delta <> 0 then
										if s_log <> "" then s_log = s_log & ";"
										s_log = s_log & log_estoque_monta_incremento((.qtde - qtde_aux), "", .produto)
										'Log de movimentação do estoque
										if Not grava_log_estoque_v2(id_usuario, r_estoque.id_nfe_emitente, rs("fabricante"), rs("produto"), qtde_delta, qtde_delta, OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA, "", ID_ESTOQUE_VENDA, "", "", "", "", Trim(r_estoque.documento), strComplemento, "") then
											msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
											exit function
											end if
										end if
									End If
								
								if converte_numero(preco_fabricante_aux) <> converte_numero(.preco_fabricante) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": preco_fabricante: " & formata_moeda(preco_fabricante_aux) & " => " & formata_moeda(.preco_fabricante)
									end if

								if converte_numero(vl_custo2_aux) <> converte_numero(.vl_custo2) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_custo2: " & formata_moeda(vl_custo2_aux) & " => " & formata_moeda(.vl_custo2)
									end if

								if converte_numero(vl_BC_ICMS_ST_aux) <> converte_numero(.vl_BC_ICMS_ST) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_BC_ICMS_ST: " & formata_moeda(vl_BC_ICMS_ST_aux) & " => " & formata_moeda(.vl_BC_ICMS_ST)
									end if
								
								if converte_numero(vl_ICMS_ST_aux) <> converte_numero(.vl_ICMS_ST) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_ICMS_ST: " & formata_moeda(vl_ICMS_ST_aux) & " => " & formata_moeda(.vl_ICMS_ST)
									end if
								
								if Trim(ncm_aux) <> Trim(.ncm) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": NCM: " & Trim(ncm_aux) & " => " & Trim(.ncm)
									end if
								
								if Trim(cst_aux) <> Trim(.cst) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": CST: " & Trim(cst_aux) & " => " & Trim(.cst)
									end if

								if converte_numero(aliq_ipi_aux) <> converte_numero(.aliq_ipi) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": aliq_ipi: " & formata_moeda(aliq_ipi_aux) & " => " & formata_moeda(.aliq_ipi)
									end if

								if converte_numero(vl_ipi_aux) <> converte_numero(.vl_ipi) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_ipi: " & formata_moeda(vl_ipi_aux) & " => " & formata_moeda(.vl_ipi)
									end if

								if converte_numero(aliq_icms_aux) <> converte_numero(.aliq_icms) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": aliq_icms: " & formata_moeda(aliq_icms_aux) & " => " & formata_moeda(.aliq_icms)
									end if

								end if
							end if
						end with
					end if
				end if
			end if
		next


'	INFORMAÇÕES P/ O LOG
	if s_log <> "" then s_log = "alteração nos dados de produtos:" & s_log


'	PRODUTOS QUE ALTERARAM A QUANTIDADE P/ ZERO: A QTDE IGUAL A ZERO ASSEGURA QUE 
'	NENHUMA OPERAÇÃO DE SAÍDA DE ESTOQUE SERÁ FEITA NESSES REGISTROS, ENTÃO PODEM 
'	SER ELIMINADOS DIRETAMENTE.
	s = "DELETE FROM t_ESTOQUE_ITEM WHERE" & _
		" (id_estoque='" & Trim(r_estoque.id_estoque) & "') AND (qtde=0) AND (qtde_utilizada=0)"
	cn.execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
  '	SE NÃO RESTOU NENHUM PRODUTO, PODE REMOVER O REGISTRO DE ENTRADA NO ESTOQUE
	s = "SELECT COUNT(*) AS total FROM t_ESTOQUE_ITEM WHERE" & _
		" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
	n_item = -1
	if rs.State <> 0 then rs.Close
	rs.Open s, cn 
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	if Not rs.Eof then
		if IsNumeric(rs("total")) then n_item = CLng(rs("total"))
		end if
	
'	VERIFICA SE HÁ VÍNCULOS NA TABELA DE MOVIMENTOS DO ESTOQUE
'	CASO SIM, MANTÉM O REGISTRO DE ENTRADA NO ESTOQUE P/ FINS DE HISTÓRICO
	s = "SELECT COUNT(*) AS total FROM t_ESTOQUE_MOVIMENTO WHERE" & _
		" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
	n_movto = -1
	if rs.State <> 0 then rs.Close
	rs.Open s, cn
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	if Not rs.Eof then
		if IsNumeric(rs("total")) then n_movto = CLng(rs("total"))
		end if
	
	If (n_item = 0) And (n_movto = 0) Then
	'	EXCLUI O REGISTRO DE ESTOQUE!!
		s = "DELETE FROM t_ESTOQUE WHERE" & _
			" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
		cn.execute(s)
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		
	'	VERIFICA SE CONSEGUIU EXCLUIR 
		s = "SELECT id_estoque FROM t_ESTOQUE WHERE" & _
			" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
		if rs.State <> 0 then rs.Close
		rs.Open s, cn
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		
		if Not rs.Eof then
			msg_erro="Falha ao tentar remover o registro do lote do estoque nº " & r_estoque.id_estoque
			exit function
			end if
	
	else
	'	ATUALIZA INFORMAÇÕES DO REGISTRO DE ENTRADA NO ESTOQUE
		if gravou_item Or _
		   (Trim(r_estoque.documento)<>Trim(r_estoque_bd.documento)) Or _
		   (Trim(r_estoque.obs)<>Trim(r_estoque_bd.obs)) Or _
		   (r_estoque.entrada_especial<>r_estoque_bd.entrada_especial) Or _
		   (r_estoque.perc_agio<>r_estoque_bd.perc_agio) then
			s = "SELECT * FROM t_ESTOQUE WHERE" & _
				" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
				
			if rs.Eof then
				msg_erro="O registro do lote do estoque nº " & r_estoque.id_estoque & " não foi encontrado."
				if rs.State <> 0 then rs.Close
				exit function
				end if

			log_via_vetor_carrega_do_recordset rs, vLog1, campos_a_omitir
			rs("documento") = Trim(r_estoque.documento)
			rs("obs") = Trim(r_estoque.obs)
			rs("entrada_especial") = r_estoque.entrada_especial
            rs("perc_agio") = r_estoque.perc_agio
			if gravou_item then rs("data_ult_movimento") = Date
			rs.Update

			log_via_vetor_carrega_do_recordset rs, vLog2, campos_a_omitir
			s = log_via_vetor_monta_alteracao(vLog1, vLog2)
			if s <> "" then 
				if s_log <> "" then s_log = "; " & s_log
				s_log = s & s_log
				end if
			end if
		
	'	COMPACTA A SEQUÊNCIA
		i_seq = 0
		s = "SELECT * FROM t_ESTOQUE_ITEM WHERE (id_estoque='" & Trim(r_estoque.id_estoque) & "') ORDER BY sequencia"
		if rs.State <> 0 then rs.Close
		rs.Open s, cn
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		
		do while Not rs.Eof
			i_seq = i_seq + 1
			rs("sequencia") = i_seq
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			rs.movenext
			loop
		end if
	
	
	if msg_erro <> "" then exit function
	
	if s_log <> "" then
		s_log = "Alteração no registro de estoque (" & Trim(r_estoque.id_estoque) & ")" & _
				" do fabricante " & Trim(r_estoque.fabricante) & "," & _
				" entrada em " & formata_data(r_estoque.data_entrada) & "," & _
				" documento " & Trim(r_estoque.documento) & ": " & _
				s_log
		info_log = s_log
		end if
		
	estoque_atualiza_agio = True
end function


' --------------------------------------------------------------------
'   ESTOQUE ATUALIZA XML
'   Retorno da função:
'      False - Ocorreu falha ao tentar alterar os dados do estoque.
'      True - Conseguiu alterar os dados do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função altera os dados cadastrais do "lote" de mercadorias 
'	do estoque e/ou a quantidade dos produtos cadastrados.
function estoque_atualiza_xml(byval id_usuario, byval r_estoque, byval v_item, byref info_log, byref msg_erro)
dim s
dim i
dim j
dim i_ref
dim achou
dim i_seq
dim n_item
dim n_movto
dim rs
dim s_log
dim gravou_item
dim r_estoque_bd
dim v_item_bd
dim qtde_aux
dim qtde_utilizada_aux
dim vl_BC_ICMS_ST_aux
dim vl_ICMS_ST_aux
dim preco_fabricante_aux
dim vl_custo2_aux
dim ncm_aux
dim cst_aux
dim ean_aux
dim aliq_ipi_aux
dim vl_ipi_aux
dim aliq_icms_aux
dim qtde_delta
dim vLog1()
dim vLog2()
dim campos_a_omitir
dim strComplemento

	estoque_atualiza_xml = False
	msg_erro = ""
	info_log = ""

	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

	if Not le_estoque_agio(r_estoque.id_estoque, r_estoque_bd, msg_erro) then exit function
	if Not le_estoque_item_xml(r_estoque.id_estoque, v_item_bd, msg_erro) then exit function

	strComplemento=""
	if Cstr(r_estoque.entrada_especial) <> Cstr(0) then strComplemento="ENTRADA_ESPECIAL"
			
	gravou_item = False
	s_log = ""
	campos_a_omitir = ""
	
'	PRODUTOS NOVOS NA LISTA
'	=======================
'	COLOCA OS PRODUTOS NOVOS NO FINAL, SENDO QUE OS ÍNDICES DE SEQUENCIAÇÃO SERÃO 
'	COMPACTADOS MAIS ADIANTE.
	i_seq = UBound(v_item)
	for i = Lbound(v_item) to Ubound(v_item)
		if Trim(v_item(i).produto) <> "" then
			achou = False
			for j = Lbound(v_item_bd) to Ubound(v_item_bd)
				if Trim(v_item(i).produto) = Trim(v_item_bd(j).produto) then
					achou = True
					exit for
					end if
				next
			
		'	É UM PRODUTO NOVO NA LISTA
			If Not achou Then
				i_seq = i_seq + 1
				with v_item(i)
					s = "SELECT * FROM t_PRODUTO WHERE" & _
						" (fabricante='" & r_estoque.fabricante & "')" 
					if IsEAN(.produto) then
						s = s & " AND (ean='" & .produto & "')"
					else
						s = s & " AND (produto='" & .produto & "')"
						end if

					if rs.State <> 0 then rs.Close
					rs.Open s, cn
					if Err <> 0 then
						msg_erro = Cstr(Err) & ": " & Err.Description
						exit function
						end if
					
					if rs.Eof then
						msg_erro = "Produto " & .produto & " do fabricante " & r_estoque.fabricante & " não está cadastrado."
						if rs.State <> 0 then rs.Close
						exit function
					else
						.fabricante = Trim(r_estoque.fabricante)
					'	CARREGA CÓDIGO INTERNO DO PRODUTO
						.produto = Trim("" & rs("produto"))
						if (.preco_fabricante = 0) And (rs("preco_fabricante") <> 0) then .preco_fabricante = rs("preco_fabricante")
						if (.vl_custo2 = 0) And (rs("vl_custo2") <> 0) then .vl_custo2 = rs("vl_custo2")
						end if
					
					s = "INSERT INTO T_ESTOQUE_ITEM" & _
						" (id_estoque, fabricante, produto, qtde, preco_fabricante, vl_custo2," & _
						" vl_BC_ICMS_ST, vl_ICMS_ST," & _
						" ncm, cst," & _
						" data_ult_movimento, sequencia, " & _
                        " ean, aliq_ipi, vl_ipi, aliq_icms " & _
						") VALUES (" & _
						"'" & r_estoque.id_estoque & "'" & _
						",'" & .fabricante & "'" & _
						",'" & .produto & "'" & _
						"," & CStr(.qtde) & _
						"," & bd_formata_numero(.preco_fabricante) & _
						"," & bd_formata_numero(.vl_custo2) & _
						"," & bd_formata_numero(.vl_BC_ICMS_ST) & _
						"," & bd_formata_numero(.vl_ICMS_ST) & _
						",'" & Trim(.ncm) & "'" & _
						",'" & Trim(.cst) & "'" & _
						"," & bd_formata_data(Date) & _
						"," & CStr(i_seq) & _
						",'" & Trim(.ean) & "'" & _
						"," & bd_formata_numero(.aliq_ipi) & _
						"," & bd_formata_numero(.vl_ipi) & _
						"," & bd_formata_numero(.aliq_icms) & _
						")"
					cn.Execute(s)
					if Err <> 0 then
						msg_erro=Cstr(Err) & ": " & Err.Description
						exit function
						end if
					
					gravou_item = True
				'	INFORMAÇÕES P/ O LOG
					s_log = s_log & log_estoque_monta_incremento(.qtde, "", .produto)
					
					'Log de movimentação do estoque
					if Not grava_log_estoque_v2(id_usuario, r_estoque.id_nfe_emitente, .fabricante, .produto, .qtde, .qtde, OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM, "", ID_ESTOQUE_VENDA, "", "", "", "", Trim(r_estoque.documento), strComplemento, "") then
						msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
						exit function
						end if
					end with
				end if
			end if
		next
		
		
'	PRODUTOS ALTERADOS/EXCLUÍDOS
'	============================
'	PRODUTOS ALTERADOS: VERIFICA SE A NOVA QUANTIDADE ESTÁ CONSISTENTE C/ RELAÇÃO À QTDE UTILIZADA.
'	PRODUTOS EXCLUÍDOS: NÃO EXISTE EXCLUSÃO DIRETA, DEVE-SE CADASTRAR COMO UMA ALTERAÇÃO P/ 
'						"QUANTIDADE = 0", E MAIS ADIANTE HAVERÁ UMA ROTINA QUE REMOVE OS REGISTROS 
'						C/ QUANTIDADE ZERADA.
	for i = Lbound(v_item) to Ubound(v_item)
		if Trim(v_item(i).produto) <> "" then
			achou = False
			for j = Lbound(v_item_bd) to Ubound(v_item_bd)
				if Trim(v_item(i).produto) = Trim(v_item_bd(j).produto) then
					achou = True
					i_ref = j
					exit for
					end if
				next
		
		'	É UMA ALTERAÇÃO NO PRODUTO?
			if achou then
				if (v_item(i).qtde <> v_item_bd(i_ref).qtde) Or _
				   (v_item(i).vl_BC_ICMS_ST <> v_item_bd(i_ref).vl_BC_ICMS_ST) Or _
				   (v_item(i).vl_ICMS_ST <> v_item_bd(i_ref).vl_ICMS_ST) Or _
				   (v_item(i).preco_fabricante <> v_item_bd(i_ref).preco_fabricante) Or _
				   (v_item(i).vl_custo2 <> v_item_bd(i_ref).vl_custo2) Or _
				   (v_item(i).ncm <> v_item_bd(i_ref).ncm) Or _
				   (v_item(i).cst <> v_item_bd(i_ref).cst) Or _
				   (v_item(i).ean <> v_item_bd(i_ref).ean) Or _
				   (v_item(i).aliq_ipi <> v_item_bd(i_ref).aliq_ipi) Or _
				   (v_item(i).vl_ipi <> v_item_bd(i_ref).vl_ipi) Or _
				   (v_item(i).aliq_icms <> v_item_bd(i_ref).aliq_icms) then
					with v_item(i)
						s = "SELECT * FROM t_ESTOQUE_ITEM WHERE" & _
							" (id_estoque='" & r_estoque.id_estoque & "')" & _
							" AND (fabricante='" & .fabricante & "')" & _
							" AND (produto='" & .produto & "')"
						if rs.State <> 0 then rs.Close
						rs.Open s, cn
						if Err <> 0 then
							msg_erro=Cstr(Err) & ": " & Err.Description
							exit function
							end if
						
						if rs.Eof then
							msg_erro = "Não foi encontrado o registro do produto " & .produto & " do lote do estoque nº " & r_estoque.id_estoque
							if rs.State <> 0 then rs.Close
							exit function
						else
							qtde_aux = rs("qtde")
							qtde_utilizada_aux = rs("qtde_utilizada")
							vl_BC_ICMS_ST_aux = rs("vl_BC_ICMS_ST")
							vl_ICMS_ST_aux = rs("vl_ICMS_ST")
							preco_fabricante_aux = rs("preco_fabricante")
							vl_custo2_aux = rs("vl_custo2")
							ncm_aux = Trim("" & rs("ncm"))
							cst_aux = Trim("" & rs("cst"))
							ean_aux = Trim("" & rs("ean"))
							aliq_ipi_aux = rs("aliq_ipi")
							vl_ipi_aux = rs("vl_ipi")
							aliq_icms_aux = rs("aliq_icms")
						'	QUANTIDADE ESTÁ CONSISTENTE, ENTÃO PODE ATUALIZAR O REGISTRO
							if .qtde < rs("qtde_utilizada") then
								msg_erro = texto_add_br(msg_erro)
								msg_erro = msg_erro & "A quantidade do produto " & .produto & " NÃO foi alterada de " & CStr(qtde_aux) & " para " & CStr(.qtde) & ", pois " & CStr(qtde_utilizada_aux) & " unidades já foram utilizadas!!"
							else
								rs("qtde") = .qtde
								rs("vl_BC_ICMS_ST") = converte_numero(.vl_BC_ICMS_ST)
								rs("vl_ICMS_ST") = converte_numero(.vl_ICMS_ST)
								rs("preco_fabricante") = converte_numero(.preco_fabricante)
								rs("vl_custo2") = converte_numero(.vl_custo2)
								rs("ncm") = Trim(.ncm)
								rs("cst") = Trim(.cst)
								rs("data_ult_movimento") = Date
								rs("ean") = Trim(.ean)
								rs("aliq_ipi") = converte_numero(.aliq_ipi)
								rs("vl_ipi") = converte_numero(.vl_ipi)
								rs("aliq_icms") = converte_numero(.aliq_icms)
								rs.Update
								if Err <> 0 then
									msg_erro=Cstr(Err) & ": " & Err.Description
									exit function
									end if
								
								gravou_item = True
							'	INFORMAÇÕES P/ O LOG
								If qtde_aux > .qtde Then
									qtde_delta = qtde_aux - .qtde
									if qtde_delta <> 0 then
										if s_log <> "" then s_log = s_log & ";"
										s_log = s_log & log_estoque_monta_decremento((qtde_aux - .qtde), "", .produto)
										'Log de movimentação do estoque
										if Not grava_log_estoque_v2(id_usuario, r_estoque.id_nfe_emitente, rs("fabricante"), rs("produto"), qtde_delta, qtde_delta, OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA, ID_ESTOQUE_VENDA, "", "", "", "", "", Trim(r_estoque.documento), strComplemento, "") then
											msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
											exit function
											end if
										end if
								Else
									qtde_delta = .qtde - qtde_aux
									if qtde_delta <> 0 then
										if s_log <> "" then s_log = s_log & ";"
										s_log = s_log & log_estoque_monta_incremento((.qtde - qtde_aux), "", .produto)
										'Log de movimentação do estoque
										if Not grava_log_estoque_v2(id_usuario, r_estoque.id_nfe_emitente, rs("fabricante"), rs("produto"), qtde_delta, qtde_delta, OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA, "", ID_ESTOQUE_VENDA, "", "", "", "", Trim(r_estoque.documento), strComplemento, "") then
											msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
											exit function
											end if
										end if
									End If
								
								if converte_numero(preco_fabricante_aux) <> converte_numero(.preco_fabricante) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": preco_fabricante: " & formata_moeda(preco_fabricante_aux) & " => " & formata_moeda(.preco_fabricante)
									end if

								if converte_numero(vl_custo2_aux) <> converte_numero(.vl_custo2) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_custo2: " & formata_moeda(vl_custo2_aux) & " => " & formata_moeda(.vl_custo2)
									end if

								if converte_numero(vl_BC_ICMS_ST_aux) <> converte_numero(.vl_BC_ICMS_ST) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_BC_ICMS_ST: " & formata_moeda(vl_BC_ICMS_ST_aux) & " => " & formata_moeda(.vl_BC_ICMS_ST)
									end if
								
								if converte_numero(vl_ICMS_ST_aux) <> converte_numero(.vl_ICMS_ST) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_ICMS_ST: " & formata_moeda(vl_ICMS_ST_aux) & " => " & formata_moeda(.vl_ICMS_ST)
									end if
								
								if Trim(ncm_aux) <> Trim(.ncm) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": NCM: " & Trim(ncm_aux) & " => " & Trim(.ncm)
									end if
								
								if Trim(cst_aux) <> Trim(.cst) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": CST: " & Trim(cst_aux) & " => " & Trim(.cst)
									end if

    							if Trim(ean_aux) <> Trim(.ean) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": EAN: " & Trim(ean_aux) & " => " & Trim(.ean)
									end if

								if converte_numero(aliq_ipi_aux) <> converte_numero(.aliq_ipi) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": aliq_ipi: " & formata_moeda(aliq_ipi_aux) & " => " & formata_moeda(.aliq_ipi)
									end if

								if converte_numero(vl_ipi_aux) <> converte_numero(.vl_ipi) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": vl_ipi: " & formata_moeda(vl_ipi_aux) & " => " & formata_moeda(.vl_ipi)
									end if

								if converte_numero(aliq_icms_aux) <> converte_numero(.aliq_icms) then
									if s_log <> "" then s_log = s_log & "; "
									s_log = s_log & "Prod " & .produto & ": aliq_icms: " & formata_moeda(aliq_icms_aux) & " => " & formata_moeda(.aliq_icms)
									end if

								end if
							end if
						end with
					end if
				end if
			end if
		next


'	INFORMAÇÕES P/ O LOG
	if s_log <> "" then s_log = "alteração nos dados de produtos:" & s_log


'	PRODUTOS QUE ALTERARAM A QUANTIDADE P/ ZERO: A QTDE IGUAL A ZERO ASSEGURA QUE 
'	NENHUMA OPERAÇÃO DE SAÍDA DE ESTOQUE SERÁ FEITA NESSES REGISTROS, ENTÃO PODEM 
'	SER ELIMINADOS DIRETAMENTE.
	s = "DELETE FROM t_ESTOQUE_ITEM WHERE" & _
		" (id_estoque='" & Trim(r_estoque.id_estoque) & "') AND (qtde=0) AND (qtde_utilizada=0)"
	cn.execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
  '	SE NÃO RESTOU NENHUM PRODUTO, PODE REMOVER O REGISTRO DE ENTRADA NO ESTOQUE
	s = "SELECT COUNT(*) AS total FROM t_ESTOQUE_ITEM WHERE" & _
		" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
	n_item = -1
	if rs.State <> 0 then rs.Close
	rs.Open s, cn 
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	if Not rs.Eof then
		if IsNumeric(rs("total")) then n_item = CLng(rs("total"))
		end if
	
'	VERIFICA SE HÁ VÍNCULOS NA TABELA DE MOVIMENTOS DO ESTOQUE
'	CASO SIM, MANTÉM O REGISTRO DE ENTRADA NO ESTOQUE P/ FINS DE HISTÓRICO
	s = "SELECT COUNT(*) AS total FROM t_ESTOQUE_MOVIMENTO WHERE" & _
		" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
	n_movto = -1
	if rs.State <> 0 then rs.Close
	rs.Open s, cn
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	if Not rs.Eof then
		if IsNumeric(rs("total")) then n_movto = CLng(rs("total"))
		end if
	
	If (n_item = 0) And (n_movto = 0) Then
	'	EXCLUI O REGISTRO DE ESTOQUE!!
		s = "DELETE FROM t_ESTOQUE WHERE" & _
			" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
		cn.execute(s)
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		
	'	VERIFICA SE CONSEGUIU EXCLUIR 
		s = "SELECT id_estoque FROM t_ESTOQUE WHERE" & _
			" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
		if rs.State <> 0 then rs.Close
		rs.Open s, cn
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		
		if Not rs.Eof then
			msg_erro="Falha ao tentar remover o registro do lote do estoque nº " & r_estoque.id_estoque
			exit function
			end if
	
	else
	'	ATUALIZA INFORMAÇÕES DO REGISTRO DE ENTRADA NO ESTOQUE
		if gravou_item Or _
		   (Trim(r_estoque.documento)<>Trim(r_estoque_bd.documento)) Or _
		   (Trim(r_estoque.obs)<>Trim(r_estoque_bd.obs)) Or _
		   (r_estoque.entrada_especial<>r_estoque_bd.entrada_especial) Or _
		   (r_estoque.perc_agio<>r_estoque_bd.perc_agio) then
			s = "SELECT * FROM t_ESTOQUE WHERE" & _
				" (id_estoque='" & Trim(r_estoque.id_estoque) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
				
			if rs.Eof then
				msg_erro="O registro do lote do estoque nº " & r_estoque.id_estoque & " não foi encontrado."
				if rs.State <> 0 then rs.Close
				exit function
				end if

			log_via_vetor_carrega_do_recordset rs, vLog1, campos_a_omitir
			rs("documento") = Trim(r_estoque.documento)
			rs("obs") = Trim(r_estoque.obs)
			rs("entrada_especial") = r_estoque.entrada_especial
            rs("perc_agio") = r_estoque.perc_agio
			if gravou_item then rs("data_ult_movimento") = Date
			rs.Update

			log_via_vetor_carrega_do_recordset rs, vLog2, campos_a_omitir
			s = log_via_vetor_monta_alteracao(vLog1, vLog2)
			if s <> "" then 
				if s_log <> "" then s_log = "; " & s_log
				s_log = s & s_log
				end if
			end if
		
	'	COMPACTA A SEQUÊNCIA
		i_seq = 0
		s = "SELECT * FROM t_ESTOQUE_ITEM WHERE (id_estoque='" & Trim(r_estoque.id_estoque) & "') ORDER BY sequencia"
		if rs.State <> 0 then rs.Close
		rs.Open s, cn
		if Err <> 0 then
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		
		do while Not rs.Eof
			i_seq = i_seq + 1
			rs("sequencia") = i_seq
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			rs.movenext
			loop
		end if
	
	
	if msg_erro <> "" then exit function
	
	if s_log <> "" then
		s_log = "Alteração no registro de estoque (" & Trim(r_estoque.id_estoque) & ")" & _
				" do fabricante " & Trim(r_estoque.fabricante) & "," & _
				" entrada em " & formata_data(r_estoque.data_entrada) & "," & _
				" documento " & Trim(r_estoque.documento) & ": " & _
				s_log
		info_log = s_log
		end if
		
	estoque_atualiza_xml = True
end function

' --------------------------------------------------------------------
'   ESTOQUE PRODUTO SAIDA PARA KIT V2
'   Retorno da função:
'      False - Ocorreu falha ao tentar alterar os dados do estoque.
'      True - Conseguiu alterar os dados do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função processa a saída do produto do estoque que foi 
'	usado para compor um kit.
'	Esta rotina foi projetada p/ processar a saída de produtos que 
'	compõem 1 unidade do kit.
'	Ou seja, se forem cadastrados 50 unidades de um kit composto 
'	por 3 produtos, esta rotina deve ser chamada 50 vezes para 
'	cada um dos 3 produtos.
'	Isto se deve ao cálculo do preço fabricante, já que o parâmetro 
'	retorna o valor acumulado dos produtos que saíram do estoque.
'	Lembre-se do problema que está sendo tratado aqui: ao cadastrar 
'	o kit, podem ser usados produtos que deram entrada no estoque com 
'	valores variados de preço fabricante.  O preço fabricante do kit 
'	é a soma do preço fabricante dos produtos usados na sua composição.
'	O sistema deve calcular o preço fabricante p/ cada unidade do kit
'	e agrupar aqueles que possuem o mesmo valor.  Se todas as unidades
'	do kit possuírem o mesmo valor p/ preço fabricante, então será 
'	criado um único registro de entrada no estoque.  Caso contrário,
'	serão criados automaticamente diferentes registros de entrada 
'	no estoque.
function estoque_produto_saida_para_kit_v2(ByVal id_usuario, ByVal id_estoque_do_kit, _
										ByVal id_nfe_emitente, _
										ByVal id_fabricante, ByVal id_produto, ByVal qtde_a_sair, _
										ByRef preco_fabricante_acumulado, _
										ByRef vl_custo2_acumulado, _
										ByRef vl_BC_ICMS_ST_acumulado, _
										ByRef vl_ICMS_ST_acumulado, _
										ByRef msg_erro)
dim s
dim rs
dim qtde_disponivel
dim qtde_movimentada
dim v_estoque
dim iv
dim qtde_aux
dim qtde_utilizada_aux
dim preco_fabricante_aux
dim vl_custo2_aux
dim vl_BC_ICMS_ST_aux
dim vl_ICMS_ST_aux
dim qtde_movto
dim s_chave

	estoque_produto_saida_para_kit_v2 = False
	msg_erro = ""
	preco_fabricante_acumulado = 0
	vl_custo2_acumulado = 0
	vl_BC_ICMS_ST_acumulado = 0
	vl_ICMS_ST_acumulado = 0
	
'	NENHUMA UNIDADE SERÁ RETIRADA!!
	If (qtde_a_sair<=0) Or (Trim(id_produto)="") Or (Trim(id_estoque_do_kit)="") Then
		estoque_produto_saida_para_kit_v2 = True
		Exit Function
		End If

	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

'	OBTÉM OS "LOTES" DO PRODUTO DISPONÍVEIS NO ESTOQUE (POLÍTICA FIFO)
	s = "SELECT" & _
			" t_ESTOQUE.id_estoque, (qtde-qtde_utilizada) AS saldo" & _
		" FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
		" WHERE" & _
			" (t_ESTOQUE.id_nfe_emitente = " & Trim("" & id_nfe_emitente) & ")" & _
			" AND (t_ESTOQUE_ITEM.fabricante='" & id_fabricante & "')" & _
			" AND (produto='" & id_produto & "')" & _
			" AND ((qtde-qtde_utilizada) > 0)" & _
		" ORDER BY" & _
			" data_entrada, t_ESTOQUE.id_estoque"
	rs.open s, cn

	qtde_disponivel = 0
	ReDim v_estoque(0)
	v_estoque(UBound(v_estoque)) = ""

	do while Not rs.Eof
	'	ARMAZENA AS ENTRADAS NO ESTOQUE CANDIDATAS À SAÍDA DE PRODUTOS
		If v_estoque(UBound(v_estoque)) <> "" Then
			ReDim Preserve v_estoque(UBound(v_estoque) + 1)
			v_estoque(UBound(v_estoque)) = ""
			End If
		v_estoque(UBound(v_estoque)) = Trim("" & rs("id_estoque"))
		qtde_disponivel = qtde_disponivel + CLng(rs("saldo"))
		rs.movenext
		loop

'	NÃO HÁ PRODUTOS SUFICIENTES NO ESTOQUE!!
	If qtde_a_sair > qtde_disponivel Then
		msg_erro = "Produto " & id_produto & " do fabricante " & id_fabricante & ": faltam " & _
					formata_inteiro(qtde_a_sair-qtde_disponivel) & " unidades no estoque (" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente) & ")."
		Exit Function
		End If

'	REALIZA A SAÍDA DO ESTOQUE!!
	qtde_movimentada = 0
	For iv = LBound(v_estoque) To UBound(v_estoque)
	
		If Trim(v_estoque(iv)) <> "" Then
		
		'	A QUANTIDADE NECESSÁRIA JÁ FOI RETIRADA DO ESTOQUE!!
			If qtde_movimentada >= qtde_a_sair Then Exit For
			
		'	T_ESTOQUE_ITEM: SAÍDA DE PRODUTOS
			s = "SELECT " & _
					"*" & _
				" FROM t_ESTOQUE_ITEM" & _
				" WHERE" & _
					" (id_estoque = '" & Trim(v_estoque(iv)) & "')" & _
					" AND (fabricante = '" & id_fabricante & "')" & _
					" AND (produto = '" & id_produto & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

			if rs.Eof then
				msg_erro = "Falha ao acessar o registro no estoque do produto " & id_produto & " do fabricante " & id_fabricante & " (id_estoque = '" & Trim(v_estoque(iv)) & "')"
				Exit Function
			else
				qtde_aux = rs("qtde")
				qtde_utilizada_aux = rs("qtde_utilizada")
				preco_fabricante_aux = rs("preco_fabricante")
				vl_custo2_aux = rs("vl_custo2")
				vl_BC_ICMS_ST_aux = rs("vl_BC_ICMS_ST")
				vl_ICMS_ST_aux = rs("vl_ICMS_ST")
				End If
			
			If (qtde_a_sair - qtde_movimentada) > (qtde_aux - qtde_utilizada_aux) Then
			'	QUANTIDADE DE PRODUTOS DESTE ITEM DE ESTOQUE É INSUFICIENTE P/ ATENDER O PEDIDO
				qtde_movto = qtde_aux - qtde_utilizada_aux
			Else
			'	QUANTIDADE DE PRODUTOS DESTE ITEM SOZINHO É SUFICIENTE P/ ATENDER O PEDIDO
				qtde_movto = qtde_a_sair - qtde_movimentada
				End If

			rs("qtde_utilizada") = rs("qtde_utilizada") + qtde_movto
			rs("data_ult_movimento") = Date
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
		
		'	CONTABILIZA QUANTIDADE MOVIMENTADA
			qtde_movimentada = qtde_movimentada + qtde_movto
		
		'	TOTALIZA O PREÇO DO FABRICANTE DESTE PRODUTO P/ SER USADO
		'	NO CÁLCULO DO PREÇO FABRICANTE DO KIT
			preco_fabricante_acumulado = preco_fabricante_acumulado + (qtde_movto * preco_fabricante_aux)
			vl_custo2_acumulado = vl_custo2_acumulado + (qtde_movto * vl_custo2_aux)
			vl_BC_ICMS_ST_acumulado = vl_BC_ICMS_ST_acumulado + (qtde_movto * vl_BC_ICMS_ST_aux)
			vl_ICMS_ST_acumulado = vl_ICMS_ST_acumulado + (qtde_movto * vl_ICMS_ST_aux)
			
		'	T_ESTOQUE_MOVIMENTO: REGISTRA O MOVIMENTO DE SAÍDA DO ESTOQUE
			If Not gera_id_estoque_movto(s_chave, msg_erro) Then
				msg_erro = "Falha ao tentar obter um nº de identificação único para este registro de movimentação no estoque!!" & _
							chr(13) & msg_erro
				Exit Function
				End If
			
			s = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
				" (id_movimento, data, hora, operacao, estoque, usuario, pedido, loja," & _
				" fabricante, produto, id_estoque, qtde, kit, kit_id_estoque) VALUES" & _
				" ('" & s_chave & "'" & _
				"," & bd_formata_data(Date) & _
				",'" & retorna_so_digitos(formata_hora(Now)) & "'" & _
				",'" & OP_ESTOQUE_CONVERSAO_KIT & "'" & _
				",'" & ID_ESTOQUE_KIT & "'" & _
				",'" & id_usuario & "'" & _
				",'" & "" & "'" & _
				",'" & "" & "'" & _
				",'" & id_fabricante & "'" & _
				",'" & id_produto & "'" & _
				",'" & Trim(v_estoque(iv)) & "'" & _
				"," & CStr(qtde_movto) & _
				"," & "1" & _
				",'" & Trim(id_estoque_do_kit) & "')"
			cn.Execute(s)
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

		'	T_ESTOQUE: ATUALIZA DATA DO ÚLTIMO MOVIMENTO
			s = "SELECT " & _
					"*" & _
				" FROM t_ESTOQUE" & _
				" WHERE" & _
					" (id_estoque = '" & v_estoque(iv) & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.Eof then
				msg_erro = "Falha ao acessar o registro principal no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				Exit Function
			else
				rs("data_ult_movimento") = Date
				rs.Update
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				End If

		'	JÁ CONSEGUIU ALOCAR TUDO?
			If qtde_movimentada >= qtde_a_sair Then Exit For
			End If
		Next
	
	if rs.State <> 0 then rs.Close
	set rs=nothing
	
	'Log de movimentação do estoque
	if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, id_fabricante, id_produto, qtde_a_sair, qtde_a_sair, OP_ESTOQUE_LOG_CONVERSAO_KIT, ID_ESTOQUE_VENDA, ID_ESTOQUE_KIT, "", "", "", "", "", "", "") then
		msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
		exit function
		end if
	
	estoque_produto_saida_para_kit_v2 = True
end function

' --------------------------------------------------------------------
'   ESTOQUE PRODUTO TRANSFERÊNCIA ENTRE CD (COMPARAÇÃO DE QUANTIDADES)
'   Retorno da função:
'      False - Ocorreu falha ao tentar alterar os dados do estoque.
'      True - Conseguiu alterar os dados do estoque.
'   Esta função verifica se a quantidade permanece a mesma antes de gravar
'   as informações.
'	
'	(melhorar esta descrição)
'	
function estoque_produto_transf_consiste_quantidades(ByVal id_nfe_emitente_origem, _
                                                    ByVal id_nfe_emitente_destino, _
                                                    ByRef v_1, _
                                                    ByRef v_2, _
										            ByRef msg_erro)
dim s
dim s_cod_prod1
dim s_cod_prod2
dim iv
dim soma

	estoque_produto_transf_consiste_quantidades = False
	
    msg_erro = ""	
    s_cod_prod1 = ""
    s_cod_prod2 = ""
    soma = 0

    'REAGRUPANDO OS PRODUTOS PARA REDIVIDIR E COMPARAR 
    for iv = LBound(v_1) to UBound(v_1)
        with v_1(iv)
            s_cod_prod1 = .produto
            if (s_cod_prod1 <> s_cod_prod2) And (s_cod_prod2 <> "") then                
                if not estoque_produto_transf_cd_monta(id_nfe_emitente_origem, _
                                                        v_1(iv-1).fabricante, _
                                                        v_1(iv-1).produto, _
                                                        soma, _
                                                        v_2, _
                                                        s) then
                    msg_erro = "Problemas na conferência da transferência: " & s
                    exit function
                    end if
                soma = 0
                end if
            s_cod_prod2 = s_cod_prod1
            soma = soma + .qtde
            end with
        next
    'FAZENDO O REAGRUPAMENTO DO ÚLTIMO PRODUTO DA LISTA (POIS O LOOP PULA ESTE PRODUTO)
    if not estoque_produto_transf_cd_monta(id_nfe_emitente_origem, _
                                            v_1(UBound(v_1)).fabricante, _
                                            v_1(UBound(v_1)).produto, _
                                            soma, _
                                            v_2, _
                                            s) then
        msg_erro = "Problemas na conferência da transferência: " & s
        exit function
        end if

    if msg_erro = "" then
        'primeira comparação: se os vetores não tiverem o mesmo número de elementos, os estoques mudaram
        if UBound(v_1) <> UBound(v_2) then
            msg_erro = "Mudanças nas quantidades dos estoques, reiniciar o processo!!!"
            exit function
        else
        'segunda comparação: se os conteúdos dos vetores forem diferentes, os estoques mudaram
            for iv = LBound(v_1) to UBound(v_2)
                if (v_1(iv).id_estoque_origem <> v_2(iv).id_estoque_origem) Or _
                    (v_1(iv).fabricante <> v_2(iv).fabricante) Or _
                    (v_1(iv).produto <> v_2(iv).produto) Or _
                    (v_1(iv).qtde <> v_2(iv).qtde) then
                    msg_erro = "Mudanças em características dos estoques, reiniciar o processo!!!"
                    exit function
                    end if
                next
            end if
        end if	
	
	estoque_produto_transf_consiste_quantidades = True
end function

' --------------------------------------------------------------------
'   ESTOQUE PRODUTO TRANSFERÊNCIA ENTRE CD (REAGRUPAMENTO DE ITENS)
'   Retorno da função:
'      False - Ocorreu falha ao tentar alterar os dados do estoque.
'      True - Conseguiu alterar os dados do estoque.
'   Esta função pega o vetor onde foi feita a distribuição de quantidades
'   dos produtos e os reagrupa conforme informado na tela inicial.
'	
'	(melhorar esta descrição)
'	
function estoque_produto_transf_reagrupa_itens(ByRef v_distribuido, _
                                                    ByRef v_reagrupado, _
										            ByRef msg_erro)
dim s
dim s_cod_prod1
dim s_cod_prod2
dim s_fabr1
dim s_fabr2
dim iv
dim v_aliq_ipi
dim v_vl_ipi
dim v_aliq_icms
dim soma

	estoque_produto_transf_reagrupa_itens = False
	
    msg_erro = ""	
    s_cod_prod1 = ""
    s_cod_prod2 = ""
    s_fabr1 = ""
    s_fabr2 = ""
    v_aliq_ipi = 0
    v_vl_ipi = 0
    v_aliq_icms = 0
    soma = 0

    'REAGRUPANDO OS PRODUTOS
    for iv = LBound(v_distribuido) to UBound(v_distribuido)
        with v_distribuido(iv)
            s_cod_prod1 = .produto
            s_fabr1 = .fabricante
            if (s_cod_prod1 <> s_cod_prod2) And (s_cod_prod2 <> "") then                
			    if Trim("" & v_reagrupado(Ubound(v_reagrupado)).produto) <> "" then
				    redim preserve v_reagrupado(ubound(v_reagrupado)+1)
				    set v_reagrupado(UBound(v_reagrupado)) = New cl_ESTOQUE_TRANSFERENCIA_ITEM
				    end if			
				v_reagrupado(UBound(v_reagrupado)).fabricante = v_distribuido(iv-1).fabricante
				v_reagrupado(UBound(v_reagrupado)).produto = v_distribuido(iv-1).produto
				v_reagrupado(UBound(v_reagrupado)).qtde = soma
                v_reagrupado(UBound(v_reagrupado)).aliq_ipi = v_aliq_ipi
                v_reagrupado(UBound(v_reagrupado)).vl_ipi = v_vl_ipi
                v_reagrupado(UBound(v_reagrupado)).aliq_icms = v_aliq_icms
                soma = 0
                end if
            s_cod_prod2 = s_cod_prod1
            s_fabr2 = s_fabr1
            soma = soma + .qtde
            v_aliq_ipi = .aliq_ipi
            v_vl_ipi = .vl_ipi
            v_aliq_icms = .aliq_icms
            end with
        next
    'FAZENDO O REAGRUPAMENTO DO ÚLTIMO PRODUTO DA LISTA (POIS O LOOP PULA ESTE PRODUTO)
	if Trim("" & v_reagrupado(Ubound(v_reagrupado)).produto) <> "" then
		redim preserve v_reagrupado(ubound(v_reagrupado)+1)
		set v_reagrupado(UBound(v_reagrupado)) = New cl_ESTOQUE_TRANSFERENCIA_ITEM
		end if			
	v_reagrupado(UBound(v_reagrupado)).fabricante = v_distribuido(UBound(v_distribuido)).fabricante
	v_reagrupado(UBound(v_reagrupado)).produto = v_distribuido(UBound(v_distribuido)).produto
	v_reagrupado(UBound(v_reagrupado)).qtde = soma
    v_reagrupado(UBound(v_reagrupado)).aliq_ipi = v_aliq_ipi
    v_reagrupado(UBound(v_reagrupado)).vl_ipi = v_vl_ipi
    v_reagrupado(UBound(v_reagrupado)).aliq_icms = v_aliq_icms
 

	estoque_produto_transf_reagrupa_itens = True
end function

' --------------------------------------------------------------------
'   ESTOQUE PRODUTO TRANSFERÊNCIA ENTRE CD (MONTAGEM)
'   Retorno da função:
'      False - Ocorreu falha ao tentar alterar os dados do estoque.
'      True - Conseguiu alterar os dados do estoque.
'   Esta função simula a saída do produto do estoque que será 
'	usado em um novo estoque, em outro CD.
'	Para cada produto transferido, a rotina é executada.
'	
'	(melhorar esta descrição comparando com o arquivo "rotina_calcula_transf.vb")
'	

function estoque_produto_transf_cd_monta(ByVal id_nfe_emitente, _
										ByVal id_fabricante, ByVal id_produto, ByVal qtde_a_sair, _
										ByRef v_item, _
										ByRef msg_erro)
dim s
dim rs
dim qtde_disponivel
dim qtde_movimentada
dim v_estoque
dim v_documento
dim v_entrada_tipo
dim iv
dim descricao_html_aux
dim qtde_aux
dim qtde_utilizada_aux
dim preco_fabricante_aux
dim vl_custo2_aux
dim vl_BC_ICMS_ST_aux
dim vl_ICMS_ST_aux
dim ncm_aux
dim cst_aux
dim st_ncm_cst_herdado_tabela_produto_aux
dim ean_aux
dim aliq_ipi_aux
dim aliq_icms_aux
dim vl_ipi_aux
dim preco_origem_aux
dim produto_xml_aux
dim qtde_movto
dim s_chave

	estoque_produto_transf_cd_monta = False
	msg_erro = ""
	
'	NENHUMA UNIDADE SERÁ RETIRADA!!
	If (qtde_a_sair<=0) Or (Trim(id_produto)="") Then
		estoque_produto_transf_cd_monta = True
		Exit Function
		End If

	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

'	OBTÉM OS "LOTES" DO PRODUTO DISPONÍVEIS NO ESTOQUE (POLÍTICA FIFO)
	s = "SELECT" & _
			" t_ESTOQUE.id_estoque, t_ESTOQUE.documento, t_ESTOQUE.entrada_tipo, (qtde-qtde_utilizada) AS saldo" & _
		" FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
		" WHERE" & _
			" (t_ESTOQUE.id_nfe_emitente = " & Trim("" & id_nfe_emitente) & ")" & _
			" AND (t_ESTOQUE_ITEM.fabricante='" & id_fabricante & "')" & _
			" AND (produto='" & id_produto & "')" & _
			" AND ((qtde-qtde_utilizada) > 0)" & _
		" ORDER BY" & _
			" data_entrada, t_ESTOQUE.id_estoque"
	rs.open s, cn

	qtde_disponivel = 0
	ReDim v_estoque(0)
	ReDim v_documento(0)
    ReDim v_entrada_tipo(0)
	v_estoque(UBound(v_estoque)) = ""
	v_documento(UBound(v_documento)) = ""
    v_entrada_tipo(UBound(v_entrada_tipo)) = 0

	do while Not rs.Eof
	'	ARMAZENA AS ENTRADAS NO ESTOQUE CANDIDATAS À SAÍDA DE PRODUTOS
		If v_estoque(UBound(v_estoque)) <> "" Then
			ReDim Preserve v_estoque(UBound(v_estoque) + 1)
			v_estoque(UBound(v_estoque)) = ""
			End If
		v_estoque(UBound(v_estoque)) = Trim("" & rs("id_estoque"))
		If v_documento(UBound(v_documento)) <> "" Then
			ReDim Preserve v_documento(UBound(v_documento) + 1)
			v_documento(UBound(v_documento)) = ""
			End If
		v_documento(UBound(v_documento)) = Trim("" & rs("documento"))
        If v_entrada_tipo(UBound(v_entrada_tipo)) <> "" Then
			ReDim Preserve v_entrada_tipo(UBound(v_entrada_tipo) + 1)
			v_entrada_tipo(UBound(v_entrada_tipo)) = ""
			End If
		v_entrada_tipo(UBound(v_entrada_tipo)) = rs("entrada_tipo")
		qtde_disponivel = qtde_disponivel + CLng(rs("saldo"))
		rs.movenext
		loop

'	NÃO HÁ PRODUTOS SUFICIENTES NO ESTOQUE!!
	If qtde_a_sair > qtde_disponivel Then
		msg_erro = "Produto " & id_produto & " do fabricante " & id_fabricante & ": faltam " & _
					formata_inteiro(qtde_a_sair-qtde_disponivel) & " unidades no estoque (" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente) & ")."
		Exit Function
		End If

'	SIMULA A SAÍDA DO ESTOQUE!!
	qtde_movimentada = 0
	For iv = LBound(v_estoque) To UBound(v_estoque)
	
		If Trim(v_estoque(iv)) <> "" Then
		
		'	A QUANTIDADE NECESSÁRIA JÁ FOI RETIRADA DO ESTOQUE!!
			If qtde_movimentada >= qtde_a_sair Then Exit For
			
		'	T_ESTOQUE_ITEM: SAÍDA DE PRODUTOS
			s = "SELECT " & _
					"ei.*, p.descricao_html, p.ean as ean_produto" & _
				" FROM t_ESTOQUE_ITEM ei" & _
                " INNER JOIN t_PRODUTO p ON (ei.fabricante = p.fabricante AND ei.produto = p.produto)" & _
				" WHERE" & _
					" (ei.id_estoque = '" & Trim(v_estoque(iv)) & "')" & _
					" AND (ei.fabricante = '" & id_fabricante & "')" & _
					" AND (ei.produto = '" & id_produto & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

			if rs.Eof then
				msg_erro = "Falha ao acessar o registro no estoque do produto " & id_produto & " do fabricante " & id_fabricante & " (id_estoque = '" & Trim(v_estoque(iv)) & "')"
				Exit Function
			else
				qtde_aux = rs("qtde")
				qtde_utilizada_aux = rs("qtde_utilizada")
				preco_fabricante_aux = rs("preco_fabricante")
				vl_custo2_aux = rs("vl_custo2")
                vl_BC_ICMS_ST_aux = rs("vl_BC_ICMS_ST")
                vl_ICMS_ST_aux = rs("vl_ICMS_ST")
				ncm_aux = Trim(rs("ncm"))
				cst_aux = Trim(rs("cst"))
                st_ncm_cst_herdado_tabela_produto_aux = rs("st_ncm_cst_herdado_tabela_produto")
				ean_aux = Trim(rs("ean"))
                aliq_ipi_aux = rs("aliq_ipi")
                aliq_icms_aux = rs("aliq_icms")
                vl_ipi_aux = rs("vl_ipi")
                preco_origem_aux = rs("preco_origem")
                produto_xml_aux = Trim(rs("produto_xml"))
                descricao_html_aux = Trim(rs("descricao_html"))
				End If
			
			If (qtde_a_sair - qtde_movimentada) > (qtde_aux - qtde_utilizada_aux) Then
			'	QUANTIDADE DE PRODUTOS DESTE ITEM DE ESTOQUE É INSUFICIENTE P/ ATENDER O PEDIDO
				qtde_movto = qtde_aux - qtde_utilizada_aux
			Else
			'	QUANTIDADE DE PRODUTOS DESTE ITEM SOZINHO É SUFICIENTE P/ ATENDER O PEDIDO
				qtde_movto = qtde_a_sair - qtde_movimentada
				End If

			If v_item(Ubound(v_item)).produto <> "" Then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ESTOQUE_TRANSFERENCIA_ITEM_SUB
				End if
			v_item(Ubound(v_item)).documento = Trim(v_documento(iv))
            v_item(Ubound(v_item)).entrada_tipo = v_entrada_tipo(iv)
			v_item(Ubound(v_item)).id_estoque_origem = Trim(v_estoque(iv))
			v_item(Ubound(v_item)).fabricante = id_fabricante
			v_item(Ubound(v_item)).produto = id_produto
            v_item(Ubound(v_item)).descricao_html = descricao_html_aux
			v_item(Ubound(v_item)).qtde = qtde_movto
            v_item(Ubound(v_item)).preco_fabricante = preco_fabricante_aux
			v_item(Ubound(v_item)).vl_custo2 = vl_custo2_aux
            v_item(Ubound(v_item)).vl_BC_ICMS_ST = vl_BC_ICMS_ST_aux
            v_item(Ubound(v_item)).vl_ICMS_ST = vl_ICMS_ST_aux
			v_item(Ubound(v_item)).ncm = ncm_aux
			v_item(Ubound(v_item)).cst = cst_aux
			v_item(Ubound(v_item)).ean = ean_aux
            v_item(Ubound(v_item)).st_ncm_cst_herdado_tabela_produto = st_ncm_cst_herdado_tabela_produto_aux
            v_item(Ubound(v_item)).aliq_ipi = aliq_ipi_aux
            v_item(Ubound(v_item)).aliq_icms = aliq_icms_aux
            v_item(Ubound(v_item)).vl_ipi = vl_ipi_aux
            v_item(Ubound(v_item)).preco_origem = preco_origem_aux
            v_item(Ubound(v_item)).produto_xml = produto_xml_aux
		
		'	CONTABILIZA QUANTIDADE MOVIMENTADA
			qtde_movimentada = qtde_movimentada + qtde_movto
		
		'	JÁ CONSEGUIU ALOCAR TUDO?
			If qtde_movimentada >= qtde_a_sair Then Exit For
			End If
		Next
	
	if rs.State <> 0 then rs.Close
	set rs=nothing
	
	
	estoque_produto_transf_cd_monta = True
end function

' --------------------------------------------------------------------
'   ESTOQUE TRANSFEÊNCIA REMOVE
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função marca os registros de transferência entre CDs, como
'   excluídos, desde que isso seja possível.

function transferencia_estoque_remove(byval id_usuario, byval transf_estoque, byref info_log, byref msg_erro)
dim s
dim rs
dim n_item
dim s_log_base

	transferencia_estoque_remove = False
	msg_erro = ""
	info_log = ""
	s_log_base = ""

	transf_estoque = Trim("" & transf_estoque)
	
	s = "SELECT * FROM t_ESTOQUE_TRANSFERENCIA WHERE (id='" & transf_estoque & "')"
	set rs = cn.execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

	if rs.Eof then
		msg_erro = "Registro de transferência nº " & transf_estoque & " não está cadastrado."
	else
		if rs("st_confirmada") <> 0 then msg_erro = "Não é possível reverter a transferência após a confirmação!!"
		s_log_base = "Exclusão da transferência (" & transf_estoque & ")" & _
					" do CD " & Trim("" & CStr(rs("id_nfe_emitente_origem"))) & _
                    " para o CD " & Trim("" & CStr(rs("id_nfe_emitente_destino"))) & _
					 ", registrada em " & formata_data(rs("data"))
		
		s_log_base = s_log_base & ", documento " & Trim("" & rs("documento")) & _
					 ", cadastrada por " & Trim("" & rs("usuario"))
		
		end if
	if rs.State <> 0 then rs.Close	
		
'	ERRO!!
	if msg_erro <> "" then exit function
	
'	TENTA MARCAR COMO EXCLUÍDOS OS ITENS DA TRANSFERÊNCIA

    s = "UPDATE t_ESTOQUE_TRANSFERENCIA_ITEM_SUB SET st_exclusao = 1 WHERE"  & _
		" (id_estoque_transferencia in (SELECT id FROM t_ESTOQUE_TRANSFERENCIA WHERE "& _
		" (id = '" & transf_estoque & "'))) "
	cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

    s = "UPDATE t_ESTOQUE_TRANSFERENCIA_ITEM SET st_exclusao = 1 WHERE"  & _
		" (id_estoque_transferencia in (SELECT id FROM t_ESTOQUE_TRANSFERENCIA WHERE "& _
		" (id = '" & transf_estoque & "'))) "
	cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if

'	TENTA ELIMINAR O REGISTRO DA TRANSFERÊNCIA

    s = "UPDATE t_ESTOQUE_TRANSFERENCIA SET st_exclusao = 1 WHERE" & _
		" (id = '" & transf_estoque & "') "
	cn.Execute(s)
	if Err <> 0 then
		msg_erro = Cstr(Err) & ": " & Err.Description
		exit function
		end if


	info_log = s_log_base & ":" & info_log
	
	transferencia_estoque_remove = True
end function


' --------------------------------------------------------------------
'   ESTOQUE PRODUTO VENDIDO SEM PRESENCA SAIDA V2
'   Revisado p/ controle de estoque por empresa (auto-split)
'   Retorno da função:
'      False - Ocorreu falha ao tentar alterar os dados do estoque.
'      True - Conseguiu alterar os dados do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função processa o produto que consta na lista de produtos
'	vendidos sem presença no estoque de modo a alocar para ele as
'	unidades que já estejam disponíveis.
'	A quantidade que faltar irá continuar constando da lista de 
'	produtos vendidos sem presença no estoque.
function estoque_produto_vendido_sem_presenca_saida_v2(byval id_usuario, byval id_pedido, _
													byval id_fabricante, byval id_produto, _
													byref qtde_estoque_vendido, _
													byref qtde_estoque_sem_presenca, _
													byref msg_erro)
dim s
dim iv
dim rs
dim v_estoque
dim qtde_a_sair
dim qtde_disponivel
dim qtde_movimentada
dim qtde_movto
dim qtde_aux
dim qtde_utilizada_aux
dim s_chave
dim id_nfe_emitente

	estoque_produto_vendido_sem_presenca_saida_v2 = False
	
	msg_erro = ""
	qtde_estoque_vendido = 0
	qtde_estoque_sem_presenca = 0
	
	If (Trim(id_produto) = "") Or (Trim(id_pedido) = "") Then
		estoque_produto_vendido_sem_presenca_saida_v2 = True
		exit function
		end if

'	OBTÉM A QUANTIDADE VENDIDA SEM PRESENÇA NO ESTOQUE
	s = "SELECT" & _
			" Sum(qtde) AS total" & _
		" FROM t_ESTOQUE_MOVIMENTO" & _
		" WHERE" & _
			" (anulado_status = 0)" & _
			" AND (estoque='" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
			" AND (pedido='" & id_pedido & "')" & _
			" AND (fabricante='" & id_fabricante & "')" & _
			" AND (produto='" & id_produto & "')"
	set rs = cn.Execute(s)
	if Err<>0 then 
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	qtde_a_sair = 0
	if Not rs.Eof then
		if IsNumeric(rs("total")) then qtde_a_sair = CLng(rs("total"))
		end if
	
	If qtde_a_sair <= 0 Then
	'	NÃO HÁ PRODUTOS PENDENTES NA LISTA DE "SEM PRESENÇA"
		estoque_produto_vendido_sem_presenca_saida_v2 = True
		exit function
		end if

'	OBTÉM A EMPRESA (CD) DO PEDIDO
	s = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" & id_pedido & "')"
	if rs.State <> 0 then rs.Close
	set rs = cn.Execute(s)
	if Err<>0 then 
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	if rs.Eof then
		msg_erro="Falha ao tentar localizar o registro do pedido " & id_pedido
		exit function
	else
		id_nfe_emitente = CLng(rs("id_nfe_emitente"))
		end if

'	OBTÉM OS "LOTES" DO PRODUTO DISPONÍVEIS NO ESTOQUE (POLÍTICA FIFO)
	ReDim v_estoque(0)
	v_estoque(UBound(v_estoque)) = ""

	s = "SELECT" & _
			" t_ESTOQUE.id_estoque," & _
			" (qtde - qtde_utilizada) AS saldo" & _
		" FROM t_ESTOQUE" & _
			" INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
		" WHERE" & _
			" (t_ESTOQUE.id_nfe_emitente = " & id_nfe_emitente & ") AND" & _
			" (t_ESTOQUE.fabricante='" & id_fabricante & "') AND" & _
			" (produto='" & id_produto & "') AND" & _
			" ((qtde - qtde_utilizada) > 0)" & _
		" ORDER BY" & _
			" data_entrada," & _
			" t_ESTOQUE.id_estoque"
	if rs.State <> 0 then rs.Close
	set rs = cn.Execute(s)
	if Err<>0 then 
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

    qtde_disponivel = 0
    do while Not rs.Eof
	'	ARMAZENA AS ENTRADAS NO ESTOQUE CANDIDATAS À SAÍDA DE PRODUTOS
		If v_estoque(UBound(v_estoque)) <> "" Then
          ReDim Preserve v_estoque(UBound(v_estoque) + 1)
          v_estoque(UBound(v_estoque)) = ""
          End If
      v_estoque(UBound(v_estoque)) = Trim("" & rs("id_estoque"))
      qtde_disponivel = qtde_disponivel + CLng(rs("saldo"))
      rs.MoveNext 
      Loop

'	HÁ PRODUTOS DISPONÍVEIS?
	if qtde_disponivel <= 0 then
	'   RETORNA TRUE!!
        estoque_produto_vendido_sem_presenca_saida_v2 = True
        exit function
        end if
		
	if rs.State <> 0 then rs.Close
	set rs = nothing
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function


'	REALIZA A SAÍDA DO ESTOQUE!!
    qtde_movimentada = 0
    For iv = LBound(v_estoque) To UBound(v_estoque)
        If Trim(v_estoque(iv)) <> "" Then
          ' A QUANTIDADE NECESSÁRIA JÁ FOI RETIRADA DO ESTOQUE!!
            If qtde_movimentada >= qtde_a_sair Then Exit For

          ' T_ESTOQUE_ITEM: SAÍDA DE PRODUTOS
            s = "SELECT" & _
					" qtde," & _
					" qtde_utilizada," & _
					" data_ult_movimento" & _
				" FROM t_ESTOQUE_ITEM" & _
				" WHERE" & _
					" (id_estoque = '" & Trim(v_estoque(iv)) & "')" & _
					" AND (fabricante = '" & id_fabricante & "')" & _
					" AND (produto = '" & id_produto & "')"
			qtde_movto = 0
			qtde_aux = 0
			qtde_utilizada_aux = 0
			rs.Open s, cn
			if Not rs.EOF then
				qtde_aux = CLng(rs("qtde"))
				qtde_utilizada_aux = CLng(rs("qtde_utilizada"))
				If (qtde_a_sair - qtde_movimentada) > (qtde_aux - qtde_utilizada_aux) Then
				  ' QUANTIDADE DE PRODUTOS DESTE ITEM DE ESTOQUE É INSUFICIENTE P/ ATENDER O PEDIDO
					qtde_movto = qtde_aux - qtde_utilizada_aux
				Else
				  ' QUANTIDADE DE PRODUTOS DESTE ITEM SOZINHO É SUFICIENTE P/ ATENDER O PEDIDO
					qtde_movto = qtde_a_sair - qtde_movimentada
					End If
				rs("qtde_utilizada") = qtde_utilizada_aux + qtde_movto
				rs("data_ult_movimento") = Date
				rs.Update 
				if Err<>0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				end if
			if rs.State <> 0 then rs.Close
				
          ' CONTABILIZA QUANTIDADE MOVIMENTADA
			qtde_movimentada = qtde_movimentada + qtde_movto

          ' REGISTRA O MOVIMENTO DE SAÍDA NO ESTOQUE
			if Not gera_id_estoque_movto(s_chave, msg_erro) then 
				msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
				exit function
				end if

			s = "INSERT INTO t_ESTOQUE_MOVIMENTO (" & _
					"id_movimento, data, hora, usuario, id_estoque, fabricante, produto," & _
					" qtde, operacao, estoque, pedido, loja, kit, kit_id_estoque" & _
				") VALUES (" & _
					"'" & s_chave & "'," & _
					bd_formata_data(Date) & "," & _
					"'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
					"'" & id_usuario & "'," & _
					"'" & Trim(v_estoque(iv)) & "'," & _
					"'" & id_fabricante & "'," & _
					"'" & id_produto & "'," & _
					CStr(qtde_movto) & "," & _
					"'" & OP_ESTOQUE_VENDA & "'," & _
					"'" & ID_ESTOQUE_VENDIDO & "'," & _
					"'" & id_pedido & "'," & _
					"'', 0, ''" & _
				")"
			cn.Execute(s)
			if Err<>0 then 
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

          ' T_ESTOQUE: ATUALIZA DATA DO ÚLTIMO MOVIMENTO
			s = "SELECT" & _
					" data_ult_movimento" & _
				" FROM t_ESTOQUE" & _
				" WHERE" & _
					" (id_estoque = '" & v_estoque(iv) & "')"
			rs.Open s, cn
			if Not rs.EOF then
				rs("data_ult_movimento") = Date
				rs.Update 
				if Err<>0 then 
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				end if
			if rs.State <> 0 then rs.Close

		'	JÁ CONSEGUIU ALOCAR TUDO?
			If qtde_movimentada >= qtde_a_sair Then Exit For
			end if
		next


'	ANULA O REGISTRO DO PRODUTO DESTE PEDIDO NA LISTA "SEM PRESENÇA NO ESTOQUE"
	s = "UPDATE t_ESTOQUE_MOVIMENTO SET" & _
			" anulado_status=1" & _
			", anulado_data=" & bd_formata_data(Date) & _
			", anulado_hora='" & retorna_so_digitos(formata_hora(Now)) & "'" & _
			", anulado_usuario='" & id_usuario & "'" & _
		" WHERE" & _
			" (anulado_status = 0)" & _
			" AND (estoque='" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
			" AND (pedido='" & id_pedido & "')" & _
			" AND (fabricante='" & id_fabricante & "')" & _
			" AND (produto='" & id_produto & "')"
	cn.Execute(s)
	if Err<>0 then 
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
		
'   RESÍDUO FALTANTE: REGISTRA A VENDA SEM PRESENÇA NO ESTOQUE P/ A DIFERENÇA QUE AINDA FALTA
	if qtde_movimentada < qtde_a_sair then
	'	REGISTRA O MOVIMENTO DE SAÍDA NO ESTOQUE
		if Not gera_id_estoque_movto(s_chave, msg_erro) then 
			msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
			exit function
			end if

		qtde_estoque_sem_presenca = qtde_a_sair - qtde_movimentada
		s = "INSERT INTO t_ESTOQUE_MOVIMENTO (" & _
				"id_movimento, data, hora, usuario, id_estoque, fabricante, produto," & _
				" qtde, operacao, estoque, pedido, loja, kit, kit_id_estoque" & _
			") VALUES (" & _
				"'" & s_chave & "'," & _
				bd_formata_data(Date) & "," & _
				"'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
				"'" & id_usuario & "'," & _
				"''," & _
				"'" & id_fabricante & "'," & _
				"'" & id_produto & "'," & _
				CStr(qtde_estoque_sem_presenca) & "," & _
				"'" & OP_ESTOQUE_VENDA & "'," & _
				"'" & ID_ESTOQUE_SEM_PRESENCA & "'," & _
				"'" & id_pedido & "'," & _
				"'', 0, ''" & _
			")"
		cn.Execute(s)
		if Err<>0 then 
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		end if
			
	qtde_estoque_vendido = qtde_movimentada
	
	'Log de movimentação do estoque
	if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, id_fabricante, id_produto, qtde_a_sair, qtde_estoque_vendido, OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA, ID_ESTOQUE_SEM_PRESENCA, ID_ESTOQUE_VENDIDO, "", "", id_pedido, id_pedido, "", "", "") then
		msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
		exit function
		end if
	
	estoque_produto_vendido_sem_presenca_saida_v2 = True
end function



' --------------------------------------------------------------------
'   ESTOQUE PROCESSA PRODUTOS VENDIDOS SEM PRESENCA V2
'   Revisado p/ controle de estoque por empresa (auto-split)
'   Retorno da função:
'      False - Ocorreu falha ao tentar alterar os dados do estoque.
'      True - Conseguiu alterar os dados do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função verifica a lista de produtos que foram vendidos sem
'	presença no estoque para alocar os produtos que já estejam dis-
'	poníveis aos pedidos mais antigos primeiro.
'	O log da movimentação do estoque (T_ESTOQUE_LOG) é gravado
'	dentro das rotinas chamadas por esta rotina:
'		1) estoque_produto_vendido_sem_presenca_saida_v2()
'	Se o parâmetro 'id_nfe_emitente' for igual a zero ou nulo, serão
'	processados os estoques de todos os CD's, caso contrário, somente
'	o estoque do CD especificado.
function estoque_processa_produtos_vendidos_sem_presenca_v2(byval id_nfe_emitente, byval id_usuario, byref msg_erro)
dim rs, s, v, v_pedido, achou, i, j, s_log, s_log_aux
dim qtde_estoque_sem_presenca, qtde_estoque_vendido
dim total_estoque_sem_presenca, total_estoque_vendido

	estoque_processa_produtos_vendidos_sem_presenca_v2 = False

	id_nfe_emitente = converte_numero(Trim("" & id_nfe_emitente))

	msg_erro = ""
	s_log = ""
	s = "SELECT" & _
			" t_ESTOQUE_MOVIMENTO.pedido," & _
			" t_ESTOQUE_MOVIMENTO.fabricante," & _
			" t_ESTOQUE_MOVIMENTO.produto" & _
		" FROM t_ESTOQUE_MOVIMENTO" & _
			" INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
			" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" & _
		" WHERE" & _
			" (anulado_status=0)" & _
			" AND (estoque='" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
			" AND ((t_ESTOQUE_ITEM.qtde - t_ESTOQUE_ITEM.qtde_utilizada) > 0)"

	if id_nfe_emitente > 0 then
		s = s & _
			" AND (t_PEDIDO.id_nfe_emitente = " & id_nfe_emitente & ")"
		end if

	s = s & _
		" ORDER BY" & _
			" t_PEDIDO.data," & _
			" t_PEDIDO.hora"
	set rs = cn.Execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

'	NÃO HÁ O QUE PROCESSAR!!
	if rs.Eof then
		estoque_processa_produtos_vendidos_sem_presenca_v2 = True
		exit function
		end if
		
	redim v(0)
	set v(Ubound(v)) = New cl_PROCESSA_PRODUTOS_VENDIDOS_SEM_PRESENCA
	v(Ubound(v)).pedido = ""

	do while Not rs.Eof
		if v(Ubound(v)).pedido <> "" then
			redim preserve v(Ubound(v)+1)
			set v(Ubound(v)) = New cl_PROCESSA_PRODUTOS_VENDIDOS_SEM_PRESENCA
			end if
		with v(Ubound(v))
			.pedido = Trim("" & rs("pedido"))
			.fabricante = Trim("" & rs("fabricante"))
			.produto = Trim("" & rs("produto"))
			end with
		rs.movenext
		loop
	if rs.State <> 0 then rs.Close
		
	redim v_pedido(0)
	v_pedido(Ubound(v_pedido)) = ""
	
'	OS PEDIDOS MAIS ANTIGOS DEVEM SER ATENDIDOS PRIMEIRO
	for i = Lbound(v) to Ubound(v)
		with v(i)
			if .pedido <> "" then
				if Not estoque_produto_vendido_sem_presenca_saida_v2(id_usuario, .pedido, .fabricante, .produto, qtde_estoque_vendido, qtde_estoque_sem_presenca, msg_erro) then 
					exit function
					end if
					
			'   SE HOUVE PRODUTO ALOCADO P/ ESTE PEDIDO, ENTÃO INCLUI O PEDIDO NA LISTA QUE SERÁ ANALISADA QUANTO AO "STATUS DE ENTREGA"
				if qtde_estoque_vendido > 0 then
					achou=False
					for j = Ubound(v_pedido) to Lbound(v_pedido) step -1
						if v_pedido(j) = .pedido then
							achou = True
							exit for
							end if
						next
					
					if Not achou then
						if v_pedido(Ubound(v_pedido)) <> "" then
							redim preserve v_pedido(Ubound(v_pedido)+1)
							end if
						v_pedido(Ubound(v_pedido)) = .pedido
						end if
					
				'	INFORMAÇÕES PARA O LOG
					s = .pedido & log_produto_monta(qtde_estoque_vendido, .fabricante, .produto) & " SPE=" & Cstr(qtde_estoque_sem_presenca)
					if s_log <> "" then s_log = s_log & "; "
					s_log = s_log & s
					end if
				end if
			end with
		next

'   ATUALIZA O "STATUS DE ENTREGA" DOS PEDIDOS
	if rs.State <> 0 then rs.Close
	set rs = nothing
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

	for i = Lbound(v_pedido) to Ubound(v_pedido)
		if v_pedido(i) <> "" then
			total_estoque_sem_presenca = 0
			s = "SELECT" & _
					" Sum(qtde) AS total" & _
				" FROM t_ESTOQUE_MOVIMENTO" & _
				" WHERE" & _
					" (anulado_status=0)" & _
					" AND (estoque = '" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
					" AND (pedido = '" & v_pedido(i) & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if Not rs.Eof then
				if IsNumeric(rs("total")) then total_estoque_sem_presenca = CLng(rs("total"))
				end if
			
			total_estoque_vendido = 0
			s = "SELECT" & _
					" Sum(qtde) AS total" & _
				" FROM t_ESTOQUE_MOVIMENTO" & _
				" WHERE" & _
					" (anulado_status=0)" & _
					" AND (estoque = '" & ID_ESTOQUE_VENDIDO & "')" & _
					" AND (pedido = '" & v_pedido(i) & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if Not rs.Eof then
				if IsNumeric(rs("total")) then total_estoque_vendido = CLng(rs("total"))
				end if
			
			s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & v_pedido(i) & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				msg_erro = "Pedido " & v_pedido(i) & " não foi encontrado."
				exit function
			else
			'	STATUS DE ENTREGA
				if total_estoque_vendido = 0 then
					s = ST_ENTREGA_ESPERAR
				elseif total_estoque_sem_presenca = 0 then
					s = ST_ENTREGA_SEPARAR
				else
					s = ST_ENTREGA_SPLIT_POSSIVEL
					end if
				
				if Trim("" & rs("st_entrega")) <> s then
					rs("st_entrega") = s
					rs.Update
					if Err <> 0 then
						msg_erro=Cstr(Err) & ": " & Err.Description
						exit function
						end if
					end if
				end if
			end if
		next
		
	if s_log <> "" then 
		s_log_aux = "Processamento automático da lista de produtos vendidos sem presença no estoque"
		if id_nfe_emitente > 0 then
			s_log_aux = s_log_aux & " (id_nfe_emitente = " & id_nfe_emitente & ")"
		else
			s_log_aux = s_log_aux & " (id_nfe_emitente = N.I.)"
			end if
		s_log = s_log_aux & ": " & s_log
		grava_log id_usuario, "", "", "", OP_LOG_ESTOQUE_PROCESSA_SP, s_log
		end if
	
	estoque_processa_produtos_vendidos_sem_presenca_v2 = True
end function



' --------------------------------------------------------------------
'   ESTOQUE PRODUTO SAIDA POR TRANSFERENCIA V2
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função processa a saída dos produtos do "estoque de venda"
'   para o estoque especificado pelo parâmetro "id_estoque_destino".
function estoque_produto_saida_por_transferencia_v2(byval id_usuario, _
												 byval id_estoque_destino, _
												 byval id_loja_destino, _
												 byval id_nfe_emitente, _
												 byval id_fabricante, byval id_produto, _
												 byval qtde_a_sair, _
												 byval id_ordem_servico_destino, _
												 ByVal id_pedido_destino, _
												 byref msg_erro)
dim s_sql
dim s_chave
dim qtde_disponivel
Dim v_estoque()
Dim iv
dim rs
Dim qtde_aux
Dim qtde_utilizada_aux
Dim qtde_movto
Dim qtde_movimentada

	estoque_produto_saida_por_transferencia_v2=False

	msg_erro=""
		
    id_usuario = Trim("" & id_usuario)
    id_estoque_destino = Trim("" & id_estoque_destino)
    id_loja_destino = Trim("" & id_loja_destino)
    id_fabricante = Trim("" & id_fabricante)
    id_produto = Trim("" & id_produto)
    id_ordem_servico_destino = Trim("" & id_ordem_servico_destino)
    if id_ordem_servico_destino <> "" then id_ordem_servico_destino=normaliza_codigo(retorna_so_digitos(id_ordem_servico_destino), TAM_MAX_NSU)
    id_pedido_destino = Trim("" & id_pedido_destino)
	id_nfe_emitente = converte_numero(id_nfe_emitente)
    
    If (qtde_a_sair <= 0) Or (id_produto = "") Then
        estoque_produto_saida_por_transferencia = True
        exit function
        end if

	if id_estoque_destino = "" then
		msg_erro = "Estoque de destino da transferência é inválido."
		exit function
		end if
	
	if id_nfe_emitente = 0 then
		msg_erro = "Não foi informado o CD"
		exit function
		end if

'	OBTÉM OS "LOTES" DO PRODUTO DISPONÍVEIS NO ESTOQUE (POLÍTICA FIFO)
	s_sql = "SELECT" & _
				" t_ESTOQUE.id_estoque," & _
				" (qtde - qtde_utilizada) AS saldo" & _
			" FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON" & _
				" (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
			" WHERE" & _
				" (t_ESTOQUE.id_nfe_emitente = " & id_nfe_emitente & ") AND" & _
				" (t_ESTOQUE.fabricante='" & id_fabricante & "') AND" & _
				" (produto='" & id_produto & "') AND" & _
				" ((qtde - qtde_utilizada) > 0)" & _
			" ORDER BY data_entrada, t_ESTOQUE.id_estoque"

    ReDim v_estoque(0)
    v_estoque(UBound(v_estoque)) = ""

    set rs=cn.Execute(s_sql)
    if Err<>0 then 
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
    qtde_disponivel = 0
    do while Not rs.Eof
	'	ARMAZENA AS ENTRADAS NO ESTOQUE CANDIDATAS À SAÍDA DE PRODUTOS
		If v_estoque(UBound(v_estoque)) <> "" Then
          ReDim Preserve v_estoque(UBound(v_estoque) + 1)
          v_estoque(UBound(v_estoque)) = ""
          End If
      v_estoque(UBound(v_estoque)) = Trim("" & rs("id_estoque"))
      qtde_disponivel = qtde_disponivel + CLng(rs("saldo"))
      rs.MoveNext 
      Loop

'	NÃO HÁ PRODUTOS SUFICIENTES NO ESTOQUE!!
    if qtde_a_sair > qtde_disponivel then 
		msg_erro="Produto " & id_produto & " do fabricante " & id_fabricante & ": faltam " & _
				Cstr(qtde_a_sair-qtde_disponivel) & " unidades no estoque."
		exit function
		end if

	if rs.State <> 0 then rs.Close
	set rs=nothing
	
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

'	REALIZA A SAÍDA (TRANSFERÊNCIA) DO ESTOQUE!!
	qtde_movimentada = 0
	For iv = LBound(v_estoque) To UBound(v_estoque)
		If Trim(v_estoque(iv)) <> "" Then
		  ' A QUANTIDADE NECESSÁRIA JÁ FOI RETIRADA DO ESTOQUE!!
			If qtde_movimentada >= qtde_a_sair Then Exit For

		  ' T_ESTOQUE_ITEM: SAÍDA DE PRODUTOS
			s_sql = "SELECT" & _
						" qtde," & _
						" qtde_utilizada," & _
						" data_ult_movimento" & _
					" FROM t_ESTOQUE_ITEM" & _
					" WHERE" & _
						" (id_estoque = '" & Trim(v_estoque(iv)) & "')" & _
						" AND (fabricante = '" & id_fabricante & "')" & _
						" AND (produto = '" & id_produto & "')"
			qtde_movto=0
			qtde_aux = 0
			qtde_utilizada_aux = 0
			rs.Open s_sql, cn
			if Not rs.EOF then
				qtde_aux = CLng(rs("qtde"))
				qtde_utilizada_aux = CLng(rs("qtde_utilizada"))
				If (qtde_a_sair - qtde_movimentada) > (qtde_aux - qtde_utilizada_aux) Then
				  ' QUANTIDADE DE PRODUTOS DESTE ITEM DE ESTOQUE É INSUFICIENTE
					qtde_movto = qtde_aux - qtde_utilizada_aux
				Else
				  ' QUANTIDADE DE PRODUTOS DESTE ITEM SOZINHO É SUFICIENTE
					qtde_movto = qtde_a_sair - qtde_movimentada
					End If
				rs("qtde_utilizada")=qtde_utilizada_aux + qtde_movto
				rs("data_ult_movimento")=Date
				rs.Update 
				if Err<>0 then 
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				end if
			if rs.State <> 0 then rs.Close
			
		  ' CONTABILIZA QUANTIDADE MOVIMENTADA
			qtde_movimentada = qtde_movimentada + qtde_movto

		  ' REGISTRA O MOVIMENTO DE SAÍDA NO ESTOQUE
			if Not gera_id_estoque_movto(s_chave, msg_erro) then 
				msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
				exit function
				end if

			s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO (" & _
						"id_movimento," & _
						" data," & _
						" hora," & _
						" usuario," & _
						" id_estoque," & _
						" fabricante," & _
						" produto," & _
						" qtde," & _
						" operacao," & _
						" estoque," & _
						" pedido," & _
						" loja," & _
						" kit," & _
						" kit_id_estoque," & _
						" id_ordem_servico" & _
					") VALUES (" & _
						"'" & s_chave & "'," & _
						bd_formata_data(Date) & "," & _
						"'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
						"'" & id_usuario & "'," & _
						"'" & Trim(v_estoque(iv)) & "'," & _
						"'" & id_fabricante & "'," & _
						"'" & id_produto & "'," & _
						CStr(qtde_movto) & "," & _
						"'" & OP_ESTOQUE_TRANSFERENCIA & "'," & _
						"'" & id_estoque_destino & "'," & _
						"'" & id_pedido_destino & "'," & _
						"'" & id_loja_destino & "'," & _
						"0, " & _
						"'', " & _
						"'" & id_ordem_servico_destino & "'" & _
					")"
			cn.Execute(s_sql)
			if Err<>0 then 
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

		  ' T_ESTOQUE: ATUALIZA DATA DO ÚLTIMO MOVIMENTO
			s_sql = "SELECT data_ult_movimento FROM t_ESTOQUE WHERE" & _
					" (id_estoque = '" & v_estoque(iv) & "')"

			rs.Open s_sql, cn
			if Not rs.EOF then
				rs("data_ult_movimento")=Date
				rs.Update 
				if Err<>0 then 
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				end if
			if rs.State <> 0 then rs.Close

		  ' JÁ CONSEGUIU ALOCAR TUDO?
			If qtde_movimentada >= qtde_a_sair Then Exit For
			end if
		next

	
'   NÃO CONSEGUIU MOVIMENTAR A QUANTIDADE SUFICIENTE	
	if qtde_movimentada < qtde_a_sair then 
		msg_erro="Produto " & id_produto & " do fabricante " & id_fabricante & ": faltam " & _
				 Cstr(qtde_a_sair-qtde_movimentada) & " unidades no estoque."
		exit function
		end if
	
	'Log de movimentação do estoque
	if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, id_fabricante, id_produto, qtde_a_sair, qtde_movimentada, OP_ESTOQUE_LOG_TRANSFERENCIA, ID_ESTOQUE_VENDA, id_estoque_destino, "", id_loja_destino, "", "", "", "", id_ordem_servico_destino) then
		msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
		exit function
		end if
	
	estoque_produto_saida_por_transferencia_v2=True
end function



' --------------------------------------------------------------------
'   ESTOQUE PRODUTO ESTORNA POR TRANSFERENCIA V2
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função estorna a quantidade de produtos indicada pelo 
'   parâmetro "qtde_a_estornar" do estoque especificado pelo 
'	parâmetro "id_estoque_origem" de volta para o "estoque de venda".
'   Se o parâmetro "qtde_a_estornar" for especificado com o valor
'   "COD_NEGATIVO_UM", então o estorno será integral.
function estoque_produto_estorna_por_transferencia_v2(ByVal id_usuario, _
												   ByVal id_estoque_origem, _
												   ByVal id_loja, _
												   ByVal id_pedido, _
												   ByVal id_nfe_emitente, _
												   ByVal id_fabricante, ByVal id_produto, _
												   ByVal qtde_a_estornar, ByRef qtde_estornada, _
												   ByVal id_ordem_servico, _
												   ByRef msg_erro)
dim iv
dim rs
dim s_chave
dim s_sql
dim v_estoque
dim id_estoque_aux
dim qtde_aux
dim qtde_utilizada_aux
dim qtde_movto
dim operacao_aux
dim estoque_aux
dim loja_aux
dim id_ordem_servico_aux
dim pedido_aux
dim blnGravarLog

	estoque_produto_estorna_por_transferencia_v2 = False
	msg_erro = ""
	qtde_estornada = 0

	id_usuario = Trim("" & id_usuario)
	id_estoque_origem = Trim("" & id_estoque_origem)
	id_loja = Trim("" & id_loja)
	id_pedido = Trim("" & id_pedido)
	id_fabricante = Trim("" & id_fabricante)
	id_produto = Trim("" & id_produto)
	id_ordem_servico = Trim("" & id_ordem_servico)
	if id_ordem_servico <> "" then id_ordem_servico=normaliza_codigo(retorna_so_digitos(id_ordem_servico), TAM_MAX_NSU)
	id_nfe_emitente = converte_numero(id_nfe_emitente)

  ' 1) LEMBRE-SE DE QUE PODE HAVER MAIS DE UM REGISTRO EM T_ESTOQUE_MOVIMENTO 
  '    P/ CADA PRODUTO, POIS PODEM TER SIDO USADOS DIFERENTES LOTES DO ESTOQUE 
  '    P/ ATENDER A UM ÚNICO PEDIDO!!
  ' 2) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
  '    OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
  '    FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	ReDim v_estoque(0)
	v_estoque(UBound(v_estoque)) = ""
	
	s_sql = "SELECT" & _
				" id_movimento" & _
			" FROM t_ESTOQUE" & _
				" INNER JOIN t_ESTOQUE_MOVIMENTO ON (t_ESTOQUE.id_estoque = t_ESTOQUE_MOVIMENTO.id_estoque)" & _
			" WHERE" & _
				" (anulado_status = 0)" & _
				" AND (estoque = '" & id_estoque_origem & "')" & _
				" AND (t_ESTOQUE_MOVIMENTO.fabricante = '" & id_fabricante & "')" & _
				" AND (produto = '" & id_produto & "')"
	if id_nfe_emitente <> 0 then s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente = " & id_nfe_emitente & ")"
	if id_loja <> "" then s_sql = s_sql & " AND (loja = '" & id_loja & "')"
	if id_pedido <> "" then s_sql = s_sql & " AND (pedido = '" & id_pedido & "')"
	if id_ordem_servico <> "" then s_sql = s_sql & " AND (id_ordem_servico = '" & id_ordem_servico & "')"

	if id_estoque_origem = ID_ESTOQUE_ROUBO then
	'	NO CASO DE ESTORNO DE ROUBO/PERDA, ESTORNA P/ OS ESTOQUES MAIS RECENTES PRIMEIRO.
	'	ISSO EVITA O PROBLEMA DE DISTORÇÃO NO CÁLCULO DO CMV, POIS AO ESTORNAR UM PRODUTO DO ROUBO/PERDA, PROCESSAR OS ESTOQUES MAIS ANTIGOS PRIMEIRO PODE RESULTAR EM VALORES DE AQUISIÇÃO MUITO MAIS BAIXOS.
	'	ALÉM DISSO, EM SITUAÇÕES EM QUE CÓDIGOS DE PRODUTOS SÃO REAPROVEITADOS, CORRE-SE O RISCO DE RESTAURAR O ESTOQUE DE PRODUTOS QUE ERAM DIFERENTES DO ATUAL, SENDO QUE TAL SITUAÇÃO JÁ CHEGOU A OCORRER.
		s_sql = s_sql & _
				" ORDER BY" & _
					" t_ESTOQUE.data_entrada DESC," & _
					" t_ESTOQUE.id_estoque DESC"
	else
		s_sql = s_sql & _
				" ORDER BY" & _
					" t_ESTOQUE.data_entrada," & _
					" t_ESTOQUE.id_estoque"
		end if

	set rs=cn.execute(s_sql)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	do while Not rs.EOF 
		If v_estoque(UBound(v_estoque)) <> "" Then
			ReDim Preserve v_estoque(UBound(v_estoque) + 1)
			v_estoque(UBound(v_estoque)) = ""
			End If
		v_estoque(UBound(v_estoque)) = Trim("" & rs("id_movimento"))
		rs.MoveNext 
		loop
		
	if rs.State <> 0 then rs.Close
	set rs=nothing
	
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function
			
	for iv=LBound(v_estoque) To UBound(v_estoque)
		If Trim(v_estoque(iv)) <> "" Then
		  
		  ' JÁ ESTORNOU TUDO?
			If qtde_a_estornar <> COD_NEGATIVO_UM Then
				If qtde_estornada >= qtde_a_estornar Then Exit For
				End If
			
		  ' T_ESTOQUE_MOVIMENTO: ANULA O MOVIMENTO	
		  ' ======================================
			s_sql = "SELECT " & _
						"*" & _
					" FROM t_ESTOQUE_MOVIMENTO" & _
					" WHERE" & _
						" (anulado_status = 0)" & _
						" AND (id_movimento = '" & Trim(v_estoque(iv)) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro="Falha ao acessar o registro de movimento no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				exit function
				end if

			id_estoque_aux = Trim("" & rs("id_estoque"))
			qtde_aux = CLng(rs("qtde"))
			operacao_aux = Trim("" & rs("operacao"))
			estoque_aux = Trim("" & rs("estoque"))
			loja_aux = Trim("" & rs("loja"))
			pedido_aux = Trim("" & rs("pedido"))
			id_ordem_servico_aux = Trim("" & rs("id_ordem_servico"))
			
			qtde_movto = qtde_aux
			
		  ' É PARA ESTORNAR TUDO OU UMA QUANTIDADE ESPECIFICADA?
			If qtde_a_estornar <> COD_NEGATIVO_UM Then
			  ' A QUANTIDADE QUE FALTA SER ESTORNADA É MENOR QUE A QUANTIDADE DO MOVIMENTO
				If (qtde_a_estornar - qtde_estornada) < qtde_aux Then
					qtde_movto = qtde_a_estornar - qtde_estornada
					End If
				End If
			
		  ' ANULA O MOVIMENTO
			rs("anulado_status") = 1
			rs("anulado_data") = Date
			rs("anulado_hora") = retorna_so_digitos(formata_hora(Now))
			rs("anulado_usuario") = id_usuario
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

		  ' ESTORNO PARCIAL: O MOVIMENTO ORIGINAL FOI ANULADO E UM NOVO MOVIMENTO 
		  ' C/ A QUANTIDADE RESTANTE DEVE SER GRAVADO!!
			If qtde_movto < qtde_aux Then
			  ' REGISTRA O MOVIMENTO DE SAÍDA NO ESTOQUE
				if Not gera_id_estoque_movto(s_chave, msg_erro) then 
					msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
					exit function
					end if
				
				s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
						" (id_movimento, data, hora, usuario, pedido, fabricante, produto, id_estoque," & _
						" qtde, operacao, estoque, loja, kit, kit_id_estoque, id_ordem_servico) VALUES (" & _
						"'" & s_chave & "'," & _
						bd_formata_data(Date) & "," & _
						"'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
						"'" & id_usuario & "'," & _
						"'" & pedido_aux & "'," & _
						"'" & id_fabricante & "'," & _
						"'" & id_produto & "'," & _
						"'" & id_estoque_aux & "'," & _
						CStr(qtde_aux - qtde_movto) & "," & _
						"'" & operacao_aux & "'," & _
						"'" & estoque_aux & "'," & _
						"'" & loja_aux & "'," & _
						"0, ''," & _
						"'" & id_ordem_servico_aux & "'" & _
						")"
				cn.Execute(s_sql)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				End If
			
		  
		  ' T_ESTOQUE_ITEM: ESTORNA PRODUTOS AO SALDO
		  ' =========================================
			s_sql = "SELECT" & _
						" data_ult_movimento," & _
						" qtde_utilizada" & _
					" FROM t_ESTOQUE_ITEM" & _
					" WHERE" & _
						" (id_estoque = '" & id_estoque_aux & "') AND" & _
						" (fabricante = '" & id_fabricante & "') AND" & _
						" (produto = '" & id_produto & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro = "Falha ao acessar o registro no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				exit function
				end if
			
			qtde_utilizada_aux = CLng(rs("qtde_utilizada"))

		  ' PRECAUÇÃO (P/ GARANTIR QUE "QTDE_UTILIZADA" NUNCA FICARÁ C/ VALOR NEGATIVO)!!
			If qtde_utilizada_aux < qtde_movto Then
				msg_erro = "Falha ao processar o estorno ao estoque por transferência: a quantidade utilizada do estoque é menor do que o esperado (id_estoque=" & id_estoque_aux & "; fabricante=" & id_fabricante & "; produto=" & id_produto & "; qtde_utilizada=" & qtde_utilizada_aux & "; qtde estorno=" & qtde_movto & ")"
				exit function
				end if
			
			rs("qtde_utilizada") = rs("qtde_utilizada") - qtde_movto
			rs("data_ult_movimento") = Date
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
		  
		  ' CONTABILIZA QUANTIDADE ESTORNADA
			qtde_estornada = qtde_estornada + qtde_movto
		
		  
		  ' T_ESTOQUE: ATUALIZA DATA DO ÚLTIMO MOVIMENTO
		  ' ============================================
			s_sql = "SELECT" & _
						" data_ult_movimento" & _
					" FROM t_ESTOQUE" & _
					" WHERE" & _
						" (id_estoque = '" & id_estoque_aux & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro = "Falha ao acessar o registro principal no estoque do produto " & id_produto & " do fabricante " & id_fabricante & " (id_estoque=" & id_estoque_aux & ")"
				exit function
				end if
			
			rs("data_ult_movimento") = Date
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			end if
		next

'	CONSEGUIU ESTORNAR TUDO?
	if (qtde_a_estornar <> COD_NEGATIVO_UM) And (qtde_a_estornar > 0) then
		if qtde_estornada < qtde_a_estornar then
			msg_erro="Não foi possível estornar a quantidade solicitada (qtde solicitada = " & qtde_a_estornar & "; qtde estornada = " & qtde_estornada & ")"
			exit function
			end if
		end if

	blnGravarLog=True
	if (qtde_a_estornar = COD_NEGATIVO_UM) And (qtde_estornada = 0) then blnGravarLog=False

	if blnGravarLog then
		'Log de movimentação do estoque
		if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, id_fabricante, id_produto, qtde_a_estornar, qtde_estornada, OP_ESTOQUE_LOG_TRANSFERENCIA, id_estoque_origem, ID_ESTOQUE_VENDA, id_loja, "", id_pedido, "", "", "", "") then
			msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
			exit function
			end if
		end if
		
	estoque_produto_estorna_por_transferencia_v2 = True
end function



' --------------------------------------------------------------------
'   ESTOQUE PRODUTO TRANSFERE ENTRE ESTOQUES V2
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função transfere a quantidade de produtos indicada pelo 
'   parâmetro "qtde_a_transferir" do estoque especificado pelo 
'	parâmetro "cod_estoque_origem" para o estoque "cod_estoque_destino".
'	As informações passadas nos parâmetros "id_loja_origem" e
'	"id_loja_destino" podem estar vazias, pois depende do tipo de
'	estoque sendo movimentado ("show-room" e "devolução" possuem a informação
'	da loja).
function estoque_produto_transfere_entre_estoques_v2(ByVal id_usuario, _
												  ByVal id_nfe_emitente, _
												  ByVal id_fabricante, _
												  ByVal id_produto, _
												  ByVal qtde_a_transferir, _
												  ByRef qtde_transferida, _
												  ByVal cod_estoque_origem, _
												  ByVal id_loja_origem, _
												  ByVal id_ordem_servico_origem, _
												  ByVal id_pedido_origem, _
												  ByVal cod_estoque_destino, _
												  ByVal id_loja_destino, _
												  ByVal id_ordem_servico_destino, _
												  ByVal id_pedido_destino, _
												  ByRef msg_erro)
dim iv
dim rs
dim s_chave
dim s_sql
dim v_estoque
dim id_estoque_aux
dim qtde_aux
dim qtde_movto
dim operacao_aux
dim estoque_aux
dim loja_aux
dim pedido_aux
dim id_ordem_servico_aux
dim id_ordem_servico_log

	estoque_produto_transfere_entre_estoques_v2 = False
	msg_erro = ""
	qtde_transferida = 0

	id_usuario = Trim("" & id_usuario)
	cod_estoque_origem = Trim("" & cod_estoque_origem)
	id_loja_origem = Trim("" & id_loja_origem)
	cod_estoque_destino = Trim("" & cod_estoque_destino)
	id_loja_destino = Trim("" & id_loja_destino)
	id_fabricante = Trim("" & id_fabricante)
	id_produto = Trim("" & id_produto)
	id_ordem_servico_origem = Trim("" & id_ordem_servico_origem)
	if id_ordem_servico_origem <> "" then id_ordem_servico_origem=normaliza_codigo(retorna_so_digitos(id_ordem_servico_origem), TAM_MAX_NSU)
	id_ordem_servico_destino = Trim("" & id_ordem_servico_destino)
	if id_ordem_servico_destino <> "" then id_ordem_servico_destino=normaliza_codigo(retorna_so_digitos(id_ordem_servico_destino), TAM_MAX_NSU)
	id_pedido_origem = Trim("" & id_pedido_origem)
	if id_pedido_origem <> "" then id_pedido_origem = normaliza_num_pedido(id_pedido_origem)
	id_pedido_destino = Trim("" & id_pedido_destino)
	if id_pedido_destino <> "" then id_pedido_destino = normaliza_num_pedido(id_pedido_destino)
	id_nfe_emitente = converte_numero(id_nfe_emitente)

	if qtde_a_transferir <= 0 then
		msg_erro="Quantidade a transferir é inválida (" & Cstr(qtde_a_transferir) & ")"
		exit function
	elseif cod_estoque_origem = "" then
		msg_erro="Não foi informado o estoque de origem"
		exit function
	elseif cod_estoque_destino = "" then
		msg_erro="Não foi informado o estoque de destino"
		exit function
	elseif id_fabricante = "" then
		msg_erro="Não foi informado o fabricante"
		exit function
	elseif id_produto = "" then
		msg_erro="Não foi informado o produto"
		exit function
		end if

  ' 1) LEMBRE-SE DE QUE PODE HAVER MAIS DE UM REGISTRO EM T_ESTOQUE_MOVIMENTO 
  '    P/ CADA PRODUTO, POIS PODEM TER SIDO USADOS DIFERENTES LOTES DO ESTOQUE 
  '    P/ ATENDER A UM ÚNICO PEDIDO/TRANSFERÊNCIA!!
  ' 2) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
  '    OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
  '    FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	ReDim v_estoque(0)
	v_estoque(UBound(v_estoque)) = ""
	
	s_sql = "SELECT" & _
				" id_movimento" & _
			" FROM t_ESTOQUE" & _
				" INNER JOIN t_ESTOQUE_MOVIMENTO ON (t_ESTOQUE.id_estoque = t_ESTOQUE_MOVIMENTO.id_estoque)" & _
			" WHERE" & _
				" (anulado_status = 0)" & _
				" AND (estoque = '" & cod_estoque_origem & "')" & _
				" AND (t_ESTOQUE_MOVIMENTO.fabricante = '" & id_fabricante & "')" & _
				" AND (produto = '" & id_produto & "')"
	if id_nfe_emitente <> 0 then s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente = " & id_nfe_emitente & ")"
	if id_loja_origem <> "" then s_sql = s_sql & " AND (loja = '" & id_loja_origem & "')"
	if id_ordem_servico_origem <> "" then s_sql = s_sql & " AND (id_ordem_servico = '" & id_ordem_servico_origem & "')"
	if id_pedido_origem <> "" then s_sql = s_sql & " AND (pedido = '" & id_pedido_origem & "')"
	s_sql = s_sql & " ORDER BY" & _
					" t_ESTOQUE.data_entrada DESC," & _
					" t_ESTOQUE.id_estoque DESC"

	set rs=cn.execute(s_sql)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	do while Not rs.EOF 
		If v_estoque(UBound(v_estoque)) <> "" Then
			ReDim Preserve v_estoque(UBound(v_estoque) + 1)
			v_estoque(UBound(v_estoque)) = ""
			End If
		v_estoque(UBound(v_estoque)) = Trim("" & rs("id_movimento"))
		rs.MoveNext 
		loop
		
	if rs.State <> 0 then rs.Close
	set rs=nothing
	
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function
			
	for iv=LBound(v_estoque) To UBound(v_estoque)
		If Trim(v_estoque(iv)) <> "" Then
		  
		  ' JÁ TRANSFERIU TUDO?
			If qtde_transferida >= qtde_a_transferir Then Exit For
			
		  ' T_ESTOQUE_MOVIMENTO: ANULA O MOVIMENTO
		  ' ======================================
			s_sql = "SELECT " & _
						"*" & _
					" FROM t_ESTOQUE_MOVIMENTO" & _
					" WHERE" & _
						" (anulado_status = 0)" & _
						" AND (id_movimento = '" & Trim(v_estoque(iv)) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro="Falha ao acessar o registro de movimento no estoque do produto " & id_produto & " do fabricante " & id_fabricante & " (id_movimento=" & Trim(v_estoque(iv)) & ")"
				exit function
				end if

			id_estoque_aux = Trim("" & rs("id_estoque"))
			qtde_aux = CLng(rs("qtde"))
			operacao_aux = Trim("" & rs("operacao"))
			estoque_aux = Trim("" & rs("estoque"))
			loja_aux = Trim("" & rs("loja"))
			pedido_aux = Trim("" & rs("pedido"))
			id_ordem_servico_aux = Trim("" & rs("id_ordem_servico"))
			
			qtde_movto = qtde_aux
			
		  ' A QUANTIDADE QUE FALTA SER TRANSFERIDA É MENOR QUE A QUANTIDADE DO MOVIMENTO
			If (qtde_a_transferir - qtde_transferida) < qtde_aux Then
				qtde_movto = qtde_a_transferir - qtde_transferida
				End If
			
		  ' ANULA O MOVIMENTO
			rs("anulado_status") = 1
			rs("anulado_data") = Date
			rs("anulado_hora") = retorna_so_digitos(formata_hora(Now))
			rs("anulado_usuario") = id_usuario
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

		  ' TRANSFERÊNCIA PARCIAL: O MOVIMENTO ORIGINAL FOI ANULADO E UM NOVO MOVIMENTO 
		  ' C/ A QUANTIDADE RESTANTE DEVE SER GRAVADO!!
			If qtde_movto < qtde_aux Then
			  ' REGISTRA O NOVO MOVIMENTO C/ A QUANTIDADE REMANESCENTE
				if Not gera_id_estoque_movto(s_chave, msg_erro) then 
					msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
					exit function
					end if
				
				s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
						" (id_movimento, data, hora, usuario, pedido, fabricante, produto, id_estoque," & _
						" qtde, operacao, estoque, loja, kit, kit_id_estoque, id_ordem_servico) VALUES (" & _
						"'" & s_chave & "'," & _
						bd_formata_data(Date) & "," & _
						"'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
						"'" & id_usuario & "'," & _
						"'" & pedido_aux & "'," & _
						"'" & id_fabricante & "'," & _
						"'" & id_produto & "'," & _
						"'" & id_estoque_aux & "'," & _
						CStr(qtde_aux - qtde_movto) & "," & _
						"'" & operacao_aux & "'," & _
						"'" & estoque_aux & "'," & _
						"'" & loja_aux & "'," & _
						"0, ''," & _
						"'" & id_ordem_servico_aux & "'" & _
						")"
				cn.Execute(s_sql)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				End If
			
		  
		  ' GERA O REGISTRO DE MOVIMENTO ATRIBUINDO A 
		  ' QUANTIDADE TRANSFERIDA P/ O ESTOQUE DE DESTINO
		  ' ==============================================
			if Not gera_id_estoque_movto(s_chave, msg_erro) then 
				msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
				exit function
				end if
				
			s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
					" (id_movimento, data, hora, usuario, pedido, fabricante, produto, id_estoque," & _
					" qtde, operacao, estoque, loja, kit, kit_id_estoque, id_ordem_servico) VALUES (" & _
					"'" & s_chave & "'," & _
					bd_formata_data(Date) & "," & _
					"'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
					"'" & id_usuario & "'," & _
					"'" & id_pedido_destino & "'," & _
					"'" & id_fabricante & "'," & _
					"'" & id_produto & "'," & _
					"'" & id_estoque_aux & "'," & _
					CStr(qtde_movto) & "," & _
					"'" & OP_ESTOQUE_TRANSFERENCIA & "'," & _
					"'" & cod_estoque_destino & "'," & _
					"'" & id_loja_destino & "'," & _
					"0, ''," & _
					"'" & id_ordem_servico_destino & "'" & _
					")"
			cn.Execute(s_sql)
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
		  
		  
		  ' CONTABILIZA QUANTIDADE TRANSFERIDA
			qtde_transferida = qtde_transferida + qtde_movto
															   
		  
		  ' T_ESTOQUE_ITEM: DATA DO ÚLTIMO MOVIMENTO
		  ' ========================================
			s_sql = "SELECT" & _
						" data_ult_movimento" & _
					" FROM t_ESTOQUE_ITEM" & _
					" WHERE" & _
						" (id_estoque = '" & id_estoque_aux & "') AND" & _
						" (fabricante = '" & id_fabricante & "') AND" & _
						" (produto = '" & id_produto & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro = "Falha ao acessar o registro no estoque do produto " & id_produto & " do fabricante " & id_fabricante & " (id_estoque=" & id_estoque_aux & ")"
				exit function
				end if

			rs("data_ult_movimento") = Date
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
		  ' T_ESTOQUE: ATUALIZA DATA DO ÚLTIMO MOVIMENTO
		  ' ============================================
			s_sql = "SELECT" & _
						" data_ult_movimento" & _
					" FROM t_ESTOQUE" & _
					" WHERE" & _
						" (id_estoque = '" & id_estoque_aux & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro = "Falha ao acessar o registro principal no estoque do produto " & id_produto & " do fabricante " & id_fabricante & " (id_estoque=" & id_estoque_aux & ")"
				exit function
				end if
			
			rs("data_ult_movimento") = Date
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			end if
		next
		
	'Log de movimentação do estoque
	if id_ordem_servico_origem <> "" then
		id_ordem_servico_log = id_ordem_servico_origem
	elseif id_ordem_servico_destino <> "" then
		id_ordem_servico_log = id_ordem_servico_destino
	else
		id_ordem_servico_log = ""
		end if
	
	if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, id_fabricante, id_produto, qtde_a_transferir, qtde_transferida, OP_ESTOQUE_LOG_TRANSFERENCIA, cod_estoque_origem, cod_estoque_destino, id_loja_origem, id_loja_destino, id_pedido_origem, id_pedido_destino, "", "", id_ordem_servico_log) then
		msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
		exit function
		end if
		
	estoque_produto_transfere_entre_estoques_v2 = True
end function



' --------------------------------------------------------------------
'   ESTOQUE PROCESSA ENTREGA MERCADORIA
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função processa a entrega das mercadorias ao cliente de
'   modo que os produtos não constem mais do estoque de produtos
'	"vendidos".
'   27/01/2017: revisado p/ estar em conformidade c/ o controle de estoque por empresa.
function estoque_processa_entrega_mercadoria(byval id_usuario, byval id_pedido, byref msg_erro)
dim s_sql
dim s_chave
dim rs
dim iv
dim v_lista()
dim id_nfe_emitente

	estoque_processa_entrega_mercadoria = False

	msg_erro=""
		
    id_usuario = Trim("" & id_usuario)
    id_pedido = Trim("" & id_pedido)
		
	if id_pedido = "" then 
		msg_erro = "Pedido não foi especificado!!"
		exit function
		end if
	
	s_sql = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" & id_pedido & "')"
	set rs=cn.Execute(s_sql)
	if rs.Eof then
		msg_erro = "Falha ao tentar localizar o registro do pedido " & id_pedido & "!!"
		exit function
		end if

	id_nfe_emitente = rs("id_nfe_emitente")

	if rs.State <> 0 then rs.Close
	set rs=nothing

'	OBTÉM OS MOVIMENTOS REFERENTES A ESTE PEDIDO DOS PRODUTOS QUE ESTÃO NO ESTOQUE "VENDIDO"
' 	OBS: LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'		 OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'		 FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT id_movimento" & _
            " FROM t_ESTOQUE_MOVIMENTO" & _
            " WHERE (anulado_status = 0)" & _
            " AND (pedido='" & id_pedido & "')" & _
            " AND (estoque='" & ID_ESTOQUE_VENDIDO & "')" & _
            " ORDER BY id_movimento"

    ReDim v_lista(0)
    v_lista(UBound(v_lista)) = ""

    set rs=cn.Execute(s_sql)
    if Err<>0 then 
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
		
    do while Not rs.Eof
	'	ARMAZENA O ID DOS MOVIMENTOS
		If v_lista(UBound(v_lista)) <> "" Then
          ReDim Preserve v_lista(UBound(v_lista) + 1)
          v_lista(UBound(v_lista)) = ""
          End If
      v_lista(UBound(v_lista)) = Trim("" & rs("id_movimento"))
      rs.MoveNext 
      Loop

	if rs.State <> 0 then rs.Close
	set rs=nothing
	
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

'	PROCESSA A ENTREGA DAS MERCADORIAS AO CLIENTE!!
    For iv = LBound(v_lista) To UBound(v_lista)
        If Trim(v_lista(iv)) <> "" Then
          ' OBTÉM O ID DO NOVO REGISTRO
            if Not gera_id_estoque_movto(s_chave, msg_erro) then 
				msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
				exit function
				end if
          
			s_sql = "SELECT * FROM t_ESTOQUE_MOVIMENTO WHERE" & _
					" (id_movimento='" & v_lista(iv) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err<>0 then 
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.Eof then
				msg_erro="Falha ao tentar acessar o registro de movimento do estoque (id=" & v_lista(iv) & ")"
				exit function
				end if
			
		'	PREPARA INCLUSÃO DO NOVO MOVIMENTO
			s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
					" (id_movimento, data, hora, usuario, operacao, estoque, pedido, id_estoque, fabricante, produto, qtde, loja, kit, kit_id_estoque)" & _
					" VALUES (" & _
					"'" & s_chave & "'" & _
					", " & bd_formata_data(Date) & _
					", '" & retorna_so_digitos(formata_hora(Now)) & "'" & _
					", '" & id_usuario & "'" & _
					", '" & OP_ESTOQUE_ENTREGA & "'" & _
					", '" & ID_ESTOQUE_ENTREGUE & "'" & _
					", '" & id_pedido & "'" & _
					", '" & Trim("" & rs("id_estoque")) & "'" & _
					", '" & Trim("" & rs("fabricante")) & "'" & _
					", '" & Trim("" & rs("produto")) & "'" & _
					", " & rs("qtde") & _
					", '" & Trim("" & rs("loja")) & "'" & _
					", " & rs("kit") & _
					", '" & Trim("" & rs("kit_id_estoque")) & "'" & _
					")"
			
		'	ANULA O MOVIMENTO ATUAL
			rs("anulado_status") = 1
			rs("anulado_data") = Date
			rs("anulado_hora") = retorna_so_digitos(formata_hora(Now))
			rs("anulado_usuario") = id_usuario
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

		'	INCLUI O NOVO MOVIMENTO
			cn.Execute(s_sql)
			if Err <> 0 then 
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			'Log de movimentação do estoque
			if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, rs("fabricante"), rs("produto"), rs("qtde"), rs("qtde"), OP_ESTOQUE_LOG_ENTREGA, ID_ESTOQUE_VENDIDO, ID_ESTOQUE_ENTREGUE, "", "", id_pedido, id_pedido, "", "", "") then
				msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
				exit function
				end if
			end if
		next

	estoque_processa_entrega_mercadoria = True
end function



' --------------------------------------------------------------------
'   ESTOQUE TRANSFERE PRODUTO VENDIDO ENTRE PEDIDOS V2
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função transfere do pedido de origem para o pedido de 
'	destino uma determinada quantidade de produtos presentes no
'	estoque de produtos vendidos.
function estoque_transfere_produto_vendido_entre_pedidos_v2(ByVal id_usuario, _
											ByVal pedido_origem, ByVal pedido_destino, _
											ByVal id_fabricante, ByVal id_produto, _
											ByVal qtde_a_transferir, ByRef msg_erro)
dim rs
dim n
dim iv
dim s
dim s_sql
dim s_chave
dim v_estoque
dim qtde_transferida
dim id_estoque_aux
dim qtde_aux
dim operacao_aux
dim loja_aux
dim id_ordem_servico_aux
dim qtde_movto
dim total_estoque_sem_presenca
dim total_estoque_vendido
dim id_nfe_emitente_pedido_origem, id_nfe_emitente_pedido_destino

	estoque_transfere_produto_vendido_entre_pedidos_v2 = False
    msg_erro = ""
	qtde_transferida = 0

	id_usuario = Trim("" & id_usuario)
	pedido_origem = Trim("" & pedido_origem)
	pedido_destino = Trim("" & pedido_destino)
	id_fabricante = Trim("" & id_fabricante)
	id_produto = Trim("" & id_produto)

'	VERIFICA SE O PEDIDO DE ORIGEM E DE DESTINO ESTÃO VINCULADOS AO ESTOQUE DA MESMA EMPRESA
	id_nfe_emitente_pedido_origem = 0
	id_nfe_emitente_pedido_destino = 0
	s_sql = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" & pedido_origem & "')"
	set rs=cn.Execute(s_sql)
	if Not rs.Eof then
		id_nfe_emitente_pedido_origem = CLng(rs("id_nfe_emitente"))
		end if

	if rs.State <> 0 then rs.Close
	set rs=nothing

	s_sql = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" & pedido_destino & "')"
	set rs=cn.Execute(s_sql)
	if Not rs.Eof then
		id_nfe_emitente_pedido_destino = CLng(rs("id_nfe_emitente"))
		end if

	if rs.State <> 0 then rs.Close
	set rs=nothing

	if id_nfe_emitente_pedido_origem <> id_nfe_emitente_pedido_destino then
		msg_erro="A operação não pode ser realizada porque os pedidos estão associados a estoques de empresas diferentes:" & _
				"<br />Pedido de origem (" & c_pedido_origem & "): " & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_pedido_origem) & _
				"<br />Pedido de destino (" & c_pedido_destino & "): " & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente_pedido_destino)
		exit function
		end if

  ' 1) LEMBRE-SE DE QUE PODE HAVER MAIS DE UM REGISTRO EM T_ESTOQUE_MOVIMENTO 
  '    P/ CADA PRODUTO, POIS PODEM TER SIDO USADOS DIFERENTES LOTES DO ESTOQUE 
  '    P/ ATENDER A UM ÚNICO PEDIDO!!
  ' 2) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
  '    OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
  '    FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
    ReDim v_estoque(0)
    v_estoque(UBound(v_estoque)) = ""
	
    s_sql = "SELECT id_movimento, qtde FROM t_ESTOQUE INNER JOIN t_ESTOQUE_MOVIMENTO ON (t_ESTOQUE.id_estoque = t_ESTOQUE_MOVIMENTO.id_estoque)" & _
			" WHERE (anulado_status = 0)" & _
            " AND (estoque = '" & ID_ESTOQUE_VENDIDO & "')" & _
            " AND (pedido = '" & pedido_origem & "')" & _
            " AND (t_ESTOQUE_MOVIMENTO.fabricante = '" & id_fabricante & "')" & _
            " AND (produto = '" & id_produto & "')" & _
			" ORDER BY t_ESTOQUE.data_entrada, t_ESTOQUE.id_estoque"
	set rs=cn.Execute(s_sql)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	n = 0
	do while Not rs.EOF 
        If v_estoque(UBound(v_estoque)) <> "" Then
            ReDim Preserve v_estoque(UBound(v_estoque) + 1)
            v_estoque(UBound(v_estoque)) = ""
            End If
		v_estoque(UBound(v_estoque)) = Trim("" & rs("id_movimento"))
		n = n + CLng(rs("qtde"))
		rs.MoveNext 
		loop

	if rs.State <> 0 then rs.Close
	set rs=nothing
		
'   VERIFICA SE O PEDIDO DE ORIGEM DISPÕE DA QUANTIDADE NECESSÁRIA
	if n < qtde_a_transferir then
		msg_erro = "Não é possível transferir " & formata_inteiro(qtde_a_transferir) & _
				   " unidades do produto " & id_produto & " (fabricante " & id_fabricante & _
				   ") porque o pedido de origem " & pedido_origem & " dispõe de apenas " & _
				   formata_inteiro(n) & " unidades!!"
		exit function
		end if

'   VERIFICA SE O PEDIDO DE DESTINO NECESSITA DA QUANTIDADE ESPECIFICADA
	s_sql = "SELECT Sum(qtde) AS total FROM t_ESTOQUE_MOVIMENTO WHERE (anulado_status=0)" & _
			" AND (pedido='" & pedido_destino & "')" & _
			" AND (fabricante='" & id_fabricante & "')" & _
			" AND (produto='" & id_produto & "')" & _
			" AND (estoque='" & ID_ESTOQUE_SEM_PRESENCA & "')"
	set rs=cn.Execute(s_sql)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	n = 0
	if Not rs.Eof then
		if Not IsNull(rs("total")) then n = CLng(rs("total"))
		end if
		
	if n < qtde_a_transferir then
		msg_erro = "Não é possível transferir " & formata_inteiro(qtde_a_transferir) & _
				   " unidades do produto " & id_produto & " (fabricante " & id_fabricante & _
				   ") porque o pedido de destino " & pedido_destino & " aguarda apenas " & _
				   formata_inteiro(n) & " unidades!!"
		exit function
		end if
		
	if rs.State <> 0 then rs.Close
	set rs=nothing
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function
			
	for iv=LBound(v_estoque) To UBound(v_estoque)
		If Trim(v_estoque(iv)) <> "" Then
          
          ' JÁ TRANSFERIU TUDO?
			If qtde_transferida >= qtde_a_transferir Then Exit For
			
		  ' PEDIDO ORIGEM: REGISTRO DE MOVIMENTO DO ESTOQUE VENDIDO
		  ' =======================================================
            s_sql = "SELECT *" & _
					" FROM t_ESTOQUE_MOVIMENTO" & _
					" WHERE (anulado_status = 0)" & _
                    " AND (id_movimento = '" & Trim(v_estoque(iv)) & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro="Falha ao acessar o registro de movimento no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				exit function
				end if

			id_estoque_aux = Trim("" & rs("id_estoque"))
			qtde_aux = CLng(rs("qtde"))
			operacao_aux = Trim("" & rs("operacao"))
			loja_aux = Trim("" & rs("loja"))
			id_ordem_servico_aux = Trim("" & rs("id_ordem_servico"))
			
			qtde_movto = qtde_aux
			
          ' A QUANTIDADE QUE FALTA SER TRANSFERIDA É MENOR QUE A QUANTIDADE DO MOVIMENTO
            If (qtde_a_transferir - qtde_transferida) < qtde_aux Then
                qtde_movto = qtde_a_transferir - qtde_transferida
                End If

          ' ANULA O MOVIMENTO
			rs("anulado_status") = 1
			rs("anulado_data") = Date
			rs("anulado_hora") = retorna_so_digitos(formata_hora(Now))
			rs("anulado_usuario") = id_usuario
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

          ' TRANSFERÊNCIA PARCIAL: O MOVIMENTO ORIGINAL FOI ANULADO E UM NOVO MOVIMENTO 
          ' C/ A QUANTIDADE RESTANTE DEVE SER GRAVADO!!
            If qtde_movto < qtde_aux Then
              ' REGISTRA O MOVIMENTO DE SAÍDA NO ESTOQUE
				if Not gera_id_estoque_movto(s_chave, msg_erro) then 
					msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
					exit function
					end if
				
                s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
                        " (id_movimento, data, hora, usuario, pedido, fabricante, produto, id_estoque," & _
                        " qtde, operacao, estoque, loja, kit, kit_id_estoque, id_ordem_servico) VALUES (" & _
                        "'" & s_chave & "'," & _
                        bd_formata_data(Date) & "," & _
                        "'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
                        "'" & id_usuario & "'," & _
                        "'" & pedido_origem & "'," & _
                        "'" & id_fabricante & "'," & _
                        "'" & id_produto & "'," & _
                        "'" & id_estoque_aux & "'," & _
                        CStr(qtde_aux - qtde_movto) & "," & _
                        "'" & operacao_aux & "'," & _
                        "'" & ID_ESTOQUE_VENDIDO & "'," & _
                        "'" & loja_aux & "'," & _
                        "0, ''," & _
                        "'" & id_ordem_servico_aux & "'" & _
                        ")"
				cn.Execute(s_sql)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
                End If


		  ' PEDIDO DESTINO: REGISTRO DE MOVIMENTO DO ESTOQUE VENDIDO
		  ' ========================================================
		  ' GERA O MOVIMENTO DE SAÍDA DO ESTOQUE PARA O PEDIDO DE DESTINO (QUE RECEBEU OS PRODUTOS)
			if Not gera_id_estoque_movto(s_chave, msg_erro) then 
				msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
				exit function
				end if
				
            s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
                    " (id_movimento, data, hora, usuario, pedido, fabricante, produto, id_estoque," & _
                    " qtde, operacao, estoque, loja, kit, kit_id_estoque, id_ordem_servico) VALUES (" & _
                    "'" & s_chave & "'," & _
                    bd_formata_data(Date) & "," & _
                    "'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
                    "'" & id_usuario & "'," & _
                    "'" & pedido_destino & "'," & _
                    "'" & id_fabricante & "'," & _
                    "'" & id_produto & "'," & _
                    "'" & id_estoque_aux & "'," & _
                    CStr(qtde_movto) & "," & _
                    "'" & operacao_aux & "'," & _
                    "'" & ID_ESTOQUE_VENDIDO & "'," & _
                    "'" & loja_aux & "'," & _
                    "0, ''," & _
                    "'" & id_ordem_servico_aux & "'" & _
                    ")"
			cn.Execute(s_sql)
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if


		 '  TOTALIZAÇÃO
			qtde_transferida = qtde_transferida + qtde_movto

		
          ' T_ESTOQUE: ATUALIZA DATA DO ÚLTIMO MOVIMENTO
          ' ============================================
            s_sql = "SELECT data_ult_movimento FROM t_ESTOQUE WHERE" & _
                    " (id_estoque = '" & id_estoque_aux & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			
			if rs.EOF then
				msg_erro = "Falha ao acessar o registro principal no estoque do produto " & id_produto & " do fabricante " & id_fabricante
				exit function
				end if
			
			rs("data_ult_movimento") = Date
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			end if
		next


'	TRANSFERIU TUDO?
'   =================
	If qtde_transferida < qtde_a_transferir Then 
		msg_erro = "Não foi possível concluir a transferência porque o pedido de origem não dispõe de quantidade suficiente!!"
		exit function
		end if


'	PEDIDO ORIGEM: LISTA DE PRODUTOS SEM PRESENÇA NO ESTOQUE
'   ========================================================
'	OBTÉM A QUANTIDADE DE PRODUTOS EM ESPERA
	s_sql = "SELECT Sum(qtde) AS total FROM t_ESTOQUE_MOVIMENTO" & _
			" WHERE (anulado_status = 0)" & _
			" AND (estoque='" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
			" AND (pedido='" & pedido_origem & "')" & _
			" AND (fabricante='" & id_fabricante & "')" & _
			" AND (produto='" & id_produto & "')"

	if rs.State <> 0 then rs.Close
	rs.Open s_sql, cn
	qtde_movto = 0
	if Not rs.EOF then
		if Not IsNull(rs("total")) then qtde_movto = CLng(rs("total"))
		end if

	qtde_movto = qtde_movto + qtde_transferida
	
'	ANULA O REGISTRO ANTERIOR DESTE PEDIDO DA LISTA DE ESPERA DOS "PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE"
	s_sql = "UPDATE t_ESTOQUE_MOVIMENTO SET" & _
			" anulado_status=1" & _
			", anulado_data=" & bd_formata_data(Date) & _
			", anulado_hora='" & retorna_so_digitos(formata_hora(Now)) & "'" & _
			", anulado_usuario='" & id_usuario & "'" & _
			" WHERE (anulado_status = 0)" & _
			" AND (estoque='" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
			" AND (pedido='" & pedido_origem & "')" & _
			" AND (fabricante='" & id_fabricante & "')" & _
			" AND (produto='" & id_produto & "')"
	cn.Execute(s_sql)
	if Err<>0 then 
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if

'	CRIA NOVO REGISTRO
	if Not gera_id_estoque_movto(s_chave, msg_erro) then 
		msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
		exit function
		end if
	s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
	        " (id_movimento, data, hora, usuario, id_estoque, fabricante, produto," & _
	        " qtde, operacao, estoque, pedido, loja, kit, kit_id_estoque) VALUES (" & _
	        "'" & s_chave & "'," & _
	        bd_formata_data(Date) & "," & _
	        "'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
	        "'" & id_usuario & "'," & _
	        "''," & _
	        "'" & id_fabricante & "'," & _
	        "'" & id_produto & "'," & _
	        CStr(qtde_movto) & "," & _
	        "'" & OP_ESTOQUE_VENDA & "'," & _
	        "'" & ID_ESTOQUE_SEM_PRESENCA & "'," & _
	        "'" & pedido_origem & "'," & _
	        "''," & _
	        "0, '')"
	cn.Execute(s_sql)
	if Err<>0 then 
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
			
			
'	PEDIDO DESTINO: LISTA DE PRODUTOS SEM PRESENÇA NO ESTOQUE
'   =========================================================
'	DECREMENTA A QUANTIDADE DE PRODUTOS EM ESPERA NA LISTA DE PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE
	s_sql = "SELECT Sum(qtde) AS total FROM t_ESTOQUE_MOVIMENTO" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque='" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
			" AND (pedido='" & pedido_destino & "')" & _
			" AND (fabricante='" & id_fabricante & "')" & _
			" AND (produto='" & id_produto & "')" 
	if rs.State <> 0 then rs.Close
	rs.Open s_sql, cn
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
			
	qtde_movto = 0
	if Not rs.EOF then
		if Not IsNull(rs("total")) then qtde_movto = CLng(rs("total"))
		end if

	if qtde_movto < qtde_transferida then
		msg_erro = "Não foi possível concluir a transferência porque o pedido de destino não está aguardando a quantidade especificada de produtos!!"
		exit function
		end if
		
	qtde_movto = qtde_movto - qtde_transferida

'	ANULA O REGISTRO ANTERIOR DESTE PEDIDO DA LISTA DE ESPERA DOS "PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE"
	s_sql = "UPDATE t_ESTOQUE_MOVIMENTO SET" & _
			" anulado_status=1" & _
			", anulado_data=" & bd_formata_data(Date) & _
			", anulado_hora='" & retorna_so_digitos(formata_hora(Now)) & "'" & _
			", anulado_usuario='" & id_usuario & "'" & _
			" WHERE (anulado_status = 0)" & _
			" AND (estoque='" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
			" AND (pedido='" & pedido_destino & "')" & _
			" AND (fabricante='" & id_fabricante & "')" & _
			" AND (produto='" & id_produto & "')"
	cn.Execute(s_sql)
	if Err<>0 then 
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	if qtde_movto > 0 then
	'	CRIA NOVO REGISTRO
		if Not gera_id_estoque_movto(s_chave, msg_erro) then 
			msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
			exit function
			end if
		s_sql = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
		        " (id_movimento, data, hora, usuario, id_estoque, fabricante, produto," & _
		        " qtde, operacao, estoque, pedido, loja, kit, kit_id_estoque) VALUES (" & _
		        "'" & s_chave & "'," & _
		        bd_formata_data(Date) & "," & _
		        "'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
		        "'" & id_usuario & "'," & _
		        "''," & _
		        "'" & id_fabricante & "'," & _
		        "'" & id_produto & "'," & _
		        CStr(qtde_movto) & "," & _
		        "'" & OP_ESTOQUE_VENDA & "'," & _
		        "'" & ID_ESTOQUE_SEM_PRESENCA & "'," & _
		        "'" & pedido_destino & "'," & _
		        "''," & _
		        "0, '')"
		cn.Execute(s_sql)
		if Err<>0 then 
			msg_erro=Cstr(Err) & ": " & Err.Description
			exit function
			end if
		end if
		

'	PEDIDO ORIGEM: ATUALIZA STATUS DE ENTREGA
'   =========================================
	total_estoque_sem_presenca = 0
	s_sql = "SELECT Sum(qtde) AS total FROM t_ESTOQUE_MOVIMENTO" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque = '" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
			" AND (pedido = '" & pedido_origem & "')"
	if rs.State <> 0 then rs.Close
	rs.open s_sql, cn
	if Not rs.Eof then
		if IsNumeric(rs("total")) then total_estoque_sem_presenca = CLng(rs("total"))
		end if

	total_estoque_vendido = 0
	s_sql = "SELECT Sum(qtde) AS total FROM t_ESTOQUE_MOVIMENTO" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque = '" & ID_ESTOQUE_VENDIDO & "')" & _
			" AND (pedido = '" & pedido_origem & "')"
	if rs.State <> 0 then rs.Close
	rs.open s_sql, cn
	if Not rs.Eof then
		if IsNumeric(rs("total")) then total_estoque_vendido = CLng(rs("total"))
		end if

	s_sql = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & pedido_origem & "')"
	if rs.State <> 0 then rs.Close
	rs.open s_sql, cn
	if rs.Eof then
		msg_erro = "Pedido " & pedido_origem & " não foi encontrado."
		exit function
	else
	'	STATUS DE ENTREGA
		if total_estoque_vendido = 0 then
			s = ST_ENTREGA_ESPERAR
		elseif total_estoque_sem_presenca = 0 then
			s = ST_ENTREGA_SEPARAR
		else
			s = ST_ENTREGA_SPLIT_POSSIVEL
			end if
				
		if Trim("" & rs("st_entrega")) <> s then
			rs("st_entrega") = s
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			end if
		end if


'	PEDIDO DESTINO: ATUALIZA STATUS DE ENTREGA
'   ==========================================
	total_estoque_sem_presenca = 0
	s_sql = "SELECT Sum(qtde) AS total FROM t_ESTOQUE_MOVIMENTO" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque = '" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
			" AND (pedido = '" & pedido_destino & "')"
	if rs.State <> 0 then rs.Close
	rs.open s_sql, cn
	if Not rs.Eof then
		if IsNumeric(rs("total")) then total_estoque_sem_presenca = CLng(rs("total"))
		end if

	total_estoque_vendido = 0
	s_sql = "SELECT Sum(qtde) AS total FROM t_ESTOQUE_MOVIMENTO" & _
			" WHERE (anulado_status=0)" & _
			" AND (estoque = '" & ID_ESTOQUE_VENDIDO & "')" & _
			" AND (pedido = '" & pedido_destino & "')"
	if rs.State <> 0 then rs.Close
	rs.open s_sql, cn
	if Not rs.Eof then
		if IsNumeric(rs("total")) then total_estoque_vendido = CLng(rs("total"))
		end if

	s_sql = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & pedido_destino & "')"
	if rs.State <> 0 then rs.Close
	rs.open s_sql, cn
	if rs.Eof then
		msg_erro = "Pedido " & pedido_destino & " não foi encontrado."
		exit function
	else
	'	STATUS DE ENTREGA
		if total_estoque_vendido = 0 then
			s = ST_ENTREGA_ESPERAR
		elseif total_estoque_sem_presenca = 0 then
			s = ST_ENTREGA_SEPARAR
		else
			s = ST_ENTREGA_SPLIT_POSSIVEL
			end if
				
		if Trim("" & rs("st_entrega")) <> s then
			rs("st_entrega") = s
			rs.Update
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			end if
		end if

	'Log de movimentação do estoque
	if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente_pedido_origem, id_fabricante, id_produto, qtde_a_transferir, qtde_transferida, OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS, ID_ESTOQUE_VENDIDO, ID_ESTOQUE_VENDIDO, "", "", pedido_origem, pedido_destino, "", "", "") then
		msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
		exit function
		end if
						
	estoque_transfere_produto_vendido_entre_pedidos_v2 = True
end function



' --------------------------------------------------------------------
'   ESTOQUE VERIFICA MERCADORIAS PARA DEVOLUCAO
'   Retorno da função:
'      False - Ocorreu falha ao tentar consultar o banco de dados.
'      True - Conseguiu consultar o banco de dados.
'   Esta função consulta o banco de dados para contabilizar a
'   quantidade de produtos que já foram entregues ao cliente e a
'	quantidade de produtos que o cliente já devolveu.
'   Note que os itens de pedido a serem analisados são passados
'   pelo vetor do parâmetro v_item.
'   27/01/2017: revisado p/ estar em conformidade c/ o controle de estoque por empresa.
'   14/03/2018: revisado para incluir na consulta os itens que foram solicitados devolução mas que ainda estão em processo de análise para aprovação.
function estoque_verifica_mercadorias_para_devolucao(byref v_item, byref msg_erro)
dim s
dim s_sql
dim i
dim rs
	estoque_verifica_mercadorias_para_devolucao = False
	msg_erro = ""
	
	for i=Lbound(v_item) to Ubound(v_item)
		with v_item(i)
			.qtde = 0
			.qtde_devolvida_anteriormente = 0
            .qtde_devolucao_pendente = 0
			
		'  LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
		'  OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
		'  FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
			s_sql = "SELECT Sum(qtde) AS total FROM t_ESTOQUE_MOVIMENTO WHERE (anulado_status=0)" & _
					" AND (pedido='" & .pedido & "')" & _
					" AND (fabricante='" & .fabricante & "')" & _
					" AND (produto='" & .produto & "')" & _
					" AND (estoque='" & ID_ESTOQUE_ENTREGUE & "')"
			set rs=cn.execute(s_sql)
			if Err<>0 then 
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			if Not rs.EOF then if IsNumeric(rs("total")) then .qtde = CLng(rs("total"))
			if rs.State <> 0 then rs.Close
		
			s_sql = "SELECT Sum(qtde) AS total FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
					" WHERE (devolucao_status<>0)" & _
					" AND (devolucao_pedido='" & .pedido & "')" & _
					" AND (t_ESTOQUE_ITEM.fabricante='" & .fabricante & "')" & _
					" AND (produto='" & .produto & "')"
			set rs=cn.execute(s_sql)
			if Err<>0 then 
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			if Not rs.EOF then if IsNumeric(rs("total")) then .qtde_devolvida_anteriormente = CLng(rs("total"))
			if rs.State <> 0 then rs.Close
		
			s_sql = "SELECT Sum(qtde) AS total FROM t_PEDIDO_DEVOLUCAO_ITEM INNER JOIN t_PEDIDO_DEVOLUCAO ON (t_PEDIDO_DEVOLUCAO_ITEM.id_pedido_devolucao=t_PEDIDO_DEVOLUCAO.id)" & _
					" WHERE (status=" & COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA & ")" & _
					" AND (pedido='" & .pedido & "')" & _
					" AND (fabricante='" & .fabricante & "')" & _
					" AND (produto='" & .produto & "')"
			set rs=cn.execute(s_sql)
			if Err<>0 then 
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if
			if Not rs.EOF then if IsNumeric(rs("total")) then .qtde_devolucao_pendente = CLng(rs("total"))
			if rs.State <> 0 then rs.Close
			end with
		next

	estoque_verifica_mercadorias_para_devolucao = True
end function



' --------------------------------------------------------------------
'   ESTOQUE PROCESSA DEVOLUCAO MERCADORIAS V2
'   Retorno da função:
'      False - Ocorreu falha ao tentar movimentar o estoque.
'      True - Conseguiu fazer a movimentação do estoque.
'   IMPORTANTE: sempre chame esta rotina dentro de uma transação para 
'      garantir a consistência dos registros.
'   Esta função processa a devolução de mercadorias pelo cliente,
'   fazendo a entrada no estoque da quantidade de produtos devolvida
'	através de "registros de entrada no estoque por devolução".
function estoque_processa_devolucao_mercadorias_v2(byval id_usuario, byval id_pedido, _
												byval id_fabricante, byval id_produto, _
												byval id_item_devolvido, _
												byval qtde_a_devolver, byref msg_erro)
dim s
dim rs
dim v_devolvido
dim v_estoque
dim i
dim iv
dim qtde_aux
dim qtde_movto
dim qtde_devolvida
dim preco_fabricante
dim vl_custo2
dim vl_BC_ICMS_ST
dim vl_ICMS_ST
dim ncm
dim cst
dim id_estoque
dim id_movimento
dim id_loja
dim id_nfe_emitente

	estoque_processa_devolucao_mercadorias_v2 = False
	msg_erro = ""
	id_nfe_emitente = 0
	
	id_usuario = Trim("" & id_usuario)
	id_pedido = Trim("" & id_pedido)
	id_fabricante = Trim("" & id_fabricante)
	id_produto = Trim("" & id_produto)
	id_item_devolvido = Trim("" & id_item_devolvido)

	if Not IsNumeric(qtde_a_devolver) then exit function
	qtde_a_devolver = CLng(qtde_a_devolver)
	
'	OBTÉM NÚMERO DA LOJA DO PEDIDO
	s = "SELECT pedido, loja FROM t_PEDIDO WHERE (pedido='" & id_pedido & "')"
	set rs=cn.execute(s)
	if rs.Eof then
		msg_erro = "Pedido " & id_pedido & " não foi encontrado."
		exit function
		end if
	
	id_loja = Trim("" & rs("loja"))
	
	if rs.State <> 0 then rs.Close
	set rs=nothing
	
  ' 1) LEMBRE-SE DE QUE PODE HAVER MAIS DE UM REGISTRO EM T_ESTOQUE_MOVIMENTO 
  '    P/ CADA PRODUTO, POIS PODEM TER SIDO USADOS DIFERENTES LOTES DO ESTOQUE 
  '    P/ ATENDER A UM ÚNICO PEDIDO!!
  ' 2) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
  '    OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
  '    FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
  ' 3) LEMBRE-SE DE QUE JÁ PODEM HAVER PRODUTOS DESTE PEDIDO DEVOLVIDOS 
  '	   ANTERIORMENTE.
    
'	OBTÉM PRODUTOS DEVOLVIDOS ANTERIORMENTE
	ReDim v_devolvido(0)
	set v_devolvido(UBound(v_devolvido)) = New cl_DUAS_COLUNAS
	v_devolvido(UBound(v_devolvido)).c1 = ""
	
	s = "SELECT devolucao_id_estoque, Sum(qtde) AS total" & _
		" FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque" & _
		" WHERE (devolucao_status<>0)" & _
		" AND (devolucao_pedido='" & id_pedido & "')" & _
		" AND (t_ESTOQUE_ITEM.fabricante='" & id_fabricante & "')" & _
		" AND (produto='" & id_produto & "')" & _
		" GROUP BY devolucao_id_estoque" & _
		" ORDER BY devolucao_id_estoque"
	set rs=cn.execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	do while Not rs.EOF 
		If v_devolvido(UBound(v_devolvido)).c1 <> "" Then
			ReDim Preserve v_devolvido(UBound(v_devolvido) + 1)
			set v_devolvido(UBound(v_devolvido)) = New cl_DUAS_COLUNAS
			v_devolvido(UBound(v_devolvido)).c1 = ""
			End If
		with v_devolvido(UBound(v_devolvido))
			.c1 = Trim("" & rs("devolucao_id_estoque"))
			if IsNumeric(rs("total")) then .c2 = CLng(rs("total")) else .c2 = 0
			end with
		rs.MoveNext
		loop
		
	if rs.State <> 0 then rs.Close
	set rs=nothing

'	OBTÉM PRODUTOS ENTREGUES AO CLIENTE
	ReDim v_estoque(0)
	set v_estoque(UBound(v_estoque)) = New cl_DUAS_COLUNAS
	v_estoque(UBound(v_estoque)).c1 = ""

	s = "SELECT id_estoque, Sum(qtde) AS total" & _
		" FROM t_ESTOQUE_MOVIMENTO" & _
		" WHERE (anulado_status = 0)" & _
		" AND (estoque = '" & ID_ESTOQUE_ENTREGUE & "')" & _
		" AND (pedido = '" & id_pedido & "')" & _
		" AND (fabricante = '" & id_fabricante & "')" & _
		" AND (produto = '" & id_produto & "')" & _
		" GROUP BY id_estoque" & _
		" ORDER BY id_estoque DESC"
	set rs=cn.execute(s)
	if Err <> 0 then
		msg_erro=Cstr(Err) & ": " & Err.Description
		exit function
		end if
	
	do while Not rs.EOF 
		If v_estoque(UBound(v_estoque)).c1 <> "" Then
			ReDim Preserve v_estoque(UBound(v_estoque) + 1)
			set v_estoque(UBound(v_estoque)) = New cl_DUAS_COLUNAS
			v_estoque(UBound(v_estoque)).c1 = ""
			End If
		with v_estoque(UBound(v_estoque))
			.c1 = Trim("" & rs("id_estoque"))
			if IsNumeric(rs("total")) then .c2 = CLng(rs("total")) else .c2 = 0
			end with
		rs.MoveNext
		loop
	
	if rs.State <> 0 then rs.Close
	set rs=nothing
	
	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

	qtde_devolvida = 0
	for iv=LBound(v_estoque) To UBound(v_estoque)
		If Trim(v_estoque(iv).c1) <> "" Then
		'	JÁ MOVIMENTOU TUDO?
			If qtde_devolvida >= qtde_a_devolver Then Exit For
			
			qtde_aux = v_estoque(iv).c2
			
		'	VERIFICA SE JÁ HOUVE DEVOLUÇÕES ANTERIORES
		'	IMPORTANTE: OS REGISTROS DE MOVIMENTO REFERENTES AO ESTOQUE 'ETG' (ID_ESTOQUE_ENTREGUE) NÃO SÃO ANULADOS, POIS O PEDIDO NÃO ALTERA O PASSADO, A DEVOLUÇÃO IRÁ CONSTAR COMO UM VALOR NEGATIVO A PARTIR DA DATA EM QUE OCORREU
			for i = Lbound(v_devolvido) to Ubound(v_devolvido)
				if v_devolvido(i).c1 = v_estoque(iv).c1 then
					qtde_aux = qtde_aux - v_devolvido(i).c2
					exit for
					end if
				next
			
		'	AINDA HÁ SALDO DE PRODUTOS QUE PODEM SER DEVOLVIDOS
			if qtde_aux > 0 then
				qtde_movto=qtde_aux
			' 	A QUANTIDADE QUE FALTA SER DEVOLVIDA É MENOR QUE A QUANTIDADE DO MOVIMENTO
				If (qtde_a_devolver - qtde_devolvida) < qtde_movto Then
					qtde_movto = qtde_a_devolver - qtde_devolvida
					End If
				
			'	OBTÉM A EMPRESA PROPRIETÁRIA DO ESTOQUE
				s = "SELECT id_nfe_emitente FROM t_ESTOQUE WHERE (id_estoque = '" & v_estoque(iv).c1 & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if

				if rs.EOF then
					msg_erro = "Falha ao acessar o registro de estoque (id_estoque = '" & v_estoque(iv).c1 & "') do produto " & id_produto & " do fabricante " & id_fabricante
					exit function
					end if

				id_nfe_emitente = CLng(rs("id_nfe_emitente"))

			'	OBTÉM DADOS DO PRODUTO P/ REGISTRAR A NOVA ENTRADA C/ OS MESMOS VALORES
				s = "SELECT * FROM t_ESTOQUE_ITEM" & _
					" WHERE (id_estoque='" & v_estoque(iv).c1 & "')" & _
					" AND (fabricante='" & id_fabricante & "')" & _
					" AND (produto='" & id_produto & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
				
				if rs.EOF then
					msg_erro = "Falha ao acessar o registro principal no estoque do produto " & id_produto & " do fabricante " & id_fabricante
					exit function
					end if
				
				preco_fabricante = rs("preco_fabricante")
				vl_custo2 = rs("vl_custo2")
				vl_BC_ICMS_ST = rs("vl_BC_ICMS_ST")
				vl_ICMS_ST = rs("vl_ICMS_ST")
				ncm = Trim("" & rs("ncm"))
				cst = Trim("" & rs("cst"))
			
				If Not gera_id_estoque(id_estoque, msg_erro) Then Exit Function
			
			' 	GRAVA INFORMAÇÕES BÁSICAS DA ENTRADA NO ESTOQUE
				s = "INSERT INTO t_ESTOQUE" & _
					" (id_estoque, data_entrada, hora_entrada, fabricante, documento," & _
					" usuario, data_ult_movimento, kit, entrada_especial," & _
					" devolucao_status, devolucao_data, devolucao_hora, devolucao_usuario," & _
					" devolucao_loja, devolucao_pedido, devolucao_id_item_devolvido, devolucao_id_estoque, obs," & _
					" id_nfe_emitente" & _
					") VALUES (" & _
					"'" & id_estoque & "'" & _
					"," & bd_formata_data(Date) & _
					",'" & retorna_so_digitos(formata_hora(Now)) & "'" & _
					",'" & id_fabricante & "'" & _
					",'DEVOLUÇÃO: (" & id_loja & ") " & id_pedido & "'" & _
					",'" & id_usuario & "'" & _
					"," & bd_formata_data(Date) & _
					", 0" & _
					", 0" & _
					", 1, " & bd_formata_data(Date) & _
					", '" & retorna_so_digitos(formata_hora(Now)) & "'" & _
					", '" & id_usuario & "'" & _ 
					", '" & id_loja & "'" & _
					", '" & id_pedido & "'" & _
					", '" & id_item_devolvido & "'" & _
					", '" & v_estoque(iv).c1 & "'" & _
					", ''" & _
					", " & Cstr(id_nfe_emitente) & _
					")"
				cn.Execute(s)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if

				s = "INSERT INTO T_ESTOQUE_ITEM" & _
					" (id_estoque, fabricante, produto, qtde, preco_fabricante, vl_custo2, qtde_utilizada," & _
					" vl_BC_ICMS_ST, vl_ICMS_ST," & _
					" ncm, cst," & _
					" data_ult_movimento, sequencia)" & _
					" VALUES (" & _
					"'" & id_estoque & "'" & _
					",'" & id_fabricante & "'" & _
					",'" & id_produto & "'" & _
					"," & CStr(qtde_movto) & _
					"," & bd_formata_numero(preco_fabricante) & _
					"," & bd_formata_numero(vl_custo2) & _
					"," & CStr(qtde_movto) & _
					"," & bd_formata_numero(vl_BC_ICMS_ST) & _
					"," & bd_formata_numero(vl_ICMS_ST) & _
					",'" & ncm & "'" & _
					",'" & cst & "'" & _
					"," & bd_formata_data(Date) & _
					", 1" & _
					")"
				cn.Execute(s)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
			
			' 	COLOCA NO ESTOQUE DE DEVOLUÇÃO
				if Not gera_id_estoque_movto(id_movimento, msg_erro) then 
					msg_erro="Falha ao tentar gerar um número identificador para o registro de movimento no estoque. " & msg_erro
					exit function
					end if
			
				s = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
					" (id_movimento, data, hora, usuario, pedido, fabricante, produto, id_estoque," & _
					" qtde, operacao, estoque, loja, kit, kit_id_estoque) VALUES (" & _
					"'" & id_movimento & "'," & _
					bd_formata_data(Date) & "," & _
					"'" & retorna_so_digitos(formata_hora(Now)) & "'," & _
					"'" & id_usuario & "'," & _
					"'" & id_pedido & "'," & _
					"'" & id_fabricante & "'," & _
					"'" & id_produto & "'," & _
					"'" & id_estoque & "'," & _
					CStr(qtde_movto) & "," & _
					"'" & OP_ESTOQUE_DEVOLUCAO & "'," & _
					"'" & ID_ESTOQUE_DEVOLUCAO & "'," & _
					"'" & id_loja & "', 0, '')"
				cn.Execute(s)
				if Err <> 0 then
					msg_erro=Cstr(Err) & ": " & Err.Description
					exit function
					end if
					
			' 	CONTABILIZA QUANTIDADE DEVOLVIDA
				qtde_devolvida = qtde_devolvida + qtde_movto
				end if
			end if
		next

'	CONSEGUIU DEVOLVER TUDO?
	if qtde_devolvida < qtde_a_devolver then 
		msg_erro="Produto " & id_produto & " do fabricante " & id_fabricante & ": " & Cstr(qtde_a_devolver - qtde_devolvida) & " unidades não foram devolvidas."
		exit function
		end if
		
	'Log de movimentação do estoque
	if Not grava_log_estoque_v2(id_usuario, id_nfe_emitente, id_fabricante, id_produto, qtde_a_devolver, qtde_devolvida, OP_ESTOQUE_LOG_DEVOLUCAO, ID_ESTOQUE_ENTREGUE, ID_ESTOQUE_DEVOLUCAO, id_loja, id_loja, id_pedido, id_pedido, "", "", "") then
		msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
		exit function
		end if
		
	estoque_processa_devolucao_mercadorias_v2 = True
end function

%>
