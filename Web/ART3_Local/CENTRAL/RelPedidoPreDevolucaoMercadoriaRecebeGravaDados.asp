<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/Global.asp" -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================================================
'	  RelPedidoPreDevolucaoMercadoriaRecebeGravaDados.asp
'     ===============================================================================
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

	class cl_REL_PEDIDO_DEVOL_GRAVA_ITENS_RECEBIDOS
		dim id_devolucao
		dim pedido
        dim fabricante
        dim produto
		dim qtde
		dim qtde_estoque_venda
		dim qtde_estoque_danificados
		end class
		
	dim s, msg_erro
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim alerta
	alerta=""
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_PRE_DEVOLUCAO_RECEBIMENTO_MERCADORIA, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim c_qtde_pre_devolucao_itens, intQtdePreDevolucaoItens, vPreDevolucaoItens
	c_qtde_pre_devolucao_itens=Trim(Request("c_qtde_pre_devolucoes_itens"))
	intQtdePreDevolucaoItens=CInt(c_qtde_pre_devolucao_itens)
	
	redim vPreDevolucaoItens(0)
	set vPreDevolucaoItens(Ubound(vPreDevolucaoItens)) = new cl_REL_PEDIDO_DEVOL_GRAVA_ITENS_RECEBIDOS
	vPreDevolucaoItens(Ubound(vPreDevolucaoItens)).id_devolucao = ""
	
	dim i, j
	dim c_id_devolucao, c_pedido, c_fabricante, c_produto, c_qtde, c_qtde_estoque_venda, c_qtde_estoque_danificados
    dim qtde_itens, blnHaItens, s_id_item_devolvido, intNsuNovoBlocoNotas, devolucao_a
    blnHaItens = False
    devolucao_a = "XXX"
	for i = 1 to intQtdePreDevolucaoItens
		c_id_devolucao = Trim(Request.Form("c_devolucao_id_" & Cstr(i)))
		c_pedido = Trim(Request.Form("c_pedido_" & Cstr(i)))
		c_fabricante = Trim(Request.Form("c_fabricante_" & Cstr(i)))
		c_produto = Trim(Request.Form("c_produto_" & Cstr(i)))
		c_qtde = Trim(Request.Form("c_qtde_" & Cstr(i)))
		c_qtde_estoque_venda = Trim(Request.Form("c_qtde_estoque_venda_" & Cstr(i)))
		c_qtde_estoque_danificados = Trim(Request.Form("c_qtde_estoque_danificado_" & Cstr(i)))
		if (c_id_devolucao<>"") And ( (c_qtde_estoque_venda<>"") Or (c_qtde_estoque_danificados<>"") ) then
			if vPreDevolucaoItens(Ubound(vPreDevolucaoItens)).id_devolucao <> "" then
				redim preserve vPreDevolucaoItens(Ubound(vPreDevolucaoItens)+1)
				set vPreDevolucaoItens(Ubound(vPreDevolucaoItens)) = new cl_REL_PEDIDO_DEVOL_GRAVA_ITENS_RECEBIDOS
				end if
			vPreDevolucaoItens(Ubound(vPreDevolucaoItens)).id_devolucao = c_id_devolucao
			vPreDevolucaoItens(Ubound(vPreDevolucaoItens)).pedido = c_pedido
			vPreDevolucaoItens(Ubound(vPreDevolucaoItens)).fabricante = c_fabricante
			vPreDevolucaoItens(Ubound(vPreDevolucaoItens)).produto = c_produto
			vPreDevolucaoItens(Ubound(vPreDevolucaoItens)).qtde = CInt(c_qtde)
            if c_qtde_estoque_venda <> "" then
			    vPreDevolucaoItens(Ubound(vPreDevolucaoItens)).qtde_estoque_venda = converte_numero(c_qtde_estoque_venda)
            else
			    vPreDevolucaoItens(Ubound(vPreDevolucaoItens)).qtde_estoque_venda = 0
                end if
            if c_qtde_estoque_danificados <> "" then
			    vPreDevolucaoItens(Ubound(vPreDevolucaoItens)).qtde_estoque_danificados = converte_numero(c_qtde_estoque_danificados)
            else
			    vPreDevolucaoItens(Ubound(vPreDevolucaoItens)).qtde_estoque_danificados = 0
                end if
            blnHaItens = True
			end if
		next

    if Not blnHaItens then
        alerta=texto_add_br(alerta)
		alerta=alerta & "Não foi informada nenhuma unidade para o estoque."
        end if

	for i=Lbound(vPreDevolucaoItens) to Ubound(vPreDevolucaoItens)
		if Trim(vPreDevolucaoItens(i).id_devolucao)<>"" then
			if (vPreDevolucaoItens(i).qtde_estoque_venda+vPreDevolucaoItens(i).qtde_estoque_danificados) > vPreDevolucaoItens(i).qtde then
                alerta=texto_add_br(alerta)
					alerta=alerta & "Quantidade atribuída ao estoque do produto " & vPreDevolucaoItens(i).produto & " referente o pedido " & vPreDevolucaoItens(i).pedido & " é superior a quantidade devolvida."
                end if
            if (vPreDevolucaoItens(i).qtde_estoque_venda+vPreDevolucaoItens(i).qtde_estoque_danificados) < vPreDevolucaoItens(i).qtde then
                alerta=texto_add_br(alerta)
					alerta=alerta & "Quantidade atribuída ao estoque do produto " & vPreDevolucaoItens(i).produto & " referente o pedido " & vPreDevolucaoItens(i).pedido & " é inferior a quantidade devolvida."
                end if
			end if
		next
	
	
	dim campos_a_omitir
	dim vLog()
	dim s_log
	s_log = ""
	campos_a_omitir = "|dt_cadastro|dt_hr_cadastro|finalizado_data|finalizado_data_hora|"

	dim id_email, corpo_mensagem, msg_erro_grava_email, dtHrMensagem
	dim s_email_remetente, s_email_vendedor
	dim r_usuario
	dim s_dados_cliente
	dim s_unidade_negocio
	dim s_descricao_status_devolucao, s_cor_status_devolucao


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rsMail
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	s_email_remetente = getParametroFromCampoTexto(ID_PARAMETRO_EMAILSNDSVC_REMETENTE__PEDIDO_DEVOLUCAO)
	call le_usuario(usuario, r_usuario, msg_erro)

	if alerta = "" then
	'	INICIA A TRANSAÇÃO
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

		if Not cria_recordset_otimista(rsMail, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

        for i=Lbound(vPreDevolucaoItens) to Ubound(vPreDevolucaoItens)
            qtde_itens = 0            
			if Trim(vPreDevolucaoItens(i).id_devolucao)<>"" then
                s = "SELECT Count(*) AS qtde_itens FROM t_PEDIDO_DEVOLUCAO_ITEM WHERE (id_pedido_devolucao = '" & vPreDevolucaoItens(i).id_devolucao & "')"
                if rs.State <> 0 then rs.Close
				rs.Open s, cn
                if CInt(rs("qtde_itens")) > 1 then
                    for j=Lbound(vPreDevolucaoItens) to Ubound(vPreDevolucaoItens)
			            if Trim(vPreDevolucaoItens(j).id_devolucao)<>"" then
                            if vPreDevolucaoItens(j).id_devolucao = vPreDevolucaoItens(i).id_devolucao then qtde_itens=qtde_itens+1
                            end if
                        next
                        if CInt(rs("qtde_itens")) > qtde_itens then
                            alerta = "Não foram especificadas todas as unidades de todos os itens do pedido " & vPreDevolucaoItens(i).pedido & "."
                            Exit For
                            end if
                    end if
                end if 'Trim(vPreDevolucaoItens(i).id_devolucao)<>""
            next

		for i=Lbound(vPreDevolucaoItens) to Ubound(vPreDevolucaoItens)
			if Trim(vPreDevolucaoItens(i).id_devolucao)<>"" then
					if alerta = "" then
						s = "SELECT * FROM t_PEDIDO_DEVOLUCAO_ITEM WHERE (id_pedido_devolucao = '" & vPreDevolucaoItens(i).id_devolucao & "'" & _
                                " AND fabricante = '" & vPreDevolucaoItens(i).fabricante & "'" & _
                                " AND produto = '" & vPreDevolucaoItens(i).produto & "'" & _
                                ")"
                        if rs.State <> 0 then rs.Close
						rs.Open s, cn
						if Not rs.Eof then
                            rs("qtde_estoque_venda")=converte_numero(vPreDevolucaoItens(i).qtde_estoque_venda)
                            rs("qtde_estoque_danificado")=converte_numero(vPreDevolucaoItens(i).qtde_estoque_danificados)
                            rs.Update
                        else
                            alerta = "Não foi encontrada a devolução."
                            end if

						if Err <> 0 then
						'	~~~~~~~~~~~~~~~~
							cn.RollbackTrans
						'	~~~~~~~~~~~~~~~~
							Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
							end if
						
						log_via_vetor_carrega_do_recordset rs, vLog, campos_a_omitir
						s_log = log_via_vetor_monta_inclusao(vLog)
						
						if rs.State <> 0 then rs.Close
							
						if s_log <> "" then grava_log usuario, "", vPreDevolucaoItens(i).pedido, "", "PED DEVOL RECEBE", s_log
						end if

				    
                    if alerta = "" then
			            s = "SELECT * FROM t_PEDIDO_DEVOLUCAO WHERE (id = '" & vPreDevolucaoItens(i).id_devolucao & "')"
						if rs.State <> 0 then rs.Close
                        rs.Open s, cn
                        if Not rs.Eof then
                            rs("status")=COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA
                            rs("status_usuario")=usuario
                            rs("status_data")=Date
                            rs("status_data_hora")=Now
                            rs("st_mercadoria_recebida")=1
                            rs("usuario_mercadoria_recebida")=usuario
                            rs("dt_mercadoria_recebida")=Date
                            rs("dt_hr_mercadoria_recebida")=Now
                            rs("usuario_ult_atualizacao")=usuario
                            rs("dt_ult_atualizacao")=Date
                            rs("dt_hr_ult_atualizacao")=Now
                            rs.Update
                        else
                            alerta = "Não foi encontrado o ID da devolução."
                            end if
                        end if 'if alerta = ""

					'Este é o primeiro item de um novo bloco de devolução?
                    if vPreDevolucaoItens(i).id_devolucao <> devolucao_a then
						'Envia email de aviso para o vendedor
						if alerta = "" then
							'Foi encontrado o email para ser usado como remetente da mensagem?
							if s_email_remetente <> "" then
								s_email_vendedor = ""
								s_dados_cliente = ""
								s_unidade_negocio = ""
								s_descricao_status_devolucao = ""
								s_cor_status_devolucao = ""

								s = "SELECT" & _
										" tP.vendedor," & _
										" tU.email" & _
									" FROM t_PEDIDO tP" & _
										" INNER JOIN t_USUARIO tU ON (tP.vendedor = tU.usuario)" & _
									" WHERE" & _
										" (tP.pedido = '" & vPreDevolucaoItens(i).pedido & "')"
								if rsMail.State <> 0 then rsMail.Close
								rsMail.Open s, cn
								if Not rsMail.Eof then
									s_email_vendedor = LCase(Trim("" & rsMail("email")))
									end if

								'Se encontrou e-mail do vendedor para enviar mensagem de aviso, obtém demais informações para a montagem da mensagem
								if s_email_vendedor <> "" then
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
											" (p.pedido = '" & vPreDevolucaoItens(i).pedido & "')"
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

									if (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__BS) Or (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__VRF) then
										obtem_descricao_status_devolucao COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA, s_descricao_status_devolucao, s_cor_status_devolucao

										dtHrMensagem = Now

										corpo_mensagem = "O status da devolução nº " & vPreDevolucaoItens(i).id_devolucao & " do pedido " & vPreDevolucaoItens(i).pedido & " foi alterado para '" & s_descricao_status_devolucao & "' por '" & usuario & "' (" & r_usuario.nome_iniciais_em_maiusculas & ") em " & formata_data_hora_sem_seg(dtHrMensagem) & _
														vbCrLf & _
														"Pedido: " & vPreDevolucaoItens(i).pedido & _
														vbCrLf & _
														"Devolução nº " & vPreDevolucaoItens(i).id_devolucao & _
														vbCrLf & _
														s_dados_cliente & _
														vbCrLf & vbCrLf & _
														"Atenção: esta é uma mensagem automática, NÃO responda a este e-mail!"

										'Envia e-mail para o vendedor
										EmailSndSvcGravaMensagemParaEnvio s_email_remetente, _
																		"", _
																		s_email_vendedor, _
																		"", _
																		"", _
																		"Status da devolução nº " & vPreDevolucaoItens(i).id_devolucao & " do pedido " & vPreDevolucaoItens(i).pedido & " alterado para '" & s_descricao_status_devolucao & "'", _
																		corpo_mensagem, _
																		Now, _
																		id_email, _
																		msg_erro_grava_email
										end if 'if (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__BS) Or (s_unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__VRF)
									end if 'if s_email_vendedor <> ""
								end if 'if s_email_remetente <> ""
							end if 'if alerta = ""

                        ' grava mensagem no bloco de notas de devolução
                        if alerta = "" then
                            s_id_item_devolvido = ""
                        
	                        s = "SELECT" & _
			                        " id" & _
		                        " FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
		                        " WHERE" & _
			                        " (pedido = '" & vPreDevolucaoItens(i).pedido & "')" & _
		                        " ORDER BY" & _
			                        " id DESC"
						    if rs.State <> 0 then rs.Close
	                        rs.Open s, cn
	                        if Not rs.Eof then s_id_item_devolvido = Trim("" & rs("id"))
                
                            ' gera NSU para gravar mensagem no bloco de notas de devolução
                            if Not fin_gera_nsu(T_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS, intNsuNovoBlocoNotas, msg_erro) then
			                    alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO NO BLOCO DE NOTAS DE DEVOLUÇÃO (" & msg_erro & ")"
		                    else
			                    if intNsuNovoBlocoNotas <= 0 then
				                    alerta = "NSU DO BLOCO DE NOTAS DE DEVOLUÇÃO GERADO É INVÁLIDO (" & intNsuNovoBlocoNotas & ")"
				                    end if
			                    end if
                            if alerta = "" then
                                s = "SELECT * FROM t_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS WHERE (id = -1)"
                                if rs.State <> 0 then rs.Close
						        rs.Open s, cn
			                    rs.AddNew 
			                    rs("id")=intNsuNovoBlocoNotas
			                    rs("id_item_devolvido")=s_id_item_devolvido
			                    rs("usuario")="SISTEMA"
			                    rs("mensagem")="Mercadoria(s) da devolução nº " & vPreDevolucaoItens(i).id_devolucao & " foi recebida no CD."
			                    rs.Update 

						        if Err <> 0 then
						        '	~~~~~~~~~~~~~~~~
							        cn.RollbackTrans
						        '	~~~~~~~~~~~~~~~~
							        Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
							        end if
                                end if
                            end if       ' if alerta = ""      
                        end if   ' vPreDevolucaoItens(i).id_devolucao <> devolucao_a

                devolucao_a = vPreDevolucaoItens(i).id_devolucao
				end if  'if Trim(vPreDevolucaoItens(i).id_devolucao)<>""
			next
			
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("RelPedidoPreDevolucaoMercadoriaRecebe.asp?origem=A&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
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