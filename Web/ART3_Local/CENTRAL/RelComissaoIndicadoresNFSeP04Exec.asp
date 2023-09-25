<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================
'	  RelComissaoIndicadoresNFSeP04Exec.asp
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
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	class cl_REL_COMISSAO_NFSe_N2
		dim vendedor
		dim indicador
		dim id_indicador
		dim desempenho_nota
		dim razao_social_nome
		dim NFSe_razao_social
		dim comissao_cartao_status
		dim comissao_cartao_cpf
		dim comissao_cartao_titular
		dim banco
		dim banco_nome
		dim agencia
		dim agencia_dv
		dim tipo_conta
		dim conta_operacao
		dim conta
		dim conta_dv
		dim favorecido
		dim favorecido_cnpj_cpf
		dim vl_total_preco_venda
		dim vl_total_preco_NF
		dim vl_total_RT
		dim vl_total_RA_bruto
		dim vl_total_RA_liquido
		dim vl_total_RA_dif
		dim cor_linha_total
		dim mensagem_desconto

		public vN3Pedido
		public vN3Desconto

		Private Sub Class_Initialize
			vN3Pedido = Array()
			vN3Desconto = Array()
		End Sub

		Sub AddN3Pedido(newItem)
		'	INICIALMENTE, O ARRAY ENCONTRA-SE EM UM ESTADO EM QUE LBOUND() RETORNA 0 (ZERO) E UBOUND() RETORNA -1 (UM NEGATIVO)
			ReDim Preserve vN3Pedido(UBound(vN3Pedido) + 1)
			set vN3Pedido(UBound(vN3Pedido)) = newItem
		End Sub
		
		Sub AddN3Desconto(newItem)
		'	INICIALMENTE, O ARRAY ENCONTRA-SE EM UM ESTADO EM QUE LBOUND() RETORNA 0 (ZERO) E UBOUND() RETORNA -1 (UM NEGATIVO)
			ReDim Preserve vN3Desconto(UBound(vN3Desconto) + 1)
			set vN3Desconto(UBound(vN3Desconto)) = newItem
		End Sub
		end class

	class cl_REL_COMISSAO_NFSe_N3_DESCONTO
		dim id_orcamentista_e_indicador_desconto
		dim descricao
		dim valor
		dim ordenacao
		end class

	class cl_REL_COMISSAO_NFSe_N3_PEDIDO
		dim pedido
		dim orcamento
		dim operacao
		dim id_cfg_tabela_origem
		dim id_registro_tabela_origem
		dim loja
		dim st_comissao_original
		dim st_pagto
		dim nome_cliente
		dim data_pedido
		dim perc_RT
		dim vl_preco_venda
		dim vl_preco_NF
		dim vl_RT
		dim vl_RA_bruto
		dim vl_RA_liquido
		dim vl_RA_dif
		dim vl_comissao
		dim sinal
		dim cor_sinal
		dim cor_linha
		end class

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, rs2, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	alerta = ""
	
	dim blnErroFatal
	blnErroFatal = False

	dim s, s_aux, s_filtro, s_filtro_indicador
	s_filtro_indicador = ""

'	FILTROS
	dim c_cnpj_nfse
	dim ckb_id_indicador
	dim c_dt_entregue_termino, dt_entregue_termino
	dim rb_visao, blnVisaoSintetica, proc_comissao_request_guid

	c_cnpj_nfse = retorna_so_digitos(Request.Form("c_cnpj_nfse"))
	ckb_id_indicador = Trim(Request.Form("ckb_id_indicador"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))
	dt_entregue_termino = StrToDate(c_dt_entregue_termino)
	rb_visao = Trim(Request.Form("rb_visao"))
	proc_comissao_request_guid = Trim(Request.Form("proc_comissao_request_guid"))
	
	blnVisaoSintetica = False
	if rb_visao = "SINTETICA" then blnVisaoSintetica = True
	
	if alerta = "" then
		if c_dt_entregue_termino <> "" then
			if Not IsDate(StrToDate(c_dt_entregue_termino)) then
				alerta = "DATA DE TÉRMINO DO PERÍODO É INVÁLIDA."
				end if
			end if
		end if
	
	if alerta = "" then
		if c_cnpj_nfse = "" then
			alerta = "CNPJ do emitente da NFS-e não foi informado"
		elseif Not cnpj_cpf_ok(c_cnpj_nfse) then
			alerta = "CNPJ do emitente da NFS-e é inválido"
			end if
		end if
	
	if alerta = "" then
		if ckb_id_indicador = "" then
			alerta = "Nenhum indicador foi selecionado"
			end if
		end if
	
	'MONTA DESCRIÇÃO DE INDICADOR(ES) SELECIONADO(S) P/ EXIBIÇÃO NO FILTRO
	s = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (Id IN (" & ckb_id_indicador & ")) ORDER BY apelido"
	if rs.State <> 0 then rs.Close
	rs.Open s, cn
	do while Not rs.Eof
		if s_filtro_indicador <> "" then s_filtro_indicador = s_filtro_indicador & "<br />"
		s_filtro_indicador = s_filtro_indicador & "<span class='N'>" & Trim("" & rs("apelido")) & " - " & Trim("" & rs("razao_social_nome_iniciais_em_maiusculas")) & "</span>"
		rs.MoveNext
		loop
	if rs.State <> 0 then rs.Close

	if alerta = "" then
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
		end if 'if alerta = ""




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' monta html bloco msg erro
'
function monta_html_bloco_msg_erro(byval mensagem_erro)
dim s_html
	s_html = "<br />" & _
		"<p class='T' style='color:red;'>Mensagem de Erro</p>" & _
		"<div class='MtAlerta' style='width:600px;font-weight:bold;' align='center'><p style='margin:5px 2px 5px 2px;'>" & mensagem_erro & "</p></div>" & _
		"<br />"
	monta_html_bloco_msg_erro = s_html
end function


' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const VENDA_NORMAL = "VEN"
const DEVOLUCAO = "DEV"
const PERDA = "PER"
dim r, tN1, tN2, tN3Desc, tN3Ped
dim s, s_aux, s_sql, x, cab_table, cab, indicador_a, vendedor_a, n_reg, n_reg_total, qtde_indicadores
dim vl_preco_venda, vl_sub_total_preco_venda, vl_total_preco_venda
dim vl_preco_NF, vl_sub_total_preco_NF, vl_total_preco_NF
dim vl_RT, vl_sub_total_RT, vl_total_RT
dim vl_RA_bruto, vl_sub_total_RA_bruto, vl_total_RA_bruto
dim vl_RA_liquido, vl_sub_total_RA_liquido, vl_total_RA_liquido
dim vl_RA_diferenca, vl_sub_total_RA_diferenca, vl_total_RA_diferenca
dim perc_RT
dim s_where, s_where_aux, s_where_venda, s_where_devolucao, s_where_perdas
dim s_where_comissao_paga, s_where_comissao_descontada, s_where_st_pagto, s_where_dt_st_pagto
dim s_cor, s_sinal, s_cor_sinal
dim s_banco, s_banco_nome, s_agencia, s_conta, s_favorecido, s_favorecido_cnpj_cpf
dim s_nome_cliente, s_desempenho_nota
dim s_id, s_id_base, s_class, s_class_td, idx_bloco, s_new_cab
dim s_lista_completa_venda_normal, s_lista_completa_devolucao, s_lista_completa_perda
dim v_desconto_descricao(), v_desconto_valor(), contador, valor_desconto, qtde_registro_desc, msg_desconto, IdIndicador_anterior
dim s_log, s_erro_fatal, lngNsuN1, lngNsuN2
dim vN2, rxN3Desconto, rxN3Pedido
dim iN2, iN3Desc, iN3Ped
dim s_indicador_nome

'	CRITÉRIOS COMUNS
	s_where = ""

	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.Id IN (" & ckb_id_indicador & "))"

'	CRITÉRIO: COMISSÃO PAGA
'	A) VENDAS
	s_where_comissao_paga = " (t_PEDIDO.comissao_paga = " & COD_COMISSAO_NAO_PAGA & ")"
		
'	B) PERDAS/DEVOLUÇÕES
	s_where_comissao_descontada = " (comissao_descontada = " & COD_COMISSAO_NAO_DESCONTADA & ")"

'	CRITÉRIO: STATUS DE PAGAMENTO
	s_where_dt_st_pagto = ""
	if IsDate(c_dt_entregue_termino) then
		s_where_dt_st_pagto = s_where_dt_st_pagto & " AND (t_PEDIDO__BASE.dt_st_pagto < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		end if

	s_where_st_pagto = " ((t_PEDIDO__BASE.st_pagto = '" & ST_PAGTO_PAGO & "')" & s_where_dt_st_pagto & ")"

	
'	CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	s_where_venda = ""
	if IsDate(c_dt_entregue_termino) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.entregue_data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		end if
		
	if s_where_comissao_paga <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (" & s_where_comissao_paga & ")"
		end if

	if s_where_st_pagto <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (" & s_where_st_pagto & ")"
		end if
	
'	CRITÉRIOS PARA DEVOLUÇÕES
	s_where_devolucao = ""
	if IsDate(c_dt_entregue_termino) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		end if

	if s_where_comissao_descontada <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (" & s_where_comissao_descontada & ")"
		end if

'	CRITÉRIOS PARA PERDAS
	s_where_perdas = ""
	if IsDate(c_dt_entregue_termino) then
		if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		s_where_perdas = s_where_perdas & " (t_PEDIDO_PERDA.data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		end if
	
	if s_where_comissao_descontada <> "" then
		if s_where_perdas <> "" then s_where_perdas = s_where_perdas & " AND"
		s_where_perdas = s_where_perdas & " (" & s_where_comissao_descontada & ")"
		end if
	
	
'	VENDAS NORMAIS
	s_where_aux = s_where
	if (s_where_aux <> "") And (s_where_venda <> "") then s_where_aux = s_where_aux & " AND"
	s_where_aux = s_where_aux & s_where_venda
	if s_where_aux <> "" then s_where_aux = " AND" & s_where_aux

	s_sql = "SELECT" & _
			" '" & VENDA_NORMAL & "' AS operacao," & _
			" t_CFG_TABELA_ORIGEM.id AS id_cfg_tabela_origem," & _
			" t_PEDIDO.pedido AS id_registro," & _
			" t_PEDIDO.comissao_paga AS status_comissao," & _
			" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota," & _
			" t_ORCAMENTISTA_E_INDICADOR.banco," & _
			" t_ORCAMENTISTA_E_INDICADOR.Id AS IdIndicador," & _
			" t_LOJA.id_plano_contas_empresa_comissao_indicador," & _
			" t_PEDIDO__BASE.indicador," & _
			" t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
			" t_PEDIDO.endereco_nome AS nome," & _
			" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
			" t_CLIENTE.nome," & _
			" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO.entregue_data AS data," & _
			" t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			" t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			" t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA," & _
			" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
			" Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.preco_venda) AS total_preco_venda," & _
			" Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.preco_NF) AS total_preco_NF" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			" INNER JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			" LEFT JOIN t_CFG_TABELA_ORIGEM ON (nome_tabela = 't_PEDIDO')" & _
			" LEFT JOIN t_LOJA ON (t_PEDIDO.loja = t_LOJA.loja)" & _
			" WHERE (t_PEDIDO.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
			s_where_aux & _
			" GROUP BY t_CFG_TABELA_ORIGEM.id, t_PEDIDO.pedido, t_PEDIDO.comissao_paga, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_ORCAMENTISTA_E_INDICADOR.Id, t_LOJA.id_plano_contas_empresa_comissao_indicador, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome," & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO.entregue_data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"

'	ITENS DEVOLVIDOS
	s_where_aux = s_where
	if (s_where_aux <> "") And (s_where_devolucao <> "") then s_where_aux = s_where_aux & " AND"
	s_where_aux = s_where_aux & s_where_devolucao
	if s_where_aux <> "" then s_where_aux = " WHERE " & s_where_aux

	s_sql = s_sql & " UNION ALL " & _
			"SELECT" & _
			" '" & DEVOLUCAO & "' AS operacao," & _
			" t_CFG_TABELA_ORIGEM.id AS id_cfg_tabela_origem," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.id AS id_registro," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.comissao_descontada AS status_comissao," & _
			" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota," & _
			" t_ORCAMENTISTA_E_INDICADOR.banco," & _
			" t_ORCAMENTISTA_E_INDICADOR.Id AS IdIndicador," & _
			" t_LOJA.id_plano_contas_empresa_comissao_indicador," & _
			" t_PEDIDO__BASE.indicador," & _
			" t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome AS nome," & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data," & _
			" t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			" t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			" t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA," & _
			" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
			" Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS total_preco_venda," & _
			" Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_NF) AS total_preco_NF" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido)" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			" INNER JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			" LEFT JOIN t_CFG_TABELA_ORIGEM ON (nome_tabela = 't_PEDIDO_ITEM_DEVOLVIDO')" & _
			" LEFT JOIN t_LOJA ON (t_PEDIDO.loja = t_LOJA.loja)" & _
			s_where_aux & _
			" GROUP BY t_CFG_TABELA_ORIGEM.id, t_PEDIDO_ITEM_DEVOLVIDO.id, t_PEDIDO_ITEM_DEVOLVIDO.comissao_descontada, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_ORCAMENTISTA_E_INDICADOR.Id, t_LOJA.id_plano_contas_empresa_comissao_indicador, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome," & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"

'	PERDAS
	s_where_aux = s_where
	if (s_where_aux <> "") And (s_where_perdas <> "") then s_where_aux = s_where_aux & " AND"
	s_where_aux = s_where_aux & s_where_perdas
	if s_where_aux <> "" then s_where_aux = " WHERE " & s_where_aux

	s_sql = s_sql & " UNION ALL " & _
			"SELECT" & _
			" '" & PERDA & "' AS operacao," & _
			" t_CFG_TABELA_ORIGEM.id AS id_cfg_tabela_origem," & _
			" t_PEDIDO_PERDA.id AS id_registro," & _
			" t_PEDIDO_PERDA.comissao_descontada AS status_comissao," & _
			" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota," & _
			" t_ORCAMENTISTA_E_INDICADOR.banco," & _
			" t_ORCAMENTISTA_E_INDICADOR.Id AS IdIndicador," & _
			" t_LOJA.id_plano_contas_empresa_comissao_indicador," & _
			" t_PEDIDO__BASE.indicador," & _
			" t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
			" t_PEDIDO.endereco_nome AS nome," & _
			" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
			" t_CLIENTE.nome," & _
			" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
			" t_PEDIDO.loja AS loja, t_PEDIDO.numero_loja," & _
			" t_PEDIDO_PERDA.data AS data," & _
			" t_PEDIDO.pedido AS pedido, t_PEDIDO.orcamento AS orcamento," & _
			" t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto," & _
			" t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA," & _
			" t_PEDIDO__BASE.perc_desagio_RA_liquida," & _
			" Sum(-t_PEDIDO_PERDA.valor) AS total_preco_venda," & _
			" Sum(-t_PEDIDO_PERDA.valor) AS total_preco_NF" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" INNER JOIN t_PEDIDO_PERDA ON (t_PEDIDO.pedido=t_PEDIDO_PERDA.pedido)" & _
			" INNER JOIN t_CLIENTE ON (t_PEDIDO__BASE.id_cliente=t_CLIENTE.id)" & _
			" INNER JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
			" LEFT JOIN t_CFG_TABELA_ORIGEM ON (nome_tabela = 't_PEDIDO_PERDA')" & _
			" LEFT JOIN t_LOJA ON (t_PEDIDO.loja = t_LOJA.loja)" & _
			s_where_aux & _
			" GROUP BY t_CFG_TABELA_ORIGEM.id, t_PEDIDO_PERDA.id, t_PEDIDO_PERDA.comissao_descontada, t_ORCAMENTISTA_E_INDICADOR.desempenho_nota, t_ORCAMENTISTA_E_INDICADOR.banco, t_ORCAMENTISTA_E_INDICADOR.Id, t_LOJA.id_plano_contas_empresa_comissao_indicador, t_PEDIDO__BASE.indicador, t_PEDIDO__BASE.vendedor,"
	
	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome," & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if
		
	s_sql = s_sql & _
			" t_PEDIDO.loja, t_PEDIDO.numero_loja, t_PEDIDO_PERDA.data, t_PEDIDO.pedido, t_PEDIDO.orcamento, t_PEDIDO__BASE.perc_RT, t_PEDIDO__BASE.st_pagto, t_PEDIDO__BASE.vl_total_RA_liquido, t_PEDIDO__BASE.st_tem_desagio_RA, t_PEDIDO__BASE.perc_desagio_RA_liquida"
	
	s_sql = "SELECT " & _
				"*" & _
			" FROM (" & _
				s_sql & _
				") t" & _
			" ORDER BY t.vendedor, t.desempenho_nota, t.indicador, t.numero_loja, t.data, t.pedido, t.total_preco_venda DESC"

	'CABEÇALHO
	cab_table = "<table cellspacing='0' id='tableDados'>" & chr(13)
	cab = "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='MDTE tdCkb' align='center' valign='bottom' nowrap><input type='checkbox' name='ckb_comissao_paga_tit_bloco' id='ckb_comissao_paga_tit_bloco' class='CKB_COM VISAO_ANALIT' onclick='trata_ckb_onclick();' /></td>" & chr(13) & _
		  "		<td class='MTD tdLoja' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>Loja</span></td>" & chr(13) & _
		  "		<td class='MTD tdOrcamento' align='left' valign='bottom' nowrap><span class='R VISAO_ANALIT'>Nº Orçam</span></td>" & chr(13) & _
		  "		<td class='MTD tdPedido' align='left' valign='bottom' nowrap><span class='R VISAO_ANALIT'>Nº Pedido</span></td>" & chr(13) & _
		  "		<td class='MTD tdData' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>Data</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlPedido' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL Pedido</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRT' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>COM (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRABruto' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Bruto (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRALiq' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Líq (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRADif' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Dif (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdStPagto' align='left' valign='bottom'><span class='R VISAO_ANALIT' style='font-weight:bold;'>St Pagto</span></td>" & chr(13) & _
		  "		<td class='MTD tdSinal' align='center' valign='bottom'><span class='Rc VISAO_ANALIT' style='font-weight:bold;'>+/-</span></td>" & chr(13) & _
		  "		<td valign='bottom' class='notPrint BkgWhite' align='left'>&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & "_NNNNN_" & chr(34) & ");' title='exibe ou oculta os dados'><img src='../botao/view_bottom.png' border='0'></a></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	idx_bloco = 0
	qtde_indicadores = 0
	vl_sub_total_preco_venda = 0
	vl_total_preco_venda = 0
	vl_sub_total_preco_NF = 0
	vl_total_preco_NF = 0
	vl_sub_total_RT = 0
	vl_total_RT = 0
	vl_sub_total_RA_bruto = 0
	vl_total_RA_bruto = 0
	vl_sub_total_RA_liquido = 0
	vl_total_RA_liquido = 0
	vl_sub_total_RA_diferenca = 0
	vl_total_RA_diferenca = 0
	s_lista_completa_venda_normal = ""
	s_lista_completa_devolucao = ""
	s_lista_completa_perda = ""
	indicador_a = "XXXXXXXXXXXX"
	vendedor_a = "XXXXXXXXXXXX"
	
	s_log = ""
	redim vN2(0)
	set vN2(UBound(vN2)) = new cl_REL_COMISSAO_NFSe_N2
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE INDICADOR?
	'	A MUDANÇA DE VENDEDOR TAMBÉM É INTERPRETADA COMO MUDANÇA DE INDICADOR JÁ QUE É NECESSÁRIO INICIAR UM NOVO BLOCO
		if (Trim("" & r("indicador")) <> indicador_a) Or (Trim("" & r("vendedor")) <> vendedor_a) then
			' FECHA TABELA DO INDICADOR ANTERIOR
			if n_reg_total > 0 then
				s_cor="black"
				if vl_sub_total_preco_venda < 0 then s_cor="red"
				if vl_sub_total_RT < 0 then s_cor="red"
				if vl_sub_total_RA_bruto < 0 then s_cor="red"
				if vl_sub_total_RA_liquido < 0 then s_cor="red"

				s_sql = "SELECT" & _
							" tDESC.id" & _
							", tDESC.descricao" & _
							", tDESC.valor" & _
							", tDESC.ordenacao" & _
						" FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO tDESC" & _
							" INNER JOIN t_ORCAMENTISTA_E_INDICADOR tOI ON (tDESC.apelido=tOI.apelido)" & _
						" WHERE" & _
							" (tOI.Id = " & IdIndicador_anterior & ")" & _
						" ORDER BY" & _
							" tDESC.ordenacao"
				if rs2.State <> 0 then rs2.Close
				rs2.Open s_sql, cn
				msg_desconto = ""
				if Not rs2.Eof then
					valor_desconto = 0
					contador = 0
					qtde_registro_desc = 0
					Erase v_desconto_descricao
					Erase v_desconto_valor
					do while Not rs2.EoF
						redim preserve v_desconto_descricao(contador)
						redim preserve  v_desconto_valor(contador)
						v_desconto_descricao(contador) = rs2("descricao")
						v_desconto_valor(contador) = rs2("valor")
						valor_desconto = valor_desconto + v_desconto_valor(contador)
						qtde_registro_desc = qtde_registro_desc + 1
						contador = contador + 1
						
						set rxN3Desconto = new cl_REL_COMISSAO_NFSe_N3_DESCONTO
						with rxN3Desconto
							.id_orcamentista_e_indicador_desconto = rs2("id")
							.descricao = Trim("" & rs2("descricao"))
							.valor = rs2("valor")
							.ordenacao = rs2("ordenacao")
							end with
						vN2(UBound(vN2)).AddN3Desconto(rxN3Desconto)

						rs2.MoveNext
						loop
					msg_desconto =  "Planilha de Desconto:&nbsp; " & Cstr(qtde_registro_desc) & " registro(s) no valor total de&nbsp;" & formata_moeda(valor_desconto)
				else
					msg_desconto = ""
					end if

				x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
						"		<td class='MTBE' colspan='5' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>TOTAL:</span></td>" & chr(13) & _
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_bruto) & "</span></td>" & chr(13) & _
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _
						"		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
						"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
						"	</tr>" & chr(13) & _
						"	<tr>" & chr(13)
				
				with vN2(UBound(vN2))
					.vl_total_preco_venda = vl_sub_total_preco_venda
					.vl_total_preco_NF = vl_sub_total_preco_NF
					.vl_total_RT = vl_sub_total_RT
					.vl_total_RA_bruto = vl_sub_total_RA_bruto
					.vl_total_RA_liquido = vl_sub_total_RA_liquido
					.vl_total_RA_dif = vl_sub_total_RA_diferenca
					.cor_linha_total = s_cor
					.mensagem_desconto = msg_desconto
					end with

				x = x & "		<td align='left' colspan='10' nowrap>"
				if msg_desconto <> "" then
					x = x & "<span class='Cd' ><a href='javascript:abreDesconto(" & Cstr(idx_bloco) & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>" & msg_desconto & "</a></span>"
				else
					x = x & "<span class='Rd' style='color: black;'></span>"
					end if
				x = x & "</td>" & chr(13)
				
				s_cor = "black"

				x = x & "		<td align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13)

				if (vl_sub_total_RT + vl_sub_total_RA_liquido) >= 0  then
					x = x & "		<td align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT + vl_sub_total_RA_liquido) & "</span></td>" & chr(13)
					end if
				
				x = x & "	</tr>" & chr(13)
				if msg_desconto <> "" then
					x = x & "	<tr>" & chr(13) & _
							"		<td  class='VISAO_ANALIT' id='table_Desconto_" & Cstr(idx_bloco) & "' colspan='15'>" & chr(13) & _
							"			<table colspan='2' align='left' >"& chr(13)
					for contador = 0 to Ubound(v_desconto_descricao)
						x = x & "				<tr>" & chr(13) & _
								"					<td width='15'>&nbsp;</td>" & chr(13) & _
								"					<td  align='left' width='400'><span class='Cd'style='color: red;'>" & v_desconto_descricao(contador) & "</span></td>" & chr(13) & _
								"					<td align='left'><span class='Cd' style='color:red;'> R$ " & formata_moeda(v_desconto_valor(contador)) & "</span></td>" & chr(13) & _
								"				</tr>" & chr(13)
						next
					x = x & "			</table>" & chr(13) & _
							"		</td>" & chr(13) & _
							"	</tr>" & chr(13)
					end if 'if msg_desconto <> ""
				
				x = x & "</table>" & chr(13)
				Response.Write x
				x = "<br />" & chr(13)
				end if 'if n_reg_total > 0


			idx_bloco = idx_bloco + 1
			qtde_indicadores = qtde_indicadores + 1

			indicador_a = Trim("" & r("indicador"))
			vendedor_a = Trim("" & r("vendedor"))

			if Trim("" & vN2(UBound(vN2)).indicador) <> "" then
				redim preserve vN2(UBound(vN2)+1)
				set vN2(UBound(vN2)) = new cl_REL_COMISSAO_NFSe_N2
				end if
			with vN2(UBound(vN2))
				.indicador = Trim("" & r("indicador"))
				.vendedor = Trim("" & r("vendedor"))
				.id_indicador = r("IdIndicador")
				end with

			n_reg = 0
			vl_sub_total_preco_venda = 0
			vl_sub_total_preco_NF = 0
			vl_sub_total_RT = 0
			vl_sub_total_RA_bruto = 0
			vl_sub_total_RA_liquido = 0
			vl_sub_total_RA_diferenca = 0

			s_sql = "SELECT " & _
						"*" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (Id = " & Trim("" & r("IdIndicador")) & ")"
			if rs.State <> 0 then rs.Close
			rs.Open s_sql, cn
			if Not rs.Eof then
				s_banco = Trim("" & rs("banco"))
				s_agencia = Trim("" & rs("agencia"))
				s_conta = Trim("" & rs("conta"))
				s_favorecido = Trim("" & rs("favorecido"))
				s_favorecido_cnpj_cpf = Trim("" & rs("favorecido_cnpj_cpf"))
				if s_favorecido_cnpj_cpf <> "" then s_favorecido_cnpj_cpf = cnpj_cpf_formata(s_favorecido_cnpj_cpf)
				s_banco_nome = x_banco(s_banco)
				if (s_banco <> "") And (s_banco_nome <> "") then s_banco = s_banco & " - " & s_banco_nome
			else
				s_banco = ""
				s_banco_nome = ""
				s_agencia = ""
				s_conta = ""
				s_favorecido = ""
				s_favorecido_cnpj_cpf = ""
				end if
			
			s_indicador_nome = x_orcamentista_e_indicador(Trim("" & r("indicador")))

			with vN2(UBound(vN2))
				.razao_social_nome = s_indicador_nome
				.banco = Trim("" & rs("banco"))
				.banco_nome = s_banco_nome
				.agencia = Trim("" & rs("agencia"))
				.agencia_dv = Trim("" & rs("agencia_dv"))
				.tipo_conta = Trim("" & rs("tipo_conta"))
				.conta_operacao = Trim("" & rs("conta_operacao"))
				.conta = Trim("" & rs("conta"))
				.conta_dv = Trim("" & rs("conta_dv"))
				.favorecido = Trim("" & rs("favorecido"))
				.favorecido_cnpj_cpf = Trim("" & rs("favorecido_cnpj_cpf"))
				end with

			if n_reg_total > 0 then x = x & "<br />" & chr(13)
			s_desempenho_nota = Trim("" & r("desempenho_nota"))
			vN2(UBound(vN2)).desempenho_nota = Trim("" & r("desempenho_nota"))
			if s_desempenho_nota = "" then
				s_desempenho_nota = "&nbsp;"
			else
				s_desempenho_nota = "(" & s_desempenho_nota & ") "
				end if
			s = Trim("" & r("indicador"))
			if (s<>"") And (s_indicador_nome<>"") then s = s & " - "
			s = s & s_indicador_nome
			x = x & Replace(cab_table, "tableDados", "tableDados_" & idx_bloco)
			if s <> "" then x = x & "	<tr>" & chr(13) & _
									"		<td class='MDTE' colspan='12' align='left' valign='bottom' style='background:azure;'>" & chr(13) & _
									"			<table cellpadding='0' cellspacing='0' width='100%'>" & chr(13) & _
									"				<tr>" & chr(13) & _
									"					<td class='MD' width='85%' align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;" & s_desempenho_nota & s & "</span></td>" & chr(13) & _
									"					<td align='left' valign='bottom' style='background:azure;'><span class='N'>&nbsp;" & Trim("" & r("vendedor")) & "</span></td>" & chr(13) & _
									"				</tr>" & chr(13) & _
									"			</table>" & chr(13) & _
									"		</td>" & chr(13) & _
									"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"	<tr>" & chr(13) & _
									"		<td class='MDTE' colspan='12' align='left' valign='bottom' class='MB' style='background:white;'>" & chr(13) & _
									"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
									"				<tr>" & chr(13) & _
									"					<td align='right' valign='bottom' nowrap><span class='Cn'>Pagamento da Comissão via Cartão:</span></td>" & chr(13) & _
									"					<td width='90%' align='left' valign='bottom' nowrap><span class='Cn'>"
			if rs("comissao_cartao_status") = 1 then
				x = x & "Sim" & " &nbsp; " & cnpj_cpf_formata(Trim("" & rs("comissao_cartao_cpf"))) & " - " & Trim("" & rs("comissao_cartao_titular"))
			else
				x = x & "Não"
				end if

			x = x & _
				"</span></td>" & chr(13) & _
				"				</tr>" & chr(13) & _
				"				<tr>" & chr(13) & _
				"					<td class='MC' align='right' valign='bottom' nowrap><span class='Cn'>Pagamento da Comissão via NFSe:</span></td>" & chr(13) & _
				"					<td class='MC' width='90%' align='left' valign='bottom' nowrap><span class='Cn'>" & chr(13)

			if Trim("" & rs("comissao_NFSe_cnpj")) <> "" then
				x = x & cnpj_cpf_formata(Trim("" & rs("comissao_NFSe_cnpj"))) & " - " & Trim("" & rs("comissao_NFSe_razao_social"))
			else
				x = x & "N.I."
				end if

			with vN2(UBound(vN2))
				.comissao_cartao_status = rs("comissao_cartao_status")
				.comissao_cartao_cpf = Trim("" & rs("comissao_cartao_cpf"))
				.comissao_cartao_titular = Trim("" & rs("comissao_cartao_titular"))
				.NFSe_razao_social = Trim("" & rs("comissao_NFSe_razao_social"))
				end with

			x = x & _
				"</span></td>" & chr(13) & _
				"				</tr>" & chr(13) & _
				"			</table>" & chr(13) & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr>" & chr(13) & _
				"		<td class='MDTE' colspan='12' align='left' valign='bottom' class='MB' style='background:whitesmoke;'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<tr>" & chr(13) & _
				"					<td colspan='3' align='left' valign='bottom'><span class='Cn'>Banco: " & s_banco & "</span></td>" & chr(13) & _
				"				</tr>" & chr(13) & _
				"				<tr>" & chr(13) & _
				"					<td class='MTD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>Agência: " & Trim("" & rs("agencia"))
			if Trim("" & rs("agencia_dv")) <> "" then
				x = x & "-" & rs("agencia_dv")
				end if
	
			x = x & "</span></td>" & chr(13)
			
			x = x & "					<td class='MC MD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>"
			if Trim("" & rs("tipo_conta")) <> "" then
				if rs("tipo_conta") = "P" then
					x = x & "C/P: "
				elseif rs("tipo_conta") = "C" then
					x = x & "C/C: "
					end if
			else
				x = x & "Conta: "
				end if

			if Trim("" & rs("conta_operacao")) <> "" then
				x = x & rs("conta_operacao") & "-"
				end if
	
			x = x & Trim("" & rs("conta"))
	
			if Trim("" & rs("conta_dv")) <> "" then
				x = x & "-" & rs("conta_dv")
				end if
			x = x & "</span></td>" & chr(13)

			x = x & "					<td class='MC' width='60%' align='left' valign='bottom'><span class='Cn'>Favorecido: " & s_favorecido & "</span></td>" & chr(13) & _
					"				</tr>" & chr(13)

			if Len(retorna_so_digitos(s_favorecido_cnpj_cpf)) = 11 then
				s_aux = "CPF"
			else
				s_aux = "CNPJ"
				end if

			x = x & _
				"				<tr>" & chr(13) & _
				"					<td colspan='3' class='MC' align='left' valign='bottom'><span class='Cn'>" & s_aux & ": " & s_favorecido_cnpj_cpf & "</span></td>" & chr(13) & _
				"				</tr>" & chr(13)

			x = x & _
				"			</table>" & chr(13) & _
				"		</td>" & chr(13) & _
				"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13)

			s_new_cab = Replace(cab, "ckb_comissao_paga_tit_bloco", "ckb_comissao_paga_tit_bloco_" & idx_bloco)
			s_new_cab = Replace(s_new_cab, "trata_ckb_onclick();", "trata_ckb_onclick(" & chr(34) & idx_bloco & chr(34) & ");")
			s_new_cab = Replace(s_new_cab, "_NNNNN_", CStr(idx_bloco))
			x = x & s_new_cab
			end if 'if (Trim("" & r("indicador")) <> indicador_a) Or (Trim("" & r("vendedor")) <> vendedor_a)
		

		' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		set rxN3Pedido = new cl_REL_COMISSAO_NFSe_N3_PEDIDO

		x = x & "	<tr nowrap class='VISAO_ANALIT'>" & chr(13)

	'	EVITA DIFERENÇAS DE ARREDONDAMENTO
		vl_preco_venda = converte_numero(formata_moeda(r("total_preco_venda")))
		vl_preco_NF = converte_numero(formata_moeda(r("total_preco_NF")))
		perc_RT = r("perc_RT")
		vl_RT = (perc_RT/100) * vl_preco_venda
		vl_RT = converte_numero(formata_moeda(vl_RT))
		vl_RA_bruto = vl_preco_NF - vl_preco_venda
		if Not calcula_total_RA_liquido(r("perc_desagio_RA_liquida"), vl_RA_bruto, vl_RA_liquido) then
			s_erro_fatal = "FALHA AO CALCULAR O RA LÍQUIDO DO PEDIDO " & Trim("" & r("pedido"))
			exit do
			end if
		
		vl_RA_liquido = converte_numero(formata_moeda(vl_RA_liquido))
		vl_RA_diferenca = vl_RA_bruto - vl_RA_liquido
		
		if (vl_preco_venda < 0) Or (vl_RT < 0) Or (vl_RA_bruto < 0) Or (vl_RA_liquido < 0) then
			s_cor = "red"
			s_cor_sinal = "red"
			s_sinal = "-"
		else
			s_cor = "black"
			s_cor_sinal = "green"
			s_sinal = "+"
			end if

		'> CHECK BOX
		'	É USADO O CÓDIGO DA OPERAÇÃO (VENDA NORMAL, DEVOLUÇÃO, PERDA) P/ NÃO CORRER O RISCO DE HAVER CONFLITO DEVIDO A ID'S REPETIDOS ENTRE AS OPERAÇÕES
		s_id_base = "comissao_paga_" & Trim("" & r("operacao")) & "_" & Trim("" & r("id_registro"))
		s_class = " CKB_COM_BL_" & idx_bloco
		s_class_td = ""
		if Trim("" & r("operacao")) = VENDA_NORMAL then
			if s_lista_completa_venda_normal <> "" then s_lista_completa_venda_normal = s_lista_completa_venda_normal & ";"
			s_lista_completa_venda_normal = s_lista_completa_venda_normal & Trim("" & r("id_registro"))
			s_class = s_class & " CKB_COM_VDNORM"
		elseif Trim("" & r("operacao")) = DEVOLUCAO then
			if s_lista_completa_devolucao <> "" then s_lista_completa_devolucao = s_lista_completa_devolucao & ";"
			s_lista_completa_devolucao = s_lista_completa_devolucao & Trim("" & r("id_registro"))
			s_class = s_class & " CKB_COM_DEV"
		elseif Trim("" & r("operacao")) = PERDA then
			if s_lista_completa_perda <> "" then s_lista_completa_perda = s_lista_completa_perda & ";"
			s_lista_completa_perda = s_lista_completa_perda & Trim("" & r("id_registro"))
			s_class = s_class & " CKB_COM_PERDA"
			end if

		s_id = "ckb_" & s_id_base
		x = x & "		<td class='MDTE tdCkb" & s_class_td & "' align='center'>" & _
							"<input type='checkbox' class='CKB_COM CKB_COM_REG " & s_class & "' name='" & s_id & "' id='" & s_id & "' value='" & Trim("" & r("id_registro")) & "|" & Trim("" & r("operacao")) & "' />" & _
							"<input type='hidden' name='" & s_id & "_original" & "' id='" & s_id & "_original" & "' value='" & Trim("" & r("status_comissao")) & "' />" & _
				"</td>" & chr(13)
	
		'> LOJA
		s_id = "spn_loja_" & s_id_base
		x = x & "		<td class='MTD tdLoja' align='center'><span class='Cnc' style='color:" & s_cor & ";' name='" & s_id & "' id='" & s_id & "'>" & Trim("" & r("loja")) & "</span>" & chr(13)
		'EMPRESA DO LANÇAMENTO NO FLUXO DE CAIXA
		s_id = "c_empresa_lancamento_" & s_id_base
		x = x & "		<input type='hidden' name='" & s_id & "' id='" & s_id & "' value='" & Trim("" & r("id_plano_contas_empresa_comissao_indicador")) & "' />" & chr(13)
		x = x & "</td>" & chr(13)

		'> Nº ORÇAMENTO
		s = Trim("" & r("orcamento"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdOrcamento' align='left'><span class='Cn'><a style='color:" & s_cor & ";' href='javascript:fORCConsulta(" & _
				chr(34) & s & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o orçamento'>" & _
				s & "</a></span></td>" & chr(13)

		'> Nº PEDIDO
		s_nome_cliente = Trim("" & r("nome_iniciais_em_maiusculas"))
		s_nome_cliente = Left(s_nome_cliente, 15)
		
		x = x & "		<td class='MTD tdPedido' align='left'><span class='Cn'><a style='color:" & s_cor & ";' href='javascript:fPEDConsulta(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o pedido'>" & _
				Trim("" & r("pedido")) & "<br />" & s_nome_cliente & "</a></span></td>" & chr(13)

		'> DATA
		s = formata_data(r("data"))
		x = x & "		<td align='center' class='MTD tdData'><span class='Cnc' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)

		'> VALOR DO PEDIDO (PREÇO DE VENDA)
		s_id = "spn_vl_preco_venda_" & s_id_base
		x = x & "		<td align='right' class='MTD tdVlPedido'><span class='Cnd' style='color:" & s_cor & ";' name='" & s_id & "' id='" & s_id & "'>" & formata_moeda(vl_preco_venda) & "</span></td>" & chr(13)

		'> PREÇO NF (CAMPO OCULTO, SOMENTE P/ CÁLCULO DE TOTAIS)
		s_id = "spn_vl_preco_NF_" & s_id_base
		x = x & "<span style='display:none;' name='" & s_id & "' id='" & s_id & "'>" & formata_moeda(vl_preco_NF) & "</span>"

		'> COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
		s_id = "spn_vl_RT_" & s_id_base
		x = x & "		<td align='right' class='MTD tdVlRT'><span class='Cnd' style='color:" & s_cor & ";' name='" & s_id & "' id='" & s_id & "'>" & formata_moeda(vl_RT) & "</span></td>" & chr(13)

		'> RA BRUTO
		s_id = "spn_vl_RA_bruto_" & s_id_base
		x = x & "		<td align='right' class='MTD tdVlRABruto'><span class='Cnd' style='color:" & s_cor & ";' name='" & s_id & "' id='" & s_id & "'>" & formata_moeda(vl_RA_bruto) & "</span></td>" & chr(13)

		'> RA LÍQUIDO
		s_id = "spn_vl_RA_liq_" & s_id_base
		x = x & "		<td align='right' class='MTD tdVlRALiq'><span class='Cnd' style='color:" & s_cor & ";' name='" & s_id & "' id='" & s_id & "'>" & formata_moeda(vl_RA_liquido) & "</span></td>" & chr(13)

		'> RA DIFERENÇA
		s_id = "spn_vl_RA_dif_" & s_id_base
		x = x & "		<td align='right' class='MTD tdVlRADif'><span class='Cnd' style='color:" & s_cor & ";' name='" & s_id & "' id='" & s_id & "'>" & formata_moeda(vl_RA_diferenca) & "</span></td>" & chr(13)

		'> STATUS DE PAGAMENTO
		s_id = "spn_st_pagto_" & s_id_base
		x = x & "		<td class='MTD tdStPagto' align='left'><span class='Cn' style='color:" & s_cor & ";' name='" & s_id & "' id='" & s_id & "'>" & x_status_pagto(Trim("" & r("st_pagto"))) & "</span></td>" & chr(13)

		'> +/-
		s_id = "spn_sinal_" & s_id_base
		x = x & "		<td align='center' class='MTD tdSinal'><span class='C' style='font-family:Courier,Arial;color:" & s_cor_sinal & "' name='" & s_id & "' id='" & s_id & "'>" & s_sinal & "</span></td>" & chr(13)
		
		'> COLUNA DA FIGURA (EXPANDE/RECOLHE)
		x = x & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13)
		
		vl_sub_total_preco_venda = vl_sub_total_preco_venda + r("total_preco_venda")
		vl_total_preco_venda = vl_total_preco_venda + r("total_preco_venda")
		vl_sub_total_preco_NF = vl_sub_total_preco_NF + r("total_preco_NF")
		vl_total_preco_NF = vl_total_preco_NF + r("total_preco_NF")
		vl_sub_total_RT = vl_sub_total_RT + vl_RT
		vl_total_RT = vl_total_RT + vl_RT
		vl_sub_total_RA_bruto = vl_sub_total_RA_bruto + vl_RA_bruto
		vl_total_RA_bruto = vl_total_RA_bruto + vl_RA_bruto
		vl_sub_total_RA_liquido = vl_sub_total_RA_liquido + vl_RA_liquido
		vl_total_RA_liquido = vl_total_RA_liquido + vl_RA_liquido
		vl_sub_total_RA_diferenca = vl_sub_total_RA_diferenca + vl_RA_diferenca
		vl_total_RA_diferenca = vl_total_RA_diferenca + vl_RA_diferenca
		
		x = x & "	</tr>" & chr(13)
		
		if (n_reg_total mod 50) = 0 then
			Response.Write x
			x = ""
			end if

		with rxN3Pedido
			.pedido = Trim("" & r("pedido"))
			.orcamento = Trim("" & r("orcamento"))
			.operacao = Trim("" & r("operacao"))
			.id_cfg_tabela_origem = r("id_cfg_tabela_origem")
			.id_registro_tabela_origem = Trim("" & r("id_registro"))
			.st_comissao_original = CLng(r("status_comissao"))
			.st_pagto = Trim("" & r("st_pagto"))
			.loja = Trim("" & r("loja"))
			.nome_cliente = Trim("" & r("nome_iniciais_em_maiusculas"))
			.data_pedido = r("data")
			.perc_RT = r("perc_RT")
			.vl_preco_venda = vl_preco_venda
			.vl_preco_NF = vl_preco_NF
			.vl_RT = vl_RT
			.vl_RA_bruto = vl_RA_bruto
			.vl_RA_liquido = vl_RA_liquido
			.vl_RA_dif = vl_RA_diferenca
			.vl_comissao = vl_RT + vl_RA_liquido
			.sinal = s_sinal
			.cor_sinal = s_cor_sinal
			.cor_linha = s_cor
			end with

		vN2(UBound(vN2)).AddN3Pedido(rxN3Pedido)

		IdIndicador_anterior = r("IdIndicador")
		r.MoveNext
		loop


  ' MOSTRA TOTAL DO ÚLTIMO INDICADOR
	if n_reg <> 0 then
		s_cor="black"
		if vl_sub_total_preco_venda < 0 then s_cor="red"
		if vl_sub_total_RT < 0 then s_cor="red"
		if vl_sub_total_RA_bruto < 0 then s_cor="red"
		if vl_sub_total_RA_liquido < 0 then s_cor="red"

		s_sql = "SELECT" & _
					" tDESC.id" & _
					", tDESC.descricao" & _
					", tDESC.valor" & _
					", tDESC.ordenacao" & _
				" FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO tDESC" & _
					" INNER JOIN t_ORCAMENTISTA_E_INDICADOR tOI ON (tDESC.apelido=tOI.apelido)" & _
				" WHERE" & _
					" (tOI.Id = " & IdIndicador_anterior & ")" & _
				" ORDER BY" & _
					" tDESC.ordenacao"
		if rs2.State <> 0 then rs2.Close
		rs2.Open s_sql, cn
		msg_desconto = ""
		if Not rs2.Eof then
			valor_desconto = 0
			contador = 0
			qtde_registro_desc = 0
			Erase v_desconto_descricao
			Erase v_desconto_valor
			do while Not rs2.EoF
				redim preserve v_desconto_descricao(contador)
				redim preserve  v_desconto_valor(contador)
				v_desconto_descricao(contador) = rs2("descricao")
				v_desconto_valor(contador) = rs2("valor")
				valor_desconto = valor_desconto + v_desconto_valor(contador)
				qtde_registro_desc = qtde_registro_desc + 1
				contador = contador + 1

				set rxN3Desconto = new cl_REL_COMISSAO_NFSe_N3_DESCONTO
				with rxN3Desconto
					.id_orcamentista_e_indicador_desconto = rs2("id")
					.descricao = Trim("" & rs2("descricao"))
					.valor = rs2("valor")
					.ordenacao = rs2("ordenacao")
					end with
				vN2(UBound(vN2)).AddN3Desconto(rxN3Desconto)

				rs2.MoveNext
				loop
			msg_desconto = "Planilha de Desconto:&nbsp; " & Cstr(qtde_registro_desc) & " registro(s) no valor total de&nbsp;" & formata_moeda(valor_desconto)
		else
			msg_desconto = ""
			end if

		x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
				"		<td colspan='5' class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>TOTAL:</span></td>" & chr(13) & _
				"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _
				"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _
				"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_bruto) & "</span></td>" & chr(13) & _
				"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _
				"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _
				"		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
				"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr>" & chr(13)

		with vN2(UBound(vN2))
			.vl_total_preco_venda = vl_sub_total_preco_venda
			.vl_total_preco_NF = vl_sub_total_preco_NF
			.vl_total_RT = vl_sub_total_RT
			.vl_total_RA_bruto = vl_sub_total_RA_bruto
			.vl_total_RA_liquido = vl_sub_total_RA_liquido
			.vl_total_RA_dif = vl_sub_total_RA_diferenca
			.cor_linha_total = s_cor
			.mensagem_desconto = msg_desconto
			end with

		x = x & "		<td align='left' colspan='10' nowrap>"
		if msg_desconto <> "" then
			x = x & "<span class='Cd'><a href='javascript:abreDesconto(" & idx_bloco & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>" & msg_desconto & "</a></span>"
		else
			x = x & "<span class='Rd' style='color: black;'></span>"
			end if
		x = x & "</td>"

		s_cor = "black"

		x = x & "		<td align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>"

		if (vl_sub_total_RT + vl_sub_total_RA_liquido) >= 0 then
			x = x & "		<td align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT + vl_sub_total_RA_liquido) & "</span></td>" & chr(13)
			end if
		x = x & "	</tr>" & chr(13)
		if msg_desconto <> "" then
			x = x & "	<tr>" & chr(13) & _
				"		<td class='VISAO_ANALIT' id='table_Desconto_" & idx_bloco & "' colspan='15'>" & chr(13) & _
				"			<table colspan='2' align='left'>" & chr(13)
			for contador = 0 to Ubound(v_desconto_descricao)
				x = x & "				<tr>" & chr(13) & _
						"					<td width='15'>&nbsp;</td>" & chr(13) & _
						"					<td align='left' width='400'><span class='Cd' style='color:red;'>" & v_desconto_descricao(contador) & "</span></td>" & chr(13) & _
						"					<td align='left'><span class='Cd' style='color:red;'> R$ " & formata_moeda(v_desconto_valor(contador)) & "</span></td>" & chr(13) & _
						"				</tr>" & chr(13)
				next
			x = x & "			</table>"& chr(13) & _
					"		</td>"& chr(13)& _
					"	</tr>"
			end if
	
	'>	TOTAL GERAL
		if qtde_indicadores >= 1 then
			x = x & "	<tr>" & chr(13) & _
					"		<td colspan='12' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<tr>" & chr(13) & _
					"		<td colspan='12' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<tr nowrap style='background:whitesmoke'>" & chr(13) & _
					"		<td class='MTBE' colspan='5' align='right' nowrap><span class='Cd' style='color:black;'>TOTAL GERAL:</span></td>" & chr(13)
			
			if vl_total_preco_venda < 0 then s_cor="red" else s_cor="black"
			x = x & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_preco_venda) & "</span></td>" & chr(13)
			if vl_total_RT < 0 then s_cor="red" else s_cor="black"
			x = x & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13)
			if vl_total_RA_bruto < 0 then s_cor="red" else s_cor="black"
			x = x & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_bruto) & "</span></td>" & chr(13)
			if vl_total_RA_liquido < 0 then s_cor="red" else s_cor="black"
			x = x & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13)
			if vl_total_RA_diferenca < 0 then s_cor="red" else s_cor="black"
			x = x & _
					"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_diferenca) & "</span></td>" & chr(13)
			x = x & _
					"		<td class='MTBD' colspan='2' align='right'><span class='Cd' style='color:black;'>&nbsp;</span></td>" & chr(13) & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<tr>" & chr(13) & _
					"		<td colspan='12' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<tr nowrap style='background:honeydew'>" & chr(13) & _
					"		<td class='MTBE' colspan='5' align='right' nowrap><span class='Cd SpnRowTotalGeralComissao' style='color:black;'>TOTAL COMISSÃO:</span></td>" & chr(13)
			if vl_total_RT < 0 then s_cor="red" else s_cor="black"
			x = x & _
					"		<td class='MTB' colspan='2' align='right'>" & _
								"<span class='Cd SpnRowTotalGeralComissao' style='color:black;'>COM: &nbsp; </span>" & _
								"<span name='spnTotalGeralRT SpnRowTotalGeralComissao' id='spnTotalGeralRT' class='Cd SpnRowTotalGeralComissao' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span>" & _
							"</td>" & chr(13)
			if vl_total_RA_liquido < 0 then s_cor="red" else s_cor="black"
			x = x & _
					"		<td class='MTB' colspan='2' align='right'>" & _
								"<span class='Cd SpnRowTotalGeralComissao' style='color:black;'>RA: &nbsp; </span>" & _
								"<span name='spnTotalGeralRALiq' id='spnTotalGeralRALiq' class='Cd SpnRowTotalGeralComissao' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span>" & _
							"</td>" & chr(13)
			if (vl_total_RT+vl_total_RA_liquido) < 0 then s_cor="red" else s_cor="black"
			x = x & _
					"		<td class='MTBD' colspan='3' align='right'>" & _
								"<span class='Cd SpnRowTotalGeralComissao' style='color:black;'> = &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>" & _
								"<span name='spnTotalGeral_RT_RALiq' id='spnTotalGeral_RT_RALiq' class='Cd SpnRowTotalGeralComissao' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT+vl_total_RA_liquido) & "</span>" & _
							"</td>" & chr(13)
			x = x & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13)
			end if 'if qtde_indicadores >= 1
		end if 'if n_reg <> 0


  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<tr nowrap>" & chr(13) & _
				"		<td class='MT ALERTA' colspan='12' align='center'><span class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS NO PERÍODO ESPECIFICADO&nbsp;</span></td>" & chr(13) & _
				"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO INDICADOR
	x = x & "</table>" & chr(13) & chr(13)
	
	Response.write x

	x = "<br />" & chr(13) & _
		"<input type='hidden' name='c_lista_completa_venda_normal' id='c_lista_completa_venda_normal' value='" & s_lista_completa_venda_normal & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_completa_devolucao' id ='c_lista_completa_devolucao' value='" & s_lista_completa_devolucao & "' />" & chr(13) & _
		"<input type='hidden' name='c_lista_completa_perda' id='c_lista_completa_perda' value='" & s_lista_completa_perda & "' />" & chr(13)
	Response.write x


	s_erro_fatal = ""

'	~~~~~~~~~~~~~
	cn.BeginTrans
'	~~~~~~~~~~~~~
	if Not cria_recordset_otimista(tN1, msg_erro) then
		s_erro_fatal = "FALHA AO TENTAR CRIAR UM OBJETO ADO PARA ACESSO AO BANCO DE DADOS PARA A TABELA: t_COMISSAO_INDICADOR_NFSe_N1"
		end if

	if s_erro_fatal = "" then
		if Not cria_recordset_otimista(tN2, msg_erro) then
			s_erro_fatal = "FALHA AO TENTAR CRIAR UM OBJETO ADO PARA ACESSO AO BANCO DE DADOS PARA A TABELA: t_COMISSAO_INDICADOR_NFSe_N2"
			end if
		end if

	if s_erro_fatal = "" then
		if Not cria_recordset_otimista(tN3Desc, msg_erro) then
			s_erro_fatal = "FALHA AO TENTAR CRIAR UM OBJETO ADO PARA ACESSO AO BANCO DE DADOS PARA A TABELA: t_COMISSAO_INDICADOR_NFSe_N3_DESCONTO"
			end if
		end if

	if s_erro_fatal = "" then
		if Not cria_recordset_otimista(tN3Ped, msg_erro) then
			s_erro_fatal = "FALHA AO TENTAR CRIAR UM OBJETO ADO PARA ACESSO AO BANCO DE DADOS PARA A TABELA: t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO"
			end if
		end if

	if s_erro_fatal = "" then
		s_sql = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N1 WHERE (id = -1)"
		tN1.open s_sql, cn
		tN1.AddNew
		tN1("usuario_cadastro") = usuario
		tN1("competencia_ano") = Year(dt_entregue_termino)
		tN1("competencia_mes") = Month(dt_entregue_termino)
		tN1("NFSe_cnpj") = c_cnpj_nfse
		tN1("vl_total_geral_preco_venda") = vl_total_preco_venda
		tN1("vl_total_geral_preco_NF") = vl_total_preco_NF
		tN1("vl_total_geral_RT") = vl_total_RT
		tN1("vl_total_geral_RA_bruto") = vl_total_RA_bruto
		tN1("vl_total_geral_RA_liquido") = vl_total_RA_liquido
		tN1("vl_total_geral_RA_dif") = vl_total_RA_diferenca
		tN1("cor_linha_total_geral") = s_cor
		tN1.Update
		lngNsuN1 = tN1("id")
		
		for iN2=LBound(vN2) to UBound(vN2)
			if Trim("" & vN2(iN2).indicador) <> "" then
				with vN2(iN2)
					s_sql = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N2 WHERE (id = -1)"
					if tN2.State <> 0 then tN2.Close
					tN2.open s_sql, cn
					tN2.AddNew
					tN2("id_comissao_indicador_nfse_n1") = lngNsuN1
					tN2("vendedor") = .vendedor
					tN2("indicador") = .indicador
					tN2("id_indicador") = .id_indicador
					tN2("desempenho_nota") = .desempenho_nota
					tN2("razao_social_nome") = .razao_social_nome
					tN2("NFSe_razao_social") = .NFSe_razao_social
					tN2("comissao_cartao_status") = .comissao_cartao_status
					tN2("comissao_cartao_cpf") = .comissao_cartao_cpf
					tN2("comissao_cartao_titular") = .comissao_cartao_titular
					tN2("banco") = .banco
					tN2("banco_nome") = .banco_nome
					tN2("agencia") = .agencia
					tN2("agencia_dv") = .agencia_dv
					tN2("tipo_conta") = .tipo_conta
					tN2("conta_operacao") = .conta_operacao
					tN2("conta") = .conta
					tN2("conta_dv") = .conta_dv
					tN2("favorecido") = .favorecido
					tN2("favorecido_cnpj_cpf") = .favorecido_cnpj_cpf
					tN2("vl_total_preco_venda") = .vl_total_preco_venda
					tN2("vl_total_preco_NF") = .vl_total_preco_NF
					tN2("vl_total_RT") = .vl_total_RT
					tN2("vl_total_RA_bruto") = .vl_total_RA_bruto
					tN2("vl_total_RA_liquido") = .vl_total_RA_liquido
					tN2("vl_total_RA_dif") = .vl_total_RA_dif
					tN2("cor_linha_total") = .cor_linha_total
					tN2("mensagem_desconto") = .mensagem_desconto
					tN2.Update
					lngNsuN2 = tN2("id")
					end with

				for iN3Desc=LBound(vN2(iN2).vN3Desconto) to UBound(vN2(iN2).vN3Desconto)
					if Trim("" & vN2(iN2).vN3Desconto(iN3Desc).id_orcamentista_e_indicador_desconto) <> "" then
						with vN2(iN2).vN3Desconto(iN3Desc)
							s_sql = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N3_DESCONTO WHERE (id = -1)"
							if tN3Desc.State <> 0 then tN3Desc.Close
							tN3Desc.open s_sql, cn
							tN3Desc.AddNew
							tN3Desc("id_comissao_indicador_nfse_n2") = lngNsuN2
							tN3Desc("id_orcamentista_e_indicador_desconto") = .id_orcamentista_e_indicador_desconto
							tN3Desc("descricao") = .descricao
							tN3Desc("valor") = .valor
							tN3Desc("ordenacao") = .ordenacao
							tN3Desc.Update
							end with
						end if
					next

				for iN3Ped=LBound(vN2(iN2).vN3Pedido) to UBound(vN2(iN2).vN3Pedido)
					if Trim("" & vN2(iN2).vN3Pedido(iN3Ped).pedido) <> "" then
						with vN2(iN2).vN3Pedido(iN3Ped)
							s_sql = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO WHERE (id = -1)"
							if tN3Ped.State <> 0 then tN3Ped.Close
							tN3Ped.open s_sql, cn
							tN3Ped.AddNew
							tN3Ped("id_comissao_indicador_nfse_n2") = lngNsuN2
							tN3Ped("pedido") = .pedido
							tN3Ped("operacao") = .operacao
							tN3Ped("id_cfg_tabela_origem") = .id_cfg_tabela_origem
							tN3Ped("id_registro_tabela_origem") = .id_registro_tabela_origem
							tN3Ped("loja") = .loja
							tN3Ped("st_comissao_original") = .st_comissao_original
							tN3Ped("st_pagto") = .st_pagto
							tN3Ped("nome_cliente") = .nome_cliente
							tN3Ped("data_pedido") = .data_pedido
							tN3Ped("perc_RT") = .perc_RT
							tN3Ped("vl_preco_venda") = .vl_preco_venda
							tN3Ped("vl_preco_NF") = .vl_preco_NF
							tN3Ped("vl_RT") = .vl_RT
							tN3Ped("vl_RA_bruto") = .vl_RA_bruto
							tN3Ped("vl_RA_liquido") = .vl_RA_liquido
							tN3Ped("vl_RA_dif") = .vl_RA_dif
							tN3Ped("vl_comissao") = .vl_comissao
							tN3Ped("sinal") = .sinal
							tN3Ped("cor_sinal") = .cor_sinal
							tN3Ped("cor_linha") = .cor_linha
							tN3Ped.Update
							end with
						end if
					next
				end if
			next
		end if 'if s_erro_fatal = ""

	if s_erro_fatal <> "" then
	'	~~~~~~~~~~~~~~~~
		cn.RollbackTrans
	'	~~~~~~~~~~~~~~~~
		blnErroFatal = True
		s = monta_html_bloco_msg_erro(s_erro_fatal)
		Response.Write s
		exit sub
		end if
	
	if s_erro_fatal = "" then
		if s_log <> "" then
			s_log = "t_COMISSAO_INDICADOR_NFSe_N1.id=" & Cstr(lngNsuN1) & ";  " & s_log
			grava_log usuario, "", "", "", OP_LOG_REL_COMISSAO_INDICADORES_NFSe_CONSULTA, s_log
			end if
	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
		end if

	if s_erro_fatal = "" then
		x = "<input type='hidden' name='id_nsu_N1' id='id_nsu_N1' value='" & CStr(lngNsuN1) & "' />" & chr(13)
		Response.Write x
		end if

	if tN1.State <> 0 then tN1.Close
	if tN2.State <> 0 then tN2.Close
	if tN3Ped.State <> 0 then tN3Ped.Close
	if tN3Desc.State <> 0 then tN3Desc.Close
	set tN1=nothing
	set tN2=nothing
	set tN3Ped=nothing
	set tN3Desc=nothing

	if r.State <> 0 then r.Close
	if rs2.State <> 0 then rs2.Close
	set r=nothing
	set rs2=nothing
end sub

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
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	var windowScrollTopAnterior;
	window.status = 'Aguarde, executando a consulta ...';

$(function() {
	$("#divPedidoConsulta").hide();

	sizeDivPedidoConsulta();

	$('#divInternoPedidoConsulta').addClass('divFixo');

	$(document).keyup(function(e) {
		if (e.keyCode == 27) fechaDivPedidoConsulta();
	});

	$("#divPedidoConsulta").click(function() {
		fechaDivPedidoConsulta();
	});

	$("#imgFechaDivPedidoConsulta").click(function() {
		fechaDivPedidoConsulta();
	});

	// EXIBE O REALCE NOS CHECKBOXES QUE SÃO EXIBIDOS INICIALMENTE ASSINALADOS
	$(".CKB_COM:enabled").each(function() {
		if ($(this).is(":checked")) {
			$(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
		}
		else {
			$(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
		}
	});

	// EVENTO P/ REALÇAR OU NÃO CONFORME SE MARCA/DESMARCA O CHECKBOX
	$(".CKB_COM:enabled").click(function () {
		if ($(this).is(":checked")) {
			$(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
		}
		else {
			$(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
		}
		recalculaComissaoTotalGeral();
	});

	// VISÃO SINTÉTICA?
	if ($("#rb_visao").val() == "SINTETICA") {
		$(".CKB_COM").attr("disabled", true);
	}

	recalculaComissaoTotalGeral();
});

	function recalculaComissaoTotalGeral() {
		var vlTotalPrecoVenda = 0;
		var vlTotalPrecoNF = 0;
		var vlTotalRT = 0;
		var vlTotalRABruto = 0;
		var vlTotalRALiq = 0;
		var vlTotalRADif = 0;
		var s_id, s_id_base, s_value;
		$(".CKB_COM_REG:enabled:checked").each(function () {
			s_id = $(this).attr("id");
			s_id_base = s_id.replace("ckb_", "");
			// Preço venda
			s_id = "spn_vl_preco_venda_" + s_id_base;
			s_value = $("#" + s_id).text();
			vlTotalPrecoVenda += converte_numero(s_value);
			// Preço NF
			s_id = "spn_vl_preco_NF_" + s_id_base;
			s_value = $("#" + s_id).text();
			vlTotalPrecoNF += converte_numero(s_value);
			// RT
			s_id = "spn_vl_RT_" + s_id_base;
			s_value = $("#" + s_id).text();
			vlTotalRT += converte_numero(s_value);
			// RA Bruto
			s_id = "spn_vl_RA_bruto_" + s_id_base;
			s_value = $("#" + s_id).text();
			vlTotalRABruto += converte_numero(s_value);
			// RA Líquido
			s_id = "spn_vl_RA_liq_" + s_id_base;
			s_value = $("#" + s_id).text();
			vlTotalRALiq += converte_numero(s_value);
			// RA Dif
			s_id = "spn_vl_RA_dif_" + s_id_base;
			s_value = $("#" + s_id).text();
			vlTotalRADif += converte_numero(s_value);
		});
		// Atualiza valores na linha "TOTAL COMISSÃO"
		s_id = "#spnTotalGeralRT";
		$(s_id).text(formata_moeda(vlTotalRT));
		$(s_id).css("color", (vlTotalRT < 0 ? "red" : "black"));
		s_id = "#spnTotalGeralRALiq";
		$(s_id).text(formata_moeda(vlTotalRALiq));
		$(s_id).css("color", (vlTotalRALiq < 0 ? "red" : "black"));
		s_id = "#spnTotalGeral_RT_RALiq";
		$(s_id).text(formata_moeda(vlTotalRT + vlTotalRALiq));
		$(s_id).css("color", ((vlTotalRT + vlTotalRALiq) < 0 ? "red" : "black"));
		// Atualiza campos hidden que armazenam os totais dos registros selecionados
		s_id = "#c_total_geral_selecionado_preco_venda";
		$(s_id).val(formata_moeda(vlTotalPrecoVenda));
		s_id = "#c_total_geral_selecionado_preco_NF";
		$(s_id).val(formata_moeda(vlTotalPrecoNF));
		s_id = "#c_total_geral_selecionado_RT";
		$(s_id).val(formata_moeda(vlTotalRT));
		s_id = "#c_total_geral_selecionado_RA_bruto";
		$(s_id).val(formata_moeda(vlTotalRABruto));
		s_id = "#c_total_geral_selecionado_RA_liq";
		$(s_id).val(formata_moeda(vlTotalRALiq));
		s_id = "#c_total_geral_selecionado_RA_dif";
		$(s_id).val(formata_moeda(vlTotalRADif));
	}

function abreDesconto(idx_bloco) {
	var s_seletor = "#table_Desconto_" + idx_bloco;
	$(s_seletor).toggle();
	}

//Every resize of window
$(window).resize(function() {
	sizeDivPedidoConsulta();
});

function sizeDivPedidoConsulta() {
	var newHeight = $(document).height() + "px";
	$("#divPedidoConsulta").css("height", newHeight);
}

function fechaDivPedidoConsulta() {
	$(window).scrollTop(windowScrollTopAnterior);
	$("#divPedidoConsulta").fadeOut();
	$("#iframePedidoConsulta").attr("src", "");
}

function marcar_todos() {
	$(".CKB_COM:enabled")
		.prop("checked", true)
		.parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
	recalculaComissaoTotalGeral();
}

function desmarcar_todos() {
	$(".CKB_COM:enabled")
		.prop("checked", false)
		.parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
	recalculaComissaoTotalGeral();
}

function trata_ckb_onclick(idx_bloco) {
var s_id, s_class;
	s_id = "#ckb_comissao_paga_tit_bloco_" + idx_bloco;
	s_class = ".CKB_COM_BL_" + idx_bloco;
	if ($(s_id).is(":checked")) {
		$(s_class).prop("checked", true);
		$(s_class).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
	}
	else {
		$(s_class).prop("checked", false);
		$(s_class).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
	}
}

function fExibeOcultaCampos(indice_bloco) {
	var s_seletor = "#tableDados_" + indice_bloco + " .VISAO_ANALIT";
	$(s_seletor).toggle();
}

function expandir_todos() {
	$(".VISAO_ANALIT").show();
}

function recolher_todos() {
	$(".VISAO_ANALIT").hide();
}

function fPEDConsulta(id_pedido, usuario) {
	windowScrollTopAnterior = $(window).scrollTop();
	sizeDivPedidoConsulta();
	$("#iframePedidoConsulta").attr("src", "PedidoConsultaView.asp?pedido_selecionado=" + id_pedido + "&pedido_selecionado_inicial=" + id_pedido + "&usuario=" + usuario);
	$("#divPedidoConsulta").fadeIn();
}

function fORCConsulta(id_orcamento, usuario) {
	windowScrollTopAnterior = $(window).scrollTop();
	sizeDivPedidoConsulta();
	$("#iframePedidoConsulta").attr("src", "OrcamentoConsultaView.asp?orcamento_selecionado=" + id_orcamento + "&orcamento_selecionado_inicial=" + id_orcamento + "&usuario=" + usuario);
	$("#divPedidoConsulta").fadeIn();
}

function fRetornaInicio(f) {
	f.action = "RelComissaoIndicadoresNFSeP01Filtro.asp";
	f.submit();
}

function fRelSumario(f, id_nsu_N1) {
	f.id_nsu_N1.value = id_nsu_N1;
	f.action = "RelComissaoIndicadoresNFSeP06BotaoMagico.asp";
	f.submit();
}

function fRELGravaDados(f) {
	var vl_total_RT_RALiq;
	var s_confirmacao;
	var v, s_id, s_id_base, s_value, s_search, s_lista_loja, s_lista_empresa, qtde_lojas, qtde_empresas, s_lojas, s_empresas;

	vl_total_RT_RALiq = converte_numero($("#c_total_geral_selecionado_RT").val()) + converte_numero($("#c_total_geral_selecionado_RA_liq").val());

	if (vl_total_RT_RALiq == 0) {
		alert("O valor total da comissão (RT + RA líquido) é zero!");
		return;
	}

	if (vl_total_RT_RALiq < 0) {
		alert("O valor total da comissão (RT + RA líquido) é negativo!");
		return;
	}

	// Verifica se há mais de uma loja entre os pedidos/devoluções/perdas selecionados e se há mais de uma empresa p/ o lançamento no fluxo de caixa
	s_lista_loja = "|";
	s_lista_empresa = "|";
	qtde_lojas = 0;
	qtde_empresas = 0;
	$(".CKB_COM_REG:enabled:checked").each(function () {
		s_id = $(this).attr("id");
		s_id_base = s_id.replace("ckb_", "");
		// Loja
		s_id = "spn_loja_" + s_id_base;
		s_value = $("#" + s_id).text();
		if (s_value.length > 0) {
			s_search = "|" + s_value + "|";
			if (s_lista_loja.indexOf(s_search) == -1) {
				qtde_lojas++;
				s_lista_loja += s_value + "|";
			}
		}
		// Empresa do lançamento no fluxo de caixa
		s_id = "c_empresa_lancamento_" + s_id_base;
		s_value = $("#" + s_id).val();
		if (s_value.length > 0) {
			s_search = "|" + s_value + "|";
			if (s_lista_empresa.indexOf(s_search) == -1) {
				qtde_empresas++;
				s_lista_empresa += s_value + "|";
			}
		}
	});

	s_confirmacao = "Prosseguir com a gravação dos dados?";
	if (qtde_lojas > 1) {
		v = s_lista_loja.split("|");
		s_lojas = "";
		for (var i = 0; i < v.length; i++) {
			if (("" + v[i]).length > 0) {
				if (s_lojas.length > 0) s_lojas += ", ";
				s_lojas += v[i];
			}
		}
		s_confirmacao += "\n\nAtenção: há mais de uma loja entre os registros selecionados (lojas: " + s_lojas + ")!";
	}

	if (qtde_empresas > 1) {
		v = s_lista_empresa.split("|");
		s_empresas = "";
		for (var i = 0; i < v.length; i++) {
			if (("" + v[i]).length > 0) {
				if (s_empresas.length > 0) s_empresas += ", ";
				s_empresas += v[i];
			}
		}
		s_confirmacao += "\n\nAtenção: há mais de uma empresa para o lançamento no fluxo de caixa entre os registros selecionados (empresas: " + s_empresas + ")!";
	}

	if (!confirm(s_confirmacao)) {
		return;
	}

	window.status = "Aguarde ...";
	dCONFIRMA.style.visibility = "hidden";
	f.action = "RelComissaoIndicadoresNFSeP05GravaDados.asp";
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
.tdCkb
{
	width: 20px;
}
.tdLoja{
	width: 28px;
	}
.tdOrcamento{
	width: 58px;
	}
.tdPedido{
	width: 83px;
	}
.tdData{
	width: 60px;
	}
.tdVlPedido{
	width: 70px;
	}
.tdVlRT{
	width: 60px;
	}
.tdVlRABruto{
	width: 60px;
	}
.tdVlRALiq{
	width: 60px;
	}
.tdVlRADif{
	width: 60px;
	}
.tdStPagto{
	width: 60px;
	}
.tdSinal{
	width: 18px;
	}
.BTN_LNK
{
	min-width:140px;
}
.CKB_HIGHLIGHT
{
	background-color:#90EE90;
}
#divPedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsulta
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsulta.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivPedidoConsulta
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframePedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
.BkgWhite
{
	background-color:#ffffff;
}
.VISAO_ANALIT
{
	<%if blnVisaoSintetica then Response.Write "display:none;"%>
}
</style>


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
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

<form name="fSumario" id="fSumario" method="post">
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="ckb_id_indicador" id="ckb_id_indicador" value="<%=ckb_id_indicador%>" />
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>">
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>">
<input type="hidden" name="id_nsu_N1" id="id_nsu_N1" />
<input type="hidden" name="proc_comissao_request_guid" id="proc_comissao_request_guid" value="<%=proc_comissao_request_guid%>" />
</form>

</center>
</body>



<% else
     %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<input type="hidden" name="orcamento_selecionado" id="orcamento_selecionado" value="">

<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="ckb_id_indicador" id="ckb_id_indicador" value="<%=ckb_id_indicador%>" />
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>">
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>">
<input type="hidden" name="c_usuario_sessao" id="c_usuario_sessao" value="<%=usuario%>" />
<input type="hidden" name="c_total_geral_selecionado_preco_venda" id="c_total_geral_selecionado_preco_venda" />
<input type="hidden" name="c_total_geral_selecionado_preco_NF" id="c_total_geral_selecionado_preco_NF" />
<input type="hidden" name="c_total_geral_selecionado_RT" id="c_total_geral_selecionado_RT" />
<input type="hidden" name="c_total_geral_selecionado_RA_bruto" id="c_total_geral_selecionado_RA_bruto" />
<input type="hidden" name="c_total_geral_selecionado_RA_liq" id="c_total_geral_selecionado_RA_liq" />
<input type="hidden" name="c_total_geral_selecionado_RA_dif" id="c_total_geral_selecionado_RA_dif" />
<input type="hidden" name="proc_comissao_request_guid" id="proc_comissao_request_guid" value="<%=proc_comissao_request_guid%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (via NFS-e)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<%
	s_filtro = "<table width='709' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	MÊS DE COMPETÊNCIA
	s = normaliza_a_esq(Cstr(Month(dt_entregue_termino)),2) & " / " & Cstr(Year(dt_entregue_termino))
	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Competência:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	CNPJ EMITENTE NFS-e
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>CNPJ NFS-e:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & cnpj_cpf_formata(c_cnpj_nfse) & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	INDICADOR(ES) SELECIONADO(S)
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Indicador(es):&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'>" & s_filtro_indicador & "</td>" & chr(13) & _
				"	</tr>" & chr(13)

'	EMISSÃO
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora_sem_seg(Now) & "</span></td></tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br />
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td class="Rc" align="left">&nbsp;</td>
</tr>
</table>

<% if blnVisaoSintetica then %>
<table class="notPrint" width="709" cellpadding="0" cellspacing="0" style="margin-top:5px;">
<tr>
	<td align="right">
		<button type="button" name="bExpandirTodos" id="bExpandirTodos" class="Button BTN_LNK" onclick="expandir_todos();" title="expandir todas as linhas de dados" style="margin-left:6px;margin-bottom:2px">Expandir Tudo</button>
		&nbsp;
		<button type="button" name="bRecolherTodos" id="bRecolherTodos" class="Button BTN_LNK" onclick="recolher_todos();" title="recolher todas as linhas de dados" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Recolher Tudo</button>
	</td>
</tr>
</table>
<br />
<table class="notPrint" width="709" cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="A1" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>
<% else %>
<table class="notPrint" width="709" cellpadding="0" cellspacing="0" style="margin-top:5px;">
<tr>
	<td align="right">
		<button type="button" name="bExpandirTodos" id="bExpandirTodos" class="Button BTN_LNK" onclick="expandir_todos();" title="expandir todas as linhas de dados" style="margin-left:6px;margin-bottom:2px">Expandir Tudo</button>
		&nbsp;
		<button type="button" name="bRecolherTodos" id="bRecolherTodos" class="Button BTN_LNK" onclick="recolher_todos();" title="recolher todas as linhas de dados" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Recolher Tudo</button>
		&nbsp;
		<button type="button" name="bMarcarTodos" id="bMarcarTodos" class="Button BTN_LNK" onclick="marcar_todos();" title="assinala todos os pedidos para gravar o status da comissão como paga" style="margin-left:6px;margin-bottom:2px">Marcar todos</button>
		&nbsp;
		<button type="button" name="bDesmarcarTodos" id="bDesmarcarTodos" class="Button BTN_LNK" onclick="desmarcar_todos();" title="desmarca todos os pedidos para gravar o status da comissão como não-paga" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Desmarcar todos</button>
	</td>
</tr>
</table>

<br />
<table class="notPrint" width="709" cellspacing="0">
<tr>
	<td align="left"><a name="bANTERIOR" id="bANTERIOR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fRetornaInicio(fREL)" title="retorna para o início do relatório">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<% if Not blnErroFatal then %>
	<td align="right">
		<div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
	<% end if %>
</tr>
</table>
<% end if %>

</form>

</center>

<div id="divPedidoConsulta"><center><div id="divInternoPedidoConsulta"><img id="imgFechaDivPedidoConsulta" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframePedidoConsulta"></iframe></div></center></div>

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
