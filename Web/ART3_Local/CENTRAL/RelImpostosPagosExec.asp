<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L I M P O S T O S P A G O S E X E C . A S P
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

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_EXTENDED_TIMEOUT_EM_SEG
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs,msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_CONTROLE_IMPOSTOS, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim url_back, strUrlBotaoVoltar
	url_back = Trim(Request("url_back"))
	if url_back <> "" then
		strUrlBotaoVoltar = "RelImpostosPagosFiltro.asp?url_back=X&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	else
		strUrlBotaoVoltar = "javascript:history.back()"
		end if
	
	dim alerta
	dim s, s_aux, s_filtro
	dim c_transportadora, s_nome_transportadora, c_dt_coleta_inicio, c_dt_coleta_termino, c_uf, c_nfe_emitente
	dim ckb_pedidos_cancelados, ckb_pedidos_com_devolucao

	alerta = ""

	c_dt_coleta_inicio = Trim(Request.Form("c_dt_coleta_inicio"))
	c_dt_coleta_termino = Trim(Request.Form("c_dt_coleta_termino"))
	c_transportadora = Trim(Request.Form("c_transportadora"))
	c_uf = Trim(Request.Form("c_uf"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
	ckb_pedidos_cancelados = Trim(Request.Form("ckb_pedidos_cancelados"))
	ckb_pedidos_com_devolucao = Trim(Request.Form("ckb_pedidos_com_devolucao"))
	
	s_nome_transportadora = ""
	if c_transportadora <> "" then s_nome_transportadora = x_transportadora(c_transportadora)

	dim rNfeEmitente
	set rNfeEmitente = le_nfe_emitente(c_nfe_emitente)

	dim qtde_notas
	qtde_notas = 0

	if alerta = "" then
		if c_nfe_emitente = "" then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Não foi informado o CD"
		elseif converte_numero(c_nfe_emitente) = 0 then
			alerta=texto_add_br(alerta)
			alerta = alerta & "É necessário definir um CD válido"
			end if
		end if






' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
Const COD_COR_NAO_DEFINIDO = 0
Const COD_COR_PRETO = 1
Const COD_COR_AZUL = 2
Const COD_COR_VERMELHO = 3
dim r
dim blnDisabled
dim s, s_sql, s_select_devolucao_grupo_1, s_select_devolucao_grupo_2, cab_table, cab, n_reg_geral, n_reg_uf, s_num_nfe, s_serie_nfe, s_link_nfe, s_row, s_html_color, s_link_habilita_nfe
dim s_pedido, s_transportadora, s_data_entrega_yyyymmdd
dim s_where, s_where_data, s_where_pedido_grupo_1, s_where_pedido_grupo_2, s_from
dim i, intCodCor, intOrdenacaoCor
dim blnRegistroOk
dim x
dim total_fcp_uf,total_icms_origem_uf,total_icms_destino_uf, valor
dim total_fcp_geral,total_icms_origem_geral,total_icms_destino_geral
dim total_fcp_uf_proporcional, total_icms_origem_uf_proporcional, total_icms_destino_uf_proporcional
dim total_fcp_geral_proporcional, total_icms_origem_geral_proporcional, total_icms_destino_geral_proporcional
dim vl_fcp_uf_proporcional, vl_icms_origem_uf_proporcional, vl_icms_destino_uf_proporcional
dim s_uf, s_uf_anterior
dim ChaveAcesso
dim percProporcao, colSpan

total_fcp_uf = 0
total_icms_origem_uf = 0
total_icms_destino_uf = 0
total_fcp_geral = 0
total_icms_origem_geral = 0
total_icms_destino_geral = 0
total_fcp_uf_proporcional = 0
total_icms_origem_uf_proporcional = 0
total_icms_destino_uf_proporcional = 0
total_fcp_geral_proporcional = 0
total_icms_origem_geral_proporcional = 0
total_icms_destino_geral_proporcional = 0
vl_fcp_uf_proporcional = 0
vl_icms_origem_uf_proporcional = 0
vl_icms_destino_uf_proporcional = 0

'	MONTA CLÁUSULA WHERE
	s_where = " AND (t_NFE_IMAGEM.ide__idDest = '2') " & _
			" AND (t_NFE_IMAGEM.ide__tpNF = '1') " & _
			" AND (t_NFE_EMISSAO.st_anulado = 0) " & _
			" AND (t_NFE_EMISSAO.codigo_retorno_NFe_T1 = 1) " & _
			" AND (t_NFE_EMISSAO.controle_impostos_status = " & COD_CONTROLE_IMPOSTOS_STATUS__OK & ")"

				
'	CRITÉRIO: TRANSPORTADORA
	if c_transportadora <> "" then
		s = " (transportadora_id = '" & c_transportadora & "')"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & s
		end if
	
'	CRITÉRIO: PERÍODO
	s_where_data = ""
	if c_dt_coleta_inicio <> "" then
		s = " (a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_coleta_inicio)) & ")"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & s
		s_where_data = s
		end if
	if c_dt_coleta_termino <> "" then
		s = " (a_entregar_data_marcada <= " & bd_formata_data(StrToDate(c_dt_coleta_termino)) & ")"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & s
		if s_where_data <> "" then s_where_data = s_where_data & " AND"
		s_where_data = s_where_data & s
		end if

'	CRITÉRIO: UF
	if c_uf <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_NFE_IMAGEM.dest__UF = '" & c_uf & "')"
		end if
	
'	OWNER DO PEDIDO
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (t_NFE_EMISSAO.id_nfe_emitente = " & rNfeEmitente.id & ")"

'	PEDIDOS CANCELADOS OU PEDIDOS COM DEVOLUÇÃO
	s_where_pedido_grupo_1 = ""
	s_where_pedido_grupo_2 = ""
	if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then
		if ckb_pedidos_cancelados <> "" then
			if s_where_pedido_grupo_1 <> "" then s_where_pedido_grupo_1 = s_where_pedido_grupo_1 & " OR"
			s_where_pedido_grupo_1 = s_where_pedido_grupo_1 & " (t_PEDIDO.st_entrega = '" & ST_ENTREGA_CANCELADO & "')"

			if s_where_pedido_grupo_2 <> "" then s_where_pedido_grupo_2 = s_where_pedido_grupo_2 & " OR"
			s_where_pedido_grupo_2 = s_where_pedido_grupo_2 & " (pedidos_nf_ok.st_entrega = '" & ST_ENTREGA_CANCELADO & "')"
			end if
		
		if ckb_pedidos_com_devolucao <> "" then
			if s_where_pedido_grupo_1 <> "" then s_where_pedido_grupo_1 = s_where_pedido_grupo_1 & " OR"
			s_where_pedido_grupo_1 = s_where_pedido_grupo_1 & " (t_PEDIDO.pedido IN (SELECT DISTINCT pedido FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (t_PEDIDO_ITEM_DEVOLVIDO.pedido = t_PEDIDO.pedido)))"

			if s_where_pedido_grupo_2 <> "" then s_where_pedido_grupo_2 = s_where_pedido_grupo_2 & " OR"
			s_where_pedido_grupo_2 = s_where_pedido_grupo_2 & " (pedidos_nf_ok.pedido IN (SELECT DISTINCT pedido FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (t_PEDIDO_ITEM_DEVOLVIDO.pedido = pedidos_nf_ok.pedido)))"
			end if
		end if

	if s_where_pedido_grupo_1 <> "" then
		s_where_pedido_grupo_1 = " AND (" & s_where_pedido_grupo_1 & ")"
		end if

	if s_where_pedido_grupo_2 <> "" then
		s_where_pedido_grupo_2 = " AND (" & s_where_pedido_grupo_2 & ")"
		end if

	s_select_devolucao_grupo_1 = ""
	s_select_devolucao_grupo_2 = ""
	if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then
		s_select_devolucao_grupo_1 = ", (SELECT Coalesce(Sum(qtde),0) FROM t_PEDIDO_ITEM WHERE (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)) AS qtde_itens_pedido," & _
								" (SELECT Coalesce(Sum(qtde),0) FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)) AS qtde_itens_devolucao," & _
								" (SELECT Coalesce(Sum(qtde*preco_NF),0) FROM t_PEDIDO_ITEM WHERE (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)) AS vl_pedido_NF, " & _
								" (SELECT Coalesce(Sum(qtde*preco_NF),0) FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)) AS vl_devolucao_NF"

		s_select_devolucao_grupo_2 = ", (SELECT Coalesce(Sum(qtde),0) FROM t_PEDIDO_ITEM WHERE (t_PEDIDO_ITEM.pedido=pedidos_nf_ok.pedido)) AS qtde_itens_pedido," & _
								" (SELECT Coalesce(Sum(qtde),0) FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (t_PEDIDO_ITEM_DEVOLVIDO.pedido=pedidos_nf_ok.pedido)) AS qtde_itens_devolucao," & _
								" (SELECT Coalesce(Sum(qtde*preco_NF),0) FROM t_PEDIDO_ITEM WHERE (t_PEDIDO_ITEM.pedido=pedidos_nf_ok.pedido)) AS vl_pedido_NF, " & _
								" (SELECT Coalesce(Sum(qtde*preco_NF),0) FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (t_PEDIDO_ITEM_DEVOLVIDO.pedido=pedidos_nf_ok.pedido)) AS vl_devolucao_NF"
		end if

'	Primeiro grupo selecionado: NFes interestaduais emitidas automaticamente
	s_sql = "SELECT" & _
				" t_NFE_EMISSAO.id," & _
				" t_NFE_EMISSAO.NFe_serie_NF," & _
				" t_NFE_EMISSAO.NFe_numero_NF," & _
				" t_NFE_EMISSAO.controle_impostos_status," & _
				" t_PEDIDO.pedido," & _
				" t_PEDIDO.a_entregar_data_marcada," & _
				" t_PEDIDO.st_entrega," & _
				" t_PEDIDO.transportadora_id," & _
				" t_NFE_IMAGEM.dest__UF AS UF," & _
				" t_NFE_IMAGEM.total__vFCPUFDest," & _
				" t_NFE_IMAGEM.total__vICMSUFDest," & _
				" t_NFE_IMAGEM.total__vICMSUFRemet" & _
				s_select_devolucao_grupo_1 & _
			" FROM t_NFE_EMISSAO" & _
				" INNER JOIN t_NFE_IMAGEM ON (t_NFE_EMISSAO.NFe_numero_NF=t_NFE_IMAGEM.NFe_numero_NF AND t_NFE_EMISSAO.NFe_serie_NF=t_NFE_IMAGEM.NFe_serie_NF AND t_NFE_EMISSAO.id_nfe_emitente=t_NFE_IMAGEM.id_nfe_emitente)" & _
				" INNER JOIN (SELECT id_nfe_emitente, NFe_serie_NF, NFe_numero_NF, max(id) AS id FROM t_NFE_IMAGEM GROUP BY id_nfe_emitente, NFe_serie_NF, NFe_numero_NF) img_max_id ON (t_NFE_IMAGEM.id = img_max_id.id AND t_NFE_IMAGEM.id_nfe_emitente=img_max_id.id_nfe_emitente)" & _
				" INNER JOIN (SELECT id_nfe_emitente, NFe_serie_NF, NFe_numero_NF, max(id) AS id FROM t_NFE_EMISSAO GROUP BY id_nfe_emitente, NFe_serie_NF, NFe_numero_NF) emi_max_id ON (t_NFE_EMISSAO.id = emi_max_id.id AND t_NFE_EMISSAO.id_nfe_emitente=emi_max_id.id_nfe_emitente)" & _
				" INNER JOIN t_PEDIDO ON (t_NFE_EMISSAO.pedido=t_PEDIDO.pedido AND t_NFE_EMISSAO.id_nfe_emitente=t_PEDIDO.id_nfe_emitente)" & _
			" WHERE t_NFE_EMISSAO.pedido IS NOT NULL " & _
			s_where & _
			s_where_pedido_grupo_1

'	Segundo grupo selecionado: NFes interestaduais emitidas manualmente, com um conjunto específico de CFOPs relacionados
			s_sql = s_sql & _
			"UNION " & _
			"SELECT" & _
				" t_NFE_EMISSAO.id," & _
				" t_NFE_EMISSAO.NFe_serie_NF," & _
				" t_NFE_EMISSAO.NFe_numero_NF," & _
				" t_NFE_EMISSAO.controle_impostos_status," & _
				" pedidos_nf_ok.pedido," & _
				" pedidos_nf_ok.a_entregar_data_marcada," & _
				" pedidos_nf_ok.st_entrega," & _
				" pedidos_nf_ok.transportadora_id," & _
				" t_NFE_IMAGEM.dest__UF AS UF," & _
				" t_NFE_IMAGEM.total__vFCPUFDest," & _
				" t_NFE_IMAGEM.total__vICMSUFDest," & _
				" t_NFE_IMAGEM.total__vICMSUFRemet" & _
				s_select_devolucao_grupo_2 & _
			" FROM t_NFE_EMISSAO" & _
				" INNER JOIN t_NFE_IMAGEM ON (t_NFE_EMISSAO.NFe_numero_NF=t_NFE_IMAGEM.NFe_numero_NF AND t_NFE_EMISSAO.NFe_serie_NF=t_NFE_IMAGEM.NFe_serie_NF AND t_NFE_EMISSAO.id_nfe_emitente=t_NFE_IMAGEM.id_nfe_emitente)" & _
				" INNER JOIN (SELECT id_nfe_emitente, NFe_serie_NF, NFe_numero_NF, max(id) AS id FROM t_NFE_IMAGEM GROUP BY id_nfe_emitente, NFe_serie_NF, NFe_numero_NF) img_max_id ON (t_NFE_IMAGEM.id = img_max_id.id AND t_NFE_IMAGEM.id_nfe_emitente=img_max_id.id_nfe_emitente)" & _
				" INNER JOIN (SELECT id_nfe_emitente, NFe_serie_NF, NFe_numero_NF, max(id) AS id FROM t_NFE_EMISSAO GROUP BY id_nfe_emitente, NFe_serie_NF, NFe_numero_NF) emi_max_id ON (t_NFE_EMISSAO.id = emi_max_id.id AND t_NFE_EMISSAO.id_nfe_emitente=emi_max_id.id_nfe_emitente)" & _
				" INNER JOIN (SELECT * FROM t_PEDIDO WHERE (ISNUMERIC(t_PEDIDO.obs_2) = 1) AND (LEN(t_PEDIDO.obs_2) < 10) AND " & s_where_data &") pedidos_nf_ok" & _
				"		ON t_NFE_EMISSAO.NFe_numero_NF=CONVERT(INT, pedidos_nf_ok.obs_2) AND t_NFE_EMISSAO.id_nfe_emitente = pedidos_nf_ok.id_nfe_emitente " & _
			" WHERE t_NFE_EMISSAO.pedido IS NULL " & _
				" AND " & _
				" (EXISTS (SELECT 1  " & _
							" FROM t_NFe_IMAGEM_ITEM  " & _
							" WHERE t_NFE_IMAGEM.id = t_NFe_IMAGEM_ITEM.id_nfe_imagem " & _
							" AND t_NFe_IMAGEM_ITEM.det__CFOP IN ('5102','6102','6108','5119','6119','5910','6910')))" & _
			s_where & _
			s_where_pedido_grupo_2
	
	s_sql = s_sql & " ORDER BY t_NFE_IMAGEM.dest__UF, pedido, t_NFE_EMISSAO.NFe_serie_NF, t_NFE_EMISSAO.NFe_numero_NF"
	
  ' CABEÇALHO
	cab_table = "<table cellspacing=0 id='tabelaRelatorio'>" & chr(13)
	cab = "	<tr style='background:azure'>" & chr(13) & _
		  "		<td class='ME MC MD' style='width:70px' align='center' valign='bottom' nowrap><span class='R'>Pedido</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:70px' align='center' valign='bottom' nowrap><span class='R'>Nº NF</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:180px' align='left' valign='bottom' nowrap><span class='R'>Transportadora</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:30px' align='center' valign='bottom' nowrap><span class='R'>Data Coleta</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:30px' align='right' valign='bottom' nowrap><span class='R'>FCP</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:70px' align='right' valign='bottom' nowrap><span class='R'>ICMS UF </br> Destino</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:70px' align='right' valign='bottom' nowrap><span class='R'>ICMS UF </br> Origem</span></td>" & chr(13)
	
	if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then
		cab = cab & _
		  "		<td class='MC MD' style='width:30px' align='right' valign='bottom' nowrap><span class='R'>FCP<br />(Devol/Cancel)</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:70px' align='right' valign='bottom' nowrap><span class='R'>ICMS UF Destino<br />(Devol/Cancel)</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:70px' align='right' valign='bottom' nowrap><span class='R'>ICMS UF Origem<br />(Devol/Cancel)</span></td>" & chr(13)
		end if

	cab = cab & _
		  "	</tr>" & chr(13)

	n_reg_geral = 0
	n_reg_uf = 0
	s_uf = ""
	s_uf_anterior = ""
	'x = cab_table & cab
	x = ""

	set r = cn.execute(s_sql)
	do while Not r.Eof
		
	'	SE A NOTA NÃO FOI COMPLETAMENTE EMITIDA, PULAR
		
		s_num_nfe = NFeFormataNumeroNF(Trim("" & r("NFe_numero_NF")))
		s_serie_nfe = NFeFormataSerieNF(Trim("" & r("NFe_serie_NF")))

		if IsNFeCompletamenteEmitida(rNfeEmitente.id, s_serie_nfe, s_num_nfe, ChaveAcesso) then

			s_uf_anterior = s_uf
			s_uf = Trim("" & r("UF"))
			if (s_uf <> s_uf_anterior) And (s_uf_anterior = "") then
				  ' SE FOR A PRIMEIRA ITERAÇÃO, INCLUIR CABEÇALHO
					x = x & cab_table 
					colSpan = 7
					if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then colSpan = colSpan + 3
					x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
							"       <td class='MDTE' nowrap colspan='" & CStr(colSpan) & "' align='left'><span class='C'>" & _
							"UF:  &nbsp; " & s_uf & "</span></td>" & chr(13) & _
							"	</tr>" & chr(13)
					x = x & cab
			    end if
			if (s_uf <> s_uf_anterior) And (s_uf_anterior <> "") then
				colSpan = 4
				'SE MUDOU UF, MOSTRA SUBTOTAL
				x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
						"       <td class='MTBE' nowrap colspan='" & CStr(colSpan) & "' align='right'><span class='C'>" & _
						"TOTAL UF:   </span></td>" & chr(13) & _
						"       <td class='MTB' nowrap align='right'><span class='C'>"& formata_moeda(total_fcp_uf) & "</span></td>" & chr(13) & _
						"       <td class='MTB' nowrap align='right'><span class='C'>"& formata_moeda(total_icms_destino_uf) & "</span></td>" & chr(13) & _
						"       <td class='MTBD' nowrap align='right'><span class='C'>"& formata_moeda(total_icms_origem_uf) & "</span></td>" & chr(13)

				if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then
					x = x & _
						"       <td class='MTB' nowrap align='right'><span class='C'>"& formata_moeda(total_fcp_uf_proporcional) & "</span></td>" & chr(13) & _
						"       <td class='MTB' nowrap align='right'><span class='C'>"& formata_moeda(total_icms_destino_uf_proporcional) & "</span></td>" & chr(13) & _
						"       <td class='MTBD' nowrap align='right'><span class='C'>"& formata_moeda(total_icms_origem_uf_proporcional) & "</span></td>" & chr(13)
					end if

				x = x & _
						"	</tr>" & chr(13)
						
				' MOSTRA O TOTAL DE REGISTROS
				colSpan = 7
				if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then colSpan = colSpan + 3
				x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
						"       <td class='MDBE' nowrap colspan='" & CStr(colSpan) & "' align='right'><span class='C'>" & _
						"TOTAL DE REGISTRO(S):  &nbsp; " & formata_inteiro(n_reg_uf) & "</span></td>" & chr(13) & _
						"	</tr>" & chr(13)
				'FECHA TABELA
				x = x & "</table>" & chr(13)

				x = x & "<br>" & chr(13)

			  ' CABEÇALHO
				colSpan = 7
				if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then colSpan = colSpan + 3
				x = x & cab_table 
				x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
						"       <td class='MDTE' nowrap colspan='" & CStr(colSpan) & "' align='left'><span class='C'>" & _
						"UF:  &nbsp; " & s_uf & "</span></td>" & chr(13) & _
						"	</tr>" & chr(13)
				x = x & cab

			  ' ZERA CONTAGEM
				n_reg_uf = 0
				total_fcp_uf = 0
				total_icms_origem_uf = 0
				total_icms_destino_uf = 0
				total_fcp_uf_proporcional = 0
				total_icms_origem_uf_proporcional = 0
				total_icms_destino_uf_proporcional = 0
				end if

		 '	CONTAGEM
			n_reg_geral = n_reg_geral + 1
			n_reg_uf = n_reg_uf + 1

			intCodCor = COD_COR_PRETO
			intOrdenacaoCor = 0
			s_html_color = "black"
			
			s_html_color = " style='color:" & s_html_color & ";'"
			
		'	MONTA O HTML DA LINHA DA TABELA
		'	===============================
			s_row = "	<tr onmouseover='realca_cor_mouse_over(this);' onmouseout='realca_cor_mouse_out(this);'>" & chr(13)

		'> Nº PEDIDO
			s_pedido = Trim("" & r("pedido"))
			s = s_pedido
			if s = "" then s = "&nbsp;"
			s_row = s_row & _
					"		<td align='center' valign='top' class='ME MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">&nbsp;<a href='javascript:fRELConcluir(" & chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'" & s_html_color & ">" & Trim("" & r("pedido")) & "</a></span>" & chr(13) & _
					"			<input type='hidden' name='c_numero_pedido' value='" & s & "' />" & chr(13) & _
					"		</td>" & chr(13)

		'	Nº NFe
			s_num_nfe = NFeFormataNumeroNF(Trim("" & r("NFe_numero_NF")))
			if s_num_nfe <> "" then
				s = "<span class='C'" & s_html_color & ">" & NFeFormataNumeroNF(s_num_nfe) & "</span>"
				s_link_nfe = monta_link_para_DANFE(s_pedido, MAX_PERIODO_LINK_DANFE_DISPONIVEL_NO_PEDIDO_EM_DIAS, s)
				s_link_habilita_nfe = s_link_nfe
				if s_link_nfe = "" then s_link_nfe = "<span class='C' style='color:gray;'>" & s_num_nfe & "</span>"
			else
				s_link_nfe = "&nbsp;"
				end if
				
			s_row = s_row & _
					"		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					"			" & s_link_nfe & chr(13) & _
					"		</td>" & chr(13)

		'	TRANSPORTADORA
			s = Trim("" & r("transportadora_id"))
			if s = "" then s = "&nbsp;"
			s_row = s_row & _
					"		<td align='left' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					"			<input type='hidden' name='c_pedido_transportadora' value='" & Trim("" & r("transportadora_id")) & "' />" & chr(13) & _
					"		</td>" & chr(13)

		'	DATA DE COLETA
			s = formata_data(r("a_entregar_data_marcada"))
			if s = "" then s = "&nbsp;"
			s_row = s_row & _
					"		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					"		</td>" & chr(13)

		'	FCP
			valor = converte_numero(Trim("" & r("total__vFCPUFDest")))
			s = formata_moeda(valor)
			if s = "" then s = "&nbsp;"
			s_row = s_row & _
					"		<td align='right' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					"		</td>" & chr(13)

		'	ICMS DESTINO
			valor = converte_numero(Trim("" & r("total__vICMSUFDest")))
			s = formata_moeda(valor)
			if s = "" then s = "&nbsp;"
			s_row = s_row & _
					"		<td align='right' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					"		</td>" & chr(13)

		'	ICMS ORIGEM
			valor = converte_numero(Trim("" & r("total__vICMSUFRemet")))
			s = formata_moeda(valor)
			if s = "" then s = "&nbsp;"
			s_row = s_row & _
					"		<td align='right' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					"		</td>" & chr(13)

			if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then
				if r("qtde_itens_devolucao") > 0 then
					if r("vl_pedido_NF") = 0 then
					'	SITUAÇÃO QUE NÃO DEVE OCORRER NA PRÁTICA, MAS SE ACONTECER, EXIBE O VALOR INTEGRAL DOS IMPOSTOS
						percProporcao = 1
					else
						percProporcao = r("vl_devolucao_NF") / r("vl_pedido_NF")
						end if
				else
				'	PEDIDOS CANCELADOS EXIBEM O VALOR INTEGRAL DOS IMPOSTOS
					percProporcao = 1
					end if

			'	FCP (Proporcional)
				valor = percProporcao * converte_numero(Trim("" & r("total__vFCPUFDest")))
				s = formata_moeda(valor)
			'	USA O VALOR ATRAVÉS DA CONVERSÃO DO VALOR FORMATADO P/ EVITAR DIFERENÇAS DE ARREDONDAMENTO
				vl_fcp_uf_proporcional = converte_numero(s)
				if s = "" then s = "&nbsp;"
				s_row = s_row & _
						"		<td align='right' valign='top' class='MC MD'>" & chr(13) & _
						"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
						"		</td>" & chr(13)

			'	ICMS DESTINO (Proporcional)
				valor = percProporcao * converte_numero(Trim("" & r("total__vICMSUFDest")))
				s = formata_moeda(valor)
			'	USA O VALOR ATRAVÉS DA CONVERSÃO DO VALOR FORMATADO P/ EVITAR DIFERENÇAS DE ARREDONDAMENTO
				vl_icms_destino_uf_proporcional = converte_numero(s)
				if s = "" then s = "&nbsp;"
				s_row = s_row & _
						"		<td align='right' valign='top' class='MC MD'>" & chr(13) & _
						"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
						"		</td>" & chr(13)

			'	ICMS ORIGEM (Proporcional)
				valor = percProporcao * converte_numero(Trim("" & r("total__vICMSUFRemet")))
				s = formata_moeda(valor)
			'	USA O VALOR ATRAVÉS DA CONVERSÃO DO VALOR FORMATADO P/ EVITAR DIFERENÇAS DE ARREDONDAMENTO
				vl_icms_origem_uf_proporcional = converte_numero(s)
				if s = "" then s = "&nbsp;"
				s_row = s_row & _
						"		<td align='right' valign='top' class='MC MD'>" & chr(13) & _
						"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
						"		</td>" & chr(13)
				end if

			s_row = s_row & "	</tr>" & chr(13)

			x = x & s_row
			if (n_reg_geral mod 100) = 0 then
				Response.Write x
				x = ""
				end if

			total_fcp_uf = total_fcp_uf + CCur(converte_numero(r("total__vFCPUFDest")))
			total_icms_destino_uf = total_icms_destino_uf + CCur(converte_numero(r("total__vICMSUFDest")))
			total_icms_origem_uf = total_icms_origem_uf + CCur(converte_numero(r("total__vICMSUFRemet")))
			total_fcp_geral = total_fcp_geral + CCur(converte_numero(r("total__vFCPUFDest")))
			total_icms_destino_geral = total_icms_destino_geral + CCur(converte_numero(r("total__vICMSUFDest")))
			total_icms_origem_geral = total_icms_origem_geral + CCur(converte_numero(r("total__vICMSUFRemet")))
			total_fcp_uf_proporcional = total_fcp_uf_proporcional + vl_fcp_uf_proporcional
			total_icms_origem_uf_proporcional = total_icms_origem_uf_proporcional + vl_icms_origem_uf_proporcional
			total_icms_destino_uf_proporcional = total_icms_destino_uf_proporcional + vl_icms_destino_uf_proporcional
			total_fcp_geral_proporcional = total_fcp_geral_proporcional + vl_fcp_uf_proporcional
			total_icms_origem_geral_proporcional = total_icms_origem_geral_proporcional + vl_icms_origem_uf_proporcional
			total_icms_destino_geral_proporcional = total_icms_destino_geral_proporcional + vl_icms_destino_uf_proporcional
			end if
		
		r.MoveNext
		loop
	
	'	ÚLTIMO SUBTOTAL NO RELATÓRIO
	colSpan = 4
	x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
			"       <td class='MTBE' nowrap colspan='" & CStr(colSpan) & "' align='right'><span class='C'>" & _
			"TOTAL UF:   </span></td>" & chr(13) & _
			"       <td class='MTB' nowrap  align='right'><span class='C'>"& formata_moeda(total_fcp_uf) & "</span></td>" & chr(13) & _
			"       <td class='MTB' nowrap  align='right'><span class='C'>"& formata_moeda(total_icms_destino_uf) & "</span></td>" & chr(13) & _
			"       <td class='MTBD' nowrap  align='right'><span class='C'>"& formata_moeda(total_icms_origem_uf) & "</span></td>" & chr(13)

	if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then
		x = x & _
			"       <td class='MTB' nowrap align='right'><span class='C'>"& formata_moeda(total_fcp_uf_proporcional) & "</span></td>" & chr(13) & _
			"       <td class='MTB' nowrap align='right'><span class='C'>"& formata_moeda(total_icms_destino_uf_proporcional) & "</span></td>" & chr(13) & _
			"       <td class='MTBD' nowrap align='right'><span class='C'>"& formata_moeda(total_icms_origem_uf_proporcional) & "</span></td>" & chr(13)
		end if

	x = x & _
			"	</tr>" & chr(13)

	'	MOSTRA O TOTAL DE REGISTROS
	colSpan = 7
	if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then colSpan = colSpan + 3
	x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
			"       <td class='MDBE' nowrap colspan='" & CStr(colSpan) & "' align='right'><span class='C'>" & _
			"TOTAL DE REGISTRO(S):  &nbsp; " & formata_inteiro(n_reg_uf) & "</span></td>" & chr(13) & _
			"	</tr>" & chr(13)

	'	FECHA TABELA
	x = x & "</table>" & chr(13)

	x = x & "<br>" & chr(13)

	'	TOTAL NO RELATÓRIO
	colSpan = 7
	if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then colSpan = colSpan + 3
	x = x & cab_table 
	x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
			"       <td class='MDTE' nowrap colspan='" & CStr(colSpan) & "' align='left'><span class='C'>" & _
			"TOTAL GERAL  &nbsp;</span></td>" & chr(13) & _
			"	</tr>" & chr(13)
	x = x & cab

	colSpan = 4
	x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
			"       <td class='MTBE' nowrap colspan='" & CStr(colSpan) & "' align='right'><span class='C'>" & _
			"TOTAL :   </span></td>" & chr(13) & _
			"       <td class='MTB' nowrap align='right'><span class='C'>"& formata_moeda(total_fcp_geral) & "</span></td>" & chr(13) & _
			"       <td class='MTB' nowrap align='right'><span class='C'>"& formata_moeda(total_icms_destino_geral) & "</span></td>" & chr(13) & _
			"       <td class='MTBD' nowrap align='right'><span class='C'>"& formata_moeda(total_icms_origem_geral) & "</span></td>" & chr(13)

	if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then
		x = x & _
			"       <td class='MTB' nowrap align='right'><span class='C'>"& formata_moeda(total_fcp_geral_proporcional) & "</span></td>" & chr(13) & _
			"       <td class='MTB' nowrap align='right'><span class='C'>"& formata_moeda(total_icms_destino_geral_proporcional) & "</span></td>" & chr(13) & _
			"       <td class='MTBD' nowrap align='right'><span class='C'>"& formata_moeda(total_icms_origem_geral_proporcional) & "</span></td>" & chr(13)
		end if

	x = x & _
			"	</tr>" & chr(13)

	' TOTAL GERAL DE REGISTROS
	colSpan = 7
	if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then colSpan = colSpan + 3
	x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
			"       <td class='MDBE' nowrap colspan='" & CStr(colSpan) & "' align='right'><span class='C'>" & _
			"TOTAL DE REGISTRO(S):  &nbsp; " & formata_inteiro(n_reg_geral) & "</span></td>" & chr(13) & _
			"	</tr>" & chr(13)

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	colSpan = 11
	if (ckb_pedidos_cancelados <> "") OR (ckb_pedidos_com_devolucao <> "") then colSpan = colSpan + 3
	if n_reg_geral = 0 then
		x = cab_table & cab
		x = x & "	<tr>" & chr(13) & _
				"		<td class='MT ALERTA' colspan='" & CStr(colSpan) & "' align='center'><span class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</table>" & chr(13)
	
	Response.write x

	qtde_notas = n_reg_geral
	
	if r.State <> 0 then r.Close
	set r=nothing
	
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
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function realca_cor_mouse_over(c) {
	c.style.backgroundColor = 'palegreen';
}

function realca_cor_mouse_out(c) {
	c.style.backgroundColor = '';
}

function fRELConcluir( id_pedido ){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=id_pedido;
	fREL.action = "Pedido.asp";
	fREL.submit();
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
</center>
</body>



<% else %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- Nº DO PEDIDO P/ CONSULTAR O PEDIDO AO CLICAR SOBRE O NÚMERO -->
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<!-- FILTROS -->
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>" />
<input type="hidden" name="c_dt_coleta_inicio" id="c_dt_coleta_inicio" value="<%=c_dt_coleta_inicio%>" />
<input type="hidden" name="c_dt_coleta_termino" id="c_dt_coleta_termino" value="<%=c_dt_coleta_termino%>" />
<input type="hidden" name="c_uf" id="c_uf" value="<%=c_uf%>" />
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>" />
<input type="hidden" name="ckb_pedidos_cancelados" id="ckb_pedidos_cancelados" value="<%=ckb_pedidos_cancelados%>" />
<input type="hidden" name="ckb_pedidos_com_devolucao" id="ckb_pedidos_com_devolucao" value="<%=ckb_pedidos_com_devolucao%>" />
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="840" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Impostos Pagos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='840' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Tipo de Relatório:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>Impostos Pagos</span></td></tr>"

	s = ""
	if ckb_pedidos_cancelados <> "" then
		if s <> "" then s = s & ", "
		s = s & "Pedidos Cancelados"
		end if
	
	if ckb_pedidos_com_devolucao <> "" then
		if s <> "" then s = s & ", "
		s = s & "Pedidos com Devolução"
		end if

	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Pedidos:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"

	s = c_transportadora
	if s = "" then 
		s = "todas"
	else
		if (s_nome_transportadora <> "") And (Ucase(s_nome_transportadora) <> Ucase(c_transportadora)) then s = s & "  (" & s_nome_transportadora & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Transportadora:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"
	
	s = c_uf
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>UF:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"
	
	s = Trim("" & rNfeEmitente.apelido)
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>CD:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"

	s = c_dt_coleta_inicio
	if s <> "" then s = s & " a " & c_dt_coleta_termino
	if s = "" then s = "N.I."
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Período:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"
	
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

</form>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>


<!-- ************   SEPARADOR   ************ -->
<table width="840" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="840" cellspacing="0">
<tr>
	<td align="center">
		<a name="bVOLTA" id="bVOLTA" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>

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
