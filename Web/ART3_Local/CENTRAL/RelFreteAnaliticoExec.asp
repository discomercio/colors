<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=true %>
<% Response.ContentType = "text/html" %>

<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelFreteAnaliticoExec.asp
'     ======================================================
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
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_FRETE_ANALITICO, s_lista_operacoes_permitidas) then
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	dim s, s_aux
	dim c_dt_entregue_inicio, c_dt_entregue_termino
	dim c_transportadora, c_fabricante, c_loja, c_vendedor, c_indicador, c_uf, c_tipo_frete
	dim s_nome_vendedor, s_nome_indicador, s_nome_loja, s_nome_fabricante, s_nome_transportadora
	dim rb_frete_status, rb_tipo_saida, codigo_tipo_frete, rb_tipo_nf
	dim c_empresa

	alerta = ""

	rb_frete_status = Trim(Request.Form("rb_frete_status"))
	rb_tipo_saida = Ucase(Trim(Request.Form("rb_tipo_saida")))
	rb_tipo_nf = Trim(Request.Form("rb_tipo_nf"))

	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))
	
	c_transportadora = Trim(Request.Form("c_transportadora"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_loja = retorna_so_digitos(Trim(Request.Form("c_loja")))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	c_indicador = Ucase(Trim(Request.Form("c_indicador")))
	c_uf = Ucase(Trim(Request.Form("c_uf")))
    c_tipo_frete = Trim(Request.Form("c_tipo_frete"))
	c_empresa = Trim(Request.Form("c_empresa"))

	if alerta = "" then
		s_nome_fabricante = ""
		if c_fabricante <> "" then
			s = "SELECT fabricante, nome FROM t_FABRICANTE WHERE (fabricante='" & c_fabricante & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "FABRICANTE " & c_fabricante & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_fabricante = iniciais_em_maiusculas(Trim("" & rs("nome")))
				end if
			end if
		end if
	
	if alerta = "" then
		s_nome_transportadora = ""
		if c_transportadora <> "" then
			s = "SELECT nome FROM t_TRANSPORTADORA WHERE (id='" & c_transportadora & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "TRANSPORTADORA " & c_transportadora & " NÃO ESTÁ CADASTRADA."
			else
				s_nome_transportadora = iniciais_em_maiusculas(Trim("" & rs("nome")))
				end if
			end if
		end if
	
	if alerta = "" then
		s_nome_vendedor = ""
		if c_vendedor <> "" then
			s = "SELECT nome_iniciais_em_maiusculas FROM t_USUARIO WHERE (usuario='" & c_vendedor & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "VENDEDOR " & c_vendedor & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_vendedor = Trim("" & rs("nome_iniciais_em_maiusculas"))
				end if
			end if
		end if
	
	if alerta = "" then
		s_nome_indicador = ""
		if c_indicador <> "" then
			s = "SELECT razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido='" & c_indicador & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "INDICADOR " & c_indicador & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_indicador = Trim("" & rs("razao_social_nome_iniciais_em_maiusculas"))
				end if
			end if
		end if
	
	if alerta = "" then
		s_nome_loja = ""
		if c_loja <> "" then
			s = "SELECT * FROM t_LOJA WHERE (CONVERT(smallint,loja) = " & c_loja & ")"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "LOJA " & c_loja & " NÃO ESTÁ CADASTRADA."
			else
				s_nome_loja = iniciais_em_maiusculas(Trim("" & rs("nome")))
				end if
			end if
		end if
	
'	Período de consulta está restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
	'	PERÍODO DE ENTREGA
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_entregue_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if
		
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_entregue_termino
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if
		
		if alerta = "" then
			if c_dt_entregue_inicio = "" then c_dt_entregue_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
			end if
	
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if
	
	dim blnSaidaExcel
	blnSaidaExcel = False
	if alerta = "" then
		if rb_tipo_saida = "XLS" then
			blnSaidaExcel = True
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=FreteAnalitico_" & formata_data_yyyymmdd(Now) & "_" & formata_hora_hhnnss(Now) & ".xls"
			Response.Write "<h2>Relatório de Frete (Analítico)</h2>"
			Response.Write monta_texto_filtro
			Response.Write "<br><br>"
			consulta_executa
			Response.End
			end if
		end if

	'MEMORIZA OPÇÃO DE CONSULTA NO BD
	call set_default_valor_texto_bd(usuario, "RelFreteAnalitico|rb_tipo_nf", rb_tipo_nf)



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S
' _____________________________________________________________________________________________

' _____________________________________
' MONTA TEXTO FILTRO
'
function monta_texto_filtro
dim s, s_aux, s_filtro

	s_filtro = ""
	
	if Not blnSaidaExcel then s_filtro = "<table width='836' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>"

	s = ""
	s_aux = c_dt_entregue_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_entregue_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	
	if blnSaidaExcel then
		s_filtro = s_filtro & "<span class='N'>Período de Entrega:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>Período de Entrega:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	s = c_transportadora
	if s = "" then 
		s = "todas"
	else
		if (s_nome_transportadora <> "") And (s_nome_transportadora <> c_transportadora) then s = s & "  (" & s_nome_transportadora & ")"
		end if
	
	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>Transportadora:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>Transportadora:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if

    s = c_tipo_frete
	if s = "" then 
		s = "todos"
	else
		if (c_tipo_frete <> "") then s = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE,c_tipo_frete)
		end if
	
	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>Tipo de Frete:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>Tipo de Frete:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	s = c_fabricante
	if s = "" then 
		s = "todos"
	else
		s = s & "  (" & s_nome_fabricante & ")"
		end if
	
	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>Fabricante:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>Fabricante:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	s = c_loja
	if s = "" then 
		s = "todas"
	else
		s = s & "  (" & s_nome_loja & ")"
		end if
	
	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>Loja:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>Loja:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	s = c_vendedor
	if s = "" then 
		s = "todos"
	else
		if (s_nome_vendedor <> "") And (s_nome_vendedor <> c_vendedor) then s = s & "  (" & s_nome_vendedor & ")"
		end if
	
	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>Vendedor:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>Vendedor:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	s = c_indicador
	if s = "" then 
		s = "todos"
	else
		if (s_nome_indicador <> "") And (s_nome_indicador <> c_indicador) then s = s & "  (" & s_nome_indicador & ")"
		end if
	
	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>Indicador:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>Indicador:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	s = c_uf
	if s = "" then s = "todos"
	
	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>UF:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>UF:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	s = c_empresa
	if s = "" then 
		s = "todas"
	else
		s = obtem_apelido_empresa_NFe_emitente(c_empresa)
		end if

	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>Empresa:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>Empresa:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if

	s = rb_frete_status
	if s = "" then 
		s = "Ambos"
	elseif s = "0" then
		s = "Frete não preenchido"
	elseif s = "1" then
		s = "Frete já preenchido"
		end if
	
	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>Status do Frete:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>Status do Frete:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if
	
	s = rb_tipo_nf
	if s = "1" then
		s = "Somente de Remessa, quando houver"
	else
		s = "Fatura e de Remessa"
		end if

	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>Tipo de NF:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>Tipo de NF:&nbsp;</span></td><td align='left' valign='top'>" & _
				   "<span class='N'>" & s & "</span></td></tr>"
		end if

	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>Emissão:&nbsp;" & formata_data_hora(Now) & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
				   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>"
		s_filtro = s_filtro & "</table>"
		end if
	
	monta_texto_filtro = s_filtro
end function



' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
const MSO_NUMBER_FORMAT_FLOAT = "\#\#\#\,\#\#\#\,\#\#0\.0"
const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
const MSO_NUMBER_FORMAT_TEXTO = "\@"
dim r
dim x,cont
dim cab, cab_table
dim s_aux, s_sql, s_where,s_sql_2
dim n_reg
dim vlFrete, vlPrecoNF, vlPrecoVenda
dim percPrecoVenda, percPrecoNF
dim vlTotalFrete, vlTotalPrecoNF, vlTotalPrecoVenda, vlTotalPrecoNFAjustado, vlTotalPrecoVendaAjustado
dim vlTotalPercPrecoNF, vlTotalPercPrecoVenda
dim strTransportadora, strTransportadoraAnterior, strTransportadoraAux, strPlural,strPluralFrete,pedidoAnterior
dim strCidade, strUf, strCidadeUf, strCnpjCpfCliente, strNF
dim intQtdeTotalPedidos, intQtdeTransportadoras
dim intQtdeSubTotalPedidos, vlSubTotalFrete, vlSubTotalPrecoNF, vlSubTotalPrecoVenda, vlSubTotalPercPrecoNF, vlSubTotalPercPrecoVenda
dim strPercPrecoNF, strPercPrecoVenda, strVlSubTotalPercPrecoNF, strVlSubTotalPercPrecoVenda, strVlTotalPercPrecoNF, strVlTotalPercPrecoVenda
dim vlCubagem, intQtdeVolumes, vlPeso
dim vlTotalCubagem, intQtdeTotalVolumes, vlTotalPeso
dim vlSubTotalCubagem, intQtdeSubTotalVolumes, vlSubTotalPeso,intQtdeSubTotalFretes,intQtdeTotalFretes,s_where_tipo_frete
dim qtde_Pedido,lista_pedidos_total_geral,strPercPrecoVendaSomado,vlTotalPercPrecoVendaAjustado,vlTotalPercPrecoNFAjustado,strVlTotalPercPrecoVendaAjustado,strVlTotalPercPrecoNFAjustado
dim pedidoatual, s_where_externo


'	CRITÉRIOS DE RESTRIÇÃO
	s_where = "(p.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')"
    s_where_externo = "WHERE (0 = 0)"
	
'	FILTRO: PERÍODO DE ENTREGA
	if IsDate(c_dt_entregue_inicio) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (p.entregue_data >= " & bd_formata_data(StrToDate(c_dt_entregue_inicio)) & ")"
		end if
	
	if IsDate(c_dt_entregue_termino) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (p.entregue_data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		end if
	
'	FILTRO: TRANSPORTADORA
	if c_transportadora <> "" then
		if s_where_externo <> "" then s_where_externo = s_where_externo & " AND"
		s_where_externo = s_where_externo & " (s.transportadora_id = '" & c_transportadora & "')"
		end if

'	FILTRO: TIPO DE FRETE
	if c_tipo_frete <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (pf.codigo_tipo_frete = '" & c_tipo_frete & "')"
        s_where_tipo_frete = "AND (codigo_tipo_frete = pf.codigo_tipo_frete)"
		end if

'	FILTRO: FABRICANTE
	if c_fabricante <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & _
			" (" & _
				"p.pedido IN " & _
					"(" & _
						"SELECT " & _
							"pedido" & _
						" FROM t_PEDIDO_ITEM i" & _
						" WHERE" & _
							" (i.pedido = p.pedido)" & _
							" AND (i.fabricante = '" & c_fabricante & "')" & _
					")" & _
			")"
		end if
	
'	FILTRO: LOJA
	if c_loja <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (CONVERT(smallint,p.loja) = " & c_loja & ")"
		end if
	
'	FILTRO: VENDEDOR
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (p.vendedor = '" & c_vendedor & "')"
		end if
	
'	FILTRO: INDICADOR
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (p.indicador = '" & c_indicador & "')"
		end if
	
'	FILTRO: UF
	if c_uf <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
			s_where = s_where & _
				" (" & _
					"((p.st_end_entrega <> 0) And (p.EndEtg_uf = '" & c_uf & "'))" & _
					" OR " & _
					"((p.st_end_entrega = 0) And (p.endereco_uf = '" & c_uf & "'))" & _
				")"
		else
			s_where = s_where & _
				" (" & _
					"((p.st_end_entrega <> 0) And (p.EndEtg_uf = '" & c_uf & "'))" & _
					" OR " & _
					"((p.st_end_entrega = 0) And (c.uf = '" & c_uf & "'))" & _
				")"
			end if
		end if
	
'	FILTRO: CD
	 if c_empresa <> "" then
        if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (p.id_nfe_emitente = " & c_empresa & ")"
		end if

'	FILTRO: STATUS DO FRETE
	if rb_frete_status <> "" then
		if s_where_externo <> "" then s_where_externo = s_where_externo & " AND"
        if rb_frete_status = "0" then        
		    s_where_externo = s_where_externo & " (qtde_fretes = 0)"
        elseif rb_frete_status = "1" then
		    s_where_externo = s_where_externo & " (qtde_fretes > 0)"
        end if
	end if
	
'	MONTA SQL DE CONSULTA
    s_sql = "SELECT transportadora_id " & _
	            ",data " & _
	            ",pedido " & _
				",obs_2 " & _
				",obs_3 " & _
	            ",qtde_fretes " & _
	            ",qtde_pedido " & _
	            ",st_end_entrega " & _
	            ",EndEtg_cidade " & _
	            ",EndEtg_uf " & _
				",EndEtg_cnpj_cpf " & _
	            ",cidade " & _
	            ",uf " & _
				",cliente_cnpj_cpf " & _
	            ",vl_frete " & _
	            ",codigo_tipo_frete " & _
	            ",descricao " & _
	            ",preco_NF " & _
	            ",cubagem " & _
	            ",qtde_volumes " & _
	            ",peso " & _
	            ",preco_venda " & _
	            ",numeros_NF " & _
            "FROM (  "

	s_sql = s_sql & "SELECT " & _
				"Coalesce(pf.transportadora_id, p.transportadora_id) AS transportadora_id,  " & _
				"p.data, " & _
				"p.pedido, " & _
				"p.obs_2, " & _
				"p.obs_3, " & _
                "( SELECT COUNT(*) FROM t_PEDIDO_FRETE WHERE pedido = p.pedido ) AS qtde_fretes, " & _
                "( SELECT COUNT(*) FROM t_PEDIDO_FRETE WHERE pedido=p.pedido AND transportadora_id=pf.transportadora_id " 
                if s_where_tipo_frete <> "" then s_sql = s_sql + s_where_tipo_frete 
                s_sql = s_sql + ") as qtde_pedido," & _  
				"p.st_end_entrega, " & _
				"p.EndEtg_cidade, " & _
				"p.EndEtg_uf, " & _
				"p.EndEtg_cnpj_cpf,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				"p.endereco_cidade AS cidade, " & _
				"p.endereco_uf AS uf, " & _
				"p.endereco_cnpj_cpf AS cliente_cnpj_cpf,"
	else
		s_sql = s_sql & _
				"c.cidade, " & _
				"c.uf, " & _
				"c.cnpj_cpf AS cliente_cnpj_cpf,"
		end if

	s_sql = s_sql & _
				"Coalesce(pf.vl_frete, 0) as vl_frete, " & _
	            "pf.codigo_tipo_frete, " & _
                "tCD.descricao, " & _
				"(" & _
					"(" & _
						"SELECT " & _
							"Coalesce(SUM(qtde*preco_NF),0) AS preco_NF" & _
						" FROM t_PEDIDO_ITEM i" &_ 
						" WHERE" & _
							" (i.pedido=p.pedido)" & _
					")" & _
					" - " & _
					"(" & _
						"SELECT " & _
							"Coalesce(SUM(qtde*preco_NF),0) AS preco_NF" & _
						" FROM t_PEDIDO_ITEM_DEVOLVIDO d" &_ 
						" WHERE" & _
							" (d.pedido=p.pedido)" & _
					")" & _
				") AS preco_NF, " & _
				"(" & _
					"SELECT " & _
						"Coalesce(SUM(qtde*cubagem),0) AS cubagem" & _
					" FROM t_PEDIDO_ITEM i" &_ 
					" WHERE" & _
						" (i.pedido=p.pedido)" & _
				") AS cubagem," & _
				"(" & _
					"SELECT " & _
						"Coalesce(SUM(qtde*qtde_volumes),0) AS qtde_volumes" & _
					" FROM t_PEDIDO_ITEM i" &_ 
					" WHERE" & _
						" (i.pedido=p.pedido)" & _
				") AS qtde_volumes," & _
				"(" & _
					"SELECT " & _
						"Coalesce(SUM(qtde*peso),0) AS peso" & _
					" FROM t_PEDIDO_ITEM i" &_ 
					" WHERE" & _
						" (i.pedido=p.pedido)" & _
				") AS peso," & _
				"(" & _
					"(" & _
						"SELECT " & _
							"Coalesce(SUM(qtde*preco_venda),0) AS preco_venda" & _
						" FROM t_PEDIDO_ITEM i" &_ 
						" WHERE" & _
							" (i.pedido=p.pedido)" & _
					")" & _
					" - " & _
					"(" & _
						"SELECT " & _
							"Coalesce(SUM(qtde*preco_venda),0) AS preco_venda" & _
						" FROM t_PEDIDO_ITEM_DEVOLVIDO d" &_ 
						" WHERE" & _
							" (d.pedido=p.pedido)" & _
					")" & _
				") AS preco_venda," & _
				"STUFF((" & _
					"SELECT" & _
						" ', ' + CONVERT(varchar(9), NFe_numero_NF) AS [text()]" & _
					" FROM " & _
						"(" & _
							"SELECT DISTINCT" & _
								" NFe_numero_NF" & _
							" FROM t_NFe_EMISSAO tNE" & _
							" WHERE" & _
								" (tNE.pedido = p.pedido)" & _
								" AND (tNE.tipo_NF = '1')" & _
								" AND (tNE.st_anulado = 0)" & _
								" AND (tNE.codigo_retorno_NFe_T1 = 1)" & _
						") tNEAux" & _
					" ORDER BY" & _
						" NFe_numero_NF" & _
					" FOR XML PATH('')" & _
				"), 1, 2, '') AS numeros_NF" & _
			" FROM t_PEDIDO p LEFT JOIN t_PEDIDO_FRETE pf ON (pf.pedido=p.pedido) INNER JOIN t_CLIENTE c ON (p.id_cliente=c.id)" & _
            " LEFT JOIN t_CODIGO_DESCRICAO tCD ON (tCD.grupo='" & GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE & "') AND (tCD.codigo = pf.codigo_tipo_frete)" & _
			" WHERE " & _
				s_where & _
            ") s " & _  
		    s_where_externo & _                      
            " ORDER BY" & _
				" s.transportadora_id," & _
	            " s.pedido," & _			
                " s.data"

	
	
'	CABEÇALHO
	cab_table = "<table cellspacing='0' cellpadding='0' class='MB'>" & chr(13)
	
	cab = _
		"	<TR style='background:azure' NOWRAP>" & chr(13) & _
		"		<TD class='MDTE tdPedido' align='left' valign='bottom' NOWRAP><span class='R titColL'>Pedido</span></TD>" &  chr(13) & _
		"		<TD class='MTD tdNF' align='left' valign='bottom'><span class='R titColL'>NF</span></TD>" & chr(13) & _
		"		<TD class='MTD tdCidade' align='left' valign='bottom'><span class='R titColL'>Cidade</span></TD>" & chr(13) & _
		"		<TD class='MTD tdCnpjCpfCliente' align='left' valign='bottom'><span class='R titColL'>CNPJ/CPF</span><br /><span class='R titColL'>Cliente</span></TD>" & chr(13) & _
		"		<TD class='MTD tdTipoFrete' align='right' valign='bottom'><span class='Rd titColR'>Tipo Frete</span></TD>" & chr(13) & _
		"		<TD class='MTD tdValorFrete' align='right' valign='bottom'><span class='Rd titColR'>Frete (" & SIMBOLO_MONETARIO & ")</span></TD>" & chr(13) & _
		"		<TD class='MTD tdPrecoNF' align='right' valign='bottom'><span class='Rd titColR'>Preço (" & SIMBOLO_MONETARIO & ")</span></TD>" & chr(13) & _
		"		<TD class='MTD tdPercPrecoNF' align='right' valign='bottom'><span class='Rd titColR'>%</span><br /><span class='Rd titColR'>Preço</span></TD>" & chr(13) & _
		"		<TD class='MTD tdCubagem' align='right' valign='bottom'><span class='Rd titColR'>Cubagem</span><br /><span class='Rd titColR'>(m3)</span></TD>" & chr(13) & _
		"		<TD class='MTD tdVolume' align='right' valign='bottom'><span class='Rd titColR'>Qtde</span><br /><span class='Rd titColR'>Volumes</span></TD>" & chr(13) & _
		"		<TD class='MTD tdPeso' align='right' valign='bottom'><span class='Rd titColR'>Peso</span><br /><span class='Rd titColR'>(kg)</span></TD>" & chr(13) & _
		"		<TD class='MTD tdPrecoVenda' align='right' valign='bottom'><span class='Rd titColR'>Venda (" & SIMBOLO_MONETARIO & ")</span></TD>" & chr(13) & _
		"		<TD class='MTD tdPercPrecoVenda' align='right' valign='bottom'><span class='Rd titColR'>%</span><br /><span class='Rd titColR'>Venda</span></TD>" & chr(13) & _
        "		<TD class='MTD tdPercPrecoVenda2' align='right' valign='bottom'><span class='Rd titColR'>%</span><br /><span class='Rd titColR'>Venda</span></TD>" & chr(13) & _
		"	</TR>" & chr(13)
	
'	LAÇO P/ LEITURA DO RECORDSET
	x = cab_table
	
	n_reg = 0
	intQtdeTotalPedidos = 0
	intQtdeTransportadoras = 0
	vlTotalFrete = 0
	vlTotalPrecoNF = 0
	vlTotalCubagem = 0
	intQtdeTotalVolumes = 0
	vlTotalPeso = 0
	vlTotalPrecoVenda = 0
    cont = 0

	pedidoAnterior = ""
	strTransportadoraAnterior = "XXXXXXXXXXXXXXXXXXXX"	
    lista_pedidos_total_geral = ""

		
	set r = cn.execute(s_sql)
	do while Not r.Eof

		strTransportadora = Trim("" & r("transportadora_id"))
        codigo_tipo_frete = Trim("" & r("codigo_tipo_frete"))
		if strTransportadora <> strTransportadoraAnterior then
			intQtdeTransportadoras = intQtdeTransportadoras + 1
		'	SUB-TOTAIS POR TRANSPORTADORA
		'	EXIBE SUB-TOTAL DA TRANSPORTADORA ANTERIOR?
			if intQtdeTotalPedidos > 0 then
				if (vlSubTotalFrete = 0) Or (vlSubTotalPrecoNF = 0) then
					vlSubTotalPercPrecoNF = 0
				else
				'	LEMBRANDO QUE O EXCEL EXIBE O VALOR MULTIPLICADO POR 100 (EX: 0.5 -> 50%)
					vlSubTotalPercPrecoNF = vlSubTotalFrete / vlSubTotalPrecoNF
					if Not blnSaidaExcel then vlSubTotalPercPrecoNF = 100 * vlSubTotalPercPrecoNF
					end if
				
				if blnSaidaExcel then 
					strVlSubTotalPercPrecoNF = formata_perc4dec(vlSubTotalPercPrecoNF)
				else 
					strVlSubTotalPercPrecoNF = formata_perc1dec(vlSubTotalPercPrecoNF)
					end if
				
				if (vlSubTotalFrete = 0) Or (vlSubTotalPrecoVenda = 0) then
					vlSubTotalPercPrecoVenda = 0
				else
				'	LEMBRANDO QUE O EXCEL EXIBE O VALOR MULTIPLICADO POR 100 (EX: 0.5 -> 50%)
					vlSubTotalPercPrecoVenda = vlSubTotalFrete / vlSubTotalPrecoVenda
					if Not blnSaidaExcel then vlSubTotalPercPrecoVenda = 100 * vlSubTotalPercPrecoVenda
					end if
				
				if blnSaidaExcel then
					strVlSubTotalPercPrecoVenda = formata_perc4dec(vlSubTotalPercPrecoVenda)
				else
					strVlSubTotalPercPrecoVenda = formata_perc1dec(vlSubTotalPercPrecoVenda)
					end if
				
				if intQtdeSubTotalPedidos > 1 then strPlural = "s" else strPlural = ""
                if intQtdeSubTotalFretes > 1 then strPluralFrete = "s" else strPluralFrete = ""
				x = x & _
					"	<TR style='background:ivory;'>" & chr(13) & _
					"		<TD colspan='2' class='MTE'>" & _
								"<P class='C' style=text-align:left;>" & formata_inteiro(intQtdeSubTotalPedidos) & " pedido" & strPlural & "</p>" & _
					"		</TD>" & chr(13) & _
                    "		<TD colspan='3' class='MTD '>" & _
								"<P class='C' style=text-align:left;>" & formata_inteiro(intQtdeSubTotalFretes) & " frete" & strPluralFrete & "</p>" & _
					"		</TD>" & chr(13) & _
					"		<TD class='MTD tdValorFrete'>" & _
								"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlSubTotalFrete) & "</p>" & _
					"		</TD>" & chr(13) & _
					"		<TD class='MTD tdPrecoNF'>" & _
								"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlSubTotalPrecoNF) & "</p>" & _
					"		</TD>" & chr(13) & _
					"		<TD class='MTD tdPercPrecoNF'>" & _
								"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlSubTotalPercPrecoNF & "</p>" & _
					"		</TD>" & chr(13) & _
					"		<TD class='MTD tdCubagem'>" & _
								"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlSubTotalCubagem) & "</p>" & _
					"		</TD>" & chr(13) & _
					"		<TD class='MTD tdVolume'>" & _
								"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(intQtdeSubTotalVolumes) & "</p>" & _
					"		</TD>" & chr(13) & _
					"		<TD class='MTD tdPeso'>" & _
								"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_FLOAT & chr(34) & ";'>" & formata_numero(vlSubTotalPeso, 1) & "</p>" & _
					"		</TD>" & chr(13) & _
					"		<TD class='MTD tdPrecoVenda'>" & _
								"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlSubTotalPrecoVenda) & "</p>" & _
					"		</TD>" & chr(13) & _
					"		<TD class='MTD tdPercPrecoVenda'>" & _
								"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlSubTotalPercPrecoVenda & "</p>" & _
					"		</TD>" & chr(13) & _
                    "		<TD class='MTD tdPercPrecoVenda2'>" & _
								"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlSubTotalPercPrecoVenda & "</p>" & _
					"		</TD>" & chr(13) & _
					"	</TR>" & chr(13)
				end if
			
			strTransportadoraAux = strTransportadora
			if strTransportadoraAux = "" then strTransportadoraAux = "SEM TRANSPORTADORA"
			if blnSaidaExcel then strTransportadoraAux = "Transportadora: " & strTransportadoraAux
			
			if intQtdeTotalPedidos > 0 then
			x = x & _
					"	<TR>" & chr(13) & _
					"		<TD colspan=14 class='MC'>&nbsp;</TD>" & chr(13) & _
					"	</TR>" & chr(13)
				end if
			
			x = x & _
				"	<TR style='background:azure'>" & chr(13) & _
				"		<TD colspan=14 class='MC ME MD'><P class='C' style='font-weight:bold;text-align:left;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & strTransportadoraAux & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			
		'	TÍTULO DAS COLUNAS
			x = x & _
				cab
			
			intQtdeSubTotalPedidos = 0
            intQtdeSubTotalFretes = 0
			vlSubTotalFrete = 0
			vlSubTotalPrecoNF = 0
			vlSubTotalCubagem = 0
			intQtdeSubTotalVolumes = 0
			vlSubTotalPeso = 0
			vlSubTotalPrecoVenda = 0
            pedidoAnterior = ""			
			end if
	
		n_reg = n_reg + 1
		intQtdeSubTotalFretes = intQtdeSubTotalFretes + 1
        intQtdeTotalFretes = intQtdeTotalFretes + 1

		x = x & "	<TR>" & chr(13)
		
		vlFrete = CCur(r("vl_frete"))
		vlPrecoNF = CCur(r("preco_NF"))
		vlCubagem = r("cubagem")
		intQtdeVolumes = r("qtde_volumes")
		vlPeso = r("peso")
		vlPrecoVenda = CCur(r("preco_venda"))
		qtde_pedido = r("qtde_pedido")
	'>  PEDIDO
    pedidoatual = Trim("" & r("pedido"))
		if blnSaidaExcel then
			s_aux = pedidoatual
		else
			s_aux = "<a href='javascript:fRELConcluir(" & chr(34) & pedidoatual & chr(34) & ")' title='clique para consultar o pedido'>" & _
					pedidoatual & _
					"</a>"
			end if
        if pedidoatual <> pedidoAnterior  AND qtde_pedido > 1  then
		    x = x & "   <TD ROWSPAN=" & qtde_pedido & " class='MDTE tdPedido'>" & _
							"<P class='C' style='text-align:left;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & _
								s_aux & _
							"</P>" & _
						"</TD>" & chr(13)         
        elseif pedidoatual = pedidoAnterior AND strTransportadora = strTransportadoraAnterior then
            
        else
            x = x & "		<TD class='MDTE tdPedido'>" & _
							"<P class='C' style='text-align:left;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & _
								s_aux & _
							"</P>" & _
						"</TD>" & chr(13)
        end if
	'>  NF
		' ALTERADO EM 28/02/2020		strNF = Trim("" & r("numeros_NF"))
		if rb_tipo_nf = "1" then
			'Somente NF de Remessa, quando houver
			strNF = Trim("" & r("obs_3"))
			if strNF = "" then strNF = Trim("" & r("obs_2"))
		else
			'NF de Fatura e de Remessa
			strNF = Trim("" & r("obs_2"))
			if (strNF <> "") And (Trim("" & r("obs_3")) <> "") then strNF = strNF & ", "
			strNF = strNF & Trim("" & r("obs_3"))
			end if

		if Not blnSaidaExcel then
			if strNF = "" then strNF = "&nbsp;"
			end if
		
		x = x & "		<TD class='MTD tdNF'>" & _
							"<p class='C' style='text-align:left;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & strNF & "</p>" & _
						"</TD>" & chr(13)
		
	'>  CIDADE
		if CInt(r("st_end_entrega")) = 0 then
			strCidade = Trim("" & r("cidade"))
			strUf = Trim("" & r("uf"))
			strCnpjCpfCliente = Trim("" & r("cliente_cnpj_cpf"))
		else
			strCidade = Trim("" & r("EndEtg_cidade"))
			strUf = Trim("" & r("EndEtg_uf"))
			strCnpjCpfCliente = Trim("" & r("EndEtg_cnpj_cpf"))
			end if
		
		if strCidade <> "" then strCidade = iniciais_em_maiusculas(strCidade)
		
		if (strCidade <> "") And (strUf <> "") then
			strCidadeUf = strCidade & " / " & strUf
		else
			strCidadeUf = strCidade & strUf
			end if
		
		x = x & "		<TD class='MTD tdCidade'>" & _
							"<p class='C' style='text-align:left;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & strCidadeUf & "</p>" & _
						"</TD>" & chr(13)
		
	'>  CNPJ/CPF CLIENTE
		if strCnpjCpfCliente <> "" then
			strCnpjCpfCliente = cnpj_cpf_formata(strCnpjCpfCliente)
		else
			strCnpjCpfCliente = "&nbsp;"
			end if
		x = x & "		<TD class='MTD tdCnpjCpfCliente'>" & _
							"<p class='C' style='text-align:left;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & strCnpjCpfCliente & "</p>" & _
						"</TD>" & chr(13)

    '>  TIPO DE FRETE
        x = x & "		<TD class='MTD tdValorFrete'>" & _
							"<P class='Cd' style='text-align:center'>" & r("descricao") & "</P>" & _
						"</TD>" & chr(13)

	'>  VALOR DO FRETE
		x = x & "		<TD class='MTD tdValorFrete'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlFrete) & "</P>" & _
						"</TD>" & chr(13)

	'>  PREÇO DE NF
		x = x & "		<TD class='MTD tdPrecoNF'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlPrecoNF) & "</P>" & _
						"</TD>" & chr(13)
					
	'>  PERCENTUAL SOBRE O PREÇO DE NF
		if (vlFrete = 0) Or (vlPrecoNF = 0) then
			percPrecoNF = 0
		else
		'	LEMBRANDO QUE O EXCEL EXIBE O VALOR MULTIPLICADO POR 100 (EX: 0.5 -> 50%)
			percPrecoNF = vlFrete / vlPrecoNF
			if Not blnSaidaExcel then percPrecoNF = 100 * percPrecoNF
			end if
		
		if blnSaidaExcel then
			strPercPrecoNF = formata_perc4dec(percPrecoNF)
		else
			strPercPrecoNF = formata_perc1dec(percPrecoNF)
			end if
		
		x = x & "		<TD class='MTD tdPercPrecoNF'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strPercPrecoNF & "</P>" & _
						"</TD>" & chr(13)
		
	'>  CUBAGEM
		x = x & "		<TD class='MTD tdCubagem'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlCubagem) & "</P>" & _
						"</TD>" & chr(13)
					
	'>  VOLUMES
		x = x & "		<TD class='MTD tdVolume'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(intQtdeVolumes) & "</P>" & _
						"</TD>" & chr(13)
					
	'>  PESO
		x = x & "		<TD class='MTD tdPeso'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_FLOAT & chr(34) & ";'>" & formata_numero(vlPeso, 1) & "</P>" & _
						"</TD>" & chr(13)
					
	'>  PREÇO DE VENDA
  
        if pedidoatual <> pedidoAnterior  AND qtde_pedido > 1  then

            x = x & "	<TD ROWSPAN=" & qtde_pedido & " class='MTD tdPrecoVenda'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlPrecoVenda) & "</P>" & _
						"</TD>" & chr(13)		            
        elseif pedidoatual = pedidoAnterior AND strTransportadora = strTransportadoraAnterior then

        else
            x = x & "	<TD class='MTD tdPrecoVenda'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlPrecoVenda) & "</P>" & _
						"</TD>" & chr(13)
        end if
		
		
	'>  PERCENTUAL SOBRE O PREÇO DE VENDA
		if (vlFrete = 0) Or (vlPrecoVenda = 0) then
			percPrecoVenda = 0
		else
		'	LEMBRANDO QUE O EXCEL EXIBE O VALOR MULTIPLICADO POR 100 (EX: 0.5 -> 50%)
			percPrecoVenda = vlFrete / vlPrecoVenda
			if Not blnSaidaExcel then percPrecoVenda = 100 * percPrecoVenda
			end if
		
		if blnSaidaExcel then
			strPercPrecoVenda = formata_perc4dec(percPrecoVenda)
		else
			strPercPrecoVenda = formata_perc1dec(percPrecoVenda)
			end if
		
		x = x & "		<TD class='MTD tdPercPrecoVenda'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strPercPrecoVenda & "</P>" & _
						"</TD>" & chr(13)
		
	'< PERCENTUAL SOBRE O PREÇO DE VENDA 2
        
        if pedidoatual <> pedidoAnterior then x = Replace(x,"#PERC#",strPercPrecoVendaSomado) 

        if pedidoatual <> pedidoAnterior  AND qtde_pedido > 1  then
               
		   x = x & "	<TD  ROWSPAN=" & qtde_pedido & " class='MTD tdPercPrecoVenda2'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'> #PERC# </P>" & _
						"</TD>" & chr(13)
           strPercPrecoVendaSomado = 0		
		   strPercPrecoVendaSomado = strPercPrecoVendaSomado + strPercPrecoVenda 
                 
        elseif pedidoatual = pedidoAnterior AND strTransportadora = strTransportadoraAnterior then
           strPercPrecoVendaSomado = strPercPrecoVendaSomado + strPercPrecoVenda  
                                         
        else 
                            
           x = x & "		<TD class='MTD tdPercPrecoVenda2'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strPercPrecoVenda & "</P>" & _
						"</TD>" & chr(13)		
		   strPercPrecoVendaSomado = 0
        end if
    

        x = x & "	</TR>" & chr(13)

    if Instr(x,"#PERC#") = 0 then
		if (n_reg mod 100) = 0 then          
			Response.Write x
			x = ""
			end if
	end if

        
		vlTotalFrete = vlTotalFrete + vlFrete
		vlSubTotalFrete = vlSubTotalFrete + vlFrete
		
		vlTotalCubagem = vlTotalCubagem + vlCubagem
		vlSubTotalCubagem = vlSubTotalCubagem + vlCubagem
		
		intQtdeTotalVolumes = intQtdeTotalVolumes + intQtdeVolumes
		intQtdeSubTotalVolumes = intQtdeSubTotalVolumes + intQtdeVolumes
		
		vlTotalPeso = vlTotalPeso + vlPeso
		vlSubTotalPeso = vlSubTotalPeso + vlPeso
		
        if Instr(lista_pedidos_total_geral, "|" & pedidoatual & "|") = 0 then
            vlTotalPrecoVendaAjustado = vlTotalPrecoVendaAjustado + vlPrecoVenda
			vlTotalPrecoNFAjustado = vlTotalPrecoNFAjustado + vlPrecoNF
            intQtdeTotalPedidos = intQtdeTotalPedidos + 1
		    
            lista_pedidos_total_geral = lista_pedidos_total_geral & "|" & pedidoatual & "|"
        end if
        
        if pedidoatual <> pedidoanterior then
         vlTotalPrecoVenda = vlTotalPrecoVenda + vlPrecoVenda
         vlTotalPrecoNF = vlTotalPrecoNF + vlPrecoNF
         vlSubTotalPrecoNF = vlSubTotalPrecoNF + vlPrecoNF
         vlSubTotalPrecoVenda = vlSubTotalPrecoVenda + vlPrecoVenda
         intQtdeSubTotalPedidos = intQtdeSubTotalPedidos + 1
		end if
        
		
		

        pedidoAnterior = pedidoatual
        strTransportadoraAnterior = strTransportadora
		r.MoveNext
		loop
	
	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeTotalPedidos = 0 then
		x = cab_table & _
			cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD class='MC MD ME ALERTA' colspan='14'><span class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</span></TD>" & chr(13) & _
			"	</TR>" & chr(13)
	else
	'	SUB-TOTAL DA ÚLTIMA TRANSPORTADORA
		if (vlSubTotalFrete = 0) Or (vlSubTotalPrecoNF = 0) then
			vlSubTotalPercPrecoNF = 0
		else
		'	LEMBRANDO QUE O EXCEL EXIBE O VALOR MULTIPLICADO POR 100 (EX: 0.5 -> 50%)
			vlSubTotalPercPrecoNF = vlSubTotalFrete / vlSubTotalPrecoNF
			if Not blnSaidaExcel then vlSubTotalPercPrecoNF = 100 * vlSubTotalPercPrecoNF
			end if
		
		if blnSaidaExcel then
			strVlSubTotalPercPrecoNF = formata_perc4dec(vlSubTotalPercPrecoNF)
		else
			strVlSubTotalPercPrecoNF = formata_perc1dec(vlSubTotalPercPrecoNF)
			end if
		
		if (vlSubTotalFrete = 0) Or (vlSubTotalPrecoVenda = 0) then
			vlSubTotalPercPrecoVenda = 0
		else
		'	LEMBRANDO QUE O EXCEL EXIBE O VALOR MULTIPLICADO POR 100 (EX: 0.5 -> 50%)
			vlSubTotalPercPrecoVenda = vlSubTotalFrete / vlSubTotalPrecoVenda
			if Not blnSaidaExcel then vlSubTotalPercPrecoVenda = 100 * vlSubTotalPercPrecoVenda
			end if
		
		if blnSaidaExcel then
			strVlSubTotalPercPrecoVenda = formata_perc4dec(vlSubTotalPercPrecoVenda)
		else
			strVlSubTotalPercPrecoVenda = formata_perc1dec(vlSubTotalPercPrecoVenda)
			end if
		
		if intQtdeSubTotalPedidos > 1 then strPlural = "s" else strPlural = ""
        if intQtdeSubTotalFretes > 1 then strPluralFrete = "s" else strPluralFrete = ""
		x = x & _
			"	<TR style='background:ivory;'>" & chr(13) & _
			"		<TD colspan='2' class='MTE'>" & _
								"<P class='C' style=text-align:left;>" & formata_inteiro(intQtdeSubTotalPedidos) & " pedido" & strPlural & "</p>" & _
			"		</TD>" & chr(13) & _
            "		<TD colspan='3' class='MTD '>" & _
						        "<P class='C' style=text-align:left;>" & formata_inteiro(intQtdeSubTotalFretes) & " frete" & strPluralFrete & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD tdValorFrete'>" & _
						"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlSubTotalFrete) & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD tdPrecoNF'>" & _
						"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlSubTotalPrecoNF) & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD tdPercPrecoNF'>" & _
						"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlSubTotalPercPrecoNF & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD tdCubagem'>" & _
						"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlSubTotalCubagem) & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD tdVolume'>" & _
						"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(intQtdeSubTotalVolumes) & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD tdPeso'>" & _
						"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_FLOAT & chr(34) & ";'>" & formata_numero(vlSubTotalPeso, 1) & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD tdPrecoVenda'>" & _
						"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlSubTotalPrecoVenda) & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD tdPercPrecoVenda'>" & _
						"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlSubTotalPercPrecoVenda & "</p>" & _
			"		</TD>" & chr(13) & _
            "		<TD class='MTD tdPercPrecoVenda2'>" & _
						"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlSubTotalPercPrecoVenda & "</p>" & _
			"		</TD>" & chr(13) & _
			"	</TR>" & chr(13)
		
	'	TOTAL GERAL
		if intQtdeTransportadoras > 1 then
			if (vlTotalFrete = 0) Or (vlTotalPrecoNF = 0) then
				vlTotalPercPrecoNF = 0
				vlTotalPercPrecoNFAjustado = 0
			else
			'	LEMBRANDO QUE O EXCEL EXIBE O VALOR MULTIPLICADO POR 100 (EX: 0.5 -> 50%)
				vlTotalPercPrecoNF = vlTotalFrete / vlTotalPrecoNF
				vlTotalPercPrecoNFAjustado = vlTotalFrete / vlTotalPrecoNFAjustado
				if Not blnSaidaExcel then
					vlTotalPercPrecoNF = 100 * vlTotalPercPrecoNF
					vlTotalPercPrecoNFAjustado = 100 * vlTotalPercPrecoNFAjustado
					end if
				end if
				
			if blnSaidaExcel then
				strVlTotalPercPrecoNF = formata_perc4dec(vlTotalPercPrecoNF)
				strVlTotalPercPrecoNFAjustado = formata_perc4dec(vlTotalPercPrecoNFAjustado)
			else
				strVlTotalPercPrecoNF = formata_perc1dec(vlTotalPercPrecoNF)
				strVlTotalPercPrecoNFAjustado = formata_perc1dec(vlTotalPercPrecoNFAjustado)
				end if
			
			if (vlTotalFrete = 0) Or (vlTotalPrecoVenda = 0) then
				vlTotalPercPrecoVenda = 0
                vlTotalPercPrecoVendaAjustado = 0
			else
			'	LEMBRANDO QUE O EXCEL EXIBE O VALOR MULTIPLICADO POR 100 (EX: 0.5 -> 50%)
				vlTotalPercPrecoVenda = vlTotalFrete / vlTotalPrecoVenda
                vlTotalPercPrecoVendaAjustado = vlTotalFrete / vlTotalPrecoVendaAjustado
				if Not blnSaidaExcel then 
                vlTotalPercPrecoVenda = 100 * vlTotalPercPrecoVenda
                vlTotalPercPrecoVendaAjustado = 100 * vlTotalPercPrecoVendaAjustado
                end if
				
            end if
			
			if blnSaidaExcel then
				strVlTotalPercPrecoVenda = formata_perc4dec(vlTotalPercPrecoVenda)
                strVlTotalPercPrecoVendaAjustado = formata_perc4dec(vlTotalPercPrecoVendaAjustado)
			else
				strVlTotalPercPrecoVenda = formata_perc1dec(vlTotalPercPrecoVenda)
                strVlTotalPercPrecoVendaAjustado = formata_perc1dec(vlTotalPercPrecoVendaAjustado)
				end if
			
			if intQtdeTotalPedidos > 1 then strPlural = "s" else strPlural = ""
            if intQtdeTotalFretes > 1 then strPluralFrete = "s" else strPluralFrete = ""
			x = x & _
				"	<TR>" & chr(13) & _
				"		<TD colspan=14 class='MC'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR>" & chr(13) & _
				"		<TD colspan=14><P class='C' style='font-weight:bold;text-align:left;'>TOTAL GERAL</P></TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR style='background:ivory;'>" & chr(13) & _
				"		<TD colspan='2' class='MTE'>" & _
							"<P class='C' style='text-align:left;'>" & formata_inteiro(intQtdeTotalPedidos) & " pedido" & strPlural & "</p>" & _
				"		</TD>" & chr(13) & _
                "		<TD colspan='3' class='MTD '>" & _
							"<P class='C' style='text-align:left;'>" & formata_inteiro(intQtdeTotalFretes) & " frete" & strPluralFrete & "</p>" & _
				"		</TD>" & chr(13) & _
				"		<TD class='MTD tdValorFrete'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlTotalFrete) & "</p>" & _
				"		</TD>" & chr(13) & _
				"		<TD class='MTD tdPrecoNF'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlTotalPrecoNF) & "</p>" & _
				"		</TD>" & chr(13) & _
				"		<TD class='MTD tdPercPrecoNF'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlTotalPercPrecoNF & "</p>" & _
				"		</TD>" & chr(13) & _
				"		<TD class='MTD tdCubagem'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlTotalCubagem) & "</p>" & _
				"		</TD>" & chr(13) & _
				"		<TD class='MTD tdVolume'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(intQtdeTotalVolumes) & "</p>" & _
				"		</TD>" & chr(13) & _
				"		<TD class='MTD tdPeso'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_FLOAT & chr(34) & ";'>" & formata_numero(vlTotalPeso, 1) & "</p>" & _
				"		</TD>" & chr(13) & _
				"		<TD class='MTD tdPrecoVenda'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlTotalPrecoVenda) & "</p>" & _
				"		</TD>" & chr(13) & _
				"		<TD colspan='2' class='MTD tdPercPrecoVenda'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlTotalPercPrecoVenda & "</p>" & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13) &_
                "   <TR style='background:ivory;'> " & chr(13) & _
                "		<TD colspan=5 class='MTD ME tdPrecoVenda2'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'> AJUSTADO </p>" & _
				"		</TD>" & chr(13) & _
				"		<TD class='MTD tdValorFrete'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlTotalFrete) & "</p>" & _
				"		</TD>" & chr(13) & _
                "		<TD class='MTD'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlTotalPrecoNFAjustado) & "</p>" & _
				"		</TD>" & chr(13) & _
                "       <TD class='MTD'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlTotalPercPrecoNFAjustado & "</p>" & _
				"		</TD>" & chr(13) & _
                "		<TD class='MTD' colspan='3'>" & _
							"&nbsp;" & _
				"		</TD>" & chr(13) & _
                "		<TD class='MTD tdPrecoVenda2'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlTotalPrecoVendaAjustado) & "</p>" & _
				"		</TD>" & chr(13) & _
                "       <TD colspan='2' class='MTD tdPrecoVenda2'>" & _
							"<P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlTotalPercPrecoVendaAjustado & "</p>" & _
				"		</TD>" & chr(13) & _
                "	</TR>" & chr(13)
			end if
		end if
	
  ' FECHA TABELA
	x = x & "</table>" & chr(13)
	
	Response.write x
	
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



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir(s_id){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=s_id;
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style type="text/css">
.tdPedido{
	width: 70px;
	}
.tdNF{
	width: 50px;
	}
.tdCidade{
	width: 160px;
	}
.tdCnpjCpfCliente{
	width: 105px;
	}
.tdValorFrete{
	width: 60px;
	}
.tdTipoFrete{
    width: 100px;
}
.tdPrecoNF{
	width: 60px;
	}
.tdPercPrecoNF{
	width: 40px;
	}
.tdPrecoVenda{
	width: 80px;
	}
.tdCubagem{
	width: 60px;
	}
.tdVolume{
	width: 60px;
	}
.tdPeso{
	width: 60px;
	}
.tdPercPrecoVenda{
	width: 40px;
	}
.tdPercPrecoVenda2{
	width: 40px;
	}
.titColL
{
	font-weight:bold;
	text-align:left;
}
.titColR
{
	font-weight:bold;
	text-align:right;
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
<table cellSpacing="0">
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

<form id="fREL" name="fREL" method="post" action="Pedido.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="<%=c_dt_entregue_inicio%>">
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>">
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">
<input type="hidden" name="c_uf" id="c_uf" value="<%=c_uf%>">
<input type="hidden" name="rb_frete_status" id="rb_frete_status" value="<%=rb_frete_status%>">
<input type="hidden" name="rb_tipo_nf" id="rb_tipo_nf" value="<%=rb_tipo_nf%>" />
<input type="hidden" name="c_empresa" id="c_empresa" value="<%=c_empresa%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="836" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Frete (Analítico)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s = monta_texto_filtro
	Response.Write s
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="836" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="836" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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
