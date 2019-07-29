<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.ContentType = "text/html" %>

<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelFreteSinteticoExec.asp
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
	if Not operacao_permitida(OP_CEN_REL_FRETE_SINTETICO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux
	dim c_dt_entregue_inicio, c_dt_entregue_termino
	dim c_transportadora, c_fabricante, c_loja, c_vendedor, c_indicador, c_uf
	dim s_nome_vendedor, s_nome_indicador, s_nome_loja, s_nome_fabricante, s_nome_transportadora
	dim rb_frete_status, rb_tipo_saida, c_tipo_frete

	alerta = ""

	rb_frete_status = Trim(Request.Form("rb_frete_status"))
	rb_tipo_saida = Ucase(Trim(Request.Form("rb_tipo_saida")))

	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))
	
	c_transportadora = Trim(Request.Form("c_transportadora"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_loja = retorna_so_digitos(Trim(Request.Form("c_loja")))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	c_indicador = Ucase(Trim(Request.Form("c_indicador")))
	c_uf = Ucase(Trim(Request.Form("c_uf")))
    c_tipo_frete = Trim(Request.Form("c_tipo_frete"))

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
			Response.AddHeader "Content-Disposition", "attachment; filename=FreteSintetico_" & formata_data_yyyymmdd(Now) & "_" & formata_hora_hhnnss(Now) & ".xls"
			Response.Write "<h2>Relatório de Frete (Sintético)</h2>"
			Response.Write monta_texto_filtro
			Response.Write "<br><br>"
			consulta_executa
			Response.End
			end if
		end if
	




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
		
	if Not blnSaidaExcel then s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>"

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
				   "<p class='N'>Período de Entrega:&nbsp;</p></td><td valign='top' width='99%'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
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
				   "<p class='N'>Transportadora:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
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
				   "<p class='N'>Fabricante:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
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
				   "<p class='N'>Loja:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
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
				   "<p class='N'>Vendedor:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
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
				   "<p class='N'>Indicador:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if
		
	s = c_uf
	if s = "" then s = "todos"
	
	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>UF:&nbsp;" & s & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>UF:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
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
				   "<p class='N'>Status do Frete:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if
	
	if blnSaidaExcel then
		if s_filtro <> "" then s_filtro = s_filtro & "<br>"
		s_filtro = s_filtro & "<span class='N'>Emissão:&nbsp;" & formata_data_hora(Now) & "</span>"
	else
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
				   "<p class='N'>" & formata_data_hora(Now) & "</p></td></tr>"
		s_filtro = s_filtro & "</table>"
		end if
	
	monta_texto_filtro = s_filtro
end function



' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
const MSO_NUMBER_FORMAT_TEXTO = "\@"
dim r
dim x
dim cab, cab_table
dim s_aux, s_sql, s_where
dim n_reg, n_reg_total
dim intLargTransportadora, intLargQtdePedidos, intLargValorFrete, intLargPrecoNF, intLargPercPrecoNF, intLargPrecoVenda, intLargPercPrecoVenda
dim vlFrete, vlPrecoNF, vlPrecoVenda
dim percPrecoVenda, percPrecoNF
dim vlTotalFrete, vlTotalPrecoNF, vlTotalPrecoVenda
dim vlTotalPercPrecoNF, vlTotalPercPrecoVenda
dim lngQtdeTotalPedidos, lngQtdePedidos
dim strTransportadora, intQtdeTransportadoras
dim strPercPrecoNF, strPercPrecoVenda, strVlTotalPercPrecoNF, strVlTotalPercPrecoVenda, s_where_externo

'	CRITÉRIOS DE RESTRIÇÃO
	s_where = "(p.st_entrega <> '" & ST_ENTREGA_CANCELADO & "')"
    s_where_externo = " WHERE (0=0)"
	
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
		s_where_externo = s_where_externo & " (tAux.transportadora_id = '" & c_transportadora & "')"
		end if

'	FILTRO: TIPO DE FRETE
	if c_tipo_frete <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (pf.codigo_tipo_frete = '" & c_tipo_frete & "')"
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
		s_where = s_where & _
			" (" & _
				"((p.st_end_entrega <> 0) And (p.EndEtg_uf = '" & c_uf & "'))" & _
				" OR " & _
				"((p.st_end_entrega = 0) And (c.uf = '" & c_uf & "'))" & _
			")"
		end if
		
'	FILTRO: STATUS DO FRETE
	if rb_frete_status <> "" then
		if s_where <> "" then s_where = s_where & " AND"
        if rb_frete_status = "0" then
            s_where = s_where & " (Coalesce(pf.vl_frete, 0) = 0)"
        elseif rb_frete_status = "1" then 
            s_where = s_where & " (Coalesce(pf.vl_frete, 0) > 0)"
        end if
		
	end if
		
'	MONTA SQL DE CONSULTA
	s_sql = "SELECT " & _
				"transportadora_id, " & _
				"Coalesce(COUNT(*),0) AS qtde_pedidos, " & _
				"Coalesce(SUM(preco_NF),0) AS preco_NF, " & _
				"Coalesce(SUM(preco_venda),0) AS preco_venda, " & _
				"Coalesce(SUM(vl_frete),0) AS frete_valor" & _
			" FROM " & _
				"(" & _
					"SELECT " & _
						"Coalesce(pf.transportadora_id, p.transportadora_id) AS transportadora_id, " & _
						"p.pedido, " & _
						"pf.vl_frete, " & _
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
						") AS preco_venda " & _
					" FROM t_PEDIDO p LEFT JOIN t_PEDIDO_FRETE pf ON (pf.pedido=p.pedido) INNER JOIN t_CLIENTE c ON (p.id_cliente=c.id)" & _
					" WHERE " & _
						s_where & _
				") tAux" & _
                s_where_externo & _            
			" GROUP BY" & _
				" tAux.transportadora_id " & _
			" ORDER BY" & _
				" frete_valor DESC, qtde_pedidos DESC"
	
	
  ' CABEÇALHO
	intLargTransportadora = 120
	intLargQtdePedidos = 70
	intLargValorFrete = 80
	intLargPrecoNF = 80
	intLargPercPrecoNF = 40
	intLargPrecoVenda = 80
	intLargPercPrecoVenda = 40
	
	cab_table = "<TABLE CellSpacing=0 CellPadding=0 class='MB'>" & chr(13)
	
	cab = _
		"	<TR style='background:azure' NOWRAP>" & chr(13) & _
		"		<TD class='MDTE' valign='bottom' NOWRAP><P style='width:" & CStr(intLargTransportadora) & "px;font-weight:bold;' class='R'>Transportadora</P></TD>" &  chr(13) & _
		"		<TD class='MTD' valign='bottom'><P style='width:" & CStr(intLargQtdePedidos) & "px;font-weight:bold;text-align:right;' class='Rd'>Qtde Pedidos</P></TD>" & chr(13) & _
		"		<TD class='MTD' valign='bottom'><P style='width:" & CStr(intLargValorFrete) & "px;font-weight:bold;text-align:right;' class='Rd'>Frete (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		"		<TD class='MTD' valign='bottom'><P style='width:" & CStr(intLargPrecoNF) & "px;font-weight:bold;text-align:right;' class='Rd'>Preço (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		"		<TD class='MTD' valign='bottom'><P style='width:" & CStr(intLargPercPrecoNF) & "px;font-weight:bold;text-align:right;' class='Rd'>% Preço</P></TD>" & chr(13) & _
		"		<TD class='MTD' valign='bottom'><P style='width:" & CStr(intLargPrecoVenda) & "px;font-weight:bold;text-align:right;' class='Rd'>Venda (" & SIMBOLO_MONETARIO & ")</P></TD>" & chr(13) & _
		"		<TD class='MTD' valign='bottom'><P style='width:" & CStr(intLargPercPrecoVenda) & "px;font-weight:bold;text-align:right;' class='Rd'>% Venda</P></TD>" & chr(13) & _
		"	</TR>" & chr(13)
	
'	LAÇO P/ LEITURA DO RECORDSET
	x = cab_table & _
		cab
	
	n_reg = 0
	n_reg_total = 0
	vlTotalFrete = 0
	vlTotalPrecoNF = 0
	vlTotalPrecoVenda = 0
	lngQtdeTotalPedidos = 0
	intQtdeTransportadoras = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR>" & chr(13)
		
		vlFrete = CCur(r("frete_valor"))
		vlPrecoNF = CCur(r("preco_NF"))
		vlPrecoVenda = CCur(r("preco_venda"))
		lngQtdePedidos = CLng(r("qtde_pedidos"))
		
	'>  TRANSPORTADORA
		strTransportadora = Trim("" & r("transportadora_id"))
		if strTransportadora <> "" then intQtdeTransportadoras = intQtdeTransportadoras + 1
		if strTransportadora = "" then strTransportadora = "&nbsp;"
		x = x & "		<TD class='MDTE'>" & _
							"<P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & _
								strTransportadora & _
							"</P>" & _
						"</TD>" & chr(13)

	'>  QTDE PEDIDOS
		x = x & "		<TD class='MTD'>" & _
							"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & _
								formata_inteiro(lngQtdePedidos) & _
							"</P>" & _
						"</TD>" & chr(13)

	'>  VALOR DO FRETE
		x = x & "		<TD class='MTD'>" & _
							"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & _
								formata_moeda(vlFrete) & _
							"</P>" & _
						"</TD>" & chr(13)

	'>  PREÇO DE NF
		x = x & "		<TD class='MTD'>" & _
							"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & _
								formata_moeda(vlPrecoNF) & _
							"</P>" & _
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
			
		x = x & "		<TD class='MTD'>" & _
							"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & _
								strPercPrecoNF & _
							"</P>" & _
						"</TD>" & chr(13)

	'>  PREÇO DE VENDA
		x = x & "		<TD class='MTD'>" & _
							"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & _
								formata_moeda(vlPrecoVenda) & _
							"</P>" & _
						"</TD>" & chr(13)
			
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
			
		x = x & "		<TD class='MTD'>" & _
							"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & _
								strPercPrecoVenda & _
							"</P>" & _
						"</TD>" & chr(13)
		
		x = x & "	</TR>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		vlTotalFrete = vlTotalFrete + vlFrete
		vlTotalPrecoNF = vlTotalPrecoNF + vlPrecoNF
		vlTotalPrecoVenda = vlTotalPrecoVenda + vlPrecoVenda
		lngQtdeTotalPedidos = lngQtdeTotalPedidos + lngQtdePedidos
		
		r.MoveNext
		loop


  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & _
			cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD class='MC MD ME' colspan=7><P class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
	else
		if (vlTotalFrete = 0) Or (vlTotalPrecoNF = 0) then
			vlTotalPercPrecoNF = 0
		else
			vlTotalPercPrecoNF = vlTotalFrete / vlTotalPrecoNF
			if Not blnSaidaExcel then vlTotalPercPrecoNF = 100 * vlTotalPercPrecoNF
			end if

		if blnSaidaExcel then
			strVlTotalPercPrecoNF = formata_perc4dec(vlTotalPercPrecoNF)
		else
			strVlTotalPercPrecoNF = formata_perc1dec(vlTotalPercPrecoNF)
			end if
			
		if (vlTotalFrete = 0) Or (vlTotalPrecoVenda = 0) then
			vlTotalPercPrecoVenda = 0
		else
			vlTotalPercPrecoVenda = vlTotalFrete / vlTotalPrecoVenda
			if Not blnSaidaExcel then vlTotalPercPrecoVenda = 100 * vlTotalPercPrecoVenda
			end if
		
		if blnSaidaExcel then
			strVlTotalPercPrecoVenda = formata_perc4dec(vlTotalPercPrecoVenda)
		else
			strVlTotalPercPrecoVenda = formata_perc1dec(vlTotalPercPrecoVenda)
			end if
		
	'	TOTAIS
		x = x & _
			"	<TR>" & chr(13) & _
			"		<TD colspan=7 class='MC'>&nbsp;</TD>" & chr(13) & _
			"	</TR>" & chr(13) & _
			"	<TR>" & chr(13) & _
			"		<TD colspan=7><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>TOTAL</P></TD>" & chr(13) & _
			"	</TR>" & chr(13) & _
			"	<TR style='background:ivory;'>" & chr(13) & _
			"		<TD class='MTD ME'>" & _
						"<P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & formata_inteiro(intQtdeTransportadoras) & " transportadoras</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD'>" & _
						"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & formata_inteiro(lngQtdeTotalPedidos) & " pedidos</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD'>" & _
						"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlTotalFrete) & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD'>" & _
						"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlTotalPrecoNF) & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD'>" & _
						"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlTotalPercPrecoNF & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD'>" & _
						"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vlTotalPrecoVenda) & "</p>" & _
			"		</TD>" & chr(13) & _
			"		<TD class='MTD'>" & _
						"<P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_PERC & chr(34) & ";'>" & strVlTotalPercPrecoVenda & "</p>" & _
			"		</TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
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



<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir(s_id){
	window.status = "Aguarde ...";
	fREL.id_selecionado.value=s_id;
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

<form id="fREL" name="fREL" method="post" action="">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="id_selecionado" id="id_selecionado" value=''>
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="<%=c_dt_entregue_inicio%>">
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>">
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">
<input type="hidden" name="c_uf" id="c_uf" value="<%=c_uf%>">
<input type="hidden" name="rb_frete_status" id="rb_frete_status" value="<%=rb_frete_status%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Frete (Sintético)</span>
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
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
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
