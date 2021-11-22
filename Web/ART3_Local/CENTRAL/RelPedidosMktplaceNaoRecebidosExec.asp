<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelPedidosMktplaceNaoRecebidosExec.asp
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
	if Not operacao_permitida(OP_CEN_REL_PEDIDO_MARKETPLACE_NAO_RECEBIDO_CLIENTE, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim url_back, strUrlBotaoVoltar
	url_back = Trim(Request("url_back"))
	if url_back <> "" then
		strUrlBotaoVoltar = "RelPedidosMktplaceNaoRecebidos.asp?url_back=X&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	else
		strUrlBotaoVoltar = "javascript:history.back()"
		end if

	dim alerta
	dim s, s_aux, s_filtro
	dim c_dt_entregue_inicio, c_dt_entregue_termino
	dim c_transportadora, c_loja, c_grupo_pedido_origem, c_pedido_origem
	dim s_nome_loja, s_nome_transportadora, s_nome_grupo_pedido_origem, s_nome_pedido_origem
	dim qtde_total_pedidos

	alerta = ""

	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))
	c_transportadora = Trim(Request.Form("c_transportadora"))
	c_loja = retorna_so_digitos(Trim(Request.Form("c_loja")))
    c_grupo_pedido_origem = Trim(Request.Form("c_grupo_pedido_origem"))
    c_pedido_origem = Trim(Request.Form("c_pedido_origem"))

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
		s_nome_grupo_pedido_origem = ""
		if c_grupo_pedido_origem <> "" then
			s = "SELECT descricao FROM t_CODIGO_DESCRICAO WHERE (grupo='PedidoECommerce_Origem_Grupo' AND codigo='" & c_grupo_pedido_origem & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_grupo_pedido_origem & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_grupo_pedido_origem = Trim("" & rs("descricao"))
				end if
			end if
		end if

    if alerta = "" then
		s_nome_pedido_origem = ""
		if c_pedido_origem <> "" then
			s = "SELECT descricao FROM t_CODIGO_DESCRICAO WHERE (grupo='PedidoECommerce_Origem' AND codigo='" & c_pedido_origem & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_pedido_origem & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_pedido_origem = Trim("" & rs("descricao"))
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
		
	dim s_sessionToken
	s_sessionToken = ""
	s = "SELECT Convert(varchar(36), SessionTokenModuloCentral) AS SessionTokenModuloCentral FROM t_USUARIO WHERE (usuario = '" & usuario & "')"
	if rs.State <> 0 then rs.Close
	rs.open s, cn
	if Not rs.Eof then s_sessionToken = Trim("" & rs("SessionTokenModuloCentral"))
	if rs.State <> 0 then rs.Close

'	Período de consulta está restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	strMinDtInicialFiltroPeriodoYYYYMMDD = ""
	strMinDtInicialFiltroPeriodoDDMMYYYY = ""
	' RESTRIÇÃO DESABILITADA A PEDIDO DA KARINA EM 24/07/2017
	if False And operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
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
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim x
dim cab, cab_table
dim s_sql, s_where
dim n_reg
dim strTransportadora, strTransportadoraAnterior, strTransportadoraAux, strPlural, strObs2, strObs3
dim strCidade, strUf, strCidadeUf
dim intQtdeTotalPedidos, intQtdeTransportadoras
dim intQtdeSubTotalPedidos, s_grupo_origem
dim s_link_rastreio
dim nColSpan, s_cor, s_dias

	nColSpan = 11

'	CRITÉRIOS DE RESTRIÇÃO
	s_where = "(p.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
              " AND (p.PedidoRecebidoStatus = 0)" & _
			  " AND (p.MarketplacePedidoRecebidoRegistrarStatus = " & COD_ST_PEDIDO_RECEBIDO_NAO & ")" & _
              " AND (p.marketplace_codigo_origem IS NOT NULL) AND (LEN(Coalesce(p.marketplace_codigo_origem,'')) > 0)"
	
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
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (p.transportadora_id = '" & c_transportadora & "')"
		end if

'   FILTRO: ORIGEM DO PEDIDO (GRUPO)
    s = ""
    if c_grupo_pedido_origem <> "" then
        s_grupo_origem = "SELECT codigo FROM t_CODIGO_DESCRICAO WHERE (codigo_pai = '" & c_grupo_pedido_origem & "') AND grupo='PedidoECommerce_Origem'"
        if rs.State <> 0 then rs.Close
	    rs.open s_grupo_origem, cn
		if rs.Eof then
            alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_grupo_pedido_origem & " NÃO EXISTE."
        else
            do while Not rs.Eof
                if s <> "" then s = s & ", "
                s = s & "'" & rs("codigo") & "'"      
                rs.MoveNext
            loop
             s = " p.marketplace_codigo_origem IN (" & s & ")"
        end if
        if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
		end if
    end if

'	FILTRO: ORIGEM DO PEDIDO
	if c_pedido_origem <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (p.marketplace_codigo_origem = '" & c_pedido_origem & "')"
		end if

'	FILTRO: LOJA
	if c_loja <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (p.numero_loja = " & c_loja & ")"
		end if

'	MONTA SQL DE CONSULTA
	s_sql = "SELECT" & _
				" p.transportadora_id," & _
				" p.entregue_data," & _
				" p.data," & _
				" p.pedido," & _
				" p.obs_2," & _
				" p.obs_3," & _
				" p.loja," & _
				" p.PrevisaoEntregaTranspData," & _
				" p.st_end_entrega," & _
				" p.EndEtg_cidade," & _
				" p.EndEtg_uf,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" p.endereco_cidade AS cidade," & _
				" p.endereco_uf AS uf," & _
				" p.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas"
	else
		s_sql = s_sql & _
				" c.cidade," & _
				" c.uf," & _
				" c.nome_iniciais_em_maiusculas"
		end if

	s_sql = s_sql & _
			" FROM t_PEDIDO p INNER JOIN t_CLIENTE c ON (p.id_cliente=c.id)" & _
			" WHERE " & _
				s_where & _
			" ORDER BY" & _
				" p.transportadora_id," & _
				" p.entregue_data," & _
				" p.data," & _
				" p.pedido"
	
	
  ' CABEÇALHO
	cab_table = "<table cellspacing=0 cellpadding=0 class='MB'>" & chr(13)
	
	cab = _
		"	<tr style='background:azure' nowrap>" & chr(13) & _
		"		<td class='MDTE tdDataEntrega' align='center' valign='bottom' nowrap><span class='Rc'>Data Coleta</span></td>" & chr(13) & _
		"		<td class='MTD tdPrevEtg' align='center' valign='bottom' nowrap><span class='Rc'>Prev Etg</span></td>" & chr(13) & _
		"		<td class='MTD tdAtraso' align='right' valign='bottom' nowrap><span class='Rd' style='font-weight:bold;'>Atraso</span></td>" & chr(13) & _
		"		<td class='MTD tdRecebido' align='center' valign='bottom' nowrap><span class='Rc'>Receb</span></td>" &  chr(13) & _
		"		<td class='MTD tdDtRecebido' align='center' valign='bottom' nowrap><span class='Rc'>Data Recebido</span></td>" &  chr(13) & _
		"		<td class='MTD tdPedido' valign='bottom' nowrap><span class='R'>Pedido</span></td>" &  chr(13) & _
		"		<td class='MTD tdCidade' valign='bottom' nowrap><span class='R' style='text-align:left;'>Cidade</span></td>" &  chr(13) & _
		"		<td class='MTD tdObs2' valign='bottom' nowrap><span class='R' style='text-align:left;'>Nota Fiscal</span></td>" &  chr(13) & _
		"		<td class='MTD tdObs2' valign='bottom'><span class='R' style='text-align:left;'>NF Simples Remessa</span></td>" &  chr(13) & _
		"		<td class='MTD tdLoja' align='center' valign='bottom' nowrap><span class='Rc'>Loja</span></td>" &  chr(13) & _
		"		<td class='MTD tdCliente' valign='bottom' ><span style='font-weight:bold; text-align:left;' class='R'>Cliente</span></td>" & chr(13) & _
		"	</tr>" & chr(13)
	
'	LAÇO P/ LEITURA DO RECORDSET
	x = cab_table
	
	n_reg = 0
	intQtdeTotalPedidos = 0
	intQtdeTransportadoras = 0
	
	strTransportadoraAnterior = "XXXXXXXXXXXXXXXXXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

		strTransportadora = Trim("" & r("transportadora_id"))
		if strTransportadora <> strTransportadoraAnterior then
			intQtdeTransportadoras = intQtdeTransportadoras + 1
		'	SUB-TOTAIS POR TRANSPORTADORA
		'	EXIBE SUB-TOTAL DA TRANSPORTADORA ANTERIOR?
			if intQtdeTotalPedidos > 0 then
				if intQtdeSubTotalPedidos > 1 then strPlural = "s" else strPlural = ""
				x = x & _
					"	<tr style='background:ivory;'>" & chr(13) & _
					"		<td class='MDTE' colspan='" & CStr(nColSpan) & "'>" & _
								"<span class='C' style='text-align:left;'>" & formata_inteiro(intQtdeSubTotalPedidos) & " pedido" & strPlural & "</span>" & _
					"		</td>" & chr(13) & _
					"	</tr>" & chr(13)
				end if
			
			strTransportadoraAux = strTransportadora
			if strTransportadoraAux = "" then strTransportadoraAux = "SEM TRANSPORTADORA"
			
			if intQtdeTotalPedidos > 0 then
			x = x & _
					"	<tr>" & chr(13) & _
					"		<td colspan='" & CStr(nColSpan) & "' class='MC'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13)
				end if
				
			x = x & _
				"	<tr style='background:azure'>" & chr(13) & _
				"		<td colspan='" & CStr(nColSpan) & "' class='MC ME MD'><span class='C'>" & strTransportadoraAux & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
			
		'	TÍTULO DAS COLUNAS
			x = x & _
				cab
			
			intQtdeSubTotalPedidos = 0
			strTransportadoraAnterior = strTransportadora
			end if
	
		n_reg = n_reg + 1
		intQtdeTotalPedidos = intQtdeTotalPedidos + 1
		intQtdeSubTotalPedidos = intQtdeSubTotalPedidos + 1

		x = x & "	<tr onmouseover='realca_cor_mouse_over(this);' onmouseout='realca_cor_mouse_out(this);'>" & chr(13)
		
	'>  DATA DE COLETA (RÓTULO ANTIGO: DATA DA ENTREGA)
		x = x & "		<td class='MDTE tdDataEntrega' align='center'>" & _
							"<input type='text' class='Cc cDtColeta' style='border:0;width:70px;' name='c_dt_coleta' id='c_dt_coleta' " & _
							"value = '" & formata_data(r("entregue_data")) & "' readonly" & _
							">" & _
						"</td>" & chr(13)

	'>  DATA PREVISÃO DE ENTREGA
		s_cor = "black"
		if Trim("" & r("PrevisaoEntregaTranspData")) <> "" then
			if r("PrevisaoEntregaTranspData") > Date then
				s_cor = "green"
			elseif r("PrevisaoEntregaTranspData") < Date then
				s_cor = "red"
				end if
			end if

		x = x & "		<td class='MTD tdPrevEtg' align='center'>" & _
				"<input type='text' class='Cc cDtPrevEtg' style='border:0;width:60px;color:" & s_cor & ";' name='c_dt_prev_etg' id='c_dt_prev_etg' " & _
				"value = '" & formata_data(r("PrevisaoEntregaTranspData")) & "' readonly" & _
				" />" & _
				"</td>" & chr(13)

	'>  ATRASO (EM DIAS)
		s_dias = ""
		if Trim("" & r("PrevisaoEntregaTranspData")) <> "" then
			if r("PrevisaoEntregaTranspData") < Date then
				s_dias = DateDiff("d", r("PrevisaoEntregaTranspData"), Date)
				end if
			end if

		x = x & "		<td class='MTD tdAtraso' align='center'>" & _
				"<input type='text' class='Cd cAtraso' style='border:0;width:30px;color:" & s_cor & ";' name='c_atraso' id='c_atraso' " & _
				"value = '" & s_dias & "' readonly" & _
				" />" & _
				"</td>" & chr(13)

	'>  RECEBIDO
		x = x & "		<td class='MTD tdCkb tdRecebido' align='center'>" & _
							"<input type='checkbox' name='ckb_recebido' id='ckb_recebido' class='Cc CKB_REC' " & _
								" value='" & Trim("" & r("pedido")) & "'" & _
								">" & _
						"</td>" & chr(13)

    '>  DATA RECEBIDO
		x = x & "		<td class='MTD tdDtRecebido' align='center'>" & _
							"<input class='Cc' name='c_dt_recebimento_pedido' id='c_dt_recebimento_pedido' maxlength='10' style='width:70px;margin-left:2px;' onblur='if (!isDate(this)) {alert(""Data inválida!""); this.focus();}' disabled=disabled />" & _
						"</td>" & chr(13)

	'>  PEDIDO
		x = x & "		<td class='MTD tdPedido'>" & _
							"<span class='Cc'>" & _
								"<a href='javascript:fPEDConsulta(" & chr(34) & r("pedido") & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o pedido'>" & _
								Trim("" & r("pedido")) & _
								"</a>" & _
							"</span>" & _
						"</td>" & chr(13)

	'>  CIDADE
		if CInt(r("st_end_entrega")) = 0 then
			strCidade = Trim("" & r("cidade"))
			strUf = Trim("" & r("uf"))
		else
			strCidade = Trim("" & r("EndEtg_cidade"))
			strUf = Trim("" & r("EndEtg_uf"))
			end if
			
		if strCidade <> "" then strCidade = iniciais_em_maiusculas(strCidade)
		
		if (strCidade <> "") And (strUf <> "") then
			strCidadeUf = strCidade & " / " & strUf
		else
			strCidadeUf = strCidade & strUf
			end if
			
		x = x & "		<td class='MTD tdCidade'>" & _
							"<span class='C' style='text-align:left;'>" & _
								strCidadeUf & _
							"</span>" & _
						"</td>" & chr(13)

	'>  Nº NOTA FISCAL
        s_link_rastreio = ""
		strObs2 = Trim("" & r("obs_2"))
		if strObs2 = "" then
            strObs2 = "&nbsp;"
        else
            s_link_rastreio = monta_link_rastreio(Trim("" & r("pedido")), strObs2, Trim("" & r("transportadora_id")), Trim("" & r("loja")))
        end if
        if s_link_rastreio <> "" then s_link_rastreio = "&nbsp;" & s_link_rastreio
		x = x & "		<td class='MTD tdObs2'>" & _
							"<span class='C' style='text-align:left;'>" & _
								"<a href='javascript:fRELConcluir(" & chr(34) & r("pedido") & chr(34) & ")' title='clique para consultar o pedido'>" & _
								strObs2 & s_link_rastreio & _
								"</a>" & _
							"</span>" & _
						"</td>" & chr(13)

    '>  Nº NOTA FISCAL SIMPLES REMESSA
		s_link_rastreio = ""
		strObs3 = Trim("" & r("obs_3"))
		if strObs3 = "" then
            strObs3 = "&nbsp;"
        else
            s_link_rastreio = monta_link_rastreio(Trim("" & r("pedido")), strObs3, Trim("" & r("transportadora_id")), Trim("" & r("loja")))
        end if
        if s_link_rastreio <> "" then s_link_rastreio = "&nbsp;" & s_link_rastreio
		x = x & "		<td class='MTD tdObs2'>" & _
							"<span class='C' style='text-align:left;'>" & _
								"<a href='javascript:fRELConcluir(" & chr(34) & r("pedido") & chr(34) & ")' title='clique para consultar o pedido'>" & _
								strObs3 & s_link_rastreio & _
								"</a>" & _
							"</span>" & _
						"</td>" & chr(13)
						
	'>  Nº LOJA
		x = x & "		<td class='MTD tdLoja' align='center'>" & _
							"<span class='C' style='text-align:left;'>" & _
								Trim("" & r("loja")) & _
							"</span>" & _
						"</td>" & chr(13)

	'>  NOME DO CLIENTE
		x = x & "		<td class='MTD tdCliente'>" & _
							"<span class='C' style='text-align:left;'>" & _
								Trim("" & r("nome_iniciais_em_maiusculas")) & _
							"</span>" & _
						"</td>" & chr(13)

		x = x & "	</tr>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		r.MoveNext
		loop

	qtde_total_pedidos = intQtdeTotalPedidos

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeTotalPedidos = 0 then
		x = cab_table & _
			cab & _
			"	<tr nowrap>" & chr(13) & _
			"		<td class='MC MD ME ALERTA' colspan='" & CStr(nColSpan) & "' align='center'><span class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</span></td>" & chr(13) & _
			"	</tr>" & chr(13)
	else
	'	SUB-TOTAL DA ÚLTIMA TRANSPORTADORA
		if intQtdeSubTotalPedidos > 1 then strPlural = "s" else strPlural = ""
		x = x & _
			"	<tr style='background:ivory;'>" & chr(13) & _
			"		<td class='MDTE' colspan='" & CStr(nColSpan) & "' align='left'>" & _
						"<span class='C' style='text-align:left;'>" & formata_inteiro(intQtdeSubTotalPedidos) & " pedido" & strPlural & "</span>" & _
			"		</td>" & chr(13) & _
			"	</tr>" & chr(13)
		
	'	TOTAL GERAL
		if intQtdeTransportadoras > 1 then
			if intQtdeTotalPedidos > 1 then strPlural = "s" else strPlural = ""
			x = x & _
				"	<tr>" & chr(13) & _
				"		<td colspan='" & CStr(nColSpan) & "' class='MC' align='left'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr>" & chr(13) & _
				"		<td colspan='" & CStr(nColSpan) & "' align='left'><span class='C' style='text-align:left;'>TOTAL GERAL</span></td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr style='background:ivory;'>" & chr(13) & _
				"		<td class='MDTE' colspan='" & CStr(nColSpan) & "' align='left'>" & _
							"<span class='C' style='text-align:left;'>" & formata_inteiro(intQtdeTotalPedidos) & " pedido" & strPlural & "</span>" & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13)
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



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__RASTREIO_VIA_WEBAPI_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	var urlBaseSsw = '<%=URL_SSW_BASE%>';
	var urlWebApiRastreio;
	var serverVariableUrl;
	serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
	serverVariableUrl = serverVariableUrl.toUpperCase();
	serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));
	urlWebApiRastreio = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/GetData/PageContentViaHttpGet';

	$(document).ready(function () {
	    $("#c_dt_recebimento").hUtilUI('datepicker_padrao');

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
		$(".CKB_REC:enabled").each(function() {
			if ($(this).is(":checked")) {
			    $(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
			    $(this).parents("td.tdCkb ").siblings(".tdDtRecebido").children("#c_dt_recebimento_pedido").removeAttr("disabled");
			}
			else {
			    $(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
			    $(this).parents("td.tdCkb ").siblings(".tdDtRecebido").children("#c_dt_recebimento_pedido").attr("disabled", "disabled");
			}
		})

		// EVENTO P/ REALÇAR OU NÃO CONFORME SE MARCA/DESMARCA O CHECKBOX
		$(".CKB_REC:enabled").click(function() {
			if ($(this).is(":checked")) {
			    $(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
			    $(this).parents("td.tdCkb ").siblings(".tdDtRecebido").children("#c_dt_recebimento_pedido").removeAttr("disabled");
			}
			else {
			    $(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
			    $(this).parents("td.tdCkb ").siblings(".tdDtRecebido").children("#c_dt_recebimento_pedido").attr("disabled", "disabled");
			}
		})
	});
</script>

<script language="JavaScript" type="text/javascript">
var windowScrollTopAnterior;
window.status='Aguarde, executando a consulta ...';

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

function realca_cor_mouse_over(c) {
	c.style.backgroundColor = 'palegreen';
}

function realca_cor_mouse_out(c) {
	c.style.backgroundColor = '';
}

function fPEDConsulta(id_pedido, usuario) {
	windowScrollTopAnterior = $(window).scrollTop();
	sizeDivPedidoConsulta();
	$("#iframePedidoConsulta").attr("src", "PedidoConsultaView.asp?pedido_selecionado=" + id_pedido + "&pedido_selecionado_inicial=" + id_pedido + "&usuario=" + usuario);
	$("#divPedidoConsulta").fadeIn();
}

function fRELConcluir(s_id){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=s_id;
	fREL.submit(); 
}

function VerificaDtRecebimento(campodata) {
	if (!isDate(campodata)) {
		alert('Data inválida!');
		campodata.focus();
	}
}

function fRELGravaDados(f) {
	var i, intQtdeTratados;
	var s, p, d, t, hasDtR, dtR;

	hasDtR = false;
	intQtdeTratados = 0;
	for (i = 1; i < f.ckb_recebido.length; i++) {
	    if (f.ckb_recebido[i].checked) {
	        p = f.ckb_recebido[i].value;
	        d = f.c_dt_coleta[i].value;
	        t = f.c_dt_recebimento_pedido[i].value;
	        if (t != "") {
	            if (converte_data(t) < converte_data(d)) {
	                alert("A data de recebimento informada no pedido " + p + " é anterior à data de coleta!!");
	                f.c_dt_recebimento_pedido[i].focus();
	                return;
	            }
	            if (converte_data(t) > converte_data(f.c_dt_hoje.value)) {
	                alert("A data de recebimento informada no pedido " + p + " não pode ser uma data futura!!");
	                f.c_dt_recebimento_pedido[i].focus();
	                return;
	            }
	        }
	        else {
	            hasDtR = true;
	        }
	        intQtdeTratados++;
	    }
	}

	if (intQtdeTratados == 0) {
	    alert('Nenhum pedido foi selecionado!!');
	    return;
	}

    //verificando se a data de recebimento está preenchida
	if (hasDtR) {
	    if (f.c_dt_recebimento.value == "") {
	        alert("A data de recebimento não está preenchida");
	        f.c_dt_recebimento.focus();
	        return;
	    }
	}

	dtR = f.c_dt_recebimento.value;

	if (dtR != "") {
	    // verifica se a data de recebimento não é uma data futura
	    if (converte_data(dtR) > converte_data(f.c_dt_hoje.value)) {
	        alert("A data de recebimento informada não pode ser uma data futura!!");
	        f.c_dt_recebimento.focus();
	        return;
	    }

	    //desprezando a pozição zero, referente ao campo hidden da página usado apenas p/ assegurar a criação de um array de campos mesmo quando houver apenas 1 linha
	    for (i = 1; i < f.ckb_recebido.length; i++) {
	        if (f.ckb_recebido[i].checked) {
	            //verificando se a data de recebimento é posterior à data da coleta
	            p = f.ckb_recebido[i].value;
	            d = f.c_dt_coleta[i].value;
	            if (converte_data(dtR) < converte_data(d)) {
	                if (f.c_dt_recebimento_pedido[i].value == "") {
	                    alert("A data de recebimento informada é anterior à data de coleta do pedido " + p);
	                    return;
	                }
	            }
	        }
	    }
	}

	window.status = "Aguarde ...";
	f.action = "RelPedidosMktplaceNaoRecebidosGravaDados.asp";
	f.submit();
}

</script>
<script type="text/javascript">
    $(document).ready(function () {
        $("#divRastreioConsultaView").hide();
        $('#divInternoRastreioConsultaView').addClass('divFixo');
        sizeDivRastreioConsultaView();
        $(document).keyup(function (e) {
            if (e.keyCode == 27) {
                fechaDivRastreioConsultaView();
            }
        });
        $("#divRastreioConsultaView").click(function () {
            fechaDivRastreioConsultaView();
        });
        $("#imgFechaDivRastreioConsultaView").click(function () {
            fechaDivRastreioConsultaView();
        });
    });
    //Every resize of window
    $(window).resize(function () {
        sizeDivRastreioConsultaView();
    });
    function fRastreioConsultaView(url) {
        sizeDivRastreioConsultaView();
        $("#divRastreioConsultaView").fadeIn();
        frame = document.getElementById("iframeRastreioConsultaView");
        frame.contentWindow.location.replace(url);
    }

	function fRastreioConsultaViaWebApiView(url) {
		executaRastreioConsultaViaWebApiView(url, urlBaseSsw, urlWebApiRastreio, "<%=usuario%>", "<%=s_sessionToken%>", "#iframeRastreioConsultaView", "#divRastreioConsultaView");
	}

	function fechaDivRastreioConsultaView() {
        $("#divRastreioConsultaView").fadeOut();
    }
    function sizeDivRastreioConsultaView() {
        var newHeight = $(document).height() + "px";
        $("#divRastreioConsultaView").css("height", newHeight);
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
.tdDataEntrega{
	width: 60px;
	}
.tdPrevEtg{
	width: 60px;
	}
.tdAtraso{
	width: 40px;
}
.tdPedido{
	width: 70px;
	}
.tdCidade{
	width: 160px;
	}
	
.tdObs2{
	width: 70px;
	}
.tdLoja{
	width: 32px;
	}
.tdCliente{
	width: 240px;
	}
.tdRecebido{
	width: 40px;
	}
.tdDtRecebido{
    width: 90px;
}
.cDtColeta
{
	background-color:transparent;
}
.cDtPrevEtg
{
	background-color:transparent;
}
.cAtraso
{
	background-color:transparent;
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
input:disabled{
    background-color: #fff;
    border: 1px solid #777;
    height: 14.2px;
}

#divRastreioConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoRastreioConsultaView
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#fff;
	opacity: 1;
}
#divInternoRastreioConsultaView.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivRastreioConsultaView
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframeRastreioConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
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
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">
<input type="hidden" name="c_dt_hoje" id="c_dt_hoje" value="<%=formata_data(Date)%>" />
<input type="hidden" name="c_usuario_sessao" id="c_usuario_sessao" value="<%=usuario%>" />

<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="ckb_recebido" id="ckb_recebido" value="">
<input type="hidden" name="c_dt_coleta" id="c_dt_coleta" value="">
<input type="hidden" name="c_dt_recebimento_pedido" id="c_dt_recebimento_pedido" value="XXXXX" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório Pedidos de Marketplace Não Recebidos Pelo Cliente</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='709' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>"

	s = ""
	s_aux = c_dt_entregue_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_entregue_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Período de Entrega:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	s = c_transportadora
	if s = "" then 
		s = "todas"
	else
		if (s_nome_transportadora <> "") And (s_nome_transportadora <> c_transportadora) then s = s & "  (" & s_nome_transportadora & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Transportadora:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = c_grupo_pedido_origem
	if s = "" then 
		s = "todos"
	else
		s = s_nome_grupo_pedido_origem
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Origem do Pedido (Grupo):&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = c_pedido_origem
	if s = "" then 
		s = "todos"
	else
		s = s_nome_pedido_origem
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Origem do Pedido:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	s = c_loja
	if s = "" then 
		s = "todas"
	else
		s = s & "  (" & s_nome_loja & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Loja:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<table cellpadding="0" cellspacing="0">
	<tr>
		<td align="right" valign="baseline">
			<span class="C">Data de recebimento</span>
		</td>
		<td align="left" valign="baseline">
			<input class="Cc" name="c_dt_recebimento" id="c_dt_recebimento" maxlength="10" style="width:90px;margin-left:2px;" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}">
		</td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<table class="notPrint" width="709" cellspacing="0" border="0">
<tr>
	<% if qtde_total_pedidos > 0 then %>
	<td align="left">
		<a name="bVOLTAR" id="bVOLTAR" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td>&nbsp;</td>
	<td align="right">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0"></a>
	</td>
	<% else %>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	<% end if %>
	</td>
</tr>
</table>

</form>

</center>

<div id="divPedidoConsulta"><center><div id="divInternoPedidoConsulta"><img id="imgFechaDivPedidoConsulta" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframePedidoConsulta"></iframe></div></center></div>
<div id="divRastreioConsultaView"><center><div id="divInternoRastreioConsultaView"><img id="imgFechaDivRastreioConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframeRastreioConsultaView"></iframe></div></center></div>

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
