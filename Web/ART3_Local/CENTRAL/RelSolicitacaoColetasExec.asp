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
'	  R E L S O L I C I T A C A O C O L E T A S E X E C . A S P
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

	Const COD_TIPO_RELATORIO_SOLICITACAO_COLETA = "SOLICITACAO_COLETA"
	Const COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO = "PRONTO_PARA_ROMANEIO"

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs,rs2, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_SOLICITACAO_COLETAS, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim url_back, strUrlBotaoVoltar
	url_back = Trim(Request("url_back"))
	if url_back <> "" then
		strUrlBotaoVoltar = "RelSolicitacaoColetasFiltro.asp?url_back=X&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	else
		strUrlBotaoVoltar = "javascript:history.back()"
		end if
	
	dim alerta
	dim i, s, s_aux, s_filtro
	dim rb_loja, c_loja, c_loja_de, c_loja_ate, c_filtro_transportadora, s_filtro_nome_transportadora, c_filtro_dt_entrega, c_nfe_emitente
	dim rb_tipo_relatorio
    dim c_fabricante_permitido, c_fabricante_proibido, c_zona_permitida, c_zona_proibida
    dim v_fabricante_permitido, v_fabricante_proibido, v_zona_permitida, v_zona_proibida

	alerta = ""

	rb_loja = Ucase(Trim(Request.Form("rb_loja")))
	c_loja = Trim(Request.Form("c_loja"))
	c_loja_de = Trim(Request.Form("c_loja_de"))
	c_loja_ate = Trim(Request.Form("c_loja_ate"))
	rb_tipo_relatorio = Trim(Request.Form("rb_tipo_relatorio"))
	c_filtro_transportadora = Trim(Request.Form("c_filtro_transportadora"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
    c_fabricante_permitido = Trim(Request.Form("c_fabricante_permitido"))
    c_fabricante_proibido = Trim(Request.Form("c_fabricante_proibido"))
    c_zona_permitida = Trim(Request.Form("c_zona_permitida"))
    c_zona_proibida = Trim(Request.Form("c_zona_proibida"))
	
	if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
		c_filtro_dt_entrega = Trim(Request.Form("c_filtro_dt_entrega"))
	else
		c_filtro_dt_entrega = ""
		end if
	
	s_filtro_nome_transportadora = ""
	if c_filtro_transportadora <> "" then s_filtro_nome_transportadora = x_transportadora(c_filtro_transportadora)
	
	if alerta = "" then
		if (rb_tipo_relatorio <> COD_TIPO_RELATORIO_SOLICITACAO_COLETA) And _
		   (rb_tipo_relatorio <> COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO) then
			alerta = "É necessário informar o tipo de saída do relatório."
			end if
		end if
		
	if alerta = "" then
		if rb_loja = "UMA" then
			if c_loja = "" then
				alerta = "Especifique o número da loja."
			else
				s = "SELECT loja FROM t_LOJA WHERE (loja='" & c_loja & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta = "Loja " & c_loja & " não está cadastrada."
					end if
				end if
		elseif rb_loja = "FAIXA" then
			if (c_loja_de="") And (c_loja_ate="") then
				alerta = "Especifique o intervalo de lojas para consulta."
			else
				if c_loja_de <> "" then
					s = "SELECT loja FROM t_LOJA WHERE (loja='" & c_loja_de & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta = alerta & "Loja " & c_loja_de & " não está cadastrada."
						end if
					end if
				
				if c_loja_ate <> "" then
					s = "SELECT loja FROM t_LOJA WHERE (loja='" & c_loja_ate & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta = alerta & "Loja " & c_loja_ate & " não está cadastrada."
						end if
					end if
				end if
			end if
		end if

	dim qtde_pedidos
	qtde_pedidos = 0

	if alerta = "" then
		if c_nfe_emitente = "" then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Não foi informado o CD"
		elseif converte_numero(c_nfe_emitente) = 0 then
			alerta=texto_add_br(alerta)
			alerta = alerta & "É necessário definir um CD válido"
			end if
		end if

    if alerta = "" then
		call set_default_valor_texto_bd(usuario, "RelSolicitacaoColetasFiltro|c_fabricante_permitido", c_fabricante_permitido)
		call set_default_valor_texto_bd(usuario, "RelSolicitacaoColetasFiltro|c_fabricante_proibido", c_fabricante_proibido)
		call set_default_valor_texto_bd(usuario, "RelSolicitacaoColetasFiltro|c_zona_permitida", c_zona_permitida)
		call set_default_valor_texto_bd(usuario, "RelSolicitacaoColetasFiltro|c_zona_proibida", c_zona_proibida)
        end if






' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function decodifica_status_contribuinte_icms(byval tipo, byval st_pj, byval st_pf)
dim strResp

	tipo = Trim(tipo)
	st_pj = Trim(st_pj)
	st_pf = Trim(st_pf)

	if tipo = ID_PJ then
		select case st_pj
			case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO
				strResp = "Não"
			case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM
				strResp = "Sim"
			case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO
				strResp = "Isento"
			case else
				strResp = ""
		end select
	elseif tipo = ID_PF then
		select case st_pf
			case COD_ST_CLIENTE_PRODUTOR_RURAL_NAO
				strResp = "Não"
			case COD_ST_CLIENTE_PRODUTOR_RURAL_SIM
				strResp = "Sim"
			case else
				strResp = ""
		end select
	end if

	decodifica_status_contribuinte_icms = strResp
end function


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
dim s, s_sql, s_lista, cab_table, cab, n_reg, s_num_nfe, s_link_nfe, s_row, s_html_color, s_link_habilita_print, s_link_indicador
dim s_cidade, s_uf, s_transportadora, s_data_entrega_yyyymmdd, s_data_credito_ok
dim s_where, s_where_aux, s_from
dim i, intCodCor, intOrdenacaoCor
dim blnRegistroOk
dim vRel()
dim rNfeEmitente
dim x
dim qtde_produto,valor,total_qtde,total_valor

total_qtde = 0
total_valor = 0

'	AS ROTINAS DE ORDENAÇÃO USAM VETORES QUE SE INICIAM NA POSIÇÃO 1
	redim vRel(1)
	for i = Lbound(vRel) to Ubound(vRel)
		set vRel(i) = New cl_DUAS_COLUNAS
		with vRel(i)
			.c1 = ""
			.c2 = 0
			end with
		next

'	MONTA CLÁUSULA WHERE
	s_where = " (" & _
					"(t_PEDIDO.st_entrega = '" & ST_ENTREGA_SEPARAR & "')" & _
					" OR " & _
					"(t_PEDIDO.st_entrega = '" & ST_ENTREGA_A_ENTREGAR & "')" & _
				")" & _
			  " AND (" & _
					"(t_PEDIDO.danfe_impressa_status=" & COD_DANFE_IMPRESSA_STATUS__INICIAL & ")" & _
					" OR " & _
					"(t_PEDIDO.danfe_impressa_status=" & COD_DANFE_IMPRESSA_STATUS__NAO_DEFINIDO & ")" & _
				")" & _
			  " AND (t_PEDIDO.danfe_a_imprimir_status<>" & COD_DANFE_A_IMPRIMIR_STATUS__IMPRESSA & ")" & _
			  " AND (t_PEDIDO__BASE.analise_credito = " & COD_AN_CREDITO_OK & ")" & _
			  " AND (t_PEDIDO.st_etg_imediata = " & COD_ETG_IMEDIATA_SIM & ")"

'	IGNORA PEDIDOS DA LOJA OLD03
	s = " t_PEDIDO.numero_loja NOT IN (" & NUMERO_LOJA_OLD03 & "," & NUMERO_LOJA_OLD03_BONIFICACAO & "," & NUMERO_LOJA_OLD03_ASSISTENCIA & ")"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"
	
'	CRITÉRIO: LOJA
	if rb_loja = "UMA" then
		if c_loja <> "" then
			s = " (t_PEDIDO.numero_loja = " & c_loja & ")"
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
	elseif rb_loja = "FAIXA" then
		s = ""
		if c_loja_de <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.numero_loja >= " & c_loja_de & ")"
			end if

		if c_loja_ate <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO.numero_loja <= " & c_loja_ate & ")"
			end if
		
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CRITÉRIO: FABRICANTE PERMITIDO
    if c_fabricante_permitido <> "" then
        v_fabricante_permitido = Split(c_fabricante_permitido, ", ")
        s_lista = ""
        for i=LBound(v_fabricante_permitido) to UBound(v_fabricante_permitido)
            if Trim("" & v_fabricante_permitido(i)) <> "" then
                if s_lista <> "" then s_lista = s_lista & ", "
                s_lista = s_lista & "'" & Trim("" & v_fabricante_permitido(i)) & "'"
            end if
        next
        
        if s_lista <> "" then
            'DEVE-SE SELECIONAR SOMENTE OS PEDIDOS QUE CONTENHAM EXCLUSIVAMENTE PRODUTOS DOS FABRICANTES SELECIONADOS, OU SEJA, OS PEDIDOS NÃO PODEM CONTER PRODUTOS DE OUTROS FABRICANTES
            s_where_aux = "COALESCE((SELECT Count(*) FROM t_PEDIDO_ITEM tPI WHERE (tPI.pedido=t_PEDIDO.pedido) AND (tPI.fabricante NOT IN (" & s_lista & "))), 0) = 0"
            if s_where <> "" then s_where = s_where & " AND"
            s_where = s_where & " (" & s_where_aux & ")"
        end if
    end if

'	CRITÉRIO: FABRICANTE PROIBIDO
    if c_fabricante_proibido <> "" then
        v_fabricante_proibido = Split(c_fabricante_proibido, ", ")
        s_lista = ""
        for i=LBound(v_fabricante_proibido) to UBound(v_fabricante_proibido)
            if Trim("" & v_fabricante_proibido(i)) <> "" then
                if s_lista <> "" then s_lista = s_lista & ", "
                s_lista = s_lista & "'" & Trim("" & v_fabricante_proibido(i)) & "'"
            end if
        next
        
        if s_lista <> "" then
            'DEVE-SE SELECIONAR SOMENTE OS PEDIDOS QUE NÃO CONTENHAM OS PRODUTOS DOS FABRICANTES SELECIONADOS
            s_where_aux = "COALESCE((SELECT Count(*) FROM t_PEDIDO_ITEM tPI WHERE (tPI.pedido=t_PEDIDO.pedido) AND (tPI.fabricante IN (" & s_lista & "))), 0) = 0"
            if s_where <> "" then s_where = s_where & " AND"
            s_where = s_where & " (" & s_where_aux & ")"
        end if
    end if

'	CRITÉRIO: ZONA PERMITIDA
    if c_zona_permitida <> "" then
        v_zona_permitida = Split(c_zona_permitida, ", ")
        s_lista = ""
        for i=LBound(v_zona_permitida) to UBound(v_zona_permitida)
            if Trim("" & v_zona_permitida(i)) <> "" then
                if s_lista <> "" then s_lista = s_lista & ", "
                s_lista = s_lista & "'" & Trim("" & v_zona_permitida(i)) & "'"
            end if
        next
        
        if s_lista <> "" then
            'DEVE-SE SELECIONAR SOMENTE OS PEDIDOS QUE CONTENHAM EXCLUSIVAMENTE PRODUTOS ARMAZENADOS NAS ZONAS DO DEPÓSITO SELECIONADAS, OU SEJA, OS PEDIDOS NÃO PODEM CONTER PRODUTOS DE OUTRAS ZONAS
            s_where_aux = "COALESCE((SELECT Count(*) FROM t_PEDIDO_ITEM tPI INNER JOIN t_PRODUTO tPROD ON (tPI.fabricante = tPROD.fabricante) AND (tPI.produto = tPROD.produto) WHERE (tPI.pedido=t_PEDIDO.pedido) AND (tPROD.deposito_zona_id NOT IN (" & s_lista & "))), 0) = 0"
            if s_where <> "" then s_where = s_where & " AND"
            s_where = s_where & " (" & s_where_aux & ")"
        end if
    end if

'	CRITÉRIO: ZONA PROIBIDA
    if c_zona_proibida <> "" then
        v_zona_proibida = Split(c_zona_proibida, ", ")
        s_lista = ""
        for i=LBound(v_zona_proibida) to UBound(v_zona_proibida)
            if Trim("" & v_zona_proibida(i)) <> "" then
                if s_lista <> "" then s_lista = s_lista & ", "
                s_lista = s_lista & "'" & Trim("" & v_zona_proibida(i)) & "'"
            end if
        next
        
        if s_lista <> "" then
            'DEVE-SE SELECIONAR SOMENTE OS PEDIDOS QUE NÃO CONTENHAM OS PRODUTOS DAS ZONAS DO DEPÓSITO SELECIONADAS
            s_where_aux = "COALESCE((SELECT Count(*) FROM t_PEDIDO_ITEM tPI INNER JOIN t_PRODUTO tPROD ON (tPI.fabricante = tPROD.fabricante) AND (tPI.produto = tPROD.produto) WHERE (tPI.pedido=t_PEDIDO.pedido) AND (tPROD.deposito_zona_id IN (" & s_lista & "))), 0) = 0"
            if s_where <> "" then s_where = s_where & " AND"
            s_where = s_where & " (" & s_where_aux & ")"
        end if
    end if

'	CRITÉRIO: TRANSPORTADORA
	if c_filtro_transportadora <> "" then
		s = " (t_PEDIDO.transportadora_id = '" & c_filtro_transportadora & "')"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & s
		end if
	
'	CRITÉRIO: DATA DE COLETA (RÓTULO ANTIGO: DATA ENTREGA)
	if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
		if c_filtro_dt_entrega <> "" then
			s = " (t_PEDIDO.a_entregar_data_marcada = " & bd_formata_data(StrToDate(c_filtro_dt_entrega)) & ")"
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & s
			end if
		end if
	
'	OWNER DO PEDIDO
	set rNfeEmitente = le_nfe_emitente(c_nfe_emitente)
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (t_PEDIDO.id_nfe_emitente = " & rNfeEmitente.id & ")"

	if s_where <> "" then s_where = " WHERE" & s_where
	
	
'	MONTA CLÁUSULA FROM
	s_from = " FROM t_PEDIDO" & _
			 " INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			 " INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)"

'	Tipo de NFe: 0-Entrada  1-Saída
	s_sql = "SELECT" & _
				" t_PEDIDO.pedido," & _
				" t_PEDIDO.data," & _
				" t_PEDIDO.loja," & _
				" t_PEDIDO.a_entregar_data_marcada," & _
				" t_PEDIDO.transportadora_id," & _
				" t_PEDIDO.st_end_entrega," & _
				" t_PEDIDO.EndEtg_endereco," & _
				" t_PEDIDO.EndEtg_endereco_numero," & _
				" t_PEDIDO.EndEtg_endereco_complemento," & _
				" t_PEDIDO.EndEtg_bairro," & _
				" t_PEDIDO.EndEtg_cidade," & _
				" t_PEDIDO.EndEtg_uf," & _
				" t_PEDIDO.EndEtg_cep," & _
				" t_PEDIDO.danfe_a_imprimir_status," & _
                " t_PEDIDO.indicador," & _
				" t_PEDIDO__BASE.analise_credito," & _
				" t_PEDIDO__BASE.analise_credito_data," & _
				" t_CLIENTE.endereco," & _
				" t_CLIENTE.endereco_numero," & _
				" t_CLIENTE.endereco_complemento," & _
				" t_CLIENTE.bairro," & _
				" t_CLIENTE.cidade," & _
				" t_CLIENTE.uf," & _
				" t_CLIENTE.cep," & _
                " t_CLIENTE.tipo," & _
	            " t_CLIENTE.ie," & _
	            " t_CLIENTE.contribuinte_icms_status," & _
	            " t_CLIENTE.produtor_rural_status," & _
				" (" & _
					"SELECT" & _
						" Count(*)" & _
					" FROM t_PEDIDO_NFe_EMISSAO_SOLICITADA tPNES" & _
					" WHERE" & _
						" (tPNES.pedido=t_PEDIDO.pedido)" & _
						" AND (" & _
							"(nfe_emitida_status=" & COD_NFE_EMISSAO_SOLICITADA__PENDENTE & ")" & _
							" OR " & _
							"(nfe_emitida_status=" & COD_NFE_EMISSAO_SOLICITADA__ATENDIDA & ")" & _
							")" & _
				") AS qtde_solicitacao_emissao_nfe," & _
				" (" & _
					"SELECT" & _
						" TOP 1 NFe_numero_NF" & _
					" FROM t_NFe_EMISSAO tNE" & _
					" WHERE" & _
						" (tNE.pedido=t_PEDIDO.pedido)" & _
						" AND (tipo_NF = '1')" & _
						" AND (st_anulado = 0)" & _
						" AND (codigo_retorno_NFe_T1 = 1)" & _
					" ORDER BY" & _
						" id DESC" & _
				") AS numeroNFe" & _
			s_from & _
			s_where
	
	s_sql = s_sql & " ORDER BY"
	
	if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
		s_sql = s_sql & " t_PEDIDO.transportadora_id,"
		end if
		
	s_sql = s_sql & " t_PEDIDO.data, t_PEDIDO.pedido"

  ' CABEÇALHO
	cab_table = "<table cellspacing=0 id='tabelaRelatorio'>" & chr(13)
	cab = "	<tr style='background:azure'>" & chr(13) & _
		  "		<td class='ME MC MD' style='width:70px' align='left' valign='bottom' nowrap><span class='R'>Nº Pedido</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:35px' align='center' valign='bottom' nowrap><span class='R'>Loja</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:180px' align='left' valign='bottom' nowrap><span class='R'>Cidade</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:30px' align='center' valign='bottom' nowrap><span class='R'>UF</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:80px' align='center' valign='bottom' nowrap><span class='R'>Insc Estadual</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:50px' align='center' valign='bottom' nowrap><span class='R'>Contribuinte</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:70px' align='center' valign='bottom' nowrap><span class='R'>Indicador</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:60px' align='center' valign='bottom' nowrap><span class='R'>Crédito</br>OK</span></td>" & chr(13) & _
          "		<td class='MC MD' style='width:30px' align='right' valign='bottom' nowrap><span class='R'>Qtde </br> Vol</span></td>" & chr(13) & _
          "		<td class='MC MD' style='width:70px' align='right' valign='bottom' nowrap><span class='R'>Valor</span></td>" & chr(13) 
		
	if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then
		cab = cab & _
			"		<td class='MC MD' style='width:70px' align='center' valign='bottom' nowrap><span class='R'>Gravar </br> Transp+Data</span></td>" & chr(13)
		end if
	
	cab = cab & _
			"		<td class='MC MD' style='width:70px' align='center' valign='bottom' nowrap><span class='R'>Data Coleta</span></td>" & chr(13) & _
			"		<td class='MC MD' style='width:70px' align='center' valign='bottom' nowrap><span class='R'>Transp</span></td>" & chr(13)
		
	if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then
		cab = cab & _
			"		<td class='MC MD' style='width:30px' align='center' valign='bottom' nowrap><span class='R'>Emitir </br> NFe</span></td>" & chr(13)
		end if
	
	cab = cab & _
			"		<td class='MC MD' style='width:70px' align='center' valign='bottom' nowrap><span class='R'>Nº NFe</span></td>" & chr(13)
	
	if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
		cab = cab & _
			"		<td class='MC MD' style='width:50px' align='center' valign='bottom' nowrap><span class='R'>DANFE </br> Impressa</span></td>" & chr(13)
		end if
	
	if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
		cab = cab & _
			"		<td class='MC MD' style='width:50px' align='center' valign='bottom' nowrap><span class='R'>Imprimir </br> PDF</span></td>" & chr(13)
		end if
	
	cab = cab & _
			"	</tr>" & chr(13)
	
	n_reg = 0
	x = cab_table & cab

	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	ANALISA A SITUAÇÃO DO PEDIDO E VERIFICA SE ELE PERTENCE OU NÃO AO TIPO DE RELATÓRIO SELECIONADO
		blnRegistroOk = False
		if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
			if (Trim("" & r("a_entregar_data_marcada")) <> "") And _
			   (Trim("" & r("transportadora_id")) <> "") And _
			   (Trim("" & r("numeroNFe")) <> "") then
				blnRegistroOk = True
				end if
		elseif rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then
			if (Trim("" & r("a_entregar_data_marcada")) = "") Or _
			   (Trim("" & r("transportadora_id")) = "") Or _
			   (Trim("" & r("numeroNFe")) = "") then
				blnRegistroOk = True
				end if
			end if
		
		if blnRegistroOk then

		 '	CONTAGEM
			n_reg = n_reg + 1

		'	ANALISA A SITUAÇÃO DO PEDIDO E DEFINE A COR DE EXIBIÇÃO
		'	=======================================================
			if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
			   intCodCor = COD_COR_PRETO
			   intOrdenacaoCor = 0
			   s_html_color = "black"
			else
			'	VERMELHO: SÓ FALTA GERAR A NFe (A SOLICITAÇÃO DE EMISSÃO JÁ FOI FEITA)
				if (Trim("" & r("a_entregar_data_marcada")) <> "") And _
				   (Trim("" & r("transportadora_id")) <> "") And _
				   (CLng(r("qtde_solicitacao_emissao_nfe")) > 0) And _
				   (Trim("" & r("numeroNFe")) = "") then
					intCodCor = COD_COR_VERMELHO
					intOrdenacaoCor = 30
					s_html_color = "red"
			'	AZUL: JÁ POSSUI TRANSPORTADORA E DATA DE COLETA PREENCHIDOS, FALTA FAZER A SOLICITAÇÃO DE EMISSÃO DA NFe
				elseif (Trim("" & r("a_entregar_data_marcada")) <> "") And _
				   (Trim("" & r("transportadora_id")) <> "") And _
				   (CLng(r("qtde_solicitacao_emissao_nfe")) = 0) And _
				   (Trim("" & r("numeroNFe")) = "") then
					intCodCor = COD_COR_AZUL
					intOrdenacaoCor = 20
					s_html_color = "blue"
			'	PRETO: NÃO TEM NADA PREENCHIDO
			'	LEMBRANDO QUE A TRANSPORTADORA E A DATA DE COLETA PODEM SER PREENCHIDOS INDIVIDUALMENTE ATRAVÉS DE OUTRAS OPERAÇÕES ESPECÍFICAS P/ ISSO (JÁ EXISTENTES ANTERIORMENTE)
				elseif ( (Trim("" & r("a_entregar_data_marcada")) = "") Or (Trim("" & r("transportadora_id")) = "") ) And _
				   (CLng(r("qtde_solicitacao_emissao_nfe")) = 0) And _
				   (Trim("" & r("numeroNFe")) = "") then
					intCodCor = COD_COR_PRETO
					intOrdenacaoCor = 10
					s_html_color = "black"
				else
					intCodCor = COD_COR_NAO_DEFINIDO
					intOrdenacaoCor = 0
					s_html_color = "darkorange"
					end if
				end if
			
			s_html_color = " style='color:" & s_html_color & ";'"
			
		'	MONTA O HTML DA LINHA DA TABELA
		'	===============================
			s_row = "	<tr onmouseover='realca_cor_mouse_over(this);' onmouseout='realca_cor_mouse_out(this);'>" & chr(13)

		'> Nº PEDIDO
			s_row = s_row & _
					"		<td align='left' valign='top' class='ME MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">&nbsp;<a href='javascript:fRELConcluir(" & chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'" & s_html_color & ">" & Trim("" & r("pedido")) & "</a></span>" & chr(13) & _
					"			<input type='hidden' name='c_numero_pedido' value='" & Trim("" & r("pedido")) & "' />" & chr(13) & _
					"		</td>" & chr(13)
		'	LOJA
			s_row = s_row & _
					"		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & Trim("" & r("loja")) & "</span>" & chr(13) & _
					"		</td>" & chr(13)
			
		'	CIDADE E UF
			if Trim("" & r("st_end_entrega")) <> "0" then
				s_cidade = Ucase(Trim("" & r("EndEtg_cidade")))
				s_uf = Ucase(Trim("" & r("EndEtg_uf")))
			else
				s_cidade = Ucase(Trim("" & r("cidade")))
				s_uf = Ucase(Trim("" & r("uf")))
				end if

			s_row = s_row & _
					"		<td align='left' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s_cidade & "</span>" & chr(13) & _
					"		</td>" & chr(13) & _
					"		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s_uf & "</span>" & chr(13) & _
					"		</td>" & chr(13)

        '	INSC ESTADUAL
            s = retorna_so_digitos(Trim("" & r("ie")))
			s_row = s_row & _
					"		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					"		</td>" & chr(13)

        '	CONTRIBUINTE DE ICMS?
            s = decodifica_status_contribuinte_icms(r("tipo"), _
													Trim("" & r("contribuinte_icms_status")), _
													Trim("" & r("produtor_rural_status")))
			s_row = s_row & _
					"		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					"		</td>" & chr(13)

        '	INDICADOR
            s_link_indicador = ""
			s = Trim("" & r("indicador"))
            if (Not operacao_permitida(OP_CEN_CADASTRO_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) And _
	            (Not operacao_permitida(OP_CEN_PESQUISA_INDICADORES, s_lista_operacoes_permitidas)) then 
                s_link_indicador = "<span class='C'" & s_html_color & ">" & s & "</span>"
            else
                s_link_indicador = "<a href='javascript:fOrcamentistaEIndicadorConsultaView(" & chr(34) & s & chr(34) & ")' title='clique para consultar o cadastro do indicador'><span class='C'" & s_html_color & ">" & s & "</span></a>"
            end if
			s_row = s_row & _
					"		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					"			" & s_link_indicador & chr(13) & _
					"		</td>" & chr(13)
			
		'	DATA DA ANÁLISE DE CRÉDITO OK
			s_data_credito_ok = "&nbsp;"
			if Trim("" & r("analise_credito")) = Trim("" & COD_AN_CREDITO_OK) then s_data_credito_ok = formata_data(r("analise_credito_data"))
			s_row = s_row & _
					"		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s_data_credito_ok & "</span>" & chr(13) & _
					"		</td>" & chr(13)
			
        '	QTDE E VALOR
            qtde_produto = 0
            valor = 0
            rs2.Open "SELECT qtde,qtde_volumes,preco_nf from t_PEDIDO_ITEM WHERE pedido = '" & Trim("" & r("pedido")) & "'",cn
            do while not rs2.Eof
                qtde_produto = qtde_produto + (Trim("" & rs2("qtde")) * Trim("" & rs2("qtde_volumes")))
                valor = valor + (Trim("" & rs2("qtde")) * Trim("" & rs2("preco_nf")))
                rs2.MoveNext
            loop
            total_qtde = total_qtde + qtde_produto
            total_valor = total_valor + valor
            s_row = s_row & _
					"		<td align='right' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & qtde_produto & "</span>" & chr(13) & _
					"		</td>" & chr(13) & _
					"		<td align='right' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & formata_moeda(valor) & "</span>" & chr(13) & _
					"		</td>" & chr(13)
            
            
			if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then
			'	CHECK BOX: GRAVAR TRANSP+DATA
				s_row = s_row & _
						"		<td align='center' valign='top' class='MC MD' style='padding:0px;'>" & chr(13) & _
						"			<input type='checkbox' name='ckb_gravar_transp_e_dt_entrega' id='ckb_gravar_transp_e_dt_entrega' value='" & Trim("" & r("pedido")) & "'"
						
				if CLng(r("qtde_solicitacao_emissao_nfe")) > 0 then s_row = s_row & " disabled"

				s_row = s_row & ">" & chr(13) & _
						"		</td>" & chr(13)
				end if

		'	DATA DE COLETA (RÓTULO ANTIGO: DATA DA ENTREGA)
			s = formata_data(r("a_entregar_data_marcada"))
			if s = "" then s = "&nbsp;"
			s_row = s_row & _
					"		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					"		</td>" & chr(13)

		'	TRANSPORTADORA
			s = Trim("" & r("transportadora_id"))
			if s = "" then s = "&nbsp;"
			s_row = s_row & _
					"		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					"			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					"			<input type='hidden' name='c_pedido_transportadora' value='" & Trim("" & r("transportadora_id")) & "' />" & chr(13) & _
					"		</td>" & chr(13)

			if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then
			'	CHECK BOX: EMITIR NFe
				if CLng(r("qtde_solicitacao_emissao_nfe")) > 0 then s="S" else s="N"
				s_row = s_row & _
						"		<td align='center' valign='top' class='MC MD' style='padding:0px;'>" & chr(13) & _
						"			<input type='checkbox' name='ckb_emitir_nfe' id='ckb_emitir_nfe' value='" & Trim("" & r("pedido")) & "|" & s & "'"
				
				blnDisabled = False
				if CLng(r("qtde_solicitacao_emissao_nfe")) > 0 then 
					s_row = s_row & " checked disabled"
					blnDisabled = True
					end if
				
			'	PARA FORÇAR O PREENCHIMENTO DOS CAMPOS NA SEQUENCIA CORRETA
				if Not blnDisabled then
					if (Trim("" & r("transportadora_id")) = "") Or (Trim("" & r("a_entregar_data_marcada")) = "") then
						s_row = s_row & " disabled"
						end if
					end if
					
				s_row = s_row & ">" & chr(13) & _
						"		</td>" & chr(13)
				end if

		'	Nº NFe
			s_num_nfe = Trim("" & r("numeroNFe"))
			if s_num_nfe <> "" then
				s = "<span class='C'" & s_html_color & ">" & NFeFormataNumeroNF(s_num_nfe) & "</span>"
				s_link_nfe = monta_link_para_ultima_DANFE(Trim("" & r("pedido")), MAX_PERIODO_LINK_DANFE_DISPONIVEL_NO_PEDIDO_EM_DIAS, s)
				'aproveitar o link, se existir, para ver se habilita o checkbox Imprimir PDF
				s_link_habilita_print = s_link_nfe
				if s_link_nfe = "" then s_link_nfe = "<span class='C' style='color:gray;'>" & NFeFormataNumeroNF(s_num_nfe) & "</span>"
			else
				s_link_nfe = "&nbsp;"
				end if
				
			s_row = s_row & _
					"		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					"			" & s_link_nfe & chr(13) & _
					"		</td>" & chr(13)

			if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
			'	CHECK BOX: DANFE IMPRESSA?
				s_row = s_row & _
						"		<td align='center' valign='top' class='MC MD' style='padding:0px;'>" & chr(13) & _
						"			<input type='checkbox' name='ckb_danfe_impressa' id='ckb_danfe_impressa' value='" & Trim("" & r("pedido")) & "'>" & chr(13) & _
						"		</td>" & chr(13)
				end if

			if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
			'	Imprimir PDF
				s_row = s_row & _
					"		<td align='center' valign='top' class='MC MD' style='padding:0px;'>" & chr(13)
				s_row = s_row & _
					"			<input type='checkbox' name='ckb_danfe_a_imprimir' id='ckb_danfe_a_imprimir' value='" & Trim("" & r("pedido")) & "'"

				if r("danfe_a_imprimir_status") = CInt(COD_DANFE_A_IMPRIMIR_STATUS__MARCADA) then s_row = s_row & " checked disabled"

				if s_link_habilita_print = "" then s_row = s_row & " disabled"

				s_row = s_row & chr(13)

				s_row = s_row & _
					"		</td>" & chr(13)
				end if

			s_row = s_row & "	</tr>" & chr(13)
			
			if Trim(vRel(Ubound(vRel)).c1) <> "" then
				redim preserve vRel(Ubound(vRel)+1)
				set vRel(Ubound(vRel)) = New cl_DUAS_COLUNAS
				end if

			s_data_entrega_yyyymmdd = formata_data_yyyymmdd(r("a_entregar_data_marcada"))
			if s_data_entrega_yyyymmdd = "" then s_data_entrega_yyyymmdd = String(8,"0")
			s_transportadora = Trim("" & r("transportadora_id"))
			s_transportadora = PadRight(s_transportadora, 10)
			
			with vRel(Ubound(vRel))
				if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then
					.c1 = normaliza_codigo(intOrdenacaoCor,2) & "|" & _
						  PadRight(s_uf,2) & "|" & _
						  retira_acentuacao(PadRight(s_cidade,60)) & "|" & _
						  s_data_entrega_yyyymmdd & "|" & _
						  s_transportadora & "|" & _
						  normaliza_codigo(n_reg,6)
				else
					.c1 = normaliza_codigo(intOrdenacaoCor,2) & "|" & _
						  s_data_entrega_yyyymmdd & "|" & _
						  s_transportadora & "|" & _
						  normaliza_codigo(n_reg,6)
					end if
				.c2 = s_row
				end with

			end if	'if blnRegistroOk then
        
        if rs2.State <> 0 then rs2.Close
		r.MoveNext
		loop
		
  ' MOSTRA TOTAL DE PEDIDOS
	if n_reg <> 0 then 
		ordena_cl_duas_colunas vRel, 1, Ubound(vRel)
		
		for i = Lbound(vRel) to Ubound(vRel)
			with vRel(i)
				if Trim("" & .c1) <> "" then
					x = x & .c2
					if (i mod 100) = 0 then
						Response.Write x
						x = ""
						end if
					end if
				end with
			next

		x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
                "       <td class='MTBE' nowrap colspan=8 align='right'><span class='C'>" & _
				"TOTAL:   </span></td>" & chr(13) & _
                "       <td class='MTB' nowrap colspan=1 align='right'><span class='C'>"& formata_inteiro(total_qtde) & "</span></td>" & chr(13) & _
				"       <td class='MTB' nowrap colspan=1 align='right'><span class='C'>"& formata_moeda(total_valor) & "</span></td>" & chr(13) & _          
				"		<td class='MTBD' nowrap colspan=6 align='right'><span class='C'>" & _
				"TOTAL DE REGISTRO(S):  &nbsp; " & formata_inteiro(n_reg) & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab
		x = x & "	<tr>" & chr(13) & _
				"		<td class='MT ALERTA' colspan='15' align='center'><span class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</table>" & chr(13)
	
	Response.write x

	qtde_pedidos = n_reg
	
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
var COD_TIPO_RELATORIO_SOLICITACAO_COLETA = "SOLICITACAO_COLETA";
var COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO = "PRONTO_PARA_ROMANEIO";

$(function () {
    $("#divOrcamentistaEIndicadorConsultaView").hide();
    $('#divInternoOrcamentistaEIndicadorConsultaView').addClass('divFixo');
    sizeDivOrcamentistaEIndicadorConsultaView();

    $("#divOrcamentistaEIndicadorConsultaView").click(function () {
        fechaDivOrcamentistaEIndicadorConsultaView();
    });

    $("#imgFechaDivOrcamentistaEIndicadorConsultaView").click(function () {
        fechaDivOrcamentistaEIndicadorConsultaView();
    });

    $("#tBodyCodCores").hide();
    $("#tBodyCodCores").addClass("TBODYCODCORES_HIDDEN");
});

$(window).resize(function () {
    sizeDivOrcamentistaEIndicadorConsultaView();
});

function TBodyCodCoresToggle() {
	if ($("#tBodyCodCores").hasClass("TBODYCODCORES_HIDDEN")) {
		$("#tBodyCodCores").show();
		$("#tBodyCodCores").removeClass("TBODYCODCORES_HIDDEN");
		document.getElementById("imgToggleCodCores").src = document.getElementById("imgToggleCodCores").src.replace("double-down-20.png", "double-up-20.png");
	}
	else {
		$("#tBodyCodCores").hide();
		$("#tBodyCodCores").addClass("TBODYCODCORES_HIDDEN");
		document.getElementById("imgToggleCodCores").src = document.getElementById("imgToggleCodCores").src.replace("double-up-20.png", "double-down-20.png");
	}
}

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

function fRELGravaDados(f) {
	var i, intQtdeEmissao, intQtdeTransp, intQtdeDanfeImpressa, intQtdeDanfeAImprimir, dtEntrega, dtHoje;
	var s, s_pedido_sem_transp, s_pedido_com_transp;

	if (f.rb_tipo_relatorio.value == COD_TIPO_RELATORIO_SOLICITACAO_COLETA) {
		intQtdeEmissao = 0;
		for (i = 0; i < f.ckb_emitir_nfe.length; i++) {
			if (f.ckb_emitir_nfe[i].checked && !f.ckb_emitir_nfe[i].disabled) intQtdeEmissao++;
		}

		intQtdeTransp = 0;
		for (i = 0; i < f.ckb_gravar_transp_e_dt_entrega.length; i++) {
			if (f.ckb_gravar_transp_e_dt_entrega[i].checked) intQtdeTransp++;
		}

		if ((intQtdeEmissao == 0) && (intQtdeTransp == 0)) {
			alert('Nenhum pedido foi selecionado!!');
			return;
		}

		if (intQtdeTransp > 0) {
			if (trim(f.c_dt_entrega.value) == "") {
				alert('Informe a data!!');
				f.c_dt_entrega.focus();
				return;
			}
			if (!isDate(f.c_dt_entrega)) {
				alert('Data inválida!!');
				f.c_dt_entrega.focus();
				return;
			}
			dtHoje = new Date();
			dtHoje = new Date(dtHoje.getFullYear(), dtHoje.getMonth(), dtHoje.getDate());
			dtEntrega = converte_data(f.c_dt_entrega.value);
			if (dtEntrega < dtHoje) {
				alert('Não é permitido informar uma data passada!!');
				f.c_dt_entrega.focus();
				return;
			}
			// Se os pedidos já possuem transportadora anotada (podem ser transportadoras variadas), permite gravar apenas a data de coleta.
			// Mas se houver pedido sem transportadora, obriga a selecionar uma transportadora da lista.
			if (trim(f.c_transportadora.value) == "") {
				s_pedido_sem_transp = "";
				for (i = 0; i < f.ckb_gravar_transp_e_dt_entrega.length; i++) {
					if (f.ckb_gravar_transp_e_dt_entrega[i].checked) {
						if (trim(f.c_pedido_transportadora[i].value) == "") {
							if (s_pedido_sem_transp.length > 0) s_pedido_sem_transp += "\n";
							s_pedido_sem_transp += f.c_numero_pedido[i].value;
						}
					}
				}
				if (s_pedido_sem_transp.length > 0) {
					s = "É necessário selecionar uma transportadora, pois os seguintes pedidos estão sem nenhuma transportadora:" + "\n" + s_pedido_sem_transp;
					alert(s);
					f.c_transportadora.focus();
					return;
				}
			}
			// Se os pedidos já possuem transportadora anotada (podem ser transportadoras variadas), permite gravar apenas a data de coleta.
			// Se uma transportadora for selecionada, verifica se algum pedido já possui uma transportadora diferente anotada p/ exibir uma confirmação.
			if (trim(f.c_transportadora.value) != "") {
				s_pedido_com_transp = "";
				for (i = 0; i < f.ckb_gravar_transp_e_dt_entrega.length; i++) {
					if (f.ckb_gravar_transp_e_dt_entrega[i].checked) {
						if (trim(f.c_pedido_transportadora[i].value) != "") {
							if (trim(f.c_pedido_transportadora[i].value.toUpperCase()) != trim(f.c_transportadora.value.toUpperCase())) {
								if (s_pedido_com_transp.length > 0) s_pedido_com_transp += "\n";
								s_pedido_com_transp += f.c_numero_pedido[i].value + " (" + f.c_pedido_transportadora[i].value + ")";
							}
						}
					}
				}
				if (s_pedido_com_transp.length > 0) {
					s = "Os seguintes pedidos já possuem uma transportadora diferente da transportadora selecionada:" + "\n" + s_pedido_com_transp + "\n\n" + "Continua e grava a nova transportadora?";
					if (!confirm(s)) return;
				}
			}
		}
	}

	if (f.rb_tipo_relatorio.value == COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO) {
		intQtdeDanfeImpressa = 0;
		intQtdeDanfeAImprimir = 0;
		for (i = 0; i < f.ckb_danfe_impressa.length; i++) {
			if (f.ckb_danfe_impressa[i].checked) intQtdeDanfeImpressa++;
			if (f.ckb_danfe_a_imprimir[i].checked) {
				if (!f.ckb_danfe_a_imprimir[i].disabled) intQtdeDanfeAImprimir++;
			}
		}

		if ((intQtdeDanfeImpressa == 0) && (intQtdeDanfeAImprimir == 0)) {
			alert('Nenhum pedido foi selecionado!!');
			return;
		}
	}

	window.status = "Aguarde ...";
	f.action = "RelSolicitacaoColetasGravaDados.asp";
	f.submit();
}

function fRELMarcarTodos(f) {
	var i;
	for (i = 0; i < f.ckb_danfe_a_imprimir.length; i++) {
		if (!f.ckb_danfe_a_imprimir[i].disabled) f.ckb_danfe_a_imprimir[i].checked=true;
	}
}

function fOrcamentistaEIndicadorConsultaView(apelido) {
    sizeDivOrcamentistaEIndicadorConsultaView();
    $("#iframeOrcamentistaEIndicadorConsultaView").attr("src", "OrcamentistaEIndicadorConsultaView.asp?id_selecionado=" + encodeURIComponent(apelido));
    $("#divOrcamentistaEIndicadorConsultaView").fadeIn();
}

function fechaDivOrcamentistaEIndicadorConsultaView() {
    $("#divOrcamentistaEIndicadorConsultaView").fadeOut();
    $("#iframeOrcamentistaEIndicadorConsultaView").attr("src", "");
}

function sizeDivOrcamentistaEIndicadorConsultaView() {
    var newHeight = $(document).height() + "px";
    $("#divOrcamentistaEIndicadorConsultaView").css("height", newHeight);
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
body{
	overflow-y:scroll;
}
.TdCodCoresMargem{
	width:20px;
}
#divOrcamentistaEIndicadorConsultaView
{
    position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoOrcamentistaEIndicadorConsultaView
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
#divInternoOrcamentistaEIndicadorConsultaView.divFixo
{
    position:fixed;
	top:6%;
}
#iframeOrcamentistaEIndicadorConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
#imgFechaDivOrcamentistaEIndicadorConsultaView
{
    position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
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

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- Nº DO PEDIDO P/ CONSULTAR O PEDIDO AO CLICAR SOBRE O NÚMERO -->
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<!-- FILTROS -->
<input type="hidden" name="rb_loja" id="rb_loja" value="<%=rb_loja%>" />
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>" />
<input type="hidden" name="c_loja_de" id="c_loja_de" value="<%=c_loja_de%>" />
<input type="hidden" name="c_loja_ate" id="c_loja_ate" value="<%=c_loja_ate%>" />
<input type="hidden" name="rb_tipo_relatorio" id="rb_tipo_relatorio" value="<%=rb_tipo_relatorio%>" />
<input type="hidden" name="c_filtro_transportadora" id="c_filtro_transportadora" value="<%=c_filtro_transportadora%>" />
<input type="hidden" name="c_filtro_dt_entrega" id="c_filtro_dt_entrega" value="<%=c_filtro_dt_entrega%>" />
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>" />
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="ckb_emitir_nfe" id="ckb_emitir_nfe" value="">
<input type="hidden" name="ckb_gravar_transp_e_dt_entrega" id="ckb_gravar_transp_e_dt_entrega" value="">
<input type="hidden" name="ckb_danfe_impressa" id="ckb_danfe_impressa" value="">
<input type="hidden" name="ckb_danfe_a_imprimir" id="ckb_danfe_a_imprimir" value="">
<input type="hidden" name="c_numero_pedido" id="c_numero_pedido" value="">
<input type="hidden" name="c_pedido_transportadora" id="c_pedido_transportadora" value="">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="918" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Solicitação de Coletas</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='918' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>"

	if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then
		s = "Solicitação de Coleta"
	elseif rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
		s = "Pedidos Prontos para Romaneio"
	else
		s = ""
		end if
	
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Tipo de Relatório:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = c_nfe_emitente
    if s = "" then
        s = "N.I."
    else
        s = obtem_apelido_empresa_NFe_emitente(s)
        end if

	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>CD:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	select case rb_loja
		case "TODAS": s = "todas"
		case "UMA": s = c_loja
		case "FAIXA"
			s = ""
			s_aux = c_loja_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " a "
			s_aux = c_loja_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
		case else: s = ""
		end select
		
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Lojas:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
    s = c_fabricante_permitido
    if s = "" then
        s = "todos"
    else
        v_fabricante_permitido = split(c_fabricante_permitido, ", ")
        s = ""
        for i = Lbound(v_fabricante_permitido) to Ubound(v_fabricante_permitido)
            if s <> "" then s = s & ", "
		    s = s & x_fabricante(v_fabricante_permitido(i))
        next
        end if

	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Fabricantes <span style='color:green;'>Permitidos</span>:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"

    s = c_fabricante_proibido
    if s = "" then
        s = "N.I."
    else
        v_fabricante_proibido = split(c_fabricante_proibido, ", ")
        s = ""
        for i = Lbound(v_fabricante_proibido) to Ubound(v_fabricante_proibido)
            if s <> "" then s = s & ", "
		    s = s & x_fabricante(v_fabricante_proibido(i))
        next
        end if

	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Fabricantes <span style='color:red;'>Proibidos</span>:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"

    s = c_zona_permitida
    if s = "" then
        s = "todas"
    else
        v_zona_permitida = split(c_zona_permitida, ", ")
        s = ""
        for i = Lbound(v_zona_permitida) to Ubound(v_zona_permitida)
            if s <> "" then s = s & ", "
		    s = s & wms_deposito_zona_obtem_descricao(v_zona_permitida(i))
        next
        end if

	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Zonas <span style='color:green;'>Permitidas</span>:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"

    s = c_zona_proibida
    if s = "" then
        s = "N.I."
    else
        v_zona_proibida = split(c_zona_proibida, ", ")
        s = ""
        for i = Lbound(v_zona_proibida) to Ubound(v_zona_proibida)
            if s <> "" then s = s & ", "
		    s = s & wms_deposito_zona_obtem_descricao(v_zona_proibida(i))
        next
        end if

	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Zonas <span style='color:red;'>Proibidas</span>:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"

	s = c_filtro_transportadora
	if s = "" then 
		s = "todas"
	else
		if (s_filtro_nome_transportadora <> "") And (Ucase(s_filtro_nome_transportadora) <> Ucase(c_filtro_transportadora)) then s = s & "  (" & s_filtro_nome_transportadora & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Transportadora:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"
	
	if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
		s = c_filtro_dt_entrega
		if s = "" then s = "N.I."
		s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
					"<span class='N'>Data Coleta:&nbsp;</span></td><td align='left' valign='top'>" & _
					"<span class='N'>" & s & "</span></td></tr>"
		end if
	
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<% if (qtde_pedidos > 0) And (rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA) then %>
<br />
<table>
	<tr>
		<td align="right">
			<span class="C">Data de coleta</span>
		</td>
		<td align="left">
			<input class="Cc" name="c_dt_entrega" id="c_dt_entrega" maxlength="10" style="width:90px;margin-left:2px;" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) {fREL.c_transportadora.focus(); event.preventDefault();} filtra_data();">
		</td>
	</tr>
	<tr>
		<td align="right">
			<span class="C">Transportadora</span>
		</td>
		<td align="left">
			<select id="c_transportadora" name="c_transportadora" style="margin-left:2px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =transportadora_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="918" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br class="notPrint">

<table class="notPrint">
	<thead>
		<tr>
			<td colspan="2" align="center"><span class="N">Código de cores do relatório</span>&nbsp;&nbsp;<a href="javascript:TBodyCodCoresToggle();" title="Exibe detalhes sobre o código de cores do relatório"><img id="imgToggleCodCores" src="../IMAGEM/double-down-20.png" style="vertical-align:bottom;"/></a></td>
		</tr>
	</thead>
	<tbody id="tBodyCodCores">
		<tr>
			<td class="TdCodCoresMargem">&nbsp;</td>
			<td><span class="C" style="color:black;">Preto:</span><span class="Cn">situação inicial</span></td>
		</tr>
		<tr>
			<td class="TdCodCoresMargem">&nbsp;</td>
			<td><span class="C" style="color:blue;">Azul:</span><span class="Cn">possui transportadora e data da coleta preenchidos, mas falta solicitar a emissão da NFe</span></td>
		</tr>
		<tr>
			<td class="TdCodCoresMargem">&nbsp;</td>
			<td><span class="C" style="color:red;">Vermelho:</span><span class="Cn">aguardando a emissão da NFe ser concluída</span></td>
		</tr>
		<tr>
			<td class="TdCodCoresMargem">&nbsp;</td>
			<td><span class="C" style="color:darkorange;">Laranja:</span><span class="Cn">situação não prevista (ex: solicitação de emissão da NFe já realizada, mas os campos transportadora e/ou data da coleta estão em branco)</span></td>
		</tr>
	</tbody>
</table>
<% end if %>

<% if (qtde_pedidos > 0) And (rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO) then %>
<br />
<table>
	<tr>
		<td align="right">
		<input name="bMarcarPDFs" id="bMarcarPDFs" type="button" class="Button" onclick="fRELMarcarTodos(fREL)" value="Marcar todos os PDFs para impressão" title="assinala todos os PDFs" style="margin-left:6px;margin-bottom:10px">
		</td>
	</tr>
</table>
<% end if %>

<!-- ************   SEPARADOR   ************ -->
<table width="918" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="918" cellspacing="0">
<tr>
	<% if qtde_pedidos > 0 then %>
	<td align="left">
		<a name="bVOLTA" id="bVOLTA" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="left">&nbsp;</td>
	<td align="right">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0"></a>
	</td>
	<% else %>
	<td align="center">
		<a name="bVOLTA" id="bVOLTA" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<% end if %>
</tr>
</table>
</form>

</center>

<div id="divOrcamentistaEIndicadorConsultaView"><center><div id="divInternoOrcamentistaEIndicadorConsultaView"><img id="imgFechaDivOrcamentistaEIndicadorConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframeOrcamentistaEIndicadorConsultaView"></iframe></div></center></div>
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
