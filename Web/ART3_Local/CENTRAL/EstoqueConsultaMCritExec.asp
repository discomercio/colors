<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================================
'	  E S T O Q U E C O N S U L T A M C R I T E X E C . A S P
'     =======================================================
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

	const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
	const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
	const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
	const MSO_NUMBER_FORMAT_TEXTO = "\@"
    
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISS�O DE ACESSO DO USU�RIO
	if Not (operacao_permitida(OP_CEN_REL_REGISTROS_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) Or _
			operacao_permitida(OP_CEN_EDITA_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas)) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_fabricante, s_nome_fabricante, s_produto, s_nome_produto, s_nome_produto_html, s_cadastrado_por
	dim s_entrada_de, s_entrada_ate, ckb_especial, ckb_saldo, ckb_compras, ckb_kit, ckb_devolucao, s_nf_entrada_de, s_nf_entrada_ate
	dim rb_saida
    dim c_empresa

	s_fabricante = retorna_so_digitos(Request.Form("c_fabricante"))
	if s_fabricante <> "" then s_fabricante = normaliza_codigo(s_fabricante, TAM_MIN_FABRICANTE)
	s_produto = UCase(Trim(Request.Form("c_produto")))
	s_cadastrado_por = UCase(Trim(Request.Form("c_cadastrado_por")))
	s_entrada_de = Trim(Request.Form("c_entrada_de"))
	s_entrada_ate = Trim(Request.Form("c_entrada_ate"))
	ckb_especial = Trim(Request.Form("ckb_especial"))
	ckb_saldo = Trim(Request.Form("ckb_saldo"))
	ckb_compras = Trim(Request.Form("ckb_compras"))
	ckb_kit = Trim(Request.Form("ckb_kit"))
	ckb_devolucao = Trim(Request.Form("ckb_devolucao"))
	s_nf_entrada_de = Trim(Request.Form("c_nf_entrada_de"))
	s_nf_entrada_ate = Trim(Request.Form("c_nf_entrada_ate"))
	rb_saida = Ucase(Trim(Request.Form("rb_saida")))
	c_empresa = Trim(Request.Form("c_empresa"))

	alerta = ""
	if (s_produto<>"") And (Not IsEAN(s_produto)) then
		if s_fabricante = "" then alerta = "PARA PESQUISAR PELO C�DIGO INTERNO DO PRODUTO � NECESS�RIO ESPECIFICAR O FABRICANTE."
		end if
	
	if alerta = "" then
		if s_fabricante <> "" then
			s_nome_fabricante = fabricante_descricao(s_fabricante)
		else
			s_nome_fabricante = ""
			end if
		
		if s_produto <> "" then
			s_nome_produto = produto_descricao(s_fabricante, s_produto)
			s_nome_produto_html = produto_formata_descricao_em_html(produto_descricao_html(s_fabricante, s_produto))
		else
			s_nome_produto = ""
			s_nome_produto_html = ""
			end if
		end if

'	Per�odo de consulta est� restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
		if alerta = "" then
			strDtRefDDMMYYYY = s_entrada_de
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			strDtRefDDMMYYYY = s_entrada_ate
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inv�lida para consulta: " & strDtRefDDMMYYYY & ".  O per�odo de consulta n�o pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			if s_entrada_de = "" then s_entrada_de = strMinDtInicialFiltroPeriodoDDMMYYYY
			end if
		
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if

	dim blnSaidaExcel
	blnSaidaExcel = False
	if alerta = "" then
		if rb_saida = "XLS" then
			blnSaidaExcel = True
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=EntradaEstoque_" & formata_data_yyyymmdd(Now) & "_" & formata_hora_hhnnss(Now) & ".xls"
			Response.Write "<h2>Registros Entrada Estoque</h2>"

			s = ""
			s = s_fabricante
			if (s<>"") And (s_nome_fabricante<>"") then s = s & " - " & s_nome_fabricante
			if (s<>"") then
				Response.Write "Fabricante: " & s
				Response.Write "<br>"
				end if

			s = ""
			s = s_produto
			if (s<>"") And (s_nome_produto<>"") then s = s & " - " & s_nome_produto
			if (s<>"") then
				Response.Write "Produto: " & s
				Response.Write "<br>"
				end if

			s = ""
			s = Trim(c_empresa)
			if s = "" then
				s = "N.I."
			else
				s = obtem_apelido_empresa_NFe_emitente(s)
				end if
			if (s<>"") then
				Response.Write "Empresa: " & s
				Response.Write "<br>"
				end if

			s = ""
			s = s_cadastrado_por
			if (s<>"") then
				Response.Write "Cadastrado por: " & s
				Response.Write "<br>"
				end if

			s = ""
			s_aux = s_entrada_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " e "
			s_aux = s_entrada_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
			if (s<>"") then
				Response.Write "Data de entrada no estoque entre " & s
				Response.Write "<br>"
				end if

			s = ""
			if ckb_compras <> "" then
				if s <> "" then s = s & ", "
				s = s & "Compras de fornecedor"
				end if
			if ckb_especial <> "" then
				if s <> "" then s = s & ", "
				s = s & "Entrada especial"
				end if
			if ckb_kit <> "" then
				if s <> "" then s = s & ", "
				s = s & "Kit"
				end if
			if ckb_devolucao <> "" then
				if s <> "" then s = s & ", "
				s = s & "Devolu��o"
				end if
			if s <> "" then
				Response.Write "Tipo de cadastramento: " & s
				Response.Write "<br>"
				end if

			s = ""
			if ckb_saldo = "TODOS" then
				if s <> "" then s = s & ", "
				s = s & "Todos"
				end if
			if ckb_saldo = "COM_SALDO" then
				if s <> "" then s = s & ", "
				s = s & "Somente produtos com saldo dispon�vel"
				end if
			if ckb_saldo = "SEM_SALDO" then
				if s <> "" then s = s & ", "
				s = s & "Somente produtos sem saldo dispon�vel"
				end if
			if s <> "" then
				Response.Write "Saldo de produtos: " & s
				Response.Write "<br>"
				end if

			s = ""
			s_aux = s_nf_entrada_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " e "
			s_aux = s_nf_entrada_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
			if (s<>"") then
				Response.Write "Data de NF entrada entre " & s
				Response.Write "<br>"
				end if

			s = "Emiss�o: " & formata_data_hora(Now)
			Response.Write s
			Response.Write "<br><br>"
			executa_consulta()
			Response.End
			end if
		end if




' ________________________________
' EXECUTA CONSULTA
'
Sub executa_consulta ()
dim s, h, x, s_sql, s_where, s_where_tipo_or, s_where_tipo_and, n_reg, rs, s_link_open, s_link_close, s_nowrap
dim w_dt_entrada, w_documento, w_dt_nf_entrada, w_empresa, w_fabricante, w_produto, w_qtde, w_saldo, w_vl_unitario, w_vl_referencia, w_operador

	if blnSaidaExcel then
		w_dt_entrada = 90
		w_documento = 100
		w_dt_nf_entrada = 90
		w_empresa = 100
		w_fabricante = 150
		w_produto = 250
		w_qtde = 50
		w_saldo = 50
		w_vl_unitario = 100
		w_vl_referencia = 100
		w_operador = 100
	else
		w_dt_entrada = 50
		w_documento = 80
        w_dt_nf_entrada = 50
		w_empresa = 70
		w_fabricante = 110
		w_produto = 200
		w_qtde = 45
		w_saldo = 45
		w_vl_unitario = 90
		w_vl_referencia = 90
		w_operador = 80
		end if
	
	h = "<table class='Q' style='border-bottom:0px;' cellspacing=0 cellpadding=0>" & chr(13)
	if blnSaidaExcel then
		h = h & _
			"<TR style='background:azure'>" & _
			chr(13) & "	<TD style='width:" & w_dt_entrada & "px;' align='center'><P class='R' style='font-weight:bold;'>Entrada</P></TD>" & _
			chr(13) & "	<TD style='width:" & w_documento & "px;' NOWRAP><P class='R' style='font-weight:bold;margin-right:2pt;'>Documento</P></TD>" & _
			chr(13) & "	<TD style='width:" & w_dt_nf_entrada & "px;' align='center'><P class='R' style='font-weight:bold;'>Data NF</P></TD>" & _
			chr(13) & "	<TD style='width:" & w_empresa & "px;'><P class='R' style='font-weight:bold;'>Empresa</P></TD>" & _
			chr(13) & "	<TD style='width:" & w_fabricante & "px;'><P class='R' style='font-weight:bold;'>Fabricante</P></TD>" & _
			chr(13) & "	<TD style='width:" & w_produto & "px;'><P class='R' style='font-weight:bold;'>Produto</P></TD>" & _
			chr(13) & "	<TD style='width:" & w_qtde & "px;' align='right'><P class='R' style='font-weight:bold;text-align:right;margin-left:2pt;margin-right:2pt;'>Qtde</P></TD>" & _
			chr(13) & "	<TD style='width:" & w_saldo & "px;' align='right'><P class='R' style='font-weight:bold;text-align:right;margin-left:2pt;margin-right:2pt;'>Saldo</P></TD>" & _
			chr(13) & "	<TD style='width:" & w_vl_unitario & "px;' align='right' NOWRAP><P class='R' style='font-weight:bold;text-align:right;margin-left:2pt;margin-right:2pt;'>Valor Unit</P></TD>" & _
			chr(13) & "	<TD style='width:" & w_vl_referencia & "px;' align='right' NOWRAP><P class='R' style='font-weight:bold;text-align:right;margin-left:2pt;margin-right:2pt;'>Valor Ref</P></TD>" & _
			chr(13) & "	<TD style='width:" & w_operador & "px;' NOWRAP><P class='R' style='font-weight:bold;margin-right:2pt;'>Operador</P></TD>" & _
			"</TR>" & chr(13)
	else
		h = h & _
			"<TR style='background:azure'>" & _
			chr(13) & "	<TD align='center' valign='bottom' class='MD MB' style='width:" & w_dt_entrada & "px;'><P class='R' style='font-weight:bold;'>entrada</P></TD>" & _
			chr(13) & "	<TD valign='bottom' class='MD MB' style='width:" & w_documento & "px;' NOWRAP><P class='R' style='font-weight:bold;margin-right:2pt;'>documento</P></TD>" & _
			chr(13) & "	<TD align='center' valign='bottom' class='MD MB' style='width:" & w_dt_nf_entrada & "px;' NOWRAP><P class='R' style='font-weight:bold;margin-right:2pt;'>data nf</P></TD>" & _
			chr(13) & "	<TD valign='bottom' class='MD MB' style='width:" & w_empresa & "px;'><P class='R' style='font-weight:bold;'>empresa</P></TD>" & _
			chr(13) & "	<TD valign='bottom' class='MD MB' style='width:" & w_fabricante & "px;'><P class='R' style='font-weight:bold;'>fabricante</P></TD>" & _
			chr(13) & "	<TD valign='bottom' class='MD MB' style='width:" & w_produto & "px;'><P class='R' style='font-weight:bold;'>produto</P></TD>" & _
			chr(13) & "	<TD valign='bottom' class='MD MB' style='width:" & w_qtde & "px;' align='right'><P class='R' style='font-weight:bold;text-align:right;margin-left:2pt;margin-right:2pt;'>qtde</P></TD>" & _
			chr(13) & "	<TD valign='bottom' class='MD MB' style='width:" & w_saldo & "px;' align='right'><P class='R' style='font-weight:bold;text-align:right;margin-left:2pt;margin-right:2pt;'>saldo</P></TD>" & _
			chr(13) & "	<TD valign='bottom' class='MD MB' style='width:" & w_vl_unitario & "px;' align='right' NOWRAP><P class='R' style='font-weight:bold;text-align:right;margin-left:2pt;margin-right:2pt;'>valor unit</P></TD>" & _
			chr(13) & "	<TD valign='bottom' class='MD MB' style='width:" & w_vl_referencia & "px;' align='right' NOWRAP><P class='R' style='font-weight:bold;text-align:right;margin-left:2pt;margin-right:2pt;'>valor ref</P></TD>" & _
			chr(13) & "	<TD valign='bottom' class='MB' style='width:" & w_operador & "px;' NOWRAP><P class='R' style='font-weight:bold;margin-right:2pt;'>operador</P></TD>" & _
			"</TR>" & chr(13)
		end if
	
	s_sql = "SELECT" & _
			" t_ESTOQUE.id_estoque, t_ESTOQUE.data_entrada," & _
			" t_ESTOQUE.id_nfe_emitente, t_NFe_EMITENTE.apelido AS empresa_apelido," & _
			" t_ESTOQUE.fabricante," & _
			" t_ESTOQUE.documento, t_ESTOQUE.usuario," & _
			" t_ESTOQUE_ITEM.produto, t_ESTOQUE_ITEM.preco_fabricante, t_ESTOQUE_ITEM.vl_custo2," & _
			" t_ESTOQUE_ITEM.qtde, t_ESTOQUE_ITEM.qtde_utilizada," & _
			" t_PRODUTO.descricao," & _
			" t_PRODUTO.descricao_html," & _
			" t_FABRICANTE.razao_social, t_FABRICANTE.nome," & _
			" t_ESTOQUE.data_emissao_NF_entrada" & _
            " FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
			" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
			" LEFT JOIN t_FABRICANTE ON (t_ESTOQUE.fabricante=t_FABRICANTE.fabricante)" & _
			" LEFT JOIN t_NFe_EMITENTE ON (t_ESTOQUE.id_nfe_emitente=t_NFe_EMITENTE.id)"

	s_where = ""
	if s_fabricante <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ESTOQUE.fabricante='" & s_fabricante & "')"
		end if

	if s_produto <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		if IsEAN(s_produto) then
			s_where = s_where & " (t_PRODUTO.ean='" & s_produto & "')"
		else
		'	PESQUISA PELO C�DIGO INTERNO: OBRIGA RESTRI��O PELO FABRICANTE, QUE � PARTE DA CHAVE PRIM�RIA DO PRODUTO
			s_where = s_where & " ((t_ESTOQUE_ITEM.fabricante='" & s_fabricante & "') AND (t_ESTOQUE_ITEM.produto='" & s_produto & "'))"
			end if
		end if

	if s_cadastrado_por <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ESTOQUE.usuario='" & s_cadastrado_por & "')"
		end if
	
	if s_entrada_de <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ESTOQUE.data_entrada >= " & bd_formata_data(StrToDate(s_entrada_de)) & ")"
		end if
	
	if s_entrada_ate <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ESTOQUE.data_entrada <= " & bd_formata_data(StrToDate(s_entrada_ate)) & ")"
		end if

    if c_empresa <> "" then
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_ESTOQUE.id_nfe_emitente = " & c_empresa & ")"
    end if
	
'	TIPO DE CADASTRAMENTO
	s_where_tipo_or = ""
	s_where_tipo_and = ""
	if ckb_especial <> "" then
		if s_where_tipo_or <> "" then s_where_tipo_or = s_where_tipo_or & " OR"
		s_where_tipo_or = s_where_tipo_or & " (t_ESTOQUE.entrada_especial<>0)"
	else
		if s_where_tipo_and <> "" then s_where_tipo_and = s_where_tipo_and & " AND"
		s_where_tipo_and = s_where_tipo_and & " (t_ESTOQUE.entrada_especial=0)"
		end if

	if ckb_kit <> "" then
		if s_where_tipo_or <> "" then s_where_tipo_or = s_where_tipo_or & " OR"
		s_where_tipo_or = s_where_tipo_or & " (t_ESTOQUE.kit<>0)"
	else
		if s_where_tipo_and <> "" then s_where_tipo_and = s_where_tipo_and & " AND"
		s_where_tipo_and = s_where_tipo_and & " (t_ESTOQUE.kit=0)"
		end if

	if ckb_devolucao <> "" then
		if s_where_tipo_or <> "" then s_where_tipo_or = s_where_tipo_or & " OR"
		s_where_tipo_or = s_where_tipo_or & " (t_ESTOQUE.devolucao_status<>0)"
	else
		if s_where_tipo_and <> "" then s_where_tipo_and = s_where_tipo_and & " AND"
		s_where_tipo_and = s_where_tipo_and & " (t_ESTOQUE.devolucao_status=0)"
		end if

	if ckb_compras <> "" then
		if s_where_tipo_or <> "" then s_where_tipo_or = s_where_tipo_or & " OR"
		s_where_tipo_or = s_where_tipo_or & " ((t_ESTOQUE.entrada_especial=0) AND (t_ESTOQUE.kit=0) AND (t_ESTOQUE.devolucao_status=0))"
	else
		if s_where_tipo_and <> "" then s_where_tipo_and = s_where_tipo_and & " AND"
		s_where_tipo_and = s_where_tipo_and & " (NOT ((t_ESTOQUE.entrada_especial=0) AND (t_ESTOQUE.kit=0) AND (t_ESTOQUE.devolucao_status=0)))"
		end if

	if s_where_tipo_or <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where_tipo_or = " (" & s_where_tipo_or & ")"
		s_where = s_where & s_where_tipo_or
		end if

	if s_where_tipo_and <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where_tipo_and = " (" & s_where_tipo_and & ")"
		s_where = s_where & s_where_tipo_and
		end if
	
	if ckb_saldo = "COM_SALDO" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " ((t_ESTOQUE_ITEM.qtde - t_ESTOQUE_ITEM.qtde_utilizada) > 0)"
	elseif ckb_saldo = "SEM_SALDO" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " ((t_ESTOQUE_ITEM.qtde - t_ESTOQUE_ITEM.qtde_utilizada) = 0)"
		end if

	if s_nf_entrada_de <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ESTOQUE.data_emissao_NF_entrada >= " & bd_formata_data(StrToDate(s_nf_entrada_de)) & ")"
		end if
	
	if s_nf_entrada_ate <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ESTOQUE.data_emissao_NF_entrada <= " & bd_formata_data(StrToDate(s_nf_entrada_ate)) & ")"
		end if

	if s_where <> "" then s_where = " WHERE" & s_where
	s_sql = s_sql & s_where
	s_sql = s_sql & " ORDER BY t_ESTOQUE.data_entrada, t_ESTOQUE.id_estoque, t_ESTOQUE.documento, t_ESTOQUE.fabricante, t_ESTOQUE_ITEM.sequencia"
	

'	EXECUTA CONSULTA
	set rs = cn.Execute( s_sql )
	
	x = h
	n_reg = 0
	do while Not rs.eof 
		n_reg = n_reg + 1
		if ((n_reg AND 1)=0) And (Not blnSaidaExcel) then
			x = x & "<TR style='background: #FFF0E0'>"
		else
			x = x & "<TR>"
			end if

		if blnSaidaExcel then
			s_link_open = ""
			s_link_close = ""
		else
			s_link_open = "<a href='javascript:fConcluir(" & chr(34) & Trim("" & rs("id_estoque")) & chr(34) & _
					 ")' title='clique para consultar este registro de entrada no estoque'>"
			s_link_close = "</a>"
			end if
		
	'	DATA ENTRADA
		x = x & chr(13) & "	<TD align='center' valign='middle' class='MDB' NOWRAP style='width:" & w_dt_entrada & "px;'><P class='Cn'>" & s_link_open & formata_data(rs("data_entrada")) & s_link_close & "</P>" & "</TD>"
		
	'	DOCUMENTO
		if blnSaidaExcel then s_nowrap = " NOWRAP" else s_nowrap = ""
		s = Trim("" & rs("documento"))
		if Not blnSaidaExcel then
			if s = "" then s = "&nbsp;"
			end if
		x = x & chr(13) & "	<TD valign='middle' class='MDB'" & s_nowrap & " style='width:" & w_documento & "px;'><P class='Cn' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_link_open & s & s_link_close & "</P></TD>"

	'	DATA NF
		'if blnSaidaExcel then s_nowrap = " NOWRAP" else s_nowrap = ""
		's = Trim("" & rs("documento"))
		'if Not blnSaidaExcel then
		'	if s = "" then s = "&nbsp;"
		'	end if
		x = x & chr(13) & "	<TD valign='middle' class='MDB'" & s_nowrap & " style='width:" & w_dt_nf_entrada & "px;'><P class='Cn'>" & s_link_open & formata_data(rs("data_emissao_NF_entrada")) & s_link_close & "</P></TD>"

	'	EMPRESA
		if blnSaidaExcel then s_nowrap = " NOWRAP" else s_nowrap = ""
		s = Trim("" & rs("empresa_apelido"))
		x = x & chr(13) & "	<TD valign='middle' class='MDB'" & s_nowrap & " style='width:" & w_empresa & "px;'><P class='Cn' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s & "</P></TD>"

	'	FABRICANTE
		if blnSaidaExcel then s_nowrap = " NOWRAP" else s_nowrap = ""
		s = Trim("" & rs("nome"))
		if s = "" then s = Trim("" & rs("razao_social"))
		if s <> "" then 
			s = iniciais_em_maiusculas(s)
			s = " - " & s
			end if
		s = Trim("" & rs("fabricante")) & s
		x = x & chr(13) & "	<TD valign='middle' class='MDB'" & s_nowrap & " style='width:" & w_fabricante & "px;'><P class='Cn' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s & "</P></TD>"

	'	PRODUTO
		if blnSaidaExcel then s_nowrap = " NOWRAP" else s_nowrap = ""
		s = Trim("" & rs("descricao_html"))
		if s <> "" then 
			s = produto_formata_descricao_em_html(s)
			s = " - " & s
			end if
		s = Trim("" & rs("produto")) & s
		x = x & chr(13) & "	<TD valign='middle' class='MDB'" & s_nowrap & " style='width:" & w_produto & "px;'><P class='Cn' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s & "</P></TD>"

	'	QTDE
		x = x & chr(13) & "	<TD align='right' valign='middle' class='MDB' NOWRAP style='width:" & w_qtde & "px;'><P class='Cnd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & Cstr(rs("qtde")) & "</P></TD>"

	'	SALDO
		x = x & chr(13) & "	<TD align='right' valign='middle' class='MDB' NOWRAP style='width:" & w_saldo & "px;'><P class='Cnd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & Cstr(CLng(rs("qtde"))-CLng(rs("qtde_utilizada"))) & "</P></TD>"

	'	VALOR UNIT�RIO
		x = x & chr(13) & "	<TD align='right' valign='middle' class='MDB' NOWRAP style='width:" & w_vl_unitario & "px;'><P class='Cnd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(rs("preco_fabricante")) & "</P></TD>"

	'	VALOR REFER�NCIA
		x = x & chr(13) & "	<TD align='right' valign='middle' class='MDB' NOWRAP style='width:" & w_vl_referencia & "px;'><P class='Cnd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(rs("vl_custo2")) & "</P></TD>"

	'	OPERADOR
		if blnSaidaExcel then s_nowrap = " NOWRAP" else s_nowrap = ""
		x = x & chr(13) & "	<TD valign='middle' class='MB'" & s_nowrap & " style='width:" & w_operador & "px;'><P class='Cn' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & Trim("" & rs("usuario")) & "</P></TD>"

		x = x & "</TR>" & chr(13)

		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if

		rs.MoveNext
		loop

	if n_reg = 0 then 
		x = h & "<TR NOWRAP >" & _
			"<TD colspan='10' align='center' class='MB'><P class='ALERTA'>&nbsp;NENHUM REGISTRO DO ESTOQUE SATISFAZ AS CONDI��ES ESPECIFICADAS&nbsp;</P></TD>" & _
			"</TR>"
		end if

	x = x & "</TABLE>"

	Response.write x

	if rs.State <> 0 then rs.Close
	set rs=nothing
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

function fConcluir ( id ) {
	fESTOQ.action = "EstoqueConsultaEAN.asp";
	fESTOQ.estoque_selecionado.value = id;
	fESTOQ.submit();
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


<style TYPE="text/css">
a
{
	text-decoration: none;
	color: black;
}
#ckb_especial_aux {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#ckb_saldo_aux {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#ckb_compras_aux {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#ckb_kit_aux {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#ckb_devolucao_aux {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
</style>


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;FONT-WEIGHT:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<BR><BR>
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
<!-- **********  P�GINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Conclu�do';">

<center>

<form id="fESTOQ" name="fESTOQ" METHOD="POST">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type=HIDDEN name="estoque_selecionado" id="estoque_selecionado" value="">
<input type=HIDDEN name="ckb_especial" id="ckb_especial" value="<%=ckb_especial%>">
<input type=HIDDEN name="ckb_saldo" id="ckb_saldo" value="<%=ckb_saldo%>">
<input type=HIDDEN name="c_fabricante" id="c_fabricante" value="<%=s_fabricante%>">
<input type=HIDDEN name="c_produto" id="c_produto" value="<%=s_produto%>">
<input type=HIDDEN name="ckb_compras" id="ckb_compras" value="<%=ckb_compras%>">
<input type=HIDDEN name="ckb_kit" id="ckb_kit" value="<%=ckb_kit%>">
<input type=HIDDEN name="ckb_devolucao" id="ckb_devolucao" value="<%=ckb_devolucao%>">

<!--  I D E N T I F I C A � � O   D A   T E L A  -->
<table width="884" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="RIGHT" vAlign="BOTTOM"><span class="PEDIDO">Registros Entrada Estoque</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PAR�METROS DA CONSULTA MULTICRIT�RIOS  -->
<table class="Qx" cellSpacing="0">
<!--  EMPRESA  -->
	<tr bgColor="#FFFFFF">
	<td class="MT" NOWRAP><span class="PLTe">Empresa</span>
        <%	s = c_empresa
			if (s<>"") then s = obtem_apelido_empresa_NFe_emitente(c_empresa)  %>
		<br><input name="c_empresa" id="c_empresa" readonly tabindex=-1 class="PLLe" style="margin-left:2pt;width:100px;"
				value="<%=s%>"></td>
	</tr>

<!--  FABRICANTE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">Fabricante</span>
		<%	s = s_fabricante
			if (s<>"") And (s_nome_fabricante<>"") then s = s & " - " & s_nome_fabricante %>
		<br><input name="c_fabricante_aux" id="c_fabricante_aux" readonly tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=s%>"></td>
	</tr>
	
<!--  PRODUTO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">Produto</span>
		<%	s = s_produto
			if (s<>"") And (s_nome_produto_html<>"") then s = s & " - " & s_nome_produto_html %>
		<br>
		<span class="PLLe" style="width:460px;margin-left:2pt;"><%=s%></span>
		<%	s = s_produto
			if (s<>"") And (s_nome_produto<>"") then s = s & " - " & s_nome_produto %>
		<input type=hidden name="c_produto_aux" id="c_produto_aux" value="<%=s%>">
	</td>
	</tr>

<!--  CADASTRADO POR  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">Cadastrado por</span>
		<br><input name="c_cadastrado_por" id="c_cadastrado_por" readonly tabindex=-1 class="PLLe" style="margin-left:2pt;width:100px;"
				value="<%=s_cadastrado_por%>"></td>
	</tr>

<!--  PER�ODO DE ENTRADA NO ESTOQUE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">Data de Entrada no Estoque Entre</span>
		<br><input name="c_entrada_de" id="c_entrada_de" readonly tabindex=-1 class="PLLc" style="margin-left:2pt;width:150px;"
						value="<%=s_entrada_de%>"
			><span class="PLTe">&nbsp;e&nbsp;</span><input name="c_entrada_ate" id="c_entrada_ate" readonly tabindex=-1 class="PLLc" style="margin-left:2pt;width:150px;"
						value="<%=s_entrada_ate%>"></td>
	</tr>

<!--  TIPO DE CADASTRAMENTO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">Tipo de Cadastramento</span>
		<br><input type="checkbox" disabled tabindex="-1" id="ckb_compras_aux" name="ckb_compras_aux" value="COMPRAS_ON"
				<%if ckb_compras<>"" then Response.Write " checked"%>><span class="C" style="cursor:default;vertical-align:bottom;">Compras de Fornecedor</span>
		<br><input type="checkbox" disabled tabindex="-1" id="ckb_especial_aux" name="ckb_especial_aux" value="ESPECIAL_ON"
				<%if ckb_especial<>"" then Response.Write " checked"%>><span class="C" style="cursor:default;vertical-align:bottom;">Entrada Especial</span>
		<br><input type="checkbox" disabled tabindex="-1" id="ckb_kit_aux" name="ckb_kit_aux" value="KIT_ON"
				<%if ckb_kit<>"" then Response.Write " checked"%>><span class="C" style="cursor:default;vertical-align:bottom;">Kit</span>
		<br><input type="checkbox" disabled tabindex="-1" id="ckb_devolucao_aux" name="ckb_devolucao_aux" value="DEVOLUCAO_ON"
				<%if ckb_devolucao<>"" then Response.Write " checked"%>><span class="C" style="cursor:default;vertical-align:bottom;">Devolu��o</span>
	</td>
	</tr>

<!--  SOMENTE PRODUTOS COM SALDO DISPON�VEL  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">Saldo de Produtos</span>
		<br><input type="radio" disabled tabindex="-1" id="ckb_saldo_aux" name="ckb_saldo_aux" value="TODOS"
				<%if ckb_saldo="TODOS" then Response.Write " checked"%>><span class="C" style="cursor:default;margin-right:10pt;vertical-align:bottom;">Todos</span>
		<br><input type="radio" disabled tabindex="-1" id="ckb_saldo_aux" name="ckb_saldo_aux" value="COM_SALDO"
				<%if ckb_saldo="COM_SALDO" then Response.Write " checked"%>><span class="C" style="cursor:default;margin-right:10pt;vertical-align:bottom;">Somente Produtos Com Saldo Dispon�vel</span>
		<br><input type="radio" disabled tabindex="-1" id="ckb_saldo_aux" name="ckb_saldo_aux" value="SEM_SALDO"
				<%if ckb_saldo="SEM_SALDO" then Response.Write " checked"%>><span class="C" style="cursor:default;margin-right:10pt;vertical-align:bottom;">Somente Produtos Sem Saldo Dispon�vel</span>
	</td>
	</tr>
</table>

<!--  RELAT�RIO  -->
<br>
<% executa_consulta %>

<!-- ************   SEPARADOR   ************ -->
<table width="884" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="884" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a p�gina anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
