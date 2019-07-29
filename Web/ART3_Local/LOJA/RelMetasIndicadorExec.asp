<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelMetasIndicadorExec.asp
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

	Const COD_CONSULTA_POR_PERIODO_CADASTRO = "CADASTRO"
	Const COD_CONSULTA_POR_PERIODO_ENTREGA = "ENTREGA"
	
	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_REL_METAS_INDICADOR, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro
	dim c_vendedor
	dim rb_periodo
	dim c_dt_cadastro_mes, c_dt_cadastro_ano, c_dt_entregue_mes, c_dt_entregue_ano

	alerta = ""

	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	rb_periodo = Trim(Request.Form("rb_periodo"))
	c_dt_cadastro_mes = retorna_so_digitos(Trim(Request.Form("c_dt_cadastro_mes")))
	c_dt_cadastro_ano = retorna_so_digitos(Trim(Request.Form("c_dt_cadastro_ano")))
	c_dt_entregue_mes = retorna_so_digitos(Trim(Request.Form("c_dt_entregue_mes")))
	c_dt_entregue_ano = retorna_so_digitos(Trim(Request.Form("c_dt_entregue_ano")))

	if alerta = "" then
		if rb_periodo = "" then
			alerta = "Selecione o tipo de consulta: 'por pedidos cadastrados' ou 'por pedidos entregues'"
		elseif (rb_periodo <> COD_CONSULTA_POR_PERIODO_CADASTRO) and (rb_periodo <> COD_CONSULTA_POR_PERIODO_ENTREGA) then
			alerta = "Opção inválida para tipo de período de consulta."
			end if
		end if
	
	if alerta = "" then
		if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
			if Not IsNumeric(c_dt_cadastro_mes) then
				alerta = "PEDIDOS CADASTRADOS EM: MÊS INVÁLIDO."
			elseif Not IsNumeric(c_dt_cadastro_ano) then
				alerta = "PEDIDOS CADASTRADOS EM: ANO INVÁLIDO."
			elseif (CLng(c_dt_cadastro_mes)<=0) Or (CLng(c_dt_cadastro_mes)>12) then
				alerta = "PEDIDOS CADASTRADOS EM: MÊS INVÁLIDO."
			elseif (CLng(c_dt_cadastro_ano)<2000) then
				alerta = "PEDIDOS CADASTRADOS EM: ANO INVÁLIDO."
			elseif (c_dt_cadastro_ano & c_dt_cadastro_mes) > Left(formata_data_yyyymmdd(Date),6) then
				alerta = "PEDIDOS CADASTRADOS EM: NÃO É POSSÍVEL CONSULTAR MÊS FUTURO"
				end if
			end if
		end if

	if alerta = "" then
		if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
			if Not IsNumeric(c_dt_entregue_mes) then
				alerta = "PEDIDOS ENTREGUES EM: MÊS INVÁLIDO."
			elseif Not IsNumeric(c_dt_entregue_ano) then
				alerta = "PEDIDOS ENTREGUES EM: ANO INVÁLIDO."
			elseif (CLng(c_dt_entregue_mes)<=0) Or (CLng(c_dt_entregue_mes)>12) then
				alerta = "PEDIDOS ENTREGUES EM: MÊS INVÁLIDO."
			elseif (CLng(c_dt_entregue_ano)<2000) then
				alerta = "PEDIDOS ENTREGUES EM: ANO INVÁLIDO."
			elseif (c_dt_entregue_ano & c_dt_entregue_mes) > Left(formata_data_yyyymmdd(Date),6) then
				alerta = "PEDIDOS ENTREGUES EM: NÃO É POSSÍVEL CONSULTAR MÊS FUTURO"
				end if
			end if
		end if

'	Período de consulta está restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	if operacao_permitida(OP_LJA_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
		if alerta = "" then
			if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
				strDtRefDDMMYYYY = "01/" & c_dt_cadastro_mes & "/" & c_dt_cadastro_ano
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if
			
			if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
				strDtRefDDMMYYYY = "01/" & c_dt_entregue_mes & "/" & c_dt_entregue_ano
				if strDtRefDDMMYYYY <> "" then
					if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
						alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
						end if
					end if
				end if
			end if
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
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
dim dtInicio, dtTermino
dim s_where_base, s_where_comum, s_where_venda, s_where_devolucao
dim vl_saldo, vl_efetuado, vl_meta, s_cor
dim intQtdeVendedores, intQtdeSubTotalIndicadores, intQtdeTotalIndicadores
dim s, s_aux, s_sql, strSqlVenda, strSqlDevolucao
dim x, cab_table, cab, vendedor_a, n_reg, n_reg_total
dim vlTotalMeta, vlSubTotalMeta, vlTotalEfetuado, vlSubTotalEfetuado, vlTotalSaldo, vlSubTotalSaldo

	if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
		dtInicio = StrToDate("01/" & c_dt_cadastro_mes & "/" & c_dt_cadastro_ano)
		dtTermino = DateAdd("m", 1, dtInicio)
		end if

	if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
		dtInicio = StrToDate("01/" & c_dt_entregue_mes & "/" & c_dt_entregue_ano)
		dtTermino = DateAdd("m", 1, dtInicio)
		end if
		
'	CRITÉRIO DA CONSULTA BASE
	s_where_base = ""
	if c_vendedor <> "" then
		if s_where_base <> "" then s_where_base = s_where_base & " AND"
		s_where_base = s_where_base & " (vendedor = '" & c_vendedor & "')"
	else
		if s_where_base <> "" then s_where_base = s_where_base & " AND"
		s_where_base = s_where_base & " " & SCHEMA_BD & ".UsuarioPossuiAcessoLoja(vendedor, '" & loja & "') = 'S'"
		end if
	
'	CRITÉRIOS COMUNS (VENDAS E DEVOLUCOES)
	s_where_comum = " (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')"

'	CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	s_where_venda = ""
	
	if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
		if IsDate(dtInicio) then
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (data >= " & bd_formata_data(dtInicio) & ")"
			end if

		if IsDate(dtTermino) then
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (data < " & bd_formata_data(dtTermino) & ")"
			end if
		end if
	
	if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
		if IsDate(dtInicio) then
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (entregue_data >= " & bd_formata_data(dtInicio) & ")"
			end if

		if IsDate(dtTermino) then
			if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
			s_where_venda = s_where_venda & " (entregue_data < " & bd_formata_data(dtTermino) & ")"
			end if
		end if
	
'	CRITÉRIOS PARA DEVOLUÇÕES
	s_where_devolucao = ""
	if IsDate(dtInicio) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (devolucao_data >= " & bd_formata_data(dtInicio) & ")"
		end if

	if IsDate(dtTermino) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (devolucao_data < " & bd_formata_data(dtTermino) & ")"
		end if
	
'	AJUNTA COM OS CRITÉRIOS COMUNS
	if s_where_comum <> "" then
		if s_where_venda <> "" then s_where_venda = " AND " & s_where_venda
		s_where_venda = s_where_comum & s_where_venda

		if s_where_devolucao <> "" then s_where_devolucao = " AND " & s_where_devolucao
		s_where_devolucao = s_where_comum & s_where_devolucao
		end if

	if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
	s_where_venda = s_where_venda & " (indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)"

	if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
	s_where_devolucao = s_where_devolucao & " (indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)"
	
	if s_where_base <> "" then s_where_base = " WHERE" & s_where_base
	if s_where_venda <> "" then s_where_venda = " WHERE" & s_where_venda
	if s_where_devolucao <> "" then s_where_devolucao = " WHERE" & s_where_devolucao

	strSqlVenda = _
		"SELECT" & _
			" Coalesce(Sum(qtde*preco_venda),0) AS vl_efetuado" & _
		" FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
		s_where_venda
		
	strSqlDevolucao = _
		"SELECT" & _
			" Coalesce(Sum(qtde*preco_venda),0) AS vl_devolucao" & _
		" FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido)" & _
		s_where_devolucao
		
	s_sql = "SELECT" & _
				" vendedor," & _
				" apelido," & _
				" vl_meta," & _
				" (" & strSqlVenda & ") AS vl_efetuado," & _
				" (" & strSqlDevolucao & ") AS vl_devolucao" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR" & _
			s_where_base & _
			" ORDER BY" & _
				" vendedor," & _
				" apelido"

  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE' valign='bottom' NOWRAP><P style='width:120px' class='R'>INDICADOR</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:80px;text-align:right;' class='R'>VL META</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:80px;text-align:right;' class='R'>VL EFETUADO</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:80px;text-align:right;' class='R'>SALDO</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	intQtdeVendedores = 0
	intQtdeTotalIndicadores = 0
	intQtdeSubTotalIndicadores = 0
	vlTotalMeta = 0
	vlTotalEfetuado = 0
	vlTotalSaldo = 0
	
	vendedor_a = "XXXXXXXXXXXX"
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE VENDEDOR?
		if Trim("" & r("vendedor"))<>vendedor_a then
			vendedor_a = Trim("" & r("vendedor"))
			intQtdeVendedores = intQtdeVendedores + 1
		  ' FECHA TABELA DO VENDEDOR ANTERIOR
			if n_reg_total > 0 then 
				x = x & _
						"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
						"		<TD class='MTE MD' NOWRAP><p class='Cd'>TOTAL</p></td>" & chr(13) & _
						"		<TD class='MTD'><p class='Cd'>" & formata_moeda(vlSubTotalMeta) & "</p></td>" & chr(13) & _
						"		<TD class='MTD'><p class='Cd' style='color:" & DecodificaCorHtmlValorMonetario(vlSubTotalEfetuado) & ";'>" & formata_moeda(vlSubTotalEfetuado) & "</p></td>" & chr(13) & _
						"		<TD class='MTD'><p class='Cd' style='color:" & DecodificaCorHtmlValorMonetario(vlSubTotalSaldo) & ";'>" & formata_moeda(vlSubTotalSaldo) & "</p></td>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
						"		<TD class='MT' colspan='4'><p class='Cc'>" & formata_inteiro(intQtdeSubTotalIndicadores) & " indicador(es)</p></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			n_reg = 0
			intQtdeSubTotalIndicadores = 0
			vlSubTotalMeta = 0
			vlSubTotalEfetuado = 0
			vlSubTotalSaldo = 0

			if n_reg_total > 0 then x = x & "<BR>"
			s = Trim("" & r("vendedor"))
			s_aux = x_usuario(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then x = x & "	<TR>" & chr(13) & _
									"		<TD class='MDTE' COLSPAN='4' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
									"	</tr>" & chr(13)
			x = x & cab
			end if
		

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		intQtdeSubTotalIndicadores = intQtdeSubTotalIndicadores + 1
		intQtdeTotalIndicadores = intQtdeTotalIndicadores + 1
		
		x = x & "	<TR NOWRAP>"  & chr(13)

	 '> INDICADOR
		x = x & "		<TD class='MDTE' valign='top'>" & _
				"<P class='Cn' style='font-weight:bold;'>" & _
				"<a href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("apelido")) & chr(34) & _
				")' title='clique para consultar o cadastro do indicador'>" & _
				Trim("" & r("apelido")) & _
				"</a>" & _
				"</P></TD>" & chr(13)

	 '> VALOR DA META
		vl_meta = r("vl_meta")
		s = formata_moeda(vl_meta)
		x = x & "		<TD class='MTD' valign='top'><P class='Cd'>" & s & "</P></TD>" & chr(13)

	 '> VALOR EFETUADO
		vl_efetuado = r("vl_efetuado")-r("vl_devolucao")
		s = formata_moeda(vl_efetuado)
		s_cor = DecodificaCorHtmlValorMonetario(vl_efetuado)
		x = x & "		<TD class='MTD' valign='top'><P class='Cd' style='color:" & s_cor & ";'>" & s & "</P></TD>" & chr(13)

	 '> SALDO
		vl_saldo = vl_efetuado - vl_meta
		s = formata_moeda(vl_saldo)
		s_cor = DecodificaCorHtmlValorMonetario(vl_saldo)
		x = x & "		<TD class='MTD' valign='top'><P class='Cd' style='color:" & s_cor & ";'>" & s & "</P></TD>" & chr(13)

		x = x & "	</TR>" & chr(13)

		vlSubTotalMeta = vlSubTotalMeta + vl_meta
		vlTotalMeta = vlTotalMeta + vl_meta
		vlSubTotalEfetuado = vlSubTotalEfetuado + vl_efetuado
		vlTotalEfetuado = vlTotalEfetuado + vl_efetuado
		vlSubTotalSaldo = vlSubTotalSaldo + vl_saldo
		vlTotalSaldo = vlTotalSaldo + vl_saldo
		
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
  ' MOSTRA TOTAL DO ÚLTIMO VENDEDOR
	if n_reg <> 0 then 
		x = x & _
				"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
				"		<TD class='MTE MD' NOWRAP><p class='Cd'>TOTAL</p></td>" & chr(13) & _
				"		<TD class='MTD'><p class='Cd'>" & formata_moeda(vlSubTotalMeta) & "</p></td>" & chr(13) & _
				"		<TD class='MTD'><p class='Cd' style='color:" & DecodificaCorHtmlValorMonetario(vlSubTotalEfetuado) & ";'>" & formata_moeda(vlSubTotalEfetuado) & "</p></td>" & chr(13) & _
				"		<TD class='MTD'><p class='Cd' style='color:" & DecodificaCorHtmlValorMonetario(vlSubTotalSaldo) & ";'>" & formata_moeda(vlSubTotalSaldo) & "</p></td>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
				"		<TD class='MT' colspan='4'><p class='Cc'>" & formata_inteiro(intQtdeSubTotalIndicadores) & " indicador(es)</p></td>" & chr(13) & _
				"	</TR>" & chr(13)
		
	'>	TOTAL GERAL
		if intQtdeVendedores > 1 then
			x = x & "	<TR>" & chr(13) & _
					"		<TD COLSPAN='4' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD COLSPAN='4' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
					"		<TD class='MTE MD' NOWRAP><p class='Cd'>TOTAL GERAL</p></td>" & chr(13) & _
					"		<TD class='MTD'><p class='Cd'>" & formata_moeda(vlTotalMeta) & "</p></td>" & chr(13) & _
					"		<TD class='MTD'><p class='Cd' style='color:" & DecodificaCorHtmlValorMonetario(vlTotalEfetuado) & ";'>" & formata_moeda(vlTotalEfetuado) & "</p></td>" & chr(13) & _
					"		<TD class='MTD'><p class='Cd' style='color:" & DecodificaCorHtmlValorMonetario(vlTotalSaldo) & ";'>" & formata_moeda(vlTotalSaldo) & "</p></td>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
					"		<TD class='MT' colspan='4'><p class='Cc'>" & formata_inteiro(intQtdeTotalIndicadores) & " indicador(es)</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='4'><P class='ALERTA'>&nbsp;NENHUM REGISTRO SATISFAZ AOS CRITÉRIOS&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO VENDEDOR
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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir( id_selecionado ) {
	fREL.action = "OrcamentistaEIndicadorConsulta.asp";
	fREL.id_selecionado.value = id_selecionado;
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
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="id_selecionado" id="id_selecionado" value=''>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<input type="hidden" name="c_dt_cadastro_mes" id="c_dt_cadastro_mes" value="<%=c_dt_cadastro_mes%>">
<input type="hidden" name="c_dt_cadastro_ano" id="c_dt_cadastro_ano" value="<%=c_dt_cadastro_ano%>">
<input type="hidden" name="c_dt_entregue_mes" id="c_dt_entregue_mes" value="<%=c_dt_entregue_mes%>">
<input type="hidden" name="c_dt_entregue_ano" id="c_dt_entregue_ano" value="<%=c_dt_entregue_ano%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="rb_periodo" id="rb_periodo" value="<%=rb_periodo%>">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Metas do Indicador</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
	if rb_periodo = COD_CONSULTA_POR_PERIODO_CADASTRO then
		s = iniciais_em_maiusculas(mes_por_extenso(c_dt_cadastro_mes, True)) & "/" & c_dt_cadastro_ano
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Período de Cadastro:&nbsp;</p></td><td valign='top' width='99%'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if

	if rb_periodo = COD_CONSULTA_POR_PERIODO_ENTREGA then
		s = iniciais_em_maiusculas(mes_por_extenso(c_dt_entregue_mes, True)) & "/" & c_dt_entregue_ano
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Período de Entrega:&nbsp;</p></td><td valign='top' width='99%'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if

	s = c_vendedor
	if s = "" then 
		s = "todos"
	elseif (Instr(1,s,"*")=0) And (Instr(1,s,BD_CURINGA_TODOS)=0) then
		s_aux = x_usuario(c_vendedor)
		if s_aux <> "" then s = s & " (" & s_aux & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Vendedor:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>" & chr(13)

	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & formata_data_hora(Now) & "</p></td></tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
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
