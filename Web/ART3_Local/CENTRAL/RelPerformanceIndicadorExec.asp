<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  R E L P E R F O R M A N C E I N D I C A D O R E X E C . A S P
'     =============================================================
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
	if Not operacao_permitida(OP_CEN_REL_PERFORMANCE_INDICADOR, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim c_mes_inicio, c_ano_inicio, c_mes_termino, c_ano_termino
	dim dtPeriodoInicial, dtPeriodoFinal
	dim c_vendedor, c_indicador
	dim s_nome_vendedor, s_nome_indicador, s_filtro, s, s_aux
	dim qtMeses

	alerta = ""
	
	c_mes_inicio = retorna_so_digitos(Request.Form("c_mes_inicio"))
	c_ano_inicio = retorna_so_digitos(Request.Form("c_ano_inicio"))
	c_mes_termino = retorna_so_digitos(Request.Form("c_mes_termino"))
	c_ano_termino = retorna_so_digitos(Request.Form("c_ano_termino"))

	c_indicador = Ucase(Trim(Request.Form("c_indicador")))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	
	if alerta = "" then
		s_nome_vendedor = ""
		if c_vendedor <> "" then
			s = "SELECT nome FROM t_USUARIO WHERE (usuario='" & c_vendedor & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "VENDEDOR " & c_vendedor & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_vendedor = Ucase(Trim("" & rs("nome")))
				end if
			end if
		end if
		
	if alerta = "" then
		s_nome_indicador = ""
		if c_indicador <> "" then
			s = "SELECT razao_social_nome FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido='" & c_indicador & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "INDICADOR " & c_indicador & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_indicador = Ucase(Trim("" & rs("razao_social_nome")))
				end if
			end if
		end if
	
	if alerta = "" then
		if c_mes_inicio = "" then
			alerta = "INFORME O MÊS DO PERÍODO INICIAL."
		elseif Not Isnumeric(c_mes_inicio) then
			alerta = "MÊS DO PERÍODO INICIAL NÃO É UM NÚMERO VÁLIDO."
		elseif (CInt(c_mes_inicio) <= 0) Or (CInt(c_mes_inicio) > 12) then
			alerta = "MÊS DO PERÍODO INICIAL É INVÁLIDO."
		elseif c_ano_inicio = "" then
			alerta = "INFORME O ANO DO PERÍODO INICIAL."
		elseif Not Isnumeric(c_ano_inicio) then
			alerta = "ANO DO PERÍODO INICIAL NÃO É UM NÚMERO VÁLIDO."
		elseif (CInt(c_ano_inicio) < 1900) Or (CInt(c_ano_inicio) > 2100) then
			alerta = "ANO DO PERÍODO INICIAL É INVÁLIDO."
		elseif c_mes_termino = "" then
			alerta = "INFORME O MÊS DO PERÍODO FINAL."
		elseif Not Isnumeric(c_mes_termino) then
			alerta = "MÊS DO PERÍODO FINAL NÃO É UM NÚMERO VÁLIDO."
		elseif (CInt(c_mes_termino) <= 0) Or (CInt(c_mes_termino) > 12) then
			alerta = "MÊS DO PERÍODO FINAL É INVÁLIDO."
		elseif c_ano_termino = "" then
			alerta = "INFORME O ANO DO PERÍODO FINAL."
		elseif Not Isnumeric(c_ano_termino) then
			alerta = "ANO DO PERÍODO FINAL NÃO É UM NÚMERO VÁLIDO."
		elseif (CInt(c_ano_termino) < 1900) Or (CInt(c_ano_termino) > 2100) then
			alerta = "ANO DO PERÍODO FINAL É INVÁLIDO."
			end if
		end if
	
	if alerta = "" then
		do while len(c_mes_inicio) < 2 : c_mes_inicio = "0" & c_mes_inicio : loop
		do while len(c_mes_termino) < 2: c_mes_termino = "0" & c_mes_termino : loop
		
		s = "01/" & c_mes_inicio & "/" & c_ano_inicio
		dtPeriodoInicial = StrToDate(s)
		s = "01/" & c_mes_termino & "/" & c_ano_termino
		dtPeriodoFinal = StrToDate(s)
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
		if alerta = "" then
			strDtRefDDMMYYYY = "01/" & c_mes_inicio & "/" & c_ano_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if
	
	'Calcula número de meses
	if alerta = "" then
		qtMeses = DateDiff("m", dtPeriodoInicial, dtPeriodoFinal) + 1
		end if





' _____________________________________________________________________________________________
'
'								D E C L A R A Ç Õ E S
' _____________________________________________________________________________________________
const ROW_VL_LISTA = 1
const ROW_VL_VENDA = 2

class cl_TOTAL
	dim vendedor
	dim vl_venda_ano_anterior
	dim vl_lista_ano_anterior
	dim vl_venda_mes(12)
	dim vl_lista_mes(12)
	dim vl_venda_ano_atual
	dim vl_lista_ano_atual
	dim perc
end class



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' TOTAL_INICIALIZA
'
sub Total_inicializa (byRef Total_vendedor)
dim i
	
	with Total_vendedor
		.vendedor = ""
		.vl_venda_ano_anterior = 0
		.vl_lista_ano_anterior = 0
		for i = Lbound(.vl_venda_mes) to Ubound(.vl_venda_mes)
			.vl_venda_mes(i) = 0
			.vl_lista_mes(i) = 0
			next
		.vl_venda_ano_atual = 0
		.vl_lista_ano_atual = 0
		.perc = 0
		end with

end sub


' _____________________________________
' TABELA_MONTA
'
function tabela_monta (byVal vRelat(), Total_vendedor)
dim i, intLargColIndicador, intLargColMonetario, intLargColPerc, intLargVendedor
dim cab_table, cab, cab_vendedor, linha_branco_peq, linha_branco_grande, x
dim s, s_mes, s_ano, s_cor, s_class
dim intIdxVetor
dim dtAux

'	ORDENA O VETOR COM RESULTADOS
	ordena_cl_vinte_colunas vRelat, 0, Ubound(vRelat)

'	CABEÇALHO
	intLargColIndicador = 90
	intLargColMonetario = 70
	intLargColPerc = 50
	intLargVendedor = 270

	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab_vendedor = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
					"		<TD class='MDTE' ColSpan='" & CStr(qtMeses + 4) & "' valign='bottom' NOWRAP><P class='R'>Vendedor: " & Total_vendedor.vendedor & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13)

	cab = cab & _
		"		<TD class='MT' valign='bottom' NOWRAP><P style='width:" & CStr(intLargColIndicador) & "px' class='R'>Indicador</P></TD>" & Chr(13) & _
		"		<TD class='MTBD' valign='bottom' NOWRAP><P style='width:" & CStr(intLargColMonetario) & "px' class='R'>Ano Anterior</P></TD>" & Chr(13)
	
	dtAux = dtPeriodoInicial
	for i = 1 to qtMeses
		s_mes = Month(dtAux)
		do while len(s_mes) < 2 : s_mes = "0" & s_mes : loop
		s_ano = Year(dtAux)
		cab = cab & "		<TD class='MTBD' valign='bottom' NOWRAP><P style='width:" & CStr(intLargColMonetario) & "px; text-align:center;' class='R' style='font-weight:bold'>" & s_mes & "/" & s_ano & "</P></TD>" & Chr(13)
		dtAux = DateAdd("m", 1, dtAux)
		next
	
	cab = cab & _
		"		<TD class='MTBD' valign='bottom' NOWRAP><P style='width:" & CStr(intLargColMonetario) & "px; text-align:center;' class='R'>Total</P></TD>" & Chr(13) & _
		"		<TD class='MTBD' valign='bottom' NOWRAP><P style='width:" & CStr(intLargColPerc) & "px; text-align:center;' class='R'>%</P></TD>" &  Chr(13) & _
		"	</TR>" & chr(13)

	linha_branco_peq = "	<TR NOWRAP>" & chr(13) & _
						"		<TD ColSpan='" & CStr(qtMeses + 4) & "'><P style='font-family: Arial, Helvetica, sans-serif;color:white;font-size:2pt;font-style:normal;'>&nbsp;</P></TD>" & chr(13) & _
						"	</TR>" & chr(13)

	linha_branco_grande = "	<TR NOWRAP>" & chr(13) & _
							"		<TD ColSpan='" & CStr(qtMeses + 4) & "'><P style='font-family: Arial, Helvetica, sans-serif;color:white;font-size:8pt;font-style:normal;'>&nbsp;</P></TD>" & chr(13) & _
							"	</TR>" & chr(13)

	x = cab_table & cab_vendedor & cab & linha_branco_peq
	
	intIdxVetor = Ubound(vRelat)
	do while (intIdxVetor >= 0)

		with vRelat(intIdxVetor)

			'Indicador
			s = ""
			if (.c17 = ROW_VL_VENDA) then
				s = Trim(.c13)
				x = x & "	<TR NOWRAP>" & chr(13)
				x = x & "		<TD class='MDTE' valign='bottom'><P class='Cn' style='width:" & CStr(intLargColIndicador) & "px;'>" & s & "</P></TD>" & chr(13)
				end if

			if (.c17 = ROW_VL_LISTA) then
				x = x & "	<TR NOWRAP>" & chr(13)
				x = x & "		<TD class='MTD' valign='bottom'><P class='Cn' style='width:" & CStr(intLargColIndicador) & "px;'>" & s & "</P></TD>" & chr(13)
				end if

			'Ano Anterior
			s = formata_moeda(.c14)
			if (.c17 = ROW_VL_LISTA) then
				if (.c14 <> 0 ) then
					s = formata_perc(100 * ((.c14-vRelat(intIdxVetor+1).c14) / .c14)) & "%"
					end if
				s_class = "MTBD"
				end if

			if (.c17 = ROW_VL_VENDA) then s_class = "MTD"

			s_cor = "black"
			if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"

			x = x & "		<TD class='" & s_class & "' valign='bottom' NOWRAP><P class='Cnd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s & "</P></TD>" & chr(13)

			s = ""
			for i = 1 to qtMeses
				select case i
					case 1 'Mês 1
						s = formata_moeda(.c1)
					case 2 'Mês 2
						s = formata_moeda(.c2)
					case 3 'Mês 3
						s = formata_moeda(.c3)
					case 4 'Mês 4
						s = formata_moeda(.c4)
					case 5 'Mês 5
						s = formata_moeda(.c5)
					case 6 'Mês 6
						s = formata_moeda(.c6)
					case 7 'Mês 7
						s = formata_moeda(.c7)
					case 8 'Mês 8
						s = formata_moeda(.c8)
					case 9 'Mês 9
						s = formata_moeda(.c9)
					case 10 'Mês 10
						s = formata_moeda(.c10)
					case 11 'Mês 11
						s = formata_moeda(.c11)
					case 12 'Mês 12
						s = formata_moeda(.c12)
					end select

				s_cor = "black"
				if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"

				if (.c17 = ROW_VL_LISTA) then
					if (s <> "") then s = s & "%"
					x = x & "		<TD class='MTBD' valign='bottom' NOWRAP><P class='Cnd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s &  "</P></TD>" & chr(13)
					end if

				if (.c17 = ROW_VL_VENDA) then
					x = x & "		<TD class='MTD' valign='bottom' NOWRAP><P class='Cnd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s &  "</P></TD>"  & chr(13)
					end if
				next

			'Total
			if (.c17 = ROW_VL_LISTA) then
				if (.c15 = 0) then
					s = ""
				else
					s = formata_perc(100 * (.c15-vRelat(intIdxVetor+1).c15) / .c15)
					s = s & "%"
					end if
				s_class = "MTBD"
				end if

			if (.c17 = ROW_VL_VENDA) then
				s = formata_moeda(.c15)
				s_class = "MTD"
				end if

			s_cor = "black"
			if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
			x = x & "		<TD class='" & s_class & "' valign='bottom' NOWRAP><P class='Cnd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s & "</P></TD>" & chr(13)

			' %
			s = ""
			if (.c17 = ROW_VL_VENDA) then
				s = formata_perc(.c16)
				if (s <> "") then s = s & "%"
				s_class = "MTD"
				end if

			if (.c17 = ROW_VL_LISTA) then
				s_class = "MC"
				end if

			s_cor = "black"
			if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
			x = x & "		<TD class='" & s_class & "' valign='bottom' NOWRAP><P class='Cnd' style='color:" & s_cor & ";width:" & CStr(intLargColPerc) & "px'>" & s & "</P></TD>" & chr(13)

			x = x & "	</TR>" & chr(13)
			if (.c17 = ROW_VL_LISTA) then
				x = x & linha_branco_peq
				end if
			end with
			
		intIdxVetor = intIdxVetor - 1
		loop


	' Totais
	x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13)
	with Total_vendedor
		' Indicador
		x = x & "		<TD class='MDTE' valign='bottom' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColIndicador) & "px;'> Total </P></TD>"	& chr(13)
		' Ano anterior
		s_cor = "black"
		if (len(.vl_venda_ano_anterior)>0) and (Mid(.vl_venda_ano_anterior,1,1) = "-") then
			s_cor = "red"
			end if

		x = x & "		<TD class='MTD' valign='bottom' NOWRAP><P class='Cd' style='width:" & CStr(intLargColMonetario) & "px'>" & formata_moeda(.vl_venda_ano_anterior) & "</P></TD>" & chr(13)

		' Meses
		s = ""
		for i = 1 to qtMeses
			s = formata_moeda(.vl_venda_mes(i))
			s_cor = "black"
			if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
			x = x & "		<TD class='MTD' valign='bottom' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s & "</P></TD>" & chr(13)
			next

		' Total (ano atual)
		s_cor = "black"
		if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
		x = x & "		<TD class='MTD' valign='bottom' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & formata_moeda(.vl_venda_ano_atual) & "</P></TD>" & chr(13)
		' %
		s = ""
		if (.vl_venda_ano_anterior <> 0) then
			.perc = ((.vl_venda_ano_atual- .vl_venda_ano_anterior )/.vl_venda_ano_anterior ) *100
			s = formata_perc(.perc) & "%"
			s_cor = "black"
			if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
			end if
		x = x & "		<TD class='MTD' valign='bottom' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColPerc) & "px'>" & s & "</P></TD>" & chr(13)

		x = x & "	</TR>" & chr(13)
		
	'	2ª LINHA DO TOTAL (PERCENTUAIS)
		x = x & "	<TR NOWRAP>" & chr(13)
		x = x & "		<TD class='MTD' valign='bottom'>&nbsp;</TD>" & chr(13)
		
		'Percentual total de desconto no ano anterior
		if .vl_lista_ano_anterior = 0 then
			s = ""
		else
			s = formata_perc(100 * (.vl_lista_ano_anterior - .vl_venda_ano_anterior)/.vl_lista_ano_anterior) & "%"
			end if
		s_cor = "black"
		if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
		x = x & "		<TD class='MTBD' valign='bottom' style='background: #FFFFDD' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s & "</P></TD>" & chr(13)
		
		'Percentual total de cada mês
		for i = 1 to qtMeses
			if .vl_lista_mes(i) = 0 then
				s = ""
			else
				s = formata_perc(100 * (.vl_lista_mes(i) - .vl_venda_mes(i))/.vl_lista_mes(i)) & "%"
				end if
			s_cor = "black"
			if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
			x = x & "		<TD class='MTBD' valign='bottom' style='background: #FFFFDD' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s & "</P></TD>" & chr(13)
			next
		
		'Percentual total do ano atual
		if .vl_lista_ano_atual = 0 then
			s = ""
		else
			s = formata_perc(100 * (.vl_lista_ano_atual - .vl_venda_ano_atual)/.vl_lista_ano_atual) & "%"
			end if
		s_cor = "black"
		if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
		x = x & "		<TD class='MTBD' valign='bottom' style='background: #FFFFDD' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s & "</P></TD>" & chr(13)
		
		x = x & "		<TD class='MC' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13)
		end with


	x = x & "	</TR>" & chr(13)
	
'	PULA LINHA
	x = x & linha_branco_grande & linha_branco_grande
	
	tabela_monta = x

end function


' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa

dim r
dim s_mes, s_ano
dim s_where, s_where_venda, s_where_devolucao, s_where_loja
dim s_sql, n_reg_BD, x
dim strSqlCampoSaida
dim perc_desconto
dim vRelat()
dim Total_geral, Total_vendedor
dim strIndicador, strIndicadorAtual, strIndicadorAnterior
dim strVendedorAtual, strVendedorAnterior, s
dim vl_venda, vl_lista, vl_total_indicador, vl_lista_indicador
dim intIdxVetor
dim i_mes, n_reg
dim i, intLargColIndicador, intLargColMonetario, intLargColPerc, intLargVendedor
dim s_cor
dim dtAux, dtInicio, dtTermino, dtInicioAnoAnterior, dtTerminoAnoAnterior

'	SELECTs INTERNOS

'	CRITÉRIOS COMUNS
	s_where = ""
	
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.vendedor = '" & c_vendedor & "')"
		end if
	
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if


'	CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	s_where_venda = ""
	
'	PERÍODO
	dtInicio = dtPeriodoInicial
	dtInicioAnoAnterior = DateAdd("yyyy", -1, dtPeriodoInicial)
	dtTermino = DateAdd("m", 1, dtPeriodoFinal)
	dtTerminoAnoAnterior = DateAdd("yyyy", -1, dtTermino)
	
	if s_where_venda <> "" then s_where_venda = s_where_venda & " AND "
	s_where_venda = s_where_venda & _
					"(" & _
						"(" & _
							"(t_PEDIDO.entregue_data >= " & bd_formata_data(dtInicio) & ")" & _
							" AND " & _
							"(t_PEDIDO.entregue_data < " & bd_formata_data(dtTermino) & ")" & _
						")" & _
						" OR " & _
						"(" & _
							"(t_PEDIDO.entregue_data >= " & bd_formata_data(dtInicioAnoAnterior) & ")" & _
								" AND " & _
							"(t_PEDIDO.entregue_data < " & bd_formata_data(dtTerminoAnoAnterior) & ")" & _
						")" & _
					")"
	
'	CRITÉRIOS PARA DEVOLUÇÕES
	s_where_devolucao = ""
	
'	PERÍODO
	if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND "
	s_where_devolucao = s_where_devolucao & _
						"(" & _
							"(" & _
								"(t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(dtInicio) & ")" & _
								" AND " & _
								"(t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(dtTermino) & ")" & _
							")" & _
							" OR " & _
							"(" & _
								"(t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(dtInicioAnoAnterior) & ")" & _
								" AND " & _
								"(t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(dtTerminoAnoAnterior) & ")" & _
							")" & _
						")"

'	MONTA SQL DE CONSULTA
	
'	A) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
'	B) O CAMPO 'QTDE' A SER USADO DEVE SER DA TABELA T_ESTOQUE_MOVIMENTO, JÁ
'	QUE UM PEDIDO PODE TER USADO DIVERSOS LOTES DO ESTOQUE PARA ATENDER A
'	UM ÚNICO PRODUTO.  NESSE CASO, HAVERÁ MAIS DE UM REGISTRO EM
'	T_ESTOQUE_MOVIMENTO SE RELACIONANDO COM O MESMO REGISTRO DE T_PEDIDO_ITEM.
'	A SOMA DE 'QTDE' DOS REGISTROS DE T_ESTOQUE_MOVIMENTO RESULTAM NO VALOR
'	DE 'QTDE' DE T_PEDIDO_ITEM.
	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then s = " AND" & s
	
	s_sql = "SELECT " & _
				"t_PEDIDO.vendedor as vendedor, " & _
				" t_ORCAMENTISTA_E_INDICADOR.apelido as Indicador, " & _
				" YEAR(t_PEDIDO.entregue_data) as ano, " & _
				" MONTH(t_PEDIDO.entregue_data) as mes, " & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_venda) AS valor_venda, " & _
				" Sum(t_ESTOQUE_MOVIMENTO.qtde*t_PEDIDO_ITEM.preco_lista) AS preco_lista " & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" INNER JOIN t_PEDIDO" & _
					" ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
					" ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" INNER JOIN t_PEDIDO_ITEM" & _
					" ON ((t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO_ITEM.pedido)AND(t_ESTOQUE_MOVIMENTO.fabricante=t_PEDIDO_ITEM.fabricante)AND(t_ESTOQUE_MOVIMENTO.produto=t_PEDIDO_ITEM.produto))" & _
				" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR" & _
					" ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido) " & _
			" WHERE" & _
				" (anulado_status=0)" & _
				" AND (estoque='" & ID_ESTOQUE_ENTREGUE & "')" & _
				s & _
			" GROUP BY " & _
				" t_PEDIDO.vendedor," & _
				" t_ORCAMENTISTA_E_INDICADOR.apelido," & _
				" YEAR(t_PEDIDO.entregue_data)," & _
				" MONTH(t_PEDIDO.entregue_data)"
	
'	LEMBRE-SE: NA DEVOLUÇÃO DO PRODUTO, É CRIADA UMA ENTRADA NO ESTOQUE DE VENDA P/
'	REPRESENTAR A ENTRADA DA MERCADORIA NO ESTOQUE. ENTRETANTO, A QUANTIDADE
'	DEVOLVIDA FICA INICIALMENTE TODA ALOCADA P/ O ESTOQUE DE DEVOLUÇÃO, DEVIDO
'	À NECESSIDADE DE TRATAR A MERCADORIA ANTES DE DISPONIBILIZA-LA P/ VENDA.
'	IMPORTANTE: NO CASO DE OCORRER A DEVOLUÇÃO DE VÁRIAS UNIDADES, PODEM SER
'	CRIADOS VÁRIOS REGISTROS DE ESTOQUE A DIFERENTES CUSTOS DE AQUISIÇÃO.
	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & _
			" UNION " & _
			"SELECT " & _
				"t_PEDIDO.vendedor as vendedor, " & _
				"t_ORCAMENTISTA_E_INDICADOR.apelido as Indicador, " & _
				"YEAR(t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data) as ano, " & _
				"MONTH(t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data) as mes, " & _
				"Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS valor_venda, " & _
				"Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_lista) AS preco_lista " & _
			"FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
				" INNER JOIN t_PEDIDO" & _
					" ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
					" ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR" & _
					" ON (t_PEDIDO.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
				s & _
			" GROUP BY " & _
				" t_PEDIDO.vendedor," & _
				" t_ORCAMENTISTA_E_INDICADOR.apelido," & _
				" YEAR(t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data)," & _
				" MONTH(t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data) "

'	SELECT EXTERNO
	strsqlCampoSaida = "vendedor, indicador, ano,  mes "
	s_sql = "SELECT vendedor, indicador, ano,  mes, " & _
				" SUM(valor_venda) as valor_venda, SUM(preco_lista) as preco_lista " & _
			"FROM " & _
				"(" & _
					s_sql & _
				") as Base " & _
			" GROUP BY " & strsqlCampoSaida & _
			" ORDER BY " & strsqlCampoSaida


'	Inicializa Variáveis
	set Total_geral = New cl_TOTAL
	Total_inicializa(Total_geral)
	set Total_vendedor = New cl_TOTAL
	Total_inicializa(Total_vendedor)
	
'	Descrição dos Dados em vRelat
'	Os dados de cada Indicador ocupam 2 linhas em vRelat, sendo uma com os
'	valores de venda (.c17 = 2), correspondente à primeira linha no relatório
'	e a outra com os valores de lista (.c17 = 1), correspondente à segunda linha no relatório
'	Os dados em vRelat são:
'	.c1 a .c12: valor do mês correspondente
'	.c13: Nome do Indicador
'	.c14: Valores acumulados do ano anterior
'	.c15: Valor acumulado por Indicador
'	.c16: Somente p/ .c17 = 2 (valor de venda): Variação em relação ao ano anterior (em porcentagem)
'	.c17: 1 para valor de lista
'		  2 para valor de venda
	redim vRelat(1)
	for intIdxVetor = Lbound(vRelat) to Ubound(vRelat)
		set vRelat(intIdxVetor) = New cl_VINTE_COLUNAS
		vRelat(intIdxVetor).CampoOrdenacao = ""
		if (intIdxVetor = 0) then
			vRelat(intIdxVetor).c17 = ROW_VL_LISTA
			end if
		if (intIdxVetor = 1) then
			vRelat(intIdxVetor).c17 = ROW_VL_VENDA
			end if
		next


'	EXECUTA CONSULTA SQL
	set r = cn.execute(s_sql)
	
'	PERCORRE OS REGISTROS OBTIDOS DO BD E CONSOLIDA OS VALORES
'	EM UMA ÚNICA LINHA POR VENDEDOR E INDICADOR
	strVendedorAnterior  = "-VV-VV-VV-VV-VV-VV"
	strIndicadorAnterior = "-II-II-II-II-II-II"
	vl_total_indicador = 0
	vl_lista_indicador = 0
	n_reg = 0
	do while Not r.Eof
		n_reg = n_reg + 1
	'	MUDOU DE VENDEDOR?
		strVendedorAtual = ""
		if not IsNull(r("vendedor")) then strVendedorAtual = r("vendedor")
		
		strIndicadorAtual = ""
		if not IsNull(r("indicador")) then strIndicadorAtual = r("indicador")

		if strVendedorAtual <> strVendedorAnterior then
			if (strVendedorAnterior <> "-VV-VV-VV-VV-VV-VV") then
				'vl_lista
				with  vRelat(Ubound(vRelat)-1)
					.CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_total_indicador)), 20) & _
									  strIndicador & .c17
					.c15 = vl_lista_indicador
					end with

				'vl_venda
				with  vRelat(Ubound(vRelat))
					.CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_total_indicador)), 20) & _
									  strIndicador & .c17
					.c15 = vl_total_indicador
					if (.c14 <> 0) then
						.c16 = ((.c15 - .c14)/.c14 ) *100
						end if
					end with

				'Totalização Geral
				with Total_geral
					.vl_venda_ano_anterior = .vl_venda_ano_anterior + Total_vendedor.vl_venda_ano_anterior
					.vl_lista_ano_anterior = .vl_lista_ano_anterior + Total_vendedor.vl_lista_ano_anterior

					for i = LBOUND(.vl_venda_mes) to UBOUND(.vl_venda_mes)
						.vl_venda_mes(i) = .vl_venda_mes(i) + Total_vendedor.vl_venda_mes(i)
						.vl_lista_mes(i) = .vl_lista_mes(i) + Total_vendedor.vl_lista_mes(i)
						next
					.vl_venda_ano_atual = .vl_venda_ano_atual + Total_vendedor.vl_venda_ano_atual
					.vl_lista_ano_atual = .vl_lista_ano_atual + Total_vendedor.vl_lista_ano_atual
					end with

				x = tabela_monta(vRelat, Total_vendedor)
				Response.Write x
				x = ""

				Total_inicializa(Total_vendedor)
				
				'Inicializa vetor
				redim vRelat(1)
				for intIdxVetor = Lbound(vRelat) to Ubound(vRelat)
					set vRelat(intIdxVetor) = New cl_VINTE_COLUNAS
					vRelat(intIdxVetor).CampoOrdenacao = ""
					if (intIdxVetor = 0) then
						vRelat(intIdxVetor).c17 = ROW_VL_LISTA
						end if
					if (intIdxVetor = 1) then
						vRelat(intIdxVetor).c17 = ROW_VL_VENDA
						end if
					next
				end if

				strVendedorAnterior = strVendedorAtual
				strIndicadorAnterior = strIndicadorAtual
				vl_total_indicador = 0
				vl_lista_indicador = 0
				end if

	'	MUDOU DE INDICADOR?
		if strIndicadorAtual <> strIndicadorAnterior then
			if (strIndicadorAnterior <> "-II-II-II-II-II-II") then
				'vl_lista
				with  vRelat(Ubound(vRelat)-1)
					.CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_total_indicador)), 20) & _
									  strIndicador & .c17
					.c15 = vl_lista_indicador
					end with

				'vl_venda
				with  vRelat(Ubound(vRelat))
					.CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_total_indicador)), 20) & _
									  strIndicador & .c17
					.c15 = vl_total_indicador
					if (.c14 <> 0) then
						.c16 = ((.c15 - .c14)/.c14 ) * 100
						end if
					.c14 = formata_moeda(.c14)
					end with

				redim preserve vRelat(Ubound(vRelat)+2)
				set vRelat(Ubound(vRelat)-1) = New cl_VINTE_COLUNAS
				vRelat(Ubound(vRelat)-1).CampoOrdenacao = ""
				vRelat(Ubound(vRelat)-1).c17 = ROW_VL_LISTA
				set vRelat(Ubound(vRelat)) = New cl_VINTE_COLUNAS
				vRelat(Ubound(vRelat)).CampoOrdenacao = ""
				vRelat(Ubound(vRelat)).c17 = ROW_VL_VENDA
				end if
			
			strIndicadorAnterior = strIndicadorAtual
			vl_total_indicador = 0
			vl_lista_indicador = 0
			end if

	'>	VENDEDOR
		Total_vendedor.vendedor = r("vendedor")


	'>	INDICADOR
		strIndicador = " "
		if not IsNull(r("indicador")) then strIndicador = r("indicador")

		vRelat(Ubound(vRelat)-1).c13 = strIndicador
		vRelat(Ubound(vRelat)).c13 = strIndicador

	'>	MÊS
		s_mes = Trim(r("mes"))

	'>	ANO
		s_ano = Trim(r("ano"))

	'>	VALOR DE VENDA
		vl_venda = r("valor_venda")

	'>	VALOR DE LISTA
		vl_lista = r("preco_lista")

		dtAux = StrToDate ("01/" & s_mes & "/" & s_ano)
	'	ANO ANTERIOR
		if (dtAux < dtPeriodoInicial) then
			vRelat(Ubound(vRelat)).c14 = vRelat(Ubound(vRelat)).c14 + vl_venda
			Total_vendedor.vl_venda_ano_anterior = Total_vendedor.vl_venda_ano_anterior + vl_venda
			vRelat(Ubound(vRelat)-1).c14 = vRelat(Ubound(vRelat)-1).c14 + vl_lista
			Total_vendedor.vl_lista_ano_anterior = Total_vendedor.vl_lista_ano_anterior + vl_lista
	'	ANO ATUAL
		else
		'	% DESCONTO
			if vl_lista = 0 then
				perc_desconto = 0
			else
				perc_desconto = 100 * (vl_lista-vl_venda) / vl_lista
				end if
			i_mes = DateDiff("m", dtPeriodoInicial, dtAux) + 1
			Select Case i_mes
				Case 1 ' Mês 1
					vRelat(Ubound(vRelat)).c1 = vl_venda
					vRelat(Ubound(vRelat)-1).c1 = perc_desconto
				Case 2 ' Mês 2
					vRelat(Ubound(vRelat)).c2 = vl_venda
					vRelat(Ubound(vRelat)-1).c2 = perc_desconto
				Case 3 ' Mês 3
					vRelat(Ubound(vRelat)).c3 = vl_venda
					vRelat(Ubound(vRelat)-1).c3 = perc_desconto
				Case 4 ' Mês 4
					vRelat(Ubound(vRelat)).c4 = vl_venda
					vRelat(Ubound(vRelat)-1).c4 = perc_desconto
				Case 5 ' Mês 5
					vRelat(Ubound(vRelat)).c5 = vl_venda
					vRelat(Ubound(vRelat)-1).c5 = perc_desconto
				Case 6 ' Mês 6
					vRelat(Ubound(vRelat)).c6 = vl_venda
					vRelat(Ubound(vRelat)-1).c6 = perc_desconto
				Case 7 ' Mês 7
					vRelat(Ubound(vRelat)).c7 = vl_venda
					vRelat(Ubound(vRelat)-1).c7 = perc_desconto
				Case 8 ' Mês 8
					vRelat(Ubound(vRelat)).c8 = vl_venda
					vRelat(Ubound(vRelat)-1).c8 = perc_desconto
				Case 9 ' Mês 9
					vRelat(Ubound(vRelat)).c9 = vl_venda
					vRelat(Ubound(vRelat)-1).c9 = perc_desconto
				Case 10 ' Mês 10
					vRelat(Ubound(vRelat)).c10 = vl_venda
					vRelat(Ubound(vRelat)-1).c10 = perc_desconto
				Case 11 ' Mês 11
					vRelat(Ubound(vRelat)).c11 = vl_venda
					vRelat(Ubound(vRelat)-1).c11 = perc_desconto
				Case 12 ' Mês 12
					vRelat(Ubound(vRelat)).c12 = vl_venda
					vRelat(Ubound(vRelat)-1).c12 = perc_desconto
				End Select

			Total_vendedor.vl_venda_mes(i_mes) = Total_vendedor.vl_venda_mes(i_mes) + vl_venda
			Total_vendedor.vl_lista_mes(i_mes) = Total_vendedor.vl_lista_mes(i_mes) + vl_lista
			Total_vendedor.vl_venda_ano_atual = Total_vendedor.vl_venda_ano_atual + vl_venda
			Total_vendedor.vl_lista_ano_atual = Total_vendedor.vl_lista_ano_atual + vl_lista
			vl_total_indicador = vl_total_indicador + vl_venda
			vl_lista_indicador = vl_lista_indicador + vl_lista
			end if

		r.MoveNext
		loop

'	MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = "	<TABLE><TR NOWRAP>" & chr(13) & _
			"		<TD class='MT' ><P class='ALERTA'>&nbsp;NENHUM REGISTRO SATISFAZ AOS CRITÉRIOS&nbsp;</P></TD>" & chr(13) & _
			"	</TR></TABLE>" & chr(13)
		Response.Write x
	else
		'vl_lista
		with  vRelat(Ubound(vRelat)-1)
			.CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_total_indicador)), 20) & _
							  strIndicador & .c17
			.c15 = vl_lista_indicador
			end with

		'vl_venda
		with vRelat(Ubound(vRelat))
			.CampoOrdenacao = normaliza_codigo(retorna_so_digitos(formata_moeda(vl_total_indicador)), 20) & _
							  strIndicador & .c17
			.c15 = vl_total_indicador
			if (.c14 <> 0) then
				.c16 = ((.c15 - .c14)/.c14 ) * 100
				end if
			end with

		x = tabela_monta (vRelat, Total_vendedor)
		Response.Write x
		x = ""

		'Totalização Geral
		with Total_geral
			.vl_venda_ano_anterior = .vl_venda_ano_anterior + Total_vendedor.vl_venda_ano_anterior
			.vl_lista_ano_anterior = .vl_lista_ano_anterior + Total_vendedor.vl_lista_ano_anterior
			for i = Lbound(.vl_venda_mes) to Ubound(.vl_venda_mes)
				.vl_venda_mes(i) = .vl_venda_mes(i) + Total_vendedor.vl_venda_mes(i)
				.vl_lista_mes(i) = .vl_lista_mes(i) + Total_vendedor.vl_lista_mes(i)
				next
			.vl_venda_ano_atual = .vl_venda_ano_atual + Total_vendedor.vl_venda_ano_atual
			.vl_lista_ano_atual = .vl_lista_ano_atual + Total_vendedor.vl_lista_ano_atual
			end with

		intLargColIndicador = 120
		intLargColMonetario = 70
		intLargColPerc = 50
		intLargVendedor = 270

		x = x & "	<TR NOWRAP style='background:aquamarine'>" & chr(13)
		with Total_geral
			' Indicador
			x = x & "		<TD class='MDTE' valign='bottom'><P class='Cd' style='width:" & CStr(intLargColIndicador) & "px;'> Total Geral</P></TD>" & chr(13)
			' Ano anterior
			s_cor = "black"
			if (len(.vl_venda_ano_anterior)>0) and (Mid(.vl_venda_ano_anterior,1,1) = "-") then s_cor = "red"
			x = x & "		<TD class='MTD' valign='bottom' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & formata_moeda(.vl_venda_ano_anterior) & "</P></TD>" & chr(13)

			' Meses
			s = ""
			for i = 1 to qtMeses
				s = formata_moeda(.vl_venda_mes(i))
				s_cor = "black"
				if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
				x = x & "		<TD class='MTD' valign='bottom' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s & "</P></TD>" & chr(13)
				next

			' Total (ano atual)
			x = x & "		<TD class='MTD' valign='bottom' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & formata_moeda(.vl_venda_ano_atual) & "</P></TD>"  & chr(13)
			
			' %
			s = ""
			if (.vl_venda_ano_anterior <> 0) then
				.perc = ((.vl_venda_ano_atual- .vl_venda_ano_anterior )/.vl_venda_ano_anterior ) * 100
				s = formata_perc(.perc) & "%"
				s_cor = "black"
				if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
				end if
			x = x & "		<TD class='MTD' valign='bottom' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColPerc) & "px'>" & s & "</P></TD>"  & chr(13)

			x = x & "	</TR>" & chr(13)

		'	2ª LINHA DO TOTAL (PERCENTUAIS)
			x = x & "	<TR NOWRAP>" & chr(13)
			x = x & "		<TD class='MTD' valign='bottom'>&nbsp;</TD>" & chr(13)
			
			'Percentual total de desconto no ano anterior
			if .vl_lista_ano_anterior = 0 then
				s = ""
			else
				s = formata_perc(100 * (.vl_lista_ano_anterior - .vl_venda_ano_anterior)/.vl_lista_ano_anterior) & "%"
				end if
			s_cor = "black"
			if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
			x = x & "		<TD class='MTBD' valign='bottom' style='background:aquamarine' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s & "</P></TD>" & chr(13)
			
			'Percentual total de cada mês
			for i = 1 to qtMeses
				if .vl_lista_mes(i) = 0 then
					s = ""
				else
					s = formata_perc(100 * (.vl_lista_mes(i) - .vl_venda_mes(i))/.vl_lista_mes(i)) & "%"
					end if
				s_cor = "black"
				if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
				x = x & "		<TD class='MTBD' valign='bottom' style='background:aquamarine' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s & "</P></TD>" & chr(13)
				next
			
			'Percentual total do ano atual
			if .vl_lista_ano_atual = 0 then
				s = ""
			else
				s = formata_perc(100 * (.vl_lista_ano_atual - .vl_venda_ano_atual)/.vl_lista_ano_atual) & "%"
				end if
			s_cor = "black"
			if (len(s)>0) and (Mid(s,1,1) = "-") then s_cor = "red"
			x = x & "		<TD class='MTBD' valign='bottom' style='background:aquamarine' NOWRAP><P class='Cd' style='color:" & s_cor & ";width:" & CStr(intLargColMonetario) & "px'>" & s & "</P></TD>" & chr(13)
			
			x = x & "		<TD class='MC' valign='bottom' NOWRAP>&nbsp;</TD>" & chr(13)
			end with

		x = x & "	</TR>" & chr(13)


	'	FECHA TABELA
		x = x & "</TABLE>" & chr(13)

		Response.Write x

		end if

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
<body onload="window.status='Concluído';bVOLTAR.focus();">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_mes_inicio" id="c_mes_inicio" value="<%=c_mes_inicio%>">
<input type="hidden" name="c_mes_termino" id="c_mes_termino" value="<%=c_mes_termino%>">
<input type="hidden" name="c_ano_inicio" id="c_ano_inicio" value="<%=c_ano_inicio%>">
<input type="hidden" name="c_ano_termino" id="c_ano_termino" value="<%=c_ano_termino%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A    -->
<table width="947" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Performance por Indicador</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='947' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>"
	
	s = ""
	s_aux = c_mes_inicio & "/" & c_ano_inicio
	s = s & s_aux & " a "
	s_aux = c_mes_termino & "/" & c_ano_termino
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Entregues entre:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s = c_vendedor
	if s = "" then
		s = "todos"
	else
		if (s_nome_vendedor <> "") And (s_nome_vendedor <> c_vendedor) then s = s & " (" & s_nome_vendedor & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Vendedor:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s = c_indicador
	if s = "" then
		s = "todos"
	else
		if (s_nome_indicador <> "") And (s_nome_indicador <> c_indicador) then s = s & " (" & s_nome_indicador & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Indicador:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & formata_data_hora(Now) & "</p></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="947" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
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
