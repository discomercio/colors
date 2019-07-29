<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelAuditoriaEstoqueExec.asp
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

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
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
	if Not operacao_permitida(OP_CEN_REL_AUDITORIA_ESTOQUE, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux
	dim c_dt_inicio, c_dt_termino, c_fabricante, c_produto
	dim flag_ok, s_filtro

	alerta = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))	

	if alerta = "" then
		if c_fabricante <> "" then
			s = "SELECT fabricante FROM t_FABRICANTE WHERE (fabricante='" & c_fabricante & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "FABRICANTE " & c_fabricante & " NÃO ESTÁ CADASTRADO."
				end if
			end if
		end if
		
	if alerta = "" then
		if c_produto <> "" then
			if (Not IsEAN(c_produto)) And (c_fabricante="") then
				alerta=texto_add_br(alerta)
				alerta=alerta & "NÃO FOI ESPECIFICADO O FABRICANTE DO PRODUTO A SER CONSULTADO."
			else
				s = "SELECT * FROM t_PRODUTO WHERE"
				if IsEAN(c_produto) then
					s = s & " (ean='" & c_produto & "')"
				else
					s = s & " (fabricante='" & c_fabricante & "') AND (produto='" & c_produto & "')"
					end if
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if Not rs.Eof then
					flag_ok = True
					if IsEAN(c_produto) And (c_fabricante<>"") then
						if (c_fabricante<>Trim("" & rs("fabricante"))) then
							flag_ok = False
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto a ser consultado " & c_produto & " NÃO pertence ao fabricante " & c_fabricante & "."
							end if
						end if
					if flag_ok then
					'	CARREGA CÓDIGO INTERNO DO PRODUTO
						c_fabricante = Trim("" & rs("fabricante"))
						c_produto = Trim("" & rs("produto"))
						end if
					end if
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
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_termino
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			if c_dt_inicio = "" then c_dt_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
			end if
		
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if
	
	
	
' _____________________________________________________________________________________________
'
'								C O N S T A N T E S
' _____________________________________________________________________________________________
	
const TP_OPERACAO = "Operação"
const TP_SALDO = "Saldo"
const TP_MENSAGEM = "Mensagem"


' _____________________________________________________________________________________________
'
'								D E C L A R A Ç Õ E S
' _____________________________________________________________________________________________

class cl_ESTOQUE_LOG
	dim tipo 'TP_OPERACAO, TP_SALDO, TP_MENSAGEM
	dim fabricante
	dim produto
	dim operacao
	dim estoque_origem
	dim estoque_destino
	dim qtde
	dim saldo
end class







' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' SALDO_OK
'
function saldo_OK (byVal data, byRef msg_erro)
dim r
dim s_sql

	saldo_OK = false

	If Not cria_recordset_otimista(r, msg_erro) then
		exit function
		end if

	s_sql = "SELECT " & _
				"*" & _
			" FROM t_ESTOQUE_SALDO_DIARIO" & _
			" WHERE" & _
				" (data = " & bd_formata_data(StrToDate(data)) &  ")" & _
				" AND (fabricante = '----') AND (produto = '--------') AND (estoque = '---')" 
	r.open s_sql, cn
	if Not r.Eof then saldo_OK = true
	if r.State <> 0 then r.Close

end function



' _____________________________________
' PRODUTO_FINALIZA
'

sub prooduto_finaliza (byRef vEstoqueLog(), byVal vSaldoFinal, produto, SaldoFinalOk)
dim indice_localizado
dim SaldoLocalizado
dim SaldoFinalBD

	redim preserve vEstoqueLog(Ubound(vEstoqueLog)+1)
	set vEstoqueLog(Ubound(vEstoqueLog)) = New cl_ESTOQUE_LOG
	with vEstoqueLog(Ubound(vEstoqueLog))
		.tipo = TP_SALDO
		.produto = produto
		.operacao = "Saldo final"
		.estoque_origem = " "
		.estoque_destino = " "
		.qtde = " "
		.saldo = vEstoqueLog(Ubound(vEstoqueLog)-1).saldo
		end with
	
	if (SaldoFinalOk) then
		SaldoFinalBD = 0
		if  localiza_cl_tres_colunas(vSaldoFinal, produto, indice_localizado) then
			SaldoFinalBD = vSaldoFinal(indice_localizado).c2
			end if
	
		if (vEstoqueLog(Ubound(vEstoqueLog)).saldo <> SaldoFinalBD) then
			redim preserve vEstoqueLog(Ubound(vEstoqueLog)+2)
			set vEstoqueLog(Ubound(vEstoqueLog)-1) = New cl_ESTOQUE_LOG
			with vEstoqueLog(Ubound(vEstoqueLog)-1)
				.produto = produto
				.tipo = TP_MENSAGEM
				.Operacao = "ATENÇÃO: divergência entre o saldo calculado pelo relatório e o saldo consolidado armazenado no sistema."
				end with
			set vEstoqueLog(Ubound(vEstoqueLog)) = New cl_ESTOQUE_LOG
			with vEstoqueLog(Ubound(vEstoqueLog))
				.produto = produto
				.tipo = TP_SALDO
				.Operacao = "Saldo consolidado no sistema" 
				.Saldo = SaldoFinalBD
				end with
			end if
	else
		redim preserve vEstoqueLog(Ubound(vEstoqueLog)+1)
		set vEstoqueLog(Ubound(vEstoqueLog)) = New cl_ESTOQUE_LOG
		with vEstoqueLog(Ubound(vEstoqueLog))
			.produto = produto
			.tipo = TP_MENSAGEM
			.Operacao = "Resultado do saldo final não confirmado por falta de dados consolidados no sistema em " & c_dt_termino
			end with
		end if
	
end sub


' _____________________________________
' TABELA_MONTA
'
function tabela_monta (byVal vEstoqueLog())
dim cab
dim x
dim linha_em_branco
dim s_produto
dim intIdxVetor
dim s
dim s_aux


  ' PRODUTO
	s = vEstoqueLog(Lbound(vEstoqueLog)).produto
	s_aux = produto_formata_descricao_em_html(produto_descricao_html(c_fabricante, s))
	if (s<>"") And (s_aux<>"") then s = trim(s) & " - "
	s = s & Trim(s_aux)
	s_produto = 	"	<TR NOWRAP>" & chr(13) & _
					"		<TD ColSpan='5' align='left' valign='top' NOWRAP>" & _
					"<p class='Np'>Produto: " & s & "</p></td>" & Chr(13) & _
					"	</TR>" & chr(13)
	x = s_produto

  ' CABEÇALHO
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) 
	cab = cab & _
		  "		<TD class='MTBE tdOp' style='vertical-align:bottom'><P class='R'>Operação</P></TD>" & chr(13) & _
		  "		<TD class='MTBE tdEstOrigem' style='vertical-align:bottom'><P class='R'>Estoque Origem</P></TD>" & chr(13) & _
		  "		<TD class='MTBE tdEstDestino' style='vertical-align:bottom'><P class='R'>Estoque Destino</P></TD>" & chr(13) & _
		  "		<TD class='MTBE tdQtd' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Qtde</P></TD>" & chr(13) & _
		  "		<TD class='MT tdSaldo' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Saldo</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = x & cab
	
	for intIdxVetor=Lbound(vEstoqueLog) to Ubound(vEstoqueLog) 
	
		x = x & "	<TR NOWRAP>" & chr(13)

		with vEstoqueLog(intIdxVetor)
			
			'> OPERAÇÃO
			if ( .tipo = TP_OPERACAO ) then
				s = substitui_caracteres(x_operacao_log_estoque(Trim("" & .operacao)), " ", "&nbsp;")
			else
				s = Trim("" & .operacao)
				end if
			if s = "" then s = "&nbsp;"
			
			select case (.tipo)
				case TP_MENSAGEM
					x = x & "		<TD class='MDBE' ColSpan='5'><P class='Cn'>" & s & "</P></TD>" & chr(13)
				case TP_OPERACAO
					x = x & "		<TD class='MEB tdOp'><P class='Cn'>" & s & "</P></TD>" & chr(13)
				case TP_SALDO
					x = x & "		<TD class='MEB' ColSpan='4'><P class='Cn'>" & s & "</P></TD>" & chr(13)
				end select
				
			'> ESTOQUE ORIGEM
			if ( .tipo = TP_OPERACAO ) then
				s = x_estoque(Trim("" & .estoque_origem))
				if s = "" then s = "&nbsp;"
				x = x & "		<TD class='MEB tdEstOrigem'><P class='Cn'>" & s & "</P></TD>" & chr(13)
				end if

			'> ESTOQUE DESTINO
			if ( .tipo = TP_OPERACAO ) then
				s = x_estoque(Trim("" & .estoque_destino))
				if s = "" then s = "&nbsp;"
				x = x & "		<TD class='MEB tdEstDestino'><P class='Cn'>" & s & "</P></TD>" & chr(13)
				end if

			'> QUANTIDADE
			if ( .tipo = TP_OPERACAO )  then
				s = formata_inteiro(.qtde)
				if (.qtde) > 0 then s = "+" & s
				if s = "" then s = "&nbsp;"
				x = x & "		<TD class='MEB tdQtd'><P class='Cd'>" & s & "</P></TD>" & chr(13)
				end if

			'>	SALDO
			if ( .tipo = TP_OPERACAO ) or ( .tipo = TP_SALDO ) then
				s = retorna_so_digitos(.saldo)
				if ( s= "" ) then
					s = .saldo
				else s = formata_inteiro(.saldo)
					end if
				x = x & "		<TD class='MDBE tdSaldo'><P class='Cd'>" & s & "</P></TD>" & chr(13)
				end if
			
			end with

			x = x & "	</TR>" & chr(13)
		
		next

	linha_em_branco = "	<TR NOWRAP>" & chr(13) & _
					  "		<TD ColSpan='5'><P style='font-family: Arial, Helvetica, sans-serif;color:white;font-size:8pt;font-style:normal;'>&nbsp;</P></TD>" & chr(13) & _
					  "	</TR>" & chr(13)
	
	tabela_monta = x & linha_em_branco

end function



' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa

dim r
dim n_reg
dim x, s, s2, s_estoque, s_sql, msg_erro
dim s_where 
dim vSaldoInicial(), vSaldoFinal(), vEstoqueLog()
dim SaldoInicialOk, SaldoFinalOk
dim strProdutoAnterior, strProdutoAtual
dim indice_localizado
dim entrada_estoque, saida_estoque


'	CRITÉRIOS COMUNS
	s_where = ""

'	FABRICANTE
	if c_fabricante <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (fabricante = '" & c_fabricante & "')"
		end if

'	PRODUTO
	if c_produto <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (produto = '" & c_produto & "')"
		end if

'	MONTA A CONSULTA DO SALDO NO DIA ANTERIOR
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	msg_erro = ""
	SaldoInicialOk = saldo_OK (StrToDate(c_dt_inicio), msg_erro)
	if  (msg_erro <> "") then
		Response.Write msg_erro
		exit sub
		end if
		
	if not SaldoInicialOk then
		x = x & "<TABLE style='border-bottom:0px;' CellSpacing=0 CellPadding=0>" & chr(13) & _
				"	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MB' ><P class='ALERTA'>&nbsp;Saldo inicial não disponível em " & c_dt_inicio & "&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"</TABLE>"
		Response.write x
		exit sub
		end if
	
	s_sql = " SELECT " & _
				" fabricante, " & _
				" produto, " & _
				" sum(qtde) as saldo " & _
			" FROM t_ESTOQUE_SALDO_DIARIO " & _
			" WHERE " & s_where & _
				" AND  (data = " & bd_formata_data(StrToDate(c_dt_inicio)) &  ") " & _
				" AND estoque in ('" & ID_ESTOQUE_VENDA & "', '" & ID_ESTOQUE_VENDIDO & "')" & _
			" GROUP BY fabricante, produto " & _
			" ORDER by fabricante, produto "


	If Not cria_recordset_otimista(r, msg_erro) then 
		Response.Write msg_erro
		exit sub
		end if
	
	r.open s_sql, cn
	redim vSaldoInicial(0)
	do while Not r.Eof
		redim preserve vSaldoInicial(Ubound(vSaldoInicial)+1)
		set vSaldoInicial(Ubound(vSaldoInicial)) = New cl_TRES_COLUNAS
		vSaldoInicial(Ubound(vSaldoInicial)).c1 = r("produto")
		vSaldoInicial(Ubound(vSaldoInicial)).c2 = r("saldo")
		r.MoveNext
		loop
	
	if r.State <> 0 then r.Close

'	MONTA A CONSULTA DO SALDO FINAL
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	msg_erro = ""
	SaldoFinalOk = saldo_OK (StrToDate(c_dt_termino)+1, msg_erro)
	if  (msg_erro <> "") then
		Response.Write msg_erro
		exit sub
		end if

	s_sql = " SELECT " & _
				" fabricante, " & _
				" produto, " & _
				" sum(qtde) as saldo " & _
			" FROM t_ESTOQUE_SALDO_DIARIO " & _
			" WHERE " & s_where & _
				" AND  (data = " & bd_formata_data(StrToDate(c_dt_termino)+1) &  ") " & _
				" AND estoque in ('" & ID_ESTOQUE_VENDA & "', '" & ID_ESTOQUE_VENDIDO & "')" & _
			" GROUP BY fabricante, produto " & _
			" ORDER by fabricante, produto "

	If Not cria_recordset_otimista(r, msg_erro) then 
		Response.Write msg_erro
		exit sub
		end if
	
	r.open s_sql, cn
	redim vSaldoFinal(0)
	do while Not r.Eof
		redim preserve vSaldoFinal(Ubound(vSaldoFinal)+1)
		set vSaldoFinal(Ubound(vSaldoFinal)) = New cl_TRES_COLUNAS
		vSaldoFinal(Ubound(vSaldoFinal)).c1 = r("produto")
		vSaldoFinal(Ubound(vSaldoFinal)).c2 = r("saldo")
		r.MoveNext
		loop
	
	if r.State <> 0 then r.Close
	
'	MONTA A CONSULTA P/ OBTER DADOS DO LOG DE MOVIMENTAÇÃO NO ESTOQUE
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'	PERÍODO
	if IsDate(c_dt_inicio) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
		
'	OPERAÇÕES QUE SENSIBILIZAM ESTOQUE DE VENDA / VENDIDO
	if s_where <> "" then 
		s_where = s_where & " AND OPERACAO IN (" & _
			"'" & OP_ESTOQUE_LOG_ENTRADA & "', " & _
			"'" & OP_ESTOQUE_LOG_ENTRADA_VIA_KIT & "', " & _
			"'" & OP_ESTOQUE_LOG_CONVERSAO_KIT & "', " & _
			"'" & OP_ESTOQUE_LOG_ENTREGA & "', " & _
			"'" & OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE & "', " & _
			"'" & OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM & "', " & _
			"'" & OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA & "', " & _
			"'" & OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA & "', " & _
			"'" & OP_ESTOQUE_LOG_TRANSFERENCIA & "')" 
		
		end if

	if s_where <> "" then s_where = " WHERE" & s_where
	
	s_sql = "SELECT " & _
				"produto, " & _
				"SUM(qtde_atendida) as qtde, " & _
				"operacao, " & _
				"cod_estoque_origem, " & _
				"cod_estoque_destino " & _
			"FROM t_ESTOQUE_LOG" & _
			s_where & _
			" GROUP BY produto, operacao, cod_estoque_origem, cod_estoque_destino" & _
			" ORDER BY produto "

	n_reg = 0

	r.open s_sql, cn
	
	strProdutoAnterior  = "-PP-PP-PP-PP-PP-PP"
	redim vEstoqueLog(0)
	set vEstoqueLog(Lbound(vEstoqueLog)) = New cl_ESTOQUE_LOG
	with vEstoqueLog(Lbound(vEstoqueLog))
		.tipo = " "
		.produto = " "
		.operacao = " "
		.estoque_origem = " "
		.estoque_destino = " "
		.qtde = " "
		.saldo = " "
		end with

	'CABEÇALHO DA TABELA
	x = "<TABLE style='border-bottom:0px;' width='649' CellSpacing=0 CellPadding=0>" & chr(13)
	Response.Write x

	do while Not r.Eof
		
		' CONTAGEM
		n_reg = n_reg + 1
		
		' MUDOU DE PRODUTO?
		strProdutoAtual = ""
		if not IsNull(r("produto")) then strProdutoAtual = r("produto")
		
		if strProdutoAtual <> strProdutoAnterior then
			if (strProdutoAnterior <> "-PP-PP-PP-PP-PP-PP") then
				prooduto_finaliza vEstoqueLog, vSaldoFinal, strProdutoAnterior, SaldoFinalOk
				x = tabela_monta (vEstoqueLog)
				Response.Write x
				redim preserve vEstoqueLog(0)
				end if
			with vEstoqueLog(Ubound(vEstoqueLog))
				.produto = strProdutoAtual
				.tipo = TP_SALDO
				.Operacao = "Saldo inicial"
				if localiza_cl_tres_colunas(vSaldoInicial, strProdutoAtual, indice_localizado) then
					.Saldo = vSaldoInicial(indice_localizado).c2
				else 
					.Saldo = 0
					end if
				end with
			strProdutoAnterior = strProdutoAtual
			end if
		

		saida_estoque = (Trim(r("cod_estoque_origem")) = ID_ESTOQUE_VENDA) or (Trim(r("cod_estoque_origem")) = ID_ESTOQUE_VENDIDO) 
		entrada_estoque = (Trim(r("cod_estoque_destino")) = ID_ESTOQUE_VENDA) or (Trim(r("cod_estoque_destino")) = ID_ESTOQUE_VENDIDO)
		if ( entrada_estoque xor saida_estoque ) then
			redim preserve vEstoqueLog(Ubound(vEstoqueLog)+1)
			set vEstoqueLog(Ubound(vEstoqueLog)) = New cl_ESTOQUE_LOG
			with vEstoqueLog(Ubound(vEstoqueLog))
				.tipo = TP_OPERACAO
				.produto = r("produto")
				.operacao = r("operacao")
				.estoque_origem = r("cod_estoque_origem")
				.estoque_destino = r("cod_estoque_destino")
				if (saida_estoque) then
					.qtde = -r("qtde")
					end if
				if (entrada_estoque) then
					.qtde = r("qtde")
					end if
				.saldo = vEstoqueLog(Ubound(vEstoqueLog)-1).saldo + .qtde
				end with
			end if
		
		
		r.MoveNext
		loop

	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		'x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MB' colspan='9'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		Response.write x
	else
		prooduto_finaliza vEstoqueLog, vSaldoFinal, strProdutoAtual, SaldoFinalOk
		x = tabela_monta (vEstoqueLog)
		Response.Write x
		
		end if

  ' FECHA TABELA
	x = "</TABLE>" & chr(13)
	Response.Write x


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
.tdOp{
	vertical-align: top;
	width: 170px;
	}
.tdEstOrigem{
	vertical-align: top;
	width: 95px;
	}
.tdEstDestino{
	vertical-align: top;
	width: 95px;
	}
.tdQtd{
	vertical-align: top;
	width: 50px;
	}
.tdSaldo{
	vertical-align: top;
	width: 60px;
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

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Auditoria do Estoque</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Período:&nbsp;</p></td><td valign='top' width='99%'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s = c_fabricante
	if s <> "" then
		s_aux = x_fabricante(s)
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		end if
	if s = "" then s = "todos"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Fabricante:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)
	
	s = c_produto
	if s <> "" then
		s_aux = produto_formata_descricao_em_html(produto_descricao_html(c_fabricante, s))
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		end if
	if s = "" then s = "todos"
	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='baseline' NOWRAP>" & _
					"<p class='N'>Produto:&nbsp;</p></td><td valign='baseline'>" & _
					"<p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP>" & _
					"<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
					"<p class='N'>" & formata_data_hora(Now) & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
	
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
