<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelDivergenciaClienteIndicadorExec.asp
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
	if Not operacao_permitida(OP_CEN_REL_DIVERGENCIA_CLIENTE_INDICADOR, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro
	dim c_dt_inicio, c_dt_termino, c_vendedor

	alerta = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))

	if c_dt_inicio = "" then
		alerta = "DATA DE INÍCIO DO PERÍODO NÃO FOI PREENCHIDA."
	elseif Not IsDate(StrToDate(c_dt_inicio)) then
		alerta = "DATA DE INÍCIO DO PERÍODO É INVÁLIDA."
	elseif (c_dt_termino<>"") And (Not IsDate(StrToDate(c_dt_termino))) then
		alerta = "DATA DE TÉRMINO DO PERÍODO É INVÁLIDA."
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
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim s, s_aux, s_sql, x, cab_table, cab, vendedor_a, n_reg, n_reg_total
dim intQtdeVendedores, intQtdeSubTotalPedidos, intQtdeTotalPedidos, intLargColCliente
dim s_where

'	CRITÉRIOS COMUNS
	s_where = _
			" (t_CLIENTE.indicador <> t_PEDIDO.indicador)" & _
			" AND (t_CLIENTE.indicador IS NOT NULL)" & _
			" AND (LEN(t_CLIENTE.indicador) > 0)" & _
			" AND (t_PEDIDO.indicador IS NOT NULL)" & _
			" AND (LEN(t_PEDIDO.indicador) > 0)"
				

	if c_vendedor <> "" then
		s = substitui_caracteres(c_vendedor, "*", BD_CURINGA_TODOS)
		s_aux = "="
		if Instr(1, s, BD_CURINGA_TODOS) <> 0 then s_aux = "LIKE"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.vendedor " & s_aux & " '" & s & "'" & SQL_COLLATE_CASE_ACCENT & ")"
		end if
	
'	CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	if IsDate(c_dt_inicio) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.data >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
		
	if IsDate(c_dt_termino) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.data < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
		
	if s_where <> "" then s_where = " WHERE" & s_where
	
	s_sql = "SELECT" & _
				" t_PEDIDO.vendedor," & _
				" t_PEDIDO.pedido," & _
				" t_PEDIDO.indicador AS indicador_novo," & _
				" t_CLIENTE.indicador AS indicador_original," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas AS nome_cliente" & _
			" FROM t_PEDIDO INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
			s_where & _
			" ORDER BY" & _
				" t_PEDIDO.vendedor," & _
				" t_PEDIDO.data," & _
				" t_PEDIDO.pedido"

  ' CABEÇALHO
	intLargColCliente=240
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE' valign='bottom' NOWRAP><P style='width:80px' class='R'>Nº PEDIDO</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:" & Cstr(intLargColCliente) & "px' class='R'>CLIENTE</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:100px' class='R'>INDICADOR<br>ORIGINAL</P></TD>" & chr(13) & _
		  "		<TD class='MTD' valign='bottom' NOWRAP><P style='width:100px' class='R'>NOVO<br>INDICADOR</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	intQtdeVendedores = 0
	intQtdeTotalPedidos = 0
	intQtdeSubTotalPedidos = 0

	vendedor_a = "XXXXXXXXXXXX"
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE VENDEDOR?
		if Trim("" & r("vendedor"))<>vendedor_a then
			vendedor_a = Trim("" & r("vendedor"))
			intQtdeVendedores = intQtdeVendedores + 1
		  ' FECHA TABELA DO VENDEDOR ANTERIOR
			if n_reg_total > 0 then 
				x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
						"		<TD class='MTBE' NOWRAP><p class='Cd'>" & _
						"TOTAL:</p></td>" & chr(13) & _
						"		<TD class='MTBD' colspan='3'><p class='C'>" & formata_inteiro(intQtdeSubTotalPedidos) & " pedido(s)</p></TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

			n_reg = 0
			intQtdeSubTotalPedidos = 0

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

		intQtdeSubTotalPedidos = intQtdeSubTotalPedidos + 1
		intQtdeTotalPedidos = intQtdeTotalPedidos + 1
		
		x = x & "	<TR NOWRAP>"  & chr(13)

	 '> Nº PEDIDO
		x = x & "		<TD class='MDTE' valign='top'><P class='Cn' style='font-weight:bold;'><a href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & _
				Trim("" & r("pedido")) & "</a></P></TD>" & chr(13)

	 '> CLIENTE
		s = Trim("" & r("nome_cliente"))
		x = x & "		<TD class='MTD' valign='top'><P class='Cn' style='width:" & Cstr(intLargColCliente) & "'>" & s & "</P></TD>" & chr(13)

	 '> INDICADOR ORIGINAL
		s = Trim("" & r("indicador_original"))
		x = x & "		<TD class='MTD' valign='top'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	 '> INDICADOR NOVO
		s = Trim("" & r("indicador_novo"))
		x = x & "		<TD class='MTD' valign='top'><P class='Cn'>" & s & "</P></TD>" & chr(13)

		x = x & "	</TR>" & chr(13)

		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
  ' MOSTRA TOTAL DO ÚLTIMO VENDEDOR
	if n_reg <> 0 then 
		x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
				"		<TD class='MTBE' NOWRAP><p class='Cd'>" & _
				"TOTAL:</p></td>" & chr(13) & _
				"		<TD class='MTBD' colspan='3'><p class='C'>" & formata_inteiro(intQtdeSubTotalPedidos) & " pedido(s)</p></td>" & chr(13) & _
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
					"		<TD class='MTBE' NOWRAP><p class='Cd'>" & _
					"TOTAL GERAL:</p></td>" & chr(13) & _
					"		<TD class='MTBD' colspan='3'><p class='C'>" & formata_inteiro(intQtdeTotalPedidos) & " pedido(s)</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='4'><P class='ALERTA'>&nbsp;NENHUM PEDIDO SATISFAZ AOS CRITÉRIOS&nbsp;</P></TD>" & chr(13) & _
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
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir( id_pedido ) {
	fREL.action = "pedido.asp";
	fREL.pedido_selecionado.value = id_pedido;
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
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Divergência Cliente/Indicador</span>
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
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Período:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & s & "</p></td></tr>" & chr(13)

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
