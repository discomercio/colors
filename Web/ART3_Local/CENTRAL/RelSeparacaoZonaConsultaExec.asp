<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================================
'	  RelSeparacaoZonaConsultaExec.asp
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_SEPARACAO_ZONA, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro
	dim c_dt_inicio, c_dt_termino, c_nsu, c_nfe_emitente
	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_nsu = Trim(Request.Form("c_nsu"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))

	alerta = ""
	if c_nsu = "" then
		if (c_dt_inicio = "") And (c_dt_termino = "") then
			alerta=texto_add_br(alerta)
			alerta=alerta & "É necessário informar o período da consulta ou o NSU do relatório."
			end if
		
		if (c_dt_inicio <> "") Or (c_dt_termino <> "") then
			if c_dt_inicio = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Informe a data de início."
				end if
			if c_dt_termino = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Informe a data de término."
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

	if alerta = "" then
		if c_nfe_emitente = "" then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Não foi informado o CD"
		elseif converte_numero(c_nfe_emitente) = 0 then
			alerta=texto_add_br(alerta)
			alerta = alerta & "É necessário definir um CD válido"
			end if
		end if





' ________________________________
' EXECUTA CONSULTA
'
Sub executa_consulta ()
dim s, cab, x, s_sql, s_where, n_reg, s_link_open, s_link_close
dim rs
dim rNfeEmitente

	cab = "<table id='tableRelatorio' class='Q' style='border-bottom:0px;' cellspacing='0' cellpadding='2'>" & chr(13) & _
		"	<tr style='background:azure'>" & chr(13) & _
		"		<th align='center' valign='bottom' class='MD MB tdData' style='vertical-align:bottom;'><span class='Rc'>Data</span></th>" & chr(13) & _
		"		<th align='center' valign='bottom' class='MD MB tdUsuario' style='vertical-align:bottom;'><span class='Rc'>Usuário</span></th>" & chr(13) & _
		"		<th align='right' valign='bottom' class='MD MB tdNsu' style='vertical-align:bottom;'><span class='Rd'>NSU</span></th>" & chr(13) & _
		"		<th align='right' valign='bottom' class='MB tdQtdePed' style='vertical-align:bottom;'><span class='Rd'>Qtde</span><br><span class='Rd'>Pedidos</span></th>" & chr(13) & _
		"	</tr>" & chr(13)
	
	s_sql = "SELECT" & _
				" id," & _
				" dt_hr_emissao," & _
				" usuario," & _
				" (SELECT Count(*) FROM t_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO WHERE (id_wms_etq_n1=tN1.id)) AS qtde_pedidos" & _
			" FROM t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO tN1"

	s_where = ""
	
	if c_dt_inicio <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tN1.dt_cadastro >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
	
	if c_dt_termino <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tN1.dt_cadastro < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
	if c_nsu <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tN1.id = " & c_nsu & ")"
		end if
	
'	OWNER DO PEDIDO
	set rNfeEmitente = le_nfe_emitente(c_nfe_emitente)
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & _
				" (" & _
					"EXISTS (" & _
						"SELECT TOP 1" & _
							" tN2.pedido" & _
						" FROM t_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO tN2" & _
							" INNER JOIN t_PEDIDO ON (tN2.pedido = t_PEDIDO.pedido)" & _
						" WHERE" & _
							" (id_wms_etq_n1 = tN1.id)" & _
							" AND (t_PEDIDO.id_nfe_emitente = " & rNfeEmitente.id & ")" & _
					")" & _
				")"

	if s_where <> "" then s_where = " WHERE" & s_where
	s_sql = s_sql & s_where
	s_sql = s_sql & " ORDER BY tN1.id"
	
'	EXECUTA CONSULTA
	set rs = cn.Execute( s_sql )
	
	x = cab
	n_reg = 0
	do while Not rs.eof 
		n_reg = n_reg + 1
		
		x = x & "	<tr>" & chr(13)
		
		s_link_open = "<a href='javascript:fConcluir(" & chr(34) & Trim("" & rs("id")) & chr(34) & _
				 ")' title='clique para consultar o relatório'>"
		s_link_close = "</a>"
		
	'	DATA
		x = x & "		<td align='center' class='MD MB tdData'><span class='Cc'>" & s_link_open & formata_data_hora_sem_seg(rs("dt_hr_emissao")) & s_link_close & "</span>" & "</td>" & chr(13)
		
	'	USUÁRIO
		x = x & "		<td align='center' class='MD MB tdUsuario'><span class='Cc'>" & s_link_open & Trim("" & rs("usuario")) & s_link_close & "</span></td>" & chr(13)
		
	'	NSU
		x = x & "		<td align='right' class='MD MB tdNsu'><span class='Cd'>" & s_link_open & Trim("" & rs("id")) & s_link_close & "</span></td>" & chr(13)
		
	'	QTDE PEDIDOS
		x = x & "		<td align='right' class='MB tdQtdePed'><span class='Cd'>" & s_link_open & Trim("" & rs("qtde_pedidos")) & s_link_close & "</span></td>" & chr(13)
		
		x = x & "	</tr>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		rs.MoveNext
		loop
	
	if n_reg = 0 then
		x = cab & _
			"	<tr nowrap>" & chr(13) & _
			"		<td colspan='4' align='center' class='MB ALERTA'><span class='ALERTA'>&nbsp;NENHUM REGISTRO SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</span></td>" & chr(13) & _
			"	</tr>" & chr(13)
		end if

	x = x & "</table>" & chr(13)

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

<script type="text/javascript">
	$(document).ready(function() {
		$("#tableRelatorio tr").not(':first').hover(
			function() {
				$(this).css("background", "#98FB98");
			},
			function() {
				$(this).css("background", "");
			}
		)
	});
</script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fConcluir ( id ) {
	fCONS.action = "RelSeparacaoZonaConsultaDetalhe.asp";
	fCONS.nsu_selecionado.value = id;
	fCONS.submit();
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
a
{
	text-decoration: none;
	color: black;
}
.Nni
{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10pt;
	font-weight: normal;
	font-style: italic;
}
.tdData
{
	vertical-align:middle;
	width:70px;
}
.tdUsuario
{
	vertical-align:middle;
	width:100px;
}
.tdNsu
{
	vertical-align:middle;
	width:80px;
}
.tdQtdePed
{
	vertical-align:middle;
	width:80px;
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
<body onload="window.status='Concluído';">

<center>

<form id="fCONS" name="fCONS" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">
<input type="hidden" name="c_nsu" id="c_nsu" value="<%=c_nsu%>">
<input type="hidden" name="nsu_selecionado" id="nsu_selecionado" value="">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Separação (Zona) - Consulta</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<%
	s_filtro ="<table width='649' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black'>" & chr(13)
	
'	PERÍODO
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>Período:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	NSU DO RELATÓRIO
	s = ""
	s_aux = c_nsu
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>NSU do Relatório:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
	
	s_filtro = s_filtro & "</table>" & chr(13)

	Response.Write s_filtro
%>
<br>

<!--  RELATÓRIO  -->
<% executa_consulta %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
