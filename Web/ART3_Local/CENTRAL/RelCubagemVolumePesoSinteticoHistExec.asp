<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =========================================================================
'	  RelCubagemVolumePesoSinteticoHistExec.asp
'     =========================================================================
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
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_SINTETICO_CUBAGEM_VOLUME_PESO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, v, i
	dim c_transportadora, c_dt_entregue_inicio, c_dt_entregue_termino

	alerta = ""

	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))
	c_transportadora = Trim(Request("c_transportadora"))

	if (c_dt_entregue_inicio <> "") And (c_dt_entregue_termino = "") then c_dt_entregue_termino = c_dt_entregue_inicio
	
	if (c_dt_entregue_inicio = "") And (c_dt_entregue_termino = "") then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Nenhuma data foi informada."
		end if
	
	if (c_dt_entregue_inicio <> "") And (c_dt_entregue_termino <> "") then
		if StrToDate(c_dt_entregue_inicio) > StrToDate(c_dt_entregue_termino) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "A data de término do período é anterior à data de início."
			end if
		end if
	
	dim s_nome_transportadora
	s_nome_transportadora = ""
	
	if alerta = "" then
		if c_transportadora <> "" then
			s = "SELECT id, nome FROM t_TRANSPORTADORA WHERE (id='" & c_transportadora & "')"
			set rs = cn.execute(s)
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Transportadora '" & c_transportadora & "' não está cadastrada."
			else
				s_nome_transportadora = Trim("" & rs("nome"))
				end if
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
dim s_sql
dim s, s_aux, x, i, v, cab_table, cab
dim vl_total_cubagem, vl_total_qtde_volumes, vl_total_peso, n_reg, n_reg_total
dim com_projecao, vl_projecao, qtde_dias_mes, qtde_dias_projecao
dim s_where, s_where_devolucao

'	CRITÉRIOS COMUNS
	s_where = ""

'	CRITÉRIOS: ENTREGUE EM
	if IsDate(c_dt_entregue_inicio) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.entregue_data >= " & bd_formata_data(StrToDate(c_dt_entregue_inicio)) & ")"
		end if
	
	if IsDate(c_dt_entregue_termino) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.entregue_data < " & bd_formata_data(StrToDate(c_dt_entregue_termino)+1) & ")"
		end if
	
'	CRITÉRIOS: TRANSPORTADORA
	if c_transportadora <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.transportadora_id = '" & c_transportadora & "')"
		end if
	
	s = s_where
	if s <> "" then s = " AND" & s
	s_sql = "SELECT" & _
				" t_TRANSPORTADORA.id AS id_transportadora, " & _
				" t_TRANSPORTADORA.nome AS nome_transportadora, " & _
				" Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.cubagem) AS cubagem," & _
				" Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.qtde_volumes) AS qtde_volumes," & _
				" Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.peso) AS peso" & _
			" FROM t_PEDIDO" & _
				" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
				" LEFT JOIN t_TRANSPORTADORA ON (t_PEDIDO.transportadora_id=t_TRANSPORTADORA.id)" & _
			" WHERE" & _
				" (st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
				s & _
			" GROUP BY" & _
				" t_TRANSPORTADORA.id, " & _
				" t_TRANSPORTADORA.nome" & _
			" ORDER BY" & _
				" t_TRANSPORTADORA.id, " & _
				" t_TRANSPORTADORA.nome"

  ' CABEÇALHO
	cab_table = "<table cellspacing='0'>" & chr(13)
	cab = "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='MT' align='left' valign='bottom' style='width:300px' nowrap><span class='R'>TRANSPORTADORA</span></td>" & chr(13) & _
		  "		<td class='MTBD' align='right' valign='bottom' style='width:80px'><span class='Rd'>CUBAGEM</span><br /><span class='Rd'>(m3)</span></td>" & chr(13) & _
		  "		<td class='MTBD' align='right' valign='bottom' style='width:80px'><span class='Rd'>QTDE</span><br /><span class='Rd'>VOLUMES</span></td>" & chr(13) & _
		  "		<td class='MTBD' align='right' valign='bottom' style='width:80px'><span class='Rd'>PESO</span><br /><span class='Rd'>(KG)</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)

'	LAÇO P/ LEITURA DO RECORDSET
	x = cab_table & _
		cab

	n_reg = 0
	n_reg_total = 0
	vl_total_cubagem = 0
	vl_total_qtde_volumes = 0
	vl_total_peso = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

		n_reg = 0

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<tr nowrap>" & chr(13)

	 '> TRANSPORTADORA
		x = x & "		<td align='left' valign='middle' class='MDBE'><span class='C'>" & Ucase(Trim("" & r("id_transportadora"))) & " - " & iniciais_em_maiusculas(Trim("" & r("nome_transportadora"))) & "</span></td>" & chr(13)
	
	 '> CUBAGEM
		x = x & "		<td align='right' valign='middle' class='MDB'><span class='Cd'>" & formata_numero(r("cubagem"), 2) & "</span></td>" & chr(13)
		
	 '> QTDE VOLUMES
		x = x & "		<td align='right' valign='middle' class='MDB'><span class='Cd'>" & formata_inteiro(r("qtde_volumes")) & "</span></td>" & chr(13)
		
	 '> PESO
		x = x & "		<td align='right' valign='middle' class='MDB'><span class='Cd'>" & formata_numero(r("peso"), 1) & "</span></td>" & chr(13)

		vl_total_cubagem = vl_total_cubagem + r("cubagem")
		vl_total_qtde_volumes = vl_total_qtde_volumes + r("qtde_volumes")
		vl_total_peso = vl_total_peso + r("peso")
		
		x = x & "	</tr>" & chr(13)
			
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
	'>	TOTAL GERAL
		if n_reg_total  > 1 then
			x = x & "	<tr nowrap style='background:honeydew;'>" & chr(13) & _
					"		<td class='MDBE' align='right' nowrap><span class='Cd'>" & _
					"TOTAL:</span></td>" & chr(13) & _
					"		<td class='MDB' align='right'><span class='Cd'>" & formata_numero(vl_total_cubagem, 2) & "</span></td>" & chr(13) & _
					"		<td class='MDB' align='right'><span class='Cd'>" & formata_inteiro(vl_total_qtde_volumes) & "</span></td>" & chr(13) & _
					"		<td class='MDB' align='right'><span class='Cd'>" & formata_numero(vl_total_peso, 1) & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
			end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS !!
	if n_reg_total = 0 then
		x = cab_table & cab & _
			"	<tr nowrap>" & chr(13) & _
			"		<td class='ME MD MB ALERTA' align='center' colspan='4'><span class='ALERTA'>&nbsp;NENHUM REGISTRO SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</span></td>" & chr(13) & _
			"	</tr>" & chr(13)
		end if

  ' FECHA TABELA DA ÚLTIMA LOJA
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



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">


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

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="<%=c_dt_entregue_inicio%>">
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>">
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório Histórico de Cubagem, Volume e Peso (Sintético)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black;' border='0'>" & chr(13)

	s = ""
	s_aux = Trim(c_dt_entregue_inicio)
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = Trim(c_dt_entregue_termino)
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Entregue em:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & s & "</span></td></tr>" & chr(13)

	s = Trim(c_transportadora)
	if s = "" then
		s = "Todas"
	else
		s = s & " - " & iniciais_em_maiusculas(s_nome_transportadora)
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Transportadora:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>" & chr(13)

	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>" & chr(13)
	
	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td align='left' class="Rc">&nbsp;</td></tr>
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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
