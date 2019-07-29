<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L P E D I D O S C O L O C A D O S E X E C . A S P
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
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_PEDIDOS_COLOCADOS_NO_MES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, lista_loja, s_filtro_loja, v_loja, v, i
	dim c_mes, c_ano, c_loja, c_empresa

	alerta = ""

	c_mes = retorna_so_digitos(Trim(Request.Form("c_mes")))
	c_ano = retorna_so_digitos(Trim(Request.Form("c_ano")))
    c_empresa = Trim(Request.Form("c_empresa"))

	c_loja = Trim(Request.Form("c_loja"))
	lista_loja = substitui_caracteres(c_loja,chr(10),"")
	v_loja = split(lista_loja,chr(13),-1)
	
	if Not IsNumeric(c_mes) then
		alerta = "MÊS INVÁLIDO."
	elseif Not IsNumeric(c_ano) then
		alerta = "ANO INVÁLIDO."
	elseif (CLng(c_mes)<=0) Or (CLng(c_mes)>12) then
		alerta = "MÊS INVÁLIDO."
	elseif (CLng(c_ano)<2003) then
		alerta = "ANO INVÁLIDO."
		end if

	if alerta = "" then
		s = Cstr(Month(Date))
		do while len(s)<2: s = "0" & s: loop
		s = Cstr(Year(Date)) & s
		s_aux = c_mes
		do while len(s_aux)<2: s_aux = "0" & s_aux: loop
		s_aux = c_ano & s_aux
		if (s_aux) > s then alerta = "MÊS DE REFERÊNCIA FUTURO É INVÁLIDO."
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
			strDtRefDDMMYYYY = "01/" & c_mes & "/" & c_ano
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
dim s, s_aux, x, dt, s_where_loja, i, v, cab_table, cab, loja_a
dim vl_total_saida, vl_sub_total_saida, qtde_lojas, n_reg, n_reg_total
dim com_projecao, vl_projecao, qtde_dias_mes, qtde_dias_projecao
dim s_where, s_where_venda, s_where_devolucao, s_cor

'	MÊS ATUAL: CONSULTA APENAS ATÉ A DATA DE ONTEM
	if (CLng(c_mes)=Month(Date)) And (CLng(c_ano)=Year(Date)) then
		dt = Date
		com_projecao = True
		qtde_dias_projecao = DateDiff("d",StrToDate("01/" & c_mes & "/" & c_ano), dt)
		qtde_dias_mes = DateDiff("d", StrToDate("01/" & c_mes & "/" & c_ano), DateAdd("m",1,StrToDate("01/" & c_mes & "/" & c_ano)))
	else
		dt = DateAdd("m", 1, StrToDate("01/" & c_mes & "/" & c_ano))
		com_projecao = False
		end if

'	CRITÉRIOS COMUNS
	s_where = ""
	s_where_loja = ""
	for i=Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
				s_where_loja = s_where_loja & " (CONVERT(smallint, t_PEDIDO.loja) = " & v_loja(i) & ")"
			else
				s = ""
				if v(Lbound(v))<>"" then 
					if s <> "" then s = s & " AND"
					s = s & " (CONVERT(smallint, t_PEDIDO.loja) >= " & v(Lbound(v)) & ")"
					end if
				if v(Ubound(v))<>"" then
					if s <> "" then s = s & " AND"
					s = s & " (CONVERT(smallint, t_PEDIDO.loja) <= " & v(Ubound(v)) & ")"
					end if
				if s <> "" then 
					if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
					s_where_loja = s_where_loja & " (" & s & ")"
					end if
				end if
			end if
		next

    if c_empresa <> "" then
        if s_where <> "" then s_where = s_where & " AND"
        s_where = s_where & " (t_PEDIDO.id_nfe_emitente = " & c_empresa & ")"
    end if
		
	if s_where_loja <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s_where_loja & ")"
		end if

'	CRITÉRIOS PARA PEDIDOS DE VENDA NORMAIS
	s_where_venda = ""
	if IsDate(dt) then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.data < " & bd_formata_data(dt) & ")"
		end if
		
	if c_mes <> "" then
		if s_where_venda <> "" then s_where_venda = s_where_venda & " AND"
		s_where_venda = s_where_venda & " (t_PEDIDO.data >= " & bd_formata_data(StrToDate("01/" & c_mes & "/" & c_ano)) & ")"
		end if
	
'	CRITÉRIOS PARA DEVOLUÇÕES
	s_where_devolucao = ""
	if IsDate(dt) then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(dt) & ")"
		end if
		
	if c_mes <> "" then
		if s_where_devolucao <> "" then s_where_devolucao = s_where_devolucao & " AND"
		s_where_devolucao = s_where_devolucao & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(StrToDate("01/" & c_mes & "/" & c_ano)) & ")"
		end if

	s = s_where
	if (s <> "") And (s_where_venda <> "") then s = s & " AND"
	s = s & s_where_venda
	if s <> "" then s = " AND" & s
	s_sql = "SELECT t_PEDIDO.loja AS loja, CONVERT(smallint,t_PEDIDO.loja) AS numero_loja," & _
			" t_PEDIDO.data AS data," & _
			" Sum(t_PEDIDO_ITEM.qtde*t_PEDIDO_ITEM.preco_venda) AS valor_saida" & _
			" FROM t_PEDIDO INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
			" WHERE (st_entrega<>'" & ST_ENTREGA_CANCELADO & "')" & _
			s & _
			" GROUP BY t_PEDIDO.loja, t_PEDIDO.data"

	s = s_where
	if (s <> "") And (s_where_devolucao <> "") then s = s & " AND"
	s = s & s_where_devolucao
	if s <> "" then s = " WHERE" & s
	s_sql = s_sql & " UNION " & _
			"SELECT t_PEDIDO.loja AS loja, CONVERT(smallint,t_PEDIDO.loja) AS numero_loja," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data AS data," & _
			" Sum(-t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS valor_saida" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.pedido=t_PEDIDO.pedido)" & _
			s & _
			" GROUP BY t_PEDIDO.loja, t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data"

	s_sql = s_sql & " ORDER BY numero_loja, data, valor_saida DESC"

  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & _
		  "		<TD class='MDTE' align='center' valign='bottom' NOWRAP><P style='width:90px' class='Rc'>DATA</P></TD>" & _
		  "		<TD class='MTD' align='center' valign='bottom' NOWRAP><P style='width:90px' class='Rc'>DIA DA SEMANA</P></TD>" & _
		  "		<TD class='MTD' align='right' valign='bottom'><P style='width:120px' class='Rd' style='font-weight:bold;'>FATURAMENTO</P></TD>" & _
		  "	</TR>" & chr(13)

	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_lojas = 0
	vl_total_saida = 0
	vl_sub_total_saida = 0
	
	loja_a = "XXXXX"
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE LOJA?
		if Trim("" & r("loja"))<>loja_a then
			loja_a = Trim("" & r("loja"))
			qtde_lojas = qtde_lojas + 1
		  ' FECHA TABELA DA LOJA ANTERIOR
			if n_reg_total > 0 then 
				s_cor="black"
				if vl_sub_total_saida < 0 then s_cor="red"
				x = x & "<TR NOWRAP style='background: #FFFFDD'>" & _
						"<TD COLSPAN='2' class='MTBE' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
						"Faturamento Acumulado:</p></td>" & _
						"<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida) & "</p></td>" & _
						"</TR>"
				if com_projecao then
					if qtde_dias_projecao <= 0 then 
						vl_projecao = 0
					else
						vl_projecao = qtde_dias_mes * (vl_sub_total_saida / qtde_dias_projecao)
						s_cor="black"
						if vl_projecao < 0 then s_cor="red"
						x = x & "<TR NOWRAP style='background: #FFFFDD'>" & _
								"<TD COLSPAN='2' class='MEB' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
								"Faturamento Estimado:</p></td>" & _
								"<TD class='MDB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_projecao) & "</p></td>" & _
								"</TR>"
						end if
					end if
					
				x = x & "</TABLE>" & chr(13)
				Response.Write x
				x="<BR>"
				end if

			n_reg = 0
			vl_sub_total_saida = 0

			if n_reg_total > 0 then x = x & "<BR>"
			s = Trim("" & r("loja"))
			s_aux = x_loja(s)
			if s_aux <> "" then s_aux = iniciais_em_maiusculas(s_aux)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then x = x & "<TR><TD class='MDTE' NOWRAP COLSPAN='3' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td></tr>" & chr(13)
			x = x & cab
			end if
		

	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>"

		s_cor="black"
		if IsNumeric(r("valor_saida")) then if CCur(r("valor_saida")) < 0 then s_cor="red"

	 '> DATA
		x = x & "		<TD align='center' class='MDTE'><P class='Cc' style='color:" & s_cor & ";'>" & formata_data(r("data")) & "</P></TD>"
	
	 '> DIA DA SEMANA
		x = x & "		<TD class='MTD'><P class='Cc' style='color:" & s_cor & ";'>" & LCase(dia_da_semana(r("data"),True)) & "</P></TD>"
		
	 '> VALOR SAÍDA
		x = x & "		<TD align='right' class='MTD'><P class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(r("valor_saida")) & "</P></TD>"

		vl_sub_total_saida = vl_sub_total_saida + r("valor_saida")
		vl_total_saida = vl_total_saida + r("valor_saida")
		
		x = x & "</TR>" & chr(13)
			
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		s_cor="black"
		if vl_sub_total_saida < 0 then s_cor="red"
		x = x & "<TR NOWRAP style='background: #FFFFDD'>" & _
				"<TD COLSPAN='2' class='MTBE' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
				"Faturamento Acumulado:</p></td>" & _
				"<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_saida) & "</p></td>" & _
				"</TR>"

		if com_projecao then
			if qtde_dias_projecao <= 0 then 
				vl_projecao = 0
			else
				vl_projecao = qtde_dias_mes * (vl_sub_total_saida / qtde_dias_projecao)
				s_cor="black"
				if vl_projecao < 0 then s_cor="red"
				x = x & "<TR NOWRAP style='background: #FFFFDD'>" & _
						"<TD COLSPAN='2' class='MEB' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
						"Faturamento Estimado:</p></td>" & _
						"<TD class='MDB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_projecao) & "</p></td>" & _
						"</TR>"
				end if
			end if
		
	'>	TOTAL GERAL
		if qtde_lojas > 1 then
			s_cor="black"
			if vl_total_saida < 0 then s_cor="red"
			x = x & "<TR><TD COLSPAN='3' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
					"<TR><TD COLSPAN='3' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
					"<TR NOWRAP style='background:honeydew'>" & _
					"<TD COLSPAN='2' class='MTBE' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
					"Faturamento Total Acumulado:</p></td>" & _
					"<TD class='MTBD'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_saida) & "</p></td>" & _
					"</TR>"
			if com_projecao then
				if qtde_dias_projecao <= 0 then
					vl_projecao = 0
				else
					vl_projecao = qtde_dias_mes * (vl_total_saida / qtde_dias_projecao)
					s_cor="black"
					if vl_projecao < 0 then s_cor="red"
					x = x & "<TR NOWRAP style='background:honeydew'>" & _
							"<TD COLSPAN='2' class='MEB' NOWRAP><p class='Cd' style='color:" & s_cor & ";'>" & _
							"Faturamento Total Estimado:</p></td>" & _
							"<TD class='MDB'><p class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_projecao) & "</p></td>" & _
							"</TR>"
					end if
				end if
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & _
				"		<TD class='MT' colspan='3'><P class='ALERTA'>&nbsp;NENHUM PEDIDO SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & _
				"	</TR>"
		end if

  ' FECHA TABELA DA ÚLTIMA LOJA
	x = x & "</TABLE>"
	
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
<input type="hidden" name="c_mes" id="c_mes" value="<%=c_mes%>">
<input type="hidden" name="c_ano" id="c_ano" value="<%=c_ano%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pedidos Colocados no Mês</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>"

	s = mes_por_extenso(c_mes,True) & " / " & c_ano
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Mês de Referência:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

    s = c_empresa
	if s = "" then
		s = "todas"
	else
		s =  obtem_apelido_empresa_NFe_emitente(c_empresa)
    end if
        s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Empresa:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"  

	s_filtro_loja = ""
	for i = Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
				s_filtro_loja = s_filtro_loja & v_loja(i)
			else
				if (v(Lbound(v))<>"") And (v(Ubound(v))<>"") then 
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Lbound(v)) & " a " & v(Ubound(v))
				elseif (v(Lbound(v))<>"") And (v(Ubound(v))="") then
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Lbound(v)) & " e acima"
				elseif (v(Lbound(v))="") And (v(Ubound(v))<>"") then
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Ubound(v)) & " e abaixo"
					end if
				end if
			end if
		next
	s = s_filtro_loja
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Loja(s):&nbsp;</p></td><td valign='top'>" & _
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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
