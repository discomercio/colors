<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L S E P A R A C A O . A S P
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
	if Not operacao_permitida(OP_CEN_REL_SEPARACAO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta, s, s_aux, c_dt_inicio, c_dt_termino, c_transportadora, c_nfe_emitente
	alerta = ""

	c_dt_inicio = Trim(Request.Form("c_dt_inicio"))
	c_dt_termino = Trim(Request.Form("c_dt_termino"))
	c_transportadora = Trim(Request.Form("c_transportadora"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))


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





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA RELATORIO
'
sub consulta_relatorio
dim r
dim s, s_aux, s_sql, x, pedido_a, fabricante_a, s_obs_2, s_transportadora
dim i, n_reg_total, idx
dim vProd()
dim v
dim rNfeEmitente

	s_sql = "SELECT" & _
				" t_PEDIDO.pedido," & _
				" t_PEDIDO.obs_2," & _
				" t_PEDIDO.loja," & _
				" t_PEDIDO.transportadora_id," & _
				" t_PEDIDO_ITEM.fabricante," & _
				" t_PEDIDO_ITEM.produto," & _
				" t_PEDIDO_ITEM.qtde,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" t_PEDIDO.endereco_nome_iniciais_em_maiusculas AS nome_iniciais_em_maiusculas,"
	else
		s_sql = s_sql & _
				" t_CLIENTE.nome_iniciais_em_maiusculas,"
		end if

	s_sql = s_sql & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html" & _
			" FROM t_PEDIDO" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
				" LEFT JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
				" LEFT JOIN t_PRODUTO ON ((t_PEDIDO_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_PEDIDO_ITEM.produto=t_PRODUTO.produto))" & _
			" WHERE" & _
				" (t_PEDIDO.st_entrega='" & ST_ENTREGA_SEPARAR & "')" & _
				" AND (t_PEDIDO.a_entregar_data_marcada IS NOT NULL)" & _
				" AND (t_PEDIDO.st_etg_imediata = " & COD_ETG_IMEDIATA_SIM & ")" & _
				" AND (t_PEDIDO__BASE.analise_credito = " & COD_AN_CREDITO_OK & ")"
	
	if IsDate(c_dt_inicio) then
		s_sql = s_sql & " AND (t_PEDIDO.a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_inicio)) & ")"
		end if
	
	if IsDate(c_dt_termino) then
		s_sql = s_sql & " AND (t_PEDIDO.a_entregar_data_marcada < " & bd_formata_data(StrToDate(c_dt_termino)+1) & ")"
		end if
	
	if c_transportadora <> "" then
		s_sql = s_sql & " AND (t_PEDIDO.transportadora_id = '" & c_transportadora & "')"
		end if
	
'	OWNER DO PEDIDO
	set rNfeEmitente = le_nfe_emitente(c_nfe_emitente)
	s_sql = s_sql & " AND (t_PEDIDO.id_nfe_emitente = " & rNfeEmitente.id & ")"
	
	s_sql = s_sql & " ORDER BY t_PEDIDO.data, t_PEDIDO.hora, t_PEDIDO.pedido, t_PEDIDO_ITEM.fabricante, t_PEDIDO_ITEM.produto"


'	AS ROTINAS DE ORDENAÇÃO USAM VETORES QUE SE INICIAM NA POSIÇÃO 1
	redim vProd(1)
	for i = Lbound(vProd) to Ubound(vProd)
		set vProd(i) = New cl_DUAS_COLUNAS
		with vProd(i)
			.c1 = ""
			.c2 = 0
			end with
		next
		
	x = ""
	n_reg_total = 0
	pedido_a = "XXXXXXXXXX"
	set r = cn.execute(s_sql)
	
	do while Not r.Eof
	  ' CONTAGEM
		n_reg_total = n_reg_total + 1

		if Trim("" & r("pedido")) <> pedido_a then
			pedido_a = Trim("" & r("pedido"))
		'	FECHA TABELA DO PEDIDO ANTERIOR
			if x <> "" then x = x & _ 
								"		</TABLE>" & chr(13) & _
								"	</TD></TR>" & chr(13) & _
								"</TABLE>" & chr(13) & _
								"<br>" & chr(13)

		'	Nº DA NF
			s_obs_2 = Trim("" & r("obs_2"))
			if s_obs_2 = "" then s_obs_2 = "&nbsp;"
			
		'	TRANSPORTADORA
			s_transportadora = iniciais_em_maiusculas(Trim("" & r("transportadora_id")))
			if s_transportadora = "" then s_transportadora = "&nbsp;"
			
		'	TABELA P/ O PRÓXIMO PEDIDO
			x = x & chr(13) & _
				"<TABLE cellSpacing=0 cellPadding=0>" & chr(13) & _
				"	<TR><TD>" & chr(13) & _
				"		<TABLE class='Q' cellSpacing=0 cellPadding=0>" & chr(13) & _
				"			<TR style='background:#FFF0E0' NOWRAP>" & chr(13)& _
				"				<TD valign='bottom' nowrap class='MD' style='width:75px;'><P class='C'><a href='javascript:fRELConcluir(" & _
								chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & _
									Trim("" & r("pedido")) & "</a></P></TD>" & chr(13) & _
				"				<TD valign='bottom' nowrap class='MD' style='width:78px;'><P class='C'>" & _
									s_obs_2 & "</P></TD>" & chr(13) & _
				"				<TD valign='bottom' nowrap class='MD' style='width:45px;'><P class='C'>Lj&nbsp;" & _
									Trim("" & r("loja")) & "</P></TD>" & chr(13) & _
				"				<TD valign='bottom' class='MD' style='width:365px;'><P class='C'>" & _
									Trim("" & r("nome_iniciais_em_maiusculas")) & "</P></TD>" & chr(13) & _
				"				<TD valign='bottom' nowrap style='width:80px;'><P class='C'>" & _
									s_transportadora & "</P></TD>" & chr(13) & _
				"			</TR>" & chr(13) & _
				"		</TABLE>" & chr(13) & _
				"	</TD></TR>" & chr(13) & _
				"	<TR><TD>" & chr(13) & _
				"		<TABLE width='100%' cellSpacing=0 cellPadding=0>" & chr(13)
			end if
		
	'	LISTAGEM
		x = x & "			<TR NOWRAP>" & chr(13)

	 '> QTDE
		x = x & "				<TD class='MDBE' valign='bottom' style='width:40px;'><P class='Cd'>&nbsp;" & _
			formata_inteiro(r("qtde")) & "</P></TD>"  & chr(13)

	 '> PRODUTO
		x = x & "				<TD class='MDB' valign='bottom' NOWRAP><P class='C' NOWRAP>&nbsp;" & _
			produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & _
			"&nbsp;&nbsp;(Cód:&nbsp;" & Trim("" & r("produto")) & _
			"&nbsp;&nbsp;Fabr:&nbsp;" & Trim("" & r("fabricante")) & ")" & _
			"</P></TD>" & chr(13)
		
		x = x & "			</TR>" & chr(13)

	'	TOTALIZAÇÃO
		s = Trim("" & r("fabricante")) & "|" & Trim("" & r("produto"))
		if localiza_cl_duas_colunas(vProd, s, idx) then
			with vProd(idx)
				.c2 = .c2 + CLng(r("qtde"))
				end with
		else
			if (vProd(Ubound(vProd)).c1<>"") then
				redim preserve vProd(Ubound(vProd)+1)
				set vProd(Ubound(vProd)) = New cl_DUAS_COLUNAS
				end if
			with vProd(Ubound(vProd))
				.c1 = Trim("" & r("fabricante")) & "|" & Trim("" & r("produto"))
				.c2 = CLng(r("qtde"))
				end with
			ordena_cl_duas_colunas vProd, 1, Ubound(vProd)
			end if

		r.movenext
		loop
		
'	FINALIZAÇÃO
	if n_reg_total <> 0 then 
	'	FECHA ÚLTIMA TABELA
		x = x & "		</TABLE>" & chr(13) & _
				"	</TD></TR>" & chr(13) & _
				"</TABLE>" & chr(13)
	'	TOTAIS
		x = x & chr(13) & "<br><br>" & chr(13) & _
			"<TABLE class='Q' style='border-bottom:0px;' cellSpacing=0 cellPadding=0>" & chr(13) & _
			"	<TR style='background:azure' NOWRAP>" & chr(13) & _
			"		<TD class='MB' align='center' COLSPAN='3'><P class='C'>&nbsp;TOTAL</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
	'	LEMBRANDO QUE O VETOR ESTÁ ORDENADO
		fabricante_a = "XXXXX"
		for i = Lbound(vProd) to Ubound(vProd)
			with vProd(i)
				if Trim("" & .c1) <> "" then
					v = Split(.c1, "|", -1)
					if Trim("" & v(0)) <> fabricante_a then
						fabricante_a = Trim("" & v(0))
						s = Trim("" & v(0))
						s_aux = ucase(x_fabricante(s))
						if (s<>"") And (s_aux<>"") then s = s & " - "
						s = s & s_aux
						x = x & "	<TR NOWRAP>" & chr(13) & _
							"		<TD colspan='3' class='MB'><P class='Cc'>" & _
							s & _
							"	</TR>" & chr(13)
						end if
					
					x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD class='MDB' valign='bottom' style='width:65px;'><P class='C'>&nbsp;" & _
						Trim("" & v(1)) & "</P></TD>"  & chr(13) & _
						"		<TD class='MDB' valign='bottom' style='width:400px;' NOWRAP><P class='C' NOWRAP>&nbsp;" & _
						produto_formata_descricao_em_html(produto_descricao_html(v(0), v(1))) & _
						"</P></TD>" & chr(13) & _
						"		<TD class='MB' valign='bottom' style='width:50px;'><P class='Cd'>&nbsp;" & _
						formata_inteiro(.c2) & "</P></TD>"  & chr(13) & _
						"	</TR>" & chr(13)
					end if
				end with
			next
		x = x & "</TABLE>"
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = "<TABLE class='Q' cellSpacing=0>" & chr(13) & _
			"	<TR>" & chr(13) & _
			"		<TD style='width:500px;'><P class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS PARA SEPARAR.&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13) & _
			"</TABLE>"
		end if

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

<style type="text/css">
P.C { font-size:10pt; }
P.Cc { font-size:10pt; }
P.Cd { font-size:10pt; }
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
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>">
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Separação</span>
	<br>
	<%	s = "Período: "
		s_aux = c_dt_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " a "
		s_aux = c_dt_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s = "<span class='STP'>" & s & "</span>"
		s_aux = c_transportadora
		if s_aux = "" then s_aux = "N.I."
		s = s & "<br>" & "<span class='STP'>Transportadora: " & s_aux & "</span>"
		s = s & "<br>" & "<span class='N'>Emissão:&nbsp;" & formata_data_hora(Now) & "</span>"
		Response.Write s
	%>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

	
<!--  RELATÓRIO  -->
<% consulta_relatorio %>

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
