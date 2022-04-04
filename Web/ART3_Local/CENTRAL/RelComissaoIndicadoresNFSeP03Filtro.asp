<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelComissaoIndicadoresNFSeP03Filtro.asp
'     ========================================================
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

	dim s_sql, qtde_indicadores_encontrados
	dim alerta
	alerta = ""
	qtde_indicadores_encontrados = 0
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	FILTROS
	dim c_cnpj_nfse
	dim ckb_id_indicador
	dim ckb_st_entrega_entregue, c_dt_entregue_inicio, c_dt_entregue_termino
	dim ckb_comissao_paga_sim, ckb_comissao_paga_nao
	dim ckb_st_pagto_pago, ckb_st_pagto_nao_pago, ckb_st_pagto_pago_parcial
	dim rb_visao
	
	c_cnpj_nfse = retorna_so_digitos(Request.Form("c_cnpj_nfse"))
	ckb_id_indicador = Trim(Request.Form("ckb_id_indicador"))
	ckb_st_entrega_entregue = Trim(Request.Form("ckb_st_entrega_entregue"))
	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))

	ckb_comissao_paga_sim = Trim(Request.Form("ckb_comissao_paga_sim"))
	ckb_comissao_paga_nao = Trim(Request.Form("ckb_comissao_paga_nao"))

	ckb_st_pagto_pago = Trim(Request.Form("ckb_st_pagto_pago"))
	ckb_st_pagto_nao_pago = Trim(Request.Form("ckb_st_pagto_nao_pago"))
	ckb_st_pagto_pago_parcial = Trim(Request.Form("ckb_st_pagto_pago_parcial"))

	rb_visao = Trim(Request.Form("rb_visao"))
	
	if c_cnpj_nfse = "" then
		alerta = "CNPJ do emitente da NFS-e não foi informado"
	elseif Not cnpj_cpf_ok(c_cnpj_nfse) then
		alerta = "CNPJ do emitente da NFS-e é inválido"
		end if

	if alerta = "" then
		if ckb_id_indicador = "" then
			alerta = "Nenhum indicador foi selecionado"
			end if
		end if

	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________


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



<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function () {
		$(".aviso").css('display', 'none');
		$("#c_dt_entregue_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_entregue_termino").hUtilUI('datepicker_filtro_final');
	});
</script>

<script language="JavaScript" type="text/javascript">
	function fRetornaInicio(f) {
		f.action = "RelComissaoIndicadoresNFSeP01Filtro.asp";
		f.submit();
	}

	function fFILTROConfirma(f) {
		var strDtRefYYYYMMDD, strDtRefDDMMYYYY;

		if (f.ckb_st_entrega_entregue.checked) {
			if (!consiste_periodo(f.c_dt_entregue_inicio, f.c_dt_entregue_termino)) return;
		}

		//  Período de consulta está restrito por perfil de acesso?
		if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) != "") {
			strDtRefDDMMYYYY = trim(f.c_dt_entregue_inicio.value);
			if (trim(strDtRefDDMMYYYY) != "") {
				strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
				if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
					alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
					return;
				}
			}
			strDtRefDDMMYYYY = trim(f.c_dt_entregue_termino.value);
			if (trim(strDtRefDDMMYYYY) != "") {
				strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
				if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
					alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
					return;
				}
			}
		}

		dCONFIRMA.style.visibility = "hidden";

		f.action = "RelComissaoIndicadoresNFSeP04Exec.asp";

		window.status = "Aguarde ...";
		f.submit();
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">

	.aviso {
		font-family: Arial, Helvetica, sans-serif;
		font-size: 8pt;
		font-weight: bold;
		font-style: normal;
		margin: 0pt 0pt 0pt 0pt;
		color: #f00;
		display: none;
	}

	.TdSpn {
		display: block;
		padding:2px;
		font-weight:bold;
	}

	.TdApelido {
		width: 160px;
		text-align:left;
		vertical-align:middle;
	}

	.TdCnpjCadastro {
		width: 110px;
		text-align:left;
		vertical-align:middle;
	}

	.TdRazaoSocialCadastro {
		width: 190px;
		text-align:left;
		vertical-align:middle;
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
<body>
<center>

<form id="fFILTRO" name="fFILTRO" method="post" onsubmit="if (!fFILTROConfirma(fFILTRO)) return false;">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="ckb_id_indicador" id="ckb_id_indicador" value="<%=ckb_id_indicador%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="690" cellpadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (via NFS-e)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table width="690" class="Qx" cellspacing="0">
<!--  STATUS DE ENTREGA  -->
<tr bgcolor="#FFFFFF">
<td class="MT" align="left" nowrap><span class="PLTe">PERÍODO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_entregue" name="ckb_st_entrega_entregue" onclick="if (fFILTRO.ckb_st_entrega_entregue.checked) fFILTRO.c_dt_entregue_inicio.focus();"
			value="<%=ST_ENTREGA_ENTREGUE%>"
			<% if (c_dt_entregue_inicio <> "") Or (c_dt_entregue_termino <> "") then Response.Write " checked"%>
			><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_entregue.click();">Pedidos entregues entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entregue_termino.focus(); else fFILTRO.ckb_st_entrega_entregue.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_entregue.checked=true;" onchange="if (trim(this.value)!='') fFILTRO.ckb_st_entrega_entregue.checked=true;"
			<% if c_dt_entregue_inicio <> "" then Response.Write " value=" & chr(34) & c_dt_entregue_inicio & chr(34)%>
			/>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entregue_termino" id="c_dt_entregue_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_st_entrega_entregue.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_entregue.checked=true;"  onchange="if (trim(this.value)!='') fFILTRO.ckb_st_entrega_entregue.checked=true;"
			<% if c_dt_entregue_termino <> "" then Response.Write " value=" & chr(34) & c_dt_entregue_termino & chr(34)%>
			/>
		</td>
	</tr>
	</table>
</td></tr>

<!--  COMISSÃO PAGA  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">COMISSÃO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_comissao_paga_sim" name="ckb_comissao_paga_sim"
			value="ON"
			<% if ckb_comissao_paga_sim <> "" then Response.Write " checked" %>
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_comissao_paga_sim.click();">Paga</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_comissao_paga_nao" name="ckb_comissao_paga_nao"
			value="ON"
			<% if ckb_comissao_paga_nao <> "" then Response.Write " checked" %>
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_comissao_paga_nao.click();">Não-Paga</span>
		</td>
	</tr>
	</table>
</td>
</tr>

<!--  STATUS DE PAGAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">STATUS DE PAGAMENTO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_pago" name="ckb_st_pagto_pago"
			value="<%=ST_PAGTO_PAGO%>"
			<% if ckb_st_pagto_pago <> "" then Response.Write " checked" %>
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_pago.click();">Pago</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_nao_pago" name="ckb_st_pagto_nao_pago"
			value="<%=ST_PAGTO_NAO_PAGO%>"
			<% if ckb_st_pagto_nao_pago <> "" then Response.Write " checked" %>
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_nao_pago.click();">Não-Pago</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_pago_parcial" name="ckb_st_pagto_pago_parcial"
			value="<%=ST_PAGTO_PARCIAL%>"
			<% if ckb_st_pagto_pago_parcial <> "" then Response.Write " checked" %>
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_pago_parcial.click();">Pago Parcial</span>
		</td>
	</tr>
	</table>
</td>
</tr>

<!--  NFS-e  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">NFS-e</span>
	<br />
	<table cellspacing="3" cellpadding="0" style="margin-bottom:4px; width:100%">
	<tr bgcolor="#FFFFFF">
		<td align="right" style="width: 20px"><span class="PLTd" style="margin-left:20px;">CNPJ do Emitente da NFS-e:</span></td>
		<td align="left"><span class="C"><%=cnpj_cpf_formata(c_cnpj_nfse)%></span></td>
	</tr>
	</table>
</td>
</tr>

<!--  INDICADOR(ES) SELECIONADO(S)  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">INDICADOR(ES) SELECIONADO(S)</span>
	<br />
	<table cellspacing="0" cellpadding="2" style="padding:6px 12px 10px 12px; width:100%">
	<%	s_sql = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (Id IN (" & ckb_id_indicador & ")) ORDER BY apelido"
		set rs = cn.Execute(s_sql)
		if rs.Eof then %>
	<tr bgcolor="#FFFFFF">
		<td class="MT ALERTA" align="center" valign="middle"><span class="ALERTA">NENHUM INDICADOR ENCONTRADO</span></td>
	</tr>
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td class="MT TdApelido" valign='bottom'><span class="R TdSpn">Apelido</span></td>
		<td class="MTBD TdCnpjCadastro" valign='bottom'><span class="R TdSpn">CNPJ/CPF (Cadastro)</span></td>
		<td class="MTBD TdRazaoSocialCadastro" valign='bottom'><span class="R TdSpn">Razão Social (Cadastro)</span></td>
	</tr>
	<%		do while Not rs.Eof
				qtde_indicadores_encontrados = qtde_indicadores_encontrados + 1
	%>
	<tr bgcolor="#FFFFFF">
		<td class="MDBE TdApelido"><span class="Cn TdSpn"><%=Trim("" & rs("apelido"))%></span></td>
		<td class="MDB TdCnpjCadastro"><span class="Cn TdSpn"><%=cnpj_cpf_formata(Trim("" & rs("cnpj_cpf")))%></span></td>
		<td class="MDB TdRazaoSocialCadastro"><span class="Cn TdSpn"><%=Trim("" & rs("razao_social_nome_iniciais_em_maiusculas"))%></span></td>
	</tr>
	<%
				rs.MoveNext
				loop
			end if
	%>
	</table>
</td>
</tr>

<!--  VISÃO: SINTÉTICA/ANALÍTICA  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">VISÃO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:4px;">
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="radio" tabindex="-1" id="rb_visao" name="rb_visao"
			value="ANALITICA"
			<% if (rb_visao = "ANALITICA") OR (rb_visao = "") then Response.Write " checked" %>
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_visao[0].click();">Analítica</span>
		    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="radio" tabindex="-1" id="rb_visao" name="rb_visao"
			value="SINTETICA"
			<% if (rb_visao = "SINTETICA") then Response.Write " checked" %>
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_visao[1].click();">Sintética</span>
		</td>
	</tr>
	</table>
</td>
</tr>

</table>

<!-- ************   SEPARADOR   ************ -->
<table width="690" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table width="690" cellspacing="0">
<% if qtde_indicadores_encontrados > 0 then %>
<tr>
	<td align="left"><a name="bANTERIOR" id="bANTERIOR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fRetornaInicio(fFILTRO)" title="retorna para o início do relatório">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="executa a consulta">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% else %>
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
<% end if %>
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