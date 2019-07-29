<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================================
'	  R E L C O M I S S A O I N D I C A D O R E S C O N S U L T A P A G O S . A S P
'     =============================================================================
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

	dim s_script, strSql
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))


'	FILTROS
	dim ckb_st_entrega_entregue, c_dt_entregue_mes, c_dt_entregue_ano
	dim ckb_comissao_paga_sim, ckb_comissao_paga_nao
	dim ckb_st_pagto_pago, ckb_st_pagto_nao_pago, ckb_st_pagto_pago_parcial
	dim c_vendedor, c_indicador
	dim c_loja, rb_visao
	
	ckb_st_entrega_entregue = Trim(Request.Form("ckb_st_entrega_entregue"))

	c_vendedor = Trim(Request.Form("c_vendedor"))
	c_indicador = Trim(Request.Form("c_indicador"))

	ckb_comissao_paga_sim = Trim(Request.Form("ckb_comissao_paga_sim"))
	ckb_comissao_paga_nao = Trim(Request.Form("ckb_comissao_paga_nao"))

	ckb_st_pagto_pago = Trim(Request.Form("ckb_st_pagto_pago"))
	ckb_st_pagto_nao_pago = Trim(Request.Form("ckb_st_pagto_nao_pago"))
	ckb_st_pagto_pago_parcial = Trim(Request.Form("ckb_st_pagto_pago_parcial"))

	c_loja = Trim(Request.Form("c_loja"))
	rb_visao = Trim(Request.Form("rb_visao"))
	
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

' ____________________________________________________________________________
' VENDEDORES MONTA ITENS SELECT
'


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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var s_ult_vendedor_selecionado = "--XX--XX--XX--XX--XX--";

function fFILTROConfirma( f ) {
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;

	if (f.c_dt_entregue_mes.value != "" || f.c_dt_entregue_ano.value != "") {
		if (f.c_dt_entregue_mes.value == "") {
		    alert("Selecione o mês de competência!");
		    f.c_dt_entregue_mes.focus();
		    return;
		}
		if (f.c_dt_entregue_ano.value == "") {
		    alert("Selecione o ano referente ao mês de competência!");
		    f.c_dt_entregue_ano.focus();
		    return;
		}
	}
	else {
		alert("Selecione o mês de competência !");
		f.c_dt_entregue_mes.focus();
		return;
	}

	$("#c_vendedor").children().prop('selected', true);
	bCONFIRMA.style.visibility = "hidden";
		
	window.status = "Aguarde ...";
	f.submit();
}

function ind_new(vendedor, apelido, nome) {
	this.vendedor = vendedor;
	this.apelido = apelido;
	this.nome = nome;
	return this;
}

</script>

<script type="text/javascript">
    $(function () {
        var data, ano, i, opt;

        $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');

        data = new Date();
        ano = data.getFullYear();

        for (i = 1; i <= 12; i++) {
            opt = document.createElement("option");
            fFILTRO.c_dt_entregue_mes.options.add(opt);
            opt.innerText = i;
            opt.value = i;

        }
        for (i = ano; i >= 2000; i--) {
            opt = document.createElement("option");
            fFILTRO.c_dt_entregue_ano.options.add(opt);
            opt.innerText = i;
            opt.value = i;
        }

        $("#btnAdiciona").click(function () {
            var x = $("#c_vendedor_escolher option:selected");
            $("#c_vendedor").append(x);
            reOrdenarEscolhidos();
        });

        $("#btnRemove").click(function () {
            var x = $("#c_vendedor option:selected");
            $("#c_vendedor_escolher").append(x);
            reOrdenarAEscolher();
        });

        $("#c_vendedor_escolher").dblclick(function () {
            var x = $("#c_vendedor_escolher option:selected");
            $("#c_vendedor").append(x);
            reOrdenarEscolhidos();
        });

        $("#c_vendedor").dblclick(function () {
            var x = $("#c_vendedor option:selected");
            $("#c_vendedor_escolher").append(x);
            reOrdenarAEscolher();
        });
                    
    });

    function reOrdenarAEscolher() {
        $("#c_vendedor_escolher").html($("#c_vendedor_escolher option").sort(function (a, b) {
            return a.text.toUpperCase() == b.text.toUpperCase() ? 0 : a.text.toUpperCase() < b.text.toUpperCase() ? -1 : 1
        }))
    }

    function reOrdenarEscolhidos() {
        $("#c_vendedorr").html($("#c_vendedor option").sort(function (a, b) {
            return a.text.toUpperCase() == b.text.toUpperCase() ? 0 : a.text.toUpperCase() < b.text.toUpperCase() ? -1 : 1
        }))
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
 
</style>


<body>
<center>
<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>
<form id="fFILTRO" name="fFILTRO" method="post" action="RelComissaoIndicadoresConsultaPagosExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" name="c_hidden_indice_indicador" id="c_hidden_indice_indicador" value="" />
<input type="hidden" name="c_hidden_reload" id="c_hidden_reload" value="0" />
<input type="hidden" name="ckb_comissao_paga_nao" id="ckb_comissao_paga_nao" value="ON" />
<input type="hidden" name="ckb_st_pagto_pago" id="ckb_st_pagto_pago" value="<%=ST_PAGTO_PAGO%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="690" cellpadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (Processado)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table width="690" class="Qx" cellspacing="0">
<!--  MÊS  -->
<tr bgcolor="#FFFFFF">
<td class="MT" align="left" nowrap><span class="PLTe">MÊS DE COMPETÊNCIA</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF">
		<td align="left">
	<!--	<input type="checkbox" tabindex="-1" id="ckb_st_entrega_entregue" name="ckb_st_entrega_entregue" onclick="if (fFILTRO.ckb_st_entrega_entregue.checked) fFILTRO.c_dt_entregue_mes.focus();"
			value="<%=ST_ENTREGA_ENTREGUE%>"
			>--><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_entregue.click();">Mês de competência:</span>
			 <select class="Cc" style="width:40px;" name="c_dt_entregue_mes" id="c_dt_entregue_mes" onchange="if (trim(this.value)!='') fFILTRO.ckb_st_entrega_entregue.checked=true;else fFILTRO.ckb_st_entrega_entregue.checked=false;" />
                <option value=""></option>
             </select>
            <span class="C">/</span>
            <select class="Cc" style="width:50px;" name="c_dt_entregue_ano" id="c_dt_entregue_ano" onchange="if (trim(this.value)!='') fFILTRO.ckb_st_entrega_entregue.checked=true;else fFILTRO.ckb_st_entrega_entregue.checked=false;" />
                <option value=""></option>
            </select>
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
		<input type="checkbox" tabindex="-1" id="comissao_paga_nao" name="comissao_paga_nao" value="ON" checked disabled="disabled" />
            <span class="C" style="cursor:default" onclick="fFILTRO.ckb_comissao_paga_nao.click();">Paga</span>
		</td>
	</tr>
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
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="executa a consulta">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>