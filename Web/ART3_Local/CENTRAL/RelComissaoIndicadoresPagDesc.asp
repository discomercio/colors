<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L C O M I S S A O I N D I C A D O R E S P A G D E S C . A S P
'     =================================================================
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

    if Not operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
        Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
    end if

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

function mes_monta_itens_select()
dim i, x, m
    
    m = DateAdd("m", -1, Date)
    m = Month(m)
    x = ""

    for i=1 to 12
        x = x & "<option value='" & i & "'"
        if i = m then x = x & " selected"
        x = x & ">" & i & "</option>" & chr(13)
    next
    mes_monta_itens_select = x

end function

function ano_monta_itens_select()
dim i, x, a, aa
    a = Year(Date)
    aa = DateAdd("m", -1, Date)
    aa = Year(aa)

    x = ""
    for i=a to 2014 step -1
        x = x & "<option "
        if i = aa then x = x & "selected "
        x = x & "value='" & i & "'>" & i & "</option>" & chr(13)
    next
    ano_monta_itens_select = x
end function

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
    <meta charset="utf-8" />
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
var data;

data = new Date();
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

	if (f.c_dt_entregue_mes.value >= (data.getMonth()+1) && f.c_dt_entregue_ano.value == data.getFullYear()) {
	    alert("O mês de competência deve ser inferior ao mês atual!");
	    f.c_dt_entregue_mes.focus();
	    return;
	}

	if (f.c_vendedor.value == "") {
	    alert("Selecione pelo menos 1 (um) vendedor da lista!!");
	    return;
	}
        
	dCONFIRMA.style.visibility = "hidden";

	$("#c_vendedor").children().prop('selected', true);
		
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
<script type="text/javascript">
    function CarregaListaVendedores(a, m) {
        var strUrl, xmlhttp;
        xmlhttp = GetXmlHttpObject();
        if (xmlhttp == null) {
            alert("O browser NÃO possui suporte ao AJAX!!");
            return;
        }

        window.status = "Aguarde, pesquisando vendedores de  " + m + "/" + a + " ...";
        divMsgAguardeObtendoDados.style.visibility = "";

        strUrl = "../Global/AjaxRelComissaoIndicadoresListaVendedores.asp";
        strUrl = strUrl + "?ano=" + a;
        strUrl = strUrl + "&mes=" + m;
        strUrl = strUrl + "&id=" + Math.random();
        xmlhttp.onreadystatechange = function () {
            var strResp;

            if (xmlhttp.readyState == 4) {
                strResp = xmlhttp.responseText;
                if (strResp == "") {
                    $('#spn_aviso').css('display', 'block');
                    $("#c_vendedor_escolher").children().empty();
                    divMsgAguardeObtendoDados.style.visibility = "hidden";
                }
                if (strResp != "") {
                    try {
                        $('#c_vendedor_escolher').html(xmlhttp.responseText);
                        $('#spn_aviso').css('display', 'none');
                        window.status = "Concluído"
                        divMsgAguardeObtendoDados.style.visibility = "hidden";
                    }
                    catch (e) {
                        alert("Falha na consulta!!");
                    }
                }
            }
        }
        xmlhttp.open("GET", strUrl, true);
        xmlhttp.send();
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
<style type="text/css">
 
 #spn_aviso {
    display: none;
 }

</style>

<body onload="if (trim(fFILTRO.c_dt_entregue_mes.value)!='' && trim(fFILTRO.c_dt_entregue_ano.value)!='') { CarregaListaVendedores(fFILTRO.c_dt_entregue_ano.value, fFILTRO.c_dt_entregue_mes.value);}">
<center>
<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>
<form id="fFILTRO" name="fFILTRO" method="post" action="RelComissaoIndicadoresPagDescExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" name="c_hidden_indice_indicador" id="c_hidden_indice_indicador" value="" />
<input type="hidden" name="ckb_comissao_paga_nao" id="ckb_comissao_paga_nao" value="ON" />
<input type="hidden" name="ckb_st_pagto_pago" id="ckb_st_pagto_pago" value="<%=ST_PAGTO_PAGO%>" />
<input type="hidden" name="mes_reload" id="mes_reload" value="" />
<input type="hidden" name="ano_reload" id="ano_reload" value="" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="690" cellpadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores Com Desconto (Processamento)</span>
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
			<span class="C">Mês de competência:</span>
			 <select class="Cc" style="width:40px;" name="c_dt_entregue_mes" id="c_dt_entregue_mes" onchange="if (trim(this.value)!='' && trim(c_dt_entregue_ano.value)!='') { CarregaListaVendedores(c_dt_entregue_ano.value, this.value);}">
                <%=mes_monta_itens_select%>
             </select>
            <span class="C">/</span>
            <select class="Cc" style="width:50px;" name="c_dt_entregue_ano" id="c_dt_entregue_ano" onchange="if (trim(this.value)!='' && trim(c_dt_entregue_mes.value)!='') { CarregaListaVendedores(this.value, c_dt_entregue_mes.value);}">
                <%=ano_monta_itens_select%>
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
            <span class="C" style="cursor:default">Não-Paga</span>
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
		<input type="checkbox" tabindex="-1" id="st_pagto_pago" name="st_pagto_pago" value="<%=ST_PAGTO_PAGO%>" checked disabled="disabled"	/>
            <span class="C" style="cursor:default">Pago</span>
		</td>
	</tr>
	</table>
</td>
</tr>

<!--  CADASTRAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">CADASTRAMENTO</span>
	<br>
	<table cellspacing="3" cellpadding="0" style="margin-bottom:10px; width: 100%">
	<tr bgcolor="#FFFFFF">
		<td align="left" style="width:47%;"><span class="C" style="margin-left:0px;">Selecione o(s) vendedor(es)</span></td>
        <td style="width:6%;">&nbsp;</td>
        <td align="left" style="width:47%"><span class="C" style="margin-left:0px;">Vendedor(es) selecionado(s)</span></td>
    </tr>
    <tr>
		<td align="left">
			<select id="c_vendedor_escolher" name="c_vendedor_escolher" style="width:95%" size="10" multiple>
			
			</select>
            <br />
            <span class="C" id="spn_aviso" style="color:red;width:100%">Nenhum vendedor a ser processado.</span><span class="C">&nbsp;&nbsp;</span>
		</td>
        <td>
            <input type="button" id="btnAdiciona" value="&raquo;" />
            <br />
            <input type="button" id="btnRemove" value="&laquo;" />
        </td>
        <td align="left">
			<select id="c_vendedor" name="c_vendedor" style="width:95%" size="10" multiple>
			</select>
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