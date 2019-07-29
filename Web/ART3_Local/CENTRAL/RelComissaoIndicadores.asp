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
'	  R E L C O M I S S A O I N D I C A D O R E S . A S P
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

	dim s_script, strSql
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	FILTROS
	dim ckb_st_entrega_entregue, c_dt_entregue_inicio, c_dt_entregue_termino
	dim ckb_comissao_paga_sim, ckb_comissao_paga_nao
	dim ckb_st_pagto_pago, ckb_st_pagto_nao_pago, ckb_st_pagto_pago_parcial
	dim c_vendedor, c_indicador
	dim c_loja, rb_visao
	
	ckb_st_entrega_entregue = Trim(Request.Form("ckb_st_entrega_entregue"))
	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))

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
function vendedores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT usuario, nome, nome_iniciais_em_maiusculas FROM" & _
			 " (" & _
			 "SELECT usuario, nome, nome_iniciais_em_maiusculas FROM t_USUARIO" & _
				" WHERE (vendedor_loja <> 0)" & _
			 " UNION" & _
			 " SELECT t_USUARIO.usuario AS usuario, t_USUARIO.nome AS nome, t_USUARIO.nome_iniciais_em_maiusculas FROM t_USUARIO" & _
				" INNER JOIN t_PERFIL_X_USUARIO ON (t_USUARIO.usuario=t_PERFIL_X_USUARIO.usuario)" & _
				" INNER JOIN t_PERFIL ON (t_PERFIL_X_USUARIO.id_perfil=t_PERFIL.id)" & _
				" INNER JOIN t_PERFIL_ITEM ON (t_PERFIL.id=t_PERFIL_ITEM.id_perfil)" & _
				" WHERE (t_PERFIL_ITEM.id_operacao=" & OP_CEN_ACESSO_TODAS_LOJAS & ")" & _
			 ") AS t" & _
			 " ORDER BY usuario"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	
		
		
	vendedores_monta_itens_select = strResp
	r.close
	set r=nothing
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

	if (f.ckb_st_entrega_entregue.checked) {
		if (!consiste_periodo(f.c_dt_entregue_inicio, f.c_dt_entregue_termino)) return;
		}

//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
		strDtRefDDMMYYYY = trim(f.c_dt_entregue_inicio.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_entregue_termino.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		}

		dCONFIRMA.style.visibility = "hidden";

		
		    if (('localStorage' in window) && window['localStorage'] !== null) {
		        var d = $("#c_indicador").html();
		        localStorage.setItem('c_indicador', d);
		    }

		fFILTRO.c_hidden_reload.value = 1
		fFILTRO.c_hidden_indice_indicador.value = $("#c_indicador option:selected").index();
		
	window.status = "Aguarde ...";
	f.submit();
}

function ind_new(vendedor, apelido, nome) {
	this.vendedor = vendedor;
	this.apelido = apelido;
	this.nome = nome;
	return this;
}

function LimpaListaIndicadores() {
    var f, oOption;
    f = fFILTRO;
    $("#c_indicador").empty();

    //  Cria um item vazio
    oOption = document.createElement("OPTION");
    f.c_indicador.add(oOption);
    oOption.innerText = "                                                                                 ";
    oOption.value = "";
    oOption.selected = true;
}

function TrataRespostaAjaxListaIndicadores() {
    var f, i, strApelido, strNome, strResp, xmlDoc, oOption, oNodes;
    f = fFILTRO;
    if (objAjaxListaIndicadores.readyState == AJAX_REQUEST_IS_COMPLETE) {
        strResp = objAjaxListaIndicadores.responseText;
        if (strResp == "") {
            window.status = "Concluído";
            divMsgAguardeObtendoDados.style.visibility = "hidden";
                $(".aviso").css('display', 'inline');
            return;
        }

        if (strResp != "") {
            $(".aviso").css('display', 'none');
            try {
                xmlDoc = objAjaxListaIndicadores.responseXML.documentElement;
                for (i = 0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
                    oOption = document.createElement("OPTION");
                    f.c_indicador.options.add(oOption);

                    oNodes = xmlDoc.getElementsByTagName("apelido")[i];
                    if (oNodes.childNodes.length > 0) strApelido = oNodes.childNodes[0].nodeValue; else strApelido = "";
                    if (strApelido == null) strApelido = "";
                    oOption.value = strApelido;

                    oNodes = xmlDoc.getElementsByTagName("razao_social_nome")[i];
                    if (oNodes.childNodes.length > 0) strNome = oNodes.childNodes[0].nodeValue; else strNome = "";
                    if (strNome == null) strNome = "";

                    oOption.value = strApelido;
                    oOption.innerText = strApelido + " - " + strNome;
                }
            }
            catch (e) {
                alert("Falha na consulta de indicadores!!" + "\n" + e.description);
            }
        }
        window.status = "Concluído";
        divMsgAguardeObtendoDados.style.visibility = "hidden";

        
    }
}

function CarregaListaIndicadores(strVendedor) {
    var f, strUrl;
    f = fFILTRO;
    
    objAjaxListaIndicadores = GetXmlHttpObject();
    if (objAjaxListaIndicadores == null) {
        alert("O browser NÃO possui suporte ao AJAX!!");
        return;
    }

    //  Limpa lista de Indicadores
    LimpaListaIndicadores();

    window.status = "Aguarde, pesquisando os indicadores do vendedor " + strVendedor + " ...";
    divMsgAguardeObtendoDados.style.visibility = "";

    strUrl = "../Global/AjaxListaIndicadoresLojaPesqBD.asp";
    strUrl = strUrl + "?vendedor=" + strVendedor;
    //  Prevents server from using a cached file
    strUrl = strUrl + "&sid=" + Math.random() + Math.random();

    objAjaxListaIndicadores.onreadystatechange = TrataRespostaAjaxListaIndicadores;
    objAjaxListaIndicadores.open("GET", strUrl, true);
    objAjaxListaIndicadores.send(null);
}
</script>

<script type="text/javascript">
    $(function() {

    $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');
    
        $("#c_dt_entregue_inicio").hUtilUI('datepicker_filtro_inicial');
        $("#c_dt_entregue_termino").hUtilUI('datepicker_filtro_final');

        $(".aviso").css('display', 'none');
        if (fFILTRO.c_hidden_reload.value == 1) {
            if (('localStorage' in window) && window['localStorage'] !== null) {
                if ('c_indicador' in localStorage) {
                    $("#c_indicador").html(localStorage.getItem('c_indicador'));
                    $("#c_indicador").prop('selectedIndex', fFILTRO.c_hidden_indice_indicador.value);
                }
            }
        }
        //Every resize of window
        $(window).resize(function() {
            sizeDivAjaxRunning();
        });

        //Every scroll of window
        $(window).scroll(function() {
            sizeDivAjaxRunning();
        });

        //Dynamically assign height
        function sizeDivAjaxRunning() {
            var newTop = $(window).scrollTop() + "px";
            $("#divMsgAguardeObtendoDados").css("top", newTop);
        }
                    
    });

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
<form id="fFILTRO" name="fFILTRO" method="post" action="RelComissaoIndicadoresExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" name="c_hidden_indice_indicador" id="c_hidden_indice_indicador" value="" />
<input type="hidden" name="c_hidden_reload" id="c_hidden_reload" value="0" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="690" cellpadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores</span>
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

<!--  CADASTRAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">CADASTRAMENTO</span>
	<br>
	<table cellspacing="6" cellpadding="0" style="margin-bottom:10px; width: 100%">
	<tr bgcolor="#FFFFFF" style="width: 70px">
		<td align="right"><span class="C" style="margin-left:20px;">Vendedor</span></td>
		<td align="left">
			<select id="c_vendedor" name="c_vendedor" style="margin-right:10px;" onchange="CarregaListaIndicadores(this.value);">
			<option onclick="CarregaListaIndicadores(this.value)" value="">&nbsp;</option>
			<% =vendedores_monta_itens_select(c_vendedor) %>
			</select></td><td align="left" style="width: 300px"><a href="javascript:CarregaListaIndicadores(fFILTRO.c_vendedor.value)"><img style="margin:0;border:0;" src="../IMAGEM/lupa_20x20.jpg" /></a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="right" valign="top" style="width: 70px"><span class="C" style="margin-left:20px;">Indicador</span></td>
		<td align="left" colspan="2">
			<select id="c_indicador" name="c_indicador" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<option selected value=''>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
			</select><br />
			<span class="aviso">Vendedor selecionado não possui indicadores.</span>&nbsp;
		</td>
	</tr>
	</table>
</td>
</tr>

<!--  LOJA(S)  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">LOJA(S)</span>
<br>
	<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 30px;">
	<tr bgcolor="#FFFFFF">
		<td align="left">
			<textarea class="PLBe" style="width:100px;font-size:9pt;margin-bottom:4px;" rows="8" name="c_loja" id="c_loja" onkeypress="if (!digitou_enter(false) && !digitou_char('-')) filtra_numerico();" onblur="this.value=normaliza_lista_lojas(this.value);"><% if c_loja <> "" then Response.Write c_loja%></textarea>
		</td>
	</tr>
	</table>
</td></tr>

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
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
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