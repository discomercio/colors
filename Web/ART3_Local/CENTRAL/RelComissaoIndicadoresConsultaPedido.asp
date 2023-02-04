<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp" -->

<%
'     ===============================================================================
'	  R E L C O M I S S A O I N D I C A D O R E S C O N S U L T A P E D I D O . A S P
'     ===============================================================================
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


function mes_monta_itens_select()
dim i, x

    x = "<option value=''>&nbsp;</option>"
    for i=1 to 12
        x = x & "<option value='" & i & "'>" & i & "</option>"
    next
    mes_monta_itens_select = x

end function

function ano_monta_itens_select()
dim i, x, a
    a = Year(Date)

    x = "<option value=''>&nbsp;</option>"
    for i=a to 2014 step -1
        x = x & "<option value='" & i & "'>" & i & "</option>"
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

$(function () {
	if (fFILTRO.c_hidden_reload.value == 1) {
		try {
			if (('localStorage' in window) && window['localStorage'] !== null) {
				if ('lista_id' in localStorage) {
					$("#id").html(localStorage.getItem('lista_id'));
				}
			}
		}
		catch (e) {
			// NOP
		}
	}

	// Trata o problema em que os campos do formulário são limpos após retornar à esta página c/ o history.back() pela 2ª vez quando ocorre erro de consistência
	if (trim(fFILTRO.c_FormFieldValues.value) != "") {
		stringToForm(fFILTRO.c_FormFieldValues.value, $('#fFILTRO'));
	}
});

function fFILTROConfirma(f) {
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
		alert("Selecione o mês de competência e o ano");
		f.c_dt_entregue_mes.focus();
		return;
	}

	if (f.id.value == "") {
		alert("Selecione um relatório para consulta!!");
		return;
		f.id.focus();
	}

	try {
		if (('localStorage' in window) && window['localStorage'] !== null) {
			var d = $("#id").html();
			localStorage.setItem('lista_id', d);
		}
	}
	catch (e) {
		// NOP
	}

	fFILTRO.c_hidden_reload.value = 1;
	fFILTRO.c_FormFieldValues.value = formToString($("#fFILTRO"));

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
                    
    });

</script>
<script type="text/javascript">
    function CarregaListaRelatorios(a,m) {
        var strUrl, xmlhttp;
        xmlhttp = GetXmlHttpObject();
        if (xmlhttp == null) {
            alert("O browser NÃO possui suporte ao AJAX!!");
            return;
        }
        
        window.status = "Aguarde, pesquisando relatórios de  " + m + "/" + a + " ...";
        divMsgAguardeObtendoDados.style.visibility = "";

        strUrl = "../Global/AjaxRelatorioComissaoIndicadores.asp";
        strUrl = strUrl + "?ano=" + a;
        strUrl = strUrl + "&mes=" + m;
        strUrl = strUrl + "&id=" + Math.random();
        xmlhttp.onreadystatechange = function () {
            var strResp;
            
            if (xmlhttp.readyState == 4) {
                strResp = xmlhttp.responseText;
                if (strResp == "") {
                    $('#spn_aviso').css('display', 'block');
                    $("#id").children().empty();
                    divMsgAguardeObtendoDados.style.visibility = "hidden";
                    window.status = "Concluído"
                }
                if (strResp != "") {
                    try {
                        $('#id').html(xmlhttp.responseText);
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
 
 #spn_aviso {
    display: none;
 }

</style>


<body>
<center>
<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>
<form id="fFILTRO" name="fFILTRO" method="get" action="RelComissaoIndicadoresConsultaPedidoExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" name="c_hidden_indice_indicador" id="c_hidden_indice_indicador" value="" />
<input type="hidden" name="c_hidden_reload" id="c_hidden_reload" value="0" />
<input type="hidden" name="ckb_comissao_paga_nao" id="ckb_comissao_paga_nao" value="ON" />
<input type="hidden" name="ckb_st_pagto_pago" id="ckb_st_pagto_pago" value="<%=ST_PAGTO_PAGO%>" />
<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="690" cellpadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (Consulta)</span>
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

			<span class="C" style="cursor:default" onclick="fFILTRO.ckb_st_entrega_entregue.click();">Mês de competência:</span>
			 <select class="Cc" style="width:40px;" name="c_dt_entregue_mes" id="c_dt_entregue_mes" onchange="if (trim(this.value)!='' && trim(c_dt_entregue_ano.value)!='') { CarregaListaRelatorios(c_dt_entregue_ano.value, this.value);}" />
                <%=mes_monta_itens_select%>
             </select>
            <span class="C">/</span>
            <select class="Cc" style="width:50px;" name="c_dt_entregue_ano" id="c_dt_entregue_ano" onchange="if (trim(this.value)!='' && trim(c_dt_entregue_mes.value)!='') { CarregaListaRelatorios(this.value, c_dt_entregue_mes.value);}" />
                <%=ano_monta_itens_select%>
            </select>
		</td>
	</tr>
	</table>
</td></tr>

<!--  CADASTRAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">RELATÓRIO</span>
	<br>
	<table cellspacing="3" cellpadding="0" style="margin-bottom:10px; width: 100%">
	<tr bgcolor="#FFFFFF">
		<td align="left" style="width:100%;"><span class="C" style="margin-left:0px;">Selecione o relatório</span></td>
    </tr>
    <tr>
		<td align="left">
			<select id="id" name="id" style="width:95%">
			
			</select>
            <br />
            <span class="C" id="spn_aviso" style="color:red;">Não há lançamentos para o mês de competência informado.</span><span class="C">&nbsp;</span>
		</td>
	</tr>
    <tr>
        <td>
            <input type="checkbox" class="rbOpt" id="ckb_Desc" name="ckb_Desc" value="1" ><label for="ckb_Desc"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_saida[0].click();dCONFIRMA.style.visibility='';"
			>Somente indicadores com descontos</span></label>
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