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

	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_script, strSql
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	if operacao_permitida(OP_LJA_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if


 dim lst_indicadores_carrega
    lst_indicadores_carrega = Request.Form("ckb_carrega_indicadores_rel")

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
	strSql = "SELECT DISTINCT t_USUARIO.usuario, nome, nome_iniciais_em_maiusculas FROM" & _
			 " t_USUARIO INNER JOIN t_USUARIO_X_LOJA ON t_USUARIO.usuario=t_USUARIO_X_LOJA.usuario" & _
			 " WHERE (vendedor_loja <> 0) AND " & _
			 SCHEMA_BD & ".UsuarioPossuiAcessoLoja(t_USUARIO.usuario, '" & loja & "') = 'S'" & _
			 " ORDER BY t_USUARIO.usuario"
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
' ____________________________________________________________________________
' INDICADORES DESTA LOJA MONTA ITENS SELECT
' LEMBRE-SE: O ORÇAMENTISTA É CONSIDERADO AUTOMATICAMENTE UM INDICADOR!!
function indicadores_desta_loja_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql="SELECT" & _
				" apelido," & _
				" razao_social_nome_iniciais_em_maiusculas" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR" & _
			" WHERE " & _
				"(" & _
					"(loja = '" & loja & "')" & _
					" OR " & _
					"(vendedor IN " & _
						"(" & _
							"SELECT DISTINCT " & _
								"usuario" & _
							" FROM t_USUARIO_X_LOJA" & _
							" WHERE" & _
								" (loja = '" & loja & "')" & _
						")" & _
					")" & _
				")"
'	SE HÁ RESTRIÇÃO NO PERÍODO DE CONSULTA, ENTÃO TRATA-SE DE UM VENDEDOR
	if operacao_permitida(OP_LJA_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		strSql = strSql & _
				" AND (vendedor = '" & usuario & "')"
		end if
	
	strSql = strSql & _
			" ORDER BY apelido"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("apelido")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("razao_social_nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
	
	indicadores_desta_loja_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' GRAVA ÚLTIMA OPÇÃO DE CONSULTA NO BD

call set_default_valor_texto_bd(usuario, "RelComissaoIndicadores|c_carrega_indicadores_estatico", lst_indicadores_carrega)

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

	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";

	try {
		if (('localStorage' in window) && window['localStorage'] !== null) {
			var d = $("#c_indicador").html();
			localStorage.setItem('c_indicador', d);
		}
	}
	catch (e) {
		// NOP
	}

	fFILTRO.c_hidden_reload.value = 1
	fFILTRO.c_hidden_indice_indicador.value = $("#c_indicador option:selected").index();
	
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
    oOption.innerText = "                                        ";
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

<% if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then %>

<script type="text/javascript">
    $(function() {

    $("#c_dt_entregue_inicio").hUtilUI('datepicker_filtro_inicial');
    $("#c_dt_entregue_termino").hUtilUI('datepicker_filtro_final');

    $(".aviso").css('display', 'none');
    <% if lst_indicadores_carrega = "" then %>
    $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50');    

		try {
			if (fFILTRO.c_hidden_reload.value == 1) {
				if (('localStorage' in window) && window['localStorage'] !== null) {
					if ('c_indicador' in localStorage) {
						$("#c_indicador").html(localStorage.getItem('c_indicador'));
						$("#c_indicador").prop('selectedIndex', fFILTRO.c_hidden_indice_indicador.value);
					}
				}
			}
		}
		catch (e) {
			// NOP
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
        <% end if %>

    });
</script>

<%else %>

<script type="text/javascript">
    $(function() {
    var usuario = "<%=usuario %>";

    $("#c_dt_entregue_inicio").hUtilUI('datepicker_filtro_inicial');
    $("#c_dt_entregue_termino").hUtilUI('datepicker_filtro_final');

        $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50');
        $(".aviso").css('display', 'none');

		if (fFILTRO.c_hidden_reload.value == 1) {
			try {
				if (('localStorage' in window) && window['localStorage'] !== null) {
					if ('c_indicador' in localStorage) {
						$("#c_indicador").html(localStorage.getItem('c_indicador'));
						$("#c_indicador").prop('selectedIndex', fFILTRO.c_hidden_indice_indicador.value);
					}
				}
			}
			catch (e) {
				// NOP
			}
		}
		else {
			CarregaListaIndicadores(usuario);
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

<%end if %>



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

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.5;visibility:hidden;vertical-align: middle">

	</div>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelComissaoIndicadoresExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" name="c_hidden_indice_indicador" id="c_hidden_indice_indicador" value="" />
<input type="hidden" name="c_hidden_reload" id="c_hidden_reload" value="0" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="690" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
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
<table width="690" class="Qx" cellSpacing="0">
<!--  STATUS DE ENTREGA  -->
<tr bgColor="#FFFFFF">
<td class="MT" NOWRAP><span class="PLTe">PERÍODO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_entregue" name="ckb_st_entrega_entregue" onclick="if (fFILTRO.ckb_st_entrega_entregue.checked) fFILTRO.c_dt_entregue_inicio.focus();"
			value="<%=ST_ENTREGA_ENTREGUE%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_entregue.click();">Pedidos entregues entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entregue_termino.focus(); else fFILTRO.ckb_st_entrega_entregue.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_entregue.checked=true;" onchange="if (trim(this.value)!='') fFILTRO.ckb_st_entrega_entregue.checked=true;"
			>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entregue_termino" id="c_dt_entregue_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_st_entrega_entregue.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_entregue.checked=true;" onchange="if (trim(this.value)!='') fFILTRO.ckb_st_entrega_entregue.checked=true;">
		</td></tr>
	</table>
</td></tr>

<!--  COMISSÃO PAGA  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">COMISSÃO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_comissao_paga_sim" name="ckb_comissao_paga_sim"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_comissao_paga_sim.click();">Paga</span>
		</td></tr>
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_comissao_paga_nao" name="ckb_comissao_paga_nao"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_comissao_paga_nao.click();">Não-Paga</span>
		</td></tr>
	</table>
</td></tr>

<!--  STATUS DE PAGAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">STATUS DE PAGAMENTO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_pago" name="ckb_st_pagto_pago"
			value="<%=ST_PAGTO_PAGO%>"
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_pago.click();">Pago</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_nao_pago" name="ckb_st_pagto_nao_pago"
			value="<%=ST_PAGTO_NAO_PAGO%>"
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_nao_pago.click();">Não-Pago</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_pago_parcial" name="ckb_st_pagto_pago_parcial"
			value="<%=ST_PAGTO_PARCIAL%>"
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_pago_parcial.click();">Pago Parcial</span>
		</td>
	</tr>
	</table>
</td>
</tr>

<!--  PAGAMENTO DA COMISSÃO VIA CARTÃO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">COMISSÃO VIA CARTÃO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:4px;">
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="radio" tabindex="-1" name="rb_pagto_comissao_via_cartao" id="rb_pagto_comissao_via_cartao_nao"
			value="0"
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_pagto_comissao_via_cartao[0].click();">Não</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="radio" tabindex="-1" name="rb_pagto_comissao_via_cartao" id="rb_pagto_comissao_via_cartao_sim"
			value="1"
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_pagto_comissao_via_cartao[1].click();">Sim</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="radio" tabindex="-1" name="rb_pagto_comissao_via_cartao" id="rb_pagto_comissao_via_cartao_todos"
			value="" checked
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_pagto_comissao_via_cartao[2].click();">Todos</span>
		</td>
	</tr>
	</table>
</td>
</tr>

<!--  CADASTRAMENTO  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">CADASTRAMENTO</span>
	<br>
	<table cellSpacing="6" cellPadding="0" style="margin-bottom:0px;">
	<% if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then %>
	<tr bgColor="#FFFFFF">
		<td align="right"><span class="C" style="margin-left:20px;">Vendedor</span></td>
		<td>
			<select id="c_vendedor" name="c_vendedor" style="margin-right:10px;" <% if lst_indicadores_carrega = "" then %> onchange="CarregaListaIndicadores(this.value);" <% end if %>>
			<option <% if lst_indicadores_carrega = "" then %>onclick="CarregaListaIndicadores(this.value)" <% end if %>value="">&nbsp;</option>
			<% =vendedores_monta_itens_select(Null) %>
			</select></td><td align="left" style="width: 300px"><% if lst_indicadores_carrega = "" then %><a href="javascript:CarregaListaIndicadores(fFILTRO.c_vendedor.value)"><img style="margin:0;border:0;" src="../IMAGEM/lupa_20x20.jpg" /></a><% end if %>
		</td>
		</td>
	</tr>
	<% else %>
	<input type="hidden" name="c_vendedor" id="c_vendedor" value=''>
	<% end if %>
	<tr bgColor="#FFFFFF">
		<td align="right" valign="top"><span class="C" style="margin-left:20px;">Indicador</span></td>
		<td colspan="2">
			<select id="c_indicador" name="c_indicador" style="margin-right:10px;width:321px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% if lst_indicadores_carrega <> "" then
			Response.Write indicadores_desta_loja_monta_itens_select(null) 
			end if%>
			</select><br />
			<span class="aviso">Vendedor selecionado não possui indicadores.</span>&nbsp;
			&nbsp;
		</td>
	</tr>
	</table>
</td>
</tr>
<% if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then %>
<tr>
    <td class="MDBE" nowrap>
        <span class="PLTe">VISÃO</span>
        <br />
        <table cellpadding="0" cellspacing="0" style="margin-top: 5px; margin-bottom: 5px">
            <tr>
                <td align="left">
                    <input type="radio" tabindex="-1" id="rb_visao" name="rb_visao" value="ANALITICA" checked="checked" />
                    <span class="C" style="cursor:default" onclick="fFILTRO.rb_visao[0].click();">Analítica</span>
                </td>
            </tr>
            <tr>
                <td align="left">
                <input type="radio" tabindex="-1" id="rb_visao" name="rb_visao" value="SINTETICA" />
                <span class="C" style="cursor:default" onclick="fFILTRO.rb_visao[1].click();">Sintética</span>
                </td>
            </tr>
        </table>

    </td>
</tr>
<% else %>
<input type="hidden" id="rb_visao" name="rb_visao" value="ANALITICA" />
<% end if %>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="690" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="690" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
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