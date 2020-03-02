<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================================
'	  RelFreteSinteticoFiltro.asp
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

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	' PREENCHIMENTO DA LISTA DE INDICADORES: GRAVA ÚLTIMA OPÇÃO DE CONSULTA NO BD
	dim lst_indicadores_carrega
	lst_indicadores_carrega = Request.Form("ckb_rel_frete_sint_carrega_indicadores")
	call set_default_valor_texto_bd(usuario, "RelFreteSintetico|c_carrega_indicadores_estatico", lst_indicadores_carrega)

	dim intIdx
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





' _____________________________________________
' TIPO_FRETE_MONTA_ITENS_SELECT
'
function tipo_frete_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE grupo='Pedido_TipoFrete' AND st_inativo=0")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("codigo"))
		if (id_default=x) then
			strResp = strResp & "<option selected"
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
	
    if id_default = "" Or id_default = null then strResp = "<option selected value=''>&nbsp;</option>" & strResp   	

	tipo_frete_monta_itens_select = strResp
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
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
    $(function () {

	<% if lst_indicadores_carrega = "" then %>

            $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');

        if (fFILTRO.c_hidden_reload.value == 1) {
            if (('localStorage' in window) && window['localStorage'] !== null) {
                if ('c_indicador' in localStorage) {
                    $("#c_indicador").html(localStorage.getItem('c_indicador'));
                    $("#c_indicador").prop('selectedIndex', fFILTRO.c_hidden_indice_indicador.value);
                }
            }
        }
        
     <% end if %>

        $("#c_dt_entregue_inicio").hUtilUI('datepicker_filtro_inicial');
        $("#c_dt_entregue_termino").hUtilUI('datepicker_filtro_final');

        //Every resize of window
        $(window).resize(function () {
            sizeDivAjaxRunning();
        });

        //Every scroll of window
        $(window).scroll(function () {
            sizeDivAjaxRunning();
        });

        //Dynamically assign height
        function sizeDivAjaxRunning() {
            var newTop = $(window).scrollTop() + "px";
            $("#divMsgAguardeObtendoDados").css("top", newTop);
        }
        $(document).tooltip();
    });
</script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
var s_de, s_ate;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;
var i;

//  PERÍODO DE ENTREGA
	if (trim(f.c_dt_entregue_inicio.value)=="") {
		alert("Informe a data inicial do período de entrega!!");
		f.c_dt_entregue_inicio.focus();
		return;
		}
	
	if (trim(f.c_dt_entregue_termino.value)=="") {
		alert("Informe a data final do período de entrega!!");
		f.c_dt_entregue_termino.focus();
		return;
		}
		
	if (trim(f.c_dt_entregue_inicio.value)!="") {
		if (!isDate(f.c_dt_entregue_inicio)) {
			alert("Data inválida!!");
			f.c_dt_entregue_inicio.focus();
			return;
			}
		}

	if (trim(f.c_dt_entregue_termino.value)!="") {
		if (!isDate(f.c_dt_entregue_termino)) {
			alert("Data inválida!!");
			f.c_dt_entregue_termino.focus();
			return;
			}
		}

	s_de = trim(f.c_dt_entregue_inicio.value);
	s_ate = trim(f.c_dt_entregue_termino.value);
	if ((s_de!="")&&(s_ate!="")) {
		s_de=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
		s_ate=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
		if (s_de > s_ate) {
			alert("Data de término é menor que a data de início!!");
			f.c_dt_entregue_termino.focus();
			return;
			}
		}

//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
	// PERÍODO DE ENTREGA
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
	window.status = "Aguarde ...";

	if (f.rb_tipo_saida[1].checked) setTimeout('exibe_botao_confirmar()', 15000);

    <% if lst_indicadores_carrega = "" then %>
	    if (('localStorage' in window) && window['localStorage'] !== null) {
        var d = $("#c_indicador").html();
        localStorage.setItem('c_indicador', d);
    }
	<% end if %>

        fFILTRO.c_hidden_reload.value = 1;
    fFILTRO.c_hidden_indice_indicador.value = $("#c_indicador option:selected").index();
    fFILTRO.ultimoVendedor.value = fFILTRO.c_vendedor.value;

    f.c_FormFieldValues.value = formToString($("#fFILTRO"));

	f.submit();
}

function exibe_botao_confirmar() {
	dCONFIRMA.style.visibility = "";
	window.status = "";
}
</script>

<script type="text/javascript">

    function LimpaListaIndicadores() {
        var f, oOption;
        f = fFILTRO;
        $("#c_indicador").empty();
        $(".aviso").css('display', 'none');

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

    function CarregaListaIndicadores() {
        var f, strUrl;
        f = fFILTRO;
        if (fFILTRO.ultimoVendedor.value == trim(fFILTRO.c_vendedor.value)) {
            return;
        }
        objAjaxListaIndicadores = GetXmlHttpObject();
        if (objAjaxListaIndicadores == null) {
            alert("O browser NÃO possui suporte ao AJAX!!");
            return;
        }

        //  Limpa lista de Indicadores
        LimpaListaIndicadores();
        divMsgAguardeObtendoDados.style.visibility = "";

        strUrl = "../Global/AjaxListaIndicadoresLojaPesqBD.asp?";
        //  Prevents server from using a cached file
        strUrl = strUrl + "sid=" + Math.random() + Math.random();
        if (trim(fFILTRO.c_vendedor.value) != "") {
            strUrl = strUrl + "&vendedor=" + fFILTRO.c_vendedor.value;
        }
        fFILTRO.ultimoVendedor.value = fFILTRO.c_vendedor.value;
        objAjaxListaIndicadores.onreadystatechange = TrataRespostaAjaxListaIndicadores;
        objAjaxListaIndicadores.open("GET", strUrl, true);
        objAjaxListaIndicadores.send(null);
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">


<body onload="fFILTRO.c_dt_entregue_inicio.focus();">
<center>

<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity: 0.6;visibility:hidden;vertical-align: middle">

	</div>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelFreteSinteticoExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />
<input type="hidden" id="ultimoVendedor" name="ultimoVendedor" value="x-x-x-x-x-x" />
<input type="hidden" name="c_hidden_reload" id="c_hidden_reload" value="0" />
<input type="hidden" name="c_hidden_indice_indicador" id="c_hidden_indice_indicador" value="" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Frete (Sintético)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0">
<!--  ENTREGUE ENTRE  -->
	<tr bgColor="#FFFFFF">
	<td class="MT" NOWRAP>
		<table cellSpacing="0" cellPadding="0"><tr bgColor="#FFFFFF"><td>
		<span class="PLTe" style="cursor:default">ENTREGUES ENTRE</span>
		<br>
		<input class="PLLc" maxlength="10" style="width:100px;" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio"
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entregue_termino.focus(); filtra_data();"
			>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;até&nbsp;</span>&nbsp;
			<input class="PLLc" maxlength="10" style="width:100px;" name="c_dt_entregue_termino" id="c_dt_entregue_termino"
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_transportadora.focus(); filtra_data();" 
			>
			</td></tr>
		</table>
		</td></tr>

<!--  TRANSPORTADORA  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe">TRANSPORTADORA</span>
		<br>
			<select id="c_transportadora" name="c_transportadora" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =transportadora_monta_itens_select(Null) %>
			</select>
			</td></tr>

<!--  TIPO DE FRETE  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" align="left" NOWRAP><span class="PLTe">TIPO DE FRETE</span>
		<br>
			<select id="c_tipo_frete" name="c_tipo_frete" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=tipo_frete_monta_itens_select(Null) %>
			</select>
			</td></tr>

<!--  FABRICANTE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">FABRICANTE</span>
	<br>
		<input maxlength="4" class="PLLe" style="width:150px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_loja.focus(); filtra_fabricante();">
		</td></tr>

<!--  LOJA  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">LOJA</span>
	<br>
		<input class="PLLe" maxlength="3" style="width:150px;" name="c_loja" id="c_loja" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true)) fFILTRO.c_vendedor.focus(); filtra_numerico();">
		</td></tr>

<!--  VENDEDOR  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">VENDEDOR</span>
	<br>
		<select id="c_vendedor" name="c_vendedor" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" <% if lst_indicadores_carrega = "" then %>onchange="LimpaListaIndicadores()" <% end if %>>
		<% =vendedores_monta_itens_select(Null) %>
		</select>
		</td></tr>

<!--  INDICADOR  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe"><% if lst_indicadores_carrega = "" then %><img id="exclamacao" src="../IMAGEM/exclamacao_14x14.png" title="Reduza o tempo de carregamento da lista de indicadores, filtrando por vendedor." style="cursor:pointer;" />&nbsp;<% end if %>INDICADOR</span>
		<br>
			<select id="c_indicador" name="c_indicador" style="margin:1px 10px 6px 10px;min-width:600px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" <% if lst_indicadores_carrega = "" then %> onfocus="CarregaListaIndicadores();" <% end if %>>
			    <option selected value=''>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
			<% if lst_indicadores_carrega <> "" then
				Response.Write indicadores_monta_itens_select(Null)
				end if%>
			</select>
			</td></tr>

<!--  UF  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe">UF</span>
		<br>
			<select id="c_uf" name="c_uf" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =uf_monta_itens_select(Null) %>
			</select>
			</td></tr>
			
<!--  STATUS DO FRETE  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe">STATUS DO FRETE</span>
		<br>
			<% intIdx=-1 %>
			<input type="radio" id="rb_frete_status" name="rb_frete_status" value="0" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_frete_status[<%=Cstr(intIdx)%>].click();">Frete <b style="color:red;">não</b> preenchido</span>
			<br>
			<input type="radio" id="rb_frete_status" name="rb_frete_status" value="1" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_frete_status[<%=Cstr(intIdx)%>].click();">Frete <b style="color:green;">já</b> preenchido</span>
			<br>
			<input type="radio" id="rb_frete_status" name="rb_frete_status" checked value="" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_frete_status[<%=Cstr(intIdx)%>].click();">Ambos</span>
			</td></tr>
<!--  SAÍDA DO RELATÓRIO  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe">SAÍDA DO RELATÓRIO</span>
		<br>
			<% intIdx=-1 %>
			<input type="radio" id="rb_tipo_saida" name="rb_tipo_saida" value="HTML" class="CBOX" style="margin-left:20px;" checked>
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_tipo_saida[<%=Cstr(intIdx)%>].click();">Html</span>
			<br />
			<input type="radio" id="rb_tipo_saida" name="rb_tipo_saida" value="XLS" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_tipo_saida[<%=Cstr(intIdx)%>].click();">Excel</span>
			</td></tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
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
