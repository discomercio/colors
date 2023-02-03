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
'	  R E L P E D I D O S M C R I T . A S P
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

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
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

	dim s_memoria

	dim url_origem
	url_origem = Trim(Request("url_origem"))

	dim url_back
	url_back = Trim(Request("url_back"))

	' PREENCHIMENTO DA LISTA DE INDICADORES: GRAVA ÚLTIMA OPÇÃO DE CONSULTA NO BD
    dim lst_indicadores_carrega
    lst_indicadores_carrega = Request.Form("ckb_carrega_indicadores")
	if url_origem = "" then
	' GRAVA ÚLTIMA OPÇÃO DE CONSULTA NO BD
	call set_default_valor_texto_bd(usuario, "RelPedidosMCrit|c_carrega_indicadores_estatico", lst_indicadores_carrega)
	end if

	' SE ESTA PÁGINA FOI ACIONADA COMO RETORNO DE OUTRA PÁGINA DECORRENTE DA CONSULTA DE UM PEDIDO DA LISTA DE RESULTADOS, RESTAURA OS FILTROS
	dim strJS, c_FormFieldValues
	strJS = ""
	c_FormFieldValues = ""
	if (url_origem <> "") Or (url_back <> "") then
		c_FormFieldValues = get_default_valor_texto_bd(usuario, "LOJA/RelPedidosMCrit|FormFields")
		if c_FormFieldValues <> "" then
			strJS = "	var formString = '" & c_FormFieldValues & "';" & chr(13) & _
					"	stringToForm(formString, $('#fFILTRO'));" & chr(13)
			end if
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________


' ____________________________________________________________________________
' VENDEDORES DESTA LOJA MONTA ITENS SELECT
'
function vendedores_desta_loja_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" usuario, nome_iniciais_em_maiusculas" & _
			" FROM t_USUARIO" & _
			" WHERE" & _
				" (vendedor_loja <> 0)"

'	TRATA-SE DE UM VENDEDOR NORMAL?
	if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
		strSql = strSql & _
				" AND (usuario = '" & usuario & "')"
		end if
	
	strSql = strSql & _
				" AND (" & _
					"usuario IN (" & _
						"SELECT DISTINCT" & _
							" usuario" & _
						" FROM t_USUARIO_X_LOJA" & _
						" WHERE" & _
							" (loja = '" & loja & "')" & _
						")" & _
					")" & _
			" ORDER BY" & _
				" usuario"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
	
	vendedores_desta_loja_monta_itens_select = strResp
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
'	TRATA-SE DE UM VENDEDOR NORMAL?
	if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
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

' ____________________________________________________________________________
' GRUPO PRODUTO MONTA ITENS SELECT
'
function t_produto_grupo_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, v, i
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "select distinct Coalesce(grupo, '') as codigo, tPG.descricao from t_produto tP" & _
                   " left join t_produto_grupo tPG on (tP.grupo=tPG.codigo)" & _
                   " where (Coalesce(grupo,'')<>'')" & _
                   " order by Coalesce(grupo,'')"

	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
	    
		x = Trim("" & r("codigo"))
		strResp = strResp & "<option "
            for i=LBound(v) to UBound(v) 
		        if (id_default<>"") And (v(i)=x) then
		            strResp = strResp & "selected"
		         end if
		   	next
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("codigo"))        
        if Trim("" & r("descricao")) <> "" then strResp = strResp & "&nbsp;-&nbsp;" & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext	
 	loop
		
	t_produto_grupo_monta_itens_select = strResp
	r.close
	set r=nothing
end function

'----------------------------------------------------------------------------------------------
' grupo_origem_pedido_monta_itens_select
function grupo_origem_pedido_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo='PedidoECommerce_Origem_Grupo') AND (st_inativo=0) ORDER BY ordenacao")
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
	
    strResp = "<option value=''>&nbsp;</option>" & strResp

	grupo_origem_pedido_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' __________________________________________________
' origem_pedido_monta_itens_select
'
function origem_pedido_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo='PedidoECommerce_Origem') AND (st_inativo=0) ORDER BY ordenacao")
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
	
    strResp = "<option value=''>&nbsp;</option>" & strResp

	origem_pedido_monta_itens_select = strResp
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
	<title>LOJA</title>
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
    $(function() {
	
    <% if strJS <> "" then Response.Write strJS %>

    $(document).tooltip();

    $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');
    
    <% if lst_indicadores_carrega = "" then %>

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
    
    <% end if %>
		$("#c_dt_entregue_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_entregue_termino").hUtilUI('datepicker_peq_filtro_final');
		$("#c_dt_cancelado_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_cancelado_termino").hUtilUI('datepicker_peq_filtro_final');
		$("#c_dt_cadastro_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_cadastro_termino").hUtilUI('datepicker_peq_filtro_final');
		$("#c_dt_entrega_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_entrega_termino").hUtilUI('datepicker_peq_filtro_final');

		$("#c_dt_previsao_entrega_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_previsao_entrega_termino").hUtilUI('datepicker_peq_filtro_final');

		$("#c_dt_coleta_a_separar_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_coleta_a_separar_termino").hUtilUI('datepicker_peq_filtro_final');

		$("#c_dt_coleta_st_a_entregar_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_coleta_st_a_entregar_termino").hUtilUI('datepicker_peq_filtro_final');

		$(".CkbPagAntQuitSt").change(function () {
			if ($(this).is(":checked")) {
				$("#ckb_pagto_antecipado_status_nao").prop("checked", false);
				$("#ckb_pagto_antecipado_status_sim").prop("checked", true);
			}
		});

		$("#ckb_pagto_antecipado_status_nao").change(function () {
			if ($(this).is(":checked")) {
				$("#ckb_pagto_antecipado_status_sim").prop("checked", false);
				$(".CkbPagAntQuitSt").prop("checked", false);
			}
		});

		$("#ckb_pagto_antecipado_status_sim").change(function () {
			if ($(this).is(":checked")) {
				$("#ckb_pagto_antecipado_status_nao").prop("checked", false);
			}
			else {
				$(".CkbPagAntQuitSt").prop("checked", false);
			}
		});

		$("#c_grupo").change(function () {
			$("#spnCounterGrupo").text($("#c_grupo :selected").length);
		});

		$("#spnCounterGrupo").text($("#c_grupo :selected").length);

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

<!-- ***** INICIO CONSULTA INDICADORES AJAX ------->
<script type="text/javascript">

	<% if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then %>
        var vendedor = "<%=usuario %>";
        var flag_vendedor_normal = "S";
    <% else %>
        var vendedor = "";
        var flag_vendedor_normal = "X";
    <% end if %>

    var loja = "<%=loja %>";

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

	function limpaCampoSelectGrupo() {
		$("#c_grupo").children().prop("selected", false);
		$("#spnCounterGrupo").text($("#c_grupo :selected").length);
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
        if (flag_vendedor_normal == "X") {
            strUrl = strUrl + "&vendedor=" + fFILTRO.c_vendedor.value;
        }
        else {
            strUrl = strUrl + "&vendedor=" + vendedor;
        }
        strUrl = strUrl + "&loja=" + loja;
        fFILTRO.ultimoVendedor.value = fFILTRO.c_vendedor.value;
        objAjaxListaIndicadores.onreadystatechange = TrataRespostaAjaxListaIndicadores;
        objAjaxListaIndicadores.open("GET", strUrl, true);
        objAjaxListaIndicadores.send(null);

    }
</script>

<!-- **** FIM TRECHO CONSULTA INDICADORES AJAX ----->

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;

	if (f.ckb_st_entrega_separar_com_marc.checked) {
		if ((trim(f.c_dt_coleta_a_separar_inicio.value) != "") && (trim(f.c_dt_coleta_a_separar_termino.value) != "")) {
			if (!consiste_periodo(f.c_dt_coleta_a_separar_inicio, f.c_dt_coleta_a_separar_termino)) return;
		}
	}

	if (f.ckb_st_entrega_a_entregar_com_marc.checked) {
		if ((trim(f.c_dt_coleta_st_a_entregar_inicio.value) != "") && (trim(f.c_dt_coleta_st_a_entregar_termino.value) != "")) {
			if (!consiste_periodo(f.c_dt_coleta_st_a_entregar_inicio, f.c_dt_coleta_st_a_entregar_termino)) return;
		}
	}

	if (f.ckb_st_entrega_entregue.checked) {
		if (!consiste_periodo(f.c_dt_entregue_inicio, f.c_dt_entregue_termino)) return;
		}

	if (f.ckb_st_entrega_cancelado.checked) {
		if (!consiste_periodo(f.c_dt_cancelado_inicio, f.c_dt_cancelado_termino)) return;
		}
		
	if (f.ckb_periodo_cadastro.checked) {
		if (trim(f.c_dt_cadastro_inicio.value)=="") {
			alert("Preencha a data!!");
			f.c_dt_cadastro_inicio.focus();
			return;
			}
		if (trim(f.c_dt_cadastro_termino.value)=="") {
			alert("Preencha a data!!");
			f.c_dt_cadastro_termino.focus();
			return;
			}
		if (!consiste_periodo(f.c_dt_cadastro_inicio, f.c_dt_cadastro_termino)) return;
		}

	if (f.ckb_entrega_marcada_para.checked) {
		if (trim(f.c_dt_entrega_inicio.value)=="") {
			alert("Preencha a data!!");
			f.c_dt_entrega_inicio.focus();
			return;
			}
		if (trim(f.c_dt_entrega_termino.value)=="") {
			alert("Preencha a data!!");
			f.c_dt_entrega_termino.focus();
			return;
			}
		if (!consiste_periodo(f.c_dt_entrega_inicio, f.c_dt_entrega_termino)) return;
		}

    if (f.ckb_entrega_imediata_nao.checked) {
        if (!consiste_periodo(f.c_dt_previsao_entrega_inicio, f.c_dt_previsao_entrega_termino)) return;
    }

	if (f.ckb_produto.checked) {
		if (trim(f.c_produto.value)!="") {
			if (!isEAN(f.c_produto.value)) {
				if (trim(f.c_fabricante.value)=="") {
					alert("Preencha o código do fabricante!!");
					f.c_fabricante.focus();
					return;
					}
				}
			}
		if ((trim(f.c_produto.value)=="")&&(trim(f.c_fabricante.value)=="")) {
			alert("Preencha o código do produto!!");
			f.c_produto.focus();
			return;
			}
		}

//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
	//  ENTREGUE ENTRE
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
	//  CANCELADO ENTRE
		strDtRefDDMMYYYY = trim(f.c_dt_cancelado_inicio.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_cancelado_termino.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
	//  COLOCADOS ENTRE
		strDtRefDDMMYYYY = trim(f.c_dt_cadastro_inicio.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_cadastro_termino.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
	//  DATA DE COLETA (RÓTULO ANTIGO: ENTREGA MARCADA ENTRE)
		strDtRefDDMMYYYY = trim(f.c_dt_entrega_inicio.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_entrega_termino.value);
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
	
	<% if lst_indicadores_carrega = "" then %>
	try {
		if (('localStorage' in window) && window['localStorage'] !== null) {
			var d = $("#c_indicador").html();
			localStorage.setItem('c_indicador', d);
		}
	}
	catch (e) {
		// NOP
	}
	<% end if %>

	fFILTRO.c_hidden_reload.value = 1;
	fFILTRO.c_hidden_indice_indicador.value = $("#c_indicador option:selected").index();
	fFILTRO.ultimoVendedor.value = fFILTRO.c_vendedor.value;
	
	f.c_FormFieldValues.value=formToString($("#fFILTRO"));

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
 </style>

<body>
<center>

<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity: 0.6;visibility:hidden;vertical-align: middle">

	</div>


<form id="fFILTRO" name="fFILTRO" method="post" action="RelPedidosMCritExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_loja" id="c_loja" value="<%=loja%>">
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>

<input type="hidden" id="ultimoVendedor" name="ultimoVendedor" value="x-x-x-x-x-x" />
<input type="hidden" name="c_hidden_reload" id="c_hidden_reload" value="0" />
<input type="hidden" name="c_hidden_indice_indicador" id="c_hidden_indice_indicador" value="" />

<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório Multicritério de Pedidos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table class="Qx" cellspacing="0" width="650">
<!--  STATUS DE ENTREGA  -->
<tr bgcolor="#FFFFFF">
<td class="MT" align="left" nowrap><span class="PLTe">STATUS DE ENTREGA</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_esperar" name="ckb_st_entrega_esperar"
			value="<%=ST_ENTREGA_ESPERAR%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_esperar.click();">Esperar</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_split" name="ckb_st_entrega_split"
			value="<%=ST_ENTREGA_SPLIT_POSSIVEL%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_split.click();">Split Possível</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_separar_sem_marc" name="ckb_st_entrega_separar_sem_marc"
			value="<%=ST_ENTREGA_SEPARAR%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_separar_sem_marc.click();">A Separar (sem data de coleta)</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_separar_com_marc" name="ckb_st_entrega_separar_com_marc"
			value="<%=ST_ENTREGA_SEPARAR%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_separar_com_marc.click();">A Separar (com data de coleta)</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_coleta_a_separar_inicio" id="c_dt_coleta_a_separar_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_coleta_a_separar_termino.focus(); else fFILTRO.ckb_st_entrega_separar_com_marc.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_separar_com_marc.checked=true;" onchange="fFILTRO.ckb_st_entrega_separar_com_marc.checked=true;"
			/>&nbsp;<span class="C">a</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_coleta_a_separar_termino" id="c_dt_coleta_a_separar_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_st_entrega_separar_com_marc.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_separar_com_marc.checked=true;" onchange="fFILTRO.ckb_st_entrega_separar_com_marc.checked=true;" />
		<span style="display:inline-block;width:2px;"></span>
		<a name="bLimparStEtgSeparar" id="bLimparStEtgSeparar" href="javascript:limpaMultiplosCampos(fFILTRO.ckb_st_entrega_separar_com_marc,fFILTRO.c_dt_coleta_a_separar_inicio,fFILTRO.c_dt_coleta_a_separar_termino);" title="limpa os campos deste filtro">
					<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_a_entregar_sem_marc" name="ckb_st_entrega_a_entregar_sem_marc"
			value="<%=ST_ENTREGA_A_ENTREGAR%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_a_entregar_sem_marc.click();">A Entregar (sem data de coleta)</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_a_entregar_com_marc" name="ckb_st_entrega_a_entregar_com_marc"
			value="<%=ST_ENTREGA_A_ENTREGAR%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_a_entregar_com_marc.click();">A Entregar (com data de coleta)</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_coleta_st_a_entregar_inicio" id="c_dt_coleta_st_a_entregar_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_coleta_st_a_entregar_termino.focus(); else fFILTRO.ckb_st_entrega_a_entregar_com_marc.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_a_entregar_com_marc.checked=true;" onchange="fFILTRO.ckb_st_entrega_a_entregar_com_marc.checked=true;"
			/>&nbsp;<span class="C">a</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_coleta_st_a_entregar_termino" id="c_dt_coleta_st_a_entregar_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_st_entrega_a_entregar_com_marc.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_a_entregar_com_marc.checked=true;" onchange="fFILTRO.ckb_st_entrega_a_entregar_com_marc.checked=true;" />
		<span style="display:inline-block;width:2px;"></span>
		<a name="bLimparStEtgAEntregar" id="bLimparStEtgAEntregar" href="javascript:limpaMultiplosCampos(fFILTRO.ckb_st_entrega_a_entregar_com_marc,fFILTRO.c_dt_coleta_st_a_entregar_inicio,fFILTRO.c_dt_coleta_st_a_entregar_termino);" title="limpa os campos deste filtro">
					<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_entregue" name="ckb_st_entrega_entregue"
			onclick="if (fFILTRO.ckb_st_entrega_entregue.checked) fFILTRO.c_dt_entregue_inicio.focus();"
			value="<%=ST_ENTREGA_ENTREGUE%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_entregue.click();">Entregue entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entregue_termino.focus(); else fFILTRO.ckb_st_entrega_entregue.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_entregue.checked=true;" onchange="fFILTRO.ckb_st_entrega_entregue.checked=true;"
			>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entregue_termino" id="c_dt_entregue_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_st_entrega_entregue.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_entregue.checked=true;" onchange="fFILTRO.ckb_st_entrega_entregue.checked=true;">
		<span style="display:inline-block;width:2px;"></span>
		<a name="bLimparStEtgEntregue" id="bLimparStEtgEntregue" href="javascript:limpaMultiplosCampos(fFILTRO.ckb_st_entrega_entregue,fFILTRO.c_dt_entregue_inicio,fFILTRO.c_dt_entregue_termino);" title="limpa os campos deste filtro">
					<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_cancelado" name="ckb_st_entrega_cancelado" onclick="if (fFILTRO.ckb_st_entrega_cancelado.checked) fFILTRO.c_dt_cancelado_inicio.focus();"
			value="<%=ST_ENTREGA_CANCELADO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_cancelado.click();">Cancelado entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cancelado_inicio" id="c_dt_cancelado_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cancelado_termino.focus(); else fFILTRO.ckb_st_entrega_cancelado.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_cancelado.checked=true;" onchange="fFILTRO.ckb_st_entrega_cancelado.checked=true;"
			>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cancelado_termino" id="c_dt_cancelado_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_st_entrega_cancelado.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_cancelado.checked=true;" onchange="fFILTRO.ckb_st_entrega_cancelado.checked=true;">
	    <span class="C" style="cursor:default">ordenado por</span>
        <select name="c_cancelados_ordena" id="c_cancelados_ordena">
            <option value="VENDEDOR" selected>Vendedor</option>
            <option value="PEDIDO">Pedido</option>
        </select>
		<span style="display:inline-block;width:2px;"></span>
		<a name="bLimparStEtgCancelado" id="bLimparStEtgCancelado" href="javascript:limpaMultiplosCampos(fFILTRO.ckb_st_entrega_cancelado,fFILTRO.c_dt_cancelado_inicio,fFILTRO.c_dt_cancelado_termino);" title="limpa os campos deste filtro">
					<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
    </td>
	</tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_exceto_cancelados" name="ckb_st_entrega_exceto_cancelados"
			value="<%=ST_ENTREGA_CANCELADO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_exceto_cancelados.click();">Exceto Cancelados</span>
		</td></tr>
	</table>
</td></tr>

<!--  PEDIDOS RECEBIDOS PELO CLIENTE  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">PEDIDOS RECEBIDOS PELO CLIENTE</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
    <tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_pedido_nao_recebido_pelo_cliente" name="ckb_pedido_nao_recebido_pelo_cliente"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_pedido_nao_recebido_pelo_cliente.click();">Não recebido pelo cliente</span>
		</td></tr>
    <tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_pedido_recebido_pelo_cliente" name="ckb_pedido_recebido_pelo_cliente"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_pedido_recebido_pelo_cliente.click();">Recebido pelo cliente</span>
		</td></tr>
	</table>
</td></tr>

<!--  STATUS DE PAGAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">STATUS DE PAGAMENTO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_pago" name="ckb_st_pagto_pago"
			value="<%=ST_PAGTO_PAGO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_pago.click();">Pago</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_nao_pago" name="ckb_st_pagto_nao_pago"
			value="<%=ST_PAGTO_NAO_PAGO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_nao_pago.click();">Não-Pago</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_pago_parcial" name="ckb_st_pagto_pago_parcial"
			value="<%=ST_PAGTO_PARCIAL%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_pago_parcial.click();">Pago Parcial</span>
		</td></tr>
	</table>
</td></tr>

<!--  PAGAMENTO ANTECIPADO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;" width="100%">
	<tr>
		<td width="50%" valign="top">
			<span class="PLTe">PAGAMENTO ANTECIPADO</span>
			<br />
			<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF"><td align="left">
				<input type="checkbox" tabindex="-1" id="ckb_pagto_antecipado_status_nao" name="ckb_pagto_antecipado_status_nao"
					value="0"><span class="C" style="cursor:default" 
					onclick="fFILTRO.ckb_pagto_antecipado_status_nao.click();">Não</span>
				</td></tr>
			<tr bgcolor="#FFFFFF"><td align="left">
				<input type="checkbox" tabindex="-1" id="ckb_pagto_antecipado_status_sim" name="ckb_pagto_antecipado_status_sim"
					value="1"><span class="C" style="cursor:default" 
					onclick="fFILTRO.ckb_pagto_antecipado_status_sim.click();">Sim</span>
				</td></tr>
			</table>
		</td>
		<td width="50%" valign="top">
			<span class="PLTe">STATUS PAGAMENTO ANTECIPADO</span>
			<br />
			<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF"><td align="left">
				<input type="checkbox" class="CkbPagAntQuitSt" tabindex="-1" id="ckb_pagto_antecipado_quitado_status_pendente" name="ckb_pagto_antecipado_quitado_status_pendente"
					value="<%=COD_PAGTO_ANTECIPADO_QUITADO_STATUS_PENDENTE%>"><span class="C" style="cursor:default;color:<%=pagto_antecipado_quitado_cor(COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO, COD_PAGTO_ANTECIPADO_QUITADO_STATUS_PENDENTE)%>;" 
					onclick="fFILTRO.ckb_pagto_antecipado_quitado_status_pendente.click();"><%=pagto_antecipado_quitado_descricao(COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO, COD_PAGTO_ANTECIPADO_QUITADO_STATUS_PENDENTE)%></span>
				</td></tr>
			<tr bgcolor="#FFFFFF"><td align="left">
				<input type="checkbox" class="CkbPagAntQuitSt" tabindex="-1" id="ckb_pagto_antecipado_quitado_status_quitado" name="ckb_pagto_antecipado_quitado_status_quitado"
					value="<%=COD_PAGTO_ANTECIPADO_QUITADO_STATUS_QUITADO%>"><span class="C" style="cursor:default;color:<%=pagto_antecipado_quitado_cor(COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO, COD_PAGTO_ANTECIPADO_QUITADO_STATUS_QUITADO)%>;" 
					onclick="fFILTRO.ckb_pagto_antecipado_quitado_status_quitado.click();"><%=pagto_antecipado_quitado_descricao(COD_PAGTO_ANTECIPADO_STATUS_ANTECIPADO, COD_PAGTO_ANTECIPADO_QUITADO_STATUS_QUITADO)%></span>
				</td></tr>
			</table>
		</td>
	</tr>
	</table>
</td></tr>

<!--  ANÁLISE DE CRÉDITO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">ANÁLISE DE CRÉDITO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;" width="100%">
	<tr>
		<td width="50%" valign="top">
			<table cellspacing="0" cellpadding="0">
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_st_inicial" name="ckb_analise_credito_st_inicial"
						value="<%=COD_AN_CREDITO_ST_INICIAL%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_ST_INICIAL)%>;" 
						onclick="fFILTRO.ckb_analise_credito_st_inicial.click();">Status Inicial</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_pendente_vendas" name="ckb_analise_credito_pendente_vendas"
						value="<%=COD_AN_CREDITO_PENDENTE_VENDAS%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_PENDENTE_VENDAS)%>;" 
						onclick="fFILTRO.ckb_analise_credito_pendente_vendas.click();">Pendente Vendas</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_pendente_endereco" name="ckb_analise_credito_pendente_endereco"
						value="<%=COD_AN_CREDITO_PENDENTE_ENDERECO%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_PENDENTE_ENDERECO)%>;" 
						onclick="fFILTRO.ckb_analise_credito_pendente_endereco.click();">Pendente Endereço</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_pendente" name="ckb_analise_credito_pendente"
						value="<%=COD_AN_CREDITO_PENDENTE%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_PENDENTE)%>;" 
						onclick="fFILTRO.ckb_analise_credito_pendente.click();">Pendente</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_pendente_cartao" name="ckb_analise_credito_pendente_cartao"
						value="<%=COD_AN_CREDITO_PENDENTE_CARTAO%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_PENDENTE_CARTAO)%>;" 
						onclick="fFILTRO.ckb_analise_credito_pendente_cartao.click();">Pendente Cartão de Crédito</span>
					</td></tr>
			</table>
		</td>
		<td width="50%" valign="top">
			<table cellspacing="0" cellpadding="0">
				<tr bgcolor="#FFFFFF"><td align="left">
				<input type="checkbox" tabindex="-1" id="ckb_analise_credito_pendente_pagto_antecipado_boleto" name="ckb_analise_credito_pendente_pagto_antecipado_boleto"
					value="<%=COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO%>" /><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)%>;"
					onclick="fFILTRO.ckb_analise_credito_pendente_pagto_antecipado_boleto.click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)%></span>
				</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_ok_aguardando_deposito" name="ckb_analise_credito_ok_aguardando_deposito"
						value="<%=COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)%>;" 
						onclick="fFILTRO.ckb_analise_credito_ok_aguardando_deposito.click();">Crédito OK (aguardando depósito)</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_ok_deposito_aguardando_desbloqueio" name="ckb_analise_credito_ok_deposito_aguardando_desbloqueio"
						value="<%=COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)%>;" 
						onclick="fFILTRO.ckb_analise_credito_ok_deposito_aguardando_desbloqueio.click();">Crédito OK (depósito aguardando desbloqueio)</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_ok_aguardando_pagto_boleto_av" name="ckb_analise_credito_ok_aguardando_pagto_boleto_av"
						value="<%=COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV%>" /><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)%>;"
						onclick="fFILTRO.ckb_analise_credito_ok_aguardando_pagto_boleto_av.click();"><%=x_analise_credito(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)%></span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_ok" name="ckb_analise_credito_ok"
						value="<%=COD_AN_CREDITO_OK%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_OK)%>;" 
						onclick="fFILTRO.ckb_analise_credito_ok.click();">Crédito OK</span>
					</td></tr>
			</table>
		</td>
	</tr>
	</table>
</td></tr>

<!--  ENTREGA IMEDIATA  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">ENTREGA IMEDIATA</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_entrega_imediata_sim" name="ckb_entrega_imediata_sim"
			value="<%=COD_ETG_IMEDIATA_SIM%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_entrega_imediata_sim.click();">Sim</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_entrega_imediata_nao" name="ckb_entrega_imediata_nao"
			value="<%=COD_ETG_IMEDIATA_NAO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_entrega_imediata_nao.click();">Não</span>
			<span style="width:50px;">&nbsp;</span>
			<span class="C" style="cursor:default;" onclick="fFILTRO.ckb_entrega_imediata_nao.click();">Previsão de Entrega entre</span>
			<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_previsao_entrega_inicio" id="c_dt_previsao_entrega_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_previsao_entrega_termino.focus(); else fFILTRO.ckb_entrega_imediata_nao.checked=true; filtra_data();" onclick="fFILTRO.ckb_entrega_imediata_nao.checked = true;" onchange="fFILTRO.ckb_entrega_imediata_nao.checked=true;"
			/>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_previsao_entrega_termino" id="c_dt_previsao_entrega_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_entrega_imediata_nao.checked=true; filtra_data();" onclick="fFILTRO.ckb_entrega_imediata_nao.checked = true;" onchange="fFILTRO.ckb_entrega_imediata_nao.checked=true;" />
			<span style="display:inline-block;width:2px;"></span>
			<a name="bLimparEtgImediataNao" id="bLimparEtgImediataNao" href="javascript:limpaMultiplosCampos(fFILTRO.ckb_entrega_imediata_nao,fFILTRO.c_dt_previsao_entrega_inicio,fFILTRO.c_dt_previsao_entrega_termino);" title="limpa os campos deste filtro">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
	</tr>
	</table>
</td></tr>

<!--  GERAL  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">GERAL</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;" width="100%">
	<tr>
		<td width="50%" valign="top">
			<table cellspacing="0" cellpadding="0">
				<tr bgcolor="#FFFFFF"><td align="left">
					<%	s_memoria = get_default_valor_texto_bd(usuario, "LOJA/RelPedidosMCrit|ckb_nao_exibir_links") %>
					<input type="checkbox" tabindex="-1" id="ckb_nao_exibir_links" name="ckb_nao_exibir_links"
						value="ON" <%if s_memoria <> "" then Response.Write " checked"%> /><span class="C" style="cursor:default" 
						onclick="fFILTRO.ckb_nao_exibir_links.click();">Não exibir links</span>
					</td></tr>
			</table>
		</td>
	</tr>
	</table>
</td></tr>

<!--  PEDIDOS COM OU SEM INDICADOR  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">INDICADOR</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_indicador_preenchido" name="ckb_indicador_preenchido"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_indicador_preenchido.click();">Indicador preenchido</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_indicador_nao_preenchido" name="ckb_indicador_nao_preenchido"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_indicador_nao_preenchido.click();">Indicador não preenchido</span>
		</td></tr>
	</table>
</td></tr>

<!--  PERÍODO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">PERÍODO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_periodo_cadastro" name="ckb_periodo_cadastro" onclick="if (fFILTRO.ckb_periodo_cadastro.checked) fFILTRO.c_dt_cadastro_inicio.focus();"
			value="PERIODO_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_periodo_cadastro.click();">Somente pedidos colocados entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cadastro_inicio" id="c_dt_cadastro_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cadastro_termino.focus(); else fFILTRO.ckb_periodo_cadastro.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_cadastro.checked=true;" onchange="fFILTRO.ckb_periodo_cadastro.checked=true;"
			>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cadastro_termino" id="c_dt_cadastro_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_periodo_cadastro.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_cadastro.checked=true;" onchange="fFILTRO.ckb_periodo_cadastro.checked=true;">
		<span style="display:inline-block;width:2px;"></span>
		<a name="bLimparPeriodoCadastro" id="bLimparPeriodoCadastro" href="javascript:limpaMultiplosCampos(fFILTRO.ckb_periodo_cadastro,fFILTRO.c_dt_cadastro_inicio,fFILTRO.c_dt_cadastro_termino);" title="limpa os campos deste filtro">
					<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_entrega_marcada_para" name="ckb_entrega_marcada_para" onclick="if (fFILTRO.ckb_entrega_marcada_para.checked) fFILTRO.c_dt_entrega_inicio.focus();"
			value="ENTREGA_MARCADA_PARA_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_entrega_marcada_para.click();">Data de coleta no período entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entrega_inicio" id="c_dt_entrega_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entrega_termino.focus(); else fFILTRO.ckb_entrega_marcada_para.checked=true; filtra_data();" onclick="fFILTRO.ckb_entrega_marcada_para.checked=true;" onchange="fFILTRO.ckb_entrega_marcada_para.checked=true;"
			>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entrega_termino" id="c_dt_entrega_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_entrega_marcada_para.checked=true; filtra_data();" onclick="fFILTRO.ckb_entrega_marcada_para.checked=true;" onchange="fFILTRO.ckb_entrega_marcada_para.checked=true;">
		<span style="display:inline-block;width:2px;"></span>
		<a name="bLimparPeriodoColeta" id="bLimparPeriodoColeta" href="javascript:limpaMultiplosCampos(fFILTRO.ckb_entrega_marcada_para,fFILTRO.c_dt_entrega_inicio,fFILTRO.c_dt_entrega_termino);" title="limpa os campos deste filtro">
					<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
	</tr>
	</table>
</td></tr>

<!--  PRODUTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">PRODUTO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_produto" name="ckb_produto" onclick="if (fFILTRO.ckb_produto.checked) fFILTRO.c_fabricante.focus();"
			value="PRODUTO_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_produto.click();">Somente pedidos que incluam:</span
			><br><span class="C" style="margin-left:30px;">Fabricante</span><input maxlength="4" class="Cc" style="width:50px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto.focus(); else fFILTRO.ckb_produto.checked=true; filtra_fabricante();" onclick="fFILTRO.ckb_produto.checked=true;">
			<span class="C">&nbsp;&nbsp;&nbsp;Produto</span><input maxlength="13" class="Cc" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_produto.checked=true; filtra_produto();" onclick="fFILTRO.ckb_produto.checked=true;">
		<span style="display:inline-block;width:2px;"></span>
		<a name="bLimparProduto" id="bLimparProduto" href="javascript:limpaMultiplosCampos(fFILTRO.ckb_produto,fFILTRO.c_fabricante,fFILTRO.c_produto);" title="limpa os campos deste filtro">
					<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
	</tr>
	</table>
</td></tr>

<!-- GRUPO DE PRODUTOS -->
<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" colspan="2" align="left" nowrap>
		<span class="PLTe">GRUPO DE PRODUTOS</span>
		<br>
		<table cellpadding="0" cellspacing="0" style="margin:1px 3px 6px 10px;">
		<tr>
		<td>
			<select id="c_grupo" name="c_grupo" class="LST" size="5" style="width:200px" multiple>
			<% =t_produto_grupo_monta_itens_select(Null) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparGrupo" id="bLimparGrupo" href="javascript:limpaCampoSelectGrupo()" title="limpa o filtro 'Grupo de Produtos'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
                        <br />
                        (<span class="Lbl" id="spnCounterGrupo"></span>)
		</td>
		</tr>
		</table>
	</td>
	</tr>

<%if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then%>
<!-- ORIGEM DO PEDIDO (GRUPO) -->
    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">ORIGEM DO PEDIDO (GRUPO)</span>
		<br>           
			<select id="c_grupo_pedido_origem" name="c_grupo_pedido_origem" style="margin:1px 3px 6px 10px;width: 200px" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =grupo_origem_pedido_monta_itens_select(Null) %>
			</select>
			
        </td></tr>

<!-- ORIGEM DO PEDIDO -->
    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">ORIGEM DO PEDIDO</span>
		<br>
			<select id="c_pedido_origem" name="c_pedido_origem" style="margin:1px 3px 6px 10px;width: 200px" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =origem_pedido_monta_itens_select(Null) %>
			</select>
			
        </td></tr>
<%end if%>

<!-- EMPRESA -->

    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">EMPRESA</span>
		<br>
			<select id="c_empresa" name="c_empresa" style="margin:1px 3px 6px 10px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
			</select>
			
        </td>
    </tr>

<!--  NOVA VERSÃO DA FORMA DE PAGAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">FORMA DE PAGAMENTO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<span class="C" style="margin-left:30px;">Forma de Pagamento</span>
			<select id="op_forma_pagto" name="op_forma_pagto">
			  <% =forma_pagto_monta_itens_select(Null) %>
			</select>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<span class="C" style="margin-left:30px;">Nº Parcelas</span>
			<input class="Cc" maxlength="2" style="width:40px;" name="c_forma_pagto_qtde_parc" id="c_forma_pagto_qtde_parc" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); filtra_numerico();">
		</td></tr>
	</table>
</td></tr>

<!--  CLIENTE  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">CLIENTE</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<span class="C" style="margin-left:30px;">CNPJ/CPF</span>
			<input class="C" maxlength="18" style="width:140px;" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="if (digitou_enter(true)&&((!tem_info(this.value))||(tem_info(this.value)&&cnpj_cpf_ok(this.value)))) {this.value=cnpj_cpf_formata(this.value); bCONFIRMA.focus();} filtra_cnpj_cpf();">
		</td></tr>
    <tr bgcolor="#FFFFFF"><td align="left">
		<span class="C" style="margin-left:30px;">UF</span>
			<select name="c_cliente_uf" id="c_cliente_uf">
                <%=UF_monta_itens_select(Null) %>
			</select>
		</td></tr>
	</table>
</td></tr>

<!--  CARTÃO DE CRÉDITO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">CARTÃO DE CRÉDITO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_visanet" name="ckb_visanet"
			value="VISANET_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_visanet.click();">Somente pedidos pagos usando cartão de crédito</span>
		</td></tr>
	</table>
</td></tr>

<!--  PERCENTUAL DE COMISSÃO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">COMISSÃO (%)</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF">
		<td align="left" valign="bottom">
			<input type="checkbox" tabindex="-1" id="ckb_perc_RT" name="ckb_perc_RT"
				value="COMISSAO_ON"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_perc_RT.click();fFILTRO.c_perc_RT.focus();">Somente pedidos com comissão abaixo de </span>
			<input class="Cd" maxlength="5" style="width:60px;" name="c_perc_RT" id="c_perc_RT" 
				onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); else fFILTRO.ckb_perc_RT.checked=true; filtra_percentual();"
				onblur="this.value=formata_perc_RT(this.value); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();}"
				onclick="fFILTRO.ckb_perc_RT.checked=true;"
				onchange="fFILTRO.ckb_perc_RT.checked=true;"
				onfocus="fFILTRO.c_perc_RT.select();"
				value="<%=formata_perc_RT(5)%>"
				/><span class="C">%</span>
		</td>
	</tr>
	</table>
</td></tr>

<!--  TRANSPORTADORA  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">TRANSPORTADORA</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<span class="C" style="margin-left:30px;">Identificação</span>
			<input class="C" maxlength="10" style="width:110px;" name="c_transportadora" id="c_transportadora" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_nome_identificador();">
		</td></tr>
	</table>
</td></tr>

<!--  CADASTRAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">CADASTRAMENTO</span>
	<br>
	<table cellspacing="6" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF">
		<td align="right"><span class="C" style="margin-left:20px;">Vendedor</span></td>
		<td align="left">
			<select id="c_vendedor" name="c_vendedor" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" <% if lst_indicadores_carrega = "" then %> onchange="LimpaListaIndicadores()" <% end if %>>
			<% =vendedores_desta_loja_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="right" valign="top" style="width: 70px"><span class="C" style="margin-left:2px;"><% if lst_indicadores_carrega = "" then %><img id="exclamacao" src="../IMAGEM/exclamacao_14x14.png" title="Reduza o tempo de carregamento da lista de indicadores, filtrando por vendedor." style="cursor:pointer;" />&nbsp;<% end if %>Indicador</span></td>
		<td align="left">
			<select id="c_indicador" name="c_indicador" style="margin-right:10px;" <% if lst_indicadores_carrega = "" then %> onfocus="CarregaListaIndicadores()" <% end if %> onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			
			<option selected value=''>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
			<% if lst_indicadores_carrega <> "" then
			Response.Write  indicadores_desta_loja_monta_itens_select(null)
			end if%>
			</select><br />
			<span class="aviso">Vendedor selecionado não possui indicadores.</span>&nbsp;
		</td>
	</tr>
	</table>
</td></tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0">
<tr><td class="Rc" align="left" style="border-bottom:1px solid black">&nbsp;</td></tr>
<tr><td align="right"><a href="RelPedidosMCrit.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="Limpar filtros" class="LSessaoEncerra">Limpar filtros</a></td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" 
		<% if url_back <> "" then %>
		href="Resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>"
		<% else %>
		href="javascript:history.back()"
		<% end if %>
		>
		<img src="../botao/voltar.gif" width="176" height="55" border="0" title="volta para a página anterior"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0" title="executa a consulta"></a></div>
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