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
'	  R E L O R C A M E N T O S M C R I T . A S P
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

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

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

	dim url_origem
	url_origem = Trim(Request("url_origem"))

	' PREENCHIMENTO DA LISTA DE INDICADORES: GRAVA ÚLTIMA OPÇÃO DE CONSULTA NO BD
	dim lst_indicadores_carrega
	lst_indicadores_carrega = Request.Form("ckb_rel_mcrit_orc_carrega_indicadores")
	if url_origem = "" then
		' GRAVA PARÂMETRO APENAS SE O ACIONAMENTO FOI REALIZADO A PARTIR DA PÁGINA INICIAL
		call set_default_valor_texto_bd(usuario, "RelOrcamentosMCrit|c_carrega_indicadores_estatico", lst_indicadores_carrega)
		end if

	' SE ESTA PÁGINA FOI ACIONADA COMO RETORNO DE OUTRA PÁGINA DECORRENTE DA CONSULTA DE UM PEDIDO DA LISTA DE RESULTADOS, RESTAURA OS FILTROS
	dim strJS, c_FormFieldValues
	strJS = ""
	c_FormFieldValues = ""
	if url_origem <> "" then
		c_FormFieldValues = get_default_valor_texto_bd(usuario, "CENTRAL/RelOrcamentosMCrit|FormFields")
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
' VENDEDORES MONTA ITENS SELECT
'
function vendedores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT usuario, nome_iniciais_em_maiusculas FROM" & _
			 " (" & _
			 "SELECT usuario, nome_iniciais_em_maiusculas FROM t_USUARIO" & _
				" WHERE (vendedor_loja <> 0)" & _
			 " UNION" & _
			 " SELECT t_USUARIO.usuario AS usuario, t_USUARIO.nome_iniciais_em_maiusculas FROM t_USUARIO" & _
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

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	vendedores_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' ____________________________________________________________________________
' INDICADORES MONTA ITENS SELECT

function indicadores_monta_itens_select(byval id_default, byval incluirItemBrancoSeNaoHouverDefault)
    dim x, r, strResp, ha_default
	    id_default = Trim("" & id_default)
	    ha_default=False
	    set r = cn.Execute("SELECT apelido, razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR WHERE (Id NOT IN (" & Cstr(ID_NSU_ORCAMENTISTA_E_INDICADOR__RESTRICAO_FP_TODOS) & "," & Cstr(ID_NSU_ORCAMENTISTA_E_INDICADOR__SEM_INDICADOR) & ")) ORDER BY apelido")
	    strResp = ""
	    do while Not r.eof 
		    x = UCase(Trim("" & r("apelido")))
		    if (id_default<>"") And (id_default=x) then
			    strResp = strResp & "<OPTION SELECTED"
			    ha_default=True
		    else
			    strResp = strResp & "<OPTION"
			    end if
		    strResp = strResp & " VALUE='" & x & "'>"
		    strResp = strResp & x & " - " & Trim("" & r("razao_social_nome_iniciais_em_maiusculas"))
		    strResp = strResp & "</OPTION>" & chr(13)
		    r.MoveNext
		    loop

	    if (Not ha_default) And incluirItemBrancoSeNaoHouverDefault then
		    strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		    end if
    		
	    indicadores_monta_itens_select = strResp
	    r.close
	    set r=nothing
end function

'----------------------------------------------------------------------------------------------
' grupo_origem_pedido_monta_itens_select
function grupo_origem_pedido_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE grupo='PedidoECommerce_Origem_Grupo' AND st_inativo=0")
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
	
    'strResp = "<option value=''>&nbsp;</option>" & strResp

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

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE grupo='PedidoECommerce_Origem' AND st_inativo=0")
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
	
    'strResp = "<option value=''>&nbsp;</option>" & strResp

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
	<title>CENTRAL</title>
	</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function () {
		<% if strJS <> "" then Response.Write strJS %>

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

		$("#c_dt_cadastro_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_cadastro_termino").hUtilUI('datepicker_filtro_final');

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
		$(document).tooltip();
	});
</script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;

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
		
	if (f.rb_loja[1].checked) {
		if (converte_numero(f.c_loja.value)==0) {
			alert("Especifique o número da loja!!");
			f.c_loja.focus();
			return;
			}
		}

	if (f.rb_loja[2].checked) {
		if (trim(f.c_loja_de.value)!="") {
			if (converte_numero(f.c_loja_de.value)==0) {
				alert("Número de loja inválido!!");
				f.c_loja_de.focus();
				return;
				}
			}
		if (trim(f.c_loja_ate.value)!="") {
			if (converte_numero(f.c_loja_ate.value)==0) {
				alert("Número de loja inválido!!");
				f.c_loja_ate.focus();
				return;
				}
			}
		if ((trim(f.c_loja_de.value)=="")&&(trim(f.c_loja_ate.value)=="")) {
			alert("Preencha pelo menos um dos campos!!");
			f.c_loja_de.focus();
			return;
			}
		if ((trim(f.c_loja_de.value)!="")&&(trim(f.c_loja_ate.value)!="")) {
			if (converte_numero(f.c_loja_ate.value)<converte_numero(f.c_loja_de.value)) {
				alert("Faixa de lojas inválida!!");
				f.c_loja_ate.focus();
				return;
				}
			}
		}
		
//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
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
		}
		
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";

	<% if lst_indicadores_carrega = "" then %>
	    if (('localStorage' in window) && window['localStorage'] !== null) {
		var d = $("#c_indicador").html();
		localStorage.setItem('c_indicador', d);
	}
	<% end if %>

	fFILTRO.c_hidden_reload.value = 1;
	fFILTRO.c_hidden_indice_indicador.value = $("#c_indicador option:selected").index();
	fFILTRO.ultimoVendedor.value = fFILTRO.c_vendedor.value;
    
	f.c_FormFieldValues.value=formToString($("#fFILTRO"));

	f.submit();
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

<form id="fFILTRO" name="fFILTRO" method="post" action="RelOrcamentosMCritExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" id="ultimoVendedor" name="ultimoVendedor" value="x-x-x-x-x-x" />
<input type="hidden" name="c_hidden_reload" id="c_hidden_reload" value="0" />
<input type="hidden" name="c_hidden_indice_indicador" id="c_hidden_indice_indicador" value="" />
<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório Multicritério de Orçamentos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table class="Qx" cellSpacing="0">

<!--  STATUS DO ORÇAMENTO  -->
<tr bgColor="#FFFFFF">
<td class="MT" NOWRAP><span class="PLTe">STATUS DO ORÇAMENTO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_orcamento_em_aberto" name="ckb_orcamento_em_aberto"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_orcamento_em_aberto.click();">Orçamento em aberto</span>
		</td></tr>
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_orcamento_virou_pedido" name="ckb_orcamento_virou_pedido"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_orcamento_virou_pedido.click();">Orçamento virou pedido</span>
		</td></tr>
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_orcamento_cancelado" name="ckb_orcamento_cancelado"
			value="<%=ST_ORCAMENTO_CANCELADO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_orcamento_cancelado.click();">Cancelado</span>
		</td></tr>
	</table>
</td></tr>

<!--  PERÍODO  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">PERÍODO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;margin-right:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_periodo_cadastro" name="ckb_periodo_cadastro" onclick="if (fFILTRO.ckb_periodo_cadastro.checked) fFILTRO.c_dt_cadastro_inicio.focus();"
			value="PERIODO_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_periodo_cadastro.click();">Somente orçamentos cadastrados entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cadastro_inicio" id="c_dt_cadastro_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cadastro_termino.focus(); else fFILTRO.ckb_periodo_cadastro.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_cadastro.checked = true;" onchange="fFILTRO.ckb_periodo_cadastro.checked = true;"
			/>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cadastro_termino" id="c_dt_cadastro_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_periodo_cadastro.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_cadastro.checked = true;" onchange="fFILTRO.ckb_periodo_cadastro.checked = true;" />
		</td></tr>
	</table>
</td></tr>

<!--  PRODUTO  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">PRODUTO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_produto" name="ckb_produto" onclick="if (fFILTRO.ckb_produto.checked) fFILTRO.c_fabricante.focus();"
			value="PRODUTO_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_produto.click();">Somente orçamentos que incluam:</span
			><br><span class="C" style="margin-left:30px;">Fabricante</span><input maxlength="4" class="Cc" style="width:50px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto.focus(); else fFILTRO.ckb_produto.checked=true; filtra_fabricante();" onclick="fFILTRO.ckb_produto.checked=true;">
			<span class="C">&nbsp;&nbsp;&nbsp;Produto</span><input maxlength="13" class="Cc" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_produto.checked=true; filtra_produto();" onclick="fFILTRO.ckb_produto.checked=true;">
		</td></tr>
	</table>
</td></tr>

<!--  LOJAS  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">LOJAS</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="radio" tabindex="-1" id="rb_loja" name="rb_loja"
			value="TODAS" checked><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_loja[0].click();">Todas as lojas</span>
		</td></tr>
	<tr bgColor="#FFFFFF"><td>
		<input type="radio" tabindex="-1" id="rb_loja" name="rb_loja" onclick="if (fFILTRO.rb_loja[1].checked) fFILTRO.c_loja.focus();"
			value="UMA"><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_loja[1].click();">Loja</span>
			<input class="Cc" maxlength="3" style="width:40px;" name="c_loja" id="c_loja" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); else this.click(); filtra_numerico();" onclick="fFILTRO.rb_loja[1].checked=true;">
		</td></tr>
	<tr bgColor="#FFFFFF"><td>
		<input type="radio" tabindex="-1" id="rb_loja" name="rb_loja" onclick="if (fFILTRO.rb_loja[2].checked) fFILTRO.c_loja_de.focus();"
			value="FAIXA"><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_loja[2].click();">Lojas</span>
			<input class="Cc" maxlength="3" style="width:40px;" name="c_loja_de" id="c_loja_de" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fFILTRO.c_loja_ate.focus(); else this.click(); filtra_numerico();" onclick="fFILTRO.rb_loja[2].checked=true;">
			<span class="C">a</span>
			<input class="Cc" maxlength="3" style="width:40px;" name="c_loja_ate" id="c_loja_ate" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); filtra_numerico();" onclick="fFILTRO.rb_loja[2].checked=true;">
		</td></tr>
	</table>
</td></tr>

<!--  CADASTRAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">CADASTRAMENTO</span>
    	<br>
	<table cellspacing="6" cellpadding="0" style="margin-bottom:0px;">
	<tr bgcolor="#FFFFFF">
		<td style="width:70px; text-align:right"><span class="C" style="text-align: right; margin-left: 2px">Vendedor</span></td>
		<td align="left">
			<select id="c_vendedor" name="c_vendedor" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" <% if lst_indicadores_carrega = "" then %>onchange="LimpaListaIndicadores()" <% end if %>>
			<% =vendedores_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="right" valign="top" style="width: 70px"><span class="C" style="margin-left:2px;"><% if lst_indicadores_carrega = "" then %><img id="exclamacao" src="../IMAGEM/exclamacao_14x14.png" title="Reduza o tempo de carregamento da lista de indicadores, filtrando por vendedor." style="cursor:pointer;" />&nbsp;<% end if %>Indicador</span></td>
		<td align="left">
			<select id="c_indicador" name="c_indicador" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" <% if lst_indicadores_carrega = "" then %> onfocus="CarregaListaIndicadores();" <% end if %>>
			    <option selected value=''>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
			<% if lst_indicadores_carrega <> "" then
			    Response.Write indicadores_monta_itens_select(Null, False)
			   end if
			 %>
			</select><br />
			<span class="aviso">Vendedor selecionado não possui indicadores.</span>&nbsp;
		</td>
	</tr>
	</table>
</td></tr>

<!--  NOVA VERSÃO DA FORMA DE PAGAMENTO  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">FORMA DE PAGAMENTO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<span class="C" style="margin-left:30px;">Forma de Pagamento</span>
			<select id="op_forma_pagto" name="op_forma_pagto">
				<% =forma_pagto_monta_itens_select(Null) %>
			</select>
		</td></tr>
	<tr bgColor="#FFFFFF"><td>
		<span class="C" style="margin-left:30px;">Nº Parcelas</span>
			<input class="Cc" maxlength="2" style="width:40px;" name="c_forma_pagto_qtde_parc" id="c_forma_pagto_qtde_parc" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); filtra_numerico();">
		</td></tr>
	</table>
</td></tr>

<!--  Nº ORÇAMENTO  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">ORÇAMENTO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<span class="C" style="margin-left:30px;">Nº Orçamento</span>
			<input class="C" maxlength="10" style="width:70px;" name="c_orcamento" id="c_orcamento" onblur="if (normaliza_num_orcamento(this.value)!='') this.value=normaliza_num_orcamento(this.value);" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {if (normaliza_num_orcamento(this.value)!='') this.value=normaliza_num_orcamento(this.value); bCONFIRMA.focus();} filtra_orcamento();">
		</td></tr>
	</table>
</td></tr>

<!--  CLIENTE  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">CLIENTE</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<span class="C" style="margin-left:30px;">CNPJ/CPF</span>
			<input class="C" maxlength="18" style="width:140px;" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="if (digitou_enter(true)&&((!tem_info(this.value))||(tem_info(this.value)&&cnpj_cpf_ok(this.value)))) {this.value=cnpj_cpf_formata(this.value); bCONFIRMA.focus();} filtra_cnpj_cpf();">
		</td></tr>
	</table>
</td></tr>

</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()">
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