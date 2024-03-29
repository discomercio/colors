<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================
'	  R E L E S T O Q U E V E N D A C M V P V . A S P
'     =================================================
'     A p�gina foi renomeada em 24/01/2018, anteriormente chamava-se RelPosicaoEstoqueCmvPv.asp
'     Este relat�rio foi duplicado da Loja p/ a Central, mas como na Central j� havia uma p�gina c/ o mesmo nome, optou-se por renomear p/ que este relat�rio mantivesse o mesmo nome na Loja e na Central.
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

	const ID_RELATORIO = "LOJA/RelEstoqueVendaCmvPv"

	dim usuario, loja, s
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISS�O DE ACESSO DO USU�RIO
	if Not ( _
			operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_SINTETICO_CMVPV, s_lista_operacoes_permitidas) Or _
			operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO_CMVPV, s_lista_operacoes_permitidas) _
			) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

    dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
    If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_sessionToken
	s_sessionToken = ""
	s = "SELECT Convert(varchar(36), SessionTokenModuloLoja) AS SessionTokenModuloLoja FROM t_USUARIO WHERE (usuario = '" & usuario & "')"
    if rs.State <> 0 then rs.Close
    rs.Open s, cn
	if Not rs.Eof then s_sessionToken = Trim("" & rs("SessionTokenModuloLoja"))
	if rs.State <> 0 then rs.Close






' _____________________________________________________________________________________________
'
'									F  U  N  �  �  E  S 
' _____________________________________________________________________________________________


' ____________________________________________________________________________
' FABRICANTE MONTA ITENS SELECT
'
function fabricante_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, i
dim v
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" Coalesce(t_PRODUTO.fabricante,'') AS fabricante" & _
                " ,Coalesce(nome,'') AS nome" & _
			" FROM t_PRODUTO" & _
            " INNER JOIN t_FABRICANTE ON (t_PRODUTO.fabricante=t_FABRICANTE.fabricante)" & _
			" WHERE" & _
				" (Coalesce(t_PRODUTO.fabricante,'') <> '')" & _
			" ORDER BY" & _
				" Coalesce(t_PRODUTO.fabricante,'')"
	set r = cn.Execute(strSql)
	strResp = ""
  
	do while Not r.eof 
	    
		x = Trim("" & r("fabricante"))
		strResp = strResp & "<option "
            for i=LBound(v) to UBound(v) 
		        if (id_default<>"") And (v(i)=x) then
		            strResp = strResp & "selected"
		         end if
		   	 next
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("fabricante")) & " &nbsp;(" & Trim("" & r("nome")) & ")"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext	
 	loop

	fabricante_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' ____________________________________________________________________________
' GRUPO MONTA ITENS SELECT
'
function t_produto_grupo_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, v, i
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "SELECT" & _
				" codigo," & _
                " descricao" & _
			" FROM t_PRODUTO_GRUPO" & _
			" WHERE" & _
				" (Coalesce(codigo,'') <> '')" & _
				" AND (inativo = 0)" & _
			" ORDER BY" & _
				" Coalesce(codigo,'')"
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
		strResp = strResp & Trim("" & r("codigo")) & " &nbsp;(" & Trim("" & r("descricao")) & ")"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext	
 	loop
		
	t_produto_grupo_monta_itens_select = strResp
	r.close
	set r=nothing
end function


'----------------------------------------------------------------------------------------------
' T_PRODUTO SUBGRUPO MONTA ITENS SELECT
function t_produto_subgrupo_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default, v, i, sDescricao
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "SELECT DISTINCT tP.subgrupo, tPS.descricao FROM t_PRODUTO tP LEFT JOIN t_PRODUTO_SUBGRUPO tPS ON (tP.subgrupo = tPS.codigo) WHERE LEN(Coalesce(tP.subgrupo,'')) > 0 ORDER by tP.subgrupo"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("subgrupo")))
		sDescricao = Trim("" & r("descricao"))
		strResp = strResp & "<option "
		for i=LBound(v) to UBound(v) 
			if (id_default<>"") And (v(i)=x) then
				strResp = strResp & "selected"
				end if
			next
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x
		if sDescricao <> "" then strResp = strResp & " &nbsp;(" & sDescricao & ")"
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop
	
	t_produto_subgrupo_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' ____________________________________________________________________________
' POTENCIA BTU MONTA ITENS SELECT
'
function potencia_BTU_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" potencia_BTU" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (potencia_BTU <> 0)" & _
			" ORDER BY" & _
				" potencia_BTU"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("potencia_BTU"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & formata_inteiro(r("potencia_BTU")) & "&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		
	potencia_BTU_monta_itens_select = strResp
	r.close
	set r=nothing
end function


' ____________________________________________________________________________
' CICLO MONTA ITENS SELECT
'
function ciclo_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" Coalesce(ciclo,'') AS ciclo" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (Coalesce(ciclo,'') <> '')" & _
			" ORDER BY" & _
				" Coalesce(ciclo,'')"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("ciclo")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & UCase(Trim("" & r("ciclo"))) & "&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		
	ciclo_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________
' POSICAO MERCADO MONTA ITENS SELECT
'
function posicao_mercado_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" Coalesce(posicao_mercado,'') AS posicao_mercado" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (Coalesce(posicao_mercado,'') <> '')" & _
			" ORDER BY" & _
				" Coalesce(posicao_mercado,'')"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("posicao_mercado")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & UCase(Trim("" & r("posicao_mercado"))) & "&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		
	posicao_mercado_monta_itens_select = strResp
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


<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
    $(function () {
        $("#c_fabricante_multiplo").change(function () {
            $("#spnCounterFabricante").text($("#c_fabricante_multiplo :selected").length);
        });

        $("#c_grupo").change(function () {
            $("#spnCounterGrupo").text($("#c_grupo :selected").length);
        });

        $("#c_subgrupo").change(function () {
            $("#spnCounterSubgrupo").text($("#c_subgrupo :selected").length);
        });

        $("#spnCounterFabricante").text($("#c_fabricante_multiplo :selected").length);
        $("#spnCounterGrupo").text($("#c_grupo :selected").length);
        $("#spnCounterSubgrupo").text($("#c_subgrupo :selected").length);

		$("#divAjaxRunning").hide();

		$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPAR�NCIA NO IE8

		$(document).ajaxStart(function () {
			$("#divAjaxRunning").show();
		})
			.ajaxStop(function () {
				$("#divAjaxRunning").hide();
			});

		//Every resize of window
		$(window).resize(function () {
			sizeDivAjaxRunning();
		});

		//Every scroll of window
		$(window).scroll(function () {
			sizeDivAjaxRunning();
		});
	});

//Dynamically assign height
function sizeDivAjaxRunning() {
	var newTop = $(window).scrollTop() + "px";
	$("#divAjaxRunning").css("top", newTop);
}

	function geraArquivoXLS(f) {
		var estoque, detalhe, consolidacao_codigos;
		var fabricante, produto, empresa, potencia_BTU, ciclo, posicao_mercado;
		var fabricante_multiplo, grupo, subgrupo;
		var usuario, sessionToken;
		var serverVariableUrl, strUrl, strUrlDownload;

		if (!consisteCamposFiltro(f)) return;

		estoque = $("input[name='rb_estoque']:checked").val();
		detalhe = $("input[name='rb_detalhe']:checked").val();
		consolidacao_codigos = $("input[name='rb_exportacao']:checked").val();

		fabricante = trim($("#c_fabricante").val());
		produto = trim($("#c_produto").val());
		empresa = trim($("#c_empresa").val());
		potencia_BTU = trim($("#c_potencia_BTU").val());
		ciclo = trim($("#c_ciclo").val());
		posicao_mercado = trim($("#c_posicao_mercado").val());

		fabricante_multiplo = "";
		for (i = 0; i < f.c_fabricante_multiplo.length; i++) {
			if (f.c_fabricante_multiplo[i].selected) {
				if (fabricante_multiplo != "") fabricante_multiplo += "_";
				fabricante_multiplo += f.c_fabricante_multiplo[i].value;
			}
		}

		grupo = "";
		for (i = 0; i < f.c_grupo.length; i++) {
			if (f.c_grupo[i].selected) {
				if (grupo != "") grupo += "_";
				grupo += f.c_grupo[i].value;
			}
		}

		subgrupo = "";
		for (i = 0; i < f.c_subgrupo.length; i++) {
			if (f.c_subgrupo[i].selected) {
				if (subgrupo != "") subgrupo += "_";
				subgrupo += f.c_subgrupo[i].value;
			}
		}

		usuario = "<%=usuario%>";
		sessionToken = $("#sessionToken").val();
		loja = trim($("#c_loja").val());

		serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
		serverVariableUrl = serverVariableUrl.toUpperCase();
		serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("LOJA"));

		strUrl = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/RelEstoqueVendaXLS/';
		strUrlDownload = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/DownloadRelEstoqueVendaXLS/';

		$("#divAjaxRunning").show();
		var jqxhr = $.ajax({
			url: strUrl,
			type: "GET",
			async: true,
			dataType: "json",
			data: {
				usuario: usuario,
				loja: loja,
				sessionToken: sessionToken,
				filtro_estoque: estoque,
				filtro_detalhe: detalhe,
				filtro_consolidacao_codigos: consolidacao_codigos,
				filtro_empresa: empresa,
				filtro_fabricante: fabricante,
				filtro_produto: produto,
				filtro_fabricante_multiplo: fabricante_multiplo,
				filtro_grupo: grupo,
				filtro_subgrupo: subgrupo,
				filtro_potencia_BTU: potencia_BTU,
				filtro_ciclo: ciclo,
				filtro_posicao_mercado: posicao_mercado
			}
		})
			.success(function (response) {
				$("#divAjaxRunning").hide();
				if (response.Status == "OK") {
					fDOWNLOAD.action = strUrlDownload + "?fileName=" + response.fileName;
					fDOWNLOAD.submit();
				}
				else if (response.Status == "Falha") {
					alert("Falha ao gerar o arquivo XLS!\n" + response.Exception);
				}
				else if (response.Status == "Vazio") {
					alert("Nenhum registro encontrado!");
				}
			})
			.fail(function (jqXHR, textStatus) {
				$("#divAjaxRunning").hide();
				var msgErro = "";
				if (textStatus.toString().length > 0) msgErro = "Mensagem de Status: " + textStatus.toString();
				try {
					if (jqXHR.status.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Status: " + jqXHR.status.toString(); }
				} catch (e) { }

				try {
					if (jqXHR.statusText.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Descri��o do Status: " + jqXHR.statusText.toString(); }
				} catch (e) { }

				try {
					if (jqXHR.responseText.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Mensagem de Resposta: " + jqXHR.responseText.toString(); }
				} catch (e) { }

				alert("Falha ao tentar processar a consulta!!\n\n" + msgErro);
			});
	}

function consisteCamposFiltro(f) {
var i, b;
	b = false;
	for (i = 0; i < f.rb_detalhe.length; i++) {
		if (f.rb_detalhe[i].checked) {
			b = true;
			break;
		}
	}
	if (!b) {
		alert("Selecione o tipo de detalhamento da consulta!!");
		return false;
	}

	if (trim(f.c_produto.value) != "") {
		if (!isEAN(trim(f.c_produto.value))) {
			if (trim(f.c_fabricante.value) == "") {
				alert("Informe o fabricante do produto!!");
				f.c_fabricante.focus();
				return false;
			}
		}
	}

	return true;
}

function fESTOQConsulta(f) {
	if (!consisteCamposFiltro(f)) return;

	if (f.rb_saida[1].checked) {
		geraArquivoXLS(f);
	}
	else {
		dCONFIRMA.style.visibility = "hidden";
		window.status = "Aguarde ...";
		f.submit();
	}
}

function limpaCampoSelect(c) {
	c.options[0].selected = true;
}
function limpaCampoSelectFabricante() {
    $("#c_fabricante_multiplo").children().prop("selected", false);
    $("#spnCounterFabricante").text($("#c_fabricante_multiplo :selected").length);
}
function limpaCampoSelectProduto() {
    $("#c_grupo").children().prop("selected", false);
    $("#spnCounterGrupo").text($("#c_grupo :selected").length);
}
function limpaCampoSelectSubgrupo() {
    $("#c_subgrupo").children().prop('selected', false);
    $("#spnCounterSubgrupo").text($("#c_subgrupo :selected").length);
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

<style TYPE="text/css">
#divAjaxRunning
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	z-index:1001;
	background-color:grey;
	opacity: .6;
}
.AjaxImgLoader
{
	position: absolute;
	left: 50%;
	top: 50%;
	margin-left: -128px; /* -1 * image width / 2 */
	margin-top: -128px;  /* -1 * image height / 2 */
	display: block;
}
#rb_estoque {
	margin: 0px 0px 0px 20px;
	vertical-align: top;
	}
#rb_detalhe {
	margin: 0px 0px 0px 20px;
	vertical-align: top;
	}
.rbOpt {
	margin-left:20px;
}
.LST
{
	margin:6px 6px 6px 6px;
}
</style>


<body onload="if (trim(fESTOQ.c_fabricante.value)=='') fESTOQ.c_fabricante.focus();">

<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>

<center>

<form id="fESTOQ" name="fESTOQ" method="post" action="RelEstoqueVendaCmvPvExec.asp">
<input type="hidden" name="sessionToken" id="sessionToken" value="<%=s_sessionToken%>" />
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_loja" id="c_loja" value="<%=loja%>" />

<!--  I D E N T I F I C A � � O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque de Venda</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PAR�METROS DA CONSULTA  -->
<table class="Qx" cellSpacing="0">
<!--  ESTOQUE  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MT" NOWRAP><span class="PLTe">Estoque de Interesse</span>
		<br><input type="radio" checked tabindex="-1" id="rb_estoque" name="rb_estoque" value="<%=ID_ESTOQUE_VENDA%>">
			<span class="C" style="cursor:default">Venda</span>
	</td>
	</tr>

<!-- EMPRESA -->
    <tr bgcolor="#FFFFFF">
		<td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Empresa</span>
		<br>
			<select id="c_empresa" name="c_empresa" style="margin:1px 3px 6px 10px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
			</select>
			
        </td>
    </tr>

<!--  TIPO DE DETALHAMENTO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MDBE" NOWRAP><span class="PLTe">Tipo de Detalhamento</span>
		<%
			s=" disabled" 
			if operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO_CMVPV, s_lista_operacoes_permitidas) then s=""
			if operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_SINTETICO_CMVPV, s_lista_operacoes_permitidas) And (Not operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO_CMVPV, s_lista_operacoes_permitidas)) then s=" checked"
		%>
		<br><input type="radio" <%=s%> tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="SINTETICO" <% if get_default_valor_texto_bd(usuario, ID_RELATORIO & "|" & "rb_detalhe") = "SINTETICO" then Response.Write " checked" %>>
			<span class="C" style="cursor:default" onclick="fESTOQ.rb_detalhe[0].click();">Sint�tico (sem custos)</span>
			
		<%	if operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO_CMVPV, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
		<br><input type="radio" <%=s%> tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="INTERMEDIARIO" <% if get_default_valor_texto_bd(usuario, ID_RELATORIO & "|" & "rb_detalhe") = "INTERMEDIARIO" then Response.Write " checked" %>>
			<span class="C" style="cursor:default" onclick="fESTOQ.rb_detalhe[1].click();">Intermedi�rio (custos m�dios)</span>
	</td>
	</tr>

<!--  FABRICANTE/PRODUTO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">Fabricante</span>
		<br><input name="c_fabricante" id="c_fabricante" class="PLLe" maxlength="4" style="margin-left:2pt;width:50px;" onkeypress="if (digitou_enter(true)) fESTOQ.c_produto.focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"></td>
	<td class="MDB" style="border-left:0pt;"><span class="PLTe">Produto</span>
		<br><input name="c_produto" id="c_produto" class="PLLe" maxlength="13" style="margin-left:2pt;width:100px;" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_produto();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_PRODUTO); this.value=ucase(trim(this.value));"></td>
	</tr>

<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" colspan="2" align="left" nowrap>
		<span class="PLTe">Fabricantes</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_fabricante_multiplo" name="c_fabricante_multiplo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="5"style="width:250px" multiple>
			<% =fabricante_monta_itens_select(get_default_valor_texto_bd(usuario, ID_RELATORIO & "|" & "c_fabricante_multiplo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparFabricante" id="bLimparFabricante" href="javascript:limpaCampoSelectFabricante()" title="limpa o filtro 'Fabricante'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
                        <br />
                        (<span class="Lbl" id="spnCounterFabricante"></span>)
		</td>
		</tr>
		</table>
	</td>
	</tr>
<!-- GRUPO DE PRODUTOS -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" colspan="2" align="left" nowrap>
		<span class="PLTe">Grupo de Produtos</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_grupo" name="c_grupo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="5"style="width:200px" multiple>
			<% =t_produto_grupo_monta_itens_select(get_default_valor_texto_bd(usuario, ID_RELATORIO & "|" & "c_grupo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparGrupo" id="bLimparGrupo" href="javascript:limpaCampoSelectProduto()" title="limpa o filtro 'Grupo de Produtos'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
                        <br />
                        (<span class="Lbl" id="spnCounterGrupo"></span>)
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<!-- SUBGRUPO DE PRODUTOS -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" colspan="2" align="left" nowrap>
		<span class="PLTe">Subgrupo de Produtos</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_subgrupo" name="c_subgrupo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="6" style="min-width:250px" multiple>
			<% =t_produto_subgrupo_monta_itens_select(get_default_valor_texto_bd(usuario, ID_RELATORIO & "|" & "c_subgrupo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparSubgrupo" id="bLimparSubgrupo" href="javascript:limpaCampoSelectSubgrupo()" title="limpa o filtro 'Subgrupo de Produtos'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
                        <br />
                        (<span class="Lbl" id="spnCounterSubgrupo"></span>)
		</td>
		</tr>
		</table>
	</td>
	</tr>
<!-- BTU/h -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" colspan="2" align="left" nowrap>
		<span class="PLTe">BTU/H</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_potencia_BTU" name="c_potencia_BTU" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =potencia_BTU_monta_itens_select(Null) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparPotenciaBTU" id="bLimparPotenciaBTU" href="javascript:limpaCampoSelect(fESTOQ.c_potencia_BTU)" title="limpa o filtro 'BTU/h'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
<!-- CICLO -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" colspan="2" align="left" nowrap>
		<span class="PLTe">Ciclo</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_ciclo" name="c_ciclo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =ciclo_monta_itens_select(Null) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparCiclo" id="bLimparCiclo" href="javascript:limpaCampoSelect(fESTOQ.c_ciclo)" title="limpa o filtro 'Ciclo'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
<!-- POSI��O MERCADO -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" colspan="2" align="left" nowrap>
		<span class="PLTe">Posi��o Mercado</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_posicao_mercado" name="c_posicao_mercado" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =posicao_mercado_monta_itens_select(Null) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparPosicaoMercado" id="bLimparPosicaoMercado" href="javascript:limpaCampoSelect(fESTOQ.c_posicao_mercado)" title="limpa o filtro 'Posi��o Mercado'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
<!--  OP��ES DE CONSULTA  -->
    <tr bgcolor="#FFFFFF">
	<td colspan="2" class="MDBE" nowrap><span class="PLTe">Op��es de Consulta</span>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_exportacao" name="rb_exportacao" value="Normais"<% if get_default_valor_texto_bd(usuario, ID_RELATORIO & "|" & "rb_exportacao") = "Normais" then Response.Write " checked" %>>
			<span class="C" style="cursor:default" onclick="fESTOQ.rb_exportacao[0].click();" >C�digos normais</span>			
        	
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_exportacao" name="rb_exportacao" value="Compostos" <% if get_default_valor_texto_bd(usuario, ID_RELATORIO & "|" & "rb_exportacao") = "Compostos" then Response.Write " checked" %>>
			<span class="C" style="cursor:default" onclick="fESTOQ.rb_exportacao[1].click();">C�digos unificados</span>
	</td>
	</tr>

<!--  SA�DA DO RELAT�RIO  -->
	<tr bgColor="#FFFFFF" NOWRAP>
		<td colspan="2" class="ME MB MD">
		<span class="PLTe">Sa�da</span>
		<br />
		<input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="Html" <% if get_default_valor_texto_bd(usuario, ID_RELATORIO & "|" & "rb_saida") = "Html" then Response.Write " checked" %>><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_saida[0].click();"
			>Html</span>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="XLS" <% if get_default_valor_texto_bd(usuario, ID_RELATORIO & "|" & "rb_saida") = "XLS" then Response.Write " checked" %>><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_saida[1].click();"
			>Excel</span>
		</td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a opera��o">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fESTOQConsulta(fESTOQ)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>

<form method="POST" name="fDOWNLOAD" id="fDOWNLOAD">
<input type="hidden" name="usuario" value="<%=usuario%>" />
<input type="hidden" name="loja" value="" />
<input type="hidden" name="sessionToken" value="<%=s_sessionToken%>" />
<input type="hidden" name="fileName" />
</form>

</body>
</html>
