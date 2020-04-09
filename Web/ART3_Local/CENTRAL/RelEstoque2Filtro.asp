<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  RelEstoque2Filtro.asp
'     ===============================================
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

	Const COD_TIPO_AGRUPAMENTO__GRUPO = "Grupo"
	Const COD_TIPO_AGRUPAMENTO__PRODUTO = "Produto"

	dim usuario, s, intIdx
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if (Not operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas)) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

    '	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
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

<script language="JavaScript" type="text/javascript">
function fESTOQConsulta( f ) {
var i, b;
	b=false;
	for (i=0; i<f.rb_estoque.length; i++) {
		if (f.rb_estoque[i].checked) {
			b=true;
			break;
			}
		}
	if (!b) {
		alert("Selecione o estoque a ser consultado!!");
		return;
		}

	b=false;
	for (i=0; i<f.rb_detalhe.length; i++) {
		if (f.rb_detalhe[i].checked) {
			b=true;
			break;
			}
		}
	if (!b) {
		alert("Selecione o tipo de detalhamento da consulta!!");
		return;
		}

	b = false;
	for (i = 0; i < f.rb_saida.length; i++) {
		if (f.rb_saida[i].checked) {
			b = true;
			break;
		}
	}
	if (!b) {
		alert("Selecione o tipo de saída do relatório!!");
		return;
	}

	if (trim(f.c_produto.value)!="") {
		if (!isEAN(trim(f.c_produto.value))) {
			if (trim(f.c_fabricante.value)=="") {
				alert("Informe o fabricante do produto!!");
				f.c_fabricante.focus();
				return;
				}
			}
		}
	
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";

	if (f.rb_saida[1].checked) setTimeout('exibe_botao_confirmar()', 10000);

	f.submit();
}

function exibe_botao_confirmar() {
	dCONFIRMA.style.visibility = "";
	window.status = "";
}

function alternaTipoDetalhamento() {
    if($(".venda_show_room").is(":checked")) {
        $(".detalhamento_completo").attr("disabled", true);
        $(".detalhamento_completo").attr("checked", false);
    }
    else {
        $(".detalhamento_completo").attr("disabled", false);
    }
}
$(function () {
    alternaTipoDetalhamento();

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
});
</script>
<script type="text/javascript">
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
<script type="text/javascript">
    function alternaTipoAgrupamento() {
        if ($(".rbOptSemSaidaPorGrupo").is(":checked")) {
            $("#rb_tipo_agrupamento_por_produto").prop("checked", true);
            $("#div_tipo_agrupamento_por_grupo").hide();
        }
        else {
            $("#div_tipo_agrupamento_por_grupo").show();
        }
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

<style type="text/css">
#rb_estoque {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
#rb_detalhe {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
#rb_saida {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
.rb_tipo_agrupamento {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
.rbOpt
{
	vertical-align:bottom;
}
.lblOpt
{
	vertical-align:bottom;
}
</style>


<body onload="if (trim(fESTOQ.c_fabricante.value)=='') fESTOQ.c_fabricante.focus(); alternaTipoAgrupamento()"">
<center>

<form id="fESTOQ" name="fESTOQ" method="post" action="RelEstoque2Exec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque II</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS DA CONSULTA  -->
<table class="Qx" cellspacing="0">
<!--  ESTOQUE  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="2" class="MT" nowrap><span class="PLTe">Estoque de Interesse</span>
		<% intIdx = -1 %>
		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_estoque" onchange="alternaTipoAgrupamento()" name="rb_estoque" value="<%=ID_ESTOQUE_VENDA%>" onclick="alternaTipoDetalhamento();"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[<%=Cstr(intIdx)%>].click();"
			>Venda</span>
			
		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_estoque" name="rb_estoque" onchange="alternaTipoAgrupamento()" value="<%=ID_ESTOQUE_VENDIDO%>" onclick="alternaTipoDetalhamento();"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[<%=Cstr(intIdx)%>].click();"
			>Vendido</span>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_estoque" name="rb_estoque" onchange="alternaTipoAgrupamento()" value="<%=ID_ESTOQUE_SHOW_ROOM%>" onclick="alternaTipoDetalhamento();"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[<%=Cstr(intIdx)%>].click();"
			>Show-Room</span>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt rbOptSemSaidaPorGrupo" tabindex="-1" id="rb_estoque" name="rb_estoque" onchange="alternaTipoAgrupamento()" value="<%=ID_ESTOQUE_DANIFICADOS%>" onclick="alternaTipoDetalhamento();"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[<%=Cstr(intIdx)%>].click();"
			>Produtos Danificados</span>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt rbOptSemSaidaPorGrupo" tabindex="-1" id="rb_estoque" name="rb_estoque" onchange="alternaTipoAgrupamento()" value="<%=ID_ESTOQUE_DEVOLUCAO%>" onclick="alternaTipoDetalhamento();"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[<%=Cstr(intIdx)%>].click();"
			>Devolução</span>

        <% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt venda_show_room" tabindex="-1" id="rb_estoque" name="rb_estoque" onchange="alternaTipoAgrupamento()" value="VENDA_SHOW_ROOM" onclick="alternaTipoDetalhamento();"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[<%=Cstr(intIdx)%>].click();"
			>Venda + Show-Room</span>
	</td>
	</tr>

<!--  TIPO DE DETALHAMENTO  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="2" class="MDBE" nowrap><span class="PLTe">Tipo de Detalhamento</span>
		<% intIdx = -1 %>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="SINTETICO"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Sintético (sem custos)</span>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="INTERMEDIARIO"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Intermediário (custos médios)</span>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt detalhamento_completo" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="COMPLETO"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Completo (custos diferenciados)</span>
	</td>
	</tr>

<!--  TIPO DE AGRUPAMENTO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MDBE" NOWRAP><span class="PLTe">Tipo de Agrupamento</span>
		<div id="div_tipo_agrupamento_por_produto"><input type="radio" class="rbOpt rb_tipo_agrupamento" tabindex="-1" id="rb_tipo_agrupamento_por_produto" name="rb_tipo_agrupamento" value="<%=COD_TIPO_AGRUPAMENTO__PRODUTO%>" onclick="dCONFIRMA.style.visibility='';" checked><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_tipo_agrupamento[0].click();dCONFIRMA.style.visibility='';"
			>Produto</span>
        </div>
		<div id="div_tipo_agrupamento_por_grupo"><input type="radio" class="rbOpt rb_tipo_agrupamento" tabindex="-1" id="rb_tipo_agrupamento_por_grupo" name="rb_tipo_agrupamento" value="<%=COD_TIPO_AGRUPAMENTO__GRUPO%>" onclick="dCONFIRMA.style.visibility='';"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_tipo_agrupamento[1].click();dCONFIRMA.style.visibility='';"
			>Grupo de Produtos</span>
        </div>
	</td>
	</tr>

<!--  SAÍDA DO RELATÓRIO  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="2" class="MDBE" nowrap><span class="PLTe">Saída do Relatório</span>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="Html" onclick="dCONFIRMA.style.visibility='';" checked><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_saida[0].click();dCONFIRMA.style.visibility='';"
			>Html</span>

		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="XLS" onclick="dCONFIRMA.style.visibility='';"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_saida[1].click();dCONFIRMA.style.visibility='';"
			>Excel</span>
	</td>
	</tr>

<!-- EMPRESA -->
    <tr bgcolor="#FFFFFF">
        <td colspan="2" class="MDBE" NOWRAP><span class="PLTe">Empresa</span>
            <br>
			<select id="c_empresa" name="c_empresa" style="margin-right:15px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
			</select>
        </td>
    </tr>

<!--  FABRICANTE/PRODUTO  -->
	<tr bgcolor="#FFFFFF">
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
			<select id="c_fabricante_multiplo" name="c_fabricante_multiplo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="6"style="min-width:250px" multiple>
			<% =fabricante_monta_itens_select(get_default_valor_texto_bd(usuario, "RelEstoque2Filtro|c_fabricante_multiplo")) %>
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
			<select id="c_grupo" name="c_grupo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="6"style="min-width:250px" multiple>
			<% =t_produto_grupo_monta_itens_select(get_default_valor_texto_bd(usuario, "RelEstoque2Filtro|c_grupo")) %>
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
		<span class="PLTe">SUBGRUPO DE PRODUTOS</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_subgrupo" name="c_subgrupo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="6" style="min-width:250px" multiple>
			<% =t_produto_subgrupo_monta_itens_select(get_default_valor_texto_bd(usuario, "RelEstoque2Filtro|c_subgrupo")) %>
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
	<!-- POSIÇÃO MERCADO -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" colspan="2" align="left" nowrap>
		<span class="PLTe">Posição Mercado</span>
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
			<a name="bLimparPosicaoMercado" id="bLimparPosicaoMercado" href="javascript:limpaCampoSelect(fESTOQ.c_posicao_mercado)" title="limpa o filtro 'Posição Mercado'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>

</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fESTOQConsulta(fESTOQ)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>
