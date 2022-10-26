<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelTabelaDinamicaFiltro.asp
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

	const ID_RELATORIO = "RelTabelaDinamicaFiltro"

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim intIdx
	
	dim s_campos_saida_default, s_checked
	s_campos_saida_default = get_default_valor_texto_bd(usuario, ID_RELATORIO & "|campos_saida_selecionados")





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________________________________________________
' FABRICANTE MONTA ITENS SELECT
'
function fabricante_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, sSelected
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" Coalesce(fabricante,'') AS fabricante" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (Coalesce(fabricante,'') <> '')" & _
			" ORDER BY" & _
				" Coalesce(fabricante,'')"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("fabricante"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("fabricante")) & "&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if ha_default then
		sSelected = ""
	else
		sSelected = " selected"
		end if
	strResp = "<option" & sSelected & " value=''>&nbsp;</option>" & chr(13) & strResp
		
	fabricante_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________
' GRUPO MONTA ITENS SELECT
'
function grupo_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, sSelected
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" Coalesce(grupo,'') AS grupo" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (Coalesce(grupo,'') <> '')" & _
			" ORDER BY" & _
				" Coalesce(grupo,'')"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("grupo")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & UCase(Trim("" & r("grupo"))) & "&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if ha_default then
		sSelected = ""
	else
		sSelected = " selected"
		end if
	strResp = "<option" & sSelected & " value=''>&nbsp;</option>" & chr(13) & strResp
		
	grupo_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________
' POTENCIA BTU MONTA ITENS SELECT
'
function potencia_BTU_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, sSelected
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

	if ha_default then
		sSelected = ""
	else
		sSelected = " selected"
		end if
	strResp = "<option" & sSelected & " value=''>&nbsp;</option>" & chr(13) & strResp
		
	potencia_BTU_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________
' CICLO MONTA ITENS SELECT
'
function ciclo_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, sSelected
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

	if ha_default then
		sSelected = ""
	else
		sSelected = " selected"
		end if
	strResp = "<option" & sSelected & " value=''>&nbsp;</option>" & chr(13) & strResp
		
	ciclo_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________
' POSICAO MERCADO MONTA ITENS SELECT
'
function posicao_mercado_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, sSelected
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

	if ha_default then
		sSelected = ""
	else
		sSelected = " selected"
		end if
	strResp = "<option" & sSelected & " value=''>&nbsp;</option>" & chr(13) & strResp
		
	posicao_mercado_monta_itens_select = strResp
	r.close
	set r=nothing
end function

'----------------------------------------------------------------------------------------------
' grupo_origem_pedido_monta_itens_select
function grupo_origem_pedido_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo='PedidoECommerce_Origem_Grupo') AND (st_inativo=0) ORDER BY descricao")
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

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo='PedidoECommerce_Origem') AND (st_inativo=0) ORDER BY descricao")
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

' ____________________________________________________________________________
' ENTREGA IMEDIATA MONTA ITENS SELECT
'
function entrega_imediata_monta_itens_select(byval id_default)
dim x, strResp, ha_default, i, sOpcoes, vOpcoes, sDescricao, sSelected
	id_default = Trim("" & id_default)
	ha_default=False
	strResp = ""
	sOpcoes = COD_ETG_IMEDIATA_NAO & "|" & COD_ETG_IMEDIATA_SIM
	vOpcoes = Split(sOpcoes, "|")
	for i=LBound(vOpcoes) to Ubound(vOpcoes)
		x = Trim("" & vOpcoes(i))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		sDescricao = decodifica_etg_imediata(vOpcoes(i))
		strResp = strResp & sDescricao
		strResp = strResp & "</option>" & chr(13)
		next

	if ha_default then
		sSelected = ""
	else
		sSelected = " selected"
		end if
	strResp = "<option" & sSelected & " value=''>&nbsp;</option>" & chr(13) & strResp
		
	entrega_imediata_monta_itens_select = strResp
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
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
		$("input[type=radio]").hUtil('fix_radios');
		$("#c_dt_faturamento_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_faturamento_termino").hUtilUI('datepicker_peq_filtro_final');
		$("#c_dt_NF_venda_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_NF_venda_termino").hUtilUI('datepicker_peq_filtro_final');
		$("#c_dt_NF_remessa_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_NF_remessa_termino").hUtilUI('datepicker_peq_filtro_final');

        $("#c_grupo_pedido_origem").change(function () {
            $("#spnCounterGrupoOrigemPedido").text($("#c_grupo_pedido_origem :selected").length);
        });

        $("#spnCounterGrupoOrigemPedido").text($("#c_grupo_pedido_origem :selected").length);

		if ($("#ckb_CONSOLIDAR_PEDIDO").is(":checked")) {
			$(".SpnDETPROD").css("color", "red");
			$(".SpnDETPED").css("color", "black");
		}
		else {
			$(".SpnDETPROD").css("color", "black");
			$(".SpnDETPED").css("color", "red");
		}

		if ($("#ckb_COL_FRETE_DETALHADO").is(":checked")) {
			$("#ckb_COL_FRETE").prop("checked", true);
		}

		if (!$("#ckb_COL_FRETE").is(":checked")) {
			$("#ckb_COL_FRETE_DETALHADO").prop("checked", false);
		}

		$("#ckb_CONSOLIDAR_PEDIDO").change(function () {
			if ($(this).is(":checked")) {
				$(".SpnDETPROD").css("color", "red");
				$(".SpnDETPED").css("color", "black");
			}
			else {
				$(".SpnDETPROD").css("color", "black");
				$(".SpnDETPED").css("color", "red");
			}
		});

		$("#ckb_COL_FRETE").change(function () {
			if (!$(this).is(":checked")) {
				$("#ckb_COL_FRETE_DETALHADO").prop("checked", false);
			}
		});

		$("#ckb_COL_FRETE_DETALHADO").change(function () {
			if ($(this).is(":checked")) {
				$("#ckb_COL_FRETE").prop("checked", true);
			}
		});
	});
</script>

<script language="JavaScript" type="text/javascript">
function limpaCampoSelect(c) {
	c.options[0].selected = true;
}

function marcarDesmarcarCadastro() {
   
    if ($("#cadastro").is(":checked")) {
        $(".CKB_CADASTRO").prop("checked", true);
    }
    else {
        $(".CKB_CADASTRO").prop("checked", false);
    }
}

function marcarDesmarcarComercial() {

    if ($("#comercial").is(":checked")) {
        $(".CKB_COMERCIAL").prop("checked", true);
    }
    else {
        $(".CKB_COMERCIAL").prop("checked", false);
		$(".CKB_COMERCIAL_SUB_OPCAO").prop("checked", false);
    }
}

function marcarDesmarcarFinanceiro() {

    if ($("#financeiro").is(":checked")) {
        $(".CKB_FINANCEIRO").prop("checked", true);
    }
    else {
        $(".CKB_FINANCEIRO").prop("checked", false);
    }
}

function marcarTodos() {
    $(".CKB_CADASTRO, .CKB_COMERCIAL, .CKB_FINANCEIRO").each(function() {
        if (!$(this).is(":checked")) {
            $(this).trigger('click');
        }
    });
}

function desmarcarTodos() {
	$(".CKB_CADASTRO, .CKB_COMERCIAL, .CKB_COMERCIAL_SUB_OPCAO, .CKB_FINANCEIRO").each(function() {
        if ($(this).is(":checked")) {
            $(this).trigger('click');
        }
    });
}

function fFILTROConfirma( f ) {
var i_qtde_campos, bTemFiltroPeriodo, msg_erro_consistencia;

	bTemFiltroPeriodo = false;
	msg_erro_consistencia = "";

	//  PERÍODO DE FATURAMENTO
	if ((trim(f.c_dt_faturamento_inicio.value) != "") || (trim(f.c_dt_faturamento_termino.value) != "")) {
		if (trim(f.c_dt_faturamento_inicio.value) == "") {
			alert("Informe a data de início do período!!");
			f.c_dt_faturamento_inicio.focus();
			return;
		}

		if (trim(f.c_dt_faturamento_termino.value) == "") {
			alert("Informe a data de término do período!!");
			f.c_dt_faturamento_termino.focus();
			return;
		}

		if (!consiste_periodo(f.c_dt_faturamento_inicio, f.c_dt_faturamento_termino)) return;
		bTemFiltroPeriodo = true;
	}

	// PERÍODO NF VENDA
	if ((trim(f.c_dt_NF_venda_inicio.value) != "") || (trim(f.c_dt_NF_venda_termino.value) != "")) {
		if (trim(f.c_dt_NF_venda_inicio.value) == "") {
			alert("Informe a data de início do período!!");
			f.c_dt_NF_venda_inicio.focus();
			return;
		}

		if (trim(f.c_dt_NF_venda_termino.value) == "") {
			alert("Informe a data de término do período!!");
			f.c_dt_NF_venda_termino.focus();
			return;
		}

		if (!consiste_periodo(f.c_dt_NF_venda_inicio, f.c_dt_NF_venda_termino)) return;
		bTemFiltroPeriodo = true;
	}

	// PERÍODO NF REMESSA
	if ((trim(f.c_dt_NF_remessa_inicio.value) != "") || (trim(f.c_dt_NF_remessa_termino.value) != "")) {
		if (trim(f.c_dt_NF_remessa_inicio.value) == "") {
			alert("Informe a data de início do período!!");
			f.c_dt_NF_remessa_inicio.focus();
			return;
		}

		if (trim(f.c_dt_NF_remessa_termino.value) == "") {
			alert("Informe a data de término do período!!");
			f.c_dt_NF_remessa_termino.focus();
			return;
		}

		if (!consiste_periodo(f.c_dt_NF_remessa_inicio, f.c_dt_NF_remessa_termino)) return;
		bTemFiltroPeriodo = true;
	}

	// ALGUM FILTRO POR PERÍODO DE DATAS FOI FORNECIDO?
	// O FILTRO POR STATUS DE ENTREGA IMEDIATA NÃO DEVE TER NENHUM FILTRO POR DATA PARA FUNCIONAR CORRETAMENTE, JÁ QUE OS PEDIDOS NESSE CONTEXTO AINDA NÃO ESTÃO COMO ENTREGUES
	if ((!bTemFiltroPeriodo) && (trim(f.c_entrega_imediata.value) == "")) {
		alert("É necessário preencher um dos períodos de data como filtro!!");
		return;
	}

	// O FILTRO DE STATUS DE ENTREGA IMEDIATA NÃO PODE OPERAR JUNTO COM OS CAMPOS DE SAÍDA 'VL Custo (Real)', 'VL Custo Total (Real)' E 'Nacional/Importado'
	// PORQUE O PROCESSAMENTO DESSES CAMPOS IMPLICA EM USAR AS TABELAS T_ESTOQUE_MOVIMENTO E T_ESTOQUE_ITEM, SENDO QUE NO CONTEXTO DA ENTREGA IMEDIATA
	// O PEDIDO AINDA NÃO ESTÁ ENTREGUE E PODE ESTAR ESPERANDO A CHEGADA DE PRODUTOS NO ESTOQUE
	if (trim(f.c_entrega_imediata.value) != "") {
		if ((trim(f.c_dt_faturamento_inicio.value) != "") || (trim(f.c_dt_faturamento_termino.value) != "")) {
			if (msg_erro_consistencia.length > 0) msg_erro_consistencia += "\n\n";
			msg_erro_consistencia += "O filtro de status de Entrega Imediata não pode operar em conjunto com o filtro 'PERÍODO (ENTREGUE)'!";
		}
		if (f.ckb_COL_VL_CUSTO_REAL.checked) {
			if (msg_erro_consistencia.length > 0) msg_erro_consistencia += "\n\n";
			msg_erro_consistencia += "O filtro de status de Entrega Imediata não pode operar em conjunto com o campo de saída 'VL Custo (Real)'!";
		}
		if (f.ckb_COL_VL_CUSTO_REAL_TOTAL.checked) {
			if (msg_erro_consistencia.length > 0) msg_erro_consistencia += "\n\n";
			msg_erro_consistencia += "O filtro de status de Entrega Imediata não pode operar em conjunto com o campo de saída 'VL Custo Total (Real)'!";
		}
		if (f.ckb_COL_NAC_IMP.checked) {
			if (msg_erro_consistencia.length > 0) msg_erro_consistencia += "\n\n";
			msg_erro_consistencia += "O filtro de status de Entrega Imediata não pode operar em conjunto com o campo de saída 'Nacional/Importado'!";
		}

		if (msg_erro_consistencia.length > 0) {
			alert(msg_erro_consistencia);
			return;
		}
	}

	//	CAMPOS DE SAÍDA
	i_qtde_campos = 0;
	$(".CKB_CADASTRO, .CKB_COMERCIAL, .CKB_FINANCEIRO").each(function() {
		if ($(this).is(":checked")) {
			i_qtde_campos++;
		}
	});

	if (i_qtde_campos == 0) {
		alert("Nenhum campo de saída foi assinalado!!");
		return;
	}
	
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";

	setTimeout('exibe_botao_confirmar()', 10000);

	f.submit();
}

function exibe_botao_confirmar() {
	dCONFIRMA.style.visibility = "";
	window.status = "";
}
</script>

<script type="text/javascript">
    function limpaCampoSelectGrupoOrigemPedido() {
        $("#c_grupo_pedido_origem").children().prop('selected', false);
        $("#spnCounterGrupoOrigemPedido").text($("#c_grupo_pedido_origem :selected").length);
    }
    function limpaCampoSelectOrigemPedido() {
        $("#c_pedido_origem").children().prop('selected', false);
        $("#spnCounterOrigemPedido").text($("#c_pedido_origem :selected").length);
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
.LST
{
	margin:6px 6px 6px 6px;
}
.tdColSaida
{
	width:48%;
}
</style>


<body onload="focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelTabelaDinamicaExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_usuario_sessao" id="c_usuario_sessao" value="<%=usuario%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Dados para Tabela Dinâmica</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellspacing="0" style="width: 400px">
<!--  PERÍODO (FATURAMENTO)  -->
	<tr bgcolor="#FFFFFF">
	<td class="MT" align="left" nowrap>
		<table cellspacing="0" cellpadding="0">
		<tr>
			<td align="left" valign="bottom">
				<span class="PLTe" style="cursor:default">PERÍODO (ENTREGUE)</span>
			</td>
		</tr>
		</table>
		<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF">
			<td align="left">
				<input class="PLLc" maxlength="10" style="width:76px;" name="c_dt_faturamento_inicio" id="c_dt_faturamento_inicio"
					onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
					onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_faturamento_termino.focus(); filtra_data();"
					value="<%=get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_faturamento_inicio")%>"
					>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;até&nbsp;</span>&nbsp;
					<input class="PLLc" maxlength="10" style="width:76px;" name="c_dt_faturamento_termino" id="c_dt_faturamento_termino"
					onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
					onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); filtra_data();"
					value="<%=get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_faturamento_termino")%>"
					/>
			</td>
			<td style="width:10px;"></td>
			<td align="left" valign="middle">
				<a name="bLimparPeriodoFaturamento" id="bLimparPeriodoFaturamento" href="javascript:limpaMultiplosCampos(fFILTRO.c_dt_faturamento_inicio,fFILTRO.c_dt_faturamento_termino);" title="limpa o filtro 'Período (Entregue)'">
							<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
			</td>
			</tr>
		</table>
	</td>
	</tr>
<!--  PERÍODO NF VENDA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		<table cellspacing="0" cellpadding="0">
		<tr>
			<td align="left" valign="bottom">
				<span class="PLTe" style="cursor:default">PERÍODO NF VENDA</span>
			</td>
		</tr>
		</table>
		<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF">
			<td align="left">
				<input class="PLLc" maxlength="10" style="width:76px;" name="c_dt_NF_venda_inicio" id="c_dt_NF_venda_inicio"
					onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
					onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_NF_venda_termino.focus(); filtra_data();"
					value="<%=get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_NF_venda_inicio")%>"
					>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;até&nbsp;</span>&nbsp;
					<input class="PLLc" maxlength="10" style="width:76px;" name="c_dt_NF_venda_termino" id="c_dt_NF_venda_termino"
					onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
					onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); filtra_data();"
					value="<%=get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_NF_venda_termino")%>"
					/>
			</td>
			<td style="width:10px;"></td>
			<td align="left" valign="middle">
				<a name="bLimparPeriodoNfVenda" id="bLimparPeriodoNfVenda" href="javascript:limpaMultiplosCampos(fFILTRO.c_dt_NF_venda_inicio,fFILTRO.c_dt_NF_venda_termino);" title="limpa o filtro 'Período NF Venda'">
							<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
			</td>
			</tr>
		</table>
	</td>
	</tr>
<!--  PERÍODO NF REMESSA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		<table cellspacing="0" cellpadding="0">
		<tr>
			<td align="left" valign="bottom">
				<span class="PLTe" style="cursor:default">PERÍODO NF REMESSA</span>
			</td>
		</tr>
		</table>
		<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF">
			<td align="left">
				<input class="PLLc" maxlength="10" style="width:76px;" name="c_dt_NF_remessa_inicio" id="c_dt_NF_remessa_inicio"
					onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
					onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_NF_remessa_termino.focus(); filtra_data();"
					value="<%=get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_NF_remessa_inicio")%>"
					>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;até&nbsp;</span>&nbsp;
					<input class="PLLc" maxlength="10" style="width:76px;" name="c_dt_NF_remessa_termino" id="c_dt_NF_remessa_termino"
					onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
					onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); filtra_data();"
					value="<%=get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_dt_NF_remessa_termino")%>"
					/>
			</td>
			<td style="width:10px;"></td>
			<td align="left" valign="middle">
				<a name="bLimparPeriodoNfRemessa" id="bLimparPeriodoNfRemessa" href="javascript:limpaMultiplosCampos(fFILTRO.c_dt_NF_remessa_inicio,fFILTRO.c_dt_NF_remessa_termino);" title="limpa o filtro 'Período NF Remessa'">
							<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
			</td>
			</tr>
		</table>
	</td>
	</tr>

<!--  PEDIDOS RECEBIDOS PELO CLIENTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">PEDIDOS RECEBIDOS PELO CLIENTE</span>
		<br>
		<table cellspacing="0" cellpadding="0" style="margin-left:30px;margin-bottom:4px;">
		<tr bgcolor="#FFFFFF"><td align="left">
			<input type="checkbox" tabindex="-1" id="ckb_pedido_nao_recebido_pelo_cliente" name="ckb_pedido_nao_recebido_pelo_cliente"
				value="ON"
				<% if get_default_valor_texto_bd(usuario, ID_RELATORIO & "|ckb_pedido_nao_recebido_pelo_cliente") <> "" then Response.Write " checked" %>
				><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_pedido_nao_recebido_pelo_cliente.click();">Não recebido pelo cliente</span>
			</td></tr>
		<tr bgcolor="#FFFFFF"><td align="left">
			<input type="checkbox" tabindex="-1" id="ckb_pedido_recebido_pelo_cliente" name="ckb_pedido_recebido_pelo_cliente"
				value="ON"
				<% if get_default_valor_texto_bd(usuario, ID_RELATORIO & "|ckb_pedido_recebido_pelo_cliente") <> "" then Response.Write " checked" %>
				><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_pedido_recebido_pelo_cliente.click();">Recebido pelo cliente</span>
			</td></tr>
		</table>
	</td></tr>

<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		<span class="PLTe">FABRICANTE</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_fabricante" name="c_fabricante" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =fabricante_monta_itens_select(get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_fabricante")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparFabricante" id="bLimparFabricante" href="javascript:limpaCampoSelect(fFILTRO.c_fabricante)" title="limpa o filtro 'Fabricante'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<!-- GRUPO DE PRODUTOS -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">GRUPO DE PRODUTOS</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_grupo" name="c_grupo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =grupo_monta_itens_select(get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_grupo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparGrupo" id="bLimparGrupo" href="javascript:limpaCampoSelect(fFILTRO.c_grupo)" title="limpa o filtro 'Grupo de Produtos'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<!-- BTU/h -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">BTU/H</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_potencia_BTU" name="c_potencia_BTU" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =potencia_BTU_monta_itens_select(get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_potencia_BTU")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparPotenciaBTU" id="bLimparPotenciaBTU" href="javascript:limpaCampoSelect(fFILTRO.c_potencia_BTU)" title="limpa o filtro 'BTU/h'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<!-- CICLO -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">CICLO</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_ciclo" name="c_ciclo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =ciclo_monta_itens_select(get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_ciclo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparCiclo" id="bLimparCiclo" href="javascript:limpaCampoSelect(fFILTRO.c_ciclo)" title="limpa o filtro 'Ciclo'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<!-- POSIÇÃO MERCADO -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">POSIÇÃO MERCADO</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_posicao_mercado" name="c_posicao_mercado" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =posicao_mercado_monta_itens_select(get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_posicao_mercado")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparPosicaoMercado" id="bLimparPosicaoMercado" href="javascript:limpaCampoSelect(fFILTRO.c_posicao_mercado)" title="limpa o filtro 'Posição Mercado'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
<!--  ENTREGA IMEDIATA  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">ENTREGA IMEDIATA</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_entrega_imediata" name="c_entrega_imediata" class="LST" style="min-width:70px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =entrega_imediata_monta_itens_select(get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_entrega_imediata")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparEntregaImediata" id="bLimparEntregaImediata" href="javascript:limpaCampoSelect(fFILTRO.c_entrega_imediata)" title="limpa o filtro 'Entrega Imediata'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
<!--  TIPO DE CLIENTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		<span class="PLTe">TIPO DE CLIENTE</span>
		<br>
		<% intIdx=-1 %>
		<input type="radio" id="rb_tipo_cliente" name="rb_tipo_cliente" value=<%=ID_PF%> style="margin-left:30px;">
		<% intIdx=intIdx+1 %>
		<span style="cursor:default" class="Np" onclick="fFILTRO.rb_tipo_cliente[<%=Cstr(intIdx)%>].click();">Pessoa Física</span>
		<br />
		<input type="radio" id="rb_tipo_cliente" name="rb_tipo_cliente" value=<%=ID_PJ%> style="margin-left:30px;">
		<% intIdx=intIdx+1 %>
		<span style="cursor:default" class="Np" onclick="fFILTRO.rb_tipo_cliente[<%=Cstr(intIdx)%>].click();">Pessoa Jurídica</span>
		<br />
		<input type="radio" id="rb_tipo_cliente" name="rb_tipo_cliente" value="" style="margin-left:30px;" checked>
		<% intIdx=intIdx+1 %>
		<span style="cursor:default" class="Np" onclick="fFILTRO.rb_tipo_cliente[<%=Cstr(intIdx)%>].click();">Ambos</span>
	</td>
	</tr>

<!--  LOJA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		
				<span class="PLTe">LOJA(S)</span>
				<br />
					<textarea class="PLBe" style="width:100px;font-size:9pt;margin-top:4px;margin-bottom:4px;margin-left: 7px" rows="8" name="c_loja" id="c_loja" onkeypress="if (!digitou_enter(false) && !digitou_char('-')) filtra_numerico();" onblur="this.value=normaliza_lista_lojas(this.value);"></textarea>			
	</td>
	</tr>

<!-- ORIGEM DO PEDIDO (GRUPO) -->
    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">ORIGEM DO PEDIDO (GRUPO)</span>
		<br>
		<table cellspacing="0" cellpadding="0"><tr bgcolor="#FFFFFF"><td align="left">            
			<select id="c_grupo_pedido_origem" name="c_grupo_pedido_origem" style="margin:1px 3px 6px 10px;width: 200px" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="5" multiple>
			<% =grupo_origem_pedido_monta_itens_select(Null) %>
			</select>
			</td>
        <td align="left" valign="top">
			<a href="javascript:limpaCampoSelectGrupoOrigemPedido()" title="limpa o filtro 'Origem do Pedido (Grupo)'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
                        <br />
                        (<span class="Lbl" id="spnCounterGrupoOrigemPedido"></span>)
		</td></tr></table>
        </td></tr>

<!-- PEDIDOS COM VALOR PAGO ATRAVÉS DE CARTÃO INTERNET -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		
				<span class="PLTe">CARTÃO (INTERNET)</span>
				<br />
				<% s_checked = ""
					if get_default_valor_texto_bd(usuario, ID_RELATORIO & "|ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET") = "ON" then s_checked = " checked"
				%>
				<input type="checkbox" tabindex="-1" id="ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET" name="ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET" class="DETPED"
						value="ON" <%=s_checked%> style="margin-left:30px;margin-bottom: 5px;margin-top: 5px;" /><span class="C SpnDETPED" style="cursor:default" onclick="fFILTRO.ckb_PEDIDOS_VL_PAGO_CARTAO_INTERNET.click();">Somente pedidos com valor pago através de cartão internet</span><br />
	</td>
	</tr>

<!--  AGRUPAMENTO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		
				<span class="PLTe">AGRUPAMENTO</span>
				<br />
				<input type="checkbox" tabindex="-1" id="ckb_AGRUPAMENTO" name="ckb_AGRUPAMENTO" class="DETPROD"
						value="ON" style="margin-left:30px;margin-bottom: 5px;margin-top: 5px;" /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_AGRUPAMENTO.click();">Desagrupar itens por quantidade</span><br />
	</td>
	</tr>

<!--  CONSOLIDAR POR PEDIDO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>

				<span class="PLTe">CONSOLIDAR POR PEDIDO</span>
				<br />
				<% s_checked = ""
					if get_default_valor_texto_bd(usuario, ID_RELATORIO & "|ckb_CONSOLIDAR_PEDIDO") = "ON" then s_checked = " checked"
				%>
				<input type="checkbox" tabindex="-1" id="ckb_CONSOLIDAR_PEDIDO" name="ckb_CONSOLIDAR_PEDIDO"
					value="ON" <%=s_checked%> style="margin-left:30px;margin-bottom: 5px;margin-top: 5px;" /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_CONSOLIDAR_PEDIDO.click();">Consolidar por pedido</span><img src="../IMAGEM/exclamacao_14x14.png" id="optConsolidarPedidoExclamacao" style="cursor:pointer" title="Esta opção impossibilita a exibição no resultado de dados específicos do produto" /><br />
	</td>
	</tr>

<!--  COMPATIBILIDADE DO EXCEL  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		
				<span class="PLTe">COMPATIBILIDADE DO EXCEL</span>
				<br />
				<% s_checked = ""
					if get_default_valor_texto_bd(usuario, ID_RELATORIO & "|ckb_COMPATIBILIDADE") = "ON" then s_checked = " checked"
				%>
				<input type="checkbox" tabindex="-1" id="ckb_COMPATIBILIDADE" name="ckb_COMPATIBILIDADE"
						value="ON" <%=s_checked%> style="margin-left:30px;margin-bottom: 5px;margin-top: 5px;" /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COMPATIBILIDADE.click();">Compatibilidade com versões anteriores do Excel</span><br />
	</td>
	</tr>

<!--  CAMPOS DE SAÍDA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		<span class="PLTe">CAMPOS DE SAÍDA</span>
		<br>
		<table width="100%" cellpadding="2" cellspacing="2">
			<tr>	
			    <td rowspan="2" class="tdColSaida" align="left" valign="top" style="margin-left:2px; margin-right:2px">	
			        <fieldset style="height:602px; border: solid 1px #555; padding: auto"><legend><input id="cadastro" type="checkbox" onclick="marcarDesmarcarCadastro()"/><label for="cadastro">Cadastro</label></legend>	   
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_DT_CADASTRO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_DT_CADASTRO" name="ckb_COL_DT_CADASTRO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_DT_CADASTRO.click();">Data (Cadastro)</span><br />
					
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_NF|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_NF" name="ckb_COL_NF"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_NF.click();">NF</span><br />
						
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_DT_EMISSAO_NF|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_DT_EMISSAO_NF" name="ckb_COL_DT_EMISSAO_NF"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_DT_EMISSAO_NF.click();">Data Emissão NF</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_NF_REMESSA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_NF_REMESSA" name="ckb_COL_NF_REMESSA"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_NF_REMESSA.click();">NF Remessa</span><br />
						
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_DT_EMISSAO_NF_REMESSA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_DT_EMISSAO_NF_REMESSA" name="ckb_COL_DT_EMISSAO_NF_REMESSA"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_DT_EMISSAO_NF_REMESSA.click();">Data Emissão NF Remessa</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_LOJA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_LOJA" name="ckb_COL_LOJA"
						        value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_LOJA.click();">Loja</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_PEDIDO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_PEDIDO" name="ckb_COL_PEDIDO"
						        value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_PEDIDO.click();">Pedido</span><br />
				
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_PEDIDO_MARKETPLACE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_PEDIDO_MARKETPLACE" name="ckb_COL_PEDIDO_MARKETPLACE"
						        value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_PEDIDO_MARKETPLACE.click();">Pedido Marketplace</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_GRUPO_PEDIDO_ORIGEM|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_GRUPO_PEDIDO_ORIGEM" name="ckb_COL_GRUPO_PEDIDO_ORIGEM"
						        value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_GRUPO_PEDIDO_ORIGEM.click();">Origem do Pedido (Grupo)</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_CPF_CNPJ_CLIENTE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_CPF_CNPJ_CLIENTE" name="ckb_COL_CPF_CNPJ_CLIENTE"
						value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_CPF_CNPJ_CLIENTE.click();">CPF/CNPJ Cliente</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_CONTRIBUINTE_ICMS|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_CONTRIBUINTE_ICMS" name="ckb_COL_CONTRIBUINTE_ICMS"
						value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_CONTRIBUINTE_ICMS.click();">Contribuinte ICMS</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_NOME_CLIENTE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_NOME_CLIENTE" name="ckb_COL_NOME_CLIENTE"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_NOME_CLIENTE.click();">Nome Cliente</span><br />
				
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_CIDADE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
						    <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_CIDADE" name="ckb_COL_CIDADE"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_CIDADE.click();">Cidade (cadastro)</span><br />
			
			            <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_UF|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_UF" name="ckb_COL_UF"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_UF.click();">UF (cadastro)</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_CIDADE_ETG|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
						    <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_CIDADE_ETG" name="ckb_COL_CIDADE_ETG"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_CIDADE_ETG.click();">Cidade (entrega)</span><br />
			
			            <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_UF_ETG|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_UF_ETG" name="ckb_COL_UF_ETG"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_UF_ETG.click();">UF (entrega)</span><br />

                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_TEL|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_TEL" name="ckb_COL_TEL"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_TEL.click();">Telefone</span><br />

                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_EMAIL|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_EMAIL" name="ckb_COL_EMAIL"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_EMAIL.click();">E-mail</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VENDEDOR|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_VENDEDOR" name="ckb_COL_VENDEDOR"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_VENDEDOR.click();">Vendedor</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_INDICADOR|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="Checkbox1" name="ckb_COL_INDICADOR"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_INDICADOR.click();">Indicador</span><br />
				
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_TRANSPORTADORA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_TRANSPORTADORA" name="ckb_COL_TRANSPORTADORA"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_TRANSPORTADORA.click();">Transportadora</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_ENTREGA_IMEDIATA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_ENTREGA_IMEDIATA" name="ckb_COL_ENTREGA_IMEDIATA"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_ENTREGA_IMEDIATA.click();">Entrega Imediata</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_DT_ENTREGA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_DT_ENTREGA" name="ckb_COL_DT_ENTREGA"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_DT_ENTREGA.click();">Data de Entrega</span><br />

                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_DT_PREVISAO_ETG_TRANSP|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_DT_PREVISAO_ETG_TRANSP" name="ckb_COL_DT_PREVISAO_ETG_TRANSP"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_DT_PREVISAO_ETG_TRANSP.click();">Previsão de Entrega da Transportadora</span><br />

                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_DT_RECEB_CLIENTE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_DT_RECEB_CLIENTE" name="ckb_COL_DT_RECEB_CLIENTE"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_DT_RECEB_CLIENTE.click();">Data de Recebimento pelo Cliente</span><br />
                        <hr />
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_INDICADOR_CPF_CNPJ|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_INDICADOR_CPF_CNPJ" name="ckb_COL_INDICADOR_CPF_CNPJ"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_INDICADOR_CPF_CNPJ.click();">CPF/CNPJ Indicador</span><br />
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_INDICADOR_ENDERECO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_INDICADOR_ENDERECO" name="ckb_COL_INDICADOR_ENDERECO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_INDICADOR_ENDERECO.click();">Endereço</span><br />
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_INDICADOR_CIDADE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_INDICADOR_CIDADE" name="ckb_COL_INDICADOR_CIDADE"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_INDICADOR_CIDADE.click();">Cidade</span><br />
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_INDICADOR_UF|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_INDICADOR_UF" name="ckb_COL_INDICADOR_UF"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_INDICADOR_UF.click();">UF</span><br />
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_INDICADOR_EMAILS|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_INDICADOR_EMAILS" name="ckb_COL_INDICADOR_EMAILS"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_INDICADOR_EMAILS.click();">E-mails</span><br />
					</fieldset>
				</td>
				<td class="tdColSaida" align="left" valign="middle" style="margin-left:2px; margin-right:2px">
                    <fieldset style="border: solid 1px #555; padding: auto"><legend><input id="comercial" type="checkbox" onclick="marcarDesmarcarComercial()" /><label for="comercial">Comercial</label></legend>
				
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_MARCA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
							<input type="checkbox" class="CKB_COMERCIAL DETPROD" tabindex="-1" id="ckb_COL_MARCA" name="ckb_COL_MARCA"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_MARCA.click();">Marca</span><br />
						
			            <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_GRUPO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL DETPROD" tabindex="-1" id="ckb_COL_GRUPO" name="ckb_COL_GRUPO"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_GRUPO.click();">Grupo</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_POTENCIA_BTU|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL DETPROD" tabindex="-1" id="ckb_COL_POTENCIA_BTU" name="ckb_COL_POTENCIA_BTU"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_POTENCIA_BTU.click();">BTU/h</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_CICLO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL DETPROD" tabindex="-1" id="ckb_COL_CICLO" name="ckb_COL_CICLO"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_CICLO.click();">Ciclo</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_POSICAO_MERCADO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL DETPROD" tabindex="-1" id="ckb_COL_POSICAO_MERCADO" name="ckb_COL_POSICAO_MERCADO"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_POSICAO_MERCADO.click();">Posição Mercado</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_PRODUTO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL DETPROD" tabindex="-1" id="ckb_COL_PRODUTO" name="ckb_COL_PRODUTO"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_PRODUTO.click();">Produto</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_NAC_IMP|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL DETPROD" tabindex="-1" id="ckb_COL_NAC_IMP" name="ckb_COL_NAC_IMP"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_NAC_IMP.click();">Nacional/Importado</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_DESCRICAO_PRODUTO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL DETPROD" tabindex="-1" id="ckb_COL_DESCRICAO_PRODUTO" name="ckb_COL_DESCRICAO_PRODUTO"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_DESCRICAO_PRODUTO.click();">Descrição Produto</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_QTDE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL DETPROD" tabindex="-1" id="ckb_COL_QTDE" name="ckb_COL_QTDE"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_QTDE.click();">Quantidade</span><br />

                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_PERC_DESC|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL DETPROD" tabindex="-1" id="ckb_COL_PERC_DESC" name="ckb_COL_PERC_DESC"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_PERC_DESC.click();">Percentual Desconto</span><br />
			
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_CUBAGEM|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_CUBAGEM" name="ckb_COL_CUBAGEM"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_CUBAGEM.click();">Cubagem</span><br />

                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_PESO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_PESO" name="ckb_COL_PESO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_PESO.click();">Peso</span><br />

                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_QTDE_VOLUMES|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_QTDE_VOLUMES" name="ckb_COL_QTDE_VOLUMES"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_QTDE_VOLUMES.click();">Qtde Volumes</span><br />

                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_FRETE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_FRETE" name="ckb_COL_FRETE"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_FRETE.click();">Valor Frete</span
							><span class="C" style="cursor:default">&nbsp(</span
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_FRETE_DETALHADO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
							><input type="checkbox" class="CKB_COMERCIAL_SUB_OPCAO" tabindex="-1" id="ckb_COL_FRETE_DETALHADO" name="ckb_COL_FRETE_DETALHADO" valign="bottom" style="padding:0px;margin:0px;"
							value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_FRETE_DETALHADO.click();">detalhado)</span><br />
					</fieldset>
				</td>
			</tr>
			<tr>
				<td class="tdColSaida" align="left" valign="middle" style="margin-left:2px; margin-right:2px">
				    <fieldset style="border: solid 1px #555;padding: auto"><legend><input id="financeiro" type="checkbox" onclick="marcarDesmarcarFinanceiro()" /><label for="financeiro">Financeiro</label></legend>
				    
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_CUSTO_ULT_ENTRADA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO DETPROD" tabindex="-1" id="ckb_COL_VL_CUSTO_ULT_ENTRADA" name="ckb_COL_VL_CUSTO_ULT_ENTRADA"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_CUSTO_ULT_ENTRADA.click();">VL Custo (Últ Entrada)</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_CUSTO_REAL|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO DETPROD" tabindex="-1" id="ckb_COL_VL_CUSTO_REAL" name="ckb_COL_VL_CUSTO_REAL"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_CUSTO_REAL.click();">VL Custo (Real)</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_LISTA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
			    	        <input type="checkbox" class="CKB_FINANCEIRO DETPROD" tabindex="-1" id="ckb_COL_VL_LISTA" name="ckb_COL_VL_LISTA"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_LISTA.click();">VL Lista</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_NF|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO DETPROD" tabindex="-1" id="ckb_COL_VL_NF" name="ckb_COL_VL_NF"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_NF.click();">VL NF</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_UNITARIO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO DETPROD" tabindex="-1" id="ckb_COL_VL_UNITARIO" name="ckb_COL_VL_UNITARIO"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPROD" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_UNITARIO.click();">VL Unitário</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_CUSTO_REAL_TOTAL|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_VL_CUSTO_REAL_TOTAL" name="ckb_COL_VL_CUSTO_REAL_TOTAL"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_CUSTO_REAL_TOTAL.click();">VL Custo Total (Real)</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_TOTAL_NF|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_VL_TOTAL_NF" name="ckb_COL_VL_TOTAL_NF"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_TOTAL_NF.click();">VL Total NF</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_TOTAL|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_VL_TOTAL" name="ckb_COL_VL_TOTAL"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_TOTAL.click();">VL Total</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_RA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
                	        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_VL_RA" name="ckb_COL_VL_RA"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_RA.click();">VL RA</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_RT|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
                	        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_RT" name="ckb_COL_RT"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_RT.click();">RT</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_ICMS_UF_DEST|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
                	        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_ICMS_UF_DEST" name="ckb_COL_ICMS_UF_DEST"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_ICMS_UF_DEST.click();">ICMS UF Destino</span><img src="../IMAGEM/exclamacao_14x14.png" id="colIcmsUfDestExclamacao" style="cursor:pointer" title="A inclusão desse campo aumenta consideravelmente o tempo de processamento" /><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_QTDE_PARCELAS|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_QTDE_PARCELAS" name="ckb_COL_QTDE_PARCELAS"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_QTDE_PARCELAS.click();">Quantidade Parcelas</span><br />
				
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_MEIO_PAGAMENTO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_MEIO_PAGAMENTO" name="ckb_COL_MEIO_PAGAMENTO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_MEIO_PAGAMENTO.click();">Meio de Pagamento</span><br />

				        <%	s_checked = ""
					        if InStr(s_campos_saida_default, "|ckb_COL_VL_PAGO_CARTAO_INTERNET|") <> 0 then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO DETPED" tabindex="-1" id="ckb_COL_VL_PAGO_CARTAO_INTERNET" name="ckb_COL_VL_PAGO_CARTAO_INTERNET"
						    value="ON" <%=s_checked%> /><span class="C SpnDETPED" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_PAGO_CARTAO_INTERNET.click();">VL Pago Cartão (Internet)</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_CHAVE_NFE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_CHAVE_NFE" name="ckb_COL_CHAVE_NFE"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_CHAVE_NFE.click();">Chave de Acesso NFe</span><br />

					</fieldset>
				</td>
			</tr>
		</table>
		<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:8px;">
		<tr>
		<td align="left">
			<input name="bMarcarTodos" id="bMarcarTodos" type="button" class="Button" onclick="marcarTodos();" value="Marcar todos" title="assinala todos os campos de saída" style="margin-left:6px;margin-bottom:10px">
		</td>
		<td align="right">
			<input name="bDesmarcarTodos" id="bDesmarcarTodos" type="button" class="Button" onclick="desmarcarTodos();" value="Desmarcar todos" title="desmarca todos os campos de saída" style="margin-left:6px;margin-right:6px;margin-bottom:10px">
		</td>
		</tr>
		</table>
	</td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
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
