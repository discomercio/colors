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
'	  R E L G E R E N C I A L V E N D A S F I L T R O . A S P
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

	Const COD_CONSULTA_POR_PERIODO_CADASTRO = "CADASTRO"
	Const COD_CONSULTA_POR_PERIODO_ENTREGA = "ENTREGA"
	Const COD_SAIDA_REL_FABRICANTE = "FABRICANTE"
	Const COD_SAIDA_REL_PRODUTO = "PRODUTO"
	Const COD_SAIDA_REL_VENDEDOR = "VENDEDOR"
	Const COD_SAIDA_REL_INDICADOR = "INDICADOR"
	Const COD_SAIDA_REL_UF = "UF"
	Const COD_SAIDA_REL_INDICADOR_UF = "INDICADOR_UF"
	Const COD_SAIDA_REL_CIDADE_UF = "CIDADE_UF"
    Const COD_SAIDA_REL_ORIGEM_PEDIDO = "ORIGEM_PEDIDO"
    Const COD_SAIDA_REL_EMPRESA = "EMPRESA"
    Const COD_SAIDA_REL_GRUPO_PRODUTO = "GRUPO_PRODUTO"

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

	dim intIdx
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

	dim strScript
	strScript = _
		"<script language='JavaScript'>" & chr(13) & _
		"var COD_SAIDA_REL_PRODUTO = '" & COD_SAIDA_REL_PRODUTO & "';" & chr(13) & _
		"var COD_SAIDA_REL_VENDEDOR = '" & COD_SAIDA_REL_VENDEDOR & "';" & chr(13) & _
		"var COD_SAIDA_REL_INDICADOR = '" & COD_SAIDA_REL_INDICADOR & "';" & chr(13) & _
		"var COD_SAIDA_REL_UF = '" & COD_SAIDA_REL_UF & "';" & chr(13) & _
		"var COD_SAIDA_REL_INDICADOR_UF = '" & COD_SAIDA_REL_INDICADOR_UF & "';" & chr(13) & _
		"var COD_SAIDA_REL_CIDADE_UF = '" & COD_SAIDA_REL_CIDADE_UF & "';" & chr(13) & _
        "var COD_SAIDA_REL_ORIGEM_PEDIDO = '" & COD_SAIDA_REL_ORIGEM_PEDIDO & "';" & chr(13) & _
        "var COD_SAIDA_REL_GRUPO_PRODUTO = '" & COD_SAIDA_REL_GRUPO_PRODUTO & "';" & chr(13) & _
        "var COD_CONSULTA_POR_PERIODO_ENTREGA = '" & COD_CONSULTA_POR_PERIODO_ENTREGA & "';" & chr(13) & _
		"</script>" & chr(13)




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________________________________________________
' INDICADORES MONTA ITENS SELECT
' LEMBRE-SE: O ORÇAMENTISTA É CONSIDERADO AUTOMATICAMENTE UM INDICADOR!!
function indicadores_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql="SELECT" & _
				" apelido," & _
				" razao_social_nome_iniciais_em_maiusculas" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR" & _
			" WHERE " & _
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
			" ORDER BY apelido"
	set r = cn.Execute(strSql)
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

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	indicadores_monta_itens_select = strResp
	r.close
	set r=nothing
end function


' ____________________________________________________________________________
' CAPTADORES MONTA ITENS SELECT
function captadores_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT" & _
				" usuario," & _
				" nome_iniciais_em_maiusculas" & _
			" FROM t_USUARIO" & _
			" WHERE" & _
				" usuario IN " & _
					"(" & _
						"SELECT DISTINCT" & _
							" captador" & _
						" FROM t_ORCAMENTISTA_E_INDICADOR" & _
						" WHERE" & _
							" (captador IS NOT NULL)" & _
							" AND (" & _
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
							")" & _
					")"
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
		
	captadores_monta_itens_select = strResp
	r.close
	set r=nothing
end function
'----------------------------------------------------------------------------------------------
' GRUPOS MONTA ITENS SELECT
function grupos_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT grupo from t_PRODUTO WHERE grupo <> '' ORDER by grupo"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("grupo")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x 
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	grupos_monta_itens_select = strResp
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
	
    strResp = "<option value=''>&nbsp;</option>" & strResp

	grupo_origem_pedido_monta_itens_select = strResp
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
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<%=strScript%>

<script type="text/javascript">
	$(function() {
		var vOptionFields;
		var vOptionEscolher;
		var vOptionEscolhida;

		$("input[type=radio]").hUtil('fix_radios');
		$("#c_dt_cadastro_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_cadastro_termino").hUtilUI('datepicker_filtro_final');
		$("#c_dt_entregue_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_entregue_termino").hUtilUI('datepicker_filtro_final');
		$("#msgAguarde").hide();
		$("#tr_cnpj_cpf").hide();
		$("#tr_cidade_comeca_com_sep").hide();
		$("#tr_cidade_comeca_com").hide();

		$("#btnAdiciona").click(function() {
			var x = $("#c_loc_a_escolher option:selected");
			$("#c_loc_escolhidas").append(x);
			$("#c_loc_digitada").val("");
			reOrdenarEscolhidos();
		});

		$("#btnRemove").click(function() {
			var x = $("#c_loc_escolhidas option:selected");
			$("#c_loc_a_escolher").append(x);
			reOrdenarAEscolher();
		});

		$("#c_loc_a_escolher").dblclick(function() {
			var x = $("#c_loc_a_escolher option:selected");
			$("#c_loc_escolhidas").append(x);
			$("#c_loc_digitada").val("");
			reOrdenarEscolhidos();
		});

		$("#c_loc_escolhidas").dblclick(function() {
			var x = $("#c_loc_escolhidas option:selected");
			$("#c_loc_a_escolher").append(x);
			reOrdenarAEscolher();
		});

		$("#c_loc_digitada").keyup(function() {
			if ($("#c_loc_digitada").val() != "") {
				$("#c_loc_escolhidas option").each(function() {
					$("#c_loc_a_escolher").append(this);
				});
				$("#c_loc_escolhidas").empty();
				reOrdenarAEscolher();
			}
		});

		if ($("#c_loc_uf").val().length > 0) {
			$("#c_loc_escolhidas").empty();
			$("#c_loc_a_escolher").empty();
			if ($("#c_hidden_loc_a_escolher").val().length > 0) {
				vOptionEscolher = fFILTRO.c_hidden_loc_a_escolher.value.split("|");
				for (i = 0; i < vOptionEscolher.length; i++) {
					vOptionFields = vOptionEscolher[i].split("¥");
					oOption = document.createElement("OPTION");
					oOption.value = vOptionFields[0];
					oOption.innerText = vOptionFields[1];
					$("#c_loc_a_escolher").append(oOption);
				}
				reOrdenarAEscolher();
			}
			if ($("#c_hidden_loc_escolhidas").val().length > 0) {
				vOptionEscolhida = fFILTRO.c_hidden_loc_escolhidas.value.split("|");
				for (i = 0; i < vOptionEscolhida.length; i++) {
					vOptionFields = vOptionEscolhida[i].split("¥");
					oOption = document.createElement("OPTION");
					oOption.value = vOptionFields[0];
					oOption.innerText = vOptionFields[1];
					$("#c_loc_escolhidas").append(oOption);
				}
				reOrdenarEscolhidos();
			}
		}
	});

	//início bloco cidades

	function obtemLocalidades() {
		$.ajax({
			type: "GET",
			url: "../GLOBAL/AjaxClienteUFLocPesqBD.asp",
			cache: false,
			async: false,
			success: function(response) {
				if (response != "") {
					fFILTRO.lista_uf_cidade.value = response;
				}
			},
			error: function(response) {
				alert("Erro ao pesquisar lista de locais!");
			}
		});
	}

	function carregaLocalidadesLenta(estado) {
		//divide a lista de estados e cidades para utilização; a lista possui os seguintes caracteres de separação:
		// £ - separa UFs
		// ¥ - separa cidades
		//
		//exemplo: se a lista contivesse o estado de SP, com a cidade de São Paulo e a cidade de Jales, a lista seria:
		//SP¥São Paulo¥Jales
		//
		//exemplo 2: se a lista anterior tivesse ainda o estado de MG, com a cidade de Belo Horizonte, seria:
		//SP¥São Paulo¥Jales£MG¥Belo Horizonte
		//
		//

		var f, i, j, k;
		var suf, scidade;
		var strUF, strCidade;
		var boAchou;

		f = fFILTRO;

		$("#c_loc_a_escolher").empty();
		$("#c_loc_escolhidas").empty();
		$("#msgAguarde").show();

		if ((estado != "") && (estado != undefined)) {
			suf = f.lista_uf_cidade.value.split('£');
			for (i = 0; (i < suf.length) && (strUF != estado); i++) {

				if (suf[i].substring(0, 2) == estado) {

					scidade = suf[i].split('¥');
					//a primeira posição desta lista será sempre a UF
					strUF = scidade[0];
					if (strUF == estado) {
						for (j = 1; (j < scidade.length); j++) {
							strCidade = scidade[j];
							//adicionar as cidades à caixa de seleção
							//verificar se a localidade já se encontra na caixa de seleção; senão, acrescentar
							boAchou = false;
							$("#c_loc_a_escolher").find("option").each(function(l) {
								if ($(this).val() == strCidade) {
									boAchou = true;
								}
							});
							if (!boAchou) {
								$("#c_loc_a_escolher").append("<option>" + strCidade + "</option>");
							}
						}
					}
				}
			}
		}

		$("#msgAguarde").hide();
	}

	function reOrdenarAEscolher() {
		$("#c_loc_a_escolher").html($("#c_loc_a_escolher option").sort(function(a, b) {
			return a.text.toUpperCase() == b.text.toUpperCase() ? 0 : a.text.toUpperCase() < b.text.toUpperCase() ? -1 : 1
		}))
	}

	function reOrdenarEscolhidos() {
		$("#c_loc_escolhidas").html($("#c_loc_escolhidas option").sort(function(a, b) {
			return a.text.toUpperCase() == b.text.toUpperCase() ? 0 : a.text.toUpperCase() < b.text.toUpperCase() ? -1 : 1
		}))
	}

	//fim bloco cidades
</script>

<script language="JavaScript" type="text/javascript">
	function LimpaListaLocalidades() {
		var f, i, oOption;
		f = fFILTRO;
		for (i = f.c_loc_a_escolher.length - 1; i >= 0; i--) {
			f.c_loc_a_escolher.remove(i);
		}

		for (i = f.c_loc_escolhidas.length - 1; i >= 0; i--) {
			f.c_loc_escolhidas.remove(i);
		}
	}

	function TrataRespostaAjaxPesquisaLocalidades() {
		var f, i, strAux, strResp, xmlDoc, oOption, oNodes;
		f = fFILTRO;
		if (objAjaxPesqLocalidades.readyState == AJAX_REQUEST_IS_COMPLETE) {
			strResp = objAjaxPesqLocalidades.responseText;
			if (strResp == "") {
				window.status = "Concluído";
				$("#msgAguarde").hide();
				alert("Nenhuma localidade encontrada!!");
				return;
			}

			if (strResp != "") {
				try {
					xmlDoc = objAjaxPesqLocalidades.responseXML.documentElement;
					for (i = 0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
						oOption = document.createElement("OPTION");
						f.c_loc_a_escolher.options.add(oOption);

						oNodes = xmlDoc.getElementsByTagName("localidade")[i];
						if (oNodes.childNodes.length > 0) strAux = oNodes.childNodes[0].nodeValue; else strAux = "";
						if (strAux == null) strAux = "";
						oOption.innerText = strAux;
						oOption.value = strAux;
					}
				}
				catch (e) {
					alert("Falha na consulta!!");
				}
			}
			window.status = "Concluído";
			$("#msgAguarde").hide();
			f.c_loc_a_escolher.focus();
		}
	}

	function CarregaLocalidades() {
		var f, strUrl, strUF;
		f = fFILTRO;
		objAjaxPesqLocalidades = GetXmlHttpObject();
		if (objAjaxPesqLocalidades == null) {
			alert("O browser NÃO possui suporte ao AJAX!!");
			return;
		}

		//  Limpa lista de localidades
		LimpaListaLocalidades();

		strUF = trim(f.c_loc_uf.value);
		if (strUF == "") {
			return;
		}

		window.status = "Aguarde, pesquisando as localidades de " + f.c_loc_uf.value + " ...";
		$("#msgAguarde").show();

		strUrl = "../Global/AjaxCepLocalidadesPesqBD.asp";
		strUrl = strUrl + "?uf=" + f.c_loc_uf.value + "&retira_acentuacao=S";
		//  Prevents server from using a cached file
		strUrl = strUrl + "&sid=" + Math.random() + Math.random();
		objAjaxPesqLocalidades.onreadystatechange = TrataRespostaAjaxPesquisaLocalidades;
		objAjaxPesqLocalidades.open("GET", strUrl, true);
		objAjaxPesqLocalidades.send(null);
	}
</script>

<script language="JavaScript" type="text/javascript">
	$(document).ready(function() {
		if (($("input[name='rb_saida']:checked").val() == COD_SAIDA_REL_INDICADOR) || ($("input[name='rb_saida']:checked").val() == COD_SAIDA_REL_INDICADOR_UF)) {
			$("#tr_FormaComoConheceu").show();
		}
		else {
			$("#tr_FormaComoConheceu").hide();
		}

		if ($("input[name='rb_saida']:checked").val() == COD_SAIDA_REL_CIDADE_UF) {
		    
		    EscondeUF();
		}
		if ($("input[name='rb_saida']:checked").val() == COD_SAIDA_REL_CIDADE_UF) {
			$("#tr_Cidade").show();
		}
		else {
			$("#tr_Cidade").hide();
		}

		if ($("input[name='rb_saida']:checked").val() == COD_SAIDA_REL_GRUPO_PRODUTO) {
		    $("#ckb_ordenar_marg_contrib").prop("disabled", false);
		}
		else {
		    $("#ckb_ordenar_marg_contrib").prop("checked", false);
		    $("#ckb_ordenar_marg_contrib").prop("disabled", true);
		}
		
		$("input[name='rb_saida']").change(function() {
			if (($("input[name='rb_saida']:checked").val() == COD_SAIDA_REL_INDICADOR) || ($("input[name='rb_saida']:checked").val() == COD_SAIDA_REL_INDICADOR_UF)) {
				$("#tr_FormaComoConheceu").show();
			}
			else {
				$("#tr_FormaComoConheceu").hide();
			}
			if ($("input[name='rb_saida']:checked").val() == COD_SAIDA_REL_CIDADE_UF) {
			    $("#tr_Cidade").show();
			    EscondeUF();
			}
			else {
			    $("#tr_Cidade").hide();
			    MostraUf();
			}
			if ($("input[name='rb_saida']:checked").val() == COD_SAIDA_REL_GRUPO_PRODUTO) {
			    $("#ckb_ordenar_marg_contrib").prop("disabled", false);
			}
			else {
			    $("#ckb_ordenar_marg_contrib").prop("disabled", true);
			    $("#ckb_ordenar_marg_contrib").prop("checked", false);
            }
		});
		$("input[name='rb_periodo']").change(function() {
		    if ($("input[name='rb_saida']:checked").val() == COD_SAIDA_REL_GRUPO_PRODUTO) {
		        $("#ckb_ordenar_marg_contrib").prop("disabled", false);
		    }
		    else {
		        $("#ckb_ordenar_marg_contrib").prop("disabled", true);
		        $("#ckb_ordenar_marg_contrib").prop("checked", false);
		    }
		});
	});
</script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
var s_de, s_ate;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;
var i, blnFlagOk;

//  TIPO DE CONSULTA: PEDIDOS CADASTRADOS OU PEDIDOS ENTREGUES
	blnFlagOk=false;
	for (i=0; i<f.rb_periodo.length; i++) {
		if (f.rb_periodo[i].checked) blnFlagOk=true;
		}
	if (!blnFlagOk) {
		alert("Selecione o tipo de consulta:\n    Por pedidos cadastrados\n    Por pedidos entregues");
		return;
		}
		
//  COLUNA DE SAÍDA DO RELATÓRIO
	blnFlagOk=false;
	for (i=0; i<f.rb_saida.length; i++) {
		if (f.rb_saida[i].checked) blnFlagOk=true;
		}
	if (!blnFlagOk) {
		alert("Selecione a coluna de saída do relatório!!");
		return;
		}
	
//  PERÍODO DE CADASTRO
	if (f.rb_periodo[0].checked) {
		if (trim(f.c_dt_cadastro_inicio.value)!="") {
			if (!isDate(f.c_dt_cadastro_inicio)) {
				alert("Data inválida!!");
				f.c_dt_cadastro_inicio.focus();
				return;
				}
			}

		if (trim(f.c_dt_cadastro_termino.value)!="") {
			if (!isDate(f.c_dt_cadastro_termino)) {
				alert("Data inválida!!");
				f.c_dt_cadastro_termino.focus();
				return;
				}
			}
		
		s_de = trim(f.c_dt_cadastro_inicio.value);
		s_ate = trim(f.c_dt_cadastro_termino.value);
		if ((s_de!="")&&(s_ate!="")) {
			s_de=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
			s_ate=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
			if (s_de > s_ate) {
				alert("Data de término é menor que a data de início!!");
				f.c_dt_cadastro_termino.focus();
				return;
				}
			}
		}

//  PERÍODO DE ENTREGA
	if (f.rb_periodo[1].checked) {
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
		}

	if ((trim(f.c_produto.value)!="")&&(trim(f.c_fabricante.value)=="")) {
		if (!isEAN(f.c_produto.value)) {
			alert("Preencha o código do fabricante do produto " + f.c_produto.value + "!!");
			f.c_fabricante.focus();
			return;
			}
		}

//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
	//  PERÍODO DE CADASTRO
		if (f.rb_periodo[0].checked) {
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
	
	// PERÍODO DE ENTREGA
		if (f.rb_periodo[1].checked) {
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
		}

	dCONFIRMA.style.visibility = "hidden";

//	SELECIONA TODOS OS OPTIONS DA LISTA DE CIDADES ESCOLHIDAS, POIS SOMENTE OS ITENS SELECIONADOS SÃO RECUPERADOS AO OBTER OS DADOS DO FORMULÁRIO
	$("#c_loc_escolhidas option").prop("selected", "selected");

	f.c_hidden_loc_a_escolher.value = "";
	$("#c_loc_a_escolher option").each(function() {
		if ($("#c_hidden_loc_a_escolher").val().length > 0) f.c_hidden_loc_a_escolher.value += "|";
		f.c_hidden_loc_a_escolher.value += $(this).val() + "¥" + $(this).text();
	});

	f.c_hidden_loc_escolhidas.value = "";
	$("#c_loc_escolhidas option").each(function() {
		if ($("#c_hidden_loc_escolhidas").val().length > 0) f.c_hidden_loc_escolhidas.value += "|";
		f.c_hidden_loc_escolhidas.value += $(this).val() + "¥" + $(this).text();
	});

	
	window.status = "Aguarde ...";
	f.submit();
}
</script>
<script language="JavaScript" type="text/javascript">
    function limpaCampoSelect(c) {
        c.options[0].selected = true;
    }
    function limpaCampoSelectGrupo() {
        $("#c_grupo").children().prop('selected', false);
    }
    function deselecionar_checkbox(){      
        
        $("#ckb_colocado_mesmo_periodo").prop("checked", false);
    } 

</script>
<script language="JavaScript" type="text/javascript">

    function MostraUf(){
        $("#UF").show();
    }
    function EscondeUF(){       
        $("#UF").hide();             
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
html
{
	overflow-y: scroll;
}
</style>


<body onload="fFILTRO.c_dt_cadastro_inicio.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelGerencialVendasExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_loja" id="c_loja" value="<%=loja%>">
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" name="lista_uf_cidade" id="lista_uf_cidade" value="" />
<input type="hidden" name="c_hidden_loc_a_escolher" id="c_hidden_loc_a_escolher" value="" />
<input type="hidden" name="c_hidden_loc_escolhidas" id="c_hidden_loc_escolhidas" value="" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório Gerencial de Vendas</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellspacing="0">
<!--  CADASTRADOS ENTRE  -->
	<tr bgcolor="#FFFFFF">
	<td class="MT" align="left" nowrap>
		<% intIdx=-1 %>
		<table cellspacing="0" cellpadding="0">
		<tr>
			<td align="left">
				<input type="radio" id="rb_periodo" name="rb_periodo" value="<%=COD_CONSULTA_POR_PERIODO_CADASTRO%>">
			</td>
			<td align="left" valign="bottom">
				<% intIdx=intIdx+1 %>
				<span class="PLTe" style="cursor:default" onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();deselecionar_checkbox();">CADASTRADOS ENTRE</span>
			</td>
		</tr>
		</table>
		<table cellspacing="0" cellpadding="0"><tr bgcolor="#FFFFFF"><td align="left">
		<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_cadastro_inicio" id="c_dt_cadastro_inicio" 
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cadastro_termino.focus(); else {if (!fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].checked) fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();deselecionar_checkbox();} filtra_data();"
			onchange="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();deselecionar_checkbox();"
			onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();deselecionar_checkbox();"
			>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;até&nbsp;</span>&nbsp;
			<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_cadastro_termino" id="c_dt_cadastro_termino" 
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entregue_inicio.focus(); else {if (!fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].checked) fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();deselecionar_checkbox();} filtra_data();" 
			onchange="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();deselecionar_checkbox();"
			onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();deselecionar_checkbox();">
			</td></tr>
		</table>
		</td></tr>

<!--  ENTREGUE ENTRE  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		<table cellspacing="0" cellpadding="0">
		<tr>
			<td align="left">
				<input type="radio" id="rb_periodo" name="rb_periodo" value="<%=COD_CONSULTA_POR_PERIODO_ENTREGA%>">
			</td>
			<td align="left" valign="bottom">
				<% intIdx=intIdx+1 %>
				<span class="PLTe" style="cursor:default" onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();">ENTREGUES ENTRE</span>
			</td>
		</tr>
		</table>
		<table cellspacing="0" cellpadding="0"><tr bgcolor="#FFFFFF"><td align="left">
		<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" 
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entregue_termino.focus(); else {if (!fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].checked) fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();} filtra_data();"
			onchange="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();"
			onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();"
			>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;até&nbsp;</span>&nbsp;
			<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_entregue_termino" id="c_dt_entregue_termino" 
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); else {if (!fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].checked) fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();} filtra_data();" 
			onchange="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();"
			onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();">
            <input type="checkbox" name="ckb_colocado_mesmo_periodo" id="ckb_colocado_mesmo_periodo" value="1" onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();"/> <span class="PLLc" style="cursor:default" onclick="fFILTRO.ckb_colocado_mesmo_periodo.click();fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();">&nbsp;Colocado no mesmo período&nbsp;</span>            
			</td></tr>
		</table>
		</td></tr>

<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">FABRICANTE</span>
	<br>
		<select id="c_fabricante" name="c_fabricante" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
		<%=fabricante_monta_itens_select(Null) %>
		</select>
		</td></tr>

<!--  PRODUTO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">PRODUTO</span>
	<br>
		<input maxlength="13" class="PLLe" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) fFILTRO.c_vendedor.focus(); filtra_produto();">
		</td></tr>

<!--  VENDEDOR  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">VENDEDOR</span>
	<br>
        <select id="c_vendedor" name="c_vendedor" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =vendedores_monta_itens_select(Null) %>
			</select>
		</td></tr>

<!--  INDICADOR  -->
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">INDICADOR</span>
		<br>
			<select id="c_indicador" name="c_indicador" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =indicadores_monta_itens_select(Null) %>
			</select>
			</td></tr>

<!--  CAPTADOR  -->
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">CAPTADOR</span>
		<br>
			<select id="c_captador" name="c_captador" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =captadores_monta_itens_select(Null) %>
			</select>
			</td></tr>

<!-- EMPRESA -->

    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">EMPRESA</span>
		<br>
			<select id="c_empresa" name="c_empresa" style="margin:1px 3px 6px 10px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
			</select>
			
        </td>
    </tr>

<!--  GRUPO DE ORIGEM DO PEDIDO  -->
    <%if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then%>
    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">ORIGEM DO PEDIDO (GRUPO)</span>
		<br>
			<select id="c_grupo_pedido_origem" name="c_grupo_pedido_origem" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =grupo_origem_pedido_monta_itens_select(Null) %>
			</select>
			</td></tr>
    <%end if%>

<!--  ORIGEM DO PEDIDO  -->
    <%if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then%>
    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">ORIGEM DO PEDIDO</span>
		<br>
			<select id="c_pedido_origem" name="c_pedido_origem" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =origem_pedido_monta_itens_select(Null) %>
			</select>
			</td></tr>
    <%end if%>

<!--  GRUPO  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">GRUPO</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_grupo" name="c_grupo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="10"style="width:100px;margin:1px 10px 6px 10px;" multiple>
			<% =grupos_monta_itens_select(get_default_valor_texto_bd(usuario, "RelGerencialVendasFiltro|c_grupo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparGrupo" id="bLimparGrupo" href="javascript:limpaCampoSelectGrupo()" title="limpa o filtro 'Fabricante'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
    <!--  CNPJ/CPF  -->
	<tr bgcolor="#FFFFFF" id="tr_cnpj_cpf">
	<td class="MDBE" align="left" nowrap><span class="PLTe">CNPJ/CPF</span>
	<br>
		<input maxlength="18" class="PLLe" style="width:150px;" name="c_cnpj_cpf" id="c_cnpj_cpf" onkeypress="if (digitou_enter(true) && tem_info(this.value) && cnpj_cpf_ok(this.value)) {bCONFIRMA.focus(); this.value=cnpj_cpf_formata(this.value);} filtra_cnpj_cpf();" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);">
	</td></tr>

<!--  TIPO DE CLIENTE  -->
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">TIPO DE CLIENTE</span>
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
	</td></tr>

    <!--  FORMA DE PAGAMENTO  -->
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
<!--  UF  -->
	<tr id='UF' bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">UF</span>
		<br>
			<select id="c_uf_saida" name="c_uf_saida" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =UF_monta_itens_select(Null) %>
			</select>
	</td></tr>
<!--  SAÍDA DO RELATÓRIO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		<table width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="left" valign="top">
				<span class="PLTe">SAÍDA DO RELATÓRIO</span>
				<br>
					<% intIdx=-1 %>
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_FABRICANTE%>"  class="CBOX" style="margin-left:20px;">
					<% intIdx=intIdx+1 %>
					<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">Fabricante</span>
					<br>
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_PRODUTO%>" " class="CBOX" style="margin-left:20px;">
					<% intIdx=intIdx+1 %>
					<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">Produto</span>
					<br>
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_VENDEDOR%>" " class="CBOX" style="margin-left:20px;">
					<% intIdx=intIdx+1 %>
					<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">Vendedor</span>
					<br>
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_INDICADOR%>"  class="CBOX" style="margin-left:20px;">
					<% intIdx=intIdx+1 %>
					<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">Indicador</span>
					<br>
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_UF%>"  class="CBOX" style="margin-left:20px;">
					<% intIdx=intIdx+1 %>
					<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">UF</span>
					<br>
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_INDICADOR_UF%>"  class="CBOX" style="margin-left:20px;">
					<% intIdx=intIdx+1 %>
					<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">Indicador/UF</span>
					<br>
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_CIDADE_UF%>"  class="CBOX" style="margin-left:20px;">
					<% intIdx=intIdx+1 %>
					<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">Cidade/UF</span>
                    <%if loja = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then%>
                    <br />
                    <input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_ORIGEM_PEDIDO%>"  class="CBOX" style="margin-left:20px;">
					<% intIdx=intIdx+1 %>
					<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">Origem do Pedido</span>                                        
                    <%end if %>
                    <br />
                    <input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_EMPRESA%>"  class="CBOX" style="margin-left:20px;">
					<% intIdx=intIdx+1 %>
					<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">Empresa</span>
                    <br />
                    <input type="radio" id="rb_saida_produto_grupo" name="rb_saida" value="<%=COD_SAIDA_REL_GRUPO_PRODUTO%>" checked class="CBOX" style="margin-left:20px;">
					<% intIdx=intIdx+1 %>
					<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();habilitar_ckb_ordenar_marg_contrib();">Grupo de Produtos</span>
                    &nbsp;&nbsp;&nbsp;
                    <input type="checkbox" name="ckb_ordenar_marg_contrib" id="ckb_ordenar_marg_contrib" value="1" />
                        <span class="PLLc" style="cursor:default" onclick="fFILTRO.ckb_ordenar_marg_contrib.click();fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();">&nbsp;Ordenar pela Margem Contrib</span>            
			</td>
		</tr>
		</table>
	</td></tr>

<!--  INDICADOR: FORMA COMO CONHECEU  -->
	<tr bgcolor="#FFFFFF" id="tr_FormaComoConheceu">
	<td class="MDBE" align="left" nowrap><span class="PLTe">FORMA COMO CONHECEU A BONSHOP</span>
		<br>
		<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
		<tr bgcolor="#FFFFFF"><td align="left">
			<p class="C">
			<select id="c_forma_como_conheceu_codigo" name="c_forma_como_conheceu_codigo" style="margin-top:4pt; margin-bottom:4pt;width:490px;">
				<%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__CAD_ORCAMENTISTA_E_INDICADOR__FORMA_COMO_CONHECEU, "")%>
			</select>
			</p>
		</td></tr>
		</table>
	</td></tr>

<!--  CIDADE  -->
	<tr bgcolor="#FFFFFF" id="tr_Cidade">
		<td class="MDBE" align="left" nowrap><span class="PLTe">CIDADE</span>
		<br>
				<table>
					<tr>
						<td valign="middle" align="left">
							<span class="Cd" style="margin:1px 10px 6px 10px;">UF</span>
							<br><select id="c_loc_uf" name="c_loc_uf" style="margin:1px 10px 6px 10px;" 
									onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;CarregaLocalidades();}" 
									onchange="CarregaLocalidades();" 
									onkeypress="if (digitou_enter(true)) fFILTRO.c_loc_a_escolher.focus();">
									<% =UF_monta_itens_select(Null) %>
								</select>
							</td>
						<td style="width:10px;" align="left">&nbsp;</td>
						<td id="msgAguarde">
							<table cellpadding="0" cellspacing="0">
								<tr>
								<td valign="middle" align="center"><span style="color:orangered;font-weight:bold;font-style:italic;font-size:10pt;">Aguarde, pesquisando cidades...</span></td>
								<td style="width:10px;" align="left">&nbsp;</td>
								<td align="left"><img src="../imagem/aguarde.gif"border="0"></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td align="left">
							<br><span id="txtAEscolher" name="txtAEscolher" class="C" style="margin:1px 10px 6px 10px;">Selecionar a Cidade</span>
							<br><select id="c_loc_a_escolher" name="c_loc_a_escolher" size="10" style="width: 200px;margin:1px 10px 6px 10px;" multiple>
							</select>
						</td>
						<td align="center" style="width:60px">
							<input type="button" id="btnAdiciona" name="btnAdiciona" style="width:40px; margin-bottom:2px;" value="&gt;" />
							<br />
							<input type="button" id="btnRemove" name="btnRemove" style="width:40px; margin-bottom:2px;" value="&lt;" />
						</td>
						<td align="left">
							<br><span id="txtEscolhidos" name="txtEscolhidos" class="C" style="margin:1px 10px 6px 10px;">Cidades Selecionadas</span>
							<br><select id="c_loc_escolhidas" name="c_loc_escolhidas" size="10" style="width: 200px;margin:1px 10px 6px 10px;" multiple>
							</select>
						</td>
					</tr>
					<tr id="tr_cidade_comeca_com_sep">
						<td align="left">
							<br><span class="C" style="margin:1px 10px 6px 10px;">OU</span>
						</td>
					</tr>
					<tr id="tr_cidade_comeca_com">
						<td align="left">
							<br><span class="C" style="margin:1px 10px 6px 10px;">Cidade Começa Com</span>
							<br><input maxlength="40" class="C" style="width:200px; margin:1px 10px 6px 10px;" name="c_loc_digitada" id="c_loc_digitada" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus();">
						</td>
					</tr>
				</table>
			</td></tr>

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
