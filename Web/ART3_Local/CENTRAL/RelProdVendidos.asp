<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<!-- #include file = "../global/global.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  R E L P R O D V E N D I D O S . A S P
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

    Const COD_CONSULTA_POR_PERIODO_CADASTRO = "CADASTRO"
	Const COD_CONSULTA_POR_PERIODO_ENTREGA = "ENTREGA"

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

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
	    $("#c_dt_cadastro_inicio").hUtilUI('datepicker_filtro_inicial');
	    $("#c_dt_cadastro_termino").hUtilUI('datepicker_filtro_final');
	    $("#c_dt_entregue_inicio").hUtilUI('datepicker_filtro_inicial');
	    $("#c_dt_entregue_termino").hUtilUI('datepicker_filtro_final');
	});
</script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
var s_de, s_ate;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;
var b;

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

	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";

	if (f.rb_saida[1].checked) setTimeout('exibe_botao_confirmar()', 10000);
	
	f.submit();
}

function exibe_botao_confirmar() {
	dCONFIRMA.style.visibility = "";
	window.status = "";
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


<body onload="fFILTRO.c_dt_inicio.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelProdVendidosExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Produtos Vendidos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PERÍODO  -->
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
				<span class="PLTe" style="cursor:default" onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();">CADASTRADOS ENTRE</span>
			</td>
		</tr>
		</table>
		<table cellspacing="0" cellpadding="0"><tr bgcolor="#FFFFFF"><td align="left">
		<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_cadastro_inicio" id="c_dt_cadastro_inicio"
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cadastro_termino.focus(); else {if (!fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].checked) fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();} filtra_data();"
			onchange="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();"
			onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();"
			>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;até&nbsp;</span>&nbsp;
			<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_cadastro_termino" id="c_dt_cadastro_termino"
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entregue_inicio.focus(); else {if (!fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].checked) fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();} filtra_data();" 
			onchange="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();"
			onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();">
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
			</td></tr>
		</table>
		</td></tr>

<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">FABRICANTE</span>
	<br>
		<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 30px;">
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_fabricante" name="rb_fabricante"
					value="UM"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_fabricante[0].click();">Fabricante</span>
			</td>
			<td align="left">
				<input maxlength="4" class="Cc" style="width:50px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante_de.focus(); else fFILTRO.rb_fabricante[0].click(); filtra_fabricante();">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_fabricante" name="rb_fabricante"
					value="FAIXA"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_fabricante[1].click();">Fabricantes</span>
			</td>
			<td align="left">
				<input maxlength="4" class="Cc" style="width:50px;" name="c_fabricante_de" id="c_fabricante_de" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante_ate.focus(); else fFILTRO.rb_fabricante[1].click(); filtra_fabricante();">
				<span class="C">a</span>
				<input maxlength="4" class="Cc" style="width:50px;" name="c_fabricante_ate" id="c_fabricante_ate" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto.focus(); else fFILTRO.rb_fabricante[1].click(); filtra_fabricante();">
			</td>
		</tr>
		</table>
	</td></tr>

<!--  PRODUTO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">PRODUTO</span>
	<br>
		<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 30px;">
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_produto" name="rb_produto"
					value="UM"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_produto[0].click();">Produto</span>
			</td>
			<td align="left">
				<input maxlength="13" class="Cc" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto_de.focus(); else fFILTRO.rb_produto[0].click(); filtra_produto();">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_produto" name="rb_produto"
					value="FAIXA"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_produto[1].click();">Produtos</span>
			</td>
			<td align="left">
				<input maxlength="13" class="Cc" style="width:100px;" name="c_produto_de" id="c_produto_de" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto_ate.focus(); else fFILTRO.rb_produto[1].click(); filtra_produto();">
				<span class="C">a</span>
				<input maxlength="13" class="Cc" style="width:100px;" name="c_produto_ate" id="c_produto_ate" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) fFILTRO.c_grupo.focus(); else fFILTRO.rb_produto[1].click(); filtra_produto();">
			</td>
		</tr>
		</table>
	</td></tr>

<!--  GRUPO DE PRODUTOS  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">GRUPO DE PRODUTOS</span>
	<br>
		<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 30px;">
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_grupo" name="rb_grupo"
					value="UM"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_grupo[0].click();">Grupo</span>
			</td>
			<td align="left">
				<input maxlength="2" class="Cc" style="width:60px;" name="c_grupo" id="c_grupo" onkeypress="if (digitou_enter(true)) fFILTRO.c_grupo_de.focus(); else fFILTRO.rb_grupo[0].click();" onblur="this.value=ucase(this.value);">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_grupo" name="rb_grupo"
					value="FAIXA"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_grupo[1].click();">Grupos</span>
			</td>
			<td align="left">
				<input maxlength="2" class="Cc" style="width:60px;" name="c_grupo_de" id="c_grupo_de" onkeypress="if (digitou_enter(true)) fFILTRO.c_grupo_ate.focus(); else fFILTRO.rb_grupo[1].click();" onblur="this.value=ucase(this.value);">
				<span class="C">a</span>
				<input maxlength="2" class="Cc" style="width:60px;" name="c_grupo_ate" id="c_grupo_ate" onkeypress="if (digitou_enter(true)) fFILTRO.c_loja.focus(); else fFILTRO.rb_grupo[1].click();" onblur="this.value=ucase(this.value);">
			</td>
		</tr>
		</table>
	</td></tr>

<!--  EMPRESA  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe">EMPRESA</span>
		<br>
			<select id="c_empresa" name="c_empresa" style="margin:1px 10px 6px 10px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>

<!--  LOJA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">LOJA(S)</span>
	<br>
		<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 30px;">
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<textarea class="PLBe" style="width:100px;font-size:9pt;margin-bottom:4px;" rows="8" name="c_loja" id="c_loja" onkeypress="if (!digitou_enter(false) && !digitou_char('-')) filtra_numerico();" onblur="this.value=normaliza_lista_lojas(this.value);"></textarea>
			</td>
		</tr>
		</table>
	</td></tr>

<!--  SAÍDA DO RELATÓRIO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">SAÍDA DO RELATÓRIO</span>
	<br><input type="radio" tabindex="-1" id="rb_saida" name="rb_saida" value="Html" onclick="dCONFIRMA.style.visibility='';" checked><span class="C" style="cursor:default" onclick="fFILTRO.rb_saida[0].click(); dCONFIRMA.style.visibility='';"
		>Html</span>

	<br><input type="radio" tabindex="-1" id="rb_saida" name="rb_saida" value="XLS" onclick="dCONFIRMA.style.visibility='';"><span class="C" style="cursor:default" onclick="fFILTRO.rb_saida[1].click(); dCONFIRMA.style.visibility='';"
		>Excel</span>
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
