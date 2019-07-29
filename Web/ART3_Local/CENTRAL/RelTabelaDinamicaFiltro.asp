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
	s_campos_saida_default = get_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|campos_saida_selecionados")





' _____________________________________________________________________________________________
'
'									F  U  N  �  �  E  S 
' _____________________________________________________________________________________________

' ____________________________________________________________________________
' FABRICANTE MONTA ITENS SELECT
'
function fabricante_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
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

	strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		
	fabricante_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________
' GRUPO MONTA ITENS SELECT
'
function grupo_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
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

	strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		
	grupo_monta_itens_select = strResp
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
		$("#c_dt_faturamento_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_faturamento_termino").hUtilUI('datepicker_filtro_final');
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
    $(":checkbox").each(function() {
        if (!$(this).is(":checked")) {
            $(this).trigger('click');
        }
    });
}

function desmarcarTodos() {
    $(":checkbox").each(function() {
        if ($(this).is(":checked")) {
            $(this).trigger('click');
        }
    });
}

function fFILTROConfirma( f ) {
var s_de, s_ate, i_qtde_campos;

//  PER�ODO DE FATURAMENTO
	if (trim(f.c_dt_faturamento_inicio.value) == "") {
		alert("Informe a data de in�cio do per�odo!!");
		f.c_dt_faturamento_inicio.focus();
		return;
	}

	if (trim(f.c_dt_faturamento_termino.value) == "") {
		alert("Informe a data de t�rmino do per�odo!!");
		f.c_dt_faturamento_termino.focus();
		return;
	}

	if (trim(f.c_dt_faturamento_inicio.value) != "") {
		if (!isDate(f.c_dt_faturamento_inicio)) {
			alert("Data inv�lida!!");
			f.c_dt_faturamento_inicio.focus();
			return;
		}
	}

	if (trim(f.c_dt_faturamento_termino.value) != "") {
		if (!isDate(f.c_dt_faturamento_termino)) {
			alert("Data inv�lida!!");
			f.c_dt_faturamento_termino.focus();
			return;
		}
	}

	s_de = trim(f.c_dt_faturamento_inicio.value);
	s_ate = trim(f.c_dt_faturamento_termino.value);
	if ((s_de != "") && (s_ate != "")) {
		s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
		s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
		if (s_de > s_ate) {
			alert("Data de t�rmino � menor que a data de in�cio!!");
			f.c_dt_faturamento_termino.focus();
			return;
		}
	}

	//	CAMPOS DE SA�DA
	i_qtde_campos = 0;
	$(".CKB_CADASTRO, .CKB_COMERCIAL, .CKB_FINANCEIRO").each(function() {
		if ($(this).is(":checked")) {
			i_qtde_campos++;
		}
	});

	if (i_qtde_campos == 0) {
		alert("Nenhum campo de sa�da foi assinalado!!");
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


<!--  I D E N T I F I C A � � O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Dados para Tabela Din�mica</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellspacing="0" style="width: 400px">
<!--  PER�ODO (FATURAMENTO)  -->
	<tr bgcolor="#FFFFFF">
	<td class="MT" align="left" nowrap>
		<table cellspacing="0" cellpadding="0">
		<tr>
			<td align="left" valign="bottom">
				<span class="PLTe" style="cursor:default">PER�ODO</span>
			</td>
		</tr>
		</table>
		<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF">
			<td align="left">
				<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_faturamento_inicio" id="c_dt_faturamento_inicio"
					onblur="if (!isDate(this)) {alert('Data inv�lida!'); this.focus();}" 
					onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_faturamento_termino.focus(); filtra_data();"
					value="<%=get_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_dt_faturamento_inicio")%>"
					>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;at�&nbsp;</span>&nbsp;
					<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_faturamento_termino" id="c_dt_faturamento_termino"
					onblur="if (!isDate(this)) {alert('Data inv�lida!'); this.focus();}" 
					onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); filtra_data();"
					value="<%=get_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_dt_faturamento_termino")%>"
					/>
			</td>
			</tr>
		</table>
	</td>
	</tr>

<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		<span class="PLTe">FABRICANTE</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_fabricante" name="c_fabricante" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =fabricante_monta_itens_select(get_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_fabricante")) %>
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
			<% =grupo_monta_itens_select(get_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_grupo")) %>
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
			<% =potencia_BTU_monta_itens_select(get_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_potencia_BTU")) %>
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
			<% =ciclo_monta_itens_select(get_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_ciclo")) %>
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
	<!-- POSI��O MERCADO -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">POSI��O MERCADO</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_posicao_mercado" name="c_posicao_mercado" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =posicao_mercado_monta_itens_select(get_default_valor_texto_bd(usuario, "RelTabelaDinamicaFiltro|c_posicao_mercado")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparPosicaoMercado" id="bLimparPosicaoMercado" href="javascript:limpaCampoSelect(fFILTRO.c_posicao_mercado)" title="limpa o filtro 'Posi��o Mercado'">
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
		<span style="cursor:default" class="Np" onclick="fFILTRO.rb_tipo_cliente[<%=Cstr(intIdx)%>].click();">Pessoa F�sica</span>
		<br />
		<input type="radio" id="rb_tipo_cliente" name="rb_tipo_cliente" value=<%=ID_PJ%> style="margin-left:30px;">
		<% intIdx=intIdx+1 %>
		<span style="cursor:default" class="Np" onclick="fFILTRO.rb_tipo_cliente[<%=Cstr(intIdx)%>].click();">Pessoa Jur�dica</span>
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

<!--  AGRUPAMENTO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		
				<span class="PLTe">AGRUPAMENTO</span>
				<br />
				<input type="checkbox" tabindex="-1" id="ckb_AGRUPAMENTO" name="ckb_AGRUPAMENTO"
						value="ON" style="margin-left:30px;margin-bottom: 5px;margin-top: 5px;" /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_AGRUPAMENTO.click();">Desagrupar itens por quantidade</span><br />
	</td>
	</tr>

<!--  CAMPOS DE SA�DA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap>
		<span class="PLTe">CAMPOS DE SA�DA</span>
		<br>
		<table width="100%" cellpadding="2" cellspacing="2">
			<tr>	
			    <td rowspan="2" class="tdColSaida" align="left" valign="top" style="margin-left:2px; margin-right:2px">	
			        <fieldset style="height:400px; border: solid 1px #555; padding: auto"><legend><input id="cadastro" type="checkbox" onclick="marcarDesmarcarCadastro()"/><label for="cadastro">Cadastro</label></legend>	   
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_DATA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
				
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_DATA" name="ckb_COL_DATA"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_DATA.click();">Data</span><br />
					
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_NF|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_NF" name="ckb_COL_NF"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_NF.click();">NF</span><br />
						
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_DT_EMISSAO_NF|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_DT_EMISSAO_NF" name="ckb_COL_DT_EMISSAO_NF"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_DT_EMISSAO_NF.click();">Data Emiss�o NF</span><br />

				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_PEDIDO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_PEDIDO" name="ckb_COL_PEDIDO"
						        value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_PEDIDO.click();">Pedido</span><br />		
				
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_CPF_CNPJ_CLIENTE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_CPF_CNPJ_CLIENTE" name="ckb_COL_CPF_CNPJ_CLIENTE"
						value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_CPF_CNPJ_CLIENTE.click();">CPF/CNPJ Cliente</span><br />
				
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_NOME_CLIENTE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_NOME_CLIENTE" name="ckb_COL_NOME_CLIENTE"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_NOME_CLIENTE.click();">Nome Cliente</span><br />
				
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_CIDADE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
						    <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_CIDADE" name="ckb_COL_CIDADE"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_CIDADE.click();">Cidade</span><br />
			
			            <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_UF|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_UF" name="ckb_COL_UF"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_UF.click();">UF</span><br />

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
                        <hr />
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_INDICADOR_CPF_CNPJ|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_INDICADOR_CPF_CNPJ" name="ckb_COL_INDICADOR_CPF_CNPJ"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_INDICADOR_CPF_CNPJ.click();">CPF/CNPJ Indicador</span><br />
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_INDICADOR_ENDERECO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_CADASTRO" tabindex="-1" id="ckb_COL_INDICADOR_ENDERECO" name="ckb_COL_INDICADOR_ENDERECO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_INDICADOR_ENDERECO.click();">Endere�o</span><br />
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
							<input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_MARCA" name="ckb_COL_MARCA"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_MARCA.click();">Marca</span><br />
						
			            <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_GRUPO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_GRUPO" name="ckb_COL_GRUPO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_GRUPO.click();">Grupo</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_POTENCIA_BTU|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_POTENCIA_BTU" name="ckb_COL_POTENCIA_BTU"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_POTENCIA_BTU.click();">BTU/h</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_CICLO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_CICLO" name="ckb_COL_CICLO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_CICLO.click();">Ciclo</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_POSICAO_MERCADO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_POSICAO_MERCADO" name="ckb_COL_POSICAO_MERCADO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_POSICAO_MERCADO.click();">Posi��o Mercado</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_PRODUTO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_PRODUTO" name="ckb_COL_PRODUTO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_PRODUTO.click();">Produto</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_DESCRICAO_PRODUTO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_DESCRICAO_PRODUTO" name="ckb_COL_DESCRICAO_PRODUTO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_DESCRICAO_PRODUTO.click();">Descri��o Produto</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_QTDE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_QTDE" name="ckb_COL_QTDE"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_QTDE.click();">Quantidade</span><br />

                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_PERC_DESC|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_PERC_DESC" name="ckb_COL_PERC_DESC"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_PERC_DESC.click();">Percentual Desconto</span><br />
			
                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_CUBAGEM|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_CUBAGEM" name="ckb_COL_CUBAGEM"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_CUBAGEM.click();">Cubagem</span><br />

                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_PESO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_PESO" name="ckb_COL_PESO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_PESO.click();">Peso</span><br />

                        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_FRETE|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_COMERCIAL" tabindex="-1" id="ckb_COL_FRETE" name="ckb_COL_FRETE"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_FRETE.click();">Valor Frete</span><br />
					</fieldset>
				</td>
			</tr>
			<tr>
				<td class="tdColSaida" align="left" valign="middle" style="margin-left:2px; margin-right:2px">
				    <fieldset style="border: solid 1px #555;padding: auto"><legend><input id="financeiro" type="checkbox" onclick="marcarDesmarcarFinanceiro()" /><label for="financeiro">Financeiro</label></legend>
				    
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_CUSTO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_VL_CUSTO" name="ckb_COL_VL_CUSTO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_CUSTO.click();">VL Custo</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_LISTA|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
			    	        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_VL_LISTA" name="ckb_COL_VL_LISTA"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_LISTA.click();">VL Lista</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_VL_UNITARIO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_VL_UNITARIO" name="ckb_COL_VL_UNITARIO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_VL_UNITARIO.click();">VL Unit�rio</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_RT|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
                	        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_RT" name="ckb_COL_RT"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_RT.click();">RT</span><br />
						
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_QTDE_PARCELAS|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_QTDE_PARCELAS" name="ckb_COL_QTDE_PARCELAS"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_QTDE_PARCELAS.click();">Quantidade Parcelas</span><br />
				
				        <%	s_checked = ""
					        if (InStr(s_campos_saida_default, "|ckb_COL_MEIO_PAGAMENTO|") <> 0) Or (s_campos_saida_default = "") then s_checked = " checked" %>
					        <input type="checkbox" class="CKB_FINANCEIRO" tabindex="-1" id="ckb_COL_MEIO_PAGAMENTO" name="ckb_COL_MEIO_PAGAMENTO"
						    value="ON" <%=s_checked%> /><span class="C" style="cursor:default" onclick="fFILTRO.ckb_COL_MEIO_PAGAMENTO.click();">Meio de Pagamento</span><br />

					</fieldset>
				</td>
			</tr>
		</table>
		<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:8px;">
		<tr>
		<td align="left">
			<input name="bMarcarTodos" id="bMarcarTodos" type="button" class="Button" onclick="marcarTodos();" value="Marcar todos" title="assinala todos os campos de sa�da" style="margin-left:6px;margin-bottom:10px">
		</td>
		<td align="right">
			<input name="bDesmarcarTodos" id="bDesmarcarTodos" type="button" class="Button" onclick="desmarcarTodos();" value="Desmarcar todos" title="desmarca todos os campos de sa�da" style="margin-left:6px;margin-right:6px;margin-bottom:10px">
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
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a p�gina anterior">
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
