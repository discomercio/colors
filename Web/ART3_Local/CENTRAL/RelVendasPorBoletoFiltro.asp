<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================================
'	  RelVendasPorBoletoFiltro.asp
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

	Const COD_SAIDA_REL_VENDEDOR = "VENDEDOR"
	Const COD_SAIDA_REL_INDICADOR = "INDICADOR"
	Const COD_SAIDA_REL_INDICADORES_DO_VENDEDOR = "INDICADORES_DO_VENDEDOR"
	Const COD_SAIDA_REL_UF = "UF"
	Const COD_SAIDA_REL_ANALISTA_CREDITO = "ANALISTA_CREDITO"
	
	Const COD_ORDENACAO_VL_BOLETO = "ORD_POR_VL_BOLETO"
	Const COD_ORDENACAO_PERC_ATRASO = "ORD_POR_PERC_ATRASO"
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_VENDAS_POR_BOLETO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim intIdx, intOpcao
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

	dim strScript
	strScript = _
		"<script language='JavaScript'>" & chr(13) & _
		"var COD_SAIDA_REL_VENDEDOR = '" & COD_SAIDA_REL_VENDEDOR & "';" & chr(13) & _
		"var COD_SAIDA_REL_INDICADOR = '" & COD_SAIDA_REL_INDICADOR & "';" & chr(13) & _
		"var COD_SAIDA_REL_INDICADORES_DO_VENDEDOR = '" & COD_SAIDA_REL_INDICADORES_DO_VENDEDOR & "';" & chr(13) & _
		"var COD_SAIDA_REL_UF = '" & COD_SAIDA_REL_UF & "';" & chr(13) & _
		"var COD_SAIDA_REL_ANALISTA_CREDITO = '" & COD_SAIDA_REL_ANALISTA_CREDITO & "';" & chr(13) & _
		"</script>" & chr(13)




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
' LEMBRE-SE: O ORÇAMENTISTA É CONSIDERADO AUTOMATICAMENTE UM INDICADOR!!
function indicadores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT apelido, razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR ORDER BY apelido")
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
' ANALISTA CREDITO MONTA ITENS SELECT
function analista_credito_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT" & _
				" analise_credito_usuario," & _
				" nome_iniciais_em_maiusculas" & _
			" FROM (" & _
					"SELECT DISTINCT analise_credito_usuario FROM t_PEDIDO WHERE (analise_credito <> " & COD_AN_CREDITO_ST_INICIAL & ") AND (analise_credito <> " & COD_AN_CREDITO_NAO_ANALISADO & ")" & _
				") tP" & _
				" LEFT JOIN t_USUARIO tU ON (tP.analise_credito_usuario = tU.usuario)" & _
			" WHERE" & _
				"(LEN(analise_credito_usuario) > 0)" & _
			" ORDER BY" & _
				" analise_credito_usuario"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("analise_credito_usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x
		if Trim("" & r("nome_iniciais_em_maiusculas")) <> "" then strResp = strResp & " - "
		strResp = strResp & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
	else
		strResp = "<option value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	analista_credito_monta_itens_select = strResp
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

<%=strScript%>

<script type="text/javascript">
	$(function() {
		$("input[type=radio]").hUtil('fix_radios');
		$("#c_dt_entregue_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_entregue_termino").hUtilUI('datepicker_filtro_final');
	});
</script>


<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
var s_de, s_ate;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;
var i, blnFlagOk;
var radioButtonSelecionado;

//  COLUNA DE SAÍDA DO RELATÓRIO
	blnFlagOk=false;
	for (i=0; i<f.rb_saida.length; i++) {
		if (f.rb_saida[i].checked) blnFlagOk=true;
		}
	if (!blnFlagOk) {
		alert("Selecione o tipo de saída do relatório!!");
		return;
		}

//  PERÍODO DE ENTREGA
	if ((trim(f.c_dt_entregue_inicio.value) == "") && (trim(f.c_dt_entregue_termino.value) == "")) {
		alert("Informe o período de entrega!!");
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
	
	radioButtonSelecionado = $('#fFILTRO input[name=rb_saida]:checked').val();
	if (radioButtonSelecionado == COD_SAIDA_REL_INDICADORES_DO_VENDEDOR) {
		if (trim(fFILTRO.c_indicadores_do_vendedor.value) == "") {
			alert("Selecione um vendedor!!");
			fFILTRO.c_indicadores_do_vendedor.focus();
			return;
		}
	}
	
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
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


<body onload="fFILTRO.c_dt_entregue_inicio.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelVendasPorBoletoExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Vendas por Boleto</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<% intIdx=-1 %>
<table class="Qx" cellspacing="0">
<!--  ENTREGUE ENTRE  -->
	<tr bgcolor="#FFFFFF">
	<td class="MT" align="left" nowrap>
		<table cellspacing="0" cellpadding="0">
		<tr>
			<td align="left" valign="bottom">
				<span class="PLTe" style="cursor:default">ENTREGUES ENTRE</span>
			</td>
		</tr>
		</table>
		<table cellspacing="0" cellpadding="0"><tr bgcolor="#FFFFFF"><td align="left">
		<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio"
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entregue_termino.focus(); filtra_data();"
			>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;até&nbsp;</span>&nbsp;
			<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_entregue_termino" id="c_dt_entregue_termino"
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_loja.focus(); filtra_data();"
			>
			</td></tr>
		</table>
		</td>
	</tr>

<!--  LOJA  -->
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap>
			<table cellspacing="0" cellpadding="0">
			<tr>
				<td align="left" valign="bottom">
					<span class="PLTe" style="cursor:default">LOJA</span>
				</td>
			</tr>
			</table>
			<table cellspacing="0" cellpadding="0">
				<tr bgcolor="#FFFFFF">
				<td align="left">
					<input class="PLLe" maxlength="3" style="width:100px;" name="c_loja" id="c_loja" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fFILTRO.c_vendedor.focus(); else this.click(); filtra_numerico();" />
				</td>
				</tr>
			</table>
		</td>
	</tr>

<!--  VENDEDOR  -->
	<% intIdx=intIdx+1 %>
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap>
			<table cellspacing="0" cellpadding="0">
			<tr>
				<td align="left">
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_VENDEDOR%>">
				</td>
				<td align="left" valign="bottom">
				<span class="PLTe" style="cursor:default" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">VENDEDOR</span>
				</td>
			</tr>
			</table>
			<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF">
				<td align="left">
					<select id="c_vendedor" name="c_vendedor" style="cursor:default;margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" onkeypress="if (digitou_enter(true)) fFILTRO.c_indicador.focus();" onchange="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">
					<% =vendedores_monta_itens_select(Null) %>
					</select>
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  INDICADOR  -->
	<% intIdx=intIdx+1 %>
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap>
			<table cellspacing="0" cellpadding="0">
			<tr>
				<td align="left">
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_INDICADOR%>">
				</td>
				<td align="left" valign="bottom">
				<span class="PLTe" style="cursor:default" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">INDICADOR</span>
				</td>
			</tr>
			</table>
			<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF">
				<td align="left">
					<select id="c_indicador" name="c_indicador" style="cursor:default;margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" onkeypress="if (digitou_enter(true)) fFILTRO.c_indicadores_do_vendedor.focus();" onchange="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">
					<% =indicadores_monta_itens_select(Null) %>
					</select>
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  INDICADORES DO VENDEDOR  -->
	<% intIdx=intIdx+1 %>
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap>
			<table cellspacing="0" cellpadding="0">
			<tr>
				<td align="left">
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_INDICADORES_DO_VENDEDOR%>">
				</td>
				<td align="left" valign="bottom">
				<span class="PLTe" style="cursor:default" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">INDICADORES DO VENDEDOR</span>
				</td>
			</tr>
			</table>
			<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF">
				<td align="left">
					<select id="c_indicadores_do_vendedor" name="c_indicadores_do_vendedor" style="cursor:default;margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" onkeypress="if (digitou_enter(true)) fFILTRO.c_analista.focus();" onchange="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">
					<% =vendedores_monta_itens_select(Null) %>
					</select>
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  UF  -->
	<% intIdx=intIdx+1 %>
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap>
			<table cellspacing="0" cellpadding="0">
			<tr>
				<td align="left">
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_UF%>">
				</td>
				<td align="left" valign="bottom">
				<span class="PLTe" style="cursor:default" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">UF</span>
				</td>
			</tr>
			</table>
			<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF">
				<td style="width:12px;" align="left">&nbsp;</td>
				<td align="left">
				<span class="PLTe" style="cursor:default" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">TODAS</span>
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  ANALISTA DE CRÉDITO  -->
	<% intIdx=intIdx+1 %>
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap>
			<table cellspacing="0" cellpadding="0">
			<tr>
				<td align="left">
					<input type="radio" id="rb_saida" name="rb_saida" value="<%=COD_SAIDA_REL_ANALISTA_CREDITO%>">
				</td>
				<td align="left" valign="bottom">
				<span class="PLTe" style="cursor:default" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">ANALISTA</span>
				</td>
			</tr>
			</table>
			<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF">
				<td align="left">
					<select id="c_analista" name="c_analista" style="cursor:default;margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus();" onchange="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">
					<% =analista_credito_monta_itens_select(Null) %>
					</select>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left">
			<table cellspacing="0" cellpadding="0" width="100%">
				<tr bgcolor="#FFFFFF">
					<td class="MD" align="left" valign="top" width="33%" nowrap>
						<span class="PLTe">TIPO DE CLIENTE</span>
						<br />
						<% intOpcao=-1 %>
						<input type="radio" id="rb_tipo_cliente" name="rb_tipo_cliente" value=<%=ID_PF%> style="margin-left:20px;">
						<% intOpcao=intOpcao+1 %>
						<span style="cursor:default" class="Np" onclick="fFILTRO.rb_tipo_cliente[<%=Cstr(intOpcao)%>].click();">Pessoa Física</span>
						<br />
						<input type="radio" id="rb_tipo_cliente" name="rb_tipo_cliente" value=<%=ID_PJ%> style="margin-left:20px;">
						<% intOpcao=intOpcao+1 %>
						<span style="cursor:default" class="Np" onclick="fFILTRO.rb_tipo_cliente[<%=Cstr(intOpcao)%>].click();">Pessoa Jurídica</span>
						<br />
						<input type="radio" id="rb_tipo_cliente" name="rb_tipo_cliente" value="" style="margin-left:20px;" checked>
						<% intOpcao=intOpcao+1 %>
						<span style="cursor:default" class="Np" onclick="fFILTRO.rb_tipo_cliente[<%=Cstr(intOpcao)%>].click();">Ambos</span>
					</td>
					<td class="MD" align="left" valign="top" width="33%">
						<span class="PLTe">UF</span>
						<br />
						<select id="c_uf" name="c_uf" style="width:60px;margin-left:20px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}">
						<% =UF_monta_itens_select(Null) %>
						</select>
					</td>
					<td align="left" valign="top" nowrap>
						<span class="PLTe">ORDENAÇÃO</span>
						<br />
						<% intOpcao=-1 %>
						<input type="radio" id="rb_ordenacao" name="rb_ordenacao" value=<%=COD_ORDENACAO_VL_BOLETO%> style="margin-left:20px;" checked>
						<% intOpcao=intOpcao+1 %>
						<span style="cursor:default;margin-right:8px;" class="Np" onclick="fFILTRO.rb_ordenacao[<%=Cstr(intOpcao)%>].click();">coluna [VL Boleto (<%=SIMBOLO_MONETARIO%>)]</span>
						<br />
						<input type="radio" id="rb_ordenacao" name="rb_ordenacao" value=<%=COD_ORDENACAO_PERC_ATRASO%> style="margin-left:20px;">
						<% intOpcao=intOpcao+1 %>
						<span style="cursor:default" class="Np" onclick="fFILTRO.rb_ordenacao[<%=Cstr(intOpcao)%>].click();">coluna [% Atraso]</span>
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
