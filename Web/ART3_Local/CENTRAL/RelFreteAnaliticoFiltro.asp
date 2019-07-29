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
'	  RelFreteAnaliticoFiltro.asp
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

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script type="text/javascript">
    $(function() {
        $("#c_dt_entregue_inicio").hUtilUI('datepicker_filtro_inicial');
        $("#c_dt_entregue_termino").hUtilUI('datepicker_filtro_final');
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<body onload="fFILTRO.c_dt_entregue_inicio.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelFreteAnaliticoExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Frete (Analítico)</span>
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
	<td class="MT" align="left" NOWRAP>
		<table cellSpacing="0" cellPadding="0"><tr bgColor="#FFFFFF"><td>
		<span class="PLTe" style="cursor:default">ENTREGUES ENTRE</span>
		<br>
		<input class="PLLc" maxlength="10" style="width:90px;" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio"
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entregue_termino.focus(); filtra_data();"
			>&nbsp;<span class="PLLc" style="color:#808080;margin-left:10px;">&nbsp;até&nbsp;</span>&nbsp;
			<input class="PLLc" maxlength="10" style="width:90px;" name="c_dt_entregue_termino" id="c_dt_entregue_termino"
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_transportadora.focus(); filtra_data();" 
			>
			</td></tr>
		</table>
		</td></tr>

<!--  TRANSPORTADORA  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" align="left" NOWRAP><span class="PLTe">TRANSPORTADORA</span>
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
	<td class="MDBE" align="left" NOWRAP><span class="PLTe">FABRICANTE</span>
	<br>
		<input maxlength="4" class="PLLe" style="width:150px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_loja.focus(); filtra_fabricante();">
		</td></tr>

<!--  LOJA  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" align="left" NOWRAP><span class="PLTe">LOJA</span>
	<br>
		<input class="PLLe" maxlength="3" style="width:150px;" name="c_loja" id="c_loja" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fFILTRO.c_vendedor.focus(); filtra_numerico();">
		</td></tr>

<!--  VENDEDOR  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" align="left" NOWRAP><span class="PLTe">VENDEDOR</span>
	<br>
		<input maxlength="10" class="PLLe" style="width:150px;" name="c_vendedor" id="c_vendedor" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_indicador.focus(); filtra_nome_identificador();">
		</td></tr>

<!--  INDICADOR  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" align="left" NOWRAP><span class="PLTe">INDICADOR</span>
		<br>
			<select id="c_indicador" name="c_indicador" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =indicadores_monta_itens_select(Null) %>
			</select>
			</td></tr>

<!--  UF  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" align="left" NOWRAP><span class="PLTe">UF</span>
		<br>
			<select id="c_uf" name="c_uf" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =uf_monta_itens_select(Null) %>
			</select>
			</td></tr>
			
<!--  STATUS DO FRETE  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" align="left" NOWRAP><span class="PLTe">STATUS DO FRETE</span>
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
		<td class="MDBE" align="left" NOWRAP><span class="PLTe">SAÍDA DO RELATÓRIO</span>
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
