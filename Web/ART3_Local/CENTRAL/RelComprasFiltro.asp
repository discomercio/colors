<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  RelComprasFiltro.asp
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

	dim usuario, s, intIdx
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if (Not operacao_permitida(OP_CEN_REL_COMPRAS, s_lista_operacoes_permitidas)) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

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



<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fFILTROConsulta( f ) {
var i, b;
var s_de, s_ate;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;

//  PERÍODO
	if (trim(f.c_dt_inicio.value)=="") {
		alert("Informe a data inicial do período!!");
		f.c_dt_inicio.focus();
		return;
		}
	
	if (trim(f.c_dt_termino.value)=="") {
		alert("Informe a data final do período!!");
		f.c_dt_termino.focus();
		return;
		}
		
	if (trim(f.c_dt_inicio.value)!="") {
		if (!isDate(f.c_dt_inicio)) {
			alert("Data inválida!!");
			f.c_dt_inicio.focus();
			return;
			}
		}

	if (trim(f.c_dt_termino.value)!="") {
		if (!isDate(f.c_dt_termino)) {
			alert("Data inválida!!");
			f.c_dt_termino.focus();
			return;
			}
		}

	s_de = trim(f.c_dt_inicio.value);
	s_ate = trim(f.c_dt_termino.value);
	if ((s_de!="")&&(s_ate!="")) {
		s_de=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
		s_ate=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
		if (s_de > s_ate) {
			alert("Data de término é menor que a data de início!!");
			f.c_dt_termino.focus();
			return;
			}
		}

//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
	// PERÍODO
		strDtRefDDMMYYYY = trim(f.c_dt_inicio.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_termino.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">

<style type="text/css">
#rb_detalhe {
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


<body onload="if (trim(fFILTRO.c_dt_inicio.value)=='') fFILTRO.c_dt_inicio.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelComprasExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Compras</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS DA CONSULTA  -->
<table class="Qx" cellSpacing="0">
<!--  PERÍODO  -->
	<tr bgColor="#FFFFFF">
	<td class="MT" colspan="2" NOWRAP>
		<table cellSpacing="0" cellPadding="0"><tr bgColor="#FFFFFF"><td>
		<span class="PLTe" style="cursor:default">Período</span>
		<br>
		<input class="PLLc" maxlength="10" style="width:100px;" name="c_dt_inicio" id="c_dt_inicio" 
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_termino.focus(); filtra_data();"
			>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;até&nbsp;</span>&nbsp;
			<input class="PLLc" maxlength="10" style="width:100px;" name="c_dt_termino" id="c_dt_termino" 
			onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" 
			onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); filtra_data();" 
			>
			</td></tr>
		</table>
		</td></tr>

<!--  FABRICANTE/PRODUTO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">Fabricante</span>
		<br><input name="c_fabricante" id="c_fabricante" class="PLLe" maxlength="4" style="margin-left:2pt;width:50px;" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto.focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"></td>
	<td class="MDB" style="border-left:0pt;"><span class="PLTe">Produto</span>
		<br><input name="c_produto" id="c_produto" class="PLLe" maxlength="13" style="margin-left:2pt;width:100px;" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_produto();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_PRODUTO); this.value=ucase(trim(this.value));"></td>
	</tr>

<!--  TIPO DE DETALHAMENTO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MDBE" NOWRAP><span class="PLTe">Tipo de Detalhamento</span>
		<% intIdx = -1 %>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="SINTETICO_FABR"><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Sintético por Fabricante</span>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="SINTETICO_PROD"><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Sintético por Produto</span>
		
		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="CUSTO_MEDIO"><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Custo Médio</span>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="CUSTO_INDIVIDUAL"><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Custo Individual</span>
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
	<td><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConsulta(fFILTRO)" title="executa a consulta">
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
