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
'	  E S T O Q U E C O N S U L T A M C R I T . A S P
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not (operacao_permitida(OP_CEN_REL_REGISTROS_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) Or _
			operacao_permitida(OP_CEN_EDITA_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas)) then 
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
		$("#c_entrada_de").hUtilUI('datepicker_filtro_inicial');
		$("#c_entrada_ate").hUtilUI('datepicker_filtro_final');
	});
</script>

<script language="JavaScript" type="text/javascript">
function fESTOQConsulta( f ) {
var f, s_de, s_ate, i, b;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;
	f=fESTOQ;
	s_de=trim(f.c_entrada_de.value);
	s_ate=trim(f.c_entrada_ate.value);
	if ((s_de!="")&&(s_ate!="")) {
		s_de=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
		s_ate=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
		if (s_de > s_ate) {
			alert("Data final é menor que a data inicial!!");
			f.c_entrada_ate.focus();
			return;
			}
		}
		
	b=false;
	if (f.ckb_compras.checked) b=true;
	if (f.ckb_especial.checked) b=true;
	if (f.ckb_kit.checked) b=true;
	if (f.ckb_devolucao.checked) b=true;
	if (!b) {
		alert("Selecione pelo menos um tipo de cadastramento!!");
		return;
		}
	
//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
		strDtRefDDMMYYYY = trim(f.c_entrada_de.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_entrada_ate.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
#ckb_especial {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
#ckb_saldo {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
#ckb_compras {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
#ckb_kit {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
#ckb_devolucao {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
#rb_saida {
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


<body onload="if (trim(fESTOQ.c_fabricante.value)=='') fESTOQ.c_fabricante.focus();">
<center>

<form id="fESTOQ" name="fESTOQ" METHOD="POST" ACTION="EstoqueConsultaMCritExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<INPUT type=HIDDEN name='c_MinDtInicialFiltroPeriodoYYYYMMDD' id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<INPUT type=HIDDEN name='c_MinDtInicialFiltroPeriodoDDMMYYYY' id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="RIGHT" vAlign="BOTTOM"><span class="PEDIDO">Registros Entrada Estoque</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS DA CONSULTA MULTICRITÉRIOS  -->
<table class="Qx" cellspacing="0">
<!-- EMPRESA -->
    <tr bgcolor="#FFFFFF">
        <td colspan="2" class="MT" NOWRAP><span class="PLTe">Empresa</span>
            <br>
			<select id="c_empresa" name="c_empresa" style="margin:1px 10px 6px 5px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
			</select>
        </td>
    </tr>

<!--  FABRICANTE/PRODUTO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">Fabricante</span>
		<br><input name="c_fabricante" id="c_fabricante" class="PLLe" maxlength="4" style="margin-left:2pt;width:120px;" onkeypress="if (digitou_enter(true)) fESTOQ.c_produto.focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"></td>
	<td class="MDBE" align="left" style="border-left:0pt;"><span class="PLTe">Produto</span>
		<br><input name="c_produto" id="c_produto" class="PLLe" maxlength="13" style="margin-left:2pt;width:160px;" onkeypress="if (digitou_enter(true)) fESTOQ.c_cadastrado_por.focus(); filtra_produto();" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));"></td>
	</tr>

<!--  CADASTRADO POR  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Cadastrado por</span>
		<br><input name="c_cadastrado_por" id="c_cadastrado_por" class="PLLe" maxlength="10" style="margin-left:2pt;width:150px;" onkeypress="if (digitou_enter(true)) fESTOQ.c_entrada_de.focus();" onblur="this.value=ucase(trim(this.value));"></td>
	</tr>

<!--  PERÍODO DE ENTRADA NO ESTOQUE  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MDBE" align="left" valign="bottom" nowrap><span class="PLTe">Data de Entrada no Estoque Entre</span>
		<br><input name="c_entrada_de" id="c_entrada_de" class="PLLc" maxlength="10" style="margin-left:2pt;width:80px;" onkeypress="if (digitou_enter(true)) fESTOQ.c_entrada_ate.focus(); filtra_data();" onblur="this.value=trim(this.value); if (!isDate(this)) {alert('Data inválida!!'); this.focus();}"
			><span class="PLTe" style="vertical-align:baseline;">&nbsp;&nbsp;e&nbsp;</span><input name="c_entrada_ate" id="c_entrada_ate" class="PLLc" maxlength="10" style="margin-left:2pt;width:80px;" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_data();" onblur="this.value=trim(this.value); if (!isDate(this)) {alert('Data inválida!!'); this.focus();}"></td>
	</tr>

<!--  TIPO DE CADASTRAMENTO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Tipo de Cadastramento</span>
		<!--  COMPRAS  -->
		<br><input type="checkbox" tabindex="-1" id="ckb_compras" name="ckb_compras" value="COMPRAS_ON" checked
		><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.ckb_compras.click();">Compras de Fornecedor</span>
		<!--  ENTRADA ESPECIAL  -->
		<br><input type="checkbox" tabindex="-1" id="ckb_especial" name="ckb_especial" value="ESPECIAL_ON"
		><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.ckb_especial.click();">Entrada Especial</span>
		<!--  KIT  -->
		<br><input type="checkbox" tabindex="-1" id="ckb_kit" name="ckb_kit" value="KIT_ON"
		><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.ckb_kit.click();">Kit</span>
		<!--  DEVOLUÇÃO  -->
		<br><input type="checkbox" tabindex="-1" id="ckb_devolucao" name="ckb_devolucao" value="DEVOLUCAO_ON"
		><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.ckb_devolucao.click();">Devolução</span>
	</td>
	</tr>

<!--  SOMENTE PRODUTOS COM SALDO DISPONÍVEL  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Saldo de Produtos</span>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="ckb_saldo" name="ckb_saldo" value="TODOS" checked
		><span class="C lblOpt" style="cursor:default;margin-right:10pt;" onclick="fESTOQ.ckb_saldo[0].click();">Todos</span>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="ckb_saldo" name="ckb_saldo" value="COM_SALDO"
		><span class="C lblOpt" style="cursor:default;margin-right:10pt;" onclick="fESTOQ.ckb_saldo[1].click();">Somente Produtos Com Saldo Disponível</span>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="ckb_saldo" name="ckb_saldo" value="SEM_SALDO"
		><span class="C lblOpt" style="cursor:default;margin-right:10pt;" onclick="fESTOQ.ckb_saldo[2].click();">Somente Produtos Sem Saldo Disponível</span>
	</td>
	</tr>

<!--  SAÍDA DO RELATÓRIO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Saída do Relatório</span>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="Html" onclick="dCONFIRMA.style.visibility='';" checked><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_saida[0].click();dCONFIRMA.style.visibility='';"
			>Html</span>

		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="XLS" onclick="dCONFIRMA.style.visibility='';"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_saida[1].click();dCONFIRMA.style.visibility='';"
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
	<td align="left"><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
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

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
