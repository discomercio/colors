<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  R E L O R C A M E N T O S M C R I T . A S P
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

	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;

	if (f.ckb_periodo_cadastro.checked) {
		if (trim(f.c_dt_cadastro_inicio.value)=="") {
			alert("Preencha a data!!");
			f.c_dt_cadastro_inicio.focus();
			return;
			}
		if (trim(f.c_dt_cadastro_termino.value)=="") {
			alert("Preencha a data!!");
			f.c_dt_cadastro_termino.focus();
			return;
			}
		if (!consiste_periodo(f.c_dt_cadastro_inicio, f.c_dt_cadastro_termino)) return;
		}
	
	if (f.ckb_produto.checked) {
		if (trim(f.c_produto.value)!="") {
			if (!isEAN(f.c_produto.value)) {
				if (trim(f.c_fabricante.value)=="") {
					alert("Preencha o código do fabricante!!");
					f.c_fabricante.focus();
					return;
					}
				}
			}
		if ((trim(f.c_produto.value)=="")&&(trim(f.c_fabricante.value)=="")) {
			alert("Preencha o código do produto!!");
			f.c_produto.focus();
			return;
			}
		}
	
//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
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


<body>
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelOrcamentosMCritExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório Multicritério de Pré-Pedidos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table class="Qx" cellSpacing="0">

<!--  STATUS DO ORÇAMENTO  -->
<tr bgColor="#FFFFFF">
<td class="MT" NOWRAP><span class="PLTe">STATUS DO PRÉ-PEDIDO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_orcamento_em_aberto" name="ckb_orcamento_em_aberto"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_orcamento_em_aberto.click();">Pré-Pedido em aberto</span>
		</td></tr>
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_orcamento_virou_pedido" name="ckb_orcamento_virou_pedido"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_orcamento_virou_pedido.click();">Pré-Pedido virou pedido</span>
		</td></tr>
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_orcamento_cancelado" name="ckb_orcamento_cancelado"
			value="<%=ST_ORCAMENTO_CANCELADO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_orcamento_cancelado.click();">Cancelado</span>
		</td></tr>
	</table>
</td></tr>

<!--  PERÍODO  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">PERÍODO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_periodo_cadastro" name="ckb_periodo_cadastro" onclick="if (fFILTRO.ckb_periodo_cadastro.checked) fFILTRO.c_dt_cadastro_inicio.focus();"
			value="PERIODO_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_periodo_cadastro.click();">Somente pré-pedidos cadastrados entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cadastro_inicio" id="c_dt_cadastro_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cadastro_termino.focus(); else fFILTRO.ckb_periodo_cadastro.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_cadastro.checked=true;"
			>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cadastro_termino" id="c_dt_cadastro_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_periodo_cadastro.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_cadastro.checked=true;">
		</td></tr>
	</table>
</td></tr>

<!--  PRODUTO  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">PRODUTO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_produto" name="ckb_produto" onclick="if (fFILTRO.ckb_produto.checked) fFILTRO.c_fabricante.focus();"
			value="PRODUTO_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_produto.click();">Somente orçamentos que incluam:</span
			><br><span class="C" style="margin-left:30px;">Fabricante</span><input maxlength="4" class="Cc" style="width:50px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto.focus(); else fFILTRO.ckb_produto.checked=true; filtra_fabricante();" onclick="fFILTRO.ckb_produto.checked=true;">
			<span class="C">&nbsp;&nbsp;&nbsp;Produto</span><input maxlength="13" class="Cc" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_produto.checked=true; filtra_produto();" onclick="fFILTRO.ckb_produto.checked=true;">
		</td></tr>
	</table>
</td></tr>

<!--  Nº ORÇAMENTO  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">PRÉ-PEDIDO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<span class="C" style="margin-left:30px;">Nº Pré-Pedido</span>
			<input class="C" maxlength="10" style="width:70px;" name="c_orcamento" id="c_orcamento" onblur="if (normaliza_num_orcamento(this.value)!='') this.value=normaliza_num_orcamento(this.value);" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {if (normaliza_num_orcamento(this.value)!='') this.value=normaliza_num_orcamento(this.value); bCONFIRMA.focus();} filtra_orcamento();">
		</td></tr>
	</table>
</td></tr>

<!--  CLIENTE  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">CLIENTE</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<span class="C" style="margin-left:30px;">CNPJ/CPF</span>
			<input class="C" maxlength="18" style="width:140px;" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="if (digitou_enter(true)&&((!tem_info(this.value))||(tem_info(this.value)&&cnpj_cpf_ok(this.value)))) {this.value=cnpj_cpf_formata(this.value); bCONFIRMA.focus();} filtra_cnpj_cpf();">
		</td></tr>
	</table>
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
