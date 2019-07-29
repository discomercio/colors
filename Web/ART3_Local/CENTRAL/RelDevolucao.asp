<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  R E L D E V O L U C A O . A S P
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
function fFILTROConfirma( f ) {
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;

	if (f.ckb_periodo_devolucao.checked) {
		if (!consiste_periodo(f.c_dt_devolucao_inicio, f.c_dt_devolucao_termino)) return;
		}

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
		
	if (f.rb_loja[1].checked) {
		if (converte_numero(f.c_loja.value)==0) {
			alert("Especifique o número da loja!!");
			f.c_loja.focus();
			return;
			}
		}

	if (f.rb_loja[2].checked) {
		if (trim(f.c_loja_de.value)!="") {
			if (converte_numero(f.c_loja_de.value)==0) {
				alert("Número de loja inválido!!");
				f.c_loja_de.focus();
				return;
				}
			}
		if (trim(f.c_loja_ate.value)!="") {
			if (converte_numero(f.c_loja_ate.value)==0) {
				alert("Número de loja inválido!!");
				f.c_loja_ate.focus();
				return;
				}
			}
		if ((trim(f.c_loja_de.value)=="")&&(trim(f.c_loja_ate.value)=="")) {
			alert("Preencha pelo menos um dos campos!!");
			f.c_loja_de.focus();
			return;
			}
		if ((trim(f.c_loja_de.value)!="")&&(trim(f.c_loja_ate.value)!="")) {
			if (converte_numero(f.c_loja_ate.value)<converte_numero(f.c_loja_de.value)) {
				alert("Faixa de lojas inválida!!");
				f.c_loja_ate.focus();
				return;
				}
			}
		}

//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
		strDtRefDDMMYYYY = trim(f.c_dt_devolucao_inicio.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_devolucao_termino.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
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

<form id="fFILTRO" name="fFILTRO" method="post" action="RelDevolucaoExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Devolução de Produtos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table class="Qx" cellSpacing="0">
<!--  PERÍODO DE DEVOLUÇÃO  -->
<tr bgColor="#FFFFFF">
<td class="MT" NOWRAP><span class="PLTe">PERÍODO DE DEVOLUÇÃO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_periodo_devolucao" name="ckb_periodo_devolucao"
			value="PERIODO_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_periodo_devolucao.click();">Devolvido entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_devolucao_inicio" id="c_dt_devolucao_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_devolucao_termino.focus(); else fFILTRO.ckb_periodo_devolucao.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_devolucao.checked=true;"
			>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_devolucao_termino" id="c_dt_devolucao_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cadastro_inicio.focus(); else fFILTRO.ckb_periodo_devolucao.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_devolucao.checked=true;">
		</td></tr>
	</table>
</td></tr>

<!--  PERÍODO  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">PERÍODO</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="checkbox" tabindex="-1" id="ckb_periodo_cadastro" name="ckb_periodo_cadastro"
			value="PERIODO_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_periodo_cadastro.click();">Somente pedidos colocados entre</span
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
		<input type="checkbox" tabindex="-1" id="ckb_produto" name="ckb_produto"
			value="PRODUTO_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_produto.click();">Produto devolvido</span
			><br><span class="C" style="margin-left:30px;">Fabricante</span><input maxlength="4" class="Cc" style="width:50px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto.focus(); else fFILTRO.ckb_produto.checked=true; filtra_fabricante();" onclick="fFILTRO.ckb_produto.checked=true;">
			<span class="C">&nbsp;&nbsp;&nbsp;Produto</span><input maxlength="13" class="Cc" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_produto.checked=true; filtra_produto();" onclick="fFILTRO.ckb_produto.checked=true;">
		</td></tr>
	</table>
</td></tr>

<!--  LOJAS  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">LOJAS</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="radio" tabindex="-1" id="rb_loja" name="rb_loja"
			value="TODAS" checked><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_loja[0].click();">Todas as lojas</span>
		</td></tr>
	<tr bgColor="#FFFFFF"><td>
		<input type="radio" tabindex="-1" id="rb_loja" name="rb_loja"
			value="UMA"><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_loja[1].click();">Loja</span>
			<input class="Cc" maxlength="3" style="width:40px;" name="c_loja" id="c_loja" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); else fFILTRO.rb_loja[1].click(); filtra_numerico();" onclick="fFILTRO.rb_loja[1].click();">
		</td></tr>
	<tr bgColor="#FFFFFF"><td>
		<input type="radio" tabindex="-1" id="rb_loja" name="rb_loja"
			value="FAIXA"><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_loja[2].click();">Lojas</span>
			<input class="Cc" maxlength="3" style="width:40px;" name="c_loja_de" id="c_loja_de" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fFILTRO.c_loja_ate.focus(); else fFILTRO.rb_loja[2].click(); filtra_numerico();" onclick="fFILTRO.rb_loja[2].click();">
			<span class="C">a</span>
			<input class="Cc" maxlength="3" style="width:40px;" name="c_loja_ate" id="c_loja_ate" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="fFILTRO.rb_loja[2].click(); if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); filtra_numerico();" onclick="fFILTRO.rb_loja[2].click();">
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
