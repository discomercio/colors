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
'	  RelPesquisaOrdemServicoFiltro.asp
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_REL_PESQUISA_ORDEM_SERVICO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
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
function fFILTROConfirma(f) {
var s_de, s_ate;
var blnTemRestricaoPeriodo = false;

	if ((trim(f.c_dt_abertura_inicio.value) != "") || (trim(f.c_dt_abertura_termino.value) != "")) {
		blnTemRestricaoPeriodo = true;
		if (trim(f.c_dt_abertura_inicio.value) == "") {
			alert("Informe a data de início para o período de abertura!!");
			f.c_dt_abertura_inicio.focus();
			return;
		}
		if (trim(f.c_dt_abertura_termino.value) == "") {
			alert("Informe a data de término para o período de abertura!!");
			f.c_dt_abertura_termino.focus();
			return;
		}
	}

	if ((trim(f.c_dt_encerramento_inicio.value) != "") || (trim(f.c_dt_encerramento_termino.value) != "")) {
		blnTemRestricaoPeriodo = true;
		if (trim(f.c_dt_encerramento_inicio.value) == "") {
			alert("Informe a data de início para o período de encerramento!!");
			f.c_dt_encerramento_inicio.focus();
			return;
		}
		if (trim(f.c_dt_encerramento_termino.value) == "") {
			alert("Informe a data de término para o período de encerramento!!");
			f.c_dt_encerramento_termino.focus();
			return;
		}
	}

	if (!blnTemRestricaoPeriodo) {
		alert("É necessário especificar pelo menos um dos períodos de consulta!!");
		return;
	}
	
	if (trim(f.c_dt_abertura_inicio.value) != "") {
		if (!isDate(f.c_dt_abertura_inicio)) {
			alert("Data de início para o período de abertura é inválida!!");
			f.c_dt_abertura_inicio.focus();
			return;
		}
	}

	if (trim(f.c_dt_abertura_termino.value) != "") {
		if (!isDate(f.c_dt_abertura_termino)) {
			alert("Data de término para o período de abertura é inválida!!");
			f.c_dt_abertura_termino.focus();
			return;
		}
	}
	if (trim(f.c_dt_encerramento_inicio.value) != "") {
		if (!isDate(f.c_dt_encerramento_inicio)) {
			alert("Data de início para o período de encerramento é inválida!!");
			f.c_dt_encerramento_inicio.focus();
			return;
		}
	}

	if (trim(f.c_dt_encerramento_termino.value) != "") {
		if (!isDate(f.c_dt_encerramento_termino)) {
			alert("Data de término para o período de encerramento é inválida!!");
			f.c_dt_encerramento_termino.focus();
			return;
		}
	}

	s_de = trim(f.c_dt_abertura_inicio.value);
	s_ate = trim(f.c_dt_abertura_termino.value);
	if ((s_de != "") && (s_ate != "")) {
		s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
		s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
		if (s_de > s_ate) {
			alert("Data de término é menor que a data de início para o período de abertura!!");
			f.c_dt_abertura_termino.focus();
			return;
		}
	}

	s_de = trim(f.c_dt_encerramento_inicio.value);
	s_ate = trim(f.c_dt_encerramento_termino.value);
	if ((s_de != "") && (s_ate != "")) {
		s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
		s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
		if (s_de > s_ate) {
			alert("Data de término é menor que a data de início para o período de encerramento!!");
			f.c_dt_encerramento_termino.focus();
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">


<body onload="fFILTRO.c_dt_abertura_inicio.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelPesquisaOrdemServicoExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pesquisa Ordem de Serviço</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0" style="width:240px;">
<!--  PERÍODO DE ABERTURA  -->
	<tr>
		<td class="ME MD MC" NOWRAP><span class="PLTe">PERÍODO DE ABERTURA DA O.S.</span></td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MD MB">
			<table cellSpacing="0" cellPadding="0" style="margin:0px 20px 6px 30px;">
			<tr bgColor="#FFFFFF">
				<td>
					<input class="Cc" maxlength="10" style="width:80px;" name="c_dt_abertura_inicio" id="c_dt_abertura_inicio" onblur="if (!isDate(this)) {alert('Data inválida!!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_abertura_termino.focus(); filtra_data();"
						>&nbsp;
					<span class="C">&nbsp;até&nbsp;</span>&nbsp;
					<input class="Cc" maxlength="10" style="width:80px;" name="c_dt_abertura_termino" id="c_dt_abertura_termino" onblur="if (!isDate(this)) {alert('Data inválida!!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_encerramento_inicio.focus(); filtra_data();">
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  PERÍODO DE ENCERRAMENTO  -->
	<tr>
		<td class="ME MD" NOWRAP><span class="PLTe">PERÍODO DE ENCERRAMENTO DA O.S.</span></td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MD MB">
			<table cellSpacing="0" cellPadding="0" style="margin:0px 20px 6px 30px;">
			<tr bgColor="#FFFFFF">
				<td>
					<input class="Cc" maxlength="10" style="width:80px;" name="c_dt_encerramento_inicio" id="c_dt_encerramento_inicio" onblur="if (!isDate(this)) {alert('Data inválida!!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_encerramento_termino.focus(); filtra_data();"
						>&nbsp;
					<span class="C">&nbsp;até&nbsp;</span>&nbsp;
					<input class="Cc" maxlength="10" style="width:80px;" name="c_dt_encerramento_termino" id="c_dt_encerramento_termino" onblur="if (!isDate(this)) {alert('Data inválida!!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); filtra_data();">
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  FABRICANTE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">FABRICANTE</span>
	<br>
		<input maxlength="4" class="PLLe" style="width:100px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto.focus(); filtra_fabricante();">
		</td></tr>

<!--  PRODUTO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">PRODUTO</span>
	<br>
		<input maxlength="13" class="PLLe" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) fFILTRO.c_pedido.focus(); filtra_produto();">
		</td></tr>

<!--  PEDIDO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">PEDIDO</span>
	<br>
		<input maxlength="10" class="PLLe" style="width:100px;" name="c_pedido" id="c_pedido" onblur="if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_vendedor.focus(); filtra_pedido();">
		</td></tr>

<!--  VENDEDOR  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe">VENDEDOR</span>
		<br>
			<select id="c_vendedor" name="c_vendedor" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =vendedores_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>

<!--  INDICADOR  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe">INDICADOR</span>
		<br>
			<select id="c_indicador" name="c_indicador" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =indicadores_monta_itens_select(Null) %>
			</select>
			</td></tr>

<!-- ************   LOJAS   ************ -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">LOJA(S)</span>
		<br>
			<textarea class="PLBe" style="font-size:9pt;width:110px;margin-left:33px;margin-bottom:4px;" rows="6" name="c_lista_loja" id="c_lista_loja" onkeypress="if (!digitou_enter(false) && !digitou_char('-')) filtra_numerico();" onblur="this.value=normaliza_lista_lojas(this.value);"></textarea>
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
