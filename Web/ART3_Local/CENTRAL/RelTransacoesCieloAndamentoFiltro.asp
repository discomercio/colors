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
'	  RelTransacoesCieloAndamentoFiltro.asp
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

	if Not operacao_permitida(OP_CEN_REL_TRANSACOES_CIELO_ANDAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim intIdx
	dim url_back, strUrlBotaoVoltar
	url_back = Trim(Request("url_back"))
	if url_back <> "" then
		strUrlBotaoVoltar = "Resumo.asp?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	else
		strUrlBotaoVoltar = "javascript:history.back()"
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S
' _____________________________________________________________________________________________

' __________________________________________________________
' REL_TRANSACOES_CIELO_STATUS_TRANSACAO_MONTA_ITENS_SELECT
'
function rel_transacoes_cielo_status_transacao_monta_itens_select(byval id_default)
dim ha_default
dim listaOpcoes
dim i, x, strResp
	listaOpcoes = Array(COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_ANDAMENTO, _
						COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_AUTENTICACAO)
	id_default = Trim("" & id_default)
	ha_default=False
	strResp = ""
	for i = Lbound(listaOpcoes) to Ubound(listaOpcoes)
		x = Trim("" & listaOpcoes(i))
		if x <> "" then
			if (id_default <> "") And (id_default = x) then
				strResp = strResp & "<option selected"
				ha_default = True
			else
				strResp = strResp & "<option"
				end if
			strResp = strResp & " value='" & x & "'>" & _
								descricao_cod_rel_transacoes_cielo(x) & _
								"</option>" & chr(13)
			end if
		next
	
	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
	
	rel_transacoes_cielo_status_transacao_monta_itens_select = strResp
end function



' __________________________________________________________
' REL_TRANSACOES_CIELO_BANDEIRA_MONTA_ITENS_SELECT
'
function rel_transacoes_cielo_bandeira_monta_itens_select(byval id_default)
dim ha_default
dim listaOpcoes
dim i, x, strResp
	listaOpcoes = CieloArrayBandeiras
	id_default = Trim("" & id_default)
	ha_default=False
	strResp = ""
	for i = Lbound(listaOpcoes) to Ubound(listaOpcoes)
		x = Trim("" & listaOpcoes(i))
		if x <> "" then
			if (id_default <> "") And (id_default = x) then
				strResp = strResp & "<option selected"
				ha_default = True
			else
				strResp = strResp & "<option"
				end if
			strResp = strResp & " value='" & x & "'>" & _
								CieloDescricaoBandeira(x) & _
								"</option>" & chr(13)
			end if
		next
	
	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
	
	rel_transacoes_cielo_bandeira_monta_itens_select = strResp
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
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
		$("#c_dt_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_termino").hUtilUI('datepicker_filtro_final');
	});
</script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma(f) {
var s_de, s_ate;

	if (trim(f.c_dt_inicio.value) == "") {
		alert("Informe a data de início do período!!");
		f.c_dt_inicio.focus();
		return;
	}

	if (trim(f.c_dt_termino.value) == "") {
		alert("Informe a data de término do período!!");
		f.c_dt_termino.focus();
		return;
	}

	if (trim(f.c_dt_inicio.value) != "") {
		if (!isDate(f.c_dt_inicio)) {
			alert("Data de início inválida!!");
			f.c_dt_inicio.focus();
			return;
		}
	}

	if (trim(f.c_dt_termino.value) != "") {
		if (!isDate(f.c_dt_termino)) {
			alert("Data de término inválida!!");
			f.c_dt_termino.focus();
			return;
		}
	}

	s_de = trim(f.c_dt_inicio.value);
	s_ate = trim(f.c_dt_termino.value);
	if ((s_de != "") && (s_ate != "")) {
		s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
		s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
		if (s_de > s_ate) {
			alert("Data de término é menor que a data de início!!");
			f.c_dt_termino.focus();
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">


<body onload="fFILTRO.c_dt_inicio.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelTransacoesCieloAndamentoExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Transações Cielo em Andamento</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellspacing="0" style="width:240px;">
<!--  PERÍODO: DATA INICIAL  -->
	<tr>
		<td class="ME MD MC" align="left" nowrap><span class="PLTe">PERÍODO DA CONSULTA</span></td>
	</tr>
	<tr bgcolor="#FFFFFF" nowrap>
		<td class="ME MD MB" align="left">
			<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 10px;">
			<tr bgcolor="#FFFFFF">
				<td align="left">
					<input class="Cc" maxlength="10" style="width:90px;" name="c_dt_inicio" id="c_dt_inicio" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_termino.focus(); filtra_data();"
						>&nbsp;
					<span class="C">&nbsp;até&nbsp;</span>&nbsp;
					<input class="Cc" maxlength="10" style="width:90px;" name="c_dt_termino" id="c_dt_termino" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_data();">
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  RESULTADO DA TRANSAÇÃO  -->
	<tr>
		<td class="ME MD PLTe" align="left" valign="bottom" nowrap>&nbsp;RESULTADO DA TRANSAÇÃO</td>
	</tr>
	<tr bgcolor="#FFFFFF" nowrap>
		<td class="ME MB MD" align="left">
			<select id="c_resultado_transacao" name="c_resultado_transacao" style="width:400px;margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=rel_transacoes_cielo_status_transacao_monta_itens_select("")%>
			</select>
		</td>
	</tr>

<!--  BANDEIRA  -->
	<tr>
		<td class="ME MD PLTe" align="left" valign="bottom" nowrap>&nbsp;BANDEIRA</td>
	</tr>
	<tr bgcolor="#FFFFFF" nowrap>
		<td class="ME MB MD" align="left">
			<select id="c_bandeira" name="c_bandeira" style="width:200px;margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=rel_transacoes_cielo_bandeira_monta_itens_select("")%>
			</select>
		</td>
	</tr>

<!--  Nº PEDIDO  -->
	<tr>
		<td class="ME MD" align="left" nowrap><span class="PLTe">PEDIDO</span></td>
	</tr>
	<tr>
		<td class="ME MD MB" align="left">
			<table cellspacing="0" cellpadding="0" style="margin-bottom:2px;">
			<tr bgcolor="#FFFFFF">
				<td align="right" style="width:80px;">
					<span class="C">Nº Pedido</span>
				</td>
				<td align="left">
					<input class="Cc" maxlength="10" style="width:80px;" name="c_pedido" id="c_pedido" onblur="if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value);" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value); bCONFIRMA.focus();} filtra_pedido();">
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  CLIENTE  -->
	<tr>
		<td class="ME MD" align="left" nowrap><span class="PLTe">CLIENTE</span></td>
	</tr>
	<tr>
		<td class="ME MD MB" align="left">
			<table cellspacing="0" cellpadding="0" style="margin-bottom:2px;">
			<tr bgcolor="#FFFFFF">
				<td align="right" style="width:80px;">
					<span class="C">CNPJ/CPF</span>
				</td>
				<td align="left">
					<input class="Cc" maxlength="18" style="width:150px;" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="if (digitou_enter(true)&&((!tem_info(this.value))||(tem_info(this.value)&&cnpj_cpf_ok(this.value)))) {this.value=cnpj_cpf_formata(this.value); bCONFIRMA.focus();} filtra_cnpj_cpf();">
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  LOJA  -->
	<tr>
		<td class="ME MD" align="left" nowrap><span class="PLTe">LOJA</span></td>
	</tr>
	<tr>
		<td class="ME MD MB" align="left">
			<table cellspacing="0" cellpadding="0" style="margin-bottom:2px;">
			<tr bgcolor="#FFFFFF">
				<td align="right" style="width:80px;">
					<span class="C">Nº Loja</span>
				</td>
				<td align="left">
					<input class="Cc" maxlength="3" style="width:60px;" name="c_loja" id="c_loja" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); filtra_numerico();">
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  ORDENAÇÃO  -->
	<tr>
		<td class="ME MD" align="left" nowrap><span class="PLTe">ORDENAÇÃO DO RESULTADO</span></td>
	</tr>
	<tr>
		<td class="ME MD MB" align="left">
			<% intIdx=-1 %>
			<input type="radio" id="rb_ordenacao_saida" name="rb_ordenacao_saida" value="ORD_POR_PEDIDO" class="CBOX" style="margin-left:20px;" checked>
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_ordenacao_saida[<%=Cstr(intIdx)%>].click();">Pedido</span>
			<br />
			<input type="radio" id="rb_ordenacao_saida" name="rb_ordenacao_saida" value="ORD_POR_DATA" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_ordenacao_saida[<%=Cstr(intIdx)%>].click();">Data</span>
		</td>
	</tr>

</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
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
