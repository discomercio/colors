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
'	  RelTransacoesCieloFiltro.asp
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

	if Not operacao_permitida(OP_CEN_REL_TRANSACOES_CIELO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
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
	listaOpcoes = Array(COD_REL_TRANSACOES_CIELO__TRANSACAO_AUTORIZADA, _
						COD_REL_TRANSACOES_CIELO__TRANSACAO_NAO_AUTORIZADA, _
						COD_REL_TRANSACOES_CIELO__TRANSACAO_CANCELADA_PELO_USUARIO, _
						COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_SITUACAO_DESCONHECIDA)
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



<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

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


<body onload="fFILTRO.c_dt_inicio.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelTransacoesCieloExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Transações Cielo</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0" style="width:240px;">
<!--  PERÍODO: DATA INICIAL E FINAL  -->
	<tr>
		<td class="ME MD MC" NOWRAP><span class="PLTe">PERÍODO DA CONSULTA</span></td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MD MB">
			<table cellSpacing="0" cellPadding="0" style="margin:0px 20px 6px 10px;">
			<tr bgColor="#FFFFFF">
				<td>
					<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_inicio" id="c_dt_inicio" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_termino.focus(); filtra_data();"
						>&nbsp;
					<span class="C">&nbsp;até&nbsp;</span>&nbsp;
					<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_termino" id="c_dt_termino" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_pedido.focus(); filtra_data();">
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  RESULTADO DA TRANSAÇÃO  -->
	<tr>
		<td class="ME MD PLTe" NOWRAP valign="bottom">&nbsp;RESULTADO DA TRANSAÇÃO</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<select id="c_resultado_transacao" name="c_resultado_transacao" style="width:400px;margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=rel_transacoes_cielo_status_transacao_monta_itens_select("")%>
			</select>
		</td>
	</tr>

<!--  BANDEIRA  -->
	<tr>
		<td class="ME MD PLTe" NOWRAP valign="bottom">&nbsp;BANDEIRA</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<select id="c_bandeira" name="c_bandeira" style="width:200px;margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=rel_transacoes_cielo_bandeira_monta_itens_select("")%>
			</select>
		</td>
	</tr>

<!--  Nº PEDIDO  -->
	<tr>
		<td class="ME MD" NOWRAP><span class="PLTe">PEDIDO</span></td>
	</tr>
	<tr>
		<td class="ME MD MB">
			<table cellSpacing="0" cellPadding="0" style="margin-bottom:2px;">
			<tr bgColor="#FFFFFF">
				<td align="right" style="width:80px;">
					<span class="C">Nº Pedido</span>
				</td>
				<td>
					<input class="Cc" maxlength="10" style="width:80px;" name="c_pedido" id="c_pedido" onblur="if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value);" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value); fFILTRO.c_cliente_cnpj_cpf.focus();} filtra_pedido();">
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  CLIENTE  -->
	<tr>
		<td class="ME MD" NOWRAP><span class="PLTe">CLIENTE</span></td>
	</tr>
	<tr>
		<td class="ME MD MB">
			<table cellSpacing="0" cellPadding="0" style="margin-bottom:2px;">
			<tr bgColor="#FFFFFF">
				<td align="right" style="width:80px;">
					<span class="C">CNPJ/CPF</span>
				</td>
				<td>
					<input class="Cc" maxlength="18" style="width:150px;" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="if (digitou_enter(true)&&((!tem_info(this.value))||(tem_info(this.value)&&cnpj_cpf_ok(this.value)))) {this.value=cnpj_cpf_formata(this.value); fFILTRO.c_loja.focus();} filtra_cnpj_cpf();">
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  LOJA  -->
	<tr>
		<td class="ME MD" NOWRAP><span class="PLTe">LOJA</span></td>
	</tr>
	<tr>
		<td class="ME MD MB">
			<table cellSpacing="0" cellPadding="0" style="margin-bottom:2px;">
			<tr bgColor="#FFFFFF">
				<td align="right" style="width:80px;">
					<span class="C">Nº Loja</span>
				</td>
				<td>
					<input class="Cc" maxlength="3" style="width:60px;" name="c_loja" id="c_loja" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); filtra_numerico();">
				</td>
			</tr>
			</table>
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
