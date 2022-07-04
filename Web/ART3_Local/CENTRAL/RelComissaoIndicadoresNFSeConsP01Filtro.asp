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
'	  RelComissaoIndicadoresNFSeConsP01Filtro.asp
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

	dim intIdx
	dim c_competencia_mes, c_competencia_ano, c_dt_proc_comissao_inicio, c_dt_proc_comissao_termino, c_vendedor, c_cnpj_nfse, s_cnpj_nfse, c_numero_nfse, rb_proc_fluxo_caixa
	c_competencia_mes = Trim(Request("c_competencia_mes"))
	c_competencia_ano = Trim(Request("c_competencia_ano"))
	c_dt_proc_comissao_inicio = Trim(Request("c_dt_proc_comissao_inicio"))
	c_dt_proc_comissao_termino = Trim(Request("c_dt_proc_comissao_termino"))
	c_vendedor = Trim(Request("c_vendedor"))
	c_cnpj_nfse = Trim(Request("c_cnpj_nfse"))
	s_cnpj_nfse = ""
	if c_cnpj_nfse <> "" then s_cnpj_nfse = cnpj_cpf_formata(c_cnpj_nfse)
	c_numero_nfse = Trim(Request("c_numero_nfse"))
	rb_proc_fluxo_caixa = Trim(Request("rb_proc_fluxo_caixa"))

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function mes_monta_itens_select(byval mes_default)
dim i, x, s_selected
	mes_default = Trim("" & mes_default)
	x = "<option value=''>&nbsp;</option>"
	for i=1 to 12
		if CStr(i) = mes_default then
			s_selected = " selected"
		else
			s_selected = ""
			end if
		x = x & "<option value='" & i & "'" & s_selected & ">" & i & "</option>"
	next

	mes_monta_itens_select = x
end function


function ano_monta_itens_select(byval ano_default)
dim i, x, ano_atual, s_selected
	ano_default = Trim("" & ano_default)
	ano_atual = Year(Date)
	x = "<option value=''>&nbsp;</option>"
	for i=ano_atual to (ano_atual - 10) step -1
		if CStr(i) = ano_default then
			s_selected = " selected"
		else
			s_selected = ""
			end if
		x = x & "<option value='" & i & "'" & s_selected & ">" & i & "</option>"
	next

	ano_monta_itens_select = x
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

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function () {
		$("#c_dt_proc_comissao_inicio").hUtilUI('datepicker_peq_filtro_inicial');
		$("#c_dt_proc_comissao_termino").hUtilUI('datepicker_peq_filtro_final');
	});
</script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma(f) {
var s_de, s_ate, blnHaFiltroPeriodo;

	blnHaFiltroPeriodo = false;

	if ((trim($("#c_competencia_mes").val()) != "") || (trim($("#c_competencia_ano").val()) != "")) {
		if (trim($("#c_competencia_mes").val()) == "") {
			alert("Mês de competência não foi informado corretamente!");
			$("#c_competencia_mes").focus();
			return;
		}

		if (trim($("#c_competencia_ano").val()) == "") {
			alert("Mês de competência não foi informado corretamente!");
			$("#c_competencia_ano").focus();
			return;
		}

		blnHaFiltroPeriodo = true;
	}

	if ((trim(f.c_dt_proc_comissao_inicio.value) != "") || (trim(f.c_dt_proc_comissao_termino.value) != "")) {
		if (trim(f.c_dt_proc_comissao_inicio.value) == "") {
			alert("Informe a data de início do período!");
			f.c_dt_proc_comissao_inicio.focus();
			return;
		}

		if (trim(f.c_dt_proc_comissao_termino.value) == "") {
			alert("Informe a data de término do período!");
			f.c_dt_proc_comissao_termino.focus();
			return;
		}

		if (trim(f.c_dt_proc_comissao_inicio.value) != "") {
			if (!isDate(f.c_dt_proc_comissao_inicio)) {
				alert("Data de início inválida!");
				f.c_dt_proc_comissao_inicio.focus();
				return;
			}
		}

		if (trim(f.c_dt_proc_comissao_termino.value) != "") {
			if (!isDate(f.c_dt_proc_comissao_termino)) {
				alert("Data de término inválida!");
				f.c_dt_proc_comissao_termino.focus();
				return;
			}
		}

		s_de = trim(f.c_dt_proc_comissao_inicio.value);
		s_ate = trim(f.c_dt_proc_comissao_termino.value);
		if ((s_de != "") && (s_ate != "")) {
			s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
			s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
			if (s_de > s_ate) {
				alert("Data de término é menor que a data de início!");
				f.c_dt_proc_comissao_termino.focus();
				return;
			}
		}

		blnHaFiltroPeriodo = true;
	}

	if (!blnHaFiltroPeriodo) {
		alert("É necessário informar um dos filtros de período:\nMês de competência\nou\nData processamento da comissão");
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<body>
<center>
<form id="fFILTRO" name="fFILTRO" method="post" action="RelComissaoIndicadoresNFSeConsP02Exec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (via NFS-e) (Consulta)</span>
	<br /><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br />

<table class="Qx" cellSpacing="0" style="width:600px;">
<!--  PERÍODO: MÊS DE COMPETÊNCIA DO RELATÓRIO  -->
	<tr>
		<td class="ME MD MC" nowrap><span class="PLTe">MÊS DE COMPETÊNCIA</span></td>
	</tr>
	<tr bgColor="#FFFFFF" nowrap>
		<td class="ME MD MB">
			<table cellSpacing="0" cellPadding="0" style="margin:0px 20px 6px 10px;">
			<tr bgColor="#FFFFFF">
				<td>
					<select class="Cc" style="font-size:10pt;width:50px;" name="c_competencia_mes" id="c_competencia_mes" />
						<%=mes_monta_itens_select(c_competencia_mes)%>
					</select>
					<span class="C" style="font-size:12pt;">/</span>
					<select class="Cc" style="font-size:10pt;width:80px;" name="c_competencia_ano" id="c_competencia_ano" />
						<%=ano_monta_itens_select(c_competencia_ano)%>
					</select>
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  PERÍODO: DATA DE PROCESSAMENTO DA COMISSÃO  -->
	<tr>
		<td class="ME MD" nowrap><span class="PLTe">DATA PROCESSAMENTO DA COMISSÃO</span></td>
	</tr>
	<tr bgColor="#FFFFFF" nowrap>
		<td class="ME MD MB">
			<table cellSpacing="0" cellPadding="0" style="margin:0px 20px 6px 10px;">
			<tr bgColor="#FFFFFF">
				<td>
					<input class="Cc" maxlength="10" style="font-size:9pt;width:80px;" name="c_dt_proc_comissao_inicio" id="c_dt_proc_comissao_inicio" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.select(); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_proc_comissao_termino.focus(); filtra_data();"
						value="<%=c_dt_proc_comissao_inicio%>" />&nbsp;
					<span class="C">&nbsp;até&nbsp;</span>&nbsp;
					<input class="Cc" maxlength="10" style="font-size:9pt;width:80px;" name="c_dt_proc_comissao_termino" id="c_dt_proc_comissao_termino" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.select(); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_data();"
						value="<%=c_dt_proc_comissao_termino%>" />
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  VENDEDOR  -->
	<tr>
		<td class="ME MD" nowrap valign="bottom"><span class="PLTe">VENDEDOR</span></td>
	</tr>
	<tr bgColor="#FFFFFF" nowrap>
		<td class="ME MB MD">
			<input maxlength="10" class="PLLe" style="font-size:9pt;width:220px;" name="c_vendedor" id="c_vendedor" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_indicador.focus(); filtra_nome_identificador();" value="<%=c_vendedor%>" />
		</td>
	</tr>

<!--  CNPJ NFS-e  -->
	<tr>
		<td class="ME MD" nowrap valign="bottom"><span class="PLTe">CNPJ NFS-e</span></td>
	</tr>
	<tr bgColor="#FFFFFF" nowrap>
		<td class="ME MB MD">
			<input name="c_cnpj_nfse" id="c_cnpj_nfse" maxlength="18" class="PLLe" style="font-size:9pt;width:220px;" onblur="if (cnpj_cpf_ok(this.value)) {this.value=cnpj_cpf_formata(this.value);}" onkeypress="if (digitou_enter(true)&&((!tem_info(this.value))||(tem_info(this.value)&&cnpj_cpf_ok(this.value)))) {this.value=cnpj_cpf_formata(this.value); bCONFIRMA.focus();} filtra_cnpj_cpf();" value="<%=s_cnpj_nfse%>" />
		</td>
	</tr>

<!--  Nº NFS-e  -->
	<tr>
		<td class="ME MD" nowrap valign="bottom"><span class="PLTe">Nº NFS-e</span></td>
	</tr>
	<tr bgColor="#FFFFFF" nowrap>
		<td class="ME MB MD">
			<input name="c_numero_nfse" id="c_numero_nfse" maxlength="9" class="PLLe" style="font-size:9pt;width:220px;" onblur="this.value=trim(this.value);" onkeypress="filtra_numerico();" value="<%=c_numero_nfse%>" />
		</td>
	</tr>

<!--  Status de processamento dos lançamentos no fluxo de caixa  -->
	<tr>
		<td class="ME MD" nowrap valign="bottom"><span class="PLTe">STATUS PROCESSAMENTO FLUXO CAIXA</span></td>
	</tr>
	<tr bgColor="#FFFFFF" nowrap>
		<td class="ME MB MD">
			<% intIdx=-1 %>
			<input type="radio" id="rb_proc_fluxo_caixa" name="rb_proc_fluxo_caixa" value="1" class="CBOX" style="margin-left:20px;"
				<% if (rb_proc_fluxo_caixa = "1") then Response.Write " checked" %>
				/>
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_proc_fluxo_caixa[<%=Cstr(intIdx)%>].click();">Somente Já Processados</span>
			<br />
			<input type="radio" id="rb_proc_fluxo_caixa" name="rb_proc_fluxo_caixa" value="0" class="CBOX" style="margin-left:20px;"
				<% if (rb_proc_fluxo_caixa = "0") then Response.Write " checked" %>
				/>
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_proc_fluxo_caixa[<%=Cstr(intIdx)%>].click();">Somente Não Processados</span>
			<br />
			<input type="radio" id="rb_proc_fluxo_caixa" name="rb_proc_fluxo_caixa" value="" class="CBOX" style="margin-left:20px;"
				<% if (rb_proc_fluxo_caixa = "") then Response.Write " checked" %>
				/>
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_proc_fluxo_caixa[<%=Cstr(intIdx)%>].click();">Ambos</span>
		</td>
	</tr>

</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br />


<table width="649" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="resumo.asp?<%=MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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
