<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  RelResumoOperacoesEntreEstoques.asp
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
function marcar_todos() {
	fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_VENDA.checked = true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA.checked = true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSFERENCIA.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_ENTREGA.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_DEVOLUCAO.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_ESTORNO.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_SPLIT.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA.checked=true;
	fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS.checked=true;
}

function desmarcar_todos() {
	fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_VENDA.checked = false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA.checked = false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSFERENCIA.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_ENTREGA.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_DEVOLUCAO.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_ESTORNO.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_SPLIT.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA.checked=false;
	fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS.checked=false;
}

function quantidade_operacoes_assinaladas( ){
var n;
	n=0;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_VENDA.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSFERENCIA.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_ENTREGA.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_DEVOLUCAO.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_ESTORNO.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_SPLIT.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA.checked) n++;
	if (fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS.checked) n++;
	return n;
}

function fFILTROConfirma( f ) {
var s_de, s_ate;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;

	if (trim(f.c_dt_inicio.value)!="") {
		if (!isDate(f.c_dt_inicio)) {
			alert("Data de início inválida!!");
			f.c_dt_inicio.focus();
			return;
			}
		}
	else {
		alert("Data de início inválida!!");
		f.c_dt_inicio.focus();
		return;
		}

	if (trim(f.c_dt_termino.value)!="") {
		if (!isDate(f.c_dt_termino)) {
			alert("Data de término inválida!!");
			f.c_dt_termino.focus();
			return;
			}
		}
	else {
		alert("Data de término inválida!!");
		f.c_dt_termino.focus();
		return;
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
		
	if ((trim(f.c_produto.value)!="")&&(trim(f.c_fabricante.value)=="")) {
		if (!isEAN(f.c_produto.value)) {
			alert("Preencha o código do fabricante do produto " + f.c_produto.value + "!!");
			f.c_fabricante.focus();
			return;
			}
		}
		
	if (quantidade_operacoes_assinaladas()==0) {
		alert("Nenhuma operação foi selecionada!!");
		return;
		}

//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
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

<form id="fFILTRO" name="fFILTRO" method="post" action="RelResumoOperacoesEntreEstoquesExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Resumo de Operações Entre Estoques</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PERÍODO  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MT" NOWRAP><span class="PLTe">PERÍODO</span>
	<br>
		<table cellSpacing="0" cellPadding="0"><tr bgColor="#FFFFFF"><td>
		<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_inicio" id="c_dt_inicio" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_termino.focus(); filtra_data();"
			>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;até&nbsp;</span>&nbsp;<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_termino" id="c_dt_termino" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); filtra_data();">
			</td></tr>
		</table>
		</td></tr>

<!--  EMPRESA (CD)  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">EMPRESA (CD)</span>
	<br>
		<select id="c_id_nfe_emitente" name="c_id_nfe_emitente" style="margin:4px 4px 4px 4px;">
		<% =apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
		</select>
		</td></tr>

<!--  FABRICANTE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">FABRICANTE</span>
	<br>
		<input maxlength="4" class="PLLe" style="width:50px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto.focus(); filtra_fabricante();">
		</td></tr>

<!--  PRODUTO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">PRODUTO</span>
	<br>
		<input maxlength="13" class="PLLe" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_produto();">
		</td></tr>

<!-- ************   LOJAS   ************ -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">LOJA(S)</span>
		<br><center>
			<textarea class="PLBe" style="font-size:9pt;width:110px;margin-bottom:4px;" rows="6" name="c_lista_loja" id="c_lista_loja" onkeypress="if (!digitou_enter(false) && !digitou_char('-')) filtra_numerico();" onblur="this.value=normaliza_lista_lojas(this.value);"></textarea>
		</center>
	</td></tr>

<!-- ************   OPERAÇÕES DE MOVIMENTAÇÃO NO ESTOQUE   ************ -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">OPERAÇÕES</span>
		<br>
		<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_ENTRADA" name="ckb_OP_ESTOQUE_LOG_ENTRADA"
				value="<%=OP_ESTOQUE_LOG_ENTRADA%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_ENTRADA)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_VENDA" name="ckb_OP_ESTOQUE_LOG_VENDA"
				value="<%=OP_ESTOQUE_LOG_VENDA%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_VENDA.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_VENDA)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA" name="ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA"
				value="<%=OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_TRANSFERENCIA" name="ckb_OP_ESTOQUE_LOG_TRANSFERENCIA"
				value="<%=OP_ESTOQUE_LOG_TRANSFERENCIA%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSFERENCIA.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_TRANSFERENCIA)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_DEVOLUCAO" name="ckb_OP_ESTOQUE_LOG_DEVOLUCAO"
				value="<%=OP_ESTOQUE_LOG_DEVOLUCAO%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_DEVOLUCAO.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_DEVOLUCAO)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS" name="ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS"
				value="<%=OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_SPLIT" name="ckb_OP_ESTOQUE_LOG_SPLIT"
				value="<%=OP_ESTOQUE_LOG_SPLIT%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_SPLIT.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_SPLIT)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA" name="ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA"
				value="<%=OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_ESTORNO" name="ckb_OP_ESTOQUE_LOG_ESTORNO"
				value="<%=OP_ESTOQUE_LOG_ESTORNO%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_ESTORNO.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_ESTORNO)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA" name="ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA"
				value="<%=OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_ENTREGA" name="ckb_OP_ESTOQUE_LOG_ENTREGA"
				value="<%=OP_ESTOQUE_LOG_ENTREGA%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_ENTREGA.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_ENTREGA)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA" name="ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA"
				value="ON"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_TRANSFERENCIA_ROUBO_PERDA.click();">Roubo ou Perda Total</span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT" name="ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT"
				value="<%=OP_ESTOQUE_LOG_CONVERSAO_KIT%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_CONVERSAO_KIT.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_CONVERSAO_KIT)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE" name="ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE"
				value="<%=OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM" name="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM"
				value="<%=OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA" name="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA"
				value="<%=OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA)%></span>
			</td></tr>
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA" name="ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA"
				value="<%=OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA%>"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA.click();"><%=x_operacao_log_estoque(OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA)%></span>
			</td></tr>
		</table>
		<input name="bMarcarTodos" id="bMarcarTodos" type="button" class="Button" onclick="marcar_todos();" value="Marcar todos" title="assinala todas as operações" style="margin-left:6px;margin-bottom:10px">
		<input name="bDesmarcarTodos" id="bDesmarcarTodos" type="button" class="Button" onclick="desmarcar_todos();" value="Desmarcar todos" title="desmarca todas as operações" style="margin-left:6px;margin-right:6px;margin-bottom:10px">
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
	<td align="RIGHT"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="executa a consulta">
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
