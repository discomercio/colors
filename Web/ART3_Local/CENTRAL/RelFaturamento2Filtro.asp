<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  RelFaturamento2Filtro.asp
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


    Const SAIDA_FABRICANTE = "FABRICANTE"
    Const SAIDA_LOJA = "LOJA" 



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________________________________________________
' INDICADORES MONTA ITENS SELECT
' LEMBRE-SE: O ORÇAMENTISTA É CONSIDERADO AUTOMATICAMENTE UM INDICADOR!!
function indicadores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT apelido, razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR ORDER BY apelido")
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("apelido")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("razao_social_nome_iniciais_em_maiusculas"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	indicadores_monta_itens_select = strResp
	r.close
	set r=nothing
end function

'----------------------------------------------------------------------------------------------
' grupo_origem_pedido_monta_itens_select
function grupo_origem_pedido_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE grupo='PedidoECommerce_Origem_Grupo' AND st_inativo=0")
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
	
    ' strResp = "<option value=''>&nbsp;</option>" & strResp

	grupo_origem_pedido_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' __________________________________________________
' origem_pedido_monta_itens_select
'
function origem_pedido_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE grupo='PedidoECommerce_Origem' AND st_inativo=0")
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
	
    ' strResp = "<option value=''>&nbsp;</option>" & strResp

	origem_pedido_monta_itens_select = strResp
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
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
		$("#c_dt_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_termino").hUtilUI('datepicker_filtro_final');
	});
</script>
<script type="text/javascript">
    function limpaCampoSelectGrupoOrigemPedido() {
        $("#c_grupo_pedido_origem").children().prop('selected', false);
    }
    function limpaCampoSelectOrigemPedido() {
        $("#c_pedido_origem").children().prop('selected', false);
    }
</script>

<script language="JavaScript" type="text/javascript">
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

	if (trim(f.c_dt_termino.value)!="") {
		if (!isDate(f.c_dt_termino)) {
			alert("Data de término inválida!!");
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

	if ((trim(f.c_produto.value)!="")&&(trim(f.c_fabricante.value)=="")) {
		if (!isEAN(f.c_produto.value)) {
			alert("Preencha o código do fabricante do produto " + f.c_produto.value + "!!");
			f.c_fabricante.focus();
			return;
			}
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">


<body onload="fFILTRO.c_dt_inicio.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelFaturamento2Exec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Faturamento II</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PERÍODO  -->
<table class="Qx" cellspacing="0">
	<tr bgcolor="#FFFFFF">
	<td class="MT" align="left" nowrap><span class="PLTe">PERÍODO</span>
	<br>
		<table cellspacing="0" cellpadding="0"><tr bgColor="#FFFFFF"><td align="left">
		<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_inicio" id="c_dt_inicio" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_termino.focus(); filtra_data();"
			>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;até&nbsp;</span>&nbsp;<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_termino" id="c_dt_termino" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); filtra_data();">
			</td></tr>
		</table>
		</td></tr>

<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">FABRICANTE</span>
	<br>
		<select id="c_fabricante" name="c_fabricante" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
		<%=fabricante_monta_itens_select(Null) %>
		</select>
		</td></tr>

<!--  PRODUTO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">PRODUTO</span>
	<br>
		<input maxlength="13" class="PLLe" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) fFILTRO.c_grupo.focus(); filtra_produto();">
		</td></tr>

<!--  GRUPO DE PRODUTOS  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">GRUPO DE PRODUTOS</span>
	<br>
		<input maxlength="2" class="PLLe" style="width:60px;" name="c_grupo" id="c_grupo" onkeypress="if (digitou_enter(true)) fFILTRO.c_vendedor.focus();" onblur="this.value=ucase(this.value);">
		</td></tr>

<!--  VENDEDOR  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">VENDEDOR</span>
	<br>
		<input maxlength="10" class="PLLe" style="width:100px;" name="c_vendedor" id="c_vendedor" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_pedido.focus(); filtra_nome_identificador();">
		</td></tr>
		
<!--  INDICADOR  -->
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">INDICADOR</span>
		<br>
			<select id="c_indicador" name="c_indicador" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =indicadores_monta_itens_select(Null) %>
			</select>
			</td></tr>
		
<!--  PEDIDO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">PEDIDO</span>
	<br>
		<input maxlength="10" class="PLLe" style="width:100px;" name="c_pedido" id="c_pedido" onblur="if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.op_forma_pagto.focus(); filtra_pedido();">
		</td></tr>

<!-- ORIGEM DO PEDIDO (GRUPO) -->
    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">ORIGEM DO PEDIDO (GRUPO)</span>
		<br>
		<table cellspacing="0" cellpadding="0"><tr bgcolor="#FFFFFF"><td align="left">            
			<select id="c_grupo_pedido_origem" name="c_grupo_pedido_origem" style="margin:1px 3px 6px 10px;width: 200px" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="5" multiple>
			<% =grupo_origem_pedido_monta_itens_select(Null) %>
			</select>
			</td>
        <td align="left" valign="top">
			<a href="javascript:limpaCampoSelectGrupoOrigemPedido()" title="limpa o filtro 'Origem do Pedido (Grupo)'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td></tr></table>
        </td></tr>

<!-- ORIGEM DO PEDIDO -->
    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">ORIGEM DO PEDIDO</span>
		<br>
            <table cellspacing="0" cellpadding="0"><tr bgcolor="#FFFFFF"><td align="left">  
			<select id="c_pedido_origem" name="c_pedido_origem" style="margin:1px 3px 6px 10px;width: 200px" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="5" multiple>
			<% =origem_pedido_monta_itens_select(Null) %>
			</select>
			</td>
            <td align="left" valign="top">
			<a href="javascript:limpaCampoSelectOrigemPedido()" title="limpa o filtro 'Origem do Pedido'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td></tr></table>
        </td></tr>
		

<!--  FORMA DE PAGAMENTO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">FORMA DE PAGAMENTO</span>
		<br>
			<span class="C" style="margin-left:30px;">Forma de Pagamento</span>
				<select id="op_forma_pagto" name="op_forma_pagto" 
					onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;"
					onkeypress="if (digitou_enter(true)) fFILTRO.c_forma_pagto_qtde_parc.focus();"
					>
				  <% =forma_pagto_monta_itens_select(Null) %>
				</select>
		<br>
			<span class="C" style="margin-left:30px;">Nº Parcelas</span>
				<input class="Cc" maxlength="2" style="width:40px;" name="c_forma_pagto_qtde_parc" id="c_forma_pagto_qtde_parc" onkeypress="if (digitou_enter(true)) fFILTRO.rb_tipo_cliente[0].focus(); filtra_numerico();">
	</td></tr>

<!--  TIPO DE CLIENTE  -->
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">TIPO DE CLIENTE</span>
		<br>
		<% intIdx=-1 %>
		<input type="radio" id="rb_tipo_cliente" name="rb_tipo_cliente" value=<%=ID_PF%> style="margin-left:30px;">
		<% intIdx=intIdx+1 %>
		<span style="cursor:default" class="Np" onclick="fFILTRO.rb_tipo_cliente[<%=Cstr(intIdx)%>].click();">Pessoa Física</span>
		<br />
		<input type="radio" id="rb_tipo_cliente" name="rb_tipo_cliente" value=<%=ID_PJ%> style="margin-left:30px;">
		<% intIdx=intIdx+1 %>
		<span style="cursor:default" class="Np" onclick="fFILTRO.rb_tipo_cliente[<%=Cstr(intIdx)%>].click();">Pessoa Jurídica</span>
		<br />
		<input type="radio" id="rb_tipo_cliente" name="rb_tipo_cliente" value="" style="margin-left:30px;" checked>
		<% intIdx=intIdx+1 %>
		<span style="cursor:default" class="Np" onclick="fFILTRO.rb_tipo_cliente[<%=Cstr(intIdx)%>].click();">Ambos</span>
	</td></tr>

<!--  UF DO CLIENTE  -->
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">UF DO CLIENTE</span>
		<br>
		  <select id="c_uf_pesq" name="c_uf_pesq" style="margin-left:30px; margin-bottom:6px;"
				onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;"
				onkeypress="if (digitou_enter(true)) fFILTRO.c_loja.focus();">
		<% =UF_monta_itens_select(Null) %>
		  </select>
		</td>
	</tr>

<!--  EMPRESA  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe">EMPRESA</span>
		<br>
			<select id="c_empresa" name="c_empresa" style="margin:1px 10px 6px 30px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			    <%=apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>

<!--  LOJA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">LOJA(S)</span>
	<br>
		<textarea class="PLBe" style="width:100px;font-size:9pt;margin-left:30px;margin-bottom:4px;" rows="8" name="c_loja" id="c_loja" onkeypress="if (!digitou_enter(false) && !digitou_char('-')) filtra_numerico();" onblur="this.value=normaliza_lista_lojas(this.value);"></textarea>
	</td></tr>

<!--  TIPO DE AGRUPAMENTO  -->
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">AGRUPAMENTO</span>
		<br>
		<% intIdx=-1 %>
		<input type="radio" id="rb_saida" name="rb_saida" value=<%=SAIDA_FABRICANTE%> style="margin-left:30px;" checked>
		<% intIdx=intIdx+1 %>
		<span style="cursor:default" class="Np" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">Fabricante</span>
		<br />
		<input type="radio" id="rb_saida" name="rb_saida" value=<%=SAIDA_LOJA%> style="margin-left:30px;">
		<% intIdx=intIdx+1 %>
		<span style="cursor:default" class="Np" onclick="fFILTRO.rb_saida[<%=Cstr(intIdx)%>].click();">Loja</span>
	</td></tr>

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
