<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  '	  R E L I M P O S T O S P A G O S F I L T R O . A S P
'     =============================================================
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

'	CD
	dim i, qtde_nfe_emitente
	dim v_usuario_x_nfe_emitente
	dim id_nfe_emitente_selecionado
	v_usuario_x_nfe_emitente = obtem_lista_usuario_x_nfe_emitente(usuario)
	
	qtde_nfe_emitente = 0
	for i=Lbound(v_usuario_x_nfe_emitente) to UBound(v_usuario_x_nfe_emitente)
		if Not Isnull(v_usuario_x_nfe_emitente(i)) then
			qtde_nfe_emitente = qtde_nfe_emitente + 1
			id_nfe_emitente_selecionado = v_usuario_x_nfe_emitente(i)
			end if
		next
	
	if qtde_nfe_emitente > 1 then
	'	HÁ MAIS DO QUE 1 CD, ENTÃO SERÁ EXIBIDA A LISTA P/ O USUÁRIO SELECIONAR UM CD
		id_nfe_emitente_selecionado = 0
		end if
	
	if qtde_nfe_emitente = 0 then
	'	NÃO HÁ NENHUM CD CADASTRADO P/ ESTE USUÁRIO!!
		Response.Redirect("aviso.asp?id=" & ERR_NENHUM_CD_HABILITADO_PARA_USUARIO)
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
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function () {
	$("#c_dt_coleta_inicio").hUtilUI('datepicker_filtro_inicial');
	$("#c_dt_coleta_termino").hUtilUI('datepicker_filtro_final');
	});
</script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
var s_de, s_ate;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;
var data;

	if ((trim(f.c_dt_coleta_inicio.value)=="") && (trim(f.c_dt_coleta_termino.value)=="")) {
		alert("Preencha o período!!");
		f.c_dt_coleta_inicio.focus();
		return;
		}

	if (trim(f.c_dt_coleta_inicio.value)!="") {
		if (!isDate(f.c_dt_coleta_inicio)) {
			alert("Data de início inválida!!");
			f.c_dt_coleta_inicio.focus();
			return;
		}
	}

	if (trim(f.c_dt_coleta_termino.value)!="") {
		if (!isDate(f.c_dt_coleta_termino)) {
			alert("Data de término inválida!!");
			f.c_dt_coleta_termino.focus();
			return;
		}
	}

	s_de = trim(f.c_dt_coleta_inicio.value);
	s_ate = trim(f.c_dt_coleta_termino.value);

	s_de = trim(f.c_dt_coleta_inicio.value);
	s_ate = trim(f.c_dt_coleta_termino.value);
	if ((s_de != "") && (s_ate != "")) {
		s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
		s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
		if (s_de > s_ate) {
			alert("Data de término é menor que a data de início!!");
			f.c_dt_coleta_termino.focus();
			return;
		}
	}

	//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) != "") {
		strDtRefDDMMYYYY = trim(f.c_dt_coleta_inicio.value);
		if (trim(strDtRefDDMMYYYY) != "") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
			}
		}
		strDtRefDDMMYYYY = trim(f.c_dt_coleta_termino.value);
		if (trim(strDtRefDDMMYYYY) != "") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
			}
		}
	}

	if (trim(f.c_nfe_emitente.value) == "") {
		alert("É necessário selecionar um CD!!");
		return;
	}

	if (converte_numero(f.c_nfe_emitente.value) == 0) {
		alert("CD selecionado é inválido!!");
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">


<body onload="fFILTRO.c_dt_coleta_inicio.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelImpostosPagosExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<% if qtde_nfe_emitente = 1 then %>
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=Cstr(id_nfe_emitente_selecionado)%>" />
<% end if %>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Impostos Pagos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  FILTRO  -->
<table class="Qx" cellspacing="0">
<!--  PERÍODO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MC ME MD" align="left" nowrap><span class="PLTe">PERÍODO DE COLETA</span></td></tr>
	<tr bgcolor="#FFFFFF"><td class="MDBE" align="left">
		<table style="margin: 4px 8px 4px 8px;" cellspacing="0" cellpadding="0"><tr bgcolor="#FFFFFF"><td align="left">
		<input size="12" class="Cc" maxlength="10" name="c_dt_coleta_inicio" id="c_dt_coleta_inicio" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)&&tem_info(this.value)&&isDate(this)) fFILTRO.c_dt_coleta_termino.focus(); filtra_data();"
			>&nbsp;<span class="C">&nbsp;até&nbsp;</span>&nbsp;<input class="Cc" size="12" maxlength="10" name="c_dt_coleta_termino" id="c_dt_coleta_termino" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)&&tem_info(this.value)&&isDate(this)) bCONFIRMA.click(); filtra_data();">
			</td></tr>
		</table>
		</td>
	</tr>
<!--  PEDIDOS CANCELADOS/DEVOLUÇÃO  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD" align="left" nowrap><span class="PLTe">PEDIDOS</span></td></tr>
	<tr bgcolor="#FFFFFF"><td class="MDBE" align="left">
		<table style="margin: 4px 8px 4px 8px;" cellspacing="0" cellpadding="0">
		<tr>
			<td>
			<input type="checkbox" name="ckb_pedidos_cancelados" id="ckb_pedidos_cancelados" value="ON" /><span class="C" style="cursor:default;" onclick="fFILTRO.ckb_pedidos_cancelados.click();">Pedidos Cancelados</span>
			</td>
		</tr>
		<tr>
			<td>
			<input type="checkbox" name="ckb_pedidos_com_devolucao" id="ckb_pedidos_com_devolucao" value="ON" /><span class="C" style="cursor:default;" onclick="fFILTRO.ckb_pedidos_com_devolucao.click();">Pedidos com Devolução</span>
			</td>
		</tr>
		</table>
		</td>
	</tr>

<!--  OPÇÕES P/ INCLUIR/EXCLUIR OS PEDIDOS JÁ VERIFICADOS  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD" align="left" nowrap><span class="PLTe">TIPO DE CONSULTA</span></td></tr>
	<tr bgcolor="#FFFFFF"><td class="MDBE" align="left">
		<input type="radio" tabindex="-1" name="rb_tipo_consulta" id="rb_tipo_consulta_somente_nao_verificados" style="margin:6pt 2pt 0pt 9pt;"
			value="<%=COD_CONTROLE_IMPOSTOS_STATUS__INICIAL%>"
            /><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_tipo_consulta[0].click();">Somente NÃO Baixados</span>
		<br />
        <input type="radio" tabindex="-1" name="rb_tipo_consulta" id="rb_tipo_consulta_somente_ja_verificados" style="margin:2pt 2pt 0pt 9pt;"
			value="<%=COD_CONTROLE_IMPOSTOS_STATUS__OK%>"
            /><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_tipo_consulta[1].click();">Somente JÁ Baixados</span>
		<br />
        <input type="radio" tabindex="-1" name="rb_tipo_consulta" id="rb_tipo_consulta_todos" style="margin:2pt 2pt 8pt 9pt;"
			value="TODOS"
            checked
			/><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_tipo_consulta[2].click();">Todos</span>
	</td></tr>

<!--  TRANSPORTADORA  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD" align="left" nowrap><span class="PLTe">TRANSPORTADORA</span></td></tr>
	<tr bgcolor="#FFFFFF"><td class="MDBE" align="left">
		<select id="c_transportadora" name="c_transportadora" style="margin:6pt 9pt 8pt 9pt;">
		<% =transportadora_monta_itens_select("") %>
		</select>
		</td>
	</tr>

<!--  UF  -->
	<tr id='UF' bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">UF</span>
		<br>
			<select id="c_uf" name="c_uf" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =UF_monta_itens_select(Null) %>
			</select>
	</td></tr>

<% if qtde_nfe_emitente > 1 then %>
<tr>
	<td class="MB ME MD" align="left">
	<table class="Qx" cellspacing="0" cellpadding="0">
	<tr bgcolor="#FFFFFF">
		<td align="left" class="MD" width="60%" nowrap>
			<span class="PLTe">CD</span>
		</td>
		<td align="left" width="40%">
			<span class="PLTe">Consultar por CNPJ do CD</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="left" class="MD">
			<table style="margin: 4px 8px 4px 8px;" cellspacing="0" cellpadding="0">
				<tr bgcolor="#FFFFFF">
				<td align="left">
					<select id="c_nfe_emitente" name="c_nfe_emitente" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" style="margin-left:5px;margin-top:4pt; margin-bottom:4pt;">
						<%=wms_usuario_x_nfe_emitente_monta_itens_select(usuario, "")%>
					</select>
				</td>
				</tr>
			</table>
		</td>
		<td align="left">
			<input type="radio" tabindex="-1" name="rb_nfe_emitente_consultar_cnpj" id="rb_nfe_emitente_consultar_cnpj_nao" value="CONSULTAR_NAO" style="margin:6pt 2pt 0pt 9pt;"
				/>
				<span class="C" style="cursor:default" onclick="fFILTRO.rb_nfe_emitente_consultar_cnpj[0].click();">Não</span>
			<br />
			<input type="radio" tabindex="-1" name="rb_nfe_emitente_consultar_cnpj" id="rb_nfe_emitente_consultar_cnpj_sim" value="CONSULTAR_SIM" style="margin:2pt 2pt 0pt 9pt;"
				checked />
				<span class="C" style="cursor:default" onclick="fFILTRO.rb_nfe_emitente_consultar_cnpj[1].click();">Sim</span>
		</td>
	</tr>
	</table>
	</td>
</tr>
<% end if %>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
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