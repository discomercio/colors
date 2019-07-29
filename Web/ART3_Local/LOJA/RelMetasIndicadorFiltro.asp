<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  RelMetasIndicadorFiltro.asp
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

	Const COD_CONSULTA_POR_PERIODO_CADASTRO = "CADASTRO"
	Const COD_CONSULTA_POR_PERIODO_ENTREGA = "ENTREGA"

	dim usuario, loja, s, intIdx
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
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




		
' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________


' ____________________________________________________________________________
' VENDEDORES MONTA ITENS SELECT
'
function vendedores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT t_USUARIO.usuario, nome_iniciais_em_maiusculas FROM" & _
			 " t_USUARIO INNER JOIN t_USUARIO_X_LOJA ON t_USUARIO.usuario=t_USUARIO_X_LOJA.usuario" & _
			 " WHERE (vendedor_loja <> 0) AND " & _
			 SCHEMA_BD & ".UsuarioPossuiAcessoLoja(t_USUARIO.usuario, '" & loja & "') = 'S'" & _
			 " ORDER BY t_USUARIO.usuario"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	vendedores_monta_itens_select = strResp
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



<html>


<head>
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
var s_de, s_ate;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;
var i, blnFlagOk;

//  TIPO DE CONSULTA: PEDIDOS CADASTRADOS OU PEDIDOS ENTREGUES
	blnFlagOk=false;
	for (i=0; i<f.rb_periodo.length; i++) {
		if (f.rb_periodo[i].checked) blnFlagOk=true;
		}
	if (!blnFlagOk) {
		alert("Selecione o tipo de consulta:\n    Por pedidos cadastrados\n    Por pedidos entregues");
		return;
		}

	if (f.rb_periodo[0].checked) {
		if ((converte_numero(f.c_dt_cadastro_mes.value)<=0)||(converte_numero(f.c_dt_cadastro_mes.value)>12)) {
			alert("Mês inválido!!");
			f.c_dt_cadastro_mes.focus();
			return;
			}

		if (converte_numero(f.c_dt_cadastro_ano.value)<2000) {
			alert("Ano inválido!!");
			f.c_dt_cadastro_ano.focus();
			return;
			}
		}

	if (f.rb_periodo[1].checked) {
		if ((converte_numero(f.c_dt_entregue_mes.value)<=0)||(converte_numero(f.c_dt_entregue_mes.value)>12)) {
			alert("Mês inválido!!");
			f.c_dt_entregue_mes.focus();
			return;
			}

		if (converte_numero(f.c_dt_entregue_ano.value)<2000) {
			alert("Ano inválido!!");
			f.c_dt_entregue_ano.focus();
			return;
			}
		}
		
//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
	//  PERÍODO DE CADASTRO
		if (f.rb_periodo[0].checked) {
			strDtRefDDMMYYYY = "01/" + trim(f.c_dt_cadastro_mes.value) + "/" + trim(f.c_dt_cadastro_ano.value);
			if (trim(strDtRefDDMMYYYY)!="") {
				strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
				if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
					alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
					return;
					}
				}
			}
	
	// PERÍODO DE ENTREGA
		if (f.rb_periodo[1].checked) {
			strDtRefDDMMYYYY = "01/" + trim(f.c_dt_entregue_mes.value) + "/" + trim(f.c_dt_entregue_ano.value);
			if (trim(strDtRefDDMMYYYY)!="") {
				strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
				if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
					alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
					return;
					}
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


<body onload="fFILTRO.c_dt_cadastro_mes.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelMetasIndicadorExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Metas do Indicador</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PEDIDOS CADASTRADOS EM  -->
<% intIdx=-1 %>
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MT" NOWRAP valign="bottom">
		<input type="radio" id="rb_periodo" name="rb_periodo" value="<%=COD_CONSULTA_POR_PERIODO_CADASTRO%>">
		<% intIdx=intIdx+1 %>
		<span class="PLTe" style="vertical-align:middle;cursor:default;margin-right:10px;" onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();">PEDIDOS CADASTRADOS EM</span>
	<br>
		<table cellSpacing="0" cellPadding="0" style="margin-top:6px;margin-bottom:4px;"><tr bgColor="#FFFFFF"><td>
		<span class="PLLc" style="color:#808080;" style="margin-left:30px;">&nbsp;Mês</span>
		<input class="PLBc" maxlength="2" style="width:24px;" name="c_dt_cadastro_mes" id="c_dt_cadastro_mes" 
			onblur="if ( (trim(this.value)!='') && ( (converte_numero(this.value)<=0)||(converte_numero(this.value)>12) ) ) {alert('Mês inválido!'); this.focus();} else {if (converte_numero(this.value)>0) {while (this.value.length<2) this.value='0'+this.value;}}"
			onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fFILTRO.c_dt_cadastro_ano.focus(); filtra_numerico();"
			<%	s=Cstr(Month(Date))
				do while len(s)<2: s = "0" & s: loop%>
				value='<%=s%>'>
			&nbsp;
			<span class="PLLc" style="color:#808080;">&nbsp;&nbsp;Ano</span>
			<input class="PLBc" maxlength="4" style="width:40px;" name="c_dt_cadastro_ano" id="c_dt_cadastro_ano" 
				onblur="if ((trim(this.value)!='')&&(converte_numero(this.value)<2000)) {alert('Ano inválido!'); this.focus();}"
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fFILTRO.c_vendedor.focus(); filtra_numerico();"
				value='<%=Cstr(Year(Date))%>'>
			</td></tr>
		</table>
		</td></tr>

	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP valign="bottom">
		<input type="radio" id="rb_periodo" name="rb_periodo" value="<%=COD_CONSULTA_POR_PERIODO_ENTREGA%>">
		<% intIdx=intIdx+1 %>
		<span class="PLTe" style="vertical-align:middle;cursor:default;margin-right:10px;" onclick="fFILTRO.rb_periodo[<%=Cstr(intIdx)%>].click();">PEDIDOS ENTREGUES EM</span>
	<br>
		<table cellSpacing="0" cellPadding="0" style="margin-top:6px;margin-bottom:4px;"><tr bgColor="#FFFFFF"><td>
		<span class="PLLc" style="color:#808080;" style="margin-left:30px;">&nbsp;Mês</span>
		<input class="PLBc" maxlength="2" style="width:24px;" name="c_dt_entregue_mes" id="c_dt_entregue_mes" 
			onblur="if ( (trim(this.value)!='') && ( (converte_numero(this.value)<=0)||(converte_numero(this.value)>12) ) ) {alert('Mês inválido!'); this.focus();} else {if (converte_numero(this.value)>0) {while (this.value.length<2) this.value='0'+this.value;}}"
			onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fFILTRO.c_dt_entregue_ano.focus(); filtra_numerico();"
			<%	s=Cstr(Month(Date))
				do while len(s)<2: s = "0" & s: loop%>
				value='<%=s%>'>
			&nbsp;
			<span class="PLLc" style="color:#808080;">&nbsp;&nbsp;Ano</span>
			<input class="PLBc" maxlength="4" style="width:40px;" name="c_dt_entregue_ano" id="c_dt_entregue_ano" 
				onblur="if ((trim(this.value)!='')&&(converte_numero(this.value)<2000)) {alert('Ano inválido!'); this.focus();}"
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fFILTRO.c_vendedor.focus(); filtra_numerico();"
				value='<%=Cstr(Year(Date))%>'>
			</td></tr>
		</table>
		</td></tr>

	<tr>
		<td>&nbsp;</td>
	</tr>
	
<!--  VENDEDOR  -->
	<% if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then %>
	<tr bgColor="#FFFFFF">
	<td class="MT" NOWRAP><span class="PLTe">VENDEDOR</span>
	<br>
		<select id="c_vendedor" name="c_vendedor" style="margin-left:10px;margin-right:10px;margin-bottom:8px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
		<% =vendedores_monta_itens_select(Null) %>
		</select>
		</td></tr>
	<% else %>
	<input type="hidden" name="c_vendedor" id="c_vendedor" value='<%=usuario%>'>
	<% end if %>

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
