<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  R E L A N A L I S E C R E D I T O F I L T R O . A S P
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

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)






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
	strSql = "SELECT DISTINCT usuario, nome, nome_iniciais_em_maiusculas FROM" & _
			 " (" & _
			 "SELECT usuario, nome, nome_iniciais_em_maiusculas FROM t_USUARIO" & _
				" WHERE (vendedor_loja <> 0)" & _
			 " UNION" & _
			 " SELECT t_USUARIO.usuario AS usuario, t_USUARIO.nome AS nome, t_USUARIO.nome_iniciais_em_maiusculas FROM t_USUARIO" & _
				" INNER JOIN t_PERFIL_X_USUARIO ON (t_USUARIO.usuario=t_PERFIL_X_USUARIO.usuario)" & _
				" INNER JOIN t_PERFIL ON (t_PERFIL_X_USUARIO.id_perfil=t_PERFIL.id)" & _
				" INNER JOIN t_PERFIL_ITEM ON (t_PERFIL.id=t_PERFIL_ITEM.id_perfil)" & _
				" WHERE (t_PERFIL_ITEM.id_operacao=" & OP_CEN_ACESSO_TODAS_LOJAS & ")" & _
			 ") AS t" & _
			 " ORDER BY usuario"
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
	// Consistência
	if ((trim(f.c_valor_inferior.value) != "") && (trim(f.c_valor_superior.value) != "")) {
		if (converte_numero(trim(f.c_valor_superior.value)) < converte_numero(trim(f.c_valor_inferior.value))) {
			alert("O valor limite superior é menor que o valor limite inferior!!");
			f.c_valor_superior.focus();
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


<body onload="fFILTRO.c_pedidos.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelAnaliseCredito.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Análise de Crédito</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0" style="width:190px;">
<!--  PEDIDOS  -->
	<tr bgColor="#FFFFFF">
	<td class="MT" NOWRAP><span class="PLTe">PEDIDO(S)</span>
	<br>
		<textarea rows="8" name="c_pedidos" id="c_pedidos" class="PLBe" style="font-size:9pt;width:110px;margin-bottom:4px;margin-left:20px;" onkeypress="if (!digitou_enter(false)) filtra_pedido();" onblur="this.value=normaliza_lista_pedidos(this.value);"></textarea>
	</td></tr>
<!-- ************   LOJAS   ************ -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">LOJA(S)</span>
		<br>
			<textarea class="PLBe" style="font-size:9pt;width:110px;margin-bottom:4px;margin-left:20px;" rows="6" name="c_lista_loja" id="c_lista_loja" onkeypress="if (!digitou_enter(false) && !digitou_char('-')) filtra_numerico();" onblur="this.value=normaliza_lista_lojas(this.value);"></textarea>
	</td></tr>
<!-- ************   VALOR   ************ -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">VALOR (<%=SIMBOLO_MONETARIO%>)</span>
		<br>
			<input maxlength="12" name="c_valor_inferior" id="c_valor_inferior" style="width:120px;text-align:right;margin-left:15px;margin-bottom:4px;margin-left:20px;" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fFILTRO.c_valor_superior.focus(); filtra_moeda();" onblur="this.value=formata_moeda(this.value);"><span class="PLTe" style="vertical-align:middle;margin-left:8px;margin-right:8px;">A</span>
			<input maxlength="12" name="c_valor_superior" id="c_valor_superior" style="width:120px;text-align:right;margin-right:15px;margin-bottom:4px;" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) bCONFIRMA.focus(); filtra_moeda();" onblur="this.value=formata_moeda(this.value);">
	</td></tr>
	
<!-- ************   VENDEDOR / INDICADOR   ************ -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">CADASTRAMENTO</span>
	<br>
	<table cellSpacing="6" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF">
		<td align="right"><span class="C" style="margin-left:8px;">Vendedor</span></td>
		<td>
			<select id="c_vendedor" name="c_vendedor" style="margin-right:8px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =vendedores_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td align="right"><span class="C" style="margin-left:8px;">Indicador</span></td>
		<td>
			<select id="c_indicador" name="c_indicador" style="margin-right:8px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =indicadores_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>
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
	<td><a name="bVOLTAR" id="bVOLTAR" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="confirma a operação">
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
