<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  R E L P E D I D O S C R E D P E N D F I L T R O . A S P
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

	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

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
	
	strSql = "SELECT DISTINCT id, nome, nome_iniciais_em_maiusculas FROM" & _
			 " (" & _
			 "SELECT DISTINCT t_USUARIO.usuario AS id, nome, nome_iniciais_em_maiusculas FROM" & _
			 " t_USUARIO INNER JOIN t_USUARIO_X_LOJA ON t_USUARIO.usuario=t_USUARIO_X_LOJA.usuario" & _
			 " WHERE (vendedor_loja <> 0) AND " & _
			 SCHEMA_BD & ".UsuarioPossuiAcessoLoja(t_USUARIO.usuario, '" & loja & "') = 'S'" & _
			 " UNION" & _
			 " SELECT DISTINCT t_PEDIDO__BASE.vendedor AS id, nome, nome_iniciais_em_maiusculas FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido) INNER JOIN t_USUARIO ON t_PEDIDO__BASE.vendedor=t_USUARIO.usuario WHERE (t_PEDIDO__BASE.numero_loja = " & loja & ") AND (t_PEDIDO__BASE.analise_credito = " & COD_AN_CREDITO_PENDENTE & ")" & _
			 ") AS t" & _
			 " ORDER BY id"
	
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("id")))
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

<form id="fFILTRO" name="fFILTRO" method="post" action="RelPedidosCredPendExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos com Crédito Pendente</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table class="Qx" cellSpacing="0">
<!--  CADASTRAMENTO  -->
<tr bgColor="#FFFFFF">
<td class="MT" NOWRAP><span class="PLTe">VENDEDOR</span>
	<br>
	<table cellSpacing="6" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF">
		<td>
			<select id="vendedor_selecionado" name="vendedor_selecionado" style="margin-left:10px;margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =vendedores_monta_itens_select(Null) %>
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