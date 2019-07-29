<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================
'	  A N A L I S E C R E D I T O C A D A P R O V A U T O . A S P
'     ===========================================================
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


' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
'	OBTEM USUÁRIO
	dim usuario
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s, m_valor
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
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando consulta ...';

function AtualizaCadastro( f ){
	dATUALIZA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit(); 
}

</script>

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">



<body onload="window.status='Concluído';fCAD.c_valor.focus();" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="RIGHT" vAlign="BOTTOM" NOWRAP><span class="PEDIDO">Valor para Aprovação Automática da Análise de Crédito</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  VALOR P/ APROVAÇÃO AUTOMÁTICA DA ANÁLISE DE CRÉDITO  -->
<br>
<center>
<form METHOD="POST" ACTION="AnaliseCreditoCadAprovAutoConfirma.asp" id="fCAD" name="fCAD">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<%
'	RECUPERA VALOR ANTERIOR
	s = "SELECT nsu FROM t_CONTROLE WHERE (id_nsu = '" & ID_PARAM_CAD_VL_APROV_AUTO_ANALISE_CREDITO & "')"
	if rs.State <> 0 then rs.Close
	rs.open s, cn
	if Not rs.Eof then
		m_valor = converte_numero(rs("nsu"))
	else
		m_valor = 0
		end if
%>
<table border="0" cellspacing="0" cellpadding="0">
<tr style="background:#FFF0E0">
	<td class="MT" align="center" style="width:160px;"><p class="C" style="font-size:10pt;">Valor Limite</p></td>
	</tr>
<tr>
	<td NOWRAP class="MDE" align="right"><p class="PLTd"><%=SIMBOLO_MONETARIO%></p></td>
	</tr>
<tr>
	<td class="MDBE" align="right"><input id="c_valor" name="c_valor" class="PLLd" style="font-size:10pt;" value="<%=formata_moeda(m_valor)%>" maxlength="18" onkeypress="if (digitou_enter(true)) bATUALIZA.focus(); filtra_moeda();" onblur="this.value=formata_moeda(this.value);"></td>
	</tr>
</table>

</form>

<BR>

<p class="TracoBottom"></p>

<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="LEFT"><a href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="RIGHT"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaCadastro(fCAD)" title="salva as alterações">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>

</center>


</body>
</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>