<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  V I S A N E T O P C O E S P A G T O . A S P
'     ===========================================
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

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_OPCOES_PAGTO_VISANET, s_lista_operacoes_permitidas) then
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim s, s_descricao, n_qtde, m_valor
	dim iQtdeBandeira, ic
	dim vBandeira
	iQtdeBandeira = 0
	vBandeira = CieloArrayBandeiras
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
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" Language="JavaScript" Type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_GLOBAL%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" Language="JavaScript" Type="text/javascript"></script>
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



<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Opções de Prazo de Pagamento para Cartão de Crédito</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  OPÇÕES DE PRAZO DE PAGAMENTO PARA CARTÃO DE CRÉDITO  -->
<br>
<center>
<form method="post" action="VisanetOpcoesPagtoAtualiza.asp" id="fCAD" name="fCAD">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<table border="0" cellspacing="0" cellpadding="0">

<%	for ic=Lbound(vBandeira) to Ubound(vBandeira)
		if Trim(vBandeira(ic)) <> "" then
			iQtdeBandeira = iQtdeBandeira + 1
			if iQtdeBandeira > 1 then
%>
	<tr>
		<td align="left" colspan="4">&nbsp;</td>
	</tr>

	<tr>
		<td align="left" colspan="4">&nbsp;</td>
	</tr>
			<% end if %>

	<tr style="background:azure">
		<td align="right" rowspan="7" style="background:white;"><img src="../Imagem/Cielo/<%=CieloObtemNomeArquivoLogo(vBandeira(ic))%>" border="0" style="margin-right:10px;"></td>
		<td colspan="3" class="MT" align="center" style="width:330px;"><span class="C" style="font-size:12pt;">Bandeira: <%=CieloDescricaoBandeira(vBandeira(ic))%></span></td>
	</tr>

<%
'	PARCELAMENTO PELA LOJA
	s = "SELECT * FROM T_PRAZO_PAGTO_VISANET WHERE (tipo = '" & CieloObtemIdRegistroBdPrazoPagtoLoja(vBandeira(ic)) & "')"
	if rs.State <> 0 then rs.Close
	rs.open s, cn
	if Not rs.Eof then
		s_descricao = Trim("" & rs("descricao"))
		n_qtde = rs("qtde_parcelas")
		m_valor = rs("vl_min_parcela")
	else
		s_descricao = "Parcelamento pela loja (sem juros)"
		n_qtde = 0
		m_valor = 0
		end if
%>
<tr style="background:gainsboro">
	<td colspan="4" class="ME MD MB" align="center" style="width:330px;"><span class="C" style="font-size:10pt;"><%=s_descricao%></span></td>
</tr>
<tr>
	<td align="left" nowrap class="ME"><span class="PLTe">Qtde de Parcelas</span></td>
	<td align="left" class="MD" style="width:20px;">&nbsp;</td>
	<td align="right" nowrap class="MD"><span class="PLTd">Parcela Mínima (<%=SIMBOLO_MONETARIO%>)</span></td>
</tr>
<tr>
	<td align="left" class="MEB"><input id="<%="C_QTDE_" & CieloObtemIdRegistroBdPrazoPagtoLoja(vBandeira(ic))%>" name="<%="C_QTDE_" & CieloObtemIdRegistroBdPrazoPagtoLoja(vBandeira(ic))%>" class="PLLc" style="width:50px;font-size:10pt;" value="<%=formata_inteiro(n_qtde)%>" maxlength="2" onfocus="this.select();" onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();" onblur="this.value=formata_inteiro(this.value);"></td>
	<td align="left" class="MDB" style="width:20px;">&nbsp;</td>
	<td align="right" class="MDB"><input id="<%="C_VL_" & CieloObtemIdRegistroBdPrazoPagtoLoja(vBandeira(ic))%>" name="<%="C_VL_" & CieloObtemIdRegistroBdPrazoPagtoLoja(vBandeira(ic))%>" class="PLLd" style="font-size:10pt;" value="<%=formata_moeda(m_valor)%>" maxlength="18" onfocus="this.select();" onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_moeda();" onblur="this.value=formata_moeda(this.value);"></td>
</tr>

<%
'	PARCELAMENTO PELO EMISSOR DO CARTÃO
	s = "SELECT * FROM T_PRAZO_PAGTO_VISANET WHERE (tipo = '" & CieloObtemIdRegistroBdPrazoPagtoEmissor(vBandeira(ic)) & "')"
	if rs.State <> 0 then rs.Close
	rs.open s, cn
	if Not rs.Eof then
		s_descricao = Trim("" & rs("descricao"))
		n_qtde = rs("qtde_parcelas")
		m_valor = rs("vl_min_parcela")
	else
		s_descricao = "Parcelamento pelo emissor do cartão (com juros)"
		n_qtde = 0
		m_valor = 0
		end if
%>
<tr style="background:gainsboro">
	<td colspan="4" class="ME MD MB" align="center" style="width:330px;"><span class="C" style="font-size:10pt;"><%=s_descricao%></span></td>
</tr>
<tr>
	<td align="left" nowrap class="ME"><span class="PLTe">Qtde de Parcelas</span></td>
	<td align="left" class="MD" style="width:20px;">&nbsp;</td>
	<td align="right" nowrap class="MD"><span class="PLTd">Parcela Mínima (<%=SIMBOLO_MONETARIO%>)</span></td>
</tr>
<tr>
	<td align="left" class="MEB"><input id="<%="C_QTDE_" & CieloObtemIdRegistroBdPrazoPagtoEmissor(vBandeira(ic))%>" name="<%="C_QTDE_" & CieloObtemIdRegistroBdPrazoPagtoEmissor(vBandeira(ic))%>" class="PLLc" style="width:50px;font-size:10pt;" value="<%=formata_inteiro(n_qtde)%>" maxlength="2" onfocus="this.select();" onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();" onblur="this.value=formata_inteiro(this.value);"></td>
	<td align="left" class="MDB" style="width:20px;">&nbsp;</td>
	<td align="right" class="MDB"><input id="<%="C_VL_" & CieloObtemIdRegistroBdPrazoPagtoEmissor(vBandeira(ic))%>" name="<%="C_VL_" & CieloObtemIdRegistroBdPrazoPagtoEmissor(vBandeira(ic))%>" class="PLLd" style="font-size:10pt;" value="<%=formata_moeda(m_valor)%>" maxlength="18" onfocus="this.select();" onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_moeda();" onblur="this.value=formata_moeda(this.value);"></td>
</tr>
<%		end if
	next
%>

</table>

</form>

<br />

<p class="TracoBottom"></p>

<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="left"><a href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
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