<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%

'     =========================================
'	  M E N U O U T R A S F U N C O E S . A S P
'     =========================================
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
'						I N I C I A L I Z A     P Á G I N A     A S P
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear

'	OBTEM USUÁRIO
	dim s, usuario, usuario_nome
	usuario = trim(Session("usuario_atual"))
	usuario_nome = Trim(Session("usuario_nome_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim idx
	
'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
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

<script language="JavaScript" type="text/javascript">
window.focus();
</script>

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fOFConcluir( f ){
var s, iop;

	iop=-1;
	s="";

 // LEITURA DO QUADRO DE AVISOS (SOMENTE NÃO LIDOS)
	iop++;
	if (f.rb_op[iop].checked) {
		s="quadroavisomostra.asp";
		f.opcao_selecionada.value="";
		}

 // LEITURA DO QUADRO DE AVISOS (TODOS OS AVISOS)
	iop++;
	if (f.rb_op[iop].checked) {
		s="quadroavisomostra.asp";
		f.opcao_selecionada.value="S";
		}

 // PERDA (CADASTRAMENTO DE VALOR A TÍTULO DE FUNDO PERDIDO)
	iop++;
	if (f.rb_op[iop].checked) {
		s="Perda.asp";
		}

	// EDITAR DADOS DE ETIQUETA (WMS)
	iop++;
	if (f.rb_op[iop].checked) {
		s = "EtqWmsEtiquetaObtemId.asp";
	}
	
	if (s=="") {
		alert("Escolha uma das funções!!");
		return false;
		}

	window.status = "Aguarde ...";
	f.action=s;
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
<!--  MENU SUPERIOR -->

<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="RIGHT" vAlign="BOTTOM"><p class="PEDIDO">CENTRAL&nbsp;&nbsp;ADMINISTRATIVA<br>
	<%	s = usuario_nome
		if s = "" then s = usuario
		s = x_saudacao & ", " & s
		s = "<span class='Cd'>" & s & "</span><br>"
	%>
	<%=s%>
	<span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="senha.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="altera a senha atual do usuário" class="LAlteraSenha">altera senha</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></p></td>
	</tr>

</table>

<br>
<center>


<!--  ***********************************************************************************************  -->
<!--  O U T R A S   F U N Ç Õ E S                         											   -->
<!--  ***********************************************************************************************  -->
<form method="post" id="fOF" name="fOF" onsubmit="if (!fOFConcluir(fOF)) return false">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="opcao_selecionada" id="opcao_selecionada" value=''>
<input type="hidden" name="opcao_alerta_se_nao_ha_aviso" id="opcao_alerta_se_nao_ha_aviso" value='S'>
<span class="T">OUTRAS FUNÇÕES</span>
<div class="QFn" align="CENTER">
<table class="TFn">
	<tr>
		<td NOWRAP>
			<%	idx = 0	%>
			<%  idx=idx+1 %>
			<% if operacao_permitida(OP_CEN_LER_AVISOS_NAO_LIDOS, s_lista_operacoes_permitidas) then s="" else s=" DISABLED" %>
			<input type="RADIO" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOF.rb_op[<%=Cstr(idx-1)%>].click(); fOF.bEXECUTAR.click();"
				>Ler Quadro de Avisos (somente não lidos)</span><br>
			<%  idx=idx+1 %>
			<% if operacao_permitida(OP_CEN_LER_AVISOS_TODOS, s_lista_operacoes_permitidas) then s="" else s=" DISABLED" %>
			<input type="RADIO" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOF.rb_op[<%=Cstr(idx-1)%>].click(); fOF.bEXECUTAR.click();"
				>Ler Quadro de Avisos (todos os avisos)</span><br>
			<%  idx=idx+1 %>
			<% if operacao_permitida(OP_CEN_CADASTRA_PERDA, s_lista_operacoes_permitidas) then s="" else s=" DISABLED" %>
			<input type="RADIO" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOF.rb_op[<%=Cstr(idx-1)%>].click(); if (fOF.rb_op[<%=Cstr(idx-1)%>].checked) fOF.bEXECUTAR.click();"
				>Perda</span><br>
			<%  idx=idx+1 %>
			<% if operacao_permitida(OP_CEN_ETQWMS_EDITA_DADOS_ETIQUETA, s_lista_operacoes_permitidas) then s="" else s=" DISABLED" %>
			<input type="RADIO" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOF.rb_op[<%=Cstr(idx-1)%>].click(); if (fOF.rb_op[<%=Cstr(idx-1)%>].checked) fOF.bEXECUTAR.click();"
				>Editar Dados de Etiqueta (WMS)</span>
			</td>
		</tr>
	</table>

	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bEXECUTAR" id="bEXECUTAR" type="SUBMIT" class="Botao" value="EXECUTAR" title="executa">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
	<td align="CENTER"><a href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>

</body>
</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
