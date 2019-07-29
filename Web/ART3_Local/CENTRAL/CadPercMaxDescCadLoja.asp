<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================
'	  CadPercMaxDescCadLoja.asp
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

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_CAD_PARAMETROS_GLOBAIS, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim s, s_nome_loja
	dim n_reg, n_reg_total
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

function realca_cor_row(id) {
var c;
	c = document.getElementById(id);
	c.style.backgroundColor = 'palegreen';
}

function normaliza_cor_row(id) {
var c;
	c = document.getElementById(id);
	c.style.backgroundColor = '';
}

function AtualizaCadastro( f ){
	dATUALIZA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}

</script>

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style TYPE="text/css">
.tdNumLoja{
	vertical-align: top;
	width: 70px;
	}
.tdNomeLoja{
	vertical-align: top;
	width: 250px;
	}
.tdPerc{
	vertical-align: top;
	width: 90px;
	}
</style>


<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="RIGHT" vAlign="BOTTOM" NOWRAP><span class="PEDIDO">Percentual Máximo da Senha de Desconto<BR>para Cadastramento na Loja</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<br>
<center>
<form METHOD="POST" ACTION="CadPercMaxDescCadLojaConfirma.asp" id="fCAD" name="fCAD">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="HIDDEN" id="c_loja" name="c_loja" value=''>
<input type="HIDDEN" id="c_nome_loja" name="c_nome_loja" value=''>
<input type="HIDDEN" id="c_percentual" name="c_percentual" value=''>

<%
'	RECUPERA VALOR ANTERIOR
	s = "SELECT" & _
			" loja," & _
			" PercMaxSenhaDesconto," & _
			" nome," & _
			" razao_social" & _
		" FROM t_LOJA" & _
		" ORDER BY" & _
			" Convert(smallint, loja)"
	if rs.State <> 0 then rs.Close
	rs.open s, cn
%>


<table border="0" cellspacing="0" cellpadding="0">
<tr style="background-color:whitesmoke;">
	<td class="MT tdNumLoja C" align="center" style="font-size:10pt;">Nº Loja</td>
	<td class="MC MB MD tdNomeLoja C" align="left" style="font-size:10pt;">Nome Loja</td>
	<td class="MC MB MD tdPerc Cd" align="right" style="font-size:10pt;">Percentual</td>
</tr>
<%  n_reg_total = 0
	do while Not rs.Eof
		n_reg_total = n_reg_total + 1
		rs.MoveNext
		loop
		
	rs.MoveFirst
	
	n_reg = 0
	do while Not rs.Eof
		n_reg = n_reg + 1
		s_nome_loja = Trim("" & rs("nome"))
		if s_nome_loja = "" then s_nome_loja = Trim("" & rs("razao_social"))
		if s_nome_loja <> "" then s_nome_loja = iniciais_em_maiusculas(s_nome_loja)
%>
<tr id="TR_<%=Cstr(n_reg)%>">
	<td class="MB MD ME tdNumLoja" align="center"><input id="c_loja" name="c_loja" readonly tabindex=-1 class="PLLc" style="font-size:10pt;width:70px;background-color:Transparent;" value="<%=Trim("" & rs("loja"))%>"></td>
	<td class="MB MD tdNomeLoja" align="left"><input id="c_nome_loja" name="c_nome_loja" readonly tabindex=-1 class="PLLe" style="font-size:10pt;width:300px;background-color:Transparent;" value="<%=s_nome_loja%>"></td>
	<td class="MB MD tdPerc" align="right"><input id="c_percentual" name="c_percentual" class="PLLd" style="font-size:10pt;width:70px;background-color:Transparent;" value="<%=formata_perc(rs("PercMaxSenhaDesconto"))%>" maxlength="5" 
		<% if n_reg = n_reg_total then %>
		onkeypress="if (digitou_enter(true)) bATUALIZA.focus(); filtra_percentual();"
		<% else %>
		onkeypress="if (digitou_enter(true)) fCAD.c_percentual[<%=Cstr(n_reg+1)%>].focus(); filtra_percentual();"
		<% end if %>
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
</tr>
<%		rs.MoveNext
		loop
%>
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