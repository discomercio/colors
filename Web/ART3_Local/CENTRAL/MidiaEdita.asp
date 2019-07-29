<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  M I D I A E D I T A . A S P
'     =====================================
'
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
	
'	OBTEM O ID
	dim s, usuario, midia_selecionada, operacao_selecionada, midia_indisponivel
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	MÍDIA A EDITAR
	midia_selecionada = trim(request("midia_selecionada"))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if operacao_selecionada=OP_INCLUI then
		midia_selecionada=retorna_so_digitos(midia_selecionada)
		end if

	midia_selecionada=normaliza_codigo(midia_selecionada, TAM_MIN_MIDIA)
	
	if (midia_selecionada="") Or (midia_selecionada="000") then Response.Redirect("aviso.asp?id=" & ERR_MIDIA_NAO_ESPECIFICADA) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)

'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	set rs = cn.Execute("select * from t_MIDIA where (id='" & midia_selecionada & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_MIDIA_JA_CADASTRADA)
	'	GARANTE QUE O Nº DO VEÍCULO DE MÍDIA NÃO ESTÁ EM USO
		rs.Close
		set rs = cn.Execute("select * from t_MIDIA where (CONVERT(smallint, id) = " & midia_selecionada & ")")
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_MIDIA_JA_CADASTRADA)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_MIDIA_NAO_CADASTRADA)
		end if
	
%>

<html>


<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>


<%
'		C L I E N T   S I D E   S C R I P T I N G
'
'      CCCCCCC   LLL        IIIII  EEEEEEEEE  NNN   NNN  TTTTTTTTT EEEEEEEEE
'     CCC   CCC  LLL         III   EEE        NNNN  NNN     TTT    EEE
'     CCC        LLL         III   EEE        NNNNN NNN     TTT    EEE
'     CCC        LLL         III   EEEEEE     NNN NNNNN     TTT    EEEEEE
'     CCC        LLL         III   EEE        NNN  NNNN     TTT    EEE
'     CCC   CCC  LLL   LLL   III   EEE        NNN   NNN     TTT    EEE
'      CCCCCCC   LLLLLLLLL  IIIII  EEEEEEEEE  NNN   NNN     TTT    EEEEEEEEE
'
%>

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function RemoveMidia( f ) {
var b;
	b=window.confirm('Confirma a exclusão deste veículo de mídia?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaMidia( f ) {
	if (trim(f.nome.value)=="") {
		alert('Preencha o nome!!');
		f.nome.focus();
		return;
		}
	dATUALIZA.style.visibility="hidden";
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style TYPE="text/css">
#rb_indisponivel {
	margin: 0pt 2pt 1pt 15pt;
	}
</style>


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.nome.focus()"
	else
		s = "focus()"
		end if
%>
<body onload="<%=s%>">
<center>



<!--  CADASTRO DO VEÍCULO DE MÍDIA -->

<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Novo Veículo de Mídia"
	else
		s = "Consulta/Edição de Veículo de Mídia Cadastrado"
		end if
%>
	<td align="center" valign="bottom"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="MidiaAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>

<!-- ************   NÚMERO/NOME   ************ -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
		<td class="MD" width="15%"><p class="R">MÍDIA</p><p class="C"><input id="midia_selecionada" name="midia_selecionada" class="TA" value="<%=midia_selecionada%>" readonly size="6" style="text-align:center; color:#0000ff"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("apelido")) else s=""%>
		<td width="85%"><p class="R">NOME (apelido)</p><p class="C"><input id="nome" name="nome" class="TA" type="TEXT" maxlength="30" size="60" value="<%=s%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bATUALIZA.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   INDISPONÍVEL?   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%
	midia_indisponivel=false
	if operacao_selecionada=OP_CONSULTA then
		if rs("indisponivel") <> 0 then midia_indisponivel=true
		end if
%>
		<td width="100%">
		<p class="R">ESTADO</p>
		<p class="C"><input type="radio" id="rb_indisponivel" name="rb_indisponivel" value="0" class="TA"<%if not midia_indisponivel then Response.Write(" checked")%>><span onclick="fCAD.rb_indisponivel[0].click();" style="cursor:default; color:#006600">Disponível</span>&nbsp;</p>
		<p class="C"><input type="radio" id="rb_indisponivel" name="rb_indisponivel" value="1" class="TA"<%if midia_indisponivel then Response.Write(" checked")%>><span onclick="fCAD.rb_indisponivel[1].click();" style="cursor:default; color:#ff0000">Indisponível</span>&nbsp;</p>
		</td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td><a href="javascript:history.back()" title="cancela as alterações no cadastro">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		s = "<td align='CENTER'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveMidia(fCAD)' "
		s =s + "title='remove o veículo de mídia cadastrado'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaMidia(fCAD)" title="atualiza o cadastro do veículo de mídia">
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
	rs.Close
	set rs = nothing
	
	cn.Close
	set cn = nothing
%>