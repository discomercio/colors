<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ============================================
'	  CadIndicadorOpcoesFormaComoConheceuEdita.asp
'     ============================================
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
	dim s, usuario, id_selecionado, operacao_selecionada, blnStInativo
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	REGISTRO A EDITAR
	id_selecionado = trim(request("id_selecionado"))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if operacao_selecionada=OP_INCLUI then
		id_selecionado=retorna_so_digitos(id_selecionado)
		end if

	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	if operacao_selecionada<>OP_INCLUI then
		if (id_selecionado="") Or (converte_numero(id_selecionado)=0) then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICADOR_NAO_FORNECIDO)
		end if

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	if operacao_selecionada=OP_CONSULTA then
		s = "SELECT " & _
				"*" & _
			" FROM t_CODIGO_DESCRICAO" & _
			" WHERE" & _
				" (grupo = '" & GRUPO_T_CODIGO_DESCRICAO__CAD_ORCAMENTISTA_E_INDICADOR__FORMA_COMO_CONHECEU & "')" & _
				" AND (codigo = '" & id_selecionado & "')"
		rs.Open s, cn
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICADOR_NAO_CADASTRADO)
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
function RemoveRegistro( f ) {
var b;
	b=window.confirm('Confirma a exclusão deste registro?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaRegistro( f ) {
	if (trim(f.c_descricao.value)=="") {
		alert('Preencha a descrição!!');
		f.c_descricao.focus();
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
#rb_st_inativo {
	margin: 0pt 2pt 1pt 15pt;
	}
</style>


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.c_descricao.focus()"
	else
		s = "focus()"
		end if
%>
<body onLoad="<%=s%>">
<center>



<!--  DADOS DO CADASTRO  -->

<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Nova Opção para<br />""Forma como conheceu a DIS""<br />(Cadastro de Indicadores)"
	else
		s = "Consulta/Edição de Opção para<br />""Forma como conheceu a DIS""<br />(Cadastro de Indicadores)"
		end if
%>
	<td align="CENTER" vAlign="BOTTOM"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" METHOD="POST" ACTION="CadIndicadorOpcoesFormaComoConheceuAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<INPUT type=HIDDEN name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>
<INPUT type=HIDDEN name='id_selecionado' id="id_selecionado" value='<%=id_selecionado%>'>

<!-- ************   DESCRIÇÃO   ************ -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("descricao")) else s=""%>
		<td width="85%"><p class="R">DESCRIÇÃO</p><p class="C"><input id="c_descricao" name="c_descricao" class="TA" type="TEXT" maxlength="60" size="70" value="<%=s%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bATUALIZA.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   INDISPONÍVEL?   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%
	blnStInativo=False
	if operacao_selecionada=OP_CONSULTA then
		if rs("st_inativo") <> 0 then blnStInativo=True
		end if
%>
		<td width="100%">
		<p class="R">STATUS</p>
		<p class="C"><input type="RADIO" id="rb_st_inativo" name="rb_st_inativo" value="0" class="TA"<%if Not blnStInativo then Response.Write(" CHECKED")%>><span onclick="fCAD.rb_st_inativo[0].click()" style="cursor:default; color:#006600">Disponível</span>&nbsp;</p>
		<p class="C"><input type="RADIO" id="rb_st_inativo" name="rb_st_inativo" value="1" class="TA"<%if blnStInativo then Response.Write(" CHECKED")%>><span onclick="fCAD.rb_st_inativo[1].click()" style="cursor:default; color:#ff0000">Indisponível</span>&nbsp;</p>
		</td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a href="javascript:history.back()" title="cancela as alterações no cadastro">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		s = "<td align='CENTER'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveRegistro(fCAD)' "
		s =s + "title='exclui o registro'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="RIGHT"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaRegistro(fCAD)" title="grava os dados">
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
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
	cn.Close
	set cn = nothing
%>