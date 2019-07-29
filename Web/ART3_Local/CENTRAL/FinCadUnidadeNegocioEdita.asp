<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  FinCadUnidadeNegocioEdita.asp
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
	dim s, strSql, usuario, id_selecionado, operacao_selecionada
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	REGISTRO A EDITAR
	id_selecionado = trim(request("id_selecionado"))
	operacao_selecionada = trim(request("operacao_selecionada"))

	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)

	if (operacao_selecionada<>OP_INCLUI) then
		if (id_selecionado="") Or (converte_numero(id_selecionado)=0) then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_INFORMADO)
		end if

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	if (operacao_selecionada<>OP_INCLUI) then
		strSql = "SELECT " & _
					"*" & _
				" FROM t_FIN_UNIDADE_NEGOCIO" & _
				" WHERE" & _
					" (id = " & id_selecionado & ")"
		set rs = cn.Execute(strSql)
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_CADASTRADO)
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
	b=window.confirm('Confirma a exclusão?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaRegistro(f) {
	if (trim(f.c_apelido.value) == "") {
		alert("Informe o apelido da unidade de negócio!!");
		f.c_apelido.focus();
		return;
		}
	if (trim(f.c_descricao.value)=="") {
		alert('Preencha o nome!!');
		f.c_descricao.focus();
		return;
		}
//  PARA O CASO DE TER CLICADO NO BOTÃO BACK APÓS TER CLICADO NA OPERAÇÃO EXCLUIR
	f.operacao_selecionada.value=f.operacao_selecionada_original.value;
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
#rb_st_ativo {
	margin: 0pt 2pt 1pt 15pt;
	}
</style>


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.c_apelido.focus()"
	else
		s = "focus()"
		end if
%>
<body onLoad="<%=s%>">
<center>



<!--  FORMULÁRIO DE CADASTRO  -->

<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Unidades de Negócio: Cadastro de Nova Unidade"
	else
		s = "Unidades de Negócio: Consulta/Edição"
		end if
%>
	<td align="CENTER" vAlign="BOTTOM"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" METHOD="POST" ACTION="FinCadUnidadeNegocioAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<INPUT type=HIDDEN name='operacao_selecionada_original' id="operacao_selecionada_original" value='<%=operacao_selecionada%>'>
<INPUT type=HIDDEN name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>
<INPUT type=HIDDEN name='id_selecionado' id="id_selecionado" value='<%=id_selecionado%>'>

<!-- ************   APELIDO / DESCRIÇÃO   ************ -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
		<td class="MD" width="25%">
			<p class="R">APELIDO</p>
			<p class="C">
				<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("apelido")) else s=""%>
				<input id="c_apelido" name="c_apelido" class="TA" value="<%=s%>" 
					maxlength="10" style="width:140px;"
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_descricao.focus(); filtra_nome_identificador();">
			</p>
		</td>
		<td width="75%">
			<p class="R">DESCRIÇÃO</p>
			<p class="C">
				<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("descricao")) else s=""%>
				<input id="c_descricao" name="c_descricao" class="TA" type="TEXT" maxlength="40" size="60" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) bATUALIZA.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   STATUS ATIVO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%
	dim st_ativo
	st_ativo=false
	if operacao_selecionada=OP_CONSULTA then
		if Cstr(rs("st_ativo")) = Cstr(COD_FIN_ST_ATIVO__ATIVO) then st_ativo=true
	elseif operacao_selecionada=OP_INCLUI then
		st_ativo=true
		end if
%>
		<td width="100%">
		<p class="R">STATUS</p>
		<p class="C">
			<input type="RADIO" id="rb_st_ativo" name="rb_st_ativo" 
				value="<%=COD_FIN_ST_ATIVO__INATIVO%>" 
				class="TA" <%if Not st_ativo then Response.Write(" CHECKED")%>
				><span onclick="fCAD.rb_st_ativo[0].click();" 
				style="cursor:default;color:<%=finStAtivoCor(COD_FIN_ST_ATIVO__INATIVO)%>;"
				><%=finStAtivoDescricao(COD_FIN_ST_ATIVO__INATIVO)%></span
				>&nbsp;</p>
		<p class="C">
			<input type="RADIO" id="rb_st_ativo" name="rb_st_ativo" 
				value="<%=COD_FIN_ST_ATIVO__ATIVO%>" 
				class="TA" <%if st_ativo then Response.Write(" CHECKED")%>
				><span onclick="fCAD.rb_st_ativo[1].click();" 
				style="cursor:default;color:<%=finStAtivoCor(COD_FIN_ST_ATIVO__ATIVO)%>;"
				><%=finStAtivoDescricao(COD_FIN_ST_ATIVO__ATIVO)%></span
				>&nbsp;</p>
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
		s = "<td align='CENTER'>" & chr(13) & _
				"		<div name='dREMOVE' id='dREMOVE'>" & chr(13) & _
					"			<a href='javascript:RemoveRegistro(fCAD)' title='exclui do banco de dados'>" & chr(13) & _
						"				<img src='../botao/remover.gif' width=176 height=55 border=0>" & chr(13) & _
					"			</a>" & chr(13) & _
				"		</div>" & chr(13) & _
			"	</td>" & chr(13)
		end if
	%><%=s%>
	<td align="RIGHT"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaRegistro(fCAD)" title="atualiza o cadastro">
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
	if operacao_selecionada <> OP_INCLUI then
		rs.Close
		set rs = nothing
		end if
	
	cn.Close
	set cn = nothing
%>