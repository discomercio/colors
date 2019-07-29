<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  G R U P O L O J A S E D I T A . A S P
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
	dim s, usuario, grupo_lojas_selecionado, operacao_selecionada
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	GRUPO DE LOJAS A EDITAR
	grupo_lojas_selecionado = trim(request("grupo_lojas_selecionado"))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if operacao_selecionada=OP_INCLUI then
		grupo_lojas_selecionado=retorna_so_digitos(grupo_lojas_selecionado)
		end if

	grupo_lojas_selecionado=normaliza_codigo(grupo_lojas_selecionado, TAM_MIN_GRUPO_LOJAS)
	
	if (grupo_lojas_selecionado="") Or (grupo_lojas_selecionado="00") then Response.Redirect("aviso.asp?id=" & ERR_GRUPO_LOJAS_NAO_ESPECIFICADO) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	
'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	set rs = cn.Execute("select * from t_LOJA_GRUPO where (grupo='" & grupo_lojas_selecionado & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_GRUPO_LOJAS_JA_CADASTRADO)
	'	GARANTE QUE O Nº DO GRUPO DE LOJAS NÃO ESTÁ EM USO
		rs.Close
		set rs = cn.Execute("select * from t_LOJA_GRUPO where (CONVERT(smallint,grupo) = " & grupo_lojas_selecionado & ")")
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_GRUPO_LOJAS_JA_CADASTRADO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_GRUPO_LOJAS_NAO_CADASTRADO)
		end if
	
	
	
	

	
' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' MONTA RELACAO LOJAS CADASTRADAS
'
sub monta_relacao_lojas_cadastradas
const COR_ITEM = "navy"
dim r, s, x
	
	s = "SELECT t_LOJA_GRUPO.grupo, t_LOJA_GRUPO_ITEM.loja," & _
		" t_LOJA_GRUPO.descricao, t_LOJA.nome, t_LOJA.razao_social" & _
		" FROM t_LOJA_GRUPO INNER JOIN t_LOJA_GRUPO_ITEM ON (t_LOJA_GRUPO.grupo=t_LOJA_GRUPO_ITEM.grupo)" & _
		" LEFT JOIN t_LOJA ON (t_LOJA_GRUPO_ITEM.loja=t_LOJA.loja)" & _
		" WHERE (t_LOJA_GRUPO.grupo='" & grupo_lojas_selecionado & "')" & _
		" ORDER BY CONVERT(smallint,t_LOJA_GRUPO.grupo), CONVERT(smallint,t_LOJA_GRUPO_ITEM.loja)"

	x = ""
	set r = cn.execute(s)
	do while Not r.Eof
		x = x & "	<TR>" & chr(13)
	
	'	IDENTAÇÃO
		x = x & "		<TD style='width:20px;'>&nbsp;</td>" & chr(13)

	'	NÚMERO DA LOJA
		x = x & "		<TD align='right'><p class='C' style='color:" & COR_ITEM & "'>" & Trim("" & r("loja")) & "</p></td>" & chr(13)

	'	TRAÇO
		x = x & "		<TD><p class='C' style='color:" & COR_ITEM & "'>-</p></td>" & chr(13)
	
	'	NOME DA LOJA
		s = Trim("" & r("nome"))
		if s = "" then s = Trim("" & r("razao_social"))
		s = iniciais_em_maiusculas(s)
		if s = "" then s = "&nbsp;"
		x = x & "		<TD><p class='C' style='color:" & COR_ITEM & "'>" & s & "</p></td>" & chr(13)

		x = x & "	</TR>" & chr(13)
		r.movenext
		loop

	if x <> "" then
		x = "<TABLE cellSpacing='0' cellPadding='0'>" & chr(13) & _
			x & _
			"</TABLE>" & chr(13)
		end if

	if x = "" then x = "&nbsp;"
	Response.Write x

end sub



' _____________________________________
' MONTA LISTA LOJAS INCLUSAO
'
sub monta_lista_lojas_inclusao
const COR_ITEM = "black"
dim r, s, x, n_reg
	
'	SELECIONA APENAS AS LOJAS QUE AINDA NÃO PERTENÇAM A NENHUM GRUPO
	s = "SELECT t_LOJA.loja, t_LOJA.nome, t_LOJA.razao_social" & _
		" FROM t_LOJA" & _
		" WHERE (" & _
					"loja NOT IN" & _
					" (SELECT DISTINCT loja FROM t_LOJA_GRUPO_ITEM)" & _
				")" & _
		" ORDER BY CONVERT(smallint,loja)"

	x = ""
	n_reg = 0
	set r = cn.execute(s)
	do while Not r.Eof
		n_reg = n_reg + 1
		
		x = x & "	<TR>" & chr(13)

	'	IDENTAÇÃO
		x = x & "		<TD style='width:20px;'>&nbsp;</td>" & chr(13)

	'	CHECKBOX
		x = x & "		<TD><input type='checkbox' tabindex='-1' id='ckb_inclui' name='ckb_inclui'" & _
				" value='" & Trim("" & r("loja")) & "'></TD>" & chr(13)
	
	'	NÚMERO DA LOJA
		x = x & "		<TD align='right' valign='bottom'><p class='C' style='color:" & COR_ITEM & "'>" & _
				Trim("" & r("loja")) & "</p></td>" & chr(13)
	
	'	TRAÇO
		x = x & "		<TD valign='bottom'><p class='C' style='color:" & COR_ITEM & "'>-</p></td>" & chr(13)

	'	NOME DA LOJA
		s = Trim("" & r("nome"))
		if s = "" then s = Trim("" & r("razao_social"))
		s = iniciais_em_maiusculas(s)
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='bottom'><p class='C' style='color:" & COR_ITEM & "'>" & _
				s & "</p></td>" & chr(13)

		x = x & "	</TR>" & chr(13)
		r.movenext
		loop


	if x <> "" then
		x = "<TABLE cellSpacing='0' cellPadding='0'>" & chr(13) & _
			x & _
			"</TABLE>" & chr(13)
		end if

	if x = "" then x = "&nbsp;"
	Response.Write x

end sub



' _____________________________________
' MONTA LISTA LOJAS EXCLUSAO
'
sub monta_lista_lojas_exclusao
const COR_ITEM = "darkred"
dim r, s, x, n_reg
	
	s = "SELECT t_LOJA_GRUPO_ITEM.grupo, t_LOJA_GRUPO_ITEM.loja," & _
		" t_LOJA.nome, t_LOJA.razao_social" & _
		" FROM t_LOJA_GRUPO_ITEM LEFT JOIN t_LOJA ON (t_LOJA_GRUPO_ITEM.loja=t_LOJA.loja)" & _
		" WHERE (t_LOJA_GRUPO_ITEM.grupo = '" & grupo_lojas_selecionado & "')" & _
		" ORDER BY CONVERT(smallint,t_LOJA_GRUPO_ITEM.loja)"

	x = ""
	n_reg = 0
	set r = cn.execute(s)
	do while Not r.Eof
		n_reg = n_reg + 1
		
		x = x & "	<TR>" & chr(13)

	'	IDENTAÇÃO
		x = x & "		<TD style='width:20px;'>&nbsp;</td>" & chr(13)
		
	'	CHECKBOX
		x = x & "		<TD><input type='checkbox' tabindex='-1' id='ckb_exclui' name='ckb_exclui'" & _
				" value='" & Trim("" & r("loja")) & "'></TD>" & chr(13)
	
	'	NÚMERO DA LOJA
		x = x & "		<TD align='right' valign='bottom'><p class='C' style='color:" & COR_ITEM & "'>" & _
				Trim("" & r("loja")) & "</p></td>" & chr(13)
	
	'	TRAÇO
		x = x & "		<TD valign='bottom'><p class='C' style='color:" & COR_ITEM & "'>-</p></td>" & chr(13)

	'	NOME DA LOJA
		s = Trim("" & r("nome"))
		if s = "" then s = Trim("" & r("razao_social"))
		s = iniciais_em_maiusculas(s)
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='bottom'><p class='C' style='color:" & COR_ITEM & "'>" & _
				s & "</p></td>" & chr(13)

		x = x & "	</TR>" & chr(13)
		r.movenext
		loop


	if x <> "" then
		x = "<TABLE cellSpacing='0' cellPadding='0'>" & chr(13) & _
			x & _
			"</TABLE>" & chr(13)
		end if

	if x = "" then x = "&nbsp;"
	Response.Write x

end sub


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
function RemoveGrupoLojas( f ) {
var b,s;
	s = "Confirma a exclusão do Grupo de Lojas?\n";
	s=s+"=================================";
	s=s+"\n\n";
	s=s+"Observação:\n";
	s=s+"Para excluir uma loja deste grupo, selecione-a na lista de lojas disponíveis para exclusão e clique no botão CONFIRMAR";
	b=window.confirm(s);
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaGrupoLojas( f ) {
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


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.c_descricao.focus()"
	else
		s = "focus()"
		end if
%>
<body onLoad="<%=s%>">
<center>



<!--  CADASTRO DO GRUPO DE LOJAS -->

<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Novo Grupo de Lojas"
	else
		s = "Consulta/Edição de Grupo de Lojas Cadastrado"
		end if
%>
	<td align="CENTER" vAlign="BOTTOM"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="GrupoLojasAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<INPUT type=hidden name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>

<!-- ************   NÚMERO/DESCRIÇÃO   ************ -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
		<td class="MD" style='width:115px;'><p class="R">GRUPO DE LOJAS</p><p class="C"><input id="grupo_lojas_selecionado" name="grupo_lojas_selecionado" class="TA" value="<%=grupo_lojas_selecionado%>" readonly tabindex=-1 size="6" style="text-align:center; color:#0000ff"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("descricao")) else s=""%>
		<td><p class="R">DESCRIÇÃO</p><p class="C"><input id="c_descricao" name="c_descricao" class="TA" type="TEXT" maxlength="30" size="60" value="<%=s%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bATUALIZA.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>


<!-- ************   RELAÇÃO DE LOJAS DO GRUPO   ************ -->
<%if operacao_selecionada=OP_CONSULTA then%>
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td width="100%"><p class="R">RELAÇÃO DE LOJAS</p>
<%monta_relacao_lojas_cadastradas%>
		</td>
	</tr>
</table>
<%end if%>


<!-- ************   LISTA DE LOJAS PARA INCLUSÃO   ************ -->
<br>
<table width="649" cellSpacing="0">
	<tr>
		<td width="100%">
			<p class="R">INCLUSÃO DE LOJAS NO GRUPO</p>
		</td>
	</tr>
	<tr>
		<td class="MT" width="100%">
<%monta_lista_lojas_inclusao%>
		</td>
	</tr>
</table>


<!-- ************   LISTA DE LOJAS PARA EXCLUSÃO   ************ -->
<%if operacao_selecionada<>OP_INCLUI then%>
<br>
<table width="649" cellSpacing="0">
	<tr>
		<td width="100%">
			<p class="R">EXCLUSÃO DE LOJAS DO GRUPO</p>
		</td>
	</tr>
	<tr>
		<td class="MT" width="100%">
<%monta_lista_lojas_exclusao%>
		</td>
	</tr>
</table>
<%end if%>


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
		s = "<td align='CENTER'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveGrupoLojas(fCAD)' "
		s =s + "title='remove o grupo de lojas cadastrado'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="RIGHT"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaGrupoLojas(fCAD)" title="atualiza o cadastro do grupo de lojas">
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