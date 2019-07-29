<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  EquipeVendasEdita.asp
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
	dim s, strSql, usuario, apelido_selecionado, operacao_selecionada
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	REGISTRO A EDITAR
	apelido_selecionado = filtra_nome_identificador(ucase(trim(request("id_selecionado"))))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if (apelido_selecionado="") then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_INFORMADO) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)

'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim intIdEquipeVendas, intIndice
	intIdEquipeVendas = 0
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_EQUIPE_VENDAS" & _
			" WHERE" & _
				" (apelido = '" & apelido_selecionado & "')"
	set rs = cn.Execute(strSql)
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICADOR_JA_CADASTRADO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICADOR_NAO_CADASTRADO)
		intIdEquipeVendas = rs("id")
		end if




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________________
' SUPERVISOR_MONTA_ITENS_SELECT
'
function supervisor_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT" & _
				" usuario," & _
				" nome_iniciais_em_maiusculas" & _
			" FROM t_USUARIO" & _
			" WHERE" & _
				" (bloqueado = 0)" & _
			" ORDER BY" & _
				" usuario"
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
		
	supervisor_monta_itens_select = strResp
	r.close
	set r=nothing
end function

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

function AtualizaRegistro( f ) {
	if (trim(f.c_descricao.value)=="") {
		alert('Preencha a descrição!!');
		f.c_descricao.focus();
		return;
		}
	if (trim(f.c_supervisor.value)=="") {
		alert('Selecione o supervisor!!');
		f.c_supervisor.focus();
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
#c_supervisor {
	margin: 4pt 4pt 4pt 10pt;
	vertical-align: top;
	}
</style>


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.c_descricao.focus()"
	else
		s = "focus()"
		end if
%>
<body onload="<%=s%>">
<center>



<!--  FORMULÁRIO DE CADASTRO  -->

<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Nova Equipe de Vendas"
	else
		s = "Consulta/Edição de Equipe de Vendas"
		end if
%>
	<td align="center" valign="bottom"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" METHOD="POST" ACTION="EquipeVendasAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<INPUT type=HIDDEN name='operacao_selecionada_original' id="operacao_selecionada_original" value='<%=operacao_selecionada%>'>
<INPUT type=HIDDEN name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>
<!-- FORÇA A CRIAÇÃO DE UM ARRAY MESMO QUANDO NÃO HÁ ITENS NA LISTA DE MEMBROS -->
<input type="HIDDEN" id="chk_membros" name="chk_membros" value=''>
<input type="HIDDEN" id="chk_membros" name="chk_membros" value=''>


<!-- ************   APELIDO / DESCRIÇÃO   ************ -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
		<td class="MD" align="center" width="15%">
			<p class="R">ID</p>
			<p class="C">
				<input id="id_selecionado" name="id_selecionado" class="TA" value="<%=apelido_selecionado%>" READONLY size="14" style="text-align:center; color:#0000ff">
			</p>
		</td>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("descricao")) else s=""%>
		<td width="85%">
			<p class="R">DESCRIÇÃO</p>
			<p class="C">
				<input id="c_descricao" name="c_descricao" class="TA" type="TEXT" maxlength="40" size="60" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) bATUALIZA.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   SUPERVISOR   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%		dim s_supervisor
		s_supervisor = ""
		if operacao_selecionada=OP_CONSULTA then s_supervisor = Trim("" & rs("supervisor"))
%>
		<td width="100%">
		<span class="R">SUPERVISOR</span>
		<br>
		<select id="c_supervisor" name="c_supervisor" style="width:490px;">
		<% =supervisor_monta_itens_select(s_supervisor) %>
		</select>
		</td>
	</tr>
</table>

<!-- ************   MEMBROS DA EQUIPE   ************ -->
<%
	strSql = "SELECT " & _
				"*" & _
			"FROM (" & _
					"SELECT" & _
						" 'S' AS assinalado," & _
						" tU.usuario," & _
						" tU.nome," & _
						" tU.nome_iniciais_em_maiusculas" & _
					" FROM t_USUARIO tU INNER JOIN t_EQUIPE_VENDAS_X_USUARIO tEVU ON (tU.usuario=tEVU.usuario)" & _
					" WHERE" & _
						" (id_equipe_vendas = " & intIdEquipeVendas & ")" & _
					" UNION " & _
					"SELECT" & _
						" 'N' AS assinalado," & _
						" tU.usuario," & _
						" tU.nome," & _
						" tU.nome_iniciais_em_maiusculas" & _
					" FROM t_USUARIO tU" & _
					" WHERE" & _
						" (bloqueado = 0)" & _
						" AND (usuario NOT IN (SELECT DISTINCT usuario FROM t_EQUIPE_VENDAS_X_USUARIO))" & _
				  ") t" & _
			" ORDER BY" & _
				" usuario"
	set rs = cn.Execute(strSql)
	if Not rs.EOF then
%>
		<br>
		<table width="649" class="Q" cellSpacing="0">
			<tr>
				<td colspan="2" width="100%">
					<span class="R">MEMBROS DA EQUIPE</span>
				</td>
			</tr>
<%
		intIndice = 0
	'	Devido aos 2 campos Hidden existentes p/ forçar a criação do array
		intIndice = intIndice + 2
		do while Not rs.EOF
%>
			<tr>
				<td>
					<input type="CHECKBOX" id="chk_membros" name="chk_membros" value="<%=Trim("" & rs("usuario"))%>"
						<%if Trim("" & rs("assinalado")) = "S" then Response.Write " CHECKED"%>>
				</td>
				<td width="99%">
					<span class="C" onclick="fCAD.chk_membros[<%=intIndice%>].click();" style="cursor:default;"><%=Trim("" & rs("usuario")) & " - " & Trim("" & rs("nome_iniciais_em_maiusculas"))%></span>
				</td>
			</tr>
<%			
			intIndice = intIndice + 1
			rs.MoveNext
			loop
%>
		</table>
<%
		end if
%>


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
		s = "<td align='CENTER'>" & _
				"<div name='dREMOVE' id='dREMOVE'>" & _
					"<a href='javascript:RemoveRegistro(fCAD)' title='exclui do banco de dados'>" & _
						"<img src='../botao/remover.gif' width=176 height=55 border=0>" & _
					"</a>" & _
				"</div>" & _
			"</td>"
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
	rs.Close
	set rs = nothing
	
	cn.Close
	set cn = nothing
%>