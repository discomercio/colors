<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%

'     ========================
'	  MultiCDRegraMenu.asp
'     ========================
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
	
	if Not operacao_permitida(OP_CEN_MULTI_CD_CADASTRO_REGRAS_CONSUMO_ESTOQUE, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________________
' MULTI_CD_REGRA_MONTA_SELECT
'
function multi_cd_regra_monta_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_WMS_REGRA_CD ORDER BY apelido")
	strResp = ""
	do while Not r.Eof
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("apelido"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
	else
		strResp = "<option value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	multi_cd_regra_monta_select = strResp
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fOPConcluir( f ){
var s_dest, s_op, s_id, s_apelido;
	
	s_dest = "";
	s_op = "";
	s_id = "";
	s_apelido = "";
	
	if (f.rb_op[0].checked) {
		s_dest = "MultiCDRegraEdita.asp";
		s_op = OP_INCLUI;
		s_apelido = f.c_novo.value;
		if (trim(s_apelido) == "") {
			alert("Informe o apelido para a nova regra de consumo do estoque!");
			f.c_novo.focus();
			return false;
			}
		}
		
	if (f.rb_op[1].checked) {
		s_dest = "MultiCDRegraEdita.asp";
		s_op = OP_CONSULTA;
		s_id = f.c_cons.value;
		if (trim(s_id) == "") {
			alert("Selecione uma regra para consultar!");
			f.c_cons.focus();
			return false;
		}
		if (converte_numero(s_id) == 0) {
			alert("Regra selecionada para ser consultada possui identificação inválida!");
			f.c_cons.focus();
			return false;
		}
	}

	if (s_dest == "") {
		alert("Escolha uma das opções!");
		return false;
	}
	
	f.id_selecionado.value = s_id;
	f.apelido_selecionado.value = s_apelido;
	f.operacao_selecionada.value=s_op;
	
	window.status = "Aguarde ...";
	f.action = s_dest;
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

<style type="text/css">
.QFn
{
	width: 540px;
}
</style>


<body onload="focus();">

<!--  MENU SUPERIOR -->
<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">CENTRAL&nbsp;&nbsp;ADMINISTRATIVA</span><br>
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
		</span>
	</td>
	</tr>

</table>


<center>
<!--  ***********************************************************************************************  -->
<!--  F O R M U L Á R I O                         												       -->
<!--  ***********************************************************************************************  -->
<form method="post" id="fOP" name="fOP" onsubmit="if (!fOPConcluir(fOP)) return false">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="id_selecionado" id="id_selecionado" value='' />
<input type="hidden" name="apelido_selecionado" id="apelido_selecionado" value='' />
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='' />

<br />
<span class="T">CADASTRO DE REGRAS DO CONSUMO DO ESTOQUE</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td NOWRAP>
			<input type="radio" name="rb_op" value="1" class="CBOX" onclick="fOP.c_novo.focus();"><span style="cursor:default" onclick="fOP.rb_op[0].click(); fOP.c_novo.focus();">Cadastrar Nova</span>&nbsp;
				<input name="c_novo" id="c_novo" type="text" maxlength="30" size="30" onblur="this.value=trim(this.value);" onclick="fOP.rb_op[0].click();" onkeypress="this.click(); filtra_nome_identificador(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit();">
				<br>
			<input type="radio" name="rb_op" value="2" class="CBOX" onclick="fOP.c_cons.focus();"><span style="cursor:default" onclick="fOP.rb_op[1].click(); fOP.c_cons.focus();">Consultar</span>&nbsp;
				<select id="c_cons" name="c_cons" style="margin-left:5px;margin-top:4pt; margin-bottom:4pt;"
					onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" 
					onclick="fOP.rb_op[1].click();">
					<%=multi_cd_regra_monta_select(Null)%>
				</select>
				<br>
			</td>
		</tr>
	</table>

	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bEXECUTAR" id="bEXECUTAR" type="SUBMIT" class="Botao" value="EXECUTAR" title="executa">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>

<br />
<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
	<td align="center"><a href="MenuCadastro.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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