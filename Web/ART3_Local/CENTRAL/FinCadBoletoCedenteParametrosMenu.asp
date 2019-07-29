<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%

'     =======================================
'	  FinCadBoletoCedenteParametrosMenu.asp
'     =======================================
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

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	if Not operacao_permitida(OP_CEN_FIN_CADASTRO_PLANO_CONTAS, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________________________________
' CONTA CORRENTE NAO CONFIGURADA MONTA ITENS SELECT
'
function conta_corrente_nao_configurada_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT " & _
				"*" & _
			" FROM t_FIN_CONTA_CORRENTE" & _
			" WHERE" & _
				" id NOT IN " & _
					"(" & _
						"SELECT" & _
							" id_conta_corrente" & _
						" FROM t_FIN_BOLETO_CEDENTE" & _
					")" & _
			" ORDER BY" & _
				" id"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & Trim("" & r("conta")) & "  -  " & Trim("" & r("descricao"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	conta_corrente_nao_configurada_monta_itens_select = strResp
	r.close
	set r=nothing
end function


' _____________________________________________________________
' CONTA CORRENTE CEDENTE MONTA ITENS SELECT
'
function conta_corrente_cedente_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT " & _
				"*" & _
			" FROM t_FIN_BOLETO_CEDENTE" & _
			" ORDER BY" & _
				" id"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & Trim("" & r("codigo_empresa")) & " - " & Trim("" & r("nome_empresa"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
	
	conta_corrente_cedente_monta_itens_select = strResp
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

<html>


<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fOPConcluir( f ){
var s_dest, s_op, s_id;
	
	s_dest="";
	s_op="";
	s_id="";
	
	if (f.rb_op[0].checked) {
		s_dest="FinCadBoletoCedenteParametrosEdita.asp";
		s_op=OP_INCLUI;
		s_id=f.c_novo.value;
		if (trim(f.c_novo.value)=="") {
			alert("Selecione a conta corrente que deve ser usada como conta do cedente!!");
			f.c_novo.focus();
			return false;
			}
		}
	
	if (f.rb_op[1].checked) {
		s_dest="FinCadBoletoCedenteParametrosEdita.asp";
		s_op=OP_CONSULTA;
		s_id=f.c_cons.value;
		if (trim(f.c_cons.value)=="") {
			alert("Selecione a conta corrente a ser consultada!!");
			f.c_cons.focus();
			return false;
			}
		}
	
	if (f.rb_op[2].checked) {
		s_dest="FinCadBoletoCedenteParametrosLista.asp";
		}

	if (s_dest=="") {
		alert("Escolha uma das opções!!");
		return false;
		}
	
	f.id_selecionado.value=s_id;
	f.operacao_selecionada.value=s_op;
	
	window.status = "Aguarde ...";
	f.action=s_dest;
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">



<body onload="focus()">

<!--  MENU SUPERIOR -->  
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">CENTRAL&nbsp;&nbsp;ADMINISTRATIVA<br>
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


<center>
<!--  ***********************************************************************************************  -->
<!--  F O R M U L Á R I O                         												       -->
<!--  ***********************************************************************************************  -->
<form method="post" id="fOP" name="fOP" onsubmit="if (!fOPConcluir(fOP)) return false">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='id_selecionado' id="id_selecionado" value=''>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value=''>

<span class="T">BOLETO - PARÂMETROS DO CEDENTE</span>
<div class="QFn" align="center">
<table class="TFn">
	<tr>
		<td nowrap>
			<input type="radio" id="rb_op" name="rb_op" value="1" class="CBOX" onclick="fOP.c_novo.focus();">
				<span style="cursor:default" onclick="fOP.rb_op[0].click(); fOP.c_novo.focus();">Configurar nova conta do cedente</span>&nbsp;
				<br>
				<span style="width:22px;">&nbsp;</span>
				<select id="c_novo" name="c_novo" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" onclick="fOP.rb_op[0].click();">
				<% =conta_corrente_nao_configurada_monta_itens_select(Null) %>
				</select>
			<br><br>
			<input type="radio" id="rb_op" name="rb_op" value="2" class="CBOX" onclick="fOP.c_cons.focus();">
				<span style="cursor:default" onclick="fOP.rb_op[1].click(); fOP.c_cons.focus();">Consultar conta já configurada</span>&nbsp;
				<br>
				<span style="width:22px;">&nbsp;</span>
				<select id="c_cons" name="c_cons" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" onclick="fOP.rb_op[1].click();">
				<% =conta_corrente_cedente_monta_itens_select(Null) %>
				</select>
			<br><br>
			<input type="radio" id="rb_op" name="rb_op" value="3" class="CBOX">
				<span class="rbLink" onclick="fOP.rb_op[2].click(); fOP.bEXECUTAR.click();">Consultar lista</span>
			</td>
		</tr>
	</table>

	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" value="EXECUTAR" title="executa">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>

<p class="TracoBottom"></p>

<table cellspacing="0">
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