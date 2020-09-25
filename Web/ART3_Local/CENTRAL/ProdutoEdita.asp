<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  ProdutoEdita.asp
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
	dim s, strSql, usuario, fabricante_selecionado, produto_selecionado, operacao_selecionada, s_cor
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	REGISTRO A EDITAR
	fabricante_selecionado = trim(request("fabricante_selecionado"))
	produto_selecionado = trim(request("produto_selecionado"))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if operacao_selecionada=OP_INCLUI then
		fabricante_selecionado=retorna_so_digitos(fabricante_selecionado)
		produto_selecionado=retorna_so_digitos(produto_selecionado)
		end if

	if (fabricante_selecionado="") then Response.Redirect("aviso.asp?id=" & ERR_FABRICANTE_NAO_ESPECIFICADO) 
	if (produto_selecionado="") then Response.Redirect("aviso.asp?id=" & ERR_PRODUTO_NAO_ESPECIFICADO) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)

	fabricante_selecionado = normaliza_codigo(fabricante_selecionado, TAM_MIN_FABRICANTE)
	produto_selecionado = normaliza_codigo(produto_selecionado, TAM_MIN_PRODUTO)

'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	strSql = "SELECT " & _
				"*" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (fabricante = '" & fabricante_selecionado & "')" & _
				" AND (produto = '" & produto_selecionado & "')"
	set rs = cn.Execute(strSql)
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ID_JA_CADASTRADO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_CADASTRADO)
		end if
	
%>


<%=DOCTYPE_LEGADO%>

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


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
    $(function () {
        $("#spn_descricao_html").html($("#c_descricao_html").val());
        $("#spn_descricao").text($("#spn_descricao_html").text());
        $("#c_descricao").text($("#spn_descricao_html").text());

        $("#c_descricao_html").on('input', function () {
            $("#spn_descricao_html").html($("#c_descricao_html").val());
            $("#spn_descricao").text($("#spn_descricao_html").text());
            $("#c_descricao").text($("#spn_descricao_html").text());
        })
    });
</script>

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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
#rb_st_ativo {
	margin: 0pt 4pt 1pt 6pt;
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

<table width="849" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Novo Produto"
	else
		s = "Consulta/Edição de Produto"
		end if
%>
	<td align="center" valign="bottom"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="FinCadContaCorrenteAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='operacao_selecionada_original' id="operacao_selecionada_original" value='<%=operacao_selecionada%>'>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>

<!-- ************   FABRICANTE / PRODUTO   ************ -->
<table width="849" class="Q" cellspacing="0" cellpadding="0">
	<tr>
		<td width="100%">
			<table width="100%" cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td class="MB MD" align="center" width="50%">
						<p class="R">FABRICANTE</p>
						<p class="C">
							<input id="fabricante_selecionado" name="fabricante_selecionado" class="TA" value="<%=fabricante_selecionado%>" readonly size="4" style="text-align:center; color:#0000ff">
						</p>
					</td>
					<td class="MB MD" align="center" width="50%">
						<p class="R">PRODUTO</p>
						<p class="C">
							<input id="produto_selecionado" name="produto_selecionado" class="TA" value="<%=produto_selecionado%>" readonly size="8" style="text-align:center; color:#0000ff">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!--  DESCRIÇÃO HTML (EDIÇÃO)  -->
	<tr>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("descricao_html")) else s=""%>
		<td width="100%" class="MB">
			<p class="R">DESCRIÇÃO (HTML)</p>
			<p class="C">
				<input id="c_descricao_html" name="c_descricao_html" class="TA" type="text" maxlength="4000" style="width:800px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_descricao_html.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
	<!--  DESCRIÇÃO HTML (EXIBIÇÃO)  -->
	<tr>
		<td width="100%" class="MB">
			<p class="R">EXIBIÇÃO DA DESCRIÇÃO (HTML)</p>
			<p class="C">
				<span id="spn_descricao_html" name="spn_descricao_html" class="TA" style="color:#0000ff;font-size:10pt;"></span>
			</p>
		</td>
	</tr>
	<!--  DESCRIÇÃO -->
	<tr>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("descricao")) else s=""%>
		<input type="hidden" name="c_descricao" id="c_descricao" value="<%=s%>" />
		<td width="100%" class="MB">
			<p class="R">DESCRIÇÃO</p>
			<p class="C">
				<span id="spn_descricao" name="spn_descricao" class="TA" style="color:#0000ff;font-size:10pt;"><%=s%></span>
			</p>
		</td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="849" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="849" cellspacing="0">
<tr>
	<td><a href="javascript:history.back()" title="cancela as alterações no cadastro">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		s = "<td align='center'>" & _
				"<div name='dREMOVE' id='dREMOVE'>" & _
					"<a href='javascript:RemoveRegistro(fCAD)' title='exclui do banco de dados'>" & _
						"<img src='../botao/remover.gif' width=176 height=55 border=0>" & _
					"</a>" & _
				"</div>" & _
			"</td>"
		end if
	%><%=s%>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
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