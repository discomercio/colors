<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================================
'	  RelPedidoPreDevolucaoMercadoriaRecebe.asp
'     =======================================================
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


	On Error GoTo 0
	Err.Clear

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_PRE_DEVOLUCAO_RECEBIMENTO_MERCADORIA, s_lista_operacoes_permitidas) then ' TODO
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if
		
	dim intIdx
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
	<title>CENTRAL</title>
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_GLOBAL%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
<%=monta_funcao_js_normaliza_numero_pedido_e_sufixo%>

function fFILTROConfirma(f) {
var i, blnFlag;
	blnFlag = false;
	for (i = 0; i < f.rb_status.length; i++) {
		if (f.rb_status[i].checked) {
			blnFlag = true;
			break;
		}
	}
	if (!blnFlag) {
		alert("Selecione o status da pré-devolução!!");
		return;
	}
	
	dCONFIRMA.style.visibility="hidden";
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


<body onload="focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelPedidoPreDevolucaoMercadoriaRecebeExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Registrar Mercadoria Recebida</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0" style="width:240px;">
<!--  STATUS DA PRÉ-DEVOLUÇÃO  -->
	<tr>
		<td class="MT PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;STATUS DA PRÉ-DEVOLUÇÃO</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD" style="padding:5px;">
			<% intIdx=-1 %>
			<input type="radio" id="rb_status" name="rb_status" value="EM_ANDAMENTO" class="CBOX" style="margin-left:20px;" checked>
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Em Andamento</span>
			<br>
			<input type="radio" id="rb_status" name="rb_status" value="MERCADORIA_RECEBIDA" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Mercadoria Recebida</span>
		</td>
	</tr>

<!-- PEDIDO -->
    <tr>
		<td class="MB MD ME PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;PEDIDO</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD" style="padding:5px;">
		<input name="c_pedido" id="c_pedido" class="PLLe" maxlength="10" style="width:100px;margin-left:4px;font-size:10pt;"
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) $(this).hUtil('focusNext'); filtra_pedido();"
				onblur="if (normaliza_numero_pedido_e_sufixo(this.value)!='') {this.value=normaliza_numero_pedido_e_sufixo(this.value);}">
		</td>
	</tr>

<!-- NOTA FISCAL -->
    <tr>
		<td class="MB MD ME PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;Nº NOTA FISCAL</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD" style="padding:5px;">
		<input name="c_nota_fiscal" id="c_nota_fiscal" class="PLLe" maxlength="10" style="width:100px;margin-left:4px;font-size:10pt;"
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) $(this).hUtil('focusNext'); filtra_numerico();">
		</td>
	</tr>

<!-- CPF/CNPJ CLIENTE -->
    <tr>
		<td class="MB MD ME PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;CPF/CNPJ CLIENTE</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD" style="padding:5px;">
		<input name="c_cpf_cnpj" id="c_cpf_cnpj" class="PLLe" maxlength="18" style="width:180px;margin-left:4px;font-size:10pt;"
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) $(this).hUtil('focusNext'); filtra_numerico();"
                onblur="if (retorna_so_digitos(this.value).length==14) { this.value=cnpj_formata(this.value);} else if (retorna_so_digitos(this.value).length==11){ this.value=cpf_formata(this.value);} else alert('Formato de CPF/CNPJ inválido!');">
		</td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="executa a consulta">
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
	cn.Close
	set cn = nothing
%>
