<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  P E D I D O B L O C O N O T A S N O V O . A S P
'     ===============================================
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

	dim usuario, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    dim url_origem
    url_origem = Trim(Request("url_origem"))

	if Not operacao_permitida(OP_LJA_BLOCO_NOTAS_PEDIDO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if
		
	dim nivel_acesso_bloco_notas
	nivel_acesso_bloco_notas = Session("nivel_acesso_bloco_notas")
	if Trim(nivel_acesso_bloco_notas) = "" then
		nivel_acesso_bloco_notas = obtem_nivel_acesso_bloco_notas_pedido(cn, usuario)
		Session("nivel_acesso_bloco_notas") = nivel_acesso_bloco_notas
		end if

	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
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
	<title>LOJA<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function calcula_tamanho_restante() {
	var f, s;
	f = fPED;
	s = "" + fPED.c_mensagem.value;
	f.c_tamanho_restante.value = MAX_TAM_MENSAGEM_BLOCO_NOTAS - s.length;
}

function fPEDBlocoNotasNovoConfirma(f) {
var s;

	s = "" + f.c_mensagem.value;
	if (s.length == 0) {
		alert('É necessário escrever o texto da mensagem!!');
		f.c_mensagem.focus();
		return;
		}
	if (s.length > MAX_TAM_MENSAGEM_BLOCO_NOTAS) {
		alert('Conteúdo da mensagem excede em ' + (s.length - MAX_TAM_MENSAGEM_BLOCO_NOTAS) + ' caracteres o tamanho máximo de ' + MAX_TAM_MENSAGEM_BLOCO_NOTAS + '!!');
		f.c_mensagem.focus();
		return;
		}

	if (trim(f.c_nivel_acesso_bloco_notas.value) == "") {
		alert('É necessário definir o nível de acesso para a leitura da mensagem!!');
		f.c_nivel_acesso_bloco_notas.focus();
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<body onload="fPED.c_mensagem.focus();">
<center>

<form id="fPED" name="fPED" method="post" action="PedidoBlocoNotasNovoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="url_origem" id="url_origem" value="<%=url_origem%>" />

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->  
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="left" valign="bottom"><p class="PEDIDO">Bloco de Notas</p></td>
	<td align="right" valign="bottom"><p class="PEDIDO" style="font-size:14pt;">Pedido <%=pedido_selecionado%></p></td>
</tr>
</table>
<br>

<table>
<tr>
	<td align="right" valign="bottom">
		<span class="PLLd">Tamanho restante:</span><input name="c_tamanho_restante" id="c_tamanho_restante" tabindex=-1 readonly class="TA" style="width:35px;text-align:right;" value='<%=Cstr(MAX_TAM_MENSAGEM_BLOCO_NOTAS)%>' />
	</td>
</tr>
<tr>
	<td>
	<table class="Q" style="width:649px;" cellSpacing="0">
		<tr>
			<td><p class="Rf">MENSAGEM</p>
				<textarea name="c_mensagem" id="c_mensagem" class="PLLe" rows="<%=Cstr(MAX_LINHAS_MENSAGEM_BLOCO_NOTAS)%>" 
					style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_MENSAGEM_BLOCO_NOTAS);" onblur="this.value=trim(this.value);calcula_tamanho_restante();"
					onkeyup="calcula_tamanho_restante();"
					></textarea>
			</td>
		</tr>
	</table>
	</td>
</tr>
<% if converte_numero(nivel_acesso_bloco_notas) = converte_numero(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO) then %>
<input type="hidden" name="c_nivel_acesso_bloco_notas" id="c_nivel_acesso_bloco_notas" value='<%=Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO)%>'>
<% else %>
<tr>
	<td>
		<br />
		<p class="Rf">NÍVEL DE ACESSO PARA LEITURA</p>
		<select id="c_nivel_acesso_bloco_notas" name="c_nivel_acesso_bloco_notas" style="margin-top:3px;margin-bottom:4px;width:180px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}">
		<% =nivel_acesso_bloco_notas_pedido_monta_itens_select(Null, nivel_acesso_bloco_notas) %>
		</select>
	</td>
</tr>
<% end if %>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a name="bCANCELA" id="bCANCELA" href="javascript:history.back()" title="cancela o cadastramento de nova mensagem no bloco de notas">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="RIGHT"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDBlocoNotasNovoConfirma(fPED)" title="grava a mensagem no bloco de notas">
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