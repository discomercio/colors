<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================================
'	  P E D I D O C H A M A D O F I N A L I Z A . A S P
'     ===============================================================
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

	dim usuario, pedido_selecionado, id_chamado, id_depto
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
	id_chamado = Trim(request("id_chamado"))
	if (id_chamado = "") then Response.Redirect("aviso.asp?id=" & ERR_PARAMETRO_OBRIGATORIO_NAO_ESPECIFICADO)

    id_depto = Trim(request("id_depto"))
	if (id_depto = "") then Response.Redirect("aviso.asp?id=" & ERR_PARAMETRO_OBRIGATORIO_NAO_ESPECIFICADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim blnIsUsuarioResponsavelDepto, blnIsUsuarioCadastroChamado
    blnIsUsuarioCadastroChamado = CBool(Request.Form("blnIsUsuarioCadastroChamado"))
    blnIsUsuarioResponsavelDepto = CBool(Request.Form("blnIsUsuarioResponsavelDepto"))

	if Not operacao_permitida(OP_CEN_REL_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) And _
    Not blnIsUsuarioResponsavelDepto And _
    Not blnIsUsuarioCadastroChamado then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

    dim nivel_acesso_chamado
	nivel_acesso_chamado = Session("nivel_acesso_chamado")
	if Trim(nivel_acesso_chamado) = "" then
		nivel_acesso_chamado = obtem_nivel_acesso_chamado_pedido(cn, usuario)
		Session("nivel_acesso_chamado") = nivel_acesso_chamado
		end if

' _____________________________________________
' MOTIVO_FINALIZACAO_CHAMADO_MONTA_ITENS_SELECT

function motivo_finalizacao_chamado_monta_itens_select(byval depto)
dim x, r, strResp, strSql
    strSql = "SELECT * FROM t_CODIGO_DESCRICAO" & _
                " WHERE grupo='" & GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_FINALIZACAO & "' AND codigo_pai = '" & Cstr(depto) & "' AND st_inativo=0" & _
                " ORDER BY ordenacao"

    set r = cn.Execute(strSql)
	strResp = "<option value='' selected>&nbsp;</option>"
	do while Not r.EOF 
        x = r("codigo")
        strResp = strResp & "<option"
	    strResp = strResp & " value='" & x & "'>"
        strResp = strResp & r("descricao")
        strResp = strResp & "</option>"
		r.MoveNext        
    loop

    motivo_finalizacao_chamado_monta_itens_select = strResp
	r.Close
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
	<title>CENTRAL<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function calcula_tamanho_restante() {
var f, s;
	f = fPED;
	s = "" + fPED.c_texto.value;
	f.c_tamanho_restante.value = MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS - s.length;
}

function fPEDChamadoFinalizaConfirma(f) {
var s;

	s = "" + f.c_texto.value;
	if (s.length == 0) {
		alert('É necessário escrever a solução referente a finalização deste chamado!!');
		f.c_texto.focus();
		return;
		}
	if (s.length > MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS) {
	    alert('Conteúdo do texto excede em ' + (s.length - MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS) + ' caracteres o tamanho máximo de ' + MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS + '!!');
		f.c_texto.focus();
		return;
	}

	s = "" + f.c_motivo_finalizacao.value;
	if (s.length == 0) {
	    alert('É necessário selecionar o motivo da finalização deste chamado!!');
	    f.c_motivo_finalizacao.focus();
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

<body onload="fPED.c_motivo_finalizacao.focus();">
<center>

<form id="fPED" name="fPED" method="post" action="PedidoChamadoFinalizaConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="id_chamado" id="id_chamado" value='<%=id_chamado%>'>
<input type="hidden" name="blnIsUsuarioCadastroChamado" id="blnIsUsuarioCadastroChamado" value="<%=blnIsUsuarioCadastroChamado%>" />
<input type="hidden" name="blnIsUsuarioResponsavelDepto" id="blnIsUsuarioResponsavelDepto" value="<%=blnIsUsuarioResponsavelDepto%>" />

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="left" valign="bottom"><p class="PEDIDO">Finalizar Chamado do Pedido <%=pedido_selecionado%></p></td>
</tr>
</table>
<br>

<table>
<tr>
    <td>
		<br />
		<p class="Rf">MOTIVO DA FINALIZAÇÃO</p>
		<select id="c_motivo_finalizacao" name="c_motivo_finalizacao" style="margin-top:3px;margin-bottom:4px;width:300px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}">
		<% =motivo_finalizacao_chamado_monta_itens_select(id_depto) %>
		</select>
	</td>
	
</tr>
<tr>
<tr>
	<td align="right" valign="bottom">
		<span class="PLLd">Tamanho restante:</span><input name="c_tamanho_restante" id="c_tamanho_restante" tabindex=-1 readonly class="TA" style="width:35px;text-align:right;" value='<%=Cstr(MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS)%>' />
	</td>
</tr>
	<td>
	<table class="Q" style="width:649px;" cellSpacing="0">
		<tr>
			<td><p class="Rf">SOLUÇÃO</p>
				<textarea name="c_texto" id="c_texto" class="PLLe" rows="7" 
					style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS);" onblur="this.value=trim(this.value);calcula_tamanho_restante();"
					onkeyup="calcula_tamanho_restante();"
					></textarea>
			</td>
		</tr>
	</table>
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
	<td><a name="bCANCELA" id="bCANCELA" href="javascript:history.back()" title="cancela a finalização do chamado">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDChamadoFinalizaConfirma(fPED)" title="confirma a finalização do chamado">
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