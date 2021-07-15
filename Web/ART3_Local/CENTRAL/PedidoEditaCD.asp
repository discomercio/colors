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
'	  PedidoEditaCD.asp
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

	pedido_selecionado = Trim(Request.Form("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

    dim url_origem
    url_origem = Trim(Request("url_origem"))

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_EDITA_PEDIDO_CD, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim cn, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim alerta
	alerta=""

	dim r_pedido
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then
		alerta = msg_erro
		end if

	if r_pedido.st_entrega <> ST_ENTREGA_ESPERAR then
		alerta=texto_add_br(alerta)
		alerta=alerta & "O status do pedido � inv�lido para realizar esta opera��o (" & x_status_entrega(r_pedido.st_entrega) & ")"
		end if

	dim sCnpjMustBeEqual
	sCnpjMustBeEqual = ""

	dim rNfeEmitente
	if alerta = "" then
		set rNfeEmitente = le_nfe_emitente(r_pedido.id_nfe_emitente)
		'Se houver NFe emitida para o pedido, somente permite alterar para outro CD que possua o mesmo CNPJ
		if (Trim("" & r_pedido.obs_2) <> "") Or (Trim("" & r_pedido.obs_3) <> "") Or (Trim("" & r_pedido.obs_4) <> "") then
			sCnpjMustBeEqual = rNfeEmitente.cnpj
			end if
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  �  �  E  S 
' _____________________________________________________________________________________________

' _____________________________________________________________
' CD monta itens select
' Usado para montar a lista de op��es de novo CD para o pedido.
' A lista � montada de forma que:
'	1) Somente os CDs habilitados no perfil do usu�rio s�o exibidos
'	2) O CD que j� est� cadastrado no pedido n�o � exibido na lista
'	3) Se o pedido j� possui NFe emitida, s�o exibidos como op��es de novo CD somente aqueles que possuem o mesmo CNPJ que o CD atual
'
' Subs�dios:
' Cada empresa emitente de NFe que esteja com os campos 'st_ativo'
' e 'st_habilitado_ctrl_estoque' habilitados (valor 1) deve ser
' considerado como um CD (Centro de Distribui��o).
'
' Par�metros:
'	id_usuario = ID do usu�rio, esta informa��o � usada para determinar os CDs habilitados em seu perfil
'	id_nfe_emitente_ignorado = ID do emitente que deve ser exclu�do da lista de op��es, ou seja, o CD que j� consta no pedido
'	cnpjMustBeEqual = se informado, s�o exibidos somente os CDs que possuam este CNPJ
function CD_monta_itens_select(ByVal id_usuario, ByVal id_nfe_emitente_ignorado, ByVal cnpjMustBeEqual)
dim x, r, strSql, strResp
	strSql = "SELECT" & _
				" tNE.id," & _
				" tNE.apelido," & _
				" tNE.razao_social" & _
			" FROM t_NFe_EMITENTE tNE" & _
				" INNER JOIN t_USUARIO_X_NFe_EMITENTE tUXNE ON (tUXNE.id_nfe_emitente = tNE.id)" & _
			" WHERE" & _
				" (tNE.st_ativo <> 0)" & _
				" AND (tNE.st_habilitado_ctrl_estoque <> 0)" & _
				" AND (tUXNE.usuario = '" & id_usuario & "')" & _
				" AND (tUXNE.excluido_status = 0)"

	if Trim("" & cnpjMustBeEqual) <> "" then
		strSql = strSql & _
				" AND (tNE.cnpj = '" & retorna_so_digitos(Trim("" & cnpjMustBeEqual)) & "')"
		end if

	if Trim("" & id_nfe_emitente_ignorado) <> "" then
		if converte_numero(id_nfe_emitente_ignorado) > 0 then
			strSql = strSql & _
					" AND (tNE.id <> " & Trim("" & id_nfe_emitente_ignorado) & ")"
			end if
		end if

	strSql = strSql & _
			" ORDER BY" & _
				" id"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		strResp = strResp & "<option"
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("apelido"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
	
'	SE N�O H� NENHUM ITEM DEFAULT, INCLUI UM ITEM EM BRANCO P/ SER O DEFAULT
	strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp

	CD_monta_itens_select = strResp
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
	<title>CENTRAL<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	function fPedEditaCDConfirma(f) {
		if (trim(f.c_novo_CD.value) == "") {
			alert("Selecione o novo CD!");
			f.c_novo_CD.focus();
			return;
		}

		dCONFIRMA.style.visibility = "hidden";
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
<link href="<%=URL_FILE__ESCREEN_CSS%>" Rel="stylesheet" Type="text/css" media="screen">



<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<body>
<center>

<form id="fPED" name="fPED" method="post" action="PedidoEditaCDConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="url_origem" id="url_origem" value="<%=url_origem%>" />

<!--  I D E N T I F I C A � � O   D O   P E D I D O -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="left" valign="bottom"><span class="PEDIDO">Edi��o do CD</span></td>
	<td align="right" valign="bottom"><span class="PEDIDO" style="font-size:14pt;">Pedido <%=pedido_selecionado%></span></td>
</tr>
</table>

<br />

<table class="Q" style="width:649px;" cellSpacing="0">
	<tr>
		<td colspan="2">
			<p class="Rf">CD ATUAL</p>
			<p class="C" style="font-size:10pt; margin: 4px 0px 2px 4px;"><%=rNfeEmitente.apelido%></p>
		</td>
	</tr>
	<tr>
		<td align="left" valign="bottom" colspan="2" class="MC">
			<p class="Rf">CD NOVO</p>
			<p class="C">
			<select name="c_novo_CD" id="c_novo_CD" style="width:180px;margin:4px 0px 8px 4px;">
			<%=CD_monta_itens_select(usuario, r_pedido.id_nfe_emitente, sCnpjMustBeEqual)%>
			</select>
			</p>
		</td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br />


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a name="bCANCELA" id="bCANCELA" href="javascript:history.back()" title="cancela a edi��o do CD">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPedEditaCDConfirma(fPED)" title="grava a altera��o do CD">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>