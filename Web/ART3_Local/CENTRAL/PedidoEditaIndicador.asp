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
'	  PedidoEditaIndicador.asp
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

	if Not operacao_permitida(OP_CEN_EDITA_PEDIDO_INDICADOR, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim alerta
	alerta=""

	dim r_pedido
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then
		alerta = msg_erro
		end if

	dim s
	dim blnIndicadorEdicaoLiberada
	blnIndicadorEdicaoLiberada = False
	s = "SELECT * FROM t_COMISSAO_INDICADOR_N4 WHERE (pedido='" & r_pedido.pedido & "')"
	set rs = cn.Execute(s)
	if operacao_permitida(OP_CEN_EDITA_PEDIDO_INDICADOR, s_lista_operacoes_permitidas) then
		if r_pedido.st_entrega<>ST_ENTREGA_CANCELADO And rs.Eof then
			blnIndicadorEdicaoLiberada = True
		else
			if r_pedido.st_entrega = ST_ENTREGA_CANCELADO then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O status do pedido é inválido para realizar esta operação (" & x_status_entrega(r_pedido.st_entrega) & ")"
				end if
			if Not rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O indicador não pode ser alterado porque este pedido já foi processado no relatório de comissões"
				end if
			end if
		end if
	if rs.State <> 0 then rs.Close

	dim r_orcamento_cotacao
	if converte_numero(Trim("" & r_pedido.IdOrcamentoCotacao)) > 0 then
		if le_orcamento_cotacao(r_pedido.IdOrcamentoCotacao, r_orcamento_cotacao, msg_erro) then
			if (r_orcamento_cotacao.IdIndicador <> ID_NSU_ORCAMENTISTA_E_INDICADOR__SEM_INDICADOR) And (Trim("" & r_orcamento_cotacao.IdIndicador) <> "") then
				blnIndicadorEdicaoLiberada = False
				alerta=texto_add_br(alerta)
				alerta=alerta & "Não é possível alterar o indicador porque o pedido foi gerado a partir do orçamento nº " & formata_inteiro(r_pedido.IdOrcamentoCotacao)
				end if
			end if
		end if

	dim r_loja
	if alerta = "" then
		set r_loja = New cl_LOJA
			if Not x_loja_bd(r_pedido.loja, r_loja) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "A loja do pedido (" & r_pedido.loja & ") não foi encontrada"
				end if
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________________
' INDICADORES MONTA ITENS SELECT
'
function indicadores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, s_sql
	id_default = Trim("" & id_default)
	ha_default=False

	s_sql = "SELECT" & _
				" apelido" & _
				", razao_social_nome_iniciais_em_maiusculas" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR" & _
			" WHERE" & _
				" (Id NOT IN (" & Cstr(ID_NSU_ORCAMENTISTA_E_INDICADOR__RESTRICAO_FP_TODOS) & "," & Cstr(ID_NSU_ORCAMENTISTA_E_INDICADOR__SEM_INDICADOR) & "))"

	if Trim("" & r_loja.unidade_negocio) = COD_UNIDADE_NEGOCIO_LOJA__AC then
		s_sql = s_sql & _
				" AND (loja IN (SELECT loja FROM t_LOJA WHERE unidade_negocio = '" & COD_UNIDADE_NEGOCIO_LOJA__AC & "'))"
	elseif (Trim("" & r_loja.unidade_negocio) = COD_UNIDADE_NEGOCIO_LOJA__BS) Or (Trim("" & r_loja.unidade_negocio) = COD_UNIDADE_NEGOCIO_LOJA__VRF) then
		s_sql = s_sql & _
				" AND (loja IN (SELECT loja FROM t_LOJA WHERE unidade_negocio IN ('" & COD_UNIDADE_NEGOCIO_LOJA__BS & "','" & COD_UNIDADE_NEGOCIO_LOJA__VRF & "')))"
		end if

	if id_default <> "" then
		s_sql = s_sql & _
					" OR (apelido = '" & QuotedStr(id_default) & "')"
		end if

	s_sql = s_sql & _
			" ORDER BY" & _
				" apelido"

	set r = cn.Execute(s_sql)
	strResp = "<option value=''>&nbsp;</option>"
	do while Not r.eof 
		x = UCase(Trim("" & r("apelido")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("razao_social_nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
		
	indicadores_monta_itens_select = strResp
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
	function fPedEditaIndicadorConfirma(f) {
		if (trim(f.c_indicador_novo.value) == "") {
			alert("Selecione o novo indicador!");
			f.c_indicador_novo.focus();
			return;
		}

		if (trim(f.c_indicador_novo.value) == trim(f.c_indicador_original.value)) {
			alert("Não houve alteração do indicador!");
			f.c_indicador_novo.focus();
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
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
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

<form id="fPED" name="fPED" method="post" action="PedidoEditaIndicadorConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="url_origem" id="url_origem" value="<%=url_origem%>" />
<input type="hidden" name="c_indicador_original" id="c_indicador_original" value="<%=r_pedido.indicador%>" />

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="left" valign="bottom"><span class="PEDIDO">Edição do Indicador</span></td>
	<td align="right" valign="bottom"><span class="PEDIDO" style="font-size:14pt;">Pedido <%=pedido_selecionado%></span></td>
</tr>
</table>

<br />

<table class="Q" style="width:649px;" cellSpacing="0">
	<tr>
		<td colspan="2">
			<p class="Rf">INDICADOR ATUAL</p>
			<p class="C" style="font-size:10pt; margin:4px 0px 2px 4px;"><%=r_pedido.indicador%></p>
			<p class="C" style="font-size:10pt;margin: 0px 0px 2px 4px;font-style:italic;"><%=x_orcamentista_e_indicador(r_pedido.indicador)%></p>
		</td>
	</tr>
	<tr>
		<td align="left" valign="bottom" colspan="2" class="MC">
			<p class="Rf">INDICADOR NOVO</p>
			<p class="C">
			<select name="c_indicador_novo" id="c_indicador_novo" style="width:500px;margin:4px 0px 8px 4px;">
			<%=indicadores_monta_itens_select(r_pedido.indicador)%>
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
	<td><a name="bCANCELA" id="bCANCELA" href="javascript:history.back()" title="cancela a edição do indicador">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPedEditaIndicadorConfirma(fPED)" title="grava a alteração do indicador">
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