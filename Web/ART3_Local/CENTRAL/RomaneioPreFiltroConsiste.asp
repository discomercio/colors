<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RomaneioPreFiltroConsiste.asp
'     ========================================================
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

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim alerta
	alerta=""
	
	dim s, s_sql, s_filtro, s_nome_transportadora
	dim c_transportadora, c_dt_entrega, dt_entrega, c_nfe_emitente
	c_transportadora = Trim(Request("c_transportadora"))
	c_dt_entrega = Trim(Request("c_dt_entrega"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
	
	if c_transportadora = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Nenhuma transportadora foi selecionada."
		end if
	
	if c_dt_entrega = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "A data de coleta não foi informada."
	elseif Not IsDate(c_dt_entrega) then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Data de coleta é inválida"
		end if
	
	if alerta = "" then
		dt_entrega = StrToDate(c_dt_entrega)
		end if
	
	if alerta = "" then
		s_nome_transportadora = ""
		if c_transportadora <> "" then
			s = "SELECT nome FROM t_TRANSPORTADORA WHERE (id='" & c_transportadora & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta=alerta & "TRANSPORTADORA " & c_transportadora & " NÃO ESTÁ CADASTRADA."
			else
				s_nome_transportadora = iniciais_em_maiusculas(Trim("" & rs("nome")))
				end if
			end if
		end if
	
	if alerta = "" then
		if c_nfe_emitente = "" then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Não foi informado o CD"
		elseif converte_numero(c_nfe_emitente) = 0 then
			alerta=texto_add_br(alerta)
			alerta = alerta & "É necessário definir um CD válido"
			end if
		end if
	
	dim rNfeEmitente
	dim qtdePedidos, intCounter, intIndex
	qtdePedidos = 0
	
	dim vPedido()
	redim vPedido(0)
	vPedido(UBound(vPedido)) = ""
	
	if alerta = "" then
		s_sql = "SELECT" & _
					" pedido" & _
				" FROM t_PEDIDO" & _
				" WHERE" & _
					" (st_entrega = '" & ST_ENTREGA_A_ENTREGAR & "')" & _
					" AND (danfe_impressa_status = " & COD_DANFE_IMPRESSA_STATUS__OK & ")" & _
					" AND (a_entregar_data_marcada = " & bd_formata_data(dt_entrega) & ")" & _
					" AND (transportadora_id = '" & c_transportadora & "')"
		
		set rNfeEmitente = le_nfe_emitente(c_nfe_emitente)
		s_sql = s_sql & " AND (t_PEDIDO.id_nfe_emitente = " & rNfeEmitente.id & ")"
		
		s_sql = s_sql & _
				" ORDER BY" & _
					" data_hora"
		
		if rs.State <> 0 then rs.Close
		rs.open s_sql, cn
		do while Not rs.Eof
			qtdePedidos = qtdePedidos + 1
			if Trim(vPedido(UBound(vPedido))) <> "" then
				redim preserve vPedido(UBound(vPedido)+1)
				end if
			
			vPedido(UBound(vPedido)) = Trim("" & rs("pedido"))
			rs.MoveNext
			loop
		
		if qtdePedidos = 0 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Não há pedidos para os parâmetros informados."
			end if
		end if
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
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
		$(".CKBPED").change(function() {
			if ($(this).is(":checked")) {
				$(this).parent().addClass("REALCADO");
			}
			else {
				$(this).parent().removeClass("REALCADO");
			}
		});
	});
</script>

<script language="JavaScript" type="text/javascript">
function marcarTodos() {
	$(".CKBPED").prop("checked", true);
	$(".CKBPED").parent().addClass("REALCADO");
}

function desmarcarTodos() {
	$(".CKBPED").prop("checked", false);
	$(".CKBPED").parent().removeClass("REALCADO");
}

function corrigeCorBackground() {
	// Caso o usuário volte p/ esta página através de um history.back(), a cor de realce do background não é restaurada automaticamente pelo IE.
	$.each($(".CKBPED"), function() {
		if ($(this).is(":checked")) {
			$(this).parent().addClass("REALCADO");
		}
		else {
			$(this).parent().removeClass("REALCADO");
		}
	});
}

function fFILTROConfirma( f ) {
var sid, lista_pedidos_selecionados;
var qtde_pedidos, qtde_pedidos_selecionada;
	
	if (trim(f.c_transportadora.value)=="") {
		alert("Não há transportadora selecionada!!");
		return;
		}
	
	if (trim(f.c_dt_entrega.value)=="") {
		alert("A data de coleta não foi informada!!");
		return;
		}
	
	qtde_pedidos = converte_numero(f.c_qtde_pedidos.value);
	if (qtde_pedidos == 0) {
		alert("Nenhum pedido atende aos critérios!!");
		return;
	}
	
	qtde_pedidos_selecionada = 0;
	lista_pedidos_selecionados = "";
	for (i = 1; i <= qtde_pedidos; i++) {
		sid = "#ckb_pedido_" + i;
		if ($(sid).is(":checked")) {
			qtde_pedidos_selecionada++;
			if (lista_pedidos_selecionados.length > 0) lista_pedidos_selecionados += "|";
			lista_pedidos_selecionados += $(sid).val();
		}
	}

	if (qtde_pedidos_selecionada == 0) {
		alert("Nenhum pedido foi selecionado!!");
		return;
	}
	
	f.c_lista_pedidos_selecionados.value = lista_pedidos_selecionados;
	
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__ESCREEN_CSS%>" rel="stylesheet" type="text/css" media="screen">

<style type="text/css">
.C2
{
	font-size:9pt;
}
.REALCADO
{
	background-color:#98FB98;
}
.TRNSP
{
	background-color:transparent;
}
</style>


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
	<td align="center"><a name="bVOLTAR" id="A1" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>
<% else %>

<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="corrigeCorBackground();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RomaneioFiltro.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_entrega" id="c_dt_entrega" value="<%=c_dt_entrega%>" />
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>" />
<input type="hidden" name="c_qtde_pedidos" id="c_qtde_pedidos" value="<%=Cstr(qtdePedidos)%>" />
<input type="hidden" name="c_lista_pedidos_selecionados" id="c_lista_pedidos_selecionados" value="" />
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Romaneio de Entrega</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
	s = c_transportadora
	if s = "" then 
		s = "todas"
	else
		if (s_nome_transportadora <> "") And (s_nome_transportadora <> c_transportadora) then s = s & "  (" & s_nome_transportadora & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Transportadora:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"
	
	s = c_dt_entrega
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Data de Coleta:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
	
	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<br />

<table class="Qx" width="130" cellspacing="0" cellpadding="2">
<!--  PEDIDOS  -->
	<tr>
		<td class="ME MD MC MB PLTe" nowrap align="center" valign="bottom" style="background-color:#F0FFFF;">&nbsp;PEDIDOS</td>
	</tr>
<%	intIndex = 0
	for intCounter = LBound(vPedido) to UBound(vPedido) %>
	<% if Trim(vPedido(intCounter)) <> "" then 
			intIndex = intIndex + 1%>
	<tr bgcolor="#FFFFFF" nowrap>
	<td class="MDBE" align="center" valign="bottom" nowrap>
		<input type="checkbox" class="TA C2 TRNSP CKBPED" name="ckb_pedido_<%=Cstr(intIndex)%>" id="ckb_pedido_<%=Cstr(intIndex)%>" value="<%=Trim("" & vPedido(intCounter))%>" /><span class="C C2 TRNSP" style="cursor:default" onclick="fFILTRO.ckb_pedido_<%=Cstr(intIndex)%>.click();"><%=Trim("" & vPedido(intCounter))%></span>
	</td>
	</tr>
	<% end if %>
<% next %>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width='649' cellpadding='0' cellspacing='0' border='0' style="margin-top:5px;margin-bottom:20px;">
<tr>
	<td width="100%" align="right">
		<table class="notPrint" cellpadding='0' cellspacing='0' border='0'>
			<tr>
			<td align="right" nowrap><a id="linkMarcarTodos" href="javascript:marcarTodos();"><span class="Button" style="margin-bottom:0px;">&nbsp;&nbsp;&nbsp;Marcar Todos&nbsp;&nbsp;&nbsp;</span></a></td>
			<td style="width:10px;" nowrap>&nbsp;</td>
			<td align="right" nowrap><a id="linkDesmarcarTodos" href="javascript:desmarcarTodos();"><span class="Button" style="margin-bottom:0px;">&nbsp;&nbsp;&nbsp;Desmarcar Todos&nbsp;&nbsp;&nbsp;</span></a></td>
			</tr>
		</table>
	</td>
</tr>
</table>

<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="confirma a operação">
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
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
