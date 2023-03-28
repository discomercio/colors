<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================
'	  CadPercMaxComissaoEDescPorLoja.asp
'     ===========================================================
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
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear

'	OBTEM USUÁRIO
	dim usuario
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_CAD_PERC_MAX_COMISSAO_E_DESC_POR_LOJA, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim s, s_nome_loja
	dim n_reg, n_reg_total
	
	dim alerta
	alerta = ""
	
	dim i, rP, vMPN2, strTextoObsNivel2
	set rP = get_registro_t_parametro(ID_PARAMETRO_PercMaxComissaoEDesconto_Nivel2_MeiosPagto)
	
	if Trim("" & rP.id) = "" then
		alerta = "Não foi localizado o registro da tabela de parâmetros que informa os meios de pagamento que usufruem do percentual de Comissão+Desconto de nível 2!!"
		end if
	
	strTextoObsNivel2 = ""
	if alerta = "" then
		vMPN2 = Split(rP.campo_texto, ",")
		for i=Lbound(vMPN2) to Ubound(vMPN2)
			if Trim("" & vMPN2(i)) <> "" then
				if strTextoObsNivel2 <> "" then strTextoObsNivel2 = strTextoObsNivel2 & ", "
				strTextoObsNivel2 = strTextoObsNivel2 & x_opcao_forma_pagamento(Trim("" & vMPN2(i)))
				end if
			next
		if strTextoObsNivel2 <> "" then
			strTextoObsNivel2 = "Meios de pagamento selecionados para o percentual de Comissão+Desconto do nível 2:<br />" & strTextoObsNivel2
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
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_GLOBAL%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	$(document).ready(function() {
		$("#divDialogBox").hide();
		$("#divDialogBox").hUtilUI('dialog_modal');

		// Tratamento p/ bug do jQuery-ui Dialog: ao tentar mover o dialog em uma tela que está c/ scroll
		// vertical, o dialog é "redesenhado" mais abaixo da posição do cursor na mesma medida do deslocamento do
		// scroll vertical. A movimentação do dialog ocorre c/ esse espaço em branco entre o cursor e o dialog.
		$(document).scroll(function(e) {
			if ($(".ui-widget-overlay")) //the dialog has popped up in modal view
			{
				//fix the overlay so it scrolls down with the page
				$(".ui-widget-overlay").css({
					position: 'fixed',
					top: '0'
				});
				//get the current popup position of the dialog box
				pos = $(".ui-dialog").position();
				//adjust the dialog box so that it scrolls as you scroll the page
				$(".ui-dialog").css({
					position: 'fixed',
					top: pos.y
				});
			}
		});
	});
</script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando consulta ...';

function realca_cor_row(id) {
var c;
	c = document.getElementById(id);
	c.style.backgroundColor = 'palegreen';
}

function normaliza_cor_row(id) {
var c;
	c = document.getElementById(id);
	c.style.backgroundColor = '';
}

function AtualizaCadastro(f) {
	var i, s_msg;

	s_msg = "";
	for (i = 0; i < f.c_perc_comissao_e_desconto.length; i++) {
		if (trim(f.c_loja[i].value) != "") {
			if (converte_numero(f.c_perc_comissao_e_desconto_nivel2[i].value) < converte_numero(f.c_perc_comissao_e_desconto[i].value)) {
				if (s_msg.length > 0) s_msg += "\n";
				s_msg += "Loja " + f.c_loja[i].value + ": " + f.c_perc_comissao_e_desconto[i].value + "% e " + f.c_perc_comissao_e_desconto_nivel2[i].value + "% (PF)";
			}
			if (converte_numero(f.c_perc_comissao_e_desconto_nivel2_pj[i].value) < converte_numero(f.c_perc_comissao_e_desconto_pj[i].value)) {
				if (s_msg.length > 0) s_msg += "\n";
				s_msg += "Loja " + f.c_loja[i].value + ": " + f.c_perc_comissao_e_desconto_pj[i].value + "% e " + f.c_perc_comissao_e_desconto_nivel2_pj[i].value + "% (PJ)";
			}
		}
	}

	if (s_msg.length > 0) {
		s_msg = "<b>O percentual de Comissão+Desconto do nível 1 NÃO pode ser maior que o percentual do nível 2!!\n\nIsso ocorreu nas seguinte(s) loja(s):</b>" + "\n" + s_msg;
		s_msg = s_msg.replace(/\n/g, '<br />');
		$("#divDialogBox div").html(s_msg);
		$("#divDialogBox").dialog("option", "title", "Erro!");
		$("#divDialogBox").dialog("open");
		return;
	}
	
	dATUALIZA.style.visibility = "hidden";
	window.status = "Aguarde ...";
	f.submit();
}

</script>

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
.tdNumLoja{
	vertical-align: top;
	width: 60px;
	}
.tdNomeLoja{
	vertical-align: top;
	width: 175px;
	}
.tdPercComissao{
	vertical-align: top;
	width: 80px;
	}
.tdPercComissaoEDesc{
	vertical-align: top;
	width: 85px;
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
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>
<% else %>

<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Percentual Máximo de Comissão e Desconto por Loja</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<br>
<center>
<form method="post" action="CadPercMaxComissaoEDescPorLojaConfirma.asp" id="fCAD" name="fCAD">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" id="c_loja" name="c_loja" value=''>
<input type="hidden" id="c_nome_loja" name="c_nome_loja" value=''>
<input type="hidden" id="c_perc_comissao" name="c_perc_comissao" value=''>
<input type="hidden" id="c_perc_comissao_e_desconto" name="c_perc_comissao_e_desconto" value=''>
<input type="hidden" id="c_perc_comissao_e_desconto_pj" name="c_perc_comissao_e_desconto_pj" value=''>
<input type="hidden" id="c_perc_comissao_e_desconto_nivel2" name="c_perc_comissao_e_desconto_nivel2" value=''>
<input type="hidden" id="c_perc_comissao_e_desconto_nivel2_pj" name="c_perc_comissao_e_desconto_nivel2_pj" value=''>
<input type="hidden" id="c_perc_alcada1_pf" name="c_perc_alcada1_pf" value="" />
<input type="hidden" id="c_perc_alcada1_pj" name="c_perc_alcada1_pj" value="" />
<input type="hidden" id="c_perc_alcada2_pf" name="c_perc_alcada2_pf" value="" />
<input type="hidden" id="c_perc_alcada2_pj" name="c_perc_alcada2_pj" value="" />
<input type="hidden" id="c_perc_alcada3_pf" name="c_perc_alcada3_pf" value="" />
<input type="hidden" id="c_perc_alcada3_pj" name="c_perc_alcada3_pj" value="" />

<%
'	RECUPERA VALOR ANTERIOR
	s = "SELECT" & _
			" loja," & _
			" perc_max_comissao," & _
			" perc_max_comissao_e_desconto," & _
			" perc_max_comissao_e_desconto_pj," & _
			" perc_max_comissao_e_desconto_nivel2," & _
			" perc_max_comissao_e_desconto_nivel2_pj," & _
			" perc_max_comissao_e_desconto_alcada1_pf," & _
			" perc_max_comissao_e_desconto_alcada1_pj," & _
			" perc_max_comissao_e_desconto_alcada2_pf," & _
			" perc_max_comissao_e_desconto_alcada2_pj," & _
			" perc_max_comissao_e_desconto_alcada3_pf," & _
			" perc_max_comissao_e_desconto_alcada3_pj," & _
			" nome," & _
			" razao_social" & _
		" FROM t_LOJA" & _
		" ORDER BY" & _
			" Convert(smallint, loja)"
	if rs.State <> 0 then rs.Close
	rs.open s, cn
%>


<table border="0" cellspacing="0" cellpadding="0">
<tr style="background-color:#ffffff;">
	<td class="MD" colspan="3">&nbsp;</td>
	<td colspan="2" class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>NÍVEL 1</span></td>
	<td colspan="2" class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>NÍVEL 2</span></td>
	<td colspan="2" class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>ALÇADA 1</span></td>
	<td colspan="2" class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>ALÇADA 2</span></td>
	<td colspan="2" class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>ALÇADA 3</span></td>
</tr>
<tr style="background-color:#ffffff;">
	<td class="MD" colspan="3">&nbsp;</td>
	<td class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>PF</span></td>
	<td class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>PJ</span></td>
	<td class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>PF</span></td>
	<td class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>PJ</span></td>
	<td class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>PF</span></td>
	<td class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>PJ</span></td>
	<td class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>PF</span></td>
	<td class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>PJ</span></td>
	<td class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>PF</span></td>
	<td class="MC MD tdPercComissaoEDesc Cc" align="center" style="font-size:10pt;padding-right:4px;background-color:whitesmoke;"><span>PJ</span></td>
</tr>
<tr style="background-color:whitesmoke;">
	<td class="MT tdNumLoja C" align="center" style="font-size:10pt;vertical-align:bottom;">Nº Loja</td>
	<td class="MC MB MD tdNomeLoja C" align="left" style="font-size:10pt;vertical-align:bottom;padding-left:2px;">Nome Loja</td>
	<td class="MC MB MD tdPercComissao Cd" align="right" style="font-size:10pt;padding-right:4px;"><span>Máx</span><br /><span>Comissão</span></td>
	<td class="MC MB MD tdPercComissaoEDesc Cd" align="right" style="font-size:10pt;padding-right:4px;"><span>Máx</span><br /><span>Comis + Desc</span></td>
	<td class="MC MB MD tdPercComissaoEDesc Cd" align="right" style="font-size:10pt;padding-right:4px;"><span>Máx</span><br /><span>Comis + Desc</span></td>
	<td class="MC MB MD tdPercComissaoEDesc Cd" align="right" style="font-size:10pt;padding-right:4px;"><span>Máx</span><br /><span>Comis + Desc</span></td>
	<td class="MC MB MD tdPercComissaoEDesc Cd" align="right" style="font-size:10pt;padding-right:4px;"><span>Máx</span><br /><span>Comis + Desc</span></td>
	<td class="MC MB MD tdPercComissaoEDesc Cd" align="right" style="font-size:10pt;padding-right:4px;"><span>Máx</span><br /><span>Comis + Desc</span></td>
	<td class="MC MB MD tdPercComissaoEDesc Cd" align="right" style="font-size:10pt;padding-right:4px;"><span>Máx</span><br /><span>Comis + Desc</span></td>
	<td class="MC MB MD tdPercComissaoEDesc Cd" align="right" style="font-size:10pt;padding-right:4px;"><span>Máx</span><br /><span>Comis + Desc</span></td>
	<td class="MC MB MD tdPercComissaoEDesc Cd" align="right" style="font-size:10pt;padding-right:4px;"><span>Máx</span><br /><span>Comis + Desc</span></td>
	<td class="MC MB MD tdPercComissaoEDesc Cd" align="right" style="font-size:10pt;padding-right:4px;"><span>Máx</span><br /><span>Comis + Desc</span></td>
	<td class="MC MB MD tdPercComissaoEDesc Cd" align="right" style="font-size:10pt;padding-right:4px;"><span>Máx</span><br /><span>Comis + Desc</span></td>
</tr>
<%  n_reg_total = 0
	do while Not rs.Eof
		n_reg_total = n_reg_total + 1
		rs.MoveNext
		loop
		
	rs.MoveFirst
	
	n_reg = 0
	do while Not rs.Eof
		n_reg = n_reg + 1
		s_nome_loja = Trim("" & rs("nome"))
		if s_nome_loja = "" then s_nome_loja = Trim("" & rs("razao_social"))
		if s_nome_loja <> "" then s_nome_loja = iniciais_em_maiusculas(s_nome_loja)
%>
<tr id="TR_<%=Cstr(n_reg)%>">
	<td class="MB MD ME tdNumLoja" align="center"><input id="c_loja" name="c_loja" readonly tabindex=-1 class="PLLc" style="font-size:10pt;width:55px;background-color:Transparent;" value="<%=Trim("" & rs("loja"))%>"></td>
	<td class="MB MD tdNomeLoja" align="left"><input id="c_nome_loja" name="c_nome_loja" readonly tabindex=-1 class="PLLe" style="font-size:10pt;width:170px;background-color:Transparent;" value="<%=s_nome_loja%>"></td>
	<td class="MB MD tdPercComissao" align="right"><input id="c_perc_comissao" name="c_perc_comissao" class="PLLd" style="font-size:10pt;width:60px;background-color:Transparent;"
		value="<%=formata_perc(rs("perc_max_comissao"))%>" maxlength="5" 
		onkeypress="if (digitou_enter(true)) fCAD.c_perc_comissao_e_desconto[<%=Cstr(n_reg)%>].focus(); filtra_percentual();"
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
	<!-- Nível 1 -->
	<td class="MB MD tdPercComissaoEDesc" align="right"><input id="c_perc_comissao_e_desconto" name="c_perc_comissao_e_desconto" class="PLLd" style="font-size:10pt;width:60px;background-color:Transparent;"
		value="<%=formata_perc(rs("perc_max_comissao_e_desconto"))%>" maxlength="5" 
		onkeypress="if (digitou_enter(true)) fCAD.c_perc_comissao_e_desconto_pj[<%=Cstr(n_reg)%>].focus(); filtra_percentual();"
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
	<td class="MB MD tdPercComissaoEDesc" align="right"><input id="c_perc_comissao_e_desconto_pj" name="c_perc_comissao_e_desconto_pj" class="PLLd" style="font-size:10pt;width:60px;background-color:Transparent;"
		value="<%=formata_perc(rs("perc_max_comissao_e_desconto_pj"))%>" maxlength="5" 
		onkeypress="if (digitou_enter(true)) fCAD.c_perc_comissao_e_desconto_nivel2[<%=Cstr(n_reg)%>].focus(); filtra_percentual();"
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
	<!-- Nível 2 -->
	<td class="MB MD tdPercComissaoEDesc" align="right"><input id="c_perc_comissao_e_desconto_nivel2" name="c_perc_comissao_e_desconto_nivel2" class="PLLd" style="font-size:10pt;width:60px;background-color:Transparent;"
		value="<%=formata_perc(rs("perc_max_comissao_e_desconto_nivel2"))%>" maxlength="5" 
		onkeypress="if (digitou_enter(true)) fCAD.c_perc_comissao_e_desconto_nivel2_pj[<%=Cstr(n_reg)%>].focus(); filtra_percentual();"
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
	<td class="MB MD tdPercComissaoEDesc" align="right"><input id="c_perc_comissao_e_desconto_nivel2_pj" name="c_perc_comissao_e_desconto_nivel2_pj" class="PLLd" style="font-size:10pt;width:60px;background-color:Transparent;"
		value="<%=formata_perc(rs("perc_max_comissao_e_desconto_nivel2_pj"))%>" maxlength="5"
		onkeypress="if (digitou_enter(true)) fCAD.c_perc_alcada1_pf[<%=Cstr(n_reg)%>].focus(); filtra_percentual();"
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
	<!-- Alçada 1 -->
	<td class="MB MD tdPercComissaoEDesc" align="right"><input id="c_perc_alcada1_pf" name="c_perc_alcada1_pf" class="PLLd" style="font-size:10pt;width:60px;background-color:Transparent;"
		value="<%=formata_perc(rs("perc_max_comissao_e_desconto_alcada1_pf"))%>" maxlength="5"
		onkeypress="if (digitou_enter(true)) fCAD.c_perc_alcada1_pj[<%=Cstr(n_reg)%>].focus(); filtra_percentual();"
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
	<td class="MB MD tdPercComissaoEDesc" align="right"><input id="c_perc_alcada1_pj" name="c_perc_alcada1_pj" class="PLLd" style="font-size:10pt;width:60px;background-color:Transparent;"
		value="<%=formata_perc(rs("perc_max_comissao_e_desconto_alcada1_pj"))%>" maxlength="5"
		onkeypress="if (digitou_enter(true)) fCAD.c_perc_alcada2_pf[<%=Cstr(n_reg)%>].focus(); filtra_percentual();"
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
	<!-- Alçada 2 -->
	<td class="MB MD tdPercComissaoEDesc" align="right"><input id="c_perc_alcada2_pf" name="c_perc_alcada2_pf" class="PLLd" style="font-size:10pt;width:60px;background-color:Transparent;"
		value="<%=formata_perc(rs("perc_max_comissao_e_desconto_alcada2_pf"))%>" maxlength="5"
		onkeypress="if (digitou_enter(true)) fCAD.c_perc_alcada2_pj[<%=Cstr(n_reg)%>].focus(); filtra_percentual();"
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
	<td class="MB MD tdPercComissaoEDesc" align="right"><input id="c_perc_alcada2_pj" name="c_perc_alcada2_pj" class="PLLd" style="font-size:10pt;width:60px;background-color:Transparent;"
		value="<%=formata_perc(rs("perc_max_comissao_e_desconto_alcada2_pj"))%>" maxlength="5"
		onkeypress="if (digitou_enter(true)) fCAD.c_perc_alcada3_pf[<%=Cstr(n_reg)%>].focus(); filtra_percentual();"
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
	<!-- Alçada 3 -->
	<td class="MB MD tdPercComissaoEDesc" align="right"><input id="c_perc_alcada3_pf" name="c_perc_alcada3_pf" class="PLLd" style="font-size:10pt;width:60px;background-color:Transparent;"
		value="<%=formata_perc(rs("perc_max_comissao_e_desconto_alcada3_pf"))%>" maxlength="5"
		onkeypress="if (digitou_enter(true)) fCAD.c_perc_alcada3_pj[<%=Cstr(n_reg)%>].focus(); filtra_percentual();"
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
	<td class="MB MD tdPercComissaoEDesc" align="right"><input id="c_perc_alcada3_pj" name="c_perc_alcada3_pj" class="PLLd" style="font-size:10pt;width:60px;background-color:Transparent;"
		value="<%=formata_perc(rs("perc_max_comissao_e_desconto_alcada3_pj"))%>" maxlength="5"
		<% if n_reg = n_reg_total then %>
		onkeypress="if (digitou_enter(true)) bATUALIZA.focus(); filtra_percentual();"
		<% else %>
		onkeypress="if (digitou_enter(true)) fCAD.c_perc_comissao[<%=Cstr(n_reg+1)%>].focus(); filtra_percentual();"
		<% end if %>
		onfocus="this.select();realca_cor_row('TR_<%=Cstr(n_reg)%>');"
		onblur="this.value=formata_numero(this.value,2);normaliza_cor_row('TR_<%=Cstr(n_reg)%>');"><span class="PLTd" style='vertical-align:middle;'>&nbsp;%</span></td>
</tr>
<%		rs.MoveNext
		loop
%>
<tr>
	<td colspan="7">&nbsp;</td>
</tr>
<tr>
	<td colspan="7" align="left" style="font-style:italic;"><%=strTextoObsNivel2%></td>
</tr>
</table>

<!--  DIV P/ DIALOG BOX -->
<div id="divDialogBox">
<div></div>
</div>

</form>

<br />

<p class="TracoBottom"></p>

<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="left"><a href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaCadastro(fCAD)" title="salva as alterações">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>

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