<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%

'     ===========================================================
'	  M E N U O R C A M E N T I S T A E I N D I C A D O R . A S P
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
'						I N I C I A L I Z A     P Á G I N A     A S P
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
	Const COD_ID_FORMA_PAGTO_SEM_RESTRICOES = 999
	
'	OBTEM USUÁRIO
	dim s, usuario, usuario_nome,permissao
	usuario = Trim(Session("usuario_atual"))
	usuario_nome = Trim(Session("usuario_nome_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	if Not (operacao_permitida(OP_CEN_CADASTRO_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_GER_LIST_CAD_INDICADORES, s_lista_operacoes_permitidas)) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	CONECTA COM O BANCO DE DADOS
	dim cn, t
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

    if operacao_permitida(OP_CEN_CADASTRO_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) then permissao="" else permissao=" disabled" 

	dim intIdx
	dim s_ckb_PF_id, s_ckb_PJ_id
	dim s_span_PF_id, s_span_PJ_id
	dim s_ckb_PF_value, s_ckb_PJ_value
	dim s_lista_id_forma_pagto
	s_lista_id_forma_pagto = ""

	'RECUPERA OPÇÃO MEMORIZADA
	dim ordenacao_default
	ordenacao_default = get_default_valor_texto_bd(usuario, "MenuOrcamentistaEIndicador|ordenacao")
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
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	var s_ckb_id, s_spn_id;

	$(function() {
		$(".CKB_PF, .CKB_PJ").each(function() {
			s_ckb_id = $(this).attr('id');
			s_spn_id = s_ckb_id.replace("ckb_", "spn_");
			if ($(this).is(':checked')) {
				$("#" + s_spn_id).css('color', 'red');
			}
			else {
				$("#" + s_spn_id).css('color', 'darkgreen');
			}
		});

		$(".CKB_PF, .CKB_PJ").change(function() {
			s_ckb_id = $(this).attr('id');
			s_spn_id = s_ckb_id.replace("ckb_", "spn_");
			if ($(this).is(':checked')) {
				$("#" + s_spn_id).css('color', 'red');
			}
			else {
				$("#" + s_spn_id).css('color', 'darkgreen');
			}
		});

		$(".CKB_PF, .CKB_PJ").click(function() {
			fOP.rb_op[fOP.rb_op.length - 1].click();
		});
	});
</script>

<script language="JavaScript" type="text/javascript">
function fOPConcluir( f ){
	var s_dest, s_op, s_id_selecionado, intIdx;
    
	s_dest="";
	s_op="";
	s_id_selecionado = "";
	intIdx = -1;
	
	intIdx++;
	if (f.rb_op[intIdx].checked) {
		s_dest="OrcamentistaEIndicadorEdita.asp";
		s_op=OP_INCLUI;
		s_id_selecionado=f.c_novo.value;
		if (trim(f.c_novo.value)=="") {
			alert("Forneça a identificação para o novo orçamentista!!");
			f.c_novo.focus();
			return false;
			}
		if ((!f.rb_tipo[0].checked)&&(!f.rb_tipo[1].checked)) {
			alert("Informe se o novo orçamentista / indicador é PF ou PJ!!");
			return false;
			}
		}
	
	intIdx++;
	if (f.rb_op[intIdx].checked) {
		s_dest="OrcamentistaEIndicadorConsulta.asp";
		s_op=OP_CONSULTA;
		s_id_selecionado=f.c_cons.value;
		if (trim(f.c_cons.value)=="") {
			alert("Forneça a identificação do orçamentista a ser consultado!!");
			f.c_cons.focus();
			return false;
			}
	}

	intIdx++;
	if (f.rb_op[intIdx].checked) {
	    s_dest = "OrcamentistaEIndicadorEdita.asp";
	    s_op = OP_CONSULTA;
	    s_id_selecionado = f.c_edit.value;
	    if (trim(f.c_edit.value) == "") {
	        alert("Forneça a identificação do orçamentista a ser editado!!");
	        f.c_edit.focus();
	        return false;
	    }
	}
	
	intIdx++;
	if (f.rb_op[intIdx].checked) {
        s_dest = "OrcamentistaEIndicadorLista.asp?op=A";
        f.filtro_loja.value = f.c_loja_ativos.value;
		}
	
	intIdx++;
	if (f.rb_op[intIdx].checked) {
        s_dest = "OrcamentistaEIndicadorLista.asp?op=I";
        f.filtro_loja.value = f.c_loja_inativos.value;
		}
	
	intIdx++;
	if (f.rb_op[intIdx].checked) {
        s_dest = "OrcamentistaEIndicadorLista.asp?op=T";
        f.filtro_loja.value = f.c_loja_todos.value;
		}
	
	intIdx++;
	if (f.rb_op[intIdx].checked) {
		if (trim(f.vendedor.value)=='') {
			alert('Selecione o vendedor!!');
			f.vendedor.focus();
			return;
			}
		s_dest="OrcamentistaEIndicadorAssocAoVendedor.asp";
		}

	intIdx++;
	if (f.rb_op[intIdx].checked) {
		s_dest = "OrcamentistaEIndicadorListaRestricaoFormaPagto.asp";
	}
	
	if (s_dest=="") {
		alert("Escolha uma das opções!!");
		return false;
		}
	
	f.id_selecionado.value=s_id_selecionado;
	f.operacao_selecionada.value=s_op;
	
	f.action=s_dest;
	dEXECUTAR.style.visibility = "hidden";
	window.status = "Aguarde ...";
	f.submit(); 
}



function geraArquivoCSV(f) {
    var serverVariableUrl, strUrl, xmlHttp;
    var loja = $("#loja").val();
    if (loja == "") {
        loja = "vazio";
    }
    serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
    serverVariableUrl = serverVariableUrl.toUpperCase();
    serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));

    xmlhttp = GetXmlHttpObject();
    if (xmlhttp == null) {
        alert("O browser NÃO possui suporte ao AJAX!!");
        return;
    }

    window.status = "Aguarde, gerando arquivo ...";
    divMsgAguardeObtendoDados.style.visibility = "";

	strUrl = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/GetCadIndicadoresListagemCSV/?loja=' + loja + '&usuario=<%=usuario%>';

    xmlhttp.onreadystatechange = function () {
        var xmlResp;

        if (xmlhttp.readyState == AJAX_REQUEST_IS_COMPLETE) {
            xmlResp = JSON.parse(xmlhttp.responseText);

            if (xmlResp.Status == "OK") {

				gerarRelatorio.action = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/downloadCadIndicadoresListagemCSV/?fileName=' + xmlResp.fileName;
                gerarRelatorio.submit();

                window.status = "Concluído";
                divMsgAguardeObtendoDados.style.visibility = "hidden";
            }
            else {
                window.status = "Concluído";
                divMsgAguardeObtendoDados.style.visibility = "hidden";

                alert("Falha ao gerar o arquivo CSV\n" + xmlResp.Exception);
                return;
            }

        }
    }

    xmlhttp.open("POST", strUrl, true);
    xmlhttp.send();

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
.TitTipoCli
{
	color:black;
}
.CKB_PF
{
}
.CKB_PJ
{
}
.Spantext {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: 10pt;
	border: 0px;
	}
</style>


<body onload="focus()">

<!--  MENU SUPERIOR -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">CENTRAL&nbsp;&nbsp;ADMINISTRATIVA<br>
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
		</span></span></td>
	</tr>

</table>

<br />

<center>
<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>

    <form name="gerarRelatorio" id="gerarRelatorio" method="POST">
    <input type="hidden" name="idRel" id="idRel" value="" />
    </form>
<!--  ***********************************************************************************************  -->
<!--  F O R M U L Á R I O                         												       -->
<!--  ***********************************************************************************************  -->
<form method="post" id="fOP" name="fOP" onsubmit="if (!fOPConcluir(fOP)) return false">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='id_selecionado' id="id_selecionado" value=''>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value=''>
<input type="hidden" name="url_origem" id="url_origem" value="MenuOrcamentistaEIndicador.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" />
<input type="hidden" name="filtro_loja" id="filtro_loja" value="" />



<span class="T">CADASTRO DE ORÇAMENTISTAS / INDICADORES</span>
<div class="QFn" align="CENTER">
<table class="TFn">
	<tr>
		<% intIdx = -1 %>
		<td align="left" nowrap>
			<br />
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(intIdx+1)%>" class="CBOX" onclick="fOP.c_novo.focus()"<%=permissao%>><span style="cursor:default" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click(); fOP.c_novo.focus();">Cadastrar Novo</span>&nbsp;
				<input name="c_novo" id="c_novo" type="text" maxlength="<%=MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR%>" size="20" onblur="this.value=trim(this.value);" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit(); filtra_nome_identificador();"<%=permissao%>
				><span style='width:12px;'></span><input type="radio" id="rb_tipo" name="rb_tipo" value="<%=ID_PF%>" class="CBOX" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click();"<%=permissao%>><span class="rbLink" onclick="fOP.rb_tipo[0].click(); fOP.rb_op[<%=Cstr(intIdx)%>].click();"
				><%=ID_PF%></span>
				<input type="radio" id="rb_tipo" name="rb_tipo" value="<%=ID_PJ%>" class="CBOX" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click();" <%=permissao%>><span class="rbLink" onclick="fOP.rb_tipo[1].click(); fOP.rb_op[<%=Cstr(intIdx)%>].click();"
				><%=ID_PJ%></span>
			<br />
			<br />
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(intIdx+1)%>" class="CBOX" onclick="fOP.c_cons.focus()" <%=permissao%>><span style="cursor:default" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click(); fOP.c_cons.focus();">Consultar</span>&nbsp;
				<input name="c_cons" id="c_cons" type="text" maxlength="<%=MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR%>" size="18" onblur="this.value=trim(this.value);" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit(); filtra_nome_identificador();" <%=permissao%>>
			<br />
			<br />
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(intIdx+1)%>" class="CBOX" onclick="fOP.c_edit.focus()"<%=permissao%>><span style="cursor:default" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click(); fOP.c_edit.focus();">Editar</span>&nbsp;
				<input name="c_edit" id="c_edit" type="text" maxlength="<%=MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR%>" size="18" onblur="this.value=trim(this.value);" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit(); filtra_nome_identificador();"<%=permissao%>>
            <br />
			<br />
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(intIdx+1)%>" class="CBOX"<%=permissao%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click(); fOP.bEXECUTAR.click();">Consultar Ativos</span>&nbsp;&nbsp;
				<span class="Lbl">Loja:</span>&nbsp;<input name="c_loja_ativos" id="c_loja_ativos" type="text" maxlength="3" size="4" onblur="this.value=trim(this.value);" onclick="fOP.rb_op[<%= Cstr(intIdx) %>].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit(); filtra_numerico();">
			<br />
			<br />
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(intIdx+1)%>" class="CBOX"<%=permissao%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click(); fOP.bEXECUTAR.click();">Consultar Inativos</span>&nbsp;&nbsp;
				<span class="Lbl">Loja:</span>&nbsp;<input name="c_loja_inativos" id="c_loja_inativos" type="text" maxlength="3" size="4" onblur="this.value=trim(this.value);" onclick="fOP.rb_op[<%= Cstr(intIdx) %>].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit(); filtra_numerico();">
			<br />
			<br />
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(intIdx+1)%>" class="CBOX"<%=permissao%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click(); fOP.bEXECUTAR.click();">Consultar Ativos e Inativos</span>&nbsp;&nbsp;
				<span class="Lbl">Loja:</span>&nbsp;<input name="c_loja_todos" id="c_loja_todos" type="text" maxlength="3" size="4" onblur="this.value=trim(this.value);" onclick="fOP.rb_op[<%= Cstr(intIdx) %>].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit(); filtra_numerico();">
			<br />
			<br />
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(intIdx+1)%>" class="CBOX"<%=permissao%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click(); fOP.vendedor.focus();">Associados ao Vendedor</span>
			<br />
				<select id="vendedor" name="vendedor" style="margin-left:25px;" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click();" <%=permissao%>>
				  <% =vendedor_do_indicador_monta_itens_select(Null) %>
				</select>
			<br />
			<br />
			<% intIdx = intIdx+1 %>
            <% if operacao_permitida(OP_CEN_GER_LIST_CAD_INDICADORES, s_lista_operacoes_permitidas) then  %>        
                <span class="Spantext" style="margin-left:20px;" >Listagem em CSV</span>
            <br />
				<select id="loja" name="loja" style="margin-left:25px;" >
				  <% =lojas_monta_itens_select(Null) %>
				</select>
            <input type="button" title="Gerar CSV" value="Gerar CSV" onclick="geraArquivoCSV(gerarRelatorio);" />
			<br />
			<br />
            <%end if%>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(intIdx+1)%>" class="CBOX"<%=permissao%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(intIdx)%>].click();">Restrição na Forma de Pagamento</span>
				<br />
				<table class="Q" style="background:transparent;margin-left:25px;" cellpadding="0" cellspacing="0">
					<tr>
						<td style="width:15px;" align="left">&nbsp;</td>
						<td colspan="4"><input type="checkbox" id="ckb_somente_ativos" name="ckb_somente_ativos" value="S" <%=permissao%>><span class="C" style="cursor:default;" onclick="fOP.ckb_somente_ativos.click();">Consultar somente ativos</span></td>
					</tr>
					<tr>
						<td style="width:15px;" align="left">&nbsp;</td>
						<td align="left"><span class="R TitTipoCli">Pessoa Física</span></td>
						<td style="width:40px;" align="left">&nbsp;</td>
						<td align="left"><span class="R TitTipoCli">Pessoa Jurídica</span></td>
						<td style="width:15px;" align="left">&nbsp;</td>
					</tr>
					<tr>
					<%
						s_ckb_PF_id = "ckb_" & ID_PF & "_" & COD_ID_FORMA_PAGTO_SEM_RESTRICOES
						s_ckb_PJ_id = "ckb_" & ID_PJ & "_" & COD_ID_FORMA_PAGTO_SEM_RESTRICOES
						s_span_PF_id = "spn_" & ID_PF & "_" & COD_ID_FORMA_PAGTO_SEM_RESTRICOES
						s_span_PJ_id = "spn_" & ID_PJ & "_" & COD_ID_FORMA_PAGTO_SEM_RESTRICOES
						s_ckb_PF_value = ID_PF & "_" & COD_ID_FORMA_PAGTO_SEM_RESTRICOES
						s_ckb_PJ_value = ID_PJ & "_" & COD_ID_FORMA_PAGTO_SEM_RESTRICOES
						if s_lista_id_forma_pagto <> "" then s_lista_id_forma_pagto = s_lista_id_forma_pagto & "|"
						s_lista_id_forma_pagto = s_lista_id_forma_pagto & COD_ID_FORMA_PAGTO_SEM_RESTRICOES
					%>
						<td align="left">&nbsp;</td>
						<td align="left"><input type="checkbox" id="<%=s_ckb_PF_id%>" name="<%=s_ckb_PF_id%>" value="<%=s_ckb_PF_value%>" class="CKB_PF" <%=permissao%>><span id="<%=s_span_PF_id%>" class="C" style="cursor:default;" onclick="fOP.<%=s_ckb_PF_id%>.click();">Sem Restrições</span></td>
						<td align="left">&nbsp;</td>
						<td align="left"><input type="checkbox" id="<%=s_ckb_PJ_id%>" name="<%=s_ckb_PJ_id%>" value="<%=s_ckb_PJ_value%>" class="CKB_PJ" <%=permissao%>><span id="<%=s_span_PJ_id%>" class="C" style="cursor:default;" onclick="fOP.<%=s_ckb_PJ_id%>.click();">Sem Restrições</span></td>
						<td align="left">&nbsp;</td>
					</tr>
					<% s = "SELECT " & _
								"*" & _
							" FROM t_FORMA_PAGTO" & _
							" WHERE" & _
								" ((hab_a_vista <> 0) OR (hab_entrada <> 0) OR (hab_prestacao <> 0))" & _
							" ORDER BY" & _
								" ordenacao"
					set t = cn.Execute(s)
					do while Not t.Eof
						if s_lista_id_forma_pagto <> "" then s_lista_id_forma_pagto = s_lista_id_forma_pagto & "|"
						s_lista_id_forma_pagto = s_lista_id_forma_pagto & Trim("" & t("id"))
						s_ckb_PF_id = "ckb_" & ID_PF & "_" & Trim("" & t("id"))
						s_ckb_PJ_id = "ckb_" & ID_PJ & "_" & Trim("" & t("id"))
						s_span_PF_id = "spn_" & ID_PF & "_" & Trim("" & t("id"))
						s_span_PJ_id = "spn_" & ID_PJ & "_" & Trim("" & t("id"))
						s_ckb_PF_value = ID_PF & "_" & Trim("" & t("id"))
						s_ckb_PJ_value = ID_PJ & "_" & Trim("" & t("id"))
					%>
					<tr>
						<td align="left">&nbsp;</td>
						<td align="left"><input type="checkbox" id="<%=s_ckb_PF_id%>" name="<%=s_ckb_PF_id%>" value="<%=s_ckb_PF_value%>" class="CKB_PF" <%=permissao%>><span id="<%=s_span_PF_id%>" class="C" style="cursor:default;" onclick="fOP.<%=s_ckb_PF_id%>.click();"><%=Trim("" & t("descricao"))%></span></td>
						<td align="left">&nbsp;</td>
						<td align="left"><input type="checkbox" id="<%=s_ckb_PJ_id%>" name="<%=s_ckb_PJ_id%>" value="<%=s_ckb_PJ_value%>" class="CKB_PJ" <%=permissao%>><span id="<%=s_span_PJ_id%>" class="C" style="cursor:default;" onclick="fOP.<%=s_ckb_PJ_id%>.click();"><%=Trim("" & t("descricao"))%></span></td>
						<td align="left">&nbsp;</td>
					</tr>
				<%
					t.MoveNext
					loop
				%>
				</table>
			<input type="hidden" name="c_lista_id_forma_pagto" id="c_lista_id_forma_pagto" value="<%=s_lista_id_forma_pagto%>" />
			</td>
		</tr>
	</table>
	<br />
	<table width="100%" cellpadding="0" cellspacing="0"><tr><td class="MC" style="height:8px;"></td></tr></table>

	<table cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td><span class="Lbl">Ordenação</span></td>
		</tr>
		<tr>
			<td>
				<select id="ordenacao" name="ordenacao" style="min-width:140px;">
					<% =ordenacao_lista_indicadores_monta_itens_select(ordenacao_default) %>
				</select>
			</td>
		</tr>
	</table>
	
	<table width="100%" cellpadding="0" cellspacing="0"><tr><td class="MB" style="height:12px;"></td></tr></table>

	<br />
	<div name="dEXECUTAR" id="dEXECUTAR">
	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" value="EXECUTAR" title="executa">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
	</div>

</div>
</form>

<br />
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