<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===================================================================
'	  O R C A M E N T I S T A E I N D I C A D O R C O N S U L T A . A S P
'     ===================================================================
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
	dim s, usuario, id_selecionado, tipo_PJ_PF, rs2, s2, cont, novo_bloco, url_origem, url_back, i
	dim s_label, s_parametro, chave, senha_descripto, s_selected, s_color
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
		
    if (Not operacao_permitida(OP_CEN_CADASTRO_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) And _
	   (Not operacao_permitida(OP_CEN_PESQUISA_INDICADORES, s_lista_operacoes_permitidas)) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
'	ORÇAMENTISTA/INDICADOR A EDITAR
	id_selecionado = ucase(trim(request("id_selecionado")))

	if (id_selecionado="") then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTISTA_INDICADOR_NAO_ESPECIFICADO) 
	
'	CONECTA COM O BANCO DE DADOS
	dim cn,rs,r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim alerta
	alerta = ""
	
	set rs = cn.Execute("SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido='" & id_selecionado & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTISTA_INDICADOR_NAO_CADASTRADO)
	tipo_PJ_PF = Trim("" & rs("tipo"))

	dim r_loja_indicador
	if alerta = "" then
		set r_loja_indicador = New cl_LOJA
		if Not x_loja_bd(Trim("" & rs("loja")), r_loja_indicador) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "A loja cadastrada para o indicador (" & Trim("" & rs("loja")) & ") não foi encontrada"
			end if
		end if 'if alerta = ""

	dim blnVisivelIdMagentoB2B
	blnVisivelIdMagentoB2B = False
	if alerta = "" then
		if (r_loja_indicador.unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__AC) _
			Or ( (Trim("" & rs("id_magento_b2b")) <> "") And (Trim("" & rs("id_magento_b2b")) <> "0") ) then
			blnVisivelIdMagentoB2B = True
			end if
		end if 'if alerta = ""

	novo_bloco = Request("NovoBloco")
	url_back = Request("url_back")
    url_origem = Request("url_origem")
%>

<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL</title>
	
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

<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<% if tipo_PJ_PF = ID_PF then %>
<script language="JavaScript" type="text/javascript">
var tipo_PJ_PF = ID_PF;
</script>
<% else %>
<script language="JavaScript" type="text/javascript">
var tipo_PJ_PF = ID_PJ;
</script>
<% end if %>

<script type="text/javascript">

    function fCADAdicionaBlocoNotas(f) {
        f.action = "OrcamentistaEIndicadorBlocoNotasNovo.asp";
        window.status = "Aguarde ...";
        fCAD.url_origem.value = "<%=url_origem%>";
        f.submit();
    }

    function fCADBlocoNotasAlteraImpressao(f) {
        if (document.getElementById("tableBlocoNotas").className == "notPrint") {
            document.getElementById("tableBlocoNotas").className = "";
            document.getElementById("imgPrinterBlocoNotas").src = document.getElementById("imgPrinterBlocoNotas").src.replace("PrinterError.png", "Printer.png");
        }
        else {
            document.getElementById("tableBlocoNotas").className = "notPrint";
            document.getElementById("imgPrinterBlocoNotas").src = document.getElementById("imgPrinterBlocoNotas").src.replace("Printer.png", "PrinterError.png");
        }
    }
    function fCADHistoricoAlteraImpressao(f) {
        if (document.getElementById("tableHistorico").className == "notPrint") {
            document.getElementById("tableHistorico").className = "";
            document.getElementById("imgPrinterHistorico").src = document.getElementById("imgPrinterHistorico").src.replace("PrinterError.png", "Printer.png");
        }
        else {
            document.getElementById("tableHistorico").className = "notPrint";
            document.getElementById("imgPrinterHistorico").src = document.getElementById("imgPrinterHistorico").src.replace("Printer.png", "PrinterError.png");
        }
    }

    function fCADTblDescontosAlteraImpressao(f) {
        if (document.getElementById("tableDescontos").className == "notPrint") {
            document.getElementById("tableDescontos").className = "";
            document.getElementById("tblDesc").className = "";
            document.getElementById("imgPrinterTblDescontos").src = document.getElementById("imgPrinterTblDescontos").src.replace("PrinterError.png", "Printer.png");
        }
        else {
            document.getElementById("tableDescontos").className = "notPrint";
            document.getElementById("tblDesc").className = "notPrint";
            document.getElementById("imgPrinterTblDescontos").src = document.getElementById("imgPrinterTblDescontos").src.replace("Printer.png", "PrinterError.png");
        }
    }

    function fCADDadosEtiquetaAlteraImpressao(f) {
        if (document.getElementById("tableDadosEtiqueta").className == "notPrint") {
            document.getElementById("tableDadosEtiqueta").className = "";
            document.getElementById("Etq1").className = "";
            document.getElementById("Etq2").className = "";
            document.getElementById("Etq3").className = "";
            document.getElementById("Etq4").className = "";
            document.getElementById("Etq5").className = "";
            document.getElementById("Etq6").className = "";
            document.getElementById("imgPrinterDadosEtiqueta").src = document.getElementById("imgPrinterDadosEtiqueta").src.replace("PrinterError.png", "Printer.png");
        }
        else {
            document.getElementById("tableDadosEtiqueta").className = "notPrint";
            document.getElementById("Etq1").className = "notPrint";
            document.getElementById("Etq2").className = "notPrint";
            document.getElementById("Etq3").className = "notPrint";
            document.getElementById("Etq4").className = "notPrint";
            document.getElementById("Etq5").className = "notPrint";
            document.getElementById("Etq6").className = "notPrint";
            document.getElementById("imgPrinterDadosEtiqueta").src = document.getElementById("imgPrinterDadosEtiqueta").src.replace("Printer.png", "PrinterError.png");
        }
    }

    function mostraOcultaMeses(x) {
        if ($('.tableBlocoMes' + x).is(':visible')) {
            $('.tableBlocoMes' + x).hide();
            $('.classeFecha' + x).hide();
            $('#img' + x).attr({ src: '../imagem/plus.gif' });
            $('#img' + x).attr({ title: 'expandir' });
        }
        else {
            $('.tableBlocoMes' + x).show();
            $('#img' + x).attr({ src: '../imagem/minus.gif' });
            $('#img' + x).attr({ title: 'ocultar' });
            $('.imgClasse' + x).attr({ src: '../imagem/plus.gif' });
        }
    }

    function mostraOcultaNotas(ano, mes) {
        if ($("#" + ano + "" + mes).is(':visible')) {
            $("#" + ano + "" + mes).css('display', 'none');
            $("#img" + ano + "" + mes).attr({ src: '../imagem/plus.gif' });
            $("#img" + ano + "" + mes).attr({ title: 'mostrar anotações' });
        }
        else {
            $("#" + ano + "" + mes).css('display', 'block');
            $('#img' + ano + "" + mes).attr({ src: '../imagem/minus.gif' });        
            $("#img" + ano + "" + mes).attr({ title: 'fechar anotações' });
            CarregaBlocoNotas(ano, mes);
        }
    }
    
</script>

<!-- CONSULTA BLOCO DE NOTAS VIA AJAX ---->
<script type="text/javascript">

    function CarregaBlocoNotas(ano, mes) {
        var strUrl, strApelido, xmlhttp;
        strApelido = "<%=id_selecionado%>"
        xmlhttp = GetXmlHttpObject();
        if (xmlhttp == null) {
            alert("O browser NÃO possui suporte ao AJAX!!");
            return;
        }
        if (strApelido == "") {
            return;
        }

        window.status = "Aguarde, pesquisando blocos de notas de  " + strApelido + " ...";
        divMsgAguardeObtendoDados.style.visibility = "";

        strUrl = "../Global/AjaxIndicadoresBlocoNotasConsulta.asp";
        strUrl = strUrl + "?apelido=" + encodeURIComponent(strApelido);
        strUrl = strUrl + "&ano=" + ano;
        strUrl = strUrl + "&mes=" + mes;
        strUrl = strUrl + "&id=" + Math.random();
        xmlhttp.onreadystatechange = function() {
            if (xmlhttp.readyState == 4) {
                $('#' + ano+""+mes).html(xmlhttp.responseText);
                window.status = "Concluído"
                divMsgAguardeObtendoDados.style.visibility = "hidden";
            }
        }
        xmlhttp.open("GET", strUrl, true);
        xmlhttp.send();
    }

</script>

<script type="text/javascript">
    $(function() {
    var mes, ano, data, novoBloco
    novoBloco = "<%=novo_bloco%>";
    data = new Date();
    mes = data.getMonth() + 1;
    ano = data.getFullYear();

    if (novoBloco == 1) {
        mostraOcultaMeses(ano);
        mostraOcultaNotas(ano, mes);
    }
    
    $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');
    
    //Every resize of window
    $(window).resize(function() {
        sizeDivAjaxRunning();
    });

    //Every scroll of window
    $(window).scroll(function() {
        sizeDivAjaxRunning();
    });

    //Dynamically assign height
    function sizeDivAjaxRunning() {
        var newTop = $(window).scrollTop() + "px";
        $("#divMsgAguardeObtendoDados").css("top", newTop);
    }
    });

</script>

<script type="text/javascript">
    //Every resize of window
    $(window).resize(function () {
        sizeDivEtiqueta();
    });

    //Every scroll of window
    $(window).scroll(function () {
        sizeDivEtiqueta();
    });

    //Dynamically assign height
    function sizeDivEtiqueta() {
        var newTop = $(window).scrollTop() + "px";
        $("#div_etiqueta").css("top", newTop);
        $("#etiqueta_layout").css("top", newTop);
    }

    function AbreJanelaEtiqueta() {
        if ($("#etq_endereco").val() == "") {
            alert("Não há dados suficiente!");
            fCAD.etq_endereco.focus();
            return;
        }
        if ($("#etq_endereco_numero").val() == "") {
            alert("Não há dados suficiente!");
            fCAD.etq_endereco_numero.focus();
            return;
        }
        if ($("#etq_cidade").val() == "") {
            alert("Não há dados suficiente!");
            fCAD.etq_cidade.focus();
            return;
        }
        if ($("#etq_uf").val() == "") {
            alert("Não há dados suficiente!");
            fCAD.etq_uf.focus();
            return;
        }
        if ($("#etq_ddd_1").val() != "" || $("#etq_tel_1").val() != "") {
            if ($("#etq_ddd_1").val() == "") {
                alert("Não há dados suficiente!");
                fCAD.etq_ddd_1.focus();
                return;
            }
            if ($("#etq_tel_1").val() == "") {
                alert("Não há dados suficiente!");
                fCAD.etq_tel_1.focus();
                return;
            }
        }
        if ($("#etq_ddd_2").val() != "" || $("#etq_tel_2").val() != "") {
            if ($("#etq_ddd_2").val() == "") {
                alert("Não há dados suficiente!");
                fCAD.etq_ddd_2.focus();
                return;
            }
            if ($("#etq_tel_2").val() == "") {
                alert("Não há dados suficiente!");
                fCAD.etq_tel_2.focus();
                return;
            }
        }

        // torna a etiqueta visível
        $("#div_etiqueta").css('display', 'block');
        $("#etiqueta_layout").css('display', 'block');

        if ($("#c_nome_fantasia").val() != "") {
            $("#etq_nome_fantasia").text($("#c_nome_fantasia").val());
        }
        else {
            $("#etq_nome_fantasia").text($("#razao_social_nome").val());
        }
        if ($("#etq_endereco_complemento").val() == "") {
            $("#separa_complemento").text("");
        }
        else {
            $("#separa_complemento").text(" - ");
        }
        if ($("#etq_bairro").val() == "") {
            $("#separa_bairro").text("");
        }
        else {
            $("#separa_bairro").text(" - ");
        }
        if ($("#etq_cep").val() == "") {
            $("#separa_cep").text("");
        }
        else {
            $("#separa_cep").text(" - ");
        }
        if ($("#etq_ddd_1").val() == "") {
            $("#spn_label_fone").text("");
            $("#spn_fecha_ddd_1").text("");
        }
        else {
            $("#spn_label_fone").text("Fone: (");
            $("#spn_fecha_ddd_1").text(") ");
        }
        if ($("#etq_ddd_2").val() == "") {
            $("#separa_tel").text("");
            $("#spn_abre_ddd_2").text("");
            $("#spn_fecha_ddd_2").text("");
        }
        else {
            $("#separa_tel").text(" / ");
            $("#spn_abre_ddd_2").text("(");
            $("#spn_fecha_ddd_2").text(") ");
        }
        if ($("#etq_email").val() == "") {
            $("#spn_label_email").text("");
        }
        else {
            $("#spn_label_email").text("Email: ");
        }

        $("#spn_etq_endereco").text($("#etq_endereco").val());
        $("#spn_etq_numero").text($("#etq_endereco_numero").val());
        $("#spn_etq_complemento").text($("#etq_endereco_complemento").val());
        $("#spn_etq_bairro").text($("#etq_bairro").val());
        $("#spn_etq_cidade").text($("#etq_cidade").val());
        $("#spn_etq_uf").text($("#etq_uf").val());
        $("#spn_etq_cep").text($("#etq_cep").val());
        $("#spn_etq_ddd_1").text($("#etq_ddd_1").val());
        $("#spn_etq_tel_1").text($("#etq_tel_1").val());
        $("#spn_etq_ddd_2").text($("#etq_ddd_2").val());
        $("#spn_etq_tel_2").text($("#etq_tel_2").val());
        $("#spn_etq_email").text($("#etq_email").val());

        if ($("#etq_ddd_1").val() == $("#etq_ddd_2").val()) {
            $("#spn_abre_ddd_2").text("");
            $("#spn_fecha_ddd_2").text("");
            $("#spn_etq_ddd_2").text("");
        }
        if ($("#etq_ddd_2").val() != "") {
            if ($("#etq_ddd_1").val() == "") {
                $("#spn_etq_ddd_1").text($("#etq_ddd_2").val());
                $("#spn_etq_tel_1").text($("#etq_tel_2").val());
                $("#spn_fecha_ddd_1").text(") ");
                $("#spn_etq_ddd_2").text("");
                $("#spn_etq_tel_2").text("");
                $("#separa_tel").text("");
                $("#spn_abre_ddd_2").text("");
                $("#spn_fecha_ddd_2").text("");
                $("#spn_label_fone").text("Fone: (");
            }
        }

    }

    function fechaEtiqueta() {
        $("#div_etiqueta").css('display', 'none');
        $("#etiqueta_layout").css('display', 'none');
    }


</script>
<script type="text/javascript">

    function calcTotal() {
        var i, total, n;
        total = 0;

        for (i = 1; i <= fCAD.desc_valor.length; i++) {
            n = converte_numero($("#desc_valor" + i).val());

            if (n == "") {
                n = 0;
                n = parseFloat(n);
            }

            total += n;
        }
        $("#spn_total").text("<%=SIMBOLO_MONETARIO%> " + formata_moeda(total));
    }
</script>
<script type="text/javascript">

    $(function () {

        $("#div_etiqueta").css('filter', 'alpha(opacity=30)');

        calcTotal();

    });
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
<link href="<%=URL_FILE__E_JANELABUSCACEP_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
#loja,#vendedor {
	margin: 4pt 0pt 4pt 10pt;
	vertical-align: top;
	}
#rb_acesso,#rb_status {
	margin-left:10pt;
	}
#rb_estabelecimento 
{
	margin-left:10pt;
}
#lbl_estabelecimento 
{
	font-size:9pt;
}

</style>


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
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
<body onload="focus()">

<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div class="notPrint" id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>
<center>

    <div id="div_etiqueta" style="width:100%;height:100%;position:absolute;left:0;top:0;display:none;background-color:#000;opacity:0.3"></div>
    <div id="etiqueta_layout" style="display:none;z-index:100;position:absolute;width:500px;height:150px;background-color:#fff;left:50%;top:50%;margin-left:-250px;margin-top:20%;box-shadow:2px 2px 2px #000;border-radius:8px;">
        <a href="javascript:fechaEtiqueta();" title="Fechar" style="font-size:21pt;font-weight:bolder;color:#555;position:relative;right:-240px;top:-30px;margin:0">&times;</a>
        <h1 id="etq_nome_fantasia" style="font-size:12pt;margin-top:0px;font-weight:bolder;text-transform:uppercase"></h1>
        <span id="spn_etq_endereco"></span><span>&nbsp;nº&nbsp;</span><span id="spn_etq_numero"></span><span id="separa_complemento">&nbsp;-&nbsp;</span><span id="spn_etq_complemento"></span><span id="separa_bairro">&nbsp;-&nbsp;</span><span id="spn_etq_bairro"></span>
        <br /><span id="spn_etq_cidade"></span><span>&nbsp;-&nbsp;</span><span id="spn_etq_uf"></span><span id="separa_cep">&nbsp;-&nbsp;</span><span id="spn_etq_cep"></span>
        <br /><span id="spn_label_fone">Fone:&nbsp;(</span><span id="spn_etq_ddd_1"></span><span id="spn_fecha_ddd_1">)&nbsp;</span><span id="spn_etq_tel_1"></span>
        <span id="separa_tel">&nbsp;/&nbsp;</span><span id="spn_abre_ddd_2">(</span><span id="spn_etq_ddd_2"></span><span id="spn_fecha_ddd_2">)&nbsp;</span><span id="spn_etq_tel_2"></span>
        <br /><span id="spn_label_email">Email:&nbsp;</span><span id="spn_etq_email"></span>
    </div>

    <div id="caixa-confirmacao" title="Deseja realmente sair?">
  <span id="msgEtq" style="display:none">Você fez alterações nos dados para etiqueta. Tem certeza que deseja sair sem salvá-las?</span>
</div>

<!--  CADASTRO DO ORÇAMENTISTA / INDICADOR -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="center" valign="bottom"><p class="PEDIDO">Consulta de Orçamentista/Indicador Cadastrado<br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>

<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<input type="hidden" name='tipo_PJ_PF' id="tipo_PJ_PF" value='<%=tipo_PJ_PF%>'>
<input type="hidden" name="url_origem" id="url_origem" value='<%=url_origem%>' />
<input type="hidden" name="desc_valor" id="desc_valor" value="0" />

<!-- ************   NOME/RAZÃO SOCIAL   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td align="left" class="MD" width="30%"><p class="R">APELIDO</p><p class="C">
		<input id="id_selecionado" name="id_selecionado" class="TA" value="<%=id_selecionado%>" 
			readonly tabindex=-1
			size="25" style="text-align:center; color:#0000ff"></p></td>
<%if tipo_PJ_PF=ID_PJ then s_label = "RAZÃO SOCIAL" else s_label="NOME" %>
		<td align="left" width="70%"><p class="R"><%=s_label%></p><p class="C">
		<input id="razao_social_nome" name="razao_social_nome" class="TA" type="text" maxlength="60" size="60" 
			value="<%=Trim("" & rs("razao_social_nome"))%>" 
			readonly tabindex=-1
			></p></td>
	</tr>
</table>

<% if blnVisivelIdMagentoB2B then
		s = Trim("" & rs("id_magento_b2b"))
		if s = "0" then s = ""
%>
<!-- ************  ID DO PARCEIRO NO MAGENTO B2B   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td align="left" width="100%"><p class="R">ID MAGENTO B2B</p><p class="C">
		<input id="c_id_magento_b2b" name="c_id_magento_b2b" class="TA" 
			value="<%=s%>" maxlength="60" size="60"
			readonly tabindex=-1
			></p></td>
	</tr>
</table>
<% end if %>

<!-- ************   RESPONSÁVEL PRINCIPAL   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td align="left" width="100%"><p class="R">PRINCIPAL</p><p class="C">
		<input id="c_responsavel_principal" name="c_responsavel_principal" class="TA" 
			value="<%=Trim("" & rs("responsavel_principal"))%>" maxlength="60" size="60"
			readonly tabindex=-1
			></p></td>
	</tr>
</table>

<!-- ************   NOME FANTASIA   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left" width="100%"><p class="R">NOME FANTASIA</p><p class="C">
		<input id="c_nome_fantasia" name="c_nome_fantasia" class="TA" 
			value="<%=Trim("" & rs("nome_fantasia"))%>" maxlength="60" size="60"
			readonly tabindex=-1
			></p></td>
	</tr>
</table>

<!-- ************   CNPJ/CPF + IE/RG   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if tipo_PJ_PF=ID_PJ then s_label = "CNPJ" else s_label="CPF" %>
	<td align="left" class="MD" width="50%"><p class="R"><%=s_label%></p><p class="C">
		<input id="cnpj_cpf" name="cnpj_cpf" class="TA" 
			value="<%=cnpj_cpf_formata(Trim("" & rs("cnpj_cpf")))%>" 
			readonly tabindex=-1
			maxlength="18" size="24" 
		></p></td>
<%if tipo_PJ_PF=ID_PJ then s_label = "IE" else s_label="RG" %>
		<td align="left" width="50%"><p class="R"><%=s_label%></p><p class="C">
		<input id="ie_rg" name="ie_rg" class="TA" type="text" maxlength="20" size="25" 
			value="<%=Trim("" & rs("ie_rg"))%>" 
			readonly tabindex=-1
			></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left" width="100%"><p class="R">ENDEREÇO</p><p class="C">
		<input id="endereco" name="endereco" class="TA" 
			value="<%=Trim("" & rs("endereco"))%>" maxlength="60" style="width:635px;"
			readonly tabindex=-1
			></p></td>
	</tr>
</table>

<!-- ************   Nº/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left" width="50%" class="MD"><p class="R">Nº</p><p class="C">
		<input id="endereco_numero" name="endereco_numero" class="TA" 
			value="<%=Trim("" & rs("endereco_numero"))%>" maxlength="20" style="width:310px;"
			readonly tabindex=-1
			></p></td>
		<td align="left" width="50%"><p class="R">COMPLEMENTO</p><p class="C">
		<input id="endereco_complemento" name="endereco_complemento" class="TA" 
			value="<%=Trim("" & rs("endereco_complemento"))%>" maxlength="60" style="width:310px;"
			readonly tabindex=-1
			></p></td>
	</tr>
</table>

<!-- ************   BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left" width="50%" class="MD"><p class="R">BAIRRO</p><p class="C">
		<input id="bairro" name="bairro" class="TA" 
			value="<%=Trim("" & rs("bairro"))%>" maxlength="72" style="width:310px;"
			readonly tabindex=-1
			></p></td>
		<td align="left" width="50%"><p class="R">CIDADE</p><p class="C">
		<input id="cidade" name="cidade" class="TA" 
			value="<%=Trim("" & rs("cidade"))%>" maxlength="60" style="width:310px;"
			readonly tabindex=-1
			></p></td>
	</tr>
</table>

<!-- ************   UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left" class="MD" width="50%"><p class="R">UF</p><p class="C">
		<input id="uf" name="uf" class="TA" value="<%=Trim("" & rs("uf"))%>" 
			maxlength="2" size="3" 
			readonly tabindex=-1
			></p></td>
		<td align="left"><p class="R">CEP</p><p class="C">
		<input id="cep" name="cep" class="TA" value="<%=cep_formata(Trim("" & rs("cep")))%>" 
			maxlength="9" size="11" 
			readonly tabindex=-1
			></p></td>
	</tr>
</table>

<!-- ************   DDD/TELEFONE/FAX/NEXTEL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left" class="MD" width="15%"><p class="R">DDD</p><p class="C">
		<input id="ddd" name="ddd" class="TA" value="<%=Trim("" & rs("ddd"))%>" 
			maxlength="4" size="5" 
			readonly tabindex=-1
			></p></td>
		<td align="left" width="25%" class="MD"><p class="R">TELEFONE</p><p class="C">
		<input id="telefone" name="telefone" class="TA" 
			value="<%=telefone_formata(Trim("" & rs("telefone")))%>" 
			maxlength="11" size="12" 
			readonly tabindex=-1
			></p></td>
		<td align="left" width="25%" class="MD"><p class="R">FAX</p><p class="C">
		<input id="fax" name="fax" class="TA" 
			value="<%=telefone_formata(Trim("" & rs("fax")))%>" maxlength="11" size="12" 
			readonly tabindex=-1
			></p></td>
		<td align="left"><p class="R">NEXTEL</p><p class="C">
		<input id="c_nextel" name="c_nextel" class="TA" 
			value="<%=Trim("" & rs("nextel"))%>" maxlength="15" size="12" 
			readonly tabindex=-1
			></p></td>
	</tr>
</table>

<!-- ************   TEL CEL / CONTATO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left" width="15%" class="MD" nowrap><p class="R">DDD (CEL)</p><p class="C">
		<input id="ddd_cel" name="ddd_cel" class="TA" value="<%=Trim("" & rs("ddd_cel"))%>" 
			maxlength="2" size="3" 
			readonly tabindex=-1
			></p></td>
		<td align="left" width="25%" class="MD"><p class="R">TELEFONE (CEL)</p><p class="C">
		<input id="tel_cel" name="tel_cel" class="TA" value="<%=telefone_formata(Trim("" & rs("tel_cel")))%>" 
			maxlength="10" size="11" 
			readonly tabindex=-1
			></p></td>
		<td align="left"><p class="R">CONTATO</p><p class="C">
		<input id="contato" name="contato" class="TA" value="<%=Trim("" & rs("contato"))%>" 
			maxlength="40" size="55" 
			readonly tabindex=-1
			></p></td>
	</tr>
</table>

<!-- ************   BANCO/AGÊNCIA/CONTA   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>

		<td width="15%" class="MD" nowrap align="left"><p class="R">BANCO</p><p class="C"><input id="banco" name="banco" class="TA" value="<%=rs("banco")%>" maxlength="4" size="3" readonly tabindex=-1></p></td>

		<td width="17%" class="MD" align="left"><p class="R">AGÊNCIA</p><p class="C"><input id="agencia" name="agencia" class="TA" value="<%=rs("agencia")%>" maxlength="8" size="5" readonly tabindex=-1></p></td>

		<td width="5%" class="MD" align="left"><p class="R">DÍG.</p><p class="C"><input id="agencia_dv" name="agencia_dv" class="TA" value="<%=rs("agencia_dv")%>" maxlength="1" size="1" readonly tabindex=-1></p></td>

        <td width="15%" class="MD" align="left"><p class="R">TIPO OPERAÇÃO</p><p class="C"><input id="tipo_operacao" name="tipo_operacao" class="TA" value="<%=rs("conta_operacao")%>" maxlength="12" size="12" readonly tabindex=-1></p></td>

		<td width="17%" class="MD" align="left"><p class="R">CONTA</p><p class="C"><input id="conta" name="conta" class="TA" value="<%=rs("conta")%>" maxlength="12" size="12" readonly tabindex=-1></p></td>

		<td width="5%" class="MD" align="left"><p class="R">DÍG.</p><p class="C"><input id="conta_dv" name="conta_dv" class="TA" value="<%=rs("conta_dv")%>" maxlength="1" size="1" readonly tabindex=-1></p></td>

		<td width="15%" align="left"><p class="R">TIPO CONTA</p><p class="C">
            <%s_selected="" %>
            <select name="tipo_conta" id="tipo_conta" disabled>
                <%if Trim("" & rs("tipo_conta")) ="" then  s_selected=" selected"%>
                <option value=""<%=s_selected%>>&nbsp;</option>
                <%s_selected=""
                    if Trim("" & rs("tipo_conta"))="C" then s_selected=" selected" %>
                <option value="C"<%=s_selected%>>Corrente</option>
                <%s_selected=""
                    if Trim("" & rs("tipo_conta"))="P" then s_selected=" selected" %>
                <option value="P"<%=s_selected%>>Poupança</option>
            </select> </p></td>

	</tr>
</table>

<!-- ************   FAVORECIDO / CNPJ/CPF FAVORECIDO    *******************  -->
<table width="649" class="QS" cellspacing="0">
    <tr>
		<td class="MD" width="70%" align="left"><p class="R">FAVORECIDO</p><p class="C"><input id="favorecido" name="favorecido" class="TA" size="60" value="<%=rs("favorecido")%>" readonly tabindex=-1></p></td>
		<td width="30%" align="left"><p class="R">CPF/CNPJ DO FAVORECIDO</p><p class="C"><input id="favorecido_cnpjcpf" name="favorecido_cnpjcpf" class="TA" readonly tabindex=-1 maxlength="18" size="25" value="<%=cnpj_cpf_formata(Trim("" & rs("favorecido_cnpj_cpf")))%>"></p></td>
    </tr>
</table>

<!-- ************   DADOS P/ PAGTO COMISSÃO: CARTÃO / NFSe   *******************  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td width="100%" style="padding-bottom:10px;" align="left">
			<p class="R" style="padding-bottom:8px;">PAGAMENTO DA COMISSÃO</p>
			<table width="607" class="Q" cellspacing="0" style="margin-left:20px;">
			<tr>
				<td width="100%">
					<p class="R">CARTÃO</p>
					<table width="100%" border="0">
					<tr>
						<td colspan="2" align="left">
							<input type="checkbox" id="ckb_comissao_cartao_status" name="ckb_comissao_cartao_status" value="ON" class="TA CKB_COM_CAR" disabled
								<% if Trim("" & rs("comissao_cartao_status")) = "1" then Response.Write " checked"%>
								/><span id="spn_comissao_cartao_status" class="C" style="cursor:default;">Pagamento Via Cartão</span>
						</td>
					</tr>
					<tr>
						<td style="width:20px;">&nbsp;</td>
						<td width="95%" style="padding-bottom:8px;padding-right:12px;">
							<table class="Q" width="100%" cellspacing="0">
								<tr>
									<td>
									<p class="R">CPF</p>
									<input type="text" id="c_comissao_cartao_cpf" name="c_comissao_cartao_cpf" class="TA" value="<%=cnpj_cpf_formata(Trim("" & rs("comissao_cartao_cpf")))%>" maxlength="14" size="18" 
										readonly tabindex=-1 />
									</td>
								</tr>
								<tr>
									<td class="MC">
									<p class="R">TITULAR DO CARTÃO</p>
									<input type="text" id="c_comissao_cartao_titular" name="c_comissao_cartao_titular" class="TA" value="<%=Trim("" & rs("comissao_cartao_titular"))%>" maxlength="60" size="70"
										readonly tabindex=-1 />
									</td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td width="100%" class="MC">
					<p class="R">EMITENTE NFSe</p>
					<table width="100%" border="0">
					<tr>
						<td style="width:20px;">&nbsp;</td>
						<td width="95%" style="padding-bottom:8px;padding-right:12px;">
							<table class="Q" width="100%" cellspacing="0">
								<tr>
									<td>
									<p class="R">CNPJ</p>
									<input type="text" id="c_comissao_NFSe_cnpj" name="c_comissao_NFSe_cnpj" class="TA" value="<%=cnpj_cpf_formata(Trim("" & rs("comissao_NFSe_cnpj")))%>" maxlength="18" size="24"
										readonly tabindex=-1 />
									</td>
								</tr>
								<tr>
									<td class="MC">
									<p class="R">RAZÃO SOCIAL DO EMITENTE</p>
									<input type="text" id="c_comissao_NFSe_razao_social" name="c_comissao_NFSe_razao_social" class="TA" value="<%=Trim("" & rs("comissao_NFSe_razao_social"))%>" maxlength="60" size="70"
										readonly tabindex=-1 />
									</td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>

<!-- ************   LOJA (DO ORÇAMENTISTA)   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left"><p class="R">LOJA&nbsp;&nbsp;(ORÇAMENTISTAS)</p><p class="C">
			<select id="loja" name="loja" style="width:490px;" disabled tabindex=-1>
			  <% =loja_do_orcamentista_monta_itens_select(Trim("" & rs("loja"))) %>
			</select>
		</p></td>
	</tr>
</table>

<!-- ************   ATENDIDO PELO VENDEDOR (P/ INDICADORES)   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td align="left"><p class="R">ATENDIDO POR&nbsp;&nbsp;(INDICADORES)</p><p class="C">
			<select id="vendedor" name="vendedor" style="width:490px;" disabled tabindex=-1>
			  <% =vendedor_do_indicador_monta_itens_select(Trim("" & rs("vendedor"))) %>
			</select>
		</p></td>
	</tr>
</table>

<!-- ************   ACESSO AO SISTEMA/STATUS   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%s_parametro=Cstr(rs("hab_acesso_sistema"))%>
		<td align="left" width="35%" class="MD"><p class="R">ACESSO AO SISTEMA</p><p class="C">
			<input type="radio" id="rb_acesso" name="rb_acesso" value="1" 
				class="TA"<%if s_parametro = "1" then Response.Write(" checked")%>
				disabled tabindex=-1
				><span style="cursor:default; color:#006600">Liberado</span>
			<br><input type="radio" id="rb_acesso" name="rb_acesso" value="0" 
				class="TA"<%if (s_parametro<>"1") And (s_parametro<>"") then Response.Write(" checked")%>
				disabled tabindex=-1
				><span style="cursor:default; color:#ff0000">Bloqueado</span>
			</p></td>
<%s_parametro=Trim("" & rs("status"))%>
		<td align="left" width="35%" class="MD"><p class="R">STATUS</p><p class="C">
			<input type="radio" id="rb_status" name="rb_status" value="A" 
				class="TA"<%if s_parametro = "A" then Response.Write(" checked")%>
				disabled tabindex=-1
				><span style="cursor:default; color:#006600">Ativo</span>
			<br><input type="radio" id="rb_status" name="rb_status" value="I" 
				class="TA"<%if (s_parametro<>"A") And (s_parametro<>"") then Response.Write(" checked")%>
				disabled tabindex=-1
				><span style="cursor:default; color:#ff0000">Inativo</span>
			</p></td>
<%s_parametro=Trim("" & rs("desempenho_nota"))%>
		<td align="left" width="30%" valign="Top"><p class="R">AVALIAÇÃO DESEMPENHO</p><p class="C">
			<select id="c_desempenho_nota" name="c_desempenho_nota" style="margin-top:4pt; margin-bottom:4pt;width:45px;" disabled>
				<% =desempenho_nota_monta_itens_select(s_parametro) %>
			</select>
		</p></td>
	</tr>
</table>

<!-- ************   SENHA / CONFIRMAÇÃO DA SENHA   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	senha_descripto= ""
	s = Trim("" & rs("datastamp"))
	chave = gera_chave(FATOR_BD)
	decodifica_dado s, senha_descripto, chave
%>
		<td class="MD" width="50%" align="left"><p class="R">SENHA</p><p class="C"><input id="senha" name="senha" class="TA" type="password" maxlength="15" size="18" value="<%=senha_descripto%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.senha2.focus();"></p></td>
		<td width="50%" align="left"><p class="R">SENHA (CONFIRMAÇÃO)</p><p class="C"><input id="senha2" name="senha2" class="TA" type="password" maxlength="15" size="18" value="<%=senha_descripto%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.loja.focus();"></p></td>
</table>

<!-- ************   LOGIN BLOQUEADO AUTOMATICAMENTE?   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	s = "&nbsp;"
	s_color = "black"
	if rs("StLoginBloqueadoAutomatico") <> 0 then
		s = "Bloqueado em " & formata_data_hora_sem_seg(rs("DataHoraBloqueadoAutomatico")) & " (" & Trim("" & rs("QtdeConsecutivaFalhaLogin")) & " tentativas consecutivas com senha errada)"
		s_color = "red"
		end if
%>
		<td width="100%" align="left">
		<p class="R">LOGIN BLOQUEADO AUTOMATICAMENTE</p>
		<p class="C" id="pMsgStLoginBloqueadoAutomatico" style="color:<%=s_color%>;"><%=s%></p>
		</td>
	</tr>
</table>

<!-- ************   PERCENTUAL DE DESÁGIO DO RA / VALOR DA META   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%s=formata_perc(rs("perc_desagio_RA"))%>
		<td align="left" width="50%" class="MD"><p class="R">PERCENTUAL DESÁGIO DO RA&nbsp;&nbsp;(INDICADORES)</p><p class="C">
			<input id="c_perc_desagio_RA" name="c_perc_desagio_RA" 
				class="TA" value="<%=s%>" 
				maxlength="5" style="text-align:right;width:60px;"
				readonly tabindex=-1
				><span style="margin-left:2px;">%</span>
		</p></td>

<%s=formata_moeda(rs("vl_limite_mensal"))%>
<input type="hidden" name="c_vl_limite_mensal" id="c_vl_limite_mensal" value="<%=s%>">

<%s=formata_moeda(rs("vl_meta"))%>
		<td align="left" width="50%"><p class="R">VL META&nbsp;&nbsp;(<%=SIMBOLO_MONETARIO%>)</p><p class="C">
			<input id="c_vl_meta" name="c_vl_meta" 
				class="TA" value="<%=s%>" 
				maxlength="18" style="text-align:left;width:180px;"
				readonly tabindex=-1
				>
		</p></td>
	</tr>
</table>

<!-- ************   E-MAILS   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left"><p class="R">E-MAIL (1)</p><p class="C">
			<input id="c_email" name="c_email" class="TA" value="<%=Trim("" & rs("email"))%>" 
				maxlength="60" style="text-align:left;" size="74"
				readonly tabindex=-1
				>
		</p></td>
	</tr>
	<tr>
		<td align="left" class="MC"><p class="R">E-MAIL (2)</p><p class="C">
			<input id="c_email2" name="c_email2" class="TA" value="<%=Trim("" & rs("email2"))%>" 
				maxlength="60" style="text-align:left;" size="74"
				readonly tabindex=-1
				>
		</p></td>
	</tr>
	<tr>
		<td align="left" class="MC"><p class="R">E-MAIL (3)</p><p class="C">
			<input id="c_email3" name="c_email3" class="TA" value="<%=Trim("" & rs("email3"))%>" 
				maxlength="60" style="text-align:left;" size="74"
				readonly tabindex=-1
				>
		</p></td>
	</tr>
</table>

<!-- ************   TIPO DE ESTABELECIMENTO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%s_parametro=Trim("" & rs("tipo_estabelecimento"))%>
		<td align="left"><p class="R">ESTABELECIMENTO</p><p class="C">
			<input type="radio" id="rb_estabelecimento" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__CASA%>" 
				class="TA"<%if s_parametro = COD_PARCEIRO_TIPO_ESTABELECIMENTO__CASA then Response.Write(" checked")%>
				disabled tabindex=-1
				><span id="lbl_estabelecimento" style="cursor:default;">Casa</span>
			<br><input type="radio" id="rb_estabelecimento" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__ESCRITORIO%>" 
				class="TA"<%if s_parametro = COD_PARCEIRO_TIPO_ESTABELECIMENTO__ESCRITORIO then Response.Write(" checked")%>
				disabled tabindex=-1
				><span id="lbl_estabelecimento" style="cursor:default;">Escritório</span>
			<br><input type="radio" id="rb_estabelecimento" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__LOJA%>" 
				class="TA"<%if s_parametro = COD_PARCEIRO_TIPO_ESTABELECIMENTO__LOJA then Response.Write(" checked")%>
				disabled tabindex=-1
				><span id="lbl_estabelecimento" style="cursor:default;">Loja</span>
			<br><input type="radio" id="rb_estabelecimento" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__OFICINA%>" 
				class="TA"<%if s_parametro = COD_PARCEIRO_TIPO_ESTABELECIMENTO__OFICINA then Response.Write(" checked")%>
				disabled tabindex=-1
				><span id="lbl_estabelecimento" style="cursor:default;">Oficina</span>
			</p></td>
	</tr>
</table>

<!-- ************   CAPTADOR   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%s=Trim("" & rs("captador"))%>
		<td align="left"><p class="R">CAPTADOR</p><p class="C">
			<select id="c_captador" name="c_captador" style="margin-top:4pt; margin-bottom:4pt;" disabled tabindex=-1>
				<%=captadores_monta_itens_select(s)%>
			</select>
		</p></td>
	</tr>
</table>

<!-- ************   VENDEDORES   **************** -->

<% set rs2 = cn.Execute("SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_CONTATOS WHERE (indicador='" & id_selecionado & "') ORDER BY dt_cadastro DESC") %>
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left" class="MB" colspan="2"><p class="R">VENDEDORES</p></td>
	</tr>
    <tr>
        <td align="left"><p class="R" style="margin-bottom:3px;margin-top:3px">NOME</p></td>
        <td align="left"><p class="R" style="margin-bottom:3px;margin-top:3px;margin-right:5px">CADASTRO</p></td>
    </tr>
<% if rs2.Eof then %>
    <tr>
        <td align="left" colspan="2"><p class="R">&nbsp;</p></td>
    </tr>
<% end if %>

<% i = 0
    do while Not rs2.Eof
    i = i + 1
%>
    <tr>
        <td align="left" width="40%">
            <input id="c_indicador_contato_<%=i%>" name="c_indicador_contato_<%=i%>" class="TA" value='<%=Trim("" & rs2("nome"))%>' style="text-align: left;margin-left: 5px;" size="40" readonly tabindex=-1 />
        </td>
        <td align="left">
            <input id="c_indicador_contato_data_<%=i%>" name="c_indicador_contato_data_<%=i%>" class="TA" value='<%=formata_data(Trim("" & rs2("dt_cadastro")))%>' style="text-align: left;margin-left: 5px;" size="20" readonly tabindex=-1 />
        </td>
    </tr>
<% rs2.MoveNext
loop %>
</table>


<!-- ************   OBS   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%s=Trim("" & rs("obs"))%>
		<td align="left"><p class="R">OBSERVAÇÕES</p><p class="C">
			<textarea name="c_obs" id="c_obs" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS_ORCAMENTISTA_INDICADOR)%>" 
				style="width:635px;margin-left:1pt;"
				readonly tabindex=-1
				><%=s%></textarea>
		</p></td>
	</tr>
</table>

<!-- ************   CHECADO / PARCEIRO DESDE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%s_parametro=Cstr(rs("checado_status"))%>
		<td align="left" width="50%" class="MD" valign="Top"><p class="R">CHECADO</p>
			<%if s_parametro = "1" then %>
				<span class="C" style="color:#006600;">SIM (checado) por <%=Trim("" & rs("checado_usuario")) & " - " & formata_data_hora(rs("checado_data"))%></span>
			<% else %>
				<span class="C" style="color:#ff0000;">NÃO (não-checado)</span>
			<% end if %>
			</td>
		<td align="left" width="50%" valign="Top"><p class="R">PARCEIRO DESDE</p>
			<span class="C"><%=formata_data(rs("dt_cadastro"))%></span>
		</td>
	</tr>
</table>

<!-- ************   DADOS PARA ETIQUETA   **************** -->
<br />
<table width="649" class="notPrint" style="border: 1pt solid #C0C0C0;border-top: 0pt;margin: 0pt;" cellSpacing="0" id="tableDadosEtiqueta">
	<tr>
		<td align="left" class="MC" valign="middle"><p class="R">DADOS PARA ETIQUETA</p></td>
		</tr>
</table>


<table id="Etq1" width="649" class="notPrint" style="border: 1pt solid #C0C0C0;border-top: 0pt;margin: 0pt;" cellSpacing="0">
	<tr>
		<td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C"><input readonly tabindex="-1" id="etq_endereco" name="etq_endereco" class="TA" value="<%=rs("etq_endereco")%>" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true)) fCAD.etq_endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>


<table id="Etq2" width="649" class="notPrint" style="border: 1pt solid #C0C0C0;border-top: 0pt;margin: 0pt;" cellSpacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">Nº</p><p class="C">
		
		<input id="etq_endereco_numero" readonly tabindex="-1" name="etq_endereco_numero" class="TA" value="<%=rs("etq_endereco_numero")%>" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.etq_endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<input id="etq_endereco_complemento" readonly tabindex="-1" name="etq_endereco_complemento" class="TA" value="<%=rs("etq_endereco_complemento")%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.etq_bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>


<table id="Etq3" width="649" class="notPrint" style="border: 1pt solid #C0C0C0;border-top: 0pt;margin: 0pt;" cellSpacing="0">
	<tr>
		<td width="50%" class="MD" align="left"><p class="R">BAIRRO</p><p class="C"><input id="etq_bairro" name="etq_bairro" readonly tabindex="-1" class="TA" value="<%=rs("etq_bairro")%>" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.etq_cidade.focus(); filtra_nome_identificador();"></p></td>
		<td width="50%" align="left"><p class="R">CIDADE</p><p class="C"><input id="etq_cidade" name="etq_cidade" readonly tabindex="-1" class="TA" value="<%=rs("etq_cidade")%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.etq_uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>


<table id="Etq4" width="649" class="notPrint" style="border: 1pt solid #C0C0C0;border-top: 0pt;margin: 0pt;" cellSpacing="0">
	<tr>
		<td class="MD"  width="50%" align="left"><p class="R">UF</p><p class="C"><input id="etq_uf" name="etq_uf" readonly tabindex="-1" class="TA" value="<%=rs("etq_uf")%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && uf_ok(this.value)) fCAD.etq_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
		<td width="50%" align="left"><p class="R">CEP</p><p class="C"><input id="etq_cep" name="etq_cep" class="TA" readonly tabindex="-1" value="<%=cep_formata(rs("etq_cep"))%>" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) fCAD.etq_ddd_1.focus(); filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
		
	</tr>
</table>


<table id="Etq5" width="649" class="notPrint" style="border: 1pt solid #C0C0C0;border-top: 0pt;margin: 0pt;" cellSpacing="0">
	<tr>
		<td width="15%" class="MD" align="left"><p class="R">DDD</p><p class="C"><input id="etq_ddd_1" name="etq_ddd_1" readonly tabindex="-1" class="TA" value="<%=rs("etq_ddd_1")%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.etq_tel_1.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
		<td width="35%" class="MD" align="left"><p class="R">TELEFONE</p><p class="C"><input id="etq_tel_1" name="etq_tel_1" readonly tabindex="-1" class="TA" value="<%=telefone_formata(rs("etq_tel_1"))%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.etq_email.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
		
		<td width="15%" class="MD" align="left" nowrap><p class="R">DDD</p><p class="C"><input id="etq_ddd_2" name="etq_ddd_2" readonly tabindex="-1" class="TA" value="<%=rs("etq_ddd_2")%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.etq_tel_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
		<td width="35%" align="left"><p class="R">TELEFONE</p><p class="C"><input id="etq_tel_2" name="etq_tel_2" class="TA" readonly tabindex="-1" value="<%=telefone_formata(rs("etq_tel_2"))%>" maxlength="10" size="11" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.contato.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	</tr>
</table>


<table id="Etq6" width="649" class="notPrint" style="border: 1pt solid #C0C0C0;border-top: 0pt;margin: 0pt;" cellSpacing="0">
	<tr>
        <td width="90%" align="left"><p class="R">E-MAIL</p><p class="C">
			<input id="etq_email" name="etq_email" class="TA" value="<%=rs("etq_email")%>" maxlength="60" 
			style="text-align:left;" size="50" readonly tabindex="-1"
			onkeypress="if (digitou_enter(true)) fCAD.etq_ddd_2.focus(); filtra_email();"
			onblur="this.value=trim(this.value);">
		</p></td>
        
        <td width="10%" align="center"><a href="javascript:AbreJanelaEtiqueta()"><img src="../imagem/lupa_20x20.png" style="width:20px;height:20px" title="Gerar etiqueta" border="0"></a></td>
	</tr>
</table>
<table class="notPrint" width="649" cellspacing="0" cellpadding="1">
   
    <tr>
		<td colspan="4" style="padding:0px;" align="left">
			<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
			<td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="DadosEtiquetaAlteraImpressao" id="DadosEtiquetaAlteraImpressao" href="javascript:fCADDadosEtiquetaAlteraImpressao(fCAD)" title="configura as dados de etiqueta para serem impressas ou não"><img name="imgPrinterDadosEtiqueta" id="imgPrinterDadosEtiqueta" src="../botao/PrinterError.png" border="0"></a></td>
			<td align="left">&nbsp;</td>
			
			</tr>
			</table>
		</td>
	</tr>
</table>

<!-- ************   TABELA DE DESCONTOS   **************** -->

<% set rs2 = cn.Execute("SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (apelido='" & id_selecionado & "') ORDER BY ordenacao") %>
<% dim inc, sid
    inc = 1 
   s = ""
   sid="-1"
    %>
<br />
<table id="tableDescontos" width="649" class="notPrint" style="border: 1pt solid #C0C0C0;border-top: 0pt;margin: 0pt;" cellSpacing="0">
	<tr>
		<td class="MC" align="left" valign="middle"><p class="R">TABELA DE DESCONTOS</p></td>
		</tr>
</table>

<table id="tblDesc" width="649" class="notPrint" style="border: 1pt solid #C0C0C0;border-top: 0pt;margin: 0pt;" cellSpacing="0">
	<tr>

		<td width="490px" align="left"><p class="R" style="margin-bottom:3px;margin-top:3px">DESCRIÇÃO</p></td>
        <td width="129px" align="right"><p class="R" style="margin-bottom:3px;margin-top:3px;margin-right:5px">VALOR</p></td>
    </tr>
    
    <% do while Not rs2.Eof %>
    <tr>
        <td>
            <p class="C"><input id="desc_descricao<%=inc%>" name="desc_descricao" class="TA" value="<%=rs2("descricao")%>" maxlength="100" style="width:490px;border:1px solid #c0c0c0" readonly></p>
		</td>
        <td><input type="text" name="desc_valor" style="display:none" />
            <p class="C">R$&nbsp;<input id="desc_valor<%=inc%>" name="desc_valor" class="TA" value="<%=formata_moeda(rs2("valor"))%>" maxlength="10" style="width:107px;border:1px solid #c0c0c0;text-align:right" readonly></p>
            <input type="hidden" name="id_desc" id="id_desc_<%=inc%>" value="<%=rs2("id")%>" />
        </td>
    </tr>
    <% inc = inc + 1
         rs2.MoveNext
         loop %>

    <tr>
        <td align="right"><span class="C">TOTAL:</span></td>
        <td align="right"><span id="spn_total" class="C"></span></td>
    </tr>
 
</table>
<table class="notPrint" width="649" cellspacing="0" cellpadding="1">
   
    <tr>
		<td colspan="4" style="padding:0px;" align="left">
			<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
			<td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="TblDescontosAlteraImpressao" id="TblDescontosAlteraImpressao" href="javascript:fCADTblDescontosAlteraImpressao(fCAD)" title="configura a tabela de descontos para ser impressa ou não"><img name="imgPrinterTblDescontos" id="imgPrinterTblDescontos" src="../botao/PrinterError.png" border="0"></a></td>
			<td align="left">&nbsp;</td>
			
			</tr>
			</table>
		</td>
	</tr>
</table>

<!-- ************   BLOCO DE RELACIONAMENTO   *********** -->

<%
 dim v_meses(12), mes, qtde_mes

 s = "SELECT " & _
			"YEAR(dt_cadastro) AS ano, COUNT(YEAR(dt_cadastro)) AS qtde_ano " & _
	   " FROM t_ORCAMENTISTA_E_INDICADOR_BLOCO_NOTAS" & _
	   " WHERE" & _
			" (apelido = '" & id_selecionado & "')" & _
			" AND dt_cadastro <= GETDATE() " & _
			" AND (anulado_status = 0)" & _
			" GROUP BY YEAR(dt_cadastro) " & _
			" ORDER BY YEAR(dt_cadastro)"
			
	set rs = cn.execute(s) 
	
%>

<br />
<table id="tableBlocoNotas" class="notPrint" width="649" cellspacing="0" cellpadding="1">
<tr>
	<td align="left" class="ME MD MC MB" align="left"><span class="Rf">BLOCO DE RELACIONAMENTO</span></td>    
</tr><tr>
        <td style="padding:0">

<% 	
	do while Not rs.Eof
    
%>
        <table id='tableBlocoAno<%=rs("ano") %>' width="649" cellspacing="0" cellpadding="1">
		<tr><td class="ME MB" valign="middle" width="10" style="background-color:#eee"><a href='javascript:mostraOcultaMeses(<%=rs("ano")%>)'><img id='img<%=rs("ano")%>' src="../imagem/plus.gif" style="border:0" title="expandir"></a></td>
		<td colspan="2" class="MD MB" align="left" style="background-color:#eee"><a href='javascript:mostraOcultaMeses(<%=rs("ano")%>)'><span class="Rf"><%=rs("ano") & "&nbsp;(" & rs("qtde_ano") & ")" %></span></a></td></tr>
		</table>
		<% s = "SELECT " & _
	    "MONTH(dt_cadastro) AS mes, COUNT(MONTH(dt_cadastro)) AS qtde_mes FROM t_ORCAMENTISTA_E_INDICADOR_BLOCO_NOTAS " & _
	    "WHERE apelido='" & id_selecionado & "' AND YEAR(dt_cadastro)='" & rs("ano") & "' " & _
	    "GROUP BY MONTH(dt_cadastro)"
	
	set rs2 = cn.Execute(s)%>
		
	<%
	v_meses(1) = 0
    v_meses(2) = 0
    v_meses(3) = 0
    v_meses(4) = 0
    v_meses(5) = 0
    v_meses(6) = 0
    v_meses(7) = 0
    v_meses(8) = 0
    v_meses(9) = 0
    v_meses(10) = 0
    v_meses(11) = 0
    v_meses(12) = 0
		 do while Not rs2.Eof 
		 v_meses(rs2("mes")) = rs2("qtde_mes") 
		 
		rs2.MoveNext
		    loop 
		    
		        for cont=1 to UBound(v_meses)
		          if rs("ano") = 2015 And cont < 7 then cont=7
		                		       
		     %>
		 <table class='tableBlocoMes<%=rs("ano")%>' width="649" cellspacing="0" cellpadding="1" style="display:none">
		<tr>
		    <td class="ME MB" valign="middle" width="5">&nbsp;</td><td class="MB" valign="middle" width="10"><a href='javascript:mostraOcultaNotas(<%=rs("ano")%>,<%=cont%>)'><img id='img<%=rs("ano") & cont%>' src="../imagem/plus.gif" title="mostrar anotações" style="border:0" class='imgClasse<%=rs("ano")%>' /></a></td>
		    <td class="MD MB" align="left"><a href='javascript:mostraOcultaNotas(<%=rs("ano")%>,<%=cont%>)'><span class="Rf"><%=mes_por_extenso(cont,true) & "&nbsp;(" & v_meses(cont) & ")" %></span></a></td>
		</tr>
		</table>
		<table id="<%=rs("ano") & cont%>" class="classeFecha<%=rs("ano")%>" cellspacing="0" cellpadding="1" style="display:none;width:649px">
		</table>
		<%	if rs("ano") = Year(Date) And cont = Month(Date) then Exit For					
		next %>
		
		 
<%
		rs.MoveNext
		loop
%>
    
    </table>
    <table class="notPrint" width="649" cellspacing="0" cellpadding="1">
   
    <tr>
		<td colspan="4" style="padding:0px;" align="left">
			<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
			<td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="bBlocoNotasAlteraImpressao" id="bBlocoNotasAlteraImpressao" href="javascript:fCADBlocoNotasAlteraImpressao(fCAD)" title="configura as mensagens do bloco de notas para serem impressas ou não"><img name="imgPrinterBlocoNotas" id="imgPrinterBlocoNotas" src="../botao/PrinterError.png" border="0"></a></td>
			<td align="left">&nbsp;</td>
			<td align="center" class="ME MB MD" style="width:45px;padding:2px;">
				<a name="bBlocoNotasAdiciona" id="bBlocoNotasAdiciona" href="javascript:fCADAdicionaBlocoNotas(fCAD)" title="Adiciona um novo bloco de notas"><img src="../botao/Add.png" border="0"></a>
			</td>
			</tr>
			</table>
		</td>
	</tr>
</table>

<!-- **************    HISTÓRICO DE ALTERAÇÕES NO CADASTRO   ******************** -->
<br />
<table id="tableHistorico" class="notPrint" width="649" cellspacing="0" cellpadding="1">
    <tr>
        <td colspan="3" class="ME MD MC MB" align="left"><span class="Rf">HISTÓRICO DE ALTERAÇÕES NO CADASTRO</span></td>
    </tr>
    <% 
        s = "SELECT " & _
                "*" & _
                " FROM t_ORCAMENTISTA_E_INDICADOR_LOG" & _
                " WHERE (apelido = '" & id_selecionado & "')" & _
                " ORDER BY dt_hr_cadastro"

        set rs = cn.Execute(s)
        if rs.Eof then  %>
    <tr class="notVisible">
		<td colspan="4" class="ME MD MB" align="left">&nbsp;</td>
	</tr>
    
    <% end if 
        do while Not rs.Eof
    %>
    <tr>
        <td class="C ME MD MB" style="width:60px" align="center" valign="top"><%=formata_data_hora(rs("dt_hr_cadastro"))%></td>
        <td class="C MD MB" style="width:80px" align="center" valign="top"><%
            s = rs("usuario")
            if Trim("" & (rs("loja")) <> "") then s = s & " (Loja&nbsp;" & Trim("" & rs("loja")) & ")"
            Response.Write s
             %></td>
        <td class="C MD MB" align="left" valign="top"><%=substitui_caracteres(rs("mensagem"), "|", "&nbsp;&nbsp;&nbsp;")%></td>
    <% rs.MoveNext
        loop %>
    </tr>

</table>
<table class="notPrint" width="649" cellspacing="0" cellpadding="1">
   
    <tr>
		<td colspan="4" style="padding:0px;" align="left">
			<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
			<td align="center" class="ME MB MD" style="width:45px;padding:2px;"><a name="HistoricotaAlteraImpressao" id="HistoricoAlteraImpressao" href="javascript:fCADHistoricoAlteraImpressao(fCAD)" title="configura o histórico de alterações para ser impresso ou não"><img name="imgPrinterHistorico" id="imgPrinterHistorico" src="../botao/PrinterError.png" border="0"></a></td>
			<td align="left">&nbsp;</td>
			
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
<td align="center"><a href='<% if url_back <> "" then Response.Write (url_origem) else Response.Write ("javascript:history.back()")%>' title="retorna para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	rs.Close
	set rs = nothing

    rs2.Close
	set rs2 = nothing
	
	cn.Close
	set cn = nothing
%>