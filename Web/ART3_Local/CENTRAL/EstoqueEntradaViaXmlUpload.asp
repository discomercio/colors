<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ==================================
'	  EstoqueEntradaViaXmlUpload.asp
'     ==================================
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

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_ENTRADA_MERCADORIAS_ESTOQUE, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	VERIFICA SE O TIPO DE UPLOAD FOI PREVIAMENTE SELECIONADO
'   (em caso de retorno da página de edição)
    dim c_op_upload
    c_op_upload = Trim(Request("c_op_upload"))
    if c_op_upload = "" then c_op_upload = "N"

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s, s_sessionToken
	s_sessionToken = ""
	s = "SELECT Convert(varchar(36), SessionTokenModuloCentral) AS SessionTokenModuloCentral FROM t_USUARIO WHERE (usuario = '" & usuario & "')"
    if rs.State <> 0 then rs.Close
    rs.Open s,cn
	if Not rs.Eof then s_sessionToken = Trim("" & rs("SessionTokenModuloCentral"))

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
	var nfe;
	var nfe2;
	var jsonFileInfo;
	var jsonFileInfo2;
	var serverVariableUrl;

	serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
	serverVariableUrl = serverVariableUrl.toUpperCase();
	serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));

//Dynamically assign height
function sizeDivAjaxRunning() {
	var newTop = $(window).scrollTop() + "px";
	$("#divAjaxRunning").css("top", newTop);
}

$(function () {
	$("#txtArquivoEnviado").hide();
	$("#imgUploadOk").hide();
	$("#imgUploadPending").show();
	$("#divAjaxRunning").hide();

	$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8

	$("#divArquivoXML2").hide();
	$("#linhaStatusUpload2").hide();
	$("#linhaArquivoUpload2").hide();

	//Every resize of window
	$(window).resize(function() {
		sizeDivAjaxRunning();
	});

	//Every scroll of window
	$(window).scroll(function() {
		sizeDivAjaxRunning();
	});

	// Trata o problema em que os campos do formulário são limpos após retornar à esta página c/ o history.back() pela 2ª vez quando ocorre erro de consistência
	if (trim(fUPLOAD.c_FormFieldValues.value) != "") {
		stringToForm(fUPLOAD.c_FormFieldValues.value, $('#fUPLOAD'));
		if ($("#c_status_upload_ok").val() == "1") {
			$("#txtArquivoEnviado").show();
			$("#imgUploadOk").show();
			$("#imgUploadPending").hide();
		}
	}

    if (trim(fESTOQ.c_FormFieldValues.value) != "")
    {
    	stringToForm(fESTOQ.c_FormFieldValues.value, $('#fESTOQ'));
    }

	$('#fUPLOAD').submit(function (e) {
		var form = $(this);
		var fd = new FormData(form[0]);
		var blnDoisXML = false;

		e.preventDefault();

		if ($("#linhaArquivoUpload2").is(":hidden")) {
			if ($("#c_arquivo").val() == "")
			{
				alert("Selecione um arquivo XML de NFe!!");
				return false;
			}
		}
		else {
			if (($("#c_arquivo").val() == "") || ($("#c_arquivo2").val() == ""))
			{
				alert("Selecione os arquivos XML primário e secundário de NFe!!");
				return false;
			}
			blnDoisXML = true;
		}

		$("#uploaded_file_guid").val("");
		$("#txtArquivoEnviado").val("");
		$("#c_status_upload_ok").val("");
		$("#c_status_get_nfe_ok").val("");
		$("#txtArquivoEnviado").hide();
		$("#imgUploadOk").hide();
		$("#imgUploadPending").show();

		if (blnDoisXML) {
			$("#uploaded_file_guid2").val("");
			$("#txtArquivoEnviado2").val("");
			$("#c_status_upload_ok2").val("");
			$("#c_status_get_nfe_ok2").val("");
			$("#txtArquivoEnviado2").hide();
			$("#imgUploadOk2").hide();
			$("#imgUploadPending2").show();
		}

		$("#divAjaxRunning").show();

		$.ajax({
			type: form.attr('method'),
			url: 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/UploadFile/PostFile',
			data: fd,
			enctype: 'multipart/form-data',
			processData: false,
			contentType: false,
			success: function (resp) {
				$("#divAjaxRunning").hide();
				jsonFileInfo = JSON.parse(resp);
				if (jsonFileInfo.Status == "OK")
				{
					$("#uploaded_file_guid").val(jsonFileInfo.files[0].stored_file_guid);
					$("#txtArquivoEnviado").val(jsonFileInfo.files[0].original_file_name);
					$("#c_status_upload_ok").val("1");
					$("#txtArquivoEnviado").show();
					$("#imgUploadOk").show();
					$("#imgUploadPending").hide();

					if (blnDoisXML) {
					$("#uploaded_file_guid2").val(jsonFileInfo.files[1].stored_file_guid);
					$("#txtArquivoEnviado2").val(jsonFileInfo.files[1].original_file_name);
					$("#c_status_upload_ok2").val("1");
					$("#txtArquivoEnviado2").show();
					$("#imgUploadOk2").show();
					$("#imgUploadPending2").hide();
					}
					
					// Obtém o objeto JSON com os dados da NFe
					$("#divAjaxRunning").show();
					$.ajax({
						url: 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/UploadedFile/ConvertXmlToJson',
						type: "GET",
						dataType: 'json',
						data: {
							id: jsonFileInfo.files[0].stored_file_guid
						}
					})
					.done(function (response) {
						$("#divAjaxRunning").hide();
						nfe = response;
						//arquivo_nfe = nfe;
					    //$("#arquivo_nfe").val(nfe);						
						preencheCamposNFe();
						$("#c_status_get_nfe_ok").val("1");
					})
					.fail(function (jqXHR, textStatus) {
						$("#divAjaxRunning").hide();
						var msgErro = "";
						if (textStatus.toString().length > 0) msgErro = "Mensagem de Status: " + textStatus.toString();
						try {
							if (jqXHR.status.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Status: " + jqXHR.status.toString(); }
						} catch (e) { }

						try {
							if (jqXHR.statusText.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Descrição do Status: " + jqXHR.statusText.toString(); }
						} catch (e) { }

						try {
							if (jqXHR.responseText.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Mensagem de Resposta: " + jqXHR.responseText.toString(); }
						} catch (e) { }

						alert("Falha ao tentar processar a requisição!!\nArquivo XML não pode ser importado devido a inconsistências internas!!\n\n" + msgErro);
					});
					
					if (blnDoisXML) {
		
					// Obtém o objeto JSON com os dados da NFe
					$("#divAjaxRunning").show();
					$.ajax({
						url: 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/UploadedFile/ConvertXmlToJson',
						type: "GET",
						dataType: 'json',
						data: {
							id: jsonFileInfo.files[1].stored_file_guid
						}
					})
					.done(function (response) {
						$("#divAjaxRunning").hide();
						nfe2 = response;
						arquivo_nfe2 = nfe2;
						preencheCamposNFe2();
						$("#c_status_get_nfe_ok2").val("1");
					})
					.fail(function (jqXHR, textStatus) {
						$("#divAjaxRunning").hide();
						var msgErro = "";
						if (textStatus.toString().length > 0) msgErro = "Mensagem de Status: " + textStatus.toString();
						try {
							if (jqXHR.status.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Status: " + jqXHR.status.toString(); }
						} catch (e) { }

						try {
							if (jqXHR.statusText.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Descrição do Status: " + jqXHR.statusText.toString(); }
						} catch (e) { }

						try {
							if (jqXHR.responseText.toString().length > 0) { if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Mensagem de Resposta: " + jqXHR.responseText.toString(); }
						} catch (e) { }

						alert("Falha ao tentar processar a requisição!!\nArquivo XML não pode ser importado devido a inconsistências internas!!\n\n" + msgErro);
					});

		
					}
					
				}
				else
				{
					alert("Falha no upload do arquivo!!\n\n" + jsonFileInfo.Message);
				}
			},
			error: function (resp, cod, msgErro) {
				$("#divAjaxRunning").hide();
				alert("Erro no upload do arquivo para o servidor!!\n\n" + msgErro);
			}
		});
		
	});
});


function trataOpcoesUpload(op) {
    if (op == "N") {
        $("#c_op_upload").val("N");
		$("#divArquivoXML2").hide();
		$("#linhaStatusUpload2").hide();
		$("#linhaArquivoUpload2").hide();
		$("#c_arquivo2").val("");
	}
	else if (op == "I") {
	    $("#c_op_upload").val("I");
	    $("#divArquivoXML2").show();
		$("#linhaStatusUpload2").show();
		$("#linhaArquivoUpload2").show();
	}
	else if (op == "L") {
	    $("#c_op_upload").val("L");
		$("#divArquivoXML2").show();
		$("#linhaStatusUpload2").show();
		$("#linhaArquivoUpload2").show();
	}
	else if (op == "M") {
	    $("#c_op_upload").val("M");
	    $("#divArquivoXML2").hide();
	    $("#linhaStatusUpload2").hide();
	    $("#linhaArquivoUpload2").hide();
	    $("#c_arquivo2").val("");
    }
	else {
	    $("#c_op_upload").val("");
	    $("#divArquivoXML2").hide();
	    $("#linhaStatusUpload2").hide();
	    $("#linhaArquivoUpload2").hide();
	    $("#c_arquivo2").val("");
	}
};

function preencheCamposNFe() {
    var vetnfe;

    //conforme detectado, há situações em que o XML não contém a tag 'nfeProc',
    //portanto, o teste abaixo verifica a existência da mesma
    if (nfe.hasOwnProperty('nfeProc')) {
        vetnfe = nfe.nfeProc;
    }
    else {
        vetnfe = nfe;
    }

    if (vetnfe.NFe.infNFe.hasOwnProperty('det')) {
        if (Array.isArray(vetnfe.NFe.infNFe.det)) {
            $("#c_nfe_qtde_itens").val(vetnfe.NFe.infNFe.det.length.toString());
        }
        else {
            $("#c_nfe_qtde_itens").val('1');
        }
    }
    else {
        $("#c_nfe_qtde_itens").val('1');
    }

    $("#c_nfe_numero_nf").val(vetnfe.NFe.infNFe.ide.nNF.toString());
    $("#c_nfe_emitente_cnpj").val(vetnfe.NFe.infNFe.emit.CNPJ.toString());
    $("#c_nfe_emitente_nome").val(vetnfe.NFe.infNFe.emit.xNome.toString());
    $("#c_nfe_emitente_nome_fantasia").val(vetnfe.NFe.infNFe.emit.xFant.toString());
    $("#c_nfe_destinatario_cnpj").val(vetnfe.NFe.infNFe.dest.CNPJ.toString());
    $("#c_nfe_dt_hr_emissao").val(vetnfe.NFe.infNFe.ide.dhEmi.toString());
}

function preencheCamposNFe2() {
    var vetnfe;

    //conforme detectado, há situações em que o XML não contém a tag 'nfeProc',
    //portanto, o teste abaixo verifica a existência da mesma
    if (nfe2.hasOwnProperty('nfeProc')) {
        vetnfe = nfe2.nfeProc;
    }
    else {
        vetnfe = nfe2;
    }


	//$("#c_nfe_qtde_itens2").val(nfe2.nfeProc.NFe.infNFe.det.length.toString());
	//havendo dois arquivos, Documento será no formato 'doc1_doc2'
	$("#c_nfe_numero_nf2").val(nfe2.nfeProc.NFe.infNFe.ide.nNF.toString());
    //$("#c_nfe_numero_nf").val(nfe.nfeProc.NFe.infNFe.ide.nNF.toString() + "_" + nfe2.nfeProc.NFe.infNFe.ide.nNF.toString());
    //$("#c_nfe_numero_nf").val($("#c_nfe_numero_nf").val() + "_" + vetnfe.NFe.infNFe.ide.nNF.toString());
    $("#c_nfe_emitente_cnpj2").val(vetnfe.NFe.infNFe.emit.CNPJ.toString());
    $("#c_nfe_emitente_nome2").val(vetnfe.NFe.infNFe.emit.xNome.toString());
    $("#c_nfe_emitente_nome_fantasia2").val(vetnfe.NFe.infNFe.emit.xFant.toString());
    $("#c_nfe_destinatario_cnpj2").val(vetnfe.NFe.infNFe.dest.CNPJ.toString());
    $("#c_nfe_dt_hr_emissao2").val(vetnfe.NFe.infNFe.ide.dhEmi.toString());
}

function fESTOQConfirma(f) {
	if (trim($("#uploaded_file_guid").val()) == "")
	{
		alert("Não há nenhum arquivo XML transferido e pronto para a operação de entrada no estoque!!");
		return;
	}
    
	if (trim($("#c_status_get_nfe_ok").val()) != "1") {
		alert("Os dados da NFe não foram recuperados corretamente!!");
		return;
	}

	fUPLOAD.c_FormFieldValues.value = formToStringAll($("#fUPLOAD"));
	fESTOQ.c_FormFieldValues.value = formToStringAll($("#fESTOQ"));

	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";

	f.action="EstoqueEntradaViaXMLDataEntry.asp";
	
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

<style type="text/css">
#divAjaxRunning
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	z-index:1001;
	background-color:grey;
	opacity: .6;
}
.AjaxImgLoader
{
	position: absolute;
	left: 50%;
	top: 50%;
	margin-left: -128px; /* -1 * image width / 2 */
	margin-top: -128px;  /* -1 * image height / 2 */
	display: block;
}
.TdTitStatus
{
	text-align:right;
	vertical-align:middle;
	width:100px;
}
</style>


<body onload="trataOpcoesUpload('<%=c_op_upload%>')">
<center>

<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>

<form id="fUPLOAD" name="fUPLOAD" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="upload_parameter__is_temp_file" value="1" />
<input type="hidden" name="upload_parameter__folder_name" value="ENTRADA_ESTOQUE_XML_NFE" />
<input type="hidden" name="upload_parameter__user_id" value="<%=usuario%>" />
<input type="hidden" name="upload_parameter__sessionToken" value="<%=s_sessionToken%>" />
<input type="hidden" name="upload_parameter__is_confirmation_required" value="1" />
<input type="hidden" name="upload_parameter__save_file_content_in_db_as_text" value="1" />
<input type="hidden" name="c_status_upload_ok" id="c_status_upload_ok" />
<input type="hidden" name="c_status_upload_ok2" id="c_status_upload_ok2" />
<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />
<!--<input type="hidden" name="arquivo_nfe" id="arquivo_nfe" value=""/>-->


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Entrada de Mercadorias no Estoque via XML</span>
	<br /><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br />

<!--  UPLOAD DO XML  -->
<table cellspacing="0" cellpadding="0" border="0">
<!--  OPÇÕES DE UPLOAD  -->
	<tr>
		<td nowrap>
			<input type="radio" id="rb_op_upload" name="rb_op_upload" class="CBOX" value="N" checked 
                onclick="trataOpcoesUpload('N');" >
			<span style="cursor:default">Normal</span>
			<br><br>
			<input type="radio" id="rb_op_upload" name="rb_op_upload" class="CBOX" value="I" 
                onclick="trataOpcoesUpload('I');" >
			<span style="cursor:default">Incentivo Fiscal</span>
			<br><br>
			<input type="radio" id="rb_op_upload" name="rb_op_upload" class="CBOX" value="L" 
                onclick="trataOpcoesUpload('L');" >
			<span style="cursor:default">Operador Logístico</span>
			<br><br>
			<input type="radio" id="rb_op_upload" name="rb_op_upload" class="CBOX" value="M" 
                onclick="trataOpcoesUpload('M');" >
			<span style="cursor:default">Manual</span>
			<br><br>
		</td>
	</tr>
<!--  ARQUIVOS XML  -->
	<tr>
		<td align="left" nowrap>
			<div id="divArquivoXML1" name="divArquivoXML1">
				<span class="PLTe">Arquivo XML da NFe</span>
				<br /><input type="file" name="c_arquivo" id="c_arquivo" accept=".xml" class="PLLe" style="width:700px;text-align:left;font-size:9pt;font-weight:normal;" />
			</div>
		</td>
	</tr>
    <tr><td></td></tr>
	<tr>
		<td align="left" nowrap>
			<div id="divArquivoXML2" name="divArquivoXML2">
				<span class="PLTe">Arquivo Secundário XML da NFe</span>
				<br /><input type="file" name="c_arquivo2" id="c_arquivo2" accept=".xml" class="PLLe" style="width:700px;text-align:left;font-size:9pt;font-weight:normal;" />
			</div>
		</td>
	</tr>
	<tr style="height:15px;">
		<td></td>
	</tr>
	<tr>
		<td align="right">
			<input type="submit" name="btn_upload" id="btn_upload" class="Button" style="font-size:12pt;color:green;" value="UPLOAD" />
		</td>
	</tr>
	<tr style="height:40px;">
		<td></td>
	</tr>
	<tr>
		<td align="left" nowrap>
			<table cellspacing="0" cellpadding="4" border="0">
				<tr>
					<td class="MC MB ME MD TdTitStatus" align="right" valign="middle">
						<span class="PLTd" style="magin-left:4px;margin-right:4px;">Status do upload</span>
					</td>
					<td class="MC MB MD" align="center" style="width:30px;">
						<img name="imgUploadPending" id="imgUploadPending" src="../IMAGEM/pending_peq.png" height="22" width="22" style="width:22px;height:22px;margin-left:4px;margin-right:4px;" />
						<img name="imgUploadOk" id="imgUploadOk" src="../IMAGEM/Ok_redondo_peq.jpg" height="22" width="22" style="width:22px;height:22px;margin-left:4px;margin-right:4px;display:none;" />
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr style="height:8px;">
		<td></td>
	</tr>
	<tr>
		<td align="left" nowrap>
			<table cellspacing="0" cellpadding="4" border="0">
				<tr>
					<td class="MC MB ME MD TdTitStatus" align="right" valign="middle">
						<span class="PLTd" style="magin-left:4px;margin-right:4px;">Arquivo enviado</span>
					</td>
					<td class="MC MB MD" align="left" style="width:540px;">
						<input type="text" name="txtArquivoEnviado" id="txtArquivoEnviado" readonly tabindex="-1" class="TA" style="width:520px;font-size:9pt;" />
					</td>
				</tr>
			</table>
		</td>
	</tr>
	

	<tr style="height:40px;">
		<td></td>
		</tr>
		<tr id="linhaStatusUpload2">
			<td align="left" nowrap>
				<table cellspacing="0" cellpadding="4" border="0">
					<tr>
						<td class="MC MB ME MD TdTitStatus" align="right" valign="middle">
							<span class="PLTd" style="magin-left:4px;margin-right:4px;">Status do upload Secundário</span>
						</td>
						<td class="MC MB MD" align="center" style="width:30px;">
							<img name="imgUploadPending2" id="imgUploadPending2" src="../IMAGEM/pending_peq.png" height="22" width="22" style="width:22px;height:22px;margin-left:4px;margin-right:4px;" />
							<img name="imgUploadOk2" id="imgUploadOk2" src="../IMAGEM/Ok_redondo_peq.jpg" height="22" width="22" style="width:22px;height:22px;margin-left:4px;margin-right:4px;display:none;" />
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr style="height:8px;">
			<td></td>
		</tr>
		<tr  id="linhaArquivoUpload2">
			<td align="left" nowrap>
				<table cellspacing="0" cellpadding="4" border="0">
					<tr>
						<td class="MC MB ME MD TdTitStatus" align="right" valign="middle">
							<span class="PLTd" style="magin-left:4px;margin-right:4px;">Arquivo secundário enviado</span>
						</td>
						<td class="MC MB MD" align="left" style="width:540px;">
							<input type="text" name="txtArquivoEnviado2" id="txtArquivoEnviado2" readonly tabindex="-1" class="TA" style="width:520px;font-size:9pt;" />
						</td>
					</tr>
				</table>
			</td>
		</tr>

		
	
</table>
</form>

<form id="fESTOQ" name="fESTOQ" method="post"">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="uploaded_file_guid" id="uploaded_file_guid" />
<input type="hidden" name="c_status_get_nfe_ok" id="c_status_get_nfe_ok" />
<input type="hidden" name="c_nfe_qtde_itens" id="c_nfe_qtde_itens" />
<input type="hidden" name="c_nfe_numero_nf" id="c_nfe_numero_nf" />
<input type="hidden" name="c_nfe_emitente_cnpj" id="c_nfe_emitente_cnpj" />
<input type="hidden" name="c_nfe_destinatario_cnpj" id="c_nfe_destinatario_cnpj" />
<input type="hidden" name="c_nfe_emitente_nome" id="c_nfe_emitente_nome" />
<input type="hidden" name="c_nfe_emitente_nome_fantasia" id="c_nfe_emitente_nome_fantasia" />
<input type="hidden" name="c_nfe_dt_hr_emissao" id="c_nfe_dt_hr_emissao" />
<input type="hidden" name="arquivo_nfe" id="arquivo_nfe" value=""/>

<input type="hidden" name="uploaded_file_guid2" id="uploaded_file_guid2" />
<input type="hidden" name="c_status_get_nfe_ok2" id="c_status_get_nfe_ok2" />
<input type="hidden" name="c_nfe_qtde_itens2" id="c_nfe_qtde_itens2" />
<input type="hidden" name="c_nfe_numero_nf2" id="c_nfe_numero_nf2" />
<input type="hidden" name="c_nfe_emitente_cnpj2" id="c_nfe_emitente_cnpj2" />
<input type="hidden" name="c_nfe_destinatario_cnpj2" id="c_nfe_destinatario_cnpj2" />
<input type="hidden" name="c_nfe_emitente_nome2" id="c_nfe_emitente_nome2" />
<input type="hidden" name="c_nfe_emitente_nome_fantasia2" id="c_nfe_emitente_nome_fantasia2" />
<input type="hidden" name="c_nfe_dt_hr_emissao2" id="c_nfe_dt_hr_emissao2" />
<input type="hidden" name="arquivo_nfe2" id="arquivo_nfe2" value=""/>

<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />

<input type="hidden" name="c_op_upload" id="c_op_upload" value="<%=c_op_upload%>" />

<input type="hidden" name="upload_parameter__is_temp_file" value="1" />
<input type="hidden" name="upload_parameter__folder_name" value="ENTRADA_ESTOQUE_XML_NFE" />
<input type="hidden" name="upload_parameter__user_id" value="<%=usuario%>" />
<input type="hidden" name="upload_parameter__sessionToken" value="<%=s_sessionToken%>" />
<input type="hidden" name="upload_parameter__is_confirmation_required" value="1" />
<input type="hidden" name="upload_parameter__save_file_content_in_db_as_text" value="1" />

</form>

<br />

<!-- ************   SEPARADOR   ************ -->
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br />

<table width="749" cellspacing="0">
<tr>
	<td align="left"><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fESTOQConfirma(fESTOQ)" title="vai para a página seguinte">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>

</center>
</body>
</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>