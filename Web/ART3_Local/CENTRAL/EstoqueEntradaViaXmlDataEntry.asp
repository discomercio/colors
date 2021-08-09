<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ====================================
'	  EstoqueEntradaViaXmlDataEntry.asp
'     ====================================
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

	dim uploaded_file_guid
    dim uploaded_file_guid2
	uploaded_file_guid = Trim(Request("uploaded_file_guid"))
    uploaded_file_guid2 = Trim(Request("uploaded_file_guid2"))
	if uploaded_file_guid = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Nenhum identificador de arquivo foi informado."
		end if

	dim s, i, iQtdeItens, iQtdeLinhas
	dim c_nfe_qtde_itens, c_nfe_numero_nf, c_nfe_numero_nf2, c_nfe_emitente_cnpj, c_nfe_destinatario_cnpj, c_nfe_emitente_nome, c_nfe_emitente_nome_fantasia, c_nfe_dt_hr_emissao, c_nfe_dt_hr_emissao2
    dim s_nfe_numero_nf
    dim rb_op_upload, c_op_upload
    dim arquivo_nfe, arquivo_nfe2
    dim s_valor_readonly
    dim s_classe_editavel
	c_nfe_qtde_itens = Trim(Request("c_nfe_qtde_itens"))
	c_nfe_numero_nf = Trim(Request("c_nfe_numero_nf"))
    c_nfe_numero_nf2 = Trim(Request("c_nfe_numero_nf2"))
	c_nfe_emitente_cnpj = retorna_so_digitos(Trim(Request("c_nfe_emitente_cnpj")))
	c_nfe_destinatario_cnpj = retorna_so_digitos(Trim(Request("c_nfe_destinatario_cnpj")))
	c_nfe_emitente_nome = Trim(Request("c_nfe_emitente_nome"))
	c_nfe_emitente_nome_fantasia = Trim(Request("c_nfe_emitente_nome_fantasia"))
	c_nfe_dt_hr_emissao = Trim(Request("c_nfe_dt_hr_emissao"))
	c_nfe_dt_hr_emissao2 = Trim(Request("c_nfe_dt_hr_emissao2"))
    arquivo_nfe = Trim(Request.Form("arquivo_nfe"))
    arquivo_nfe2 = Trim(Request.Form("arquivo_nfe2"))
    
    'rb_op_upload = Trim(Request.Form("rb_op_upload"))
    c_op_upload = Trim(Request("c_op_upload"))

	iQtdeItens = converte_numero(c_nfe_qtde_itens)

    if c_op_upload = "M" then
        iQtdeLinhas = MAX_PRODUTOS_ENTRADA_ESTOQUE
        s_valor_readonly = " "        
        s_classe_editavel = " TxtEditavel"
    else
        iQtdeLinhas = iQtdeItens
        s_valor_readonly = " readonly tabindex=-1"
        s_classe_editavel = " "
    end if

    s_nfe_numero_nf = c_nfe_numero_nf
    if c_nfe_numero_nf2 <> "" then s_nfe_numero_nf = s_nfe_numero_nf & "_" & c_nfe_numero_nf2

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

    dim s_sessionToken
	s_sessionToken = ""
	s = "SELECT Convert(varchar(36), SessionTokenModuloCentral) AS SessionTokenModuloCentral FROM t_USUARIO WHERE (usuario = '" & usuario & "')"
    if rs.State <> 0 then rs.Close
    rs.Open s,cn
	if Not rs.Eof then s_sessionToken = Trim("" & rs("SessionTokenModuloCentral"))


	dim alerta
	alerta = ""

	dim s_id_nfe_emitente
	s_id_nfe_emitente = ""

	if alerta = "" then
		if c_nfe_destinatario_cnpj <> "" then
			s = "SELECT id FROM t_NFe_EMITENTE WHERE (cnpj = '" & c_nfe_destinatario_cnpj & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Not rs.Eof then
				s_id_nfe_emitente = Trim("" & rs("id"))
				rs.MoveNext
				if Not rs.Eof then
				'	HÁ MAIS DE UM REGISTRO COM O MESMO CNPJ
					s_id_nfe_emitente = ""
					end if
				end if
			end if
		end if

	dim s_fabricante_codigo, s_fabricante_nome, s_nfe_dt_hr_emissao, s_perc_agio
	s_fabricante_codigo = ""
	s_fabricante_nome = ""
    s_perc_agio = "0,0000"

	if alerta = "" then
		if c_nfe_emitente_cnpj <> "" then
			s = "SELECT fabricante, nome, razao_social, perc_agio FROM t_FABRICANTE WHERE (cnpj = '" & c_nfe_emitente_cnpj & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Not rs.Eof then
				s_fabricante_codigo = Trim("" & rs("fabricante"))
				s_fabricante_nome = Trim("" & rs("nome"))
				if s_fabricante_nome = "" then s_fabricante_nome = Trim("" & rs("razao_social"))
                s_perc_agio = formata_numero(rs("perc_agio"), 4)
				rs.MoveNext
				if Not rs.Eof then
				'	HÁ MAIS DE UM REGISTRO COM O MESMO CNPJ
					s_fabricante_codigo = ""
					s_fabricante_nome = ""
                    s_perc_agio = "0,0000"
					end if
                if trim(s_perc_agio) = "" then s_perc_agio = "0,0000"
				end if
			end if
		
		if (s_fabricante_codigo = "") And (c_nfe_emitente_nome_fantasia <> "") then
			s = "SELECT fabricante, nome, razao_social FROM t_FABRICANTE WHERE (nome LIKE '" & BD_CURINGA_TODOS & c_nfe_emitente_nome_fantasia & BD_CURINGA_TODOS & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Not rs.Eof then
				s_fabricante_codigo = Trim("" & rs("fabricante"))
				s_fabricante_nome = Trim("" & rs("nome"))
				if s_fabricante_nome = "" then s_fabricante_nome = Trim("" & rs("razao_social"))
				rs.MoveNext
				if Not rs.Eof then
				'	HÁ MAIS DE UM REGISTRO NO RESULTADO
					s_fabricante_codigo = ""
					s_fabricante_nome = ""
					end if
				end if
			end if
		end if

	dim vDtHr, vDt, vHr
	if alerta = "" then
		s_nfe_dt_hr_emissao = c_nfe_dt_hr_emissao
		if c_nfe_dt_hr_emissao <> "" then
			vDtHr = Split(c_nfe_dt_hr_emissao, "T")
			vDt = Split(vDtHr(LBound(vDtHr)), "-")
			s_nfe_dt_hr_emissao = vDt(LBound(vDt)+2) & "/" & vDt(LBound(vDt)+1) & "/" & vDt(LBound(vDt))
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
<script src="../Global/jquery-my-janelaean.js?v=001" language="JavaScript" type="text/javascript"></script>  

<script type="text/javascript">
	var serverVariableUrl;
	var uploaded_file_guid;
	var uploaded_file_guid2;
	var nfe;
	var nfe2;
	var iQtdeItens = '<%=iQtdeItens%>';
	var fEANPopup;
	var blnDoisXML = false; 
	var jqxhr;

	serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
	serverVariableUrl = serverVariableUrl.toUpperCase();
	serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));
	uploaded_file_guid = '<%=uploaded_file_guid%>';
	uploaded_file_guid2 = '<%=uploaded_file_guid2%>';

//Dynamically assign height
function sizeDivAjaxRunning() {
	var newTop = $(window).scrollTop() + "px";
	$("#divAjaxRunning").css("top", newTop);
}

$(function () {
	$("#divAjaxRunning").hide();

	$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8

	$(document).ajaxStart(function () {
		$("#divAjaxRunning").show();
	})
	.ajaxStop(function () {
		$("#divAjaxRunning").hide();
	});

	//Every resize of window
	$(window).resize(function() {
		sizeDivAjaxRunning();
	});

	//Every scroll of window
	$(window).scroll(function() {
		sizeDivAjaxRunning();
	});

	// Trata o problema em que os campos do formulário são limpos após retornar à esta página c/ o history.back() pela 2ª vez quando ocorre erro de consistência
	if (trim(fESTOQ.c_FormFieldValues.value) != "")
	{
		stringToForm(fESTOQ.c_FormFieldValues.value, $('#fESTOQ'));
	}

	jqxhr = $.ajax({
		url: '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/UploadedFile/ConvertXmlToJson',
		type: "GET",
		dataType: 'json',
		data: {
			id: uploaded_file_guid
		}
	})
	.done(function (response) {
		$("#divAjaxRunning").hide();
		nfe = response;
		preencheForm();
	})
	.fail(function (jqXHR, textStatus) {
		$("#divAjaxRunning").hide();
		var msgErro = "";
		if (textStatus.toString().length > 0) msgErro = "Mensagem de Status: " + textStatus.toString();
		try {
			if (jqXHR.status.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Status: " + jqXHR.status.toString();}
		} catch (e) { }

		try {
			if (jqXHR.statusText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Descrição do Status: " + jqXHR.statusText.toString();}
		} catch (e) { }
		
		try {
			if (jqXHR.responseText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Mensagem de Resposta: " + jqXHR.responseText.toString();}
		} catch (e) { }
		
		alert("Falha ao tentar processar a requisição!!\n\n" + msgErro);
	});

	if (uploaded_file_guid2 != "") {

	    jqxhr = $.ajax({
			url: '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/UploadedFile/ConvertXmlToJson',
	        type: "GET",
	        dataType: 'json',
	        data: {
	            id: uploaded_file_guid2
	        }
	    })
        .done(function (response) {
            $("#divAjaxRunning").hide();
            nfe2 = response;
            complementaForm();
            //alert("pegou dois");
        })
        .fail(function (jqXHR, textStatus) {
            $("#divAjaxRunning").hide();
            var msgErro = "";
            if (textStatus.toString().length > 0) msgErro = "Mensagem de Status: " + textStatus.toString();
            try {
                if (jqXHR.status.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Status: " + jqXHR.status.toString();}
            } catch (e) { }

            try {
                if (jqXHR.statusText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Descrição do Status: " + jqXHR.statusText.toString();}
            } catch (e) { }
		
            try {
                if (jqXHR.responseText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Mensagem de Resposta: " + jqXHR.responseText.toString();}
            } catch (e) { }
		
            alert("Falha ao tentar processar a requisição do segundo XML!!\n\n" + msgErro);
        });

	}

});

//function retorna_descricao_produto(linha)
//{
//        var form = fESTOQ;

//        var produto = $("#c_erp_codigo_" + linha.toString()).val();

//        var jqxhr = $.ajax({
//            url: '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>/ART3/WebAPI/api/GetData/ProdutoBySku',
//            type: "GET",
//            dataType: 'json',
//            data: {
//                codProduto: produto,
//                usuario: form.usuarioid,
//                sessionToken: form.sessionToken
//            }
//        })
//		.done(function (response) {
//		    alert("faiô");
//		    $("#c_descricao_erp_" + linha.toString()).text(response.descricao);
//		})
//		.fail(function (jqXHR, textStatus) {
//		    alert("rolô");
//		    $("#c_descricao_erp_" + linha.toString()).text(jqXHR.responseText.toString());
//		});

//        alert("terminô");
//}

function preencheForm()
{
    var f, i, sIdx, childProp;
    var s;
    var iQtdeItens = '<%=iQtdeItens%>';
    var vetnfe;

	if (nfe == null) {
		alert("Os dados da NFe não foram recuperados corretamente (preenchimento)!!");
		return;
	}

	f = fESTOQ;	

    //conforme detectado, há situações em que o XML não contém a tag 'nfeProc',
    //portanto, o teste abaixo verifica a existência da mesma
	if (nfe.hasOwnProperty('nfeProc')) {
	    vetnfe = nfe.nfeProc;
	}
	else {
	    vetnfe = nfe;
	}

	$("#c_xml_ide__cNF_1").val(vetnfe.NFe.infNFe.ide.cNF);
	$("#c_xml_ide__serie_1").val(vetnfe.NFe.infNFe.ide.serie);
	$("#c_xml_ide__nNF_1").val(vetnfe.NFe.infNFe.ide.nNF);
	$("#c_xml_emit__CNPJ_1").val(vetnfe.NFe.infNFe.emit.CNPJ);
	$("#c_xml_emit__xNome_1").val(vetnfe.NFe.infNFe.emit.xNome);
	$("#c_xml_dest__CNPJ_1").val(vetnfe.NFe.infNFe.dest.CNPJ);
	$("#c_xml_dest__xNome_1").val(vetnfe.NFe.infNFe.dest.xNome);
	//$("#c_xml_transp__CNPJ_1").val(vetnfe.NFe.transp.transporta.CNPJ);
	//$("#c_xml_transp__xNome_1").val(vetnfe.NFe.transp.transporta.xNome);


//SE TIVER APENAS UM ELEMENTO
	if (!Array.isArray(vetnfe.NFe.infNFe.det)) {
	    i = 0;
	    sIdx = (i + 1).toString();
	    if (vetnfe.NFe.infNFe.hasOwnProperty('det')) {
	        $("#c_nfe_nItem_" + sIdx).val("1");
	        $("#c_ean_" + sIdx).val(retorna_so_digitos(vetnfe.NFe.infNFe.det.prod.cEAN));
	        $("#b_ean_exibe_" + sIdx).attr('title', retorna_so_digitos(vetnfe.NFe.infNFe.det.prod.cEAN));
	        $("#c_nfe_codigo_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.cProd);
	        //removendo descrição da nota, a pedido do Adailton
	        //$("#c_nfe_descricao_" + sIdx).text(nfe.nfeProc.NFe.infNFe.det.prod.xProd);
	        $("#c_nfe_ncm_sh_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.NCM);
	        $("#c_nfe_cfop_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.CFOP);
	        $("#c_nfe_qtde_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det.prod.qCom, 1));
            $("#c_nfe_vl_unitario_nota_" + sIdx).val(formata_moeda_xml(vetnfe.NFe.infNFe.det.prod.vUnCom));
            $("#c_nfe_vl_unitario_" + sIdx).val(formata_moeda_xml(vetnfe.NFe.infNFe.det.prod.vUnCom));
	        $("#c_nfe_vl_total_nota_" + sIdx).text(formata_numero(vetnfe.NFe.infNFe.det.prod.vProd, 2));
            $("#c_nfe_vl_total_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det.prod.vProd, 2));
            if (vetnfe.NFe.infNFe.det.prod.hasOwnProperty('vFrete')) {
                $("#c_nfe_vl_frete_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det.prod.vFrete, 2))
                $("#c_nfe_vl_frete_ori_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det.prod.vFrete, 2))
            }
            else {
                $("#c_nfe_vl_frete_" + sIdx).val("0,00");
                $("#c_nfe_vl_frete_ori_" + sIdx).val("0,00");
            }
	        for (var key in vetnfe.NFe.infNFe.det.imposto.ICMS) {
	            if (vetnfe.NFe.infNFe.det.imposto.ICMS.hasOwnProperty(key)) {
	                childProp = vetnfe.NFe.infNFe.det.imposto.ICMS[key];
	                $("#c_nfe_cst_" + sIdx).val(childProp.orig.toString() + childProp.CST.toString());
	                if ($("#c_op_upload").val()=="I") {
	                    $("#c_erp_cst_" + sIdx).val("200");
	                }
	                else {
	                    $("#c_erp_cst_" + sIdx).val(converte_cst_nfe_fabricante_para_entrada_estoque($("#c_nfe_cst_" + sIdx).val()));
	                }
	                $("#c_nfe_aliq_icms_" + sIdx).val(formata_numero(childProp.pICMS, 0));
	                break;
	            }
	        }
	        //incluído o teste para ver se a propriedade IPITrib existe, pois estava dando problema
	        if (vetnfe.NFe.infNFe.det.imposto.hasOwnProperty('IPI')) { 
	            if (vetnfe.NFe.infNFe.det.imposto.IPI.hasOwnProperty('IPITrib')) {
	                if (vetnfe.NFe.infNFe.det.imposto.IPI.IPITrib.hasOwnProperty('vIPI')) {
	                    $("#c_nfe_vl_ipi_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det.imposto.IPI.IPITrib.vIPI, 2));
	                } 
	                else {			        
	                    $("#c_nfe_vl_ipi_" + sIdx).val("0,00");
	                }
	                if (vetnfe.NFe.infNFe.det.imposto.IPI.IPITrib.hasOwnProperty('pIPI')) {
	                    s = vetnfe.NFe.infNFe.det.imposto.IPI.IPITrib.pIPI;
	                    if (s.indexOf(".") > 0) s = s.substring(0, s.indexOf(".") + 2);
	                    $("#c_nfe_aliq_ipi_" + sIdx).val(formata_numero(s, 0));
	                }
	                else {
	                    $("#c_nfe_aliq_ipi_" + sIdx).val("0");
	                }
	            }
	            else {
	                if (vetnfe.NFe.infNFe.det.imposto.IPI.hasOwnProperty('vIPI')) {			        
	                    $("#c_nfe_vl_ipi_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det.imposto.IPI.vIPI, 2));
	                } 
	                else {
	                    $("#c_nfe_vl_ipi_" + sIdx).val("0,00");
	                }
	                if (vetnfe.NFe.infNFe.det.imposto.IPI.hasOwnProperty('pIPI')) {
	                    s = vetnfe.NFe.infNFe.det.imposto.IPI.pIPI;
	                    if (s.indexOf(".") > 0) s = s.substring(0, s.indexOf(".") + 2);
	                    $("#c_nfe_aliq_ipi_" + sIdx).val(formata_numero(s, 0));
	                }
	                else {
	                    $("#c_nfe_aliq_ipi_" + sIdx).val("0");
	                }
	            }
	        }
	        else {
	            $("#c_nfe_vl_ipi_" + sIdx).val("0,00");
	            $("#c_nfe_aliq_ipi_" + sIdx).val("0");
	        }
	        $("#c_nfe_vl_ipi_ori_" + sIdx).val($("#c_nfe_vl_ipi_" + sIdx).val());
	        $("#c_nfe_aliq_ipi_ori_" + sIdx).val($("#c_nfe_aliq_ipi_" + sIdx).val());


	        $("#c1_xml_prod_cProd_" + sIdx).val($("#c_nfe_codigo_" + sIdx).val());
	        $("#c1_xml_prod_cEAN_" + sIdx).val($("#c_ean_" + sIdx).val());
	        $("#c1_xml_prod__NCM_" + sIdx).val($("#c_nfe_ncm_sh_" + sIdx).val());
	        $("#c1_xml_prod__CFOP_" + sIdx).val($("#c_nfe_cfop_" + sIdx).val());
	        $("#c1_xml_prod__qCom_" + sIdx).val($("#c_nfe_qtde_" + sIdx).val());
	        $("#c1_xml_prod__vUnCom_" + sIdx).val($("#c_nfe_vl_unitario_nota_" + sIdx).val());
            $("#c1_xml_prod__vProd_" + sIdx).val($("#c_nfe_vl_total_" + sIdx).val());
            $("#c1_xml_prod__vFrete_" + sIdx).val($("#c_nfe_vl_frete_" + sIdx).val());            
	        $("#c1_xml_imposto__pICMS_" + sIdx).val($("#c_nfe_aliq_icms_" + sIdx).val());
	        $("#c1_xml_imposto__vIPI_" + sIdx).val($("#c_nfe_vl_ipi_" + sIdx).val());
	        $("#c1_xml_imposto__pIPI_" + sIdx).val($("#c_nfe_aliq_ipi_" + sIdx).val());
	    }

	}
//SE TIVER MAIS DE UM ELEMENTO
	else {
	    //for (i = 0; i < f.c_erp_codigo.length; i++) {
	    for (i = 0; i < iQtdeItens; i++) {
	        sIdx = (i + 1).toString();
	        if (i < vetnfe.NFe.infNFe.det.length) {
	            $("#c_nfe_nItem_" + sIdx).val(vetnfe.NFe.infNFe.det[i]['@nItem']);
	            $("#c_ean_" + sIdx).val(retorna_so_digitos(vetnfe.NFe.infNFe.det[i].prod.cEAN));
	            $("#b_ean_exibe_" + sIdx).attr('title', retorna_so_digitos(vetnfe.NFe.infNFe.det[i].prod.cEAN));
	            $("#c_nfe_codigo_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.cProd);
	            //removendo descrição da nota, a pedido do Adailton
	            //$("#c_nfe_descricao_" + sIdx).text(nfe.nfeProc.NFe.infNFe.det[i].prod.xProd);
	            $("#c_nfe_ncm_sh_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.NCM);
	            $("#c_nfe_cfop_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.CFOP);
	            //$("#c_nfe_unid_" + sIdx).text(nfe.nfeProc.NFe.infNFe.det[i].prod.uCom);
	            //$("#c_nfe_unid_" + sIdx).val(nfe.nfeProc.NFe.infNFe.det[i].prod.uCom);
	            //$("#c_nfe_qtde_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].prod.qCom, 1));
	            $("#c_nfe_qtde_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det[i].prod.qCom, 1));
	            //$("#c_nfe_vl_unitario_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].prod.vUnCom, 4));
	            //$("#c_nfe_vl_unitario_nota_" + sIdx).val(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].prod.vUnCom, 4));
                $("#c_nfe_vl_unitario_nota_" + sIdx).val(formata_moeda_xml(vetnfe.NFe.infNFe.det[i].prod.vUnCom));
                $("#c_nfe_vl_unitario_" + sIdx).val(formata_moeda_xml(vetnfe.NFe.infNFe.det[i].prod.vUnCom));
	            $("#c_nfe_vl_total_nota_" + sIdx).text(formata_numero(vetnfe.NFe.infNFe.det[i].prod.vProd, 2));
                $("#c_nfe_vl_total_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det[i].prod.vProd, 2));
                if (vetnfe.NFe.infNFe.det[i].prod.hasOwnProperty('vFrete')) {
                    $("#c_nfe_vl_frete_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det[i].prod.vFrete, 2))
                    $("#c_nfe_vl_frete_ori_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det[i].prod.vFrete, 2))
                }
                else {
                    $("#c_nfe_vl_frete_" + sIdx).val("0,00");
                    $("#c_nfe_vl_frete_ori_" + sIdx).val("0,00");
                }
	            for (var key in vetnfe.NFe.infNFe.det[i].imposto.ICMS) {
	                if (vetnfe.NFe.infNFe.det[i].imposto.ICMS.hasOwnProperty(key)) {
	                    childProp = vetnfe.NFe.infNFe.det[i].imposto.ICMS[key];
	                    //$("#c_nfe_cst_" + sIdx).text(childProp.orig.toString() + childProp.CST.toString());
	                    $("#c_nfe_cst_" + sIdx).val(childProp.orig.toString() + childProp.CST.toString());
	                    //$("#c_erp_cst_" + sIdx).val(converte_cst_nfe_fabricante_para_entrada_estoque($("#c_nfe_cst_" + sIdx).text()));
	                    if ($("#c_op_upload").val()=="I") {
	                        $("#c_erp_cst_" + sIdx).val("200");
	                    }
	                    else {
	                        $("#c_erp_cst_" + sIdx).val(converte_cst_nfe_fabricante_para_entrada_estoque($("#c_nfe_cst_" + sIdx).val()));
	                    }
	                    //$("#c_nfe_vl_base_icms_" + sIdx).text(formata_numero(childProp.vBC, 2));
	                    //$("#c_nfe_vl_base_icms_nota_" + sIdx).text(formata_numero(childProp.vBC, 2));
	                    //$("#c_nfe_vl_base_icms_" + sIdx).val(formata_numero(childProp.vBC, 2));
	                    //$("#c_nfe_vl_icms_" + sIdx).text(formata_numero(childProp.vICMS, 2));
	                    //$("#c_nfe_vl_icms_nota_" + sIdx).text(formata_numero(childProp.vICMS, 2));
	                    //$("#c_nfe_vl_icms_" + sIdx).val(formata_numero(childProp.vICMS, 2));
	                    //$("#c_nfe_aliq_icms_" + sIdx).text(formata_numero(childProp.pICMS, 2));
	                    //$("#c_nfe_aliq_icms_nota_" + sIdx).text(formata_numero(childProp.pICMS, 2));
	                    $("#c_nfe_aliq_icms_" + sIdx).val(formata_numero(childProp.pICMS, 0));
	                    break;
	                }
	            }
	            //incluído o teste para ver se a propriedade IPITrib existe, pois estava dando problema
	            //rec ini
	            if (vetnfe.NFe.infNFe.det[i].imposto.hasOwnProperty('IPI')) { 
	                if (vetnfe.NFe.infNFe.det[i].imposto.IPI.hasOwnProperty('IPITrib')) {
	                    if (vetnfe.NFe.infNFe.det[i].imposto.IPI.IPITrib.hasOwnProperty('vIPI')) {
	                        //$("#c_nfe_vl_ipi_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.IPITrib.vIPI, 2));
	                        //$("#c_nfe_vl_ipi_nota_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.IPITrib.vIPI, 2));
	                        $("#c_nfe_vl_ipi_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det[i].imposto.IPI.IPITrib.vIPI, 2));
	                    } 
	                    else {			        
	                        //$("#c_nfe_vl_ipi_nota_" + sIdx).text("0,00");
	                        $("#c_nfe_vl_ipi_" + sIdx).val("0,00");
	                    }
	                    if (vetnfe.NFe.infNFe.det[i].imposto.IPI.IPITrib.hasOwnProperty('pIPI')) {
	                        //$("#c_nfe_aliq_ipi_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.IPITrib.pIPI, 2));
	                        //$("#c_nfe_aliq_ipi_nota_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.IPITrib.pIPI, 2));
	                        s = vetnfe.NFe.infNFe.det[i].imposto.IPI.IPITrib.pIPI;
	                        if (s.indexOf(".") > 0) s = s.substring(0, s.indexOf(".") + 2);
	                        $("#c_nfe_aliq_ipi_" + sIdx).val(formata_numero(s, 0));
	                    }
	                    else {
	                        //$("#c_nfe_aliq_ipi_nota_" + sIdx).text("0,00");
	                        $("#c_nfe_aliq_ipi_" + sIdx).val("0");
	                    }
	                }
	                else {
	                    if (vetnfe.NFe.infNFe.det[i].imposto.IPI.hasOwnProperty('vIPI')) {			        
	                        //$("#c_nfe_vl_ipi_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.vIPI, 2));
	                        //$("#c_nfe_vl_ipi_nota_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.vIPI, 2));
	                        $("#c_nfe_vl_ipi_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det[i].imposto.IPI.vIPI, 2));
	                    } 
	                    else {
	                        //$("#c_nfe_vl_ipi_nota_" + sIdx).text("0,00");
	                        $("#c_nfe_vl_ipi_" + sIdx).val("0,00");
	                    }
	                    if (vetnfe.NFe.infNFe.det[i].imposto.IPI.hasOwnProperty('pIPI')) {
	                        //$("#c_nfe_aliq_ipi_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.pIPI, 2));
	                        //$("#c_nfe_aliq_ipi_nota_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.pIPI, 2));
	                        s = vetnfe.NFe.infNFe.det[i].imposto.IPI.pIPI;
	                        if (s.indexOf(".") > 0) s = s.substring(0, s.indexOf(".") + 2);
	                        $("#c_nfe_aliq_ipi_" + sIdx).val(formata_numero(s, 0));
	                    }
	                    else {
	                        //$("#c_nfe_aliq_ipi_nota_" + sIdx).val("0,00");
	                        $("#c_nfe_aliq_ipi_" + sIdx).val("0");
	                    }
	                }
	            }
	            else {
	                $("#c_nfe_vl_ipi_" + sIdx).val("0,00");
	                $("#c_nfe_aliq_ipi_" + sIdx).val("0");
	            }
	            $("#c_nfe_vl_ipi_ori_" + sIdx).val($("#c_nfe_vl_ipi_" + sIdx).val());
	            $("#c_nfe_aliq_ipi_ori_" + sIdx).val($("#c_nfe_aliq_ipi_" + sIdx).val());

	            //rec fim

	        }
	        else {
	            $("#c_erp_codigo_" + sIdx).prop("readonly", true);
	            $("#c_erp_cst_" + sIdx).prop("readonly", true);
	            $("#c_erp_codigo_" + sIdx).attr("tabindex", -1);
	            $("#c_erp_cst_" + sIdx).attr("tabindex", -1);
	        }

	        //recalcula_itens();
	        //recalcula_total_nf();
	        $("#c1_xml_prod_cProd_" + sIdx).val($("#c_nfe_codigo_" + sIdx).val());
	        $("#c1_xml_prod_cEAN_" + sIdx).val($("#c_ean_" + sIdx).val());
	        $("#c1_xml_prod__NCM_" + sIdx).val($("#c_nfe_ncm_sh_" + sIdx).val());
	        $("#c1_xml_prod__CFOP_" + sIdx).val($("#c_nfe_cfop_" + sIdx).val());
	        $("#c1_xml_prod__qCom_" + sIdx).val($("#c_nfe_qtde_" + sIdx).val());
	        $("#c1_xml_prod__vUnCom_" + sIdx).val($("#c_nfe_vl_unitario_nota_" + sIdx).val());
            $("#c1_xml_prod__vProd_" + sIdx).val($("#c_nfe_vl_total_" + sIdx).val()); -
            $("#c1_xml_prod__vFrete_" + sIdx).val($("#c_nfe_vl_frete_" + sIdx).val());            
	        $("#c1_xml_imposto__pICMS_" + sIdx).val($("#c_nfe_aliq_icms_" + sIdx).val());
	        $("#c1_xml_imposto__vIPI_" + sIdx).val($("#c_nfe_vl_ipi_" + sIdx).val());
	        $("#c1_xml_imposto__pIPI_" + sIdx).val($("#c_nfe_aliq_ipi_" + sIdx).val());

	    }

	}

	
	recalcula_itens();
    //recalcula_total_nf();
	$("#c_total_nf").val(formata_moeda(vetnfe.NFe.infNFe.total.ICMSTot.vNF));
    
    // Ajusta a altura dos campos input para ficar na mesma altura da linha da tabela
	$(".TxtErpCodigo").each(function () {
		$(this).height($(this).parent().height());
	});

	$(".TxtErpCst").each(function () {
		$(this).height($(this).parent().height());
	});

	// Aceita somente dígitos
	$(".TxtErpFabr, .TxtErpCodigo, .TxtErpCst").keydown(function (e) {
		// Allow: delete, backspace, tab, escape, enter
		if ($.inArray(e.keyCode, [46, 8, 9, 27, 13]) !== -1 ||
			// Allow: Ctrl+A, Command+A, Ctrl+C, Ctrl+V, Ctrl+X
			(((e.keyCode === 65) || (e.keyCode === 67) || (e.keyCode === 86) || (e.keyCode === 88)) && (e.ctrlKey === true || e.metaKey === true)) ||
			// Allow: home, end, left, right, down, up
			(e.keyCode >= 35 && e.keyCode <= 40)) {
			// let it happen, don't do anything
			return;
		}
		// Ensure that it is a number and stop the keypress
		if ((e.shiftKey || (e.keyCode < 48 || e.keyCode > 57)) && (e.keyCode < 96 || e.keyCode > 105)) {
			e.preventDefault();
		}
	});

	$('.TxtErpCodigo').blur(function (e) {
	    var linha = retorna_so_digitos($(this).attr("id")).toString();
	    var usuid = '<%=usuario%>';
	    var sessionToken = '<%=s_sessionToken%>';
	    var produto = $(this).val();
	    
	    e.preventDefault();

	    if (produto =="") {
	        $("#c_descricao_erp_" + linha.toString()).css("color","black");
	        $("#c_descricao_erp_" + linha.toString()).val("");
	    }
	    else
	    {
	        var jqxhr = $.ajax({
				url: '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/GetData/ProdutoBySku',
	            type: 'GET',
	            dataType: 'json',
	            data: {
	                codProduto: produto,
	                usuario: usuid,
	                sessionToken: sessionToken
	            }
	        })
		    .done(function (response) {
		        $("#c_descricao_erp_" + linha.toString()).css("color","black");
		        $("#c_descricao_erp_" + linha.toString()).val(response.descricao);
		    })
		    .fail(function (jqXHR, textStatus) {
		        //$("#c_descricao_erp_" + linha.toString()).text(jqXHR.responseText.toString());
		        //$("#c_descricao_erp_" + linha.toString()).text(response.descricao);
		        $("#c_descricao_erp_" + linha.toString()).css("color","red");
		        $("#c_descricao_erp_" + linha.toString()).val("Problemas para localizar o produto " + produto);
		    });
	    }

	});

}

function complementaForm()
{
    var f, i, sIdx, childProp;
    var iQtdeItens = '<%=iQtdeItens%>';
    var c_op_upload = '<%=c_op_upload%>';
    var s;
    var vetnfe;

    if (nfe2 == null) {
        alert("Os dados do segundo XML não foram recuperados corretamente!!");
        return;
    }

    f = fESTOQ;
    
    //conforme detectado, há situações em que o XML não contém a tag 'nfeProc',
    //portanto, o teste abaixo verifica a existência da mesma
    if (nfe2.hasOwnProperty('nfeProc')) {
        vetnfe = nfe2.nfeProc;
    }
    else {
        vetnfe = nfe2;
    }

    $("#c_xml_ide__cNF_2").val(vetnfe.NFe.infNFe.ide.cNF);
    $("#c_xml_ide__serie_2").val(vetnfe.NFe.infNFe.ide.serie);
    $("#c_xml_ide__nNF_2").val(vetnfe.NFe.infNFe.ide.nNF);
    $("#c_xml_emit__CNPJ_2").val(vetnfe.NFe.infNFe.emit.CNPJ);
    $("#c_xml_emit__xNome_2").val(vetnfe.NFe.infNFe.emit.xNome);
    $("#c_xml_dest__CNPJ_2").val(vetnfe.NFe.infNFe.dest.CNPJ);
    $("#c_xml_dest__xNome_2").val(vetnfe.NFe.infNFe.dest.xNome);
    //$("#c_xml_transp__CNPJ_2").val(vetnfe.NFe.transp.transporta.CNPJ);
    //$("#c_xml_transp__xNome_2").val(vetnfe.NFe.transp.transporta.xNome);

    //SE TIVER APENAS UM ELEMENTO
    if (!Array.isArray(vetnfe.NFe.infNFe.det)) {
        i = 0;
        sIdx = (i + 1).toString();
        if (vetnfe.NFe.infNFe.hasOwnProperty('det')) {
            //$("#c_nfe_nItem_" + sIdx).val("1");
            //primeiro caso: tratamento de arquivos de Incentivo Fiscal
            if (c_op_upload == "I") {
                $("#c_nfe_cfop_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.CFOP);
                for (var key in vetnfe.NFe.infNFe.det.imposto.ICMS) {
                    if (vetnfe.NFe.infNFe.det.imposto.ICMS.hasOwnProperty(key)) {
                        childProp = vetnfe.NFe.infNFe.det.imposto.ICMS[key];
                        $("#c_nfe_cst_" + sIdx).val(childProp.orig.toString() + childProp.CST.toString());
                        //conforme solicitação do Adailton, sempre altera o CST de entrada para 200
                        $("#c_erp_cst_" + sIdx).val("200");
                        $("#c_nfe_aliq_icms_" + sIdx).val(formata_numero(childProp.pICMS, 0));
                        break;
                    }
                }
            }
                //segundo caso: tratamento de arquivos de Operador Logístico
            else {
                $("#c_nfe_cfop_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.CFOP);
                for (var key in vetnfe.NFe.infNFe.det.imposto.ICMS) {
                    if (vetnfe.NFe.infNFe.det.imposto.ICMS.hasOwnProperty(key)) {
                        childProp = vetnfe.NFe.infNFe.det.imposto.ICMS[key];
                        $("#c_nfe_cst_" + sIdx).val(childProp.orig.toString() + childProp.CST.toString());
                        $("#c_erp_cst_" + sIdx).val(converte_cst_nfe_fabricante_para_entrada_estoque($("#c_nfe_cst_" + sIdx).val()));
                        $("#c_nfe_aliq_icms_" + sIdx).val(formata_numero(childProp.pICMS, 0));
                        break;
                    }
                }
                //incluído o teste para ver se a propriedade IPITrib existe, pois estava dando problema
                if (vetnfe.NFe.infNFe.det.imposto.hasOwnProperty('IPI')) { 
                    if (vetnfe.NFe.infNFe.det.imposto.IPI.hasOwnProperty('IPITrib')) {
                        if (vetnfe.NFe.infNFe.det.imposto.IPI.IPITrib.hasOwnProperty('vIPI')) {
                            $("#c_nfe_vl_ipi_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det.imposto.IPI.IPITrib.vIPI, 2));
                        } 
                        if (vetnfe.NFe.infNFe.det.imposto.IPI.IPITrib.hasOwnProperty('pIPI')) {
                            s = vetnfe.NFe.infNFe.det.imposto.IPI.IPITrib.pIPI;
                            if (s.indexOf(".") > 0) s = s.substring(0, s.indexOf(".") + 2);
                            $("#c_nfe_aliq_ipi_" + sIdx).val(formata_numero(s, 0));
                        }
                    }
                    else {
                        if (vetnfe.NFe.infNFe.det.imposto.IPI.hasOwnProperty('vIPI')) {			        
                            $("#c_nfe_vl_ipi_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det.imposto.IPI.vIPI, 2));
                        } 
                        if (vetnfe.NFe.infNFe.det.imposto.IPI.hasOwnProperty('pIPI')) {
                            s = vetnfe.NFe.infNFe.det.imposto.IPI.pIPI;
                            if (s.indexOf(".") > 0) s = s.substring(0, s.indexOf(".") + 2);
                            $("#c_nfe_aliq_ipi_" + sIdx).val(formata_numero(s, 0));
                        }
                    }
                }
                $("#c_nfe_vl_ipi_ori_" + sIdx).val($("#c_nfe_vl_ipi_" + sIdx).val());
                $("#c_nfe_aliq_ipi_ori_" + sIdx).val($("#c_nfe_aliq_ipi_" + sIdx).val());
            }        
            $("#c2_xml_prod_cProd_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.cProd);
            $("#c2_xml_prod_cEAN_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.cEAN);
            $("#c2_xml_prod__NCM_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.NCM);
            $("#c2_xml_prod__CFOP_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.CFOP);
            $("#c2_xml_prod__qCom_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.qCom);
            $("#c2_xml_prod__vUnCom_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.vUnCom);
            $("#c2_xml_prod__vProd_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.vProd);
            $("#c2_xml_prod__vFrete_" + sIdx).val(vetnfe.NFe.infNFe.det.prod.vFrete);            
            $("#c2_xml_imposto__pICMS_" + sIdx).val($("#c_nfe_aliq_icms_" + sIdx).val());
            $("#c2_xml_imposto__vIPI_" + sIdx).val($("#c_nfe_vl_ipi_" + sIdx).val());
            $("#c2_xml_imposto__pIPI_" + sIdx).val($("#c_nfe_aliq_ipi_" + sIdx).val());
        }

    }
        //SE TIVER MAIS DE UM ELEMENTO
    else {
        
        for (i = 0; i < iQtdeItens; i++) {
            sIdx = (i + 1).toString();
            if (i < vetnfe.NFe.infNFe.det.length) {
                //primeiro caso: tratamento de arquivos de Incentivo Fiscal
                if (c_op_upload == "I") {
                    $("#c_nfe_cfop_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.CFOP);
                    for (var key in vetnfe.NFe.infNFe.det[i].imposto.ICMS) {
                        if (vetnfe.NFe.infNFe.det[i].imposto.ICMS.hasOwnProperty(key)) {
                            childProp = vetnfe.NFe.infNFe.det[i].imposto.ICMS[key];
                            $("#c_nfe_cst_" + sIdx).val(childProp.orig.toString() + childProp.CST.toString());
                            //conforme solicitação do Adailton, sempre altera o CST de entrada para 200
                            $("#c_erp_cst_" + sIdx).val("200");
                            $("#c_nfe_aliq_icms_" + sIdx).val(formata_numero(childProp.pICMS, 0));
                            break;
                        }
                    }
                }
                    //segundo caso: tratamento de arquivos de Operador Logístico
                else {
                    $("#c_nfe_cfop_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.CFOP);
                    for (var key in vetnfe.NFe.infNFe.det[i].imposto.ICMS) {
                        if (vetnfe.NFe.infNFe.det[i].imposto.ICMS.hasOwnProperty(key)) {
                            childProp = vetnfe.NFe.infNFe.det[i].imposto.ICMS[key];
                            $("#c_nfe_cst_" + sIdx).val(childProp.orig.toString() + childProp.CST.toString());
                            $("#c_erp_cst_" + sIdx).val(converte_cst_nfe_fabricante_para_entrada_estoque($("#c_nfe_cst_" + sIdx).val()));
                            $("#c_nfe_aliq_icms_" + sIdx).val(formata_numero(childProp.pICMS, 0));
                            break;
                        }
                    }
                    //incluído o teste para ver se a propriedade IPITrib existe, pois estava dando problema
                    if (vetnfe.NFe.infNFe.det[i].imposto.hasOwnProperty('IPI')) { 
                        if (vetnfe.NFe.infNFe.det[i].imposto.IPI.hasOwnProperty('IPITrib')) {
                            if (vetnfe.NFe.infNFe.det[i].imposto.IPI.IPITrib.hasOwnProperty('vIPI')) {
                                $("#c_nfe_vl_ipi_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det[i].imposto.IPI.IPITrib.vIPI, 2));
                            } 
                            if (vetnfe.NFe.infNFe.det[i].imposto.IPI.IPITrib.hasOwnProperty('pIPI')) {
                                s = vetnfe.NFe.infNFe.det[i].imposto.IPI.IPITrib.pIPI;
                                if (s.indexOf(".") > 0) s = s.substring(0, s.indexOf(".") + 2);
                                $("#c_nfe_aliq_ipi_" + sIdx).val(formata_numero(s, 0));
                            }
                        }
                        else {
                            if (vetnfe.NFe.infNFe.det[i].imposto.IPI.hasOwnProperty('vIPI')) {			        
                                $("#c_nfe_vl_ipi_" + sIdx).val(formata_numero(vetnfe.NFe.infNFe.det[i].imposto.IPI.vIPI, 2));
                            } 
                            if (vetnfe.NFe.infNFe.det[i].imposto.IPI.hasOwnProperty('pIPI')) {
                                s = vetnfe.NFe.infNFe.det[i].imposto.IPI.pIPI;
                                if (s.indexOf(".") > 0) s = s.substring(0, s.indexOf(".") + 2);
                                $("#c_nfe_aliq_ipi_" + sIdx).val(formata_numero(s, 0));
                            }
                        }
                    }
                    $("#c_nfe_vl_ipi_ori_" + sIdx).val($("#c_nfe_vl_ipi_" + sIdx).val());
                    $("#c_nfe_aliq_ipi_ori_" + sIdx).val($("#c_nfe_aliq_ipi_" + sIdx).val());
                }        
                $("#c2_xml_prod_cProd_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.cProd);
                $("#c2_xml_prod_cEAN_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.cEAN);
                $("#c2_xml_prod__NCM_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.NCM);
                $("#c2_xml_prod__CFOP_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.CFOP);
                $("#c2_xml_prod__qCom_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.qCom);
                $("#c2_xml_prod__vUnCom_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.vUnCom);
                $("#c2_xml_prod__vProd_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.vProd);
                $("#c2_xml_prod__vFrete_" + sIdx).val(vetnfe.NFe.infNFe.det[i].prod.vFrete);                
                $("#c2_xml_imposto__pICMS_" + sIdx).val($("#c_nfe_aliq_icms_" + sIdx).val());
                $("#c2_xml_imposto__vIPI_" + sIdx).val($("#c_nfe_vl_ipi_" + sIdx).val());
                $("#c2_xml_imposto__pIPI_" + sIdx).val($("#c_nfe_aliq_ipi_" + sIdx).val());
            }
        }
    }

    recalcula_itens();
}

function recalcula_itens() {
    var v_agio;
    var v_perc_agio;
	var v_calculo;
	var v_ipi;
    var v_aliq_ipi;
    var v_frete;
	var iQtdeItens = '<%=iQtdeLinhas%>';
	var s = "";

	//if (nfe == null) {
	//	alert("Os dados da NFe não foram recuperados corretamente (recálculo)!!");
	//	return;
	//}
	
	s = $("#c_perc_agio").val();
	if ((s == "") || (s == "0,0000")) {
	    s = "0";
	};
	v_perc_agio = converte_numero(s) / 100;
	if (v_perc_agio > 1) {
	    alert('Percentual de ágio maior que 100%!'); 
	    $("#c_perc_agio").focus();
	    //return;
	} 

	for (var i = 1; i <= iQtdeItens; i++) {
		if ($("#ckb_importa_" + trim(i.toString())).is(":checked")) {
			//calculo do valor do produto com IPI, frete e ágio
            v_calculo = converte_numero(formata_moeda_xml($("#c_nfe_vl_unitario_nota_" + trim(i.toString())).val()));
            v_frete = converte_numero($("#c_nfe_vl_frete_" + trim(i.toString())).val());
            v_frete = v_frete / converte_numero($("#c_nfe_qtde_" + trim(i.toString())).val());
            v_calculo = v_calculo + v_frete;
		    s = $("#c_nfe_aliq_ipi_" + trim(i.toString())).val();
		    v_aliq_ipi = converte_numero(s) / 100;
		    if (v_aliq_ipi > 0) {
		        v_ipi = converte_numero(formata_moeda(v_calculo * v_aliq_ipi));
		        $("#c_nfe_vl_ipi_" + trim(i.toString())).val(formata_moeda(v_ipi.toString()));
		    }
		    else {
		        v_ipi = converte_numero($("#c_nfe_vl_ipi_" + trim(i.toString())).val());
		    }
		    $("#c_nfe_vl_ipi_" + trim(i.toString())).val(formata_moeda(v_ipi.toString()));
            v_calculo = v_calculo + v_ipi;
		    v_agio = converte_numero(formata_moeda(v_calculo  * v_perc_agio));
		    v_calculo = converte_numero(formata_moeda(v_calculo + v_agio));
		    $("#c_nfe_vl_unitario_" + trim(i.toString())).val(formata_moeda(v_calculo.toString()));
		    //v_calculo = converte_numero(nfe.nfeProc.NFe.infNFe.det[i-1].prod.qCom) * converte_numero(formata_moeda($("#c_nfe_vl_unitario_" + trim(i.toString())).val()));
			//$("#c_nfe_vl_total_" + trim(i.toString())).val(formata_moeda(v_calculo.toString()));
			//v_calculo = converte_numero($("#c_nfe_vl_base_icms_" + trim(i.toString())).val()) * (1 + v_agio);
			//$("#c_nfe_vl_base_icms_" + trim(i.toString())).val(formata_numero(v_calculo.toString(), 2));
			//$("#c_nfe_vl_icms_" + trim(i.toString())).val(formata_numero(v_calculo.toString(), 2));
			recalcula_linha(i);
		}
		else {
		}
	}
	recalcula_total();
}

function recalcula_linha(i) {
	var v_calculo;

	v_calculo = converte_numero($("#c_nfe_qtde_" + trim(i.toString())).val()) * 
								converte_numero(formata_moeda($("#c_nfe_vl_unitario_" + trim(i.toString())).val()));
	$("#c_nfe_vl_total_" + trim(i.toString())).val(formata_moeda(v_calculo.toString()));
	recalcula_total();
}

function recalcula_total_nf() {
    var v_calculo;
    var v_total;
    var iQtdeItens = '<%=iQtdeLinhas%>';
    var f;
    var i;

    //não recalcular, ao invés disso, pegar o valor direto do XML na rotina preencheForm
    return;

    v_calculo = 0;
    v_total = 0;
    for (i = 1; i <= iQtdeItens; i++)
    {
        v_calculo = converte_numero($("#c_nfe_qtde_" + trim(i.toString())).val()) * 
                    (converte_numero(formata_numero($("#c_nfe_vl_unitario_nota_" + trim(i.toString())).val(), 2)) +
            converte_numero(formata_numero($("#c_nfe_vl_ipi_" + trim(i.toString())).val(), 2)) +
            converte_numero(formata_numero($("#c_nfe_vl_frete_" + trim(i.toString())).val(), 2)));
		v_total = v_total + v_calculo;
    }
	$("#c_total_nf").val(formata_moeda(v_total));
		

}

function recalcula_total() {
	var m, i;
	var s;
	var v_perc_agio;
	var iQtdeItens = '<%=iQtdeLinhas%>';

	s = $("#c_perc_agio").val();
	v_perc_agio = converte_numero(s) / 100;
	//if (v_perc_agio == 0) {
	//    $("#c_nfe_vl_total_geral").val(formata_moeda(nfe.nfeProc.NFe.infNFe.total.ICMSTot.vNF));
	//    return;
	//}

	m=0;
	for (i=1; i<=iQtdeItens; i++) 
	{
	    if ($("#ckb_importa_" + trim(i.toString())).is(":checked")) {
			m=m+converte_numero(formata_moeda($("#c_nfe_vl_total_" + trim(i.toString())).val()));
		}
	}
	$("#c_nfe_vl_total_geral").val(formata_moeda(m));
}

function realca_cor_linha(c, indice_row) {
	$("#TR_" + indice_row).css("background-color","palegreen");
	$("#TR_"+indice_row + " td input").css("background-color","palegreen");
	$("#TR_"+indice_row + " td span").css("background-color","palegreen");
	$(c).css("background-color","lightgray");
}

function normaliza_cor_linha(c, indice_row) {
	$("#TR_" + indice_row).css("background-color","");
	$("#TR_"+indice_row + " td input").css("background-color","");
	$("#TR_"+indice_row + " td span").css("background-color","");
	$(c).css("background-color","");
}

function inativa_cor_linha(c, indice_row) {
	$("#TR_" + indice_row).css("background-color","palegray");
	$("#TR_"+indice_row + " td input").css("background-color","palegray");
	$("#TR_"+indice_row + " td span").css("background-color","palegray");
	$(c).css("background-color","lightgray");
}

function alterna_cor_linha(c, indice_row) {
	if ($(c).is(":checked")) {
		normaliza_cor_linha(c, indice_row);
	}
	else {
		inativa_cor_linha(c, indice_row);
	}
}

function exibe_EAN(i) {
    var s;
    var s_ant;
    //alert("EAN: " + $("#c_ean_" + trim(i.toString())).val());

    s_ant = $("#c_ean_" + trim(i.toString())).val();
    s = prompt("EAN:", $("#c_ean_" + trim(i.toString())).val());

    if (s!=s_ant) {
        $("#c_ean_" + trim(i.toString())).val(s);
    }

}

function fESTOQConfirma(f) {
	var s_id;
	var s_aux;
	var s;
	var s_perc_agio = '<%=s_perc_agio%>';
	var iQtdeLinhas = '<%=iQtdeLinhas%>';
	var iQtdeLinhasPreenchidas;


	s_id = "#c_id_nfe_emitente";
	if ($(s_id).val() == "") {
		alert("Selecione o CD!");
		$(s_id).focus();
		return;
	}

	s_id = "#c_fabricante";
	if ($(s_id).val() == "") {
		alert("Informe o código do fabricante!");
		$(s_id).focus();
		return;
	}

	s_id = "#c_documento";
	if ($(s_id).val() == "") {
		alert("Informe o número do documento!");
		$(s_id).focus();
		return;
	}

	iQtdeLinhasPreenchidas = 0;
    for (var i = 1; i <= iQtdeLinhas; i++) {
		s_id = "#c_erp_codigo_" + i.toString();
		s_aux = "#ckb_importa_" + i.toString();
		if (($(s_aux).is(":checked")) && ($(s_id).val() == "")) {
			alert("Informe o código do produto no ERP!");
			$(s_id).focus();
			return;
		}

		if ($(s_aux).is(":checked")) {
		    $(s_aux).val("IMPORTA_ON");
		    s_id = "#c_erp_cst_" + i.toString();
		    if ($(s_id).val() == "") {
		        alert("Informe o CST para entrada no estoque!");
		        $(s_id).focus();
		        return;
		    }
		    iQtdeLinhasPreenchidas = iQtdeLinhasPreenchidas + 1;		    
		}
	}
	$("#iQtdeItensPreenchidos").val(iQtdeLinhasPreenchidas);

	s_aux = $("#c_perc_agio").val();
	if (s_aux == "") {
	    s_aux = "0,0000";
	}
	if (converte_numero(s_aux) != converte_numero(s_perc_agio)) {
	    s = "O ágio esperado para o fabricante atual é " + s_perc_agio + ". Confirma o ágio de " + s_aux + "?";
	    if (!confirm(s)){
	        $("#c_perc_agio").focus();
	        return;
	    }
	}

	fESTOQ.c_FormFieldValues.value = formToStringAll($("#fESTOQ"));

	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";

	f.submit();
}
</script>


<script type="text/javascript">
    function exibeJanelaEAN(idx) {
        var campo_ean_edit;
        var campo_ean_btn;
        var campo_cod_xml;
        campo_ean_edit = "c_ean_" + idx.toString();
        campo_ean_btn = "b_ean_exibe_" + idx.toString();
        campo_cod_xml = "c_nfe_codigo_" + idx.toString();
        $.mostraJanelaEAN(campo_ean_edit, campo_ean_btn, campo_cod_xml);
        ed_EAN.focus();
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
<link href="../Global/eJanelaEditaEAN.css?v=001" rel="stylesheet" type="text/css">

<style type="text/css">
select
{
	margin-left:8px;
}
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
.PLTe{
	margin-left:1pt;
}
.TxtEditavel{
	color: blue;
}
.TxtNfeEmitNome{
	width:640px;
}
.TxtErpFabr{
	width:100px;
	text-align:left;
}
.TxtErpDocumento{
	width:270px;
	margin-left:2pt;
}
.TxtNfeDtHrEmissao{
	width:80px;
	margin-left:2pt;
}
.TxtErpObs{
	width:642px;
	margin-left:2pt;
}
.TxtErpCodigo{
	/*width: 50px;*/
    width: 40px;
	padding-left:4px;
}
.TxtSubtitulo{
	width: 30px;
	padding-left:4px;
    text-align:left;
    color: grey;
}
.TxtErpCst{
	/*width: 30px;*/
    width: 20px;
	text-align:center;
}
.TdErpCodigo{
	/*width:50px;*/
    width: 40px;
	vertical-align: middle;
}
.TdNfeCodigo{
	/*width: 150px;*/
    width: 130px;
	vertical-align: middle;
}
.TdNfeDescricao{
	width: 320px;
	vertical-align: middle;
}
.TdNfeNcm{
	/*width: 60px;*/
    width: 50px;
	vertical-align: middle;
	text-align:center;
}
.TdErpCst{
	/*width: 30px;*/
    width: 20px;
	vertical-align: middle;
	text-align:center;
}
.TdNfeCst{
	/*width: 30px;*/
    width: 20px;
	vertical-align: middle;
	text-align:center;
}
.TdNfeCfop{
	/*width: 40px;*/
    width: 30px;
	vertical-align: middle;
	text-align:center;
}
.TdNfeUnid{
	width: 30px;
	vertical-align: middle;
	text-align:center;
}
.TdNfeQtde{
	/*width: 50px;*/
    width: 40px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeVlUnit{
	width: 70px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeVlTotal{
	width: 70px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeVlBcIcms{
	width: 70px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeVlIcms{
	width: 70px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeVlIpi{
	width: 70px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeAliqIcms{
	/*width: 40px;*/
    width: 30px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeAliqIpi{
	/*width: 40px;*/
    width: 30px;
	vertical-align: middle;
	text-align:right;
}
</style>


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="recalcula_itens()">
<!--<center>-->
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

<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>

<!-- #include file = "../global/JanelaEditaEAN.htm"    -->

<form id="fESTOQ" name="fESTOQ" method="post" action="EstoqueEntradaViaXmlConsiste.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />
<input type="hidden" name="uploaded_file_guid" id="uploaded_file_guid" value="<%=uploaded_file_guid%>" />
<input type="hidden" name="uploaded_file_guid2" id="uploaded_file_guid2" value="<%=uploaded_file_guid2%>" />
<input type="hidden" name="c_nfe_qtde_itens" id="c_nfe_qtde_itens" value="<%=c_nfe_qtde_itens%>"/>
<input type="hidden" name="iQtdeItens" id="iQtdeItens" />
<input type="hidden" name="iQtdeItensPreenchidos" id="iQtdeItensPreenchidos" />
<input type="hidden" name="c_nfe_numero_nf" id="c_nfe_numero_nf" value="<%=s_nfe_numero_nf%>"/>
<input type="hidden" name="c_nfe_emitente_cnpj" id="c_nfe_emitente_cnpj" />
<input type="hidden" name="c_nfe_destinatario_cnpj" id="c_nfe_destinatario_cnpj" value="<%=c_nfe_destinatario_cnpj%>" />
<!--<input type="hidden" name="c_nfe_emitente_nome" id="c_nfe_emitente_nome" />-->
<input type="hidden" name="c_nfe_emitente_nome_fantasia" id="c_nfe_emitente_nome_fantasia" />
<input type="hidden" name="c_dt_hr_emissao" id="c_dt_hr_emissao" value="<%=c_nfe_dt_hr_emissao%>"/>
<input type="hidden" name="c_dt_hr_emissao2" id="c_dt_hr_emissao2" value="<%=c_nfe_dt_hr_emissao2%>"/>
<!--<input type="hidden" name="rb_op_upload" id="rb_op_upload" value="<%=rb_op_upload%>"/>-->
<input type="hidden" name="c_op_upload" id="c_op_upload" value="<%=c_op_upload%>"/>
<input type="hidden" name="arquivo_nfe" id="arquivo_nfe" value="<%=arquivo_nfe%>"/>
<input type="hidden" name="arquivo_nfe2" id="arquivo_nfe2" value="<%=arquivo_nfe2%>"/>

<input type="hidden" name="c_xml_ide__cNF_1" id="c_xml_ide__cNF_1"  value="" />
<input type="hidden" name="c_xml_ide__serie_1" id="c_xml_ide__serie_1" value="" />
<input type="hidden" name="c_xml_ide__nNF_1" id="c_xml_ide__nNF_1" value="" />
<input type="hidden" name="c_xml_emit__CNPJ_1" id="c_xml_emit__CNPJ_1" value="" />
<input type="hidden" name="c_xml_emit__xNome_1" id="c_xml_emit__xNome_1" value="" />
<input type="hidden" name="c_xml_dest__CNPJ_1" id="c_xml_dest__CNPJ_1" value="" />
<input type="hidden" name="c_xml_dest__xNome_1" id="c_xml_dest__xNome_1" value="" />
<input type="hidden" name="c_xml_transp__CNPJ_1" id="c_xml_transp__CNPJ_1" value="" />
<input type="hidden" name="c_xml_det_nItem_1" id="c_xml_det_nItem_1" value="" />
<input type="hidden" name="c_xml_transp__xNome_1" id="c_xml_transp__xNome_1" value="" />
<input type="hidden" name="c_xml_ide__cNF_2" id="c_xml_ide__cNF_2"  value="" />
<input type="hidden" name="c_xml_ide__serie_2" id="c_xml_ide__serie_2" value="" />
<input type="hidden" name="c_xml_ide__nNF_2" id="c_xml_ide__nNF_2" value="" />
<input type="hidden" name="c_xml_emit__CNPJ_2" id="c_xml_emit__CNPJ_2" value="" />
<input type="hidden" name="c_xml_emit__xNome_2" id="c_xml_emit__xNome_2" value="" />
<input type="hidden" name="c_xml_dest__CNPJ_2" id="c_xml_dest__CNPJ_2" value="" />
<input type="hidden" name="c_xml_dest__xNome_2" id="c_xml_dest__xNome_2" value="" />
<input type="hidden" name="c_xml_transp__CNPJ_2" id="c_xml_transp__CNPJ_2" value="" />
<input type="hidden" name="c_xml_det_nItem_2" id="c_xml_det_nItem_2" value="" />
<input type="hidden" name="c_xml_transp__xNome_2" id="c_xml_transp__xNome_2" value="" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Entrada de Mercadorias no Estoque via XML</span>
	<br /><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br />

<!--  CADASTRO DA ENTRADA DE MERCADORIAS NO ESTOQUE VIA XML  -->
<table class="Qx" cellspacing="0" cellpadding="0">
<!--  EMPRESA COMPRADORA / CENTRO DE DISTRIBUIÇÃO  -->
	<tr bgcolor="#FFFFFF" class="trWmsCd">
		<td colspan="3">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td class="MT" align="left" width="50%"><span class="PLTe">Empresa</span>
			<br />
			<select id="c_id_nfe_emitente" name="c_id_nfe_emitente" style="margin-top:4pt;margin-bottom:4pt;min-width:100px;">
			<%=wms_apelido_empresa_nfe_emitente_monta_itens_select(s_id_nfe_emitente)%>
			</select>
			</td>
		</tr>
		</table>
		</td>
	</tr>

<!--  FABRICANTE/DOCUMENTO  -->
	<tr bgcolor="#FFFFFF">
		<td colspan="2" class="MDBE" align="left"><span class="PLTe">Fabricante</span>
			<br /><input name="c_nfe_emitente_nome" id="c_nfe_emitente_nome" class="PLLe TxtNfeEmitNome" readonly tabindex="-1" value="<%=filtra_nome_identificador(c_nfe_emitente_nome)%>" />
		</td>
    	<td class="MDB" style="border-left:0pt;" align="left"><span class="PLTe">% Ágio</span>
		<br><input name="c_perc_agio" id="c_perc_agio" class="PLLe TxtEditavel" maxlength="8" value="<%=s_perc_agio%>" 
            onkeypress="if (digitou_enter(true)) $('#c_fabricante').focus();" onblur="this.value=formata_numero(this.value, 4); recalcula_itens();"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">Cód Fabricante (ERP)</span>
		<br><input name="c_fabricante" id="c_fabricante" class="PLLe TxtErpFabr TxtEditavel" maxlength="4" value="<%=s_fabricante_codigo%>" onkeypress="if ((digitou_enter(true))&&tem_info(this.value)) fESTOQ.c_documento.focus();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"></td>
	<td class="MDB" style="border-left:0pt;" align="left"><span class="PLTe">Documento</span>
		<br><input name="c_documento" id="c_documento" class="PLLe TxtErpDocumento TxtEditavel" maxlength="30" value="<%=s_nfe_numero_nf%>" onkeypress="if (digitou_enter(true)) $('#c_erp_codigo_1').focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"></td>
	<td class="MDB" style="border-left:0pt;" align="left"><span class="PLTe">Emissão</span>
		<br /><input name="c_nfe_dt_emissao" id="c_nfe_dt_emissao" class="PLLe TxtNfeDtHrEmissao" readonly tabindex="-1" value="<%=s_nfe_dt_hr_emissao%>" />
	</td>
	</tr>

<!--  ENTRADA ESPECIAL  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="3" class="MDBE" align="left" nowrap><span class="PLTe">Tipo de Cadastramento</span>
		<br><input type="checkbox" class="rbOpt" tabindex="-1" id="ckb_especial" name="ckb_especial" value="ESPECIAL_ON"
		<%if Not operacao_permitida(OP_CEN_ENTRADA_ESPECIAL_ESTOQUE, s_lista_operacoes_permitidas) then Response.Write " disabled" %>
		><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.ckb_especial.click();">Entrada Especial</span>
	</td>
	</tr>

<!--  OBS -->
	<tr bgColor="#FFFFFF">
	<td colspan="3" class="MDBE" align="left" nowrap><span class="PLTe">Observações</span>
		<br><textarea name="c_obs" id="c_obs" class="PLLe TxtErpObs TxtEditavel" rows="<%=Cstr(MAX_LINHAS_ESTOQUE_OBS)%>"
				onkeypress="limita_tamanho(this,MAX_TAM_T_ESTOQUE_CAMPO_OBS);" onblur="this.value=trim(this.value);"
				></textarea>
	</td>
	</tr>
</table>

<br />
<br />

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0">
	<thead>
	<tr bgColor="#FFFFFF">
	<td align="left">&nbsp;</td>
	<td class="MB TdErpCodigo" align="left" style="vertical-align:bottom;"><span class="PLTe">EAN</span></td>
	<td class="MB TdErpCodigo" align="left" style="vertical-align:bottom;"><span class="PLTe">(FABR)<br />CÓD PROD</span></td>
    <td class="MB TdNfeCodigo" align="center" style="vertical-align:bottom;"><span class="PLTe">(XML)<br />CÓD PROD</span></td>
	<td class="MB TdNfeDescricao" align="center" style="vertical-align:bottom;"><span class="PLTe">DESCRIÇÃO (ERP)</span></td>
	<td class="MB TdNfeNcm" align="left" style="vertical-align:bottom;"><span class="PLTe">NCM/SH</span></td>
	<td class="MB TdNfeCst" align="left" style="vertical-align:bottom;"><span class="PLTe">CST</span></td>
	<td class="MB TdNfeCst" align="left" style="vertical-align:bottom;"><span class="PLTe">CST (ENTR)</span></td>
	<td class="MB TdNfeCfop" align="left" style="vertical-align:bottom;"><span class="PLTe">CFOP</span></td>
	<td class="MB TdNfeQtde" align="left" style="vertical-align:bottom;"><span class="PLTe">QUANT</span></td>
	<td class="MB TdNfeVlUnit" align="left" style="vertical-align:bottom;"><span class="PLTe">VL Nota</span></td>
	<td class="MB TdNfeVlUnit" align="left" style="vertical-align:bottom;"><span class="PLTe">VL Referência</span></td>
	<td class="MB TdNfeAliqIpi" align="left" style="vertical-align:bottom;"><span class="PLTe">A.IPI</span></td>
	<td class="MB TdNfeVlIpi" align="left" style="vertical-align:bottom;"><span class="PLTe">VL IPI</span></td>
	<td class="MB TdNfeAliqIcms" align="left" style="vertical-align:bottom;"><span class="PLTe">A.ICMS</span></td>
    <td class="MB TdNfeVlIpi" align="left" style="vertical-align:bottom;"><span class="PLTe">VL Frete</span></td>
	<td class="MB TdNfeVlTotal" align="left" style="vertical-align:bottom;"><span class="PLTe">VL TOTAL</span></td>
	</tr>
	</thead>
	<tbody>
<% for i=1 to iQtdeLinhas %>
	<tr id="TR_<%=Cstr(i)%>">
	<td align="left">
		<% if i <= iQtdeItens then %>
		    <input type="checkbox" name="ckb_importa_<%=Cstr(i)%>" id="ckb_importa_<%=Cstr(i)%>" 
		    value="IMPORTA_ON" checked="checked">
        <% else %>
            <input type="checkbox" name="ckb_importa_<%=Cstr(i)%>" id="ckb_importa_<%=Cstr(i)%>" 
            value="">
        <% end if %>
        <input type="hidden" name="c1_xml_prod_cProd_<%=Cstr(i)%>" id="c1_xml_prod_cProd_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c1_xml_prod_cEAN_<%=Cstr(i)%>" id="c1_xml_prod_cEAN_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c1_xml_prod__NCM_<%=Cstr(i)%>" id="c1_xml_prod__NCM_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c1_xml_prod__CFOP_<%=Cstr(i)%>" id="c1_xml_prod__CFOP_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c1_xml_prod__qCom_<%=Cstr(i)%>" id="c1_xml_prod__qCom_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c1_xml_prod__vUnCom_<%=Cstr(i)%>" id="c1_xml_prod__vUnCom_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c1_xml_prod__vProd_<%=Cstr(i)%>" id="c1_xml_prod__vProd_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c1_xml_prod__vFrete_<%=Cstr(i)%>" id="c1_xml_prod__vFrete_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c1_xml_imposto__pICMS_<%=Cstr(i)%>" id="c1_xml_imposto__pICMS_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c1_xml_imposto__pIPI_<%=Cstr(i)%>" id="c1_xml_imposto__pIPI_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c1_xml_imposto__vIPI_<%=Cstr(i)%>" id="c1_xml_imposto__vIPI_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c2_xml_prod_cProd_<%=Cstr(i)%>" id="c2_xml_prod_cProd_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c2_xml_prod_cEAN_<%=Cstr(i)%>" id="c2_xml_prod_cEAN_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c2_xml_prod__NCM_<%=Cstr(i)%>" id="c2_xml_prod__NCM_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c2_xml_prod__CFOP_<%=Cstr(i)%>" id="c2_xml_prod__CFOP_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c2_xml_prod__qCom_<%=Cstr(i)%>" id="c2_xml_prod__qCom_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c2_xml_prod__vUnCom_<%=Cstr(i)%>" id="c2_xml_prod__vUnCom_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c2_xml_prod__vProd_<%=Cstr(i)%>" id="c2_xml_prod__vProd_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c2_xml_prod__vFrete_<%=Cstr(i)%>" id="c2_xml_prod__vFrete_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c2_xml_imposto__pICMS_<%=Cstr(i)%>" id="c2_xml_imposto__pICMS_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c2_xml_imposto__pIPI_<%=Cstr(i)%>" id="c2_xml_imposto__pIPI_<%=Cstr(i)%>" value="" />
        <input type="hidden" name="c2_xml_imposto__vIPI_<%=Cstr(i)%>" id="c2_xml_imposto__vIPI_<%=Cstr(i)%>" value="" />
	</td>
    <td class="MDBE" align="center">
        <input type="hidden" name="c_ean_<%=Cstr(i)%>" id="c_ean_<%=Cstr(i)%>" value=""/>
        <a name="b_ean_exibe_<%=Cstr(i)%>" id="b_ean_exibe_<%=Cstr(i)%>" href="javascript:exibeJanelaEAN(<%=Cstr(i)%>)" >
		<img src="../botao/view_bottom.PNG" border="0"></a>
	</td>
	<td class="MDBE TdErpCodigo" align="left">
        <!--<input type="hidden" name="c_ean_<%=Cstr(i)%>" id="c_ean_<%=Cstr(i)%>" value=""/>-->
		<input name="c_erp_codigo_<%=Cstr(i)%>" id="c_erp_codigo_<%=Cstr(i)%>" class="PLLe TxtErpCodigo TxtEditavel" maxlength="8" value=""
			onkeypress="if (digitou_enter(true)) {$('#c_erp_cst_<%=Cstr(i)%>').focus();}"
			onfocus="this.select();realca_cor_linha(this,<%=Cstr(i)%>);"
			onblur="this.value=normaliza_produto(this.value);normaliza_cor_linha(this,<%=Cstr(i)%>); " />
	</td>
	<td class="MDB TdNfeCodigo" align="left">
		<input type="hidden" name="c_nfe_nItem_<%=Cstr(i)%>" id="c_nfe_nItem_<%=Cstr(i)%>" />
		<input name="c_nfe_codigo_<%=Cstr(i)%>" id="c_nfe_codigo_<%=Cstr(i)%>" class="PLLe TdNfeCodigo <%=s_classe_editavel%>" <%=s_valor_readonly%> />
	</td>
	<td class="MDB TdNfeDescricao" align="left">
        <input name="c_descricao_erp_<%=Cstr(i)%>" id="c_descricao_erp_<%=Cstr(i)%>" class="PLLe" style="width:99%;" readonly />
	</td>
	<td class="MDB TdNfeNcm" align="left">
        <input name="c_nfe_ncm_sh_<%=Cstr(i)%>" id="c_nfe_ncm_sh_<%=Cstr(i)%>" class="PLLe TdNfeNcm <%=s_classe_editavel%>" <%=s_valor_readonly%> maxlength="8" />
	</td>
	<td class="MDB TdNfeCst" align="left">
        <input name="c_nfe_cst_<%=Cstr(i)%>" id="c_nfe_cst_<%=Cstr(i)%>" class="PLLe TdNfeNcm <%=s_classe_editavel%>" <%=s_valor_readonly%> />
	</td>
	<td class="MDB TdErpCst" align="left">
		<input name="c_erp_cst_<%=Cstr(i)%>" id="c_erp_cst_<%=Cstr(i)%>" class="PLLe TxtErpCst TxtEditavel" maxlength="3"
			onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==<%=Cstr(iQtdeItens)%>) {bCONFIRMA.focus();} else {$('#c_erp_codigo_<%=Cstr(i+1)%>').focus();}}"
			onfocus="this.select();realca_cor_linha(this,<%=Cstr(i)%>);"
			onblur="this.value=trim(this.value);normaliza_cor_linha(this,<%=Cstr(i)%>);" />
	</td>
	<td class="MDB TdNfeCfop" align="left">
        <input name="c_nfe_cfop_<%=Cstr(i)%>" id="c_nfe_cfop_<%=Cstr(i)%>" class="PLLe TdNfeCfop <%=s_classe_editavel%>" <%=s_valor_readonly%> />
	</td>
	<td class="MDB TdNfeQtde" align="left">
        <input name="c_nfe_qtde_<%=Cstr(i)%>" id="c_nfe_qtde_<%=Cstr(i)%>" class="PLLe TdNfeQtde <%=s_classe_editavel%>" <%=s_valor_readonly%>
		onblur="recalcula_linha(<%=Cstr(i)%>); recalcula_total_nf(); recalcula_total();" />
	</td>
	<td class="MDB TdNfeVlUnit" align="left">
        <input name="c_nfe_vl_unitario_nota_<%=Cstr(i)%>" id="c_nfe_vl_unitario_nota_<%=Cstr(i)%>" class="PLLe TdNfeVlUnit <%=s_classe_editavel%>" <%=s_valor_readonly%> />
	</td>
	<td class="MDB TdNfeVlUnit" align="left">
        <input name="c_nfe_vl_unitario_<%=Cstr(i)%>" id="c_nfe_vl_unitario_<%=Cstr(i)%>" class="PLLe TdNfeVlUnit TxtEditavel"
        onblur="this.value=formata_moeda(this.value); recalcula_linha(<%=Cstr(i)%>); recalcula_total_nf(); recalcula_total();" />
	</td>
	<td class="MDB TdNfeAliqIpi" align="left">
        <input type="hidden" name="c_nfe_aliq_ipi_ori_<%=Cstr(i)%>" id="c_nfe_aliq_ipi_ori_<%=Cstr(i)%>" />
        <input name="c_nfe_aliq_ipi_<%=Cstr(i)%>" id="c_nfe_aliq_ipi_<%=Cstr(i)%>" class="PLLe TdNfeAliqIpi TxtEditavel"
         onblur="this.value=formata_numero(this.value, 0); if (converte_numero(this.value) > 100) {alert('Alíquota de IPI maior que 100%!');}; if (this.value != c_nfe_aliq_ipi_ori_<%=Cstr(i)%>.value) {recalcula_itens(); recalcula_total_nf();}" />
	</td>
	<td class="MDB TdNfeVlIpi" align="left">
        <input type="hidden" name="c_nfe_vl_ipi_ori_<%=Cstr(i)%>" id="c_nfe_vl_ipi_ori_<%=Cstr(i)%>" />
        <input name="c_nfe_vl_ipi_<%=Cstr(i)%>" id="c_nfe_vl_ipi_<%=Cstr(i)%>" class="PLLe TdNfeVlIpi TxtEditavel" 
        onblur="this.value=formata_moeda(this.value); if (this.value != c_nfe_vl_ipi_ori_<%=Cstr(i)%>.value) {recalcula_itens(); recalcula_total();}" />
	</td>
	<td class="MDB TdNfeAliqIcms" align="left">
        <input name="c_nfe_aliq_icms_<%=Cstr(i)%>" id="c_nfe_aliq_icms_<%=Cstr(i)%>" class="PLLe TdNfeAliqIcms TxtEditavel"
         onblur="this.value=formata_numero(this.value, 0); if (converte_numero(this.value) > 100) {alert('Alíquota de ICMS maior que 100%!');}" />

	</td>
	<td class="MDB TdNfeVlIpi" align="left">
        <input type="hidden" name="c_nfe_vl_frete_ori_<%=Cstr(i)%>" id="c_nfe_vl_frete_ori_<%=Cstr(i)%>" />
        <input name="c_nfe_vl_frete_<%=Cstr(i)%>" id="c_nfe_vl_frete_<%=Cstr(i)%>" class="PLLe TdNfeVlIpi TxtEditavel" 
        onblur="this.value=formata_moeda(this.value); if (this.value != c_nfe_vl_frete_ori_<%=Cstr(i)%>.value) {recalcula_itens(); recalcula_total();}" />
	</td>
	<td class="MDB TdNfeVlTotal" align="left">
        <input name="c_nfe_vl_total_<%=Cstr(i)%>" id="c_nfe_vl_total_<%=Cstr(i)%>" class="PLLe TdNfeVlTotal" readonly />
	</td>
	</tr>
<% next %>
	<tbody>
	<tfoot>
	<tr>
	
	<td colspan="10" class="MD">&nbsp;</td>

	<td class="MDB" align="left"><p class="Cd">Total NF</p></td>
	
	<td class="MDB" align="right"><input name="c_total_nf" id="c_total_nf" class="PLLd" style="width:62px;color:black;" 
	value=""></td>
	<td>&nbsp;</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
	<td class="MD">&nbsp;</td>
	<td class="MDB" align="right"><input name="c_nfe_vl_total_geral" id="c_nfe_vl_total_geral" class="PLLd" style="width:70px;color:black;"
		value="" readonly tabindex=-1 /></td>
	
	</tr>
	</tfoot>
</table>

</form>

<br />

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br />

<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="retorna para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fESTOQConfirma(fESTOQ)" title="vai para a página seguinte">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>

</center>
</body>
<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS

	cn.Close
	set cn = nothing
%>