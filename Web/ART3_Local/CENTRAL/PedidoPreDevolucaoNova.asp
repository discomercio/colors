<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ==================================================
'	  P E D I D O P R E D E V O L U C A O N O V A. A S P
'     ==================================================
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

	dim s, usuario, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_PRE_DEVOLUCAO_CADASTRAMENTO, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
	dim request_guid
	request_guid = Trim(Request("request_guid"))

	dim i, n
	dim s_fabricante, s_produto, s_descricao, s_descricao_html, s_qtde, s_vl_unitario, s_vl_TotalItem, m_TotalItem, m_TotalDestePedido
	dim s_cor, s_devolucao_anterior, s_devolucao_pendente, s_readonly
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
    If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

    dim s_sessionToken
	s_sessionToken = ""
	s = "SELECT Convert(varchar(36), SessionTokenModuloCentral) AS SessionTokenModuloCentral FROM t_USUARIO WHERE (usuario = '" & usuario & "')"
    if rs.State <> 0 then rs.Close
    rs.Open s,cn
	if Not rs.Eof then s_sessionToken = Trim("" & rs("SessionTokenModuloCentral"))

	dim r_pedido, v_item, v_devol, alerta, msg_erro
	alerta=""
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
	else
		if Not le_pedido_item(pedido_selecionado, v_item, msg_erro) then alerta = msg_erro
		end if

    dim r_indicador, ha_indicador
    ha_indicador = False
    if alerta = "" then
        ha_indicador = le_indicador(r_pedido.indicador, r_indicador, msg_erro)
        end if

	if alerta = "" then
		if r_pedido.st_entrega <> ST_ENTREGA_ENTREGUE then
			alerta = "Pedido " & pedido_selecionado & " não consta como entregue, portanto, não é possível processar a sua devolução."
			end if
		end if
	
	if alerta = "" then
		redim v_devol(Ubound(v_item))
		for i=Lbound(v_devol) to Ubound(v_devol)
			set v_devol(i) = New cl_ITEM_DEVOLUCAO_MERCADORIAS
			v_devol(i).pedido		= v_item(i).pedido
			v_devol(i).fabricante	= v_item(i).fabricante
			v_devol(i).produto		= v_item(i).produto
			v_devol(i).qtde			= v_item(i).qtde
			next
		
		if Not estoque_verifica_mercadorias_para_devolucao(v_devol, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
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



<html>


<head>
	<title>CENTRAL<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>



<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
var LOCAL_COLETA__CLIENTE = "001";
var LOCAL_COLETA__PARCEIRO = "002";
var LOCAL_COLETA__TRANSPORTADORA = "003";
var LOCAL_COLETA__NAO_HAVERA_COLETA = "004";
var CREDITO_TRANSACAO__ESTORNO = "002";
var CREDITO_TRANSACAO__REEMBOLSO = "003";
var TAXA_ADMINISTRATIVA__NAO = "0";
var TAXA_ADMINISTRATIVA__SIM = "1";
var TAXA_FORMA_PAGAMENTO__ABATIMENTO_CREDITO = "003";
var TAXA_RESPONSAVEL__CLIENTE = "003";
var PROCEDIMENTO__ACERTO_INTERNO = "002";

    $(function () {
        $('#c_taxa_valor').val(formata_moeda(0));
        $('#c_taxa_forma_pagto').attr('disabled', true);
        $('#c_taxa_percentual').attr('disabled', true);
        $('#c_taxa_responsavel').attr('disabled', true);
        $('#c_taxa_valor').attr('disabled', true);

        $('#c_taxa').change(function () {
            if ($(this).val() == TAXA_ADMINISTRATIVA__SIM) {
                $('#c_taxa_forma_pagto').attr('disabled', false);
                $('#c_taxa_percentual').attr('disabled', false);
                $('#c_taxa_responsavel').attr('disabled', false);
                $('#c_taxa_valor').attr('disabled', false);
            }
            else {
                $("#c_taxa_forma_pagto").prop('selectedIndex', 0);
                $("#c_taxa_percentual").prop('selectedIndex', 0);
                $("#c_taxa_responsavel").prop('selectedIndex', 0);
                $('#c_taxa_forma_pagto').attr('disabled', true);
                $('#c_taxa_percentual').attr('disabled', true);
                $('#c_taxa_responsavel').attr('disabled', true);
                $('#c_taxa_valor').attr('disabled', true);
            }
            recalcula_valor_taxa();
        });

        $('#c_credito_transacao').change(function () {
            if ($(this).val() == CREDITO_TRANSACAO__REEMBOLSO) {
                $('#trDadosBancarios').show();
                $('#c_cliente_banco').attr('disabled', false);
                $('#c_cliente_agencia').attr('disabled', false);
                $('#c_cliente_conta').attr('disabled', false);
                $('#c_cliente_favorecido').attr('disabled', false);
            }
            else {
                $('#trDadosBancarios').hide();
                $('#c_cliente_banco').attr('disabled', true);
                $('#c_cliente_agencia').attr('disabled', true);
                $('#c_cliente_conta').attr('disabled', true);
                $('#c_cliente_favorecido').attr('disabled', true);
            }
        });

        // em caso de reload da página, assegurar que os campos de taxa
        // estão habilitados em caso que taxa = sim
        if ($('#c_taxa').val() == TAXA_ADMINISTRATIVA__SIM) {
            $('#c_taxa_forma_pagto').attr('disabled', false);
            $('#c_taxa_percentual').attr('disabled', false);
            $('#c_taxa_responsavel').attr('disabled', false);
            $('#c_taxa_valor').attr('disabled', false);
            recalcula_valor_taxa();
        }
    });
</script>

<script language="JavaScript" type="text/javascript">

<%=monta_funcao_js_normaliza_numero_pedido_e_sufixo%>

function posiciona_foco_prox_linha( idx ) {
var f,i, i_base;
	f=fPED;
	i_base=idx;
	if (i_base<0) i_base=0;
	for (i=i_base; i<f.c_qtde_devolucao.length; i++) {
		if (!f.c_qtde_devolucao[i].readOnly) {
			f.c_qtde_devolucao[i].focus();
			return true;
			}
		}
	return false;
}

function verifica_qtde_max( campo, idx ) {
var f,p,n;
	if (campo.readOnly) return true;
	f=fPED;
	p=converte_numero(idx)-1;
	n=converte_numero(f.c_qtde[p].value)-(converte_numero(f.c_devolucao_anterior[p].value)+converte_numero(f.c_devolucao_pendente[p].value));
	if (converte_numero(campo.value)>n) {
		alert("A quantidade máxima para devolução é de " + n + " unidade(s)!!");
		campo.focus();
		return false;
		}
	return true;
}

function recalcula_valores() {
var f, i, n, m, t;
	f=fPED;
	t=0
	for (i=0; i<f.c_qtde_devolucao.length; i++) {
		if (!f.c_qtde_devolucao[i].readOnly) {
			n=converte_numero(f.c_qtde_devolucao[i].value);
			m=converte_numero(f.c_vl_unitario[i].value);
			m=n*m;
			t=t+m;
			f.c_vl_total[i].value=formata_moeda(m);
			}
		}
	f.c_total_geral.value=formata_moeda(t);
    
	recalcula_valor_taxa();
}

function recalcula_valor_taxa() {
var f, t, n, m;
    f=fPED;
    t=converte_numero(f.c_total_geral.value);    
    if (f.c_taxa_percentual.value != "") {
        n=converte_numero(f.c_taxa_percentual.value);
        m=(n/100)*t;
        f.c_taxa_valor.value=formata_moeda(m);
    }
    else {
        f.c_taxa_valor.value = formata_moeda(0);
    }
}

function preencheEndColeta(f, local_coleta) {
    var pedido_st_end_entrega;
    pedido_st_end_entrega = "<%=r_pedido.st_end_entrega%>";

    f.c_coleta_endereco.value = "";
    f.c_coleta_endereco_numero.value = "";
    f.c_coleta_bairro.value = "";
    f.c_coleta_cep.value = "";
    f.c_coleta_cidade.value = "";
    f.c_coleta_complemento.value = "";
    f.c_coleta_uf.selectedIndex = -1;
    document.getElementById("c_transportadora").innerText = "";

    if (local_coleta == LOCAL_COLETA__PARCEIRO) {
        f.c_coleta_endereco.value = "<%=substitui_caracteres(r_indicador.endereco, chr(34), chr(39))%>";
        f.c_coleta_endereco_numero.value = "<%=r_indicador.endereco_numero%>";
        f.c_coleta_bairro.value = "<%=substitui_caracteres(r_indicador.bairro, chr(34), chr(39))%>";
        f.c_coleta_cep.value = "<%=cep_formata(r_indicador.cep)%>";
        f.c_coleta_cidade.value = "<%=r_indicador.cidade%>";
        f.c_coleta_complemento.value = "<%=substitui_caracteres(r_indicador.endereco_complemento, chr(34), chr(39))%>";
        f.c_coleta_uf.value = "<%=r_indicador.uf%>";
    }
    else if (local_coleta == LOCAL_COLETA__CLIENTE) {
        if (pedido_st_end_entrega != "0") {
            f.c_coleta_endereco.value = "<%=substitui_caracteres(r_pedido.EndEtg_endereco, chr(34), chr(39))%>";
            f.c_coleta_endereco_numero.value = "<%=r_pedido.EndEtg_endereco_numero%>";
            f.c_coleta_bairro.value = "<%=substitui_caracteres(r_pedido.EndEtg_bairro, chr(34), chr(39))%>";
            f.c_coleta_cep.value = "<%=cep_formata(r_pedido.EndEtg_cep)%>";
            f.c_coleta_cidade.value = "<%=r_pedido.EndEtg_cidade%>";
            f.c_coleta_complemento.value = "<%=substitui_caracteres(r_pedido.EndEtg_endereco_complemento, chr(34), chr(39))%>";
            f.c_coleta_uf.value = "<%=r_pedido.EndEtg_uf%>";
        }
        else {
            f.c_coleta_endereco.value = "<%=substitui_caracteres(r_pedido.endereco_logradouro, chr(34), chr(39))%>";
            f.c_coleta_endereco_numero.value = "<%=r_pedido.endereco_numero%>";
            f.c_coleta_bairro.value = "<%=substitui_caracteres(r_pedido.endereco_bairro, chr(34), chr(39))%>";
            f.c_coleta_cep.value = "<%=cep_formata(r_pedido.endereco_cep)%>";
            f.c_coleta_cidade.value = "<%=r_pedido.endereco_cidade%>";
            f.c_coleta_complemento.value = "<%=substitui_caracteres(r_pedido.endereco_complemento, chr(34), chr(39))%>";
            f.c_coleta_uf.value = "<%=r_pedido.endereco_uf%>";
        }
    }
    else if (local_coleta == LOCAL_COLETA__TRANSPORTADORA) {
        document.getElementById("c_transportadora").innerText = "<%=r_pedido.transportadora_id%> - <%=x_transportadora(r_pedido.transportadora_id)%>";
    }

}
function reloadFileReader(idx) {
    var file_image_extensions = "|jpeg|jpg|png|bmp|gif|tif|tiff|";
    var img_src, img_width, img_height;
    img_src = "";
    if (typeof (FileReader) != "undefined") {
        var image_holder = $('#image-holder-arquivo' + idx);
        var file_name = $('#arquivo' + idx).val();
        image_holder.empty();

        var file_extension = file_name.substring(file_name.lastIndexOf(".")+1, file_name.length);
        file_extension = file_extension.toLowerCase();

        var reader = new FileReader();
        reader.onload = function (e) {
            if (file_image_extensions.indexOf(file_extension) == -1) {
                if (file_extension == "pdf") {
                    img_src = "../IMAGEM/file_pdf_150x150.png";
                    img_width = "150";
                    img_height = "150";
                }
                else {
                    img_src = "../IMAGEM/file_150x150.png";
                    img_width = "150";
                    img_height = "150";
                }
            }
            else {
                img_src = e.target.result;
                img_width = "200";
                img_height = "180";
            }


            $("<img />", {
                "src": img_src,
                "class": "thumb-image",
                "width": img_width,
                "height": img_height
            }).appendTo(image_holder);
        }
        image_holder.show();
        reader.readAsDataURL($('#arquivo' + idx)[0].files[0]);
        blnFotoPendente = true;
    }
}

function fPEDConfirma( f ) {
var b, i, n;
var serverVariableUrl;
var jsonResponse;
var s_stored_file_guid, blnUploadOk, blnEnviaRequisicaoUpload, inputFile, qtde_fotos;
serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
serverVariableUrl = serverVariableUrl.toUpperCase();
serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));

	b=false;
	for (i=0; i<f.c_qtde_devolucao.length; i++) {
		if (!f.c_qtde_devolucao[i].readOnly) {
			if (converte_numero(f.c_qtde_devolucao[i].value)>0) {
				b=true;
				}
			n=converte_numero(f.c_qtde[i].value)-converte_numero(f.c_devolucao_anterior[i].value);
			if (converte_numero(f.c_qtde_devolucao[i].value) > n) {
				alert("A quantidade máxima para devolução é de " + n + " unidades!!");
				f.c_qtde_devolucao[i].focus();
				return;
				}
			}
		}
	if (!b) {
		alert("Não foi especificada nenhuma mercadoria para devolução!!");
		return;
	}

	if (f.c_procedimento.value==""){
	    alert("Informe o procedimento!!");
	    f.c_procedimento.focus();
	    return;
	}

	if (f.c_local_coleta.value==""){
	    alert("Informe o local de coleta!!");
	    f.c_local_coleta.focus();
	    return;
	}

	if (f.c_motivo_devolucao.value==""){
	    alert("Informe o motivo da devolução!!");
	    f.c_motivo_devolucao.focus();
	    return;
	}

	if(trim(f.c_motivo_descricao.value)==""){
	    alert("Insira uma descrição para o motivo da devolução!!");
	    f.c_motivo_descricao.focus();
	    return;
	}

	if(f.c_taxa.value==TAXA_ADMINISTRATIVA__SIM){
	    if (f.c_taxa_forma_pagto.value==""){
	        alert("Informe a forma de pagamento da taxa administrativa!!");
	        f.c_taxa_forma_pagto.focus();
	        return;
	    }
	    if (f.c_taxa_forma_pagto.value==TAXA_FORMA_PAGAMENTO__ABATIMENTO_CREDITO) {
	        if(f.c_taxa_responsavel.value!=TAXA_RESPONSAVEL__CLIENTE) {
	            alert("Quando a forma de pagamento da taxa for 'abatimento no crédito', obrigatoriamente o responsável deve ser o cliente!!");
	            f.c_taxa_responsavel.focus();
	            return;
	        }
	    }
	    if (f.c_taxa_percentual.value==""){
	        alert("Informe o percentual da taxa administrativa!!");
	        f.c_taxa_percentual.focus();
	        return;
	    }
	    if (f.c_taxa_responsavel.value==""){
	        alert("Informe o responsável pelo pagamento da taxa administrativa!!");
	        f.c_taxa_responsavel.focus();
	        return;
	    }
	}

	if (f.c_credito_transacao.value == "") {
	    alert("Informe o tipo de transação do crédito!!");
	    f.c_credito_transacao.focus();
	    return;
    }

	if (f.c_credito_transacao.value == CREDITO_TRANSACAO__REEMBOLSO) {
		if (retorna_so_digitos(f.c_cliente_banco.value)==""){
	        alert("Preencha o código do banco!!");
	        f.c_cliente_banco.focus();
	        return;
		}
		if (f.c_cliente_banco.value != retorna_so_digitos(f.c_cliente_banco.value)) {
			alert("Código do banco é inválido!!");
			f.c_cliente_banco.focus();
			return;
		}
	    if (f.c_cliente_agencia.value==""){
	        alert("Preencha o número da agência!!");
	        f.c_cliente_agencia.focus();
	        return;
	    }
	    if (f.c_cliente_conta.value==""){
	        alert("Preencha o número da conta!!");
	        f.c_cliente_conta.focus();
	        return;
	    }
	    if (f.c_cliente_favorecido.value==""){
	        alert("Preencha o favorecido da conta!!");
	        f.c_cliente_favorecido.focus();
	        return;
	    }
	}
	else if (f.c_credito_transacao.value == CREDITO_TRANSACAO__ESTORNO) {
	    if (f.c_pedido_possui_parcela_cartao.value != "1") {
	        alert("A transação do crédito não pode ser 'Estorno' porque o pedido não possui pagamento via cartão de crédito!!");
	        f.c_credito_transacao.focus();
	        return;
	    }
	}
	
	blnEnviaRequisicaoUpload = false;
	qtde_fotos = 0;
	n = $('input[type=file]').length;
	for (i = 0; i< n;i++) {
	    inputFile = $('input[type=file]')[i];
	    if (inputFile.value != "") {
	        blnEnviaRequisicaoUpload = true;
	        qtde_fotos++;
	    }
	}
	if ((f.c_local_coleta.value != LOCAL_COLETA__TRANSPORTADORA) && (f.c_local_coleta.value != LOCAL_COLETA__NAO_HAVERA_COLETA)) {
	    if (f.c_procedimento.value != PROCEDIMENTO__ACERTO_INTERNO) {
	        if (qtde_fotos < 2) {
	            alert("Insira no mínimo 2 (duas) fotos!!");
	            return;
	        }
	    }
	}

	b=window.confirm("Confirma o cadastro da pré-devolução?");
	if (!b) return;

	window.status = "Aguarde ...";
	dCONFIRMA.style.visibility="hidden";

	if (blnEnviaRequisicaoUpload) {
	    var form = $('#fPED');
	    var fd = new FormData(form[0]);
	    blnUploadOk = false;
	    s_stored_file_guid = "";

	    $.ajax({
	        type: form.attr('method'),
			url: '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/UploadFile/PostFile',
	        enctype: 'multipart/form-data',
	        processData: false,
	        contentType: false,
	        data: fd,
	        async: false,
	        success: function (resp) {
	            jsonResponse = JSON.parse(resp);
	            if (jsonResponse.Status == "OK")
	            {
	                $.each(jsonResponse.files, function() {
	                    $.each(this, function(k, v) {
	                        if (k == "stored_file_guid") {
	                            s_stored_file_guid = s_stored_file_guid + v + "|";
	                        }
	                    });
	                });
	                blnUploadOk = true;
	                $('#upload_file_guid_returned').val(s_stored_file_guid);
	            }
	        },
	        error: function (resp, cod, msgErro) {
	            alert("Erro ao salvar arquivos de foto no servidor!!\n\n" + msgErro);
	        }
	    });
	}
	if ((blnUploadOk) || (!blnEnviaRequisicaoUpload)) {
	    f.action="pedidopredevolucaonovaconfirma.asp";
	    f.submit();
	}
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
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ************************************************* -->
<!-- **********  PÁGINA EDITAR QUANTIDADES  ********** -->
<!-- ************************************************* -->
<body onload="fPED.c_procedimento.focus();">
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="c_pedido_cliente" id="c_pedido_cliente" value="<%=r_pedido.id_cliente%>" />
<input type="hidden" name="c_pedido_indicador" id="c_pedido_indicador" value="<%=r_pedido.indicador%>" />
<input type="hidden" name="c_pedido_possui_parcela_cartao" id="c_pedido_possui_parcela_cartao" value="<%=r_pedido.st_forma_pagto_possui_parcela_cartao%>" />
<input type="hidden" name="upload_parameter__is_temp_file" value="0" />
<input type="hidden" name="upload_parameter__is_confirmation_required" value="1" />
<input type="hidden" name="upload_parameter__folder_name" value="PEDIDO_DEVOLUCAO_FOTOS" />
<input type="hidden" name="upload_parameter__user_id" value="<%=usuario%>" />
<input type="hidden" name="upload_parameter__sessionToken" value="<%=s_sessionToken%>" />
<input type="hidden" name="upload_file_guid_returned" id="upload_file_guid_returned" />
<input type="hidden" name="request_guid" id="request_guid" value="<%=request_guid%>" />


<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<%=MontaHeaderIdentificacaoPedido(pedido_selecionado, r_pedido, 852)%>
<br>

<!--  DESCRIÇÃO DA OPERAÇÃO -->
<table width="649" cellPadding="0" CellSpacing="0">
<tr><td><p class="Expl">PRÉ-DEVOLUÇÃO</p></td></tr>
<tr><td>
	<p class="Expl">A devolução de mercadorias só é possível em pedido que já tenha sido entregue ao cliente.</p>
	</td>
</tr>
</table>
<br>

<table class="Q" style="width: 649px;" cellpadding="0" cellspacing="0">
    <tr>
        <td colspan="4" class="MB" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">DADOS DO PARCEIRO</span></td>
    </tr>
    <tr>
        <% if ha_indicador then
            s = r_indicador.apelido & " - " & r_indicador.razao_social_nome_iniciais_em_maiusculas
            s_cor = "black"
        else
            s = "Não há parceiro vinculado ao pedido."
            s_cor = "gray"
            end if%>
        <td align="left" colspan="4" style="padding:5px;">
            <input type="text" readonly tabindex=-1 class="PLLe" style="width:320px; margin-left: 10px; color:<%=s_cor%>;" value="<%=s%>">
        </td>
    </tr>

    <tr>
        <td colspan="4" class="MB MC" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">PROCEDIMENTO</span></td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;"><p class="Cd">Procedimento</p></td>
        <td style="padding: 5px;" colspan="3">
            <select id="c_procedimento" name="c_procedimento" style='width:240px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;'>
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__PROCEDIMENTO, "")%>
            </select>
        </td>
    </tr>

    <tr>
        <td colspan="4" class="MB MC" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">DADOS DE COLETA</span></td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;"><p class="Cd">Local de Coleta</p></td>
        <td style="padding: 5px;" colspan="3">
            <select id="c_local_coleta" name="c_local_coleta" style='width:240px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;'
                onchange="preencheEndColeta(fPED, this.value);">
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__LOCAL_COLETA, "")%>
            </select>&nbsp;&nbsp;&nbsp;
            <span id="c_transportadora" class="C"></span>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Endereço:</p></td>
        <td align="left" style="padding:5px;width:260px;" class="MC">
            <input type="text" maxlength="80" name="c_coleta_endereco" id="c_coleta_endereco" class="PLLe" style="width:240px; margin-left: 10px;">
        </td>
        <td style="padding: 5px;width:100px;" class="MC ME"><p class="Cd">Nº:</p></td>
        <td align="left" style="padding:5px;width:189px;" class="MC">
            <input type="text" maxlength="20" name="c_coleta_endereco_numero" id="c_coleta_endereco_numero" class="PLLe" style="width:100px; margin-left: 10px;">
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Bairro:</p></td>
        <td align="left" style="padding:5px;width:260px;" class="MC">
            <input type="text" maxlength="72" name="c_coleta_bairro" id="c_coleta_bairro" class="PLLe" style="width:240px; margin-left: 10px;">
        </td>
        <td style="padding: 5px;width:100px;" class="MC ME"><p class="Cd">CEP:</p></td>
        <td align="left" style="padding:5px;width:189px;" class="MC">
            <input type="text" maxlength="9" name="c_coleta_cep" id="c_coleta_cep" class="PLLe" style="width:100px; margin-left: 10px;" onblur="if (!cep_ok(this.value)){ alert('CEP inválido!!'); this.focus(); };">
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Cidade:</p></td>
        <td align="left" style="padding:5px;width:260px;" class="MC">
            <input type="text" maxlength="60" name="c_coleta_cidade" id="c_coleta_cidade" class="PLLe" style="width:240px; margin-left: 10px;">
        </td>
        <td style="padding: 5px;width:100px;" class="MC ME"><p class="Cd">UF:</p></td>
        <td align="left" style="padding:5px;width:189px;" class="MC">
            <select name="c_coleta_uf" id="c_coleta_uf" class="PLLe" style="width:50px; margin-left: 10px;">
                <%=UF_monta_itens_select(Null)%>                
            </select>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Complemento:</p></td>
        <td align="left" style="padding:5px;width:260px;" class="MC" colspan="3">
            <input type="text" maxlength="60" name="c_coleta_complemento" id="c_coleta_complemento" class="PLLe" style="width:380px; margin-left: 10px;">
        </td>
    </tr>
</table>
<br />
<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB" valign="bottom"><p class="PLTe">Fabr</p></td>
	<td class="MB" valign="bottom"><p class="PLTe">Produto</p></td>
	<td class="MB" valign="bottom"><p class="PLTe">Descrição</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Qtde</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Devol<br>Anter</p></td>
    <td class="MB" valign="bottom"><p class="PLTd" title="Devolução pendente">Devol<br />Pend</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Devolver</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Valor<br>Unitário</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Total<br>Devolução</p></td>
	</tr>

<% m_TotalDestePedido=0
   n = Lbound(v_item)-1
   for i=1 to MAX_ITENS 
	 n = n+1
	 s_cor = "black"
	 s_readonly = "readonly tabindex=-1"
	 if n <= Ubound(v_item) then
		with v_item(n)
			s_fabricante=.fabricante
			s_produto=.produto
			s_descricao=.descricao
			s_descricao_html=produto_formata_descricao_em_html(.descricao_html)
			s_qtde=.qtde
			s_vl_unitario=formata_moeda(.preco_NF)
			m_TotalItem=0
			s_vl_TotalItem=formata_moeda(m_TotalItem)
			m_TotalDestePedido=m_TotalDestePedido + m_TotalItem
			end with
		s_devolucao_anterior=""
        s_devolucao_pendente=""
		with v_devol(n)
			if .qtde_devolvida_anteriormente<>0 then 
				s_devolucao_anterior=Cstr(.qtde_devolvida_anteriormente)
				s_cor = "darkorange"
				end if
            if .qtde_devolucao_pendente<>0 then
                s_devolucao_pendente=CStr(.qtde_devolucao_pendente)
                end if
			if ((.qtde - .qtde_devolvida_anteriormente) > 0) And ((.qtde - .qtde_devolucao_pendente) > 0) then s_readonly = ""
			end with
		
	 else
		s_fabricante=""
		s_produto=""
		s_descricao=""
		s_descricao_html=""
		s_qtde=""
		s_devolucao_anterior=""
        s_devolucao_pendente=""
		s_vl_unitario=""
		s_vl_TotalItem=""
		s_readonly = "readonly tabindex=-1"
		end if
%>
	<tr>
	<td class="MDBE"><input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:25px; color:<%=s_cor%>"
		value='<%=s_fabricante%>' readonly tabindex=-1></td>
	<td class="MDB"><input name="c_produto" id="c_produto" class="PLLe" style="width:54px; color:<%=s_cor%>"
		value='<%=s_produto%>' readonly tabindex=-1></td>
	<td class="MDB" style="width:269px;">
		<span class="PLLe" style="color:<%=s_cor%>"><%=s_descricao_html%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value='<%=s_descricao%>'>
	</td>
	<td class="MDB" align="right"><input name="c_qtde" id="c_qtde" class="PLLd" style="width:38px; color:<%=s_cor%>"
		value='<%=s_qtde%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_devolucao_anterior" id="c_devolucao_anterior" class="PLLd" style="width:40px; color:<%=s_cor%>"
		value='<%=s_devolucao_anterior%>' readonly tabindex=-1></td>
    <td class="MDB" align="right"><input name="c_devolucao_pendente" id="c_devolucao_pendente" class="PLLd" style="width:40px; color:blue"
		value='<%=s_devolucao_pendente%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_qtde_devolucao" id="c_qtde_devolucao" class="PLLd" maxlength="4" style="width:40px;color:red" onkeypress="filtra_numerico();" onblur="if (verifica_qtde_max(this,<%=Cstr(i)%>)) recalcula_valores();"
		value='' <%=s_readonly%>></td>
	
	<td class="MDB" align="right"><input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_vl_unitario%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px; color:<%=s_cor%>" 
		value='<%=s_vl_TotalItem%>' readonly tabindex=-1></td>
	</tr>
<% next %>
	<tr>
	<td colspan="8" class="MD">&nbsp;</td>
	<td class="MDB" align="right"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:red;" 
		value='<%=formata_moeda(m_TotalDestePedido)%>' readonly tabindex=-1></td>
	</tr>
</table>
<br />
<table class="Q" style="width: 649px;" cellpadding="0" cellspacing="0">
    <tr>
        <td colspan="4" class="MB" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">MOTIVO DA DEVOLUÇÃO</span></td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;"><p class="Cd">Motivo</p></td>
        <td style="padding: 5px;" colspan="3">
            <select id="c_motivo_devolucao" name="c_motivo_devolucao" style='width:380px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;'>
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__MOTIVO, "")%>
            </select>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" valign="top" class="MC"><p class="Cd">Descrição / Observações</p></td>
        <td style="padding: 5px;" colspan="3" class="MC">
            <textarea name="c_motivo_descricao" id="c_motivo_descricao" class="PLLe" rows="<%=MAX_LINHAS_DESCRICAO_OBSERVACAO_DEVOLUCAO%>" style="width: 520px;"></textarea>
        </td>
    </tr>
    <tr>
        <td colspan="4" class="MB MC" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">TAXA ADMINISTRATIVA</span></td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;"><p class="Cd">Taxa</p></td>
        <td style="padding: 5px;width:180px;">
            <select id="c_taxa" name="c_taxa" style='width:80px;' onchange="recalcula_valor_taxa();">
                <option value="<%=TAXA_ADMINISTRATIVA__SIM%>">Sim</option>
                <option value="<%=TAXA_ADMINISTRATIVA__NAO%>" selected>Não</option>
            </select>
        </td>
        <td style="padding: 5px;width:140px"><p class="Cd">Forma de pagto</p></td>
        <td style="padding: 5px;">
            <select id="c_taxa_forma_pagto" name="c_taxa_forma_pagto" style='width:180px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;'>
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__TAXA_FORMA_PAGAMENTO, "")%>
            </select>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Percentual</p></td>
        <td style="padding:5px;width:100px;" class="MC">
            <select name="c_taxa_percentual" id="c_taxa_percentual" style="width:60px;" onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;' onchange="recalcula_valor_taxa();">
                <option value="">&nbsp;</option>               
                <option value="10">10%</option>               
                <option value="15">15%</option>               
                <option value="25">25%</option>
                <option value="30">30%</option>               
            </select>
        </td>
        <td style="padding: 5px;" class="MC"><p class="Cd">Cobrar valor:</p></td>
        <td align="left" style="padding:5px;" class="MC">
            <input type="text" readonly tabindex=-1 name="c_taxa_valor" id="c_taxa_valor" class="PLLd" style="width:80px; margin-left: 10px;">
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Responsável</p></td>
        <td style="padding: 5px;" colspan="3" class="MC">
            <select id="c_taxa_responsavel" name="c_taxa_responsavel" style='width:220px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;'>
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__TAXA_RESPONSAVEL, "")%>
            </select>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" valign="top" class="MC"><p class="Cd">Observações da taxa</p></td>
        <td style="padding: 5px;" colspan="3" class="MC">
            <textarea name="c_taxa_observacoes" id="c_taxa_observacoes" class="PLLe" rows="<%=MAX_LINHAS_DESCRICAO_OBSERVACAO_DEVOLUCAO%>" style="width: 520px;"></textarea>
        </td>
    </tr>
    <tr>
        <td colspan="4" class="MC" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">CRÉDITO</span></td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Transação</p></td>
        <td style="padding: 5px;" class="MC">
            <select id="c_credito_transacao" name="c_credito_transacao" style='width:180px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;'>
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__CREDITO_TRANSACAO, "")%>
            </select>
        </td>
        <td style="padding: 5px;" class="MC"><p class="Cd">Pedido novo:</p></td>
        <td align="left" style="padding:5px;" class="MC">
            <input type="text" name="c_pedido_novo" id="c_pedido_novo" maxlength="10" class="PLLe" style="width:80px; margin-left: 10px;" onkeypress="filtra_pedido();" onblur="if (normaliza_numero_pedido_e_sufixo(this.value)!='') {this.value=normaliza_numero_pedido_e_sufixo(this.value);}"">
        </td>
    </tr>
    <tr id="trDadosBancarios" style="display: none;">
        <td colspan="4">
            <table style="width: 649px;" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="padding: 5px;" class="MC"><p class="Cd">Nº Banco:</p></td>
                    <td align="left" style="padding:5px;" class="MC">
                        <input type="text" maxlength="4" name="c_cliente_banco" id="c_cliente_banco" class="PLLe" style="width:60px; margin-left: 10px;">
                    </td>
                    <td style="padding: 5px;" class="MC"><p class="Cd">Agência:</p></td>
                    <td align="left" style="padding:5px;" class="MC">
                        <input type="text" maxlength="10" name="c_cliente_agencia" id="c_cliente_agencia" class="PLLe" style="width:80px; margin-left: 10px;">
                    </td>
                    <td style="padding: 5px;" class="MC"><p class="Cd">Conta:</p></td>
                    <td align="left" style="padding:5px;" class="MC">
                        <input type="text" maxlength="14" name="c_cliente_conta" id="c_cliente_conta" class="PLLe" style="width:110px; margin-left: 10px;">
                    </td>
                </tr>
                <tr>
                    <td style="padding: 5px;" class="MC"><p class="Cd">Favorecido:</p></td>
                    <td align="left" style="padding:5px;" class="MC" colspan="5">
                        <input type="text" maxlength="40" name="c_cliente_favorecido" id="c_cliente_favorecido" class="PLLe" style="width:340px; margin-left: 10px;">
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" valign="top" class="MC"><p class="Cd">Observações do crédito</p></td>
        <td style="padding: 5px;" colspan="3" class="MC">
            <textarea name="c_credito_observacoes" id="c_credito_observacoes" class="PLLe" rows="<%=MAX_LINHAS_DESCRICAO_OBSERVACAO_DEVOLUCAO%>" style="width: 520px;"></textarea>
        </td>
    </tr>
    <tr>
        <td colspan="4" class="MC" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">FOTOS</span></td>
    </tr>
    <tr>
        <td colspan="4" style="padding: 5px;width:100px;" valign="top" class="MC"><p class="C">Selecione até 6 (seis) arquivos de foto:</p></td>
    </tr>
    <tr>
        <td colspan="4">
            <table cellspacing="0" cellpadding="3" width="100%">
            <tr>
                <td>
                    <input type="file" name="arquivo1" id="arquivo1" onchange='reloadFileReader("1")' class="PLLd" style="font-weight: normal; width: 200px" />
                </td>
                <td>
                    <input type="file" name="arquivo2" id="arquivo2" onchange='reloadFileReader("2")' class="PLLd" style="font-weight: normal; width: 200px" />
                </td>
                <td>
                    <input type="file" name="arquivo3" id="arquivo3" onchange='reloadFileReader("3")' class="PLLd" style="font-weight: normal; width: 200px" />
                </td>
            </tr>
            <tr>
                <td>
                    <div id="image-holder-arquivo1" style="width: 200px; height: 180px; border: 1px dashed #000" onclick="fPED.arquivo1.click();"></div>
                </td>
                <td>
                    <div id="image-holder-arquivo2" style="width: 200px; height: 180px; border: 1px dashed #000" onclick="fPED.arquivo2.click();"></div>
                </td>
                <td>
                    <div id="image-holder-arquivo3" style="width: 200px; height: 180px; border: 1px dashed #000" onclick="fPED.arquivo3.click();"></div>
                </td>
            </tr>
            <tr>
                <td>
                    <input type="file" name="arquivo4" id="arquivo4" onchange='reloadFileReader("4")' class="PLLd" style="font-weight: normal; width: 200px" />
                </td>
                <td>
                    <input type="file" name="arquivo5" id="arquivo5" onchange='reloadFileReader("5")' class="PLLd" style="font-weight: normal; width: 200px" />
                </td>
                <td>
                    <input type="file" name="arquivo6" id="arquivo6" onchange='reloadFileReader("6")' class="PLLd" style="font-weight: normal; width: 200px" />
                </td>
            </tr>
            <tr>
                <td>
                    <div id="image-holder-arquivo4" style="width: 200px; height: 180px; border: 1px dashed #000" onclick="fPED.arquivo4.click();"></div>
                </td>
                <td>
                    <div id="image-holder-arquivo5" style="width: 200px; height: 180px; border: 1px dashed #000" onclick="fPED.arquivo5.click();"></div>
                </td>
                <td>
                    <div id="image-holder-arquivo6" style="width: 200px; height: 180px; border: 1px dashed #000" onclick="fPED.arquivo6.click();"></div>
                </td>
            </tr>
        </table>
        </td>
    </tr>
    <tr>
        <td colspan="4" class="MC" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">OBSERVAÇÕES</span></td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" valign="top" class="MC"><p class="Cd">Observações gerais</p></td>
        <td style="padding: 5px;" colspan="3" class="MC">
            <textarea name="c_observacoes" id="c_observacoes" class="PLLe" rows="<%=MAX_LINHAS_DESCRICAO_OBSERVACAO_DEVOLUCAO%>" style="width: 520px;"></textarea>
        </td>
    </tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="852" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<!-- ************   BOTÕES   ************ -->
<table width="649" cellSpacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDConfirma(fPED)" title="confirma a devolução">
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
