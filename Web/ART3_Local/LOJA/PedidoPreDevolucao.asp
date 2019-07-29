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
'     =====================================================
'	  P E D I D O P R E D E V O L U C A O E D I T A . A S P
'     =====================================================
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

    Const LOCAL_COLETA__TRANSPORTADORA = "003"

	dim s, usuario, pedido_selecionado, id_devolucao, st_codigo, st_devolucao_descricao, st_devolucao_cor, st_usuario, st_data_hora, loja
    loja = Trim(Session("loja_atual"))
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    dim url_back, rb_status, c_loja
    url_back = Request.Form("url_back")
    rb_status = Request.Form("rb_status")
    c_loja = Request.Form("c_loja")

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_PRE_DEVOLUCAO_LEITURA, s_lista_operacoes_permitidas) And _
        Not operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

    id_devolucao = Request("id_devolucao")
    if id_devolucao = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
	dim i, n, cont
	dim s_fabricante, s_produto, s_descricao, s_descricao_html, s_qtde, s_vl_unitario, s_vl_TotalItem, m_TotalItem, m_TotalDestePedido
	dim s_cor, s_devolucao_anterior, s_devolucao_pendente, s_readonly, s_display, s_disabled, blnExibeBotaoConfirma, s_disabled_taxa_credito, s_readonly_taxa_credito
    dim devolucao_status, devolucao_usuario
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
    If Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

    dim s_sessionToken
	s_sessionToken = ""
	s = "SELECT Convert(varchar(36), SessionTokenModuloLoja) AS SessionTokenModuloLoja FROM t_USUARIO WHERE (usuario = '" & usuario & "')"
    if r.State <> 0 then r.Close
    r.Open s,cn
	if Not r.Eof then s_sessionToken = Trim("" & r("SessionTokenModuloLoja"))

	dim r_pedido, alerta, msg_erro
	alerta=""
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
		end if

    dim r_indicador, ha_indicador
    ha_indicador = False
    if alerta = "" then
        ha_indicador = le_indicador(r_pedido.indicador, r_indicador, msg_erro)
        end if

    dim serverVariablesUrl, serverVariablesServerName
    serverVariablesServerName = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
    serverVariablesUrl = Request.ServerVariables("URL")
    serverVariablesUrl = Ucase(serverVariablesUrl)
    serverVariablesUrl = Mid(serverVariablesUrl, 1, CInt(InStr(serverVariablesUrl, "LOJA")-1))
    serverVariablesUrl = serverVariablesServerName & serverVariablesUrl

    if alerta = "" then
        s = "SELECT * FROM t_PEDIDO_DEVOLUCAO " & _
                "INNER JOIN t_PEDIDO on t_PEDIDO_DEVOLUCAO.pedido = t_PEDIDO.pedido WHERE (id='" & id_devolucao & "')"
        set rs = cn.Execute(s)
	    if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)

        if rs.Eof then
            alerta = "Devolução não localizada (ID = " & id_devolucao & ")."
            end if
        end if

    devolucao_usuario = rs("usuario_cadastro")
    st_codigo = CStr(rs("status"))
    st_usuario = Trim("" & rs("status_usuario"))
    st_data_hora = formata_data_hora_sem_seg(rs("status_data_hora"))
    obtem_descricao_status_devolucao st_codigo, st_devolucao_descricao, st_devolucao_cor

    blnExibeBotaoConfirma = False
    s_readonly = "readonly tabindex=-1"
    s_readonly_taxa_credito = "readonly tabindex=-1"
    s_disabled = "disabled"
    s_disabled_taxa_credito = "disabled"
    if operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) And loja = rs("loja") then
        if st_codigo <> COD_ST_PEDIDO_DEVOLUCAO__FINALIZADA And st_codigo <> COD_ST_PEDIDO_DEVOLUCAO__CANCELADA And _ 
                st_codigo <> COD_ST_PEDIDO_DEVOLUCAO__REPROVADA then
            s_readonly = ""
            s_disabled = ""
            blnExibeBotaoConfirma = True
            end if
        if st_codigo = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA then
            s_disabled_taxa_credito = ""
            s_readonly_taxa_credito = ""
            end if
        end if

' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ___________________________________
' OBTÉM DESCRIÇÃO STATUS DEVOLUÇÃO
'
function obtem_descricao_status_devolucao(byval st_codigo, byref st_devolucao_descricao, byref st_devolucao_cor)
dim s_descricao, s_cor
    st_codigo = Trim("" & st_codigo)
    if st_codigo = "" then exit function

    st_devolucao_descricao = ""
    st_devolucao_cor = ""

    s_descricao = ""
    s_cor = ""
    select case st_codigo
        case COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA
                s_descricao = "Cadastrada"
                s_cor = "#0348E1"
            case COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO
                s_descricao = "Em Andamento"
                s_cor = "#E26534"
            case COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA
                s_descricao = "Mercadoria Recebida"
                s_cor = "#008080"
            case COD_ST_PEDIDO_DEVOLUCAO__FINALIZADA
                s_descricao = "Finalizada"
                s_cor = "#4FAB5B"
            case COD_ST_PEDIDO_DEVOLUCAO__REPROVADA
                s_descricao = "Reprovada"
                s_cor = "#B91832"
            case COD_ST_PEDIDO_DEVOLUCAO__CANCELADA
                s_descricao = "Cancelada"
                s_cor = "#C7465A"
            case else
                s_descricao = "Indefinido"
                s_cor = "#000000"
    end select
    st_devolucao_descricao = s_descricao
    st_devolucao_cor = s_cor
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
	<title>LOJA<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>



<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
var LOCAL_COLETA__CLIENTE = "001";
var LOCAL_COLETA__PARCEIRO = "002";
var LOCAL_COLETA__TRANSPORTADORA = "003";
var LOCAL_COLETA__OUTROS = "004";
var CREDITO_TRANSACAO__ESTORNO = "002";
var CREDITO_TRANSACAO__REEMBOLSO = "003";
var TAXA_ADMINISTRATIVA__NAO = "0";
var TAXA_ADMINISTRATIVA__SIM = "1";
var TAXA_FORMA_PAGAMENTO__ABATIMENTO_CREDITO = "003";
var TAXA_RESPONSAVEL__CLIENTE = "003";
var COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA = "1";
var COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO = "2";
var COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA = "3";
var COD_ST_PEDIDO_DEVOLUCAO__FINALIZADA = "4";
var COD_ST_PEDIDO_DEVOLUCAO__REPROVADA = "5";
var COD_ST_PEDIDO_DEVOLUCAO__CANCELADA = "6";
var COD_ST_PEDIDO_DEVOLUCAO__INDEFINIDO = "0";
var blnFotoPendente, blnAcionaBtnConfirma;
    
    $(function () {
        
        blnFotoPendente = false;
        blnAcionaBtnConfirma = false;

            $('#c_taxa_valor').val(formata_moeda(0));
            $('#c_taxa_forma_pagto').attr('disabled', true);
            $('#c_taxa_percentual').attr('readonly', true);
            $('#c_taxa_responsavel').attr('disabled', true);
            $('#c_taxa_valor').attr('disabled', true);


        $('#c_taxa').change(function () {
            if ($(this).val() == TAXA_ADMINISTRATIVA__SIM) {
                $('#c_taxa_forma_pagto').attr('disabled', false);
                $('#c_taxa_percentual').attr('readonly', false);
                $('#c_taxa_responsavel').attr('disabled', false);
                $('#c_taxa_valor').attr('disabled', false);
            }
            else {
                $("#c_taxa_forma_pagto").prop('selectedIndex', -1);
                $("#c_taxa_percentual").val('0');
                $("#c_taxa_responsavel").prop('selectedIndex', -1);
                $('#c_taxa_forma_pagto').attr('disabled', true);
                $('#c_taxa_percentual').attr('readonly', true);
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
        // estão habilitados em caso de taxa = sim        
        if ($('#c_taxa').val() == TAXA_ADMINISTRATIVA__SIM) {
            recalcula_valor_taxa();
            <% if operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) And loja = rs("loja") then %>
            if ($('#cod_status_devolucao').val() == COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA) {
                $('#c_taxa_forma_pagto').attr('disabled', false);
                $('#c_taxa_percentual').attr('readonly', false);
                $('#c_taxa_responsavel').attr('disabled', false);
                $('#c_taxa_valor').attr('disabled', false);
            }
            <% end if %>
        }
        
        var local_coleta = $('#c_local_coleta :selected').val();
        if (local_coleta == LOCAL_COLETA__TRANSPORTADORA){
            $('#c_transportadora').text("<%=r_pedido.transportadora_id%> - <%=x_transportadora(r_pedido.transportadora_id)%>");
        }

        window.onbeforeunload = function(event) {
            if (blnFotoPendente) {
                if (!blnAcionaBtnConfirma) {
                    if(!confirm("É necessário clicar em 'Confirmar' para salvar as novas fotos adicionadas!!\n\nTem certeza de que deseja sair da página sem salvar as novas fotos?")){
                        event.preventDefault();
                        return;
                    }
                }
            }
        };
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

function fPEDDevolucaoMensagemCadastra(f, id_devolucao) {
    f.action = "PedidoPreDevolucaoMensagemNova.asp";
    f.id_devolucao.value = id_devolucao;
    window.status = "Aguarde ...";
    f.submit();
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

function fPEDConfirma( f ) {
var b, i, n, serverVariableUrl;
var s_stored_file_guid, blnUploadOk, blnEnviaRequisicaoUpload, inputFile; 
serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
serverVariableUrl = serverVariableUrl.toUpperCase();
serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("LOJA"));

    blnAcionaBtnConfirma = true;

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
	    else {
	        if ((converte_numero(f.c_taxa_percentual.value) > 100) || (converte_numero(f.c_taxa_percentual.value) < 0)) {
	            alert("O percentual da taxa administrativa não pode ser menor que 0 (zero) e nem maior que 100 (cem)!!");
	            f.c_taxa_percentual.focus();
	            return;
	        }
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
	    if (f.c_cliente_banco.value==""){
	        alert("Preencha o código do banco!!");
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

	if (f.cod_status_devolucao.value == COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA) {
	    if (f.rb_status_reprova.checked){
	        if (!confirm("Tem certeza de que deseja REPROVAR o pedido de devolução?")) return;
	    }
	    if (f.rb_status_aprova.checked){
	        if (!confirm("Tem certeza de que deseja APROVAR o pedido de devolução?")) return;
	    }
	}

	blnEnviaRequisicaoUpload = false;
	window.status = "Aguarde ...";
	dCONFIRMA.style.visibility="hidden";

	n = $('input[type=file]').length;
	for (i = 0; i< n;i++) {
	    inputFile = $('input[type=file]')[i];
	    if (inputFile.value != "") {
	        blnEnviaRequisicaoUpload = true;
	        break;
	    }
	}

	if (blnEnviaRequisicaoUpload) {
	    var form = $('#fPED');
	    var fd = new FormData(form[0]);
	    blnUploadOk = false;
	    s_stored_file_guid = "";

	    $.ajax({
	        type: form.attr('method'),
	        url: 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/UploadFile/PostFile',
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
	    f.action="pedidopredevolucaoatualiza.asp";
	    f.submit();
	}
}
function filtra_digito_percentual_taxa() {
    var letra;
    letra = String.fromCharCode(window.event.keyCode);
    if (((letra < "0") || (letra > "9")) && (letra != "-") && (letra != ".") && (letra != ",")) window.event.keyCode = 0;
}
function filtra_percentual_taxa(valor) {
    if ((converte_numero(valor) > 100) || (converte_numero(valor) < 0)) {
        alert("O percentual da taxa administrativa não pode ser menor que 0 (zero) e nem maior que 100 (cem)!!");
        fPED.c_taxa_percentual.focus();
    }
}

function reloadFileReader(idx) {
var file_image_extensions = "|jpeg|jpg|png|bmp|gif|tif|tiff|";
var img_src, img_width;
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
                }
                else {
                    img_src = "../IMAGEM/file_150x150.png";
                    img_width = "150";
                }
            }
            else {
                img_src = e.target.result;
                img_width = "180";
            }


            $("<img />", {
                "src": img_src,
                "class": "thumb-image",
                "width": img_width,
                "height": "140"
            }).appendTo(image_holder);
        }
        image_holder.show();
        reader.readAsDataURL($('#arquivo' + idx)[0].files[0]);
        blnFotoPendente = true;
    }
}

function fPEDCancelar( f ) {
    if (!confirm("Tem certeza de que deseja cancelar a pré-devolução?")) return;
    f.action = "pedidopredevolucaocancela.asp";
    dREMOVER.style.visibility = "hidden";
    window.status = "Aguarde ...";
    f.submit();
}

function fPEDFinalizar( f ) {
    if (!confirm("Tem certeza de que deseja finalizar o processo de devolução?")) return;
    f.action = "pedidopredevolucaofinaliza.asp";
    dFINALIZAR.style.visibility = "hidden";
    window.status = "Aguarde ...";
    f.submit();
}

function fPEDRemoverFoto(id_upload_file_param) {
var id_devolucao_param
id_devolucao_param = $('#id_devolucao').val();
    if(!confirm("Tem certeza de que deseja remover essa foto?")) return;

    $.ajax({
        type: "post",
        url: "../Global/AjaxPedidoDevolucaoRemoveFoto.asp?id_devolucao=" + id_devolucao_param + "&id_upload_file=" + id_upload_file_param + "&usuario=<%=usuario%>",
        contentType: false,
        processData: false,
        success: function (resp) {
            $('#tblFotos').empty();
            $('#tblFotos').html(resp);
        },
        error: function (resp) {
            alert("Erro ao tentar remover a foto!" + resp);
        }
    });
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
select[disabled]{
    background-color: white;
    border: 0px;
    font-size: 8pt;
}
select[disabled]::-ms-expand{
    display:none;
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
<body>
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="c_pedido_cliente" id="c_pedido_cliente" value="<%=r_pedido.id_cliente%>" />
<input type="hidden" name="id_devolucao" id="id_devolucao" value="<%=id_devolucao%>" />
<input type="hidden" name="c_devolucao_usuario" id="c_devolucao_usuario" value="<%=Trim("" & rs("usuario_cadastro"))%>" />
<input type="hidden" name="cod_status_devolucao" id="cod_status_devolucao" value="<%=st_codigo%>" />
<input type="hidden" name="c_pedido_possui_parcela_cartao" id="c_pedido_possui_parcela_cartao" value="<%=r_pedido.st_forma_pagto_possui_parcela_cartao%>" />
<input type="hidden" name="c_pedido_indicador" id="c_pedido_indicador" value="<%=r_pedido.indicador%>" />
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=r_pedido.vendedor%>" />
<input type="hidden" name="url_origem" id="url_origem" value="PedidoPreDevolucao.asp" />
<input type="hidden" name="url_back" id="url_back" value="<%=url_back%>" />
<input type="hidden" name="rb_status_back" id="rb_status_back" value="<%=rb_status%>" />
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>" />
<input type="hidden" name="upload_file_guid_returned" id="upload_file_guid_returned" />
<input type="hidden" name="upload_parameter__is_temp_file" value="0" />
<input type="hidden" name="upload_parameter__is_confirmation_required" value="1" />
<input type="hidden" name="upload_parameter__folder_name" value="PEDIDO_DEVOLUCAO_FOTOS" />
<input type="hidden" name="upload_parameter__user_id" value="<%=usuario%>" />
<input type="hidden" name="upload_parameter__sessionToken" value="<%=s_sessionToken%>" />
<input type="hidden" name="st_devolucao" value="<%=st_codigo%>" />

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<%=MontaHeaderIdentificacaoPedido(pedido_selecionado, r_pedido, 852)%>
<br>

<!--  DESCRIÇÃO DA OPERAÇÃO -->
<table width="649" cellPadding="0" CellSpacing="0">
<tr><td><p class="Expl">CONSULTAR/EDITAR PRÉ-DEVOLUÇÃO - (ID: <%=rs("id")%>)</p></td></tr>
</table>
<br />

<table class="Q" style="width: 649px;" cellpadding="0" cellspacing="0">
    <tr>
        <td colspan="4" class="MB"><span class="C" style="color:<%=st_devolucao_cor%>;">Status: <%=UCase(st_devolucao_descricao)%> por <%=st_usuario%> em <%=formata_data_hora_sem_seg(st_data_hora)%></span></td>
    </tr>
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
            <select id="c_procedimento" name="c_procedimento" style='width:240px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;' <%=s_disabled%>>
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__PROCEDIMENTO, Trim("" & rs("cod_procedimento")))%>
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
                onchange="preencheEndColeta(fPED, this.value);" <%=s_disabled%>>
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__LOCAL_COLETA, Trim("" & rs("cod_local_coleta")))%>
            </select>&nbsp;&nbsp;&nbsp;
            <span id="c_transportadora" class="C"></span>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Endereço:</p></td>
        <td align="left" style="padding:5px;width:260px;" class="MC">
            <input type="text" maxlength="80" name="c_coleta_endereco" id="c_coleta_endereco" class="PLLe" style="width:240px; margin-left: 10px;"
                value="<%=Trim("" & rs("endereco_coleta_logradouro"))%>" <%=s_readonly%> />
        </td>
        <td style="padding: 5px;width:100px;" class="MC ME"><p class="Cd">Nº:</p></td>
        <td align="left" style="padding:5px;width:189px;" class="MC">
            <input type="text" maxlength="20" name="c_coleta_endereco_numero" id="c_coleta_endereco_numero" class="PLLe" style="width:100px; margin-left: 10px;"
                value="<%=Trim("" & rs("endereco_coleta_numero"))%>" <%=s_readonly%>>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Bairro:</p></td>
        <td align="left" style="padding:5px;width:260px;" class="MC">
            <input type="text" maxlength="72" name="c_coleta_bairro" id="c_coleta_bairro" class="PLLe" style="width:240px; margin-left: 10px;"
                value="<%=Trim("" & rs("endereco_coleta_bairro"))%>" <%=s_readonly%>>
        </td>
        <td style="padding: 5px;width:100px;" class="MC ME"><p class="Cd">CEP:</p></td>
        <td align="left" style="padding:5px;width:189px;" class="MC">
            <input type="text" maxlength="9" name="c_coleta_cep" id="c_coleta_cep" class="PLLe" style="width:100px; margin-left: 10px;" onblur="if (!cep_ok(this.value)){ alert('CEP inválido!!'); this.focus(); };"
                value="<%=Trim("" & rs("endereco_coleta_cep"))%>" <%=s_readonly%>>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Cidade:</p></td>
        <td align="left" style="padding:5px;width:260px;" class="MC">
            <input type="text" maxlength="60" name="c_coleta_cidade" id="c_coleta_cidade" class="PLLe" style="width:240px; margin-left: 10px;"
                value="<%=Trim("" & rs("endereco_coleta_cidade"))%>" <%=s_readonly%>>
        </td>
        <td style="padding: 5px;width:100px;" class="MC ME"><p class="Cd">UF:</p></td>
        <td align="left" style="padding:5px;width:189px;" class="MC">
            <select name="c_coleta_uf" id="c_coleta_uf" class="PLLe" style="width:50px; margin-left: 10px;" <%=s_disabled%>>
                <%=UF_monta_itens_select(Trim("" & rs("endereco_coleta_uf")))%>                
            </select>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Complemento:</p></td>
        <td align="left" style="padding:5px;width:260px;" class="MC" colspan="3">
            <input type="text" maxlength="60" name="c_coleta_complemento" id="c_coleta_complemento" class="PLLe" style="width:380px; margin-left: 10px;"
                value="<%=Trim("" & rs("endereco_coleta_complemento"))%>" <%=s_readonly%>>
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
	<td class="MB" valign="bottom"><p class="PLTd" title="quantidade a ser devolvida">Qtde<br />Devol</p></td>
	<td class="MB" valign="bottom"><p class="PLTd" title="quantidade para o estoque de venda">Estoque<br />Venda</p></td>
	<td class="MB" valign="bottom"><p class="PLTd" title="quantidade para o estoque de danificados">Estoque<br />Danif</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Valor<br>Unitário</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Total<br>Devolução</p></td>
	</tr>

<% 
    s = _
	    "SELECT " & _
		    "tPDI.fabricante," & _
		    "tPDI.produto," & _
		    "tPDI.qtde," & _
		    "tPDI.qtde_estoque_venda," & _
		    "tPDI.qtde_estoque_danificado," & _
            "tPDI.vl_unitario," & _
		    "tP.descricao," & _
            "tP.descricao_html" & _
	    " FROM t_PEDIDO_DEVOLUCAO_ITEM tPDI" & _
        " INNER JOIN t_PRODUTO tP ON ((tPDI.fabricante=tP.fabricante) AND (tPDI.produto=tP.produto))" & _
	    " WHERE" & _
		    " (id_pedido_devolucao = " & Trim("" & rs("id")) & ")" & _
	    " ORDER BY" & _
		    " tPDI.produto," & _
		    " tPDI.fabricante"
    if r.State <> 0 then r.Close
    r.open s, cn
    m_TotalDestePedido=0
    m_TotalItem=0
    do while Not r.Eof

        s_cor = "black"
        s_fabricante=Trim("" & r("fabricante"))
        s_produto=Trim("" & r("produto"))
        s_descricao=Trim("" & r("descricao"))
        s_descricao_html=produto_formata_descricao_em_html(Trim("" & r("descricao_html")))
        s_qtde=converte_numero(r("qtde"))
        s_vl_unitario=formata_moeda(Trim("" & r("vl_unitario")))
        m_TotalItem=converte_numero(s_vl_unitario)*s_qtde
        s_vl_TotalItem=formata_moeda(m_TotalItem)
        m_TotalDestePedido=m_TotalDestePedido + m_TotalItem		
%>
	<tr>
	<td class="MDBE"><input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:25px; color:<%=s_cor%>"
		value='<%=s_fabricante%>' readonly tabindex=-1></td>
	<td class="MDB"><input name="c_produto" id="c_produto" class="PLLe" style="width:54px; color:<%=s_cor%>"
		value='<%=s_produto%>' readonly tabindex=-1></td>
	<td class="MDB" style="width:290px;">
		<span class="PLLe" style="color:<%=s_cor%>"><%=s_descricao_html%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value='<%=s_descricao%>'>
	</td>
	<td class="MDB" align="right"><input name="c_qtde_devolucao" id="c_qtde_devolucao" class="PLLd" style="width:38px; color:<%=s_cor%>"
		value='<%=s_qtde%>' readonly tabindex=-1></td>
    <td class="MDB" align="right"><input name="c_qtde_estoque_venda" id="c_qtde_estoque_venda" class="PLLd" style="width:38px; color:<%=s_cor%>"
		value='<%=Trim("" & r("qtde_estoque_venda"))%>' readonly tabindex=-1></td>
    <td class="MDB" align="right"><input name="c_qtde_estoque_danificado" id="c_qtde_estoque_danificado" class="PLLd" style="width:38px; color:<%=s_cor%>"
		value='<%=Trim("" & r("qtde_estoque_danificado"))%>' readonly tabindex=-1></td>
	
	<td class="MDB" align="right"><input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_vl_unitario%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px; color:<%=s_cor%>" 
		value='<%=s_vl_TotalItem%>' readonly tabindex=-1></td>
	</tr>
<% r.MoveNext
    loop %>
	<tr>
	<td colspan="7" class="MD">&nbsp;</td>
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
            <select id="c_motivo_devolucao" name="c_motivo_devolucao" style='width:380px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;' <%=s_disabled%>>
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__MOTIVO, Trim("" & rs("cod_devolucao_motivo")))%>
            </select>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" valign="top" class="MC"><p class="Cd">Descrição / Observações</p></td>
        <td style="padding: 5px;" colspan="3" class="MC">
            <textarea name="c_motivo_descricao" id="c_motivo_descricao" class="PLLe" rows="<%=MAX_LINHAS_DESCRICAO_OBSERVACAO_DEVOLUCAO%>" style="width: 520px;" <%=s_readonly%>><%=Trim("" & rs("motivo_observacao"))%></textarea>
        </td>
    </tr>
    <tr>
        <td colspan="4" class="MB MC" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">TAXA ADMINISTRATIVA</span></td>
    </tr>
	<tr>
		<% s = Trim("" & rs("taxa_usuario_ult_atualizacao"))
			if s <> "" then s = formata_data_hora_sem_seg(rs("taxa_dt_hr_ult_atualizacao")) & " por " & s
			s = "Última edição: " & s
			%>
		<td style="padding: 5px;width:100px;">&nbsp;</td>
		<td colspan="3"><span class="C" style="font-style:italic;color:gray;"><%=s%></span></td>
	</tr>
    <tr>
        <td class="MC" style="padding: 5px;width:100px;"><p class="Cd">Taxa</p></td>
        <td class="MC" style="padding: 5px;width:180px;">
            <select id="c_taxa" name="c_taxa" style='width:80px;' onchange="recalcula_valor_taxa();" <%=s_disabled_taxa_credito%>>
                <option value="<%=TAXA_ADMINISTRATIVA__SIM%>" <%=iif( (CInt(rs("taxa_flag"))=CInt(TAXA_ADMINISTRATIVA__SIM)), "selected", "")%>>Sim</option>
                <option value="<%=TAXA_ADMINISTRATIVA__NAO%>" <%=iif( (CInt(rs("taxa_flag"))=CInt(TAXA_ADMINISTRATIVA__NAO)), "selected", "")%>>Não</option>
            </select>
        </td>
        <td class="MC" style="padding: 5px;width:140px"><p class="Cd">Forma de pagto</p></td>
        <td class="MC" style="padding: 5px;">
            <select id="c_taxa_forma_pagto" name="c_taxa_forma_pagto" style='width:180px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;' <%=s_disabled_taxa_credito%>>
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__TAXA_FORMA_PAGAMENTO, Trim("" & rs("cod_taxa_forma_pagto")))%>
            </select>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Percentual</p></td>
        <td style="padding:5px;width:100px;" class="MC">
            <input type="number" name="c_taxa_percentual" maxlength="5" id="c_taxa_percentual" class="PLLd" style="width:60px;" onchange="recalcula_valor_taxa();" <%=s_readonly_taxa_credito%> value="<%=Trim( "" & rs("taxa_percentual"))%>"
                onblur="this.value=formata_numero(this.value,1); filtra_percentual_taxa(this.value);" onkeypress="filtra_digito_percentual_taxa();" /><span class="C">&nbsp;%</span>
        </td>
        <td style="padding: 5px;" class="MC"><p class="Cd">Cobrar valor:</p></td>
        <td align="left" style="padding:5px;" class="MC">
            <input type="text" readonly tabindex=-1 name="c_taxa_valor" id="c_taxa_valor" class="PLLd" style="width:80px; margin-left: 10px;">
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Responsável</p></td>
        <td style="padding: 5px;" colspan="3" class="MC">
            <select id="c_taxa_responsavel" name="c_taxa_responsavel" style='width:220px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;' <%=s_disabled_taxa_credito%>>
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__TAXA_RESPONSAVEL, Trim("" & rs("cod_taxa_responsavel")))%>
            </select>
        </td>
    </tr>
    <tr>
        <td style="padding: 5px;width:100px;" valign="top" class="MC"><p class="Cd">Observações da taxa</p></td>
        <td style="padding: 5px;" colspan="3" class="MC">
            <textarea name="c_taxa_observacoes" id="c_taxa_observacoes" class="PLLe" rows="<%=MAX_LINHAS_DESCRICAO_OBSERVACAO_DEVOLUCAO%>" style="width: 520px;" <%=s_readonly_taxa_credito%>><%=Trim("" & rs("taxa_observacoes"))%></textarea>
        </td>
    </tr>

    <tr>
        <td colspan="4" class="MC MB" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">CRÉDITO</span></td>
    </tr>
	<tr>
		<% s = Trim("" & rs("credito_usuario_ult_atualizacao"))
			if s <> "" then s = formata_data_hora_sem_seg(rs("credito_dt_hr_ult_atualizacao")) & " por " & s
			s = "Última edição: " & s
			%>
		<td style="padding: 5px;width:100px;">&nbsp;</td>
		<td colspan="3"><span class="C" style="font-style:italic;color:gray;"><%=s%></span></td>
	</tr>
    <tr>
        <td style="padding: 5px;width:100px;" class="MC"><p class="Cd">Transação</p></td>
        <td style="padding: 5px;" class="MC">
            <select id="c_credito_transacao" name="c_credito_transacao" style='width:180px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;' <%=s_disabled_taxa_credito%>>
                <%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__CREDITO_TRANSACAO, Trim("" & rs("cod_credito_transacao")))%>
            </select>
        </td>
        <td style="padding: 5px;" class="MC"><p class="Cd">Pedido novo:</p></td>
        <td align="left" style="padding:5px;" class="MC">
            <input type="text" name="c_pedido_novo" id="c_pedido_novo" maxlength="10" class="PLLe" style="width:80px; margin-left: 10px;" onkeypress="filtra_pedido();" onblur="if (normaliza_numero_pedido_e_sufixo(this.value)!='') {this.value=normaliza_numero_pedido_e_sufixo(this.value);}"
                value="<%=Trim("" & rs("pedido_novo"))%>" <%=s_readonly_taxa_credito%> />
        </td>
    </tr>
    <% s_display = ""
       if Trim("" & rs("cod_credito_transacao"))<>CREDITO_TRANSACAO__REEMBOLSO then s_display = "style=display:none" %>
    <tr id="trDadosBancarios" <%=s_display%>>
        <td colspan="4">
            <table style="width: 649px;" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="padding: 5px;" class="MC"><p class="Cd">Banco:</p></td>
                    <td align="left" style="padding:5px;" class="MC">
                        <input type="text" maxlength="4" name="c_cliente_banco" id="c_cliente_banco" class="PLLe" style="width:60px; margin-left: 10px;"
                            value="<%=Trim("" & rs("banco"))%>" <%=s_readonly_taxa_credito%>>
                    </td>
                    <td style="padding: 5px;" class="MC"><p class="Cd">Agência:</p></td>
                    <td align="left" style="padding:5px;" class="MC">
                        <input type="text" maxlength="10" name="c_cliente_agencia" id="c_cliente_agencia" class="PLLe" style="width:80px; margin-left: 10px;"
                        value="<%=Trim("" & rs("agencia"))%>" <%=s_readonly_taxa_credito%>>
                    </td>
                    <td style="padding: 5px;" class="MC"><p class="Cd">Conta:</p></td>
                    <td align="left" style="padding:5px;" class="MC">
                        <input type="text" maxlength="14" name="c_cliente_conta" id="c_cliente_conta" class="PLLe" style="width:110px; margin-left: 10px;"
                        value="<%=Trim("" & rs("conta"))%>" <%=s_readonly_taxa_credito%>>
                    </td>
                </tr>
                <tr>
                    <td style="padding: 5px;" class="MC"><p class="Cd">Favorecido:</p></td>
                    <td align="left" style="padding:5px;" class="MC" colspan="5">
                        <input type="text" maxlength="40" name="c_cliente_favorecido" id="c_cliente_favorecido" class="PLLe" style="width:340px; margin-left: 10px;"
                        value="<%=Trim("" & rs("favorecido"))%>" <%=s_readonly_taxa_credito%>>
                    </td>
                </tr>
            </table>
        </td>
    </tr>

    <tr>
        <td style="padding: 5px;width:100px;" valign="top" class="MC"><p class="Cd">Observações do crédito</p></td>
        <td style="padding: 5px;" colspan="3" class="MC">
            <textarea name="c_credito_observacoes" id="c_credito_observacoes" class="PLLe" rows="<%=MAX_LINHAS_DESCRICAO_OBSERVACAO_DEVOLUCAO%>" style="width: 520px;" <%=s_readonly_taxa_credito%>><%=Trim("" & rs("credito_observacoes"))%></textarea>
        </td>
    </tr>

<% 
    dim x, full_url_file_src, full_url_file_href, file_attr_title, file_extension
    s = "SELECT * FROM t_UPLOAD_FILE" & _
            " INNER JOIN t_PEDIDO_DEVOLUCAO_IMAGEM ON (t_UPLOAD_FILE.id=t_PEDIDO_DEVOLUCAO_IMAGEM.id_upload_file)" & _
            " WHERE (" & _
                "id_pedido_devolucao = '" & id_devolucao & "'" & _
            ") ORDER BY dt_hr_cadastro"
    if r.State <> 0 then r.Close
    r.open s, cn
    i = 0
%>
    <tr>
	    <td colspan="4" class="Rf tdWithPadding MB MC" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">FOTOS</span></td>
    </tr>
    <tr>
        <td colspan="4">
            <table id="tblFotos" cellpadding="5" cellspacing="0">
                <tr>
<% if r.Eof then 
    if Not operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) Or st_codigo <> COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA then
    %>
                    <td style="width: 100%;">
                        <span class="C" style="color: gray;text-align: center">Não há nenhuma foto vinculada a esta devolução.</span>
                    </td>
    <% end if %>
<% end if %>
<% do while Not r.Eof 
    i = i + 1
    if Trim("" & r("st_file_deleted")) = "1" then
        full_url_file_src = "../IMAGEM/No_Image_Available.png"
    else
        x = r("stored_file_name")
        file_extension = Mid(x, Instr(x, ".")+1, Len(x))
        full_url_file_href = "http://"
        full_url_file_href = full_url_file_href & serverVariablesUrl
        full_url_file_href = full_url_file_href & "FileServer/"
        full_url_file_href = full_url_file_href & r("stored_relative_path")
        full_url_file_href = full_url_file_href & "/" & r("stored_file_name")
        select case LCase(file_extension)
            case "jpeg", "jpg", "png", "bmp", "gif", "tif", "tiff"
                full_url_file_src = full_url_file_href
                file_attr_title = "clique para visualizar a imagem no tamanho original"
            case "pdf"
                full_url_file_src = "../IMAGEM/file_pdf_150x150.png"
                file_attr_title = "clique para visualizar o conteúdo do PDF"
            case else
                full_url_file_src = "../IMAGEM/file_150x150.png"
                file_attr_title = "clique para visualizar o conteúdo do arquivo"
        end select
    end if
    %>
<% if i = 4 then %>
                </tr>
                <tr>
<% end if %>
                <% if Trim("" & r("st_file_deleted")) = "1" then %>
                    <td style="width: 220px;" valign="top">
                        <img src="<%=full_url_file_src%>" style="width: 150px; height: 150px;border:1px dashed black;" />
                    </td>
                <% else %>
                    <td style="width: 220px;" valign="top">
                        <a href="<%=full_url_file_href%>" target="_blank" title="<%=file_attr_title%>">
                        <img src="<%=full_url_file_src%>" style="width: 150px; height: 150px;border:1px dashed black;" /></a>
                    <% if operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) And loja = rs("loja") And st_codigo = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA then %>
                        <a href="javascript:fPEDRemoverFoto('<%=CStr(r("id_upload_file"))%>')" title="remover arquivo">
                        <img src="../BOTAO/botao_X_red.gif" style="margin-left: 0px;vertical-align: top" /></a>
                    <% end if %>
                    </td>
                <% end if %>
<% r.MoveNext
    loop %>

<% if i < PEDIDO_DEVOLUCAO_QTDE_FOTO And operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) And loja = rs("loja") And st_codigo = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA then 
    for cont = i to PEDIDO_DEVOLUCAO_QTDE_FOTO -1
    i = i + 1%>
        
        <% if i = 4 then %>
                </tr>
                <tr>
        <% end if %>
                    <td>
                       <input type='file' name='arquivo<%=Cstr(i)%>' id='arquivo<%=Cstr(i)%>' onchange='reloadFileReader("<%=Cstr(i)%>")' class='PLLd' style='font-weight: normal; width: 180px' /><br />
                       <div id='image-holder-arquivo<%=Cstr(i)%>' style='width: 180px; height: 140px; border: 1px dashed #000; margin-top: 3px' onclick='fPED.arquivo<%=Cstr(i)%>.click();'></div>
                    </td>
<% next
end if %>
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
            <textarea name="c_observacoes" id="c_observacoes" class="PLLe" rows="<%=MAX_LINHAS_DESCRICAO_OBSERVACAO_DEVOLUCAO%>" style="width: 520px;" <%=s_readonly%>><%=Trim("" & rs("observacao"))%></textarea>
        </td>
    </tr>

<% 
    s = "SELECT * FROM t_PEDIDO_DEVOLUCAO_MENSAGEM" & _
        " WHERE (id_pedido_devolucao = '" & id_devolucao & "')" & _
        " ORDER BY dt_hr_cadastro DESC"
    
    if r.State <> 0 then r.Close
    r.open s, cn
%>
    <tr>
		<td colspan="4" class="MC" align="left">
            <a name="aMensagens"></a>
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td class="Rf tdWithPadding" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">MENSAGENS</span></td>
			</tr>
			<% if r.Eof then %>
			<tr>
				<td align="left">&nbsp;</td>
			</tr>
			<% end if %>

			<%	do while not r.Eof %>
			<tr>
				<td align="left">
					<table width="100%" cellspacing="0" cellpadding="0">
					<tr>
					<td class="C MD MC tdWithPadding" style="width:60px;" align="center" valign="top"><%=formata_data_hora_sem_seg(r("dt_hr_cadastro"))%></td>
					<td class="C MD MC tdWithPadding" style="width:80px;" align="center" valign="top"><%
						s = r("usuario_cadastro")
						if Trim("" & r("loja")) <> "" then s = s & " (Loja&nbsp;" & Trim("" & r("loja")) & ")"
						Response.Write s
						%></td>
					<td class="C MC tdWithPadding" align="left" valign="top"><%=substitui_caracteres(Trim("" & r("texto_mensagem")), chr(13), "<br>")%></td>
					</tr>
					</table>
				</td>
			</tr>

<%
    r.MoveNext
    loop
%>
            <% if CInt(rs("st_finalizado"))=0 And CInt(rs("st_reprovado"))=0 And CInt(rs("st_cancelado"))=0 then %>
			<tr class="notPrint">
				<td class="MC" style="padding:0px;" align="left">
					<table width="100%" cellpadding="0" cellspacing="0">
					<tr>
					<td align="left">&nbsp;</td>
                    <% if operacao_permitida(OP_CEN_PRE_DEVOLUCAO_ESCREVER_MSG, s_lista_operacoes_permitidas) then %>
					<td align="center" class="ME MB" style="width:124px;padding:6px;">
						<a name="bMensagemDevolucaoAdiciona" id="bMensagemDevolucaoAdiciona" href="javascript:fPEDDevolucaoMensagemCadastra(fPED,'<%=Trim("" & rs("id"))%>')" title="Grava uma nova mensagem referente a esta pré-devolução">
							<span class="Button" style="width:120px;">Nova Mensagem</span>
						</a>
					</td>
                    <% end if %>
					</tr>
					</table>
				</td>
			</tr>

			<tr class="notPrint">
				<td class="" align="left"><span style='font-family: Arial, Helvetica, sans-serif;color:white;font-size:6pt;font-style:normal;'>&nbsp;</span></td>
			</tr>
			<tr class="notVisible">
				<td align="left"><span style='font-family: Arial, Helvetica, sans-serif;color:white;font-size:6pt;font-style:normal;'>&nbsp;</span></td>
			</tr>
			<% end if %>
    
            </table>
        </td>
    </tr>

    <tr>
        <td colspan="4" class="MB MC" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">HISTÓRICO DE STATUS</span></td>
    </tr>
<% if st_codigo <> COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA then %>
    <% if CInt(rs("st_cancelado")) <> 0 then%>
    <tr>
		<td align="left" colspan="4">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td class="C MD tdWithPadding" style="width:100px;" align="center" valign="top"><%=formata_data_hora_sem_seg(Trim("" & rs("dt_hr_cancelado")))%></td>
			<td class="C MD tdWithPadding" style="width:120px;" align="center" valign="top"><%
				s = Trim("" & rs("usuario_cancelado"))
				Response.Write s
				%></td>
			<td class="C tdWithPadding" align="left" style="padding-left: 5px;" valign="top">CANCELADA</td>
			</tr>
			</table>
		</td>
	</tr>
    <% end if %>
    <% if CInt(rs("st_reprovado")) <> 0 then%>
    <tr>
		<td align="left" colspan="4">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td class="C MD tdWithPadding" style="width:100px;" align="center" valign="top"><%=formata_data_hora_sem_seg(Trim("" & rs("dt_hr_reprovado")))%></td>
			<td class="C MD tdWithPadding" style="width:120px;" align="center" valign="top"><%
				s = Trim("" & rs("usuario_reprovado"))
				Response.Write s
				%></td>
			<td class="C tdWithPadding" align="left" style="padding-left: 5px;" valign="top">REPROVADA</td>
			</tr>
			</table>
		</td>
	</tr>
    <% end if %>
    <% if CInt(rs("st_finalizado")) <> 0 then%>
    <tr>
		<td align="left" colspan="4">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td class="C MD tdWithPadding" style="width:100px;" align="center" valign="top"><%=formata_data_hora_sem_seg(Trim("" & rs("dt_hr_finalizado")))%></td>
			<td class="C MD tdWithPadding" style="width:120px;" align="center" valign="top"><%
				s = Trim("" & rs("usuario_finalizado"))
				Response.Write s
				%></td>
			<td class="C tdWithPadding" align="left" style="padding-left: 5px;" valign="top">FINALIZADA</td>
			</tr>
			</table>
		</td>
	</tr>
    <% end if %>
    <% if CInt(rs("st_mercadoria_recebida")) <> 0 then%>
    <tr>
		<td align="left" colspan="4">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td class="C MD tdWithPadding" style="width:100px;" align="center" valign="top"><%=formata_data_hora_sem_seg(Trim("" & rs("dt_hr_mercadoria_recebida")))%></td>
			<td class="C MD tdWithPadding" style="width:120px;" align="center" valign="top"><%
				s = Trim("" & rs("usuario_mercadoria_recebida"))
				Response.Write s
				%></td>
			<td class="C tdWithPadding" align="left" style="padding-left: 5px;" valign="top">MERCADORIA RECEBIDA</td>
			</tr>
			</table>
		</td>
	</tr>
    <% end if %>
    <% if CInt(rs("st_aprovado")) <> 0 then%>
    <tr>
		<td align="left" colspan="4">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td class="C MD tdWithPadding" style="width:100px;" align="center" valign="top"><%=formata_data_hora_sem_seg(Trim("" & rs("dt_hr_aprovado")))%></td>
			<td class="C MD tdWithPadding" style="width:120px;" align="center" valign="top"><%
				s = Trim("" & rs("usuario_aprovado"))
				Response.Write s
				%></td>
			<td class="C tdWithPadding" align="left" style="padding-left: 5px;" valign="top">EM ANDAMENTO</td>
			</tr>
			</table>
		</td>
	</tr>
    <% end if %>
    <tr>
		<td align="left" colspan="4">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td class="C MD tdWithPadding" style="width:100px;" align="center" valign="top"><%=formata_data_hora_sem_seg(Trim("" & rs("dt_hr_cadastro")))%></td>
			<td class="C MD tdWithPadding" style="width:120px;" align="center" valign="top"><%
				s = Trim("" & rs("usuario_cadastro"))
				Response.Write s
				%></td>
			<td class="C tdWithPadding" align="left" style="padding-left: 5px;" valign="top">CADASTRADA</td>
			</tr>
			</table>
		</td>
	</tr>

<% else %>
    <tr>
		<td align="left" colspan="4">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td class="C MD tdWithPadding" style="width:100px;" align="center" valign="top"><%=formata_data_hora_sem_seg(st_data_hora)%></td>
			<td class="C MD tdWithPadding" style="width:120px;" align="center" valign="top"><%
				s = st_usuario
				Response.Write s
				%></td>
			<td class="C tdWithPadding" align="left" style="padding-left: 5px;" valign="top"><%=UCase(st_devolucao_descricao)%></td>
			</tr>
			</table>
		</td>
	</tr>
<% end if %>

<% if operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) And loja = rs("loja") And st_codigo = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA then %>
    <tr>
        <td colspan="4" class="MB MC" align="left" style="padding:2px;"><span class="Rf" style="padding-left: 5px;">-</span></td>
    </tr>
    <tr>
		<td class="C tdWithPadding" align="left" style="width:100px;" align="center" valign="top">
			<input type="radio" name="rb_status" id="rb_status_aprova" value="<%=COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO%>" />
            <label for="rb_status_aprova"><span class="C" style="color:olivedrab;">APROVAR</span></label>
		</td>
        <td class="C tdWithPadding" colspan="3" align="left" style="width:100px;" align="center" valign="top">
			<input type="radio" name="rb_status" id="rb_status_reprova" value="<%=COD_ST_PEDIDO_DEVOLUCAO__REPROVADA%>" />
            <label for="rb_status_reprova"><span class="C" style="color:darkred;">REPROVAR</span></label>
		</td>
	</tr>
<% end if %>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="852" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<!-- ************   BOTÕES   ************ -->
<table width="649" cellSpacing="0">
<tr>
    <% if url_back = "" then
        s = "javascript:history.back()"
        else
        s = "resumo.asp"
        end if %>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
<% if blnExibeBotaoConfirma then %>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDConfirma(fPED)" title="confirma a atualização da devolução">
		<img src="../botao/confirmar.gif" width="176" height="55" id="btnCONFIRMAR" border="0"></a></div>
	</td>
<% end if %>
<% if ((usuario = rs("usuario_cadastro") Or usuario = r_pedido.vendedor) And st_codigo = COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA) then %>
<% if blnExibeBotaoConfirma then %>
</tr>
<tr>
<% end if %>
    <td align="left">&nbsp;</td>
    <td align="right"><div name="dREMOVER" id="dREMOVER" style="display: none;">
		<a name="bREMOVER" id="bREMOVER" href="javascript:fPEDCancelar(fPED)" title="cancelar a devolução">
		<img src="../botao/remover.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% end if %>
<% if operacao_permitida(OP_LJA_PRE_DEVOLUCAO_ADMINISTRACAO, s_lista_operacoes_permitidas) And loja = rs("loja") And st_codigo = COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA then %>
<% if blnExibeBotaoConfirma then %>
</tr>
<tr>
<% end if %>
    <td align="left">&nbsp;</td>
    <td align="right"><div name="dFINALIZAR" id="dFINALIZAR">
		<a name="bFINALIZAR" id="bFINALIZAR" href="javascript:fPEDFinalizar(fPED)" title="finalizar a devolução">
		<img src="../botao/finalizar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% end if %>

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
